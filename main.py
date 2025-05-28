from helpers import *

# Nos conectamos a la bandeja (otra vez con Davmail)
try:
    mailbox = MailBox(
        IP_MAILBOX, port=PORT_MAILBOX, ssl_context=ssl._create_unverified_context()
    ).login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
    mailbox.folder.set(FOLDER_MAILBOX)

    for msg in mailbox.fetch(limit=MSG_LIMIT, reverse=False):
        # primero validamos que el correo venga de un remitente @imss.gob, si sí, lo almacenamos
        if (DOMINIO_MAILBOX in msg.from_) and (BAD_MAIL_STRING not in msg.subject.strip()):
            try:
                # limpiamos el asunto
                asunto = (
                    msg.subject.strip()
                    .replace("RV: ", "")
                    .replace("RE: ", "")
                    .replace("_", "-")
                    .replace("<", "")
                    .replace(">", "")
                    .replace(" ", "")
                    .replace("|", "")
                )
            except Exception:
                asunto = msg.subject.strip()
            # guardamos el archivo

            nombre_filename = msg.subject.strip()
            nombre_filename = nombre_filename.replace('\r\n', '')
            nombre_filename = nombre_filename.replace('\t', '')
            nombre_filename = nombre_filename.replace('--', '')
            nombre_filename = nombre_filename.replace('"', '')
            nombre_filename = nombre_filename.replace('|', '')
            filename = (
                f"{msg.date}-{nombre_filename}.eml".strip()
                .replace(" ", "-")
                .replace(":", "")
                .replace("/", "-")
            )

            with open(
                f"{PATH_ARCHIVO}/{filename}",
                "w+",
            ) as out:
                gen = email.generator.Generator(out)
                gen.flatten(msg.obj)
            if "atendid" not in msg.subject.lower():
                if msg.from_ in subdelegados:
                    tipo_operacion = ""
                    try:
                        # la estrategia va a ser ésta, como IMAP no tiene ID's especificos por correo, vamos a generar una llave
                        # ésta está compuesta por: fecha-encabezado-remitente-cuerpo en un JWT (para que después podamos descifrar)
                        try:
                            tipo_operacion = asunto.split("-")[2]
                        except Exception:
                            excepcion_asunto = "Asunto inválido"
                            excepcion_anexos = excepcion_cuerpo = excepcion_tamanio = ""
                            # print(f"SOL. NO VALIDA de: {msg.from_} | Asunto inválido: {asunto} ")
                            raise Exception(f"Asunto {asunto} no válido")
                        #
                        # cuerpo_correo = msg.text.split("Enviado el:")[0]
                        # evaluamos el asunto, cuerpo y anexos, en orden, si hay excepciones en alguno, las guardamos
                        asunto_valido, excepcion_asunto = validar_asunto(asunto=asunto, remitente=msg.from_)
                        cuerpo_valido, excepcion_cuerpo = validar_cuerpo_correo(
                            msg.text, tipo_operacion
                        )

                        tamanio_valido, excepcion_tamanio = validar_tamanio(msg.attachments, asunto)

                        # dado que en esta funcion validamos la bitacora, la regresamos de una vez :))
                        anexos_validos, excepcion_anexos, bitacora = validar_anexos(
                            msg.attachments, tipo_operacion, asunto, msg.text, msg.size
                        )


                        # si falló al menos una validación, rechazamos
                        if asunto_valido and cuerpo_valido and anexos_validos and tamanio_valido:
                            usuario_atencion = (
                                df_usuarios.loc[
                                    df_usuarios.cve_solicitud.str.lower()
                                    == tipo_operacion.lower()
                                ]
                                .reset_index(drop=True)
                                .iloc[0]["responsable"]
                            )
                            # preparamos el objeto a insertar en la bd
                            operation_upper = asunto.split("-")[2].upper()
                            mongo_object = {
                                "fecha": msg.date,
                                "asunto": asunto.upper(),
                                "remitente": msg.from_,
                                "len_msg": len(msg.text),
                                "atendido_por": usuario_atencion,
                                "atendido": 0,
                                "delegacion": asunto.split("-")[0],
                                "subdelegacion": asunto.split("-")[1],
                                "operacion": operation_upper,
                                "sujeto": asunto.split("-")[3],
                            }
                            # movemos el archivo a la carpeta /SOLICITUDES VALIDAS/ y enviamos los correos de respuesta y atención
                            mailbox.move(f"{msg.uid}", "INBOX/SOLICITUDES VALIDAS")
                            correo_respuesta(True, "", msg.from_, asunto, operation_upper)
                            # inseramos en la bd
                            # col_solicitudes.insert_one(mongo_object)
                            col_solicitudes2.insert_one(mongo_object)
                            # le ponemos fecha y asunto a la bitácora y la insertamos
                            bitacora["asunto"] = asunto
                            bitacora["fecha"] = msg.date
                            # col_bitacora.insert_many(bitacora.to_dict("records"))
                            col_bitacora2.insert_many(bitacora.to_dict("records"))

                            correo_atender(msg.obj, asunto, tipo_operacion, msg.from_)
                            print("SOL. VALIDA: ", asunto)
                        else:
                            # mailbox.(f"{msg.uid}", "INBOX/NO-SOLICITUDES")
                            # print("SOL. NO VALIDA: ", asunto)
                            # print(', '.join([str(x) for x in [excepcion_asunto, excepcion_cuerpo, excepcion_anexos] if x is not None]))
                            raise Exception(" ".join([f"<li>{str(x)}</li>" for x in [excepcion_asunto, excepcion_cuerpo, excepcion_anexos, excepcion_tamanio] if x is not None]))
                    except Exception as e:
                        # print(e)
                        mailbox.move(f"{msg.uid}", "INBOX/NO-SOLICITUDES")
                        correo_respuesta(False, e, msg.from_, asunto, tipo_operacion.upper())
                        mongo_object = {
                            "fecha": msg.date,
                            "asunto": asunto.upper(),
                            "remitente": msg.from_,
                            "len_msg": len(msg.text),
                            "excepcion_asunto": str(excepcion_asunto),
                            "excepcion_anexo": str(excepcion_anexos),
                            "excepcion_cuerpo": str(excepcion_cuerpo),
                            "excepcion_tamanio": str(excepcion_tamanio),
                            "remitente_permitido": msg.from_ in subdelegados,
                        }
                        # insertamos el error en otra colección de la BD
                        col_errores.insert_one(mongo_object)
                        print("SOL. NO VALIDA: ", asunto)
                elif (
                    msg.from_
                    == EMAIL_MICROSOFT
                ):
                    # los correos de buzón lleno, los eliminamos
                    mensaje_buzon_lleno = msg.text[:150].replace('\r\n', '')
                    print(f"buzon lleno {mensaje_buzon_lleno}")
                    mailbox.move(f"{msg.uid}", "INBOX/TMP")
                    #mailbox.delete(f"{msg.uid}")
                    pass
                else:
                    try:
                        raise Exception(
                            "Sólo el o la subdelegada y el o la encargada de la subdelegación pueden enviar correos a esta dirección."
                        )
                    except Exception as e:
                        print(f"Remitente FUERA DE LISTA {msg.from_}")
                        asunto = msg.subject.strip()
                        correo_respuesta(False, e, msg.from_, asunto, "")
                        mailbox.move(f"{msg.uid}", "INBOX/NO-SOLICITUDES")
                        print("SOL. NO VALIDA: ", asunto)

            else:
                try:
                    asunto = (
                        msg.subject.replace("\r\n", "")
                        .split(" - Enviado por:")[0]
                        .split(":")[-1]
                        .strip()
                        .replace(" ", "")
                    )
                    operacion = asunto.split("-")[2].upper()
                    solicitudes_usuario = pd.DataFrame(
                        col_solicitudes2.find(
                            {
                                "operacion": { "$regex": operacion, "$options": "i" },
                                "$or": [
                                    {"estatus": {"$in": ["Parcial", "Rechazado",
                                                         "Reasignado"]}},
                                    {"atendido": 0},
                                ],
                            }
                        )
                    )
                    solicitudes_usuario["asunto"] = solicitudes_usuario[
                        "asunto"
                    ].str.lower()
                    try:
                        id_update = (
                            solicitudes_usuario.loc[
                                solicitudes_usuario.asunto.str.lower() == asunto.lower()
                            ]
                            .head(1)["_id"]
                            .tolist()[0]
                        )
                    except Exception:
                        raise Exception(
                            f"O no encontramos la solicitud con id {asunto} o éste ya ha sido marcado como atendido."
                        )
                    status = ""
                    if "rechaz" in msg.subject.lower():
                        status = "Rechazado"
                    elif "parcial" in msg.subject.lower():
                        status = "Parcial"
                    # col_solicitudes.update_one({'_id':id_update},{'$set':{'atendido':1, 'fecha_atencion':datetime.today(), 'estatus':status}})
                    col_solicitudes2.update_one(
                        {"_id": id_update},
                        {
                            "$set": {
                                "atendido": 1,
                                "atendido_por": msg.from_,
                                "fecha_atencion": msg.date,
                                # "fecha_atencion": datetime.today(),
                                "estatus": status,
                            }
                        },
                    )
                    mailbox.move(f"{msg.uid}", "INBOX/ATENDIDOS")
                    correo_respuesta_atencion(True, msg.from_, asunto)
                except Exception as e:
                    print(e)
                    mailbox.move(f"{msg.uid}", "INBOX/ERROR-AL-MARCAR")
                    print("ERROR AL MARCAR: ", asunto)
                    # correo_respuesta_atencion(False, EMAIL_ADMINISTRADOR, asunto)
                    correo_respuesta_atencion(False, msg.from_, asunto)
        else:
            print("no-imss")
            mailbox.move(f"{msg.uid}", "Junk")
            pass

    print("No hay más mensajes. FIN")

except NameError:
    print(f"Error en nombre de Carpeta {FOLDER_MAILBOX}")
except TimeoutError:
    print("Time Out en el servidor de correo")
#except imaplib.IMAP4.abort:
    #print("imaplib.IMAP4.abort Exception ocurred")
except Exception as e:
    print(e)
