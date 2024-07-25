from helpers import *

# Nos conectamos a la bandeja (otra vez con Davmail)
mailbox = MailBox(
    "0.0.0.0", port=0, ssl_context=ssl._create_unverified_context()
).login("a@imss.gob", password)
mailbox.folder.set("INBOX")
for msg in mailbox.fetch():
    # primero validamos que el correo venga de un remitente @imss.gob, si sí, lo almacenamos 
    if "@imss.gob" in msg.from_:
        try:
            # limpiamos el asunto
            asunto = msg.subject.strip().replace("RV: ","").replace("RE: ","").replace('_','-').replace(' ','')
        except Exception:
            asunto = msg.subject.strip()
        # guardamos el archivo
        with open(
            f'./archivo/{msg.date}-{asunto.replace("/","-")}.elm'.strip().replace(
                " ", "-"
            ),
            "w+",
        ) as out:
            gen = email.generator.Generator(out)
            gen.flatten(msg.obj)
        if "atendid" not in msg.subject.lower():
            if msg.from_ in subdelegados:
                try:
                    # la estrategia va a ser ésta, como IMAP no tiene ID's especificos por correo, vamos a generar una llave
                    # ésta está compuesta por: fecha-encabezado-remitente-cuerpo en un JWT (para que después podamos descifrar)
                    try:
                        tipo_operacion = asunto.split("-")[2]
                    except Exception:
                        # mailbox.move(f"{msg.uid}", "NO-SOLICITUDES")
                        raise Exception(f"Asunto {asunto} no válido")
                    # cuerpo_correo = msg.text.split("Enviado el:")[0]
                    # evaluamos el asunto, cuerpo y anexos, en orden, si hay excepciones en alguno, las guardamos
                    asunto_valido, excepcion_asunto = validar_asunto(asunto=asunto)
                    cuerpo_valido, excepcion_cuerpo = validar_cuerpo_correo(
                        msg.text, tipo_operacion
                    )
                    # dado que en esta funcion validamos la bitacora, la regresamos de una vez :))
                    anexos_validos, excepcion_anexos, bitacora = validar_anexos(
                        msg.attachments, tipo_operacion, asunto, msg.text
                    )
                    # si falló al menos una validación, rechazamos
                    if asunto_valido and cuerpo_valido and anexos_validos:
                        usuario_atencion = (
                            df_usuarios.loc[
                                df_usuarios.cve_solicitud.str.lower()
                                == tipo_operacion.lower()
                            ]
                            .reset_index(drop=True)
                            .iloc[0]["responsable"]
                        )
                        # preparamos el objeto a insertar en la bd
                        mongo_object = {
                            "fecha": msg.date,
                            "asunto": asunto,
                            "remitente": msg.from_,
                            "len_msg": len(msg.text),
                            "atendido_por": usuario_atencion,
                            "atendido": 0,
                            "delegacion": asunto.split("-")[0],
                            "subdelegacion": asunto.split("-")[1],
                            "operacion": asunto.split("-")[2],
                            "sujeto": asunto.split("-")[3],
                        }
                        # movemos el archivo a la carpeta /SOLICITUDES VALIDAS/ y enviamos los correos de respuesta y atención
                        mailbox.move(f"{msg.uid}", "SOLICITUDES VALIDAS")
                        correo_respuesta(True, "", msg.from_, asunto)
                        # inseramos en la bd
                        col_solicitudes.insert_one(mongo_object)
                        col_solicitudes2.insert_one(mongo_object)
                        # le ponemos fecha y asunto a la bitácora y la insertamos
                        bitacora["asunto"] = asunto
                        bitacora["fecha"] = msg.date
                        col_bitacora.insert_many(bitacora.to_dict('records'))
                        col_bitacora2.insert_many(bitacora.to_dict("records"))

                        correo_atender(msg.obj, asunto, tipo_operacion, msg.from_)
                    else:
                        # mailbox.move(f"{msg.uid}", "NO-SOLICITUDES")
                        # print(', '.join([str(x) for x in [excepcion_asunto, excepcion_cuerpo, excepcion_anexos] if x is not None]))
                        raise Exception(
                            " ".join(
                                [
                                    f"<li>{str(x)}</li>"
                                    for x in [
                                        excepcion_asunto,
                                        excepcion_cuerpo,
                                        excepcion_anexos,
                                    ]
                                    if x is not None
                                ]
                            )
                        )
                except Exception as e:
                    # print(e)
                    mailbox.move(f"{msg.uid}", "NO-SOLICITUDES")
                    correo_respuesta(False, e, msg.from_, asunto)
                    mongo_object = {
                        "fecha": msg.date,
                        "asunto": asunto,
                        "remitente": msg.from_,
                        "len_msg": len(msg.text),
                        "excepcion_asunto": str(excepcion_asunto),
                        "excepcion_anexo": str(excepcion_anexos),
                        "excepcion_cuerpo": str(excepcion_cuerpo),
                        "remitente_permitido": msg.from_ in subdelegados,
                    }
                    # insertamos el error en otra colección de la BD
                    col_errores.insert_one(mongo_object)
                    # correo_respuesta(False, e, "b@imss.gob", asunto)
            elif (
                msg.from_
                == "M@imss.gob"
            ):
                # los correos de buzón lleno, los eliminamos
                print("buzon lleno")
                mailbox.delete(f"{msg.uid}")
                pass
            else:
                try:
                    raise Exception(
                        "Sólo el o la subdelegada y el o la encargada de la subdelegación pueden enviar correos a esta dirección."
                    )
                except Exception as e:
                    print("FUERA DE LISTA")
                    asunto = msg.subject.strip()
                    correo_respuesta(False, e, msg.from_, asunto)
                    # correo_respuesta(False, e, "b@imss.gob", asunto.strip().replace("\\r\\n","")[0:65])
                    mailbox.move(f"{msg.uid}", "NO-SOLICITUDES")

        else:
            try:
                asunto = (
                    msg.subject.replace("\r\n", "")
                    .split(" - Enviado por:")[0]
                    .split(":")[-1]
                    .strip()
                    .replace(" ","")
                )
                operacion = asunto.split("-")[2]
                solicitudes_usuario = pd.DataFrame(
                    col_solicitudes2.find(
                        {
                            "operacion": operacion,
                            "$or": [
                                {"estatus": {"$in": ["Parcial", "Rechazado"]}},
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
                            solicitudes_usuario.asunto == asunto.lower()
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
                            "fecha_atencion": datetime.today(),
                            "estatus": status,
                        }
                    },
                )
                mailbox.move(f"{msg.uid}", "ATENDIDOS")
                correo_respuesta_atencion(True, msg.from_, asunto)
            except Exception as e:
                print(e)
                correo_respuesta_atencion(False, "b@imss.gob", asunto)
    else:
        print("no-imss")
        pass
