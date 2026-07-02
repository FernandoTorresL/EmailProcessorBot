from helpers import *

# Nos conectamos a la bandeja (otra vez con Davmail)
try:
    mailbox = MailBox(
        IP_MAILBOX, port=PORT_MAILBOX, ssl_context=ssl._create_unverified_context()
    ).login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
    mailbox.folder.set(FOLDER_MAILBOX)

    cifras = ''
    i_total_correos = 0
    
    i_total_validos = 0
    i_total_atendidos = 0
    i_total_no_solicitudes = 0

    i_total_fuera_de_lista = 0
    i_total_error_al_marcar = 0
    i_total_buzon_lleno = 0
    i_total_junk = 0

    i_total_nuevos_motivos = 0

    for msg in mailbox.fetch(limit=MSG_LIMIT, reverse=False):

        i_total_correos += 1
        # primero validamos que el correo venga de un remitente @imss.gob, si sí, lo almacenamos
        if (DOMINIO_MAILBOX in msg.from_) and (BAD_MAIL_STRING not in msg.subject.strip()):
            try:
                # limpiamos el asunto
                asunto = (
                    msg.subject.strip()
                    # .replace("RV: ", "")
                    # .replace("RE: ", "")
                    # .replace("RV:", "")
                    # .replace("RE:", "")
                    .replace("_", "-")
                    .replace("<", "")
                    .replace(">", "")
                    .replace(" ", "")
                    .replace("|", "")
                    .replace("*", "")
                )
            except Exception:
                asunto = msg.subject.strip()

            # comprobamos que exista la carpeta por año
            backup_folder = datetime.now().strftime("%Y")
            backup_folder = PATH_ARCHIVO + "/" + backup_folder

            if not os.path.exists(backup_folder):
                # Enviar correo avisando la creación de la carpeta de año
                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: No existe carpeta {backup_folder}", f"Se intentará crear carpeta {backup_folder}")
                
                os.mkdir(backup_folder)

                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: Se ha creado carpeta {backup_folder}", f"Se ha creado la nueva carpeta {backup_folder}")

            # comprobamos que exista la carpeta por mes
            backup_folder = backup_folder + "/" + datetime.now().strftime("%m")

            if not os.path.exists(backup_folder):
                # Enviar correo avisando la creación de la carpeta
                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: No existe carpeta {backup_folder}", f"Se intentará crear carpeta {backup_folder}")

                os.mkdir(backup_folder)

                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: Se ha creado carpeta {backup_folder}", f"Se ha creado la nueva carpeta {backup_folder}")

            # comprobamos que exista la carpeta por día
            backup_folder = backup_folder + "/" + datetime.now().strftime("%d")

            if not os.path.exists(backup_folder):
                # Enviar correo avisando la creación de la carpeta
                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: No existe carpeta {backup_folder}", f"Se intentará crear carpeta {backup_folder}")

                os.mkdir(backup_folder)

                correo_debug(EMAIL_ADMINISTRADOR, f"LOG: Se ha creado carpeta {backup_folder}", f"Se ha creado la nueva carpeta {backup_folder}")

            # Definimos un prefijo
            prefijo_filename = (f"{msg.date}")
            prefijo_filename = prefijo_filename[:16]
            prefijo_filename = prefijo_filename.replace(':','_')
            prefijo_filename = prefijo_filename.replace('-','_')
            prefijo_filename = prefijo_filename.replace(' ','_')

            # Recortamos el nombre del archivo
            filename = msg.subject.strip().upper()
            filename = filename.replace(DOMINIO_MAILBOX.upper(), '')
            filename = filename.replace('"', '')
            filename = filename.replace("'", "")
            filename = filename.replace('RV: ', '')
            filename = filename.replace('RE: ', '')
            filename = filename.replace('RV:', '')
            filename = filename.replace('RE:', '')
            filename = filename.replace('ATENDIDO-CON RECHAZO ', 'AR-')
            filename = filename.replace('ATENDIDO CON RECHAZO ', 'AR-')
            filename = filename.replace('ATENDIDO ', 'AT-')
            filename = filename.replace('A ATENDER: ', '')
            filename = filename.replace('CON RECHAZO ', 'RECHAZO ')
            filename = filename.replace('ENVIADO POR: ', '-')
            filename = filename.replace('/', '-')
            filename = filename.replace(' ', '-')
            filename = filename.replace('_', '-')
            filename = filename.replace('<', '-')
            filename = filename.replace('>', '-')
            filename = filename.replace('\r\n', '')
            filename = filename.replace('\t', '')
            filename = filename.replace("'", '')
            filename = filename.replace('|', '')
            filename = filename.replace('*', '')
            filename = filename.replace(':', '')
            filename = filename.replace('+0000', '')
            filename = filename.replace('---', '-')
            filename = filename.replace('--', '-')

            filename = filename[:28]
            filename = filename.strip()

            dice_roll = random.randint(1,100)
            dice_roll = f"{dice_roll:03d}"

            filename = prefijo_filename + "-" + filename[:47] + "_" + dice_roll

            path_filename = backup_folder + "/" + filename + '.eml'
            path_filename_zip = backup_folder + "/" + filename + '.zip'

            if not SOLO_REENVIO_POR_BUZON_LLENO:
                # guardamos el archivo
                with open(
                    path_filename,
                    "w+",
                ) as out:
                    gen = email.generator.Generator(out)
                    gen.flatten(msg.obj)

                with zipfile.ZipFile(path_filename_zip,'w', zipfile.ZIP_DEFLATED) as zipf:
                    # Añadir el archivo EML
                    zipf.write(path_filename)

                os.remove(path_filename)

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
                            i_total_validos += 1

                            #### 
                            if not SOLO_REENVIO_POR_BUZON_LLENO:
                                correo_respuesta(True, "", msg.from_, asunto, operation_upper)
                                # inseramos en la bd
                                # col_solicitudes.insert_one(mongo_object)
                                col_solicitudes2.insert_one(mongo_object)
                                # le ponemos fecha y asunto a la bitácora y la insertamos

                                # Si NO se trata de los 3 nuevos asuntos, realiza:
                                if not tipo_operacion.lower() in ["uiss88", "ufc1127f", "ufc1127c"]:
                                    bitacora["asunto"] = asunto
                                    bitacora["fecha"] = msg.date

                                    data = bitacora.to_dict("records")
                                    enviar_correo = False

                                    # Conversión a texto y con ceros a la izquierda
                                    for item in data:
                                        for key, value in item.items():
                                            if key == "CLAVE_OOAD":
                                                # Convertir a string y limpiar espacios
                                                s = str(value).strip()

                                                # Extraer solo los dígitos
                                                match = re.search(r'\d+', s)
                                                if match:
                                                    num = int(match.group())
                                                    
                                                    # Opcional: validar rango 1–40
                                                    if 1 <= num <= 40:
                                                        item[key] = f"_{num:02d}"
                                                    else:
                                                        # Si está fuera de rango puedes decidir qué hacer
                                                        item[key] = None  # o dejarlo igual, o lanzar error
                                                        print(f"CLAVE_OOAD {num} inválida en bitácora")
                                                        # correo_respuesta_bitacora(True, f"CLAVE_OOAD {num} inválida en bitácora", msg.from_, asunto, operation_upper)
                                                else:
                                                    item[key] = None  # o manejar caso sin números
                                                    print(f"CLAVE_OOAD {num} inválida en bitácora")
                                                    # correo_respuesta_bitacora(True, f"CLAVE_OOAD {num} inválida en bitácora", msg.from_, asunto, operation_upper)
                                            if key == "CLAVE_SUBDELEGACIÓN":
                                                item[key] = str(value).replace("_", "").zfill(2)   # siempre 2 caracteres
                                            elif key == "CONSECUTIVO_5_POSICIONES":
                                                item[key] = str(value).zfill(5)   # siempre 5 caracteres
                                            elif key == "NSS_11_POSICIONES":
                                                item[key] = str(value).zfill(11)  # siempre 11 caracteres
                                            elif isinstance(value, (int, float)) and not isinstance(value, pd.Timestamp):
                                                item[key] = "_" + str(value).zfill(2)            
                                                enviar_correo = True
                                                # otros numéricos, solo convertir a texto
                                                # correo_respuesta_bitacora(True,"Revisar CLAVE_OOAD para futuras solicitudes, la cual debe estar en formato _XX, como se definió en la plantilla. ", msg.from_, asunto, operation_upper)

                                    col_bitacora2.insert_many(data)

                                # if enviar_correo:
                                    # correo_respuesta_bitacora(True,"Revisar CLAVE_OOAD para futuras solicitudes, la cual debe estar en formato _XX, como se definió en la plantilla. ", msg.from_, asunto, operation_upper)

                            datetime_correo_original = msg.date-timedelta(hours=6)

                            correo_atender(msg.obj, asunto, tipo_operacion, msg.from_, datetime_correo_original)

                            # Código para solicitud de CA (revisión por OOAD)
                            if tipo_operacion.lower() in ["uiss88"]:
                                # "delegacion": asunto.split("-")[0],
                                # "subdelegacion": asunto.split("-")[1],

                                ooad_revisores = asunto.split("-")[0]
                                subdel_revisores = asunto.split("-")[1]

                                correo_atender_revisores(msg.obj, asunto, tipo_operacion, msg.from_, datetime_correo_original)

                                i_total_nuevos_motivos += 1
                            # Fin código para solicitud de CA (revisión por OOAD)

                            # print("SOL. VALIDA: ", asunto)
                        else:
                            # mailbox.(f"{msg.uid}", "INBOX/NO-SOLICITUDES")
                            # print("SOL. NO VALIDA: ", asunto)
                            # print(', '.join([str(x) for x in [excepcion_asunto, excepcion_cuerpo, excepcion_anexos] if x is not None]))
                            raise Exception(" ".join([f"<li>{str(x)}</li>" for x in [excepcion_asunto, excepcion_cuerpo, excepcion_anexos, excepcion_tamanio] if x is not None]))
                    except Exception as e:
                        # print(e)
                        mailbox.move(f"{msg.uid}", "INBOX/NO-SOLICITUDES")
                        i_total_no_solicitudes += 1
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
                    mailbox.move(f"{msg.uid}", "INBOX/TMP")
                    i_total_buzon_lleno += 1

                    # correo_debug(EMAIL_ADMINISTRADOR, f"WARNING: Buzón lleno", f"Se ha llenado el buzón {mensaje_buzon_lleno}")

                    # print(f"Buzón lleno {mensaje_buzon_lleno}")
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
                        i_total_fuera_de_lista += 1
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

                    ayudante_email = ""
                    operacion = asunto.split("-")[2].upper()

                    iniciales = msg.subject.replace("\r\n", "")

                    if operacion == "MOTIVO1" and "(" in iniciales and ")" in iniciales:
                        iniciales = iniciales.split("(")[1].split(")")[0]

                        ayudante_email = AYUDANTES[iniciales]

                    solicitudes_usuario = pd.DataFrame(
                        col_solicitudes2.find({
                            "operacion": {
                                "$regex": operacion, "$options": "i"},
                            "atendido": 0,
                            }
                        )
                    )

                    solicitudes_usuario["asunto"] = solicitudes_usuario["asunto"].str.lower()

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

                    if operacion == "MOTIVO1":
                        col_solicitudes2.update_one(
                            {"_id": id_update},
                            {
                                "$set": {
                                    "atendido": 1,
                                    "atendido_por": msg.from_,
                                    "fecha_atencion": msg.date,
                                    "ayudante": ayudante_email,
                                    "estatus": status,
                                }
                            },
                        )
                    else:
                        col_solicitudes2.update_one(
                            {"_id": id_update},
                            {
                                "$set": {
                                    "atendido": 1,
                                    "atendido_por": msg.from_,
                                    "fecha_atencion": msg.date,
                                    "estatus": status,
                                }
                            },
                        )

                    mailbox.move(f"{msg.uid}", "INBOX/ATENDIDOS")
                    i_total_atendidos += 1
                    correo_respuesta_atencion(True, msg.from_, asunto)
                    # print("SOL. DE ATENDIDOS VALIDA: ", asunto)

                except Exception as e:
                    print(e)
                    mailbox.move(f"{msg.uid}", "INBOX/ERROR-AL-MARCAR")
                    i_total_error_al_marcar += 1
                    print("ERROR AL MARCAR: ", asunto)
                    # correo_respuesta_atencion(False, EMAIL_ADMINISTRADOR, asunto)
                    correo_respuesta_atencion(False, msg.from_, asunto)
                    print("")
        else:
            print("no-imss")
            mailbox.move(f"{msg.uid}", "Junk")
            i_total_junk += 1

            correo_debug(EMAIL_ADMINISTRADOR, "WARNING EmailProcessorBot - Correo Junk", f"Correo externo de {msg.from_}")
            pass

        # Removemos el archivo original .eml
        # os.remove(path_filename)

    i_total_correos = f"{i_total_correos:03d}"
    print ('Correos procesados: ', i_total_correos)
    print("No hay más mensajes. FIN")

    i_total_validos = f"{i_total_validos:03d}"
    i_total_atendidos = f"{i_total_atendidos:03d}"
    i_total_no_solicitudes = f"{i_total_no_solicitudes:03d}"
    i_total_fuera_de_lista = f"{i_total_fuera_de_lista:03d}"
    i_total_error_al_marcar = f"{i_total_error_al_marcar:03d}"
    i_total_buzon_lleno = f"{i_total_buzon_lleno:03d}"
    i_total_junk = f"{i_total_junk:03d}"

    i_total_nuevos_motivos = f"{i_total_nuevos_motivos:03d}"

    cifras = f"""
        <p>CIFRAS {datetime.now().strftime("%A, %d de %B de %Y %H:%M:%S")}</p>
        <p><strong>{i_total_correos} Correos procesados</strong></p>
            <p>
                <ol>{i_total_validos} Validos</ol>
                <ol>{i_total_atendidos} Atendidos</ol>
                <ol>{i_total_no_solicitudes} No_solicitudes</ol>
                    <ol>{i_total_fuera_de_lista} Fuera_de_lista</ol>
                <br>
                <ol>{i_total_error_al_marcar} Error_al_marcar</ol>
                <ol>{i_total_buzon_lleno} Buzon_lleno</ol>
                <ol>{i_total_junk} Junk</ol>
                <br>
                <ol>{i_total_nuevos_motivos} Nuevos motivos</ol>
            </p>
            <br>
    """
    
    correo_debug(EMAIL_ADMINISTRADOR, "LOG EmailProcessorBot - Finalización exitosa", cifras)

    print('Total_correos: ', i_total_correos)
    print('Validos: ', i_total_validos)
    print('Atendidos: ', i_total_atendidos)
    print('No_solicitudes: ', i_total_no_solicitudes)
    print('Fuera_de_lista: ', i_total_fuera_de_lista)
    print('Error_al_marcar: ', i_total_error_al_marcar)
    print('Buzon_lleno: ', i_total_buzon_lleno)
    print('Junk: ', i_total_junk)
    print('Nuevos_motivos: ', i_total_nuevos_motivos)

except FileNotFoundError as fnf_error:
    desc_error = f"Error: {fnf_error}"

    if not path_filename:
        path_filename = "|SIN-DATO|"

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - FNF", f"{desc_error}. Archivo {path_filename} con nombre original {msg.subject} no localizado")

except zipfile.BadZipFile as bzf_error:
    desc_error = f"Error: {bzf_error}"

    if not path_filename_zip:
        path_filename_zip = "|SIN-DATO2|"

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - Error zip file", f"{desc_error}. Archivo {path_filename_zip} con nombre original {msg.subject}")

except IOError as e:
    desc_error = f"Error al manejar el archivo: {e}"
    print(desc_error)

    if not path_filename:
        path_filename = "|SIN-DATO|"

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - Guardar correo", f"{desc_error}. Archivo {path_filename} con nombre original {msg.subject}")

except Exception as e:
    desc_error = f"Ocurrió un error inesperado: {e}"
    print(desc_error)

    if not path_filename:
        path_filename = "|SIN-DATO|"
    if not path_filename_zip:
        path_filename_zip = "|SIN-DATO2|"

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - Error inesperado", f"{desc_error}. Archivo {path_filename}, Zip {path_filename_zip} con nombre original {msg.subject}")

except NameError:
    desc_error = f"Error en nombre de Carpeta {FOLDER_MAILBOX}"
    print(desc_error)

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - Folder(PST)", desc_error)

except TimeoutError:
    desc_error = "Time Out en el servidor de correo"
    print(desc_error)

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - TimeOut", desc_error)

except Exception as e:
    desc_error = f"Ocurrió un error inesperado: {e}"
    print(desc_error)

    if not path_filename:
        path_filename = "|SIN-DATO|"
    if not path_filename_zip:
        path_filename_zip = "|SIN-DATO2|"

    correo_debug(EMAIL_ADMINISTRADOR, "ERROR EmailProcessorBot - Error inesperado", f"{desc_error}. Archivo {path_filename}, Zip {path_filename_zip} con nombre original {msg.subject}")
