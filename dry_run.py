# Uso:
# python dry_run.py --destinatario "ejemplo@imss.gob"
# si no se proporciona ningún argumento, se envían los correos a b@imss.gob


import argparse

from helpers import *

parser = argparse.ArgumentParser()
parser.add_argument('--destinatario', '-d', help="Correo a enviar los mensajes de confirmación", type= str, nargs='?', const=EMAIL_ADMINISTRADOR, default=EMAIL_ADMINISTRADOR)
args=parser.parse_args()

# Nos conectamos a la bandeja (otra vez con Davmail)
mailbox = MailBox(
    IP_MAILBOX, port=PORT_MAILBOX, ssl_context=ssl._create_unverified_context()
).login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
mailbox.folder.set("INBOX")
for msg in mailbox.fetch():
    # primero validamos que el correo venga de un remitente @imss.gob, si sí, lo almacenamos 
    if DOMINIO_MAILBOX in msg.from_:
        try:
            # limpiamos el asunto
            asunto = msg.subject.strip().replace("RV: ","").replace("RE: ","").replace('_','-').replace(' ','')
        except Exception:
            asunto = msg.subject.strip()
        if "atendid" not in msg.subject.lower():
            if msg.from_ in subdelegados:
                tipo_operacion = ""
                try:
                    # la estrategia va a ser ésta, como IMAP no tiene ID's especificos por correo, vamos a generar una llave
                    # ésta está compuesta por: fecha-encabezado-remitente-cuerpo en un JWT (para que después podamos descifrar)
                    try:
                        tipo_operacion = asunto.split("-")[2]
                    except Exception:
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
                            "operacion": asunto.split("-")[2].upper(),
                            "sujeto": asunto.split("-")[3],
                        }
                        correo_respuesta(True, "", args.destinatario, asunto, operation_upper)
                        # le ponemos fecha y asunto a la bitácora y la insertamos
                        bitacora["asunto"] = asunto
                        bitacora["fecha"] = msg.date
                    else:
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
                    correo_respuesta(False, e, args.destinatario, asunto, tipo_operacion.upper())
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
            elif (
                msg.from_
                == EMAIL_MICROSOFT
            ):
                # los correos de buzón lleno, los eliminamos
                # mailbox.delete(f"{msg.uid}")
                pass
            else:
                try:
                    raise Exception(
                        "Sólo el o la subdelegada y el o la encargada de la subdelegación pueden enviar correos a esta dirección."
                    )
                except Exception as e:
                    asunto = msg.subject.strip()
                    correo_respuesta(False, e, args.destinatario, asunto, tipo_operacion.upper())
        else:
            try:
                asunto = (
                    msg.subject.replace("\r\n", "")
                    .split(" - Enviado por:")[0]
                    .split(":")[-1]
                    .strip()
                    .replace(" ", "")
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
                correo_respuesta_atencion(True, args.destinatario, asunto)
            except Exception as e:
                correo_respuesta_atencion(False, args.destinatario, asunto)
    else:
        pass
