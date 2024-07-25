# bueno, ya intenté varias y parece que la más sencilla es IMAPTOOLS
import imaplib
import socket
import ssl
import time
import traceback

from imap_tools import AND, A, MailBox, MailboxLoginError, MailboxLogoutError
# establecemos conexión con PyMongo
from pymongo import MongoClient

conn_nvo = MongoClient(
    IP_MONGO_CLIENT,
    username=USER_NAME_MONGO,
    password=PWD_MONGO,
    authsource="admin",
    authMechanism="SCRAM-SHA-256",
)
db2 = conn_nvo["afiliacion"]
col_solicitudes2 = db2["solicitudes"]
col_bitacora2 = db2["bitacora"]
col_errores = db2["errores_sol"]

import os

# claves de operacion válidas
cves_solicitud = [
    "CDA01",
    "CDA02",
    "CDA03",
    "CDA04",
    "CDA05",
    "CDA06",
    "CDA07",
    "CDA08",
    "MOTIVO1",
    "MOTIVO2",
    "MOTIVO7",
    "EP",
    "LAUDO",
    "IVRO",
    "TEC",
    "MOD32",
    "MOD33",
    "MOD40",
    "PTH",
    "PTI",
    "CDA00",
]
import email
import mimetypes
import re
import smtplib
import string
from collections import Counter
from datetime import datetime, timedelta, timezone
from email import encoders
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import make_msgid
from io import BytesIO, StringIO

import pandas as pd
import smbclient
from tqdm import tqdm
from unidecode import unidecode

from credenciales import (DOMINIO_MAILBOX, EMAIL_ADMINISTRADOR, EMAIL_DUDAS,
                          EMAIL_MAILBOX, EMAIL_MICROSOFT, FOLDER_MAILBOX,
                          IP_MAILBOX, IP_MONGO_CLIENT, IP_SMTP, MSG_LIMIT,
                          PASSWORD_MAILBOX, PATH_ARCHIVO, PORT_MAILBOX,
                          PORT_SMTP, PWD_MONGO, URL_CIRCULAR, USER_NAME_MONGO)
# archivo que contiene la lista de las subdelegaciones válidas
from cves_subdelegacion import cves_subdel

# path de ejecución, ACTUALIZAR
path = "."
# leemos el archivo con los remitentes
df_subdelegados = pd.read_csv(f"{path}/Directorio Nacional Subdelegados.csv")
subdelegados = df_subdelegados.Email.str.strip().tolist()
del df_subdelegados

smbclient.ClientConfig(username=EMAIL_MAILBOX, password=PASSWORD_MAILBOX)

# leemos el archivo que contiene los operadores por tipo de operación
df_usuarios = pd.read_csv(f"{path}/Destinatarios en CA.csv")

# mucha flexibilidad con la responsiva jajaj, con que tengan alguno de estos textos (quitando puntuación y espacios dobles) se las damos por buena
str_responsiva_1 = "de conformidad con el artículo 150 y 155"
str_responsiva_1 = (
    str_responsiva_1.translate(str.maketrans("", "", string.punctuation))
    .strip()
    .translate(str.maketrans("", "", string.punctuation))
    .strip()
)
str_responsiva_1_art = "de conformidad con el art 150 y 155 del reglamento"
str_responsiva_1_art = str_responsiva_1_art.translate(
    str.maketrans("", "", string.punctuation)
).strip()
str_responsiva_1_alt = "de conformidad con los artículos 150 y 155"
str_responsiva_1_alt = str_responsiva_1_alt.translate(
    str.maketrans("", "", string.punctuation)
).strip()

str_responsiva_2 = (
    "que se encuentra debidamente soportado con la documentación que se adjunta".strip()
)
str_responsiva_2 = str_responsiva_2.translate(
    str.maketrans("", "", string.punctuation)
).strip()


# función que valida el último elemento del asunto del correo (es decir, que sea un NSS, un RP o una fecha, según sea el caso)
# los regex están feos pero son eficientes
def regex_operaciones(tipo_operacion: str, value: str) -> bool:
    try:
        if tipo_operacion.upper() in [
            x for x in cves_solicitud if x not in ["EP", "MOD40"]
        ]:
            excepcion = "El NSS debe de tener 11 dígitos"
            res = re.match(r"^\d{11}$", value)
        elif tipo_operacion.upper() == "EP":
            excepcion = f"El registro patronal debe de tener 8 o 10 caracteres. El registro patronal provisto es {value}"
            res = (
                re.match(r"^[a-zA-Z0-9]{8}$", value)
                or re.match(r"^[a-zA-Z0-9]{10}$", value)
                or re.match(r"^[a-zA-Z0-9]{11}$", value)
            )
        elif tipo_operacion.upper() == "MOD40":
            excepcion = "La fecha debe de estar en el formato dd/mm/yyyy"
            res = re.match(r"^(0[1-9]|[1-2][0-9]|3[0-1])/(0[1-9]|1[0-2])/\d{4}$", value)
        if res:
            return True, None
        else:
            raise Exception(excepcion)
    except Exception as e:
        return False, e


# función que valida el asunto
def validar_asunto(asunto: str) -> bool:
    try:
        asunto_split = asunto.split("-")
        # validamos componente por componente, primero vemos que al separar por guiones el asunto, éste tenga 4 elementos, si no, tiramos excepción
        if len(asunto_split) != 4:
            raise Exception("Asunto inválido. Consulte el formato.")
        else:
            # si pasamos el primer check, validamos que los primeros dos elementos resulten en una cve de subdelgación válida
            cve_compue = (
                f"{str(asunto_split[0]).zfill(2)}{str(asunto_split[1]).zfill(2)}"
            )
            if cve_compue not in cves_subdel:
                raise Exception(
                    "La combinación de cve_delegacion y cve_subdelegacion proporcionadas no forman una subdelegación válida."
                )
            else:
                # validamos que el tercer elemento sea un tipo de operación válida
                if asunto_split[2].upper() not in cves_solicitud:
                    raise Exception(
                        f"{asunto_split[2]} no es un tipo de operación válida"
                    )
                # finalmente, revisamos que el cuarto elemento sea del tipo correcto (según el tipo de operación), si todo bien, regresamos True
                else:
                    return regex_operaciones(asunto_split[2], asunto_split[3])
    # si hubo alguna excepción, enviamos False y la excepción (para reportarla al solicitante)
    except Exception as e:
        return False, e


# ciertas operaciones tienen requisitos específicos para el cuerpo
# (por ejemplo, una tabla con ciertas columnas), esta función valida que, según el tipo de operación, eso se cumpla
def validar_ops(cuerpo: str, tipo_operacion: str) -> [bool, str]:
    try:
        # primero, le quitamos acentos, puntuación y mayúsculas al cuerpo del correo y al tipo de operación
        cuerpo = unidecode(
            cuerpo.lower().translate(str.maketrans("", "", string.punctuation)).strip()
        ).replace("\r\n", " ")
        tipo_operacion = tipo_operacion.lower()
        # ya aquí según el cuerpo de operación pedimos que existan los elementos necesarios
        # si existen, regresamos True y None, si no, False y la excepción
        if tipo_operacion in ["cda07", "motivo2", "motivo7"]:
            if (
                ("ciz" in cuerpo or "cicz" in cuerpo)
                and "nss" in cuerpo.replace(".", "")
                and (
                    ("registro" in cuerpo and "patronal" in cuerpo)
                    or "reg patronal" in cuerpo
                    or "rp" in cuerpo
                    or "registro pat" in cuerpo
                    or "reg pat" in cuerpo
                )
                and "dice" in cuerpo
                and ("debe" in cuerpo and "decir" in cuerpo)
            ):
                if ( ("tc11" in cuerpo) or ("tc 11" in cuerpo) or ("tc-11"in cuerpo) ):
                    raise Exception(
                        "No puede referirse a TC11 en las solicitudes CDA07. Revise su solicitud y/o CIZ en tabla"
                    )
                else:
                    return True, None
            else:
                raise Exception(
                    "Debe de incluir en el cuerpo una tabla con las siguientes columnas: CIZ, NSS, REGISTRO PATRONAL, DICE, DEBE DECIR. La tabla no puede ser una imagen."
                )
        elif tipo_operacion == "motivo1":
            cuerpo_limpio = cuerpo.replace(" ", "").replace("\t", "")
            if (
                ("ciz" in cuerpo or "cicz" in cuerpo)
                and "nss" in cuerpo.replace(".", "")
                and (
                    ("registro" in cuerpo and "patronal" in cuerpo)
                    or "reg patronal" in cuerpo
                    or "rp" in cuerpo
                    or "registro pat" in cuerpo
                    or "reg pat" in cuerpo
                )
                and "dice" in cuerpo
                and ("debe" in cuerpo and "decir" in cuerpo)
                and "baja" in cuerpo
                and "causa" in cuerpo
                and (
                    "bajadelrpcausadebaja" in cuerpo_limpio
                    or "bajadelrpcausadelabaja" in cuerpo_limpio
                    or "bajarpcausadebaja" in cuerpo_limpio
                    or "bajadelregistropatronalcausadelabaja" in cuerpo_limpio
                    or "bajadelregistropatronalcausadebaja" in cuerpo_limpio
                )
            ):
                return True, None
            else:
                raise Exception(
                    "Debe de incluir en el cuerpo una tabla con las siguientes columnas: CIZ, NSS, REGISTRO PATRONAL, BAJA DEL RP, CAUSA DE BAJA DEL RP, DICE, DEBE DECIR. La tabla no puede ser una imagen."
                )
        elif tipo_operacion == "mod40":
            cuerpo_limpio = cuerpo.replace(".", "").replace(" ", "").replace("\t", "")
            if (
                "nss" in cuerpo_limpio
                and ("tipodemov" in cuerpo_limpio or "tipomov" in cuerpo_limpio)
                and "salario" in cuerpo
                and ("ciz" in cuerpo or "cicz" in cuerpo or "cisz" in cuerpo)
            ):
                return True, None
            else:
                raise Exception(
                    "Debe de incluir en el cuerpo una tabla con las siguientes columnas: NSS, Tipo de movimiento Reingreso o baja, SALARIO, CIZ 1, 2 o 3"
                )
    except Exception as e:
        return False, e


# esta función simplemente combina la validación de la nota responsiva con la validación de la tabla contenida en el cuerpo - según el tipo de solicitud
def validar_cuerpo_correo(cuerpo: str, tipo_operacion: str) -> bool:
    try:
        if (
            (
                unidecode(" ".join(str_responsiva_1.lower().split()))
                not in unidecode(
                    " ".join(
                        cuerpo.replace("“", "")
                        .replace("”", "")
                        .replace("  ", " ")
                        .translate(str.maketrans("", "", string.punctuation))
                        .lower()
                        .strip()
                        .split()
                    )
                )
            )
            and (
                unidecode(" ".join(str_responsiva_1_art.lower().split()))
                not in unidecode(
                    " ".join(
                        cuerpo.replace("“", "")
                        .replace("”", "")
                        .replace("  ", " ")
                        .translate(str.maketrans("", "", string.punctuation))
                        .lower()
                        .strip()
                        .split()
                    )
                )
            )
            and (
                unidecode(" ".join(str_responsiva_1_alt.lower().split()))
                not in unidecode(
                    " ".join(
                        cuerpo.replace("“", "")
                        .replace("”", "")
                        .replace("  ", " ")
                        .translate(str.maketrans("", "", string.punctuation))
                        .lower()
                        .strip()
                        .split()
                    )
                )
            )
        ):
            raise Exception("No se incluyó la nota responsiva en el cuerpo")
        else:
            if tipo_operacion in ["CDA07", "MOTIVO1", "MOTIVO2", "MOTIVO7", "MOD40"]:
                result, excepcion = validar_ops(cuerpo, tipo_operacion)
                return result, excepcion
            else:
                return True, None
    except Exception as e:
        return False, e


# columnas que DEBE tener la bitácora
columnas_bitacora = [
    "CLAVE_OOAD",
    "CLAVE_SUBDELEGACIÓN",
    "CLAVE_DE_TIPO_DE_SOLICITUD",
    "CONSECUTIVO_5_POSICIONES",
    "NSS_11_POSICIONES",
    "PRIMER_APELLIDO",
    "SEGUNDO_APELLIDO",
    "NOMBRE",
    "LA MODIFICACIÓN SOLICITADA AFECTA PERIODOS CON FECHA PREVIA AL 1 DE JULIO DE 1997",
]


# función para validar la forma de la bitácora cuando se envía una carpeta compartida
def validar_bitacora_smb(cuerpo: str) -> pd.DataFrame:
    # primero, buscamos la IP con este regex
    ip_pattern = r"\b(?:\d{1,3}\.){3}\d{1,3}\b"
    match = re.search(ip_pattern, cuerpo)
    # genermaos la URL
    url = f"{match.group()}{cuerpo.split(match.group())[1].split(' ')[0].split('<')[0]}"
    # abrimos conexión con Samba
    archivos_smb = smbclient.listdir(rf"\\{url}")
    # buscamos en la carpeta el archivo que empiece con BCA, BCP o BC40 y sea formato excel (.xls o .xlsx)
    bitacora = [
        x
        for x in archivos_smb
        if ("BCA" in x or "BCP" in x or "BC40" in x) and (".xls" in x)
    ]
    # SÓLO DEBE HABER 1 ARCHIVO CON ESE NOMBRE EN LA CARPETA, SI NO, MANDAMOS ERROR
    if len(bitacora) == 1:
        with smbclient.open_file(
            rf"\\{url}\\{bitacora[0]}", mode="rb", encoding="ISO-8859-1"
        ) as fd:
            bitacora = pd.read_excel(fd)
            # leemos la bitácora y validamos que las columans sean idénticas a las del formato (importa el orden)
            # y que tenga al menos 1 fila de información (si solo mandan un archivo con los encabezados, se les regresa)
        if bitacora.columns.tolist() == columnas_bitacora and not bitacora.empty:
            # si cumple todo, regresamos la bitácora (para después subirla a la BD)
            return bitacora
        else:
            raise Exception(
                "O la bitácora está vacía o no tiene las mismas columnas que el formato original."
            )
    else:
        raise Exception(
            "No encontramos la bitácora en la carpeta (recuerde que el nombre debe empezar con BCA, BCP o BC40, según el tipo de solicitud) o no pudimos acceder al recurso compartido."
        )


nota_carpeta = "No debe de combinar más de un anexo en un solo archivo. Si en conjunto los archivos pesan más de 10MB incluír una carpeta con el mismo nombre que el asunto del correo."


def validar_anexos(
    attachments: list, tipo_operacion: str, asunto: str, cuerpo: str
) -> bool:
    # quitamos el archivo image001 de los anexos (este corresponde a la firma)
    attachments = [
        x for x in attachments if ("inline" not in x.content_disposition)
    ]
    attachments = [
        x for x in attachments if ("image00" not in x.filename.split(".")[0])
    ]
    try:
        if (
            len(
                [
                    x.filename
                    for x in attachments
                    if (
                        "BCA" in x.filename.upper()
                        or "BCP" in x.filename.upper()
                        or "BC40" in x.filename.upper()
                    )
                    and (".xls" in x.filename.lower())
                ]
            )
            == 1
        ):
            nomb_archivo = [
                x.filename
                for x in attachments
                if (
                    "BCA" in x.filename.upper()
                    or "BCP" in x.filename.upper()
                    or "BC40" in x.filename.upper()
                )
                and (".xls" in x.filename.lower())
            ]
            # tipo_archivos = Counter([x.content_type for x in attachments])
            # print(nomb_archivo)
            # leemos la bitácora (sólo debe haber UN archvio excel que empiece con BCA, BCP o BC40 en el correo, si hay más o menos de 1, rechazamos
            bitacora = pd.read_excel(
                BytesIO(
                    [
                        x.payload
                        for x in attachments
                        if (
                            "BCA" in x.filename.upper()
                            or "BCP" in x.filename.upper()
                            or "BC40" in x.filename.upper()
                        )
                        and (".xls" in x.filename.lower())
                    ][0]
                )
            )
            #if (bitacora.shape[0] > 10) or (bitacora.shape[1] > 9):
            print(f"Tamaño bitacora para {asunto}:{bitacora.shape}")
            if bitacora.shape[0] > 49:
               excepcion = "La bitácora parece exceder de tamaño. No incluir toda la historia de su bitácora, enviar sólo el renglón correspondiente a la petición de su correo. O puede ser que inadvertidamente haya agregado/modificado renglones/columnas vacías. Revise la cantidad de filas o columnas."
               print(f"Excepcion para {asunto}: {excepcion}")
               raise Exception(excepcion)
            # Si es un CDA07, validamos que el NSS del asunto esté incluido en el nombre del Excel/bitacora
            if tipo_operacion.lower() == "cda07":
                if asunto.split("-")[3] not in nomb_archivo[0]:
                    excepcion = "El NSS contenido en el asunto no es el mismo indicado en el nombre de la bitácora"
                    print(f"Excepcion para {asunto}: {excepcion}")
                    raise Exception(excepcion)
            # validamos columnas y que la bitácora tenga al menos una fila de información
            if bitacora.columns.tolist() != columnas_bitacora or bitacora.empty:
                excepcion = "O la bitácora está vacía o no tiene las columnas del formato original. Revise también que no haya incluido columnas adicionales al final (incluso vacías)"
                raise Exception(excepcion)
        else:
            excepcion = "No adjuntó la bitácora a su solicitud o incluyó más de una bitácora o bien, ésta no tiene el nombre correcto. Recuerde que, según el tipo de operación, el nombre de la bitácora debe empezar con BCA, BCP o BC40 (consulte la circular para más detalles)."
            raise Exception(excepcion)
        # acá va la paarte tediosa de esta función, para cada tipo de operación, debemos validar que la suma de anexos por tipo de archivo sea la indicada en la circular
        # dada la extensión, no comentaré, cualquier cosa, escribir al administrador
        if tipo_operacion.lower() in ["cda01", "cda02"]:
            if len(attachments) >= 3:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    == 1
                    and tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 2
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg) y la solicitud de regularización y/o corrección en formado .docx o .pdf"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "\\" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 3 anexos, la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg) y la solicitud de regularización y/o corrección en formado .docx o .pdf. {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "cda03":
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 2
                    and tipo_archivos.get(
                        "application/vnd.ms-excel.sheet.macroenabled.12",
                        0,
                    )
                    == 1
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg), la solicitud de regularización y/o corrección en formado .docx o .pdf y el documento Solicitud Eliminación MPC Nombre DelXX SubdelYY aaaammdd.xlsm debidamente requisitado (en formato .xlsm)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = "Debe incluir 4 anexos la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg), la solicitud de regularización y/o corrección en formado .docx o .pdf y el documento Solicitud Eliminación MPC Nombre DelXX SubdelYY aaaammdd.xlsm debidamente requisitado (en formato .xlsm)"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() in ["cda04", "cda07"]:
            if len(attachments) >= 6:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and (
                        tipo_archivos.get(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            0,
                        )
                        + tipo_archivos.get("image/jpeg", 0)
                        + tipo_archivos.get("application/pdf", 0)
                        + tipo_archivos.get(
                            "image/png",
                            0,
                        )
                    )
                    >= 5
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg), la solicitud de regularización y/o corrección (en formado .docx o .pdf), así como CURP, acta de nacimiento e identificación oficial (por separado cada una en formato .pdf)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 6 archivos, la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada (en PDF o imágen .jpg), la solicitud de regularización y/o corrección (en formado .docx o .pdf), así como CURP, acta de nacimiento e identificación oficial (por separado cada una en formato .pdf). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "cda05":
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get("application/pdf", 0)
                    + tipo_archivos.get("image/jpeg", 0)
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 3
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, la captura de pantalla de la cuenta individual original certificada (en un archivo pdf o una imagen .jpg), la solicitud de regularización y/o corrección en formato .docx o pdf y el documento que justifique la eliminación (como archivo PDF)."
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe anexar 4 archivos, la bitácora de control en formato excel, la captura de pantalla de la cuenta individual original certificada (en un archivo pdf o una imagen .jpg), la solicitud de regularización y/o corrección en formato .docx o pdf y el documento que justifique la eliminación (como archivo PDF). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "cda06":
            if len(attachments) >= 6:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get("image/jpeg", 0)
                    + tipo_archivos.get("application/pdf", 0)
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 5
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir un archivo en formato excel, la bitácora de control, la solicitud de regularización y/o corrección en formato word o PDF y la CURP, identificación oficial, acta de nacimiento y aviso afiliatorio cotejado con el original y/o el formato 1073-33 debidamente requisitado (los 4 en formato pdf)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 6 archivos anexos, la bitácora de control, la solicitud de regularización y/o corrección en formato word o PDF y la CURP, identificación oficial, acta de nacimiento y aviso afiliatorio cotejado con el original y/o el formato 1073-33 debidamente requisitado (los 4 en formato pdf). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "cda08":
            if len(attachments) >= 6:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and (
                        tipo_archivos.get(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            0,
                        )
                        + tipo_archivos.get("image/jpeg", 0)
                        + tipo_archivos.get("application/pdf", 0)
                        + tipo_archivos.get(
                            "image/png",
                            0,
                        )
                    )
                    >= 5
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada donde se indiquen los periodos a regularizar (en PDF o imágen .jpg), la solicitud de regularización y/o corrección (en formado .docx o .pdf), así como CURP, acta de nacimiento e identificación oficial (por separado en formato .pdf)."
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 6 archivos, la bitácora de control en formato excel, la captura de pantalla de la cuenta individual certificada donde se indiquen los periodos a regularizar (en PDF o imágen .jpg), la solicitud de regularización y/o corrección (en formado .docx o .pdf), así como CURP, acta de nacimiento e identificación oficial (por separado cada una en formato .pdf). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "motivo1":
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if tipo_archivos.get(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    0,
                ) + tipo_archivos.get(
                    "application/vnd.ms-excel",
                    0,
                ) >= 1 and (
                    tipo_archivos.get("image/jpeg", 0)
                    + tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get("application/pdf", 0)
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 3
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir la bitácora de control en formato excel, las capturas de pantalla de SINDO del F7 y F3 (2 imágenes) y la cuenta individual filtrada por registro patronal original certificada (en formato pdf o como imagen)."
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 4 anexos, la bitácora de control en formato excel, las capturas de pantalla de SINDO del F7 y F3 (2 imágenes) y la cuenta individual filtrada por registro patronal original certificada (en formato pdf o como imagen). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "motivo2":
            if len(attachments) >= 2:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                ) and (
                    tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 1
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir dos archivos, la cuenta individual filtrada por registro patronal original certificada del caso (en formato PDF o como imagen) y la bitácora de control en formato excel"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir dos archivos, la cuenta individual filtrada por registro patronal original certificada del caso (en formato PDF o como imagen) y la bitácora de control en formato excel. {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "ep":
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    >= 2
                    and tipo_archivos.get(
                        "text/plain",
                        0,
                    )
                    >= 1
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir un archivo en formato excel (bitácora de control), 2 pdf's o imágenes (Oficio de petición a la División de Afiliación a Régimen Obligatorio y Escrito Patronal) y un archivo DISPMAG.txt por cada registro patronal (al menos 1)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir al menos 4 archivos adjuntos, un archivo en formato excel (bitácora de control), 2 pdf's o imágenes (Oficio de petición a la División de Afiliación a Régimen Obligatorio y Escrito Patronal) y un archivo .txt por cada registro patronal (al menos 1). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "laudo":
            if len(attachments) >= 5:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 3
                    and tipo_archivos.get(
                        "text/plain",
                        0,
                    )
                    >= 1
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir un archivo en formato excel (bitácora de control), 3 pdf's o imágenes (Oficio de petición a la División de Afiliación a Régimen Obligatorio, Reporte de juicio concluído o instrucción en materia de afiliación de la Jefatura de Servicios Jurídicos y certificación de pagos, emitida por la jefatura de departamento de cobranza) y un archivo .txt (DISPMAG.txt)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 5 archivos, un archivo en formato excel (bitácora de control), 3 pdf's o imágenes (Oficio de petición a la División de Afiliación a Régimen Obligatorio, Reporte de juicio concluído o instrucción en materia de afiliación de la Jefatura de Servicios Jurídicos y certificación de pagos) y un archivo .txt (DISPMAG.txt). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() in ["ivro", "mod32", "mod33"]:
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 2
                    and tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 2
                    and (
                        len(
                            [
                                x.filename
                                for x in attachments
                                if "modalidad" in x.filename.lower()
                                and "ivro" in x.filename.lower()
                            ]
                        )
                        == 1
                    )
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir 2 archivo en formato excel la bitácora de control, el archivo denominado Modalidad_IVRO_MOD_32_33_35_43_44_OOAD_XX_SUB_XX_ddmmyyyy, así como el comprobante de pago y el formato AFIL-05en imagen o PDF"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 4 archivos: 2 archivo en formato excel la bitácora de control, el archivo denominado Modalidad_IVRO_MOD_32_33_35_43_44_OOAD_XX_SUB_XX_ddmmyyyy, así como el comprobante de pago y el formato AFIL-05en imagen o PDF. {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "tec":
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    >= 1
                    and tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    >= 2
                    and tipo_archivos.get(
                        "text/plain",
                        0,
                    )
                    >= 1
                ):
                    return True, None, bitacora
                else:
                    excepcion = "Debe incluir un archivo en formato excel (bitácora de control), 2 pdf's o imágenes (Oficio de petición a la División de Incorporación Voluntaria y Convenios y el Escrito Patronal) y un archivo .txt (DISPMAG.txt)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 4 archivos, un archivo en formato excel (bitácora de control), 2 pdf's o imágenes (Oficio de petición a la División de Incorporación Voluntaria y Convenios y el Escrito Patronal) y un archivo .txt (DISPMAG.txt). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() in ["pth", "pti"]:
            if len(attachments) >= 4:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.ms-excel",
                        0,
                    )
                    == 1
                    and tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/octet-stream",
                        0,
                    )
                    >= 3
                ):
                    return True, None, bitacora
                else:
                    excepcion = f"Debe incluir un archivo en formato excel (bitácora de control) y 3 pdf's o imágenes (formato de incorporación de {tipo_operacion.upper()}, línea de captura y comprobante de pago)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir 4 archivos, un archivo en formato excel (bitácora de control) y 3 pdf's o imágenes (formato de incorporación de {tipo_operacion.upper()}, línea de captura y comprobante de pago). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        elif tipo_operacion.lower() == "mod40":
            return True, None, bitacora
        elif tipo_operacion.lower() == "motivo7":
            return True, None, bitacora
        elif tipo_operacion.lower() == "cda00":
            if len(attachments) >= 3:
                tipo_archivos = Counter([x.content_type for x in attachments])
                if (
                    tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        0,
                    )
                    + tipo_archivos.get("application/vnd.ms-excel", 0)
                    >= 1
                    and tipo_archivos.get(
                        "image/jpeg",
                        0,
                    )
                    + tipo_archivos.get(
                        "image/png",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/msword",
                        0,
                    )
                    + tipo_archivos.get(
                        "application/pdf",
                        0,
                    )
                    >= 2
                ):
                    return True, None, bitacora
                else:
                    excepcion = f"Debe incluir un archivo en formato excel (bitácora de control) y 2 imágenes (una captura de la bitácora de solicitudes y una de la bandeja de solicitudes)"
                    raise Exception(excepcion)
            else:
                if (
                    f"{asunto.lower()}" not in cuerpo.lower()
                    or "file://" not in cuerpo.lower()
                ):
                    excepcion = f"Debe incluir un archivo en formato excel (bitácora de control) y 2 imágenes (una captura de la bitácora de solicitudes y una de la bandeja de solicitudes). {nota_carpeta}"
                    raise Exception(excepcion)
                else:
                    return True, None, bitacora
        else:
            excepcion = f"Tipo de operación {tipo_operacion} inválida. Revise el título de su correo."
            raise Exception(excepcion)
        # ya validamos tipo y número de anexos, ahora validamos en otra función la bitácora
    # si no hay bitácora o no cumple con requisitos, vemos si tiene carpeta compartida
    except Exception as e:
        # si hay una IP en el cuerpo
        ip_pattern = r"\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}\b"
        match = re.search(ip_pattern, cuerpo)
        if match:
            try:
                bitacora = validar_bitacora_smb(msg.text)
                return True, None, bitacora
            except Exception as e:
                return False, e, None
        else:
            return False, e, None


# función que envía correo de respuesta al subdelegado, bien informando que se envío a proceso o que se rechazó su solicitud (explicando las razones de rechazo)
def correo_respuesta(
    exito: bool, excepciones: str, destinatario: str, asunto_original: str
) -> None:
    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_MAILBOX
    sender = EMAIL_MAILBOX
    msg["To"] = destinatario
    receiver = destinatario
    if not exito:
        msg["Subject"] = f"Error al procesar - {asunto_original}"

        cuerpo_correo = f"""
            Estimado o estimada,
            Recibimos su correo con el identificador {asunto_original}, éste no pudo ser procesado, debido a las siguientes razones:
            {excepciones}

            Quedamos a la espera de las correcciones para poder procesar el movimiento.
        """
        ruta_Circular = URL_CIRCULAR
        html = f"""\
        <html>
        <body>
                <p>Estimado o estimada,<br></p>
                <p>Recibimos su correo con el identificador {asunto_original}, éste no pudo ser procesado, debido a las siguientes razones:</p>
                <ol>{excepciones}</ol>
                <p>Puede <a href={ruta_Circular}> descargar aquí el oficio circular y las plantillas</a> con las instrucciones a seguir.</p>
                <p>Si tiene dudas, puede escribir un correo a {EMAIL_DUDAS}</p>
                <p>
                    Cordiales saludos.
                </p>

            </body>
        </html>
        """
        part2 = MIMEText(html, "html")
        msg.attach(part2)

    elif exito:
        msg["Subject"] = f"Éxito al procesar - {asunto_original}"

        cuerpo_correo = f"""
            Estimado o estimada,
            Recibimos su correo con el identificador {asunto_original}, éste fue enviado a revisión de manera exitosa.

        """
        html = f"""\
        <html>
        <body>
                <p>Estimado o estimada,<br></p>
                <p>Recibimos su correo con el identificador {asunto_original}, éste fue enviado a revisión de manera exitosa.</p>
                <p>
                    Cordiales saludos.
                </p>
            </body>
        </html>
        """
        part2 = MIMEText(html, "html")
        msg.attach(part2)

    with smtplib.SMTP(IP_SMTP, PORT_SMTP) as server:
        server.login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
        server.sendmail(sender, receiver, msg.as_string())


# correo que confirma a le operadore que se marcó como atendida la solicitud indicada
def correo_respuesta_atencion(exito, destinatario, asunto):
    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_MAILBOX
    sender = EMAIL_MAILBOX
    msg["To"] = destinatario
    receiver = destinatario
    if not exito:
        msg["Subject"] = f"Error al marcar como atendido - {asunto}"

        html = f"""
            <p>Estimado o estimada,</p>
            <p>Recibimos su correo con el identificador {asunto}, éste no pudo ser marcado como atendido. O bien no pudimos identificar un ID válido en el asunto o esta solicitud ya había sido marcada como completa.</p>    
            <br><p>Quedamos a la espera de las correcciones para poder procesar el movimiento.</p><br><br><br>
        """
    elif exito:
        msg["Subject"] = f"Éxito al marcar como atendido - {asunto}"

        html = f"""
            Estimado o estimada,
            Recibimos su correo con el identificador {asunto}, éste fue marcado exitosamente como atendido.
<br>
            Gracias y buen día.<br><br><br>
            """
    part2 = MIMEText(html, "html")
    msg.attach(part2)
    with smtplib.SMTP(IP_SMTP, PORT_SMTP) as server:
        server.login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
        server.sendmail(sender, receiver, msg.as_string())


# función para reenviar las solicitudes a le operadore asignade
def correo_atender(
    mensaje: EmailMessage, id_mensaje: str, tipo_operacion: str, enviado_por: str
):
    df_responsable = (
        df_usuarios.loc[df_usuarios.cve_solicitud.str.lower() == tipo_operacion.lower()]
        .reset_index(drop=True)
        .iloc[0]
    )
    responsable = df_responsable["responsable"]
    msg = mensaje
    msg["CC"] = [df_responsable["ccp_1"], df_responsable["ccp_2"]]
    for key in [
        x
        for x in msg.keys()
        if x not in ["Subject", "Content-Type", "MIME-Version", "X-MS-Has-Attach"]
    ]:
        del msg[key]
    msg["From"] = EMAIL_MAILBOX
    sender = EMAIL_MAILBOX
    msg["To"] = f"{responsable},{df_responsable['ccp_1']},{df_responsable['ccp_2']}"
    asunto_orig = msg["Subject"]
    del msg["Subject"]
    msg["Subject"] = f"A atender: {id_mensaje} - Enviado por: {enviado_por}"
    del msg["Received"]
    receiver = [responsable, df_responsable["ccp_1"], df_responsable["ccp_2"]]
    with smtplib.SMTP(IP_SMTP, PORT_SMTP) as server:
        server.login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
        server.sendmail(sender, receiver, msg.as_string())


def limpiar_carpeta(carpeta: str = "Sent"):
    mailbox = MailBox(
        "localhost", port=PORT_MAILBOX, ssl_context=ssl._create_unverified_context()
    ).login(EMAIL_MAILBOX, PASSWORD_MAILBOX)
    mailbox.folder.set(carpeta)
    enviados = [x.uid for x in mailbox.fetch()]
    mailbox.delete(enviados)
