# EmailProcessorBot

<a href="https://github.com/FernandoTorresL/EmailProcessorBot/commits/main" target="_blank">![GitHub last commit](https://img.shields.io/github/last-commit/FernandoTorresL/EmailProcessorBot)</a>

<a href="https://github.com/FernandoTorresL/EmailProcessorBot" target="_blank">![GitHub repo size](https://img.shields.io/github/repo-size/FernandoTorresL/EmailProcessorBot)</a>

## Proyecto interno para IMSS-DIR, México

> Automatically reads, process, saves and response emails from an Outlook mailbox.

Automáticamente lee, procesa, almacena y responde a correos de un buzón Outlook.

Valida correos recibidos en un buzón IMSS en particular, valida requisitos/contenido, destinatarios, etc., y lo reenvía a personal responsable para su atención, llevando un registro en una base de datos MongoDB.


## Introducción

Este es un proyecto interno para la _Dirección de Incorporación y Recaudación_ - _Coordinación de Afiliación_, una oficina del _Instituto Mexicano del Seguro Social_

El código original fue desarrollado por personal de la _USE_ y posteriormente entregado para su operación y mantenimiento a personal de la _Coordinación de Afiliación_

## Tecnología/Software utilizado

El proyecto fue construido usando:

- Python v3.11 (original)
- Python v3.12.3 (W10 - Pruebas)
- Python v3.12.3 (OS - Pruebas)
- [DavMail](https://davmail.sourceforge.net) (Requiere Java JDK en el equipo donde se instale. Ver [setup](https:davmail.sourceforge.net/windowssetup.html))

### Clonar el repositorio

```sh
git clone git@github.com:FernandoTorresL/EmailProcessorBot.git <my_folder> 
```
Opcionalmente puedes cambiar <my_folder> con el nombre de la carpeta de tu elección.

> Optional: You can change *<my_folder>* on this instruction to create a new folder


### Crea/actualiza los archivos iniciales (sólo hay ejemplos en el repositorio)

Debes crear y actualizar los siguientes archivos:

- A csv file "Destinatarios en CA.csv"
- A csv file "Directorio Nacional Subdelegados.csv"
- A python file "credenciales.py"

> Todos basados en sus versiones .example.py incluidos en este repositorio.

- Si no se detecta Java, deberá crear un archivo davmail64.ini y usar la línea siguiente con la ruta de exe de DavMail

```sh
vm.location=C:\path_to_java_folder\bin\server\jvm.dll
```

También debe crearse un certificado siguiendo las [instrucciones siguientes](https://davmail.sourceforge.net/sslsetup.html)

```sh
keytool -genkey -keyalg rsa -keysize 2048 -storepass password -keystore davmail.p12 -storetype
pkcs12 -validity 3650 -dname cn=davmailhostname.company.com,ou=davmail,o=sf,o=net
```
El archivo generado davmail.p12 se colocó en C:\Users\<usuario> junto con el archivo que guarda las propiedades de DavMail ".davmail.properties", el cual contiene el valor de la URL del servidor de Exchange y los puertos utilizados.

### Change to working directory and create a Python virtual environment
### Cambia al directorio de trabajo y crea un ambiente virtual de Python

OS X & Linux:

```sh
$ cd EmailProcessorBot
$ python3 -m venv ./venv
$ source ./venv/bin/activate
$ pip install --upgrade pip
$ pip3 install -r requirements.txt
(venv) $
```

Windows:
```sh
$ cd EmailProcessorBot
$ python -m venv venv
$ .\venv\Scripts\activate
$ pip3 install -r requirements.txt
(venv) $
```

Windows 10 with Git bash terminal:
```sh
$ cd EmailProcessorBot
$ python -m venv venv
$ source ./venv/Scripts/activate
$ pip3 install -r requirements.txt
(venv) $
```

Windows 10 with powershell terminal:
```sh
PS> cd EmailProcessorBot
PS> python -m venv venv
PS> .\.venv\Scripts\Activate.ps1
PS> pip3 install -r requirements.txt
(.venv) PS>
```

Windows 10 with WSL shell:
```sh
user@pc_name: cd EmailProcessorBot
user@pc_name: python3 -m venv venv
user@pc_name: source venv/bin/activate
user@pc_name: pip install --upgrade pip
user@pc_name: pip3 install -r requirements.txt
(venv) user@pc_name:
```

El prompt puede varias dependiendo de la versión de terminal o shell utilizada.

> This prompt may vary if you use another shell configuration, like pk10 or git bash

Para desactivar el ambiente:
> Later, to deactivate the virtual environment
OS X & Linux & Windows:

```sh
(venv) $ deactivate
$
```

## Ejecuta el proyecto

Antes de ejecutar el programa, debes crear y actualizar el contenido de los archivos siguientes:

> Before run this project, edit or copy the following files and update content with real values:

* credenciales.py
* Destinatarios en CA.csv
* Directorio Nacional Subdelegados.csv

Y podrás ejecutar el programa:
> Then, you can execute the program:

```sh
python3 main.py
```


## Contributing to this repo

1. [Fork this project](https://github.com/FernandoTorresL/EmailProcessorBot/fork)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

---

<div align="center">
    <a href="https://fertorresmx.dev/">
      <img height="150em" src="https://raw.githubusercontent.com/FernandoTorresL/FernandoTorresL/main/media/FerTorres-dev1.png">
  </a>
</div>



## Follow me 
[fertorresmx.dev](https://fertorresmx.dev/)

#### :globe_with_meridians: [Twitter](https://twitter.com/FerTorresMx), [Instagram](https://www.instagram.com/fertorresmx/): @fertorresmx
