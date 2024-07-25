# EmailProcessorBot
Automatically reads, process, saves and response emails from an Outlook mailbox
Automáticamente lee, procesa, almacena y responde a correos de un buzón Outlook

## Validador de correos CDA-IMSS (Afiliacion_Correos)
## Proyecto privado para IMSS-DIR, México

Valida correos recibidos en un buzón IMSS en particular, valida requisitos/contenido, destinatarios, etc., y lo reenvía a personal responsable para su atención.


## Introducción

Este es un proyecto interno y confidencial para la _Dirección de Incorporación y Recaudación_ - _Coordinación de Afiliación_, una oficnia del _Instituto Mexicano del Seguro Social_

## Tech utilizada

El proyecto fue construido con:

- Python v3.??
- Python v3.12.3 (W10 - Pruebas)

### Create/Copy initial files (only placeholder_file.txt on GitHub)

You must create and update some files on place:

- A csv file "Destinatarios en CA.csv"
- A csv file "Directorio Nacional Subdelegados.csv"
- A credenciales.py file based on credenciales.example.py


### Change to working directory and create a Python virtual environment

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

> This prompt may vary if you use another shell configuration, like pk10 or git bash

Later, to deactivate the virtual environment
OS X & Linux & Windows:

```sh
(venv) $ deactivate
$
```

```sh
usage: dry_run.py [-destinatario <email>]

positional arguments:
  email            email for testing

options:
  -d, --destinatario  testing with <email>

```
> If using another Python version try: python dry_run.py

## Run the project

Before run this project, edit or copy the following files and update content with real values:

* credenciales.py 
* Destinatarios en CA.csv
* Directorio Nacional Subdelegados.csv

Then, you can execute the program:

```sh
python main.py
```

### Example

```sh
python ??
```

### Example output

```sh
??
```

## Output files


## Contributing to this repo

1. Fork it (<http://github.com/FernandoTorresL/EmailProcessorBot/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

---
