# Uso: 
# python limpiar_carpeta.py --carpeta "CARPETA A VACIAR"
# si no se proporciona ningún argumento, se vacía la carpeta de elementos enviados

from helpers import limpiar_carpeta

import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--carpeta', '-c', help="Carpeta a despejar", type= str, nargs='?', const="Sent", default="Sent")
args=parser.parse_args()
#print(f"Args: {args}\nCarpeta: {args.carpeta}")
limpiar_carpeta(args.carpeta)
