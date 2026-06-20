import os
import time
from src.init import main as iniciarAplicacion
from colorama import Fore, Style, init
init(autoreset=True)
from src.helpers import borrarPantallas as bp
from src.init import *

CARPETAS_REQUERIDAS = ["Canjes", "Descuentos", "Mayoristas", "Mercado Libre"]

def crear_carpetas():
    for carpeta in CARPETAS_REQUERIDAS:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)
            print(Fore.YELLOW + f"  Carpeta creada: {carpeta}/")

if __name__ == "__main__":
    bp()
    print(Fore.GREEN + Style.BRIGHT + "Iniciando aplicación...")
    crear_carpetas()
    time.sleep(2)
    print(Fore.GREEN + Style.BRIGHT + "La aplicación está ahora en ejecución.")
    time.sleep(2)
    bp()
    iniciarAplicacion()


