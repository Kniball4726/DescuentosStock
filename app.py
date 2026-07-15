import os
import time
import traceback
from src.init import main as iniciarAplicacion
from colorama import Fore, Style, init
init(autoreset=True)
from src.helpers import borrarPantallas as bp

CARPETAS_REQUERIDAS = ["Canjes", "Descuentos", "Descontados", "Mayoristas", "Mercado Libre"]

def crear_carpetas():
    for carpeta in CARPETAS_REQUERIDAS:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)
            print(Fore.YELLOW + f"  Carpeta creada: {carpeta}/")

if __name__ == "__main__":
    try:
        bp()
        iniciarAplicacion()
    except Exception:
        tb = traceback.format_exc()
        try:
            with open("error.log", "w", encoding="utf-8") as f:
                f.write(tb)
        except Exception:
            pass
        print(Fore.RED + Style.BRIGHT + "Se produjo un error. Revisa el archivo error.log en la carpeta del ejecutable.")


