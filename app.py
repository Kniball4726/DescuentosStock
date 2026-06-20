import time
from src.init import main as iniciarAplicacion
from colorama import Fore, Style, init
init(autoreset=True)
from src.helpers import borrarPantallas as bp
from src.init import *


if __name__ == "__main__":
    bp()
    print(Fore.GREEN + Style.BRIGHT + "Iniciando aplicación...")
    time.sleep(2)  # Simulate some startup time
    print(Fore.GREEN + Style.BRIGHT + "La aplicación está ahora en ejecución.")
    time.sleep(2)
    bp()
    iniciarAplicacion()

