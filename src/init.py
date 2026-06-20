from src.helpers import borrarPantallas as bp
from src.funciones import *
import time
from colorama import Fore, Style, init
init(autoreset=True)

def main():
    opcion = ""
    try:
        while opcion != "4":
            print(Fore.CYAN + Style.BRIGHT + "Menú de opciones:\n")
            print("1.- Descontar Mayoristas")
            print("2.- Descontar Canjes")
            print("3.- Descontar Mercado libre")
            print("4.- Salir")
            opcion = input(Fore.CYAN + Style.BRIGHT + "\nSeleccione una opción: ").strip()

            match opcion:
                case "1":
                    descontarMayoristas()
                case "2":
                    descontarCanjes()
                case "3":
                    descontarMercado()
                case "4":
                    print(Fore.RED + "\nsaliendo . . .")
                    time.sleep(2)
                    exit()
    except KeyboardInterrupt:
        print("\nSaliendo de la aplicación...")
    except Exception as e:
        print(f"\nOcurrió un error: {e}")
