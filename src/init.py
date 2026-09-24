import sys
import time
from colorama import Fore, Style, init
init(autoreset=True)

try:
    from .helpers import borrarPantallas as bp
    from .funciones import descontarMayoristas, descontarCanjes, descontarMercado, guardarDescuentos
except ImportError:
    try:
        from src.helpers import borrarPantallas as bp
        from src.funciones import descontarMayoristas, descontarCanjes, descontarMercado, guardarDescuentos
    except ImportError:
        from helpers import borrarPantallas as bp
        from funciones import descontarMayoristas, descontarCanjes, descontarMercado, guardarDescuentos

def main():
    opcion = ""
    try:
        while opcion != "5":
            bp()
            print(Fore.CYAN + Style.BRIGHT + "Menú de opciones:\n")
            print("1.- Descontar Mayoristas")
            print("2.- Descontar Canjes")
            print("3.- Descontar Mercado libre")
            print("4.- Guardar Descuentos (Backup + Limpiar Plantilla)")
            print("5.- Salir")
            opcion = input(Fore.CYAN + Style.BRIGHT + "\nSeleccione una opción: ").strip()

            match opcion:
                case "1":
                    descontarMayoristas()
                case "2":
                    descontarCanjes()
                case "3":
                    descontarMercado()
                case "4":
                    guardarDescuentos()
                case "5":
                    print(Fore.RED + "\nsaliendo . . .")
                    time.sleep(2)
                    sys.exit(0)
    except KeyboardInterrupt:
        print(Fore.RED +"\n\nSaliendo de la aplicación...")
    except SystemExit:
        pass
    except Exception as e:
        print(f"\nOcurrió un error: {e}")
