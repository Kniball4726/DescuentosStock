import os
import time
import pdfplumber
import pandas as pd
from colorama import Fore, Style
try:
    from src.helpers import borrarPantallas as bp
except ModuleNotFoundError:
    from helpers import borrarPantallas as bp

def descontarMayoristas():
    bp()
    print(Fore.CYAN + Style.BRIGHT + "=== PROCESANDO PEDIDOS MAYORISTAS ===\n")
    
    folder = "Mayoristas"
    plantilla_path = "Plantilla.xlsx"
    
    if not os.path.exists(folder):
        print(Fore.RED + f"Error: La carpeta '{folder}' no existe.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return
        
    if not os.path.exists(plantilla_path):
        print(Fore.RED + f"Error: El archivo '{plantilla_path}' no existe.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
    if not pdf_files:
        print(Fore.YELLOW + f"No se encontraron archivos PDF en la carpeta '{folder}'.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    print(Fore.BLUE + f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")
    
    all_products = {}
    
    # 1. Procesar cada PDF
    for file_name in pdf_files:
        file_path = os.path.join(folder, file_name)
        print(Fore.WHITE + f"Procesando: {file_name}...")
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    lines = text.split("\n")
                    in_detalle = False
                    for line in lines:
                        line_stripped = line.strip()
                        if "DETALLE" in line_stripped:
                            in_detalle = True
                            continue
                        if in_detalle and ("Cant. Articulos:" in line_stripped or "OBSERVACIONES" in line_stripped):
                            in_detalle = False
                        
                        if in_detalle:
                            tokens = line_stripped.split()
                            if not tokens:
                                continue
                            # Las líneas de producto empiezan con un código numérico y terminan con una cantidad numérica
                            if tokens[0].isdigit() and tokens[-1].isdigit():
                                code = int(tokens[0])
                                qty = int(tokens[-1])
                                all_products[code] = all_products.get(code, 0) + qty
        except Exception as e:
            print(Fore.RED + f"  Error al leer {file_name}: {e}")

    if not all_products:
        print(Fore.YELLOW + "No se extrajo ningún producto de los archivos PDF.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    print(Fore.GREEN + f"\nExtracción completa. Se encontraron {len(all_products)} productos distintos.")

    # 2. Cargar la plantilla
    print(Fore.WHITE + f"Cargando {plantilla_path}...")
    try:
        # Intentar cargar como XLSX estándar primero
        df = pd.read_excel(plantilla_path, engine='openpyxl')
    except Exception:
        try:
            # Reintento con xlrd (formato antiguo OLE/XLS)
            df = pd.read_excel(plantilla_path, engine='xlrd')
        except Exception as e:
            print(Fore.RED + f"Error al leer la plantilla: {e}")
            input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
            return

    # Limpiar columna Descuento
    df['Descuento'] = None

    # 3. Emparejar y actualizar
    descuentos_actualizados = 0
    total_descuento_unidades = 0

    for idx, row in df.iterrows():
        code_val = row['CODIGO']
        if pd.notna(code_val):
            # Ignorar la fila de totalizador
            if str(code_val).strip().upper() == 'TOTAL UNIDADES':
                continue
            try:
                # Convertir a int de forma segura (soporta floats como 330.0)
                code_int = int(float(str(code_val).strip()))
                if code_int in all_products:
                    qty = all_products[code_int]
                    df.at[idx, 'Descuento'] = qty
                    descuentos_actualizados += 1
                    total_descuento_unidades += qty
            except (ValueError, TypeError):
                # Ignorar encabezados de categoría u otros textos
                pass

    # 4. Actualizar gran total en la fila de 'TOTAL UNIDADES'
    total_row_mask = df['CODIGO'].astype(str).str.strip().str.upper() == 'TOTAL UNIDADES'
    if total_row_mask.any():
        df.loc[total_row_mask, 'Descuento'] = total_descuento_unidades
        df.loc[total_row_mask, 'PRODUCTOS'] = 'TOTAL PRODUCTOS'
        print(Fore.GREEN + f"Fila de 'TOTAL UNIDADES' actualizada con {total_descuento_unidades} unidades en Descuento y texto 'TOTAL PRODUCTOS'.")
    else:
        print(Fore.YELLOW + "Advertencia: No se encontró la fila 'TOTAL UNIDADES' en la plantilla.")

    # 5. Guardar archivo como XLSX moderno
    print(Fore.WHITE + "Guardando cambios en Plantilla.xlsx...")
    try:
        df.to_excel(plantilla_path, index=False, engine='openpyxl')
        print(Fore.GREEN + Style.BRIGHT + f"\n¡Éxito! Se actualizaron {descuentos_actualizados} productos en la columna 'Descuento'.")
        print(Fore.GREEN + f"Total de unidades descontadas: {total_descuento_unidades}")
    except PermissionError:
        print(Fore.RED + f"\n[ERROR] No se pudo guardar el archivo '{plantilla_path}'.")
        print(Fore.YELLOW + "El archivo está abierto en otro programa (como Excel) o está bloqueado. Ciérralo e intenta de nuevo.")
    except Exception as e:
        print(Fore.RED + f"Error al guardar el archivo: {e}")

    input(Fore.YELLOW + "\nPresione Enter para volver al menú...")

def descontarCanjes():
    bp()
    print(Fore.CYAN + Style.BRIGHT + "=== PROCESANDO PEDIDOS CANJES ===\n")
    
    folder = "Canjes"
    plantilla_path = "Plantilla.xlsx"
    
    if not os.path.exists(folder):
        print(Fore.RED + f"Error: La carpeta '{folder}' no existe.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return
        
    if not os.path.exists(plantilla_path):
        print(Fore.RED + f"Error: El archivo '{plantilla_path}' no existe.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    excel_files = [f for f in os.listdir(folder) if f.endswith('.xlsx') and not f.startswith('~$')]
    if not excel_files:
        print(Fore.YELLOW + f"No se encontraron archivos Excel en la carpeta '{folder}'.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    print(Fore.BLUE + f"Se encontraron {len(excel_files)} archivos Excel para procesar.")

    # 1. Cargar la plantilla actual para mapear los códigos válidos e índices de fila
    print(Fore.WHITE + f"Cargando {plantilla_path}...")
    try:
        df_plantilla = pd.read_excel(plantilla_path, engine='openpyxl')
    except Exception:
        try:
            df_plantilla = pd.read_excel(plantilla_path, engine='xlrd')
        except Exception as e:
            print(Fore.RED + f"Error al leer la plantilla: {e}")
            input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
            return

    # Crear mapa de códigos válidos de la plantilla
    # code_int -> row_index
    valid_codes = {}
    for idx, row in df_plantilla.iterrows():
        code_val = row['CODIGO']
        if pd.notna(code_val) and str(code_val).strip().upper() != 'TOTAL UNIDADES':
            try:
                code_int = int(float(str(code_val).strip()))
                valid_codes[code_int] = idx
            except (ValueError, TypeError):
                pass

    # 2. Procesar cada archivo de Canjes
    canjes_products = {}  # code_int -> accumulated_qty
    missing_products = []  # list of dicts: {'code': code, 'name': name, 'qty': qty, 'file': file_name}
    
    for file_name in excel_files:
        file_path = os.path.join(folder, file_name)
        print(Fore.WHITE + f"Procesando: {file_name}...")
        try:
            df_canjes = pd.read_excel(file_path, engine='openpyxl')
            
            # Recorrer filas desde el índice 4
            for idx, row in df_canjes.iterrows():
                # Fila index 3 es el encabezado 'PEDIDO', 'CODIGO'
                # Las filas de datos válidas están a partir de la fila index 4
                if idx <= 3:
                    continue
                
                # Column index 0: Pedido/Qty
                # Column index 1: Código
                # Column index 2: Producto/Name
                if len(row) < 3:
                    continue
                
                qty_val = row.iloc[0]
                code_val = row.iloc[1]
                name_val = row.iloc[2]
                
                if pd.isna(qty_val) or pd.isna(code_val):
                    continue
                
                # Extraer cantidad y código de forma segura
                try:
                    qty = int(float(str(qty_val).strip()))
                    code_int = int(float(str(code_val).strip()))
                except (ValueError, TypeError):
                    continue
                
                # Limpiar el nombre del producto
                name = str(name_val).strip() if pd.notna(name_val) else "Producto sin nombre"

                if code_int in valid_codes:
                    canjes_products[code_int] = canjes_products.get(code_int, 0) + qty
                else:
                    missing_products.append({
                        'code': code_int,
                        'name': name,
                        'qty': qty,
                        'file': file_name
                    })
        except Exception as e:
            print(Fore.RED + f"  Error al leer {file_name}: {e}")

    if not canjes_products and not missing_products:
        print(Fore.YELLOW + "No se extrajo ningún producto o cantidad de los archivos de canje.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # 3. Aplicar las sumas a la plantilla acumulando sobre el valor existente
    descuentos_actualizados = 0
    total_descuento_canjes = 0

    for code_int, qty in canjes_products.items():
        idx = valid_codes[code_int]
        current_descuento = df_plantilla.at[idx, 'Descuento']
        
        # Si es NaN o nulo, tratarlo como 0 para acumular correctamente
        base_descuento = float(current_descuento) if pd.notna(current_descuento) else 0.0
        df_plantilla.at[idx, 'Descuento'] = base_descuento + qty
        descuentos_actualizados += 1
        total_descuento_canjes += qty

    # 4. Recalcular el gran total de la columna Descuento
    # Sumar todo lo que sea numérico en la columna Descuento excepto la fila TOTAL UNIDADES
    suma_total_descuentos = 0.0
    for idx, row in df_plantilla.iterrows():
        code_val = row['CODIGO']
        if pd.notna(code_val) and str(code_val).strip().upper() == 'TOTAL UNIDADES':
            continue
        val = row['Descuento']
        if pd.notna(val):
            try:
                suma_total_descuentos += float(val)
            except (ValueError, TypeError):
                pass

    total_row_mask = df_plantilla['CODIGO'].astype(str).str.strip().str.upper() == 'TOTAL UNIDADES'
    if total_row_mask.any():
        df_plantilla.loc[total_row_mask, 'Descuento'] = suma_total_descuentos
        df_plantilla.loc[total_row_mask, 'PRODUCTOS'] = 'TOTAL PRODUCTOS'
        print(Fore.GREEN + f"Fila de 'TOTAL UNIDADES' actualizada con el gran total de {suma_total_descuentos} en Descuento.")
    else:
        print(Fore.YELLOW + "Advertencia: No se encontró la fila 'TOTAL UNIDADES' en la plantilla.")

    # 5. Guardar archivo como XLSX moderno
    print(Fore.WHITE + "Guardando cambios en Plantilla.xlsx...")
    try:
        df_plantilla.to_excel(plantilla_path, index=False, engine='openpyxl')
        print(Fore.GREEN + Style.BRIGHT + f"\n¡Éxito! Se procesaron {descuentos_actualizados} códigos de canjes coincidentes.")
        print(Fore.GREEN + f"Total de unidades agregadas por canjes: {total_descuento_canjes}")
        print(Fore.GREEN + f"Suma total actual en columna Descuento: {suma_total_descuentos}")
    except PermissionError:
        print(Fore.RED + f"\n[ERROR] No se pudo guardar el archivo '{plantilla_path}'.")
        print(Fore.YELLOW + "El archivo está abierto en otro programa (como Excel) o está bloqueado. Ciérralo e intenta de nuevo.")
    except Exception as e:
        print(Fore.RED + f"Error al guardar el archivo: {e}")

    # 6. Reporte de productos no descontados por ausencia de código en la plantilla
    if missing_products:
        print(Fore.RED + Style.BRIGHT + "\n" + "="*70)
        print(Fore.RED + Style.BRIGHT + "=== PRODUCTOS NO DESCONTADOS POR AUSENCIA DE CÓDIGO EN LA PLANTILLA ===")
        print(Fore.RED + Style.BRIGHT + "="*70)
        
        # Agrupar por archivo
        by_file = {}
        for item in missing_products:
            by_file.setdefault(item['file'], []).append(item)
            
        for file_name, items in by_file.items():
            print(Fore.YELLOW + f"\nArchivo: {file_name}")
            for item in items:
                print(Fore.WHITE + f"  Código: {item['code']:<6} | Cantidad: {item['qty']:<3} | Producto: {item['name']}")
        print(Fore.RED + Style.BRIGHT + "="*70)
    else:
        print(Fore.GREEN + "\nNo hubo productos no descontados. Todos los códigos existían en la plantilla.")

    input(Fore.YELLOW + "\nPresione Enter para volver al menú...")

def descontarMercado():
    bp()
    print("descontarMercado")
    time.sleep(2)
