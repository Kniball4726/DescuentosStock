import os
import time
import shutil
import pdfplumber
import pandas as pd
import xlrd
import datetime
from xlutils.copy import copy
from colorama import Fore, Style
from .helpers import borrarPantallas as bp


def normalize_header(value):
    """Normaliza un nombre de columna para compararlo de forma robusta."""
    return str(value).strip().lower().replace(" ", "")


def get_column_name(columns, candidates, default=None):
    """Devuelve el nombre real de una columna mediante comparación flexible."""
    normalized_candidates = {normalize_header(candidate): candidate for candidate in candidates}
    for col in columns:
        normalized_col = normalize_header(col)
        if normalized_col in normalized_candidates:
            return col
    return default


def mover_archivos_a_descontados(source_folder, file_names, destino_folder="Descontados"):
    """Mueve archivos procesados desde una carpeta fuente a una carpeta de salida."""
    os.makedirs(destino_folder, exist_ok=True)

    moved_files = []
    for file_name in file_names:
        source_path = os.path.join(source_folder, file_name)
        if not os.path.exists(source_path):
            continue

        destination_path = os.path.join(destino_folder, file_name)
        if os.path.exists(destination_path):
            base_name, extension = os.path.splitext(file_name)
            counter = 1
            while os.path.exists(destination_path):
                destination_path = os.path.join(destino_folder, f"{base_name}_{counter}{extension}")
                counter += 1

        shutil.move(source_path, destination_path)
        moved_files.append(file_name)

    return moved_files


def set_cell_value_preserve_format(out_sheet, row, col, value):
    """
    Escribe un valor en una celda de un Worksheet de xlwt (de xlutils.copy)
    preservando el índice de formato (xf_idx) original de la celda.
    """
    row_obj = out_sheet._Worksheet__rows.get(row)
    previous_cell = row_obj._Row__cells.get(col) if row_obj else None
    
    if value is None or pd.isna(value) or value == "":
        out_sheet.write(row, col, "")
    else:
        out_sheet.write(row, col, value)
        
    if previous_cell:
        new_row_obj = out_sheet._Worksheet__rows.get(row)
        new_cell = new_row_obj._Row__cells.get(col)
        if new_cell:
            new_cell.xf_idx = previous_cell.xf_idx

def update_missing_products_sheet(rb, wb, missing_products):
    """
    Agrega (acumula) entradas en la pestaña 'No Encontrados' del workbook.
    Los registros anteriores se conservan; los nuevos se agregan al final.
    Para limpiar la pestaña usar clear_missing_products_sheet().
    """
    sheet_name = 'No Encontrados'
    sheet_names = rb.sheet_names()

    if sheet_name in sheet_names:
        sheet_index = sheet_names.index(sheet_name)
        sheet = wb.get_sheet(sheet_index)
        r_sheet = rb.sheet_by_index(sheet_index)
        sheet.cell_overwrite_ok = True
        # Determinar la primera fila vacía para agregar al final
        next_row = r_sheet.nrows  # las filas existentes ya están copiadas por xlutils
    else:
        sheet = wb.add_sheet(sheet_name)
        sheet.cell_overwrite_ok = True
        # Escribir encabezados en la nueva pestaña
        headers = ['Código', 'Producto', 'Cantidad', 'Archivo Origen', 'Fecha Procesamiento']
        for col_idx, header in enumerate(headers):
            sheet.write(0, col_idx, header)
        next_row = 1

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if missing_products:
        for item in missing_products:
            sheet.write(next_row, 0, item['code'])
            sheet.write(next_row, 1, item['name'])
            sheet.write(next_row, 2, item['qty'])
            sheet.write(next_row, 3, item['file'])
            sheet.write(next_row, 4, current_time)
            next_row += 1
        print(Fore.GREEN + f"Pestaña 'No Encontrados' actualizada: se agregaron {len(missing_products)} productos faltantes.")
    else:
        print(Fore.GREEN + "No hubo productos faltantes en esta ejecución; pestaña 'No Encontrados' sin cambios.")


def clear_missing_products_sheet(rb, wb):
    """
    Limpia por completo la pestaña 'No Encontrados' dejando solo los encabezados.
    Se usa después de hacer el backup en guardarDescuentos.
    """
    sheet_name = 'No Encontrados'
    sheet_names = rb.sheet_names()

    if sheet_name in sheet_names:
        sheet_index = sheet_names.index(sheet_name)
        sheet = wb.get_sheet(sheet_index)
        r_sheet = rb.sheet_by_index(sheet_index)
        sheet.cell_overwrite_ok = True
        # Borrar todas las filas de datos (conservar fila 0 de encabezados)
        for r in range(1, r_sheet.nrows):
            for c in range(5):
                sheet.write(r, c, "")
        print(Fore.GREEN + "Pestaña 'No Encontrados' limpiada correctamente.")
    else:
        # Si no existe, crearla con encabezados para el próximo ciclo
        sheet = wb.add_sheet(sheet_name)
        sheet.cell_overwrite_ok = True
        headers = ['Código', 'Producto', 'Cantidad', 'Archivo Origen', 'Fecha Procesamiento']
        for col_idx, header in enumerate(headers):
            sheet.write(0, col_idx, header)
        print(Fore.GREEN + "Pestaña 'No Encontrados' creada lista para el próximo ciclo.")

def descontarMayoristas():
    """Descuenta los archivos cargados en la carpeta Mayoristas"""
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
    
    all_products = {} # code_int -> {'qty': qty, 'name': name, 'file': file_name}
    
    # 1. Procesar cada PDF
    for file_name in pdf_files:
        file_path = os.path.join(folder, file_name)
        print(Fore.WHITE + f"Procesando: {file_name}...")
        productos_en_archivo = 0
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
                            if tokens[0].isdigit() and tokens[-1].isdigit():
                                code = int(tokens[0])
                                qty = int(tokens[-1])
                                name = " ".join(tokens[1:-1])
                                if code in all_products:
                                    all_products[code]['qty'] += qty
                                else:
                                    all_products[code] = {'qty': qty, 'name': name, 'file': file_name}
                                productos_en_archivo += 1
        except Exception as e:
            print(Fore.RED + f"  Error al leer {file_name}: {e}")
            continue

        if productos_en_archivo == 0:
            print(Fore.YELLOW + f"  Archivo vacío o sin productos detectados, se omite: {file_name}")

    if not all_products:
        print(Fore.YELLOW + "No se extrajo ningún producto de los archivos PDF.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    print(Fore.GREEN + f"\nExtracción completa. Se encontraron {len(all_products)} productos distintos.")

    # 2. Cargar la plantilla con pandas para mapear índices
    print(Fore.WHITE + f"Cargando {plantilla_path}...")
    try:
        df_plantilla = pd.read_excel(plantilla_path, engine='xlrd')
    except Exception as e:
        print(Fore.RED + f"Error al leer la plantilla: {e}")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # Cargar workbook con xlrd (formato antiguo) para copiar estilos
    try:
        rb = xlrd.open_workbook(plantilla_path, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
    except Exception as e:
        print(Fore.RED + f"Error al cargar estilos de la plantilla: {e}")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # Detectar dinámicamente el índice de las columnas Descuento y Productos
    r_sheet = rb.sheet_by_index(0)
    headers = [str(val).strip() for val in r_sheet.row_values(0)]
    
    col_descuento = 4  # Valor por defecto
    descuento_col_name = get_column_name(headers, ['DESCUENTO', 'Descuento', 'descuento'], None)
    if descuento_col_name is not None:
        col_descuento = headers.index(descuento_col_name)
    else:
        print(Fore.YELLOW + "Advertencia: No se encontró la columna 'Descuento' por nombre. Usando columna index 4 por defecto.")
        
    col_productos = 2  # Valor por defecto
    productos_col_name = get_column_name(headers, ['PRODUCTOS', 'Productos', 'productos'], None)
    if productos_col_name is not None:
        col_productos = headers.index(productos_col_name)

    codigo_col_name = get_column_name(df_plantilla.columns.tolist(), ['CODIGO', 'Codigo', 'codigo'], None)
    if codigo_col_name is None:
        codigo_col_name = df_plantilla.columns[0]

    # Mapear códigos a índices de fila en pandas
    valid_codes = {}
    for idx, row in df_plantilla.iterrows():
        code_val = row[codigo_col_name]
        if pd.notna(code_val) and str(code_val).strip().upper() != 'TOTAL UNIDADES':
            try:
                code_int = int(float(str(code_val).strip()))
                valid_codes[code_int] = idx
            except (ValueError, TypeError):
                pass

    # Limpiar columna Descuento en la plantilla
    # Excel row index = pandas row index + 1
    for idx in range(len(df_plantilla)):
        set_cell_value_preserve_format(sheet, idx + 1, col_descuento, None)

    # 3. Emparejar y actualizar
    descuentos_actualizados = 0
    total_descuento_unidades = 0
    missing_products = []

    for code_int, info in all_products.items():
        qty = info['qty']
        name = info['name']
        file_origin = info['file']
        
        if code_int in valid_codes:
            idx = valid_codes[code_int]
            set_cell_value_preserve_format(sheet, idx + 1, col_descuento, qty)
            descuentos_actualizados += 1
            total_descuento_unidades += qty
        else:
            missing_products.append({
                'code': code_int,
                'name': name,
                'qty': qty,
                'file': file_origin
            })

    # 4. No actualizar la fila de total en la plantilla

    # Actualizar pestaña 'No Encontrados'
    update_missing_products_sheet(rb, wb, missing_products)

    # 5. Guardar archivo preservando el formato original
    print(Fore.WHITE + "Guardando cambios en Plantilla.xlsx...")
    try:
        wb.save(plantilla_path)
        print(Fore.GREEN + Style.BRIGHT + f"\n¡Éxito! Se actualizaron {descuentos_actualizados} productos en la columna 'Descuento'.")
        print(Fore.GREEN + f"Total de unidades descontadas: {total_descuento_unidades}")

        moved_files = mover_archivos_a_descontados(folder, pdf_files)
        if moved_files:
            print(Fore.GREEN + f"Archivos movidos a la carpeta 'Descontados': {', '.join(moved_files)}")
        else:
            print(Fore.YELLOW + "No se movieron archivos a la carpeta 'Descontados'.")
    except PermissionError:
        print(Fore.RED + f"\n[ERROR] No se pudo guardar el archivo '{plantilla_path}'.")
        print(Fore.YELLOW + "El archivo está abierto en otro programa (como Excel) o está bloqueado. Ciérralo e intenta de nuevo.")
    except Exception as e:
        print(Fore.RED + f"Error al guardar el archivo: {e}")

    # 6. Reporte por consola
    if missing_products:
        print(Fore.RED + Style.BRIGHT + "\n" + "="*70)
        print(Fore.RED + Style.BRIGHT + "=== PRODUCTOS NO DESCONTADOS POR AUSENCIA DE CÓDIGO EN LA PLANTILLA ===")
        print(Fore.RED + Style.BRIGHT + "="*70)
        
        by_file = {}
        for item in missing_products:
            by_file.setdefault(item['file'], []).append(item)
            
        for file_name, items in by_file.items():
            print(Fore.YELLOW + f"\nArchivo: {file_name}")
            for item in items:
                print(Fore.WHITE + f"  Código: {item['code']:<6} | Cantidad: {item['qty']:<3} | Producto: {item['name']}")
        print(Fore.RED + Style.BRIGHT + "="*70)

    input(Fore.YELLOW + "\nPresione Enter para volver al menú...")

def descontarCanjes():
    """Descuenta los archivos colocados en la carpeta Canjes que tienen formato .xls o xlsx"""
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
        df_plantilla = pd.read_excel(plantilla_path, engine='xlrd')
    except Exception as e:
        print(Fore.RED + f"Error al leer la plantilla: {e}")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    codigo_col_name = get_column_name(df_plantilla.columns.tolist(), ['CODIGO', 'Codigo', 'codigo'], None)
    if codigo_col_name is None:
        codigo_col_name = df_plantilla.columns[0]

    # Crear mapa de códigos válidos de la plantilla
    # code_int -> row_index
    valid_codes = {}
    for idx, row in df_plantilla.iterrows():
        code_val = row[codigo_col_name]
        if pd.notna(code_val) and str(code_val).strip().upper() != 'TOTAL UNIDADES':
            try:
                code_int = int(float(str(code_val).strip()))
                valid_codes[code_int] = idx
            except (ValueError, TypeError):
                pass

    # Cargar workbook con xlrd (formato antiguo) para copiar estilos
    try:
        rb = xlrd.open_workbook(plantilla_path, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
    except Exception as e:
        print(Fore.RED + f"Error al cargar estilos de la plantilla: {e}")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # Detectar dinámicamente el índice de las columnas Descuento y Productos
    r_sheet = rb.sheet_by_index(0)
    headers = [str(val).strip() for val in r_sheet.row_values(0)]
    
    col_descuento = 4  # Valor por defecto
    descuento_col_name = get_column_name(headers, ['DESCUENTO', 'Descuento', 'descuento'], None)
    if descuento_col_name is not None:
        col_descuento = headers.index(descuento_col_name)
    else:
        print(Fore.YELLOW + "Advertencia: No se encontró la columna 'Descuento' por nombre. Usando columna index 4 por defecto.")
        
    col_productos = 2  # Valor por defecto
    productos_col_name = get_column_name(headers, ['PRODUCTOS', 'Productos', 'productos'], None)
    if productos_col_name is not None:
        col_productos = headers.index(productos_col_name)

    descuento_val_col_name = get_column_name(df_plantilla.columns.tolist(), ['DESCUENTO', 'Descuento', 'descuento'], None)
    if descuento_val_col_name is None:
        descuento_val_col_name = df_plantilla.columns[3] if len(df_plantilla.columns) > 3 else df_plantilla.columns[0]

    # 2. Procesar cada archivo de Canjes
    canjes_products = {}  # code_int -> accumulated_qty
    missing_products = []  # list of dicts: {'code': code, 'name': name, 'qty': qty, 'file': file_name}
    
    for file_name in excel_files:
        file_path = os.path.join(folder, file_name)
        print(Fore.WHITE + f"Procesando: {file_name}...")
        try:
            df_canjes = pd.read_excel(file_path, engine='openpyxl')

            # Filtrar filas válidas (desde índice 4, con qty y code no nulos)
            filas_validas = [
                row for idx, row in df_canjes.iterrows()
                if idx > 3
                and len(row) >= 3
                and pd.notna(row.iloc[0])
                and pd.notna(row.iloc[1])
            ]

            if not filas_validas:
                print(Fore.YELLOW + f"  Archivo vacío o sin datos válidos, se omite: {file_name}")
                continue

            for row in filas_validas:
                qty_val = row.iloc[0]
                code_val = row.iloc[1]
                name_val = row.iloc[2]

                try:
                    qty = int(float(str(qty_val).strip()))
                    code_int = int(float(str(code_val).strip()))
                except (ValueError, TypeError):
                    continue

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
        current_descuento = df_plantilla.at[idx, descuento_val_col_name]
        
        # Si es NaN o nulo, tratarlo como 0 para acumular correctamente
        base_descuento = float(current_descuento) if pd.notna(current_descuento) else 0.0
        new_val = base_descuento + qty
        set_cell_value_preserve_format(sheet, idx + 1, col_descuento, new_val)
        df_plantilla.at[idx, descuento_val_col_name] = new_val
        descuentos_actualizados += 1
        total_descuento_canjes += qty

    # 4. Recalcular el gran total de la columna Descuento
    suma_total_descuentos = 0.0
    for idx, row in df_plantilla.iterrows():
        code_val = row[codigo_col_name]
        if pd.notna(code_val) and str(code_val).strip().upper() == 'TOTAL UNIDADES':
            continue
        val = row[descuento_val_col_name]
        if pd.notna(val):
            try:
                suma_total_descuentos += float(val)
            except (ValueError, TypeError):
                pass

    # 4. No actualizar la fila de total en la plantilla

    # Actualizar pestaña 'No Encontrados'
    update_missing_products_sheet(rb, wb, missing_products)

    # 5. Guardar archivo preservando el formato original
    print(Fore.WHITE + "Guardando cambios en Plantilla.xlsx...")
    try:
        wb.save(plantilla_path)
        print(Fore.GREEN + Style.BRIGHT + f"\n¡Éxito! Se procesaron {descuentos_actualizados} códigos de canjes coincidentes.")
        print(Fore.GREEN + f"Total de unidades agregadas por canjes: {total_descuento_canjes}")
        print(Fore.GREEN + f"Suma total actual en columna Descuento: {suma_total_descuentos}")

        moved_files = mover_archivos_a_descontados(folder, excel_files)
        if moved_files:
            print(Fore.GREEN + f"Archivos movidos a la carpeta 'Descontados': {', '.join(moved_files)}")
        else:
            print(Fore.YELLOW + "No se movieron archivos a la carpeta 'Descontados'.")
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

def guardarDescuentos():
    """Hace un respaldo del descuento en la carpeta Descuentos con la fecha actual y el formato .xlsx este suma los descuentos de mayoristas y canjes"""
    bp()
    print(Fore.CYAN + Style.BRIGHT + "=== GUARDAR DESCUENTOS ===\n")

    plantilla_path = "Plantilla.xlsx"
    descuentos_folder = "descuentos"

    # Verificar que Plantilla.xlsx existe
    if not os.path.exists(plantilla_path):
        print(Fore.RED + f"Error: El archivo '{plantilla_path}' no existe.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # Crear la carpeta 'descuentos' si no existe
    os.makedirs(descuentos_folder, exist_ok=True)

    # Definir nombre de la copia con la fecha actual
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
    destino_path = os.path.join(descuentos_folder, f"{fecha_actual}.xlsx")

    # Si ya existe una copia para hoy, agregar hora para no sobrescribir
    if os.path.exists(destino_path):
        fecha_actual_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        destino_path = os.path.join(descuentos_folder, f"{fecha_actual_hora}.xlsx")

    # 1. Copiar Plantilla.xlsx a descuentos/{fecha}.xlsx
    try:
        shutil.copy2(plantilla_path, destino_path)
        print(Fore.GREEN + f"Copia guardada exitosamente en: {destino_path}")
    except PermissionError:
        print(Fore.RED + f"\n[ERROR] No se pudo copiar el archivo '{plantilla_path}'.")
        print(Fore.YELLOW + "El archivo está abierto en Excel u otro programa. Ciérralo e intenta de nuevo.")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return
    except Exception as e:
        print(Fore.RED + f"Error al copiar el archivo: {e}")
        input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
        return

    # 2. Limpiar la columna 'Descuento' en la Plantilla original preservando el formato
    print(Fore.WHITE + "\nLimpiando columna 'Descuento' en Plantilla.xlsx...")
    try:
        rb = xlrd.open_workbook(plantilla_path, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        r_sheet = rb.sheet_by_index(0)

        # Detectar columna Descuento dinámicamente
        headers = [str(val).strip().upper() for val in r_sheet.row_values(0)]
        col_descuento = 4  # Valor por defecto
        if 'DESCUENTO' in headers:
            col_descuento = headers.index('DESCUENTO')
        else:
            print(Fore.YELLOW + "Advertencia: columna 'Descuento' no encontrada por nombre. Usando índice 4.")

        # Limpiar cada celda de datos de la columna Descuento (omitir fila 0 de encabezados)
        celdas_limpiadas = 0
        for row_idx in range(1, r_sheet.nrows):
            set_cell_value_preserve_format(sheet, row_idx, col_descuento, None)
            celdas_limpiadas += 1

        # Limpiar también la pestaña 'No Encontrados' para el nuevo ciclo
        print(Fore.WHITE + "Limpiando pestaña 'No Encontrados'...")
        clear_missing_products_sheet(rb, wb)

        wb.save(plantilla_path)
        print(Fore.GREEN + Style.BRIGHT + f"\n¡Éxito! Se limpiaron {celdas_limpiadas} filas de la columna 'Descuento' en Plantilla.xlsx.")
        print(Fore.GREEN + f"Backup guardado en: {destino_path}")

    except PermissionError:
        print(Fore.RED + f"\n[ERROR] No se pudo guardar los cambios en '{plantilla_path}'.")
        print(Fore.YELLOW + "El archivo está abierto en otro programa. Ciérralo e intenta de nuevo.")
        print(Fore.YELLOW + f"Nota: La copia de backup ya fue guardada en '{destino_path}'.")
    except Exception as e:
        print(Fore.RED + f"Error al limpiar la plantilla: {e}")
        print(Fore.YELLOW + f"Nota: La copia de backup ya fue guardada en '{destino_path}'.")

    input(Fore.YELLOW + "\nPresione Enter para volver al menú...")
