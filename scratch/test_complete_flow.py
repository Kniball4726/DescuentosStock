import os
import xlrd
from xlutils.copy import copy
import pandas as pd

def set_cell_value_preserve_format(out_sheet, row, col, value):
    """Mantiene la configuración original de las filas y columnas pre establecidas"""
    row_obj = out_sheet._Worksheet__rows.get(row)
    previous_cell = row_obj._Row__cells.get(col) if row_obj else None
    
    if value is None or pd.isna(value):
        out_sheet.write(row, col, "")
    else:
        out_sheet.write(row, col, value)
        
    if previous_cell:
        new_row_obj = out_sheet._Worksheet__rows.get(row)
        new_cell = new_row_obj._Row__cells.get(col)
        if new_cell:
            new_cell.xf_idx = previous_cell.xf_idx

try:
    print("Reading template...")
    # Read with pandas first to locate indexes easily
    df_plantilla = pd.read_excel('Plantilla.xlsx', engine='xlrd')
    
    # Load with xlrd for xlutils
    rb = xlrd.open_workbook('Plantilla.xlsx', formatting_info=True)
    wb = copy(rb)
    sheet = wb.get_sheet(0)
    
    # Let's map valid codes
    valid_codes = {}
    for idx, row in df_plantilla.iterrows():
        code_val = row['CODIGO']
        if pd.notna(code_val) and str(code_val).strip().upper() != 'TOTAL UNIDADES':
            try:
                code_int = int(float(str(code_val).strip()))
                valid_codes[code_int] = idx
            except (ValueError, TypeError):
                pass
                
    print(f"Mapped {len(valid_codes)} valid codes.")
    
    # Let's test clearing and setting values
    # For descontarMayoristas: clear Descuento column
    for idx in range(len(df_plantilla)):
        # Excel row is idx + 1
        set_cell_value_preserve_format(sheet, idx + 1, 4, "")
        
    # Write some dummy quantities
    dummy_updates = {330: 10, 366: 15, 2156: 8}
    total_qty = 0
    for code, qty in dummy_updates.items():
        if code in valid_codes:
            idx = valid_codes[code]
            set_cell_value_preserve_format(sheet, idx + 1, 4, qty)
            total_qty += qty
            
    # Write total row
    # TOTAL UNIDADES is at df_plantilla last row
    total_row_mask = df_plantilla['CODIGO'].astype(str).str.strip().str.upper() == 'TOTAL UNIDADES'
    if total_row_mask.any():
        total_idx = df_plantilla[total_row_mask].index[0]
        set_cell_value_preserve_format(sheet, total_idx + 1, 4, total_qty)
        set_cell_value_preserve_format(sheet, total_idx + 1, 2, "TOTAL PRODUCTOS")
        print(f"Updated total row at Excel row {total_idx + 2} with qty {total_qty}.")
        
    wb.save('Plantilla_flow_test.xlsx')
    print("Success saving flow test!")
    
except Exception as e:
    import traceback
    traceback.print_exc()
