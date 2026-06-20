import os
import xlrd
from xlutils.copy import copy
import pandas as pd
import datetime

def update_missing_products_sheet(rb, wb, missing_products):
    sheet_name = 'No Encontrados'
    sheet_names = rb.sheet_names()
    
    if sheet_name in sheet_names:
        sheet_index = sheet_names.index(sheet_name)
        sheet = wb.get_sheet(sheet_index)
        r_sheet = rb.sheet_by_index(sheet_index)
        sheet.cell_overwrite_ok = True
        # Clear existing rows
        for r in range(r_sheet.nrows):
            for c in range(5):
                sheet.write(r, c, "")
    else:
        sheet = wb.add_sheet(sheet_name)
        sheet.cell_overwrite_ok = True

    # Headers
    headers = ['Código', 'Producto', 'Cantidad', 'Archivo Origen', 'Fecha Procesamiento']
    for col_idx, header in enumerate(headers):
        sheet.write(0, col_idx, header)

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if missing_products:
        for row_idx, item in enumerate(missing_products, start=1):
            sheet.write(row_idx, 0, item['code'])
            sheet.write(row_idx, 1, item['name'])
            sheet.write(row_idx, 2, item['qty'])
            sheet.write(row_idx, 3, item['file'])
            sheet.write(row_idx, 4, current_time)
    else:
        sheet.write(1, 0, "No se registraron productos sin código en la última ejecución.")

try:
    print("Testing sheet creation...")
    rb = xlrd.open_workbook('Plantilla.xlsx', formatting_info=True)
    wb = copy(rb)
    
    missing_test = [
        {'code': 2720, 'name': 'DUBAI protein bar x 12 unidades', 'qty': 2, 'file': 'test_file_1.xlsx'},
        {'code': 2552, 'name': 'Liquid Energy Gel', 'qty': 4, 'file': 'test_file_2.xlsx'}
    ]
    
    update_missing_products_sheet(rb, wb, missing_test)
    wb.save('Plantilla_sheet_test.xlsx')
    print("Saved successfully to Plantilla_sheet_test.xlsx")
    
    # Let's verify by loading it again
    rb2 = xlrd.open_workbook('Plantilla_sheet_test.xlsx')
    print("Sheets in new file:", rb2.sheet_names())
    sheet_missing = rb2.sheet_by_name('No Encontrados')
    print("Rows in missing sheet:", sheet_missing.nrows)
    for r in range(sheet_missing.nrows):
        print(f"Row {r}:", sheet_missing.row_values(r))
        
except Exception as e:
    import traceback
    traceback.print_exc()
