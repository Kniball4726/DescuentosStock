import xlrd
import xlwt
from xlutils.copy import copy

def set_cell_value_preserve_format(out_sheet, row, col, value):
    # 1. Get the previous cell to extract its style index
    row_obj = out_sheet._Worksheet__rows.get(row)
    previous_cell = row_obj._Row__cells.get(col) if row_obj else None
    
    # 2. Write the new value
    out_sheet.write(row, col, value)
    
    # 3. Re-apply the original style index
    if previous_cell:
        new_row_obj = out_sheet._Worksheet__rows.get(row)
        new_cell = new_row_obj._Row__cells.get(col)
        if new_cell:
            new_cell.xf_idx = previous_cell.xf_idx

try:
    print("Opening workbook...")
    rb = xlrd.open_workbook('Plantilla.xlsx', formatting_info=True)
    wb = copy(rb)
    sheet = wb.get_sheet(0)
    
    # Let's test writing to row 246 (TOTAL UNIDADES), column 4 (Descuento)
    set_cell_value_preserve_format(sheet, 246, 4, 1234)
    # Write to row 246 (TOTAL UNIDADES), column 2 (PRODUCTOS)
    set_cell_value_preserve_format(sheet, 246, 2, "TOTAL PRODUCTOS HACK")
    
    wb.save('Plantilla_test.xlsx')
    print("Success! Saved to Plantilla_test.xlsx")
except Exception as e:
    print(f"Error occurred: {e}")
