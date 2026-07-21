import os
import tempfile
import unittest

import xlrd
import xlwt
from xlutils.copy import copy

from src.funciones import update_pedidos_descontados_section


class UpdatePedidosDescontadosSectionTests(unittest.TestCase):
    def test_writes_remitos_and_clients_in_template_section(self):
        temp_dir = tempfile.mkdtemp(prefix="template_test_", dir=".")
        try:
            xls_path = os.path.join(temp_dir, "plantilla_test.xls")
            wb = xlwt.Workbook()
            ws = wb.add_sheet("Hoja1")
            ws.write(0, 0, "Codigo")
            ws.write(353, 0, "PEDIDOS DESCONTADOS")
            ws.write(354, 0, "REMITO")
            ws.write(354, 2, "CLIENTE")
            wb.save(xls_path)

            rb = xlrd.open_workbook(xls_path, formatting_info=True)
            out_wb = copy(rb)
            out_sheet = out_wb.get_sheet(0)

            entries = [
                {"remito": "0001-00022639", "cliente": "NATURALDIST"},
                {"remito": "0001-00022640", "cliente": "ULTRA TECH"},
            ]

            update_pedidos_descontados_section(rb, out_wb, entries)
            out_wb.save(xls_path)

            updated_wb = xlrd.open_workbook(xls_path)
            updated_sheet = updated_wb.sheet_by_index(0)

            self.assertEqual(updated_sheet.cell_value(355, 0), "Remito: 0001-00022639")
            self.assertEqual(updated_sheet.cell_value(355, 2), "Cliente: NATURALDIST")
            self.assertEqual(updated_sheet.cell_value(356, 0), "Remito: 0001-00022640")
            self.assertEqual(updated_sheet.cell_value(356, 2), "Cliente: ULTRA TECH")
        finally:
            for root, dirs, files in os.walk(temp_dir, topdown=False):
                for name in files:
                    os.remove(os.path.join(root, name))
                for name in dirs:
                    os.rmdir(os.path.join(root, name))
            os.rmdir(temp_dir)

    def test_preserves_existing_entries_and_appends_new_ones(self):
        temp_dir = tempfile.mkdtemp(prefix="template_test_", dir=".")
        try:
            xls_path = os.path.join(temp_dir, "plantilla_test.xls")
            wb = xlwt.Workbook()
            ws = wb.add_sheet("Hoja1")
            ws.write(353, 0, "PEDIDOS DESCONTADOS")
            ws.write(354, 0, "REMITO")
            ws.write(354, 2, "CLIENTE")
            ws.write(355, 0, "0001-00022639")
            ws.write(355, 2, "NATURALDIST")
            wb.save(xls_path)

            rb = xlrd.open_workbook(xls_path, formatting_info=True)
            out_wb = copy(rb)

            entries = [{"remito": "0001-00022640", "cliente": "ULTRA TECH"}]
            update_pedidos_descontados_section(rb, out_wb, entries)
            out_wb.save(xls_path)

            updated_wb = xlrd.open_workbook(xls_path)
            updated_sheet = updated_wb.sheet_by_index(0)

            self.assertEqual(updated_sheet.cell_value(355, 0), "Remito: 0001-00022639")
            self.assertEqual(updated_sheet.cell_value(355, 2), "Cliente: NATURALDIST")
            self.assertEqual(updated_sheet.cell_value(356, 0), "Remito: 0001-00022640")
            self.assertEqual(updated_sheet.cell_value(356, 2), "Cliente: ULTRA TECH")
        finally:
            for root, dirs, files in os.walk(temp_dir, topdown=False):
                for name in files:
                    os.remove(os.path.join(root, name))
                for name in dirs:
                    os.rmdir(os.path.join(root, name))
            os.rmdir(temp_dir)


if __name__ == "__main__":
    unittest.main()
