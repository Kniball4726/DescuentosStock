import os
import shutil
import tempfile
import unittest
import datetime

from src.funciones import mover_archivos_a_descontados


class MoveFilesToDiscountedFolderTest(unittest.TestCase):
    def test_mover_archivos_a_descontados(self):
        temp_dir = tempfile.mkdtemp()
        try:
            origen = os.path.join(temp_dir, "origen")
            destino = os.path.join(temp_dir, "Descontados")
            os.makedirs(origen, exist_ok=True)
            os.makedirs(destino, exist_ok=True)

            archivo = os.path.join(origen, "remito.pdf")
            with open(archivo, "w", encoding="utf-8") as handle:
                handle.write("demo")

            moved_files = mover_archivos_a_descontados(origen, ["remito.pdf"], destino_folder=destino)

            self.assertEqual(len(moved_files), 1)
            self.assertFalse(os.path.exists(archivo))
            self.assertTrue(os.path.exists(os.path.join(destino, datetime.date.today().strftime("%Y-%m-%d"), "remito.pdf")))
        finally:
            shutil.rmtree(temp_dir)


if __name__ == "__main__":
    unittest.main()
