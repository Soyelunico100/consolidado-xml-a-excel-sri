import importlib.util
import shutil
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def load_module_from(path):
    spec = importlib.util.spec_from_file_location("consolidado_test", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class RutasRelativasTest(unittest.TestCase):
    def test_main_usa_rutas_relativas_y_genera_excel(self):
        with tempfile.TemporaryDirectory() as tmp:
            temp_root = Path(tmp)
            shutil.copy2(ROOT / "consolidado_xml_a_excel.py", temp_root / "consolidado_xml_a_excel.py")

            entrada = temp_root / "XML"
            entrada.mkdir()
            shutil.copy2(ROOT / "ejemplos" / "factura_ficticia.xml", entrada / "factura_ficticia.xml")

            module = load_module_from(temp_root / "consolidado_xml_a_excel.py")
            module.REQUIERE_ACTIVACION = False
            module.GENERAR_XML_ATS = False
            module.MOVER_XML_PROCESADOS = False
            module.main()

            self.assertEqual(module.XML_FOLDER, temp_root / "XML")
            self.assertEqual(module.SALIDA_EXCEL_DIR, temp_root)
            self.assertTrue((temp_root / "XML" / "factura_ficticia.xml").exists())
            self.assertGreater(len(list(temp_root.glob("*.xlsx"))), 0)


if __name__ == "__main__":
    unittest.main()
