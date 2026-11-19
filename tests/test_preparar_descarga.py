"""Último paso de toda descarga (exportacion_formatos.preparar_descarga): sin protección, sin cadena de cálculo vieja, y el código de ensayo
(los dígitos del código del formato) en su celda con el «Informe ID No.» completo."""
import re
import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from exportacion_formatos import partes_hojas, preparar_descarga  # noqa: E402


def _hoja(contenido, nombre):
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        return z.read(partes_hojas(contenido)[nombre]).decode("utf-8"), z.namelist()


def _celda(xml, ref):
    m = re.search(r'<c r="' + ref + r'"[^>]*?(?:/>|>(.*?)</c>)', xml, re.S)
    return m.group(1) if m else None


class PrepararDescargaTest(unittest.TestCase):
    def test_terrones_codigo_034_en_k10_e_informe_id_completo(self):
        original = (ROOT / "templates" / "GDA-FLC-034_terrones_arcilla.xlsx").read_bytes()
        self.assertIn(b"<sheetProtection", zipfile.ZipFile(BytesIO(original)).read(partes_hojas(original)["FORMATO"]))
        xml, nombres = _hoja(preparar_descarga(original), "FORMATO")
        self.assertNotIn("<sheetProtection", xml)
        self.assertNotIn("xl/calcChain.xml", nombres)
        self.assertIn("034", _celda(xml, "K10"))
        self.assertIn('CONCATENATE(I9," ",J9,"-",K10,"-",D12,"-",F12)', _celda(xml, "H10"))

    def test_formato_sin_lugar_para_el_codigo_lo_lleva_en_la_formula(self):
        original = (ROOT / "templates" / "GDA-FLC-018_carga_puntual.xlsx").read_bytes()
        xml, _ = _hoja(preparar_descarga(original), "FORMATO")
        self.assertIn('&quot;018&quot;', _celda(xml, "L10").replace('"018"', '&quot;018&quot;'))

    def test_clasificacion_cada_informe_con_su_codigo(self):
        original = (ROOT / "templates" / "CLASIFICACION_DE_SUELOS.xlsm").read_bytes()
        xml, _ = _hoja(preparar_descarga(original), "FORMATO")
        self.assertIn(">001<", _celda(xml, "B50"))     # GDA-FLC-001
        self.assertIn(">001A<", _celda(xml, "AC10"))   # GDA-FLC-001A (límites)


if __name__ == "__main__":
    unittest.main()
