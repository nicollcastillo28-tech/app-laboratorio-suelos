"""Descarga de Granulometría (GDA-FLC-001): los tamices que no retuvieron nada (sin dato) salen en 0, aunque su casilla nunca se haya guardado;
con solo Pasa 200 (sin tamices) no se inventa una curva."""
import os
import sys
import unittest
from io import BytesIO

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_consolidacion_descarga import ROOT, _namespace  # noqa: E402


class CerosGranulometriaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _namespace()
        cls.ns["TEMPLATE_GRANULOMETRIA"] = str(ROOT / "templates" / "CLASIFICACION_DE_SUELOS.xlsm")
        cls.muestra = {"numero": "3", "profundidad_de": 1.0, "profundidad_hasta": 1.5, "id_unico": "GDA-088-26-S1-M3"}
        cls.proyecto = {"numero": "088", "anio": "26", "nombre": "PROYECTO", "cliente": "CLIENTE", "localizacion": "YUMBO"}

    def _tamices(self, gran_data):
        b = self.ns["_generar_excel_clasificacion"]("GDA-088-26", "S1", self.muestra, self.proyecto, gran_data=gran_data, lim_data=None)
        ws = load_workbook(BytesIO(b), keep_vba=True)["FORMATO"]
        return {k: ws[celda].value for k, _l, _a, celda in self.ns["SIEVES"]}

    def test_tamices_sin_dato_salen_en_cero(self):
        # Solo se guardaron los tamices digitados (las casillas vacías nunca llegaron a la base).
        tamices = self._tamices({"masa_inicial_seca": "500", "s_38": "12.5", "s_4": "20", "s_200": "80.3"})
        self.assertEqual(tamices.pop("s_38"), 12.5)
        self.assertEqual(tamices.pop("s_4"), 20)
        self.assertEqual(tamices.pop("s_200"), 80.3)
        self.assertEqual(set(tamices.values()), {0})
        # Casillas guardadas vacías: también 0.
        self.assertEqual(set(self._tamices({"masa_inicial_seca": "500", "s_3": "", "s_10": None, "s_200": "5"}).values()) - {5}, {0})

    def test_solo_pasa_200_no_inventa_tamices(self):
        tamices = self._tamices({"masa_inicial_seca": "500", "p200_masa_recipiente_antes": "20"})
        self.assertTrue(all(v in (None, "") for v in tamices.values()), tamices)


if __name__ == "__main__":
    unittest.main()
