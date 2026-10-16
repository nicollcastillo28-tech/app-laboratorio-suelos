from pathlib import Path
from io import BytesIO
import unittest, zipfile
from exportacion_formatos import escribir_entradas,partes_hojas

ROOT=Path(__file__).resolve().parents[1]


class EstilosLotesTest(unittest.TestCase):
    def test_muestras_con_distintos_campos_comparten_los_mismos_estilos(self):
        original=(ROOT/'templates/CLASIFICACION_DE_SUELOS.xlsm').read_bytes()
        parte=partes_hojas(original)['FORMATO']
        a=escribir_entradas(original,parte,{'D6':'CLIENTE','D17':123.45,'S19':17})
        b=escribir_entradas(original,parte,{'D6':'CLIENTE','R41':'NO PLÁSTICO','R42':'N.P.'},celdas_formula_editables=('R41','R42'))
        with zipfile.ZipFile(BytesIO(a)) as za,zipfile.ZipFile(BytesIO(b)) as zb:
            self.assertEqual(za.read('xl/styles.xml'),zb.read('xl/styles.xml'))


if __name__=='__main__':unittest.main()
