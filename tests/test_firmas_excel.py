"""Firmas del Director Técnico y del Jefe de Laboratorio en los formatos de Excel: python -m unittest discover -s tests."""
from io import BytesIO
from pathlib import Path
import unittest
import zipfile

from openpyxl import load_workbook

from firmas_excel import agregar_firmas, NOMBRE_JEFE, _bytes_firma, FIRMA_DIRECTOR, FIRMA_JEFE

ROOT = Path(__file__).resolve().parents[1]
PLANTILLAS = ["GDA-FLC-014_humedad_natural.xlsx", "GDA-FLC-013_cbr.xlsx", "GDA-FLC-003_materia_organica.xlsx", "GDA-FLC-004_masa_unitaria.xlsx"]


def _firmas_en_dibujos(contenido):
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        return sum(z.read(n).count(b'name="Firma ') for n in z.namelist() if 'drawings/drawing' in n and n.endswith('.xml'))


@unittest.skipUnless(_bytes_firma(FIRMA_DIRECTOR) and _bytes_firma(FIRMA_JEFE), 'las firmas no están disponibles (están en los secretos, no en el repositorio)')
class FirmasExcelTest(unittest.TestCase):
    def test_pone_las_dos_firmas_y_el_nombre_del_jefe(self):
        for nombre in PLANTILLAS:
            with self.subTest(plantilla=nombre):
                original = (ROOT / 'templates' / nombre).read_bytes()
                con_firmas = agregar_firmas(original)
                self.assertGreaterEqual(_firmas_en_dibujos(con_firmas), 2)
                with zipfile.ZipFile(BytesIO(con_firmas)) as z:
                    self.assertIn('xl/media/gd_firma_director.png', z.namelist())
                    self.assertIn('xl/media/gd_firma_jefe.png', z.namelist())
                libro = load_workbook(BytesIO(con_firmas))
                textos = [c.value for ws in libro.worksheets for fila in ws.iter_rows() for c in fila if isinstance(c.value, str)]
                self.assertIn(NOMBRE_JEFE, textos)

    def test_es_idempotente(self):
        original = (ROOT / 'templates/GDA-FLC-014_humedad_natural.xlsx').read_bytes()
        una_vez = agregar_firmas(original)
        self.assertEqual(agregar_firmas(una_vez), una_vez)

    def test_no_toca_el_resto_del_formato(self):
        original = (ROOT / 'templates/GDA-FLC-014_humedad_natural.xlsx').read_bytes()
        con_firmas = agregar_firmas(original)
        with zipfile.ZipFile(BytesIO(original)) as a, zipfile.ZipFile(BytesIO(con_firmas)) as b:
            for nombre in a.namelist():
                if 'drawings/' in nombre or nombre.startswith('xl/worksheets/sheet') or nombre == '[Content_Types].xml':
                    continue
                self.assertEqual(a.read(nombre), b.read(nombre), nombre)


if __name__ == '__main__':
    unittest.main()
