"""Las bitácoras de papel (GDA-FL-xxx) se llenan con los datos de ejemplo sin romper el archivo (no necesita LibreOffice)."""
import os
import sys
import zipfile
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bitacoras
import bitacoras_formatos as bf
import pruebas_datos

CTX = dict(fecha_ejecucion="21/10/2026", descripcion="Arcilla de baja plasticidad color café", prof_de="1.0", prof_a="1.5",
           localizacion="Yumbo", tipo_perf="Sondeo", perf_codigo="S-1", muestra="M-3", numero="088", anio="26", realizo="Laura Gómez")
CI = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300,
      330, 360, 390, 420, 450, 480, 510, 540, 650, 700, 750, 800, 900, 1000]


def _casos():
    d = pruebas_datos.datos_ejemplo
    return {
        "FL-004": lambda: bf.llenar_fl004(CTX, d("granulometria"), d("humedad"), d("masa-unitaria"), ["INV-214-13", "INV.E-213-13", "INV.E 123-13"]),
        "FL-005": lambda: bf.llenar_fl005(CTX, d("compresion-inconfinada"), CI),
        "FL-006": lambda: bf.llenar_fl006(CTX, d("corte-directo")),
        "FL-007": lambda: bf.llenar_fl007(CTX, d("compresion-roca")),
        "FL-008": lambda: bf.llenar_fl008(CTX, d("proctor"), ["10", "25", "56"]),
        "FL-017": lambda: bf.llenar_fl017(CTX, d("consolidacion")),
        "FL-019": lambda: bf.llenar_fl019(CTX, d("gravedad-especifica")),
        "FL-020": lambda: bf.llenar_fl020(CTX, d("carga-puntual")),
        "FL-022": lambda: bf.llenar_fl022(CTX, d("terrones-arcilla")),
        "FL-043": lambda: bf.llenar_fl043(CTX, d("solidez-sulfatos")),
        "FL-046": lambda: bf.llenar_fl046(CTX, d("alargamiento-aplanamiento")),
        "FLC-029": lambda: bf.llenar_flc029(CTX, {"mub_altura": "75.2", "mub_diametro": "38.1", "mub_masa": "181.5"}),
    }


def test_todas_las_bitacoras_se_llenan():
    for clave, armar in _casos().items():
        P, celdas, casillas = armar()
        assert celdas, clave
        xlsx = bitacoras.llenar(P, celdas, casillas)
        with zipfile.ZipFile(BytesIO(xlsx)) as z:
            assert z.testzip() is None, clave
            hoja = z.read(P.hoja_xml).decode("utf-8")
        assert "21/10/2026" in hoja, clave   # la fecha de ejecución llegó a la hoja


def test_marca_de_casillas():
    P, celdas, casillas = bf.llenar_fl005(CTX, dict(pruebas_datos.datos_ejemplo("compresion-inconfinada"), ci_condicion="Inalterada"), CI)
    assert casillas
    xlsx = bitacoras.llenar(P, celdas, casillas)
    with zipfile.ZipFile(BytesIO(xlsx)) as z:
        dibujo = z.read(P.dibujo_xml).decode("utf-8")
    assert dibujo.count("<a:t>X</a:t>") >= len(set(casillas))
