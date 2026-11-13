"""La descarga de Consolidación (GDA-FLC-009) se genera con los mismos imports que app.py, sin que la prueba le pase funciones: así un nombre
sin importar (pasó con partes_hojas: NameError al pulsar el botón) hace fallar la prueba."""
import ast
import math
import os
import re
import sys
import time
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import load_workbook  # noqa: E402

import pruebas_datos  # noqa: E402

GENERADORES = {"generar_excel_consolidacion"}


def _namespace():
    """Funciones y constantes de app.py (sin Streamlit) más sus imports de exportacion_formatos y presentacion_excel."""
    arbol = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    ns = {"math": math, "re": re, "zipfile": zipfile, "BytesIO": BytesIO, "os": os, "load_workbook": load_workbook}
    for n in arbol.body:
        if isinstance(n, ast.Import) and all(a.name in ("base64", "html", "json", "math", "contextlib", "os", "re", "time", "zipfile") for a in n.names):
            exec(compile(ast.Module(body=[n], type_ignores=[]), "app.py", "exec"), ns)
        elif isinstance(n, ast.ImportFrom) and n.module in ("exportacion_formatos", "presentacion_excel", "io"):
            exec(compile(ast.Module(body=[n], type_ignores=[]), "app.py", "exec"), ns)
        elif isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            n.decorator_list = []
            exec(compile(ast.Module(body=[n], type_ignores=[]), "app.py", "exec"), ns)
        elif isinstance(n, ast.Assign) and all(not isinstance(x, ast.Call) or (isinstance(x.func, ast.Attribute) and x.func.attr in ("compile", "join"))
                                               for x in ast.walk(n.value)):
            try:
                exec(compile(ast.Module(body=[n], type_ignores=[]), "app.py", "exec"), ns)
            except Exception:
                pass
    # Lo que depende de la base de datos o de la sesión se reemplaza por valores fijos.
    ns.update({
        "TEMPLATE_CONSOLIDACION": str(ROOT / "templates" / "GDA-FLC-009_consolidacion.xlsx"),
        "TEMPLATE_GESP_ARCILLA": str(ROOT / "templates" / "GDA-FLC-006_gravedad_arcilla.xlsx"),
        "_fecha_ddmmaaaa": lambda v: v, "_fecha_recepcion_perf": lambda *a: "12/10/2026", "fecha_ejecucion_informe": lambda *a: "13/10/2026",
        "fecha_emision_hoy": lambda: "16/10/2026", "get_perforacion": lambda *a: {"tipo": "Sondeo", "consecutivo": 1},
        "descripcion_para_informe": lambda m: "ARCILLA", "perf_para_encabezado": lambda pc, m, d: "S-1",
        "muestra_para_encabezado": lambda pc, m: "M-3", "TIPO_PERFORACION_EXCEL": {"Sondeo": "SONDEO"},
    })
    return ns


class DescargaConsolidacionTest(unittest.TestCase):
    def test_genera_el_excel_con_las_lecturas_y_t90(self):
        ns = _namespace()
        datos = pruebas_datos.datos_ejemplo("consolidacion")
        muestra = {"numero": "3", "profundidad_de": 1.0, "profundidad_hasta": 1.5}
        proyecto = {"numero": "088", "anio": "26", "nombre": "P", "cliente": "C", "localizacion": "Y"}
        inicio = time.time()
        contenido = ns["generar_excel_consolidacion"]("GDA-088-26", "S1", muestra, proyecto, datos)
        # Antes tardaba minutos (y el servidor cortaba la descarga): reescribía el XML entero por cada lectura de la máquina.
        self.assertLess(time.time() - inicio, 60)
        wb = load_workbook(BytesIO(contenido))
        self.assertEqual(wb["FORMATO"]["AB18"].value, 3)
        self.assertEqual(wb["DATOS MAQUINA "]["C5"].value, ns["_cons_deformacion_en"](datos["cons_maq_1"], wb["DATOS MAQUINA "]["B5"].value * 60))
        self.assertEqual(wb["0.5 kg"]["L29"].value, 0.5)
        self.assertEqual(wb["FORMATO"]["K30"].value, "=AB25")


if __name__ == "__main__":
    unittest.main()
