"""Casilla de perforación del encabezado de los Excel: el código de la perforación de la orden (no su número consecutivo ni otro nombre); el
código del apique solo cuando los apiques están agrupados en una perforación."""
import ast
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _funciones():
    ns = {"re": re}
    arbol = ast.parse(open(os.path.join(ROOT, "app.py"), encoding="utf-8").read())
    for n in arbol.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("_apique_y_muestra", "perf_para_encabezado", "muestra_para_encabezado"):
            exec(compile(ast.Module(body=[n], type_ignores=[]), "app.py", "exec"), ns)
    return ns


class PerforacionEncabezadoTest(unittest.TestCase):
    ns = _funciones()

    def _celdas(self, perf, numero):
        m = {"numero": numero}
        return self.ns["perf_para_encabezado"](perf, m, perf), self.ns["muestra_para_encabezado"](perf, m)

    def test_sondeos_y_perforaciones_llevan_su_codigo(self):
        for perf, numero in (("T2-SN-P-05", "3"), ("T2-SN-P-05", "P-05-1"), ("T2-SN-P-05", "SN-P-05-2"), ("S-1", "S-10-2"), ("S1", "1A"),
                             ("PZ-3", "2"), ("T2-SN-SE-05", "4"), ("APIQUE 1", "2")):
            with self.subTest(perf=perf, numero=numero):
                self.assertEqual(self._celdas(perf, numero)[0], perf)

    def test_apiques_agrupados_llevan_el_codigo_del_apique(self):
        self.assertEqual(self._celdas("AP-S", "AP-S-01-3"), ("AP-S-01", "3"))
        self.assertEqual(self._celdas("T2", "AP-S-01-3"), ("AP-S-01", "3"))      # perforación renombrada
        self.assertEqual(self._celdas("AP-S", "3"), ("AP-S", "3"))
        self.assertEqual(self._celdas("APQ", "APQ1-2"), ("APQ1", "2"))


if __name__ == "__main__":
    unittest.main()
