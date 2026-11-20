"""Datos mal digitados (validacion_datos): errores de digitación («11..1», «11-12»), rangos absurdos (115 golpes) y masas que no cuadran no
dejan enviar el ensayo a revisión. Las reglas se prueban solas y, de punta a punta, con la app real contra la base falsa de
test_guardado_ensayo."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db  # noqa: E402
import pruebas_datos  # noqa: E402
import validacion_datos  # noqa: E402
from test_guardado_ensayo import TABLAS, _sembrar, _sesion  # noqa: E402

TIPOS = ["alargamiento-aplanamiento", "caras-fracturadas", "carga-puntual", "cbr", "compresion-inconfinada", "compresion-roca", "consolidacion",
         "corte-directo", "granulometria", "gravedad-especifica", "humedad", "limite-contraccion", "limites", "masa-unitaria", "materia-organica",
         "pasa200", "proctor", "solidez-sulfatos", "terrones-arcilla"]


def _claves(tipo, data):
    return {x["clave"] for x in validacion_datos.valores_mal_digitados(tipo, data)}


class ReglasTest(unittest.TestCase):
    def test_los_datos_de_ejemplo_de_todos_los_ensayos_pasan(self):
        for tipo in TIPOS:
            for _ in range(3):   # los ejemplos tienen algo de azar
                with self.subTest(tipo=tipo):
                    self.assertEqual(validacion_datos.valores_mal_digitados(tipo, pruebas_datos.datos_ejemplo(tipo)), [])

    def test_errores_de_digitacion(self):
        lim = pruebas_datos.datos_ejemplo("limites")
        for valor in ("11..1", "11-12", "12 5", "1.234,5", "11,,1"):
            with self.subTest(valor=valor):
                self.assertEqual(_claves("limites", dict(lim, lim_ll_humedo_1=valor)), {"lim_ll_humedo_1"})
        # Lo que no es un número a propósito no se toca: recipientes con guion, fechas, temperaturas como texto.
        hum = dict(pruebas_datos.datos_ejemplo("humedad"), hum_recipiente="12-3", fecha_ensayo="2026-11-20", hum_temp_horno="110 ± 5 °C")
        self.assertEqual(_claves("humedad", hum), set())

    def test_golpes_del_limite_liquido(self):
        lim = pruebas_datos.datos_ejemplo("limites")
        for golpes in ("45", "115", "5", "25.5"):
            with self.subTest(golpes=golpes):
                self.assertEqual(_claves("limites", dict(lim, lim_ll_golpes_1=golpes)), {"lim_ll_golpes_1"})
        self.assertEqual(_claves("limites", dict(lim, lim_ll_golpes_1="35", lim_ll_golpes_2="25", lim_ll_golpes_3="15")), set())
        self.assertIn("lim_ll_golpes_1", _claves("limites", dict(lim, lim_metodo="Método B", lim_ll_golpes_1="35")))   # un punto: 20 a 30
        self.assertEqual(_claves("limites", dict(lim, lim_nl=True, lim_ll_golpes_1="115")), set())   # no líquido: no se revisa

    def test_rangos_de_otros_ensayos(self):
        cbr = pruebas_datos.datos_ejemplo("cbr")
        self.assertEqual(_claves("cbr", dict(cbr, cbrc_1_golpes="560", cbrc_2_capas="55")), {"cbrc_1_golpes", "cbrc_2_capas"})
        gs = dict(pruebas_datos.datos_ejemplo("gravedad-especifica"), gesp_a_temp="220", gesp_a_pct_retenido="150")
        self.assertEqual(_claves("gravedad-especifica", gs), {"gesp_a_temp", "gesp_a_pct_retenido"})

    def test_masas_que_no_cuadran(self):
        hum = pruebas_datos.datos_ejemplo("humedad")
        self.assertIn("hum_seco_mas_recipiente", _claves("humedad", dict(hum, hum_seco_mas_recipiente="999")))
        self.assertIn("hum_masa_recipiente", _claves("humedad", dict(hum, hum_masa_recipiente="-3")))
        self.assertIn("hum_masa_recipiente", _claves("humedad", dict(hum, hum_masa_recipiente="500")))
        cf = pruebas_datos.datos_ejemplo("caras-fracturadas")
        clave = next(k for k in cf if k.endswith("_una") and cf[k] not in (None, ""))
        total = clave.replace("_una", "_ret")
        self.assertIn(clave, _claves("caras-fracturadas", dict(cf, **{clave: str(float(cf[total]) + 10)})))


class EnvioBloqueadoTest(unittest.TestCase):
    def setUp(self):
        _sembrar()
        db._COLUMNAS_LIVIANAS = None
        db._CACHE_PESADO.clear()

    def test_no_se_envia_con_un_dato_mal_digitado_y_si_al_corregirlo(self):
        at = _sesion()
        at.run()
        for clave, valor in (("hum_recipiente", "A-12"), ("hum_masa_recipiente", "18.45"), ("hum_masa_humedo_mas_recipiente", "117.65"),
                             ("hum_seco_mas_recipiente", "98..45")):
            at.text_input(key=f"{clave}_a1").set_value(valor).run()
        self.assertTrue(any("mal digitados" in e.value and "98..45" in e.value for e in at.error))   # aviso en vivo, antes de enviar
        next(b for b in at.button if b.label == "Enviar a revisión").click().run()
        self.assertNotEqual(TABLAS["assays"][0]["status"], "finalizado")
        self.assertTrue(any("No se envió a revisión" in e.value for e in at.error))
        at.text_input(key="hum_seco_mas_recipiente_a1").set_value("98.45").run()
        self.assertFalse(any("mal digitados" in e.value for e in at.error))
        next(b for b in at.button if b.label == "Enviar a revisión").click().run()
        self.assertEqual(TABLAS["assays"][0]["status"], "finalizado")


if __name__ == "__main__":
    unittest.main()
