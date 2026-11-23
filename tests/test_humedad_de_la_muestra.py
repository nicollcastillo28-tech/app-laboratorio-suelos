"""La humedad natural de la muestra no se digita dos veces: Peso unitario Método B, Compresión inconfinada, Compresión en roca, Carga puntual,
Consolidación (humedad inicial) y Límite de contracción (natural) copian solas las masas del ensayo de Contenido de humedad de la muestra."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db  # noqa: E402
from test_guardado_ensayo import TABLAS, _sembrar, _sesion, ahora  # noqa: E402

HUMEDAD = {"hum_recipiente": "A-12", "hum_masa_recipiente": "18.45", "hum_masa_humedo_mas_recipiente": "117.65",
           "hum_seco_mas_recipiente": "98.45", "hum_seco_14h": "98.45", "hum_seco_15h": "98.45", "hum_seco_16h": "98.45"}


class HumedadDeLaMuestraTest(unittest.TestCase):
    def setUp(self):
        _sembrar()
        TABLAS["assays"][0].update(data=dict(HUMEDAD), status="finalizado")
        db._COLUMNAS_LIVIANAS = None
        db._CACHE_PESADO.clear()

    def _abrir(self, tipo, data=None):
        if not any(a["id"] == "a2" for a in TABLAS["assays"]):
            TABLAS["assays"].append({"id": "a2", "muestra_id": "m1", "tipo": tipo, "status": "sin-iniciar", "data": dict(data or {}),
                                     "observations": "", "laboratorist": "", "etapa_revision": None, "historial": [], "archived": False,
                                     "created_at": ahora(), "updated_at": ahora()})
        at = _sesion()
        at.session_state["selected_assay_id"] = "a2"
        at.session_state["selected_assay_type"] = tipo
        at.run()
        self.assertEqual(len(at.exception), 0, [e.value for e in at.exception])
        return at

    def _valores(self, at, claves):
        return [at.text_input(key=f"{k}_a2").value for k in claves]

    def test_peso_unitario_metodo_b_se_copia_y_se_actualiza(self):
        claves = ("mub_hum_recipiente", "mub_hum_humedo", "mub_hum_seco_19", "mub_hum_masa_rec")
        at = self._abrir("masa-unitaria", {"mu_metodo": "Método B"})
        self.assertEqual(self._valores(at, claves), ["A-12", "117.65", "98.45", "18.45"])
        self.assertTrue(any("Copiados del ensayo de Contenido de humedad" in c.value for c in at.caption))
        # Corregido a mano aquí: la corrección manda…
        at.text_input(key="mub_hum_masa_rec_a2").set_value("18.50").run()
        self.assertEqual(at.text_input(key="mub_hum_masa_rec_a2").value, "18.50")
        # …hasta que la humedad cambie allá.
        TABLAS["assays"][0]["data"]["hum_masa_recipiente"] = "18.40"
        TABLAS["assays"][0]["updated_at"] = ahora()
        self.assertEqual(self._valores(self._abrir("masa-unitaria"), claves)[3], "18.40")

    def test_compresion_carga_puntual_consolidacion_y_contraccion(self):
        casos = {
            "compresion-inconfinada": ("ci_hum_recipiente", "ci_hum_humedo", "ci_hum_seco_19", "ci_hum_masa_rec"),
            "compresion-roca": ("roca_hum_recipiente", "roca_hum_humedo", "roca_hum_seco_19", "roca_hum_masa_rec"),
            "carga-puntual": ("cp_hum_recipiente", "cp_hum_humedo", "cp_hum_seco_16", "cp_hum_masa_rec"),
            "consolidacion": ("cons_ini_recipiente", "cons_ini_humedo", "cons_ini_seco_19", "cons_ini_recipiente_masa"),
        }
        for tipo, claves in casos.items():
            with self.subTest(tipo=tipo):
                TABLAS["assays"][1:] = []
                at = self._abrir(tipo)
                self.assertEqual(self._valores(at, claves), ["A-12", "117.65", "98.45", "18.45"])
                self.assertTrue(any("se copian del ensayo de Contenido de humedad" in c.value for c in at.caption))
        TABLAS["assays"][1:] = []
        at = self._abrir("limite-contraccion")
        self.assertEqual(self._valores(at, ("lc_natural_hum_humedo", "lc_natural_hum_seco", "lc_natural_hum_recipiente")),
                         ["117.65", "98.45", "18.45"])
        self.assertEqual(at.text_input(key="lc_contraccion_hum_humedo_a2").value, "")   # la de contracción es propia del ensayo

    def test_la_humedad_despues_del_ensayo_no_se_copia(self):
        at = self._abrir("compresion-inconfinada", {"ci_hum_antes": "Después del ensayo"})
        self.assertEqual(at.text_input(key="ci_hum_humedo_a2").value, "")


if __name__ == "__main__":
    unittest.main()
