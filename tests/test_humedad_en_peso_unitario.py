"""Peso unitario Método B (GDA-FLC-030): las masas de humedad se copian solas del ensayo de Contenido de humedad de la muestra."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db  # noqa: E402
from test_guardado_ensayo import TABLAS, _sembrar, _sesion, ahora  # noqa: E402

HUMEDAD = {"hum_recipiente": "A-12", "hum_masa_recipiente": "18.45", "hum_masa_humedo_mas_recipiente": "117.65",
           "hum_seco_mas_recipiente": "98.45"}


class HumedadEnPesoUnitarioTest(unittest.TestCase):
    def setUp(self):
        _sembrar()
        TABLAS["assays"][0].update(data=dict(HUMEDAD), status="finalizado")
        TABLAS["assays"].append({"id": "a2", "muestra_id": "m1", "tipo": "masa-unitaria", "status": "sin-iniciar", "data": {"mu_metodo": "Método B"},
                                 "observations": "", "laboratorist": "", "etapa_revision": None, "historial": [], "archived": False,
                                 "created_at": ahora(), "updated_at": ahora()})
        db._COLUMNAS_LIVIANAS = None
        db._CACHE_PESADO.clear()

    def _abrir(self):
        at = _sesion()
        at.session_state["selected_assay_id"] = "a2"
        at.session_state["selected_assay_type"] = "masa-unitaria"
        at.run()
        self.assertEqual(len(at.exception), 0, [e.value for e in at.exception])
        return at

    def _valores(self, at):
        return [at.text_input(key=f"{k}_a2").value for k in ("mub_hum_recipiente", "mub_hum_humedo", "mub_hum_seco_19", "mub_hum_masa_rec")]

    def test_las_masas_de_la_humedad_se_copian_y_se_actualizan(self):
        at = self._abrir()
        self.assertEqual(self._valores(at), ["A-12", "117.65", "98.45", "18.45"])
        self.assertTrue(any("Copiados del ensayo de Contenido de humedad" in c.value for c in at.caption))
        # Corregido a mano aquí: la corrección manda…
        at.text_input(key="mub_hum_masa_rec_a2").set_value("18.50").run()
        self.assertEqual(at.text_input(key="mub_hum_masa_rec_a2").value, "18.50")
        # …hasta que la humedad cambie allá.
        TABLAS["assays"][0]["data"]["hum_masa_recipiente"] = "18.40"
        TABLAS["assays"][0]["updated_at"] = ahora()
        at2 = self._abrir()
        self.assertEqual(self._valores(at2)[3], "18.40")


if __name__ == "__main__":
    unittest.main()
