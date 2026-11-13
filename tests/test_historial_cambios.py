"""Historial de cambios de un ensayo (historial_cambios.py): reconstruir cada cambio desde las filas de assay_versions, agrupar por rato de
trabajo y no repetir las pesadas por hora que el formulario copia solas."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import historial_cambios as hc  # noqa: E402


class HistorialCambiosTest(unittest.TestCase):
    def test_filas_nuevas_y_viejas_se_reconstruyen_en_orden(self):
        actual = {"hum_recipiente": "A-3", "hum_masa_recipiente": "18.5", "_lastsync": "x", "foto": {"b64": "..."}}
        filas = [   # la más reciente primero
            {"saved_at": "2026-11-13T15:10:00+00:00", "autor": "Ana (Laboratorista)", "digitado_por": "Ana (Laboratorista)",
             "cambios": {"hum_masa_recipiente": {"a": "18.4", "d": "18.5"}}},
            {"saved_at": "2026-11-13T15:05:00+00:00", "autor": "Ana (Laboratorista)", "digitado_por": "Luis (Laboratorista)",
             "cambios": {"hum_recipiente": {"a": "A-1", "d": "A-3"}, "hum_masa_recipiente": {"a": None, "d": "18.4"}}},
            # fila vieja: la versión completa anterior (sin `cambios`)
            {"saved_at": "2026-11-12T10:00:00+00:00", "digitado_por": "Luis (Laboratorista)", "digitado_en": "2026-11-12T09:00:00+00:00",
             "data": {"hum_recipiente": "A-0"}},
        ]
        eventos = hc.reconstruir_eventos(filas, actual, "Ana (Laboratorista)", "2026-11-13T15:10:00+00:00")
        self.assertEqual([e["despues"].get("hum_recipiente") for e in eventos], ["A-0", "A-1", "A-3", "A-3"])
        self.assertTrue(eventos[0]["inicial"])                       # lo más viejo guardado ya traía datos
        self.assertEqual(eventos[1]["autor"], "Luis (Laboratorista)")   # A-0 -> A-1 lo hizo quien digitó el estado siguiente (fila vieja)
        self.assertNotIn("foto", eventos[-1]["despues"])
        self.assertNotIn("_lastsync", eventos[-1]["despues"])

    def test_pesadas_por_hora_copiadas_no_se_cuentan_aparte(self):
        antes = {}
        despues = {"hum_seco_mas_recipiente": "95.2", "hum_seco_14h": "95.2", "hum_seco_15h": "95.2", "hum_seco_16h": "95.2",
                   "lim_ll_seco_1": "20.1", "lim_ll_seco_14h_1": "20.1", "lim_ll_seco_15h_1": "19.9"}
        cambios = hc.cambios_de(antes, despues)
        self.assertEqual(cambios["hum_seco_mas_recipiente"][2], [14, 15, 16])
        self.assertNotIn("hum_seco_14h", cambios)
        self.assertEqual(cambios["lim_ll_seco_1"][2], [14])
        self.assertIn("lim_ll_seco_15h_1", cambios)                  # esta pesada se digitó distinta: sí es un cambio propio

    def test_agrupa_el_rato_de_trabajo_y_muestra_todos_los_valores(self):
        actual = {"lim_ll_golpes_1": "25"}
        filas = [
            {"saved_at": "2026-11-13T15:20:00+00:00", "autor": "Ana", "cambios": {"lim_ll_golpes_1": {"a": "24", "d": "25"}}},
            {"saved_at": "2026-11-13T15:10:00+00:00", "autor": "Ana", "cambios": {"lim_ll_golpes_1": {"a": "2", "d": "24"}}},
            {"saved_at": "2026-11-13T15:09:00+00:00", "autor": "Ana", "cambios": {"lim_ll_golpes_1": {"a": None, "d": "2"}}},
            {"saved_at": "2026-11-12T08:00:00+00:00", "autor": "Ana", "cambios": {"lim_metodo": {"a": None, "d": "A"}}},
        ]
        actual["lim_metodo"] = "A"
        bloques = hc.agrupar(hc.reconstruir_eventos(filas, actual, "Ana", None))
        self.assertEqual(len(bloques), 2)                            # otro día = otro rato de trabajo
        self.assertEqual(bloques[-1]["datos"]["lim_ll_golpes_1"]["valores"], [None, "2", "24", "25"])
        self.assertEqual(bloques[0]["estado"], {"lim_metodo": "A"})    # cómo quedó el ensayo al final de ese rato (para restaurar)

    def test_mismo_numero_escrito_distinto_no_es_cambio(self):
        self.assertEqual(hc.cambios_de({"x": "12.5"}, {"x": "12.50"}), {})
        self.assertEqual(hc.texto_valor([[0, 1.2], [60, 1.4]]), "2 lecturas")


if __name__ == "__main__":
    unittest.main()
