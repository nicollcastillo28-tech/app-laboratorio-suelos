"""Guardado de un ensayo, de punta a punta: corre la app REAL contra una base falsa en memoria (el db.py real con un cliente de Supabase falso)
y comprueba que lo digitado se guarda, sobrevive a las recargas de datos y a una sesión nueva, que el estado (rojo/naranja) sigue a lo digitado y
que el freno de seguridad impide que un fallo de la pantalla borre datos."""
import copy
import os
import sys
import time
import types
import unittest
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from streamlit.testing.v1 import AppTest

import db

RELOJ = [0]


def ahora():
    RELOJ[0] += 1
    return datetime(2026, 10, 23, 12, 0, 0, RELOJ[0], tzinfo=timezone.utc).isoformat()


TABLAS = {"projects": [], "perforaciones": [], "muestras": [], "assays": [], "notifications": [], "balance_checks": [], "profiles": []}
FALLAR_ESCRITURAS = [False]   # simula que la base no responde al guardar


class Q:
    def __init__(s, tabla):
        s.t, s.cols, s.f, s.orden, s.upd, s.ins = tabla, "*", [], None, None, None

    def select(s, cols="*"):
        s.cols = cols
        return s

    def eq(s, c, v):
        s.f.append((c, "eq", v))
        return s

    def neq(s, c, v):
        s.f.append((c, "neq", v))
        return s

    def in_(s, c, vs):
        s.f.append((c, "in", list(vs)))
        return s

    def order(s, c, desc=False):
        s.orden = (c, desc)
        return s

    def limit(s, n):
        return s

    def update(s, campos):
        s.upd = campos
        return s

    def insert(s, fila):
        s.ins = fila
        return s

    def execute(s):
        filas = TABLAS.setdefault(s.t, [])
        if s.ins is not None:
            nueva = dict(s.ins)
            nueva.setdefault("id", f"{s.t}-{len(filas) + 1}")
            nueva.setdefault("created_at", ahora())
            nueva.setdefault("updated_at", ahora())
            filas.append(nueva)
            return types.SimpleNamespace(data=[copy.deepcopy(nueva)])

        def cumple(r):
            for c, op, v in s.f:
                if op == "eq" and r.get(c) != v:
                    return False
                if op == "neq" and r.get(c) == v:
                    return False
                if op == "in" and r.get(c) not in v:
                    return False
            return True
        sel = [r for r in filas if cumple(r)]
        if s.upd is not None:
            if FALLAR_ESCRITURAS[0]:
                raise ConnectionError("sin conexión")
            for r in sel:
                r.update(copy.deepcopy(s.upd))
                r["updated_at"] = ahora()
            return types.SimpleNamespace(data=[copy.deepcopy(r) for r in sel])
        if s.orden:
            sel = sorted(sel, key=lambda r: r.get(s.orden[0]) or "", reverse=s.orden[1])
        cols = None if s.cols == "*" else [c.strip() for c in s.cols.split(",")]
        return types.SimpleNamespace(data=[{k: copy.deepcopy(v) for k, v in r.items() if cols is None or k in cols} for r in sel])


class Cliente:
    def table(self, nombre):
        return Q(nombre)

    def rpc(self, *a, **k):
        return Q("rpc")


db.get_client = lambda: Cliente()


def _sesion():
    at = AppTest.from_file(os.path.join(ROOT, "app.py"), default_timeout=120)
    at.secrets["SUPABASE_URL"] = "http://localhost:1"
    at.secrets["SUPABASE_ANON_KEY"] = "x"
    for k, v in dict(initialized=True, role="laboratorista", profile={"full_name": "Laura Gómez", "role": "laboratorista", "id": "u1"}, screen="assay-form",
                     projects=[], perforaciones={}, muestras={}, assays=[], notifications=[], balance_checks=[], nav_stack=[], bitacora_draft={},
                     draft_perforaciones=[], draft_muestras={}, selected_codigo="GDA-001-26", selected_perforacion="S1",
                     selected_muestra_id="GDA-001-26-S1-M1", selected_assay_id="a1", selected_assay_type="humedad", read_only_view=False).items():
        at.session_state[k] = v
    return at


def _sembrar():
    for t in TABLAS.values():
        t.clear()
    TABLAS["projects"].append({"id": "p1", "codigo_interno": "GDA-001-26", "numero": "001", "anio": "26", "nombre": "Proyecto de prueba", "archived": False,
                               "created_at": ahora(), "updated_at": ahora(), "localizacion": "Yumbo", "cliente": "Cliente", "norma": "INVIAS"})
    TABLAS["perforaciones"].append({"id": "pf1", "project_id": "p1", "tipo": "Sondeo", "consecutivo": 1, "codigo": "S1", "archived": False,
                                    "created_at": ahora(), "updated_at": ahora()})
    TABLAS["muestras"].append({"id": "m1", "perforacion_id": "pf1", "numero": "1", "id_unico": "GDA-001-26-S1-M1", "profundidad_de": 1.0,
                               "profundidad_hasta": 1.5, "tipo_muestra": "SPT", "ensayos": {"Humedad": True}, "observaciones": "", "archived": False,
                               "created_at": ahora(), "updated_at": ahora(), "desc_tipo_suelo": "ARCILLA", "desc_color": "MARRÓN",
                               "desc_consistencia": "FIRME", "descripcion_visual": None})
    TABLAS["assays"].append({"id": "a1", "muestra_id": "m1", "tipo": "humedad", "status": "sin-iniciar", "data": {}, "observations": "", "laboratorist": "",
                             "etapa_revision": None, "historial": [], "archived": False, "created_at": ahora(), "updated_at": ahora()})


def _valores_bd():
    a = TABLAS["assays"][0]
    return a["status"], {k: v for k, v in a["data"].items() if k.startswith("hum_") and v not in ("", True, [])}


MASA, REC, HUM = "hum_masa_recipiente_a1", "hum_recipiente_a1", "hum_masa_humedo_mas_recipiente_a1"
SECO = "hum_seco_mas_recipiente_a1"


class GuardadoEnsayoTest(unittest.TestCase):
    def setUp(self):
        _sembrar()
        db._COLUMNAS_LIVIANAS = None
        db._CACHE_PESADO.clear()

    def test_lo_digitado_se_guarda_y_sobrevive_a_recargas_y_a_una_sesion_nueva(self):
        at = _sesion()
        at.run()
        self.assertEqual(len(at.exception), 0)
        at.text_input(key=MASA).set_value("18.45").run()
        self.assertEqual(_valores_bd()[0], "en-proceso")          # un valor digitado: pasa a «en ejecución»
        at.session_state["_datos_ts"] = time.monotonic() - 1000   # vence el tiempo de los datos en memoria
        at.text_input(key=HUM).set_value("117.65").run()
        at.session_state["_datos_sucios"] = True
        at.text_input(key=REC).set_value("A-12").run()
        self.assertEqual([at.text_input(key=k).value for k in (MASA, HUM, REC)], ["18.45", "117.65", "A-12"])
        otra = _sesion()                                          # la pantalla se recarga: sesión nueva
        otra.run()
        self.assertEqual([otra.text_input(key=k).value for k in (MASA, HUM, REC)], ["18.45", "117.65", "A-12"])
        otra.text_input(key=MASA).set_value("19.00").run()
        self.assertEqual(_valores_bd()[1]["hum_masa_recipiente"], "19.00")
        self.assertEqual(_valores_bd()[1]["hum_masa_humedo_mas_recipiente"], "117.65")   # lo demás no se pierde

    def test_el_estado_depende_de_lo_digitado(self):
        at = _sesion()
        at.run()
        self.assertEqual(_valores_bd()[0], "sin-iniciar")          # abrir el ensayo no lo inicia
        at.text_input(key=REC).set_value("A-12").run()
        self.assertEqual(_valores_bd()[0], "en-proceso")           # basta el número de recipiente
        at.text_input(key=REC).set_value("").run()
        self.assertEqual(_valores_bd()[0], "sin-iniciar")          # todo borrado: vuelve a rojo

    def test_freno_si_muchos_datos_aparecen_en_blanco_de_golpe(self):
        at = _sesion()
        at.run()
        for k, v in ((REC, "A-12"), (MASA, "18.45"), (HUM, "117.65"), (SECO, "80.00")):
            at.text_input(key=k).set_value(v).run()
        antes = _valores_bd()
        for k in (REC, MASA, HUM, SECO, "hum_seco_14h_a1", "hum_seco_15h_a1", "hum_seco_16h_a1"):
            at.text_input(key=k).set_value("")
        at.run()
        self.assertEqual(_valores_bd(), antes)                     # la base queda intacta
        self.assertTrue(any("NO se guardó" in e.value for e in at.error))
        at.button(key="recuperar_a1").click().run()
        self.assertEqual(at.text_input(key=MASA).value, "18.45")   # lo guardado vuelve a la pantalla

    def test_borrar_de_a_un_dato_no_activa_el_freno(self):
        at = _sesion()
        at.run()
        for k, v in ((REC, "A-12"), (MASA, "18.45"), (HUM, "117.65")):
            at.text_input(key=k).set_value(v).run()
        at.text_input(key=MASA).set_value("").run()
        self.assertNotIn("hum_masa_recipiente", _valores_bd()[1])
        self.assertEqual(len(at.error), 0)

    def test_el_respaldo_permite_restaurar_la_version_anterior(self):
        at = _sesion()
        at.run()
        at.text_input(key=MASA).set_value("18.45").run()
        at.text_input(key=HUM).set_value("117.65").run()
        self.assertNotIn("_respaldo", TABLAS["assays"][0]["data"])      # agregar valores no genera respaldo
        at.text_input(key=MASA).set_value("20.00").run()                 # cambiar un valor ya guardado sí
        self.assertEqual(TABLAS["assays"][0]["data"]["_respaldo"]["hum_masa_recipiente"], "18.45")
        at.button(key="restaurar_a1").click().run()
        self.assertEqual(_valores_bd()[1]["hum_masa_recipiente"], "18.45")
        self.assertEqual(at.text_input(key=MASA).value, "18.45")

    def test_campos_viejos_tras_reconectar_no_pisan_lo_guardado_despues(self):
        """La tableta se durmió y al reconectar la pantalla trae los valores VIEJOS en los campos, pero otro dispositivo ya guardó datos más nuevos."""
        a = _sesion()
        a.run()
        a.text_input(key=MASA).set_value("18.45").run()
        a.text_input(key=HUM).set_value("117.65").run()
        TABLAS["assays"][0]["data"]["hum_masa_recipiente"] = "99.00"               # otro dispositivo corrigió
        TABLAS["assays"][0]["data"]["hum_masa_humedo_mas_recipiente"] = "200.00"
        TABLAS["assays"][0]["updated_at"] = ahora()
        b = _sesion()                                                              # sesión nueva: el navegador reenvía los campos viejos
        b.session_state[MASA] = "18.45"
        b.session_state[HUM] = "117.65"
        b.run()
        self.assertEqual(_valores_bd()[1]["hum_masa_recipiente"], "99.00")         # no se pisó
        self.assertEqual(_valores_bd()[1]["hum_masa_humedo_mas_recipiente"], "200.00")
        b.text_input(key=REC).set_value("A-12").run()                              # la persona digita OTRO campo
        self.assertEqual(_valores_bd()[1]["hum_recipiente"], "A-12")
        self.assertEqual(_valores_bd()[1]["hum_masa_recipiente"], "99.00")         # y lo viejo de los otros campos tampoco se pisa
        self.assertEqual(_valores_bd()[1]["hum_masa_humedo_mas_recipiente"], "200.00")

    def test_dos_sesiones_abiertas_no_se_borran_entre_si(self):
        """El mismo ensayo abierto en dos pantallas: lo que una guarda después de que la otra lo abrió NO se borra cuando la otra sigue digitando
        (el guardado se mezcla sobre lo que hay AHORA en la base, no sobre la copia que la pantalla trajo al abrir)."""
        a = _sesion()
        a.run()
        a.text_input(key=MASA).set_value("18.45").run()
        TABLAS["assays"][0]["data"]["hum_masa_humedo_mas_recipiente"] = "200.00"   # la otra pantalla guardó esto después de que esta abrió
        TABLAS["assays"][0]["data"]["hum_seco_mas_recipiente"] = "150.00"
        TABLAS["assays"][0]["updated_at"] = ahora()
        a.text_input(key=REC).set_value("A-12").run()                              # esta pantalla (con copia vieja) sigue digitando OTRO campo
        bd = _valores_bd()[1]
        self.assertEqual(bd["hum_recipiente"], "A-12")
        self.assertEqual(bd["hum_masa_recipiente"], "18.45")
        self.assertEqual(bd["hum_masa_humedo_mas_recipiente"], "200.00")           # lo que guardó la otra pantalla sigue ahí
        self.assertEqual(bd["hum_seco_mas_recipiente"], "150.00")

    def test_guardar_borrador_con_copia_vieja_no_borra_lo_de_la_otra_pantalla(self):
        a = _sesion()
        a.run()
        a.text_input(key=MASA).set_value("18.45").run()
        TABLAS["assays"][0]["data"]["hum_masa_humedo_mas_recipiente"] = "200.00"   # la otra pantalla guardó después
        TABLAS["assays"][0]["updated_at"] = ahora()
        [b for b in a.button if "Guardar borrador" in b.label][0].click().run()
        bd = _valores_bd()[1]
        self.assertEqual(bd["hum_masa_recipiente"], "18.45")
        self.assertEqual(bd["hum_masa_humedo_mas_recipiente"], "200.00")

    def test_quien_solo_mira_el_ensayo_no_queda_como_el_ultimo(self):
        a = _sesion()
        a.run()
        a.text_input(key=MASA).set_value("18.45").run()                             # Laura digita
        self.assertEqual(TABLAS["assays"][0]["laboratorist"], "Laura Gómez")
        self.assertIn("Laura Gómez", TABLAS["assays"][0].get("editado_por", ""))
        b = _sesion()
        b.session_state["profile"] = {"full_name": "Nicoll Castillo", "role": "laboratorista", "id": "u2"}
        b.run()                                                                     # Nicoll solo entra a mirar
        self.assertEqual(TABLAS["assays"][0]["laboratorist"], "Laura Gómez")
        b.text_input(key=REC).set_value("A-12").run()                               # y cuando digita algo, sí queda él
        self.assertEqual(TABLAS["assays"][0]["laboratorist"], "Nicoll Castillo")
        self.assertIn("Nicoll Castillo", TABLAS["assays"][0]["editado_por"])

    def test_si_la_base_falla_se_avisa_y_se_reintenta(self):
        at = _sesion()
        at.run()
        FALLAR_ESCRITURAS[0] = True
        try:
            at.text_input(key=MASA).set_value("18.45").run()
            self.assertTrue(any("NO SE GUARDÓ" in e.value for e in at.error))
            self.assertNotIn("hum_masa_recipiente", _valores_bd()[1])
        finally:
            FALLAR_ESCRITURAS[0] = False
        at.text_input(key=HUM).set_value("117.65").run()                           # vuelve la conexión: se guarda TODO lo pendiente
        self.assertEqual(_valores_bd()[1]["hum_masa_recipiente"], "18.45")
        self.assertEqual(_valores_bd()[1]["hum_masa_humedo_mas_recipiente"], "117.65")

    def test_al_entrar_se_parte_de_lo_guardado_en_la_base(self):
        at = _sesion()
        at.run()
        at.text_input(key=MASA).set_value("18.45").run()
        TABLAS["assays"][0]["data"]["hum_masa_humedo_mas_recipiente"] = "200.00"   # otro dispositivo guardó después
        TABLAS["assays"][0]["updated_at"] = ahora()
        at.session_state["screen"] = "home"
        at.run()
        at.session_state["screen"] = "assay-form"
        at.session_state["_nav_token"] = 99
        at.run()
        self.assertEqual(at.text_input(key=HUM).value, "200.00")


if __name__ == "__main__":
    unittest.main()
