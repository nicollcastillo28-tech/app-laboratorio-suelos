"""Genera etiquetas_campos.json: el nombre legible de cada dato de los ensayos, tal como se ve en el formulario.

Dibuja cada formulario de ensayo (Streamlit AppTest, con una base de datos simulada en memoria y los datos de ejemplo de
pruebas_datos) y, para cada casilla, toma su etiqueta y el título de la tarjeta donde está:
    lim_lp_recipiente_1 -> "Límite Plástico › Recipiente No. 1"
El historial de cambios de los ensayos usa estos nombres. Volver a correrlo cuando se agreguen o renombren casillas:

    python tools/generar_etiquetas_campos.py
"""
import copy
import json
import os
import re
import sys
import types
from datetime import datetime, timedelta, timezone

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import db  # noqa: E402
import pruebas_datos  # noqa: E402
from streamlit.testing.v1 import AppTest  # noqa: E402

TABLAS = {}
_reloj = [0]


def _ahora():
    _reloj[0] += 1
    return (datetime(2026, 10, 1, 12, tzinfo=timezone.utc) + timedelta(seconds=_reloj[0])).isoformat()


class _Consulta:
    """Lo mínimo de la API de supabase-py que usa la app, sobre listas en memoria."""

    def __init__(self, tabla):
        self.t, self.filtros, self.orden, self.cambios, self.nuevas, self.limite = tabla, [], None, None, None, None

    def select(self, *_a, **_k): return self
    def eq(self, c, v): self.filtros.append((c, "eq", v)); return self
    def neq(self, c, v): self.filtros.append((c, "neq", v)); return self
    def in_(self, c, vs): self.filtros.append((c, "in", list(vs))); return self
    def is_(self, c, _v): self.filtros.append((c, "eq", None)); return self
    def order(self, c, desc=False): self.orden = (c, desc); return self
    def limit(self, n): self.limite = n; return self
    def update(self, campos): self.cambios = campos; return self
    def insert(self, filas): self.nuevas = filas; return self
    def upsert(self, filas, **_k): self.nuevas = filas; return self
    def delete(self): self.cambios = "__borrar__"; return self

    def execute(self):
        filas = TABLAS.setdefault(self.t, [])
        if self.nuevas is not None:
            lista = self.nuevas if isinstance(self.nuevas, list) else [self.nuevas]
            for f in lista:
                f = dict(f)
                f.setdefault("id", f"{self.t}-{len(filas) + 1}")
                f.setdefault("created_at", _ahora())
                f.setdefault("updated_at", _ahora())
                filas.append(f)
            return types.SimpleNamespace(data=[copy.deepcopy(f) for f in filas[-len(lista):]])
        sel = [r for r in filas if all((op == "eq" and r.get(c) == v) or (op == "neq" and r.get(c) != v) or (op == "in" and r.get(c) in v)
                                       for c, op, v in self.filtros)]
        if self.cambios == "__borrar__":
            for r in sel:
                filas.remove(r)
            return types.SimpleNamespace(data=[])
        if self.cambios is not None:
            for r in sel:
                r.update(copy.deepcopy(self.cambios))
            return types.SimpleNamespace(data=[copy.deepcopy(r) for r in sel])
        if self.orden:
            sel = sorted(sel, key=lambda r: r.get(self.orden[0]) or "", reverse=self.orden[1])
        return types.SimpleNamespace(data=[copy.deepcopy(r) for r in sel[:self.limite or None]])


class _Cliente:
    def table(self, nombre): return _Consulta(nombre)
    def rpc(self, *_a, **_k): return _Consulta("rpc")


db.get_client = lambda: _Cliente()

# (tipo de ensayo, nombre en la bitácora, variantes de datos para que aparezcan todas las casillas)
ENSAYOS = [
    ("humedad", "Humedad", [{}]),
    ("granulometria", "Granulometría", [{}]),
    ("pasa200", "Pasa 200", [{}]),
    ("limites", "Límites de Atterberg", [{}]),
    ("masa-unitaria", "Peso unitario", [{}]),
    ("cbr", "CBR", [{}]),
    ("corte-directo", "Corte CD", [{}]),
    ("gravedad-especifica", "Gravedad específica", [{"gesp_sel": "Ambos"}, {"gesp_sel": "Arcillas y limos (INV E-128)"}]),
    ("proctor", "Proctor", [{}]),
    ("materia-organica", "Materia orgánica", [{}]),
    ("limite-contraccion", "Límite de contracción", [{}]),
    ("consolidacion", "Consolidación", [{}]),
    ("compresion-inconfinada", "Compresión inconfinada", [{}]),
    ("compresion-roca", "Compresión en roca", [{}]),
    ("carga-puntual", "Carga puntual", [{}]),
    ("solidez-sulfatos", "Solidez en sulfatos", [{}]),
    ("terrones-arcilla", "Terrones de arcilla", [{}]),
    ("alargamiento-aplanamiento", "Alargamiento y aplanamiento", [{}]),
]
TIPOS_CASILLA = {"TextInput", "NumberInput", "Radio", "Selectbox", "Multiselect", "Checkbox", "Toggle", "TextArea", "DateInput", "ButtonGroup",
                 "SelectSlider"}


def _titulo_tarjeta(texto_html):
    texto = re.sub(r"<[^>]+>", " ", texto_html)
    texto = re.sub(r"\s+", " ", texto).strip()
    texto = re.sub(r"^[a-z_]+ ", "", texto)   # el nombre del ícono de Material que va antes del título
    texto = re.sub(r"\s*\((?:INV|I\.N\.V|ASTM)[^)]*\)", "", texto, flags=re.I)
    return texto.strip()


def _etiqueta_limpia(etiqueta):
    etiqueta = re.sub(r"\s+", " ", etiqueta or "").strip()
    # Pesadas del secado: «(14 hrs)», «(16 horas)» -> «, pesada de las 14 h»
    etiqueta = re.sub(r"\s*\((\d{1,2})\s*(?:hrs?|horas?)\)", r", pesada de las \1 h", etiqueta)
    return etiqueta


# Casillas con una etiqueta interna corta al final («2 1/2" a 1 1/2" ini»): qué es cada una según el ensayo (prefijo de la clave).
_FINALES = {
    "sulf_": {"ini": "masa inicial seca (g)", "fin": "masa final seca (g)"},
    "ter_": {"ini": "masa inicial lavada seca (g)", "fin": "masa final lavada seca (g)"},
    "aa_": {"ini": "masa de la fracción (g)", "par": "masa de las partículas (g)"},
}


def _afinar(dato, etiqueta):
    for prefijo, finales in _FINALES.items():
        if dato.startswith(prefijo):
            m = re.match(r"^(.*?)\s+(ini|fin|par)$", etiqueta)
            if m and m.group(2) in finales:
                return f"{m.group(1)}, {finales[m.group(2)]}"
    etiqueta = re.sub(r"\s+—\s+ini$", ", inicial", etiqueta)
    etiqueta = re.sub(r"\s+—\s+fin$", ", final", etiqueta)
    return etiqueta


def _recorrer(nodo, tarjeta, salida, assay_id):
    hijos = getattr(nodo, "children", None)
    hijos = list(hijos.values()) if isinstance(hijos, dict) else (hijos or [])
    for hijo in hijos:
        tipo = type(hijo).__name__
        if tipo == "Markdown" and "font-size:19px" in (hijo.value or ""):
            tarjeta = _titulo_tarjeta(hijo.value)
            continue
        clave = getattr(hijo, "key", None)
        if tipo in TIPOS_CASILLA and clave and clave.endswith(f"_{assay_id}"):
            dato = clave[: -len(assay_id) - 1]
            etiqueta = _afinar(dato, _etiqueta_limpia(getattr(hijo, "label", "")))
            if etiqueta and dato not in salida:
                salida[dato] = f"{tarjeta} › {etiqueta}" if tarjeta and tarjeta.lower() not in etiqueta.lower() else etiqueta
        # Los bloques (contenedores, columnas, desplegables) se recorren por dentro; un título de tarjeta que aparezca dentro de un
        # contenedor solo vale para lo que está dentro de él.
        _recorrer(hijo, tarjeta, salida, assay_id)
    return tarjeta


def etiquetas_de(tipo, nombre, extra):
    TABLAS.clear()
    TABLAS.update({k: [] for k in ("projects", "perforaciones", "muestras", "assays", "notifications", "balance_checks", "profiles",
                                   "assay_versions", "balanzas")})
    TABLAS["projects"].append({"id": "p1", "codigo_interno": "GDA-001-26", "numero": "001", "anio": "26", "nombre": "Proyecto",
                               "archived": False, "created_at": _ahora(), "updated_at": _ahora(), "localizacion": "Yumbo",
                               "cliente": "Cliente", "norma": "INVIAS"})
    TABLAS["perforaciones"].append({"id": "pf1", "project_id": "p1", "tipo": "Sondeo", "consecutivo": 1, "codigo": "S-1", "archived": False,
                                    "created_at": _ahora(), "updated_at": _ahora(), "fecha_ingreso": "2026-10-01"})
    TABLAS["muestras"].append({"id": "m1", "perforacion_id": "pf1", "numero": "1", "id_unico": "GDA-001-26-S-1-M1", "profundidad_de": 1,
                               "profundidad_hasta": 1.5, "tipo_muestra": "SPT", "ensayos": {nombre: True}, "observaciones": "",
                               "desc_tipo_suelo": "ARCILLA", "desc_color": "GRIS", "archived": False, "created_at": _ahora(),
                               "updated_at": _ahora()})
    TABLAS["assays"].append({"id": "a1", "muestra_id": "m1", "tipo": tipo, "status": "en-proceso",
                             "data": {**pruebas_datos.datos_ejemplo(tipo), **extra}, "observations": "", "laboratorist": "Ana",
                             "etapa_revision": None, "historial": [], "archived": False, "created_at": _ahora(), "updated_at": _ahora()})
    at = AppTest.from_file(os.path.join(RAIZ, "app.py"), default_timeout=300)
    at.secrets["SUPABASE_URL"] = "http://localhost:1"
    at.secrets["SUPABASE_ANON_KEY"] = "x"
    estado = dict(initialized=True, role="laboratorista", profile={"full_name": "Ana", "role": "laboratorista", "id": "u1"},
                  screen="assay-form", projects=[], perforaciones={}, muestras={}, assays=[], notifications=[], balance_checks=[],
                  nav_stack=[], bitacora_draft={}, draft_perforaciones=[], draft_muestras={}, selected_codigo="GDA-001-26",
                  selected_perforacion="S-1", selected_muestra_id="GDA-001-26-S-1-M1", selected_assay_id="a1",
                  selected_assay_type=tipo, read_only_view=False, _cookie_restore_attempted=True)
    for k, v in estado.items():
        at.session_state[k] = v
    at.run()
    if at.exception:
        raise RuntimeError(f"{tipo}: {at.exception[0].value}")
    salida = {}
    _recorrer(at.main, "", salida, "a1")
    return salida


def main():
    todas = {}
    for tipo, nombre, variantes in ENSAYOS:
        for extra in variantes:
            encontradas = etiquetas_de(tipo, nombre, extra)
            print(f"{tipo} {extra or ''}: {len(encontradas)} casillas", flush=True)
            for k, v in encontradas.items():
                todas.setdefault(k, v)
    ruta = os.path.join(RAIZ, "etiquetas_campos.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(todas.items())), f, ensure_ascii=False, indent=1)
    print(f"{len(todas)} etiquetas en {ruta}")


if __name__ == "__main__":
    main()
