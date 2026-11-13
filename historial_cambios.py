"""Historial de cambios de un ensayo, para consultarlo por fechas.

La base guarda en `assay_versions` una fila por cada cambio de `assays.data`:
  * filas nuevas (migración 0042): solo lo que cambió, `cambios` = {clave: {"a": antes, "d": después}}, y `autor` = quién lo cambió;
  * filas viejas: la versión completa ANTERIOR en `data` y `digitado_por` = de quién era esa versión.
Partiendo de los datos actuales y deshaciendo los cambios hacia atrás se sabe cómo estaba el ensayo en cada momento.
"""
import re
from datetime import datetime, timezone

_HORA = re.compile(r"^(.*?)_(\d{1,2})h(_.*)?$")


def datos_comparables(data):
    """Los datos del ensayo sin claves internas (_...), sin las de control (…_lastsync) y sin fotos (la base no las guarda en el historial)."""
    return {k: v for k, v in (data or {}).items()
            if not k.startswith("_") and not k.endswith("_lastsync") and not (isinstance(v, dict) and "b64" in v)}


def _vacio(v):
    return v is None or v == "" or v == []


def _numero(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def mismo_valor(a, b):
    """"12.5" y "12.50" son el mismo dato; vacío, None y [] también."""
    if _vacio(a) and _vacio(b):
        return True
    if isinstance(a, str) and isinstance(b, str):
        na, nb = _numero(a), _numero(b)
        if na is not None and nb is not None:
            return na == nb
    return a == b


def reconstruir_eventos(filas, datos_actuales, autor_actual=None, fecha_actual=None):
    """Cambios del más viejo al más nuevo: [{"fecha", "autor", "antes", "despues", "inicial"}].

    `filas`: las de assay_versions, la más reciente primero. Si lo más viejo guardado ya traía datos (historial anterior a guardar todos
    los cambios), se agrega al comienzo un evento «inicial» con esos datos."""
    despues = datos_comparables(datos_actuales)
    autor_despues = autor_actual
    eventos = []
    for fila in filas:
        cambios = fila.get("cambios")
        if isinstance(cambios, dict):
            antes = dict(despues)
            for clave, cambio in cambios.items():
                valor = (cambio or {}).get("a")
                if valor is None:
                    antes.pop(clave, None)
                else:
                    antes[clave] = valor
            antes = datos_comparables(antes)
            autor = fila.get("autor") or autor_despues
        else:
            antes = datos_comparables(fila.get("data"))
            autor = autor_despues
        eventos.append({"fecha": fila.get("saved_at"), "autor": autor, "antes": antes, "despues": despues, "inicial": False})
        despues = antes
        autor_despues = fila.get("digitado_por") or None
    eventos.reverse()
    if despues:   # lo más antiguo que se conoce ya tenía datos: se muestra cómo estaba
        mas_viejo = filas[-1] if filas else {}
        eventos.insert(0, {"fecha": mas_viejo.get("digitado_en") or (mas_viejo.get("saved_at") if filas else fecha_actual),
                           "autor": autor_despues if filas else autor_actual, "antes": {}, "despues": despues, "inicial": True})
    return eventos


def cambios_de(antes, despues):
    """{clave: (antes, después, horas copiadas)} de lo que cambió. Las pesadas por hora (…_14h, …_15h_1, …_16h_antes) que el formulario
    llena solas con la misma masa seca digitada no se cuentan aparte: se anotan en el dato que se digitó."""
    claves = sorted(k for k in set(antes) | set(despues) if not mismo_valor(antes.get(k), despues.get(k)))
    resultado = {k: [antes.get(k), despues.get(k), []] for k in claves}
    for k in claves:
        m = _HORA.match(k)
        if not m:
            continue
        prefijo, hora, sufijo = m.group(1), m.group(2), m.group(3) or ""
        origen = next((c for c in claves if c != k and not _HORA.match(c) and c.startswith(prefijo) and c.endswith(sufijo)
                       and mismo_valor(despues.get(c), despues.get(k))), None)
        if origen:
            resultado[origen][2].append(int(hora))
            resultado.pop(k)
    return {k: (a, d, sorted(h)) for k, (a, d, h) in resultado.items()}


def _instante(iso):
    try:
        f = datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
    except (TypeError, ValueError):
        return None
    return f if f.tzinfo else f.replace(tzinfo=timezone.utc)


def agrupar(eventos, minutos=20):
    """Junta los cambios seguidos de la misma persona (con menos de `minutos` entre uno y otro) en un «rato de trabajo»:
    [{"inicio", "fin", "autor", "inicial", "datos": {clave: {"valores": [antes, ..., final], "horas": [...]}}, "estado": datos al final}].
    Si un dato cambió varias veces en el mismo rato se ven todos sus valores en orden."""
    bloques = []
    for ev in eventos:
        cambios = cambios_de(ev["antes"], ev["despues"])
        if not cambios:
            continue
        cuando = _instante(ev["fecha"])
        ultimo = bloques[-1] if bloques else None
        junto = (ultimo is not None and not ev["inicial"] and not ultimo["inicial"] and ultimo["autor"] == ev["autor"]
                 and cuando is not None and ultimo["fin"] is not None and (cuando - ultimo["fin"]).total_seconds() <= minutos * 60)
        if not junto:
            ultimo = {"inicio": cuando, "fin": cuando, "autor": ev["autor"], "inicial": ev["inicial"], "datos": {}, "estado": None}
            bloques.append(ultimo)
        ultimo["fin"] = cuando or ultimo["fin"]
        ultimo["estado"] = ev["despues"]
        for clave, (a, d, horas) in cambios.items():
            dato = ultimo["datos"].setdefault(clave, {"valores": [a], "horas": []})
            if not mismo_valor(dato["valores"][-1], d):
                dato["valores"].append(d)
            dato["horas"] = sorted(set(dato["horas"]) | set(horas))
    for b in bloques:   # un dato que en el mismo rato volvió a su valor inicial y no pasó por otros no es un cambio
        b["datos"] = {k: v for k, v in b["datos"].items() if len(v["valores"]) > 1}
    return [b for b in bloques if b["datos"]]


def texto_valor(v):
    if _vacio(v):
        return "—"
    if isinstance(v, bool):
        return "Sí" if v else "No"
    if isinstance(v, list):
        if all(isinstance(x, (str, int, float)) for x in v):
            return ", ".join(str(x) for x in v)
        return f"{len(v)} lecturas"
    if isinstance(v, dict):
        return "(datos)"
    return str(v)
