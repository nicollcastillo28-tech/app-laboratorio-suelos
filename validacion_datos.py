"""Datos mal digitados en un ensayo: lo que no deja enviarlo a revisión.

Tres revisiones, en este orden:
  1. Errores de digitación en una casilla numérica: «11..1», «11-12», «12 5», «1.234,5»… (el Excel queda sin el dato o con el resultado mal).
  2. Rangos absurdos: golpes del Límite Líquido fuera del método (15 a 35 en el Método A, 20 a 30 en el B), golpes del CBR o del Proctor,
     capas, temperaturas, porcentajes, masas negativas, alturas o diámetros en cero.
  3. Masas que no cuadran entre sí: la masa seca + recipiente mayor que la húmeda + recipiente, el recipiente que pesa igual o más que el
     suelo seco + recipiente, o una parte que pesa más que el total (caras fracturadas, alargamiento y aplanamiento).

Las masas de los Límites que no cuadran ya las revisa app.limites_incoherentes (en amarillo): aquí no se repiten.
"""
import re

# Solo dígitos, puntos, comas, signos y espacios: lo que alguien quiso escribir como número. Un texto con letras («R-12», «110 ± 5»), con «/»
# (fechas, «3/4») o con «:» (horas) no se revisa como número.
_PARECE_NUMERO = re.compile(r"^[\d\s.,+\-]+$")
_FECHA_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}")
# Casillas que son un número de identificación (recipiente, picnómetro, molde…), no una medida: pueden llevar guiones.
_IDENTIFICADOR = re.compile(r"(^|_)(recipiente|picnometro|pic_no|molde|consolidometro)(_\d+|_antes|_despues|_inicial|_final)?$")

# (patrón de la clave, mínimo, máximo, entero, qué es) — el primero que coincide manda.
RANGOS = [
    (r"^cbrc_\d+_golpes$", 5, 75, True, "golpes por capa del CBR (lo normal: 10, 25 y 56)"),
    (r"^proc_\d+_golpes$", 10, 75, True, "golpes por capa del Proctor (lo normal: 25 o 56)"),
    (r"_capas$", 1, 10, True, "número de capas (lo normal: 3 o 5)"),
    (r"^mo_temp$", 200, 1100, False, "temperatura de ignición (°C)"),
    (r"(temp_secado|temp_horno)$", 30, 130, False, "temperatura de secado del horno (°C)"),
    (r"(_temp|_temperatura|_temp_ini|_temp_fin|_temp_inicial|_temp_final|_temp_agua)$", 5, 45, False, "temperatura del ensayo (°C)"),
    (r"(pct_retenido|sobretamano_pct)$", 0, 100, False, "porcentaje"),
    (r"^corte_hum_(inicial|final)$", 0, 500, False, "humedad (%)"),
]
_RANGOS = [(re.compile(p), mn, mx, ent, que) for p, mn, mx, ent, que in RANGOS]

# Golpes del Límite Líquido según el método (INV E-125): multipunto entre 15 y 35 (lo normal, cerca de 35, 25 y 15); un solo punto, 20 a 30.
GOLPES_LL = {"Método A": (15, 35), "Método B": (20, 30)}

_MASA = re.compile(r"(^|_)(masa|seco|humedo|peso)(_|$)|^s_(\d|p|fondo)|_(ret|una|dos|ini|par)$|_hum_masa_rec$|_recipiente_masa$")
_DIMENSION = re.compile(r"(^|_)(altura|diametro)(_|$)|^(ci|roca)_[dh]_\d+$")
_FUERZA = re.compile(r"_pen_\d+$")


def _numero(valor):
    try:
        return float(str(valor).strip().replace(",", "."))
    except (ValueError, TypeError):
        return None


def _bonito(valor):
    v = float(valor)
    return str(int(v)) if v == int(v) else f"{v:.3f}".rstrip("0").rstrip(".")


def _corta(etiqueta):
    """La casilla sin la tarjeta («Determinación de Humedad › Masa del recipiente (g)» → «Masa del recipiente (g)»)."""
    return etiqueta.split("›")[-1].strip()


def _triplos():
    """(húmeda + recipiente, patrón de las secas + recipiente, recipiente) de cada humedad de los ensayos."""
    yield "hum_masa_humedo_mas_recipiente", r"^hum_seco_(mas_recipiente|\d+h)$", "hum_masa_recipiente"
    for i in range(1, 4):
        for t in ("hc", "hd"):
            yield f"cbrc_{i}_{t}_masa_humedo", rf"^cbrc_{i}_{t}_seco_\d+h$", f"cbrc_{i}_{t}_masa_recipiente"
    for i in range(1, 7):
        yield f"proc_{i}_hum_masa_humedo", rf"^proc_{i}_hum_seco_\d+h$", f"proc_{i}_hum_masa_recipiente"
    for t in ("natural", "contraccion"):
        yield f"lc_{t}_hum_humedo", rf"^lc_{t}_hum_seco$", f"lc_{t}_hum_recipiente"
    yield "mo_masa_crisol_seco", r"^mo_masa_crisol_ignicion$", "mo_masa_crisol"
    for p in ("ci", "roca", "mu", "cp"):
        yield f"{p}_hum_humedo", rf"^{p}_hum_seco_\d+$", f"{p}_hum_masa_rec"
    for t in ("ini", "fin"):
        yield f"cons_{t}_humedo", rf"^cons_{t}_seco_\d+$", f"cons_{t}_recipiente_masa"
    for i in range(1, 4):
        for t in ("inicial", "final"):
            yield f"corte_m{i}_hum_masa_humedo_{t}", rf"^corte_m{i}_hum_seco_\d+h_{t}$", f"corte_m{i}_hum_masa_recipiente_{t}"


def _partes_del_total():
    """(parte, total): una parte de la muestra no puede pesar más que el total."""
    for n in range(19, 27):   # caras fracturadas: dos o más caras ≤ una o más caras ≤ masa retenida
        yield f"cf_{n}_una", f"cf_{n}_ret"
        yield f"cf_{n}_dos", f"cf_{n}_una"
    for t in ("alar", "apla"):
        for n in range(1, 8):
            yield f"aa_{t}_{n}_par", f"aa_{t}_{n}_ini"


def _se_ignora(tipo, data, clave):
    if clave.startswith("_") or clave.endswith("_lastsync") or "fecha" in clave or "hora" in clave:
        return True
    if tipo == "limites" and clave.startswith("lim_"):
        if data.get("lim_nl"):
            return True     # no líquido: no se digita nada de Límites
        if data.get("lim_np") and clave.startswith("lim_lp_"):
            return True     # no plástico: el Límite Plástico no se usa
    return False


def valores_mal_digitados(tipo, data, etiqueta=None):
    """[{clave, etiqueta, valor, motivo}] con lo mal digitado en los datos `data` de un ensayo de tipo `tipo`; vacío si todo está bien.
    `etiqueta(clave)` da el nombre de la casilla para el aviso."""
    etiqueta = etiqueta or (lambda k: k)
    data = data or {}
    salida, vistos = [], set()

    def agregar(clave, valor, motivo):
        if clave in vistos:
            return
        vistos.add(clave)
        salida.append({"clave": clave, "etiqueta": etiqueta(clave), "valor": "" if valor is None else str(valor).strip(), "motivo": motivo})

    numeros = {}
    for clave, valor in data.items():
        if not isinstance(valor, (str, int, float)) or isinstance(valor, bool) or _se_ignora(tipo, data, clave):
            continue
        texto = str(valor).strip()
        if not texto:
            continue
        n = _numero(texto)
        if n is None:
            # 1. Error de digitación: parece un número pero no lo es.
            if isinstance(valor, str) and _PARECE_NUMERO.match(texto) and re.search(r"\d", texto) and not _FECHA_ISO.match(texto) \
                    and not (_IDENTIFICADOR.search(clave) and "(g)" not in etiqueta(clave)):
                agregar(clave, texto, "no es un número válido: revisa si quedó un punto, una coma, un guion o un espacio de más")
            continue
        numeros[clave] = n

    # 2. Rangos absurdos.
    if tipo == "limites" and not data.get("lim_nl"):
        metodo = data.get("lim_metodo") if data.get("lim_metodo") in GOLPES_LL else "Método A"
        mn, mx = GOLPES_LL[metodo]
        for clave in sorted(k for k in numeros if re.match(r"^lim_ll_golpes_\d+$", k)):
            g = numeros[clave]
            if g != int(g) or not mn <= g <= mx:
                agregar(clave, data.get(clave), f"con el {metodo} los golpes deben ser un número entero entre {mn} y {mx}"
                        + (" (lo normal: cerca de 35, 25 y 15)" if metodo == "Método A" else ""))
    for clave, n in numeros.items():
        for patron, mn, mx, entero, que in _RANGOS:
            if patron.search(clave):
                if (entero and n != int(n)) or not mn <= n <= mx:
                    agregar(clave, data.get(clave), f"fuera de rango para {que}: debe ser {'un número entero ' if entero else ''}entre "
                                                    f"{_bonito(mn)} y {_bonito(mx)}")
                break
        else:
            texto_etq = etiqueta(clave)
            if n < 0 and (_MASA.search(clave) or "(g)" in texto_etq or _FUERZA.search(clave)):
                agregar(clave, data.get(clave), "no puede ser negativo")
            elif n <= 0 and _DIMENSION.search(clave):
                agregar(clave, data.get(clave), "una altura o un diámetro no puede ser cero ni negativo")

    # 3. Masas que no cuadran entre sí.
    for humeda, patron_secas, recipiente in _triplos():
        h, r = numeros.get(humeda), numeros.get(recipiente)
        secas = [k for k in numeros if re.match(patron_secas, k)]
        for seca in sorted(secas):
            s = numeros[seca]
            if h is not None and s > h + 1e-9:
                agregar(seca, data.get(seca), f"es mayor que «{_corta(etiqueta(humeda))}» ({_bonito(h)} g): la muestra no puede pesar más "
                                              "después de secarla")
            if r is not None and r >= s - 1e-9:
                agregar(recipiente, data.get(recipiente), f"es igual o mayor que «{_corta(etiqueta(seca))}» ({_bonito(s)} g): el recipiente "
                                                          "solo no puede pesar lo mismo o más que el recipiente con el suelo")
        if h is not None and r is not None and r >= h - 1e-9 and recipiente not in vistos:
            agregar(recipiente, data.get(recipiente), f"es igual o mayor que «{_corta(etiqueta(humeda))}» ({_bonito(h)} g)")
    for parte, total in _partes_del_total():
        p, t = numeros.get(parte), numeros.get(total)
        if p is not None and t is not None and p > t + 1e-9:
            agregar(parte, data.get(parte), f"es mayor que «{_corta(etiqueta(total))}» ({_bonito(t)} g): una parte no puede pesar más que el total")
    return salida
