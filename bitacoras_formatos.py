"""Cómo se pasa lo digitado de cada ensayo a su bitácora de papel (GDA-FL-xxx). Cada `llenar_*` recibe el contexto de la muestra (`ctx`) y los datos
del ensayo (`d`, el mismo diccionario que guarda la app) y devuelve (celdas, casillas): lo que se escribe y las casillas que se marcan."""
from bitacoras import Plantilla, Izq, Libre, llenar, xlsx_a_pdf, to_float, _num, _norm
import re

_CACHE = {}


def _plantilla(archivo):
    if archivo not in _CACHE:
        _CACHE[archivo] = Plantilla(archivo)
    return _CACHE[archivo]


def _fmt(v):
    """Texto de un número (sin ceros de más) o el texto tal cual."""
    return _num(v)


def _prof(ctx):
    de, a = ctx.get("prof_de"), ctx.get("prof_a")
    if de in (None, "") and a in (None, ""):
        return None
    f = lambda x: ("%g" % to_float(x)) if to_float(x) is not None else str(x or "")
    return f"{f(de)} - {f(a)} m"


def _ref(col, fila):
    return f"{col}{fila}"


def _cod_equipo(texto):
    m = re.search(r"GDA\s*-\s*E\s*-\s*(\d+)", str(texto), re.I)
    return m.group(1) if m else None


def casillas_equipos(P, equipos, desde_fila=1, hasta_fila=None):
    """Casillas de los equipos usados. Cada etiqueta de equipo de la bitácora lleva su código GDA-E-xxx (a veces dos en la misma celda, cada
    uno con su casilla): se marcan las casillas de los códigos que están en `equipos`, dentro del rango de filas dado.
    Cada código toma la casilla más cercana a su etiqueta (la que está justo a su derecha)."""
    codigos = {c for c in (_cod_equipo(e) for e in (equipos or [])) if c}
    if not codigos:
        return []
    etiquetas = []   # (fila, col, [códigos de la celda])
    for fila in P.ws.iter_rows(min_row=desde_fila, max_row=hasta_fila):
        for c in fila:
            if isinstance(c.value, str):
                en_celda = re.findall(r"GDA\s*-\s*E\s*-\s*(\d+)", c.value, re.I)
                if en_celda:
                    etiquetas.append((c.row, c.column, en_celda, c.value))
    if not etiquetas:
        return []
    ids = []
    for fila, col, en_celda, texto in etiquetas:
        if not (set(en_celda) & codigos):
            continue
        try:
            if len(en_celda) == 1:
                ids.append(P.casilla_de(texto, desde_fila=fila, hasta_fila=fila))
            else:
                cajas = P.casillas_en_celda(fila, col)
                for k, cod in enumerate(en_celda):
                    if cod in codigos and k < len(cajas):
                        ids.append(cajas[k])
        except KeyError:
            pass
    return ids


def todos_equipos(*datas):
    """Todos los equipos marcados en uno o varios ensayos (las listas guardadas en las claves *_equipos)."""
    salida = []
    for d in datas:
        for k, v in (d or {}).items():
            if k.endswith("_equipos") and isinstance(v, list):
                salida += v
    return salida


def encabezado(P, ctx, celdas, *, apique="APIQUE", muestra="MUESTRA:", profundidad="PROFUNDIDAD", realizo=None, gda=None):
    """Datos comunes del encabezado de todas las bitácoras (fecha, descripción, profundidad, localización, apique/sondeo, muestra y código)."""
    def poner(etq, valor, **kw):
        if valor in (None, ""):
            return
        try:
            celdas[P.valor_de(etq, **kw)] = Libre(valor)
        except KeyError:
            pass
    poner("FECHA DE EJECUCIÓN", ctx.get("fecha_ejecucion"))
    if ctx.get("descripcion"):
        try:
            celdas.update(P.lineas(P.valor_de("DESCRIPCIÓN VISUAL"), ctx["descripcion"], max_lineas=2))
        except KeyError:
            pass
    poner(profundidad, _prof(ctx))
    poner("LOCALIZACIÓN", ctx.get("localizacion"))
    poner(apique, f"{ctx.get('tipo_perf', '')} {ctx.get('perf_codigo', '')}".strip())
    poner(muestra, ctx.get("muestra"))
    # Código interno del proyecto: la celda que ya trae "GDA - ... - ..." (o "GDA:" y el guion en la celda siguiente)
    try:
        f, c = P.buscar("GDA")
        from openpyxl.utils import get_column_letter as L
        if ctx.get("numero"):
            txt = P.ws.cell(row=f, column=c).value or ""
            if gda:
                celdas[gda[0]] = ctx.get("numero")
                celdas[gda[1]] = ctx.get("anio")
            elif txt.strip().upper().startswith("GDA:"):
                celdas[L(c + 1) + str(f)] = ctx.get("numero")
            else:
                celdas[f"{L(c)}{f}"] = f"GDA - {ctx.get('numero')} - {ctx.get('anio')}"
    except KeyError:
        pass
    if realizo and ctx.get("realizo"):
        celdas[realizo] = Libre(f"NOMBRE: {ctx['realizo']}")


def _temp_hum(P, celdas, ini_temp, fin_temp, ini_hum=None, fin_hum=None, etq_t="TEMPERATURA", etq_h="HUMEDAD %"):
    for etq, a, b in ((etq_t, ini_temp, fin_temp), (etq_h, ini_hum, fin_hum)):
        try:
            ref = P.celda_a_la_derecha(_ref_de(P, etq))
        except KeyError:
            continue
        if a not in (None, ""):
            celdas[ref] = _fmt(a)
        if b not in (None, ""):
            celdas[P.celda_a_la_derecha(ref)] = _fmt(b)


def _ref_de(P, etq):
    from openpyxl.utils import get_column_letter as L
    f, c = P.buscar(etq)
    return f"{L(c)}{f}"


# ═══ GDA-FL-005 · Compresión inconfinada ═════════════════════════════════════════════════════════════════════════════════
def _ci_carga_en(d, def_mm):
    """Carga (kN) a esa deformación (mm): la digitada en la bitácora o, si se cargó la tabla de la máquina, interpolada en ella."""
    maq = d.get("ci_maq")
    if not maq:
        return None
    pts = sorted((float(x[2]), float(x[1])) for x in maq if x[1] is not None and x[2] is not None)
    if not pts or def_mm > pts[-1][0] + 1e-9:
        return None
    for (d0, f0), (d1, f1) in zip(pts, pts[1:]):
        if d0 <= def_mm <= d1:
            return f0 if d1 == d0 else f0 + (f1 - f0) * (def_mm - d0) / (d1 - d0)
    return pts[0][1]


def llenar_fl005(ctx, d, CI_DEFORMACIONES):
    P = _plantilla("GDA-FL-005_compresion_inconfinada.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, realizo="B56")
    _temp_hum(P, celdas, d.get("ci_temp_ini"), d.get("ci_temp_fin"))
    casillas.append(P.casilla_de({"Inalterada": "INALTERADA", "Compactada": "COMPACTADA", "Remodelada": "REMODELADA"}.get(d.get("ci_condicion"), "INALTERADA")))
    # Tabla deformación / carga (filas 19 a 53 de la plantilla, en el mismo orden de CI_DEFORMACIONES)
    filas = [f for f in range(19, 54) if f != 25 and f != 52]
    fila_def = {}
    for r in range(19, 54):
        v = P.ws.cell(row=r, column=3).value
        if v is not None:
            fila_def[r] = v
    for i, defo in enumerate(CI_DEFORMACIONES, start=1):
        fila = next((r for r, v in fila_def.items() if (isinstance(v, (int, float)) and v == defo)), None)
        if fila is None:
            continue
        carga = to_float(d.get(f"ci_carga_{i}"))
        if carga is None:
            carga = _ci_carga_en(d, defo * 0.0254)
        if carga is not None:
            celdas[f"E{fila}"] = round(carga, 3)
    for i in (1, 2, 3):
        if to_float(d.get(f"ci_h_{i}")) is not None:
            celdas[f"J{18 + i}"] = to_float(d.get(f"ci_h_{i}"))
        if to_float(d.get(f"ci_d_{i}")) is not None:
            celdas[f"K{18 + i}"] = to_float(d.get(f"ci_d_{i}"))
    hs = [to_float(d.get(f"ci_h_{i}")) for i in (1, 2, 3) if to_float(d.get(f"ci_h_{i}")) is not None]
    ds = [to_float(d.get(f"ci_d_{i}")) for i in (1, 2, 3) if to_float(d.get(f"ci_d_{i}")) is not None]
    if hs:
        celdas["J22"] = round(sum(hs) / len(hs), 3)
    if ds:
        celdas["K22"] = round(sum(ds) / len(ds), 3)
    celdas["J23"] = _fmt(d.get("ci_peso"))
    for r, key in zip(range(26, 32), ("ci_hum_recipiente", "ci_hum_humedo", "ci_hum_seco_17", "ci_hum_seco_18", "ci_hum_seco_19", "ci_hum_masa_rec")):
        celdas[f"L{r}"] = _fmt(d.get(key))
    celdas["K33"] = _fmt(d.get("ci_penetrometro"))
    celdas["O30"] = _fmt(d.get("ci_velocidad"))
    celdas["O31"] = _fmt(d.get("ci_tiempo_falla"))
    # Casillas: temperatura de secado, método, humedad obtenida, sobre qué
    casillas.append(P.casilla_de("60", desde_fila=37, hasta_fila=37) if "60" in str(d.get("ci_temp_secado", "")) else P.casilla_de("110", desde_fila=38, hasta_fila=38))
    casillas.append(P.casilla_de("A", desde_fila=37, hasta_fila=37) if d.get("ci_metodo", "A") == "A" else P.casilla_de("B", desde_fila=38, hasta_fila=38))
    casillas.append(P.casilla_de("ANTES DEL ENSAYO") if d.get("ci_hum_antes", "Antes del ensayo").startswith("Antes") else P.casilla_de("DESPUÉS DEL ENSAYO"))
    casillas.append(P.casilla_de("MUESTRA COMPLETA") if d.get("ci_hum_muestra", "Muestra completa").startswith("Muestra") else P.casilla_de("CORTES DE MUESTRA"))
    casillas += casillas_equipos(P, d.get("ci_equipos"), desde_fila=40)
    if d.get("observaciones"):
        celdas.update(P.lineas("J44", d["observaciones"], max_lineas=5))
    return P, celdas, casillas


def _poner(celdas, ref, valor, dec=None):
    v = _num(valor, dec)
    if v not in (None, ""):
        celdas[ref] = v


def _tabla_humedad(celdas, d, col, filas, claves):
    """Escribe una columna de datos de humedad: `filas` = filas de la plantilla, `claves` = claves de `d` en el mismo orden."""
    for fila, clave in zip(filas, claves):
        if clave:
            _poner(celdas, f"{col}{fila}", d.get(clave))


# ═══ GDA-FL-008 · Proctor y CBR (compactado) ════════════════════════════════════════════════════════════════════════════
def llenar_fl008(ctx, d, CBRC_GOLPES):
    P = _plantilla("GDA-FL-008_proctor_cbr.xlsx")
    celdas, casillas = {}, {}
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", muestra="MUESTRA", realizo="A44")
    # Profundidad: la plantilla trae el guion entre dos casillas (B8 y D8)
    celdas.pop(P.valor_de("PROFUNDIDAD"), None)
    _poner(celdas, "B8", ctx.get("prof_de"))
    _poner(celdas, "D8", ctx.get("prof_a"))
    _temp_hum(P, celdas, d.get("proc_temp_ini"), d.get("proc_temp_fin"), d.get("proc_hum_ini"), d.get("proc_hum_fin"))
    # Proctor (columnas B a E = pruebas 1 a 4)
    for i, col in enumerate("BCDE", start=1):
        v = lambda c: d.get(f"proc_{i}_{c}")
        for fila, campo in ((18, "golpes"), (19, "molde"), (20, "capas"), (21, "masa_humedo_molde"), (22, "masa_molde"), (23, "volumen_molde"),
                            (27, "hum_recipiente"), (28, "hum_masa_humedo"), (29, "hum_seco_16h"), (30, "hum_seco_17h"), (31, "hum_seco_18h"),
                            (32, "hum_seco_19h"), (33, "hum_masa_recipiente")):
            _poner(celdas, f"{col}{fila}", v(campo))
    # CBR compactado (columnas H a J = moldes 1 a 3)
    for i, col in enumerate("HIJ", start=1):
        c = lambda k: d.get(f"cbrc_{i}_{k}")
        _poner(celdas, f"{col}18", c("golpes") or CBRC_GOLPES[i - 1])
        for fila, campo in ((19, "molde"), (20, "capas"), (21, "masa_muestra_molde"), (22, "masa_molde"), (23, "altura"), (24, "diametro"),
                            (27, "hc_recipiente"), (28, "hc_masa_humedo"), (29, "hc_seco_16h"), (30, "hc_seco_17h"), (31, "hc_seco_18h"),
                            (32, "hc_seco_19h"), (33, "hc_masa_recipiente")):
            _poner(celdas, f"{col}{fila}", c(campo))
    # Expansión y humedad después de la inmersión (columnas M a O = moldes 1 a 3)
    for i, col in enumerate("MNO", start=1):
        c = lambda k: d.get(f"cbrc_{i}_{k}")
        _poner(celdas, f"{col}18", c("exp_inicial"))
        _poner(celdas, f"{col}20", c("exp_final"))
        for fila, campo in ((24, "hd_recipiente"), (25, "hd_masa_humedo"), (26, "hd_seco_16h"), (27, "hd_seco_17h"), (28, "hd_seco_18h"),
                            (29, "hd_seco_19h"), (30, "hd_masa_recipiente")):
            _poner(celdas, f"{col}{fila}", c(campo))
        _poner(celdas, f"{col}32", c("molde"))
    # Norma, método, sobretamaños y equipos
    norma = str(d.get("proc_norma", ""))
    if "141" in norma:
        casillas.append(P.casilla_de("INV E - 141"))
    elif "142" in norma:
        casillas.append(P.casilla_de("INV E - 142"))
    metodo = d.get("proc_metodo")
    if metodo in ("A", "B", "C"):
        casillas.append(P.casilla_de(metodo, desde_fila=35 + "ABC".index(metodo), hasta_fila=35 + "ABC".index(metodo)))
    tamiz = d.get("proc_sobretamano_tamiz")
    if tamiz and tamiz != "—":
        _poner(celdas, "B39", tamiz)
        _poner(celdas, "E39", d.get("proc_sobretamano_pct"))
    casillas += casillas_equipos(P, todos_equipos(d), desde_fila=35, hasta_fila=39)
    return P, celdas, casillas


# ═══ GDA-FL-004 · Clasificación (pasa 200, granulometría, humedad y peso unitario parafinado) ═════════════════════════════
FL004_TAMICES = {"s_3": 14, "s_2p5": 15, "s_2": 16, "s_1p5": 17, "s_1": 18, "s_34": 19, "s_12": 20, "s_38": 21, "s_4": 22, "s_10": 24,
                 "s_20": 25, "s_40": 26, "s_60": 27, "s_100": 29, "s_200": 30}


def llenar_fl004(ctx, gran, hum, mu, NORMAS_GRAN):
    P = _plantilla("GDA-FL-004_clasificacion.xlsx")
    gran, hum, mu = gran or {}, hum or {}, mu or {}
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, realizo="B30")
    for ref in ("B50", "M47"):
        if ctx.get("realizo"):
            from bitacoras import Libre
            celdas[ref] = Libre(f"NOMBRE: {ctx['realizo']}")
    datos_temp = [d.get(k) for d in (gran, hum, mu) for k in ("temp_ini", "temp_fin") if d.get(k)]
    # Pasa 200 (antes / después del lavado) y granulometría
    for suf, col in (("antes", "G"), ("despues", "J")):
        for fila, campo in ((16, "recipiente"), (17, "seco_mas_recipiente"), (18, "seco_14h"), (19, "seco_15h"), (20, "seco_16h"), (21, "masa_recipiente")):
            _poner(celdas, f"{col}{fila}", gran.get(f"p200_{campo}_{suf}"))
    for clave, fila in FL004_TAMICES.items():
        if clave in gran:
            v = to_float(gran.get(clave))
            celdas[f"P{fila}"] = v if v is not None else 0
    norma = gran.get("gran_norma")
    if norma in NORMAS_GRAN:
        cajas = P.casillas_en_celda(12, 2)  # B12 (combinada) cubre la primera casilla; las otras dos están sobre G12 e I12
        cajas = sorted({*cajas, *P.casillas_en_celda(12, 7), *P.casillas_en_celda(12, 9)}, key=lambda i: next(f["x"] for f in P.formas if f["id"] == i))
        k = NORMAS_GRAN.index(norma)
        if k < len(cajas):
            casillas.append(cajas[k])
    casillas += casillas_equipos(P, gran.get("gran_equipos"), 23, 28)
    # Humedad
    for fila, clave in ((35, "hum_recipiente"), (36, "hum_masa_humedo_mas_recipiente"), (39, "hum_seco_mas_recipiente"), (41, "hum_masa_recipiente")):
        _poner(celdas, f"G{fila}", hum.get(clave))
    cajas_met = P.casillas_en_celda(34, 2)
    if hum.get("hum_metodo") in ("Método A", "Método B") and len(cajas_met) >= 2:
        casillas.append(cajas_met[0 if hum["hum_metodo"] == "Método A" else 1])
    casillas += casillas_equipos(P, hum.get("hum_equipos"), 44, 48)
    # Peso unitario parafinado
    if mu.get("mu_metodo", "Parafinado") == "Parafinado":
        for fila, clave in ((35, "mu_peso_aire"), (36, "mu_peso_aire_par"), (37, "mu_peso_agua_par"), (38, "mu_temp_agua")):
            _poner(celdas, f"P{fila}", mu.get(clave))
        casillas += casillas_equipos(P, mu.get("mu_equipos"), 40, 44)
    obs = [d.get("observaciones") for d in (gran, hum, mu) if d.get("observaciones")]
    if obs:
        celdas.update(P.lineas("M51", " ".join(obs), max_lineas=2))
    return P, celdas, casillas


def _x_en_equipos(P, celdas, equipos, filas, col_etiqueta, col_casilla):
    """Cuando las casillas de equipos son celdas (no formas): pone una X en la celda de la fila cuyo equipo (por código GDA-E-xxx) se usó."""
    codigos = {c for c in (_cod_equipo(e) for e in (equipos or [])) if c}
    from openpyxl.utils import column_index_from_string as N
    for fila in filas:
        texto = P.ws.cell(row=fila, column=N(col_etiqueta)).value
        if isinstance(texto, str) and _cod_equipo(texto) in codigos:
            celdas[f"{col_casilla}{fila}"] = "X"


# ═══ GDA-FL-006 · Corte directo ══════════════════════════════════════════════════════════════════════════════════════════
def llenar_fl006(ctx, d):
    P = _plantilla("GDA-FL-006_corte_directo.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, realizo="B48")
    _poner(celdas, "H10", ctx.get("muestra"))
    celdas["H10"] = Libre(ctx["muestra"]) if ctx.get("muestra") else None
    _temp_hum(P, celdas, d.get("corte_temp_inicial"), d.get("corte_temp_final"), d.get("corte_hum_inicial"), d.get("corte_hum_final"))
    if d.get("corte_tipo") in ("CD", "CU", "UU"):
        casillas.append(P.casilla_de(d["corte_tipo"]))
    cond = str(d.get("corte_condicion", ""))
    for etiqueta in ("INALTERADA", "REMOLDADA", "COMPACTADA"):
        if cond.upper().startswith(etiqueta[:5]):
            casillas.append(P.casilla_de(etiqueta))
    for k, (c1, c2) in enumerate((("D", "E"), ("I", "J"), ("N", "O")), start=1):
        v = lambda campo: d.get(f"corte_m{k}_{campo}")
        for fila, campo in ((20, "masa_inicial_anillo"), (21, "masa_anillo"), (22, "altura_anillo"), (23, "diametro_anillo"),
                            (25, "velocidad_corte"), (26, "esfuerzo_normal")):
            _poner(celdas, f"{c1}{fila}", v(campo))
        for col, suf in ((c1, "inicial"), (c2, "final")):
            for fila, campo in ((30, "hum_recipiente"), (31, "hum_masa_humedo"), (32, "hum_seco_17h"), (34, "hum_seco_18h"), (35, "hum_seco_19h"),
                                (36, "hum_masa_recipiente")):
                _poner(celdas, f"{col}{fila}", v(f"{campo}_{suf}"))
        if v("temp_secado"):
            celdas[f"{c2}{39 if '60' in str(v('temp_secado')) else 40}"] = "X"
        _x_en_equipos(P, celdas, d.get(f"corte_m{k}_hum_equipos"), range(43, 47), {"D": "B", "I": "G", "N": "L"}[c1], c2)
    for fila, campo in ((20, "masa_suelo_seco"), (21, "masa_pic_muestra"), (22, "masa_pic"), (23, "temperatura")):
        _poner(celdas, f"U{fila}", d.get(f"corte_ge_{campo}"))
    _poner(celdas, "T26", d.get("corte_ge_picnometro"))
    casillas += casillas_equipos(P, d.get("corte_ge_equipos"), 25, 32)
    casillas += casillas_equipos(P, d.get("corte_equipos"), 33, 37)
    return P, celdas, casillas


# ═══ GDA-FL-007 · Compresión en roca ═════════════════════════════════════════════════════════════════════════════════════
def _carga_en_maquina(maq, def_mm):
    """Carga (kN) a esa deformación (mm) interpolada en la tabla de la máquina [[tiempo, fuerza, deformación], ...]."""
    if not maq:
        return None
    pts = sorted((float(x[2]), float(x[1])) for x in maq if x[1] is not None and x[2] is not None)
    if not pts or def_mm > pts[-1][0] + 1e-9:
        return None
    for (d0, f0), (d1, f1) in zip(pts, pts[1:]):
        if d0 <= def_mm <= d1:
            return f0 if d1 == d0 else f0 + (f1 - f0) * (def_mm - d0) / (d1 - d0)
    return pts[0][1]


def llenar_fl007(ctx, d):
    P = _plantilla("GDA-FL-007_compresion_roca.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, realizo="B55")
    _temp_hum(P, celdas, d.get("roca_temp_ini"), d.get("roca_temp_fin"))
    for fila in range(19, 53):
        deformacion = to_float(P.ws.cell(row=fila, column=5).value)   # 0.001 in
        if deformacion is None:
            continue
        carga = _carga_en_maquina(d.get("roca_maq"), deformacion * 0.0254)
        if carga is not None:
            celdas[f"C{fila}"] = round(carga, 3)
    for i in (1, 2, 3):
        _poner(celdas, f"J{18 + i}", d.get(f"roca_h_{i}"))
        _poner(celdas, f"K{18 + i}", d.get(f"roca_d_{i}"))
    hs = [to_float(d.get(f"roca_h_{i}")) for i in (1, 2, 3) if to_float(d.get(f"roca_h_{i}")) is not None]
    ds = [to_float(d.get(f"roca_d_{i}")) for i in (1, 2, 3) if to_float(d.get(f"roca_d_{i}")) is not None]
    if hs:
        celdas["J22"] = round(sum(hs) / len(hs), 3)
    if ds:
        celdas["K22"] = round(sum(ds) / len(ds), 3)
    _poner(celdas, "J23", d.get("roca_peso"))
    for fila, clave in zip(range(26, 32), ("roca_hum_recipiente", "roca_hum_humedo", "roca_hum_seco_17", "roca_hum_seco_18", "roca_hum_seco_19", "roca_hum_masa_rec")):
        _poner(celdas, f"L{fila}", d.get(clave))
    _poner(celdas, "P17", d.get("roca_esfuerzo_maximo_manual"))
    _poner(celdas, "P19", d.get("roca_tiempo_falla"))
    if d.get("roca_falla"):
        celdas["P20"] = Libre(str(d["roca_falla"]))
    casillas.append(P.casilla_de("60", desde_fila=37, hasta_fila=37) if "60" in str(d.get("roca_temp_secado", "")) else P.casilla_de("110", desde_fila=38, hasta_fila=38))
    casillas.append(P.casilla_de("A", desde_fila=37, hasta_fila=37) if d.get("roca_metodo", "A") == "A" else P.casilla_de("B", desde_fila=38, hasta_fila=38))
    casillas.append(P.casilla_de("ANTES DEL ENSAYO") if str(d.get("roca_hum_antes", "Antes")).startswith("Antes") else P.casilla_de("DESPUÉS DEL ENSAYO"))
    casillas.append(P.casilla_de("MUESTRA COMPLETA") if str(d.get("roca_hum_muestra", "Muestra completa")).startswith("Muestra") else P.casilla_de("CORTES DE MUESTRA"))
    casillas += casillas_equipos(P, d.get("roca_equipos"), 40)
    if d.get("observaciones"):
        celdas.update(P.lineas("J44", d["observaciones"], max_lineas=5))
    return P, celdas, casillas


# ═══ GDA-FL-017 · Consolidación unidimensional ═══════════════════════════════════════════════════════════════════════════
def llenar_fl017(ctx, d):
    P = _plantilla("GDA-FL-017_consolidacion.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", muestra="MUESTRA", realizo="A39")
    _temp_hum(P, celdas, d.get("cons_temp_ini"), d.get("cons_temp_fin"))
    cond = str(d.get("cons_condicion", ""))
    for etiqueta in ("INALTERADA", "COMPACTADA", "REMOLDADA"):
        if cond.upper().startswith(etiqueta[:5]):
            casillas.append(P.casilla_de(etiqueta))
    for ref, clave in (("D14", "cons_masa_anillo_muestra_ini"), ("D15", "cons_masa_anillo"), ("D16", "cons_masa_anillo_muestra_fin"),
                       ("D17", "cons_diametro"), ("D19", "cons_altura")):
        _poner(celdas, ref, d.get(clave))
    for col, pref in (("C", "cons_ini"), ("G", "cons_fin")):
        for fila, clave in ((22, "recipiente"), (23, "humedo"), (24, "seco_17"), (25, "seco_18"), (26, "seco_19"), (27, "recipiente_masa")):
            _poner(celdas, f"{col}{fila}", d.get(f"{pref}_{clave}"))
    casillas.append(P.casilla_de("A", desde_fila=30, hasta_fila=30) if d.get("cons_metodo", "A") == "A" or "60" in str(d.get("cons_temp_secado", "")) else P.casilla_de("B", desde_fila=31, hasta_fila=31))
    for ref, clave in (("M14", "cons_pic_masa_agua_suelo"), ("M15", "cons_pic_temp"), ("M16", "cons_pic_masa_seco")):
        _poner(celdas, ref, d.get(clave))
    if d.get("cons_pic_no"):
        celdas["F31"] = Libre(f"PICNÓMETRO N. {d['cons_pic_no']}")
    _poner(celdas, "M31", d.get("cons_consolidometro"))
    for i, fila in enumerate(range(21, 28), start=1):
        _poner(celdas, f"K{fila}", d.get(f"cons_def_{i}_carga"))
        _poner(celdas, f"L{fila}", d.get(f"cons_def_{i}_descarga"))
    _poner(celdas, "L28", d.get("cons_precarga"))
    casillas += casillas_equipos(P, d.get("cons_equipos"), 30, 34)
    if d.get("observaciones"):
        celdas.update(P.lineas("B35", d["observaciones"], max_lineas=2))
    return P, celdas, casillas


# ═══ GDA-FL-019 · Gravedad específica ════════════════════════════════════════════════════════════════════════════════════
def llenar_fl019(ctx, d):
    P = _plantilla("GDA-FL-019_gravedad_especifica.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", muestra="MUESTRA:", profundidad="PROFUNDIDAD (m)", gda=("AF6", "AJ6"))
    celdas.pop(P.valor_de("PROFUNDIDAD (m)"), None)
    _poner(celdas, "F11", ctx.get("prof_de"))
    _poner(celdas, "J11", ctx.get("prof_a"))
    # Pasa el tamiz No. 4 (finos)
    for ref, clave in (("U18", "gesp_f_masa_sss"), ("U19", "gesp_f_masa_pic_agua_suelo"), ("U20", "gesp_f_masa_pic_agua"), ("U22", "gesp_f_temp"),
                       ("U24", "gesp_f_masa_seco"), ("U26", "gesp_f_pct_retenido")):
        _poner(celdas, ref, d.get(clave))
    casillas += casillas_equipos(P, d.get("gesp_finos_equipos"), 29, 37)
    # Retiene el tamiz No. 4 (gruesos)
    for ref, clave in (("T48", "gesp_g_masa_sss"), ("T50", "gesp_g_masa_sumergido"), ("T52", "gesp_g_masa_seco"), ("T54", "gesp_g_temp")):
        _poner(celdas, ref, d.get(clave))
    casillas += casillas_equipos(P, d.get("gesp_gruesos_equipos"), 57, 65)
    if ctx.get("realizo"):
        for ref in ("J40", "J69"):
            celdas[ref] = Libre(f"NOMBRE: {ctx['realizo']}")
        celdas["A79"] = None
    return P, celdas, casillas


# ═══ GDA-FL-020 · Carga puntual ══════════════════════════════════════════════════════════════════════════════════════════
def llenar_fl020(ctx, d):
    P = _plantilla("GDA-FL-020_carga_puntual.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="PERFORACIÓN", muestra="MUESTRA", profundidad="PROFUNDIDAD (m)", gda=("AF6", "AJ6"))
    celdas.pop(P.valor_de("PROFUNDIDAD (m)"), None)
    _poner(celdas, "V14", ctx.get("prof_de"))
    _poner(celdas, "Z14", ctx.get("prof_a"))
    _temp_hum(P, celdas, d.get("cp_temp_ini"), d.get("cp_temp_fin"))
    for n, fila in enumerate(range(37, 57, 2), start=1):
        for col, clave in (("E", "carga"), ("J", "d"), ("O", "l1"), ("AB", "masa"), ("AG", "sentido")):
            _poner(celdas, f"{col}{fila}", d.get(f"cp_{n}_{clave}"))
    for fila, clave in ((61, "cp_hum_recipiente"), (63, "cp_hum_humedo"), (65, "cp_hum_seco_14"), (67, "cp_hum_seco_15"), (69, "cp_hum_seco_16"),
                        (71, "cp_hum_masa_rec")):
        _poner(celdas, f"K{fila}", d.get(clave))
    casillas += casillas_equipos(P, d.get("cp_equipos"), 59, 69)
    if ctx.get("realizo"):
        celdas["A80"] = Libre(f"NOMBRE: {ctx['realizo']}")
    if d.get("observaciones"):
        celdas.update(P.lineas("F74", d["observaciones"], max_lineas=2))
    return P, celdas, casillas


# ═══ GDA-FL-022 · Terrones de arcilla y partículas deleznables (la parte de partículas fracturadas no se digita en la app) ══
def llenar_fl022(ctx, d):
    P = _plantilla("GDA-FL-022_caras_fracturadas_terrones.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", muestra="MUESTRA No", realizo="G39")
    celdas.pop("G39", None)
    if ctx.get("realizo"):
        celdas["H39"] = Libre(ctx["realizo"])
    for fila, clave in zip(range(31, 36), (18, 19, 20, 21, 22)):
        _poner(celdas, f"G{fila}", d.get(f"ter_{clave}_ini"))
        _poner(celdas, f"I{fila}", d.get(f"ter_{clave}_fin"))
    casillas += casillas_equipos(P, d.get("ter_equipos"), 37, 40)
    if d.get("observaciones"):
        celdas.update(P.lineas("H36", d["observaciones"], max_lineas=2))
    return P, celdas, casillas


# ═══ GDA-FL-043 · Solidez de agregados en sulfatos ═══════════════════════════════════════════════════════════════════════
def llenar_fl043(ctx, d):
    P = _plantilla("GDA-FL-043_solidez_sulfatos.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="PERFORACIÓN", muestra="MUESTRA", profundidad="PROFUNDIDAD (m)", gda=("AF6", "AJ6"))
    celdas.pop(P.valor_de("PROFUNDIDAD (m)"), None)
    _poner(celdas, "V13", ctx.get("prof_de"))
    _poner(celdas, "Z13", ctx.get("prof_a"))
    _temp_hum(P, celdas, d.get("sulf_temp_ini"), d.get("sulf_temp_fin"))
    for fila, clave in zip((21, 23, 25, 27, 29), (21, 22, 23, 24, 25)):          # agregado grueso
        _poner(celdas, f"R{fila}", d.get(f"sulf_{clave}_ini"))
        _poner(celdas, f"AH{fila}", d.get(f"sulf_{clave}_fin"))
    for fila, clave in zip((38, 40, 42, 44, 46, 48), (30, 31, 32, 33, 34, 35)):   # agregado fino
        _poner(celdas, f"R{fila}", d.get(f"sulf_{clave}_ini"))
        _poner(celdas, f"AH{fila}", d.get(f"sulf_{clave}_fin"))
    casillas += casillas_equipos(P, d.get("sulf_equipos"), 63, 70)
    if ctx.get("realizo"):
        celdas["A81"] = Libre(f"NOMBRE: {ctx['realizo']}")
    return P, celdas, casillas


# ═══ GDA-FL-046 · Índices de aplanamiento y alargamiento ═════════════════════════════════════════════════════════════════
def llenar_fl046(ctx, d):
    P = _plantilla("GDA-FL-046_alargamiento_aplanamiento.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", muestra="MUESTRA:", realizo=None)
    for i, fila in enumerate(range(16, 23), start=1):
        _poner(celdas, f"D{fila}", d.get(f"aa_alar_{i}_ini"))
        _poner(celdas, f"E{fila}", d.get(f"aa_alar_{i}_par"))
        _poner(celdas, f"F{fila}", d.get(f"aa_apla_{i}_ini"))
        _poner(celdas, f"G{fila}", d.get(f"aa_apla_{i}_par"))
    casillas += casillas_equipos(P, d.get("aa_equipos"), 31, 35)
    if ctx.get("realizo"):
        try:
            celdas[P.valor_de("NOMBRE:", desde_fila=37)] = Libre(ctx["realizo"])
        except KeyError:
            pass
    if d.get("observaciones"):
        try:
            celdas.update(P.lineas(P.valor_de("observaciones"), d["observaciones"], max_lineas=3))
        except KeyError:
            pass
    return P, celdas, casillas


# ═══ GDA-FLC-029 · Peso unitario método B ════════════════════════════════════════════════════════════════════════════════
def llenar_flc029(ctx, d):
    P = _plantilla("GDA-FLC-029_bitacora_peso_unitario_b.xlsx")
    celdas, casillas = {}, []
    encabezado(P, ctx, celdas, apique="APIQUE", profundidad="PROFUNDIDAD (m)")
    celdas.pop(P.valor_de("PROFUNDIDAD (m)"), None)
    _poner(celdas, "F11", ctx.get("prof_de"))
    _poner(celdas, "J11", ctx.get("prof_a"))
    _poner(celdas, "N28", d.get("mub_altura"))
    _poner(celdas, "W28", d.get("mub_diametro"))
    _poner(celdas, "T31", d.get("mub_masa"))
    casillas += casillas_equipos(P, d.get("mub_equipos"), 36, 44)
    if ctx.get("realizo"):
        celdas["J55"] = Libre(f"NOMBRE: {ctx['realizo']}")
    return P, celdas, casillas
