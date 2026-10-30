"""Alertas de resultados "fuera de lo esperado" para los ensayos de laboratorio.

Módulo PURO (sin Streamlit ni acceso a la base de datos): recibe los valores ya calculados de un ensayo (`valores`),
el material de la muestra (`material`) y algo de contexto de los demás ensayos de la misma muestra (`ctx`), y devuelve
una lista de alertas. `app.py` se encarga de sacar esos valores de los datos digitados y de mostrarlas.

Dos niveles:
  - "revisar": el resultado se sale del rango HABITUAL del material -- puede ser real (suelo atípico), pero conviene
    consultar/revisar los datos antes de aprobarlo.
  - "alto": el resultado es físicamente improbable o imposible (humedad negativa, punto sobre la línea U de Casagrande,
    densidad seca por encima de la curva de saturación...) -- casi seguro hay un error de digitación, de pesada o de ensayo.

IMPORTANTE: los umbrales son RANGOS TÍPICOS de la literatura geotécnica (ver FUENTES), no criterios de aceptación ni
de norma -- salvo los marcados como "especificación" (terrones de arcilla INV 320-13, solidez ASTM C33). Un resultado
alertado no es un resultado incorrecto: es uno que merece una segunda mirada. Los umbrales viven todos en las tablas
de abajo para que el Jefe/Director Técnico los puedan ajustar sin tocar la lógica.
"""
import math

# ════════════════════════════════════════════════════════════════════
# FUENTES (consultadas en la web; `verificada` = se leyó la página/tabla, `secundaria` = apuntes/resúmenes de
# búsqueda que repiten la tabla original -- conviene contrastarlas con el texto de la norma o el libro citado).
# ════════════════════════════════════════════════════════════════════
FUENTES = {
    "agua_suelos": {
        "titulo": "Contenidos de humedad naturales típicos por tipo de suelo (manual de laboratorio de mecánica de suelos)",
        "url": "https://alg.manifoldapp.org/system/actioncallout/3/7/d/37dbe2c2-dc6f-48cd-b6c1-6dd4dc9bc189/attachment/05cfc782c0169d034c8ab614c3e92f08.docx",
        "nota": "Arenas/gravas hasta ~15-20 %; arena suelta saturada 25-30 %, densa 12-16 %; arcilla dura ~20 %, blanda 30-50 %; "
                "suelos finos naturales hasta 50-80 %; arcilla orgánica blanda 80-130 %; turbas y suelos muy orgánicos hasta ~500 %.",
        "estado": "secundaria"},
    "agua_minima": {
        "titulo": "Humedad mínima razonable: el agua higroscópica de un suelo SECADO AL AIRE ya es ≈0.3–6.3 % (más alta en suelos arcillosos); "
                  "una arcilla natural no puede tener menos que eso",
        "url": "https://publications.goettingen-research-online.de/ID-26373",
        "nota": "Estudio de suelos europeos (agua de muestras secadas al aire 2.8–63.3 g/kg y relación lineal con el contenido de arcilla). "
                "En arenas secas al aire la humedad es <0.1 % (apunte de enseñanza hallado en la búsqueda, no se pudo fijar la página exacta). "
                "Además, bajo el límite de contracción (≈12–25 % en arcillas, ver 'limite_contraccion') el suelo ya está seco. "
                "Los umbrales de la app (arcilla <6 % revisar, <2 % improbable; limo <5 % y <1.5 %; arena/grava <0.1 %; roca <0.01 %) "
                "son la lectura práctica de esos rangos, no valores publicados tal cual.",
        "estado": "secundaria"},
    "porosidad_roca": {
        "titulo": "Porosidad / absorción típica de roca intacta (granito 0.1-0.5 %, basalto 0.1-1 %, arenisca 5-25 %, caliza 5-20 %, lutita ~8 %)",
        "url": "https://lbas.lboro.ac.uk/chart/porosity-ratings-natural-stone-chart.html",
        "nota": "Complementa: https://lkouniv.ac.in/site/writereaddata/siteContent/202004032240236202gaurav_sriv_engg_Geology_Engg_properties_of_rocks.pdf. "
                "La humedad máxima posible de una roca es su humedad de saturación (porosidad / densidad seca): con n = 25 % y "
                "ρd ≈ 2.0 g/cm³ da ≈ 12 %.",
        "estado": "secundaria"},
    "linea_u": {
        "titulo": "Carta de plasticidad de Casagrande: línea A (PI = 0.73·(LL−20)) y línea U (PI = 0.9·(LL−8)), límite superior empírico; "
                  "un punto por encima de la línea U debe verificarse",
        "url": "https://uomus.edu.iq/img/lectures21/MUCLecture_2023_112510526.pdf",
        "nota": "Línea A en ASTM D2487: https://softbank.iust.ac.ir/Railway/Codes/ASTM-standard/ASTM_2004_Standards/ASTM%202004%20PART%20IB/ASTM%202004%20Volume%2004.08%20Soil%20and%20Rock%20(I)%20D%20420%20-%20D%205611/D%202487%20-%2098%20%20_RDI0ODCTUKVE.pdf",
        "estado": "secundaria"},
    "proctor_rangos": {
        "titulo": "Rangos típicos de compactación: γd,máx 1.6-2.0 Mg/m³ (extremos 1.3-2.4) y w óptima 10-20 % (extremos 5-40 %)",
        "url": "https://www.sciencedirect.com/science/article/abs/pii/S0013795220317270",
        "nota": "\"An overview on the compaction characteristics of soils by laboratory tests\" (Engineering Geology). Por tipo: NCHRP Report 544 "
                "(VDOT): granulares 1.84-2.16 g/cm³ y w ≈ 10-12 %; limos-arcillas 1.36-1.84 g/cm³ y w ≈ 15-20 %. "
                "http://www.extranet.vdot.state.va.us/locdes/hydraulic_design/nchrp_rpt544/content/html/Special_Topics/Soil_Compaction.html",
        "estado": "secundaria"},
    "cbr_uscs": {
        "titulo": "CBR típico por clase USCS (GW 40-80, GP 30-60, GM 20-60, GC 20-40, SW 20-40, SP 10-40, SM 10-40, SC 5-20, ML ≤15, CL ≤15, OL ≤5, MH ≤10, CH ≤15, OH ≤5)",
        "url": "https://pavementinteractive.org/reference-desk/design/design-parameters/california-bearing-ratio/",
        "nota": "Tabla leída directamente de la página (Pavement Interactive); la página no cita la fuente primaria de la tabla.",
        "estado": "verificada"},
    "isrm_ucs": {
        "titulo": "Clasificación ISRM de la resistencia de la roca intacta (R0 0.25-1 MPa, R1 1-5, R2 5-25, R3 25-50, R4 50-100, R5 100-250, R6 >250); "
                  "Is(50) R4 2-4, R5 4-10, R6 >10 MPa",
        "url": "https://rockmass.net/files/compr_strength_table.pdf",
        "nota": "También: BGS https://earthwise.bgs.ac.uk/index.php/OR/15/065_Appendix_2_-_Strength,_consistency_and_density_classification_terms_and_descriptions",
        "estado": "secundaria"},
    "terzaghi_qu": {
        "titulo": "Consistencia de arcillas según qu (Terzaghi & Peck, 1967): muy blanda <24, blanda 24-48, media 48-96, firme 96-192, muy firme 192-383, dura >383 kPa",
        "url": "https://yadda.icm.edu.pl/baztech/element/bwmeta1.element.baztech-565a5ca6-0332-4e34-aef6-ab48bbbbb2d5/c/Widodo.pdf",
        "nota": "Tabla reproducida en Widodo (Pontianak soft clay); el límite superior de la clase \"dura\" es el de Terzaghi & Peck.",
        "estado": "secundaria"},
    "is50_ucs": {
        "titulo": "Relación UCS ≈ 20-25 × Is(50) (ISRM; Broch & Franklin 1972 usan 24); es muy variable: 15-50 según IS 8764, 8-50 en rocas anisótropas",
        "url": "https://www.geomechanics.io/free-tool/free-point-load-strength-uniaxial-compressive-strength-plt-ucs-calculator",
        "nota": "También: https://infralens.in/code/IS-8764-1998 y https://presentations.copernicus.org/EGU2018/EGU2018-6054_presentation.pdf",
        "estado": "secundaria"},
    "gs_suelos": {
        "titulo": "Gravedad específica típica de sólidos: arena 2.63-2.68, limo 2.65-2.70, arcilla inorgánica 2.67-2.90, suelos orgánicos 1.2-2.0",
        "url": "https://ocw.ehu.eus/pluginfile.php/39477/mod_resource/content/7/OCW%20Lesson2%20Theory%20v2.pdf",
        "nota": "Contrastado con apuntes de la Universidad de Patras: https://eclass.upatras.gr/modules/document/file.php/GEO345/ENGINEERING%20GEOLOGY%20%28LABORATORIES%202025-2026%29/Lab1_2_3_Moisture%20content-Density-SpGravity.pdf",
        "estado": "secundaria"},
    "peso_unitario": {
        "titulo": "Pesos unitarios naturales: grava 19-22, arena 17-20, limo 17-20, arcilla 15-22 kN/m³; roca: granito 2.6-2.7, caliza 2.3-2.7, arenisca 2.2-2.8, lutita 2.4-2.8 g/cm³",
        "url": "https://ocw.ehu.eus/pluginfile.php/39477/mod_resource/content/7/OCW%20Lesson2%20Theory%20v2.pdf",
        "nota": "Roca: https://finesoftware.eu/help/geo5/en/unit-weight-of-rocks-01/ y https://literature.wolkersdorfer.info/literature/Density%20of%20Rocks.pdf",
        "estado": "secundaria"},
    "phi_suelos": {
        "titulo": "Ángulo de fricción de arenas y gravas: arenas 26-45° (suelta 28-30, media 30-36, densa 36-41), gravas 36-42°, SM 32-35°, SC 30-40°, cohesión efectiva ≈ 0",
        "url": "https://geotechdata.info/parameter/angle-of-friction",
        "nota": "También: https://www.gf.uns.ac.rs/~zbornik/doc/NS2018.34.pdf. Para arcillas/limos no se halló un rango fiable de φ' en la búsqueda: "
                "solo se alerta por encima de 45° (valor extremo para un suelo fino).",
        "estado": "secundaria"},
    "organica_loi": {
        "titulo": "Materia orgánica por pérdida por ignición: 3-20 % ligeramente orgánico, 20-75 % orgánico, >75 % turba; DOTs limitan 2-7 % en subrasantes",
        "url": "https://eprints.um.edu.my/8870",
        "nota": "Límites de DOT: https://docs.lib.purdue.edu/jtrp/1186 (el método de ignición tiende a sobrestimar en suelos poco orgánicos). Método: ASTM D2974.",
        "estado": "secundaria"},
    "limite_contraccion": {
        "titulo": "Límite de contracción: por definición menor que el límite plástico; arcillas británicas 12-25 %, arcillas tropicales/bentoníticas 30-50 %",
        "url": "https://nora.nerc.ac.uk/508951/1/geotechnique2_FINAL.pdf",
        "nota": "Definición SL < PL: https://en.wikipedia.org/wiki/Atterberg_limits",
        "estado": "secundaria"},
    "indice_liquidez": {
        "titulo": "Índice de liquidez IL = (w − LP)/IP: IL > 1 significa w > LL (suelo \"líquido\" en sitio), solo posible en arcillas muy sensibles/\"quick\"",
        "url": "https://www.emerald.com/books/book/12901/chapter/83274853/The-Sensitivity-of-Clays",
        "nota": "Houston & Mitchell (1969): la sensibilidad crece con el índice de liquidez. IL > 1.6 en arcilla sensible de Tiller-Flotten.",
        "estado": "secundaria"},
    "astm_c33_c88": {
        "titulo": "ASTM C33 / C88: pérdida por sulfatos en agregado fino ≤ 10 % (Na₂SO₄) / 15 % (MgSO₄); grueso ≤ 12 % / 18 %",
        "url": "https://store.astm.org/c0088_c0088m-24.html",
        "nota": "La búsqueda web confirmó C88 (método) y C33 (especificación) pero los porcentajes provienen de resúmenes secundarios "
                "(https://www.hufocw.org/Download/file/32334); verificar contra el texto vigente de la norma o la especificación INVIAS del proyecto.",
        "estado": "secundaria"},
    "inv_terrones": {
        "titulo": "Terrones de arcilla y partículas deleznables: criterio ≤ 2 % (INV Art. 320-13), el mismo que ya usa la plantilla GDA-FLC-034",
        "url": "https://store.astm.org/c0033_c0033m-18.html",
        "nota": "ASTM C142/C33 fija 3.0 % máx. en agregado fino. Es un criterio de ESPECIFICACIÓN (cumple / no cumple), no un rango típico.",
        "estado": "secundaria"},
    "fisica": {
        "titulo": "Relaciones físicas básicas de suelos (sin fuente externa)",
        "url": "",
        "nota": "Curva de saturación (cero vacíos de aire): γd = Gs·γw / (1 + w·Gs); grado de saturación Sr ≤ 100 %; relación de vacíos e > 0; "
                "densidad húmeda < Gs·γw; humedad ≥ 0.",
        "estado": "criterio"},
}

NIVEL_ETIQUETA = {"alto": "Resultado improbable", "revisar": "Fuera del rango habitual"}

# Texto corto para mostrar el material en la interfaz.
MATERIAL_NOMBRE = {"roca": "roca", "arcilla": "arcilla", "limo": "limo", "arena": "arena", "grava": "grava", "organico": "suelo orgánico"}
FINOS = ("arcilla", "limo")
GRANULARES = ("arena", "grava")

# ── Tablas de umbrales ─────────────────────────────────────────────
# Humedad natural (%): (revisar si >, alto si >). Fuentes: agua_suelos / porosidad_roca / fisica (saturación).
#   arena: saturada suelta ≈ e/Gs ≈ 1.0/2.65 ≈ 38 % -> "alto" por encima de 40.   grava: e ≈ 0.9 -> ≈ 34 % -> 35.
HUMEDAD_MAX = {
    "roca": (5.0, 15.0, "porosidad_roca"),
    "arcilla": (80.0, 150.0, "agua_suelos"),
    "limo": (80.0, 150.0, "agua_suelos"),
    "arena": (30.0, 40.0, "agua_suelos"),
    "grava": (20.0, 35.0, "agua_suelos"),
    "organico": (500.0, 800.0, "agua_suelos"),
}
# Humedad natural MÍNIMA (%): (revisar si <, improbable si <). Un suelo natural no puede tener menos agua que la higroscópica de uno
# secado al aire (fuente agua_minima); 0.01 % en una arcilla es casi seguro un error de pesada.
HUMEDAD_MIN = {
    "arcilla": (6.0, 2.0, "agua_minima"),
    "limo": (5.0, 1.5, "agua_minima"),
    "arena": (0.1, None, "agua_minima"),
    "grava": (0.1, None, "agua_minima"),
    "roca": (0.01, None, "agua_minima"),
}
# Gravedad específica Gs: (mín típico, máx típico) para "revisar"; fuera de (GS_ABSOLUTO) -> "alto".
GS_TIPICO = {"arena": (2.55, 2.80), "grava": (2.55, 2.85), "limo": (2.55, 2.90), "arcilla": (2.55, 2.95),
             "organico": (1.2, 2.55), "roca": (2.2, 3.1)}
GS_ABSOLUTO = (1.0, 3.6)
# Densidad natural (húmeda, g/cm³): (mín, máx) típicos para "revisar"; "alto" fuera de DENS_ABSOLUTA.
DENS_HUMEDA_TIPICA = {"roca": (2.0, 3.0), "arcilla": (1.4, 2.25), "limo": (1.5, 2.25), "arena": (1.5, 2.3),
                      "grava": (1.7, 2.4), "organico": (0.9, 1.9)}
DENS_HUMEDA_ABSOLUTA = {"roca": (1.2, 3.6)}
DENS_SUELO_ABSOLUTA = (0.8, 2.9)
# CBR (%) por clase USCS: (mín, máx). Fuente cbr_uscs.
CBR_USCS = {"GW": (40, 80), "GP": (30, 60), "GM": (20, 60), "GC": (20, 40), "SW": (20, 40), "SP": (10, 40), "SM": (10, 40),
            "SC": (5, 20), "ML": (0, 15), "CL": (0, 15), "OL": (0, 5), "MH": (0, 10), "CH": (0, 15), "OH": (0, 5)}
CBR_POR_MATERIAL = {"grava": (20, 80), "arena": (10, 40), "limo": (0, 15), "arcilla": (0, 15), "organico": (0, 5)}
# Resistencia a la compresión de ROCA (MPa): fuera de (0.25, 250) "revisar" (límites de R0 y R6 de la ISRM).
UCS_ROCA = (0.25, 250.0)
UCS_ROCA_ABSOLUTO = (0.05, 450.0)
IS50_ROCA = (0.03, 15.0)
IS50_ROCA_ABSOLUTO = (0.005, 25.0)
UCS_IS50_RAZON = (8.0, 50.0)
# Ángulo de fricción (°): granulares "revisar" fuera de (25, 48); cualquier suelo "alto" por encima de 55.
PHI_GRANULAR = (25.0, 48.0)
PHI_MAXIMO_FINO = 45.0
PHI_ABSOLUTO = 55.0
COHESION_GRANULAR_KPA = 10.0
# Sulfatos (ASTM C33): (límite Na₂SO₄, límite MgSO₄) en %.
SULFATOS_FINO = (10.0, 15.0)
SULFATOS_GRUESO = (12.0, 18.0)
TERRONES_MAX = 2.0   # INV Art. 320-13, igual que la plantilla GDA-FLC-034
# Materia orgánica (%): suelo inorgánico "revisar" por encima de 7 (límite alto de los DOTs), "alto" por encima de 20.
MO_INORGANICO = (7.0, 20.0)
MO_ORGANICO_MIN = 3.0


def _a(nivel, mensaje, *fuentes):
    return {"nivel": nivel, "mensaje": mensaje, "fuentes": [f for f in fuentes if f]}


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def _f(v, d=1):
    s = f"{v:.{d}f}"
    if "." in s:   # solo se quitan ceros de los decimales: sin esto 20 se escribía «2» y 100 «1»
        s = s.rstrip("0").rstrip(".")
    return s if s not in ("", "-") else "0"


def _nombre(material):
    return MATERIAL_NOMBRE.get(material, "este material")


# ════════════════════════════════════════════════════════════════════
# Reglas por magnitud (se reutilizan entre ensayos: la humedad aparece en 8 formularios distintos)
# ════════════════════════════════════════════════════════════════════
def alertas_humedad(w, material, ctx=None, etiqueta="Humedad"):
    out = []
    if not _num(w):
        return out
    if w < 0:
        return [_a("alto", f"{etiqueta} negativa ({_f(w, 2)} %): la masa seca es mayor que la húmeda. Revisa las pesadas y los recipientes.", "fisica")]
    piso = HUMEDAD_MIN.get(material)
    if piso:
        rev_min, alto_min, fuente_min = piso
        if alto_min is not None and w < alto_min:
            out.append(_a("alto", f"{etiqueta} = {_f(w, 2)} % es demasiado baja para {_nombre(material)}: un suelo natural no baja de ≈{_f(alto_min, 0)} % "
                                  "(eso es un suelo secado al horno). Revisa que la masa seca y la húmeda no sean casi iguales o que no se haya "
                                  "digitado la misma pesada.", fuente_min))
        elif w < rev_min:
            out.append(_a("revisar", f"{etiqueta} = {_f(w, 2)} % es baja para {_nombre(material)} (lo habitual es más de ≈{_f(rev_min, 1)} %).", fuente_min))
    elif w == 0:
        out.append(_a("revisar", f"{etiqueta} = 0 %: no hay pérdida de masa al secar. Revisa las pesadas.", "agua_minima"))
    lc = (ctx or {}).get("lc")
    if material in FINOS and _num(lc) and 0 < w < lc and not out:
        out.append(_a("revisar", f"{etiqueta} ({_f(w, 1)} %) está por debajo del límite de contracción ({_f(lc, 1)} %): el suelo estaría más seco que su estado de máxima contracción.",
                      "limite_contraccion"))
    tope = HUMEDAD_MAX.get(material)
    if tope:
        rev, alto, fuente = tope
        if w > alto:
            out.append(_a("alto", f"{etiqueta} = {_f(w, 1)} % es demasiado alta para {_nombre(material)} "
                                  f"(por encima de la humedad de saturación habitual, ≈{_f(alto, 0)} %).", fuente))
        elif w > rev:
            extra = (" Si es una roca muy porosa o meteorizada (arenisca, lutita) puede ser real." if material == "roca" else "")
            out.append(_a("revisar", f"{etiqueta} = {_f(w, 1)} % es alta para {_nombre(material)} "
                                     f"(lo habitual es hasta ≈{_f(rev, 0)} %).{extra}", fuente))
    ctx = ctx or {}
    ll, lp, ip = ctx.get("ll"), ctx.get("lp"), ctx.get("ip")
    if material in FINOS and _num(ll) and _num(lp) and _num(ip) and ip > 0 and w > ll:
        il = (w - lp) / ip
        out.append(_a("revisar", f"{etiqueta} ({_f(w, 1)} %) supera el límite líquido ({_f(ll, 0)} %): índice de liquidez {_f(il, 2)} > 1. "
                                 "Solo ocurre en arcillas muy sensibles; si no se espera eso, revisa humedad y límites.", "indice_liquidez"))
    return out


def alertas_gs(gs, material, etiqueta="Gravedad específica"):
    if not _num(gs):
        return []
    lo_a, hi_a = GS_ABSOLUTO
    if gs <= lo_a or gs > hi_a:
        return [_a("alto", f"{etiqueta} = {_f(gs, 3)} es físicamente improbable (valores reales de sólidos: ≈1.2–3.4). Revisa las pesadas del picnómetro y la temperatura.", "gs_suelos")]
    rango = GS_TIPICO.get(material, (2.4, 3.0))
    if gs < rango[0] or gs > rango[1]:
        return [_a("revisar", f"{etiqueta} = {_f(gs, 3)} fuera del rango habitual para {_nombre(material)} "
                              f"({_f(rango[0], 2)}–{_f(rango[1], 2)}).", "gs_suelos")]
    return []


def alertas_densidad_humeda(rho, material, etiqueta="Densidad húmeda"):
    if not _num(rho):
        return []
    lo_a, hi_a = DENS_HUMEDA_ABSOLUTA.get(material, DENS_SUELO_ABSOLUTA)
    if rho <= lo_a or rho > hi_a:
        return [_a("alto", f"{etiqueta} = {_f(rho, 3)} g/cm³ es improbable para {_nombre(material)}. Revisa masa, altura y diámetro del espécimen.", "peso_unitario")]
    lo, hi = DENS_HUMEDA_TIPICA.get(material, (1.4, 2.4))
    if rho < lo or rho > hi:
        return [_a("revisar", f"{etiqueta} = {_f(rho, 3)} g/cm³ fuera del rango habitual para {_nombre(material)} ({_f(lo, 2)}–{_f(hi, 2)} g/cm³).", "peso_unitario")]
    return []


def _zav(gs, w):
    """Densidad seca (g/cm³) con cero vacíos de aire para humedad w (%) y gravedad específica gs."""
    return gs / (1 + w / 100 * gs)


# ════════════════════════════════════════════════════════════════════
# Reglas por ensayo
# ════════════════════════════════════════════════════════════════════
def _humedad(v, material, ctx):
    return alertas_humedad(v.get("w"), material, ctx)


def _limites(v, material, ctx):
    out = []
    ll, lp, ip = v.get("ll"), v.get("lp"), v.get("ip")
    if not (_num(ll) and _num(lp)):
        return out
    if lp > ll or (_num(ip) and ip < 0):
        out.append(_a("alto", f"El límite plástico ({_f(lp, 0)}) es mayor que el líquido ({_f(ll, 0)}): índice de plasticidad negativo. Revisa los ensayos de LL y LP.", "linea_u"))
        return out
    if ll > 16 and _num(ip) and ip > 0.9 * (ll - 8) + 1:
        out.append(_a("alto", f"LL = {_f(ll, 0)}, IP = {_f(ip, 0)} cae por encima de la línea U de Casagrande (IP máx. ≈ {_f(0.9 * (ll - 8), 0)}): "
                              "ningún suelo natural suele ubicarse ahí. Verifica el método del LL y el rodillo del LP.", "linea_u"))
    if material == "arcilla" and _num(ip) and 0 <= ip < 4:
        out.append(_a("revisar", f"La muestra se trata como arcilla pero IP = {_f(ip, 0)} (<4): esa plasticidad es la de un limo o un suelo no plástico "
                                 "(en la carta de Casagrande una arcilla tiene IP > 7). Revisa el límite plástico o la descripción visual.", "linea_u"))
    if material == "roca":
        out.append(_a("revisar", "Esta muestra se trata como roca, pero tiene Límites de Atterberg (aplican a suelos finos). Verifica el tipo de muestra.", "linea_u"))
    return out


def _granulometria(v, material, ctx):
    out = []
    finos, grava, arena, cu = v.get("finos"), v.get("grava"), v.get("arena"), v.get("cu")
    if _num(v.get("retenido_pct")) and v["retenido_pct"] > 100.3:
        out.append(_a("alto", f"Los retenidos suman {_f(v['retenido_pct'], 1)} % de la masa inicial (más del 100 %): hay masas mal digitadas o la masa inicial está mal.", "fisica"))
    if _num(finos) and finos < -0.3:
        out.append(_a("alto", f"El % que pasa el No. 200 sale negativo ({_f(finos, 1)} %). Revisa la masa inicial seca y los retenidos.", "fisica"))
    if _num(cu) and cu < 1:
        out.append(_a("alto", f"Cu = {_f(cu, 2)} < 1: D60 no puede ser menor que D10. Revisa los retenidos por tamiz.", "fisica"))
    desc = (ctx or {}).get("desc_tipo")
    if desc and _num(finos):
        if desc in FINOS and finos < 30:
            out.append(_a("revisar", f"La descripción visual dice {desc.lower()} pero la granulometría solo tiene {_f(finos, 0)} % de finos "
                                     "(un suelo fino tiene más de 50 %). Revisa la descripción o los tamices.", "linea_u"))
        elif desc == "arena" and finos > 50:
            out.append(_a("revisar", f"La descripción visual dice arena pero la granulometría tiene {_f(finos, 0)} % de finos (>50 %: suelo fino).", "linea_u"))
        elif desc == "grava" and _num(grava) and grava < 30:
            out.append(_a("revisar", f"La descripción visual dice grava pero solo hay {_f(grava, 0)} % de grava en la granulometría.", "linea_u"))
    if material == "roca":
        out.append(_a("revisar", "Esta muestra se trata como roca, pero tiene Granulometría (aplica a suelos/agregados). Verifica el tipo de muestra.", "fisica"))
    return out


def _proctor(v, material, ctx):
    out = []
    opt_w, max_d = v.get("opt_w"), v.get("max_d")
    puntos = v.get("puntos") or []
    if _num(opt_w) and _num(max_d):
        lo, hi = 1.3, 2.4
        if max_d <= 0.9 or max_d > 2.9:
            out.append(_a("alto", f"Densidad máxima = {_f(max_d, 3)} g/cm³ es improbable para cualquier suelo compactado. Revisa volumen del molde y masas.", "proctor_rangos"))
        elif max_d < lo or max_d > hi:
            out.append(_a("revisar", f"Densidad máxima = {_f(max_d, 3)} g/cm³ fuera del rango extremo de la literatura ({_f(lo, 1)}–{_f(hi, 1)} g/cm³).", "proctor_rangos"))
        if opt_w < 0 or opt_w > 60:
            out.append(_a("alto", f"Humedad óptima = {_f(opt_w, 1)} % es improbable. Revisa las humedades de cada prueba.", "proctor_rangos"))
        elif opt_w < 5 or opt_w > 40:
            out.append(_a("revisar", f"Humedad óptima = {_f(opt_w, 1)} % fuera del rango extremo de la literatura (5–40 %).", "proctor_rangos"))
        if material in GRANULARES and (max_d < 1.6 or opt_w > 20):
            out.append(_a("revisar", f"Para {_nombre(material)} se espera densidad ≈1.8–2.2 g/cm³ y humedad óptima ≈10–12 % "
                                     f"(resultado: {_f(max_d, 2)} g/cm³ y {_f(opt_w, 1)} %).", "proctor_rangos"))
        elif material in FINOS and (max_d > 2.05 or opt_w < 8):
            out.append(_a("revisar", f"Para {_nombre(material)} se espera densidad ≈1.4–1.85 g/cm³ y humedad óptima ≈15–20 % "
                                     f"(resultado: {_f(max_d, 2)} g/cm³ y {_f(opt_w, 1)} %).", "proctor_rangos"))
        ws = [w for w, _d in puntos]
        if ws and (opt_w < min(ws) or opt_w > max(ws)):
            out.append(_a("revisar", f"La humedad óptima ({_f(opt_w, 1)} %) queda fuera de las humedades ensayadas ({_f(min(ws), 1)}–{_f(max(ws), 1)} %): "
                                     "la curva se extrapola. Agrega pruebas para que el óptimo quede encerrado.", "proctor_rangos"))
    # Cada punto debe quedar por debajo de la curva de saturación (cero vacíos de aire).
    for i, (w, ds) in enumerate(puntos, 1):
        if ds > _zav(3.0, w):
            out.append(_a("alto", f"La prueba {i} (w = {_f(w, 1)} %, γd = {_f(ds, 3)} g/cm³) queda por encima de la curva de saturación incluso con Gs = 3.0: imposible. Revisa masa/volumen/humedad.", "fisica"))
            break
        if ds > _zav(2.75, w):
            out.append(_a("revisar", f"La prueba {i} (w = {_f(w, 1)} %, γd = {_f(ds, 3)} g/cm³) queda por encima de la curva de saturación para Gs = 2.75: "
                                     "solo sería posible con sólidos muy densos. Revisa masa/volumen/humedad.", "fisica"))
            break
    return out


def _cbr(v, material, ctx):
    out = []
    cbr = v.get("cbr")
    simbolo = (ctx or {}).get("uscs")
    if _num(cbr):
        rango, ref = None, None
        if simbolo:
            clave = simbolo.split("-")[0]
            rango, ref = CBR_USCS.get(clave), clave
        if rango is None and material in CBR_POR_MATERIAL:
            rango, ref = CBR_POR_MATERIAL[material], _nombre(material)
        if cbr < 0:
            out.append(_a("alto", f"CBR negativo ({_f(cbr, 1)} %). Revisa las lecturas de fuerza.", "fisica"))
        elif rango:
            lo, hi = rango
            if cbr > 2 * hi and hi < 100:
                out.append(_a("alto", f"CBR = {_f(cbr, 1)} % es más del doble del máximo típico para {ref} ({_f(hi, 0)} %). Revisa fuerzas y lecturas de penetración.", "cbr_uscs"))
            elif cbr > hi or cbr < lo:
                rango_txt = f"{_f(lo, 0)}–{_f(hi, 0)} %" if lo else f"hasta {_f(hi, 0)} %"
                out.append(_a("revisar", f"CBR = {_f(cbr, 1)} % fuera del rango típico para {ref} ({rango_txt}).", "cbr_uscs"))
        if cbr > 200:
            out.append(_a("alto", f"CBR = {_f(cbr, 0)} % supera con mucho el de la piedra triturada patrón (100 %). Revisa las lecturas.", "cbr_uscs"))
    exp = v.get("exp_pct")
    if _num(exp) and exp < 0:
        out.append(_a("revisar", f"Expansión negativa ({_f(exp, 2)} %): el dial final es menor que el inicial (contracción). Verifica las lecturas del extensómetro.", "fisica"))
    out += alertas_humedad(v.get("w_antes"), material, ctx, "Humedad antes de inmersión")
    wa, wd = v.get("w_antes"), v.get("w_despues")
    if _num(wd):
        out += alertas_humedad(wd, material, ctx, "Humedad después de inmersión")
        if _num(wa) and wd < wa - 0.5:
            out.append(_a("revisar", f"La humedad después de la inmersión ({_f(wd, 1)} %) es menor que la de antes ({_f(wa, 1)} %): la muestra sumergida "
                                     "debería ganar agua, no perderla. Revisa las masas de humedad.", "fisica"))
    return out


def _corte_directo(v, material, ctx):
    out = []
    phi, c_kpa, modo = v.get("phi"), v.get("c_kpa"), (v.get("modo") or "")
    if _num(phi):
        if phi > PHI_ABSOLUTO:
            out.append(_a("alto", f"Ángulo de fricción = {_f(phi, 1)}° es improbable (>55°). Revisa esfuerzos normales y cortantes.", "phi_suelos"))
        elif material in GRANULARES and (phi < PHI_GRANULAR[0] or phi > PHI_GRANULAR[1]) and "UU" not in modo:
            out.append(_a("revisar", f"φ = {_f(phi, 1)}° fuera del rango habitual para {_nombre(material)} ({_f(PHI_GRANULAR[0], 0)}–{_f(PHI_GRANULAR[1], 0)}°).", "phi_suelos"))
        elif material in FINOS and phi > PHI_MAXIMO_FINO:
            out.append(_a("revisar", f"φ = {_f(phi, 1)}° es muy alto para {_nombre(material)} (>45°).", "phi_suelos"))
    if _num(c_kpa) and material in GRANULARES and c_kpa > COHESION_GRANULAR_KPA and "UU" not in modo:
        out.append(_a("revisar", f"Cohesión = {_f(c_kpa, 1)} kPa en {_nombre(material)}: lo normal es ≈ 0 (puede ser cohesión aparente por humedad o cementación).", "phi_suelos"))
    out += alertas_gs(v.get("gs"), material)
    for p in v.get("probetas") or []:
        etq = f"Probeta {p.get('i')}"
        out += alertas_humedad(p.get("w_i"), material, ctx, f"{etq}: humedad inicial")
        sr, e0 = p.get("sr"), p.get("e0")
        if _num(sr) and sr > 105:
            out.append(_a("alto", f"{etq}: grado de saturación = {_f(sr, 0)} % (>100 %). Revisa masas, dimensiones del anillo y Gs.", "fisica"))
        if _num(e0) and e0 <= 0:
            out.append(_a("alto", f"{etq}: relación de vacíos ≤ 0. Revisa masas, dimensiones del anillo y Gs.", "fisica"))
        elif _num(e0) and material in GRANULARES and e0 > 1.2:
            out.append(_a("revisar", f"{etq}: relación de vacíos = {_f(e0, 2)} es alta para {_nombre(material)} (lo habitual es 0.4–1.0).", "fisica"))
        out += alertas_densidad_humeda(p.get("dh"), material, f"{etq}: densidad húmeda")
    return _sin_repetidos(out)


def _masa_unitaria(v, material, ctx):
    out = alertas_densidad_humeda(v.get("dens_h"), material)
    out += alertas_humedad(v.get("w"), material, ctx)
    ds, dh = v.get("dens_s"), v.get("dens_h")
    if _num(ds) and _num(dh) and ds > dh + 1e-9:
        out.append(_a("alto", "La densidad seca sale mayor que la húmeda (humedad negativa). Revisa la humedad usada.", "fisica"))
    return out


def _gravedad(v, material, ctx):
    out = alertas_gs(v.get("gs"), material)
    ab = v.get("absorcion")
    if _num(ab) and ab < 0:
        out.append(_a("alto", f"Absorción negativa ({_f(ab, 2)} %): la masa SSS es menor que la seca. Revisa las pesadas.", "fisica"))
    return out


def _materia_organica(v, material, ctx):
    mo = v.get("mo")
    if not _num(mo):
        return []
    if mo < 0 or mo > 100:
        return [_a("alto", f"Materia orgánica = {_f(mo, 1)} % es imposible (fuera de 0–100 %). Revisa las masas del crisol.", "fisica")]
    out = []
    if material in FINOS + GRANULARES:
        if mo > MO_INORGANICO[1]:
            out.append(_a("alto", f"Materia orgánica = {_f(mo, 1)} % corresponde a un suelo orgánico (>20 %), pero la muestra se trata como {_nombre(material)}. "
                                  "Revisa el tipo de suelo o las masas.", "organica_loi"))
        elif mo > MO_INORGANICO[0]:
            out.append(_a("revisar", f"Materia orgánica = {_f(mo, 1)} % es alta para {_nombre(material)} (los DOT limitan 2–7 %; >3 % ya es ligeramente orgánico).", "organica_loi"))
    elif material == "organico" and mo < MO_ORGANICO_MIN:
        out.append(_a("revisar", f"La muestra está descrita como orgánica pero la materia orgánica es solo {_f(mo, 1)} % (<3 %).", "organica_loi"))
    return out


def _limite_contraccion(v, material, ctx):
    out = []
    lc, lp = v.get("lc"), (ctx or {}).get("lp")
    if _num(lc):
        if lc < 0 or lc > 100:
            out.append(_a("alto", f"Límite de contracción = {_f(lc, 1)} % es improbable. Revisa volúmenes y masas.", "limite_contraccion"))
        elif _num(lp) and lc > lp:
            out.append(_a("alto", f"El límite de contracción ({_f(lc, 1)} %) es mayor que el límite plástico ({_f(lp, 0)} %): por definición debe ser menor.", "limite_contraccion"))
        elif lc > 50:
            out.append(_a("revisar", f"Límite de contracción = {_f(lc, 1)} % está por encima de lo típico (12–25 % en arcillas, hasta 30–50 % en tropicales/bentoníticas).", "limite_contraccion"))
    r = v.get("r")
    if _num(r) and (r < 0.5 or r > 2.9):
        out.append(_a("alto", f"Relación de contracción R = {_f(r, 2)} es improbable (R es una densidad relativa, típicamente 1.2–2.0).", "limite_contraccion"))
    return out


def _consolidacion(v, material, ctx):
    out = alertas_gs(v.get("gs"), material)
    out += alertas_humedad(v.get("w_i"), material, ctx, "Humedad inicial")
    out += alertas_densidad_humeda(v.get("dens_h"), material, "Peso unitario total inicial")
    sr, e0 = v.get("sr"), v.get("e0")
    if _num(sr) and sr > 105:
        out.append(_a("alto", f"Saturación inicial = {_f(sr, 0)} % (>100 %). Revisa masas, dimensiones del anillo y Gs.", "fisica"))
    if _num(e0) and e0 <= 0:
        out.append(_a("alto", "Relación de vacíos inicial ≤ 0. Revisa masas, dimensiones del anillo y Gs.", "fisica"))
    elif _num(e0) and e0 > 4 and material != "organico":
        out.append(_a("revisar", f"Relación de vacíos inicial = {_f(e0, 2)} es muy alta para un suelo inorgánico.", "fisica"))
    return out


def _compresion_inconfinada(v, material, ctx):
    out = []
    qu = v.get("qu_kpa")
    if _num(qu):
        if qu <= 0:
            out.append(_a("alto", "qu ≤ 0: revisa las lecturas de carga y las dimensiones del espécimen.", "terzaghi_qu"))
        elif qu > 1000:
            out.append(_a("revisar", f"qu = {_f(qu / 1000, 2)} MPa ya está en el rango de roca extremadamente débil (R0 ≥ 0.25–1 MPa, ISRM) y es muy alto para un suelo: "
                                     "verifica que el material y el ensayo sean los correctos (¿es roca? ¿carga en kN vs kgf?).", "isrm_ucs", "terzaghi_qu"))
        elif material in GRANULARES and qu > 100:
            out.append(_a("revisar", f"qu = {_f(qu, 0)} kPa en {_nombre(material)}: un suelo granular casi no tiene resistencia inconfinada (solo por cementación o humedad aparente).", "terzaghi_qu"))
        elif qu > 383 and material in FINOS:
            out.append(_a("revisar", f"qu = {_f(qu, 0)} kPa corresponde a arcilla dura (>383 kPa, Terzaghi & Peck): poco común; verifica la carga máxima y el área corregida.", "terzaghi_qu"))
    out += alertas_humedad(v.get("w"), material, ctx)
    out += alertas_densidad_humeda(v.get("dens_h"), material)
    return out


def _compresion_roca(v, material, ctx):
    out = []
    ucs = v.get("ucs_mpa")
    if _num(ucs):
        lo_a, hi_a = UCS_ROCA_ABSOLUTO
        lo, hi = UCS_ROCA
        if ucs <= lo_a or ucs > hi_a:
            out.append(_a("alto", f"Resistencia = {_f(ucs, 2)} MPa es improbable para roca (rango real ≈ 0.25–400 MPa). Revisa unidades de la carga (kN) y el diámetro del núcleo.", "isrm_ucs"))
        elif ucs < lo:
            out.append(_a("revisar", f"Resistencia = {_f(ucs, 2)} MPa está por debajo de la roca extremadamente débil (R0 ≥ 0.25 MPa): se comporta como suelo.", "isrm_ucs"))
        elif ucs > hi:
            out.append(_a("revisar", f"Resistencia = {_f(ucs, 0)} MPa supera 250 MPa (R6, extremadamente fuerte): rara; verifica la carga máxima.", "isrm_ucs"))
        is50 = (ctx or {}).get("is50")
        if _num(is50) and is50 > 0:
            ratio = ucs / is50
            if ratio < UCS_IS50_RAZON[0] or ratio > UCS_IS50_RAZON[1]:
                out.append(_a("revisar", f"UCS/Is(50) = {_f(ratio, 1)} (UCS {_f(ucs, 1)} MPa, Is50 {_f(is50, 2)} MPa de carga puntual) se sale del rango observado 8–50 "
                                         "(típico 20–25). Los dos ensayos de esta muestra no concuerdan.", "is50_ucs"))
    out += alertas_humedad(v.get("w"), "roca", ctx)
    out += alertas_densidad_humeda(v.get("dens_h"), "roca")
    return out


def _carga_puntual(v, material, ctx):
    out = []
    is50 = v.get("is50_prom")
    if _num(is50):
        lo_a, hi_a = IS50_ROCA_ABSOLUTO
        lo, hi = IS50_ROCA
        if is50 <= lo_a or is50 > hi_a:
            out.append(_a("alto", f"Is(50) = {_f(is50, 3)} MPa es improbable (rango real ≈ 0.01–15 MPa). Revisa la carga (kN) y el diámetro equivalente.", "isrm_ucs"))
        elif is50 < lo or is50 > hi:
            out.append(_a("revisar", f"Is(50) = {_f(is50, 3)} MPa fuera del rango habitual de roca ({_f(lo, 2)}–{_f(hi, 0)} MPa; R6 empieza en 10 MPa).", "isrm_ucs"))
        ucs = (ctx or {}).get("ucs_mpa")
        if _num(ucs) and is50 > 0:
            ratio = ucs / is50
            if ratio < UCS_IS50_RAZON[0] or ratio > UCS_IS50_RAZON[1]:
                out.append(_a("revisar", f"UCS/Is(50) = {_f(ratio, 1)} se sale del rango observado 8–50 (típico 20–25): la carga puntual y la compresión de esta muestra no concuerdan.", "is50_ucs"))
    out += alertas_humedad(v.get("w"), "roca", ctx)
    return out


def _solidez(v, material, ctx):
    out = []
    for etq, clave, (na, mg) in (("agregado grueso", "perdida_grueso", SULFATOS_GRUESO), ("agregado fino", "perdida_fino", SULFATOS_FINO)):
        p = v.get(clave)
        if not _num(p):
            continue
        if p < 0 or p > 100:
            out.append(_a("alto", f"Pérdida en {etq} = {_f(p, 1)} % es imposible. Revisa las masas antes/después de los ciclos.", "fisica"))
        elif p > mg:
            out.append(_a("revisar", f"Pérdida en {etq} = {_f(p, 1)} % supera el límite de ASTM C33 con cualquiera de las dos sales ({_f(na, 0)} % Na₂SO₄ / {_f(mg, 0)} % MgSO₄): no cumple.", "astm_c33_c88"))
        elif p > na:
            out.append(_a("revisar", f"Pérdida en {etq} = {_f(p, 1)} % supera el límite de ASTM C33 para sulfato de sodio ({_f(na, 0)} %); "
                                     f"si el ensayo fue con sulfato de magnesio el límite es {_f(mg, 0)} %.", "astm_c33_c88"))
    return out


def _terrones(v, material, ctx):
    p = v.get("ponderado")
    if not _num(p):
        return []
    if p < 0 or p > 100:
        return [_a("alto", f"% de terrones ponderado = {_f(p, 2)} % es imposible. Revisa las masas inicial/final.", "fisica")]
    if p > TERRONES_MAX:
        return [_a("revisar", f"% de terrones y partículas deleznables = {_f(p, 2)} % supera el criterio de {_f(TERRONES_MAX, 0)} % (INV Art. 320-13): no cumple.", "inv_terrones")]
    return []


def _alargamiento_aplanamiento(v, material, ctx):
    """Chequeos físicos (no hay un rango "típico" confiable sin fijarse en la especificación del proyecto): la masa de partículas alargadas /
    aplanadas no puede pasar de la masa de su fracción, y el índice del ensayo está entre 0 y 100 %."""
    out = []
    nombres = {"alar": "alargadas", "apla": "aplanadas"}
    for clave, t in (v.get("tablas") or {}).items():
        for i, (ini, par) in enumerate(t.get("filas") or [], start=1):
            if _num(ini) and _num(par) and par > ini:
                out.append(_a("alto", f"Fracción {i}: la masa de partículas {nombres.get(clave, '')} ({_f(par, 1)} g) es mayor que la masa de la fracción "
                                      f"({_f(ini, 1)} g). Revisa las pesadas.", "fisica"))
        idx = t.get("indice")
        if _num(idx) and not (0 <= idx <= 100):
            out.append(_a("alto", f"Índice de {'alargamiento' if clave == 'alar' else 'aplanamiento'} = {_f(idx, 1)} % fuera de 0–100 %.", "fisica"))
    return out


_REGLAS = {
    "alargamiento-aplanamiento": _alargamiento_aplanamiento,
    "humedad": _humedad, "limites": _limites, "granulometria": _granulometria, "pasa200": _granulometria,
    "proctor": _proctor, "cbr": _cbr, "corte-directo": _corte_directo, "masa-unitaria": _masa_unitaria,
    "gravedad-especifica": _gravedad, "materia-organica": _materia_organica, "limite-contraccion": _limite_contraccion,
    "consolidacion": _consolidacion, "compresion-inconfinada": _compresion_inconfinada, "compresion-roca": _compresion_roca,
    "carga-puntual": _carga_puntual, "solidez-sulfatos": _solidez, "terrones-arcilla": _terrones,
}


def _sin_repetidos(alertas):
    vistos, out = set(), []
    for a in alertas:
        if a["mensaje"] not in vistos:
            vistos.add(a["mensaje"])
            out.append(a)
    return out


def evaluar(tipo, valores, material=None, ctx=None):
    """Alertas del ensayo `tipo` (clave interna: 'humedad', 'proctor'...). `valores`: dict con las magnitudes ya calculadas
    del ensayo; `material`: 'roca'|'arcilla'|'limo'|'arena'|'grava'|'organico'|None; `ctx`: valores de otros ensayos de la
    misma muestra (ll, lp, ip, is50, ucs_mpa, uscs, desc_tipo). Ordena primero las de nivel 'alto'."""
    regla = _REGLAS.get(tipo)
    if regla is None or not valores:
        return []
    try:
        res = _sin_repetidos(regla(valores, material, ctx or {}))
    except Exception:   # una regla defectuosa nunca debe tumbar la pantalla del ensayo
        return []
    return sorted(res, key=lambda a: 0 if a["nivel"] == "alto" else 1)


def fuentes_de(alertas):
    """FUENTES usadas por las alertas dadas (sin repetir, en orden de aparición): lista de (id, dict)."""
    ids = []
    for a in alertas:
        for f in a["fuentes"]:
            if f not in ids and f in FUENTES:
                ids.append(f)
    return [(i, FUENTES[i]) for i in ids]
