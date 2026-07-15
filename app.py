"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Replica del diseño hecho en Figma, construida en Streamlit + Python.

Cómo correrla en tu computador:
    streamlit run app.py
"""

import streamlit as st
from datetime import datetime
import uuid
import pandas as pd

# ════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE LA PÁGINA
# ════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Geodelta Lab",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "v1.0.0"

# ════════════════════════════════════════════════════════════════════
# ESTILOS (paleta de colores igual a la del diseño en Figma)
# ════════════════════════════════════════════════════════════════════
NAVY = "#1B2E4B"
NAVY_HOVER = "#243D62"
BLUE = "#2563EB"
BLUE_LIGHT = "#EFF4FF"
SUCCESS = "#16A34A"
SUCCESS_LIGHT = "#DCFCE7"
WARNING = "#D97706"
WARNING_LIGHT = "#FEF3C7"
DANGER = "#DC2626"
SURFACE = "#FFFFFF"
BG = "#F2F4F7"
BORDER = "#E2E6ED"
TEXT = "#1A2332"
MUTED = "#6B7A8D"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"]  {{
        font-family: 'Inter', 'Segoe UI', sans-serif;
    }}
    .stApp {{
        background-color: {BG};
    }}
    section[data-testid="stSidebar"] {{
        background-color: {NAVY};
    }}
    section[data-testid="stSidebar"] * {{
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] .stButton button {{
        background-color: transparent;
        border: 1px solid rgba(255,255,255,0.15);
        color: #FFFFFF !important;
        text-align: left;
        width: 100%;
        border-radius: 8px;
        margin-bottom: 6px;
    }}
    section[data-testid="stSidebar"] .stButton button:hover {{
        background-color: {NAVY_HOVER};
        border-color: {NAVY_HOVER};
    }}
    .card {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 12px;
        padding: 20px 22px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .section-title {{
        font-size: 12px;
        font-weight: 700;
        color: {MUTED};
        text-transform: uppercase;
        letter-spacing: 0.06em;
        border-bottom: 1px solid {BORDER};
        padding-bottom: 8px;
        margin-bottom: 14px;
        margin-top: 4px;
    }}
    .badge {{
        display: inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
    }}
    .badge-success {{ background: {SUCCESS_LIGHT}; color: {SUCCESS}; }}
    .badge-warning {{ background: {WARNING_LIGHT}; color: {WARNING}; }}
    .badge-muted {{ background: #EEF1F5; color: {MUTED}; }}
    .result-box {{
        background: {BLUE_LIGHT};
        border: 1px solid #BFDBFE;
        border-radius: 8px;
        padding: 14px 18px;
        margin-top: 10px;
    }}
    .result-value {{ font-size: 26px; font-weight: 800; color: #1D4ED8; }}
    .result-label {{ font-size: 12px; font-weight: 600; color: #1D4ED8; text-transform: uppercase; letter-spacing: 0.05em; }}
    div.stButton > button[kind="primary"] {{
        background-color: {SUCCESS};
        border: none;
    }}
    h1, h2, h3 {{ color: {TEXT}; letter-spacing: -0.02em; }}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════
# DATOS DE REFERENCIA (tamices, normas, tipos de ensayo)
# ════════════════════════════════════════════════════════════════════
SIEVES = [
    ("s_3p", '3"', "76.20"), ("s_2p", '2"', "50.80"), ("s_1p5", '1½"', "38.10"),
    ("s_1p", '1"', "25.40"), ("s_34", '¾"', "19.05"), ("s_12", '½"', "12.70"),
    ("s_38", '⅜"', "9.525"), ("s_4", "No. 4", "4.750"), ("s_10", "No. 10", "2.000"),
    ("s_20", "No. 20", "0.850"), ("s_40", "No. 40", "0.425"), ("s_60", "No. 60", "0.250"),
    ("s_100", "No. 100", "0.150"), ("s_200", "No. 200", "0.075"),
]

ASSAY_LABELS = {
    "granulometria": "Granulometría",
    "humedad": "Contenido de humedad",
    "masa-unitaria": "Masa unitaria parafinada",
}

NORMAS = {
    "granulometria": ["INV E-213", "INV E-214", "ASTM D422", "ASTM D7928"],
    "humedad": ["INV E-122", "ASTM D2216"],
    "masa-unitaria": ["INV E-202", "ASTM D1188"],
}

STATUS_LABELS = {"sin-iniciar": "Sin iniciar", "en-proceso": "En proceso", "finalizado": "Finalizado"}
STATUS_BADGE = {"sin-iniciar": "badge-muted", "en-proceso": "badge-warning", "finalizado": "badge-success"}


# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL (datos de ejemplo, igual a los del Figma)
# ════════════════════════════════════════════════════════════════════
def init_state():
    if "initialized" in st.session_state:
        return
    st.session_state.initialized = True
    st.session_state.screen = "home"
    st.session_state.projects = [
        {"code": "GEO-2024-001", "name": "Estudio de suelos vía Bogotá-Medellín Km 14"},
        {"code": "GEO-2024-002", "name": "Cimentación edificio Torre Colina"},
    ]
    st.session_state.headers = {
        "GEO-2024-001::M1": {
            "project": "GEO-2024-001", "sample": "M1", "date": "2024-11-15",
            "apique": "AP-01", "location": "Sector Norte, Km 14+200",
            "depthStart": "0.00", "depthEnd": "1.50",
            "visualDescription": "Suelo arcilloso de color café oscuro, consistencia media.",
            "tempStart": "22", "tempEnd": "23", "humStart": "65", "humEnd": "63",
        },
    }
    st.session_state.assays = [
        {"id": "a001", "project": "GEO-2024-001", "sample": "M1", "type": "granulometria",
         "status": "finalizado", "data": {}, "observations": "Muestra en buen estado.",
         "laboratorist": "Ing. Carlos Mendoza", "signature": "",
         "lastModified": "2024-11-16T09:45:00", "createdAt": "2024-11-15T14:30:00"},
        {"id": "a002", "project": "GEO-2024-001", "sample": "M1", "type": "humedad",
         "status": "en-proceso", "data": {}, "observations": "",
         "laboratorist": "Ing. Carlos Mendoza", "signature": "",
         "lastModified": "2024-11-16T10:20:00", "createdAt": "2024-11-16T10:00:00"},
    ]
    st.session_state.selected_project_code = ""
    st.session_state.selected_sample = ""
    st.session_state.selected_assay_type = None
    st.session_state.selected_assay_id = None
    st.session_state.form_draft = {}


init_state()


def navigate(screen):
    st.session_state.screen = screen
    st.rerun()


def to_float(v, default=None):
    try:
        return float(str(v).replace(",", "."))
    except (ValueError, TypeError):
        return default


def now_iso():
    return datetime.now().isoformat()


# ════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════
def render_sidebar():
    with st.sidebar:
        st.markdown("## 🧪 GEODELTA LAB")
        st.markdown("---")
        if st.button("🏠  Inicio", use_container_width=True):
            navigate("home")
        if st.button("📂  Continuar ensayo", use_container_width=True):
            navigate("continue")
        if st.button("🔎  Buscar ensayos", use_container_width=True):
            navigate("search")
        st.markdown("---")
        if st.session_state.selected_project_code:
            st.caption("Proyecto activo")
            st.markdown(f"**{st.session_state.selected_project_code}**")
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.caption(f"Geodelta Lab · {APP_VERSION}")


# ════════════════════════════════════════════════════════════════════
# PANTALLA 1 · INICIO (lista de proyectos)
# ════════════════════════════════════════════════════════════════════
def render_home():
    st.title("GEODELTA LAB")
    st.caption("Selecciona un proyecto o inicia una acción")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🧪  Nuevo proyecto", use_container_width=True):
            navigate("new-project")
    with c2:
        if st.button("📂  Continuar ensayo", use_container_width=True):
            navigate("continue")
    with c3:
        if st.button("🔎  Buscar ensayos", use_container_width=True):
            navigate("search")

    st.markdown("### Proyectos")
    if not st.session_state.projects:
        st.info("Todavía no hay proyectos. Crea el primero con el botón **Nuevo proyecto**.")
    for p in st.session_state.projects:
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            cols = st.columns([4, 1])
            with cols[0]:
                st.markdown(f"**{p['code']}**")
                st.caption(p["name"])
            with cols[1]:
                if st.button("Abrir", key=f"open_{p['code']}", use_container_width=True):
                    st.session_state.selected_project_code = p["code"]
                    navigate("project-detail")
            st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PANTALLA · NUEVO PROYECTO
# ════════════════════════════════════════════════════════════════════
def render_new_project():
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Nuevo proyecto")

    existing_codes = [p["code"] for p in st.session_state.projects]
    code = st.text_input("Código del proyecto", placeholder="GEO-2024-001")
    name = st.text_input("Nombre del proyecto", placeholder="Estudio de suelos vía Bogotá-Medellín")

    if code:
        if code in existing_codes:
            st.error("Ese código de proyecto ya existe.")
        else:
            st.success("Código disponible.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            navigate("home")
    with col2:
        disabled = not code or not name or code in existing_codes
        if st.button("Continuar →", type="primary", use_container_width=True, disabled=disabled):
            st.session_state.projects.append({"code": code, "name": name})
            st.session_state.selected_project_code = code
            navigate("project-detail")


# ════════════════════════════════════════════════════════════════════
# PANTALLA · DETALLE DE PROYECTO (lista de muestras)
# ════════════════════════════════════════════════════════════════════
def render_project_detail():
    code = st.session_state.selected_project_code
    project = next((p for p in st.session_state.projects if p["code"] == code), None)
    if not project:
        navigate("home")
        return

    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown(f"## {project['code']}")
    st.caption(project["name"])

    samples = sorted({k.split("::")[1] for k in st.session_state.headers if k.startswith(code + "::")})

    if st.button("➕ Nueva muestra", type="primary"):
        navigate("new-sample")

    st.markdown("### Muestras")
    if not samples:
        st.info("Este proyecto todavía no tiene muestras.")
    for s in samples:
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            cols = st.columns([4, 1])
            with cols[0]:
                st.markdown(f"**Muestra {s}**")
                header = st.session_state.headers.get(f"{code}::{s}", {})
                st.caption(f"Apique {header.get('apique', '—')} · Profundidad {header.get('depthStart','—')} - {header.get('depthEnd','—')} m")
            with cols[1]:
                if st.button("Abrir", key=f"open_sample_{s}", use_container_width=True):
                    st.session_state.selected_sample = s
                    navigate("sample-summary")
            st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PANTALLA · NUEVA MUESTRA
# ════════════════════════════════════════════════════════════════════
def render_new_sample():
    code = st.session_state.selected_project_code
    st.button("← Atrás", on_click=lambda: navigate("project-detail"))
    st.markdown(f"## Nueva muestra · {code}")

    existing = {k.split("::")[1] for k in st.session_state.headers if k.startswith(code + "::")}
    sample = st.text_input("Nombre de la muestra", placeholder="M1, M3, M12…")
    sample = sample.strip().upper()

    if sample and sample in existing:
        st.warning(f"La muestra **{sample}** ya existe en este proyecto. Se abrirá su información.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            navigate("project-detail")
    with col2:
        if st.button("Continuar →", type="primary", use_container_width=True, disabled=not sample):
            st.session_state.selected_sample = sample
            if sample in existing:
                navigate("sample-summary")
            else:
                navigate("header-form")


# ════════════════════════════════════════════════════════════════════
# PANTALLA · ENCABEZADO (solo para muestra nueva)
# ════════════════════════════════════════════════════════════════════
def render_header_form():
    code = st.session_state.selected_project_code
    sample = st.session_state.selected_sample
    st.button("← Atrás", on_click=lambda: navigate("new-sample"))
    st.markdown(f"## Encabezado · {code} · Muestra {sample}")
    st.info(f"Muestra nueva: **{code} · {sample}** — Complete el encabezado para continuar.")

    st.markdown('<div class="section-title">Información general</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        date = st.text_input("Fecha de ejecución", placeholder="2024-11-15")
        apique = st.text_input("Apique / Sondeo", placeholder="AP-01")
        depth_start = st.text_input("Profundidad inicial (m)", placeholder="0.00")
        temp_start = st.text_input("Temperatura inicial (°C)", placeholder="22")
        hum_start = st.text_input("Humedad inicial (%)", placeholder="65")
    with c2:
        location = st.text_input("Localización", placeholder="Sector Norte, Km 14+200")
        depth_end = st.text_input("Profundidad final (m)", placeholder="1.50")
        temp_end = st.text_input("Temperatura final (°C)", placeholder="23")
        hum_end = st.text_input("Humedad final (%)", placeholder="63")

    visual_desc = st.text_area("Descripción visual", placeholder="Suelo arcilloso de color café oscuro, de consistencia media…", height=100)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            navigate("new-sample")
    with col2:
        if st.button("Guardar encabezado →", type="primary", use_container_width=True):
            key = f"{code}::{sample}"
            st.session_state.headers[key] = {
                "project": code, "sample": sample, "date": date, "apique": apique,
                "location": location, "depthStart": depth_start, "depthEnd": depth_end,
                "visualDescription": visual_desc, "tempStart": temp_start, "tempEnd": temp_end,
                "humStart": hum_start, "humEnd": hum_end,
            }
            navigate("sample-summary")


# ════════════════════════════════════════════════════════════════════
# PANTALLA · RESUMEN DE LA MUESTRA
# ════════════════════════════════════════════════════════════════════
def render_sample_summary():
    code = st.session_state.selected_project_code
    sample = st.session_state.selected_sample
    header = st.session_state.headers.get(f"{code}::{sample}")

    st.button("← Atrás", on_click=lambda: navigate("project-detail"))
    st.markdown(f"## Muestra {sample}")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    cols = st.columns(5)
    cols[0].metric("Proyecto", code)
    cols[1].metric("Muestra", sample)
    if header:
        cols[2].metric("Profundidad", f"{header['depthStart']}–{header['depthEnd']} m")
        cols[3].metric("Apique", header["apique"])
    cols[4].metric("Estado", "Activa")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("➕ Nuevo ensayo", type="primary"):
        navigate("select-assay-type")

    st.markdown("### Ensayos")
    sample_assays = [a for a in st.session_state.assays if a["project"] == code and a["sample"] == sample]
    if not sample_assays:
        st.info("Esta muestra todavía no tiene ensayos. Usa **Nuevo ensayo** para agregar uno.")
    for a in sample_assays:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 1.4, 1])
        with cols[0]:
            st.markdown(f"**{ASSAY_LABELS[a['type']]}**")
        with cols[1]:
            badge_class = STATUS_BADGE[a["status"]]
            st.markdown(f'<span class="badge {badge_class}">{STATUS_LABELS[a["status"]]}</span>', unsafe_allow_html=True)
        with cols[2]:
            if st.button("Abrir ensayo", key=f"open_assay_{a['id']}", use_container_width=True):
                st.session_state.selected_assay_id = a["id"]
                st.session_state.selected_assay_type = a["type"]
                navigate("assay-form")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PANTALLA · SELECCIONAR TIPO DE ENSAYO
# ════════════════════════════════════════════════════════════════════
def render_select_assay_type():
    st.button("← Atrás", on_click=lambda: navigate("sample-summary"))
    st.markdown("## Selecciona el tipo de ensayo")

    cols = st.columns(3)
    icons = {"granulometria": "📊", "humedad": "💧", "masa-unitaria": "⚖️"}
    for col, (t, label) in zip(cols, ASSAY_LABELS.items()):
        with col:
            st.markdown('<div class="card" style="text-align:center; height:150px;">', unsafe_allow_html=True)
            st.markdown(f"<div style='font-size:34px;'>{icons[t]}</div>", unsafe_allow_html=True)
            st.markdown(f"**{label}**")
            if st.button("Seleccionar", key=f"select_{t}", use_container_width=True):
                code = st.session_state.selected_project_code
                sample = st.session_state.selected_sample
                existing = next((a for a in st.session_state.assays
                                  if a["project"] == code and a["sample"] == sample and a["type"] == t), None)
                if existing:
                    st.session_state.selected_assay_id = existing["id"]
                else:
                    new_id = f"a-{uuid.uuid4().hex[:8]}"
                    st.session_state.assays.append({
                        "id": new_id, "project": code, "sample": sample, "type": t,
                        "status": "sin-iniciar", "data": {}, "observations": "",
                        "laboratorist": "", "signature": "",
                        "lastModified": now_iso(), "createdAt": now_iso(),
                    })
                    st.session_state.selected_assay_id = new_id
                st.session_state.selected_assay_type = t
                navigate("assay-form")
            st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PANTALLA · FORMULARIO DEL ENSAYO
# ════════════════════════════════════════════════════════════════════
def field(label, key, data, placeholder="", numeric=True):
    val = st.text_input(label, value=data.get(key, ""), placeholder=placeholder, key=f"f_{key}")
    data[key] = val
    return to_float(val) if numeric else val


def render_norma_selector(assay_type, data, key_prefix):
    st.markdown('<div class="section-title">Norma aplicada</div>', unsafe_allow_html=True)
    options = NORMAS[assay_type]
    current = data.get(f"{key_prefix}_norma", "")
    idx = options.index(current) + 1 if current in options else 0
    choice = st.radio("Norma", ["(ninguna)"] + options, index=idx, horizontal=True,
                       key=f"norma_{key_prefix}", label_visibility="collapsed")
    data[f"{key_prefix}_norma"] = "" if choice == "(ninguna)" else choice


def render_equipo(data, prefix):
    st.markdown('<div class="section-title">Equipo utilizado</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, col in enumerate(cols, start=1):
        with col:
            data[f"{prefix}_equipo{i}_nombre"] = st.text_input(f"Equipo {i}", value=data.get(f"{prefix}_equipo{i}_nombre", ""),
                                                                 placeholder=f"Nombre del equipo {i}", key=f"{prefix}_eq{i}n")
            data[f"{prefix}_equipo{i}_serie"] = st.text_input(f"N° serie {i}", value=data.get(f"{prefix}_equipo{i}_serie", ""),
                                                                placeholder="N° de serie / código", key=f"{prefix}_eq{i}s",
                                                                label_visibility="collapsed")


def render_granulometria_form(data):
    st.markdown('<div class="section-title">Pasa No. 200 por lavado</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        data["pasa200_recipiente"] = st.text_input("Recipiente No.", value=data.get("pasa200_recipiente", ""), placeholder="R-01")
        antes = field("Muestra seca antes del lavado (g)", "pasa200_antes", data, "350.5")
    with c2:
        despues = field("Muestra seca después del lavado (g)", "pasa200_despues", data, "315.2")
        field("Masa suelo seco + recipiente (g)", "pasa200_suelo_rec", data, "0.0")
    with c3:
        field("Masa suelo seco + recipiente (74h) (g)", "pasa200_suelo_rec_74", data, "0.0")
        field("Masa del recipiente (g)", "pasa200_masa_recipiente", data, "0.0")

    if antes and despues and antes > 0:
        pasa200 = ((antes - despues) / antes) * 100
        st.markdown(f'<div class="result-box"><span class="result-label">% Pasa No. 200 por lavado</span><br>'
                    f'<span class="result-value">{pasa200:.2f}%</span></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Granulometría por tamizado</div>', unsafe_allow_html=True)
    peso_total = field("Peso total seco de la muestra (g)", "gran_peso_total", data, "350.5")

    rows = []
    acum = 0.0
    for key, label, apert in SIEVES:
        retenido = field(f"{label} — Retenido (g)", key, data, "0.0")
        retenido = retenido or 0.0
        acum += retenido
        pct_ret = (retenido / peso_total * 100) if peso_total else None
        pct_acum = (acum / peso_total * 100) if peso_total else None
        pct_pasa = (100 - pct_acum) if pct_acum is not None else None
        rows.append({
            "Tamiz": label, "Apert. (mm)": apert, "Retenido (g)": retenido,
            "Ret. acum. (g)": round(acum, 2),
            "% Retenido": f"{pct_ret:.2f}%" if pct_ret is not None else "—",
            "% Ret. acum.": f"{pct_acum:.2f}%" if pct_acum is not None else "—",
            "% Pasa": f"{pct_pasa:.2f}%" if pct_pasa is not None else "—",
        })

    finos = field("Finos < 0.075 mm (g)", "s_finos", data, "0.0") or 0.0
    total_acum = acum + finos
    pct_finos = (finos / peso_total * 100) if peso_total else None
    pct_acum_total = (total_acum / peso_total * 100) if peso_total else None
    rows.append({
        "Tamiz": "Finos", "Apert. (mm)": "< 0.075", "Retenido (g)": finos,
        "Ret. acum. (g)": round(total_acum, 2),
        "% Retenido": f"{pct_finos:.2f}%" if pct_finos is not None else "—",
        "% Ret. acum.": f"{pct_acum_total:.2f}%" if pct_acum_total is not None else "—",
        "% Pasa": f"{100 - pct_acum_total:.2f}%" if pct_acum_total is not None else "—",
    })

    st.markdown("**Tabla de resultados**")
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    render_equipo(data, "gran")
    render_norma_selector("granulometria", data, "gran")


def render_humedad_form(data):
    st.markdown('<div class="section-title">Método de ensayo</div>', unsafe_allow_html=True)
    metodo = data.get("hum_metodo", "")
    idx = ["(ninguno)", "Método A", "Método B"].index(metodo) if metodo in ["Método A", "Método B"] else 0
    choice = st.radio("Método", ["(ninguno)", "Método A", "Método B"], index=idx, horizontal=True,
                       key="hum_metodo_radio", label_visibility="collapsed")
    data["hum_metodo"] = "" if choice == "(ninguno)" else choice

    st.markdown('<div class="section-title">Determinación de humedad natural</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        data["hum_recipiente"] = st.text_input("Recipiente No.", value=data.get("hum_recipiente", ""), placeholder="R-01")
    with c2:
        tara = field("Masa del recipiente (g)", "hum_tara", data, "25.30")
    with c3:
        wh = field("M. suelo húmedo + recipiente (g)", "hum_suelo_humedo_tara", data, "148.60")
    with c4:
        ws = field("M. suelo seco + recipiente 74h (g)", "hum_suelo_seco_tara", data, "132.40")

    c1, c2 = st.columns(2)
    masa_seca = (ws - tara) if (ws is not None and tara is not None) else None
    masa_agua = (wh - ws) if (wh is not None and ws is not None) else None
    with c1:
        st.metric("Masa suelo seco (g)", f"{masa_seca:.2f}" if masa_seca is not None else "—")
    with c2:
        st.metric("Masa agua evaporada (g)", f"{masa_agua:.2f}" if masa_agua is not None else "—")

    if tara is not None and wh is not None and ws is not None and ws > tara:
        calc = ((wh - ws) / (ws - tara)) * 100
        st.markdown(f'<div class="result-box"><span class="result-label">Contenido de humedad calculado</span><br>'
                    f'<span class="result-value">{calc:.2f}%</span></div>', unsafe_allow_html=True)

    render_equipo(data, "hum")
    render_norma_selector("humedad", data, "hum")


def render_masa_unitaria_form(data):
    st.markdown('<div class="section-title">Determinación de masa unitaria parafinada</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        ps = field("Masa en el aire (g)", "mu_peso_aire", data, "245.80")
        pa = field("Masa en el agua parafinado (g)", "mu_peso_agua_par", data, "138.20")
        pp = field("Masa de la parafina (g)", "mu_peso_parafina", data, "12.50")
    with c2:
        field("Masa en el aire parafinado (g)", "mu_peso_aire_par", data, "258.30")
        field("Temperatura del agua (°C)", "mu_temp_agua", data, "22.0")
        dp = field("Densidad de la parafina (g/cm³)", "mu_dens_parafina", data, "0.90") or 0.9

    if ps is not None and pa is not None and pp is not None and dp:
        vol = (ps - pa) - pp / dp
        if vol and vol > 0:
            gamma_b = (ps - pp) / vol
            c1, c2 = st.columns(2)
            with c1:
                st.metric("Volumen de la muestra (cm³)", f"{vol:.3f}")
            with c2:
                st.markdown(f'<div class="result-box"><span class="result-label">Peso unitario bruto (g/cm³)</span><br>'
                            f'<span class="result-value">{gamma_b:.3f}</span></div>', unsafe_allow_html=True)

    render_equipo(data, "mu")
    render_norma_selector("masa-unitaria", data, "mu")


def render_assay_form():
    code = st.session_state.selected_project_code
    sample = st.session_state.selected_sample
    assay_id = st.session_state.selected_assay_id
    assay = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    header = st.session_state.headers.get(f"{code}::{sample}")

    if not assay:
        navigate("sample-summary")
        return

    st.button("← Atrás", on_click=lambda: navigate("sample-summary"))

    # Barra fija superior
    bar_cols = st.columns(5)
    bar_cols[0].markdown(f"**Proyecto**<br>{code}", unsafe_allow_html=True)
    bar_cols[1].markdown(f"**Muestra**<br>{sample}", unsafe_allow_html=True)
    bar_cols[2].markdown(f"**Ensayo**<br>{ASSAY_LABELS[assay['type']]}", unsafe_allow_html=True)
    bar_cols[3].markdown(f"**Apique**<br>{header['apique'] if header else '—'}", unsafe_allow_html=True)
    badge_class = STATUS_BADGE[assay["status"]]
    bar_cols[4].markdown(f"**Estado**<br><span class='badge {badge_class}'>{STATUS_LABELS[assay['status']]}</span>", unsafe_allow_html=True)

    st.markdown(f"## {ASSAY_LABELS[assay['type']]}")

    data = dict(assay.get("data", {}))  # copia editable

    if assay["type"] == "granulometria":
        render_granulometria_form(data)
    elif assay["type"] == "humedad":
        render_humedad_form(data)
    elif assay["type"] == "masa-unitaria":
        render_masa_unitaria_form(data)

    st.markdown('<div class="section-title">Observaciones</div>', unsafe_allow_html=True)
    observations = st.text_area("Observaciones", value=assay.get("observations", ""),
                                 placeholder="Observaciones generales del ensayo…", label_visibility="collapsed")

    st.markdown('<div class="section-title">Laboratorista</div>', unsafe_allow_html=True)
    laboratorist = st.text_input("Laboratorista", value=assay.get("laboratorist", ""),
                                  placeholder="Nombre completo del laboratorista", label_visibility="collapsed")

    st.markdown('<div class="section-title">Firma digital</div>', unsafe_allow_html=True)
    try:
        from streamlit_drawable_canvas import st_canvas
        canvas_result = st_canvas(
            fill_color="rgba(255,255,255,0)", stroke_width=2, stroke_color="#1B2E4B",
            background_color="#FFFFFF", height=130, width=600, drawing_mode="freedraw",
            key=f"signature_{assay_id}",
        )
        signature = "firmado" if canvas_result.image_data is not None and canvas_result.json_data["objects"] else ""
    except ImportError:
        st.caption("Para habilitar la firma con mouse/pantalla táctil instala `streamlit-drawable-canvas` (ver README).")
        signature = st.text_input("Nombre de quien firma (temporal, mientras se activa la firma digital)",
                                   value=assay.get("signature", ""))

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾  Guardar borrador", use_container_width=True):
            assay["data"] = data
            assay["observations"] = observations
            assay["laboratorist"] = laboratorist
            assay["signature"] = signature
            assay["status"] = "en-proceso"
            assay["lastModified"] = now_iso()
            navigate("sample-summary")
    with col2:
        if st.button("✅  Finalizar ensayo", type="primary", use_container_width=True):
            assay["data"] = data
            assay["observations"] = observations
            assay["laboratorist"] = laboratorist
            assay["signature"] = signature
            assay["status"] = "finalizado"
            assay["lastModified"] = now_iso()
            navigate("confirmation")


# ════════════════════════════════════════════════════════════════════
# PANTALLA · CONTINUAR ENSAYO
# ════════════════════════════════════════════════════════════════════
def render_continue():
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Continuar ensayo")

    in_progress = [a for a in st.session_state.assays if a["status"] == "en-proceso"]
    if not in_progress:
        st.info("No hay ensayos en proceso por el momento.")
    for a in in_progress:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 2, 2, 1])
        cols[0].markdown(f"**{a['project']}**")
        cols[1].markdown(f"Muestra {a['sample']}")
        cols[2].markdown(ASSAY_LABELS[a["type"]])
        with cols[3]:
            if st.button("Continuar", key=f"cont_{a['id']}", use_container_width=True):
                st.session_state.selected_project_code = a["project"]
                st.session_state.selected_sample = a["sample"]
                st.session_state.selected_assay_type = a["type"]
                st.session_state.selected_assay_id = a["id"]
                navigate("assay-form")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PANTALLA · BUSCAR ENSAYOS
# ════════════════════════════════════════════════════════════════════
def render_search():
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Buscar ensayos")

    c1, c2, c3 = st.columns(3)
    with c1:
        f_project = st.text_input("Proyecto")
    with c2:
        f_sample = st.text_input("Muestra")
    with c3:
        f_type = st.selectbox("Tipo de ensayo", ["(todos)"] + list(ASSAY_LABELS.values()))

    results = st.session_state.assays
    if f_project:
        results = [a for a in results if f_project.lower() in a["project"].lower()]
    if f_sample:
        results = [a for a in results if f_sample.lower() in a["sample"].lower()]
    if f_type != "(todos)":
        results = [a for a in results if ASSAY_LABELS[a["type"]] == f_type]

    if results:
        df = pd.DataFrame([{
            "Proyecto": a["project"], "Muestra": a["sample"], "Ensayo": ASSAY_LABELS[a["type"]],
            "Estado": STATUS_LABELS[a["status"]], "Última modificación": a["lastModified"],
        } for a in results])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No se encontraron ensayos con esos filtros.")


# ════════════════════════════════════════════════════════════════════
# PANTALLA · CONFIRMACIÓN
# ════════════════════════════════════════════════════════════════════
def render_confirmation():
    assay_id = st.session_state.selected_assay_id
    assay = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    if not assay:
        navigate("home")
        return

    st.markdown("## ✅ Ensayo finalizado")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f"**Proyecto:** {assay['project']}")
    st.markdown(f"**Muestra:** {assay['sample']}")
    st.markdown(f"**Ensayo:** {ASSAY_LABELS[assay['type']]}")
    st.markdown("✔️ Encabezado completo &nbsp;&nbsp; ✔️ Datos completos &nbsp;&nbsp; ✔️ Firma agregada")
    st.markdown('</div>', unsafe_allow_html=True)

    st.info("La generación de Excel y PDF llegará en la próxima versión de la app. Por ahora los datos quedan guardados dentro de la aplicación.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Volver a la muestra", use_container_width=True):
            navigate("sample-summary")
    with col2:
        if st.button("🏠 Ir al inicio", type="primary", use_container_width=True):
            navigate("home")


# ════════════════════════════════════════════════════════════════════
# ENRUTADOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════
render_sidebar()

SCREENS = {
    "home": render_home,
    "new-project": render_new_project,
    "project-detail": render_project_detail,
    "new-sample": render_new_sample,
    "header-form": render_header_form,
    "sample-summary": render_sample_summary,
    "select-assay-type": render_select_assay_type,
    "assay-form": render_assay_form,
    "continue": render_continue,
    "search": render_search,
    "confirmation": render_confirmation,
}

SCREENS.get(st.session_state.screen, render_home)()
