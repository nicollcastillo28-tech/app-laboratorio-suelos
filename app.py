"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Roles: Jefe (control total, genera bitácora, descarga Excel) y Auxiliar (digita resultados).

Cómo correrla en tu computador:
    streamlit run app.py
"""

import os
import uuid
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE LA PÁGINA
# ════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Geodelta Lab", page_icon="🧪", layout="wide", initial_sidebar_state="expanded")

APP_VERSION = "v2.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_GRANULOMETRIA = os.path.join(BASE_DIR, "templates", "CLASIFICACION_DE_SUELOS.xlsm")

# Claves de acceso. CÁMBIALAS aquí por las que tú quieras usar en tu laboratorio.
PASSWORDS = {
    "jefe": "geodelta2024",
    "auxiliar": "aux2024",
}

# ════════════════════════════════════════════════════════════════════
# ESTILOS
# ════════════════════════════════════════════════════════════════════
NAVY, NAVY_HOVER = "#1B2E4B", "#243D62"
SUCCESS, SUCCESS_LIGHT = "#16A34A", "#DCFCE7"
WARNING, WARNING_LIGHT = "#D97706", "#FEF3C7"
SURFACE, BG, BORDER, TEXT, MUTED = "#FFFFFF", "#F2F4F7", "#E2E6ED", "#1A2332", "#6B7A8D"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', 'Segoe UI', sans-serif; }}
    .stApp {{ background-color: {BG}; }}
    section[data-testid="stSidebar"] {{ background-color: {NAVY}; }}
    section[data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
    section[data-testid="stSidebar"] .stButton button {{
        background-color: transparent; border: 1px solid rgba(255,255,255,0.15);
        color: #FFFFFF !important; text-align: left; width: 100%; border-radius: 8px; margin-bottom: 6px;
    }}
    section[data-testid="stSidebar"] .stButton button:hover {{ background-color: {NAVY_HOVER}; border-color: {NAVY_HOVER}; }}
    .card {{
        background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 12px;
        padding: 20px 22px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .section-title {{
        font-size: 12px; font-weight: 700; color: {MUTED}; text-transform: uppercase;
        letter-spacing: 0.06em; border-bottom: 1px solid {BORDER}; padding-bottom: 8px;
        margin-bottom: 14px; margin-top: 4px;
    }}
    .badge {{ display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }}
    .badge-success {{ background: {SUCCESS_LIGHT}; color: {SUCCESS}; }}
    .badge-warning {{ background: {WARNING_LIGHT}; color: {WARNING}; }}
    .badge-muted {{ background: #EEF1F5; color: {MUTED}; }}
    .role-pill {{
        display: inline-block; padding: 4px 12px; border-radius: 999px; font-size: 12px;
        font-weight: 700; background: rgba(255,255,255,0.12); margin-bottom: 12px;
    }}
    div.stButton > button[kind="primary"] {{ background-color: {SUCCESS}; border: none; }}
    h1, h2, h3 {{ color: {TEXT}; letter-spacing: -0.02em; }}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════
# TAMICES — MISMOS EXACTOS DE LA PLANTILLA DE EXCEL (hoja MUESTRA, col. E20:E34)
# ════════════════════════════════════════════════════════════════════
SIEVES = [
    ("s_3", '3"', "76.2", "E20"), ("s_2p5", '2 1/2"', "63.5", "E21"), ("s_2", '2"', "50.8", "E22"),
    ("s_1p5", '1 1/2"', "38.1", "E23"), ("s_1", '1"', "25.4", "E24"), ("s_34", '3/4"', "19.05", "E25"),
    ("s_12", '1/2"', "12.7", "E26"), ("s_38", '3/8"', "9.52", "E27"), ("s_4", "No. 4", "4.76", "E28"),
    ("s_10", "No. 10", "2.00", "E29"), ("s_20", "No. 20", "0.841", "E30"), ("s_40", "No. 40", "0.42", "E31"),
    ("s_60", "No. 60", "0.25", "E32"), ("s_100", "No. 100", "0.149", "E33"), ("s_200", "No. 200", "0.075", "E34"),
]

ASSAY_LABELS = {"granulometria": "Granulometría", "humedad": "Contenido de humedad", "masa-unitaria": "Masa unitaria parafinada"}
NORMAS = {
    "granulometria": ["INV E-213", "INV E-214", "ASTM D422", "ASTM D7928"],
    "humedad": ["INV E-122", "ASTM D2216"],
    "masa-unitaria": ["INV E-202", "ASTM D1188"],
}
STATUS_LABELS = {"sin-iniciar": "Sin iniciar", "en-proceso": "En proceso", "finalizado": "Finalizado"}
STATUS_BADGE = {"sin-iniciar": "badge-muted", "en-proceso": "badge-warning", "finalizado": "badge-success"}

BITACORA_ENSAYOS = [
    "Granulometría", "Pasa 200", "Humedad", "Límites de Atterberg", "Límite de contracción",
    "Materia orgánica", "Proctor", "CBR", "Compresión inconfinada", "Compresión en roca",
    "Peso unitario", "Gravedad específica", "Consolidación", "Corte CD", "Corte CU", "Corte UU", "Otro",
]

# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL
# ════════════════════════════════════════════════════════════════════
def init_state():
    if "initialized" in st.session_state:
        return
    st.session_state.initialized = True
    st.session_state.role = None
    st.session_state.screen = "home"
    st.session_state.projects = [
        {"code": "GEO-2024-001", "name": "Estudio de suelos vía Bogotá-Medellín Km 14"},
    ]
    st.session_state.headers = {
        "GEO-2024-001::M1": {
            "project": "GEO-2024-001", "sample": "M1", "date": "2024-11-15", "apique": "AP-01",
            "location": "Sector Norte, Km 14+200", "depthStart": "0.00", "depthEnd": "1.50",
            "visualDescription": "Suelo arcilloso de color café oscuro, consistencia media.",
        },
    }
    st.session_state.sample_status = {"GEO-2024-001::M1": "en-proceso"}
    st.session_state.assays = [
        {"id": "a001", "project": "GEO-2024-001", "sample": "M1", "type": "granulometria",
         "status": "en-proceso", "data": {}, "observations": "", "laboratorist": "",
         "lastModified": datetime.now().isoformat(), "createdAt": datetime.now().isoformat()},
    ]
    st.session_state.selected_project_code = ""
    st.session_state.selected_sample = ""
    st.session_state.selected_assay_type = None
    st.session_state.selected_assay_id = None
    st.session_state.bitacoras = {}  # code -> {"meta": {...}, "rows": [...]}


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


def require_role(*allowed):
    if st.session_state.role not in allowed:
        st.warning("🔒 No tienes permiso para ver esta sección.")
        if st.button("← Volver al inicio"):
            navigate("home")
        st.stop()


# ════════════════════════════════════════════════════════════════════
# LOGIN
# ════════════════════════════════════════════════════════════════════
def render_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col = st.columns([1, 1.4, 1])[1]
    with col:
        st.markdown("## 🧪 GEODELTA LAB")
        st.caption("Ingresa tu clave de acceso")
        role_choice = st.radio("Tipo de usuario", ["Auxiliar", "Jefe"], horizontal=True)
        password = st.text_input("Clave de acceso", type="password")
        if st.button("Ingresar", type="primary", use_container_width=True):
            role_key = "jefe" if role_choice == "Jefe" else "auxiliar"
            if password == PASSWORDS[role_key]:
                st.session_state.role = role_key
                navigate("home")
            else:
                st.error("Clave incorrecta.")


# ════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════
def render_sidebar():
    with st.sidebar:
        st.markdown("## 🧪 GEODELTA LAB")
        role_label = "Jefe de laboratorio" if st.session_state.role == "jefe" else "Auxiliar"
        st.markdown(f'<span class="role-pill">👤 {role_label}</span>', unsafe_allow_html=True)
        st.markdown("---")
        if st.button("🏠  Inicio", use_container_width=True):
            navigate("home")
        if st.session_state.role == "jefe":
            if st.button("📋  Generar bitácora", use_container_width=True):
                navigate("bitacora")
        if st.button("📂  Continuar ensayo", use_container_width=True):
            navigate("continue")
        if st.button("🔎  Buscar ensayos", use_container_width=True):
            navigate("search")
        st.markdown("---")
        if st.session_state.selected_project_code:
            st.caption("Proyecto activo")
            st.markdown(f"**{st.session_state.selected_project_code}**")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪  Cerrar sesión", use_container_width=True):
            st.session_state.role = None
            navigate("home")
        st.caption(f"Geodelta Lab · {APP_VERSION}")


# ════════════════════════════════════════════════════════════════════
# INICIO
# ════════════════════════════════════════════════════════════════════
def render_home():
    st.title("GEODELTA LAB")
    st.caption("Selecciona un proyecto o inicia una acción")

    if st.session_state.role == "jefe":
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
    else:
        st.info("Como auxiliar puedes abrir los proyectos existentes, digitar resultados de ensayos y actualizar el estado de las muestras.")

    st.markdown("### Proyectos")
    if not st.session_state.projects:
        st.info("Todavía no hay proyectos.")
    for p in st.session_state.projects:
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
# NUEVO PROYECTO (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_new_project():
    require_role("jefe")
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Nuevo proyecto")

    existing_codes = [p["code"] for p in st.session_state.projects]
    code = st.text_input("Código del proyecto", placeholder="GEO-2024-001")
    name = st.text_input("Nombre del proyecto", placeholder="Estudio de suelos vía Bogotá-Medellín")

    if code:
        st.error("Ese código ya existe.") if code in existing_codes else st.success("Código disponible.")

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
# DETALLE DE PROYECTO
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

    if st.session_state.role == "jefe":
        if st.button("➕ Nueva muestra", type="primary"):
            navigate("new-sample")

    st.markdown("### Muestras")
    if not samples:
        st.info("Este proyecto todavía no tiene muestras.")
    for s in samples:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 1.4, 1])
        with cols[0]:
            st.markdown(f"**Muestra {s}**")
            header = st.session_state.headers.get(f"{code}::{s}", {})
            st.caption(f"Apique {header.get('apique', '—')} · Profundidad {header.get('depthStart','—')}-{header.get('depthEnd','—')} m")
        with cols[1]:
            status = st.session_state.sample_status.get(f"{code}::{s}", "sin-iniciar")
            st.markdown(f'<span class="badge {STATUS_BADGE[status]}">{STATUS_LABELS[status]}</span>', unsafe_allow_html=True)
        with cols[2]:
            if st.button("Abrir", key=f"open_sample_{s}", use_container_width=True):
                st.session_state.selected_sample = s
                navigate("sample-summary")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# NUEVA MUESTRA (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_new_sample():
    require_role("jefe")
    code = st.session_state.selected_project_code
    st.button("← Atrás", on_click=lambda: navigate("project-detail"))
    st.markdown(f"## Nueva muestra · {code}")

    existing = {k.split("::")[1] for k in st.session_state.headers if k.startswith(code + "::")}
    sample = st.text_input("Nombre de la muestra", placeholder="M1, M3, M12…").strip().upper()

    if sample and sample in existing:
        st.warning(f"La muestra **{sample}** ya existe. Se abrirá su información.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            navigate("project-detail")
    with col2:
        if st.button("Continuar →", type="primary", use_container_width=True, disabled=not sample):
            st.session_state.selected_sample = sample
            navigate("sample-summary" if sample in existing else "header-form")


# ════════════════════════════════════════════════════════════════════
# ENCABEZADO (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_header_form():
    require_role("jefe")
    code = st.session_state.selected_project_code
    sample = st.session_state.selected_sample
    st.button("← Atrás", on_click=lambda: navigate("new-sample"))
    st.markdown(f"## Encabezado · {code} · Muestra {sample}")

    st.markdown('<div class="section-title">Información general</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        date = st.text_input("Fecha de ejecución", placeholder="2024-11-15")
        apique = st.text_input("Apique / Sondeo", placeholder="AP-01")
        depth_start = st.text_input("Profundidad inicial (m)", placeholder="0.00")
    with c2:
        location = st.text_input("Localización", placeholder="Sector Norte, Km 14+200")
        depth_end = st.text_input("Profundidad final (m)", placeholder="1.50")

    visual_desc = st.text_area("Descripción visual", placeholder="Suelo arcilloso de color café oscuro…", height=100)

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
                "visualDescription": visual_desc,
            }
            st.session_state.sample_status[key] = "sin-iniciar"
            navigate("sample-summary")


# ════════════════════════════════════════════════════════════════════
# RESUMEN DE LA MUESTRA
# ════════════════════════════════════════════════════════════════════
def render_sample_summary():
    code = st.session_state.selected_project_code
    sample = st.session_state.selected_sample
    key = f"{code}::{sample}"
    header = st.session_state.headers.get(key)

    st.button("← Atrás", on_click=lambda: navigate("project-detail"))
    st.markdown(f"## Muestra {sample}")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    cols = st.columns(4)
    cols[0].metric("Proyecto", code)
    cols[1].metric("Muestra", sample)
    if header:
        cols[2].metric("Profundidad", f"{header['depthStart']}–{header['depthEnd']} m")
        cols[3].metric("Apique", header["apique"])
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Estado de la muestra</div>', unsafe_allow_html=True)
    current_status = st.session_state.sample_status.get(key, "sin-iniciar")
    options = list(STATUS_LABELS.keys())
    new_status = st.selectbox("Estado", options, index=options.index(current_status),
                               format_func=lambda k: STATUS_LABELS[k])
    if new_status != current_status:
        st.session_state.sample_status[key] = new_status
        st.rerun()

    if st.session_state.role == "jefe":
        if st.button("➕ Nuevo ensayo", type="primary"):
            navigate("select-assay-type")

    st.markdown("### Ensayos")
    sample_assays = [a for a in st.session_state.assays if a["project"] == code and a["sample"] == sample]
    if not sample_assays:
        st.info("Esta muestra todavía no tiene ensayos.")
    for a in sample_assays:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 1.4, 1])
        cols[0].markdown(f"**{ASSAY_LABELS[a['type']]}**")
        cols[1].markdown(f'<span class="badge {STATUS_BADGE[a["status"]]}">{STATUS_LABELS[a["status"]]}</span>', unsafe_allow_html=True)
        with cols[2]:
            if st.button("Abrir ensayo", key=f"open_assay_{a['id']}", use_container_width=True):
                st.session_state.selected_assay_id = a["id"]
                st.session_state.selected_assay_type = a["type"]
                navigate("assay-form")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# SELECCIONAR TIPO DE ENSAYO (solo Jefe crea ensayos nuevos)
# ════════════════════════════════════════════════════════════════════
def render_select_assay_type():
    require_role("jefe")
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
                        "status": "sin-iniciar", "data": {}, "observations": "", "laboratorist": "",
                        "lastModified": now_iso(), "createdAt": now_iso(),
                    })
                    st.session_state.selected_assay_id = new_id
                st.session_state.selected_assay_type = t
                navigate("assay-form")
            st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE GRANULOMETRÍA (plantilla real del laboratorio)
# ════════════════════════════════════════════════════════════════════
def generar_excel_granulometria(code, sample, header, data):
    wb = load_workbook(TEMPLATE_GRANULOMETRIA, keep_vba=True)
    ws = wb["MUESTRA"]

    project = next((p for p in st.session_state.projects if p["code"] == code), None)
    ws["D7"] = project["name"] if project else code
    if header:
        ws["D9"] = header.get("location", "")
        ws["K7"] = header.get("date", "")
        ws["D12"] = header.get("apique", "")
        ws["K12"] = to_float(header.get("depthStart"))
        ws["M12"] = to_float(header.get("depthEnd"))
        ws["D13"] = header.get("visualDescription", "")
    ws["H12"] = sample
    ws["D17"] = to_float(data.get("masa_inicial_seca"))

    for key, _label, _apert, cell in SIEVES:
        ws[cell] = to_float(data.get(key)) or 0

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ════════════════════════════════════════════════════════════════════
# FORMULARIOS DE ENSAYO (solo captura de datos, sin cálculos)
# ════════════════════════════════════════════════════════════════════
def render_equipo(data, prefix):
    st.markdown('<div class="section-title">Equipo utilizado</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, col in enumerate(cols, start=1):
        with col:
            data[f"{prefix}_equipo{i}_nombre"] = st.text_input(f"Equipo {i}", value=data.get(f"{prefix}_equipo{i}_nombre", ""), placeholder=f"Nombre del equipo {i}", key=f"{prefix}_eq{i}n")
            data[f"{prefix}_equipo{i}_serie"] = st.text_input(f"N° serie {i}", value=data.get(f"{prefix}_equipo{i}_serie", ""), placeholder="N° de serie", key=f"{prefix}_eq{i}s", label_visibility="collapsed")


def render_norma_selector(assay_type, data, key_prefix):
    st.markdown('<div class="section-title">Norma aplicada</div>', unsafe_allow_html=True)
    options = NORMAS[assay_type]
    current = data.get(f"{key_prefix}_norma", "")
    idx = options.index(current) + 1 if current in options else 0
    choice = st.radio("Norma", ["(ninguna)"] + options, index=idx, horizontal=True, key=f"norma_{key_prefix}", label_visibility="collapsed")
    data[f"{key_prefix}_norma"] = "" if choice == "(ninguna)" else choice


def render_granulometria_form(data):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — los cálculos y la clasificación USCS los hace el Excel, no la app.")
    st.markdown('<div class="section-title">Datos generales</div>', unsafe_allow_html=True)
    data["masa_inicial_seca"] = st.text_input("Masa inicial seca (g)", value=data.get("masa_inicial_seca", ""), placeholder="350.5")

    st.markdown('<div class="section-title">Pesos retenidos por tamiz (g)</div>', unsafe_allow_html=True)
    rows = [{"Tamiz": label, "Abertura (mm)": apert, "Retenido (g)": to_float(data.get(key), 0.0)} for key, label, apert, _cell in SIEVES]
    df = pd.DataFrame(rows)
    edited = st.data_editor(
        df, hide_index=True, use_container_width=True, disabled=["Tamiz", "Abertura (mm)"],
        column_config={"Retenido (g)": st.column_config.NumberColumn(format="%.2f", step=0.1)},
        key="gran_sieve_editor",
    )
    for i, (key, _label, _apert, _cell) in enumerate(SIEVES):
        data[key] = edited.iloc[i]["Retenido (g)"]

    render_equipo(data, "gran")
    render_norma_selector("granulometria", data, "gran")


def render_humedad_form(data):
    st.info("Estos datos se guardan tal cual, sin calcular el % de humedad dentro de la app.")
    metodo = data.get("hum_metodo", "")
    idx = ["(ninguno)", "Método A", "Método B"].index(metodo) if metodo in ["Método A", "Método B"] else 0
    choice = st.radio("Método de ensayo", ["(ninguno)", "Método A", "Método B"], index=idx, horizontal=True, key="hum_metodo_radio")
    data["hum_metodo"] = "" if choice == "(ninguno)" else choice

    st.markdown('<div class="section-title">Determinación de humedad natural</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        data["hum_recipiente"] = st.text_input("Recipiente No.", value=data.get("hum_recipiente", ""), placeholder="R-01")
    with c2:
        data["hum_tara"] = st.text_input("Masa del recipiente (g)", value=data.get("hum_tara", ""), placeholder="25.30")
    with c3:
        data["hum_suelo_humedo_tara"] = st.text_input("M. suelo húmedo + recipiente (g)", value=data.get("hum_suelo_humedo_tara", ""), placeholder="148.60")
    with c4:
        data["hum_suelo_seco_tara"] = st.text_input("M. suelo seco + recipiente 74h (g)", value=data.get("hum_suelo_seco_tara", ""), placeholder="132.40")

    render_equipo(data, "hum")
    render_norma_selector("humedad", data, "hum")


def render_masa_unitaria_form(data):
    st.info("Estos datos se guardan tal cual, sin calcular el peso unitario dentro de la app.")
    c1, c2 = st.columns(2)
    with c1:
        data["mu_peso_aire"] = st.text_input("Masa en el aire (g)", value=data.get("mu_peso_aire", ""), placeholder="245.80")
        data["mu_peso_agua_par"] = st.text_input("Masa en el agua parafinado (g)", value=data.get("mu_peso_agua_par", ""), placeholder="138.20")
        data["mu_peso_parafina"] = st.text_input("Masa de la parafina (g)", value=data.get("mu_peso_parafina", ""), placeholder="12.50")
    with c2:
        data["mu_peso_aire_par"] = st.text_input("Masa en el aire parafinado (g)", value=data.get("mu_peso_aire_par", ""), placeholder="258.30")
        data["mu_temp_agua"] = st.text_input("Temperatura del agua (°C)", value=data.get("mu_temp_agua", ""), placeholder="22.0")
        data["mu_dens_parafina"] = st.text_input("Densidad de la parafina (g/cm³)", value=data.get("mu_dens_parafina", ""), placeholder="0.90")

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

    bar_cols = st.columns(5)
    bar_cols[0].markdown(f"**Proyecto**<br>{code}", unsafe_allow_html=True)
    bar_cols[1].markdown(f"**Muestra**<br>{sample}", unsafe_allow_html=True)
    bar_cols[2].markdown(f"**Ensayo**<br>{ASSAY_LABELS[assay['type']]}", unsafe_allow_html=True)
    bar_cols[3].markdown(f"**Apique**<br>{header['apique'] if header else '—'}", unsafe_allow_html=True)
    bar_cols[4].markdown(f"**Estado**<br><span class='badge {STATUS_BADGE[assay['status']]}'>{STATUS_LABELS[assay['status']]}</span>", unsafe_allow_html=True)

    st.markdown(f"## {ASSAY_LABELS[assay['type']]}")
    data = dict(assay.get("data", {}))

    if assay["type"] == "granulometria":
        render_granulometria_form(data)
    elif assay["type"] == "humedad":
        render_humedad_form(data)
    elif assay["type"] == "masa-unitaria":
        render_masa_unitaria_form(data)

    st.markdown('<div class="section-title">Observaciones</div>', unsafe_allow_html=True)
    observations = st.text_area("Observaciones", value=assay.get("observations", ""), label_visibility="collapsed", placeholder="Observaciones generales del ensayo…")

    st.markdown('<div class="section-title">Laboratorista</div>', unsafe_allow_html=True)
    laboratorist = st.text_input("Laboratorista", value=assay.get("laboratorist", ""), label_visibility="collapsed", placeholder="Nombre completo")

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾  Guardar borrador", use_container_width=True):
            assay.update(data=data, observations=observations, laboratorist=laboratorist, status="en-proceso", lastModified=now_iso())
            navigate("sample-summary")
    with col2:
        if st.button("✅  Finalizar ensayo", type="primary", use_container_width=True):
            assay.update(data=data, observations=observations, laboratorist=laboratorist, status="finalizado", lastModified=now_iso())
            navigate("sample-summary")

    if st.session_state.role == "jefe" and assay["type"] == "granulometria":
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_granulometria(code, sample, header, data)
        st.download_button(
            "📥  Descargar Excel (plantilla oficial de Granulometría)",
            data=excel_bytes,
            file_name=f"Granulometria_{code}_{sample}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
        )


# ════════════════════════════════════════════════════════════════════
# CONTINUAR / BUSCAR
# ════════════════════════════════════════════════════════════════════
def render_continue():
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Continuar ensayo")
    in_progress = [a for a in st.session_state.assays if a["status"] == "en-proceso"]
    if not in_progress:
        st.info("No hay ensayos en proceso.")
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
        df = pd.DataFrame([{"Proyecto": a["project"], "Muestra": a["sample"], "Ensayo": ASSAY_LABELS[a["type"]],
                             "Estado": STATUS_LABELS[a["status"]], "Última modificación": a["lastModified"]} for a in results])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No se encontraron ensayos con esos filtros.")


# ════════════════════════════════════════════════════════════════════
# BITÁCORA DE ORDEN (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_bitacora():
    require_role("jefe")
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## 📋 Bitácora orden para ensayos de laboratorio")

    codes = [p["code"] for p in st.session_state.projects] or ["(sin proyectos)"]
    project_code = st.selectbox("Proyecto", codes)
    bitacora = st.session_state.bitacoras.setdefault(project_code, {"meta": {}, "rows": []})

    st.markdown('<div class="section-title">Datos generales</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        bitacora["meta"]["fecha"] = st.text_input("Fecha", value=bitacora["meta"].get("fecha", ""))
        bitacora["meta"]["localizacion"] = st.text_input("Localización", value=bitacora["meta"].get("localizacion", ""))
    with c2:
        bitacora["meta"]["fuente_cantera"] = st.text_input("Fuente / Cantera", value=bitacora["meta"].get("fuente_cantera", ""))
        bitacora["meta"]["orden_asignada"] = st.text_input("Orden asignada a", value=bitacora["meta"].get("orden_asignada", ""))
    with c3:
        bitacora["meta"]["norma"] = st.text_input("Norma (NTC / IDU / RAS / GDA / Otro)", value=bitacora["meta"].get("norma", ""))
        bitacora["meta"]["revisada_por"] = st.text_input("Orden revisada por", value=bitacora["meta"].get("revisada_por", ""))

    st.markdown('<div class="section-title">Muestras y ensayos solicitados</div>', unsafe_allow_html=True)
    base_cols = {"N° Perforación": "", "N° Muestra": "", "Tipo de muestra": "", "Prof. De": None, "Prof. A": None}
    for e in BITACORA_ENSAYOS:
        base_cols[e] = False
    base_cols["Observaciones"] = ""

    if bitacora["rows"]:
        df = pd.DataFrame(bitacora["rows"])
        for col in base_cols:
            if col not in df.columns:
                df[col] = base_cols[col]
        df = df[list(base_cols.keys())]
    else:
        df = pd.DataFrame([base_cols])

    column_config = {e: st.column_config.CheckboxColumn(e) for e in BITACORA_ENSAYOS}
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config=column_config, key="bitacora_editor")
    bitacora["rows"] = edited.to_dict("records")

    st.markdown("<br>", unsafe_allow_html=True)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame([bitacora["meta"]]).to_excel(writer, index=False, sheet_name="Datos generales")
        edited.to_excel(writer, index=False, sheet_name="Muestras")
    st.download_button("📥  Descargar bitácora (Excel)", data=bio.getvalue(),
                        file_name=f"Bitacora_{project_code}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# ENRUTADOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════
if st.session_state.role is None:
    render_login()
else:
    render_sidebar()
    SCREENS = {
        "home": render_home, "new-project": render_new_project, "project-detail": render_project_detail,
        "new-sample": render_new_sample, "header-form": render_header_form, "sample-summary": render_sample_summary,
        "select-assay-type": render_select_assay_type, "assay-form": render_assay_form,
        "continue": render_continue, "search": render_search, "bitacora": render_bitacora,
    }
    SCREENS.get(st.session_state.screen, render_home)()
