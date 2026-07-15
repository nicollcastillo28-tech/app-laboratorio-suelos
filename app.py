"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Estructura: Proyecto -> Perforación (Sondeo/Apique/Fuente-Cantera) -> Muestra -> Ensayos

Cómo correrla en tu computador:
    streamlit run app.py
"""

import os
import uuid
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE LA PÁGINA
# ════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Geodelta Lab", page_icon="🧪", layout="wide", initial_sidebar_state="expanded")

APP_VERSION = "v3.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_GRANULOMETRIA = os.path.join(BASE_DIR, "templates", "CLASIFICACION_DE_SUELOS.xlsm")

# Claves de acceso. CÁMBIALAS aquí por las que tú quieras usar en tu laboratorio.
PASSWORDS = {"jefe": "geodelta2024", "auxiliar": "aux2024"}

# ════════════════════════════════════════════════════════════════════
# ESTILOS
# ════════════════════════════════════════════════════════════════════
NAVY, NAVY_HOVER, PRIMARY = "#1B2E4B", "#243D62", "#002046"
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

    /* ---- LOGIN ---- */
    .login-wrap {{ max-width: 420px; margin: 40px auto 0 auto; }}
    .login-icon {{
        width: 72px; height: 72px; border-radius: 999px; background: {SURFACE};
        border: 1px solid {BORDER}; display: flex; align-items: center; justify-content: center;
        font-size: 32px; margin: 0 auto 14px auto; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .login-title {{ text-align: center; color: {PRIMARY}; font-weight: 800; font-size: 26px; letter-spacing: -0.02em; margin-bottom: 28px; }}
    .login-card [data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 16px !important; border: 1px solid {BORDER} !important;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03) !important;
    }}
    .login-card div[role="radiogroup"] {{ gap: 10px; }}
    .login-card div[role="radiogroup"] label {{
        border: 1px solid {BORDER}; border-radius: 10px; padding: 10px 14px; flex: 1;
    }}
    .login-footer {{ text-align: center; color: {MUTED}; font-size: 12px; text-transform: uppercase;
        letter-spacing: 0.06em; margin-top: 22px; }}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════
# CONSTANTES DEL DOMINIO
# ════════════════════════════════════════════════════════════════════
SIEVES = [
    ("s_3", '3"', "76.2", "E20"), ("s_2p5", '2 1/2"', "63.5", "E21"), ("s_2", '2"', "50.8", "E22"),
    ("s_1p5", '1 1/2"', "38.1", "E23"), ("s_1", '1"', "25.4", "E24"), ("s_34", '3/4"', "19.05", "E25"),
    ("s_12", '1/2"', "12.7", "E26"), ("s_38", '3/8"', "9.52", "E27"), ("s_4", "No. 4", "4.76", "E28"),
    ("s_10", "No. 10", "2.00", "E29"), ("s_20", "No. 20", "0.841", "E30"), ("s_40", "No. 40", "0.42", "E31"),
    ("s_60", "No. 60", "0.25", "E32"), ("s_100", "No. 100", "0.149", "E33"), ("s_200", "No. 200", "0.075", "E34"),
]

ASSAY_LABELS = {"granulometria": "Granulometría", "humedad": "Contenido de humedad", "masa-unitaria": "Peso unitario"}
NORMAS_ENSAYO = {
    "granulometria": ["INV E-213", "INV E-214", "ASTM D422", "ASTM D7928"],
    "humedad": ["INV E-122", "ASTM D2216"],
    "masa-unitaria": ["INV E-202", "ASTM D1188"],
}
STATUS_LABELS = {"sin-iniciar": "Sin iniciar", "en-proceso": "En proceso", "finalizado": "Finalizado"}
STATUS_BADGE = {"sin-iniciar": "badge-muted", "en-proceso": "badge-warning", "finalizado": "badge-success"}

TIPO_PERFORACION_PREFIX = {"Sondeo": "S", "Apique": "AP", "Fuente/Cantera": "F"}
TIPO_MUESTRA_OPTIONS = ["Shelby", "NQ", "SS", "N/A"]
NORMA_PROYECTO_OPTIONS = ["NTC", "IDU", "RAS", "GDA", "Otro"]

BITACORA_ENSAYOS = [
    "Granulometría", "Pasa 200", "Humedad", "Límites de Atterberg", "Límite de contracción",
    "Materia orgánica", "Proctor", "CBR", "Compresión inconfinada", "Compresión en roca",
    "Peso unitario", "Gravedad específica", "Consolidación", "Corte CD", "Corte CU", "Corte UU", "Otro",
]
# Ensayos de la lista de arriba que ya tienen formulario propio en la app
SUPPORTED_ASSAY_MAP = {"Granulometría": "granulometria", "Humedad": "humedad", "Peso unitario": "masa-unitaria"}


# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL
# ════════════════════════════════════════════════════════════════════
def init_state():
    if "initialized" in st.session_state:
        return
    st.session_state.initialized = True
    st.session_state.role = None
    st.session_state.screen = "home"

    codigo_demo = "GDA-001-24"
    st.session_state.projects = [{
        "codigo_interno": codigo_demo, "numero": "001", "anio": "24",
        "nombre": "Estudio de suelos vía Bogotá-Medellín Km 14", "localizacion": "Sector Norte, Km 14+200",
        "fecha_bitacora": "2024-11-15", "fecha_ingreso_muestra": "2024-11-15", "norma": "GDA",
    }]
    st.session_state.perforaciones = {codigo_demo: [{"tipo": "Sondeo", "consecutivo": 1, "codigo": "S1"}]}
    st.session_state.muestras = {
        f"{codigo_demo}::S1": [{
            "numero": "1", "id_unico": f"{codigo_demo}_S1_M1", "profundidad_de": 0.0, "profundidad_hasta": 1.5,
            "tipo_muestra": "Shelby", "ensayos": {"Granulometría": True, "Humedad": True},
            "estado": "en-proceso",
        }]
    }
    st.session_state.assays = [{
        "id": "a001", "muestra_id": f"{codigo_demo}_S1_M1", "tipo": "granulometria", "status": "en-proceso",
        "data": {}, "observations": "", "laboratorist": "",
        "codigo_interno": codigo_demo, "perforacion_codigo": "S1", "muestra_numero": "1",
        "lastModified": datetime.now().isoformat(), "createdAt": datetime.now().isoformat(),
    }]

    st.session_state.selected_codigo = ""
    st.session_state.selected_perforacion = ""
    st.session_state.selected_muestra_id = ""
    st.session_state.selected_assay_id = None
    st.session_state.selected_assay_type = None


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


def get_project(codigo):
    return next((p for p in st.session_state.projects if p["codigo_interno"] == codigo), None)


def get_muestra(codigo, perforacion_codigo, muestra_id):
    for m in st.session_state.muestras.get(f"{codigo}::{perforacion_codigo}", []):
        if m["id_unico"] == muestra_id:
            return m
    return None


# ════════════════════════════════════════════════════════════════════
# LOGIN (estilo tarjeta centrada, inspirado en el mockup compartido)
# ════════════════════════════════════════════════════════════════════
def render_login():
    st.markdown("<br>", unsafe_allow_html=True)
    col = st.columns([1, 1.3, 1])[1]
    with col:
        st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
        st.markdown('<div class="login-icon">🧪</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">GEODELTA LAB</div>', unsafe_allow_html=True)

        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("#### Bienvenido de nuevo")
            st.caption("Ingresa tus credenciales para acceder al sistema.")
            role_choice = st.radio("Tipo de usuario", ["Auxiliar", "Jefe"], horizontal=True, label_visibility="visible")
            password = st.text_input("Clave de acceso", type="password", placeholder="••••••••")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("INGRESAR", type="primary", use_container_width=True):
                role_key = "jefe" if role_choice == "Jefe" else "auxiliar"
                if password == PASSWORDS[role_key]:
                    st.session_state.role = role_key
                    navigate("home")
                else:
                    st.error("Clave incorrecta.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="login-footer">🛠️ Geodelta Lab Engineering</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


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
            if st.button("📋  Bitácora de proyecto", use_container_width=True):
                navigate("bitacora")
        if st.button("📂  Continuar ensayo", use_container_width=True):
            navigate("continue")
        if st.button("🔎  Buscar ensayos", use_container_width=True):
            navigate("search")
        st.markdown("---")
        if st.session_state.selected_codigo:
            st.caption("Proyecto activo")
            st.markdown(f"**{st.session_state.selected_codigo}**")
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
        st.info("Como auxiliar puedes abrir proyectos existentes, digitar resultados de ensayos y actualizar el estado de las muestras.")

    st.markdown("### Proyectos")
    if not st.session_state.projects:
        st.info("Todavía no hay proyectos.")
    for p in st.session_state.projects:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([4, 1])
        with cols[0]:
            st.markdown(f"**{p['codigo_interno']}**")
            st.caption(p["nombre"])
        with cols[1]:
            if st.button("Abrir", key=f"open_{p['codigo_interno']}", use_container_width=True):
                st.session_state.selected_codigo = p["codigo_interno"]
                navigate("project-detail")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# NUEVO PROYECTO (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_new_project():
    require_role("jefe")
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## Nuevo proyecto")

    st.markdown('<div class="section-title">Código interno</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.text_input("Prefijo", value="GDA", disabled=True)
    with c2:
        numero = st.text_input("Número", placeholder="001")
    with c3:
        anio = st.text_input("Año", placeholder="24")

    codigo_interno = f"GDA-{numero}-{anio}" if numero and anio else ""
    existing_codes = [p["codigo_interno"] for p in st.session_state.projects]
    if codigo_interno:
        if codigo_interno in existing_codes:
            st.error(f"El código **{codigo_interno}** ya existe.")
        else:
            st.success(f"Código interno: **{codigo_interno}**")

    st.markdown('<div class="section-title">Información del proyecto</div>', unsafe_allow_html=True)
    nombre = st.text_input("Nombre del proyecto", placeholder="Estudio de suelos vía Bogotá-Medellín")
    localizacion = st.text_input("Localización", placeholder="Km 14+200")
    norma = st.selectbox("Norma", NORMA_PROYECTO_OPTIONS)

    c1, c2 = st.columns(2)
    with c1:
        fecha_bitacora = st.date_input("Fecha de bitácora", value=date.today())
    with c2:
        fecha_ingreso = st.date_input("Fecha de ingreso de muestra", value=date.today())

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            navigate("home")
    with col2:
        disabled = not codigo_interno or not nombre or codigo_interno in existing_codes
        if st.button("Crear proyecto →", type="primary", use_container_width=True, disabled=disabled):
            st.session_state.projects.append({
                "codigo_interno": codigo_interno, "numero": numero, "anio": anio, "nombre": nombre,
                "localizacion": localizacion, "norma": norma,
                "fecha_bitacora": str(fecha_bitacora), "fecha_ingreso_muestra": str(fecha_ingreso),
            })
            st.session_state.perforaciones[codigo_interno] = []
            st.session_state.selected_codigo = codigo_interno
            navigate("bitacora")


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PROYECTO → LISTA DE PERFORACIONES
# ════════════════════════════════════════════════════════════════════
def render_project_detail():
    codigo = st.session_state.selected_codigo
    project = get_project(codigo)
    if not project:
        navigate("home")
        return

    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown(f"## {project['codigo_interno']}")
    st.caption(project["nombre"])

    st.markdown('<div class="card">', unsafe_allow_html=True)
    cols = st.columns(4)
    cols[0].metric("Localización", project.get("localizacion") or "—")
    cols[1].metric("Norma", project.get("norma") or "—")
    cols[2].metric("Fecha bitácora", project.get("fecha_bitacora") or "—")
    cols[3].metric("Fecha ingreso muestra", project.get("fecha_ingreso_muestra") or "—")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.role == "jefe":
        if st.button("📋  Ir a la bitácora (agregar perforaciones y muestras)", type="primary"):
            navigate("bitacora")

    st.markdown("### Perforaciones")
    perforaciones = st.session_state.perforaciones.get(codigo, [])
    if not perforaciones:
        st.info("Este proyecto todavía no tiene perforaciones. Usa la Bitácora para agregarlas.")
    for perf in perforaciones:
        muestras = st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 2, 1])
        with cols[0]:
            st.markdown(f"**{perf['codigo']}** — {perf['tipo']}")
        with cols[1]:
            st.caption(f"{len(muestras)} muestra(s)")
        with cols[2]:
            if st.button("Abrir", key=f"open_perf_{perf['codigo']}", use_container_width=True):
                st.session_state.selected_perforacion = perf["codigo"]
                navigate("perforacion-detail")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PERFORACIÓN → LISTA DE MUESTRAS
# ════════════════════════════════════════════════════════════════════
def render_perforacion_detail():
    codigo = st.session_state.selected_codigo
    perf_codigo = st.session_state.selected_perforacion
    project = get_project(codigo)
    if not project:
        navigate("home")
        return

    st.button("← Atrás", on_click=lambda: navigate("project-detail"))
    st.markdown(f"## Perforación {perf_codigo}")
    st.caption(f"{project['codigo_interno']} · {project['nombre']}")

    muestras = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    if not muestras:
        st.info("Esta perforación todavía no tiene muestras. Agrégalas desde la Bitácora.")
    for m in muestras:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([2.5, 2, 1.3, 1])
        with cols[0]:
            st.markdown(f"**Muestra {m['numero']}**")
            st.caption(m["id_unico"])
        with cols[1]:
            st.caption(f"Prof. {m['profundidad_de']}–{m['profundidad_hasta']} m · {m['tipo_muestra']}")
        with cols[2]:
            st.markdown(f'<span class="badge {STATUS_BADGE[m["estado"]]}">{STATUS_LABELS[m["estado"]]}</span>', unsafe_allow_html=True)
        with cols[3]:
            if st.button("Abrir", key=f"open_muestra_{m['id_unico']}", use_container_width=True):
                st.session_state.selected_muestra_id = m["id_unico"]
                navigate("muestra-detail")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# BITÁCORA — crea perforaciones y muestras (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_bitacora():
    require_role("jefe")
    st.button("← Atrás", on_click=lambda: navigate("home"))
    st.markdown("## 📋 Bitácora orden para ensayos de laboratorio")

    codes = [p["codigo_interno"] for p in st.session_state.projects]
    if not codes:
        st.info("Primero crea un proyecto.")
        return
    default_idx = codes.index(st.session_state.selected_codigo) if st.session_state.selected_codigo in codes else 0
    codigo = st.selectbox("Proyecto", codes, index=default_idx)
    st.session_state.selected_codigo = codigo
    project = get_project(codigo)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f"**{project['nombre']}** · {project.get('localizacion','—')} · Norma {project.get('norma','—')}")
    st.markdown('</div>', unsafe_allow_html=True)

    perforaciones = st.session_state.perforaciones.setdefault(codigo, [])

    st.markdown('<div class="section-title">Agregar perforación</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([2, 1])
    with c1:
        tipo = st.selectbox("Tipo de perforación", list(TIPO_PERFORACION_PREFIX.keys()))
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Agregar perforación", use_container_width=True):
            prefix = TIPO_PERFORACION_PREFIX[tipo]
            consecutivo = len([p for p in perforaciones if p["tipo"] == tipo]) + 1
            codigo_perf = f"{prefix}{consecutivo}"
            perforaciones.append({"tipo": tipo, "consecutivo": consecutivo, "codigo": codigo_perf})
            st.session_state.muestras[f"{codigo}::{codigo_perf}"] = []
            st.rerun()

    st.markdown('<div class="section-title">Perforaciones y muestras</div>', unsafe_allow_html=True)
    if not perforaciones:
        st.info("Todavía no hay perforaciones en este proyecto. Agrega la primera arriba.")

    for perf in perforaciones:
        with st.expander(f"**{perf['codigo']}** — {perf['tipo']}", expanded=True):
            key = f"{codigo}::{perf['codigo']}"
            muestras = st.session_state.muestras.setdefault(key, [])

            base_cols = {"Número": "", "Prof. De": 0.0, "Prof. A": 0.0, "Tipo de muestra": "Shelby"}
            for e in BITACORA_ENSAYOS:
                base_cols[e] = False

            if muestras:
                rows = []
                for m in muestras:
                    row = {"Número": m["numero"], "Prof. De": m["profundidad_de"], "Prof. A": m["profundidad_hasta"], "Tipo de muestra": m["tipo_muestra"]}
                    for e in BITACORA_ENSAYOS:
                        row[e] = m["ensayos"].get(e, False)
                    rows.append(row)
                df = pd.DataFrame(rows)
            else:
                df = pd.DataFrame([base_cols])

            column_config = {
                "Tipo de muestra": st.column_config.SelectboxColumn(options=TIPO_MUESTRA_OPTIONS),
                "Prof. De": st.column_config.NumberColumn(format="%.2f"),
                "Prof. A": st.column_config.NumberColumn(format="%.2f"),
            }
            for e in BITACORA_ENSAYOS:
                column_config[e] = st.column_config.CheckboxColumn(e)

            edited = st.data_editor(
                df, num_rows="dynamic", use_container_width=True,
                column_config=column_config, key=f"editor_{key}",
            )

            nuevas_muestras = []
            for _, row in edited.iterrows():
                numero = str(row["Número"]).strip()
                if not numero or numero == "nan":
                    continue
                id_unico = f"{codigo}_{perf['codigo']}_M{numero}"
                existente = next((m for m in muestras if m["id_unico"] == id_unico), None)
                nuevas_muestras.append({
                    "numero": numero, "id_unico": id_unico,
                    "profundidad_de": row["Prof. De"], "profundidad_hasta": row["Prof. A"],
                    "tipo_muestra": row["Tipo de muestra"],
                    "ensayos": {e: bool(row[e]) for e in BITACORA_ENSAYOS},
                    "estado": existente["estado"] if existente else "sin-iniciar",
                })
            st.session_state.muestras[key] = nuevas_muestras

    st.markdown("<br>", unsafe_allow_html=True)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame([project]).to_excel(writer, index=False, sheet_name="Proyecto")
        all_rows = []
        for perf in perforaciones:
            for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
                r = {"Perforación": perf["codigo"], "Tipo": perf["tipo"], "N° Muestra": m["numero"],
                     "ID único": m["id_unico"], "Prof. De": m["profundidad_de"], "Prof. A": m["profundidad_hasta"],
                     "Tipo de muestra": m["tipo_muestra"], "Estado": STATUS_LABELS[m["estado"]]}
                r.update(m["ensayos"])
                all_rows.append(r)
        pd.DataFrame(all_rows).to_excel(writer, index=False, sheet_name="Muestras")
    st.download_button("📥  Descargar bitácora (Excel)", data=bio.getvalue(),
                        file_name=f"Bitacora_{codigo}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# DETALLE DE MUESTRA → LISTA DE ENSAYOS SOLICITADOS
# ════════════════════════════════════════════════════════════════════
def render_muestra_detail():
    codigo = st.session_state.selected_codigo
    perf_codigo = st.session_state.selected_perforacion
    muestra_id = st.session_state.selected_muestra_id
    muestra = get_muestra(codigo, perf_codigo, muestra_id)
    if not muestra:
        navigate("home")
        return

    st.button("← Atrás", on_click=lambda: navigate("perforacion-detail"))
    st.markdown(f"## Muestra {muestra['numero']}")
    st.caption(muestra["id_unico"])

    st.markdown('<div class="card">', unsafe_allow_html=True)
    cols = st.columns(4)
    cols[0].metric("Perforación", perf_codigo)
    cols[1].metric("Profundidad", f"{muestra['profundidad_de']}–{muestra['profundidad_hasta']} m")
    cols[2].metric("Tipo de muestra", muestra["tipo_muestra"])
    cols[3].metric("Proyecto", codigo)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Estado de la muestra</div>', unsafe_allow_html=True)
    options = list(STATUS_LABELS.keys())
    new_status = st.selectbox("Estado", options, index=options.index(muestra["estado"]), format_func=lambda k: STATUS_LABELS[k])
    if new_status != muestra["estado"]:
        muestra["estado"] = new_status
        st.rerun()

    st.markdown("### Ensayos solicitados")
    solicitados = [e for e, v in muestra["ensayos"].items() if v]
    if not solicitados:
        st.info("Esta muestra no tiene ensayos marcados en la bitácora.")
    for ensayo_label in solicitados:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        cols = st.columns([3, 1.4, 1])
        cols[0].markdown(f"**{ensayo_label}**")
        tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
        if tipo_interno:
            existing = next((a for a in st.session_state.assays if a["muestra_id"] == muestra_id and a["tipo"] == tipo_interno), None)
            status = existing["status"] if existing else "sin-iniciar"
            cols[1].markdown(f'<span class="badge {STATUS_BADGE[status]}">{STATUS_LABELS[status]}</span>', unsafe_allow_html=True)
            with cols[2]:
                if st.button("Abrir", key=f"open_ensayo_{ensayo_label}", use_container_width=True):
                    if existing:
                        st.session_state.selected_assay_id = existing["id"]
                    else:
                        new_id = f"a-{uuid.uuid4().hex[:8]}"
                        st.session_state.assays.append({
                            "id": new_id, "muestra_id": muestra_id, "tipo": tipo_interno, "status": "sin-iniciar",
                            "data": {}, "observations": "", "laboratorist": "",
                            "codigo_interno": codigo, "perforacion_codigo": perf_codigo, "muestra_numero": muestra["numero"],
                            "lastModified": now_iso(), "createdAt": now_iso(),
                        })
                        st.session_state.selected_assay_id = new_id
                    st.session_state.selected_assay_type = tipo_interno
                    navigate("assay-form")
        else:
            cols[1].markdown('<span class="badge badge-muted">Sin formulario aún</span>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE GRANULOMETRÍA (plantilla real del laboratorio)
# ════════════════════════════════════════════════════════════════════
def generar_excel_granulometria(codigo, perf_codigo, muestra, project, data):
    wb = load_workbook(TEMPLATE_GRANULOMETRIA, keep_vba=True)
    ws = wb["MUESTRA"]

    ws["D7"] = project["nombre"] if project else codigo
    ws["D9"] = project.get("localizacion", "") if project else ""
    ws["K7"] = project.get("fecha_ingreso_muestra", "") if project else ""
    ws["D12"] = perf_codigo
    ws["H12"] = muestra["numero"]
    ws["K12"] = to_float(muestra.get("profundidad_de"))
    ws["M12"] = to_float(muestra.get("profundidad_hasta"))
    ws["D13"] = f"Tipo de muestra: {muestra.get('tipo_muestra','')}"
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
    options = NORMAS_ENSAYO[assay_type]
    current = data.get(f"{key_prefix}_norma", "")
    idx = options.index(current) if current in options else 0
    choice = st.radio("Norma", options, index=idx, horizontal=True, key=f"norma_{key_prefix}", label_visibility="collapsed")
    data[f"{key_prefix}_norma"] = choice


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
    metodo = data.get("hum_metodo", "Método A")
    choice = st.radio("Método de ensayo", ["Método A", "Método B"], index=["Método A", "Método B"].index(metodo) if metodo in ["Método A", "Método B"] else 0, horizontal=True, key="hum_metodo_radio")
    data["hum_metodo"] = choice

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
    assay_id = st.session_state.selected_assay_id
    assay = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    if not assay:
        navigate("muestra-detail")
        return

    codigo, perf_codigo, muestra_id = assay["codigo_interno"], assay["perforacion_codigo"], assay["muestra_id"]
    project = get_project(codigo)
    muestra = get_muestra(codigo, perf_codigo, muestra_id)

    st.button("← Atrás", on_click=lambda: navigate("muestra-detail"))

    bar_cols = st.columns(5)
    bar_cols[0].markdown(f"**Proyecto**<br>{codigo}", unsafe_allow_html=True)
    bar_cols[1].markdown(f"**Muestra**<br>{muestra['numero'] if muestra else '—'}", unsafe_allow_html=True)
    bar_cols[2].markdown(f"**Ensayo**<br>{ASSAY_LABELS[assay['tipo']]}", unsafe_allow_html=True)
    bar_cols[3].markdown(f"**Perforación**<br>{perf_codigo}", unsafe_allow_html=True)
    bar_cols[4].markdown(f"**Estado**<br><span class='badge {STATUS_BADGE[assay['status']]}'>{STATUS_LABELS[assay['status']]}</span>", unsafe_allow_html=True)

    st.markdown(f"## {ASSAY_LABELS[assay['tipo']]}")
    data = dict(assay.get("data", {}))

    if assay["tipo"] == "granulometria":
        render_granulometria_form(data)
    elif assay["tipo"] == "humedad":
        render_humedad_form(data)
    elif assay["tipo"] == "masa-unitaria":
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
            navigate("muestra-detail")
    with col2:
        if st.button("✅  Finalizar ensayo", type="primary", use_container_width=True):
            assay.update(data=data, observations=observations, laboratorist=laboratorist, status="finalizado", lastModified=now_iso())
            navigate("muestra-detail")

    if st.session_state.role == "jefe" and assay["tipo"] == "granulometria" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_granulometria(codigo, perf_codigo, muestra, project, data)
        st.download_button(
            "📥  Descargar Excel (plantilla oficial de Granulometría)",
            data=excel_bytes, file_name=f"Granulometria_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True,
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
        cols[0].markdown(f"**{a['codigo_interno']}**")
        cols[1].markdown(f"{a['perforacion_codigo']} · Muestra {a['muestra_numero']}")
        cols[2].markdown(ASSAY_LABELS[a["tipo"]])
        with cols[3]:
            if st.button("Continuar", key=f"cont_{a['id']}", use_container_width=True):
                st.session_state.selected_codigo = a["codigo_interno"]
                st.session_state.selected_perforacion = a["perforacion_codigo"]
                st.session_state.selected_muestra_id = a["muestra_id"]
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
        f_muestra = st.text_input("Muestra")
    with c3:
        f_type = st.selectbox("Tipo de ensayo", ["(todos)"] + list(ASSAY_LABELS.values()))

    results = st.session_state.assays
    if f_project:
        results = [a for a in results if f_project.lower() in a["codigo_interno"].lower()]
    if f_muestra:
        results = [a for a in results if f_muestra.lower() in str(a["muestra_numero"]).lower()]
    if f_type != "(todos)":
        results = [a for a in results if ASSAY_LABELS[a["tipo"]] == f_type]

    if results:
        df = pd.DataFrame([{"Proyecto": a["codigo_interno"], "Perforación": a["perforacion_codigo"], "Muestra": a["muestra_numero"],
                             "Ensayo": ASSAY_LABELS[a["tipo"]], "Estado": STATUS_LABELS[a["status"]],
                             "Última modificación": a["lastModified"]} for a in results])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No se encontraron ensayos con esos filtros.")


# ════════════════════════════════════════════════════════════════════
# ENRUTADOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════
if st.session_state.role is None:
    render_login()
else:
    render_sidebar()
    SCREENS = {
        "home": render_home, "new-project": render_new_project, "project-detail": render_project_detail,
        "perforacion-detail": render_perforacion_detail, "muestra-detail": render_muestra_detail,
        "bitacora": render_bitacora, "assay-form": render_assay_form,
        "continue": render_continue, "search": render_search,
    }
    SCREENS.get(st.session_state.screen, render_home)()
