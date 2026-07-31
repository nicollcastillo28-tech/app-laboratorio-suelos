"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Estructura: Proyecto -> Perforación (Sondeo/Apique/Fuente-Cantera) -> Muestra -> Ensayo

Cómo correrla en tu computador:
    streamlit run app.py
"""

import html
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
st.set_page_config(page_title="Geodelta Lab", page_icon="🧪", layout="wide", initial_sidebar_state="collapsed")

APP_VERSION = "v5.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_GRANULOMETRIA = os.path.join(BASE_DIR, "templates", "CLASIFICACION_DE_SUELOS.xlsm")
TEMPLATE_BITACORA_ORDEN = os.path.join(BASE_DIR, "templates", "GDA-FL-003_bitacora_orden.xlsx")

PASSWORDS = {"jefe": "geodelta2024", "auxiliar": "aux2024"}

# ════════════════════════════════════════════════════════════════════
# ESTILOS — paleta del brief SoilLab Pro (Primary #1B365D · Secondary #4A6278 · Tertiary #005EB8 · Neutral #64748B)
# ════════════════════════════════════════════════════════════════════
PRIMARY, PRIMARY_DARK, PRIMARY_CONTAINER = "#002046", "#001B3D", "#1B365D"
SECONDARY, SECONDARY_CONTAINER = "#496177", "#C9E2FD"
TERTIARY = "#005EB8"
NEUTRAL = "#64748B"
SUCCESS, SUCCESS_LIGHT = "#16A34A", "#DCFCE7"
WARNING, WARNING_LIGHT = "#D97706", "#FEF3C7"
SURFACE, BG, BORDER, TEXT = "#FFFFFF", "#F8F9FF", "#C4C6CF", "#0B1C30"
MUTED = NEUTRAL

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap');
    .material-symbols-outlined, [data-testid="stMarkdownContainer"] span.material-symbols-outlined,
    [data-testid="stMarkdownContainer"] p span.material-symbols-outlined {{
        font-family: 'Material Symbols Outlined' !important;
        font-weight: normal; font-style: normal; text-transform: none;
        letter-spacing: normal; word-wrap: normal; white-space: nowrap; direction: ltr;
        -webkit-font-smoothing: antialiased;
        font-variation-settings: 'FILL' 0, 'wght' 400, 'GRAD' 0, 'opsz' 24;
        vertical-align: middle; line-height: 1; font-size: 1.15em; display: inline-block;
    }}
    .msi-fill {{ font-variation-settings: 'FILL' 1, 'wght' 400, 'GRAD' 0, 'opsz' 24; }}
    html, body, [class*="css"] {{ font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif; }}
    [data-testid="stAppViewContainer"], [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] span, [data-testid="stMarkdownContainer"] li {{
        font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif !important;
    }}
    .stApp {{ background-color: {BG}; }}
    [data-testid="collapsedControl"] {{ display: none; }}
    section[data-testid="stSidebar"] {{ display: none; }}
    .font-mono {{ font-family: 'JetBrains Mono', monospace; }}

    /* ---- TOP APP BAR (desktop / tablet ancho) ---- */
    .st-key-topbar {{
        position: sticky; top: 0; z-index: 999; background: {SURFACE};
        border-bottom: 1px solid {BORDER}; padding: 10px 4px 6px 4px; margin-bottom: 8px;
    }}
    .st-key-topbar .stButton button {{
        font-family: 'JetBrains Mono', monospace; font-weight: 700;
        letter-spacing: 0.04em; text-transform: uppercase; white-space: nowrap;
        font-size: clamp(10px, 1.1vw, 12px); padding-left: 8px; padding-right: 8px;
    }}
    .topbar-brand {{ display: flex; align-items: center; gap: 10px; height: 38px; }}
    .topbar-brand .brand-title {{
        font-size: clamp(15px, 2vw, 20px); font-weight: 700; color: {PRIMARY}; letter-spacing: -0.02em; white-space: nowrap;
    }}
    .topbar-avatar {{
        width: 36px; height: 36px; border-radius: 999px; background: {PRIMARY_CONTAINER}; color: #FFFFFF;
        display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px;
        border: 1px solid {BORDER}; margin-left: auto; flex-shrink: 0;
    }}

    /* ---- BOTTOM NAV (celular y tablet en vertical) ---- */
    .st-key-bottomnav {{ display: none; }}
    @media (max-width: 900px) {{
        .st-key-topbar-nav {{ display: none; }}
        div[data-testid="stColumn"]:has(.st-key-topbar-nav) {{ display: none; }}
        .st-key-bottomnav {{
            display: block; position: fixed; bottom: 0; left: 0; width: 100%; z-index: 999;
            background: {SURFACE}; border-top: 1px solid {BORDER}; padding: 6px 8px 8px 8px; box-shadow: 0 -2px 8px rgba(0,0,0,0.04);
        }}
        .st-key-bottomnav .stButton button {{
            font-family: 'JetBrains Mono', monospace; font-size: clamp(8px, 2.6vw, 10px); text-transform: uppercase;
            letter-spacing: 0.02em; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
            padding-top: 10px; padding-bottom: 10px; padding-left: 2px; padding-right: 2px;
        }}
        .st-key-bottomnav [data-testid="stHorizontalBlock"] {{ flex-direction: row !important; flex-wrap: nowrap !important; gap: 6px !important; }}
        .st-key-bottomnav [data-testid="stColumn"] {{ width: auto !important; flex: 1 1 0 !important; min-width: 0 !important; }}
        .main .block-container {{ padding-bottom: 76px; }}
    }}
    @media (max-width: 420px) {{
        .topbar-brand .brand-title {{ display: none; }}
        .st-key-bottomnav .stButton button {{ font-size: 9px; }}
    }}

    /* Contenedores con borde nativos de Streamlit = nuestras "tarjetas" (sin bugs de HTML suelto) */
    /* OJO: en esta versión de Streamlit ya no existe stVerticalBlockBorderWrapper como wrapper
    aparte — st.container(border=True) marca el propio stVerticalBlock con
    data-test-scroll-behavior="normal" (no lo tienen los stVerticalBlock sin borde). */
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"] {{
        border-radius: 12px !important; border: 1px solid {BORDER} !important;
        box-shadow: 0 1px 4px rgba(11,28,48,0.08) !important; background: {SURFACE} !important;
    }}

    .section-title {{
        font-size: 12px; font-weight: 700; color: {MUTED}; text-transform: uppercase;
        letter-spacing: 0.06em; border-bottom: 1px solid {BORDER}; padding-bottom: 8px;
        margin-bottom: 14px; margin-top: 4px;
    }}
    .badge {{
        display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700;
        font-family: 'IBM Plex Sans', sans-serif;
    }}
    .badge-success {{ background: {SUCCESS_LIGHT}; color: {SUCCESS}; }}
    .badge-warning {{ background: {WARNING_LIGHT}; color: {WARNING}; }}
    .badge-muted {{ background: #EEF1F5; color: {MUTED}; }}
    .role-pill {{
        display: inline-block; padding: 4px 12px; border-radius: 999px; font-size: 11px;
        font-family: 'JetBrains Mono', monospace; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase;
        background: {SECONDARY_CONTAINER}; color: {PRIMARY};
    }}
    .timestamp-caption {{ color: {MUTED}; font-size: 12px; margin-top: 2px; }}
    div.stButton > button[kind="primary"] {{ background-color: {PRIMARY}; border: 1px solid {PRIMARY}; }}
    div.stButton > button[kind="primary"]:hover {{ background-color: {PRIMARY_DARK}; border-color: {PRIMARY_DARK}; }}
    h1, h2, h3, h4, h5, h6 {{ color: {TEXT}; letter-spacing: -0.02em; font-family: 'IBM Plex Sans', sans-serif !important; }}

    /* Campos digitables con fondo distinto al de la página, para que se note qué se puede editar */
    .stTextInput input, .stTextArea textarea, .stNumberInput input,
    .stDateInput input, .stSelectbox > div > div, .stMultiSelect > div > div {{
        background-color: {SURFACE} !important;
        border: 1px solid {BORDER} !important;
        border-radius: 8px !important;
    }}
    .stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus {{
        border-color: {PRIMARY} !important; box-shadow: 0 0 0 1px {PRIMARY} !important;
    }}
    [data-testid="stDataFrameResizable"], [data-testid="stDataEditorGrid"] {{
        background-color: {SURFACE} !important; border: 1px solid {BORDER} !important; border-radius: 8px;
    }}

    .login-icon {{
        width: 56px; height: 56px; border-radius: 16px; background: {SECONDARY_CONTAINER};
        display: flex; align-items: center; justify-content: center; color: {PRIMARY};
        font-size: 26px; margin: 0 auto 10px auto;
    }}
    .login-title {{ text-align: center; color: {TEXT}; font-weight: 600; font-size: 16px; letter-spacing: -0.01em; margin-bottom: 20px; }}
    .login-footer {{ text-align: center; color: {NEUTRAL}; font-size: 12px; margin-top: 18px; }}

    /* Selector de rol en el login, como tarjetas seleccionables */
    .st-key-login-card [data-testid="stRadio"] > div[role="radiogroup"] {{ display: flex; gap: 10px; }}
    .st-key-login-card [data-testid="stRadio"] [role="radiogroup"] label {{
        flex: 1 1 0; border: 1px solid {BORDER}; border-radius: 10px; padding: 8px 12px !important;
        margin: 0 !important; background: {SURFACE}; white-space: nowrap;
    }}
    .st-key-login-card [data-testid="stRadio"] [role="radiogroup"] label p {{ white-space: nowrap; }}
    .st-key-login-card [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
        border-color: {PRIMARY}; background: {SECONDARY_CONTAINER};
    }}
    .st-key-login-card [data-testid="stRadio"] input[type="radio"] {{ accent-color: {PRIMARY}; }}
    .st-key-login-card [data-testid="stWidgetLabel"] p {{
        font-family: 'JetBrains Mono', monospace; font-size: 11px; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.05em; color: {NEUTRAL};
    }}
    .st-key-login-card .stButton:has(button[kind="secondary"]) button {{
        color: {NEUTRAL}; border: none; background: transparent;
    }}

    /* ---- BENTO CARDS (inspirado en el diseño de Stitch) ---- */
    .bento-primary {{
        background: linear-gradient(135deg, {PRIMARY_CONTAINER} 0%, {PRIMARY} 100%);
        color: #FFFFFF; border-radius: 16px; padding: 26px 28px; min-height: 168px;
        display: flex; flex-direction: column; justify-content: space-between; margin-bottom: 16px;
    }}
    .bento-primary .bento-icon {{
        background: rgba(255,255,255,0.14); width: 44px; height: 44px; border-radius: 10px;
        display: flex; align-items: center; justify-content: center; font-size: 22px; margin-bottom: 14px;
    }}
    .bento-primary .bento-eyebrow {{ font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.65; }}
    .bento-primary h3 {{ color: #FFFFFF; margin: 4px 0 6px 0; }}
    .bento-primary p {{ opacity: 0.75; font-size: 13px; margin: 0; }}

    .bento-light {{
        background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 16px; padding: 24px 26px;
        min-height: 168px; display: flex; flex-direction: column; justify-content: space-between; margin-bottom: 16px;
    }}
    .bento-light .bento-icon {{
        background: {SECONDARY_CONTAINER}; color: {PRIMARY}; width: 44px; height: 44px; border-radius: 10px;
        display: flex; align-items: center; justify-content: center; font-size: 20px; margin-bottom: 14px;
    }}
    .bento-light h3 {{ color: {PRIMARY}; margin: 4px 0 6px 0; font-size: 18px; }}

    /* Tarjetas de la fila de acciones de Inicio, todas del mismo alto y alineadas */
    .st-key-home-actions [data-testid="stHorizontalBlock"] {{ align-items: stretch; }}
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-primary),
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-light) {{ flex: 1 1 auto; }}
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-primary) .stMarkdown,
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-light) .stMarkdown,
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-primary) .stMarkdown > div,
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-light) .stMarkdown > div,
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-primary) [data-testid="stMarkdownContainer"],
    .st-key-home-actions [data-testid="stElementContainer"]:has(.bento-light) [data-testid="stMarkdownContainer"] {{
        height: 100%;
    }}
    .st-key-home-actions .bento-primary, .st-key-home-actions .bento-light {{ height: 100%; }}
    .bento-light p {{ color: {MUTED}; font-size: 13px; margin: 0; }}

    .stat-chip {{
        background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 12px; padding: 14px 16px;
        display: flex; align-items: center; gap: 12px; margin-bottom: 12px;
    }}
    .stat-chip .stat-icon {{
        width: 40px; height: 40px; border-radius: 999px; background: {SECONDARY_CONTAINER};
        display: flex; align-items: center; justify-content: center; font-size: 18px;
    }}
    .stat-chip .stat-label {{ font-size: 11px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: {MUTED}; }}
    .stat-chip .stat-value {{ font-size: 20px; font-weight: 800; color: {TEXT}; }}

    /* ---- ACTIVITY TABLE (inspirado en el diseño de Stitch) ---- */
    .activity-table-wrap {{ overflow-x: auto; }}
    .activity-table {{ width: 100%; border-collapse: collapse; font-family: 'IBM Plex Sans', sans-serif; }}
    .activity-table thead th {{
        background: {SECONDARY_CONTAINER}; color: {PRIMARY}; font-family: 'JetBrains Mono', monospace;
        font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
        padding: 10px 14px; text-align: left; white-space: nowrap; border-bottom: 1px solid {BORDER};
    }}
    .activity-table tbody td {{
        padding: 12px 14px; border-bottom: 1px solid {BORDER}; font-size: 14px; color: {TEXT}; vertical-align: middle;
    }}
    .activity-table tbody tr:last-child td {{ border-bottom: none; }}
    .activity-table tbody tr:hover {{ background: {BG}; }}
    .activity-table .cell-id {{ font-family: 'JetBrains Mono', monospace; color: {PRIMARY}; font-weight: 800; font-size: 15px; }}
    .activity-table .cell-title {{ font-weight: 600; color: {TEXT}; }}
    .activity-table .cell-sub {{ font-size: 12px; color: {NEUTRAL}; margin-top: 1px; }}
    .activity-table .cell-muted {{ color: {NEUTRAL}; font-size: 13px; }}
    .activity-footer {{
        display: flex; justify-content: space-between; align-items: center; padding: 10px 14px;
        color: {NEUTRAL}; font-size: 13px;
    }}

    /* ---- ENSAYOS ASIGNADOS (panel de Auxiliar) ---- */
    .assigned-th {{
        background: {SECONDARY_CONTAINER}; color: {PRIMARY}; font-family: 'JetBrains Mono', monospace;
        font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
        padding: 8px 10px; border-radius: 6px; margin-bottom: 4px; white-space: nowrap;
    }}
    .assigned-chip {{
        background: {BG}; border: 1px solid {BORDER}; color: {PRIMARY}; font-size: 12px; font-weight: 700;
        padding: 3px 10px; border-radius: 6px; display: inline-block;
    }}

    /* ---- TARJETAS DE PROYECTO (Proyectos en ejecución) ---- */
    .code-badge {{
        display: inline-block; background: {SECONDARY_CONTAINER}; color: {PRIMARY};
        font-family: 'JetBrains Mono', monospace; font-weight: 800; font-size: 16px;
        letter-spacing: 0.02em; padding: 4px 14px; border-radius: 6px;
    }}
    [class*="st-key-projcard_"] {{
        border-left: 4px solid {PRIMARY} !important;
        background: {SURFACE} !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
    }}

    .perf-code-box {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 30px; height: 30px; border-radius: 7px; background: {PRIMARY_CONTAINER};
        color: #fff; font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 12px;
    }}

    /* Botón flotante "+" para crear proyecto desde Proyectos en ejecución */
    .st-key-fab-new-project {{
        position: fixed; right: 24px; bottom: 28px; z-index: 998; width: 56px !important;
    }}
    .st-key-fab-new-project .stButton button {{
        width: 56px; height: 56px; border-radius: 999px !important; padding: 0 !important;
        background: {PRIMARY} !important; border: none !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.28) !important;
    }}
    .st-key-fab-new-project .stButton button span[data-testid="stIconMaterial"] {{
        color: #fff !important; font-size: 26px !important;
    }}
    @media (max-width: 900px) {{
        .st-key-fab-new-project {{ bottom: 92px; right: 16px; }}
    }}
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
STATUS_ICON = {"sin-iniciar": "radio_button_unchecked", "en-proceso": "autorenew", "finalizado": "check_circle"}

TIPO_PERFORACION_PREFIX = {"Sondeo": "S", "Apique": "AP", "Fuente/Cantera": "F"}
TIPO_MUESTRA_OPTIONS = ["Shelby", "NQ", "SS", "N/A"]
NORMA_PROYECTO_OPTIONS = ["IDU", "NTC", "INVIAS", "Otro"]

# Lista de equipos del laboratorio. Por ahora sin código — agrega o edita los que tengas aquí.
EQUIPO_LIST = [
    "Balanza digital 0.01g", "Balanza digital 0.1g", "Horno de secado", "Tamices serie gruesa",
    "Tamices serie fina", "Tamizadora mecánica", "Cazuela de Casagrande", "Ranurador", "Copa de Casagrande",
    "Molde Proctor estándar", "Molde Proctor modificado", "Prensa CBR", "Balanza hidrostática",
    "Horno de parafinado", "Cronómetro", "Termómetro", "Extractor de muestras", "Otro",
]

BITACORA_ENSAYOS = [
    "Granulometría", "Pasa 200", "Humedad", "Límites de Atterberg", "Límite de contracción",
    "Materia orgánica", "Proctor", "CBR", "Compresión inconfinada", "Compresión en roca",
    "Peso unitario", "Gravedad específica", "Consolidación", "Corte CD", "Corte CU", "Corte UU", "Otro",
]
SUPPORTED_ASSAY_MAP = {"Granulometría": "granulometria", "Humedad": "humedad", "Peso unitario": "masa-unitaria"}

BITACORA_BASE_COLS = ["Número", "Prof. De", "Prof. A", "Tipo de muestra"] + BITACORA_ENSAYOS + ["Observaciones"]

# Celdas de la plantilla oficial GDA-FL-003 (hoja "S1"): columna por ensayo, checkbox de
# norma y checkbox de tipo de perforación. Ensayos sin columna en la plantilla (p. ej.
# Carga puntual, Desgaste) simplemente no se marcan.
BITACORA_XLSX_ENSAYO_COL = {
    "Granulometría": "F", "Pasa 200": "G", "Humedad": "H", "Límites de Atterberg": "I",
    "Límite de contracción": "J", "Materia orgánica": "K", "Proctor": "L", "CBR": "M",
    "Compresión inconfinada": "N", "Compresión en roca": "O", "Peso unitario": "P",
    "Gravedad específica": "Q", "Consolidación": "R", "Corte CD": "S", "Corte CU": "T",
    "Corte UU": "U", "Otro": "AG",
}
BITACORA_XLSX_NORMA_CELL = {"IDU": "AG10", "INVIAS": "AI10", "NTC": "AG12", "Otro": "AI12"}
BITACORA_XLSX_TIPO_CELL = {"Sondeo": "H14", "Apique": "D14", "Fuente/Cantera": "AH14"}
BITACORA_XLSX_MAX_ROWS = 14  # la plantilla trae 14 filas fijas (18 a 31)


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
            "numero": "1", "id_unico": f"{codigo_demo}-S1-M1", "profundidad_de": 0.0, "profundidad_hasta": 1.5,
            "tipo_muestra": "Shelby", "ensayos": {"Granulometría": True, "Humedad": True}, "observaciones": "",
        }]
    }
    st.session_state.assays = [{
        "id": "a001", "muestra_id": f"{codigo_demo}-S1-M1", "tipo": "granulometria", "status": "en-proceso",
        "data": {}, "observations": "", "laboratorist": "",
        "codigo_interno": codigo_demo, "perforacion_codigo": "S1", "muestra_numero": "1",
        "lastModified": datetime.now().isoformat(), "createdAt": datetime.now().isoformat(),
    }]

    st.session_state.nav_stack = []
    st.session_state.bitacora_draft = {}
    st.session_state.sieve_draft = {}
    st.session_state.selected_codigo = ""
    st.session_state.selected_perforacion = ""
    st.session_state.selected_muestra_id = ""
    st.session_state.selected_assay_id = None
    st.session_state.selected_assay_type = None
    st.session_state.read_only_view = False


init_state()


def navigate(screen):
    actual = st.session_state.get("screen")
    if actual and actual != screen:
        st.session_state.nav_stack.append(actual)
    st.session_state.screen = screen
    st.rerun()


def go_back(fallback="home"):
    """Vuelve a la pantalla realmente anterior (pila de navegación) en vez de un destino fijo."""
    if st.session_state.nav_stack:
        st.session_state.screen = st.session_state.nav_stack.pop()
    else:
        st.session_state.screen = fallback
    st.rerun()


def to_float(v, default=None):
    try:
        return float(str(v).replace(",", "."))
    except (ValueError, TypeError):
        return default


def icon(name, size=18, fill=False, color=None):
    """Ícono de Material Symbols para insertar dentro de HTML propio (st.markdown con unsafe_allow_html)."""
    cls = "material-symbols-outlined msi-fill" if fill else "material-symbols-outlined"
    style = f"font-size:{size}px;"
    if color:
        style += f"color:{color};"
    return f'<span class="{cls}" style="{style}">{name}</span>'


def status_badge_html(status, font_size=None):
    style = f' style="font-size:{font_size}px;"' if font_size else ""
    return (f'<span class="badge {STATUS_BADGE[status]}"{style}>'
            f'{icon(STATUS_ICON[status], size=13)} {STATUS_LABELS[status]}</span>')


def now_iso():
    return datetime.now().isoformat()


def format_dt(iso_str):
    try:
        return datetime.fromisoformat(iso_str).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return "—"


def require_role(*allowed):
    if st.session_state.role not in allowed:
        st.warning("No tienes permiso para ver esta sección.")
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


def get_assay(muestra_id, tipo_interno):
    return next((a for a in st.session_state.assays if a["muestra_id"] == muestra_id and a["tipo"] == tipo_interno), None)


def compute_muestra_estado(muestra):
    """El estado de la muestra se calcula solo, a partir del estado de cada ensayo solicitado."""
    statuses = []
    for label, activo in muestra["ensayos"].items():
        if not activo:
            continue
        tipo_interno = SUPPORTED_ASSAY_MAP.get(label)
        if not tipo_interno:
            continue
        a = get_assay(muestra["id_unico"], tipo_interno)
        statuses.append(a["status"] if a else "sin-iniciar")
    if not statuses:
        return "sin-iniciar"
    if all(s == "finalizado" for s in statuses):
        return "finalizado"
    if any(s in ("en-proceso", "finalizado") for s in statuses):
        return "en-proceso"
    return "sin-iniciar"


def project_progress(codigo):
    counts = {"sin-iniciar": 0, "en-proceso": 0, "finalizado": 0}
    for perf in st.session_state.perforaciones.get(codigo, []):
        for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
            counts[compute_muestra_estado(m)] += 1
    return counts


def project_status(codigo):
    """'ejecutado' solo si el proyecto tiene al menos una muestra y TODAS están finalizadas."""
    counts = project_progress(codigo)
    total = sum(counts.values())
    if total > 0 and counts["finalizado"] == total:
        return "ejecutado"
    return "ejecucion"


def confirm_delete(action_key, label):
    """Botón de eliminar con confirmación en dos pasos. Devuelve True solo cuando se confirma."""
    flag = f"confirm_{action_key}"
    if st.session_state.get(flag):
        st.warning(f"¿Eliminar {label}? Esta acción no se puede deshacer.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Sí, eliminar", key=f"yes_{action_key}", type="primary", use_container_width=True):
                st.session_state[flag] = False
                return True
        with c2:
            if st.button("Cancelar", key=f"no_{action_key}", use_container_width=True):
                st.session_state[flag] = False
                st.rerun()
        return False
    if st.button("Eliminar", key=f"del_{action_key}", use_container_width=True, icon=":material/delete:"):
        st.session_state[flag] = True
        st.rerun()
    return False


# ════════════════════════════════════════════════════════════════════
# LOGIN
# ════════════════════════════════════════════════════════════════════
def render_login():
    st.markdown("<br>", unsafe_allow_html=True)
    col = st.columns([1, 1.3, 1])[1]
    with col:
        st.markdown(f'<div class="login-icon">{icon("biotech", size=26)}</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Geodelta Lab</div>', unsafe_allow_html=True)
        with st.container(border=True, key="login-card"):
            st.markdown("#### Bienvenido de nuevo")
            st.caption("Ingresa tus credenciales para acceder al sistema.")
            role_choice = st.radio("Tipo de usuario", ["Auxiliar", "Jefe"], horizontal=True)
            password = st.text_input("Clave de acceso", type="password", placeholder="••••••••")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("INGRESAR", type="primary", use_container_width=True):
                role_key = "jefe" if role_choice == "Jefe" else "auxiliar"
                if password == PASSWORDS[role_key]:
                    st.session_state.role = role_key
                    st.session_state.nav_stack = []
                    navigate("home")
                else:
                    st.error("Clave incorrecta.")
            st.markdown('<hr style="margin:16px 0 4px 0;">', unsafe_allow_html=True)
            if st.button("¿Olvidaste tu clave?", key="forgot_pwd", type="secondary", use_container_width=True):
                st.info("Contacta al Jefe de laboratorio para restablecer tu clave de acceso.")
        st.markdown(f'<div class="login-footer">{icon("build", size=14)} Geodelta Lab Engineering</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# NAVEGACIÓN — TopAppBar + BottomNav (reemplaza el sidebar)
# ════════════════════════════════════════════════════════════════════
NAV_ITEMS = [
    ("home", "Inicio", "home"), ("projects-active", "Proyectos", "folder"), ("search", "Buscar", "search"),
]
ACTIVE_MAP = {
    "home": "home",
    "projects-active": "projects-active", "projects-done": "projects-active", "new-project": "projects-active",
    "project-detail": "projects-active", "edit-project": "projects-active",
    "perforacion-detail": "projects-active", "muestra-detail": "projects-active",
    "bitacora": "projects-active", "continue": "projects-active", "assay-form": "projects-active",
    "search": "search",
}


def render_topbar():
    active = ACTIVE_MAP.get(st.session_state.screen)
    with st.container(key="topbar"):
        c_brand, c_nav, c_avatar, c_logout = st.columns([2.4, 4.6, 0.7, 0.7])
        with c_brand:
            st.markdown(f'<div class="topbar-brand">{icon("biotech", size=24)}'
                        f'<span class="brand-title">Geodelta Lab</span></div>', unsafe_allow_html=True)
        with c_nav:
            with st.container(key="topbar-nav"):
                cols = st.columns(len(NAV_ITEMS))
                for col, (key, label, icono) in zip(cols, NAV_ITEMS):
                    with col:
                        if st.button(label, key=f"nav_{key}", use_container_width=True, icon=f":material/{icono}:",
                                     type="primary" if active == key else "secondary"):
                            navigate(key)
        with c_avatar:
            iniciales = "JL" if st.session_state.role == "jefe" else "AX"
            st.markdown(f'<div class="topbar-avatar">{iniciales}</div>', unsafe_allow_html=True)
        with c_logout:
            if st.button("", key="logout_top", help="Cerrar sesión", use_container_width=True, icon=":material/logout:"):
                st.session_state.role = None
                st.session_state.nav_stack = []
                navigate("home")


def render_bottomnav():
    active = ACTIVE_MAP.get(st.session_state.screen)
    with st.container(key="bottomnav"):
        cols = st.columns(len(NAV_ITEMS))
        for col, (key, label, icono) in zip(cols, NAV_ITEMS):
            with col:
                if st.button(label, key=f"bnav_{key}", use_container_width=True, icon=f":material/{icono}:",
                             type="primary" if active == key else "secondary"):
                    navigate(key)


# ════════════════════════════════════════════════════════════════════
# INICIO
# ════════════════════════════════════════════════════════════════════
def render_home():
    es_jefe = st.session_state.role == "jefe"
    if es_jefe:
        st.markdown("## Bienvenido, Jefe de Laboratorio")
        st.caption("Resumen de operaciones y control de calidad geotécnica para hoy.")
    else:
        st.markdown("## Panel de Auxiliar")
        st.caption("Gestiona tus proyectos asignados y registra los resultados de los ensayos de suelo.")

    st.markdown("<br>", unsafe_allow_html=True)

    with st.container(key="home-actions"):
        if es_jefe:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="bento-primary"><div class="bento-icon">{icon("add")}</div>'
                             '<div><h3>Crear nuevo proyecto</h3><p>Registrar nuevo cliente y parámetros de sitio.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Crear proyecto →", key="cta_new_project", use_container_width=True):
                    navigate("new-project")
            with c2:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("sync")}</div>'
                             f'<div><h3>Proyectos en ejecución</h3><p>{sum(1 for p in st.session_state.projects if project_status(p["codigo_interno"])=="ejecucion")} proyecto(s) activos en laboratorio.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Ver proyectos →", key="cta_active", use_container_width=True):
                    navigate("projects-active")
            with c3:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("archive")}</div>'
                             '<div><h3>Proyectos ejecutados</h3><p>Revisar reportes finales y resultados certificados.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Explorar archivo →", key="cta_done", use_container_width=True):
                    navigate("projects-done")
        else:
            c1, c2 = st.columns([2, 1])
            with c1:
                activos = sum(1 for p in st.session_state.projects if project_status(p["codigo_interno"]) == "ejecucion")
                st.markdown(f'<div class="bento-primary"><div class="bento-icon">{icon("assignment")}</div>'
                             f'<div><span class="bento-eyebrow">Tareas prioritarias</span>'
                             f'<h3>Proyectos en ejecución</h3><p>Accede a los proyectos activos para registrar granulometría, humedad y peso unitario.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button(f"Ver proyectos → ({activos} activos)", key="cta_active_aux", use_container_width=True):
                    navigate("projects-active")
            with c2:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("archive")}</div>'
                             '<div><h3>Proyectos ejecutados</h3><p>Consulta el historial. Solo lectura.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Explorar archivo →", key="cta_done_aux", use_container_width=True):
                    navigate("projects-done")

    st.markdown("<br>", unsafe_allow_html=True)
    todos_los_ensayos = sorted(st.session_state.assays, key=lambda a: a["lastModified"], reverse=True)

    if es_jefe:
        recientes = todos_los_ensayos[:5]
        with st.container(border=True):
            h1, h2 = st.columns([4, 1])
            with h1:
                st.markdown(f'<div class="section-title" style="border-bottom:none;margin-bottom:0;padding-bottom:0;">'
                            f'{icon("history", size=15)} Actividad reciente</div>', unsafe_allow_html=True)
            with h2:
                if st.button("Ver todo →", key="cta_ver_todo_actividad", use_container_width=True):
                    navigate("search")

            if not recientes:
                st.info("Todavía no hay actividad registrada.")
            else:
                col_ratios = [1.4, 2.6, 1.8, 1.3, 0.9]
                headers = st.columns(col_ratios)
                for col, label in zip(headers, ["ID proyecto", "Cliente / Ubicación", "Última actualización", "Estado", "Acción"]):
                    col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
                for a in recientes:
                    proyecto = get_project(a["codigo_interno"])
                    titulo = html.escape(proyecto["nombre"] if proyecto else a["codigo_interno"])
                    subtitulo = html.escape(f'{a["perforacion_codigo"]} · Muestra {a["muestra_numero"]} · {ASSAY_LABELS[a["tipo"]]}')
                    actualizacion = format_dt(a["lastModified"])
                    if a.get("laboratorist"):
                        actualizacion += f' · {html.escape(a["laboratorist"])}'
                    cols = st.columns(col_ratios, vertical_alignment="center")
                    cols[0].markdown(f'<span class="cell-id">{html.escape(a["codigo_interno"])}</span>', unsafe_allow_html=True)
                    cols[1].markdown(f'<div class="cell-title">{titulo}</div><div class="cell-sub">{subtitulo}</div>',
                                      unsafe_allow_html=True)
                    cols[2].markdown(f'<span class="cell-muted">{html.escape(actualizacion)}</span>', unsafe_allow_html=True)
                    cols[3].markdown(status_badge_html(a["status"]), unsafe_allow_html=True)
                    with cols[4]:
                        if st.button("Abrir", key=f"open_recent_{a['id']}", use_container_width=True):
                            st.session_state.selected_codigo = a["codigo_interno"]
                            navigate("project-detail")
                st.markdown(f'<div class="activity-footer">Mostrando {len(recientes)} de {len(todos_los_ensayos)} ensayo(s)</div>',
                            unsafe_allow_html=True)
    else:
        pendientes = [a for a in todos_los_ensayos if a["status"] != "finalizado"]
        with st.container(border=True):
            h1, h2 = st.columns([4, 1])
            with h1:
                st.markdown(f'<div class="section-title" style="border-bottom:none;margin-bottom:0;padding-bottom:0;">'
                            f'{icon("assignment", size=15)} Ensayos asignados</div>', unsafe_allow_html=True)
            with h2:
                st.markdown(f'<div style="text-align:right;"><span class="badge badge-muted">Total: {len(pendientes)}</span></div>',
                            unsafe_allow_html=True)

            if not pendientes:
                st.info("No tienes ensayos pendientes por ahora.")
            else:
                col_ratios = [1.5, 2.6, 1.6, 1.8, 1.2, 0.9]
                headers = st.columns(col_ratios)
                for col, label in zip(headers, ["ID ensayo", "Proyecto", "Tipo de ensayo", "Última actualización", "Estado", "Acción"]):
                    col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
                for a in pendientes:
                    proyecto = get_project(a["codigo_interno"])
                    cols = st.columns(col_ratios, vertical_alignment="center")
                    ensayo_id = f'{a["codigo_interno"]}-{a["perforacion_codigo"]}-M{a["muestra_numero"]}'
                    cols[0].markdown(f'<span class="cell-id">{html.escape(ensayo_id)}</span>', unsafe_allow_html=True)
                    titulo = html.escape(proyecto["nombre"] if proyecto else a["codigo_interno"])
                    subtitulo = html.escape(proyecto.get("localizacion", "")) if proyecto else ""
                    cols[1].markdown(f'<div class="cell-title">{titulo}</div><div class="cell-sub">{subtitulo}</div>',
                                      unsafe_allow_html=True)
                    cols[2].markdown(f'<span class="assigned-chip">{ASSAY_LABELS[a["tipo"]]}</span>', unsafe_allow_html=True)
                    cols[3].markdown(f'<span class="cell-muted">{html.escape(format_dt(a["lastModified"]))}</span>',
                                      unsafe_allow_html=True)
                    cols[4].markdown(status_badge_html(a["status"]), unsafe_allow_html=True)
                    with cols[5]:
                        if st.button("Abrir", key=f"open_assigned_{a['id']}", use_container_width=True):
                            st.session_state.selected_assay_id = a["id"]
                            st.session_state.selected_codigo = a["codigo_interno"]
                            st.session_state.selected_perforacion = a["perforacion_codigo"]
                            st.session_state.selected_muestra_id = a["muestra_id"]
                            st.session_state.selected_assay_type = a["tipo"]
                            navigate("assay-form")


def _render_project_list(codes, empty_msg, allow_delete, mark_read_only=False):
    if not codes:
        st.info(empty_msg)
        return
    for p in st.session_state.projects:
        if p["codigo_interno"] not in codes:
            continue
        counts = project_progress(p["codigo_interno"])
        with st.container(border=True):
            cols = st.columns([3, 2, 1, 1] if allow_delete else [3, 2, 1])
            with cols[0]:
                st.markdown(f"**{p['codigo_interno']}**")
                st.caption(p["nombre"])
            with cols[1]:
                st.markdown(f'<span class="cell-muted">{icon(STATUS_ICON["sin-iniciar"], size=14)} {counts["sin-iniciar"]}'
                            f'&nbsp;&nbsp;·&nbsp;&nbsp;{icon(STATUS_ICON["en-proceso"], size=14)} {counts["en-proceso"]}'
                            f'&nbsp;&nbsp;·&nbsp;&nbsp;{icon(STATUS_ICON["finalizado"], size=14)} {counts["finalizado"]}</span>',
                            unsafe_allow_html=True)
            with cols[2]:
                if st.button("Abrir", key=f"openlist_{p['codigo_interno']}", use_container_width=True):
                    st.session_state.selected_codigo = p["codigo_interno"]
                    navigate("project-detail")
            if allow_delete:
                with cols[3]:
                    if confirm_delete(f"project_{p['codigo_interno']}", f"el proyecto {p['codigo_interno']}"):
                        codigo = p["codigo_interno"]
                        st.session_state.projects = [x for x in st.session_state.projects if x["codigo_interno"] != codigo]
                        st.session_state.perforaciones.pop(codigo, None)
                        st.session_state.muestras = {k: v for k, v in st.session_state.muestras.items() if not k.startswith(codigo + "::")}
                        st.session_state.assays = [a for a in st.session_state.assays if a["codigo_interno"] != codigo]
                        st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo + "::")}
                        st.rerun()


def _resumen_tecnico_perforaciones(codigo):
    """Una línea por perforación: sus muestras y los ensayos solicitados en ellas."""
    lineas = []
    for perf in st.session_state.perforaciones.get(codigo, []):
        muestras = st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        if not muestras:
            lineas.append(f"<strong>{perf['codigo']}</strong>: sin muestras")
            continue
        ids = ", ".join(f"M-{m['numero']}" for m in muestras)
        ensayos = sorted({e for m in muestras for e, activo in m["ensayos"].items() if activo})
        linea = f"<strong>{perf['codigo']}</strong>: {ids}"
        if ensayos:
            linea += f" · {', '.join(ensayos)}"
        lineas.append(linea)
    return lineas


def render_projects_active():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Proyectos en ejecución")
    st.caption("Monitoreo técnico de sondeos y análisis geotécnico.")

    proyectos = [p for p in st.session_state.projects if project_status(p["codigo_interno"]) == "ejecucion"]

    en_ensayo = sum(1 for p in proyectos if project_progress(p["codigo_interno"])["en-proceso"] > 0)
    c1, c2 = st.columns(2)
    for col, icono, label, valor in [
        (c1, "folder", "Activos", len(proyectos)), (c2, "science", "En ensayo", en_ensayo),
    ]:
        with col:
            st.markdown(f'<div class="stat-chip"><div class="stat-icon">{icon(icono, size=20)}</div>'
                        f'<div><div class="stat-label">{label}</div><div class="stat-value">{valor}</div></div></div>',
                        unsafe_allow_html=True)

    busqueda = st.text_input("Buscar", placeholder="Buscar por código o nombre...", label_visibility="collapsed", icon=":material/search:")
    if busqueda:
        q = busqueda.lower()
        proyectos = [p for p in proyectos if q in p["codigo_interno"].lower() or q in p["nombre"].lower()]

    if st.session_state.role == "jefe":
        with st.container(key="fab-new-project"):
            if st.button("", icon=":material/add:", key="fab_new_project_btn", help="Crear nuevo proyecto"):
                navigate("new-project")

    if not proyectos:
        st.info("No hay proyectos en ejecución en este momento.")
        return

    for p in proyectos:
        codigo = p["codigo_interno"]
        counts = project_progress(codigo)
        total = sum(counts.values())
        if counts["en-proceso"] > 0:
            estado_badge, estado_label = "badge-warning", "En ensayo"
        elif total == 0:
            estado_badge, estado_label = "badge-muted", "Sin muestras"
        else:
            estado_badge, estado_label = "badge-muted", "Por iniciar"

        with st.container(border=True, key=f"projcard_{codigo}"):
            top = st.columns([3, 1])
            top[0].markdown(f'<span class="code-badge">{html.escape(codigo)}</span>', unsafe_allow_html=True)
            top[1].markdown(f'<div style="text-align:right;"><span class="badge {estado_badge}">{estado_label}</span></div>',
                             unsafe_allow_html=True)
            st.markdown(f"**{p['nombre']}**")
            st.markdown(f'<span class="cell-muted">{icon("location_on", size=14)} {html.escape(p.get("localizacion") or "—")}</span>',
                        unsafe_allow_html=True)
            st.markdown(f'<div class="section-title" style="margin-bottom:4px;">Norma</div>'
                        f'<div style="margin-bottom:10px;">{html.escape(p.get("norma") or "—")}</div>', unsafe_allow_html=True)
            m1, m2 = st.columns(2)
            m1.markdown(f'<span class="cell-muted">Fecha bitácora:</span><br><span style="font-weight:600;">'
                        f'{html.escape(p.get("fecha_bitacora") or "—")}</span>', unsafe_allow_html=True)
            m2.markdown(f'<span class="cell-muted">Fecha ingreso:</span><br><span style="font-weight:600;">'
                        f'{html.escape(p.get("fecha_ingreso_muestra") or "—")}</span>', unsafe_allow_html=True)
            resumen = _resumen_tecnico_perforaciones(codigo)
            if resumen:
                st.markdown(f'<div class="section-title" style="margin-bottom:6px;">'
                            f'Resumen técnico de perforaciones ({len(resumen)})</div>', unsafe_allow_html=True)
                for linea in resumen:
                    st.markdown(f'<div class="cell-sub" style="margin-bottom:4px;">{linea}</div>', unsafe_allow_html=True)
            if st.button("Ver proyecto →", key=f"veractivo_{codigo}", type="primary", use_container_width=True):
                st.session_state.selected_codigo = codigo
                navigate("project-detail")


def render_projects_done():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Proyectos ejecutados")
    if st.session_state.role == "auxiliar":
        st.info("Modo consulta: puedes ver los resultados, pero no editarlos.")
    codes = [p["codigo_interno"] for p in st.session_state.projects if project_status(p["codigo_interno"]) == "ejecutado"]
    _render_project_list(codes, "Todavía no hay proyectos completamente finalizados.",
                          allow_delete=(st.session_state.role == "jefe"), mark_read_only=True)


# ════════════════════════════════════════════════════════════════════
# NUEVO PROYECTO (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def render_new_project():
    require_role("jefe")
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Nuevo proyecto")

    st.markdown('<div class="section-title">Código interno</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.text_input("Prefijo", value="GDA", disabled=True, autocomplete="off")
    with c2:
        numero = st.text_input("Número", placeholder="001", autocomplete="off")
    with c3:
        anio = st.text_input("Año", placeholder="24", autocomplete="off")

    codigo_interno = f"GDA-{numero}-{anio}" if numero and anio else ""
    existing_codes = [p["codigo_interno"] for p in st.session_state.projects]
    codigo_valido = bool(codigo_interno) and codigo_interno not in existing_codes
    if codigo_interno:
        if not codigo_valido:
            st.error(f"El código **{codigo_interno}** ya existe.")
        else:
            st.success(f"Código interno: **{codigo_interno}**")

    st.markdown('<div class="section-title">Información del proyecto</div>', unsafe_allow_html=True)
    nombre = st.text_input("Nombre del proyecto", placeholder="Estudio de suelos vía Bogotá-Medellín")
    localizacion = st.text_input("Localización", placeholder="Km 14+200")
    norma = st.radio("Norma", NORMA_PROYECTO_OPTIONS, horizontal=True)

    c1, c2 = st.columns(2)
    with c1:
        fecha_bitacora = st.date_input("Fecha de bitácora", value=date.today())
    with c2:
        fecha_ingreso = st.date_input("Fecha de ingreso de muestra", value=date.today())

    laboratorista_asignado = st.text_input(
        "Asignar bitácora a laboratorista (opcional)", placeholder="Nombre del laboratorista",
        help="Es solo una referencia informativa: como todos los auxiliares comparten la misma clave, "
             "esto no restringe quién puede ver o digitar el proyecto.",
    )

    # Perforaciones y muestras se arman aquí mismo, antes de crear el proyecto formalmente.
    # Se usa el código interno (ya calculado en vivo) como llave temporal en session_state.
    perforaciones = []
    edited_frames = {}
    if codigo_valido and nombre:
        perforaciones = st.session_state.perforaciones.setdefault(codigo_interno, [])

        st.markdown('<div class="section-title">Perforación</div>', unsafe_allow_html=True)
        pc1, pc2 = st.columns([2, 1])
        with pc1:
            tipo = st.selectbox("Tipo de perforación", list(TIPO_PERFORACION_PREFIX.keys()))
        with pc2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Nueva perforación", use_container_width=True, icon=":material/add:"):
                prefix = TIPO_PERFORACION_PREFIX[tipo]
                consecutivo = len([p for p in perforaciones if p["tipo"] == tipo]) + 1
                codigo_perf = f"{prefix}{consecutivo}"
                perforaciones.append({"tipo": tipo, "consecutivo": consecutivo, "codigo": codigo_perf})
                st.session_state.muestras[f"{codigo_interno}::{codigo_perf}"] = []
                st.rerun()

        if perforaciones:
            st.markdown('<div class="section-title">Perforaciones y muestras</div>', unsafe_allow_html=True)
        for perf in perforaciones:
            key = f"{codigo_interno}::{perf['codigo']}"
            muestras = st.session_state.muestras.setdefault(key, [])
            with st.expander(f"**{perf['codigo']}** — {perf['tipo']}  ·  {len(muestras)} muestra(s)", expanded=True):
                if key not in st.session_state.bitacora_draft:
                    df_init = pd.DataFrame(_muestras_to_rows(muestras))
                    for col in BITACORA_BASE_COLS:
                        if col not in df_init.columns:
                            df_init[col] = _bitacora_row_defaults()[col]
                    st.session_state.bitacora_draft[key] = df_init[BITACORA_BASE_COLS]
                df_source = st.session_state.bitacora_draft[key]

                column_config = {
                    "Número": st.column_config.TextColumn(default=""),
                    "Prof. De": st.column_config.NumberColumn(default=0.0, step=0.01),
                    "Prof. A": st.column_config.NumberColumn(default=0.0, step=0.01),
                    "Tipo de muestra": st.column_config.SelectboxColumn(options=TIPO_MUESTRA_OPTIONS, default=TIPO_MUESTRA_OPTIONS[0]),
                }
                for e in BITACORA_ENSAYOS:
                    column_config[e] = st.column_config.CheckboxColumn(e, default=False)
                column_config["Observaciones"] = st.column_config.TextColumn(
                    default="", width="medium", help="Cómo llegó la muestra o cualquier condición que impida el ensayo.")

                st.caption("Usa el ícono para agregar fila sobre la tabla para sumar una muestra nueva. Para eliminar una, selecciona el cuadro a la izquierda de su fila y usa el ícono de basura que aparece sobre la tabla.")
                edited = st.data_editor(
                    df_source, num_rows="dynamic", use_container_width=True,
                    column_config=column_config, key=f"neweditor_{key}",
                )
                edited_frames[key] = edited
                if confirm_delete(f"newperf_{key}", f"la perforación {perf['codigo']}"):
                    st.session_state.perforaciones[codigo_interno] = [p for p in perforaciones if p["codigo"] != perf["codigo"]]
                    st.session_state.muestras.pop(key, None)
                    st.session_state.bitacora_draft.pop(key, None)
                    st.rerun()
    elif nombre or numero or anio:
        st.info("Completa un código interno válido y el nombre del proyecto para agregar perforaciones y muestras.")

    if perforaciones:
        filas_preview, tipos_usados = [], set()
        for perf in perforaciones:
            tipos_usados.add(perf["tipo"])
            df_rows = edited_frames.get(f"{codigo_interno}::{perf['codigo']}")
            rows = df_rows.to_dict("records") if df_rows is not None else []
            for row in rows:
                numero_m = str(row.get("Número", "")).strip()
                if not numero_m or numero_m.lower() == "none" or numero_m == "nan":
                    continue
                filas_preview.append({
                    "perf_codigo": perf["codigo"], "numero": numero_m,
                    "tipo_muestra": row.get("Tipo de muestra") or TIPO_MUESTRA_OPTIONS[0],
                    "profundidad_de": row.get("Prof. De") or 0.0, "profundidad_hasta": row.get("Prof. A") or 0.0,
                    "ensayos": {e: bool(row.get(e, False)) for e in BITACORA_ENSAYOS},
                    "observaciones": row.get("Observaciones") or "",
                })
        project_preview = {
            "codigo_interno": codigo_interno, "numero": numero, "anio": anio, "nombre": nombre,
            "localizacion": localizacion, "norma": norma, "fecha_bitacora": str(fecha_bitacora),
        }
        excel_bytes, truncado = generar_excel_bitacora_orden(project_preview, filas_preview, tipos_usados)
        st.download_button("Descargar bitácora de orden (Excel)", data=excel_bytes, icon=":material/download:",
                            file_name=f"{codigo_interno} Bitacora de orden.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)
        if truncado:
            st.caption(f"El formato oficial admite hasta {BITACORA_XLSX_MAX_ROWS} muestras; se incluyeron las primeras {BITACORA_XLSX_MAX_ROWS}.")

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            if codigo_interno:
                st.session_state.perforaciones.pop(codigo_interno, None)
                st.session_state.muestras = {k: v for k, v in st.session_state.muestras.items() if not k.startswith(codigo_interno + "::")}
                st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo_interno + "::")}
            navigate("home")
    with col2:
        if st.button("Guardar bitácora", type="primary", use_container_width=True, icon=":material/save:",
                      disabled=not codigo_valido or not nombre):
            st.session_state.projects.append({
                "codigo_interno": codigo_interno, "numero": numero, "anio": anio, "nombre": nombre,
                "localizacion": localizacion, "norma": norma,
                "fecha_bitacora": str(fecha_bitacora), "fecha_ingreso_muestra": str(fecha_ingreso),
                "laboratorista_asignado": laboratorista_asignado,
            })
            st.session_state.perforaciones.setdefault(codigo_interno, [])
            for perf in perforaciones:
                key = f"{codigo_interno}::{perf['codigo']}"
                df_rows = edited_frames.get(key)
                rows = df_rows.to_dict("records") if df_rows is not None else []
                nuevas = []
                for row in rows:
                    numero_m = str(row.get("Número", "")).strip()
                    if not numero_m or numero_m.lower() == "none" or numero_m == "nan":
                        continue
                    id_unico = f"{codigo_interno}-{perf['codigo']}-M{numero_m}"
                    nuevas.append({
                        "numero": numero_m, "id_unico": id_unico,
                        "profundidad_de": row.get("Prof. De") or 0.0, "profundidad_hasta": row.get("Prof. A") or 0.0,
                        "tipo_muestra": row.get("Tipo de muestra") or TIPO_MUESTRA_OPTIONS[0],
                        "ensayos": {e: bool(row.get(e, False)) for e in BITACORA_ENSAYOS},
                        "observaciones": row.get("Observaciones") or "",
                    })
                st.session_state.muestras[key] = nuevas
                # El draft cacheado quedó vacío desde que se creó la perforación (antes de guardar);
                # se descarta para que la próxima vez que se abra la Bitácora se reconstruya a partir
                # de las muestras recién guardadas y no muestre/reescriba una tabla vacía.
                st.session_state.bitacora_draft.pop(key, None)
            st.session_state.selected_codigo = codigo_interno
            navigate("project-detail")


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PROYECTO → PERFORACIONES + PROGRESO
# ════════════════════════════════════════════════════════════════════
def render_project_detail():
    codigo = st.session_state.selected_codigo
    project = get_project(codigo)
    if not project:
        navigate("home")
        return

    if st.button("← Atrás"):
        go_back()

    progreso = project_progress(codigo)
    total = sum(progreso.values())
    pct_general = round(progreso["finalizado"] / total * 100) if total else 0
    perforaciones = st.session_state.perforaciones.get(codigo, [])
    sondeos_txt = f'{len(perforaciones)} sondeo(s) registrado(s)'

    st.markdown(f'''
        <div class="bento-primary" style="margin-bottom:16px;">
            <span class="bento-eyebrow">Proyecto ID</span>
            <h2 style="color:#fff;margin:4px 0 2px 0;">{html.escape(project["codigo_interno"])}</h2>
            <p style="opacity:0.85;font-size:15px;margin:0 0 12px 0;">{html.escape(project["nombre"])}</p>
            <span class="badge" style="background:rgba(255,255,255,0.15);color:#fff;">{sondeos_txt}</span>
        </div>
    ''', unsafe_allow_html=True)

    with st.container(border=True):
        info_rows = [
            ("location_on", "Ubicación", project.get("localizacion")),
            ("rule", "Norma", project.get("norma")),
            ("calendar_month", "Fecha de orden", project.get("fecha_bitacora")),
            ("move_to_inbox", "Ingreso de muestras", project.get("fecha_ingreso_muestra")),
            ("person", "Asignado a", project.get("laboratorista_asignado")),
        ]
        for i, (icono, label, valor) in enumerate(info_rows):
            margen = "margin-top:14px;" if i else ""
            st.markdown(f'<div class="cell-muted" style="{margen}text-transform:uppercase;letter-spacing:0.04em;font-size:11px;">'
                        f'{icon(icono, size=14)} {label}</div>'
                        f'<div style="font-weight:600;font-size:15px;">{html.escape(valor or "—")}</div>', unsafe_allow_html=True)

    if st.session_state.role == "jefe":
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Editar proyecto", icon=":material/edit:", use_container_width=True):
                navigate("edit-project")
        with c2:
            if confirm_delete(f"project_{codigo}", f"el proyecto {codigo} y todas sus perforaciones y muestras"):
                st.session_state.projects = [p for p in st.session_state.projects if p["codigo_interno"] != codigo]
                st.session_state.perforaciones.pop(codigo, None)
                st.session_state.muestras = {k: v for k, v in st.session_state.muestras.items() if not k.startswith(codigo + "::")}
                st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo + "::")}
                st.session_state.assays = [a for a in st.session_state.assays if a["codigo_interno"] != codigo]
                navigate("home")

    with st.container(border=True):
        st.markdown('<div class="section-title">Progreso general (así avanzan los auxiliares)</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 3])
        with c1:
            st.markdown(f'<div style="font-size:24px;font-weight:800;color:{PRIMARY};">{pct_general}%</div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="cell-muted" style="margin-top:8px;">{progreso["finalizado"]} de {total} muestras completadas</div>',
                        unsafe_allow_html=True)
        st.progress(pct_general / 100)
        cols = st.columns(3)
        for col, status_key, label in zip(cols, ["sin-iniciar", "en-proceso", "finalizado"],
                                           ["Sin iniciar", "En proceso", "Finalizado"]):
            with col:
                st.markdown(f'''
                    <div style="background:{SECONDARY_CONTAINER};border-radius:8px;
                                padding:6px 6px;text-align:center;margin-top:10px;margin-bottom:10px;">
                        <div style="font-size:9px;text-transform:uppercase;letter-spacing:0.03em;color:{PRIMARY};">{label}</div>
                        <div style="font-size:15px;font-weight:800;color:{PRIMARY};">{progreso[status_key]}</div>
                    </div>
                ''', unsafe_allow_html=True)

    if st.session_state.role == "jefe":
        tiene_bitacora = bool(perforaciones)
        label_bitacora = "Editar bitácora de orden" if tiene_bitacora else "Generar bitácora de orden"
        icon_bitacora = "edit_document" if tiene_bitacora else "assignment"
        c1, c2 = st.columns(2)
        with c1:
            if st.button(label_bitacora, type="primary", icon=f":material/{icon_bitacora}:", use_container_width=True):
                navigate("bitacora")
        with c2:
            filas_proy, tipos_usados = _bitacora_filas_proyecto(codigo, perforaciones)
            excel_bytes, truncado = generar_excel_bitacora_orden(project, filas_proy, tipos_usados)
            st.download_button("Descargar bitácora de orden", data=excel_bytes, icon=":material/download:",
                                file_name=f"{codigo} Bitacora de orden.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True)
        if truncado:
            st.caption(f"El formato oficial admite hasta {BITACORA_XLSX_MAX_ROWS} muestras; se incluyeron las primeras {BITACORA_XLSX_MAX_ROWS}.")

    st.markdown(f'<div class="section-title">Perforaciones realizadas ({len(perforaciones)} elemento(s) identificado(s))</div>',
                unsafe_allow_html=True)
    if not perforaciones:
        st.info("Este proyecto todavía no tiene perforaciones. Usa la Bitácora para agregarlas.")
    for perf in perforaciones:
        muestras = st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        counts = {"sin-iniciar": 0, "en-proceso": 0, "finalizado": 0}
        for m in muestras:
            counts[compute_muestra_estado(m)] += 1
        perf_total = len(muestras)
        perf_pct = round(counts["finalizado"] / perf_total * 100) if perf_total else 0
        if perf_total == 0:
            estado_badge, estado_label = "badge-muted", "Pendiente"
        elif counts["finalizado"] == perf_total:
            estado_badge, estado_label = "badge-success", "Completado"
        else:
            estado_badge, estado_label = "badge-warning", "En progreso"
        if muestras:
            prof_txt = f'{min(m["profundidad_de"] for m in muestras):.2f}m – {max(m["profundidad_hasta"] for m in muestras):.2f}m'
        else:
            prof_txt = "—"

        with st.container(border=True):
            top = st.columns([1, 3, 1.3])
            with top[0]:
                st.markdown(f'<div class="perf-code-box">{html.escape(perf["codigo"])}</div>', unsafe_allow_html=True)
            with top[1]:
                st.markdown(f'<span style="font-weight:700;">{perf_pct}%</span>&nbsp;&nbsp;'
                            f'<span class="badge {estado_badge}">{estado_label}</span>', unsafe_allow_html=True)
            with top[2]:
                st.markdown(f'<div style="text-align:right;"><span class="assigned-chip">{html.escape(perf["tipo"])}</span></div>',
                            unsafe_allow_html=True)
            st.markdown(f'<div class="cell-muted"><strong>Profundidad</strong> {prof_txt} · {len(muestras)} muestra(s)</div>',
                        unsafe_allow_html=True)
            st.progress(perf_pct / 100)
            if st.button("Ver muestras →", key=f"open_perf_{perf['codigo']}", use_container_width=True):
                st.session_state.selected_perforacion = perf["codigo"]
                navigate("perforacion-detail")


# ════════════════════════════════════════════════════════════════════
# EDITAR PROYECTO (el código interno no se puede editar: es la llave
# que usan perforaciones, muestras y ensayos en session_state)
# ════════════════════════════════════════════════════════════════════
def render_edit_project():
    require_role("jefe")
    codigo = st.session_state.selected_codigo
    project = get_project(codigo)
    if not project:
        navigate("home")
        return

    if st.button("← Atrás"):
        go_back(fallback="project-detail")
    st.markdown("## Editar proyecto")
    st.markdown(f'<span class="code-badge">{html.escape(codigo)}</span>', unsafe_allow_html=True)
    st.caption("El código interno no se puede modificar.")

    nombre = st.text_input("Nombre del proyecto", value=project.get("nombre", ""))
    localizacion = st.text_input("Localización", value=project.get("localizacion", ""))
    norma_actual = project.get("norma")
    idx = NORMA_PROYECTO_OPTIONS.index(norma_actual) if norma_actual in NORMA_PROYECTO_OPTIONS else 0
    norma = st.radio("Norma", NORMA_PROYECTO_OPTIONS, index=idx, horizontal=True)

    def _parse_fecha(valor):
        try:
            return date.fromisoformat(valor)
        except (TypeError, ValueError):
            return date.today()

    c1, c2 = st.columns(2)
    with c1:
        fecha_bitacora = st.date_input("Fecha de bitácora", value=_parse_fecha(project.get("fecha_bitacora")))
    with c2:
        fecha_ingreso = st.date_input("Fecha de ingreso de muestra", value=_parse_fecha(project.get("fecha_ingreso_muestra")))
    laboratorista_asignado = st.text_input(
        "Asignar bitácora a laboratorista (opcional)", value=project.get("laboratorista_asignado", ""))

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Cancelar", use_container_width=True):
            go_back(fallback="project-detail")
    with c2:
        if st.button("Guardar cambios", type="primary", use_container_width=True, icon=":material/save:", disabled=not nombre):
            project["nombre"] = nombre
            project["localizacion"] = localizacion
            project["norma"] = norma
            project["fecha_bitacora"] = str(fecha_bitacora)
            project["fecha_ingreso_muestra"] = str(fecha_ingreso)
            project["laboratorista_asignado"] = laboratorista_asignado
            navigate("project-detail")


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PERFORACIÓN → LISTA DE MUESTRAS
# ════════════════════════════════════════════════════════════════════
def _perforacion_ensayos_progress(codigo, perf_codigo):
    muestras = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    total_ensayos, completados = 0, 0
    for m in muestras:
        for label, activo in m["ensayos"].items():
            if not activo:
                continue
            total_ensayos += 1
            tipo_interno = SUPPORTED_ASSAY_MAP.get(label)
            a = get_assay(m["id_unico"], tipo_interno) if tipo_interno else None
            if a and a["status"] == "finalizado":
                completados += 1
    return completados, total_ensayos


def render_perforacion_detail():
    codigo = st.session_state.selected_codigo
    perf_codigo = st.session_state.selected_perforacion
    project = get_project(codigo)
    if not project:
        navigate("home")
        return

    if st.button("← Atrás"):
        go_back(fallback="project-detail")

    perf = next((p for p in st.session_state.perforaciones.get(codigo, []) if p["codigo"] == perf_codigo), None)
    muestras = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    ensayos_completados, ensayos_total = _perforacion_ensayos_progress(codigo, perf_codigo)
    pct = round(ensayos_completados / ensayos_total * 100) if ensayos_total else 0

    st.markdown(f'<div class="cell-muted" style="text-transform:uppercase;letter-spacing:0.04em;font-size:11px;margin-bottom:4px;">Proyecto</div>'
                f'<span class="code-badge">{html.escape(project["codigo_interno"])}</span> '
                f'<span style="font-weight:600;">{html.escape(project["nombre"])}</span>',
                unsafe_allow_html=True)
    st.markdown(f"### Muestras de Perforación {html.escape(perf_codigo)}")

    with st.container(border=True):
        top = st.columns([3, 1])
        with top[0]:
            st.markdown(f"#### Sondeo {html.escape(perf_codigo)}")
            st.markdown(f'<span class="assigned-chip">{html.escape(perf["tipo"] if perf else "—")}</span>',
                        unsafe_allow_html=True)
            st.markdown(f'<div class="cell-muted" style="margin-top:10px;">'
                        f'{icon("location_on", size=13)} {html.escape(project.get("localizacion") or "—")}'
                        f'&nbsp;&nbsp;&nbsp;{icon("calendar_month", size=13)} {html.escape(project.get("fecha_bitacora") or "—")}</div>',
                        unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            filas_perf = [{
                "perf_codigo": perf_codigo, "numero": m["numero"], "tipo_muestra": m["tipo_muestra"],
                "profundidad_de": m["profundidad_de"], "profundidad_hasta": m["profundidad_hasta"],
                "ensayos": m["ensayos"], "observaciones": m.get("observaciones", ""),
            } for m in muestras]
            excel_bytes, _truncado = generar_excel_bitacora_orden(project, filas_perf, {perf["tipo"]} if perf else set())
            st.download_button("Exportar perfil", data=excel_bytes, icon=":material/download:",
                                file_name=f"{codigo} Bitacora de orden {perf_codigo}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True)
        with c2:
            if st.session_state.role == "jefe":
                if st.button("Nueva muestra", icon=":material/add:", type="primary", use_container_width=True):
                    navigate("bitacora")

    with st.container(border=True):
        st.markdown('<div class="section-title">Avance del sondeo</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 3])
        with c1:
            st.markdown(f'<div style="font-size:32px;font-weight:800;color:{PRIMARY};">{pct}%</div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="cell-muted" style="margin-top:16px;">'
                        f'{icon("check_circle", size=14)} {ensayos_completados} de {ensayos_total} ensayos completados</div>',
                        unsafe_allow_html=True)
        st.progress(pct / 100)

    st.markdown(f'''
        <div style="background:{PRIMARY};color:#fff;border-radius:10px;padding:10px 16px;
                    display:flex;align-items:center;gap:8px;font-weight:700;font-size:14px;margin-bottom:10px;">
            {icon("science", size=16)} Orden de Laboratorio
        </div>
    ''', unsafe_allow_html=True)
    if not muestras:
        st.info("Esta perforación todavía no tiene muestras. Usa la Bitácora para agregarlas.")
    else:
        with st.container(border=True):
            col_ratios = [0.9, 1.1, 1.4, 2.5, 1.2, 1.1]
            headers = st.columns(col_ratios)
            for col, label in zip(headers, ["Muestra ID", "Tipo", "Profundidad", "Ensayos asignados", "Estado", "Acción"]):
                col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
            for i, m in enumerate(muestras):
                if i:
                    st.markdown('<hr style="margin:8px 0;border-color:#C4C6CF;">', unsafe_allow_html=True)
                cols = st.columns(col_ratios, vertical_alignment="center")
                cols[0].markdown(f'<span class="cell-id">M-{html.escape(str(m["numero"]))}</span>', unsafe_allow_html=True)
                cols[1].markdown(f'<span class="cell-muted">{html.escape(m["tipo_muestra"])}</span>', unsafe_allow_html=True)
                cols[2].markdown(f'<span class="cell-muted">{m["profundidad_de"]}–{m["profundidad_hasta"]} m</span>', unsafe_allow_html=True)
                ensayos_sol = [e for e, v in m["ensayos"].items() if v]
                chips = "".join(f'<span class="assigned-chip" style="margin-right:4px;">{html.escape(e)}</span>' for e in ensayos_sol) \
                    or '<span class="cell-muted">—</span>'
                cols[3].markdown(chips, unsafe_allow_html=True)
                cols[4].markdown(status_badge_html(compute_muestra_estado(m)), unsafe_allow_html=True)
                with cols[5]:
                    if st.button("Abrir", key=f"open_muestra_{m['id_unico']}", use_container_width=True):
                        st.session_state.selected_muestra_id = m["id_unico"]
                        navigate("muestra-detail")


# ════════════════════════════════════════════════════════════════════
# BITÁCORA — crea perforaciones y muestras
# ════════════════════════════════════════════════════════════════════
def _bitacora_row_defaults():
    row = {"Número": "", "Prof. De": 0.0, "Prof. A": 0.0, "Tipo de muestra": TIPO_MUESTRA_OPTIONS[0]}
    for e in BITACORA_ENSAYOS:
        row[e] = False
    row["Observaciones"] = ""
    return row


def _muestras_to_rows(muestras):
    rows = []
    for m in muestras:
        row = {"Número": m["numero"], "Prof. De": m["profundidad_de"], "Prof. A": m["profundidad_hasta"], "Tipo de muestra": m["tipo_muestra"]}
        for e in BITACORA_ENSAYOS:
            row[e] = m["ensayos"].get(e, False)
        row["Observaciones"] = m.get("observaciones", "")
        rows.append(row)
    return rows or [_bitacora_row_defaults()]


def render_bitacora():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Bitácora orden para ensayos de laboratorio")

    codes = [p["codigo_interno"] for p in st.session_state.projects]
    if not codes:
        st.info("Todavía no hay proyectos.")
        return
    default_idx = codes.index(st.session_state.selected_codigo) if st.session_state.selected_codigo in codes else 0
    codigo = st.selectbox("Proyecto", codes, index=default_idx)
    st.session_state.selected_codigo = codigo
    project = get_project(codigo)

    with st.container(border=True):
        st.markdown(f"**{project['nombre']}** · {project.get('localizacion','—')} · Norma {project.get('norma','—')}")

    perforaciones = st.session_state.perforaciones.setdefault(codigo, [])
    es_jefe = st.session_state.role == "jefe"

    if es_jefe:
        st.markdown('<div class="section-title">Agregar perforación</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([2, 1])
        with c1:
            tipo = st.selectbox("Tipo de perforación", list(TIPO_PERFORACION_PREFIX.keys()))
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Agregar perforación", use_container_width=True, icon=":material/add:"):
                prefix = TIPO_PERFORACION_PREFIX[tipo]
                consecutivo = len([p for p in perforaciones if p["tipo"] == tipo]) + 1
                codigo_perf = f"{prefix}{consecutivo}"
                perforaciones.append({"tipo": tipo, "consecutivo": consecutivo, "codigo": codigo_perf})
                st.session_state.muestras[f"{codigo}::{codigo_perf}"] = []
                st.rerun()
    else:
        st.info("Estás viendo la bitácora en modo lectura. Solo el Jefe puede editarla.")

    st.markdown('<div class="section-title">Perforaciones y muestras</div>', unsafe_allow_html=True)
    if not perforaciones:
        st.info("Todavía no hay perforaciones en este proyecto.")

    edited_frames = {}
    for perf in perforaciones:
        key = f"{codigo}::{perf['codigo']}"
        muestras = st.session_state.muestras.setdefault(key, [])

        with st.expander(f"**{perf['codigo']}** — {perf['tipo']}  ·  {len(muestras)} muestra(s)", expanded=True):
            # OJO: el DataFrame se crea UNA sola vez y se reutiliza el mismo objeto en cada rerun.
            # Reconstruirlo desde cero (dict -> DataFrame) en cada actualización es lo que causaba
            # que la primera edición se perdiera y tocara escribir dos veces.
            if key not in st.session_state.bitacora_draft:
                df_init = pd.DataFrame(_muestras_to_rows(muestras))
                for col in BITACORA_BASE_COLS:
                    if col not in df_init.columns:
                        df_init[col] = _bitacora_row_defaults()[col]
                st.session_state.bitacora_draft[key] = df_init[BITACORA_BASE_COLS]

            df_source = st.session_state.bitacora_draft[key]

            column_config = {
                "Número": st.column_config.TextColumn(default=""),
                # Sin `format`: con NumberColumn + format printf-style, Streamlit reformatea el valor
                # mostrado a mitad de la edición y descarta la primera pulsación, obligando a digitar
                # dos veces. Sin `format` el editor no interfiere y el valor se guarda a la primera.
                "Prof. De": st.column_config.NumberColumn(default=0.0, step=0.01),
                "Prof. A": st.column_config.NumberColumn(default=0.0, step=0.01),
                "Tipo de muestra": st.column_config.SelectboxColumn(options=TIPO_MUESTRA_OPTIONS, default=TIPO_MUESTRA_OPTIONS[0]),
            }
            for e in BITACORA_ENSAYOS:
                column_config[e] = st.column_config.CheckboxColumn(e, default=False)
            column_config["Observaciones"] = st.column_config.TextColumn(
                default="", width="medium", help="Cómo llegó la muestra o cualquier condición que impida el ensayo.")

            if es_jefe:
                st.caption("Usa el ícono para agregar fila sobre la tabla para sumar una muestra nueva. Para eliminar una, selecciona el cuadro a la izquierda de su fila y usa el ícono de basura que aparece sobre la tabla.")
                # OJO: `data` que se le pasa a st.data_editor debe permanecer estable entre reruns
                # (bitacora_draft[key] solo cambia por acciones explícitas nuestras, como "Agregar
                # muestra"). El resultado editado NO se vuelve a guardar ahí — hacerlo generaba el
                # bug de tener que digitar dos veces, porque el editor detectaba la fuente como
                # "cambiada" y descartaba la edición recién hecha.
                edited = st.data_editor(
                    df_source, num_rows="dynamic", use_container_width=True,
                    column_config=column_config, key=f"editor_{key}",
                )
                edited_frames[key] = edited
                if confirm_delete(f"perf_{key}", f"la perforación {perf['codigo']} y todas sus muestras"):
                    st.session_state.perforaciones[codigo] = [p for p in st.session_state.perforaciones[codigo] if p["codigo"] != perf["codigo"]]
                    st.session_state.muestras.pop(key, None)
                    st.session_state.bitacora_draft.pop(key, None)
                    st.session_state.assays = [a for a in st.session_state.assays if not (a["codigo_interno"] == codigo and a["perforacion_codigo"] == perf["codigo"])]
                    st.rerun()
            else:
                st.dataframe(df_source, use_container_width=True, hide_index=True)



    if es_jefe:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Guardar bitácora", type="primary", use_container_width=True, icon=":material/save:"):
            for perf in perforaciones:
                key = f"{codigo}::{perf['codigo']}"
                df_rows = edited_frames.get(key)
                rows = df_rows.to_dict("records") if df_rows is not None else []
                nuevas = []
                for row in rows:
                    numero = str(row.get("Número", "")).strip()
                    if not numero or numero.lower() == "none" or numero == "nan":
                        continue
                    id_unico = f"{codigo}-{perf['codigo']}-M{numero}"
                    nuevas.append({
                        "numero": numero, "id_unico": id_unico,
                        "profundidad_de": row.get("Prof. De") or 0.0, "profundidad_hasta": row.get("Prof. A") or 0.0,
                        "tipo_muestra": row.get("Tipo de muestra") or TIPO_MUESTRA_OPTIONS[0],
                        "ensayos": {e: bool(row.get(e, False)) for e in BITACORA_ENSAYOS},
                        "observaciones": row.get("Observaciones") or "",
                    })
                st.session_state.muestras[key] = nuevas
                # Se descarta el draft cacheado para que, si se vuelve a abrir esta perforación,
                # se reconstruya desde las muestras recién guardadas (evita mostrar/pisar con una
                # tabla vieja lo que ya se guardó).
                st.session_state.bitacora_draft.pop(key, None)
            st.success("Bitácora guardada. Los auxiliares ya pueden ver y digitar las muestras.")

    st.markdown("<br>", unsafe_allow_html=True)
    filas_proy, tipos_usados = _bitacora_filas_proyecto(codigo, perforaciones)
    excel_bytes, truncado = generar_excel_bitacora_orden(project, filas_proy, tipos_usados)
    st.download_button("Descargar bitácora de orden", data=excel_bytes, icon=":material/download:",
                        file_name=f"{codigo} Bitacora de orden.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
    if truncado:
        st.caption(f"El formato oficial admite hasta {BITACORA_XLSX_MAX_ROWS} muestras; se incluyeron las primeras {BITACORA_XLSX_MAX_ROWS}.")


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

    project = get_project(codigo)
    if st.button("← Atrás"):
        go_back(fallback="perforacion-detail")

    with st.container(border=True):
        top = st.columns([3, 1])
        with top[0]:
            st.markdown(f"### Muestra {muestra['numero']}")
            st.caption(f"{codigo} · {project['nombre'] if project else ''}")
        with top[1]:
            st.markdown(f'<div style="text-align:right;"><span class="assigned-chip">{html.escape(muestra["tipo_muestra"])}</span></div>',
                        unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.markdown(f'<div class="cell-muted">Identificador</div><div style="font-weight:600;">{html.escape(muestra["id_unico"])}</div>',
                    unsafe_allow_html=True)
        c2.markdown(f'<div class="cell-muted">Profundidad</div><div style="font-weight:600;">'
                    f'{muestra["profundidad_de"]}–{muestra["profundidad_hasta"]} m</div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="cell-muted">Perforación</div><div style="font-weight:600;">{html.escape(perf_codigo)}</div>',
                    unsafe_allow_html=True)
        estado = compute_muestra_estado(muestra)
        st.markdown(f'<div style="margin-top:10px;">{status_badge_html(estado, font_size=13)}</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<div class="section-title">Observaciones de la muestra</div>', unsafe_allow_html=True)
        st.caption("Cómo llegó la muestra, o cualquier condición que impida continuar con el ensayo. "
                   "Se guarda para todos los ensayos de esta muestra y la puede editar tanto el Jefe como el laboratorista.")
        observacion = st.text_area(
            "Observaciones de la muestra", value=muestra.get("observaciones", ""), label_visibility="collapsed",
            placeholder="Ej: Muestra con humedad visible, sin alteraciones aparentes...", key=f"obs_{muestra_id}",
        )
        if st.button("Guardar observación", icon=":material/save:", key=f"obs_save_{muestra_id}"):
            muestra["observaciones"] = observacion
            st.success("Observación guardada.")

    solicitados = [e for e, v in muestra["ensayos"].items() if v]
    finalizados = sum(
        1 for e in solicitados
        if SUPPORTED_ASSAY_MAP.get(e) and (get_assay(muestra_id, SUPPORTED_ASSAY_MAP[e]) or {}).get("status") == "finalizado"
    )
    total_sol = len(solicitados)
    pct_ensayos = round(finalizados / total_sol * 100) if total_sol else 0

    with st.container(border=True):
        st.markdown('<div class="section-title">Avance de ensayos</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 3])
        with c1:
            st.markdown(f'<div style="font-size:32px;font-weight:800;color:{PRIMARY};">{pct_ensayos}%</div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="cell-muted" style="margin-top:16px;">{finalizados} de {total_sol} ensayos programados</div>',
                        unsafe_allow_html=True)
        st.progress(pct_ensayos / 100)

    st.markdown('<div class="section-title">Ensayos asignados</div>', unsafe_allow_html=True)
    if not solicitados:
        st.info("Esta muestra no tiene ensayos marcados en la bitácora.")
    for ensayo_label in solicitados:
        with st.container(border=True):
            cols = st.columns([3, 1.6, 1])
            cols[0].markdown(f"**{ensayo_label}**")
            tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
            if tipo_interno:
                existing = get_assay(muestra_id, tipo_interno)
                status = existing["status"] if existing else "sin-iniciar"
                cols[1].markdown(status_badge_html(status), unsafe_allow_html=True)
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
                if existing and existing.get("laboratorist"):
                    st.markdown(f'<div class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(existing["lastModified"])} · {existing["laboratorist"]}</div>', unsafe_allow_html=True)
                elif existing:
                    st.markdown(f'<div class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(existing["lastModified"])}</div>', unsafe_allow_html=True)
            else:
                cols[1].markdown('<span class="badge badge-muted">Sin formulario aún</span>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE BITÁCORA DE ORDEN (plantilla oficial GDA-FL-003)
# ════════════════════════════════════════════════════════════════════
def _bitacora_filas_proyecto(codigo, perforaciones):
    """Junta todas las muestras de todas las perforaciones ya guardadas de un proyecto,
    en el formato que espera generar_excel_bitacora_orden."""
    filas, tipos_usados = [], set()
    for perf in perforaciones:
        tipos_usados.add(perf["tipo"])
        for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
            filas.append({
                "perf_codigo": perf["codigo"], "numero": m["numero"], "tipo_muestra": m["tipo_muestra"],
                "profundidad_de": m["profundidad_de"], "profundidad_hasta": m["profundidad_hasta"],
                "ensayos": m["ensayos"], "observaciones": m.get("observaciones", ""),
            })
    return filas, tipos_usados


def generar_excel_bitacora_orden(project, filas, tipos_usados):
    """filas: lista de dicts con perf_codigo, numero, tipo_muestra, profundidad_de,
    profundidad_hasta, ensayos (dict) y observaciones. Devuelve (bytes, truncado)."""
    wb = load_workbook(TEMPLATE_BITACORA_ORDEN)
    ws = wb["S1"]

    # La plantilla ya trae "GDA" impreso en su propio recuadro (AC8); solo se llena el
    # recuadro siguiente (AG8) con número-año, tal como se digita en Código interno.
    numero_proy = str(project.get("numero") or "").strip()
    anio_proy = str(project.get("anio") or "").strip()
    if numero_proy or anio_proy:
        ws["AG8"] = f"{numero_proy}-{anio_proy}"
    fecha_partes = str(project.get("fecha_bitacora") or "").split("-")  # "AAAA-MM-DD"
    if len(fecha_partes) == 3:
        anio, mes, dia = fecha_partes
        ws["E8"] = dia
        ws["F8"] = mes
        ws["H8"] = anio
    ws["F10"] = project.get("nombre") or ""
    ws["E12"] = project.get("localizacion") or ""

    norma_cell = BITACORA_XLSX_NORMA_CELL.get(project.get("norma"))
    if norma_cell:
        ws[norma_cell] = "X"
    for tipo in tipos_usados:
        tipo_cell = BITACORA_XLSX_TIPO_CELL.get(tipo)
        if tipo_cell:
            ws[tipo_cell] = "X"

    truncado = len(filas) > BITACORA_XLSX_MAX_ROWS
    for i, fila in enumerate(filas[:BITACORA_XLSX_MAX_ROWS]):
        r = 18 + i
        ws[f"A{r}"] = fila["perf_codigo"]
        ws[f"B{r}"] = fila["numero"]
        ws[f"C{r}"] = fila["tipo_muestra"]
        ws[f"D{r}"] = to_float(fila.get("profundidad_de"))
        ws[f"E{r}"] = to_float(fila.get("profundidad_hasta"))
        for label, activo in fila.get("ensayos", {}).items():
            col = BITACORA_XLSX_ENSAYO_COL.get(label)
            if activo and col:
                ws[f"{col}{r}"] = "X"
        ws[f"AH{r}"] = fila.get("observaciones") or ""

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue(), truncado


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
    seleccionados = data.get(f"{prefix}_equipos", [])
    seleccionados = [e for e in seleccionados if e in EQUIPO_LIST]  # por si cambia la lista
    data[f"{prefix}_equipos"] = st.multiselect(
        "Marca todo el equipo utilizado en este ensayo", EQUIPO_LIST, default=seleccionados,
        key=f"{prefix}_equipos_ms", placeholder="Selecciona uno o varios equipos…",
    )


def render_norma_selector(assay_type, data, key_prefix):
    st.markdown('<div class="section-title">Norma aplicada</div>', unsafe_allow_html=True)
    options = NORMAS_ENSAYO[assay_type]
    current = data.get(f"{key_prefix}_norma", "")
    idx = options.index(current) if current in options else 0
    choice = st.radio("Norma", options, index=idx, horizontal=True, key=f"norma_{key_prefix}", label_visibility="collapsed")
    data[f"{key_prefix}_norma"] = choice


def render_granulometria_form(data, assay_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — los cálculos y la clasificación USCS los hace el Excel, no la app.")
    st.markdown('<div class="section-title">Datos generales</div>', unsafe_allow_html=True)
    data["masa_inicial_seca"] = st.text_input("Masa inicial seca (g)", value=data.get("masa_inicial_seca", ""), placeholder="350.5")

    st.markdown('<div class="section-title">Pesos retenidos por tamiz (g)</div>', unsafe_allow_html=True)
    # Igual que en la Bitácora: la fuente que se le pasa a st.data_editor debe permanecer estable
    # entre reruns (si no, el editor descarta la primera edición y toca digitar dos veces). Por eso
    # se arma una sola vez por ensayo y se cachea en session_state.
    sieve_key = f"sieve_{assay_id}"
    if sieve_key not in st.session_state.sieve_draft:
        rows = [{"Tamiz": label, "Abertura (mm)": apert, "Retenido (g)": to_float(data.get(key), 0.0)} for key, label, apert, _cell in SIEVES]
        st.session_state.sieve_draft[sieve_key] = pd.DataFrame(rows)
    df_source = st.session_state.sieve_draft[sieve_key]
    edited = st.data_editor(
        df_source, hide_index=True, use_container_width=True, disabled=["Tamiz", "Abertura (mm)"],
        column_config={"Retenido (g)": st.column_config.NumberColumn(step=0.1, default=0.0)},
        key=f"gran_sieve_editor_{assay_id}",
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


def render_read_only_summary(tipo, data):
    if tipo == "granulometria":
        st.markdown(f"**Masa inicial seca (g):** {data.get('masa_inicial_seca') or '—'}")
        rows = [{"Tamiz": label, "Retenido (g)": data.get(key, "—")} for key, label, _apert, _cell in SIEVES]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        equipos, norma = data.get("gran_equipos", []), data.get("gran_norma", "—")
    elif tipo == "humedad":
        campos = [("hum_metodo", "Método"), ("hum_recipiente", "Recipiente No."), ("hum_tara", "Masa recipiente (g)"),
                  ("hum_suelo_humedo_tara", "M. suelo húmedo + recipiente (g)"), ("hum_suelo_seco_tara", "M. suelo seco + recipiente 74h (g)")]
        for key, label in campos:
            st.markdown(f"**{label}:** {data.get(key) or '—'}")
        equipos, norma = data.get("hum_equipos", []), data.get("hum_norma", "—")
    else:
        campos = [("mu_peso_aire", "Masa en el aire (g)"), ("mu_peso_aire_par", "Masa en el aire parafinado (g)"),
                  ("mu_peso_agua_par", "Masa en el agua parafinado (g)"), ("mu_temp_agua", "Temperatura del agua (°C)"),
                  ("mu_peso_parafina", "Masa de la parafina (g)"), ("mu_dens_parafina", "Densidad de la parafina (g/cm³)")]
        for key, label in campos:
            st.markdown(f"**{label}:** {data.get(key) or '—'}")
        equipos, norma = data.get("mu_equipos", []), data.get("mu_norma", "—")
    st.markdown(f"**Equipo utilizado:** {', '.join(equipos) if equipos else '—'}")
    st.markdown(f"**Norma aplicada:** {norma or '—'}")


def render_assay_form():
    assay_id = st.session_state.selected_assay_id
    assay = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    if not assay:
        navigate("muestra-detail")
        return

    codigo, perf_codigo, muestra_id = assay["codigo_interno"], assay["perforacion_codigo"], assay["muestra_id"]
    project = get_project(codigo)
    muestra = get_muestra(codigo, perf_codigo, muestra_id)
    read_only = st.session_state.role == "auxiliar" and project_status(codigo) == "ejecutado"

    if st.button("← Atrás"):
        go_back(fallback="muestra-detail")

    bar_cols = st.columns(5)
    bar_cols[0].markdown(f"**Proyecto**<br>{codigo}", unsafe_allow_html=True)
    bar_cols[1].markdown(f"**Muestra**<br>{muestra['numero'] if muestra else '—'}", unsafe_allow_html=True)
    bar_cols[2].markdown(f"**Ensayo**<br>{ASSAY_LABELS[assay['tipo']]}", unsafe_allow_html=True)
    bar_cols[3].markdown(f"**Perforación**<br>{perf_codigo}", unsafe_allow_html=True)
    bar_cols[4].markdown(f"**Estado**<br>{status_badge_html(assay['status'])}", unsafe_allow_html=True)
    st.markdown(f'<div class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(assay["lastModified"])}'
                + (f' · {html.escape(assay["laboratorist"])}' if assay.get("laboratorist") else "") + '</div>', unsafe_allow_html=True)

    st.markdown(f"## {ASSAY_LABELS[assay['tipo']]}")
    data = dict(assay.get("data", {}))

    if read_only:
        st.info("Este proyecto ya fue ejecutado. Estás en modo consulta — no puedes editar estos datos.")
        render_read_only_summary(assay["tipo"], data)
        st.markdown('<div class="section-title">Observaciones</div>', unsafe_allow_html=True)
        st.write(assay.get("observations") or "—")
        st.markdown('<div class="section-title">Laboratorista</div>', unsafe_allow_html=True)
        st.write(assay.get("laboratorist") or "—")
        return

    if assay["tipo"] == "granulometria":
        render_granulometria_form(data, assay_id)
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
        if st.button("Guardar borrador", use_container_width=True, icon=":material/save:"):
            assay.update(data=data, observations=observations, laboratorist=laboratorist, status="en-proceso", lastModified=now_iso())
            navigate("muestra-detail")
    with col2:
        if st.button("Finalizar ensayo", type="primary", use_container_width=True, icon=":material/check_circle:"):
            assay.update(data=data, observations=observations, laboratorist=laboratorist, status="finalizado", lastModified=now_iso())
            navigate("muestra-detail")

    if st.session_state.role == "jefe" and assay["tipo"] == "granulometria" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_granulometria(codigo, perf_codigo, muestra, project, data)
        st.download_button(
            "Descargar Excel (plantilla oficial de Granulometría)", icon=":material/download:",
            data=excel_bytes, file_name=f"Granulometria_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True,
        )


# ════════════════════════════════════════════════════════════════════
# CONTINUAR / BUSCAR
# ════════════════════════════════════════════════════════════════════
def render_continue():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Continuar ensayo")
    in_progress = [a for a in st.session_state.assays if a["status"] == "en-proceso"]
    if not in_progress:
        st.info("No hay ensayos en proceso.")
    for a in in_progress:
        with st.container(border=True):
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
                    st.session_state.read_only_view = False
                    navigate("assay-form")


def render_search():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Buscar ensayos")

    codes = [p["codigo_interno"] for p in st.session_state.projects]
    if not codes:
        st.info("Todavía no hay proyectos.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        default_idx = codes.index(st.session_state.selected_codigo) if st.session_state.selected_codigo in codes else 0
        codigo = st.selectbox("Proyecto", codes, index=default_idx)
    perforaciones = st.session_state.perforaciones.get(codigo, [])
    with c2:
        perf_options = ["(todas)"] + [p["codigo"] for p in perforaciones]
        perf_choice = st.selectbox("Perforación", perf_options)
    with c3:
        f_type = st.selectbox("Tipo de ensayo", ["(todos)"] + list(ASSAY_LABELS.values()))

    if not perforaciones:
        st.info("Este proyecto todavía no tiene perforaciones. Ve a la Bitácora para agregarlas.")
        return

    perfs_to_show = perforaciones if perf_choice == "(todas)" else [p for p in perforaciones if p["codigo"] == perf_choice]

    project = get_project(codigo)
    any_shown = False
    for perf in perfs_to_show:
        muestras = st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        for m in muestras:
            solicitados = [e for e, v in m["ensayos"].items() if v]
            if f_type != "(todos)":
                solicitados = [e for e in solicitados if ASSAY_LABELS.get(SUPPORTED_ASSAY_MAP.get(e), None) == f_type]
            if not solicitados:
                continue
            any_shown = True
            with st.container(border=True):
                st.markdown(f"**{m['id_unico']}**  ·  Prof. {m['profundidad_de']}–{m['profundidad_hasta']} m  ·  {m['tipo_muestra']}")
                for ensayo_label in solicitados:
                    cols = st.columns([2.2, 1.4, 1.3, 1.3])
                    cols[0].markdown(ensayo_label)
                    tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
                    if tipo_interno:
                        existing = get_assay(m["id_unico"], tipo_interno)
                        status = existing["status"] if existing else "sin-iniciar"
                        cols[1].markdown(status_badge_html(status), unsafe_allow_html=True)
                        cols[1].caption(format_dt(existing["lastModified"]) if existing else "—")
                        with cols[2]:
                            if st.button("Abrir", key=f"search_open_{m['id_unico']}_{tipo_interno}", use_container_width=True):
                                if existing:
                                    st.session_state.selected_assay_id = existing["id"]
                                else:
                                    new_id = f"a-{uuid.uuid4().hex[:8]}"
                                    st.session_state.assays.append({
                                        "id": new_id, "muestra_id": m["id_unico"], "tipo": tipo_interno, "status": "sin-iniciar",
                                        "data": {}, "observations": "", "laboratorist": "",
                                        "codigo_interno": codigo, "perforacion_codigo": perf["codigo"], "muestra_numero": m["numero"],
                                        "lastModified": now_iso(), "createdAt": now_iso(),
                                    })
                                    st.session_state.selected_assay_id = new_id
                                st.session_state.selected_codigo = codigo
                                st.session_state.selected_perforacion = perf["codigo"]
                                st.session_state.selected_muestra_id = m["id_unico"]
                                st.session_state.selected_assay_type = tipo_interno
                                navigate("assay-form")
                        with cols[3]:
                            if st.session_state.role == "jefe" and tipo_interno == "granulometria" and project:
                                excel_bytes = generar_excel_granulometria(codigo, perf["codigo"], m, project, existing.get("data", {}) if existing else {})
                                st.download_button("Excel", icon=":material/download:", data=excel_bytes, file_name=f"Granulometria_{m['id_unico']}.xlsm",
                                                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                                                    key=f"search_dl_{m['id_unico']}", use_container_width=True)
                    else:
                        cols[1].markdown('<span class="badge badge-muted">Sin formulario aún</span>', unsafe_allow_html=True)

    if not any_shown:
        st.info("No se encontraron ensayos con esos filtros.")


# ════════════════════════════════════════════════════════════════════
# ENRUTADOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════
if st.session_state.role is None:
    render_login()
else:
    render_topbar()
    SCREENS = {
        "home": render_home, "new-project": render_new_project, "project-detail": render_project_detail,
        "edit-project": render_edit_project,
        "perforacion-detail": render_perforacion_detail, "muestra-detail": render_muestra_detail,
        "bitacora": render_bitacora, "assay-form": render_assay_form,
        "continue": render_continue, "search": render_search,
        "projects-active": render_projects_active, "projects-done": render_projects_done,
    }
    SCREENS.get(st.session_state.screen, render_home)()
    render_bottomnav()
