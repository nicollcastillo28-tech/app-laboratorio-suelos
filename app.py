"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Estructura: Proyecto -> Perforación (Sondeo/Apique/Fuente-Cantera) -> Muestra -> Ensayo

Cómo correrla en tu computador:
    streamlit run app.py
"""

import html
import os
import re
import uuid
import zipfile
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
TEMPLATE_HUMEDAD = os.path.join(BASE_DIR, "templates", "GDA-FLC-014_humedad_natural.xlsx")
TEMPLATE_MASA_UNITARIA = os.path.join(BASE_DIR, "templates", "GDA-FLC-004_masa_unitaria.xlsx")

PASSWORDS = {"jefe": "geodelta2024", "auxiliar": "aux2024", "ingeniero": "ing2024"}
ROLE_LABELS = {"jefe": "Jefe de Laboratorio", "auxiliar": "Auxiliar", "ingeniero": "Ingeniero"}
ROLE_INICIALES = {"jefe": "JL", "auxiliar": "AX", "ingeniero": "ING"}

# ════════════════════════════════════════════════════════════════════
# ESTILOS — paleta del brief SoilLab Pro (Primary #1B365D · Secondary #4A6278 · Tertiary #005EB8 · Neutral #64748B)
# ════════════════════════════════════════════════════════════════════
PRIMARY, PRIMARY_DARK, PRIMARY_CONTAINER = "#002046", "#001B3D", "#1B365D"
SECONDARY, SECONDARY_CONTAINER = "#496177", "#C9E2FD"
TERTIARY = "#005EB8"
NEUTRAL = "#64748B"
SUCCESS, SUCCESS_LIGHT = "#16A34A", "#DCFCE7"
WARNING, WARNING_LIGHT = "#D97706", "#FEF3C7"
DANGER, DANGER_LIGHT = "#DC2626", "#FEE2E2"
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
    .badge-danger {{ background: {DANGER_LIGHT}; color: {DANGER}; }}
    .badge-muted {{ background: #EEF1F5; color: {MUTED}; }}
    .status-circle {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 36px; height: 36px; border-radius: 999px; flex-shrink: 0;
    }}
    .status-circle-success {{ background: {SUCCESS_LIGHT}; color: {SUCCESS}; }}
    .status-circle-warning {{ background: {WARNING_LIGHT}; color: {WARNING}; }}
    .status-circle-danger {{ background: {DANGER_LIGHT}; color: {DANGER}; }}
    .status-circle-muted {{ background: #EEF1F5; color: {MUTED}; }}

    /* Tarjeta con acento a la izquierda, para encabezados de detalle (ej. Detalle de Muestra) */
    .st-key-muestra-header-card {{ border-left: 4px solid {PRIMARY} !important; }}
    .st-key-muestra-obs-box .stTextArea textarea {{ background-color: {SECONDARY_CONTAINER} !important; }}
    /* El laboratorista puede escribir en minúscula — se ve en mayúscula de una vez (el texto
       que se guarda también se normaliza a mayúscula al hacer clic en "Guardar"). */
    .st-key-muestra-desc-visual-box .stTextArea textarea {{ text-transform: uppercase; }}
    /* Botones dentro de la campana de notificaciones — por defecto salen blancos/planos y se
       pierden contra el fondo del popover. */
    .st-key-notif-popover-body .stButton button {{
        background-color: {SECONDARY_CONTAINER}; color: {PRIMARY}; border-color: {SECONDARY_CONTAINER};
    }}
    .st-key-notif-popover-body .stButton button:hover {{
        background-color: {PRIMARY}; color: {SURFACE}; border-color: {PRIMARY};
    }}
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

ASSAY_LABELS = {"granulometria": "Granulometría", "humedad": "Contenido de humedad", "masa-unitaria": "Peso unitario", "limites": "Límites de Atterberg", "pasa200": "Pasa 200"}
NORMAS_ENSAYO = {
    "granulometria": ["INV-214-13", "INV.E-213-13", "INV.E 123-13"],
    "humedad": ["INV E-122", "ASTM D2216"],
    "masa-unitaria": ["INV E-202", "ASTM D1188"],
}
STATUS_LABELS = {"sin-iniciar": "Sin iniciar", "en-proceso": "En proceso", "finalizado": "Finalizado"}
STATUS_BADGE = {"sin-iniciar": "badge-success", "en-proceso": "badge-warning", "finalizado": "badge-danger"}
STATUS_ICON = {"sin-iniciar": "radio_button_unchecked", "en-proceso": "autorenew", "finalizado": "check_circle"}

TIPO_PERFORACION_PREFIX = {"Sondeo": "S", "Apique": "AP", "Fuente/Cantera": "F"}
# Texto que espera la lista desplegable de "tipo de perforación" en la plantilla
# CLASIFICACION_DE_SUELOS.xlsm (celda D12, validada contra AG6:AG10: SONDEO/APIQUE/TRINCHERA/NQ/N.A.).
TIPO_PERFORACION_EXCEL = {"Sondeo": "SONDEO", "Apique": "APIQUE", "Fuente/Cantera": "CANTERA"}
TIPO_MUESTRA_OPTIONS = ["Shelby", "NQ", "SS", "N/A"]
NORMA_PROYECTO_OPTIONS = ["IDU", "NTC", "INVIAS", "Otro"]

# Lista de equipos del laboratorio. Por ahora sin código — agrega o edita los que tengas aquí.
EQUIPO_LIST = [
    "Balanza digital 0.01g", "Balanza digital 0.1g", "Horno de secado", "Tamices serie gruesa",
    "Tamices serie fina", "Tamizadora mecánica", "Cazuela de Casagrande", "Ranurador", "Copa de Casagrande",
    "Molde Proctor estándar", "Molde Proctor modificado", "Prensa CBR", "Balanza hidrostática",
    "Horno de parafinado", "Cronómetro", "Termómetro", "Extractor de muestras", "Otro",
]

# Equipos reales de laboratorio con su código interno, tal como aparecen en el formato físico
# "EQUIPOS UTILIZADOS" para el ensayo de Granulometría (incluye el lavado por Tamiz No. 200).
EQUIPO_GRANULOMETRIA = [
    "Balanza GDA-E-010", "Balanza GDA-E-011", "Balanza GDA-E-012",
    "Horno GDA-E-007", "Horno GDA-E-404",
    "Tamiz de lavado GDA-E-", "Serie de tamices GDA-E-030 a GDA-E-045",
]

# Equipos reales usados en el ensayo de Contenido de Humedad Natural.
EQUIPO_HUMEDAD = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Balanza GDA-E-014", "Horno GDA-E-007", "Horno GDA-E-404"]

# Método del ensayo de humedad (INV E-122), tal como aparece en la plantilla oficial (celda C28).
METODO_HUMEDAD = ["Método A", "Método B"]

# Equipos reales usados en el ensayo de Límites de Atterberg (Líquido + Plástico).
EQUIPO_LIMITES = [
    "Balanza GDA-E-012", "Balanza GDA-E-011", "Balanza GDA-E-010",
    "Cazuela Casagrande GDA-E-081", "Cazuela Casagrande GDA-E-060", "Cazuela Casagrande GDA-E-400",
    "Horno GDA-E-404", "Horno GDA-E-007", "Tamiz No. 40 GDA-E-054",
]

# Equipos reales usados en el ensayo de Peso Unitario Parafinado.
EQUIPO_MASA_UNITARIA = ["Balanza GDA-E-011", "Termómetro GDA-E-126"]

# Filas de Límite Líquido (INV. E-125-13) y Límite Plástico (INV. E-126-13), con las celdas
# reales de la plantilla CLASIFICACION_DE_SUELOS.xlsm (sección "LIMITES DE ATTERBERG", a la
# derecha de la tabla de Granulometría en la hoja "MUESTRA"): 3 columnas de ensayo para el
# Líquido (S/T/U) y 2 para el Plástico (Y/Z) — el propio Excel calcula LL, LP e IP.
LIMITE_LIQUIDO_FILAS = [
    ("lim_ll_recipiente", "Recipiente No.", ["S20", "T20", "U20"]),
    ("lim_ll_golpes", "No. de Golpes", ["S19", "T19", "U19"]),
    ("lim_ll_humedo", "Masa suelo húmedo + rec. (g)", ["S21", "T21", "U21"]),
    ("lim_ll_seco", "Masa suelo seco + rec. (g)", ["S22", "T22", "U22"]),
    # Las lecturas a 14/15/16 horas son solo un dato de apoyo/verificación del laboratorista —
    # no tienen celda propia en el Excel (la plantilla solo usa "lim_ll_seco").
    ("lim_ll_seco_14h", "Masa suelo seco + rec. (g) (14 hrs)", []),
    ("lim_ll_seco_15h", "Masa suelo seco + rec. (g) (15 hrs)", []),
    ("lim_ll_seco_16h", "Masa suelo seco + rec. (g) (16 hrs)", []),
    ("lim_ll_recip_masa", "Masa recipiente (g)", ["S23", "T23", "U23"]),
]
LIMITE_LIQUIDO_N = 3

LIMITE_PLASTICO_FILAS = [
    ("lim_lp_recipiente", "Recipiente No.", ["Y20", "Z20"]),
    ("lim_lp_humedo", "Masa suelo húmedo + rec. (g)", ["Y21", "Z21"]),
    ("lim_lp_seco", "Masa suelo seco + rec. (g)", ["Y22", "Z22"]),
    ("lim_lp_seco_14h", "Masa suelo seco + rec. (g) (14 hrs)", []),
    ("lim_lp_seco_15h", "Masa suelo seco + rec. (g) (15 hrs)", []),
    ("lim_lp_seco_16h", "Masa suelo seco + rec. (g) (16 hrs)", []),
    ("lim_lp_recip_masa", "Masa recipiente (g)", ["Y23", "Z23"]),
]
LIMITE_PLASTICO_N = 2

# "Pasa 200" sí aparece como ensayo aparte (se puede solicitar sin Granulometría), pero
# comparte plantilla y datos con Granulometría: si ambos se solicitan para la misma muestra,
# los dos leen y escriben el mismo diccionario de datos (ver render_assay_form), así que lo que
# se digite en cualquiera de los dos se refleja igual en ambos.
BITACORA_ENSAYOS = [
    "Granulometría", "Pasa 200", "Humedad", "Límites de Atterberg", "Límite de contracción",
    "Materia orgánica", "Proctor", "CBR", "Compresión inconfinada", "Compresión en roca",
    "Peso unitario", "Gravedad específica", "Consolidación", "Corte CD", "Corte CU", "Corte UU", "Otro",
]
SUPPORTED_ASSAY_MAP = {
    "Granulometría": "granulometria", "Humedad": "humedad", "Peso unitario": "masa-unitaria",
    "Límites de Atterberg": "limites", "Pasa 200": "pasa200",
}

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
# ALMACÉN COMPARTIDO ENTRE SESIONES (jefe y auxiliares deben ver los mismos datos
# aunque estén en pestañas/dispositivos distintos — st.session_state por sí solo es
# privado de cada sesión de navegador, así que los datos "de negocio" viven aquí,
# en un recurso cacheado que vive mientras el proceso del servidor siga corriendo.
# Nota: esto NO sobrevive un reinicio del servidor — eso es tarea de la migración a
# Supabase, que queda pendiente aparte).
# ════════════════════════════════════════════════════════════════════
@st.cache_resource
def get_shared_store():
    codigo_demo = "GDA-001-24"
    return {
        "projects": [{
            "codigo_interno": codigo_demo, "numero": "001", "anio": "24",
            "nombre": "Estudio de suelos vía Bogotá-Medellín Km 14", "localizacion": "Sector Norte, Km 14+200",
            "fecha_bitacora": "2024-11-15", "fecha_ingreso_muestra": "2024-11-15", "norma": "GDA",
        }],
        "perforaciones": {codigo_demo: [{"tipo": "Sondeo", "consecutivo": 1, "codigo": "S1"}]},
        "muestras": {
            f"{codigo_demo}::S1": [{
                "numero": "1", "id_unico": f"{codigo_demo}-S1-M1", "profundidad_de": 0.0, "profundidad_hasta": 1.5,
                "tipo_muestra": "Shelby", "ensayos": {"Granulometría": True, "Humedad": True}, "observaciones": "",
            }]
        },
        "assays": [{
            "id": "a001", "muestra_id": f"{codigo_demo}-S1-M1", "tipo": "granulometria", "status": "en-proceso",
            "data": {}, "observations": "", "laboratorist": "",
            "codigo_interno": codigo_demo, "perforacion_codigo": "S1", "muestra_numero": "1",
            "lastModified": datetime.now().isoformat(), "createdAt": datetime.now().isoformat(),
        }],
        "notifications": [],
    }


# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL
# ════════════════════════════════════════════════════════════════════
def _sync_query_params():
    """Guarda el rol y la pantalla actual en la URL. Si la conexión se corta (celular que se
    bloquea, wifi inestable) y el navegador reconecta, Streamlit abre una sesión nueva pero
    puede restaurar dónde estaba el usuario leyendo esto — sin pedirle la clave otra vez ni
    mandarlo al Inicio en medio de un ensayo. Al cerrar sesión se limpia todo, así que la
    sesión solo termina cuando el usuario le da a "Cerrar sesión"."""
    if st.session_state.role:
        st.query_params["role"] = st.session_state.role
        st.query_params["screen"] = st.session_state.screen
        st.query_params["codigo"] = st.session_state.selected_codigo or ""
        st.query_params["perf"] = st.session_state.selected_perforacion or ""
        st.query_params["muestra"] = st.session_state.selected_muestra_id or ""
        st.query_params["assay"] = st.session_state.selected_assay_id or ""
        st.query_params["atipo"] = st.session_state.selected_assay_type or ""
    else:
        st.query_params.clear()


def init_state():
    if "initialized" in st.session_state:
        return
    st.session_state.initialized = True
    st.session_state.role = None
    st.session_state.screen = "home"

    store = get_shared_store()
    st.session_state.projects = store["projects"]
    st.session_state.perforaciones = store["perforaciones"]
    st.session_state.muestras = store["muestras"]
    st.session_state.assays = store["assays"]
    st.session_state.notifications = store["notifications"]

    st.session_state.nav_stack = []
    st.session_state.bitacora_draft = {}
    st.session_state.selected_codigo = ""
    st.session_state.selected_perforacion = ""
    st.session_state.selected_muestra_id = ""
    st.session_state.selected_assay_id = None
    st.session_state.selected_assay_type = None
    st.session_state.read_only_view = False

    # Restaura la sesión desde la URL si venía de una reconexión (ver _sync_query_params).
    qp_role = st.query_params.get("role")
    if qp_role in PASSWORDS:
        st.session_state.role = qp_role
        st.session_state.screen = st.query_params.get("screen") or "home"
        st.session_state.selected_codigo = st.query_params.get("codigo") or ""
        st.session_state.selected_perforacion = st.query_params.get("perf") or ""
        st.session_state.selected_muestra_id = st.query_params.get("muestra") or ""
        st.session_state.selected_assay_id = st.query_params.get("assay") or None
        st.session_state.selected_assay_type = st.query_params.get("atipo") or None


init_state()


def navigate(screen):
    actual = st.session_state.get("screen")
    if actual and actual != screen:
        st.session_state.nav_stack.append(actual)
    st.session_state.screen = screen
    _sync_query_params()
    st.rerun()


def go_back(fallback="home"):
    """Vuelve a la pantalla realmente anterior (pila de navegación) en vez de un destino fijo."""
    if st.session_state.nav_stack:
        st.session_state.screen = st.session_state.nav_stack.pop()
    else:
        st.session_state.screen = fallback
    _sync_query_params()
    st.rerun()


def to_float(v, default=None):
    try:
        return float(str(v).replace(",", "."))
    except (ValueError, TypeError):
        return default


def fmt_num(v, decimals=3):
    """Formatea un derivado numérico permitiendo hasta `decimals` decimales, sin ceros de más."""
    if v is None:
        return None
    s = f"{v:.{decimals}f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-") else "0"


def calcular_humedad_pct(data):
    """% de humedad a partir de los datos digitados en el ensayo de Humedad, con la misma
    fórmula que trae la plantilla oficial (GDA-FLC-014, celda I24):
    (masa húmeda - masa seca) / (masa seca - masa recipiente) * 100."""
    masa_humedo = to_float(data.get("hum_masa_humedo_mas_recipiente"))
    masa_seco = to_float(data.get("hum_seco_mas_recipiente"))
    masa_recip = to_float(data.get("hum_masa_recipiente"))
    if masa_humedo is None or masa_seco is None or masa_recip is None:
        return None
    masa_suelo_seco = masa_seco - masa_recip
    if not masa_suelo_seco:
        return None
    return (masa_humedo - masa_seco) / masa_suelo_seco * 100


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


def status_circle_html(status, size=20):
    circle_class = STATUS_BADGE[status].replace("badge-", "status-circle-")
    return f'<span class="status-circle {circle_class}">{icon(STATUS_ICON[status], size=size, fill=True)}</span>'


def card_header_html(icon_name, title, extra_html=""):
    """Encabezado de tarjeta con ícono + título (y opcionalmente un badge a la derecha),
    usado en las tarjetas de los formularios de ensayo (Norma, Equipos, Pasa 200, etc.)."""
    return (f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">'
            f'<div style="display:flex;align-items:center;gap:8px;font-weight:700;color:{PRIMARY};font-size:15px;">'
            f'{icon(icon_name, size=18)} {title}</div>{extra_html}</div>')


def param_table_html(rows, header_left="PARÁMETRO", header_right="VALOR REGISTRADO"):
    """Tabla de 2 columnas (etiqueta/valor) para la vista de solo lectura ('Resultados de Ensayo')."""
    body = "".join(
        f'<tr><td style="padding:10px 14px;border-bottom:1px solid {BORDER};color:{TEXT};">{html.escape(str(label))}</td>'
        f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;font-weight:600;color:{PRIMARY};">'
        f'{html.escape(str(value)) if value not in (None, "") else "—"}</td></tr>'
        for label, value in rows
    )
    return (f'<table style="width:100%;border-collapse:collapse;font-size:14px;">'
            f'<thead><tr style="background:{SECONDARY_CONTAINER};">'
            f'<th style="padding:10px 14px;text-align:left;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{header_left}</th>'
            f'<th style="padding:10px 14px;text-align:right;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{header_right}</th>'
            f'</tr></thead><tbody>{body}</tbody></table>')


def param_table_3col_html(rows, headers=("PARÁMETRO", "ANTES DEL LAVADO (g)", "DESPUÉS DEL LAVADO (g)")):
    """Tabla de 3 columnas (etiqueta + dos valores), usada para el Pasa No. 200 de Granulometría
    en la vista de solo lectura ('Resultados de Ensayo')."""
    def celda(v):
        return html.escape(str(v)) if v not in (None, "") else "—"
    body = "".join(
        f'<tr><td style="padding:10px 14px;border-bottom:1px solid {BORDER};color:{TEXT};">{html.escape(str(label))}</td>'
        f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;font-weight:600;color:{PRIMARY};">{celda(v1)}</td>'
        f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;font-weight:600;color:{PRIMARY};">{celda(v2)}</td></tr>'
        for label, v1, v2 in rows
    )
    head_left, head_mid, head_right = headers
    return (f'<table style="width:100%;border-collapse:collapse;font-size:14px;">'
            f'<thead><tr style="background:{SECONDARY_CONTAINER};">'
            f'<th style="padding:10px 14px;text-align:left;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{head_left}</th>'
            f'<th style="padding:10px 14px;text-align:center;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{head_mid}</th>'
            f'<th style="padding:10px 14px;text-align:center;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{head_right}</th>'
            f'</tr></thead><tbody>{body}</tbody></table>')


def param_table_ncol_html(headers, rows):
    """Tabla con cantidad arbitraria de columnas de valores (primer elemento de cada fila es la
    etiqueta, el resto son valores) — usada para Límite Líquido/Plástico, que tienen 3 y 2
    columnas de ensayo respectivamente."""
    def celda(v):
        return html.escape(str(v)) if v not in (None, "") else "—"
    body = "".join(
        '<tr>' + f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};color:{TEXT};">{html.escape(str(row[0]))}</td>'
        + "".join(f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;font-weight:600;color:{PRIMARY};">{celda(v)}</td>' for v in row[1:])
        + '</tr>'
        for row in rows
    )
    head_cells = "".join(
        f'<th style="padding:10px 14px;text-align:{"left" if i == 0 else "center"};font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">{h}</th>'
        for i, h in enumerate(headers)
    )
    return (f'<table style="width:100%;border-collapse:collapse;font-size:14px;">'
            f'<thead><tr style="background:{SECONDARY_CONTAINER};">{head_cells}</tr></thead><tbody>{body}</tbody></table>')


def condicion_table_html(muestra):
    """Tabla 'CONDICIÓN / TEMPERATURA °C / HUMEDAD %' para la vista de solo lectura."""
    def fila(cond_key, label):
        temp = muestra.get(f"cond_{cond_key}_temp") or "—"
        hum = muestra.get(f"cond_{cond_key}_hum") or "—"
        return (f'<tr><td style="padding:10px 14px;border-bottom:1px solid {BORDER};font-weight:600;color:{TEXT};">{label}</td>'
                f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;">{html.escape(str(temp))}</td>'
                f'<td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;">{html.escape(str(hum))}</td></tr>')
    body = fila("inicial", "Inicial") + fila("final", "Final")
    return (f'<table style="width:100%;border-collapse:collapse;font-size:14px;">'
            f'<thead><tr style="background:{SECONDARY_CONTAINER};">'
            f'<th style="padding:10px 14px;text-align:left;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">CONDICIÓN</th>'
            f'<th style="padding:10px 14px;text-align:center;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">TEMPERATURA °C</th>'
            f'<th style="padding:10px 14px;text-align:center;font-size:11px;letter-spacing:0.04em;color:{PRIMARY};">HUMEDAD %</th>'
            f'</tr></thead><tbody>{body}</tbody></table>')


def split_equipo_codigo(equipo):
    """Separa 'Balanza GDA-E-011' en ('Balanza', 'GDA-E-011') para mostrarlo en dos líneas."""
    idx = equipo.find("GDA-E")
    if idx == -1:
        return equipo, ""
    return equipo[:idx].strip(), equipo[idx:].strip()


def equipos_readonly_html(equipos):
    """Lista de equipos utilizados (ícono + nombre + código) para la vista de solo lectura."""
    if not equipos:
        return f'<div class="cell-muted">Ningún equipo seleccionado.</div>'
    items = "".join(
        f'<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid {BORDER};">'
        f'{icon("construction", size=18, color=PRIMARY)}'
        f'<div><div style="font-weight:600;">{html.escape(nombre)}</div>'
        f'<div class="cell-muted" style="font-size:12px;">{html.escape(codigo) if codigo else "—"}</div></div></div>'
        for nombre, codigo in (split_equipo_codigo(e) for e in equipos)
    )
    return f'<div>{items}</div>'


def now_iso():
    return datetime.now().isoformat()


def add_notification(role, mensaje, codigo=None, perf=None, muestra_id=None):
    """Notificación compartida entre todas las sesiones logueadas con `role` (ver
    get_shared_store: notifications vive en el mismo store con @st.cache_resource)."""
    st.session_state.notifications.append({
        "id": f"n-{uuid.uuid4().hex[:8]}", "role": role, "mensaje": mensaje, "leida": False,
        "fecha": now_iso(), "codigo_interno": codigo, "perforacion_codigo": perf, "muestra_id": muestra_id,
    })


def add_historial_muestra(muestra, texto):
    """Registro de auditoría por muestra: cuándo se entregó cada ensayo, cuándo confirmó el
    Jefe, cuándo aprobó (o devolvió) el Ingeniero, etc. Se guarda en la propia muestra."""
    muestra.setdefault("historial", []).append({"fecha": now_iso(), "texto": texto})


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


def get_perforacion(codigo, perforacion_codigo):
    return next((p for p in st.session_state.perforaciones.get(codigo, []) if p["codigo"] == perforacion_codigo), None)


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
    """'ejecutado' solo si el proyecto tiene al menos una muestra, TODAS están finalizadas
    (el laboratorista terminó) Y TODAS tienen el visto bueno final del Ingeniero — no basta
    con que el laboratorio haya terminado, tiene que estar aprobado para poder entregarse."""
    counts = project_progress(codigo)
    total = sum(counts.values())
    if total == 0 or counts["finalizado"] != total:
        return "ejecucion"
    for perf in st.session_state.perforaciones.get(codigo, []):
        for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
            if m.get("etapa_revision") != "aprobado":
                return "ejecucion"
    return "ejecutado"


def desarchivar_proyecto(codigo):
    """Reabre un proyecto ejecutado: revierte a 'en-proceso' todos sus ensayos finalizados,
    para que el laboratorista pueda volver a digitar o el Jefe agregar nuevas muestras/
    perforaciones. También limpia la aprobación de cada muestra (etapa_revision y demás) —
    si se reabre, tiene que volver a pasar por Jefe e Ingeniero antes de poder archivarse de
    nuevo. El estado 'ejecutado' se recalcula solo (ver project_status)."""
    ahora = now_iso()
    for a in st.session_state.assays:
        if a["codigo_interno"] == codigo and a["status"] == "finalizado":
            a["status"] = "en-proceso"
            a["lastModified"] = ahora
    for perf in st.session_state.perforaciones.get(codigo, []):
        for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
            for campo in ("etapa_revision", "confirmado_por_jefe", "confirmado_por_jefe_fecha",
                          "aprobado_por_ing", "aprobado_por_ing_fecha", "motivo_rechazo", "rechazado_por"):
                m.pop(campo, None)


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
            role_choice = st.radio("Tipo de usuario", ["Auxiliar", "Jefe", "Ingeniero"], horizontal=True)
            password = st.text_input("Clave de acceso", type="password", placeholder="••••••••")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("INGRESAR", type="primary", use_container_width=True):
                role_key = {"Jefe": "jefe", "Ingeniero": "ingeniero"}.get(role_choice, "auxiliar")
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
        c_brand, c_nav, c_bell, c_avatar, c_logout = st.columns([2.2, 4.2, 0.7, 0.7, 0.7])
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
        with c_bell:
            mis_notifs = sorted(
                (n for n in st.session_state.notifications if n["role"] == st.session_state.role),
                key=lambda n: n["fecha"], reverse=True)
            no_leidas = sum(1 for n in mis_notifs if not n["leida"])
            with st.popover(str(no_leidas) if no_leidas else "", icon=":material/notifications:", use_container_width=True):
                st.markdown("**Notificaciones**")
                if not mis_notifs:
                    st.caption("No tienes notificaciones.")
                else:
                    with st.container(key="notif-popover-body"):
                        if no_leidas and st.button("Marcar todas como leídas", key="notif_marcar_todas", use_container_width=True):
                            for n in mis_notifs:
                                n["leida"] = True
                            st.rerun()
                        for n in mis_notifs[:15]:
                            with st.container(border=True):
                                estilo_msg = "font-weight:700;" if not n["leida"] else f"font-weight:400;font-size:13px;color:{MUTED};"
                                st.markdown(f'<div style="{estilo_msg}">{html.escape(n["mensaje"])}</div>'
                                            f'<div class="timestamp-caption">{format_dt(n["fecha"])}</div>', unsafe_allow_html=True)
                                if n.get("muestra_id") and st.button("Ir a la muestra →", key=f"notif_go_{n['id']}", use_container_width=True):
                                    n["leida"] = True
                                    st.session_state.selected_codigo = n["codigo_interno"]
                                    st.session_state.selected_perforacion = n["perforacion_codigo"]
                                    st.session_state.selected_muestra_id = n["muestra_id"]
                                    navigate("muestra-detail")
        with c_avatar:
            iniciales = ROLE_INICIALES.get(st.session_state.role, "AX")
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
def _muestras_pendientes_ing():
    """Muestras ya finalizadas por el laboratorista, confirmadas por el Jefe y esperando el
    visto bueno del Ingeniero. Devuelve tuplas (codigo, perf_codigo, muestra)."""
    pendientes = []
    for p in st.session_state.projects:
        codigo = p["codigo_interno"]
        for perf in st.session_state.perforaciones.get(codigo, []):
            for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
                if compute_muestra_estado(m) == "finalizado" and m.get("etapa_revision") == "pendiente_ing":
                    pendientes.append((codigo, perf["codigo"], m))
    return pendientes


def render_home():
    es_jefe = st.session_state.role == "jefe"
    es_ingeniero = st.session_state.role == "ingeniero"
    es_supervisor = es_jefe or es_ingeniero
    if es_jefe:
        st.markdown("## Bienvenido, Jefe de Laboratorio")
        st.caption("Resumen de operaciones y control de calidad geotécnica para hoy.")
    elif es_ingeniero:
        st.markdown("## Bienvenido, Ingeniero")
        st.caption("Revisión final y aprobación de muestras antes de entregarlas al cliente.")
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
        elif es_ingeniero:
            pendientes_ing = _muestras_pendientes_ing()
            c1, c2 = st.columns([2, 1])
            with c1:
                st.markdown(f'<div class="bento-primary"><div class="bento-icon">{icon("fact_check")}</div>'
                             f'<div><span class="bento-eyebrow">Tareas prioritarias</span>'
                             f'<h3>Muestras pendientes de tu aprobación</h3><p>{len(pendientes_ing)} muestra(s) '
                             f'confirmadas por el Jefe de Laboratorio, esperando tu visto bueno final.</p></div></div>',
                             unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("archive")}</div>'
                             '<div><h3>Proyectos ejecutados</h3><p>Consulta el historial certificado.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Explorar archivo →", key="cta_done_ing", use_container_width=True):
                    navigate("projects-done")
            if pendientes_ing:
                st.markdown("<br>", unsafe_allow_html=True)
                for codigo, perf_codigo, m in pendientes_ing[:5]:
                    proyecto = get_project(codigo)
                    with st.container(border=True):
                        cols = st.columns([3, 1])
                        cols[0].markdown(
                            f'<div class="cell-title">{html.escape(proyecto["nombre"] if proyecto else codigo)}</div>'
                            f'<div class="cell-sub">{html.escape(codigo)} · {html.escape(perf_codigo)} · Muestra {m["numero"]}</div>',
                            unsafe_allow_html=True)
                        with cols[1]:
                            if st.button("Revisar →", key=f"revisar_ing_{m['id_unico']}", use_container_width=True):
                                st.session_state.selected_codigo = codigo
                                st.session_state.selected_perforacion = perf_codigo
                                st.session_state.selected_muestra_id = m["id_unico"]
                                navigate("muestra-detail")
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

    if es_supervisor:
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


def _render_project_list(codes, empty_msg, allow_delete, mark_read_only=False, allow_unarchive=False):
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
                        st.session_state.projects[:] = [x for x in st.session_state.projects if x["codigo_interno"] != codigo]
                        st.session_state.perforaciones.pop(codigo, None)
                        for k in [k for k in st.session_state.muestras if k.startswith(codigo + "::")]:
                            del st.session_state.muestras[k]
                        st.session_state.assays[:] = [a for a in st.session_state.assays if a["codigo_interno"] != codigo]
                        st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo + "::")}
                        st.rerun()
            if allow_unarchive:
                if st.button("Desarchivar proyecto", icon=":material/unarchive:", key=f"unarchive_{p['codigo_interno']}", use_container_width=True):
                    desarchivar_proyecto(p["codigo_interno"])
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
        ensayos = sorted({e for m in muestras for e, activo in m["ensayos"].items() if activo and e in BITACORA_ENSAYOS})
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
                          allow_delete=(st.session_state.role == "jefe"), mark_read_only=True,
                          allow_unarchive=(st.session_state.role == "jefe"))


# ════════════════════════════════════════════════════════════════════
# NUEVO PROYECTO (solo Jefe)
# ════════════════════════════════════════════════════════════════════
def _str_excel(v):
    """str() de una celda de Excel, sin el '.0' feo que deja Python en números enteros
    guardados como float (típico en teléfonos digitados en una celda numérica)."""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v or "").strip()


def _fecha_excel(v):
    """Convierte el valor de una celda de fecha de Excel a un date de Python, o None si
    no se pudo leer como fecha (celda vacía, texto libre, etc.)."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _leer_bitacora_cliente_xlsx(file_obj):
    """Lee el formato GDA-FL-021 (Bitácora de Proyecto) que envía el cliente y devuelve
    los campos que sirven para precargar Nuevo Proyecto."""
    wb = load_workbook(file_obj, data_only=True)
    ws = wb["HOJA1"] if "HOJA1" in wb.sheetnames else wb.active
    return {
        "cliente": _str_excel(ws["E11"].value),
        "nombre": _str_excel(ws["E12"].value),
        "localizacion": _str_excel(ws["E13"].value),
        "direccion_cliente": _str_excel(ws["E14"].value),
        "telefono_contacto": _str_excel(ws["E15"].value),
        "correo_cliente": _str_excel(ws["Q15"].value),
        "nombre_contacto": _str_excel(ws["E16"].value),
        "fecha_inicio_proyecto": _fecha_excel(ws["E17"].value),
        "fecha_final_proyecto": _fecha_excel(ws["Q17"].value),
    }


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

    st.markdown('<div class="section-title">Cargar bitácora de proyecto del cliente (opcional)</div>', unsafe_allow_html=True)
    uploaded_cliente_xlsx = st.file_uploader(
        "Bitácora de proyecto del cliente (Excel)", type=["xlsx"], key="cliente_xlsx_uploader",
        help="Si el cliente te envió el formato GDA-FL-021 (Bitácora de Proyecto), súbelo aquí para "
             "precargar Cliente, Nombre del proyecto, Localización, Dirección, Teléfono, Correo, "
             "Nombre de contacto y fechas de inicio/fin del proyecto.",
    )
    if uploaded_cliente_xlsx is not None and st.session_state.get("_cliente_xlsx_last") != uploaded_cliente_xlsx.name:
        try:
            datos_cliente = _leer_bitacora_cliente_xlsx(uploaded_cliente_xlsx)
            st.session_state["new_nombre"] = datos_cliente["nombre"]
            st.session_state["new_localizacion"] = datos_cliente["localizacion"]
            st.session_state["new_cliente"] = datos_cliente["cliente"]
            st.session_state["new_correo_cliente"] = datos_cliente["correo_cliente"]
            st.session_state["new_direccion_cliente"] = datos_cliente["direccion_cliente"]
            st.session_state["new_telefono_contacto"] = datos_cliente["telefono_contacto"]
            st.session_state["new_nombre_contacto"] = datos_cliente["nombre_contacto"]
            if datos_cliente["fecha_inicio_proyecto"]:
                st.session_state["new_fecha_inicio_proyecto"] = datos_cliente["fecha_inicio_proyecto"]
            if datos_cliente["fecha_final_proyecto"]:
                st.session_state["new_fecha_final_proyecto"] = datos_cliente["fecha_final_proyecto"]
            st.session_state["_cliente_xlsx_last"] = uploaded_cliente_xlsx.name
            st.success("Datos del cliente cargados desde el Excel. Revísalos abajo antes de guardar.")
        except Exception:
            st.error("No se pudo leer el archivo. Verifica que sea el formato GDA-FL-021 (Bitácora de Proyecto).")

    st.markdown('<div class="section-title">Información del proyecto</div>', unsafe_allow_html=True)
    nombre = st.text_input("Nombre del proyecto", key="new_nombre", placeholder="Estudio de suelos vía Bogotá-Medellín")
    localizacion = st.text_input("Localización", key="new_localizacion", placeholder="Km 14+200")
    norma = st.radio("Norma", NORMA_PROYECTO_OPTIONS, horizontal=True)

    c1, c2 = st.columns(2)
    with c1:
        fecha_bitacora = st.date_input("Fecha de bitácora", value=date.today(), format="DD/MM/YYYY")
    with c2:
        fecha_ingreso = st.date_input("Fecha de ingreso de muestra", value=date.today(), format="DD/MM/YYYY")

    ec1, ec2, ec3 = st.columns(3)
    with ec1:
        fecha_recepcion = st.date_input("Fecha de recepción", value=date.today(), format="DD/MM/YYYY")
    with ec2:
        fecha_ejecucion = st.date_input("Fecha de ejecución", value=date.today(), format="DD/MM/YYYY")
    with ec3:
        fecha_emision = st.date_input("Fecha de emisión", value=date.today(), format="DD/MM/YYYY")

    st.markdown('<div class="section-title">Datos del cliente (para el encabezado de los informes — solo el Jefe los ve)</div>', unsafe_allow_html=True)
    cliente = st.text_input("Cliente", key="new_cliente", placeholder="Nombre del cliente")
    direccion_cliente = st.text_input("Dirección cliente", key="new_direccion_cliente", placeholder="Dirección del cliente")
    dc1, dc2 = st.columns(2)
    with dc1:
        telefono_contacto = st.text_input("Teléfono de contacto", key="new_telefono_contacto", placeholder="300 000 0000")
    with dc2:
        correo_cliente = st.text_input("Correo electrónico", key="new_correo_cliente", placeholder="correo@cliente.com")
    nombre_contacto = st.text_input("Nombre de contacto", key="new_nombre_contacto", placeholder="Nombre de quien coordina con el cliente")
    muestra_tomada_por = st.text_input("Muestra tomada por", placeholder="Nombre de quien tomó la muestra")

    dc3, dc4 = st.columns(2)
    with dc3:
        if "new_fecha_inicio_proyecto" not in st.session_state:
            st.session_state["new_fecha_inicio_proyecto"] = date.today()
        fecha_inicio_proyecto = st.date_input("Fecha inicio proyecto", key="new_fecha_inicio_proyecto", format="DD/MM/YYYY")
    with dc4:
        if "new_fecha_final_proyecto" not in st.session_state:
            st.session_state["new_fecha_final_proyecto"] = date.today()
        fecha_final_proyecto = st.date_input("Fecha final proyecto", key="new_fecha_final_proyecto", format="DD/MM/YYYY")

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
                df_source = _bitacora_draft_df(key, muestras)

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

                # Cada perforación se descarga en su propio Excel: la plantilla oficial
                # representa UN sondeo (hoja "S1"), así que no se mezclan varias en un archivo.
                filas_perf_preview = []
                for row in edited.to_dict("records"):
                    numero_m = str(row.get("Número", "")).strip()
                    if not numero_m or numero_m.lower() == "none" or numero_m == "nan":
                        continue
                    filas_perf_preview.append({
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
                excel_bytes, truncado = generar_excel_bitacora_orden(project_preview, filas_perf_preview, {perf["tipo"]})
                st.download_button(f"Descargar bitácora — {perf['codigo']}", data=excel_bytes, icon=":material/download:",
                                    file_name=f"{codigo_interno} Bitacora de orden {perf['codigo']}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key=f"dl_newperf_{key}")
                if truncado:
                    st.caption(f"El formato oficial admite hasta {BITACORA_XLSX_MAX_ROWS} muestras; se incluyeron las primeras {BITACORA_XLSX_MAX_ROWS}.")

                if confirm_delete(f"newperf_{key}", f"la perforación {perf['codigo']}"):
                    st.session_state.perforaciones[codigo_interno] = [p for p in perforaciones if p["codigo"] != perf["codigo"]]
                    st.session_state.muestras.pop(key, None)
                    st.session_state.bitacora_draft.pop(key, None)
                    st.rerun()
    elif nombre or numero or anio:
        st.info("Completa un código interno válido y el nombre del proyecto para agregar perforaciones y muestras.")

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            if codigo_interno:
                st.session_state.perforaciones.pop(codigo_interno, None)
                for k in [k for k in st.session_state.muestras if k.startswith(codigo_interno + "::")]:
                    del st.session_state.muestras[k]
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
                "cliente": cliente, "correo_cliente": correo_cliente, "muestra_tomada_por": muestra_tomada_por,
                "direccion_cliente": direccion_cliente, "telefono_contacto": telefono_contacto,
                "nombre_contacto": nombre_contacto,
                "fecha_inicio_proyecto": str(fecha_inicio_proyecto), "fecha_final_proyecto": str(fecha_final_proyecto),
                "fecha_recepcion": str(fecha_recepcion), "fecha_ejecucion": str(fecha_ejecucion), "fecha_emision": str(fecha_emision),
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
            ("event", "Fecha inicio proyecto", project.get("fecha_inicio_proyecto")),
            ("event_available", "Fecha final proyecto", project.get("fecha_final_proyecto")),
            ("inbox", "Fecha de recepción", project.get("fecha_recepcion")),
            ("science", "Fecha de ejecución", project.get("fecha_ejecucion")),
            ("outbox", "Fecha de emisión", project.get("fecha_emision")),
            ("person", "Asignado a", project.get("laboratorista_asignado")),
        ]
        # Datos del cliente: los ven Jefe e Ingeniero (roles de consulta/revisión), nunca el laboratorista.
        if st.session_state.role in ("jefe", "ingeniero"):
            info_rows.insert(1, ("badge", "Cliente", project.get("cliente")))
            info_rows.insert(2, ("mail", "Correo electrónico", project.get("correo_cliente")))
            info_rows.insert(3, ("home_pin", "Dirección cliente", project.get("direccion_cliente")))
            info_rows.insert(4, ("call", "Teléfono de contacto", project.get("telefono_contacto")))
            info_rows.insert(5, ("contact_page", "Nombre de contacto", project.get("nombre_contacto")))
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
                st.session_state.projects[:] = [p for p in st.session_state.projects if p["codigo_interno"] != codigo]
                st.session_state.perforaciones.pop(codigo, None)
                for k in [k for k in st.session_state.muestras if k.startswith(codigo + "::")]:
                    del st.session_state.muestras[k]
                st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo + "::")}
                st.session_state.assays[:] = [a for a in st.session_state.assays if a["codigo_interno"] != codigo]
                navigate("home")
        if project_status(codigo) == "ejecutado":
            st.caption("Este proyecto ya está completamente ejecutado (todas sus muestras finalizadas).")
            if st.button("Desarchivar proyecto", icon=":material/unarchive:", use_container_width=True):
                desarchivar_proyecto(codigo)
                st.success("Proyecto desarchivado — vuelve a aparecer en Proyectos en ejecución.")
                st.rerun()

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
        if st.button(label_bitacora, type="primary", icon=f":material/{icon_bitacora}:", use_container_width=True):
            navigate("bitacora")

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
            bc1, bc2 = st.columns(2)
            with bc1:
                if st.button("Ver muestras →", key=f"open_perf_{perf['codigo']}", use_container_width=True):
                    st.session_state.selected_perforacion = perf["codigo"]
                    navigate("perforacion-detail")
            with bc2:
                filas_perf = _bitacora_filas_perforacion(codigo, perf["codigo"])
                excel_bytes, _truncado = generar_excel_bitacora_orden(project, filas_perf, {perf["tipo"]})
                st.download_button("Descargar bitácora", data=excel_bytes, icon=":material/download:",
                                    file_name=f"{codigo} Bitacora de orden {perf['codigo']}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key=f"dl_perf_{perf['codigo']}")


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

    def _parse_fecha(valor):
        try:
            return date.fromisoformat(valor)
        except (TypeError, ValueError):
            return date.today()

    # Campos con `key=` (namespaded por código de proyecto) para poder precargarlos desde el
    # Excel del cliente sin pisar los de otro proyecto que se haya editado en la misma sesión.
    for campo, valor_defecto in (
        ("nombre", project.get("nombre", "")), ("localizacion", project.get("localizacion", "")),
        ("cliente", project.get("cliente", "")), ("direccion_cliente", project.get("direccion_cliente", "")),
        ("telefono_contacto", project.get("telefono_contacto", "")), ("correo_cliente", project.get("correo_cliente", "")),
        ("nombre_contacto", project.get("nombre_contacto", "")),
    ):
        wkey = f"edit_{campo}_{codigo}"
        if wkey not in st.session_state:
            st.session_state[wkey] = valor_defecto

    st.markdown('<div class="section-title">Cargar bitácora de proyecto del cliente (opcional)</div>', unsafe_allow_html=True)
    uploaded_cliente_xlsx_edit = st.file_uploader(
        "Bitácora de proyecto del cliente (Excel)", type=["xlsx"], key=f"cliente_xlsx_uploader_edit_{codigo}",
        help="Si el cliente te envió el formato GDA-FL-021 (Bitácora de Proyecto), súbelo aquí para "
             "precargar Cliente, Nombre del proyecto, Localización, Dirección, Teléfono, Correo, "
             "Nombre de contacto y fechas de inicio/fin del proyecto.",
    )
    guard_key = f"_cliente_xlsx_last_edit_{codigo}"
    if uploaded_cliente_xlsx_edit is not None and st.session_state.get(guard_key) != uploaded_cliente_xlsx_edit.name:
        try:
            datos_cliente = _leer_bitacora_cliente_xlsx(uploaded_cliente_xlsx_edit)
            st.session_state[f"edit_nombre_{codigo}"] = datos_cliente["nombre"]
            st.session_state[f"edit_localizacion_{codigo}"] = datos_cliente["localizacion"]
            st.session_state[f"edit_cliente_{codigo}"] = datos_cliente["cliente"]
            st.session_state[f"edit_correo_cliente_{codigo}"] = datos_cliente["correo_cliente"]
            st.session_state[f"edit_direccion_cliente_{codigo}"] = datos_cliente["direccion_cliente"]
            st.session_state[f"edit_telefono_contacto_{codigo}"] = datos_cliente["telefono_contacto"]
            st.session_state[f"edit_nombre_contacto_{codigo}"] = datos_cliente["nombre_contacto"]
            if datos_cliente["fecha_inicio_proyecto"]:
                st.session_state[f"edit_fecha_inicio_proyecto_{codigo}"] = datos_cliente["fecha_inicio_proyecto"]
            if datos_cliente["fecha_final_proyecto"]:
                st.session_state[f"edit_fecha_final_proyecto_{codigo}"] = datos_cliente["fecha_final_proyecto"]
            st.session_state[guard_key] = uploaded_cliente_xlsx_edit.name
            st.success("Datos del cliente cargados desde el Excel. Revísalos abajo antes de guardar.")
        except Exception:
            st.error("No se pudo leer el archivo. Verifica que sea el formato GDA-FL-021 (Bitácora de Proyecto).")

    nombre = st.text_input("Nombre del proyecto", key=f"edit_nombre_{codigo}")
    localizacion = st.text_input("Localización", key=f"edit_localizacion_{codigo}")
    norma_actual = project.get("norma")
    idx = NORMA_PROYECTO_OPTIONS.index(norma_actual) if norma_actual in NORMA_PROYECTO_OPTIONS else 0
    norma = st.radio("Norma", NORMA_PROYECTO_OPTIONS, index=idx, horizontal=True)

    c1, c2 = st.columns(2)
    with c1:
        fecha_bitacora = st.date_input("Fecha de bitácora", value=_parse_fecha(project.get("fecha_bitacora")), format="DD/MM/YYYY")
    with c2:
        fecha_ingreso = st.date_input("Fecha de ingreso de muestra", value=_parse_fecha(project.get("fecha_ingreso_muestra")), format="DD/MM/YYYY")

    ec1, ec2, ec3 = st.columns(3)
    with ec1:
        fecha_recepcion = st.date_input("Fecha de recepción", value=_parse_fecha(project.get("fecha_recepcion")), format="DD/MM/YYYY")
    with ec2:
        fecha_ejecucion = st.date_input("Fecha de ejecución", value=_parse_fecha(project.get("fecha_ejecucion")), format="DD/MM/YYYY")
    with ec3:
        fecha_emision = st.date_input("Fecha de emisión", value=_parse_fecha(project.get("fecha_emision")), format="DD/MM/YYYY")

    st.markdown('<div class="section-title">Datos del cliente (para el encabezado de los informes)</div>', unsafe_allow_html=True)
    cliente = st.text_input("Cliente", key=f"edit_cliente_{codigo}")
    direccion_cliente = st.text_input("Dirección cliente", key=f"edit_direccion_cliente_{codigo}")
    dc1, dc2 = st.columns(2)
    with dc1:
        telefono_contacto = st.text_input("Teléfono de contacto", key=f"edit_telefono_contacto_{codigo}")
    with dc2:
        correo_cliente = st.text_input("Correo electrónico", key=f"edit_correo_cliente_{codigo}")
    nombre_contacto = st.text_input("Nombre de contacto", key=f"edit_nombre_contacto_{codigo}")
    muestra_tomada_por = st.text_input("Muestra tomada por", value=project.get("muestra_tomada_por", ""))

    if f"edit_fecha_inicio_proyecto_{codigo}" not in st.session_state:
        st.session_state[f"edit_fecha_inicio_proyecto_{codigo}"] = _parse_fecha(project.get("fecha_inicio_proyecto"))
    if f"edit_fecha_final_proyecto_{codigo}" not in st.session_state:
        st.session_state[f"edit_fecha_final_proyecto_{codigo}"] = _parse_fecha(project.get("fecha_final_proyecto"))

    dc3, dc4 = st.columns(2)
    with dc3:
        fecha_inicio_proyecto = st.date_input("Fecha inicio proyecto", key=f"edit_fecha_inicio_proyecto_{codigo}", format="DD/MM/YYYY")
    with dc4:
        fecha_final_proyecto = st.date_input("Fecha final proyecto", key=f"edit_fecha_final_proyecto_{codigo}", format="DD/MM/YYYY")

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
            project["cliente"] = cliente
            project["correo_cliente"] = correo_cliente
            project["muestra_tomada_por"] = muestra_tomada_por
            project["direccion_cliente"] = direccion_cliente
            project["telefono_contacto"] = telefono_contacto
            project["nombre_contacto"] = nombre_contacto
            project["fecha_inicio_proyecto"] = str(fecha_inicio_proyecto)
            project["fecha_final_proyecto"] = str(fecha_final_proyecto)
            project["fecha_recepcion"] = str(fecha_recepcion)
            project["fecha_ejecucion"] = str(fecha_ejecucion)
            project["fecha_emision"] = str(fecha_emision)
            navigate("project-detail")


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PERFORACIÓN → LISTA DE MUESTRAS
# ════════════════════════════════════════════════════════════════════
def _perforacion_ensayos_progress(codigo, perf_codigo):
    muestras = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    total_ensayos, completados = 0, 0
    for m in muestras:
        for label, activo in m["ensayos"].items():
            if not activo or label not in BITACORA_ENSAYOS:
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
            filas_perf = _bitacora_filas_perforacion(codigo, perf_codigo)
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
                ensayos_sol = [e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS]
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


def _bitacora_draft_df(key, muestras):
    """Inicializa (o repara) el borrador de bitácora de esta perforación en session_state,
    agregando cualquier columna de BITACORA_BASE_COLS que falte —por ejemplo un ensayo nuevo
    que se agregó a BITACORA_ENSAYOS después de que este borrador ya existía en la sesión—
    sin perder lo que el usuario ya haya digitado."""
    if key not in st.session_state.bitacora_draft:
        df = pd.DataFrame(_muestras_to_rows(muestras))
    else:
        df = st.session_state.bitacora_draft[key]
    for col in BITACORA_BASE_COLS:
        if col not in df.columns:
            df[col] = _bitacora_row_defaults()[col]
    st.session_state.bitacora_draft[key] = df[BITACORA_BASE_COLS]
    return st.session_state.bitacora_draft[key]


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
            df_source = _bitacora_draft_df(key, muestras)

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
                filas_perf_rows = edited.to_dict("records")
            else:
                st.dataframe(df_source, use_container_width=True, hide_index=True)
                filas_perf_rows = df_source.to_dict("records")

            # Cada perforación se descarga en su propio Excel: la plantilla oficial
            # representa UN sondeo (hoja "S1"), así que no se mezclan varias en un archivo.
            filas_perf = []
            for row in filas_perf_rows:
                numero_m = str(row.get("Número", "")).strip()
                if not numero_m or numero_m.lower() == "none" or numero_m == "nan":
                    continue
                filas_perf.append({
                    "perf_codigo": perf["codigo"], "numero": numero_m,
                    "tipo_muestra": row.get("Tipo de muestra") or TIPO_MUESTRA_OPTIONS[0],
                    "profundidad_de": row.get("Prof. De") or 0.0, "profundidad_hasta": row.get("Prof. A") or 0.0,
                    "ensayos": {e: bool(row.get(e, False)) for e in BITACORA_ENSAYOS},
                    "observaciones": row.get("Observaciones") or "",
                })
            excel_bytes, truncado = generar_excel_bitacora_orden(project, filas_perf, {perf["tipo"]})
            st.download_button(f"Descargar bitácora — {perf['codigo']}", data=excel_bytes, icon=":material/download:",
                                file_name=f"{codigo} Bitacora de orden {perf['codigo']}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True, key=f"dl_bitacora_{key}")
            if truncado:
                st.caption(f"El formato oficial admite hasta {BITACORA_XLSX_MAX_ROWS} muestras; se incluyeron las primeras {BITACORA_XLSX_MAX_ROWS}.")

            if es_jefe and confirm_delete(f"perf_{key}", f"la perforación {perf['codigo']} y todas sus muestras"):
                st.session_state.perforaciones[codigo] = [p for p in st.session_state.perforaciones[codigo] if p["codigo"] != perf["codigo"]]
                st.session_state.muestras.pop(key, None)
                st.session_state.bitacora_draft.pop(key, None)
                st.session_state.assays[:] = [a for a in st.session_state.assays if not (a["codigo_interno"] == codigo and a["perforacion_codigo"] == perf["codigo"])]
                st.rerun()



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

    estado = compute_muestra_estado(muestra)
    with st.container(border=True, key="muestra-header-card"):
        top = st.columns([3, 1])
        with top[0]:
            st.markdown(f'<div style="display:flex;align-items:center;gap:8px;">'
                        f'<h3 style="margin:0;">Muestra {muestra["numero"]}</h3>'
                        f'{status_badge_html(estado, font_size=13)}</div>', unsafe_allow_html=True)
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

    with st.container(border=True):
        st.markdown('<div class="section-title">Descripción visual de la muestra</div>', unsafe_allow_html=True)
        st.caption("Cómo se ve físicamente la muestra (color, textura, humedad, etc.). Esta es la que se lleva "
                   "al campo \"DESCRIPCIÓN VISUAL\" del Excel oficial — es independiente de las observaciones.")
        with st.container(key="muestra-desc-visual-box"):
            descripcion_visual = st.text_area(
                "Descripción visual de la muestra", value=muestra.get("descripcion_visual", ""), label_visibility="collapsed",
                placeholder="Ej: Suelo arcilloso color gris, humedad media, sin fragmentos de roca...", key=f"desc_visual_{muestra_id}",
            )
        if st.button("Guardar descripción visual", icon=":material/save:", key=f"desc_visual_save_{muestra_id}"):
            muestra["descripcion_visual"] = descripcion_visual.upper()
            st.success("Descripción visual guardada.")

    with st.container(border=True):
        st.markdown('<div class="section-title">Observaciones</div>', unsafe_allow_html=True)
        st.caption("El Jefe la digita antes de que el laboratorista inicie el ensayo (instrucciones, contexto del "
                   "cliente, etc.); el laboratorista la digita si la muestra presenta fisuras o no se puede "
                   "realizar el ensayo por alguna razón. Se guarda para todos los ensayos de esta muestra y la "
                   "puede editar tanto el Jefe como el laboratorista.")
        with st.container(key="muestra-obs-box"):
            observacion = st.text_area(
                "Observaciones", value=muestra.get("observaciones", ""), label_visibility="collapsed",
                placeholder="Ej: Muestra con fisuras visibles, no fue posible completar el ensayo...", key=f"obs_{muestra_id}",
            )
        if st.button("Guardar observación", icon=":material/save:", key=f"obs_save_{muestra_id}"):
            muestra["observaciones"] = observacion
            st.success("Observación guardada.")

    # Filtra por si la muestra guarda un ensayo que ya no es seleccionable (p. ej. "Pasa 200",
    # que quedó incluido dentro de Granulometría) — no se muestra aunque quede marcado en datos viejos.
    solicitados = [e for e, v in muestra["ensayos"].items() if v and e in BITACORA_ENSAYOS]
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
            tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
            existing = get_assay(muestra_id, tipo_interno) if tipo_interno else None
            status = existing["status"] if existing else "sin-iniciar"

            cols = st.columns([0.6, 2.4, 1.6, 1])
            cols[0].markdown(status_circle_html(status), unsafe_allow_html=True)
            cols[1].markdown(f"**{ensayo_label}**")
            if tipo_interno:
                cols[2].markdown(status_badge_html(status), unsafe_allow_html=True)
                with cols[3]:
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
                cols[2].markdown('<span class="badge badge-muted">Sin formulario aún</span>', unsafe_allow_html=True)

    if estado == "finalizado":
        st.markdown('<div class="section-title">Aprobación de la muestra</div>', unsafe_allow_html=True)
        etapa = muestra.get("etapa_revision")
        motivo = muestra.get("motivo_rechazo")
        if motivo:
            quien = "el Jefe de Laboratorio" if muestra.get("rechazado_por") == "jefe" else "el Ingeniero"
            st.warning(f"Devuelto por {quien}: {motivo}")

        with st.container(border=True):
            if etapa == "aprobado":
                st.markdown(card_header_html("verified", "Aprobado — listo para entregar al cliente"), unsafe_allow_html=True)
                st.markdown(
                    f'<div class="cell-muted">Confirmado por Jefe</div><div style="font-weight:600;">'
                    f'{html.escape(muestra.get("confirmado_por_jefe") or "—")} · {format_dt(muestra.get("confirmado_por_jefe_fecha"))}</div>'
                    f'<div class="cell-muted" style="margin-top:10px;">Aprobado por Ingeniero</div><div style="font-weight:600;">'
                    f'{html.escape(muestra.get("aprobado_por_ing") or "—")} · {format_dt(muestra.get("aprobado_por_ing_fecha"))}</div>',
                    unsafe_allow_html=True)
            elif etapa == "pendiente_ing":
                st.markdown(card_header_html("hourglass_top", "Enviado al Ingeniero — esperando revisión"), unsafe_allow_html=True)
                st.markdown(
                    f'<div class="cell-muted">Confirmado por Jefe</div><div style="font-weight:600;">'
                    f'{html.escape(muestra.get("confirmado_por_jefe") or "—")} · {format_dt(muestra.get("confirmado_por_jefe_fecha"))}</div>',
                    unsafe_allow_html=True)
                if st.session_state.role == "ingeniero":
                    st.markdown("<br>", unsafe_allow_html=True)
                    nombre_ing = st.text_input("Tu nombre", key=f"ing_nombre_{muestra_id}", placeholder="Nombre completo")
                    if st.button("Dar visto bueno final", type="primary", icon=":material/check_circle:",
                                 use_container_width=True, key=f"ing_aprobar_{muestra_id}"):
                        muestra["etapa_revision"] = "aprobado"
                        muestra["aprobado_por_ing"] = nombre_ing
                        muestra["aprobado_por_ing_fecha"] = now_iso()
                        muestra["motivo_rechazo"] = ""
                        muestra["rechazado_por"] = ""
                        add_notification("jefe", f"El Ingeniero aprobó la Muestra {muestra['numero']} de {codigo} "
                                                  f"— ya se puede entregar al cliente.", codigo, perf_codigo, muestra_id)
                        quien = f" ({nombre_ing})" if nombre_ing else ""
                        add_historial_muestra(muestra, f"Aprobado por el Ingeniero{quien}.")
                        st.success("Aprobado.")
                        st.rerun()
                    motivo_ing = st.text_area("Motivo si vas a devolver al Jefe", key=f"ing_motivo_{muestra_id}",
                                               placeholder="Qué hay que corregir...")
                    if st.button("Devolver al Jefe", icon=":material/undo:", use_container_width=True, key=f"ing_devolver_{muestra_id}"):
                        if motivo_ing.strip():
                            muestra["etapa_revision"] = None
                            muestra["motivo_rechazo"] = motivo_ing
                            muestra["rechazado_por"] = "ing"
                            add_notification("jefe", f"El Ingeniero devolvió la Muestra {muestra['numero']} de {codigo}: "
                                                      f"{motivo_ing}", codigo, perf_codigo, muestra_id)
                            add_historial_muestra(muestra, f"Devuelto por el Ingeniero: {motivo_ing}")
                            st.success("Devuelto al Jefe.")
                            st.rerun()
                        else:
                            st.error("Escribe el motivo antes de devolver.")
            else:
                st.markdown(card_header_html("fact_check", "Pendiente de confirmación del Jefe de Laboratorio"), unsafe_allow_html=True)
                if st.session_state.role == "jefe":
                    nombre_jefe = st.text_input("Tu nombre", key=f"jefe_nombre_{muestra_id}", placeholder="Nombre completo")
                    if st.button("Confirmar y enviar al Ingeniero", type="primary", icon=":material/check_circle:",
                                 use_container_width=True, key=f"jefe_confirmar_{muestra_id}"):
                        muestra["etapa_revision"] = "pendiente_ing"
                        muestra["confirmado_por_jefe"] = nombre_jefe
                        muestra["confirmado_por_jefe_fecha"] = now_iso()
                        muestra["motivo_rechazo"] = ""
                        muestra["rechazado_por"] = ""
                        add_notification("ingeniero", f"El Jefe de Laboratorio envió la Muestra {muestra['numero']} de "
                                                       f"{codigo} para tu revisión final.", codigo, perf_codigo, muestra_id)
                        quien = f" ({nombre_jefe})" if nombre_jefe else ""
                        add_historial_muestra(muestra, f"Confirmado por el Jefe de Laboratorio{quien} y enviado al Ingeniero.")
                        st.success("Enviado al Ingeniero.")
                        st.rerun()
                    motivo_jefe = st.text_area("Motivo si vas a devolver al laboratorista", key=f"jefe_motivo_{muestra_id}",
                                                placeholder="Qué hay que corregir...")
                    if st.button("Devolver al laboratorista", icon=":material/undo:", use_container_width=True, key=f"jefe_devolver_{muestra_id}"):
                        if motivo_jefe.strip():
                            ahora = now_iso()
                            for e in solicitados:
                                tipo_i = SUPPORTED_ASSAY_MAP.get(e)
                                a = get_assay(muestra_id, tipo_i) if tipo_i else None
                                if a and a["status"] == "finalizado":
                                    a["status"] = "en-proceso"
                                    a["lastModified"] = ahora
                            muestra["etapa_revision"] = None
                            muestra["motivo_rechazo"] = motivo_jefe
                            muestra["rechazado_por"] = "jefe"
                            add_notification("auxiliar", f"El Jefe de Laboratorio devolvió la Muestra {muestra['numero']} "
                                                          f"de {codigo}: {motivo_jefe}", codigo, perf_codigo, muestra_id)
                            add_historial_muestra(muestra, f"Devuelto por el Jefe de Laboratorio: {motivo_jefe}")
                            st.success("Devuelto al laboratorista.")
                            st.rerun()
                        else:
                            st.error("Escribe el motivo antes de devolver.")
                else:
                    st.caption("El Jefe de Laboratorio debe confirmar esta muestra antes de enviarla al Ingeniero.")

    historial = muestra.get("historial", [])
    if historial:
        with st.expander(f"Historial de la muestra ({len(historial)})", icon=":material/history:"):
            for h in sorted(historial, key=lambda h: h["fecha"], reverse=True):
                st.markdown(f'<div style="margin-bottom:10px;"><div style="font-weight:600;">{html.escape(h["texto"])}</div>'
                            f'<div class="timestamp-caption">{format_dt(h["fecha"])}</div></div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE BITÁCORA DE ORDEN (plantilla oficial GDA-FL-003)
# ════════════════════════════════════════════════════════════════════
def _bitacora_filas_perforacion(codigo, perf_codigo):
    """Muestras de UNA perforación ya guardada, en el formato que espera
    generar_excel_bitacora_orden. Cada perforación se exporta a su propio Excel —
    la plantilla oficial (hoja "S1") representa un solo sondeo, no varios a la vez."""
    filas = []
    for m in st.session_state.muestras.get(f"{codigo}::{perf_codigo}", []):
        filas.append({
            "perf_codigo": perf_codigo, "numero": m["numero"], "tipo_muestra": m["tipo_muestra"],
            "profundidad_de": m["profundidad_de"], "profundidad_hasta": m["profundidad_hasta"],
            "ensayos": m["ensayos"], "observaciones": m.get("observaciones", ""),
        })
    return filas


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
# GENERAR EXCEL DE GRANULOMETRÍA Y HUMEDAD (plantillas reales del laboratorio,
# ambas comparten el mismo diseño de encabezado — filas 1 a 13)
# ════════════════════════════════════════════════════════════════════
def _fecha_ddmmaaaa(iso_str):
    """Convierte una fecha guardada en formato ISO ("AAAA-MM-DD") a DD/MM/AAAA para que
    se vea igual que en la app al pasarla al Excel."""
    try:
        return date.fromisoformat(iso_str).strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return iso_str or ""


def _llenar_encabezado_informe(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo="", perf_numero_cell="F12"):
    ws["D6"] = project.get("cliente", "") if project else ""  # Cliente
    ws["D7"] = project["nombre"] if project else codigo          # Proyecto
    ws["D8"] = project.get("correo_cliente", "") if project else ""  # Correo electrónico
    ws["D9"] = project.get("localizacion", "") if project else ""  # Localización
    ws["D10"] = project.get("muestra_tomada_por", "") if project else ""  # Muestra tomada por
    ws["K6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""  # Fecha de recepción
    ws["K7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""  # Fecha de ejecución
    ws["K8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""  # Fecha de emisión
    ws["L9"] = project.get("numero", "") if project else ""  # Código interno — número (K9 ya trae "GDA")
    ws["M9"] = project.get("anio", "") if project else ""  # Código interno — año

    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""  # Tipo de perforación (lista desplegable)
    # El número de perforación cae en una celda distinta según la plantilla: F12 en la de
    # Humedad (GDA-FLC-014), E12 en la de Granulometría/Límites (GDA-FLC-001, actualizada 2026).
    ws[perf_numero_cell] = perf["consecutivo"] if perf else ""
    ws["H12"] = muestra["numero"]
    ws["K12"] = to_float(muestra.get("profundidad_de"))
    ws["M12"] = to_float(muestra.get("profundidad_hasta"))
    # Descripción visual: primero lo que se digitó en "Descripción visual de la muestra" (Bitácora)
    # —independiente de las Observaciones—, si no hay, lo que el laboratorista escribió en
    # "Observaciones" del propio ensayo, y solo si ninguna de las dos existe, el tipo de muestra
    # como último recurso.
    ws["D13"] = muestra.get("descripcion_visual") or observaciones_ensayo or f"Tipo de muestra: {muestra.get('tipo_muestra','')}"


def _escribir_limites(ws, data):
    def _escribir(filas):
        for key, _label, cells in filas:
            es_recipiente = key.endswith("recipiente")
            for i, cell in enumerate(cells, start=1):
                valor = data.get(f"{key}_{i}", "")
                if es_recipiente:
                    ws[cell] = valor
                else:
                    num = to_float(valor)
                    if num is not None:
                        ws[cell] = num
    _escribir(LIMITE_LIQUIDO_FILAS)
    _escribir(LIMITE_PLASTICO_FILAS)


def _reparar_graficos_perdidos(xlsm_bytes, template_path):
    """openpyxl no conserva las 'chartUserShapes' de un gráfico (las anotaciones de texto
    dibujadas a mano encima, ej. las etiquetas LÍNEA U/LÍNEA A/CH/CL-ML de la Carta de
    Plasticidad): se pierden tanto los archivos (drawingN.xml, chartN.xml.rels) como —esto
    es lo que de verdad hace que no se vean— la propia referencia <c:userShapes r:id="..."/>
    dentro del chartN.xml, que es lo que le dice a Excel que busque esas formas. El gráfico
    y sus series/ejes sí sobreviven intactos porque no dependen de esa referencia.
    Esta función restaura los archivos faltantes desde la plantilla original Y vuelve a
    insertar la referencia dentro del/de los chartN.xml correspondientes."""
    with zipfile.ZipFile(template_path) as tpl:
        tpl_names = set(tpl.namelist())
        with zipfile.ZipFile(BytesIO(xlsm_bytes)) as out:
            out_names = set(out.namelist())
            faltantes = {n for n in tpl_names - out_names
                         if n.startswith("xl/drawings/") or n.startswith("xl/charts/_rels/")}
            if not faltantes:
                return xlsm_bytes

            content_types_tpl = tpl.read("[Content_Types].xml").decode("utf-8")
            content_types_out = out.read("[Content_Types].xml").decode("utf-8")
            for nombre in faltantes:
                if not nombre.endswith(".xml"):
                    continue  # los .rels usan el Default de extensión "rels", no necesitan Override
                parte = "/" + nombre
                m = re.search(r'<Override PartName="' + re.escape(parte) + r'"[^>]*?/>', content_types_tpl)
                if m and m.group(0) not in content_types_out:
                    content_types_out = content_types_out.replace("</Types>", m.group(0) + "</Types>")

            # Para cada chartN.xml.rels restaurado, extraemos el rId de su relación
            # "chartUserShapes" y lo reinsertamos dentro del chartN.xml de salida (que openpyxl
            # ya escribió, pero sin esa referencia).
            R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            charts_a_parchar = {}  # "xl/charts/chartN.xml" -> rId
            for nombre in faltantes:
                m = re.match(r"xl/charts/_rels/(chart\d+)\.xml\.rels$", nombre)
                if not m:
                    continue
                rels_xml = tpl.read(nombre).decode("utf-8")
                rm = re.search(
                    r'<Relationship[^>]*Type="[^"]*chartUserShapes"[^>]*Id="([^"]+)"'
                    r'|<Relationship[^>]*Id="([^"]+)"[^>]*Type="[^"]*chartUserShapes"',
                    rels_xml,
                )
                if rm:
                    rid = rm.group(1) or rm.group(2)
                    charts_a_parchar[f"xl/charts/{m.group(1)}.xml"] = rid

            chart_bytes_parchados = {}
            for chart_nombre, rid in charts_a_parchar.items():
                if chart_nombre not in out_names:
                    continue
                chart_xml = out.read(chart_nombre).decode("utf-8")
                if "userShapes" in chart_xml:
                    continue  # ya la tiene, nada que hacer
                # La etiqueta raíz puede no declarar el namespace "r:" (openpyxl no lo usa en
                # el cuerpo del chart), así que lo agregamos si hace falta.
                if 'xmlns:r=' not in chart_xml.split(">", 1)[0]:
                    chart_xml = chart_xml.replace(
                        "<chartSpace ", f'<chartSpace xmlns:r="{R_NS}" ', 1
                    )
                userShapes_tag = f'<userShapes r:id="{rid}"/>'
                chart_xml = chart_xml.replace("</chartSpace>", userShapes_tag + "</chartSpace>")
                chart_bytes_parchados[chart_nombre] = chart_xml.encode("utf-8")

            bio = BytesIO()
            with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
                for item in out.infolist():
                    if item.filename == "[Content_Types].xml":
                        z.writestr(item, content_types_out)
                    elif item.filename in chart_bytes_parchados:
                        z.writestr(item, chart_bytes_parchados[item.filename])
                    else:
                        z.writestr(item, out.read(item.filename))
                for nombre in faltantes:
                    z.writestr(nombre, tpl.read(nombre))
            bio.seek(0)
            return bio.getvalue()


def _generar_excel_clasificacion(codigo, perf_codigo, muestra, project, gran_data=None, lim_data=None, observaciones_ensayo=""):
    """Granulometría y Límites de Atterberg comparten la misma plantilla y hoja ("GUIA") —
    por muestra van juntos en un solo archivo, sin importar si se descarga desde el ensayo de
    Granulometría o desde el de Límites."""
    wb = load_workbook(TEMPLATE_GRANULOMETRIA, keep_vba=True)
    ws = wb["GUIA"]
    _llenar_encabezado_informe(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo, perf_numero_cell="E12")

    if gran_data is not None:
        masa_inicial_seca = to_float(gran_data.get("masa_inicial_seca"))
        if masa_inicial_seca is not None:
            ws["D17"] = masa_inicial_seca
        # Si solo se solicitó Pasa 200 (sin Granulometría), `gran_data` no trae los tamices —
        # se dejan las celdas tal cual trae la plantilla en vez de escribir 0, para no simular
        # una curva granulométrica falsa ("pasa 100%" en todos los tamices).
        for key, _label, _apert, cell in SIEVES:
            valor = to_float(gran_data.get(key))
            if valor is not None:
                ws[cell] = valor

    if lim_data is not None:
        _escribir_limites(ws, lim_data)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _reparar_graficos_perdidos(bio.getvalue(), TEMPLATE_GRANULOMETRIA)


def generar_excel_granulometria(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    lim_assay = get_assay(muestra["id_unico"], "limites")
    lim_data = lim_assay.get("data", {}) if lim_assay else None
    return _generar_excel_clasificacion(codigo, perf_codigo, muestra, project, gran_data=data, lim_data=lim_data,
                                         observaciones_ensayo=observaciones_ensayo)


def generar_excel_pasa200(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Pasa 200 usa la misma plantilla de Granulometría. `data` ya viene resuelto por
    render_assay_form: si la muestra también tiene Granulometría, es el mismo diccionario de
    esa muestra (con tamices incluidos si se digitaron); si no, son solo los datos propios de
    Pasa 200."""
    lim_assay = get_assay(muestra["id_unico"], "limites")
    lim_data = lim_assay.get("data", {}) if lim_assay else None
    return _generar_excel_clasificacion(codigo, perf_codigo, muestra, project, gran_data=data, lim_data=lim_data,
                                         observaciones_ensayo=observaciones_ensayo)


def generar_excel_humedad(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    wb = load_workbook(TEMPLATE_HUMEDAD)
    ws = wb["GUIA"]

    _llenar_encabezado_informe(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo)
    ws["I19"] = data.get("hum_recipiente", "")
    ws["I20"] = to_float(data.get("hum_masa_humedo_mas_recipiente"))
    ws["I21"] = to_float(data.get("hum_seco_mas_recipiente"))
    ws["I22"] = to_float(data.get("hum_masa_recipiente"))
    # I23 (masa seca) e I24 (% humedad) son fórmulas de la propia plantilla; no se tocan.

    metodo = data.get("hum_metodo", "")
    ws["C28"] = "MÉTODO A" if metodo == "Método A" else ("MÉTODO B" if metodo == "Método B" else "")
    temp_horno = data.get("hum_temp_horno", "")
    ws["E28"] = "110°C" if "110" in temp_horno else ("60°C" if "60" in temp_horno else "")

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def generar_excel_limites(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    gran_assay = get_assay(muestra["id_unico"], "granulometria")
    gran_data = gran_assay.get("data", {}) if gran_assay else None
    return _generar_excel_clasificacion(codigo, perf_codigo, muestra, project, gran_data=gran_data, lim_data=data,
                                         observaciones_ensayo=observaciones_ensayo)


def _llenar_encabezado_masa_unitaria(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo=""):
    """La plantilla de Peso Unitario Parafinado (GDA-FLC-004) usa una distribución de celdas
    de encabezado propia, distinta a la de Granulometría/Límites/Humedad."""
    ws["C6"] = project.get("cliente", "") if project else ""  # Cliente
    ws["C7"] = project["nombre"] if project else codigo  # Proyecto
    ws["C8"] = project.get("correo_cliente", "") if project else ""  # Correo electrónico
    ws["C9"] = project.get("localizacion", "") if project else ""  # Localización
    ws["C10"] = project.get("muestra_tomada_por", "") if project else ""  # Muestra tomada por
    ws["I6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""  # Fecha de recepción
    ws["I7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""  # Fecha de ejecución
    ws["I8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""  # Fecha de emisión
    ws["J9"] = project.get("numero", "") if project else ""  # Código interno — número (I9 ya trae "GDA")
    ws["K9"] = project.get("anio", "") if project else ""  # Código interno — año

    perf = get_perforacion(codigo, perf_codigo)
    ws["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""  # Tipo de perforación (lista desplegable)
    ws["D12"] = perf["consecutivo"] if perf else ""  # Número de perforación
    ws["F12"] = muestra["numero"]  # Muestra No.
    ws["I12"] = to_float(muestra.get("profundidad_de"))
    ws["K12"] = to_float(muestra.get("profundidad_hasta"))
    ws["C13"] = muestra.get("descripcion_visual") or observaciones_ensayo or f"Tipo de muestra: {muestra.get('tipo_muestra','')}"


def generar_excel_masa_unitaria(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    wb = load_workbook(TEMPLATE_MASA_UNITARIA)
    ws = wb["GUIA"]
    _llenar_encabezado_masa_unitaria(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo)

    ws["E20"] = to_float(data.get("mu_peso_aire"))  # B = masa en el aire
    ws["F20"] = to_float(data.get("mu_peso_aire_par"))  # C = masa en el aire parafinado
    ws["G20"] = to_float(data.get("mu_peso_agua_par"))  # D = masa parafinada sumergida
    # La densidad de la parafina (L24) no se digita en la app — se deja el 0.86 por defecto
    # que ya trae la plantilla.

    # La humedad (G28) la necesita la fórmula de "densidad seca" pero esta plantilla no la
    # digita — se toma del ensayo de Humedad de la misma muestra, igual que Granulometría y
    # Límites comparten datos entre sí.
    hum_assay = get_assay(muestra["id_unico"], "humedad")
    humedad_pct = calcular_humedad_pct(hum_assay.get("data", {})) if hum_assay else None
    if humedad_pct is not None:
        ws["G28"] = humedad_pct
    # A (masa de la cuerda, D20) y temperatura del agua no tienen celda equivalente en esta
    # plantilla — quedan para completar manualmente en el Excel.

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ════════════════════════════════════════════════════════════════════
# FORMULARIOS DE ENSAYO (solo captura de datos, sin cálculos)
# ════════════════════════════════════════════════════════════════════
def render_equipo(data, prefix, equipo_list=None):
    lista = equipo_list or EQUIPO_LIST
    with st.container(border=True):
        st.markdown(card_header_html("construction", "Equipos Utilizados"), unsafe_allow_html=True)
        seleccionados = set(data.get(f"{prefix}_equipos", []))
        cols = st.columns(2)
        nuevos = []
        for i, equipo in enumerate(lista):
            with cols[i % 2]:
                if st.checkbox(equipo, value=equipo in seleccionados, key=f"{prefix}_equipo_{i}"):
                    nuevos.append(equipo)
        data[f"{prefix}_equipos"] = nuevos


def render_norma_selector(assay_type, data, key_prefix):
    with st.container(border=True):
        st.markdown(card_header_html("rule", "Norma a Utilizar"), unsafe_allow_html=True)
        options = NORMAS_ENSAYO[assay_type]
        current = data.get(f"{key_prefix}_norma", "")
        idx = options.index(current) if current in options else 0
        choice = st.selectbox("Norma", options, index=idx, key=f"norma_{key_prefix}", label_visibility="collapsed")
        data[f"{key_prefix}_norma"] = choice


PASA_200_FILAS = [
    ("p200_recipiente", "Recipiente No."),
    ("p200_seco_mas_recipiente", "Masa suelo seco + recipiente"),
    ("p200_seco_14h", "Masa suelo seco (14 hrs)"),
    ("p200_seco_15h", "Masa suelo seco (15 hrs)"),
    ("p200_seco_16h", "Masa suelo seco (16 hrs)"),
    ("p200_masa_recipiente", "Masa del recipiente"),
]


def render_pasa200_section(data, assay_id, requerido=True):
    """Campos de "Determinación Pasa No. 200". Se usa tanto embebido dentro del formulario de
    Granulometría como en el formulario del ensayo "Pasa 200" independiente — en ambos casos
    `data` puede terminar siendo el mismo diccionario compartido (ver render_assay_form), así
    que lo que se digite en cualquiera de las dos pantallas se refleja en la otra."""
    badge = '<span class="badge badge-warning">Requerido</span>' if requerido else ""
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Determinación Pasa No. 200", badge), unsafe_allow_html=True)
        head = st.columns([2.2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Antes del lavado (g)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Después del lavado (g)</div>', unsafe_allow_html=True)
        for key, label in PASA_200_FILAS:
            row = st.columns([2.2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            for suffix, col in (("antes", row[1]), ("despues", row[2])):
                field_key = f"{key}_{suffix}"
                widget_key = f"{field_key}_{assay_id}"
                # No se pasa `value=` junto con un key que también se controla por session_state
                # (el autocompletado de abajo lo hace) — Streamlit no permite mezclar ambos.
                if widget_key not in st.session_state:
                    st.session_state[widget_key] = data.get(field_key, "")
                data[field_key] = col.text_input(
                    f"{label} {suffix}", key=widget_key, label_visibility="collapsed", placeholder="0.00")
            if key == "p200_seco_mas_recipiente":
                # Autocompleta las 3 lecturas de horas con esta masa apenas se digita, pero si el
                # laboratorista ya las cambió a mano, no se vuelven a pisar en el siguiente rerun —
                # solo se repite el autocompletado cuando el valor de origen vuelve a cambiar.
                for suffix in ("antes", "despues"):
                    src_key = f"{key}_{suffix}"
                    current_val = data[src_key]
                    lastsync_key = f"{src_key}_lastsync"
                    if data.get(lastsync_key) != current_val:
                        for hkey in ("p200_seco_14h", "p200_seco_15h", "p200_seco_16h"):
                            st.session_state[f"{hkey}_{suffix}_{assay_id}"] = current_val
                            data[f"{hkey}_{suffix}"] = current_val
                        data[lastsync_key] = current_val

        # La plantilla de Excel necesita la masa inicial seca (neta, sin el recipiente) para calcular
        # el % que pasa cada tamiz. Se deriva de las lecturas "antes del lavado" ya digitadas arriba
        # (no es un campo nuevo en la interfaz, solo cómo se arma el dato para el Excel).
        masa_seco_mas_recip = to_float(data.get("p200_seco_mas_recipiente_antes"))
        masa_recip = to_float(data.get("p200_masa_recipiente_antes"))
        data["masa_inicial_seca"] = (masa_seco_mas_recip - masa_recip) if (masa_seco_mas_recip is not None and masa_recip is not None) else ""


def render_pasa200_form(data, assay_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel de Granulometría — "
            "si la muestra también tiene Granulometría, los dos ensayos comparten los mismos datos.")
    render_norma_selector("granulometria", data, "gran")
    render_equipo(data, "gran", EQUIPO_GRANULOMETRIA)
    render_pasa200_section(data, assay_id, requerido=False)


def render_granulometria_form(data, assay_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — los cálculos y la clasificación USCS los hace el Excel, no la app.")

    render_norma_selector("granulometria", data, "gran")
    render_equipo(data, "gran", EQUIPO_GRANULOMETRIA)

    render_pasa200_section(data, assay_id)

    with st.container(border=True):
        st.markdown(card_header_html("grid_view", "Granulometría (Masa de Suelo Retenido)"), unsafe_allow_html=True)
        # Campos de texto libre (no st.data_editor con NumberColumn) porque el editor de tabla
        # de Streamlit borra el punto decimal apenas se escribe más de un dígito después de él
        # — el mismo patrón confiable que ya se usa en Pasa No. 200 y en Humedad.
        head = st.columns([1.2, 1, 1.4])
        head[0].markdown('<div class="cell-muted" style="font-weight:700;">Tamiz</div>', unsafe_allow_html=True)
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Abertura (mm)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Retenido (g)</div>', unsafe_allow_html=True)
        for key, label, apert, _cell in SIEVES:
            row = st.columns([1.2, 1, 1.4])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            row[1].markdown(f'<div style="padding-top:8px;text-align:center;">{apert}</div>', unsafe_allow_html=True)
            widget_key = f"retenido_{key}_{assay_id}"
            if widget_key not in st.session_state:
                # Ensayos guardados antes de este cambio pueden tener el valor como float
                # (venía del st.data_editor con NumberColumn) — text_input necesita un str.
                raw = data.get(key, "")
                st.session_state[widget_key] = "" if raw in (None, "") else str(raw)
            data[key] = row[2].text_input(f"Retenido {label}", key=widget_key, label_visibility="collapsed", placeholder="0.00")
        st.caption("El % retenido y la clasificación USCS se calculan en la plantilla de Excel, no aquí.")


def render_humedad_form(data, assay_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — el % de humedad lo calcula el Excel, no la app.")

    render_norma_selector("humedad", data, "hum")
    render_equipo(data, "hum", EQUIPO_HUMEDAD)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Determinación de Humedad"), unsafe_allow_html=True)

        def _campo(key, label, placeholder="0.00"):
            row = st.columns([2.2, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder=placeholder)

        _campo("hum_recipiente", "Recipiente no.", placeholder="Ej: 839")
        _campo("hum_masa_recipiente", "Masa del recipiente (g)")
        _campo("hum_masa_humedo_mas_recipiente", "Masa suelo húmedo + recipiente (g)")

        # "Masa suelo seco + recipiente" se autocompleta en las 3 lecturas de horas apenas se
        # digita, igual que en Pasa No. 200 de Granulometría — si el laboratorista cambia una
        # lectura a mano, ya no se vuelve a pisar hasta que el valor de origen vuelva a cambiar.
        src_key = "hum_seco_mas_recipiente"
        src_widget_key = f"{src_key}_{assay_id}"
        if src_widget_key not in st.session_state:
            st.session_state[src_widget_key] = data.get(src_key, "")
        row = st.columns([2.2, 1])
        row[0].markdown('<div style="padding-top:8px;">Masa suelo seco + recipiente (g)</div>', unsafe_allow_html=True)
        data[src_key] = row[1].text_input("Masa suelo seco + recipiente (g)", key=src_widget_key,
                                           label_visibility="collapsed", placeholder="0.00")
        current_val = data[src_key]
        lastsync_key = f"{src_key}_lastsync"
        if data.get(lastsync_key) != current_val:
            for hkey in ("hum_seco_14h", "hum_seco_15h", "hum_seco_16h"):
                st.session_state[f"{hkey}_{assay_id}"] = current_val
                data[hkey] = current_val
            data[lastsync_key] = current_val
        for hkey, hlabel in (("hum_seco_14h", "Masa suelo seco + recipiente (g) (14 hrs)"),
                              ("hum_seco_15h", "Masa suelo seco + recipiente (g) (15 hrs)"),
                              ("hum_seco_16h", "Masa suelo seco + recipiente (g) (16 hrs)")):
            hwidget_key = f"{hkey}_{assay_id}"
            if hwidget_key not in st.session_state:
                st.session_state[hwidget_key] = data.get(hkey, "")
            hrow = st.columns([2.2, 1])
            hrow[0].markdown(f'<div style="padding-top:8px;">{hlabel}</div>', unsafe_allow_html=True)
            data[hkey] = hrow[1].text_input(hlabel, key=hwidget_key, label_visibility="collapsed", placeholder="0.00")

        st.caption("La masa del agua y la masa de suelo seco se calculan solas (restando la masa del recipiente) — no se digitan aquí.")

    with st.container(border=True):
        st.markdown(card_header_html("local_fire_department", "Datos del Laboratorio"), unsafe_allow_html=True)
        temp_actual = data.get("hum_temp_horno", "110 ± 5 °C")
        opciones_temp = ["110 ± 5 °C", "60 °C"]
        idx = opciones_temp.index(temp_actual) if temp_actual in opciones_temp else 0
        c1, c2 = st.columns(2)
        with c1:
            data["hum_temp_horno"] = st.selectbox("Temperatura Horno", opciones_temp, index=idx, key=f"hum_temp_horno_{assay_id}")
        with c2:
            metodo_actual = data.get("hum_metodo", METODO_HUMEDAD[0])
            midx = METODO_HUMEDAD.index(metodo_actual) if metodo_actual in METODO_HUMEDAD else 0
            data["hum_metodo"] = st.selectbox("Método del Ensayo", METODO_HUMEDAD, index=midx, key=f"hum_metodo_{assay_id}")


def render_masa_unitaria_form(data):
    st.info("Estos datos se guardan tal cual, sin calcular el peso unitario dentro de la app.")
    c1, c2 = st.columns(2)
    with c1:
        data["mu_peso_aire"] = st.text_input("Masa en el aire (g)", value=data.get("mu_peso_aire", ""), placeholder="245.80")
        data["mu_peso_agua_par"] = st.text_input("Masa en el agua parafinado (g)", value=data.get("mu_peso_agua_par", ""), placeholder="138.20")
    with c2:
        data["mu_peso_aire_par"] = st.text_input("Masa en el aire parafinado (g)", value=data.get("mu_peso_aire_par", ""), placeholder="258.30")
        data["mu_temp_agua"] = st.text_input("Temperatura del agua (°C)", value=data.get("mu_temp_agua", ""), placeholder="22.0")

    render_equipo(data, "mu", EQUIPO_MASA_UNITARIA)
    render_norma_selector("masa-unitaria", data, "mu")


def render_limites_form(data, assay_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — el Límite Líquido, el Límite Plástico y el Índice de Plasticidad los calcula el Excel, no la app.")

    with st.container(border=True):
        st.markdown(card_header_html("info", "Información de Ensayo"), unsafe_allow_html=True)
        metodo_actual = data.get("lim_metodo", METODO_HUMEDAD[0])
        midx = METODO_HUMEDAD.index(metodo_actual) if metodo_actual in METODO_HUMEDAD else 0
        data["lim_metodo"] = st.radio("Método de Ensayo", METODO_HUMEDAD, index=midx, horizontal=True, key=f"lim_metodo_{assay_id}")

    def _tabla_limite(icono, titulo, filas, n):
        with st.container(border=True):
            st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
            head = st.columns([2] + [1] * n)
            head[0].markdown('<div class="cell-muted" style="font-weight:700;">Parámetro</div>', unsafe_allow_html=True)
            for i in range(n):
                head[i + 1].markdown(f'<div class="cell-muted" style="text-align:center;font-weight:700;">Ensayo {i + 1}</div>', unsafe_allow_html=True)
            for key, label, _cells in filas:
                row = st.columns([2] + [1] * n)
                row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                for i in range(n):
                    field_key = f"{key}_{i + 1}"
                    widget_key = f"{field_key}_{assay_id}"
                    if widget_key not in st.session_state:
                        raw = data.get(field_key, "")
                        st.session_state[widget_key] = "" if raw in (None, "") else str(raw)
                    data[field_key] = row[i + 1].text_input(f"{label} {i + 1}", key=widget_key, label_visibility="collapsed", placeholder="0.00")

                if key.endswith("_seco"):
                    # "Masa suelo seco + rec." se autocompleta en las lecturas de 14/15/16 horas
                    # de cada columna, igual que en Humedad y Pasa No. 200 — si el laboratorista
                    # cambia una lectura a mano, no se vuelve a pisar hasta que el valor de
                    # origen vuelva a cambiar.
                    for i in range(n):
                        field_key = f"{key}_{i + 1}"
                        current_val = data[field_key]
                        lastsync_key = f"{field_key}_lastsync"
                        if data.get(lastsync_key) != current_val:
                            for suffix in ("14h", "15h", "16h"):
                                hkey = f"{key}_{suffix}_{i + 1}"
                                st.session_state[f"{hkey}_{assay_id}"] = current_val
                                data[hkey] = current_val
                            data[lastsync_key] = current_val

    _tabla_limite("water_drop", "Límite Líquido (INV. 125 - 13)", LIMITE_LIQUIDO_FILAS, LIMITE_LIQUIDO_N)
    _tabla_limite("gesture", "Límite Plástico (INV. 126 - 13)", LIMITE_PLASTICO_FILAS, LIMITE_PLASTICO_N)

    render_equipo(data, "lim", EQUIPO_LIMITES)


def render_read_only_summary(tipo, data, laboratorista="—"):
    """Vista de solo lectura ('Resultados de Ensayo') — la misma para el Jefe (siempre) y para
    el auxiliar cuando el proyecto ya fue ejecutado. Sin casillas de digitación, solo tarjetas
    y tablas con los datos ya registrados."""
    if tipo == "granulometria":
        # Toda orden de granulometría incluye el Pasa No. 200 — se muestra tal cual se digitó,
        # con sus dos columnas (antes/después del lavado), igual que en el formulario editable.
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Determinación Pasa No. 200"), unsafe_allow_html=True)
            pasa200_rows = [(label, data.get(f"{key}_antes"), data.get(f"{key}_despues")) for key, label in PASA_200_FILAS]
            st.markdown(param_table_3col_html(pasa200_rows), unsafe_allow_html=True)
        # "Masa inicial seca" no se muestra en la app (ni aquí ni en el formulario editable) —
        # se deriva solo al momento de generar el Excel (ver generar_excel_granulometria), tal
        # como se llena a mano en la plantilla física: masa suelo seco + recipiente, menos recipiente.
        with st.container(border=True):
            st.markdown(card_header_html("grid_view", "Granulometría (Masa de Suelo Retenido)"), unsafe_allow_html=True)
            sieve_rows = [(label, data.get(key, "—")) for key, label, _apert, _cell in SIEVES]
            st.markdown(param_table_html(sieve_rows, header_left="TAMIZ", header_right="RETENIDO (g)"), unsafe_allow_html=True)
        equipos, norma = data.get("gran_equipos", []), data.get("gran_norma", "—")
    elif tipo == "pasa200":
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Determinación Pasa No. 200"), unsafe_allow_html=True)
            pasa200_rows = [(label, data.get(f"{key}_antes"), data.get(f"{key}_despues")) for key, label in PASA_200_FILAS]
            st.markdown(param_table_3col_html(pasa200_rows), unsafe_allow_html=True)
        equipos, norma = data.get("gran_equipos", []), data.get("gran_norma", "—")
    elif tipo == "humedad":
        masa_humedo = to_float(data.get("hum_masa_humedo_mas_recipiente"))
        masa_seco = to_float(data.get("hum_seco_mas_recipiente"))
        masa_recip = to_float(data.get("hum_masa_recipiente"))
        masa_agua = (masa_humedo - masa_seco) if (masa_humedo is not None and masa_seco is not None) else None
        masa_suelo_seco = (masa_seco - masa_recip) if (masa_seco is not None and masa_recip is not None) else None
        humedad_pct = calcular_humedad_pct(data)
        rows = [
            ("Recipiente no.", data.get("hum_recipiente")),
            ("Masa del recipiente (g)", data.get("hum_masa_recipiente")),
            ("Masa suelo húmedo + recipiente (g)", data.get("hum_masa_humedo_mas_recipiente")),
            ("Masa suelo seco + recipiente (g) (14 hrs)", data.get("hum_seco_14h")),
            ("Masa suelo seco + recipiente (g) (15 hrs)", data.get("hum_seco_15h")),
            ("Masa suelo seco + recipiente (g) (16 hrs)", data.get("hum_seco_16h")),
            ("Masa del agua (g)", fmt_num(masa_agua)),
            ("Masa suelo seco (g)", fmt_num(masa_suelo_seco)),
            ("Humedad (%)", fmt_num(humedad_pct, decimals=2)),
        ]
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("local_fire_department", "Datos del Laboratorio"), unsafe_allow_html=True)
            lab_rows = [
                ("Temperatura Horno", data.get("hum_temp_horno")),
                ("Método del Ensayo", data.get("hum_metodo")),
                ("Laboratorista", laboratorista),
            ]
            st.markdown(param_table_html(lab_rows, header_left="DATO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("hum_equipos", []), data.get("hum_norma", "—")
    elif tipo == "limites":
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Límite Líquido (INV. 125 - 13)"), unsafe_allow_html=True)
            headers = ["PARÁMETRO"] + [f"ENSAYO {i}" for i in range(1, LIMITE_LIQUIDO_N + 1)]
            ll_rows = [(label, *[data.get(f"{key}_{i}") for i in range(1, LIMITE_LIQUIDO_N + 1)]) for key, label, _c in LIMITE_LIQUIDO_FILAS]
            st.markdown(param_table_ncol_html(headers, ll_rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("gesture", "Límite Plástico (INV. 126 - 13)"), unsafe_allow_html=True)
            headers = ["PARÁMETRO"] + [f"ENSAYO {i}" for i in range(1, LIMITE_PLASTICO_N + 1)]
            lp_rows = [(label, *[data.get(f"{key}_{i}") for i in range(1, LIMITE_PLASTICO_N + 1)]) for key, label, _c in LIMITE_PLASTICO_FILAS]
            st.markdown(param_table_ncol_html(headers, lp_rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("info", "Información de Ensayo"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Método de Ensayo", data.get("lim_metodo"))], header_left="DATO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("lim_equipos", []), "INV. E-125-13 / INV. E-126-13"
    else:
        rows = [("Masa en el aire (g)", data.get("mu_peso_aire")), ("Masa en el aire parafinado (g)", data.get("mu_peso_aire_par")),
                ("Masa en el agua parafinado (g)", data.get("mu_peso_agua_par")), ("Temperatura del agua (°C)", data.get("mu_temp_agua"))]
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        equipos, norma = data.get("mu_equipos", []), data.get("mu_norma", "—")

    with st.container(border=True):
        st.markdown(card_header_html("rule", "Norma Aplicada"), unsafe_allow_html=True)
        st.markdown(f'<div style="font-weight:600;">{html.escape(norma or "—")}</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(card_header_html("construction", "Equipos Utilizados"), unsafe_allow_html=True)
        st.markdown(equipos_readonly_html(equipos), unsafe_allow_html=True)


def render_assay_form():
    assay_id = st.session_state.selected_assay_id
    assay = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    if not assay:
        navigate("muestra-detail")
        return

    codigo, perf_codigo, muestra_id = assay["codigo_interno"], assay["perforacion_codigo"], assay["muestra_id"]
    project = get_project(codigo)
    muestra = get_muestra(codigo, perf_codigo, muestra_id)
    es_jefe = st.session_state.role == "jefe"
    es_supervisor = st.session_state.role in ("jefe", "ingeniero")
    # El Jefe y el Ingeniero solo consultan los ensayos — quien digita los datos de laboratorio es el laboratorista.
    read_only = es_supervisor or (st.session_state.role == "auxiliar" and project_status(codigo) == "ejecutado")

    if st.button("← Atrás"):
        go_back(fallback="muestra-detail")

    st.markdown(f"## {ASSAY_LABELS[assay['tipo']]}")
    st.caption("Resultados de Ensayo" if read_only else "Registro de Ensayo")
    st.markdown(f'<div style="margin-bottom:10px;">{status_badge_html(assay["status"])}&nbsp;&nbsp;'
                f'<span class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(assay["lastModified"])}'
                + (f' · {html.escape(assay["laboratorist"])}' if assay.get("laboratorist") else "") + '</span></div>',
                unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(card_header_html("info", "Información General"), unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        g1.markdown(f'<div class="cell-muted">Proyecto</div><div style="font-weight:600;">{html.escape(codigo)}</div>', unsafe_allow_html=True)
        g2.markdown(f'<div class="cell-muted">Sondeo</div><div style="font-weight:600;">{html.escape(perf_codigo)}</div>', unsafe_allow_html=True)
        g3, g4 = st.columns(2)
        g3.markdown(f'<div class="cell-muted" style="margin-top:12px;">Muestra</div>'
                    f'<div style="font-weight:600;">M-{html.escape(str(muestra["numero"])) if muestra else "—"}</div>', unsafe_allow_html=True)
        profundidad_txt = f'{muestra["profundidad_de"]:.2f}m - {muestra["profundidad_hasta"]:.2f}m' if muestra else "—"
        g4.markdown(f'<div class="cell-muted" style="margin-top:12px;">Profundidad</div><div style="font-weight:600;">{profundidad_txt}</div>',
                    unsafe_allow_html=True)

    if muestra is not None:
        with st.container(border=True):
            st.markdown(card_header_html("thermostat", "Condición del Ensayo"), unsafe_allow_html=True)
            st.caption("Se digita una sola vez por muestra: la inicial al empezar el ensayo y la final al terminarlo. Se comparte entre todos los ensayos de esta muestra.")
            if read_only:
                st.markdown(condicion_table_html(muestra), unsafe_allow_html=True)
            else:
                head = st.columns([1.4, 1, 1])
                head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Temperatura °C</div>', unsafe_allow_html=True)
                head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Humedad %</div>', unsafe_allow_html=True)
                for cond_key, cond_label in (("inicial", "Inicial"), ("final", "Final")):
                    row = st.columns([1.4, 1, 1])
                    row[0].markdown(f'<div style="padding-top:8px;">{cond_label}</div>', unsafe_allow_html=True)
                    muestra[f"cond_{cond_key}_temp"] = row[1].text_input(
                        f"Temperatura {cond_label}", value=muestra.get(f"cond_{cond_key}_temp", ""),
                        key=f"cond_{cond_key}_temp_{muestra_id}", label_visibility="collapsed", placeholder="0.0")
                    muestra[f"cond_{cond_key}_hum"] = row[2].text_input(
                        f"Humedad {cond_label}", value=muestra.get(f"cond_{cond_key}_hum", ""),
                        key=f"cond_{cond_key}_hum_{muestra_id}", label_visibility="collapsed", placeholder="0")

    # Pasa 200 comparte plantilla y datos con Granulometría: si la muestra también tiene un
    # ensayo de Granulometría, "Pasa 200" lee y escribe directamente sobre ESE diccionario de
    # datos (no el suyo propio), así que lo digitado en cualquiera de las dos pantallas se ve
    # reflejado en la otra. Si no hay Granulometría, Pasa 200 usa sus propios datos, igual que
    # cualquier otro ensayo independiente.
    pasa200_gran_sibling = get_assay(muestra_id, "granulometria") if assay["tipo"] == "pasa200" else None
    data = dict(pasa200_gran_sibling["data"]) if pasa200_gran_sibling else dict(assay.get("data", {}))

    if read_only:
        if es_supervisor:
            st.info("Estás viendo el ensayo en modo consulta — solo el laboratorista puede digitar estos datos.")
        else:
            st.info("Este proyecto ya fue ejecutado. Estás en modo consulta — no puedes editar estos datos.")
        render_read_only_summary(assay["tipo"], data, assay.get("laboratorist") or "—")
        with st.container(border=True):
            st.markdown(card_header_html("notes", "Observaciones"), unsafe_allow_html=True)
            st.markdown(f'<div>{html.escape(assay.get("observations") or "—")}</div>', unsafe_allow_html=True)
        if assay["tipo"] != "humedad":
            with st.container(border=True):
                st.markdown(card_header_html("person", "Laboratorista"), unsafe_allow_html=True)
                st.markdown(f'<div style="font-weight:600;">{html.escape(assay.get("laboratorist") or "—")}</div>', unsafe_allow_html=True)
        if es_jefe and assay["status"] == "finalizado":
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Habilitar edición para el laboratorista", icon=":material/lock_open:", use_container_width=True):
                assay.update(status="en-proceso", lastModified=now_iso())
                st.success("Ensayo habilitado — el laboratorista ya puede volver a digitar los datos.")
                st.rerun()
    else:
        if assay["tipo"] == "granulometria":
            render_granulometria_form(data, assay_id)
        elif assay["tipo"] == "pasa200":
            render_pasa200_form(data, assay_id)
        elif assay["tipo"] == "humedad":
            render_humedad_form(data, assay_id)
        elif assay["tipo"] == "limites":
            render_limites_form(data, assay_id)
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
                if pasa200_gran_sibling:
                    pasa200_gran_sibling["data"] = data
                navigate("muestra-detail")
        with col2:
            if st.button("Enviar a revisión", type="primary", use_container_width=True, icon=":material/send:"):
                ya_estaba_finalizado = assay["status"] == "finalizado"
                assay.update(data=data, observations=observations, laboratorist=laboratorist, status="finalizado", lastModified=now_iso())
                if pasa200_gran_sibling:
                    pasa200_gran_sibling["data"] = data
                if muestra:
                    quien = f" por {laboratorist}" if laboratorist else ""
                    add_historial_muestra(muestra, f"{ASSAY_LABELS[assay['tipo']]} enviado a revisión{quien}.")
                    # Si este era el último ensayo pendiente, la muestra completa su ciclo en el
                    # laboratorio (semáforo en rojo) — se avisa al Jefe para que la confirme.
                    if not ya_estaba_finalizado and compute_muestra_estado(muestra) == "finalizado":
                        add_notification("jefe", f"La Muestra {muestra['numero']} de {codigo} ya completó todos sus "
                                                  f"ensayos — está lista para tu confirmación.", codigo, perf_codigo, muestra_id)
                        add_historial_muestra(muestra, "Todos los ensayos completados — pendiente de confirmación del Jefe de Laboratorio.")
                navigate("muestra-detail")

    if es_supervisor and assay["tipo"] == "granulometria" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_granulometria(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (Granulometría y Límites de Atterberg — mismo archivo por muestra)", icon=":material/download:",
            data=excel_bytes, file_name=f"Clasificacion_de_suelos_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True,
        )

    if es_supervisor and assay["tipo"] == "pasa200" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_pasa200(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (Granulometría — mismo archivo por muestra)", icon=":material/download:",
            data=excel_bytes, file_name=f"Clasificacion_de_suelos_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True,
        )

    if es_supervisor and assay["tipo"] == "humedad" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_humedad(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (plantilla oficial de Humedad)", icon=":material/download:",
            data=excel_bytes, file_name=f"Humedad_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )

    if es_supervisor and assay["tipo"] == "limites" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_limites(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (Granulometría y Límites de Atterberg — mismo archivo por muestra)", icon=":material/download:",
            data=excel_bytes, file_name=f"Clasificacion_de_suelos_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True,
        )

    if es_supervisor and assay["tipo"] == "masa-unitaria" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_masa_unitaria(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (plantilla oficial de Peso Unitario Parafinado)", icon=":material/download:",
            data=excel_bytes, file_name=f"Peso_unitario_parafinado_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
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


SEARCH_PAGE_SIZE = 8


def render_search():
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Buscar ensayos")

    codes = [p["codigo_interno"] for p in st.session_state.projects]
    if not codes:
        st.info("Todavía no hay proyectos.")
        return

    with st.container(border=True):
        st.markdown(card_header_html("filter_list", "Filtros de Búsqueda"), unsafe_allow_html=True)
        default_idx = codes.index(st.session_state.selected_codigo) if st.session_state.selected_codigo in codes else 0
        codigo = st.selectbox("Proyecto", codes, index=default_idx)

        perforaciones = st.session_state.perforaciones.get(codigo, [])
        perf_options = ["(todas)"] + [p["codigo"] for p in perforaciones]
        perf_choice = st.selectbox("Perforación", perf_options)

        perfs_to_show = perforaciones if perf_choice == "(todas)" else [p for p in perforaciones if p["codigo"] == perf_choice]
        # El código del proyecto ya se eligió arriba — en este desplegable solo hace falta la
        # perforación y el número de muestra, no el id_unico completo repitiendo el proyecto.
        muestra_label_by_id = {
            m["id_unico"]: f"{perf['codigo']} · Muestra {m['numero']}"
            for perf in perfs_to_show for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        }
        muestra_options = ["(todas)"] + list(muestra_label_by_id.keys())
        muestra_choice = st.selectbox("Muestra", muestra_options,
                                       format_func=lambda v: "(todas)" if v == "(todas)" else muestra_label_by_id[v])

        f_type = st.selectbox("Tipo de ensayo", ["(todos)"] + list(ASSAY_LABELS.values()))

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Aplicar filtros", type="primary", use_container_width=True, icon=":material/search:"):
            st.session_state["search_page"] = 0

    if not perforaciones:
        st.info("Este proyecto todavía no tiene perforaciones. Ve a la Bitácora para agregarlas.")
        return

    project = get_project(codigo)
    rows = []
    for perf in perfs_to_show:
        muestras = st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", [])
        for m in muestras:
            if muestra_choice != "(todas)" and m["id_unico"] != muestra_choice:
                continue
            solicitados = [e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS]
            if f_type != "(todos)":
                solicitados = [e for e in solicitados if ASSAY_LABELS.get(SUPPORTED_ASSAY_MAP.get(e), None) == f_type]
            for ensayo_label in solicitados:
                tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
                if not tipo_interno:
                    continue
                rows.append((perf, m, ensayo_label, tipo_interno))

    with st.container(border=True):
        col_ratios = [1.7, 2.2, 1.8, 1.3, 0.6]
        headers = st.columns(col_ratios)
        for col, label in zip(headers, ["ID ensayo", "Proyecto", "Tipo / Muestra", "Estado", ""]):
            col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)

        if not rows:
            st.info("No se encontraron ensayos con esos filtros.")
        else:
            total = len(rows)
            total_pages = max(1, (total + SEARCH_PAGE_SIZE - 1) // SEARCH_PAGE_SIZE)
            page = min(st.session_state.get("search_page", 0), total_pages - 1)
            st.session_state["search_page"] = page
            start = page * SEARCH_PAGE_SIZE
            for perf, m, ensayo_label, tipo_interno in rows[start:start + SEARCH_PAGE_SIZE]:
                existing = get_assay(m["id_unico"], tipo_interno)
                status = existing["status"] if existing else "sin-iniciar"
                ensayo_id = f'{codigo}-{perf["codigo"]}-M{m["numero"]}'
                cols = st.columns(col_ratios, vertical_alignment="center")
                cols[0].markdown(f'<span class="cell-id">{html.escape(ensayo_id)}</span>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="cell-title">{html.escape(project["nombre"] if project else codigo)}</div>'
                                  f'<div class="cell-sub">{html.escape(codigo)}</div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="cell-title">{html.escape(ensayo_label)}</div>'
                                  f'<div class="cell-sub">Muestra {html.escape(str(m["numero"]))}</div>', unsafe_allow_html=True)
                cols[3].markdown(status_badge_html(status), unsafe_allow_html=True)
                with cols[4]:
                    if st.button("", key=f"search_open_{m['id_unico']}_{tipo_interno}", icon=":material/chevron_right:",
                                 use_container_width=True, help="Abrir"):
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

            st.markdown("<hr style='margin:8px 0;'>", unsafe_allow_html=True)
            f1, f2, f3 = st.columns([3, 1, 1])
            f1.caption(f"Mostrando {len(rows[start:start + SEARCH_PAGE_SIZE])} de {total} resultado(s)")
            with f2:
                if st.button("", key="search_prev", icon=":material/chevron_left:", use_container_width=True, disabled=page == 0):
                    st.session_state["search_page"] = page - 1
                    st.rerun()
            with f3:
                if st.button("", key="search_next", icon=":material/chevron_right:", use_container_width=True,
                             disabled=page >= total_pages - 1):
                    st.session_state["search_page"] = page + 1
                    st.rerun()


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
