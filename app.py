"""
GEODELTA LAB - App para digitar ensayos de laboratorio de suelos
Estructura: Proyecto -> Perforación (Sondeo/Apique/Fuente-Cantera) -> Muestra -> Ensayo

Cómo correrla en tu computador:
    streamlit run app.py
"""

import base64
import html
import json
import math
import os
import re
import time
import zipfile
from datetime import date, datetime, timedelta
from io import BytesIO

import extra_streamlit_components as stx
import pandas as pd
from PIL import Image, ImageOps
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

import db

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
TEMPLATE_MASA_UNITARIA_B = os.path.join(BASE_DIR, "templates", "GDA-FLC-030_peso_unitario_b.xlsx")
TEMPLATE_CBR = os.path.join(BASE_DIR, "templates", "GDA-FLC-013_cbr.xlsx")
TEMPLATE_CORTE_DIRECTO = os.path.join(BASE_DIR, "templates", "GDA-FLC-007_corte_directo.xlsx")
TEMPLATE_GESP_FINO = os.path.join(BASE_DIR, "templates", "GDA-FLC-027_gravedad_fino.xlsx")
TEMPLATE_GESP_GRUESO = os.path.join(BASE_DIR, "templates", "GDA-FLC-028_gravedad_grueso.xlsx")
TEMPLATE_GESP_ARCILLA = os.path.join(BASE_DIR, "templates", "GDA-FLC-006_gravedad_arcilla.xlsx")
TEMPLATE_PROCTOR = os.path.join(BASE_DIR, "templates", "GDA-FLC-002_proctor_cbr.xlsm")
TEMPLATE_MATERIA_ORGANICA = os.path.join(BASE_DIR, "templates", "GDA-FLC-003_materia_organica.xlsx")
TEMPLATE_LIMITE_CONTRACCION = os.path.join(BASE_DIR, "templates", "GDA-FLC-022_limite_contraccion.xlsx")
TEMPLATE_CONSOLIDACION = os.path.join(BASE_DIR, "templates", "GDA-FLC-009_consolidacion.xlsx")
TEMPLATE_COMPRESION_INCONFINADA = os.path.join(BASE_DIR, "templates", "GDA-FLC-008_compresion_inconfinada.xlsm")
TEMPLATE_COMPRESION_ROCA = os.path.join(BASE_DIR, "templates", "GDA-FLC-043_compresion_roca.xlsx")
TEMPLATE_CARGA_PUNTUAL = os.path.join(BASE_DIR, "templates", "GDA-FLC-018_carga_puntual.xlsx")
TEMPLATE_SOLIDEZ_SULFATOS = os.path.join(BASE_DIR, "templates", "GDA-FLC-033_solidez_sulfatos.xlsx")
TEMPLATE_TERRONES_ARCILLA = os.path.join(BASE_DIR, "templates", "GDA-FLC-034_terrones_arcilla.xlsx")
TEMPLATE_VERIFICACION_BALANZAS = os.path.join(BASE_DIR, "templates", "GDA-FL-029_verificacion_balanzas.xlsx")

ROLE_LABELS = {"jefe": "Jefe de Laboratorio", "laboratorista": "Laboratorista", "ingeniero": "Director Técnico"}
ROLE_INICIALES = {"jefe": "JL", "laboratorista": "LB", "ingeniero": "DT"}

# ════════════════════════════════════════════════════════════════════
# ESTILOS — paleta "Verdant Precision" (Primary #007A33 · Secondary #4A7862 · Tertiary #D1E8D5 · Neutral #212121)
# ════════════════════════════════════════════════════════════════════
PRIMARY, PRIMARY_DARK, PRIMARY_CONTAINER = "#007A33", "#00591F", "#0B3D22"
SECONDARY, SECONDARY_CONTAINER = "#4A7862", "#D1E8D5"
TERTIARY = "#D1E8D5"
NEUTRAL = "#6B7570"
SUCCESS, SUCCESS_LIGHT = "#16A34A", "#DCFCE7"
WARNING, WARNING_LIGHT = "#D97706", "#FEF3C7"
DANGER, DANGER_LIGHT = "#DC2626", "#FEE2E2"
SURFACE, BG, BORDER, TEXT = "#FFFFFF", "#F7FAF8", "#D6D9D5", "#212121"
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
    /* Fondo de cuadrícula sutil (papel milimetrado) solo en el fondo de la página — las tarjetas
       son opacas ({SURFACE}) así que la cuadrícula solo asoma en los márgenes/espacios entre ellas. */
    .stApp {{
        background-color: {BG};
        background-image:
            linear-gradient(rgba(33,33,33,0.05) 1px, transparent 1px),
            linear-gradient(90deg, rgba(33,33,33,0.05) 1px, transparent 1px);
        background-size: 26px 26px;
    }}
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

    /* ---- BOTTOM NAV (celular y tablet, ambas orientaciones) ----
       900px solo cubría tablet en vertical; en horizontal (~1024-1194px, iPad/Android típico)
       caía en el layout de escritorio con la nav de arriba apretada contra el padding por
       defecto de Streamlit — de ahí se veía "apretada" incluso en "modo escritorio". */
    .st-key-bottomnav {{ display: none; }}
    @media (max-width: 1180px) {{
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
        /* Streamlit deja ~5rem de aire a cada lado por defecto (clase ".main" ya no existe en
           esta versión, por eso el selector viejo nunca aplicaba) — en una tablet eso se come
           casi el 20% del ancho útil y hace ver todo más apretado de lo que hace falta. */
        [data-testid="stMainBlockContainer"] {{
            padding-left: 1.25rem !important; padding-right: 1.25rem !important; padding-bottom: 76px !important;
        }}
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
    /* st.expander (tarjetas de "Perforaciones y muestras", "Historial de Cambios", etc.) trae de
    fábrica un borde casi invisible (20% de opacidad) y sin sombra ni fondo propio — se pierde
    contra el fondo cuadriculado de la app. Se le da el mismo tratamiento de tarjeta que a los
    st.container(border=True) de arriba, para que se distinga igual de bien. */
    [data-testid="stExpander"] details {{
        border-radius: 12px !important; border: 1px solid {BORDER} !important;
        box-shadow: 0 1px 4px rgba(11,28,48,0.08) !important; background: {SURFACE} !important;
    }}
    /* Barra de título del expander (ej. "S1 — Sondeo · 2 muestra(s)") en verde, para que cada
    perforación se distinga de un vistazo dentro de la lista. */
    [data-testid="stExpander"] summary {{
        background: {SECONDARY_CONTAINER} !important; border-radius: 11px 11px 0 0 !important;
    }}
    [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {{
        color: {PRIMARY} !important;
    }}
    /* Las filas "etiqueta + campo" de los formularios de ensayo (humedad, límites, granulometría...)
       arman el layout con st.columns para que etiqueta y campo queden lado a lado — pero Streamlit
       las apila solo (etiqueta arriba, campo abajo a todo el ancho) en pantallas angostas, que es
       justo lo que se ve "apachurrado"/mal alineado en tablet. Se fuerza a que sigan en fila dentro
       de las tarjetas; en celular muy angosto (420px) se deja el apilado nativo como respaldo.
       min-width:0 es clave: sin eso, la columna de la etiqueta no se encoge por debajo del ancho
       de su texto más largo y la fila se desborda de la tarjeta en vez de acomodarse (la etiqueta
       larga como "Masa suelo seco + recipiente (g) (16 hrs)" empuja el campo fuera de pantalla).
       Ojo: 0 puro también aplastaba las columnas angostas de badges/botones ("Abrir", el badge de
       Estado) en tablas como "Ensayos asignados" — el botón/badge se encogía por debajo de su
       contenido y el texto quedaba amontonado. Esas columnas (con :has) se excluyen del encogido:
       mantienen su tamaño natural y son las columnas de texto/etiqueta las que ceden espacio.
       Los encabezados de tabla (.assigned-th, ej. "ID PROYECTO") NO se excluyen del encogido —
       su columna tiene que encogerse igual que la columna de datos de abajo (misma proporción),
       si no, encabezado y dato quedan desalineados. En vez de eso, se le quita el nowrap más abajo
       para que el texto pase a dos líneas dentro del fondo verde en vez de desbordarse cortado.
       :not(.st-key-home-actions): la fila de tarjetas de Inicio también tiene
       data-test-scroll-behavior="normal" (no es exclusivo de tarjetas con borde, como se asumía
       antes) y quedaba atrapada por estas reglas de "no encoger" pensadas para tablas — con más
       especificidad que las reglas propias de home-actions más abajo, ganaban ellas y la fila de
       Inicio nunca podía hacer wrap (se quedaba desbordada, exigiendo scroll horizontal). */
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:not(.st-key-home-actions) [data-testid="stHorizontalBlock"] {{
        flex-wrap: nowrap !important;
    }}
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:not(.st-key-home-actions) [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {{
        min-width: 0 !important;
    }}
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:not(.st-key-home-actions) [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:has(.stButton),
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:not(.st-key-home-actions) [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:has(.badge) {{
        min-width: fit-content !important;
        flex-shrink: 0 !important;
    }}
    /* La regla de arriba (:has) protege la columna de un botón/badge en la fila de DATOS, pero la
       fila de ENCABEZADO (.assigned-th, sin botón ni badge) no calzaba con ese :has y se encogía
       distinto → encabezado y dato quedaban desalineados (ej. "ESTADO"/"ACCIÓN" corridos respecto
       a su columna real). Se protegen las últimas 2 columnas por posición, no por contenido, y solo
       dentro de las tarjetas que son tablas (tienen algún .assigned-th) para no tocar los formularios
       de 2 columnas etiqueta+campo. */
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:has(.assigned-th)
        [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-last-child(-n+2) {{
        min-width: fit-content !important;
        flex-shrink: 0 !important;
    }}
    @media (max-width: 420px) {{
        div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"] [data-testid="stHorizontalBlock"] {{
            flex-wrap: wrap !important;
        }}
    }}
    /* Acento azul a la izquierda solo en tarjetas de contenido — se excluyen la barra de
       navegación, las tarjetas de notificaciones y las cajas donde se digita información
       (esas no son "otra tarjeta más", son campos de captura). home-actions tampoco: es un
       contenedor que envuelve varias tarjetas propias (cada una ya tiene su estilo, gradiente
       o borde), no una tarjeta en sí — el acento le quedaba como una barra verde de borde a
       borde tapando el grupo entero en vez de resaltar una sola tarjeta. bell-alert/bell-quiet
       (el botón-contenedor de la campana en el topbar) tampoco: es un ícono, no una tarjeta —
       se veía como una rayita verde suelta pegada a la campana. */
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"]:not(.st-key-topbar):not(.st-key-topbar-nav):not(.st-key-bottomnav):not(.st-key-notif-popover-body):not(.st-key-fab-new-project):not(.st-key-muestra-obs-box):not(.st-key-home-actions):not(.st-key-bell-alert):not(.st-key-bell-quiet):not([class*="st-key-notif-card-"]) {{
        border-left: 4px solid {PRIMARY} !important;
    }}
    /* Tarjetas de "Ensayos asignados" (una por ensayo): más separación entre sí y sombra
       más marcada para que no se confundan entre ellas ni con el fondo de la página. */
    div[data-testid="stVerticalBlock"][data-test-scroll-behavior="normal"][class*="st-key-ensayo-card-"] {{
        margin-bottom: 12px !important;
        box-shadow: 0 2px 6px rgba(11,28,48,0.10) !important;
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
    .status-circle-primary {{ background: {SECONDARY_CONTAINER}; color: {PRIMARY}; }}

    /* Línea de tiempo del historial de la muestra (marcador + línea conectora + contenido) */
    .timeline-item {{ display: flex; gap: 14px; }}
    .timeline-marker-col {{ display: flex; flex-direction: column; align-items: center; }}
    .timeline-line {{ width: 2px; flex: 1; min-height: 14px; background: {BORDER}; margin: 4px 0; }}
    .timeline-content {{ flex: 1; padding-bottom: 22px; }}
    .timeline-item:last-child .timeline-content {{ padding-bottom: 2px; }}
    .timeline-titulo {{ font-weight: 700; font-size: 14px; color: {TEXT}; }}
    .timeline-actor {{ font-size: 13px; color: {PRIMARY}; margin-top: 1px; }}
    .timeline-fecha {{ font-size: 12px; color: {MUTED}; margin-top: 4px; }}

    /* Tarjeta con acento a la izquierda, para encabezados de detalle (ej. Detalle de Muestra) */
    .st-key-muestra-header-card {{ border-left: 4px solid {PRIMARY} !important; }}
    .st-key-muestra-obs-box .stTextArea textarea {{ background-color: {SECONDARY_CONTAINER} !important; }}
    /* Botones dentro de la campana de notificaciones — por defecto salen blancos/planos y se
       pierden contra el fondo del popover. */
    .st-key-notif-popover-body .stButton button {{
        background-color: {SECONDARY_CONTAINER}; color: {PRIMARY}; border-color: {SECONDARY_CONTAINER};
    }}
    .st-key-notif-popover-body .stButton button:hover {{
        background-color: {PRIMARY}; color: {SURFACE}; border-color: {PRIMARY};
    }}
    /* Campana: roja mientras haya notificaciones sin leer, para que se note de un vistazo. */
    .st-key-bell-alert button {{
        background-color: {DANGER} !important; color: {SURFACE} !important; border-color: {DANGER} !important;
    }}
    .st-key-bell-alert button:hover {{
        background-color: {DANGER_LIGHT} !important; color: {DANGER} !important; border-color: {DANGER} !important;
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
    .stDateInput input, [data-testid="stDateInputField"],
    .stSelectbox > div > div, .stMultiSelect > div > div {{
        background-color: {SURFACE} !important;
        border: 1px solid {BORDER} !important;
        border-radius: 8px !important;
    }}
    /* El date_input de Streamlit ya no usa un <input> normal, sino un grupo de "spinbuttons"
       (día/mes/año) sin caja propia — sin esto se ve como texto plano sobre el fondo gris. */
    [data-testid="stDateInputField"] {{ padding: 8px 12px !important; }}
    .stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus,
    [data-testid="stDateInputField"]:focus-within {{
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

    /* Tarjetas de la fila de acciones de Inicio, todas del mismo alto y alineadas.
       El wrap nativo de Streamlit para st.columns() se activa según el ancho de la ventana, no
       según el espacio real disponible en esta fila — en una tablet ancha (pero no tan ancha como
       para que quepan cómodas 3 tarjetas con su texto) no llegaba a activarse y la fila se
       desbordaba, obligando a hacer scroll horizontal para ver la tercera tarjeta. Se le da a cada
       columna un ancho mínimo propio y se fuerza el wrap explícitamente: sobran 3 en pantallas
       anchas, se acomodan 2+1 o 1+1+1 en las angostas. */
    .st-key-home-actions [data-testid="stHorizontalBlock"] {{
        align-items: stretch; flex-wrap: wrap !important; row-gap: 16px;
    }}
    .st-key-home-actions [data-testid="stColumn"] {{
        flex: 1 1 260px !important; min-width: 260px !important; width: auto !important;
    }}
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
    /* OJO: estas 4 clases se usan sueltas (Actividad reciente, Ensayos asignados, Buscar...)
    fuera de cualquier ".activity-table" — ese wrapper ya no existe en ningún lado del código,
    así que iban sin aplicar (texto a tamaño por defecto, más grande de lo pensado) hasta que
    se les quitó el prefijo muerto. */
    .cell-id {{ font-family: 'JetBrains Mono', monospace; color: {PRIMARY}; font-weight: 800; font-size: 14px; }}
    .cell-title {{ font-weight: 600; color: {TEXT}; font-size: 14px; }}
    .cell-sub {{ font-size: 11.5px; color: {NEUTRAL}; margin-top: 1px; }}
    .cell-muted {{ color: {NEUTRAL}; font-size: 12px; }}
    .activity-footer {{
        display: flex; justify-content: space-between; align-items: center; padding: 10px 14px;
        color: {NEUTRAL}; font-size: 13px;
    }}

    /* ---- ENSAYOS ASIGNADOS (panel de Laboratorista) ---- */
    .assigned-th {{
        background: {SECONDARY_CONTAINER}; color: {PRIMARY}; font-family: 'JetBrains Mono', monospace;
        font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
        padding: 8px 10px; border-radius: 6px; margin-bottom: 4px; white-space: normal;
        overflow-wrap: break-word; line-height: 1.3;
    }}
    .assigned-chip {{
        background: {BG}; border: 1px solid {BORDER}; color: {PRIMARY}; font-size: 12px; font-weight: 700;
        padding: 3px 10px; border-radius: 6px; display: inline-block;
    }}
    /* Variantes de color para la columna "Ensayos asignados" de la tabla de muestras — cada
       chip refleja el estado de ESE ensayo puntual (verde finalizado / amarillo en proceso /
       rojo sin iniciar), reemplazando la columna "Estado" aparte que mostraba solo un estado
       agregado de toda la muestra. */
    .assigned-chip-success {{ background: {SUCCESS_LIGHT}; border-color: {SUCCESS}; color: {SUCCESS}; }}
    .assigned-chip-warning {{ background: {WARNING_LIGHT}; border-color: {WARNING}; color: {WARNING}; }}
    .assigned-chip-danger {{ background: {DANGER_LIGHT}; border-color: {DANGER}; color: {DANGER}; }}
    /* Versión compacta para la columna "Ensayos asignados" de la lista de muestras: uno debajo
       del otro (Humedad, Granulometría, Límites...) en vez de lado a lado — así la columna solo
       necesita el ancho del chip más largo, no el de los 3 juntos, y le deja espacio de sobra a
       "Tipo"/"Profundidad" al lado para no partirse en varias líneas. */
    .assigned-chip-sm {{ font-size: 10px; padding: 2px 6px; white-space: nowrap; }}
    .assigned-chip-row {{ display: flex; flex-direction: column; align-items: flex-start; gap: 3px; }}

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
    @media (max-width: 1180px) {{
        .st-key-fab-new-project {{ bottom: 92px; right: 16px; }}
    }}
</style>
""", unsafe_allow_html=True)

# Los campos numéricos de la app (pesos, temperaturas, lecturas de tamiz, etc.) son todos
# st.text_input con placeholder "0.00" / "0.0" / "0" / "001" — Streamlit no tiene un widget de
# texto con teclado numérico nativo, así que se marca por JS (vía un iframe de components.html,
# que sí puede tocar el DOM del documento padre por ser mismo origen) usando esa convención de
# placeholder para poner inputmode="decimal"/"numeric" solo ahí. El resto de campos de texto
# (nombre, dirección, correo...) no calzan el patrón y se quedan con el teclado normal.
# El mismo script hace que Enter salte al siguiente campo de texto en vez de quedarse quieto
# (útil digitando tamiz tras tamiz) — el setTimeout deja que Streamlit primero registre el
# valor tecleado (su propio manejador de Enter) antes de mover el foco.
# También hace que el gesto de "atrás" (botón del navegador, o el gesto nativo de Android)
# funcione igual que el botón "← Atrás" de la app — ver _push_history_entry() más abajo, que es
# quien realmente dispara el pushState (justo después de que navigate()/go_back() terminan de
# sincronizar la URL). Este bloque solo pone el listener de popstate: al presionar atrás, dispara
# un recargo completo de la página, que Python resuelve leyendo la URL ya restaurada por el
# navegador (init_state) — por eso esto necesita ir de la mano con la restauración de sesión por
# cookie, si no cada "atrás" mandaría de vuelta al login.
components.html("""
<script>
(function() {
    if (window.parent.__geodeltaInputModeInit) return;
    window.parent.__geodeltaInputModeInit = true;

    function applyInputMode() {
        var inputs = window.parent.document.querySelectorAll(
            'div[data-testid="stTextInput"] input[type="text"]'
        );
        inputs.forEach(function(input) {
            var ph = (input.getAttribute('placeholder') || '').trim();
            var mode = null;
            if (/^-?\\d+[.,]\\d+$/.test(ph)) {
                mode = 'decimal';
            } else if (/^\\d[\\d\\s]*$/.test(ph)) {
                mode = 'numeric';
            }
            if (mode) {
                if (input.getAttribute('inputmode') !== mode) input.setAttribute('inputmode', mode);
            } else if (input.hasAttribute('inputmode')) {
                input.removeAttribute('inputmode');
            }
        });
    }

    function focusNextOnEnter(e) {
        if (e.key !== 'Enter') return;
        var target = e.target;
        if (!target || !target.matches || !target.matches('div[data-testid="stTextInput"] input')) return;
        var inputs = Array.prototype.slice.call(
            window.parent.document.querySelectorAll('div[data-testid="stTextInput"] input')
        );
        var idx = inputs.indexOf(target);
        if (idx === -1 || idx === inputs.length - 1) return;
        var next = inputs[idx + 1];
        setTimeout(function() {
            next.focus();
            next.select();
        }, 0);
    }

    applyInputMode();
    new MutationObserver(applyInputMode).observe(window.parent.document.body, {childList: true, subtree: true});
    window.parent.document.addEventListener('keydown', focusNextOnEnter, true);
    window.parent.addEventListener('popstate', function() {
        window.parent.location.reload();
    });
})();
</script>
""", height=0)

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

ASSAY_LABELS = {"granulometria": "Granulometría", "humedad": "Contenido de humedad", "masa-unitaria": "Peso unitario", "limites": "Límites de Atterberg", "pasa200": "Pasa 200", "cbr": "CBR", "corte-directo": "Corte Directo", "gravedad-especifica": "Gravedad específica", "proctor": "Proctor", "materia-organica": "Materia orgánica", "limite-contraccion": "Límite de contracción", "consolidacion": "Consolidación", "compresion-inconfinada": "Compresión inconfinada", "compresion-roca": "Compresión en roca", "carga-puntual": "Carga puntual", "solidez-sulfatos": "Solidez en sulfatos", "terrones-arcilla": "Terrones de arcilla"}
NORMAS_ENSAYO = {
    "granulometria": ["INV-214-13", "INV.E-213-13", "INV.E 123-13"],
    "humedad": ["INV E-122", "ASTM D2216"],
    "masa-unitaria": ["INV E-202", "ASTM D1188"],
    "cbr": ["INV E-148", "ASTM D1883"],
    "corte-directo": ["INV E-154", "ASTM D3080"],
    "gravedad-especifica": ["INV E-222", "INV E-223", "INV E-128", "ASTM C128", "ASTM C127", "ASTM D854"],
    "proctor": ["INV E-141-13", "INV E-142-13"],
    "materia-organica": ["INV E-121-13", "ASTM D2974"],
    "limite-contraccion": ["INV E-129-13", "ASTM D4943"],
    "consolidacion": ["INV E-151-13", "ASTM D2435"],
    "compresion-inconfinada": ["INV E-152-13", "ASTM D2166"],
    "compresion-roca": ["ASTM D7012"],
    "carga-puntual": ["ASTM D5731"],
    "solidez-sulfatos": ["INV E-220-13", "ASTM C88"],
    "terrones-arcilla": ["INV E-211-13"],
}
STATUS_LABELS = {"sin-iniciar": "Sin iniciar", "en-proceso": "En proceso", "finalizado": "Finalizado"}
STATUS_BADGE = {"sin-iniciar": "badge-danger", "en-proceso": "badge-warning", "finalizado": "badge-success"}
STATUS_ICON = {"sin-iniciar": "radio_button_unchecked", "en-proceso": "autorenew", "finalizado": "check_circle"}

TIPO_PERFORACION_PREFIX = {"Sondeo": "S", "Apique": "AP", "Fuente/Cantera": "F"}
# Texto que espera la lista desplegable de "tipo de perforación" en la plantilla
# CLASIFICACION_DE_SUELOS.xlsm (celda D12, validada contra AG6:AG10: SONDEO/APIQUE/TRINCHERA/NQ/N.A.).
TIPO_PERFORACION_EXCEL = {"Sondeo": "SONDEO", "Apique": "APIQUE", "Fuente/Cantera": "CANTERA"}
TIPO_MUESTRA_OPTIONS = ["Shelby", "NQ", "SS", "SPT", "Lona", "Bolsa", "CBR", "N/A"]
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
# Peso unitario Método B (bitácora GDA-FLC-029 y plantilla GDA-FLC-030 — por volumen conocido: altura y diámetro
# de la muestra, en vez de parafinado).
EQUIPO_MASA_UNITARIA_B = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Horno GDA-E-007", "Horno GDA-E-404",
                         "Calibrador pie de rey GDA-E-110"]
MUB_HUMEDAD_FILAS = [
    ("mub_hum_recipiente", "Recipiente No."), ("mub_hum_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("mub_hum_seco_17", "Masa suelo seco + recipiente (g) (17 horas)"), ("mub_hum_seco_18", "Masa suelo seco + recipiente (g) (18 horas)"),
    ("mub_hum_seco_19", "Masa suelo seco + recipiente (g) (19 horas)"), ("mub_hum_masa_rec", "Masa del recipiente (g)"),
]

# Mismos campos que MUB_HUMEDAD_FILAS, pero para Peso Unitario Parafinado cuando la muestra no
# tiene un ensayo de Humedad asignado (prefijo mu_hum_ en vez de mub_hum_).
MU_HUMEDAD_FILAS = [
    ("mu_hum_recipiente", "Recipiente No."), ("mu_hum_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("mu_hum_seco_17", "Masa suelo seco + recipiente (g) (17 horas)"), ("mu_hum_seco_18", "Masa suelo seco + recipiente (g) (18 horas)"),
    ("mu_hum_seco_19", "Masa suelo seco + recipiente (g) (19 horas)"), ("mu_hum_masa_rec", "Masa del recipiente (g)"),
]

# Equipos reales del CBR (INV E-148 / ASTM D1883), tal como aparecen en el formato físico
# "EQUIPOS UTILIZADOS", en el mismo orden (fila por fila).
EQUIPO_CBR = [
    "Horno GDA-E-007", "Balanza GDA-E-010", "Martillo GDA-E-387", "Tamiz N°4 GDA-E-038",
    "Horno GDA-E-404", "Balanza GDA-E-011", "Martillo GDA-E-111", "Tamiz 3/8\" GDA-E-037",
    "Pie de rey GDA-E-110", "Balanza GDA-E-012", "Celda de carga GDA-E-016", "Tamiz 3/4\" GDA-E-035",
    "Máquina MULT GDA-E-008", "Deformímetro GDA-E-086", "Celda de carga GDA-E-017",
]

# Equipos usados en Corte Directo (INV E-154), leídos del formato físico en papel que se llena a
# mano — todavía no hay una plantilla de Excel oficial conectada a este ensayo (ver
# render_corte_directo_form), así que estos campos pueden necesitar ajuste cuando se consiga esa
# plantilla, igual que le pasó a CBR con su primera versión armada solo a partir de una captura.
EQUIPO_CORTE_DIRECTO = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Horno GDA-E-007",
                        "Horno GDA-E-404", "Máquina de corte GDA-E-001", "Máquina de corte GDA-E-002"]
EQUIPO_CORTE_HUMEDAD = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Horno GDA-E-007", "Horno GDA-E-404"]

# Proctor (INV E-141 / E-142) — bitácora GDA-FL-008, que comparte hoja con el CBR. Solo los equipos
# que se usan en el Proctor (se dejan fuera los del CBR: celdas de carga, máquina MULT, deformímetro).
EQUIPO_PROCTOR = ["Horno GDA-E-007", "Horno GDA-E-404", "Balanza GDA-E-010", "Balanza GDA-E-011", "Balanza GDA-E-012",
                  "Martillo GDA-E-387", "Martillo GDA-E-111", "Tamiz N°4 GDA-E-038", "Tamiz 3/8\" GDA-E-037",
                  "Tamiz 3/4\" GDA-E-035", "Pie de rey GDA-E-110"]
# CBR de suelos compactados que va junto al Proctor (lado derecho de la bitácora GDA-FL-008 y de la
# plantilla GDA-FLC-002): 3 moldes de 10, 25 y 56 golpes. Es un CBR APARTE del ensayo "CBR"
# (inalterado, plantilla GDA-FLC-013) — sus datos viven dentro del ensayo de Proctor.
CBRC_MOLDES = 3
CBRC_GOLPES = ["10", "25", "56"]
CBRC_FILAS = [
    ("molde", "Molde No."), ("capas", "No. de capas"), ("masa_muestra_molde", "Masa muestra + molde (g)"),
    ("masa_molde", "Masa del molde (g)"), ("altura", "Altura de la muestra (cm)"), ("diametro", "Diámetro de la muestra (cm)"),
]
CBRC_HUM_FILAS = [
    ("recipiente", "Recipiente No."), ("masa_humedo", "Masa húmeda + recipiente (g)"),
    ("seco_16h", "Masa seca + recipiente (g) (16 horas)"), ("seco_17h", "Masa seca + recipiente (g) (17 horas)"),
    ("seco_18h", "Masa seca + recipiente (g) (18 horas)"), ("seco_19h", "Masa seca + recipiente (g) (19 horas)"),
    ("masa_recipiente", "Masa del recipiente (g)"),
]
CBRC_EXP_FILAS = [("exp_inicial", "Lectura inicial expansión (in)"), ("exp_final", "Lectura final expansión (in)")]
# Filas de la plantilla (columna W/AA/AE) para las 12 profundidades de CBR_PENETRACION_FILAS.
CBRC_PEN_FILAS_EXCEL = [22, 23, 24, 25, 26, 28, 30, 31, 33, 34, 35, 37]
# Materia orgánica (INV E-121) — armado a partir de la plantilla GDA-FLC-003 (no hay bitácora en papel).
MO_METODOS = ["A", "B"]
MO_TEMPERATURAS = ["110 °C", "60 °C"]
MO_CAMPOS = [
    ("mo_masa_crisol_seco", "Masa del crisol + muestra de suelo seco (g)"),
    ("mo_masa_crisol_ignicion", "Masa del crisol + muestra de suelo después de ignición (g)"),
    ("mo_masa_crisol", "Masa del crisol (g)"),
]
# Límite de contracción (INV E-129, método de la parafina) — plantilla GDA-FLC-022.
LC_DATOS_RECIPIENTE = [("lc_volumen", "Volumen del recipiente (cm³)"), ("lc_masa_recipiente", "Masa del recipiente (g)")]
LC_DATOS_PASTILLA = [
    ("lc_masa_aire", "Masa en el aire (g)"), ("lc_masa_aire_parafinada", "Masa en el aire parafinada (g)"),
    ("lc_masa_sumergida_parafinada", "Masa sumergida parafinada (g)"),
]
LC_HUMEDAD_FILAS = [("hum_humedo", "Masa suelo húmedo + recipiente (g)"), ("hum_seco", "Masa del suelo seco + recipiente (g)"),
                    ("hum_recipiente", "Masa del recipiente (g)")]
# Consolidación unidimensional (INV E-151) — bitácora GDA-FL-017 y plantilla GDA-FLC-009. Las lecturas de
# deformación por carga/tiempo salen de la máquina (hojas "DATOS MAQUINA" y de cada carga): no se digitan acá.
CONS_CONDICIONES = ["Inalterada", "Compactada", "Remoldada"]
CONS_TEMP_SECADO = ["60 °C", "110 °C"]
CONS_METODOS = ["A", "B"]
CONS_CARGAS = ["0.5", "1.0", "2.0", "4.0", "8.0", "16.0", "32.0"]
# (etiqueta, clave inicial, clave final). En "Masa muestra + anillo" cada columna lleva su valor; en el resto la
# columna final repite la inicial mientras no se digite otro valor.
CONS_DATOS_MUESTRA = [
    ("Masa muestra + anillo (g)", "cons_masa_anillo_muestra_ini", "cons_masa_anillo_muestra_fin"),
    ("Masa del anillo (g)", "cons_masa_anillo", "cons_fin_masa_anillo"),
    ("Diámetro del anillo (mm)", "cons_diametro", "cons_fin_diametro"),
    ("Altura del anillo (mm)", "cons_altura", "cons_fin_altura"),
]
CONS_HUMEDAD_FILAS = [
    ("recipiente", "Recipiente No."), ("humedo", "Masa muestra húmeda + recipiente (g)"),
    ("seco_17", "Masa suelo seco + recipiente (g) (17 horas)"), ("seco_18", "Masa suelo seco + recipiente (g) (18 horas)"),
    ("seco_19", "Masa suelo seco + recipiente (g) (19 horas)"), ("recipiente_masa", "Masa del recipiente (g)"),
]
CONS_GS_CAMPOS = [
    ("cons_pic_no", "Picnómetro No. (1 a 6)"),
    ("cons_pic_masa_agua_suelo", "Masa picnómetro + agua + suelo a T °C (g)"),
    ("cons_pic_temp", "Temperatura de aforo (°C)"),
    ("cons_pic_masa_seco", "Masa suelo seco (g)"),
]
# Cargas de la hoja "DATOS MAQUINA " de la plantilla: (encabezado, columna del tiempo en min, columna donde se pega
# la deformación de la máquina). El tiempo ya viene en la plantilla; el bloque de 32 kg comparte la rejilla del de 16 kg.
CONS_MAQ_BLOQUES = [("0,25 kg", "B", "C"), ("0,5 kg", "E", "F"), ("1,0 kg", "H", "I"), ("2,0 kg", "K", "L"),
                    ("4,0 kg", "N", "O"), ("8,0 kg", "Q", "R"), ("16 kg", "T", "U"), ("32 kg", "T", "X")]
CONS_CONSOLIDOMETROS = ["", "GDA-E-003", "GDA-E-004", "GDA-E-005", "GDA-E-006", "GDA-E-393", "GDA-E-394"]
EQUIPO_CONSOLIDACION = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Termómetro GDA-E-126", "Horno GDA-E-007", "Horno GDA-E-404"]
# Calibración de los picnómetros (hoja Resultados de la plantilla): masa del picnómetro lleno de agua
# = a·T + b, con T en °C.
CONS_PIC_CALIBRACION = {1: (-0.1428, 688.64), 2: (-0.1158, 694.37), 3: (-0.0614, 349.31),
                        4: (-0.0655, 344.51), 5: (-0.0652, 344.03), 6: (-0.061, 355.9)}
# Compresión inconfinada (INV E-152) — bitácora GDA-FL-005 y plantilla GDA-FLC-008 (.xlsm).
CI_DEFORMACIONES = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300,
                    330, 360, 390, 420, 450, 480, 510, 540, 650, 700, 750, 800, 900, 1000]  # 0.001 in
CI_FILA_MAQUINA, CI_MAX_MAQUINA = 26, 33  # los datos de la máquina empiezan en la fila del cero (26)
CI_FILAS_EXCEL = list(range(27, 59))  # la plantilla trae 32 filas de lectura (la 26 es el cero)
CI_CONDICIONES = ["Inalterada", "Compactada", "Remodelada"]
CI_MUESTREOS = ["Tubo Shelby", "SPT", "NQ - Barrena", "HQ - Barrena"]
CI_FALLAS = ["PLANO INCLINADO", "ABOMBAMIENTO", "CONO Y GRIETAS VERTICALES", "DESMORONAMIENTO"]
CI_HUMEDAD_FILAS = [
    ("ci_hum_recipiente", "Recipiente No."), ("ci_hum_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("ci_hum_seco_17", "Masa suelo seco + recipiente (g) (17 horas)"), ("ci_hum_seco_18", "Masa suelo seco + recipiente (g) (18 horas)"),
    ("ci_hum_seco_19", "Masa suelo seco + recipiente (g) (19 horas)"), ("ci_hum_masa_rec", "Masa del recipiente (g)"),
]
EQUIPO_COMPRESION_INCONFINADA = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Máquina multiusos GDA-E-008", "Máquina manual GDA-E-014",
                                 "Horno GDA-E-007", "Horno GDA-E-404", "Pie de rey GDA-E-028"]
PROCTOR_METODOS = ["A", "B", "C"]
PROCTOR_TAMICES = ["", "3/4\"", "3/8\"", "No. 4"]
PROCTOR_TAMIZ_EXCEL = {"3/4\"": "¾ \"", "3/8\"": "⅜ \"", "No. 4": "N.4"}  # lista desplegable de L25
PROCTOR_PRUEBAS = 4
PROCTOR_FILAS = [
    ("golpes", "No. de golpes"), ("molde", "Molde No."), ("capas", "No. de capas"),
    ("masa_humedo_molde", "Masa de la muestra húmeda + molde (g)"), ("masa_molde", "Masa molde (g)"),
    ("volumen_molde", "Volumen del molde (cm³)"),
]
PROCTOR_HUMEDAD_FILAS = [
    ("hum_recipiente", "Recipiente No."), ("hum_masa_humedo", "Masa recipiente + muestra húmeda (g)"),
    ("hum_seco_16h", "Masa recipiente + muestra seca (g) (16 horas)"), ("hum_seco_17h", "Masa recipiente + muestra seca (g) (17 horas)"),
    ("hum_seco_18h", "Masa recipiente + muestra seca (g) (18 horas)"), ("hum_seco_19h", "Masa recipiente + muestra seca (g) (19 horas)"),
    ("hum_masa_recipiente", "Masa del recipiente (g)"),
]
EQUIPO_CORTE_GRAVEDAD = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Termómetro GDA-E-126", "Baño de María GDA-E-009"]
CORTE_GRAVEDAD_CAMPOS = [
    ("masa_suelo_seco", "Masa del suelo seco (g)"),
    ("masa_pic_muestra", "Masa pic. + muestra aforado a temp. (g)"),
    ("masa_pic", "Masa pic. aforado a temp. (g)"),
    ("temperatura", "Temperatura (°C)"),
]

# Gravedad Específica — dos variantes en el formato físico: la que pasa el tamiz No. 4 (finos,
# con picnómetro) y la que lo retiene (gruesos, material sumergido).
EQUIPO_GESP_FINOS = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Horno GDA-E-007", "Picnómetro",
                     "Termómetro GDA-E-116", "Baño de María GDA-E-009"]
EQUIPO_GESP_GRUESOS = ["Balanza GDA-E-012", "Balanza GDA-E-011", "Horno GDA-E-007",
                       "Termómetro GDA-E-116", "Horno GDA-E-404"]
GESP_OPCIONES = ["Pasa el tamiz No. 4 (finos)", "Retiene el tamiz No. 4 (gruesos)", "Ambos",
                 "Arcillas y limos (INV E-128)"]
GESP_ARCILLA_CAMPOS = [
    ("gesp_a_picnometro", "Picnómetro No."),
    ("gesp_a_masa_aire", "Masa en el aire de la muestra (g)"),
    ("gesp_a_masa_pic_agua", "Masa del picnómetro lleno de agua (g)"),
    ("gesp_a_masa_pic_muestra", "Masa del picnómetro aforado + muestra (g)"),
    ("gesp_a_temp", "Temperatura del ensayo (°C)"),
    ("gesp_a_pct_retenido", "% material retenido en el tamiz No. 4"),
]
GESP_FINOS_CAMPOS = [
    ("gesp_f_masa_seco", "Masa del suelo seco (g)"),
    ("gesp_f_masa_sss", "Masa de la muestra saturada superficialmente seca (g)"),
    ("gesp_f_masa_pic_agua_suelo", "Masa del picnómetro + agua + suelo a la temperatura del ensayo (g)"),
    ("gesp_f_masa_pic_agua", "Masa del picnómetro + agua a la temperatura del ensayo (g)"),
    ("gesp_f_temp", "Temperatura del ensayo (°C)"),
    ("gesp_f_pct_retenido", "% material retenido en el tamiz No. 4"),
]
GESP_GRUESOS_CAMPOS = [
    ("gesp_g_masa_sss", "Masa del material saturado superficialmente seco (g)"),
    ("gesp_g_masa_sumergido", "Masa del material sumergido en agua (g)"),
    ("gesp_g_masa_seco", "Masa del material seco (g)"),
    ("gesp_g_temp", "Temperatura del ensayo (°C)"),
]

# ════════════════════════════════════════════════════════════════════
# DESCRIPCIÓN VISUAL ESTRUCTURADA (menús desplegables en vez de texto libre) — para poder
# comparar de un vistazo la clasificación que hace el laboratorista a ojo con la clasificación
# USCS/AASHTO que calcula la app a partir de los datos de Granulometría/Límites (ver
# clasificar_uscs/clasificar_aashto). Todas las opciones van en MAYÚSCULA (así se escribe en el
# informe oficial) y siguen la nomenclatura estándar de descripción visual-manual de suelos
# (INV E-102 / ASTM D2488), no una lista inventada. No hay campo de texto libre — la frase final
# se arma solo con lo elegido en estos menús (ver descripcion_visual_estructurada).
DESC_TIPO_SUELO_OPTIONS = ["", "GRAVA", "ARENA", "LIMO", "ARCILLA", "ORGÁNICO", "OTROS"]
# Componente secundario ("grava CON ALGO DE arena", "arcilla CON ALGO DE arena") — mismas
# opciones que el tipo principal (sin la casilla en blanco, se filtra el tipo principal ya
# elegido para no dejar escoger "grava con algo de grava"). Es opcional: la casilla
# "¿tiene un componente secundario?" decide si se muestra o no.
DESC_TIPO_SECUNDARIO_OPTIONS = ["GRAVA", "ARENA", "LIMO", "ARCILLA", "ORGÁNICO", "OTROS"]
# Color principal y subtonalidad van separados (antes venían mezclados como "Café oscuro"): el
# color base es el matiz dominante, la subtonalidad es el matiz que lo modifica (ej. "MARRÓN
# ROJIZO", "GRIS AMARILLENTO") — así se puede combinar cualquier color con cualquier matiz en vez
# de tener que enumerar cada combinación como una opción aparte.
DESC_COLOR_OPTIONS = ["", "GRIS", "MARRÓN", "AMARILLO", "ROJO", "NEGRO", "BLANCO", "BEIGE", "NARANJA", "VERDE"]
DESC_SUBTONALIDAD_OPTIONS = ["", "CLARO", "OSCURO", "ROJIZO", "AMARILLENTO", "VERDOSO", "GRISÁCEO", "BLANQUECINO"]
# Suelos gruesos (grava/arena) se describen por cementación; suelos finos (limo/arcilla/orgánico)
# por consistencia — son dos propiedades distintas, no dos escalas de lo mismo, por eso solo se
# despliega una u otra según el tipo de grano elegido (ver _es_grueso).
DESC_CEMENTACION_OPTIONS = ["", "DÉBIL", "MODERADA", "FUERTE"]
DESC_CONSISTENCIA_OPTIONS = ["", "MUY BLANDA", "BLANDA", "FIRME", "DURA", "MUY DURA"]
# Forma solo aplica a grava (partículas lo bastante grandes para juzgar su forma a ojo);
# angulosidad aplica a cualquier suelo grueso (grava o arena).
DESC_FORMA_OPTIONS = ["", "PLANAS", "ALARGADAS", "PLANAS Y ALARGADAS"]
DESC_ANGULOSIDAD_OPTIONS = ["", "ANGULOSA", "SUB ANGULOSA", "SUB REDONDEADA", "REDONDEADA"]
DESC_HUMEDAD_OPTIONS = ["", "SECA", "HÚMEDA", "SATURADA"]


def _es_grueso(tipo_suelo):
    """Grava/Arena = grano grueso (cementación, angulosidad); Limo/Arcilla/Orgánico = grano fino
    (consistencia). Ver INV E-102."""
    return tipo_suelo in ("GRAVA", "ARENA")


def descripcion_visual_estructurada(muestra, tipo_override=None):
    """Arma la frase legible en mayúscula (ej. 'LIMO DE COLOR MARRÓN ROJIZO DE CONSISTENCIA DURA
    EN CONDICIÓN SECA') a partir de los menús desplegables de la muestra — se usa tanto en la
    vista de solo lectura como en el Excel oficial. None si todavía no se ha elegido ninguna
    opción.

    `tipo_override`: reemplaza SOLO la palabra inicial (ej. "GRAVA ARCILLOSA" en vez de "GRAVA")
    sin tocar el resto de la frase — para armar la versión "según la clasificación calculada" (ver
    descripcion_visual_calculada) sin duplicar toda esta lógica. El resto de la frase (forma,
    angulosidad, cementación o consistencia) se sigue decidiendo con el tipo de grano que
    realmente eligió el laboratorista a ojo, no con el override: son justo los campos que él
    mismo llenó bajo ese tipo, y cambiar de escala a mitad de la frase dejaría campos vacíos."""
    # .upper(): datos de antes de este cambio (migración 0015) se guardaron en minúscula/mixta
    # (ej. "Limo", "Café oscuro") — se normalizan al leer para que la frase salga toda en
    # mayúscula igual que los datos nuevos, sin tener que migrar filas viejas en la base de datos.
    tipo = (muestra.get("desc_tipo_suelo") or "").upper() or None
    partes = [tipo_override or tipo] if (tipo_override or tipo) else []

    # Componente secundario (ej. "GRAVA CON ALGO DE ARENA", "ARCILLA CON ALGO DE ARENA") — un
    # segundo tipo de grano que también está presente en la muestra, sin ser el dominante.
    secundario = (muestra.get("desc_tipo_secundario") or "").upper() or None
    if secundario:
        partes.append(f"CON ALGO DE {secundario}")

    angulosidad = (muestra.get("desc_angulosidad") or "").upper() if _es_grueso(tipo) else None
    if angulosidad:
        partes.append(angulosidad)
    forma = (muestra.get("desc_forma") or "").upper() if tipo == "GRAVA" else None
    if forma:
        partes.append(f"DE FORMA {forma}")

    color = (muestra.get("desc_color") or "").upper() or None
    subtonalidad = (muestra.get("desc_subtonalidad") or "").upper() or None
    if color:
        partes.append(f"DE COLOR {color}" + (f" {subtonalidad}" if subtonalidad else ""))
    elif subtonalidad:
        partes.append(f"DE COLOR {subtonalidad}")

    if _es_grueso(tipo):
        cementacion = (muestra.get("desc_cementacion") or "").upper() or None
        if cementacion:
            partes.append(f"CON CEMENTACIÓN {cementacion}")
    else:
        consistencia = (muestra.get("desc_consistencia") or "").upper() or None
        if consistencia:
            partes.append(f"DE CONSISTENCIA {consistencia}")

    # "CONDICIÓN {ESTADO}", no "CONDICIÓN DE HUMEDAD {ESTADO}" — el estado (SECA/HÚMEDA/
    # SATURADA) ya funciona como adjetivo de "condición", "condición de humedad húmeda" suena
    # redundante.
    humedad = (muestra.get("desc_humedad") or "").upper() or None
    if humedad:
        partes.append(f"EN CONDICIÓN {humedad}")

    return " ".join(partes) if partes else None


def descripcion_visual_para_excel(muestra):
    """Texto que va al campo "DESCRIPCIÓN VISUAL" del Excel oficial: la frase armada de los
    menús desplegables — None si todavía no hay ninguna opción elegida, para que el llamador
    pueda caer al respaldo de siempre (observaciones del ensayo, o el tipo de muestra). Si la
    muestra viene de antes de quitar el campo de notas libres (ver migración 0016), esas notas
    viejas se siguen mostrando como respaldo en vez de perderse."""
    return descripcion_visual_estructurada(muestra) or (muestra.get("descripcion_visual") or "").strip() or None


def descripcion_visual_calculada(muestra):
    """Segunda versión de la descripción visual, con el tipo de suelo que de verdad salió en la
    clasificación USCS (calculada con los datos ya digitados de Granulometría/Límites, ver
    clasificar_uscs) en vez del tipo que el laboratorista eligió a ojo antes de tener esos datos
    — combinado con el resto de características, que sí siguen siendo juicio visual del
    laboratorista (color, cementación o consistencia, humedad...). Se muestra AL LADO de
    descripcion_visual_estructurada, no en su lugar — la inicial a ojo se conserva tal cual.
    None mientras no haya datos suficientes para calcular la USCS."""
    gran_assay = get_assay(muestra["id_unico"], "granulometria")
    if not gran_assay:
        return None
    lim_assay = get_assay(muestra["id_unico"], "limites")
    resultado = clasificar_uscs(gran_assay.get("data"), lim_assay.get("data") if lim_assay else None)
    simbolo = resultado.get("simbolo")
    if not simbolo:
        return None
    nombre = USCS_NOMBRES.get(simbolo, simbolo).upper()
    return descripcion_visual_estructurada(muestra, tipo_override=nombre)

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
    "Peso unitario", "Gravedad específica", "Consolidación", "Corte CD", "Corte CU", "Corte UU",
    "Carga puntual", "Solidez en sulfatos", "Terrones de arcilla", "Alargamiento y aplanamiento",
    "Caras fracturadas", "Azul de metileno", "Desgaste", "Micro deval", "Angularidad",
    "Expansión Lambe", "Hidrometría", "Otro",
]
SUPPORTED_ASSAY_MAP = {
    "Granulometría": "granulometria", "Humedad": "humedad", "Peso unitario": "masa-unitaria",
    "Límites de Atterberg": "limites", "Pasa 200": "pasa200", "CBR": "cbr",
    # Las 3 casillas de la bitácora (Corte CD/CU/UU) comparten un solo ensayo de "Corte Directo"
    # con un selector de modo adentro (ver render_corte_directo_form) — no son 3 formularios
    # separados. Si una muestra tiene más de una marcada, "Abrir" en cualquiera de las 3 lleva al
    # mismo ensayo, exactamente igual que Granulometría/Pasa 200 comparten datos hoy.
    "Corte CD": "corte-directo", "Corte CU": "corte-directo", "Corte UU": "corte-directo",
    "Gravedad específica": "gravedad-especifica",
    "Corte Directo": "corte-directo",
    "Proctor": "proctor",
    "Materia orgánica": "materia-organica",
    "Límite de contracción": "limite-contraccion",
    "Consolidación": "consolidacion",
    "Compresión inconfinada": "compresion-inconfinada",
    "Compresión en roca": "compresion-roca",
    "Carga puntual": "carga-puntual",
    "Solidez en sulfatos": "solidez-sulfatos",
    "Terrones de arcilla": "terrones-arcilla",
}


def unificar_ensayos(labels):
    """Las casillas Corte CD/CU/UU comparten un solo ensayo (ver SUPPORTED_ASSAY_MAP) — en las
    listas se muestra una sola entrada "Corte Directo" aunque la bitácora tenga varias marcadas;
    el modo (CD/CU/UU) se elige adentro del ensayo. El resto de labels queda igual."""
    vistos, resultado = set(), []
    for e in labels:
        tipo = SUPPORTED_ASSAY_MAP.get(e)
        if tipo == "corte-directo":
            e = "Corte Directo"
        if e in vistos:
            continue
        vistos.add(e)
        resultado.append(e)
    return resultado


BITACORA_BASE_COLS = ["Número", "Prof. De", "Prof. A", "Tipo de muestra"] + BITACORA_ENSAYOS + ["Observaciones"]

# Celdas de la plantilla oficial GDA-FL-003 (hoja "S1"): columna por ensayo, checkbox de
# norma y checkbox de tipo de perforación.
BITACORA_XLSX_ENSAYO_COL = {
    "Granulometría": "F", "Pasa 200": "G", "Humedad": "H", "Límites de Atterberg": "I",
    "Límite de contracción": "J", "Materia orgánica": "K", "Proctor": "L", "CBR": "M",
    "Compresión inconfinada": "N", "Compresión en roca": "O", "Peso unitario": "P",
    "Gravedad específica": "Q", "Consolidación": "R", "Corte CD": "S", "Corte CU": "T",
    "Corte UU": "U", "Carga puntual": "V", "Solidez en sulfatos": "W", "Terrones de arcilla": "X",
    "Alargamiento y aplanamiento": "Y", "Caras fracturadas": "Z", "Azul de metileno": "AA",
    "Desgaste": "AB", "Micro deval": "AC", "Angularidad": "AD", "Expansión Lambe": "AE",
    "Hidrometría": "AF", "Otro": "AG",
}
BITACORA_XLSX_NORMA_CELL = {"IDU": "AG10", "INVIAS": "AI10", "NTC": "AG12", "Otro": "AI12"}
BITACORA_XLSX_TIPO_CELL = {"Sondeo": "H14", "Apique": "D14", "Fuente/Cantera": "AH14"}
BITACORA_XLSX_MAX_ROWS = 14  # la plantilla trae 14 filas fijas (18 a 31)


# ════════════════════════════════════════════════════════════════════
# CALIBRACIÓN DE BALANZAS (Jefe de Laboratorio) — GDA-FLC-029 · GDA-LABI-002
# La pantalla es el HTML aprobado por el Jefe de Laboratorio, incrustado TAL CUAL como componente de
# Streamlit (balanzas_component/, generado con tools/build_balanzas_component.py) — el diseño y los
# cálculos en pantalla son los suyos. La app solo le agrega el puente: le manda el registro guardado y
# recibe los cambios para guardarlos en Supabase. Los cálculos de abajo (verificados contra las
# fórmulas reales de VERIFICACION DE BALANZAS.xlsx, hoja "Balanzas", filas 11-31: Error = indicación −
# carga de referencia; Resultado según |error| <= EMP) solo se usan para el estado guardado y la
# bitácora en Excel.
BALANZAS = [
    {"codigo": "GDA-E-010", "nombre": "Balanza Cap. 600g", "marca": "TRUMAX", "serie": "MIX-H", "resolucion": "0,01"},
    {"codigo": "GDA-E-011", "nombre": "Balanza Cap. 3200 g Pionner", "marca": "OHAUS", "serie": "Px3202/E", "resolucion": "0,01"},
    {"codigo": "GDA-E-012", "nombre": "Balanza Cap. 30 Kg", "marca": "TRUMAX", "serie": "FENIX", "resolucion": "1"},
    {"codigo": "GDA-E-013", "nombre": "Balanza Cap. 3000 g", "marca": "TS", "serie": "T200", "resolucion": "0,1"},
]
BALANZAS_POR_CODIGO = {b["codigo"]: b for b in BALANZAS}   # respaldo si la tabla `balanzas` (migración 0032) aún no existe


def _bal_lista():
    """Catálogo de balanzas (tabla `balanzas`, se administra desde la pantalla). Si la migración 0032 todavía
    no se corrió, se usan las 4 de siempre para que la pantalla siga funcionando."""
    lista = st.session_state.get("balanzas")
    if lista is None:
        return [dict(b, id=None, activa=True) for b in BALANZAS]
    return lista


def _bal_por_codigo(codigo):
    return next((b for b in _bal_lista() if b["codigo"] == codigo), None)


def _bal_codigos_selector():
    """Balanzas que se ofrecen en el selector: las activas y, para poder consultar su historial, las que están de baja
    pero tienen registros. Activas primero."""
    con_registros = {c["codigo_equipo"] for c in st.session_state.get("balance_checks", [])}
    visibles = [b for b in _bal_lista() if b.get("activa", True) or b["codigo"] in con_registros]
    visibles.sort(key=lambda b: (not b.get("activa", True), b["codigo"]))
    return [b["codigo"] for b in visibles]
BAL_CONDICIONES_PREVIAS = [
    ("cond_limpieza", "Limpieza del receptor y del entorno"),
    ("cond_nivelacion", "Nivelación verificada (burbuja centrada)"),
    ("cond_cero_tara", "Cero / tara estable antes de cargar"),
    ("cond_estabilizacion", "Tiempo de estabilización cumplido"),
    ("cond_masas_limpias", "Masas patrón limpias y en buen estado"),
]


def _bal_resultado(error, emp):
    """Mismo criterio que la plantilla: "Pendiente" si falta la indicación o el ±EMP (o el EMP
    es 0/negativo), si no "Cumple"/"No cumple" según |error| <= EMP."""
    if error is None or emp is None or emp <= 0:
        return "Pendiente"
    return "Cumple" if abs(error) <= emp else "No cumple"


def resultados_bal_excentricidad(data):
    carga = to_float(data.get("exc_carga_usada"))
    filas = []
    for p in range(1, 6):
        ind = to_float(data.get(f"exc_p{p}_indicacion"))
        emp = to_float(data.get(f"exc_p{p}_emp"))
        error = (ind - carga) if (ind is not None and carga is not None) else None
        filas.append({"punto": p, "indicacion": ind, "emp": emp, "error": error,
                       "resultado": _bal_resultado(error, emp)})
    indicaciones = [f["indicacion"] for f in filas if f["indicacion"] is not None]
    errores = [f["error"] for f in filas if f["error"] is not None]
    error_maximo = max((abs(e) for e in errores), default=None)
    diferencia_max = (max(indicaciones) - min(indicaciones)) if len(indicaciones) == 5 else None
    return filas, error_maximo, diferencia_max


def resultados_bal_repetibilidad(data):
    carga = to_float(data.get("rep_carga_usada"))
    limite_r = to_float(data.get("rep_limite_r"))
    filas = []
    for p in range(1, 6):
        ind = to_float(data.get(f"rep_r{p}_indicacion"))
        emp = to_float(data.get(f"rep_r{p}_emp"))
        error = (ind - carga) if (ind is not None and carga is not None) else None
        filas.append({"punto": p, "indicacion": ind, "emp": emp, "error": error,
                       "resultado": _bal_resultado(error, emp)})
    indicaciones = [f["indicacion"] for f in filas if f["indicacion"] is not None]
    rango_r = (max(indicaciones) - min(indicaciones)) if len(indicaciones) == 5 else None
    if rango_r is None or limite_r is None or limite_r <= 0:
        resultado_r = "Pendiente"
    else:
        resultado_r = "Cumple" if rango_r <= limite_r else "No cumple"
    return filas, rango_r, resultado_r


def resultados_bal_exactitud(data):
    filas = []
    for p in range(1, 9):
        aplica = data.get(f"exact_p{p}_aplica", True)
        carga = to_float(data.get(f"exact_p{p}_carga"))
        asc = to_float(data.get(f"exact_p{p}_asc"))
        desc = to_float(data.get(f"exact_p{p}_desc"))
        emp = to_float(data.get(f"exact_p{p}_emp"))
        error_asc = (asc - carga) if (asc is not None and carga is not None) else None
        error_desc = (desc - carga) if (desc is not None and carga is not None) else None
        if not aplica:
            resultado_asc = resultado_desc = "No aplica"
        else:
            resultado_asc, resultado_desc = _bal_resultado(error_asc, emp), _bal_resultado(error_desc, emp)
        filas.append({"punto": p, "aplica": aplica, "carga": carga, "asc": asc, "desc": desc, "emp": emp,
                       "error_asc": error_asc, "error_desc": error_desc,
                       "resultado_asc": resultado_asc, "resultado_desc": resultado_desc})
    return filas


def decision_sugerida_balanza(data):
    """"apto"/"no-apto"/None(pendiente) según los resultados de las 3 pruebas: cualquier
    "No cumple" → no apto; todo "Cumple" sin pendientes → apto; si falta algo, pendiente."""
    resultados = []
    exc_filas, _, _ = resultados_bal_excentricidad(data)
    resultados += [f["resultado"] for f in exc_filas]
    rep_filas, _, resultado_r = resultados_bal_repetibilidad(data)
    resultados += [f["resultado"] for f in rep_filas] + [resultado_r]
    for f in resultados_bal_exactitud(data):
        if f["aplica"]:
            resultados += [f["resultado_asc"], f["resultado_desc"]]
    if any(r == "No cumple" for r in resultados):
        return "no-apto"
    if resultados and all(r == "Cumple" for r in resultados):
        return "apto"
    return None


def _bal_lunes(d):
    """Lunes de la semana ISO de `d` (date)."""
    return d - timedelta(days=d.weekday())


# ════════════════════════════════════════════════════════════════════
# IMPORTAR BITÁCORA DE ORDEN DESDE EXCEL (plantilla oficial GDA-FL-003 ya
# diligenciada, ej. por el cliente) — lee el mismo mapeo de celdas de arriba,
# en sentido inverso a generar_excel_bitacora_orden().
# ════════════════════════════════════════════════════════════════════
def _hoja_es_bitacora_orden(ws):
    """True si esta hoja tiene la forma de la plantilla GDA-FL-003 — para no intentar leer
    como bitácora una hoja cualquiera de un archivo que no es esta plantilla."""
    texto = " ".join(str(v) for v in (ws["AH1"].value, ws["G1"].value) if v).upper()
    return "GDA-FL-003" in texto or ("BIT" in texto and "ORDEN" in texto)


def _celda_marcada(valor):
    """Cualquier contenido no vacío cuenta como 'marcado' — a mano la gente pone X, x, un
    visto, 'si', lo que sea; no necesariamente la 'X' exacta que pone la app al exportar."""
    return valor is not None and str(valor).strip() != ""


def _normalizar_numero_muestra(valor):
    """El cliente a veces ya escribe 'M1'/'m-1' en la columna NRO. DE MUESTRA — como la app le
    antepone su propia 'M-' en toda pantalla donde se muestra el número de una muestra (ej. la
    tabla "Ver muestras"), importar ese valor tal cual se veía duplicado ('M-M1'). Se le quita el
    prefijo M de una sola letra al importar para que quede solo el número/código, igual que si
    alguien lo hubiera digitado a mano sin ese prefijo."""
    s = str(valor).strip()
    # Solo se quita si justo después del prefijo viene un dígito (ej. "M1", "m-23") — si no, es
    # más probable que la "M" sea parte real del código (ej. "MA6") y no un prefijo a quitar.
    m = re.match(r"^[Mm][-.\s]*(\d.*)$", s)
    return m.group(1).strip() if m else s


def _fila_sin_ensayos(fila):
    """True si esta fila no trae ningún ensayo marcado — esas filas no se importan: sin ensayo
    asignado no hay nada que un laboratorista pueda digitar, así que solo estorbarían en la
    bitácora en vez de ayudar."""
    return not any(fila.get(e) for e in BITACORA_ENSAYOS)


def parse_bitacora_orden_xlsx(nombre_archivo, file_bytes):
    """Lee un archivo .xlsx ya diligenciado con la plantilla oficial GDA-FL-003 y devuelve
    (perforaciones, advertencias, encabezado).

    Un archivo "de fábrica" trae un solo sondeo/apique en su hoja "S1", pero en la práctica no
    todos los clientes mandan los archivos igual (uno por sondeo, o varias hojas duplicadas y
    renombradas dentro del mismo archivo) — por eso se recorren TODAS las hojas del archivo y se
    toma cualquiera que tenga la forma de esta plantilla (_hoja_es_bitacora_orden).

    perforaciones: lista de dicts {tipo, codigo, filas}. 'filas' ya viene en el mismo formato de
    columnas que usa el editor de la bitácora (BITACORA_BASE_COLS), lista para pd.DataFrame(...).
    encabezado: datos del encabezado de la primera hoja válida (nombre, localización...), solo
    informativos, para que el Jefe compare contra el proyecto seleccionado antes de guardar —
    no se usan para crear ni editar el proyecto."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=True)
    except Exception:
        return [], [f"{nombre_archivo}: no se pudo abrir como Excel (¿el archivo está dañado o no es .xlsx?)."], None

    hojas_validas = [ws for ws in wb.worksheets if _hoja_es_bitacora_orden(ws)]
    if not hojas_validas:
        return [], [f"{nombre_archivo}: ninguna hoja parece ser la plantilla GDA-FL-003 (Bitácora Orden) — se omitió."], None

    perforaciones, advertencias, encabezado = [], [], None
    for ws in hojas_validas:
        if encabezado is None:
            encabezado = {
                "nombre": ws["F10"].value or "", "localizacion": ws["E12"].value or "",
                "numero_anio": ws["AG8"].value or "",
            }

        tipo_hoja = next((t for t, celda in BITACORA_XLSX_TIPO_CELL.items() if _celda_marcada(ws[celda].value)), None)

        # Se agrupa por el código que trae CADA fila en la columna N.° DE PERFORACIÓN, no por un
        # solo código "dominante" para toda la hoja — algunos clientes usan una sola hoja para
        # listar varias perforaciones distintas, una por fila (ej. apiques AP-1, AP-2, AP-3...),
        # no solo el uso "de fábrica" de una hoja = un sondeo con varias muestras de profundidad.
        filas_por_codigo, sin_codigo, sin_ensayo = {}, [], 0
        for i in range(BITACORA_XLSX_MAX_ROWS):
            r = 18 + i
            numero = ws[f"B{r}"].value
            if numero is None or str(numero).strip() == "":
                continue
            perf_r = ws[f"A{r}"].value
            perf_r = str(perf_r).strip() if perf_r and str(perf_r).strip() else None
            fila = {
                "Número": _normalizar_numero_muestra(numero),
                "Prof. De": to_float(ws[f"D{r}"].value, 0.0),
                "Prof. A": to_float(ws[f"E{r}"].value, 0.0),
                "Tipo de muestra": (str(ws[f"C{r}"].value).strip() if _celda_marcada(ws[f"C{r}"].value)
                                    else TIPO_MUESTRA_OPTIONS[0]),
            }
            for label, col in BITACORA_XLSX_ENSAYO_COL.items():
                fila[label] = _celda_marcada(ws[f"{col}{r}"].value)
            fila["Observaciones"] = ws[f"AH{r}"].value or ""
            # Una muestra sin ningún ensayo marcado no se importa — no hay nada que digitar.
            if _fila_sin_ensayos(fila):
                sin_ensayo += 1
                continue
            (filas_por_codigo.setdefault(perf_r, []) if perf_r else sin_codigo).append(fila)

        if sin_ensayo:
            advertencias.append(f"{nombre_archivo} — hoja '{ws.title}': {sin_ensayo} muestra(s) sin ningún "
                                 "ensayo marcado no se importaron.")

        if not filas_por_codigo and not sin_codigo:
            if not sin_ensayo:
                advertencias.append(f"{nombre_archivo} — hoja '{ws.title}': no se encontró ninguna muestra "
                                     "(columna NRO. DE MUESTRA vacía) — se omitió.")
            continue

        if sin_codigo:
            codigo_default = max(filas_por_codigo, key=lambda c: len(filas_por_codigo[c])) if filas_por_codigo else ws.title
            filas_por_codigo.setdefault(codigo_default, []).extend(sin_codigo)
            advertencias.append(f"{nombre_archivo} — hoja '{ws.title}': alguna(s) fila(s) no traían código en "
                                 f"la columna N.° DE PERFORACIÓN — se agruparon bajo '{codigo_default}', revísalo.")

        for codigo, filas_codigo in filas_por_codigo.items():
            tipo = tipo_hoja or next((t for t, prefix in TIPO_PERFORACION_PREFIX.items()
                                       if codigo.upper().startswith(prefix)), None)
            if not tipo:
                tipo = "Sondeo"
                advertencias.append(f"{nombre_archivo} — hoja '{ws.title}' ({codigo}): no se marcó Sondeo/"
                                     "Apique/Fuente-Cantera — se asumió Sondeo, revísalo antes de guardar.")
            perforaciones.append({"tipo": tipo, "codigo": codigo, "filas": filas_codigo})

    return perforaciones, advertencias, encabezado


# ════════════════════════════════════════════════════════════════════
# CARGA DE DATOS DESDE SUPABASE (una vez por rerun, en el enrutador principal)
#
# El resto de la app sigue leyendo st.session_state.projects/perforaciones/
# muestras/assays exactamente con la misma forma que tenían en el store en
# memoria (perforaciones/muestras siguen agrupadas por código interno, los
# ensayos siguen usando el id_unico de la muestra como "muestra_id") — así
# que ninguno de los ~50 sitios de lectura repartidos por la app tuvo que
# cambiar. Lo único que cambia es de dónde viene el dato (Supabase, no un
# dict en memoria) y que ahora SÍ sobrevive un reinicio del servidor.
# Los sitios de ESCRITURA sí cambiaron: mutan vía db.py y luego hacen
# st.rerun(), que dispara una nueva llamada a _load_data() con el dato fresco.
# ════════════════════════════════════════════════════════════════════
def _load_data(solo_balanzas=False):
    """Carga desde Supabase lo que las pantallas necesitan. `solo_balanzas`: la pantalla de Calibración
    de Balanzas no usa proyectos/perforaciones/muestras/ensayos (la barra superior solo usa las
    notificaciones), así que no se piden — en cada navegación eso son 4 consultas menos y se acorta el
    rato en que Streamlit deja el contenido de la pantalla anterior atenuado."""
    if solo_balanzas:
        notifications = db.list_notifications(st.session_state.role) if st.session_state.role else []
        for n in notifications:
            n["role"] = n["target_role"]
            n["muestra_id"] = n.get("muestra_id_unico")
        st.session_state.notifications = notifications
        st.session_state.balance_checks = db.list_balance_checks() if st.session_state.role == "jefe" else []
        try:
            st.session_state.balanzas = db.list_balanzas() if st.session_state.role == "jefe" else None
            st.session_state.pop("_bal_sin_tabla", None)
        except Exception:   # la migración 0032 todavía no se corrió
            st.session_state.balanzas = None
            st.session_state["_bal_sin_tabla"] = True
        return
    projects = db.list_projects()
    proj_by_id = {p["id"]: p for p in projects}
    st.session_state.projects = projects

    perfs_raw = db.list_all_perforaciones()
    perf_by_id = {p["id"]: p for p in perfs_raw}
    perforaciones = {}
    for p in perfs_raw:
        proj = proj_by_id.get(p["project_id"])
        if proj:
            perforaciones.setdefault(proj["codigo_interno"], []).append(p)
    st.session_state.perforaciones = perforaciones

    muestras_raw = db.list_all_muestras()
    muestra_by_id = {m["id"]: m for m in muestras_raw}
    muestras = {}
    for m in muestras_raw:
        perf = perf_by_id.get(m["perforacion_id"])
        proj = proj_by_id.get(perf["project_id"]) if perf else None
        if perf and proj:
            key = f"{proj['codigo_interno']}::{perf['codigo']}"
            muestras.setdefault(key, []).append(m)
    st.session_state.muestras = muestras

    assays = []
    for a in db.list_all_assays():
        muestra = muestra_by_id.get(a["muestra_id"])
        perf = perf_by_id.get(muestra["perforacion_id"]) if muestra else None
        proj = proj_by_id.get(perf["project_id"]) if perf else None
        if not (muestra and perf and proj):
            continue
        a = dict(a)
        a["muestra_id"] = muestra["id_unico"]
        a["codigo_interno"] = proj["codigo_interno"]
        a["perforacion_codigo"] = perf["codigo"]
        a["muestra_numero"] = muestra["numero"]
        a["lastModified"] = a["updated_at"]
        a["createdAt"] = a["created_at"]
        assays.append(a)
    st.session_state.assays = assays

    notifications = db.list_notifications(st.session_state.role) if st.session_state.role else []
    for n in notifications:
        n["role"] = n["target_role"]
        n["muestra_id"] = n.get("muestra_id_unico")
    st.session_state.notifications = notifications

    # Solo el Jefe de Laboratorio usa Calibración de Balanzas (RLS también lo exige) — no pedirle
    # esto a Supabase para los otros roles, que igual no podrían leerlo.
    st.session_state.balance_checks = db.list_balance_checks() if st.session_state.role == "jefe" else []


# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL
# ════════════════════════════════════════════════════════════════════
SESSION_COOKIE_MAX_AGE_DAYS = 30

# OJO: Streamlit Community Cloud filtra casi todas las cookies en su capa de proxy antes de que
# lleguen al backend de la app — st.context.cookies (lo que se usaba antes acá) da un dict vacío
# una vez desplegado, aunque funcione perfecto corriendo local. Por eso la sesión nunca se
# restauraba al recargar la página en producción, solo en las pruebas locales. CookieManager
# (extra_streamlit_components) evita el problema porque lee la cookie con JS en el navegador y
# se la manda a Python como el valor de un componente — nunca pasa por el proxy del servidor.
# NO envolver esto en @st.fragment: cuando el componente entrega su valor real de forma asíncrona,
# el rerun automático que dispara Streamlit quedaría acotado al fragmento y el resto del script
# (_try_restore_session_from_cookie, más abajo) nunca se enteraría del valor nuevo.
cookie_manager = stx.CookieManager(key="gdl_cookie_manager")


def _set_session_cookie(access_token, refresh_token):
    """Guarda los tokens de la sesión de Supabase en una cookie del navegador (nunca en la
    URL — la pantalla actual sí va en la URL, ver _sync_query_params, pero el token no) para
    poder restaurar el login solo tras un recargo de página o una reconexión, en vez de
    mandar siempre a la persona de vuelta a pedirle código+clave (ver init_state)."""
    expira = datetime.now() + timedelta(days=SESSION_COOKIE_MAX_AGE_DAYS)
    cookie_manager.batch_set({"gdl_at": access_token, "gdl_rt": refresh_token}, expires_at=expira)


def _delete_cookie_safe(nombre, key):
    """cookie_manager.delete() hace `del self.cookies[nombre]` sin verificar que exista, y
    truena con KeyError si ya no está (p. ej. la cookie nunca llegó a existir en este navegador,
    o el componente todavía no había entregado su valor real cuando se llamó) — se verifica antes
    para que cerrar sesión no reviente si no hay nada que borrar."""
    if nombre in cookie_manager.cookies:
        cookie_manager.delete(nombre, key=key)


def _clear_session_cookie():
    _delete_cookie_safe("gdl_at", key="del_gdl_at")
    _delete_cookie_safe("gdl_rt", key="del_gdl_rt")


def _set_remember_user_cookie(codigo):
    """Guarda el código de usuario (no la clave) en una cookie aparte de la de sesión, para
    precargar el campo "Código de usuario" del login la próxima vez que haga falta iniciar
    sesión — checkbox "Recordar mi usuario" en render_login()."""
    expira = datetime.now() + timedelta(days=SESSION_COOKIE_MAX_AGE_DAYS)
    cookie_manager.set("gdl_user", codigo, key="set_gdl_user", expires_at=expira)


def _clear_remember_user_cookie():
    _delete_cookie_safe("gdl_user", key="del_gdl_user")


def _push_history_entry():
    """Convierte la navegación que acaba de terminar en una entrada de historial de verdad
    (ver navigate()/go_back() y el router principal, que llaman a esto en el rerun SIGUIENTE
    al que cambió de pantalla — no en el mismo, por la misma razón que _set_session_cookie no
    se llama justo antes de un st.rerun(): el iframe no alcanza a montarse y correr su script
    antes de que el próximo rerun lo reemplace). Para cuando esto corre, _sync_query_params()
    ya terminó de aplicar sus 6 asignaciones en un rerun previo y completo, así que la URL del
    navegador ya es la definitiva — no hace falta ninguna espera ni detección de estabilidad.
    Se vio en pruebas que esto se termina llamando 2 veces por una sola navegación (probable
    doble rerun del click del botón) — se compara contra la última URL empujada para no dejar
    una entrada de historial duplicada, que obligaría a presionar "atrás" dos veces seguidas
    para moverse una sola pantalla."""
    components.html("""
    <script>
    (function() {
        var href = window.parent.location.href;
        if (window.parent.__geodeltaLastPushedHref === href) return;
        window.parent.__geodeltaLastPushedHref = href;
        window.parent.history.pushState({geodelta: true}, '', href);
    })();
    </script>
    """, height=0)


def _sync_query_params():
    """Guarda la pantalla actual en la URL (nunca el token de sesión — eso vive en una
    cookie aparte, ver _set_session_cookie). Al cerrar sesión se limpia todo, así que la
    sesión solo termina cuando el usuario le da a "Cerrar sesión" o la cookie vence (30 días).
    Se actualizan los 6 parámetros de una sola vez con .update(): asignarlos uno por uno
    (st.query_params["x"] = y) manda un ForwardMsg — y dispara una re-sincronización del
    navegador — por cada asignación; con 6 asignaciones seguidas eso se notó como hasta 7
    entradas de historial por cada navegación (ver _push_history_entry). .update() está
    documentado en el propio Streamlit para mandar un solo mensaje."""
    if st.session_state.role:
        st.query_params.update({
            "screen": st.session_state.screen,
            "codigo": st.session_state.selected_codigo or "",
            "perf": st.session_state.selected_perforacion or "",
            "muestra": st.session_state.selected_muestra_id or "",
            "assay": st.session_state.selected_assay_id or "",
            "atipo": st.session_state.selected_assay_type or "",
        })
    else:
        st.query_params.clear()


def init_state():
    if "initialized" in st.session_state:
        return
    st.session_state.initialized = True
    st.session_state.role = None
    st.session_state.profile = None
    st.session_state.screen = "home"

    st.session_state.projects = []
    st.session_state.perforaciones = {}
    st.session_state.muestras = {}
    st.session_state.assays = []
    st.session_state.notifications = []
    st.session_state.balance_checks = []

    st.session_state.nav_stack = []
    st.session_state.bitacora_draft = {}
    st.session_state.draft_perforaciones = []
    st.session_state.draft_muestras = {}
    st.session_state.selected_codigo = ""
    st.session_state.selected_perforacion = ""
    st.session_state.selected_muestra_id = ""
    st.session_state.selected_assay_id = None
    st.session_state.selected_assay_type = None
    st.session_state.read_only_view = False

    # Restaura la posición de navegación desde la URL tras un recargo o reconexión (ver
    # _sync_query_params) — la persona vuelve a la misma pantalla en vez de a Inicio.
    st.session_state.screen = st.query_params.get("screen") or "home"
    st.session_state.selected_codigo = st.query_params.get("codigo") or ""
    st.session_state.selected_perforacion = st.query_params.get("perf") or ""
    st.session_state.selected_muestra_id = st.query_params.get("muestra") or ""
    st.session_state.selected_assay_id = st.query_params.get("assay") or None
    st.session_state.selected_assay_type = st.query_params.get("atipo") or None


def _try_restore_session_from_cookie():
    """Restaura el login desde la cookie del navegador (ver _set_session_cookie) — antes,
    cualquier recargo de página (F5, reconexión) mandaba de vuelta al login aunque la sesión de
    Supabase siguiera siendo válida. Si el refresh_token ya no sirve (venció, se cerró sesión en
    otro dispositivo), restore_session devuelve None y simplemente se queda en la pantalla de
    login, como antes.

    OJO: a diferencia del resto de init_state(), esto NO puede correr una sola vez al principio
    de la sesión — CookieManager lee la cookie de forma asíncrona con JS, así que en el primer
    rerun todavía no tiene el valor real (solo un default vacío) y recién lo entrega un par de
    reruns después. Por eso esta función se llama en CADA rerun mientras no haya sesión, hasta
    que la cookie real llegue. Una vez se intenta con un token real (funcione o no), no se
    reintenta más en esta sesión de navegador, para no golpear a Supabase en cada rerun si de
    verdad venció."""
    if st.session_state.role is not None or st.session_state.get("_cookie_restore_attempted"):
        return
    at = cookie_manager.get("gdl_at")
    rt = cookie_manager.get("gdl_rt")
    if not (at and rt):
        return
    st.session_state["_cookie_restore_attempted"] = True
    profile = db.restore_session(at, rt)
    if profile:
        st.session_state.profile = profile
        st.session_state.role = profile["role"]
        # restore_session() refresca el token si el access_token de la cookie ya había
        # vencido — Supabase rota el refresh_token en ese refresh, así que hay que
        # reescribir la cookie con los tokens nuevos (ver _tokens_rotated en el router
        # principal) o el siguiente recargo fallaría con un refresh_token ya inválido.
        st.session_state["_tokens_rotated"] = True


init_state()
_try_restore_session_from_cookie()


def _navegacion_muy_seguida():
    """True si navigate()/go_back() se acaba de llamar hace instantes (menos de 400ms) — guarda
    contra un solo toque que dispara el evento de click dos veces seguidas casi al mismo tiempo,
    algo que se ha visto con los widgets nuevos de esta versión de Streamlit en pantallas
    táctiles (ver también la investigación sobre checkboxes/date_input más arriba en el código).
    Reportado como pantallas/tarjetas repetidas al navegar (una parece quedar "pegada" encima de
    la otra) — la sospecha es que el segundo disparo del mismo toque arranca un rerun mientras el
    del primero todavía no había terminado de reflejarse. No se puede reproducir de forma
    confiable fuera de una tablet real, así que esto es una salvaguarda razonable, no una
    corrección confirmada de la causa exacta: si el segundo disparo llega casi pegado al primero,
    simplemente no hace nada — la persona solo tiene que volver a tocar si de verdad quería
    navegar dos veces en menos de 400ms (poco probable para un toque humano intencional)."""
    ahora = time.monotonic()
    si_seguida = (ahora - st.session_state.get("_last_navigate_ts", 0)) < 0.4
    st.session_state["_last_navigate_ts"] = ahora
    return si_seguida


def navigate(screen):
    if _navegacion_muy_seguida():
        return
    actual = st.session_state.get("screen")
    if actual and actual != screen:
        st.session_state.nav_stack.append(actual)
    st.session_state.screen = screen
    _sync_query_params()
    st.session_state._pending_history_push = True
    st.rerun()


def go_back(fallback="home"):
    """Vuelve a la pantalla realmente anterior (pila de navegación) en vez de un destino fijo."""
    if _navegacion_muy_seguida():
        return
    if st.session_state.nav_stack:
        st.session_state.screen = st.session_state.nav_stack.pop()
    else:
        st.session_state.screen = fallback
    _sync_query_params()
    st.session_state._pending_history_push = True
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
    return (f'<span class="status-circle {circle_class}" title="{html.escape(STATUS_LABELS[status])}">'
            f'{icon(STATUS_ICON[status], size=size, fill=True)}</span>')


APROBACION_INFO = {
    None: ("No Confirmada", "badge-danger"),
    "pendiente_ing": ("En Proceso", "badge-warning"),
    "aprobado": ("Confirmada", "badge-success"),
}


def aprobacion_badge_html(etapa):
    """Estado de aprobación (Jefe → Director Técnico) de un ensayo individual — independiente
    del semáforo de laboratorio (sin-iniciar/en-proceso/finalizado)."""
    label, clase = APROBACION_INFO.get(etapa, APROBACION_INFO[None])
    return f'<span class="badge {clase}">DT: {label}</span>'


def card_header_html(icon_name, title, extra_html=""):
    """Encabezado de tarjeta con ícono + título (y opcionalmente un badge a la derecha),
    usado en las tarjetas de los formularios de ensayo (Norma, Equipos, Pasa 200, etc.)."""
    return (f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;">'
            f'<div style="display:flex;align-items:center;gap:9px;font-weight:800;color:{PRIMARY};font-size:19px;">'
            f'{icon(icon_name, size=22)} {title}</div>{extra_html}</div>')


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
    """Notificación compartida entre todas las sesiones logueadas con `role` (tabla
    notifications en Supabase — ver _load_data)."""
    db.add_notification(role, mensaje, codigo_interno=codigo, perforacion_codigo=perf, muestra_id_unico=muestra_id)


def add_historial(obj, titulo, subtitulo="", icono="history", tono="muted"):
    """Registro de auditoría de un ensayo: cuándo se entregó, cuándo confirmó el Jefe, cuándo
    aprobó (o devolvió) el Director Técnico, etc. Se muestra como línea de tiempo (ver
    historial_timeline_html). `obj` es el dict de assay tal como lo devuelve get_assay."""
    db.add_historial(obj["id"], titulo, subtitulo, icono, tono)


def historial_timeline_html(historial):
    items = sorted(historial, key=lambda h: h["fecha"], reverse=True)
    filas = []
    for i, h in enumerate(items):
        linea = '<div class="timeline-line"></div>' if i < len(items) - 1 else ""
        marcador = (f'<div class="timeline-marker-col">'
                    f'<div class="status-circle status-circle-{h.get("tono", "muted")}">'
                    f'{icon(h.get("icono", "history"), size=17, fill=True)}</div>{linea}</div>')
        # Compatibilidad con entradas antiguas (formato previo: solo {"fecha","texto"}) que
        # puedan seguir en el store en memoria de una sesión anterior al rediseño del historial.
        titulo = h.get("titulo") or h.get("texto") or "Cambio registrado"
        actor_html = (f'<div class="timeline-actor">{html.escape(h["subtitulo"])}</div>'
                      if h.get("subtitulo") else "")
        filas.append(
            f'<div class="timeline-item">{marcador}'
            f'<div class="timeline-content"><div class="timeline-titulo">{html.escape(titulo)}</div>'
            f'{actor_html}<div class="timeline-fecha">{format_dt(h["fecha"])}</div></div></div>'
        )
    return "".join(filas)


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


@st.cache_data(ttl=30)
def list_laboratoristas():
    """Nombres de los laboratoristas activos, para el selector de asignación —
    reemplaza el texto libre de antes, que aceptaba cualquier cosa."""
    return sorted(
        p["full_name"] for p in db.list_profiles()
        if p["role"] == "laboratorista" and p.get("active", True)
    )


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
    (el laboratorista terminó) Y TODOS sus ensayos tienen el visto bueno final del Director
    Técnico (aprobación por ensayo individual) — no basta con que el laboratorio haya
    terminado, cada ensayo tiene que estar aprobado para poder entregarse."""
    counts = project_progress(codigo)
    total = sum(counts.values())
    if total == 0 or counts["finalizado"] != total:
        return "ejecucion"
    for a in st.session_state.assays:
        if a["codigo_interno"] == codigo and a.get("etapa_revision") != "aprobado":
            return "ejecucion"
    return "ejecutado"


def desarchivar_proyecto(codigo):
    """Reabre un proyecto ejecutado: revierte a 'en-proceso' todos sus ensayos finalizados,
    para que el laboratorista pueda volver a digitar o el Jefe agregar nuevas muestras/
    perforaciones. También limpia la aprobación de cada ensayo (etapa_revision y demás) —
    si se reabre, tiene que volver a pasar por Jefe y Director Técnico antes de poder archivarse
    de nuevo. El estado 'ejecutado' se recalcula solo (ver project_status). El llamador hace
    st.rerun() después, que recarga todo fresco desde Supabase (ver _load_data)."""
    for a in st.session_state.assays:
        if a["codigo_interno"] == codigo:
            db.reset_confirmacion(a["id"], reset_status=(a["status"] == "finalizado"))


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
            st.caption("Ingresa tu código de usuario y tu clave para acceder al sistema.")
            codigo_recordado = cookie_manager.get("gdl_user") or ""
            codigo = st.text_input("Código de usuario", value=codigo_recordado, placeholder="ej. jperez", autocomplete="off")
            password = st.text_input("Clave de acceso", type="password", placeholder="••••••••")
            recordar = st.checkbox("Recordar mi usuario", value=True)
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("INGRESAR", type="primary", use_container_width=True):
                if not codigo or not password:
                    st.error("Ingresa tu código de usuario y tu clave.")
                else:
                    try:
                        profile = db.sign_in(codigo, password)
                    except db.AuthError as e:
                        st.error(str(e))
                    else:
                        st.session_state.profile = profile
                        st.session_state.role = profile["role"]
                        # No se guarda la cookie aquí mismo: el st.rerun() de abajo corta la
                        # ejecución antes de que el iframe de components.html llegue a montarse
                        # y correr su script en el navegador (se probó y la cookie nunca quedaba
                        # puesta). Se guarda el token pendiente y se escribe la cookie en el
                        # siguiente rerun, cuando ya no hay un rerun inmediato después que lo corte.
                        st.session_state._pending_cookie_tokens = db.get_session_tokens()
                        st.session_state._pending_remember_user = codigo.strip() if recordar else ""
                        st.rerun()
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
            with st.container(key="bell-alert" if no_leidas else "bell-quiet"):
                with st.popover(str(no_leidas) if no_leidas else "", icon=":material/notifications:", use_container_width=True):
                    st.markdown("**Notificaciones**")
                    if not mis_notifs:
                        st.caption("No tienes notificaciones.")
                    else:
                        with st.container(key="notif-popover-body"):
                            if no_leidas and st.button("Marcar todas como leídas", key="notif_marcar_todas", use_container_width=True):
                                db.mark_all_notifications_read(st.session_state.role)
                                st.rerun()
                            for n in mis_notifs[:15]:
                                with st.container(border=True, key=f"notif-card-{n['id']}"):
                                    estilo_msg = "font-weight:700;" if not n["leida"] else f"font-weight:400;font-size:13px;color:{MUTED};"
                                    st.markdown(f'<div style="{estilo_msg}">{html.escape(n["mensaje"])}</div>'
                                                f'<div class="timestamp-caption">{format_dt(n["fecha"])}</div>', unsafe_allow_html=True)
                                    if n.get("muestra_id") and st.button("Ir a la muestra →", key=f"notif_go_{n['id']}", use_container_width=True):
                                        db.mark_notification_read(n["id"])
                                        st.session_state.selected_codigo = n["codigo_interno"]
                                        st.session_state.selected_perforacion = n["perforacion_codigo"]
                                        st.session_state.selected_muestra_id = n["muestra_id"]
                                        navigate("muestra-detail")
        with c_avatar:
            iniciales = ROLE_INICIALES.get(st.session_state.role, "LB")
            st.markdown(f'<div class="topbar-avatar">{iniciales}</div>', unsafe_allow_html=True)
        with c_logout:
            if st.button("", key="logout_top", help="Cerrar sesión", use_container_width=True, icon=":material/logout:"):
                db.sign_out()
                st.session_state.role = None
                st.session_state.profile = None
                st.session_state.nav_stack = []
                st.session_state._pending_logout_cookie_clear = True
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
def _ensayos_pendientes_dt():
    """Ensayos ya finalizados por el laboratorista, confirmados por el Jefe y esperando el
    visto bueno del Director Técnico (aprobación por ensayo individual). Devuelve tuplas
    (codigo, perf_codigo, muestra, ensayo_label)."""
    pendientes = []
    for p in st.session_state.projects:
        codigo = p["codigo_interno"]
        for perf in st.session_state.perforaciones.get(codigo, []):
            for m in st.session_state.muestras.get(f"{codigo}::{perf['codigo']}", []):
                for ensayo_label in unificar_ensayos([e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS]):
                    tipo_i = SUPPORTED_ASSAY_MAP.get(ensayo_label)
                    a = get_assay(m["id_unico"], tipo_i) if tipo_i else None
                    if a and a.get("etapa_revision") == "pendiente_ing":
                        pendientes.append((codigo, perf["codigo"], m, ensayo_label))
    return pendientes


def render_home():
    es_jefe = st.session_state.role == "jefe"
    es_ingeniero = st.session_state.role == "ingeniero"
    es_supervisor = es_jefe or es_ingeniero
    if es_jefe:
        st.markdown("## Bienvenido, Jefe de Laboratorio")
        st.caption("Resumen de operaciones y control de calidad geotécnica para hoy.")
    elif es_ingeniero:
        st.markdown("## Bienvenido, Director Técnico")
        st.caption("Revisión final y aprobación de muestras antes de entregarlas al cliente.")
    else:
        st.markdown("## Panel de Laboratorista")
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
            st.markdown("<br>", unsafe_allow_html=True)
            c4, _c5 = st.columns([1, 2])
            with c4:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("balance")}</div>'
                             '<div><h3>Calibración de balanzas</h3><p>Comprobación intermedia semanal — GDA-FLC-029.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Abrir →", key="cta_balanzas", use_container_width=True):
                    navigate("balanzas")
        elif es_ingeniero:
            pendientes_ing = _ensayos_pendientes_dt()
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="bento-primary"><div class="bento-icon">{icon("fact_check")}</div>'
                             f'<div><span class="bento-eyebrow">Tareas prioritarias</span>'
                             f'<h3>Ensayos pendientes de tu aprobación</h3><p>{len(pendientes_ing)} ensayo(s) '
                             f'confirmados por el Jefe de Laboratorio, esperando tu visto bueno final.</p></div></div>',
                             unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("sync")}</div>'
                             f'<div><h3>Proyectos en ejecución</h3><p>{sum(1 for p in st.session_state.projects if project_status(p["codigo_interno"])=="ejecucion")} proyecto(s) activos en laboratorio.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Ver proyectos →", key="cta_active_ing", use_container_width=True):
                    navigate("projects-active")
            with c3:
                st.markdown(f'<div class="bento-light"><div class="bento-icon">{icon("archive")}</div>'
                             '<div><h3>Proyectos ejecutados</h3><p>Consulta el historial certificado.</p></div></div>',
                             unsafe_allow_html=True)
                if st.button("Explorar archivo →", key="cta_done_ing", use_container_width=True):
                    navigate("projects-done")
            if pendientes_ing:
                st.markdown("<br>", unsafe_allow_html=True)
                for codigo, perf_codigo, m, ensayo_label in pendientes_ing[:5]:
                    proyecto = get_project(codigo)
                    with st.container(border=True):
                        cols = st.columns([3, 1])
                        cols[0].markdown(
                            f'<div class="cell-title">{html.escape(proyecto["nombre"] if proyecto else codigo)} · {html.escape(ensayo_label)}</div>'
                            f'<div class="cell-sub">{html.escape(codigo)} · {html.escape(perf_codigo)} · Muestra {m["numero"]}</div>',
                            unsafe_allow_html=True)
                        with cols[1]:
                            if st.button("Revisar →", key=f"revisar_ing_{m['id_unico']}_{ensayo_label}", use_container_width=True):
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

    if es_jefe and st.session_state.projects:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(f'<div class="section-title" style="border-bottom:none;margin-bottom:0;padding-bottom:0;">'
                        f'{icon("edit", size=15)} Editar un proyecto</div>', unsafe_allow_html=True)
            opciones = {p["codigo_interno"]: f'{p["codigo_interno"]} — {p["nombre"]}' for p in st.session_state.projects}
            elegido = st.selectbox("Proyecto", list(opciones), format_func=opciones.get,
                                    key="home_editar_proyecto", label_visibility="collapsed")
            e1, e2 = st.columns(2)
            with e1:
                if st.button("Editar proyecto", key="home_editar_btn", icon=":material/edit:", use_container_width=True):
                    st.session_state.selected_codigo = elegido
                    navigate("edit-project")
            with e2:
                if st.button("Editar bitácora de orden", key="home_editar_bitacora_btn",
                              icon=":material/assignment:", use_container_width=True):
                    st.session_state.selected_codigo = elegido
                    navigate("bitacora")

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
                col_ratios = [1.2, 2.0, 1.9, 1.5, 1.1, 0.9]
                headers = st.columns(col_ratios)
                for col, label in zip(headers, ["ID proyecto", "Nombre del proyecto", "Sondeo / Muestra", "Última actualización", "Estado", "Acción"]):
                    col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
                for i, a in enumerate(recientes):
                    if i:
                        st.markdown(f'<hr style="margin:8px 0;border-color:{BORDER};">', unsafe_allow_html=True)
                    proyecto = get_project(a["codigo_interno"])
                    titulo = html.escape(proyecto["nombre"] if proyecto else a["codigo_interno"])
                    subtitulo = html.escape(f'{a["perforacion_codigo"]} · Muestra {a["muestra_numero"]} · {ASSAY_LABELS[a["tipo"]]}')
                    actualizacion = format_dt(a["lastModified"])
                    if a.get("laboratorist"):
                        actualizacion += f' · {html.escape(a["laboratorist"])}'
                    cols = st.columns(col_ratios, vertical_alignment="center")
                    cols[0].markdown(f'<span class="cell-id">{html.escape(a["codigo_interno"])}</span>', unsafe_allow_html=True)
                    cols[1].markdown(f'<div class="cell-title">{titulo}</div>', unsafe_allow_html=True)
                    cols[2].markdown(f'<div class="cell-sub">{subtitulo}</div>', unsafe_allow_html=True)
                    cols[3].markdown(f'<span class="cell-muted">{html.escape(actualizacion)}</span>', unsafe_allow_html=True)
                    with cols[4]:
                        st.markdown(f'<div style="text-align:center;">{status_circle_html(a["status"], size=16)}</div>', unsafe_allow_html=True)
                    with cols[5]:
                        if st.button("Abrir", key=f"open_recent_{a['id']}", use_container_width=True):
                            st.session_state.selected_codigo = a["codigo_interno"]
                            navigate("project-detail")
                st.markdown(f'<div class="activity-footer">Mostrando {len(recientes)} de {len(todos_los_ensayos)} ensayo(s)</div>',
                            unsafe_allow_html=True)
    else:
        # Cualquier laboratorista puede abrir cualquier ensayo pendiente — no hay asignación
        # previa del Jefe, cada quien indica su propio nombre en el campo "Laboratorista" del
        # formulario del ensayo al hacerlo. Por eso se recorren las muestras y su checklist de
        # ensayos directamente, en vez de filtrar st.session_state.assays (que solo trae ensayos
        # que YA tienen una fila creada, es decir que alguien ya abrió al menos una vez), para que
        # un ensayo recién marcado en la bitácora aparezca aquí de inmediato aunque nadie lo haya
        # abierto todavía.
        pendientes = []
        for key, muestras_perf in st.session_state.muestras.items():
            codigo_proyecto, perf_codigo = key.split("::", 1)
            for m in muestras_perf:
                for ensayo_label in unificar_ensayos([e for e, v in m["ensayos"].items() if v]):
                    tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
                    if not tipo_interno:
                        continue  # sin formulario propio, no hay nada que "Abrir"
                    existing = get_assay(m["id_unico"], tipo_interno)
                    status = existing["status"] if existing else "sin-iniciar"
                    if status == "finalizado":
                        continue
                    pendientes.append({
                        "id": existing["id"] if existing else None,
                        "codigo_interno": codigo_proyecto, "perforacion_codigo": perf_codigo,
                        "muestra_id": m["id_unico"], "muestra_db_id": m["id"], "muestra_numero": m["numero"],
                        "tipo": tipo_interno, "status": status,
                        "lastModified": existing["lastModified"] if existing else m.get("updated_at", ""),
                    })
        pendientes.sort(key=lambda a: a["lastModified"], reverse=True)
        with st.container(border=True):
            h1, h2 = st.columns([4, 1])
            with h1:
                st.markdown(f'<div class="section-title" style="border-bottom:none;margin-bottom:0;padding-bottom:0;">'
                            f'{icon("assignment", size=15)} Ensayos pendientes</div>', unsafe_allow_html=True)
            with h2:
                st.markdown(f'<div style="text-align:right;"><span class="badge badge-muted">Total: {len(pendientes)}</span></div>',
                            unsafe_allow_html=True)

            if not pendientes:
                st.info("No hay ensayos pendientes por ahora.")
            else:
                col_ratios = [1.2, 2.0, 1.9, 1.5, 1.1, 0.9]
                headers = st.columns(col_ratios)
                for col, label in zip(headers, ["ID ensayo", "Proyecto", "Tipo de ensayo", "Última actualización", "Estado", "Acción"]):
                    col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
                for i, a in enumerate(pendientes):
                    if i:
                        st.markdown(f'<hr style="margin:8px 0;border-color:{BORDER};">', unsafe_allow_html=True)
                    proyecto = get_project(a["codigo_interno"])
                    cols = st.columns(col_ratios, vertical_alignment="center")
                    ensayo_id = f'{a["codigo_interno"]}-{a["perforacion_codigo"]}-M{a["muestra_numero"]}'
                    cols[0].markdown(f'<span class="cell-id">{html.escape(ensayo_id)}</span>', unsafe_allow_html=True)
                    titulo = html.escape(proyecto["nombre"] if proyecto else a["codigo_interno"])
                    subtitulo = html.escape(proyecto.get("localizacion", "")) if proyecto else ""
                    cols[1].markdown(f'<div class="cell-title">{titulo}</div><div class="cell-sub">{subtitulo}</div>',
                                      unsafe_allow_html=True)
                    # Texto plano (mismo estilo que "Sondeo / Muestra" en Actividad reciente del
                    # Jefe, .cell-sub) en vez del chip con borde de antes — se ve más liviano.
                    cols[2].markdown(f'<div class="cell-sub">{html.escape(ASSAY_LABELS[a["tipo"]])}</div>', unsafe_allow_html=True)
                    cols[3].markdown(f'<span class="cell-muted">{html.escape(format_dt(a["lastModified"]))}</span>',
                                      unsafe_allow_html=True)
                    with cols[4]:
                        st.markdown(f'<div style="text-align:center;">{status_circle_html(a["status"], size=16)}</div>', unsafe_allow_html=True)
                    with cols[5]:
                        if st.button("Abrir", key=f"open_assigned_{a['muestra_id']}_{a['tipo']}", use_container_width=True):
                            if a["id"]:
                                st.session_state.selected_assay_id = a["id"]
                            else:
                                nuevo = db.create_assay(a["muestra_db_id"], a["tipo"])
                                st.session_state.selected_assay_id = nuevo["id"]
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
                        db.archive_project(p["id"])
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


_bal_component = components.declare_component("bal_comprobacion", path=os.path.join(BASE_DIR, "balanzas_component"))
# Claves del estado del mockup (Component.init() del HTML) que sí se guardan en cada registro.
BAL_ESTADO_KEYS = ("plate", "fecha", "proxima", "eq", "exc", "rep", "exa", "pat", "cond", "tec", "fuente", "obs",
                    "tnc", "decision", "decisionDate", "elaboro", "reviso")


def _bal_estado_inicial(balanza, hoy):
    return {"eq": {"codigo": balanza["codigo"], "nombre": balanza["nombre"], "marca": balanza["marca"],
                    "serie": balanza["serie"], "d": balanza["resolucion"]},
            "fecha": str(hoy), "proxima": str(hoy + timedelta(days=8))}


def _bal_cargar_estado(check):
    """Estado que se le manda al mockup: los datos de la balanza + lo ya guardado del registro (los
    registros de la versión anterior de esta pantalla no traen estas claves y arrancan limpios)."""
    b = _bal_por_codigo(check["codigo_equipo"])
    try:
        base = _bal_estado_inicial(b, date.fromisoformat(check.get("fecha_comprobacion") or check["semana_lunes"])) if b else {}
    except ValueError:
        base = {}
    guardado = {k: v for k, v in (check.get("data") or {}).items() if k in BAL_ESTADO_KEYS}
    return {**base, **guardado}


def _bal_state_a_data(state):
    """Estado del mockup -> formato plano que usan los cálculos y la bitácora en Excel."""
    exc, rep, exa = state.get("exc") or {}, state.get("rep") or {}, state.get("exa") or {}
    cond, tec, eq = state.get("cond") or {}, state.get("tec") or {}, state.get("eq") or {}
    d = {"eq_nombre": eq.get("nombre", ""), "eq_marca": eq.get("marca", ""), "eq_serie": eq.get("serie", ""),
         "eq_resolucion": eq.get("d", ""),
         "exc_forma": {"circular": "Circular", "triangular": "Triangular"}.get(state.get("plate"), "Cuadrado"),
         "exc_carga_usada": exc.get("carga", ""), "rep_carga_usada": rep.get("carga", ""),
         "rep_limite_r": rep.get("limite", "")}
    for i, r in enumerate(exc.get("rows") or [], start=1):
        d[f"exc_p{i}_indicacion"], d[f"exc_p{i}_emp"] = r.get("ind", ""), r.get("emp", "")
    for i, r in enumerate(rep.get("rows") or [], start=1):
        d[f"rep_r{i}_indicacion"], d[f"rep_r{i}_emp"] = r.get("ind", ""), r.get("emp", "")
    for i, r in enumerate(exa.get("rows") or [], start=1):
        d[f"exact_p{i}_aplica"] = r.get("aplica", True) is not False
        d[f"exact_p{i}_carga"], d[f"exact_p{i}_asc"] = r.get("carga", ""), r.get("asc", "")
        d[f"exact_p{i}_desc"], d[f"exact_p{i}_emp"] = r.get("desc", ""), r.get("emp", "")
        d[f"exact_p{i}_just"] = r.get("just", "")
    d["patrones"] = [{"codigo": p.get("codigo", ""), "instrumento": p.get("instrumento", ""),
                       "fecha_calibracion": p.get("ultima", ""), "proxima_calibracion": p.get("proxima", ""),
                       "no_certificado": p.get("cert", "")} for p in (state.get("pat") or []) if any(p.values())]
    d.update(cond_limpieza=bool(cond.get("limpieza")), cond_nivelacion=bool(cond.get("nivelacion")),
             cond_cero_tara=bool(cond.get("cero")), cond_estabilizacion=bool(cond.get("estab")),
             cond_masas_limpias=bool(cond.get("patrones")), cond_temperatura=cond.get("temp", ""),
             cond_humedad=cond.get("hum", ""), dt_capacidad_maxima=tec.get("cap", ""), dt_division_e=tec.get("div", ""),
             dt_clase=tec.get("clase", ""), dt_certificado_balanza=tec.get("cert", ""),
             fuente_criterios=state.get("fuente", ""), observaciones=state.get("obs", ""),
             tnc_numero=state.get("tnc", ""), decision_fecha=state.get("decisionDate", ""),
             decision={"apto": "apto", "noapto": "no-apto"}.get(state.get("decision")))
    for prefijo in ("elaboro", "reviso"):
        o = state.get(prefijo) or {}
        d[f"{prefijo}_nombre"], d[f"{prefijo}_cargo"], d[f"{prefijo}_firmo"] = o.get("nombre", ""), o.get("cargo", ""), bool(o.get("firmado"))
    return d


def _bal_estado_db(state):
    """Estado que se guarda en la columna `estado` (para listar): la decisión del Jefe si ya la tomó;
    si no, "no-apto" cuando alguna prueba ya no cumple, y "pendiente" en cualquier otro caso."""
    decision = state.get("decision")
    if decision in ("apto", "noapto"):
        return "apto" if decision == "apto" else "no-apto"
    return "no-apto" if decision_sugerida_balanza(_bal_state_a_data(state)) == "no-apto" else "pendiente"


def _bal_control():
    """Filas del tablero "Control semanal" (lista única compartida), leídas una vez por sesión."""
    if "_bal_control_cache" not in st.session_state:
        try:
            st.session_state["_bal_control_cache"] = db.get_balance_control()
        except Exception:
            st.session_state["_bal_control_cache"] = []
            st.session_state["_bal_control_error"] = True
    return st.session_state["_bal_control_cache"]


def _bal_rotulo(check):
    estado = {"apto": "Apto", "no-apto": "No apto"}.get(check.get("estado"), "Pendiente")
    try:
        fecha = date.fromisoformat(check.get("fecha_comprobacion") or check["semana_lunes"]).strftime("%d/%m/%Y")
        semana = date.fromisoformat(check["semana_lunes"]).strftime("%d/%m")
    except ValueError:
        return estado
    return f"{fecha} · semana del {semana} · {estado}"


def _bal_nuevo_registro(codigo):
    """Crea el registro de esta semana para la balanza (uno por equipo y semana — si ya existe, lo abre)."""
    balanza, hoy = _bal_por_codigo(codigo), date.today()
    if not balanza or not balanza.get("activa", True):
        st.session_state["_bal_aviso"] = "Esa balanza está dada de baja: no se le pueden crear registros nuevos (los anteriores se pueden consultar)."
        st.rerun()
        return
    semana = str(_bal_lunes(hoy))
    destino = next((c for c in st.session_state.balance_checks
                    if c["codigo_equipo"] == codigo and c["semana_lunes"] == semana), None)
    if destino:
        st.session_state["_bal_aviso"] = "Esta balanza ya tenía registro de esta semana (uno por equipo y semana): se abrió ese."
    else:
        destino = db.create_balance_check(codigo, semana, fecha_comprobacion=str(hoy),
                                            fecha_proxima=str(hoy + timedelta(days=8)), data=_bal_estado_inicial(balanza, hoy))
        st.session_state.balance_checks.append(destino)
    st.session_state["_bal_pend_equipo"] = codigo
    st.session_state["_bal_pend_registro"] = destino["id"]
    st.rerun()


def _bal_guardar(check, valor):
    """Guarda en Supabase lo que devolvió el mockup (registro + tablero de Control semanal)."""
    state, control = valor.get("state") or {}, valor.get("control")
    try:
        if state and state != (check.get("data") or {}):
            fecha = state.get("fecha") or check.get("fecha_comprobacion")
            cambios = {"data": state, "estado": _bal_estado_db(state), "fecha_comprobacion": fecha}
            if state.get("proxima"):
                cambios["fecha_proxima"] = state["proxima"]
            try:
                cambios["semana_lunes"] = str(_bal_lunes(date.fromisoformat(fecha)))
            except (TypeError, ValueError):
                pass
            check.update(db.update_balance_check(check["id"], **cambios))
        if control is not None and control != _bal_control():
            try:
                db.save_balance_control(control)
                st.session_state["_bal_control_cache"] = control
            except Exception:
                st.session_state["_bal_control_error"] = True
    except Exception:
        st.error("No se pudo guardar el último cambio (revisa tu conexión). Sigue en pantalla — vuelve a intentarlo "
                 "modificando algo más.")


def _bal_admin():
    """Agregar, editar, dar de baja / reactivar y eliminar balanzas del catálogo."""
    with st.expander("Administrar balanzas (agregar, editar, dar de baja)", icon=":material/tune:",
                     expanded=bool(st.session_state.get("_bal_admin_open"))):
        if st.session_state.get("_bal_sin_tabla"):
            st.warning("Para administrar las balanzas falta correr la migración 0032_balanzas.sql en Supabase "
                       "(mientras tanto se usan las 4 de siempre).")
            return
        lista = _bal_lista()
        conteo = {}
        for c in st.session_state.balance_checks:
            conteo[c["codigo_equipo"]] = conteo.get(c["codigo_equipo"], 0) + 1
        st.caption("Una balanza dada de baja deja de ofrecerse para registros nuevos, pero su historial se sigue "
                   "pudiendo consultar y descargar. Los registros ya guardados conservan los datos del equipo que tenían.")
        for b in lista:
            n = conteo.get(b["codigo"], 0)
            with st.container(border=True):
                estado = "" if b["activa"] else " · **de baja**"
                st.markdown(f"**{html.escape(b['codigo'])}** — {html.escape(b['nombre'])}{estado}  \n"
                            f"{html.escape(b['marca'] or '—')} · Serie {html.escape(b['serie'] or '—')} · "
                            f"d = {html.escape(b['resolucion'] or '—')} g · {n} registro(s)")
                c1, c2, c3 = st.columns(3)
                if c1.button("Editar", key=f"bal_ed_{b['id']}", use_container_width=True, icon=":material/edit:"):
                    st.session_state["_bal_edit"] = b["id"]
                    st.session_state["_bal_admin_open"] = True
                    st.rerun()
                if b["activa"]:
                    if c2.button("Dar de baja", key=f"bal_baja_{b['id']}", use_container_width=True):
                        db.update_balanza(b["id"], activa=False)
                        st.session_state["_bal_admin_open"] = True
                        st.rerun()
                elif c2.button("Reactivar", key=f"bal_alta_{b['id']}", use_container_width=True):
                    db.update_balanza(b["id"], activa=True)
                    st.session_state["_bal_admin_open"] = True
                    st.rerun()
                with c3:
                    if n:
                        st.caption("Tiene registros: en vez de eliminarla, dala de baja.")
                    elif confirm_delete(f"balanza_{b['id']}", f"la balanza {b['codigo']}"):
                        db.delete_balanza(b["id"])
                        st.session_state.pop("_bal_edit", None)
                        st.session_state["_bal_admin_open"] = True
                        st.rerun()

        editando = st.session_state.get("_bal_edit")
        base = next((b for b in lista if b["id"] == editando), None)
        n_base = conteo.get(base["codigo"], 0) if base else 0
        st.markdown("##### " + (f"Editar balanza {base['codigo']}" if base else "Agregar balanza"))
        with st.form(f"bal_form_{editando or 'nuevo'}", clear_on_submit=not base):
            codigo = st.text_input("Código interno", value=base["codigo"] if base else "", placeholder="GDA-E-014",
                                   disabled=bool(base and n_base))
            if base and n_base:
                st.caption("El código no se puede cambiar porque la balanza ya tiene registros.")
            nombre = st.text_input("Nombre del equipo", value=base["nombre"] if base else "", placeholder="Balanza Cap. 600g")
            f1, f2 = st.columns(2)
            marca = f1.text_input("Marca", value=base["marca"] if base else "", placeholder="TRUMAX")
            serie = f2.text_input("Serie", value=base["serie"] if base else "", placeholder="MIX-H")
            resolucion = st.text_input("Resolución d (g)", value=base["resolucion"] if base else "", placeholder="0,01")
            b1, b2 = st.columns(2)
            guardar = b1.form_submit_button("Guardar cambios" if base else "Agregar balanza", type="primary",
                                            use_container_width=True)
            cancelar = b2.form_submit_button("Cancelar", use_container_width=True) if base else False
        if cancelar:
            st.session_state.pop("_bal_edit", None)
            st.session_state["_bal_admin_open"] = True
            st.rerun()
        if guardar:
            codigo = (base["codigo"] if (base and n_base) else codigo).strip().upper()
            nombre, marca, serie, resolucion = nombre.strip(), marca.strip(), serie.strip(), resolucion.strip()
            if not codigo or not nombre:
                st.error("El código interno y el nombre del equipo son obligatorios.")
            elif not resolucion or (to_float(resolucion) or 0) <= 0:
                st.error("La resolución d debe ser un número mayor que 0 (por ejemplo 0,01).")
            elif any(o["codigo"].upper() == codigo and (not base or o["id"] != base["id"]) for o in lista):
                st.error(f"Ya existe una balanza con el código {codigo}.")
            else:
                try:
                    if base:
                        db.update_balanza(base["id"], codigo=codigo, nombre=nombre, marca=marca, serie=serie, resolucion=resolucion)
                        st.session_state.pop("_bal_edit", None)
                    else:
                        db.create_balanza(codigo, nombre, marca, serie, resolucion)
                    st.session_state["_bal_pend_equipo"] = codigo
                    st.session_state["_bal_aviso"] = f"Balanza {codigo} guardada."
                    st.session_state["_bal_admin_open"] = True
                    st.rerun()
                except Exception:
                    st.error("No se pudo guardar la balanza (revisa tu conexión).")


def render_balanzas():
    require_role("jefe")
    # El mockup es un tablero de ~1440 px: se reduce el margen lateral de Streamlit solo en esta pantalla.
    st.markdown("<style>[data-testid='stMainBlockContainer']{padding-left:1rem !important;"
                "padding-right:1rem !important;max-width:none !important;}</style>", unsafe_allow_html=True)
    if st.button("← Atrás"):
        go_back(fallback="home")

    # Selecciones pedidas en el run anterior (no se pueden asignar después de crear el selector).
    for destino, pendiente in (("bal_sel_equipo", "_bal_pend_equipo"), ("bal_sel_registro", "_bal_pend_registro")):
        if pendiente in st.session_state:
            st.session_state[destino] = st.session_state.pop(pendiente)
    codigos = _bal_codigos_selector()
    if st.session_state.get("_bal_aviso"):
        st.info(st.session_state.pop("_bal_aviso"))
    if not codigos:
        st.info("No hay balanzas registradas: agrega la primera en «Administrar balanzas».")
        _bal_admin()
        return
    if st.session_state.get("bal_sel_equipo") not in codigos:
        st.session_state["bal_sel_equipo"] = codigos[0]

    # Dos filas de dos (en una tablet vertical cuatro columnas quedan demasiado apretadas).
    t1, t2 = st.columns(2)
    with t1:
        equipo = st.selectbox("Balanza", codigos, key="bal_sel_equipo",
                               format_func=lambda c: f"{c} — {_bal_por_codigo(c)['nombre']}"
                               + ("" if _bal_por_codigo(c).get("activa", True) else " (de baja)"))
    registros = sorted([c for c in st.session_state.balance_checks if c["codigo_equipo"] == equipo],
                       key=lambda c: (c["semana_lunes"], c.get("created_at") or ""), reverse=True)
    ids = [c["id"] for c in registros]
    if ids and st.session_state.get("bal_sel_registro") not in ids:
        st.session_state["bal_sel_registro"] = ids[0]
    with t2:
        sel_id = None
        if ids:
            sel_id = st.selectbox("Registro guardado", ids, key="bal_sel_registro",
                                   format_func=lambda i: _bal_rotulo(next(c for c in registros if c["id"] == i)))
        else:
            st.markdown('<div style="padding-top:8px;color:#6B7570;">Esta balanza todavía no tiene registros.</div>',
                        unsafe_allow_html=True)
    t3, t4 = st.columns(2)
    with t3:
        if st.button("Nuevo registro", key="bal_nuevo", type="primary", use_container_width=True, icon=":material/add:",
                     disabled=not _bal_por_codigo(equipo).get("activa", True)):
            _bal_nuevo_registro(equipo)
    with t4:
        slot_descarga = st.empty()

    if st.session_state.get("_bal_control_error"):
        st.warning("El Control semanal no se pudo leer/guardar: falta correr la migración 0031_balance_control.sql en Supabase.")
    _bal_admin()
    if not sel_id:
        st.info("Crea el primer registro de esta balanza con «Nuevo registro» (uno por equipo y semana).")
        return

    check = next(c for c in registros if c["id"] == sel_id)
    token = check["id"]
    valor = _bal_component(load_token=token, state=_bal_cargar_estado(check), control=_bal_control(),
                            key="bal_comp", default=None)
    # El componente devuelve su último valor en cada rerun: solo se procesa uno nuevo y de ESTE registro.
    if valor and valor.get("token") == token and st.session_state.get("_bal_ultimo_valor") != (token, valor.get("seq")):
        st.session_state["_bal_ultimo_valor"] = (token, valor.get("seq"))
        if valor.get("action") == "new":
            _bal_nuevo_registro(equipo)
        else:
            _bal_guardar(check, valor)

    balanza = _bal_por_codigo(check["codigo_equipo"]) or {}
    with slot_descarga:
        st.download_button(
            "Descargar Excel", icon=":material/download:", use_container_width=True, key="bal_dl",
            data=generar_excel_balanza({**check, "data": _bal_state_a_data(_bal_cargar_estado(check))}, balanza,
                                       control=_bal_control()),
            file_name=f"Verificacion_balanza_{check['codigo_equipo']}_{check['semana_lunes']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("Opciones del registro", icon=":material/settings:"):
        if confirm_delete(f"bal_{check['id']}", "este registro de comprobación"):
            db.delete_balance_check(check["id"])
            st.session_state.balance_checks[:] = [c for c in st.session_state.balance_checks if c["id"] != check["id"]]
            st.session_state.pop("bal_sel_registro", None)
            st.rerun()


def _bal_dmy(iso):
    """"2026-09-23" -> "23/09/2026" ("" si no es una fecha)."""
    try:
        return date.fromisoformat(str(iso)).strftime("%d/%m/%Y")
    except ValueError:
        return ""


def _bal_serial_excel(iso):
    """Fecha ISO -> número de serie de Excel (así la celda conserva su formato de fecha), o None."""
    try:
        return (date.fromisoformat(str(iso)) - date(1899, 12, 30)).days
    except ValueError:
        return None


def _bal_estado_control(fila):
    """Misma regla del tablero del mockup: realizada <= programada -> "A tiempo"; después -> "Retraso · TNC"."""
    programada, realizada = fila.get("programada"), fila.get("realizada")
    if realizada and programada:
        estado = "A tiempo" if realizada <= programada else "Retraso · TNC"
    elif realizada:
        estado = "Realizada"
    else:
        estado = "Pendiente"
    return f"{estado} — {fila['just']}" if fila.get("just") else estado


def _xlsx_combinar_con_ajuste(xlsx_bytes, hoja_xml, ref_celda, rango):
    """Combina `rango` (ej. "E38:AV39") y deja la celda `ref_celda` con el texto alineado arriba a la izquierda y con
    ajuste de línea: en la plantilla, Observaciones son dos filas de celdas sueltas centradas, donde un texto largo
    se sale de la hoja. Se clona el estilo que ya tenía la celda (bordes, fuente) y solo se cambia la alineación."""
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zin:
        sheet = zin.read(hoja_xml).decode("utf-8")
        styles = zin.read("xl/styles.xml").decode("utf-8")
        m_cell = re.search(r'<c r="' + ref_celda + r'"([^>]*?)(/>|>)', sheet)
        estilo = int(re.search(r'\bs="(\d+)"', m_cell.group(1)).group(1))
        m_xfs = re.search(r'(<cellXfs count=")(\d+)(">)(.*?)(</cellXfs>)', styles, re.S)
        xfs = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', m_xfs.group(4), re.S)
        xf, alineacion = xfs[estilo], '<alignment horizontal="left" vertical="top" wrapText="1"/>'
        if "<alignment" in xf:
            xf = re.sub(r'<alignment[^>]*/>', alineacion, xf)
        elif xf.endswith("/>"):
            xf = xf[:-2] + ">" + alineacion + "</xf>"
        else:
            xf = xf.replace("</xf>", alineacion + "</xf>")
        if "applyAlignment" not in xf:
            xf = xf.replace("<xf ", '<xf applyAlignment="1" ', 1)
        nuevo_idx = len(xfs)
        styles = (styles[:m_xfs.start()] + m_xfs.group(1) + str(len(xfs) + 1) + m_xfs.group(3) + m_xfs.group(4) + xf
                  + m_xfs.group(5) + styles[m_xfs.end():])
        sheet = sheet[:m_cell.start()] + re.sub(r'\bs="\d+"', f's="{nuevo_idx}"', m_cell.group(0), count=1) + sheet[m_cell.end():]
        m_mc = re.search(r'<mergeCells count="(\d+)">', sheet)
        sheet = sheet.replace(m_mc.group(0), f'<mergeCells count="{int(m_mc.group(1)) + 1}"><mergeCell ref="{rango}"/>', 1)
        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                contenido = zin.read(item.filename)
                if item.filename == hoja_xml:
                    contenido = sheet.encode("utf-8")
                elif item.filename == "xl/styles.xml":
                    contenido = styles.encode("utf-8")
                zout.writestr(item, contenido)
        return bio.getvalue()


def generar_excel_balanza(check, balanza, control=None):
    """Bitácora GDA-FL-029 "Verificación de equipos — Balanzas" (la plantilla oficial de VERIFICACION DE BALANZAS.xlsx)
    llena con un registro de Comprobación de Balanzas. Se escribe directo en el XML de la hoja "Balanzas" solo en las
    celdas de datos (amarillas/azules y los textos de la parte de abajo) — el código interno trae solo nombre, marca,
    serie y resolución (BUSCARV contra el catálogo de equipos de la propia hoja) y los errores, "Cumple/No cumple", el
    rango R y la fecha próxima los calcula el propio Excel con sus fórmulas. `check["data"]` viene en el formato plano
    de _bal_state_a_data; `control` son las filas del tablero de Control semanal (hoja "Control semanal")."""
    data = check.get("data", {})
    num = to_float
    c = {}
    fecha = check.get("fecha_comprobacion")
    c["J6"] = _bal_serial_excel(fecha)
    proxima = check.get("fecha_proxima")
    try:
        if proxima and date.fromisoformat(proxima) != date.fromisoformat(fecha) + timedelta(days=8):
            c["AG6"] = _bal_serial_excel(proxima)   # la plantilla calcula fecha + 8 días; solo se pisa si se cambió a mano
    except (TypeError, ValueError):
        pass
    c["J8"] = check["codigo_equipo"]
    # Nombre, marca, resolución y serie: la plantilla los busca por código en su propio catálogo (solo trae las 4
    # balanzas originales), así que se escriben directo con los datos del equipo tal como quedaron en el registro.
    balanza = balanza or {}
    c["Z8"] = data.get("eq_nombre") or balanza.get("nombre")
    c["J9"] = data.get("eq_marca") or balanza.get("marca")
    resolucion = data.get("eq_resolucion") or balanza.get("resolucion")
    c["Z9"] = num(resolucion) if num(resolucion) is not None else resolucion
    c["AP9"] = data.get("eq_serie") or balanza.get("serie")

    # 1. Excentricidad / 2. Repetibilidad
    c["S12"] = num(data.get("exc_carga_usada"))
    c["AM12"] = num(data.get("rep_carga_usada"))
    for p in range(1, 6):
        fila = 14 + p
        c[f"K{fila}"], c[f"Q{fila}"] = num(data.get(f"exc_p{p}_indicacion")), num(data.get(f"exc_p{p}_emp"))
        c[f"AE{fila}"], c[f"AK{fila}"] = num(data.get(f"rep_r{p}_indicacion")), num(data.get(f"rep_r{p}_emp"))
    # La plantilla trae la fórmula del error en los puntos 2 a 5 pero no en el 1 (queda siempre "Pendiente"): se completa.
    c["N15"] = _FormulaXlsx('IF(COUNT(K15,$S$12)<2,"",K15-$S$12)')
    c["O51"] = num(data.get("rep_limite_r"))

    # 3. Exactitud — un punto "No aplica" queda en blanco (sale "Pendiente" en la plantilla) y se justifica en Observaciones
    excluidos = []
    for p in range(1, 9):
        fila = 23 + p
        if data.get(f"exact_p{p}_aplica", True):
            c[f"K{fila}"], c[f"N{fila}"] = num(data.get(f"exact_p{p}_carga")), num(data.get(f"exact_p{p}_asc"))
            c[f"Q{fila}"], c[f"Z{fila}"] = num(data.get(f"exact_p{p}_desc")), num(data.get(f"exact_p{p}_emp"))
        else:
            excluidos.append(f"Exactitud punto {p}: No aplica — {data.get(f'exact_p{p}_just') or 'sin justificación'}.")

    # Instrumentos patrón (la plantilla trae 3 filas)
    patrones = data.get("patrones") or []
    for i, pat in enumerate(patrones[:3]):
        fila = 34 + i
        c[f"A{fila}"], c[f"F{fila}"] = pat.get("codigo"), pat.get("instrumento")
        c[f"T{fila}"], c[f"AC{fila}"] = _bal_dmy(pat.get("fecha_calibracion")) or pat.get("fecha_calibracion"), \
            _bal_dmy(pat.get("proxima_calibracion")) or pat.get("proxima_calibracion")
        c[f"AL{fila}"] = pat.get("no_certificado")
    extras = [f"{p.get('codigo') or '—'} ({p.get('instrumento') or '—'}, cert. {p.get('no_certificado') or '—'})" for p in patrones[3:]]

    # Observaciones (texto libre + lo que la plantilla no tiene dónde poner)
    partes = [str(data.get("observaciones") or "").strip(),
              f"Excentricidad: receptor {str(data.get('exc_forma') or 'Cuadrado').lower()}."] + excluidos
    if extras:
        partes.append("Patrones adicionales: " + "; ".join(extras) + ".")
    observaciones = " ".join(t for t in partes if t)
    c["E38"] = observaciones if len(observaciones) <= 700 else observaciones[:697] + "…"

    # Firmas (la etiqueta de la plantilla — NOMBRE / FIRMA / CARGO — se conserva y se le agrega el dato)
    for prefijo, col in (("elaboro", "A"), ("reviso", "Y")):
        if data.get(f"{prefijo}_nombre"):
            c[f"{col}42"] = f"NOMBRE: {data[f'{prefijo}_nombre']}"
        if data.get(f"{prefijo}_firmo"):
            c[f"{col}43"] = "FIRMA: Firmado (registro digital)"
        if data.get(f"{prefijo}_cargo"):
            c[f"{col}44"] = f"CARGO: {data[f'{prefijo}_cargo']}"

    # Criterios y decisión (filas 49 a 56, columna O)
    if str(data.get("fuente_criterios") or "").strip():
        c["O49"] = data["fuente_criterios"].strip()
    condiciones = [f"{label.split(' (')[0]}: {'Sí' if data.get(key) else 'No'}" for key, label in BAL_CONDICIONES_PREVIAS]
    if data.get("cond_temperatura"):
        condiciones.append(f"Temperatura: {data['cond_temperatura']} °C")
    if data.get("cond_humedad"):
        condiciones.append(f"Humedad relativa: {data['cond_humedad']} %")
    c["O53"] = "Condiciones previas — " + "; ".join(condiciones) + "."
    tec = [("Capacidad máxima (g)", "dt_capacidad_maxima"), ("División e (si aplica)", "dt_division_e"),
           ("Clase (si aplica)", "dt_clase"), ("Certificado de balanza", "dt_certificado_balanza")]
    if any(data.get(k) for _l, k in tec):
        c["O54"] = "  ".join(f"{label}: {data.get(key) or '____'}" for label, key in tec)
    if data.get("decision"):
        quien = data.get("reviso_nombre") or data.get("elaboro_nombre") or "Jefe de Laboratorio"
        cuando = _bal_dmy(data.get("decision_fecha"))
        c["O55"] = (f"{'APTO' if data['decision'] == 'apto' else 'NO APTO'}. Decisión de {quien}"
                    + (f" el {cuando}." if cuando else "."))
    if str(data.get("tnc_numero") or "").strip():
        c["O56"] = ("Si falla o se omite la verificación: detener uso, identificar fuera de servicio, evaluar impacto y "
                    f"relacionar GDA-FC-023 No.: {data['tnc_numero'].strip()}")

    with open(TEMPLATE_VERIFICACION_BALANZAS, "rb") as f:
        libro = f.read()
    libro = _xlsx_escribir_celdas(libro, "xl/worksheets/sheet1.xml", c)
    libro = _xlsx_combinar_con_ajuste(libro, "xl/worksheets/sheet1.xml", "E38", "E38:AV39")

    # Hoja "Control semanal": una fila por registro del tablero (la plantilla trae las filas 6 a 57)
    filas = {}
    for i, r in enumerate((control or [])[:52]):
        n = 6 + i
        filas.update({f"A{n}": _bal_dmy(r.get("semana")) or r.get("semana"), f"B{n}": r.get("codigo"),
                      f"C{n}": _bal_dmy(r.get("programada")) or r.get("programada"),
                      f"D{n}": _bal_dmy(r.get("realizada")) or r.get("realizada"), f"E{n}": r.get("operador"),
                      f"F{n}": r.get("revision"), f"G{n}": r.get("registro"), f"H{n}": _bal_estado_control(r)})
    if filas:
        libro = _xlsx_escribir_celdas(libro, "xl/worksheets/sheet2.xml", filas)
    return libro


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
    if st.session_state.role == "laboratorista":
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
    st.markdown("## Bitácora de proyecto")

    st.markdown('<div class="section-title">Código interno</div>', unsafe_allow_html=True)
    # Sugerencia de consecutivo: año actual (2 dígitos) y, dentro de los proyectos ya creados con
    # ese año, el número más alto + 1 — igual que armaría el consecutivo el Jefe a mano. Se
    # precargan como valor editable (no solo placeholder) para no tener que digitarlos siempre;
    # el Jefe puede borrarlos y poner otros si el proyecto es de un año distinto.
    anio_sugerido = str(date.today().year)[-2:]
    if "new_anio" not in st.session_state:
        st.session_state["new_anio"] = anio_sugerido
    if "new_numero" not in st.session_state:
        numeros_mismo_anio = [
            int(p["numero"]) for p in st.session_state.projects
            if str(p.get("anio", "")) == st.session_state["new_anio"] and str(p.get("numero", "")).isdigit()
        ]
        st.session_state["new_numero"] = f"{(max(numeros_mismo_anio) + 1):03d}" if numeros_mismo_anio else "001"

    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.text_input("Prefijo", value="GDA", disabled=True, autocomplete="off")
    with c2:
        numero = st.text_input("Número", key="new_numero", autocomplete="off")
    with c3:
        anio = st.text_input("Año", key="new_anio", autocomplete="off")

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
    # A diferencia de las demás fechas, esta normalmente no se sabe todavía al crear el
    # proyecto (es cuándo terminó de verdad, no la fecha final planeada) — arranca vacía en
    # vez de con la fecha de hoy, y se puede completar después desde "Editar proyecto".
    fecha_final_real = st.date_input(
        "Fecha final real", key="new_fecha_final_real", format="DD/MM/YYYY", value=None,
        help="Cuándo terminó de verdad el proyecto — distinta de la fecha final planeada. "
             "Déjala vacía si todavía no ha terminado.",
    )

    # La asignación ya no es por proyecto entero — se asigna ensayo por ensayo desde el
    # detalle de cada muestra (ver render_muestra_detail), una vez que la bitácora existe.

    # Perforaciones y muestras se arman aquí mismo, antes de crear el proyecto formalmente —
    # en estado de borrador propio (no en st.session_state.perforaciones/muestras, que ahora
    # se recargan desde Supabase en cada rerun y no existen todavía para un proyecto sin crear).
    perforaciones = st.session_state.setdefault("draft_perforaciones", [])
    edited_frames = {}
    if codigo_valido and nombre:
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
                st.session_state.setdefault("draft_muestras", {})[codigo_perf] = []
                st.rerun()

        if perforaciones:
            st.markdown('<div class="section-title">Perforaciones y muestras</div>', unsafe_allow_html=True)
        for perf in perforaciones:
            key = f"{codigo_interno}::{perf['codigo']}"
            muestras = st.session_state.draft_muestras.setdefault(perf["codigo"], [])
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
                    st.session_state.draft_perforaciones = [p for p in perforaciones if p["codigo"] != perf["codigo"]]
                    st.session_state.draft_muestras.pop(perf["codigo"], None)
                    st.session_state.bitacora_draft.pop(key, None)
                    st.rerun()
    elif nombre or numero or anio:
        st.info("Completa un código interno válido y el nombre del proyecto para agregar perforaciones y muestras.")

    def _limpiar_borrador():
        st.session_state.draft_perforaciones = []
        st.session_state.draft_muestras = {}
        st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items()
                                            if not k.startswith(f"{codigo_interno}::")}
        # El número sugerido se calcula una sola vez por visita a esta pantalla (ver arriba) — hay
        # que olvidarlo al salir para que la próxima vez se recalcule contra la lista de proyectos
        # ya actualizada (si no, seguiría sugiriendo el mismo número recién usado).
        st.session_state.pop("new_numero", None)

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancelar", use_container_width=True):
            _limpiar_borrador()
            navigate("home")
    with col2:
        if st.button("Guardar bitácora", type="primary", use_container_width=True, icon=":material/save:",
                      disabled=not codigo_valido or not nombre):
            perforaciones_payload = []
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
                perforaciones_payload.append({**perf, "muestras": nuevas})

            db.commit_new_project({
                "codigo_interno": codigo_interno, "numero": numero, "anio": anio, "nombre": nombre,
                "localizacion": localizacion, "norma": norma,
                "fecha_bitacora": str(fecha_bitacora), "fecha_ingreso_muestra": str(fecha_ingreso),
                "cliente": cliente, "correo_cliente": correo_cliente, "muestra_tomada_por": muestra_tomada_por,
                "direccion_cliente": direccion_cliente, "telefono_contacto": telefono_contacto,
                "nombre_contacto": nombre_contacto,
                "fecha_inicio_proyecto": str(fecha_inicio_proyecto), "fecha_final_proyecto": str(fecha_final_proyecto),
                "fecha_final_real": str(fecha_final_real) if fecha_final_real else "",
                "fecha_recepcion": str(fecha_recepcion), "fecha_ejecucion": str(fecha_ejecucion), "fecha_emision": str(fecha_emision),
            }, perforaciones_payload)
            _limpiar_borrador()
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

    # Alerta de plazo: solo Jefe y Director Técnico la ven, y solo mientras el proyecto sigue sin
    # entregarse del todo — project_status devuelve "ejecutado" solo cuando TODAS las muestras
    # están finalizadas Y TODOS los ensayos tienen el visto bueno del Director Técnico; si ya se
    # entregó todo, seguir advirtiendo sobre la fecha límite no aporta nada, aunque haya pasado.
    # Se basa en "Fecha final real" (el plazo que de verdad hay que cumplir, ajustado a
    # contratiempos) y no en "Fecha final proyecto" (la referencia inicial que se le da al cliente).
    if st.session_state.role in ("jefe", "ingeniero") and project_status(codigo) != "ejecutado":
        deadline_raw = project.get("fecha_final_real")
        dias_restantes = None
        if deadline_raw:
            try:
                dias_restantes = (date.fromisoformat(deadline_raw) - date.today()).days
            except ValueError:
                dias_restantes = None
        if dias_restantes is not None and dias_restantes <= 0:
            # Vencida (o vence hoy) y todavía faltan entregas: se resalta mucho más fuerte que
            # los demás estados — encabezado propio en mayúsculas + detalle, en vez de la línea
            # sencilla que usan los casos "faltan X días".
            sub = (f"Pasaron {abs(dias_restantes)} día(s) desde la fecha final real y todavía "
                   f"faltan ensayos por entregar." if dias_restantes < 0 else
                   "Hoy es la fecha final real y todavía faltan ensayos por entregar.")
            st.markdown(f'''
                <div style="display:flex;align-items:flex-start;gap:10px;background:{DANGER_LIGHT};
                            color:{DANGER};border-radius:10px;border:1.5px solid {DANGER};
                            padding:12px 14px;margin-bottom:16px;">
                    {icon("report", size=22)}
                    <div>
                        <div style="font-weight:800;font-size:15px;">
                            {"¡FECHA LÍMITE VENCIDA!" if dias_restantes < 0 else "¡VENCE HOY!"}
                        </div>
                        <div style="font-size:13px;margin-top:2px;">{sub} ({deadline_raw})</div>
                    </div>
                </div>
            ''', unsafe_allow_html=True)
        elif dias_restantes is not None:
            if dias_restantes <= 5:
                tono, fondo, icono_alerta = WARNING, WARNING_LIGHT, "schedule"
            else:
                tono, fondo, icono_alerta = SUCCESS, SUCCESS_LIGHT, "event_available"
            texto = f"Faltan {dias_restantes} día(s)"
            st.markdown(f'''
                <div style="display:flex;align-items:center;gap:10px;background:{fondo};color:{tono};
                            border-radius:10px;padding:10px 14px;margin-bottom:16px;font-weight:600;font-size:14px;">
                    {icon(icono_alerta, size=18)}
                    <span>{texto} para la fecha final real del proyecto ({deadline_raw})</span>
                </div>
            ''', unsafe_allow_html=True)

    with st.container(border=True):
        info_rows = [
            ("location_on", "Ubicación", project.get("localizacion")),
            ("rule", "Norma", project.get("norma")),
        ]
        # Datos del cliente: los ven Jefe y Director Técnico (roles de consulta/revisión), nunca el laboratorista.
        if st.session_state.role in ("jefe", "ingeniero"):
            info_rows.insert(1, ("badge", "Cliente", project.get("cliente")))
            info_rows.insert(2, ("mail", "Correo electrónico", project.get("correo_cliente")))
            info_rows.insert(3, ("home_pin", "Dirección cliente", project.get("direccion_cliente")))
            info_rows.insert(4, ("call", "Teléfono de contacto", project.get("telefono_contacto")))
            info_rows.insert(5, ("contact_page", "Nombre de contacto", project.get("nombre_contacto")))

        def _info_row(icono, label, valor, primera):
            margen = "" if primera else "margin-top:14px;"
            return (f'<div class="cell-muted" style="{margen}text-transform:uppercase;letter-spacing:0.04em;font-size:11px;">'
                    f'{icon(icono, size=14)} {label}</div>'
                    f'<div style="font-weight:600;font-size:15px;">{html.escape(valor or "—")}</div>')

        for i, (icono, label, valor) in enumerate(info_rows):
            st.markdown(_info_row(icono, label, valor, i == 0), unsafe_allow_html=True)

        # Fechas agrupadas en paquetes separados (orden/ingreso, plazos del proyecto, ejecución/emisión)
        # para que no se vean como una sola lista larga y confusa.
        date_groups = [
            [("calendar_month", "Fecha de orden", project.get("fecha_bitacora")),
             ("move_to_inbox", "Ingreso de muestras", project.get("fecha_ingreso_muestra"))],
            [("event", "Fecha inicio proyecto", project.get("fecha_inicio_proyecto")),
             ("event_available", "Fecha final proyecto", project.get("fecha_final_proyecto")),
             ("event_available", "Fecha final real", project.get("fecha_final_real"))],
            [("inbox", "Fecha de recepción", project.get("fecha_recepcion")),
             ("science", "Fecha de ejecución", project.get("fecha_ejecucion")),
             ("outbox", "Fecha de emisión", project.get("fecha_emision"))],
        ]
        for gi, group in enumerate(date_groups):
            margen_inferior = "margin-bottom:6px;" if gi == len(date_groups) - 1 else ""
            rows_html = "".join(_info_row(icono, label, valor, i == 0) for i, (icono, label, valor) in enumerate(group))
            st.markdown(
                f'<div style="margin-top:16px;{margen_inferior}padding:10px 12px 12px;border-radius:10px;'
                f'background:rgba(74,120,98,0.06);border:1px solid rgba(74,120,98,0.16);">{rows_html}</div>',
                unsafe_allow_html=True,
            )

    if st.session_state.role == "jefe":
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Editar proyecto", icon=":material/edit:", use_container_width=True):
                navigate("edit-project")
        with c2:
            if confirm_delete(f"project_{codigo}", f"el proyecto {codigo} y todas sus perforaciones y muestras"):
                db.archive_project(project["id"])
                st.session_state.bitacora_draft = {k: v for k, v in st.session_state.bitacora_draft.items() if not k.startswith(codigo + "::")}
                navigate("home")
        if project_status(codigo) == "ejecutado":
            st.caption("Este proyecto ya está completamente ejecutado (todas sus muestras finalizadas).")
            if st.button("Desarchivar proyecto", icon=":material/unarchive:", use_container_width=True):
                desarchivar_proyecto(codigo)
                st.success("Proyecto desarchivado — vuelve a aparecer en Proyectos en ejecución.")
                st.rerun()

    with st.container(border=True):
        st.markdown('<div class="section-title">Progreso general (así avanzan los laboratoristas)</div>', unsafe_allow_html=True)
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
            # Avance por ENSAYO (no por muestra completa) — antes solo se veía entrando a la
            # perforación ("Avance del sondeo"); se trae acá para no tener que entrar solo para
            # ver esto.
            ensayos_completados, ensayos_total = _perforacion_ensayos_progress(codigo, perf["codigo"])
            st.markdown(f'<div class="cell-muted">{icon("check_circle", size=13)} {ensayos_completados} de '
                        f'{ensayos_total} ensayos completados</div>', unsafe_allow_html=True)
            st.progress(perf_pct / 100)

            # Muestras de esta perforación, con chips de estado por ensayo (semáforo: rojo/
            # amarillo/verde) y botón "Abrir" directo a cada una — ya no hace falta pasar por la
            # pantalla de la perforación solo para abrir un ensayo (ver
            # _render_tabla_muestras_con_semaforo, la misma tabla de "Orden de Laboratorio").
            # Esa pantalla se conserva para lo que sí sigue necesitando por fuera de esto:
            # exportar el perfil completo, o que el Jefe agregue una muestra nueva.
            if muestras:
                with st.expander(f"Ver muestras ({len(muestras)})", icon=":material/visibility:"):
                    _render_tabla_muestras_con_semaforo(sorted(muestras, key=lambda m: m["numero"]))

            bc1, bc2 = st.columns(2)
            with bc1:
                if st.button("Ver detalle del sondeo →", key=f"open_perf_{perf['codigo']}", use_container_width=True):
                    st.session_state.selected_perforacion = perf["codigo"]
                    navigate("perforacion-detail")
            with bc2:
                filas_perf = _bitacora_filas_perforacion(codigo, perf["codigo"])
                excel_bytes, _truncado = generar_excel_bitacora_orden(project, filas_perf, {perf["tipo"]})
                st.download_button("Descargar orden", data=excel_bytes, icon=":material/download:",
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
    # A diferencia de las demás, esta fecha se queda vacía (no en hoy) si el proyecto todavía
    # no ha terminado de verdad — por eso no usa _parse_fecha (esa cae a hoy si no hay valor).
    if f"edit_fecha_final_real_{codigo}" not in st.session_state:
        try:
            st.session_state[f"edit_fecha_final_real_{codigo}"] = date.fromisoformat(project.get("fecha_final_real"))
        except (TypeError, ValueError):
            st.session_state[f"edit_fecha_final_real_{codigo}"] = None
    fecha_final_real = st.date_input(
        "Fecha final real", key=f"edit_fecha_final_real_{codigo}", format="DD/MM/YYYY",
        help="Cuándo terminó de verdad el proyecto — distinta de la fecha final planeada.",
    )

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Cancelar", use_container_width=True):
            go_back(fallback="project-detail")
    with c2:
        if st.button("Guardar cambios", type="primary", use_container_width=True, icon=":material/save:", disabled=not nombre):
            db.update_project(
                project["id"], nombre=nombre, localizacion=localizacion, norma=norma,
                fecha_bitacora=str(fecha_bitacora), fecha_ingreso_muestra=str(fecha_ingreso),
                cliente=cliente, correo_cliente=correo_cliente,
                muestra_tomada_por=muestra_tomada_por, direccion_cliente=direccion_cliente,
                telefono_contacto=telefono_contacto, nombre_contacto=nombre_contacto,
                fecha_inicio_proyecto=str(fecha_inicio_proyecto), fecha_final_proyecto=str(fecha_final_proyecto),
                fecha_final_real=str(fecha_final_real) if fecha_final_real else None,
                fecha_recepcion=str(fecha_recepcion), fecha_ejecucion=str(fecha_ejecucion), fecha_emision=str(fecha_emision),
            )
            navigate("project-detail")


# ════════════════════════════════════════════════════════════════════
# DETALLE DE PERFORACIÓN → LISTA DE MUESTRAS
# ════════════════════════════════════════════════════════════════════
def _perforacion_ensayos_progress(codigo, perf_codigo):
    muestras = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    total_ensayos, completados = 0, 0
    for m in muestras:
        for label in unificar_ensayos([e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS]):
            total_ensayos += 1
            tipo_interno = SUPPORTED_ASSAY_MAP.get(label)
            a = get_assay(m["id_unico"], tipo_interno) if tipo_interno else None
            if a and a["status"] == "finalizado":
                completados += 1
    return completados, total_ensayos


def _render_tabla_muestras_con_semaforo(muestras):
    """Tabla de muestras con chips de color por ensayo (semáforo: rojo sin iniciar/amarillo en
    proceso/verde finalizado, ver STATUS_BADGE) y un botón "Abrir" que va directo al detalle de
    esa muestra — usada tanto en "Orden de Laboratorio" (Perforación) como en el desplegable
    "Ver muestras" del detalle de Proyecto, para no mantener la misma tabla duplicada en dos
    lugares con lógicas separadas."""
    # Los ensayos asignados van apilados uno debajo del otro (ver .assigned-chip-row), así que la
    # columna ya no necesita ancho para 3 chips lado a lado — se le puede devolver espacio a
    # "Tipo"/"Profundidad" sin que sus encabezados se partan en varias líneas.
    col_ratios = [0.8, 1.3, 1.4, 2.6, 1.0]
    headers = st.columns(col_ratios)
    for col, label in zip(headers, ["ID", "Tipo", "Profundidad", "Ensayos asignados", "Acción"]):
        col.markdown(f'<div class="assigned-th">{label}</div>', unsafe_allow_html=True)
    for i, m in enumerate(muestras):
        if i:
            st.markdown(f'<hr style="margin:8px 0;border-color:{BORDER};">', unsafe_allow_html=True)
        cols = st.columns(col_ratios, vertical_alignment="center")
        cols[0].markdown(f'<span class="cell-id">M-{html.escape(str(m["numero"]))}</span>', unsafe_allow_html=True)
        cols[1].markdown(f'<span class="cell-muted">{html.escape(m["tipo_muestra"])}</span>', unsafe_allow_html=True)
        cols[2].markdown(f'<span class="cell-muted">{m["profundidad_de"]}–{m["profundidad_hasta"]} m</span>', unsafe_allow_html=True)
        ensayos_sol = unificar_ensayos([e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS])
        chip_parts = []
        for e in ensayos_sol:
            tipo_interno = SUPPORTED_ASSAY_MAP.get(e)
            existing = get_assay(m["id_unico"], tipo_interno) if tipo_interno else None
            status = existing["status"] if existing else "sin-iniciar"
            chip_class = "assigned-chip assigned-chip-sm " + STATUS_BADGE[status].replace("badge-", "assigned-chip-")
            chip_parts.append(f'<span class="{chip_class}">{html.escape(e)}</span>')
        chips = (f'<div class="assigned-chip-row">{"".join(chip_parts)}</div>' if chip_parts
                 else '<span class="cell-muted">—</span>')
        cols[3].markdown(chips, unsafe_allow_html=True)
        with cols[4]:
            if st.button("Abrir", key=f"open_muestra_{m['id_unico']}", use_container_width=True):
                st.session_state.selected_muestra_id = m["id_unico"]
                navigate("muestra-detail")


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
            _render_tabla_muestras_con_semaforo(muestras)


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


def _sync_muestras_perforacion(perforacion_id, codigo, perf_codigo, nuevas_rows):
    """Reconcilia la tabla editada de muestras de una perforación existente contra lo que ya
    hay en Supabase: filas con id_unico ya conocido se actualizan, las nuevas se crean, y las
    que ya no aparecen (se borraron en el editor) se archivan."""
    actuales = st.session_state.muestras.get(f"{codigo}::{perf_codigo}", [])
    actuales_by_id_unico = {m["id_unico"]: m for m in actuales}
    vistos = set()
    for row in nuevas_rows:
        id_unico = row["id_unico"]
        vistos.add(id_unico)
        existente = actuales_by_id_unico.get(id_unico)
        campos = {k: v for k, v in row.items() if k != "id_unico"}
        if existente:
            db.update_muestra(existente["id"], **campos)
        else:
            db.create_muestra(perforacion_id, id_unico=id_unico, **campos)
    for id_unico, m in actuales_by_id_unico.items():
        if id_unico not in vistos:
            db.archive_muestra(m["id"])


def render_bitacora():
    require_role("jefe")
    if st.button("← Atrás"):
        go_back()
    st.markdown("## Orden de ensayos para laboratorio")

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
                db.create_perforacion(project["id"], tipo, consecutivo, codigo_perf)
                st.rerun()

        with st.expander("Importar desde Excel (plantilla GDA-FL-003 ya diligenciada)", icon=":material/upload_file:"):
            st.caption("Sube uno o varios archivos de la plantilla oficial \"Bitácora Orden para Ensayos "
                       "de Laboratorio\" que te haya mandado el cliente ya diligenciada — un archivo puede "
                       "traer un solo sondeo o varias hojas (una por sondeo/apique). Las perforaciones y "
                       "muestras que traiga se agregan a la tabla de abajo para que las revises antes de "
                       "guardar — con solo subir el archivo todavía no se guarda nada.")
            archivos = st.file_uploader("Archivos .xlsx", type=["xlsx"], accept_multiple_files=True,
                                          key=f"import_bitacora_{codigo}")
            if archivos and st.button("Importar de estos archivos", icon=":material/publish:",
                                        key=f"btn_import_bitacora_{codigo}"):
                advertencias_totales = []
                importadas = 0
                for archivo in archivos:
                    perfs_parseadas, advertencias, _hdr = parse_bitacora_orden_xlsx(archivo.name, archivo.getvalue())
                    advertencias_totales.extend(advertencias)
                    for pp in perfs_parseadas:
                        codigo_norm = pp["codigo"].strip().upper()
                        perf = next((p for p in perforaciones if p["codigo"].strip().upper() == codigo_norm), None)
                        if not perf:
                            consecutivo = len([p for p in perforaciones if p["tipo"] == pp["tipo"]]) + 1
                            perf = db.create_perforacion(project["id"], pp["tipo"], consecutivo, pp["codigo"])
                            perforaciones.append(perf)
                        key = f"{codigo}::{perf['codigo']}"
                        df_actual = _bitacora_draft_df(key, st.session_state.muestras.setdefault(key, []))
                        numeros_ya = set(df_actual["Número"].astype(str).str.strip())
                        filas_nuevas = [f for f in pp["filas"] if f["Número"] not in numeros_ya]
                        if filas_nuevas:
                            df_actual = pd.concat([df_actual, pd.DataFrame(filas_nuevas)], ignore_index=True)
                            df_actual = df_actual[df_actual["Número"].astype(str).str.strip() != ""]  # quita la fila vacía de arranque
                            st.session_state.bitacora_draft[key] = df_actual[BITACORA_BASE_COLS]
                            importadas += len(filas_nuevas)
                if importadas:
                    st.success(f"Se importaron {importadas} muestra(s). Revísalas en la tabla de abajo antes de guardar la bitácora.")
                for adv in advertencias_totales:
                    st.warning(adv)
                if importadas:
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
                db.archive_perforacion(perf["id"])
                st.session_state.bitacora_draft.pop(key, None)
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
                _sync_muestras_perforacion(perf["id"], codigo, perf["codigo"], nuevas)
                # Se descarta el draft cacheado para que, si se vuelve a abrir esta perforación,
                # se reconstruya desde las muestras recién guardadas (evita mostrar/pisar con una
                # tabla vieja lo que ya se guardó).
                st.session_state.bitacora_draft.pop(key, None)
            st.success("Bitácora guardada. Los laboratoristas ya pueden ver y digitar las muestras.")
            st.rerun()


# ════════════════════════════════════════════════════════════════════
# CLASIFICACIÓN USCS (ASTM D2487 / INV E-102), calculada en la app a partir de los datos ya
# digitados de Granulometría y Límites de Atterberg — para no tener que abrir el Excel solo para
# ver a qué grupo pertenece la muestra. OJO: no hay ningún dato digitado en la app que permita
# distinguir un suelo orgánico (color, olor) del inorgánico, así que siempre se asume inorgánico
# (igual que ya asume la descripción visual del laboratorista); por eso conviene que el Jefe/
# Director Técnico verifiquen los primeros resultados contra el Excel antes de confiar en ellos.
# ════════════════════════════════════════════════════════════════════
USCS_NOMBRES = {
    "GW": "Grava bien gradada", "GP": "Grava mal gradada",
    "GM": "Grava limosa", "GC": "Grava arcillosa",
    "GW-GM": "Grava bien gradada con limo", "GW-GC": "Grava bien gradada con arcilla",
    "GP-GM": "Grava mal gradada con limo", "GP-GC": "Grava mal gradada con arcilla",
    "SW": "Arena bien gradada", "SP": "Arena mal gradada",
    "SM": "Arena limosa", "SC": "Arena arcillosa",
    "SW-SM": "Arena bien gradada con limo", "SW-SC": "Arena bien gradada con arcilla",
    "SP-SM": "Arena mal gradada con limo", "SP-SC": "Arena mal gradada con arcilla",
    "CL": "Arcilla de baja plasticidad", "ML": "Limo de baja plasticidad",
    "CL-ML": "Arcilla limosa de baja plasticidad",
    "CH": "Arcilla de alta plasticidad", "MH": "Limo de alta plasticidad",
}


def _calcular_limites_atterberg(data):
    """LL, LP e IP a partir de las lecturas digitadas (INV E-125/E-126, equivalente a ASTM D4318):
    humedad = (masa húmeda - masa seca) / (masa seca - masa recipiente) x 100 por cada ensayo. El
    Límite Líquido es la humedad interpolada a 25 golpes sobre la curva de fluidez (humedad vs.
    log de golpes) — si algún ensayo se hizo exactamente a 25 golpes se usa esa lectura directa en
    vez de la regresión, igual que la fórmula de la plantilla de Excel. Devuelve (None, None, None)
    si no hay lecturas suficientes."""
    puntos = []
    for i in range(1, LIMITE_LIQUIDO_N + 1):
        golpes = to_float(data.get(f"lim_ll_golpes_{i}"))
        humedo = to_float(data.get(f"lim_ll_humedo_{i}"))
        seco = to_float(data.get(f"lim_ll_seco_{i}"))
        recip = to_float(data.get(f"lim_ll_recip_masa_{i}"))
        if None in (golpes, humedo, seco, recip) or golpes <= 0 or (seco - recip) <= 0:
            continue
        puntos.append((golpes, (humedo - seco) / (seco - recip) * 100))

    ll = None
    if puntos:
        exacto25 = [w for g, w in puntos if abs(g - 25) < 0.5]
        if exacto25:
            ll = exacto25[0]
        elif len(puntos) >= 2:
            xs = [math.log10(g) for g, _ in puntos]
            ys = [w for _, w in puntos]
            n = len(xs)
            mean_x, mean_y = sum(xs) / n, sum(ys) / n
            sxx = sum((x - mean_x) ** 2 for x in xs)
            if sxx == 0:
                ll = ys[0]
            else:
                pendiente = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / sxx
                intercepto = mean_y - pendiente * mean_x
                ll = intercepto + pendiente * math.log10(25)
        else:
            ll = puntos[0][1]

    humedades_lp = []
    for i in range(1, LIMITE_PLASTICO_N + 1):
        humedo = to_float(data.get(f"lim_lp_humedo_{i}"))
        seco = to_float(data.get(f"lim_lp_seco_{i}"))
        recip = to_float(data.get(f"lim_lp_recip_masa_{i}"))
        if None in (humedo, seco, recip) or (seco - recip) <= 0:
            continue
        humedades_lp.append((humedo - seco) / (seco - recip) * 100)
    lp = sum(humedades_lp) / len(humedades_lp) if humedades_lp else None

    if ll is None or lp is None:
        return None, None, None
    ll_i, lp_i = int(ll), int(lp)
    return ll_i, lp_i, max(ll_i - lp_i, 0)


def _calcular_curva_granulometrica(gran_data):
    """% que pasa cada tamiz, a partir de las masas retenidas digitadas y la masa inicial seca
    (derivada de las lecturas de Pasa No. 200 — la misma base que ya usa el Excel en D17). Un
    tamiz sin digitar cuenta como 0 g retenido, igual que en la tabla de solo lectura. Devuelve
    None si todavía no hay masa inicial válida (Pasa No. 200 sin digitar)."""
    masa_seco_mas_recip = to_float(gran_data.get("p200_seco_mas_recipiente_antes"))
    masa_recip = to_float(gran_data.get("p200_masa_recipiente_antes"))
    if masa_seco_mas_recip is None or masa_recip is None:
        return None
    masa_inicial = masa_seco_mas_recip - masa_recip
    if masa_inicial <= 0:
        return None

    puntos = []  # (apertura_mm, % que pasa), en orden de tamiz más grande a más chico
    acumulado = 0.0
    for key, _label, apert, _cell in SIEVES:
        retenido = to_float(gran_data.get(key)) or 0.0
        acumulado += retenido
        puntos.append((float(apert), max(0.0, 100 - acumulado / masa_inicial * 100)))

    pct_finos = puntos[-1][1]
    pct_pasa_4 = next((p for d, p in puntos if abs(d - 4.76) < 0.001), None)
    pct_grava = (100 - pct_pasa_4) if pct_pasa_4 is not None else None
    pct_arena = (pct_pasa_4 - pct_finos) if pct_pasa_4 is not None else None
    return {"puntos": puntos, "pct_finos": pct_finos, "pct_grava": pct_grava, "pct_arena": pct_arena}


def _interpolar_diametro(puntos, objetivo):
    """Diámetro (mm) para un % que pasa dado, interpolando log-lineal sobre la curva granulométrica
    (igual que se lee a mano sobre el papel semilogarítmico). None si el % pedido queda fuera del
    rango de tamices digitados."""
    crecientes = sorted(puntos, key=lambda t: t[0])  # de tamiz más chico a más grande
    for (d1, p1), (d2, p2) in zip(crecientes, crecientes[1:]):
        if p1 == p2 == objetivo:
            return d1
        if p1 <= objetivo <= p2 and p2 > p1:
            frac = (objetivo - p1) / (p2 - p1)
            return 10 ** (math.log10(d1) + frac * (math.log10(d2) - math.log10(d1)))
    return None


def clasificar_uscs(gran_data, lim_data):
    """Clasificación USCS (ASTM D2487) a partir de los datos ya digitados. Devuelve un dict con
    el símbolo y los valores intermedios (para verificarlos contra el Excel), o con la lista
    "faltantes" si todavía no hay datos suficientes para completarla."""
    curva = _calcular_curva_granulometrica(gran_data) if gran_data else None
    if curva is None:
        return {"faltantes": ["Faltan las lecturas de Pasa No. 200 (masa inicial de la muestra)."]}

    pct_finos, pct_grava, pct_arena = curva["pct_finos"], curva["pct_grava"], curva["pct_arena"]
    if pct_grava is None:
        return {"faltantes": ["Falta el retenido del tamiz No. 4."]}

    ll = lp = ip = None
    if lim_data:
        ll, lp, ip = _calcular_limites_atterberg(lim_data)

    def _simbolo_fino(ll, ip):
        a_line = 0.73 * (ll - 20)
        if ip < 4 or ip < a_line:
            return "M"
        if ip > 7 and ip >= a_line:
            return "C"
        return "C-M"  # zona rayada CL-ML

    resultado = {"faltantes": [], "pct_grava": pct_grava, "pct_arena": pct_arena, "pct_finos": pct_finos,
                 "ll": ll, "lp": lp, "ip": ip, "cu": None, "cc": None}

    if pct_finos >= 50:
        if ll is None:
            resultado["faltantes"].append("Falta digitar Límites de Atterberg — la muestra tiene 50% o más "
                                           "de finos y la clasificación depende de ellos.")
            return resultado
        base = _simbolo_fino(ll, ip)
        resultado["simbolo"] = "CL-ML" if base == "C-M" else f"{base}{'H' if ll >= 50 else 'L'}"
        return resultado

    d10 = _interpolar_diametro(curva["puntos"], 10)
    d30 = _interpolar_diametro(curva["puntos"], 30)
    d60 = _interpolar_diametro(curva["puntos"], 60)
    cu = (d60 / d10) if (d10 and d60) else None
    cc = ((d30 ** 2) / (d10 * d60)) if (d10 and d30 and d60) else None
    resultado["cu"], resultado["cc"] = cu, cc

    prefijo = "G" if pct_grava >= pct_arena else "S"
    umbral_cu = 4 if prefijo == "G" else 6
    bien_gradada = cu is not None and cc is not None and cu >= umbral_cu and 1 <= cc <= 3
    simbolo_gradacion = f"{prefijo}{'W' if bien_gradada else 'P'}"

    if pct_finos < 5:
        resultado["simbolo"] = simbolo_gradacion
    elif pct_finos > 12:
        if ll is None:
            resultado["faltantes"].append("Falta digitar Límites de Atterberg — la fracción fina de esta "
                                           "muestra supera el 12% y la clasificación depende de ellos.")
            return resultado
        base = _simbolo_fino(ll, ip)
        resultado["simbolo"] = f"{prefijo}{'C' if base == 'C-M' else base}"
    else:
        if ll is None:
            resultado["faltantes"].append("Falta digitar Límites de Atterberg — la fracción fina de esta "
                                           "muestra está entre 5% y 12% y la clasificación depende de ellos.")
            return resultado
        base = _simbolo_fino(ll, ip)
        resultado["simbolo"] = f"{simbolo_gradacion}-{prefijo}{'C' if base == 'C-M' else base}"

    return resultado


AASHTO_NOMBRES = {
    "A-1-a": "Fragmentos de piedra, grava y arena", "A-1-b": "Grava y arena fina",
    "A-3": "Arena fina",
    "A-2-4": "Grava y arena limosa o arcillosa", "A-2-5": "Grava y arena limosa o arcillosa",
    "A-2-6": "Grava y arena limosa o arcillosa", "A-2-7": "Grava y arena limosa o arcillosa",
    "A-4": "Suelo limoso", "A-5": "Suelo limoso",
    "A-6": "Suelo arcilloso", "A-7-5": "Suelo arcilloso", "A-7-6": "Suelo arcilloso",
}


def _grupo_aashto_a2(ll, indice_p):
    if indice_p < 10.5:
        return "A-2-4" if ll < 40.5 else "A-2-5"
    return "A-2-6" if ll < 40.5 else "A-2-7"


def clasificar_aashto(gran_data, lim_data):
    """Clasificación AASHTO (M 145) a partir de los mismos datos de Granulometría y Límites de
    Atterberg que la USCS — reimplementación fiel de la función AASH() de la plantilla oficial
    CLASIFICACION_DE_SUELOS.xlsm (Módulo11 del macro; ver oletools.olevba si hace falta releerla),
    para poder mostrarla también dentro de la app sin tener que abrir el Excel. Devuelve un dict
    con el símbolo y los valores intermedios, o con "faltantes" si aún no hay datos suficientes."""
    curva = _calcular_curva_granulometrica(gran_data) if gran_data else None
    if curva is None:
        return {"faltantes": ["Faltan las lecturas de Pasa No. 200 (masa inicial de la muestra)."]}

    pasa200 = curva["pct_finos"]
    pasa40 = next((p for d, p in curva["puntos"] if abs(d - 0.42) < 0.001), None)
    pasa10 = next((p for d, p in curva["puntos"] if abs(d - 2.00) < 0.001), None)
    if pasa40 is None or pasa10 is None:
        return {"faltantes": ["Faltan los retenidos de los tamices No. 10 y No. 40."]}

    ll = lp = None
    if lim_data:
        ll, lp, _ip = _calcular_limites_atterberg(lim_data)

    resultado = {"faltantes": [], "pasa200": pasa200, "pasa40": pasa40, "pasa10": pasa10, "ll": ll, "lp": lp}
    if ll is None or lp is None:
        resultado["faltantes"].append("Falta digitar Límites de Atterberg — la clasificación AASHTO "
                                       "siempre los necesita, incluso para suelos granulares.")
        return resultado

    indice_p = ll - lp
    if pasa200 <= 35:
        if pasa200 <= 25 and indice_p <= 6:
            if pasa40 <= 30:
                simbolo = ("A-1-a" if pasa10 <= 50 else "A-1-b") if pasa200 <= 15 else "A-1-b"
            elif pasa40 <= 50:
                simbolo = "A-1-b"
            elif pasa200 <= 10 and indice_p == 0:
                simbolo = "A-3"
            else:
                simbolo = _grupo_aashto_a2(ll, indice_p)
        else:
            simbolo = _grupo_aashto_a2(ll, indice_p)
    elif indice_p < 10.5:
        simbolo = "A-4" if ll < 40.5 else "A-5"
    elif ll < 40.5:
        simbolo = "A-6"
    else:
        simbolo = "A-7-5" if lp >= 30 else "A-7-6"

    resultado["simbolo"] = simbolo
    return resultado


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
        st.caption("Cómo se ve la muestra a ojo — para comparar con la clasificación USCS calculada de "
                   "los datos de laboratorio, justo abajo.")
        if st.session_state.role == "laboratorista":
            # Modo de digitación: con menús (estructurado, permite armar la comparación contra
            # la clasificación calculada — ver descripcion_visual_calculada) o texto libre, para
            # cuando la muestra no encaja bien en las opciones de los menús. No hace falta una
            # columna nueva para recordar el modo: se infiere de qué datos ya tiene la muestra —
            # si hay tipo de grano elegido, está en modo menús; si no hay tipo pero sí texto
            # libre guardado, está en modo texto; si no hay nada de nada, se parte de menús (el
            # modo recomendado por defecto, para poder comparar contra la USCS calculada).
            modos = ["Con menús (recomendado)", "Escribirla directamente"]
            modo_actual = modos[1] if (not muestra.get("desc_tipo_suelo") and muestra.get("descripcion_visual")) else modos[0]
            modo = st.radio("¿Cómo la vas a digitar?", modos, index=modos.index(modo_actual),
                             key=f"desc_modo_{muestra_id}", horizontal=True)

            if modo == modos[1]:
                st.caption("Al guardar en texto libre se borran los menús que hayas elegido antes — no se pueden "
                           "combinar los dos modos en la misma muestra.")
                texto_actual = muestra.get("descripcion_visual", "")
                texto = st.text_area(
                    "Descripción visual", value=texto_actual, key=f"desc_texto_libre_{muestra_id}",
                    label_visibility="collapsed", height=100,
                    placeholder="Ej: Grava arcillosa de color gris oscuro, débilmente cementada, en condición húmeda...",
                )
                if st.button("Guardar descripción visual", icon=":material/save:", key=f"desc_visual_save_texto_{muestra_id}"):
                    db.update_muestra(
                        muestra["id"], descripcion_visual=texto.strip().upper(),
                        desc_tipo_suelo=None, desc_tipo_secundario=None, desc_color=None, desc_subtonalidad=None,
                        desc_forma=None, desc_angulosidad=None, desc_cementacion=None, desc_consistencia=None,
                        desc_humedad=None,
                    )
                    st.success("Descripción visual guardada.")
                    st.rerun()
            else:
                # .upper() al comparar: datos de antes de este cambio (migración 0015) se
                # guardaron en minúscula/mixta (ej. "Limo") — así se siguen preseleccionando en
                # el menú en vez de aparecer en blanco solo porque el case no coincide con las
                # opciones nuevas.
                tipo_actual = (muestra.get("desc_tipo_suelo") or "").upper()
                tipo_idx = DESC_TIPO_SUELO_OPTIONS.index(tipo_actual) if tipo_actual in DESC_TIPO_SUELO_OPTIONS else 0
                desc_tipo = st.selectbox("Tipo de grano", DESC_TIPO_SUELO_OPTIONS, index=tipo_idx,
                                          key=f"desc_tipo_{muestra_id}", format_func=lambda v: v or "— Seleccionar —")
                grueso = _es_grueso(desc_tipo)

                # Componente secundario (ej. "grava con algo de arena", "arcilla con algo de
                # arena") — opcional, se oculta el selectbox si la casilla no está marcada en vez
                # de forzar a elegir "— Seleccionar —" cada vez. Se excluye el tipo principal de
                # la lista para no dejar armar "grava con algo de grava".
                secundario_actual = (muestra.get("desc_tipo_secundario") or "").upper()
                tiene_secundario = st.checkbox(
                    "¿Tiene un componente secundario? (ej. grava con algo de arena, arcilla con algo de arena)",
                    value=bool(secundario_actual), key=f"desc_tiene_sec_{muestra_id}")
                desc_secundario = None
                if tiene_secundario:
                    opciones_sec = [o for o in DESC_TIPO_SECUNDARIO_OPTIONS if o != desc_tipo]
                    sec_idx = opciones_sec.index(secundario_actual) if secundario_actual in opciones_sec else 0
                    desc_secundario = st.selectbox("Componente secundario", opciones_sec, index=sec_idx,
                                                    key=f"desc_sec_{muestra_id}")

                # Qué campos aparecen después depende del tipo de grano elegido arriba — grueso
                # (grava/arena) se describe por angulosidad y cementación, fino (limo/arcilla/
                # orgánico) por consistencia, y forma solo aplica a grava específicamente. Se
                # arman como lista en vez de columnas fijas porque cuáles aparecen cambia, y se
                # van dibujando en parejas de 2 columnas.
                campos = []
                if desc_tipo == "GRAVA":
                    campos.append(("Forma", DESC_FORMA_OPTIONS, "desc_forma", "forma"))
                if grueso:
                    campos.append(("Angulosidad", DESC_ANGULOSIDAD_OPTIONS, "desc_angulosidad", "angulosidad"))
                campos.append(("Color", DESC_COLOR_OPTIONS, "desc_color", "color"))
                campos.append(("Subtonalidad", DESC_SUBTONALIDAD_OPTIONS, "desc_subtonalidad", "subton"))
                campos.append(("Cementación", DESC_CEMENTACION_OPTIONS, "desc_cementacion", "cem") if grueso
                               else ("Consistencia", DESC_CONSISTENCIA_OPTIONS, "desc_consistencia", "cons"))
                campos.append(("Condición de humedad", DESC_HUMEDAD_OPTIONS, "desc_humedad", "hum"))

                valores = {}
                for i in range(0, len(campos), 2):
                    for col, (label, opciones, campo_db, campo_key) in zip(st.columns(2), campos[i:i + 2]):
                        with col:
                            actual = (muestra.get(campo_db) or "").upper()
                            idx = opciones.index(actual) if actual in opciones else 0
                            # La key incluye si el grano es grueso/fino, para forzar un widget
                            # nuevo si la persona cambia de escala — evita arrastrar en pantalla
                            # un valor que ya no aplica (ej. una cementación al pasar de grava a
                            # limo).
                            valores[campo_db] = st.selectbox(
                                label, opciones, index=idx, key=f"{campo_key}_{muestra_id}_{grueso}",
                                format_func=lambda v: v or "— Seleccionar —")

                if st.button("Guardar descripción visual", icon=":material/save:", key=f"desc_visual_save_{muestra_id}"):
                    guardar = {"desc_tipo_suelo": desc_tipo, "desc_tipo_secundario": desc_secundario,
                               "descripcion_visual": None, **valores}
                    # Los campos que no aparecieron arriba para este tipo de grano (ej. forma si
                    # no es grava) se limpian en vez de dejar guardado un valor viejo que ya no
                    # se ve.
                    for campo_db in ("desc_forma", "desc_angulosidad", "desc_cementacion", "desc_consistencia"):
                        guardar.setdefault(campo_db, None)
                    db.update_muestra(muestra["id"], **guardar)
                    st.success("Descripción visual guardada.")
                    st.rerun()
        else:
            descripcion_val = descripcion_visual_para_excel(muestra)
            if descripcion_val:
                st.markdown(f'<div style="display:flex;gap:10px;align-items:flex-start;background:{BG};'
                             f'border-radius:10px;padding:12px 14px;margin-bottom:4px;">'
                             f'<span style="margin-top:2px;">{icon("visibility", size=18)}</span>'
                             f'<div style="font-weight:600;line-height:1.5;">{html.escape(descripcion_val)}</div></div>',
                             unsafe_allow_html=True)
            else:
                st.markdown(f'<div style="display:flex;align-items:center;gap:6px;color:{NEUTRAL};font-style:italic;">'
                             f'{icon("visibility_off", size=16)} El laboratorista aún no la digita</div>', unsafe_allow_html=True)
            # Segunda versión, con el tipo de suelo que de verdad salió en la clasificación USCS
            # en vez del que se eligió a ojo — se muestra debajo de la inicial, no en su lugar
            # (ver descripcion_visual_calculada). Nada que mostrar hasta que haya datos de
            # Granulometría/Límites suficientes para calcularla.
            descripcion_calc = descripcion_visual_calculada(muestra)
            if descripcion_calc:
                st.markdown(f'<div style="display:flex;gap:10px;align-items:flex-start;background:{SECONDARY_CONTAINER};'
                             f'border-radius:10px;padding:12px 14px;margin-top:8px;">'
                             f'<span style="margin-top:2px;">{icon("science", size=18)}</span>'
                             f'<div><div class="cell-muted" style="margin-bottom:2px;">Según la clasificación USCS calculada</div>'
                             f'<div style="font-weight:600;line-height:1.5;color:{PRIMARY};">{html.escape(descripcion_calc)}</div></div></div>',
                             unsafe_allow_html=True)

    if muestra["ensayos"].get("Granulometría"):
        with st.container(border=True):
            st.markdown('<div class="section-title">Clasificación USCS</div>', unsafe_allow_html=True)
            st.caption("Calculada en la app con los datos ya digitados de Granulometría y Límites de "
                       "Atterberg — verifícala contra el Excel antes de usarla en un informe.")
            gran_assay = get_assay(muestra_id, "granulometria")
            lim_assay = get_assay(muestra_id, "limites")
            resultado = clasificar_uscs(
                gran_assay.get("data") if gran_assay else None,
                lim_assay.get("data") if lim_assay else None,
            )
            simbolo = resultado.get("simbolo")
            if simbolo:
                nombre = USCS_NOMBRES.get(simbolo, "")
                st.markdown(f'''
                    <div style="display:flex;align-items:center;gap:14px;background:{SECONDARY_CONTAINER};
                                border-radius:10px;padding:14px 16px;">
                        <div style="font-size:28px;font-weight:800;color:{PRIMARY};">{simbolo}</div>
                        <div style="font-weight:600;color:{PRIMARY};">{html.escape(nombre)}</div>
                    </div>
                ''', unsafe_allow_html=True)
                detalles = []
                if resultado.get("pct_grava") is not None:
                    detalles.append(f'Grava {resultado["pct_grava"]:.0f}% · Arena {resultado["pct_arena"]:.0f}% '
                                     f'· Finos {resultado["pct_finos"]:.0f}%')
                if resultado.get("ll") is not None:
                    detalles.append(f'LL {resultado["ll"]} · LP {resultado["lp"]} · IP {resultado["ip"]}')
                if resultado.get("cu") is not None and resultado.get("cc") is not None:
                    detalles.append(f'Cu {resultado["cu"]:.1f} · Cc {resultado["cc"]:.1f}')
                if detalles:
                    st.markdown(f'<div class="cell-muted" style="margin-top:10px;">{" · ".join(detalles)}</div>',
                                unsafe_allow_html=True)
            else:
                razones = resultado.get("faltantes") or ["Aún no hay datos suficientes para calcularla."]
                st.markdown(
                    f'<div style="display:flex;flex-direction:column;gap:8px;background:{BG};border-radius:10px;'
                    f'padding:12px 14px;margin-bottom:4px;">' + "".join(
                        f'<div style="display:flex;align-items:center;gap:8px;color:{NEUTRAL};font-style:italic;">'
                        f'{icon("hourglass_empty", size=18)} {html.escape(razon)}</div>' for razon in razones
                    ) + '</div>', unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown('<div class="section-title">Clasificación AASHTO</div>', unsafe_allow_html=True)
            st.caption("Calculada igual que la USCS de arriba, con los mismos datos de Granulometría y "
                       "Límites de Atterberg — verifícala contra el Excel antes de usarla en un informe.")
            resultado_aashto = clasificar_aashto(
                gran_assay.get("data") if gran_assay else None,
                lim_assay.get("data") if lim_assay else None,
            )
            simbolo_aashto = resultado_aashto.get("simbolo")
            if simbolo_aashto:
                nombre_aashto = AASHTO_NOMBRES.get(simbolo_aashto, "")
                st.markdown(f'''
                    <div style="display:flex;align-items:center;gap:14px;background:{SECONDARY_CONTAINER};
                                border-radius:10px;padding:14px 16px;">
                        <div style="font-size:28px;font-weight:800;color:{PRIMARY};">{simbolo_aashto}</div>
                        <div style="font-weight:600;color:{PRIMARY};">{html.escape(nombre_aashto)}</div>
                    </div>
                ''', unsafe_allow_html=True)
                detalles_aashto = []
                if resultado_aashto.get("pasa200") is not None:
                    detalles_aashto.append(f'Pasa No. 200 {resultado_aashto["pasa200"]:.0f}% · '
                                            f'Pasa No. 40 {resultado_aashto["pasa40"]:.0f}% · '
                                            f'Pasa No. 10 {resultado_aashto["pasa10"]:.0f}%')
                if resultado_aashto.get("ll") is not None:
                    detalles_aashto.append(f'LL {resultado_aashto["ll"]} · LP {resultado_aashto["lp"]}')
                if detalles_aashto:
                    st.markdown(f'<div class="cell-muted" style="margin-top:10px;">{" · ".join(detalles_aashto)}</div>',
                                unsafe_allow_html=True)
            else:
                razones_aashto = resultado_aashto.get("faltantes") or ["Aún no hay datos suficientes para calcularla."]
                st.markdown(
                    f'<div style="display:flex;flex-direction:column;gap:8px;background:{BG};border-radius:10px;'
                    f'padding:12px 14px;margin-bottom:4px;">' + "".join(
                        f'<div style="display:flex;align-items:center;gap:8px;color:{NEUTRAL};font-style:italic;">'
                        f'{icon("hourglass_empty", size=18)} {html.escape(razon)}</div>' for razon in razones_aashto
                    ) + '</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<div class="section-title">Observaciones</div>', unsafe_allow_html=True)
        st.caption("El laboratorista la digita si la muestra presenta fisuras o no se puede realizar el ensayo "
                   "por alguna razón.")
        if st.session_state.role == "laboratorista":
            with st.container(key="muestra-obs-box"):
                observacion = st.text_area(
                    "Observaciones", value=muestra.get("observaciones", ""), label_visibility="collapsed",
                    placeholder="Ej: Muestra con fisuras visibles, no fue posible completar el ensayo...", key=f"obs_{muestra_id}",
                )
            if st.button("Guardar observación", icon=":material/save:", key=f"obs_save_{muestra_id}"):
                db.update_muestra(muestra["id"], observaciones=observacion)
                st.success("Observación guardada.")
                st.rerun()
        else:
            observaciones_val = muestra.get("observaciones")
            if observaciones_val:
                st.markdown(f'<div class="cell-muted">{html.escape(observaciones_val)}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div style="display:flex;align-items:center;gap:8px;background:{BG};'
                             f'border-radius:10px;padding:12px 14px;margin-bottom:4px;color:{NEUTRAL};font-style:italic;">'
                             f'{icon("inbox", size=18)} Sin observaciones</div>', unsafe_allow_html=True)

    # Filtra por si la muestra guarda un ensayo que ya no es seleccionable (p. ej. "Pasa 200",
    # que quedó incluido dentro de Granulometría) — no se muestra aunque quede marcado en datos viejos.
    solicitados = unificar_ensayos([e for e, v in muestra["ensayos"].items() if v and e in BITACORA_ENSAYOS])
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
    else:
        finalizables = [e for e in solicitados if SUPPORTED_ASSAY_MAP.get(e)]
        todos_finalizados = bool(finalizables) and all(
            (get_assay(muestra_id, SUPPORTED_ASSAY_MAP[e]) or {}).get("status") == "finalizado" for e in finalizables)
        todos_aprobados = todos_finalizados and all(
            (get_assay(muestra_id, SUPPORTED_ASSAY_MAP[e]) or {}).get("etapa_revision") == "aprobado" for e in finalizables)
        if todos_aprobados:
            st.success("Todos los ensayos fueron aprobados por el Director Técnico — esta muestra está lista "
                        "para entregar al cliente.")
        elif todos_finalizados:
            st.info("Verifica que todas las aprobaciones del Director Técnico estén completas antes de dar "
                     "esta muestra por lista para el cliente.")

        for ensayo_label in solicitados:
            tipo_interno = SUPPORTED_ASSAY_MAP.get(ensayo_label)
            existing = get_assay(muestra_id, tipo_interno) if tipo_interno else None
            status = existing["status"] if existing else "sin-iniciar"
            etapa = existing.get("etapa_revision") if existing else None
            mostrar_aprobacion = bool(tipo_interno and existing and status == "finalizado")
            slug = re.sub(r"[^a-z0-9]+", "-", ensayo_label.lower())

            with st.container(border=True, key=f"ensayo-card-{slug}"):
                cols = st.columns([0.5, 2.0, 1.3, 0.9], vertical_alignment="center")
                cols[0].markdown(status_circle_html(status), unsafe_allow_html=True)
                with cols[1]:
                    st.markdown(f"**{ensayo_label}**")
                    if existing and existing.get("laboratorist"):
                        st.markdown(f'<div class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(existing["lastModified"])} · {existing["laboratorist"]}</div>', unsafe_allow_html=True)
                    elif existing:
                        st.markdown(f'<div class="timestamp-caption">{icon("history", size=13)} Última actualización: {format_dt(existing["lastModified"])}</div>', unsafe_allow_html=True)

                if not tipo_interno:
                    cols[2].markdown('<span class="badge badge-muted">Sin formulario aún</span>', unsafe_allow_html=True)
                else:
                    with cols[2]:
                        if mostrar_aprobacion:
                            st.markdown(aprobacion_badge_html(etapa), unsafe_allow_html=True)
                        else:
                            st.markdown(status_badge_html(status), unsafe_allow_html=True)

                    with cols[3]:
                        if st.button("Abrir", key=f"open_ensayo_{ensayo_label}", use_container_width=True):
                            if existing:
                                st.session_state.selected_assay_id = existing["id"]
                            else:
                                nuevo = db.create_assay(muestra["id"], tipo_interno)
                                st.session_state.selected_assay_id = nuevo["id"]
                            st.session_state.selected_assay_type = tipo_interno
                            navigate("assay-form")

                    # Fila de acciones (Confirmar/Devolver/Desconfirmar) en su propia línea a todo
                    # el ancho de la tarjeta, separada del título/badge/Abrir — antes compartían la
                    # misma fila angosta y "Confirmar"+"Devolver" quedaban tan apretados que el
                    # texto de "Devolver" se partía en dos líneas en pantallas angostas.
                    if not mostrar_aprobacion:
                        pass
                    elif etapa == "aprobado":
                        if st.session_state.role == "ingeniero":
                            with st.popover("Desconfirmar", use_container_width=True):
                                st.caption("Si los resultados no satisfacen al cliente, esto reabre el ensayo "
                                           "para el laboratorista y reinicia el ciclo de confirmación.")
                                motivo_desconf = st.text_area("Motivo", key=f"desconfirmar_motivo_{ensayo_label}",
                                                               placeholder="Qué hay que corregir...")
                                if st.button("Confirmar desconfirmación", key=f"desconfirmar_{ensayo_label}", use_container_width=True):
                                    if motivo_desconf.strip():
                                        db.ing_desconfirmar(existing["id"], st.session_state.profile, motivo_desconf)
                                        add_notification("laboratorista", f"El Director Técnico desconfirmó {ensayo_label} de la Muestra "
                                                                      f"{muestra['numero']} de {codigo}: {motivo_desconf}", codigo, perf_codigo, muestra_id)
                                        add_notification("jefe", f"El Director Técnico desconfirmó {ensayo_label} de la Muestra "
                                                                  f"{muestra['numero']} de {codigo}: {motivo_desconf}", codigo, perf_codigo, muestra_id)
                                        add_historial(existing, "Desconfirmado por el Director Técnico", f"Director Técnico: {motivo_desconf}",
                                                      icono="undo", tono="danger")
                                        st.success("Desconfirmado — vuelve al laboratorista.")
                                        st.rerun()
                                    else:
                                        st.error("Escribe el motivo antes de desconfirmar.")
                        else:
                            st.button("Confirmado", disabled=True, use_container_width=True, key=f"aprobado_ro_{ensayo_label}")
                    elif etapa == "pendiente_ing":
                        if st.session_state.role == "ingeniero":
                            b1, b2 = st.columns(2)
                            with b1:
                                if st.button("Confirmar", type="primary",
                                             use_container_width=True, key=f"ing_aprobar_{ensayo_label}"):
                                    db.ing_aprobar(existing["id"], st.session_state.profile)
                                    add_notification("jefe", f"El Director Técnico aprobó {ensayo_label} de la Muestra "
                                                              f"{muestra['numero']} de {codigo}.", codigo, perf_codigo, muestra_id)
                                    add_historial(existing, "Aprobación Final del Director Técnico", "Director Técnico",
                                                  icono="verified", tono="success")
                                    st.success("Aprobado.")
                                    st.rerun()
                            with b2:
                                with st.popover("Devolver", use_container_width=True):
                                    motivo_ing = st.text_area("Motivo", key=f"ing_motivo_{ensayo_label}",
                                                               placeholder="Qué hay que corregir...")
                                    if st.button("Confirmar devolución", key=f"ing_devolver_{ensayo_label}", use_container_width=True):
                                        if motivo_ing.strip():
                                            db.ing_devolver(existing["id"], st.session_state.profile, motivo_ing)
                                            add_notification("jefe", f"El Director Técnico devolvió {ensayo_label} de la Muestra "
                                                                      f"{muestra['numero']} de {codigo}: {motivo_ing}", codigo, perf_codigo, muestra_id)
                                            add_historial(existing, "Devuelto al Jefe de Laboratorio", f"Director Técnico: {motivo_ing}",
                                                          icono="undo", tono="danger")
                                            st.success("Devuelto al Jefe.")
                                            st.rerun()
                                        else:
                                            st.error("Escribe el motivo antes de devolver.")
                        elif st.session_state.role == "jefe":
                            with st.popover("Desconfirmar", use_container_width=True):
                                st.caption("Retira tu confirmación — el ensayo deja de estar pendiente de "
                                           "revisión del Director Técnico.")
                                motivo_jefe_desconf = st.text_area("Motivo", key=f"jefe_desconfirmar_motivo_{ensayo_label}",
                                                                    placeholder="Por qué te retractas...")
                                if st.button("Confirmar desconfirmación", key=f"jefe_desconfirmar_{ensayo_label}", use_container_width=True):
                                    if motivo_jefe_desconf.strip():
                                        db.jefe_desconfirmar(existing["id"], st.session_state.profile, motivo_jefe_desconf)
                                        add_notification("ingeniero", f"El Jefe de Laboratorio desconfirmó {ensayo_label} de la Muestra "
                                                                       f"{muestra['numero']} de {codigo} — ya no está pendiente de tu revisión.",
                                                          codigo, perf_codigo, muestra_id)
                                        add_historial(existing, "Desconfirmado por el Jefe de Laboratorio", f"Jefe de Laboratorio: {motivo_jefe_desconf}",
                                                      icono="undo", tono="danger")
                                        st.success("Desconfirmado.")
                                        st.rerun()
                                    else:
                                        st.error("Escribe el motivo antes de desconfirmar.")
                        else:
                            st.caption("Esperando al Director Técnico")
                    else:
                        if st.session_state.role == "jefe":
                            b1, b2 = st.columns(2)
                            with b1:
                                if st.button("Confirmar", type="primary",
                                             use_container_width=True, key=f"jefe_confirmar_{ensayo_label}"):
                                    db.jefe_confirmar(existing["id"], st.session_state.profile)
                                    add_notification("ingeniero", f"El Jefe de Laboratorio envió {ensayo_label} de la Muestra "
                                                                   f"{muestra['numero']} de {codigo} para tu confirmación final.",
                                                      codigo, perf_codigo, muestra_id)
                                    add_historial(existing, "Ensayo Confirmado", "Jefe de Laboratorio",
                                                  icono="check_circle", tono="success")
                                    st.success("Enviado al Director Técnico.")
                                    st.rerun()
                            with b2:
                                with st.popover("Devolver", use_container_width=True):
                                    motivo_jefe = st.text_area("Motivo", key=f"jefe_motivo_{ensayo_label}",
                                                                placeholder="Qué hay que corregir...")
                                    if st.button("Confirmar devolución", key=f"jefe_devolver_{ensayo_label}", use_container_width=True):
                                        if motivo_jefe.strip():
                                            db.jefe_devolver(existing["id"], st.session_state.profile, motivo_jefe)
                                            add_notification("laboratorista", f"El Jefe de Laboratorio devolvió {ensayo_label} de la Muestra "
                                                                          f"{muestra['numero']} de {codigo}: {motivo_jefe}", codigo, perf_codigo, muestra_id)
                                            add_historial(existing, "Devuelto al Laboratorista", f"Jefe de Laboratorio: {motivo_jefe}",
                                                          icono="undo", tono="danger")
                                            st.success("Devuelto al laboratorista.")
                                            st.rerun()
                                        else:
                                            st.error("Escribe el motivo antes de devolver.")
                        else:
                            st.caption("Pendiente del Jefe")

                if existing:
                    motivo = existing.get("motivo_rechazo")
                    if motivo:
                        quien = "el Jefe de Laboratorio" if existing.get("rechazado_por") == "jefe" else "el Director Técnico"
                        st.warning(f"Devuelto por {quien}: {motivo}")
                    # El historial se muestra siempre que exista (no solo cuando el ensayo está
                    # "Finalizado") — un desconfirmar/devolver reabre el ensayo pero no borra su
                    # registro de auditoría, y debe seguir siendo visible.
                    ens_historial = existing.get("historial", [])
                    if ens_historial:
                        with st.expander(f"Historial de Cambios ({len(ens_historial)})", icon=":material/history:"):
                            st.markdown(historial_timeline_html(ens_historial), unsafe_allow_html=True)


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
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_BITACORA_ORDEN), truncado


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
    # Descripción visual: primero la frase armada de los menús desplegables (+ notas
    # adicionales si hay, ver descripcion_visual_para_excel) — independiente de las
    # Observaciones—, si no hay nada de eso, lo que el laboratorista escribió en "Observaciones"
    # del propio ensayo, y solo si ninguna de las dos existe, el tipo de muestra como último recurso.
    ws["D13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or f"Tipo de muestra: {muestra.get('tipo_muestra','')}"


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


def _restaurar_orden_formato_condicional(xlsx_bytes, template_path):
    """openpyxl reordena los rangos de las reglas de formato condicional con varias áreas (ej.
    sqref="I6:M6 B6:B10 D6:D10 ...") al volver a guardar. La fórmula de esas reglas (ej.
    LEN(TRIM(B6))=0, las casillas verdes/grises que se apagan al escribir) es relativa a la
    primera celda del rango, así que con otro orden Excel las evalúa contra celdas equivocadas y
    el verde puede quedarse en casillas ya llenas. Se devuelve cada sqref al orden exacto de la
    plantilla original."""
    patron = re.compile(r'<conditionalFormatting sqref="([^"]*)">(.*?)</conditionalFormatting>', re.S)

    # Excel evalúa la fórmula relativa a la PRIMERA celda del sqref. Varias plantillas traen
    # reglas de varias áreas cuya fórmula apunta a otra celda (ej. sqref="G19:G20 G22:G23 E33:G34"
    # con LEN(TRIM(E19))=0), así que las casillas G19:G23 se quedan verdes aunque estén llenas —
    # pasa incluso escribiendo directo en la plantilla original. Se reescribe la fórmula para que
    # siempre revise la celda misma (la primera del sqref), que es lo que la regla quiere decir.
    # La fórmula relativa depende de qué celda toma Excel como base en reglas de varias áreas (varía
    # según el orden de los rangos), así que en vez de apuntar a una celda se usa una fórmula que
    # siempre revisa la celda misma: INDIRECT(ADDRESS(ROW(),COLUMN())).
    def reemplazo(m):
        sqref, cuerpo = m.group(1), m.group(2)
        nuevo = re.sub(r"LEN\(TRIM\(\$?[A-Z]+\$?\d+\)\)=0", "LEN(TRIM(INDIRECT(ADDRESS(ROW(),COLUMN()))))=0", cuerpo)
        return f'<conditionalFormatting sqref="{sqref}">{nuevo}</conditionalFormatting>'

    bio = BytesIO()
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as out, zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        for item in out.infolist():
            contenido = out.read(item.filename)
            if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                contenido = patron.sub(reemplazo, contenido.decode("utf-8")).encode("utf-8")
            z.writestr(item, contenido)
    return bio.getvalue()


def _reparar_enlaces_externos(xlsx_bytes):
    """Si la plantilla se guardó desde Excel con un vínculo a otro libro, openpyxl deja el
    externalLinkN.xml apuntando a un id (rId1) que ya no coincide con el del archivo de relaciones
    (rId2) — Excel dice que el archivo está dañado y no lo abre. Se iguala el id de la relación al
    que usa el externalLink."""
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zin:
        cambios = {}
        for nombre in zin.namelist():
            m = re.match(r"xl/externalLinks/externalLink(\d+)\.xml$", nombre)
            if not m:
                continue
            rels_nombre = f"xl/externalLinks/_rels/externalLink{m.group(1)}.xml.rels"
            if rels_nombre not in zin.namelist():
                continue
            usado = re.search(r'<externalBook[^>]*r:id="([^"]+)"', zin.read(nombre).decode("utf-8"))
            rels = zin.read(rels_nombre).decode("utf-8")
            ids = re.findall(r'Id="([^"]+)"', rels)
            if usado and len(ids) == 1 and ids[0] != usado.group(1):
                cambios[rels_nombre] = rels.replace(f'Id="{ids[0]}"', f'Id="{usado.group(1)}"').encode("utf-8")
        if not cambios:
            return xlsx_bytes
        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, cambios.get(item.filename, zin.read(item.filename)))
        return bio.getvalue()


def _restaurar_imagenes_perdidas(xlsx_bytes, template_path):
    reparado = _reparar_enlaces_externos(xlsx_bytes)
    return _restaurar_orden_formato_condicional(_restaurar_drawings_perdidos(reparado, template_path), template_path)


def _restaurar_drawings_perdidos(xlsx_bytes, template_path):
    """openpyxl vacía las <xdr:pic> (imágenes incrustadas — logo, firmas, diagramas de
    referencia de la hoja GUIA) de cualquier drawingN.xml al volver a guardar el archivo: el
    archivo drawingN.xml sigue existiendo y las formas/gráficos de adentro se conservan, pero
    las imágenes puntuales se pierden en blanco porque openpyxl no sabe reconstruir esa parte
    del XML al reescribirlo (a diferencia de _reparar_graficos_perdidos, que arregla archivos
    que openpyxl bota por completo, acá el archivo sigue estando pero vacío de imágenes).

    Restaurar solo el drawingN.xml (y su .rels) NO basta: openpyxl vuelve a numerar y a mezclar
    TODOS los archivos xl/media/imageN.png al guardar, así que el "image7.png" que deja openpyxl
    casi nunca es el mismo contenido que el "image7.png" de la plantilla, aunque el nombre
    coincida — restaurar solo el drawing dejaba el nombre correcto apuntando a la imagen
    equivocada (una firma se veía reemplazada por la etiqueta de otra parte de la hoja). Por eso
    también se restauran, con su contenido original, los xl/media/*.png que el drawing
    restaurado referencia — así drawing + rels + imagen quedan siempre como un trío consistente
    sacado íntegro de la plantilla, sin importar qué numeración haya usado openpyxl."""
    with zipfile.ZipFile(template_path) as tpl:
        tpl_names = set(tpl.namelist())
        drawings = {n for n in tpl_names if re.match(r"xl/drawings/drawing\d+\.xml$", n)}
        parches = {}
        with zipfile.ZipFile(BytesIO(xlsx_bytes)) as out:
            out_names = set(out.namelist())
            for nombre in drawings:
                tpl_xml = tpl.read(nombre).decode("utf-8")
                if "<xdr:pic>" not in tpl_xml or nombre not in out_names:
                    continue  # sin imágenes en la plantilla, o el archivo falta entero (lo arregla _reparar_graficos_perdidos)
                out_xml = out.read(nombre).decode("utf-8")
                if out_xml.count("<xdr:pic>") >= tpl_xml.count("<xdr:pic>"):
                    continue  # openpyxl sí las conservó, nada que restaurar
                parches[nombre] = tpl.read(nombre)
                rels_nombre = nombre.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels"
                if rels_nombre not in tpl_names:
                    continue
                rels_bytes = tpl.read(rels_nombre)
                parches[rels_nombre] = rels_bytes
                for target in re.findall(r'Target="(\.\./media/[^"]+)"', rels_bytes.decode("utf-8")):
                    media_nombre = "xl/" + target[len("../"):]
                    if media_nombre in tpl_names:
                        parches[media_nombre] = tpl.read(media_nombre)

            if not parches:
                return xlsx_bytes

            bio = BytesIO()
            with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
                for item in out.infolist():
                    if item.filename in parches:
                        z.writestr(item, parches.pop(item.filename))
                    else:
                        z.writestr(item, out.read(item.filename))
                for nombre, data in parches.items():
                    z.writestr(nombre, data)  # por si openpyxl nunca escribió ese archivo
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
        # Si solo se solicitó Pasa 200 (sin Granulometría), `gran_data` no trae los tamices en
        # absoluto (nunca pasó por el formulario de Granulometría) — ahí se dejan las celdas tal
        # cual trae la plantilla, para no simular una curva granulométrica falsa ("pasa 100%" en
        # todos los tamices). Pero si el tamiz SÍ se digitó (el ensayo de Granulometría existe y
        # su formulario ya se abrió), un tamiz que quedó en blanco significa "no quedó nada
        # retenido ahí" y se escribe como 0, igual que en la vista de solo lectura de la app.
        for key, _label, _apert, cell in SIEVES:
            if key not in gran_data:
                continue
            valor = to_float(gran_data.get(key))
            ws[cell] = valor if valor is not None else 0

    if lim_data is not None:
        _escribir_limites(ws, lim_data)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    data = _reparar_graficos_perdidos(bio.getvalue(), TEMPLATE_GRANULOMETRIA)
    return _restaurar_imagenes_perdidas(data, TEMPLATE_GRANULOMETRIA)


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
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_HUMEDAD)


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
    ws["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or f"Tipo de muestra: {muestra.get('tipo_muestra','')}"


def generar_excel_masa_unitaria(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    wb = load_workbook(TEMPLATE_MASA_UNITARIA)
    ws = wb["GUIA"]
    _llenar_encabezado_masa_unitaria(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo)

    ws["E20"] = to_float(data.get("mu_peso_aire"))  # B = masa en el aire
    ws["F20"] = to_float(data.get("mu_peso_aire_par"))  # C = masa en el aire parafinado
    ws["G20"] = to_float(data.get("mu_peso_agua_par"))  # D = masa parafinada sumergida
    # La densidad de la parafina (L24) no se digita en la app — se deja el 0.86 por defecto
    # que ya trae la plantilla.

    # La humedad (G28) la necesita la fórmula de "densidad seca". Se toma del ensayo de Humedad
    # de la misma muestra si tiene uno asignado (igual que Granulometría y Límites comparten
    # datos entre sí); si no tiene, se usa la que se digitó manualmente en este mismo ensayo.
    humedad_pct, _fuente = _mu_humedad_parafinado(data, muestra["id_unico"])
    if humedad_pct is not None:
        ws["G28"] = humedad_pct
    # A (masa de la cuerda, D20) y temperatura del agua no tienen celda equivalente en esta
    # plantilla — quedan para completar manualmente en el Excel.

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_MASA_UNITARIA)


def _llenar_encabezado_cbr(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo=""):
    """La plantilla oficial de CBR (GDA-FLC-013) usa una distribución de encabezado propia,
    distinta a la de Granulometría/Límites/Humedad/Masa Unitaria: etiquetas en B/H, valores en
    C/J en vez de D/K (ver hoja "GUIA" de GDA-FLC-013 CBR INALTERADO.xlsx)."""
    ws["C6"] = project.get("cliente", "") if project else ""  # Cliente
    ws["C7"] = project["nombre"] if project else codigo  # Proyecto
    ws["C8"] = project.get("correo_cliente", "") if project else ""  # Correo electrónico
    ws["C9"] = project.get("localizacion", "") if project else ""  # Localización
    ws["C10"] = project.get("muestra_tomada_por", "") if project else ""  # Muestra tomada por
    ws["J6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""  # Fecha de recepción
    ws["J7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""  # Fecha de ejecución
    ws["J8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""  # Fecha de emisión
    ws["K9"] = project.get("numero", "") if project else ""  # Código interno — número (J9 ya trae "GDA")
    ws["L9"] = project.get("anio", "") if project else ""  # Código interno — año

    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""  # Tipo de perforación (lista desplegable)
    ws["F12"] = perf["consecutivo"] if perf else ""  # Número de perforación
    ws["H12"] = muestra["numero"]  # Muestra No.
    prof_de, prof_hasta = to_float(muestra.get("profundidad_de")), to_float(muestra.get("profundidad_hasta"))
    ws["K12"] = f"{prof_de}-{prof_hasta}" if prof_de is not None and prof_hasta is not None else ""
    ws["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or f"Tipo de muestra: {muestra.get('tipo_muestra','')}"


def generar_excel_cbr(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """CBR (GDA-FLC-013, INV E-148-13). La plantilla real trae ~20 hojas ocultas más (proyectos
    anteriores, cada uno una copia de "GUIA" ya diligenciada) que no tienen nada que ver con esta
    muestra — se descartan todas menos "GUIA" antes de guardar, para no mandar datos de otros
    clientes en cada descarga ni inflar el archivo."""
    wb = load_workbook(TEMPLATE_CBR)
    for nombre in list(wb.sheetnames):
        if nombre != "GUIA":
            del wb[nombre]
    ws = wb["GUIA"]
    _llenar_encabezado_cbr(ws, codigo, perf_codigo, muestra, project, observaciones_ensayo)

    ws["D20"] = data.get("cbr_molde", "")
    ws["D21"] = to_float(data.get("cbr_diametro"))
    ws["D22"] = to_float(data.get("cbr_altura"))
    ws["D24"] = to_float(data.get("cbr_masa_molde"))
    ws["D23"] = to_float(data.get("cbr_masa_muestra_molde_antes"))
    ws["E23"] = to_float(data.get("cbr_masa_muestra_molde_despues"))

    # Humedad antes de inmersión: no se digita en el formulario de CBR, se comparte con el
    # ensayo de Contenido de Humedad de la misma muestra (ver render_cbr_form) — se copia acá
    # porque la plantilla real necesita el valor en su propia celda, no una referencia cruzada
    # a otro archivo.
    hum_assay = get_assay(muestra["id_unico"], "humedad")
    hum_data = hum_assay.get("data", {}) if hum_assay else {}
    ws["D30"] = hum_data.get("hum_recipiente", "")
    ws["D31"] = to_float(hum_data.get("hum_masa_humedo_mas_recipiente"))
    ws["D32"] = to_float(hum_data.get("hum_seco_mas_recipiente"))
    ws["D33"] = to_float(hum_data.get("hum_masa_recipiente"))

    ws["E30"] = data.get("cbr_desp_recipiente", "")
    ws["E31"] = to_float(data.get("cbr_desp_masa_humedo"))
    ws["E32"] = to_float(data.get("cbr_desp_masa_seco"))
    ws["E33"] = to_float(data.get("cbr_desp_masa_recipiente"))

    ws["D37"] = to_float(data.get("cbr_exp_lectura_inicial"))
    ws["D38"] = to_float(data.get("cbr_exp_lectura_final"))

    # Pesas de sobrecarga y tiempo de inmersión: la plantilla ya trae un valor por defecto
    # (4554 g / 4 días) — solo se pisa si el laboratorista digitó algo distinto.
    pesas_antes = to_float(data.get("cbr_pesas_antes"))
    if pesas_antes is not None:
        ws["I39"] = pesas_antes
    if data.get("cbr_rango"):
        ws["I37"] = data["cbr_rango"]
    pesas_despues = to_float(data.get("cbr_pesas_despues"))
    if pesas_despues is not None:
        ws["I40"] = pesas_despues
    tiempo_antes = to_float(data.get("cbr_tiempo_inmersion_antes"))
    if tiempo_antes is not None:
        ws["J39"] = tiempo_antes
    tiempo_despues = to_float(data.get("cbr_tiempo_inmersion_despues"))
    if tiempo_despues is not None:
        ws["J40"] = tiempo_despues

    # Tabla de penetración: fila 22 es la profundidad "0" — la plantilla trae sus celdas de
    # fuerza vacías, así que se llenan con 0; las 12 profundidades reales (CBR_PENETRACION_FILAS)
    # caen en las filas 23 a 34, en el mismo orden.
    ws["I22"] = to_float(data.get("cbr_pen_antes_0"), 0)
    ws["K22"] = to_float(data.get("cbr_pen_despues_0"), 0)
    for i in range(1, len(CBR_PENETRACION_FILAS) + 1):
        fila = 22 + i
        antes = to_float(data.get(f"cbr_pen_antes_{i}"))
        if antes is not None:
            ws[f"I{fila}"] = antes
        despues = to_float(data.get(f"cbr_pen_despues_{i}"))
        if despues is not None:
            ws[f"K{fila}"] = despues

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_CBR)


def _generar_excel_gravedad(template, hoja, celdas, codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Gravedad específica y absorción del agregado fino (GDA-FLC-027, INV E-222) o grueso
    (GDA-FLC-028, INV E-223). `celdas` mapea cada celda de entrada de la plantilla a su valor;
    las densidades y la absorción las calcula el propio Excel."""
    wb = load_workbook(template)
    ws = wb[hoja]
    ws["C6"] = project.get("cliente", "") if project else ""
    ws["C7"] = project["nombre"] if project else codigo
    ws["C8"] = project.get("correo_cliente", "") if project else ""
    ws["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["C10"] = project["muestra_tomada_por"]
    ws["I6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["I7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["I8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["J9"] = project.get("numero", "") if project else ""
    ws["K9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["D12"] = perf_codigo
    ws["F12"] = muestra["numero"]
    ws["I12"] = to_float(muestra.get("profundidad_de"))
    ws["K12"] = to_float(muestra.get("profundidad_hasta"))
    ws["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""
    for celda, valor in celdas.items():
        ws[celda] = to_float(valor)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), template)


def generar_excel_gravedad_fino(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    celdas = {"G19": data.get("gesp_f_masa_seco"), "G20": data.get("gesp_f_masa_pic_agua"),
              "G21": data.get("gesp_f_masa_pic_agua_suelo"), "G22": data.get("gesp_f_masa_sss"),
              "N54": data.get("gesp_f_pct_retenido"), "AD30": data.get("gesp_f_temp")}
    return _generar_excel_gravedad(TEMPLATE_GESP_FINO, "GUIA", celdas, codigo, perf_codigo, muestra, project, data,
                                    observaciones_ensayo)


def generar_excel_limite_contraccion(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Límite de contracción (GDA-FLC-022, INV E-129). Se llenan encabezado y datos; LC, R, Cv y Cl
    los calcula el Excel."""
    wb = load_workbook(TEMPLATE_LIMITE_CONTRACCION)
    ws = wb["GUIA"]
    ws["D6"] = project.get("cliente", "") if project else ""
    ws["D7"] = project["nombre"] if project else codigo
    ws["D8"] = project.get("correo_cliente", "") if project else ""
    ws["D9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["D10"] = project["muestra_tomada_por"]
    ws["L6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["L7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["L8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["M9"] = project.get("numero", "") if project else ""
    ws["N9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["F12"] = perf_codigo
    ws["I12"] = muestra["numero"]
    ws["L12"] = to_float(muestra.get("profundidad_de"))
    ws["N12"] = to_float(muestra.get("profundidad_hasta"))
    ws["D13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""
    ws["F20"] = data.get("lc_recipiente") or None
    ws["F21"] = to_float(data.get("lc_volumen"))
    ws["F22"] = to_float(data.get("lc_masa_recipiente"))
    ws["L20"] = to_float(data.get("lc_masa_aire"))
    ws["L21"] = to_float(data.get("lc_masa_aire_parafinada"))
    ws["L22"] = to_float(data.get("lc_masa_sumergida_parafinada"))
    for col, tipo in (("E", "natural"), ("F", "contraccion")):
        ws[f"{col}29"] = to_float(data.get(f"lc_{tipo}_hum_humedo"))
        ws[f"{col}30"] = to_float(data.get(f"lc_{tipo}_hum_seco"))
        ws[f"{col}31"] = to_float(data.get(f"lc_{tipo}_hum_recipiente"))
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_LIMITE_CONTRACCION)


def generar_excel_materia_organica(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Contenido de materia orgánica (GDA-FLC-003, INV E-121). Se llenan encabezado y las cuatro
    casillas de datos; la masa seca y el contenido (%) los calcula el Excel."""
    wb = load_workbook(TEMPLATE_MATERIA_ORGANICA)
    ws = wb["GUIA"]
    ws["D6"] = project.get("cliente", "") if project else ""
    ws["D7"] = project["nombre"] if project else codigo
    ws["D8"] = project.get("correo_cliente", "") if project else ""
    ws["D9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["D10"] = project["muestra_tomada_por"]
    ws["K6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["K7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["K8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["L9"] = project.get("numero", "") if project else ""
    ws["M9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["F12"] = perf_codigo
    ws["H12"] = muestra["numero"]
    ws["K12"] = to_float(muestra.get("profundidad_de"))
    ws["M12"] = to_float(muestra.get("profundidad_hasta"))
    ws["D13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""
    ws["I18"] = data.get("mo_recipiente") or None
    ws["I19"] = to_float(data.get("mo_masa_crisol_seco"))
    ws["I20"] = to_float(data.get("mo_masa_crisol_ignicion"))
    ws["I21"] = to_float(data.get("mo_masa_crisol"))
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_MATERIA_ORGANICA)


def generar_excel_proctor(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Proctor (GDA-FLC-002, INV E-141/E-142). La plantilla es un .xlsm que trae juntos el informe
    de Proctor (izquierda) y el de CBR (derecha) — acá solo se llena el Proctor; el CBR se descarga
    desde su propio ensayo. Los cálculos (humedad, densidades, humedad óptima, curva) los hace el
    Excel. El método (A/B/C), la preparación de la muestra, el martillo y el molde usado no se
    llenan: no se digitan en la bitácora."""
    wb = load_workbook(TEMPLATE_PROCTOR, keep_vba=True)
    ws = wb["Hoja1"]
    ws["D6"] = project.get("cliente", "") if project else ""
    ws["D7"] = project["nombre"] if project else codigo
    ws["D8"] = project.get("correo_cliente", "") if project else ""
    ws["D9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["D10"] = project["muestra_tomada_por"]
    ws["K6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["K7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["K8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["L9"] = project.get("numero", "") if project else ""
    ws["M9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["F12"] = perf_codigo
    ws["H12"] = muestra["numero"]
    ws["K12"] = to_float(muestra.get("profundidad_de"))
    ws["M12"] = to_float(muestra.get("profundidad_hasta"))
    ws["D13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    for i, col in enumerate("EFGH", start=1):
        def v(campo):
            return data.get(f"proc_{i}_{campo}")
        ws[f"{col}20"] = to_float(v("golpes"))
        ws[f"{col}21"] = to_float(v("capas"))
        ws[f"{col}22"] = to_float(v("masa_humedo_molde"))
        ws[f"{col}23"] = to_float(v("masa_molde"))
        ws[f"{col}24"] = to_float(v("volumen_molde"))
        ws[f"{col}28"] = (v("hum_recipiente") or None)
        ws[f"{col}29"] = to_float(v("hum_masa_humedo"))
        # La plantilla trae una sola fila de masa seca: se toma la última lectura (19, 18, 17 o 16 h).
        ws[f"{col}31"] = next((to_float(v(f"hum_seco_{h}h")) for h in (19, 18, 17, 16)
                               if to_float(v(f"hum_seco_{h}h")) is not None), None)
        ws[f"{col}32"] = to_float(v("hum_masa_recipiente"))
    tamiz = PROCTOR_TAMIZ_EXCEL.get(data.get("proc_sobretamano_tamiz"))
    if tamiz:
        ws["L25"] = tamiz
    ws["M25"] = to_float(data.get("proc_sobretamano_pct"))

    # CBR de suelos compactados (lado derecho de la plantilla): moldes en las columnas Q, R, S;
    # penetración en W, AA, AE; expansión en AF, AG, AH. Las celdas de expansión traen valores de
    # ejemplo (0.1 / 0.5), por eso se escriben siempre (vacías si no hay dato).
    for i, (col, col_pen, col_exp) in enumerate((("Q", "W", "AF"), ("R", "AA", "AG"), ("S", "AE", "AH")), start=1):
        def c(campo):
            return data.get(f"cbrc_{i}_{campo}")
        ws[f"{col}19"] = to_float(c("golpes"), to_float(CBRC_GOLPES[i - 1]))
        ws[f"{col}20"] = c("molde") or None
        ws[f"{col}21"] = to_float(c("capas"))
        ws[f"{col}22"] = to_float(c("masa_muestra_molde"))
        ws[f"{col}23"] = to_float(c("masa_molde"))
        altura, diametro = to_float(c("altura")), to_float(c("diametro"))
        ws[f"{col}24"] = altura
        ws[f"{col}25"] = diametro
        ws[f"{col}27"] = round(math.pi * (diametro / 2) ** 2 * altura, 2) if altura and diametro else None
        for fila_hum, pref, filas_xl in (("hc", "hc", (34, 35, 37, 38)), ("hd", "hd", (42, 43, 45, 46))):
            f_rec, f_hum, f_seco, f_masa = filas_xl
            ws[f"{col}{f_rec}"] = c(f"{pref}_recipiente") or None
            ws[f"{col}{f_hum}"] = to_float(c(f"{pref}_masa_humedo"))
            ws[f"{col}{f_seco}"] = next((to_float(c(f"{pref}_seco_{h}h")) for h in (19, 18, 17, 16)
                                         if to_float(c(f"{pref}_seco_{h}h")) is not None), None)
            ws[f"{col}{f_masa}"] = to_float(c(f"{pref}_masa_recipiente"))
        ws[f"{col_exp}45"] = to_float(c("exp_inicial"))
        ws[f"{col_exp}46"] = to_float(c("exp_final"))
        for j, fila in enumerate(CBRC_PEN_FILAS_EXCEL, start=1):
            ws[f"{col_pen}{fila}"] = to_float(c(f"pen_{j}"))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(_reparar_graficos_perdidos(bio.getvalue(), TEMPLATE_PROCTOR), TEMPLATE_PROCTOR)


def _fijar_celda_formula(xlsx_bytes, celda, formula):
    """Reemplaza la fórmula de una celda ya guardada, editando el XML de la primera hoja para no
    pasar otra vez por openpyxl (que rompería lo que _restaurar_imagenes_perdidas ya restauró)."""
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zin:
        nombre = "xl/worksheets/sheet1.xml"
        xml = zin.read(nombre).decode("utf-8")
        patron = re.compile(r'(<c r="' + celda + r'"[^>]*>)<f>[^<]*</f>(<v>[^<]*</v>|<v\s*/>)?')
        nuevo, n = patron.subn(lambda m: m.group(1) + f"<f>{formula}</f>", xml, count=1)
        if not n:
            return xlsx_bytes
        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, nuevo.encode("utf-8") if item.filename == nombre else zin.read(item.filename))
        return bio.getvalue()


def generar_excel_gravedad_arcilla(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """GDA-FLC-006 (INV E-128). G21 (masa del picnómetro lleno de agua) trae una fórmula que la
    busca en la calibración del picnómetro — solo se pisa si el laboratorista digitó el valor."""
    celdas = {"G19": data.get("gesp_a_picnometro"), "G20": data.get("gesp_a_masa_aire"),
              "G22": data.get("gesp_a_masa_pic_muestra"), "G23": data.get("gesp_a_temp"),
              "D32": data.get("gesp_a_pct_retenido")}
    if to_float(data.get("gesp_a_masa_pic_agua")) is not None:
        celdas["G21"] = data.get("gesp_a_masa_pic_agua")
    # G26 (total corregida) combinaba el % retenido con F51/F52, que en la plantilla son valores de
    # ejemplo (30 y 11) — daba un número que no correspondía a la muestra. Se deja el resultado
    # de la gravedad específica de la muestra ensayada, calculado en F48.
    excel = _generar_excel_gravedad(TEMPLATE_GESP_ARCILLA, "Hoja1", celdas, codigo, perf_codigo, muestra, project,
                                     data, observaciones_ensayo)
    return _fijar_celda_formula(excel, "G26", "F48")


def generar_excel_gravedad_grueso(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    celdas = {"G19": data.get("gesp_g_masa_seco"), "G20": data.get("gesp_g_masa_sumergido"),
              "G21": data.get("gesp_g_masa_sss"), "AD29": data.get("gesp_g_temp")}
    return _generar_excel_gravedad(TEMPLATE_GESP_GRUESO, "Hoja1 (2)", celdas, codigo, perf_codigo, muestra, project,
                                    data, observaciones_ensayo)


CORTE_TIPO_EXCEL = {"CD": "CONSOLIDADO DRENADO (CD)", "CU": "CONSOLIDADO NO DRENADO (CU)",
                    "UU": "NO CONSOLIDADO NO DRENADO (UU)"}
CORTE_CONDICION_EXCEL = {"Inalterada": "INALTERADA", "Remoldada": "REMOLDEADA", "Compactada": "COMPACTADA"}


def generar_excel_corte_directo(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Corte Directo (GDA-FLC-007, INV E-154). Se llena la hoja "1" (informe): encabezado,
    dimensiones y masas de las 3 probetas, humedad inicial/final y gravedad específica. Las hojas
    "2", "CARGA1" y "Fuente" (lecturas de deformación/carga de la máquina) NO se llenan: la app
    todavía no captura esas lecturas, así que el esfuerzo cortante, la cohesión y el ángulo de
    fricción quedan para completar en el Excel."""
    wb = load_workbook(TEMPLATE_CORTE_DIRECTO)
    ws = wb["1"]
    ws["C6"] = project.get("cliente", "") if project else ""
    ws["C7"] = project["nombre"] if project else codigo
    ws["C8"] = project.get("correo_cliente", "") if project else ""
    ws["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["C10"] = project["muestra_tomada_por"]
    ws["J6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["J7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["J8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["K9"] = project.get("numero", "") if project else ""
    ws["L9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["D12"] = perf_codigo
    ws["G12"] = muestra["numero"]
    ws["J12"] = to_float(muestra.get("profundidad_de"))
    ws["L12"] = to_float(muestra.get("profundidad_hasta"))
    ws["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    # Dimensiones del anillo: la hoja trae una sola casilla (el mismo anillo para las 3 probetas),
    # se toma la de la probeta 1.
    ws["D17"] = to_float(data.get("corte_m1_diametro_anillo"))
    ws["D18"] = to_float(data.get("corte_m1_altura_anillo"))
    ws["D19"] = to_float(data.get("corte_m1_masa_anillo"))
    for i, col in ((1, "I"), (2, "J"), (3, "K")):
        ws[f"{col}19"] = to_float(data.get(f"corte_m{i}_masa_inicial_anillo"))
        esfuerzo = to_float(data.get(f"corte_m{i}_esfuerzo_normal"))
        if esfuerzo is not None:
            ws[f"D{48 + i}"] = esfuerzo

    # Gravedad específica (INV E-128)
    ws["D25"] = to_float(data.get("corte_ge_temperatura"))
    ws["D26"] = to_float(data.get("corte_ge_masa_pic"))
    ws["D27"] = to_float(data.get("corte_ge_masa_pic_muestra"))
    ws["D28"] = to_float(data.get("corte_ge_masa_suelo_seco"))

    # Humedad: columnas G/H/I = inicial de las probetas 1/2/3, J/K/L = final.
    for i, (c_ini, c_fin) in ((1, ("G", "J")), (2, ("H", "K")), (3, ("I", "L"))):
        for sufijo, col in (("inicial", c_ini), ("final", c_fin)):
            base = f"corte_m{i}_"
            ws[f"{col}48"] = data.get(f"{base}hum_recipiente_{sufijo}", "") or None
            ws[f"{col}50"] = to_float(data.get(f"{base}hum_masa_humedo_{sufijo}"))
            seco = next((to_float(data.get(f"{base}hum_seco_{h}h_{sufijo}")) for h in (19, 18, 17)
                         if to_float(data.get(f"{base}hum_seco_{h}h_{sufijo}")) is not None), None)
            ws[f"{col}51"] = seco
            ws[f"{col}52"] = to_float(data.get(f"{base}hum_masa_recipiente_{sufijo}"))

    ws["C57"] = CORTE_TIPO_EXCEL.get(data.get("corte_tipo"), ws["C57"].value)
    ws["C58"] = CORTE_CONDICION_EXCEL.get(data.get("corte_condicion"), ws["C58"].value)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_CORTE_DIRECTO)


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

# ════════════════════════════════════════════════════════════════════
# CAMPOS REQUERIDOS POR ENSAYO — antes de "Enviar a revisión" se valida que estén digitados; si
# falta alguno, no se manda y ese campo se resalta en rojo (ver render_assay_form/campos_faltantes).
# Los 14h/15h/16h de Pasa 200 y Humedad NO están acá: son solo lecturas de verificación que se
# autocompletan de la lectura base, no algo que el laboratorista tenga que digitar aparte. Igual
# los tamices de Granulometría: un tamiz vacío es un dato válido (nada quedó retenido ahí), no un
# dato faltante — ver el "0 en vez de guion" ya implementado para esa tabla.
# ════════════════════════════════════════════════════════════════════
CAMPOS_REQUERIDOS_PASA200 = [
    (f"p200_{campo}_{suf}", f"{label} ({'antes' if suf == 'antes' else 'después'} del lavado)")
    for campo, label in (("recipiente", "Recipiente No."), ("seco_mas_recipiente", "Masa suelo seco + recipiente"),
                          ("masa_recipiente", "Masa del recipiente"))
    for suf in ("antes", "despues")
]
CAMPOS_REQUERIDOS_HUMEDAD = [
    ("hum_recipiente", "Recipiente No."),
    ("hum_masa_recipiente", "Masa del recipiente (g)"),
    ("hum_masa_humedo_mas_recipiente", "Masa suelo húmedo + recipiente (g)"),
    ("hum_seco_mas_recipiente", "Masa suelo seco + recipiente (g)"),
]
CAMPOS_REQUERIDOS_LIMITES = (
    [(f"lim_ll_{campo}_{i}", f"{label} — Límite Líquido, ensayo {i}")
     for campo, label in (("recipiente", "Recipiente No."), ("golpes", "No. de golpes"),
                           ("humedo", "Masa húmedo + recipiente"), ("seco", "Masa seco + recipiente"),
                           ("recip_masa", "Masa recipiente"))
     for i in range(1, LIMITE_LIQUIDO_N + 1)]
    + [(f"lim_lp_{campo}_{i}", f"{label} — Límite Plástico, ensayo {i}")
       for campo, label in (("recipiente", "Recipiente No."), ("humedo", "Masa húmedo + recipiente"),
                             ("seco", "Masa seco + recipiente"), ("recip_masa", "Masa recipiente"))
       for i in range(1, LIMITE_PLASTICO_N + 1)]
)
CAMPOS_REQUERIDOS_MASA_UNITARIA = [
    ("mu_peso_aire", "Masa en el aire (g)"),
    ("mu_peso_agua_par", "Masa en el agua parafinado (g)"),
    ("mu_peso_aire_par", "Masa en el aire parafinado (g)"),
    ("mu_temp_agua", "Temperatura del agua (°C)"),
]
# Reconstruido a partir de la plantilla oficial real (GDA-FLC-013 CBR INALTERADO.xlsx, hoja
# "GUIA", INV E-148-13) — la primera versión de este formulario se armó a partir de una captura
# de pantalla que resultó no coincidir con el formato oficial (traía "No. de golpes"/"No. de
# capas", que no existen en la plantilla real, y una humedad "después de inmersión" con lecturas
# de verificación a 16/17/18/19 horas que tampoco están ahí). Ver TEMPLATE_CBR/generar_excel_cbr
# para el mapeo exacto celda por celda.
# La "Humedad antes de inmersión" NO está acá ni se digita en el formulario del CBR: se comparte
# con el ensayo de Contenido de Humedad de la misma muestra (ver render_cbr_form) y se copia sola
# a la plantilla al exportar — si falta, se marca como faltante allá, no acá. La tabla de
# penetración (24 lecturas de fuerza) y las "Condiciones del Ensayo" (pesas/tiempo de inmersión,
# que ya traen un valor por defecto en la plantilla) tampoco son obligatorias.
CAMPOS_REQUERIDOS_CBR = [
    ("cbr_molde", "Molde No."),
    ("cbr_diametro", "Diámetro de la muestra (cm)"),
    ("cbr_altura", "Altura de la muestra (cm)"),
    ("cbr_masa_molde", "Masa molde (g)"),
    ("cbr_masa_muestra_molde_antes", "Masa de la muestra + molde (g) — antes de inmersión"),
    ("cbr_masa_muestra_molde_despues", "Masa de la muestra + molde (g) — después de inmersión"),
    ("cbr_desp_recipiente", "Recipiente — humedad después de inmersión"),
    ("cbr_desp_masa_humedo", "Peso recipiente + suelo húmedo (g) — después de inmersión"),
    ("cbr_desp_masa_seco", "Peso recipiente + suelo seco (g) — después de inmersión"),
    ("cbr_desp_masa_recipiente", "Peso recipiente (g) — después de inmersión"),
    ("cbr_exp_lectura_inicial", "Lectura inicial (in) — expansión"),
    ("cbr_exp_lectura_final", "Lectura final (in) — expansión"),
]
# Las 12 profundidades de penetración estándar de INV E-148 (pulgadas), con su equivalente en mm
# tal como aparece impreso en la plantilla — no se digitan, son fijas; lo que se digita es la
# Fuerza (kN) leída en cada una, antes y después de inmersión (ver render_cbr_form).
CBR_PENETRACION_FILAS = [
    ("0.005", "0.127"), ("0.025", "0.635"), ("0.05", "1.27"), ("0.075", "1.905"),
    ("0.1", "2.54"), ("0.125", "3.175"), ("0.15", "3.81"), ("0.175", "4.445"),
    ("0.2", "5.08"), ("0.3", "7.62"), ("0.4", "10.16"), ("0.5", "12.7"),
]

# Corte Directo (INV E-154 / ASTM D3080) — armado a partir de un formato físico en papel (todavía
# sin plantilla de Excel oficial conectada), así que estos campos pueden necesitar ajuste más
# adelante. "Página 1" del formato: datos de la muestra + humedad inicial/final por cada una de
# las 3 probetas ensayadas a distintos esfuerzos normales — no incluye las lecturas de
# deformación/carga del ensayo en sí (esas quedan para cuando se digitalice esa parte).
CORTE_TIPOS = ["CD", "CU", "UU"]
CORTE_TIPO_LABELS = {
    "CD": "Consolidada Drenada (CD)", "CU": "Consolidada No Drenada (CU)", "UU": "No Consolidada No Drenada (UU)",
}
CORTE_CONDICION_OPTIONS = ["Inalterada", "Remoldada", "Compactada"]
CORTE_TEMP_SECADO_OPTIONS = ["60 °C (Método A)", "110 °C (Método B)"]
CORTE_MUESTRA_CAMPOS = [
    ("masa_inicial_anillo", "Masa muestra inicial + anillo (g)"),
    ("masa_anillo", "Masa anillo (g)"),
    ("altura_anillo", "Altura anillo (cm)"),
    ("diametro_anillo", "Diámetro anillo (cm)"),
    ("velocidad_corte", "Velocidad de corte (mm/min)"),
    ("esfuerzo_normal", "Esfuerzo normal aplicado (kg/cm²)"),
]
CORTE_HUMEDAD_CAMPOS = [
    ("hum_recipiente", "Recipiente No."),
    ("hum_masa_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("hum_seco_17h", "Masa suelo seco + recipiente (g) — 17 horas"),
    ("hum_seco_18h", "Masa suelo seco + recipiente (g) — 18 horas"),
    ("hum_seco_19h", "Masa suelo seco + recipiente (g) — 19 horas"),
    ("hum_masa_recipiente", "Masa del recipiente (g)"),
]

def parse_cbr_penetracion_xlsx(file_bytes):
    """Lee el Excel que genera la prensa de CBR (hoja "Informe", tabla "RESULTADOS DEL ENSAYO":
    penetración en pulgadas en la columna V y fuerza en kN en la columna AB, filas 19 en
    adelante) y devuelve ({índice de CBR_PENETRACION_FILAS: fuerza kN como texto}, aviso)."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return {}, "No se pudo abrir el archivo como Excel."
    ws = wb["Informe"] if "Informe" in wb.sheetnames else wb.worksheets[0]
    por_pulgada = {}
    for row in ws.iter_rows(min_row=19, max_row=45, min_col=22, max_col=28, values_only=True):
        pulg, kn = row[0], row[6]
        if isinstance(pulg, (int, float)) and isinstance(kn, (int, float)):
            por_pulgada[round(float(pulg), 4)] = kn
    valores = {}
    for i, (pulg, _mm) in enumerate(CBR_PENETRACION_FILAS, start=1):
        kn = por_pulgada.get(round(float(pulg), 4))
        if kn is not None:
            valores[i] = f"{kn:.3f}".rstrip("0").rstrip(".") or "0"
    if not valores:
        return {}, ("No encontré la tabla de penetración/fuerza en este archivo (se espera la hoja \"Informe\" "
                    "con penetración en la columna V y fuerza kN en la columna AB).")
    faltan = len(CBR_PENETRACION_FILAS) - len(valores)
    return valores, (f"Faltaron {faltan} profundidad(es) en el archivo — complétalas a mano." if faltan else "")


CAMPOS_REQUERIDOS_POR_TIPO = {
    "humedad": CAMPOS_REQUERIDOS_HUMEDAD,
    "pasa200": CAMPOS_REQUERIDOS_PASA200,
    "granulometria": CAMPOS_REQUERIDOS_PASA200,  # comparte los mismos campos de Pasa 200 (embebido y "Requerido")
    "limites": CAMPOS_REQUERIDOS_LIMITES,
    "masa-unitaria": CAMPOS_REQUERIDOS_MASA_UNITARIA,
    "cbr": CAMPOS_REQUERIDOS_CBR,
}


def campos_faltantes(tipo, data):
    """Campos requeridos que todavía están vacíos para este tipo de ensayo — lista de
    (clave, etiqueta) en el mismo orden en que se digitan en el formulario."""
    # Peso Unitario Método B es un formulario nuevo (sin plantilla de Excel todavía) — como los demás ensayos
    # nuevos, todavía no tiene validación de campos obligatorios; los del Parafinado no aplican acá.
    if tipo == "masa-unitaria" and data.get("mu_metodo") == "Método B":
        return []
    return [(key, label) for key, label in CAMPOS_REQUERIDOS_POR_TIPO.get(tipo, [])
            if not str(data.get(key, "")).strip()]


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

        _campo("hum_recipiente", "Recipiente no.", placeholder="839")
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


def _mub_humedad(data):
    """Humedad (%) a partir de las masas digitadas (recipiente, húmeda, seca a 17/18/19h) — misma fórmula que los
    demás ensayos. None si todavía faltan datos."""
    humedo, rec = to_float(data.get("mub_hum_humedo")), to_float(data.get("mub_hum_masa_rec"))
    seco = next((v for v in (to_float(data.get(f"mub_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    if None in (humedo, seco, rec) or (seco - rec) == 0:
        return None
    return (humedo - seco) / (seco - rec) * 100


def _mu_humedad_manual(data):
    """Humedad (%) a partir de las masas digitadas aquí mismo (recipiente, húmeda, seca a
    17/18/19h) cuando la muestra no tiene un ensayo de Humedad asignado — misma fórmula que
    _mub_humedad y que el resto de ensayos. None si todavía faltan datos."""
    humedo, rec = to_float(data.get("mu_hum_humedo")), to_float(data.get("mu_hum_masa_rec"))
    seco = next((v for v in (to_float(data.get(f"mu_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    if None in (humedo, seco, rec) or (seco - rec) == 0:
        return None
    return (humedo - seco) / (seco - rec) * 100


def _mu_humedad_parafinado(data, muestra_id):
    """Humedad (%) a usar en Peso Unitario Parafinado: si la muestra tiene un ensayo de Humedad
    asignado, se copia de ahí (igual que en CBR); si no tiene ninguno asignado, se calcula a
    partir de las masas (recipiente, húmeda, seca) digitadas en este mismo ensayo. Devuelve
    (valor, fuente), con fuente en {"ensayo", "manual", None}."""
    hum_assay = get_assay(muestra_id, "humedad") if muestra_id else None
    if hum_assay:
        return calcular_humedad_pct(hum_assay.get("data", {})), "ensayo"
    manual = _mu_humedad_manual(data)
    return manual, ("manual" if manual is not None else None)


def resultados_masa_unitaria_parafinado(data, muestra_id):
    """Densidad húmeda y seca del método Parafinado (GDA-FLC-004), misma fórmula de la plantilla
    (G24, G25, G26, G27): densidad húmeda (g/cm³) = B / ( -(D-A) + C - ((C-B)/densidad_parafina) ),
    con A = masa de la cuerda (0, no se digita en la app), B = masa en el aire, C = masa en el aire
    parafinado, D = masa en el agua parafinado, densidad_parafina = 0.86 (valor por defecto de la
    plantilla). La humedad se toma del ensayo de Humedad de la misma muestra si tiene uno asignado
    (igual que en CBR); si no tiene, se usa la que se digite manualmente aquí — sin ninguna de las
    dos solo se puede mostrar la densidad húmeda, no la seca."""
    b = to_float(data.get("mu_peso_aire"))
    c = to_float(data.get("mu_peso_aire_par"))
    d = to_float(data.get("mu_peso_agua_par"))
    if b is None or c is None or d is None:
        return [], None, None
    denominador = -d + c - ((c - b) / 0.86)
    if not denominador:
        return [], None, None
    dens_humeda = b / denominador
    filas = [("Densidad húmeda (g/cm³)", fmt_num(dens_humeda, 3)),
             ("Densidad húmeda (kN/m³)", fmt_num(dens_humeda * 10, 2))]
    humedad_pct, fuente = _mu_humedad_parafinado(data, muestra_id)
    if humedad_pct is not None:
        etiqueta = "Humedad (%) — del ensayo de Humedad" if fuente == "ensayo" else "Humedad (%) — digitada aquí"
        filas.insert(0, (etiqueta, fmt_num(humedad_pct, 2)))
        if humedad_pct != -100:
            dens_seca = dens_humeda / (1 + humedad_pct / 100)
            filas += [("Densidad seca (g/cm³)", fmt_num(dens_seca, 3)),
                      ("Densidad seca (kN/m³)", fmt_num(dens_seca * 10, 2))]
    return filas, humedad_pct, fuente


def resultados_masa_unitaria_b(data):
    """Densidad húmeda y seca del Método B (INV/ASTM D7263-09, plantilla GDA-FLC-030 — una sola lectura, no
    promedio): volumen (cm³) = π·altura(mm)·(diámetro(mm)/2)²/1000, densidad húmeda = masa/volumen, densidad
    seca = densidad húmeda/(1+w/100), con w calculada de las masas de humedad. El kN/m³ es la propia plantilla:
    densidad (g/cm³) × 10 (no × 9.80665)."""
    altura, diametro, masa = to_float(data.get("mub_altura")), to_float(data.get("mub_diametro")), to_float(data.get("mub_masa"))
    if not altura or not diametro or masa is None:
        return []
    volumen = math.pi * altura * (diametro / 2) ** 2 / 1000
    if volumen == 0:
        return []
    dens_humeda = masa / volumen
    filas = [("Volumen de la muestra (cm³)", fmt_num(volumen, 2)),
             ("Densidad húmeda (g/cm³)", fmt_num(dens_humeda, 3)),
             ("Densidad húmeda (kN/m³)", fmt_num(dens_humeda * 10, 2))]
    w = _mub_humedad(data)
    if w is not None:
        filas.insert(0, ("Humedad (%)", fmt_num(w, 2)))
    if w is not None and w != -100:
        dens_seca = dens_humeda / (1 + w / 100)
        filas += [("Densidad seca (g/cm³)", fmt_num(dens_seca, 3)),
                  ("Densidad seca (kN/m³)", fmt_num(dens_seca * 10, 2))]
    return filas


def generar_excel_masa_unitaria_b(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Peso unitario Método B (GDA-FLC-030, D7263-09). Se escribe directo en el XML de la hoja (sin fórmulas
    compartidas ni calcChain en esta plantilla — más simple que Compresión); volumen y densidades las calcula
    el propio Excel a partir de altura, diámetro, masa y humedad."""
    c = {}
    c["C6"] = project.get("cliente", "") if project else ""
    c["C7"] = project["nombre"] if project else codigo
    c["C8"] = project.get("correo_cliente", "") if project else ""
    c["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["C10"] = project["muestra_tomada_por"]
    c["I6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["I7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["I8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["J9"] = project.get("numero", "") if project else ""
    c["K9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["D12"] = perf_codigo
    c["F12"] = muestra["numero"]
    c["I12"] = to_float(muestra.get("profundidad_de"))
    c["K12"] = to_float(muestra.get("profundidad_hasta"))
    c["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    c["D20"] = to_float(data.get("mub_altura"))
    c["E20"] = to_float(data.get("mub_diametro"))
    c["F20"] = to_float(data.get("mub_masa"))
    c["G28"] = _mub_humedad(data)

    with open(TEMPLATE_MASA_UNITARIA_B, "rb") as f:
        plantilla = f.read()
    return _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet1.xml", c),
                                                TEMPLATE_MASA_UNITARIA_B)


def render_masa_unitaria_form(data, assay_id, muestra_id=None):
    actual = data.get("mu_metodo", "Parafinado")
    data["mu_metodo"] = st.radio("Método", ["Parafinado", "Método B"], horizontal=True,
                                  index=["Parafinado", "Método B"].index(actual) if actual in ("Parafinado", "Método B") else 0,
                                  key=f"mu_metodo_{assay_id}")
    if data["mu_metodo"] == "Parafinado":
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Masas del Ensayo"), unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                data["mu_peso_aire"] = st.text_input("Masa en el aire (g)", value=data.get("mu_peso_aire", ""),
                                                      key=f"mu_peso_aire_{assay_id}", placeholder="245.80")
                data["mu_peso_agua_par"] = st.text_input("Masa en el agua parafinado (g)", value=data.get("mu_peso_agua_par", ""),
                                                          key=f"mu_peso_agua_par_{assay_id}", placeholder="138.20")
            with c2:
                data["mu_peso_aire_par"] = st.text_input("Masa en el aire parafinado (g)", value=data.get("mu_peso_aire_par", ""),
                                                          key=f"mu_peso_aire_par_{assay_id}", placeholder="258.30")
                data["mu_temp_agua"] = st.text_input("Temperatura del agua (°C)", value=data.get("mu_temp_agua", ""),
                                                      key=f"mu_temp_agua_{assay_id}", placeholder="22.0")
        hum_assay_mu = get_assay(muestra_id, "humedad") if muestra_id else None
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Humedad"), unsafe_allow_html=True)
            if hum_assay_mu:
                humedad_copiada = calcular_humedad_pct(hum_assay_mu.get("data", {}))
                st.markdown(param_table_html([("Humedad (%) — copiada del ensayo de Humedad de esta muestra",
                                               fmt_num(humedad_copiada, 2) if humedad_copiada is not None else None)]),
                            unsafe_allow_html=True)
                if humedad_copiada is None:
                    st.caption("El ensayo de Humedad de esta muestra todavía no tiene datos suficientes.")
            else:
                st.caption("Esta muestra no tiene un ensayo de Humedad asignado — digita las masas para que se "
                           "calcule aquí mismo (necesaria para la densidad seca).")

                def _campo_hum_mu(key, label, placeholder="0.00"):
                    row = st.columns([2.2, 1])
                    row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                    data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed", placeholder=placeholder)

                for key, label in MU_HUMEDAD_FILAS:
                    _campo_hum_mu(key, label, placeholder="" if key == "mu_hum_recipiente" else "0.00")
                    if key == "mu_hum_seco_17" and data.get("mu_hum_seco_17"):
                        for siguiente in ("mu_hum_seco_18", "mu_hum_seco_19"):
                            if not data.get(siguiente):
                                st.session_state[f"{siguiente}_{assay_id}"] = data["mu_hum_seco_17"]
                                data[siguiente] = data["mu_hum_seco_17"]
                humedad_calc = _mu_humedad_manual(data)
                if humedad_calc is not None:
                    st.markdown(param_table_html([("Humedad calculada (%)", fmt_num(humedad_calc, 2))]),
                                unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
            filas, humedad_pct, _fuente = resultados_masa_unitaria_parafinado(data, muestra_id)
            if filas:
                st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
                if humedad_pct is None:
                    st.caption("Falta la humedad (del ensayo de Humedad o digitada arriba) para calcular "
                               "la densidad seca — por ahora solo se muestra la húmeda.")
            else:
                st.caption("Se muestran a medida que se digitan las masas de arriba.")
        render_equipo(data, "mu", EQUIPO_MASA_UNITARIA)
    else:
        st.info("Formulario armado sobre la plantilla oficial GDA-FLC-030. El Excel para descargar está al final del ensayo.")

        def _campo_b(key, label, placeholder="0.00"):
            row = st.columns([2.2, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder=placeholder)

        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Peso Unitario Volumétrico"), unsafe_allow_html=True)
            _campo_b("mub_altura", "Altura de la muestra (mm)")
            _campo_b("mub_diametro", "Diámetro de la muestra (mm)")
            _campo_b("mub_masa", "Masa de la muestra (g)")
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
            st.caption("Con esto se calcula también la densidad seca; sin ella, solo la húmeda.")
            for key, label in MUB_HUMEDAD_FILAS:
                _campo_b(key, label, placeholder="" if key == "mub_hum_recipiente" else "0.00")
                if key == "mub_hum_seco_17" and data.get("mub_hum_seco_17"):
                    for siguiente in ("mub_hum_seco_18", "mub_hum_seco_19"):
                        if not data.get(siguiente):
                            st.session_state[f"{siguiente}_{assay_id}"] = data["mub_hum_seco_17"]
                            data[siguiente] = data["mub_hum_seco_17"]
        with st.container(border=True):
            st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
            filas = resultados_masa_unitaria_b(data)
            if filas:
                st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
            else:
                st.caption("Se muestran a medida que se digitan los datos de arriba.")
        render_equipo(data, "mub", EQUIPO_MASA_UNITARIA_B)
    render_norma_selector("masa-unitaria", data, "mu")


def render_cbr_form(data, assay_id, muestra_id):
    st.info("Estos datos se guardan tal cual y se llevan a la plantilla oficial de Excel — el CBR a 0.1\" y 0.2\" "
            "de penetración, igual que el resto de valores calculados, los saca el Excel, no la app.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    def _campo_antes_despues(key_base, label):
        """Un mismo dato con lectura antes Y después de inmersión (ej. masa de la muestra +
        molde, que cambia porque la muestra absorbe agua) — dos casillas lado a lado."""
        row = st.columns([2, 1, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[f"{key_base}_antes"] = row[1].text_input(f"{label} (antes)", value=data.get(f"{key_base}_antes", ""),
                                                        key=f"{key_base}_antes_{assay_id}", label_visibility="collapsed",
                                                        placeholder="Antes")
        data[f"{key_base}_despues"] = row[2].text_input(f"{label} (después)", value=data.get(f"{key_base}_despues", ""),
                                                          key=f"{key_base}_despues_{assay_id}", label_visibility="collapsed",
                                                          placeholder="Después")

    with st.container(border=True):
        st.markdown(card_header_html("science", "Datos Iniciales"), unsafe_allow_html=True)
        st.caption("Molde, diámetro, altura y masa del molde son los mismos antes y después de inmersión — "
                   "se digitan en \"Antes\" y se copian solos a \"Después\". Solo la masa de la muestra + "
                   "molde cambia (la muestra absorbe agua).")
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Antes</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Después</div>', unsafe_allow_html=True)

        def _campo_replicado(key, label, placeholder="0.00"):
            row = st.columns([2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder=placeholder)
            copia = html.escape(str(data[key])) if data[key] else "—"
            row[2].markdown(f'<div style="padding:8px 12px;border:1px solid {BORDER};border-radius:8px;'
                             f'background:{BG};color:{NEUTRAL};min-height:38px;">{copia}</div>', unsafe_allow_html=True)

        _campo_replicado("cbr_molde", "Molde No.", placeholder="1")
        _campo_replicado("cbr_diametro", "Diámetro de la muestra (cm)")
        _campo_replicado("cbr_altura", "Altura de la muestra (cm)")
        _campo_antes_despues("cbr_masa_muestra_molde", "Masa de la muestra + molde (g)")
        _campo_replicado("cbr_masa_molde", "Masa molde (g)")

        antes_mm = str(data.get("cbr_masa_muestra_molde_antes", "")).strip()
        despues_mm = str(data.get("cbr_masa_muestra_molde_despues", "")).strip()
        if antes_mm and antes_mm == despues_mm:
            st.markdown(f"<style>.st-key-cbr_masa_muestra_molde_despues_{assay_id} input {{ border: 2px solid #d32f2f !important; "
                        f"background-color: #fdecea !important; }}</style>", unsafe_allow_html=True)
            st.error("La masa de la muestra + molde después de inmersión es igual a la de antes — la muestra "
                     "debería haber absorbido agua. Revisa el dato.")

    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Humedad de inmersión"), unsafe_allow_html=True)
        st.caption("\"Antes\" se comparte con el ensayo de Contenido de Humedad de esta muestra — no se "
                   "digita aquí; si falta o está mal, corrígelo desde ese ensayo. \"Después\" sí es un dato "
                   "propio del CBR.")
        hum_assay = get_assay(muestra_id, "humedad")
        hum_data = hum_assay.get("data", {}) if hum_assay else {}
        hay_antes = bool(hum_data.get("hum_recipiente") or hum_data.get("hum_seco_mas_recipiente"))

        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Antes</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Después</div>', unsafe_allow_html=True)

        def _campo_humedad(label, valor_antes, key_despues, placeholder="0.00"):
            row = st.columns([2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            texto_antes = html.escape(str(valor_antes)) if valor_antes not in (None, "") else "—"
            color_antes = TEXT if valor_antes not in (None, "") else NEUTRAL
            row[1].markdown(f'<div style="padding-top:8px;text-align:center;color:{color_antes};">{texto_antes}</div>',
                             unsafe_allow_html=True)
            data[key_despues] = row[2].text_input(label, value=data.get(key_despues, ""), key=f"{key_despues}_{assay_id}",
                                                    label_visibility="collapsed", placeholder=placeholder)

        _campo_humedad("Recipiente", hum_data.get("hum_recipiente") if hay_antes else None,
                       "cbr_desp_recipiente", placeholder="839")
        _campo_humedad("Peso recipiente + suelo húmedo (g)",
                       hum_data.get("hum_masa_humedo_mas_recipiente") if hay_antes else None, "cbr_desp_masa_humedo")
        _campo_humedad("Peso recipiente + suelo seco (g)",
                       hum_data.get("hum_seco_mas_recipiente") if hay_antes else None, "cbr_desp_masa_seco")
        _campo_humedad("Peso recipiente (g)", hum_data.get("hum_masa_recipiente") if hay_antes else None,
                       "cbr_desp_masa_recipiente")

        if hay_antes:
            row = st.columns([2, 1, 1])
            row[0].markdown('<div style="padding-top:8px;">Humedad (%)</div>', unsafe_allow_html=True)
            row[1].markdown(f'<div style="padding-top:8px;text-align:center;font-weight:700;">'
                             f'{fmt_num(calcular_humedad_pct(hum_data), decimals=2)}</div>', unsafe_allow_html=True)
            row[2].markdown(f'<div style="padding-top:8px;text-align:center;color:{NEUTRAL};">—</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div style="display:flex;align-items:center;gap:6px;color:{NEUTRAL};font-style:italic;">'
                         f'{icon("visibility_off", size=16)} El ensayo de Contenido de Humedad de esta muestra '
                         f'todavía no tiene datos</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(card_header_html("straighten", "Datos de Expansión"), unsafe_allow_html=True)
        _campo("cbr_exp_lectura_inicial", "Lectura inicial (in)")
        _campo("cbr_exp_lectura_final", "Lectura final (in)")

    with st.container(border=True):
        st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
        st.caption("La plantilla ya trae un valor por defecto (4554 g / 4 días) — solo digita algo aquí si es distinto.")
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Antes</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Después</div>', unsafe_allow_html=True)
        _campo_antes_despues("cbr_pesas", "Pesas de sobrecarga (g)")
        _campo_antes_despues("cbr_tiempo_inmersion", "Tiempo de inmersión (días)")
        opciones_rango = ["", "5 kN", "50 kN"]
        actual_rango = data.get("cbr_rango", "")
        data["cbr_rango"] = st.selectbox("Rango de la prensa", opciones_rango,
                                          index=opciones_rango.index(actual_rango) if actual_rango in opciones_rango else 0,
                                          key=f"cbr_rango_{assay_id}", format_func=lambda x: x or "— elegir —")

    with st.container(border=True):
        st.markdown(card_header_html("show_chart", "Penetración"), unsafe_allow_html=True)
        st.caption("Fuerza (kN) leída en cada profundidad — el esfuerzo (MPa) y el CBR a 0.1\"/0.2\" los calcula el Excel.")
        with st.expander("Importar resultados desde el Excel de la prensa", icon=":material/upload_file:"):
            st.caption("Sube el Excel que genera la prensa (hoja \"Informe\") — se llenan solas las fuerzas en kN "
                       "de esa columna. Puedes corregir cualquier valor después.")
            for suf, titulo in (("antes", "Antes de inmersión"), ("despues", "Después de inmersión")):
                archivo = st.file_uploader(titulo, type=["xlsx"], key=f"cbr_pen_upload_{suf}_{assay_id}")
                if archivo and st.button(f"Cargar {titulo.lower()}", key=f"cbr_pen_cargar_{suf}_{assay_id}",
                                          icon=":material/publish:"):
                    valores, aviso = parse_cbr_penetracion_xlsx(archivo.getvalue())
                    if not valores:
                        st.error(aviso)
                    else:
                        st.session_state[f"cbr_pen_{suf}_0_{assay_id}"] = "0"
                        for i, v in valores.items():
                            st.session_state[f"cbr_pen_{suf}_{i}_{assay_id}"] = v
                        if aviso:
                            st.warning(aviso)
                        st.success(f"Se cargaron {len(valores)} valores ({titulo.lower()}).")

        head = st.columns([1.2, 1, 1])
        head[0].markdown('<div class="cell-muted" style="font-weight:700;">Profundidad</div>', unsafe_allow_html=True)
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Fuerza antes (kN)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Fuerza después (kN)</div>', unsafe_allow_html=True)
        row0 = st.columns([1.2, 1, 1])
        row0[0].markdown('<div style="padding-top:8px;">0.000" (0 mm)</div>', unsafe_allow_html=True)
        for c, suf in ((row0[1], "antes"), (row0[2], "despues")):
            data[f"cbr_pen_{suf}_0"] = c.text_input(f"Fuerza {suf} 0in", value=data.get(f"cbr_pen_{suf}_0", "0"),
                                                     key=f"cbr_pen_{suf}_0_{assay_id}", label_visibility="collapsed")
        for i, (pulg, mm) in enumerate(CBR_PENETRACION_FILAS, start=1):
            row = st.columns([1.2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{pulg}" ({mm} mm)</div>', unsafe_allow_html=True)
            data[f"cbr_pen_antes_{i}"] = row[1].text_input(f"Fuerza antes {pulg}in", value=data.get(f"cbr_pen_antes_{i}", ""),
                                                             key=f"cbr_pen_antes_{i}_{assay_id}", label_visibility="collapsed",
                                                             placeholder="kN")
            data[f"cbr_pen_despues_{i}"] = row[2].text_input(f"Fuerza después {pulg}in", value=data.get(f"cbr_pen_despues_{i}", ""),
                                                               key=f"cbr_pen_despues_{i}_{assay_id}", label_visibility="collapsed",
                                                               placeholder="kN")

    render_equipo(data, "cbr", EQUIPO_CBR)
    render_norma_selector("cbr", data, "cbr")


def render_corte_directo_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora oficial GDA-FL-006. Los datos se llevan a la plantilla "
            "de Excel GDA-FLC-007 al descargar (abajo, al final del ensayo).")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    def _radio(key, label, options):
        actual = data.get(key, options[0])
        data[key] = st.radio(label, options, index=options.index(actual) if actual in options else 0,
                              key=f"{key}_{assay_id}", horizontal=True)

    def _campo_inicial_final(key_base, label, placeholder="0.00"):
        row = st.columns([2, 1, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        key_ini, key_fin = f"{key_base}_inicial", f"{key_base}_final"
        data[key_ini] = row[1].text_input(f"{label} (inicial)", value=data.get(key_ini, ""),
                                            key=f"{key_ini}_{assay_id}", label_visibility="collapsed",
                                            placeholder=placeholder)
        data[key_fin] = row[2].text_input(f"{label} (final)", value=data.get(key_fin, ""),
                                            key=f"{key_fin}_{assay_id}", label_visibility="collapsed",
                                            placeholder=placeholder)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Datos del Ensayo"), unsafe_allow_html=True)
        _radio("corte_tipo", "Tipo de corte directo", CORTE_TIPOS)
        _radio("corte_condicion", "Condición de la muestra", CORTE_CONDICION_OPTIONS)
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Inicial</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Final</div>', unsafe_allow_html=True)
        _campo_inicial_final("corte_temp", "Temperatura (°C)")
        _campo_inicial_final("corte_hum", "Humedad (%)")

    for i in (1, 2, 3):
        with st.container(border=True):
            st.markdown(card_header_html("science", f"Muestra {i} — Datos de la Muestra"), unsafe_allow_html=True)
            for campo, label in CORTE_MUESTRA_CAMPOS:
                _campo(f"corte_m{i}_{campo}", label)

        with st.container(border=True):
            st.markdown(card_header_html("water_drop", f"Muestra {i} — Contenido de Humedad"), unsafe_allow_html=True)
            head = st.columns([2, 1, 1])
            head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Inicial</div>', unsafe_allow_html=True)
            head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Final</div>', unsafe_allow_html=True)
            for campo, label in CORTE_HUMEDAD_CAMPOS:
                _campo_inicial_final(f"corte_m{i}_{campo}", label)
            _radio(f"corte_m{i}_temp_secado", "Temperatura de secado", CORTE_TEMP_SECADO_OPTIONS)
            st.markdown('<div class="cell-muted" style="font-weight:700;margin-top:8px;">Equipos utilizados (humedad)</div>',
                        unsafe_allow_html=True)
            sel = set(data.get(f"corte_m{i}_hum_equipos", []))
            nuevos = []
            cols_eq = st.columns(2)
            for j, equipo in enumerate(EQUIPO_CORTE_HUMEDAD):
                with cols_eq[j % 2]:
                    if st.checkbox(equipo, value=equipo in sel, key=f"corte_m{i}_hum_equipo_{j}_{assay_id}"):
                        nuevos.append(equipo)
            data[f"corte_m{i}_hum_equipos"] = nuevos

    with st.container(border=True):
        st.markdown(card_header_html("science", "Gravedad Específica (INV E-128-13)"), unsafe_allow_html=True)
        for campo, label in CORTE_GRAVEDAD_CAMPOS:
            _campo(f"corte_ge_{campo}", label)
        st.markdown('<div class="cell-muted" style="font-weight:700;margin-top:8px;">Equipos utilizados (gravedad)</div>',
                    unsafe_allow_html=True)
        _campo("corte_ge_picnometro", "Picnómetro No.", placeholder="1")
        sel = set(data.get("corte_ge_equipos", []))
        nuevos = []
        cols_eq = st.columns(2)
        for j, equipo in enumerate(EQUIPO_CORTE_GRAVEDAD):
            with cols_eq[j % 2]:
                if st.checkbox(equipo, value=equipo in sel, key=f"corte_ge_equipo_{j}_{assay_id}"):
                    nuevos.append(equipo)
        data["corte_ge_equipos"] = nuevos

    render_equipo(data, "corte", EQUIPO_CORTE_DIRECTO)
    render_norma_selector("corte-directo", data, "corte")


@st.cache_data
def _tabla_agua_temperatura():
    """Tabla 1 de la plantilla GDA-FLC-006: {temperatura °C: (densidad del agua, coeficiente K)}."""
    tabla = {}
    ws = load_workbook(TEMPLATE_GESP_ARCILLA, data_only=True, read_only=True)["Hoja1"]
    for row in ws.iter_rows(min_row=4, max_row=163, min_col=41, max_col=43, values_only=True):
        t, dens, k = row
        if isinstance(t, (int, float)) and isinstance(dens, (int, float)) and isinstance(k, (int, float)):
            tabla[round(float(t), 1)] = (dens, k)
    return tabla


def resultados_gravedad(data, modo):
    """Mismas fórmulas que las plantillas de Excel (GDA-FLC-027/028/006), para mostrar el
    resultado en la app antes de descargar. Devuelve [(etiqueta, valor)]; lista vacía si todavía
    faltan datos. modo: 'fino' (INV E-222), 'grueso' (INV E-223) o 'arcilla' (INV E-128)."""
    def f(k):
        return to_float(data.get(k))
    if modo == "fino":
        a, b, c, s = f("gesp_f_masa_seco"), f("gesp_f_masa_pic_agua"), f("gesp_f_masa_pic_agua_suelo"), f("gesp_f_masa_sss")
        if None in (a, b, c, s) or a == 0 or (b + s - c) == 0 or (b + a - c) == 0:
            return []
        seca, sss, aparente = a / (b + s - c), s / (b + s - c), a / (b + a - c)
        return [("Densidad relativa seca al horno", fmt_num(seca)), ("Densidad relativa SSS", fmt_num(sss)),
                ("Densidad relativa aparente", fmt_num(aparente)),
                ("Densidad seca al horno (kg/m³)", fmt_num(997.5 * seca, 1)),
                ("Densidad SSS (kg/m³)", fmt_num(997.5 * sss, 1)),
                ("Densidad aparente (kg/m³)", fmt_num(997.5 * aparente, 1)),
                ("Absorción (%)", fmt_num((s - a) / a * 100, 2))]
    if modo == "grueso":
        a, b, c = f("gesp_g_masa_seco"), f("gesp_g_masa_sumergido"), f("gesp_g_masa_sss")
        if None in (a, b, c) or a == 0 or (c - b) == 0 or (a - b) == 0:
            return []
        seca, sss, aparente = a / (c - b), c / (c - b), a / (a - b)
        return [("Densidad relativa seca al horno", fmt_num(seca)), ("Densidad relativa SSS", fmt_num(sss)),
                ("Densidad relativa aparente", fmt_num(aparente)),
                ("Densidad seca al horno (kg/m³)", fmt_num(997.5 * seca, 1)),
                ("Densidad SSS (kg/m³)", fmt_num(997.5 * sss, 1)),
                ("Densidad aparente (kg/m³)", fmt_num(997.5 * aparente, 1)),
                ("Absorción (%)", fmt_num((c - a) / a * 100, 2))]
    ws_, wpw, wpws, temp = f("gesp_a_masa_aire"), f("gesp_a_masa_pic_agua"), f("gesp_a_masa_pic_muestra"), f("gesp_a_temp")
    if None in (ws_, wpw, wpws, temp):
        return []
    dens_k = _tabla_agua_temperatura().get(round(temp, 1))
    if not dens_k or (wpw + ws_ - wpws) == 0:
        return []
    gt = ws_ * dens_k[0] / (wpw + ws_ - wpws)
    return [("Peso específico del agua (g/cm³)", fmt_num(dens_k[0], 5)),
            ("Gravedad específica", fmt_num(gt, 4))]


def resultados_materia_organica(data):
    """Mismas fórmulas de la plantilla GDA-FLC-003: masa seca y contenido de materia orgánica (%)."""
    a, b, c = (to_float(data.get(k)) for k, _ in MO_CAMPOS)
    if None in (a, b, c) or (b - c) == 0:
        return []
    return [("Masa seca (g)", fmt_num(b - c)), ("Contenido de materia orgánica (%)", fmt_num((a - b) / (b - c) * 100, 2))]


def resultados_limite_contraccion(data):
    """Mismas fórmulas de la plantilla GDA-FLC-022 (parafina: 0.86 g/cm³, agua: 1 g/cm³)."""
    def f(k):
        return to_float(data.get(k))
    vol, aire, aire_paraf, sum_paraf = f("lc_volumen"), f("lc_masa_aire"), f("lc_masa_aire_parafinada"), f("lc_masa_sumergida_parafinada")
    h_c, s_c, r_c = f("lc_contraccion_hum_humedo"), f("lc_contraccion_hum_seco"), f("lc_contraccion_hum_recipiente")
    if None in (vol, aire, aire_paraf, sum_paraf, h_c, s_c, r_c) or (s_c - r_c) == 0:
        return []
    vo = (aire_paraf - sum_paraf) - (aire_paraf - aire) / 0.86
    if vo == 0:
        return []
    w_c = (h_c - s_c) / (s_c - r_c)
    lc = w_c - (vol - vo) / (s_c - r_c)
    r = (s_c - r_c) / vo
    cv = (w_c - lc) * 100 * r
    cl = 100 * (1 - (100 / (100 + cv)) ** (1 / 3))
    return [("Volumen de la muestra sin parafina, Vo (cm³)", fmt_num(vo, 2)), ("Límite de contracción, LC (%)", fmt_num(lc * 100, 2)),
            ("Relación de contracción, R", fmt_num(r, 3)), ("Cambio volumétrico, Cv (%)", fmt_num(cv, 2)),
            ("Contracción lineal, Cl (%)", fmt_num(cl, 2))]

def parse_maquina_consolidacion_xlsx(file_bytes):
    """Excel de la máquina de consolidación: hoja "Data2", tiempo en segundos (columna A) y deformación en mm
    (columna D). Devuelve ([[segundos, deformación mm]], aviso)."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return [], "No se pudo abrir el archivo como Excel."
    nombre = next((n for n in wb.sheetnames if n.strip().lower() == "data2"), None)
    if nombre is None:
        return [], "No encontré la hoja \"Data2\" (la que arroja la máquina)."
    puntos = {}
    for row in wb[nombre].iter_rows(min_row=2, max_col=4, values_only=True):
        t, d = row[0], row[3]
        if isinstance(t, (int, float)) and isinstance(d, (int, float)):
            puntos[float(t)] = round(float(d), 4)
    if not puntos:
        return [], "La hoja \"Data2\" no trae tiempo (columna A) y deformación (columna D)."
    return [[t, puntos[t]] for t in sorted(puntos)], ""


def _cons_deformacion_en(puntos, t_seg):
    """Deformación interpolada en el tiempo t (segundos); None si la máquina ya no tiene datos a ese tiempo."""
    if not puntos or t_seg > puntos[-1][0]:
        return None
    if t_seg <= puntos[0][0]:
        return puntos[0][1]
    lo, hi = 0, len(puntos) - 1
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if puntos[mid][0] <= t_seg:
            lo = mid
        else:
            hi = mid
    (t0, d0), (t1, d1) = puntos[lo], puntos[hi]
    return d0 if t1 == t0 else d0 + (d1 - d0) * (t_seg - t0) / (t1 - t0)


def _cons_final(data, clave_fin, clave_ini):
    """Valor de la columna final; si no se digitó, repite el de la columna inicial."""
    v = to_float(data.get(clave_fin))
    return v if v is not None else to_float(data.get(clave_ini))


def _cons_humedad(data, prefijo):
    """Humedad (%) de la bitácora de consolidación: usa la última pesada de secado digitada."""
    def f(k):
        return to_float(data.get(f"cons_{prefijo}_{k}"))
    seco = next((v for v in (f("seco_19"), f("seco_18"), f("seco_17")) if v is not None), None)
    humedo, rec = f("humedo"), f("recipiente_masa")
    if None in (humedo, seco, rec) or (seco - rec) == 0:
        return None
    return (humedo - seco) / (seco - rec) * 100


def _cons_gravedad(data):
    """Gravedad específica corregida (misma fórmula de la plantilla GDA-FLC-009: Ws·ρ/(Wpw+Ws−Wpws)·K)."""
    ws_, wpws, temp = (to_float(data.get(k)) for k in ("cons_pic_masa_seco", "cons_pic_masa_agua_suelo", "cons_pic_temp"))
    pic = to_float(data.get("cons_pic_no"))
    if None in (ws_, wpws, temp, pic) or int(pic) not in CONS_PIC_CALIBRACION:
        return None
    dens_k = _tabla_agua_temperatura().get(round(temp, 1))
    if not dens_k:
        return None
    a, b = CONS_PIC_CALIBRACION[int(pic)]
    wpw = a * temp + b
    if (wpw + ws_ - wpws) == 0:
        return None
    return ws_ * dens_k[0] / (wpw + ws_ - wpws) * dens_k[1]


def resultados_consolidacion(data):
    """Mismas fórmulas de la hoja Resultados de la plantilla GDA-FLC-009 que dependen solo de la
    bitácora (datos de la muestra, humedades y gravedad específica). Índices de compresión, Cv, etc.
    dependen de las lecturas de la máquina y se calculan en el Excel."""
    def f(k):
        return to_float(data.get(k))
    filas = []
    ini, fin, anillo = f("cons_masa_anillo_muestra_ini"), f("cons_masa_anillo_muestra_fin"), f("cons_masa_anillo")
    anillo_fin = _cons_final(data, "cons_fin_masa_anillo", "cons_masa_anillo")
    mt = ini - anillo if None not in (ini, anillo) else None
    mf = fin - anillo_fin if None not in (fin, anillo_fin) else None
    if mt is not None:
        filas.append(("Masa de la muestra inicial (g)", fmt_num(mt, 2)))
    if mf is not None:
        filas.append(("Masa de la muestra final (g)", fmt_num(mf, 2)))
    wi, wf = _cons_humedad(data, "ini"), _cons_humedad(data, "fin")
    if wi is not None:
        filas.append(("Humedad inicial de la muestra (%)", fmt_num(wi, 2)))
    if wf is not None:
        filas.append(("Humedad final de la muestra (%)", fmt_num(wf, 2)))
    gs = _cons_gravedad(data)
    if gs is not None:
        filas.append(("Gravedad específica, Gs", fmt_num(gs, 4)))
    d, h = f("cons_diametro"), f("cons_altura")
    if None in (d, h) or d <= 0 or h <= 0:
        return filas
    d, h = d / 10, h / 10
    area = math.pi * d ** 2 / 4
    vol = area * h
    filas += [("Área de la muestra (cm²)", fmt_num(area, 2)), ("Volumen inicial (cm³)", fmt_num(vol, 2))]
    if mt is not None:
        filas.append(("Peso unitario total inicial (g/cm³)", fmt_num(mt / vol, 3)))
        if wi is not None:
            filas.append(("Peso unitario total seco inicial (g/cm³)", fmt_num(mt / vol / (1 + wi / 100), 3)))
    if mt is not None and mf is not None and wf is not None and gs:
        md = mf / (1 + wf / 100)
        vs = md / gs
        hs = vs / area
        if hs > 0 and h > hs:
            filas += [("Masa seca, Md (g)", fmt_num(md, 2)), ("Volumen de sólidos (cm³)", fmt_num(vs, 2)),
                      ("Hs (cm)", fmt_num(hs, 3)), ("Relación de vacíos inicial, e₀", fmt_num((h - hs) / hs, 3)),
                      ("Saturación inicial (%)", fmt_num((mt - md) / (area * (h - hs)) * 100, 1))]
    return filas


def render_consolidacion_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora GDA-FL-017. Las lecturas de deformación de la máquina no se digitan acá: "
            "se pegan en el Excel de la plantilla oficial (GDA-FLC-009), que se descarga al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    def _radio(key, label, opciones):
        actual = data.get(key, opciones[0])
        data[key] = st.radio(label, opciones, horizontal=True, index=opciones.index(actual) if actual in opciones else 0,
                              key=f"{key}_{assay_id}")

    with st.container(border=True):
        st.markdown(card_header_html("science", "Muestra"), unsafe_allow_html=True)
        _radio("cons_condicion", "Condición inicial", CONS_CONDICIONES)
        for key, label in (("cons_temp_ini", "Temperatura inicial (°C)"), ("cons_temp_fin", "Temperatura final (°C)")):
            _campo(key, label, placeholder="0.0")
    with st.container(border=True):
        st.markdown(card_header_html("science", "Datos de la Muestra"), unsafe_allow_html=True)
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Inicial</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Final</div>', unsafe_allow_html=True)
        for label, k_ini, k_fin in CONS_DATOS_MUESTRA:
            row = st.columns([2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            data[k_ini] = row[1].text_input(f"{label} — inicial", value=data.get(k_ini, ""), key=f"{k_ini}_{assay_id}",
                                             label_visibility="collapsed", placeholder="0.00")
            repite = k_fin not in ("cons_masa_anillo_muestra_fin",)
            data[k_fin] = row[2].text_input(f"{label} — final", value=data.get(k_fin, ""), key=f"{k_fin}_{assay_id}",
                                             label_visibility="collapsed",
                                             placeholder=(data.get(k_ini) or "igual al inicial") if repite else "0.00")
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Humedad de la Muestra"), unsafe_allow_html=True)
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Inicial</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Final</div>', unsafe_allow_html=True)
        for campo, label in CONS_HUMEDAD_FILAS:
            row = st.columns([2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            for col_i, pref in ((1, "ini"), (2, "fin")):
                key = f"cons_{pref}_{campo}"
                data[key] = row[col_i].text_input(f"{label} — {pref}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed")
        _radio("cons_temp_secado", "Temperatura de secado", CONS_TEMP_SECADO)
        _radio("cons_metodo", "Método utilizado", CONS_METODOS)
    with st.container(border=True):
        st.markdown(card_header_html("science", "Gravedad Específica"), unsafe_allow_html=True)
        for key, label in CONS_GS_CAMPOS:
            _campo(key, label)
    with st.container(border=True):
        st.markdown(card_header_html("straighten", "Deformaciones (mm)"), unsafe_allow_html=True)
        st.caption("Se guardan en la app como respaldo de la bitácora; la plantilla toma las lecturas de la máquina.")
        head = st.columns([1.4, 1, 1])
        head[0].markdown('<div class="cell-muted" style="font-weight:700;">Carga (kg)</div>', unsafe_allow_html=True)
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Carga</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Descarga</div>', unsafe_allow_html=True)
        for i, carga in enumerate(CONS_CARGAS, start=1):
            row = st.columns([1.4, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{carga}</div>', unsafe_allow_html=True)
            for col_i, tipo in ((1, "carga"), (2, "descarga")):
                key = f"cons_def_{i}_{tipo}"
                data[key] = row[col_i].text_input(f"{carga} kg — {tipo}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed")
        _campo("cons_precarga", "Precarga (g)")
    with st.container(border=True):
        st.markdown(card_header_html("show_chart", "Lecturas de la Máquina"), unsafe_allow_html=True)
        st.caption("Sube el Excel que arroja la máquina (hoja \"Data2\", columna D): se guardan las lecturas de cada carga y "
                   "van directo a la hoja \"DATOS MAQUINA\" del Excel que descargas. Un archivo por carga.")
        for par in range(0, len(CONS_MAQ_BLOQUES), 2):
            cols = st.columns(2)
            for col, i in zip(cols, (par + 1, par + 2)):
                enc = CONS_MAQ_BLOQUES[i - 1][0]
                with col:
                    st.markdown(f'<div style="font-weight:700;padding-top:6px;">Carga {enc}</div>', unsafe_allow_html=True)
                    archivo = st.file_uploader(f"Excel de la máquina — {enc}", type=["xlsx"], key=f"cons_maq_upload_{i}_{assay_id}",
                                                label_visibility="collapsed")
                    if archivo and st.button(f"Cargar {enc}", key=f"cons_maq_cargar_{i}_{assay_id}", icon=":material/publish:",
                                              use_container_width=True):
                        puntos, aviso = parse_maquina_consolidacion_xlsx(archivo.getvalue())
                        if not puntos:
                            st.error(aviso)
                        else:
                            data[f"cons_maq_{i}"] = puntos
                            if _guardar_inmediato(assay_id, data):
                                st.success(f"Se cargaron y guardaron {len(puntos)} lecturas.")
                            else:
                                data.pop(f"cons_maq_{i}", None)
                    pts = data.get(f"cons_maq_{i}")
                    if pts:
                        st.markdown(f'<div class="cell-muted">Cargado: {len(pts)} lecturas, hasta {pts[-1][0] / 60:.0f} min</div>',
                                    unsafe_allow_html=True)
                        if st.button("Quitar", key=f"cons_maq_quitar_{i}_{assay_id}", use_container_width=True):
                            data.pop(f"cons_maq_{i}", None)
                            if _guardar_inmediato(assay_id, data):
                                st.rerun()
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_consolidacion(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digitan los datos de arriba.")
    with st.container(border=True):
        st.markdown(card_header_html("construction", "Consolidómetro"), unsafe_allow_html=True)
        actual = data.get("cons_consolidometro", "")
        data["cons_consolidometro"] = st.selectbox(
            "Consolidómetro", CONS_CONSOLIDOMETROS, index=CONS_CONSOLIDOMETROS.index(actual) if actual in CONS_CONSOLIDOMETROS else 0,
            key=f"cons_consolidometro_{assay_id}", label_visibility="collapsed", format_func=lambda x: x or "Selecciona el equipo")
    render_equipo(data, "cons", EQUIPO_CONSOLIDACION)
    render_norma_selector("consolidacion", data, "cons")


def generar_excel_consolidacion(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Consolidación unidimensional (GDA-FLC-009, INV E-151). Se llena la hoja "Resultados": encabezado,
    datos de la muestra, humedades, gravedad específica y condiciones. Las lecturas de la máquina (hojas de
    cada carga, "Datos " y "DATOS MAQUINA ") quedan para los humanos; con ellas el Excel calcula Cc, Cr, Cv…
    Además corrige tres referencias de la plantilla que la dejaban sin gravedad específica: la calibración
    del picnómetro leía la temperatura de una celda vacía (G23), K30 leía una hoja GS vacía y el % que pasa
    el tamiz No. 4 (Z30) restaba de X30 en vez de Z28."""
    wb = load_workbook(TEMPLATE_CONSOLIDACION)
    ws = wb["Resultados"]
    ws["D6"] = project.get("cliente", "") if project else ""
    ws["D7"] = project["nombre"] if project else codigo
    ws["D8"] = project.get("correo_cliente", "") if project else ""
    ws["D9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        ws["D10"] = project["muestra_tomada_por"]
    ws["M6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    ws["M7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    ws["M8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    ws["N9"] = project.get("numero", "") if project else ""
    ws["P9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    ws["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    ws["F12"] = perf_codigo
    ws["I12"] = muestra["numero"]
    ws["M12"] = to_float(muestra.get("profundidad_de"))
    ws["O12"] = to_float(muestra.get("profundidad_hasta"))
    ws["D13"] = descripcion_visual_para_excel(muestra) or ""
    ws["D14"] = observaciones_ensayo or ""

    diametro, altura = to_float(data.get("cons_diametro")), to_float(data.get("cons_altura"))
    ws["E19"] = diametro / 10 if diametro is not None else None
    ws["E20"] = altura / 10 if altura is not None else None
    anillo = to_float(data.get("cons_masa_anillo"))
    ws["E23"] = to_float(data.get("cons_masa_anillo_muestra_ini"))
    ws["F23"] = to_float(data.get("cons_masa_anillo_muestra_fin"))
    ws["E24"] = anillo
    ws["F24"] = _cons_final(data, "cons_fin_masa_anillo", "cons_masa_anillo")
    for col, pref in (("E", "ini"), ("F", "fin")):
        ws[f"{col}28"] = data.get(f"cons_{pref}_recipiente") or None
        ws[f"{col}29"] = to_float(data.get(f"cons_{pref}_humedo"))
        seco = next((v for v in (to_float(data.get(f"cons_{pref}_seco_{h}")) for h in (19, 18, 17)) if v is not None), None)
        ws[f"{col}30"] = seco
        ws[f"{col}31"] = to_float(data.get(f"cons_{pref}_recipiente_masa"))
    ws["K51"] = data.get("cons_consolidometro") or None
    ws["K52"] = data.get("cons_metodo") or None
    ws["K53"] = (data.get("cons_condicion") or "").upper() or None

    ws["AB18"] = to_float(data.get("cons_pic_no"))
    ws["AB19"] = to_float(data.get("cons_pic_masa_seco"))
    ws["AB21"] = to_float(data.get("cons_pic_masa_agua_suelo"))
    ws["AB22"] = to_float(data.get("cons_pic_temp"))
    ws["Z28"] = 0  # la bitácora no captura el % retenido en el tamiz No. 4
    for col in "BCDEFG":
        ws[f"{col}109"] = "=$AB$22"
    ws["Z30"] = "=100-Z28"
    ws["K30"] = "=AB25"

    # Lecturas de la máquina: la plantilla ya trae la rejilla de tiempos (min) de cada carga en "DATOS MAQUINA ";
    # se escribe la deformación de la máquina interpolada en cada tiempo, hasta donde haya datos.
    hoja_maq = wb["DATOS MAQUINA "]
    # En la plantilla la columna V (16 kg) leía la deformación pasando por la columna Y (32 kg), que a su vez copiaba
    # la U: para poder cargar el 32 kg por separado, V se calcula directo desde U (igual que sus filas de más abajo).
    for fila in range(4, hoja_maq.max_row + 1):
        v = hoja_maq[f"V{fila}"].value
        if isinstance(v, str) and v.startswith("=Y"):
            hoja_maq[f"V{fila}"] = f"=U{fila}+$AB$7"
    for i, (_enc, col_t, col_d) in enumerate(CONS_MAQ_BLOQUES, start=1):
        puntos = data.get(f"cons_maq_{i}")
        if not puntos:
            continue
        for fila in range(4, hoja_maq.max_row + 1):
            t_min = hoja_maq[f"{col_t}{fila}"].value
            if isinstance(t_min, (int, float)):
                hoja_maq[f"{col_d}{fila}"] = _cons_deformacion_en(puntos, t_min * 60)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return _restaurar_imagenes_perdidas(bio.getvalue(), TEMPLATE_CONSOLIDACION)

def _ci_promedio(data, prefijo):
    valores = [v for v in (to_float(data.get(f"{prefijo}_{i}")) for i in (1, 2, 3)) if v is not None]
    return sum(valores) / len(valores) if valores else None


def _ci_lecturas(data):
    """[(deformación en 0.001 in, carga en kN)] de las filas de la bitácora que tienen carga digitada."""
    lecturas = []
    for i, deformacion in enumerate(CI_DEFORMACIONES, start=1):
        carga = to_float(data.get(f"ci_carga_{i}"))
        if carga is not None:
            lecturas.append((deformacion, carga))
    return lecturas


def _ci_remuestrear(maquina, n=CI_MAX_MAQUINA):
    """La plantilla tiene n filas: si la máquina trae más, se toman n puntos igualmente espaciados en el tiempo,
    desde el cero hasta el final del ensayo (así quedan ~29 s entre filas en un ensayo de 15 min), interpolando
    la fuerza y la deformación entre las lecturas vecinas."""
    if len(maquina) <= n:
        return [list(x) for x in maquina]
    puntos = sorted(maquina, key=lambda x: x[0])
    t_fin = puntos[-1][0]
    res, j = [], 0
    for k in range(n):
        t = puntos[0][0] + (t_fin - puntos[0][0]) * k / (n - 1)
        while j < len(puntos) - 2 and puntos[j + 1][0] <= t:
            j += 1
        (t0, f0, d0), (t1, f1, d1) = puntos[j], puntos[j + 1]
        w = 0 if t1 == t0 else min(max((t - t0) / (t1 - t0), 0), 1)
        f = f0 if f1 is None else f1 if f0 is None else f0 + (f1 - f0) * w
        res.append([round(t, 1), None if f is None else round(f, 4), round(d0 + (d1 - d0) * w, 3)])
    return res


def _ci_puntos(data):
    """[(fila de la plantilla, tiempo en s o None, deformación mm, fuerza kN o None)]. Si se cargaron los datos de la
    máquina se usan esos (desde la fila 26, la del cero); si no, las lecturas de la bitácora (desde la fila 27; el
    tiempo solo se calcula si hay velocidad de falla)."""
    maquina = data.get("ci_maq")
    if maquina:
        return [(CI_FILA_MAQUINA + i, t, d, f) for i, (t, f, d) in enumerate(_ci_remuestrear(maquina))]
    velocidad = to_float(data.get("ci_velocidad"))
    puntos = []
    for fila, (deformacion, carga) in zip(CI_FILAS_EXCEL, _ci_lecturas(data)):
        mm = round(deformacion * 0.0254, 4)
        puntos.append((fila, round(mm / velocidad * 60, 1) if velocidad else None, mm, carga))
    return puntos


def parse_maquina_ci_xlsx(file_bytes):
    """Excel con la tabla de la máquina: se busca en cualquier hoja la fila de encabezados (Tiempo / Fuerza /
    Deformación o Desplazamiento) y se leen los datos que están debajo. Algunas máquinas traen una hoja "Informe"
    con fórmulas rotas (#N/A en todas las filas al abrir sin Excel) antes de la hoja real de datos — esa se salta
    sola porque no encuentra filas válidas. También pueden traer, además de la columna en mm, otra de "deformación
    unitaria (%)" que también empieza por "deform"; se prioriza la que trae "mm" en el encabezado, que es la que
    hace falta acá. Devuelve [[t, fuerza|None, deformación mm]]."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return []
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(min_row=1, max_row=3000, values_only=True))
        for r, fila in enumerate(filas):
            pos_t = pos_f = None
            candidatos_d = []  # (trae "mm" en el encabezado, columna)
            for c, v in enumerate(fila):
                txt = str(v).strip().lower() if v is not None else ""
                if pos_t is None and txt.startswith("tiempo"):
                    pos_t = c
                elif pos_f is None and txt.startswith("fuerza"):
                    pos_f = c
                elif txt.startswith("deform") or txt.startswith("desplazamiento"):
                    candidatos_d.append(("mm" in txt, c))
            if pos_t is None or pos_f is None or not candidatos_d:
                continue
            candidatos_d.sort(key=lambda x: not x[0])
            pos_d = candidatos_d[0][1]
            datos = []
            for fila_dato in filas[r + 1:]:
                if max(pos_t, pos_f, pos_d) >= len(fila_dato):
                    continue
                t, f, d = to_float(fila_dato[pos_t]), to_float(fila_dato[pos_f]), to_float(fila_dato[pos_d])
                if t is not None and d is not None:
                    datos.append([t, f, d])
            if datos:
                return datos
    return []


def resultados_compresion_inconfinada(data):
    """Mismas fórmulas de la plantilla GDA-FLC-008 (GUIA): humedad, densidades, esfuerzo con área corregida,
    qu, Su y módulo de elasticidad (entre las dos primeras lecturas, como la plantilla)."""
    filas = []
    d, h, masa = _ci_promedio(data, "ci_d"), _ci_promedio(data, "ci_h"), to_float(data.get("ci_peso"))
    humedo, rec = to_float(data.get("ci_hum_humedo")), to_float(data.get("ci_hum_masa_rec"))
    seco = next((v for v in (to_float(data.get(f"ci_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    w = (humedo - seco) / (seco - rec) * 100 if None not in (humedo, seco, rec) and (seco - rec) != 0 else None
    if w is not None:
        filas.append(("Humedad (%)", fmt_num(w, 2)))
    if not d or not h:
        return filas
    area = math.pi * d ** 2 / 4
    vol = area * h
    filas += [("Diámetro promedio (cm)", fmt_num(d, 2)), ("Altura promedio (cm)", fmt_num(h, 2)),
              ("Área (cm²)", fmt_num(area, 2)), ("Volumen (cm³)", fmt_num(vol, 2))]
    if masa:
        rho_h = masa / vol
        filas.append(("Densidad húmeda ρh (g/cm³)", fmt_num(rho_h, 3)))
        if w is not None:
            filas.append(("Densidad seca ρd (g/cm³)", fmt_num(rho_h / (1 + w / 100), 3)))
    puntos = {}  # fila de la plantilla -> (deformación unitaria %, esfuerzo kPa)
    for fila, _t, mm, carga in _ci_puntos(data):
        eps = mm / (h * 10) * 100
        if eps >= 100:
            continue
        if fila == CI_FILA_MAQUINA:
            puntos[fila] = (eps, 0.0)  # la fila 26 de la plantilla trae el esfuerzo en 0
        elif carga is not None:
            puntos[fila] = (eps, carga / (area / (1 - eps / 100) / 10000))
    if puntos:
        qu = max(s for _e, s in puntos.values())
        filas += [("Resistencia a la compresión inconfinada qu (kPa)", fmt_num(qu, 1)),
                  ("Resistencia al corte Su (kPa)", fmt_num(qu / 2, 1))]
        if 27 in puntos and 28 in puntos and puntos[28][0] != puntos[27][0]:
            filas.append(("Módulo de elasticidad (kPa)", fmt_num((puntos[28][1] - puntos[27][1]) / (puntos[28][0] - puntos[27][0]) * 100, 0)))
    return filas


def _procesar_foto(imagen_bytes, ancho_max=1280, peso_max_kb=300):
    """Comprime la foto del plano de falla antes de guardarla en el ensayo (jsonb): se reduce a un ancho máximo y se
    guarda como JPEG. La foto de una cámara de celular pesa varios MB (a veces HEIC/RAW) — subirla y guardarla así de
    pesada es lo que dejaba la app "cargando" sin terminar en redes lentas y, si la escritura a Supabase fallaba a
    medio camino, la foto nunca quedaba guardada de verdad. Se baja la calidad hasta quedar bajo `peso_max_kb`
    (o hasta una calidad mínima razonable, para no dejar una foto irreconocible). Devuelve {"b64", "ext", "ancho", "alto"}."""
    img = Image.open(BytesIO(imagen_bytes))
    img = ImageOps.exif_transpose(img).convert("RGB")
    if img.width > ancho_max:
        img = img.resize((ancho_max, round(img.height * ancho_max / img.width)))
    for calidad in (70, 55, 40, 30):
        bio = BytesIO()
        img.save(bio, format="JPEG", quality=calidad)
        if bio.tell() <= peso_max_kb * 1024 or calidad == 30:
            break
    return {"b64": base64.b64encode(bio.getvalue()).decode("ascii"), "ext": "jpeg", "ancho": img.width, "alto": img.height}


def _guardar_inmediato(assay_id, data):
    """Guarda `data` en Supabase ya mismo y sincroniza la copia en memoria del ensayo (`st.session_state.assays`) —
    se usa antes de un `st.rerun()` que hace falta disparar EN EL ACTO (para que la interfaz cambie de una, como al
    subir una foto o quitar un dato importado), en vez de esperar al autoguardado normal, que corre más abajo en el
    flujo del formulario y por eso nunca llega a ejecutarse si el rerun corta el guion antes de esa línea.
    Sin este guardado explícito, el cambio se perdía apenas se recargaba la página: la corrida siguiente volvía a
    leer `data` desde la copia vieja en memoria, así que con un selector de archivo (que sigue "lleno" en cada
    corrida, a diferencia de un botón) la foto se procesaba y se intentaba guardar una y otra vez sin parar nunca —
    el "muñeco" de arriba quedaba corriendo para siempre y la foto jamás llegaba a guardarse de verdad. Devuelve
    True si guardó bien."""
    try:
        db.update_assay_data(assay_id, data=data)
    except Exception:
        st.error("No se pudo guardar (revisa tu conexión) — vuelve a intentarlo.")
        return False
    assay_ref = next((a for a in st.session_state.assays if a["id"] == assay_id), None)
    if assay_ref is not None:
        assay_ref["data"] = data
    return True


def render_compresion_inconfinada_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora GDA-FL-005. El Excel para descargar (plantilla oficial GDA-FLC-008) "
            "está al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    def _radio(key, label, opciones):
        actual = data.get(key, opciones[0])
        data[key] = st.radio(label, opciones, horizontal=True, index=opciones.index(actual) if actual in opciones else 0,
                              key=f"{key}_{assay_id}")

    with st.container(border=True):
        st.markdown(card_header_html("science", "Muestra"), unsafe_allow_html=True)
        _radio("ci_condicion", "Tipo de muestra", CI_CONDICIONES)
        if data.get("ci_condicion") == "Inalterada":
            _radio("ci_muestreo", "Método de muestreo", CI_MUESTREOS)
        for key, label in (("ci_temp_ini", "Temperatura inicial (°C)"), ("ci_temp_fin", "Temperatura final (°C)")):
            _campo(key, label, placeholder="0.0")
    with st.container(border=True):
        st.markdown(card_header_html("straighten", "Dimensiones de la Muestra"), unsafe_allow_html=True)
        head = st.columns([1, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Altura (cm)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Diámetro (cm)</div>', unsafe_allow_html=True)
        for i in (1, 2, 3):
            row = st.columns([1, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{i}</div>', unsafe_allow_html=True)
            for col_i, pref in ((1, "ci_h"), (2, "ci_d")):
                key = f"{pref}_{i}"
                data[key] = row[col_i].text_input(f"{pref} {i}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed", placeholder="0.00")
        _campo("ci_peso", "Peso de la muestra (g)")
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
        for key, label in CI_HUMEDAD_FILAS:
            _campo(key, label, placeholder="" if key == "ci_hum_recipiente" else "0.00")
            if key == "ci_hum_seco_17" and data.get("ci_hum_seco_17"):
                # Normalmente el peso ya se estabilizó a las 17 horas y se repite igual a las 18 y 19 — se copia
                # solo, sin pisar un valor distinto que ya se hubiera digitado a mano.
                for siguiente in ("ci_hum_seco_18", "ci_hum_seco_19"):
                    if not data.get(siguiente):
                        st.session_state[f"{siguiente}_{assay_id}"] = data["ci_hum_seco_17"]
                        data[siguiente] = data["ci_hum_seco_17"]
        _radio("ci_temp_secado", "Temperatura de secado", ["60 °C", "110 °C"])
        _radio("ci_metodo", "Método", ["A", "B"])
        _radio("ci_hum_antes", "Humedad obtenida", ["Antes del ensayo", "Después del ensayo"])
        _radio("ci_hum_muestra", "Sobre", ["Muestra completa", "Cortes de muestra"])
    with st.container(border=True):
        st.markdown(card_header_html("show_chart", "Datos de la Máquina"), unsafe_allow_html=True)
        st.caption("Tiempo (s), fuerza (kN) y deformación (mm) que arroja la máquina: van a la tabla de datos de la máquina del "
                   f"Excel ({CI_MAX_MAQUINA} filas; si traen más, se reparten en el tiempo). Si los cargas, se usan en lugar de la tabla de la bitácora.")
        archivo = st.file_uploader("Sube el Excel de la máquina", type=["xlsx"], key=f"ci_maq_archivo_{assay_id}")
        if archivo and st.button("Cargar datos de la máquina", key=f"ci_maq_cargar_{assay_id}",
                                  icon=":material/publish:", use_container_width=True):
            filas_maq = parse_maquina_ci_xlsx(archivo.getvalue())
            if not filas_maq:
                st.error("No encontré filas con tiempo, fuerza y deformación. Revisa que estén las 3 columnas.")
            else:
                data["ci_maq"] = filas_maq
                # Se guarda ya mismo (no se espera al autoguardado normal, que corre después en el formulario): sin
                # esto, si el laboratorista recargaba la página o volvía a entrar antes de descargar, lo cargado
                # se perdía y el Excel salía sin los datos de la máquina, aunque acá arriba mostrara "cargado".
                if _guardar_inmediato(assay_id, data):
                    st.success(f"Se cargaron y guardaron {len(filas_maq)} filas.")
                    if len(filas_maq) > CI_MAX_MAQUINA:
                        st.info(f"La plantilla admite {CI_MAX_MAQUINA} filas: se exportan {CI_MAX_MAQUINA} puntos igualmente "
                                "espaciados en el tiempo, desde el cero hasta el final del ensayo.")
                else:
                    data.pop("ci_maq", None)
        if data.get("ci_maq"):
            st.markdown(f'<div class="cell-muted">Cargado: {len(data["ci_maq"])} filas, hasta {data["ci_maq"][-1][0]:.0f} s</div>',
                        unsafe_allow_html=True)
            if st.button("Quitar datos de la máquina", key=f"ci_maq_quitar_{assay_id}"):
                data.pop("ci_maq", None)
                if _guardar_inmediato(assay_id, data):
                    st.rerun()
    with st.container(border=True):
        st.markdown(card_header_html("tune", "Falla"), unsafe_allow_html=True)
        _campo("ci_penetrometro", "Resistencia al penetrómetro (kg/cm²)")
        st.markdown('<div style="padding-top:4px;">Tipo de falla (diagrama)</div>', unsafe_allow_html=True)
        actual = data.get("ci_falla", CI_FALLAS[0])
        data["ci_falla"] = st.selectbox("Tipo de falla (diagrama)", CI_FALLAS,
                                         index=CI_FALLAS.index(actual) if actual in CI_FALLAS else 0,
                                         key=f"ci_falla_{assay_id}", label_visibility="collapsed")
        _campo("ci_velocidad", "Velocidad de falla (mm/min)", placeholder="1")
        _campo("ci_tiempo_falla", "Tiempo de falla (min)")
    with st.container(border=True):
        st.markdown(card_header_html("photo_camera", "Foto del Plano de Falla"), unsafe_allow_html=True)
        st.caption("Se agrega al Excel dentro del recuadro \"PLANO DE FALLA\", centrada y con margen (queda encima del "
                   "dibujo de la plantilla): la puedes mover o redimensionar a mano después de descargar.")
        if data.get("ci_foto_falla"):
            st.image(base64.b64decode(data["ci_foto_falla"]["b64"]), width=220)
            if st.button("Quitar foto", key=f"ci_foto_quitar_{assay_id}"):
                data.pop("ci_foto_falla", None)
                # Se cambia el número de intento para que el selector de archivo vuelva a nacer limpio: con la misma
                # key, Streamlit sigue acordándose del archivo ya subido aunque el widget estuviera oculto (mientras
                # se mostraba la foto), así que "Quitar" no quitaba nada de verdad — la misma foto volvía a aparecer.
                data["_ci_foto_intento"] = data.get("_ci_foto_intento", 0) + 1
                if _guardar_inmediato(assay_id, data):
                    st.rerun()
        else:
            # En el celular, este selector de archivo ya abre la cámara del sistema (y ahí el laboratorista escoge
            # trasera o delantera) además de la galería — no hace falta el widget de cámara en vivo de Streamlit.
            captura = st.file_uploader("Foto del plano de falla", type=["png", "jpg", "jpeg"],
                                       key=f"ci_foto_archivo_{assay_id}_{data.get('_ci_foto_intento', 0)}",
                                       label_visibility="collapsed")
            if captura is not None:
                with st.spinner("Guardando la foto…"):
                    data["ci_foto_falla"] = _procesar_foto(captura.getvalue())
                    if _guardar_inmediato(assay_id, data):
                        st.rerun()
                    else:
                        data.pop("ci_foto_falla", None)
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_compresion_inconfinada(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digitan los datos de arriba.")
    render_equipo(data, "ci", EQUIPO_COMPRESION_INCONFINADA)
    render_norma_selector("compresion-inconfinada", data, "ci")


def _col_a_numero(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


class _FormulaXlsx(str):
    """Texto de una fórmula de Excel (sin el "=" inicial) para _xlsx_escribir_celdas: se escribe como <f>, no como texto."""


def _xlsx_escribir_celdas(xlsx_bytes, hoja_xml, celdas):
    """Escribe valores en una hoja editando su XML directamente (sin pasar por openpyxl), así se conserva TODO lo demás del
    archivo tal cual: formas, grupos, casillas de verificación, imágenes, gráficos, macros. `celdas` = {"C6": valor}; un
    número queda como número, un texto como texto y None deja la celda vacía. Conserva el estilo que ya tenga la celda."""
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zin:
        xml = zin.read(hoja_xml).decode("utf-8")

        def celda_xml(ref, estilo, valor):
            attr_s = f' s="{estilo}"' if estilo else ""
            if valor is None or valor == "":
                return f'<c r="{ref}"{attr_s}/>'
            if isinstance(valor, _FormulaXlsx):
                return f'<c r="{ref}"{attr_s}><f>{html.escape(str(valor), quote=False)}</f></c>'
            if isinstance(valor, bool) or not isinstance(valor, (int, float)):
                texto = html.escape(str(valor), quote=False)
                return f'<c r="{ref}"{attr_s} t="inlineStr"><is><t xml:space="preserve">{texto}</t></is></c>'
            return f'<c r="{ref}"{attr_s}><v>{repr(float(valor)) if isinstance(valor, float) else valor}</v></c>'

        for ref, valor in celdas.items():
            m_ref = re.match(r"([A-Z]+)(\d+)$", ref)
            col, fila = m_ref.group(1), int(m_ref.group(2))
            patron = re.compile(r'<c r="' + ref + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
            m = patron.search(xml)
            if m:
                est = re.search(r'\bs="(\d+)"', m.group(1))
                xml = xml[:m.start()] + celda_xml(ref, est.group(1) if est else "", valor) + xml[m.end():]
                continue
            # La celda no existe: se inserta en su fila respetando el orden de columnas.
            m_fila = re.search(r'<row r="' + str(fila) + r'"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
            if not m_fila or m_fila.group(1) is None:
                continue
            cuerpo, inicio = m_fila.group(1), m_fila.start(1)
            pos = len(cuerpo)
            for mc in re.finditer(r'<c r="([A-Z]+)\d+"', cuerpo):
                if _col_a_numero(mc.group(1)) > _col_a_numero(col):
                    pos = mc.start()
                    break
            xml = xml[:inicio + pos] + celda_xml(ref, "", valor) + xml[inicio + pos:]

        # xl/calcChain.xml es un caché de qué celdas tienen fórmula y en qué orden se recalculan — cuando esta
        # función le quita la fórmula a una celda (por ejemplo, para reemplazarla por un valor ya calculado en
        # Python, como en carga puntual) esa caché queda apuntando a una celda que ya no tiene fórmula. openpyxl
        # abre el archivo igual, pero Excel lo rechaza directo ("no se puede abrir") sin decir por qué. Se quita
        # el archivo entero (es solo una optimización, no hace falta) y sus referencias — con fullCalcOnLoad
        # activado más abajo, Excel recalcula todo de cero al abrir sin necesitarlo.
        rels_wb = "xl/_rels/workbook.xml.rels"
        tiene_calc_chain = "xl/calcChain.xml" in zin.namelist()

        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue
                contenido = zin.read(item.filename)
                if item.filename == hoja_xml:
                    contenido = xml.encode("utf-8")
                elif item.filename == "xl/workbook.xml":  # que Excel recalcule las fórmulas al abrir
                    texto = contenido.decode("utf-8")
                    if "fullCalcOnLoad" not in texto:
                        texto = texto.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
                    contenido = texto.encode("utf-8")
                elif item.filename == "[Content_Types].xml" and tiene_calc_chain:
                    contenido = re.sub(rb'<Override PartName="/xl/calcChain\.xml"[^>]*/>', b"", contenido)
                elif item.filename == rels_wb and tiene_calc_chain:
                    contenido = re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b"", contenido)
                zout.writestr(item, contenido)
        return bio.getvalue()


def _insertar_imagen_hoja(xlsx_bytes, drawing_xml, imagen_bytes, imagen_ext, col, fila, ancho_emu, alto_emu,
                          col_off=0, fila_off=0):
    """Agrega una imagen suelta (sin encogerla a ninguna celda) al drawing ya existente de una hoja — se usa para
    fotos que el laboratorista sube, que quedan junto a (o dentro de) un recuadro de la plantilla para que las
    acomode a mano tras descargar. `col`/`fila` son 0-indexados; `col_off`/`fila_off` son un desplazamiento en EMU
    desde el borde de esa celda (914400 EMU = 1 pulgada), para meter la imagen unos milímetros adentro de un
    recuadro en vez de pegarla justo al borde. `drawing_xml` es la ruta del drawingN.xml de esa hoja."""
    rels_xml = re.sub(r"([^/]+)\.xml$", r"_rels/\1.xml.rels", drawing_xml)
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zin:
        nombres = set(zin.namelist())
        existentes = [int(m.group(1)) for n in nombres for m in [re.match(r"xl/media/image(\d+)\.\w+$", n)] if m]
        media_nombre = f"xl/media/image{(max(existentes) + 1) if existentes else 1}.{imagen_ext}"

        rels = zin.read(rels_xml).decode("utf-8") if rels_xml in nombres else \
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        ids = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels)]
        rid = f"rId{(max(ids) + 1) if ids else 1}"
        rels = rels.replace("</Relationships>",
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="../media/{media_nombre.rsplit("/", 1)[1]}"/></Relationships>')

        drawing = zin.read(drawing_xml).decode("utf-8")
        pic_id = len(re.findall(r"<xdr:cNvPr ", drawing)) + 1
        ancla = (
            f'<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>'
            f'<xdr:row>{fila}</xdr:row><xdr:rowOff>{fila_off}</xdr:rowOff></xdr:from>'
            f'<xdr:ext cx="{ancho_emu}" cy="{alto_emu}"/>'
            f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{pic_id}" name="Foto plano de falla"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>'
            f'<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{ancho_emu}" cy="{alto_emu}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
        )
        drawing = drawing.replace("</xdr:wsDr>", ancla + "</xdr:wsDr>")

        content_types = zin.read("[Content_Types].xml").decode("utf-8")
        if f'Extension="{imagen_ext}"' not in content_types:
            mime = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg"}.get(imagen_ext, "image/png")
            content_types = content_types.replace(
                "</Types>", f'<Default Extension="{imagen_ext}" ContentType="{mime}"/></Types>')

        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == drawing_xml:
                    zout.writestr(item, drawing.encode("utf-8"))
                elif item.filename == rels_xml:
                    zout.writestr(item, rels.encode("utf-8"))
                elif item.filename == "[Content_Types].xml":
                    zout.writestr(item, content_types.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
            if rels_xml not in nombres:
                zout.writestr(rels_xml, rels.encode("utf-8"))
            zout.writestr(media_nombre, imagen_bytes)
        return bio.getvalue()


# Recuadro "PLANO DE FALLA" de las plantillas de Compresión Inconfinada y en Roca (el mismo, ambas comparten esa
# parte de la plantilla): ancla (columna/fila 0-indexadas + desplazamiento en EMU) y tamaño del recuadro completo,
# leídos del propio grupo de dibujo de la plantilla ("Grupo 31").
PLANO_FALLA_ANCLA = (11, 15, 62843, 92610)
PLANO_FALLA_CX, PLANO_FALLA_CY = 2351187, 3789642


def _foto_dentro_de_recuadro(foto, margen_emu=180000):
    """Calcula col/fila/desplazamiento/tamaño para que la foto quede DENTRO del recuadro "PLANO DE FALLA", con
    margen parejo a los cuatro lados y sin deformarla (se encoge al lado que le sobre para no salirse del
    recuadro). Devuelve los kwargs listos para _insertar_imagen_hoja."""
    col, fila, col_off, fila_off = PLANO_FALLA_ANCLA
    interior_cx, interior_cy = PLANO_FALLA_CX - 2 * margen_emu, PLANO_FALLA_CY - 2 * margen_emu
    aspecto = foto["alto"] / foto["ancho"]
    if interior_cx * aspecto <= interior_cy:
        ancho_emu, alto_emu = interior_cx, round(interior_cx * aspecto)
    else:
        alto_emu, ancho_emu = interior_cy, round(interior_cy / aspecto)
    return {
        "col": col, "fila": fila, "ancho_emu": ancho_emu, "alto_emu": alto_emu,
        "col_off": col_off + margen_emu + (interior_cx - ancho_emu) // 2,
        "fila_off": fila_off + margen_emu + (interior_cy - alto_emu) // 2,
    }


def _textos_compartidos(xlsx_path):
    with zipfile.ZipFile(xlsx_path) as z:
        xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    textos = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        textos.append(html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    return textos


def generar_excel_compresion_inconfinada(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Compresión inconfinada (GDA-FLC-008, INV E-152). Se llena la hoja GUIA: encabezado, dimensiones, masa, humedad y las
    lecturas de carga/deformación (columnas P, R y U, que en la plantilla reciben los datos de la máquina); esfuerzo, qu y
    Su los calcula el Excel. La plantilla trae formas, grupos (plano de falla, firmas), casillas y un gráfico que openpyxl
    no conserva, por eso los valores se escriben directo en el XML de la hoja."""
    c = {}
    c["C6"] = project.get("cliente", "") if project else ""
    c["C7"] = project["nombre"] if project else codigo
    c["C8"] = project.get("correo_cliente", "") if project else ""
    c["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["C10"] = project["muestra_tomada_por"]
    c["K6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["K7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["K8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["L9"] = project.get("numero", "") if project else ""
    c["N9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["D12"] = perf_codigo
    c["F12"] = muestra["numero"]
    c["H12"] = to_float(muestra.get("profundidad_de"))
    c["J12"] = to_float(muestra.get("profundidad_hasta"))
    c["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    c["C18"] = _ci_promedio(data, "ci_d")
    c["C19"] = _ci_promedio(data, "ci_h")
    c["C21"] = to_float(data.get("ci_peso"))
    c["I18"] = to_float(data.get("ci_hum_humedo"))
    c["I19"] = next((v for v in (to_float(data.get(f"ci_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    c["I20"] = to_float(data.get("ci_hum_masa_rec"))

    # Las listas desplegables de la plantilla traen espacios al final de algunas opciones: se usa el texto exacto.
    textos = {t.strip(): t for t in _textos_compartidos(TEMPLATE_COMPRESION_INCONFINADA)}
    c["F20"] = textos.get(data.get("ci_falla", ""), data.get("ci_falla") or None)
    condicion = data.get("ci_condicion", "Inalterada")
    muestreo = {"Compactada": "Compactada", "Remodelada": "Remoldeada"}.get(condicion, data.get("ci_muestreo") or "Tubo Shelby")
    c["F21"] = textos.get(muestreo, muestreo)
    c["F22"] = to_float(data.get("ci_velocidad"))

    for fila, t, mm, fuerza in _ci_puntos(data):
        c[f"P{fila}"] = t
        c[f"R{fila}"] = fuerza
        c[f"U{fila}"] = mm

    with open(TEMPLATE_COMPRESION_INCONFINADA, "rb") as f:
        plantilla = f.read()
    salida = _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet1.xml", c),
                                                   TEMPLATE_COMPRESION_INCONFINADA)
    foto = data.get("ci_foto_falla")
    if foto:
        imagen = base64.b64decode(foto["b64"])
        # Se mete DENTRO del recuadro "PLANO DE FALLA" de la plantilla, con margen — no reemplaza el dibujo del
        # cilindro (queda encima) ni se ajusta al recuadro exacto: se centra y se encoge lo justo para no salirse.
        salida = _insertar_imagen_hoja(salida, "xl/drawings/drawing1.xml", imagen, foto["ext"],
                                       **_foto_dentro_de_recuadro(foto))
    return salida

# Compresión simple en roca (ASTM D7012 método B) — bitácora GDA-FL-007 y plantilla GDA-FLC-043. La plantilla de
# descarga es prácticamente igual a la de Compresión inconfinada (mismas columnas de la tabla de la máquina, mismos
# datos de la muestra y de humedad) salvo que no reporta resistencia al corte Su (no aplica a roca) y tiene su
# propia lista de tipos de falla.
ROCA_DEFORMACIONES = list(range(5, 166, 5))  # 0.001 in — bitácora GDA-FL-007, filas E19:E52 (34 filas, una en blanco)
ROCA_FILA_MAQUINA, ROCA_MAX_MAQUINA = 26, 33  # igual que compresión inconfinada: la plantilla trae 33 filas (la del cero incluida)
ROCA_FILAS_EXCEL = list(range(27, 59))
ROCA_FALLAS = ["FALLA CONTROLADA POR DISCONTINUIDADES", "FALLA POR DIVISIÓN AXIAL", "FALLA POR CORTE", "FALLA CÓNICA",
               "FALLA COLUMNAR O MÚLTIPLE", "FALLA ESCALONADA"]
ROCA_HUMEDAD_FILAS = [
    ("roca_hum_recipiente", "Recipiente No."), ("roca_hum_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("roca_hum_seco_17", "Masa suelo seco + recipiente (g) (17 horas)"), ("roca_hum_seco_18", "Masa suelo seco + recipiente (g) (18 horas)"),
    ("roca_hum_seco_19", "Masa suelo seco + recipiente (g) (19 horas)"), ("roca_hum_masa_rec", "Masa del recipiente (g)"),
]
EQUIPO_COMPRESION_ROCA = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Máquina de compresión GDA-E-008",
                          "Máquina de compresión GDA-E-014", "Pie de rey GDA-E-110", "Horno GDA-E-007", "Horno GDA-E-404"]


def _roca_promedio(data, prefijo):
    valores = [v for v in (to_float(data.get(f"{prefijo}_{i}")) for i in (1, 2, 3)) if v is not None]
    return sum(valores) / len(valores) if valores else None


def _roca_lecturas(data):
    """[(deformación en 0.001 in, carga en kN)] de las filas de la bitácora que tienen carga digitada."""
    lecturas = []
    for i, deformacion in enumerate(ROCA_DEFORMACIONES, start=1):
        carga = to_float(data.get(f"roca_carga_{i}"))
        if carga is not None:
            lecturas.append((deformacion, carga))
    return lecturas


def _roca_puntos(data):
    """Igual que _ci_puntos: usa los datos de la máquina si se cargaron, si no las lecturas de la bitácora."""
    maquina = data.get("roca_maq")
    if maquina:
        return [(ROCA_FILA_MAQUINA + i, t, d, f) for i, (t, f, d) in enumerate(_ci_remuestrear(maquina, ROCA_MAX_MAQUINA))]
    velocidad = to_float(data.get("roca_velocidad"))
    puntos = []
    for fila, (deformacion, carga) in zip(ROCA_FILAS_EXCEL, _roca_lecturas(data)):
        mm = round(deformacion * 0.0254, 4)
        puntos.append((fila, round(mm / velocidad * 60, 1) if velocidad else None, mm, carga))
    return puntos


def resultados_compresion_roca(data):
    """Mismas fórmulas de la plantilla GDA-FLC-043 (GUIA): humedad, densidades, esfuerzo con área corregida y qu
    (no reporta Su ni módulo de elasticidad aparte de qu, igual que en la plantilla de roca)."""
    filas = []
    d, h, masa = _roca_promedio(data, "roca_d"), _roca_promedio(data, "roca_h"), to_float(data.get("roca_peso"))
    humedo, rec = to_float(data.get("roca_hum_humedo")), to_float(data.get("roca_hum_masa_rec"))
    seco = next((v for v in (to_float(data.get(f"roca_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    w = (humedo - seco) / (seco - rec) * 100 if None not in (humedo, seco, rec) and (seco - rec) != 0 else None
    if w is not None:
        filas.append(("Humedad (%)", fmt_num(w, 2)))
    if not d or not h:
        return filas
    area = math.pi * d ** 2 / 4
    vol = area * h
    filas += [("Diámetro promedio (cm)", fmt_num(d, 2)), ("Altura promedio (cm)", fmt_num(h, 2)),
              ("Área (cm²)", fmt_num(area, 2)), ("Volumen (cm³)", fmt_num(vol, 2))]
    if masa:
        rho_h = masa / vol
        filas.append(("Densidad húmeda ρh (g/cm³)", fmt_num(rho_h, 3)))
        if w is not None:
            filas.append(("Densidad seca ρd (g/cm³)", fmt_num(rho_h / (1 + w / 100), 3)))
    puntos = {}
    for fila, _t, mm, carga in _roca_puntos(data):
        eps = mm / (h * 10) * 100
        if eps >= 100:
            continue
        if fila == ROCA_FILA_MAQUINA and data.get("roca_maq"):
            puntos[fila] = (eps, 0.0)
        elif carga is not None:
            puntos[fila] = (eps, carga / (area / (1 - eps / 100) / 10000))
    if puntos:
        qu = max(s for _e, s in puntos.values())
        filas.append(("Resistencia a la compresión qu (kPa)", fmt_num(qu, 1)))
        if 27 in puntos and 28 in puntos and puntos[28][0] != puntos[27][0]:
            filas.append(("Módulo de elasticidad (kPa)", fmt_num((puntos[28][1] - puntos[27][1]) / (puntos[28][0] - puntos[27][0]) * 100, 0)))
    return filas


def render_compresion_roca_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora GDA-FL-007. El Excel para descargar (plantilla oficial GDA-FLC-043) "
            "está al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    def _radio(key, label, opciones):
        actual = data.get(key, opciones[0])
        data[key] = st.radio(label, opciones, horizontal=True, index=opciones.index(actual) if actual in opciones else 0,
                              key=f"{key}_{assay_id}")

    with st.container(border=True):
        st.markdown(card_header_html("science", "Muestra"), unsafe_allow_html=True)
        for key, label in (("roca_temp_ini", "Temperatura inicial (°C)"), ("roca_temp_fin", "Temperatura final (°C)")):
            _campo(key, label, placeholder="0.0")
    with st.container(border=True):
        st.markdown(card_header_html("straighten", "Dimensiones de la Muestra"), unsafe_allow_html=True)
        head = st.columns([1, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Altura (cm)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Diámetro (cm)</div>', unsafe_allow_html=True)
        for i in (1, 2, 3):
            row = st.columns([1, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{i}</div>', unsafe_allow_html=True)
            for col_i, pref in ((1, "roca_h"), (2, "roca_d")):
                key = f"{pref}_{i}"
                data[key] = row[col_i].text_input(f"{pref} {i}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed", placeholder="0.00")
        _campo("roca_peso", "Peso de la muestra (g)")
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
        for key, label in ROCA_HUMEDAD_FILAS:
            _campo(key, label, placeholder="" if key == "roca_hum_recipiente" else "0.00")
            if key == "roca_hum_seco_17" and data.get("roca_hum_seco_17"):
                for siguiente in ("roca_hum_seco_18", "roca_hum_seco_19"):
                    if not data.get(siguiente):
                        st.session_state[f"{siguiente}_{assay_id}"] = data["roca_hum_seco_17"]
                        data[siguiente] = data["roca_hum_seco_17"]
        _radio("roca_temp_secado", "Temperatura de secado", ["60 °C", "110 °C"])
        _radio("roca_metodo", "Método", ["A", "B"])
        _radio("roca_hum_antes", "Humedad obtenida", ["Antes del ensayo", "Después del ensayo"])
        _radio("roca_hum_muestra", "Sobre", ["Cortes de muestra", "Muestra completa"])
    with st.container(border=True):
        st.markdown(card_header_html("show_chart", "Datos de la Máquina"), unsafe_allow_html=True)
        st.caption("Tiempo (s), fuerza (kN) y deformación (mm) que arroja la máquina: van a la tabla de datos de la máquina del "
                   f"Excel ({ROCA_MAX_MAQUINA} filas; si traen más, se reparten en el tiempo). Si los cargas, se usan en lugar de "
                   "la tabla de la bitácora.")
        archivo = st.file_uploader("Sube el Excel de la máquina", type=["xlsx"], key=f"roca_maq_archivo_{assay_id}")
        if archivo and st.button("Cargar datos de la máquina", key=f"roca_maq_cargar_{assay_id}",
                                  icon=":material/publish:", use_container_width=True):
            filas_maq = parse_maquina_ci_xlsx(archivo.getvalue())
            if not filas_maq:
                st.error("No encontré filas con tiempo, fuerza y deformación. Revisa que estén las 3 columnas.")
            else:
                data["roca_maq"] = filas_maq
                if _guardar_inmediato(assay_id, data):
                    st.success(f"Se cargaron y guardaron {len(filas_maq)} filas.")
                    if len(filas_maq) > ROCA_MAX_MAQUINA:
                        st.info(f"La plantilla admite {ROCA_MAX_MAQUINA} filas: se exportan {ROCA_MAX_MAQUINA} puntos igualmente "
                                "espaciados en el tiempo, desde el cero hasta el final del ensayo.")
                else:
                    data.pop("roca_maq", None)
        if data.get("roca_maq"):
            st.markdown(f'<div class="cell-muted">Cargado: {len(data["roca_maq"])} filas, hasta {data["roca_maq"][-1][0]:.0f} s</div>',
                        unsafe_allow_html=True)
            if st.button("Quitar datos de la máquina", key=f"roca_maq_quitar_{assay_id}"):
                data.pop("roca_maq", None)
                if _guardar_inmediato(assay_id, data):
                    st.rerun()
    with st.container(border=True):
        st.markdown(card_header_html("tune", "Falla"), unsafe_allow_html=True)
        st.markdown('<div style="padding-top:4px;">Tipo de falla (diagrama)</div>', unsafe_allow_html=True)
        actual = data.get("roca_falla", ROCA_FALLAS[0])
        data["roca_falla"] = st.selectbox("Tipo de falla (diagrama)", ROCA_FALLAS,
                                           index=ROCA_FALLAS.index(actual) if actual in ROCA_FALLAS else 0,
                                           key=f"roca_falla_{assay_id}", label_visibility="collapsed")
        _campo("roca_velocidad", "Velocidad de falla (mm/min)", placeholder="1")
        _campo("roca_tiempo_falla", "Tiempo de falla (min)")
        _campo("roca_esfuerzo_maximo_manual", "Esfuerzo máximo leído en la máquina (MPa, opcional)", placeholder="0.0")
    with st.container(border=True):
        st.markdown(card_header_html("photo_camera", "Foto del Plano de Falla"), unsafe_allow_html=True)
        st.caption("Se agrega al Excel dentro del recuadro \"PLANO DE FALLA\", centrada y con margen (queda encima del "
                   "dibujo de la plantilla): la puedes mover o redimensionar a mano después de descargar.")
        if data.get("roca_foto_falla"):
            st.image(base64.b64decode(data["roca_foto_falla"]["b64"]), width=220)
            if st.button("Quitar foto", key=f"roca_foto_quitar_{assay_id}"):
                data.pop("roca_foto_falla", None)
                data["_roca_foto_intento"] = data.get("_roca_foto_intento", 0) + 1
                if _guardar_inmediato(assay_id, data):
                    st.rerun()
        else:
            captura = st.file_uploader("Foto del plano de falla", type=["png", "jpg", "jpeg"],
                                       key=f"roca_foto_archivo_{assay_id}_{data.get('_roca_foto_intento', 0)}",
                                       label_visibility="collapsed")
            if captura is not None:
                with st.spinner("Guardando la foto…"):
                    data["roca_foto_falla"] = _procesar_foto(captura.getvalue())
                    if _guardar_inmediato(assay_id, data):
                        st.rerun()
                    else:
                        data.pop("roca_foto_falla", None)
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_compresion_roca(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digitan los datos de arriba.")
    render_equipo(data, "roca", EQUIPO_COMPRESION_ROCA)
    render_norma_selector("compresion-roca", data, "roca")


def generar_excel_compresion_roca(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Compresión simple en roca (GDA-FLC-043, ASTM D7012 método B). Se escribe directo en el XML de la hoja (igual
    que compresión inconfinada) para conservar el plano de falla, las firmas, las casillas y el gráfico de la
    plantilla — pasar por openpyxl los perdía."""
    c = {}
    c["C6"] = project.get("cliente", "") if project else ""
    c["C7"] = project["nombre"] if project else codigo
    c["C8"] = project.get("correo_cliente", "") if project else ""
    c["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["C10"] = project["muestra_tomada_por"]
    c["K6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["K7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["K8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["L9"] = project.get("numero", "") if project else ""
    c["N9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["D12"] = perf_codigo
    c["F12"] = muestra["numero"]
    c["H12"] = to_float(muestra.get("profundidad_de"))
    c["J12"] = to_float(muestra.get("profundidad_hasta"))
    c["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    c["C18"] = _roca_promedio(data, "roca_d")
    c["C19"] = _roca_promedio(data, "roca_h")
    c["C21"] = to_float(data.get("roca_peso"))
    c["I18"] = to_float(data.get("roca_hum_humedo"))
    c["I19"] = next((v for v in (to_float(data.get(f"roca_hum_seco_{x}")) for x in (19, 18, 17)) if v is not None), None)
    c["I20"] = to_float(data.get("roca_hum_masa_rec"))

    textos = {t.strip(): t for t in _textos_compartidos(TEMPLATE_COMPRESION_ROCA)}
    c["F20"] = textos.get(data.get("roca_falla", ""), data.get("roca_falla") or None)
    c["F21"] = textos.get("NQ - Barrena", "NQ - Barrena")  # la bitácora de roca no distingue método de muestreo
    c["F22"] = to_float(data.get("roca_velocidad"))

    for fila, t, mm, fuerza in _roca_puntos(data):
        c[f"P{fila}"] = t
        c[f"R{fila}"] = fuerza
        c[f"U{fila}"] = mm

    with open(TEMPLATE_COMPRESION_ROCA, "rb") as f:
        plantilla = f.read()
    salida = _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet1.xml", c),
                                                   TEMPLATE_COMPRESION_ROCA)
    foto = data.get("roca_foto_falla")
    if foto:
        imagen = base64.b64decode(foto["b64"])
        salida = _insertar_imagen_hoja(salida, "xl/drawings/drawing1.xml", imagen, foto["ext"],
                                       **_foto_dentro_de_recuadro(foto))
    return salida

# Carga puntual — índice de fuerza de carga puntual de la roca (ASTM D5731) — bitácora GDA-FL-020 y plantilla
# GDA-FLC-018. Hasta 10 ensayos por muestra; cada uno puede ser diametral, axial, en bloque o irregular, y esa
# elección cambia cómo se calcula el diámetro equivalente.
CP_MAX_ENSAYOS = 10
CP_FILAS_EXCEL = list(range(21, 31))
CP_SENTIDOS = ["DIAMETRAL", "AXIAL", "BLOQUE", "IRREGULAR"]
CP_HUMEDAD_FILAS = [
    ("cp_hum_recipiente", "Recipiente No."), ("cp_hum_humedo", "Masa muestra húmeda + recipiente (g)"),
    ("cp_hum_seco_14", "Masa suelo seco + recipiente (g) (14 horas)"), ("cp_hum_seco_15", "Masa suelo seco + recipiente (g) (15 horas)"),
    ("cp_hum_seco_16", "Masa suelo seco + recipiente (g) (16 horas)"), ("cp_hum_masa_rec", "Masa del recipiente (g)"),
]
EQUIPO_CARGA_PUNTUAL = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Horno GDA-E-007", "Acople carga puntual GDA-E-021",
                        "Pie de rey GDA-E-110", "Celda de carga GDA-E-016", "Celda de carga GDA-E-017", "Máquina multiusos GDA-E-008"]


def _cp_resultado_fila(data, i):
    """(De², De, K, Is, Is50) de un ensayo, con las mismas fórmulas de la plantilla GDA-FLC-018. None si faltan datos."""
    carga = to_float(data.get(f"cp_{i}_carga"))
    d = to_float(data.get(f"cp_{i}_d"))
    sentido = data.get(f"cp_{i}_sentido", "DIAMETRAL")
    if carga is None or d is None:
        return None
    if sentido == "DIAMETRAL":
        de2 = d * d
    else:
        l1, w2 = to_float(data.get(f"cp_{i}_l1"), 0) or 0, to_float(data.get(f"cp_{i}_w2"), 0) or 0
        h = (l1 + w2) / 2
        if h == 0:
            return None
        de2 = (4 * h * d) / math.pi
    if de2 <= 0:
        return None
    de = math.sqrt(de2)
    k = (de / 50) ** 0.45 if (d < 49 or d > 51) else math.sqrt(de / 50)
    is_ = carga * 1000 / de2
    is50 = is_ * k
    return de2, de, k, is_, is50


def resultados_carga_puntual(data):
    """[(etiqueta, valor)] con el Is50 promedio y la humedad, más una fila por ensayo con su Is50."""
    filas = []
    valores_is50 = []
    for i in range(1, CP_MAX_ENSAYOS + 1):
        r = _cp_resultado_fila(data, i)
        if r:
            valores_is50.append(r[4])
            filas.append((f"Ensayo {i} — Is50 (MPa)", fmt_num(r[4], 3)))
    if valores_is50:
        filas.insert(0, ("Is50 promedio (MPa)", fmt_num(sum(valores_is50) / len(valores_is50), 3)))
    humedo, rec = to_float(data.get("cp_hum_humedo")), to_float(data.get("cp_hum_masa_rec"))
    seco = next((v for v in (to_float(data.get(f"cp_hum_seco_{x}")) for x in (16, 15, 14)) if v is not None), None)
    if None not in (humedo, seco, rec) and (seco - rec) != 0:
        filas.append(("Humedad (%)", fmt_num((humedo - seco) / (seco - rec) * 100, 2)))
    return filas


def render_carga_puntual_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora GDA-FL-020. El Excel para descargar (plantilla oficial GDA-FLC-018) "
            "está al final del ensayo. Hasta 10 ensayos por muestra.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Ensayos"), unsafe_allow_html=True)
        head = st.columns([0.6, 1, 1, 1, 1, 1.3])
        for j, texto in enumerate(("#", "Carga P (kN)", "Altura D (mm)", "L1 ó W1 (mm)", "W2 (mm)", "Sentido de falla")):
            head[j].markdown(f'<div class="cell-muted" style="text-align:center;font-weight:700;">{texto}</div>', unsafe_allow_html=True)
        for i in range(1, CP_MAX_ENSAYOS + 1):
            row = st.columns([0.6, 1, 1, 1, 1, 1.3])
            row[0].markdown(f'<div style="padding-top:8px;text-align:center;">{i}</div>', unsafe_allow_html=True)
            for col_i, campo in ((1, "carga"), (2, "d"), (3, "l1"), (4, "w2")):
                key = f"cp_{i}_{campo}"
                data[key] = row[col_i].text_input(f"{campo} {i}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed")
            actual = data.get(f"cp_{i}_sentido", "DIAMETRAL")
            data[f"cp_{i}_sentido"] = row[5].selectbox(
                f"Sentido {i}", CP_SENTIDOS, index=CP_SENTIDOS.index(actual) if actual in CP_SENTIDOS else 0,
                key=f"cp_{i}_sentido_{assay_id}", label_visibility="collapsed")
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
        for key, label in CP_HUMEDAD_FILAS:
            _campo(key, label, placeholder="" if key == "cp_hum_recipiente" else "0.00")
            if key == "cp_hum_seco_14" and data.get("cp_hum_seco_14"):
                for siguiente in ("cp_hum_seco_15", "cp_hum_seco_16"):
                    if not data.get(siguiente):
                        st.session_state[f"{siguiente}_{assay_id}"] = data["cp_hum_seco_14"]
                        data[siguiente] = data["cp_hum_seco_14"]
    with st.container(border=True):
        st.markdown(card_header_html("photo_camera", "Foto de la Muestra"), unsafe_allow_html=True)
        st.caption("Se agrega al Excel, junto a la tabla de ensayos (no queda ajustada a ningún recuadro): la acomodas a "
                   "mano después de descargar.")
        if data.get("cp_foto_falla"):
            st.image(base64.b64decode(data["cp_foto_falla"]["b64"]), width=220)
            if st.button("Quitar foto", key=f"cp_foto_quitar_{assay_id}"):
                data.pop("cp_foto_falla", None)
                data["_cp_foto_intento"] = data.get("_cp_foto_intento", 0) + 1
                if _guardar_inmediato(assay_id, data):
                    st.rerun()
        else:
            captura = st.file_uploader("Foto de la muestra", type=["png", "jpg", "jpeg"],
                                       key=f"cp_foto_archivo_{assay_id}_{data.get('_cp_foto_intento', 0)}",
                                       label_visibility="collapsed")
            if captura is not None:
                with st.spinner("Guardando la foto…"):
                    data["cp_foto_falla"] = _procesar_foto(captura.getvalue())
                    if _guardar_inmediato(assay_id, data):
                        st.rerun()
                    else:
                        data.pop("cp_foto_falla", None)
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_carga_puntual(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digitan los datos de arriba.")
    render_equipo(data, "cp", EQUIPO_CARGA_PUNTUAL)
    render_norma_selector("carga-puntual", data, "cp")


def generar_excel_carga_puntual(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Carga puntual (GDA-FLC-018, ASTM D5731). Se escribe directo en el XML de la hoja "FORMATO" (sheet2.xml en
    esta plantilla) para conservar los diagramas y las firmas.

    Los ensayos 1 a 6 (filas 21-26) tienen fórmulas vivas para De², De, K, Is e Is50 — pero encadenadas entre sí con
    "fórmulas compartidas" de Excel (una fila trae la fórmula completa y las demás solo la referencian por índice);
    si se borra o reemplaza la fila que trae la fórmula completa, las que la referencian quedan apuntando a nada y
    el archivo no vuelve a abrir en Excel (se detectó así: openpyxl seguía abriéndolo bien, pero Excel no). Por eso
    esas columnas (H a M) NUNCA se tocan en esas filas — solo se escriben los datos de entrada (C, D, E, G, N) y
    Excel las calcula solo. Los ensayos 7 a 10 (filas 27-30) son un simple "-" sin fórmula, así que ahí sí se
    escriben ya calculados (misma fórmula que la plantilla). El promedio final (E33, =AVERAGE(M21:M30) en la
    plantilla) se reemplaza por el promedio calculado en Python, porque esa fórmula se rompe (#¡DIV/0!) apenas
    algún ensayo queda sin digitar, sin importar cuántos otros sí tengan resultado."""
    c = {}
    c["D6"] = project.get("cliente", "") if project else ""
    c["D7"] = project["nombre"] if project else codigo
    c["D8"] = project.get("correo_cliente", "") if project else ""
    c["D9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["D10"] = project["muestra_tomada_por"]
    c["L6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["L7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["L8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["M9"] = project.get("numero", "") if project else ""
    c["N9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["D12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["F12"] = perf_codigo
    c["I12"] = muestra["numero"]
    c["L12"] = to_float(muestra.get("profundidad_de"))
    c["N12"] = to_float(muestra.get("profundidad_hasta"))
    c["D13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    valores_is50 = []
    for i, fila in enumerate(CP_FILAS_EXCEL, start=1):
        carga, d = to_float(data.get(f"cp_{i}_carga")), to_float(data.get(f"cp_{i}_d"))
        resultado = _cp_resultado_fila(data, i)
        if resultado:
            valores_is50.append(resultado[4])
        con_formula_viva = i <= 6  # filas 21-26; ver nota de la función sobre las fórmulas compartidas
        if carga is None or d is None:
            if not con_formula_viva:  # en 21-26 no se toca nada si el ensayo quedó sin digitar
                for col in "CDEGHIJKLMN":
                    c[f"{col}{fila}"] = None
            continue
        l1, w2 = to_float(data.get(f"cp_{i}_l1")), to_float(data.get(f"cp_{i}_w2"))
        sentido = data.get(f"cp_{i}_sentido") or "DIAMETRAL"
        c[f"C{fila}"] = carga
        c[f"D{fila}"] = d
        c[f"E{fila}"] = l1
        c[f"G{fila}"] = w2
        c[f"N{fila}"] = sentido
        if not con_formula_viva:
            c[f"H{fila}"] = ((l1 or 0) + (w2 or 0)) / 2 if sentido != "DIAMETRAL" else None
            for col in "IJKLM":
                c[f"{col}{fila}"] = None
            if resultado:
                c[f"I{fila}"], c[f"J{fila}"], c[f"K{fila}"], c[f"L{fila}"], c[f"M{fila}"] = resultado
    c["E33"] = sum(valores_is50) / len(valores_is50) if valores_is50 else None
    c["E36"] = data.get("cp_hum_recipiente") or None
    c["E37"] = to_float(data.get("cp_hum_humedo"))
    c["E38"] = next((v for v in (to_float(data.get(f"cp_hum_seco_{x}")) for x in (16, 15, 14)) if v is not None), None)
    c["E39"] = to_float(data.get("cp_hum_masa_rec"))

    with open(TEMPLATE_CARGA_PUNTUAL, "rb") as f:
        plantilla = f.read()
    salida = _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet2.xml", c),
                                                   TEMPLATE_CARGA_PUNTUAL)
    foto = data.get("cp_foto_falla")
    if foto:
        imagen = base64.b64decode(foto["b64"])
        # Suelta, a la derecha de la tabla de ensayos (columna U, fila 18) — ahí no hay nada más en la plantilla.
        ancho_emu = 2286000
        alto_emu = round(ancho_emu * foto["alto"] / foto["ancho"])
        salida = _insertar_imagen_hoja(salida, "xl/drawings/drawing2.xml", imagen, foto["ext"], col=20, fila=17,
                                       ancho_emu=ancho_emu, alto_emu=alto_emu)
    return salida

# Solidez de los agregados frente a sulfato de sodio o de magnesio (INV E-220, plantilla GDA-FLC-033). No tengo
# bitácora de papel para este ensayo — el formulario se armó directo sobre la plantilla de descarga. La columna
# O (% retenido individual) y P (% retenido acumulado) de la gradación unificada (filas 22-38) son "fórmulas
# compartidas" de Excel (una fila trae la fórmula completa y las demás la referencian por índice) — igual que en
# Carga Puntual, tocarlas rompe el archivo; solo se escribe la columna N (masa retenida corregida), que si es un
# insumo normal.
SULF_TAMICES = [
    ("4\"", 22), ("3 1/2\"", 23), ("3\"", 24), ("2 1/2\"", 25), ("2\"", 26), ("1 1/2\"", 27), ("1\"", 28),
    ("3/4\"", 29), ("1/2\"", 30), ("3/8\"", 31), ("N°4", 32), ("N°8", 33), ("N°16", 34), ("N°30", 35),
    ("N°50", 36), ("N°100", 37), ("<100", 38),
]
SULF_GRUESO_FILAS = [("3\" a 2 1/2\"", 21), ("2 1/2\" a 1 1/2\"", 22), ("1 1/2\" a 3/4\"", 23),
                     ("3/4\" a 3/8\"", 24), ("3/8\" a N°4", 25)]
SULF_FINO_FILAS = [("Menor de N°100", 29), ("N°50 a N°100", 30), ("N°30 a N°50", 31), ("N°16 a N°30", 32),
                   ("N°8 a N°16", 33), ("N°4 a N°8", 34), ("3/8\" a N°4", 35)]
SULF_SOLUCIONES = ["Sulfato de sodio", "Sulfato de magnesio"]
# No hay bitácora física para confirmar los códigos exactos — se usa la misma balanza/horno/tamices que
# Granulometría más lo propio de la inmersión en sulfato; corregir si no coincide con los equipos reales.
EQUIPO_SOLIDEZ_SULFATOS = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Balanza GDA-E-012", "Horno GDA-E-007",
                           "Horno GDA-E-404", "Serie de tamices GDA-E-030 a GDA-E-045", "Recipientes de inmersión"]


def _sulf_gradacion(data):
    """[(fila, N=masa retenida corregida, O=% retenido individual)] de los 17 tamices de la gradación unificada —
    misma fórmula que la plantilla: O = N / masa_inicial_total * 100."""
    total = to_float(data.get("sulf_masa_total"))
    filas = []
    for _label, fila in SULF_TAMICES:
        n = to_float(data.get(f"sulf_grad_{fila}"))
        o = (n / total * 100) if (n is not None and total) else None
        filas.append((fila, n, o))
    return filas


def resultados_solidez_sulfatos(data):
    """Réplica de las fórmulas de la plantilla GDA-FLC-033 para el % de pérdida ponderada, agregado grueso y
    fino por separado (mismo camino: gradación unificada → agrupar por tamaño → % retenido de cada grupo →
    % de pérdida de cada grupo, que se arrastra del grupo anterior si ese grupo pesa menos del 5% → ponderar)."""
    o_por_fila = {fila: o for fila, _n, o in _sulf_gradacion(data)}

    def o(fila):
        return o_por_fila.get(fila) or 0

    l = {25: o(24) + o(25), 27: o(26) + o(27), 29: o(28) + o(29), 31: o(30) + o(31), 32: o(32)}
    l33 = sum(l.values())
    o39 = sum(o(f) for f in range(32, 39))

    def _grupo(filas_pct, filas_frac, base_pct_por_fila):
        # Misma regla que la plantilla: si la fracción pesa >=5% se calcula directo (F-G)/F; si pesa <5% se
        # arrastra el % de pérdida de la fracción anterior. Si pesa >=5% pero faltan las masas F/G, NO se debe
        # arrastrar un valor "parecido" — en Excel esa celda queda en #DIV/0! porque falta digitar esa fracción,
        # así que aquí se marca como None (faltante) y ese hueco también se arrastra tal cual a las fracciones
        # siguientes que dependan de ella (<5%), en vez de inventar un número.
        resultado, h_prev = [], 0.0
        for i, (label, fila) in enumerate(filas_frac, start=1):
            pct = base_pct_por_fila(i)
            f_ini, f_fin = to_float(data.get(f"sulf_{fila}_ini")), to_float(data.get(f"sulf_{fila}_fin"))
            if pct is None:
                h = None
            elif pct >= 5:
                h = (f_ini - f_fin) / f_ini * 100 if (f_ini and f_fin is not None) else None
            else:
                h = h_prev
            h_prev = h
            ponderado = h * pct / 100 if (h is not None and pct is not None) else None
            resultado.append((label, pct, h, ponderado))
        return resultado

    filas_grueso = _grupo(None, SULF_GRUESO_FILAS, lambda i: (l[[25, 27, 29, 31, 32][i - 1]] / l33 * 100) if l33 else None)
    filas_fino = _grupo(None, SULF_FINO_FILAS, lambda i: (o([38, 37, 36, 35, 34, 33, 32][i - 1]) / o39 * 100) if o39 else None)

    filas = []
    if any(p is not None for _l, p, _h, _pon in filas_grueso):
        filas.append(("— Agregado grueso —", ""))
        for label, pct, h, pon in filas_grueso:
            if pct is not None:
                filas.append((f"{label} — % retenido / pérdida", f"{fmt_num(pct, 1)}% / {fmt_num(h, 1) if h is not None else '—'}%"))
        pct_grueso = [p for _l, p, _h, _pon in filas_grueso if p is not None]
        if pct_grueso and all(pon is not None for _l, p, _h, pon in filas_grueso if p is not None):
            total_grueso = sum(pon for _l, _p, _h, pon in filas_grueso if pon is not None)
            filas.append(("% de pérdida ponderada — grueso", fmt_num(total_grueso, 2)))
        elif pct_grueso:
            filas.append(("% de pérdida ponderada — grueso", "Faltan masas por digitar"))
    if any(p is not None for _l, p, _h, _pon in filas_fino):
        filas.append(("— Agregado fino —", ""))
        for label, pct, h, pon in filas_fino:
            if pct is not None:
                filas.append((f"{label} — % retenido / pérdida", f"{fmt_num(pct, 1)}% / {fmt_num(h, 1) if h is not None else '—'}%"))
        pct_fino = [p for _l, p, _h, _pon in filas_fino if p is not None]
        if pct_fino and all(pon is not None for _l, p, _h, pon in filas_fino if p is not None):
            total_fino = sum(pon for _l, _p, _h, pon in filas_fino if pon is not None)
            filas.append(("% de pérdida ponderada — fino", fmt_num(total_fino, 2)))
        elif pct_fino:
            filas.append(("% de pérdida ponderada — fino", "Faltan masas por digitar"))
    return filas


def render_solidez_sulfatos_form(data, assay_id):
    st.info("Formulario armado sobre la plantilla oficial GDA-FLC-033 (no hay bitácora de papel para este ensayo). "
            "El Excel para descargar está al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Gradación de la Muestra Original"), unsafe_allow_html=True)
        _campo("sulf_masa_total", "Masa inicial seca total (g)")
        head = st.columns([1.4, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa retenida corregida (g)</div>',
                          unsafe_allow_html=True)
        for label, fila in SULF_TAMICES:
            row = st.columns([1.4, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            key = f"sulf_grad_{fila}"
            data[key] = row[1].text_input(f"Masa retenida {label}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder="0.00")
    for titulo, filas_frac in (("Ensayo sobre el Agregado Grueso", SULF_GRUESO_FILAS),
                                ("Ensayo sobre el Agregado Fino", SULF_FINO_FILAS)):
        with st.container(border=True):
            st.markdown(card_header_html("science", titulo), unsafe_allow_html=True)
            head = st.columns([1.6, 1, 1])
            head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa inicial (g)</div>', unsafe_allow_html=True)
            head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa final (g)</div>', unsafe_allow_html=True)
            for label, fila in filas_frac:
                row = st.columns([1.6, 1, 1])
                row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                for col_i, sufijo in ((1, "ini"), (2, "fin")):
                    key = f"sulf_{fila}_{sufijo}"
                    data[key] = row[col_i].text_input(f"{label} {sufijo}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                       label_visibility="collapsed", placeholder="0.00")
    with st.container(border=True):
        st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
        actual = data.get("sulf_solucion", SULF_SOLUCIONES[0])
        data["sulf_solucion"] = st.radio("Tipo de solución usada", SULF_SOLUCIONES, horizontal=True,
                                          index=SULF_SOLUCIONES.index(actual) if actual in SULF_SOLUCIONES else 0,
                                          key=f"sulf_solucion_{assay_id}")
        _campo("sulf_ciclos", "Número de ciclos", placeholder="5")
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_solidez_sulfatos(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digita la gradación y las masas de arriba.")
    render_equipo(data, "sulf", EQUIPO_SOLIDEZ_SULFATOS)
    render_norma_selector("solidez-sulfatos", data, "sulf")


def generar_excel_solidez_sulfatos(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Solidez frente a sulfatos (GDA-FLC-033, INV E-220). Se escribe directo en el XML de la hoja: la columna N
    (masa retenida corregida) de la gradación y las masas de las fracciones — el % retenido, % de pérdida y el
    ponderado final los calcula el propio Excel con sus fórmulas (varias son "compartidas", así que no se tocan)."""
    c = {}
    c["C6"] = project.get("cliente", "") if project else ""
    c["C7"] = project["nombre"] if project else codigo
    c["C8"] = project.get("correo_cliente", "") if project else ""
    c["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["C10"] = project["muestra_tomada_por"]
    c["I6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["I7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["I8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["J9"] = project.get("numero", "") if project else ""
    c["K9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["D12"] = perf_codigo
    c["F12"] = muestra["numero"]
    c["I12"] = to_float(muestra.get("profundidad_de"))
    c["K12"] = to_float(muestra.get("profundidad_hasta"))
    c["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    c["O18"] = to_float(data.get("sulf_masa_total"))
    for _label, fila in SULF_TAMICES:
        c[f"N{fila}"] = to_float(data.get(f"sulf_grad_{fila}"))
    for _label, fila in SULF_GRUESO_FILAS + SULF_FINO_FILAS:
        c[f"F{fila}"] = to_float(data.get(f"sulf_{fila}_ini"))
        c[f"G{fila}"] = to_float(data.get(f"sulf_{fila}_fin"))
    textos = {t.strip(): t for t in _textos_compartidos(TEMPLATE_SOLIDEZ_SULFATOS)}
    solucion = (data.get("sulf_solucion") or "").upper()
    c["D38"] = textos.get(solucion, solucion or None)
    c["D39"] = to_float(data.get("sulf_ciclos")) or (5 if data.get("sulf_ciclos") is None else None)

    with open(TEMPLATE_SOLIDEZ_SULFATOS, "rb") as f:
        plantilla = f.read()
    return _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet1.xml", c),
                                                TEMPLATE_SOLIDEZ_SULFATOS)


# Gradación unificada (17 tamices, igual idea que Solidez en Sulfatos) — columnas O (% retenido
# individual) y P (% retenido acumulado) son fórmulas "compartidas" de Excel, así que solo se
# escribe la columna N (masa retenida corregida). El propio formato trae un aviso: estos valores
# se leen del Excel de Granulometría de la misma muestra y se digitan aquí a mano (no hay una
# forma confiable de copiarlos solos porque son "corregidos", no el retenido crudo).
TER_TAMICES = [
    ('6"', 18), ('4"', 19), ('3"', 20), ('2 1/2"', 21), ('2"', 22), ('1 1/2"', 23), ('1"', 24),
    ('3/4"', 25), ('1/2"', 26), ('3/8"', 27), ("N°4", 28), ("N°10", 29), ("N°20", 30), ("N°40", 31),
    ("N°60", 32), ("N°100", 33), ("N°200", 34),
]
# (fracción, fila E/G, tamices de la gradación que se suman para el % retenido acumulado (H),
# masa mínima de la fracción según INV E-211 (g), tamiz designado para lavar esa fracción) — las
# dos últimas ya vienen fijas en la plantilla, se muestran solo como referencia.
TER_FRACCIONES = [
    ('- a 1 1/2"', 18, (18, 19, 20, 21, 22, 23), 5000, "N° 4"),
    ('1 1/2" a 3/4"', 19, (24, 25), 3000, "N° 4"),
    ('3/4" a 3/8"', 20, (26, 27), 2000, "N° 4"),
    ('3/8" a N° 4', 21, (28,), 1000, "N° 8"),
    ("N° 4 a N° 10", 22, (29,), 25, "N° 20"),
]
# No hay bitácora física para confirmar los códigos exactos — se usa la misma balanza/horno/tamices
# que Granulometría/Solidez en Sulfatos; corregir si no coincide con los equipos reales.
EQUIPO_TERRONES_ARCILLA = ["Balanza GDA-E-010", "Balanza GDA-E-011", "Balanza GDA-E-012", "Horno GDA-E-007",
                           "Horno GDA-E-404", "Serie de tamices GDA-E-030 a GDA-E-045", "Recipientes de lavado"]


def _ter_gradacion(data):
    """[(fila, N=masa retenida corregida, O=% retenido individual)] de los 17 tamices — misma
    fórmula que la plantilla: O = N / masa_inicial_seca_total * 100."""
    total = to_float(data.get("ter_masa_total"))
    filas = []
    for _label, fila in TER_TAMICES:
        n = to_float(data.get(f"ter_grad_{fila}"))
        o = (n / total * 100) if (n is not None and total) else None
        filas.append((fila, n, o))
    return filas


def resultados_terrones_arcilla(data):
    """Réplica de las fórmulas de la plantilla GDA-FLC-034 (INV E-211): con la gradación unificada
    se calcula el % retenido acumulado de cada una de las 5 fracciones fijas de la plantilla (H,
    sumando los tamices que le corresponden a cada una) y se combina con el % de terrones y
    partículas deleznables de esa fracción (I = (masa inicial lavada - masa final lavada) / masa
    inicial lavada × 100) para dar el % de pérdida ponderada final = Σ(H_i·I_i)/100, que se compara
    contra el criterio de aceptación INV-320-13 (≤ 2 %, fijo en la plantilla en D41)."""
    o_por_fila = {fila: o for fila, _n, o in _ter_gradacion(data)}

    def o(fila):
        return o_por_fila.get(fila) or 0

    filas, ponderado_total, completas, alguna = [], 0.0, True, False
    for label, fila_eg, tamices_k, _masa_min, _tamiz_perdida in TER_FRACCIONES:
        h = sum(o(f) for f in tamices_k)
        e_i = to_float(data.get(f"ter_{fila_eg}_ini"))
        g_i = to_float(data.get(f"ter_{fila_eg}_fin"))
        if e_i is not None or g_i is not None:
            alguna = True
        i_pct = (e_i - g_i) / e_i * 100 if (e_i and g_i is not None) else None
        if i_pct is None:
            completas = False
        else:
            ponderado_total += h * i_pct / 100
        filas.append((f"{label} — % retenido / % terrones",
                      f"{fmt_num(h, 1)}% / {fmt_num(i_pct, 1) if i_pct is not None else '—'}%"))
    if not alguna:
        return []
    if completas:
        filas.append(("% de pérdida ponderada", fmt_num(ponderado_total, 2)))
        filas.append(("Cumple INV - Art. 320-13 (≤ 2 %)", "Sí" if ponderado_total <= 2 else "No"))
    else:
        filas.append(("% de pérdida ponderada", "Faltan masas por digitar"))
    return filas


def render_terrones_arcilla_form(data, assay_id):
    st.info("Formulario armado sobre la plantilla oficial GDA-FLC-034 (no hay bitácora de papel para este ensayo). "
            "La gradación se lee del Excel de Granulometría de esta misma muestra (la propia plantilla lo indica) "
            "y se digita aquí. El Excel para descargar está al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Gradación de la Muestra (de Granulometría)"), unsafe_allow_html=True)
        _campo("ter_masa_total", "Masa inicial seca total (g)")
        head = st.columns([1.4, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa retenida corregida (g)</div>',
                          unsafe_allow_html=True)
        for label, fila in TER_TAMICES:
            row = st.columns([1.4, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            key = f"ter_grad_{fila}"
            data[key] = row[1].text_input(f"Masa retenida {label}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder="0.00")
    with st.container(border=True):
        st.markdown(card_header_html("science", "Ensayo por Fracción (Terrones y Partículas Deleznables)"),
                    unsafe_allow_html=True)
        head = st.columns([1.8, 1, 1, 1.3])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa inicial lavada seca (g)</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa final lavada seca (g)</div>', unsafe_allow_html=True)
        head[3].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Masa mín. / tamiz pérdida</div>', unsafe_allow_html=True)
        for label, fila, _tamices_k, masa_min, tamiz_perdida in TER_FRACCIONES:
            row = st.columns([1.8, 1, 1, 1.3])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            for col_i, sufijo in ((1, "ini"), (2, "fin")):
                key = f"ter_{fila}_{sufijo}"
                data[key] = row[col_i].text_input(f"{label} {sufijo}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed", placeholder="0.00")
            row[3].markdown(f'<div class="cell-muted" style="text-align:center;padding-top:8px;">'
                             f'{fmt_num(masa_min, 0)} g / {tamiz_perdida}</div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_terrones_arcilla(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran a medida que se digita la gradación y las masas de arriba.")
    render_equipo(data, "ter", EQUIPO_TERRONES_ARCILLA)
    render_norma_selector("terrones-arcilla", data, "ter")


def generar_excel_terrones_arcilla(codigo, perf_codigo, muestra, project, data, observaciones_ensayo=""):
    """Terrones de arcilla y partículas deleznables (GDA-FLC-034, INV E-211). Se escribe directo en
    el XML de la hoja: la gradación unificada (columna N) y las masas de cada fracción — el %
    retenido, el % de terrones y la pérdida ponderada final los calcula el propio Excel con sus
    fórmulas (varias son "compartidas", así que no se tocan)."""
    c = {}
    c["C6"] = project.get("cliente", "") if project else ""
    c["C7"] = project["nombre"] if project else codigo
    c["C8"] = project.get("correo_cliente", "") if project else ""
    c["C9"] = project.get("localizacion", "") if project else ""
    if project and project.get("muestra_tomada_por"):
        c["C10"] = project["muestra_tomada_por"]
    c["H6"] = _fecha_ddmmaaaa(project.get("fecha_recepcion", "")) if project else ""
    c["H7"] = _fecha_ddmmaaaa(project.get("fecha_ejecucion", "")) if project else ""
    c["H8"] = _fecha_ddmmaaaa(project.get("fecha_emision", "")) if project else ""
    c["I9"] = project.get("numero", "") if project else ""
    c["J9"] = project.get("anio", "") if project else ""
    perf = get_perforacion(codigo, perf_codigo)
    c["C12"] = TIPO_PERFORACION_EXCEL.get(perf["tipo"], "") if perf else ""
    c["D12"] = perf_codigo
    c["F12"] = muestra["numero"]
    c["H12"] = to_float(muestra.get("profundidad_de"))
    c["J12"] = to_float(muestra.get("profundidad_hasta"))
    c["C13"] = descripcion_visual_para_excel(muestra) or observaciones_ensayo or ""

    c["O15"] = to_float(data.get("ter_masa_total"))
    for _label, fila in TER_TAMICES:
        c[f"N{fila}"] = to_float(data.get(f"ter_grad_{fila}"))
    for _label, fila, _tamices_k, _masa_min, _tamiz_perdida in TER_FRACCIONES:
        c[f"E{fila}"] = to_float(data.get(f"ter_{fila}_ini"))
        c[f"G{fila}"] = to_float(data.get(f"ter_{fila}_fin"))

    with open(TEMPLATE_TERRONES_ARCILLA, "rb") as f:
        plantilla = f.read()
    return _restaurar_orden_formato_condicional(_xlsx_escribir_celdas(plantilla, "xl/worksheets/sheet1.xml", c),
                                                TEMPLATE_TERRONES_ARCILLA)


def render_limite_contraccion_form(data, assay_id):
    st.info("Formulario armado sobre la plantilla oficial GDA-FLC-022. El Excel para descargar está al final del ensayo.")

    def _campo(key, label, placeholder="0.00"):
        row = st.columns([2.2, 1])
        row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
        data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                       label_visibility="collapsed", placeholder=placeholder)

    with st.container(border=True):
        st.markdown(card_header_html("science", "Datos del Recipiente"), unsafe_allow_html=True)
        _campo("lc_recipiente", "N. recipiente", placeholder="1")
        for key, label in LC_DATOS_RECIPIENTE:
            _campo(key, label)
    with st.container(border=True):
        st.markdown(card_header_html("science", "Peso Unitario de la Muestra"), unsafe_allow_html=True)
        for key, label in LC_DATOS_PASTILLA:
            _campo(key, label)
    with st.container(border=True):
        st.markdown(card_header_html("water_drop", "Humedad"), unsafe_allow_html=True)
        head = st.columns([2, 1, 1])
        head[1].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Natural</div>', unsafe_allow_html=True)
        head[2].markdown('<div class="cell-muted" style="text-align:center;font-weight:700;">Contracción</div>', unsafe_allow_html=True)
        for campo, label in LC_HUMEDAD_FILAS:
            row = st.columns([2, 1, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            for col_i, tipo in ((1, "natural"), (2, "contraccion")):
                key = f"lc_{tipo}_{campo}"
                data[key] = row[col_i].text_input(f"{label} — {tipo}", value=data.get(key, ""), key=f"{key}_{assay_id}",
                                                   label_visibility="collapsed", placeholder="0.00")
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_limite_contraccion(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran cuando estén digitados todos los datos de arriba.")
    render_norma_selector("limite-contraccion", data, "lc")


def render_materia_organica_form(data, assay_id):
    st.info("Formulario armado sobre la plantilla oficial GDA-FLC-003. El Excel para descargar está al final del ensayo.")
    with st.container(border=True):
        st.markdown(card_header_html("science", "Contenido de Materia Orgánica (w %)"), unsafe_allow_html=True)
        row = st.columns([2.2, 1])
        row[0].markdown('<div style="padding-top:8px;">Recipiente No.</div>', unsafe_allow_html=True)
        data["mo_recipiente"] = row[1].text_input("Recipiente No.", value=data.get("mo_recipiente", ""),
                                                   key=f"mo_recipiente_{assay_id}", label_visibility="collapsed")
        for key, label in MO_CAMPOS:
            row = st.columns([2.2, 1])
            row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
            data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                           label_visibility="collapsed", placeholder="0.00")
    with st.container(border=True):
        st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
        filas = resultados_materia_organica(data)
        if filas:
            st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        else:
            st.caption("Se muestran cuando estén digitadas las tres masas.")
    with st.container(border=True):
        st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
        st.caption("Se guardan en la app; la plantilla no tiene casilla para ellos.")
        for key, label, opciones in (("mo_metodo", "Método", MO_METODOS), ("mo_temp", "Temperatura de ignición", MO_TEMPERATURAS)):
            actual = data.get(key, opciones[0])
            data[key] = st.radio(label, opciones, horizontal=True, index=opciones.index(actual) if actual in opciones else 0,
                                  key=f"{key}_{assay_id}")
    render_norma_selector("materia-organica", data, "mo")


def render_proctor_form(data, assay_id):
    st.info("Formulario armado sobre la bitácora oficial GDA-FL-008: el Proctor y, abajo, su CBR de suelos "
            "compactados. Los dos salen en el mismo Excel (GDA-FLC-002) al final del ensayo.")

    def _tabla(titulo, icono, filas):
        with st.container(border=True):
            st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
            head = st.columns([2.2] + [1] * PROCTOR_PRUEBAS)
            for i in range(PROCTOR_PRUEBAS):
                head[i + 1].markdown(f'<div class="cell-muted" style="text-align:center;font-weight:700;">Prueba {i + 1}</div>',
                                      unsafe_allow_html=True)
            for campo, label in filas:
                row = st.columns([2.2] + [1] * PROCTOR_PRUEBAS)
                row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                for i in range(1, PROCTOR_PRUEBAS + 1):
                    key = f"proc_{i}_{campo}"
                    data[key] = row[i].text_input(f"{label} — prueba {i}", value=data.get(key, ""),
                                                   key=f"{key}_{assay_id}", label_visibility="collapsed")

    _tabla("Compactación", "science", PROCTOR_FILAS)
    _tabla("Humedad", "water_drop", PROCTOR_HUMEDAD_FILAS)

    with st.container(border=True):
        st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
        actual = data.get("proc_metodo", PROCTOR_METODOS[0])
        data["proc_metodo"] = st.radio("Método", PROCTOR_METODOS, horizontal=True,
                                        index=PROCTOR_METODOS.index(actual) if actual in PROCTOR_METODOS else 0,
                                        key=f"proc_metodo_{assay_id}")
        row = st.columns([2.2, 1])
        row[0].markdown('<div style="padding-top:8px;">Sobretamaños — tamiz</div>', unsafe_allow_html=True)
        actual_t = data.get("proc_sobretamano_tamiz", "")
        data["proc_sobretamano_tamiz"] = row[1].selectbox(
            "Sobretamaños tamiz", PROCTOR_TAMICES, index=PROCTOR_TAMICES.index(actual_t) if actual_t in PROCTOR_TAMICES else 0,
            key=f"proc_sobretamano_tamiz_{assay_id}", label_visibility="collapsed", format_func=lambda x: x or "—")
        row = st.columns([2.2, 1])
        row[0].markdown('<div style="padding-top:8px;">% retenido de sobretamaños</div>', unsafe_allow_html=True)
        data["proc_sobretamano_pct"] = row[1].text_input("% retenido sobretamaños", value=data.get("proc_sobretamano_pct", ""),
                                                           key=f"proc_sobretamano_pct_{assay_id}", label_visibility="collapsed")

    st.markdown('<div class="section-title">CBR de la muestra compactada (3 moldes)</div>', unsafe_allow_html=True)
    st.caption("Es un CBR aparte del ensayo \"CBR\" (inalterado): va junto al Proctor y sale en el mismo Excel.")

    def _tabla_cbrc(titulo, icono, filas, key_fn):
        with st.container(border=True):
            st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
            head = st.columns([2.2] + [1] * CBRC_MOLDES)
            for i in range(CBRC_MOLDES):
                head[i + 1].markdown(f'<div class="cell-muted" style="text-align:center;font-weight:700;">Molde {i + 1}</div>',
                                      unsafe_allow_html=True)
            for campo, label in filas:
                row = st.columns([2.2] + [1] * CBRC_MOLDES)
                row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                for i in range(1, CBRC_MOLDES + 1):
                    key = key_fn(i, campo)
                    data[key] = row[i].text_input(f"{label} — molde {i}", value=data.get(key, ""),
                                                   key=f"{key}_{assay_id}", label_visibility="collapsed")

    with st.container(border=True):
        st.markdown(card_header_html("science", "Número de golpes por molde"), unsafe_allow_html=True)
        cols_g = st.columns(CBRC_MOLDES)
        for i in range(1, CBRC_MOLDES + 1):
            key = f"cbrc_{i}_golpes"
            data[key] = cols_g[i - 1].text_input(f"Molde {i} — No. de golpes", value=data.get(key, CBRC_GOLPES[i - 1]),
                                                  key=f"{key}_{assay_id}")
    _tabla_cbrc("Datos iniciales", "science", CBRC_FILAS, lambda i, c: f"cbrc_{i}_{c}")
    _tabla_cbrc("Humedad de compactación", "water_drop", CBRC_HUM_FILAS, lambda i, c: f"cbrc_{i}_hc_{c}")
    _tabla_cbrc("Humedad después de inmersión", "water_drop", CBRC_HUM_FILAS, lambda i, c: f"cbrc_{i}_hd_{c}")
    _tabla_cbrc("Expansión", "straighten", CBRC_EXP_FILAS, lambda i, c: f"cbrc_{i}_{c}")

    with st.container(border=True):
        st.markdown(card_header_html("show_chart", "Penetración (fuerza en kN)"), unsafe_allow_html=True)
        with st.expander("Importar resultados desde el Excel de la prensa", icon=":material/upload_file:"):
            st.caption("Un archivo por molde (hoja \"Informe\" del Excel de la prensa).")
            for i in range(1, CBRC_MOLDES + 1):
                archivo = st.file_uploader(f"Molde {i}", type=["xlsx"], key=f"cbrc_upload_{i}_{assay_id}")
                if archivo and st.button(f"Cargar molde {i}", key=f"cbrc_cargar_{i}_{assay_id}", icon=":material/publish:"):
                    valores, aviso = parse_cbr_penetracion_xlsx(archivo.getvalue())
                    if not valores:
                        st.error(aviso)
                    else:
                        for j, val in valores.items():
                            st.session_state[f"cbrc_{i}_pen_{j}_{assay_id}"] = val
                        if aviso:
                            st.warning(aviso)
                        st.success(f"Se cargaron {len(valores)} valores (molde {i}).")
        head = st.columns([1.4] + [1] * CBRC_MOLDES)
        for i in range(CBRC_MOLDES):
            head[i + 1].markdown(f'<div class="cell-muted" style="text-align:center;font-weight:700;">Molde {i + 1}</div>',
                                  unsafe_allow_html=True)
        for j, (pulg, mm) in enumerate(CBR_PENETRACION_FILAS, start=1):
            row = st.columns([1.4] + [1] * CBRC_MOLDES)
            row[0].markdown(f'<div style="padding-top:8px;">{pulg}" ({mm} mm)</div>', unsafe_allow_html=True)
            for i in range(1, CBRC_MOLDES + 1):
                key = f"cbrc_{i}_pen_{j}"
                data[key] = row[i].text_input(f"Fuerza molde {i} {pulg}in", value=data.get(key, ""),
                                               key=f"{key}_{assay_id}", label_visibility="collapsed", placeholder="kN")

    render_equipo(data, "proc", EQUIPO_PROCTOR)
    render_norma_selector("proctor", data, "proc")


def render_gravedad_especifica_form(data, assay_id):
    st.info("Los resultados se calculan aquí con las mismas fórmulas de la plantilla de Excel; el archivo "
            "para descargar está al final del ensayo.")

    def _resultados(modo):
        filas = resultados_gravedad(data, modo)
        with st.container(border=True):
            st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
            if filas:
                st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
            else:
                st.caption("Se muestran cuando estén digitados todos los datos de arriba.")

    def _campos(titulo, icono, campos):
        with st.container(border=True):
            st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
            for key, label in campos:
                row = st.columns([2.2, 1])
                row[0].markdown(f'<div style="padding-top:8px;">{label}</div>', unsafe_allow_html=True)
                data[key] = row[1].text_input(label, value=data.get(key, ""), key=f"{key}_{assay_id}",
                                               label_visibility="collapsed", placeholder="0.00")

    with st.container(border=True):
        st.markdown(card_header_html("science", "¿Cuál ensayo se va a hacer?"), unsafe_allow_html=True)
        actual = data.get("gesp_sel", GESP_OPCIONES[0])
        data["gesp_sel"] = st.radio("Ensayo", GESP_OPCIONES, index=GESP_OPCIONES.index(actual) if actual in GESP_OPCIONES else 0,
                                     key=f"gesp_sel_{assay_id}", label_visibility="collapsed")

    hace_finos = data["gesp_sel"] in (GESP_OPCIONES[0], GESP_OPCIONES[2])
    hace_gruesos = data["gesp_sel"] in (GESP_OPCIONES[1], GESP_OPCIONES[2])
    if hace_finos:
        _campos("Gravedad Específica que pasa el tamiz No. 4", "science", GESP_FINOS_CAMPOS)
        _resultados("fino")
        render_equipo(data, "gesp_finos", EQUIPO_GESP_FINOS)
    if hace_gruesos:
        _campos("Gravedad Específica que retiene el tamiz No. 4", "science", GESP_GRUESOS_CAMPOS)
        _resultados("grueso")
        render_equipo(data, "gesp_gruesos", EQUIPO_GESP_GRUESOS)
    if data["gesp_sel"] == GESP_OPCIONES[3]:
        _campos("Gravedad Específica relativa en arcillas y limos (INV E-128-13)", "science", GESP_ARCILLA_CAMPOS)
        _resultados("arcilla")
        render_equipo(data, "gesp_finos", EQUIPO_GESP_FINOS)
    render_norma_selector("gravedad-especifica", data, "gesp")


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


def render_read_only_summary(tipo, data, laboratorista="—", muestra_id=None):
    """Vista de solo lectura ('Resultados de Ensayo') — la misma para el Jefe (siempre) y para
    el laboratorista cuando el proyecto ya fue ejecutado. Sin casillas de digitación, solo tarjetas
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
            # Tamiz sin digitar = no se pesó nada retenido ahí, no un dato faltante -> se muestra 0.
            sieve_rows = [(label, data.get(key) if data.get(key) not in (None, "") else 0) for key, label, _apert, _cell in SIEVES]
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
    elif tipo == "masa-unitaria" and data.get("mu_metodo") == "Método B":
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Peso Unitario Volumétrico"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Altura de la muestra (mm)", data.get("mub_altura")),
                                          ("Diámetro de la muestra (mm)", data.get("mub_diametro")),
                                          ("Masa de la muestra (g)", data.get("mub_masa"))]), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
            st.markdown(param_table_html([(l, data.get(k)) for k, l in MUB_HUMEDAD_FILAS]), unsafe_allow_html=True)
        resultados = resultados_masa_unitaria_b(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("mub_equipos", []), data.get("mu_norma", "—")
    elif tipo == "masa-unitaria":
        rows = [("Masa en el aire (g)", data.get("mu_peso_aire")), ("Masa en el aire parafinado (g)", data.get("mu_peso_aire_par")),
                ("Masa en el agua parafinado (g)", data.get("mu_peso_agua_par")), ("Temperatura del agua (°C)", data.get("mu_temp_agua"))]
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        if not (get_assay(muestra_id, "humedad") if muestra_id else None):
            with st.container(border=True):
                st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
                st.markdown(param_table_html([(l, data.get(k)) for k, l in MU_HUMEDAD_FILAS]), unsafe_allow_html=True)
        filas, _humedad_pct, _fuente = resultados_masa_unitaria_parafinado(data, muestra_id)
        if filas:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(filas, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("mu_equipos", []), data.get("mu_norma", "—")
    elif tipo == "cbr":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Datos Iniciales"), unsafe_allow_html=True)
            rows = [
                ("Molde No.", data.get("cbr_molde")), ("Diámetro de la muestra (cm)", data.get("cbr_diametro")),
                ("Altura de la muestra (cm)", data.get("cbr_altura")), ("Masa molde (g)", data.get("cbr_masa_molde")),
                ("Masa de la muestra + molde (g) — antes", data.get("cbr_masa_muestra_molde_antes")),
                ("Masa de la muestra + molde (g) — después", data.get("cbr_masa_muestra_molde_despues")),
            ]
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Humedad de inmersión"), unsafe_allow_html=True)
            st.caption("\"Antes\" se comparte con el ensayo de Contenido de Humedad de esta muestra.")
            hum_data = (get_assay(muestra_id, "humedad") or {}).get("data", {}) if muestra_id else {}
            hay_antes = bool(hum_data.get("hum_recipiente") or hum_data.get("hum_seco_mas_recipiente"))
            rows = [
                ("Recipiente", hum_data.get("hum_recipiente") if hay_antes else None, data.get("cbr_desp_recipiente")),
                ("Peso recipiente + suelo húmedo (g)",
                 hum_data.get("hum_masa_humedo_mas_recipiente") if hay_antes else None, data.get("cbr_desp_masa_humedo")),
                ("Peso recipiente + suelo seco (g)",
                 hum_data.get("hum_seco_mas_recipiente") if hay_antes else None, data.get("cbr_desp_masa_seco")),
                ("Peso recipiente (g)", hum_data.get("hum_masa_recipiente") if hay_antes else None,
                 data.get("cbr_desp_masa_recipiente")),
            ]
            if hay_antes:
                rows.append(("Humedad (%)", fmt_num(calcular_humedad_pct(hum_data), decimals=2), None))
            st.markdown(param_table_ncol_html(("PARÁMETRO", "ANTES", "DESPUÉS"), rows), unsafe_allow_html=True)
            if not hay_antes:
                st.caption("El ensayo de Contenido de Humedad de esta muestra todavía no tiene datos.")
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Datos de Expansión"), unsafe_allow_html=True)
            rows = [
                ("Lectura inicial (in)", data.get("cbr_exp_lectura_inicial")),
                ("Lectura final (in)", data.get("cbr_exp_lectura_final")),
            ]
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
            rows = [
                ("Pesas de sobrecarga (g) — antes", data.get("cbr_pesas_antes") or "4554 (por defecto)"),
                ("Pesas de sobrecarga (g) — después", data.get("cbr_pesas_despues") or "4554 (por defecto)"),
                ("Tiempo de inmersión (días) — antes", data.get("cbr_tiempo_inmersion_antes") or "4 (por defecto)"),
                ("Tiempo de inmersión (días) — después", data.get("cbr_tiempo_inmersion_despues") or "4 (por defecto)"),
            ]
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("show_chart", "Penetración"), unsafe_allow_html=True)
            headers = ["PROFUNDIDAD (in)", "FUERZA ANTES (kN)", "FUERZA DESPUÉS (kN)"]
            pen_rows = [(pulg, data.get(f"cbr_pen_antes_{i}"), data.get(f"cbr_pen_despues_{i}"))
                        for i, (pulg, _mm) in enumerate(CBR_PENETRACION_FILAS, start=1)]
            st.markdown(param_table_ncol_html(headers, pen_rows), unsafe_allow_html=True)
        equipos, norma = data.get("cbr_equipos", []), data.get("cbr_norma", "—")
    elif tipo == "compresion-inconfinada":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            filas = [("Tipo de muestra", data.get("ci_condicion")), ("Método de muestreo", data.get("ci_muestreo") if data.get("ci_condicion") == "Inalterada" else None),
                     ("Temperatura inicial (°C)", data.get("ci_temp_ini")), ("Temperatura final (°C)", data.get("ci_temp_fin")),
                     ("Peso de la muestra (g)", data.get("ci_peso"))] + [(l, data.get(k)) for k, l in CI_HUMEDAD_FILAS] + [
                     ("Temperatura de secado", data.get("ci_temp_secado")), ("Método", data.get("ci_metodo")),
                     ("Humedad obtenida", f'{data.get("ci_hum_antes", "")} — {data.get("ci_hum_muestra", "")}'),
                     ("Resistencia al penetrómetro (kg/cm²)", data.get("ci_penetrometro")), ("Tipo de falla", data.get("ci_falla")),
                     ("Velocidad de falla (mm/min)", data.get("ci_velocidad")), ("Tiempo de falla (min)", data.get("ci_tiempo_falla"))]
            st.markdown(param_table_html(filas), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Dimensiones"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["#", "ALTURA (cm)", "DIÁMETRO (cm)"],
                                              [(i, data.get(f"ci_h_{i}"), data.get(f"ci_d_{i}")) for i in (1, 2, 3)]), unsafe_allow_html=True)
        if data.get("ci_maq"):
            with st.container(border=True):
                st.markdown(card_header_html("show_chart", "Datos de la Máquina"), unsafe_allow_html=True)
                st.markdown(param_table_ncol_html(["TIEMPO (s)", "FUERZA (kN)", "DEFORMACIÓN (mm)"],
                                                  [(t, "" if f is None else f, d) for t, f, d in data["ci_maq"]]), unsafe_allow_html=True)
        lecturas = [(d, data.get(f"ci_carga_{i}")) for i, d in enumerate(CI_DEFORMACIONES, start=1) if data.get(f"ci_carga_{i}")]
        if lecturas:
            with st.container(border=True):
                st.markdown(card_header_html("show_chart", "Deformación y Carga"), unsafe_allow_html=True)
                st.markdown(param_table_ncol_html(["DEFORMACIÓN (0.001 in)", "CARGA (kN)"], lecturas), unsafe_allow_html=True)
        resultados = resultados_compresion_inconfinada(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        if data.get("ci_foto_falla"):
            with st.container(border=True):
                st.markdown(card_header_html("photo_camera", "Foto del Plano de Falla"), unsafe_allow_html=True)
                st.image(base64.b64decode(data["ci_foto_falla"]["b64"]), width=220)
        equipos, norma = data.get("ci_equipos", []), data.get("ci_norma", "—")
    elif tipo == "compresion-roca":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            filas = [("Temperatura inicial (°C)", data.get("roca_temp_ini")), ("Temperatura final (°C)", data.get("roca_temp_fin")),
                     ("Peso de la muestra (g)", data.get("roca_peso"))] + [(l, data.get(k)) for k, l in ROCA_HUMEDAD_FILAS] + [
                     ("Temperatura de secado", data.get("roca_temp_secado")), ("Método", data.get("roca_metodo")),
                     ("Humedad obtenida", f'{data.get("roca_hum_antes", "")} — {data.get("roca_hum_muestra", "")}'),
                     ("Esfuerzo máximo leído en la máquina (MPa)", data.get("roca_esfuerzo_maximo_manual")),
                     ("Tipo de falla", data.get("roca_falla")), ("Velocidad de falla (mm/min)", data.get("roca_velocidad")),
                     ("Tiempo de falla (min)", data.get("roca_tiempo_falla"))]
            st.markdown(param_table_html(filas), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Dimensiones"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["#", "ALTURA (cm)", "DIÁMETRO (cm)"],
                                              [(i, data.get(f"roca_h_{i}"), data.get(f"roca_d_{i}")) for i in (1, 2, 3)]), unsafe_allow_html=True)
        lecturas = [(d, data.get(f"roca_carga_{i}")) for i, d in enumerate(ROCA_DEFORMACIONES, start=1) if data.get(f"roca_carga_{i}")]
        if lecturas:
            with st.container(border=True):
                st.markdown(card_header_html("show_chart", "Deformación y Carga"), unsafe_allow_html=True)
                st.markdown(param_table_ncol_html(["DEFORMACIÓN (0.001 in)", "CARGA (kN)"], lecturas), unsafe_allow_html=True)
        resultados = resultados_compresion_roca(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        if data.get("roca_foto_falla"):
            with st.container(border=True):
                st.markdown(card_header_html("photo_camera", "Foto del Plano de Falla"), unsafe_allow_html=True)
                st.image(base64.b64decode(data["roca_foto_falla"]["b64"]), width=220)
        equipos, norma = data.get("roca_equipos", []), data.get("roca_norma", "—")
    elif tipo == "carga-puntual":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Ensayos"), unsafe_allow_html=True)
            filas_tabla = []
            for i in range(1, CP_MAX_ENSAYOS + 1):
                if any(data.get(f"cp_{i}_{c}") for c in ("carga", "d", "l1", "w2")):
                    filas_tabla.append((i, data.get(f"cp_{i}_carga"), data.get(f"cp_{i}_d"), data.get(f"cp_{i}_l1"),
                                        data.get(f"cp_{i}_w2"), data.get(f"cp_{i}_sentido")))
            if filas_tabla:
                st.markdown(param_table_ncol_html(["#", "CARGA P (kN)", "ALTURA D (mm)", "L1 ó W1 (mm)", "W2 (mm)", "SENTIDO"],
                                                  filas_tabla), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Datos de Humedad"), unsafe_allow_html=True)
            st.markdown(param_table_html([(l, data.get(k)) for k, l in CP_HUMEDAD_FILAS]), unsafe_allow_html=True)
        resultados = resultados_carga_puntual(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        if data.get("cp_foto_falla"):
            with st.container(border=True):
                st.markdown(card_header_html("photo_camera", "Foto de la Muestra"), unsafe_allow_html=True)
                st.image(base64.b64decode(data["cp_foto_falla"]["b64"]), width=220)
        equipos, norma = data.get("cp_equipos", []), data.get("cp_norma", "—")
    elif tipo == "solidez-sulfatos":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Gradación de la Muestra Original"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Masa inicial seca total (g)", data.get("sulf_masa_total"))]), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["TAMIZ", "MASA RETENIDA CORREGIDA (g)"],
                                              [(label, data.get(f"sulf_grad_{fila}")) for label, fila in SULF_TAMICES]),
                        unsafe_allow_html=True)
        for titulo, filas_frac in (("Ensayo sobre el Agregado Grueso", SULF_GRUESO_FILAS),
                                    ("Ensayo sobre el Agregado Fino", SULF_FINO_FILAS)):
            with st.container(border=True):
                st.markdown(card_header_html("science", titulo), unsafe_allow_html=True)
                st.markdown(param_table_ncol_html(["FRACCIÓN", "MASA INICIAL (g)", "MASA FINAL (g)"],
                                                  [(label, data.get(f"sulf_{fila}_ini"), data.get(f"sulf_{fila}_fin"))
                                                   for label, fila in filas_frac]), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Tipo de solución usada", data.get("sulf_solucion")),
                                          ("Número de ciclos", data.get("sulf_ciclos"))]), unsafe_allow_html=True)
        resultados = resultados_solidez_sulfatos(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("sulf_equipos", []), data.get("sulf_norma", "—")
    elif tipo == "terrones-arcilla":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Gradación de la Muestra (de Granulometría)"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Masa inicial seca total (g)", data.get("ter_masa_total"))]), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["TAMIZ", "MASA RETENIDA CORREGIDA (g)"],
                                              [(label, data.get(f"ter_grad_{fila}")) for label, fila in TER_TAMICES]),
                        unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("science", "Ensayo por Fracción"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["FRACCIÓN", "MASA INICIAL LAVADA (g)", "MASA FINAL LAVADA (g)"],
                                              [(label, data.get(f"ter_{fila}_ini"), data.get(f"ter_{fila}_fin"))
                                               for label, fila, *_ in TER_FRACCIONES]), unsafe_allow_html=True)
        resultados = resultados_terrones_arcilla(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("ter_equipos", []), data.get("ter_norma", "—")
    elif tipo == "consolidacion":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            filas = ([("Condición inicial", data.get("cons_condicion")), ("Temperatura inicial (°C)", data.get("cons_temp_ini")),
                      ("Temperatura final (°C)", data.get("cons_temp_fin"))]
                     + [(l, data.get(k)) for k, l in CONS_GS_CAMPOS]
                     + [("Temperatura de secado", data.get("cons_temp_secado")), ("Método", data.get("cons_metodo")),
                        ("Consolidómetro", data.get("cons_consolidometro")), ("Precarga (g)", data.get("cons_precarga"))])
            st.markdown(param_table_html(filas), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("science", "Datos de la Muestra"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["PARÁMETRO", "INICIAL", "FINAL"],
                                              [(l, data.get(ki), data.get(kf) or (data.get(ki) if kf != "cons_masa_anillo_muestra_fin" else ""))
                                               for l, ki, kf in CONS_DATOS_MUESTRA]), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Humedad"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["PARÁMETRO", "INICIAL", "FINAL"],
                                              [(l, data.get(f"cons_ini_{c}"), data.get(f"cons_fin_{c}")) for c, l in CONS_HUMEDAD_FILAS]),
                        unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("straighten", "Deformaciones (mm)"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["CARGA (kg)", "CARGA", "DESCARGA"],
                                              [(c, data.get(f"cons_def_{i}_carga"), data.get(f"cons_def_{i}_descarga"))
                                               for i, c in enumerate(CONS_CARGAS, start=1)]), unsafe_allow_html=True)
        resultados = resultados_consolidacion(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = data.get("cons_equipos", []), data.get("cons_norma", "—")
    elif tipo == "limite-contraccion":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            filas = [("N. recipiente", data.get("lc_recipiente"))] + [(l, data.get(k)) for k, l in LC_DATOS_RECIPIENTE + LC_DATOS_PASTILLA]
            st.markdown(param_table_html(filas), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("water_drop", "Humedad"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["PARÁMETRO", "NATURAL", "CONTRACCIÓN"],
                                              [(l, data.get(f"lc_natural_{c}"), data.get(f"lc_contraccion_{c}")) for c, l in LC_HUMEDAD_FILAS]),
                        unsafe_allow_html=True)
        resultados = resultados_limite_contraccion(data)
        if resultados:
            with st.container(border=True):
                st.markdown(card_header_html("calculate", "Resultados"), unsafe_allow_html=True)
                st.markdown(param_table_html(resultados, header_left="RESULTADO", header_right="VALOR"), unsafe_allow_html=True)
        equipos, norma = [], data.get("lc_norma", "—")
    elif tipo == "materia-organica":
        with st.container(border=True):
            st.markdown(card_header_html("science", "Parámetros Registrados"), unsafe_allow_html=True)
            st.markdown(param_table_html([("Recipiente No.", data.get("mo_recipiente"))]
                                         + [(label, data.get(key)) for key, label in MO_CAMPOS]
                                         + resultados_materia_organica(data)
                                         + [("Método", data.get("mo_metodo")), ("Temperatura de ignición", data.get("mo_temp"))]),
                        unsafe_allow_html=True)
        equipos, norma = [], data.get("mo_norma", "—")
    elif tipo == "proctor":
        for titulo, icono, filas in (("Compactación", "science", PROCTOR_FILAS), ("Humedad", "water_drop", PROCTOR_HUMEDAD_FILAS)):
            with st.container(border=True):
                st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
                encabezados = ["PARÁMETRO"] + [f"PRUEBA {i}" for i in range(1, PROCTOR_PRUEBAS + 1)]
                filas_tabla = [(label, *[data.get(f"proc_{i}_{campo}") for i in range(1, PROCTOR_PRUEBAS + 1)])
                               for campo, label in filas]
                st.markdown(param_table_ncol_html(encabezados, filas_tabla), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("tune", "Condiciones del Ensayo"), unsafe_allow_html=True)
            st.markdown(param_table_html([
                ("Método", data.get("proc_metodo")), ("Sobretamaños — tamiz No.", data.get("proc_sobretamano_tamiz")),
                ("% retenido de sobretamaños", data.get("proc_sobretamano_pct"))]), unsafe_allow_html=True)
        for titulo, icono, filas, key_fn in (
                ("CBR compactado — datos iniciales", "science", CBRC_FILAS, lambda i, c: f"cbrc_{i}_{c}"),
                ("CBR compactado — humedad de compactación", "water_drop", CBRC_HUM_FILAS, lambda i, c: f"cbrc_{i}_hc_{c}"),
                ("CBR compactado — humedad después de inmersión", "water_drop", CBRC_HUM_FILAS, lambda i, c: f"cbrc_{i}_hd_{c}"),
                ("CBR compactado — expansión", "straighten", CBRC_EXP_FILAS, lambda i, c: f"cbrc_{i}_{c}")):
            with st.container(border=True):
                st.markdown(card_header_html(icono, titulo), unsafe_allow_html=True)
                encabezados = ["PARÁMETRO"] + [f"MOLDE {i} ({data.get(f'cbrc_{i}_golpes') or CBRC_GOLPES[i - 1]} golpes)"
                                               for i in range(1, CBRC_MOLDES + 1)]
                st.markdown(param_table_ncol_html(encabezados, [(label, *[data.get(key_fn(i, campo)) for i in range(1, CBRC_MOLDES + 1)])
                                                                for campo, label in filas]), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(card_header_html("show_chart", "CBR compactado — penetración (kN)"), unsafe_allow_html=True)
            st.markdown(param_table_ncol_html(["PROFUNDIDAD (in)"] + [f"MOLDE {i}" for i in range(1, CBRC_MOLDES + 1)],
                                              [(pulg, *[data.get(f"cbrc_{i}_pen_{j}") for i in range(1, CBRC_MOLDES + 1)])
                                               for j, (pulg, _mm) in enumerate(CBR_PENETRACION_FILAS, start=1)]),
                        unsafe_allow_html=True)
        equipos, norma = data.get("proc_equipos", []), data.get("proc_norma", "—")
    elif tipo == "gravedad-especifica":
        sel = data.get("gesp_sel", GESP_OPCIONES[0])
        equipos = []
        if sel in (GESP_OPCIONES[0], GESP_OPCIONES[2]):
            with st.container(border=True):
                st.markdown(card_header_html("science", "Gravedad Específica que pasa el tamiz No. 4"), unsafe_allow_html=True)
                st.markdown(param_table_html([(label, data.get(key)) for key, label in GESP_FINOS_CAMPOS]), unsafe_allow_html=True)
            equipos += data.get("gesp_finos_equipos", [])
        if sel in (GESP_OPCIONES[1], GESP_OPCIONES[2]):
            with st.container(border=True):
                st.markdown(card_header_html("science", "Gravedad Específica que retiene el tamiz No. 4"), unsafe_allow_html=True)
                st.markdown(param_table_html([(label, data.get(key)) for key, label in GESP_GRUESOS_CAMPOS]), unsafe_allow_html=True)
            equipos += data.get("gesp_gruesos_equipos", [])
        if sel == GESP_OPCIONES[3]:
            with st.container(border=True):
                st.markdown(card_header_html("science", "Gravedad Específica relativa en arcillas y limos (INV E-128-13)"), unsafe_allow_html=True)
                st.markdown(param_table_html([(label, data.get(key)) for key, label in GESP_ARCILLA_CAMPOS]), unsafe_allow_html=True)
            equipos += data.get("gesp_finos_equipos", [])
        norma = data.get("gesp_norma", "—")
    else:  # "corte-directo"
        with st.container(border=True):
            st.markdown(card_header_html("science", "Datos del Ensayo"), unsafe_allow_html=True)
            rows = [
                ("Tipo de corte directo", CORTE_TIPO_LABELS.get(data.get("corte_tipo"), data.get("corte_tipo"))),
                ("Condición de la muestra", data.get("corte_condicion")),
                ("Temperatura (°C) — inicial", data.get("corte_temp_inicial")),
                ("Temperatura (°C) — final", data.get("corte_temp_final")),
                ("Humedad (%) — inicial", data.get("corte_hum_inicial")),
                ("Humedad (%) — final", data.get("corte_hum_final")),
            ]
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
        for i in (1, 2, 3):
            with st.container(border=True):
                st.markdown(card_header_html("science", f"Muestra {i} — Datos de la Muestra"), unsafe_allow_html=True)
                rows = [(label, data.get(f"corte_m{i}_{campo}")) for campo, label in CORTE_MUESTRA_CAMPOS]
                st.markdown(param_table_html(rows), unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown(card_header_html("water_drop", f"Muestra {i} — Contenido de Humedad"), unsafe_allow_html=True)
                hum_rows = [(label, data.get(f"corte_m{i}_{campo}_inicial"), data.get(f"corte_m{i}_{campo}_final"))
                            for campo, label in CORTE_HUMEDAD_CAMPOS]
                st.markdown(param_table_ncol_html(("PARÁMETRO", "INICIAL", "FINAL"), hum_rows), unsafe_allow_html=True)
                st.caption(f"Temperatura de secado: {data.get(f'corte_m{i}_temp_secado') or '—'}")
                st.caption("Equipos (humedad): " + (", ".join(data.get(f"corte_m{i}_hum_equipos", [])) or "—"))
        with st.container(border=True):
            st.markdown(card_header_html("science", "Gravedad Específica (INV E-128-13)"), unsafe_allow_html=True)
            rows = [(label, data.get(f"corte_ge_{campo}")) for campo, label in CORTE_GRAVEDAD_CAMPOS]
            rows.append(("Picnómetro No.", data.get("corte_ge_picnometro")))
            st.markdown(param_table_html(rows), unsafe_allow_html=True)
            st.caption("Equipos (gravedad): " + (", ".join(data.get("corte_ge_equipos", [])) or "—"))
        equipos, norma = data.get("corte_equipos", []), data.get("corte_norma", "—")

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
    # El Jefe y el Director Técnico solo consultan los ensayos — quien digita los datos de laboratorio es el laboratorista.
    read_only = es_supervisor or (st.session_state.role == "laboratorista" and project_status(codigo) == "ejecutado")

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
            st.markdown(f'<div class="cell-muted" style="margin-top:12px;">Descripción visual de la muestra</div>'
                        f'<div style="font-weight:600;">{html.escape(descripcion_visual_para_excel(muestra) or "— (el laboratorista aún no la digita) —")}</div>',
                        unsafe_allow_html=True)
            # Ver descripcion_visual_calculada: la misma frase pero con el tipo de suelo que de
            # verdad salió en la clasificación USCS, en vez del que se eligió a ojo — debajo de
            # la inicial, no en su lugar.
            descripcion_calc = descripcion_visual_calculada(muestra)
            if descripcion_calc:
                st.markdown(f'<div class="cell-muted" style="margin-top:10px;">Según la clasificación USCS calculada</div>'
                            f'<div style="font-weight:600;color:{PRIMARY};">{html.escape(descripcion_calc)}</div>',
                            unsafe_allow_html=True)

    if muestra is not None:
        with st.container(border=True):
            st.markdown(card_header_html("thermostat", "Condición del Ensayo"), unsafe_allow_html=True)
            with st.expander("Ver temperatura y humedad", icon=":material/thermostat:",
                              expanded=bool(muestra.get("cond_inicial_temp") or muestra.get("cond_inicial_hum"))):
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
                        nuevo_temp = row[1].text_input(
                            f"Temperatura {cond_label}", value=muestra.get(f"cond_{cond_key}_temp", ""),
                            key=f"cond_{cond_key}_temp_{muestra_id}", label_visibility="collapsed", placeholder="0.0")
                        nuevo_hum = row[2].text_input(
                            f"Humedad {cond_label}", value=muestra.get(f"cond_{cond_key}_hum", ""),
                            key=f"cond_{cond_key}_hum_{muestra_id}", label_visibility="collapsed", placeholder="0")
                        cambios_cond = {}
                        if nuevo_temp != muestra.get(f"cond_{cond_key}_temp", ""):
                            cambios_cond[f"cond_{cond_key}_temp"] = nuevo_temp
                        if nuevo_hum != muestra.get(f"cond_{cond_key}_hum", ""):
                            cambios_cond[f"cond_{cond_key}_hum"] = nuevo_hum
                        if cambios_cond:
                            db.update_muestra(muestra["id"], **cambios_cond)
                            muestra.update(cambios_cond)

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
        render_read_only_summary(assay["tipo"], data, assay.get("laboratorist") or "—", muestra_id=muestra_id)
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
                db.update_assay_data(assay["id"], status="en-proceso")
                st.success("Ensayo habilitado — el laboratorista ya puede volver a digitar los datos.")
                st.rerun()
    else:
        intento_incompleto_key = f"_intento_incompleto_{assay_id}"
        if st.session_state.get(intento_incompleto_key):
            faltantes_actuales = campos_faltantes(assay["tipo"], data)
            if faltantes_actuales:
                css_reglas = "".join(
                    f'.st-key-{key}_{assay_id} input {{ border: 2px solid #d32f2f !important; '
                    f'background-color: #fdecea !important; }}\n'
                    for key, _ in faltantes_actuales
                )
                st.markdown(f"<style>{css_reglas}</style>", unsafe_allow_html=True)
                st.error("⚠️ Faltan datos por digitar antes de poder enviar este ensayo a revisión "
                          "(resaltados en rojo abajo):\n\n"
                          + "\n".join(f"- {label}" for _, label in faltantes_actuales))
            else:
                st.session_state.pop(intento_incompleto_key, None)

        if assay["tipo"] == "granulometria":
            render_granulometria_form(data, assay_id)
        elif assay["tipo"] == "pasa200":
            render_pasa200_form(data, assay_id)
        elif assay["tipo"] == "humedad":
            render_humedad_form(data, assay_id)
        elif assay["tipo"] == "limites":
            render_limites_form(data, assay_id)
        elif assay["tipo"] == "masa-unitaria":
            render_masa_unitaria_form(data, assay_id, muestra_id)
        elif assay["tipo"] == "cbr":
            render_cbr_form(data, assay_id, muestra_id)
        elif assay["tipo"] == "corte-directo":
            render_corte_directo_form(data, assay_id)
        elif assay["tipo"] == "gravedad-especifica":
            render_gravedad_especifica_form(data, assay_id)
        elif assay["tipo"] == "proctor":
            render_proctor_form(data, assay_id)
        elif assay["tipo"] == "materia-organica":
            render_materia_organica_form(data, assay_id)
        elif assay["tipo"] == "limite-contraccion":
            render_limite_contraccion_form(data, assay_id)
        elif assay["tipo"] == "consolidacion":
            render_consolidacion_form(data, assay_id)
        elif assay["tipo"] == "compresion-inconfinada":
            render_compresion_inconfinada_form(data, assay_id)
        elif assay["tipo"] == "compresion-roca":
            render_compresion_roca_form(data, assay_id)
        elif assay["tipo"] == "carga-puntual":
            render_carga_puntual_form(data, assay_id)
        elif assay["tipo"] == "solidez-sulfatos":
            render_solidez_sulfatos_form(data, assay_id)
        elif assay["tipo"] == "terrones-arcilla":
            render_terrones_arcilla_form(data, assay_id)

        with st.expander("Observaciones (opcional)", icon=":material/notes:", expanded=bool(assay.get("observations"))):
            observations = st.text_area("Observaciones", value=assay.get("observations", ""), label_visibility="collapsed",
                                         placeholder="Observaciones generales del ensayo, en caso de que se requiera…")

        st.markdown('<div class="section-title">Laboratorista</div>', unsafe_allow_html=True)
        # Ya no se digita a mano — se asigna solo con el nombre de la cuenta con la que se
        # inició sesión (ver profiles.full_name), así no queda a criterio de quien esté
        # digitando escribir cualquier nombre. Si por algo raro la sesión no trae el perfil
        # (no debería pasar acá, esta rama solo la ve un laboratorista con sesión activa), se
        # cae al valor que ya tuviera guardado el ensayo en vez de dejarlo en blanco.
        laboratorist = (st.session_state.profile or {}).get("full_name") or assay.get("laboratorist", "")
        st.markdown(f'<div style="font-weight:600;">{html.escape(laboratorist or "—")}</div>', unsafe_allow_html=True)

        # Autoguardado: si el laboratorista digita y se le olvida darle a "Guardar borrador"
        # antes de salir, los datos no se pierden — se persisten solos en cada rerun (cada vez
        # que se completa un campo), sin necesidad de un clic explícito.
        if (data != assay.get("data", {}) or observations != assay.get("observations", "")
                or laboratorist != assay.get("laboratorist", "")):
            nuevo_status = "en-proceso" if assay["status"] == "sin-iniciar" else assay["status"]
            # Un ensayo con foto (compresión inconfinada) manda un jsonb varias veces más pesado que uno solo con
            # números — en una red de celular lenta la escritura a Supabase puede tardar; se avisa con un spinner
            # en vez de dejar la pantalla quieta sin explicación. Si la escritura falla (señal mala, se cierra la
            # pestaña a medio guardar) se avisa con un error en vez de dejarlo pasar en silencio: sin este intento
            # explícito, la excepción tumbaba toda la página sin decir qué pasó ni qué hacer.
            hay_foto = any(isinstance(v, dict) and "b64" in v for v in data.values())
            try:
                with st.spinner("Guardando la foto…" if hay_foto else "Guardando…"):
                    db.update_assay_data(assay["id"], data=data, observations=observations, laboratorist=laboratorist, status=nuevo_status)
                    if pasa200_gran_sibling:
                        db.update_assay_shared_data(muestra["id"], ["granulometria", "pasa200"], data)
                assay["data"] = data
                assay["observations"] = observations
                assay["laboratorist"] = laboratorist
                assay["status"] = nuevo_status
            except Exception:
                st.error("No se pudo guardar el último cambio (revisa tu conexión). Sigue en pantalla — vuelve a intentarlo "
                          "digitando algo más o dale a \"Guardar borrador\".")
        st.markdown(f'<div class="timestamp-caption">{icon("cloud_done", size=13)} Los cambios se guardan automáticamente mientras digitas.</div>',
                    unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Guardar borrador", use_container_width=True, icon=":material/save:"):
                try:
                    with st.spinner("Guardando…"):
                        db.update_assay_data(assay["id"], data=data, observations=observations, laboratorist=laboratorist, status="en-proceso")
                        if pasa200_gran_sibling:
                            db.update_assay_shared_data(muestra["id"], ["granulometria", "pasa200"], data)
                    assay.update(data=data, observations=observations, laboratorist=laboratorist, status="en-proceso")
                    st.toast("Borrador guardado.", icon=":material/check_circle:")
                except Exception:
                    st.error("No se pudo guardar (revisa tu conexión) — vuelve a intentarlo.")
        with col2:
            if st.button("Enviar a revisión", type="primary", use_container_width=True, icon=":material/send:"):
                faltantes = campos_faltantes(assay["tipo"], data)
                if faltantes:
                    st.session_state[intento_incompleto_key] = True
                    st.rerun()
                else:
                    st.session_state.pop(intento_incompleto_key, None)
                    ya_estaba_finalizado = assay["status"] == "finalizado"
                    db.update_assay_data(assay["id"], data=data, observations=observations, laboratorist=laboratorist, status="finalizado")
                    assay.update(data=data, observations=observations, laboratorist=laboratorist, status="finalizado")
                    if pasa200_gran_sibling:
                        db.update_assay_shared_data(muestra["id"], ["granulometria", "pasa200"], data)
                    if muestra:
                        actor_lab = f"{laboratorist} (Laboratorista)" if laboratorist else "Laboratorista"
                        add_historial(assay, "Enviado a Revisión", actor_lab, icono="science", tono="primary")
                        # Cada ensayo que el laboratorista termina se avisa al Jefe (nunca al Director Técnico,
                        # que solo entra en juego cuando el Jefe confirma ese ensayo individual).
                        if not ya_estaba_finalizado:
                            add_notification("jefe", f"El laboratorista terminó {ASSAY_LABELS[assay['tipo']]} de la Muestra "
                                                      f"{muestra['numero']} de {codigo}.", codigo, perf_codigo, muestra_id)
                            # Si este era el último ensayo pendiente, la muestra completa su ciclo en el
                            # laboratorio (semáforo en rojo) — se avisa aparte para que la confirme.
                            if compute_muestra_estado(muestra) == "finalizado":
                                add_notification("jefe", f"La Muestra {muestra['numero']} de {codigo} ya completó todos sus "
                                                          f"ensayos — está lista para tu confirmación.", codigo, perf_codigo, muestra_id)
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

    if es_supervisor and assay["tipo"] == "masa-unitaria" and muestra and data.get("mu_metodo") != "Método B":
        # Parafinado y Método B usan plantillas de Excel distintas (GDA-FLC-004 y GDA-FLC-030) — cada una con su
        # propio botón, más abajo para Método B.
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_masa_unitaria(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (plantilla oficial de Peso Unitario Parafinado)", icon=":material/download:",
            data=excel_bytes, file_name=f"Peso_unitario_parafinado_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )

    if es_supervisor and assay["tipo"] == "masa-unitaria" and muestra and data.get("mu_metodo") == "Método B":
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Peso Unitario Método B)", icon=":material/download:",
            data=generar_excel_masa_unitaria_b(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Peso_unitario_metodo_b_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_masa_unitaria_b")

    if assay["tipo"] == "compresion-inconfinada" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Compresión Inconfinada)", icon=":material/download:",
            data=generar_excel_compresion_inconfinada(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Compresion_inconfinada_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True, key="dl_compresion_inconfinada")

    if assay["tipo"] == "compresion-roca" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Compresión en Roca)", icon=":material/download:",
            data=generar_excel_compresion_roca(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Compresion_roca_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_compresion_roca")

    if assay["tipo"] == "carga-puntual" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Carga Puntual)", icon=":material/download:",
            data=generar_excel_carga_puntual(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Carga_puntual_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_carga_puntual")

    if assay["tipo"] == "solidez-sulfatos" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Solidez en Sulfatos)", icon=":material/download:",
            data=generar_excel_solidez_sulfatos(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Solidez_sulfatos_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_solidez_sulfatos")

    if assay["tipo"] == "terrones-arcilla" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Terrones de Arcilla)", icon=":material/download:",
            data=generar_excel_terrones_arcilla(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Terrones_arcilla_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_terrones_arcilla")

    if assay["tipo"] == "consolidacion" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Consolidación)", icon=":material/download:",
            data=generar_excel_consolidacion(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Consolidacion_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_consolidacion")

    if assay["tipo"] == "limite-contraccion" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Límite de Contracción)", icon=":material/download:",
            data=generar_excel_limite_contraccion(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Limite_contraccion_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_limite_contraccion")

    if assay["tipo"] == "materia-organica" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Materia Orgánica)", icon=":material/download:",
            data=generar_excel_materia_organica(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Materia_organica_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            key="dl_materia_organica")

    if assay["tipo"] == "proctor" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (plantilla oficial de Proctor)", icon=":material/download:",
            data=generar_excel_proctor(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
            file_name=f"Proctor_{muestra['id_unico']}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True, key="dl_proctor")
        st.caption("Trae el Proctor y su CBR de suelos compactados (3 moldes). El método (A/B/C), la preparación de "
                   "la muestra y el martillo/molde usado no se digitan en la app: se marcan en el Excel.")

    if assay["tipo"] == "gravedad-especifica" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        sel = data.get("gesp_sel", GESP_OPCIONES[0])
        if sel in (GESP_OPCIONES[0], GESP_OPCIONES[2]):
            st.download_button(
                "Descargar Excel (Gravedad específica — agregado fino, INV E-222)", icon=":material/download:",
                data=generar_excel_gravedad_fino(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
                file_name=f"Gravedad_especifica_fino_{muestra['id_unico']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
                key="dl_gesp_fino")
        if sel == GESP_OPCIONES[3]:
            st.download_button(
                "Descargar Excel (Gravedad específica — arcillas y limos, INV E-128)", icon=":material/download:",
                data=generar_excel_gravedad_arcilla(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
                file_name=f"Gravedad_especifica_arcilla_{muestra['id_unico']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
                key="dl_gesp_arcilla")
        if sel in (GESP_OPCIONES[1], GESP_OPCIONES[2]):
            st.download_button(
                "Descargar Excel (Gravedad específica — agregado grueso, INV E-223)", icon=":material/download:",
                data=generar_excel_gravedad_grueso(codigo, perf_codigo, muestra, project, data, assay.get("observations", "")),
                file_name=f"Gravedad_especifica_grueso_{muestra['id_unico']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
                key="dl_gesp_grueso")

    if assay["tipo"] == "corte-directo" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_corte_directo(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (plantilla oficial de Corte Directo)", icon=":material/download:",
            data=excel_bytes, file_name=f"Corte_directo_{muestra['id_unico']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )
        st.caption("Trae los datos de las probetas, la humedad y la gravedad específica. Las lecturas de "
                   "deformación/carga de la máquina no se digitan en la app: el esfuerzo cortante, la cohesión "
                   "y el ángulo de fricción se completan en el Excel.")

    if assay["tipo"] == "cbr" and muestra:
        st.markdown("---")
        st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
        excel_bytes = generar_excel_cbr(codigo, perf_codigo, muestra, project, data, assay.get("observations", ""))
        st.download_button(
            "Descargar Excel (plantilla oficial de CBR)", icon=":material/download:",
            data=excel_bytes, file_name=f"CBR_{muestra['id_unico']}.xlsx",
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
            solicitados = unificar_ensayos([e for e, v in m["ensayos"].items() if v and e in BITACORA_ENSAYOS])
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
            for i, (perf, m, ensayo_label, tipo_interno) in enumerate(rows[start:start + SEARCH_PAGE_SIZE]):
                if i:
                    st.markdown(f'<hr style="margin:8px 0;border-color:{BORDER};">', unsafe_allow_html=True)
                existing = get_assay(m["id_unico"], tipo_interno)
                status = existing["status"] if existing else "sin-iniciar"
                ensayo_id = f'{codigo}-{perf["codigo"]}-M{m["numero"]}'
                cols = st.columns(col_ratios, vertical_alignment="center")
                cols[0].markdown(f'<span class="cell-id">{html.escape(ensayo_id)}</span>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="cell-title">{html.escape(project["nombre"] if project else codigo)}</div>'
                                  f'<div class="cell-sub">{html.escape(codigo)}</div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="cell-title">{html.escape(ensayo_label)}</div>'
                                  f'<div class="cell-sub">Muestra {html.escape(str(m["numero"]))}</div>', unsafe_allow_html=True)
                with cols[3]:
                    st.markdown(f'<div style="text-align:center;">{status_circle_html(status, size=16)}</div>', unsafe_allow_html=True)
                with cols[4]:
                    if st.button("", key=f"search_open_{m['id_unico']}_{tipo_interno}", icon=":material/chevron_right:",
                                 use_container_width=True, help="Abrir"):
                        if existing:
                            st.session_state.selected_assay_id = existing["id"]
                        else:
                            nuevo = db.create_assay(m["id"], tipo_interno)
                            st.session_state.selected_assay_id = nuevo["id"]
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
# El borrado de la cookie de sesión al cerrar sesión tiene el mismo problema de timing que
# escribirla al iniciar sesión (ver _set_session_cookie/_pending_cookie_tokens): si se llama justo
# antes de un st.rerun(), el componente no alcanza a mandarle la orden de borrado al navegador
# antes de que el rerun reemplace la página. Por eso se difiere igual, un rerun después — y como
# el logout deja st.session_state.role en None, este chequeo va ANTES del if/else de abajo, no
# adentro del "else" (que solo corre con sesión iniciada).
if st.session_state.pop("_pending_logout_cookie_clear", False):
    _clear_session_cookie()

if st.session_state.role is None:
    render_login()
else:
    if st.session_state.get("_pending_cookie_tokens"):
        tokens = st.session_state.pop("_pending_cookie_tokens")
        _set_session_cookie(tokens["access_token"], tokens["refresh_token"])
    if "_pending_remember_user" in st.session_state:
        codigo_a_recordar = st.session_state.pop("_pending_remember_user")
        if codigo_a_recordar:
            _set_remember_user_cookie(codigo_a_recordar)
        else:
            _clear_remember_user_cookie()
    if st.session_state.pop("_pending_history_push", False):
        _push_history_entry()
    _load_data(solo_balanzas=st.session_state.screen == "balanzas")
    # Si algo de lo anterior refrescó el token (ver _tokens_rotated en db._refresh_if_needed
    # y en el restore de cookie de init_state), la cookie del navegador queda con un
    # refresh_token ya rotado/vencido si no se vuelve a guardar aquí con el vigente.
    if st.session_state.pop("_tokens_rotated", False):
        tokens_vigentes = db.get_session_tokens()
        if tokens_vigentes:
            _set_session_cookie(tokens_vigentes["access_token"], tokens_vigentes["refresh_token"])
    render_topbar()
    SCREENS = {
        "home": render_home, "new-project": render_new_project, "project-detail": render_project_detail,
        "edit-project": render_edit_project,
        "perforacion-detail": render_perforacion_detail, "muestra-detail": render_muestra_detail,
        "bitacora": render_bitacora, "assay-form": render_assay_form,
        "continue": render_continue, "search": render_search,
        "projects-active": render_projects_active, "projects-done": render_projects_done,
        "balanzas": render_balanzas,
    }
    SCREENS.get(st.session_state.screen, render_home)()
    render_bottomnav()
