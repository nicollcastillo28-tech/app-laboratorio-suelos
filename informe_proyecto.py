"""Informe PDF del proyecto: aspectos generales, ensayos solicitados según la orden vs. realizados (y por qué no se hicieron),
progreso frente a las fechas, avance por perforación y por tipo de ensayo, y detalle por muestra.

Este módulo NO conoce Streamlit ni Supabase: recibe un diccionario ya armado (ver `generar_informe_pdf`) y devuelve los bytes del PDF.
Así se puede probar con datos de ejemplo. app.py arma el diccionario con `_datos_informe_proyecto`.

Solo se usan caracteres que trae la fuente estándar Helvetica (Latin-1 + guiones y comillas de Windows): nada de flechas ni símbolos
matemáticos, para no depender de fuentes externas.
"""
from datetime import date
from io import BytesIO
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as _canvas
from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Flowable, Frame, KeepTogether, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

# Paleta de la app (ver constantes PRIMARY, etc. en app.py)
VERDE = colors.HexColor("#0A7A3A")
VERDE_OSCURO = colors.HexColor("#0B3D22")
VERDE_CLARO = colors.HexColor("#E2EFE3")
PAPEL = colors.HexColor("#F3F1E8")
BORDE = colors.HexColor("#D3D6C9")
TEXTO = colors.HexColor("#10241A")
GRIS = colors.HexColor("#56625A")
AMBAR = colors.HexColor("#D98E04")
AMBAR_CLARO = colors.HexColor("#FBEBC8")
ROJO = colors.HexColor("#A51D1D")
ROJO_CLARO = colors.HexColor("#FBE1DE")
VERDE_CLARO2 = colors.HexColor("#D5EFDC")
PIZARRA = colors.HexColor("#8F9A8B")
GRIS_BARRA = colors.HexColor("#C9CEC4")

ESTADOS = {   # clave -> (etiqueta, color fuerte, color suave)
    "finalizado": ("Realizado", VERDE, VERDE_CLARO2),
    "en-proceso": ("En ejecución", AMBAR, AMBAR_CLARO),
    "sin-iniciar": ("Por iniciar", GRIS_BARRA, colors.HexColor("#ECEEE9")),
    "no-realizado": ("No realizado", ROJO, ROJO_CLARO),
    "sin-formulario": ("Sin formulario", PIZARRA, colors.HexColor("#E6E9E3")),
}

ANCHO, ALTO = A4
MARGEN = 1.6 * cm
ANCHO_UTIL = ANCHO - 2 * MARGEN


def _e(t):
    return escape(str(t if t not in (None, "") else "—"))


def _estilos():
    base = ParagraphStyle("base", fontName="Helvetica", fontSize=9, leading=12, textColor=TEXTO)
    return {
        "base": base,
        "pequeno": ParagraphStyle("pequeno", parent=base, fontSize=7.5, leading=9.5, textColor=GRIS),
        "celda": ParagraphStyle("celda", parent=base, fontSize=8, leading=10),
        "celda_c": ParagraphStyle("celda_c", parent=base, fontSize=8, leading=10, alignment=TA_CENTER),
        "celda_b": ParagraphStyle("celda_b", parent=base, fontSize=8, leading=10, fontName="Helvetica-Bold"),
        "enc": ParagraphStyle("enc", parent=base, fontSize=7.5, leading=9, fontName="Helvetica-Bold", textColor=colors.white),
        "enc_c": ParagraphStyle("enc_c", parent=base, fontSize=7.5, leading=9, fontName="Helvetica-Bold", textColor=colors.white,
                                alignment=TA_CENTER),
        "h2": ParagraphStyle("h2", parent=base, fontName="Helvetica-Bold", fontSize=12.5, leading=15, textColor=VERDE_OSCURO,
                             spaceBefore=12, spaceAfter=5, keepWithNext=1),
        "kpi_n": ParagraphStyle("kpi_n", parent=base, fontName="Helvetica-Bold", fontSize=20, leading=22, alignment=TA_CENTER),
        "kpi_t": ParagraphStyle("kpi_t", parent=base, fontSize=7.5, leading=9, alignment=TA_CENTER, textColor=GRIS),
        "titulo": ParagraphStyle("titulo", parent=base, fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=colors.white),
        "subtitulo": ParagraphStyle("subtitulo", parent=base, fontSize=10.5, leading=14, textColor=colors.HexColor("#CFE6D6")),
        "diag": ParagraphStyle("diag", parent=base, fontSize=9.5, leading=13),
    }


# ── Dibujos ───────────────────────────────────────────────────────────────────────────────────────────────────────
class Barra(Flowable):
    """Barra horizontal. `segmentos` = [(valor, color)] apilados sobre un fondo gris; el ancho total es la suma de los valores o `total`."""

    def __init__(self, segmentos, ancho, alto=8, total=None, texto=None, radio=None):
        super().__init__()
        self.segmentos, self.ancho, self.alto = segmentos, ancho, alto
        self.total = total if total is not None else sum(v for v, _ in segmentos)
        self.texto, self.radio = texto, (radio if radio is not None else alto / 2)
        self.ancho_barra = ancho - (30 if texto else 0)   # con texto, el porcentaje se escribe a la derecha de la barra

    def wrap(self, *_):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        c.setFillColor(colors.HexColor("#E7E9E2"))
        c.roundRect(0, 0, self.ancho_barra, self.alto, self.radio, stroke=0, fill=1)
        if self.total > 0:
            x = 0.0
            ult = max((i for i, (v, _c) in enumerate(self.segmentos) if v > 0), default=-1)
            for i, (v, col) in enumerate(self.segmentos):
                if v <= 0:
                    continue
                w = self.ancho_barra * v / self.total
                c.setFillColor(col)
                if i == ult and x == 0:
                    c.roundRect(x, 0, w, self.alto, self.radio, stroke=0, fill=1)
                else:
                    c.rect(x, 0, w, self.alto, stroke=0, fill=1)
                x += w
        if self.texto:
            c.setFont("Helvetica-Bold", 7.5)
            c.setFillColor(TEXTO)
            c.drawRightString(self.ancho, self.alto / 2 - 2.6, self.texto)


class LineaDeTiempo(Flowable):
    """Línea de tiempo del proyecto: inicio -> final planeado (y final real), con el tramo transcurrido y la marca de hoy."""

    def __init__(self, ancho, inicio, final, final_real, hoy):
        super().__init__()
        self.ancho, self.alto = ancho, 54
        self.inicio, self.final, self.final_real, self.hoy = inicio, final, final_real, hoy

    def wrap(self, *_):
        return self.ancho, self.alto

    def _x(self, d, x0, x1, ini, fin):
        dias = max((fin - ini).days, 1)
        return x0 + (x1 - x0) * min(max((d - ini).days / dias, 0), 1)

    def draw(self):
        c = self.canv
        x0, x1, y = 14, self.ancho - 14, 26
        ini = self.inicio
        fin = max([d for d in (self.final, self.final_real, self.hoy) if d] + [ini])
        c.setFillColor(colors.HexColor("#E7E9E2"))
        c.roundRect(x0, y - 4, x1 - x0, 8, 4, stroke=0, fill=1)
        xh = self._x(self.hoy, x0, x1, ini, fin)
        c.setFillColor(VERDE)
        c.roundRect(x0, y - 4, max(xh - x0, 6), 8, 4, stroke=0, fill=1)

        def marca(d, rotulo, color, arriba=True, negrita=False):
            x = self._x(d, x0, x1, ini, fin)
            c.setStrokeColor(color)
            c.setFillColor(color)
            c.setLineWidth(1.4)
            c.line(x, y - 9, x, y + 9)
            c.circle(x, y, 3.4, stroke=0, fill=1)
            c.setFont("Helvetica-Bold" if negrita else "Helvetica", 7)
            c.drawCentredString(min(max(x, x0 + 22), x1 - 22), y + (14 if arriba else -20), rotulo)
            c.setFont("Helvetica", 6.6)
            c.setFillColor(GRIS)
            c.drawCentredString(min(max(x, x0 + 22), x1 - 22), y + (6.5 + 14 if arriba else -28), d.strftime("%d/%m/%Y"))

        marca(ini, "Inicio", VERDE_OSCURO, arriba=False)
        if self.final:
            marca(self.final, "Final planeado", GRIS, arriba=False)
        if self.final_real:
            marca(self.final_real, "Final real", ROJO, arriba=True)
        marca(self.hoy, "Hoy", AMBAR, arriba=True, negrita=True)


class Banda(Flowable):
    """Encabezado de la primera página: franja verde oscura con el código, el nombre y el estado del proyecto."""

    def __init__(self, ancho, codigo, nombre, estado, subtitulo, est):
        super().__init__()
        self.ancho, self.alto = ancho, 3.1 * cm
        self.codigo, self.nombre, self.estado, self.subtitulo, self.est = codigo, nombre, estado, subtitulo, est

    def wrap(self, *_):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        c.setFillColor(VERDE_OSCURO)
        c.roundRect(0, 0, self.ancho, self.alto, 12, stroke=0, fill=1)
        c.setFillColor(colors.HexColor("#6FE39A"))
        c.roundRect(0, 0, 6, self.alto, 3, stroke=0, fill=1)
        c.setFont("Helvetica-Bold", 8.5)
        c.setFillColor(colors.HexColor("#6FE39A"))
        c.drawString(18, self.alto - 20, "INFORME DE AVANCE DEL PROYECTO")
        c.setFont("Helvetica-Bold", 22)
        c.setFillColor(colors.white)
        c.drawString(18, self.alto - 46, self.codigo)
        c.setFont("Helvetica", 11)
        c.setFillColor(colors.HexColor("#CFE6D6"))
        nombre = self.nombre
        while c.stringWidth(nombre, "Helvetica", 11) > self.ancho - 150 and len(nombre) > 4:
            nombre = nombre[:-2]
        c.drawString(18, self.alto - 64, nombre + ("…" if nombre != self.nombre else ""))
        c.setFont("Helvetica", 8)
        c.drawString(18, 12, self.subtitulo)
        # chip de estado
        c.setFont("Helvetica-Bold", 9)
        w = c.stringWidth(self.estado, "Helvetica-Bold", 9) + 22
        c.setFillColor(self.est)
        c.roundRect(self.ancho - w - 16, self.alto - 34, w, 20, 10, stroke=0, fill=1)
        c.setFillColor(colors.white if self.est in (VERDE, ROJO) else TEXTO)
        c.drawCentredString(self.ancho - w / 2 - 16, self.alto - 27.5, self.estado)


class _NumeradoCanvas(_canvas.Canvas):
    """Canvas que sabe cuántas páginas tiene el documento: pie 'Página X de Y'."""

    def __init__(self, *a, pie="", **k):
        super().__init__(*a, **k)
        self._paginas = []
        self._pie = pie

    def showPage(self):
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self._pie_y_cabecera(total)
            super().showPage()
        super().save()

    def _pie_y_cabecera(self, total):
        n = self._pageNumber
        if n > 1:
            self.setFillColor(VERDE_OSCURO)
            self.rect(0, ALTO - 0.95 * cm, ANCHO, 0.95 * cm, stroke=0, fill=1)
            self.setFillColor(colors.white)
            self.setFont("Helvetica-Bold", 8)
            self.drawString(MARGEN, ALTO - 0.6 * cm, "GEODELTA INGENIERÍA S.A.S.")
            self.setFont("Helvetica", 8)
            self.drawRightString(ANCHO - MARGEN, ALTO - 0.6 * cm, self._titulo_cabecera)
        self.setStrokeColor(BORDE)
        self.setLineWidth(0.6)
        self.line(MARGEN, 1.25 * cm, ANCHO - MARGEN, 1.25 * cm)
        self.setFont("Helvetica", 7.5)
        self.setFillColor(GRIS)
        self.drawString(MARGEN, 0.8 * cm, self._pie)
        self.drawRightString(ANCHO - MARGEN, 0.8 * cm, f"Página {n} de {total}")


# ── Piezas reutilizables ──────────────────────────────────────────────────────────────────────────────────────────
def _tabla(filas, anchos, encabezado=True, estilos_extra=(), repetir=1):
    t = Table(filas, colWidths=anchos, repeatRows=repetir if encabezado else 0)
    est = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, BORDE),
    ]
    if encabezado:
        est += [("BACKGROUND", (0, 0), (-1, 0), VERDE_OSCURO), ("LINEBELOW", (0, 0), (-1, 0), 0, VERDE_OSCURO)]
        for i in range(1, len(filas)):
            if i % 2 == 0:
                est.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F7F8F4")))
    t.setStyle(TableStyle(est + list(estilos_extra)))
    return t


_TEXTO_ESTADO = {"finalizado": "0A5A2A", "en-proceso": "8A4B06", "sin-iniciar": "56625A", "no-realizado": "A51D1D", "sin-formulario": "56625A"}


def _chip(clave, st):
    etiqueta = ESTADOS[clave][0]
    return Paragraph(f'<font color="#{_TEXTO_ESTADO[clave]}"><b>{etiqueta}</b></font>', st["celda_c"])


def _kpi(valor, titulo, color, fondo, st, ancho):
    t = Table([[Paragraph(f'<font color="#{color.hexval()[2:]}">{valor}</font>', st["kpi_n"])], [Paragraph(titulo, st["kpi_t"])]],
              colWidths=[ancho])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), fondo), ("BOX", (0, 0), (-1, -1), 0.6, BORDE),
        ("ROUNDEDCORNERS", [7, 7, 7, 7]), ("TOPPADDING", (0, 0), (-1, 0), 7), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
        ("TOPPADDING", (0, 1), (-1, 1), 0), ("BOTTOMPADDING", (0, 0), (-1, 0), 1),
    ]))
    return t


def _tabla_tipos(lista, st):
    """Tabla 'ensayo | solicitados | realizados | en ejecución | por iniciar | no realizados | avance' para una lista de tipos de ensayo."""
    enc = ["Ensayo", "Solicitados", "Realizados", "En ejecución", "Por iniciar", "No realizados", "Avance"]
    filas = [[Paragraph(t, st["enc"] if i < 1 else st["enc_c"]) for i, t in enumerate(enc)]]
    anchos = [4.2 * cm, 1.8 * cm, 1.8 * cm, 1.9 * cm, 1.8 * cm, 2.1 * cm, 3.3 * cm]
    escala = ANCHO_UTIL / sum(anchos)
    anchos = [x * escala for x in anchos]
    for t in lista:
        cerrados = t["finalizados"] + t["no_realizados"]
        segs = [(t["finalizados"], VERDE), (t["no_realizados"], ROJO), (t["en_curso"], AMBAR), (t["por_iniciar"] + t["sin_formulario"], GRIS_BARRA)]
        filas.append([Paragraph(_e(t["ensayo"]), st["celda_b"]), Paragraph(str(t["solicitados"]), st["celda_c"]),
                      Paragraph(str(t["finalizados"]), st["celda_c"]), Paragraph(str(t["en_curso"]), st["celda_c"]),
                      Paragraph(str(t["por_iniciar"] + t["sin_formulario"]), st["celda_c"]), Paragraph(str(t["no_realizados"]), st["celda_c"]),
                      Barra(segs, anchos[6] - 12, 9, total=max(t["solicitados"], 1), texto=f"{_pct(cerrados, t['solicitados']):.0f} %")])
    return _tabla(filas, anchos)


def _pct(a, b):
    return (100.0 * a / b) if b else 0.0


def _dias_txt(n):
    return f"{abs(n)} día{'s' if abs(n) != 1 else ''}"


def _diagnostico(plazo, tot, avance):
    """(texto, color) con el avance de ensayos frente al tiempo del proyecto."""
    hoy, ini, fin, fin_real = plazo["hoy"], plazo.get("inicio"), plazo.get("final"), plazo.get("final_real")
    meta = fin_real or fin
    if tot["solicitados"] and avance >= 99.99:
        return "Todos los ensayos solicitados están cerrados.", VERDE
    if not (ini and meta):
        return "Faltan las fechas de inicio y final del proyecto para comparar el avance con el plazo.", GRIS
    total_dias = max((meta - ini).days, 1)
    pasados = (hoy - ini).days
    tiempo = min(max(100.0 * pasados / total_dias, 0), 100)
    if hoy > meta:
        return (f"El plazo venció hace {_dias_txt((hoy - meta).days)} y faltan {tot['solicitados'] - tot['cerrados']} ensayo(s) por cerrar "
                f"({avance:.0f} % de avance)."), ROJO
    restan = (meta - hoy).days
    dif = avance - tiempo
    if pasados < 0:
        return f"El proyecto empieza en {_dias_txt(pasados)}; faltan {_dias_txt(restan)} para el final.", GRIS
    if dif >= -5:
        return f"En tiempo: {avance:.0f} % de ensayos cerrados con {tiempo:.0f} % del plazo transcurrido; quedan {_dias_txt(restan)}.", VERDE
    if dif >= -20:
        return (f"Ligero retraso: {avance:.0f} % de ensayos cerrados frente a {tiempo:.0f} % del plazo transcurrido; "
                f"quedan {_dias_txt(restan)}."), AMBAR
    return (f"Atrasado: {avance:.0f} % de ensayos cerrados frente a {tiempo:.0f} % del plazo transcurrido; "
            f"quedan {_dias_txt(restan)}."), ROJO


def generar_informe_pdf(d):
    """Devuelve los bytes del PDF. `d` (todas las fechas de texto ya en DD/MM/AAAA, y las de `plazo` como date o None):
      generado, generado_por, proyecto{codigo, nombre, localizacion, norma, cliente, contacto, telefono, correo, muestra_tomada_por,
      fecha_bitacora, fecha_ingreso, fecha_inicio, fecha_final, fecha_final_real, estado},
      plazo{inicio, final, final_real, hoy}, totales{perforaciones, muestras_orden, muestras_ensayo, solicitados, finalizados,
      no_realizados, en_curso, por_iniciar, sin_formulario, confirmados_jefe, aprobados_dt},
      perforaciones[{codigo, tipo, muestras_orden, muestras_ensayo, solicitados, finalizados, no_realizados, en_curso, por_iniciar,
      sin_formulario, recepcion, ejecucion, emision, por_tipo[{ensayo, solicitados, ...}]}], alertas[{perforacion, muestra, ensayo, nivel
      ('alto'|'revisar'), nivel_txt, mensaje}], por_tipo[{ensayo, solicitados, finalizados, no_realizados, en_curso, por_iniciar,
      sin_formulario}], no_hechos[{perforacion, muestra, ensayo, estado, motivo}], detalle[{perforacion, muestra, ensayo, estado, revision,
      actualizado, nota}], observaciones[{perforacion, muestra, texto}]."""
    st = _estilos()
    p, tot, plazo = d["proyecto"], dict(d["totales"]), d["plazo"]
    tot["cerrados"] = tot["finalizados"] + tot["no_realizados"]
    avance = _pct(tot["cerrados"], tot["solicitados"])
    historia = []

    # ── Portada: banda + aspectos generales ─────────────────────────────────────────────────────────────────────────
    est_color = {"Ejecutado": VERDE, "En ejecución": AMBAR}.get(p.get("estado"), GRIS)
    historia.append(Banda(ANCHO_UTIL, p["codigo"], p["nombre"], p.get("estado", ""),
                          f"Generado el {d['generado']} (hora de Colombia) · {d.get('generado_por', '')}", est_color))
    historia.append(Spacer(1, 6))

    historia.append(Paragraph("Aspectos generales", st["h2"]))
    kv = [("Proyecto", p["nombre"]), ("Código interno", p["codigo"]), ("Localización", p.get("localizacion")), ("Norma", p.get("norma")),
          ("Cliente", p.get("cliente")), ("Contacto", p.get("contacto")), ("Teléfono", p.get("telefono")), ("Correo", p.get("correo")),
          ("Muestra tomada por", p.get("muestra_tomada_por")), ("Fecha de la bitácora", p.get("fecha_bitacora")),
          ("Inicio del proyecto", p.get("fecha_inicio")), ("Final planeado", p.get("fecha_final")),
          ("Final real", p.get("fecha_final_real")), ("Perforaciones", tot["perforaciones"])]
    filas = []
    for i in range(0, len(kv), 2):
        izq, der = kv[i], kv[i + 1] if i + 1 < len(kv) else ("", "")
        filas.append([Paragraph(f"<b>{izq[0]}</b>", st["pequeno"]), Paragraph(_e(izq[1]), st["celda"]),
                      Paragraph(f"<b>{der[0]}</b>" if der[0] else "", st["pequeno"]), Paragraph(_e(der[1]) if der[0] else "", st["celda"])])
    historia.append(_tabla(filas, [3.0 * cm, ANCHO_UTIL / 2 - 3.0 * cm, 3.0 * cm, ANCHO_UTIL / 2 - 3.0 * cm], encabezado=False,
                           estilos_extra=[("BACKGROUND", (0, 0), (0, -1), VERDE_CLARO), ("BACKGROUND", (2, 0), (2, -1), VERDE_CLARO)]))

    # ── Resumen de ensayos ──────────────────────────────────────────────────────────────────────────────────────────
    historia.append(Paragraph("Resumen de ensayos según la orden", st["h2"]))
    w = (ANCHO_UTIL - 4 * 0.25 * cm) / 5
    kpis = [_kpi(tot["solicitados"], "Solicitados<br/>en la orden", VERDE_OSCURO, VERDE_CLARO, st, w),
            _kpi(tot["finalizados"], "Realizados<br/>con resultado", VERDE, VERDE_CLARO2, st, w),
            _kpi(tot["en_curso"], "En ejecución<br/>en el laboratorio", AMBAR, AMBAR_CLARO, st, w),
            _kpi(tot["por_iniciar"] + tot["sin_formulario"], "Por iniciar<br/>sin datos aún", GRIS, colors.HexColor("#ECEEE9"), st, w),
            _kpi(tot["no_realizados"], "No realizados<br/>con su motivo", ROJO, ROJO_CLARO, st, w)]
    k = Table([kpis], colWidths=[w] * 5, hAlign="LEFT")
    k.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0.25 * cm)]))
    historia.append(k)
    historia.append(Spacer(1, 7))
    segs = [(tot["finalizados"], VERDE), (tot["no_realizados"], ROJO), (tot["en_curso"], AMBAR), (tot["por_iniciar"] + tot["sin_formulario"], GRIS_BARRA)]
    historia.append(Paragraph(f"<b>Avance de ensayos: {avance:.0f} %</b>  ({tot['cerrados']} de {tot['solicitados']} cerrados, entre realizados y no realizados)",
                              st["base"]))
    historia.append(Spacer(1, 3))
    historia.append(Barra(segs, ANCHO_UTIL, 13, total=max(tot["solicitados"], 1)))
    leyenda = "  ".join(f'<font color="#{c.hexval()[2:]}">&#9632;</font> {t}' for t, c in
                        (("Realizados", VERDE), ("No realizados", ROJO), ("En ejecución", AMBAR), ("Por iniciar", GRIS_BARRA)))
    historia.append(Spacer(1, 2))
    historia.append(Paragraph(leyenda, st["pequeno"]))
    historia.append(Spacer(1, 8))
    filas = [[Paragraph("<b>Muestras en la orden</b>", st["pequeno"]), Paragraph(str(tot["muestras_orden"]), st["celda_b"]),
              Paragraph("<b>Muestras con ensayo</b>", st["pequeno"]), Paragraph(str(tot["muestras_ensayo"]), st["celda_b"]),
              Paragraph("<b>Confirmados por el Jefe</b>", st["pequeno"]), Paragraph(f"{tot['confirmados_jefe']} de {tot['finalizados'] + tot['no_realizados']}", st["celda_b"]),
              Paragraph("<b>Aprobados por el DT</b>", st["pequeno"]), Paragraph(f"{tot['aprobados_dt']} de {tot['finalizados'] + tot['no_realizados']}", st["celda_b"])]]
    historia.append(_tabla(filas, [2.9 * cm, 1.1 * cm, 2.9 * cm, 1.1 * cm, 3.4 * cm, 1.5 * cm, 3.0 * cm, 1.5 * cm][:8], encabezado=False,
                           estilos_extra=[("BACKGROUND", (0, 0), (-1, -1), PAPEL)]))
    if tot["muestras_orden"] > tot["muestras_ensayo"]:
        historia.append(Spacer(1, 3))
        historia.append(Paragraph(
            f"De las {tot['muestras_orden']} muestras que trae la orden, {tot['muestras_ensayo']} tienen ensayos marcados; las demás se recibieron sin ensayo asignado.",
            st["pequeno"]))

    # ── Progreso frente a las fechas ─────────────────────────────────────────────────────────────────────────────────
    historia.append(Paragraph("Progreso con respecto a las fechas", st["h2"]))
    bloque = []
    texto, color = _diagnostico(plazo, tot, avance)
    ini, fin, fin_real, hoy = plazo.get("inicio"), plazo.get("final"), plazo.get("final_real"), plazo["hoy"]
    if ini and (fin or fin_real):
        bloque.append(LineaDeTiempo(ANCHO_UTIL, ini, fin, fin_real, hoy))
        bloque.append(Spacer(1, 4))
        meta = fin_real or fin
        tiempo = min(max(100.0 * (hoy - ini).days / max((meta - ini).days, 1), 0), 100)
        mitad = (ANCHO_UTIL - 0.6 * cm) / 2
        dos = Table([[Paragraph(f"<b>Tiempo transcurrido: {tiempo:.0f} %</b>", st["celda"]),
                      Paragraph(f"<b>Avance de ensayos: {avance:.0f} %</b>", st["celda"])],
                     [Barra([(tiempo, AMBAR)], mitad, 11, total=100), Barra([(avance, VERDE)], mitad, 11, total=100)]],
                    colWidths=[mitad + 0.3 * cm, mitad + 0.3 * cm], hAlign="LEFT")
        dos.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0.3 * cm),
                                 ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
        bloque.append(dos)
        bloque.append(Spacer(1, 5))
    caja = Table([[Paragraph(f'<font color="#{color.hexval()[2:]}"><b>{_e(texto)}</b></font>', st["diag"])]], colWidths=[ANCHO_UTIL])
    caja.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, color), ("BACKGROUND", (0, 0), (-1, -1), PAPEL),
                              ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                              ("LEFTPADDING", (0, 0), (-1, -1), 9)]))
    bloque.append(caja)
    historia.append(KeepTogether(bloque))

    # ── Avance por perforación ───────────────────────────────────────────────────────────────────────────────────────
    historia.append(Paragraph("Avance por perforación", st["h2"]))
    historia.append(Paragraph("Recepción: día en que entró la orden de la perforación. Ejecución: día en que el laboratorista digitó el primer dato. "
                              "Emisión: última descarga de un Excel de resultados de la perforación.", st["pequeno"]))
    historia.append(Spacer(1, 4))
    enc = ["Perforación", "Muestras<br/>ensayo / orden", "Ensayos<br/>cerrados / solicitados", "Avance", "Recepción", "Ejecución", "Emisión"]
    filas = [[Paragraph(t, st["enc"] if i < 1 else st["enc_c"]) for i, t in enumerate(enc)]]
    anchos = [2.3 * cm, 2.3 * cm, 2.9 * cm, 3.6 * cm, 2.15 * cm, 2.15 * cm, 2.15 * cm]
    escala = ANCHO_UTIL / sum(anchos)
    anchos = [a * escala for a in anchos]
    for pf in d["perforaciones"]:
        cerrados = pf["finalizados"] + pf["no_realizados"]
        segs = [(pf["finalizados"], VERDE), (pf["no_realizados"], ROJO), (pf["en_curso"], AMBAR),
                (pf["por_iniciar"] + pf["sin_formulario"], GRIS_BARRA)]
        filas.append([Paragraph(f"<b>{_e(pf['codigo'])}</b><br/><font size=7 color='#56625A'>{_e(pf['tipo'])}</font>", st["celda"]),
                      Paragraph(f"{pf['muestras_ensayo']} / {pf['muestras_orden']}", st["celda_c"]),
                      Paragraph(f"{cerrados} / {pf['solicitados']}", st["celda_c"]),
                      Barra(segs, anchos[3] - 12, 9, total=max(pf["solicitados"], 1), texto=f"{_pct(cerrados, pf['solicitados']):.0f} %"),
                      Paragraph(_e(pf["recepcion"]), st["celda_c"]), Paragraph(_e(pf["ejecucion"]), st["celda_c"]),
                      Paragraph(_e(pf["emision"]), st["celda_c"])])
    if len(filas) == 1:
        filas.append([Paragraph("Este proyecto todavía no tiene perforaciones.", st["celda"])] + [""] * 6)
    historia.append(_tabla(filas, anchos))

    # ── Avance por tipo de ensayo ────────────────────────────────────────────────────────────────────────────────────
    historia.append(Paragraph("Ensayos solicitados y realizados, por tipo", st["h2"]))
    historia.append(_tabla_tipos(d["por_tipo"], st))

    # ── Lo mismo, separado por perforación ───────────────────────────────────────────────────────────────────────────
    con_tipos = [pf for pf in d["perforaciones"] if pf.get("por_tipo")]
    if con_tipos:
        historia.append(Paragraph("Ensayos solicitados y realizados, por perforación y tipo", st["h2"]))
        for pf in con_tipos:
            cerrados = pf["finalizados"] + pf["no_realizados"]
            historia.append(KeepTogether([
                Paragraph(f"<b>{_e(pf['codigo'])}</b> · {_e(pf['tipo'])} — {cerrados} de {pf['solicitados']} ensayos cerrados "
                          f"({_pct(cerrados, pf['solicitados']):.0f} %)", st["base"]),
                Spacer(1, 3), _tabla_tipos(pf["por_tipo"], st), Spacer(1, 9)]))

    # ── Ensayos no realizados / pendientes y por qué ────────────────────────────────────────────────────────────────
    historia.append(Paragraph("Ensayos que no se realizaron o siguen pendientes, y por qué", st["h2"]))
    if d["no_hechos"]:
        enc = ["Perforación", "Muestra", "Ensayo", "Estado", "Motivo / observación"]
        filas = [[Paragraph(t, st["enc"]) for t in enc]]
        anchos = [2.2 * cm, 1.6 * cm, 3.6 * cm, 2.3 * cm, ANCHO_UTIL - 9.7 * cm]
        extra = []
        for i, r in enumerate(d["no_hechos"], start=1):
            _et, _f, suave = ESTADOS[r["estado"]]
            filas.append([Paragraph(_e(r["perforacion"]), st["celda"]), Paragraph(_e(r["muestra"]), st["celda"]),
                          Paragraph(_e(r["ensayo"]), st["celda_b"]), _chip(r["estado"], st), Paragraph(_e(r["motivo"]), st["celda"])])
            extra.append(("BACKGROUND", (3, i), (3, i), suave))
        historia.append(_tabla(filas, anchos, estilos_extra=extra))
    else:
        historia.append(Paragraph("Todos los ensayos solicitados se realizaron.", st["base"]))

    # ── Resultados para revisar (valores raros o fuera de lo esperado) ──────────────────────────────────────────────────
    historia.append(CondPageBreak(5 * cm))
    historia.append(Paragraph("Resultados para revisar", st["h2"]))
    alertas = d.get("alertas") or []
    n_alto = sum(1 for x in alertas if x["nivel"] == "alto")
    historia.append(Paragraph("Revisión automática de los resultados digitados: se comparan con los rangos típicos de la literatura geotécnica según el material de "
                              "cada muestra, y se cruzan la descripción visual, la clasificación y los límites. Una alerta pide una segunda mirada; no "
                              "significa que el ensayo esté mal ni lo invalida.", st["pequeno"]))
    historia.append(Spacer(1, 4))
    if alertas:
        historia.append(Paragraph(f"<b>{len(alertas)} resultado(s) para revisar</b> — {n_alto} de nivel alto y {len(alertas) - n_alto} por revisar.", st["base"]))
        historia.append(Spacer(1, 4))
        enc = ["Perforación", "Muestra", "Ensayo", "Nivel", "Qué revisar"]
        filas = [[Paragraph(t, st["enc"]) for t in enc]]
        anchos = [2.0 * cm, 1.5 * cm, 3.2 * cm, 1.8 * cm, ANCHO_UTIL - 8.5 * cm]
        extra = []
        for i, r in enumerate(alertas, start=1):
            alto = r["nivel"] == "alto"
            filas.append([Paragraph(_e(r["perforacion"]), st["celda"]), Paragraph(_e(r["muestra"]), st["celda"]),
                          Paragraph(_e(r["ensayo"]), st["celda_b"]),
                          Paragraph(f'<font color="#{"A51D1D" if alto else "8A4B06"}"><b>{_e(r["nivel_txt"])}</b></font>', st["celda_c"]),
                          Paragraph(_e(r["mensaje"]), st["celda"])])
            extra.append(("BACKGROUND", (3, i), (3, i), ROJO_CLARO if alto else AMBAR_CLARO))
        historia.append(_tabla(filas, anchos, estilos_extra=extra))
    else:
        caja = Table([[Paragraph('<font color="#0A5A2A"><b>No se detectaron resultados fuera de lo esperado</b> en los ensayos con datos de este '
                                 'proyecto.</font>', st["diag"])]], colWidths=[ANCHO_UTIL])
        caja.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, VERDE), ("BACKGROUND", (0, 0), (-1, -1), VERDE_CLARO2),
                                  ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6), ("LEFTPADDING", (0, 0), (-1, -1), 9)]))
        historia.append(caja)

    # ── Detalle por muestra ──────────────────────────────────────────────────────────────────────────────────────────
    historia.append(CondPageBreak(6 * cm))
    historia.append(Paragraph("Detalle de ensayos por muestra", st["h2"]))
    enc = ["Perforación", "Muestra", "Ensayo", "Estado", "Revisión", "Última actualización"]
    filas = [[Paragraph(t, st["enc"]) for t in enc]]
    anchos = [2.0 * cm, 1.5 * cm, 3.6 * cm, 2.2 * cm, ANCHO_UTIL - 12.4 * cm, 3.1 * cm]
    extra = []
    for i, r in enumerate(d["detalle"], start=1):
        _et, _f, suave = ESTADOS[r["estado"]]
        filas.append([Paragraph(_e(r["perforacion"]), st["celda"]), Paragraph(_e(r["muestra"]), st["celda"]),
                      Paragraph(_e(r["ensayo"]), st["celda_b"]), _chip(r["estado"], st), Paragraph(_e(r["revision"]), st["celda"]),
                      Paragraph(_e(r["actualizado"]), st["celda"])])
        extra.append(("BACKGROUND", (3, i), (3, i), suave))
    if len(filas) == 1:
        filas.append([Paragraph("No hay ensayos solicitados.", st["celda"])] + [""] * 5)
    historia.append(_tabla(filas, anchos, estilos_extra=extra))

    if d.get("observaciones"):
        historia.append(Paragraph("Observaciones registradas en las muestras", st["h2"]))
        filas = [[Paragraph(t, st["enc"]) for t in ("Perforación", "Muestra", "Observación")]]
        for r in d["observaciones"]:
            filas.append([Paragraph(_e(r["perforacion"]), st["celda"]), Paragraph(_e(r["muestra"]), st["celda"]), Paragraph(_e(r["texto"]), st["celda"])])
        historia.append(_tabla(filas, [2.2 * cm, 1.8 * cm, ANCHO_UTIL - 4.0 * cm]))

    # ── Armado ──────────────────────────────────────────────────────────────────────────────────────────────────────
    buf = BytesIO()
    titulo_cab = f"Informe de avance · {p['codigo']}"
    pie = f"{p['codigo']} · Generado el {d['generado']} (hora de Colombia)"

    def fabrica(*a, **k):
        c = _NumeradoCanvas(*a, pie=pie, **k)
        c._titulo_cabecera = titulo_cab
        return c

    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=MARGEN, rightMargin=MARGEN, topMargin=MARGEN, bottomMargin=1.7 * cm,
                          title=f"Informe de avance {p['codigo']}", author="Geodelta Ingeniería S.A.S.")
    primera = Frame(MARGEN, 1.7 * cm, ANCHO_UTIL, ALTO - MARGEN - 1.7 * cm, id="f1", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    resto = Frame(MARGEN, 1.7 * cm, ANCHO_UTIL, ALTO - 1.7 * cm - 1.6 * cm, id="f2", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="primera", frames=[primera], autoNextPageTemplate="resto"),
                          PageTemplate(id="resto", frames=[resto])])
    doc.build(historia, canvasmaker=fabrica)
    return buf.getvalue()
