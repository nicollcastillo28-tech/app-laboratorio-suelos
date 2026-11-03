"""Firmas en los formatos de Excel de resultados.

Cada formato trae un bloque «APROBÓ / REVISÓ» con los rótulos «Director Técnico» y «Jefe laboratorio» al pie. Este módulo, sobre el archivo YA generado:
  * pone las imágenes de la firma del Director Técnico y del Jefe de Laboratorio en el espacio vacío que hay entre «APROBÓ / REVISÓ» y el nombre,
  * escribe «MIGUEL BARÓN» en la casilla vacía que está justo encima de «Jefe laboratorio» (el nombre del Director Técnico ya viene impreso en el formato).
Trabaja directamente sobre el XML del libro (como el resto de las exportaciones): no pasa por openpyxl, que perdería imágenes y formas. Si algo no cuadra con el
formato de una hoja, esa hoja se deja como está: una firma nunca debe impedir una descarga.
"""
from io import BytesIO
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

CARPETA = Path(__file__).parent / 'templates' / 'firmas'
FIRMA_DIRECTOR = ('xl/media/gd_firma_director.png', CARPETA / 'firma_director_tecnico.png', 'Firma Director Técnico')
FIRMA_JEFE = ('xl/media/gd_firma_jefe.png', CARPETA / 'firma_jefe_laboratorio.png', 'Firma Jefe de Laboratorio')
NOMBRE_JEFE = 'MIGUEL BARÓN'

RE_DIRECTOR = re.compile(r'^\s*director\s+t', re.I)
RE_JEFE = re.compile(r'^\s*jefe\s+(de\s+)?laboratorio', re.I)
RE_APROBO = re.compile(r'aprob|revis', re.I)
EMU_PX = 9525
EMU_PT = 12700


def _tamano_png(ruta):
    from PIL import Image
    with Image.open(ruta) as im:
        return im.size


def _textos_compartidos(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    raiz = ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter(S + 't')) for si in raiz.iter(S + 'si')]


def _hojas(z):
    """[(nombre, ruta de la hoja)] en el orden del libro."""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).iter(REL + 'Relationship')}
    salida = []
    for h in wb.iter(S + 'sheet'):
        destino = rels.get(h.get(R + 'id'))
        if destino:
            ruta = destino.lstrip('/') if destino.startswith('/') else 'xl/' + destino
            salida.append((h.get('name'), ruta, h.get('state')))
    return salida


def _anchos_px(raiz):
    fmt = raiz.find(S + 'sheetFormatPr')
    base = float(fmt.get('defaultColWidth') or 8.43) if fmt is not None else 8.43
    por_columna = {}
    cols = raiz.find(S + 'cols')
    if cols is not None:
        for c in cols.iter(S + 'col'):
            ancho = float(c.get('width') or base)
            for i in range(int(c.get('min')), min(int(c.get('max')), 200) + 1):
                por_columna[i] = ancho
    return lambda i: int(por_columna.get(i, base) * 7 + 5)


def _altos_pt(raiz):
    fmt = raiz.find(S + 'sheetFormatPr')
    base = float(fmt.get('defaultRowHeight') or 15) if fmt is not None else 15
    por_fila = {}
    for f in raiz.iter(S + 'row'):
        if f.get('ht'):
            por_fila[int(f.get('r'))] = float(f.get('ht'))
    return lambda i: por_fila.get(i, base)


def _combinadas(raiz):
    salida = []
    mc = raiz.find(S + 'mergeCells')
    if mc is not None:
        for m in mc.iter(S + 'mergeCell'):
            a, b = m.get('ref').split(':')
            f1, c1 = coordinate_to_tuple(a)
            f2, c2 = coordinate_to_tuple(b)
            salida.append((c1, f1, c2, f2))
    return salida


def _celdas_con_texto(raiz, compartidos):
    """{(fila, columna): texto} de las celdas con texto."""
    salida = {}
    for c in raiz.iter(S + 'c'):
        tipo = c.get('t')
        texto = None
        if tipo == 's':
            v = c.find(S + 'v')
            if v is not None and v.text is not None and v.text.isdigit() and int(v.text) < len(compartidos):
                texto = compartidos[int(v.text)]
        elif tipo == 'inlineStr':
            texto = ''.join(t.text or '' for t in c.iter(S + 't'))
        elif tipo == 'str':
            v = c.find(S + 'v')
            texto = v.text if v is not None else None
        if texto and texto.strip():
            fila, col = coordinate_to_tuple(c.get('r'))
            salida[(fila, col)] = texto
    return salida


def _rango_de(col, fila, combinadas):
    for c1, f1, c2, f2 in combinadas:
        if c1 <= col <= c2 and f1 <= fila <= f2:
            return c1, f1, c2, f2
    return col, fila, col, fila


def _posicion(valor, tamanos):
    """(índice base 1, desplazamiento en px dentro de ese índice) de la posición `valor` (px desde el borde), con `tamanos(i)` el tamaño del índice i."""
    i, acumulado = 1, 0
    while True:
        t = tamanos(i)
        if valor < acumulado + t or i > 2000:
            return i, max(0, valor - acumulado)
        acumulado += t
        i += 1


def _disenar(raiz, compartidos):
    """Para una hoja: [(archivo_firma, (col, off_x, fila, off_y, ancho_px, alto_px))] y [(ref, texto)] a escribir. Vacío si la hoja no tiene el bloque."""
    textos = _celdas_con_texto(raiz, compartidos)
    # cada rótulo se atiende por su cuenta: algunos formatos (Proctor) tienen más de un bloque, o los dos rótulos en filas distintas
    rotulos = []
    for (fila, col), t in sorted(textos.items()):
        if RE_DIRECTOR.match(t):
            rotulos.append(((fila, col), FIRMA_DIRECTOR))
        elif RE_JEFE.match(t):
            rotulos.append(((fila, col), FIRMA_JEFE))
    if not rotulos:
        return [], []
    combinadas = _combinadas(raiz)
    ancho, alto = _anchos_px(raiz), _altos_pt(raiz)
    imagenes, escribir = [], []
    for (fila, col), firma in rotulos:
        fila_rotulo = fila
        c1, _f1, c2, _f2 = _rango_de(col, fila, combinadas)
        # fila del «APROBÓ / REVISÓ» de este bloque: la más cercana hacia arriba (hasta 9 filas)
        fila_aprobo = next((f for f in range(fila_rotulo - 1, fila_rotulo - 10, -1)
                            if f > 0 and any(RE_APROBO.search(textos.get((f, c), '')) for c in range(c1, c2 + 1))), fila_rotulo - 5)
        f_ini, f_fin = fila_aprobo + 1, fila_rotulo - 2   # entre «APROBÓ / REVISÓ» y la fila del nombre
        if f_fin < f_ini:
            continue
        # no se pisa nada: ninguna celda de ese espacio puede tener texto
        if any(textos.get((f, c)) for f in range(f_ini, f_fin + 1) for c in range(c1, c2 + 1)):
            continue
        bloque_x = sum(ancho(i) for i in range(1, c1))
        bloque_y = sum(alto(f) * 96 / 72 for f in range(1, f_ini))
        W = sum(ancho(i) for i in range(c1, c2 + 1))
        H = sum(alto(f) * 96 / 72 for f in range(f_ini, f_fin + 1))
        iw0, ih0 = _tamano_png(firma[1])
        escala = min(W * 0.92 / iw0, H * 0.92 / ih0)
        iw, ih = max(1, int(iw0 * escala)), max(1, int(ih0 * escala))
        x = bloque_x + (W - iw) / 2
        y = bloque_y + (H - ih) / 2
        col_i, off_x = _posicion(x, ancho)
        fila_i, off_y = _posicion(y, lambda f: alto(f) * 96 / 72)
        imagenes.append((firma, (col_i - 1, int(off_x * EMU_PX), fila_i - 1, int(off_y * EMU_PX), iw, ih)))
        if firma is FIRMA_JEFE:
            # nombre del Jefe: casilla vacía justo encima del rótulo (no se pisa lo que ya haya)
            fila_nombre = fila - 1
            origen = _rango_de(c1, fila_nombre, combinadas)
            if fila_nombre > 0 and not textos.get((fila_nombre, origen[0])) and origen[1] == fila_nombre:
                escribir.append((f'{get_column_letter(origen[0])}{fila_nombre}', NOMBRE_JEFE))
    return imagenes, escribir


def _poner_texto(xml, ref, texto):
    """Escribe `texto` en la celda `ref` conservando su estilo. False si la celda no existe en el XML (se deja como está)."""
    patron = re.compile(r'<c r="%s"((?:\s[^>]*?)?)(/>|>.*?</c>)' % re.escape(ref), re.S)
    m = patron.search(xml)
    if not m:
        return xml, False
    atributos = re.sub(r'\st="[^"]*"', '', m.group(1))
    seguro = texto.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    nueva = f'<c r="{ref}"{atributos} t="inlineStr"><is><t>{seguro}</t></is></c>'
    return xml[:m.start()] + nueva + xml[m.end():], True


def agregar_firmas(contenido):
    """Devuelve el libro con las firmas y el nombre del Jefe de Laboratorio puestos. Idempotente: si ya las tiene, lo devuelve igual."""
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        nombres = z.namelist()
        if FIRMA_JEFE[0] in nombres:
            return contenido
        compartidos = _textos_compartidos(z)
        cambios = {}   # ruta -> bytes nuevos
        usadas = set()
        for _nombre, ruta, estado in _hojas(z):
            if ruta not in nombres or estado == 'hidden':
                continue
            xml_bytes = z.read(ruta)
            try:
                raiz = ET.fromstring(xml_bytes)
                imagenes, escribir = _disenar(raiz, compartidos)
            except Exception:
                continue
            if not imagenes and not escribir:
                continue
            xml = xml_bytes.decode('utf-8')
            for ref, texto in escribir:
                xml, _ = _poner_texto(xml, ref, texto)
            if imagenes:
                m_dib = re.search(r'<drawing\s[^>]*?r:id="([^"]+)"', xml)
                rels_hoja = ruta.replace('worksheets/', 'worksheets/_rels/') + '.rels'
                if m_dib and rels_hoja in nombres:
                    destino = next((r.get('Target') for r in ET.fromstring(z.read(rels_hoja)).iter(REL + 'Relationship') if r.get('Id') == m_dib.group(1)), None)
                    if destino:
                        dib = 'xl/drawings/' + destino.split('/')[-1] if destino.startswith('..') else destino.lstrip('/')
                        rels_dib = dib.replace('drawings/', 'drawings/_rels/') + '.rels'
                        if dib in nombres:
                            dib_xml = cambios.get(dib, z.read(dib)).decode('utf-8')
                            rels_xml = cambios.get(rels_dib, z.read(rels_dib) if rels_dib in nombres else
                                                   b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode('utf-8')
                            siguiente = max([int(i) for i in re.findall(r'<(?:\w+:)?cNvPr[^>]*\sid="(\d+)"', dib_xml)] + [1]) + 1
                            anclas = ''
                            for firma, (col, off_x, fila, off_y, iw, ih) in imagenes:
                                rid = f'rIdGdFirma{len(usadas) + 1}{siguiente}'
                                anclas += (
                                    f'<xdr:oneCellAnchor xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
                                    f'<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{off_x}</xdr:colOff><xdr:row>{fila}</xdr:row><xdr:rowOff>{off_y}</xdr:rowOff></xdr:from>'
                                    f'<xdr:ext cx="{iw * EMU_PX}" cy="{ih * EMU_PX}"/>'
                                    f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{siguiente}" name="{firma[2]}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                                    f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                                    f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{iw * EMU_PX}" cy="{ih * EMU_PX}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
                                    f'</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')
                                rels_xml = rels_xml.replace('</Relationships>', f'<Relationship Id="{rid}" Type="{NS_R}/image" Target="../media/{Path(firma[0]).name}"/></Relationships>')
                                usadas.add(firma[0])
                                siguiente += 1
                            cierre = re.search(r'</(?:\w+:)?wsDr>', dib_xml)
                            if cierre:
                                cambios[dib] = (dib_xml[:cierre.start()] + anclas + dib_xml[cierre.start():]).encode('utf-8')
                                cambios[rels_dib] = rels_xml.encode('utf-8')
            cambios[ruta] = xml.encode('utf-8')
        if not usadas and not cambios:
            return contenido
        # tipo de contenido para .png (si el libro todavía no lo declara)
        tipos = z.read('[Content_Types].xml').decode('utf-8')
        if usadas and not re.search(r'<Default[^>]*Extension="png"', tipos, re.I):
            tipos = tipos.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override', 1)
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                if item.filename == '[Content_Types].xml':
                    out.writestr(item, tipos.encode('utf-8'))
                elif item.filename in cambios:
                    out.writestr(item, cambios[item.filename])
                else:
                    out.writestr(item, z.read(item.filename))
            for nuevo, bytes_ in cambios.items():
                if nuevo not in nombres:
                    out.writestr(nuevo, bytes_)
            for ruta_media, origen, _n in (FIRMA_DIRECTOR, FIRMA_JEFE):
                if ruta_media in usadas:
                    out.writestr(ruta_media, origen.read_bytes())
        return salida.getvalue()


def quitar_subrayado(contenido, ruta_hoja, ref):
    """Quita el subrayado de la letra de la celda `ref` (por ejemplo la descripción visual, que el formato trae subrayada) sin tocar el estilo de las demás celdas:
    se crea un estilo nuevo, copia del de la celda pero con una letra sin subrayar. Si la celda ya no está subrayada (o no se encuentra), devuelve el libro igual."""
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        if ruta_hoja not in z.namelist() or 'xl/styles.xml' not in z.namelist():
            return contenido
        hoja = z.read(ruta_hoja).decode('utf-8')
        estilos = z.read('xl/styles.xml').decode('utf-8')
        m_celda = re.search(r'<c r="%s"((?:\s[^>]*?)?)(/>|>)' % re.escape(ref), hoja)
        if not m_celda:
            return contenido
        m_s = re.search(r'\ss="(\d+)"', m_celda.group(1))
        idx = int(m_s.group(1)) if m_s else 0
        m_xfs = re.search(r'(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)', estilos, re.S)
        m_fonts = re.search(r'(<fonts\b[^>]*>)(.*?)(</fonts>)', estilos, re.S)
        if not m_xfs or not m_fonts:
            return contenido
        xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', m_xfs.group(2), re.S)
        fuentes = re.findall(r'<font\b[^>]*?(?:/>|>.*?</font>)', m_fonts.group(2), re.S)
        if idx >= len(xfs):
            return contenido
        xf = xfs[idx]
        m_fid = re.search(r'\sfontId="(\d+)"', xf)
        fid = int(m_fid.group(1)) if m_fid else 0
        if fid >= len(fuentes) or not re.search(r'<u\b', fuentes[fid]):
            return contenido
        fuente_nueva = re.sub(r'<u\b[^>]*/>|<u\b[^>]*>.*?</u>', '', fuentes[fid], flags=re.S)
        n_fuentes, n_xfs = len(fuentes), len(xfs)
        xf_nuevo = re.sub(r'\sfontId="\d+"', f' fontId="{n_fuentes}"', xf, count=1) if m_fid else xf.replace('<xf ', f'<xf fontId="{n_fuentes}" ', 1)
        if 'applyFont=' in xf_nuevo:
            xf_nuevo = re.sub(r'applyFont="\d"', 'applyFont="1"', xf_nuevo)
        else:
            xf_nuevo = xf_nuevo.replace('<xf ', '<xf applyFont="1" ', 1)
        estilos = (estilos[:m_fonts.start(3)] + fuente_nueva + estilos[m_fonts.start(3):])
        # los índices de abajo cambian al insertar la fuente: se vuelve a buscar el bloque de estilos de celda
        m_xfs = re.search(r'(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)', estilos, re.S)
        estilos = estilos[:m_xfs.start(3)] + xf_nuevo + estilos[m_xfs.start(3):]
        estilos = re.sub(r'(<fonts\b[^>]*?\scount=")\d+(")', r'\g<1>%d\g<2>' % (n_fuentes + 1), estilos, count=1)
        estilos = re.sub(r'(<cellXfs\b[^>]*?\scount=")\d+(")', r'\g<1>%d\g<2>' % (n_xfs + 1), estilos, count=1)
        atributos = m_celda.group(1)
        atributos = re.sub(r'\ss="\d+"', f' s="{n_xfs}"', atributos) if m_s else atributos + f' s="{n_xfs}"'
        hoja = hoja[:m_celda.start()] + f'<c r="{ref}"{atributos}{m_celda.group(2)}' + hoja[m_celda.end():]
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                if item.filename == ruta_hoja:
                    out.writestr(item, hoja.encode('utf-8'))
                elif item.filename == 'xl/styles.xml':
                    out.writestr(item, estilos.encode('utf-8'))
                else:
                    out.writestr(item, z.read(item.filename))
        return salida.getvalue()
