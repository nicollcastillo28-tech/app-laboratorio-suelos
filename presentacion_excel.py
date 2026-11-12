"""Ajustes de presentación únicamente en los libros que se descargan."""
from io import BytesIO
import re
import zipfile
from openpyxl.utils.cell import range_boundaries, get_column_letter


def _rango(rect):
    c1, r1, c2, r2 = rect
    a, b = f'{get_column_letter(c1)}{r1}', f'{get_column_letter(c2)}{r2}'
    return a if a == b else f'{a}:{b}'


def _interseccion(a, b):
    c1, r1, c2, r2 = max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])
    return (c1, r1, c2, r2) if c1 <= c2 and r1 <= r2 else None


def _restar(rect, parte):
    c1, r1, c2, r2 = rect
    x1, y1, x2, y2 = parte
    return [r for r in [(c1,r1,c2,y1-1),(c1,y2+1,c2,r2),(c1,y1,x1-1,y2),(x2+1,y1,c2,y2)]
            if r[0] <= r[2] and r[1] <= r[3]]


def corregir_avisos_vacios(xml):
    """Las reglas de casillas vacías consultan su celda, o el ancla de su combinación.

    Separar los rangos evita la referencia desplazada de las reglas de varias áreas.
    Se mantienen las reglas de resultados, colores, cobertura y orden de precedencia.
    """
    merges = [range_boundaries(r) for r in re.findall(r'<mergeCell\b[^>]*\bref="([^"]+)"', xml)]
    grupos = re.compile(r'<conditionalFormatting\b([^>]*)>(.*?)</conditionalFormatting>', re.S)
    reglas = re.compile(r'<cfRule\b([^>]*?)(?:/>|>(.*?)</cfRule>)', re.S)
    def grupo(m):
        attrs, cuerpo = m.group(1), m.group(2)
        sr = re.search(r'\bsqref="([^"]+)"', attrs)
        if not sr:
            return m.group(0)
        rangos = sr.group(1).split()
        salida = []
        for regla in reglas.finditer(cuerpo):
            ra, contenido = regla.group(1), regla.group(2) or ''
            formula = re.search(r'<formula>(.*?)</formula>', contenido, re.S)
            simple = formula and re.fullmatch(r'LEN\(TRIM\(\$?[A-Z]+\$?\d+\)\)=0', formula.group(1))
            propia = formula and formula.group(1) == 'LEN(TRIM(INDIRECT(ADDRESS(ROW(),COLUMN()))))=0'
            if not (simple or propia) or not re.search(r'\btype="(?:containsBlanks|expression)"', ra):
                salida.append(f'<conditionalFormatting{attrs}>{regla.group(0)}</conditionalFormatting>')
                continue
            destinos = []
            for rango in rangos:
                pendientes = [range_boundaries(rango)]
                for combinado in merges:
                    resto = []
                    for rect in pendientes:
                        parte = _interseccion(rect, combinado)
                        if parte:
                            ancla = f'${get_column_letter(combinado[0])}${combinado[1]}'
                            destinos.append((parte, ancla))
                            resto.extend(_restar(rect, parte))
                        else:
                            resto.append(rect)
                    pendientes = resto
                destinos.extend((rect, f'{get_column_letter(rect[0])}{rect[1]}') for rect in pendientes)
            ra = re.sub(r'\btype="[^"]+"', 'type="expression"', ra)
            for rect, ancla in dict.fromkeys(destinos):
                nuevo = contenido[:formula.start()] + f'<formula>LEN(TRIM({ancla}))=0</formula>' + contenido[formula.end():]
                nuevos_attrs = re.sub(r'\bsqref="[^"]+"', f'sqref="{_rango(rect)}"', attrs)
                salida.append(f'<conditionalFormatting{nuevos_attrs}><cfRule{ra}>{nuevo}</cfRule></conditionalFormatting>')
        return ''.join(salida) if salida else m.group(0)
    resultado = grupos.sub(grupo, xml)
    if resultado == xml:
        return xml
    # Las reglas divididas deben tener prioridades únicas, conservando su orden relativo.
    encontrados = list(re.finditer(r'<cfRule\b[^>]*\bpriority="(\d+)"[^>]*>', resultado))
    orden = {m.start(): i+1 for i,m in enumerate(sorted(encontrados,key=lambda m:(int(m.group(1)),m.start())))}
    return re.sub(r'<cfRule\b[^>]*\bpriority="\d+"[^>]*>',
                  lambda m: re.sub(r'\bpriority="\d+"', f'priority="{orden[m.start()]}"', m.group(0)), resultado)


TAMANO_ENCABEZADO = "11"


class _EstilosPorContenido(dict):
    """Índice de estilos (<xf>) que reconoce como iguales dos estilos con los mismos atributos en distinto orden o escritos como <xf .../> y
    <xf ...></xf>. Sin esto, ajustar dos veces el mismo libro creaba copias repetidas de un estilo que ya estaba."""
    @staticmethod
    def _clave(xf):
        m = re.match(r'<xf\b([^>]*?)(/?)>', xf)
        cuerpo = '' if m.group(2) else xf[m.end():-len('</xf>')]
        return tuple(sorted(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))), cuerpo

    def __init__(self, pares=()):
        super().__init__()
        for xf, i in pares:
            self.setdefault(self._clave(xf), i)

    def __contains__(self, xf):
        return super().__contains__(self._clave(xf))

    def __getitem__(self, xf):
        return super().__getitem__(self._clave(xf))

    def __setitem__(self, xf, i):
        super().__setitem__(self._clave(xf), i)


def _fila_descripcion_visual(z, xml):
    """Fila del rótulo «DESCRIPCIÓN VISUAL» de la hoja (el final del encabezado de los informes), o None si la hoja no lo tiene."""
    if 'xl/sharedStrings.xml' not in z.namelist():
        return None
    compartidos = z.read('xl/sharedStrings.xml')
    guardado = getattr(z, '_indices_descripcion', None)   # en un PaqueteZip se calcula una vez para todas sus hojas
    if guardado is not None and guardado[0] is compartidos:
        indices = guardado[1]
    else:
        textos = [''.join(re.findall(r'<t\b[^>]*>([^<]*)</t>', si)) for si in
                  re.findall(r'<si>(.*?)</si>', compartidos.decode('utf-8'), re.S)]
        indices = {str(i) for i, t in enumerate(textos) if re.match(r'\s*DESCRIPCI', t, re.I)}
        if isinstance(z, PaqueteZip):
            z._indices_descripcion = (compartidos, indices)
    filas = [int(m.group(1)) for m in re.finditer(r'<c\b[^>]*\br="[A-Z]+(\d+)"[^>]*\bt="s"[^>]*><v>(\d+)</v>', xml) if m.group(2) in indices]
    return min(filas) if filas else None


class PaqueteZip:
    """El .xlsx descomprimido en memoria, para aplicarle varias transformaciones seguidas sin volver a comprimir el archivo entero entre una y
    otra (en Consolidación, con sus 15 MB de XML, eso hacía que la descarga tardara minutos). Imita lo que se usa de ZipFile (namelist, read)."""

    def __init__(self, contenido):
        with zipfile.ZipFile(BytesIO(contenido)) as z:
            self.infos = z.infolist()
            self.datos = {i.filename: z.read(i.filename) for i in self.infos}

    def namelist(self):
        return [i.filename for i in self.infos if i.filename in self.datos]

    def read(self, nombre):
        return self.datos[nombre]

    def escribir(self, nombre, datos):
        self.datos[nombre] = datos

    def quitar(self, nombre):
        self.datos.pop(nombre, None)

    def bytes(self):
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in self.infos:
                if item.filename in self.datos:
                    out.writestr(item, self.datos[item.filename])
        return salida.getvalue()


def ajustar_presentacion(contenido, parte, referencias):
    """Centra las entradas sin modificar valores, formatos numéricos ni objetos. Además, lo que la app escribe en el ENCABEZADO del informe
    (de la fila 1 a la de «DESCRIPCIÓN VISUAL») va siempre en letra 11: varias plantillas traen algunas de esas casillas en 10 (el proyecto en
    Humedad, la perforación en la 001 de antes…) y al descargar se veían más pequeñas. El estilo original no se toca: se usa una copia del mismo
    estilo con la letra en 11."""
    paquete = PaqueteZip(contenido)
    ajustar_presentacion_en(paquete, parte, referencias)
    return paquete.bytes()


def ajustar_presentacion_en(z, parte, referencias):
    """ajustar_presentacion sobre un PaqueteZip (lo modifica)."""
    xml = z.read(parte).decode('utf-8')
    estilos = z.read('xl/styles.xml').decode('utf-8')
    bloque = re.search(r'<cellXfs\b[^>]*>(.*?)</cellXfs>', estilos, re.S)
    xfs = re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*>.*?</xf>', bloque.group(1), re.S)
    existentes = _EstilosPorContenido((xf, i) for i, xf in enumerate(xfs))
    indices = {}
    bloque_f =re.search(r'<fonts\b[^>]*>(.*?)</fonts>', estilos, re.S)
    fuentes = re.findall(r'<font\b[^>]*?/>|<font\b[^>]*>.*?</font>', bloque_f.group(1), re.S) if bloque_f else []
    n_fuentes = len(fuentes)
    fuentes_11, indices_11 = {}, {}

    def fuente_11(fid):
        """Índice de una fuente igual a `fid` pero de tamaño 11 (la misma si ya lo es)."""
        if fid in fuentes_11:
            return fuentes_11[fid]
        f = fuentes[fid]
        if re.search(r'<sz\b[^>]*\bval="' + TAMANO_ENCABEZADO + r'"', f):
            fuentes_11[fid] = fid
            return fid
        if f.endswith('/>'):
            nueva = f'<font><sz val="{TAMANO_ENCABEZADO}"/></font>'
        elif re.search(r'<sz\b[^>]*/>', f):
            nueva = re.sub(r'<sz\b[^>]*/>', f'<sz val="{TAMANO_ENCABEZADO}"/>', f, count=1)
        else:
            apertura_f = re.match(r'<font\b[^>]*>', f).group(0)
            nueva = apertura_f + f'<sz val="{TAMANO_ENCABEZADO}"/>' + f[len(apertura_f):]
        if nueva not in fuentes:
            fuentes.append(nueva)
        fuentes_11[fid] = fuentes.index(nueva)
        return fuentes_11[fid]

    def en_11(indice):
        """Índice de un estilo igual a `indice` pero con la letra en 11 (el mismo si ya lo es)."""
        if indice in indices_11:
            return indices_11[indice]
        xf = xfs[indice]
        tag = re.match(r'<xf\b[^>]*>', xf).group(0)
        fid = re.search(r'\bfontId="(\d+)"', tag)
        fid = int(fid.group(1)) if fid else 0
        if fid >= len(fuentes) or fuente_11(fid) == fid:
            indices_11[indice] = indice
            return indice
        tag_nuevo = re.sub(r'\s+applyFont="[^"]*"', '', tag)
        if re.search(r'\bfontId="\d+"', tag_nuevo):
            tag_nuevo = re.sub(r'\bfontId="\d+"', f'fontId="{fuente_11(fid)}"', tag_nuevo)
        else:
            tag_nuevo = tag_nuevo.replace('<xf', f'<xf fontId="{fuente_11(fid)}"', 1)
        tag_nuevo = tag_nuevo[:-2] + ' applyFont="1"/>' if tag_nuevo.endswith('/>') else tag_nuevo[:-1] + ' applyFont="1">'
        nuevo = tag_nuevo + xf[len(tag):]
        if nuevo not in existentes:
            existentes[nuevo] = len(xfs)
            xfs.append(nuevo)
        indices_11[indice] = existentes[nuevo]
        return indices_11[indice]

    def centrado(indice):
        if indice in indices:
            return indices[indice]
        xf = xfs[indice]
        tag = re.match(r'<xf\b[^>]*>', xf).group(0)
        attrs = re.sub(r'\s+applyAlignment="[^"]*"', '', tag).rstrip('>/')
        tag_nuevo = attrs + ' applyAlignment="1">'
        cuerpo = '' if tag.endswith('/>') else xf[len(tag):-5]
        al = re.search(r'<alignment\b[^>]*?/>', cuerpo)
        alineacion = al.group(0) if al else '<alignment/>'
        alineacion = re.sub(r'\s+(?:horizontal|vertical|indent|relativeIndent)="[^"]*"', '', alineacion)
        alineacion = alineacion[:-2] + ' horizontal="center" vertical="center"/>'
        if al:
            cuerpo = cuerpo[:al.start()] + alineacion + cuerpo[al.end():]
        else:
            # alignment precede protection y extLst en CT_Xf.
            posicion = re.search(r'<(?:protection|extLst)\b', cuerpo)
            i = posicion.start() if posicion else len(cuerpo)
            cuerpo = cuerpo[:i] + alineacion + cuerpo[i:]
        nuevo = tag_nuevo + cuerpo + '</xf>'
        if nuevo not in existentes:
            existentes[nuevo] = len(xfs)
            xfs.append(nuevo)
        indices[indice] = existentes[nuevo]
        return indices[indice]
    # Identificadores estables: dos muestras del mismo formato comparten estilos
    # al agruparse, aunque una tenga campos vacíos o límites NP/NL.
    n_base = len(xfs)
    for indice in range(n_base):
        centrado(indice)
    # Las versiones con letra 11 también se crean siempre en el mismo orden, por la misma razón.
    for indice in range(n_base):
        centrado(en_11(indice))
    refs = set(referencias)
    fin_encabezado = _fila_descripcion_visual(z, xml)
    def celda(m):
        if m.group(1) not in refs:
            return m.group(0)
        tag = m.group(0)
        s = re.search(r'\bs="(\d+)"', tag)
        base = int(s.group(1)) if s else 0
        if fin_encabezado and int(re.sub(r'^[A-Z]+', '', m.group(1))) <= fin_encabezado:
            base = en_11(base)   # encabezado del informe: siempre en letra 11
        nuevo = str(centrado(base))
        if s:
            return re.sub(r'\bs="\d+"', f's="{nuevo}"', tag)
        return tag[:-2] + f' s="{nuevo}"/>' if tag.endswith('/>') else tag[:-1] + f' s="{nuevo}">'
    xml = re.sub(r'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>', celda, xml)
    xml = corregir_avisos_vacios(xml)
    if len(xfs) != len(re.findall(r'<xf\b', bloque.group(1))):
        apertura = re.sub(r'\bcount="\d+"', f'count="{len(xfs)}"', estilos[bloque.start():bloque.start(1)])
        estilos = estilos[:bloque.start()] + apertura + ''.join(xfs) + '</cellXfs>' + estilos[bloque.end():]
    if len(fuentes) != n_fuentes:   # <fonts> va antes de <cellXfs>: se busca de nuevo, ya con cellXfs reemplazado
        bloque_f = re.search(r'<fonts\b[^>]*>(.*?)</fonts>', estilos, re.S)
        apertura = re.sub(r'\bcount="\d+"', f'count="{len(fuentes)}"', estilos[bloque_f.start():bloque_f.start(1)])
        estilos = estilos[:bloque_f.start()] + apertura + ''.join(fuentes) + '</fonts>' + estilos[bloque_f.end():]
    z.escribir(parte, xml.encode('utf-8'))
    z.escribir('xl/styles.xml', estilos.encode('utf-8'))


def reparar_formulas_clasificacion(contenido, parte):
    """Corrige errores de texto en límites y CONCAT sin cambiar el diseño ni el método."""
    # (El código del «INFORME ID» —K10 y Y10— NO se toca: va tal cual lo trae la plantilla oficial, que es el código estandarizado del laboratorio.)
    formulas = {
        'R40': 'IF(COUNT(S19:U19)=0,"",IF(T19=25,IF(ISNUMBER(T25),ROUNDDOWN(T25,0),""),IF(COUNT(S25:U25)=0,"",ROUNDDOWN(Q52,0))))',
        'R41': 'IF(COUNT(Y25:Z25)=0,"",ROUNDDOWN(AVERAGE(Y25:Z25),0))',
        'R42': 'IF(R40="NO LÍQUIDO","N.L. - N.P.",IF(R41="NO PLÁSTICO","N.P.",IF(AND(ISNUMBER(R40),ISNUMBER(R41)),R40-R41,"")))',
        'U41': 'IFERROR(CONCATENATE(AE41," (",IF(AE41="A-2-6 O A-2-7",AF43,AE43),")"),"-")',
    }
    import html
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        xml = z.read(parte).decode('utf-8')
        for ref, formula in formulas.items():
            patron = re.compile(r'(<c\b[^>]*\br="'+ref+r'"[^>]*>)(.*?)(</c>)', re.S)
            def cambiar(m):
                if not re.search(r'<f\b', m.group(2)):
                    return m.group(0)  # NP/NL ya digitado tiene precedencia.
                cuerpo = re.sub(r'<f\b[^>]*>.*?</f>', '<f>'+html.escape(formula,quote=False)+'</f>', m.group(2), flags=re.S)
                cuerpo = re.sub(r'<v(?:\s[^>]*)?>.*?</v>|<v\s*/>', '', cuerpo, flags=re.S)
                return m.group(1)+cuerpo+m.group(3)
            xml, n = patron.subn(cambiar, xml)
            if n != 1:
                raise ValueError(f'La celda {ref} del formato de clasificación cambió')
        salida=BytesIO()
        with zipfile.ZipFile(salida,'w',zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item,xml.encode('utf-8') if item.filename==parte else z.read(item.filename))
        return salida.getvalue()


def reparar_punto_carta_plasticidad(contenido, ll=None, ip=None):
    """El punto de la muestra en la CARTA DE PLASTICIDAD (el triángulo verde) debe ir en X = Límite Líquido (R40), Y = Índice de Plasticidad (R42).
    La plantilla oficial GDA-FLC-001 lo traía mal: primero X e Y en FORMATO!$AI$29 (un 0: salía en la esquina) y luego Y en R41, que es el
    Límite Plástico (con LL 65 / LP 21 / IP 44 caía en MH en vez de CH). Aquí se busca en la carta la serie de un solo punto (X e Y de una sola
    celda, no un rango como las líneas A, U y las zonas) y se apunta a R40 / R42. Si `ll` e `ip` son números, también queda calculado el valor
    en el gráfico (los visores que no recalculan, como la vista previa del celular, ya lo ven en su sitio); con N.L. / N.P. no hay punto.
    El resto del archivo queda byte por byte igual."""
    def num(v):
        return isinstance(v, (int, float)) and not isinstance(v, bool)

    def cache(valor):
        if not num(valor):
            return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/></c:numCache>'
        return f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>{valor}</c:v></c:pt></c:numCache>'

    def eje(etiqueta, ref, valor):
        return f'<c:{etiqueta}><c:numRef><c:f>{ref}</c:f>{cache(valor)}</c:numRef></c:{etiqueta}>'

    hay_punto = num(ll) and num(ip)
    un_punto = r"<c:{0}><c:numRef><c:f>'?FORMATO'?!\$?[A-Z]+\$?\d+</c:f>.*?</c:numRef></c:{0}>"
    cambios = {}
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        for nombre in z.namelist():
            if not re.match(r'xl/charts/chart\d+\.xml$', nombre):
                continue
            xml = z.read(nombre).decode('utf-8')
            if 'CARTA DE PLASTICIDAD' not in xml:
                continue

            def arreglar(m):
                serie = m.group(0)
                if not (re.search(un_punto.format('xVal'), serie, re.S) and re.search(un_punto.format('yVal'), serie, re.S)):
                    return serie
                serie = re.sub(un_punto.format('xVal'), lambda _m: eje('xVal', 'FORMATO!$R$40', ll if hay_punto else None), serie, flags=re.S)
                return re.sub(un_punto.format('yVal'), lambda _m: eje('yVal', 'FORMATO!$R$42', ip if hay_punto else None), serie, flags=re.S)
            nuevo = re.sub(r'<c:ser>.*?</c:ser>', arreglar, xml, flags=re.S)
            if nuevo != xml:
                cambios[nombre] = nuevo.encode('utf-8')
        if not cambios:
            return contenido
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, cambios.get(item.filename) or z.read(item.filename))
        return salida.getvalue()


def reponer_formulas_faltantes(contenido, parte, formulas):
    """Pone en la hoja `parte` las fórmulas de `formulas` ({celda: fórmula sin «=»}) SOLO en las celdas que no traen fórmula ni valor.
    Para cálculos que la plantilla oficial perdió (p. ej. la masa seca y la humedad del GDA-FLC-014): si la plantilla ya los trae, no se toca
    nada. Se conserva el estilo de la celda; Excel calcula el resultado al abrir (el libro ya va con recálculo completo)."""
    import html
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        xml = z.read(parte).decode('utf-8')
        nuevo = xml
        for ref, formula in formulas.items():
            m = re.search(r'<c\b[^>]*\br="' + ref + r'"[^>]*?(?:/>|>(.*?)</c>)', nuevo, re.S)
            if not m or (m.group(1) and re.search(r'<f\b|<v>[^<]', m.group(1))):
                continue
            abre = re.match(r'<c\b[^>]*?(?=/?>)', m.group(0)).group(0)
            abre = re.sub(r'\s+t="[^"]*"', '', abre)
            nuevo = nuevo[:m.start()] + f'{abre}><f>{html.escape(formula, quote=False)}</f></c>' + nuevo[m.end():]
        if nuevo == xml:
            return contenido
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, nuevo.encode('utf-8') if item.filename == parte else z.read(item.filename))
        return salida.getvalue()


def cambiar_formulas_si_iguales(contenido, parte, cambios):
    """Cambia fórmulas de la hoja `parte` SOLO si todavía son las de antes: {celda: (fórmula vieja, fórmula nueva)}, sin «=». Para corregir en la
    descarga una fórmula de una plantilla que tiene la hoja protegida con contraseña (p. ej. el ×10 del kN/m³ del peso unitario parafinado):
    si la plantilla oficial ya viene corregida, no se toca nada. Se quita el valor guardado de la celda; Excel la recalcula al abrir."""
    return cambiar_formulas_varias(contenido, {parte: cambios})


def cambiar_formulas_varias(contenido, cambios_por_parte):
    """cambiar_formulas_si_iguales en varias hojas a la vez ({parte: cambios}), descomprimiendo y comprimiendo el archivo una sola vez.
    Devuelve el mismo contenido si nada cambió."""
    paquete = None
    for parte, cambios in cambios_por_parte.items():
        paquete = paquete or PaqueteZip(contenido)
        cambiar_formulas_en(paquete, parte, cambios)
    if paquete is None or not getattr(paquete, '_cambio', False):
        return contenido
    return paquete.bytes()


def cambiar_formulas_en(z, parte, cambios):
    """cambiar_formulas_si_iguales sobre un PaqueteZip (lo modifica). True si cambió alguna fórmula."""
    import html
    if parte not in z.datos:
        return False
    xml = z.read(parte).decode('utf-8')
    nuevo = xml
    for ref, (vieja, nueva) in cambios.items():
        m = re.search(r'(<c\b[^>]*\br="' + ref + r'"[^>]*>)(.*?)(</c>)', nuevo, re.S)
        if not m:
            continue
        f = re.search(r'(<f\b[^>]*>)(.*?)</f>', m.group(2), re.S)
        if not f or html.unescape(f.group(2)).replace(' ', '') != vieja.replace(' ', ''):
            continue   # (una celda de fórmula COMPARTIDA que no es la principal no trae texto: cambia sola al cambiar la principal)
        # Se conserva la etiqueta <f ...> tal cual: si es la principal de una fórmula compartida (t="shared" ref=... si=...), las demás
        # celdas del grupo siguen apuntando a ella y toman la fórmula nueva (G27 = G25*9.81 al cambiar G26 = G24*9.81).
        apertura = re.sub(r'\s+t="[^"]*"', '', m.group(1))
        nuevo = nuevo[:m.start()] + f'{apertura}{f.group(1)}{html.escape(nueva, quote=False)}</f>{m.group(3)}' + nuevo[m.end():]
    if nuevo == xml:
        return False
    z.escribir(parte, nuevo.encode('utf-8'))
    z._cambio = True
    return True


def vaciar_celdas_solo_espacios(contenido, parte):
    """Deja vacías (con su mismo formato) las celdas de la hoja `parte` cuyo texto son SOLO espacios. Algunas plantillas traen celdas con 80-90
    espacios (p. ej. I26/I29 en gravedad específica fino, I25/I28 en grueso): ese texto invisible se desborda hacia la derecha y Excel no dibuja
    el borde del marco por donde pasa, así que la línea salía cortada al imprimir. Las plantillas tienen la hoja protegida, así que se corrige
    en la descarga. Nada visible cambia."""
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        xml = z.read(parte).decode('utf-8')
        sst = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = [''.join(re.findall(r'<t\b[^>]*>([^<]*)</t>', si)) for si in
                   re.findall(r'<si>(.*?)</si>', z.read('xl/sharedStrings.xml').decode('utf-8'), re.S)]
        en_blanco = {str(i) for i, t in enumerate(sst) if t and not t.strip()}

        def limpiar(m):
            apertura, cuerpo = m.group(1), m.group(2)
            v = re.search(r'<v>(\d+)</v>', cuerpo)
            if 't="s"' in apertura and v and v.group(1) in en_blanco:
                return re.sub(r'\s+t="s"', '', apertura)[:-1] + '/>'
            t = re.search(r'<is>.*?</is>', cuerpo, re.S)
            if 't="inlineStr"' in apertura and t and not ''.join(re.findall(r'<t\b[^>]*>([^<]*)</t>', t.group(0))).strip():
                return re.sub(r'\s+t="inlineStr"', '', apertura)[:-1] + '/>'
            return m.group(0)
        nuevo = re.sub(r'(<c\b[^>]*[^/]>)(.*?)</c>', limpiar, xml, flags=re.S)   # [^/]>: no toma las celdas vacías <c .../>
        if nuevo == xml:
            return contenido
        salida = BytesIO()
        with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, nuevo.encode('utf-8') if item.filename == parte else z.read(item.filename))
        return salida.getvalue()
