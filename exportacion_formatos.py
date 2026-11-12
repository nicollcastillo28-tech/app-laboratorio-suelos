"""Diligencia entradas sin volver a guardar ni rediseñar los formatos oficiales.

openpyxl se usa únicamente para leer. La descarga conserva las partes originales
del archivo (incluidos macros, gráficos, estilos y fórmulas) y sustituye solamente
los valores de entrada solicitados por el generador. Las excepciones autorizadas
para sustituir una fórmula se declaran por hoja y celda; actualmente, G21 del 006.
"""
from io import BytesIO
import html
import math
import re
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, range_boundaries
from presentacion_excel import PaqueteZip, ajustar_presentacion, ajustar_presentacion_en

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


class FormatoIncompatibleError(ValueError):
    """La entrada no cabe en el formato original sin alterar su estructura."""


def partes_hojas(contenido):
    with zipfile.ZipFile(BytesIO(contenido)) as z:
        rels = {r.attrib['Id']: r.attrib['Target']
                for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        hojas = ET.fromstring(z.read('xl/workbook.xml')).find(S + 'sheets')
        resultado = {}
        for hoja in hojas:
            destino = rels.get(hoja.get(R + 'id'))
            if destino:
                resultado[hoja.get('name')] = (destino.lstrip('/') if destino.startswith('/') else 'xl/' + destino)
        return resultado


def escribir_entradas(contenido, parte, valores, *, celdas_formula_editables=()):
    """Cambia entradas; una fórmula sólo se sustituye con autorización explícita por celda."""
    paquete = PaqueteZip(contenido)
    escribir_entradas_en(paquete, parte, valores, celdas_formula_editables=celdas_formula_editables)
    return paquete.bytes()


def escribir_entradas_en(z, parte, valores, *, celdas_formula_editables=()):
    """escribir_entradas sobre un PaqueteZip (lo modifica): así LibroFormatoOriginal llena todas sus hojas comprimiendo el archivo una vez."""
    xml = z.read(parte).decode('utf-8')
    # Índice de celdas y filas en una sola pasada (posición en el XML). Antes se leía la hoja entera con ElementTree para saber qué celdas
    # tienen fórmula, y luego se buscaba y rearmaba el XML por cada celda: con las miles de lecturas de la máquina de Consolidación la
    # descarga tardaba minutos.
    celdas, celdas_por_fila = {}, {}
    for m in re.finditer(r'<c\b\s+r="([A-Z]+)(\d+)"([^>]*?)(/>|>)', xml):
        fin = m.end() if m.group(4) == '/>' else xml.index('</c>', m.end()) + 4
        ref_c = m.group(1) + m.group(2)
        if ref_c not in celdas:
            celdas[ref_c] = (m.start(), fin, m.group(3))
            celdas_por_fila.setdefault(int(m.group(2)), []).append((m.group(1), m.start()))
    filas = {}
    for m in re.finditer(r'<row\b\s+r="(\d+)"[^>]*?(/>|>)', xml):
        n = int(m.group(1))
        if n in filas:
            continue
        filas[n] = (m.start(), m.end(), None) if m.group(2) == '/>' else (m.start(), xml.index('</row>', m.end()) + 6, m.end())
    combinaciones = [range_boundaries(r) for r in re.findall(r'<mergeCell\b[^>]*\bref="([^"]+)"', xml)]
    if len(re.findall(r'<c\b', xml)) != len(celdas):
        # Celdas que no traen r como primer atributo (no lo hace Excel): se lee con ElementTree, como antes.
        existentes_et = {c.get('r'): c for c in ET.fromstring(xml).iter(S + 'c')}

        def formula_de(ref):
            c = existentes_et.get(ref)
            f = c.find(S + 'f') if c is not None else None
            return None if f is None else (f.get('t') or '')
    else:
        def formula_de(ref):
            """Tipo de la fórmula de la celda ('' normal, 'shared', 'array') o None si no tiene."""
            if ref not in celdas:
                return None
            inicio, fin, _a = celdas[ref]
            f = re.search(r'<f\b([^>]*)', xml[inicio:fin])
            if f is None:
                return None
            t = re.search(r'\bt="([^"]*)"', f.group(1))
            return t.group(1) if t else ''
    entradas = {}
    formulas_sustituidas = False
    for ref, valor in valores.items():
        if isinstance(valor, str) and valor.startswith('='):
            raise FormatoIncompatibleError(f'{parte}!{ref}: no se permite introducir una fórmula')
        fila, col = coordinate_to_tuple(ref)
        for c1, f1, c2, f2 in combinaciones:
            if c1 <= col <= c2 and f1 <= fila <= f2 and (col, fila) != (c1, f1):
                if valor not in (None, ''):
                    raise FormatoIncompatibleError(f'{parte}!{ref}: la casilla está combinada con otra; se conserva el formato original')
                break
        else:
            tipo_formula = formula_de(ref)
            if tipo_formula is not None:
                if ref not in celdas_formula_editables:
                    if valor not in (None, ''):
                        raise FormatoIncompatibleError(f'{parte}!{ref}: contiene una fórmula del formato original')
                    continue
                if tipo_formula in ('shared', 'array'):
                    raise FormatoIncompatibleError(f'{parte}!{ref}: pertenece a una fórmula compartida o matricial')
                formulas_sustituidas = True
            if isinstance(valor, float) and not math.isfinite(valor):
                raise ValueError(f'{parte}!{ref}: valor numérico no finito')
            entradas[ref] = valor

    def reemplazo(ref, atributos, valor):
        # Conservar todos los atributos de formato de la celda, salvo su tipo de dato.
        atributos = re.sub(r'\s+t="[^"]*"', '', atributos)
        atributos = re.sub(r'\s+(?:cm|vm)="[^"]*"', '', atributos)
        if valor is None or valor == '':
            return f'<c r="{ref}"{atributos}/>'
        if isinstance(valor, bool):
            return f'<c r="{ref}"{atributos} t="b"><v>{int(valor)}</v></c>'
        if isinstance(valor, (int, float)):
            return f'<c r="{ref}"{atributos}><v>{valor!r}</v></c>'
        texto = html.escape(str(valor), quote=False)
        return f'<c r="{ref}"{atributos} t="inlineStr"><is><t xml:space="preserve">{texto}</t></is></c>'

    # Todos los cambios se aplican de una vez sobre el XML.
    cambios = []   # (inicio, fin, columna, texto)
    nuevas_en_fila = {}
    for ref, valor in entradas.items():
        if ref in celdas:
            inicio, fin, atributos = celdas[ref]
            cambios.append((inicio, fin, 0, reemplazo(ref, atributos, valor)))
            continue
        fila, col = coordinate_to_tuple(ref)
        if fila not in filas:
            # La ausencia de una fila de entrada es un cambio de formato pendiente de revisar.
            if valor not in (None, ''):
                raise FormatoIncompatibleError(f'{parte}!{ref}: la fila de entrada no existe en el formato')
            continue
        if filas[fila][2] is None and valor in (None, ''):
            continue
        nuevas_en_fila.setdefault(fila, []).append((col, ref, valor))
    for fila, nuevas in nuevas_en_fila.items():
        inicio, fin, inicio_cuerpo = filas[fila]
        nuevas.sort()
        if inicio_cuerpo is None:   # <row .../> sin celdas: se abre la fila con las nuevas, en orden de columna
            tag = xml[inicio:fin]
            cambios.append((inicio, fin, 0, tag[:-2] + '>' + ''.join(reemplazo(r, '', v) for _c, r, v in nuevas) + '</row>'))
            continue
        existentes_fila = sorted((column_index_from_string(letras), pos) for letras, pos in celdas_por_fila.get(fila, []))
        for col, ref, valor in nuevas:
            lugar = next((pos for c, pos in existentes_fila if c > col), fin - 6)
            cambios.append((lugar, lugar, col, reemplazo(ref, '', valor)))
    if cambios:
        cambios.sort(key=lambda x: (x[0], x[1], x[2]))
        trozos, cursor = [], 0
        for inicio, fin, _col, texto in cambios:
            trozos.append(xml[cursor:inicio])
            trozos.append(texto)
            cursor = max(cursor, fin)
        trozos.append(xml[cursor:])
        xml = ''.join(trozos)

    # La cadena de cálculo anterior contiene la fórmula que acabamos de retirar. Excel debe reconstruir esa caché al abrir el libro.
    if formulas_sustituidas:
        z.quitar('xl/calcChain.xml')
        if '[Content_Types].xml' in z.datos:
            z.escribir('[Content_Types].xml', re.sub(rb'<Override\b(?=[^>]*\bPartName="/xl/calcChain\.xml")[^>]*/>', b'',
                                                      z.read('[Content_Types].xml')))
        if 'xl/_rels/workbook.xml.rels' in z.datos:
            z.escribir('xl/_rels/workbook.xml.rels', re.sub(rb'<Relationship\b(?=[^>]*\bType="[^"]*/calcChain")[^>]*/>', b'',
                                                             z.read('xl/_rels/workbook.xml.rels')))
    if 'xl/workbook.xml' in z.datos:
        # Invalidar las cachés de cálculo tras llenar entradas, sin cambiar fórmulas.
        texto = z.read('xl/workbook.xml').decode('utf-8')
        def recalcular(m):
            tag = re.sub(r'\s+(?:fullCalcOnLoad|forceFullCalc)="[^"]*"', '', m.group(0))
            return tag.replace('/>', ' fullCalcOnLoad="1" forceFullCalc="1"/>')
        z.escribir('xl/workbook.xml', re.sub(r'<calcPr\b[^>]*/>', recalcular, texto).encode('utf-8'))
    z.escribir(parte, xml.encode('utf-8'))
    ajustar_presentacion_en(z, parte, entradas)


class _HojaEntradas:
    def __init__(self, hoja):
        self._hoja = hoja
        self.entradas = {}

    def __getattr__(self, nombre):
        return getattr(self._hoja, nombre)

    def __getitem__(self, ref):
        return self._hoja[ref]

    def __setitem__(self, ref, valor):
        # Los generadores anteriores reparaban fórmulas del formato anterior.
        # Esas reparaciones quedan desactivadas para conservar el formato nuevo.
        if isinstance(valor, str) and valor.startswith('='):
            if self._hoja[ref].value == valor:
                return
            raise FormatoIncompatibleError(f'{self._hoja.title}!{ref}: se conserva la fórmula original')
        self.entradas[ref] = valor


# Leer una plantilla con openpyxl cuesta cerca de un segundo; un paquete de 30 muestras la leía 30 veces. El libro de openpyxl SOLO se consulta (hojas,
# celdas combinadas): lo que se llena queda en `_HojaEntradas.entradas` y se escribe sobre los bytes originales, así que se comparte entre usos.
_CACHE_LIBROS = {}


class LibroFormatoOriginal:
    def __init__(self, path, *, celdas_formula_editables=None, **kwargs):
        import os
        clave = (str(path), os.path.getmtime(path), tuple(sorted((k, repr(v)) for k, v in kwargs.items())))
        en_cache = _CACHE_LIBROS.get(clave)
        if en_cache is None:
            with open(path, 'rb') as f:
                original = f.read()
            en_cache = (original, load_workbook(BytesIO(original), **kwargs))
            if len(_CACHE_LIBROS) > 24:
                _CACHE_LIBROS.clear()
            _CACHE_LIBROS[clave] = en_cache
        self._original, self._libro = en_cache
        self._formulas_editables = celdas_formula_editables or {}
        self._hojas = {h.title: _HojaEntradas(h) for h in self._libro.worksheets}

    @property
    def sheetnames(self):
        return self._libro.sheetnames

    @property
    def worksheets(self):
        return list(self._hojas.values())

    def __getitem__(self, nombre):
        return self._hojas[nombre]

    def __delitem__(self, nombre):
        raise FormatoIncompatibleError(f'No se permite eliminar la hoja original {nombre}')

    def save(self, archivo):
        partes = partes_hojas(self._original)
        paquete = PaqueteZip(self._original)
        for nombre, hoja in self._hojas.items():
            if hoja.entradas:
                escribir_entradas_en(paquete, partes[nombre], hoja.entradas,
                                     celdas_formula_editables=self._formulas_editables.get(nombre, ()))
        archivo.write(paquete.bytes())
