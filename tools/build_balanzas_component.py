"""
Convierte el mockup HTML de "Comprobación intermedia de balanzas" (aprobado por el Jefe de Laboratorio;
versión tablet vertical, una sola columna) en un componente bidireccional de Streamlit, SIN tocar su
diseño ni sus cálculos:

  python tools/build_balanzas_component.py [carpeta con Main.dc.html, support.js y vendor/]

Sin argumentos usa tools/balanzas_src/ (el HTML entregado por el Jefe, guardado tal cual). Si llega un HTML
nuevo, se reemplaza esa carpeta y se vuelve a correr. Además del puente con Streamlit, se le aplican los
ajustes de diseño pedidos (ver aplicar_ajustes): fallan con un aviso claro si el HTML nuevo cambió tanto
que ya no los encuentra.

Escribe balanzas_component/ (index.html + support.js + vendor/). Solo se le agregan, dentro de la clase
`Component` del propio mockup, (1) un puente con Streamlit — recibe el registro guardado, devuelve los
cambios para que la app los guarde en Supabase y ajusta la altura del marco al contenido para que la
página (no el marco) haga el scroll — y (2) el botón "Nuevo registro" avisa a la app en vez de borrar la
pantalla (la app crea el registro nuevo en la base de datos).
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(HERE), "balanzas_component")

# Se inserta dentro de la clase Component, antes de `static init()`.
METODOS = """  componentDidMount() { Component.__bridgeStart(this); }
  componentDidUpdate() { Component.__bridgeChanged(this); }

  static init() {"""

# Se agrega después de la clase (misma función de evaluación del runtime, ve `Component`).
PUENTE = r"""
(function () {
  var inst = null, token = null, seq = 0, timer = null, lastJson = null, lastH = 0;
  var nonce = Math.random().toString(36).slice(2, 8);   // distingue cargas del marco (el contador reinicia al recargar)
  function send(m) { window.parent.postMessage(Object.assign({ isStreamlitMessage: true }, m), '*'); }
  function payload() {
    var st = Object.assign({}, inst.state);
    var control = st.control || [];
    // lo que solo es de pantalla (paso abierto, paneles desplegados) no se guarda
    delete st.control; delete st.step; delete st.focusPt; delete st.eqOpen; delete st.showIssues;
    return { state: st, control: control };
  }
  // El diseño crece hacia abajo: el marco toma la altura del contenido y la página hace el scroll.
  function height() {
    var host = document.querySelector('#dc-root .sc-host') || document.getElementById('dc-root');
    if (!host) return;
    var bottom = 0;
    Array.prototype.forEach.call(host.children, function (el) {
      var b = el.getBoundingClientRect().bottom + (window.pageYOffset || 0);
      if (b > bottom) bottom = b;
    });
    var root = document.getElementById('dc-root');
    var h = Math.ceil(bottom) + 4 + (root && root.scrollWidth > root.clientWidth + 1 ? 17 : 0);   // 17 = barra horizontal
    if (h > 200 && Math.abs(h - lastH) > 2) { lastH = h; send({ type: 'streamlit:setFrameHeight', height: h }); }
  }
  Component.__bridgeStart = function (i) {
    if (inst) return;
    inst = i;
    // El marco de Streamlit no permite scroll (scrolling="no"): si la pantalla es más angosta que el
    // diseño (celular), el propio contenido se desplaza horizontalmente en vez de quedar cortado.
    var css = document.createElement('style');
    css.textContent = '#dc-root{overflow-x:auto;overflow-y:hidden}';
    document.head.appendChild(css);
    window.addEventListener('message', function (ev) {
      var d = ev.data;
      if (!d || d.type !== 'streamlit:render') return;
      var a = d.args || {};
      if (a.load_token && a.load_token !== token) {
        token = a.load_token;
        var ini = Component.init();
        var s = Object.assign(ini, a.state || {}, { control: a.control || [], step: 1, focusPt: 1 });
        inst.setState(s);
        setTimeout(function () { lastJson = JSON.stringify(payload()); }, 50);
      }
      height();
    });
    send({ type: 'streamlit:componentReady', apiVersion: 1 });
    height();
    setInterval(height, 700);
    if (window.ResizeObserver) {
      var host = document.querySelector('#dc-root .sc-host');
      if (host) new ResizeObserver(height).observe(host.firstElementChild || host);
    }
  };
  Component.__bridgeChanged = function () {
    height();
    if (!token) return;
    clearTimeout(timer);
    timer = setTimeout(function () {
      var p = payload(), j = JSON.stringify(p);
      if (j === lastJson) return;
      lastJson = j; seq += 1;
      send({ type: 'streamlit:setComponentValue', dataType: 'json',
             value: { token: token, seq: nonce + '-' + seq, state: p.state, control: p.control } });
    }, 900);
  };
  Component.__requestNew = function () {
    if (!token) return;
    seq += 1;
    send({ type: 'streamlit:setComponentValue', dataType: 'json', value: { token: token, seq: nonce + '-' + seq, action: 'new' } });
  };
})();
"""

RESET_ORIGINAL = ("resetAll: () => this.setState(Object.assign(Component.init(), { eq: s.eq, fecha: s.fecha, "
                  "proxima: s.proxima, control: s.control, pat: s.pat })),")
RESET_NUEVO = "resetAll: () => Component.__requestNew(),"

# ── Ajustes de diseño pedidos por el Jefe (sobre tablet vertical) ───────────────────────────────────────
# Exactitud: la fila de 8 columnas dejaba muy angostos los campos de ascendente/descendente y un solo
# resultado no decía cuál de los dos fallaba. Cada punto pasa a dos líneas: (# · aplica · carga · ± EMP) y,
# debajo, Ascendente y Descendente lado a lado, cada uno con su campo ancho, su error y su propio resultado.
EXA_MARCA_INI = ('<div style="display: grid; grid-template-columns: 22px 52px minmax(0, 1fr) minmax(0, 1fr) '
                 'minmax(0, 1fr) 86px 76px 100px; gap: 6px; padding: 0 6px 2px; align-items: end;">')
EXA_MARCA_FIN = "</sc-for>"
_CHIP = ('<span style="display: inline-flex; align-items: center; gap: 5px; height: 24px; padding: 0 9px; border-radius: 999px; '
         'font-size: 11px; font-weight: 600; white-space: nowrap; background: {{r.CHIP.bg}}; color: {{r.CHIP.fg}};">'
         '<span style="width: 6px; height: 6px; border-radius: 50%; background: {{r.CHIP.fg}};"></span>{{r.CHIP.label}}</span>')


def _mitad(titulo, campo, setter, err, errc, chip, aria):
    return (
        '<div style="display: flex; flex-direction: column; gap: 4px; min-width: 0;">'
        '<div style="display: flex; align-items: center; justify-content: space-between; gap: 6px; min-height: 24px;">'
        + '<span class="th">' + titulo + '</span>' + _CHIP.replace("CHIP", chip) + '</div>'
        '<div style="display: grid; grid-template-columns: minmax(0, 1fr) 84px; gap: 6px;">'
        '<input class="cap" style="padding: 0 10px; font-size: 16px;" inputmode="decimal" placeholder="{{placeholderNum}}" '
        'value="{{r.' + campo + '}}" onChange="{{r.' + setter + '}}" aria-label="' + aria + ' punto {{r.n}}">'
        '<div class="calc" style="padding: 0 8px; font-size: 14px; color: {{r.' + errc + '}};">{{r.' + err + '}}</div>'
        '</div></div>')


EXA_FILAS_NUEVAS = (
    '<sc-for list="{{exaRows}}" as="r" hint-placeholder-count="8">\n'
    '<div style="display: flex; flex-direction: column; gap: 10px; padding: 10px 12px; border-radius: 12px; '
    'border: 1px solid #E2DED4; background: {{r.bg}}; opacity: {{r.op}};">\n'
    '<div style="display: grid; grid-template-columns: 26px 64px minmax(0, 1fr) minmax(0, 1fr); gap: 10px; align-items: end;">\n'
    '<span class="mono" style="font-weight: 700; font-size: 16px; color: #5E6167; line-height: 44px;">{{r.n}}</span>\n'
    '<button type="button" onClick="{{r.toggle}}" aria-pressed="{{r.aplica}}" aria-label="Punto {{r.n}} aplica" '
    'style="height: 44px; border-radius: 999px; border: 1px solid {{r.tbd}}; background: {{r.tbg}}; color: {{r.tfg}}; '
    'font-size: 13px; font-weight: 600; cursor: pointer;">{{r.apLabel}}</button>\n'
    '<label style="display: flex; flex-direction: column; gap: 4px;"><span class="th" style="text-align: right;">Carga (g)</span>'
    '<input class="cap" style="padding: 0 10px; font-size: 16px;" inputmode="decimal" placeholder="0" value="{{r.carga}}" '
    'onChange="{{r.setCarga}}" aria-label="Carga punto {{r.n}}"></label>\n'
    '<label style="display: flex; flex-direction: column; gap: 4px;"><span class="th" style="text-align: right;">± EMP (g)</span>'
    '<input class="cap" style="padding: 0 10px; font-size: 16px;" inputmode="decimal" placeholder="0,00" value="{{r.emp}}" '
    'onChange="{{r.setEmp}}" aria-label="EMP punto {{r.n}}"></label>\n'
    '</div>\n'
    '<div style="display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px;">\n'
    + _mitad("Ascendente", "asc", "setAsc", "errA", "errAc", "chipA", "Indicación ascendente") + "\n"
    + _mitad("Descendente", "desc", "setDesc", "errD", "errDc", "chipD", "Indicación descendente") + "\n"
    '</div>\n'
    '</div>\n'
)
EXA_CHIP_JS_ORIGINAL = "chip: on ? this.chip(x.st) : this.chip(r.just ? 'na' : 'pend', r.just ? 'No aplica' : 'Justificar'),"
EXA_CHIP_JS_NUEVO = (EXA_CHIP_JS_ORIGINAL + "\n        chipA: on ? this.chip(x.a.st) : this.chip(r.just ? 'na' : 'pend', r.just ? 'No aplica' : 'Justificar'),"
                     "\n        chipD: on ? this.chip(x.d.st) : this.chip(r.just ? 'na' : 'pend', r.just ? 'No aplica' : 'Justificar'),")

# Textos de ayuda: qué es el límite aprobado de R y qué es un instrumento patrón.
AYUDA_R_ANTES = 'onChange="{{setRepLimite}}"></label>\n</div>\n'
AYUDA_R = (AYUDA_R_ANTES + '<p style="margin: 0; font-size: 13px; line-height: 1.45; color: #5E6167;"><strong>R</strong> es el rango de '
           'repetibilidad: la lectura mayor menos la menor de las 5 repeticiones. El <strong>límite aprobado de R</strong> es el '
           'máximo R permitido para esta balanza; lo define y aprueba el Jefe de Laboratorio antes de medir. Si R es menor o igual '
           'que el límite, la repetibilidad es Conforme.</p>\n')
AYUDA_PAT_ANTES = '<span class="lbl">Instrumentos patrón</span><button type="button" class="btn" onClick="{{addPat}}">+ Agregar patrón</button></div>\n'
AYUDA_PAT = (AYUDA_PAT_ANTES + '<p style="margin: 0; font-size: 13px; line-height: 1.45; color: #5E6167;">El <strong>instrumento patrón</strong> '
             'es el juego de masas (pesas) de referencia con el que se hace la comprobación. Registre su código interno, una descripción '
             '(por ejemplo, juego de masas clase F1), la fecha de su última calibración, cuándo vence y el número del certificado. '
             'Si está vencido, la comprobación no es válida.</p>\n')


def aplicar_ajustes(html):
    a = html.index(EXA_MARCA_INI)
    b = html.index(EXA_MARCA_FIN, html.index('<sc-for list="{{exaRows}}"', a)) + len(EXA_MARCA_FIN) + 1
    html = html[:a] + EXA_FILAS_NUEVAS + html[b:]
    for antes, despues, que in ((EXA_CHIP_JS_ORIGINAL, EXA_CHIP_JS_NUEVO, "resultado de exactitud"),
                                (AYUDA_R_ANTES, AYUDA_R, "campo límite de R"),
                                (AYUDA_PAT_ANTES, AYUDA_PAT, "lista de instrumentos patrón")):
        assert html.count(antes) == 1, f"no se encontró (o está repetido) el bloque de {que} en el HTML"
        html = html.replace(antes, despues, 1)
    return html



def main(src_dir):
    with open(os.path.join(src_dir, "Main.dc.html"), encoding="utf-8") as f:
        html = f.read()
    assert "  static init() {" in html, "no se encontró static init() del mockup"
    assert RESET_ORIGINAL in html, "no se encontró el botón 'Nuevo registro' del mockup"
    html = aplicar_ajustes(html)
    html = html.replace("  static init() {", METODOS, 1)
    html = html.replace(RESET_ORIGINAL, RESET_NUEVO, 1)
    fin = html.rindex("\n}\n</script>")
    html = html[: fin + 3] + PUENTE + html[fin + 3:]

    if os.path.isdir(OUT_DIR):
        shutil.rmtree(OUT_DIR)
    os.makedirs(os.path.join(OUT_DIR, "vendor"))
    with open(os.path.join(OUT_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)
    shutil.copy(os.path.join(src_dir, "support.js"), os.path.join(OUT_DIR, "support.js"))
    for nombre in ("react.js", "react-dom.js"):
        shutil.copy(os.path.join(src_dir, "vendor", nombre), os.path.join(OUT_DIR, "vendor", nombre))
    print("ok ->", OUT_DIR)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "balanzas_src"))
