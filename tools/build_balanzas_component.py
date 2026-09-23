"""
Convierte el mockup HTML de "Comprobación intermedia de balanzas" (aprobado por el Jefe de Laboratorio;
versión tablet vertical, una sola columna) en un componente bidireccional de Streamlit, SIN tocar su
diseño ni sus cálculos:

  python tools/build_balanzas_component.py "<carpeta con Main.dc.html, support.js y vendor/>"

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


def main(src_dir):
    with open(os.path.join(src_dir, "Main.dc.html"), encoding="utf-8") as f:
        html = f.read()
    assert "  static init() {" in html, "no se encontró static init() del mockup"
    assert RESET_ORIGINAL in html, "no se encontró el botón 'Nuevo registro' del mockup"
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
    main(sys.argv[1])
