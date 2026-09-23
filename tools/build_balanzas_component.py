"""
Convierte el HTML exportado de "Comprobación intermedia de balanzas" (mockup aprobado por el Jefe de
Laboratorio) en un componente bidireccional de Streamlit, SIN tocar su diseño ni sus cálculos:

  python tools/build_balanzas_component.py "<ruta al HTML exportado>"

Escribe balanzas_component/index.html. Solo se le agregan, dentro de la clase `Component` del propio
mockup, (1) un puente con Streamlit — recibe el registro guardado y devuelve los cambios para que la
app los guarde en Supabase — y (2) el botón "Nuevo registro" avisa a la app en vez de borrar la
pantalla (la app crea el registro nuevo en la base de datos).
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "balanzas_component", "index.html")

# Se inserta dentro de la clase Component (antes de los helpers del mockup).
METODOS = """  componentDidMount() { Component.__bridgeStart(this); }
  componentDidUpdate() { Component.__bridgeChanged(this); }

  // ---------- helpers ----------"""

# Se agrega después de la clase (misma función de evaluación del runtime, ve `Component`).
PUENTE = r"""
(function () {
  var inst = null, token = null, seq = 0, timer = null, lastJson = null;
  var nonce = Math.random().toString(36).slice(2, 8);   // distingue cargas del marco (el contador reinicia al recargar)
  function send(m) { window.parent.postMessage(Object.assign({ isStreamlitMessage: true }, m), '*'); }
  function payload() {
    var st = Object.assign({}, inst.state);
    var control = st.control || [];
    delete st.control; delete st.step; delete st.focusPt;   // el paso abierto es solo de pantalla
    return { state: st, control: control };
  }
  function height() { send({ type: 'streamlit:setFrameHeight', height: 1000 }); }   // el diseño es una "app" con scroll interno
  Component.__bridgeStart = function (i) {
    if (inst) return;
    inst = i;
    window.addEventListener('message', function (ev) {
      var d = ev.data;
      if (!d || d.type !== 'streamlit:render') return;
      var a = d.args || {};
      height();
      if (a.load_token && a.load_token !== token) {
        token = a.load_token;
        var ini = Component.init();
        var s = Object.assign(ini, a.state || {}, { control: a.control || [], step: 1, focusPt: 1 });
        inst.setState(s);
        setTimeout(function () { lastJson = JSON.stringify(payload()); }, 50);
      }
    });
    send({ type: 'streamlit:componentReady', apiVersion: 1 });
    height();
  };
  Component.__bridgeChanged = function () {
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

# El mockup es un "artboard" fijo de 1440x1280: se adapta al ancho de la página y al alto del marco
# (el contenido sigue haciendo scroll adentro, igual que en el original).
ARTBOARD_ORIGINAL = "width: 1440px; height: 1280px; box-sizing: border-box; padding: 24px 28px;"
ARTBOARD_NUEVO = "width: 100%; min-width: 1360px; height: 100vh; box-sizing: border-box; padding: 24px 28px;"


def main(src):
    lines = open(src, encoding="utf-8").read().split("\n")
    idx = next(i for i, l in enumerate(lines) if '<script type="__bundler/template">' in l) + 1
    template = json.loads(lines[idx])

    assert "  // ---------- helpers ----------" in template, "no se encontró el bloque de helpers del mockup"
    assert RESET_ORIGINAL in template, "no se encontró el botón 'Nuevo registro' del mockup"
    template = template.replace("  // ---------- helpers ----------", METODOS, 1)
    template = template.replace(RESET_ORIGINAL, RESET_NUEVO, 1)
    assert ARTBOARD_ORIGINAL in template, "no se encontró el contenedor principal del mockup"
    template = template.replace(ARTBOARD_ORIGINAL, ARTBOARD_NUEVO, 1)
    fin = template.rindex("\n}\n</script>")
    template = template[: fin + 3] + PUENTE + template[fin + 3:]

    lines[idx] = json.dumps(template, ensure_ascii=False).replace("</", "<\\u002F")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("ok ->", OUT, os.path.getsize(OUT), "bytes")


if __name__ == "__main__":
    main(sys.argv[1])
