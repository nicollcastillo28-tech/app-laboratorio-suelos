"""Datos de ejemplo para el Modo de prueba (pantalla "pruebas" de app.py).

Cada ensayo trae un juego de datos ficticios pero coherentes entre sí (masas, humedades y lecturas que dan resultados dentro de
lo que se espera de un suelo arcilloso o, en roca, de un núcleo NQ), para poder ver el formulario lleno, los resultados, las
alertas y el Excel descargable sin digitar nada. Módulo puro (sin Streamlit): app.py le agrega los equipos y la norma.
Todo se devuelve como texto, igual que lo guarda un campo de texto de los formularios."""

import math


def _n(v, d=2):
    return f"{v:.{d}f}"


def _humedad(prefijo, rec, suelo_seco, w_pct, *, horas, recipiente, humedo_key="humedo", rec_key="masa_rec", rec_no_key="recipiente",
             sufijo=""):
    """Campos de humedad (recipiente, húmeda, secas y masa del recipiente) con el % `w_pct` sobre `suelo_seco` g."""
    humedo = rec + suelo_seco * (1 + w_pct / 100)
    seco = rec + suelo_seco
    d = {f"{prefijo}{rec_no_key}": recipiente, f"{prefijo}{humedo_key}": _n(humedo), f"{prefijo}{rec_key}": _n(rec)}
    for h in horas:
        d[f"{prefijo}seco_{h}{sufijo}"] = _n(seco)
    return d


def _granulometria():
    # Arcilla de baja plasticidad: 500 g de muestra seca, 70 % pasa el tamiz No. 200.
    retenido = {"s_4": 15.0, "s_10": 20.0, "s_20": 15.0, "s_40": 25.0, "s_60": 20.0, "s_100": 25.0, "s_200": 30.0}
    suma = sum(retenido.values())
    rec, seco_antes = 110.0, 610.0
    d = {"p200_recipiente_antes": "R-01", "p200_recipiente_despues": "R-02", "p200_masa_recipiente_antes": _n(rec),
         "p200_masa_recipiente_despues": _n(rec), "p200_seco_mas_recipiente_antes": _n(seco_antes),
         "p200_seco_mas_recipiente_despues": _n(rec + suma)}
    for suf, v in (("antes", seco_antes), ("despues", rec + suma)):
        for h in (14, 15, 16):
            d[f"p200_seco_{h}h_{suf}"] = _n(v)
    for k in ("s_3", "s_2p5", "s_2", "s_1p5", "s_1", "s_34", "s_12", "s_38", "s_4", "s_10", "s_20", "s_40", "s_60", "s_100", "s_200"):
        d[k] = _n(retenido.get(k, 0.0))
    d["masa_inicial_seca"] = seco_antes - rec
    return d


def _limites():
    d = {"lim_metodo": "Método A"}
    # Límite líquido: 3 puntos (17, 25 y 33 golpes) con ~20 g de suelo seco; humedades 48, 45 y 42 %.
    for i, (golpes, rec, w) in enumerate(((17, 12.10, 48.0), (25, 12.35, 45.0), (33, 12.20, 42.0)), start=1):
        seco = rec + 20.0
        d.update({f"lim_ll_recipiente_{i}": f"L{i}", f"lim_ll_golpes_{i}": str(golpes), f"lim_ll_recip_masa_{i}": _n(rec),
                  f"lim_ll_humedo_{i}": _n(seco + 20.0 * w / 100), f"lim_ll_seco_{i}": _n(seco)})
        for h in (14, 15, 16):
            d[f"lim_ll_seco_{h}h_{i}"] = _n(seco)
    # Límite plástico: 2 ensayos, ~22 %.
    for i, (rec, w) in enumerate(((12.05, 22.0), (12.40, 22.5)), start=1):
        seco = rec + 10.0
        d.update({f"lim_lp_recipiente_{i}": f"P{i}", f"lim_lp_recip_masa_{i}": _n(rec),
                  f"lim_lp_humedo_{i}": _n(seco + 10.0 * w / 100), f"lim_lp_seco_{i}": _n(seco)})
        for h in (14, 15, 16):
            d[f"lim_lp_seco_{h}h_{i}"] = _n(seco)
    return d


def _humedad_natural():
    d = {"hum_recipiente": "A-12", "hum_masa_recipiente": "18.45", "hum_masa_humedo_mas_recipiente": "117.65",
         "hum_seco_mas_recipiente": "98.45", "hum_temp_horno": "110 ± 5 °C", "hum_metodo": "Método A"}
    for h in (14, 15, 16):
        d[f"hum_seco_{h}h"] = "98.45"
    return d


def _masa_unitaria():
    d = {"mu_metodo": "Parafinado", "mu_peso_aire": "180.00", "mu_peso_aire_par": "190.50", "mu_peso_agua_par": "81.00",
         "mu_temp_agua": "22.0"}
    # Sin ensayo de humedad asignado en la muestra, la humedad se calcula con estas masas (24 %).
    d.update(_humedad("mu_hum_", 18.30, 70.0, 24.0, horas=(17, 18, 19), recipiente="A-15", rec_key="masa_rec"))
    return d


def _pasa200():
    g = _granulometria()
    return {k: v for k, v in g.items() if k.startswith("p200_") or k == "masa_inicial_seca"}


def _cbr():
    base = [0.04, 0.23, 0.57, 0.81, 1.0, 1.15, 1.27, 1.38, 1.49, 1.85, 2.13, 2.36]
    d = {"cbr_molde": "3", "cbr_diametro": "15.24", "cbr_altura": "11.64", "cbr_masa_molde": "6300.00",
         "cbr_masa_muestra_molde_antes": "10500.00", "cbr_masa_muestra_molde_despues": "10800.00",
         "cbr_exp_lectura_inicial": "0.100", "cbr_exp_lectura_final": "0.118", "cbr_pesas_antes": "4540", "cbr_pesas_despues": "4540",
         "cbr_tiempo_inmersion_antes": "0", "cbr_tiempo_inmersion_despues": "4", "cbr_rango": "5 kN",
         "cbr_pen_antes_0": "0", "cbr_pen_despues_0": "0"}
    d.update({"cbr_desp_recipiente": "C-8", "cbr_desp_masa_recipiente": "20.00", "cbr_desp_masa_humedo": "127.10",
              "cbr_desp_masa_seco": "105.00"})
    for i, f in enumerate(base, start=1):
        d[f"cbr_pen_antes_{i}"] = _n(f * 1.08)
        d[f"cbr_pen_despues_{i}"] = _n(f * 0.66)
    return d


def _corte_directo():
    d = {"corte_tipo": "CD", "corte_condicion": "Inalterada", "corte_temp_inicial": "22.0", "corte_temp_final": "22.5",
         "corte_hum_inicial": "58", "corte_hum_final": "56"}
    for i, (masa_total, normal) in enumerate(((172.1, 0.5), (172.6, 1.0), (171.8, 2.0)), start=1):
        p = f"corte_m{i}_"
        d.update({p + "masa_inicial_anillo": _n(masa_total), p + "masa_anillo": "55.00", p + "altura_anillo": "2.00",
                  p + "diametro_anillo": "6.35", p + "velocidad_corte": "0.05", p + "esfuerzo_normal": _n(normal),
                  p + "temp_secado": "110 °C (Método B)"})
        for suf, w in (("inicial", 24.0), ("final", 26.0)):
            d[p + f"hum_recipiente_{suf}"] = f"M{i}{suf[0].upper()}"
            d[p + f"hum_masa_recipiente_{suf}"] = "20.00"
            d[p + f"hum_masa_humedo_{suf}"] = _n(20.0 + 60.0 * (1 + w / 100))
            for h in (17, 18, 19):
                d[p + f"hum_seco_{h}h_{suf}"] = "80.00"
    d.update({"corte_ge_masa_suelo_seco": "50.00", "corte_ge_masa_pic_muestra": "681.10", "corte_ge_masa_pic": "650.00",
              "corte_ge_temperatura": "22.0", "corte_ge_picnometro": "1"})
    return d


def _gravedad_especifica():
    # "Ambos": agregado fino (pasa el tamiz No. 4) y grueso (lo retiene); además quedan los datos del método de arcillas y limos
    # (INV E-128, picnómetro 3) por si se cambia la opción.
    return {"gesp_sel": "Ambos", "gesp_f_masa_sss": "500.00", "gesp_f_masa_pic_agua_suelo": "977.00", "gesp_f_masa_pic_agua": "665.20",
            "gesp_f_temp": "22.0", "gesp_f_masa_seco": "492.50", "gesp_f_pct_retenido": "8.0",
            "gesp_g_masa_sss": "1520.00", "gesp_g_masa_sumergido": "960.00", "gesp_g_masa_seco": "1500.00", "gesp_g_temp": "22.0",
            "gesp_a_picnometro": "3", "gesp_a_masa_aire": "50.00", "gesp_a_masa_pic_agua": "347.96", "gesp_a_masa_pic_muestra": "379.34", "gesp_a_temp": "22.0",
            "gesp_a_pct_retenido": "0"}


def _proctor():
    d = {"proc_metodo": "A", "proc_sobretamano_tamiz": "—"}
    # 4 puntos del Proctor estándar: humedades 12, 15, 18 y 21 %; densidad seca máxima ~1.72 g/cm³ hacia el 18 %.
    seca = (1.62, 1.70, 1.72, 1.66)
    for i, (w, rd) in enumerate(zip((12.0, 15.0, 18.0, 21.0), seca), start=1):
        masa_humeda = rd * (1 + w / 100) * 944.0
        d.update({f"proc_{i}_golpes": "25", f"proc_{i}_molde": "M-1", f"proc_{i}_capas": "3", f"proc_{i}_masa_molde": "4150.0",
                  f"proc_{i}_masa_humedo_molde": _n(4150.0 + masa_humeda, 1), f"proc_{i}_volumen_molde": "944"})
        d.update(_humedad(f"proc_{i}_hum_", 20.0, 100.0, w, horas=(16, 17, 18, 19), recipiente=f"P-{i}", humedo_key="masa_humedo",
                          rec_key="masa_recipiente", rec_no_key="recipiente", sufijo="h"))
    # CBR de suelo compactado: 3 moldes (10, 25 y 56 golpes), humedad de compactación 16 % y después de la inmersión 19 %.
    base = [0.04, 0.23, 0.57, 0.81, 1.0, 1.15, 1.27, 1.38, 1.49, 1.85, 2.13, 2.36]
    for i, (rho, p01, exp_fin) in enumerate(((1.914, 0.55, 0.112), (1.995, 0.85, 0.108), (2.065, 1.20, 0.104)), start=1):
        p = f"cbrc_{i}_"
        d.update({p + "molde": str(i), p + "capas": "5", p + "masa_molde": "6300.0", p + "altura": "11.64", p + "diametro": "15.24",
                  p + "masa_muestra_molde": _n(6300.0 + rho * 2123.0, 1), p + "exp_inicial": "0.100", p + "exp_final": _n(exp_fin, 3)})
        d.update(_humedad(p + "hc_", 20.0, 100.0, 16.0, horas=(16, 17, 18, 19), recipiente=f"C{i}", humedo_key="masa_humedo",
                          rec_key="masa_recipiente", rec_no_key="recipiente", sufijo="h"))
        d.update(_humedad(p + "hd_", 20.0, 100.0, 19.0, horas=(16, 17, 18, 19), recipiente=f"D{i}", humedo_key="masa_humedo",
                          rec_key="masa_recipiente", rec_no_key="recipiente", sufijo="h"))
        for j, f in enumerate(base, start=1):
            d[p + f"pen_{j}"] = _n(f * p01)
    return d


def _materia_organica():
    return {"mo_recipiente": "C-5", "mo_masa_crisol": "25.00", "mo_masa_crisol_seco": "45.00", "mo_masa_crisol_ignicion": "44.40",
            "mo_metodo": "A", "mo_temp": "110 °C"}


def _limite_contraccion():
    d = {"lc_recipiente": "1", "lc_volumen": "16.30", "lc_masa_recipiente": "25.50", "lc_masa_aire": "24.80",
         "lc_masa_aire_parafinada": "27.00", "lc_masa_sumergida_parafinada": "11.39"}
    d.update({"lc_natural_hum_recipiente": "12.00", "lc_natural_hum_humedo": "49.20", "lc_natural_hum_seco": "42.00",
              "lc_contraccion_hum_recipiente": "12.00", "lc_contraccion_hum_humedo": "51.00", "lc_contraccion_hum_seco": "42.00"})
    return d


def _consolidacion():
    d = {"cons_condicion": "Inalterada", "cons_temp_ini": "22.0", "cons_temp_fin": "22.5", "cons_masa_anillo": "52.00",
         "cons_masa_anillo_muestra_ini": "169.10", "cons_masa_anillo_muestra_fin": "170.40", "cons_diametro": "63.50",
         "cons_altura": "20.00", "cons_temp_secado": "110 °C", "cons_metodo": "A", "cons_consolidometro": "GDA-E-003",
         "cons_pic_no": "3", "cons_pic_temp": "22.0", "cons_pic_masa_seco": "50.00", "cons_precarga": "0"}
    # Gs ~2.68 con el picnómetro 3 (masa lleno de agua = -0.0614·T + 349.31)
    wpw = -0.0614 * 22.0 + 349.31
    d["cons_pic_masa_agua_suelo"] = _n(wpw + 50.0 - 50.0 / 2.68)
    for pref, w in (("ini", 26.0), ("fin", 22.0)):
        d[f"cons_{pref}_recipiente"] = f"I-{pref[0]}"
        d[f"cons_{pref}_recipiente_masa"] = "20.00"
        d[f"cons_{pref}_humedo"] = _n(20.0 + 60.0 * (1 + w / 100))
        for h in (17, 18, 19):
            d[f"cons_{pref}_seco_{h}"] = "80.00"
    # Deformación de cada escalón de carga (mm) en la bitácora y lecturas de la máquina (un archivo por carga).
    incrementos = [0.020, 0.030, 0.050, 0.090, 0.160, 0.280, 0.450, 0.600]
    acumulada = 0.0
    tiempos = [0, 6, 15, 30, 60, 120, 240, 480, 900, 1800, 3600, 7200, 14400, 28800, 57600, 86400]
    for i, inc in enumerate(incrementos, start=1):
        puntos = [[float(t), round(acumulada + inc * (1 - math.exp(-t / 1500.0)), 4)] for t in tiempos]
        acumulada = puntos[-1][1]
        d[f"cons_maq_{i}"] = puntos
    for i, (c, ds) in enumerate(zip((0.05, 0.09, 0.17, 0.31, 0.55, 0.86, 1.30), (1.60, 1.58, 1.55, 1.50, 1.43, 1.35, 1.25)), start=1):
        d[f"cons_def_{i}_carga"] = _n(c, 3)
        d[f"cons_def_{i}_descarga"] = _n(ds, 3)
    return d


def _curva_maquina(area_cm2, altura_cm, qu_kpa, eps_falla, velocidad_mm_min, eps_final):
    """[[t s, fuerza kN, deformación mm]] de un ensayo de compresión con falla en `eps_falla` (fracción) y caída posterior."""
    puntos = []
    t_final = altura_cm * 10 * eps_final / velocidad_mm_min * 60
    n = 60
    for k in range(n + 1):
        t = t_final * k / n
        mm = velocidad_mm_min * t / 60
        eps = mm / (altura_cm * 10)
        x = eps / eps_falla
        sigma = qu_kpa * (2 * x - x * x) if x <= 1 else qu_kpa * max(1 - 0.35 * (x - 1), 0.2)
        fuerza = sigma * area_cm2 / (1 - eps) / 10000
        puntos.append([round(t, 1), round(fuerza, 4), round(mm, 4)])
    return puntos


def _compresion_inconfinada():
    d = {"ci_condicion": "Inalterada", "ci_muestreo": "Tubo Shelby", "ci_temp_ini": "22.0", "ci_temp_fin": "22.5",
         "ci_h_1": "7.62", "ci_h_2": "7.60", "ci_h_3": "7.64", "ci_d_1": "3.81", "ci_d_2": "3.80", "ci_d_3": "3.82",
         "ci_peso": "165.00", "ci_temp_secado": "110 °C", "ci_metodo": "A", "ci_hum_antes": "Antes del ensayo",
         "ci_hum_muestra": "Cortes de muestra", "ci_penetrometro": "1.20", "ci_falla": "PLANO INCLINADO", "ci_velocidad": "1",
         "ci_tiempo_falla": "3.10"}
    d.update(_humedad("ci_hum_", 15.0, 50.0, 24.0, horas=(17, 18, 19), recipiente="H-2", rec_key="masa_rec"))
    area = math.pi * 3.81 ** 2 / 4
    d["ci_maq"] = _curva_maquina(area, 7.62, 150.0, 0.04, 1.0, 0.08)
    return d


def _compresion_roca():
    d = {"roca_temp_ini": "22.0", "roca_temp_fin": "22.5", "roca_h_1": "11.90", "roca_h_2": "11.90", "roca_h_3": "11.88",
         "roca_d_1": "4.76", "roca_d_2": "4.76", "roca_d_3": "4.77", "roca_peso": "561.00", "roca_temp_secado": "110 °C",
         "roca_metodo": "A", "roca_hum_antes": "Antes del ensayo", "roca_hum_muestra": "Muestra completa",
         "roca_falla": "FALLA POR DIVISIÓN AXIAL", "roca_velocidad": "0.50", "roca_tiempo_falla": "1.20",
         "roca_esfuerzo_maximo_manual": "60.0"}
    d.update(_humedad("roca_hum_", 25.0, 100.0, 0.8, horas=(17, 18, 19), recipiente="R-7", rec_key="masa_rec"))
    area = math.pi * 4.76 ** 2 / 4
    d["roca_maq"] = _curva_maquina(area, 11.9, 60000.0, 0.0045, 0.5, 0.0065)
    return d


def _carga_puntual():
    d = {}
    for i, p in enumerate((4.8, 5.4, 4.1, 5.9, 4.6), start=1):
        d.update({f"cp_{i}_sentido": "DIAMETRAL", f"cp_{i}_carga": _n(p), f"cp_{i}_d": "47.60", f"cp_{i}_l1": "62.00",
                  f"cp_{i}_masa": _n(380 + 12 * i, 0)})
    d.update(_humedad("cp_hum_", 25.0, 100.0, 1.0, horas=(14, 15, 16), recipiente="CP-1", rec_key="masa_rec"))
    return d


def _solidez_sulfatos():
    grad = {22: 0, 23: 0, 24: 0, 25: 0, 26: 0, 27: 250, 28: 600, 29: 450, 30: 700, 31: 400, 32: 560, 33: 480, 34: 420, 35: 380,
            36: 330, 37: 250, 38: 180}
    d = {"sulf_masa_total": _n(sum(grad.values())), "sulf_solucion": "Sulfato de sodio", "sulf_ciclos": "5"}
    for k, v in grad.items():
        d[f"sulf_grad_{k}"] = _n(v)
    # fracción: (masa inicial, masa final después de los ciclos)
    fracciones = {22: (1500.0, 1451.0), 23: (1500.0, 1448.0), 24: (1000.0, 962.0), 25: (300.0, 287.5), 34: (100.0, 96.2), 33: (100.0, 95.8),
                  32: (100.0, 96.8), 31: (100.0, 97.4), 30: (100.0, 98.1), 29: (100.0, 98.6), 35: (100.0, 96.0)}
    for k, (ini, fin) in fracciones.items():
        d[f"sulf_{k}_ini"] = _n(ini)
        d[f"sulf_{k}_fin"] = _n(fin)
    return d


def _terrones_arcilla():
    grad = {23: 400, 24: 800, 25: 600, 26: 700, 27: 500, 28: 600, 29: 400, 30: 300, 31: 250, 32: 200, 33: 150, 34: 100}
    d = {"ter_masa_total": _n(sum(grad.values()))}
    for k, v in grad.items():
        d[f"ter_grad_{k}"] = _n(v)
    for k, (ini, fin) in {18: (3000.0, 2950.0), 19: (3000.0, 2955.0), 20: (2000.0, 1976.0), 21: (1000.0, 988.0), 22: (500.0, 493.8)}.items():
        d[f"ter_{k}_ini"] = _n(ini)
        d[f"ter_{k}_fin"] = _n(fin)
    return d


def _alargamiento_aplanamiento():
    ini = (500, 1200, 1500, 800, 500, 300, 200)
    d = {}
    for clave, par in (("alar", (40, 90, 120, 60, 38, 24, 16)), ("apla", (35, 85, 125, 52, 34, 20, 14))):
        for i, (a, b) in enumerate(zip(ini, par), start=1):
            d[f"aa_{clave}_{i}_ini"] = _n(a, 1)
            d[f"aa_{clave}_{i}_par"] = _n(b, 1)
    return d


_GENERADORES = {
    "granulometria": _granulometria, "humedad": _humedad_natural, "masa-unitaria": _masa_unitaria, "limites": _limites,
    "pasa200": _pasa200, "cbr": _cbr, "corte-directo": _corte_directo, "gravedad-especifica": _gravedad_especifica,
    "proctor": _proctor, "materia-organica": _materia_organica, "limite-contraccion": _limite_contraccion,
    "consolidacion": _consolidacion, "compresion-inconfinada": _compresion_inconfinada, "compresion-roca": _compresion_roca,
    "carga-puntual": _carga_puntual, "solidez-sulfatos": _solidez_sulfatos, "terrones-arcilla": _terrones_arcilla,
    "alargamiento-aplanamiento": _alargamiento_aplanamiento,
}

# Condición del ensayo (temperatura y humedad del laboratorio) que se comparte entre los ensayos de una muestra.
CONDICION_EJEMPLO = {"cond_inicial_temp": "22.0", "cond_inicial_hum": "58", "cond_final_temp": "22.5", "cond_final_hum": "56"}


def datos_ejemplo(tipo):
    """Diccionario de datos de ejemplo del ensayo `tipo` ({} si no hay)."""
    gen = _GENERADORES.get(tipo)
    return gen() if gen else {}
