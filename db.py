"""
Capa de acceso a datos de Geodelta Lab contra Supabase.

Cada sesión de Streamlit tiene su propio cliente autenticado (nunca uno
compartido a nivel de módulo) para que las políticas RLS que dependen de
auth.uid() se apliquen a la persona correcta.
"""

import streamlit as st
from postgrest.exceptions import APIError
from supabase import Client, create_client

AUTH_DOMAIN = "geodelta-lab.local"  # dominio ficticio: la persona solo ve/escribe un código corto


class AuthError(Exception):
    pass


class PermissionError(Exception):
    pass


# ════════════════════════════════════════════════════════════════════
# CLIENTE / SESIÓN
# ════════════════════════════════════════════════════════════════════
def _new_client() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_ANON_KEY"])


def get_client() -> Client:
    if "sb_client" not in st.session_state:
        st.session_state.sb_client = _new_client()
    client = st.session_state.sb_client
    _refresh_if_needed(client)
    return client


def _refresh_if_needed(client: Client):
    """Solo refresca cuando el token está por vencer (no en cada llamada) — refrescar
    de más puede invalidar momentáneamente la sesión entre dos operaciones seguidas de
    la misma acción, y esa ventana hace que la siguiente llamada corra como anónimo
    (RLS la rechaza en vez de dar un error de autenticación claro).

    OJO: Supabase rota el refresh_token cada vez que se refresca la sesión (el viejo queda
    inválido) — si no se avisa a app.py para que reescriba la cookie del navegador con el
    token nuevo, la próxima vez que la persona recargue la página, la cookie tendrá un
    refresh_token ya vencido y la mandará de vuelta al login aunque haya seguido usando la
    app activamente. `_tokens_rotated` es la señal que app.py revisa en cada rerun para
    volver a guardar la cookie con los tokens vigentes (ver router principal)."""
    import time
    session = client.auth.get_session()
    if session is None:
        return
    expires_at = getattr(session, "expires_at", None)
    if expires_at is not None and expires_at - time.time() > 60:
        return
    try:
        client.auth.refresh_session()
        st.session_state["_tokens_rotated"] = True
    except Exception:
        sign_out()


def sign_in(codigo: str, password: str) -> dict:
    """Autentica con un código corto de usuario (no un correo real) y devuelve
    el perfil (nombre, rol) de la persona. Lanza AuthError si la clave es incorrecta."""
    client = get_client()
    email = f"{codigo.strip().lower()}@{AUTH_DOMAIN}"
    try:
        client.auth.sign_in_with_password({"email": email, "password": password})
    except Exception as e:
        print(f"[db.sign_in] fallo autenticando {email!r}: {e!r}")
        raise AuthError("Código de usuario o clave incorrectos.")

    profile = get_current_profile(client)
    if profile is None:
        client.auth.sign_out()
        raise AuthError("Tu cuenta no tiene un perfil asignado. Contacta al Jefe de laboratorio.")
    return profile


def sign_out():
    client = st.session_state.get("sb_client")
    if client is not None:
        try:
            client.auth.sign_out()
        except Exception:
            pass


def get_current_profile(client: Client = None) -> dict | None:
    client = client or get_client()
    user = client.auth.get_user()
    if user is None:
        return None
    res = client.table("profiles").select("id, full_name, role, active").eq("id", user.user.id).execute()
    rows = res.data or []
    if not rows or not rows[0]["active"]:
        return None
    return rows[0]


def get_session_tokens() -> dict | None:
    """access_token/refresh_token de la sesión actual, para guardarlos en una cookie del
    navegador y poder restaurar el login tras un recargo de página (ver restore_session)."""
    session = get_client().auth.get_session()
    if session is None:
        return None
    return {"access_token": session.access_token, "refresh_token": session.refresh_token}


def restore_session(access_token: str, refresh_token: str) -> dict | None:
    """Recrea la sesión a partir de tokens guardados en la cookie del navegador. Si el
    access_token venció, set_session lo refresca solo usando el refresh_token. Devuelve el
    perfil si funciona, o None si el refresh_token ya no sirve — ahí toca loguearse de nuevo."""
    client = get_client()
    try:
        client.auth.set_session(access_token, refresh_token)
    except Exception:
        return None
    return get_current_profile(client)


# ════════════════════════════════════════════════════════════════════
# PROYECTOS
# ════════════════════════════════════════════════════════════════════
def list_projects(include_archived: bool = False) -> list[dict]:
    q = get_client().table("projects").select("*")
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("created_at", desc=True).execute().data


def get_project(codigo_interno: str) -> dict | None:
    res = get_client().table("projects").select("*").eq("codigo_interno", codigo_interno).execute()
    rows = res.data or []
    return rows[0] if rows else None


def update_project(project_id: str, **fields) -> dict:
    res = get_client().table("projects").update(fields).eq("id", project_id).execute()
    return res.data[0]


def archive_project(project_id: str):
    """Archiva el proyecto y, en cascada explícita, sus perforaciones/muestras/assays."""
    client = get_client()
    client.table("projects").update({"archived": True}).eq("id", project_id).execute()
    perf_ids = [p["id"] for p in client.table("perforaciones").select("id").eq("project_id", project_id).execute().data]
    if perf_ids:
        client.table("perforaciones").update({"archived": True}).in_("id", perf_ids).execute()
        muestra_ids = [m["id"] for m in client.table("muestras").select("id").in_("perforacion_id", perf_ids).execute().data]
        if muestra_ids:
            client.table("muestras").update({"archived": True}).in_("id", muestra_ids).execute()
            client.table("assays").update({"archived": True}).in_("muestra_id", muestra_ids).execute()


def commit_new_project(project_fields: dict, perforaciones: list[dict]) -> str:
    """Crea proyecto + perforaciones + muestras en una sola transacción (función RPC
    commit_new_project en Postgres). `perforaciones` es una lista de
    {tipo, consecutivo, codigo, muestras: [...]}."""
    res = get_client().rpc("commit_new_project", {
        "project_fields": project_fields,
        "perforaciones": perforaciones,
    }).execute()
    return res.data


# ════════════════════════════════════════════════════════════════════
# PERFORACIONES
# ════════════════════════════════════════════════════════════════════
def list_perforaciones(project_id: str, include_archived: bool = False) -> list[dict]:
    q = get_client().table("perforaciones").select("*").eq("project_id", project_id)
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("consecutivo").execute().data


def create_perforacion(project_id: str, tipo: str, consecutivo: int, codigo: str) -> dict:
    res = get_client().table("perforaciones").insert({
        "project_id": project_id, "tipo": tipo, "consecutivo": consecutivo, "codigo": codigo,
    }).execute()
    return res.data[0]


def set_perforacion_fecha_ingreso(perforacion_id: str, fecha: str | None) -> dict:
    """Fecha de ingreso de las muestras de ESTA perforación ('YYYY-MM-DD', o None para volver a la del proyecto).
    Requiere la migración 0033 (columna perforaciones.fecha_ingreso)."""
    res = get_client().table("perforaciones").update({"fecha_ingreso": fecha or None}).eq("id", perforacion_id).execute()
    return res.data[0]


def set_perforacion_fecha_emision(perforacion_id: str, fecha: str | None) -> dict:
    """Día en que se descargó por última vez un Excel de resultados de ESTA perforación ('YYYY-MM-DD').
    Requiere la migración 0037 (columna perforaciones.fecha_emision)."""
    res = get_client().table("perforaciones").update({"fecha_emision": fecha or None}).eq("id", perforacion_id).execute()
    return res.data[0]


def set_perforacion_muestras_en_orden(perforacion_id: str, total: int | None) -> dict:
    """Cuántas muestras traía la orden de ESTA perforación, con o sin ensayo marcado (las que no tienen ensayo no se importan).
    Requiere la migración 0038 (columna perforaciones.muestras_en_orden)."""
    res = get_client().table("perforaciones").update({"muestras_en_orden": total}).eq("id", perforacion_id).execute()
    return res.data[0]


def set_perforacion_muestras_en_orden_por_codigo(project_id: str, codigo: str, total: int | None):
    """Igual que set_perforacion_muestras_en_orden pero por (proyecto, código), para justo después de crear el proyecto."""
    get_client().table("perforaciones").update({"muestras_en_orden": total}).eq("project_id", project_id).eq("codigo", codigo).execute()


def set_perforacion_fecha_ingreso_por_codigo(project_id: str, codigo: str, fecha: str | None):
    """Igual que set_perforacion_fecha_ingreso pero buscando la perforación por (proyecto, código) -- para justo
    después de crear el proyecto, cuando todavía no se conocen los ids de sus perforaciones."""
    get_client().table("perforaciones").update({"fecha_ingreso": fecha or None}).eq("project_id", project_id).eq("codigo", codigo).execute()


def restore_or_create_perforacion(project_id: str, tipo: str, consecutivo: int, codigo: str) -> dict:
    """Crea la perforación, o la reactiva si ya existe archivada con ese mismo código: el índice único
    (project_id, codigo) cuenta también las archivadas, así que borrar una perforación y volver a
    importar/crear la misma fallaba con un error de llave duplicada sin agregar nada."""
    client = get_client()
    rows = client.table("perforaciones").select("*").eq("project_id", project_id).eq("codigo", codigo).execute().data or []
    if rows:
        row = rows[0]
        if row.get("archived"):
            row = client.table("perforaciones").update({"archived": False, "tipo": tipo}).eq("id", row["id"]).execute().data[0]
        return row
    return create_perforacion(project_id, tipo, consecutivo, codigo)


def rename_perforacion(perforacion_id: str, codigo: str) -> dict:
    """Cambia el nombre (código) de una perforación. Las muestras y sus ensayos cuelgan de la perforación por su id, no por el nombre."""
    return get_client().table("perforaciones").update({"codigo": codigo}).eq("id", perforacion_id).execute().data[0]


def archive_perforacion(perforacion_id: str):
    client = get_client()
    client.table("perforaciones").update({"archived": True}).eq("id", perforacion_id).execute()
    muestra_ids = [m["id"] for m in client.table("muestras").select("id").eq("perforacion_id", perforacion_id).execute().data]
    if muestra_ids:
        client.table("muestras").update({"archived": True}).in_("id", muestra_ids).execute()
        client.table("assays").update({"archived": True}).in_("muestra_id", muestra_ids).execute()


# ════════════════════════════════════════════════════════════════════
# MUESTRAS
# ════════════════════════════════════════════════════════════════════
def list_muestras(perforacion_id: str, include_archived: bool = False) -> list[dict]:
    q = get_client().table("muestras").select("*").eq("perforacion_id", perforacion_id)
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("numero").execute().data


def get_muestra(muestra_id: str) -> dict | None:
    res = get_client().table("muestras").select("*").eq("id", muestra_id).execute()
    rows = res.data or []
    return rows[0] if rows else None


def create_muestra(perforacion_id: str, **fields) -> dict:
    res = get_client().table("muestras").insert({"perforacion_id": perforacion_id, **fields}).execute()
    return res.data[0]


def upsert_muestra(perforacion_id: str, id_unico: str, **fields) -> dict:
    """Crea la muestra, o la reactiva (con sus ensayos) si ya existe archivada con ese id_unico: los
    índices únicos (id_unico y perforacion_id + numero) cuentan también las archivadas, así que volver a
    agregar un número de muestra que se había borrado fallaba con una llave duplicada."""
    client = get_client()
    rows = client.table("muestras").select("*").eq("id_unico", id_unico).execute().data or []
    if not rows:
        return create_muestra(perforacion_id, id_unico=id_unico, **fields)
    row = rows[0]
    if row.get("archived"):
        client.table("assays").update({"archived": False}).eq("muestra_id", row["id"]).execute()
    return client.table("muestras").update({**fields, "perforacion_id": perforacion_id, "archived": False}).eq("id", row["id"]).execute().data[0]


def update_muestra(muestra_id: str, **fields) -> dict:
    res = get_client().table("muestras").update(fields).eq("id", muestra_id).execute()
    return res.data[0]


def asignar_ensayo(muestra_id: str, ensayo_label: str, laboratorista: str | None) -> dict:
    """Asigna (o desasigna, si laboratorista es None) un laboratorista a un ensayo puntual de
    una muestra — no a toda la muestra ni al proyecto entero. Se guarda en el jsonb
    muestras.ensayo_asignado_a en vez de en assays porque el Jefe debe poder asignar antes de
    que exista una fila de assay (esa solo se crea cuando alguien abre el formulario)."""
    muestra = get_muestra(muestra_id)
    asignaciones = dict(muestra.get("ensayo_asignado_a") or {})
    if laboratorista:
        asignaciones[ensayo_label] = laboratorista
    else:
        asignaciones.pop(ensayo_label, None)
    return update_muestra(muestra_id, ensayo_asignado_a=asignaciones)


def archive_muestra(muestra_id: str):
    client = get_client()
    client.table("muestras").update({"archived": True}).eq("id", muestra_id).execute()
    client.table("assays").update({"archived": True}).eq("muestra_id", muestra_id).execute()


# ════════════════════════════════════════════════════════════════════
# ENSAYOS (assays)
# ════════════════════════════════════════════════════════════════════
def get_assay(muestra_id: str, tipo: str) -> dict | None:
    res = get_client().table("assays").select("*").eq("muestra_id", muestra_id).eq("tipo", tipo).execute()
    rows = res.data or []
    return rows[0] if rows else None


def get_assay_fila(assay_id: str, columnas: str = "*") -> dict | None:
    """Una fila de ensayos por id, con las columnas pedidas (para partir de lo que de verdad está guardado)."""
    filas = get_client().table("assays").select(columnas).eq("id", assay_id).execute().data or []
    return filas[0] if filas else None


def create_assay(muestra_id: str, tipo: str, **defaults) -> dict:
    res = get_client().table("assays").insert({"muestra_id": muestra_id, "tipo": tipo, **defaults}).execute()
    return res.data[0]


def update_assay_data(assay_id: str, data: dict = None, observations: str = None,
                       laboratorist: str = None, status: str = None) -> dict:
    fields = {}
    if data is not None:
        fields["data"] = data
    if observations is not None:
        fields["observations"] = observations
    if laboratorist is not None:
        fields["laboratorist"] = laboratorist
    if status is not None:
        fields["status"] = status
    res = get_client().table("assays").update(fields).eq("id", assay_id).execute()
    return res.data[0]


def update_assay_shared_data(muestra_id: str, tipos: list[str], data: dict) -> None:
    """Actualiza el mismo `data` en varios tipos de ensayo a la vez — Pasa 200 comparte
    su formulario con Granulometría, así que guardar uno actualiza ambas filas."""
    get_client().table("assays").update({"data": data}).eq("muestra_id", muestra_id).in_("tipo", tipos).execute()


def _ahora_colombia() -> str:
    """Fecha y hora actual en hora de Colombia (UTC-5, sin horario de verano), con zona, en ISO. El servidor corre en UTC: con
    datetime.now() a secas las horas del historial salían 5 horas adelantadas."""
    from datetime import datetime, timedelta, timezone
    return datetime.now(timezone(timedelta(hours=-5))).isoformat()


def add_historial(assay_id: str, titulo: str, subtitulo: str = "", icono: str = "history", tono: str = "neutral",
                  cambios: list | None = None, snapshot: dict | None = None, nota: str | None = None):
    """Agrega una entrada al historial del ensayo (jsonb, sin migración: las claves nuevas son opcionales).
      cambios:  [{campo, etiqueta, antes, despues}] -- qué valores cambiaron respecto al envío anterior.
      snapshot: copia de los datos en este envío; sirve de base para el próximo diff. Solo se conserva la del ÚLTIMO envío
                (se borra de las entradas anteriores) para que el historial no crezca con una copia de los datos por envío.
      nota:     texto corto (ej. "no hay versión anterior para comparar").
    Un envío nuevo (snapshot) reemplaza la entrada "en curso" de ediciones posteriores al envío (ver set_edicion_en_curso):
    ese mismo diff queda ahora dentro de la entrada del reenvío."""
    from datetime import datetime
    client = get_client()
    assay = client.table("assays").select("historial").eq("id", assay_id).execute().data[0]
    previo = list(assay["historial"] or [])
    if snapshot is not None:
        previo = [{k: v for k, v in h.items() if k != "snapshot"} for h in previo if not h.get("en_curso")]
    entry = {"titulo": titulo, "subtitulo": subtitulo, "icono": icono, "tono": tono, "fecha": _ahora_colombia()}
    if cambios is not None:
        entry["cambios"] = cambios
    if snapshot is not None:
        entry["snapshot"] = snapshot
    if nota:
        entry["nota"] = nota
    client.table("assays").update({"historial": previo + [entry]}).eq("id", assay_id).execute()


def set_edicion_en_curso(assay_id: str, cambios: list, actor: str = ""):
    """Mantiene UNA entrada viva "Modificado después del envío" con el diff acumulado contra el último envío -- se actualiza
    sola mientras el laboratorista corrige (autoguardado) y desaparece si deja los valores como estaban. Así queda
    registrado lo que se cambió aunque nunca vuelva a darle "Enviar a revisión"."""
    from datetime import datetime
    client = get_client()
    assay = client.table("assays").select("historial").eq("id", assay_id).execute().data[0]
    historial = list(assay["historial"] or [])
    actual = next((h for h in historial if h.get("en_curso")), None)
    resto = [h for h in historial if not h.get("en_curso")]
    if (not cambios and actual is None) or (actual is not None and actual.get("cambios") == cambios):
        return   # nada que escribir: sin cambios y sin entrada viva, o el diff sigue igual
    if cambios:
        resto.append({"titulo": "Modificado después del envío", "subtitulo": actor, "icono": "edit_note", "tono": "warning",
                      "fecha": _ahora_colombia(), "cambios": cambios, "en_curso": True})
    client.table("assays").update({"historial": resto}).eq("id", assay_id).execute()


def registrar_ultimo_cambio(assay_id: str, actor: str):
    """Deja UNA entrada en el historial con el último cambio de datos hecho por el Jefe o el Director Técnico (quién y cuándo), también cuando el
    ensayo todavía es un borrador (ahí no hay envío con el que comparar y no se puede listar qué valores cambiaron). Cada cambio nuevo
    reemplaza la entrada anterior, así el historial no se llena con una línea por cada autoguardado."""
    client = get_client()
    assay = client.table("assays").select("historial").eq("id", assay_id).execute().data[0]
    resto = [h for h in (assay["historial"] or []) if not h.get("ultimo_cambio")]
    resto.append({"titulo": "Último cambio de datos", "subtitulo": actor, "icono": "edit_note", "tono": "warning",
                  "fecha": _ahora_colombia(), "ultimo_cambio": True})
    client.table("assays").update({"historial": resto}).eq("id", assay_id).execute()


def _require_role(profile: dict, *roles: str):
    if profile["role"] not in roles:
        raise PermissionError(f"Rol '{profile['role']}' no autorizado para esta acción.")


# Las 6 transiciones reales del flujo de confirmación (ver bloque "Ensayos asignados" en
# app.py, pantalla de detalle de muestra). Ninguna llama add_historial internamente — cada
# sitio de llamada en app.py ya registra su propia entrada de historial con texto específico
# para esa transición, justo después de invocar la función correspondiente.


def jefe_confirmar(assay_id: str, profile: dict):
    """etapa None -> pendiente_ing. El laboratorio ya terminó (Finalizado); el Jefe lo envía
    al Director Técnico para el visto bueno final."""
    _require_role(profile, "jefe")
    from datetime import datetime
    get_client().table("assays").update({
        "etapa_revision": "pendiente_ing",
        "confirmado_por_jefe_id": profile["id"],
        "confirmado_por_jefe_fecha": _ahora_colombia(),
        "motivo_rechazo": None, "rechazado_por": None,
    }).eq("id", assay_id).execute()


def jefe_devolver(assay_id: str, profile: dict, motivo: str):
    """etapa None -> sigue None, pero reabre el ensayo para el laboratorista (status
    vuelve a en-proceso) porque el Jefe encontró algo que corregir antes de confirmar."""
    _require_role(profile, "jefe")
    get_client().table("assays").update({
        "status": "en-proceso", "etapa_revision": None,
        "motivo_rechazo": motivo, "rechazado_por": "jefe",
    }).eq("id", assay_id).execute()


def jefe_desconfirmar(assay_id: str, profile: dict, motivo: str):
    """etapa pendiente_ing -> None. El Jefe retracta su confirmación antes de que el
    Director Técnico lo revise; el trabajo del laboratorista (status) no se toca."""
    _require_role(profile, "jefe")
    get_client().table("assays").update({
        "etapa_revision": None, "confirmado_por_jefe_id": None, "confirmado_por_jefe_fecha": None,
        "motivo_rechazo": motivo, "rechazado_por": "jefe",
    }).eq("id", assay_id).execute()


def ing_aprobar(assay_id: str, profile: dict):
    """etapa pendiente_ing -> aprobado. Visto bueno final del Director Técnico."""
    _require_role(profile, "ingeniero")
    from datetime import datetime
    get_client().table("assays").update({
        "etapa_revision": "aprobado",
        "aprobado_por_ing_id": profile["id"],
        "aprobado_por_ing_fecha": _ahora_colombia(),
        "motivo_rechazo": None, "rechazado_por": None,
    }).eq("id", assay_id).execute()


def ing_devolver(assay_id: str, profile: dict, motivo: str):
    """etapa pendiente_ing -> None. El Director Técnico lo devuelve al Jefe sin reabrir
    el trabajo de laboratorio (status no se toca)."""
    _require_role(profile, "ingeniero")
    get_client().table("assays").update({
        "etapa_revision": None, "motivo_rechazo": motivo, "rechazado_por": "ing",
    }).eq("id", assay_id).execute()


def ing_desconfirmar(assay_id: str, profile: dict, motivo: str):
    """etapa aprobado -> None. El Director Técnico retracta su aprobación final; reabre
    el ensayo para el laboratorista (status vuelve a en-proceso) y borra ambas
    confirmaciones (Jefe y Director Técnico) — tiene que volver a pasar por las dos."""
    _require_role(profile, "ingeniero")
    get_client().table("assays").update({
        "status": "en-proceso", "etapa_revision": None,
        "confirmado_por_jefe_id": None, "confirmado_por_jefe_fecha": None,
        "aprobado_por_ing_id": None, "aprobado_por_ing_fecha": None,
        "motivo_rechazo": motivo, "rechazado_por": "ing",
    }).eq("id", assay_id).execute()


def reset_confirmacion(assay_id: str, reset_status: bool = False):
    """Usado al desarchivar un proyecto ejecutado: borra toda la confirmación de un ensayo
    para que tenga que volver a pasar por Jefe y Director Técnico. Si `reset_status`, además
    reabre el ensayo (estaba 'finalizado') para que el laboratorista pueda volver a digitar."""
    fields = {
        "etapa_revision": None, "confirmado_por_jefe_id": None, "confirmado_por_jefe_fecha": None,
        "aprobado_por_ing_id": None, "aprobado_por_ing_fecha": None,
        "motivo_rechazo": None, "rechazado_por": None,
    }
    if reset_status:
        fields["status"] = "en-proceso"
    get_client().table("assays").update(fields).eq("id", assay_id).execute()


# ════════════════════════════════════════════════════════════════════
# LECTURA MASIVA (para reconstruir projects/perforaciones/muestras/assays
# en st.session_state en cada rerun — ver _load_data() en app.py)
# ════════════════════════════════════════════════════════════════════
def list_all_perforaciones(include_archived: bool = False) -> list[dict]:
    q = get_client().table("perforaciones").select("*")
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("consecutivo").execute().data


def list_all_muestras(include_archived: bool = False) -> list[dict]:
    q = get_client().table("muestras").select("*")
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("numero").execute().data


# Columnas pesadas de los ensayos (los datos digitados con sus lecturas de máquina y fotos, y el historial con sus copias): traerlas de TODOS los
# ensayos en cada recarga era lo que hacía lenta la navegación. Se guardan en memoria por (id, updated_at) y solo se vuelven a pedir las de los
# ensayos que cambiaron desde la última vez (updated_at lo actualiza un trigger en cada cambio).
_COLUMNAS_PESADAS = ("data", "historial")
_CACHE_PESADO: dict = {}
_CACHE_PESADO_MAX = 2500
_COLUMNAS_LIVIANAS: list | None = None


def _guardar_pesado(fila: dict):
    if len(_CACHE_PESADO) >= _CACHE_PESADO_MAX:
        for clave in list(_CACHE_PESADO)[: _CACHE_PESADO_MAX // 5]:   # saca los más viejos
            _CACHE_PESADO.pop(clave, None)
    _CACHE_PESADO[(fila["id"], fila["updated_at"])] = {c: fila.get(c) for c in _COLUMNAS_PESADAS if c in fila}


def list_all_assays(include_archived: bool = False) -> list[dict]:
    """Todos los ensayos. La primera vez (en el proceso) se piden completos para conocer las columnas; después solo las livianas, y las
    pesadas (data, historial) solo de los ensayos que cambiaron."""
    global _COLUMNAS_LIVIANAS
    import copy
    client = get_client()

    def _consulta(columnas):
        q = client.table("assays").select(columnas)
        if not include_archived:
            q = q.eq("archived", False)
        return q.order("created_at").execute().data

    if not _COLUMNAS_LIVIANAS:
        filas = _consulta("*")
        if filas:
            _COLUMNAS_LIVIANAS = [c for c in filas[0] if c not in _COLUMNAS_PESADAS]
        for f in filas:
            _guardar_pesado(f)
        return filas
    filas = _consulta(",".join(_COLUMNAS_LIVIANAS))
    faltan = []
    for f in filas:
        pesado = _CACHE_PESADO.get((f["id"], f["updated_at"]))
        if pesado is None:
            faltan.append(f["id"])
        else:
            f.update(copy.deepcopy(pesado))   # copia: cada sesión tiene sus propios diccionarios
    if faltan:
        por_id = {f["id"]: f for f in filas}
        for i in range(0, len(faltan), 40):
            lote = faltan[i:i + 40]
            for p in client.table("assays").select("id,updated_at," + ",".join(_COLUMNAS_PESADAS)).in_("id", lote).execute().data:
                _guardar_pesado(p)
                destino = por_id.get(p["id"])
                if destino is not None:
                    destino.update({c: p.get(c) for c in _COLUMNAS_PESADAS if c in p})
                    destino["updated_at"] = p["updated_at"]
    for f in filas:   # por si un ensayo desapareció entre las dos consultas
        for c in _COLUMNAS_PESADAS:
            f.setdefault(c, {} if c == "data" else [])
    return filas


def list_profiles() -> list[dict]:
    return get_client().table("profiles").select("id, full_name, role, active").execute().data


# ════════════════════════════════════════════════════════════════════
# ADMINISTRACIÓN DE USUARIOS (solo Jefe de Laboratorio)
# Crear una cuenta de acceso en Supabase Auth exige la "service role" key (la anon key, que usa la app para todo lo demás, no puede).
# Esa llave vive SOLO en los secretos del servidor (SUPABASE_SERVICE_ROLE_KEY), nunca se guarda en session_state ni llega al navegador, y
# cada función de abajo comprueba primero, con el cliente de la propia sesión, que quien la invoca es un Jefe de Laboratorio activo.
# ════════════════════════════════════════════════════════════════════
ROLES_VALIDOS = ("laboratorista", "ingeniero", "jefe")


class AdminNoConfigurado(Exception):
    pass


def admin_configurado() -> bool:
    try:
        return bool(st.secrets.get("SUPABASE_SERVICE_ROLE_KEY"))
    except Exception:
        return False


def _admin_client() -> Client:
    if not admin_configurado():
        raise AdminNoConfigurado("Falta SUPABASE_SERVICE_ROLE_KEY en los secretos de la app.")
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_SERVICE_ROLE_KEY"])


def _exigir_jefe() -> dict:
    perfil = get_current_profile()
    if not perfil or perfil.get("role") != "jefe":
        raise PermissionError("Solo el Jefe de Laboratorio puede administrar usuarios.")
    return perfil


def codigo_a_email(codigo: str) -> str:
    return f"{codigo.strip().lower()}@{AUTH_DOMAIN}"


def admin_listar_usuarios() -> list[dict]:
    """[{id, codigo, full_name, role, active}] de todas las cuentas (el código sale del correo ficticio de la cuenta)."""
    _exigir_jefe()
    admin = _admin_client()
    correos = {}
    pagina = 1
    while True:
        usuarios = admin.auth.admin.list_users(page=pagina, per_page=200)
        if not usuarios:
            break
        for u in usuarios:
            correos[str(u.id)] = (u.email or "").split("@")[0]
        if len(usuarios) < 200:
            break
        pagina += 1
    perfiles = admin.table("profiles").select("id, full_name, role, active").execute().data or []
    return sorted(({"id": p["id"], "codigo": correos.get(str(p["id"]), "—"), "full_name": p["full_name"], "role": p["role"],
                    "active": p["active"]} for p in perfiles), key=lambda u: (not u["active"], u["full_name"].lower()))


def admin_crear_usuario(codigo: str, full_name: str, role: str, password: str) -> dict:
    perfil = _exigir_jefe()
    import re as _re
    codigo = codigo.strip().lower()
    if not _re.fullmatch(r"[a-z0-9][a-z0-9._-]{2,29}", codigo):
        raise ValueError("El código debe tener de 3 a 30 caracteres: letras, números, punto, guion o guion bajo (sin espacios ni tildes).")
    if not full_name.strip():
        raise ValueError("Escribe el nombre completo.")
    if role not in ROLES_VALIDOS:
        raise ValueError("Rol no válido.")
    if len(password) < 8:
        raise ValueError("La clave debe tener al menos 8 caracteres.")
    admin = _admin_client()
    try:
        res = admin.auth.admin.create_user({"email": codigo_a_email(codigo), "password": password, "email_confirm": True})
    except Exception as e:
        if "already" in str(e).lower() or "registered" in str(e).lower() or "exists" in str(e).lower():
            raise ValueError(f"Ya existe un usuario con el código «{codigo}».")
        raise
    uid = str(res.user.id)
    try:
        admin.table("profiles").insert({"id": uid, "full_name": full_name.strip(), "role": role, "active": True}).execute()
    except Exception:
        # Si no se pudo crear el perfil no se deja una cuenta de acceso huérfana.
        try:
            admin.auth.admin.delete_user(uid)
        except Exception:
            pass
        raise
    return {"id": uid, "codigo": codigo, "full_name": full_name.strip(), "role": role, "active": True, "creado_por": perfil["id"]}


def _otros_jefes_activos(admin, excepto_id: str) -> int:
    filas = admin.table("profiles").select("id").eq("role", "jefe").eq("active", True).neq("id", excepto_id).execute().data or []
    return len(filas)


def admin_cambiar_activo(user_id: str, activo: bool):
    """Da de baja (activo=False) o reactiva una cuenta: quien está de baja ya no puede iniciar sesión, pero su historial se conserva."""
    perfil = _exigir_jefe()
    if not activo:
        if str(user_id) == str(perfil["id"]):
            raise ValueError("No puedes dar de baja tu propia cuenta.")
    admin = _admin_client()
    if not activo and _otros_jefes_activos(admin, str(user_id)) == 0:
        objetivo = admin.table("profiles").select("role").eq("id", user_id).execute().data or []
        if objetivo and objetivo[0]["role"] == "jefe":
            raise ValueError("Tiene que quedar al menos un Jefe de Laboratorio activo.")
    admin.table("profiles").update({"active": activo}).eq("id", user_id).execute()


def admin_restablecer_clave(user_id: str, nueva_clave: str):
    _exigir_jefe()
    if len(nueva_clave) < 8:
        raise ValueError("La clave debe tener al menos 8 caracteres.")
    _admin_client().auth.admin.update_user_by_id(str(user_id), {"password": nueva_clave})


def admin_eliminar_usuario(user_id: str) -> str:
    """Elimina la cuenta. Devuelve 'eliminado' o, si la persona ya figura en el historial de ensayos o proyectos (no se puede borrar sin
    romper ese historial), 'baja': queda dada de baja (no puede entrar) y su nombre se conserva donde ya aparece."""
    perfil = _exigir_jefe()
    if str(user_id) == str(perfil["id"]):
        raise ValueError("No puedes eliminar tu propia cuenta.")
    admin = _admin_client()
    if _otros_jefes_activos(admin, str(user_id)) == 0:
        objetivo = admin.table("profiles").select("role").eq("id", user_id).execute().data or []
        if objetivo and objetivo[0]["role"] == "jefe":
            raise ValueError("Tiene que quedar al menos un Jefe de Laboratorio activo.")
    try:
        admin.auth.admin.delete_user(str(user_id))
        return "eliminado"
    except Exception:
        admin.table("profiles").update({"active": False}).eq("id", user_id).execute()
        return "baja"


# ════════════════════════════════════════════════════════════════════
# NOTIFICACIONES — codigo_interno/perforacion_codigo/muestra_id_unico son
# identificadores de negocio en texto plano, no FKs: una notificación es
# informativa/efímera, no necesita integridad referencial con las otras tablas.
# ════════════════════════════════════════════════════════════════════
def add_notification(target_role: str, mensaje: str, codigo_interno: str = None,
                      perforacion_codigo: str = None, muestra_id_unico: str = None):
    """Vía RPC (función security definer), no insert directo — ver migración 0012."""
    get_client().rpc("add_notification", {
        "p_target_role": target_role, "p_mensaje": mensaje, "p_codigo_interno": codigo_interno,
        "p_perforacion_codigo": perforacion_codigo, "p_muestra_id_unico": muestra_id_unico,
    }).execute()


def list_notifications(role: str, limit: int = 50) -> list[dict]:
    res = (get_client().table("notifications").select("*")
           .eq("target_role", role).order("fecha", desc=True).limit(limit).execute())
    return res.data


def mark_notification_read(notification_id: str):
    get_client().table("notifications").update({"leida": True}).eq("id", notification_id).execute()


def mark_all_notifications_read(role: str):
    get_client().table("notifications").update({"leida": True}).eq("target_role", role).eq("leida", False).execute()


# ════════════════════════════════════════════════════════════════════
# COMPROBACIÓN DE BALANZAS (jefe de laboratorio) — GDA-FLC-029
# ════════════════════════════════════════════════════════════════════
def list_balance_checks() -> list[dict]:
    res = get_client().table("balance_checks").select("*").order("semana_lunes", desc=True).execute()
    return res.data or []


def get_balance_check(check_id: str) -> dict | None:
    res = get_client().table("balance_checks").select("*").eq("id", check_id).execute()
    rows = res.data or []
    return rows[0] if rows else None


def create_balance_check(codigo_equipo: str, semana_lunes: str, fecha_comprobacion: str = None,
                          fecha_proxima: str = None, data: dict = None) -> dict:
    fields = {"codigo_equipo": codigo_equipo, "semana_lunes": semana_lunes, "data": data or {}}
    if fecha_comprobacion is not None:
        fields["fecha_comprobacion"] = fecha_comprobacion
    if fecha_proxima is not None:
        fields["fecha_proxima"] = fecha_proxima
    profile = get_current_profile()
    if profile:
        fields["created_by"] = profile["id"]
    res = get_client().table("balance_checks").insert(fields).execute()
    return res.data[0]


def update_balance_check(check_id: str, data: dict = None, estado: str = None,
                          fecha_comprobacion: str = None, fecha_proxima: str = None,
                          semana_lunes: str = None) -> dict:
    fields = {}
    if data is not None:
        fields["data"] = data
    if estado is not None:
        fields["estado"] = estado
    if fecha_comprobacion is not None:
        fields["fecha_comprobacion"] = fecha_comprobacion
    if fecha_proxima is not None:
        fields["fecha_proxima"] = fecha_proxima
    if semana_lunes is not None:
        fields["semana_lunes"] = semana_lunes
    res = get_client().table("balance_checks").update(fields).eq("id", check_id).execute()
    return res.data[0]


def delete_balance_check(check_id: str):
    get_client().table("balance_checks").delete().eq("id", check_id).execute()


def get_balance_control() -> list:
    """Filas del tablero "Control semanal" (una sola lista compartida, no una por registro)."""
    res = get_client().table("balance_control").select("rows").eq("id", 1).execute()
    rows = res.data or []
    return rows[0]["rows"] if rows else []


def save_balance_control(rows: list):
    get_client().table("balance_control").upsert({"id": 1, "rows": rows}).execute()


# ════════════════════════════════════════════════════════════════════
# CATÁLOGO DE BALANZAS (Comprobación de Balanzas)
# ════════════════════════════════════════════════════════════════════
def list_balanzas() -> list:
    res = get_client().table("balanzas").select("*").order("codigo").execute()
    return res.data or []


def create_balanza(codigo: str, nombre: str, marca: str, serie: str, resolucion: str) -> dict:
    res = get_client().table("balanzas").insert({
        "codigo": codigo, "nombre": nombre, "marca": marca, "serie": serie, "resolucion": resolucion,
    }).execute()
    return res.data[0]


def update_balanza(balanza_id: str, **fields) -> dict:
    res = get_client().table("balanzas").update(fields).eq("id", balanza_id).execute()
    return res.data[0]


def delete_balanza(balanza_id: str):
    get_client().table("balanzas").delete().eq("id", balanza_id).execute()
