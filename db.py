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
    (RLS la rechaza en vez de dar un error de autenticación claro)."""
    import time
    session = client.auth.get_session()
    if session is None:
        return
    expires_at = getattr(session, "expires_at", None)
    if expires_at is not None and expires_at - time.time() > 60:
        return
    try:
        client.auth.refresh_session()
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


def update_muestra(muestra_id: str, **fields) -> dict:
    res = get_client().table("muestras").update(fields).eq("id", muestra_id).execute()
    return res.data[0]


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


def add_historial(assay_id: str, titulo: str, subtitulo: str = "", icono: str = "history", tono: str = "neutral"):
    from datetime import datetime
    client = get_client()
    assay = client.table("assays").select("historial").eq("id", assay_id).execute().data[0]
    entry = {"titulo": titulo, "subtitulo": subtitulo, "icono": icono, "tono": tono,
              "fecha": datetime.now().isoformat()}
    nuevo_historial = (assay["historial"] or []) + [entry]
    client.table("assays").update({"historial": nuevo_historial}).eq("id", assay_id).execute()


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
        "confirmado_por_jefe_fecha": datetime.now().isoformat(),
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
        "aprobado_por_ing_fecha": datetime.now().isoformat(),
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


def list_all_assays(include_archived: bool = False) -> list[dict]:
    q = get_client().table("assays").select("*")
    if not include_archived:
        q = q.eq("archived", False)
    return q.order("created_at").execute().data


def list_profiles() -> list[dict]:
    return get_client().table("profiles").select("id, full_name, role, active").execute().data


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
