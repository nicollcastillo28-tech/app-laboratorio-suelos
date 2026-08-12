-- Geodelta Lab — reactiva RLS en notifications y mueve la escritura a una
-- función security definer (RPC), en vez de depender del INSERT directo vía
-- PostgREST que fallaba con RLS incluso con una política "to public with
-- check (true)" — la causa exacta no se identificó, pero este patrón es el
-- estándar para casos así y permite reactivar RLS con confianza.
alter table notifications enable row level security;

drop policy if exists notifications_insert_any on notifications;

create or replace function add_notification(
  p_target_role role_enum,
  p_mensaje text,
  p_codigo_interno text default null,
  p_perforacion_codigo text default null,
  p_muestra_id_unico text default null
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into notifications (target_role, mensaje, codigo_interno, perforacion_codigo, muestra_id_unico)
  values (p_target_role, p_mensaje, p_codigo_interno, p_perforacion_codigo, p_muestra_id_unico);
end;
$$;

grant execute on function add_notification(role_enum, text, text, text, text) to authenticated;

-- Sin política de INSERT en la tabla a propósito: los inserts directos quedan
-- bloqueados por RLS, todo tiene que pasar por la función de arriba.
select policyname, cmd, roles from pg_policies where tablename = 'notifications';
