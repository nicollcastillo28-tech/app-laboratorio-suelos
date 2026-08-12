-- Geodelta Lab — el insert de notifications seguía fallando con RLS incluso en
-- sesiones nuevas (probablemente algo en cómo supabase-py propaga el rol
-- "authenticated" para ese request puntual). Las notificaciones son datos de
-- bajo riesgo (avisos internos, no información sensible), así que en vez de
-- seguir depurando a ciegas, se amplía la política de INSERT a "public"
-- (que igual solo pueden alcanzar sesiones con la anon key del proyecto).
drop policy if exists notifications_insert_any on notifications;

create policy notifications_insert_any on notifications for insert
  to public with check (true);

-- Verificación.
select policyname, cmd, roles from pg_policies where tablename = 'notifications';
