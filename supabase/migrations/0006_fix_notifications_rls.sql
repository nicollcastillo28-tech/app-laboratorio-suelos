-- Geodelta Lab — recrea las políticas de notifications desde cero (diagnóstico:
-- el insert estaba fallando con "new row violates row-level security policy").
drop policy if exists notifications_select_own_role on notifications;
drop policy if exists notifications_insert_any on notifications;
drop policy if exists notifications_update_own_role on notifications;

create policy notifications_select_own_role on notifications for select
  to authenticated using (
    target_role = (select role from profiles where id = auth.uid())
  );

create policy notifications_insert_any on notifications for insert
  to authenticated with check (true);

create policy notifications_update_own_role on notifications for update
  to authenticated using (
    target_role = (select role from profiles where id = auth.uid())
  );

-- Verificación: debe listar las 3 políticas de arriba.
select policyname, cmd, roles from pg_policies where tablename = 'notifications';
