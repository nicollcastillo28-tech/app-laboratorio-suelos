-- Geodelta Lab — Row Level Security.
-- Ejecutar después de 0001_schema.sql.

alter table profiles       enable row level security;
alter table projects       enable row level security;
alter table perforaciones  enable row level security;
alter table muestras       enable row level security;
alter table assays         enable row level security;
alter table notifications  enable row level security;

-- ── profiles ──────────────────────────────────────────────────────────
-- Cualquier usuario autenticado puede leer todos los perfiles (para
-- mostrar nombres reales en el historial de confirmación). Sin
-- INSERT/UPDATE propio: el aprovisionamiento de usuarios es manual
-- (panel de Supabase o script con la service-role key).
create policy profiles_select_all on profiles for select
  to authenticated using (true);

-- ── projects / perforaciones / muestras ─────────────────────────────
-- Lectura: cualquier autenticado (igual al store compartido de hoy).
-- Escritura (insert/update/delete): solo rol "jefe".
create policy projects_select_all on projects for select
  to authenticated using (true);
create policy projects_jefe_write on projects for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));

create policy perforaciones_select_all on perforaciones for select
  to authenticated using (true);
create policy perforaciones_jefe_write on perforaciones for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));

create policy muestras_select_all on muestras for select
  to authenticated using (true);
create policy muestras_jefe_write on muestras for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));

-- ── assays ───────────────────────────────────────────────────────────
-- Lectura: cualquier autenticado. Escritura: cualquiera de los 3 roles
-- (la regla fina de quién puede tocar qué campo en qué estado vive en
-- db.py, igual que hoy vive en Python — ver plan de migración).
create policy assays_select_all on assays for select
  to authenticated using (true);
create policy assays_write on assays for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role in ('jefe','laboratorista','ingeniero')))
  with check (exists (select 1 from profiles where id = auth.uid() and role in ('jefe','laboratorista','ingeniero')));

-- ── notifications ────────────────────────────────────────────────────
-- Lectura filtrada por el rol del usuario autenticado (reemplaza el
-- filtro `n["role"] == st.session_state.role` de hoy).
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
