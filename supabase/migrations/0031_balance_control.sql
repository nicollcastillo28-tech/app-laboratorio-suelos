-- Geodelta Lab: tablero "Control semanal" de balanzas (GDA-FLC-029). Es una sola lista compartida
-- entre todos los registros de comprobación (no una por registro), así que vive en una fila única.
create table balance_control (
  id          int primary key default 1 check (id = 1),
  rows        jsonb not null default '[]'::jsonb,
  updated_at  timestamptz not null default now()
);

create trigger balance_control_set_updated_at before update on balance_control
  for each row execute function set_updated_at();

alter table balance_control enable row level security;
create policy balance_control_jefe_all on balance_control for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));
