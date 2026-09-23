-- Geodelta Lab: Comprobación intermedia de balanzas (jefe de laboratorio) — GDA-FLC-029.
-- Un registro por equipo/semana; el resto de campos (excentricidad, repetibilidad, exactitud,
-- patrones, condiciones, decisión y firmas) vive en `data` (jsonb), igual que `assays.data`.
create table balance_checks (
  id                  uuid primary key default gen_random_uuid(),
  codigo_equipo       text not null,                        -- "GDA-E-010", etc.
  semana_lunes        date not null,                         -- lunes de la semana ISO del registro
  fecha_comprobacion  date,
  fecha_proxima       date,
  estado              text not null default 'pendiente'
                         check (estado in ('pendiente', 'apto', 'no-apto')),
  data                jsonb not null default '{}'::jsonb,
  created_by          uuid references profiles(id),
  created_at          timestamptz not null default now(),
  updated_at          timestamptz not null default now()
);
create index balance_checks_equipo_semana_idx on balance_checks (codigo_equipo, semana_lunes desc);

create trigger balance_checks_set_updated_at before update on balance_checks
  for each row execute function set_updated_at();

-- Solo el Jefe de Laboratorio usa este módulo (ni lectura para los demás roles).
alter table balance_checks enable row level security;
create policy balance_checks_jefe_all on balance_checks for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));
