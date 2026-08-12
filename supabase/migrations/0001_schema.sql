-- Geodelta Lab — esquema inicial para la migración a Supabase.
-- Ejecutar completo en el SQL Editor del proyecto Supabase (una sola vez).

create extension if not exists pgcrypto; -- gen_random_uuid()

create type role_enum as enum ('jefe', 'laboratorista', 'ingeniero');

-- ════════════════════════════════════════════════════════════════════
-- PERFILES (vincula auth.users con el rol de laboratorio de la persona)
-- ════════════════════════════════════════════════════════════════════
create table profiles (
  id          uuid primary key references auth.users(id) on delete cascade,
  full_name   text not null,
  role        role_enum not null,
  active      boolean not null default true,
  created_at  timestamptz not null default now()
);

-- ════════════════════════════════════════════════════════════════════
-- PROYECTOS
-- ════════════════════════════════════════════════════════════════════
create table projects (
  id                     uuid primary key default gen_random_uuid(),
  codigo_interno         text not null unique,           -- "GDA-001-24", campo de negocio (ya no es llave)
  numero                 text not null,
  anio                   text not null,
  nombre                 text not null,
  localizacion           text,
  norma                  text,                            -- IDU / NTC / INVIAS / Otro
  fecha_bitacora         date,
  fecha_ingreso_muestra  date,
  laboratorista_asignado text,
  cliente                text,
  correo_cliente         text,
  muestra_tomada_por     text,
  direccion_cliente      text,
  telefono_contacto      text,
  nombre_contacto        text,
  fecha_inicio_proyecto  date,
  fecha_final_proyecto   date,
  fecha_recepcion        date,
  fecha_ejecucion        date,
  fecha_emision          date,
  archived               boolean not null default false,
  created_by             uuid references profiles(id),
  created_at             timestamptz not null default now(),
  updated_at             timestamptz not null default now()
);
create unique index projects_numero_anio_uk on projects (numero, anio);
create index projects_archived_idx on projects (archived);

-- ════════════════════════════════════════════════════════════════════
-- PERFORACIONES
-- ════════════════════════════════════════════════════════════════════
create table perforaciones (
  id           uuid primary key default gen_random_uuid(),
  project_id   uuid not null references projects(id) on delete cascade,
  tipo         text not null check (tipo in ('Sondeo','Apique','Fuente/Cantera')),
  consecutivo  int  not null,
  codigo       text not null,                             -- "S1", "AP2"
  archived     boolean not null default false,
  created_at   timestamptz not null default now()
);
create unique index perforaciones_project_codigo_uk on perforaciones (project_id, codigo);
create index perforaciones_project_idx on perforaciones (project_id);

-- ════════════════════════════════════════════════════════════════════
-- MUESTRAS
-- ════════════════════════════════════════════════════════════════════
create table muestras (
  id                  uuid primary key default gen_random_uuid(),
  perforacion_id      uuid not null references perforaciones(id) on delete cascade,
  numero              text not null,
  id_unico            text not null unique,               -- "GDA-001-24-S1-M1", id legible histórico
  profundidad_de      numeric,
  profundidad_hasta   numeric,
  tipo_muestra        text,
  ensayos             jsonb not null default '{}'::jsonb,  -- {"Granulometría": true, ...}
  observaciones       text,
  cond_inicial_temp   text,
  cond_inicial_hum    text,
  cond_final_temp     text,
  cond_final_hum      text,
  archived            boolean not null default false,
  created_at          timestamptz not null default now(),
  updated_at          timestamptz not null default now()
);
create unique index muestras_perforacion_numero_uk on muestras (perforacion_id, numero);
create index muestras_perforacion_idx on muestras (perforacion_id);

-- ════════════════════════════════════════════════════════════════════
-- ENSAYOS (assays)
-- ════════════════════════════════════════════════════════════════════
create table assays (
  id                          uuid primary key default gen_random_uuid(),
  muestra_id                  uuid not null references muestras(id) on delete cascade,
  tipo                        text not null check (tipo in ('granulometria','humedad','masa-unitaria','limites','pasa200')),
  status                      text not null default 'sin-iniciar'
                                 check (status in ('sin-iniciar','en-proceso','finalizado')),
  data                        jsonb not null default '{}'::jsonb,
  observations                text,
  laboratorist                text,
  etapa_revision              text check (etapa_revision in ('pendiente_ing','aprobado')),
  confirmado_por_jefe_id      uuid references profiles(id),
  confirmado_por_jefe_fecha   timestamptz,
  aprobado_por_ing_id         uuid references profiles(id),
  aprobado_por_ing_fecha      timestamptz,
  motivo_rechazo              text,
  rechazado_por               text check (rechazado_por in ('jefe','ing')),
  historial                   jsonb not null default '[]'::jsonb,
  archived                    boolean not null default false,
  created_at                  timestamptz not null default now(),
  updated_at                  timestamptz not null default now()
);
create unique index assays_muestra_tipo_uk on assays (muestra_id, tipo);
create index assays_muestra_idx on assays (muestra_id);
create index assays_pendiente_ing_idx on assays (etapa_revision) where etapa_revision = 'pendiente_ing';

-- ════════════════════════════════════════════════════════════════════
-- NOTIFICACIONES
-- ════════════════════════════════════════════════════════════════════
create table notifications (
  id                  uuid primary key default gen_random_uuid(),
  target_role         role_enum not null,
  mensaje             text not null,
  leida               boolean not null default false,
  fecha               timestamptz not null default now(),
  project_id          uuid references projects(id) on delete set null,
  perforacion_codigo  text,
  muestra_id          uuid references muestras(id) on delete set null,
  codigo_interno      text
);
create index notifications_target_role_idx on notifications (target_role, leida);

-- ════════════════════════════════════════════════════════════════════
-- updated_at automático
-- ════════════════════════════════════════════════════════════════════
create or replace function set_updated_at() returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger projects_set_updated_at before update on projects
  for each row execute function set_updated_at();
create trigger muestras_set_updated_at before update on muestras
  for each row execute function set_updated_at();
create trigger assays_set_updated_at before update on assays
  for each row execute function set_updated_at();
