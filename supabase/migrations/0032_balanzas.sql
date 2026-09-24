-- Geodelta Lab: catálogo de balanzas de la Comprobación de Balanzas (antes fijo en el código).
-- Una balanza "de baja" (activa = false) deja de ofrecerse para registros nuevos, pero sus registros
-- históricos se siguen pudiendo consultar y descargar.
create table balanzas (
  id          uuid primary key default gen_random_uuid(),
  codigo      text not null unique,                 -- código interno, ej. "GDA-E-010"
  nombre      text not null,
  marca       text not null default '',
  serie       text not null default '',
  resolucion  text not null default '',             -- d (g), tal como se digita: "0,01"
  activa      boolean not null default true,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create trigger balanzas_set_updated_at before update on balanzas
  for each row execute function set_updated_at();

alter table balanzas enable row level security;
create policy balanzas_jefe_all on balanzas for all
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'))
  with check (exists (select 1 from profiles where id = auth.uid() and role = 'jefe'));

-- Las 4 balanzas con las que ya se venía trabajando.
insert into balanzas (codigo, nombre, marca, serie, resolucion) values
  ('GDA-E-010', 'Balanza Cap. 600g',             'TRUMAX', 'MIX-H',    '0,01'),
  ('GDA-E-011', 'Balanza Cap. 3200 g Pionner',   'OHAUS',  'Px3202/E', '0,01'),
  ('GDA-E-012', 'Balanza Cap. 30 Kg',            'TRUMAX', 'FENIX',    '1'),
  ('GDA-E-013', 'Balanza Cap. 3000 g',           'TS',     'T200',     '0,1')
on conflict (codigo) do nothing;
