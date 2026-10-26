-- Historial de versiones de los ensayos: cada vez que cambia `assays.data`, la versión ANTERIOR queda guardada aquí
-- (la guarda la base con un trigger, así que cubre autoguardado, botones y cualquier otro dispositivo).
-- Se conservan las últimas 40 versiones de cada ensayo. Las fotos (b64) no se copian para no abultar la tabla.

create table if not exists public.assay_versions (
  id          bigint generated always as identity primary key,
  assay_id    uuid not null references public.assays(id) on delete cascade,
  data        jsonb not null,
  status      text,
  saved_at    timestamptz not null default now()   -- cuándo se reemplazó esta versión
);
create index if not exists assay_versions_assay_idx on public.assay_versions (assay_id, saved_at desc);

create or replace function public.guardar_version_ensayo() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  viejo jsonb;
  nuevo jsonb;
begin
  select coalesce(jsonb_object_agg(k, v), '{}'::jsonb) into viejo
    from jsonb_each(coalesce(old.data, '{}'::jsonb)) as t(k, v)
   where k <> '_respaldo' and not (jsonb_typeof(v) = 'object' and v ? 'b64');
  select coalesce(jsonb_object_agg(k, v), '{}'::jsonb) into nuevo
    from jsonb_each(coalesce(new.data, '{}'::jsonb)) as t(k, v)
   where k <> '_respaldo' and not (jsonb_typeof(v) = 'object' and v ? 'b64');
  if viejo is distinct from nuevo and viejo <> '{}'::jsonb then
    insert into public.assay_versions (assay_id, data, status) values (old.id, viejo, old.status);
    delete from public.assay_versions
     where assay_id = old.id
       and id not in (select id from public.assay_versions where assay_id = old.id order by saved_at desc, id desc limit 40);
  end if;
  return new;
end $$;

drop trigger if exists assays_guardar_version on public.assays;
create trigger assays_guardar_version after update of data on public.assays
  for each row execute function public.guardar_version_ensayo();

alter table public.assay_versions enable row level security;
drop policy if exists assay_versions_read on public.assay_versions;
create policy assay_versions_read on public.assay_versions for select to authenticated using (true);
