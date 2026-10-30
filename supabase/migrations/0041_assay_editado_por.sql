-- Quién digitó cada cambio de un ensayo.
--  * assays.editado_por: nombre (y rol) de la persona que hizo el ÚLTIMO cambio de datos. La app lo escribe en cada guardado con cambios;
--    entrar a mirar un ensayo ya no cambia «quién fue el último».
--  * assay_versions.digitado_por / digitado_en: de quién y de cuándo era el estado que se guardó como versión anterior (el que tenía el
--    ensayo justo antes de ser reemplazado).

alter table public.assays add column if not exists editado_por text;
alter table public.assay_versions add column if not exists digitado_por text;
alter table public.assay_versions add column if not exists digitado_en timestamptz;

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
    insert into public.assay_versions (assay_id, data, status, digitado_por, digitado_en)
    values (old.id, viejo, old.status, coalesce(old.editado_por, old.laboratorist), old.updated_at);
    delete from public.assay_versions
     where assay_id = old.id
       and id not in (select id from public.assay_versions where assay_id = old.id order by saved_at desc, id desc limit 40);
  end if;
  return new;
end $$;
