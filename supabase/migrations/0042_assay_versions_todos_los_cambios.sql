-- Historial de cambios de los ensayos: se guardan TODOS los cambios (antes solo las últimas 40 versiones y se borraban las más viejas).
--
-- Cada vez que cambia `assays.data`, se guarda una fila con SOLO lo que cambió:
--   cambios = {"clave": {"a": valor antes, "d": valor después}, ...}   (null = el dato no estaba / se quitó)
--   autor   = quién hizo el cambio (assays.editado_por del estado nuevo)
--   digitado_por / digitado_en = de quién y de cuándo era el estado anterior (como en la migración 0041)
-- Guardar solo lo que cambió evita copiar el ensayo completo en cada guardado automático (en Consolidación son cientos de KB por las
-- lecturas de la máquina). La app reconstruye cómo estaba el ensayo en cualquier momento partiendo de los datos actuales y deshaciendo
-- los cambios hacia atrás. Las filas viejas (versión completa en `data`) se siguen leyendo igual.

alter table public.assay_versions add column if not exists cambios jsonb;
alter table public.assay_versions add column if not exists autor text;

create or replace function public.guardar_version_ensayo() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  viejo jsonb;
  nuevo jsonb;
  dif   jsonb;
begin
  select coalesce(jsonb_object_agg(k, v), '{}'::jsonb) into viejo
    from jsonb_each(coalesce(old.data, '{}'::jsonb)) as t(k, v)
   where k <> '_respaldo' and not (jsonb_typeof(v) = 'object' and v ? 'b64');
  select coalesce(jsonb_object_agg(k, v), '{}'::jsonb) into nuevo
    from jsonb_each(coalesce(new.data, '{}'::jsonb)) as t(k, v)
   where k <> '_respaldo' and not (jsonb_typeof(v) = 'object' and v ? 'b64');
  if viejo is distinct from nuevo then
    select coalesce(jsonb_object_agg(k, jsonb_build_object('a', viejo -> k, 'd', nuevo -> k)), '{}'::jsonb) into dif
      from (select jsonb_object_keys(viejo) as k union select jsonb_object_keys(nuevo)) as claves
     where (viejo -> k) is distinct from (nuevo -> k);
    insert into public.assay_versions (assay_id, data, cambios, status, digitado_por, digitado_en, autor)
    values (old.id, '{}'::jsonb, dif, old.status, coalesce(old.editado_por, old.laboratorist), old.updated_at,
            coalesce(new.editado_por, new.laboratorist));
    -- Ya no se borran las versiones viejas: queda todo el historial.
  end if;
  return new;
end $$;

drop trigger if exists assays_guardar_version on public.assays;
create trigger assays_guardar_version after update of data on public.assays
  for each row execute function public.guardar_version_ensayo();
