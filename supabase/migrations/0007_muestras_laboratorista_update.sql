-- Geodelta Lab — la política anterior de muestras solo dejaba escribir al jefe,
-- pero "Descripción visual", "Observaciones" y "Condición del Ensayo" las digita
-- el laboratorista directamente sobre una muestra ya creada. Crear/eliminar
-- muestras (estructura) sigue siendo solo del jefe (vía Bitácora); actualizar
-- campos de una muestra existente lo puede hacer jefe o laboratorista.
drop policy if exists muestras_jefe_write on muestras;

create policy muestras_jefe_insert on muestras for insert
  to authenticated with check (
    exists (select 1 from profiles where id = auth.uid() and role = 'jefe')
  );

create policy muestras_jefe_delete on muestras for delete
  to authenticated using (
    exists (select 1 from profiles where id = auth.uid() and role = 'jefe')
  );

create policy muestras_update_jefe_o_laboratorista on muestras for update
  to authenticated
  using (exists (select 1 from profiles where id = auth.uid() and role in ('jefe','laboratorista')))
  with check (exists (select 1 from profiles where id = auth.uid() and role in ('jefe','laboratorista')));

-- Verificación: debe listar 4 políticas (select_all + las 3 de arriba).
select policyname, cmd, roles from pg_policies where tablename = 'muestras';
