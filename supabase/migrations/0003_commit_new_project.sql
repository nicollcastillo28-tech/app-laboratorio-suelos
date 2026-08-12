-- Geodelta Lab — commit atómico de un proyecto nuevo (proyecto + sus
-- perforaciones + las muestras de cada perforación, todo o nada).
-- SECURITY DEFINER: se salta RLS, así que el chequeo de rol se hace
-- explícito adentro de la función en vez de depender de las políticas.

create or replace function commit_new_project(project_fields jsonb, perforaciones jsonb)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  new_project_id uuid;
  new_perf_id    uuid;
  perf           jsonb;
  m              jsonb;
begin
  if not exists (select 1 from profiles where id = auth.uid() and role = 'jefe') then
    raise exception 'permiso denegado: solo el jefe de laboratorio puede crear proyectos';
  end if;

  insert into projects (
    codigo_interno, numero, anio, nombre, localizacion, norma,
    fecha_bitacora, fecha_ingreso_muestra, laboratorista_asignado,
    cliente, correo_cliente, muestra_tomada_por, direccion_cliente,
    telefono_contacto, nombre_contacto,
    fecha_inicio_proyecto, fecha_final_proyecto, fecha_recepcion,
    fecha_ejecucion, fecha_emision, created_by
  )
  values (
    project_fields->>'codigo_interno', project_fields->>'numero', project_fields->>'anio',
    project_fields->>'nombre', project_fields->>'localizacion', project_fields->>'norma',
    nullif(project_fields->>'fecha_bitacora','')::date, nullif(project_fields->>'fecha_ingreso_muestra','')::date,
    project_fields->>'laboratorista_asignado',
    project_fields->>'cliente', project_fields->>'correo_cliente', project_fields->>'muestra_tomada_por',
    project_fields->>'direccion_cliente', project_fields->>'telefono_contacto', project_fields->>'nombre_contacto',
    nullif(project_fields->>'fecha_inicio_proyecto','')::date, nullif(project_fields->>'fecha_final_proyecto','')::date,
    nullif(project_fields->>'fecha_recepcion','')::date,
    nullif(project_fields->>'fecha_ejecucion','')::date, nullif(project_fields->>'fecha_emision','')::date,
    auth.uid()
  )
  returning id into new_project_id;

  for perf in select * from jsonb_array_elements(coalesce(perforaciones, '[]'::jsonb))
  loop
    insert into perforaciones (project_id, tipo, consecutivo, codigo)
    values (new_project_id, perf->>'tipo', (perf->>'consecutivo')::int, perf->>'codigo')
    returning id into new_perf_id;

    for m in select * from jsonb_array_elements(coalesce(perf->'muestras', '[]'::jsonb))
    loop
      insert into muestras (
        perforacion_id, numero, id_unico, profundidad_de, profundidad_hasta,
        tipo_muestra, ensayos, observaciones
      )
      values (
        new_perf_id, m->>'numero', m->>'id_unico',
        nullif(m->>'profundidad_de','')::numeric, nullif(m->>'profundidad_hasta','')::numeric,
        m->>'tipo_muestra', coalesce(m->'ensayos', '{}'::jsonb), m->>'observaciones'
      );
    end loop;
  end loop;

  return new_project_id;
end;
$$;

grant execute on function commit_new_project(jsonb, jsonb) to authenticated;
