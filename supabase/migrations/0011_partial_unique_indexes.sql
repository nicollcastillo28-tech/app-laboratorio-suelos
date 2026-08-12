-- Geodelta Lab — los índices únicos originales no excluían filas archivadas,
-- así que archivar una perforación/muestra/proyecto y luego crear uno nuevo
-- con el mismo código (algo normal: el código se calcula solo, ej. "S1") choca
-- con la fila archivada. Se recrean como índices parciales (solo exigen
-- unicidad entre filas activas).
drop index if exists projects_numero_anio_uk;
create unique index projects_numero_anio_uk on projects (numero, anio) where not archived;

alter table projects drop constraint if exists projects_codigo_interno_key;
create unique index projects_codigo_interno_uk on projects (codigo_interno) where not archived;

drop index if exists perforaciones_project_codigo_uk;
create unique index perforaciones_project_codigo_uk on perforaciones (project_id, codigo) where not archived;

drop index if exists muestras_perforacion_numero_uk;
create unique index muestras_perforacion_numero_uk on muestras (perforacion_id, numero) where not archived;

alter table muestras drop constraint if exists muestras_id_unico_key;
create unique index muestras_id_unico_uk on muestras (id_unico) where not archived;

drop index if exists assays_muestra_tipo_uk;
create unique index assays_muestra_tipo_uk on assays (muestra_id, tipo) where not archived;
