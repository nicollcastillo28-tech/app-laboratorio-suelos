-- Geodelta Lab — descripción visual: componente secundario (ej. "GRAVA con algo de ARENA",
-- "ARCILLA con algo de ARENA") y opción "OTROS" en el tipo de grano principal para muestras que
-- se salen de la clasificación estándar. Ver DESC_TIPO_SECUNDARIO_OPTIONS y
-- descripcion_visual_estructurada en app.py. "OTROS" no necesita columna nueva: es solo una
-- opción más de la lista ya guardada en desc_tipo_suelo (migración 0015).
alter table muestras
  add column if not exists desc_tipo_secundario text;

comment on column muestras.desc_tipo_secundario is
  'Componente secundario de la muestra (ej. "ARENA" en una grava con algo de arena) — opcional, '
  'null si no aplica. Ver DESC_TIPO_SECUNDARIO_OPTIONS en app.py.';
