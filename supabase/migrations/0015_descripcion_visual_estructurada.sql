-- Geodelta Lab — descripción visual de la muestra por menús desplegables (tipo de suelo, color,
-- cementación, consistencia/compacidad, condición de humedad) en vez de un solo campo de texto
-- libre. Sigue la nomenclatura estándar de descripción visual-manual de suelos (INV E-102 /
-- ASTM D2488), para poder comparar de un vistazo contra la clasificación USCS calculada de los
-- datos de Granulometría/Límites. La columna "descripcion_visual" ya existente se conserva tal
-- cual, pero ahora funciona como "Notas adicionales" (texto libre corto) en vez del campo
-- principal — no hace falta migrar datos viejos, siguen viéndose igual.
alter table muestras
  add column if not exists desc_tipo_suelo text,
  add column if not exists desc_color text,
  add column if not exists desc_cementacion text,
  add column if not exists desc_consistencia text,
  add column if not exists desc_humedad text;

comment on column muestras.desc_tipo_suelo is 'Grava, Arena, Limo, Arcilla u Orgánico — ver DESC_TIPO_SUELO_OPTIONS en app.py.';
comment on column muestras.desc_color is 'Ver DESC_COLOR_OPTIONS en app.py.';
comment on column muestras.desc_cementacion is 'Ver DESC_CEMENTACION_OPTIONS en app.py.';
comment on column muestras.desc_consistencia is
  'Compacidad (suelos granulares) o consistencia (suelos cohesivos) — la escala depende de '
  'desc_tipo_suelo, ver DESC_CONSISTENCIA_GRANULAR_OPTIONS / DESC_CONSISTENCIA_COHESIVO_OPTIONS en app.py.';
comment on column muestras.desc_humedad is 'Húmeda o Seca — ver DESC_HUMEDAD_OPTIONS en app.py.';
