-- Descripción visual: intensidad del color (CLARO / OSCURO) como campo propio, separado de la subtonalidad
-- ("MARRÓN ROJIZO CLARO" = color + matiz + intensidad). Los datos viejos que guardaban CLARO/OSCURO en desc_subtonalidad
-- se siguen leyendo bien (ver _partes_color en app.py), no hace falta migrarlos.
-- Los componentes secundarios múltiples ("arcilla con algo de grava y arena") se guardan separados por coma en la columna
-- ya existente desc_tipo_secundario (text), así que no requieren columna nueva.
alter table muestras add column if not exists desc_intensidad text;
comment on column muestras.desc_intensidad is 'Intensidad del color: CLARO u OSCURO — ver DESC_INTENSIDAD_OPTIONS en app.py.';
