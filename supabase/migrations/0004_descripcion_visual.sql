-- Geodelta Lab — campo que faltaba en el esquema inicial: la "Descripción
-- visual de la muestra" que digita el laboratorista (independiente de
-- "observaciones"), usada también en el Excel de Granulometría.
alter table muestras add column descripcion_visual text;
