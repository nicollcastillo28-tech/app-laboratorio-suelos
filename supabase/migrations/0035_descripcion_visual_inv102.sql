-- Descripción visual alineada con el INV E-102-13 (Tabla 102-13, lista de comprobación): tres descriptores más.
--   desc_estructura: ESTRATIFICADA, LAMINADA, FISURADA, LISA, EN BLOQUES, LENTICULAR, HOMOGÉNEA (Tabla 102-7)
--   desc_hcl:        NULA, DÉBIL, FUERTE (Tabla 102-4)
--   desc_olor:       ORGÁNICO, INUSUAL (numeral 9.4: solo se menciona si es orgánico o inusual)
-- Los componentes secundarios con 30 % o más ("arcilla arenosa") se marcan con un "*" dentro de la columna ya existente desc_tipo_secundario;
-- no requieren columna nueva.
alter table muestras add column if not exists desc_estructura text;
alter table muestras add column if not exists desc_hcl text;
alter table muestras add column if not exists desc_olor text;
