-- Geodelta Lab — segunda vuelta de la descripción visual estructurada (ver 0015): se quita el
-- campo de notas libres del flujo de captura (la descripción ahora es 100% menús desplegables,
-- siempre en mayúscula), se separa "color" de "subtonalidad" (antes venían mezclados como
-- "Café oscuro"/"Café claro"), y se agregan forma (solo grava) y angulosidad (suelos gruesos:
-- grava/arena) — nomenclatura INV E-102 / ASTM D2488. Los valores viejos de desc_tipo_suelo,
-- desc_color, desc_cementacion, desc_consistencia, desc_humedad (migración 0015) quedan tal cual
-- están, aunque ahora la app los relee/escribe siempre en mayúscula hacia adelante — no hace
-- falta migrar datos viejos, la comparación de opciones ya no depende del case exacto guardado.
alter table muestras
  add column if not exists desc_subtonalidad text,
  add column if not exists desc_forma text,
  add column if not exists desc_angulosidad text;

comment on column muestras.desc_subtonalidad is 'Matiz que modifica el color principal (ej. "ROJIZO" sobre "MARRÓN") — ver DESC_SUBTONALIDAD_OPTIONS en app.py.';
comment on column muestras.desc_forma is 'Forma de las partículas — solo aplica si desc_tipo_suelo = GRAVA. Ver DESC_FORMA_OPTIONS en app.py.';
comment on column muestras.desc_angulosidad is 'Angulosidad de las partículas — solo aplica a suelos gruesos (GRAVA/ARENA). Ver DESC_ANGULOSIDAD_OPTIONS en app.py.';
