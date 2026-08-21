-- Geodelta Lab — asignación de laboratorista por ensayo individual (no por proyecto entero).
-- Vive en muestras (no en assays) porque el Jefe debe poder asignar ANTES de que exista una
-- fila de assay — esa fila solo se crea cuando alguien abre el formulario por primera vez.
alter table muestras add column if not exists ensayo_asignado_a jsonb not null default '{}'::jsonb;

comment on column muestras.ensayo_asignado_a is
  'Mapa {"Humedad": "Nicoll Castillo", ...} — a qué laboratorista está asignado cada ensayo '
  'solicitado de esta muestra (por full_name, igual que el resto de la app). Un ensayo sin '
  'entrada acá no aparece en la lista de "Ensayos asignados" de nadie hasta que el Jefe lo asigne.';
