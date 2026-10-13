-- Fecha de ingreso de las muestras POR PERFORACIÓN: un mismo proyecto puede recibir las muestras de cada
-- sondeo/apique en días distintos. Opcional: si queda vacía se sigue usando projects.fecha_ingreso_muestra.
alter table perforaciones add column if not exists fecha_ingreso date;
