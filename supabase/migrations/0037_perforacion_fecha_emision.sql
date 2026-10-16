-- Fecha de emisión POR PERFORACIÓN: el día en que se descargó el Excel de resultados de esa perforación (cada
-- sondeo/apique puede entregarse un día distinto). La app la actualiza sola en cada descarga.
alter table perforaciones add column if not exists fecha_emision date;
