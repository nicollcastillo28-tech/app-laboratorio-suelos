-- Cuántas muestras traía la orden de cada perforación, con o sin ensayo marcado. Las muestras sin ensayo no se importan,
-- así que este total es lo único que permite mostrar "ensayo de 8 muestras de 15".
alter table perforaciones add column if not exists muestras_en_orden integer;
