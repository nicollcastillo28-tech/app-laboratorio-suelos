-- Plasticidad en la descripción visual de suelos finos (limo y arcilla), INV E-102-13 Tabla 102-2: NO PLÁSTICA, BAJA, MEDIA, ALTA.
alter table muestras add column if not exists desc_plasticidad text;
