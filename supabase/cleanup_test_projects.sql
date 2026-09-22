-- Geodelta Lab — borra de verdad (no solo archiva) los proyectos de prueba
-- creados durante la migración a Supabase. El "on delete cascade" de las
-- foreign keys se encarga de perforaciones/muestras/assays asociados.
delete from projects
where codigo_interno in ('GDA-TEST-99', 'GDA-010-26', 'GDA-020-26', 'GDA-030-26');

-- Verificación: no debe quedar ninguna fila con esos códigos.
select codigo_interno from projects
where codigo_interno in ('GDA-TEST-99', 'GDA-010-26', 'GDA-020-26', 'GDA-030-26');
