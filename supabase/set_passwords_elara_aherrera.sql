-- Geodelta Lab — fija la clave de las cuentas que no tenían opción de
-- "reset password" directa en el panel. crypt()/gen_salt('bf') vienen de
-- pgcrypto (ya habilitado desde la migración 0001).
update auth.users
set encrypted_password = crypt('lara2026', gen_salt('bf'))
where email = 'elara@geodelta-lab.local';

update auth.users
set encrypted_password = crypt('herrera2026', gen_salt('bf'))
where email = 'aherrera@geodelta-lab.local';

-- Verificación: debe mostrar las 2 cuentas con updated_at recién actualizado.
select email, updated_at from auth.users
where email in ('elara@geodelta-lab.local', 'aherrera@geodelta-lab.local');
