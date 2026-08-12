-- Geodelta Lab — vincula las cuentas creadas en Authentication → Users
-- con su nombre y rol en la tabla profiles.
-- Ejecutar en el SQL Editor DESPUÉS de crear las 5 cuentas (Add user)
-- con los correos sintéticos indicados abajo.

insert into profiles (id, full_name, role)
select id, v.full_name, v.role::role_enum
from auth.users
join (values
  ('ncastillo@geodelta-lab.local', 'Nicoll Castillo', 'laboratorista'),
  ('elara@geodelta-lab.local',     'Esteban Lara',    'laboratorista'),
  ('aherrera@geodelta-lab.local',  'Andrés Herrera',  'laboratorista'),
  ('mbaron@geodelta-lab.local',    'Miguel Barón',    'jefe'),
  ('sduran@geodelta-lab.local',    'Sergio Durán',    'ingeniero')
) as v(email, full_name, role)
  on v.email = auth.users.email
on conflict (id) do update set full_name = excluded.full_name, role = excluded.role;

-- Verificación: debe devolver las 5 filas con su rol correcto.
select p.full_name, p.role, u.email
from profiles p join auth.users u on u.id = p.id
order by p.full_name;
