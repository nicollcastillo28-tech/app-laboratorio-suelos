select polname, polcmd, polroles::regrole[], polqual, polwithcheck, polpermissive
from pg_policy
where polrelid = 'notifications'::regclass;

select relrowsecurity, relforcerowsecurity
from pg_class
where oid = 'notifications'::regclass;
