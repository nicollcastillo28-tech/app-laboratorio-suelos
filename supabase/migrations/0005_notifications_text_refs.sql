-- Geodelta Lab — simplifica notifications: codigo_interno/perforacion_codigo/
-- muestra_id_unico son identificadores de negocio en texto plano (como ya
-- funcionaba en memoria), no FKs — una notificación es informativa/efímera,
-- no necesita integridad referencial con projects/muestras.
alter table notifications drop column if exists project_id;
alter table notifications drop column if exists muestra_id;
alter table notifications add column if not exists muestra_id_unico text;
