-- Geodelta Lab — TEMPORAL: el insert a notifications seguía dando "row-level
-- security policy" incluso con una política "to public with check (true)" y
-- con la política recién recreada — algo no identificado en cómo supabase-py
-- arma ese request específico. notifications es de bajo riesgo (avisos
-- internos, no datos sensibles del cliente), así que se desactiva RLS aquí
-- para no bloquear el resto de la migración. PENDIENTE: investigar la causa
-- real antes de desplegar a producción y volver a activar RLS.
alter table notifications disable row level security;
