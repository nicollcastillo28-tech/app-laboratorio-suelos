-- Geodelta Lab — habilita 'corte-directo' como tipo de ensayo soportado (antes las 3 casillas de
-- la bitácora Corte CD/Corte CU/Corte UU no tenían formulario propio, ver
-- BITACORA_ENSAYOS/SUPPORTED_ASSAY_MAP en app.py — las 3 comparten este mismo tipo, con un
-- selector CD/CU/UU adentro del formulario). El check constraint de assays.tipo (ampliado la
-- última vez en la migración 0018 para 'cbr') no incluía 'corte-directo' — hay que ampliarlo antes
-- de poder crear una fila de este tipo.
alter table assays drop constraint if exists assays_tipo_check;
alter table assays add constraint assays_tipo_check
  check (tipo in ('granulometria','humedad','masa-unitaria','limites','pasa200','cbr','corte-directo'));
