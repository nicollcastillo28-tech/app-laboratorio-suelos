-- Geodelta Lab — habilita 'cbr' como tipo de ensayo soportado (antes solo era una casilla de la
-- bitácora sin formulario propio, ver BITACORA_ENSAYOS/SUPPORTED_ASSAY_MAP en app.py). El check
-- constraint original de assays.tipo (migración 0001) no incluía 'cbr' — hay que ampliarlo antes
-- de poder crear una fila de este tipo.
alter table assays drop constraint if exists assays_tipo_check;
alter table assays add constraint assays_tipo_check
  check (tipo in ('granulometria','humedad','masa-unitaria','limites','pasa200','cbr'));
