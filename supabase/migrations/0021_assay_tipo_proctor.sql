-- Geodelta Lab: habilita 'proctor' como tipo de ensayo soportado.
alter table assays drop constraint if exists assays_tipo_check;
alter table assays add constraint assays_tipo_check
  check (tipo in ('granulometria','humedad','masa-unitaria','limites','pasa200','cbr','corte-directo','gravedad-especifica','proctor'));
