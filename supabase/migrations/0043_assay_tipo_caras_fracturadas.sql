-- Geodelta Lab: habilita 'caras-fracturadas' como tipo de ensayo soportado (INV E-227-13, bitácora GDA-FL-022, informe GDA-FLC-032).
alter table assays drop constraint if exists assays_tipo_check;
alter table assays add constraint assays_tipo_check
  check (tipo in ('granulometria','humedad','masa-unitaria','limites','pasa200','cbr','corte-directo','gravedad-especifica','proctor','materia-organica','limite-contraccion','consolidacion','compresion-inconfinada','compresion-roca','carga-puntual','solidez-sulfatos','terrones-arcilla','alargamiento-aplanamiento','caras-fracturadas'));
