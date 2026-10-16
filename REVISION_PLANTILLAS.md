# Actualización de plantillas — 16 de octubre de 2026

Se adaptan las exportaciones a los formatos oficiales de Dropbox por código GDA-FLC/GDA-FL. Se reemplazan quince archivos Excel existentes; no se agrega ningún formato Excel ni se cambia su extensión. Los archivos fuente conservan su contenido original.

Los cálculos de la aplicación se mantienen. Al exportar se escriben los datos en el XML original del libro, preservando estilos, fórmulas, imágenes, gráficos y macros. La excepción autorizada es GDA-FLC-006, G21: recibe la masa del picnómetro con agua digitada en la aplicación, incluido cero; vacía permanece vacía. Las diferencias entre los cálculos de los originales nuevos y la aplicación quedan pendientes de revisión: no se intenta hacerlos coincidir sobrescribiendo fórmulas.

Granulometría y límites siguen compartiendo el archivo completo GDA-FLC-001. Proctor conserva el informe individual con sus datos de CBR compactado. Los informes generados de Proctor y compresión en roca excluyen las hojas de ejemplos de otros proyectos que traían los originales; sus plantillas fuente no se modifican.

## Disponible

Clasificación/granulometría/límites numéricos/pasa 200; humedad; peso unitario parafinado y método B; gravedad específica de finos, gruesos y arcillas; Proctor individual; materia orgánica; límite de contracción; compresión inconfinada; compresión en roca; carga puntual; solidez en sulfatos; terrones de arcilla; alargamiento y aplanamiento; bitácora de orden y verificación de balanzas.

Compresión inconfinada GDA-FLC-008 conserva la plantilla vigente .xlsm y su exportador: la actualización entregada era .xlsx y se aplaza. Balanzas GDA-FL-029 conserva la plantilla y el exportador vigentes hasta adaptar el nuevo diseño.

## Pendiente, con aviso en la aplicación

- CBR GDA-FLC-013: I39 e I40 pertenecen a casillas combinadas H39:I39 y H40:I40; se requiere decidir la ubicación de las masas sin alterar el formato.
- Corte directo GDA-FLC-007: D26 contiene una fórmula original y el exportador pretendía escribir allí una masa digitada.
- Consolidación GDA-FLC-009: la celda X4 de parámetros de máquina contiene =U4; la aplicación maneja un valor independiente.
- Descargas agrupadas de CBR y Proctor: se aplazan; Proctor individual sigue disponible. Se han verificado seis agrupaciones: granulometría, humedad, límite de contracción, materia orgánica, masa unitaria y carga puntual.

La digitación y el guardado de ensayos permanecen disponibles. Una variante incompatible presenta un aviso de revisión y no entrega un libro parcialmente lleno.

## Validación realizada

- 18 casos de exportación individual disponibles; 950 comprobaciones de celdas contra las entradas del exportador anterior, con las posiciones nuevas correspondientes.
- Conservación de fórmulas y estructura del formato; integridad ZIP y apertura mediante openpyxl. Se verifica la conservación de cada parte retenida, con las excepciones de recálculo, G21 autorizada y exclusión de hojas de otros proyectos en las salidas.
- 7 pruebas automáticas permanentes: integridad de fuentes, escritura de humedad, cero/vacío, protección de fórmulas y celdas combinadas, rechazo de fórmulas como entradas y G21 del generador real.
- 14 pruebas locales de interfaz Streamlit sin errores: 9 ensayos disponibles y 5 casos con aviso (CBR, corte, consolidación, NP y NL). Se verificaron los botones generados.
- 3 exportaciones con fotografías y 6 archivos con dos muestras: correctos.
- Bitácora nueva: encabezado, fila de muestra, marcas de ensayos y aviso de exceso de filas. Balanzas: mismas partes y valores que la versión anterior.
- Inicio completo de la aplicación: pantalla de ingreso sin errores. Comparación del código: las funciones de cálculo permanecen iguales.

Estas comprobaciones se hicieron con datos ficticios y sin escrituras en Supabase. No equivalen a una revisión matemática del resultado recalculado ni a una inspección visual de todos los informes en Microsoft Excel.

Pruebas permanentes: `python -m unittest discover -s tests -v`.

## Corrección de presentación y clasificación — misma fecha

Las descargas centran horizontal y verticalmente los datos escritos por la aplicación. Se conservan fuentes, bordes, colores base, protección, formatos numéricos y ajuste de texto. Los estilos añadidos tienen índices estables para combinar muestras con campos distintos en un mismo libro.

Las reglas de casillas vacías se separan por áreas y por combinaciones: cada regla consulta la celda que le corresponde, o el ancla de su combinación. Así se evita que permanezcan verdes/grises las casillas ya diligenciadas. Se conservan las alertas de resultados y la precedencia de las reglas.

Por autorización del laboratorio se corrigen cinco fórmulas en la salida de clasificación 001: K10 arma el número de informe desde las celdas existentes, sin usar B50/B51 vacías; R40 y R41 evitan redondear texto o promediar datos ausentes; R42 reconoce resultados NP/NL; U41 utiliza CONCATENATE en lugar de _xludf.CONCAT. NP/NL sólo sustituyen las casillas de resultado explícitamente autorizadas. Los datos numéricos conservan el método de cálculo anterior y las macros originales.

NP y NL dejan de estar pendientes: se probaron ocho combinaciones, iniciando la descarga desde granulometría y desde límites, con límites numéricos, NP, NL y sin límites. Las descargas, las posiciones de entrada y las partes gráficas originales permanecen íntegras. Se revisó visualmente una vista de humedad y se comprobó el centrado y la lógica de colores. El visor de pruebas no admite algunas formas de los gráficos del 001, por lo que no se pudo renderizar allí el informe completo; sus objetos se verificaron comparando sus partes originales.

Hay once pruebas permanentes de integridad, casillas vacías, combinaciones, cero/vacío, alineación, estabilidad de estilos en lotes y alcance de las fórmulas autorizadas. No se ha certificado aquí el recálculo de macros en Microsoft Excel.
