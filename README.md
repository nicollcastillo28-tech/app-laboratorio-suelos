# Geodelta Lab

App para digitar ensayos de laboratorio de suelos, con estructura Proyecto → Perforación → Muestra → Ensayo.

## Cómo correrla en tu computador

```
pip install -r requirements.txt
streamlit run app.py
```

## Claves de acceso (cámbialas en app.py, variable PASSWORDS)

- Jefe: geodelta2024
- Auxiliar: aux2024

## Novedades de esta versión

- **Bug de la bitácora corregido de raíz**: antes se reconstruía la tabla desde cero en cada
  actualización, lo que hacía que la primera vez que digitabas algo se perdiera. Ahora la tabla
  vive de forma estable en la sesión y no se pierde nada.
- **Equipo utilizado** ahora es una lista para marcar (checklist), no texto libre. Edita la lista
  `EQUIPO_LIST` en `app.py` para ajustarla a tu laboratorio (por ahora sin código de equipo).
- **Contraste de color**: los campos digitables ahora tienen fondo blanco con borde, distinto al
  fondo gris de la página.
- **Buscar ensayos** ahora muestra el código completo de la muestra (ej. `GDA-001-24_S1_M1`) y
  tiene botones por fila: abrir el ensayo, o descargar directo el Excel de Granulometría (Jefe).

## Cómo funciona

1. El **Jefe** crea el Proyecto y entra a la **Bitácora**: agrega Perforaciones (código automático S1, AP1…)
   y, dentro de cada una, las Muestras en una tabla (número, profundidad, tipo, y checklist de ensayos).
2. Nada queda guardado hasta que el Jefe presiona **💾 Guardar bitácora**.
3. Los **Auxiliares** ven la bitácora en modo lectura, abren cada muestra y solo ven/digitan los ensayos
   que el Jefe marcó. El estado de la muestra se calcula solo según el estado de cada ensayo, y cada
   ensayo muestra la fecha/hora de la última actualización.
4. El Jefe ve el progreso de los auxiliares sin entrar a cada muestra, desde la pantalla del proyecto.
5. El Jefe descarga el Excel de Granulometría ya lleno con la plantilla oficial, desde el ensayo o
   directamente desde Buscar ensayos.

## Pendiente

- Excel para Humedad y Peso unitario (falta que compartas esas plantillas)
- Guardado permanente y compartido entre dispositivos (próxima fase con Google Sheets)
- Código para cada equipo de laboratorio
