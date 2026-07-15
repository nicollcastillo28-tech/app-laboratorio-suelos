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

## Cómo funciona la estructura

1. El **Jefe** crea el **Proyecto** (código interno GDA-NNN-AA, nombre, localización, norma, fechas).
2. El Jefe entra a la **Bitácora** del proyecto y agrega **Perforaciones** (Sondeo, Apique o Fuente/Cantera) —
   cada una recibe un código automático (S1, S2, AP1…).
3. Dentro de cada perforación, el Jefe agrega **Muestras** en una tabla (número que él mismo asigna,
   profundidad de/hasta, tipo de muestra, y qué ensayos aplican mediante checklist). La app arma el
   ID único de cada muestra como `GDA-NNN-AA_S1_M17`, por ejemplo.
4. Los **Auxiliares** abren el proyecto → la perforación → la muestra, y solo ven los ensayos que el
   Jefe marcó en la bitácora. Ahí digitan los datos y cambian el estado de la muestra.
5. El Jefe descarga el Excel de Granulometría ya lleno con la plantilla oficial (fórmulas incluidas).

## Estado actual

- ✅ Login con roles Jefe / Auxiliar
- ✅ Estructura Proyecto → Perforación → Muestra con códigos únicos automáticos
- ✅ Bitácora real que genera la estructura (no solo una tabla de referencia)
- ✅ Formularios de ensayo sin cálculos (los cálculos viven en el Excel)
- ✅ Descarga de Excel real de Granulometría usando la plantilla oficial
- ⏳ Excel para Humedad y Peso unitario (falta que compartas esas plantillas)
- ⏳ Guardado permanente y compartido entre todos los usuarios (próxima fase con Google Sheets)
