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

## Cómo funciona

1. El **Jefe** crea el Proyecto y entra a la **Bitácora**: agrega Perforaciones (código automático S1, AP1…)
   y, dentro de cada una, las Muestras en una tabla (número, profundidad, tipo, y checklist de ensayos).
2. Nada queda guardado hasta que el Jefe presiona **💾 Guardar bitácora** — antes de eso es solo un borrador.
3. Los **Auxiliares** ven la bitácora en modo lectura, abren cada muestra y solo ven/digitan los ensayos
   que el Jefe marcó. El estado de la muestra se calcula solo según el estado de cada ensayo (ya no se
   marca a mano), y cada ensayo muestra la fecha/hora de la última actualización.
4. El Jefe puede ver el progreso de los auxiliares sin entrar a cada muestra, desde la pantalla del
   proyecto (contador de sin iniciar / en proceso / finalizado por perforación y en general).
5. El Jefe puede borrar proyectos y perforaciones completas (con confirmación); para borrar una muestra
   suelta, se selecciona su fila en la tabla de la bitácora y se usa el ícono de basura.
6. El Jefe descarga el Excel de Granulometría ya lleno con la plantilla oficial (fórmulas incluidas).

## Estado actual

- ✅ Roles Jefe / Auxiliar, bitácora visible para ambos (edición solo Jefe)
- ✅ Guardado explícito de la bitácora (ya no se pierde nada al salir sin guardar)
- ✅ Borrado de proyectos y perforaciones con confirmación
- ✅ Estado por ensayo con fecha/hora de última actualización
- ✅ Progreso visible para el Jefe sin entrar a cada muestra
- ✅ Descarga de Excel real de Granulometría con la plantilla oficial
- ⏳ Excel para Humedad y Peso unitario (falta que compartas esas plantillas)
- ⏳ Guardado permanente y compartido entre dispositivos (próxima fase con Google Sheets)
