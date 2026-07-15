# Geodelta Lab

App para digitar ensayos de laboratorio de suelos (Granulometría, Humedad, Masa unitaria parafinada), construida en Python + Streamlit, basada en el diseño de Figma.

## Cómo correrla en tu computador

```
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá sola en el navegador en `http://localhost:8501`.

## Estado actual

- ✅ Navegación completa entre pantallas (proyectos, muestras, encabezado, ensayos)
- ✅ Formularios de los 3 tipos de ensayo con los cálculos automáticos
- ✅ Firma digital (si `streamlit-drawable-canvas` está instalado)
- ⏳ Generación de Excel y PDF (próxima fase)
- ⏳ Guardado permanente compartido entre todos los usuarios (próxima fase, hoy los datos viven mientras la app está abierta)

## Publicar para que la usen otros

Ver instrucciones detalladas en la conversación con Claude, o en `https://docs.streamlit.io/deploy/streamlit-community-cloud`.
