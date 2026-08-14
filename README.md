# Generador de Documentos RRHH

App web (Streamlit) con 5 herramientas:

1. Generador de Contratos
2. Cálculo de Finiquitos (planilla Excel)
3. Finiquito + Declaración Jurada
4. Anexo de Continuidad
5. Anexo Obrero → Capataz

En cada pestaña subes **tu Excel de datos** y **tu(s) plantilla(s)** (Word con
`«CAMPO»` o Excel con las celdas fijas), le das a generar, y descargas un ZIP.
Nada queda guardado en el servidor: todo se procesa en memoria durante esa
sesión.

## Probarla en tu computador

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre solo en `http://localhost:8501`.

## Desplegarla gratis (recomendado: Streamlit Community Cloud)

1. Crea un repositorio en GitHub (puede ser privado) y sube estos 3 archivos:
   `app.py`, `requirements.txt`, `README.md`.
2. Entra a **share.streamlit.io** con tu cuenta de GitHub.
3. Click en "New app" → elige el repo → archivo principal `app.py` → Deploy.
4. En 1-2 minutos te da una URL pública (tipo `tuapp.streamlit.app`) que
   puedes compartir con quien necesite usarla.

Es gratis, no duerme tan agresivo como otras opciones, y cada vez que hagas
`git push` con cambios se actualiza sola.

### Alternativa: Render.com
Si prefieres no usar GitHub público: crea un "Web Service" en Render, conecta
el repo (puede ser privado), y como start command pon:
```
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```
El tier gratis de Render "duerme" el servicio tras un rato sin uso y tarda
~30 seg en despertar la primera vez que alguien entra — no es un problema
grave para uso ocasional tipo RRHH.

## Nota de seguridad
El script original traía una API key de Groq escrita directo en el código.
La eliminé por completo en esta versión (no se usa IA/LLM). Si en algún
momento vuelves a meter una clave de API en un proyecto, ponla como
"Secret"/variable de entorno en la plataforma de hosting, nunca directo en
el código — sobre todo si el repo es público.
