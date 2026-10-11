# Revisor web de informes MIN328 — 2026-2

Aplicación web basada en Streamlit. Puede ejecutarse en una computadora mediante el navegador o desplegarse en un servicio de hosting compatible, por ejemplo Streamlit Community Cloud.

## Características
- Carga múltiple de informes Word (.docx) y PDF.
- Evalúa con IA usando los descriptores de la rúbrica oficial proporcionada.
- Máximo de 10.00 puntos; carátula sin puntaje.
- Los puntajes se eligen de los niveles permitidos por cada sección (incluidos incrementos de 0.25 donde la rúbrica los establece).
- Incluye evidencia, mejoras sugeridas y alertas de revisión.
- Permite confirmar/modificar puntajes por sección antes de descargar.
- Exporta Excel con hojas “Consolidado” y “Detalle por sección”.

## Ejecutar en Windows
1. Instala Python 3.11 o superior.
2. Descomprime el ZIP.
3. Abre PowerShell dentro de la carpeta del proyecto.
4. Ejecuta:
   ```powershell
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   streamlit run app.py
   ```
5. Se abrirá una dirección local en el navegador. Ingresa una clave API de OpenAI válida en la barra lateral y carga los informes.

## Publicarlo para acceder por internet
1. Sube los archivos `app.py` y `requirements.txt` a un repositorio GitHub.
2. Entra en https://share.streamlit.io/ e inicia sesión.
3. Crea una aplicación nueva seleccionando el repositorio y `app.py`.
4. En la configuración de secretos de la aplicación, añade:
   ```toml
   OPENAI_API_KEY = "tu_clave_api"
   ```
5. Publica la aplicación. Guarda la URL privada o compártela solo con las personas autorizadas.

El despliegue requiere una cuenta del servicio de hosting y una clave API con facturación/acceso habilitados. Una suscripción de ChatGPT no incluye necesariamente el uso de API.

## Límites y privacidad
- PDF escaneados sin texto seleccionable necesitan OCR.
- La extracción de texto puede omitir texto dentro de imágenes, paginación y aspectos visuales; deben revisarse manualmente.
- La evaluación de IA es una recomendación, no una nota definitiva.
- Los documentos se envían al proveedor de IA para procesarlos. Respeta las reglas de privacidad y autorización institucional.
