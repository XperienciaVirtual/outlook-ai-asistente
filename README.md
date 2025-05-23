# Asistente de Redacción para Outlook

Este complemento de Outlook te ayuda a mejorar la redacción de tus correos electrónicos, haciéndolos más formales, claros y efectivos, respetando siempre tu estilo y propósito original.

## Características
- Corrige errores ortográficos y gramaticales.
- Sugiere mejoras en estructura y expresiones.
- Adapta el tono según tus indicaciones.
- Interfaz en español.
- Integración segura con la API de OpenAI (ChatGPT) usando funciones serverless de Netlify.

## Estructura del proyecto
- `taskpane.html`: Interfaz principal del usuario.
- `taskpane.js`: Lógica de interacción y llamada a la función serverless.
- `styles.css`: Estilos visuales.
- `manifest.xml`: Manifest para el complemento de Outlook.
- `netlify/functions/mejorar-correo.js`: Función serverless para conectar con OpenAI (tu API Key va en las variables de entorno de Netlify).
- `assets/`: Carpeta para recursos estáticos.

## Instalación y despliegue
1. **Clona este repositorio en tu máquina y súbelo a GitHub.**
2. **En Netlify:**
   - Sube el proyecto y activa funciones serverless.
   - Añade tu API Key de OpenAI como variable de entorno: `OPENAI_API_KEY`.
3. **En Outlook:**
   - Usa el `manifest.xml` para cargar el complemento en tu cuenta de Outlook.

## Uso
1. Abre el complemento en Outlook.
2. Escribe o pega tu correo, indica el destinatario y el propósito.
3. Haz clic en "Mejorar correo" y espera la respuesta.
4. Copia el correo mejorado y revisa las explicaciones de las mejoras.

---

¿Dudas? ¡Contáctame para soporte!
