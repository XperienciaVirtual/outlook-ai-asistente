const fetch = require('node-fetch');

exports.handler = async function(event, context) {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Método no permitido' })
    };
  }

  const { prompt } = JSON.parse(event.body || '{}'); // Ahora esperamos 'prompt'
  if (!prompt) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Faltan datos requeridos: prompt' }) // Mensaje de error actualizado
    };
  }

  // El prompt ya viene construido desde el frontend

  try {
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
        return {
            statusCode: 500,
            body: JSON.stringify({ error: 'OPENAI_API_KEY no configurada en las variables de entorno de Netlify.' })
        };
    }
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4o',
        messages: [
          { role: 'system', content: 'Eres un asistente de redacción profesional.' },
          { role: 'user', content: prompt } // Usamos el prompt recibido directamente
        ],
        temperature: 0.3,
        max_tokens: 700
      })
    });
    const result = await response.json();
    if (!response.ok) { // Verificar si la llamada a la API de OpenAI falló
        console.error('Error de la API de OpenAI:', result);
        throw new Error(result.error?.message || 'Error desconocido de la API de OpenAI');
    }
    if (!result.choices || !result.choices[0]) throw new Error('Sin respuesta válida de OpenAI');

    const output = result.choices[0].message.content;
    // Ya no necesitamos separar explicaciones, el modelo solo debe devolver el correo mejorado
    const correoMejorado = output.trim();

    return {
      statusCode: 200,
      body: JSON.stringify({ correoMejorado: correoMejorado })
    };
  } catch (err) {
    console.error('Error en la función mejorar-correo:', err); // Registrar el error en Netlify
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message || 'Error inesperado en la función.' })
    };
  }
};