const fetch = require('node-fetch');

exports.handler = async function(event, context) {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Método no permitido' })
    };
  }

  const { correo } = JSON.parse(event.body || '{}'); // Solo esperamos 'correo'
  if (!correo) { // Solo verificamos que 'correo' exista
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Faltan datos requeridos: correo' })
    };
  }

  // Construir el prompt adaptado para solo usar el contenido del correo
  const prompt = `Eres un asistente experto en redacción de correos electrónicos. Tu tarea es mejorar el siguiente correo, corrigiendo ortografía y gramática, sugiriendo mejoras de estructura y expresiones para que sea más formal, claro y efectivo, pero respetando al 100% la primera frase y el tono original. No añadas adjetivos ni cambies el encabezado. Habla en singular si el texto está en primera persona. Explica cada mejora realizada de forma clara y profesional, en español.\n\nCorreo original:\n${correo}`;

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
          { role: 'user', content: prompt }
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

    // Esperamos que el modelo devuelva el correo mejorado y explicaciones separadas por un separador especial
    // Ejemplo de formato esperado: "<correo mejorado>\n---\nExplicaciones:\n- ...\n- ..."
    const output = result.choices[0].message.content;
    const parts = output.split(/---+|Explicaciones:/i);
    const correoMejorado = parts[0] ? parts[0].trim() : '';
    const explicacionesRaw = parts[1] || '';

    const explicaciones = explicacionesRaw
      ? explicacionesRaw.split(/\n|\r/).map(e => e.replace(/^[-•\s]+/, '')).filter(Boolean)
      : ['No se encontraron explicaciones.'];

    return {
      statusCode: 200,
      body: JSON.stringify({ correoMejorado: correoMejorado, explicaciones })
    };
  } catch (err) {
    console.error('Error en la función mejorar-correo:', err); // Registrar el error en Netlify
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message || 'Error inesperado en la función.' })
    };
  }
};