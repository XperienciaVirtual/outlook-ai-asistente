const fetch = require('node-fetch');

exports.handler = async function(event, context) {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Método no permitido' })
    };
  }

  const { destinatario, proposito, correo } = JSON.parse(event.body || '{}');
  if (!correo || !proposito || !destinatario) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Faltan datos requeridos' })
    };
  }

  // Construir el prompt según las reglas del usuario
  const prompt = `Eres un asistente experto en redacción de correos electrónicos. Tu tarea es mejorar el siguiente correo, corrigiendo ortografía y gramática, sugiriendo mejoras de estructura y expresiones para que sea más formal, claro y efectivo, pero respetando al 100% la primera frase y el tono original. Dirígete al destinatario como '${destinatario}'. No añadas adjetivos ni cambies el encabezado. Habla en singular si el texto está en primera persona. Explica cada mejora realizada de forma clara y profesional, en español.\n\nPropósito: ${proposito}\n\nCorreo original:\n${correo}`;

  try {
    const apiKey = process.env.OPENAI_API_KEY;
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4',
        messages: [
          { role: 'system', content: 'Eres un asistente de redacción profesional.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.3,
        max_tokens: 700
      })
    });
    const result = await response.json();
    if (!result.choices || !result.choices[0]) throw new Error('Sin respuesta de OpenAI');

    // Esperamos que el modelo devuelva el correo mejorado y explicaciones separadas por un separador especial
    // Ejemplo de formato esperado: "<correo mejorado>\n---\nExplicaciones:\n- ...\n- ..."
    const output = result.choices[0].message.content;
    const [correoMejorado, explicacionesRaw] = output.split(/---+|Explicaciones:/i);
    const explicaciones = explicacionesRaw
      ? explicacionesRaw.split(/\n|\r/).map(e => e.replace(/^[-•\s]+/, '')).filter(Boolean)
      : ['No se encontraron explicaciones.'];

    return {
      statusCode: 200,
      body: JSON.stringify({ correoMejorado: correoMejorado.trim(), explicaciones })
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
