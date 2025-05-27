const { OpenAI } = require('openai');

exports.handler = async (event) => {
    if (event.httpMethod !== 'POST') {
        return {
            statusCode: 405,
            body: 'Method Not Allowed',
        };
    }

    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
        console.error('OPENAI_API_KEY no está configurada en las variables de entorno de Netlify.');
        return {
            statusCode: 500,
            body: JSON.stringify({ error: 'La clave API de OpenAI no está configurada.' }),
        };
    }

    const openai = new OpenAI({ apiKey: apiKey });

    try {
        const { texto } = JSON.parse(event.body || '{}');

        if (!texto) {
            return {
                statusCode: 400,
                body: JSON.stringify({ error: 'El campo \'texto\' es requerido.' }),
            };
        }

        const prompt = `Detecta el idioma de este texto y tradúcelo al idioma opuesto (si es español, a inglés; si es inglés, a español). Es ABSOLUTAMENTE CRÍTICO que mantengas TODO el formato original, incluyendo todos los saltos de línea (simples y dobles), espacios en blanco y la estructura del texto. No añadas ni quites nada que no sea la traducción directa. Solo proporciona la traducción.`;

        const completion = await openai.chat.completions.create({
            model: 'gpt-4o',
            messages: [
                { role: 'system', content: 'Eres un traductor profesional que respeta el formato original del texto, incluyendo saltos de línea y espacios en blanco.' },
                { role: 'user', content: prompt + '\n\nTexto original:\n' + texto }
            ],
            temperature: 0.7,
            max_tokens: 4000,
        });

        const textoTraducido = completion.choices[0].message.content;

        return {
            statusCode: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ textoTraducido: textoTraducido }),
        };
    } catch (error) {
        console.error('Error en la función traducir-correo:', error);
        let errorMessage = 'Error al procesar la solicitud de traducción.';
        if (error.response && error.response.data && error.response.data.error) {
            errorMessage = `Error de la API de OpenAI: ${error.response.data.error.message}`;
        } else if (error.message) {
            errorMessage = error.message;
        }
        return {
            statusCode: 500,
            body: JSON.stringify({ error: errorMessage }),
        };
    }
};
