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

        const prompt = `Traduce el siguiente texto al inglés. Solo proporciona la traducción, sin explicaciones adicionales:

${texto}`;

        const completion = await openai.chat.completions.create({
            model: 'gpt-4o', // Usar el modelo gpt-4o para la traducción
            messages: [
                { role: 'system', content: 'Eres un traductor profesional. Solo proporcionas la traducción solicitada.' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.7,
            max_tokens: 1000,
        });

        const textoTraducido = completion.choices[0].message.content.trim();

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
