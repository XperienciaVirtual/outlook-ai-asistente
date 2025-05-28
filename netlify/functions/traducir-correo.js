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
        console.log('Contenido bruto de event.body:', event.body); // Añadido para depuración
        let body;
        try {
            body = typeof event.body === 'string' ? JSON.parse(event.body) : event.body;
        } catch (parseError) {
            console.error('Error al parsear el cuerpo del evento:', parseError);
            return {
                statusCode: 400,
                body: JSON.stringify({ error: 'Formato de solicitud inválido.' }),
            };
        }
        const { texto } = body;
        console.log('Valor de texto después del parseo:', texto); // Añadido para depuración
        console.log('Texto recibido en la función de traducción:', texto);

        if (!texto) {
            return {
                statusCode: 400,
                body: JSON.stringify({ error: 'El campo \'texto\' es requerido.' }),
            };
        }

        const prompt = `Detecta el idioma de este contenido HTML y tradúcelo al idioma opuesto (si es español, a inglés; si es inglés, a español). Es ABSOLUTAMENTE CRÍTICO que extraigas el texto del HTML y mantengas TODO el formato original, incluyendo todos los saltos de línea (simples y dobles), espacios en blanco y la estructura de los párrafos. Asegúrate de incluir TODAS las líneas del texto original en la traducción, incluso si son solo saltos de línea al final. No añadas ni quites nada que no sea la traducción directa. Solo proporciona la traducción en texto plano, sin HTML.`;

        const completion = await openai.chat.completions.create({
            model: 'gpt-4o',
            messages: [
                { role: 'system', content: 'Eres un traductor profesional que procesa contenido HTML, extrae el texto y lo traduce, respetando el formato original del texto, incluyendo saltos de línea y espacios en blanco. La salida debe ser texto plano.' },
                { role: 'user', content: prompt + '\n\nContenido HTML original:\n' + texto }
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
