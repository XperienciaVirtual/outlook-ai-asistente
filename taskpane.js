// taskpane.js: Lógica del complemento Outlook para mejorar correos usando API serverless

Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // El DOM ya está listo aquí, no necesitamos DOMContentLoaded
        const form = document.getElementById('correoForm');
        const resultado = document.getElementById('resultado');
        const correoMejorado = document.getElementById('correoMejorado');
        const cargando = document.getElementById('cargando');
        const errorDiv = document.getElementById('error');
        const volverBtn = document.getElementById('volver');
        const traducirBtn = document.getElementById('traducirBtn');
        const traduccionResultado = document.getElementById('traduccionResultado');
        const correoTraducido = document.getElementById('correoTraducido');
        const volverTraduccionBtn = document.getElementById('volverTraduccion');
        const instruccionesAdicionales = document.getElementById('instruccionesAdicionales');

        // Función para limpiar el estado y mostrar la sección de carga
        function mostrarCargando() {
            resultado.classList.add('hidden');
            traduccionResultado.classList.add('hidden');
            cargando.classList.remove('hidden');
            errorDiv.classList.add('hidden');
            correoMejorado.innerHTML = '';
            correoTraducido.innerHTML = '';
        }

        // Función para mostrar solo el formulario principal
        function mostrarFormularioPrincipal() {
            form.classList.remove('hidden');
            resultado.classList.add('hidden');
            traduccionResultado.classList.add('hidden');
            cargando.classList.add('hidden');
            errorDiv.classList.add('hidden');
            instruccionesAdicionales.value = '';
        }

        // Función de reintento con retroceso exponencial
        async function fetchWithRetry(url, options, retries = 3, delay = 1000) {
            try {
                const response = await fetch(url, options);
                if (!response.ok) {
                    // Si la respuesta no es OK, pero no es un error de red, reintentar si quedan intentos
                    if (retries > 0 && (response.status === 500 || response.status === 502 || response.status === 503 || response.status === 504 || response.status === 429)) {
                        console.warn(`Intento fallido (${response.status}). Reintentando en ${delay / 1000}s...`);
                        await new Promise(res => setTimeout(res, delay));
                        return fetchWithRetry(url, options, retries - 1, delay * 2);
                    } else {
                        // Si no es un error reintentable o no quedan reintentos
                        const errorText = await response.text();
                        throw new Error(`Error del servidor: ${response.status} - ${errorText}`);
                    }
                }
                return response;
            } catch (error) {
                if (retries > 0 && (error.message.includes('Failed to fetch') || error.message.includes('NetworkError'))) {
                    console.warn(`Error de red. Reintentando en ${delay / 1000}s...`);
                    await new Promise(res => setTimeout(res, delay));
                    return fetchWithRetry(url, options, retries - 1, delay * 2);
                } else {
                    throw error; // Re-lanzar el error si no es reintentable o no quedan intentos
                }
            }
        }

        // Inicializar la vista al cargar el complemento
        mostrarFormularioPrincipal();

        // Esta versión se usará para procesar la salida del modelo, si es necesario.
        function eliminarFirma(texto) {
            const lineas = texto.split(/\r?\n/);

            // Patrones comunes de inicio de firma o despedida, incluyendo despedidas
            const patronesDespedida = [
                /^--/,
                /saludos/i,
                /atentamente/i,
                /gracias/i,
                /un saludo/i,
                /best regards/i,
                /cordialmente/i,
                /a la espera/i,
                /esperando su respuesta/i,
                /sinceramente/i,
                /atte\./i,
                /suyo/i,
                /respetuosamente/i,
                /kind regards/i,
                /^(daniel|juan|maria|pedro|ana)[\s\S]*$/i // Nombres comunes al inicio de una firma
            ];

            for (let i = 0; i < lineas.length; i++) { // Iterar desde el principio
                const linea = lineas[i].trim();

                for (const patron of patronesDespedida) {
                    if (patron.test(linea) && linea.length > 0) { // Si la línea coincide con un patrón y no está vacía
                        return lineas.slice(0, i).join('\n').trim(); // Cortar todo a partir de esta línea
                    }
                }
            }

            return texto; // Si no se encuentra ninguna despedida, devolver el texto original
        }

        function extraerCuerpoPrincipal(texto) {
            const lineas = texto.split(/\r?\n/);
            let lastBodyLineIndex = lineas.length - 1;

            const patronesDespedidaOInicioFirma = [
                /^--/,
                /saludos/i,
                /atentamente/i,
                /gracias/i,
                /un saludo/i,
                /best regards/i,
                /cordialmente/i,
                /a la espera/i,
                /esperando su respuesta/i,
                /sinceramente/i,
                /atte\./i,
                /suyo/i,
                /respetuosamente/i,
                /kind regards/i,
                /^(daniel|juan|maria|pedro|ana)[\s\S]*$/i // Nombres comunes al inicio de una firma
            ];

            let emptyLineCount = 0; // Contador de líneas vacías consecutivas

            for (let i = lineas.length - 1; i >= 0; i--) {
                const linea = lineas[i].trim();

                if (linea.length === 0) {
                    emptyLineCount++;
                    if (emptyLineCount >= 2) { // Si hay 2 o más líneas vacías consecutivas, asumimos que es el inicio de la firma
                        lastBodyLineIndex = i - 1;
                        break;
                    }
                } else {
                    emptyLineCount = 0; // Resetear el contador si encontramos una línea no vacía
                }

                let esDespedidaOInicioFirma = false;
                for (const patron of patronesDespedidaOInicioFirma) {
                    if (patron.test(linea)) {
                        esDespedidaOInicioFirma = true;
                        break;
                    }
                }

                if (esDespedidaOInicioFirma) {
                    lastBodyLineIndex = i - 1;
                } else {
                    break; // Hemos encontrado una línea que no es vacía ni una despedida/firma
                }
            }

            if (lastBodyLineIndex < 0) lastBodyLineIndex = 0;

            return lineas.slice(0, lastBodyLineIndex + 1).join('\n').trim();
        }

        form.addEventListener('submit', async function (e) {
            e.preventDefault();
            mostrarCargando();

            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const correoContent = asyncResult.value; // Obtenemos el contenido completo
                    const correoSoloCuerpo = extraerCuerpoPrincipal(correoContent); // Pre-procesamos aquí
                    const instruccionesAdicionalesValue = instruccionesAdicionales.value;

                    // Reforzar el prompt para que el modelo elimine la firma y despedida
                    let prompt = `Mejora la redacción y ortografía de este correo electrónico. Mantén el tono profesional y el significado original. 

INSTRUCCIONES CRÍTICAS Y OBLIGATORIAS:
1. Respeta EXACTAMENTE la estructura y el contenido del saludo inicial (ej. 'Estimado Juan,', 'Hola equipo,') si lo hubiera. NO LO ALTERES.
2. La salida NO DEBE INCLUIR NINGÚN nombre de remitente, firma, o despedida final. OMITE COMPLETAMENTE ESTAS SECCIONES FINALES.`;

                    if (instruccionesAdicionalesValue) {
                        prompt += `\nInstrucciones adicionales: ${instruccionesAdicionalesValue}`; 
                    }

                    prompt += `\n\nCorreo original:\n${correoSoloCuerpo}`; // Enviamos solo el cuerpo

                    try {
                        // Usar fetchWithRetry para la llamada a la función serverless
                        const response = await fetchWithRetry('/.netlify/functions/mejorar-correo', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({ prompt: prompt })
                        });
                        if (!response.ok) throw new Error('Error al comunicarse con el servidor');
                        const data = await response.json();
                        
                        // Aplicar eliminarFirma a la respuesta del modelo como post-procesamiento
                        const correoMejoradoFinal = eliminarFirma(data.correoMejorado);
                        correoMejorado.innerHTML = correoMejoradoFinal.replace(/\r?\n/g, '<br>');
                        
                        cargando.classList.add('hidden');
                        resultado.classList.remove('hidden');
                        instruccionesAdicionales.value = '';
                    } catch (err) {
                        cargando.classList.add('hidden');
                        errorDiv.textContent = err.message || 'Error inesperado.';
                        errorDiv.classList.remove('hidden');
                        instruccionesAdicionales.value = '';
                    }
                } else {
                    cargando.classList.add('hidden');
                    errorDiv.textContent = 'Error al obtener el cuerpo del correo: ' + asyncResult.error.message;
                    errorDiv.classList.remove('hidden');
                }
            });
        });

        traducirBtn.addEventListener('click', async function (e) {
            e.preventDefault();
            mostrarCargando();

            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const htmlContent = asyncResult.value;
                    // NO aplicar eliminarFirma al HTML para la traducción. La función de traducción maneja el HTML directamente.
                    const textoParaEnviar = htmlContent; // Enviamos el HTML completo

                    try {
                        cargando.classList.remove('hidden');
                        // Usar fetchWithRetry para la llamada a la función serverless de traducción
                        const response = await fetchWithRetry('/.netlify/functions/traducir-correo', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                            },
                            body: JSON.stringify({ texto: textoParaEnviar }) // Enviamos el HTML
                        });
                        if (!response.ok) throw new Error('Error al comunicarse con el servidor de traducción');
                        const data = await response.json();
                        correoTraducido.innerHTML = data.textoTraducido.replace(/\r?\n/g, '<br>');
                        cargando.classList.add('hidden');
                        traduccionResultado.classList.remove('hidden');
                    } catch (err) {
                        cargando.classList.add('hidden');
                        errorDiv.textContent = err.message || 'Error inesperado en la traducción.';
                        errorDiv.classList.remove('hidden');
                    }
                } else {
                    cargando.classList.add('hidden');
                    errorDiv.textContent = 'Error al obtener el cuerpo del correo para traducir: ' + asyncResult.error.message;
                    errorDiv.classList.remove('hidden');
                }
            });
        });

        volverBtn.addEventListener('click', function () {
            // Obtener el contenido del correo mejorado
            const improvedEmailContent = correoMejorado.innerHTML;

            // Insertar el contenido en el cuerpo del correo de Outlook
            Office.context.mailbox.item.body.setAsync(improvedEmailContent, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Correo mejorado pegado en Outlook.');
                    // Opcional: Volver a la pantalla principal después de pegar
                    mostrarFormularioPrincipal();
                } else {
                    console.error('Error al pegar el correo mejorado en Outlook: ' + asyncResult.error.message);
                    // Mostrar un mensaje de error al usuario si es necesario
                    errorDiv.textContent = 'Error al pegar el correo en Outlook: ' + asyncResult.error.message;
                    errorDiv.classList.remove('hidden');
                }
            });
        });

        volverTraduccionBtn.addEventListener('click', function () {
            // Obtener el contenido del correo traducido
            const translatedEmailContent = correoTraducido.innerHTML;

            // Insertar el contenido en el cuerpo del correo de Outlook
            Office.context.mailbox.item.body.setAsync(translatedEmailContent, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Correo traducido pegado en Outlook.');
                    // Opcional: Volver a la pantalla principal después de pegar
                    mostrarFormularioPrincipal();
                } else {
                    console.error('Error al pegar el correo traducido en Outlook: ' + asyncResult.error.message);
                    // Mostrar un mensaje de error al usuario si es necesario
                    errorDiv.textContent = 'Error al pegar el correo traducido en Outlook: ' + asyncResult.error.message;
                    errorDiv.classList.remove('hidden');
                }
            });
        });
    }
});
