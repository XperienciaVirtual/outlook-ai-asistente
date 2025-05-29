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
            let lastBodyLineIndex = lineas.length - 1;

            // Patrones para identificar la firma real (nombres, contactos, URLs, etc.)
            const patronesFirmaOInicioFirma = [
                /^--/,
                /\d{3}[-.\s]?\d{3}[-.\s]?\d{4}/, // Números de teléfono
                /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i, // Direcciones de correo electrónico
                /(http|https):\/\/[^\s]+/i, // URLs
                /^(daniel|juan|maria|pedro|ana|jose|luis|carlos|javier|pablo|fernando|alberto|sergio|david|antonio|francisco|manuel|alejandro|miguel|rafael|ramon|roberto|santiago|vicente|angel|arturo|benito|cristian|diego|eduardo|felipe|gabriel|hector|ignacio|jaime|joaquin|jorge|julian|leonardo|marcos|martin|mateo|nicolas|oscar|pedro|quique|ricardo|ruben|salvador|tomas|victor|walter|xavi|yago|zaqueo)[\s\S]*$/i, // Nombres comunes
                /project manager/i, /director/i, /ceo/i, /gerente/i, /sales/i, /marketing/i // Títulos de cargo
            ];

            let foundSignatureStart = -1; // Índice donde la firma realmente comienza

            for (let i = lineas.length - 1; i >= 0; i--) {
                const linea = lineas[i].trim();

                // 1. Si la línea está vacía
                if (linea.length === 0) {
                    if (foundSignatureStart !== -1) {
                        // Si ya habíamos encontrado un componente de firma, y ahora hay una línea vacía,
                        // significa que el cuerpo termina antes de esta línea vacía.
                        lastBodyLineIndex = i;
                        break; // Hemos encontrado el final del cuerpo
                    }
                    continue; // Ignorar líneas vacías al principio de la búsqueda
                }

                // 2. Comprobar si la línea es una frase de despedida (y nada más)
                const isGreetingLike = /^(saludos|un saludo|saludos cordiales|cordialmente|atte\.|sinceramente|best regards|kind regards),?$/i.test(linea);
                if (isGreetingLike) {
                    // Si es una despedida y NO hemos encontrado aún un componente de firma CLARO,
                    // entonces esta línea es parte del cuerpo. Detenemos la búsqueda de firma aquí.
                    if (foundSignatureStart === -1) {
                        lastBodyLineIndex = i; // La despedida es parte del cuerpo
                        break; // Hemos encontrado el final del cuerpo
                    }
                    // Si ya habíamos encontrado un componente de firma, esta despedida es parte de la firma.
                }

                // 3. Comprobar si la línea contiene un componente de firma claro
                let isClearSignatureComponent = false;
                for (const patron of patronesFirmaOInicioFirma) {
                    if (patron.test(linea)) {
                        isClearSignatureComponent = true;
                        break;
                    }
                }

                if (isClearSignatureComponent) {
                    foundSignatureStart = i; // Marcar esta línea como el posible inicio de la firma
                    lastBodyLineIndex = i; // Por ahora, el cuerpo termina aquí
                } else if (foundSignatureStart !== -1) {
                    // Si ya habíamos encontrado un componente de firma, y esta línea NO es un componente de firma,
                    // significa que el cuerpo termina justo antes de 'foundSignatureStart'.
                    lastBodyLineIndex = foundSignatureStart; // El cuerpo termina en el inicio de la firma
                    break; // Hemos encontrado el final del cuerpo
                } else {
                    // Si no es un componente de firma, ni una despedida, ni hemos encontrado una firma aún,
                    // esta línea es parte del cuerpo. Continuar buscando hacia arriba.
                    lastBodyLineIndex = i; // Mantener el cuerpo extendido
                }
            }

            // Si no se encontró ninguna firma clara, devolver el texto original
            if (foundSignatureStart === -1) {
                return texto;
            }

            // Si se encontró un potencial inicio de firma, cortar desde ahí. De lo contrario, devolver todo.
            if (foundSignatureStart !== -1) {
                return lineas.slice(0, foundSignatureStart).join('\n').trim();
            } else {
                return texto; // No se detectó firma, devolver el texto original completo
            }
        }

        function extraerCuerpoPrincipal(texto) {
            const lineas = texto.split(/\r?\n/);
            let lastBodyLineIndex = lineas.length - 1;

            // Patrones para identificar la firma real (nombres, contactos, URLs, etc.)
            const patronesFirmaOInicioFirma = [
                /^--/,
                /\d{3}[-.\s]?\d{3}[-.\s]?\d{4}/, // Números de teléfono
                /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i, // Direcciones de correo electrónico
                /(http|https):\/\/[^\s]+/i, // URLs
                /^(daniel|juan|maria|pedro|ana|jose|luis|carlos|javier|pablo|fernando|alberto|sergio|david|antonio|francisco|manuel|alejandro|miguel|rafael|ramon|roberto|santiago|vicente|angel|arturo|benito|cristian|diego|eduardo|felipe|gabriel|hector|ignacio|jaime|joaquin|jorge|julian|leonardo|marcos|martin|mateo|nicolas|oscar|pedro|quique|ricardo|ruben|salvador|tomas|victor|walter|xavi|yago|zaqueo)[\s\S]*$/i, // Nombres comunes
                /project manager/i, /director/i, /ceo/i, /gerente/i, /sales/i, /marketing/i // Títulos de cargo
            ];

            let foundSignatureStart = -1; // Índice donde la firma realmente comienza

            for (let i = lineas.length - 1; i >= 0; i--) {
                const linea = lineas[i].trim();

                // 1. Si la línea está vacía
                if (linea.length === 0) {
                    if (foundSignatureStart !== -1) {
                        // Si ya habíamos encontrado un componente de firma, y ahora hay una línea vacía,
                        // significa que el cuerpo termina antes de esta línea vacía.
                        lastBodyLineIndex = i;
                        break; // Hemos encontrado el final del cuerpo
                    }
                    continue; // Ignorar líneas vacías al principio de la búsqueda
                }

                // 2. Comprobar si la línea es una frase de despedida (y nada más)
                const isGreetingLike = /^(saludos|un saludo|saludos cordiales|cordialmente|atte\.|sinceramente|best regards|kind regards),?$/i.test(linea);
                if (isGreetingLike) {
                    // Si es una despedida y NO hemos encontrado aún un componente de firma CLARO,
                    // entonces esta línea es parte del cuerpo. Detenemos la búsqueda de firma aquí.
                    if (foundSignatureStart === -1) {
                        lastBodyLineIndex = i; // La despedida es parte del cuerpo
                        break; // Hemos encontrado el final del cuerpo
                    }
                    // Si ya habíamos encontrado un componente de firma, esta despedida es parte de la firma.
                }

                // 3. Comprobar si la línea contiene un componente de firma claro
                let isClearSignatureComponent = false;
                for (const patron of patronesFirmaOInicioFirma) {
                    if (patron.test(linea)) {
                        isClearSignatureComponent = true;
                        break;
                    }
                }

                if (isClearSignatureComponent) {
                    foundSignatureStart = i; // Marcar esta línea como el posible inicio de la firma
                    lastBodyLineIndex = i; // Por ahora, el cuerpo termina aquí
                } else if (foundSignatureStart !== -1) {
                    // Si ya habíamos encontrado un componente de firma, y esta línea NO es un componente de firma,
                    // significa que el cuerpo termina justo antes de 'foundSignatureStart'.
                    lastBodyLineIndex = foundSignatureStart; // El cuerpo termina en el inicio de la firma
                    break; // Hemos encontrado el final del cuerpo
                } else {
                    // Si no es un componente de firma, ni una despedida, ni hemos encontrado una firma aún,
                    // esta línea es parte del cuerpo. Continuar buscando hacia arriba.
                    lastBodyLineIndex = i; // Mantener el cuerpo extendido
                }
            }

            if (lastBodyLineIndex < 0) lastBodyLineIndex = 0;

            // Si no se encontró ninguna firma clara, devolver el texto original
            if (foundSignatureStart === -1) {
                return texto;
            }

            // Si se encontró un potencial inicio de firma, cortar desde ahí. De lo contrario, devolver todo.
            if (foundSignatureStart !== -1) {
                return lineas.slice(0, foundSignatureStart).join('\n').trim();
            } else {
                return texto; // No se detectó firma, devolver el texto original completo
            }
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
                    let textoParaEnviar = extraerCuerpoPrincipal(htmlContent); // Limpiamos el HTML y la firma antes de enviar

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
