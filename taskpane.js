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

        // Función heurística para intentar eliminar la firma del correo
        function eliminarFirma(texto) {
            const lineas = texto.split(/\r?\n/);
            let cuerpoLimpio = [];
            let enFirma = false;

            for (let i = lineas.length - 1; i >= 0; i--) {
                const linea = lineas[i].trim();
                if (linea.length === 0) continue; // Ignorar líneas vacías al final

                // Patrones comunes de inicio de firma o separadores
                if (linea.startsWith('--') ||
                    linea.toLowerCase().includes('saludos') ||
                    linea.toLowerCase().includes('atentamente') ||
                    linea.toLowerCase().includes('gracias') ||
                    linea.toLowerCase().includes('un saludo') ||
                    linea.toLowerCase().includes('best regards')) {
                    enFirma = true;
                }

                if (enFirma) {
                    // Si estamos en la firma, no añadir la línea al cuerpo limpio
                    // Pero si encontramos una línea que parece ser parte del cuerpo principal, salimos
                    if (linea.length > 50 && !linea.includes('http') && !linea.includes('@')) {
                        // Esto es una heurística: si la línea es larga y no parece un enlace/email, podría ser cuerpo
                        enFirma = false; // Salir del modo firma
                        cuerpoLimpio.unshift(lineas[i]); // Añadir esta línea al cuerpo
                    }
                } else {
                    cuerpoLimpio.unshift(lineas[i]);
                }
            }
            return cuerpoLimpio.join('\n').trim();
        }

        form.addEventListener('submit', async function (e) {
            e.preventDefault();
            mostrarCargando();

            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const correoContent = asyncResult.value;
                    const instrucciones = instruccionesAdicionales.value.trim();

                    try {
                        // Usar fetchWithRetry para la llamada a la función serverless
                        const response = await fetchWithRetry('/.netlify/functions/mejorar-correo', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({ correo: correoContent, instrucciones: instrucciones })
                        });
                        if (!response.ok) throw new Error('Error al comunicarse con el servidor');
                        const data = await response.json();
                        correoMejorado.innerHTML = data.correoMejorado.replace(/\r?\n/g, '<br>');
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

                    // Eliminar la firma del correo si existe (esto aún se aplica al HTML si la firma es texto simple)
                    const textoParaEnviar = eliminarFirma(htmlContent); // Ahora enviamos el HTML

                    console.log('HTML enviado a la función de traducción:', textoParaEnviar); // <-- Actualizado para depuración
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
