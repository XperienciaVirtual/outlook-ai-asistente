// taskpane.js: Lógica del complemento Outlook para mejorar correos usando API serverless

Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // El DOM ya está listo aquí, no necesitamos DOMContentLoaded
        const form = document.getElementById('correoForm');
        const resultado = document.getElementById('resultado');
        const correoMejorado = document.getElementById('correoMejorado');
        const explicaciones = document.getElementById('explicaciones');
        const cargando = document.getElementById('cargando');
        const errorDiv = document.getElementById('error');
        const volverBtn = document.getElementById('volver');
        const traducirBtn = document.getElementById('traducirBtn');
        const traduccionResultado = document.getElementById('traduccionResultado');
        const correoTraducido = document.getElementById('correoTraducido');
        const volverTraduccionBtn = document.getElementById('volverTraduccion');

        // Función para limpiar el estado y mostrar la sección de carga
        function mostrarCargando() {
            resultado.classList.add('hidden');
            traduccionResultado.classList.add('hidden');
            cargando.classList.remove('hidden');
            errorDiv.classList.add('hidden');
            explicaciones.innerHTML = '';
            correoMejorado.textContent = '';
            correoTraducido.textContent = '';
        }

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

                    try {
                        const response = await fetch('/.netlify/functions/mejorar-correo', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({ correo: correoContent })
                        });
                        if (!response.ok) throw new Error('Error al comunicarse con el servidor');
                        const data = await response.json();
                        correoMejorado.textContent = data.correoMejorado;
                        data.explicaciones.forEach(exp => {
                            const li = document.createElement('li');
                            li.textContent = exp;
                            explicaciones.appendChild(li);
                        });
                        cargando.classList.add('hidden');
                        resultado.classList.remove('hidden');
                    } catch (err) {
                        cargando.classList.add('hidden');
                        errorDiv.textContent = err.message || 'Error inesperado.';
                        errorDiv.classList.remove('hidden');
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

            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const correoCompleto = asyncResult.value;
                    const correoSinFirma = eliminarFirma(correoCompleto);

                    try {
                        const response = await fetch('/.netlify/functions/traducir-correo', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({ texto: correoSinFirma })
                        });
                        if (!response.ok) throw new Error('Error al comunicarse con el servidor de traducción');
                        const data = await response.json();
                        correoTraducido.textContent = data.textoTraducido;
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
            resultado.classList.add('hidden');
            // form.reset(); // No es necesario resetear el formulario ya que no hay campos de entrada
        });

        volverTraduccionBtn.addEventListener('click', function () {
            traduccionResultado.classList.add('hidden');
        });
    }
});
