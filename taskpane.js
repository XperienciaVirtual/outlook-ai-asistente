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

        form.addEventListener('submit', async function (e) {
            e.preventDefault();
            resultado.classList.add('hidden');
            cargando.classList.remove('hidden');
            errorDiv.classList.add('hidden');
            explicaciones.innerHTML = '';
            correoMejorado.textContent = '';

            // Obtener el cuerpo del correo
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

        volverBtn.addEventListener('click', function () {
            resultado.classList.add('hidden');
        });
    }
});
