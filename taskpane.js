// taskpane.js: Lógica del complemento Outlook para mejorar correos usando API serverless

document.addEventListener('DOMContentLoaded', function () {
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

        const destinatario = document.getElementById('destinatario').value;
        const proposito = document.getElementById('proposito').value;
        const correo = document.getElementById('correo').value;

        try {
            const response = await fetch('/.netlify/functions/mejorar-correo', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ destinatario, proposito, correo })
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
    });

    volverBtn.addEventListener('click', function () {
        resultado.classList.add('hidden');
        form.reset();
    });
});
