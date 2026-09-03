(function() {
    'use strict';

    function initNivelDinamico() {
        const edificioSelect = document.getElementById('id_edificio');
        const nivelInput = document.getElementById('id_nivel');
        if (!edificioSelect || !nivelInput) return;

        function actualizarNiveles() {
            const edificioId = edificioSelect.value;
            if (!edificioId) {
                nivelInput.value = '';
                nivelInput.placeholder = 'Seleccione un edificio primero';
                return;
            }

            fetch('/api/niveles-edificio/?edificio_id=' + edificioId)
                .then(r => r.json())
                .then(data => {
                    const niveles = data.niveles || ['PB'];
                    const valorActual = nivelInput.value;

                    // Reemplazar input por select temporal
                    const select = document.createElement('select');
                    select.id = 'id_nivel';
                    select.name = 'nivel';
                    select.className = 'form-control';
                    select.size = Math.min(niveles.length, 8);

                    niveles.forEach(nivel => {
                        const opt = document.createElement('option');
                        opt.value = nivel;
                        opt.textContent = nivel;
                        if (nivel === valorActual) opt.selected = true;
                        select.appendChild(opt);
                    });

                    nivelInput.parentNode.replaceChild(select, nivelInput);
                })
                .catch(() => {});
        }

        // Si nivel ya es un select, no re-inicializar
        if (nivelInput.tagName === 'SELECT') {
            edificioSelect.addEventListener('change', function() {
                const nivelActual = nivelInput.value;
                const edificioId = edificioSelect.value;
                if (!edificioId) return;

                fetch('/api/niveles-edificio/?edificio_id=' + edificioId)
                    .then(r => r.json())
                    .then(data => {
                        const niveles = data.niveles || ['PB'];
                        nivelInput.innerHTML = '';
                        niveles.forEach(nivel => {
                            const opt = document.createElement('option');
                            opt.value = nivel;
                            opt.textContent = nivel;
                            nivelInput.appendChild(opt);
                        });
                        if (niveles.includes(nivelActual)) {
                            nivelInput.value = nivelActual;
                        }
                    });
            });
        } else {
            edificioSelect.addEventListener('change', actualizarNiveles);
            actualizarNiveles();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initNivelDinamico);
    } else {
        initNivelDinamico();
    }
})();
