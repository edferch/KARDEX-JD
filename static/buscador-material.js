// static/buscador-material.js
// Buscador de materiales con desplegable propio (código, nombre y descripción).
// Se usa en el index (modales de Entrada/Devolución/Salida) y en Reportes.

function normalizarBusquedaMaterial(texto) {
    return (texto || '').toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function escaparHtmlMaterial(texto) {
    const div = document.createElement('div');
    div.textContent = texto == null ? '' : texto;
    return div.innerHTML;
}

function crearBuscadorMaterial(inputId, hiddenId, dropdownId, opciones) {
    const input = document.getElementById(inputId);
    const hidden = document.getElementById(hiddenId);
    const dropdown = document.getElementById(dropdownId);
    if (!input || !hidden || !dropdown) return;

    const materiales = (opciones && opciones.materiales) || [];
    const onSeleccionar = opciones && opciones.onSeleccionar;
    let resultados = [];
    let indiceActivo = -1;

    // Si ya viene un material preseleccionado (ej. al recargar Reportes), no lo marcamos inválido.
    if (!hidden.value) {
        input.setCustomValidity("Debe seleccionar un material de la lista.");
    }

    function cerrarDropdown() {
        dropdown.style.display = 'none';
        dropdown.innerHTML = '';
        indiceActivo = -1;
    }

    function renderizarResultados() {
        if (resultados.length === 0) {
            dropdown.innerHTML = '<div class="dropdown-materiales-vacio">Sin resultados</div>';
            dropdown.style.display = 'block';
            return;
        }
        dropdown.innerHTML = resultados.map(function(mat, i) {
            return '<div class="dropdown-materiales-item' + (i === indiceActivo ? ' activo' : '') + '" data-index="' + i + '">' +
                '<span class="dropdown-materiales-codigo">' + (mat.codigo ? escaparHtmlMaterial(mat.codigo) : 'S/C') + '</span>' +
                '<span class="dropdown-materiales-nombre">' + escaparHtmlMaterial(mat.nombre) + '</span>' +
                (mat.descripcion ? '<span class="dropdown-materiales-desc">' + escaparHtmlMaterial(mat.descripcion) + '</span>' : '') +
                '</div>';
        }).join('');
        dropdown.style.display = 'block';

        dropdown.querySelectorAll('.dropdown-materiales-item').forEach(function(el) {
            el.addEventListener('mousedown', function(e) {
                e.preventDefault();
                seleccionarMaterial(resultados[parseInt(el.getAttribute('data-index'), 10)]);
            });
        });

        if (indiceActivo >= 0) {
            const activo = dropdown.querySelector('.dropdown-materiales-item.activo');
            if (activo) activo.scrollIntoView({ block: 'nearest' });
        }
    }

    function seleccionarMaterial(mat) {
        hidden.value = mat.id;
        input.value = (mat.codigo ? mat.codigo + ' - ' : '') + mat.nombre;
        input.setCustomValidity("");
        cerrarDropdown();
        if (typeof onSeleccionar === 'function') onSeleccionar(mat);
    }

    function buscar() {
        hidden.value = '';
        input.setCustomValidity("Debe seleccionar un material de la lista.");
        const filtro = normalizarBusquedaMaterial(input.value);

        if (filtro === '') { cerrarDropdown(); return; }

        resultados = materiales.filter(function(mat) {
            return normalizarBusquedaMaterial(mat.codigo).includes(filtro) ||
                   normalizarBusquedaMaterial(mat.nombre).includes(filtro) ||
                   normalizarBusquedaMaterial(mat.descripcion).includes(filtro);
        }).slice(0, 30);

        indiceActivo = -1;
        renderizarResultados();
    }

    input.addEventListener('input', buscar);
    input.addEventListener('focus', function() { input.select(); });

    input.addEventListener('keydown', function(e) {
        if (dropdown.style.display !== 'block') return;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            indiceActivo = Math.min(indiceActivo + 1, resultados.length - 1);
            renderizarResultados();
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            indiceActivo = Math.max(indiceActivo - 1, 0);
            renderizarResultados();
        } else if (e.key === 'Enter') {
            if (indiceActivo >= 0 && resultados[indiceActivo]) {
                e.preventDefault();
                seleccionarMaterial(resultados[indiceActivo]);
            }
        } else if (e.key === 'Escape') {
            cerrarDropdown();
        }
    });

    document.addEventListener('click', function(e) {
        if (e.target !== input && !dropdown.contains(e.target)) {
            cerrarDropdown();
        }
    });
}
