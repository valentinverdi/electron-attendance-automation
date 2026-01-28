const nombre = document.querySelector('.nombre');
const actividad = document.querySelector('.actividad');
const mes = document.querySelector('#mes');
const mas = document.querySelector('.mas');
const menos = document.querySelector('.menos');
const feriadosContainer = document.querySelector('.feriados');
const enviar = document.querySelector('.enviar');
const advertencia = document.querySelector('.advertencia');
const advPdfImp = document.querySelector('.advPdfImp');
const imp = document.querySelector('#imp');
const pdf = document.querySelector('#pdf');

window.electronAPI.onActividadConsultada((row)=>{
    nombre.innerHTML = row.nombre_apellido;
    actividad.innerHTML = row.nombre_actividad;
    actividad.setAttribute('id',row.id_actividad);
});

mas.addEventListener('click',()=>{
    let input = document.createElement('input');
    input.setAttribute('class','dia');
    input.setAttribute('type','text');
    feriadosContainer.appendChild(input);
    if (advertencia.innerHTML != '') {
        advertencia.innerHTML = '';
    }
});

menos.addEventListener('click',()=>{
    const ultInput = feriadosContainer.lastElementChild;
    if (ultInput) {
        ultInput.remove();
    }
});

enviar.addEventListener('click',()=>{
    advPdfImp.innerHTML = '';
    const feriadosElem = document.querySelectorAll('.dia');
    let feriadosArray = [];
    feriadosElem.forEach(dia=>{
        feriadosArray.push(dia.value);
    });

    if (verificarFeriadosRango(feriadosArray,mes.value) && verificarFeriadosRepetidos(feriadosArray,0) && (imp.checked || pdf.checked)) {
        const datos = {
            id_actividad: actividad.getAttribute('id'),
            mes: mes.value,
            feriados: feriadosArray,
            imp: imp.checked,
            pdf: pdf.checked
        }
        console.log(datos)
        window.electronAPI.sendDatosPlanilla(datos);
    } else {
        if (!imp.checked && !pdf.checked) {
            advPdfImp.innerHTML = 'Es necesario marcar al menos una opción'
        }
        if (!verificarFeriadosRango(feriadosArray,mes.value) || !verificarFeriadosRepetidos(feriadosArray,0)) {
            advertencia.innerHTML = 'Feriados ingresados erroneamente';
            feriadosContainer.innerHTML = '';
        }
    }
});

function verificarFeriadosRango(array,mes) {
    let valid = true;
    for (let i = 0; i < array.length; i++) {
        if (Number(array[i])>=1) {
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                if (Number(array[i])>31) {
                    valid = false;
                    break
                }
            } else if (['4','6','9','11'].includes(mes)) {
                if (Number(array[i])>30) {
                    valid = false;
                    break
                }
                
            } else if (mes = '2') {
                if (Number(array[i])>28) {
                    valid = false;
                    break
                }
            }
        } else {
            valid = false;
            break 
        }
    }
    return valid;
}

function verificarFeriadosRepetidos(array, inicio = 0) {
    if (array.length === 0 || array.length === 1) {
        return true;
    }

    const elem = array[inicio];
    let valid = true;

    for (let i = inicio + 1; i < array.length; i++) {
        if (elem == array[i]) {
            valid = false;
            break;
        }
    }

    if (inicio + 1 >= array.length) {
        return valid;
    }

    if (!valid) {
        return valid;
    } else {
        return verificarFeriadosRepetidos(array, inicio + 1);
    }
}
