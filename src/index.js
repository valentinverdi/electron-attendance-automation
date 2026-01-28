const container = document.querySelector('.container');

window.electronAPI.onNuevoPasajero((htmlCode)=>{
    container.innerHTML += htmlCode;
    
    const btns = document.querySelectorAll('.btn');
    btns.forEach(btn=>{
        btn.addEventListener('click',()=>{
            const id_actividad = btn.getAttribute('id');
            window.electronAPI.sendID(id_actividad);
        });
    });
});


