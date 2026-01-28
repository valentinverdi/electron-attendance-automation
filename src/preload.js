const { ipcRenderer , contextBridge } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    onNuevoPasajero: (callback) => ipcRenderer.on('nuevoPasajero',(e,value)=> callback(value)),
    sendID: (value) => ipcRenderer.send('actividadSeleccionada',value),
    onActividadConsultada: (callback) => ipcRenderer.on('actividadConsultada',(e,value)=>callback(value)),
    sendDatosPlanilla: (value) => ipcRenderer.send('datosPlanilla',value)
});