
const { app, BrowserWindow, Menu, ipcMain } = require('electron');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const xlsx_populate = require('xlsx-populate');
const { execSync } = require('child_process');

let ventanaPrincipal;
let ventanaActividad;

const db = new sqlite3.Database(path.join(__dirname, 'BaseDeDatos/baseDeDatos.db'), sqlite3.OPEN_READWRITE, (err) => {
    if (err) {
        console.log(err);
    }
});

function consultarYenviarDatos() {
    const sql = `
            SELECT nombre_apellido,id_pasajero,obra_social FROM pasajeros
            ORDER BY obra_social
        `;

    db.all(sql, [], (err, rows) => {
        if (err) {
            console.log(err)
            return;
        }
        let promises = [];
        rows.forEach(row => {
            promises.push(new Promise((resolve, reject) => {
                let htmlCode = `
                            <div class="pasajero">
                                <span>${row.nombre_apellido}</span>
                                <div class="obso">${row.obra_social}</div>
                            `;

                const sqlActividades = `
                            SELECT id_actividad, nombre_actividad FROM actividades
                            WHERE id_pasajero = ${row.id_pasajero}
                        `;
                db.all(sqlActividades, [], (e, rows2) => {
                    if (e) {
                        console.log(e);
                        reject(e);
                    } else {
                        rows2.forEach(row2 => {
                            htmlCode += `
                                    <div class="actividad">
                                        <span>${row2.nombre_actividad}</span>
                                        <i id="${row2.id_actividad}" class="btn fa-solid fa-right-to-bracket"></i>
                                    </div>
                                    `;
                        });
                        htmlCode += '</div>';
                        resolve(htmlCode);

                    }
                });
            }));
        });
        
        Promise.all(promises)
            .then(htmlCodes => {
                ventanaPrincipal.webContents.send('nuevoPasajero', htmlCodes.join(''));
            })
            .catch(e=>{
                console.error(e);
            })
        
    });
}

function consultarActividad(id_actividad) {
    const sql = `
        SELECT nombre_apellido,id_actividad,nombre_actividad FROM actividades AS a
        INNER JOIN pasajeros AS p ON a.id_pasajero = p.id_pasajero
        WHERE id_actividad = ${id_actividad}
    `;

    db.get(sql,[],(e,row)=>{
        if (e) return e;
        ventanaActividad.webContents.send('actividadConsultada',row);
    });
}

async function consultarRutaYOs(id_actividad) {
    try {
        const row = await new Promise((resolve,reject)=>{
            const sql = `
                SELECT ruta_planilla,obra_social FROM actividades as a
                INNER JOIN pasajeros as p ON a.id_pasajero = p.id_pasajero
                WHERE a.id_actividad = ${id_actividad}
            `;
            db.get(sql,[],(e,fila)=>{
                if (e) {
                    reject(e)
                } else {
                    resolve(fila)
                }
            });
        });

        return row;
        
    } catch (e) {
        console.log(e)
    }
}

function crearVentana() {
    ventanaPrincipal = new BrowserWindow({
        width: 500,
        height: 600,
        show: false,
        webPreferences: {
            preload: path.join(__dirname,'preload.js'),
        },

    })

    ventanaPrincipal.loadFile('index.html');

    ventanaPrincipal.on('ready-to-show',()=>{
        ventanaPrincipal.show()
    });

    ventanaPrincipal.webContents.on('dom-ready', () => {
        consultarYenviarDatos();
    });

    ventanaPrincipal.on('closed',()=>{
        app.quit();
    })
}

function configMenu(ventana) {
    const menuTemplate = [
        {
            label: 'devTools',
            submenu: [
                {
                    label: 'Toggle DevTools',
                    accelerator: 'Ctrl+Shift+I',
                    click: ()=>{
                        ventana.webContents.toggleDevTools();
                    }
                }
            ]
        }
    ];

    const menu = Menu.buildFromTemplate(menuTemplate);
    Menu.setApplicationMenu(menu)
}

function crearVentanaAct(id_actividad) {
    ventanaActividad = new BrowserWindow({
        width: 350,
        height: 470,
        show: false,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js')
        }
    });

    ventanaActividad.loadFile('actividad.html');

    ventanaActividad.on('ready-to-show',()=>{
        ventanaActividad.show();
    });

    ventanaActividad.webContents.on('dom-ready',()=>{
        consultarActividad(id_actividad);
    })
}

function getMes(mes) {
    const meses = ['.','enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
    const numMes = Number(mes);
    return meses[numMes];
}

function esBisiesto (anoActual){
    return ((anoActual % 4 === 0 && anoActual % 100 !== 0) || anoActual % 400 === 0);
}

async function modificarExcelOscoema(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('G1').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('H'+String(i)).value("");
    }
    sheet.cell('C5').value(String(anoActual));
    sheet.cell('B5').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('G1').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'H'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fecha = `${diaStr}/${mesFecha}/${anoActual}`;

        if (dia <= 15) {
            sheet.cell(`A${dia + 9}`).value(fecha);
        } else {
            sheet.cell(`D${dia - 6}`).value(fecha);
        }
    }

    if (esBisiesto(anoActual)) {
        sheet.cell('D23').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('D24').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('D25').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('D25').value("");
            }
        } else {
            sheet.cell('D24').value("");
            sheet.cell('D25').value("");
        }
    } else {
        if (mes != '2') {
            sheet.cell('D23').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('D24').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('D25').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('D25').value("");
            }
        } else {
            sheet.cell('D23').value("");
            sheet.cell('D24').value("");
            sheet.cell('D25').value("");
        }
    }
    

    if (pdf) {
        sheet.cell('J1').value('pdf');
    }
    if (imp) {
        sheet.cell('K1').value('imp');
    }

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function modificarExcelIoma(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('J1').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('K'+String(i)).value("");
    }

    sheet.cell('C7').value(String(anoActual));
    sheet.cell('B7').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('J1').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'K'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fila = dia + 11;
        
        sheet.cell(`A${fila}`).value(`${diaStr}/${mesFecha}/${anoActual}`);
    }
    
    if (esBisiesto(anoActual)) {
        sheet.cell('A40').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('A41').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('A42').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('A42').value("");
            }
        } else {
            sheet.cell('A41').value("");
            sheet.cell('A42').value("");
        }
    } else {
        if (mes != '2') {
            sheet.cell('A40').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('A41').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('A42').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('A42').value("");
            }
        } else {
            sheet.cell('A40').value("");
            sheet.cell('A41').value("");
            sheet.cell('A42').value("");
        }
    }

    if (pdf) {
        sheet.cell('L1').value('pdf');
    }
    if (imp) {
        sheet.cell('M1').value('imp');
    };

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function modificarExcelOsecac(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('L2').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('M'+String(i)).value("");
    }

    sheet.cell('C7').value(String(anoActual));
    sheet.cell('B7').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('L2').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'M'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fila = dia + 1;
        
        sheet.cell(`I${fila}`).value(`${diaStr}/${mesFecha}/${anoActual}`);
    }
    
    if (esBisiesto(anoActual)) {
        sheet.cell('I30').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('I31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('I32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('I32').value("");
            }
        } else {
            sheet.cell('I31').value("");
            sheet.cell('I32').value("");
        }
    } else {
        if (mes != '2') {
            sheet.cell('I30').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('I31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('I32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('I32').value("");
            }
        } else {
            sheet.cell('I30').value("");
            sheet.cell('I31').value("");
            sheet.cell('I32').value("");
        }
    }

    if (pdf) {
        sheet.cell('N1').value('pdf');
    }
    if (imp) {
        sheet.cell('O1').value('imp');
    };

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function modificarExcelAmtar(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('L2').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('M'+String(i)).value("");
    }

    sheet.cell('D9').value(String(anoActual));
    sheet.cell('C9').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('L2').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'M'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fila = dia + 1;
        
        sheet.cell(`I${fila}`).value(`${diaStr}/${mesFecha}/${anoActual}`);
    }
    
    if (esBisiesto(anoActual)) {
        sheet.cell('I30').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('I31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('I32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('I32').value("");
            }
        } else {
            sheet.cell('I31').value("");
            sheet.cell('I32').value("");
        }
    } else {
        if (mes != '2') {
            sheet.cell('I30').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('I31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('I32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('I32').value("");
            }
        } else {
            sheet.cell('I30').value("");
            sheet.cell('I31').value("");
            sheet.cell('I32').value("");
        }
    }

    if (pdf) {
        sheet.cell('N1').value('pdf');
    }
    if (imp) {
        sheet.cell('O1').value('imp');
    };

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function modificarExcelOstpcphyara(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('N2').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('O'+String(i)).value("");
    }

    sheet.cell('J6').value(String(anoActual));
    sheet.cell('I6').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('N2').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'O'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fila = dia + 1;
        
        sheet.cell(`K${fila}`).value(`${diaStr}/${mesFecha}/${anoActual}`);
    }
    
    if (esBisiesto(anoActual)) {
        sheet.cell('K30').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('K31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('K32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('K32').value("");
            }
        } else {
            sheet.cell('K31').value("");
            sheet.cell('K32').value("");
        }
    } else {
        if (mes != '2') {
            sheet.cell('K30').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('K31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('K32').value('31' + '/' + mesFecha + '/' + String(anoActual));
            } else {
                sheet.cell('K32').value("");
            }
        } else {
            sheet.cell('K30').value("");
            sheet.cell('K31').value("");
            sheet.cell('K32').value("");
        }
    }

    if (pdf) {
        sheet.cell('P1').value('pdf');
    }
    if (imp) {
        sheet.cell('Q1').value('imp');
    };

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function modificarExcelIosfa(feriados,mes,ruta_planilla,imp,pdf) {
    const ruta = path.join(__dirname,ruta_planilla);
    const anoActual = new Date().getFullYear();
    const workbook = await xlsx_populate.fromFileAsync(ruta);
    const sheet = workbook.sheet('hoja1');
    const cantFerAnt = sheet.cell('N2').value();
    for (let i = 1; i <= cantFerAnt; i++) {
        sheet.cell('O'+String(i)).value("");
    }

    sheet.cell('J1').value(getMes(mes));
    let mesFecha = mes;
    if (Number(mes)<10) {
        mesFecha = '0' + mes
    }

    sheet.cell('N2').value(String(feriados.length));
    for (let i = 0; i < feriados.length; i++) {
        let feriado;
        if (Number(feriados[i]) < 10) {
            feriado = '0' + feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        } else {
            feriado = feriados[i] + '/' + mesFecha + '/' + String(anoActual); 
        }
        celda = 'O'+ String(i+1);
        sheet.cell(celda).value(feriado)
    }

    for (let dia = 1; dia <= 28; dia++) {
        const diaStr = String(dia).padStart(2, '0');
        const fila = dia + 1;
        
        sheet.cell(`K${fila}`).value(`${diaStr}/${mesFecha}/${anoActual}`);
    }   
    
    if (esBisiesto(anoActual)) {
        sheet.cell('K30').value('29' + '/' + mesFecha + '/' + String(anoActual));
        if (mes != '2') {
            sheet.cell('K31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('K32').value('31' + '/' + mesFecha + '/' + String(anoActual));
                sheet.cell('G5').value('Mar del plata, 31 de '+ getMes(mes) + ' ' + String(anoActual));
            } else {
                sheet.cell('K32').value("");
                sheet.cell('G5').value('Mar del plata, 30 de '+ getMes(mes) + ' ' + String(anoActual));
            }
        } else {
            sheet.cell('K31').value("");
            sheet.cell('K32').value("");
            sheet.cell('G5').value('Mar del plata, 29 de '+ getMes(mes) + ' ' + String(anoActual));
        }
    } else {
        if (mes != '2') {
            sheet.cell('K30').value('29' + '/' + mesFecha + '/' + String(anoActual));
            sheet.cell('K31').value('30' + '/' + mesFecha + '/' + String(anoActual));
            if (['1','3','5','7','8','10','12'].includes(mes)) {
                sheet.cell('K32').value('31' + '/' + mesFecha + '/' + String(anoActual));
                sheet.cell('G5').value('Mar del plata, 31 de '+ getMes(mes) + ' ' +  String(anoActual));
            } else {
                sheet.cell('K32').value("");
                sheet.cell('G5').value('Mar del plata, 30 de '+ getMes(mes) + ' ' + String(anoActual));
            }
        } else {
            sheet.cell('K30').value("");
            sheet.cell('K31').value("");
            sheet.cell('K32').value("");
            sheet.cell('G5').value('Mar del plata, 28 de '+ getMes(mes) + ' ' + String(anoActual));
        }
    }

    sheet.cell('A13').value('Kilometros realizados en el mes de ' + getMes(mes) + ' ' + String(anoActual) + ':');

    if (pdf) {
        sheet.cell('N1').value('pdf');
    }
    if (imp) {
        sheet.cell('O1').value('imp');
    };

    await workbook.toFileAsync(ruta);

    await executeExcel(ruta)
}

async function executeExcel(workbookPath) {
    try {
        execSync(`start excel.exe "${workbookPath}" /e`);
    } catch (error) {
        console.error('Error al ejecutar la macro de VBA:', error);
    }
}

async function elegirOs(datos) {
    const rutaYOs = await consultarRutaYOs(datos.id_actividad);
    const { obra_social, ruta_planilla } = rutaYOs;
    const { feriados, mes, imp, pdf } = datos;

    switch (obra_social) {
        case 'OSCOEMA':
            modificarExcelOscoema(feriados, mes, ruta_planilla, imp, pdf);
            break;
        case 'IOMA':
            modificarExcelIoma(feriados, mes, ruta_planilla, imp, pdf);
            break;
        case 'OSECAC':
            modificarExcelOsecac(feriados, mes, ruta_planilla, imp, pdf);
            break;
        case 'AMTAR':
            modificarExcelAmtar(feriados, mes, ruta_planilla, imp, pdf);
            break;
        case 'OSTPCPHYARA':
            modificarExcelOstpcphyara(feriados, mes, ruta_planilla, imp, pdf);
            break;
        case 'IOSFA':
            modificarExcelIosfa(feriados, mes, ruta_planilla, imp, pdf);
            break;
        default:
            console.warn(`Obra social no reconocida: ${obra_social}`);
            break;
    }
}

app.whenReady().then(()=>{
    crearVentana();
    configMenu(ventanaPrincipal);
});

app.on('window-all-closed', () => {
    app.quit()
})

ipcMain.on('actividadSeleccionada',(e,id_actividad)=>{
    crearVentanaAct(id_actividad);
    configMenu(ventanaActividad);
});

ipcMain.on('datosPlanilla',(e,datos)=>{
    ventanaActividad.close();
    elegirOs(datos);
});
