const AVISO_CONTRATO_DIAS = 30;
const AVISO_MANTTO_DIAS = 7;

const VERDE = '#d9ead3';
const AMARILLO = '#fff2cc';
const ROJO = '#f4cccc';

function diasRestantes(fecha) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const objetivo = new Date(fecha);
  objetivo.setHours(0, 0, 0, 0);
  return Math.round((objetivo - hoy) / (1000 * 60 * 60 * 24));
}

function colorPorDias(dias, umbral) {
  if (dias <= 7 || dias < 0) return ROJO;
  if (dias <= umbral) return AMARILLO;
  return VERDE;
}

function registrar(funcion, accion, detalle) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  hoja.appendRow([new Date(), funcion, accion, detalle]);

  // ademas del log en la hoja, mando el evento a Cloud Logging del proyecto de GCP
  // asi queda un registro centralizado que el usuario no puede alterar desde la hoja
  console.info(JSON.stringify({
    origen: 'EquipaTI',
    funcion: funcion,
    accion: accion,
    detalle: detalle
  }));
}

function cargarUsuarios() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const filas = hoja.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < filas.length; i++) {
    const nombre = filas[i][0];
    if (!nombre) continue;
    // indexo por nombre porque en Contratos/Mantenimientos el responsable viene por nombre, no por correo
    mapa[nombre] = { correo: filas[i][1], rol: filas[i][2] };
  }
  return mapa;
}

function correoPorRol(usuarios, rol) {
  for (const nombre in usuarios) {
    if (usuarios[nombre].rol === rol) return usuarios[nombre].correo;
  }
  return null;
}
