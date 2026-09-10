/**
 * Configuracion del entorno de pruebas
 * Se ejecuta solo una vez al inicio
 */

// correo base para los alias del ambiente de pruebas
const CORREO_BASE = 'dabcd6011';

function configurarEntorno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  crearContratos(ss);
  crearMantenimientos(ss);
  crearUsuarios(ss);
  crearLog(ss);

  // elimina la hoja vacia que Sheets crea por defecto
  ['Hoja 1', 'Hoja1', 'Sheet1'].forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (h) ss.deleteSheet(h);
  });

  ss.setActiveSheet(ss.getSheetByName('Contratos'));
}

/** Borra la hoja si ya existe y la vuelve a crear, para poder reconstruir el entorno */
function hojaLimpia(ss, nombre) {
  const previa = ss.getSheetByName(nombre);
  if (previa) ss.deleteSheet(previa);
  return ss.insertSheet(nombre);
}

/** Aplica formato al encabezado y congela la primera fila */
function formatoEncabezado(hoja, columnas) {
  const rango = hoja.getRange(1, 1, 1, columnas);
  rango.setFontWeight('bold').setBackground('#e8eaed');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, columnas);
}

function crearContratos(ss) {
  const hoja = hojaLimpia(ss, 'Contratos');

  hoja.getRange(1, 1, 1, 11).setValues([[
    'Cliente', 'Equipo', 'Cantidad', 'Renta mensual', 'Inicio', 'Vence',
    'Ejecutivo', 'Estado', 'Dias restantes', 'Alerta enviada', 'Evento creado'
  ]]);

  const datos = [
    ['Grupo Textil MX', 'Multifuncional Ricoh MP 2555', 8, 12800, new Date(2024, 9, 1), new Date(2026, 8, 12), 'Ana Ruiz'],
    ['Farmacias del Norte', 'Multifuncional HP LaserJet E77822', 4, 9200, new Date(2024, 8, 15), new Date(2026, 8, 15), 'Luis Vega'],
    ['Constructora Aper', 'Laptop Dell Latitude 5440', 25, 37500, new Date(2024, 8, 20), new Date(2026, 8, 25), 'Ana Ruiz'],
    ['Hotel Costa Azul', 'Multifuncional Xerox VersaLink C7025', 3, 7800, new Date(2024, 9, 5), new Date(2026, 9, 5), 'Luis Vega'],
    ['Distribuidora Rios', 'Servidor Dell PowerEdge T150', 2, 15400, new Date(2024, 10, 20), new Date(2026, 10, 20), 'Ana Ruiz'],
    ['Colegio Monteverde', 'Desktop HP ProDesk 400', 40, 28000, new Date(2024, 11, 15), new Date(2026, 11, 15), 'Luis Vega'],
    ['Corporativo Lerma', 'Plotter HP DesignJet T650', 1, 6500, new Date(2025, 0, 10), new Date(2027, 0, 10), 'Ana Ruiz'],
    ['Distribuidora Rios', 'Switch Cisco Catalyst 1000', 6, 4800, new Date(2024, 8, 1), new Date(2026, 8, 8), 'Luis Vega']
  ];

  hoja.getRange(2, 1, datos.length, 7).setValues(datos);
  hoja.getRange(2, 4, datos.length, 1).setNumberFormat('$#,##0');
  hoja.getRange(2, 5, datos.length, 2).setNumberFormat('yyyy-mm-dd');
  formatoEncabezado(hoja, 11);
}

function crearMantenimientos(ss) {
  const hoja = hojaLimpia(ss, 'Mantenimientos');

  hoja.getRange(1, 1, 1, 11).setValues([[
    'Cliente', 'Equipo', 'Numero de serie', 'Ultimo servicio', 'Frecuencia (dias)',
    'Proximo servicio', 'Tecnico', 'Estado', 'Dias restantes', 'Alerta enviada', 'Evento creado'
  ]]);

  const datos = [
    ['Grupo Textil MX', 'Multifuncional Ricoh MP 2555', 'RIC-2555-018', new Date(2026, 7, 8), 30, 'Carlos Mena'],
    ['Farmacias del Norte', 'Multifuncional HP LaserJet E77822', 'HP-77822-004', new Date(2026, 7, 12), 30, 'Carlos Mena'],
    ['Constructora Aper', 'Servidor Dell PowerEdge T150', 'DEL-T150-002', new Date(2026, 7, 15), 30, 'Sofia Ramos'],
    ['Grupo Textil MX', 'Laptop Dell Latitude 5440', 'DEL-5440-L01', new Date(2026, 6, 28), 45, 'Sofia Ramos'],
    ['Hotel Costa Azul', 'Multifuncional Xerox VersaLink C7025', 'XER-7025-009', new Date(2026, 7, 20), 30, 'Carlos Mena'],
    ['Distribuidora Rios', 'NAS Synology DS923+', 'SYN-923-001', new Date(2026, 7, 28), 30, 'Sofia Ramos'],
    ['Colegio Monteverde', 'No-break APC Smart-UPS 1500', 'APC-1500-011', new Date(2026, 6, 20), 60, 'Sofia Ramos'],
    ['Corporativo Lerma', 'Plotter HP DesignJet T650', 'HP-T650-003', new Date(2026, 7, 30), 30, 'Carlos Mena'],
    ['Distribuidora Rios', 'Switch Cisco Catalyst 1000', 'CIS-C1000-007', new Date(2026, 8, 2), 30, 'Sofia Ramos'],
    ['Colegio Monteverde', 'Desktop HP ProDesk 400', 'HP-PD400-L02', new Date(2026, 6, 30), 45, 'Sofia Ramos']
  ];

  // las columnas van A,B,C,D,E y el tecnico en G, por eso se escriben en dos bloques
  const base = datos.map(function (f) { return [f[0], f[1], f[2], f[3], f[4]]; });
  const tecnicos = datos.map(function (f) { return [f[5]]; });

  hoja.getRange(2, 1, base.length, 5).setValues(base);
  hoja.getRange(2, 7, tecnicos.length, 1).setValues(tecnicos);
  hoja.getRange(2, 4, base.length, 1).setNumberFormat('yyyy-mm-dd');
  hoja.getRange(2, 6, base.length, 1).setNumberFormat('yyyy-mm-dd');
  formatoEncabezado(hoja, 11);
}

function crearUsuarios(ss) {
  const hoja = hojaLimpia(ss, 'Usuarios');

  hoja.getRange(1, 1, 1, 5).setValues([[
    'Nombre', 'Correo', 'Rol', 'Area', 'Grupo asignado'
  ]]);

  // el script no tiene correos escritos en el codigo: los lee de aqui.
  // cambiar un responsable es cambiar una fila, no tocar el script.
  const datos = [
    ['Ana Ruiz', CORREO_BASE + '+ana@gmail.com', 'Ejecutivo de cuenta', 'Comercial', 'Cartera Norte'],
    ['Luis Vega', CORREO_BASE + '+luis@gmail.com', 'Ejecutivo de cuenta', 'Comercial', 'Cartera Sur'],
    ['Carlos Mena', CORREO_BASE + '+carlos@gmail.com', 'Tecnico de soporte', 'Soporte', 'Impresion'],
    ['Sofia Ramos', CORREO_BASE + '+sofia@gmail.com', 'Tecnico de soporte', 'Soporte', 'Computo e infraestructura'],
    ['Diana Perez', CORREO_BASE + '+coord@gmail.com', 'Coordinador de soporte', 'Soporte', 'Todas'],
    ['Jorge Salas', CORREO_BASE + '+gerente@gmail.com', 'Gerente comercial', 'Comercial', 'Todas']
  ];

  hoja.getRange(2, 1, datos.length, 5).setValues(datos);
  formatoEncabezado(hoja, 5);
}

function crearLog(ss) {
  const hoja = hojaLimpia(ss, 'Log');
  hoja.getRange(1, 1, 1, 4).setValues([['Fecha y hora', 'Funcion', 'Accion', 'Detalle']]);
  formatoEncabezado(hoja, 4);
}
