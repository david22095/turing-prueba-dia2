function resumenDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usuarios = cargarUsuarios();

  const contratos = contarPorSemaforo('Contratos', 5, AVISO_CONTRATO_DIAS);
  const mantto = contarPorSemaforo('Mantenimientos', null, AVISO_MANTTO_DIAS);

  const gerente = correoPorRol(usuarios, 'Gerente comercial');
  const coord = correoPorRol(usuarios, 'Coordinador de soporte');

  if (gerente) {
    GmailApp.sendEmail(gerente, 'Resumen diario de contratos',
      'Estado de la cartera al ' + Utilities.formatDate(new Date(), 'GMT-6', 'dd/MM/yyyy') + ':\n\n' +
      'Vigentes (verde): ' + contratos.verde + '\n' +
      'Por vencer (amarillo): ' + contratos.amarillo + '\n' +
      'Criticos o vencidos (rojo): ' + contratos.rojo + '\n\n' +
      'Sistema Equipa TI');
    registrar('resumenDiario', 'Resumen contratos', 'Enviado a ' + gerente);
  }

  if (coord) {
    GmailApp.sendEmail(coord, 'Resumen diario de mantenimientos',
      'Estado de servicios al ' + Utilities.formatDate(new Date(), 'GMT-6', 'dd/MM/yyyy') + ':\n\n' +
      'Al dia (verde): ' + mantto.verde + '\n' +
      'Proximos (amarillo): ' + mantto.amarillo + '\n' +
      'Vencidos o urgentes (rojo): ' + mantto.rojo + '\n\n' +
      'Sistema Equipa TI');
    registrar('resumenDiario', 'Resumen mantto', 'Enviado a ' + coord);
  }
}

// columnaVence: indice (base 0) de la fecha en Contratos; en Mantenimientos se recalcula
function contarPorSemaforo(nombreHoja, columnaVence, umbral) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  const filas = hoja.getDataRange().getValues();
  const cuenta = { verde: 0, amarillo: 0, rojo: 0 };

  for (let i = 1; i < filas.length; i++) {
    if (!filas[i][0]) continue;

    let dias;
    if (columnaVence !== null) {
      dias = diasRestantes(filas[i][columnaVence]);
    } else {
      const proximo = new Date(filas[i][3]);
      proximo.setDate(proximo.getDate() + filas[i][4]);
      dias = diasRestantes(proximo);
    }

    const color = colorPorDias(dias, umbral);
    if (color === ROJO) cuenta.rojo++;
    else if (color === AMARILLO) cuenta.amarillo++;
    else cuenta.verde++;
  }
  return cuenta;
}
