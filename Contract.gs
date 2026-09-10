function procesarContratos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Contratos');
  const filas = hoja.getDataRange().getValues();
  const usuarios = cargarUsuarios();
  const cal = CalendarApp.getDefaultCalendar();

  let avisos = 0;

  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    const cliente = f[0];
    if (!cliente) continue;

    const equipo = f[1];
    const cantidad = f[2];
    const vence = f[5];
    const ejecutivo = f[6];
    const alertaEnviada = f[9];
    const eventoCreado = f[10];

    const dias = diasRestantes(vence);
    const fila = i + 1;

    // columnas: Estado(8) Dias restantes(9) Alerta(10) Evento(11)
    hoja.getRange(fila, 9).setValue(dias);
    hoja.getRange(fila, 8).setValue(dias < 0 ? 'Vencido' : 'Vigente');

    const color = colorPorDias(dias, AVISO_CONTRATO_DIAS);
    hoja.getRange(fila, 1, 1, 11).setBackground(color);

    // solo actua dentro de la ventana de aviso y si no se habia avisado antes
    if (dias <= AVISO_CONTRATO_DIAS && alertaEnviada !== 'Si') {
      const datos = usuarios[ejecutivo];
      if (!datos) {
        registrar('procesarContratos', 'Sin destinatario', ejecutivo + ' no esta en Usuarios');
        continue;
      }

      const asunto = 'Contrato por vencer: ' + cliente + ' (' + dias + ' dias)';
      const cuerpo =
        'Hola ' + ejecutivo + ',\n\n' +
        'El contrato de ' + cliente + ' vence el ' +
        Utilities.formatDate(new Date(vence), 'GMT-6', 'dd/MM/yyyy') + ' (' + dias + ' dias).\n' +
        'Equipo: ' + cantidad + ' x ' + equipo + '\n\n' +
        'Acciones: contactar al cliente para renovar o programar la recoleccion del equipo.\n\n' +
        'Sistema Equipa TI';

      GmailApp.sendEmail(datos.correo, asunto, cuerpo);
      hoja.getRange(fila, 10).setValue('Si');
      registrar('procesarContratos', 'Correo enviado', cliente + ' -> ' + datos.correo);
      avisos++;

      if (eventoCreado !== 'Si') {
        const fechaReunion = new Date(vence);
        fechaReunion.setDate(fechaReunion.getDate() - 5);
        cal.createAllDayEvent(
          'Renovacion: ' + cliente,
          fechaReunion,
          { description: equipo + ' (' + cantidad + '). Responsable: ' + ejecutivo, guests: datos.correo }
        );
        hoja.getRange(fila, 11).setValue('Si');
        registrar('procesarContratos', 'Evento creado', cliente + ' el ' +
          Utilities.formatDate(fechaReunion, 'GMT-6', 'dd/MM/yyyy'));
      }
    }
  }

  registrar('procesarContratos', 'Fin', avisos + ' avisos enviados');
}
