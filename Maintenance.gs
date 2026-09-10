function procesarMantenimientos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Mantenimientos');
  const filas = hoja.getDataRange().getValues();
  const usuarios = cargarUsuarios();
  const cal = CalendarApp.getDefaultCalendar();

  let avisos = 0;

  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    const cliente = f[0];
    if (!cliente) continue;

    const equipo = f[1];
    const serie = f[2];
    const ultimo = f[3];
    const frecuencia = f[4];
    const tecnico = f[6];
    const alertaEnviada = f[9];
    const eventoCreado = f[10];

    // el proximo servicio no se captura: se calcula desde el ultimo + la frecuencia
    const proximo = new Date(ultimo);
    proximo.setDate(proximo.getDate() + frecuencia);
    const dias = diasRestantes(proximo);
    const fila = i + 1;

    // columnas: Proximo(6) Estado(8) Dias(9) Alerta(10) Evento(11)
    hoja.getRange(fila, 6).setValue(proximo).setNumberFormat('yyyy-mm-dd');
    hoja.getRange(fila, 9).setValue(dias);
    hoja.getRange(fila, 8).setValue(dias < 0 ? 'Vencido' : 'Programado');

    hoja.getRange(fila, 1, 1, 11).setBackground(colorPorDias(dias, AVISO_MANTTO_DIAS));

    if (dias <= AVISO_MANTTO_DIAS && alertaEnviada !== 'Si') {
      const datos = usuarios[tecnico];
      if (!datos) {
        registrar('procesarMantenimientos', 'Sin destinatario', tecnico + ' no esta en Usuarios');
        continue;
      }

      const asunto = 'Mantenimiento programado: ' + equipo + ' (' + cliente + ')';
      const cuerpo =
        'Hola ' + tecnico + ',\n\n' +
        'El equipo ' + equipo + ' (serie ' + serie + ') del cliente ' + cliente +
        ' requiere servicio el ' +
        Utilities.formatDate(proximo, 'GMT-6', 'dd/MM/yyyy') + ' (' + dias + ' dias).\n\n' +
        'Sistema Equipa TI';

      GmailApp.sendEmail(datos.correo, asunto, cuerpo);
      hoja.getRange(fila, 10).setValue('Si');
      registrar('procesarMantenimientos', 'Correo enviado', cliente + '/' + equipo + ' -> ' + datos.correo);
      avisos++;

      if (eventoCreado !== 'Si') {
        cal.createAllDayEvent(
          'Mantenimiento: ' + equipo + ' (' + cliente + ')',
          proximo,
          { description: 'Serie ' + serie + '. Tecnico: ' + tecnico, guests: datos.correo }
        );
        hoja.getRange(fila, 11).setValue('Si');
        registrar('procesarMantenimientos', 'Evento creado', equipo + ' el ' +
          Utilities.formatDate(proximo, 'GMT-6', 'dd/MM/yyyy'));
      }
    }
  }

  registrar('procesarMantenimientos', 'Fin', avisos + ' avisos enviados');
}
