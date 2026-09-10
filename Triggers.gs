function instalarTriggers() {
  // limpio los que ya existan para no duplicar si se corre dos veces
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('rutinaDiaria')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  ScriptApp.newTrigger('alEditar')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  registrar('instalarTriggers', 'Triggers instalados', 'diario 8am + onEdit');
}

function rutinaDiaria() {
  procesarContratos();
  procesarMantenimientos();
  resumenDiario();
}

function alEditar(e) {
  if (!e || !e.range) return;

  const hoja = e.range.getSheet().getName();
  // solo recalculo si editan una fecha en Contratos o Mantenimientos, no en cada tecla
  if (hoja === 'Contratos') {
    procesarContratos();
    registrar('alEditar', 'Recalculo', 'Contratos fila ' + e.range.getRow());
  } else if (hoja === 'Mantenimientos') {
    procesarMantenimientos();
    registrar('alEditar', 'Recalculo', 'Mantenimientos fila ' + e.range.getRow());
  }
}
