function realizarCopiaConFecha() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var respaldoSheet = spreadsheet.getSheetByName('Respaldo');
  var destinoSheetName = obtenerFechaActual(); // Obtiene la fecha actual en el formato "dd-MM-yyyy".
  
  // Obtén los datos de la hoja original
  var respaldoData = respaldoSheet.getDataRange().getValues();
  var numRows = respaldoData.length;
  var numCols = respaldoData[0].length;
  
  // Crea una nueva hoja con el nombre de la fecha actual
  spreadsheet.insertSheet(destinoSheetName);
  var destinoSheet = spreadsheet.getSheetByName(destinoSheetName);

  // Copia los valores a la nueva hoja
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      destinoSheet.getRange(i + 1, j + 1).setValue(respaldoData[i][j]);
    }
  }
  
  // Opcional: Puedes establecer la nueva hoja como activa si lo deseas
  spreadsheet.setActiveSheet(destinoSheet);
}

function obtenerFechaActual() {
  var timeZone = Session.getScriptTimeZone();
  var fecha = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy");
  return fecha;
}
