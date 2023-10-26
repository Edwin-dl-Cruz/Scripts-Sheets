// Función para borrar registros en hojas específicas.
function borrarRegistros() {
  // Lista de nombres de hojas a borrar.
  var hojasABorrar = ['Primer', 'Segundo', 'Tercero'];

  // Acceder a la hoja de cálculo activa.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Iterar a través de la lista de hojas a borrar.
  for (var i = 0; i < hojasABorrar.length; i++) {
    var nombreHoja = hojasABorrar[i];

    // Obtener la hoja por nombre.
    var hoja = spreadsheet.getSheetByName(nombreHoja);

    if (hoja) {
      // Verificar si la hoja existe.
      var numRows = hoja.getLastRow();

      if (numRows > 1) {
        // Verificar si la hoja tiene más de una fila de datos.
        var numColumns = hoja.getLastColumn();

        // Limpiar el contenido de todas las filas excepto la primera (encabezados).
        hoja.getRange(2, 1, numRows - 1, numColumns).clearContent();
      }
    } else {
      // Registrar un mensaje en el registro si la hoja no se encontró.
      Logger.log("La hoja '" + nombreHoja + "' no se encontró en la hoja de cálculo.");
    }
  }
}
