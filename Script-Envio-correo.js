function enviarPDFporEmail() {
  // Obtiene la fecha actual y la formatea como "dd-MM-yyyy"
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");

  // Define el nombre de la hoja como la fecha formateada
  var nombreHoja = fecha;

  // Genera un PDF del archivo sheet
  var archivoPDF = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getAs('application/pdf');

  // Lista de destinatarios separados por comas
  var destinatarios = 'correo@escuela.edu.mx, direccionsecundaria@escuela.edu.mx ';

  // Suma los totales de las hojas con funcion CONTAR
  var totalAsistencias = 0;

  // Nombres las hojas con las fórmulas CONTAR
  var hojasAsistencias = ['Hoja 1', 'Hoja 2', 'Hoja 3', 'Hoja 4', 'Hoja 5', 'Hoja 6', 'Hoja 7'];

  // Itera a través de las hojas y suma los totales
  for (var i = 0; i < hojasAsistencias.length; i++) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojasAsistencias[i]);
    if (hoja) {
      // Obtiene el valor de la celda D30 (En donde se encuentra la funcion CONTAR)
      var total = hoja.getRange('D30').getValue(); 
      totalAsistencias += total;
    }
  }

  // Asunto del correo electrónico
  var asunto = 'Asistencias de Secundaria del ' + nombreHoja;

  // Cuerpo del correo electrónico
  var cuerpo = 'Se adjunta el PDF de las asistencias de Secundaria del dia ' + nombreHoja + '. Total de alumnos: ' + totalAsistencias;

  // Envía el correo electrónico con el archivo PDF adjunto a múltiples destinatarios
  MailApp.sendEmail(destinatarios, asunto, cuerpo, { attachments: [archivoPDF] });
}

// Ejecuta la función para generar el PDF y enviarlo por correo
function generarYEnviarPDF() {
  enviarPDFporEmail();
}
