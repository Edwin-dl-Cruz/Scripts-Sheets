// Función principal que busca y copia nombres entre hojas de Google Sheets.
function buscarYCopiarNombres() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Obtener las hojas de trabajo relevantes.
  var hoja2 = ss.getSheetByName('Asistencia'); // Hoja de asistencia
  var hoja1 = ss.getSheetByName('API'); // Hoja de datos API
  var hoja3 = ss.getSheetByName('Resultado'); // Hoja de resultados
  var hojaNoEncontrado = ss.getSheetByName('noEncontrado'); // Hoja de no encontrados

  // Obtener los datos de las hojas de trabajo.
  var dataHoja2 = hoja2.getRange("A2:C" + hoja2.getLastRow()).getValues();
  var dataHoja1 = hoja1.getRange("A2:B" + hoja1.getLastRow()).getValues();

  var grupos = {}; // Un objeto para agrupar registros por grupo.

  // Bucle a través de los registros en la hoja de asistencia.
  for (var i = 0; i < dataHoja2.length; i++) {
    var nombreBuscado = dataHoja2[i][0];
    var fecha = dataHoja2[i][1];
    var hora = dataHoja2[i][2];

    var encontrado = false; // Variable para verificar si se encontró el registro en la hoja API.

    // Bucle a través de los registros en la hoja API.
    for (var j = 0; j < dataHoja1.length; j++) {
      var nombreHoja1 = dataHoja1[j][0];

      // Normalizar cadenas para una comparación sin distinción entre mayúsculas y minúsculas.
      var nombreNormalizado = normalizeString(nombreBuscado);
      var nombreHoja1Normalizado = normalizeString(nombreHoja1);

      if (nombreHoja1Normalizado === nombreNormalizado) {
        var grupo = dataHoja1[j][1];

        // Verificar si el grupo ya existe en el objeto de grupos.
        if (!grupos[grupo]) {
          grupos[grupo] = [];
        }

        // Verificar si el registro es duplicado en el grupo.
        var registroDuplicado = false;
        for (var k = 0; k < grupos[grupo].length; k++) {
          if (grupos[grupo][k][0] === nombreBuscado) {
            registroDuplicado = true;
            break;
          }
        }

        // Si no es duplicado, agregarlo al grupo.
        if (!registroDuplicado) {
          grupos[grupo].push([nombreBuscado, grupo, fecha, hora]);
        }

        encontrado = true; // Marcar como encontrado en la hoja API.
        break;
      }
    }

    if (!encontrado) {
      // Si el registro no se encontró en hoja API, agregarlo a la hoja "noEncontrado".
      hojaNoEncontrado.appendRow([nombreBuscado, fecha, hora]);
    }
  }

  // Copiar registros a hojas de trabajo separadas según los grupos.
  for (var grupo in grupos) {
    var hoja = ss.getSheetByName(grupo);

    if (!hoja) {
      hoja = ss.insertSheet(grupo); // Crear hoja si no existe.
    } else {
      var numRows = hoja.getLastRow() - 1;
      if (numRows > 0) {
        hoja.getRange(2, 1, numRows, hoja.getLastColumn()).clear(); // Limpiar registros existentes.
      }
    }

    var registrosGrupo = grupos[grupo];

    if (registrosGrupo.length > 0) {
      hoja.getRange(2, 1, registrosGrupo.length, 4).setValues(registrosGrupo); // Copiar registros al grupo.
    }
  }
}

// Función para normalizar una cadena para comparación sin distinción entre mayúsculas y minúsculas.
function normalizeString(str) {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
}
