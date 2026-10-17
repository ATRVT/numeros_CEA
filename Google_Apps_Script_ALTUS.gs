/**
 * ============================================================================
 * GOOGLE APPS SCRIPT - LIBRO CONTABLE ALTUS 26 AUDITADO (100% AUTÓNOMO)
 * ============================================================================
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Sube el archivo "Libro_Contable_ALTUS_26_Auditado.xlsx" a tu Google Drive.
 * 2. Ábrelo y ve a Archivo -> "Guardar como Hojas de cálculo de Google".
 * 3. Ve al menú: Extensiones -> Apps Script.
 * 4. Borra el código que haya y pega todo este archivo.
 * 5. Haz clic en "Implementar" -> "Nueva implementación" -> Tipo: "Aplicación web".
 *    - Ejecutar como: "Yo"
 *    - Quién tiene acceso: "Cualquier persona"
 * 6. Copia la URL generada y pégala en el botón "Sincronizar Sheets" de tu App ALTUS.
 * ============================================================================
 */

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (payload.action === 'appendEntry' || payload.entry) {
      var entry = payload.entry || payload;
      var diario = ss.getSheetByName('Diario');
      if (!diario) {
        return jsonResponse({ status: 'error', message: 'No se encontró la pestaña Diario' });
      }
      
      // Buscar la primera fila vacía en la columna G (Monto Bruto) a partir de la fila 2
      var lastRow = Math.max(diario.getLastRow(), 2);
      var colG = diario.getRange(2, 7, lastRow, 1).getValues();
      var targetRow = lastRow + 1;
      for (var i = 0; i < colG.length; i++) {
        if (colG[i][0] === '' || colG[i][0] === null) {
          targetRow = i + 2;
          break;
        }
      }
      
      var r = targetRow;
      var prev = r - 1;
      
      // Escribir valores y fórmulas vivas (Columnas A..O de Diario)
      diario.getRange(r, 1).setFormula('=IF(G' + r + '="","","ALT-26-"&TEXT(ROW()-1,"000"))'); // A: ID Asiento
      diario.getRange(r, 2).setValue(entry.fecha || '');                                       // B: Fecha
      diario.getRange(r, 3).setValue((entry.mes || '').toLowerCase());                         // C: Mes
      diario.getRange(r, 4).setValue(entry.descripcion || '');                                 // D: Descripción
      diario.getRange(r, 5).setValue(entry.cuentaFull || entry.cuenta || '');                  // E: Cuenta Contable SAT
      
      // F: Tipo Movimiento (Fórmula conectada a Config)
      diario.getRange(r, 6).setFormula(
        '=IF(E' + r + '="","",IFERROR(VLOOKUP(E' + r + ',Config!$F$2:$G$60,2,FALSE),"Gasto"))'
      );
      
      // G: Monto Bruto (Q)
      diario.getRange(r, 7).setValue(Number(entry.montoBruto || 0));
      
      // H: IVA (Q) (Fórmula conectada a Config)
      diario.getRange(r, 8).setFormula(
        '=IF(G' + r + '="","",IF(IFERROR(VLOOKUP(E' + r + ',Config!$F$2:$H$60,3,FALSE),"Exento")="Afecto",ROUND(G' + r + '-(G' + r + '/1.12),2),0))'
      );
      
      // I: Monto Neto (Q)
      diario.getRange(r, 9).setFormula('=IF(G' + r + '="","",ROUND(G' + r + '-H' + r + ',2))');
      
      // J: Cuenta Bancaria
      diario.getRange(r, 10).setValue(entry.banco || 'BI- Monetaria');
      
      // K: Saldo BI-Monetaria (Q)
      diario.getRange(r, 11).setFormula(
        '=IF(G' + r + '="","",ROUND(K' + prev + '+IF(J' + r + '="BI-Ahorro",0,IF(F' + r + '="Gasto",-G' + r + ',G' + r + ')),2))'
      );
      
      // L: Saldo BI-Ahorro (Q)
      diario.getRange(r, 12).setFormula(
        '=IF(G' + r + '="","",ROUND(L' + prev + '+IF(LEFT(E' + r + ',4)="2102",G' + r + ',IF(J' + r + '="BI-Ahorro",-G' + r + ',0)),2))'
      );
      
      // M, N, O: Metadatos de Auditoría
      diario.getRange(r, 13).setValue(entry.documento || entry.numDoc || 'S/N');
      diario.getRange(r, 14).setValue(entry.centroCosto || '00-Administración');
      diario.getRange(r, 15).setValue(entry.notaAuditoria || ('Registrado desde App ALTUS (' + new Date().toISOString().slice(0, 16).replace('T', ' ') + ')'));
      
      return jsonResponse({
        status: 'ok',
        row: r,
        message: 'Asiento registrado en fila ' + r + '. Colegiaturas, Conciliación y Panel V2 se actualizan solos por fórmula.'
      });
    }
    
    return jsonResponse({ status: 'error', message: 'Acción no reconocida' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var diario = ss.getSheetByName('Diario');
  var lastRow = diario ? diario.getLastRow() : 1;
  return jsonResponse({
    status: 'ok',
    workbook: ss.getName(),
    diarioRows: lastRow - 1
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
