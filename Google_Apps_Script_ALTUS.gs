/**
 * ============================================================================
 * PUENTE DE SINCRONIZACIÓN EN TIEMPO REAL - CENTRO EDUCATIVO ALTUS
 * Google Sheets <-> App Contable iPhone & Mac (PWA)
 * ============================================================================
 * INSTRUCCIONES DE INSTALACIÓN (1 MINUTO):
 * 1. Sube "Libro_Contable_ALTUS_26_Auditado.xlsx" a tu Google Drive.
 * 2. Ábrelo con Google Sheets y ve al menú superior: Archivo -> Guardar como Hojas de cálculo de Google.
 * 3. En el menú superior ve a: Extensiones -> Apps Script.
 * 4. Borra el código de ejemplo, pega todo este archivo y presiona el ícono de Guardar (💾).
 * 5. Haz clic arriba a la derecha en "Implementar" -> "Nueva implementación".
 * 6. En Tipo selecciona "Aplicación web":
 *    - Ejecutar como: "Yo"
 *    - Quién tiene acceso: "Cualquier persona"
 * 7. Haz clic en "Implementar", copia la URL generada y pégala en el botón "Conectar Sheets" de tu App ALTUS.
 * ============================================================================
 */

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'status';

  if (action === 'sync') {
    const diarioSheet = ss.getSheetByName('Diario');
    const colegiaturasSheet = ss.getSheetByName('Colegiaturas');
    return jsonResponse({
      ok: true,
      spreadsheetName: ss.getName(),
      updatedAt: new Date().toISOString(),
      diarioRows: diarioSheet ? diarioSheet.getLastRow() - 1 : 0,
      colegiaturasRows: colegiaturasSheet ? colegiaturasSheet.getLastRow() - 1 : 0
    });
  }

  return jsonResponse({
    ok: true,
    status: 'ALTUS Contabilidad Bridge Active',
    spreadsheetName: ss.getName(),
    timestamp: new Date().toISOString()
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = payload.action || 'appendEntry';

    if (action === 'appendEntry') {
      const entry = payload.entry;
      const diario = ss.getSheetByName('Diario');
      if (!diario) throw new Error('No se encontró la pestaña Diario');

      diario.appendRow([
        entry.id,
        entry.fecha,
        entry.mes,
        entry.descripcion,
        entry.cuenta,
        entry.tipo,
        Number(entry.montoBruto),
        Number(entry.iva || 0),
        Number(entry.montoNeto),
        entry.banco,
        Number(entry.saldoMonetaria),
        Number(entry.saldoAhorro),
        entry.documento || '',
        entry.centroCosto || '00-Administración',
        entry.notaAuditoria || 'Registrado desde App ALTUS iPhone/Mac'
      ]);

      // Si es un pago de colegiatura, actualizar matriz en pestaña Colegiaturas
      if (entry.studentUpdate && entry.studentUpdate.idFamilia && entry.studentUpdate.mes) {
        updateStudentTuitionCell(ss, entry.studentUpdate.idFamilia, entry.studentUpdate.mes, Number(entry.montoBruto));
      }

      return jsonResponse({ ok: true, id: entry.id, message: 'Asiento registrado en Google Sheets exitosamente' });
    }

    if (action === 'batchEntries') {
      const entries = payload.entries || [];
      const diario = ss.getSheetByName('Diario');
      entries.forEach(function(entry) {
        diario.appendRow([
          entry.id,
          entry.fecha,
          entry.mes,
          entry.descripcion,
          entry.cuenta,
          entry.tipo,
          Number(entry.montoBruto),
          Number(entry.iva || 0),
          Number(entry.montoNeto),
          entry.banco,
          Number(entry.saldoMonetaria),
          Number(entry.saldoAhorro),
          entry.documento || '',
          entry.centroCosto || '00-Administración',
          entry.notaAuditoria || 'Lote Nómina Día 27 App ALTUS'
        ]);
      });
      return jsonResponse({ ok: true, count: entries.length, message: 'Lote registrado en Google Sheets' });
    }

    return jsonResponse({ ok: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function updateStudentTuitionCell(ss, idFamilia, mesNombre, monto) {
  const sheet = ss.getSheetByName('Colegiaturas');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  const colIndex = headers.indexOf(mesNombre.toLowerCase().trim());
  if (colIndex === -1) return;

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === String(idFamilia).trim()) {
      const currentVal = Number(data[r][colIndex] || 0);
      sheet.getRange(r + 1, colIndex + 1).setValue(currentVal + monto);
      break;
    }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
