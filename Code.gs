// ================================================
// SISTEMA DE FICHAJE - Google Apps Script
// Pega este código en tu proyecto de Apps Script
// ================================================

const SHEET_NAME = 'Fichajes';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    
    if (data.action === 'entrada') {
      sheet.appendRow([
        data.id,
        data.fecha,
        data.hora,
        '',
        '',
        data.usuario || 'Sin nombre'
      ]);
      return ok({ message: 'Entrada registrada' });
    }
    
    if (data.action === 'salida') {
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][0] === data.id && rows[i][3] === '') {
          sheet.getRange(i + 1, 4).setValue(data.hora);
          sheet.getRange(i + 1, 5).setValue(data.horas);
          return ok({ message: 'Salida registrada' });
        }
      }
      return ok({ message: 'No se encontró entrada activa' });
    }
    
    if (data.action === 'ping') {
      return ok({ message: 'ok' });
    }

    return error('Acción no reconocida');
  } catch(err) {
    return error(err.toString());
  }
}

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1).map(row => ({
      id: row[0],
      fecha: row[1],
      entrada: row[2],
      salida: row[3],
      horas: row[4],
      usuario: row[5]
    }));
    return ok({ records });
  } catch(err) {
    return error(err.toString());
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Fecha', 'Entrada', 'Salida', 'Horas trabajadas', 'Usuario']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 130);
    sheet.setColumnWidth(6, 140);
  }
  return sheet;
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
