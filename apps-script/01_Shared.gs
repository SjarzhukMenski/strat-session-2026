function getMainSs() { return SpreadsheetApp.getActiveSpreadsheet(); }
function getAuthSs() { return SpreadsheetApp.openById(CONFIG.AUTH_SHEET_ID); }

function ok(data) { return { ok: true, data }; }
function fail(msg, code) { return { ok: false, error: msg, code: code || 400 }; }

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function nowIso() { return new Date().toISOString(); }

function generateToken() {
  const bytes = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  return bytes.substring(0, 32);
}

function readRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h || '').trim());
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
}
