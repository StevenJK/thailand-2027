// Thailand 2027 survey backend.
// Paste into the "Thailand 2027 – Answers" sheet: Extensions > Apps Script.
// Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.

const COLS = ['id', 'updated', 'name', 'status', 'count', 'from', 'tier', 'bangkok', 'flex', 'activities', 'notes'];

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(COLS);
  return sh;
}

// Stop typed text being treated as a formula.
function clean_(v) {
  const s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const row = COLS.map(k =>
      k === 'updated' ? new Date()
      : k === 'activities' ? clean_((d.activities || []).join(', '))
      : clean_(d[k]));
    const sh = sheet_();
    const n = Math.max(sh.getLastRow() - 1, 1);
    const ids = sh.getRange(2, 1, n, 1).getValues().map(r => r[0]);
    const i = d.id ? ids.indexOf(clean_(d.id)) : -1;
    if (i >= 0) sh.getRange(i + 2, 1, 1, COLS.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// The public board: first names, yes/maybe/no, headcount and origin only.
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1)
    .filter(r => r[2])
    .map(r => ({ name: String(r[2]).trim().split(/\s+/)[0], status: r[3], count: r[4], from: r[5] }));
  return out_({ rows });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
