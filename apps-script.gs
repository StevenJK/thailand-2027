// Thailand 2027 survey + ideas backend.
// Lives in the "Thailand 2027 – Answers" sheet: Extensions > Apps Script.
// After pasting a new version: Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
// The web app URL stays the same.

const COLS = ['id', 'updated', 'name', 'status', 'count', 'from', 'tier', 'bangkok', 'flex', 'activities', 'notes', 'run'];
const IDEA_COLS = ['id', 'created', 'name', 'idea', 'link', 'keen', 'status', 'hide'];

function answers_() {
  const sh = SpreadsheetApp.getActive().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(COLS);
  return sh;
}

function tab_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); }
  return sh;
}

// Stop typed text being treated as a formula.
function clean_(v, max) {
  const s = String(v == null ? '' : v).slice(0, max || 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function isHidden_(v) {
  return v === true || /^(true|yes|x|hide)$/i.test(String(v).trim());
}

function rowIndex_(sh, id) {
  const n = Math.max(sh.getLastRow() - 1, 1);
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  return ids.indexOf(String(id));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);

    if (d.type === 'idea') {
      const text = clean_(d.idea, 600);
      if (!text.trim()) return out_({ ok: false });
      const link = /^https?:\/\//i.test(String(d.link || '')) ? clean_(d.link, 400) : '';
      const id = 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      tab_('Ideas', IDEA_COLS).appendRow([id, new Date(), clean_(d.name, 60), text, link, 0, 'Idea', '']);
      return out_({ ok: true, id });
    }

    if (d.type === 'keen') {
      const sh = tab_('Ideas', IDEA_COLS);
      const i = rowIndex_(sh, d.ideaId);
      if (i >= 0) {
        const cell = sh.getRange(i + 2, 6);
        cell.setValue(Math.max(0, (Number(cell.getValue()) || 0) + (d.on ? 1 : -1)));
      }
      return out_({ ok: true });
    }

    // Default: a survey answer, updated in place when the same browser saves again.
    const row = COLS.map(k =>
      k === 'updated' ? new Date()
      : k === 'activities' ? clean_((d.activities || []).join(', '))
      : clean_(d[k]));
    const sh = answers_();
    const i = d.id ? rowIndex_(sh, clean_(d.id)) : -1;
    if (i >= 0) sh.getRange(i + 2, 1, 1, COLS.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Public data only: first names, yes/maybe/no, headcount, origin and race choice; visible ideas; updates.
function doGet() {
  const rows = answers_().getDataRange().getValues().slice(1)
    .filter(r => r[2])
    .map(r => ({ name: String(r[2]).trim().split(/\s+/)[0], status: r[3], count: r[4], from: r[5], run: r[11] || '' }));
  const ideas = tab_('Ideas', IDEA_COLS).getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[3] && !isHidden_(r[7]))
    .map(r => ({ id: String(r[0]), created: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      name: String(r[2]).trim().split(/\s+/)[0], idea: String(r[3]), link: String(r[4] || ''),
      keen: Number(r[5]) || 0, status: String(r[6] || 'Idea') }));
  const updates = tab_('Updates', ['date', 'update', 'hide']).getDataRange().getValues().slice(1)
    .filter(r => r[1] && !isHidden_(r[2]))
    .map(r => ({ date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), text: String(r[1]) }));
  return out_({ rows, ideas, updates });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
