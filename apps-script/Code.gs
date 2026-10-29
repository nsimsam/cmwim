// Lets the board at noorsimsam.com/cmwim/ read and edit the Tasks tab of this sheet.
// Setup: Project Settings → Script properties → add PASSCODE. Then Deploy → New deployment →
// Web app, Execute as: Me, Who has access: Anyone.

const TAB = 'Tasks';
const FIELDS = { initiative: 'Initiative', title: 'Task', owner: 'Owner', due: 'Due', status: 'Status', notes: 'Notes' };

function doGet() {
  return json_({ tasks: read_(), inits: inits_() });
}

function doPost(e) {
  const req = JSON.parse(e.postData.contents);
  if (req.passcode !== PropertiesService.getScriptProperties().getProperty('PASSCODE')) {
    return json_({ error: 'passcode' });
  }
  if (req.action === 'check') return json_({ ok: true });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (req.action === 'initUpdate') {
      setInit_(String(req.name || '').trim(), String(req.description || ''));
      SpreadsheetApp.flush();
      return json_({ ok: true, tasks: read_(), inits: inits_() });
    }
    const sh = sheet_(), c = cols_(sh), task = req.task || {};
    if (req.action === 'add') {
      const row = new Array(sh.getLastColumn()).fill('');
      for (const [k, h] of Object.entries(FIELDS)) if (k in task) row[c[h]] = task[k];
      row[c.ID] = newId_();
      sh.appendRow(row);
    } else {
      const n = sh.getLastRow() - 1;
      const ids = n > 0 ? sh.getRange(2, c.ID + 1, n, 1).getValues().map(r => String(r[0])) : [];
      const i = ids.indexOf(String(req.id));
      if (i < 0) return json_({ error: 'missing', tasks: read_(), inits: inits_() });
      if (req.action === 'delete') {
        sh.deleteRow(i + 2);
      } else if (req.action === 'update') {
        for (const [k, h] of Object.entries(FIELDS)) if (k in task) sh.getRange(i + 2, c[h] + 1).setValue(task[k]);
      }
    }
    if (task.initiative) addToList_(String(task.initiative).trim());
    SpreadsheetApp.flush();
    return json_({ ok: true, tasks: read_(), inits: inits_() });
  } finally {
    lock.releaseLock();
  }
}

// Adds an initiative typed on the board to the Lists tab, so the Sheet's Initiative dropdown offers it too.
function addToList_(name) {
  const ls = SpreadsheetApp.getActive().getSheetByName('Lists');
  if (!ls || !name) return;
  const head = ls.getRange(1, 1, 1, ls.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = head.indexOf('Initiative') + 1;
  if (!col) return;
  const vals = ls.getRange(2, col, Math.max(ls.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim()).filter(String);
  if (vals.indexOf(name) >= 0) return;
  ls.getRange(vals.length + 2, col).setValue(name);
  // Point the dropdown at the whole Lists column so later additions show up as well.
  try {
    const sh = sheet_(), c = cols_(sh);
    const rule = SpreadsheetApp.newDataValidation().requireValueInRange(ls.getRange(2, col, ls.getMaxRows() - 1, 1), true).setAllowInvalid(true).build();
    sh.getRange(2, c.Initiative + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  } catch (e) {
    // A Sheets table manages its own dropdowns; the name is still on the Lists tab.
  }
}

const INIT_TAB = 'Initiatives';

// Initiative descriptions edited on the board (or typed on the Initiatives tab).
// A planned initiative without a row keeps the board's built-in text.
function inits_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(INIT_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), description: String(r[1]) }));
}

function setInit_(name, description) {
  if (!name) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(INIT_TAB);
  if (!sh) {
    sh = ss.insertSheet(INIT_TAB);
    sh.getRange(1, 1, 1, 2).setValues([['Initiative', 'Description']]).setFontWeight('bold');
    sh.getRange('A:B').setNumberFormat('@');
  }
  const n = sh.getLastRow() - 1;
  const names = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
  const i = names.indexOf(name.toLowerCase());
  if (i >= 0) sh.getRange(i + 2, 2).setValue(description);
  else sh.appendRow([name, description]);
}

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(TAB);
}

// Maps header names to 0-based column indexes, adding an ID column on first run.
function cols_(sh) {
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (head.indexOf('ID') < 0) {
    sh.getRange(1, head.length + 1).setValue('ID');
    head.push('ID');
  }
  const c = {};
  head.forEach((h, i) => (c[h] = i));
  return c;
}

function newId_() {
  return Utilities.getUuid().slice(0, 8);
}

// Returns every task row, giving rows without an ID one (rows added by hand in the sheet).
function read_() {
  const sh = sheet_(), c = cols_(sh), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  const out = [];
  vals.forEach((r, i) => {
    if (!String(r[c.Task]).trim()) return;
    if (!r[c.ID]) {
      r[c.ID] = newId_();
      sh.getRange(i + 2, c.ID + 1).setValue(r[c.ID]);
    }
    const due = r[c.Due] instanceof Date ? Utilities.formatDate(r[c.Due], tz, 'yyyy-MM-dd') : String(r[c.Due] || '');
    out.push({
      id: String(r[c.ID]),
      initiative: String(r[c.Initiative]),
      title: String(r[c.Task]),
      owner: String(r[c.Owner]),
      due: due,
      status: String(r[c.Status]),
      notes: String(r[c.Notes]),
    });
  });
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
