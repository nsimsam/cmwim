// Lets the board at noorsimsam.com/cmwim/ read and edit the Tasks tab of this sheet.
// Setup: Project Settings → Script properties → add PASSCODE. Then Deploy → New deployment →
// Web app, Execute as: Me, Who has access: Anyone.

const TAB = 'Tasks';
const FIELDS = { initiative: 'Initiative', title: 'Task', owner: 'Owner', due: 'Due', status: 'Status', notes: 'Notes' };
// Initiative descriptions edited on the board: one row per initiative, ideas one per line.
const INIT_TAB = 'Initiatives';
const INIT_HEAD = ['Initiative', 'Stream', 'When', 'Ideas'];

function doGet() {
  return json_({ tasks: read_(), initiatives: readInits_() });
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
    if (req.action === 'initiative') {
      saveInit_(req.initiative || {});
      SpreadsheetApp.flush();
      return json_({ ok: true, tasks: read_(), initiatives: readInits_() });
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
      if (i < 0) return json_({ error: 'missing', tasks: read_() });
      if (req.action === 'delete') {
        sh.deleteRow(i + 2);
      } else if (req.action === 'update') {
        for (const [k, h] of Object.entries(FIELDS)) if (k in task) sh.getRange(i + 2, c[h] + 1).setValue(task[k]);
      }
    }
    if (task.initiative) addToList_(String(task.initiative).trim());
    SpreadsheetApp.flush();
    return json_({ ok: true, tasks: read_(), initiatives: readInits_() });
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

function readInits_() {
  const s = SpreadsheetApp.getActive().getSheetByName(INIT_TAB);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, INIT_HEAD.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), stream: String(r[1]).trim(), when: String(r[2]), ideas: String(r[3]) }));
}

// Adds or replaces an initiative's row on the Initiatives tab, creating the tab on first use.
function saveInit_(it) {
  const name = String(it.name || '').trim();
  if (!name) return;
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(INIT_TAB);
  if (!s) {
    s = ss.insertSheet(INIT_TAB);
    s.getRange(1, 1, 1, INIT_HEAD.length).setValues([INIT_HEAD]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  const n = Math.max(s.getLastRow() - 1, 0);
  const names = n ? s.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
  let row = names.indexOf(name.toLowerCase()) + 2;
  if (row < 2) row = n + 2;
  // Plain text, so a line starting with = or + is not read as a formula.
  s.getRange(row, 1, 1, INIT_HEAD.length).setNumberFormat('@')
    .setValues([[name, String(it.stream || ''), String(it.when || ''), String(it.ideas || '')]]);
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
