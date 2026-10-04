// Branch Sales — Google Sheet connector (version 2: all branches in one Sheet).
// Paste this into Extensions > Apps Script of your Google Sheet, then Deploy > New deployment >
// Web app, Execute as: Me, Who has access: Anyone. Copy the Web app URL into the app (Branch & items > Google Sheet).
//
// Tabs it keeps:  Market: Daily totals, Sales, Expenses (for you to read)
//                 Other branches: "Mabuhay Daily totals", "Mabuhay Sales", "Mabuhay Expenses", and so on
//                 Data (do not edit) (what the apps read back)
//
// To update an existing connection: paste this over the old script, Save, then
// Deploy > Manage deployments > pencil icon > Version: New version > Deploy. The Web app URL stays the same.

const DATA = 'Data (do not edit)';
const VERSION = 2;
const NAMES = { market: 'Market', mabuhay: 'Mabuhay', mercedes: 'Mercedes', lugay: 'Lugay', main: 'Main Branch' };

// Market keeps the original keys and tab names, so its earlier entries stay where they are.
function branch_(b) {
  const k = String(b || 'market').toLowerCase().replace(/[^a-z0-9]/g, '') || 'market';
  const name = NAMES[k] || (k.charAt(0).toUpperCase() + k.slice(1));
  return { key: k, pre: k === 'market' ? '' : k + '|', tab: k === 'market' ? '' : name + ' ' };
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';
  if (action === 'ping') return out_({ ok: true, v: VERSION });
  const br = branch_(e.parameter.b);
  const rows = sheet_(DATA, ['Key', 'Saved data', 'Updated']).getDataRange().getValues().slice(1);
  const days = {};
  let setup = null, carry = null;
  rows.forEach(function (r) {
    let k = String(r[0]);
    if (br.pre) { if (k.indexOf(br.pre) !== 0) return; k = k.slice(br.pre.length); }
    else if (k.indexOf('|') >= 0) return;
    let v;
    try { v = JSON.parse(r[1]); } catch (x) { return; }
    if (k.indexOf('day:') === 0) days[k.slice(4)] = v;
    else if (k === 'setup') setup = v;
    else if (k === 'carry') carry = v;
  });
  return out_({ ok: true, v: VERSION, days: days, setup: setup, carry: carry });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const b = JSON.parse(e.postData.contents);
    const br = branch_(b.b);
    const store = sheet_(DATA, ['Key', 'Saved data', 'Updated']);
    const vals = store.getDataRange().getValues();
    const at = {};
    vals.forEach(function (r, i) { if (i) at[String(r[0])] = i + 1; });
    const put = function (key, obj) {
      const k = br.pre + key;
      const row = [k, JSON.stringify(obj), new Date()];
      if (at[k]) store.getRange(at[k], 1, 1, 3).setValues([row]);
      else { store.appendRow(row); at[k] = store.getLastRow(); }
    };
    if (b.setup) put('setup', b.setup);
    if (b.carry) put('carry', b.carry);
    (b.days || []).forEach(function (d) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return;
      put('day:' + d.date, d.day);
      readable_(d, br.tab);
    });
    return out_({ ok: true, v: VERSION });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Rewrites the human-readable rows for one day.
function readable_(d, tab) {
  const s = d.sum || {};
  const t = s.tot || ['', '', '', '', ''];
  replace_(sheet_(tab + 'Daily totals', ['Date', 'Sales', 'Expenses', 'Net (sales - expenses)', 'Cash at closing', 'Short (-) / over (+)']),
    d.date, [[d.date].concat(t)]);
  replace_(sheet_(tab + 'Sales', ['Date', 'Item', 'Price', 'Beginning', 'Stock in', 'Stock out', 'Ending', 'Sold', 'Sales']),
    d.date, (s.items || []).map(function (r) { return [d.date].concat(r); }));
  replace_(sheet_(tab + 'Expenses', ['Date', 'Category', 'Amount', 'Note']),
    d.date, (s.exp || []).map(function (r) { return [d.date].concat(r); }));
}

function replace_(sh, date, rows) {
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (day_(v[i][0]) === date) sh.deleteRow(i + 1);
  if (!rows.length) return;
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(r, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort({ column: 1, ascending: true });
}

function day_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
