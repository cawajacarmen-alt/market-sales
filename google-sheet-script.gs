// Market Sales — Google Sheet connector.
// Paste this into Extensions > Apps Script of your Google Sheet, then Deploy > New deployment >
// Web app, Execute as: Me, Who has access: Anyone. Copy the Web app URL into the app (Branch & items > Google Sheet).
//
// Tabs it keeps:  Daily totals, Sales, Expenses (for you to read)
//                 Data (do not edit) (what the app reads back)

const DATA = 'Data (do not edit)';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';
  if (action === 'ping') return out_({ ok: true });
  const rows = sheet_(DATA, ['Key', 'Saved data', 'Updated']).getDataRange().getValues().slice(1);
  const days = {};
  let setup = null, carry = null;
  rows.forEach(function (r) {
    const k = String(r[0]);
    let v;
    try { v = JSON.parse(r[1]); } catch (x) { return; }
    if (k.indexOf('day:') === 0) days[k.slice(4)] = v;
    else if (k === 'setup') setup = v;
    else if (k === 'carry') carry = v;
  });
  return out_({ ok: true, days: days, setup: setup, carry: carry });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const b = JSON.parse(e.postData.contents);
    const store = sheet_(DATA, ['Key', 'Saved data', 'Updated']);
    const vals = store.getDataRange().getValues();
    const at = {};
    vals.forEach(function (r, i) { if (i) at[String(r[0])] = i + 1; });
    const put = function (k, obj) {
      const row = [k, JSON.stringify(obj), new Date()];
      if (at[k]) store.getRange(at[k], 1, 1, 3).setValues([row]);
      else { store.appendRow(row); at[k] = store.getLastRow(); }
    };
    if (b.setup) put('setup', b.setup);
    if (b.carry) put('carry', b.carry);
    (b.days || []).forEach(function (d) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return;
      put('day:' + d.date, d.day);
      readable_(d);
    });
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Rewrites the human-readable rows for one day.
function readable_(d) {
  const s = d.sum || {};
  const t = s.tot || ['', '', '', '', ''];
  replace_(sheet_('Daily totals', ['Date', 'Sales', 'Expenses', 'Net (sales - expenses)', 'Cash at closing', 'Short (-) / over (+)']),
    d.date, [[d.date].concat(t)]);
  replace_(sheet_('Sales', ['Date', 'Item', 'Price', 'Beginning', 'Stock in', 'Stock out', 'Ending', 'Sold', 'Sales']),
    d.date, (s.items || []).map(function (r) { return [d.date].concat(r); }));
  replace_(sheet_('Expenses', ['Date', 'Category', 'Amount', 'Note']),
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
