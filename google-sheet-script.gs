// Bakery websites — Google Sheet connector (version 6: branch sales, deliveries, online orders, production and peddler sales in one Sheet, with staff logins).
// Paste this into Extensions > Apps Script of your Google Sheet, then Deploy > New deployment >
// Web app, Execute as: Me, Who has access: Anyone. Copy the Web app URL into the app (Branch & items > Google Sheet).
//
// Tabs it keeps:  Market: Daily totals, Sales, Expenses (for you to read)
//                 Other branches: "Mabuhay Daily totals", "Mabuhay Sales", "Mabuhay Expenses", and so on
//                 Deliveries, Online orders, Production (one row per bread line / order)
//                 Peddler sales (one row per peddler per day: taken, returned, sales, cash turned in)
//                 Data (do not edit), App data (do not edit) (what the apps read back)
//                 Staff logins: one row per person. Name, PIN, Sites (all, or e.g. market, deliveries), Manager (yes / no).
//                 Delete a row to remove someone's access.
//
// To update an existing connection: paste this over the old script, Save, then
// Deploy > Manage deployments > pencil icon > Version: New version > Deploy. The Web app URL stays the same.

const DATA = 'Data (do not edit)';
const VERSION = 6;
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
  if (action === 'login') return out_(login_(e.parameter));
  if (action === 'docs') return out_(appGet_(e.parameter.app, Number(e.parameter.since) || 0));
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
    if (b.action === 'docs') return out_(appPost_(b.app, b.docs || {}));
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

// ---------- Deliveries and Online orders websites ----------
// Each saved document is a record of pieces with change times; records from different phones are merged piece by piece.
const APPDATA = 'App data (do not edit)';
const CHUNK = 45000; // a cell holds 50,000 characters, so long records (receipt pictures) span several cells
const SPLITS = {
  dlv: [['^delivery/setup/days/', { drops: 'arr', ready: 'arr' }]],
  ord: [['^days/', { orders: 'map', collections: 'map' }]],
  prd: [['^bakery/log/days/', { orders: 'map' }]],
  pdl: [['^pdl/days/', { peddlers: 'map' }]]
};
// Shared by the website pages and the Google Sheet script: one saved record per document.
// A record keeps each piece of a document with the time it last changed, so two phones that
// change different orders or receipts on the same day both keep their change.
// rec = {u:{<unit>:{t:<ms>, v:<value>} | {t:<ms>, del:1}}}
// Unit keys: "f" = whole field f; "f.m.<key>" = one entry of map field f; "f.a.<id>" = one item (by id) of array field f.
function wsSplitOf(split, path) {
  for (var i = 0; i < split.length; i++) if (new RegExp(split[i][0]).test(path)) return split[i][1];
  return {};
}
function wsUnits(doc, sp) {
  var u = {};
  Object.keys(doc || {}).forEach(function (k) {
    var v = doc[k];
    if (sp[k] === 'map' && v && typeof v === 'object' && !Array.isArray(v)) Object.keys(v).forEach(function (s) { u[k + '.m.' + s] = v[s]; });
    else if (sp[k] === 'arr' && Array.isArray(v)) v.forEach(function (x) { if (x && x.id != null) u[k + '.a.' + x.id] = x; });
    else u[k] = v;
  });
  return u;
}
function wsAssemble(rec, sp) {
  var d = {}, any = false;
  Object.keys(sp).forEach(function (k) { d[k] = sp[k] === 'arr' ? [] : {}; });
  Object.keys((rec && rec.u) || {}).forEach(function (key) {
    var x = rec.u[key]; if (x.del) return; any = true;
    var m = /^(.*?)\.(m|a)\.(.*)$/.exec(key);
    if (m && sp[m[1]] === (m[2] === 'm' ? 'map' : 'arr')) { if (m[2] === 'm') d[m[1]][m[3]] = x.v; else d[m[1]].push(x.v); }
    else d[key] = x.v;
  });
  return any ? d : null;
}
function wsMerge(a, b) {
  var out = { u: {} }, au = (a && a.u) || {}, bu = (b && b.u) || {};
  Object.keys(au).concat(Object.keys(bu)).forEach(function (k) {
    var x = au[k], y = bu[k];
    out.u[k] = !x ? y : !y ? x : (y.t > x.t || (y.t === x.t && JSON.stringify(y) > JSON.stringify(x))) ? y : x;
  });
  return out;
}
// New record after a page writes `doc` (null = deleted) over `rec`.
function wsStamp(rec, doc, sp, now) {
  var old = (rec && rec.u) || {}, nu = doc ? wsUnits(doc, sp) : {}, out = { u: {} };
  Object.keys(old).forEach(function (k) { out.u[k] = old[k]; });
  Object.keys(nu).forEach(function (k) {
    var o = old[k];
    if (!o || o.del || JSON.stringify(o.v) !== JSON.stringify(nu[k])) out.u[k] = { t: Math.max(now, o ? o.t + 1 : 0), v: nu[k] };
  });
  Object.keys(old).forEach(function (k) { if (!old[k].del && !(k in nu)) out.u[k] = { t: Math.max(now, old[k].t + 1), del: 1 }; });
  return out;
}

function appSheet_() { return sheet_(APPDATA, ['Key', 'Updated', 'Saved data']); }
function appRead_(sh) {
  const v = sh.getDataRange().getValues(), out = {};
  for (let i = 1; i < v.length; i++) {
    const k = String(v[i][0]); if (!k) continue;
    let txt = ''; for (let c = 2; c < v[i].length; c++) txt += String(v[i][c] || '');
    out[k] = { row: i + 1, at: Number(v[i][1]) || 0, txt: txt };
  }
  return out;
}
function appGet_(app, since) {
  if (!SPLITS[app]) return { ok: false, error: 'unknown app' };
  const all = appRead_(appSheet_()), docs = {}, t = Date.now(), pre = app + '|';
  Object.keys(all).forEach(function (k) {
    if (k.indexOf(pre) !== 0 || all[k].at <= since) return;
    try { docs[k.slice(pre.length)] = JSON.parse(all[k].txt); } catch (x) {}
  });
  return { ok: true, v: VERSION, t: t - 1, docs: docs };
}
function appPost_(app, docs) {
  if (!SPLITS[app]) return { ok: false, error: 'unknown app' };
  const sh = appSheet_(), all = appRead_(sh), t = Date.now(), dates = {};
  Object.keys(docs).forEach(function (p) {
    const k = app + '|' + p, old = all[k];
    let rec = docs[p];
    if (old) { try { rec = wsMerge(JSON.parse(old.txt), rec); } catch (x) {} }
    const txt = JSON.stringify(rec), parts = [];
    for (let i = 0; i < txt.length; i += CHUNK) parts.push(txt.slice(i, i + CHUNK));
    const row = [k, t].concat(parts);
    if (old) {
      const width = Math.max(sh.getLastColumn(), row.length);
      while (row.length < width) row.push('');
      sh.getRange(old.row, 1, 1, row.length).setValues([row]);
    } else { sh.appendRow(row); }
    const sp = wsSplitOf(SPLITS[app], p), d = wsAssemble(rec, sp);
    const m = /\/(\d{4}-\d{2}-\d{2})$/.exec(p);
    if (m && (/^delivery\/setup\/days\//.test(p) || /^days\//.test(p) || /^bakery\/log\/days\//.test(p) || /^pdl\/days\//.test(p))) dates[m[1]] = d || {};
  });
  Object.keys(dates).forEach(function (date) {
    if (app === 'dlv') dlvReadable_(date, dates[date], all);
    if (app === 'ord') ordReadable_(date, dates[date], all);
    if (app === 'prd') prdReadable_(date, dates[date], all);
    if (app === 'pdl') pdlReadable_(date, dates[date], all);
  });
  return { ok: true, v: VERSION };
}
function setupOf_(all, k) { try { return wsAssemble(JSON.parse(all[k].txt), {}) || {}; } catch (x) { return {}; } }
function dlvReadable_(date, d, all) {
  const st = all['dlv|delivery/setup'] ? setupOf_(all, 'dlv|delivery/setup') : {};
  const bn = {}, inm = {};
  (st.branches || []).forEach(function (b) { bn[b.id] = b.name; });
  (st.items || []).forEach(function (i) { inm[i.id] = i.name; });
  const rows = [];
  (d.drops || []).slice().sort(function (a, b) { return (a.no || 0) - (b.no || 0); }).forEach(function (r) {
    Object.keys(r.items || {}).forEach(function (iid) {
      const q = r.items[iid] || {}, p = Number(q.price) || 0;
      rows.push([date, 'DR-' + String(r.no || 0).padStart(5, '0'), bn[r.bid] || r.bid, r.time || '', r.driver || '', r.by || '', inm[iid] || iid,
        Number(q.sent) || 0, Number(q.pulled) || 0, Number(q.damaged) || 0, p, (Number(q.sent) || 0) * p, r.void ? 'VOID' : '']);
    });
  });
  replace_(sheet_('Deliveries', ['Date', 'Receipt no.', 'Branch', 'Time', 'Driver', 'Issued by', 'Bread', 'Delivered', 'Pulled out', 'Damaged', 'Price', 'Amount', 'Void']), date, rows);
}
function ordReadable_(date, d, all) {
  const st = all['ord|app/setup'] ? setupOf_(all, 'ord|app/setup') : {}, rn = {};
  (st.riders || []).forEach(function (r) { rn[r.id] = r.name; });
  const PAY = { cod: 'Cash', gcash: 'GCash', maya: 'Maya' };
  const rows = Object.keys(d.orders || {}).map(function (id) { return d.orders[id]; }).sort(function (a, b) { return (a.no || 0) - (b.no || 0); }).map(function (o) {
    const items = (o.items || []).map(function (l) { return l.qty + ' x ' + l.name; }).join(', ');
    const tot = (o.items || []).reduce(function (s, l) { return s + (Number(l.qty) || 0) * (Number(l.price) || 0); }, 0) + (Number(o.fee) || 0);
    const p = o.pay || {};
    return [date, o.no || '', o.t ? Utilities.formatDate(new Date(o.t), Session.getScriptTimeZone(), 'h:mm a') : '', o.cust || '', o.phone || '', o.addr || '', items, tot,
      PAY[p.m] || '', p.paid ? 'Paid' : 'Not paid', p.ref || '', o.status || '', rn[o.rider] || '', o.remit ? 'Turned in' : '', o.cancel && o.cancel.reason ? o.cancel.reason : ''];
  });
  replace_(sheet_('Online orders', ['Date', 'Order no.', 'Time', 'Customer', 'Phone', 'Address', 'Breads', 'Total', 'Payment', 'Paid', 'Reference', 'Status', 'Rider', 'Rider cash', 'Cancel reason']), date, rows);
}
function prdReadable_(date, d, all) {
  const st = all['prd|bakery/log'] ? setupOf_(all, 'prd|bakery/log') : {}, pn = {}, bn = {};
  (st.products || []).forEach(function (p) { pn[p.id] = p.name; });
  (st.bakers || []).forEach(function (b) { bn[b.id] = b.name; });
  const rows = Object.keys(d.orders || {}).map(function (id) { return d.orders[id]; }).sort(function (a, b) { return (a.no || 0) - (b.no || 0); }).map(function (o) {
    const done = o.status === 'inspected';
    return [date, 'PO-' + date.replace(/-/g, '').slice(2) + '-' + (o.no || ''), pn[o.pid] || o.pid || '', Number(o.batches) || 0, bn[o.baker] || '', o.status || '',
      o.yield === undefined ? '' : Number(o.yield) || 0, done ? Number(o.passed) || 0 : '', done ? Number(o.damaged) || 0 : '', o.why || '', o.note || ''];
  });
  replace_(sheet_('Production', ['Date', 'Order no.', 'Bread', 'Batches', 'Baker', 'Status', 'Counted', 'Passed (released)', 'Damaged', 'Damage reason', 'Note']), date, rows);
}

function pdlReadable_(date, d, all) {
  const st = all['pdl|pdl/setup'] ? setupOf_(all, 'pdl|pdl/setup') : {}, pn = {}, it = {};
  (st.peddlers || []).forEach(function (p) { pn[p.id] = p.name; });
  (st.items || []).forEach(function (i) { it[i.id] = i; });
  const n = function (v) { v = parseFloat(v); return isFinite(v) ? v : 0; };
  const rows = Object.keys(d.peddlers || {}).map(function (pid) {
    const r = d.peddlers[pid], ids = {};
    (r.loads || []).forEach(function (l) { Object.keys(l.q || {}).forEach(function (k) { ids[k] = 1; }); });
    Object.keys(r.ret || {}).forEach(function (k) { ids[k] = 1; });
    let tk = 0, rt = 0, sold = 0, sales = 0;
    const bread = [];
    Object.keys(ids).forEach(function (k) {
      const t = (r.loads || []).reduce(function (a, l) { return a + n((l.q || {})[k]); }, 0), b = n((r.ret || {})[k]), s = Math.max(0, t - b);
      const p = r.price && r.price[k] != null ? n(r.price[k]) : n((it[k] || {}).price);
      tk += t; rt += b; sold += s; sales += s * p; if (t || b) bread.push(((it[k] || {}).name || k) + ' ' + s + '/' + t);
    });
    const comm = sales * n(r.comm) / 100, due = sales - comm, paid = r.cash === '' || r.cash == null ? '' : n(r.cash);
    return [date, pn[pid] || pid, (r.loads || []).length, tk, rt, sold, sales, comm, due, paid, paid === '' ? 'not settled' : paid - due, bread.join(', '), r.by || '', r.note || ''];
  });
  replace_(sheet_('Peddler sales', ['Date', 'Peddler', 'Loads', 'Taken', 'Returned', 'Sold', 'Sales', 'Commission', 'To turn in', 'Turned in', 'Short / over', 'Bread (sold/taken)', 'Recorded by', 'Note']), date, rows);
}

// ---------- Staff logins ----------
function login_(q) {
  const sh = sheet_('Staff logins', ['Name', 'PIN', 'Sites', 'Manager']);
  const rows = sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]).trim(); });
  if (!rows.length) return { ok: false, v: VERSION, setup: true };
  const name = String(q.name || '').trim().toLowerCase(), site = String(q.site || '').toLowerCase();
  const cache = CacheService.getScriptCache(), ck = 'fail:' + name, fails = Number(cache.get(ck)) || 0;
  if (fails >= 5) return { ok: false, v: VERSION, locked: true };
  const r = rows.filter(function (x) { return String(x[0]).trim().toLowerCase() === name && pinEq_(x[1], q.pin); })[0];
  if (!r) { cache.put(ck, String(fails + 1), 600); return { ok: false, v: VERSION }; }
  cache.remove(ck);
  const sites = String(r[2] || '').trim(), list = sites.toLowerCase().split(/[\s,;]+/);
  if (sites && list.indexOf('all') < 0 && list.indexOf(site) < 0) return { ok: false, v: VERSION, site: true };
  return { ok: true, v: VERSION, name: String(r[0]).trim(), sites: sites || 'all', manager: /^(y|yes|oo|true|1)$/i.test(String(r[3]).trim()) };
}
function pinEq_(a, b) {
  const x = String(a).trim(), y = String(b || '').trim();
  return !!y && (x === y || (/^\d+$/.test(y) && x !== '' && Number(x) === Number(y)));
}
