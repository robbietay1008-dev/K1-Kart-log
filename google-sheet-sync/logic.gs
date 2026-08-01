/* K1 Kart Log — sheet-side LOGIC (fetched by the shell installed in Apps Script).
   Edit this file, push to GitHub, and the sheet picks it up within ~10 minutes
   (or instantly after running refreshLogic in the Apps Script editor). */

/** K1 Kart Log receiver v3 — the sheet is the app's full storage.
 *  Writes: APP ALL DATES, APP LOG, APP PARTS USED, APP PHOTOS,
 *  kart tabs 1-53, appends to "parts used", hidden _APP DATA.
 *  Never touches inventory tabs' content or the template. */

var LOGIC_VER = 'v8.6';

var COUNT_TAB = 'APP COUNT SHEET';

var KART_TABS = (function(){ var a=[]; for (var i=1;i<=53;i++) a.push(String(i)); return a; })();

function handlePost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data || data.app !== 'k1kartlog') return txt('ignored');
    var lock = LockService.getScriptLock();
    lock.waitLock(40000);
    try {
      if (data.type === 'photo') { savePhoto(data); return txt('ok photo'); }
      if (data.type === 'count') { return txt(writeCountRows(SpreadsheetApp.getActiveSpreadsheet(), data)); }
      if (data.type === 'snapshot') {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        saveJson('snapshot', { savedAt: new Date().toISOString(), appBuild: data.appBuild || '',
                               karts: data.karts, shop: data.shop || [],
                               quicks: data.quicks || [], inv: data.inv || {},
                               tomb: data.tomb || {}, stamps: data.stamps || {},
                               invTouched: data.invTouched || {}, rc: data.rc || {},
                               invCounted: data.invCounted || {},
                               invCfg: data.invCfg || {}, parts: data.parts || [],
                               cfgTouched: data.cfgTouched || {},
                               partTomb: data.partTomb || {}, rekeys: data.rekeys || [] });
        var photoIndex = loadJson('photo_index', {});
        var errs = [];
        function step(name, fn){ try { fn(); SpreadsheetApp.flush(); } catch (err2) { errs.push(name + ': ' + err2); } }
        step('all dates', function(){ writeAllDates(ss, data); });
        step('log', function(){ writeLog(ss, data, photoIndex); });
        step('parts used', function(){ writePartsUsed(ss, data); });
        step('kart tabs', function(){ writeKartTabs(ss, data); });
        step('part config', function(){ mergeCfg(ss, data); });
        step('counts', function(){ scanCounts(ss, data); });
        step('inventory qty', function(){ writeInventoryQty(ss, data.inv, data.invCfg, data.invCounted || {}); });
        step('needed', function(){ writeNeeded(ss, data.inv); });
        step('order view', function(){ ensureOrderView(ss); });
        saveJson('lastSync', { at: new Date().toISOString(), build: data.appBuild || '', errs: errs });
        return txt(errs.length ? 'ok with errors: ' + errs.join(' | ') : 'ok');
      }
      return txt('ignored');
    } finally { lock.releaseLock(); }
  } catch (err) { return txt('error: ' + err); }
}

function handleGet(e) {
  var cb = (e && e.parameter && e.parameter.callback ? e.parameter.callback : 'callback').replace(/[^\w$.]/g, '');
  if (e && e.parameter && e.parameter.mode === 'snapshot') {
    var snap = loadJson('snapshot', null);
    var photoIndex = loadJson('photo_index', {});
    var payload = JSON.stringify({ ok: !!snap, appBuild: snap ? (snap.appBuild || '') : '',
                                   savedAt: snap ? (snap.savedAt || '') : '',
                                   karts: snap ? snap.karts : null,
                                   shop: snap ? (snap.shop || []) : [],
                                   quicks: snap ? (snap.quicks || []) : [],
                                   inv: snap ? (snap.inv || {}) : {},
                                   tomb: snap ? (snap.tomb || {}) : {},
                                   stamps: snap ? (snap.stamps || {}) : {},
                                   invTouched: snap ? (snap.invTouched || {}) : {},
                                   invCounted: snap ? (snap.invCounted || {}) : {},
                                   rc: snap ? (snap.rc || {}) : {},
                                   invCfg: snap ? (snap.invCfg || {}) : {},
                                   cfgTouched: snap ? (snap.cfgTouched || {}) : {},
                                   partTomb: snap ? (snap.partTomb || {}) : {},
                                   rekeys: snap ? (snap.rekeys || []) : [],
                                   photos: photoIndex });
    return ContentService.createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  if (e && e.parameter && e.parameter.mode === 'inv') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    /* Reconcile whatever was typed on the inventory tab against the last data
       the app sent us, so the answer already accounts for both sides. If
       another device is mid-merge, answer without config rather than guess. */
    var res2 = { names: null, del: null, at: 0 };
    var lk2 = LockService.getScriptLock();
    if (lk2.tryLock(15000)) {
      try {
        var snapI = loadJson('snapshot', null);
        res2 = mergeCfg(ss, snapI);
        /* fill the paper-count tab in from whatever the app has counted */
        try { scanCounts(ss, snapI); } catch (ec) {}
      }
      catch (em) { res2 = { names: invConfigFromSheet(ss), del: null, at: 0, err: String(em) }; }
      finally { lk2.releaseLock(); }
    }
    var payload2 = JSON.stringify({ ok: true, ver: LOGIC_VER,
                                    names: res2.names, del: res2.del, cfgAt: res2.at,
                                    cfgErr: res2.err || '',
                                    receipts: loadJson('receipts', []),
                                    have: Object.keys(loadJson('photo_index', {})),
                                    last: loadJson('lastSync', null) });
    return ContentService.createTextOutput(cb + '(' + payload2 + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  if (e && e.parameter && e.parameter.mode === 'redraw') {
    /* Rebuild tabs straight from the stored snapshot — no data round-trip,
       so nothing can be lost. ?only=log rebuilds a single tab quickly. */
    var snap3 = loadJson('snapshot', null);
    if (!snap3 || !snap3.karts) return txt('no snapshot stored');
    var ss3 = SpreadsheetApp.getActiveSpreadsheet();
    var pi3 = loadJson('photo_index', {});
    var only = String((e.parameter.only || '')).toLowerCase();
    var errs3 = [];
    var step3 = function (name, fn) {
      if (only && only !== name) return;
      try { fn(); SpreadsheetApp.flush(); } catch (err3) { errs3.push(name + ': ' + err3); }
    };
    step3('log', function () { writeLog(ss3, snap3, pi3); });
    step3('alldates', function () { writeAllDates(ss3, snap3); });
    step3('partsused', function () { writePartsUsed(ss3, snap3); });
    step3('karttabs', function () { writeKartTabs(ss3, snap3); });
    step3('needed', function () { writeNeeded(ss3, snap3.inv); });
    step3('orderview', function () { ensureOrderView(ss3); });
    /* record the outcome so the app's health check reflects reality —
       a stale failure would otherwise keep warning about a fixed problem */
    saveJson('lastSync', { at: new Date().toISOString(),
                           build: (snap3.appBuild || '') + '+redraw', errs: errs3 });
    return txt('redraw ' + (only || 'all') + ' [' + LOGIC_VER + ']: ' +
               (errs3.length ? 'ERRORS ' + errs3.join(' | ') : 'ok'));
  }
  if (e && e.parameter && e.parameter.mode === 'peek') {
    /* Read-only look at a tab: row count plus the first few rows. */
    var ssp = SpreadsheetApp.getActiveSpreadsheet();
    var tp = ssp.getSheetByName(e.parameter.tab || 'APP LOG');
    if (!tp) return txt('missing tab');
    var lr = tp.getLastRow(), lc = tp.getLastColumn();
    var want = Math.min(lr, parseInt(e.parameter.rows, 10) || 4);
    var body = '';
    if (want > 0 && lc > 0) {
      var vv = tp.getRange(1, 1, want, lc).getDisplayValues();
      for (var vi = 0; vi < vv.length; vi++) body += '\n' + vv[vi].join(' | ');
    }
    return txt(tp.getName() + ': rows=' + lr + ' cols=' + lc + body);
  }
  if (e && e.parameter && e.parameter.mode === 'diag') {
    /* Report exactly which sheet operation a tab refuses. Read-only except
       for the optional destructive steps, which must be asked for by name. */
    var ssd = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];
    var all = ssd.getSheets(), names = [];
    for (var n = 0; n < all.length; n++) names.push(all[n].getName());
    out.push('tabs(' + all.length + '): ' + names.join(','));
    var tname = e.parameter.tab || 'APP LOG';
    var t = ssd.getSheetByName(tname);
    if (!t) { out.push(tname + ': MISSING'); return txt(out.join('\n')); }
    out.push(tname + ': index=' + t.getIndex() + ' lastRow=' + t.getLastRow() +
             ' lastCol=' + t.getLastColumn() + ' maxRow=' + t.getMaxRows());
    var probe = function (label, fn) {
      try { var r = fn(); SpreadsheetApp.flush(); out.push(label + ': ok' + (r === undefined ? '' : ' -> ' + r)); }
      catch (ep) { out.push(label + ': FAIL ' + ep); }
    };
    probe('clearContents', function () { t.clearContents(); });
    probe('setNumberFormat D', function () { t.getRange('D:D').setNumberFormat('@'); });
    probe('setValues A1', function () { t.getRange(1, 1, 1, 2).setValues([['DATE', 'KART']]); });
    if (e.parameter.drop === 'yes') {
      probe('deleteSheet', function () { ssd.deleteSheet(t); });
      probe('recreate', function () { ssd.insertSheet(tname, 1); return 'made'; });
    }
    return txt(out.join('\n'));
  }
  if (e && e.parameter && e.parameter.mode === 'bust') {
    /* drop the cached copy of this file so the very next call re-fetches
       it from GitHub — no more waiting out the 10-minute window */
    try { CacheService.getScriptCache().remove('klogic'); } catch (eb) {}
    return txt('logic cache cleared (was ' + LOGIC_VER + ')');
  }
  return txt('kart log sync is running (' + LOGIC_VER + ')');
}

function txt(s) { return ContentService.createTextOutput(s); }

/* ================= ONE-TIME CLEANUP (run manually once) ================= */
function cleanupImpl() {
  /* Resumable watchdog/rebuild: summary tabs first, then kart tabs in
     time-budgeted chunks; a cursor lets successive runs continue. */
  var snap = loadJson('snapshot', null);
  if (!snap || !snap.karts) return;
  var interactive = false;
  try { SpreadsheetApp.getUi(); interactive = true; } catch (e) {}
  var cur = loadJson('wd_cursor', null);
  if (!interactive && !cur) {
    var last = loadJson('lastSync', null);
    var snapAt = new Date(snap.savedAt || 0).getTime();
    var lastAt = last ? new Date(last.at || 0).getTime() : 0;
    var lastHadErrors = last && last.errs && last.errs.length > 0;
    if (lastAt >= snapAt && !lastHadErrors) return; // everything current
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var deadline = Date.now() + 210 * 1000; // ~3.5 min budget per run
    var photoIndex = loadJson('photo_index', {});
    var errs = (cur && cur.errs) || [];
    function step(name, fn){ try { fn(); SpreadsheetApp.flush(); } catch (err2) { errs.push(name + ': ' + err2); } }
    var phase = cur ? cur.phase : 0;
    var kartPos = cur ? (cur.k || 0) : 0;
    if (phase === 0) {
      step('all dates', function(){ writeAllDates(ss, snap); });
      step('log', function(){ writeLog(ss, snap, photoIndex); });
      step('parts used', function(){ writePartsUsed(ss, snap); });
      step('inventory qty', function(){ writeInventoryQty(ss, snap.inv, snap.invCfg, snap.invCounted || {}); });
      step('needed', function(){ writeNeeded(ss, snap.inv); });
      step('order view', function(){ ensureOrderView(ss); });
      phase = 1; kartPos = 0;
    }
    var doneAll = false;
    if (phase === 1) {
      var next = writeKartTabsChunk(ss, snap, kartPos, deadline, errs);
      if (next < 0) doneAll = true; else kartPos = next;
    }
    if (doneAll) {
      saveJson('wd_cursor', null);
      saveJson('lastSync', { at: new Date().toISOString(),
                             build: (snap.appBuild || '') + '+watchdog', errs: errs });
    } else {
      saveJson('wd_cursor', { phase: 1, k: kartPos, errs: errs });
    }
    if (interactive) {
      try {
        SpreadsheetApp.getUi().alert(doneAll
          ? 'Tabs rebuilt from the latest synced data.'
          : 'Summary tabs done, kart tabs partly done — the rest finishes automatically in the background over the next few 10-minute cycles. No need to run this again.');
      } catch (e2) {}
    }
  } finally { lock.releaseLock(); }
}
function writeKartTabsChunk(ss, data, startIdx, deadline, errs) {
  var sigs = loadJson('kart_sigs', {});
  var ks = kartOrder(data.karts);
  var wrote = false;
  for (var i = startIdx; i < ks.length; i++) {
    if (Date.now() > deadline) { if (wrote) saveJson('kart_sigs', sigs); return i; }
    var k = ks[i];
    try {
      var kd = data.karts[k], st = kd.status || {};
      var sig = kartSig(JSON.stringify(kd));
      if (sigs[k] === sig) continue;
      var sh = ss.getSheetByName(k);
      if (!sh) sh = ss.insertSheet(k);
      var rows = [KART_HDR.slice()];
      rows.push(['', '', '', '', kd.knotes || '', st.chain || '', st.diff || '', st.brake || '',
                 st.bat1 || '', st.bat2 || '', st.bat3 || '', st.bat4 || '', st.weld || 0]);
      var es = (kd.entries || []).slice().sort(function (a, b) {
        return cmpKey(entryKey(a.date), entryKey(b.date));   /* oldest first */
      });
      for (var j = 0; j < es.length; j++) {
        var en = es[j];
        rows.push([en.date || '', en.action || '', en.parts || '', en.mechanic || '',
                   en.notes || '', '', '', '', '', '', '', '', '']);
      }
      sh.clearContents();
      sh.getRange('C:C').setNumberFormat('@');
      sh.getRange('I:L').setNumberFormat('@');
      sh.getRange(1, 1, rows.length, KART_HDR.length).setValues(rows);
      sh.getRange(1, 1, 1, KART_HDR.length).setFontWeight('bold');
      sigs[k] = sig;
      wrote = true;
    } catch (e3) { errs.push('kart ' + k + ': ' + e3); }
  }
  if (wrote) saveJson('kart_sigs', sigs);
  return -1;
}

/* ================= hidden-tab JSON storage ================= */
function dataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_APP DATA');
  if (!sh) { sh = ss.insertSheet('_APP DATA'); sh.hideSheet(); }
  return sh;
}
function findDataRow(sh, name) {
  var last = sh.getLastRow();
  if (!last) return 0;
  var names = sh.getRange(1, 1, last, 1).getValues();
  for (var i = 0; i < names.length; i++) if (names[i][0] === name) return i + 1;
  return 0;
}
function saveJson(name, obj) {
  var sh = dataSheet();
  var s = JSON.stringify(obj);
  var chunks = [];
  for (var i = 0; i < s.length; i += 40000) chunks.push(s.substr(i, 40000));
  var row = findDataRow(sh, name) || (sh.getLastRow() + 1);
  var wide = Math.max(sh.getLastColumn(), chunks.length + 1, 2);
  sh.getRange(row, 1, 1, wide).clearContent();
  sh.getRange(row, 1).setValue(name);
  if (chunks.length) sh.getRange(row, 2, 1, chunks.length).setNumberFormat('@').setValues([chunks]);
}
function loadJson(name, fallback) {
  var sh = dataSheet();
  var row = findDataRow(sh, name);
  if (!row) return fallback;
  var vals = sh.getRange(row, 2, 1, Math.max(sh.getLastColumn() - 1, 1)).getValues()[0];
  var s = '';
  for (var i = 0; i < vals.length; i++) if (vals[i] !== '') s += vals[i];
  try { return s ? JSON.parse(s) : fallback; } catch (err) { return fallback; }
}

/* ================= photos: embedded in APP PHOTOS tab ================= */
function photoSheet(ss) {
  var sh = ss.getSheetByName('APP PHOTOS');
  if (!sh) sh = ss.insertSheet('APP PHOTOS');
  sh.getRange(1, 1, 1, 4).setValues([['KART', 'DATE', 'PHOTO', '']]).setFontWeight('bold');
  sh.getRange('A:A').setHorizontalAlignment('center');
  sh.getRange('B:B').setHorizontalAlignment('center');
  sh.setColumnWidth(3, 260);
  return sh;
}
function savePhoto(data) {
  if (!data.id || !data.dataURL) throw 'photo missing id/dataURL';
  var index = loadJson('photo_index', {});
  if (index[data.id]) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = photoSheet(ss);
  var row = sh.getLastRow() + 1;
  var placed = false;
  try {
    // in-cell image: locked to the cell, sorts and moves with the row
    var cellImg = SpreadsheetApp.newCellImage()
      .setSourceUrl(String(data.dataURL))
      .setAltTextTitle(data.id)
      .build();
    sh.getRange(row, 3).setValue(cellImg);
    sh.setRowHeight(row, 150);
    placed = true;
  } catch (err) {
    // fallback: floating image anchored at the cell
    var base64 = String(data.dataURL).split(',')[1];
    if (!base64) throw 'bad dataURL';
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', data.id + '.jpg');
    var img = sh.insertImage(blob, 3, row);
    var w = img.getWidth(), h = img.getHeight();
    if (w > 240) { img.setHeight(Math.round(h * 240 / w)); img.setWidth(240); }
    sh.setRowHeight(row, Math.max(img.getHeight() + 10, 60));
    placed = true;
  }
  if (!placed) throw 'image placement failed';
  sh.getRange(row, 1, 1, 2).setValues([[data.kart || '?', data.date || '']]);
  index[data.id] = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=C' + row;
  saveJson('photo_index', index);
}

/* ---- sorting log entries by date ----
   Entry dates are typed by hand and a few old rows have no date at all.
   `new Date('')` is an Invalid Date, and every comparison against NaN comes
   back false, so the old comparators answered inconsistently for those rows —
   the sort then left APP LOG in several separate newest-first runs and a brand
   new entry could land in the middle of the tab instead of on top. Turning the
   date into a plain number first, with undated rows pinned to "oldest", keeps
   the ordering total so the sort can't scramble. */
function entryKey(s) {
  var t = String(s === null || s === undefined ? '' : s).trim();
  if (t) {
    var m = t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (m) {
      var y = +m[3];
      return new Date(y < 100 ? 2000 + y : y, +m[1] - 1, +m[2]).getTime();
    }
    var n = new Date(t).getTime();
    if (n === n) return n;
  }
  return -Infinity;                 /* no usable date -> sorts as oldest */
}
/* -Infinity minus -Infinity is NaN, so compare rather than subtract */
function cmpKey(a, b) { return a < b ? -1 : (a > b ? 1 : 0); }

/* ================= date helpers for coloring ================= */
function parseAnyDate(s) {
  if (!s) return null;
  var m = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
  return null;
}
function parseBat(s) { // "8-24" or "8/24" -> Date(2024, 7)
  if (!s) return null;
  var m = String(s).match(/^(\d{1,2})[\/-](\d{2})$/);
  if (m) return new Date(2000 + (+m[2]), +m[1] - 1, 1);
  return null;
}
function svcColor(s) {
  var d = parseAnyDate(s);
  if (!d) return '#ffffff';
  var days = (Date.now() - d.getTime()) / 86400000;
  if (days > 90) return '#f4c7c3';   // red — overdue
  if (days > 75) return '#fce8b2';   // yellow — coming up
  return '#d9ead3';                  // green — fresh
}
function batColor(s) {
  var d = parseBat(s);
  if (!d) return '#ffffff';
  var months = (Date.now() - d.getTime()) / (86400000 * 30.4);
  if (months > 24) return '#f4c7c3';
  if (months > 18) return '#fce8b2';
  return '#d9ead3';
}

/* ================= sheet writers ================= */
/* Apps Script batches spreadsheet writes and flushes them AFTER the calling
   function returns — so a rejected write throws outside any try/catch here
   unless we force the flush ourselves. Everything guarded below must flush. */
function tryOp(fn) {
  try { fn(); SpreadsheetApp.flush(); return ''; }
  catch (e) { return String(e) || 'failed'; }
}

/* Best-effort text formatting. Returns false if any column refused
   (a Sheets "table" makes typed columns un-formattable). Never throws. */
function fmtTextCols(sh, cols) {
  if (!cols) return true;
  var ok = true;
  for (var i = 0; i < cols.length; i++) {
    var col = cols[i];
    var err = tryOp(function () { sh.getRange(col + ':' + col).setNumberFormat('@'); });
    if (err) ok = false;
  }
  return ok;
}

function freshSheet(ss, name, textCols) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); fmtTextCols(sh, textCols); return sh; }

  var cleared = !tryOp(function () { sh.clearContents(); });
  if (cleared && fmtTextCols(sh, textCols)) return sh;

  /* The tab is a Sheets "table" (or otherwise unwritable): stand up a clean
     replacement. Rename the old one FIRST so the new tab can take the real
     name even if the delete is refused — a stray "<name> OLD" tab is a far
     better outcome than an empty tab and an aborted write. */
  var pos = 1;
  try { pos = sh.getIndex(); } catch (e1) { pos = 1; }
  var junk = name + ' OLD';
  tryOp(function () {
    var prior = ss.getSheetByName(junk);
    if (prior) ss.deleteSheet(prior);
  });
  if (tryOp(function () { sh.setName(junk); })) return sh;  /* name not freed */

  var nsh = null;
  if (tryOp(function () { nsh = ss.insertSheet(name, pos - 1); }))
    tryOp(function () { nsh = ss.insertSheet(name); });
  if (!nsh) { tryOp(function () { sh.setName(name); }); return sh; }
  tryOp(function () { ss.deleteSheet(ss.getSheetByName(junk)); });
  fmtTextCols(nsh, textCols);
  return nsh;
}
function kartOrder(karts) {
  return Object.keys(karts).sort(function (a, b) { return (+a) - (+b); });
}
function writeAllDates(ss, data) {
  var sh = freshSheet(ss, 'APP ALL DATES', ['E', 'F', 'G', 'H']);
  var rows = [['KART', 'CHAIN', 'DIFF', 'BRAKE FLUSH',
               'BAT 1', 'BAT 2', 'BAT 3', 'BAT 4', 'WELD COUNT', 'KART NOTES']];
  var colors = [['#cccccc', '#cccccc', '#cccccc', '#cccccc', '#cccccc',
                 '#cccccc', '#cccccc', '#cccccc', '#cccccc', '#cccccc']];
  var ks = kartOrder(data.karts);
  for (var i = 0; i < ks.length; i++) {
    var k = ks[i], st = data.karts[k].status || {};
    rows.push([k, st.chain || '', st.diff || '', st.brake || '',
               st.bat1 || '', st.bat2 || '', st.bat3 || '', st.bat4 || '',
               st.weld || 0, data.karts[k].knotes || '']);
    colors.push(['#ffffff', svcColor(st.chain), svcColor(st.diff), svcColor(st.brake),
                 batColor(st.bat1), batColor(st.bat2), batColor(st.bat3), batColor(st.bat4),
                 '#ffffff', '#ffffff']);
  }
  var rng = sh.getRange(1, 1, rows.length, rows[0].length);
  rng.setValues(rows);
  rng.setBackgrounds(colors);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.getRange(1, 12).setValue('updated ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d h:mm a'));
  sh.setFrozenRows(1);
}
var KART_HDR = ['DATE', 'ACTION COMPLETED', 'parts used', 'MECHANIC', 'NOTES',
                'CHAIN DATE', 'DIFF DATE', 'BRAKE FLUSH', 'BAT 1 DATE', 'BAT 2 DATE',
                'BAT 3 DATE', 'BAT 4 DATE', 'WELD COUNT'];
function kartSig(s) {
  var h = 5381;
  for (var i = 0; i < s.length; i++) h = (h * 33 + s.charCodeAt(i)) % 4294967296;
  return String(h);
}
function writeKartTabs(ss, data) {
  var sigs = loadJson('kart_sigs', {});
  var ks = kartOrder(data.karts);
  var wrote = 0;
  for (var i = 0; i < ks.length; i++) {
    var k = ks[i];
    var kd = data.karts[k], st = kd.status || {};
    var sig = kartSig(JSON.stringify(kd));
    if (sigs[k] === sig) continue;   // unchanged since last write
    var sh = ss.getSheetByName(k);
    if (!sh) sh = ss.insertSheet(k);
    var rows = [KART_HDR.slice()];
    rows.push(['', '', '', '', kd.knotes || '', st.chain || '', st.diff || '', st.brake || '',
               st.bat1 || '', st.bat2 || '', st.bat3 || '', st.bat4 || '', st.weld || 0]);
    var es = (kd.entries || []).slice().sort(function (a, b) {
      return cmpKey(entryKey(a.date), entryKey(b.date));   /* oldest first */
    });
    for (var j = 0; j < es.length; j++) {
      var en = es[j];
      rows.push([en.date || '', en.action || '', en.parts || '', en.mechanic || '',
                 en.notes || '', '', '', '', '', '', '', '', '']);
    }
    sh.clearContents();
    sh.getRange('C:C').setNumberFormat('@');
    sh.getRange('I:L').setNumberFormat('@');
    sh.getRange(1, 1, rows.length, KART_HDR.length).setValues(rows);
    sh.getRange(1, 1, 1, KART_HDR.length).setFontWeight('bold');
    sigs[k] = sig;
    wrote++;
  }
  if (wrote) saveJson('kart_sigs', sigs);
}
function photoCell(photos, index) {
  if (!photos) return '';
  if (typeof photos === 'number') return photos ? photos + ' in app' : '';
  if (!photos.length) return '';
  var locs = [], missing = 0;
  for (var i = 0; i < photos.length; i++) {
    var u = index[photos[i]];
    if (u) {
      var m = u.match(/range=[A-Z](\d+)/);
      locs.push(m ? 'APP PHOTOS row ' + m[1] : u);
    } else missing++;
  }
  var cell = locs.join(', ');
  if (missing) cell += (cell ? ' + ' : '') + missing + ' not synced yet';
  return cell;
}
function writeLog(ss, data, photoIndex) {
  var sh = freshSheet(ss, 'APP LOG', ['D']);
  var rows = [['DATE', 'KART', 'ACTION COMPLETED', 'PARTS USED',
               'MECHANIC', 'NOTES', 'PHOTOS']];
  var ks = kartOrder(data.karts);
  var all = [];
  for (var i = 0; i < ks.length; i++) {
    var es = data.karts[ks[i]].entries || [];
    for (var j = 0; j < es.length; j++) {
      var en = es[j];
      all.push([en.date || '', ks[i], en.action || '', en.parts || '',
                en.mechanic || '', en.notes || '', photoCell(en.photos, photoIndex)]);
    }
  }
  /* newest first, undated rows at the bottom; the original position is carried
     along as a last tie-break so two entries logged the same day on the same
     kart keep a fixed order instead of shuffling on every rebuild */
  for (var q = 0; q < all.length; q++) all[q].push(q);
  all.sort(function (a, b) {
    return cmpKey(entryKey(b[0]), entryKey(a[0])) ||
           ((+a[1]) - (+b[1])) || (a[7] - b[7]);
  });
  for (var q2 = 0; q2 < all.length; q2++) all[q2].pop();
  rows = rows.concat(all);
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
function writePartsUsed(ss, data) {
  var sh = freshSheet(ss, 'APP PARTS USED', ['A']);
  var desc = {};
  if (data.parts) for (var p = 0; p < data.parts.length; p++)
    desc[String(data.parts[p][0]).toUpperCase()] = data.parts[p][1];
  var tot = {}, last = {};
  var sources = [];
  var ks = Object.keys(data.karts);
  for (var i = 0; i < ks.length; i++) sources.push(data.karts[ks[i]].entries || []);
  sources.push(data.shop || []);
  for (var s = 0; s < sources.length; s++) {
    var es = sources[s];
    for (var j = 0; j < es.length; j++) {
      var en = es[j];
      if (!en.parts) continue;
      var items = String(en.parts).split(',');
      for (var t = 0; t < items.length; t++) {
        var it = items[t].trim();
        if (!it) continue;
        var m = it.match(/^(.+?)\s*[xX]\s*(\d+)$/);
        var num = m ? m[1].trim().toUpperCase() : it.toUpperCase();
        var qty = m ? (parseInt(m[2], 10) || 1) : 1;
        tot[num] = (tot[num] || 0) + qty;
        var d = new Date(en.date);
        if (!isNaN(d.getTime()) && (!last[num] || d > last[num])) last[num] = d;
      }
    }
  }
  var rows = [['PART NUMBER', 'DESCRIPTION', 'TOTAL USED', 'LAST USED']];
  var nums = Object.keys(tot).sort(function (a, b) { return tot[b] - tot[a]; });
  for (var n = 0; n < nums.length; n++) {
    var nm = nums[n];
    rows.push([nm, desc[nm] || '', tot[nm],
               last[nm] ? Utilities.formatDate(last[nm], ss.getSpreadsheetTimeZone(), 'M/d/yyyy') : '']);
  }
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/* ================= v4: INVENTORY & ORDERING ================= */
function buildMenu() {
  SpreadsheetApp.getUi().createMenu('K1 Kart Log')
    .addItem('Phase 2 → Place order (move checked to APP ORDERS)', 'placeOrder')
    .addItem('Book received quantities into stock', 'bookReceived')
    .addSeparator()
    .addItem('Rebuild tabs from latest data', 'cleanupSheet')
    .addToUi();
}

/* ---- inventory names/thresholds served to the app + receipts ---- */
function invConfigFromSheet(ss) {
  var sh = ss.getSheetByName('inventory');
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var wide = Math.min(13, Math.max(6, sh.getMaxColumns() - 1));
  var vals = sh.getRange(2, 2, last - 1, wide).getValues(); // B..N: part, name, qty, inventoried, red, green ... fits
  for (var i = 0; i < vals.length; i++) {
    var num = vals[i][0];
    if (num === '' || num === null) continue;
    var nm = String(vals[i][1] || num).trim();
    out[String(num).trim().toUpperCase()] = {
      n: nm,
      r: parseInt(vals[i][4], 10) || 0,
      g: parseInt(vals[i][5], 10) || 0,
      k: normKind(vals[i][12]),
      retired: /\(RETIRED\)/i.test(nm) ? 1 : 0
    };
  }
  return out;
}

/* ---- write app quantities back into the inventory tab QUANTITY column ----
   Rows are added, removed and renamed by mergeCfg; this only moves numbers. */
function writeInventoryQty(ss, inv, cfgFromApp, counted) {
  var sh = ss.getSheetByName('inventory');
  if (!sh || !inv) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var nums = sh.getRange(2, 2, last - 1, 1).getValues();
  var qtys = sh.getRange(2, 4, last - 1, 1).getValues();
  /* Column E INVENTORIED is the recount checkbox. It is ticked off on the app's
     INVENTORY screen, and mirrored here so the sheet shows the same thing.
     RESET in the app empties invCounted, so every box goes back to FALSE. */
  var ticks = counted ? sh.getRange(2, 5, last - 1, 1).getValues() : null;
  for (var i = 0; i < nums.length; i++) {
    var key = String(nums[i][0] || '').trim().toUpperCase();
    if (key && inv[key] !== undefined) qtys[i][0] = inv[key];
    if (ticks) ticks[i][0] = !!(key && counted[key]);
  }
  sh.getRange(2, 4, last - 1, 1).setValues(qtys);
  if (ticks) tryOp(function () { sh.getRange(2, 5, last - 1, 1).setValues(ticks); });
}

/* ================= v8: two-way part config =================
   Part names, numbers and thresholds can be changed in the app or typed
   straight onto the inventory tab. To tell which side actually moved (rather
   than letting the last writer win), we keep a mirror of the last agreed
   config in the hidden tab and use it as the common ancestor.

   Timestamps are the app's own Date.now() stamps, never the sheet's clock, so
   the two sides never have to agree on time — cfgTouched values from the app
   are compared against the newest cfgTouched the mirror was built from. */

/* inventory tab as a lookup: B=part C=name D=qty E=counted F=red G=green N=fits */
function invRows(ss) {
  var sh = ss.getSheetByName('inventory');
  var out = { sh: sh, byNum: {} };
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var wide = Math.min(13, Math.max(6, sh.getMaxColumns() - 1));
  var vals = sh.getRange(2, 2, last - 1, wide).getValues();
  for (var i = 0; i < vals.length; i++) {
    var num = String(vals[i][0] === null ? '' : vals[i][0]).trim().toUpperCase();
    if (!num || out.byNum[num]) continue;      /* first row wins on duplicates */
    var nm = String(vals[i][1] || num).trim();
    out.byNum[num] = { row: i + 2, n: nm,
                       r: parseInt(vals[i][4], 10) || 0,
                       g: parseInt(vals[i][5], 10) || 0,
                       k: normKind(vals[i][12]),
                       retired: /\(RETIRED\)/i.test(nm) ? 1 : 0 };
  }
  return out;
}

function sameCfg(a, b) {
  return !!a && !!b && a.n === b.n && (a.r || 0) === (b.r || 0) && (a.g || 0) === (b.g || 0) &&
         (a.k || '') === (b.k || '');
}

/* ---- FITS column (N): which karts a part goes on ----
   Stored in the app as 'A' (adult only), 'J' (jr only), 'N' (never on a kart --
   shop/facility stock) or '' (fits both adult and jr), and typed on the sheet as
   a word. Anything unrecognised reads as "both", so an empty column on 400-odd
   existing rows means nothing changes for them. */
function normKind(v) {
  var s = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
  if (!s) return '';
  var c = s.charAt(0);
  if (c === 'A') return 'A';
  if (c === 'J') return 'J';
  if (c === 'N') return 'N';   /* NOT KART -- shop/facility part, never on a kart */
  return '';
}
function kindLabel(k) {
  return k === 'A' ? 'ADULT' : (k === 'J' ? 'JR' : (k === 'N' ? 'NOT KART' : 'BOTH'));
}

/* make sure column N exists and is labelled before we write into it */
function ensureFitsColumn(sh) {
  if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
  var h = sh.getRange(1, 14);
  if (String(h.getValue() || '').trim() === '') {
    h.setValue('FITS');
    tryOp(function () { h.setFontWeight('bold'); });
  }
}

function mergeCfg(ss, snap) {
  var sh = ss.getSheetByName('inventory');
  if (!sh) return { names: {}, del: null, at: 0 };
  tryOp(function () { ensureFitsColumn(sh); });

  var ours = (snap && snap.invCfg) || {};
  var touched = (snap && snap.cfgTouched) || {};
  var tomb = (snap && snap.partTomb) || {};
  var rekeys = (snap && snap.rekeys) || [];
  var inv = (snap && snap.inv) || {};

  /* how current the app data in this merge is, in the app's own clock */
  var snapAt = 0, kk;
  for (kk in touched) if (touched[kk] > snapAt) snapAt = touched[kk];
  for (kk in tomb) if (tomb[kk] > snapAt) snapAt = tomb[kk];

  var base = loadJson('cfg_mirror', null);
  var haveMirror = !!base;
  if (!base) base = {};
  var mirrorAt = base._at || 0;
  delete base._at;

  var rows = invRows(ss);
  var theirs = rows.byNum;
  var del = loadJson('cfg_del', {});
  var names = {};
  var edits = [], dropRows = [], appends = [];

  /* --- 1. renames made in the app: carry the sheet row to the new number --- */
  var rkSeen = {};
  var rkDone = loadJson('cfg_rekeys', {});
  for (var q = 0; q < rekeys.length; q++) {
    var rk = rekeys[q];
    if (!rk || !rk.from || !rk.to) continue;
    var sig = rk.from + '>' + rk.to + '@' + (rk.at || 0);
    rkSeen[sig] = 1;
    if (rkDone[sig]) continue;
    var src = theirs[rk.from];
    if (src) {
      if (theirs[rk.to]) dropRows.push(src.row);          /* target row already there */
      else {
        edits.push({ row: src.row, col: 2, v: rk.to });
        theirs[rk.to] = { row: src.row, n: src.n, r: src.r, g: src.g, k: src.k,
                          retired: src.retired };
      }
      delete theirs[rk.from];
    }
    if (base[rk.from]) {
      if (!base[rk.to]) base[rk.to] = base[rk.from];
      delete base[rk.from];
    }
    delete del[rk.from];
  }
  /* the app caps its rename list, so trim ours to match and stop it growing */
  saveJson('cfg_rekeys', rkSeen);

  /* --- 2. resolve every part we know about from any of the three sides --- */
  var keys = {}, k;
  for (k in base) keys[k] = 1;
  for (k in theirs) keys[k] = 1;
  for (k in ours) keys[k] = 1;

  for (k in keys) {
    var b = base[k], t = theirs[k], o = ours[k];

    /* deleted in the app — pull the row and stop tracking it */
    if (tomb[k]) {
      if (t) dropRows.push(t.row);
      delete base[k]; delete del[k];
      continue;
    }
    /* legacy "(RETIRED)" rows: report them, never rewrite or delete them */
    if (t && t.retired) { names[k] = { n: t.n, r: t.r, g: t.g, k: t.k, retired: 1 }; continue; }

    if (t && !o) {                       /* typed onto the sheet, app doesn't have it */
      names[k] = { n: t.n, r: t.r, g: t.g, k: t.k || '' };
      base[k] = { n: t.n, r: t.r, g: t.g, k: t.k || '' };
      delete del[k];
      continue;
    }
    if (o && !t) {                       /* in the app, no row on the sheet */
      if (haveMirror && b && (touched[k] || 0) <= mirrorAt) {
        /* it had a row at the last merge and the app hasn't touched it since,
           so the row was deleted on the sheet — hand that back to the app */
        del[k] = mirrorAt || 1;
        delete base[k];
        continue;
      }
      appends.push(k);                   /* new in the app (or re-added after a delete) */
      names[k] = { n: o.n, r: o.r, g: o.g, k: o.k || '' };
      base[k] = { n: o.n, r: o.r, g: o.g, k: o.k || '' };
      delete del[k];
      continue;
    }
    /* on both sides: only the side that actually changed gets to win */
    var sheetMoved = !sameCfg(b, t);
    var appMoved = !sameCfg(b, o);
    var win;
    /* First run ever: there is no ancestor to compare against, so don't guess —
       adopt the sheet as the starting point and merge properly from here on.
       The exception is a part the app has explicitly stamped: cfgTouched only
       exists from v8.0 on, so a stamp here is an edit made in the app that the
       sheet has never seen, and dropping it would undo the user's typing. */
    if (!haveMirror) win = touched[k] ? o : t;
    else if (appMoved && !sheetMoved) win = o;
    else if (!appMoved) win = t;
    else win = ((touched[k] || 0) > mirrorAt) ? o : t;   /* both moved: newer edit */
    if (win !== t) {
      if (win.n !== t.n) edits.push({ row: t.row, col: 3, v: win.n });
      if ((win.r || 0) !== t.r) edits.push({ row: t.row, col: 6, v: win.r || 0 });
      if ((win.g || 0) !== t.g) edits.push({ row: t.row, col: 7, v: win.g || 0 });
      if ((win.k || '') !== (t.k || '')) edits.push({ row: t.row, col: 14, v: kindLabel(win.k || '') });
    }
    names[k] = { n: win.n, r: win.r, g: win.g, k: win.k || '' };
    base[k] = { n: win.n, r: win.r, g: win.g, k: win.k || '' };
    delete del[k];
  }

  /* --- 3. write it: edits before deletes, since deletes shift rows up --- */
  var i;
  for (i = 0; i < edits.length && i < 300; i++)
    sh.getRange(edits[i].row, edits[i].col).setValue(edits[i].v);
  if (dropRows.length) {
    dropRows.sort(function (a, b2) { return b2 - a; });
    var prev = -1;
    for (i = 0; i < dropRows.length; i++) {
      if (dropRows[i] === prev) continue;
      prev = dropRows[i];
      sh.deleteRow(dropRows[i]);
    }
  }
  if (appends.length) {
    var newRows = [];
    if (appends.length > 500) appends = appends.slice(0, 500);
    var newFits = [];
    for (i = 0; i < appends.length; i++) {
      var an = appends[i], ac = ours[an];
      var q2 = inv[an] !== undefined ? inv[an] : 0;
      newRows.push(['', an, ac.n || an, q2, '', ac.r || 0, ac.g || 0]);
      newFits.push([ac.k ? kindLabel(ac.k) : '']);
    }
    var start = sh.getLastRow() + 1;
    tryOp(function () { sh.getRange(start, 2, newRows.length, 1).setNumberFormat('@'); });
    sh.getRange(start, 1, newRows.length, 7).setValues(newRows);
    tryOp(function () { sh.getRange(start, 14, newFits.length, 1).setValues(newFits); });
  }
  SpreadsheetApp.flush();

  /* keep the handback list from growing without bound */
  var dk = [], dn;
  for (dn in del) dk.push(dn);
  if (dk.length > 300) {
    dk.sort(function (a, b2) { return (del[a] || 0) - (del[b2] || 0); });
    for (i = 0; i < dk.length - 300; i++) delete del[dk[i]];
  }
  saveJson('cfg_del', del);
  base._at = snapAt;
  saveJson('cfg_mirror', base);
  delete base._at;
  return { names: names, del: del, at: snapAt };
}

/* ---- APP NEEDED: red + yellow items; preserves your checks & order qtys ---- */
function writeNeeded(ss, inv) {
  var cfg = invConfigFromSheet(ss);
  var sh = ss.getSheetByName('APP NEEDED');
  var kept = {}; // part -> {chk, qty}
  if (sh) {
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last > 1 && lastCol >= 8) {
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var chkCol = -1, qtyCol = -1;
      for (var h = 0; h < hdr.length; h++) {
        if (String(hdr[h]).indexOf('ORDER ✓') > -1) chkCol = h;
        if (String(hdr[h]).indexOf('ORDER QTY') > -1) qtyCol = h;
      }
      if (chkCol > -1 && qtyCol > -1) {
        var old = sh.getRange(2, 1, last - 1, lastCol).getValues();
        for (var i = 0; i < old.length; i++) {
          var k = String(old[i][0] || '').trim().toUpperCase();
          if (k) kept[k] = { chk: old[i][chkCol] === true, qty: old[i][qtyCol] };
        }
      }
    }
    sh.clear();
  } else {
    sh = ss.insertSheet('APP NEEDED');
  }
  sh.getRange('A:A').setNumberFormat('@');
  var rows = [['PART #', 'NAME', 'QTY NOW', 'RED ≤', 'GREEN ≥', 'STATUS', 'TO GREEN', 'ORDER ✓', 'ORDER QTY']];
  var colors = [['#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc']];
  for (var num in cfg) {
    var q = (inv && inv[num] !== undefined) ? inv[num] : '';
    if (q === '') continue;
    var c = cfg[num];
    if (!c.r && !c.g) continue;      /* no levels set = not something we track */
    var status = q <= c.r ? 'NEEDED' : (q < c.g ? 'WANTED' : '');
    if (!status) continue;
    var keep = kept[num] || { chk: false, qty: '' };
    var myQty = (keep.qty === '' || keep.qty === null || keep.qty === undefined) ? '' : keep.qty;
    rows.push([num, c.n, q, c.r, c.g, status, Math.max(c.g - q, 0), keep.chk, myQty]);
    var bg = status === 'NEEDED' ? '#f4c7c3' : '#fce8b2';
    colors.push(['#ffffff', '#ffffff', bg, '#ffffff', '#ffffff', bg, '#ffffff', '#ffffff', '#fff9c4']);
  }
  sh.getRange(1, 1, rows.length, 9).setValues(rows).setBackgrounds(colors);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  if (rows.length > 1) sh.getRange(2, 8, rows.length - 1, 1).insertCheckboxes();
  sh.setFrozenRows(1);
}

/* ---- APP ORDER (phase 2): live view of what you've checked ---- */
function ensureOrderView(ss) {
  var sh = ss.getSheetByName('APP ORDER');
  if (!sh) sh = ss.insertSheet('APP ORDER');
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['PART #', 'NAME', 'QTY NOW', 'ORDERING']]).setFontWeight('bold');
  sh.getRange('A2').setFormula(
    "=IFERROR(FILTER({'APP NEEDED'!A2:C, 'APP NEEDED'!I2:I}, 'APP NEEDED'!H2:H=TRUE), \"nothing checked yet\")");
  sh.setFrozenRows(1);
}

/* ---- menu: move checked items into APP ORDERS with the date ---- */
function placeOrderImpl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var needed = ss.getSheetByName('APP NEEDED');
  if (!needed) { SpreadsheetApp.getUi().alert('No APP NEEDED tab yet.'); return; }
  var last = needed.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Nothing on APP NEEDED.'); return; }
  var vals = needed.getRange(2, 1, last - 1, 9).getValues();
  var orders = ss.getSheetByName('APP ORDERS');
  if (!orders) {
    orders = ss.insertSheet('APP ORDERS');
    orders.getRange(1, 1, 1, 8).setValues([['DATE ORDERED', 'PART #', 'NAME', 'QTY ORDERED', 'QTY RECEIVED', 'QTY BOOKED', 'LAST BOOKED', 'ORDER #']])
      .setFontWeight('bold');
    orders.getRange('B:B').setNumberFormat('@');
    orders.setFrozenRows(1);
  }
  var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  var orderNo = 'ORD-' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmm');
  var newRows = [], clearRows = [], skipped = 0;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][7] === true) {
      var qty = parseInt(vals[i][8], 10) || 0;
      if (qty > 0) { newRows.push([today, String(vals[i][0]), vals[i][1], qty, '', 0, '', orderNo]); clearRows.push(i + 2); }
      else skipped++;
    }
  }
  if (!newRows.length) {
    SpreadsheetApp.getUi().alert(skipped
      ? skipped + ' checked item(s) have no ORDER QTY — type how many you\'re ordering first.'
      : 'Nothing checked.');
    return;
  }
  orders.getRange(orders.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  for (var r = 0; r < clearRows.length; r++) needed.getRange(clearRows[r], 8, 1, 2).setValues([[false, '']]);
  var msg = 'Order ' + orderNo + ' logged: ' + newRows.length + ' items moved to APP ORDERS.';
  if (skipped) msg += '\n\n' + skipped + ' checked item(s) SKIPPED — no ORDER QTY typed.';
  msg += '\n\nWhen boxes arrive, type what you got in QTY RECEIVED, then run "Book received quantities into stock".';
  SpreadsheetApp.getUi().alert(msg);
}

/* ---- menu: book received quantities -> receipts the app applies ---- */
function bookReceivedImpl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = ss.getSheetByName('APP ORDERS');
  if (!orders) { SpreadsheetApp.getUi().alert('No APP ORDERS tab yet.'); return; }
  var last = orders.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('No orders logged.'); return; }
  var vals = orders.getRange(2, 1, last - 1, 8).getValues();
  var receipts = loadJson('receipts', []);
  var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  var booked = 0, units = 0;
  for (var i = 0; i < vals.length; i++) {
    var received = parseInt(vals[i][4], 10);
    var already = parseInt(vals[i][5], 10) || 0;
    if (isNaN(received)) continue;
    var delta = received - already;
    if (delta > 0) {
      receipts.push({ id: 'rc' + Date.now() + '_' + i, part: String(vals[i][1]).toUpperCase(), qty: delta });
      orders.getRange(i + 2, 6).setValue(received);   // QTY BOOKED = new running total
      orders.getRange(i + 2, 7).setValue(today);      // LAST BOOKED
      booked++; units += delta;
    }
  }
  if (!booked) { SpreadsheetApp.getUi().alert('Nothing new to book — QTY RECEIVED matches what\'s already booked on every row.'); return; }
  if (receipts.length > 400) receipts = receipts.slice(receipts.length - 400);
  saveJson('receipts', receipts);
  SpreadsheetApp.getUi().alert(booked + ' row(s), ' + units + ' unit(s) booked into stock. The app applies them next time it opens or syncs. Split shipments: just raise QTY RECEIVED when the rest arrives and book again — only the new amount gets added.');
}

/* ================= APP COUNT SHEET (v8.4) =================
   The monthly recount is walked on paper. This tab is that paper typed up, in
   the same running order, and it exists so Robbie can read the counted numbers
   off it and copy them onto the sheet of paper — nothing more. Three columns:

       A PART #    B NAME    C QTY

   Rows are posted in as a transcription (type:'count'). QTY is filled in by the
   script from the app: whatever a part's stock is once it has been ticked off on
   the app's recount screen shows up here. Parts not yet counted stay blank, and
   hitting RESET in the app blanks them again for the next round.                */

var COUNT_HDR = ['PART #', 'NAME', 'QTY'];

function countSheet(ss, make) {
  var sh = ss.getSheetByName(COUNT_TAB);
  if (!sh) {
    if (!make) return null;
    sh = ss.insertSheet(COUNT_TAB);
  }
  if (String(sh.getRange(1, 1).getValue() || '').trim() !== COUNT_HDR[0]) {
    /* first build, or an older/wider layout being replaced — start clean so no
       stray columns are left sitting to the right of the new ones */
    tryOp(function () { sh.clearContents(); });
    sh.getRange(1, 1, 1, COUNT_HDR.length).setValues([COUNT_HDR]);
    tryOp(function () { sh.getRange(1, 1, 1, COUNT_HDR.length).setFontWeight('bold'); });
    tryOp(function () { sh.setFrozenRows(1); });
    tryOp(function () { sh.getRange('A:A').setNumberFormat('@'); });
    tryOp(function () { sh.setColumnWidth(1, 110); });
    tryOp(function () { sh.setColumnWidth(2, 430); });
    tryOp(function () { sh.setColumnWidth(3, 70); });
  }
  return sh;
}

/* The paper prints part numbers with a leading zero ("059012") while the
   inventory tab stores them bare ("59012"). Try it as typed first, then with
   zeros stripped, then padded, so either spelling finds the same part. */
function matchPart(raw, byNum) {
  var s = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
  if (!s) return '';
  if (byNum[s]) return s;
  var bare = s.replace(/^0+/, '');
  if (bare && byNum[bare]) return bare;
  for (var z = 1; z <= 3; z++) {
    var pad = new Array(z + 1).join('0') + bare;
    if (byNum[pad]) return pad;
  }
  return '';
}

/* ---- POST type:'count' -> lay a transcribed paper page onto the tab ----
   data.rows: [{num: part number, d: name (optional)}], in paper order.
   data.replace: 1 wipes what's there first, otherwise rows are appended.
   Names are filled in from the inventory tab when not supplied, so a
   transcription only really has to carry the numbers. */
function writeCountRows(ss, data) {
  var rows = (data && data.rows) || [];
  if (!rows.length && !(data && data.replace)) return 'no rows';
  var sh = countSheet(ss, true);
  /* handlePost already holds the script lock for this request */
  if (data.replace && sh.getLastRow() > 1)
    sh.getRange(2, 1, sh.getLastRow() - 1, COUNT_HDR.length).clearContent();
  var byNum = invRows(ss).byNum;
  var out = [], miss = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || {};
    var key = matchPart(r.num, byNum);
    var known = key ? byNum[key] : null;
    if (!key) miss++;
    out.push([String(r.num === undefined ? '' : r.num),
              r.d || (known ? known.n : (key ? '' : '?? not in inventory')),
              '']);
  }
  if (out.length) {
    var start = Math.max(2, sh.getLastRow() + 1);
    if (sh.getMaxRows() < start + out.length)
      sh.insertRowsAfter(sh.getMaxRows(), start + out.length - sh.getMaxRows());
    tryOp(function () { sh.getRange(start, 1, out.length, 1).setNumberFormat('@'); });
    sh.getRange(start, 1, out.length, COUNT_HDR.length).setValues(out);
  }
  SpreadsheetApp.flush();
  return 'ok count: ' + out.length + ' rows' + (miss ? ', ' + miss + ' with no matching part' : '');
}

/* ---- fill the paper page in from the app's recount ----
   Every pass: any part the app has ticked off as counted gets its quantity
   written into QTY, parts it hasn't counted stay blank, and a part that loses
   its checkmark (RESET in the app) goes blank again. Blank names are backfilled
   from the inventory tab. Nothing flows the other way — the counting happens on
   the iPad, this tab just shows the numbers to copy onto the paper. */
function scanCounts(ss, snap) {
  var sh = countSheet(ss, false);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rows = sh.getRange(2, 1, last - 1, COUNT_HDR.length).getValues();
  var byNum = invRows(ss).byNum;
  var inv = (snap && snap.inv) || {};
  var counted = (snap && snap.invCounted) || {};
  var filled = 0, dirty = false;

  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i][0];
    if (raw === '' || raw === null) continue;
    var key = matchPart(raw, byNum);
    var known = key ? byNum[key] : null;

    if (!rows[i][1] && known) { rows[i][1] = known.n; dirty = true; }

    var want = '';
    if (key && counted[key]) {
      want = (inv[key] === undefined || inv[key] === null) ? 0 : inv[key];
      filled++;
    }
    if (rows[i][2] !== want) { rows[i][2] = want; dirty = true; }
  }

  if (dirty) sh.getRange(2, 1, rows.length, COUNT_HDR.length).setValues(rows);
  SpreadsheetApp.flush();
  return filled;
}
