/* K1 Kart Log — sheet-side LOGIC (fetched by the shell installed in Apps Script).
   Edit this file, push to GitHub, and the sheet picks it up within ~10 minutes
   (or instantly after running refreshLogic in the Apps Script editor). */

/** K1 Kart Log receiver v3 — the sheet is the app's full storage.
 *  Writes: APP ALL DATES, APP LOG, APP PARTS USED, APP PHOTOS,
 *  kart tabs 1-53, appends to "parts used", hidden _APP DATA.
 *  Never touches inventory tabs' content or the template. */

var KART_TABS = (function(){ var a=[]; for (var i=1;i<=53;i++) a.push(String(i)); return a; })();

function handlePost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data || data.app !== 'k1kartlog') return txt('ignored');
    var lock = LockService.getScriptLock();
    lock.waitLock(40000);
    try {
      if (data.type === 'photo') { savePhoto(data); return txt('ok photo'); }
      if (data.type === 'snapshot') {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        saveJson('snapshot', { savedAt: new Date().toISOString(),
                               karts: data.karts, shop: data.shop || [],
                               quicks: data.quicks || [], inv: data.inv || {},
                               parts: data.parts || [] });
        var photoIndex = loadJson('photo_index', {});
        writeAllDates(ss, data);
        writeLog(ss, data, photoIndex);
        writePartsUsed(ss, data);
        writeKartTabs(ss, data);
        writeInventoryQty(ss, data.inv);
        writeNeeded(ss, data.inv);
        ensureOrderView(ss);
        return txt('ok');
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
    var payload = JSON.stringify({ ok: !!snap, karts: snap ? snap.karts : null,
                                   shop: snap ? (snap.shop || []) : [],
                                   quicks: snap ? (snap.quicks || []) : [],
                                   inv: snap ? (snap.inv || {}) : {},
                                   photos: photoIndex });
    return ContentService.createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  if (e && e.parameter && e.parameter.mode === 'inv') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload2 = JSON.stringify({ ok: true, names: invConfigFromSheet(ss),
                                    receipts: loadJson('receipts', []) });
    return ContentService.createTextOutput(cb + '(' + payload2 + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return txt('kart log sync is running');
}

function txt(s) { return ContentService.createTextOutput(s); }

/* ================= ONE-TIME CLEANUP (run manually once) ================= */
function cleanupImpl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var goners = ['EXAMPLE', 'ALL DATES', 'Restocks', 'needed', 'ordering',
                '_invState', 'parts used', 'full', '1', '35', '36'];
  var deleted = [];
  for (var g = 0; g < goners.length; g++) {
    var sh = ss.getSheetByName(goners[g]);
    if (sh) { ss.deleteSheet(sh); deleted.push(goners[g]); }
  }
  /* only move the front tabs — kart tabs already sit in order behind them */
  var front = ['APP ALL DATES', 'APP LOG', 'APP PARTS USED', 'APP PHOTOS',
               'APP NEEDED', 'APP ORDER', 'APP ORDERS', 'inventory'];
  var pos = 1;
  for (var j = 0; j < front.length; j++) {
    var sh2 = ss.getSheetByName(front[j]);
    if (!sh2) continue;
    ss.setActiveSheet(sh2, true);
    ss.moveActiveSheet(pos++);
  }
  ss.setActiveSheet(ss.getSheetByName('APP ALL DATES') || ss.getSheets()[0], true);
  try {
    SpreadsheetApp.getUi().alert('Cleanup done. Deleted this run: ' +
      (deleted.length ? deleted.join(', ') : 'nothing (already clean)') + '.');
  } catch (err) {}
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
  if (!sh) {
    sh = ss.insertSheet('APP PHOTOS');
    sh.getRange(1, 1, 1, 4).setValues([['KART', 'DATE', 'PHOTO ID', 'PHOTO']]);
    sh.setColumnWidth(4, 260);
  }
  return sh;
}
function savePhoto(data) {
  if (!data.id || !data.dataURL) throw 'photo missing id/dataURL';
  var index = loadJson('photo_index', {});
  if (index[data.id]) return;
  var base64 = String(data.dataURL).split(',')[1];
  if (!base64) throw 'bad dataURL';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = photoSheet(ss);
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 3).setValues([[data.kart || '?', data.date || '', data.id]]);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', data.id + '.jpg');
  var img = sh.insertImage(blob, 4, row);
  var w = img.getWidth(), h = img.getHeight();
  if (w > 240) { img.setHeight(Math.round(h * 240 / w)); img.setWidth(240); }
  sh.setRowHeight(row, Math.max(img.getHeight() + 10, 60));
  index[data.id] = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=D' + row;
  saveJson('photo_index', index);
}

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
function freshSheet(ss, name, textCols) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  if (textCols) for (var i = 0; i < textCols.length; i++)
    sh.getRange(textCols[i] + ':' + textCols[i]).setNumberFormat('@');
  return sh;
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
  sh.setFrozenRows(1);
}
var KART_HDR = ['DATE', 'ACTION COMPLETED', 'parts used', 'MECHANIC', 'NOTES',
                'CHAIN DATE', 'DIFF DATE', 'BRAKE FLUSH', 'BAT 1 DATE', 'BAT 2 DATE',
                'BAT 3 DATE', 'BAT 4 DATE', 'WELD COUNT'];
function writeKartTabs(ss, data) {
  var ks = kartOrder(data.karts);
  for (var i = 0; i < ks.length; i++) {
    var k = ks[i];
    var sh = ss.getSheetByName(k);
    if (!sh) sh = ss.insertSheet(k);
    var kd = data.karts[k], st = kd.status || {};
    var rows = [KART_HDR.slice()];
    rows.push(['', '', '', '', kd.knotes || '', st.chain || '', st.diff || '', st.brake || '',
               st.bat1 || '', st.bat2 || '', st.bat3 || '', st.bat4 || '', st.weld || 0]);
    var es = (kd.entries || []).slice().sort(function (a, b) {
      return (new Date(a.date) - new Date(b.date)) || 0;
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
  }
}
function photoCell(photos, index) {
  if (!photos) return '';
  if (typeof photos === 'number') return photos ? photos + ' in app' : '';
  if (!photos.length) return '';
  var locs = [], missing = 0;
  for (var i = 0; i < photos.length; i++) {
    var u = index[photos[i]];
    if (u) {
      var m = u.match(/range=D(\d+)/);
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
  all.sort(function (a, b) {
    var d = new Date(b[0]) - new Date(a[0]);   // newest first
    return d || ((+a[1]) - (+b[1]));
  });
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
    .addItem('Run sheet cleanup (delete retired tabs)', 'cleanupSheet')
    .addToUi();
}

/* ---- inventory names/thresholds served to the app + receipts ---- */
function invConfigFromSheet(ss) {
  var sh = ss.getSheetByName('inventory');
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 2, last - 1, 6).getValues(); // B..G: part, name, qty, inventoried, red, green
  for (var i = 0; i < vals.length; i++) {
    var num = vals[i][0];
    if (num === '' || num === null) continue;
    out[String(num).trim().toUpperCase()] = {
      n: String(vals[i][1] || num).trim(),
      r: parseInt(vals[i][4], 10) || 0,
      g: parseInt(vals[i][5], 10) || 0
    };
  }
  return out;
}

/* ---- write app quantities back into the inventory tab QUANTITY column ---- */
function writeInventoryQty(ss, inv) {
  var sh = ss.getSheetByName('inventory');
  if (!sh || !inv) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var nums = sh.getRange(2, 2, last - 1, 1).getValues();
  var qtys = sh.getRange(2, 4, last - 1, 1).getValues();
  for (var i = 0; i < nums.length; i++) {
    var key = String(nums[i][0] || '').trim().toUpperCase();
    if (key && inv[key] !== undefined) qtys[i][0] = inv[key];
  }
  sh.getRange(2, 4, last - 1, 1).setValues(qtys);
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
    orders.getRange(1, 1, 1, 7).setValues([['DATE ORDERED', 'PART #', 'NAME', 'QTY ORDERED', 'QTY RECEIVED', 'DATE BOOKED', 'ORDER #']])
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
      if (qty > 0) { newRows.push([today, String(vals[i][0]), vals[i][1], qty, '', '', orderNo]); clearRows.push(i + 2); }
      else skipped++;
    }
  }
  if (!newRows.length) {
    SpreadsheetApp.getUi().alert(skipped
      ? skipped + ' checked item(s) have no ORDER QTY — type how many you\'re ordering first.'
      : 'Nothing checked.');
    return;
  }
  orders.getRange(orders.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  for (var r = 0; r < clearRows.length; r++) needed.getRange(clearRows[r], 8).setValue(false);
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
  var vals = orders.getRange(2, 1, last - 1, 7).getValues();
  var receipts = loadJson('receipts', []);
  var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  var booked = 0;
  for (var i = 0; i < vals.length; i++) {
    var rec = parseInt(vals[i][4], 10);
    var alreadyBooked = vals[i][5] !== '' && vals[i][5] !== null;
    if (!alreadyBooked && !isNaN(rec) && rec > 0) {
      receipts.push({ id: 'rc' + Date.now() + '_' + i, part: String(vals[i][1]).toUpperCase(), qty: rec });
      orders.getRange(i + 2, 6).setValue(today);
      booked++;
    }
  }
  if (!booked) { SpreadsheetApp.getUi().alert('Nothing to book — fill QTY RECEIVED on rows that arrived (and aren\'t booked yet).'); return; }
  if (receipts.length > 400) receipts = receipts.slice(receipts.length - 400);
  saveJson('receipts', receipts);
  SpreadsheetApp.getUi().alert(booked + ' item(s) booked. The app adds them to stock next time it opens or syncs (then quantities flow back to the inventory tab).');
}
