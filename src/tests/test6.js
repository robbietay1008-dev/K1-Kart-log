/* Battery scan-in, end to end: the BATTERIES screen on two devices, the merge,
   and the BATTERY TRACKING tab drawn by the real logic.gs on a fake sheet. */
const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const vm = require('vm');

const ROOT = process.env.KL_ROOT || '/home/claude/kartlog';
const APP = 'file://' + ROOT + '/kart_log.html';
const LOGIC = ROOT + '/google-sheet-sync/logic.gs';
const PORT = 8767;

let fails = 0;
function ok(label, cond, extra) {
  if (cond) console.log('  ok   ' + label);
  else { fails++; console.log('  FAIL ' + label + (extra === undefined ? '' : '  ' + JSON.stringify(extra))); }
}

/* ---------------- fake sheet (formatting calls are no-ops that chain) ---------------- */
function makeSheet(rows, name) {
  const g = rows.map(r => { const c = r.slice(); while (c.length < 14) c.push(''); return c; });
  const fmt = {};
  ['setNumberFormat', 'setBackgrounds', 'setBackground', 'setFontWeight', 'setFontSize', 'setHorizontalAlignment',
   'setVerticalAlignment', 'setBorder', 'merge', 'insertCheckboxes'].forEach(k => fmt[k] = function () { return this; });
  return {
    _g: g,
    getName: () => name || 'inventory',
    getLastRow: () => { let n = 0; for (let i = 0; i < g.length; i++) if (g[i].some(v => v !== '' && v !== null)) n = i + 1; return n; },
    getLastColumn: () => 14, getMaxColumns: () => 14,
    insertColumnsAfter() {},
    insertRowsAfter(after, n) { for (let i = 0; i < n; i++) g.push(new Array(14).fill('')); },
    setColumnWidth() {}, setRowHeight() {}, setHiddenGridlines() {},
    deleteRow(r) { g.splice(r - 1, 1); },
    getRange(a, b, c, d) {
      if (typeof a === 'string') return Object.assign({ setValue: () => {}, getValue: () => '' }, fmt);
      const r0 = a, c0 = b, nr = c === undefined ? 1 : c, nc = d === undefined ? 1 : d;
      return Object.assign({
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = g[r0 - 1 + i] || [];
            const line = [];
            for (let j = 0; j < nc; j++) line.push(row[c0 - 1 + j] === undefined ? '' : row[c0 - 1 + j]);
            out.push(line);
          }
          return out;
        },
        getDisplayValues() { return this.getValues().map(r => r.map(String)); },
        getValue() { return this.getValues()[0][0]; },
        setValues(v) {
          for (let i = 0; i < v.length; i++) {
            while (g.length < r0 - 1 + i + 1) g.push(new Array(14).fill(''));
            for (let j = 0; j < v[i].length; j++) g[r0 - 1 + i][c0 - 1 + j] = v[i][j];
          }
          return this;
        },
        setValue(v) { return this.setValues([[v]]); },
        clearContent() { return this.setValues(new Array(nr).fill(0).map(() => new Array(nc).fill(''))); }
      }, fmt);
    },
    setFrozenRows() {}, clear() {},
    clearContents() { for (let i = 0; i < g.length; i++) for (let j = 0; j < g[i].length; j++) g[i][j] = ''; },
    getIndex: () => 3, getMaxRows: () => g.length, setName() {}
  };
}

const HDR = ['', 'PART #', 'NAME', 'QUANTITY', 'INVENTORIED', 'Red ≤', 'Green ≥'];
const tabs = { inventory: makeSheet([HDR, ['', '59191', 'optima battery', 22, '', 10, 15]]) };
const store = {};
const sandbox = {
  console,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: n => tabs[n] || null,
      getSheets: () => Object.keys(tabs).map(k => tabs[k]),
      insertSheet: n => (tabs[n] = makeSheet([new Array(14).fill('')], n)),
      deleteSheet: sh => { delete tabs[sh.getName()]; }, getSpreadsheetTimeZone: () => 'UTC'
    }),
    flush: () => {}, getUi: () => { throw new Error('no ui'); },
    BorderStyle: { SOLID_MEDIUM: 'M' }
  },
  LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
  ContentService: { createTextOutput: s => ({ setMimeType: () => s }), MimeType: {} },
  Utilities: { formatDate: () => '1/1/2026' },
  Date, JSON, Object, Math, String, Number, parseInt, parseFloat, isNaN, RegExp, Array, Error
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(LOGIC, 'utf8'), sandbox);
sandbox.__store = store;
vm.runInContext(`
  loadJson = function(n, f) { return __store[n] === undefined ? f : JSON.parse(JSON.stringify(__store[n])); };
  saveJson = function(n, o) { __store[n] = JSON.parse(JSON.stringify(o)); };
`, sandbox);

function receiveSnapshot(data) {
  sandbox.__d = data;
  vm.runInContext(`
    saveJson('snapshot', { karts: __d.karts || {}, shop: __d.shop || [], quicks: __d.quicks || null,
      inv: __d.inv || {}, invCfg: __d.invCfg || {}, tomb: __d.tomb || {}, stamps: __d.stamps || {},
      invTouched: __d.invTouched || {}, invCounted: __d.invCounted || {}, cfgTouched: __d.cfgTouched || {},
      partTomb: __d.partTomb || {}, rekeys: __d.rekeys || [], bat: __d.bat || {} });
    var __ss = SpreadsheetApp.getActiveSpreadsheet();
    mergeCfg(__ss, loadJson('snapshot', null));
    writeBatteryTab(__ss, loadJson('snapshot', null));
    writeBatteryLog(__ss, loadJson('snapshot', null));
  `, sandbox);
}
function snapAnswer() {
  const s = store.snapshot || {};
  return JSON.stringify({ ok: true, karts: s.karts || {}, shop: s.shop || [], inv: s.inv || {},
    invCfg: s.invCfg || {}, tomb: s.tomb || {}, stamps: s.stamps || {}, invTouched: s.invTouched || {},
    invCounted: s.invCounted || {}, cfgTouched: s.cfgTouched || {}, partTomb: s.partTomb || {},
    rekeys: s.rekeys || [], bat: s.bat || {}, photos: {} });
}
function invAnswer() {
  return vm.runInContext(`(function(){ var r = mergeCfg(SpreadsheetApp.getActiveSpreadsheet(), loadJson('snapshot', null));
    return JSON.stringify({ ok: true, names: r.names, del: r.del, cfgAt: r.at, receipts: [] }); })()`, sandbox);
}
function batTab(name) { const t = tabs[name]; return t ? t._g.map(r => r.slice(0, 5)) : null; }
function batTabNameJS(iso) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso); return 'BATTERIES ' + m[2] + '-' + m[3] + '-' + m[1]; }

const posts = [];
const server = http.createServer((req, res) => {
  if (req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      let d = null; try { d = JSON.parse(body); } catch (e) {}
      if (d && d.type === 'snapshot') { posts.push(d); try { receiveSnapshot(d); } catch (e) { console.log('  !! receiver threw: ' + e); fails++; } }
      res.writeHead(200); res.end('ok');
    });
    return;
  }
  const url = new URL(req.url, 'http://x');
  const mode = url.searchParams.get('mode'), cb = url.searchParams.get('callback');
  let payload = '{"ok":false}';
  try { if (mode === 'inv') payload = invAnswer(); else if (mode === 'snapshot') payload = snapAnswer(); }
  catch (e) { console.log('  !! ' + mode + ' threw: ' + e); fails++; }
  res.writeHead(200, { 'Content-Type': 'application/javascript' });
  res.end(cb ? cb + '(' + payload + ')' : payload);
});

const SYNC = 'http://127.0.0.1:' + PORT + '/exec';
async function device(browser, label) {
  const ctx = await browser.newContext({ viewport: { width: 768, height: 1024 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(label + ': ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(label + ' console: ' + m.text()); });
  page.on('dialog', d => d.accept());
  await page.goto(APP);
  await page.waitForFunction('typeof DB === "object" && DB && DB.karts');
  await page.evaluate(u => { DB.meta.syncUrl = u; saveQuiet(); }, SYNC);
  return { page, errors, ctx };
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
const push = d => d.page.evaluate(() => new Promise(r => { pushSnapshot(true); setTimeout(r, 400); }));
const pull = d => d.page.evaluate(() => new Promise(r => pullSnapshot(x => { if (mergeRemote(x)) { saveQuiet(); refreshActive(); } r(); })));
const toast = d => d.page.evaluate(() => $('toast').textContent);

(async () => {
  await new Promise(r => server.listen(PORT, '127.0.0.1', r));
  const shell = fs.readdirSync('/opt/pw-browsers').filter(n => n.startsWith('chromium_headless_shell-'))[0];
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/' + shell + '/chrome-linux/headless_shell' });
  const A = await device(browser, 'A');
  const B = await device(browser, 'B');

  /* ---------- 1. scan a pallet in on device A ---------- */
  console.log('\n1. scan a pallet in (device A)');
  await A.page.click('#btnBat');
  await sleep(150);
  ok('BATTERIES screen opens', await A.page.evaluate(() => $('scrBat').className.indexOf('active') > -1));
  ok('scan box has focus', await A.page.evaluate(() => document.activeElement === $('batScan')));
  ok('date received defaults to today', await A.page.evaluate(() => $('batDefRcv').value === todayISO()));
  await A.page.fill('#batDefRcv', '2026-09-19');
  await A.page.type('#batScan', '6180409549');
  await A.page.keyboard.press('Enter');
  let s = await A.page.evaluate(() => { const l = batSorted(); return { n: l.length, b: l[0] && l[0].b, box: $('batScan').value }; });
  ok('one battery on the 9/19 pallet, nothing filled in yet', s.n === 1 && s.b.sn === '6180409549' && s.b.rcv === '2026-09-19' && s.b.kart === '' && s.b.date === '' && s.b.ini === '', s);
  ok('scan box cleared for the next one', s.box === '');
  ok('row shows #1, the serial and "on the shelf"', /#1/.test(await A.page.textContent('#batList')) && /6180409549/.test(await A.page.textContent('#batList')) && /on the shelf/.test(await A.page.textContent('#batList')));

  await A.page.type('#batScan', '6180409549');
  await A.page.keyboard.press('Enter');
  ok('re-scanning the same serial is refused', await A.page.evaluate(() => batSorted().length) === 1 && (await toast(A)).indexOf('Already scanned') > -1, await toast(A));

  await A.page.type('#batScan', '6180405554\r');
  await A.page.keyboard.press('Tab');
  ok('Tab / stray CR handled', await A.page.evaluate(() => batSorted().length) === 2 && await A.page.evaluate(() => batSorted()[1].b.sn) === '6180405554');
  await A.page.type('#batScan', '123');
  await A.page.keyboard.press('Enter');
  ok('a 3-digit scan is rejected', await A.page.evaluate(() => batSorted().length) === 2);

  /* a second pallet a week later */
  await A.page.fill('#batDefRcv', '2026-09-26');
  await A.page.type('#batScan', '6180407777');
  await A.page.keyboard.press('Enter');
  s = await A.page.evaluate(() => ({ n: batSorted().length, num: batNumber(batSorted()[2].id), rcv: batSorted()[2].b.rcv }));
  ok('second pallet starts numbering at #1 again', s.n === 3 && s.num === 1 && s.rcv === '2026-09-26', s);

  /* ---------- 2. logging a battery into a kart pops the pairing box ---------- */
  console.log('\n2. log work that used a battery');
  await A.page.evaluate(() => openKart('12'));
  await A.page.click('#btnLog');
  await A.page.fill('#fDate', '2026-09-20');
  await A.page.fill('#fAction', 'Replace battery 3');
  await A.page.evaluate(() => { selectedParts.push({ num: '59191', qty: 1 }); renderSelParts(); selectedMech = 'ROBERT'; renderMechRow(); });
  await A.page.click('#btnSaveLog');
  await sleep(150);
  ok('log saved', await A.page.evaluate(() => DB.karts['12'].entries.some(e => e.action === 'Replace battery 3' && e.parts === '59191 x1')));
  ok('pairing popup opened for kart 12', await A.page.evaluate(() => $('batUseModal').className === 'modal open' && $('batUseTitle').textContent === 'Batteries into Kart 12'));
  ok('date used pre-filled from the log', await A.page.evaluate(() => $('batUseDate').value) === '2026-09-20');
  ok('all three shelf batteries listed', await A.page.evaluate(() => $('batUseList').children.length) === 3);
  await A.page.fill('#batUseIni', 'rb');
  await A.page.click('#btnBatUseSave');
  ok('saving with nothing picked is refused', await A.page.evaluate(() => $('batUseModal').className === 'modal open') && (await toast(A)).indexOf('Scan or tap') > -1);
  await A.page.type('#batUseScan', '9549');
  ok('typing part of a serial narrows the list', await A.page.evaluate(() => $('batUseList').children.length) === 1 && (await A.page.textContent('#batUseList')).indexOf('6180409549') > -1);
  await A.page.fill('#batUseScan', '');
  await A.page.type('#batUseScan', '6180409549');
  await A.page.keyboard.press('Enter');
  ok('scanned serial is ticked', await A.page.evaluate(() => Object.keys(batUse.sel).length === 1 && $('btnBatUseSave').textContent === 'SAVE (1)'));
  ok('four position chips for an adult kart', await A.page.evaluate(() => $('batUseList').querySelectorAll('.chip').length) === 4);
  await A.page.evaluate(() => { const c = $('batUseList').querySelectorAll('.chip'); c[1].click(); });
  ok('BAT 2 chip lit', await A.page.evaluate(() => batUse.pos[Object.keys(batUse.sel)[0]] === '2'));
  await A.page.click('#btnBatUseSave');
  s = await A.page.evaluate(() => batSorted()[0].b);
  ok('battery now shows kart 12 BAT 2 / 9/20/2026 / RB', s.kart === '12' && s.pos === '2' && s.date === '2026-09-20' && s.ini === 'RB', s);
  ok('history has the install', s.h.length === 1 && s.h[0].k === '12' && s.h[0].p === '2', s.h);
  ok('kart page BAT 2 chip shows the serial', await A.page.evaluate(() => $('batRow').children[1].textContent.indexOf('6180409549') > -1 && $('batRow').children[0].textContent.indexOf('6180409549') === -1));
  ok('popup closed', await A.page.evaluate(() => $('batUseModal').className === 'modal'));
  ok('kart 12 page lists the serial', (await A.page.textContent('#batSerials')).indexOf('6180409549') > -1);

  /* two batteries, one scanned that was never received, one tapped */
  await A.page.click('#btnLog');
  await A.page.fill('#fDate', '2026-09-21');
  await A.page.fill('#fAction', 'batteries');
  await A.page.evaluate(() => { selectedParts.push({ num: '059191', qty: 2 }); renderSelParts(); selectedMech = 'ROBERT'; renderMechRow(); });
  await A.page.click('#btnSaveLog');
  await sleep(150);
  ok('popup says 2 batteries logged', (await A.page.textContent('#batUseHint')).indexOf('2 batteries logged') > -1);
  ok('the one already in this kart is not offered as "on the shelf" first', (await A.page.textContent('#batUseList')).indexOf('already this kart') > -1);
  await A.page.type('#batUseScan', '6180400001');
  await A.page.keyboard.press('Enter');          /* dialog auto-accepted: adds it */
  ok('unknown serial added on the spot, received today', await A.page.evaluate(() => { const id = batBySn('6180400001'); return !!id && DB.bat[id].rcv === todayISO() && !!batUse.sel[id]; }));
  await A.page.evaluate(() => { const id = batBySn('6180405554'); const rows = $('batUseList').children; for (const r of rows) if (r.textContent.indexOf('6180405554') > -1) r.click(); });
  ok('tapped row ticked, save shows 2 of 2', await A.page.evaluate(() => $('btnBatUseSave').textContent === 'SAVE (2 of 2)'));
  ok('BAT 2 chip shows who is in it now', (await A.page.textContent('#batUseList')).indexOf('(9549)') > -1);
  /* 5554 -> BAT 2 (kicks 9549 out), 0001 -> BAT 1 */
  await A.page.evaluate(() => {
    const rows = Array.from($('batUseList').children);
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].textContent.indexOf('6180405554') > -1) rows[i + 1].querySelectorAll('.chip')[1].click();
    }
  });
  await A.page.evaluate(() => {
    const rows = Array.from($('batUseList').children);
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].textContent.indexOf('6180400001') > -1) rows[i + 1].querySelectorAll('.chip')[0].click();
    }
  });
  ok('positions chosen', await A.page.evaluate(() => batUse.pos[batBySn('6180405554')] === '2' && batUse.pos[batBySn('6180400001')] === '1'));
  await A.page.fill('#batUseIni', 'JS');
  await A.page.click('#btnBatUseSave');
  s = await A.page.evaluate(() => [batBySn('6180405554'), batBySn('6180400001')].map(id => DB.bat[id]));
  ok('both paired to kart 12 with JS / 9/21', s.every(b => b.kart === '12' && b.ini === 'JS' && b.date === '2026-09-21') && s[0].pos === '2' && s[1].pos === '1', s);
  s = await A.page.evaluate(() => DB.bat[batBySn('6180409549')]);
  ok('the old BAT 2 came out and its history says so', s.kart === '' && s.pos === '' && s.h.length === 2 && s.h[1].out === '12', s);
  ok('kart page: BAT 1 = 0001, BAT 2 = 5554', await A.page.evaluate(() => $('batRow').children[0].textContent.indexOf('6180400001') > -1 && $('batRow').children[1].textContent.indexOf('6180405554') > -1));

  /* a log with no battery does not pop */
  await A.page.click('#btnLog');
  await A.page.fill('#fDate', '2026-09-21');
  await A.page.fill('#fAction', 'replace chain');
  await A.page.evaluate(() => { selectedMech = 'ROBERT'; renderMechRow(); });
  await A.page.click('#btnSaveLog');
  await sleep(100);
  ok('no popup for a chain', await A.page.evaluate(() => $('batUseModal').className === 'modal'));

  /* skip path */
  await A.page.click('#btnLog');
  await A.page.fill('#fDate', '2026-09-22');
  await A.page.fill('#fAction', 'battery terminal');
  await A.page.evaluate(() => { selectedMech = 'ROBERT'; renderMechRow(); });
  await A.page.click('#btnSaveLog');
  await sleep(100);
  ok('"battery" in the action alone pops it', await A.page.evaluate(() => $('batUseModal').className === 'modal open'));
  await A.page.click('#btnBatUseSkip');
  ok('skip closes it', await A.page.evaluate(() => $('batUseModal').className === 'modal' && batUse === null));

  /* ---------- 3. the sheet draws one paper per pallet ---------- */
  console.log('\n3. BATTERIES tabs');
  await push(A);
  let t = batTab('BATTERIES 09-19-2026');
  ok('9/19 tab created', !!t);
  ok('9/26 tab created', !!batTab('BATTERIES 09-26-2026'));
  ok('the never-received one landed on a today tab', !!batTab(batTabNameJS(await A.page.evaluate(() => todayISO()))));
  ok('row 1 is Center / Date Received 9/19/2026', t[0][0] === 'Center:' && t[0][3] === 'Date Received:' && t[0][4] === '9/19/2026', t[0]);
  ok('row 2 is the title', t[1][0] === 'Battery Tracking Sheet');
  ok('row 3 headers match the paper', JSON.stringify(t[2]) === JSON.stringify(['Battery', 'Serial Number', 'Kart Number', 'Date Used', 'Initials']));
  ok('battery 1 (pulled again) shows blank kart but keeps its serial', t[3][0] === 1 && t[3][1] === '6180409549' && t[3][2] === '' , t[3]);
  { const lg = tabs['BATTERY LOG']._g.map(r => r.slice(0, 9));
    ok('BATTERY LOG has the header', lg[0][1] === 'SERIAL' && lg[0][5] === 'POSITION');
    const ev = lg.slice(1).filter(r => r[1] === '6180409549').map(r => r[3]);
    ok('9549 log: received, installed, pulled (newest first)', ev[0].indexOf('pulled from kart 12') === 0 && ev[1] === 'installed' && ev[2] === 'received', ev);
    const inst = lg.slice(1).find(r => r[1] === '6180405554' && r[3] === 'installed');
    ok('5554 install row carries kart 12 BAT 2 JS', inst && inst[4] === '12' && inst[5] === '2' && inst[7] === 'JS' && inst[8] === 'kart 12 BAT 2', inst); }
  ok('battery 2', t[4][0] === 2 && t[4][1] === '6180405554' && t[4][2] === '12' && t[4][3] === '9/21/2026' && t[4][4] === 'JS', t[4]);
  ok('30 numbered rows like the paper', t[32][0] === 30 && t[32][1] === '' && t.length >= 33, t.length);
  let t2 = batTab('BATTERIES 09-26-2026');
  ok('9/26 tab has its own #1 still on the shelf', t2[3][1] === '6180407777' && t2[3][2] === '' && t2[0][4] === '9/26/2026', t2[3]);

  /* Center typed on one tab is copied to all of them and survives */
  tabs['BATTERIES 09-19-2026']._g[0][1] = 'Arlington';
  await push(A);
  ok('Center kept on the 9/19 tab', batTab('BATTERIES 09-19-2026')[0][1] === 'Arlington');
  ok('Center copied to the 9/26 tab', batTab('BATTERIES 09-26-2026')[0][1] === 'Arlington');

  /* ---------- 4. second device ---------- */
  console.log('\n4. device B pulls, edits, and the two merge');
  await pull(B);
  s = await B.page.evaluate(() => batSorted().map(x => x.b.sn));
  ok('B sees all four', s.length === 4, s);
  await B.page.evaluate(() => { const id = batBySn('6180409549'); DB.bat[id].kart = '7'; DB.bat[id].at = Date.now() + 5; save(); });
  await push(B);
  await pull(A);
  ok('A takes B\'s newer kart number', await A.page.evaluate(() => DB.bat[batBySn('6180409549')].kart) === '7');
  await push(A);
  await pull(B);
  ok('no ping-pong', await B.page.evaluate(() => DB.bat[batBySn('6180409549')].kart) === '7');

  /* both devices scan the same new battery before syncing */
  await A.page.evaluate(() => { const now = Date.now() - 1000; DB.bat['bA'] = { sn: '6180400002', rcv: '2026-09-19', kart: '9', date: '', ini: '', c: now, at: now }; save(); });
  await B.page.evaluate(() => { const now = Date.now(); DB.bat['bB'] = { sn: '6180400002', rcv: '2026-09-19', kart: '', date: '', ini: 'WT', c: now, at: now }; save(); });
  await push(A); await pull(B); await push(B); await pull(A);
  s = await B.page.evaluate(() => ({ n: batSorted().filter(x => x.b.sn === '6180400002').length, ids: Object.keys(DB.bat).filter(k => DB.bat[k].sn === '6180400002'), rec: DB.bat['bA'] }));
  ok('duplicate serial collapses onto the earlier scan and keeps both sides\' details', s.n === 1 && s.ids[0] === 'bA' && s.rec.kart === '9' && s.rec.ini === 'WT' && s.rec.rcv === '2026-09-19', s);

  /* delete on A reaches B and the sheet; an emptied pallet loses its tab */
  await A.page.evaluate(() => { const id = batBySn('6180407777'); DB.tomb[id] = Date.now(); delete DB.bat[id]; save(); });
  await push(A); await pull(B);
  ok('deletion propagates to B', await B.page.evaluate(() => batBySn('6180407777')) === null);
  ok('the now-empty 9/26 tab is gone', !batTab('BATTERIES 09-26-2026'));
  { const tt = batTab('BATTERIES 09-19-2026'); const sns = tt.slice(3).map(r => r[1]).filter(Boolean);
    ok('9/19 tab keeps its batteries, 7777 gone, no gaps', sns.indexOf('6180400002') > -1 && sns.indexOf('6180407777') === -1 && tt[3 + sns.length][1] === '', sns); }

  /* ---------- 5. restore path ---------- */
  console.log('\n5. fresh device restores from the sheet');
  const C = await device(browser, 'C');
  await C.page.evaluate(() => { DB.bat = {}; saveQuiet(); });
  await pull(C);
  ok('restored device has the batteries', await C.page.evaluate(() => batSorted().length) === 4);

  /* ---------- 6. more than 30 on one pallet just adds rows ---------- */
  console.log('\n6. beyond the paper\'s 30 lines');
  await A.page.evaluate(() => { for (let i = 0; i < 40; i++) { const now = Date.now() + i; DB.bat['x' + i] = { sn: '61804' + String(10000 + i), rcv: '2026-09-19', kart: '', date: '', ini: '', c: now, at: now }; } save(); });
  await push(A);
  t = batTab('BATTERIES 09-19-2026');
  ok('43 rows numbered through', t[3 + 42][0] === 43 && t[3 + 42][1] !== '', t[3 + 42]);

  const errs = A.errors.concat(B.errors, C.errors);
  ok('no page errors on any device', errs.length === 0, errs);
  await browser.close();
  server.close();
  console.log(fails ? '\n' + fails + ' FAILED' : '\nALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.log('CRASH ' + e.stack); process.exit(2); });
