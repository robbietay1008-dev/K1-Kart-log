/* Two-way part editing, end to end.
   A real Chromium runs the built app; a local HTTP server stands in for the
   Apps Script receiver and runs the REAL logic.gs against a fake Sheets API,
   so a change made in the app is merged by the same code that runs live. */
const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const vm = require('vm');

const APP = 'file:///home/claude/kartlog/kart_log.html';
const LOGIC = '/home/claude/kartlog/ghrepo/google-sheet-sync/logic.gs';
const PORT = 8766;

let fails = 0;
function ok(label, cond, extra) {
  if (cond) console.log('  ok   ' + label);
  else { fails++; console.log('  FAIL ' + label + (extra === undefined ? '' : '  ' + JSON.stringify(extra))); }
}

/* ---------------- fake sheet ---------------- */
function makeSheet(rows, name) {
  const g = rows.map(r => { const c = r.slice(); while (c.length < 14) c.push(''); return c; });
  return {
    _g: g,
    getName: () => name || 'inventory',
    getLastRow: () => { let n = 0; for (let i = 0; i < g.length; i++) if (g[i].some(v => v !== '' && v !== null)) n = i + 1; return n; },
    getLastColumn: () => 14,
    getMaxColumns: () => 14,
    insertColumnsAfter() {},
    insertRowsAfter(after, n) { for (let i = 0; i < n; i++) g.push(new Array(14).fill('')); },
    setColumnWidth() {},
    deleteRow(r) { g.splice(r - 1, 1); },
    getRange(a, b, c, d) {
      if (typeof a === 'string') return { setNumberFormat: () => {}, setValue: () => {}, getValue: () => '' };
      const r0 = a, c0 = b, nr = c === undefined ? 1 : c, nc = d === undefined ? 1 : d;
      return {
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
            while (g.length < r0 - 1 + i + 1) g.push(['','','','','','','','','','','','','','']);
            for (let j = 0; j < v[i].length; j++) g[r0 - 1 + i][c0 - 1 + j] = v[i][j];
          }
          return this;
        },
        setValue(v) { return this.setValues([[v]]); },
        clearContent() { return this.setValues(new Array(nr).fill(0).map(() => new Array(nc).fill(''))); },
        setNumberFormat() { return this; },
        setBackgrounds() { return this; },
        setFontWeight() { return this; },
        insertCheckboxes() { return this; }
      };
    },
    setFrozenRows() {}, clear() {}, clearContents() {}, getIndex: () => 3,
    getMaxRows: () => g.length, setName() {}
  };
}

const HDR = ['', 'PART #', 'NAME', 'QUANTITY', 'INVENTORIED', 'Red ≤', 'Green ≥'];
const sheet = makeSheet([HDR,
  ['', '59191', 'optima battery', 22, '', 10, 15],
  ['', 'CHAIN219', 'chain 219', 8, '', 2, 6],
  ['', 'SEATBOLT', 'seat bolt', 40, '', 5, 12]]);
const store = {};
const tabs = { inventory: sheet };

const sandbox = {
  console,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: n => tabs[n] || null,
      getSheets: () => Object.keys(tabs).map(k => tabs[k]),
      insertSheet: n => (tabs[n] = makeSheet([new Array(14).fill('')], n)),
      deleteSheet: () => {},
      getSpreadsheetTimeZone: () => 'UTC'
    }),
    flush: () => {}, getUi: () => { throw new Error('no ui'); }
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
  /* the same two things handlePost does for part identity */
  vm.runInContext(`
    saveJson('snapshot', { karts: __d.karts || {}, shop: __d.shop || [], quicks: __d.quicks || null,
      inv: __d.inv || {}, invCfg: __d.invCfg || {}, tomb: __d.tomb || {}, stamps: __d.stamps || {},
      invTouched: __d.invTouched || {}, cfgTouched: __d.cfgTouched || {},
      partTomb: __d.partTomb || {}, rekeys: __d.rekeys || [] });
    var __ss = SpreadsheetApp.getActiveSpreadsheet();
    mergeCfg(__ss, loadJson('snapshot', null));
    writeInventoryQty(__ss, __d.inv || {}, __d.invCfg || {});
  `, sandbox);
}
function invAnswer() {
  return vm.runInContext(`
    (function(){ var ss = SpreadsheetApp.getActiveSpreadsheet();
      var snap = loadJson('snapshot', null);
      var r = mergeCfg(ss, snap);
      scanCounts(ss, snap && snap.inv);
      return JSON.stringify({ ok: true, names: r.names, del: r.del, cfgAt: r.at,
                              receipts: loadJson('receipts', []) }); })()
  `, sandbox);
}
function snapAnswer() {
  const s = store.snapshot || {};
  return JSON.stringify({ ok: true, karts: s.karts || {}, shop: s.shop || [], inv: s.inv || {},
    invCfg: s.invCfg || {}, tomb: s.tomb || {}, stamps: s.stamps || {},
    invTouched: s.invTouched || {}, cfgTouched: s.cfgTouched || {},
    partTomb: s.partTomb || {}, rekeys: s.rekeys || [], photos: {} });
}
function rowsOf() {
  const m = {};
  for (let i = 1; i < sheet._g.length; i++)
    if (sheet._g[i][1]) m[String(sheet._g[i][1])] = { n: sheet._g[i][2], q: sheet._g[i][3], r: sheet._g[i][5],
                                                      g: sheet._g[i][6], k: sheet._g[i][13] };
  return m;
}

const posts = [];
const server = http.createServer((req, res) => {
  if (req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      let d = null;
      try { d = JSON.parse(body); } catch (e) {}
      if (d && d.type === 'snapshot') { posts.push(d); try { receiveSnapshot(d); } catch (e) { console.log('  !! receiver threw: ' + e); fails++; } }
      res.writeHead(200); res.end('ok');
    });
    return;
  }
  const url = new URL(req.url, 'http://x');
  const mode = url.searchParams.get('mode'), cb = url.searchParams.get('callback');
  let payload = '{"ok":false}';
  try {
    if (mode === 'inv') payload = invAnswer();
    else if (mode === 'snapshot') payload = snapAnswer();
  } catch (e) { console.log('  !! ' + mode + ' threw: ' + e); fails++; }
  res.writeHead(200, { 'Content-Type': 'application/javascript' });
  res.end(cb ? cb + '(' + payload + ')' : payload);
});

/* ---------------- browser side ---------------- */
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

(async () => {
  await new Promise(r => server.listen(PORT, '127.0.0.1', r));
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const A = await device(browser, 'A');
  const B = await device(browser, 'B');

  /* give both devices a log entry that uses CHAIN219, so a renumber has history to rewrite */
  for (const d of [A, B]) {
    await d.page.evaluate(() => {
      DB.karts['7'].entries.unshift({ id: 'tst1', date: '7/1/2026', action: 'REPLACE CHAIN',
        parts: 'CHAIN219 x2, SEATBOLT', mechanic: 'ROBERT', notes: '', photos: [] });
      /* the app knows these two parts already, like anything on the real sheet */
      DB.invCfg['CHAIN219'] = { n: 'chain 219', r: 2, g: 6 }; DB.inv['CHAIN219'] = 8;
      DB.invCfg['SEATBOLT'] = { n: 'seat bolt', r: 5, g: 12 }; DB.inv['SEATBOLT'] = 40;
      epBust(); saveQuiet();
    });
  }

  /* ---------- 1. add a part through the real UI ---------- */
  console.log('\n1. add a part in the app (device A, real UI)');
  await A.page.evaluate(() => { renderInv(); showScreen('scrInv'); });
  await A.page.click('#btnInvAdd');
  await A.page.fill('#naNum', 'tie999');
  await A.page.fill('#naName', 'tie rod end');
  await A.page.fill('#naQty', '4');
  await A.page.fill('#naRed', '2');
  await A.page.fill('#naGreen', '6');
  await A.page.click('#btnNaSave');
  let s = await A.page.evaluate(() => ({ cfg: DB.invCfg['TIE999'], qty: DB.inv['TIE999'],
    inCat: !!partRow('TIE999'), modal: $('invAddModal').className }));
  ok('part is in invCfg, upper-cased', !!s.cfg && s.cfg.n === 'tie rod end' && s.cfg.r === 2 && s.cfg.g === 6, s.cfg);
  ok('opening quantity kept', s.qty === 4, s.qty);
  ok('pickable from the catalog straight away', s.inCat);
  ok('modal closed', s.modal === 'modal');

  /* ---------- 2. rename + renumber through the real UI ---------- */
  console.log('\n2. rename and renumber (device A, real UI)');
  await A.page.evaluate(() => openInvModal('CHAIN219'));
  await A.page.click('#btnInvEdit');
  s = await A.page.evaluate(() => ({ open: $('invEditWrap').style.display,
    num: $('invENum').value, name: $('invEName').value, uses: partUsageCount('CHAIN219') }));
  ok('edit panel opens prefilled', s.open === 'block' && s.num === 'CHAIN219', s);
  ok('usage count sees the log entry', s.uses === 1, s.uses);
  await A.page.fill('#invEName', 'chain 219 heavy');
  await A.page.fill('#invENum', 'chn219h');
  await A.page.fill('#invERed', '3');
  await A.page.click('#btnInvEditSave');
  s = await A.page.evaluate(() => ({
    entry: DB.karts['7'].entries[0].parts,
    newCfg: DB.invCfg['CHN219H'], oldCfg: DB.invCfg['CHAIN219'],
    qty: DB.inv['CHN219H'], tomb: !!DB.partTomb['CHAIN219'],
    rekey: DB.rekeys[DB.rekeys.length - 1], oldPick: !!partRow('CHAIN219'), newPick: !!partRow('CHN219H')
  }));
  ok('history rewritten to the new number', s.entry === 'CHN219H x2, SEATBOLT', s.entry);
  ok('config carried across', s.newCfg && s.newCfg.n === 'chain 219 heavy' && s.newCfg.r === 3, s.newCfg);
  ok('old number gone from config', !s.oldCfg);
  ok('stock carried across', s.qty === 8, s.qty);
  ok('old number tombstoned', s.tomb);
  ok('rekey recorded for other devices', s.rekey && s.rekey.from === 'CHAIN219' && s.rekey.to === 'CHN219H', s.rekey);
  ok('old number no longer pickable', !s.oldPick);
  ok('new number pickable', s.newPick);

  /* ---------- 3. delete a part ---------- */
  console.log('\n3. delete a part (device A)');
  await A.page.evaluate(() => { openInvModal('SEATBOLT'); deleteInvPart(); });
  s = await A.page.evaluate(() => ({ tomb: !!DB.partTomb['SEATBOLT'], cfg: !!DB.invCfg['SEATBOLT'],
    pick: !!partRow('SEATBOLT'), entry: DB.karts['7'].entries[0].parts }));
  ok('tombstoned', s.tomb);
  ok('out of config', !s.cfg);
  ok('not pickable', !s.pick);
  ok('past log entry left alone', s.entry === 'CHN219H x2, SEATBOLT', s.entry);

  /* ---------- 4. sync A -> the sheet ---------- */
  console.log('\n4. device A syncs to the sheet');
  await A.page.evaluate(() => syncNow(true));
  await sleep(1200);
  let r = rowsOf();
  ok('receiver got the snapshot', posts.length >= 1, posts.length);
  ok('added part is a row on the sheet', r.TIE999 && r.TIE999.n === 'tie rod end' && r.TIE999.q === 4, r.TIE999);
  ok('renumbered row moved, not duplicated', !!r.CHN219H && !r.CHAIN219, Object.keys(r));
  ok('renumbered row keeps name, stock and threshold', r.CHN219H.n === 'chain 219 heavy' && r.CHN219H.q === 8 && r.CHN219H.r === 3, r.CHN219H);
  ok('deleted part pulled from the sheet', !r.SEATBOLT, Object.keys(r));
  ok('untouched row survived', r['59191'] && r['59191'].n === 'optima battery', r['59191']);

  /* ---------- 5. device B catches up ---------- */
  console.log('\n5. device B pulls and replays');
  await B.page.evaluate(() => syncNow(true));
  await sleep(1500);
  s = await B.page.evaluate(() => ({
    entry: DB.karts['7'].entries[0].parts,
    newCfg: DB.invCfg['CHN219H'], oldCfg: DB.invCfg['CHAIN219'],
    added: DB.invCfg['TIE999'], seat: !!DB.invCfg['SEATBOLT'], seatTomb: !!DB.partTomb['SEATBOLT'],
    rekeys: DB.rekeys.length, pick: !!partRow('CHN219H')
  }));
  ok('B rewrote its own history', s.entry === 'CHN219H x2, SEATBOLT', s.entry);
  ok('B has the new number', !!s.newCfg && s.newCfg.n === 'chain 219 heavy', s.newCfg);
  ok('B dropped the old number', !s.oldCfg);
  ok('B picked up the added part', !!s.added && s.added.n === 'tie rod end', s.added);
  ok('B removed the deleted part', !s.seat && s.seatTomb);
  ok('B can pick the new number', s.pick);
  ok('B recorded the rekey once', s.rekeys === 1, s.rekeys);

  /* B syncs back — nothing should ping-pong */
  const before = JSON.stringify(rowsOf());
  await B.page.evaluate(() => syncNow(true));
  await sleep(1200);
  ok('B syncing back changes nothing on the sheet', JSON.stringify(rowsOf()) === before, rowsOf());
  s = await B.page.evaluate(() => DB.rekeys.length);
  ok('rekey still not replayed', s === 1, s);

  /* ---------- 6. edits typed on the sheet ---------- */
  console.log('\n6. edits typed on the sheet come back to the app');
  for (let i = 1; i < sheet._g.length; i++) {
    if (String(sheet._g[i][1]) === '59191') { sheet._g[i][2] = 'OPTIMA BATTERY YELLOW'; sheet._g[i][5] = 8; }
  }
  for (let i = sheet._g.length - 1; i >= 1; i--) if (String(sheet._g[i][1]) === 'TIE999') sheet._g.splice(i, 1);
  await B.page.evaluate(() => pullInvConfig());
  await sleep(1200);
  s = await B.page.evaluate(() => ({ bat: DB.invCfg['59191'], tie: !!DB.invCfg['TIE999'],
    tieTomb: !!DB.partTomb['TIE999'], tiePick: !!partRow('TIE999') }));
  ok('sheet rename reached the app', s.bat && s.bat.n === 'OPTIMA BATTERY YELLOW' && s.bat.r === 8, s.bat);
  ok('hand-deleted row removed the part in the app', !s.tie && s.tieTomb);
  ok('and it is no longer pickable', !s.tiePick);
  await B.page.evaluate(() => syncNow(true));
  await sleep(1200);
  ok('deleted row stays deleted after B syncs', !rowsOf().TIE999, Object.keys(rowsOf()));

  /* ---------- 7. an unsent app edit is not clobbered ---------- */
  console.log('\n7. an app edit the sheet has not seen yet wins');
  await B.page.evaluate(() => {
    DB.invCfg['59191'].n = 'OPTIMA BATTERY RED TOP';
    DB.cfgTouched['59191'] = Date.now();   /* newer than anything the sheet knows */
    saveQuiet();
  });
  await B.page.evaluate(() => pullInvConfig());
  await sleep(1200);
  s = await B.page.evaluate(() => DB.invCfg['59191'].n);
  ok('local edit survives the pull', s === 'OPTIMA BATTERY RED TOP', s);
  await B.page.evaluate(() => syncNow(true));
  await sleep(1200);
  ok('and then reaches the sheet', rowsOf()['59191'].n === 'OPTIMA BATTERY RED TOP', rowsOf()['59191']);

  /* ---------- 8. adult / jr / both ---------- */
  console.log('\n8. FITS — adult, jr or both');
  /* let A catch up first so this section only changes the FITS field */
  await A.page.evaluate(() => pullInvConfig());
  await sleep(1200);

  await A.page.evaluate(() => { renderInv(); showScreen('scrInv'); });
  await A.page.click('#btnInvAdd');
  await A.page.fill('#naNum', 'jrtire1');
  await A.page.fill('#naName', 'jr rear tire');
  await A.page.fill('#naQty', '12');
  await A.page.evaluate(() => {
    var chips = $('naFitsRow').querySelectorAll('.chip');
    for (var i = 0; i < chips.length; i++) if (chips[i].textContent === 'JR') chips[i].click();
  });
  s = await A.page.evaluate(() => {
    var on = $('naFitsRow').querySelectorAll('.chip.on');
    return on.length === 1 ? on[0].textContent : '(' + on.length + ')';
  });
  ok('JR chip selects on its own', s === 'JR', s);
  await A.page.click('#btnNaSave');
  s = await A.page.evaluate(() => ({ cfg: DB.invCfg['JRTIRE1'], row: partRow('JRTIRE1') }));
  ok('kind stored on the new part', s.cfg && s.cfg.k === 'J', s.cfg);
  ok('kind reaches the catalog', s.row && s.row[4] === 'J', s.row);

  /* mark an existing part adult-only through the edit panel */
  await A.page.evaluate(() => openInvModal('59191'));
  await A.page.click('#btnInvEdit');
  s = await A.page.evaluate(() => {
    var on = $('invEFitsRow').querySelectorAll('.chip.on');
    return on.length === 1 ? on[0].textContent : '(' + on.length + ')';
  });
  ok('edit panel starts on BOTH', s === 'BOTH', s);
  await A.page.evaluate(() => {
    var chips = $('invEFitsRow').querySelectorAll('.chip');
    for (var i = 0; i < chips.length; i++) if (chips[i].textContent === 'ADULT') chips[i].click();
  });
  await A.page.click('#btnInvEditSave');
  s = await A.page.evaluate(() => ({ k: DB.invCfg['59191'].k, n: DB.invCfg['59191'].n }));
  ok('adult-only saved', s.k === 'A', s);
  ok('name untouched by the FITS save', !!s.n, s);

  /* the part search on the log form only offers what fits the kart */
  s = await A.page.evaluate(() => {
    var out = {}, keep = currentKart;
    currentKart = '7';   searchParts('JRTIRE1'); out.jrOnAdult = $('partResults').textContent;
    currentKart = '45';  searchParts('JRTIRE1'); out.jrOnJr = $('partResults').textContent;
    currentKart = '7';   searchParts('59191');   out.adOnAdult = $('partResults').textContent;
    currentKart = '45';  searchParts('59191');   out.adOnJr = $('partResults').textContent;
    currentKart = '7';   searchParts('CHN219H'); out.bothOnAdult = $('partResults').textContent;
    currentKart = '45';  searchParts('CHN219H'); out.bothOnJr = $('partResults').textContent;
    currentKart = keep;
    searchShopParts('JRTIRE1'); out.shop = $('sPartResults').textContent;
    return out;
  });
  ok('jr part hidden on an adult kart', s.jrOnAdult.indexOf('jr rear tire') === -1, s.jrOnAdult);
  ok('jr part offered on a jr kart', s.jrOnJr.indexOf('jr rear tire') > -1, s.jrOnJr);
  ok('adult part offered on an adult kart', s.adOnAdult.indexOf('ADULT') > -1, s.adOnAdult);
  ok('adult part hidden on a jr kart', s.adOnJr.indexOf('ADULT') === -1, s.adOnJr);
  ok('a both part shows on an adult kart', s.bothOnAdult.indexOf('chain 219 heavy') > -1, s.bothOnAdult);
  ok('a both part shows on a jr kart', s.bothOnJr.indexOf('chain 219 heavy') > -1, s.bothOnJr);
  ok('shop parts are never filtered', s.shop.indexOf('jr rear tire') > -1, s.shop);

  /* it reaches the sheet's FITS column */
  await A.page.evaluate(() => syncNow(true));
  await sleep(1200);
  r = rowsOf();
  ok('FITS header written', sheet._g[0][13] === 'FITS', sheet._g[0][13]);
  ok('new jr part lands as JR', r.JRTIRE1 && r.JRTIRE1.k === 'JR', r.JRTIRE1);
  ok('edited part lands as ADULT', r['59191'] && r['59191'].k === 'ADULT', r['59191']);
  ok('a both part is left blank', (r.CHN219H.k || '') === '', r.CHN219H);

  /* typed on the sheet, it comes back to the app */
  for (let i = 1; i < sheet._g.length; i++)
    if (String(sheet._g[i][1]) === 'CHN219H') sheet._g[i][13] = 'adult';
  await B.page.evaluate(() => pullInvConfig());
  await sleep(1200);
  s = await B.page.evaluate(() => ({ chn: DB.invCfg['CHN219H'], jr: DB.invCfg['JRTIRE1'] }));
  ok('B picked up the sheet-typed ADULT', s.chn && s.chn.k === 'A', s.chn);
  ok('B picked up the app-set JR', s.jr && s.jr.k === 'J', s.jr);

  /* and a second round trip changes nothing */
  const beforeFits = JSON.stringify(rowsOf());
  await B.page.evaluate(() => syncNow(true));
  await sleep(1200);
  {
    const a = JSON.parse(beforeFits), b2 = rowsOf(), diff = {};
    for (const k in b2) if (JSON.stringify(a[k]) !== JSON.stringify(b2[k])) diff[k] = [a[k], b2[k]];
    for (const k in a) if (!(k in b2)) diff[k] = [a[k], null];
    ok('FITS does not ping-pong', Object.keys(diff).length === 0, diff);
  }

  /* ---------- 8b. NOT KART: stock that never goes on a kart ---------- */
  console.log('\n8b. FITS — NOT KART');
  await A.page.evaluate(() => { renderInv(); showScreen('scrInv'); });
  await A.page.click('#btnInvAdd');
  await A.page.fill('#naNum', 'shoprag1');
  await A.page.fill('#naName', 'shop rags box');
  await A.page.fill('#naQty', '6');
  await A.page.evaluate(() => {
    var chips = $('naFitsRow').querySelectorAll('.chip');
    for (var i = 0; i < chips.length; i++) if (chips[i].textContent === 'NOT KART') chips[i].click();
  });
  await A.page.click('#btnNaSave');
  s = await A.page.evaluate(() => ({ cfg: DB.invCfg['SHOPRAG1'], row: partRow('SHOPRAG1') }));
  ok('NOT KART stored on the part', s.cfg && s.cfg.k === 'N', s.cfg);
  ok('NOT KART reaches the catalog', s.row && s.row[4] === 'N', s.row);

  s = await A.page.evaluate(() => {
    var out = {}, keep = currentKart, keepScope = qpScope;
    currentKart = '7';  searchParts('SHOPRAG1'); out.onAdult = $('partResults').textContent;
    currentKart = '45'; searchParts('SHOPRAG1'); out.onJr = $('partResults').textContent;
    currentKart = null; searchParts('SHOPRAG1'); out.noKart = $('partResults').textContent;
    currentKart = keep;
    qpScope = 'all';    searchQpParts('SHOPRAG1'); out.qpAll = $('qpPartResults').textContent;
    qpScope = keepScope;
    searchShopParts('SHOPRAG1'); out.shop = $('sPartResults').textContent;
    return out;
  });
  ok('NOT KART hidden on an adult kart', s.onAdult.indexOf('shop rags') === -1, s.onAdult);
  ok('NOT KART hidden on a jr kart', s.onJr.indexOf('shop rags') === -1, s.onJr);
  ok('NOT KART hidden with no kart picked', s.noKart.indexOf('shop rags') === -1, s.noKart);
  ok('NOT KART out of every quick pick', s.qpAll.indexOf('shop rags') === -1, s.qpAll);
  ok('NOT KART still usable in the shop', s.shop.indexOf('shop rags') > -1, s.shop);
  ok('NOT KART badged in the shop list', s.shop.indexOf('NOT KART') > -1, s.shop);

  await A.page.evaluate(() => syncNow(true));
  await sleep(1200);
  r = rowsOf();
  ok('NOT KART lands on the sheet as a word', r.SHOPRAG1 && r.SHOPRAG1.k === 'NOT KART', r.SHOPRAG1);

  /* typed by hand on the sheet it comes back too */
  for (let i = 1; i < sheet._g.length; i++)
    if (String(sheet._g[i][1]) === '59012') sheet._g[i][13] = 'not kart';
  await B.page.evaluate(() => pullInvConfig());
  await sleep(1200);
  s = await B.page.evaluate(() => ({ c: DB.invCfg['59012'], rag: DB.invCfg['SHOPRAG1'] }));
  ok('B picked up the sheet-typed NOT KART', s.c && s.c.k === 'N', s.c);
  ok('B picked up the app-set NOT KART', s.rag && s.rag.k === 'N', s.rag);

  const beforeNk = JSON.stringify(rowsOf());
  await B.page.evaluate(() => syncNow(true));
  await sleep(1200);
  {
    const a = JSON.parse(beforeNk), b2 = rowsOf(), diff = {};
    for (const k in b2) if (JSON.stringify(a[k]) !== JSON.stringify(b2[k])) diff[k] = [a[k], b2[k]];
    for (const k in a) if (!(k in b2)) diff[k] = [a[k], null];
    ok('NOT KART does not ping-pong', Object.keys(diff).length === 0, diff);
  }

  /* ---------- 8c. paper count sheet fills the inventory in ---------- */
  console.log('\n8c. APP COUNT SHEET');
  const post = d => vm.runInContext('writeCountRows(SpreadsheetApp.getActiveSpreadsheet(), __c)',
    Object.assign(sandbox, { __c: d }));

  /* make sure the sheet is holding the app's current numbers before we count */
  await A.page.evaluate(() => syncNow(true));
  await sleep(1200);
  const beforeQty = await A.page.evaluate(() => DB.inv['59191']);
  post({ rows: [
    { p: 1, l: 267, num: '059191' },        /* leading zero, as the paper prints it */
    { p: 1, l: 268, num: 'SHOPRAG1' },
    { p: 1, l: 269, num: 'NOSUCHPART' }
  ] });
  const cg = tabs['APP COUNT SHEET']._g;
  ok('count tab exists with the transcription', cg[1][2] === '059191' && cg[2][2] === 'SHOPRAG1', cg.slice(1, 4));
  ok('names come off the inventory tab', !!cg[1][3] && cg[1][3] === rowsOf()['59191'].n, cg[1]);
  ok('a part we do not stock is flagged', cg[3][7] === 'not in inventory', cg[3]);

  /* he walks the shelves and writes the numbers in */
  cg[1][5] = 17;
  cg[2][5] = 0;
  await A.page.evaluate(() => pullInvConfig());
  await sleep(1500);
  s = await A.page.evaluate(() => ({ b: DB.inv['59191'], c: DB.inv['SHOPRAG1'],
                                     cb: DB.invCounted['59191'], cc: DB.invCounted['SHOPRAG1'] }));
  ok('the counted number replaces the stock', s.b === 17 && beforeQty !== 17, [beforeQty, s.b]);
  ok('a count of zero empties the shelf', s.c === 0, s.c);
  ok('counted parts get ticked off', s.cb === 1 && s.cc === 1, s);
  ok('the sheet records what it applied', cg[1][6] === 17 && /counted/.test(String(cg[1][7])), cg[1]);

  /* and it goes back the other way on the next sync */
  await A.page.evaluate(() => syncNow(true));
  await sleep(1200);
  r = rowsOf();
  ok('the inventory tab agrees', r['59191'] && r['59191'].q === 17, r['59191']);

  /* pulling again must not double-apply */
  await A.page.evaluate(() => pullInvConfig());
  await sleep(1200);
  s = await A.page.evaluate(() => DB.inv['59191']);
  ok('a second pull changes nothing', s === 17, s);

  /* ---------- 9. no script errors anywhere ---------- */
  console.log('\n9. runtime');
  const errs = A.errors.concat(B.errors);
  ok('no page errors', errs.length === 0, errs.slice(0, 4));

  await browser.close();
  server.close();
  console.log(fails ? '\n' + fails + ' FAILED' : '\nall passed');
  process.exit(fails ? 1 : 0);
})();
