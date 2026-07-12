const { chromium } = require('playwright');
const http = require('http');

(async () => {
  // local server to capture the sync POST
  let captured = null;
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      captured = body;
      res.writeHead(200, { 'Access-Control-Allow-Origin': '*' });
      res.end('ok');
    });
  });
  await new Promise(r => server.listen(8765, r));

  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const ctx = await browser.newContext({ viewport: { width: 768, height: 1024 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('CONSOLE: ' + m.text()); });

  await page.goto('file:///home/claude/kartlog/kart_log.html');
  await page.waitForTimeout(600);

  // ---- kart notes ----
  await page.$$eval('#gridAdult .kbtn', els => { els[2].click(); }); // kart 3
  await page.waitForTimeout(300);
  const seedNote = await page.$eval('#knotesBox', el => el.textContent);
  console.log('kart 3 seeded note:', seedNote);
  await page.click('#knotesBox');
  await page.waitForTimeout(200);
  await page.fill('#knotesText', 'WATCH L BRAKE PINS. Pulls right after bumps. New diff 1/19.');
  await page.click('#btnKnotesSave');
  await page.waitForTimeout(300);
  const newNote = await page.$eval('#knotesBox', el => el.textContent);
  console.log('kart 3 edited note:', newNote);
  await page.screenshot({ path: 'shot_kart3_notes.png' });

  // ---- sync ----
  await page.click('#btnBack');
  await page.click('#btnExportScr');
  await page.fill('#syncUrl', 'http://127.0.0.1:8765/exec');
  await page.click('#btnSaveSync');
  await page.waitForTimeout(200);
  await page.click('#btnSyncNow');
  await page.waitForTimeout(1200);
  const status = await page.$eval('#syncStatus', el => el.textContent);
  console.log('sync status text:', status);
  if (captured) {
    const snap = JSON.parse(captured);
    console.log('captured snapshot: app =', snap.app, '| karts =', Object.keys(snap.karts).length,
      '| parts catalog =', snap.parts.length);
    console.log('kart 3 knotes in payload:', snap.karts['3'].knotes);
    console.log('kart 2 entries in payload:', snap.karts['2'].entries.length,
      '| first:', JSON.stringify(snap.karts['2'].entries[0]));
  } else {
    console.log('NO POST CAPTURED');
  }

  // ---- auto-sync on save (debounced 5s) ----
  captured = null;
  await page.click('#btnBack2');
  await page.$$eval('#gridAdult .kbtn', els => { els[0].click(); }); // kart 1
  await page.waitForTimeout(200);
  // edit a battery date (prompt)
  page.once('dialog', d => d.accept('7-26'));
  await page.$$eval('#batRow .schip', els => { els[0].click(); });
  await page.waitForTimeout(6500); // wait past debounce
  console.log('auto-sync fired after status edit:', captured ? 'YES' : 'NO');
  if (captured) console.log('kart1 bat1 in payload:', JSON.parse(captured).karts['1'].status.bat1);

  // ---- export workbook has PARTS USED + KART NOTES ----
  const exp = await page.evaluate(() => {
    const wb = buildWorkbook();
    const names = wb.SheetNames.slice(0, 4);
    const pu = XLSX.utils.sheet_to_json(wb.Sheets['PARTS USED'], { header: 1 }).slice(0, 4);
    const ad = XLSX.utils.sheet_to_json(wb.Sheets['ALL DATES'], { header: 1 });
    const k3row = ad.find(r => String(r[0]) === '3');
    const k3sheet = XLSX.utils.sheet_to_json(wb.Sheets['3'], { header: 1 });
    return { names, pu, k3row, k3status: k3sheet[1] };
  });
  console.log('sheet names:', exp.names);
  console.log('PARTS USED top rows:', JSON.stringify(exp.pu));
  console.log('ALL DATES kart3 row:', JSON.stringify(exp.k3row));
  console.log('kart3 tab status row (notes in col E):', JSON.stringify(exp.k3status));

  console.log('ERRORS:', errors.length ? errors : 'none');
  await browser.close();
  server.close();
})();
