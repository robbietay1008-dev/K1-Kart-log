const { chromium } = require('playwright');
const http = require('http');

(async () => {
  // local server to capture the sync POST
  let captured = null;
  const server = http.createServer((req, res) => {
    if (req.method !== 'POST') {
      const cb = new URL(req.url, 'http://x').searchParams.get('callback');
      res.writeHead(200, { 'Content-Type': 'application/javascript' });
      res.end(cb ? cb + '({"ok":true})' : 'ok');
      return;
    }
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      if (body.indexOf('"type":"snapshot"') > -1) captured = body;
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
  await page.waitForTimeout(800);
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
  if (captured) {
    /* karts 1, 35 and 36 were retired on request, so read whichever is first */
    const ks = JSON.parse(captured).karts, first = Object.keys(ks)[0];
    console.log('kart ' + first + ' bat1 in payload:', ks[first].status.bat1);
  }

  // (Excel export was removed from the app on request — block dropped)

  console.log('ERRORS:', errors.length ? errors : 'none');
  await browser.close();
  server.close();
})();
