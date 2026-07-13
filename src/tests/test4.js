const { chromium } = require('playwright');
const http = require('http');
const { execSync } = require('child_process');

(async () => {
  // mock receiver: capture POSTs, serve JSONP snapshot on GET
  const posts = [];
  const server = http.createServer((req, res) => {
    if (req.method === 'POST') {
      let body = '';
      req.on('data', c => body += c);
      req.on('end', () => {
        posts.push(JSON.parse(body));
        res.writeHead(200); res.end('ok');
      });
    } else {
      const url = new URL(req.url, 'http://x');
      if (url.searchParams.get('mode') === 'snapshot') {
        const cb = url.searchParams.get('callback');
        const payload = {
          ok: true,
          karts: { '7': { status: { chain: '9/9/2026', diff: '', brake: '', bat1: '', bat2: '', bat3: '', bat4: '', weld: 5 },
                          knotes: 'RESTORED NOTE', entries: [
                            { date: '7/1/2026', action: 'server entry', parts: '', mechanic: 'ROBERT', notes: '', photos: ['pServer1'] } ] } },
          photos: { pServer1: 'https://drive.google.com/file/d/FAKE/view' }
        };
        res.writeHead(200, { 'Content-Type': 'application/javascript' });
        res.end(cb + '(' + JSON.stringify(payload) + ')');
      } else { res.writeHead(200); res.end('running'); }
    }
  });
  await new Promise(r => server.listen(8765, r));

  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const page = await (await browser.newContext({ viewport: { width: 768, height: 1024 } })).newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });

  await page.goto('file:///home/claude/kartlog/kart_log.html');
  await page.waitForTimeout(700);
  // point sync at mock server
  await page.evaluate(() => {
    DB.meta.syncUrl = 'http://127.0.0.1:8765/exec';
    localStorage.setItem('k1kartlog_v1', JSON.stringify(DB));
  });

  // log work with a photo on kart 2
  execSync(`python3 -c "from PIL import Image; Image.new('RGB',(1600,1200),(30,60,200)).save('/tmp/p1.jpg')"`);
  await page.$$eval('#gridAdult .kbtn', els => { els[1].click(); });
  await page.click('#btnLog');
  await page.fill('#fAction', 'test photo entry');
  await page.$$eval('#mechRow .chip', els => { els[0].click(); });
  await page.setInputFiles('#photoInput', '/tmp/p1.jpg');
  await page.waitForTimeout(1200);
  await page.click('#btnSaveLog');
  await page.waitForTimeout(2500); // save -> photo upload (600ms) + snapshot debounce not yet

  const photoPosts = posts.filter(p => p.type === 'photo');
  console.log('photo POSTs:', photoPosts.length,
    photoPosts[0] ? `id=${photoPosts[0].id} kart=${photoPosts[0].kart} dataURL len=${photoPosts[0].dataURL.length}` : '');
  await page.waitForTimeout(4500); // let snapshot debounce fire
  const snapPosts = posts.filter(p => p.type === 'snapshot');
  const snapEntry = snapPosts.length ? snapPosts[snapPosts.length-1].karts['2'].entries.slice(-1)[0] : null;
  console.log('snapshot POSTs:', snapPosts.length, '| last kart2 entry photos field:', JSON.stringify(snapEntry && snapEntry.photos));
  const upFlag = await page.evaluate(() => JSON.parse(localStorage.getItem('k1kartlog_v1')).meta.up);
  console.log('uploaded flags:', JSON.stringify(upFlag));

  // restore-from-sheet flow (accept confirm)
  page.on('dialog', d => d.accept());
  await page.click('#btnBack');
  await page.click('#btnExportScr');
  await page.click('#btnPullSheet');
  await page.waitForTimeout(1500);
  const restored = await page.evaluate(() => {
    const db = JSON.parse(localStorage.getItem('k1kartlog_v1'));
    return { k7chain: db.karts['7'].status.chain, k7note: db.karts['7'].knotes,
             k7entries: db.karts['7'].entries.length, purl: db.meta.photoUrls['pServer1'],
             k2entries: db.karts['2'].entries.length };
  });
  console.log('after restore:', JSON.stringify(restored));

  // drive-link chip renders for non-local photo
  await page.click('#btnBack2');
  await page.$$eval('#gridAdult .kbtn', els => { els[6].click(); }); // kart 7
  await page.waitForTimeout(600);
  const chip = await page.$$eval('#historyList .photorow button.thumb', els => els.length);
  console.log('drive-link photo chips on kart 7:', chip);

  console.log('ERRORS:', errors.length ? errors : 'none');
  await browser.close();
  server.close();
})();
