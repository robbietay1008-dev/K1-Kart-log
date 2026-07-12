const { chromium } = require('playwright');
const fs = require('fs');

// make a small test JPEG via canvas in-page instead of uploading a real file?
// Simpler: generate a PNG buffer here (1x1 won't test compression path well, use 2000px to test resize)
const { execSync } = require('child_process');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const ctx = await browser.newContext({ viewport: { width: 768, height: 1024 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('CONSOLE: ' + m.text()); });

  await page.goto('file:///home/claude/kartlog/kart_log.html');
  await page.waitForTimeout(600);

  // no tire UI anywhere
  const tireMention = await page.evaluate(() => document.body.innerHTML.indexOf('TIRES') > -1);
  console.log('tire UI present (should be false):', tireMention);

  // open kart 3
  await page.$$eval('#gridAdult .kbtn', els => { els[2].click(); });
  await page.waitForTimeout(300);

  // open log form, fill, attach photo
  await page.click('#btnLog');
  await page.waitForTimeout(300);
  await page.fill('#fAction', 'WELD frame crack rear left');
  await page.$$eval('#mechRow .chip', els => { els[0].click(); });

  // create a big test image file then set on input
  execSync('python3 -c "from PIL import Image; im=Image.new(\'RGB\',(2400,1800),(200,30,30)); im.save(\'/tmp/testphoto.jpg\')" 2>/dev/null || python3 -c "import struct,zlib; import sys; sys.exit(1)"');
  await page.setInputFiles('#photoInput', '/tmp/testphoto.jpg');
  await page.waitForTimeout(1200);
  const pending = await page.$$eval('#formPhotos .thumb', els => els.length);
  console.log('pending photos in form:', pending);
  await page.screenshot({ path: 'shot_form2.png', fullPage: true });

  await page.click('#btnSaveLog');
  await page.waitForTimeout(800);
  const thumbs = await page.$$eval('#historyList .entry .thumb', els => els.length);
  console.log('thumbs in history:', thumbs);
  const weldCount = await page.$eval('#weldHint', el => el.textContent);
  console.log('weld hint (should be 1):', weldCount.slice(0, 20));

  // persistence of photo across reload
  await page.reload();
  await page.waitForTimeout(700);
  await page.$$eval('#gridAdult .kbtn', els => { els[2].click(); });
  await page.waitForTimeout(700);
  const thumbs2 = await page.$$eval('#historyList .entry .thumb', els => els.length);
  const thumbSrc = await page.$$eval('#historyList .entry .thumb', els => els.map(e => e.src.slice(0, 30)));
  console.log('thumbs after reload:', thumbs2, thumbSrc);

  // photo size check (compressed?)
  const size = await page.evaluate(() => new Promise(res => {
    photoGet(JSON.parse(localStorage.getItem('k1kartlog_v1')).karts['3'].entries.slice(-1)[0].photos[0],
      d => res(d ? d.length : -1));
  }));
  console.log('stored photo dataURL chars (should be ~100-300k, orig 2400px):', size);

  // viewer
  await page.$$eval('#historyList .entry .thumb', els => { els[0].click(); });
  await page.waitForTimeout(500);
  const viewerOpen = await page.$eval('#pviewer', el => el.className);
  console.log('viewer open:', viewerOpen);
  await page.screenshot({ path: 'shot_viewer.png' });
  await page.click('#pvClose');

  // export headers have no tires, notes has photo marker
  const row = await page.evaluate(() => {
    const wb = buildWorkbook();
    const ws = wb.Sheets['3'];
    const hdr = [];
    for (let c = 0; c < 20; c++) {
      const cell = ws[XLSX.utils.encode_cell({r:0,c:c})];
      if (cell) hdr.push(cell.v);
    }
    // find the weld row entry
    let marker = null;
    for (let r = 2; r < 15; r++) {
      const cell = ws[XLSX.utils.encode_cell({r:r,c:4})];
      if (cell && String(cell.v).indexOf('photo') > -1) marker = cell.v;
    }
    return { hdr, marker };
  });
  console.log('export headers:', row.hdr.join(' | '));
  console.log('photo marker in notes col:', row.marker);

  console.log('ERRORS:', errors.length ? errors : 'none');
  await browser.close();
})();
