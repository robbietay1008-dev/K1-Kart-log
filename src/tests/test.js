const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const ctx = await browser.newContext({ viewport: { width: 768, height: 1024 } }); // iPad mini size
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('CONSOLE: ' + m.text()); });

  await page.goto('file:///home/claude/kartlog/kart_log.html');
  await page.waitForTimeout(600);
  await page.screenshot({ path: 'shot_home.png' });

  // kart grid counts
  const adult = await page.$$eval('#gridAdult .kbtn', els => els.length);
  const jr = await page.$$eval('#gridJr .kbtn', els => els.length);
  console.log('adult karts:', adult, 'jr karts:', jr);

  // open kart 2 (has seeded history)
  await page.$$eval('#gridAdult .kbtn', els => { els[1].click(); });
  await page.waitForTimeout(400);
  const hist = await page.$$eval('#historyList .entry', els => els.length);
  console.log('kart 2 history entries:', hist);
  const chain = await page.$eval('#svcRow', el => el.textContent);
  console.log('kart2 svc row:', chain);
  await page.screenshot({ path: 'shot_kart2.png', fullPage: true });

  // open log modal, add entry
  await page.click('#btnLog');
  await page.waitForTimeout(300);
  await page.$$eval('#quickRow .chip', els => { els[0].click(); }); // BRAKE FLUSH
  await page.fill('#fPartSearch', 'chain');
  await page.waitForTimeout(300);
  const partHits = await page.$$eval('#partResults .partitem', els => els.map(e => e.textContent));
  console.log('part search "chain" hits:', partHits.length, partHits.slice(0,3));
  await page.$$eval('#partResults .padd', els => { els[0].click(); });
  await page.$$eval('#mechRow .chip', els => { els[0].click(); }); // ROBERT
  await page.screenshot({ path: 'shot_form.png', fullPage: true });
  await page.click('#btnSaveLog');
  await page.waitForTimeout(400);
  const hist2 = await page.$$eval('#historyList .entry', els => els.length);
  console.log('kart 2 history after save:', hist2);
  const brakeAfter = await page.$eval('#svcRow', el => el.textContent);
  console.log('svc row after brake flush (should show today):', brakeAfter);

  // tire cycle
  await page.$$eval('#tireRow .tirebtn', els => { els[0].click(); });
  await page.waitForTimeout(200);
  const tireTxt = await page.$eval('#tireRow', el => el.textContent);
  console.log('tires after 1 tap on FL:', tireTxt);

  // verify persistence
  await page.reload();
  await page.waitForTimeout(500);
  await page.$$eval('#gridAdult .kbtn', els => { els[1].click(); });
  await page.waitForTimeout(300);
  const hist3 = await page.$$eval('#historyList .entry', els => els.length);
  console.log('kart 2 history after reload (persistence):', hist3);

  // build workbook in-page and dump to file to verify export
  const b64 = await page.evaluate(() => {
    const wb = buildWorkbook();
    return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
  });
  require('fs').writeFileSync('test_export.xlsx', Buffer.from(b64, 'base64'));
  console.log('export bytes:', Buffer.from(b64, 'base64').length);

  // junior kart check (no diff/brake chips, 2 batteries)
  await page.click('#btnBack');
  await page.$$eval('#gridJr .kbtn', els => { els[0].click(); }); // kart 41
  await page.waitForTimeout(300);
  const jrSvc = await page.$eval('#svcRow', el => el.textContent);
  const jrBats = await page.$$eval('#batRow .schip', els => els.length);
  console.log('kart41 svc chips:', jrSvc, '| batteries:', jrBats);

  console.log('ERRORS:', errors.length ? errors : 'none');
  await browser.close();
})();
