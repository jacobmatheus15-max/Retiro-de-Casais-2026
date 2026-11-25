const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 3 });
  await p.goto('file://' + __dirname + '/capa.html', { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.screenshot({ path: __dirname + '/capa.png' });
  await b.close();
})();
