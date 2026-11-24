// Dumps visible text of each URL to <outdir>/<n>.txt. Usage: node page_text.js <outdir> <url>...
const { firefox } = require('playwright-core');
const fs = require('fs');
const [outdir, ...urls] = process.argv.slice(2);
fs.mkdirSync(outdir, { recursive: true });
(async () => {
  const b = await firefox.launch({ headless: true });
  for (const [i, u] of urls.entries()) {
    const p = await b.newPage();
    try {
      await p.goto(u, { waitUntil: 'domcontentloaded', timeout: 45000 });
      await p.waitForTimeout(2500);
      fs.writeFileSync(`${outdir}/${i}.txt`, u + '\n' + await p.innerText('body'));
      console.log(i, 'ok', u);
    } catch (e) { console.log(i, 'ERR', u, String(e).slice(0, 100)); }
    await p.close();
  }
  await b.close();
})();
