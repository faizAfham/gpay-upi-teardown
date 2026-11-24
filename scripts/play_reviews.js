// Opens the Play Store "See all reviews" dialog, scrolls it, prints reviews (name|date|stars|text) as JSON lines.
// Usage: node play_reviews.js <appId> <scrolls>
const { firefox } = require('playwright-core');
const [id, scrolls = 30] = process.argv.slice(2);
(async () => {
  const b = await firefox.launch({ headless: true });
  const p = await b.newPage();
  await p.goto(`https://play.google.com/store/apps/details?id=${id}&hl=en_IN&gl=IN&showAllReviews=true`, { waitUntil: 'networkidle', timeout: 60000 });
  await p.getByRole('button', { name: /see all reviews/i }).first().click().catch(() => {});
  await p.waitForTimeout(3000);
  for (let i = 0; i < +scrolls; i++) {
    await p.evaluate(() => { const d = document.querySelector('[role=dialog] .fysCi, [role=dialog] [jsname] div[style], [role=dialog]'); const s = [...document.querySelectorAll('[role=dialog] *')].find(e => e.scrollHeight > e.clientHeight + 50); (s || d).scrollTop = 1e9; });
    await p.waitForTimeout(700);
  }
  const rows = await p.$$eval('[role=dialog] .RHo1pe', els => els.map(e => ({
    name: e.querySelector('.X5PpBb')?.innerText, date: e.querySelector('.bp9Aid')?.innerText,
    stars: e.querySelector('[role=img]')?.getAttribute('aria-label'), text: e.querySelector('.h3YV2d')?.innerText })));
  rows.forEach(r => console.log(JSON.stringify(r)));
  await b.close();
})();
