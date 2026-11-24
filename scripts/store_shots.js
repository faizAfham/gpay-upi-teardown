// Lists screenshot image URLs from Play Store + App Store listings. Prints JSON; downloads done separately.
const { firefox } = require('playwright-core');
const pages = {
  gpay_play: 'https://play.google.com/store/apps/details?id=com.google.android.apps.nbu.paisa.user&hl=en_IN&gl=IN',
  phonepe_play: 'https://play.google.com/store/apps/details?id=com.phonepe.app&hl=en_IN&gl=IN',
  gpay_ios: 'https://apps.apple.com/in/app/google-pay-save-pay-manage/id1193357041',
  phonepe_ios: 'https://apps.apple.com/in/app/phonepe-secure-payments-app/id1170055821',
};
(async () => {
  const b = await firefox.launch({ headless: true });
  const out = {};
  for (const [k, u] of Object.entries(pages)) {
    const p = await b.newPage();
    try {
      await p.goto(u, { waitUntil: 'networkidle', timeout: 60000 });
      out[k] = { url: p.url(), title: await p.title(), imgs: await p.$$eval('img, source', els => els.map(e => ({ src: e.currentSrc || e.src || e.srcset, alt: e.alt || '' })).filter(x => x.src && /screenshot/i.test(x.alt + ' ' + x.src) || /mzstatic.*(Screenshot|PurpleSource|png)/.test(x.src))) };
    } catch (e) { out[k] = { error: String(e) }; }
    await p.close();
  }
  console.log(JSON.stringify(out, null, 1));
  await b.close();
})();
