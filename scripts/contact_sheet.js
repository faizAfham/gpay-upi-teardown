// Renders a grid of local images to one PNG for visual review. Usage: node contact_sheet.js <dir> <prefix> <out.png>
const { firefox } = require('/Users/msahmad/.npm/_npx/705bc6b22212b352/node_modules/playwright-core');
const fs = require('fs'), path = require('path');
const [dir, prefix, out] = process.argv.slice(2);
const files = fs.readdirSync(dir).filter(f => f.startsWith(prefix) && f.endsWith('.png')).sort();
const html = '<body style="margin:0;display:flex;flex-wrap:wrap;gap:6px;font:14px sans-serif">' + files.map(f =>
  `<figure style="margin:0;width:260px"><img src="file://${path.resolve(dir, f)}" style="width:260px"><figcaption>${f}</figcaption></figure>`).join('') + '</body>';
(async () => {
  const b = await firefox.launch({ headless: true });
  const p = await b.newPage({ viewport: { width: 1100, height: 800 } });
  fs.writeFileSync(path.join(dir, '..', '_sheet.html'), html);
  await p.goto('file://' + path.resolve(dir, '..', '_sheet.html'));
  await p.waitForTimeout(1500);
  await p.screenshot({ path: out, fullPage: true });
  await b.close();
})();
