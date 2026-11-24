// Downloads monthly NPCI UPI statistics into a raw JSON file ({ "YYYY-Mon": { s: status, t: body } }).
// NPCI's API rejects plain HTTP clients, so requests run inside a real (headless) Firefox page.
//
// Usage: node pull.js <tab> <out.json> [type]
//   node pull.js upi-apps    ../../data/npci_upi_apps_raw.json
//   node pull.js top-15-psps ../../data/npci_top15_payer_psps_raw.json payer
//   node pull.js credit-line ../../data/npci_credit_line_raw.json app-wise   (also: issuer-wise)
// Re-running resumes: months already saved with status 200 are skipped. The API returns 503 when hit
// too fast, so each month is retried with growing waits.
const { firefox } = require("playwright-core");
const fs = require("fs");

const [tab, outFile, type] = process.argv.slice(2);
if (!tab || !outFile) { console.error("usage: node pull.js <tab> <out.json> [type]"); process.exit(1); }
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const FIRST_YEAR = 2020, LAST = { y: 2026, m: "Oct" };

(async () => {
  const out = fs.existsSync(outFile) ? JSON.parse(fs.readFileSync(outFile)) : {};
  const browser = await firefox.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto("https://www.npci.org.in/product/ecosystem-statistics/upi", { waitUntil: "networkidle", timeout: 90000 });
  for (let y = FIRST_YEAR; y <= LAST.y; y++) {
    for (const m of MONTHS) {
      if (y === LAST.y && MONTHS.indexOf(m) > MONTHS.indexOf(LAST.m)) break;
      const key = `${y}-${m}`;
      if (out[key] && out[key].s === 200) continue;
      const qs = new URLSearchParams({ product_name: "UPI", tab_name: tab, year: y, month: m, page_no: 1, sort_by: "asc", size: 500, locale: "en" });
      if (type) qs.set("type_name", type);
      for (let attempt = 0; attempt < 4; attempt++) {
        await page.waitForTimeout(2000 + attempt * 4000);
        const res = await page.evaluate(async (q) => {
          const r = await fetch(`/api/ecosystem-statistics/get-statistics?${q}`);
          return { s: r.status, t: await r.text() };
        }, qs.toString());
        out[key] = res;
        if (res.s === 200) break;
      }
      console.log(key, out[key].s);
      fs.writeFileSync(outFile, JSON.stringify(out));
    }
  }
  await browser.close();
})();
