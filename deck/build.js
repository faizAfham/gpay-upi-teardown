// Builds gpay-teardown.pptx. Run: NODE_PATH=<dir with pptxgenjs> node build.js
// Edit text/numbers here; colours live in THEME only.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT = path.join(__dirname, "..", "gpay-teardown.pptx");
const CSV = path.join(__dirname, "..", "data", "npci_upi_apps_monthly.csv");

// Wada #17: Eugenia Red (GPay) + Sea Green (PhonePe); greys for everything else
const THEME = {
  name: "Wada 17",
  headFontFace: "Cambria",
  bodyFontFace: "Arial",
  colors: {
    dk1: "222222", lt1: "FFFFFF", dk2: "666666", lt2: "F3F3F3",
    accent1: "DA525D", accent2: "00B49B", accent3: "8C8C8C",
    accent4: "BDBDBD", accent5: "4D4D4D", accent6: "DDDDDD",
    hlink: "00B49B", folHlink: "8C8C8C",
  },
};
const HEX = { red: "DA525D", green: "00B49B", grey: "8C8C8C", light: "BDBDBD", dark: "222222", mid: "666666", panel: "F3F3F3", line: "DDDDDD" };

const W = 13.333, M = 0.5; // LAYOUT_WIDE, 0.5" margins
const SECTIONS = ["SUMMARY", "STORY", "WHY", "WHO", "OPTIONS", "PLAN", "RISKS"];

// ---------- data ----------
function loadShares() {
  const [head, ...rows] = fs.readFileSync(CSV, "utf8").trim().split("\n");
  const cols = head.split(",");
  const idx = (c) => cols.indexOf(c);
  const series = { "Google Pay": [], PhonePe: [], Paytm: [], Navi: [], Others: [] };
  const labels = [];
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  for (const r of rows) {
    const v = r.split(",");
    const [y, m] = v[0].split("-");
    labels.push(m === "01" ? y : (labels.length === 0 ? `${MON[+m - 1]} ${y}` : ""));
    let named = 0;
    for (const app of ["Google Pay", "PhonePe", "Paytm", "Navi"]) {
      const p = +v[idx(`${app}_pct`)];
      series[app].push(p);
      named += p;
    }
    series.Others.push(+(100 - named).toFixed(2));
  }
  return { labels, series };
}

// ---------- shared frame ----------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Google Pay India teardown";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.55, w: W - 2 * M, h: 0.9, fontFace: THEME.headFontFace, fontSize: 22, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "source", type: "body", x: M, y: 6.92, w: W - 2 * M - 0.6, h: 0.4, fontSize: 8, color: C.text2, valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.4, y: 6.92, w: 0.4, h: 0.25, fontSize: 8, color: C.text2, align: "right" },
});

function tracker(slide, active) {
  const runs = [];
  SECTIONS.forEach((s, i) => {
    runs.push({ text: s, options: { bold: s === active, color: s === active ? C.accent1 : C.text2 } });
    if (i < SECTIONS.length - 1) runs.push({ text: "   ·   ", options: { color: C.accent4 } });
  });
  slide.addText(runs, { x: M, y: 0.22, w: 9, h: 0.25, fontSize: 9, margin: 0, isTextBox: true, objectName: "tracker" });
}

function soWhat(slide, text, y = 6.2) {
  slide.addShape(pres.shapes.RECTANGLE, { x: M, y, w: W - 2 * M, h: 0.55, fill: { color: C.background2 }, line: { type: "none" }, objectName: "so-what box" });
  slide.addText([
    { text: "So what  ", options: { bold: true, color: C.accent1 } },
    { text, options: { color: C.text1 } },
  ], { x: M + 0.15, y, w: W - 2 * M - 0.3, h: 0.55, fontSize: 12, valign: "middle", margin: 0, isTextBox: true, objectName: "so-what text" });
}

function panelTitle(slide, text, x, y, w) {
  slide.addText(text, { x, y, w, h: 0.3, fontSize: 11, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: `panel: ${text}` });
  slide.addShape(pres.shapes.LINE, { x, y: y + 0.32, w, h: 0, line: { color: HEX.line, width: 0.75 } });
}

// ---------- slide 2: the story ----------
function slideStory() {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Teardown" });
  tracker(s, "STORY");
  s.addText("GPay lost the lead in 2020-21. Since 2024, new apps have taken share from both leaders", { placeholder: "title" });

  const { labels, series } = loadShares();
  const n = labels.length;

  // Chart geometry: plot area pinned so the phase strip and labels line up with it
  const cx = M, cy = 1.95, cw = 8.2, ch = 3.55;
  const PL = { x: 0.06, y: 0.04, w: 0.84, h: 0.84 }; // plot area as a fraction of the chart frame
  const plotX = cx + PL.x * cw, plotW = PL.w * cw, plotY = cy + PL.y * ch, plotH = PL.h * ch;
  const VMAX = 55;
  const xAt = (i) => plotX + ((i + 0.5) / n) * plotW;
  const yAt = (v) => plotY + (1 - v / VMAX) * plotH;

  panelTitle(s, "Share of UPI payments by app (%)", cx, 1.55, cw);

  const order = ["Others", "Navi", "Paytm", "PhonePe", "Google Pay"];
  const colors = { "Google Pay": HEX.red, PhonePe: HEX.green, Paytm: HEX.grey, Navi: "4D4D4D", Others: HEX.light };
  s.addChart(pres.charts.LINE, order.map((k) => ({ name: k, labels, values: series[k] })), {
    x: cx, y: cy, w: cw, h: ch, objectName: "share chart",
    layout: PL,
    chartColors: order.map((k) => colors[k]),
    lineSize: 2, lineDataSymbol: "none",
    showLegend: false,
    valAxisMinVal: 0, valAxisMaxVal: VMAX, valAxisMajorUnit: 10,
    valAxisLabelFontSize: 9, catAxisLabelFontSize: 9,
    valAxisLabelFontFace: "+mn-lt", catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: HEX.mid, catAxisLabelColor: HEX.mid,
    catAxisLabelFrequency: 1, catAxisLabelRotate: 0, catAxisLineShow: false, valAxisLineShow: false,
    valGridLine: { color: "EEEEEE", size: 0.5 }, catGridLine: { style: "none" },
  });

  // Direct end labels (colour-blind safe: no legend to decode)
  const last = n - 1;
  const endLabels = [
    ["Google Pay", `GPay ${series["Google Pay"][last].toFixed(1)}%`, 0],
    ["PhonePe", `PhonePe ${series.PhonePe[last].toFixed(1)}%`, 0],
    ["Paytm", `Paytm ${series.Paytm[last].toFixed(1)}%`, -0.09],
    ["Navi", `Navi ${series.Navi[last].toFixed(1)}%`, 0.09],
    ["Others", `Others ${series.Others[last].toFixed(1)}%`, -0.2],
  ];
  for (const [k, t, dy] of endLabels) {
    s.addText(t, { x: plotX + plotW + 0.05, y: yAt(series[k][last]) - 0.11 + dy, w: 1.2, h: 0.22, fontSize: 9, bold: k === "Google Pay" || k === "PhonePe", color: colors[k] === HEX.light ? HEX.grey : colors[k], margin: 0, isTextBox: true, objectName: `label ${k}` });
  }

  // Event markers
  const events = [
    [6, "Dec 2020: PhonePe overtakes. GPay's volume falls for 3 months."],
    [43, "Jan 2024: RBI acts on Paytm"],
  ];
  for (const [i, t] of events) {
    s.addShape(pres.shapes.LINE, { x: xAt(i), y: plotY, w: 0, h: plotH, line: { color: HEX.mid, width: 0.75, dashType: "dash" }, objectName: `event line ${t}` });
    s.addText(t, { x: xAt(i) + 0.05, y: plotY, w: i < 20 ? 2.9 : 2.3, h: i < 20 ? 0.34 : 0.2, fontSize: 8, color: HEX.mid, margin: 0, isTextBox: true, objectName: `event ${t}` });
  }

  // Phase strip under the x-axis, aligned to the plot area
  const phases = [[0, 12, "1  Overtake"], [12, 43, "2  Stable duopoly"], [43, 48, "3"], [48, n - 1, "4  New entrants"]];
  const sy = cy + ch + 0.02;
  phases.forEach(([a, b, t], k) => {
    const x0 = plotX + (a / (n - 1)) * plotW, x1 = plotX + (b / (n - 1)) * plotW;
    s.addShape(pres.shapes.RECTANGLE, { x: x0, y: sy, w: x1 - x0 - 0.02, h: 0.28, fill: { color: k === 0 ? "F8DCDF" : C.background2 }, line: { type: "none" }, objectName: `phase ${k + 1}` });
    s.addText(t, { x: x0, y: sy, w: x1 - x0 - 0.02, h: 0.28, fontSize: 9, bold: k === 0, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `phase ${k + 1} label` });
  });

  // Right column: phase table
  const rx = 9.0, rw = W - M - rx;
  panelTitle(s, "Share change by phase (points)", rx, 1.55, rw);
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.text2, fontSize: 9 } });
  const cell = (t, o = {}) => ({ text: t, options: { fontSize: 10, ...o } });
  const neg = { color: C.accent1, bold: true };
  s.addTable([
    [hdr("Phase"), hdr("PhonePe"), hdr("GPay"), hdr("Others")],
    [cell("1 Overtake\nJun 20-Jun 21", { bold: true }), cell("+8.3", { bold: true }), cell("-10.3", neg), cell("Paytm +3.0, Amazon Pay -2.7")],
    [cell("2 Stable\nJun 21-Dec 23"), cell("+1.3"), cell("+2.0"), cell("CRED +0.7, Amazon Pay -1.3")],
    [cell("3 Paytm crisis\nJan-Jun 24"), cell("+2.8"), cell("+1.4"), cell("Paytm -4.6")],
    [cell("4 New entrants\nJun 24-Oct 26"), cell("-3.9"), cell("-5.4", neg), cell("Navi +4.8, super.money +1.7")],
  ], {
    x: rx, y: 1.95, w: rw, colW: [1.25, 0.7, 0.6, rw - 2.55], fontFace: "Arial", color: C.text1,
    border: { type: "solid", pt: 0.5, color: HEX.line }, margin: 0.04, valign: "middle", objectName: "phase table",
  });

  // Right column: how our thinking changed
  const ty = 4.35;
  panelTitle(s, "How our thinking changed", rx, ty, rw);
  s.addText([
    { text: "We assumed: ", options: { bold: true } },
    { text: "GPay lost share to Navi. PhonePe stayed flat.", options: { breakLine: true } },
    { text: "We checked: ", options: { bold: true } },
    { text: "77 months of NPCI data.", options: { breakLine: true } },
    { text: "We found: ", options: { bold: true, color: C.accent1 } },
    { text: "PhonePe fell too (49.3% → 45.3%). GPay lost the lead in 2020-21, mostly to PhonePe." },
  ], { x: rx, y: ty + 0.42, w: rw, h: 1.3, fontSize: 10, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "thinking changed" });

  soWhat(s, "The lead was lost in 2020-21. That's where to look for the cause.");

  s.addText("Source: NPCI app-wise UPI data, Jun 2020-Oct 2026.", { placeholder: "source" });

  s.addNotes("GPay led UPI in mid-2020 with 44% of payments. PhonePe passed it in December 2020. GPay's volume even fell for three months, from 960 million to 828 million. The press says PhonePe stayed flat since 2024, but NPCI data shows it fell too. So the lead was lost in 2020-21; the newer decline is a cashback fight that hits both.");
  return s;
}

// ---------- helpers for the remaining slides ----------
const T = (slide, text, o) => slide.addText(text, { fontSize: 10, color: C.text1, margin: 0, valign: "top", isTextBox: true, ...o });
const box = (slide, o) => slide.addShape(pres.shapes.RECTANGLE, { line: { type: "none" }, ...o });
const hdr = (t) => ({ text: t, options: { bold: true, color: C.text2, fontSize: 9 } });
const cell = (t, o = {}) => ({ text: t, options: { fontSize: 10, ...o } });
const TABLE = { fontFace: "Arial", color: C.text1, border: { type: "solid", pt: 0.5, color: HEX.line }, margin: 0.05, valign: "middle" };
const b = (t, o = {}) => ({ text: t, options: { bold: true, ...o } });
const r = (t, o = {}) => ({ text: t, options: o });
const br = (t, o = {}) => ({ text: t, options: { breakLine: true, ...o } });

function frame(section, title, source, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section === "BACKUP" ? "Backup" : "Teardown" });
  if (section === "BACKUP") T(s, [b("BACKUP", { color: C.accent1 }), r("")], { x: M, y: 0.22, w: 9, h: 0.25, fontSize: 9, objectName: "tracker" });
  else tracker(s, section);
  s.addText(title, { placeholder: "title" });
  s.addText(source, { placeholder: "source" });
  s.addNotes(notes);
  return s;
}

// strength meter: n of 4 filled squares
function meter(s, x, y, n) {
  for (let i = 0; i < 4; i++) box(s, { x: x + i * 0.17, y, w: 0.13, h: 0.13, fill: { color: i < n ? HEX.red : HEX.line }, objectName: `meter ${i}` });
}

// ---------- slide 0: cover ----------
function slideCover() {
  const s = pres.addSlide({ sectionTitle: "Teardown" });
  box(s, { x: 9.0, y: 0, w: W - 9.0, h: 7.5, fill: { color: C.accent1 }, objectName: "cover panel" });
  T(s, "PRODUCT TEARDOWN", { x: M + 0.3, y: 2.3, w: 7.5, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2 });
  T(s, "Why did Google Pay lose India's UPI lead to PhonePe, and how does it win it back?", { x: M + 0.3, y: 2.75, w: 7.9, h: 2.0, fontFace: THEME.headFontFace, fontSize: 36, color: C.text1 });
  T(s, "Faiz  ·  November 2026", { x: M + 0.3, y: 5.0, w: 7.9, h: 0.3, fontSize: 12, color: C.text2 });
  s.addNotes("An outside-in teardown of Google Pay in India, built on NPCI's monthly app data. One question: why did GPay lose the lead, and how does it win it back?");
}

// ---------- slide 1: executive summary ----------
function slideSummary() {
  const s = frame("SUMMARY",
    "GPay grew 13.7× but lost 12 points of share. Two moves can win it back",
    "Source: NPCI app-wise UPI data, Jun 2020-Oct 2026.",
    "UPI grew almost 19 times in six years. GPay had 44% in mid-2020 and lost the lead to PhonePe by December. PhonePe went from 37% to 45% over the same period. The 3.1 billion figure only sizes the gap; the 30% cap means GPay could never hold 44% anyway. The plan: track weekly merchant payers, launch habit rewards, then bring credit to the pay button.");
  const rows = [
    ["Situation", "UPI payments grew 18.9× in six years. Merchants pay no fee, so apps compete on habit and earn from credit."],
    ["Complication", "GPay lost the lead to PhonePe in Dec 2020. Its share fell from 44.3% to 32.1%. Since 2024, cashback apps have taken share from both leaders."],
    ["Resolution", "Track weekly merchant payers. Launch habit rewards this quarter. Bring credit to the pay button within a year."],
  ];
  rows.forEach(([k, v], i) => {
    const y = 1.65 + i * 1.47;
    box(s, { x: M, y, w: 7.9, h: 1.32, fill: { color: i === 2 ? "FBE9EB" : C.background2 }, objectName: `scr ${k}` });
    T(s, k, { x: M + 0.2, y, w: 1.6, h: 1.32, valign: "middle", fontSize: 14, bold: true, color: i === 2 ? C.accent1 : C.text1 });
    T(s, v, { x: M + 1.9, y, w: 5.85, h: 1.32, valign: "middle", fontSize: 15 });
  });
  const tiles = [
    ["44.3% → 32.1%", "GPay's UPI share, 2020 → 2026"],
    ["13.7× vs 18.9×", "GPay's growth vs the market's"],
    ["~3.1B a month", "Payments a month GPay would have at its 2020 share"],
  ];
  tiles.forEach(([n, l], i) => {
    const y = 1.65 + i * 1.47;
    T(s, n, { x: 8.75, y, w: 4.08, h: 0.6, fontFace: THEME.headFontFace, fontSize: 28, color: C.accent1 });
    T(s, l, { x: 8.75, y: y + 0.65, w: 4.08, h: 0.6, fontSize: 12, color: C.text2 });
  });
  soWhat(s, "Both moves grow habit and revenue, not share. So the 30% cap can't block them.");
}

// ---------- slide 3: why PhonePe won ----------
function slideWhy() {
  const s = frame("WHY",
    "PhonePe won the shop counter while GPay pulled back on rewards",
    "Sources: Entrackr (May 2020); KrASIA (May 2021); NPCI (app-wise volume; payer banks, Jan 2022); PhonePe DRHP; Team-BHP; Business Standard (Feb 2024); Inc42/Entrackr; Outlook Business (Oct 2025).",
    "Four things line up with the 2020-21 overtake. PhonePe already led shop payments in April 2020. GPay cut cashback, and its volume fell for three months. PhonePe's users were mostly in smaller towns. Reliability may have helped too, but only one month of data shows it. Amazon Pay is the warning: a huge shopping app, and still 0.4% of UPI.");
  const lw = 8.3;
  panelTitle(s, "Four reasons, with 2020-21 evidence", M, 1.55, lw);
  const reasons = [
    ["1", "Shop-payment habit", "Apr 2020: PhonePe led shop payments (180M vs 164M), while GPay led overall. Today: ~54% vs ~26% of active merchants (est.).", 2, "Medium"],
    ["2", "GPay pulled back rewards", "GPay cut cashback before the overtake (analyst, May 2021). Its volume then fell 960M → 828M. No official date.", 3, "Medium-strong"],
    ["3", "Tier-2/3 reach", "80% of PhonePe's users were in tier-2 towns or smaller (2021). No GPay data.", 2, "Medium"],
    ["4", "Reliability, 2020-21", "Jan 2022: GPay's banks failed 30-50× more often. Only one month of data. Level since late 2023.", 1, "Weak; past"],
  ];
  const rows = [[hdr("#"), hdr("Reason"), hdr("Evidence"), hdr("Strength")]];
  reasons.forEach(([n, k, e, , lab]) => rows.push([cell(n, { bold: true, color: C.accent1 }), cell(k, { bold: true }), cell(e, { fontSize: 9.5 }), cell(`\n${lab}`, { fontSize: 9, color: C.text2 })]));
  const colW = [0.3, 1.55, 5.15, 1.3];
  s.addTable(rows, { ...TABLE, x: M, y: 1.95, w: lw, colW, rowH: [0.28, 0.72, 0.72, 0.55, 0.6], objectName: "reasons table" });
  // meters sit in the strength column, above the label
  let y = 1.95 + 0.28;
  [0.72, 0.72, 0.55, 0.6].forEach((h, i) => { meter(s, M + 7.0 + 0.08, y + 0.12, reasons[i][3]); y += h; });

  // quotes
  const qy = 5.0;
  T(s, [r("“GPay was absolutely unreliable … The switch to PhonePe is not for any kind of rewards but for reliability.”", { italic: true }), r("  Team-BHP user, Jan 2023", { color: C.text2, fontSize: 8.5 })], { x: M, y: qy, w: 4.05, h: 0.95, fontSize: 10 });
  T(s, [r("“As a trader, my concern is the money, not the medium … we will use whatever the customer uses.”", { italic: true }), r("  Delhi shop owner, Feb 2024", { color: C.text2, fontSize: 8.5 })], { x: M + 4.25, y: qy, w: 4.05, h: 0.95, fontSize: 10 });

  // Amazon Pay highlight
  const hx = 9.05, hw = W - M - hx;
  box(s, { x: hx, y: 1.55, w: hw, h: 4.45, fill: { color: "FBE9EB" }, objectName: "amazon highlight" });
  T(s, "AMAZON PAY", { x: hx + 0.2, y: 1.7, w: hw - 0.4, h: 0.25, fontSize: 9, bold: true, color: C.accent1 });
  T(s, "5.93% → 0.40%", { x: hx + 0.2, y: 1.98, w: hw - 0.4, h: 0.5, fontFace: THEME.headFontFace, fontSize: 26, color: C.text1 });
  T(s, "UPI share, Aug 2020 peak → Oct 2026", { x: hx + 0.2, y: 2.5, w: hw - 0.4, h: 0.22, fontSize: 9, color: C.text2 });
  T(s, [
    br("Joined UPI late (2019)", { bullet: true }),
    br("Stuck at ~70M payments a month", { bullet: true }),
    br("Lost ~₹9,550 Cr (FY20-26)", { bullet: true }),
    r("CEO, Oct 2025: “Our QR codes are deliberately focused on large enterprise merchants”", { bullet: true }),
  ], { x: hx + 0.2, y: 2.85, w: hw - 0.4, h: 1.95, fontSize: 10, paraSpaceAfter: 4 });
  T(s, [b("Lesson: "), r("a big shopping app doesn't create a daily payment habit.")], { x: hx + 0.2, y: 4.95, w: hw - 0.4, h: 0.95, fontSize: 10 });

  soWhat(s, "GPay lost the shop-counter habit. That's what it must rebuild.");
}

// ---------- slide 4: who + north star ----------
function slideWho() {
  const s = frame("WHO",
    "UPI grows because people pay more often. GPay should track weekly merchant payers",
    "Sources: PhonePe DRHP; NPCI; student survey (n=398). GPay's segments use PhonePe's as a proxy.",
    "Payments grew 19 times; users only 2 to 3 times. So the fight is for how often people pay, and most people keep two or more apps. The metric is weekly merchant payers: people who pay a shop in 3 of the last 4 weeks. It looks like a low bar next to the average of about 28 shop payments a month, but that average is pulled up by power users. The exact threshold should come from GPay's own retention data. PhonePe's filing calls 30-plus payments a month 'habitual'; I call those power users.");
  // left: segments
  const lx = M, lw = 4.1;
  panelTitle(s, "Segments (PhonePe's split as proxy)", lx, 1.55, lw);
  s.addTable([
    [hdr("Segment"), hdr("Pays/month"), hdr("Size")],
    [cell("Power", { bold: true }), cell("≥30"), cell("41% of active users", { fontSize: 9.5 })],
    [cell("Habitual", { bold: true, color: C.accent1 }), cell("10-29"), cell("59% combined", { fontSize: 9.5, color: C.accent1, bold: true })],
    [cell("Casual", { bold: true, color: C.accent1 }), cell("1-9"), cell("Target", { fontSize: 9.5, color: C.accent1 })],
    [cell("Dormant", { bold: true }), cell("0"), cell("~100-200M signed up, inactive (est.)", { fontSize: 9.5 })],
  ], { ...TABLE, x: lx, y: 1.95, w: lw, colW: [0.95, 0.85, 2.3], objectName: "segments" });

  T(s, [b("Who to target: ", { color: C.accent1 }), r("habitual and casual users. Cashback apps are pulling them away. Power users already pay daily.")], { x: lx, y: 4.15, w: lw, h: 1.1, fontSize: 12 });

  // middle: frequency facts
  const mx = 4.85, mw = 3.55;
  panelTitle(s, "Growth came from frequency", mx, 1.55, mw);
  const facts = [
    ["25.5 → 38.3", "payments per PhonePe user a month, 2023 → 2025"],
    ["19× vs 2-3×", "growth in payments vs growth in users"],
    ["1.8 apps", "per person (student survey, n=398)"],
  ];
  facts.forEach(([n, l], i) => {
    const y = 2.0 + i * 1.05;
    T(s, n, { x: mx, y, w: mw, h: 0.42, fontFace: THEME.headFontFace, fontSize: 22, color: C.text1 });
    T(s, l, { x: mx, y: y + 0.44, w: mw, h: 0.5, fontSize: 9.5, color: C.text2 });
  });

  // right: NSM tree
  const rx = 8.75, rw = W - M - rx;
  panelTitle(s, "North-star metric and its inputs", rx, 1.55, rw);
  box(s, { x: rx, y: 2.0, w: rw, h: 0.85, fill: { color: C.accent1 }, objectName: "nsm" });
  T(s, [br("Weekly merchant payers", { bold: true, fontSize: 13 }), r("GPay users who pay a shop in ≥3 of the last 4 weeks", { fontSize: 9.5 })], { x: rx + 0.15, y: 2.0, w: rw - 0.3, h: 0.85, color: "FFFFFF", valign: "middle" });
  const inputs = ["Active merchants", "Share of payments at shops", "Payments per payer", "Credit users"];
  const tx = rx + 0.25;
  s.addShape(pres.shapes.LINE, { x: tx, y: 2.85, w: 0, h: 0.25 + 3 * 0.56 + 0.2, line: { color: HEX.grey, width: 1 } });
  inputs.forEach((t, i) => {
    const y = 3.05 + i * 0.56;
    s.addShape(pres.shapes.LINE, { x: tx, y: y + 0.22, w: 0.25, h: 0, line: { color: HEX.grey, width: 1 } });
    box(s, { x: tx + 0.25, y, w: rw - 0.5, h: 0.44, fill: { color: C.background2 }, objectName: `input ${i}` });
    T(s, t, { x: tx + 0.35, y, w: rw - 0.7, h: 0.44, fontSize: 9.5, valign: "middle" });
  });

  // bottom: why not the obvious metrics
  const by = 5.35;
  T(s, [
    b("Not share: ", { color: C.accent1 }), r("NPCI caps it at 30%.   "),
    b("Not volume: ", { color: C.accent1 }), r("cashback splitting inflates it.   "),
    b("Not monthly users: ", { color: C.accent1 }), r("most people have 2+ apps."),
  ], { x: M, y: by, w: W - 2 * M, h: 0.72, fontSize: 12 });
  soWhat(s, "Turn casual payers into weekly payers.");
}

// ---------- slide 5: options ----------
function slideOptions() {
  const s = frame("OPTIONS",
    "RICE favours habit rewards. Adding moat, money and regulation favours credit",
    "RICE = Reach × Impact × Confidence ÷ Effort. Weighted = 1-5 per criterion. All scores are our estimates.",
    "RICE picks habit rewards because it's cheap and reaches many users. But RICE ignores moat, revenue and regulation, and those decide GPay's future. Our weighted score puts credit first and habit rewards second. If we shift weight to moat, credit stays first but Maps passes habit rewards. We still pick habit rewards because it ships in one quarter. SoundPod is last on both scores.");
  const lw = 8.55;
  panelTitle(s, "Five options, two scores", M, 1.55, lw);
  const opts = [
    ["F  Credit at the pay button", "offer credit when people pay", "40", "2", "50%", "8", "5.0", "3", "3.60", "1", "▲2  best for revenue and moat", true],
    ["B  Habit rewards", "reward weekly shop payments", "120", "1", "80%", "2", "48.0", "1", "3.50", "2", "▼1  easy to copy, earns little", true],
    ["D  Maps \"Pay here\"", "pay button on Maps listings", "40", "1", "50%", "6", "3.3", "4", "3.10", "3", "▲1  only Google has Maps", false],
    ["A  Android fast pay", "scan to pay from the lock screen", "150", "1", "30%", "3", "15.0", "2", "2.85", "4", "▼2  antitrust risk, earns nothing", false],
    ["C  SoundPod push", "cheaper speakers for shops", "60", "1", "30%", "10", "1.8", "5", "2.50", "5", "=  speaker doesn't choose the app", false],
  ];
  const rows = [[hdr("Option"), hdr("R"), hdr("I"), hdr("C"), hdr("E"), hdr("RICE"), hdr("#"), hdr("Wtd"), hdr("#"), hdr("Why the rank moves")]];
  for (const [name, desc, R, I, Cf, E, rice, rr, wt, wr, why, top] of opts) {
    const f = top ? { fill: { color: "FBE9EB" } } : {};
    rows.push([
      { text: [b(name, { breakLine: true, fontSize: 10 }), r(desc, { fontSize: 8.5, color: C.text2 })], options: f },
      cell(R, f), cell(I, f), cell(Cf, f), cell(E, f),
      cell(rice, { bold: true, ...f }), cell(rr, { color: C.text2, ...f }),
      cell(wt, { bold: true, color: top ? C.accent1 : C.text1, ...f }), cell(wr, { color: C.text2, ...f }),
      cell(why, { fontSize: 9, ...f }),
    ]);
  }
  s.addTable(rows, { ...TABLE, x: M, y: 1.95, w: lw, colW: [2.85, 0.45, 0.32, 0.5, 0.35, 0.55, 0.3, 0.5, 0.3, 2.43], rowH: [0.28, 0.66, 0.66, 0.66, 0.66, 0.66], objectName: "scores table" });

  // right panel
  const rx = 9.3, rw = W - M - rx;
  panelTitle(s, "Why RICE alone doesn't fit GPay", rx, 1.55, rw);
  T(s, [
    br("No moat: rivals copy rewards fast", { bullet: { type: "number" } }),
    br("No revenue: payments earn nothing", { bullet: { type: "number" } }),
    br("No regulation: cap, fees, antitrust", { bullet: { type: "number" } }),
    r("Reach is a guess: GPay shares no user data", { bullet: { type: "number" } }),
  ], { x: rx, y: 1.97, w: rw, h: 1.75, fontSize: 12, paraSpaceAfter: 6 });
  T(s, "Our weights", { x: rx, y: 3.8, w: rw, h: 0.22, fontSize: 9.5, bold: true });
  const weights = [["NSM", 30], ["Moat", 20], ["Money", 20], ["Reg.", 15], ["Speed", 15]];
  let x = rx;
  weights.forEach(([k, v], i) => {
    const w = (v / 100) * rw;
    box(s, { x, y: 4.05, w: w - 0.03, h: 0.42, fill: { color: i === 0 ? HEX.red : i < 3 ? HEX.grey : HEX.light }, objectName: `weight ${k}` });
    T(s, `${k}\n${v}%`, { x, y: 4.05, w: w - 0.03, h: 0.42, fontSize: 8, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    x += w;
  });
  T(s, [b("Sensitivity: "), r("shift 10 points from impact to moat: credit stays first, and Maps overtakes habit rewards. We still pick habit rewards because it ships in one quarter.")], { x: rx, y: 4.6, w: rw, h: 1.35, fontSize: 11 });
  soWhat(s, "Use the weighted score to pick the bet. Use RICE inside it.");
}

// ---------- slide 6: plan ----------
function slidePlan() {
  const s = frame("PLAN",
    "Launch habit rewards in one quarter. Bring credit to the pay button within a year",
    "Sources: PhonePe and Paytm filings (FY25); NPCI credit-line data (Oct 2026); Google Pay blog. Scope uses MoSCoW. Quarters are illustrative.",
    "Habit rewards is the quick win: a small brand-funded reward for paying shops in 3 of 4 weeks. We test it against a holdout, and check people keep paying after the reward stops. That's what failed in 2020. Credit is the long-term bet. GPay already has loans and Flex, but they sit in a separate section. The first step needs no new lending: show credit at the pay screen to users already approved.");
  const cw = (W - 2 * M - 0.3) / 2;
  const cards = [
    ["B  Habit rewards", "QUICK WIN  ·  1 QUARTER", [
      ["What", "A brand-funded reward for paying shops in 3 of 4 weeks"],
      ["Must (v1)", "Streak rule, reward cap, fraud checks"],
      ["Won't (v1)", "Blanket cashback, points, lotteries"],
      ["First step", "Test against a holdout group"],
      ["Metric", "Weekly merchant payers vs holdout at day 30; retention after rewards stop"],
      ["Money", "Brands pay for rewards"],
    ]],
    ["F  Credit at the pay button", "LONG-TERM BET  ·  2-4 QUARTERS", [
      ["What", "Offer credit when people pay, with limits from GPay history"],
      ["Must (v1)", "Credit on the pay screen, partner banks only, clear rates"],
      ["Won't (v1)", "Google as lender, BNPL for first-time borrowers"],
      ["First step", "Show it first to users already approved for credit"],
      ["Metric", "Share of payments on credit at day 30; repayment rate"],
      ["Money", "Interest share and card fees. Paytm lent ₹13,958 Cr to merchants in FY25, PhonePe ₹4,507 Cr."],
    ]],
  ];
  cards.forEach(([name, tag, rows], i) => {
    const x = M + i * (cw + 0.3);
    T(s, tag, { x, y: 1.55, w: cw, h: 0.22, fontSize: 8.5, bold: true, color: C.accent1, charSpacing: 1 });
    T(s, name, { x, y: 1.78, w: cw, h: 0.32, fontFace: THEME.headFontFace, fontSize: 15 });
    s.addTable(rows.map(([k, v]) => [cell(k, { bold: true, fontSize: 11, fill: { color: C.background2 } }), cell(v, { fontSize: 11 })]),
      { ...TABLE, x, y: 2.17, w: cw, colW: [1.05, cw - 1.05], rowH: 0.46, objectName: `plan ${name}` });
  });
  // roadmap
  const ry = 5.03, lab = 1.9, gw = W - 2 * M - lab, qw = gw / 4;
  ["Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027"].forEach((q, i) => T(s, q, { x: M + lab + i * qw, y: ry, w: qw, h: 0.2, fontSize: 8.5, color: C.text2, align: "center" }));
  const bars = [
    ["B  Habit rewards", [[0, 1, "Holdout test", HEX.red], [1, 4, "National rollout", "E8979D"]]],
    ["F  Credit", [[0, 2, "Credit for approved users", HEX.grey], [2, 4, "Merchant loans", HEX.light]]],
  ];
  bars.forEach(([k, segs], j) => {
    const y = ry + 0.25 + j * 0.33;
    T(s, k, { x: M, y, w: lab, h: 0.28, fontSize: 9.5, bold: true, valign: "middle" });
    segs.forEach(([a, z, t, c]) => {
      box(s, { x: M + lab + a * qw + 0.02, y, w: (z - a) * qw - 0.04, h: 0.28, fill: { color: c }, objectName: `bar ${t}` });
      T(s, t, { x: M + lab + a * qw + 0.1, y, w: (z - a) * qw - 0.2, h: 0.28, fontSize: 8.5, color: c === HEX.light ? C.text1 : "FFFFFF", bold: true, valign: "middle" });
    });
  });
  soWhat(s, "Rewards build the habit. Credit earns from it.");
}

// ---------- slide 7: risks ----------
function slideRisks() {
  const s = frame("RISKS",
    "Avoid cashback wars and hardware races. Watch the 30% cap and antitrust",
    "Sources: NPCI; Inc42; Entrackr; Paytm (Jun 2026); Business Today (Dec 2024, Jul 2026); SCC Online; The Week; CCI orders (Oct 2022).",
    "Three things not to build: blanket cashback, non-finance super-app tabs, and a speaker race. In fairness, GPay might be fine: volume grew almost 14 times and its average payment matches PhonePe's. The 30% cap is due on 31 December; enforced, PhonePe loses about 15 points and GPay 2. The hard question is Android: it's Google's biggest edge, and the CCI already flagged GPay in 2022.");
  const cw = (W - 2 * M - 0.5) / 3;
  const x1 = M, x2 = M + cw + 0.25, x3 = M + 2 * (cw + 0.25);
  panelTitle(s, "Don't build", x1, 1.55, cw);
  const donts = [
    ["Another blanket cashback war", "GPay's share fell 10 points after it cut cashback. Paid-for users left. PhonePe still lost ₹2,792 Cr in FY26."],
    ["Super-app tabs outside finance", "PhonePe wrote off ₹684.7 Cr on its app store. Users leaving PhonePe complain about clutter."],
    ["A SoundPod hardware race", "Paytm has the most speakers but only 8% of payments. The speaker doesn't choose the app."],
  ];
  donts.forEach(([k, v], i) => {
    const y = 1.98 + i * 1.33;
    box(s, { x: x1, y, w: cw, h: 1.22, fill: { color: C.background2 }, objectName: `dont ${i}` });
    T(s, [br(`✕  ${k}`, { bold: true, color: C.accent1, fontSize: 12 }), r(v, { fontSize: 11 })], { x: x1 + 0.12, y: y + 0.08, w: cw - 0.24, h: 1.08, paraSpaceAfter: 3 });
  });

  panelTitle(s, "Why GPay might be fine", x2, 1.55, cw);
  const fine = [
    ["13.7×", "growth in payments since 2020"],
    ["#2 at 32.1%", "4× the next app"],
    ["₹1,297 vs ₹1,302", "average payment, GPay vs PhonePe"],
    ["Top on trust", "in a 2026 user survey"],
  ];
  fine.forEach(([n, l], i) => {
    const y = 2.0 + i * 0.83;
    T(s, n, { x: x2, y, w: cw, h: 0.38, fontFace: THEME.headFontFace, fontSize: 18 });
    T(s, l, { x: x2, y: y + 0.38, w: cw, h: 0.4, fontSize: 11, color: C.text2 });
  });
  T(s, [r("But GPay still grew slower than the market.")], { x: x2, y: 5.38, w: cw, h: 0.55, fontSize: 11, color: C.accent1 });

  panelTitle(s, "Risks to watch", x3, 1.55, cw);
  T(s, [b("30% cap: "), r("no app may handle over 30% of UPI payments. Due 31 Dec 2026. Never enforced.")], { x: x3, y: 1.98, w: cw, h: 0.95, fontSize: 11 });
  s.addTable([
    [hdr("If NPCI…"), hdr("Effect")],
    [cell("Enforces", { bold: true, fontSize: 9 }), cell("PhonePe loses ~15 points, GPay ~2", { fontSize: 10.5 })],
    [cell("Extends / scraps", { bold: true, fontSize: 9 }), cell("No change", { fontSize: 10.5 })],
  ], { ...TABLE, x: x3, y: 2.95, w: cw, colW: [1.1, cw - 1.1], objectName: "cap scenarios" });
  T(s, [b("MDR: "), r("a 0.4% fee on payments over ₹2,000. Delayed, no date.")], { x: x3, y: 4.05, w: cw, h: 0.6, fontSize: 11 });
  box(s, { x: x3, y: 4.7, w: cw, h: 1.3, fill: { color: "FBE9EB" }, objectName: "hard question" });
  T(s, [br("The hard question", { bold: true, color: C.accent1 }), r("Android fast pay would top RICE without antitrust risk. CCI flagged GPay in 2022. Can Google use Android without repeating that?", { fontSize: 11 })], { x: x3 + 0.12, y: 4.78, w: cw - 0.24, h: 1.15, fontSize: 12, paraSpaceAfter: 3 });
  soWhat(s, "Grow habit and revenue, not share.");
}

// ---------- A1: reliability ----------
function slideReliability() {
  const s = frame("BACKUP",
    "GPay's banks failed more often only in Jan 2022. They're level now",
    "Source: NPCI payer bank data, technical declines, Jan 2022-Oct 2026.",
    "Apps run on partner banks: PhonePe mainly on Yes Bank, GPay on four big banks. In January 2022, GPay's banks failed 30 to 50 times more often. From early 2022 to late 2023 NPCI reports almost zero failures everywhere, which looks misreported. Since late 2023 the banks are level.");
  const rows = fs.readFileSync(path.join(__dirname, "..", "data", "npci_payer_psp_declines_monthly.csv"), "utf8").trim().split("\n").slice(1).map((l) => l.replace(/"([^"]*)"/g, (_, q) => q.replace(/,/g, "")).split(","));
  const banks = [["Yes Bank", "Yes Bank (PhonePe)", HEX.green], ["Axis Bank", "Axis (GPay)", HEX.red], ["HDFC Bank", "HDFC (GPay)", "B8434D"], ["ICICI Bank", "ICICI (GPay)", "E8979D"], ["SBI", "SBI (GPay)", "8E2F37"]];
  const by = {};
  for (const [m, p, , , , td] of rows) (by[m] ||= {})[p] = +td;
  const months = Object.keys(by).sort().filter((m) => banks.every(([k]) => by[m][k] !== undefined));
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const labels = months.map((m) => (m.endsWith("-01") ? m.slice(0, 4) : ""));
  panelTitle(s, "Technical decline rate by payer bank, % of payments", M, 1.55, 8.2);
  s.addChart(pres.charts.LINE, banks.map(([k, n]) => ({ name: n, labels, values: months.map((m) => by[m][k]) })), {
    x: M, y: 1.95, w: 8.2, h: 3.9, chartColors: banks.map(([, , c]) => c), lineSize: 1.75, lineDataSymbol: "none",
    showLegend: true, legendPos: "b", legendFontSize: 9, legendFontFace: "+mn-lt",
    valAxisLabelFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFontFace: "+mn-lt", catAxisLabelFontFace: "+mn-lt",
    valAxisLabelColor: HEX.mid, catAxisLabelColor: HEX.mid, valAxisLabelFormatCode: "0.0", catAxisLabelFrequency: 1, catAxisLabelRotate: 0,
    valGridLine: { color: "EEEEEE", size: 0.5 }, catGridLine: { style: "none" }, catAxisLineShow: false, valAxisLineShow: false, objectName: "td chart",
  });
  const rx = 9.0, rw = W - M - rx;
  panelTitle(s, "Snapshot (technical declines, %)", rx, 1.55, rw);
  const pick = ["2022-01", "2022-02", "2025-01", "2026-10"];
  s.addTable([
    [hdr("Month"), hdr("Yes"), hdr("Axis"), hdr("HDFC"), hdr("ICICI"), hdr("SBI")],
    ...pick.map((m) => [cell(`${MON[+m.slice(5) - 1]} ${m.slice(0, 4)}`, { fontSize: 9 }), ...banks.map(([k], i) => cell(by[m][k].toFixed(2), { fontSize: 9, bold: m === "2022-01", color: m === "2022-01" ? (i === 0 ? C.accent2 : C.accent1) : C.text1 }))]),
  ], { ...TABLE, x: rx, y: 1.95, w: rw, colW: [0.95, 0.56, 0.56, 0.56, 0.56, rw - 3.19], objectName: "td table" });
  T(s, [
    b("Caveats", { breakLine: true }),
    br("Bank data, not app data", { bullet: true }),
    br("2022-23 looks misreported", { bullet: true }),
    r("Business declines ignored", { bullet: true }),
  ], { x: rx, y: 4.0, w: rw, h: 1.9, fontSize: 11, paraSpaceAfter: 4 });
  soWhat(s, "Reliability may have mattered in 2020-21. It isn't a gap now.");
}

// ---------- A2: credit ----------
function slideCredit() {
  const s = frame("BACKUP",
    "GPay sells credit but hides it. Credit works best at the pay button",
    "Sources: NPCI credit-line data (Oct 2026); Paytm, PhonePe, Navi filings; Ant Group, Nubank, Mercado Libre; OJK; Google Pay blog.",
    "GPay already offers loans, gold loans, merchant loans and the Flex card, but shares no volumes. Paytm has three-quarters of credit-line value on UPI. The models that work put credit at the pay button and set limits from payment data. GPay's biggest gap is merchant loans repaid from daily QR payments.");
  const lw = 4.0;
  panelTitle(s, "Credit lines on UPI, share of value, Oct 2026", M, 1.55, lw);
  s.addTable([
    [hdr("App"), hdr("Share")],
    [cell("Paytm", { bold: true }), cell("73.8%", { bold: true })], [cell("PhonePe"), cell("7.9%")], [cell("slice"), cell("5.2%")], [cell("Navi"), cell("4.8%")], [cell("BharatPe"), cell("2.8%")],
    [cell("Others, incl. GPay", { color: C.accent1, bold: true }), cell("5.5%", { color: C.accent1, bold: true })],
  ], { ...TABLE, x: M, y: 1.95, w: lw, colW: [2.8, 1.2], rowH: 0.3, objectName: "credit share" });
  T(s, [b("GPay today: "), r("loans, gold loans, merchant loans and the Flex card. It never shares volumes.")], { x: M, y: 4.2, w: lw, h: 0.95, fontSize: 11 });
  T(s, [b("Rivals: "), r("merchant loans Paytm ₹13,958 Cr, PhonePe ₹4,507 Cr (FY25); Navi ₹23,287 Cr disbursed (FY26).")], { x: M, y: 5.0, w: lw, h: 0.75, fontSize: 11 });

  const rx = 4.85, rw = W - M - rx;
  panelTitle(s, "What works, and where GPay stands", rx, 1.55, rw);
  const P = [
    ["Credit as a default option at the pay screen", "Huabei, Paytm Postpaid", "✕", "Loans sit in a separate section"],
    ["Limits from the platform's own payment data, small then growing", "Nubank, Mercado Pago, Huabei", "✕", "Lender pre-qualifies"],
    ["Merchant loans repaid from daily QR settlements", "Paytm ₹13,958 Cr, PhonePe ₹4,507 Cr", "~", "Sachet loans exist; no scale or auto-repay disclosed"],
    ["Bank lends, platform distributes", "Huabei (98% partner-funded); RBI 2022 rules", "✓", "Strength: Axis, SBI, others"],
    ["Guardrails: caps, clear pricing, no wallet loading", "Indonesia paylater NPLs; RBI BNPL curbs", "✓", "Mostly in place"],
  ];
  s.addTable([
    [hdr("Principle"), hdr("Evidence it works"), hdr("GPay"), hdr("Detail")],
    ...P.map(([p, e, g, d]) => [cell(p, { fontSize: 9.5, bold: true }), cell(e, { fontSize: 9 }), cell(g, { fontSize: 13, bold: true, align: "center", color: g === "✓" ? C.accent2 : C.accent1 }), cell(d, { fontSize: 9 })]),
  ], { ...TABLE, x: rx, y: 1.95, w: rw, colW: [3.0, 2.2, 0.5, rw - 5.7], rowH: [0.28, 0.55, 0.55, 0.55, 0.55, 0.55], objectName: "principles" });
  T(s, [b("Why it's small in India: "), r("RBI rules and few card holders (~50M).")], { x: rx, y: 5.3, w: rw, h: 0.5, fontSize: 11 });
  soWhat(s, "GPay's biggest gap is merchant loans repaid from daily QR payments.");
}

// ---------- A3: screens ----------
function slideScreens() {
  const s = frame("BACKUP",
    "Both apps do the same jobs. PhonePe sells breadth, GPay simplicity",
    "Images: Play Store and App Store listings, © Google LLC and © PhonePe. Marketing images, not live flows.",
    "These are the official store images. GPay leads with paying people. PhonePe leads with financial services and rewards. They're marketing images, so I don't count taps from them.");
  const base = path.join(__dirname, "..", "screens");
  const sets = [
    ["Google Pay", "gpay", HEX.red, ["gpay_home_appstore.png", "gpay_home_payanyone_playstore.png", "gpay_pay_contact_people_playstore.png", "gpay_scan_pay_playstore.png", "gpay_bills_recharge_playstore.png", "gpay_bank_balance_playstore.png"]],
    ["PhonePe", "phonepe", HEX.green, ["phonepe_home_rewards_appstore.png", "phonepe_home_financial_services_playstore.png", "phonepe_pay_contact_appstore.png", "phonepe_scan_pay_playstore.png", "phonepe_bills_recharge_playstore.png", "phonepe_bank_balance_playstore.png"]],
  ];
  const caps = ["Home", "Home / main pitch", "Pay a contact", "Scan & pay", "Bills & recharge", "Bank balance"];
  const ih = 2.05, iw = ih * 0.5625;
  sets.forEach(([name, dir, col, files], j) => {
    const y = 1.6 + j * 2.25;
    T(s, name, { x: M, y: y + 0.85, w: 1.3, h: 0.35, fontSize: 13, bold: true, color: col });
    files.forEach((f, i) => {
      const x = M + 1.5 + i * ((W - 2 * M - 1.5) / 6);
      s.addImage({ path: path.join(base, dir, f), x, y, w: iw, h: ih, altText: `${name} store screenshot: ${caps[i]}`, sizing: { type: "contain", w: iw, h: ih } });
      T(s, caps[i], { x: x + iw + 0.05, y: y + ih - 0.3, w: 0.6, h: 0.3, fontSize: 7.5, color: C.text2, valign: "bottom" });
    });
  });
  soWhat(s, "The screens look alike. Habit at the counter decides.");
}

// ---------- A4: method ----------
function slideMethod() {
  const s = frame("BACKUP",
    "Shares come from NPCI data. Estimates are labelled",
    "Full sources: research files and analysis-gate1.md in the gpay-teardown folder.",
    "All share numbers come from NPCI's own table. Estimates are labelled. GPay publishes almost nothing, so some comparisons use PhonePe's numbers as a proxy.");
  const cw = (W - 2 * M - 0.3) / 2;
  panelTitle(s, "Data and estimates", M, 1.55, cw);
  s.addTable([
    [hdr("Item"), hdr("Type"), hdr("Basis")],
    [cell("App shares, volumes, growth", { fontSize: 9 }), cell("NPCI data", { fontSize: 9, color: C.accent2, bold: true }), cell("App-wise UPI table, Jun 2020-Oct 2026 (77 months)", { fontSize: 9 })],
    [cell("Failure rates", { fontSize: 9 }), cell("NPCI data", { fontSize: 9, color: C.accent2, bold: true }), cell("Top 15 PSPs, payer side, Jan 2022-Oct 2026", { fontSize: 9 })],
    [cell("Credit-line shares", { fontSize: 9 }), cell("NPCI data", { fontSize: 9, color: C.accent2, bold: true }), cell("Sep-Oct 2026, % only", { fontSize: 9 })],
    [cell("GPay ~200M monthly users", { fontSize: 9 }), cell("Estimate", { fontSize: 9, color: C.accent1, bold: true }), cell("GPay volume ≈ 0.71× PhonePe's, applied to PhonePe's 290M actives", { fontSize: 9 })],
    [cell("GPay merchant reach ~26%", { fontSize: 9 }), cell("Estimate", { fontSize: 9, color: C.accent1, bold: true }), cell("23M onboarded × PhonePe's 23.5% active rate", { fontSize: 9 })],
    [cell("UPI monthly actives 350-450M", { fontSize: 9 }), cell("Estimate", { fontSize: 9, color: C.accent1, bold: true }), cell("Onboarded users less dormant; NPCI publishes no MAU", { fontSize: 9 })],
    [cell("Segments", { fontSize: 9 }), cell("Proxy", { fontSize: 9, color: C.text2, bold: true }), cell("PhonePe DRHP split, applied to GPay", { fontSize: 9 })],
    [cell("RICE and weighted scores", { fontSize: 9 }), cell("Judgement", { fontSize: 9, color: C.text2, bold: true }), cell("Reason for every score in analysis-gate1.md §9, §16-19", { fontSize: 9 })],
  ], { ...TABLE, x: M, y: 1.95, w: cw, colW: [1.85, 0.85, cw - 2.7], objectName: "method table" });
  const rx = M + cw + 0.3;
  panelTitle(s, "Method notes and limits", rx, 1.55, cw);
  T(s, [
    br("Share = app volume ÷ all listed apps", { bullet: true }),
    br("Shows net change, not who switched", { bullet: true }),
    br("GPay publishes no users or merchants", { bullet: true }),
    br("Quotes: 25 public posts, not a sample", { bullet: true }),
    r("Unverified: date of GPay's rewards cut", { bullet: true }),
  ], { x: rx, y: 1.97, w: cw, h: 4.0, fontSize: 12, paraSpaceAfter: 8 });
  soWhat(s, "The diagnosis uses NPCI data. The plan uses labelled estimates.");
}

(async () => {
  pres.addSection({ title: "Teardown" });
  slideCover();
  slideSummary();
  slideStory();
  slideWhy();
  slideWho();
  slideOptions();
  slidePlan();
  slideRisks();
  pres.addSection({ title: "Backup" });
  slideReliability();
  slideCredit();
  slideScreens();
  slideMethod();
  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
