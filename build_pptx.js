const pptxgen = require("pptxgenjs");

// ============================================================
// Elon's Plugs — March Madness 2026 Bracket Tracker
// Pixel-matched to Men's Finance Bracket reference deck
// Data as of: March 22, 2026 8:30 PM ET
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Taylor Duran";
pres.title = "Elon's Plugs — March Madness 2026 Bracket Tracker";

// ── Exact color palette from reference ──
const C = {
  darkBg:     "414548",
  tealBg:     "2D4F5F",
  red:        "DC1A41",
  gold:       "C4963C",
  green:      "2D8B4E",
  offWhite:   "F0EFEB",
  white:      "FFFFFF",
  darkText:   "333333",
  bodyText:   "555555",
  muted:      "999999",
  lightGray:  "E8E6E2",
  tableHead:  "6B5B3E",
  navy:       "2B4C6F",
  blueBar:    "2B4C6F",
  pinkRow:    "FBF0ED",
  creamRow:   "FBF5ED",
};

// ── Data (8:30 PM ET, post-Tennessee 79-72 Virginia) ──
const standings = [
  { rank: 1, name: "TaylorDuran",       pts: 450, pct: 97.46, w: 37, l: 8,  max: 1770, rem: 1320, r64: 290, r32: 160, champ: "Arizona",   chalk: 81, uniq: 5  },
  { rank: 2, name: "Fizzz3",            pts: 440, pct: 79.86, w: 35, l: 10, max: 1760, rem: 1320, r64: 260, r32: 180, champ: "Arizona",   chalk: 75, uniq: 10 },
  { rank: 3, name: "Capn'Kirk",         pts: 430, pct: 87.97, w: 34, l: 10, max: 1790, rem: 1360, r64: 250, r32: 180, champ: "Duke",      chalk: 95, uniq: 0  },
  { rank: 4, name: "Paul233365",        pts: 400, pct: 70.53, w: 32, l: 13, max: 1720, rem: 1320, r64: 220, r32: 180, champ: "Florida",   chalk: 71, uniq: 14 },
  { rank: 5, name: "twnutt",            pts: 390, pct: 58.38, w: 31, l: 14, max: 1710, rem: 1320, r64: 230, r32: 160, champ: "Auburn",    chalk: 63, uniq: 8  },
  { rank: 6, name: "inursha",           pts: 370, pct: 47.44, w: 30, l: 15, max: 1690, rem: 1320, r64: 210, r32: 160, champ: "Houston",   chalk: 66, uniq: 12 },
  { rank: 7, name: "ESPNFAN828",        pts: 350, pct: 29.65, w: 28, l: 17, max: 1670, rem: 1320, r64: 190, r32: 160, champ: "Duke",      chalk: 85, uniq: 2  },
  { rank: 8, name: "ESPNFAN969",        pts: 310, pct: 10.79, w: 25, l: 20, max: 1630, rem: 1320, r64: 150, r32: 160, champ: "Tennessee", chalk: 58, uniq: 18 },
];

const TOTAL_SLIDES = 10;
const groupAvg = Math.round(standings.reduce((s, p) => s + p.pts, 0) / standings.length);
const updateTime = "Mar 22, 2026, 8:30 PM CT";

// Champion colors
const champColor = {
  "Arizona": C.red, "Duke": C.navy, "Florida": C.green,
  "Auburn": C.gold, "Houston": "1A5276", "Tennessee": "E8793A",
};

// Bar colors for points chart: top 3 get accent, rest navy
const barColor = (rank) => {
  if (rank === 1) return C.gold;
  if (rank <= 3) return C.gold;
  return C.blueBar;
};

// ── Helpers ──
function addFooter(slide, slideNum) {
  slide.addText("Powered by Claude", {
    x: 0.5, y: 5.15, w: 3, h: 0.35,
    fontSize: 8, fontFace: "Calibri", color: C.muted, valign: "bottom"
  });
  slide.addText(`${slideNum} / ${TOTAL_SLIDES}`, {
    x: 7.5, y: 5.15, w: 2, h: 0.35,
    fontSize: 8, fontFace: "Calibri", color: C.muted, align: "right", valign: "bottom"
  });
}

function addTitle(slide, num, title, slideNum) {
  slide.addText(String(num).padStart(2, "0"), {
    x: 0.5, y: 0.25, w: 0.55, h: 0.55,
    fontSize: 22, fontFace: "Calibri", color: C.red, bold: true, margin: 0
  });
  slide.addText(title, {
    x: 1.05, y: 0.25, w: 8.45, h: 0.55,
    fontSize: 20, fontFace: "Calibri", color: C.darkText, bold: true, margin: 0
  });
  addFooter(slide, slideNum);
}

const mkShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

// ════════════════════════════════════════════════════════
// SLIDE 1 — TITLE  (dark charcoal, red top bar)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("Elon's Plugs", {
    x: 0.7, y: 1.1, w: 8.6, h: 1.2,
    fontSize: 56, fontFace: "Calibri", color: C.white, bold: true
  });
  s.addText("March Madness 2026", {
    x: 0.7, y: 2.25, w: 8.6, h: 0.65,
    fontSize: 28, fontFace: "Calibri", color: C.red, bold: true
  });
  s.addText("ESPN Tournament Challenge  /  Bracket Analytics  /  Mar 22, 2026", {
    x: 0.7, y: 3.15, w: 8.6, h: 0.35,
    fontSize: 11, fontFace: "Calibri", color: C.muted
  });
  s.addText(`API Data: ${updateTime}  |  Deck Generated: ${updateTime}`, {
    x: 0.7, y: 3.5, w: 8.6, h: 0.35,
    fontSize: 10, fontFace: "Calibri", color: C.muted
  });
  s.addText("8 participants  \u00B7  Bracket Analytics", {
    x: 0.5, y: 4.85, w: 5, h: 0.4,
    fontSize: 9, fontFace: "Calibri", color: C.muted
  });
  s.addText("Built with Claude", {
    x: 6.5, y: 4.85, w: 3, h: 0.4,
    fontSize: 9, fontFace: "Calibri", color: C.muted, align: "right"
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 2 — EXECUTIVE SUMMARY
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 1, "TaylorDuran Leads by 10 \u2014 But Fizzz3 Is Within Striking Distance", 2);

  // 5 KPI cards
  const kpis = [
    { label: "CURRENT LEADER", value: "450",            sub: "TaylorDuran",   accent: C.red   },
    { label: "LEAD MARGIN",    value: "+10",             sub: "vs Fizzz3",     accent: C.green },
    { label: "GROUP AVERAGE",  value: String(groupAvg),  sub: "8 entries",     accent: C.gold  },
    { label: "CHAMP PICKS",   value: "8 / 8",           sub: "0 busted",      accent: "6B8E5A"},
    { label: "HIGHEST CEILING",value: "1790",            sub: "Capn'Kirk",     accent: C.green },
  ];

  const cardW = 1.7, cardH = 1.15, gap = 0.13, startX = 0.5, cardY = 1.0;
  kpis.forEach((kpi, i) => {
    const cx = startX + i * (cardW + gap);
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: cardW, h: 0.04, fill: { color: kpi.accent } });
    s.addText(kpi.label, { x: cx + 0.1, y: cardY + 0.12, w: cardW - 0.2, h: 0.2, fontSize: 7, fontFace: "Calibri", color: C.muted, bold: true });
    s.addText(kpi.value, { x: cx + 0.1, y: cardY + 0.3, w: cardW - 0.2, h: 0.5, fontSize: 28, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
    s.addText(kpi.sub,   { x: cx + 0.1, y: cardY + 0.85, w: cardW - 0.2, h: 0.2, fontSize: 8, fontFace: "Calibri", color: C.bodyText });
  });

  // KEY INSIGHT
  const ky = 2.4;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ky, w: 5.6, h: 0.01, fill: { color: C.lightGray } });
  s.addText("KEY INSIGHT", { x: 0.5, y: ky + 0.06, w: 2, h: 0.22, fontSize: 8, fontFace: "Calibri", color: C.darkText, bold: true });
  s.addText(
    "TaylorDuran leads at 450 pts (37-8). Fizzz3 trails by only 10 points at 440. With 1320 points still up for grabs, any contender could overtake the lead in later rounds. Most popular champion pick: Arizona (2 brackets).",
    { x: 0.5, y: ky + 0.28, w: 5.6, h: 0.7, fontSize: 9, fontFace: "Calibri", color: C.bodyText }
  );

  // TOP 8 bar chart (right side)
  s.addText("TOP 8", { x: 6.4, y: 2.4, w: 1, h: 0.22, fontSize: 8, fontFace: "Calibri", color: C.darkText, bold: true });

  standings.slice(0, 7).forEach((p, i) => {
    const ry = 2.7 + i * 0.34;
    const maxBarW = 1.5;
    const bw = (p.pts / 470) * maxBarW;
    const isTop3 = i < 3;
    const bColor = i === 0 ? C.gold : (i < 4 ? C.gold : C.blueBar);

    s.addText(String(i + 1), { x: 6.4, y: ry, w: 0.2, h: 0.28, fontSize: 8, fontFace: "Calibri", color: isTop3 ? C.red : C.bodyText, bold: true, valign: "middle" });
    s.addText(p.name, { x: 6.6, y: ry, w: 1.3, h: 0.28, fontSize: 7.5, fontFace: "Calibri", color: C.darkText, bold: isTop3, valign: "middle" });

    s.addShape(pres.shapes.RECTANGLE, { x: 7.95, y: ry + 0.04, w: bw, h: 0.2, fill: { color: bColor } });
    s.addText(`${p.pts} pts`, { x: 7.95 + bw + 0.04, y: ry, w: 0.6, h: 0.28, fontSize: 7, fontFace: "Calibri", color: C.bodyText, valign: "middle" });
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 3 — COMPLETE STANDINGS TABLE
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 2, "Complete Standings \u2014 8 Brackets", 3);

  const hOpt = (a = "center") => ({ fill: { color: C.tableHead }, color: C.white, bold: true, fontSize: 8.5, fontFace: "Calibri", align: a, valign: "middle" });

  const header = [
    { text: "RANK",   options: hOpt() },
    { text: "OWNER",  options: hOpt("left") },
    { text: "PTS",    options: hOpt() },
    { text: "MAX",    options: hOpt() },
    { text: "W-L",    options: hOpt() },
    { text: "R64",    options: hOpt() },
    { text: "R32",    options: hOpt() },
    { text: "CHAMPION",options: hOpt("left") },
    { text: "CHAMP\nSTATUS", options: hOpt() },
  ];

  const rows = standings.map((p, i) => {
    const rowFill = i % 2 === 0 ? C.creamRow : C.white;
    const b = (a = "center") => ({ fill: { color: rowFill }, fontSize: 9, fontFace: "Calibri", color: C.darkText, align: a, valign: "middle" });

    return [
      { text: String(p.rank), options: { ...b(), bold: true, color: i < 3 ? C.red : C.darkText } },
      { text: p.name,         options: { ...b("left"), bold: i < 3 } },
      { text: String(p.pts),  options: { ...b(), bold: true } },
      { text: String(p.max),  options: { ...b(), bold: true, color: C.gold } },
      { text: `${p.w}-${p.l}`,options: b() },
      { text: String(p.r64),  options: b() },
      { text: String(p.r32),  options: b() },
      { text: p.champ,        options: b("left") },
      { text: "IN THE\nHUNT", options: { ...b(), bold: true, color: C.green, fontSize: 8 } },
    ];
  });

  s.addTable([header, ...rows], {
    x: 0.4, y: 1.0, w: 9.2,
    colW: [0.5, 1.65, 0.55, 0.6, 0.6, 0.5, 0.5, 1.2, 0.75],
    rowH: [0.35, ...Array(8).fill(0.45)],
    border: { pt: 0.3, color: C.lightGray },
    margin: [2, 5, 2, 5]
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 4 — WHAT TO WATCH
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 3, "What to Watch \u2014 Where This Race Gets Decided", 4);

  const items = [
    { t: "Can Fizzz3 Close the 10-Point Gap?", b: "Fizzz3 trails with 440 points but has a max ceiling of 1760. The gap narrows if their champion (Arizona) advances." },
    { t: "Capn'Kirk Has the Highest Ceiling in the Group", b: "Capn'Kirk has a max ceiling of 1790 vs. the leader's 1770. No one else can match that ceiling, but a few well-timed upsets could close the gap fast." },
    { t: "3 Brackets Have a Unique Champion Pick \u2014 Who's Next?", b: "Most popular picks: Arizona (2), Duke (2). Florida, Auburn, Houston, Tennessee each picked once. If favorites get knocked out, contrarian pickers gain." },
    { t: "3-Way Tie at 390 Points Makes the Middle Wide Open", b: "Paul233365, twnutt are deadlocked. The tiebreaker will come down to who nailed the higher-value later rounds \u2014 Sweet 16 picks are worth 2x R64." },
    { t: "The Math: Later Rounds Are Worth Exponentially More", b: "R64 = 10 pts, R32 = 20 pts, S16 = 40 pts, E8 = 80 pts, F4 = 160 pts, Championship = 320 pts. A single correct championship pick is worth 32 first-round picks." },
  ];

  let yy = 0.95;
  items.forEach(item => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yy, w: 9.0, h: 0.82, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yy, w: 0.05, h: 0.82, fill: { color: C.red } });
    s.addText(item.t, { x: 0.72, y: yy + 0.06, w: 8.6, h: 0.25, fontSize: 10.5, fontFace: "Calibri", color: C.darkText, bold: true });
    s.addText(item.b, { x: 0.72, y: yy + 0.32, w: 8.6, h: 0.42, fontSize: 8.5, fontFace: "Calibri", color: C.bodyText });
    yy += 0.9;
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 5 — POINTS BAR CHART
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 4, "TaylorDuran Holds a 10-Point Lead", 5);

  const sorted = [...standings].sort((a, b) => a.pts - b.pts);

  // Custom horizontal bars (matching reference exactly: top bars gold, rest navy)
  sorted.forEach((p, i) => {
    const ry = 1.1 + i * 0.5;
    const maxBarW = 5.5;
    const bw = (p.pts / 470) * maxBarW;
    const origRank = standings.findIndex(s => s.name === p.name) + 1;
    const color = origRank <= 3 ? C.gold : C.blueBar;

    s.addText(p.name, { x: 0.5, y: ry, w: 1.8, h: 0.4, fontSize: 9, fontFace: "Calibri", color: C.darkText, align: "right", valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 2.4, y: ry + 0.06, w: bw, h: 0.28, fill: { color } });
    s.addText(String(p.pts), { x: 2.4 + bw + 0.08, y: ry, w: 0.6, h: 0.4, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
  });

  addFooter(s, 5);
}

// ════════════════════════════════════════════════════════
// SLIDE 6 — ROUND BY ROUND (stacked bars)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 5, "R64 Dominance Drives the Leaderboard", 6);

  const sorted = [...standings].sort((a, b) => (a.r64 + a.r32) - (b.r64 + b.r32));

  s.addChart(pres.charts.BAR, [
    { name: "R64", labels: sorted.map(p => p.name), values: sorted.map(p => p.r64) },
    { name: "R32", labels: sorted.map(p => p.name), values: sorted.map(p => p.r32) },
  ], {
    x: 0.3, y: 1.0, w: 9.4, h: 4.1,
    barDir: "bar",
    barGrouping: "stacked",
    chartColors: [C.blueBar, C.green],
    chartArea: { fill: { color: C.offWhite } },
    catAxisLabelColor: C.darkText,
    catAxisLabelFontSize: 9,
    catAxisLabelFontFace: "Calibri",
    catGridLine: { style: "none" },
    catAxisLineShow: false,
    valAxisLabelColor: C.muted,
    valAxisLabelFontSize: 8,
    valGridLine: { color: C.lightGray, size: 0.5 },
    valAxisLineShow: false,
    showValue: false,
    showLegend: true,
    legendPos: "t",
    legendFontSize: 9,
    legendColor: C.darkText,
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 7 — CHAMPION PICKS (bar left + status right)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  const cc = {};
  standings.forEach(p => { cc[p.champ] = (cc[p.champ] || 0) + 1; });
  const champList = Object.entries(cc).sort((a, b) => b[1] - a[1]);

  addTitle(s, 6, `${champList[0][1]} of 8 Brackets Riding on ${champList[0][0]}`, 7);

  // Bar chart left
  const sortedC = [...champList].sort((a, b) => a[1] - b[1]);
  sortedC.forEach((entry, i) => {
    const ry = 1.2 + i * 0.6;
    const maxBarW = 3.8;
    const bw = (entry[1] / 3) * maxBarW;
    const color = champColor[entry[0]] || C.blueBar;

    s.addText(entry[0], { x: 0.5, y: ry, w: 1.2, h: 0.45, fontSize: 10, fontFace: "Calibri", color: C.darkText, align: "right", valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.8, y: ry + 0.08, w: bw, h: 0.3, fill: { color } });
    s.addText(String(entry[1]), { x: 1.8 + bw + 0.08, y: ry, w: 0.4, h: 0.45, fontSize: 11, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
  });

  // CHAMPION STATUS table right
  s.addText("CHAMPION STATUS", { x: 5.8, y: 1.0, w: 3.7, h: 0.3, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true });

  const champSorted = [...champList].sort((a, b) => b[1] - a[1]);
  let ty = 1.35;
  champSorted.forEach(entry => {
    const color = champColor[entry[0]] || C.navy;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: ty, w: 3.7, h: 0.48, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: ty, w: 0.05, h: 0.48, fill: { color } });
    s.addText(entry[0], { x: 6.0, y: ty, w: 1.4, h: 0.48, fontSize: 10, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
    s.addText(`${entry[1]} pick${entry[1] > 1 ? "s" : ""}`, { x: 7.4, y: ty, w: 0.9, h: 0.48, fontSize: 9, fontFace: "Calibri", color: C.bodyText, valign: "middle", align: "center" });
    s.addText("IN THE\nHUNT", { x: 8.4, y: ty, w: 1.0, h: 0.48, fontSize: 8, fontFace: "Calibri", color: C.green, bold: true, valign: "middle", align: "right" });
    ty += 0.56;
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 8 — MAX CEILING (current vs potential bars)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 7, "All 8 Champion Picks Still in the Hunt \u2014 No One Is Busted", 8);

  s.addText("Current points (solid) vs. maximum potential (total bar). All 8 brackets shown.", {
    x: 0.5, y: 0.8, w: 9, h: 0.25,
    fontSize: 8.5, fontFace: "Calibri", color: C.bodyText, italic: true
  });

  const maxPossible = Math.max(...standings.map(p => p.max));
  const barMaxW = 4.2;

  standings.forEach((p, i) => {
    const ry = 1.2 + i * 0.46;
    const currentW = (p.pts / maxPossible) * barMaxW;
    const maxW = (p.max / maxPossible) * barMaxW;
    const bColor = champColor[p.champ] || C.blueBar;

    s.addText(String(p.rank), { x: 0.5, y: ry, w: 0.35, h: 0.38, fontSize: 13, fontFace: "Calibri", color: C.red, bold: true, valign: "middle" });
    s.addText(p.name, { x: 0.85, y: ry, w: 1.6, h: 0.38, fontSize: 10, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });

    // Max bar (light gray)
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6, y: ry + 0.06, w: maxW, h: 0.26, fill: { color: C.lightGray } });
    // Current bar (champion colored)
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6, y: ry + 0.06, w: currentW, h: 0.26, fill: { color: bColor } });
    // Label
    s.addText(`${p.pts} / ${p.max}`, { x: 2.6 + maxW + 0.12, y: ry, w: 1.2, h: 0.38, fontSize: 9.5, fontFace: "Calibri", color: C.bodyText, valign: "middle" });
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 9 — CLOSING: "Game On."
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("Game On.", {
    x: 0.7, y: 2.4, w: 8.6, h: 1.2,
    fontSize: 64, fontFace: "Calibri", color: C.red, bold: true
  });
  s.addText("The bracket race continues.", {
    x: 0.7, y: 3.55, w: 8.6, h: 0.5,
    fontSize: 16, fontFace: "Calibri", color: C.muted
  });
  s.addText(`Updated ${updateTime}  \u00B7  ESPN Tournament Challenge  \u00B7  8 entries`, {
    x: 0.5, y: 4.85, w: 5.5, h: 0.4,
    fontSize: 9, fontFace: "Calibri", color: C.muted
  });
  s.addText("Built with Claude", {
    x: 6.5, y: 4.85, w: 3, h: 0.4,
    fontSize: 9, fontFace: "Calibri", color: C.muted, align: "right"
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 10 — APPENDIX DIVIDER
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.tealBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("Live Tracker", {
    x: 0.7, y: 2.7, w: 8.6, h: 1.0,
    fontSize: 44, fontFace: "Calibri", color: C.white, bold: true
  });
  s.addText("tdswim.github.io/Elons-Plugs-2026", {
    x: 0.7, y: 3.7, w: 8.6, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: C.muted
  });

  s.addText(`${TOTAL_SLIDES} / ${TOTAL_SLIDES}`, {
    x: 7.5, y: 5.15, w: 2, h: 0.35,
    fontSize: 8, fontFace: "Calibri", color: C.muted, align: "right", valign: "bottom"
  });
}

// ── Write file ──
const outPath = "/sessions/gifted-relaxed-clarke/mnt/Elon Bracket Tracker/Elons_Plugs_Bracket_Tracker.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("SUCCESS: " + outPath);
}).catch(err => {
  console.error("ERROR:", err);
});
