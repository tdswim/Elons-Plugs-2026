const pptxgen = require("pptxgenjs");

// ============================================================
// Elon's Plugs — March Madness 2026 Bracket Tracker
// Pixel-matched to Men's Finance Bracket reference deck
// Data as of: March 29, 2026 — Elite 8 Complete
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
  busted:     "B04040",
};

// ── Data (Elite 8 Complete — Mar 29, 2026) ──
// Final Four: Arizona vs Michigan, UConn vs Illinois
// Duke ELIMINATED — all 5 Duke pickers lose championship points
const standings = [
  { rank: 1, name: "TaylorDuran",         pts: 910, pct: 97.61, w: 47, l: 14, max: 1390, rem: 480, r64: 290, r32: 220, s16: 160, e8: 240, f4max: 160, champMax: 320, champ: "Arizona", champAlive: true  },
  { rank: 2, name: "Capn'Kirk",           pts: 810, pct: 81.35, w: 43, l: 19, max: 970,  rem: 160, r64: 250, r32: 240, s16: 160, e8: 160, f4max: 160, champMax: 0,   champ: "Duke",    champAlive: false },
  { rank: 3, name: "twnutt",              pts: 750, pct: 64.89, w: 39, l: 23, max: 910,  rem: 160, r64: 230, r32: 200, s16: 160, e8: 160, f4max: 160, champMax: 0,   champ: "Duke",    champAlive: false },
  { rank: 4, name: "Wisdom TeethPicks 1", pts: 700, pct: 50.58, w: 40, l: 21, max: 1180, rem: 480, r64: 240, r32: 220, s16: 160, e8: 80,  f4max: 160, champMax: 320, champ: "Arizona", champAlive: true  },
  { rank: 5, name: "Fizzz3",              pts: 680, pct: 45.08, w: 41, l: 20, max: 1160, rem: 480, r64: 260, r32: 220, s16: 120, e8: 80,  f4max: 160, champMax: 320, champ: "Arizona", champAlive: true  },
  { rank: 6, name: "inursha",             pts: 680, pct: 45.08, w: 36, l: 27, max: 680,  rem: 0,   r64: 200, r32: 200, s16: 200, e8: 80,  f4max: 0,   champMax: 0,   champ: "Duke",    champAlive: false },
  { rank: 7, name: "DIP Harambe",         pts: 670, pct: 42.43, w: 35, l: 27, max: 830,  rem: 160, r64: 230, r32: 120, s16: 160, e8: 160, f4max: 160, champMax: 0,   champ: "Duke",    champAlive: false },
  { rank: 8, name: "Rike Myan",           pts: 520, pct: 15.80, w: 35, l: 28, max: 520,  rem: 0,   r64: 240, r32: 160, s16: 120, e8: 0,   f4max: 0,   champMax: 0,   champ: "Duke",    champAlive: false },
];

const TOTAL_SLIDES = 12;
const groupAvg = Math.round(standings.reduce((s, p) => s + p.pts, 0) / standings.length);
const updateTime = "Mar 29, 2026 — Elite 8 Complete";
const alive = standings.filter(p => p.champAlive).length;
const busted = standings.filter(p => !p.champAlive).length;

// Champion colors
const champColor = { "Arizona": C.red, "Duke": C.navy };

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
  s.addText("ESPN Tournament Challenge  /  Bracket Analytics  /  Mar 29, 2026", {
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
  addTitle(s, 1, "TaylorDuran Dominates by 100 \u2014 Duke Collapse Reshapes the Race", 2);

  const kpis = [
    { label: "CURRENT LEADER", value: "910",            sub: "TaylorDuran",      accent: C.red   },
    { label: "LEAD MARGIN",    value: "+100",            sub: "vs Capn'Kirk",     accent: C.green },
    { label: "GROUP AVERAGE",  value: String(groupAvg),  sub: "8 entries",        accent: C.gold  },
    { label: "CHAMP PICKS",   value: `${alive} / 8`,    sub: `${busted} busted`, accent: C.busted },
    { label: "HIGHEST CEILING",value: "1390",            sub: "TaylorDuran",      accent: C.green },
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
    "Duke's elimination crushed 5 of 8 brackets. TaylorDuran (Arizona) leads by 100 pts with both the highest score AND highest ceiling. Only 3 champion picks (all Arizona) are still alive. Final Four: Arizona vs Michigan, UConn vs Illinois.",
    { x: 0.5, y: ky + 0.28, w: 5.6, h: 0.7, fontSize: 9, fontFace: "Calibri", color: C.bodyText }
  );

  // TOP 8 bar chart (right side)
  s.addText("TOP 8", { x: 6.4, y: 2.4, w: 1, h: 0.22, fontSize: 8, fontFace: "Calibri", color: C.darkText, bold: true });

  standings.slice(0, 8).forEach((p, i) => {
    const ry = 2.7 + i * 0.3;
    const maxBarW = 1.5;
    const bw = (p.pts / 950) * maxBarW;
    const isTop3 = i < 3;
    const bColor = p.champAlive ? C.gold : C.blueBar;

    s.addText(String(i + 1), { x: 6.4, y: ry, w: 0.2, h: 0.28, fontSize: 8, fontFace: "Calibri", color: isTop3 ? C.red : C.bodyText, bold: true, valign: "middle" });
    s.addText(p.name, { x: 6.6, y: ry, w: 1.3, h: 0.28, fontSize: 7.5, fontFace: "Calibri", color: C.darkText, bold: isTop3, valign: "middle" });

    s.addShape(pres.shapes.RECTANGLE, { x: 7.95, y: ry + 0.04, w: bw, h: 0.2, fill: { color: bColor } });
    s.addText(`${p.pts}`, { x: 7.95 + bw + 0.04, y: ry, w: 0.6, h: 0.28, fontSize: 7, fontFace: "Calibri", color: C.bodyText, valign: "middle" });
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 3 — COMPLETE STANDINGS TABLE
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 2, "Complete Standings \u2014 8 Brackets", 3);

  const hOpt = (a = "center") => ({ fill: { color: C.tableHead }, color: C.white, bold: true, fontSize: 8, fontFace: "Calibri", align: a, valign: "middle" });

  const header = [
    { text: "RK",       options: hOpt() },
    { text: "OWNER",    options: hOpt("left") },
    { text: "PTS",      options: hOpt() },
    { text: "MAX",      options: hOpt() },
    { text: "W-L",      options: hOpt() },
    { text: "R64",      options: hOpt() },
    { text: "R32",      options: hOpt() },
    { text: "S16",      options: hOpt() },
    { text: "E8",       options: hOpt() },
    { text: "CHAMP",    options: hOpt("left") },
    { text: "STATUS",   options: hOpt() },
  ];

  const rows = standings.map((p, i) => {
    const rowFill = i % 2 === 0 ? C.creamRow : C.white;
    const b = (a = "center") => ({ fill: { color: rowFill }, fontSize: 8.5, fontFace: "Calibri", color: C.darkText, align: a, valign: "middle" });
    const statusText = p.champAlive ? "ALIVE" : "BUSTED";
    const statusColor = p.champAlive ? C.green : C.busted;

    return [
      { text: String(p.rank), options: { ...b(), bold: true, color: i < 3 ? C.red : C.darkText } },
      { text: p.name,         options: { ...b("left"), bold: i < 3 } },
      { text: String(p.pts),  options: { ...b(), bold: true } },
      { text: String(p.max),  options: { ...b(), bold: true, color: p.max === p.pts ? C.busted : C.gold } },
      { text: `${p.w}-${p.l}`,options: b() },
      { text: String(p.r64),  options: b() },
      { text: String(p.r32),  options: b() },
      { text: String(p.s16),  options: b() },
      { text: String(p.e8),   options: b() },
      { text: p.champ,        options: b("left") },
      { text: statusText,     options: { ...b(), bold: true, color: statusColor, fontSize: 8 } },
    ];
  });

  s.addTable([header, ...rows], {
    x: 0.3, y: 1.0, w: 9.4,
    colW: [0.35, 1.55, 0.5, 0.5, 0.55, 0.45, 0.45, 0.45, 0.45, 0.9, 0.65],
    rowH: [0.35, ...Array(8).fill(0.45)],
    border: { pt: 0.3, color: C.lightGray },
    margin: [2, 4, 2, 4]
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
    { t: "Duke's Collapse Changes Everything", b: "All 5 Duke pickers (Capn'Kirk, twnutt, inursha, DIP Harambe, Rike Myan) lost their 320-pt championship bonus. Two (inursha, Rike Myan) are ceiling-locked with zero upside remaining." },
    { t: "TaylorDuran Has a Stranglehold on This Group", b: "Leading by 100 pts with the highest ceiling (1390), TaylorDuran wins the group in every scenario where Arizona reaches the championship game. Even if Arizona loses in the F4, Taylor leads by 100+ over most." },
    { t: "Arizona in the Final Four Is the Biggest Swing Factor", b: "If Arizona wins it all, the 3 Arizona pickers (TaylorDuran, Wisdom TeethPicks, Fizzz3) each get 320 pts. No Duke picker can match that. TaylorDuran's projected 1390 is unreachable." },
    { t: "Capn'Kirk's Only Path: Nail the F4 Pick + Pray Arizona Loses", b: "At 810 pts with max 970, Capn'Kirk needs their remaining F4 pick (160 pts) AND needs TaylorDuran to score 0 more. Even then, 970 vs 910 is razor thin." },
    { t: "Final Four: Arizona vs Michigan, UConn vs Illinois", b: "F4 = 160 pts/game, Championship = 320 pts. Arizona is the only live champion pick in this group. The F4 games determine who gets the last 480 possible points." },
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
  addTitle(s, 4, "TaylorDuran Leads by 100 \u2014 E8 Domination (240 pts)", 5);

  const sorted = [...standings].sort((a, b) => a.pts - b.pts);

  sorted.forEach((p, i) => {
    const ry = 1.1 + i * 0.5;
    const maxBarW = 5.5;
    const bw = (p.pts / 950) * maxBarW;
    const color = p.champAlive ? C.gold : C.blueBar;

    s.addText(p.name, { x: 0.5, y: ry, w: 1.8, h: 0.4, fontSize: 9, fontFace: "Calibri", color: C.darkText, align: "right", valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 2.4, y: ry + 0.06, w: bw, h: 0.28, fill: { color } });
    s.addText(String(p.pts), { x: 2.4 + bw + 0.08, y: ry, w: 0.6, h: 0.4, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
    if (!p.champAlive) {
      s.addText("\u2716", { x: 2.4 + bw + 0.55, y: ry, w: 0.3, h: 0.4, fontSize: 8, fontFace: "Calibri", color: C.busted, valign: "middle" });
    }
  });

  addFooter(s, 5);
}

// ════════════════════════════════════════════════════════
// SLIDE 6 — ROUND BY ROUND (stacked bars — now 4 rounds)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 5, "TaylorDuran Got 3 of 4 Elite 8 Picks \u2014 Best E8 Score (240)", 6);

  const sorted = [...standings].sort((a, b) => (a.r64 + a.r32 + a.s16 + a.e8) - (b.r64 + b.r32 + b.s16 + b.e8));

  s.addChart(pres.charts.BAR, [
    { name: "R64", labels: sorted.map(p => p.name), values: sorted.map(p => p.r64) },
    { name: "R32", labels: sorted.map(p => p.name), values: sorted.map(p => p.r32) },
    { name: "S16", labels: sorted.map(p => p.name), values: sorted.map(p => p.s16) },
    { name: "E8",  labels: sorted.map(p => p.name), values: sorted.map(p => p.e8) },
  ], {
    x: 0.3, y: 1.0, w: 9.4, h: 4.1,
    barDir: "bar",
    barGrouping: "stacked",
    chartColors: [C.blueBar, C.green, C.gold, C.red],
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

  addTitle(s, 6, `Only ${alive} of 8 Champion Picks Still Alive \u2014 Duke Is Out`, 7);

  // Bar chart left
  const sortedC = [...champList].sort((a, b) => a[1] - b[1]);
  sortedC.forEach((entry, i) => {
    const ry = 1.8 + i * 0.8;
    const maxBarW = 3.8;
    const bw = (entry[1] / 6) * maxBarW;
    const color = champColor[entry[0]] || C.blueBar;

    s.addText(entry[0], { x: 0.5, y: ry, w: 1.2, h: 0.45, fontSize: 10, fontFace: "Calibri", color: C.darkText, align: "right", valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.8, y: ry + 0.08, w: bw, h: 0.3, fill: { color } });
    s.addText(String(entry[1]), { x: 1.8 + bw + 0.08, y: ry, w: 0.4, h: 0.45, fontSize: 11, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
  });

  // CHAMPION STATUS table right
  s.addText("CHAMPION STATUS", { x: 5.8, y: 1.5, w: 3.7, h: 0.3, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true });

  const statusData = [
    { team: "Arizona", count: 3, alive: true },
    { team: "Duke", count: 5, alive: false },
  ];
  let ty = 1.9;
  statusData.forEach(entry => {
    const color = champColor[entry.team] || C.navy;
    const statusText = entry.alive ? "IN THE\nHUNT" : "ELIMINATED";
    const statusColor = entry.alive ? C.green : C.busted;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: ty, w: 3.7, h: 0.48, fill: { color: C.white }, shadow: mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: ty, w: 0.05, h: 0.48, fill: { color } });
    s.addText(entry.team, { x: 6.0, y: ty, w: 1.4, h: 0.48, fontSize: 10, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });
    s.addText(`${entry.count} pick${entry.count > 1 ? "s" : ""}`, { x: 7.4, y: ty, w: 0.9, h: 0.48, fontSize: 9, fontFace: "Calibri", color: C.bodyText, valign: "middle", align: "center" });
    s.addText(statusText, { x: 8.4, y: ty, w: 1.0, h: 0.48, fontSize: 8, fontFace: "Calibri", color: statusColor, bold: true, valign: "middle", align: "right" });
    ty += 0.56;
  });

  // Extra context
  s.addText("Duke fell in the Elite 8. All 5 Duke pickers lose their 320-pt championship bonus.", {
    x: 0.5, y: 3.6, w: 9, h: 0.3,
    fontSize: 9, fontFace: "Calibri", color: C.bodyText, italic: true
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 8 — MAX CEILING (with busted status)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 7, `Only ${alive} Champion Picks Alive \u2014 ${busted} Are Busted`, 8);

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
    const bColor = p.champAlive ? (champColor[p.champ] || C.blueBar) : C.muted;

    s.addText(String(p.rank), { x: 0.5, y: ry, w: 0.35, h: 0.38, fontSize: 13, fontFace: "Calibri", color: p.champAlive ? C.red : C.muted, bold: true, valign: "middle" });
    s.addText(p.name, { x: 0.85, y: ry, w: 1.6, h: 0.38, fontSize: 10, fontFace: "Calibri", color: C.darkText, bold: true, valign: "middle" });

    // Max bar
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6, y: ry + 0.06, w: maxW, h: 0.26, fill: { color: C.lightGray } });
    // Current bar
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6, y: ry + 0.06, w: currentW, h: 0.26, fill: { color: bColor } });
    // Label
    s.addText(`${p.pts} / ${p.max}`, { x: 2.6 + maxW + 0.12, y: ry, w: 1.2, h: 0.38, fontSize: 9.5, fontFace: "Calibri", color: C.bodyText, valign: "middle" });

    // Busted tag
    if (!p.champAlive && p.max === p.pts) {
      s.addText("LOCKED", { x: 2.6 + maxW + 1.3, y: ry, w: 0.7, h: 0.38, fontSize: 7, fontFace: "Calibri", color: C.busted, bold: true, valign: "middle" });
    } else if (!p.champAlive) {
      s.addText("NO CHAMP", { x: 2.6 + maxW + 1.3, y: ry, w: 0.8, h: 0.38, fontSize: 7, fontFace: "Calibri", color: C.busted, bold: true, valign: "middle" });
    }
  });

  s.addText(`${busted} brackets with a busted champion pick`, {
    x: 0.5, y: 5.0, w: 9, h: 0.2,
    fontSize: 8, fontFace: "Calibri", color: C.bodyText, italic: true
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 9 — PATHS TO WIN (from FCC reference)
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  addTitle(s, 8, "3 Arizona Picks, 5 Ceilings Capped \u2014 Here\u2019s How It Plays Out", 9);

  s.addText("Projections include Final Four points (160/game) and championship (320 pts) based on each player\u2019s remaining live picks.", {
    x: 0.5, y: 0.8, w: 9.5, h: 0.25,
    fontSize: 8.5, fontFace: "Calibri", color: C.bodyText, italic: true
  });

  // -- Column 1: ARIZONA WINS --
  const col1x = 0.5;
  s.addShape(pres.shapes.RECTANGLE, { x: col1x, y: 1.2, w: 4.3, h: 0.32, fill: { color: C.red } });
  s.addText("ARIZONA WINS", { x: col1x, y: 1.2, w: 4.3, h: 0.32, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle" });

  // Header row
  const headerY = 1.6;
  s.addText("Player", { x: col1x, y: headerY, w: 1.8, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true });
  s.addText("Now", { x: col1x + 1.8, y: headerY, w: 0.6, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });
  s.addText("Proj.", { x: col1x + 2.4, y: headerY, w: 0.6, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });
  s.addText("Upside", { x: col1x + 3.0, y: headerY, w: 0.7, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });

  // Arizona wins: sort by max (includes championship)
  const azWins = [...standings].map(p => ({
    name: p.name,
    now: p.pts,
    proj: p.max,
    upside: p.max - p.pts,
  })).sort((a, b) => b.proj - a.proj);

  azWins.forEach((p, i) => {
    const ry = 1.9 + i * 0.3;
    const isAZ = standings.find(s => s.name === p.name).champAlive;
    s.addText(`${i + 1}. ${p.name}`, { x: col1x, y: ry, w: 1.8, h: 0.28, fontSize: 8.5, fontFace: "Calibri", color: isAZ ? C.red : C.darkText, bold: isAZ, valign: "middle" });
    s.addText(String(p.now), { x: col1x + 1.8, y: ry, w: 0.6, h: 0.28, fontSize: 9, fontFace: "Calibri", color: C.darkText, align: "center", valign: "middle" });
    s.addText(String(p.proj), { x: col1x + 2.4, y: ry, w: 0.6, h: 0.28, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true, align: "center", valign: "middle" });
    s.addText(`+${p.upside}`, { x: col1x + 3.0, y: ry, w: 0.7, h: 0.28, fontSize: 9, fontFace: "Calibri", color: p.upside > 0 ? C.green : C.muted, bold: true, align: "center", valign: "middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: col1x, y: 4.35, w: 4.3, h: 0.3, fill: { color: "FBF0ED" } });
  s.addText("TaylorDuran wins by 210+", { x: col1x, y: 4.35, w: 4.3, h: 0.3, fontSize: 9, fontFace: "Calibri", color: C.red, bold: true, align: "center", valign: "middle" });

  // -- Column 2: ARIZONA LOSES IN F4 --
  const col2x = 5.2;
  s.addShape(pres.shapes.RECTANGLE, { x: col2x, y: 1.2, w: 4.3, h: 0.32, fill: { color: C.navy } });
  s.addText("ARIZONA LOSES IN F4", { x: col2x, y: 1.2, w: 4.3, h: 0.32, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle" });

  s.addText("Player", { x: col2x, y: headerY, w: 1.8, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true });
  s.addText("Now", { x: col2x + 1.8, y: headerY, w: 0.6, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });
  s.addText("Max", { x: col2x + 2.4, y: headerY, w: 0.6, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });
  s.addText("Upside", { x: col2x + 3.0, y: headerY, w: 0.7, h: 0.25, fontSize: 8, fontFace: "Calibri", color: C.muted, bold: true, align: "center" });

  // Arizona loses: no champ bonus for anyone. F4 picks still possible.
  const azLoses = [...standings].map(p => ({
    name: p.name,
    now: p.pts,
    max: p.pts + p.f4max, // only F4 points remain, no championship
    upside: p.f4max,
  })).sort((a, b) => b.max - a.max);

  azLoses.forEach((p, i) => {
    const ry = 1.9 + i * 0.3;
    s.addText(`${i + 1}. ${p.name}`, { x: col2x, y: ry, w: 1.8, h: 0.28, fontSize: 8.5, fontFace: "Calibri", color: C.darkText, valign: "middle" });
    s.addText(String(p.now), { x: col2x + 1.8, y: ry, w: 0.6, h: 0.28, fontSize: 9, fontFace: "Calibri", color: C.darkText, align: "center", valign: "middle" });
    s.addText(String(p.max), { x: col2x + 2.4, y: ry, w: 0.6, h: 0.28, fontSize: 9, fontFace: "Calibri", color: C.darkText, bold: true, align: "center", valign: "middle" });
    s.addText(p.upside > 0 ? `+${p.upside}` : "LOCKED", { x: col2x + 3.0, y: ry, w: 0.7, h: 0.28, fontSize: p.upside > 0 ? 9 : 7, fontFace: "Calibri", color: p.upside > 0 ? C.green : C.busted, bold: true, align: "center", valign: "middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: col2x, y: 4.35, w: 4.3, h: 0.3, fill: { color: "EDF0FB" } });
  s.addText("TaylorDuran still leads \u2014 only Capn'Kirk can catch", { x: col2x, y: 4.35, w: 4.3, h: 0.3, fontSize: 9, fontFace: "Calibri", color: C.navy, bold: true, align: "center", valign: "middle" });

  // Bottom summary
  s.addText("TaylorDuran wins the group in every scenario where Arizona reaches the championship game. Even if Arizona loses in the F4, Taylor's 100-point cushion is nearly insurmountable.", {
    x: 0.5, y: 4.8, w: 9.5, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: C.bodyText, italic: true
  });
}

// ════════════════════════════════════════════════════════
// SLIDE 10 — CLOSING: "Game On."
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("Game On.", {
    x: 0.7, y: 2.4, w: 8.6, h: 1.2,
    fontSize: 64, fontFace: "Calibri", color: C.red, bold: true
  });
  s.addText("The Final Four awaits.", {
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
// SLIDE 11 — APPENDIX DIVIDER
// ════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.tealBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("Appendix", {
    x: 0.7, y: 2.2, w: 8.6, h: 1.0,
    fontSize: 44, fontFace: "Calibri", color: C.white, bold: true
  });
  s.addText("Deeper analytics: round accuracy, momentum, path scenarios", {
    x: 0.7, y: 3.2, w: 8.6, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: C.muted
  });
  addFooter(s, 11);
}

// ════════════════════════════════════════════════════════
// SLIDE 12 — LIVE TRACKER
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
