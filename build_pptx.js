/**
 * CANONICAL BUILD SCRIPT — Elons Plugs 2026 March Madness Bracket Tracker
 *
 * This is the single source of truth for generating the PPTX deck.
 * Both the scheduled task and manual rebuilds should use this script.
 *
 * Usage: node canonical_build_pptx.js
 * Output: Elons_Plugs_Deck.pptx in current directory
 *
 * Dependencies: npm install -g pptxgenjs
 */

const pptxgen = require("pptxgenjs");

// ============================================================
// COLOR PALETTE — matches the approved 7:30 PM visual style
// ============================================================
const C = {
  navy:      '34576E',
  darkNavy:  '2D4F63',
  red:       'DC3545',
  gold:      'CDA050',
  cream:     'EDEBE8',
  lightCream:'F5F3F0',
  white:     'FFFFFF',
  dark:      '2D3748',
  muted:     '6B7280',
  tableAlt:  'F0EEEB',
  greenText: '16A34A',
  redText:   'DC3545',
  goldText:  'B8860B',
};

// ============================================================
// DATA — Updated March 22, 2026 8:30 PM ET (Tennessee 79-72 Virginia)
// ============================================================
const UPDATE_TIME = 'March 22, 2026 9:17 PM ET';
const ROUND_STATUS = 'Round of 32 in Progress';
const GAMES_REMAINING = '4 Games Remaining';

const standings = [
  { rank: 1, name: 'TaylorDuran',          pts: 450, pct: 97.46, wl: '37W-8L',  max: 1770, r64: 290, r32: 160, champ: 'Arizona Wildcats' },
  { rank: 2, name: 'Fizzz3',               pts: 440, pct: 79.86, wl: '35W-10L', max: 1760, r64: 260, r32: 180, champ: 'Arizona Wildcats' },
  { rank: 3, name: "Capn'Kirk",            pts: 430, pct: 87.97, wl: '34W-10L', max: 1790, r64: 250, r32: 180, champ: 'Duke Blue Devils' },
  { rank: 4, name: 'Paul233365',           pts: 400, pct: 60.95, wl: '32W-11L', max: 1780, r64: 240, r32: 160, champ: 'Arizona Wildcats' },
  { rank: 5, name: 'twnutt',               pts: 390, pct: 52.11, wl: '31W-13L', max: 1750, r64: 230, r32: 160, champ: 'Duke Blue Devils' },
  { rank: 6, name: 'inursha',              pts: 340, pct: 20.98, wl: '27W-17L', max: 1680, r64: 200, r32: 140, champ: 'Duke Blue Devils' },
  { rank: 6, name: 'ESPNFAN8289368577',    pts: 340, pct: 20.98, wl: '29W-18L', max: 1400, r64: 240, r32: 100, champ: 'Duke Blue Devils' },
  { rank: 8, name: 'ESPNFAN9699154365',    pts: 310, pct: 13.27, wl: '27W-17L', max: 1650, r64: 230, r32: 80,  champ: 'Duke Blue Devils' },
];

const keyChanges = [
  "Fizzz3 surges +20 to 440 pts (Tennessee R32 win) — leapfrogs Capn'Kirk for #2",
  "TaylorDuran lead shrinks to 10 pts over Fizzz3, 20 over Capn'Kirk",
  "Capn'Kirk +1L (had Virginia), TaylorDuran +1L — both picked Virginia to advance",
  "4 R32 games remain: Texas Tech, UCLA, Iowa, Utah State matchups still undecided",
];

const contrarianPicks = [
  { player: 'inursha',       team: 'Northern Iowa',  round: 'R64', groupPct: '13%', result: 'WRONG' },
  { player: 'inursha',       team: "Hawai'i",        round: 'R64', groupPct: '13%', result: 'WRONG' },
  { player: 'ESPN...4365',   team: 'Kennesaw State', round: 'R64', groupPct: '13%', result: 'WRONG' },
  { player: 'ESPN...8577',   team: 'Louisville',     round: 'R32', groupPct: '25%', result: 'WRONG' },
  { player: 'ESPN...4365',   team: 'South Florida',  round: 'R32', groupPct: '13%', result: 'WRONG' },
  { player: 'ESPN...4365',   team: 'UCLA',           round: 'R32', groupPct: '25%', result: 'UNDECIDED' },
  { player: 'Fizzz3',        team: 'Nebraska',       round: 'R32', groupPct: '25%', result: 'CORRECT' },
  { player: 'ESPN...4365',   team: 'VCU',            round: 'R32', groupPct: '25%', result: 'WRONG' },
  { player: 'twnutt',        team: 'Texas',          round: 'R32', groupPct: '25%', result: 'WRONG' },
  { player: 'Fizzz3',        team: 'Texas Tech',     round: 'R32', groupPct: '25%', result: 'UNDECIDED' },
  { player: 'ESPN...4365',   team: 'Tennessee',      round: 'R32', groupPct: '25%', result: 'CORRECT' },
  { player: 'ESPN...8577',   team: 'Ole Miss',       round: 'R32', groupPct: '25%', result: 'WRONG' },
  { player: 'ESPN...4365',   team: 'Iowa',           round: 'R32', groupPct: '25%', result: 'UNDECIDED' },
  { player: 'Fizzz3',        team: 'Utah State',     round: 'R32', groupPct: '25%', result: 'UNDECIDED' },
];

const DASHBOARD_URL = 'tdswim.github.io/Elons-Plugs-2026';

// ============================================================
// HELPER — truncate long names for chart labels
// ============================================================
function shortName(name, maxLen = 12) {
  if (name.length <= maxLen) return name;
  if (name.startsWith('ESPNFAN')) return 'ESPN...' + name.slice(-4);
  return name.slice(0, maxLen);
}

// ============================================================
// BUILD PRESENTATION
// ============================================================
async function build() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.author = 'Elons Plugs Bracket Tracker';
  pres.title = 'Elons Plugs 2026 — March Madness Bracket Tracker';

  // Group stats
  const totalPts = standings.reduce((s, p) => s + p.pts, 0);
  const avgPts = Math.round(totalPts / standings.length);
  const leader = standings[0];
  const runnerUp = standings[1];
  const leadMargin = leader.pts - runnerUp.pts;

  // ----------------------------------------------------------
  // SLIDE 1: TITLE
  // ----------------------------------------------------------
  const s1 = pres.addSlide();
  s1.background = { color: C.navy };
  // Red bars top and bottom
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.12, fill: { color: C.red } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.505, w: 10, h: 0.12, fill: { color: C.red } });

  s1.addText('ELONS PLUGS 2026', {
    x: 0.5, y: 1.0, w: 9, h: 1.2,
    fontSize: 44, fontFace: 'Arial Black', color: C.white,
    bold: true, align: 'center', charSpacing: 3,
  });
  s1.addText('MARCH MADNESS BRACKET TRACKER', {
    x: 0.5, y: 2.2, w: 9, h: 0.6,
    fontSize: 18, fontFace: 'Arial', color: C.gold,
    bold: true, align: 'center', charSpacing: 4,
  });
  // Red divider line
  s1.addShape(pres.shapes.LINE, {
    x: 3.5, y: 2.95, w: 3, h: 0,
    line: { color: C.red, width: 2.5 },
  });
  s1.addText(`Men's Division  •  Updated ${UPDATE_TIME}`, {
    x: 0.5, y: 3.3, w: 9, h: 0.4,
    fontSize: 13, fontFace: 'Calibri', color: C.white, align: 'center',
  });
  s1.addText([
    { text: ROUND_STATUS, options: { color: C.gold } },
    { text: `  •  ${GAMES_REMAINING}`, options: { color: C.gold } },
  ], {
    x: 0.5, y: 3.85, w: 9, h: 0.4,
    fontSize: 13, fontFace: 'Calibri', align: 'center',
  });

  // ----------------------------------------------------------
  // SLIDE 2: EXECUTIVE SUMMARY
  // ----------------------------------------------------------
  const s2 = pres.addSlide();
  s2.background = { color: C.cream };
  // Navy header bar
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s2.addText('01  EXECUTIVE SUMMARY', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  // KPI Cards — 4 across
  const kpis = [
    { label: 'LEADER', value: leader.name, sub: `${leader.pts} pts • ${leader.pct}th percentile` },
    { label: 'LEAD MARGIN', value: `${leadMargin} pts`, sub: `vs ${runnerUp.name} (${runnerUp.pts})` },
    { label: 'GROUP AVG', value: `${avgPts} pts`, sub: `${standings.length} total entries` },
    { label: 'NATIONAL %ILE', value: `${leader.pct}%`, sub: `${leader.name} ranking` },
  ];

  const cardW = 2.05;
  const cardGap = 0.2;
  const cardStartX = 0.6;
  const cardY = 1.1;
  const cardH = 1.5;

  kpis.forEach((kpi, i) => {
    const cx = cardStartX + i * (cardW + cardGap);
    // Card background
    s2.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.white },
      shadow: { type: 'outer', blur: 4, offset: 1, angle: 135, color: '000000', opacity: 0.08 },
    });
    // Red top accent
    s2.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.05, fill: { color: C.red },
    });
    // Label
    s2.addText(kpi.label, {
      x: cx + 0.15, y: cardY + 0.15, w: cardW - 0.3, h: 0.3,
      fontSize: 9, fontFace: 'Calibri', color: C.muted, bold: true, margin: 0,
    });
    // Value
    s2.addText(kpi.value, {
      x: cx + 0.15, y: cardY + 0.45, w: cardW - 0.3, h: 0.5,
      fontSize: 22, fontFace: 'Calibri', color: C.dark, bold: true, margin: 0,
    });
    // Sub
    s2.addText(kpi.sub, {
      x: cx + 0.15, y: cardY + 1.0, w: cardW - 0.3, h: 0.3,
      fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0,
    });
  });

  // Key Changes box
  const changesY = 2.95;
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: changesY, w: 8.8, h: 2.2,
    fill: { color: C.white },
    shadow: { type: 'outer', blur: 4, offset: 1, angle: 135, color: '000000', opacity: 0.08 },
  });
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: changesY, w: 8.8, h: 0.05, fill: { color: C.red },
  });
  s2.addText('KEY CHANGES SINCE LAST UPDATE', {
    x: 0.85, y: changesY + 0.15, w: 8, h: 0.35,
    fontSize: 12, fontFace: 'Calibri', color: C.red, bold: true, margin: 0,
  });

  const bulletItems = keyChanges.map((text, idx) => ({
    text: text,
    options: {
      bullet: true,
      breakLine: idx < keyChanges.length - 1,
      fontSize: 11,
      fontFace: 'Calibri',
      color: C.dark,
      paraSpaceAfter: 6,
    },
  }));
  s2.addText(bulletItems, {
    x: 0.85, y: changesY + 0.55, w: 8.3, h: 1.5, valign: 'top', margin: 0,
  });

  // ----------------------------------------------------------
  // SLIDE 3: COMPLETE STANDINGS TABLE
  // ----------------------------------------------------------
  const s3 = pres.addSlide();
  s3.background = { color: C.cream };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s3.addText('02  COMPLETE STANDINGS', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  const headers = ['RANK', 'PLAYER', 'PTS', 'PCT', 'W-L', 'MAX', 'R64', 'R32', 'CHAMPION'];
  const colW = [0.55, 1.45, 0.55, 0.7, 0.8, 0.6, 0.55, 0.55, 1.6];
  const headerOpts = {
    fill: { color: C.navy }, color: C.white, bold: true,
    fontSize: 8, fontFace: 'Calibri', align: 'center', valign: 'middle',
  };

  const headerRow = headers.map((h, i) => ({
    text: h,
    options: { ...headerOpts, align: i === 1 || i === 8 ? 'left' : 'center' },
  }));

  const tableRows = [headerRow];
  standings.forEach((p, idx) => {
    const isEven = idx % 2 === 0;
    const rowBg = isEven ? C.white : C.tableAlt;
    const isLeader = idx === 0;
    const nameColor = isLeader ? C.red : C.dark;

    const displayName = p.name.length > 15 ? p.name.slice(0, 13) + '...' : p.name;

    const makeCell = (text, opts = {}) => ({
      text: String(text),
      options: {
        fill: { color: rowBg },
        fontSize: 9, fontFace: 'Calibri', color: C.dark,
        valign: 'middle', ...opts,
      },
    });

    tableRows.push([
      makeCell(`#${p.rank}`, { bold: true, align: 'center' }),
      makeCell(displayName, { bold: isLeader, color: nameColor, align: 'left' }),
      makeCell(p.pts, { bold: true, align: 'center', color: isLeader ? C.red : C.dark }),
      makeCell(`${p.pct}%`, { align: 'center' }),
      makeCell(p.wl, { align: 'center' }),
      makeCell(p.max, { align: 'center', bold: true, color: C.goldText }),
      makeCell(p.r64, { align: 'center' }),
      makeCell(p.r32, { align: 'center' }),
      makeCell(p.champ, { align: 'left', fontSize: 8 }),
    ]);
  });

  s3.addTable(tableRows, {
    x: 0.6, y: 1.0, w: 8.8,
    colW: colW,
    rowH: [0.35, ...Array(standings.length).fill(0.45)],
    border: { pt: 0.5, color: 'D0CCC8' },
    margin: [2, 4, 2, 4],
  });

  // ----------------------------------------------------------
  // SLIDE 4: ROUND-BY-ROUND BREAKDOWN (stacked horizontal bars)
  // ----------------------------------------------------------
  const s4 = pres.addSlide();
  s4.background = { color: C.cream };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s4.addText('03  ROUND-BY-ROUND BREAKDOWN', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  // Build stacked bars manually for precise control
  const barStartX = 1.8;
  const barMaxW = 5.5;
  const barH = 0.35;
  const barGap = 0.22;
  const barStartY = 1.15;
  const maxPts = Math.max(...standings.map(p => p.pts));

  standings.forEach((p, i) => {
    const y = barStartY + i * (barH + barGap);
    const totalW = (p.pts / maxPts) * barMaxW;
    const r64W = (p.r64 / maxPts) * barMaxW;
    const r32W = (p.r32 / maxPts) * barMaxW;

    // Player name label
    s4.addText(shortName(p.name, 11), {
      x: 0.2, y: y, w: 1.5, h: barH,
      fontSize: 10, fontFace: 'Calibri', color: C.dark, bold: true,
      align: 'right', valign: 'middle', margin: 0,
    });

    // R64 bar (red)
    s4.addShape(pres.shapes.RECTANGLE, {
      x: barStartX, y: y, w: Math.max(r64W, 0.01), h: barH,
      fill: { color: C.red },
    });
    // R64 label inside bar
    if (r64W > 0.5) {
      s4.addText(String(p.r64), {
        x: barStartX, y: y, w: r64W, h: barH,
        fontSize: 9, fontFace: 'Calibri', color: C.white, bold: true,
        align: 'center', valign: 'middle', margin: 0,
      });
    }

    // R32 bar (navy)
    s4.addShape(pres.shapes.RECTANGLE, {
      x: barStartX + r64W, y: y, w: Math.max(r32W, 0.01), h: barH,
      fill: { color: C.navy },
    });
    // R32 label inside bar
    if (r32W > 0.4) {
      s4.addText(String(p.r32), {
        x: barStartX + r64W, y: y, w: r32W, h: barH,
        fontSize: 9, fontFace: 'Calibri', color: C.white, bold: true,
        align: 'center', valign: 'middle', margin: 0,
      });
    }

    // Total label on right
    s4.addText(`${p.pts} pts`, {
      x: barStartX + totalW + 0.15, y: y, w: 1, h: barH,
      fontSize: 10, fontFace: 'Calibri', color: C.dark,
      align: 'left', valign: 'middle', margin: 0,
    });
  });

  // Legend
  const legY = barStartY + standings.length * (barH + barGap) + 0.15;
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: legY, w: 0.25, h: 0.2, fill: { color: C.red } });
  s4.addText('R64', { x: 5.8, y: legY, w: 0.5, h: 0.2, fontSize: 9, fontFace: 'Calibri', color: C.dark, margin: 0, valign: 'middle' });
  s4.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: legY, w: 0.25, h: 0.2, fill: { color: C.navy } });
  s4.addText('R32', { x: 6.8, y: legY, w: 0.5, h: 0.2, fontSize: 9, fontFace: 'Calibri', color: C.dark, margin: 0, valign: 'middle' });

  // ----------------------------------------------------------
  // SLIDE 5: CONTRARIAN PICKS ANALYSIS
  // ----------------------------------------------------------
  const s5 = pres.addSlide();
  s5.background = { color: C.cream };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s5.addText('04  CONTRARIAN PICKS ANALYSIS', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  // Summary line
  const wrongCount = contrarianPicks.filter(p => p.result === 'WRONG').length;
  const correctCount = contrarianPicks.filter(p => p.result === 'CORRECT').length;
  const undecidedCount = contrarianPicks.filter(p => p.result === 'UNDECIDED').length;
  s5.addText(
    `${wrongCount} of ${contrarianPicks.length} Contrarian Picks Failed — ${correctCount} Hit (Nebraska, Tennessee), ${undecidedCount} Undecided`,
    {
      x: 0.6, y: 0.9, w: 8.8, h: 0.35,
      fontSize: 11, fontFace: 'Calibri', color: C.dark, bold: true, margin: 0,
    }
  );

  // Build contrarian table
  const cHeaders = ['PLAYER', 'TEAM', 'ROUND', 'GROUP %', 'RESULT'];
  const cColW = [1.5, 1.6, 0.7, 0.85, 1.1];
  const cHeaderRow = cHeaders.map(h => ({
    text: h,
    options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 8, fontFace: 'Calibri', align: 'center', valign: 'middle' },
  }));

  const cRows = [cHeaderRow];
  contrarianPicks.forEach((pick, idx) => {
    const rowBg = idx % 2 === 0 ? C.white : C.tableAlt;
    const resultColor = pick.result === 'CORRECT' ? C.greenText : pick.result === 'WRONG' ? C.redText : C.goldText;
    cRows.push([
      { text: pick.player, options: { fill: { color: rowBg }, fontSize: 8, fontFace: 'Calibri', color: C.dark, valign: 'middle', align: 'left' } },
      { text: pick.team, options: { fill: { color: rowBg }, fontSize: 8, fontFace: 'Calibri', color: C.dark, valign: 'middle', align: 'left' } },
      { text: pick.round, options: { fill: { color: rowBg }, fontSize: 8, fontFace: 'Calibri', color: C.dark, valign: 'middle', align: 'center' } },
      { text: pick.groupPct, options: { fill: { color: rowBg }, fontSize: 8, fontFace: 'Calibri', color: C.dark, valign: 'middle', align: 'center' } },
      { text: pick.result, options: { fill: { color: rowBg }, fontSize: 8, fontFace: 'Calibri', color: resultColor, bold: true, valign: 'middle', align: 'center' } },
    ]);
  });

  s5.addTable(cRows, {
    x: 1.4, y: 1.35, w: 5.75,
    colW: cColW,
    rowH: [0.3, ...Array(contrarianPicks.length).fill(0.25)],
    border: { pt: 0.5, color: 'D0CCC8' },
    margin: [1, 3, 1, 3],
  });

  // ----------------------------------------------------------
  // SLIDE 6: MAX POINTS CEILING (stacked horizontal bars)
  // ----------------------------------------------------------
  const s6 = pres.addSlide();
  s6.background = { color: C.cream };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s6.addText('05  MAX POINTS CEILING', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  const maxCeiling = Math.max(...standings.map(p => p.max));
  const mBarStartX = 2.0;
  const mBarMaxW = 5.2;
  const mBarH = 0.4;
  const mBarGap = 0.2;
  const mBarStartY = 1.1;

  standings.forEach((p, i) => {
    const y = mBarStartY + i * (mBarH + mBarGap);
    const currentW = (p.pts / maxCeiling) * mBarMaxW;
    const remainW = ((p.max - p.pts) / maxCeiling) * mBarMaxW;

    // Name label
    s6.addText(shortName(p.name, 12), {
      x: 0.2, y: y, w: 1.7, h: mBarH,
      fontSize: 10, fontFace: 'Calibri', color: C.dark, bold: true,
      align: 'right', valign: 'middle', margin: 0,
    });

    // Current points bar (red)
    s6.addShape(pres.shapes.RECTANGLE, {
      x: mBarStartX, y: y, w: Math.max(currentW, 0.01), h: mBarH,
      fill: { color: C.red },
    });

    // Remaining points bar (gold)
    s6.addShape(pres.shapes.RECTANGLE, {
      x: mBarStartX + currentW, y: y, w: Math.max(remainW, 0.01), h: mBarH,
      fill: { color: C.gold },
    });

    // Max label
    s6.addText(String(p.max), {
      x: mBarStartX + currentW + remainW + 0.15, y: y, w: 0.8, h: mBarH,
      fontSize: 11, fontFace: 'Calibri', color: C.dark, bold: true,
      align: 'left', valign: 'middle', margin: 0,
    });
  });

  // Legend
  const mLegY = mBarStartY + standings.length * (mBarH + mBarGap) + 0.1;
  s6.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: mLegY, w: 0.25, h: 0.2, fill: { color: C.red } });
  s6.addText('Current', { x: 5.5, y: mLegY, w: 0.8, h: 0.2, fontSize: 9, fontFace: 'Calibri', color: C.dark, margin: 0, valign: 'middle' });
  s6.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: mLegY, w: 0.25, h: 0.2, fill: { color: C.gold } });
  s6.addText('Remaining', { x: 6.8, y: mLegY, w: 0.9, h: 0.2, fontSize: 9, fontFace: 'Calibri', color: C.dark, margin: 0, valign: 'middle' });

  // ----------------------------------------------------------
  // SLIDE 7: CHAMPION PICK DISTRIBUTION
  // ----------------------------------------------------------
  const s7 = pres.addSlide();
  s7.background = { color: C.cream };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: C.navy } });
  s7.addText('06  CHAMPION PICK DISTRIBUTION', {
    x: 0.6, y: 0.05, w: 8, h: 0.65,
    fontSize: 22, fontFace: 'Arial Black', color: C.white, bold: true, valign: 'middle',
  });

  const arizonaPickers = standings.filter(p => p.champ === 'Arizona Wildcats').map(p => p.name);
  const dukePickers = standings.filter(p => p.champ === 'Duke Blue Devils').map(p => shortName(p.name));

  // Arizona card (left)
  const cardLeft = 1.2;
  const cardRight = 5.3;
  const champCardW = 3.5;
  const champCardY = 1.3;
  const champCardH = 3.2;

  // Arizona card
  s7.addShape(pres.shapes.RECTANGLE, {
    x: cardLeft, y: champCardY, w: champCardW, h: champCardH,
    fill: { color: C.white },
    shadow: { type: 'outer', blur: 4, offset: 1, angle: 135, color: '000000', opacity: 0.08 },
  });
  // Red header
  s7.addShape(pres.shapes.RECTANGLE, {
    x: cardLeft, y: champCardY, w: champCardW, h: 0.55,
    fill: { color: C.red },
  });
  s7.addText('Arizona Wildcats', {
    x: cardLeft, y: champCardY, w: champCardW, h: 0.55,
    fontSize: 16, fontFace: 'Calibri', color: C.white, bold: true,
    align: 'center', valign: 'middle',
  });
  s7.addText(`${arizonaPickers.length} picks`, {
    x: cardLeft, y: champCardY + 0.7, w: champCardW, h: 0.7,
    fontSize: 32, fontFace: 'Calibri', color: C.red, bold: true,
    align: 'center', valign: 'middle',
  });
  s7.addText(arizonaPickers.join(', '), {
    x: cardLeft + 0.2, y: champCardY + 1.5, w: champCardW - 0.4, h: 0.7,
    fontSize: 11, fontFace: 'Calibri', color: C.dark,
    align: 'center', valign: 'middle',
  });
  s7.addText('STATUS: ALIVE', {
    x: cardLeft, y: champCardY + 2.4, w: champCardW, h: 0.5,
    fontSize: 13, fontFace: 'Calibri', color: C.greenText, bold: true,
    align: 'center', valign: 'middle',
  });

  // Duke card (right)
  s7.addShape(pres.shapes.RECTANGLE, {
    x: cardRight, y: champCardY, w: champCardW, h: champCardH,
    fill: { color: C.white },
    shadow: { type: 'outer', blur: 4, offset: 1, angle: 135, color: '000000', opacity: 0.08 },
  });
  s7.addShape(pres.shapes.RECTANGLE, {
    x: cardRight, y: champCardY, w: champCardW, h: 0.55,
    fill: { color: C.navy },
  });
  s7.addText('Duke Blue Devils', {
    x: cardRight, y: champCardY, w: champCardW, h: 0.55,
    fontSize: 16, fontFace: 'Calibri', color: C.white, bold: true,
    align: 'center', valign: 'middle',
  });
  s7.addText(`${dukePickers.length} picks`, {
    x: cardRight, y: champCardY + 0.7, w: champCardW, h: 0.7,
    fontSize: 32, fontFace: 'Calibri', color: C.navy, bold: true,
    align: 'center', valign: 'middle',
  });
  s7.addText(dukePickers.join(', '), {
    x: cardRight + 0.2, y: champCardY + 1.5, w: champCardW - 0.4, h: 0.7,
    fontSize: 11, fontFace: 'Calibri', color: C.dark,
    align: 'center', valign: 'middle', isWordWrap: true,
  });
  s7.addText('STATUS: ALIVE', {
    x: cardRight, y: champCardY + 2.4, w: champCardW, h: 0.5,
    fontSize: 13, fontFace: 'Calibri', color: C.greenText, bold: true,
    align: 'center', valign: 'middle',
  });

  // ----------------------------------------------------------
  // SLIDE 8: CLOSING / DASHBOARD LINK
  // ----------------------------------------------------------
  const s8 = pres.addSlide();
  s8.background = { color: C.navy };
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.12, fill: { color: C.red } });
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.505, w: 10, h: 0.12, fill: { color: C.red } });

  s8.addText('ELONS PLUGS 2026', {
    x: 0.5, y: 1.2, w: 9, h: 1.0,
    fontSize: 40, fontFace: 'Arial Black', color: C.white,
    bold: true, align: 'center', charSpacing: 5,
  });
  // Gold divider
  s8.addShape(pres.shapes.LINE, {
    x: 3.5, y: 2.45, w: 3, h: 0,
    line: { color: C.gold, width: 2.5 },
  });
  s8.addText('Live Dashboard', {
    x: 0.5, y: 2.7, w: 9, h: 0.5,
    fontSize: 18, fontFace: 'Calibri', color: C.gold, bold: true, italic: true,
    align: 'center',
  });
  s8.addText(DASHBOARD_URL, {
    x: 0.5, y: 3.2, w: 9, h: 0.4,
    fontSize: 13, fontFace: 'Calibri', color: C.white, align: 'center',
    hyperlink: { url: `https://${DASHBOARD_URL}` },
  });
  s8.addText(`Next update: After Round of 32 concludes`, {
    x: 0.5, y: 4.0, w: 9, h: 0.4,
    fontSize: 12, fontFace: 'Calibri', color: C.white, align: 'center',
  });

  // ----------------------------------------------------------
  // WRITE FILE
  // ----------------------------------------------------------
  await pres.writeFile({ fileName: 'Elons_Plugs_Deck.pptx' });
  console.log('✅ Deck generated: Elons_Plugs_Deck.pptx');
}

build().catch(err => { console.error('Build failed:', err); process.exit(1); });
