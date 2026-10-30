const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply_theme.js");
const CH = "../Charts/";
const OUT = "../Rexel valuation - deck.pptx";

const THEME = { name: "Rexel valuation", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1A1A1A", lt1: "FFFFFF", dk2: "0B2545", lt2: "F2F4F7", accent1: "C8692B", accent2: "4F7CAC", accent3: "8A8F98", accent4: "1F9D73", accent5: "D64545", accent6: "E9B44C", hlink: "4F7CAC", folHlink: "8A8F98" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Marwan Ferreira da Silva"; pres.title = "Rexel - Valuation of a listed company";
const C = pres.SchemeColor;
const NAVY = THEME.colors.dk2, COPPER = THEME.colors.accent1, STEEL = THEME.colors.accent2, GREY = THEME.colors.accent3, LT2 = THEME.colors.lt2, GREEN = THEME.colors.accent4, RED = THEME.colors.accent5;
const FOOT = "Rexel – valuation of a listed company  |  Dauphine M2 104  |  Valuation date 30 October 2026";

pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontSize: 20, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.28, w: 7.5, h: 0.25, fontSize: 8, color: C.accent3, margin: 0 } } },
  ],
  slideNumber: { x: 8.9, y: 5.28, w: 0.6, h: 0.25, fontSize: 8, color: C.accent3, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { color: NAVY },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } }],
  slideNumber: { x: 8.9, y: 5.28, w: 0.6, h: 0.25, fontSize: 8, color: C.accent3, align: "right" } });

// helpers --------------------------------------------------------------
const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox: true, fontSize: 11, color: C.text1, margin: 0, valign: "top" }, o));
const bullets = (s, items, o) => s.addText(items.map((t, i) => (typeof t === "string" ? { text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 } } : Object.assign({}, t, { options: Object.assign({ bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 }, t.options || {}) }))),
  Object.assign({ isTextBox: true, fontSize: 11, color: C.text1, margin: 0, valign: "top" }, o));
const card = (s, x, y, w, h, fill) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill || LT2 }, line: { color: fill || LT2 }, rectRadius: 0.06 });
const stat = (s, x, y, w, big, label, color) => { card(s, x, y, w, 1.0); T(s, big, { x: x + 0.12, y: y + 0.06, w: w - 0.24, h: 0.45, fontSize: 21, bold: true, color: color || C.text2, fontFace: "Cambria" }); T(s, label, { x: x + 0.12, y: y + 0.52, w: w - 0.24, h: 0.45, fontSize: 8.5, color: C.accent3 }); };
const head = (s, txt, x, y, w) => T(s, txt, { x, y, w, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
const hdrRow = (cells) => cells.map((c, i) => ({ text: c, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: i ? "right" : "left", fontSize: 9 } }));
const row = (cells, opt) => cells.map((c, i) => ({ text: String(c), options: Object.assign({ align: i ? "right" : "left", fontSize: 9, color: THEME.colors.dk1 }, opt || {}) }));
const tableOpts = (x, y, w, colW, o) => Object.assign({ x, y, w, colW, border: { type: "solid", pt: 0.5, color: "D9DCE1" }, fontFace: "Calibri", fontSize: 9, rowH: 0.24, margin: [2, 4, 2, 4], valign: "middle" }, o || {});
const draft = (s, who) => { s.addShape(pres.ShapeType.roundRect, { x: 7.3, y: 0.08, w: 2.2, h: 0.22, fill: { color: "FFF4E8" }, line: { color: COPPER, width: 0.75 }, rectRadius: 0.05 }); T(s, "DRAFT FRAME – " + who + " to complete", { x: 7.3, y: 0.08, w: 2.2, h: 0.22, fontSize: 8, bold: true, color: C.accent1, align: "center", valign: "middle" }); };
const src = (s, txt) => T(s, txt, { x: 0.5, y: 5.05, w: 9.0, h: 0.2, fontSize: 7.5, color: C.accent3, italic: true });

// ===================================================================== 1 TITLE
pres.addSection({ title: "Opening" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" });
T(s, "Rexel", { x: 0.6, y: 1.2, w: 8.8, h: 1.0, fontSize: 54, bold: true, color: C.background1, fontFace: "Cambria" });
T(s, "Should an investor acquire a stake?", { x: 0.6, y: 2.2, w: 8.8, h: 0.6, fontSize: 24, color: C.accent1, fontFace: "Cambria" });
T(s, "Valuation of a listed company – listed peers, DCF and LBO bid", { x: 0.6, y: 2.85, w: 8.8, h: 0.4, fontSize: 14, color: "CADCFC" });
s.addShape(pres.ShapeType.rect, { x: 0.6, y: 3.65, w: 8.8, h: 0.9, fill: { color: "13315C" }, line: { color: "13315C" } });
T(s, [{ text: "Marwan Ferreira da Silva  ·  Douae", options: { bold: true, breakLine: true } }, { text: "Introduction to Corporate Finance – Université Paris Dauphine-PSL, M2 104  ·  Prof. David Heller", options: { breakLine: true } }, { text: "Valuation date: 30 October 2026  ·  Share price €34.47  ·  Submitted 2 November 2026", options: {} }], { x: 0.8, y: 3.72, w: 8.4, h: 0.8, fontSize: 11, color: C.background1, valign: "middle" });
s.addNotes("Title slide. Replace 'Douae' with full name before sending.");

// ===================================================================== 2 EXEC SUMMARY
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
s.addText("Executive summary – Buy: three methods converge 10–45% above a share price depressed by deal overhang", { placeholder: "title" });
card(s, 0.5, 1.15, 4.2, 1.1, NAVY);
T(s, "BUY – acquire a stake", { x: 0.65, y: 1.2, w: 3.9, h: 0.45, fontSize: 20, bold: true, color: C.background1, fontFace: "Cambria" });
T(s, "Target range €43–46 per share, +25% to +35% vs €34.47. Anchors: DCF base case and peer xEBITDA median.", { x: 0.65, y: 1.65, w: 3.9, h: 0.55, fontSize: 10.5, color: "CADCFC" });
stat(s, 0.5, 2.4, 2.0, "€34.47", "Share price 30 Oct 2026, below the €34.75 issue price");
stat(s, 2.7, 2.4, 2.0, "€43.1", "DCF base case – WACC 8.9%, g 2.0% (+25%)", COPPER);
stat(s, 0.5, 3.52, 2.0, "€46–50", "Listed peers, xEBITDA median to average (+34–46%)", COPPER);
stat(s, 2.7, 3.52, 2.0, "€37.6", "Maximum LBO bid at 20% IRR (+9%): the floor");
s.addImage({ path: CH + "02_football_field.png", x: 4.95, y: 1.15, w: 4.55, h: 3.3 });
bullets(s, ["Operations are at their best since 2022 (Q3 same-day sales +11.4%, guidance raised twice in 2026), yet the share has derated since the $1.4bn GCG acquisition and the €500m capital increase.",
  "Rexel trades at 9.8x 2026e EBITDA after leases against 13–14x for US distributors: a European discount that the fundamentals do not justify."], { x: 0.5, y: 4.62, w: 9.0, h: 0.45, fontSize: 9.5 });
s.addNotes("Pyramid: the answer first. Target range = DCF base (43) to peers xEBITDA median (46). LBO bid is a floor, not a target.");

// ===================================================================== 3 MARKET
pres.addSection({ title: "Market and company" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Market overview – a consolidating distribution market with structural electrification tailwinds", { placeholder: "title" });
draft(s, "Douae");
head(s, "Largest electrical distributors – latest annual sales (bn, reported currency)", 0.5, 1.1, 4.6);
s.addChart(pres.ChartType.bar, [{ name: "Sales", labels: ["Sonepar (€)", "Wesco ($, LTM)", "Rexel (€)", "Graybar ($)", "CED ($, est.)"], values: [33.6, 25.0, 19.4, 11.0, 8.0] }],
  { x: 0.5, y: 1.4, w: 4.6, h: 2.6, barDir: "bar", chartColors: [STEEL], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: THEME.colors.dk1, dataLabelFontFace: "+mn-lt", catAxisLabelFontSize: 9, catAxisLabelColor: THEME.colors.dk1, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin", barGapWidthPct: 60 });
T(s, "US market concentration: the top-10 distributors hold c.56% of industry sales and the five national chains (Wesco, Sonepar, Graybar, Rexel, CED) at least 46% – a long tail of regional players feeds M&A (125+ distributors acquired 2021–25).", { x: 0.5, y: 4.05, w: 4.6, h: 0.9, fontSize: 9.5, color: C.text1 });
head(s, "Demand drivers", 5.4, 1.1, 4.1);
const drivers = [["Data centres & AI", "Triple-digit growth at Rexel US in Q3 2026; 13% of US sales, 25% of US backlog. Wesco data-centre sales +45% y/y."], ["Energy transition", "HVAC, solar and EV charging > 20% of Rexel's European sales, +13% in Q3 2026; electrification plans and energy-price volatility."], ["Grid & industrial automation", "Utility capex, reshoring and industrial projects; industrial automation +20% at Rexel US in Q3."]];
drivers.forEach((d, i) => { const y = 1.42 + i * 0.82; card(s, 5.4, y, 4.1, 0.72); s.addShape(pres.ShapeType.ellipse, { x: 5.52, y: y + 0.18, w: 0.36, h: 0.36, fill: { color: COPPER }, line: { color: COPPER } }); T(s, String(i + 1), { x: 5.52, y: y + 0.18, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" }); T(s, d[0], { x: 6.0, y: y + 0.07, w: 3.4, h: 0.25, fontSize: 10.5, bold: true, color: C.text2 }); T(s, d[1], { x: 6.0, y: y + 0.32, w: 3.4, h: 0.4, fontSize: 8.5, color: C.text1 }); });
head(s, "Risks", 5.4, 3.95, 4.1);
T(s, "Cyclical exposure to non-residential construction · copper (cables c.16% of sales, copper c.60% of cable) · pricing pressure and digital disintermediation · interest rates.", { x: 5.4, y: 4.22, w: 4.1, h: 0.7, fontSize: 9.5 });
src(s, "Sources: Sonepar 2025 results; Wesco LTM (Yahoo Finance); Rexel FY 2025 and Q3 2026 releases; Electrical Wholesaling 2026 Electrical Pyramid; Axia Growth distribution M&A 2026. [Douae: add market size + growth with a cited report or Rexel's URD chapter 1 / CMD]");
s.addNotes("DOUAE – to complete: (1) market size and expected growth with a citable source (Rexel URD 2025 chapter 1 'addressable market', Capital Markets Day deck, or a market report); (2) one sentence on the three end-markets (residential, non-residential, industrial) and their 2026 trend; (3) check Graybar/CED figures (press or MDM top distributors list). Keep the chart; replace '(est.)' once sourced. Remove the DRAFT tag when done.");

// ===================================================================== 4 COMPETITIVE LANDSCAPE
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Competitive landscape – Rexel is the #2 global distributor, with Sonepar the only rival of comparable reach", { placeholder: "title" });
draft(s, "Douae");
// map
s.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.15, w: 5.2, h: 3.55, fill: { color: LT2 }, line: { color: LT2 } });
s.addShape(pres.ShapeType.line, { x: 0.95, y: 4.4, w: 4.55, h: 0, line: { color: GREY, width: 1 } });
s.addShape(pres.ShapeType.line, { x: 0.95, y: 1.35, w: 0, h: 3.05, line: { color: GREY, width: 1 } });
T(s, "Geographic reach  →  (regional · national · multi-continent)", { x: 0.95, y: 4.44, w: 4.55, h: 0.22, fontSize: 8, color: C.accent3 });
T(s, "Scale (sales) ↑", { x: 0.55, y: 1.2, w: 1.2, h: 0.2, fontSize: 8, color: C.accent3 });
const bub = [["Sonepar €33.6bn", 4.7, 1.65, 0.78, NAVY], ["Wesco $25bn", 3.3, 1.95, 0.68, STEEL], ["Rexel €19.4bn", 4.6, 2.75, 0.62, COPPER], ["Graybar $11bn", 2.4, 2.75, 0.5, STEEL], ["CED c.$8bn", 1.75, 3.45, 0.44, STEEL], ["Ahlsell · Würth · Solar", 3.2, 3.6, 0.36, GREY]];
bub.forEach(b => { s.addShape(pres.ShapeType.ellipse, { x: b[1] - b[3] / 2, y: b[2] - b[3] / 2, w: b[3], h: b[3], fill: { color: b[4], transparency: 15 }, line: { color: "FFFFFF", width: 1 } }); T(s, b[0], { x: b[1] - 0.9, y: b[2] + b[3] / 2 + 0.02, w: 1.8, h: 0.2, fontSize: 8.5, bold: true, color: C.text2, align: "center" }); });
T(s, "Private: Sonepar (40 countries, family-owned), Graybar (employee-owned), CED. Listed: Rexel, Wesco. Regional leaders: Ahlsell (Nordics), Würth (DACH), Solar (Benelux/Nordics).", { x: 0.95, y: 4.0, w: 4.6, h: 0.38, fontSize: 7.5, color: C.accent3 });
head(s, "How Rexel competes", 6.0, 1.1, 3.5);
bullets(s, [{ text: "Value-added services: ", options: { bold: true } }, { text: "moving upstream in the value chain (engineering, kitting, project support); GCG adds proprietary offerings with >75% value-added content.", options: { bullet: false } },
  { text: "Digital: ", options: { bold: true } }, { text: "36% of Q3 2026 sales online (44% in Europe), AI tools lifting North America by +600 bp y/y.", options: { bullet: false } },
  { text: "Key accounts and high-growth verticals: ", options: { bold: true } }, { text: "data centres, utilities, industrial automation; backlog in North America > $3bn (+42% vs end-2025).", options: { bullet: false } },
  { text: "Bolt-on M&A: ", options: { bold: true } }, { text: "four North American deals in 2026 (TC 360, Revere, DEE Electronics, GCG), 2–3 pts of growth per year targeted.", options: { bullet: false } }], { x: 6.0, y: 1.4, w: 3.5, h: 3.3, fontSize: 9.5 });
src(s, "Sources: company reports (Sonepar, Wesco, Rexel), Rexel Q3 2026 presentation, GCG acquisition release. Bubble area is indicative of sales. [Douae: verify Graybar/CED/Ahlsell figures, add Sonepar/Wesco strategy in one line each]");
s.addNotes("DOUAE – to complete: verify the private players' sales (Graybar ~$11bn, CED ~$8bn, Ahlsell, Würth Elektrogroßhandel, Solar A/S), add one line on Sonepar's and Wesco's strategies (Sonepar: family-owned, 'Spark' digital program, US expansion; Wesco: data centre/utility supply chain), and if possible a market-share figure for Rexel in France (#1 or #2) and the US. Remove the DRAFT tag when done.");

// ===================================================================== 5 COMPANY
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Rexel at a glance – €19.4bn of sales, 1,876 branches in 17 countries, repositioning toward high-growth segments", { placeholder: "title" });
draft(s, "Douae");
head(s, "Sales by geography, FY 2025", 0.5, 1.1, 3.0);
s.addChart(pres.ChartType.doughnut, [{ name: "Sales", labels: ["North America", "Europe", "Asia-Pacific"], values: [46, 48, 6] }], { x: 0.4, y: 1.35, w: 3.1, h: 2.3, holeSize: 55, chartColors: [COPPER, NAVY, GREY], showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 9, dataLabelFontFace: "+mn-lt", showLegend: true, legendPos: "b", legendFontSize: 8, legendFontFace: "+mn-lt", legendColor: THEME.colors.dk1, showTitle: false });
s.addTable([hdrRow(["Segment", "Sales €m", "Adj. EBITA %"]), row(["North America", "8,907", "7.3%"]), row(["Europe", "9,403", "5.7%"]), row(["Asia-Pacific", "1,105", "1.1%"]), row(["Group", "19,415", "6.0%"], { bold: true })], tableOpts(0.5, 3.75, 3.0, [1.3, 0.9, 0.8]));
head(s, "Business", 3.8, 1.1, 2.9);
bullets(s, ["Multichannel B2B distributor of electrical products and services for residential, non-residential and industrial customers.", "1,876 branches, 26,306 employees, 17 countries; US 38% of group sales, France 19%.", "Warehouse sales 81%, direct sales 19%; cables c.16% of sales.", "Listed on Euronext Paris (SBF 120, STOXX 600); free float 74%, Cevian Capital 23%, employees 1%.", "Ratings Ba1 / BB+ (stable); net debt / EBITDAaL 2.4x at June 2026, c.2.0x targeted from 2027."], { x: 3.8, y: 1.4, w: 2.9, h: 3.5, fontSize: 9.5 });
head(s, "Strategy – Axelerate 2028", 7.0, 1.1, 2.5);
card(s, 7.0, 1.4, 2.5, 2.05);
bullets(s, ["Sales growth 5–8% a year, of which 2–3 pts from M&A", "Adjusted EBITA margin above 7%", "FCF conversion c.65% of EBITDAaL", "Dividend payout at least 40%, buybacks €50–150m a year"], { x: 7.12, y: 1.5, w: 2.3, h: 1.9, fontSize: 9 });
T(s, "2026 guidance (raised twice): same-day growth c.7%, adjusted EBITA margin ≥ 6.4%, FCF conversion > 65%.", { x: 7.0, y: 3.55, w: 2.5, h: 0.8, fontSize: 9, color: C.text2, bold: true });
T(s, "GCG (Chicago): $1.4bn EV, > $1.1bn sales, c.11% EBITA margin, < 8x EBITDAaL post synergies, closing expected end-2026.", { x: 7.0, y: 4.3, w: 2.5, h: 0.7, fontSize: 8.5, color: C.text1 });
src(s, "Sources: Rexel FY 2025 release and financial statements, Q3 2026 release, GCG release, Rexel shareholding page (31 Dec 2025).");
s.addNotes("DOUAE – to complete: add two lines on history (spun out of PPR in 2005, IPO 2007) and on the CEO's repositioning (Guillaume Texier, since 2021); confirm the Axelerate 2028 targets from the Capital Markets Day deck. Remove the DRAFT tag when done.");

// ===================================================================== 6 FINANCIAL PERFORMANCE
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Financial performance – growth is accelerating and margins are climbing toward the 7% ambition", { placeholder: "title" });
s.addImage({ path: CH + "03_financial_performance.png", x: 0.5, y: 1.1, w: 9.0, h: 2.75 });
stat(s, 0.5, 3.95, 2.1, "+11.4%", "Q3 2026 same-day sales growth, best quarter since 2022 (US +18.2%)", COPPER);
stat(s, 2.8, 3.95, 2.1, "6.0% → 6.4%", "Adjusted EBITA margin FY 2025 → 2026 guidance floor (consensus 7.0% in 2028)");
stat(s, 5.1, 3.95, 2.1, "66%", "FCF conversion 2025 (76% ex €124m fine); > 65% guided");
stat(s, 7.4, 3.95, 2.1, "2.4x", "Net debt / EBITDAaL at June 2026; c.2.0x from 2027 after the €500m raise");
src(s, "Sources: Rexel FY 2025, H1 2026 and Q3 2026 releases; consensus collected by Rexel (15 Oct 2026). 2022–23 sales from Yahoo Finance.");

// ===================================================================== 7 STOCK
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Stock market performance – strong results, weak share: Rexel has derated since the GCG deal and the capital increase", { placeholder: "title" });
s.addImage({ path: CH + "01_stock_performance.png", x: 0.5, y: 1.1, w: 6.2, h: 3.4 });
head(s, "Reading", 6.9, 1.1, 2.6);
bullets(s, ["+37% over two years vs +93% for Wesco and +7% for the CAC 40: Rexel has captured the electrification theme, but far less than its US peer.", "Since the GCG announcement (25 Sep) the share has fallen below the €34.75 issue price of the capital increase, despite the Q3 beat (sales 4.8% above consensus).", "Sell-side: Buy consensus, 13 analysts, mean target €41.9 (range €30–46).", "Market-implied multiple: 9.8x 2026e EBITDA after leases, vs 13–14x for Wesco, Ferguson and Watsco."], { x: 6.9, y: 1.4, w: 2.6, h: 3.5, fontSize: 9.5 });
T(s, "Beta: 1.25 (2-year weekly vs CAC 40), 0.98 (Yahoo 5-year monthly).", { x: 0.5, y: 4.6, w: 6.2, h: 0.3, fontSize: 9, color: C.accent3 });
src(s, "Sources: Yahoo Finance daily closes (price return, local currency); Rexel press releases for events; analyst targets per Yahoo Finance, 30 Oct 2026.");

// ===================================================================== 8 COMPS TABLE
pres.addSection({ title: "Valuation" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("Listed peers – five B2B distributors; Rexel trades 2–4 turns below the sector median on every multiple", { placeholder: "title" });
s.addTable([hdrRow(["Company", "Mkt cap (bn)", "EV/EBITDA LTM", "EV/EBITDA FY+1", "EV/EBIT LTM", "EV/EBIT FY+1", "P/E LTM", "P/E FY+1"]),
  row(["Wesco (US)", "$16.7", "13.8x", "12.4x", "16.2x", "14.6x", "23.5x", "16.4x"]), row(["Grainger (US)", "$64.2", "21.9x", "19.5x", "23.9x", "21.4x", "34.3x", "26.6x"]), row(["Ferguson (US/UK)", "$46.2", "13.9x", "12.9x", "15.6x", "14.4x", "19.6x", "18.8x"]), row(["Watsco (US)", "$10.9", "15.5x", "13.4x", "16.5x", "14.2x", "23.0x", "23.6x"]), row(["Bunzl (UK)", "£8.9", "8.8x", "8.3x", "13.2x", "12.5x", "19.4x", "14.5x"]),
  row(["Sector median", "", "13.9x", "12.9x", "16.2x", "14.4x", "23.0x", "18.8x"], { bold: true, fill: { color: LT2 } }), row(["Sector average", "", "14.8x", "13.3x", "17.1x", "15.4x", "24.0x", "20.0x"], { bold: true, fill: { color: LT2 } }),
  row(["Rexel at market", "€10.6", "11.0x", "9.8x", "11.6x", "10.3x", "15.7x", "14.9x"], { bold: true, color: COPPER })], tableOpts(0.5, 1.15, 6.1, [1.45, 0.75, 0.7, 0.7, 0.65, 0.65, 0.6, 0.6]));
T(s, "EV = market cap + net debt (Yahoo Finance, 30 Oct 2026). US GAAP peers expense operating leases above EBITDA, so Rexel is measured on EBITDA after leases (EBITDAaL) and adjusted EBITA, with lease liabilities excluded from net debt. FY+1 EBITDA/EBIT for peers = next-year revenue consensus × LTM margin; FY+1 net income = consensus EPS. Rexel FY+1 = consensus collected by Rexel (13 brokers, Oct 2026).", { x: 0.5, y: 3.42, w: 6.1, h: 0.85, fontSize: 7.5, color: C.accent3 });
head(s, "Sample", 0.5, 4.3, 6.1);
T(s, "Listed B2B distributors with comparable business models: electrical (Wesco), MRO (Grainger), plumbing/HVAC (Ferguson, Watsco), consumables (Bunzl). Sonepar and Graybar, the closest comparables, are private. Fastenal (29x EBITDA), Legrand and Schneider (manufacturers) are excluded.", { x: 0.5, y: 4.55, w: 6.1, h: 0.45, fontSize: 8 });
card(s, 6.85, 1.15, 2.65, 3.9, NAVY);
T(s, "Implied value of Rexel (€ per share)", { x: 7.0, y: 1.25, w: 2.4, h: 0.3, fontSize: 10.5, bold: true, color: C.background1 });
const imp = [["xEBITDA – anchor", "46 – 50", true], ["xEBIT", "52 – 56", false], ["P/E", "43 – 53", false], ["xSales", "77 – 95 (n.m.)", false]];
imp.forEach((r, i) => { const y = 1.7 + i * 0.6; T(s, r[0], { x: 7.0, y, w: 1.3, h: 0.5, fontSize: 10, color: "CADCFC", valign: "middle" }); T(s, r[1], { x: 8.1, y, w: 1.35, h: 0.5, fontSize: r[2] ? 16 : 13, bold: r[2], color: r[2] ? "F2A86F" : C.background1, align: "right", valign: "middle" }); });
T(s, "Median to average of the five peers, applied to Rexel 2025A / 2026e aggregates, less €2.8bn net debt and minorities, over 308.5m shares.", { x: 7.0, y: 4.15, w: 2.4, h: 0.85, fontSize: 8.5, color: "CADCFC" });
src(s, "Sources: Yahoo Finance (30 Oct 2026), company reports, Rexel consensus (15 Oct 2026). Model: Comps sheet.");

// ===================================================================== 9 COMPS READING
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("Reading the multiples – xEBITDA is the anchor; the sales multiple is meaningless and the discount is a listing discount", { placeholder: "title" });
s.addImage({ path: CH + "04_peer_benchmark.png", x: 0.5, y: 1.1, w: 5.6, h: 2.8 });
head(s, "Why the multiples disagree", 6.3, 1.1, 3.2);
bullets(s, [{ text: "xSales overvalues Rexel (€77–95). ", options: { bold: true } }, { text: "Peers earn 10–23% EBITDA margins on each euro of sales; Rexel earns 6%. A sales multiple built on richer margins has no meaning here.", options: { bullet: false } },
  { text: "xEBIT above xEBITDA. ", options: { bold: true } }, { text: "Rexel's D&A after leases is only 6% of EBITDAaL (asset-light, leases already deducted), lower than peers', so an EBIT multiple flatters it slightly.", options: { bullet: false } },
  { text: "P/E is wide (€43–53). ", options: { bold: true } }, { text: "Net income is after interest and tax: Rexel's 2.4x leverage and 30% 2025 tax rate are not comparable with net-cash US peers.", options: { bullet: false } },
  { text: "Keep xEBITDA: €46–50. ", options: { bold: true } }, { text: "Same margin as Wesco, three turns cheaper: the gap is a European listing and growth-history discount, widened by deal overhang, not an operating gap.", options: { bullet: false } }], { x: 6.3, y: 1.4, w: 3.2, h: 3.6, fontSize: 9 });
T(s, "Caveat: a five-name sample dominated by US MRO distributors that trade at a structural premium; Grainger alone lifts the average. The median is therefore the reference.", { x: 0.5, y: 4.0, w: 5.6, h: 0.6, fontSize: 9, color: C.text2 });
src(s, "Sources: Yahoo Finance LTM figures, 30 Oct 2026; Rexel FY 2025 release. Model: Comps sheet.");

// ===================================================================== 10 DCF ASSUMPTIONS
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("DCF assumptions – Rexel's own broker consensus to 2028, soft landing to 2% growth, WACC of 8.9%", { placeholder: "title" });
head(s, "Discount rate (CAPM / WACC at market weights)", 0.5, 1.1, 4.3);
s.addTable([hdrRow(["Input", "Value", "Source / rationale"]),
  row(["Risk-free rate", "4.13%", "50% Bund 3.33% + 50% UST 4.94% (Rexel is 51% North America)"]), row(["Market risk premium", "5.0%", "Damodaran, France, Jan 2026"]), row(["Beta", "1.20", "Own 2-yr weekly regression 1.25; Yahoo 0.98"]), row(["Cost of equity", "10.1%", "rf + β × MRP"]), row(["Pre-tax cost of debt", "5.6%", "rf + 150 bp spread (Ba1 / BB+)"]), row(["Tax rate", "27%", "Normative rate (30% in 2025 incl. French surtax)"]), row(["Debt / (D + E)", "21%", "€2.8bn net debt pro forma, €10.6bn market cap"]), row(["WACC", "8.9%", "Rexel impairment tests use 8.4–9.8% by country"], { bold: true, fill: { color: LT2 } }), row(["Perpetual growth g", "2.0%", "Inflation-based; Rexel uses 1.9–2.2%"], { bold: true, fill: { color: LT2 } })],
  tableOpts(0.5, 1.4, 4.3, [1.3, 0.6, 2.4], { rowH: 0.27 }));
head(s, "Business plan (€m)", 5.1, 1.1, 4.4);
s.addTable([hdrRow(["", "2025A", "2026e", "2027e", "2028e", "2032e"]),
  row(["Sales", "19,415", "20,578", "21,938", "22,949", "25,805"]), row(["growth", "+0.7%", "+6.0%", "+6.6%", "+4.6%", "+2.0%"], { color: GREY, italic: true }), row(["EBITDA after leases", "1,227", "1,379", "1,539", "1,669", "1,877"]), row(["Adjusted EBITA (EBIT)", "1,158", "1,305", "1,466", "1,599", "1,798"]), row(["margin", "6.0%", "6.3%", "6.7%", "7.0%", "7.0%"], { color: GREY, italic: true }), row(["Capex", "(136)", "(144)", "(154)", "(161)", "(79)"]), row(["Δ WCR", "", "(165)", "(193)", "(143)", "(72)"]), row(["Free cash flow", "", "718", "797", "933", "1,241"], { bold: true, fill: { color: LT2 } })],
  tableOpts(5.1, 1.4, 4.4, [1.5, 0.58, 0.58, 0.58, 0.58, 0.58], { rowH: 0.27 }));
bullets(s, ["2026–2028: mean of the consensus collected by Rexel from 13 brokers (5–13 Oct 2026), excluding GCG – collected before the 22 Oct guidance raise, so conservative.", "2029–2032: growth converges linearly to g; EBITDAaL margin frozen at the 2028 level (7.3%); capex converges to D&A; WCR held at 14.2% of sales (2025 financial balance sheet).", "FCF = EBIT × (1 – t) + D&A – capex – ΔWCR; flows after leases, so lease liabilities are not in net debt."], { x: 5.1, y: 3.9, w: 4.4, h: 1.15, fontSize: 8.5 });
src(s, "Sources: Rexel consensus (15 Oct 2026), 2025 financial statements (note 12 impairment tests), Trading Economics (30 Oct 2026), Damodaran. Model: DCF sheet.");

// ===================================================================== 11 DCF RESULT
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("DCF result – €43 per share, 25% above the market; above the price across the whole reasonable grid", { placeholder: "title" });
head(s, "From free cash flows to value per share (€m)", 0.5, 1.1, 4.0);
s.addTable([hdrRow(["", "€m"]), row(["Sum of discounted FCF 2026–2032", "5,241"]), row(["Terminal value: FCF2032 × (1+g) / (K – g)", "18,414"]), row(["Discounted terminal value (68% of EV)", "10,898"]), row(["Enterprise value", "16,139"], { bold: true }), row(["Net debt pro forma + minorities", "(2,833)"]), row(["Equity value", "13,306"], { bold: true, fill: { color: LT2 } }), row(["Shares outstanding (m)", "308.5"]), row(["Value per share", "€43.1"], { bold: true, fill: { color: LT2 }, color: COPPER }), row(["Upside vs €34.47", "+25%"], { bold: true })], tableOpts(0.5, 1.4, 4.0, [2.9, 1.1], { rowH: 0.28 }));
T(s, "Implied EV / EBITDAaL 2026e: 11.7x (market: 9.8x; peer median 12.9x). Terminal value is 68% of EV – normal for a distributor with 2% perpetual growth, but it makes K and g the two numbers to challenge.", { x: 0.5, y: 4.3, w: 4.0, h: 0.75, fontSize: 9, color: C.text2 });
s.addImage({ path: CH + "05_dcf_sensitivity.png", x: 4.8, y: 1.1, w: 4.7, h: 3.0 });
T(s, "Every cell with K ≤ 9.5% and g ≥ 1.5% is above €34.47. The market price corresponds to a WACC of about 10% at 2% growth – above anything Rexel itself uses.", { x: 4.8, y: 4.2, w: 4.7, h: 0.8, fontSize: 9.5, color: C.text1 });
src(s, "Model: DCF sheet, valuation date 30 Oct 2026, full-year 2026 FCF discounted 0.17 year. Net debt at 30 June 2026 less the €500m capital increase.");

// ===================================================================== 12 LBO STRUCTURE
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("LBO-style bid – a sponsor buys 100% through a holding financed 48% with debt, exits after three years", { placeholder: "title" });
const box = (x, y, w, h, fill, title, body, tcol) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.06 }); T(s, title, { x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.28, fontSize: 10.5, bold: true, color: tcol || "FFFFFF" }); T(s, body, { x: x + 0.15, y: y + 0.4, w: w - 0.3, h: h - 0.45, fontSize: 8.5, color: tcol || "FFFFFF" }); };
box(0.5, 1.15, 4.2, 0.95, NAVY, "Investment fund", "Equity €6.1bn (52% of sources) – required IRR 20%, 3-year hold, exit at the entry multiple of 10.3x EBIT.");
s.addShape(pres.ShapeType.downArrow, { x: 2.45, y: 2.15, w: 0.3, h: 0.3, fill: { color: GREY }, line: { color: GREY } });
box(0.5, 2.5, 4.2, 1.35, STEEL, "NewCo holding – acquisition debt €5.5bn (4.0x 2026e EBITDAaL)", "Senior A €2.8bn (50%), Euribor + 2.75% + swap = 5.45%, amortising over 7 years\nSenior B €1.9bn (35%), 5.95%, bullet 8 years\nJunior C €0.8bn (15%), 8.70%, bullet 9 years\nTax consolidation: acquisition interest offsets Rexel's taxable income; 100% of net income upstreamed.");
s.addShape(pres.ShapeType.downArrow, { x: 2.45, y: 3.9, w: 0.3, h: 0.3, fill: { color: GREY }, line: { color: GREY } });
box(0.5, 4.25, 4.2, 0.75, LT2, "Rexel – 100% acquired", "Existing net financial debt €2.8bn stays in the target: total leverage 6.0x EBITDAaL at entry, consistent with large-cap LBO practice.", THEME.colors.dk1);
head(s, "Sources & uses at entry (end-2026, €m)", 5.0, 1.1, 4.5);
s.addTable([hdrRow(["Sources", "€m", "Uses", "€m"]), row(["Fund equity", "6,075", "Purchase of 100% of equity (the bid)", "11,593"]), row(["Senior A", "2,759", "", ""]), row(["Senior B", "1,931", "", ""]), row(["Junior C", "828", "", ""]), row(["Total", "11,593", "Total", "11,593"], { bold: true, fill: { color: LT2 } })], tableOpts(5.0, 1.4, 4.5, [1.1, 0.8, 1.8, 0.8], { rowH: 0.27 }));
head(s, "Mechanics (course method)", 5.0, 3.15, 4.5);
bullets(s, ["Rexel's plan: consensus 2026–28, then soft landing over 7 years; interest on its own debt at 5.6%.", "Holding P&L: dividends received – 5% taxable share × 27% – acquisition interest + tax of the fiscal group; cash accumulates after Senior A amortisation.", "Exit 2029: EV = 10.3x × EBITA 2029 (€1,667m) = €17.2bn; consolidated net debt €6.7bn (Rexel + holding – cash).", "Fund entry = exit equity / 1.20³; bid = fund entry + acquisition debt."], { x: 5.0, y: 3.42, w: 4.5, h: 1.6, fontSize: 8.5 });
src(s, "Model: LBO sheet. Euribor 6M 2.2% (to verify), spreads in line with the course case (senior 2.75–3.25%, junior 6%).");

// ===================================================================== 13 LBO RESULT
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("LBO result – a fund can bid at most €37.6 per share for a 20% IRR: a 9% premium, too thin for a public-to-private", { placeholder: "title" });
head(s, "Returns and maximum bid", 0.5, 1.1, 4.2);
s.addTable([hdrRow(["", "€m"]), row(["Enterprise value at exit (2029): 10.3x × 1,667", "17,186"]), row(["Consolidated net debt at exit", "(6,688)"]), row(["Equity value at exit", "10,498"], { bold: true }), row(["Fund equity at entry = 10,498 / 1.20³", "6,075"]), row(["Acquisition debt", "5,518"]), row(["Maximum price for 100% of the equity", "11,593"], { bold: true, fill: { color: LT2 } }), row(["Bid per share (308.5m shares)", "€37.6"], { bold: true, fill: { color: LT2 }, color: COPPER }), row(["Premium vs €34.47 / money multiple", "+9% / 1.73x"], { bold: true })], tableOpts(0.5, 1.4, 4.2, [3.1, 1.1], { rowH: 0.28 }));
head(s, "Sensitivity of the bid (€ per share)", 5.0, 1.1, 4.5);
s.addTable([hdrRow(["IRR \\ acquisition debt", "3.5x", "4.0x", "4.5x", "5.0x"]), row(["15%", "39.7", "40.3", "40.8", "41.4"]), row(["20%", "36.8", "37.6", "38.4", "39.1"], { bold: true, fill: { color: LT2 } }), row(["25%", "34.4", "35.3", "36.3", "37.2"])], tableOpts(5.0, 1.4, 4.5, [1.7, 0.7, 0.7, 0.7, 0.7], { rowH: 0.27 }));
s.addTable([hdrRow(["Exit multiple (xEBIT), 4.0x / 20%", "9.0x", "10.3x", "11.0x", "12.0x"]), row(["Bid per share", "33.5", "37.6", "39.7", "42.9"])], tableOpts(5.0, 2.65, 4.5, [1.7, 0.7, 0.7, 0.7, 0.7], { rowH: 0.27 }));
bullets(s, ["Only a re-rating at exit (12x EBIT) or a 15% return target gets a sponsor above €40 – the price action, not leverage, sets the ceiling.", "With 6.0x total leverage and Ba1/BB+ ratings already, financing more aggressively is not credible; a five-year hold at 20% lowers the bid to €35.", "Conclusion: an LBO at a normal 25–30% premium (€43–45) is not feasible; the stock is cheap for a long-term holder or a strategic buyer (Sonepar, Wesco) who can pay for synergies."], { x: 5.0, y: 3.4, w: 4.5, h: 1.6, fontSize: 9 });
src(s, "Model: LBO sheet and sensitivity replica. Interest on opening net debt; dividend tax 5% × 27%.");

// ===================================================================== 14 INTERPRETATION
pres.addSection({ title: "Conclusion" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Conclusion" });
s.addText("Interpretation – the methods are consistent with each other and with the sell-side; the market is the outlier", { placeholder: "title" });
s.addImage({ path: CH + "02_football_field.png", x: 0.5, y: 1.1, w: 4.6, h: 2.2 });
head(s, "Consistency checks", 5.3, 1.1, 4.2);
bullets(s, ["DCF (€43) sits inside the analyst range (€30–46, mean €41.9) and just below the peer xEBITDA band (€46–50): intrinsic and relative values agree.", "Market price implies a 10% WACC at 2% growth, or 9.8x EBITDAaL 2026e – both outside Rexel's own and the peers' ranges.", "The LBO bid (€37.6) is a floor: a financial buyer earns 20% at that price, so downside from €34.47 is cushioned.", "Peer FY+1 EBITDA/EBIT are proxies (consensus revenue × LTM margin); the median is used to limit sample bias."], { x: 5.3, y: 1.4, w: 4.2, h: 1.95, fontSize: 9.5 });
head(s, "What would change the answer", 0.5, 3.38, 9.0);
s.addTable([hdrRow(["Risk", "Impact on value", "Mitigant / watch item"]), row(["US cycle, data-centre concentration (13% of US sales)", "2028 margin 6.5% instead of 7.0%: DCF c.–€3 per share", "Backlog $3bn (3.4 months); diversified end-markets"]), row(["GCG integration ($1.4bn, closing end-2026)", "Synergies missed: < 8x multiple holds only post synergies", "EPS accretive yr 1; leverage 2.0x from 2027"]), row(["Copper and pricing (cables 16% of sales)", "Non-recurring margin swings; 2026 helped by price", "Pass-through model, monitored monthly"]), row(["French sovereign spread (OAT 4.97%) and rates", "+50 bp WACC ≈ –€4 per share", "51% of sales in North America; blended rf"])], tableOpts(0.5, 3.64, 9.0, [3.0, 3.2, 2.8], { rowH: 0.2, fontSize: 8 }));
src(s, "Model: Summary, DCF and LBO sheets; Rexel Q3 2026 presentation; Trading Economics.");

// ===================================================================== 15 RECOMMENDATION
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("Investment recommendation – BUY: acquire a stake, target €43–46 per share (+25% to +35%)", { placeholder: "title" });
const col = (x, title, items, color) => { s.addShape(pres.ShapeType.roundRect, { x, y: 1.35, w: 2.85, h: 3.2, fill: { color: "13315C" }, line: { color: "13315C" }, rectRadius: 0.06 }); T(s, title, { x: x + 0.15, y: 1.45, w: 2.55, h: 0.35, fontSize: 13, bold: true, color: color }); bullets(s, items, { x: x + 0.15, y: 1.85, w: 2.55, h: 2.6, fontSize: 9.5, color: "FFFFFF" }); };
col(0.5, "Thesis", ["Best operating momentum since 2022: +11.4% same-day growth, guidance raised twice, margin on the way to 7%.", "Valuation gap: 9.8x EBITDAaL vs 13–14x for US distributors, with the same margin as Wesco.", "Balance sheet funded: the €500m raise keeps leverage at c.2.0x from 2027 while GCG adds a c.11%-margin platform."], "F2A86F");
col(3.575, "Catalysts", ["GCG closing by year-end and first synergy guidance.", "FY 2026 results (February 2027): delivery of the ≥6.4% margin and 2027 outlook including GCG.", "Consensus upgrades: the October consensus predates the Q3 beat and the guidance raise.", "Share price below the €34.75 placement: institutional support at that level."], "CADCFC");
col(6.65, "Risks", ["US data-centre cycle and tougher comparables in Q4 2026.", "GCG integration and synergy delivery; execution of four acquisitions in one year.", "Copper, pricing and European macro; French sovereign spread lifting the cost of capital.", "Stop-loss logic: a 2028 margin below 6.5% removes most of the upside."], "F2A86F");
T(s, "Sizing: a long-term equity stake at or below €37 (the LBO floor) with a €43–46 target offers an asymmetric 3:1 upside/downside; we would not pursue a leveraged buyout at a control premium.", { x: 0.5, y: 4.65, w: 9.0, h: 0.4, fontSize: 10, color: "FFFFFF", bold: true });
T(s, "Sources: Rexel releases and financial statements (FY 2025, H1 2026, Q3 2026, GCG, capital increase); Rexel broker consensus (15 Oct 2026); Yahoo Finance; Trading Economics; Damodaran (Jan 2026); Electrical Wholesaling. Full model and data on github.com/marwanferreira/rexel-valuation.", { x: 0.5, y: 5.05, w: 8.6, h: 0.25, fontSize: 7, color: "8DA9C4" });

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
