const pptxgen = require("pptxgenjs");
const fs = require("fs");
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
const NAVY = THEME.colors.dk2, COPPER = THEME.colors.accent1, STEEL = THEME.colors.accent2, GREY = THEME.colors.accent3, LT2 = THEME.colors.lt2;
const FOOT = "Rexel – valuation of a listed company  |  Dauphine M2 104  |  Valuation date 30 October 2026";

pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 8.0, h: 0.75, fontSize: 19, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { image: { x: 0.5, y: 5.22, w: 1.25, h: 0.27, path: CH + "logos/dauphine.png" } },
    { text: { text: FOOT, options: { x: 1.95, y: 5.28, w: 6.8, h: 0.25, fontSize: 8, color: C.accent3, margin: 0 } } },
    { image: { x: 8.55, y: 0.1, w: 0.95, h: 0.36, path: CH + "logos/rexel.png", sizing: { type: "contain", w: 0.95, h: 0.36 } } },
  ],
  slideNumber: { x: 8.9, y: 5.28, w: 0.6, h: 0.25, fontSize: 8, color: C.accent3, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { color: NAVY },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { rect: { x: 0.5, y: 5.17, w: 1.45, h: 0.37, fill: { color: "FFFFFF" } } }, { image: { x: 0.6, y: 5.22, w: 1.25, h: 0.27, path: CH + "logos/dauphine.png" } }],
  slideNumber: { x: 8.9, y: 5.28, w: 0.6, h: 0.25, fontSize: 8, color: C.accent3, align: "right" } });

// helpers --------------------------------------------------------------
const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox: true, fontSize: 11, color: C.text1, margin: 0, valign: "top" }, o));
const bullets = (s, items, o) => s.addText(items.map((t, i) => (typeof t === "string" ? { text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 } } : Object.assign({}, t, { options: Object.assign({ bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 }, t.options || {}) }))),
  Object.assign({ isTextBox: true, fontSize: 11, color: C.text1, margin: 0, valign: "top" }, o));
const card = (s, x, y, w, h, fill) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill || LT2 }, line: { color: fill || LT2 }, rectRadius: 0.06 });
const stat = (s, x, y, w, big, label, color) => { card(s, x, y, w, 1.0); T(s, big, { x: x + 0.12, y: y + 0.06, w: w - 0.24, h: 0.45, fontSize: 21, bold: true, color: color || C.text2, fontFace: "Cambria" }); T(s, label, { x: x + 0.12, y: y + 0.52, w: w - 0.24, h: 0.45, fontSize: 8.5, color: C.accent3 }); };
const head = (s, txt, x, y, w) => T(s, txt, { x, y, w, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
const hdrRow = (cells, first) => cells.map((c, i) => ({ text: c, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: first === "c" ? "center" : (i ? "right" : "left"), fontSize: 9 } }));
const row = (cells, opt) => cells.map((c, i) => ({ text: String(c), options: Object.assign({ align: i ? "right" : "left", fontSize: 9, color: THEME.colors.dk1 }, opt || {}) }));
const tableOpts = (x, y, w, colW, o) => Object.assign({ x, y, w, colW, border: { type: "solid", pt: 0.5, color: "D9DCE1" }, fontFace: "Calibri", fontSize: 9, rowH: 0.24, margin: [2, 4, 2, 4], valign: "middle" }, o || {});
const draft = (s, who) => { s.addShape(pres.ShapeType.roundRect, { x: 6.2, y: 0.12, w: 2.2, h: 0.22, fill: { color: "FFF4E8" }, line: { color: COPPER, width: 0.75 }, rectRadius: 0.05 }); T(s, "DRAFT FRAME – " + who + " to complete", { x: 6.2, y: 0.12, w: 2.2, h: 0.22, fontSize: 8, bold: true, color: C.accent1, align: "center", valign: "middle" }); };
const src = (s, txt) => T(s, txt, { x: 0.5, y: 5.05, w: 9.0, h: 0.2, fontSize: 7.5, color: C.accent3, italic: true });
const logo = (s, key, label, x, y, w, h) => { const p = CH + "logos/" + key + ".png"; if (fs.existsSync(p)) { s.addImage({ path: p, x, y, w, h, sizing: { type: "contain", w, h } }); } else { s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: "FFFFFF" }, line: { color: "D9DCE1", width: 0.75 } }); T(s, label, { x, y, w, h, fontSize: 9, bold: true, color: C.text2, align: "center", valign: "middle", fontFace: "Cambria" }); } };
const f0 = (v) => (v < 0 ? "(" + Math.round(-v).toLocaleString("en-US") + ")" : Math.round(v).toLocaleString("en-US"));
const f1 = (v) => v.toFixed(1);

// ===================================================================== model numbers (same inputs as the workbook; recomputed here for the tables)
const K = 0.0887, g = 0.02, t = 0.27, sh = 308.495, price = 34.47, mcap = price * sh, nd = 2821.9, mino = 11.5;
const yrs = ["2026e", "2027e", "2028e", "2029e", "2030e", "2031e", "2032e"];
const sales = [20578, 21938, 22949, 23856.9, 24645.2, 25298.8, 25804.8];
const ebitdaal = [1379.4, 1538.7, 1668.8, 1734.9, 1792.2, 1839.7, 1876.5];
const ebita = [1305, 1466, 1599, 1662.3, 1717.2, 1762.7, 1798.0];
const capex = [-144.0, -153.6, -160.6, -140.1, -119.6, -99.1, -78.5];
const wcrPct = 2751.6 / 19414.6; const salesPrev = [19414.6].concat(sales.slice(0, 6));
const dwcr = sales.map((v, i) => -(v - salesPrev[i]) * wcrPct);
const tax = ebita.map(v => -v * t), nopat = ebita.map((v, i) => v + tax[i]), da = ebitdaal.map((v, i) => v - ebita[i]);
const fcf = nopat.map((v, i) => v + da[i] + capex[i] + dwcr[i]);
const per = yrs.map((_, i) => 62 / 365 + i), df = per.map(p => 1 / Math.pow(1 + K, p)), pv = fcf.map((v, i) => v * df[i]);
const sumpv = pv.reduce((a, b) => a + b, 0), tv = fcf[6] * (1 + g) / (K - g), pvtv = tv * df[6], ev = sumpv + pvtv, eq = ev - nd - mino, vps = eq / sh;

// ===================================================================== 1 TITLE
pres.addSection({ title: "Opening" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" });
T(s, "Rexel", { x: 0.6, y: 1.1, w: 6.0, h: 1.0, fontSize: 54, bold: true, color: C.background1, fontFace: "Cambria" });
T(s, "Should an investor acquire a stake?", { x: 0.6, y: 2.1, w: 8.8, h: 0.6, fontSize: 24, color: C.accent1, fontFace: "Cambria" });
T(s, "Valuation of a listed company – listed peers, DCF and LBO bid", { x: 0.6, y: 2.75, w: 8.8, h: 0.4, fontSize: 14, color: "CADCFC" });
s.addShape(pres.ShapeType.roundRect, { x: 7.1, y: 1.15, w: 2.3, h: 0.95, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" }, rectRadius: 0.06 });
logo(s, "rexel", "Rexel", 7.25, 1.25, 2.0, 0.75);
s.addShape(pres.ShapeType.rect, { x: 0.6, y: 3.65, w: 8.8, h: 0.9, fill: { color: "13315C" }, line: { color: "13315C" } });
T(s, [{ text: "Marwan Ferreira da Silva  ·  Douae El Adloun", options: { bold: true, breakLine: true } }, { text: "Introduction to Corporate Finance – Université Paris Dauphine-PSL, M2 104 Research in Finance  ·  Prof. David Heller", options: { breakLine: true } }, { text: "Valuation date: 30 October 2026  ·  Share price €34.47  ·  Submitted 2 November 2026", options: {} }], { x: 0.8, y: 3.72, w: 8.4, h: 0.8, fontSize: 11, color: C.background1, valign: "middle" });
s.addNotes("Title slide.");

// ===================================================================== 2 EXEC SUMMARY
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
s.addText("Executive summary – Buy: three methods converge 10–45% above a share price depressed by deal overhang", { placeholder: "title" });
card(s, 0.5, 1.15, 4.2, 1.0, NAVY);
T(s, "BUY – acquire a stake", { x: 0.65, y: 1.2, w: 3.9, h: 0.42, fontSize: 19, bold: true, color: C.background1, fontFace: "Cambria" });
T(s, "Target range €43–46 per share (+25% to +35%): DCF base case to peer xEBITDA median. The LBO bid (€37.6) is the floor.", { x: 0.65, y: 1.6, w: 3.9, h: 0.5, fontSize: 10, color: "CADCFC" });
stat(s, 0.5, 2.28, 2.0, "€34.47", "Share price 30 Oct 2026, below the €34.75 issue price");
stat(s, 2.7, 2.28, 2.0, "€43.1", "DCF base case – WACC 8.9%, g 2.0% (+25%)", COPPER);
stat(s, 0.5, 3.38, 2.0, "€46–50", "Listed peers, xEBITDA median to average (+34–46%)", COPPER);
stat(s, 2.7, 3.38, 2.0, "€37.6", "Maximum LBO bid at 20% IRR (+9%): the floor");
s.addImage({ path: CH + "02_football_field.png", x: 4.95, y: 1.15, w: 4.55, h: 3.25 });
s.addTable([hdrRow(["Market cap", "Net debt + min.", "EV", "EV / EBITDAaL 26e", "EV / adj. EBITA 26e", "P/E 26e", "Dividend yield", "FCF yield 26e"], "c"),
  row(["€10.6bn", "€2.8bn", "€13.5bn", "9.8x", "10.3x", "14.9x", "3.5%", "4.8%"], { align: "center", bold: true, fontSize: 9.5 })], tableOpts(0.5, 4.5, 9.0, [1.1, 1.15, 1.0, 1.3, 1.3, 1.0, 1.05, 1.1], { rowH: 0.22 }));
T(s, "Market-implied: 9.8x 2026e EBITDA after leases vs 13–14x for US distributors – a European discount that the fundamentals (best growth since 2022, margin on the way to 7%) do not justify.", { x: 0.5, y: 4.97, w: 9.0, h: 0.3, fontSize: 8.5, color: C.text2 });
s.addNotes("Pyramid: the answer first. Target range = DCF base (43) to peers xEBITDA median (46). Key multiples: EV 13,467 = 10,634 + 2,822 + 12; EBITDAaL 2026e 1,379; adj. EBITA 1,305; NI 713; DPS 1.20; FCF after interest and tax 2026e 513.");

// ===================================================================== 3 MARKET
pres.addSection({ title: "Market and company" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Market overview – a consolidating distribution market with structural electrification tailwinds", { placeholder: "title" });
draft(s, "Douae");
head(s, "Largest electrical distributors – latest annual sales (bn, reported currency)", 0.5, 1.1, 4.6);
s.addChart(pres.ChartType.bar, [{ name: "Sales", labels: ["Sonepar (€)", "Wesco ($, LTM)", "Rexel (€)", "Graybar ($)", "CED ($, est.)"], values: [33.6, 25.0, 19.4, 11.0, 8.0] }],
  { x: 0.5, y: 1.4, w: 4.6, h: 2.4, barDir: "bar", chartColors: [STEEL], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: THEME.colors.dk1, dataLabelFontFace: "+mn-lt", catAxisLabelFontSize: 9, catAxisLabelColor: THEME.colors.dk1, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin", barGapWidthPct: 60 });
s.addTable([hdrRow(["US electrical distribution", "Share"]), row(["Top-5 national chains (Wesco, Sonepar, Graybar, Rexel, CED)", "≥ 46%"]), row(["Top-10 distributors", "c. 56%"]), row(["Top-20 distributors", "c. 64%"]), row(["Distributors acquired 2021–2025", "125+"])], tableOpts(0.5, 3.9, 4.6, [3.6, 1.0], { rowH: 0.21, fontSize: 8.5 }));
head(s, "Demand drivers", 5.4, 1.1, 4.1);
const drivers = [["Data centres & AI", "Triple-digit growth at Rexel US in Q3 2026; 13% of US sales, 25% of US backlog. Wesco data-centre sales +45% y/y."], ["Energy transition", "HVAC, solar and EV charging > 20% of Rexel's European sales, +13% in Q3 2026; electrification plans and energy-price volatility."], ["Grid & industrial automation", "Utility capex, reshoring and industrial projects; industrial automation +20% at Rexel US in Q3."]];
drivers.forEach((d, i) => { const y = 1.42 + i * 0.82; card(s, 5.4, y, 4.1, 0.72); s.addShape(pres.ShapeType.ellipse, { x: 5.52, y: y + 0.18, w: 0.36, h: 0.36, fill: { color: COPPER }, line: { color: COPPER } }); T(s, String(i + 1), { x: 5.52, y: y + 0.18, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" }); T(s, d[0], { x: 6.0, y: y + 0.07, w: 3.4, h: 0.25, fontSize: 10.5, bold: true, color: C.text2 }); T(s, d[1], { x: 6.0, y: y + 0.32, w: 3.4, h: 0.4, fontSize: 8.5, color: C.text1 }); });
head(s, "Risks", 5.4, 3.95, 4.1);
T(s, "Cyclical exposure to non-residential construction · copper (cables c.16% of sales, copper c.60% of cable) · pricing pressure and digital disintermediation · interest rates.", { x: 5.4, y: 4.22, w: 4.1, h: 0.7, fontSize: 9.5 });
src(s, "Sources: Sonepar 2025 results; Wesco LTM (Yahoo Finance); Rexel FY 2025 and Q3 2026 releases; Electrical Wholesaling 2026 Electrical Pyramid; Axia Growth distribution M&A 2026. [Douae: add market size + growth with a cited report or Rexel's URD chapter 1 / CMD]");
s.addNotes("DOUAE – to complete: (1) market size and expected growth with a citable source (Rexel URD 2025 chapter 1 'addressable market', Capital Markets Day deck, or a market report); (2) one sentence on the three end-markets (residential, non-residential, industrial) and their 2026 trend; (3) check Graybar/CED figures (press or MDM top distributors list). Keep the chart; replace '(est.)' once sourced. Remove the DRAFT tag when done.");

// ===================================================================== 4 COMPETITIVE LANDSCAPE
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Competitive landscape – Rexel is the #2 global distributor; Sonepar is the only rival of comparable reach", { placeholder: "title" });
draft(s, "Douae");
const comp = [["sonepar", "Sonepar", "Private (family)", "€33.6bn", "40", "#1 worldwide; US #1 via Sonepar USA; digital 'Spark' programme"], ["wesco", "Wesco", "Listed (NYSE)", "$25.0bn", "c.50", "US-centric; data-centre, utility and datacom supply chain"], ["rexel", "Rexel", "Listed (Euronext)", "€19.4bn", "17", "#2 worldwide; 51% North America; value-added services, GCG"], ["graybar", "Graybar", "Private (employees)", "$11bn", "US/Canada", "Electrical, comms and data networking, US #3"], ["ced", "CED", "Private", "c.$8bn", "US", "Decentralised branch model, 700+ locations"], ["ahlsell", "Ahlsell · Würth · Solar", "Private / listed", "€3–5bn each", "Regional", "Nordics, DACH and Benelux leaders"]];
const RH = 0.44, TY = 1.15, HH = 0.24;
s.addTable([hdrRow(["", "Company", "Ownership", "Sales", "Countries", "Positioning"])].concat(comp.map(c => row(["", c[1], c[2], c[3], c[4], c[5]], { align: "left", fontSize: 8.5, bold: c[0] === "rexel", color: c[0] === "rexel" ? COPPER : THEME.colors.dk1 }))), tableOpts(0.5, TY, 6.0, [0.95, 1.05, 1.0, 0.75, 0.65, 1.6], { rowH: RH, fontSize: 8.5 }));
comp.forEach((c, i) => logo(s, c[0], c[1], 0.56, TY + HH + RH * i + 0.07, 0.83, RH - 0.14));
head(s, "How Rexel competes", 6.8, 1.1, 2.7);
bullets(s, [{ text: "Value-added services: ", options: { bold: true } }, { text: "moving upstream (engineering, kitting, project support); GCG brings >75% value-added content.", options: { bullet: false } },
  { text: "Digital: ", options: { bold: true } }, { text: "36% of Q3 2026 sales online (44% in Europe); AI tools lifted North America by +600 bp y/y.", options: { bullet: false } },
  { text: "High-growth verticals: ", options: { bold: true } }, { text: "data centres, utilities, automation; North American backlog > $3bn (+42% vs end-2025).", options: { bullet: false } },
  { text: "Bolt-on M&A: ", options: { bold: true } }, { text: "four North American deals in 2026 (TC 360, Revere, DEE, GCG); 2–3 pts of growth a year.", options: { bullet: false } }], { x: 6.8, y: 1.4, w: 2.7, h: 3.0, fontSize: 9 });
T(s, "Scale matters: purchasing terms, logistics density and digital investment favour the top five, which keep gaining share from the long tail – the source of Rexel's 2–3 pts of annual M&A growth.", { x: 0.5, y: 4.35, w: 6.0, h: 0.55, fontSize: 9, color: C.text2 });
src(s, "Sources: company reports (Sonepar, Wesco, Rexel, Graybar), Rexel Q3 2026 presentation, GCG release; CED and regional players: trade press estimates. Logos are company trademarks (Wikimedia Commons). [Douae: verify the private players' figures]");
s.addNotes("DOUAE – to complete: verify the private players' sales (Graybar ~$11bn, CED ~$8bn, Ahlsell, Würth Elektrogroßhandel, Solar A/S) and Wesco's country count; add a market-share figure for Rexel in France (#1/#2) and in the US if available. Remove the DRAFT tag when done.");

// ===================================================================== 5 COMPANY
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Market and company" });
s.addText("Rexel at a glance – €19.4bn of sales, 1,876 branches in 17 countries, repositioning toward high-growth segments", { placeholder: "title" });
draft(s, "Douae");
head(s, "Sales by geography, FY 2025", 0.5, 1.1, 3.0);
s.addChart(pres.ChartType.doughnut, [{ name: "Sales", labels: ["North America", "Europe", "Asia-Pacific"], values: [46, 48, 6] }], { x: 0.4, y: 1.3, w: 3.1, h: 1.9, holeSize: 55, chartColors: [COPPER, NAVY, GREY], showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 9, dataLabelFontFace: "+mn-lt", showLegend: true, legendPos: "b", legendFontSize: 8, legendFontFace: "+mn-lt", legendColor: THEME.colors.dk1, showTitle: false });
s.addTable([hdrRow(["Segment", "Sales €m", "Adj. EBITA %"]), row(["North America", "8,907", "7.3%"]), row(["Europe", "9,403", "5.7%"]), row(["Asia-Pacific", "1,105", "1.1%"]), row(["Group", "19,415", "6.0%"], { bold: true })], tableOpts(0.5, 3.25, 3.0, [1.3, 0.9, 0.8], { rowH: 0.22 }));
head(s, "Key figures (€m)", 3.8, 1.1, 5.7);
s.addTable([hdrRow(["", "2022", "2023", "2024", "2025", "H1 26"]), row(["Sales", "18,702", "19,153", "19,285", "19,415", "9,989"]), row(["Reported EBITDA", "1,681", "1,582", "1,516", "1,553", "835"]), row(["Adj. EBITA (margin)", "n.a.", "n.a.", "1,137 (5.9%)", "1,158 (6.0%)", "615 (6.2%)"]), row(["Net income", "922", "775", "341", "591", "342"]), row(["Diluted EPS (€)", "3.04", "2.57", "1.13", "1.99", "–"]), row(["Net financial debt", "–", "–", "2,484", "2,631", "3,322"]), row(["Net debt / EBITDAaL", "–", "–", "–", "2.0x", "2.4x"])], tableOpts(3.8, 1.4, 5.7, [1.7, 0.75, 0.75, 0.85, 0.85, 0.8], { rowH: 0.21, fontSize: 8.5 }));
T(s, "2024 net income hit by the €124m French Competition Authority fine; 2025 operating income €1,062m (+26%).", { x: 3.8, y: 3.12, w: 5.7, h: 0.22, fontSize: 7.5, color: C.accent3, italic: true });
head(s, "Business", 3.8, 3.35, 2.8);
bullets(s, ["Multichannel B2B distributor for residential, non-residential and industrial customers; 26,306 employees; US 38% of sales, France 19%.", "Warehouse sales 81%, direct 19%; cables c.16%; digital 36% of Q3 2026 sales.", "Free float 74%, Cevian Capital 23%, employees 1%; Ba1 / BB+ stable."], { x: 3.8, y: 3.62, w: 2.8, h: 1.4, fontSize: 8.5 });
head(s, "Strategy – Axelerate 2028", 6.8, 3.35, 2.7);
bullets(s, ["Sales growth 5–8% a year (M&A 2–3 pts); adj. EBITA margin > 7%; FCF conversion c.65%; payout ≥ 40%.", "2026 guidance, raised twice: growth c.7%, margin ≥ 6.4%, FCF conversion > 65%.", "GCG (Chicago): $1.4bn EV, > $1.1bn sales, c.11% EBITA margin, < 8x EBITDAaL post synergies, closing end-2026."], { x: 6.8, y: 3.62, w: 2.7, h: 1.4, fontSize: 8.5 });
src(s, "Sources: Rexel FY 2025 release and financial statements, H1 2026 and Q3 2026 releases, GCG release, shareholding page (31 Dec 2025); 2022–23 from Yahoo Finance.");
s.addNotes("DOUAE – to complete: two lines on history (spun out of PPR in 2005, IPO 2007) and the CEO's repositioning (Guillaume Texier, since 2021); confirm the Axelerate 2028 targets from the Capital Markets Day deck. Remove the DRAFT tag when done.");

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
s.addImage({ path: CH + "01_stock_performance.png", x: 0.5, y: 1.1, w: 6.2, h: 2.85 });
head(s, "Reading", 6.9, 1.1, 2.6);
bullets(s, ["+37% over two years vs +93% for Wesco and +7% for the CAC 40: Rexel has captured the electrification theme, but far less than its US peer.", "Since the GCG announcement (25 Sep) the share has fallen below the €34.75 placement price, despite the Q3 beat (sales 4.8% above consensus).", "Sell-side: Buy consensus, 13 analysts, mean target €41.9 (range €30–46).", "Beta 1.25 (2-year weekly vs CAC 40); 0.98 on Yahoo (5-year monthly)."], { x: 6.9, y: 1.4, w: 2.6, h: 2.6, fontSize: 9 });
head(s, "Share data and enterprise value build-up (30 Oct 2026)", 0.5, 4.0, 9.0);
s.addTable([hdrRow(["Price", "52-week", "Shares out.", "Market cap", "+ Net debt", "+ Minorities", "= EV", "EV / EBITDAaL 25A", "EV / EBITDAaL 26e", "EV / adj. EBITA 26e"], "c"),
  row(["€34.47", "€28.83–39.80", "308.5m", "€10,634m", "€2,822m", "€12m", "€13,467m", "11.0x", "9.8x", "10.3x"], { align: "center", bold: true, fontSize: 8.5 })], tableOpts(0.5, 4.27, 9.0, [0.7, 1.05, 0.8, 0.95, 0.85, 0.8, 0.9, 1.0, 1.0, 0.95], { rowH: 0.22 }));
T(s, "Shares: 296,997,254 + 14,388,490 placed at €34.75 (29 Sep 2026) – 2,890,248 treasury. Net debt: €3,322m at 30 June 2026 less €500m proceeds. EBITDAaL 2025A €1,227m, 2026e €1,379m.", { x: 0.5, y: 4.84, w: 9.0, h: 0.2, fontSize: 7, color: C.accent3, italic: true });
src(s, "Sources: Yahoo Finance daily closes (price return, local currency); Rexel press releases for events and share count; analyst targets per Yahoo Finance, 30 Oct 2026.");

// ===================================================================== 8 COMPS TABLE
pres.addSection({ title: "Valuation" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("Listed peers – five B2B distributors; Rexel trades 2–4 turns below the sector median on every multiple", { placeholder: "title" });
[["wesco", "Wesco"], ["grainger", "Grainger"], ["ferguson", "Ferguson"], ["watsco", "Watsco"], ["bunzl", "Bunzl"]].forEach((l, i) => logo(s, l[0], l[1], 0.5 + i * 1.25, 1.1, 1.1, 0.36));
s.addTable([hdrRow(["Company", "Mkt cap (bn)", "Net debt (bn)", "EV/EBITDA LTM", "EV/EBITDA FY+1", "EV/EBIT LTM", "EV/EBIT FY+1", "P/E LTM", "P/E FY+1"]),
  row(["Wesco (US)", "$16.7", "$5.2", "13.8x", "12.4x", "16.2x", "14.6x", "23.5x", "16.4x"]), row(["Grainger (US)", "$64.2", "$1.9", "21.9x", "19.5x", "23.9x", "21.4x", "34.3x", "26.6x"]), row(["Ferguson (US/UK)", "$46.2", "$3.5", "13.9x", "12.9x", "15.6x", "14.4x", "19.6x", "18.8x"]), row(["Watsco (US)", "$10.9", "$0.1", "15.5x", "13.4x", "16.5x", "14.2x", "23.0x", "23.6x"]), row(["Bunzl (UK)", "£8.9", "£1.6", "8.8x", "8.3x", "13.2x", "12.5x", "19.4x", "14.5x"]),
  row(["Sector median", "", "", "13.9x", "12.9x", "16.2x", "14.4x", "23.0x", "18.8x"], { bold: true, fill: { color: LT2 } }), row(["Sector average", "", "", "14.8x", "13.3x", "17.1x", "15.4x", "24.0x", "20.0x"], { bold: true, fill: { color: LT2 } }),
  row(["Rexel at market", "€10.6", "€2.8", "11.0x", "9.8x", "11.6x", "10.3x", "15.7x", "14.9x"], { bold: true, color: COPPER })], tableOpts(0.5, 1.55, 6.1, [1.2, 0.65, 0.65, 0.65, 0.65, 0.6, 0.6, 0.55, 0.55], { rowH: 0.215, fontSize: 8.5 }));
T(s, "EV = market cap + net debt (Yahoo Finance, 30 Oct 2026). US GAAP peers expense operating leases above EBITDA, so Rexel is measured on EBITDA after leases (EBITDAaL) and adjusted EBITA, with lease liabilities excluded from net debt. FY+1 EBITDA/EBIT for peers = next-year revenue consensus × LTM margin; FY+1 net income = consensus EPS. Rexel aggregates: 2025A and consensus 2026e (13 brokers, Oct 2026).", { x: 0.5, y: 3.78, w: 6.1, h: 0.62, fontSize: 7.2, color: C.accent3 });
head(s, "Sample", 0.5, 4.4, 6.1);
T(s, "Listed B2B distributors with comparable models: electrical (Wesco), MRO (Grainger), plumbing/HVAC (Ferguson, Watsco), consumables (Bunzl). Sonepar and Graybar, the closest comparables, are private. Fastenal (29x EBITDA), Legrand and Schneider (manufacturers) are excluded.", { x: 0.5, y: 4.63, w: 6.1, h: 0.42, fontSize: 7.8 });
card(s, 6.85, 1.1, 2.65, 3.95, NAVY);
T(s, "Implied value of Rexel (€ per share)", { x: 7.0, y: 1.2, w: 2.4, h: 0.3, fontSize: 10.5, bold: true, color: C.background1 });
const imp = [["xEBITDA – anchor", "46 – 50", true], ["xEBIT", "52 – 56", false], ["P/E", "43 – 53", false], ["xSales", "77 – 95 (n.m.)", false]];
imp.forEach((r, i) => { const y = 1.62 + i * 0.58; T(s, r[0], { x: 7.0, y, w: 1.3, h: 0.5, fontSize: 10, color: "CADCFC", valign: "middle" }); T(s, r[1], { x: 8.1, y, w: 1.35, h: 0.5, fontSize: r[2] ? 16 : 13, bold: r[2], color: r[2] ? "F2A86F" : C.background1, align: "right", valign: "middle" }); });
T(s, "Worked example (xEBITDA FY+1, median): 12.9x × €1,379m = €17,766m EV; – €2,833m net debt and minorities = €14,933m equity; ÷ 308.5m shares = €48.4.", { x: 7.0, y: 4.0, w: 2.4, h: 1.0, fontSize: 8.5, color: "CADCFC" });
src(s, "Sources: Yahoo Finance (30 Oct 2026), company reports, Rexel consensus (15 Oct 2026). Model: Comps sheet. Logos are company trademarks.");

// ===================================================================== 9 COMPS READING
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("Reading the multiples – xEBITDA is the anchor; the sales multiple is meaningless and the discount is a listing discount", { placeholder: "title" });
s.addImage({ path: CH + "04_peer_benchmark.png", x: 0.5, y: 1.1, w: 5.6, h: 2.8 });
head(s, "Why the multiples disagree", 6.3, 1.1, 3.2);
bullets(s, [{ text: "xSales overvalues Rexel (€77–95). ", options: { bold: true } }, { text: "Peers earn 10–23% EBITDA margins on each euro of sales; Rexel earns 6%. A sales multiple built on richer margins has no meaning here.", options: { bullet: false } },
  { text: "xEBIT above xEBITDA. ", options: { bold: true } }, { text: "Rexel's D&A after leases is only 6% of EBITDAaL (asset-light, leases already deducted), lower than peers', so an EBIT multiple flatters it slightly.", options: { bullet: false } },
  { text: "P/E is wide (€43–53). ", options: { bold: true } }, { text: "Net income is after interest and tax: Rexel's 2.4x leverage and 30% 2025 tax rate are not comparable with net-cash US peers.", options: { bullet: false } },
  { text: "Keep xEBITDA: €46–50. ", options: { bold: true } }, { text: "Same margin as Wesco, three turns cheaper: the gap is a European listing and growth-history discount, widened by deal overhang, not an operating gap.", options: { bullet: false } }], { x: 6.3, y: 1.4, w: 3.2, h: 3.6, fontSize: 9 });
s.addTable([hdrRow(["Implied equity (€m) and € per share", "Sector median", "Sector average"]), row(["xEBITDA LTM: 13.9x × 1,227 – 2,833", "14,229 → €46.1", "15,293 → €49.6"]), row(["xEBITDA FY+1: 12.9x × 1,379 – 2,833", "14,932 → €48.4", "15,520 → €50.3"]), row(["xEBIT FY+1: 14.4x × 1,305 – 2,833", "16,017 → €51.9", "17,305 → €56.1"]), row(["P/E FY+1: 18.8x × 713", "13,410 → €43.5", "14,250 → €46.2"])], tableOpts(0.5, 4.0, 5.6, [2.8, 1.4, 1.4], { rowH: 0.2, fontSize: 8 }));
src(s, "Sources: Yahoo Finance LTM figures, 30 Oct 2026; Rexel FY 2025 release and consensus. Model: Comps sheet (median and average of the five peers).");

// ===================================================================== 10 DCF ASSUMPTIONS
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("DCF assumptions – Rexel's own broker consensus to 2028, soft landing to 2% growth, WACC of 8.9%", { placeholder: "title" });
head(s, "Discount rate", 0.5, 1.1, 4.1);
s.addTable([hdrRow(["Input", "Value", "Source / rationale"]),
  row(["Risk-free rate (rf)", "4.13%", "50% Bund 3.33% + 50% UST 4.94%"]), row(["Market risk premium", "5.0%", "Damodaran, France, Jan 2026"]), row(["Beta (b)", "1.20", "2-yr regression 1.25; Yahoo 0.98"]), row(["Cost of equity k = rf + b·MRP", "10.1%", "CAPM"]), row(["Cost of debt i = rf + spread", "5.6%", "150 bp spread (Ba1 / BB+)"]), row(["Tax rate (t)", "27%", "Normative (2025: 30% with surtax)"]), row(["D / (D + E) at market", "21.0%", "€2.8bn / (€2.8bn + €10.6bn)"]), row(["WACC = k·E/V + i(1–t)·D/V", "8.87%", "Rexel impairment tests: 8.4–9.8%"], { bold: true, fill: { color: LT2 } }), row(["Perpetual growth g", "2.0%", "Rexel uses 1.9–2.2%"], { bold: true, fill: { color: LT2 } })],
  tableOpts(0.5, 1.37, 4.1, [1.6, 0.55, 1.95], { rowH: 0.215, fontSize: 8 }));
head(s, "Cross-check (course method, Modigliani–Miller)", 0.5, 3.6, 4.1);
s.addTable([hdrRow(["Step", "Formula", "Value"]), row(["Unlevered beta b*", "b / (1 + D(1–t)/E) = 1.20 / 1.194", "1.00"]), row(["Unlevered cost r", "rf + b*·MRP", "9.16%"]), row(["k (MM)", "r + (r – i)(1–t)·D/E", "9.85%"]), row(["K (MM)", "r · (1 – D·t / (D+E))", "8.64%"])], tableOpts(0.5, 3.86, 4.1, [1.1, 2.3, 0.7], { rowH: 0.2, fontSize: 8 }));
head(s, "Free cash flow build (€m)", 4.85, 1.1, 4.65);
const fcfRows = [hdrRow([""].concat(yrs)), row(["Sales"].concat(sales.map(f0))), row(["growth"].concat(sales.map((v, i) => "+" + (100 * (v / salesPrev[i] - 1)).toFixed(1) + "%")), { color: GREY, italic: true }), row(["Adj. EBITA (EBIT)"].concat(ebita.map(f0))), row(["margin"].concat(ebita.map((v, i) => (100 * v / sales[i]).toFixed(1) + "%")), { color: GREY, italic: true }), row(["– Tax at 27%"].concat(tax.map(f0))), row(["= NOPAT"].concat(nopat.map(f0))), row(["+ D&A"].concat(da.map(f0))), row(["– Capex"].concat(capex.map(f0))), row(["– Δ WCR"].concat(dwcr.map(f0))), row(["= Free cash flow"].concat(fcf.map(f0)), { bold: true, fill: { color: LT2 } }), row(["Discount period (yrs)"].concat(per.map(v => v.toFixed(2))), { color: GREY, italic: true }), row(["Discount factor"].concat(df.map(v => v.toFixed(3))), { color: GREY, italic: true }), row(["Present value"].concat(pv.map(f0)), { bold: true })];
s.addTable(fcfRows, tableOpts(4.85, 1.37, 4.65, [1.29, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48], { rowH: 0.2, fontSize: 7 }));
T(s, "2026–28: consensus mean collected by Rexel (13 brokers, 5–13 Oct 2026, excl. GCG, before the 22 Oct guidance raise). 2029–32: growth converges linearly to g; EBITDAaL margin frozen at 7.3%; capex converges to D&A; WCR 14.2% of sales (2025 financial balance sheet).", { x: 4.85, y: 4.42, w: 4.65, h: 0.6, fontSize: 7, color: C.accent3, italic: true });
src(s, "Sources: Rexel consensus (15 Oct 2026), 2025 financial statements (note 12), Trading Economics (30 Oct 2026), Damodaran. Model: DCF sheet.");

// ===================================================================== 11 DCF RESULT
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("DCF result – €43 per share, 25% above the market; above the price across the whole reasonable grid", { placeholder: "title" });
head(s, "From free cash flows to value per share (€m)", 0.5, 1.1, 4.1);
s.addTable([hdrRow(["", "€m"]), row(["Sum of discounted FCF 2026–2032", f0(sumpv)]), row(["Terminal value = 1,241 × 1.02 / (8.87% – 2.0%)", f0(tv)]), row(["Discount factor 2032: 1 / 1.0887^6.17", df[6].toFixed(3)]), row(["Discounted terminal value (" + Math.round(100 * pvtv / ev) + "% of EV)", f0(pvtv)]), row(["Enterprise value", f0(ev)], { bold: true }), row(["– Net debt pro forma – minorities", "(2,833)"]), row(["Equity value", f0(eq)], { bold: true, fill: { color: LT2 } }), row(["÷ Shares outstanding (m)", "308.5"]), row(["Value per share", "€" + f1(vps)], { bold: true, fill: { color: LT2 }, color: COPPER }), row(["Upside vs €34.47", "+" + Math.round(100 * (vps / price - 1)) + "%"], { bold: true })], tableOpts(0.5, 1.37, 4.1, [3.0, 1.1], { rowH: 0.225, fontSize: 8.5 }));
s.addTable([hdrRow(["Sanity checks", "Value"]), row(["Implied EV / EBITDAaL 2026e (market 9.8x, peers 12.9x)", (ev / ebitdaal[0]).toFixed(1) + "x"]), row(["Implied exit multiple TV / EBITDAaL 2032", (tv / ebitdaal[6]).toFixed(1) + "x"]), row(["WACC implied by the share price (g = 2%)", "c. 10.0%"]), row(["Sell-side mean target", "€41.9"])], tableOpts(0.5, 3.95, 4.1, [3.0, 1.1], { rowH: 0.19, fontSize: 8 }));
s.addImage({ path: CH + "05_dcf_sensitivity.png", x: 4.8, y: 1.1, w: 4.7, h: 3.0 });
T(s, "Every cell with K ≤ 9.5% and g ≥ 1.5% is above €34.47. The terminal value is 68% of EV – normal for a distributor with 2% perpetual growth, which makes K and g the two numbers to challenge: ±50 bp of WACC moves the value by about €4 per share.", { x: 4.8, y: 4.2, w: 4.7, h: 0.8, fontSize: 9, color: C.text1 });
src(s, "Model: DCF sheet, valuation date 30 Oct 2026, full-year 2026 FCF discounted 0.17 year. Net debt at 30 June 2026 less the €500m capital increase.");

// ===================================================================== 12 LBO STRUCTURE
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("LBO-style bid – a sponsor buys 100% through a holding financed 48% with debt, exits after three years", { placeholder: "title" });
const box = (x, y, w, h, fill, title, body, tcol) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.06 }); T(s, title, { x: x + 0.15, y: y + 0.07, w: w - 0.3, h: 0.26, fontSize: 10, bold: true, color: tcol || "FFFFFF" }); T(s, body, { x: x + 0.15, y: y + 0.35, w: w - 0.3, h: h - 0.4, fontSize: 8, color: tcol || "FFFFFF" }); };
box(0.5, 1.12, 4.0, 0.78, NAVY, "Investment fund", "Equity €6,075m (52% of sources) – required IRR 20%, 3-year hold, exit at the entry multiple of 10.3x EBIT.");
s.addShape(pres.ShapeType.downArrow, { x: 2.35, y: 1.95, w: 0.3, h: 0.25, fill: { color: GREY }, line: { color: GREY } });
box(0.5, 2.25, 4.0, 1.25, STEEL, "NewCo – acquisition debt €5,518m (4.0x EBITDAaL 2026e)", "Senior A €2,759m (50%), Euribor 2.2% + 2.75% + 0.5% swap = 5.45%, amortising over 7 years\nSenior B €1,931m (35%), 5.95%, bullet 8 years · Junior C €828m (15%), 8.70%, bullet 9 years\nTax consolidation: acquisition interest offsets Rexel's taxable income; 100% of net income upstreamed.");
s.addShape(pres.ShapeType.downArrow, { x: 2.35, y: 3.55, w: 0.3, h: 0.22, fill: { color: GREY }, line: { color: GREY } });
box(0.5, 3.82, 4.0, 0.68, LT2, "Rexel – 100% acquired at €11,593m (€37.6 per share)", "Existing net financial debt €2,822m stays in the target: total leverage 6.0x EBITDAaL at entry, in line with large-cap LBO practice.", THEME.colors.dk1);
T(s, "Sources €11,593m = fund equity 6,075 + Senior A 2,759 + Senior B 1,931 + Junior C 828. Uses = purchase of 100% of the equity.", { x: 0.5, y: 4.56, w: 4.0, h: 0.45, fontSize: 8, color: C.text2 });
head(s, "Acquisition debt schedule (€m)", 4.8, 1.1, 4.7);
s.addTable([hdrRow(["", "2026", "2027", "2028", "2029"]), row(["Senior A outstanding (year-end)", "2,759", "2,365", "1,971", "1,577"]), row(["  amortisation", "", "(394)", "(394)", "(394)"]), row(["  interest at 5.45% on average balance", "", "140", "118", "97"]), row(["Senior B €1,931m bullet – interest at 5.95%", "", "115", "115", "115"]), row(["Junior C €828m bullet – interest at 8.70%", "", "72", "72", "72"]), row(["Total acquisition interest", "", "327", "305", "284"], { bold: true, fill: { color: LT2 } })], tableOpts(4.8, 1.37, 4.7, [2.5, 0.55, 0.55, 0.55, 0.55], { rowH: 0.17, fontSize: 7 }));
head(s, "Holding income statement and cash (€m)", 4.8, 3.0, 4.7);
s.addTable([hdrRow(["", "2027", "2028", "2029"]), row(["Dividend received (100% of Rexel net income)", "954", "1,040", "1,080"]), row(["– Tax on dividend (5% × 27%)", "(13)", "(14)", "(15)"]), row(["+ Rexel's own tax, now paid at group level", "353", "385", "399"]), row(["– Interest on acquisition debt", "(327)", "(305)", "(284)"]), row(["– Tax of the fiscal group (27% × (EBT – interest))", "(265)", "(302)", "(323)"]), row(["= Holding net income", "703", "803", "858"], { bold: true }), row(["Cumulative cash after Senior A amortisation (394 p.a.)", "309", "718", "1,182"], { bold: true, fill: { color: LT2 } })], tableOpts(4.8, 3.25, 4.7, [2.8, 0.63, 0.63, 0.64], { rowH: 0.16, fontSize: 7 }));
T(s, "Model: LBO sheet (course method, Part 5). Rexel's plan: consensus 2026–28 then soft landing; interest on its own debt at 5.6%. Euribor 6M 2.2% to verify.", { x: 0.5, y: 4.95, w: 4.1, h: 0.3, fontSize: 7, color: C.accent3, italic: true });

// ===================================================================== 13 LBO RESULT
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Valuation" });
s.addText("LBO result – a fund can bid at most €37.6 per share for a 20% IRR: a 9% premium, too thin for a public-to-private", { placeholder: "title" });
head(s, "Returns and maximum bid", 0.5, 1.1, 4.2);
s.addTable([hdrRow(["", "€m"]), row(["Rexel EBITA 2029e (consensus + soft landing)", "1,667"]), row(["× Entry multiple: EV at market 13,222 ÷ EBITA 2026e 1,305", "10.3x"]), row(["= Enterprise value at exit (2029)", "17,186"], { bold: true }), row(["– Rexel net debt 2029 (3,535) – holding debt (4,336) + cash (1,182)", "(6,688)"]), row(["= Equity value at exit", "10,498"], { bold: true }), row(["÷ (1 + 20%)³ = fund equity at entry", "6,075"]), row(["+ Acquisition debt", "5,518"]), row(["= Maximum price for 100% of the equity", "11,593"], { bold: true, fill: { color: LT2 } }), row(["÷ 308.5m shares = bid per share", "€37.6"], { bold: true, fill: { color: LT2 }, color: COPPER }), row(["Premium vs €34.47 / money multiple / equity share", "+9% / 1.73x / 52%"], { bold: true })], tableOpts(0.5, 1.37, 4.2, [3.1, 1.1], { rowH: 0.235, fontSize: 8.5 }));
head(s, "Sensitivity of the bid (€ per share)", 5.0, 1.1, 4.5);
s.addTable([hdrRow(["IRR \\ acquisition debt", "3.5x", "4.0x", "4.5x", "5.0x"]), row(["15%", "39.7", "40.3", "40.8", "41.4"]), row(["20%", "36.8", "37.6", "38.4", "39.1"], { bold: true, fill: { color: LT2 } }), row(["25%", "34.4", "35.3", "36.3", "37.2"])], tableOpts(5.0, 1.37, 4.5, [1.7, 0.7, 0.7, 0.7, 0.7], { rowH: 0.24 }));
s.addTable([hdrRow(["Exit multiple (xEBIT), 4.0x / 20%", "9.0x", "10.3x", "11.0x", "12.0x"]), row(["Bid per share", "33.5", "37.6", "39.7", "42.9"])], tableOpts(5.0, 2.5, 4.5, [1.7, 0.7, 0.7, 0.7, 0.7], { rowH: 0.24 }));
s.addTable([hdrRow(["Holding period at 20% IRR", "3 years", "5 years"]), row(["Bid per share", "37.6", "35.3"])], tableOpts(5.0, 3.15, 4.5, [1.7, 1.4, 1.4], { rowH: 0.24 }));
bullets(s, ["Only a re-rating at exit (12x EBIT) or a 15% return target gets a sponsor above €40: the exit multiple, not leverage, sets the ceiling.", "With 6.0x total leverage and Ba1/BB+ ratings already, more aggressive financing is not credible; a longer hold lowers the bid.", "Conclusion: an LBO at a normal 25–30% premium (€43–45) is not feasible. The stock is cheap for a long-term holder or a strategic buyer (Sonepar, Wesco) who can pay for synergies."], { x: 5.0, y: 3.8, w: 4.5, h: 1.25, fontSize: 8.5 });
src(s, "Model: LBO sheet and sensitivity replica. Interest on opening net debt; dividend tax 5% × 27%; Rexel net debt rises in the plan because 100% of income is upstreamed.");

// ===================================================================== 14 INTERPRETATION
pres.addSection({ title: "Conclusion" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Conclusion" });
s.addText("Interpretation – the methods are consistent with each other and with the sell-side; the market is the outlier", { placeholder: "title" });
s.addImage({ path: CH + "02_football_field.png", x: 0.5, y: 1.1, w: 4.4, h: 2.1 });
head(s, "Consistency checks", 5.1, 1.1, 4.4);
bullets(s, ["DCF (€43) sits inside the analyst range (€30–46, mean €41.9) and just below the peer xEBITDA band (€46–50): intrinsic and relative values agree.", "The share price implies a 10% WACC at 2% growth, or 9.8x EBITDAaL 2026e – outside Rexel's own range and the peers'.", "The LBO bid (€37.6) is a floor: a financial buyer earns 20% at that price, so the downside from €34.47 is cushioned.", "Peer FY+1 EBITDA/EBIT are proxies (consensus revenue × LTM margin); the median is used to limit sample bias."], { x: 5.1, y: 1.4, w: 4.4, h: 1.85, fontSize: 9 });
head(s, "Scenarios (€ per share)", 0.5, 3.28, 4.4);
s.addTable([hdrRow(["", "Bear", "Base", "Bull"]), row(["DCF: K / g", "9.5% / 1.5%", "8.9% / 2.0%", "8.5% / 2.5%"]), row(["DCF value", "36.6", "43.1", "49.5"]), row(["Peers xEBITDA", "46.1 (LTM median)", "48.4 (FY+1 median)", "50.3 (FY+1 average)"]), row(["LBO bid (20% IRR)", "35.3 (5-year hold)", "37.6", "42.9 (12x exit)"]), row(["Upside vs €34.47", "+2% to +34%", "+9% to +40%", "+24% to +46%"], { bold: true, fill: { color: LT2 } })], tableOpts(0.5, 3.53, 4.4, [1.2, 1.07, 1.06, 1.07], { rowH: 0.2, fontSize: 8 }));
head(s, "What would change the answer", 5.1, 3.28, 4.4);
s.addTable([hdrRow(["Risk", "Impact", "Watch item"]), row(["US cycle, data-centre concentration (13% of US sales)", "2028 margin 6.5% vs 7.0%: c.–€3", "Backlog $3bn (+42%)"]), row(["GCG integration ($1.4bn, closing end-2026)", "< 8x only post synergies", "Leverage 2.0x from 2027"]), row(["Copper and pricing (cables 16% of sales)", "Non-recurring margin swings", "Monthly copper effect"]), row(["French sovereign spread (OAT 4.97%)", "+50 bp WACC ≈ –€4", "51% of sales in NA"])], tableOpts(5.1, 3.53, 4.4, [1.9, 1.4, 1.1], { rowH: 0.2, fontSize: 7.5 }));
src(s, "Model: Summary, DCF and LBO sheets; Rexel Q3 2026 presentation; Trading Economics.");

// ===================================================================== 15 RECOMMENDATION
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("Investment recommendation – BUY: acquire a stake, target €43–46 per share (+25% to +35%)", { placeholder: "title" });
const col = (x, title, items, color) => { s.addShape(pres.ShapeType.roundRect, { x, y: 1.35, w: 2.85, h: 3.1, fill: { color: "13315C" }, line: { color: "13315C" }, rectRadius: 0.06 }); T(s, title, { x: x + 0.15, y: 1.45, w: 2.55, h: 0.35, fontSize: 13, bold: true, color: color }); bullets(s, items, { x: x + 0.15, y: 1.85, w: 2.55, h: 2.5, fontSize: 9.5, color: "FFFFFF" }); };
col(0.5, "Thesis", ["Best operating momentum since 2022: +11.4% same-day growth, guidance raised twice, margin on the way to 7%.", "Valuation gap: 9.8x EBITDAaL vs 13–14x for US distributors, with the same margin as Wesco.", "Balance sheet funded: the €500m raise keeps leverage at c.2.0x from 2027 while GCG adds a c.11%-margin platform."], "F2A86F");
col(3.575, "Catalysts", ["GCG closing by year-end and first synergy guidance.", "FY 2026 results (February 2027): delivery of the ≥6.4% margin and 2027 outlook including GCG.", "Consensus upgrades: the October consensus predates the Q3 beat and the guidance raise.", "Share price below the €34.75 placement: institutional support at that level."], "CADCFC");
col(6.65, "Risks", ["US data-centre cycle and tougher comparables in Q4 2026.", "GCG integration and synergy delivery; execution of four acquisitions in one year.", "Copper, pricing and European macro; French sovereign spread lifting the cost of capital.", "Stop-loss logic: a 2028 margin below 6.5% removes most of the upside."], "F2A86F");
T(s, "Sizing: a long-term equity stake at or below €37 (the LBO floor) with a €43–46 target offers an asymmetric 3:1 upside/downside; we would not pursue a leveraged buyout at a control premium.", { x: 0.5, y: 4.55, w: 9.0, h: 0.42, fontSize: 10, color: "FFFFFF", bold: true });
T(s, "Sources: Rexel releases and financial statements (FY 2025, H1 2026, Q3 2026, GCG, capital increase); Rexel broker consensus (15 Oct 2026); Yahoo Finance; Trading Economics; Damodaran (Jan 2026); Electrical Wholesaling. Full model and data on github.com/marwanferreira/rexel-valuation.", { x: 2.1, y: 5.0, w: 6.7, h: 0.28, fontSize: 7, color: "8DA9C4" });

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
