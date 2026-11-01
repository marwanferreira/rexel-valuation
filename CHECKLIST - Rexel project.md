# Rexel valuation project — checklist

Team: Marwan Ferreira da Silva + Douae El Adloun. Course: Intro to corporate finance, M2 104, Prof. David Heller.

Every line is tagged with where it comes from, so nobody mistakes a suggestion for a rule:

- **[PDF]** = written in the official instructions (`Instructions EN - Project to value a listed company 26-27 (2).pdf`). Non-negotiable.
- **[COURSE]** = not in the PDF, but it is how the prof did it in class (Part 4 Moncler / LVMH, Part 5 Carrefour LBO). He will expect it.
- **[REC]** = our own recommendation (good practice, not a rule). Drop it if short on time.

---

## 1. Hard rules — [PDF]

- [x] **Topic: Rexel.** Financial analysis of Rexel to recommend whether or not to acquire a stake. [PDF]
- [x] **Work in pairs.** [PDF]
- [x] **Format: PowerPoint converted to PDF.** [PDF]
- [x] **No more than 15 slides.** The PDF says "15 slides", not pages. Once converted, that is 15 PDF pages. [PDF]
    - [x] [REC] Title slide and sources slide count toward the 15 unless the prof says otherwise. Assume they count.
- [x] **Send to david.heller@dauphine.psl.eu.** [PDF]
- [x] **Deadline: before 3 November.** "Any delay will result in a penalty of points." [PDF]
    - [x] [REC] Send by the evening of Sunday 2 November at the latest. Do not wait for the 3rd.
    - [x] [REC] Put both names, the course name and "M2 104" on the title slide.
- [x] **Cite the sources of the data used.** [PDF]
- [x] **Each slide: essential information only, no visual overload.** [PDF]
- [x] **Use graphs, tables and diagrams to illustrate.** [PDF]

## 2. Required sections — [PDF]

The PDF lists these six, in this order. All six must be there.

- [x] **Market overview**: size, trends, competition, risks, opportunities. [PDF]
- [x] **Company overview**: main activities, operating segments, competitive advantages, strategic positioning. [PDF]
- [x] **Valuation results** from the methods used, with explanation of the assumptions: [PDF]
    - [x] DCF: justify growth rate, discount rate "etc." [PDF]
    - [x] Market comparables: justify the choice of companies in the sample. [PDF]
- [x] **LBO-style bid by an investment fund.** [PDF]
- [x] **Interpretation of results**: critical analysis, consistency with market data, sensitivity to assumptions. [PDF]
- [x] **Investment recommendation**: clear and structured argument, acquire a stake or not. [PDF]

## 3. Grading criteria — [PDF]

Check the deck against each one before sending.

- [x] Quality of market and company analysis. [PDF]
- [x] Relevance of assumptions and methodological rigour of valuations and bid. [PDF]
- [x] Clarity and accuracy of interpretation of results. [PDF]
- [x] Quality of investment recommendations. [PDF]
- [x] Visual presentation and clarity of supporting materials. [PDF]

---

## 4. What the prof expects inside each model — [COURSE]

Not written in the PDF, but this is exactly how he built the class cases. Reuse the workbooks in `Part 4/Excel/` and `Part 5/Excel/`.

### Comparables (Moncler method, Part 4)
- [x] EV of each peer = market cap + net debt. [COURSE]
- [x] Multiples: xSales, xEBITDA, xEBIT, P/E, for two years (current year + next). [COURSE]
- [x] Sector multiple = average of the peers. [COURSE]
- [x] Rexel equity = sector multiple × Rexel aggregate − Rexel net debt (P/E gives equity directly). [COURSE]
- [x] Compare to Rexel's market cap and compute the upside/downside. [COURSE]
- [x] Explain *why* the multiples disagree (margin vs peers, D&A intensity, volatile net income). [COURSE]
- [x] [REC] Use at least 3-4 peers (Moncler had only 2 and the prof flagged it as a weakness). Candidates: Sonepar is private, so think Wesco, Grainger, Legrand, Schneider, Bunzl, Fastenal. Justify each inclusion in one line.

### DCF (LVMH method, Part 4)
- [x] Business plan: broker consensus for the next 2-3 years, then "soft landing" where growth converges linearly to the long-term g. [COURSE]
- [x] FCF = EBIT × (1 − tax) + D&A − capex − ΔWCR. [COURSE]
- [x] Capex converges to D&A at steady state; WCR kept as a % of sales. [COURSE]
- [x] Cost of equity by CAPM: rf + β × MRP. Cost of debt = rf + spread. WACC weighted by market values. [COURSE]
- [x] Terminal value by Gordon-Shapiro: FCF × (1+g) / (WACC − g). [COURSE]
- [x] Bridge: EV − net debt = equity, then per share vs the share price. [COURSE]
- [x] [REC] g between 1.5 % and 2.5 %. Anything above 3 % will be challenged.
- [x] [REC] Sensitivity table: WACC on one axis, g on the other. This covers the PDF's "sensitivity to assumptions" line in one slide.
- [x] [REC] State the valuation date and the share price used on that date.

### LBO bid (Carrefour method, Part 5)
- [x] Holding company buys Rexel; sources = equity from the fund + senior debt (tranches A/B) + junior debt (C). [COURSE]
- [x] Target income statement, cash flow and balance sheet projected over the holding period. [COURSE]
- [x] Holding income statement with tax consolidation (interest on acquisition debt offsets the target's tax). [COURSE]
- [x] Debt schedule: repayment from the target's dividends / cash. [COURSE]
- [x] Exit value by a multiple, fund IRR target around 20 % over 3-5 years. [COURSE]
- [x] The bid = the maximum price per share the fund can pay and still reach its IRR, with a premium vs the current share price. [COURSE]
- [x] [REC] Show entry multiple, leverage (debt/EBITDA), exit multiple and IRR on one slide. That is the whole LBO story for a grader.

---

## 5. Data to collect — [REC]

Nothing is provided by the prof, so everything comes from public sources. Keep a list of sources as you go for the sources slide.

- [x] Rexel annual report / universal registration document (latest full year) and the last half-year report: sales, EBITDA, EBIT, net income, D&A, capex, WCR, net debt, number of shares. Source: rexel.com investor section.
- [x] Rexel share price and market cap on a fixed valuation date (Euronext Paris, ticker RXL).
- [x] Broker consensus for the next 2-3 years (sales, EBITDA, EBIT). Sources: Rexel's own consensus page if published, MarketScreener, Zonebourse.
- [x] Beta of Rexel (or sector unlevered beta re-levered, as in the Part 3 exam exercise).
- [x] Risk-free rate (10-year OAT or Bund), market risk premium, Rexel's credit spread (its bond yields or rating).
- [x] Peer financials and market caps on the same date as Rexel.
- [x] Market data: size of electrical distribution, growth drivers (electrification, data centres, energy transition), competitors (Sonepar, Wesco, Graybar), risks (cyclicality, copper prices, pricing pressure).
- [x] Typical LBO financing terms (leverage, interest rates by tranche) from the Part 5 case if nothing better is found.

## 6. Slide budget — [REC]

A way to fit everything in 15. Adjust freely.

| # | Slide | Section |
|---|-------|---------|
| 1 | Title, names, date | — |
| 2 | Executive summary: recommendation + football field chart | Recommendation |
| 3 | Market: size and trends | Market |
| 4 | Market: competition, risks, opportunities | Market |
| 5 | Company: activities and segments | Company |
| 6 | Company: competitive advantages, positioning, key financials | Company |
| 7 | Comparables: peer sample and multiples | Valuation |
| 8 | Comparables: Rexel implied value and reading | Valuation |
| 9 | DCF: assumptions (WACC build-up, g, business plan) | Valuation |
| 10 | DCF: FCF, terminal value, bridge, value per share | Valuation |
| 11 | DCF sensitivity table (WACC × g) | Interpretation |
| 12 | LBO: structure, sources and uses, debt schedule | LBO bid |
| 13 | LBO: returns, max bid price, premium | LBO bid |
| 14 | Interpretation: all methods vs market price, critical view | Interpretation |
| 15 | Recommendation and sources | Recommendation |

## 7. Before sending — [REC]

- [x] Count the pages of the exported PDF: 15 or fewer.
- [x] Every number on a slide traces back to a source or to the Excel model.
- [x] Same valuation date, same share count and same net debt used in all three methods.
- [x] Recommendation slide answers the exact question: acquire a stake, yes or no, and why.
- [x] File name is clean (e.g. `Rexel valuation - Ferreira Douae - M2 104.pdf`).
- [x] Both of you have read the final PDF, not just the PowerPoint.
- [x] Attach the Excel model to the email as a backup? Not required by the PDF, but it costs nothing and shows rigour.

---

## 8. Handoff — fill in

| Who | What | Status |
|-----|------|--------|
| Marwan (30 Oct) | Data pack: all sources collected, `Model/Rexel - Inputs & data pack.xlsx` built, source PDFs filed in `Sources/` | Done |
| Marwan (30 Oct) | Excel models built in the same workbook: sheets `Comps`, `DCF`, `LBO`, `Summary` (football-field data). Formulas evaluated and cross-checked with an independent Python replica. | Done |
| Marwan (30 Oct) | Charts in `Charts/` (PNG + PDF): 01 stock performance with events, 02 football field, 03 financial performance (sales, margin, quarterly growth), 04 peer benchmark, 05 DCF sensitivity. Git repository initialised. | Done |
| Marwan (30 Oct) | Repo published: github.com/marwanferreira/rexel-valuation. Deck built: `Rexel valuation - deck.pptx` (15 slides, all quant slides final; slides 3, 4, 5 are framed drafts tagged for Douae) + a draft PDF export. | Done |
| Marwan (31 Oct) | Deck densified to banker standard: logos (Rexel top-right, Dauphine bottom-left, peers and competitors, all on uniform tiles), charts restyled in the deck palette, weighted target price (€45.8, DCF 50% / peers 50%), EV build-up, implied-value arithmetic, MM cross-check, full FCF table, debt schedule, holding P&L, scenarios table. CED and Watsco have no usable public logo (wordmarks). | Done |
| Douae (by Sat 1 Nov evening) | Complete slides 3 (market size + growth with a cited source), 4 (verify private players' figures, Sonepar/Wesco strategy lines), 5 (history, CEO, CMD targets); review slides 2, 14, 15; remove the three orange "Draft" pills. Briefed by message on 31 Oct with the repo link. Delivered 31 Oct (see her review section below). | Done |
| Marwan (31 Oct) | Pulled Douae's deck, restored the file name, fixed layout on slides 3 and 4 (caption overlapping the bar chart, mixed fonts, source lines); PDF export refreshed. Models, charts and numbers unchanged. | Done |
| Marwan (1 Nov) | Pre-send verification: all 15 slides re-rendered and read; key figures cross-checked across the deck, README and checklist (DCF €43.1, LBO €37.6, target €43–46, price €34.47, 308.5m shares); every workbook formula evaluated outside Excel with no error (DCF L11 = 43.13, LBO C113 = 37.58); document properties and speaker notes cleaned. | Done |
| Both (1 Nov) | Final PDF exported from PowerPoint and emailed to Prof. Heller on 1 November 2026 with the Excel model and the repository link. | Done |
| Marwan (31 Oct) | Final pass: DCF bridge on slide 11 aligned to the model to the euro (€43.1), sensitivity grid centred on the base WACC (8.9%), LBO footnote states the Euribor assumption and its sensitivity. Peer forward EBITDA/EBIT stay proxies (flagged on slide 8). | Done |
| Douae | Market overview, competitive landscape, company overview, financial-performance slide, deck design | Done |
| Both | Interpretation, football field reading, investment opinion | Done (final read-through Sun 2 Nov) |

Notes for Douae (written 30 Oct 2026 by Marwan):
- **Valuation date is frozen at 30 October 2026**, share price €34.47 (Yahoo close, check it on Euronext). Use it everywhere.
- **Everything is in `Model/Rexel - Inputs & data pack.xlsx`.** README sheet = the list of sources S1 to S15 with links, and the conventions (IFRS 16 leases, net debt definition, share count). Cite "S-numbers" on slides, then expand them on the sources slide.
- **Source PDFs are in `Sources/`**: Q3 2026 release + presentation (22 Oct), broker consensus (15 Oct), GCG acquisition release (25 Sep), FY 2025 release, 2025 financial statements.
- **Big news to build the story around:** Q3 same-day growth +11.4% (best since 2022), US +18.2% with data centres at 13% of US sales, guidance raised twice in 2026 (now c.7% growth, ≥6.4% adj. EBITA margin), GCG acquisition ($1.4bn EV, largest in 15 years) financed with a €500m capital increase at €34.75 on 29 Sep. The stock trades below the issue price. That tension (strong operations vs. share price) is the investment question.
- **For the market slide:** Rexel's own description in the Q3 presentation (electrification, data centres, HVAC/solar/EV >20% of European sales), Sonepar €33.6bn (#1, private), Wesco c.$25bn LTM, Rexel €19.4bn, Graybar c.$11bn, CED c.$8bn (private). US top-10 distributors hold c.56% of industry sales (Electrical Wholesaling 2026 pyramid, S14). No reliable global market-size figure was found; use Rexel's addressable-market statements from its Capital Markets Day / URD chapter 1 and cite them.
- **For the company slide:** segment sales and margins FY 2025 and H1 2026 are in the Rexel_Hist sheet (North America 51% of Q3 sales, Europe 43%, APAC 6%). Shareholders: public 74%, Cevian 23%.
- **Do not change blue cells** in the Inputs workbook without telling Marwan; the models link to them.
- **The deck** (`Rexel valuation - deck.pptx`) opens in PowerPoint or Keynote. Your three slides carry an orange DRAFT tag and speaker notes that say exactly what to add. Everything else is final unless we change the opinion. To export the PDF: PowerPoint → File → Export → PDF (or Keynote → Export To → PDF). The deck is regenerated by `scripts/07_build_deck.js` (node, needs `npm install pptxgenjs`); if you edit in PowerPoint directly, tell Marwan so he does not overwrite it.

**Model results at 30 Oct 2026 (share price €34.47, 308.5m shares, market cap €10.6bn, net debt pro forma €2.8bn):**

| Method | € per share | vs price | Where |
|---|---|---|---|
| Peers xEBITDA (median / average of Wesco, Grainger, Ferguson, Watsco, Bunzl) | 46 – 50 | +34% to +46% | `Comps` |
| Peers xEBIT | 52 – 56 | | `Comps` |
| Peers P/E | 43 – 53 | | `Comps` |
| Peers xSales | 77 – 95 | not meaningful: Rexel's 6% margin vs 8-20% for peers | `Comps` |
| DCF base (WACC 8.9%, g 2%, consensus plan then soft landing) | 43.1 | +25% | `DCF` cell L11 |
| DCF sensitivity (K 8.5-9.5%, g 1.5-2.5%) | 37 – 50 | | `DCF` K18:Q23 |
| LBO bid (fund IRR 20%, 3 years, 4.0x acquisition debt, exit at entry 10.3x EBIT) | 37.6 | +9% premium | `LBO` cell C113 |

Reading for the deck: every method lands above the price. The market prices Rexel at 9.8x EBITDAaL 2026e against 13-14x for US distributors, i.e. a "European discount" plus the overhang of the GCG deal and the capital increase. An LBO fund could pay about a 9% premium and still earn 20%, which is thin; a strategic buyer (Sonepar, Wesco) could pay more. Investment opinion to argue: Buy / acquire a stake, target around €43-46, with risks: US cycle and data-centre concentration, copper, GCG integration, French sovereign spread.

Model conventions (say them on the assumptions slide): flows after leases (EBITDAaL, adjusted EBITA) so lease liabilities are not in net debt; rf = 50% Bund + 50% US Treasury = 4.13% (Rexel is 51% North America); MRP 5.0% (Damodaran France); beta 1.2 (own regression 1.25, Yahoo 0.98); spread 1.5% (BB+); tax 27%; WACC 8.9% vs Rexel's own impairment-test WACCs of 8.4-9.8%; peer forward EBITDA/EBIT are proxies (next-year consensus revenue × LTM margin), flagged blue in `Comps`.
## Douae — Presentation review (31 October 2026)

- [x] Slide 3: Added market size, growth outlook and end-market trends.
- [x] Slide 3: Updated Graybar revenue to $12.9bn.
- [x] Slide 4: Reviewed competitor figures and Rexel's positioning in France and the US.
- [x] Slide 5: Added corporate history and CEO's strategic repositioning; verified Axelerate 2028 targets.
- [x] Slides 2, 14 and 15: Clarified LBO maximum bid versus downside floor.
- [x] Removed unsupported 3:1 upside/downside claim.
- [x] Removed Draft labels from assigned slides.
- [x] Verified Solar A/S FY2025 revenue: DKK 12,171m (official annual report).
- [x] Verified Ahlsell Group revenue scale: approximately SEK 50–51bn (company and Moody's sources).
- [x] Reviewed CED revenue: c.$8bn remains unverified and explicitly labelled as an estimate.
- [x] Reviewed Rexel's €320bn addressable market: no official aggregate CAGR identified; qualitative growth profiles retained.
