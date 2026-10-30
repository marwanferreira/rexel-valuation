# Rexel (RXL.PA) — Equity valuation and investment opinion

**Should an investor acquire a stake in Rexel, the world's second-largest electrical distributor?**

Group project for *Introduction to Corporate Finance* — Université Paris Dauphine-PSL, Master 2 (104) Research in Finance.
Marwan Ferreira da Silva & Douae.

| | |
|---|---|
| **Valuation date** | 30 October 2026 |
| **Share price** | €34.47 (Euronext Paris: RXL) |
| **Market cap** | €10.6bn (308.5m shares after the September 2026 capital increase) |
| **Net financial debt** | €2.8bn pro forma the €500m equity raise, excl. leases |
| **Methods** | Listed peers · DCF · LBO bid by a financial sponsor |

---

## Headline

![Football field](Charts/02_football_field.png)

| Method | € per share | vs €34.47 |
|---|---|---|
| Listed peers — xEBITDA (median to average) | **46 – 50** | +34% to +46% |
| Listed peers — xEBIT | 52 – 56 | |
| Listed peers — P/E | 43 – 53 | |
| DCF — base case (WACC 8.9%, g 2.0%) | **43.1** | +25% |
| DCF — sensitivity (K 8.5–9.5%, g 1.5–2.5%) | 37 – 50 | |
| LBO bid — 20% IRR, 3 years, 4.0x acquisition debt | **37.6** | +9% premium |
| Sell-side targets (13 analysts) | 30 – 46 | mean 41.9 |

Every method lands above the market price. Rexel trades at **9.8x 2026e EBITDA after leases** against **13–14x for US distributors** (Wesco, Ferguson, Watsco): a European discount, widened by the overhang of the $1.4bn GCG acquisition and the capital increase priced at €34.75. A financial sponsor could pay only a thin premium and still clear a 20% IRR, so the LBO route is marginal; the upside belongs to a long-term holder or a strategic buyer.

---

## The story in four charts

**Operating momentum is the strongest since 2022** — Q3 2026 same-day sales +11.4%, United States +18.2% with data centres at 13% of US sales, guidance raised twice this year.

![Financial performance](Charts/03_financial_performance.png)

**…but the share has lagged its closest peer** — Rexel +37% over two years vs Wesco +93%, and a sharp drawdown since the GCG announcement and the capital increase.

![Stock performance](Charts/01_stock_performance.png)

**The discount is a multiple discount, not a margin story alone** — Rexel's EBITDA margin matches Wesco's, yet it trades 3 turns cheaper.

![Peer benchmark](Charts/04_peer_benchmark.png)

**The DCF is above the price across the reasonable grid.**

![DCF sensitivity](Charts/05_dcf_sensitivity.png)

---

## Methodology

All three methods follow the course framework and live in one workbook, `Model/Rexel - Inputs & data pack.xlsx`. Blue cells are sourced inputs, black cells are formulas; every input is tagged with a source number (S1–S15, listed in the `README` sheet).

### 1. Listed peers (`Comps`)
- Sample: **Wesco, Grainger, Ferguson, Watsco, Bunzl** — listed B2B distributors. Fastenal, Legrand and Schneider are shown for reference and excluded from the averages (premium MRO multiple; manufacturers).
- EV = market cap + net debt; multiples xSales, xEBITDA, xEBIT, P/E on LTM and next-year figures; sector median and average applied to Rexel's aggregates.
- Consistency: US GAAP peers report EBITDA after operating-lease cost, so Rexel is valued on **EBITDA after leases (EBITDAaL)** and **adjusted EBITA**, with lease liabilities excluded from net debt.
- The sales multiple is reported but not used: Rexel's 6% margin against 10–20% for the peers makes it meaningless.

### 2. DCF (`DCF`)
- Business plan 2026–2028: **broker consensus collected by Rexel itself** (13 brokers, 5–13 October 2026), then a soft landing to g = 2% by 2032 (growth converges linearly, margins frozen at the 2028 level, capex converges to D&A).
- FCF = EBIT (1 − t) + D&A − capex − ΔWCR, with WCR held at 14.2% of sales (2025 financial balance sheet).
- Discount rate: rf 4.13% (50% Bund + 50% US Treasury, Rexel being 51% North American), market risk premium 5.0% (Damodaran, France), beta 1.2 (own 2-year weekly regression), spread 150 bp (BB+/Ba1), tax 27% → **WACC 8.9%**, inside the 8.4–9.8% range Rexel uses in its own impairment tests.
- Terminal value by Gordon–Shapiro; live sensitivity grid to K and g.

### 3. LBO bid (`LBO`)
- A holding company acquires 100% of Rexel at end-2026, financed by Senior A (amortising, 7 years), Senior B (bullet, 8 years), Junior C (bullet, 9 years) and sponsor equity; acquisition debt set at 4.0x 2026e EBITDAaL on top of Rexel's own 2.0x.
- Tax consolidation: acquisition interest is deducted from Rexel's taxable income; 100% of net income is upstreamed to service the debt.
- Exit after three years at the entry EBIT multiple (10.3x). The **bid is the maximum price per share** at which the sponsor earns exactly its 20% IRR: €37.6, a 9% premium, 1.7x money multiple.

---

## Repository

```
├── README.md
├── CHECKLIST - Rexel project.md     Assignment rules, course-method checklist, slide plan, handoff notes, results
├── Model/
│   ├── Rexel - Inputs & data pack.xlsx   Sources · Rexel_Hist · Consensus · Market · Peers · WACC_inputs · Prices · Comps · DCF · LBO · Summary
│   └── prices_daily_2021-2026.csv        Daily closes: Rexel, peers, CAC 40, STOXX 600
├── Charts/                          Deck charts (PNG + vector PDF)
├── Sources/                         Rexel public documents used (press releases, consensus, financial statements)
├── scripts/                         Python: data pull, workbook build, charts, independent model check
└── Instructions EN - ....pdf        The assignment
```

### Reproduce

```bash
pip install -r scripts/requirements.txt
cd scripts
python 01_pull_prices.py             # prices + key stats from Yahoo Finance
python 02_pull_peer_financials.py    # peer financials, estimates, regression betas
python 03_build_inputs_workbook.py   # sourced Inputs workbook
python 04_build_model_sheets.py      # Comps / DCF / LBO / Summary sheets (live Excel formulas)
python 05_charts.py                  # the five charts
python 06_independent_check.py       # re-computes the three models in Python and prints the results
```

The workbook is written without cached values and recalculates when opened in Excel. The formulas were also evaluated with the `formulas` library and reconciled to the Python replica (DCF to the euro).

---

## Sources

Rexel Q3 2026 sales release and presentation (22 Oct 2026) · Rexel broker consensus (15 Oct 2026) · Q2 & H1 2026 results (27 Jul 2026) · Q4 & FY 2025 results (11 Feb 2026) · 2025 consolidated financial statements · GCG acquisition and capital-increase releases (25–29 Sep 2026) · Rexel debt & ratings page · Yahoo Finance (prices, peers, estimates) · Trading Economics (sovereign yields, 30 Oct 2026) · Damodaran (equity risk premiums, Jan 2026) · Electrical Wholesaling 2026 Electrical Pyramid.

*Academic work; not investment advice.*
