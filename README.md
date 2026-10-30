# Rexel – Valuation of a listed company

Group project for the *Introduction to Corporate Finance* course (Université Paris Dauphine-PSL, M2 104).
Marwan Ferreira and Douae.

**Question:** should an investor acquire a stake in Rexel (Euronext Paris: RXL)?
**Valuation date:** 30 October 2026 (share price €34.47).

## Contents

| Folder / file | What it is |
|---|---|
| `CHECKLIST - Rexel project.md` | Project rules (from the instructions), course-method requirements, slide plan, data sources, handoff notes and model results |
| `Model/Rexel - Inputs & data pack.xlsx` | The model: `README` (sources S1-S15), `Rexel_Hist`, `Consensus`, `Market`, `Peers`, `WACC_inputs`, `Prices`, then `Comps`, `DCF`, `LBO`, `Summary` |
| `Model/prices_daily_2021-2026.csv` | Daily closes for Rexel, peers, CAC 40 and STOXX 600 |
| `Charts/` | Deck charts (PNG and PDF): stock performance, football field, financial performance, peer benchmark, DCF sensitivity |
| `Sources/` | Rexel public documents used (press releases, consensus, financial statements) |
| `Instructions EN - ...pdf` | The assignment |

## Methods

- **Listed peers** (Wesco, Grainger, Ferguson, Watsco, Bunzl): EV = market cap + net debt; xSales, xEBITDA, xEBIT, P/E; sector median and average applied to Rexel.
- **DCF**: broker consensus 2026-2028 collected by Rexel, soft landing to g = 2% by 2032, FCF = EBIT(1-t) + D&A - capex - ΔWCR, WACC by CAPM, Gordon-Shapiro terminal value, sensitivity to K and g.
- **LBO bid**: holding company financed by senior A/B and junior C debt plus fund equity, tax consolidation, exit after 3 years at the entry multiple, 20% IRR required; the bid is the maximum price per share consistent with that return.

Blue cells in the workbook are sourced inputs; black cells are formulas. The workbook recalculates when opened in Excel.

## Results (€ per share)

| Method | Value |
|---|---|
| Peers xEBITDA (median to average) | 46 – 50 |
| DCF base case (WACC 8.9%, g 2%) | 43.1 |
| DCF sensitivity (K 8.5-9.5%, g 1.5-2.5%) | 37 – 50 |
| LBO bid (20% IRR, 4.0x acquisition debt) | 37.6 |
| Share price, 30 Oct 2026 | 34.47 |
