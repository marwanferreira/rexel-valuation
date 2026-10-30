import json, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
OUT="../Model/Rexel - Inputs & data pack.xlsx"
BLUE=Font(name="Calibri",size=10,color="0000FF"); BLK=Font(name="Calibri",size=10); BOLD=Font(name="Calibri",size=10,bold=True)
HDR=Font(name="Calibri",size=10,bold=True,color="FFFFFF"); NAVY=PatternFill("solid",fgColor="1F3864"); GREY=PatternFill("solid",fgColor="D9D9D9")
TITLE=Font(name="Calibri",size=14,bold=True,color="1F3864"); NOTE=Font(name="Calibri",size=9,italic=True,color="595959")
wb=Workbook()
def sheet(name,title,sub=None):
    ws=wb.create_sheet(name); ws["A1"]=title; ws["A1"].font=TITLE
    if sub: ws["A2"]=sub; ws["A2"].font=NOTE
    ws.sheet_view.showGridLines=False; return ws
def hdr(ws,r,vals,c0=1):
    for i,v in enumerate(vals):
        c=ws.cell(r,c0+i,v); c.font=HDR; c.fill=NAVY; c.alignment=Alignment(horizontal="center" if i else "left")
def row(ws,r,label,vals,c0=1,inp=True,fmt='#,##0.0;(#,##0.0);"-"',bold=False,src=None):
    c=ws.cell(r,c0,label); c.font=BOLD if bold else BLK
    for i,v in enumerate(vals):
        c=ws.cell(r,c0+1+i,v); c.font=BLUE if (inp and v is not None and not isinstance(v,str)) else (BOLD if bold else BLK); c.number_format=fmt
    if src: c=ws.cell(r,c0+1+len(vals),src); c.font=NOTE
def widths(ws,w): 
    for i,x in enumerate(w): ws.column_dimensions[get_column_letter(i+1)].width=x
wb.remove(wb.active)
# ---------------- README
ws=sheet("README","Rexel valuation project – data pack","Valuation date: 30 October 2026 (close). All € millions unless stated. Blue = hard-coded input from a source; black = formula. Sources listed below; every number in the deck must trace to one of them.")
r=4; hdr(ws,r,["#","Source","Date","What it gives","Link / file"]); r+=1
srcs=[
("S1","Rexel – Q3 2026 sales press release & presentation","22 Oct 2026","Q3/9m sales by geography, growth split, backlog, guidance, GCG status","Project/Sources/ ; https://www.rexel.com/en/q3-2026-sales"),
("S2","Rexel – Broker consensus collected by Rexel (13 brokers, 5-13 Oct 2026)","15 Oct 2026","2026e-2028e sales, EBITDA, adj. EBITA, net income, EPS, DPS, FCF, net debt; excl. and incl. GCG","Project/Sources/ ; https://www.rexel.com/app/uploads/2026/10/Estimates-pre-Q3-2026.pdf"),
("S3","Rexel – Q2 sales & H1 2026 results press release","27 Jul 2026","H1 P&L, cash flow, net debt 30 Jun 2026 (3,321.9), leverage 2.4x, segment margins","https://www.rexel.com/en (Financial results) ; mirror https://www.mfn.se/one/a/rexel/rexel-q2-sales-and-h1-2026-results-626f457e"),
("S4","Rexel – Q4 sales & FY 2025 results press release","11 Feb 2026","FY 2024-2025 full P&L, bridge to adj. EBITA, FCF, net debt, dividend, buybacks, 2026 guidance","Project/Sources/ ; https://www.rexel.com/app/uploads/2026/01/PR-Q4-sales-FY-2025-results.pdf"),
("S5","Rexel – 2025 Group results & consolidated financial statements (URD section 5)","16 Feb 2026","Balance sheet, cash flow, segment note, EPS note, treasury shares, impairment-test WACC/g by country","Project/Sources/ ; https://www.rexel.com/app/uploads/2026/02/Rexel-Financial-statements-2025-12-EN.pdf"),
("S6","Rexel – GCG acquisition press release","25 Sep 2026","EV $1.4bn, <8x 2026e EBITDAaL post synergies, >$1.1bn sales, c.11% EBITA margin, €800m debt + €500m equity","Project/Sources/ ; https://www.rexel.com/en/rexel-to-acquire-gcg"),
("S7","Rexel – Capital increase success press release","29 Sep 2026","14,388,490 new shares at €34.75; 296,997,254 shares before; settlement 1 Oct 2026","https://finance.yahoo.com/markets/stocks/articles/rexel-announces-success-c-500-053700628.html"),
("S8","Rexel – Debt & ratings page","30 Oct 2026","Moody's Ba1 / S&P BB+ stable; bonds 2.125% 2028, 5.25% 2030, 4.00% 2030; OCEANE 1% 2031; gross debt 4,300.5, cash 1,033.8 at 30 Jun 2026","https://www.rexel.com/en/investors-analysts/debt"),
("S9","Rexel – Shareholding structure","31 Dec 2025","Public 74.10%, Cevian 23.33%, employees 1.40%, treasury 0.98%","https://www.rexel.com/en/investors-analysts/rexel-share"),
("S10","Yahoo Finance via yfinance (Python)","30 Oct 2026","Daily closes 2021-2026 for Rexel, peers, CAC 40, STOXX 600; peer market caps, LTM financials, analyst estimates, betas","Project/Model/prices_daily_2021-2026.csv ; https://finance.yahoo.com/quote/RXL.PA"),
("S11","Trading Economics – government bond yields","30 Oct 2026","France 10Y OAT 4.97%, Germany 10Y 3.325%, US 10Y 4.936%","https://tradingeconomics.com/france/government-bond-yield"),
("S12","Damodaran – Country risk premiums","Jan 2026","Mature-market ERP 4.23%; France total ERP 5.01% (CRP 0.78%); US 4.46%","https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html"),
("S13","Sonepar – 2025 results","2026","Sonepar sales €33.6bn 2025 (private, #1 worldwide) – for competitive map only","https://www.sonepar.com/"),
("S14","Electrical Wholesaling – 2026 Electrical Pyramid","2026","US concentration: top 10 distributors c.56% of industry sales, top 5 chains ≥46%","https://www.ewweb.com/data-training/electrical-pyramid/article/55394430/ews-2026-electrical-pyramid"),
("S15","Course material – Part 4 (LVMH DCF, Moncler peers) and Part 5 (Carrefour LBO)","2026","Method and workbook templates","../Part 4/Excel ; ../Part 5/Excel"),
]
for s in srcs:
    for i,v in enumerate(s): ws.cell(r,1+i,v).font=BLK
    r+=1
r+=1; ws.cell(r,1,"Conventions to respect").font=BOLD; r+=1
for t in ["Rexel reports EBITDA post-IFRS 16 (leases inside). Either use EV = market cap + net financial debt + lease liabilities with EBITDA, or EV = market cap + net financial debt with EBITDAaL (EBITDA after leases). Never mix.",
          "Net financial debt excludes lease liabilities (Rexel definition). Leverage covenant metric = net financial debt / EBITDAaL.",
          "Adjusted EBITA = current EBITA excluding non-recurring copper effect; this is Rexel's margin KPI and the consensus line.",
          "Share count: 311,385,744 issued after the 1 Oct 2026 capital increase; less c.2.9m treasury shares = c.308.5m outstanding (matches Yahoo 308.4m).",
          "Consensus 'incl. GCG as of 01/27' columns assume GCG consolidated from Jan 2027 (4 brokers only). Base case: use 'excl. GCG' (9 brokers) and treat GCG as a separate add-on, or use incl. GCG with pro-forma net debt. State which.",
          "2022-2023 history comes from Yahoo Finance and is flagged; 2024-2025 is from Rexel's own releases."]:
    ws.cell(r,1,"• "+t).font=BLK; r+=1
widths(ws,[6,62,14,70,80])
# ---------------- Rexel history
ws=sheet("Rexel_Hist","Rexel – historical financials (€m)","FY 2024-2025 from S4/S5 (reported basis). FY 2022-2023 from S10 (Yahoo) – replace with URD figures if time. H1 from S3. 9m/Q3 from S1.")
r=4; hdr(ws,r,["Income statement","FY 2022*","FY 2023*","FY 2024","FY 2025","H1 2025","H1 2026","9m 2026","Source"]); r+=1
data=[("Sales",[18701.6,19153.4,19285.1,19414.6,9775.3,9988.9,15441.8]),
("Gross profit",[None,None,4795.8,4873.2,2447.1,2536.2,None]),
("Gross margin %",[None,None,0.249,0.251,0.250,0.254,None]),
("Operating expenses excl. depreciation",[None,None,-3280.2,-3320.7,-1687.8,-1701.5,None]),
("EBITDA (reported, post-IFRS16)",[1681.2,1581.8,1515.6,1552.5,759.4,834.7,None]),
("EBITDA margin %",[None,None,0.079,0.080,0.078,0.084,None]),
("Depreciation",[None,None,-376.3,-389.2,-195.1,-199.1,None]),
("Current EBITA",[None,None,1139.3,1163.3,564.2,635.6,None]),
("Current adjusted EBITA (constant basis)",[None,None,1137.2,1157.8,559.8,614.8,None]),
("Adj. EBITA margin %",[None,None,0.059,0.060,0.059,0.062,None]),
("PPA amortization",[None,None,-35.7,-45.6,-22.0,-23.2,None]),
("Other income and expenses",[None,None,-257.7,-56.2,-36.5,-7.7,None]),
("Operating income",[1382.9,1295.9,845.9,1061.6,505.7,604.7,None]),
("Net financial expenses",[None,None,-207.7,-214.2,-106.6,-112.5,None]),
("Income tax",[None,None,-297.2,-256.0,-137.5,-150.1,None]),
("Effective tax rate %",[None,None,0.466,0.302,0.345,0.305,None]),
("Net income (group share)",[922.3,774.7,341.0,591.4,261.6,342.1,None]),
("Recurring net income",[None,None,662.6,678.5,307.9,346.6,None]),
("Diluted EPS (€)",[3.04,2.57,1.13,1.99,None,None,None]),
("Dividend per share (€)",[None,None,1.20,1.20,None,None,None]),
("Weighted avg shares (m)",[302.9,300.9,300.0,295.1,None,None,None])]
for lab,v in data:
    f='0.0%' if '%' in lab else ('0.00' if 'EPS' in lab or 'per share' in lab else '#,##0.0;(#,##0.0);"-"')
    row(ws,r,lab,v,fmt=f,src="S4/S5; *S10"); r+=1
r+=1; hdr(ws,r,["Cash flow","","","FY 2024","FY 2025","H1 2025","H1 2026","","Source"]); r+=1
for lab,v in [("EBITDAaL (EBITDA after lease payments)",[None,None,1381.0,1226.8,595.5,669.7,None]),
("Lease payments (in EBITDAaL)",[None,None,None,-325.7,None,None,None]),
("Operating cash flow before interest & tax",[None,None,1253.2,1442.6,None,None,None]),
("Repayment of lease liabilities",[None,None,-245.1,-253.4,None,None,None]),
("Change in working capital",[None,None,34.3,-239.1,-379.7,-342.1,None]),
("Operating capex (net of disposals)",[None,None,-125.8,-136.2,-71.1,-65.5,None]),
("Gross capex",[None,None,None,None,-68.5,-68.4,None]),
("FCF before interest & tax (reported)",[None,None,916.5,813.8,127.2,247.3,None]),
("FCF before interest & tax excl. €124m fine",[None,None,916.5,937.8,None,None,None]),
("FCF conversion % (FCF bit / EBITDAaL)",[None,None,None,0.663,0.21,0.37,None]),
("Interest paid",[None,None,-129.6,-137.1,None,None,None]),
("Tax paid",[None,None,-281.0,-270.6,None,None,None]),
("FCF after interest & tax",[None,None,506.0,406.1,-73.8,72.8,None]),
("Net financial investment (M&A)",[None,None,-437.2,-256.7,-198.8,-398.3,None]),
("Dividends paid",[None,None,None,None,-354.6,-353.2,None]),
("Share buyback",[None,None,-99.5,-100.0,None,None,None])]:
    f='0.0%' if '%' in lab else '#,##0.0;(#,##0.0);"-"'; row(ws,r,lab,v,fmt=f,src="S4/S5/S3"); r+=1
# EBITDAaL 2024 = 916.5/ FCF conv? not given -> leave: compute as note
ws.cell(r-16,4).value=None
r+=1; hdr(ws,r,["Balance sheet (31 Dec)","","","2024","2025","30 Jun 2026","","","Source"]); r+=1
for lab,v in [("Goodwill",[None,None,3978.4,3907.6,None]),("Intangible assets",[None,None,1675.1,1659.3,None]),("PP&E",[None,None,368.3,353.7,None]),("Right-of-use assets",[None,None,1381.4,1342.2,None]),
("Inventories",[None,None,2485.7,2364.1,None]),("Trade receivables",[None,None,2694.4,2682.0,None]),("Other receivables",[None,None,795.5,893.6,None]),("Cash",[None,None,883.3,1037.5,1033.8]),("Total assets",[None,None,14408.2,14366.6,None]),
("Total equity (incl. minorities)",[None,None,5569.9,5405.7,None]),("Interest-bearing debt (non-current + current + accrued)",[None,None,3220.8,3625.3,4300.5]),("Lease liabilities (non-current + current)",[None,None,1521.3,1488.3,None]),
("Trade payables",[None,None,2305.4,2170.9,None]),("Other current liabilities",[None,None,1227.1,1039.5,None]),
("Net financial debt (Rexel definition, excl. leases)",[None,None,2483.9,2631.4,3321.9]),("Leverage net debt / EBITDAaL (x)",[None,None,None,2.03,2.39]),
("Trade working capital (inv + receivables - payables)",[None,None,2874.7,2875.2,None]),("Trade WC % of sales",[None,None,0.149,0.148,None]),
("Treasury shares (number)",[None,None,2847481,2890248,None]),("Issued shares (number)",[None,None,None,296997254,None])]:
    f='0.0%' if '%' in lab else ('0.00' if '(x)' in lab else ('#,##0' if 'number' in lab else '#,##0.0;(#,##0.0);"-"')); row(ws,r,lab,v,fmt=f,src="S5/S3/S7/S8"); r+=1
r+=1; hdr(ws,r,["Sales by segment","","","FY 2024","FY 2025","H1 2026","Q3 2026","9m 2026","Source"]); r+=1
for lab,v in [("North America",[None,None,8461.8,8906.8,4559.6,2756.8,7316.4]),("  o/w United States",[None,None,6975.0,7381.9,None,2317.4,6078.2]),("  o/w Canada",[None,None,1486.8,1524.9,None,439.4,1238.2]),
("Europe",[None,None,9550.6,9403.2,4802.9,2348.4,7151.3]),("  o/w France",[None,None,3654.3,3756.2,None,934.2,2902.3]),("  o/w DACH",[None,None,2258.0,2191.8,None,620.9,1765.2]),("  o/w Benelux",[None,None,1561.2,1553.5,None,383.6,1210.1]),("  o/w UK & Ireland",[None,None,1038.6,945.3,None,237.1,698.5]),("  o/w Nordics / Sweden",[None,None,830.0,732.2,None,127.7,427.9]),
("Asia-Pacific",[None,None,1272.7,1104.7,626.5,347.6,974.2]),("Group",[None,None,19285.1,19414.6,9988.9,5452.9,15441.8])]:
    row(ws,r,lab,v,src="S5/S1/S3",bold=(lab=="Group")); r+=1
r+=1; hdr(ws,r,["Adj. EBITA by segment","","","FY 2025","margin","H1 2026","margin","","Source"]); r+=1
for lab,v in [("North America",[None,None,650.8,0.073,323.9,0.071]),("Europe",[None,None,535.2,0.057,301.7,0.063]),("Asia-Pacific",[None,None,12.6,0.011,7.7,0.012]),("Corporate",[None,None,-40.8,None,-18.6,None]),("Group",[None,None,1157.8,0.060,614.8,0.062])]:
    for i,x in enumerate(v):
        pass
    row(ws,r,lab,v,src="S4/S3",bold=(lab=="Group"))
    for col in (5,7): ws.cell(r,col).number_format='0.0%'
    r+=1
r+=1; hdr(ws,r,["Same-day growth by quarter","Q1 25","Q2 25","Q3 25","Q4 25","Q1 26","Q2 26","Q3 26","Source"]); r+=1
for lab,v in [("Group",[0.018,0.030,0.030,0.038,0.034,0.067,0.114]),("North America",[0.058,0.074,0.079,0.079,0.087,0.078,0.174]),("Europe",[-0.030,-0.005,-0.005,0.0,0.006,0.044,0.045]),("Asia-Pacific",[-0.065,0.014,0.014,0.045,0.038,0.158,0.158])]:
    row(ws,r,lab,v,fmt='0.0%',src="S1 presentation p.7 (read from chart – verify Q1-Q2 25 and APAC)"); r+=1
r+=1; ws.cell(r,1,"Other KPIs (S1/S4/S6): 1,876 branches, 17 countries, 26,306 employees; digital = 36% of sales Q3 26; cables c.16% of sales, copper c.60% of cable; NA backlog > $3bn (+42% vs end-2025), data centers c.13% of US sales and c.25% of US backlog; HVAC/solar/EV > 20% of Europe sales.").font=NOTE
widths(ws,[48,11,11,11,11,11,11,11,40])
# ---------------- Consensus
ws=sheet("Consensus","Rexel broker consensus (collected by Rexel, 13 of 15 brokers, 5-13 Oct 2026) – S2","Barclays, Bernstein, BofA, CIC, Citi, Exane, Goldman Sachs, Jefferies, JP Morgan, Kepler, Morgan Stanley, Oddo, Oxcap. Pre-Q3 → does not include the 22 Oct guidance raise to c.7% growth / ≥6.4% margin (upside to 2026e).")
for blk,r0 in (("MEAN",4),("MEDIAN",28)):
    hdr(ws,r0,[blk+" (€m)","2025A","2026e","2027e excl GCG","2028e excl GCG","2027e incl GCG","2028e incl GCG"]); r=r0+1
    M={"MEAN":[("Reported sales",[19415,20578,21938,22949,22855,23906]),("Organic growth same-day %",[0.025,0.057,0.047,0.043,0.046,0.042]),("Gross margin",[4868,5192,5575,5868,5717,6046]),("Gross margin %",[0.251,0.252,0.254,0.256,0.250,0.253]),("Reported EBITDA",[1553,1712,1866,2013,1994,2131]),("Reported EBITA",[1163,1321,1467,1599,1560,1680]),("Adj. EBITA",[1158,1305,1466,1599,1559,1678]),("Adj. EBITA margin %",[0.060,0.0634,0.0668,0.0697,0.0682,0.0702]),("Net financial expenses",[-214,-232,-224,-218,-267,-259]),("Income tax",[-256,-290,-312,-349,-324,-355]),("Tax rate %",[0.302,0.289,0.270,0.270,0.271,0.269]),("Reported net income",[591,713,843,945,871,967]),("EPS adjusted (€)",[None,2.45,2.90,3.26,2.97,3.27]),("DPS (€)",[1.20,1.22,1.31,1.39,1.29,1.38]),("FCF before interest & tax",[938,938,1154,1285,1056,1272]),("FCF conversion %",[0.76,0.68,0.75,0.77,0.65,0.72]),("FCF after interest & tax",[530,513,701,830,547,734]),("Net debt excl. leases (2026e: excl GCG 2806 / incl GCG 3714)",[2631,2806,2580,2261,3501,3130]),("Lease liabilities (2026e: 1473 / 1557)",[1488,1473,1385,1295,1565,1572])],
       "MEDIAN":[("Reported sales",[19415,20558,21645,22531,22726,23765]),("Organic growth same-day %",[0.025,0.057,0.049,0.043,0.048,0.043]),("Gross margin",[4868,5207,5579,5876,5767,6158]),("Gross margin %",[0.251,0.253,0.258,0.261,0.254,0.259]),("Reported EBITDA",[1553,1711,1836,1978,1992,2122]),("Reported EBITA",[1163,1319,1470,1610,1555,1678]),("Adj. EBITA",[1158,1297,1470,1610,1555,1675]),("Adj. EBITA margin %",[0.060,0.0631,0.0679,0.0714,0.0684,0.0705]),("Net financial expenses",[-214,-230,-221,-214,-265,-258]),("Income tax",[-256,-296,-305,-337,-323,-356]),("Tax rate %",[0.302,0.300,0.270,0.270,0.270,0.270]),("Reported net income",[591,709,827,929,873,968]),("EPS adjusted (€)",[None,2.42,2.91,3.30,2.98,3.29]),("DPS (€)",[1.20,1.20,1.30,1.38,1.29,1.39]),("FCF before interest & tax",[938,982,1159,1255,1012,1295]),("FCF conversion %",[0.76,0.69,0.77,0.77,0.63,0.73]),("FCF after interest & tax",[530,550,682,816,506,735]),("Net debt excl. leases (2026e: excl GCG 2703 / incl GCG 3666)",[2631,2703,2460,2172,3558,3185]),("Lease liabilities (2026e: 1488 / 1565)",[1488,1488,1488,1488,1580,1595])]}[blk]
    for lab,v in M:
        f='0.0%' if '%' in lab else ('0.00' if '(€)' in lab else '#,##0;(#,##0);"-"'); row(ws,r,lab,v,fmt=f); r+=1
    ws.cell(r,1,"Q3 2026 consensus sales was 5,205 (mean) / 5,226 (median) vs 5,452.9 actual → beat of c.+4.8%. Same-day growth consensus 6.6-6.8% vs 11.4% actual.").font=NOTE
widths(ws,[52,12,12,16,16,16,16])
# ---------------- Market data
ws=sheet("Market","Rexel market data at valuation date","Close of 30 Oct 2026 (Yahoo Finance, S10) – cross-check on Euronext before the deck. 52-week range 28.83 – 39.80.")
r=4; hdr(ws,r,["Item","Value","Unit","Source / note"]); r+=1
for lab,v,u,s in [("Valuation date","30 Oct 2026","",""),("Share price (close)",34.47,"€","S10 – verify on live.euronext.com"),("Previous close (29 Oct)",35.03,"€","S10"),("52-week high",39.80,"€","S10"),("52-week low",28.83,"€","S10"),
("Issued shares before capital increase (28 Sep 2026)",296997254,"shares","S7"),("New shares issued (settled 1 Oct 2026)",14388490,"shares","S7"),("Issue price",34.75,"€","S7 – c.1% discount"),("Issued shares after capital increase","=B10+B11","shares","formula"),("Treasury shares (31 Dec 2025, latest disclosed)",2890248,"shares","S5"),("Shares outstanding","=B13-B14","shares","formula (Yahoo shows 308,448,880)"),
("Market capitalisation (outstanding)","=B6*B15/1000000","€m","formula"),("Market capitalisation (issued)","=B6*B13/1000000","€m","formula"),
("Net financial debt 30 Jun 2026 (excl. leases)",3321.9,"€m","S3/S8"),("Capital increase proceeds (gross)",-500,"€m","S7 – received 1 Oct 2026, after the June balance sheet"),("Net financial debt pro forma for the capital increase","=B18+B19","€m","formula – before GCG payment"),("Lease liabilities (31 Dec 2025)",1488.3,"€m","S5"),("Non-controlling interests (31 Dec 2025)",11.5,"€m","S5"),("Net employee benefit liabilities (31 Dec 2025)",100.1,"€m","S5 – optional debt-like item"),
("EV (pro forma, excl. leases) = mcap + net debt + minorities","=B16+B20+B22","€m","use with EBITDAaL / adj. EBITA"),("EV (pro forma, incl. leases)","=B24+B21","€m","use with reported EBITDA"),
("GCG enterprise value",1400,"$m","S6 – closing expected by end-2026; add c.€1.2bn net debt + c.$150m EBITDAaL (<8x post-synergies) if modelling pro forma"),("GCG expected 2026 revenue",">1,100","$m","S6"),("GCG expected 2026 EBITA margin",0.11,"%","S6"),("GCG debt financing",800,"€m","S6"),
("Analyst target price – mean (Yahoo, 13 analysts)",41.86,"€","S10 – high 46 / low 30 / median 44"),("Dividend 2025 (paid May 2026)",1.20,"€","S4 – 52% payout; policy ≥40%"),("Shareholders: public / Cevian / employees / treasury","74.1% / 23.3% / 1.4% / 1.0%","","S9"),
("Credit rating","Ba1 (Moody's) / BB+ (S&P), stable","","S8"),("Index membership","SBF 120, CAC Next 20, STOXX 600, MSCI World","","S6")]:
    c=ws.cell(r,1,lab); c.font=BLK; c=ws.cell(r,2,v); c.font=BLUE if not (isinstance(v,str) and v.startswith("=")) else BLK
    c.number_format='#,##0' if (isinstance(v,(int,float)) and abs(v)>1000) or (isinstance(v,str) and v.startswith("=") and "B6" not in v) else ('0.0%' if u=="%" else '#,##0.00')
    if isinstance(v,str) and v.startswith("=") and ("B16" in v or "B18" in v or "B24" in v or "B6*" in v): c.number_format='#,##0.0'
    ws.cell(r,3,u).font=BLK; ws.cell(r,4,s).font=NOTE; r+=1
widths(ws,[58,22,8,90])
# ---------------- Peers
info=json.load(open("info.json")); fin=json.load(open("financials.json"))
ws=sheet("Peers","Peer data (Yahoo Finance, 30 Oct 2026 – S10)","Local currency, millions. LTM = last four reported quarters. '0y'/'+1y' = analyst consensus for current and next fiscal year (revenue; EPS). EBITDA/EBIT forward estimates are NOT on Yahoo: take them from MarketScreener/Zonebourse or company guidance and paste in blue. Choose the sample and justify each name (distributors first; Legrand/Schneider are manufacturers → reference only).")
r=4; hdr(ws,r,["Company","Ticker","Ccy","Price","Shares (m)","Market cap (m)","Net debt (m, last BS)","EV (m)","LTM sales","LTM EBITDA","LTM EBIT","LTM net income","Rev est FY0","Rev est FY+1","EPS est FY0","EPS est FY+1","Beta (Yahoo 5y mthly)","Fwd P/E","EV/EBITDA (Yahoo)","Target mean","Role in sample"]); r+=1
roles={"Rexel":"Target","Wesco":"Core peer – US electrical/datacom distributor, closest comp","Grainger":"Core peer – US MRO distributor","Ferguson":"Core peer – plumbing/HVAC distributor, NA","Watsco":"Core peer – HVAC distributor, US","Bunzl":"Core peer – European B2B distributor (Price in GBp → divide by 100)","Fastenal":"Optional – high-multiple industrial distributor","Legrand":"Reference – manufacturer, same end-markets","Schneider":"Reference – manufacturer"}
for k in ["Rexel","Wesco","Grainger","Ferguson","Watsco","Bunzl","Fastenal","Legrand","Schneider"]:
    i=info[k]; f=fin[k]; m=1e6
    px=i["currentPrice"]; sh=i["sharesOutstanding"]/m; mc=i["marketCap"]/m; nd=f.get("BS_Net Debt"); nd=nd/m if nd else None
    if k=="Bunzl": px=px/100; 
    ltm=lambda n: (f.get("LTM_"+n)/m if f.get("LTM_"+n) else None)
    est=f.get("est_rev",{}); eps=f.get("est_eps",{})
    vals=[k,{"Rexel":"RXL.PA","Wesco":"WCC","Grainger":"GWW","Ferguson":"FERG","Watsco":"WSO","Bunzl":"BNZL.L","Fastenal":"FAST","Legrand":"LR.PA","Schneider":"SU.PA"}[k],"GBP" if k=="Bunzl" else i["currency"],px,sh,mc,nd,f"=F{r}+G{r}",ltm("Total Revenue"),ltm("EBITDA"),ltm("EBIT"),ltm("Net Income"),(est.get("0y") or 0)/m or None,(est.get("+1y") or 0)/m or None,eps.get("0y"),eps.get("+1y"),i["beta"],i["forwardPE"],i["enterpriseToEbitda"],(f.get("target") or {}).get("mean"),roles[k]]
    for j,v in enumerate(vals):
        c=ws.cell(r,1+j,v); c.font=BLUE if isinstance(v,(int,float)) else BLK
        c.number_format='#,##0.0' if isinstance(v,(int,float)) and abs(v)>100 else '0.00'
    r+=1
ws.cell(r+1,1,"Notes: Rexel 'BS net debt' here is Yahoo's 31 Dec 2025 figure (2,568) – use the Market sheet (3,321.9 at 30 Jun 2026, pro forma 2,821.9) in the model. Bunzl and Schneider LTM lines are empty on Yahoo (report in GBP/€ annual only) → fill from their H1 2026 reports. Sonepar (€33.6bn sales 2025, S13) and Graybar/CED are private: competitive map only.").font=NOTE
widths(ws,[12,9,5,9,11,14,16,12,11,11,11,13,12,12,10,10,10,8,10,10,60])
# ---------------- WACC
ws=sheet("WACC_inputs","Discount rate inputs","Pick ONE consistent set and justify it on the DCF assumptions slide. The prof's LVMH case used rf 3.5%, MRP 8%, spread 0.56% – the course convention is CAPM + (rf + spread) weighted at market values.")
r=4; hdr(ws,r,["Input","Value","Source / comment"]); r+=1
for lab,v,s in [("France 10Y OAT (30 Oct 2026)",0.0497,"S11 – at a 24-year high after the Oct 2026 budget turmoil (5.05% peak on 21 Oct). Argue: Rexel is 51% North America, so a blended rf is defensible."),("Germany 10Y Bund (30 Oct 2026)",0.03325,"S11 – euro risk-free floor"),("US 10Y Treasury (30 Oct 2026)",0.04936,"S11"),("Suggested blended rf (e.g. 50% OAT + 50% UST, or Bund + spread)","=0.5*B5+0.5*B7","formula – change at will"),
("Damodaran mature-market ERP (Jan 2026)",0.0423,"S12"),("Damodaran France total ERP (Jan 2026)",0.0501,"S12 – includes 0.78% country risk premium"),("Damodaran US total ERP (Jan 2026)",0.0446,"S12"),("Course MRP used by the prof (LVMH case)",0.08,"S15 – high; if used, say why (or show as sensitivity)"),
("Beta – Yahoo (5y monthly vs index)",0.98,"S10"),("Beta – own regression, 2y weekly vs CAC 40",1.25,"S10 – computed from prices.csv"),("Beta – own regression, 2y weekly vs STOXX 600",1.38,"S10"),("Beta – own regression, 5y weekly vs CAC 40",1.35,"S10"),("Peer betas (Yahoo): Wesco 1.55, Grainger 1.01, Ferguson 1.12, Watsco 1.04, Bunzl 0.32, Fastenal 0.69, Legrand 0.96, Schneider 1.13","","S10 – unlever with each peer's D/E and re-lever at Rexel's (course method, Part 3)"),
("Rexel bond 4.00% Sept 2030 (issued Sept 2026, Ba1/BB+)",0.040,"S8 – best pre-tax cost-of-debt proxy: new issue at market"),("Rexel bond 5.25% Sept 2030 (issued 2023)",0.0525,"S8"),("Effective interest rate 2025 (per Rexel)",0.0396,"S4 – 4.35% in 2024"),("Net financial expenses 2025 / avg net debt","=214.2/((2483.9+2631.4)/2)","S4/S5 – implied all-in cost incl. leases interest"),
("Normative tax rate",0.27,"S4 – 30% in 2025 incl. French exceptional contribution; consensus uses 27% from 2027"),("Debt / (Debt + Equity) at market, pro forma","=Market!B20/(Market!B20+Market!B16)","formula"),
("Long-term growth g (consensus / Rexel impairment tests)",0.02,"S5 – Rexel uses 1.9% France, 2.2% US, 2.0% most countries; inflation-based"),
("Rexel's own after-tax WACC by country (impairment test, 31 Dec 2025)","US 9.8% / France 8.7% / Canada 8.4% / Germany 8.6% / UK 9.5% / Netherlands 8.4% / Sweden 8.7% / Australia 9.5%","S5 note 12 – a group blended WACC of c.9.0-9.3% is consistent with management; a DCF WACC far outside 8-10% needs justification")]:
    ws.cell(r,1,lab).font=BLK; c=ws.cell(r,2,v); c.font=BLUE if isinstance(v,float) else BLK; c.number_format='0.00%'; ws.cell(r,3,s).font=NOTE; r+=1
widths(ws,[70,24,120])
# ---------------- Prices
px=pd.read_csv("prices.csv",index_col=0,parse_dates=True)
ws=sheet("Prices","Daily closing prices (Yahoo Finance, S10)","Local currency (Bunzl in pence). Use for the stock-performance slide (rebase to 100) and the beta regression.")
r=4; hdr(ws,r,["Date"]+list(px.columns)); r+=1
for d,rowv in px.iterrows():
    ws.cell(r,1,d.date()).number_format='yyyy-mm-dd'
    for j,v in enumerate(rowv): 
        if pd.notna(v): c=ws.cell(r,2+j,float(v)); c.number_format='0.00'; c.font=BLK
    r+=1
widths(ws,[12]+[10]*len(px.columns))
ws.freeze_panes="B5"
wb.save(OUT); print("saved",OUT)
