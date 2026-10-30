import json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties
OUT="../Model/Rexel - Inputs & data pack.xlsx"
wb=load_workbook(OUT)
for n in ["Comps","DCF","LBO","Summary"]:
    if n in wb.sheetnames: wb.remove(wb[n])
BLUE=Font(name="Calibri",size=10,color="0000FF"); BLK=Font(name="Calibri",size=10); BOLD=Font(name="Calibri",size=10,bold=True)
HDR=Font(name="Calibri",size=10,bold=True,color="FFFFFF"); NAVY=PatternFill("solid",fgColor="1F3864"); RES=PatternFill("solid",fgColor="FFF2CC")
TITLE=Font(name="Calibri",size=14,bold=True,color="1F3864"); NOTE=Font(name="Calibri",size=9,italic=True,color="595959")
N='#,##0;(#,##0);"-"'; N1='#,##0.0;(#,##0.0);"-"'; P='0.0%'; X='0.0"x"'; E2='0.00'
def sheet(name,title,sub):
    ws=wb.create_sheet(name); ws["A1"]=title; ws["A1"].font=TITLE; ws["A2"]=sub; ws["A2"].font=NOTE; ws.sheet_view.showGridLines=False; return ws
def hdr(ws,r,vals,c0=1):
    for i,v in enumerate(vals):
        c=ws.cell(r,c0+i,v); c.font=HDR; c.fill=NAVY; c.alignment=Alignment(horizontal="center" if i else "left",wrap_text=True)
def put(ws,r,c,v,fmt=N,bold=False,res=False):
    cell=ws.cell(r,c,v); isf=isinstance(v,str) and v.startswith("=")
    if isinstance(v,str) and v.startswith("= "): cell.data_type="s" 
    cell.font=BOLD if bold else (BLUE if isinstance(v,(int,float)) and not isinstance(v,bool) else BLK); cell.number_format=fmt
    if res: cell.fill=RES; cell.font=BOLD
    return cell
def lab(ws,r,t,bold=False,note=False):
    c=ws.cell(r,1,t); c.font=NOTE if note else (BOLD if bold else BLK)
    if isinstance(t,str) and t.startswith("="): c.data_type="s"
    return c
def widths(ws,w):
    for i,x in enumerate(w): ws.column_dimensions[L(i+1)].width=x
info=json.load(open("info.json")); fin=json.load(open("financials.json"))
# =============================== COMPS ===============================
ws=sheet("Comps","Valuation by listed peers (course method, Part 4 – Moncler)","EV = market cap + net debt (Yahoo, excl. leases, last balance sheet). LTM = last 4 quarters (US GAAP peers: EBITDA is after lease costs → compare with Rexel EBITDAaL and adj. EBITA, net financial debt excl. leases). FY+1 EBITDA/EBIT = analyst FY+1 revenue × LTM margin (PROXY – replace with MarketScreener consensus if possible, keep blue). FY+1 NI = consensus EPS × shares. Local currency, millions.")
cols=["Company","Ccy","Price","Shares (m)","Market cap","Net debt","EV","Sales LTM","Sales FY+1e","EBITDA LTM","EBITDA FY+1e","EBIT LTM","EBIT FY+1e","Net income LTM","Net income FY+1e","Role"]
hdr(ws,4,cols)
peers=["Wesco","Grainger","Ferguson","Watsco","Bunzl"]; refs=["Fastenal","Legrand","Schneider"]
tk={"Wesco":"WCC","Grainger":"GWW","Ferguson":"FERG","Watsco":"WSO","Bunzl":"BNZL.L","Fastenal":"FAST","Legrand":"LR.PA","Schneider":"SU.PA"}
def pdata(k):
    i=info[k]; f=fin[k]; m=1e6
    px=i["currentPrice"]/(100 if k=="Bunzl" else 1); sh=i["sharesOutstanding"]/m; nd=(f.get("BS_Net Debt") or 0)/m
    def v(n):
        x=f.get("LTM_"+n)
        if x is None:
            fy=f.get("FY_"+n) or {}
            x=list(fy.values())[0] if fy else None
        return x/m if x else None
    est=f.get("est_rev",{}); eps=f.get("est_eps",{})
    return dict(px=px,sh=sh,nd=nd,s=v("Total Revenue"),ebitda=v("EBITDA"),ebit=v("EBIT"),ni=v("Net Income"),s1=(est.get("+1y") or 0)/m,eps1=eps.get("+1y"))
r=5; rows={}
for k in peers+refs:
    d=pdata(k); rows[k]=r
    put(ws,r,1,k,bold=False).font=BLK; put(ws,r,2,"GBP" if k=="Bunzl" else info[k]["currency"]).font=BLK
    put(ws,r,3,d["px"],E2); put(ws,r,4,d["sh"],N1); put(ws,r,5,f"=C{r}*D{r}"); put(ws,r,6,d["nd"]); put(ws,r,7,f"=E{r}+F{r}")
    put(ws,r,8,d["s"]); put(ws,r,9,d["s1"] or None); put(ws,r,10,d["ebitda"]); put(ws,r,11,f"=I{r}*J{r}/H{r}"); put(ws,r,12,d["ebit"]); put(ws,r,13,f"=I{r}*L{r}/H{r}"); put(ws,r,14,d["ni"]); put(ws,r,15,f"=D{r}*{d['eps1']}" if d["eps1"] else None)
    put(ws,r,16,"Core peer" if k in peers else "Reference – excluded from averages (Fastenal: premium MRO multiple; Legrand/Schneider: manufacturers)").font=NOTE; r+=1
rx=r; rows["Rexel"]=r
put(ws,r,1,"Rexel (€)",bold=True); put(ws,r,2,"EUR").font=BLK; put(ws,r,3,"=Market!B6",E2); put(ws,r,4,"=Market!B15/1000000",N1); put(ws,r,5,"=C{0}*D{0}".format(r)); put(ws,r,6,"=Market!B20+Market!B22"); put(ws,r,7,f"=E{r}+F{r}")
put(ws,r,8,"=Rexel_Hist!E5"); put(ws,r,9,"=Consensus!C5"); put(ws,r,10,"=Rexel_Hist!E28"); put(ws,r,11,"=Consensus!C19/Consensus!C20"); put(ws,r,12,"=Rexel_Hist!E13"); put(ws,r,13,"=Consensus!C11"); put(ws,r,14,"=Rexel_Hist!E22"); put(ws,r,15,"=Consensus!C16")
put(ws,r,16,"Target. 2025A / 2026e consensus mean excl. GCG. EBITDA = EBITDAaL (2026e = FCF bit ÷ conversion). EBIT = adj. EBITA. NI = recurring NI 2025 / reported NI 2026e. Net debt = 30 Jun 2026 pro forma capital increase + minorities.").font=NOTE
# multiples
r=rx+3; hdr(ws,r,["Multiples","xSales LTM","xSales FY+1","xEBITDA LTM","xEBITDA FY+1","xEBIT LTM","xEBIT FY+1","P/E LTM","P/E FY+1"]); r+=1; m0=r
mrow={}
for k in peers+refs+["Rexel"]:
    q=rows[k]; mrow[k]=r; put(ws,r,1,k if k!="Rexel" else "Rexel (at market)",bold=(k=="Rexel")).font=BOLD if k=="Rexel" else BLK
    for j,(num,den) in enumerate([("G","H"),("G","I"),("G","J"),("G","K"),("G","L"),("G","M"),("E","N"),("E","O")]):
        put(ws,r,2+j,f"={num}{q}/{den}{q}",X)
    r+=1
    if k==peers[-1]:
        pa,pb=mrow[peers[0]],mrow[peers[-1]]; avg=r; 
        put(ws,r,1,"Sector average (core peers)",bold=True)
        for j in range(8): put(ws,r,2+j,f"=AVERAGE({L(2+j)}{pa}:{L(2+j)}{pb})",X,bold=True)
        r+=1; med=r; put(ws,r,1,"Sector median (core peers)",bold=True)
        for j in range(8): put(ws,r,2+j,f"=MEDIAN({L(2+j)}{pa}:{L(2+j)}{pb})",X,bold=True)
        r+=1
# application
r+=1; hdr(ws,r,["Application to Rexel (€m)","xSales LTM","xSales FY+1","xEBITDA LTM","xEBITDA FY+1","xEBIT LTM","xEBIT FY+1","P/E LTM","P/E FY+1"]); r+=1
aggs=["H","I","J","K","L","M","N","O"]
lab(ws,r,"Rexel aggregate"); 
for j,a in enumerate(aggs): put(ws,r,2+j,f"={a}{rx}")
ragg=r; r+=1
for name,src in (("average",avg),("median",med)):
    lab(ws,r,f"Sector multiple ({name})"); 
    for j in range(8): put(ws,r,2+j,f"={L(2+j)}{src}",X)
    rm=r; r+=1
    lab(ws,r,"Implied EV (multiple × aggregate)"); 
    for j in range(6): put(ws,r,2+j,f"={L(2+j)}{rm}*{L(2+j)}{ragg}")
    rev=r; r+=1
    lab(ws,r,"(Net debt + minorities)")
    for j in range(6): put(ws,r,2+j,f"=-$F${rx}")
    r+=1
    lab(ws,r,f"Equity value of Rexel – {name}",bold=True)
    for j in range(6): put(ws,r,2+j,f"={L(2+j)}{rev}+{L(2+j)}{r-1}",res=True)
    for j in (6,7): put(ws,r,2+j,f"={L(2+j)}{rm}*{L(2+j)}{ragg}",res=True)
    req=r; r+=1
    lab(ws,r,"Value per share (€)")
    for j in range(8): put(ws,r,2+j,f"={L(2+j)}{req}/$D${rx}",E2)
    rps=r; r+=1
    lab(ws,r,"Upside / (downside) vs €34.47")
    for j in range(8): put(ws,r,2+j,f"={L(2+j)}{rps}/$C${rx}-1",P)
    r+=2
    if name=="average": req_avg,rps_avg=req,rps
    else: req_med,rps_med=req,rps
# football-field summary rows
hdr(ws,r,["Football field (€ per share)","Low","High","Mid","Comment"]); r+=1; ff0=r
for nm,c1,c2,com in [("P/E (FY+1 / LTM)","I","H","P/E of peers is distorted by Fastenal/Grainger (premium MRO names) – treat as a high bound"),("xEBIT (FY+1 / LTM)","G","F","Rexel adj. EBITA margin 6% vs peers 8-20%: EBIT multiple on higher-margin peers overstates"),("xEBITDA (FY+1 / LTM)","E","D","Preferred multiple (course rule) – EBITDAaL vs US GAAP EBITDA is consistent"),("xSales (FY+1 / LTM)","C","B","Lowest: Rexel's margin is far below the peers → sales multiple undervalues (same lesson as Moncler, reversed)")]:
    lab(ws,r,nm); put(ws,r,2,f"=MIN({c1}{rps_avg},{c2}{rps_avg},{c1}{rps_med},{c2}{rps_med})",E2); put(ws,r,3,f"=MAX({c1}{rps_avg},{c2}{rps_avg},{c1}{rps_med},{c2}{rps_med})",E2); put(ws,r,4,f"=(B{r}+C{r})/2",E2); put(ws,r,5,com).font=NOTE; r+=1
ws.cell(r+1,1,"Reading (course method): compare the four multiples, explain the gaps by margin and D&A intensity differences, keep xEBITDA as the anchor, discuss sample bias (US MRO distributors trade at a premium to a European electrical distributor; Sonepar, the closest peer, is private).").font=NOTE
widths(ws,[34,7,9,10,11,10,11,11,11,11,11,11,11,12,12,70])
# =============================== DCF ===============================
ws=sheet("DCF","DCF valuation (course method, Part 4 – LVMH)","Business plan = Rexel-collected consensus mean excl. GCG (S2), then soft landing to g over 4 years. Flows after leases (EBITDAaL / adj. EBITA), so lease liabilities are NOT in net debt. Full-year 2026 FCF discounted from the valuation date, as in the LVMH case.")
# inputs
I=[("Risk-free rate (rf) – 50% Bund + 50% UST, 30 Oct 2026",4,"=0.5*WACC_inputs!B6+0.5*WACC_inputs!B7",P),("Beta Rexel (b) – own regression 2y vs CAC 40 (1.25) rounded; Yahoo 0.98 as low case",5,1.2,E2),("Market risk premium (Rm−rf) – Damodaran France Jan 2026",6,"=WACC_inputs!B10",P),("Spread (BB+/Ba1)",7,0.015,P),("Interest rate (i) = rf + spread",8,"=B4+B7",P),("Market cap (30 Oct 2026, outstanding shares)",9,"=Market!B16",N),("Net debt (D) – 30 Jun 2026 pro forma capital increase + minorities",10,"=Market!B20+Market!B22",N),("Corporate tax rate (t) – normative",11,"=WACC_inputs!B22",P),("Unlevered beta (b*) = b / (1 + D(1−t)/E)",12,"=B5/(1+B10*(1-B11)/B9)",E2),("Debt beta (bd) = (i − rf) / MRP",13,"=(B8-B4)/B6",E2),("r (unlevered cost of equity) = rf + b*·MRP",14,"=B4+B12*B6",P),("Cost of equity (k) by CAPM = rf + b·MRP",15,"=B4+B5*B6",P),("Cost of equity (k) by MM = r + (r − i)(1−t) D/E",16,"=B14+(B14-B8)*(1-B11)*B10/B9",P),("Cost of capital (K) by WACC (market weights)",17,"=B15*B9/(B9+B10)+B8*(1-B11)*B10/(B9+B10)",P),("Cost of capital (K) by MM = r (1 − D·t/(D+E))",18,"=B14*(1-B10*B11/(B9+B10))",P),("K used for discounting",19,"=B17",P),("Perpetual growth rate (g)",20,"=WACC_inputs!B24",P),("Shares outstanding (m)",21,"=Market!B15/1000000",N1)]
for t,rr,v,f in I: lab(ws,rr,t); put(ws,rr,2,v,f,res=(rr in (17,19)))
ws["B19"].font=BLUE
# business plan: columns B=2025A, C..E consensus, F..I soft landing
r0=24; hdr(ws,r0,["Business plan (€m)","2025A","2026e","2027e","2028e","2029e","2030e","2031e","2032e"]); 
lab(ws,r0-1,"Consensus mean excl. GCG (S2)  →  soft landing (growth converges linearly to g, margins frozen at 2028 level, capex converges to D&A)",note=True)
R={"sales":25,"g":26,"ebitda":28,"em":29,"da":31,"das":32,"ebit":34,"tax":35,"nopat":36,"pda":37,"capex":38,"dwcr":39,"fcf":40,"per":41,"dfcf":42}
lab(ws,R["sales"],"Sales"); put(ws,R["sales"],2,"=Rexel_Hist!E5"); 
for j,c in enumerate("CDE"): put(ws,R["sales"],3+j,f"=Consensus!{'CDE'[j]}5")
for c,p in zip("FGHI","EFGH"): put(ws,R["sales"],ord(c)-64,f"={p}{R['sales']}*(1+{c}{R['g']})")
lab(ws,R["g"],"Sales growth"); 
for c,p in zip("CDE","BCD"): put(ws,R["g"],ord(c)-64,f"={c}{R['sales']}/{p}{R['sales']}-1",P)
for c,p in zip("FGH","EFG"): put(ws,R["g"],ord(c)-64,f"={p}{R['g']}-($E${R['g']}-$I${R['g']})/4",P)
put(ws,R["g"],9,"=B20",P)
lab(ws,R["ebitda"],"EBITDA after leases (EBITDAaL)"); put(ws,R["ebitda"],2,"=Rexel_Hist!E28")
for j,c in enumerate("CDE"): put(ws,R["ebitda"],3+j,f"=Consensus!{'CDE'[j]}19/Consensus!{'CDE'[j]}20")
for c in "FGHI": put(ws,R["ebitda"],ord(c)-64,f"={c}{R['em']}*{c}{R['sales']}")
lab(ws,R["em"],"EBITDAaL margin")
for c in "BCDE": put(ws,R["em"],ord(c)-64,f"={c}{R['ebitda']}/{c}{R['sales']}",P)
for c,p in zip("FGHI","EFGH"): put(ws,R["em"],ord(c)-64,f"={p}{R['em']}",P)
lab(ws,R["da"],"(D&A) = EBIT − EBITDAaL")
for c in "BCDE": put(ws,R["da"],ord(c)-64,f"={c}{R['ebit']}-{c}{R['ebitda']}")
for c in "FGHI": put(ws,R["da"],ord(c)-64,f"={c}{R['das']}*{c}{R['sales']}")
lab(ws,R["das"],"(D&A) / Sales")
for c in "BCDE": put(ws,R["das"],ord(c)-64,f"={c}{R['da']}/{c}{R['sales']}",'0.00%')
for c,p in zip("FGHI","EFGH"): put(ws,R["das"],ord(c)-64,f"={p}{R['das']}",'0.00%')
lab(ws,R["ebit"],"EBIT = current adjusted EBITA"); put(ws,R["ebit"],2,"=Rexel_Hist!E13")
for j,c in enumerate("CDE"): put(ws,R["ebit"],3+j,f"=Consensus!{'CDE'[j]}11")
for c in "FGHI": put(ws,R["ebit"],ord(c)-64,f"={c}{R['ebitda']}+{c}{R['da']}")
lab(ws,R["tax"],"(Corporate tax on EBIT)")
for c in "CDEFGHI": put(ws,R["tax"],ord(c)-64,f"=-$B$11*{c}{R['ebit']}")
lab(ws,R["nopat"],"= NOPAT")
for c in "CDEFGHI": put(ws,R["nopat"],ord(c)-64,f"={c}{R['ebit']}+{c}{R['tax']}")
lab(ws,R["pda"],"+ D&A")
for c in "CDEFGHI": put(ws,R["pda"],ord(c)-64,f"=-{c}{R['da']}")
lab(ws,R["capex"],"(Capex) – 0.7% of sales 2026-28 (Rexel 2025 ratio, S4), then converges to D&A")
put(ws,R["capex"],2,"=Rexel_Hist!E33")
for c in "CDE": put(ws,R["capex"],ord(c)-64,f"=-0.007*{c}{R['sales']}")
for c,p in zip("FGH","EFG"): put(ws,R["capex"],ord(c)-64,f"={p}{R['capex']}-($E${R['capex']}-$I${R['capex']})/4")
put(ws,R["capex"],9,f"=-I{R['pda']}")
lab(ws,R["dwcr"],"(Δ WCR)")
for c,p in zip("CDEFGHI","BCDEFGH"): put(ws,R["dwcr"],ord(c)-64,f"=-({c}58-{p}58)")
lab(ws,R["fcf"],"= Annual FCF",bold=True)
for c in "CDEFGHI": put(ws,R["fcf"],ord(c)-64,f"=SUM({c}{R['nopat']}:{c}{R['dwcr']})",bold=True)
lab(ws,R["per"],"Discount period (years from valuation date)"); put(ws,R["per"],3,"=C63",E2)
for c,p in zip("DEFGHI","CDEFGH"): put(ws,R["per"],ord(c)-64,f"=1+{p}{R['per']}",E2)
lab(ws,R["dfcf"],"Discounted FCF")
for c in "CDEFGHI": put(ws,R["dfcf"],ord(c)-64,f"={c}{R['fcf']}/(1+$B$19)^{c}{R['per']}")
# financial balance sheet 2025
lab(ws,45,"Financial balance sheet 31 Dec 2025 (S5)",bold=True)
lab(ws,46,"Non-current assets (goodwill, intangibles, PP&E, ROU, LT investments, DTA)"); put(ws,46,2,"=3907.6+1659.3+353.7+1342.2+61.7+30.7")
lab(ws,47,"WCR (inventories + trade & other receivables + tax assets − payables − tax & other current liabilities)"); put(ws,47,2,"=2364.1+2682.0+34.1+893.6-2170.9-11.8-1039.5")
lab(ws,48,"= Total assets"); put(ws,48,2,"=B46+B47",bold=True)
lab(ws,50,"Equity (incl. minorities)"); put(ws,50,2,5405.7)
lab(ws,51,"Provisions (pensions, DTL, other provisions)"); put(ws,51,2,"=100.1+309.4+215.4")
lab(ws,52,"Lease liabilities"); put(ws,52,2,"=1248.2+240.1")
lab(ws,53,"Net financial debt (debt + accrued interest − cash)"); put(ws,53,2,"=2934.2+671.5+19.6-1037.5")
lab(ws,54,"= Total equity & liabilities"); put(ws,54,2,"=SUM(B50:B53)",bold=True)
lab(ws,55,"Check (assets − liabilities)",note=True); put(ws,55,2,"=B48-B54",N1)
lab(ws,57,"WCR / Sales"); put(ws,57,2,"=B47/B25",P)
for c,p in zip("CDEFGHI","BCDEFGH"): put(ws,57,ord(c)-64,f"={p}57",P)
lab(ws,58,"WCR"); put(ws,58,2,"=B47")
for c in "CDEFGHI": put(ws,58,ord(c)-64,f"={c}57*{c}25")
lab(ws,61,"Closing date"); ws["C61"]="2026-12-31"; 
import datetime; ws["C61"]=datetime.datetime(2026,12,31); ws["C61"].number_format="dd/mm/yyyy"; ws["C61"].font=BLUE
lab(ws,62,"Valuation date"); ws["C62"]=datetime.datetime(2026,10,30); ws["C62"].number_format="dd/mm/yyyy"; ws["C62"].font=BLUE
lab(ws,63,"Fraction of year to closing"); put(ws,63,3,"=(C61-C62)/365",E2)
# results block (columns K-M)
lab(ws,4,"",); 
for rr,t,v,f in [(4,"Sum of discounted FCF 2026-2032","=SUM(C42:I42)",N),(5,"Terminal value (Gordon-Shapiro) = FCF2032 × (1+g) / (K − g)","=I40*(1+B20)/(B19-B20)",N),(6,"Discounted terminal value","=L5/(1+B19)^I41",N),(7,"Enterprise value (EV)","=L4+L6",N),(8,"TV as % of EV","=L6/L7",P),(9,"(Net debt + minorities)","=-B10",N),(10,"Equity value",f"=L7+L9",N),(11,"Value per share (€)","=L10/B21",E2),(12,"Share price 30 Oct 2026 (€)","=Market!B6",E2),(13,"Upside / (downside)","=L11/L12-1",P),(14,"Implied EV / EBITDAaL 2026e","=L7/C28",X),(15,"Implied EV / adj. EBITA 2026e","=L7/C34",X)]:
    c=ws.cell(rr,11,t); c.font=BOLD if rr in (10,11) else BLK; put(ws,rr,12,v,f,res=(rr in (10,11,13)))
# sensitivity (live formulas): value per share for g (rows) × K (cols)
lab(ws,17,"")
ws.cell(17,11,"Sensitivity – value per share (€): g (down) × K (across)").font=BOLD
Ks=[0.075,0.08,0.085,0.09,0.095,0.10]; gs=[0.01,0.015,0.02,0.025,0.03]
ws.cell(18,11,"g \\ K").font=BOLD
for j,K in enumerate(Ks): put(ws,18,12+j,K,P)
for i,g in enumerate(gs):
    put(ws,19+i,11,g,P)
    for j in range(len(Ks)):
        Kc=f"{L(12+j)}$18"; gc=f"$K{19+i}"
        put(ws,19+i,12+j,f"=(SUMPRODUCT($C$40:$I$40/(1+{Kc})^$C$41:$I$41)+$I$40*(1+{gc})/({Kc}-{gc})/(1+{Kc})^$I$41-$B$10)/$B$21",E2)
ws.cell(25,11,"Second sensitivity – value per share: 2028+ EBITDAaL margin (down) × K (across)").font=BOLD
ws.cell(26,11,"margin \\ K").font=BOLD
for j,K in enumerate(Ks): put(ws,26,12+j,K,P)
ws.cell(27,11,"(static: rebuild with Excel Data Table if needed – margin case is handled by changing E28 in the plan)").font=NOTE
widths(ws,[62,11,11,11,11,11,11,11,11,3,52,12,10,10,10,10,10])
ws.freeze_panes="B4"
# =============================== LBO ===============================
ws=sheet("LBO","LBO-style bid by an investment fund (course method, Part 5 – Carrefour)","A holding (NewCo) buys 100% of Rexel at the end of 2026, financed by Senior A (amortising 7y) + Senior B (bullet 8y) + Junior C (bullet 9y) + fund equity. Tax consolidation: the holding deducts acquisition interest from Rexel's taxable income. The fund exits in 2029 (3 years) at the entry xEBIT multiple and requires a 20% IRR → the equity it can inject + the debt raised = the maximum bid.")
# target income statement
hdr(ws,4,["Rexel income statement (€m)","2025A","2026e","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,3,"Consensus mean excl. GCG 2026-28, then soft landing (7 steps to g, as in the Carrefour case)",note=True)
lab(ws,5,"Sales"); put(ws,5,2,"=DCF!B25")
for j,c in enumerate("CDE"): put(ws,5,3+j,f"=DCF!{c}25")
for c,p in zip("FGHIJ","EFGHI"): put(ws,5,ord(c)-64,f"={p}5*(1+{c}6)")
lab(ws,6,"Sales growth"); 
for c,p in zip("CDE","BCD"): put(ws,6,ord(c)-64,f"={c}5/{p}5-1",P)
for c,p in zip("FGHIJ","EFGHI"): put(ws,6,ord(c)-64,f"={p}6-($E$6-$B$7)/7",P)
lab(ws,7,"Long-term growth g"); put(ws,7,2,"=DCF!B20",P)
lab(ws,8,"EBITDAaL"); 
for c in "BCDE": put(ws,8,ord(c)-64,f"=DCF!{c}28")
for c in "FGHIJ": put(ws,8,ord(c)-64,f"={c}9*{c}5")
lab(ws,9,"EBITDAaL margin")
for c in "BCDE": put(ws,9,ord(c)-64,f"={c}8/{c}5",P)
for c,p in zip("FGHIJ","EFGHI"): put(ws,9,ord(c)-64,f"={p}9",P)
lab(ws,10,"(D&A)")
for c in "BCDE": put(ws,10,ord(c)-64,f"={c}12-{c}8")
for c in "FGHIJ": put(ws,10,ord(c)-64,f"={c}11*{c}5")
lab(ws,11,"(D&A)/Sales")
for c in "BCDE": put(ws,11,ord(c)-64,f"={c}10/{c}5",'0.00%')
for c,p in zip("FGHIJ","EFGHI"): put(ws,11,ord(c)-64,f"={p}11",'0.00%')
lab(ws,12,"EBIT (adj. EBITA)")
for c in "BCDE": put(ws,12,ord(c)-64,f"=DCF!{c}34")
for c in "FGHIJ": put(ws,12,ord(c)-64,f"={c}8+{c}10")
lab(ws,13,"(Interest on Rexel's own net debt)"); put(ws,13,2,-214.2)
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,13,ord(c)-64,f"=-$B$14*{p}40")
ws.cell(13,12,"Interest on OPENING net debt (the Carrefour case uses the average, which creates a circular reference needing iterative calculation; opening debt avoids it, difference is immaterial).").font=NOTE
lab(ws,14,"Interest rate on Rexel's net debt (i, from DCF)"); put(ws,14,2,"=DCF!B8",P)
lab(ws,15,"EBT")
for c in "CDEFGHIJ": put(ws,15,ord(c)-64,f"={c}12+{c}13")
lab(ws,16,"(Corporate tax)")
for c in "CDEFGHIJ": put(ws,16,ord(c)-64,f"=-$B$17*{c}15")
lab(ws,17,"Corporate tax rate"); put(ws,17,2,"=DCF!B11",P)
lab(ws,18,"Net income",bold=True)
for c in "CDEFGHIJ": put(ws,18,ord(c)-64,f"={c}15+{c}16",bold=True)
# target cash flow
hdr(ws,21,["Rexel cash flow = Δ net debt (€m)","2025A","2026e","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,22,"Net income"); 
for c in "CDEFGHIJ": put(ws,22,ord(c)-64,f"={c}18")
lab(ws,23,"(Dividend paid to the holding) – payout"); put(ws,23,2,1.0,P)
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,23,ord(c)-64,f"=-$B$23*{c}22")
ws.cell(23,12,"Under the LBO 100% of net income is upstreamed to service the acquisition debt (Carrefour case: dividend = previous NI). Set 100% here.").font=NOTE
lab(ws,24,"+ D&A")
for c in "CDEFGHIJ": put(ws,24,ord(c)-64,f"=-{c}10")
lab(ws,25,"(Capex)")
for c in "CDEFGHI": put(ws,25,ord(c)-64,f"=DCF!{c}38")
put(ws,25,10,"=I25")
lab(ws,26,"(Δ WCR)")
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,26,ord(c)-64,f"=-({c}35-{p}35)")
lab(ws,27,"= Δ cash flow Rexel (reduces net debt)",bold=True)
for c in "CDEFGHIJ": put(ws,27,ord(c)-64,f"=SUM({c}22:{c}26)",bold=True)
# target balance sheet
hdr(ws,30,["Rexel simplified financial balance sheet (€m)","2025A","2026e","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,31,"Non-current assets"); put(ws,31,2,"=DCF!B46")
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,31,ord(c)-64,f"={p}31-{c}25-{c}24")
lab(ws,35,"WCR"); put(ws,35,2,"=DCF!B47")
for c in "CDEFGHIJ": put(ws,35,ord(c)-64,f"={c}36*{c}5")
lab(ws,36,"WCR / Sales"); put(ws,36,2,"=B35/B5",P)
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,36,ord(c)-64,f"={p}36",P)
lab(ws,32,"= Total assets"); 
for c in "BCDEFGHIJ": put(ws,32,ord(c)-64,f"={c}31+{c}35",bold=True)
lab(ws,38,"Equity"); put(ws,38,2,"=DCF!B50")
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,38,ord(c)-64,f"={p}38+{c}22+{c}23")
lab(ws,39,"Provisions + lease liabilities (held constant)"); put(ws,39,2,"=DCF!B51+DCF!B52")
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,39,ord(c)-64,f"={p}39")
lab(ws,40,"Net financial debt"); put(ws,40,2,"=DCF!B53")
for c,p in zip("CDEFGHIJ","BCDEFGHI"): put(ws,40,ord(c)-64,f"={p}40-{c}27")
lab(ws,41,"= Total equity & liabilities")
for c in "BCDEFGHIJ": put(ws,41,ord(c)-64,f"=SUM({c}38:{c}40)",bold=True)
lab(ws,42,"Check",note=True)
for c in "BCDEFGHIJ": put(ws,42,ord(c)-64,f"={c}32-{c}41",N1)
ws.cell(42,12,"2026e net debt here is the model's roll-forward from 31 Dec 2025 (dividend = 100% of NI, no M&A). Actual 30 Jun 2026 was 3,322 after €398m M&A and €353m dividends – consistent order of magnitude.").font=NOTE
# holding income statement (acquisition at 31 Dec 2026, first year 2027 = column D)
hdr(ws,45,["Holding income statement (€m)","","","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,46,"Dividend received from Rexel")
for c in "DEFGHIJ": put(ws,46,ord(c)-64,f"=-{c}23")
lab(ws,47,"(Tax on dividend: 5% taxable share × tax rate – parent-subsidiary regime)"); put(ws,47,2,0.05,P)
for c in "DEFGHIJ": put(ws,47,ord(c)-64,f"=-$B$17*$B$47*{c}46")
lab(ws,48,"+ Tax saved: Rexel's own tax now paid at group level")
for c in "DEFGHIJ": put(ws,48,ord(c)-64,f"=-{c}16")
lab(ws,49,"(Interest paid on acquisition debt)")
for c in "DEFGHIJ": put(ws,49,ord(c)-64,f"=-{c}66-{c}72-{c}78")
lab(ws,50,"(Tax paid for the fiscal group)")
for c in "DEFGHIJ": put(ws,50,ord(c)-64,f"=-{c}57")
lab(ws,51,"= Holding net income",bold=True)
for c in "DEFGHIJ": put(ws,51,ord(c)-64,f"=SUM({c}46:{c}50)",bold=True)
hdr(ws,53,["Tax of the fiscal group (€m)","","","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,54,"EBT Rexel")
for c in "DEFGHIJ": put(ws,54,ord(c)-64,f"={c}15")
lab(ws,55,"(Interest on acquisition debt)")
for c in "DEFGHIJ": put(ws,55,ord(c)-64,f"={c}49")
lab(ws,56,"= Taxable income of the group")
for c in "DEFGHIJ": put(ws,56,ord(c)-64,f"={c}54+{c}55")
lab(ws,57,"Tax paid for the group")
for c in "DEFGHIJ": put(ws,57,ord(c)-64,f"=$B$17*{c}56")
# debt schedule
hdr(ws,60,["Acquisition debt schedule (€m)","Share / rate","2026e","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,61,"Total acquisition debt raised = x EBITDAaL 2026e"); put(ws,61,2,4.0,X); put(ws,61,3,"=B61*C8")
ws.cell(61,12,"Leverage assumption of the bid (acquisition debt only; Rexel's existing net debt stays in the target). 4.0x EBITDAaL 2026e on top of Rexel 2.0x = c.6x total. Sensitise 3.5-5.0x.").font=NOTE
lab(ws,62,"Outstanding Senior A (constant amortisation, 7 years)"); put(ws,62,2,0.5,P); put(ws,62,3,"=B62*C61")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,62,ord(c)-64,f"={p}62+{c}63")
lab(ws,63,"(Amortisation)")
for c in "DEFGHIJ": put(ws,63,ord(c)-64,"=-$C$62/7")
lab(ws,64,"Interest rate Senior A = Euribor 6M + spread + swap"); put(ws,64,2,"=SUM(B65:B65)+0.0275+0.005",P)
lab(ws,65,"Euribor 6M (30 Oct 2026, to verify on euribor-rates.eu)"); put(ws,65,2,0.022,P)
lab(ws,66,"Interest paid on Senior A")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,66,ord(c)-64,f"=$B$64*AVERAGE({p}62:{c}62)")
lab(ws,68,"Outstanding Senior B (bullet, 8 years)"); put(ws,68,2,0.35,P); put(ws,68,3,"=B68*C61")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,68,ord(c)-64,f"={p}68+{c}69")
lab(ws,69,"(Amortisation)")
for c in "DEFGHIJ": put(ws,69,ord(c)-64,0)
lab(ws,70,"Interest rate Senior B = Euribor + 3.25% + swap"); put(ws,70,2,"=B65+0.0325+0.005",P)
lab(ws,72,"Interest paid on Senior B")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,72,ord(c)-64,f"=$B$70*AVERAGE({p}68:{c}68)")
lab(ws,74,"Outstanding Junior C (bullet, 9 years)"); put(ws,74,2,"=1-B62-B68",P); put(ws,74,3,"=B74*C61")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,74,ord(c)-64,f"={p}74+{c}75")
lab(ws,75,"(Amortisation)")
for c in "DEFGHIJ": put(ws,75,ord(c)-64,0)
lab(ws,76,"Interest rate Junior C = Euribor + 6% + swap"); put(ws,76,2,"=B65+0.06+0.005",P)
lab(ws,78,"Interest paid on Junior C")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,78,ord(c)-64,f"=$B$76*AVERAGE({p}74:{c}74)")
# holding cash flow & BS
hdr(ws,81,["Holding cash flow (€m)","","","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,82,"Holding net income")
for c in "DEFGHIJ": put(ws,82,ord(c)-64,f"={c}51")
lab(ws,83,"(Debt amortisation)")
for c in "DEFGHIJ": put(ws,83,ord(c)-64,f"={c}63+{c}69+{c}75")
lab(ws,84,"= Holding cash flow",bold=True)
for c in "DEFGHIJ": put(ws,84,ord(c)-64,f"={c}82+{c}83",bold=True)
hdr(ws,86,["Holding balance sheet (€m)","","2026e","2027e","2028e","2029e","2030e","2031e","2032e","2033e"])
lab(ws,87,"Rexel shares (acquisition price)"); put(ws,87,3,"=C91+C92+C93")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,87,ord(c)-64,f"={p}87")
lab(ws,88,"Cash"); put(ws,88,3,0)
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,88,ord(c)-64,f"={p}88+{c}84")
lab(ws,89,"= Total assets")
for c in "CDEFGHIJ": put(ws,89,ord(c)-64,f"={c}87+{c}88",bold=True)
lab(ws,91,"Equity (fund capital entry)"); put(ws,91,3,"=C112")
for c,p in zip("DEFGHIJ","CDEFGHI"): put(ws,91,ord(c)-64,f"={p}91+{c}51")
lab(ws,92,"Senior A + B")
for c in "CDEFGHIJ": put(ws,92,ord(c)-64,f"={c}62+{c}68")
lab(ws,93,"Junior C")
for c in "CDEFGHIJ": put(ws,93,ord(c)-64,f"={c}74")
lab(ws,94,"= Total equity & liabilities")
for c in "CDEFGHIJ": put(ws,94,ord(c)-64,f"=SUM({c}91:{c}93)",bold=True)
# capital entry / bid
hdr(ws,97,["Determination of the fund's capital entry and bid (€m)","","Entry 2026e","","","Exit 2029e"])
lab(ws,98,"Rexel market cap (30 Oct 2026)"); put(ws,98,3,"=Market!B16")
lab(ws,99,"Rexel net financial debt (model, end 2026)"); put(ws,99,3,"=C40")
lab(ws,100,"Rexel enterprise value at market"); put(ws,100,3,"=C98+C99")
lab(ws,101,"EBIT (adj. EBITA)"); put(ws,101,3,"=C12"); put(ws,101,6,"=F12")
lab(ws,102,"xEBIT entry = exit multiple"); put(ws,102,3,"=C100/C101",X); put(ws,102,6,"=C102",X)
lab(ws,103,"Rexel enterprise value at exit"); put(ws,103,6,"=F102*F101")
lab(ws,104,"Consolidated net debt at exit = Rexel net debt + holding debt − holding cash"); put(ws,104,6,"=F40+F92+F93-F88")
lab(ws,105,"Rexel economic equity value at exit (100%)",bold=True); put(ws,105,6,"=F103-F104",res=True)
lab(ws,106,"IRR required by the fund"); put(ws,106,6,0.20,P)
lab(ws,107,"Holding period (years)"); put(ws,107,6,3)
lab(ws,109,"Fund capital entry = exit equity / (1+IRR)^n",bold=True); put(ws,109,3,"=F105/(1+F106)^F107",res=True)
lab(ws,110,"+ Acquisition debt raised"); put(ws,110,3,"=C61")
lab(ws,111,"= Maximum price the fund can pay for 100% of the equity (the bid)",bold=True); put(ws,111,3,"=C109+C110",res=True)
lab(ws,112,"(link for the holding balance sheet)",note=True); put(ws,112,3,"=C109")
lab(ws,113,"Bid per share (€)",bold=True); put(ws,113,3,"=C111/DCF!B21",E2,res=True)
lab(ws,114,"Premium / (discount) vs share price €34.47"); put(ws,114,3,"=C113/Market!B6-1",P)
lab(ws,115,"Implied entry EV / EBITDAaL 2026e"); put(ws,115,3,"=(C111+C99)/C8",X)
lab(ws,116,"Equity share of sources"); put(ws,116,3,"=C109/C111",P)
lab(ws,117,"Fund money multiple (MoM) = exit equity / entry equity"); put(ws,117,3,"=F105/C109",X)
lab(ws,118,"Net debt / EBITDAaL at entry (Rexel + acquisition debt)"); put(ws,118,3,"=(C40+C61)/C8",X)
lab(ws,120,"Reading: the bid is the price at which the fund earns exactly 20% a year over 3 years, assuming exit at the entry multiple. If the bid per share is below the market price, an LBO is not feasible at 20% – the fund would need a lower IRR, more leverage, operational improvement (margin) or a multiple re-rating. Sensitise the IRR (15-25%), leverage (4-5.5x) and exit multiple.",note=True)
widths(ws,[62,12,11,11,11,11,11,11,11,11,3,80])
ws.freeze_panes="B5"
# =============================== SUMMARY ===============================
ws=sheet("Summary","Valuation summary – football field (€ per share)","Feeds the football-field chart. Low / high per method; market reference lines: share price 34.47, 52-week range 28.83-39.80, capital increase price 34.75, analyst targets 30-46 (mean 41.9).")
hdr(ws,4,["Method","Low","High","Mid","Source cell / comment"])
items=[("52-week trading range","=Market!B9","=Market!B8","Market sheet (S10)"),("Analyst target prices (13 analysts)",30,46,"Yahoo / S10 – mean 41.86"),("Capital increase price (29 Sep 2026)",34.75,34.75,"S7"),
("Peers – xSales","=Comps!B%d"%ff0 if False else None,None,"Comps football-field block"),]
r=5
for t,lo,hi,com in [("52-week trading range","=Market!B9","=Market!B8","Market (S10)"),("Analyst target prices","=Market!B29*0+30","=46","S10 – 13 analysts, mean 41.86"),("Capital increase price (29 Sep 2026)","=Market!B12","=Market!B12","S7"),
 ("Peers – xSales",f"=Comps!B{ff0+3}",f"=Comps!C{ff0+3}","Comps"),("Peers – xEBITDA (anchor)",f"=Comps!B{ff0+2}",f"=Comps!C{ff0+2}","Comps"),("Peers – xEBIT",f"=Comps!B{ff0+1}",f"=Comps!C{ff0+1}","Comps"),("Peers – P/E",f"=Comps!B{ff0}",f"=Comps!C{ff0}","Comps"),
 ("DCF – sensitivity K 8.5-9.5%, g 1.5-2.5%","=MIN(DCF!N20:P22)","=MAX(DCF!N20:P22)","DCF sensitivity table (inner block)"),("DCF – base case","=DCF!L11","=DCF!L11","DCF"),("LBO bid (20% IRR, 4.0x acquisition debt)","=LBO!C113","=LBO!C113","LBO")]:
    lab(ws,r,t); put(ws,r,2,lo,E2); put(ws,r,3,hi,E2); put(ws,r,4,f"=(B{r}+C{r})/2",E2); ws.cell(r,5,com).font=NOTE; r+=1
lab(ws,r+1,"Share price 30 Oct 2026"); put(ws,r+1,2,"=Market!B6",E2)
widths(ws,[42,10,10,10,50])
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print("saved")
