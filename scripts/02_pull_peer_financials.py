import yfinance as yf, pandas as pd, numpy as np, json, warnings; warnings.filterwarnings("ignore")
tick = {"Rexel":"RXL.PA","Wesco":"WCC","Grainger":"GWW","Ferguson":"FERG","Watsco":"WSO","Bunzl":"BNZL.L","Fastenal":"FAST","Legrand":"LR.PA","Schneider":"SU.PA"}
rows=["Total Revenue","EBITDA","EBIT","Operating Income","Net Income","Reconciled Depreciation","Diluted EPS","Basic Average Shares"]
out={}
for k,t in tick.items():
    T=yf.Ticker(t)
    a=T.income_stmt; q=T.quarterly_income_stmt
    d={}
    for r in rows:
        if r in a.index: d["FY_"+r]=a.loc[r].dropna().iloc[:4].to_dict()
        if r in q.index: d["LTM_"+r]=float(q.loc[r].dropna().iloc[:4].sum())
    bs=T.balance_sheet
    for r in ["Total Debt","Cash And Cash Equivalents","Net Debt","Stockholders Equity","Minority Interest","Working Capital"]:
        if r in bs.index: d["BS_"+r]=float(bs.loc[r].dropna().iloc[0])
    try:
        re=T.revenue_estimate; ee=T.earnings_estimate
        d["est_rev"]=re["avg"].to_dict(); d["est_eps"]=ee["avg"].to_dict()
    except Exception as e: d["est_err"]=str(e)
    try: d["growth"]=T.growth_estimates["stock"].to_dict()
    except Exception as e: pass
    try: d["target"]=T.analyst_price_targets
    except Exception as e: pass
    out[k]={kk:({str(x):v for x,v in vv.items()} if isinstance(vv,dict) else vv) for kk,vv in d.items()}
json.dump(out,open("financials.json","w"),indent=1,default=str)
for k,v in out.items():
    print("==",k); 
    for kk in ["LTM_Total Revenue","LTM_EBITDA","LTM_EBIT","LTM_Net Income","BS_Net Debt","est_rev","est_eps","target"]:
        print("  ",kk,v.get(kk))
# regression beta for Rexel: 2y weekly vs CAC40 and STOXX600
px=pd.read_csv("prices.csv",index_col=0,parse_dates=True)
w=px.resample("W-FRI").last().loc["2024-10-01":]
r=w.pct_change().dropna()
for idx in ["CAC40","SXXP"]:
    for nm in ["Rexel","Legrand","Schneider"]:
        b=np.cov(r[nm],r[idx])[0,1]/np.var(r[idx],ddof=1); print(f"beta {nm} vs {idx} 2y weekly: {b:.2f}")
w5=px.resample("W-FRI").last().loc["2021-10-01":].pct_change().dropna()
print("beta Rexel vs CAC40 5y weekly:", round(np.cov(w5["Rexel"],w5["CAC40"])[0,1]/np.var(w5["CAC40"],ddof=1),2))
