import yfinance as yf, pandas as pd, json
tick = {"Rexel":"RXL.PA","Wesco":"WCC","Grainger":"GWW","Ferguson":"FERG","Watsco":"WSO","Bunzl":"BNZL.L","Fastenal":"FAST","Legrand":"LR.PA","Schneider":"SU.PA","Sonepar(na)":None,"CAC40":"^FCHI","SXXP":"^STOXX"}
tick = {k:v for k,v in tick.items() if v}
px = yf.download(list(tick.values()), start="2021-10-01", end="2026-10-31", auto_adjust=False, progress=False)["Close"]
px.columns = [k for v in px.columns for k,t in tick.items() if t==v]
px.to_csv("prices.csv")
print(px.tail(3).round(2).to_string())
info = {}
for k,t in tick.items():
    if k in ("CAC40","SXXP"): continue
    i = yf.Ticker(t).info
    info[k] = {f:i.get(f) for f in ["currency","currentPrice","previousClose","marketCap","sharesOutstanding","beta","enterpriseValue","totalDebt","totalCash","trailingPE","forwardPE","enterpriseToEbitda","dividendYield","fiftyTwoWeekHigh","fiftyTwoWeekLow"]}
json.dump(info, open("info.json","w"), indent=1)
print(pd.DataFrame(info).T.to_string())
