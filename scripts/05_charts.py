import pandas as pd, numpy as np, json, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as mdates
from matplotlib.patches import FancyBboxPatch
OUT="../Charts/"
BLUE,ORANGE,AQUA,YEL="#0B2545","#C8692B","#4F7CAC","#E9B44C"; INK,INK2,MUTE,GRID="#1A1A1A","#52514e","#8A8F98","#E3E5E9"; SURF="#ffffff"
plt.rcParams.update({"font.family":"Arial","font.size":9,"axes.edgecolor":GRID,"axes.linewidth":0.8,"axes.spines.top":False,"axes.spines.right":False,"axes.titleweight":"bold","axes.titlesize":10,"axes.titlecolor":INK,"xtick.color":INK2,"ytick.color":INK2,"axes.labelcolor":INK2,"figure.facecolor":SURF,"axes.facecolor":SURF,"legend.frameon":False,"legend.fontsize":8})
def save(fig,name):
    fig.savefig(OUT+name+".png",dpi=250,bbox_inches="tight"); fig.savefig(OUT+name+".pdf",bbox_inches="tight"); plt.close(fig); print("saved",name)
def src(ax,text): ax.figure.text(0.01,-0.02,text,fontsize=7,color=MUTE,ha="left",va="top")
# ---------- 1. stock performance
px=pd.read_csv("prices.csv",index_col=0,parse_dates=True).ffill()
start="2024-10-30"; w=px.loc[start:,["Rexel","Wesco","CAC40","SXXP"]]; rb=w/w.iloc[0]*100
fig,ax=plt.subplots(figsize=(9,4.2))
for col,c,lw,lab in [("SXXP","#C3C6CC",1.3,"STOXX 600"),("CAC40","#8A8F98",1.3,"CAC 40"),("Wesco",ORANGE,1.5,"Wesco"),("Rexel",BLUE,2.1,"Rexel")]:
    ax.plot(rb.index,rb[col],color=c,lw=lw,label=lab); ax.text(rb.index[-1]+pd.Timedelta(days=6),rb[col].iloc[-1],f"{lab}  {rb[col].iloc[-1]:.0f}",color=c if col in("Rexel","Wesco") else INK2,fontsize=8,va="center",fontweight="bold" if col=="Rexel" else "normal")
ax.axhline(100,color=GRID,lw=0.8)
events=[("2026-02-11","FY25 results\nguidance 3-5%",1),("2026-04-22","Q1: +3.4%",-1.3),("2026-07-27","H1: guidance\nraised to c.5%",1),("2026-09-25","GCG $1.4bn +\n€500m raise @34.75",-1.9),("2026-10-22","Q3: +11.4%,\nguidance c.7%",1)]
for d,t,s in events:
    d=pd.Timestamp(d); y=rb.loc[:d,"Rexel"].iloc[-1]
    ax.plot([d],[y],"o",ms=6,color=BLUE,mec=SURF,mew=1.2)
    ax.annotate(t,(d,y),xytext=(0,38*s),textcoords="offset points",ha="center",fontsize=7,color=INK2,arrowprops=dict(arrowstyle="-",color=MUTE,lw=0.6))
ax.set_title("Rexel share price vs peers and indices – rebased to 100 on 30 Oct 2024 (local currency)")
ax.yaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y")); ax.set_xlim(rb.index[0],rb.index[-1]+pd.Timedelta(days=75))
ax.legend(loc="upper left",ncol=4)
lo,hi=px.loc["2025-10-30":,"Rexel"].min(),px.loc["2025-10-30":,"Rexel"].max()
ax.text(0.01,0.03,f"Rexel 30 Oct 2026: €{px['Rexel'].iloc[-1]:.2f}  |  2-year performance {rb['Rexel'].iloc[-1]-100:+.0f}%  |  52-week range €28.83 – €39.80 (intraday)",transform=ax.transAxes,fontsize=8,color=INK2)
src(ax,"Source: Yahoo Finance daily closes (S10); Rexel press releases for events (S1, S3, S4, S6, S7). Price return, local currency.")
save(fig,"01_stock_performance")
# ---------- 2. football field
rows=[("52-week trading range",28.83,39.80,"#C3C6CC"),("Analyst target prices (13)",30.0,46.0,"#C3C6CC"),("Peers – P/E",43.5,52.7,AQUA),("Peers – xEBIT",51.7,56.1,AQUA),("Peers – xEBITDA (anchor)",46.1,50.3,AQUA),("DCF – K 8.5-9.5%, g 1.5-2.5%",36.6,49.5,BLUE),("LBO bid – IRR 20%, 4.0x debt",37.6,37.6,ORANGE)]
fig,ax=plt.subplots(figsize=(9,4.2))
for i,(n,lo,hi,c) in enumerate(rows):
    y=len(rows)-1-i
    if hi-lo<0.6: ax.plot([lo],[y],"D",color=c,ms=9,mec=SURF,mew=1); ax.text(lo+0.6,y,f"€{lo:.1f}",va="center",fontsize=8,color=INK2)
    else:
        ax.barh(y,hi-lo,left=lo,height=0.5,color=c); ax.text(lo-0.5,y,f"{lo:.0f}",va="center",ha="right",fontsize=8,color=INK2); ax.text(hi+0.5,y,f"{hi:.0f}",va="center",ha="left",fontsize=8,color=INK2)
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows][::-1],fontsize=8.5,color=INK)
for x,t,c in [(34.47,"Share price €34.47",INK),(34.75,"Issue price €34.75",MUTE),(43.1,"DCF base €43.1",ORANGE)]:
    ax.axvline(x,color=c,lw=1 if x!=34.75 else 0.8,ls="-" if x!=34.75 else ":"); ax.text(x+(0.3 if x==34.75 else 0),(len(rows)-0.35) if x!=34.75 else -0.65,t,rotation=0,fontsize=7.5,color=c,ha="center" if x!=34.75 else "left",va="bottom")
ax.set_xlim(24,62); ax.set_ylim(-0.7,len(rows)+0.4); ax.xaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True); ax.set_xlabel("€ per share"); ax.tick_params(axis="y",length=0)
ax.set_title("Rexel valuation summary – football field (€ per share, 30 Oct 2026)")
src(ax,"Source: team model (Comps, DCF, LBO sheets), Yahoo Finance targets. Peers: Wesco, Grainger, Ferguson, Watsco, Bunzl (median to average). xSales omitted (not meaningful: margin gap).")
save(fig,"02_football_field")
# ---------- 3. financial performance (3 small multiples)
yrs=["2022","2023","2024","2025","2026e","2027e","2028e"]; sales=[18701.6,19153.4,19285.1,19414.6,20578,21938,22949]
myrs=["2024","2025","2026e","2027e","2028e"]; marg=[5.9,6.0,6.34,6.68,6.97]
q=["Q1'25","Q2'25","Q3'25","Q4'25","Q1'26","Q2'26","Q3'26"]; grp=[1.8,3.0,3.0,3.8,3.4,6.7,11.4]; na=[5.8,7.4,7.9,7.9,8.7,7.8,17.4]; eu=[-3.0,-0.5,-0.5,0.0,0.6,4.4,4.5]
fig,axs=plt.subplots(1,3,figsize=(11,3.4),gridspec_kw=dict(wspace=0.35))
ax=axs[0]; cols=[BLUE if not y.endswith("e") else "#8DA9C4" for y in yrs]; ax.bar(yrs,[s/1000 for s in sales],color=cols,width=0.62)
for i,s in enumerate(sales): ax.text(i,s/1000+0.25,f"{s/1000:.1f}",ha="center",fontsize=7.5,color=INK2)
ax.set_title("Sales (€bn) – actual and consensus"); ax.set_ylim(0,26); ax.yaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis="x",labelsize=7.5)
ax=axs[1]; ax.plot(myrs,marg,color=BLUE,lw=2,marker="o",ms=5,mec=SURF)
for i,m in enumerate(marg): ax.text(i,m+0.09,f"{m:.1f}%",ha="center",fontsize=7.5,color=INK2)
ax.axhline(7,color=ORANGE,lw=1,ls="--"); ax.text(0.0,7.04,"Axelerate 2028 target: >7%",fontsize=7,color=ORANGE,va="bottom",ha="left")
ax.set_title("Adjusted EBITA margin (%)"); ax.set_ylim(5.4,7.5); ax.yaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis="x",labelsize=7.5)
ax=axs[2]; x=np.arange(len(q)); wd=0.27
for k,(v,c,l) in enumerate([(na,ORANGE,"North America"),(grp,BLUE,"Group"),(eu,"#8A8F98","Europe")]): ax.bar(x+(k-1)*wd,v,width=wd,color=c,label=l)
ax.axhline(0,color=GRID,lw=0.8); ax.set_xticks(x); ax.set_xticklabels(q,fontsize=7.5); ax.set_title("Same-day sales growth by quarter (%)"); ax.legend(loc="upper left",ncol=1); ax.yaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True)
ax.text(6,11.4+0.4,"+11.4%",ha="center",fontsize=7.5,color=BLUE,fontweight="bold"); ax.text(6+wd*-1,17.4+0.4,"+17.4%",ha="center",fontsize=7.5,color=ORANGE,fontweight="bold")
src(axs[0],"Source: Rexel FY 2025 and Q3 2026 releases (S1, S4), consensus collected by Rexel 15 Oct 2026 (S2); 2022-23 sales Yahoo Finance (S10). Q1-Q2 2025 growth read from the Q3 2026 presentation chart.")
save(fig,"03_financial_performance")
# ---------- 4. peer benchmark
info=json.load(open("info.json")); fin=json.load(open("financials.json"))
names=["Rexel","Wesco","Grainger","Ferguson","Watsco","Bunzl","Fastenal"]
evm=[];mg=[];ms=[]
for k in names:
    i=info[k]; f=fin[k]; m=1e6
    if k=="Rexel": ev=(34.47*308.5)+2821.9+11.5; eb=1226.8; s=19414.6; m1=1157.8/19414.6
    else:
        px_=i["currentPrice"]/(100 if k=="Bunzl" else 1); mc=px_*i["sharesOutstanding"]/m; ev=mc+(f.get("BS_Net Debt") or 0)/m
        def v(n):
            x=f.get("LTM_"+n)
            if x is None: fy=f.get("FY_"+n) or {}; x=list(fy.values())[0] if fy else None
            return x/m
        eb=v("EBITDA"); s=v("Total Revenue"); m1=v("EBIT")/s
    evm.append(ev/eb); mg.append(eb/s*100); ms.append(m1*100)
fig,axs=plt.subplots(1,2,figsize=(10,3.6),gridspec_kw=dict(wspace=0.3))
order=np.argsort(evm)[::-1]
for ax,vals,t,fmt in [(axs[0],evm,"EV / EBITDA (last twelve months, x)","{:.1f}x"),(axs[1],mg,"EBITDA margin (last twelve months, %)","{:.1f}%")]:
    o=np.argsort(vals)[::-1]; nm=[names[j] for j in o]; vv=[vals[j] for j in o]; cols=[ORANGE if n=="Rexel" else ("#C3C6CC" if n=="Fastenal" else "#4F7CAC") for n in nm]
    ax.barh(range(len(nm))[::-1],vv,color=cols,height=0.6)
    for y,(n,v) in zip(range(len(nm))[::-1],zip(nm,vv)): ax.text(v+max(vals)*0.01,y,fmt.format(v),va="center",fontsize=8,color=INK2)
    ax.set_yticks(range(len(nm))[::-1]); ax.set_yticklabels(nm,fontsize=8.5); ax.tick_params(axis="y",length=0); ax.set_title(t); ax.xaxis.grid(True,color=GRID,lw=0.6); ax.set_axisbelow(True); ax.set_xlim(0,max(vals)*1.18)
src(axs[0],"Source: Yahoo Finance (S10), 30 Oct 2026; Rexel FY 2025 release (S4). US peers report EBITDA after operating-lease cost (comparable to Rexel's EBITDAaL).\nRexel: EBITDA after leases, net debt pro forma the €500m raise. Fastenal (light grey) excluded from the peer averages.")
save(fig,"04_peer_benchmark")
# ---------- 5. DCF sensitivity heatmap
Ks=[7.5,8.0,8.5,9.0,9.5,10.0]; gs=[1.0,1.5,2.0,2.5,3.0]
M=np.array([[48.5,44.3,40.7,37.6,34.8,32.3],[52.1,47.3,43.2,39.7,36.6,33.9],[56.3,50.8,46.1,42.2,38.7,35.7],[61.3,54.9,49.5,45.0,41.1,37.8],[67.4,59.8,53.5,48.3,43.9,40.1]])
fig,ax=plt.subplots(figsize=(6.2,3.2))
from matplotlib.colors import LinearSegmentedColormap
cm=LinearSegmentedColormap.from_list("navy",["#DCE6F2","#4F7CAC","#0B2545"])
im=ax.imshow(M,cmap=cm,aspect="auto")
for i in range(len(gs)):
    for j in range(len(Ks)):
        v=M[i,j]; ax.text(j,i,f"{v:.0f}",ha="center",va="center",fontsize=9,color="white" if v>44 else INK,fontweight="bold" if (i==2 and j==3) else "normal")
ax.set_xticks(range(len(Ks))); ax.set_xticklabels([f"{k:.1f}%" for k in Ks]); ax.set_yticks(range(len(gs))); ax.set_yticklabels([f"{g_:.1f}%" for g_ in gs]); ax.set_xlabel("Cost of capital K"); ax.set_ylabel("Perpetual growth g")
ax.add_patch(plt.Rectangle((2.5,1.5),1,1,fill=False,ec=ORANGE,lw=2.5)); ax.tick_params(length=0)
for s in ax.spines.values(): s.set_visible(False)
ax.set_title("DCF value per share (€) – sensitivity to K and g (base: K 8.9%, g 2.0% → €43)")
src(ax,"Source: team DCF model. Share price €34.47: every cell with K ≤ 9.5% and g ≥ 1.5% is above the price.")
save(fig,"05_dcf_sensitivity")
