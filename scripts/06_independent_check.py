# Independent Python replica of the three models (same inputs) to check the Excel logic and report results
import json, numpy as np
info=json.load(open("info.json")); fin=json.load(open("financials.json"))
price=34.47; sh_out=(296997254+14388490-2890248)/1e6; mcap=price*sh_out; nd=3321.9-500+11.5
print(f"shares out {sh_out:.1f}m  mcap {mcap:,.0f}  net debt+min {nd:,.0f}")
# ---- COMPS
peers=["Wesco","Grainger","Ferguson","Watsco","Bunzl","Fastenal"]
rows=[]
for k in peers:
    i=info[k]; f=fin[k]; m=1e6
    px=i["currentPrice"]/(100 if k=="Bunzl" else 1); sh=i["sharesOutstanding"]/m; mc=px*sh; ndp=(f.get("BS_Net Debt") or 0)/m; ev=mc+ndp
    def v(n):
        x=f.get("LTM_"+n)
        if x is None:
            fy=f.get("FY_"+n) or {}; x=list(fy.values())[0] if fy else None
        return x/m if x else np.nan
    s,eb,ei,ni=v("Total Revenue"),v("EBITDA"),v("EBIT"),v("Net Income"); s1=f["est_rev"]["+1y"]/m; eps1=f["est_eps"]["+1y"]; ni1=sh*eps1
    eb1=s1*eb/s; ei1=s1*ei/s
    rows.append([ev/s,ev/s1,ev/eb,ev/eb1,ev/ei,ev/ei1,mc/ni,mc/ni1]); print(f"{k:9s} mcap {mc:8,.0f} nd {ndp:7,.0f} EV {ev:8,.0f} | xS {ev/s:.2f}/{ev/s1:.2f} xEBITDA {ev/eb:.1f}/{ev/eb1:.1f} xEBIT {ev/ei:.1f}/{ev/ei1:.1f} PE {mc/ni:.1f}/{mc/ni1:.1f}")
A=np.array(rows); avg=A.mean(0); med=np.median(A,0)
agg=[19414.6,20578,1226.8,938/0.68,1157.8,1305,678.5,713]
print("avg multiples",np.round(avg,2)); print("med multiples",np.round(med,2))
for nm,M in (("AVG",avg),("MED",med)):
    eq=[M[j]*agg[j]-nd for j in range(6)]+[M[6]*agg[6],M[7]*agg[7]]
    print(nm,"equity",np.round(eq,0),"per share",np.round(np.array(eq)/sh_out,1))
# ---- DCF
rf=0.5*0.03325+0.5*0.04936; beta=1.2; mrp=0.0501; spread=0.015; i=rf+spread; t=0.27; g=0.02
ke=rf+beta*mrp; K=ke*mcap/(mcap+nd)+i*(1-t)*nd/(mcap+nd)
bu=beta/(1+nd*(1-t)/mcap); r_=rf+bu*mrp; Kmm=r_*(1-nd*t/(mcap+nd))
print(f"rf {rf:.4f} ke {ke:.4f} WACC {K:.4f}  (MM K {Kmm:.4f}, unlevered beta {bu:.2f})")
sales=[19414.6,20578,21938,22949]; ebitda=[1226.8,938/0.68,1154/0.75,1285/0.77]; ebit=[1157.8,1305,1466,1599]
g28=sales[3]/sales[2]-1; grs=[g28-(g28-g)/4*k for k in (1,2,3)]+[g]
for gr in grs: sales.append(sales[-1]*(1+gr))
em=ebitda[3]/sales[3]; das=(ebit[3]-ebitda[3])/sales[3]
for k in range(4,8): ebitda.append(em*sales[k]); ebit.append(ebitda[k]+das*sales[k])
da=[ebit[k]-ebitda[k] for k in range(8)]
capex=[-136.2]+[-0.007*sales[k] for k in (1,2,3)]
c28=capex[3]; cI=-(-da[7])  # 2032 capex = D&A (positive da negative) -> capex = da (negative number)
cI=da[7]
for k in (1,2,3): capex.append(c28-(c28-cI)/4*k)
capex.append(cI)
wcr0=2364.1+2682.0+34.1+893.6-2170.9-11.8-1039.5; ws_=wcr0/sales[0]; wcr=[ws_*s for s in sales]
fcf=[]
for k in range(1,8): fcf.append(ebit[k]*(1-t)-da[k]+capex[k]-(wcr[k]-wcr[k-1]))
per=[(62/365)+k for k in range(7)]
pv=sum(f/(1+K)**p for f,p in zip(fcf,per)); tv=fcf[-1]*(1+g)/(K-g); pvtv=tv/(1+K)**per[-1]; ev=pv+pvtv; eq=ev-nd
print("sales",np.round(sales,0)); print("EBITDAaL",np.round(ebitda,0)); print("EBIT",np.round(ebit,0)); print("capex",np.round(capex,0)); print("FCF",np.round(fcf,0))
print(f"PV FCF {pv:,.0f} TV {tv:,.0f} PV TV {pvtv:,.0f} EV {ev:,.0f} ({pvtv/ev:.0%} TV) equity {eq:,.0f} per share {eq/sh_out:.2f} upside {eq/sh_out/price-1:+.1%}  EV/EBITDAaL26 {ev/ebitda[1]:.1f}x EV/EBITA26 {ev/ebit[1]:.1f}x")
def val(Kx,gx):
    pv=sum(f/(1+Kx)**p for f,p in zip(fcf,per)); return (pv+fcf[-1]*(1+gx)/(Kx-gx)/(1+Kx)**per[-1]-nd)/sh_out
print("sens g\\K", [f"{x:.1%}" for x in (0.075,0.08,0.085,0.09,0.095,0.10)])
for gx in (0.01,0.015,0.02,0.025,0.03): print(f"{gx:.1%}",[round(val(Kx,gx),1) for Kx in (0.075,0.08,0.085,0.09,0.095,0.10)])
# ---- LBO (quick replica)
tx=0.27; sales_l=sales[:4]; g28=sales_l[3]/sales_l[2]-1
grs=[g28-(g28-g)/7*k for k in (1,2,3,4,5)]
for gr in grs: sales_l.append(sales_l[-1]*(1+gr))
ebl=ebitda[:4]+[em*s for s in sales_l[4:]]; ebt_=ebit[:4]+[ebl[k]+das*sales_l[k] for k in range(4,9)]
dal=[ebt_[k]-ebl[k] for k in range(9)]
cap=[-136.2]+[-0.007*sales_l[k] for k in (1,2,3)]; c28=cap[3]; cI=dal[7]
for k in (1,2,3): cap.append(c28-(c28-cI)/4*k)
cap+= [cI,cI]
wcrl=[ws_*s for s in sales_l]
ndl=[2934.2+671.5+19.6-1037.5]; ni=[None]; div=[None]
for k in range(1,9):
    intr=-i*(ndl[k-1]+ndl[k-1])/2 if k==1 else None
    # iterate for average net debt
    ndk=ndl[k-1]
    for _ in range(20):
        intr=-i*(ndl[k-1]+ndk)/2; ebtk=ebt_[k]+intr; nik=ebtk*(1-tx); cf=nik-nik+(-dal[k])+cap[k]-(wcrl[k]-wcrl[k-1]); ndk=ndl[k-1]-cf
    ndl.append(ndk); ni.append(nik); div.append(nik)
print("LBO Rexel NI",np.round(ni[1:],0)); print("LBO Rexel net debt",np.round(ndl,0))
lev=4.5; debt=lev*ebl[1]; A=0.5*debt; B=0.35*debt; C=0.15*debt; eur=0.022; rA=eur+0.0275+0.005; rB=eur+0.0325+0.005; rC=eur+0.06+0.005
outA=A; cash=0; hni=[]
for k in range(1,4):  # 2027,2028,2029
    yr=k+1
    amort=A/7; intA=rA*(outA+(outA-amort))/2; outA-=amort; intB=rB*B; intC=rC*C; intr=intA+intB+intC
    ebtk=ebt_[yr]-i*(ndl[yr-1]+ndl[yr])/2; taxR=tx*ebtk
    grp_tax=tx*(ebtk-intr); divr=div[yr]; h=divr-tx*0.05*divr+taxR-intr-grp_tax; hni.append(h); cash+=h-amort
xebit=(mcap+ndl[1])/ebt_[1]; ev_exit=xebit*ebt_[4]; cons_nd=ndl[4]+outA+B+C-cash; eq_exit=ev_exit-cons_nd
entry=eq_exit/1.2**3; bid=entry+debt
print(f"acq debt {debt:,.0f} (A {A:,.0f} B {B:,.0f} C {C:,.0f}) rates {rA:.2%}/{rB:.2%}/{rC:.2%}  xEBIT entry {xebit:.1f}x  holding NI {np.round(hni,0)} cash end-29 {cash:,.0f}")
print(f"EV exit {ev_exit:,.0f} cons ND {cons_nd:,.0f} equity exit {eq_exit:,.0f} -> fund entry {entry:,.0f} + debt = bid {bid:,.0f} = {bid/sh_out:.2f}/share ({bid/sh_out/price-1:+.1%}) MoM {eq_exit/entry:.2f}x entry ND/EBITDAaL {(ndl[1]+debt)/ebl[1]:.1f}x")
for lev in (4.0,5.0,5.5):
    for irr in (0.15,0.20,0.25):
        d=lev*ebl[1]; e=eq_exit  # approx: ignore interest change
        print(f"  lev {lev}x irr {irr:.0%}: bid/share ~{(e/(1+irr)**3+d)/sh_out:.1f} (approx, interest not re-run)")
