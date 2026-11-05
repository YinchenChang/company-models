# Siliconomics–Wattonomics 部署量差距診斷（2027–2029）：openpyxl 寫公式，LibreOffice 重算
# 用法：python3 siliconomics/tools/gap_diag.py <輸出.xlsx>，再以 LibreOffice headless 重算
import sys, openpyxl
from openpyxl.styles import Font, PatternFill
BLUE=Font(color="0000FF"); BOLD=Font(bold=True); HDR=PatternFill("solid", fgColor="DDEBF7"); KEY=PatternFill("solid", fgColor="FFF2CC")
wb=openpyxl.Workbook()
rd=wb.active; rd.title="README"
for i,t in enumerate([
 "Siliconomics 與 Wattonomics 部署量差距診斷（2027–2029；2026-11-05；只診斷，不改任何模型）",
 "讀取：Siliconomics v0.16.1（R4 橋接 r7：siliconomics/docs/reports/r4_bridge/20261105_R4對帳橋接_Siliconomics側_r7.xlsx）；Wattonomics v0.3（data/wattonomics_snapshot_v0.3.json）；Tokenomics v21.7（每 GW 顆數、每顆 IT capex）。",
 "系統定義：比較對象＝新增上線 IT GW（毛額）。Siliconomics 側＝R4 b8（出貨年成本等值 GW → 實體 → 上線年）；Wattonomics 側＝WIF_MktGW_Add（基準＝公告時程＋已確認延遲）。比值＝Wattonomics ÷ Siliconomics；容許帶 ±15%。",
 "頁面：In（輸入、標記、低／基準／高）→ Bridge（逐項調整 → 調整後比值）→ NV（2027 NVIDIA 出貨 GW 三方對照）→ Sens（一次動一項）。",
 "標記：Verified／Interested-party（寫明是誰）／Analogy／Assumed／Derived；Analogy、Assumed 給低／高。",
 "正值＝使該側 GW 變大。所有計算為 Excel 公式；LibreOffice headless 重算後存檔。"],1):
    rd.cell(i,1,t).font=Font(bold=(i==1))
rd.column_dimensions["A"].width=150
IN=wb.create_sheet("In")
IN["A1"]="In — 輸入（藍字）。年欄：D＝2027、E＝2028、F＝2029；年度無關的輸入只填 D。低／高見 G、H（只用於 2027 或標明的年度）"
for j,h in enumerate(["鍵","項目","單位","2027／值","2028","2029","低","高","標記","來源／說明"],1):
    c=IN.cell(3,j,h); c.font=BOLD; c.fill=HDR
R={}; r=4
def inp(k,lab,u,vals,lo=None,hi=None,tag="",src=""):
    global r
    IN.cell(r,1,k); IN.cell(r,2,lab); IN.cell(r,3,u)
    for j,v in enumerate(vals):
        if v is not None: c=IN.cell(r,4+j,v); c.font=BLUE
    if lo is not None: IN.cell(r,7,lo).font=BLUE
    if hi is not None: IN.cell(r,8,hi).font=BLUE
    IN.cell(r,9,tag); IN.cell(r,10,src); R[k]=r; r+=1
def sec(t):
    global r
    IN.cell(r,1,t).font=BOLD; r+=1
sec("A. 兩側現值")
inp("s_on","Siliconomics 推得新增上線 IT GW（R4 b8）","GW",[36.173,46.499,51.906],tag="Derived",src="R4 橋接 r7 Bridge b8（Siliconomics v0.16.1）")
inp("w_on","Wattonomics 年新增 IT GW（毛額，基準）","GW",[26.597,29.780,27.970],tag="Derived",src="Wattonomics v0.3 WIF_MktGW_Add_2027–2029")
inp("w_site","Wattonomics SiteOnly（只有站點證據）","GW",[13.57,17.55,6.47],tag="Derived",src="WIF_MktGW_Add_A_D0_SiteOnly；2029 站點證據明顯不足")
inp("lagk","出貨 → 上線：延遲＋庫存係數（lag 0.375＋inv_k 0.10）","x",[0.475],0.35,0.7,"Assumed","R4 橋接 In lag、inv_k 基準（低＝3 個月＋5 週；高＝6 個月＋20%）")
sec("B. Siliconomics 側調整")
inp("s1","① NVIDIA 世代改實體 GW（出貨年）：Σ 顆數 ÷ 該世代每 GW 顆數 − Σ 顆數 × 每顆 IT capex ÷ VR200 每 GW IT capex","GW",[0.19,-2.62,-3.28],tag="Derived",
    src="Deploy 頁 2027–29 以 VR200 每 GW IT capex（37.59）當分母，Rubin Ultra（每 GW 50.09）被高估 33%、Blackwell 被低估；R4 ① 只換算非 NVIDIA。顆數 SIF_UnitsPlat_*×SIF_UnitsCoef_Base；每 GW 顆數 TK IF_GPUsPerGW（GB300 487,008、VR200 291,744、RU 194,760）")
inp("cap_on","供給上限：可支撐新增上線 GW（SIF_CapOnlineGW_Base）","GW",[33.42,46.50,51.91],tag="Derived",src="Siliconomics v0.16 Cap 頁；2027 由 HBM 決定（SIF_HBMUtil 1.08），2028–29 未觸頂（＝基準）。三層校準原則第 2 條：供給是上限")
inp("ru_dup","Rubin Ultra 與 Rubin 可能重複計算（2027，出貨年 GW）","GW",[-1.54],0,-1.54,"Interested-party",
    "Morgan Stanley 2026-07：『Rubin 及 Rubin Ultra 2027 出貨接近 700 萬顆、NVL72 約 9 萬櫃』；本庫 rb 7.0M＋ru 0.3M＝7.3M。0.3M ÷ 194,760／GW＝1.54 GW；低＝不重複")
sec("C. Wattonomics 側調整")
inp("w_cov","Wattonomics 覆蓋不足（七家資本支出高於個股模型＋非七家買方；全市場上線 GW）","GW",[4.8,None,None],1.8,8.3,"Derived",
    "算力收支 11-03《Wattonomics 2027 時差來源診斷》：3.50 NV GW → 全市場約 4.8（1.8–8.3），信心中低；2028–29 未估")
sec("D. 2027 NVIDIA 出貨 GW 三方對照（NV 頁）")
inp("s_nv","Siliconomics NVIDIA 出貨 GW（成本等值；R4 n1）","GW",[30.27],tag="Derived",src="R4 r7 n1")
inp("w_nv","Wattonomics NVIDIA 出貨 GW","GW",[16.03],tag="Derived",src="WIF_Recon_NV_NVGW_2027（逐買方 NVIDIA GW 份額 55%）")
inp("w_ship","Wattonomics 全市場出貨 GW","GW",[29.12],tag="Derived",src="WIF_Recon_NV_ShipGW_2027")
inp("nv27","NVIDIA 資料中心營收 FY2028（約 CY2027）","$B",[633.7],623,645,"Interested-party（NVIDIA 指引推得）","NVIDIA FY27 Q2 法說 FY2028 營收約 +70% × 資料中心 92.5%（R4 In nv_dc_27）")
inp("nv26","NVIDIA 資料中心營收 FY2027 錨點（約 CY2026）","$B",[375.46],tag="Verified／Assumed",src="Q1、Q2 實績 75.2／89.0＋Q3 指引中點＋Q4 假設")
inp("cvr","每 IT GW NVIDIA 營收（VR200，營收口徑）","$B/GW",[26.13],19.67,29.30,"Derived","chip_bridge v0.2.3 Bridge D48／D49／D50（低＝殘差與交換匣其他不算 NVIDIA）")
inp("cw","Wattonomics 用的每 IT GW NVIDIA 內容（2027 世代加權，混合口徑）","$B/GW",[22.29],tag="Derived",src="WIF_Recon_NV_ContentPerGW_2027（R7 改營收口徑前）")
for col,w in zip("ABCDEFGHIJ",[10,70,8,10,10,10,9,9,22,110]): IN.column_dimensions[col].width=w
def v(k,col="D"): return f"In!${col}${R[k]}"
BR=wb.create_sheet("Bridge")
BR["A1"]="Bridge — 新增上線 IT GW：兩側逐項調整（正值＝該側變大）。D–F＝2027–2029"
for j,h in enumerate(["列","項目","說明","2027","2028","2029"],1):
    c=BR.cell(3,j,h); c.font=BOLD; c.fill=HDR
B={}; br=4
YC=["D","E","F"]
def row(k,lab,note,f,bold=False,fmt="0.00"):
    global br
    BR.cell(br,1,k); BR.cell(br,2,lab).font=Font(bold=bold); BR.cell(br,3,note)
    for i,c in enumerate(YC):
        x=f(i,c)
        if x is not None:
            cc=BR.cell(br,4+i,x); cc.number_format=fmt
            if bold: cc.fill=KEY
    B[k]=br; br+=1
row("s0","Siliconomics 推得上線 GW（現值）","R4 b8",lambda i,c:f"={v('s_on',c)}",True)
row("a1","① NVIDIA 世代改實體 GW（上線年）","出貨年調整 − 延遲係數 ×（當年 − 前一年調整）",lambda i,c: f"={v('s1',c)}-{v('lagk')}*({v('s1',c)}-{('0' if i==0 else v('s1',YC[i-1]))})")
row("a2","② Rubin Ultra 可能重複（上線年）","2027 出貨 −1.54 GW；2028 因 2027 出貨減少，上線反增",lambda i,c: f"={v('ru_dup')}*(1-{v('lagk')})" if i==0 else (f"=-{v('ru_dup')}*{v('lagk')}" if i==1 else "=0"))
row("s1x","Siliconomics 調整後（需求面）","s0＋①＋②",lambda i,c: f"={c}{B['s0']}+{c}{B['a1']}+{c}{B['a2']}",True)
row("a3","③ 供給上限截頂","MIN(調整後, 可支撐上線 GW) − 調整後",lambda i,c: f"=MIN({c}{B['s1x']},{v('cap_on',c)})-{c}{B['s1x']}")
row("s2","Siliconomics 調整後（截頂）","",lambda i,c: f"={c}{B['s1x']}+{c}{B['a3']}",True)
br+=1
row("w0","Wattonomics 上線 GW（現值）","WIF_MktGW_Add",lambda i,c: f"={v('w_on',c)}",True)
row("b1","④ Wattonomics 覆蓋不足（算力收支 11-03 估）","2028–29 未估，暫 0",lambda i,c: f"=N({v('w_cov',c)})")
row("w2","Wattonomics 調整後","",lambda i,c: f"={c}{B['w0']}+{c}{B['b1']}",True)
br+=1
row("r0","比值（現值）：Wattonomics ÷ Siliconomics","",lambda i,c: f"={c}{B['w0']}/{c}{B['s0']}",True,"0.000")
row("r1","比值：只做 Siliconomics 側調整","",lambda i,c: f"={c}{B['w0']}/{c}{B['s2']}",False,"0.000")
row("r2","比值：兩側都調整","",lambda i,c: f"={c}{B['w2']}/{c}{B['s2']}",True,"0.000")
row("st","狀態（±15%）","",lambda i,c: f'=IF(ABS({c}{B["r2"]}-1)<=0.15,"一致","待診斷")',True,"@")
row("gap","剩餘差距（Siliconomics − Wattonomics，調整後）","",lambda i,c: f"={c}{B['s2']}-{c}{B['w2']}",False)
row("gsh","剩餘差距 ÷ 原差距","",lambda i,c: f"={c}{B['gap']}/({c}{B['s0']}-{c}{B['w0']})",False,"0%")
row("site","對照：Wattonomics SiteOnly ÷ 基準","2029 只有 23% 有站點證據",lambda i,c: f"={v('w_site',c)}/{v('w_on',c)}",False,"0%")
for col,w in zip("ABCDEF",[6,44,52,10,10,10]): BR.column_dimensions[col].width=w
NV=wb.create_sheet("NV")
NV["A1"]="NV — 2027 NVIDIA 出貨 GW：兩側 vs 由公司營收反推（出貨年；營收口徑每 GW 內容低／基準／高）"
for j,h in enumerate(["項目","低","基準","高","說明"],1):
    c=NV.cell(3,j,h); c.font=BOLD; c.fill=HDR
rows=[
 ("Siliconomics NVIDIA 出貨 GW",f"={v('s_nv')}",f"={v('s_nv')}",f"={v('s_nv')}","R4 n1"),
 ("Wattonomics NVIDIA 出貨 GW",f"={v('w_nv')}",f"={v('w_nv')}",f"={v('w_nv')}","WIF_Recon_NV_NVGW_2027"),
 ("由 FY2028 指引反推的 NVIDIA GW（指引 ÷ 每 GW 營收）",f"={v('nv27','G')}/{v('cvr','H')}",f"={v('nv27')}/{v('cvr')}",f"={v('nv27','H')}/{v('cvr','G')}","低＝指引低／內容高；高＝指引高／內容低"),
 ("由 2026 營收持平反推的下限 GW（2026 錨點 ÷ 每 GW 營收）",f"={v('nv26')}/{v('cvr','H')}",f"={v('nv26')}/{v('cvr')}",f"={v('nv26')}/{v('cvr','G')}","2027 營收不成長時的 NVIDIA GW"),
 ("Wattonomics 隱含 NVIDIA 2027 營收（GW × 營收口徑內容）",f"={v('w_nv')}*{v('cvr','G')}",f"={v('w_nv')}*{v('cvr')}",f"={v('w_nv')}*{v('cvr','H')}","$B"),
 ("　相對 2026 成長",f"=B8/{v('nv26')}-1",f"=C8/{v('nv26')}-1",f"=D8/{v('nv26')}-1","指引為 +69%"),
 ("Siliconomics 隱含 NVIDIA 2027 營收",f"={v('s_nv')}*{v('cvr','G')}",f"={v('s_nv')}*{v('cvr')}",f"={v('s_nv')}*{v('cvr','H')}","$B"),
 ("　相對 2026 成長",f"=B10/{v('nv26')}-1",f"=C10/{v('nv26')}-1",f"=D10/{v('nv26')}-1",""),
 ("Wattonomics NVIDIA 占出貨 GW",f"={v('w_nv')}/{v('w_ship')}",f"={v('w_nv')}/{v('w_ship')}",f"={v('w_nv')}/{v('w_ship')}","Siliconomics 為 30.27 ÷ 41.65＝73%"),
]
for i,(a,b,c,d,e) in enumerate(rows,4):
    NV.cell(i,1,a); NV.cell(i,5,e)
    for j,f in enumerate((b,c,d),2):
        x=NV.cell(i,j,f); x.number_format="0.0%" if "成長" in a or "占" in a else "0.00"
NV.column_dimensions["A"].width=58; NV.column_dimensions["E"].width=60
for col in "BCD": NV.column_dimensions[col].width=11
SE=wb.create_sheet("Sens")
SE["A1"]="Sens — 2027 調整後比值（兩側都調整），一次動一項"
for j,h in enumerate(["情況","2027 Siliconomics 調整後","2027 Wattonomics 調整後","比值"],1):
    c=SE.cell(3,j,h); c.font=BOLD; c.fill=HDR
s1_27=lambda lag: f"({v('s1')}*(1-{lag}))"
base_s=lambda lag="In!$D$%d"%R['lagk'],ru=v('ru_dup'): f"MIN({v('s_on')}+{s1_27(lag)}+{ru}*(1-{lag}),{v('cap_on')})"
cases=[("基準",base_s(),f"{v('w_on')}+{v('w_cov')}"),
 ("Wattonomics 覆蓋不足 低（1.8）",base_s(),f"{v('w_on')}+{v('w_cov','G')}"),
 ("Wattonomics 覆蓋不足 高（8.3）",base_s(),f"{v('w_on')}+{v('w_cov','H')}"),
 ("Rubin Ultra 不重複",base_s(ru="0"),f"{v('w_on')}+{v('w_cov')}"),
 ("不截頂（不採用供給上限）",f"{v('s_on')}+{s1_27(v('lagk'))}+{v('ru_dup')}*(1-{v('lagk')})",f"{v('w_on')}+{v('w_cov')}"),
 ("延遲＋庫存 低",base_s(lag=v('lagk','G')),f"{v('w_on')}+{v('w_cov')}"),
 ("延遲＋庫存 高",base_s(lag=v('lagk','H')),f"{v('w_on')}+{v('w_cov')}")]
for i,(n,s,w) in enumerate(cases,4):
    SE.cell(i,1,n); SE.cell(i,2,"="+s).number_format="0.00"; SE.cell(i,3,"="+w).number_format="0.00"; SE.cell(i,4,f"=C{i}/B{i}").number_format="0.000"
SE.cell(12,1,"註：延遲係數只影響 Siliconomics 的出貨→上線換算；供給上限值本身由 Siliconomics 以基準延遲算出，延遲低高時上限未重算（近似）")
SE.column_dimensions["A"].width=40
for col in "BCD": SE.column_dimensions[col].width=22
wb.save(sys.argv[1]); print(R)
