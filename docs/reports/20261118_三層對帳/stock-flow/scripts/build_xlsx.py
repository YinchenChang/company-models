# 分階段存量流量原型（stockflow，2026-11-18）。只產生試算表，不改任何模型。
# 用法：python3 build_xlsx.py <輸出.xlsx>；之後以 LibreOffice 重算。
# 資料來源：company-models main 4641dc9d（Siliconomics v0.20、Wattonomics v0.3.6、橋接表 v0.3.6r1、各個股成品 v20261116）。
# NVIDIA FY2027 第三季實績（107.6）沒有使用；NVIDIA 只用到 FY2027 Q2 實績（Q3 指引中點只在 LeadLag 頁作參照、標 Interested-party）。
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1]
wb = openpyxl.Workbook()
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); RED = PatternFill("solid", fgColor="F8CBAD")
WRAP = Alignment(wrap_text=True, vertical="top")
YRS = [2025, 2026, 2027, 2028, 2029]
YC = {y: CL(4 + i) for i, y in enumerate(YRS)}  # D..H


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def hdr(ws, r, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR


# =============================== README ===============================
RM = wb.active; RM.title = "README"
readme = [
    "分階段存量流量原型（出貨 → 交付／上架 → 通電；stockflow 設計代理，2026-11-18）。只是試算，不改任何模型、不進任何基準。",
    "單位：實體 GW（IT 電力）；金額：十億美元（US$B），台幣標明 NT$B。年度步長，2025 年底為起點（與橋接表 Stock 頁同一起點 stk_2025）。",
    "頁面：In（輸入，藍字，附等級與來源）→ Flow（三階段存量帳、內生時差、與單一係數並列）→ Check2027（2027 年底 25 GW 放在哪裡才說得通）→ Map（現行係數對應）→ Stages（階段表）→ Stock4（個股類別 × 認列階段）→ LeadLag（季度領先落後小檢驗）。",
    "恆等式：S1（已認列未交付，通路＝ODM／OEM 與在途）期末＝期初＋晶片商出貨 − 買方收貨；S2（已交付未通電，買方端）期末＝期初＋買方收貨 − 上線。S1＋S2＝橋接表 Stock 頁 B 節。",
    "內生時差＝期末存量 ÷ 當年流出 × 12（月），不是外生係數。現行做法（Wattonomics lag_on 4.5 月＋inv_on 10%，合計約 5.7 月）只出現在 S2；S1 現行假設為 0（橋接表把 Siliconomics 出貨直接對上線）。",
    "硬性限制：NVIDIA FY2027 第三季實績未使用；NVIDIA 存貨只用到 FY2026 年底（模型已入）；廣達存貨用到 3Q26（廣達模型已入 3Q26 財報期）。",
    "主要發現（基準）：2027 年底 S1＋S2＝25.3 GW；需求側自身一致的買方端 S2 只有約 6.8 GW，剩約 18.5 GW 要放在通路 S1。依廣達 3Q26 存貨天數，通路最多放約 8.8 GW；放 18.5 GW 需要全市場 ODM 的 AI 存貨約 5,060 億美元，是 3Q26 推估值（約 1,370 億）的 3.7 倍。兩側矛盾的所在＝晶片商出貨與買方收貨之間（ODM／通路這一段），不是時差。",
]
for i, t in enumerate(readme, 1):
    c = RM.cell(i, 1, t); c.alignment = WRAP
RM.column_dimensions["A"].width = 160

# =============================== In ===============================
IN = wb.create_sheet("In")
IN["A1"] = "In — 輸入（藍字）。A 節＝年度數列（D–H＝2025–2029）；B 節＝單值（I–K＝低／基準／高）。標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給三點"
IN["A1"].font = BOLD
hdr(IN, 3, ["鍵", "項目", "單位"] + [str(y) for y in YRS] + ["低", "基準", "高", "標記", "來源"])
r = 4
IN.cell(r, 1, "A. 年度數列").font = BOLD; r += 1
SER = {}


def ser(key, label, unit, vals, tag, src):
    global r
    IN.cell(r, 1, key); IN.cell(r, 2, label); IN.cell(r, 3, unit)
    for y, v in zip(YRS, vals):
        if v is not None:
            c = IN.cell(r, 4 + YRS.index(y), v); c.font = BLUE
    IN.cell(r, 12, tag); IN.cell(r, 13, src)
    SER[key] = r
    r += 1


ser("ship", "晶片商出貨（認列營收）實體 GW：供給側", "GW", [16.137, 24.816, 43.678, 44.106, 44.523], "Derived",
    "2025＝橋接表 In ship_2025（R4 橋接：NVIDIA FY2026 資料中心營收 ÷ 每 GW 內容＋非 NVIDIA）；2026–29＝Siliconomics v0.20 SIF_DeployGWPhys_Base（Deploy 頁實體部署 GW）")
ser("pay", "買方收貨（付款年）實體 GW：需求側", "GW", [11.867, 22.465, 28.337, 29.434, 29.711], "Derived",
    "2025–27＝Wattonomics v0.3.6 WIF_Recon_NV_ShipGW（七家與 xAI 付款年，其餘由上線年鏡像反推）；2028–29 無名稱，取橋接表 Bridge s2（平滑法同式鏡像：(1−k)·上線ₜ＋k·上線ₜ₊₁）")
ser("on", "上線（通電）實體 GW：需求側", "GW", [9.049, 19.294, 25.501, 29.311, 29.569], "Derived", "Wattonomics v0.3.6 WIF_MktGW_Add（主路徑毛額；七家付款年經 lag_on、inv_on 轉上線年）")
ser("bstk", "對照：橋接表 Stock 頁 B 節期末存量（基準）", "GW", [1.614, 7.136, 25.313, 40.108, 55.061], "Derived", "siliconomics/docs/reports/20261118_協調橋接_v0.3.6r1.xlsx Stock!b_end_b（只作守恆核對）")
ser("odmgw", "ODM 整櫃價值（每 GW）", "US$B/GW", [27.121, 27.121, 27.351, 27.351, 27.351], "Derived", "Siliconomics v0.20 Rack 頁 A 節（Tokenomics 每 GW 架數 × ODM 每櫃價格）；2025 用 2026 值")
ser("itpg", "每 GW IT 資本支出", "US$B/GW", [33.414, 33.414, 37.586, 37.586, 37.586], "Derived", "Siliconomics v0.20 Deploy 頁（Tokenomics IF_CapexIT；2026 GB300／VR200 混合，2027 起 VR200）；2025 用 2026 值")
ser("wspend", "Wattonomics IT 支出合計（上線年口徑）", "US$B", [233.578, 606.863, 886.068, 1083.877, 1092.185], "Derived", "Wattonomics v0.3.6 WIF_Spend_it_sum（上線年口徑，W2_B 報告第 4 節）")
ser("onsite", "對照：只計站點證據的上線 GW（SiteOnly）", "GW", [9.049, 11.556, 12.228, 16.482, 6.781], "Derived", "Wattonomics v0.3.6 WIF_MktGW_Add_A_D0_SiteOnly（站點證據路徑）")
for k, rr in SER.items():
    name(f"x_{k}", f"In!$D${rr}:$H${rr}")
r += 1
IN.cell(r, 1, "B. 單值").font = BOLD; r += 1
SC = {}


def sc(key, label, unit, lo, base, hi, tag, src):
    global r
    IN.cell(r, 1, key); IN.cell(r, 2, label); IN.cell(r, 3, unit)
    for j, v in zip((9, 10, 11), (lo, base, hi)):
        if v is not None:
            c = IN.cell(r, j, v); c.font = BLUE
    IN.cell(r, 12, tag); IN.cell(r, 13, src)
    name(key, f"In!$J${r}")
    if lo is not None: name(key + "_lo", f"In!$I${r}")
    if hi is not None: name(key + "_hi", f"In!$K${r}")
    SC[key] = r
    r += 1


sc("stk25", "2025 年底已出貨未上線存量（S1＋S2）", "GW", "=inv_on_lo*In!D%d" % SER["ship"], "=inv_on*In!D%d" % SER["ship"], "=inv_on_hi*In!D%d" % SER["ship"], "Assumed", "同橋接表 stk_2025＝inv_on × ship_2025（低 5%／基準 10%／高 20%）；無外部實績")
sc("s1sh", "2025 年底存量中屬通路（S1）的比例", "比", 0.3, 0.5, 0.7, "Assumed", "無依據；只影響起點分配，2027 年 S1、S2 各差 ±0.3 GW 以內")
sc("lag_on", "Wattonomics 付款到上線延遲", "月", 3, 4.5, 6, "Assumed", "Wattonomics v0.3.6 In lag_on（權威層，規則 J3）")
sc("inv_on", "Wattonomics 未上線庫存 ÷ 年出貨", "比", 0.05, 0.10, 0.20, "Assumed", "Wattonomics v0.3.6 In inv_on")
sc("lag_bldg", "Wattonomics 建物資本支出提前月數", "月", 12, 18, 24, "Assumed", "Wattonomics v0.3.6 In lag_bldg（工作單 C1：12–24 月）")
sc("qinv3", "廣達存貨（2026-09-30）", "NT$B", None, 856.1, None, "Interested-party", "優分析 2026（https://uanalyze.com.tw/articles/8736859074）：3Q26 底 8,561 億、季增 659 億；公開資訊觀測站原文未逐字核對（第二來源待補）")
sc("qinv2", "廣達存貨（2026-06-30）", "NT$B", None, "=qinv3-65.9", None, "Derived", "上列 − 季增 659 億")
sc("qcogs3", "廣達營業成本 3Q26", "NT$B", None, "=1115.54-53.02", None, "Derived", "quanta/dist 廣達 v0.4.10 Q 頁：3Q26 營收 1,115.54 − 毛利 53.02（MOPS t164sb04）")
sc("qcogs2", "廣達營業成本 2Q26（＝1H26 ÷ 2）", "NT$B", None, "=(1845.81-90.78)/2", None, "Derived", "同上 1H26 營收 1,845.81 − 毛利 90.78，平均到兩季（近似）")
sc("qai", "廣達存貨中 AI 伺服器占比", "比", 0.5, "=2719.19/3998.81", 0.85, "Assumed", "基準＝廣達模型 FY2026 AI 伺服器營收 ÷ 營收（Derived 0.68）套用到存貨；低高為判斷")
sc("qsh", "廣達 AI 伺服器組裝份額 2026", "比", None, 0.134, None, "Derived", "Siliconomics v0.20 ShareTbl（asm_qush_26，2026 公司模型校準；分母＝AI 伺服器組裝市場〔ODM 價〕）")
sc("fx", "匯率 NT$／US$", "NT$/US$", 30.5, 32, 33, "Assumed", "Siliconomics v0.20 In fx_twd")
sc("t1lo", "通路停留時間 低（組裝＋測試＋出貨）", "月", None, 1.0, None, "Assumed", "判斷：整櫃組裝測試約 4 週；只作下界")
sc("meta_cip", "Meta 在建工程（2026-09-30）", "US$B", None, 89.499, None, "Verified", "Meta 10-Q 2026Q3 不動產與設備附註（meta/data/meta_facts_2026q3.json ppe.cip.20260930）")
sc("meta_cx9", "Meta 2026 前三季資本支出（含融資租賃本金）", "US$B", None, 85.82, None, "Verified", "meta/dist Meta v0.2.2 In capex_9m")
sc("meta_sv0", "Meta 伺服器與網路設備毛額（2025-12-31）", "US$B", None, 98.04, None, "Verified", "Meta 10-Q 2026Q3（ppe.servers.20251231）")
sc("meta_sv1", "Meta 伺服器與網路設備毛額（2026-09-30）", "US$B", None, 133.626, None, "Verified", "Meta 10-Q 2026Q3（ppe.servers.20260930）")
sc("orcl_cip", "Oracle 在建工程（FY27 Q1 末，2026-08-31）", "US$B", None, 48.546, None, "Interested-party", "oracle/dist Oracle v0.5.3 In cip_q1（Oracle FY27 Q1 10-Q）")
sc("orcl_cxq", "Oracle 資本支出 FY27 Q1", "US$B", None, 28.5, None, "Interested-party", "oracle/dist Oracle v0.5.3 In capex_q1")
sc("nv_inv", "NVIDIA 存貨（FY2026 年底，2026-01-25）", "US$B", None, 21.4, None, "Verified", "nvidia/dist NVIDIA v0.3.11 Hist h_inv（10-K FY2026）")
sc("nv_gm", "NVIDIA 毛利率 FY2026", "比", None, 0.711, None, "Derived", "NVIDIA v0.3.11 Q 頁（GAAP 毛利 153.46 ÷ 營收 215.94）")
sc("nv_cpg26", "NVIDIA 每 GW 內容 2026（營收口徑）", "US$B/GW", None, 25.152, None, "Derived", "Wattonomics v0.3.6 WIF_Recon_NV_ContentPerGW_2026")
sc("nv_cpg27", "NVIDIA 每 GW 內容 2027（營收口徑）", "US$B/GW", None, 26.65, None, "Derived", "Wattonomics v0.3.6 WIF_Recon_NV_ContentPerGW_2027")
sc("nv_sil27", "Siliconomics 隱含 NVIDIA 資料中心營收 CY2027", "US$B", None, 743.8, None, "Derived", "交接檔第 5 節第 5 點凍結值（Siliconomics v0.20 Recon nv_dc）")
sc("nv_wt27", "Wattonomics 隱含 NVIDIA 資料中心營收 CY2027", "US$B", None, 408.758, None, "Derived", "Wattonomics v0.3.6 WIF_Recon_NV_ImpliedRev_2027")
sc("crwv_on", "CoreWeave 在役電力（2026-09-30）", "MW", None, 1700, None, "Interested-party", "coreweave/dist v5.0.5 In mw_q3（Q3 2026 新聞稿）")
sc("crwv_ctr", "CoreWeave 已簽約電力（2026-09-30）", "MW", None, 4600, None, "Interested-party", "coreweave/dist v5.0.5 In ctr_mw")
sc("nbis_on", "Nebius 在役電力（2026-09-30，推估）", "MW", 280, 330, 400, "Assumed", "nebius/dist v0.6.3 In mw_q3（公司未揭露，以營收推估）")
sc("nbis_ctr", "Nebius 已簽約電力（2026-11-12）", "MW", None, 4000, None, "Interested-party", "nebius/dist v0.6.3 In ctr_mw（口徑未說明）")
sc("msft_unc", "Microsoft 已簽未起租租賃（2026-06-30）", "US$B", None, 329.1, None, "Interested-party", "二手轉述 Microsoft FY2026 10-K（https://btw.media/en/microsoft-329-1bn-uncommenced-datacentre-leases）；起租 FY2027–FY2033；10-K 原文未讀")
sc("bill_lag", "通電到開始計費（驗收、燒機）", "月", 0.5, 1, 2, "Assumed", "無公開數字；CoreWeave／Nebius 模型 u_bill（計費 ÷ 在役 0.88–0.96）為同一現象的比例寫法")
sc("t0", "CoWoS 完成到 NVIDIA 認列（晶片測試、上板、出貨）", "月", 1.5, 3, 6, "Assumed", "判斷；NVIDIA 存貨天數可作上界檢查（Check2027 第 F 節）")
for col, w in zip("ABCDEFGHIJKLM", (12, 46, 9, 9, 9, 9, 9, 9, 9, 9, 9, 15, 110)):
    IN.column_dimensions[col].width = w

S = lambda k, y: f"In!{YC[y]}{SER[k]}"  # 年度數列儲存格

# =============================== Flow ===============================
F = wb.create_sheet("Flow")
F["A1"] = ("Flow — 三階段存量帳（單位：實體 GW）。S1＝已認列未交付（通路：ODM／OEM 在製與成品、在途）；S2＝已交付未通電（買方端：已收貨、上架、等電、驗收）。"
           "恆等式：期末＝期初＋流入 − 流出；不設上限、不截斷。內生時差＝期末存量 ÷ 流出 × 12。")
F["A1"].font = BOLD
hdr(F, 3, ["列", "項目", "說明"] + [str(y) for y in YRS])
FR = {}; fr = 4


def frow(key, label, note, f, fmt="0.00", bold=False, fill=None):
    global fr
    F.cell(fr, 1, key); F.cell(fr, 2, label); F.cell(fr, 3, note)
    for y in YRS:
        v = f(y)
        if v is not None:
            c = F.cell(fr, 4 + YRS.index(y), v); c.number_format = fmt
            if fill: c.fill = fill
    if bold: F.cell(fr, 2).font = BOLD
    FR[key] = fr; fr += 1


def fsec(t):
    global fr
    F.cell(fr, 1, t).font = BOLD; fr += 1


P = lambda y: CL(3 + YRS.index(y))  # 前一年欄
Cc = lambda y: YC[y]
fsec("A. 流量")
frow("ship", "① 晶片商出貨（認列營收）", "In ship（供給側）", lambda y: f"={S('ship', y)}", bold=True)
frow("pay", "② 買方收貨（付款年）", "In pay（需求側）", lambda y: f"={S('pay', y)}", bold=True)
frow("on", "③ 上線（通電）", "In on（需求側）", lambda y: f"={S('on', y)}", bold=True)
fsec("B. 存量（基準；2025＝起點）")
frow("s1", "S1 通路存量（已認列未交付）＝期初＋① − ②", "2025＝stk25 × s1sh",
     lambda y: "=stk25*s1sh" if y == 2025 else f"={P(y)}{fr}+{Cc(y)}{FR['ship']}-{Cc(y)}{FR['pay']}", bold=True, fill=KEY)
frow("s2", "S2 買方端存量（已交付未通電）＝期初＋② − ③", "2025＝stk25 ×（1 − s1sh）",
     lambda y: "=stk25*(1-s1sh)" if y == 2025 else f"={P(y)}{fr}+{Cc(y)}{FR['pay']}-{Cc(y)}{FR['on']}", bold=True, fill=KEY)
frow("tot", "合計 S1＋S2（已出貨未通電）", "", lambda y: f"={Cc(y)}{FR['s1']}+{Cc(y)}{FR['s2']}", bold=True)
frow("chk", "　守恆核對：合計 − 橋接表 Stock B 節（應為 0）", "", lambda y: f"=ROUND({Cc(y)}{FR['tot']}-{S('bstk', y)},6)", fmt="0.000000")
frow("chk2", "　守恆核對：合計 −（stk25＋累計① − 累計③）（應為 0）", "與逐年遞推獨立",
     lambda y: f"=ROUND({Cc(y)}{FR['tot']}-(stk25+SUM($E{FR['ship']}:{Cc(y)}{FR['ship']})-SUM($E{FR['on']}:{Cc(y)}{FR['on']})),6)" if y > 2025 else "=0", fmt="0.000000")
frow("tot_lo", "合計（stk25 低）", "", lambda y: f"={Cc(y)}{FR['tot']}-stk25+stk25_lo")
frow("tot_hi", "合計（stk25 高）", "", lambda y: f"={Cc(y)}{FR['tot']}-stk25+stk25_hi")
fsec("B2. 起點一致版（敏感度）：2025 年底 S2＝k × ②₂₀₂₅（與單一係數自洽）、S1＝τ1obs ÷ 12 × ①₂₀₂₅（廣達存貨天數）")
frow("s1a", "S1（起點一致版）", "橋接表 stk_2025 只含 inv_on 部分，漏了 lag_on 的在途量；本列補上",
     lambda y: f"=Check2027!$D$7/12*{Cc(y)}{FR['ship']}" if y == 2025 else f"={P(y)}{fr}+{Cc(y)}{FR['ship']}-{Cc(y)}{FR['pay']}")
frow("s2a", "S2（起點一致版）", "", lambda y: f"=(lag_on/12+inv_on)*{Cc(y)}{FR['pay']}" if y == 2025 else f"={P(y)}{fr}+{Cc(y)}{FR['pay']}-{Cc(y)}{FR['on']}")
frow("tota", "合計（起點一致版）", "", lambda y: f"={Cc(y)}{FR['s1a']}+{Cc(y)}{FR['s2a']}", bold=True)
frow("t1a", "τ1（起點一致版，月）", "", lambda y: f"={Cc(y)}{FR['s1a']}/{Cc(y)}{FR['pay']}*12", fmt="0.0")
frow("t2a", "τ2（起點一致版，月）", "≈ 現行 k 5.7 月：需求側自洽", lambda y: f"={Cc(y)}{FR['s2a']}/{Cc(y)}{FR['on']}*12", fmt="0.0")
fsec("C. 內生時差（月）＝期末存量 ÷ 當年流出 × 12")
frow("t1", "τ1 通路停留（S1 ÷ ②）", "現行做法隱含 0（橋接表把出貨直接對上線）", lambda y: f"={Cc(y)}{FR['s1']}/{Cc(y)}{FR['pay']}*12", fmt="0.0", fill=KEY)
frow("t2", "τ2 買方端停留（S2 ÷ ③）", "現行做法隱含約 5.7 月（lag_on＋inv_on × 12）", lambda y: f"={Cc(y)}{FR['s2']}/{Cc(y)}{FR['on']}*12", fmt="0.0", fill=KEY)
frow("tt", "τ 合計（S1＋S2）÷ ③", "", lambda y: f"={Cc(y)}{FR['tot']}/{Cc(y)}{FR['on']}*12", fmt="0.0")
fsec("D. 與現行單一係數並列")
frow("k", "現行 k（月）＝lag_on＋inv_on × 12", "Wattonomics 權威（J3）；全部放在 S2", lambda y: "=lag_on+inv_on*12", fmt="0.0")
frow("kst", "單一係數隱含存量＝k ÷ 12 × ①", "＝橋接表 Stock C 節 c_kb", lambda y: f"={Cc(y)}{FR['k']}/12*{Cc(y)}{FR['ship']}" if y > 2025 else None)
frow("kr", "恆等式合計 ÷ 單一係數隱含", "1＝與單一係數一致；> 1＝單一係數放不下", lambda y: f"={Cc(y)}{FR['tot']}/{Cc(y)}{FR['kst']}" if y > 2025 else None, fmt="0.00")
frow("kimp", "內生 k（年）＝S2 ÷ ③（只看買方端）", "可直接替代 lag_on÷12＋inv_on 的值", lambda y: f"={Cc(y)}{FR['s2']}/{Cc(y)}{FR['on']}", fmt="0.000")
frow("kimp_tot", "內生 k（年）＝（S1＋S2）÷ ③（兩段合計）", "若硬把全部塞進單一係數", lambda y: f"={Cc(y)}{FR['tot']}/{Cc(y)}{FR['on']}", fmt="0.000")
frow("crdk", "S1 不累積時的出貨讀數 ÷ 供給側讀數（類 crd_k）＝② ÷ ①", "Map 頁說明：crd_k 在新架構＝通路可容納後的出貨 ÷ 原始出貨", lambda y: f"={Cc(y)}{FR['pay']}/{Cc(y)}{FR['ship']}", fmt="0.000")
fsec("E. 第四段（只顯示）：通電 → 開始計費")
frow("bill", "計費中存量（已通電、驗收中）＝③ × bill_lag ÷ 12", "Tokenomics 每 GW 營收應乘在計費 GW 上；現行乘在上線 GW", lambda y: f"={Cc(y)}{FR['on']}*bill_lag/12")
fsec("F. 時點換算（只顯示；個股讀取用）：各階段流量在年度帳的近似（領先為正）")
frow("f0", "階段 0 流量（CoWoS 完成、HBM 與晶圓出貨）≈ ①ₜ＋t0 ÷ 12 ×（①ₜ₊₁ − ①ₜ）", "上游供應商認列營收的時點；2029 次年用 2029 值",
     lambda y: (f"={Cc(y)}{FR['ship']}+t0/12*({YC[y+1]}{FR['ship']}-{Cc(y)}{FR['ship']})" if y < 2029 else f"={Cc(y)}{FR['ship']}"))
frow("f2s", "ODM 交付流量（供給側一致版）≈ ①ₜ − τ1obs ÷ 12 ×（①ₜ − ①ₜ₋₁）", "通路停留取 Check2027 觀測 τ1（廣達 3Q26 存貨天數）",
     lambda y: None if y == 2025 else f"={Cc(y)}{FR['ship']}-Check2027!$D$7/12*({Cc(y)}{FR['ship']}-{P(y)}{FR['ship']})")
frow("f2d", "ODM 交付流量（需求側一致版）＝②", "", lambda y: f"={Cc(y)}{FR['pay']}")
F.column_dimensions["A"].width = 8; F.column_dimensions["B"].width = 50; F.column_dimensions["C"].width = 48
for y in YRS: F.column_dimensions[YC[y]].width = 10

# =============================== Check2027 ===============================
K = wb.create_sheet("Check2027")
K["A1"] = "Check2027 — 2027 年底「已出貨未通電」各階段要放多少才說得通，對照實體世界可查的量（存貨、在建工程）。金額 US$B；GW＝實體 GW"
K["A1"].font = BOLD
hdr(K, 3, ["列", "項目", "說明", "值", "低", "高", "等級"])
kr = 4; KR = {}


def krow(key, label, note, v, lo=None, hi=None, tag="Derived", fmt="0.00", fill=None):
    global kr
    K.cell(kr, 1, key); K.cell(kr, 2, label); K.cell(kr, 3, note)
    for j, x in zip((4, 5, 6), (v, lo, hi)):
        if x is not None:
            c = K.cell(kr, j, x); c.number_format = fmt
            if fill and j == 4: c.fill = fill
    K.cell(kr, 7, tag)
    KR[key] = kr; kr += 1


def ksec(t):
    global kr
    K.cell(kr, 1, t).font = BOLD; kr += 1


FC = "F"  # 2027 欄在 Flow
ksec("A. 觀測停留時間（月）")
krow("dio3", "廣達存貨天數（3Q26）＝存貨 ÷ 季營業成本 × 3", "全公司口徑（含非 AI、含尚未上板的原料）→ 當作通路停留的上界", "=qinv3/qcogs3*3", tag="Derived（Interested-party 存貨）", fmt="0.0")
krow("dio2", "廣達存貨天數（2Q26）", "同上", "=qinv2/qcogs2*3", tag="Derived", fmt="0.0")
krow("t1obs", "通路停留 τ1obs（基準＝3Q26；低＝判斷 1.0；高＝2Q26）", "Flow F 節引用本列 D 欄（第 7 列）", f"=D{kr-2}", "=t1lo", f"=D{kr-1}", tag="Derived／Assumed", fmt="0.0", fill=KEY)
assert KR["t1obs"] == 7, KR
krow("meta_t", "Meta 在建工程 ÷ 年化資本支出（月）", "含建物；IT 單獨停留時間更短 → 上界", "=meta_cip/(meta_cx9*4/3)*12", tag="Derived（Verified 原值）", fmt="0.0")
krow("orcl_t", "Oracle 在建工程 ÷ 年化資本支出（月）", "Oracle 機房多為租用 → 較接近 IT 停留", "=orcl_cip/(orcl_cxq*4)*12", tag="Derived（Interested-party 原值）", fmt="0.0")
krow("meta_sv", "對照：Meta 伺服器與網路設備前三季轉入在役（年化）", "毛額增量 ×4/3；與資本支出比較，看資本支出有多少比例當年轉入", "=(meta_sv1-meta_sv0)*4/3", tag="Derived", fmt="0.0")
krow("k_now", "現行單一係數 k（月）", "lag_on＋inv_on × 12（全部放買方端）", "=lag_on+inv_on*12", "=lag_on_lo+inv_on_lo*12", "=lag_on_hi+inv_on_hi*12", tag="Assumed", fmt="0.0")
ksec("B. 2027 年底需要放的量（基準情境）")
krow("tot", "S1＋S2 合計（＝橋接表 Stock B 節）", "", f"=Flow!{FC}{FR['tot']}", f"=Flow!{FC}{FR['tot_lo']}", f"=Flow!{FC}{FR['tot_hi']}", fill=KEY)
krow("s2m", "S2 買方端（需求側自身一致：② − ③ 累積）", "Wattonomics 的錢與上線互相一致的部分", f"=Flow!{FC}{FR['s2']}")
krow("s1r", "S1 通路需要放的量＝合計 − S2", "", f"=D{kr-2}-D{kr-1}", f"=E{kr-2}-D{kr-1}", f"=F{kr-2}-D{kr-1}", fill=KEY)
krow("s1ra", "S1 通路需要放的量（起點一致版）", "Flow B2 節；起點補上 lag_on 在途量後 S1 反而更大", f"=Flow!{FC}{FR['s1a']}")
krow("tota", "S1＋S2 合計（起點一致版）", "比橋接表 B 節多出 lag_on 在途量", f"=Flow!{FC}{FR['tota']}")
ksec("C. 通路（ODM／OEM）放得下多少")
krow("cap", "通路容量＝τ1obs ÷ 12 × 2027 出貨①", "停留時間不變、存貨隨出貨同比例放大（已偏寬）", f"=D{KR['t1obs']}/12*Flow!{FC}{FR['ship']}", f"=E{KR['t1obs']}/12*Flow!{FC}{FR['ship']}", f"=F{KR['t1obs']}/12*Flow!{FC}{FR['ship']}")
krow("obs26", "觀測：全市場 ODM AI 存貨換算 GW（3Q26）＝廣達存貨 ÷ 匯率 × AI 占比 ÷ 份額 ÷ 每 GW ODM 價值",
     "只有廣達一家實測；以份額外推（其他 ODM 存貨天數未查）", "=qinv3/fx*qai/qsh/In!E" + str(SER["odmgw"]), "=qinv3/fx_hi*qai_lo/qsh/In!E" + str(SER["odmgw"]),
     "=qinv3/fx_lo*qai_hi/qsh/In!E" + str(SER["odmgw"]), tag="Derived（廣達 Interested-party、份額 Derived）")
krow("obs26usd", "　同上，金額（US$B）", "", f"=D{kr-1}*In!E{SER['odmgw']}", f"=E{kr-1}*In!E{SER['odmgw']}", f"=F{kr-1}*In!E{SER['odmgw']}", fmt="0")
krow("m26", "　模型 2026 年底 S1（Flow）", "對照觀測值：2026 年模型的通路存量 ≤ 觀測 → 2026 說得通", f"=Flow!E{FR['s1']}")
krow("req_usd", "放 S1 需要的全市場 ODM AI 存貨（US$B）", "S1 × 2027 每 GW ODM 價值", f"=D{KR['s1r']}*In!F{SER['odmgw']}", f"=E{KR['s1r']}*In!F{SER['odmgw']}", f"=F{KR['s1r']}*In!F{SER['odmgw']}", fmt="0", fill=KEY)
krow("mult", "　÷ 3Q26 觀測（倍）", "", f"=D{kr-1}/D{KR['obs26usd']}", fmt="0.0", fill=KEY)
krow("q_req", "　換成廣達存貨（NT$B；份額與 AI 占比不變）", "對照 3Q26 856.1", f"=D{KR['req_usd']}*qsh/qai*fx", fmt="#,##0")
krow("exc", "放不下的量＝S1 需要 − 通路容量（GW）", "> 0：這一段放不下", f"=D{KR['s1r']}-D{KR['cap']}", f"=D{KR['s1r']}-F{KR['cap']}", f"=D{KR['s1r']}-E{KR['cap']}", fill=RED)
krow("verdict", "判定", "", f'=IF(D{kr-1}>0,"通路放不下："&TEXT(D{kr-1},"0.0")&" GW 沒有地方放","通路放得下")', fmt="@")
ksec("D. 若把放不下的量移到買方端（情境 Y）：要成立需要什麼")
krow("y_s2", "買方端需要的 S2＝S2（需求側）＋放不下的量", "", f"=D{KR['s2m']}+D{KR['exc']}")
krow("y_t2", "　隱含買方端停留（月）＝S2 ÷ 2027 上線③ × 12", "對照 Meta 上界與 Oracle（A 節）", f"=D{kr-1}/Flow!{FC}{FR['on']}*12", fmt="0.0")
krow("y_pay", "　需要多出的買方收貨（2026–27 累計 GW）", "＝放不下的量；買方得先付錢才收貨", f"=D{KR['exc']}")
krow("y_usd", "　換成買方 IT 資本支出（US$B）", "× 2027 每 GW IT 資本支出", f"=D{kr-1}*In!F{SER['itpg']}", fmt="0")
krow("y_pct", "　÷ Wattonomics 2026＋2027 IT 支出（上線年口徑，近似）", "需求側資本支出要被低估這麼多", f"=D{kr-1}/(In!E{SER['wspend']}+In!F{SER['wspend']})", fmt="0.0%", fill=KEY)
ksec("E. 用錢核對同一件事（NVIDIA 營收，2027 單年）")
krow("nvgap", "Siliconomics 隱含 − Wattonomics 隱含 NVIDIA 資料中心營收（US$B）", "交接檔凍結值", "=nv_sil27-nv_wt27", fmt="0")
krow("nvgw", "　÷ 2027 NVIDIA 每 GW 內容（GW）", "只 NVIDIA；一年就要累積這麼多", "=D{0}/nv_cpg27".format(kr - 1), fmt="0.0")
ksec("F. 上游存量（認列之前；不在「已出貨未通電」內）")
krow("nvinv", "NVIDIA 存貨換算 GW（FY2026 年底）＝存貨 ÷（每 GW 內容 ×（1 − 毛利率））", "含 HBM、晶圓等原料 → 上界；顯示 t0 的量級", "=nv_inv/(nv_cpg26*(1-nv_gm))", fmt="0.0")
krow("nvinv_m", "　相當於 2026 年出貨的月數", "對照 t0 假設（1.5／3／6 月）", f"=D{kr-1}/In!E{SER['ship']}*12", fmt="0.0")
ksec("G. 新雲：已簽約電力對在役（另一條可觀測存量，屬站點與電力，不是晶片）")
krow("crwv", "CoreWeave 已簽約 ÷ 在役（2026-09-30）", "", "=crwv_ctr/crwv_on", fmt="0.0")
krow("nbis", "Nebius 已簽約 ÷ 在役", "在役為推估", "=nbis_ctr/nbis_on", "=nbis_ctr/nbis_on_hi", "=nbis_ctr/nbis_on_lo", fmt="0.0")
K.column_dimensions["A"].width = 9; K.column_dimensions["B"].width = 62; K.column_dimensions["C"].width = 52
for col in "DEF": K.column_dimensions[col].width = 11
K.column_dimensions["G"].width = 30

# =============================== Map ===============================
M = wb.create_sheet("Map")
M["A1"] = "Map — 現行時差與水準係數對應到新架構的哪一段、能否由新架構推導（公式見各列；數值為 2027 基準）"
M["A1"].font = BOLD
hdr(M, 3, ["係數", "所在層", "現行定義", "新架構中的位置", "推導公式（新架構）", "2027 現值", "2027 新架構推得值（Flow B2 起點一致版）", "可否推導"])
maps = [
    ("lag_on", "Wattonomics", "付款到上線延遲（月），基準 4.5", "S2 的「正常流程」部分（上架、等電、驗收）",
     "lag_on ≈ 12 × S2ₜ ÷ ③ₜ − 12 × inv_on（若 inv_on 另有觀測）；或直接以 τ2 取代 lag_on＋inv_on × 12", "=lag_on", f"=Flow!F{FR['t2a']}-inv_on*12",
     "可（以 τ2 內生）；拆成 lag_on 與 inv_on 兩項需要額外觀測（在建工程分建物與 IT）"),
    ("inv_on", "Wattonomics", "未上線庫存 ÷ 年出貨，基準 10%", "S2 的「超額囤積」部分",
     "inv_on ≈ S2ₜ ÷ ②ₜ − lag_on ÷ 12（S2 期末對當年收貨）", "=inv_on", f"=Flow!F{FR['s2a']}/Flow!F{FR['pay']}-lag_on/12",
     "可，但與 lag_on 不可分；新架構建議合併成單一 τ2（月）"),
    ("k（橋接表）", "橋接表", "lag_on ÷ 12＋inv_on（年）", "S1＋S2 合計的停留時間，但現行只量 S2",
     "k₍內生₎＝（S1＋S2）ₜ ÷ ③ₜ；只看買方端時＝S2ₜ ÷ ③ₜ", "=lag_on/12+inv_on", f"=Flow!F{FR['s2a']}/Flow!F{FR['on']}",
     "可；平滑法的 k 改成逐年內生 kₜ，換算值 x2 改由恆等式反推"),
    ("crd_k", "Siliconomics", "基準 ÷ 供給側讀數（2027＝1、2028＝0.80、2029＝0.71）；水準係數，不是時差", "①（供給側出貨）中最後真的交到買方的比例",
     "crd_kₜ ≈（②ₜ＋ΔS1ₜ容量內）÷ ①ₜ；ΔS1 上限＝τ1obs ÷ 12 ×（①ₜ − ①ₜ₋₁）", 1,
     f"=(Flow!F{FR['pay']}+Check2027!D7/12*(Flow!F{FR['ship']}-Flow!E{FR['ship']}))/Flow!F{FR['ship']}",
     "部分可：新架構把 crd_k 的「為什麼不是 1」講成通路放不下；但數值仍取決於誰的流量對（②或①），不會自己消失"),
    ("lag_bldg", "Wattonomics", "建物資本支出提前月數（12–24，基準 18）", "平行的建物鏈：開工 → 外殼機電完成 → 可通電；在建工程的建物部分",
     "lag_bldg ≈ 12 × 建物在建工程 ÷ 年建物支出（需在建工程分建物與 IT；Meta、Oracle 未拆）", "=lag_bldg", "未拆分（無資料）",
     "暫不可：在建工程未拆建物與 IT；SemiAnalysis 五階段的前三階段屬這條鏈"),
]
for i, row in enumerate(maps):
    for j, v in enumerate(row, 1):
        c = M.cell(4 + i, j, v); c.alignment = WRAP
        if j in (6, 7) and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
for col, w in zip("ABCDEFGH", (12, 12, 30, 30, 50, 10, 12, 44)):
    M.column_dimensions[col].width = w

# =============================== Stages ===============================
ST = wb.create_sheet("Stages")
ST["A1"] = "Stages — 階段定義與可觀測資料。揭露狀態：有揭露＝公司或法定文件直接給；沒揭露＝已查證公司不揭露；沒找到＝本輪未找到（不等於不存在）"
ST["A1"].font = BOLD
hdr(ST, 3, ["序", "階段（流量／存量）", "定義", "單位", "擁有層", "可直接觀測的公開資料", "等級", "頻率", "落後", "揭露狀態"])
stages = [
    ("0", "流量：CoWoS 完成、HBM 出貨", "台積電 CoWoS 封裝完成、記憶體廠出貨 HBM", "千片、EB、封裝顆數", "Siliconomics（Cap、Mkt 頁）",
     "台積電季營收與 HPC 占比；SK 海力士、美光季營收；CoWoS 產能（Interested-party 說法）", "Verified（營收）／Interested-party（產能）", "季", "約 1 個月", "營收有揭露；CoWoS 量與 HBM 金額沒揭露"),
    ("0→1", "存量：NVIDIA 存貨與採購承諾", "已生產、尚未認列營收", "US$B", "Siliconomics（無變數）",
     "NVIDIA 10-Q 存貨（FY2026 年底 21.4）、供應與產能承諾（119 → 279）", "Verified", "季", "約 1 個月", "有揭露（金額，非顆數）"),
    ("1", "流量：NVIDIA 等晶片商認列營收（出貨）", "賣給直接客戶（ODM、OEM、雲端）時認列", "顆數、成本等值 GW、實體 GW", "Siliconomics（Units、Deploy）",
     "NVIDIA 資料中心營收；AMD、Broadcom AI 營收；顆數沒揭露", "Verified（營收）", "季", "約 1 個月", "營收有揭露；顆數沒揭露"),
    ("1→2", "存量 S1：通路（ODM／OEM 在製、成品、在途）", "已認列、尚未交到買方", "GW、US$B", "橋接層（新增）",
     "台廠季報存貨（廣達 3Q26 8,561 億）；月營收（10 日內公布，最高頻）；Dell、美超微季報存貨", "Interested-party（二手）／Verified（MOPS 原文待核）", "季（存貨）／月（營收）", "約 1.5 個月／10 天",
     "廣達有揭露；鴻海、緯創、緯穎本輪沒找；AI 與非 AI 存貨拆分沒揭露"),
    ("2", "流量：ODM 整櫃出貨＝買方收貨（交付）", "整櫃交到買方站點；買方以資本支出入帳", "機櫃、GW、US$B", "橋接層（新增）／Wattonomics（付款年 GW）",
     "ODM 月營收與季報；買方資本支出（現金與應計）", "Verified（營收、資本支出）", "月／季", "10 天／約 1 個月", "總額有揭露；AI 伺服器金額沒揭露（只有占比說法）"),
    ("2→3", "存量 S2：買方端（已收貨、上架、等電、驗收）", "已付款或應付、尚未通電投入使用", "GW、US$B", "Wattonomics（lag_on、inv_on）",
     "買方在建工程（Meta 89.5、Oracle 48.5）；伺服器毛額轉入在役（Meta）；「有晶片沒電」說法", "Verified（Meta）／Interested-party（Oracle 經模型）", "季", "約 1 個月",
     "Meta、Oracle 有揭露（未分建物與 IT）；Microsoft、Alphabet、Amazon 在建工程本輪沒找到；Microsoft 未起租租賃 3,291 億有揭露（二手）"),
    ("3", "流量：通電（上線）", "站點供電、IT 負載開始", "MW（IT）", "Wattonomics（主路徑、站點路徑）",
     "公用事業併網公告、空污許可（Epoch 法）、新雲在役 MW（CoreWeave 1.7 GW）", "Interested-party（公司 KPI）／Verified（許可）", "不定期／季", "數週到一季", "新雲有揭露；超大雲端沒揭露 MW"),
    ("3→4", "存量：已通電未計費（驗收、燒機）", "上線到開始收費", "MW", "Tokenomics（u_bill 類）／個股",
     "CoreWeave、Nebius 已簽約對在役（4.6／1.7 GW；4.0／約 0.33 GW）——屬電力合約，不是晶片", "Interested-party", "季", "約 1 個月", "計費 MW 沒揭露"),
    ("4", "流量：開始服務／計費", "收入開始", "token、US$", "Tokenomics（每 GW 營收）",
     "新雲營收、雲端分部營收、OpenAI／Anthropic 營收說法", "Verified（新雲）／Interested-party（實驗室）", "季", "約 1 個月", "新雲有揭露；超大雲端 AI 營收沒揭露"),
    ("平行", "建物鏈：土地許可 → 動工 → 外殼機電 → 可通電", "SemiAnalysis 五階段的前三段", "MW、US$B", "Wattonomics（lag_bldg、站點清單）",
     "建照、空污許可、衛星影像（Epoch、SemiAnalysis）、公司站點公告", "Verified（許可）／Interested-party", "不定期", "數月", "部分有；逐站系統資料本輪沒找"),
]
for i, row in enumerate(stages):
    for j, v in enumerate(row, 1):
        ST.cell(4 + i, j, v).alignment = WRAP
for col, w in zip("ABCDEFGHIJ", (5, 30, 30, 14, 22, 50, 22, 10, 10, 40)):
    ST.column_dimensions[col].width = w

# =============================== Stock4 ===============================
S4 = wb.create_sheet("Stock4")
S4["A1"] = "Stock4 — 個股類別 × 認列營收的階段（2027 方向性估計；實體 GW；不改任何個股）"
S4["A1"].font = BOLD
hdr(S4, 3, ["類別", "代表個股", "現在讀的部署量", "2027 現讀值（GW）", "應該讀的階段", "2027 應讀值（供給側一致）", "2027 應讀值（需求側一致）", "差異（供給側版 ÷ 現讀 − 1）", "差異（需求側版 ÷ 現讀 − 1）", "時差方向", "說明"])
CE27 = 39.164  # SIF_DeployGW_Base 2027
cats = [
    ("晶片設計（GPU、ASIC）", "NVIDIA、AMD、Broadcom、Marvell", "SIF_DeployGW_Base（成本等值、出貨年）", CE27, "階段 1（認列＝出貨）",
     f"=Flow!F{FR['ship']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "同階段", "階段正確；兩側差距直接落在這類（NVIDIA 2027 隱含營收 743.8 對 408.8）"),
    ("晶圓代工與先進封裝", "台積電、日月光、Amkor", "SIF_UnitsPlat_Total、SIF_UnitsCoef（顆數、出貨年）", CE27, "階段 0（CoWoS 完成，領先約 t0）",
     f"=Flow!F{FR['f0']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "現讀偏晚（2026 被算少、2027 年增被算高）", "2027→2028 出貨持平，所以 2027 水準差很小；差在 2026→2027 年增"),
    ("記憶體（HBM）", "SK 海力士、三星、美光", "markets/memory 需求位元（顆數、出貨年）", CE27, "階段 0（HBM 出貨，領先約 t0）",
     f"=Flow!F{FR['f0']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "同上", "HBM 供給受限，營收傳導最小（stockgap）"),
    ("載板（ABF）", "Ibiden、欣興、南電、景碩", "ABF 指數（顆數、出貨年）", CE27, "階段 0（領先）",
     f"=Flow!F{FR['f0']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "同上", ""),
    ("AI 高速 PCB／CCL", "勝宏、金像電、台光電、台燿", "PGW（部署 GW，出貨年）", CE27, "階段 1→2 之間（主板隨 ODM 組裝）",
     f"=Flow!F{FR['f2s']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "現讀略早", ""),
    ("光模組與網通", "中際旭創、新易盛、Coherent、智邦", "markets/network AI IT 資本支出（成本等值、出貨年）", CE27, "階段 2（隨買方收貨；交換器由買方直接採購）",
     f"=Flow!F{FR['f2s']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "現讀略早", "光模組多直接賣給買方或 ODM，時點接近整櫃交付"),
    ("ODM 組裝", "廣達、鴻海、緯創、緯穎", "SIF_AIITCapex_Base、SIF_AsmMkt_Base（成本等值、出貨年）", CE27, "階段 2（ODM 出貨＝買方收貨），應改實體 GW",
     f"=Flow!F{FR['f2s']}", f"=Flow!F{FR['pay']}", "現讀偏早約 2.4 月；單位應改實體", "月營收是階段 2 最高頻的觀測，可直接裁定兩側"),
    ("電源（機櫃內）", "台達電、光寶", "SIF_PwrMkt_Base（成本等值 GW × 每 kW 內容）", CE27, "階段 2（隨整櫃），應改實體 GW",
     f"=Flow!F{FR['f2s']}", f"=Flow!F{FR['pay']}", "同 ODM", "機房級電力（UPS、變壓器）屬建物鏈，提前 lag_bldg"),
    ("散熱（液冷）", "奇鋐、雙鴻、健策", "SIF_LiqMkt_Base（成本等值 GW × 滲透率）", CE27, "階段 2（冷板、分歧管隨整櫃），應改實體 GW",
     f"=Flow!F{FR['f2s']}", f"=Flow!F{FR['pay']}", "同 ODM", "CDU 等機房設備屬建物鏈"),
    ("商用 GPU 系統（OEM）", "Dell、美超微", "Rack B 節顆數（出貨年）", CE27, "階段 2（OEM 出貨給企業與新雲）",
     f"=Flow!F{FR['f2s']}/Flow!F{FR['ship']}*{CE27}", f"=Flow!F{FR['pay']}/Flow!F{FR['ship']}*{CE27}", "現讀略早", "會計年度錯位另計"),
    ("算力平台與雲端", "CoreWeave、Nebius、Oracle", "Wattonomics 玩家在役 GW（上線年）＋Tokenomics 每 GW", "上線", "階段 4（計費＝上線＋驗收）",
     f"=Flow!F{FR['on']}", f"=Flow!F{FR['on']}", "階段大致正確；少驗收時差", "已讀上線，另有 u_bill；可再加 bill_lag 約 1 月"),
]
for i, row in enumerate(cats):
    rr = 4 + i
    vals = list(row)
    S4.cell(rr, 1, vals[0]); S4.cell(rr, 2, vals[1]); S4.cell(rr, 3, vals[2])
    c = S4.cell(rr, 4, vals[3] if vals[3] != "上線" else f"=Flow!F{FR['on']}"); c.number_format = "0.00"
    S4.cell(rr, 5, vals[4])
    S4.cell(rr, 6, vals[5]).number_format = "0.00"
    S4.cell(rr, 7, vals[6]).number_format = "0.00"
    S4.cell(rr, 8, f"=F{rr}/D{rr}-1").number_format = "0.0%"
    S4.cell(rr, 9, f"=G{rr}/D{rr}-1").number_format = "0.0%"
    S4.cell(rr, 10, vals[7]); S4.cell(rr, 11, vals[8])
    for j in (1, 2, 3, 5, 10, 11): S4.cell(rr, j).alignment = WRAP
S4.cell(4 + len(cats) + 1, 1, "註：ODM、電源、散熱的「應讀值」直接用實體 GW（stockgap 發現這 9 家仍讀成本等值）；其他類別以成本等值 GW 乘相對比例，只呈現時點效果。需求側版＝假設兩側矛盾全部以買方收貨為準（上限型的另一端）。").alignment = WRAP
for col, w in zip("ABCDEFGHIJK", (18, 22, 30, 10, 28, 12, 12, 12, 12, 24, 40)):
    S4.column_dimensions[col].width = w

# =============================== LeadLag ===============================
LL = wb.create_sheet("LeadLag")
LL["A1"] = ("LeadLag — 季度領先／落後小檢驗（只用各個股 Q 頁既有實績；NVIDIA 只用到 FY2027 Q2 實績，Q3 指引中點標 Interested-party 只作參照；NVIDIA Q3 實績未使用）。"
            "做法：各家 3Q26（7–9 月）÷ 1H26 季平均 − 1，對照 NVIDIA 同窗與次窗的同一比值。落後型應接近 NVIDIA 前一窗，領先型應接近次窗。")
LL["A1"].font = BOLD
hdr(LL, 3, ["鍵", "項目", "單位", "1H26 或前期", "3Q26 或後期", "比值 − 1", "標記", "來源"])
ll = [
    ("nv_fq1", "NVIDIA 資料中心 FY27 Q1（2026-02～04）", "US$B", None, 75.2, None, "Verified", "nvidia v0.3.11 Q 頁"),
    ("nv_fq2", "NVIDIA 資料中心 FY27 Q2（2026-05～07）", "US$B", None, 89.0, None, "Verified", "同上"),
    ("nv_fq3g", "NVIDIA 資料中心 FY27 Q3 指引中點推得（2026-08～10）", "US$B", None, 100.42, None, "Interested-party", "同上（指引，不是實績；Q3 實績未使用）"),
    ("nv_a", "NVIDIA 同窗：FQ2 ÷ (FQ1＋FQ2)/2 − 1（約對應 5–7 月對 2–7 月）", "比", "=(E4+E5)/2", "=E5", "=E{r}/D{r}-1", "Derived", ""),
    ("nv_b", "NVIDIA 次窗：FQ3 指引 ÷ (FQ1＋FQ2)/2 − 1（約對應 8–10 月）", "比", "=(E4+E5)/2", "=E6", "=E{r}/D{r}-1", "Derived（含指引）", ""),
    ("tsmc", "台積電 HPC 營收（US$B）：3Q26 ÷ 1H26 季平均", "US$B", "=(35.9*0.61+40.2*0.66)/2", "=46.72*0.68", "=E{r}/D{r}-1", "Verified", "tsmc 成品 In rev26_1q–3q、pf26_*_hpc"),
    ("skh", "SK 海力士記憶體營收（兆韓元）：3Q26 ÷ 1H26 季平均", "KRW T", "=(52.58+79.32)/2", "=104.62", "=E{r}/D{r}-1", "Verified", "skhynix 成品 Q 頁（含 DRAM 價格上漲，非純量）"),
    ("quanta", "廣達營收（NT$B）：3Q26 ÷ 1H26 季平均", "NT$B", "=1845.81/2", "=1115.54", "=E{r}/D{r}-1", "Verified", "quanta 成品 Q 頁（含非 AI）"),
    ("foxconn", "鴻海營收（NT$B）：3Q26 ÷ 1H26 季平均", "NT$B", "=4645.43/2", "=3015.63", "=E{r}/D{r}-1", "Verified", "foxconn 成品 Q 頁（含非 AI）"),
    ("wiwynn", "緯穎營收（NT$B）：3Q26 ÷ 1H26 季平均", "NT$B", "=554.66/2", "=434.62", "=E{r}/D{r}-1", "Verified", "wiwynn 成品 Q 頁（ASIC 為主）"),
    ("delta", "台達電營收（NT$B）：3Q26 ÷ 1H26 季平均", "NT$B", "=342.6/2", "=197.06", "=E{r}/D{r}-1", "Verified", "delta 成品 Q 頁（AI 占比約四分之一）"),
    ("avc", "奇鋐營收（NT$B）：3Q26 ÷ 1H26 季平均", "NT$B", "=98.16/2", "=59.09", "=E{r}/D{r}-1", "Verified", "avc 成品 Q 頁"),
    ("inno", "中際旭創營收（RMB B）：3Q26 ÷ 1H26 季平均", "RMB B", "=41.78/2", "=27.77", "=E{r}/D{r}-1", "Verified", "innolight 成品 Q 頁"),
    ("crwv", "CoreWeave 營收（US$B）：Q3 2026 ÷ (Q1＋Q2)/2", "US$B", "=(2.08+2.58)/2", "=3.66", "=E{r}/D{r}-1", "Verified", "coreweave 成品 Q 頁"),
]
for i, row in enumerate(ll):
    rr = 4 + i
    for j, v in enumerate(row, 1):
        if isinstance(v, str) and "{r}" in v: v = v.replace("{r}", str(rr))
        c = LL.cell(rr, j, v)
        if j in (4, 5) and v is not None and not (isinstance(v, str) and not v.startswith("=")): c.font = BLUE; c.number_format = "0.00"
        if j == 6: c.number_format = "0.0%"
rr = 4 + len(ll) + 1
LL.cell(rr, 1, "判讀").font = BOLD
LL.cell(rr + 1, 1, '="9 家中 "&COUNTIF(F9:F17,">"&F7)&" 家的比值高於 NVIDIA 同窗（"&TEXT(F7,"0%")&"），"&COUNTIF(F9:F17,">"&F8)&" 家高於次窗（"&TEXT(F8,"0%")&"，含指引）。價格（記憶體）、ASIC（緯穎）、非 AI 與份額變化同時在動；每家只有 1 個比值點，不足以判斷領先或落後"')
LL.cell(rr + 2, 1, "SK 海力士另有 7 季（2025Q1–2026Q3），但同期 DRAM 價格大漲，季增無法拆出量；台積電 HPC 只有 3 季。需要月營收（台股）或出貨量揭露才可能檢驗。")
for col, w in zip("ABCDEFGH", (10, 60, 8, 12, 12, 10, 18, 50)):
    LL.column_dimensions[col].width = w

wb.save(OUT)
print("saved", OUT, FR, KR)
