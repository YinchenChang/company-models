# Vertiv（VRT）個股模型 v0.1.0（2026-11-17；AI 供應鏈：電源與散熱，美國機房設施端）
# 用法：python3 vrt/build_vrt.py raw.xlsx ；再以 LibreOffice headless 重算（LO_PROFILE=/tmp/lo_n5_vrt）
# 會計年度＝日曆年，FY(t)＝CY(t)。金額：十億美元（US$B）；EPS：美元；股數：十億股。口徑：GAAP（歸屬普通股淨利＝公司淨利）。
# 產業變數：機房設施市場讀 Wattonomics v0.3.6 快照（WIF_Spend_fac_dist／ups／cool；不讀 Siliconomics 的電源與散熱）。
# 環境變數：VRT_OVERRIDE="鍵=@lo;鍵=@hi;鍵=數值"（敏感度整本重算用）；VRT_ROWMAP=路徑（輸出列位置供 gen_html_json.py）；VRT_SENS_RUN=1（略過業務主因表）。
import sys, os, json, glob, csv, hashlib
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
import fcff_dcf as FD
import relative_multiple as RMX

VER, BUILT = "0.1.0", "2026-11-17"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
SC = ("A", "B", "C")
OVR = {}
for part in filter(None, os.environ.get("VRT_OVERRIDE", "").split(";")):
    k, v = part.split("=", 1); OVR[k] = v

# ---------------- 外部資料（唯讀）：Wattonomics 快照 v0.3.6；價格檔 ----------------
WF = os.path.join(ROOT, "wattonomics", "data", "wattonomics_snapshot_v0.3.6.json")
WS = json.load(open(WF, encoding="utf-8"))
WN = WS["names"]; WSHA = WS["source"]["sha256"]
assert WS["source"]["version"] == "v0.3.6", WS["source"]
def px_of(tk):
    out = None
    for f in ("markets/invest_map/data/closes_20261020_20261113.csv", "markets/invest_map/data/closes_extra_20261113.csv"):
        for r in csv.reader(open(os.path.join(ROOT, f), encoding="utf-8")):
            if r and r[0] == tk and r[2] == "2026-11-13": out = float(r[3])
    assert out, tk
    return out
PX = px_of("VRT")
CLOSES = "markets/invest_map/data/closes_extra_20261113.csv（2026-11-13 收盤；Yahoo Finance 日線，主串整理）"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；比例：%。每個輸入附來源（找不到者寫「未找到」）。"
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived（每列單一標記）；Interested-party 列另寫「來源誘因」。公司原文＝Vertiv 財報新聞稿（SEC 8-K 附件 99.1）與 10-Q；"
            "Yahoo Finance 只作交叉核對，差異見 Recon 與 README。SemiAnalysis 未使用。")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明", "來源誘因"])
rows = []
sec = lambda t: rows.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, s1, s2="—", inc=None):
    if key in OVR:
        v = OVR[key]
        base = {"@lo": lo, "@hi": hi}.get(v, None) if v in ("@lo", "@hi") else float(v)
        assert base is not None, (key, v)
    rows.append((key, item, unit, lo, base, hi, tag, s1, s2, inc))

Q3PR = "Vertiv 3Q26 新聞稿（SEC 8-K 附件 99.1，2026-10-28；https://www.sec.gov/Archives/edgar/data/0001674101/000162828026067707/q32026exhibit991vrt10282026.htm）"
Q310 = "Vertiv 3Q26 10-Q（2026-09-30；https://www.sec.gov/Archives/edgar/data/1674101/000162828026067893/vrt-20260930.htm）"
Q2PR = "Vertiv 2Q26 新聞稿（SEC 8-K 附件 99.1，2026-07-29；https://www.sec.gov/Archives/edgar/data/0001674101/000162828026050323/q22026exhibit991vrt07292026.htm）"
FYPR = "Vertiv 4Q25／FY2025 新聞稿（https://s205.q4cdn.com/554782763/files/doc_financials/2025/q4/Vertiv_Fourth-Quarter-2025-Earnings-Release.pdf）"
WATT = f"Wattonomics 快照 v0.3.6（wattonomics/data/wattonomics_snapshot_v0.3.6.json，sha256 {WSHA[:12]}…）"
INC_SELL = "賣方與資料商報價；與公司或承銷無直接關係，但賣方有評等與覆蓋誘因"
INC_MGMT = "公司管理層指引：有行銷與維持股價的誘因，且指引偏保守的慣例使其不等於實際"

sec("A. 機房設施市場（讀 Wattonomics v0.3.6；單位 US$B；Vertiv 可及市場＝配電＋UPS＋冷卻，不含發電機組）")
for y in (2025, 2026, 2027, 2028, 2029):
    for nm, lab in (("dist", "配電（開關櫃、PDU、母線等；機櫃外）"), ("ups", "UPS（機櫃外）"), ("cool", "冷卻（CDU、冰水機、CRAH 等；機櫃外）"), ("gen", "發電機組（Vertiv 不計入可及市場，只作對照）")):
        inp(f"mk_{nm}_{y % 100}", f"機房設施支出：{lab} {y}", "US$B", None, WN[f"WIF_Spend_fac_{nm}_{y}"], None, "Derived",
            f"{WATT}：names.WIF_Spend_fac_{nm}_{y}", "邊界：機櫃內（電源櫃、冷板、機櫃內歧管、BBU）屬 IT 端（台達電），不計入；機櫃外屬設施端。Vertiv 的 CDU 歸冷卻" if (y, nm) == (2025, "dist") else "同上")
sec("B. 情境（個股自訂；情境 A 樂觀、B 基準、C 保守；以設施市場 2028–2029 的倍數表達）")
for s, a28, a29 in (("A", 1.10, 1.15), ("B", 1.00, 1.00), ("C", 0.80, 0.65)):
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 設施市場 × 倍數", "x", None, a28, None, "Assumed", "個股自訂；B＝Wattonomics 基準原值，A／C 為上下緣（建議由 Wattonomics 提供官方上下情境）", "2027 三情境同為 1.0")
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 設施市場 × 倍數", "x", None, a29, None, "Assumed", "同上", "—")
inp("mk_adj", "設施市場 2028–2029 倍數（Sens 用；三情境同乘）", "x", 0.85, 1.0, 1.15, "Assumed", "Sens 用 ±15%", "—")
sec("C. 營收結構：資料中心設施產品營收 ＝ Vertiv 份額 × 可及市場；其餘（服務、備品、非資料中心）另算")
inp("dcp_25", "2025 年資料中心設施產品營收 ÷ 營收", "%", None, 0.70, None, "Assumed", "10-Q 3Q26 營收分類（Note 4）：產品（不含備品）3,593.8 ÷ 營收 4,312.8＝83%（Verified）；扣除非資料中心產品（商業、工業等）後推為 70%（Assumed）",
    "公司未揭露資料中心占比：未找到")
inp("dcp_26", "2026 年資料中心設施產品營收 ÷ 營收", "%", 0.65, 0.72, 0.80, "Assumed", Q310 + "：3Q26 產品（不含備品）占 83%；9M 服務占 17.7%", "Derived 份額＝本值 × FY2026 營收 ÷ 可及市場，區間進 Sens（份額基準變動）")
for y, lo, b, hi in ((27, 0.95, 1.03, 1.10), (28, 0.95, 1.02, 1.08), (29, 0.93, 1.00, 1.05)):
    inp(f"rel_{y}", f"Vertiv 份額相對值 20{y} ÷ 2026", "x", lo, b, hi, "Assumed", "backlog 與訂單支持份額小幅提升（FY2025 底 backlog 150 億美元、4Q25 book-to-bill 約 2.9；3Q26 有機訂單成長逾 100%）",
        "FY2025→FY2026 份額由訂單與出貨時差而下降（見 Coord）" if y == 27 else "同上")
inp("g_oth", "其餘營收（服務、備品、非資料中心）年增（FY2027 起）", "%", 0.02, 0.06, 0.10, "Assumed", Q310 + "：9M 服務 18.2 億美元（年增 38%，含併購）；備品與服務年增低於產品", "—")
sec("D. 損益")
inp("gm_x", "毛利率相對 FY2026 錨點的變動（FY2027 起，各情境共用）", "pt", -0.02, 0.0, 0.015, "Assumed", "FY2025 36.3%、9M26 37.9%、3Q26 38.1%（公司）；AI 機架混合與價格壓力互抵", "—")
for s, v in (("A", 0.005), ("B", 0.0), ("C", -0.015)):
    inp(f"gms_{s}", f"情境 {s}：毛利率額外加減（個股自訂）", "pt", None, v, None, "Assumed", "情境定義：下行情境含價格壓力與低稼動", "—")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.60, 0.75, 0.95, "Derived", "FY2025 營業費用 +20.3%（毛利減營業利益）÷ 營收 +27.7%＝0.73；FY2026 錨點 0.74（Q 頁）", "—")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "Sens 用 ±2 個百分點", "—")
inp("nop_f", "業外淨額（淨利息與其他；每年）", "US$B", -0.04, 0.0, 0.04, "Assumed", "9M26 淨利息費用 19.4、債務清償損失 6.2、其他收入 16.9（公司）；現金增加使淨利息趨近零", "—")
inp("tax_f", "有效稅率（FY2027 起）", "%", 0.20, 0.22, 0.24, "Assumed", "FY2025 23.5%、9M26 19.5%、3Q26 23.2%（公司，所得稅 ÷ 稅前）", "—")
inp("sh_g", "稀釋股數年變動", "%", 0.0, 0.003, 0.006, "Assumed", "FY2025 稀釋 390.7M → 3Q26 392.7M（公司）；無買回", "—")
sec("E. FY2026 錨點（1Q–3Q 實績＋4Q 假設）")
inp("q4_rev", "4Q26 營收", "US$B", 4.313, 4.413, 4.513, "Interested-party", Q3PR + "：4Q26 指引 4,313–4,513（百萬美元），中位 4,413", "FY2026 指引 14,450–14,850", INC_MGMT)
inp("q_gm", "4Q26 毛利率", "%", 0.370, 0.385, 0.395, "Assumed", "3Q26 38.1%（公司）；4Q 指引調整後營業利益率 25.3–26.3%", "—")
inp("q_opx", "4Q26 營業費用 ÷ 營收（銷管、攤銷、其他與重組合計）", "%", 0.145, 0.152, 0.162, "Assumed", "3Q26：(556.9＋65.3＋0.2＋7.1＋30.1) ÷ 4,312.8＝15.3%（公司）", "—")
inp("q_nop", "4Q26 業外淨額", "US$B", -0.01, 0.005, 0.02, "Assumed", "3Q26：−6.4（淨利息）＋17.4（其他收入）＝＋11.0（公司）", "—")
inp("q_tax", "4Q26 有效稅率", "%", 0.20, 0.22, 0.24, "Assumed", "3Q26 23.2%、9M26 19.5%（公司）", "—")
inp("q_adj", "4Q26 調整項目（調整後營業利益 − GAAP 營業利益；Non-GAAP 橋接）", "US$B", 0.09, 0.105, 0.12, "Assumed", "3Q26：1,086.9 − 983.8＝103.1（無風險：主要為攤銷 65.3 與或有對價 33.5）（公司）", "只用於 Q 頁 Non-GAAP 列；DCF 用 GAAP 營業利益")
inp("q_sh", "稀釋股數（十億股）", "十億股", None, 0.392682064, None, "Verified", Q3PR + "：3Q26 稀釋加權平均 392,682,064 股", "10-Q 封面：2026-10-26 流通 A 類股 385,060,421（只作對照）")
sec("F. 長期收斂（FY2030–FY2032）與目標價規則")
inp("tp_g", "收斂基準成長率（FY2032 營收年增；B 情境）", "%", 0.02, 0.03, 0.04, "Assumed", "永續成長率上限 3%（Andy 2026-11-05）", "A＝本值＋2pt、C＝本值−2pt")
for s, d in (("A", 0.02), ("B", 0.0), ("C", -0.02)):
    inp(f"fade_g{s}", f"情境 {s}：FY2032 營收年增", "%", None, f"=tp_g+({d})", None, "Derived", "＝tp_g＋情境差（A＋2pt、B 0、C −2pt）", "—")
inp("opm_f", "營業利益率收斂係數（FY2032 目標＝FY2029 × 係數；C 取係數、B 取 (1＋係數) ÷ 2、A 不收斂）", "x", 0.75, 0.85, 0.95, "Assumed", "毛利率循環回落與同業競爭（Schneider、Eaton、台達電擴產）", "—")
for s, v in (("A", 0.30), ("B", 0.50), ("C", 0.20)):
    inp(f"tp_w{s}", f"情境 {s} 權重", "%", None, v, None, "Assumed", "同 delta／nvidia 慣例 0.3／0.5／0.2", "Andy 可改")
inp("tp_w_dcf", "DCF 權重（Comparable 權重＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股 DCF 標準規格：DCF 與 Comparable 各半", "Sens：0.3／0.7")
inp("px", "現價（2026-11-13 收盤）", "US$", None, PX, None, "Verified", CLOSES, "Yahoo Finance 日線收盤；只讀價格檔，不上網抓價")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, 334.29, None, "Interested-party", "Yahoo quoteSummary 2026-11-17：目標價均值 334.29（30 位分析師）；docs/workorders/20261117_新增五家_INTC_GFS_VRT_ABB_TEL_Yahoo基本面快照.md", "第二來源：未找到", INC_SELL)
inp("cons_eps27", "共識 FY2027 EPS（調整後口徑）", "US$", None, 9.874, None, "Interested-party", "Yahoo quoteSummary 2026-11-17：2027-12 EPS 9.874、營收 19.54B（共識為 Non-GAAP，本模型為 GAAP）", "共識 FY2026 EPS 7.191、營收 14.72B", INC_SELL)
inp("cons_rev27", "共識 FY2027 營收", "US$B", None, 19.54, None, "Interested-party", "同上", "—", INC_SELL)

# 折現率與自由現金流（tools/fcff_dcf.py）
KD = "=(0.850*0.04125+0.600*0.0485+0.500*0.0565+0.500*0.058+0.500*0.0595)/2.95"
for _row in FD.in_rows(dict(
        date="2026-12-31", ccy="US$", amt="US$B", shu="十億股",
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=Q310 + "（資產負債表）",
                  div=(0.024, "Verified", Q310 + "：2026 年 3Q 宣告股利 24.0（百萬美元）；於 4Q 發放", "UtilityInnovation 收購（2026-10-26 交割，現金 14.5 億）見 uig_adj：基準視為價值中性，不扣也不加")),
        rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；與 NVIDIA 模型同值）",
            "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）", "無風險利率（美國 10 年期公債殖利率）"),
        betas=[("Yahoo Finance", 2.11, "Yahoo quoteSummary 2026-11-17：Beta 2.11（5 年月報酬）；其餘來源：未找到")],
        beta_lohi=(1.45, 2.05, "Analogy：同業（Eaton、Schneider）β 約 1.0–1.3，與 Vertiv 自身 Blume 調整後 1.74 之間取低 1.45；高取 2.05（Assumed）"),
        kd=(KD, "Derived", Q310 + "：票面加權（優先擔保票 2028 4.125% 850、2036 4.85% 600、2046 5.65% 500、2056 5.80% 500、2066 5.95% 500；合計 2,950）", "循環額度 2,500 未提款；舊定期貸款已於 2026-03 清償"),
        tax_w=(0.21, "Assumed", "美國聯邦法定稅率 21%（債務稅盾）", "NOPAT 用 tax_f（有效稅率）"),
        debt=(2.9285, "Verified", Q3PR + "：總債務 2,928.5（流動 3.4＋長期 2,925.1；百萬美元）", "2025-12-31：2,913.0（公司）"),
        lease=(0.3397, "Verified", Q3PR + "：長期租賃負債 339.7（百萬美元）", "流動部分未揭露：未找到（低估債務約 0.1）；2025-12-31：245.2（公司）"),
        cash=("=3.5129+0.1988", "Verified", Q3PR + "：現金 3,512.9＋短期投資 198.8（百萬美元）", "2025-12-31：1,728.4＋99.5（公司）"),
        inv=(0.0, "Verified", Q310 + "：無權益證券等非營業投資", "—"),
        mi=(0.0, "Verified", Q310 + "：資產負債表無非控制權益", "—"),
        dps=("=4*0.024/0.385060421", "Derived", Q310 + "：季度股利 24.0 ÷ 流通 385.06M 股＝每股約 0.062；評價日到目標時點假設四次、金額不變", "—"),
        sh=("=q_sh", "Derived", "＝In 頁 q_sh（3Q26 稀釋加權股數）", "流通股 385.06M 較稀釋股少約 7.6M（RSU 等）；評價用稀釋股"),
        fc_da=(None, 0.032, None, "Derived", Q310 + "：9M26 折舊 106.7＋攤銷 224.7＝331.4 ÷ 營收 10,236.6＝3.2%（FY2025 3.0%）", "—"),
        fc_cx=(None, 0.040, None, "Verified", Q310 + "：FY2026 資本支出（含資本化軟體）指引 575–600，約占營收 4.0%（公司）", "FY2025 2.2%；Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=fc_da_ss*(1+5*g_tv)", "Assumed", "長期穩態資本支出 ÷ 營收＝折舊 ÷ 營收 ×（1＋5g）", "滿足規格：資本支出 ≥ 折舊 ×（1＋g）"),
        fc_da_ss=(0.028, "Assumed", "攤銷隨併購攤完而下降；穩態折舊攤銷 2.8%", "—"),
        fc_wc=(0.02, 0.05, 0.10, "Assumed", "營運資金增加＝5% × 營收增量；9M26 貿易營運資金為現金流入 540.5（客戶預收），故取偏低值（Verified 事實＋Assumed 比例）", "Sens 低／高 2%／10%"),
        )):
    rows.append(_row)
inp("uig_adj", "UtilityInnovation 收購現金（Sens 用；基準 0＝價值中性）", "US$B", None, 0.0, 1.45, "Assumed", "2026-10-26 交割：現金 14.5 億美元、另有最多 11.5 億美元 EBITDA 對價（unite.ai 2026-09-02 轉述公告；10-Q 僅列合併協議，未載金額）", "基準：被收購業務價值≈價格，不調整淨現金也不加營收；高：現金流出但不計獲利（下緣）")

r = 4
INLAST = None
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow[:9]
    inc = srow[9] if len(srow) > 9 else None
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    if inc: IN.cell(r, 10, inc)
    if tag == "Interested-party" and not inc: IN.cell(r, 10, INC_SELL)
    defname(key, f"In!$E${r}")
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    INLAST = r; r += 1
FD.yellow(IN)
for col, w in zip("ABCDEFGHIJ", [12, 58, 10, 9, 11, 9, 18, 70, 70, 40]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億美元；EPS 美元；股數十億股）。藍字＝輸入；黑字＝公式。實績取自 Vertiv 新聞稿與 10-Q（百萬美元 ÷ 1,000）；4Q26 為假設。"
Q["A2"] = ("1Q26 為 1H26 減 2Q26（公司未單獨列示 1Q，新聞稿 6M 欄）；各季營業利益 Yahoo 與公司略有差異（Yahoo 1Q26 433.6、2Q26 637.9、3Q26 991.1；公司 440.1、637.9、983.8），本表用公司值。"
           "Non-GAAP 橋接：調整後營業利益（公司）− GAAP 營業利益＝攤銷、重組、併購與或有對價等；DCF 用 GAAP 營業利益（攤銷於 D&A 加回，股權報酬視為費用）。")
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明", "來源誘因"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1="—", s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
    if tag == "Interested-party": Q.cell(r, 11, INC_MGMT)
S_Q = f"{FYPR}（FY2025）；{Q2PR}（1H26、2Q26）；{Q3PR}（3Q26）"
qrow(4, "營收", {"B": 10.2299, "C": 2.6495, "D": 3.2743, "E": 4.3128, "F": "=q4_rev"}, s1=S_Q, s2="4Q26＝In 頁 q4_rev（公司指引中位，Interested-party）", tag="Verified")
qrow(5, "毛利", {"B": 3.7152, "C": 0.9997, "D": 1.2349, "E": 1.6434, "F": "=F4*q_gm"}, s1=S_Q, s2="毛利＝營收 − 銷貨成本（公司未列毛利）", tag="Verified")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s2="FY2025 36.3%；3Q26 38.1%")
qrow(7, "營業費用（銷管、攤銷、重組、匯兌與其他）", {c: f"={c}5-{c}8" for c in "BCDE"} | {"F": "=F4*q_opx"}, tag="Derived", s1="毛利 − 營業利益", s2="4Q26＝營收 × In 頁 q_opx")
qrow(8, "營業利益（GAAP）", {"B": 1.8297, "C": 0.4401, "D": 0.6379, "E": 0.9838, "F": "=F5-F7"}, s1=S_Q, tag="Verified", s2="1Q26＝1H26 1,078.0 − 2Q26 637.9（公司新聞稿推導）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s2="FY2025 17.9%")
qrow(10, "業外淨額（淨利息、債務清償、其他）", {"B": "=B11-B8", "C": "=C11-C8", "D": "=D11-D8", "E": "=E11-E8", "F": "=q_nop"}, tag="Derived", s1="稅前 − 營業利益", s2="4Q26＝In 頁 q_nop")
qrow(11, "稅前利益", {"B": 1.7419, "C": 0.4383, "D": 0.6200, "E": 0.9948, "F": "=F8+F10"}, s1=S_Q, tag="Verified", s2="1Q26＝6M 1,058.3 − 2Q26 620.0")
qrow(12, "所得稅", {"B": 0.4091, "C": 0.0482, "D": 0.1222, "E": 0.2303, "F": "=F11*q_tax"}, s1=S_Q, tag="Verified", s2="1Q26＝6M 170.4 − 2Q26 122.2")
qrow(13, "淨利（歸屬普通股）", {c: f"={c}11-{c}12" for c in "BCDEF"}, tag="Derived", s1="稅前 − 所得稅（公司淨利 FY2025 1,332.8；1Q26 390.1；2Q26 497.8；3Q26 764.5，逐項吻合）", s2="—")
qrow(14, "稀釋股數（十億股）", {"B": 0.3907, "C": 0.3921, "D": 0.3928, "E": "=q_sh", "F": "=q_sh"}, ann="=q_sh", tag="Verified", s1=S_Q, s2="1Q、2Q 取 Yahoo 稀釋股數（公司 1Q 未單獨列示）；FY2026 取 3Q26 稀釋加權")
qrow(15, "EPS（GAAP，稀釋）", {c: f"={c}13/{c}14" for c in "BCDEF"}, fmt="0.00", ann="=G13/G14", tag="Derived", s1="淨利 ÷ 股數", s2="公司 EPS：FY2025 3.41、2Q26 1.27、3Q26 1.95、9M26 4.21")
qrow(16, "調整後營業利益（Non-GAAP，公司口徑）", {"B": 2.0897, "C": 0.5522, "D": 0.7384, "E": 1.0869, "F": "=F8+q_adj"}, s1=S_Q, tag="Verified", s2="4Q26＝GAAP＋In 頁 q_adj；公司指引 4Q26 調整後營業利益 1,118–1,158")
qrow(17, "調整項目（調整後 − GAAP 營業利益）", {c: f"={c}16-{c}8" for c in "BCDEF"}, tag="Derived", s1="Non-GAAP 橋接", s2="主要為無形資產攤銷、重組、併購與或有對價")
qrow(18, "調整後營業利益率", {c: f"={c}16/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s2="公司：FY2026 指引 23.5–24.5%")
Q["G16"] = "=SUM(C16:F16)"; Q["G17"] = "=G16-G8"
Q.cell(20, 1, "FY2026 錨點的檢查（參考；校準未使用共識）").font = BOLD
for rr, lab, f, fm in ((21, "FY2026 營收年增", "=G4/B4-1", "0.0%"), (22, "FY2026 營業費用年增", "=G7/B7-1", "0.0%"), (23, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）", "=B22/B21", "0.00"),
                       (24, "FY2026 營收 ÷ 公司指引中位（14,650）− 1", "=G4/14.65-1", "0.0%"), (25, "FY2026 GAAP EPS ÷ 公司 GAAP 指引中位（6.29）− 1", "=G15/6.29-1", "0.0%"),
                       (26, "FY2026 調整後營業利益 ÷ 公司指引中位（3,515）− 1", "=G16/3.515-1", "0.0%")):
    Q.cell(rr, 1, lab); Q[f"B{rr}"] = f; Q[f"B{rr}"].number_format = fm
Q.cell(27, 1, "公司指引（FY2026，2026-10-28）：營收 14,450–14,850；調整後營業利益 3,475–3,555；調整後 EPS 7.10–7.20；GAAP EPS 中位 6.29；調整後自由現金流 2,600–2,800（Interested-party，管理層指引）")
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nop", 10), ("pt", 11), ("tax", 12), ("ni", 13), ("sh", 14), ("eps", 15), ("adj", 16)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 14; Q.column_dimensions["I"].width = 90; Q.column_dimensions["J"].width = 60
Q.freeze_panes = "B4"

# ---------------- Hist：歷史財務（公司原文） ----------------
HS = wb.create_sheet("Hist")
HS["A1"] = "Hist — 歷史財務（十億美元；EPS 美元；FY2023–FY2025 公司新聞稿）。FY2023、FY2024 的稅前利益含認股權證公允價值變動（FY2023 −157.9、FY2024 −449.2 百萬美元），淨利不可與 FY2025 直接比較。"
hdr(HS, 3, ["項目", "FY2023A", "FY2024A", "FY2025A", "標記", "來源", "說明"])
H_SRC = "FY2023：Vertiv 4Q23 新聞稿（SEC 8-K 附件 99.1，2024-02-21）；FY2024、FY2025：" + FYPR
HDATA = [("營收", (6.8632, 8.0118, 10.2299)), ("毛利", (2.4005, 2.9342, 3.7152)), ("營業利益（GAAP）", (0.8722, 1.3674, 1.8297)), ("稅前利益", (0.5337, 0.7654, 1.7419)),
         ("所得稅", (0.0735, 0.2696, 0.4091)), ("淨利", (0.4602, 0.4958, 1.3328)), ("稀釋股數（十億股）", (0.386226, 0.3863, 0.3907)), ("調整後營業利益（Non-GAAP）", (1.0535, 1.5516, 2.0897)),
         ("折舊攤銷（現金流量表）", (0.2710, 0.2770, 0.3086)), ("資本支出", (0.1279, 0.1670, 0.2200))]
for i, (lab, vals) in enumerate(HDATA, 4):
    HS.cell(i, 1, lab)
    for j, v in enumerate(vals):
        c = HS.cell(i, 2 + j, v); c.font = BLUE; c.number_format = "0.0000" if "股數" in lab else "#,##0.000"
    HS.cell(i, 5, "Verified"); HS.cell(i, 6, H_SRC)
HS.cell(14, 1, "EPS（稀釋，計算值）"); HS.cell(15, 1, "毛利率"); HS.cell(16, 1, "營業利益率"); HS.cell(17, 1, "營業費用（毛利 − 營業利益）")
for j in range(3):
    L = "BCD"[j]
    HS[f"{L}14"] = f"={L}9/{L}10"; HS[f"{L}14"].number_format = "0.00"
    HS[f"{L}15"] = f"={L}5/{L}4"; HS[f"{L}15"].number_format = "0.0%"
    HS[f"{L}16"] = f"={L}6/{L}4"; HS[f"{L}16"].number_format = "0.0%"
    HS[f"{L}17"] = f"={L}5-{L}6"; HS[f"{L}17"].number_format = "#,##0.000"
for rr in (14, 15, 16, 17): HS.cell(rr, 5, "Derived")
HS.cell(14, 7, "公司 EPS：FY2023 1.19、FY2024 1.28、FY2025 3.41")
HS.cell(19, 1, "Yahoo 交叉核對（Yahoo quoteSummary 2026-11-17）：營業利益 FY2023 916.8（公司 872.2）、FY2024 1,382.0（公司 1,367.4）、FY2025 1,896.2（公司 1,829.7）；差異來自 Yahoo 營業利益口徑（疑似排除重組與其他項目），本模型一律用公司值。")
HS.column_dimensions["A"].width = 40; HS.column_dimensions["F"].width = 90
# ---------------- Units：設施市場與份額 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 機房設施市場（Wattonomics v0.3.6，US$B）、Vertiv 可及市場與 Vertiv 份額（本專案推導；份額＝Vertiv 資料中心設施產品營收 ÷ 可及市場，Derived）"
UN["A2"] = ("邊界：機櫃內（PSU、電源櫃、冷板、機櫃內歧管、BBU）屬 IT 端（台達電），不計入；機櫃外（UPS、配電、CDU、冰水機、CRAH）屬設施端。Vertiv 可及市場＝配電＋UPS＋冷卻，發電機組不計（Vertiv 非主要供應商）。"
            "Wattonomics 2030–2032 的新增容量下降（市場金額 2030 起轉降），本模型 FY2030 起不讀市場而用收斂法（見 PL 說明）。")
hdr(UN, 3, ["項目", "單位", "2025", "2026", "2027", "2028", "2029", "說明"])
UC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
for i, (nm, lab) in enumerate((("dist", "配電"), ("ups", "UPS"), ("cool", "冷卻"), ("gen", "發電機組（不計入可及市場）")), 4):
    UN.cell(i, 1, f"機房設施支出：{lab}"); UN.cell(i, 2, "US$B")
    for y, c in UC.items():
        UN[f"{c}{i}"] = f"=mk_{nm}_{y % 100}"; UN[f"{c}{i}"].number_format = "#,##0.00"
UN.cell(8, 1, "可及市場（配電＋UPS＋冷卻）"); UN.cell(8, 2, "US$B"); UN.cell(9, 1, "Wattonomics 設施支出合計（含發電、土建、人工等，只作對照）"); UN.cell(9, 2, "US$B")
UN.cell(10, 1, "可及市場年增"); UN.cell(10, 2, "%")
for y, c in UC.items():
    UN[f"{c}8"] = f"=SUM({c}4:{c}6)"; UN[f"{c}8"].number_format = "#,##0.00"
    UN[f"{c}9"] = WN[f"WIF_Spend_fac_sum_{y}"]; UN[f"{c}9"].number_format = "#,##0.00"; UN[f"{c}9"].font = BLUE
    if y > 2025: UN[f"{c}10"] = f"={c}8/{chr(ord(c) - 1)}8-1"; UN[f"{c}10"].number_format = "0.0%"
UN.cell(9, 8, f"{WATT}：names.WIF_Spend_fac_sum_<年>（Derived）")
UN.cell(11, 1, "Vertiv 資料中心設施產品營收（Q 頁 × dcp，基準）"); UN.cell(11, 2, "US$B")
UN["C11"] = "=a25_rev*dcp_25"; UN["D11"] = "=a26_rev*dcp_26"
UN.cell(12, 1, "Vertiv 份額（÷ 可及市場；Derived）"); UN.cell(12, 2, "%")
UN["C12"] = "=C11/C8"; UN["D12"] = "=D11/D8"
for c in "CD": UN[f"{c}11"].number_format = "#,##0.00"; UN[f"{c}12"].number_format = "0.0%"
UN.cell(12, 8, "FY2025 份額高於 FY2026：Wattonomics 市場以裝機年度計支出，Vertiv 於出貨認列，backlog（FY2025 底 150 億美元）使營收落後支出；份額的區間進 Sens（dcp_26 與 rel 兩組鍵）")
UN.cell(13, 1, "Vertiv 營收 ÷ Wattonomics 設施支出合計（對照）"); UN["C13"] = "=a25_rev/C9"; UN["D13"] = "=a26_rev/D9"
for c in "CD": UN[f"{c}13"].number_format = "0.0%"
UN.column_dimensions["A"].width = 62; UN.column_dimensions["H"].width = 90
for c in "CDEFG": UN.column_dimensions[c].width = 12

# ---------------- PL：三情境引擎（FY2025A–FY2032） ----------------
PL = wb.create_sheet("PL")
YR = (2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032)
FC = {y: "CDEFGHIJ"[i] for i, y in enumerate(YR)}
HEAD = ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]
SCN_T = {"A": "情境 A：樂觀（設施市場 2028 ×1.10、2029 ×1.15）", "B": "情境 B：基準（Wattonomics 基準原值）", "C": "情境 C：保守（設施市場 2028 ×0.80、2029 ×0.65）"}
def engine(ws, r0, title, s):
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, HEAD); r += 1
    def row(key, label, unit, fs, fmt="#,##0.000", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for y, f in fs.items():
            c = ws[f"{FC[y]}{r}"]; c.value = f; c.number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    um = {2027: "1", 2028: f"um28_{s}*mk_adj", 2029: f"um29_{s}*mk_adj"}
    row("mkt", "可及市場（Wattonomics 配電＋UPS＋冷卻）× 情境倍數", "US$B", {2025: "=Units!C8", 2026: "=Units!D8", **{y: f"=Units!{UC[y]}8*{um[y]}" for y in (2027, 2028, 2029)}}, "#,##0.00", "FY2030 起不讀市場（收斂法）")
    row("sh0", "Vertiv 份額（FY2026 校準）", "%", {2026: f"=D{r + 2}/D{R['mkt']}"}, "0.0%", "Derived：資料中心設施產品營收 ÷ 可及市場")
    row("rel", "份額相對值（÷ FY2026）", "x", {2026: "=1", **{y: f"=rel_{y % 100}" for y in (2027, 2028, 2029)}}, "0.00")
    row("dcp", "資料中心設施產品營收", "US$B", {2025: "=a25_rev*dcp_25", 2026: "=a26_rev*dcp_26", **{y: f"={FC[y]}{R['mkt']}*$D${R['sh0']}*{FC[y]}{R['rel']}" for y in (2027, 2028, 2029)}}, "#,##0.000", "＝可及市場 × 份額（命題的驅動）")
    row("oth", "其餘營收（服務、備品、非資料中心）", "US$B", {2025: "=a25_rev-C%d" % r, 2026: "=a26_rev-D%d" % r, **{y: f"={FC[y - 1]}{r + 1}*(1+g_oth)" for y in (2027, 2028, 2029)}}, "#,##0.000")
    # 上一列 oth 的 r 已遞增；修正引用（oth 本列位置＝R['oth']）
    for y in (2027, 2028, 2029): ws[f"{FC[y]}{R['oth']}"] = f"={FC[y - 1]}{R['oth']}*(1+g_oth)"
    ws[f"C{R['oth']}"] = f"=a25_rev-C{R['dcp']}"; ws[f"D{R['oth']}"] = f"=a26_rev-D{R['dcp']}"
    ws[f"D{R['sh0']}"] = f"=D{R['dcp']}/D{R['mkt']}"
    gR = r + 1   # rev 之後的 g 列位置在 rev 下一列
    row("rev", "營收", "US$B", {2025: "=a25_rev", 2026: "=a26_rev", **{y: f"={FC[y]}{R['dcp']}+{FC[y]}{R['oth']}" for y in (2027, 2028, 2029)}, **{y: f"={FC[y - 1]}{r}*(1+{FC[y]}{gR})" for y in (2030, 2031, 2032)}}, "#,##0.000",
        "FY2027–FY2029＝資料中心設施產品＋其餘；FY2030 起收斂法：年增由 FY2029 線性收斂到 fade_g")
    row("g", "  營收年增", "%", {y: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in (2026, 2027, 2028, 2029)} | {y: f"=$G${gR}+(fade_g{s}-$G${gR})*{y - 2029}/3" for y in (2030, 2031, 2032)}, "0.0%")
    row("dcs", "  資料中心設施產品 ÷ 營收", "%", {y: f"={FC[y]}{R['dcp']}/{FC[y]}{R['rev']}" for y in (2025, 2026, 2027, 2028, 2029)}, "0.0%")
    row("gm", "毛利率", "%", {2025: "=a25_gp/a25_rev", 2026: "=a26_gp/a26_rev", **{y: f"=$D${r}+gm_x+gms_{s}" for y in (2027, 2028, 2029)}, **{y: f"=$G${r}" for y in (2030, 2031, 2032)}}, "0.0%", "FY2027 起＝FY2026 錨點毛利率＋gm_x＋情境加減；FY2030 起維持 FY2029")
    row("gp", "毛利", "US$B", {y: f"={FC[y]}{R['rev']}*{FC[y]}{R['gm']}" for y in YR}, "#,##0.000")
    ws[f"C{R['gp']}"] = "=a25_gp"; ws[f"D{R['gp']}"] = "=a26_gp"
    oiR = r + 1
    row("opb", "營業費用（基準路徑）", "US$B", {2025: "=a25_opx", 2026: "=a26_opx", **{y: f"={FC[y - 1]}{r}*(1+MAX(0,phi*{FC[y]}{R['g']}))" for y in (2027, 2028, 2029)}}, "#,##0.000")
    row("opx", "營業費用", "US$B", {2025: f"=C{R['opb']}", 2026: f"=D{R['opb']}", **{y: f"={FC[y]}{R['opb']}+opx_adj*{FC[y]}{R['rev']}" for y in (2027, 2028, 2029)}, **{y: f"={FC[y]}{R['gp']}-{FC[y]}{oiR + 1}" for y in (2030, 2031, 2032)}}, "#,##0.000")
    omR = r + 1
    row("oi", "營業利益", "US$B", {y: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in (2025, 2026, 2027, 2028, 2029)} | {y: f"={FC[y]}{R['rev']}*{FC[y]}{omR}" for y in (2030, 2031, 2032)}, "#,##0.000")
    tm = {"A": f"=$G${omR}", "B": f"=$G${omR}*(1+opm_f)/2", "C": f"=$G${omR}*opm_f"}[s]
    row("opm", "營業利益率", "%", {y: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in (2025, 2026, 2027, 2028, 2029)} | {y: f"=$G${omR}+($J${r + 1}-$G${omR})*{y - 2029}/3" for y in (2030, 2031, 2032)}, "0.0%", "FY2030 起由 FY2029 線性收斂到 FY2032 目標")
    row("tm", "  FY2032 營業利益率目標", "%", {2032: tm}, "0.0%", {"A": "不收斂（＝FY2029）", "B": "FY2029 與 FY2029 × 收斂係數的中點", "C": "FY2029 × 收斂係數"}[s])
    row("nop", "業外淨額", "US$B", {2025: "=a25_nop", 2026: "=a26_nop", **{y: "=nop_f" for y in YR[2:]}}, "#,##0.000")
    row("pbt", "稅前利益", "US$B", {y: f"={FC[y]}{R['oi']}+{FC[y]}{R['nop']}" for y in YR}, "#,##0.000")
    row("tax", "所得稅", "US$B", {2025: "=a25_tax", 2026: "=a26_tax", **{y: f"={FC[y]}{R['pbt']}*tax_f" for y in YR[2:]}}, "#,##0.000")
    row("ni", "淨利", "US$B", {y: f"={FC[y]}{R['pbt']}-{FC[y]}{R['tax']}" for y in YR}, "#,##0.000")
    row("nim", "  淨利率", "%", {y: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in YR}, "0.0%")
    row("sh_", "稀釋股數", "十億股", {2025: "=a25_sh", 2026: "=a26_sh", **{y: f"={FC[y - 1]}{r}*(1+sh_g)" for y in YR[2:]}}, "0.0000")
    row("eps", "EPS（GAAP）", "US$", {y: f"={FC[y]}{R['ni']}/{FC[y]}{R['sh_']}" for y in YR}, "0.00")
    row("epsg", "  EPS 年增", "%", {y: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in YR[1:]}, "0.0%")
    row("chk", "檢查：FY2026 營收（情境值）− 錨點（應為 0）", "US$B", {2026: f"=D{R['rev']}-a26_rev"}, "0.0000")
    R["end"] = r
    return R
PB = {}; r0 = 60
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], s)
    r0 = PB[s]["end"] + 2
PL["A1"] = ("PL — Vertiv 損益（十億美元；EPS 美元；GAAP）。FY2025A 實績、FY2026 錨點（1Q–3Q 實績＋4Q 假設）、FY2027–FY2029 由設施市場 × 份額推導、FY2030–FY2032 收斂法（灰底＝非驅動因子推導）。FY(t)＝CY(t)。三情境各一個完整區塊（第 60 列起）。")
hdr(PL, 3, HEAD)
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("dcp", "  資料中心設施產品營收", "US$B", "#,##0.0"), ("oth", "  其餘營收", "US$B", "#,##0.0"),
                                ("gm", "毛利率", "%", "0.0%"), ("opm", "營業利益率", "%", "0.0%"), ("ni", "淨利", "US$B", "#,##0.0"), ("eps", "EPS", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in YR:
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("dcp", "oth") and y > 2029: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
assert sr < 60, sr
grey = PatternFill("solid", fgColor="E7E6E6")
for rr in range(4, 60):
    for y in (2030, 2031, 2032):
        if PL[f"{FC[y]}{rr}"].value is not None: PL[f"{FC[y]}{rr}"].fill = grey
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 14
# 標準損益表（情境 B；GAAP；FY2023A–FY2032E）
SR0 = PB["C"]["end"] + 3
PL.cell(SR0 - 1, 1, "標準損益表（GAAP 口徑，十億美元；EPS 為美元）：FY2023A–FY2025A 實績（Hist 頁）、FY2026E 錨點年、FY2027E–FY2032E 情境 B（基準）；下方另列 Non-GAAP 調整後營業利益（公司口徑，只作對照）").font = BOLD
SH = ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
hdr(PL, SR0, SH)
SL = [("營收", "US$B"), ("營收年增率", "%"), ("營業成本", "US$B"), ("毛利", "US$B"), ("毛利率", "%"), ("營業費用", "US$B"), ("營業利益", "US$B"), ("營業利益率", "%"),
      ("其他收入（業外淨額）", "US$B"), ("稅前利益", "US$B"), ("所得稅", "US$B"), ("稅後淨利", "US$B"), ("淨利率", "%"), ("EPS（GAAP，稀釋）", "US$"), ("淨利年增率", "%"), ("Non-GAAP：調整後營業利益（公司口徑，對照）", "US$B")]
SRW = {lab: SR0 + 1 + i for i, (lab, _) in enumerate(SL)}
for lab, unit in SL:
    PL.cell(SRW[lab], 1, lab); PL.cell(SRW[lab], 2, unit)
BM = PB["B"]; col = lambda i: chr(ord("C") + i)
for i in range(10):
    c = col(i)
    if i < 3:
        h = "BCD"[i]
        src = {"營收": f"=Hist!{h}4", "營業成本": f"=Hist!{h}4-Hist!{h}5", "毛利": f"=Hist!{h}5", "營業費用": f"=Hist!{h}17", "營業利益": f"=Hist!{h}6", "其他收入（業外淨額）": f"=Hist!{h}7-Hist!{h}6",
               "稅前利益": f"=Hist!{h}7", "所得稅": f"=Hist!{h}8", "稅後淨利": f"=Hist!{h}9", "EPS（GAAP，稀釋）": f"=Hist!{h}14", "Non-GAAP：調整後營業利益（公司口徑，對照）": f"=Hist!{h}11"}
    else:
        e = FC[2025 + i - 2]   # FY2025→col C of engine
        e = FC[2026 + (i - 3)]
        src = {"營收": f"={e}{BM['rev']}", "營業成本": f"={e}{BM['rev']}-{e}{BM['gp']}", "毛利": f"={e}{BM['gp']}", "營業費用": f"={e}{BM['opx']}", "營業利益": f"={e}{BM['oi']}", "其他收入（業外淨額）": f"={e}{BM['nop']}",
               "稅前利益": f"={e}{BM['pbt']}", "所得稅": f"={e}{BM['tax']}", "稅後淨利": f"={e}{BM['ni']}", "EPS（GAAP，稀釋）": f"={e}{BM['eps']}"}
        if i == 3: src["Non-GAAP：調整後營業利益（公司口徑，對照）"] = "=a26_adj"
    if i == 2:   # FY2025A 取 Hist；Hist 欄 D
        pass
    for lab, f in src.items(): PL[f"{c}{SRW[lab]}"] = f
    PL[f"{c}{SRW['毛利率']}"] = f"={c}{SRW['毛利']}/{c}{SRW['營收']}"
    PL[f"{c}{SRW['營業利益率']}"] = f"={c}{SRW['營業利益']}/{c}{SRW['營收']}"
    PL[f"{c}{SRW['淨利率']}"] = f"={c}{SRW['稅後淨利']}/{c}{SRW['營收']}"
    if i > 0:
        p = col(i - 1)
        PL[f"{c}{SRW['營收年增率']}"] = f"={c}{SRW['營收']}/{p}{SRW['營收']}-1"
        PL[f"{c}{SRW['淨利年增率']}"] = f"={c}{SRW['稅後淨利']}/{p}{SRW['稅後淨利']}-1"
    for lab, unit in SL:
        PL[f"{c}{SRW[lab]}"].number_format = "0.0%" if unit == "%" else ("0.00" if unit == "US$" else "#,##0.000")
SCHK = SR0 + len(SL) + 2
PL.cell(SCHK, 1, "檢查：標準損益表恆等式（毛利 − 營業費用 − 營業利益；稅前 − 稅後 − 所得稅；應為 0）")
for i in range(10):
    c = col(i); PL[f"{c}{SCHK}"] = f"=({c}{SRW['毛利']}-{c}{SRW['營業費用']}-{c}{SRW['營業利益']})+({c}{SRW['稅前利益']}-{c}{SRW['所得稅']}-{c}{SRW['稅後淨利']})"; PL[f"{c}{SCHK}"].number_format = "0.000"
PL.cell(SCHK + 1, 1, "檢查：標準損益表 FY2027E 營收 − 情境 B 摘要（應為 0）"); PL[f"G{SCHK + 1}"] = f"=G{SRW['營收']}-E{SUMR[('B', 'rev')]}"; PL[f"G{SCHK + 1}"].number_format = "0.000"
PL.cell(SCHK + 2, 1, "FY2025A 標準損益表稅後淨利 − Q 頁（應為 0）"); PL[f"E{SCHK + 2}"] = f"=E{SRW['稅後淨利']}-a25_ni"; PL[f"E{SCHK + 2}"].number_format = "0.000"
PL.cell(SCHK + 3, 1, "分部（官方：地區；FY2025 實績 Verified）").font = BOLD
SEG = (("Americas（美洲）", 6.3863), ("Asia Pacific（亞太）", 2.0192), ("EMEA（歐洲、中東與非洲）", 1.8244))
for k, (nm, v) in enumerate(SEG, SCHK + 4):
    PL.cell(k, 1, nm); PL.cell(k, 2, "US$B"); c = PL.cell(k, 5, v); c.font = BLUE; c.number_format = "#,##0.000"
    PL.cell(k, 12, "Vertiv FY2025 新聞稿 Segment net sales（Verified）")
SEGR = SCHK + 4
PL.cell(SEGR + 3, 1, "檢查：分部合計 − 合併營收（FY2025 實績）"); PL[f"E{SEGR + 3}"] = f"=E{SEGR}+E{SEGR + 1}+E{SEGR + 2}-E{SRW['營收']}"; PL[f"E{SEGR + 3}"].number_format = "0.000"
PL.cell(SEGR + 4, 1, "產品線拆分（本模型命題，非官方分部）：資料中心設施產品營收與其餘營收見各情境區塊與摘要；Vertiv 未揭露產品線營收")
PL.cell(SEGR + 5, 1, "資料中心設施產品營收（情境 B，推導）"); PL.cell(SEGR + 6, 1, "其餘營收（服務、備品、非資料中心；情境 B，推導）")
for y in (2025, 2026, 2027, 2028, 2029):
    PL[f"{FC[y]}{SEGR + 5}"] = f"={FC[y]}{BM['dcp']}"; PL[f"{FC[y]}{SEGR + 6}"] = f"={FC[y]}{BM['oth']}"
    PL[f"{FC[y]}{SEGR + 5}"].number_format = "#,##0.000"; PL[f"{FC[y]}{SEGR + 6}"].number_format = "#,##0.000"
PL.freeze_panes = "C4"
exec(open(os.path.join(HERE, "build_part2.py"), encoding="utf-8").read())
