# ABB（ABBN；SIX 瑞士證交所，股價 CHF；財報 USD）個股模型 v0.1.0（新增五家工作單，2026-11-17；AI 供應鏈 Project：電源與散熱）
# 用法：python3 abb/build_abb.py raw.xlsx ；再以 LibreOffice headless 重算。金額：十億美元（US$B）；股數：十億股；EPS：美元；股價：瑞士法郎（CHF）。
# 結構：README、In、Hist、Q、Wat、PL、DCF、Val、RevDCF、Sens、Recon、Coord、Sources。
# 資料口徑：公司原文＝ABB Q4 2025 financial information、Q3 2026 新聞稿（2026-10-20；Nasdaq／EQS 附件）；Yahoo 只作交叉核對，差異列 Recon。
# AI 資料中心變數讀 Wattonomics（設施電力＝機櫃外配電與 UPS；機櫃內屬 IT）快照 wattonomics/data/wattonomics_snapshot_v0.3.6.json（名稱，見 Wat 頁）。
# 環境變數：ABB_OVERRIDE＝"鍵=@lo;鍵=@hi;鍵=數值"（敏感度快照用；@lo／@hi＝In 頁敏感度低／高欄）；ABB_SNAP＝1（不讀 sens.json）。
import csv, hashlib, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
import fcff_dcf as FD  # noqa: E402

VER, BUILD_DATE, PX_DATE = "v0.1.0", "2026-11-17", "2026-11-13"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
SUB = PatternFill("solid", fgColor="EDEDED")
SC = ("A", "B", "C"); SCN = {"A": "順利", "B": "基準", "C": "不如預期"}
YEARS = list(range(2023, 2033)); YC = {y: get_column_letter(3 + i) for i, y in enumerate(YEARS)}   # PL：C=FY2023 … L=FY2032
FY_X = list(range(2027, 2033))

OVR = {}
for kv in os.environ.get("ABB_OVERRIDE", "").split(";"):
    if "=" in kv:
        k, v = kv.split("=", 1); OVR[k.strip()] = v.strip()

wb = openpyxl.Workbook()
def defname(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR

# ---------------- 價格（只讀兩個 CSV） ----------------
def close(sym):
    for fn in ("closes_20261020_20261113.csv", "closes_extra_20261113.csv"):
        for r in csv.DictReader(open(os.path.join(ROOT, "markets/invest_map/data", fn), encoding="utf-8")):
            if r["symbol"] == sym and r["date"] == PX_DATE: return float(r["close"])
    raise KeyError(sym)
PX_CHF, FX = close("ABBN.SW"), close("CHF=X")      # FX＝每美元 CHF
assert (PX_CHF, FX) == (69.96, 0.83593), (PX_CHF, FX)

# ---------------- Wattonomics 快照（唯讀） ----------------
WAT_F = "wattonomics/data/wattonomics_snapshot_v0.3.6.json"
WAT = json.load(open(os.path.join(ROOT, WAT_F), encoding="utf-8"))
WN = WAT["names"]; WSHA = WAT["source"]["sha256"][:8]
def wv(nm, y): return WN[f"{nm}_{y}"]

# =============================================================== In
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，敏感度低 D、敏感度高 F；低／高只用於 Sens 頁，一次動一個參數，不進目標價）。金額：十億美元（US$B）；股數：十億股；股價：瑞士法郎（CHF）。"
IN["A2"] = ("標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場，並填來源誘因）／Analogy（寫明對照與差異）／Assumed（一律給低／高區間）／Derived（寫明換算）。"
            "ABB 的 AI 資料中心變數讀 Wattonomics（設施電力）；本頁不放 Wattonomics 數值，見 Wat 頁。")
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因", "說明"])
FMT = {"%": "0.0%", "pt": "0.0%", "x": "0.00", "US$B": "#,##0.000", "十億股": "0.000", "CHF": "#,##0.00", "US$": "#,##0.00", "年": "0.00", "CHF/US$": "0.00000", "權重": "0%"}
IR = {"r": 4}
def sec(title):
    IN.cell(IR["r"], 1, title).font = BOLD; IR["r"] += 1
def ov(key, lo, base, hi):
    if key in OVR:
        v = OVR[key]; return {"@lo": lo, "@hi": hi}.get(v, None) if v in ("@lo", "@hi") else float(v)
    return base
def inrow(key, item, unit, lo, base, hi, tag, src, inc="—", note="", yellow=False):
    r = IR["r"]; b = ov(key, lo, base, hi)
    if b is None: b = base
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for col, v in ((4, lo), (5, b), (6, hi)):
        if v is None: continue
        c = IN.cell(r, col, v)
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        if unit in FMT: c.number_format = FMT[unit]
    IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, inc if tag == "Interested-party" else "—"); IN.cell(r, 10, note)
    defname(key, f"In!$E${r}")
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    if yellow:
        for cc in range(1, 11): IN.cell(r, cc).fill = KEY
    IR["r"] += 1

# --- A. 股價與匯率（只用兩個 CSV；2026-11-13 收盤） ---
sec("A. 股價、匯率與市場共識（股價只取 markets/invest_map/data 兩個 CSV；共識與同業倍數只作對照與 Comparable 輸入）")
inrow("px_chf", "現價（ABBN.SW，2026-11-13 收盤）", "CHF", None, PX_CHF, None, "Verified", "markets/invest_map/data/closes_extra_20261113.csv：ABBN.SW 69.96（CHF）")
inrow("fx", "匯率：每美元兌瑞士法郎（CHF=X，2026-11-13）", "CHF/US$", None, FX, None, "Verified", "markets/invest_map/data/closes_extra_20261113.csv：CHF=X 0.83593", note="報表與現金流為美元；每股價值以本匯率換成瑞士法郎，不預測匯率")
inrow("px", "現價（美元；＝CHF 現價 ÷ 每美元 CHF）", "US$", None, "=px_chf/fx", None, "Derived", "px_chf ÷ fx")
inrow("cons_tp", "市場共識目標價（賣方平均）", "CHF", None, 83.71, None, "Interested-party", "Yahoo quoteSummary 2026-11-17：目標價均值 83.71（27 家）；docs/workorders/20261117_新增五家_Yahoo基本面快照.md",
      "賣方分析師；多數為投行或券商研究，與 ABB 有承銷與顧問業務往來，評等偏買進；只作對照")
inrow("cons_rev26", "共識 FY2026 營收（Yahoo；幣別為美元，與公司報表口徑一致）", "US$B", None, 38.24, None, "Interested-party", "Yahoo 共識 2026-12 期：營收 38.24B（快照註記疑為美元；與 9M 實績＋Q4 指引相符）", "賣方分析師平均，同上")
inrow("cons_eps26", "共識 FY2026 EPS（Yahoo；口徑待核）", "US$", None, 3.244, None, "Interested-party", "Yahoo 共識 2026-12 期：EPS 3.244", "賣方分析師平均，同上；是否含 Robotics 出售收益或為調整後 EPS 未知")
inrow("cons_rev27", "共識 FY2027 營收（Yahoo）", "US$B", None, 43.30, None, "Interested-party", "Yahoo 共識 2027-12 期：營收 43.30B", "賣方分析師平均，同上")
inrow("cons_eps27", "共識 FY2027 EPS（Yahoo；口徑待核）", "US$", None, 3.729, None, "Interested-party", "Yahoo 共識 2027-12 期：EPS 3.729", "賣方分析師平均，同上")
# --- B. 目標價規則 ---
sec("B. 目標價規則（黃底＝Assumed，Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inrow(f"tp_w{s_}", f"目標價情境權重 {s_}（{SCN[s_]}）", "權重", None, w_, None, "Assumed", "工作單預設 A 0.3／B 0.5／C 0.2", note="三者合計應為 100%", yellow=True)
inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "權重", 0.3, 0.5, 0.7, "Assumed", "Andy 2026-10-25：DCF／Comparable 各半", note="Sens 低／高＝0.3／0.7", yellow=True)
inrow("pe_adj", "同業前瞻 P/E 中位數的調整（Sens 用；基準 0）", "x", -2.0, 0.0, 2.0, "Assumed", "Sens：同業中位數 −2／＋2 倍", note="同業只有 4 檔且來源單一（Yahoo），中位數不穩，故留 ±2 倍")
# --- C. Wattonomics 與 AI 資料中心 ---
sec("C. AI 資料中心暴露（Wattonomics 設施電力；ABB 資料中心營收未揭露，拆分為本模型假設）")
inrow("wif_k", "Wattonomics 設施支出（配電＋UPS）係數（1＝直接採用快照；Sens 低／高）", "x", 0.8, 1.0, 1.2, "Assumed", "Wattonomics v0.3.6 快照（sha " + WSHA + "）WIF_Spend_fac_dist＋WIF_Spend_fac_ups；Wat 頁", note="快照無區間；係數只用來測 ABB 對 Wattonomics 的敏感度，不是改 Wattonomics")
inrow("dc25_pct", "FY2025 資料中心占 Electrification 營收（ABB 不揭露；本模型假設）", "%", 0.10, 0.14, 0.18, "Assumed", "ABB 只說資料中心訂單（Q3 2026 起「三位數增長」），不揭露營收；比例為本模型假設", note="校準點；用 Wattonomics 同年設施支出換算 ABB 份額（Wat 頁、PL 頁）")
inrow("dc26_pct", "FY2026E 資料中心占 Electrification 營收（本模型假設）", "%", 0.17, 0.22, 0.27, "Assumed", "同上；Electrification 9M 2026 營收 +20%、訂單 +60%（Q3 2026 新聞稿）", note="FY2026 資料中心營收＝本比例 × Electrification 營收")
inrow("shr27", "ABB 占 Wattonomics 配電＋UPS 設施支出的份額：FY2027", "%", 0.060, 0.075, 0.090, "Assumed", "FY2025 隱含份額約 8%、FY2026E 約 7%（PL 頁：資料中心營收 ÷ Wattonomics 配電＋UPS 支出，Derived）", note="FY2027→FY2032 線性移動；情境 A／C 另加減下列差距")
inrow("shr32", "ABB 占 Wattonomics 配電＋UPS 設施支出的份額：FY2032", "%", 0.050, 0.070, 0.090, "Assumed", "同上", note="市場份額假設；同業有 Vertiv、Eaton、Schneider、Siemens")
inrow("dsh_up", "情境 A：份額加幅", "pt", 0.005, 0.010, 0.015, "Assumed", "Assumed（情境定義）")
inrow("dsh_dn", "情境 C：份額減幅", "pt", 0.010, 0.015, 0.020, "Assumed", "Assumed（情境定義）")
# --- D. 分部成長 ---
sec("D. 分部營收成長（FY2027 與 FY2032 兩端，中間線性；FY2026 為 9M 實績＋Q4 假設）")
G = [("g_el27", "Electrification 非資料中心營收成長 FY2027", 0.06, 0.09, 0.11, "FY2026E 約 +9%（9M 實績＋Q4 假設）；Electrification 在手訂單 9/2026 為 160.07 億美元、年增 83%（Q3 2026 新聞稿）"),
     ("g_el32", "Electrification 非資料中心營收成長 FY2032", 0.035, 0.05, 0.065, "長期工業與電網電氣化成長；Assumed"),
     ("g_mo27", "Motion 營收成長 FY2027", 0.04, 0.06, 0.08, "9M 2026 營收 +11.6%、在手訂單 70.65 億美元（+14%）"),
     ("g_mo32", "Motion 營收成長 FY2032", 0.025, 0.04, 0.055, "Assumed"),
     ("g_au27", "Automation 營收成長 FY2027", 0.04, 0.06, 0.08, "9M 2026 營收 +10.7%、在手訂單 105.67 億美元（+11%）"),
     ("g_au32", "Automation 營收成長 FY2032", 0.025, 0.04, 0.055, "Assumed")]
for k, lab, lo, b, hi, nt in G:
    inrow(k, lab, "%", lo, b, hi, "Assumed", "ABB Q3 2026 新聞稿分部資料（Nasdaq／EQS 附件）；本欄為模型假設", note=nt)
inrow("dg_up", "情境 A：各分部成長加幅（每年）", "pt", 0.010, 0.020, 0.030, "Assumed", "Assumed（情境定義）")
inrow("dg_dn", "情境 C：各分部成長減幅（每年）", "pt", 0.020, 0.030, 0.040, "Assumed", "Assumed（情境定義）")
inrow("cnet", "公司與其他及分部間沖銷（淨）占三分部營收合計", "%", 0.010, 0.014, 0.018, "Derived", "FY2025：(462−930) ÷ 三分部營收合計 33,688＝−1.4%（ABB Q4 2025 financial information）；取正值，模型取負號", note="區間 Assumed")
# --- E. 利潤 ---
sec("E. 利潤率（Operational EBITA ÷ 分部營收；FY2027 與 FY2032 兩端，中間線性）")
M = [("m_el27", "Electrification Operational EBITA 率 FY2027", 0.245, 0.260, 0.270, "9M 2026：24.9%；Q3 2026：25.7%；FY2025：23.5%"),
     ("m_el32", "Electrification Operational EBITA 率 FY2032", 0.245, 0.265, 0.280, "Assumed"),
     ("m_mo27", "Motion Operational EBITA 率 FY2027", 0.180, 0.193, 0.205, "9M 2026：18.9%；Q3 2026：19.4%；FY2025：19.4%"),
     ("m_mo32", "Motion Operational EBITA 率 FY2032", 0.180, 0.198, 0.210, "Assumed"),
     ("m_au27", "Automation Operational EBITA 率 FY2027", 0.140, 0.152, 0.165, "9M 2026：14.9%；Q3 2026：14.8%；FY2025：14.0%"),
     ("m_au32", "Automation Operational EBITA 率 FY2032", 0.145, 0.160, 0.175, "Assumed")]
for k, lab, lo, b, hi, nt in M:
    inrow(k, lab, "%", lo, b, hi, "Assumed", "ABB Q3 2026 新聞稿分部資料；本欄為模型假設", note=nt)
inrow("dm_up", "情境 A：各分部利潤率加幅", "pt", 0.004, 0.008, 0.012, "Assumed", "Assumed（情境定義）")
inrow("dm_dn", "情境 C：各分部利潤率減幅", "pt", 0.012, 0.018, 0.024, "Assumed", "Assumed（情境定義）")
inrow("co_pct", "公司與其他 Operational EBITA 占三分部營收合計（負值＝成本）", "%", -0.018, -0.013, -0.008, "Assumed", "FY2025：−499（−1.5%）；Q3 2026：−132（−1.3%）；含 Robotics 出售後的滯留成本（FY2025 123 百萬美元）", note="9M 2026 僅 −24（含一次性項目），不外推")
inrow("gm", "毛利率（占營收）", "%", 0.395, 0.405, 0.415, "Assumed", "FY2025：41.1%；9M 2026：40.0%；Q3 2026：40.6%（公司原文）", note="營業費用為毛利與營業利益的差額（含其他收入）")
inrow("br_pct", "Operational EBITA → 營業利益（損益表）的橋：收購攤銷與其他，占營收（負值）", "%", -0.017, -0.012, -0.007, "Assumed", "FY2025：6,314→6,047＝−0.8%（含其他收入 318）；FY2024：−2.7%；Q3 2026：−1.9%", note="FY2026 前九個月橋為 −2.8%（含重組與出售相關成本），不外推")
inrow("fin_f", "稅前的業外淨額（利息收入－財務費用＋退休金貸項；每年，US$B）", "US$B", 0.00, 0.12, 0.20, "Assumed", "FY2025：+172 百萬美元（203−86＋55）；FY2026 Q3 因 Rotork 交易避險重估使財務費用增加（非現金）", note="Rotork 收購（2026-07 宣布，未計入基準）")
inrow("tax_f", "有效稅率（預測）", "%", 0.23, 0.245, 0.26, "Assumed", "FY2025：25.2%；FY2024：24.3%（公司原文）")
inrow("nci_pct", "非控制權益占持續營運淨利（ABB India 等上市子公司）", "%", 0.015, 0.020, 0.030, "Assumed", "FY2025：89 ÷ 4,649＝1.9%（公司原文）")
inrow("sh_g", "每年流通股數變動（買回；負值＝減少）", "%", -0.015, -0.010, 0.000, "Assumed", "2026-02 起 20 億美元買回計畫（Q3 2026 買回 5.7 百萬股、約 5.65 億美元）", note="只影響 EPS 與 Comparable；DCF 用評價日股數")
# --- F. FY2026 第四季 ---
sec("F. FY2026 第四季假設（前三季為公司實績；Q4 為模型；公司指引：Q4 可比營收成長與 Q3 相近、Operational EBITA 率季減幅小於歷史）")
Q4 = [("q4g_el", "Q4 Electrification 營收年增", "%", 0.15, 0.20, 0.25, "Q3 2026：+20.6%；Q4 2025 基期 47.02 億美元"),
      ("q4g_mo", "Q4 Motion 營收年增", "%", 0.08, 0.11, 0.13, "Q3 2026：+11.4%"),
      ("q4g_au", "Q4 Automation 營收年增", "%", 0.03, 0.06, 0.09, "Q3 2026：+5.6%"),
      ("q4cnet", "Q4 公司與其他及沖銷（淨）營收（負值）", "US$B", -0.17, -0.15, -0.13, "Q3 2026：−0.169；Q4 2025：−0.153"),
      ("q4m_el", "Q4 Electrification Operational EBITA 率", "%", 0.235, 0.250, 0.260, "Q3 2026：25.7%；Q4 2025：22.6%"),
      ("q4m_mo", "Q4 Motion Operational EBITA 率", "%", 0.170, 0.183, 0.195, "Q3 2026：19.4%；Q4 2025：18.2%"),
      ("q4m_au", "Q4 Automation Operational EBITA 率", "%", 0.130, 0.142, 0.155, "Q3 2026：14.8%；Q4 2025：13.9%"),
      ("q4co", "Q4 公司與其他 Operational EBITA（負值）", "US$B", -0.18, -0.14, -0.10, "Q3 2026：−0.132；Q4 2025：−0.197"),
      ("br_q4", "Q4 Operational EBITA → 營業利益的橋，占營收（負值）", "%", -0.025, -0.019, -0.013, "Q3 2026：−1.9%")]
for k, lab, u, lo, b, hi, nt in Q4:
    inrow(k, lab, u, lo, b, hi, "Assumed", "ABB Q3 2026 新聞稿（2026-10-20）指引與分部資料；本欄為模型假設", note=nt)
inrow("pg_chf", "Q4 房產出售收益（公司指引：約 2.4 億瑞士法郎，不在 Q4 Operational EBITA 率指引內；本模型計入營業利益與 stub 期現金）", "CHF", None, 0.24, None, "Verified", "ABB Q3 2026 新聞稿：Q4 約 CHF 240 million 房產出售正面影響（單位 CHF B）", note="假設全額為營業利益且現金等值；交割與稅負為假設，列 README 待判斷")
inrow("fin_q4", "Q4 業外淨額（US$B）", "US$B", -0.05, 0.03, 0.08, "Assumed", "Q3 2026 因 Rotork 避險重估財務費用增加；Q4 假設回到小幅正值")
inrow("rob_q4", "Q4 停業單位（Robotics）淨利（US$B；出售收益不計入 EPS）", "US$B", 0.030, 0.045, 0.060, "Assumed", "FY2025 停業單位淨利 174 百萬美元（≈ 每季 44）；出售預計 2026 下半年交割", note="出售收益在交割季認列；本模型 EPS 一律不含，價款由 DCF 加入")
inrow("q4_da", "Q4 折舊與攤銷（US$B）", "US$B", None, 0.230, None, "Derived", "Q3 2026 EBITDA 2,075 − 營業利益 1,846＝229 百萬美元（公司原文）", note="stub 期 FCFF 用")
inrow("q4_cx", "Q4 資本支出（US$B）", "US$B", None, 0.280, None, "Assumed", "Q3 2026 資本支出 251 百萬美元；公司：資本支出年約 10 億美元、中期不再增加", note="stub 期 FCFF 用")
inrow("sh_q", "Q3 2026 期末流通股數（十億股；Q4 EPS 與股數起點）", "十億股", None, 1.808, None, "Verified", "ABB Q3 2026 新聞稿：流通股 1,808 百萬股（2026-09-30）")

# --- G. 標準版 DCF（tools/fcff_dcf.py） ---
P = dict(date="2026-12-31", ccy="US$", amt="US$B", shu="十億股",
         rf=(0.051, "FRED DGS10（美國 10 年期公債殖利率）2026-11-12：5.10%（財政部頁面被擋，改用 FRED）", "報表與現金流為美元，折現用美元利率；g＝MIN(3%, rf)", "無風險利率（美元 10 年期公債殖利率）"),
         beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Analogy",
                     "β＝β_u 依 ABB 的 D／E 重新加槓桿；β_u 由 Yahoo 單一來源 β 1.05 去槓桿（Analogy：電氣設備同業沿用；自算產業 β 尚無 ABB 同業組）", "Sens：β_u ± 0.3 各自重新加槓桿"),
         erp=(0.04, 0.05, 0.06, "工作單預設 5%（同 SpaceX、Alphabet）"),
         crp=(0.0, "Assumed", "營收全球分散、折現用美元：0", "未做國家別 CRP；README 待判斷"),
         kd=(0.059, "Assumed", "美元 10 年期 5.10% ＋ 約 0.8pt 利差（A 級工業債）", "負債僅占資本約 5%，影響極小"),
         tax_w=("=tax_f", "Derived", "同預測有效稅率 tax_f", "—"),
         debt=(7.620, "Verified", "ABB Q3 2026 新聞稿：總債務 7,620 百萬美元（2026-09-30；短期 1,157＋長期 6,463）", "—"),
         lease=(0.786, "Verified", "ABB Q4 2025 financial information：營業租賃負債 786 百萬美元（2025-12-31；Q3 2026 未揭露，沿用）", "—"),
         cash=(6.288, "Verified", "ABB Q3 2026 新聞稿：現金 3,810＋短期投資與有價證券 2,478＝6,288 百萬美元（2026-09-30）", "—"),
         inv=(5.0, "Assumed", "Robotics 出售（SoftBank，企業價值約 54 億美元）：2026-11-05 報導中國外資安全審查禁止 SoftBank 收購其中國業務，交易未完成、雙方仍在協商（Robotics247、Supply Chain Digital；Interested-party）。價值仍以 50 億美元計：交割則收稅後淨款，不交割則保留事業，其價值約當出價（Assumed）；Sens 低 35 億（交易失敗且保留事業價值打折）／高 54。FCFF 不含 Robotics", "價款稅負與分離成本為假設；交割與否、時點與保留價值皆列 README 待判斷"),
         mi=(0.559, "Verified", "ABB Q4 2025 financial information：非控制權益 559 百萬美元（2025-12-31，帳面）；Yahoo 2026-09 為 483；ABB India 上市市值未用 CSV 定價", "帳面值；ABB India 市價可能較高（Sens 高＝3.0 十億美元，Assumed；外部事實未檢核），影響約 −2%（README 待判斷）"),
         sh=("=sh_q+0.006", "Derived", "Q3 2026 期末流通股 1,808 百萬＋稀釋約 6 百萬（Yahoo Q3 2026 稀釋加權 1,814）", "—"),
         fc_da=(None, 0.024, None, "Assumed", "Q3 2026 折舊攤銷 229 百萬 ÷ 營收 9,701＝2.4%（公司原文換算）", "—"),
         fc_da_ss=(0.025, "Assumed", "穩態略高於近期（資本支出增加）", "—"),
         fc_cx=(None, 0.026, None, "Assumed", "公司：資本支出年約 10 億美元、中期不再增加；FY2027 營收約 420 億美元 → 約 2.4–2.6%", "—"),
         fc_cx_ss=(0.027, "Assumed", "最後一年須 ≥ 折舊 ×（1＋g）", "—"),
         fc_wc=(0.05, 0.08, 0.11, "Derived", "淨營運資金 2,827 百萬美元（2026-09-30）÷ 近四季營收約 36,600 百萬＝7.7%（公司原文換算）；Sens 5%／11%", "—"),
         stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src="ABB Q3 2026 新聞稿資產負債表（2026-09-30）",
                   frac=("=(DATE(2026,12,31)-DATE(2026,9,30))/365", "Derived", "＝(2026-12-31 − 2026-09-30) ÷ 365", "stub FCFF 用 Q 頁 Q4 模型數"),
                   div=(0.0, "Assumed", "Q4 無已宣告股利；買回以每股價值中性處理，不調整現金與股數", "—")),
         dps=(0.0, "Assumed", "ABB 年度股利 2026-03 已付；評價日（2026-12-31）到 2027-12-31 的股利（約 0.9 瑞士法郎）未扣除，列 README 待判斷", "—"))
dcf_rows = FD.in_rows(P)
LOHI = {"inv_v": (3.5, 5.4), "mi_v": (0.483, 3.0)}   # Robotics 價值 35／54 億美元（35＝交易失敗且保留事業價值打折）；非控制權益：Yahoo 2026-09 帳面 0.483／ABB India 市價假設 3.0（Assumed）
sec("G. 折現率與無槓桿自由現金流（標準版 DCF；tools/fcff_dcf.py；評價日 2026-12-31）")
inrow("beta_de", "債務與租賃 ÷ 權益市值（β 重新加槓桿用；＝(債務＋租賃) ÷ (現價 × 股數)）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "In 頁債務、租賃、現價與股數")
inrow("beta_du", "β_u 敏感度幅度（± ）", "x", None, 0.3, None, "Assumed", "規格預設 ±0.3")
inrow("beta_u", "β_u（去槓桿）＝Yahoo β 1.05 ÷ (1＋(1−稅)×D／E)", "x", None, "=1.05/(1+(1-tax_w)*beta_de)", None, "Analogy", "Yahoo quoteSummary 2026-11-17：ABBN.SW β 1.05（單一來源，資料商）；ABB 自身 β 視為電氣設備同業代理", note="未自算同業 β；列 README 待判斷與建議補進")
for it in dcf_rows:
    if it[0] == "sec":
        sec(it[1]); continue
    key, lab, unit, lo, b, hi, tag, s1, s2 = it
    if key == "beta_de": continue
    if key in LOHI: lo, hi = LOHI[key]
    inrow(key, lab, unit, lo, b, hi, tag, s1, "—", s2, yellow=key in FD.JUDG)
IN.freeze_panes = "D4"
for col, w in zip("ABCDEFGHIJ", [12, 62, 9, 10, 11, 10, 15, 70, 40, 60]): IN.column_dimensions[col].width = w

# =============================================================== Hist
H = wb.create_sheet("Hist")
H["A1"] = ("Hist — 年度實績（公司原文；百萬美元，除每股與比率）。FY2024、FY2025＝Robotics 重編後的持續營運口徑（Robotics 為停業單位）；"
           "Yahoo 欄只作交叉核對（FY2022–FY2023 為未重編、含 Robotics）。差異見 Recon 頁。")
hdr(H, 3, ["項目", "FY2024", "FY2025", "單位", "標記", "來源", "來源誘因", "Yahoo FY2023（未重編）", "Yahoo FY2024", "Yahoo FY2025", "說明"])
SRC25 = "ABB Q4 2025 financial information（library.e.abb.com；2026-02）"
HR = {}
def hrow(key, lab, v24, v25, unit="US$M", tag="Verified", src=SRC25, yh=(None, None, None), note="", fmt="#,##0"):
    r = H.max_row + 1 if H.max_row >= 4 else 4
    H.cell(r, 1, lab)
    for col, v in ((2, v24), (3, v25)):
        if v is None: continue
        c = H.cell(r, col, v); c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    H.cell(r, 4, unit); H.cell(r, 5, tag); H.cell(r, 6, src); H.cell(r, 7, "—")
    for j, v in enumerate(yh):
        if v is not None: H.cell(r, 8 + j, v).number_format = fmt
    H.cell(r, 11, note); HR[key] = r
hrow("rev", "營收", 30583, 33220, yh=(32235, 30583, 33220))
hrow("cogs", "營業成本", 18582, 19580)
hrow("gp", "毛利", 12001, 13640, yh=(11214, 12001, 13640))
hrow("sga", "銷管費用", 5980, 6593)
hrow("rd", "非訂單相關研發", 1268, 1318, yh=(None, None, 1318))
hrow("oth", "其他收入（費用），淨額", -18, 318, note="FY2025 含房產與資產處分利益（公司原文未逐項列）")
hrow("ebit", "營業利益（Income from operations）", 4735, 6047, yh=(4354, 4753, 5729), note="Yahoo 營業利益＝毛利−銷管−研發，不含其他收入（淨額）；差額＝其他收入")
hrow("intinc", "利息與股利收入", 206, 203)
hrow("finexp", "利息與其他財務費用", -74, -86)
hrow("pens", "非營運退休金貸項", 56, 55)
hrow("pti", "稅前利益（持續營運）", 4923, 6219)
hrow("tax", "所得稅費用", 1197, 1570)
hrow("cont", "持續營運稅後淨利", 3726, 4649)
hrow("disc", "停業單位（Robotics）稅後淨利", 226, 174)
hrow("nci", "非控制權益", 17, 89)
hrow("ni", "歸屬 ABB 淨利", 3935, 4734, yh=(3745, 3935, 4734))
hrow("eps", "基本 EPS（美元）", 2.13, 2.59, unit="US$", fmt="0.00", yh=(2.01, 2.13, 2.59))
hrow("shb", "加權平均基本股數（百萬股）", 1844, 1827)
hrow("shd", "加權平均稀釋股數（百萬股）", 1851, 1831, yh=(1867, 1851, 1831))
hrow("da", "折舊與攤銷", 761, 813, yh=(780, 761, 813))
hrow("capex", "資本支出（PP&E 與無形資產）", 799, 1001, yh=(770, 799, 1001))
hrow("ocf", "營業活動現金流", 4675, 5469, yh=(4290, 4675, 5469))
hrow("div", "已付股利", 1769, 1907)
hrow("bb", "買回庫藏股", 1247, 1499)
hrow("etr", "有效稅率", 0.243, 0.252, unit="%", fmt="0.0%")
H.cell(H.max_row + 2, 1, "分部（Operational EBITA 為 ABB 調整後指標；分部營收未沖銷分部間交易）").font = BOLD
hrow("r_el", "Electrification 營收", 15448, 17357)
hrow("r_mo", "Motion 營收", 7787, 8247)
hrow("r_au", "Automation（原 Process Automation）營收", 7692, 8084)
hrow("r_co", "Corporate and Other 營收", 558, 462)
hrow("r_el_x", "分部間沖銷", -902, -930)
hrow("e_el", "Electrification Operational EBITA", 3520, 4081)
hrow("e_mo", "Motion Operational EBITA", 1518, 1600)
hrow("e_au", "Automation Operational EBITA", 1080, 1132)
hrow("e_co", "Corporate and Other Operational EBITA", -546, -499)
hrow("e_grp", "集團 Operational EBITA", 5572, 6314)
hrow("o_grp", "集團訂單", 31482, 36765)
hrow("rob_rev", "停業單位 Robotics 營收", 2267, 2331, note="SoftBank 約 54 億美元企業價值；2026-11-05 報導中國禁止其中國業務收購，交易未完成、協商中")
hrow("rob_ebit", "停業單位 Robotics 營業利益", 331, 294)
H.cell(H.max_row + 2, 1, "資產負債表（2025-12-31 與 2026-09-30）").font = BOLD
hdr(H, H.max_row + 1, ["項目", "2025-12-31", "2026-09-30", "單位", "標記", "來源", "來源誘因", "Yahoo 2025-12", "Yahoo 2026-09", "", "說明"])
SRCQ3 = "ABB Q3 2026 新聞稿（2026-10-20；Nasdaq／EQS 附件，瑞典語譯本 mb.cision.com）"
def brow(key, lab, a, b, src, yh=(None, None), note=""):
    r = H.max_row + 1
    H.cell(r, 1, lab)
    for col, v in ((2, a), (3, b)):
        if v is not None:
            c = H.cell(r, col, v); c.font = BLUE; c.number_format = "#,##0"
    H.cell(r, 4, "US$M"); H.cell(r, 5, "Verified"); H.cell(r, 6, src); H.cell(r, 7, "—")
    for j, v in enumerate(yh):
        if v is not None: H.cell(r, 8 + j, v).number_format = "#,##0"
    H.cell(r, 11, note); HR[key] = r
brow("cash", "現金及約當現金", 4640, 3810, SRC25 + "；" + SRCQ3)
brow("sti", "短期投資與有價證券", 1981, 2478, SRC25 + "；" + SRCQ3)
brow("std", "短期債務與一年內到期長期債務", 475, 1157, SRC25 + "；" + SRCQ3)
brow("ltd", "長期債務", 7829, 6463, SRC25 + "；" + SRCQ3)
brow("lease", "營業租賃負債", 786, None, SRC25, yh=(None, None), note="Q3 2026 未揭露")
brow("nci_b", "非控制權益（帳面）", 559, None, SRC25, yh=(559, 483), note="Yahoo 2026-09：483")
brow("eq", "歸屬 ABB 股東權益", 16087, None, SRC25, yh=(16087, 16588))
brow("pen", "淨退休金（資產）／負債（不計入淨負債）", -511, -619, SRC25 + "；" + SRCQ3, note="淨資產；DCF 保守不計入")
brow("nd", "淨負債（總債務 − 現金 − 短期投資；公司定義）", "=B{}+B{}-B{}-B{}".format(HR["std"], HR["ltd"], HR["cash"], HR["sti"]), "=C{}+C{}-C{}-C{}".format(HR["std"], HR["ltd"], HR["cash"], HR["sti"]), "＝總債務 − 現金 − 短期投資（公司 2025-12：1,683；2026-09：1,332）")
H.column_dimensions["A"].width = 52
for c, w in zip("BCDEFGHIJK", [11, 11, 8, 10, 60, 10, 14, 12, 12, 60]): H.column_dimensions[c].width = w
HC = lambda key, col="C": f"Hist!${col}${HR[key]}"

# =============================================================== Q
Q = wb.create_sheet("Q")
Q["A1"] = "Q — FY2026 錨點：前三季（9M）與 Q3 為公司實績，Q4 為模型（In 頁 F 節）。FY2026E＝9M 實績＋Q4 模型。金額：十億美元（US$B）。"
Q["A2"] = "會計年度＝日曆年。Q4 2025 實績由 FY2025 減 9M 2025 推得（公司原文：FY2025 與 Q3 2026 新聞稿的 9M 2025 欄）。"
hdr(Q, 4, ["項目（US$B）", "9M 2026A", "Q3 2026A", "Q4 2025A", "Q4 2026E", "FY2026E", "標記", "來源", "來源誘因", "說明"])
QR = {}
def qrow(key, lab, b, c, d, e, f="sum", tag="Verified", src="ABB Q3 2026 新聞稿（2026-10-20）；Q4 2025 欄＝FY2025 − 9M 2025", fmt="#,##0.000", note=""):
    r = Q.max_row + 1; QR[key] = r
    Q.cell(r, 1, lab)
    for col, v in zip("BCDE", (b, c, d, e)):
        if v is None: continue
        cc = Q[f"{col}{r}"]; cc.value = v; cc.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): cc.font = BLUE
    if f == "sum": Q[f"F{r}"] = f"=B{r}+E{r}"
    elif f: Q[f"F{r}"] = f
    Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, src); Q.cell(r, 9, "—"); Q.cell(r, 10, note)
qrow("rev_el", "Electrification 營收", 15.237, 5.424, 4.702, "=D{r}*(1+q4g_el)".format(r=5))
qrow("rev_mo", "Motion 營收", 6.679, 2.320, 2.260, "=D{r}*(1+q4g_mo)".format(r=6))
qrow("rev_au", "Automation 營收", 6.466, 2.126, 2.243, "=D{r}*(1+q4g_au)".format(r=7))
qrow("rev_co", "公司與其他及沖銷（淨）營收", "=B9-B5-B6-B7", "=C9-C5-C6-C7", "=D9-D5-D6-D7", "=q4cnet", tag="Derived", src="＝集團營收 − 三分部營收（公司原文：9M 27,910、Q3 9,701、Q4 2025 9,052）")
qrow("rev_grp", "集團營收", 27.910, 9.701, 9.052, "=E5+E6+E7+E8", tag="Verified")
Q.cell(QR["rev_grp"], 10, "Q4 2026E＝三分部＋公司與其他及沖銷；Q4 2025 實績 9,052")
qrow("eb_el", "Electrification Operational EBITA", 3.800, 1.394, 1.062, "=E5*q4m_el")
qrow("eb_mo", "Motion Operational EBITA", 1.260, 0.449, 0.412, "=E6*q4m_mo")
qrow("eb_au", "Automation Operational EBITA", 0.964, 0.315, 0.311, "=E7*q4m_au")
qrow("eb_co", "公司與其他 Operational EBITA", -0.024, -0.132, -0.197, "=q4co", note="9M 2026 僅 −24：含一次性利益（公司原文未逐項列）")
qrow("eb_grp", "集團 Operational EBITA", 6.000, 2.026, 1.588, "=SUM(E10:E13)")
qrow("pg", "房產出售收益（Q4；美元）", None, None, None, "=pg_chf/fx", tag="Derived", src="pg_chf ÷ fx（In 頁）", note="公司指引約 CHF 240 million；不在 Q4 率指引內")
qrow("ebit", "營業利益（Income from operations）", 5.211, 1.846, "=Hist!C{}/1000-4.542".format(HR["ebit"]), "=E14+br_q4*E9+E15", note="Q4 2026E＝Operational EBITA ＋ 橋 ＋ 房產收益；9M 2025＝4,542（Q3 2026 新聞稿）")
qrow("gp", "毛利", 11.172, 3.939, "=Hist!C{}/1000-9.982".format(HR["gp"]), "=gm*E9", note="9M 2025＝9,982（Q3 2026 新聞稿）")
qrow("cont", "持續營運稅後淨利", 3.827, 1.258, "=Hist!C{}/1000-(3.461-0.0)".format(HR["cont"]) if False else None, "=(E16+fin_q4)*(1-tax_f)", note="Q4 2025 持續營運淨利未揭露分季，空白")
qrow("ni", "歸屬 ABB 淨利", 3.736, 1.181, "=Hist!C{}/1000-3.461".format(HR["ni"]), "=E18*(1-nci_pct)+rob_q4", note="Q4 2026E＝持續營運淨利 ×（1 − 非控制權益比）＋ Robotics 淨利（不含出售收益）")
qrow("eps", "基本 EPS（美元）", 2.06, 0.65, 0.70, "=E19/sh_q", fmt="0.000", note="Q4 2026E＝淨利 ÷ Q3 期末股數")
qrow("da", "折舊與攤銷", None, 0.229, None, "=q4_da", f=None, tag="Derived", src="Q3 2026 EBITDA 2,075 − 營業利益 1,846")
qrow("cx", "資本支出", None, 0.251, None, "=q4_cx", f=None, note="Q3 2026 公司原文：251（PP&E 與無形資產，持續營運）")
qrow("fcff4", "Q4 stub 期 FCFF（US$B）＝營業利益 ×（1−稅）＋折舊 − 資本支出（營運資金變動假設 0）", None, None, None, "=E16*(1-tax_f)+E21-E22", f=None, tag="Derived", src="Q 頁 Q4 模型；DCF 頁 stub 期用", note="含房產出售收益（假設現金等值）")
QR["chk"] = Q.max_row + 1
Q.cell(QR["chk"], 1, "檢查：9M 分部營收與公司合計 − 集團營收（應為 0）"); Q[f"B{QR['chk']}"] = "=B5+B6+B7+B8-B9"; Q[f"C{QR['chk']}"] = "=C5+C6+C7+C8-C9"; Q[f"D{QR['chk']}"] = "=D5+D6+D7+D8-D9"
Q.cell(QR["chk"] + 1, 1, "檢查：分部 Operational EBITA 合計 − 集團（應為 0）"); Q[f"B{QR['chk']+1}"] = "=SUM(B10:B13)-B14"; Q[f"C{QR['chk']+1}"] = "=SUM(C10:C13)-C14"; Q[f"D{QR['chk']+1}"] = "=SUM(D10:D13)-D14"
assert QR["rev_el"] == 5 and QR["rev_grp"] == 9 and QR["eb_el"] == 10 and QR["eb_grp"] == 14 and QR["pg"] == 15 and QR["ebit"] == 16 and QR["cont"] == 18 and QR["ni"] == 19 and QR["da"] == 21 and QR["cx"] == 22, QR
Q.column_dimensions["A"].width = 70
for c, w in zip("BCDEFGHIJ", [11, 11, 11, 11, 11, 12, 70, 10, 60]): Q.column_dimensions[c].width = w
for r in range(5, 25):
    for c in "BCDEF":
        if Q[f"{c}{r}"].value is not None and Q[f"{c}{r}"].number_format == "General": Q[f"{c}{r}"].number_format = "#,##0.000"

# =============================================================== Wat（Wattonomics 讀數）
W = wb.create_sheet("Wat")
W["A1"] = ("Wat — Wattonomics 讀數（設施電力：機櫃外的配電與 UPS；機櫃內屬 IT，不在此）。快照 " + WAT_F + "（v0.3.6；sha256 前 8 碼 " + WSHA + "）。")
W["A2"] = "ABB 資料中心營收＝ABB 份額 × (Wattonomics 配電＋UPS 設施支出) × 係數 wif_k。冷卻與發電機（WIF_Spend_fac_cool、WIF_Spend_fac_gen）不在 ABB 資料中心營收的本模型定義內。"
hdr(W, 4, ["名稱（快照）", "單位"] + [f"FY{y}" for y in YEARS] + ["標記", "來源"])
WR = {}
def wrow(key, nm, unit, r):
    W.cell(r, 1, nm); W.cell(r, 2, unit); WR[key] = r
    for y in YEARS:
        k = f"{nm}_{y}"
        if k in WN and WN[k] is not None:
            c = W[f"{YC[y]}{r}"]; c.value = WN[k]; c.number_format = "#,##0.00"; c.font = BLUE
    W.cell(r, 13, "Derived"); W.cell(r, 14, f"{WAT_F}（Wattonomics v0.3.6）names[{nm}_<yyyy>]；下游只讀快照")
wrow("dist", "WIF_Spend_fac_dist", "US$B", 5); wrow("ups", "WIF_Spend_fac_ups", "US$B", 6)
W.cell(7, 1, "配電＋UPS 設施支出（本模型採用）"); W.cell(7, 2, "US$B"); WR["sum"] = 7
for y in YEARS:
    if f"WIF_Spend_fac_dist_{y}" in WN:
        W[f"{YC[y]}7"] = f"={YC[y]}5+{YC[y]}6"; W[f"{YC[y]}7"].number_format = "#,##0.00"
W.cell(7, 13, "Derived"); W.cell(7, 14, "＝上兩列相加")
wrow("fdist", "WIF_Facility_dist", "US$B/GW", 8); wrow("fups", "WIF_Facility_ups", "US$B/GW", 9); wrow("gw", "WIF_MktGW_Add", "GW", 10)
W.cell(11, 1, "說明：WIF_Facility_dist／ups 為每 IT GW 的配電、UPS 支出（2027 年 2.55、0.75 US$B/GW）；WIF_MktGW_Add 為全市場新增 IT GW。支出＝新增 GW × 每 GW。")
W.column_dimensions["A"].width = 44
for i in range(3, 13): W.column_dimensions[get_column_letter(i)].width = 10
W.column_dimensions["N"].width = 90

# =============================================================== PL
PLW = wb.create_sheet("PL")
PLW["A1"] = ("PL — 三情境損益（GAAP 口徑＝ABB 損益表；Operational EBITA 為 ABB 的 Non-GAAP 調整後指標，另列）。金額：十億美元；EPS：美元。FY2023A 為 Yahoo 未重編（含 Robotics）數，僅供參考；"
             "FY2024A、FY2025A 為 Robotics 重編後持續營運；FY2026E＝前九個月實績＋Q4 模型（Q 頁）；FY2027E 起為情境推導。")
PLW["A2"] = "表頭 FY…A＝實績、FY…E＝模型。三情境在 FY2026E 相同。資料中心營收為本模型拆分（ABB 不揭露），讀 Wattonomics 配電＋UPS 設施支出（Wat 頁）。"
PB = {}
def pl_block(r0, s):
    R = {}; r = r0
    PLW.cell(r, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
    for cc in range(1, 3 + len(YEARS)): PLW.cell(r, cc).fill = SUB
    r += 1; R["hdr"] = r
    for j, h in enumerate(["項目", "單位"] + [f"FY{y}{'A' if y <= 2025 else 'E'}" for y in YEARS], 1):
        c = PLW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    def row(key, lab, unit, fmt):
        nonlocal r
        PLW.cell(r, 1, lab); PLW.cell(r, 2, unit); R[key] = r; r += 1
        return r - 1
    def put(rr, y, f, fmt):
        c = PLW[f"{YC[y]}{rr}"]; c.value = f; c.number_format = fmt
    dg = {"A": "dg_up", "B": "0", "C": "-dg_dn"}[s]; dm = {"A": "dm_up", "B": "0", "C": "-dm_dn"}[s]; dsh = {"A": "dsh_up", "B": "0", "C": "-dsh_dn"}[s]
    NF, PF, XF = "#,##0.000", "0.0%", "0.000"
    rn = row("nd", "電氣化：非資料中心營收", "US$B", NF)
    rd = row("dc", "電氣化：資料中心營收（拆分為本模型假設）", "US$B", NF)
    re_ = row("el", "電氣化營收（Electrification）", "US$B", NF)
    rm = row("mo", "馬達驅動營收（Motion）", "US$B", NF)
    ra = row("au", "自動化營收（Automation）", "US$B", NF)
    rc_ = row("co", "公司與其他及沖銷（淨）", "US$B", NF)
    rr = row("rev", "營收", "US$B", NF)
    ryo = row("yoy", "營收年增率（YoY）", "%", PF)
    rcg = row("cogs", "營業成本（COGS）", "US$B", NF)
    rgp = row("gp", "毛利", "US$B", NF)
    rgm = row("gmr", "毛利率", "%", PF)
    rox = row("opx", "營業費用（銷管、研發與其他收入，淨）", "US$B", NF)
    rop = row("ebit", "營業利益（Income from operations）", "US$B", NF)
    rom = row("ebitm", "營業利益率", "%", PF)
    rfi = row("fin", "其他收入（淨額，業外：利息、財務費用與退休金）", "US$B", NF)
    rpt = row("pti", "稅前利益", "US$B", NF)
    rtx = row("tax", "所得稅", "US$B", NF)
    rnc = row("nci", "非控制權益與停業單位損益（淨；＋停業單位淨利 − 非控制權益）", "US$B", NF)
    rni = row("ni", "稅後淨利（歸屬 ABB）", "US$B", NF)
    rnm = row("nim", "淨利率", "%", PF)
    rsh = row("sh", "股數（十億股；FY2025A 為加權基本）", "十億股", XF)
    rep = row("eps", "EPS（基本；美元）", "US$", "0.000")
    rny = row("niy", "淨利年增率（YoY）", "%", PF)
    PLW.cell(r, 1, "Operational EBITA（ABB 調整後 Non-GAAP 指標）與資料中心暴露").font = BOLD; r += 1
    rbe = row("be_el", "電氣化 Operational EBITA", "US$B", NF); rbm = row("be_mo", "馬達驅動 Operational EBITA", "US$B", NF)
    rba = row("be_au", "自動化 Operational EBITA", "US$B", NF); rbc = row("be_co", "公司與其他 Operational EBITA", "US$B", NF)
    rbg = row("be_g", "集團 Operational EBITA", "US$B", NF); rbr = row("be_r", "集團 Operational EBITA 率", "%", PF)
    rbl = row("be_elr", "電氣化 Operational EBITA 率", "%", PF)
    rbr2 = row("br", "橋：營業利益 − Operational EBITA（收購攤銷與其他）", "US$B", NF)
    rw = row("wat", "Wattonomics 配電＋UPS 設施支出（Wat 頁；US$B）", "US$B", "#,##0.00")
    rsp = row("shr", "ABB 資料中心營收占 Wattonomics 配電＋UPS 設施支出（份額）", "%", PF)
    rdp = row("dcp", "資料中心占電氣化營收", "%", PF)
    rdg = row("dcg", "資料中心占集團營收（AI 暴露度）", "%", PF)
    rot = row("oth", "非電氣化營收（馬達驅動＋自動化＋公司與其他及沖銷）", "US$B", NF)
    rck = row("chk", "檢查：分部合計 − 合併營收（應為 0）", "US$B", "0.000000")
    for y in YEARS:
        c = YC[y]; p = YC.get(y - 1)
        # --- 歷史與錨點 ---
        if y == 2023:
            for k_, v_ in ((rr, 32.235), (rgp, 11.214), (rni, 3.745), (rep, 2.01)):
                PLW[f"{c}{k_}"] = v_; PLW[f"{c}{k_}"].font = BLUE
            PLW[f"{c}{rr}"].number_format = NF; PLW[f"{c}{rgp}"].number_format = NF; PLW[f"{c}{rni}"].number_format = NF; PLW[f"{c}{rep}"].number_format = "0.000"
            put(rgm, y, f"={c}{rgp}/{c}{rr}", PF); put(rnm, y, f"={c}{rni}/{c}{rr}", PF)
            continue
        if y in (2024, 2025):
            hc = "B" if y == 2024 else "C"; h = lambda k: f"Hist!{hc}{HR[k]}/1000"
            put(re_, y, f"=({h('r_el')})", NF); put(rm, y, f"=({h('r_mo')})", NF); put(ra, y, f"=({h('r_au')})", NF)
            put(rc_, y, f"=({h('r_co')})+({h('r_el_x')})", NF)
            put(rr, y, f"={h('rev')}", NF); put(rcg, y, f"={h('cogs')}", NF); put(rgp, y, f"={h('gp')}", NF)
            put(rop, y, f"={h('ebit')}", NF); put(rox, y, f"={c}{rgp}-{c}{rop}", NF)
            put(rfi, y, f"=({h('intinc')})-({h('finexp')})*-1+({h('pens')})" if False else f"=({h('intinc')})+({h('finexp')})+({h('pens')})", NF)
            put(rpt, y, f"={c}{rop}+{c}{rfi}", NF); put(rtx, y, f"={h('tax')}", NF)
            put(rnc, y, f"=({h('disc')})-({h('nci')})", NF); put(rni, y, f"={c}{rpt}-{c}{rtx}+{c}{rnc}", NF)
            put(rsh, y, f"=Hist!{hc}{HR['shb']}/1000", XF); put(rep, y, f"=Hist!{hc}{HR['eps']}", "0.000")
            put(rbe, y, f"={h('e_el')}", NF); put(rbm, y, f"={h('e_mo')}", NF); put(rba, y, f"={h('e_au')}", NF); put(rbc, y, f"={h('e_co')}", NF)
            if y == 2025:
                put(rd, y, f"=dc25_pct*{c}{re_}", NF); put(rn, y, f"={c}{re_}-{c}{rd}", NF)
        else:
            if y == 2026:
                qf = lambda k: f"=Q!$F${QR[k]}"
                put(re_, y, qf("rev_el"), NF); put(rm, y, qf("rev_mo"), NF); put(ra, y, qf("rev_au"), NF); put(rc_, y, qf("rev_co"), NF)
                put(rr, y, qf("rev_grp"), NF); put(rgp, y, qf("gp"), NF); put(rcg, y, f"={c}{rr}-{c}{rgp}", NF); put(rop, y, qf("ebit"), NF); put(rox, y, f"={c}{rgp}-{c}{rop}", NF)
                put(rni, y, qf("ni"), NF); put(rep, y, qf("eps"), "0.000"); put(rsh, y, f"={c}{rni}/{c}{rep}", XF)
                put(rpt, y, f"=Q!$B${QR['cont']}/(1-tax_f)+Q!$E${QR['cont']}/(1-tax_f)", NF)
                put(rtx, y, f"={c}{rpt}*tax_f", NF); put(rfi, y, f"={c}{rpt}-{c}{rop}", NF); put(rnc, y, f"={c}{rni}-({c}{rpt}-{c}{rtx})", NF)
                put(rbe, y, qf("eb_el"), NF); put(rbm, y, qf("eb_mo"), NF); put(rba, y, qf("eb_au"), NF); put(rbc, y, qf("eb_co"), NF)
                put(rd, y, f"=dc26_pct*{c}{re_}", NF); put(rn, y, f"={c}{re_}-{c}{rd}", NF)
            else:
                k = y - 2027
                gm_ = lambda a, b: f"({a}+({b}-{a})*{k}/5+{dg})"
                put(rn, y, f"={p}{rn}*(1+{gm_('g_el27', 'g_el32')})", NF)
                put(rd, y, f"=MAX(0,shr27+(shr32-shr27)*{k}/5+{dsh})*Wat!{c}$7*wif_k", NF)
                put(re_, y, f"={c}{rn}+{c}{rd}", NF)
                put(rm, y, f"={p}{rm}*(1+{gm_('g_mo27', 'g_mo32')})", NF); put(ra, y, f"={p}{ra}*(1+{gm_('g_au27', 'g_au32')})", NF)
                put(rc_, y, f"=-cnet*({c}{re_}+{c}{rm}+{c}{ra})", NF)
                put(rr, y, f"={c}{re_}+{c}{rm}+{c}{ra}+{c}{rc_}", NF)
                mg = lambda a, b: f"({a}+({b}-{a})*{k}/5+{dm})"
                put(rbe, y, f"={c}{re_}*{mg('m_el27', 'm_el32')}", NF); put(rbm, y, f"={c}{rm}*{mg('m_mo27', 'm_mo32')}", NF); put(rba, y, f"={c}{ra}*{mg('m_au27', 'm_au32')}", NF)
                put(rbc, y, f"=co_pct*({c}{re_}+{c}{rm}+{c}{ra})", NF)
                put(rop, y, f"=SUM({c}{rbe}:{c}{rbc})+br_pct*{c}{rr}", NF)
                put(rgp, y, f"=gm*{c}{rr}", NF); put(rcg, y, f"={c}{rr}-{c}{rgp}", NF); put(rox, y, f"={c}{rgp}-{c}{rop}", NF)
                put(rfi, y, "=fin_f", NF); put(rpt, y, f"={c}{rop}+{c}{rfi}", NF); put(rtx, y, f"={c}{rpt}*tax_f", NF)
                put(rnc, y, f"=-nci_pct*({c}{rpt}-{c}{rtx})", NF); put(rni, y, f"={c}{rpt}-{c}{rtx}+{c}{rnc}", NF)
                put(rsh, y, f"={p}{rsh}*(1+sh_g)" if y > 2027 else f"=sh_q*(1+sh_g)", XF); put(rep, y, f"={c}{rni}/{c}{rsh}", "0.000")
        # --- 共用衍生列（2024 起） ---
        put(rgm, y, f"={c}{rgp}/{c}{rr}", PF); put(rom, y, f"={c}{rop}/{c}{rr}", PF); put(rnm, y, f"={c}{rni}/{c}{rr}", PF)
        if y >= 2025:
            put(ryo, y, f"={c}{rr}/{p}{rr}-1", PF); put(rny, y, f"={c}{rni}/{p}{rni}-1", PF)
        put(rbg, y, f"=SUM({c}{rbe}:{c}{rbc})", NF); put(rbr, y, f"={c}{rbg}/{c}{rr}", PF); put(rbl, y, f"={c}{rbe}/{c}{re_}", PF)
        put(rbr2, y, f"={c}{rop}-{c}{rbg}", NF)
        if y >= 2025:
            put(rw, y, f"=Wat!{c}$7", "#,##0.00"); put(rsp, y, f"={c}{rd}/({c}{rw}*wif_k)", PF)
            put(rdp, y, f"={c}{rd}/{c}{re_}", PF); put(rdg, y, f"={c}{rd}/{c}{rr}", PF)
        put(rot, y, f"={c}{rm}+{c}{ra}+{c}{rc_}", NF)
        put(rck, y, f"={c}{re_}+{c}{rm}+{c}{ra}+{c}{rc_}-{c}{rr}", "0.000000")
    for y in (2024,):
        put(ryo, y, None, PF)
    PLW[f"{YC[2024]}{ryo}"] = "=Hist!B{0}/1000/(32.235)-1".format(HR["rev"]) if False else "n/a（FY2023 未重編）"
    PLW[f"{YC[2024]}{rny}"] = "n/a（FY2023 未重編）"
    R["end"] = r
    return R
r0 = 4
for s in SC:
    PB[s] = pl_block(r0, s); r0 = PB[s]["end"] + 2
PLW.column_dimensions["A"].width = 64; PLW.column_dimensions["B"].width = 8
for y in YEARS: PLW.column_dimensions[YC[y]].width = 11
PLW.freeze_panes = "C4"

# =============================================================== DCF
DW = wb.create_sheet("DCF")
DW["A1"] = ("DCF — 標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；FY2027–FY2032 明確期。"
            "金額：十億美元（US$B）。FY2026 第四季（2026-10-01 到 2026-12-31）用 Q 頁模型的 stub 期 FCFF，加入淨現金。")
DW["A2"] = ("EBIT＝損益表營業利益（PL 頁）：含收購攤銷與其他橋項，不含 Robotics（停業單位）；FY2027 起不含房產出售收益。稅率＝tax_f。折舊、資本支出占營收比例見 In 頁 G 節；"
            "營運資金＝fc_wc × 營收增量。Robotics 出售價款在 In 頁以非營業投資加入淨現金；Rotork 收購與三菱電機 JV 未計入（價值中性假設，見 README）。")
paths = [dict(name=s, label=f"情境 {s}（{SCN[s]}）", rev={y: f"PL!{YC[y]}{PB[s]['rev']}" for y in FY_X}, rev0=f"PL!{YC[2026]}{PB[s]['rev']}",
              ebit={y: f"PL!{YC[y]}{PB[s]['ebit']}" for y in FY_X}, tax={y: "tax_f" for y in FY_X}, stub_fcff=f"Q!$E${QR['fcff4']}") for s in SC]
DW.cell(4, 1, "（資料中心營收的 Wattonomics 讀數見 Wat 頁；情境推導見 PL 頁）").font = Font(italic=True)
R = FD.write_detail(DW, 5, paths, FY_X, 1.0, fmt="#,##0.000", stub=True, title="三情境 FCFF 明細（US$B；期中慣例）")
nc_bs = f"(Hist!$C${HR['cash']}+Hist!$C${HR['sti']})/1000+inv_k*inv_v-(Hist!$C${HR['std']}+Hist!$C${HR['ltd']})/1000-Hist!$B${HR['lease']}/1000-Hist!$B${HR['nci_b']}/1000"
R["end"] = FD.write_checks(DW, R["end"], nc_bs=nc_bs, nc_note="資產負債表（Hist 頁，2026-09-30；租賃與非控制權益為 2025-12-31 帳面）") + 1
DW.column_dimensions["A"].width = 62; DW.column_dimensions["B"].width = 34
for c in "CDEFGH": DW.column_dimensions[c].width = 12

# =============================================================== Val
V = wb.create_sheet("Val")
V["A1"] = ("Val — 估值。以美元計算（報表與現金流為美元），第 0 節以瑞士法郎呈現（現價 CHF；每股價值 × CHF=X 每美元兌瑞士法郎）。"
           "目標價＝DCF 與 Comparable 加權；評等依上檔空間固定規則（五級）。")
V["A2"] = "目標時點為評價日（2026-12-31）後 12 個月；DCF 情境價值以股權成本推到 12 個月後。同業前瞻 P/E 來源單一（Yahoo 2026-11-17），Comparable 只作交叉。"
T = {}
t0 = 4
V.cell(t0, 1, "0. 目標價、現價與投資建議（瑞士法郎）").font = BOLD
hdr(V, t0 + 1, ["項目", "A（順利）", "B（基準）", "C（不如預期）", "說明"])
# 先寫 DCF 彙總（美元）在第 28 列起，再回填第 0 節
S = FD.write_summary(wb, V, 30, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00", title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；美元；明細見 DCF 頁）",
                     beta_note="Yahoo β 1.05 去槓桿後依 ABB 的 D／E 重新加槓桿")
V.cell(S["dcf"], 1, "DCF 目標價（情境加權；美元）"); V.cell(S["v12"], 1, "DCF 情境價值推到 12 個月後（美元；以股權成本推到目標時點後扣期間股利）")
V.cell(S["g"], 1, "永續成長率 g＝MIN(上限, rf)"); V.cell(S["ronic"], 1, "永續期 RONIC＝WACC 加成")
V.cell(S["ronic"], 5, "終值再投資率＝g ÷ RONIC；基準加成 0"); V.cell(S["wacc"], 5, "(E × ke ＋ D × kd ×（1 − t）) ÷ (E ＋ D)；調整項只用於 Sens（基準 0）")
q = S["dcf"] + 2
V.cell(q, 1, "0.2 Comparable（美國、歐洲電氣設備同業前瞻 P/E；Yahoo 2026-11-17；來源單一）").font = BOLD
hdr(V, q + 1, ["同業", "代號", "前瞻 P/E（下一會計年度 EPS）", "日期", "來源", "標記", "來源誘因"])
PEERS = [("Eaton", "ETN", 20.50), ("Schneider Electric", "SU.PA", 19.18), ("nVent Electric", "NVT", 15.56), ("Vertiv", "VRT", 19.11)]
for i, (nm, tk, pe) in enumerate(PEERS, q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, "2026-11-17"); V.cell(i, 5, "Yahoo quoteSummary 同業 forwardPE（docs/workorders/20261117_新增五家_Yahoo基本面快照.md）；價格與 EPS 同幣別")
    V.cell(i, 6, "Interested-party"); V.cell(i, 7, "資料商（Yahoo）彙整賣方共識 EPS；賣方研究與投行業務有利益往來")
p1, pl_ = q + 2, q + 1 + len(PEERS)
V.cell(pl_ + 1, 1, "並列（不進中位數）：Yahoo 的 ABB 前瞻 P/E 18.57 把瑞士法郎股價除以美元 EPS，幣別錯置；正確值見第 3 節"); V.cell(pl_ + 1, 3, 18.57).number_format = "0.0"; V.cell(pl_ + 1, 6, "Interested-party")
q2 = pl_ + 2
def tset(rr, key, label, fs, fmt, note=""):
    V.cell(rr, 1, label); T[key] = rr
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
    if note: V[f"E{rr}"] = note
tset(q2, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{p1}:C{pl_})", None, None], "0.0", "ETN、SU.PA、NVT、VRT 中位數")
tset(q2 + 1, "pmed_a", "同業前瞻 P/E 中位數（含 pe_adj）", [f"=B{q2}+pe_adj", None, None], "0.0", "基準 pe_adj＝0；Sens 用")
tset(q2 + 2, "feps", "12 個月後的前瞻 EPS（情境加權；美元＝FY2028 EPS）", ["=" + "+".join(f"tp_w{s}*PL!{YC[2028]}{PB[s]['eps']}" for s in SC), None, None], "0.000",
     "目標時點 2027-12-31 的下一會計年度＝FY2028；同業倍數以現在的下一會計年度 EPS 計，時點口徑一致")
tset(q2 + 3, "comp_usd", "Comparable 目標價（美元）", [f"=B{q2+1}*B{q2+2}", None, None], "#,##0.00", "同業中位數 × 上列 EPS")
tset(q2 + 4, "tpw_usd", "加權目標價（美元）", [f"=tp_w_dcf*B{S['dcf']}+(1-tp_w_dcf)*B{q2+3}", None, None], "#,##0.00", "DCF 權重 × DCF 目標價＋（1 − DCF 權重）× Comparable 目標價")
defname("tp_w_usd", f"Val!$B${q2+4}")
# 回填第 0 節
r = t0 + 2
tset(r, "px", f"現價（{PX_DATE} 收盤；ABBN.SW，瑞士法郎）", ["=px_chf", None, None], "#,##0.00", "投資地圖資料夾的兩個收盤價 CSV；匯率 CHF=X 0.83593（每美元）"); r += 1
tset(r, "dcf_s", "DCF 情境價值推到 12 個月後（瑞士法郎）", [f"={c}{S['v12']}*fx" for c in "BCD"], "#,##0.00", "美元情境價值（本頁 0.1 節）× 每美元兌瑞士法郎"); r += 1
tset(r, "dcf", "DCF 目標價（情境加權）", [f"=B{S['dcf']}*fx", None, None], "#,##0.00", "瑞士法郎；情境權重見 In 頁"); r += 1
tset(r, "comp", "Comparable 目標價（市場主流同業法）", [f"=B{q2+3}*fx", None, None], "#,##0.00", "瑞士法郎；本模型未做相對倍數調整（ABB 沒有 10 年前瞻倍數歷史）"); r += 1
tset(r, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B{}+(1-tp_w_dcf)*B{}".format(T["dcf"], T["comp"]), None, None], "#,##0.00", "DCF 權重 × DCF 目標價＋（1 − DCF 權重）× Comparable 目標價；瑞士法郎"); r += 1
tset(r, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", "Yahoo 2026-11-17：27 家賣方平均（Interested-party）；瑞士法郎"); r += 1
tset(r, "up", "上檔空間：加權目標價", [f"=B{T['tpw']}/B{T['px']}-1", None, None], "0.0%", "現價＝In 頁現價（2026-11-13 收盤）"); r += 1
tset(r, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp']}/B{T['px']}-1", None, None], "0.0%"); r += 1
tset(r, "tier", "上檔空間級距（六級，供分組）", [f'=IF(B{T["up"]}>0.5,"大於 50%",IF(B{T["up"]}>0.25,"25%–50%",IF(B{T["up"]}>=0,"0%–25%",IF(B{T["up"]}>=-0.25,"0%～−25%",IF(B{T["up"]}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@", "Andy 2026-10-25 14:00"); r += 1
tset(r, "rec", "投資建議", [f'=IF(B{T["up"]}>0.5,"強力買進",IF(B{T["up"]}>0.25,"買進",IF(B{T["up"]}>=-0.25,"中立",IF(B{T["up"]}>=-0.5,"放空","強力放空"))))', None, None], "@",
     "固定規則：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（Andy 2026-10-25 14:17）"); r += 1
tset(r, "band", "價格帶（DCF 保守～樂觀）", [f'=TEXT(D{T["dcf_s"]},"#,##0")&"～"&TEXT(B{T["dcf_s"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後；瑞士法郎）"); r += 1
tset(r, "px_usd", "現價（美元）", ["=px", None, None], "#,##0.00", "＝CHF 現價 ÷ CHF=X"); r += 1
tset(r, "tp_usd", "加權目標價（美元）", [f"=B{q2+4}", None, None], "#,##0.00", "DCF 與 Comparable 加權；美元"); r += 1
assert r < 30, r
# --- 第 3 節 倍數檢查 ---
m0 = q2 + 7
V.cell(m0, 1, "3. 現價隱含倍數（瑞士法郎現價 ÷ 以匯率換成瑞士法郎的 EPS）").font = BOLD
hdr(V, m0 + 1, ["項目", "A", "B", "C", "說明"])
for i, y in enumerate((2026, 2027, 2028)):
    tset(m0 + 2 + i, f"pe{y}", f"現價 ÷ FY{y} EPS（模型）", [f"=px/PL!{YC[y]}{PB[s]['eps']}" for s in SC], "0.0", "現價（美元）÷ EPS（美元）" + ("；FY2026 三情境相同" if y == 2026 else ""))
tset(m0 + 5, "pe27c", "現價 ÷ FY2027 EPS（共識）", ["=px/cons_eps27", None, None], "0.0", "Yahoo 的 ABB 前瞻 P/E 18.57 把 CHF 股價除以 USD EPS；本列為幣別一致的 22 倍上下")
tset(m0 + 6, "pe_yh", "Yahoo 前瞻 P/E（幣別錯置）", ["=18.57", None, None], "0.0", "69.96 CHF ÷ 3.729 USD ≈ 18.8（約）；不採用")
V.column_dimensions["A"].width = 60
for c in "BCD": V.column_dimensions[c].width = 15
V.column_dimensions["E"].width = 80; V.column_dimensions["F"].width = 16

# =============================================================== Sens
SE = wb.create_sheet("Sens")
SE["A1"] = ("Sens — 目標價敏感度：加權目標價（美元；Val 第 0.2 節末列）一次動一個輸入。上段＝DCF 參數與權重（閉式公式，與 DCF 頁明細同式）；"
            "下段＝經營輸入的整本重算快照（abb/sens_abb.py 以環境變數覆寫 In 頁後重建、LibreOffice 重算，結果寫入 abb/sens.json）。")
SE["A2"] = "讀法：相對基準＝該案例加權目標價 ÷ 基準加權目標價 − 1；「—」＝該案例只有單邊。Comparable 在 DCF 參數案例不變（DCF 權重案例除外）。"
SJ = os.path.join(HERE, "sens.json")
SNAP = json.load(open(SJ, encoding="utf-8")) if (os.path.exists(SJ) and not os.environ.get("ABB_SNAP") and not OVR) else None
def build_id():
    h = hashlib.sha256(VER.encode())
    for fn in ("build_abb.py", "sens.json", "segments.json"):
        p = os.path.join(HERE, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    h.update(open(os.path.join(ROOT, WAT_F), "rb").read())
    return f"abb-{VER}-{h.hexdigest()[:10]}"
BID = build_id()
first, last, nxt = FD.write_sens(SE, 6, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${q2+3}", base_ref=f"Val!$B${q2+4}", fmt="#,##0.00",
                                 title="目標價參數敏感度（標準版 DCF；美元）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du")
# 同業倍數案例（閉式）
SE.cell(nxt, 1, "同業前瞻 P/E 中位數 ±2 倍"); SE.cell(nxt, 2, "Comparable＝(中位數 ± 2) × 前瞻 EPS；DCF 不變"); SE.cell(nxt, 8, "pe_adj")
SE.cell(nxt, 3, f"=tp_w_dcf*Val!$B${S['dcf']}+(1-tp_w_dcf)*(Val!$B${q2}+pe_adj_lo)*Val!$B${q2+2}"); SE.cell(nxt, 5, f"=tp_w_dcf*Val!$B${S['dcf']}+(1-tp_w_dcf)*(Val!$B${q2}+pe_adj_hi)*Val!$B${q2+2}")
for col in (3, 5):
    SE.cell(nxt, col).number_format = "#,##0.00"; SE.cell(nxt, col + 1, f"={get_column_letter(col)}{nxt}/G{nxt}-1").number_format = "+0.0%;-0.0%;0.0%"
SE.cell(nxt, 7, f"=Val!$B${q2+4}").number_format = "#,##0.00"
SE_ROWS = list(range(first, last + 1)) + [nxt]
# 快照案例
b0 = nxt + 3
SE.cell(b0 - 1, 1, "經營輸入的整本重算快照（一次動一組輸入；低／高＝In 頁敏感度低／高欄）").font = BOLD
hdr(SE, b0, ["參數", "設定（低／高）", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵"])
rb = b0 + 1
SJ_ROWS = []
if SNAP:
    for cs in SNAP["cases"]:
        k = cs["key"]; lo = SNAP["biz"].get(k, {}).get("低"); hi = SNAP["biz"].get(k, {}).get("高")
        SE.cell(rb, 1, cs["label"]); SE.cell(rb, 2, cs["setting"]); SE.cell(rb, 8, cs["keys"])
        for col, v in ((3, lo), (5, hi)):
            if v is None: SE.cell(rb, col, "—"); SE.cell(rb, col + 1, "—")
            else:
                SE.cell(rb, col, v["tp"]).number_format = "#,##0.00"
                SE.cell(rb, col + 1, f"={get_column_letter(col)}{rb}/G{rb}-1").number_format = "+0.0%;-0.0%;0.0%"
        SE.cell(rb, 7, f"=Val!$B${q2+4}").number_format = "#,##0.00"
        SJ_ROWS.append(rb); rb += 1
SE.cell(rb, 1, "註：快照案例為常數（abb/sens.json），不隨本頁公式重算；改 In 頁後須重跑 abb/sens_abb.py 並重建（否則下列模型狀態＝STALE）。")
rs = rb + 2
SE.cell(3, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(3, 4, BID)
SE.cell(4, 1, "模型狀態（CURRENT＝Sens 快照與本檔輸入一致；STALE＝In 頁已改、須重跑 abb/sens_abb.py 與建置）").font = BOLD
snapb = SNAP["_base"]["tp_usd"] if SNAP else None
SE.cell(4, 4, f'=IF(ISNUMBER(D5),IF(ABS(Val!$B${q2+4}-D5)<=0.000001*MAX(1,ABS(D5)),"CURRENT","STALE"),"STALE")')
SE.cell(5, 1, "快照基準加權目標價（美元；sens.json 的 _base）"); SE.cell(5, 4, snapb); SE["D5"].number_format = "#,##0.00"
SE.column_dimensions["A"].width = 50; SE.column_dimensions["B"].width = 56; SE.column_dimensions["H"].width = 28
for c in "CDEFG": SE.column_dimensions[c].width = 15
SJ_MAP = {cs["key"]: rr for cs, rr in zip(SNAP["cases"], SJ_ROWS)} if SNAP else {k: 1 for k in ("wts", "dc_sh", "wif")}
T.update(first=first, last=last, nxt=nxt, SE_ROWS=SE_ROWS, SJ_ROWS=SJ_ROWS, SJ_MAP=SJ_MAP)

# =============================================================== RevDCF
RV = wb.create_sheet("RevDCF")
D = R["paths"]["B"]; col = R["col"]
RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼。一次只動一個驅動、其餘維持基準（情境 B），反解到「現價」；全為公式，只讀 DCF、Val、PL、In 頁，不影響目標價。")
RV["A2"] = ("讀法：DCF 部分反解「情境 B 的 DCF 價值推到 12 個月後＝現價（美元）」；營業利益率為 FY2027–FY2032 統一加減的百分點；營收倍數為 FY2027 起各年同乘的倍數；"
            "WACC 用網格內插。反解值超出合理範圍時照實列出，不外推。")
hdr(RV, 4, ["項目", "值", "說明"])
rv = {}
def kv(key, lab, f, fmt, note=""):
    r_ = RV.max_row + 1; RV.cell(r_, 1, lab); RV.cell(r_, 2, f).number_format = fmt; RV.cell(r_, 3, note); rv[key] = r_
kv("px", "現價（美元）", "=px", "#,##0.00", "＝CHF 現價 ÷ CHF=X")
kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點")
kv("nc", "淨現金滾到評價日（US$B；含 stub 期 FCFF）", f"=Val!$C${S['ncr']}", "#,##0.000", "Val 0.1 節情境 B 欄")
kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{rv['v0']}*B{rv['sh']}-B{rv['nc']}", "#,##0.000")
kv("ev", "模型企業價值 EV（情境 B；US$B）", f"=DCF!$C${D['ev']}", "#,##0.000")
kv("evx", "EV 需放大為模型的倍數", f"=B{rv['evr']}/B{rv['ev']}", "0.00")
kv("vb", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${D['val']}", "#,##0.00")
RV.cell(RV.max_row + 2, 1, "(a) 現價隱含的營業利益率（FY2027–FY2032 統一加減；其他不變）").font = BOLD
cl = lambda y: col[y]
sl = "+".join(f"DCF!{cl(y)}{D['rev']}*DCF!{cl(y)}$%d*DCF!{cl(y)}{D['df']}" % R["frac"] for y in FY_X)
last_c = cl(FY_X[-1])
kv("slope", "EV 對利益率 1 個百分點的敏感度（US$B／pt）", f"=(1-tax_f)*(({sl})+DCF!{last_c}{D['rev']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^{R['tvt']})/100", "#,##0.000", "FCFF 與終值的 NOPAT 隨利益率線性變動")
kv("dm", "現價隱含的利益率加減（百分點；負值＝現價低於模型）", f"=(B{rv['evr']}-B{rv['ev']})/B{rv['slope']}", "0.0", "＝(所需 EV − 模型 EV) ÷ 敏感度")
kv("m0", "模型：FY2027–FY2032 平均營業利益率（情境 B）", "=AVERAGE(" + ",".join(f"PL!{YC[y]}{PB['B']['ebitm']}" for y in FY_X) + ")", "0.0%")
kv("m1", "現價隱含的平均營業利益率", f"=B{rv['m0']}+B{rv['dm']}/100", "0.0%")
RV.cell(RV.max_row + 2, 1, "(b) 現價隱含的營收規模（FY2027 起各年同乘的倍數；利益率與資本支出比例不變）").font = BOLD
kv("w1", "FY2027 營運資金增量調整項（US$B；乘 fc_wc × FY2026 營收 × 第一年現值因子）", f"=fc_wc*PL!{YC[2026]}{PB['B']['rev']}*DCF!{cl(FY_X[0])}{R['frac']}*DCF!{cl(FY_X[0])}{D['df']}", "#,##0.000")
kv("s", "現價隱含的營收倍數", f"=(B{rv['evr']}+B{rv['w1']})/(B{rv['ev']}+B{rv['w1']})", "0.00", "終值與各年 FCFF 隨營收同比放大")
kv("r27", "模型 FY2027 營收（US$B）", f"=PL!{YC[2027]}{PB['B']['rev']}", "#,##0.0")
kv("r27i", "現價隱含的 FY2027 營收（US$B）", f"=B{rv['r27']}*B{rv['s']}", "#,##0.0")
RV.cell(RV.max_row + 2, 1, "(c) 現價隱含的 WACC（網格內插；情境 B；RONIC 隨 WACC）").font = BOLD
gr = RV.max_row + 1
RV.cell(gr, 1, "WACC 網格"); RV.cell(gr + 1, 1, "情境 B 價值推到 12 個月後（美元）")
grid = [0.05 + 0.005 * i for i in range(17)]
for i, w in enumerate(grid):
    cc = get_column_letter(2 + i); RV[f"{cc}{gr}"] = w; RV[f"{cc}{gr}"].number_format = "0.0%"
    RV[f"{cc}{gr+1}"] = "=" + FD.value_expr(R, "B", w=f"{cc}{gr}"); RV[f"{cc}{gr+1}"].number_format = "#,##0.00"
c0, c1 = "B", get_column_letter(2 + len(grid) - 1)
kv("wm", "網格內索引（價值 ≥ 現價的最後一格）", f"=MATCH(B{rv['px']},B{gr+1}:{c1}{gr+1},-1)", "0", "價值隨 WACC 遞減")
kv("wi", "現價隱含的 WACC（線性內插）", f"=IFERROR(INDEX(B{gr}:{c1}{gr},B{rv['wm']})+(INDEX(B{gr+1}:{c1}{gr+1},B{rv['wm']})-B{rv['px']})/(INDEX(B{gr+1}:{c1}{gr+1},B{rv['wm']})-INDEX(B{gr+1}:{c1}{gr+1},B{rv['wm']}+1))*0.005,\"超出網格\")", "0.00%")
kv("w0", "模型 WACC", "=wacc", "0.00%")
RV.cell(RV.max_row + 2, 1, "(d) Comparable：現價隱含的倍數").font = BOLD
kv("cm", "現價隱含的前瞻 P/E（現價美元 ÷ 12 個月後前瞻 EPS）", f"=px/Val!$B${q2+2}", "0.0", "對照同業中位數")
kv("cp", "同業前瞻 P/E 中位數", f"=Val!$B${q2}", "0.0")
RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 60

# =============================================================== Recon
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 對帳：公司原文（主）對 Yahoo（交叉核對）與市場共識；差異即列。單位：百萬美元（除每股與倍數）。"
hdr(RC, 3, ["項目", "公司原文（主）", "Yahoo／共識／模型", "差異", "出處與口徑", "等級"])
def rec(lab, a, b, src, grade="Verified", fmt="#,##0", rel=False):
    r_ = RC.max_row + 1; RC.cell(r_, 1, lab)
    RC.cell(r_, 2, a).number_format = fmt; RC.cell(r_, 3, b).number_format = fmt
    RC.cell(r_, 4, f"=C{r_}-B{r_}" if not rel else f"=C{r_}/B{r_}-1").number_format = fmt if not rel else "0.0%"
    RC.cell(r_, 5, src); RC.cell(r_, 6, grade)
h = lambda key, col: f"=Hist!{col}{HR[key]}"
rec("FY2024 營收（Robotics 重編後）", h("rev", "B"), h("rev", "I"), "公司 FY2024 30,583＝Yahoo 30,583；工作單註記的「2024 年報 32,850」為未重編（含 Robotics）數，兩者相差 Robotics 2,267")
rec("FY2025 營收", h("rev", "C"), h("rev", "J"), "兩者一致")
rec("FY2025 營業利益", h("ebit", "C"), h("ebit", "J"), "Yahoo 營業利益＝毛利−銷管−研發，不含「其他收入（淨額）」318；公司損益表 Income from operations 含之")
rec("FY2024 營業利益", h("ebit", "B"), h("ebit", "I"), "差額＝其他收入（淨額）−18")
rec("FY2025 歸屬淨利", h("ni", "C"), h("ni", "J"), "一致（含停業單位 174）")
rec("FY2025 稀釋股數（百萬）", h("shd", "C"), h("shd", "J"), "一致")
rec("FY2025 資本支出", h("capex", "C"), h("capex", "J"), "一致")
rec("FY2023 營收：Yahoo 未重編（含 Robotics）", None, h("rev", "H"), "公司原文只取得 FY2024 與 FY2025 的重編數，FY2023 重編數未取得；PL 頁 FY2023A 僅供參考", "Interested-party")
rec("非控制權益（帳面）：2025-12 公司 559 對 2026-09 Yahoo 483", h("nci_b", "B"), h("nci_b", "I"), "日期不同（Q3 2026 新聞稿未揭露）；ABB India 上市市值未用 CSV 定價，用帳面值；ABB India 市價高於帳面時影響約 −2%（Sens 案例）")
rec("Q3 2026 營業利益：公司 1,846 對 Yahoo 1,856", 1846, 1856, "Yahoo 季營業利益不含其他收入（淨額），同 FY2025 差異原因", "Interested-party")
rec("FY2026 營收：本模型對共識（US$B）", "=cons_rev26", f"=PL!{YC[2026]}{PB['B']['rev']}", "本模型＝9M 實績 27.910＋Q4 模型；共識 38.24（Yahoo；疑為美元，與 9M 實績加 Q4 指引相符）", "Interested-party", "#,##0.00")
rec("FY2026 EPS：本模型對共識（美元）", "=cons_eps26", f"=PL!{YC[2026]}{PB['B']['eps']}", "共識 3.244 − 9M 實績 2.06＝隱含 Q4 EPS 1.18，高於本模型 Q4（含房產收益）；差異可能來自 Robotics 出售收益或共識為調整後口徑，口徑待核", "Interested-party", "0.000")
rec("FY2026 Operational EBITA 率較 FY2025 的改善（公司指引：至少 +270 bps；含 Q4 房產收益）", 0.027, f"=(PL!{YC[2026]}{PB['B']['be_g']}+Q!$E${QR['pg']})/PL!{YC[2026]}{PB['B']['rev']}-PL!{YC[2025]}{PB['B']['be_r']}",
    "公司 Q3 2026 新聞稿：FY2026 reported Operational EBITA 率至少改善 270 bps（Verified 指引）；本模型值為 Operational EBITA 加房產收益後的率，達標", "Verified", "0.0%")
rec("FY2027 營收：本模型情境 B 對共識（US$B）", "=cons_rev27", f"=PL!{YC[2027]}{PB['B']['rev']}", "共識 43.30（Yahoo）", "Interested-party", "#,##0.00")
rec("FY2027 EPS：本模型情境 B 對共識（美元）", "=cons_eps27", f"=PL!{YC[2027]}{PB['B']['eps']}", "共識 3.729（Yahoo）；本模型含收購攤銷與其他橋項（GAAP），共識口徑待核", "Interested-party", "0.000")
rec("市場共識目標價（瑞士法郎）對本模型加權目標價", "=cons_tp", f"=Val!$B${T['tpw']}", "Yahoo 賣方平均 83.71（27 家）；本模型為 DCF 與 Comparable 各半", "Interested-party", "#,##0.00")
rec("Yahoo 前瞻 P/E 18.57 對幣別一致的現價 ÷ 共識 FY2027 EPS", 18.57, f"=Val!$B${m0+5}", "Yahoo 把 CHF 股價除以 USD EPS；改用同幣別為約 22 倍", "Interested-party", "0.0")
RC.column_dimensions["A"].width = 66
for c, w in zip("BCDEF", [16, 18, 12, 110, 16]): RC.column_dimensions[c].width = w

# =============================================================== Coord
CO = wb.create_sheet("Coord")
CO["A1"] = "Coord — 三層基準：本模型讀哪一層（ABB 為電氣設備廠，AI 資料中心變數讀 Wattonomics 設施電力，不讀 Siliconomics；設施供應商例外）。"
CO["A2"] = "狀態：另一側讀數不適用（電力與機房範圍定案 2026-11-17）：機房設施供應商只讀 Wattonomics 廠房拆分，該數字只有 Wattonomics 一個權威層、沒有 Siliconomics 側的另一讀數，三層協調 C4 兩讀數不適用；與 Wattonomics 的差異走對帳單。本頁不含對帳中判定。"
hdr(CO, 4, ["讀取項目（Wattonomics 快照名稱）", "版本", "sha256 前 8 碼", "用在哪裡", "FY2025", "FY2026", "FY2027", "FY2030", "單位", "標記"])
for i, (nm, use) in enumerate((("WIF_Spend_fac_dist（配電）", "Wat 頁第 5 列 → PL 頁資料中心營收"), ("WIF_Spend_fac_ups（UPS）", "Wat 頁第 6 列 → PL 頁資料中心營收"),
                               ("配電＋UPS（本模型採用）", "Wat 頁第 7 列；ABB 份額分母"), ("WIF_MktGW_Add（全市場新增 IT GW）", "對照（Wat 頁第 10 列）"))):
    r_ = 5 + i; CO.cell(r_, 1, nm); CO.cell(r_, 2, "Wattonomics v0.3.6"); CO.cell(r_, 3, WSHA); CO.cell(r_, 4, use)
    wr = [5, 6, 7, 10][i]
    for j, y in enumerate((2025, 2026, 2027, 2030)):
        CO.cell(r_, 5 + j, f"=Wat!{YC[y]}{wr}").number_format = "#,##0.00"
    CO.cell(r_, 9, "GW" if wr == 10 else "US$B"); CO.cell(r_, 10, "Derived")
CO.cell(10, 1, "邊界：機櫃外的設施（配電、UPS）屬 ABB 範圍；機櫃內屬 IT。冷卻與發電機（ABB 非主要供應商）不納入。ABB 資料中心營收只是 Electrification 的一部分，其餘事業用公司成長假設（Assumed，In 頁 D、E 節）。")
CO.cell(11, 1, "ABB 份額＝資料中心營收 ÷ (配電＋UPS) 為 Derived（PL 頁），區間與敏感度見 In 頁的份額與係數及 Sens。")
CO.cell(12, 1, "尚未做：Wattonomics 另一側讀數、對帳中旗標。建議補進 Wattonomics：以 ABB 為樣本的「機櫃外設施供應商份額」對照（見 README）。")
CO.column_dimensions["A"].width = 56; CO.column_dimensions["D"].width = 50

# =============================================================== Sources
SO = wb.create_sheet("Sources")
SO["A1"] = "Sources — 來源與等級（公司原文以網頁擷取工具讀取，表格錯位處以算術核對；頁碼待人工複核）"; SO["A1"].font = BOLD
hdr(SO, 3, ["來源", "內容", "等級", "日期", "網址或檔案"])
for i, row_ in enumerate([
    ("ABB Q4 2025 financial information", "FY2025／FY2024 損益表（持續營運重編）、分部、現金流、資產負債表；Robotics 停業單位", "Verified（公司原文）", "2026-02", "library.e.abb.com/public/efa424a5e55545ef99275bfafc6ff638/ABB-Q4-2025-financial-information.pdf"),
    ("ABB Q3 2026 results（新聞稿與財務附件）", "Q3／9M 2026：分部、集團、資產負債表、淨負債、股數、買回、指引（Q4、FY2026）、三菱電機 JV、Rotork、資料中心", "Verified（公司原文）", "2026-10-20", "abb.com/global/en/news/139696/q3-2026-results；附件 mb.cision.com/Main/417/4411979/4326359.pdf（瑞典語譯本）"),
    ("Wattonomics 快照 v0.3.6", "WIF_Spend_fac_dist、WIF_Spend_fac_ups、WIF_Facility_*、WIF_MktGW_Add", "Derived（本專案上游層）", "2026-11-16", WAT_F),
    ("markets/invest_map/data 兩個 CSV", "ABBN.SW 69.96 CHF；CHF=X 0.83593（2026-11-13 收盤）", "Verified", "2026-11-13", "closes_20261020_20261113.csv、closes_extra_20261113.csv"),
    ("Yahoo quoteSummary 快照", "共識 EPS／營收／目標價、β 1.05、同業前瞻 P/E；年度與季度基本面（交叉核對）", "Interested-party（資料商）", "2026-11-17", "docs/workorders/20261117_新增五家_Yahoo基本面快照.md"),
    ("FRED DGS10", "美國 10 年期公債殖利率 5.10%", "Verified", "2026-11-12", "fred.stlouisfed.org/series/DGS10"),
], 4):
    for j, v in enumerate(row_, 1): SO.cell(i, j, v)
SO.column_dimensions["A"].width = 40; SO.column_dimensions["B"].width = 90; SO.column_dimensions["E"].width = 90

# =============================================================== README（最後建立，插到最前）
import readme_abb  # noqa: E402
readme_abb.build(wb, dict(VER=VER, BUILD_DATE=BUILD_DATE, PX_DATE=PX_DATE, BID=BID, T=T, S=S, q2=q2, m0=m0, PB=PB, YC=YC, QR=QR, rv=rv, HR=HR, WSHA=WSHA, SJ_MAP=SJ_MAP))
wb.save(sys.argv[1])
