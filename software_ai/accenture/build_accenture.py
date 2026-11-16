#!/usr/bin/env python3
"""Accenture 個股模型 v0.1（原型 L：按人力計價；命題＋官方分部營收 → 損益表 → FCFF DCF＋分部加總 Comparable）。
用法：python3 build_accenture.py <raw.xlsx> [--sens data/sens_v01_20261116.json] [--rowmap rowmap.json] [--rev data/rev_solve_v01.json] [--set key=value ...]
所有計算為活公式；輸入在 In 頁（藍字）。LibreOffice 重算後才有快取值。本版未做先驗。
"""
import datetime, hashlib, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1]
def arg(name): return sys.argv[sys.argv.index(name) + 1] if name in sys.argv else None
SENS, ROWMAP, REVS = arg("--sens"), arg("--rowmap"), arg("--rev")
SETS = {}
for i, a in enumerate(sys.argv):
    if a == "--set":
        k, v = sys.argv[i + 1].split("=", 1); SETS[k] = float(v)
VER, VDATE = "0.1", "20261116"
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKI = TK["items"]


def task_cost(tier):  # 每成功任務推論成本（Task 5：Coding agent 長程，人類 8 小時）＝每次嘗試成本 ÷ 成功率
    g = lambda n: TKI[n]["values"]
    f, c, d = (g(k)[4] for k in ("IF_TaskTokFresh", "IF_TaskTokCached", "IF_TaskTokDec"))
    return (f * g(f"IF_PriceFresh_{tier}") + c * g(f"IF_PriceCached_{tier}") + d * g(f"IF_PriceOut_{tier}")) / 1e6 / g(f"IF_TaskSucc_{tier}")[4]


CPS = {t: task_cost(t) for t in ("Terra", "Sol", "Astra")}
TLEN = TKI["IF_TaskLen"]["values"][4]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
SH = {n: wb.active if n == "README" else wb.create_sheet(n) for n in ["README", "In", "Q", "Seg", "PL", "DCF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
SH["README"].title = "README"

D0 = "Verified"; IP = "Interested-party"; AN = "Analogy"; AS = "Assumed"; DE = "Derived"
K26 = "10-K FY2026（SEC 0001467373-26-000049，申報 2026-10-16）https://www.sec.gov/Archives/edgar/data/1467373/000146737326000049/acn-20260831.htm"
PRQ4 = "8-K 2026-10-01 Q4 FY26 財報新聞稿與法說（https://www.sec.gov/Archives/edgar/data/0001467373/000146737326000037/q4fy26earnings8k.htm）"
RES = "software_ai/accenture/data/research_accenture.md、facts_accenture.csv"
INP = []


def add(key, label, unit, lo, base, hi, tag, src, inc=""):
    if tag == IP and not inc: raise SystemExit(f"{key}: Interested-party 須填來源誘因")
    INP.append((key, label, unit, lo, base, hi, tag, src, inc))


# ---------------- 市場與資本結構
add("px_date", "股價基準日（收盤）", "日期", None, datetime.date(2026, 11, 13), None, D0, "StockAnalysis ACN 歷史股價頁（單一來源）；研究包 facts 擷取日 2026-11-16")
add("px", "現價（2026-11-13 收盤）", "US$/股", None, 250.49, None, D0, "https://stockanalysis.com/stocks/acn/history/（NYSE 收盤；單一來源）")
add("bs_date", "資產負債表日（FY2026 年底）", "日期", None, datetime.date(2026, 8, 31), None, D0, K26 + " 資產負債表")
add("sh_basic", "Class A 流通股數（庫藏股已扣；2026-10-02）", "M股", None, 590.737, None, D0, K26 + " 封面：Class A 流通 668,722,814 股含 77,985,339 股庫藏股，相減 590,737,475")
add("sh_x", "Class X 普通股（可交換股；2026-08-31）", "M股", None, 0.301, None, D0, K26 + " 資產負債表：Class X 300,673 股")
add("rsu", "未歸屬 RSU（2026-08-31；績效型以最高績效計）", "M股", None, 21.110, None, D0, K26 + " Note 股份獎酬 F-40：Nonvested balance 21,109,630")
add("rsu_fx", "RSU 預期沒收比例（以全額計入評價股數再扣沒收）", "比例", 0.0, 0.08, 0.15, DE, "FY2026 沒收 2,370,134 ÷（期初 15,985,074＋新授予 14,289,753）＝7.8%（10-K F-40；Derived）")
add("cash", "現金及約當現金（2026-08-31）", "US$M", None, 12831.471, None, D0, K26 + " 資產負債表")
add("debt", "債務本金（長期債務 9,998.689＋一年內到期 113.069；不含營業租賃）", "US$M", None, 10111.758, None, D0, K26 + " 資產負債表；營業租賃負債不計入淨債務（原則：只計債務本金與融資租賃）；融資租賃未另列（未找到）")
add("nci", "非控制權益（1,157.098）＋可贖回非控制權益（532.762）", "US$M", None, 1689.860, None, D0, K26 + " 資產負債表；自企業價值扣除")
add("dps", "每股股利（年）", "US$/股", None, 6.48, None, DE, "FY2026 股利 4.0B（公司揭露：回購 7.5B＋股利 4.0B；8-K 2026-10-01）÷ 稀釋加權股數 617.5M")
add("rf_date", "無風險利率基準日", "日期", None, datetime.date(2026, 11, 16), None, D0, "https://tradingeconomics.com/united-states/government-bond-yield（2026-11-16 讀取）")
add("rf", "無風險利率（10 年美債）", "比例", 0.045, 0.0512, 0.055, D0, "tradingeconomics 美國 10 年殖利率 5.12%（2026-11-16；單一來源）")
add("beta_raw", "Beta（5 年，單一來源；Blume 調整後用）", "—", 0.9, 1.10, 1.3, D0, "StockAnalysis ACN statistics 頁 Beta (5Y) 1.10（2026-10-31 版；本次只取得一個來源，規格要兩來源，未找到第二來源）")
add("erp", "股權風險溢酬", "比例", 0.04, 0.05, 0.06, AS, "DCF 標準規格預設 5%（4–6%）")
add("kd_sp", "稅前債務信用利差（相對 rf）", "比例", 0.006, 0.010, 0.02, AS, "未找到公司債票面與有效利率；以利差假設")
add("g_tv", "永續成長率", "比例", 0.02, 0.03, 0.04, AS, "DCF 標準規格預設 3%（2–4%）")
add("ronic_sp", "RONIC 減 WACC 的利差", "比例", -0.02, 0.0, 0.02, AS, "DCF 標準規格預設 RONIC＝WACC（新投資價值中性）")
add("tax", "稅率（NOPAT 與損益）", "比例", 0.22, 0.251, 0.27, AS, "FY2026 GAAP 有效稅率 25.1%（10-K MD&A；FY2025 23.7%）；預測期持平於實績，區間為敏感度")
add("w_dcf", "DCF 權重（其餘為 Comparable）", "比例", 0.3, 0.5, 0.7, AS, "原則第 4 節預設 50／50")
add("cm_adj", "Comparable 倍數調整係數（對同業中位數）", "倍", 0.8, 1.0, 1.2, AS, "同業僅 3 家、日期不一致，不做迴歸；以係數作敏感度")
# ---------------- 命題（原型 L）
add("hc_y25", "期末員工數 FY2025", "千人", None, 779, None, IP, "10-K FY2026 MD&A People Metrics（約數）", "公司自揭 KPI，約數")
add("hc_y26", "期末員工數 FY2026", "千人", None, 814, None, IP, "10-K FY2026 MD&A People Metrics；8-K 2026-10-01", "公司自揭 KPI，約數")
add("util", "利用率（FY2026；全預測期持平）", "比例", 0.88, 0.93, 0.95, IP, "10-K MD&A People Metrics 93%（公司定義未揭露細節）；預測期持平", "公司自揭 KPI；定義未揭露")
add("hrs", "每人年可用工時", "小時/年", 1800, 1880, 1960, AS, "unit_economics L_hrs 基準 1,880（約 47 週 × 40 小時）；公司未揭露；與 L_util、L_rate 同名")
add("hg", "AI 前淨人力成長（含併購併入）", "比例/年", 0.010, 0.025, 0.040, AS, "期末員工數年增：FY2023 +1.7%、FY2024 +5.6%、FY2025 +0.7%、FY2026 +4.5%（Derived，平均 3.1%）；基準取 2.5%；對應併購現金 acq27、acq_run")
add("rg", "AI 前每工時價格與組合成長（含工資轉嫁）", "比例/年", 0.000, 0.015, 0.030, AS, "每員工營收 FY2025→FY2026 +1.9%（Derived）；公司稱 FY2026 全年定價 relatively stable（10-K p.31，以合約毛利率定義）；公司不追蹤量價拆分")
add("a27", "AI 自動化工時占比 FY2027（相對 FY2026 基期的增量；情境共同）", "比例", 0.0, 0.02, 0.05, AS, "基期營收已含 FY2026 以前的 AI 影響；10-K 風險因子稱部分任務已被取代並使需求與利用率下降（Verified 文字），無量化；起點取 unit_economics L_a26 區間下緣的一部分")
add("n27", "AI 新專案營收占 AI 前營收 FY2027（情境共同）", "比例", 0.0, 0.015, 0.04, AS, "公司 FY2026 Q1 揭露 Advanced AI 營收約 1.1B／季口述（Interested-party，之後停止揭露）；但為 AI 專屬營收含舊工作轉型，不得相加；起點為假設")
add("tlen", "代理任務長度（人類完成時間）", "小時", None, TLEN, None, D0, "Tokenomics v22.0 IF_TaskLen 第 5 項（8 小時；長程 Coding agent）；快照 data/tokenomics_snapshot_v22.0.json")
add("cps", "每成功任務推論成本（Sol 層級市場價；Terra 為低、Astra 為高）", "US$/件", CPS["Terra"], CPS["Sol"], CPS["Astra"], AN,
    "Tokenomics v22.0：任務 5 每次嘗試 token（新鮮 200,400／快取 1,263,600／decode 30,000；IF_TaskTok*）× 層級有效牌價（IF_PriceFresh／Cached／Out）÷ 成功率（IF_TaskSucc）；公司未揭露 token 量與代理任務量可校準；選哪一層屬類比。高端＝只動成本一邊的上界")
add("px_dec", "每件推論成本逐年變動（FY2028 起）", "比例/年", -0.25, 0.0, 0.0, AS, "價格路徑待 AI供應鏈 登錄；基準持平（保守）；低端每年降 25%")
add("cs", "自動化工時省下的交付成本比例（相對交付成本全額）", "比例", 0.3, 0.6, 1.0, AS, "無揭露；FY2026 員工數與營收同步成長、payroll 占營收由 65.6% 降至 64.5%（Derived）；0.6＝省下的人力一部分被再培訓、閒置與工具成本吃掉")
add("c0_dr", "營業成本率（不含推論）逐年變動（FY2027 起）", "百分點/年", -0.003, 0.0, 0.003, AS, "規則：比率無證據時持平於 FY2026 實績（67.95%）；改善或惡化只放敏感度")
add("opx_dr", "營業費用率逐年變動（FY2027 起）", "百分點/年", -0.003, -0.0011, 0.0, DE, "G3 第 1 條：歷史有規模經濟證據——銷售與行銷＋一般與管理占營收 FY2024 17.15%→FY2025 16.37%→FY2026 16.26%（不含 business optimization；Derived）；基準取最近兩年斜率 −0.11 點／年延伸；高端 0＝全期持平（下緣）；低端＝更快改善")
add("d_em", "EMEA 營收占比逐年變動", "百分點/年", -0.005, 0.0, 0.005, AS, "持平於 FY2026（36.34%）；FY2024→FY2026 +0.9 點（Derived，新口徑）")
add("d_ap", "Asia Pacific 營收占比逐年變動", "百分點/年", -0.003, 0.0, 0.003, AS, "持平於 FY2026（14.39%）")
add("da_pct", "折舊與攤銷占營收", "比例", 0.019, 0.02105, 0.023, DE, "FY2026 1,561.81÷74,183.445＝2.11%（10-K Note 16 分部附註 D&A；持平）")
add("capex_pct", "資本支出占營收", "比例", 0.009, 0.0115, 0.014, AS, "FY2026 1.0%（742.7÷74,183）；公司 FY2027 指引約 0.9B（Q4 FY26 法說）≈ 1.2%")
add("wc_k", "營運資金增加占營收增量比例（負值＝流入）", "比例", -0.05, -0.02, 0.05, DE, "歷史：營業現金流減（淨利＋折舊＋股票薪酬）FY2025 約 +23M、FY2026 約 +178M（流入），除以營收增量 4,776／4,510＝約 −0.5%／−3.9%，平均約 −2%（Derived；第 23 條，注意正負號）")
add("acq27", "FY2027 併購現金支出（FCFF 內）", "US$M", 3000, 5000, 5000, IP, "法說口述：FY2027 預計約 $5B（Q4 FY26 法說；FY2026 實際 4.94B）；另有約 3B（Dragos 相關）由 FY2026Q4 移到 FY2027Q1", "公司口述計畫，未簽約部分可變")
add("acq_run", "FY2028 起每年併購現金支出（FCFF 內）", "US$M", 1500, 3500, 5000, AS, "FY2023–FY2026 併購與投資現金 2.53B／6.58B／1.47B／4.94B（Derived，平均 3.9B）；與 hg 的人力併入配套")
add("oth_net", "業外淨額（利息收入－利息費用＋其他）", "US$M", -150, -30, 100, AS, "FY2026：363.2－303.4－93.6＝−33.7（10-K 損益表，Derived）；債務增加使利息費用上升")
add("nci_pct", "非控制權益占淨利比例", "比例", 0.015, 0.0183, 0.022, DE, "FY2026 NCI 155.8÷淨利 8,522.9＝1.83%（10-K）")
add("sh27", "FY2027 加權稀釋股數（僅 EPS 顯示）", "M股", 590, 600, 615, AS, "流通 590.7M＋RSU 稀釋；FY2026 加權稀釋 617.5M（Verified）但年末流通已降至約 591M；僅影響 EPS 顯示")
add("bb", "FY2028 起每年股數淨減少（僅 EPS 顯示）", "比例/年", 0.0, 0.01, 0.025, AS, "FY2027 股東回饋指引至少 9.5B（IP）；僅影響 EPS 顯示")
SC = {
 "a32": ("AI 自動化工時占比 FY2032（占 AI 前工時；線性自 a27）", "比例", 0.15, 0.20, 0.30, AS, "三情境分化依據：自動化強度。unit_economics L_a32 區間 0.15–0.60（Assumed）；10-K 風險因子稱任務已被取代；CEO 稱 AI 機會大於效率衝擊（Interested-party）", ""),
 "r": ("價格保留率（被自動化工時中仍向客戶收取的比例）", "比例", 0.60, 0.45, 0.25, AS, "分化依據：定價模式。10-K：FY2026 全年定價持平、Q4 多處價格較低；固定價格合約占比逾 60% 且上升（IP）；unit_economics L_r 0.10–0.70", ""),
 "n32": ("AI 新專案營收占 AI 前營收 FY2032（線性自 n27）", "比例", 0.12, 0.08, 0.04, AS, "分化依據：新工作量。公司稱 emerging AI 夥伴訂單成長逾 3 倍、AI 與資料專業人員近 11 萬人（Interested-party）；unit_economics L_n32 0.03–0.25", ""),
 "tp_w": ("情境權重", "比例", 0.25, 0.5, 0.25, AS, "原則第 4 節預設 25／50／25", ""),
}
for k, v in SETS.items():
    for i, row in enumerate(INP):
        if row[0] == k: INP[i] = row[:4] + (v,) + row[5:]

ws = SH["In"]
ws["A1"] = "Accenture 個股模型 v0.1 輸入頁（藍字＝輸入；黑字＝公式）。金額 US$M，會計年度 9 月制（FY2027＝2026-09-01～2027-08-31）。"
ws["A2"] = "基準欄 E 進主模型；D／F（敏感度低／高）只用於 Sens 頁一次動一個參數，不進目標價。情境專屬列（鍵尾 _A／_B／_C）由「情境驅動對照」節並列。本版未做先驗。"
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因"], 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
IR = {}
r = 4
for key, label, unit, lo, base, hi, tag, src, inc in INP:
    IR[key] = r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, unit)
    for col, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = ws.cell(r, col, v); c.font = BLUE
            if isinstance(v, datetime.date): c.number_format = "yyyy-mm-dd"
    ws.cell(r, 7, tag); ws.cell(r, 8, __import__("re").sub(r"https?://\S+", "（網址見 facts 總帳）", src) if "未找到" in src else src); ws.cell(r, 9, inc)
    r += 1
for pre, (label, unit, a, b, c_, tag, src, inc) in SC.items():
    for s, v in zip("ABC", (a, b, c_)):
        k = f"{pre}{s}" if pre == "tp_w" else f"{pre}_{s}"
        IR[k] = r
        ws.cell(r, 1, k); ws.cell(r, 2, f"情境 {s}：{label}"); ws.cell(r, 3, unit)
        ws.cell(r, 5, v).font = BLUE; ws.cell(r, 7, tag); ws.cell(r, 8, __import__("re").sub(r"https?://\S+", "（網址見 facts 總帳）", src) if "未找到" in src else src); ws.cell(r, 9, inc)
        r += 1
r += 1
ws.cell(r, 1, "J. 情境驅動對照（純展示；A 樂觀／B 基準／C 保守）").font = BOLD
r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A", "情境 B", "情境 C", "標記", "說明"], 1): ws.cell(r, j, h).font = BOLD
SCN_H = r; r += 1
SCN_ROWS = {}
for pre, key in (("a32", "scn_a32"), ("r", "scn_r"), ("n32", "scn_n32"), ("tp_w", "scn_w")):
    ks = [f"{pre}{s}" if pre == "tp_w" else f"{pre}_{s}" for s in "ABC"]
    ws.cell(r, 1, key); ws.cell(r, 2, SC[pre][0]); ws.cell(r, 3, SC[pre][1])
    for j, k in zip((4, 5, 6), ks): ws.cell(r, j, f"=E{IR[k]}")
    ws.cell(r, 7, "Derived"); ws.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；三情境並排，只引用上方輸入列")
    SCN_ROWS[key] = r; r += 1


def I(key): return f"In!$E${IR[key]}"


# ---------------- Q 頁
YEARS = ["FY2024", "FY2025", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"]
YC = {y: "CDEFGHIJK"[i] for i, y in enumerate(YEARS)}
HIST = ["FY2024", "FY2025", "FY2026"]
Q = SH["Q"]
Q["A1"] = "Q 頁：公司法定申報歷史數字（US$M；FY2024–FY2026 為會計年度，9 月制）。每列單一標記；Interested-party 列有來源誘因。FY2024 為 10-K FY2026 的重編值（新分部口徑）。"
Q["A2"] = "來源頁碼：分部附註 Note 16 p.F-46、股份獎酬 p.F-40；損益表與資產負債表頁碼未逐頁核對（10-K FY2026 Item 8）。"
for j, h in enumerate(["指標", "FY2024", "FY2025", "FY2026", "單位", "來源", "標記", "來源誘因"], 1): Q.cell(3, j, h).font = BOLD
QR = {}
qdata = [
 ("seg_am", "Americas 營收", 32552.489, 35056.715, 36546.614, "US$M", K26 + " Note 16 分部（p.F-46）", D0, ""),
 ("seg_em", "EMEA 營收", 22817.879, 24643.957, 26958.994, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("seg_ap", "Asia Pacific 營收", 9526.096, 9972.305, 10677.837, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("rev", "合併營收", 64896.464, 69672.977, 74183.445, "US$M", K26 + " 損益表", D0, ""),
 ("oi_am", "Americas 營業利益", 5079.651, 5324.339, 6244.303, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("oi_em", "EMEA 營業利益", 2803.610, 3090.993, 3436.425, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("oi_ap", "Asia Pacific 營業利益", 1712.586, 1810.332, 1725.232, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("cogs", "營業成本（Cost of services）", 43734.147, 47437.576, 50407.847, "US$M", K26 + " 損益表", D0, ""),
 ("sm", "銷售與行銷", 6846.714, 7043.445, 7243.476, "US$M", K26 + " 損益表", D0, ""),
 ("ga", "一般與管理", 4281.316, 4350.968, 4818.621, "US$M", K26 + " 損益表", D0, ""),
 ("bo", "Business optimization costs", 438.440, 615.324, 307.541, "US$M", K26 + " 損益表", D0, ""),
 ("ebit", "營業利益（GAAP）", 9595.847, 10225.664, 11405.960, "US$M", K26 + " 損益表", D0, ""),
 ("intinc", "利息收入", 272.256, 336.324, 363.205, "US$M", K26 + " 損益表", D0, ""),
 ("intexp", "利息費用", -58.969, -228.555, -303.397, "US$M", K26 + " 損益表", D0, ""),
 ("othx", "其他收入（費用）淨額", -109.811, -63.040, -93.553, "US$M", K26 + " 損益表", D0, ""),
 ("ebt", "稅前利益", 9699.323, 10270.393, 11372.215, "US$M", K26 + " 損益表", D0, ""),
 ("tax", "所得稅", 2280.126, 2437.993, 2849.333, "US$M", K26 + " 損益表", D0, ""),
 ("ni_all", "淨利（含非控制權益）", 7419.197, 7832.400, 8522.882, "US$M", K26 + " 損益表", D0, ""),
 ("ni", "歸屬 Accenture plc 淨利", 7264.787, 7678.433, 8367.056, "US$M", K26 + " 損益表", D0, ""),
 ("eps", "稀釋 EPS（GAAP）", 11.44, 12.15, 13.56, "US$/股", K26 + " 損益表", D0, ""),
 ("sh", "稀釋加權股數", 635.940, 632.435, 617.498, "M股", K26 + " 損益表", D0, ""),
 ("da", "折舊與攤銷（分部附註）", 1438.874, 1525.462, 1561.810, "US$M", K26 + " Note 16（p.F-46）", D0, ""),
 ("capex", "資本支出", 516.509, 600.039, 742.748, "US$M", K26 + " 現金流量表", D0, ""),
 ("sbc", "股票薪酬費用", 1941.590, 2093.878, 2100.433, "US$M", K26 + " 現金流量表", D0, ""),
 ("hc", "期末員工數（千人，約）", 774, 779, 814, "千人", K26 + " MD&A People Metrics", IP, "公司自揭 KPI，約數"),
 ("utl", "利用率", 0.92, 0.92, 0.93, "比例", K26 + " MD&A People Metrics", IP, "公司自揭 KPI；定義未揭露"),
]
rr = 4
for key, lab, a, b, c_, unit, src, tag, inc in qdata:
    QR[key] = rr
    Q.cell(rr, 1, lab)
    for j, v in zip((2, 3, 4), (a, b, c_)): Q.cell(rr, j, v).font = BLUE
    Q.cell(rr, 5, unit); Q.cell(rr, 6, src); Q.cell(rr, 7, tag); Q.cell(rr, 8, inc)
    rr += 1
CHKQ = []
for lab, f in (("核對：三分部營收合計－合併營收（應為 0）", "{c}%d+{c}%d+{c}%d-{c}%d" % (QR["seg_am"], QR["seg_em"], QR["seg_ap"], QR["rev"])),
               ("核對：三分部營業利益合計－合併營業利益（應為 0）", "{c}%d+{c}%d+{c}%d-{c}%d" % (QR["oi_am"], QR["oi_em"], QR["oi_ap"], QR["ebit"])),
               ("核對：營收－成本與費用合計－營業利益（應為 0）", "{c}%d-{c}%d-{c}%d-{c}%d-{c}%d-{c}%d" % (QR["rev"], QR["cogs"], QR["sm"], QR["ga"], QR["bo"], QR["ebit"])),
               ("核對：營業利益＋利息與其他－稅前利益（應為 0）", "{c}%d+{c}%d+{c}%d+{c}%d-{c}%d" % (QR["ebit"], QR["intinc"], QR["intexp"], QR["othx"], QR["ebt"])),
               ("核對：稅前利益－所得稅－淨利（含非控制權益）（應為 0）", "{c}%d-{c}%d-{c}%d" % (QR["ebt"], QR["tax"], QR["ni_all"]))):
    Q.cell(rr, 1, lab)
    for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, "=" + f.format(c=col))
    CHKQ.append(rr); rr += 1


def QH(key, y): return f"Q!${'BCD'[HIST.index(y)]}${QR[key]}"


# ---------------- Seg 頁（命題模組；原型 L）
S = SH["Seg"]
S["A1"] = ("命題模組（原型 L：按人力計價）：營收＝AI 前營收 R0（人力 × 利用率 × 可用工時 × 時薪）−（被自動化工時 × 時薪 ×（1−價格保留率））＋ AI 新專案營收；推論成本＝被自動化工時 ÷ 任務長度 × 每成功任務成本（Tokenomics v22.0 快照，不自行估價）。"
           "AI 淨效果＝新專案增量 − 定價侵蝕 − 推論成本。")
S["A2"] = ("變數沿用 software_ai/unit_economics：L_util、L_hrs、L_rate、a（自動化比例）、r（價格保留率）、n（新專案占比）、tier；該模型 v0.1 完成前在本模型 In 頁自建（v0.2 改讀 unit_economics）。"
           "FY2026 基期營收已含既有 AI 影響，a 與 n 皆為相對基期的增量；三情境自 FY2027 起以 a32、r、n32 分化。")
SEGLAB = ["t", "hc0", "rate", "r0", "a", "n", "hauto", "tasks", "cps", "eros", "incr", "inf", "net", "hc_ai", "rev", "sh_em", "sh_ap", "sh_am", "seg_am", "seg_em", "seg_ap", "chk1"]
LABS = {"t": "年序（FY2026＝0）", "hc0": "AI 前有效人力（年平均，千人）", "rate": "混合計費時薪（FY2026 營收反推；US$/小時）", "r0": "AI 前營收 R0（人力 × 利用率 × 可用工時 × 時薪）",
        "a": "自動化工時占比 a（相對 FY2026 基期）", "n": "AI 新專案營收占 R0 比例 n", "hauto": "被自動化工時（百萬小時）", "tasks": "代理任務數（百萬件）", "cps": "每成功任務推論成本（US$；Tokenomics）",
        "eros": "AI 淨效果之一：定價侵蝕（R0 × a ×（1−價格保留率））", "incr": "AI 淨效果之二：新專案營收增量（R0 × n）", "inf": "AI 淨效果之三：推論成本（Tokenomics 市場價）",
        "net": "AI 淨效果（增量 − 侵蝕 − 推論成本）", "hc_ai": "AI 後所需交付人力（千人；人力替代的結果）", "rev": "合併營收（R0 − 侵蝕 + 增量）",
        "sh_em": "EMEA 營收占比", "sh_ap": "Asia Pacific 營收占比", "sh_am": "Americas 營收占比（殘差）", "seg_am": "Americas", "seg_em": "EMEA", "seg_ap": "Asia Pacific",
        "chk1": "檢查：三分部營收合計 − 合併營收"}
SR = {}
for si, s in enumerate("ABC"):
    r0 = 4 + si * (len(SEGLAB) + 4)
    S.cell(r0, 1, f"情境 {s}（{ {'A':'樂觀：自動化低、價格保留高、新專案多','B':'基準','C':'保守：自動化高、價格保留低、新專案少'}[s] }）").font = BOLD
    for y, c in YC.items():
        S[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); S[f"{c}{r0+1}"].font = BOLD
    S.cell(r0 + 1, 1, "US$M 除另註")
    rows = {k: r0 + 2 + i for i, k in enumerate(SEGLAB)}
    SR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items(): S.cell(rw, 1, LABS[k])
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        P = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            S[R("rev")] = f"={QH('rev', y)}"
            for k, q in (("seg_am", "seg_am"), ("seg_em", "seg_em"), ("seg_ap", "seg_ap")): S[R(k)] = f"={QH(q, y)}"
            S[R("sh_em")] = f"={R('seg_em')}/{R('rev')}"; S[R("sh_ap")] = f"={R('seg_ap')}/{R('rev')}"; S[R("sh_am")] = f"={R('seg_am')}/{R('rev')}"
            S[R("chk1")] = f"={R('seg_am')}+{R('seg_em')}+{R('seg_ap')}-{R('rev')}"
            if y == "FY2026":
                S[R("t")] = 0
                S[R("hc0")] = f"=({I('hc_y25')}+{I('hc_y26')})/2"
                S[R("rate")] = f"={R('rev')}*1000000/({R('hc0')}*1000*{I('util')}*{I('hrs')})"
                S[R("r0")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('rate')}/1000000"
            continue
        S[R("t")] = f"={P('t')}+1"
        S[R("hc0")] = f"={P('hc0')}*(1+{I('hg')})"
        S[R("rate")] = f"={P('rate')}*(1+{I('rg')})"
        S[R("r0")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('rate')}/1000000"
        S[R("a")] = f"={I('a27')}+({I('a32_'+s)}-{I('a27')})*({R('t')}-1)/5"
        S[R("n")] = f"={I('n27')}+({I('n32_'+s)}-{I('n27')})*({R('t')}-1)/5"
        S[R("hauto")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('a')}/1000000"
        S[R("tasks")] = f"={R('hauto')}/{I('tlen')}"
        S[R("cps")] = f"={I('cps')}*(1+{I('px_dec')})^MAX({R('t')}-1,0)"
        S[R("eros")] = f"={R('r0')}*{R('a')}*(1-{I('r_'+s)})"
        S[R("incr")] = f"={R('r0')}*{R('n')}"
        S[R("inf")] = f"={R('tasks')}*{R('cps')}"
        S[R("net")] = f"={R('incr')}-{R('eros')}-{R('inf')}"
        S[R("hc_ai")] = f"={R('hc0')}*(1-{R('a')})+{R('incr')}*1000000/({R('rate')}*{I('util')}*{I('hrs')})/1000"
        S[R("rev")] = f"={R('r0')}-{R('eros')}+{R('incr')}"
        S[R("sh_em")] = f"={P('sh_em')}+{I('d_em')}"; S[R("sh_ap")] = f"={P('sh_ap')}+{I('d_ap')}"; S[R("sh_am")] = f"=1-{R('sh_em')}-{R('sh_ap')}"
        S[R("seg_am")] = f"={R('rev')}*{R('sh_am')}"; S[R("seg_em")] = f"={R('rev')}*{R('sh_em')}"; S[R("seg_ap")] = f"={R('rev')}*{R('sh_ap')}"
        S[R("chk1")] = f"={R('seg_am')}+{R('seg_em')}+{R('seg_ap')}-{R('rev')}"
SEG_END = 4 + 3 * (len(SEGLAB) + 4)
S.cell(SEG_END, 1, "註：推論成本取 Tokenomics 市場價（有效牌價）；公司未揭露 token 量與代理任務量（Accenture、Fujitsu 皆然），故每任務 token 取 Tokenomics 任務 5 值（類比），不以聊天基數乘倍數。全成本對照（IF_FullCostDefault）本版未列。").font = BOLD

# ---------------- PL 頁
P = SH["PL"]
P["A1"] = "損益表：官方報告分部三個（Americas、EMEA、Asia Pacific；10-K Note 16），營收與營業利益皆為公司法定附註口徑（GAAP）。US$M。"
P["A2"] = "FY2024A–FY2026A 取自 Q 頁；FY2027E–FY2032E 為情境預測（FY2027 不含 business optimization 費用）。分部營收＝合併營收 × 占比；分部營業利益＝分部營收 ×（FY2026 分部利益率＋全公司共同平移），平移使三分部合計恆等於合併營業利益。"
PLB = ["rev", "c0", "cogs_x", "inf", "cogs", "gp", "opx_r", "opex", "ebit", "da", "capex", "nopat", "dwc", "acq", "fcff", "oth", "ebt", "taxx", "ni", "nci", "ni_own", "shs", "eps", "gate", "kshift", "oi_am", "oi_em", "oi_ap", "chk_oi"]
PLL = {"rev": "營收", "c0": "營業成本率（不含推論）", "cogs_x": "營業成本（不含代理推論）", "inf": "代理推論成本", "cogs": "營業成本", "gp": "毛利", "opx_r": "營業費用率", "opex": "營業費用（銷售與行銷＋一般與管理）",
       "ebit": "營業利益", "da": "折舊與攤銷", "capex": "資本支出", "nopat": "NOPAT（營業利益×(1−稅率)；股票薪酬已在費用內，不加回）", "dwc": "營運資金增加", "acq": "併購現金支出",
       "fcff": "無槓桿自由現金流 FCFF（NOPAT＋折舊−資本支出−營運資金增加−併購）", "oth": "業外淨額", "ebt": "稅前利益", "taxx": "所得稅", "ni": "淨利（含非控制權益）", "nci": "非控制權益",
       "ni_own": "歸屬 Accenture plc 淨利", "shs": "稀釋股數（M）", "eps": "EPS（GAAP 稀釋，美元）", "gate": "收支（FCFF＋稅後業外淨額−股利）", "kshift": "分部利益率共同平移",
       "oi_am": "Americas 營業利益", "oi_em": "EMEA 營業利益", "oi_ap": "Asia Pacific 營業利益", "chk_oi": "檢查：三分部營業利益合計 − 合併營業利益"}
PR = {}
BASE0 = 70
for si, s in enumerate("ABC"):
    r0 = BASE0 + si * (len(PLB) + 4)
    P.cell(r0, 1, f"情境 {s} 損益與現金流").font = BOLD
    for y, c in YC.items():
        P[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); P[f"{c}{r0+1}"].font = BOLD
    rows = {k: r0 + 2 + i for i, k in enumerate(PLB)}
    PR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items(): P.cell(rw, 1, f"情境 {s}｜{PLL[k]}")
    sr = SR[s]
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        Pv = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            P[R("rev")] = f"={QH('rev', y)}"; P[R("cogs")] = f"={QH('cogs', y)}"; P[R("gp")] = f"={R('rev')}-{R('cogs')}"
            P[R("opex")] = f"={QH('sm', y)}+{QH('ga', y)}+{QH('bo', y)}"; P[R("ebit")] = f"={R('gp')}-{R('opex')}"
            P[R("da")] = f"={QH('da', y)}"; P[R("capex")] = f"={QH('capex', y)}"
            P[R("oth")] = f"={QH('ebt', y)}-{QH('ebit', y)}"; P[R("ebt")] = f"={QH('ebt', y)}"; P[R("taxx")] = f"={QH('tax', y)}"
            P[R("ni")] = f"={QH('ni_all', y)}"; P[R("nci")] = f"={QH('ni_all', y)}-{QH('ni', y)}"; P[R("ni_own")] = f"={QH('ni', y)}"
            P[R("shs")] = f"={QH('sh', y)}"; P[R("eps")] = f"={QH('eps', y)}"
            P[R("oi_am")] = f"={QH('oi_am', y)}"; P[R("oi_em")] = f"={QH('oi_em', y)}"; P[R("oi_ap")] = f"={QH('oi_ap', y)}"
            P[R("chk_oi")] = f"={R('oi_am')}+{R('oi_em')}+{R('oi_ap')}-{R('ebit')}"
            continue
        P[R("rev")] = f"=Seg!{c}{sr['rev']}"
        P[R("c0")] = f"=({QH('cogs', 'FY2026')}/{QH('rev', 'FY2026')})+{I('c0_dr')}*Seg!{c}{sr['t']}"
        P[R("cogs_x")] = f"={R('c0')}*(Seg!{c}{sr['r0']}*(1-{I('cs')}*Seg!{c}{sr['a']})+Seg!{c}{sr['incr']})"
        P[R("inf")] = f"=Seg!{c}{sr['inf']}"
        P[R("cogs")] = f"={R('cogs_x')}+{R('inf')}"
        P[R("gp")] = f"={R('rev')}-{R('cogs')}"
        P[R("opx_r")] = f"=(({QH('sm', 'FY2026')}+{QH('ga', 'FY2026')})/{QH('rev', 'FY2026')})+{I('opx_dr')}*Seg!{c}{sr['t']}"
        P[R("opex")] = f"={R('rev')}*{R('opx_r')}"
        P[R("ebit")] = f"={R('gp')}-{R('opex')}"
        P[R("da")] = f"={R('rev')}*{I('da_pct')}"; P[R("capex")] = f"={R('rev')}*{I('capex_pct')}"
        P[R("nopat")] = f"={R('ebit')}*(1-{I('tax')})"
        P[R("dwc")] = f"={I('wc_k')}*({R('rev')}-{Pv('rev')})"
        P[R("acq")] = f"={I('acq27')}" if y == "FY2027" else f"={I('acq_run')}"
        P[R("fcff")] = f"={R('nopat')}+{R('da')}-{R('capex')}-{R('dwc')}-{R('acq')}"
        P[R("oth")] = f"={I('oth_net')}"
        P[R("ebt")] = f"={R('ebit')}+{R('oth')}"; P[R("taxx")] = f"={R('ebt')}*{I('tax')}"; P[R("ni")] = f"={R('ebt')}-{R('taxx')}"
        P[R("nci")] = f"={R('ni')}*{I('nci_pct')}"; P[R("ni_own")] = f"={R('ni')}-{R('nci')}"
        P[R("shs")] = f"={I('sh27')}" if y == "FY2027" else f"={Pv('shs')}*(1-{I('bb')})"
        P[R("eps")] = f"={R('ni_own')}/{R('shs')}"
        P[R("gate")] = f"={R('fcff')}+{R('oth')}*(1-{I('tax')})-{I('dps')}*{R('shs')}"
        P[R("kshift")] = f"=({R('ebit')}-(Seg!{c}{sr['seg_am']}*{QH('oi_am','FY2026')}/{QH('seg_am','FY2026')}+Seg!{c}{sr['seg_em']}*{QH('oi_em','FY2026')}/{QH('seg_em','FY2026')}+Seg!{c}{sr['seg_ap']}*{QH('oi_ap','FY2026')}/{QH('seg_ap','FY2026')}))/{R('rev')}"
        for sg in ("am", "em", "ap"):
            P[R(f"oi_{sg}")] = f"=Seg!{c}{sr['seg_'+sg]}*({QH('oi_'+sg,'FY2026')}/{QH('seg_'+sg,'FY2026')}+{R('kshift')})"
        P[R("chk_oi")] = f"={R('oi_am')}+{R('oi_em')}+{R('oi_ap')}-{R('ebit')}"
PL_END = BASE0 + 3 * (len(PLB) + 4)
b = PR["B"]
STD = ["營收", "營收年增率", "營業成本", "毛利", "毛利率", "營業費用", "營業利益", "營業利益率", "其他收入", "稅前利益", "所得稅", "稅後淨利", "淨利率", "EPS（GAAP 稀釋，美元）", "淨利年增率"]
P.cell(4, 1, "標準損益表（GAAP；US$M，EPS 為美元；情境 B）：FY2024A–FY2026A 實績、FY2027E–FY2032E 預測；稅後淨利＝歸屬 Accenture plc").font = BOLD
for y, c in YC.items():
    P[f"{c}5"] = y + ("A" if y in HIST else "E"); P[f"{c}5"].font = BOLD
P.cell(5, 1, "項目")
STR = {n: 6 + i for i, n in enumerate(STD)}
for n, rw in STR.items(): P.cell(rw, 1, n)
for y, c in YC.items():
    p = chr(ord(c) - 1)
    ref = lambda k: f"{c}{b[k]}"
    P[f"{c}{STR['營收']}"] = f"={ref('rev')}"
    P[f"{c}{STR['營收年增率']}"] = f"={c}{STR['營收']}/{p}{STR['營收']}-1" if y != "FY2024" else "未找到"
    P[f"{c}{STR['營業成本']}"] = f"={ref('cogs')}"; P[f"{c}{STR['毛利']}"] = f"={ref('gp')}"
    P[f"{c}{STR['毛利率']}"] = f"={c}{STR['毛利']}/{c}{STR['營收']}"
    P[f"{c}{STR['營業費用']}"] = f"={ref('opex')}"; P[f"{c}{STR['營業利益']}"] = f"={ref('ebit')}"
    P[f"{c}{STR['營業利益率']}"] = f"={c}{STR['營業利益']}/{c}{STR['營收']}"
    P[f"{c}{STR['其他收入']}"] = f"={ref('oth')}"; P[f"{c}{STR['稅前利益']}"] = f"={ref('ebt')}"; P[f"{c}{STR['所得稅']}"] = f"={ref('taxx')}"
    P[f"{c}{STR['稅後淨利']}"] = f"={ref('ni_own')}"; P[f"{c}{STR['淨利率']}"] = f"={c}{STR['稅後淨利']}/{c}{STR['營收']}"
    P[f"{c}{STR['EPS（GAAP 稀釋，美元）']}"] = f"={ref('eps')}"
    P[f"{c}{STR['淨利年增率']}"] = f"={c}{STR['稅後淨利']}/{p}{STR['稅後淨利']}-1" if y != "FY2024" else "未找到"
P.cell(STR["淨利年增率"] + 1, 1, "FY2024A–FY2026A 營業費用含 business optimization 費用（438.4／615.3／307.5）；預測期不含。FY2024 年增率需 FY2023 同口徑資料，本模型未列（未找到）。")
P.cell(23, 1, "分部營收與營業利益（情境 B；官方報告分部 Note 16）").font = BOLD
SEGP = {"Americas": ("seg_am", "rev"), "EMEA": ("seg_em", "rev"), "Asia Pacific": ("seg_ap", "rev"), "營收合計（三分部）": (None, "rev"),
        "Americas 營業利益": ("oi_am", "oi"), "EMEA 營業利益": ("oi_em", "oi"), "Asia Pacific 營業利益": ("oi_ap", "oi"), "營業利益合計（三分部）": (None, "oi")}
SPR = {}
for i, (lab, (k, kind)) in enumerate(SEGP.items()):
    rw = 24 + i; SPR[lab] = rw; P.cell(rw, 1, lab)
    for y, c in YC.items():
        if k and kind == "rev": P[f"{c}{rw}"] = f"=Seg!{c}{SR['B'][k]}"
        elif k: P[f"{c}{rw}"] = f"={c}{b[k]}"
        elif kind == "rev": P[f"{c}{rw}"] = f"=SUM({c}{rw-3}:{c}{rw-1})"
        else: P[f"{c}{rw}"] = f"=SUM({c}{rw-3}:{c}{rw-1})"
CK = 24 + len(SEGP) + 1
P.cell(CK - 1, 1, "對帳（應為 0）").font = BOLD
CHKL = ["檢查：分部營收合計 − 合併營收", "檢查：分部營業利益合計 − 合併營業利益", "檢查：分部營收合計 − 公司揭露合併營收（歷史年度）", "檢查：營收 − 營業成本 − 營業費用 − 營業利益（歷史年度）", "檢查：三情境分部營收合計 − 合併營收"]
for i, l in enumerate(CHKL): P.cell(CK + i, 1, l)
for y, c in YC.items():
    P[f"{c}{CK}"] = f"={c}{SPR['營收合計（三分部）']}-{c}{STR['營收']}"
    P[f"{c}{CK+1}"] = f"={c}{SPR['營業利益合計（三分部）']}-{c}{b['ebit']}"
    P[f"{c}{CK+2}"] = f"={c}{SPR['營收合計（三分部）']}-{QH('rev', y)}" if y in HIST else 0
    P[f"{c}{CK+3}"] = f"={c}{STR['營收']}-{c}{STR['營業成本']}-{c}{STR['營業費用']}-{c}{STR['營業利益']}" if y in HIST else 0
    P[f"{c}{CK+4}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['chk1']})" for s in "ABC") + ")"
CHK_ROWS = (CK, CK + 4)

# ---------------- DCF 頁
Dp = SH["DCF"]
Dp["A1"] = "DCF（FCFF 標準規格）：FY2027E～FY2032E；永續價值＝NOPAT(FY2033)×(1−g÷RONIC)÷(WACC−g)；期中折現。現金流自資產負債表日（2026-08-31）起算，折現到股價基準日；三情境各算一次。"
Dp["A2"] = "WACC 由 Val 頁推出（讀 In 頁基準欄）。年度起訖日：FY(y)＝(y−1)-09-01～y-08-31。"
DR = {}
KEYS = ["start", "end", "fcff", "frac", "t", "df", "pv"]
KLAB = ["年度起日", "年度迄日", "FCFF", "計入比例（資產負債表日之後）", "折現時點（年；自股價基準日到期中）", "折現因子", "現值"]
for si, s in enumerate("ABC"):
    r0 = 4 + si * 11
    DR[s] = {k: r0 + 1 + i for i, k in enumerate(KEYS)}
    Dp.cell(r0, 1, f"情境 {s}").font = BOLD
    for y in YEARS[3:]:
        Dp[f"{YC[y]}{r0}"] = y + "E"
    for k, lab in zip(KEYS, KLAB): Dp.cell(DR[s][k], 1, f"情境 {s}｜{lab}")
    for y in YEARS[3:]:
        c = YC[y]; yy = int(y[2:])
        Dp[f"{c}{DR[s]['start']}"] = f"=DATE({yy-1},9,1)"; Dp[f"{c}{DR[s]['end']}"] = f"=DATE({yy},8,31)"
        for k in ("start", "end"): Dp[f"{c}{DR[s][k]}"].number_format = "yyyy-mm-dd"
        Dp[f"{c}{DR[s]['fcff']}"] = f"=PL!{c}{PR[s]['fcff']}"
        st, en = f"{c}{DR[s]['start']}", f"{c}{DR[s]['end']}"
        Dp[f"{c}{DR[s]['frac']}"] = f"=MAX(0,MIN(1,({en}-MAX({st}-1,{I('bs_date')}))/({en}-{st}+1)))"
        Dp[f"{c}{DR[s]['t']}"] = f"=MAX(0,((MAX({st},{I('bs_date')}+1)+{en})/2-{I('px_date')})/365)"
        Dp[f"{c}{DR[s]['df']}"] = f"=1/(1+Val!$B$%d)^{c}{DR[s]['t']}"  # 補 WACC 列號
        Dp[f"{c}{DR[s]['pv']}"] = f"={c}{DR[s]['fcff']}*{c}{DR[s]['frac']}*{c}{DR[s]['df']}"
    Dp.cell(DR[s]["pv"] + 1, 1, f"情境 {s}｜永續價值折現時點（年；FY2032 年底）"); Dp[f"K{DR[s]['pv']+1}"] = f"=(K{DR[s]['end']}-{I('px_date')})/365"
    Dp.cell(DR[s]["pv"] + 2, 1, f"情境 {s}｜檢查：FCFF 各列加總（NOPAT＋折舊−資本支出−營運資金增加−併購）−FCFF（應為 0）")
    for y in YEARS[3:]:
        c = YC[y]
        Dp[f"{c}{DR[s]['pv']+2}"] = f"={c}{DR[s]['fcff']}-(PL!{c}{PR[s]['nopat']}+PL!{c}{PR[s]['da']}-PL!{c}{PR[s]['capex']}-PL!{c}{PR[s]['dwc']}-PL!{c}{PR[s]['acq']})"
GATE = 4 + 3 * 11 + 1
Dp.cell(GATE, 1, "收支關卡：自由現金流扣股利（FY2027–FY2031 五年合計；US$M；欄 B／C／D＝情境 A／B／C）")
for col, s in zip("BCD", "ABC"):
    Dp[f"{col}{GATE}"] = f"=SUM(PL!F{PR[s]['gate']}:J{PR[s]['gate']})"
Dp.cell(GATE + 1, 1, "檢查：收支關卡 B − 逐年加總（應為 0）"); Dp[f"B{GATE+1}"] = f"=C{GATE}-(PL!F{PR['B']['gate']}+PL!G{PR['B']['gate']}+PL!H{PR['B']['gate']}+PL!I{PR['B']['gate']}+PL!J{PR['B']['gate']})"

RCR_REV27, RCR_OM27 = 5, 6   # Recon 列號（h1=3、g27=4、rev27=5、om27=6；建置後斷言）
# ---------------- Val 頁
V = SH["Val"]
VR = {}
vrows = ["intro", "hint", "blank0", "cap", "px", "px_date", "tp_date", "dcf_tp", "comp_tp", "tp", "up", "valr", "gate", "rec", "tp_ref", "sp1", "wh", "ke", "kd", "kdat", "wacc", "ronic", "g", "emkt", "dmkt", "ewt",
         "sp2", "ndh", "cashx", "ndebt", "nciv", "netcash", "sh", "dps", "sp3", "sumh", "pvx", "tvpv", "ev", "eq", "vps0", "dcf12", "tvsh", "sp4", "ch", "ttm_am", "ttm_em", "ttm_ap", "ttm_tot", "mult", "ev_c", "comp12",
         "sp8", "g3h", "g3_dcf", "g3_ev", "g3_rate", "g3_rev", "g3_om", "g3_flag", "sp7", "gaph", "gap", "gap_t", "gap_m", "gap_sbc", "gap_lev", "ownh", "own_ev", "own_mult", "own_ratio", "rerate", "sp5", "scenh", "wa", "wb", "wc", "chk_w",
         "sp6", "peerh", "pe1", "pe2", "pe3", "pe_med", "pe_n", "pe_own", "stale_rf", "pe_note"]
for i, k in enumerate(vrows): VR[k] = 1 + i
V.cell(VR["intro"], 1, "估值：DCF（FCFF）與 Comparable（分部加總：各分部目標時點 TTM 營業利益 × 同業 EV／EBIT）加權；目標價為 12 個月後（股價基準日＋12 個月）。三情境 25／50／25。上檔空間不是年化報酬。")
V.cell(VR["hint"], 1, "Comparable 同業只有 3 家（少於 8 家，不做迴歸）、日期不一致；倍數分母為 TTM 營業利益（EV／EBIT，無槓桿對無槓桿），故乘目標日的 TTM 營業利益；不用 FCF 倍數，避免股票薪酬與槓桿口徑落差。")
V.cell(VR["cap"], 1, "結論（情境加權）").font = BOLD


def vrow(k, label, b=None, c=None, d=None, e=None):
    r_ = VR[k]; V.cell(r_, 1, label)
    for col, v in zip("BCDE", (b, c, d, e)):
        if v is not None: V[f"{col}{r_}"] = v


wr = lambda s: I("tp_w" + s)
vrow("px", "現價（US$）", f"={I('px')}")
vrow("px_date", "股價基準日", f"={I('px_date')}"); V[f"B{VR['px_date']}"].number_format = "yyyy-mm-dd"
vrow("tp_date", "目標時點（股價基準日＋12 個月）", f"=EDATE({I('px_date')},12)"); V[f"B{VR['tp_date']}"].number_format = "yyyy-mm-dd"
vrow("dcf_tp", "DCF 目標價（情境加權；US$）", f"={wr('A')}*C{VR['dcf12']}+{wr('B')}*D{VR['dcf12']}+{wr('C')}*E{VR['dcf12']}")
vrow("comp_tp", "Comparable 目標價（分部加總；情境加權；US$）", f"={wr('A')}*C{VR['comp12']}+{wr('B')}*D{VR['comp12']}+{wr('C')}*E{VR['comp12']}")
vrow("tp", "加權目標價（DCF 與 Comparable 加權；US$）", f"={I('w_dcf')}*B{VR['dcf_tp']}+(1-{I('w_dcf')})*B{VR['comp_tp']}")
vrow("up", "上檔空間（相對現價）", f"=B{VR['tp']}/B{VR['px']}-1")
tier = lambda u: f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'
vrow("valr", "估值評等（DCF 與 Comparable 加權，只看上檔空間）", "=" + tier(f"B{VR['up']}"))
vrow("gate", "入不敷出判定用：基準情境 FY2027–FY2031 收支合計（US$M）", f"=DCF!C{GATE}")
vrow("rec", "投資建議", f'=IF(AND(B{VR["gate"]}<0,OR(B{VR["valr"]}="買進",B{VR["valr"]}="強力買進")),"收支不足",B{VR["valr"]})')
vrow("tp_ref", "只看 Comparable 的目標價（上檔空間見右）", f"=B{VR['comp_tp']}", f"=B{VR['comp_tp']}/B{VR['px']}-1")
V.cell(VR["wh"], 1, "折現率（DCF）").font = BOLD
vrow("ke", "股權成本 ke（rf＋β×ERP；β＝Blume 調整＝0.67×原始＋0.33）", f"={I('rf')}+(0.67*{I('beta_raw')}+0.33)*{I('erp')}")
vrow("kd", "稅前債務成本（rf＋利差）", f"={I('rf')}+{I('kd_sp')}")
vrow("kdat", "稅後債務成本", f"=B{VR['kd']}*(1-{I('tax')})")
vrow("emkt", "股權市值（US$M；評價股數 × 現價）", f"={I('px')}*B{VR['sh']}")
vrow("dmkt", "債務本金（US$M）", f"={I('debt')}")
vrow("ewt", "股權權重", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})")
vrow("wacc", "WACC", f"=B{VR['ewt']}*B{VR['ke']}+(1-B{VR['ewt']})*B{VR['kdat']}")
vrow("ronic", "RONIC（＝WACC＋利差）", f"=B{VR['wacc']}+{I('ronic_sp')}")
vrow("g", "永續成長率", f"={I('g_tv')}")
V.cell(VR["ndh"], 1, "淨現金（2026-08-31 資產負債表；US$M；營業租賃不計）").font = BOLD
vrow("cashx", "現金及約當現金", f"={I('cash')}")
vrow("ndebt", "債務本金", f"={I('debt')}")
vrow("nciv", "非控制權益（含可贖回；自股權價值扣除）", f"={I('nci')}")
vrow("netcash", "淨現金減非控制權益（負值＝淨債務）", f"=B{VR['cashx']}-B{VR['ndebt']}-B{VR['nciv']}")
vrow("sh", "評價股數（M；流通＋Class X＋RSU×(1−預期沒收)）", f"={I('sh_basic')}+{I('sh_x')}+{I('rsu')}*(1-{I('rsu_fx')})")
vrow("dps", "期間股利（US$/股；持有期間）", f"={I('dps')}")
V.cell(VR["sumh"], 1, "DCF 彙總（三情境）").font = BOLD
for col, lab in zip("CDE", ("情境 A", "情境 B", "情境 C")): V[f"{col}{VR['sumh']}"] = lab
for si, s in enumerate("ABC"):
    col = "CDE"[si]; pvr = DR[s]["pv"]
    V[f"{col}{VR['pvx']}"] = f"=SUM(DCF!F{pvr}:K{pvr})"
    V[f"{col}{VR['tvpv']}"] = f"=PL!K{PR[s]['nopat']}*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^DCF!K{pvr+1}"
    V[f"{col}{VR['ev']}"] = f"={col}{VR['pvx']}+{col}{VR['tvpv']}"
    V[f"{col}{VR['eq']}"] = f"={col}{VR['ev']}+$B${VR['netcash']}"
    V[f"{col}{VR['vps0']}"] = f"={col}{VR['eq']}/$B${VR['sh']}"
    V[f"{col}{VR['dcf12']}"] = f"={col}{VR['vps0']}*(1+$B${VR['ke']})-$B${VR['dps']}"
    V[f"{col}{VR['tvsh']}"] = f"={col}{VR['tvpv']}/{col}{VR['ev']}"
    # Comparable：目標日 TTM（窗口＝[股價基準日, 目標日]）：FY2027 占 w、FY2028 占 1−w
    w = f"MIN(1,(DCF!$F${DR[s]['end']}-{I('px_date')})/365)"
    for sg, rowk in (("am", "ttm_am"), ("em", "ttm_em"), ("ap", "ttm_ap")):
        V[f"{col}{VR[rowk]}"] = f"={w}*PL!F{PR[s]['oi_'+sg]}+(1-{w})*PL!G{PR[s]['oi_'+sg]}"
    V[f"{col}{VR['ttm_tot']}"] = f"=SUM({col}{VR['ttm_am']}:{col}{VR['ttm_ap']})"
    V[f"{col}{VR['mult']}"] = f"=$B${VR['pe_med']}*{I('cm_adj')}"
    V[f"{col}{VR['ev_c']}"] = f"={col}{VR['ttm_tot']}*{col}{VR['mult']}"
    V[f"{col}{VR['comp12']}"] = f"=({col}{VR['ev_c']}+$B${VR['netcash']})/$B${VR['sh']}"
    V[f"{col}{VR['wb']}"] = f"={I('w_dcf')}*{col}{VR['dcf12']}+(1-{I('w_dcf')})*{col}{VR['comp12']}"
    V[f"{col}{VR['wc']}"] = f"={col}{VR['wb']}/$B${VR['px']}-1"
vrow("pvx", "明確期 FCFF 現值（US$M）"); vrow("tvpv", "永續價值現值（US$M）"); vrow("ev", "企業價值 EV（US$M）"); vrow("eq", "股權價值（EV＋淨現金減非控制權益；US$M）")
vrow("vps0", "每股價值（評價日；US$）"); vrow("dcf12", "DCF 情境目標價（12 個月後＝每股價值×(1+ke)−股利；US$）"); vrow("tvsh", "永續價值占 EV 比例")
V.cell(VR["ch"], 1, "Comparable（分部加總：目標日 TTM 營業利益 × 同業 EV／EBIT〔分母＝TTM 營業利益〕）").font = BOLD
vrow("ttm_am", "Americas 目標日 TTM 營業利益（US$M）"); vrow("ttm_em", "EMEA 目標日 TTM 營業利益（US$M）"); vrow("ttm_ap", "Asia Pacific 目標日 TTM 營業利益（US$M）"); vrow("ttm_tot", "三分部 TTM 營業利益合計（US$M）")
vrow("mult", "套用倍數（同業 EV／EBIT 中位數×調整係數；三分部同倍數：無分區同業）"); vrow("ev_c", "Comparable 企業價值（US$M）"); vrow("comp12", "Comparable 情境目標價（US$；淨現金取資產負債表日值）")
V.cell(VR["gaph"], 1, "DCF 與 Comparable 落差分解（情境 B；每股 US$；公式組成）").font = BOLD
vrow("gap", "落差：DCF 情境 B 目標價 − Comparable 情境 B 目標價", f"=D{VR['dcf12']}-D{VR['comp12']}")
vrow("gap_t", "其中時點：DCF 自評價日滾到 12 個月後（每股價值×ke−股利）", f"=D{VR['dcf12']}-D{VR['vps0']}")
vrow("gap_m", "其中倍數：評價日 DCF 每股價值 − Comparable 目標價（DCF 隱含的 EV／EBIT 對同業倍數）", f"=D{VR['vps0']}-D{VR['comp12']}")
vrow("gap_sbc", "其中股票薪酬：兩法皆以已扣股票薪酬的營業利益為基礎", 0)
vrow("gap_lev", "其中槓桿：兩法皆為企業價值（無槓桿）加同一淨現金", 0)
vrow("own_ev", "Accenture 現行企業價值（現價 × 評價股數 ＋ 債務 ＋ 非控制權益 − 現金；US$M）", f"=B{VR['emkt']}+B{VR['dmkt']}+B{VR['nciv']}-B{VR['cashx']}")
V.cell(VR["ownh"], 1, "本公司現行倍數 ÷ 同業中位數與隱含重估（原則第 47 條）").font = BOLD
vrow("own_mult", "Accenture 現行 EV／TTM 營業利益（FY2026 GAAP 營業利益 11,405.96）", f"=B{VR['own_ev']}/PL!E{PR['B']['ebit']}")
vrow("own_ratio", "現行倍數 ÷ 同業中位數", f"=B{VR['own_mult']}/B{VR['pe_med']}")
vrow("rerate", "若回到同業中位數（調整前）的隱含股價變動", f"=(B{VR['pe_med']}*PL!E{PR['B']['ebit']}+B{VR['netcash']})/B{VR['sh']}/B{VR['px']}-1")
V.cell(VR["g3h"], 1, "合理性關卡（工作單 G3 第 3 條；發布前必查）").font = BOLD
vrow("g3_dcf", "DCF 目標價 ÷ 現價（小於 0.25 或大於 4 觸發）", f"=B{VR['dcf_tp']}/B{VR['px']}")
vrow("g3_ev", "三情境企業價值最小值（US$M；負值觸發）", f"=MIN(C{VR['ev']}:E{VR['ev']})")
vrow("g3_rate", "估值評等為強力買進或強力放空（1＝觸發）", f'=IF(OR(B{VR["valr"]}="強力買進",B{VR["valr"]}="強力放空"),1,0)')
vrow("g3_rev", "FY2027 營收（情境 B）對指引中點差距（超過 10% 觸發）", f"=ABS(Recon!B{RCR_REV27}/AVERAGE(Recon!C{RCR_REV27}:D{RCR_REV27})-1)")
vrow("g3_om", "FY2027 營業利益率對指引中點差距（相對值，超過 10% 觸發）", f"=ABS(Recon!B{RCR_OM27}/AVERAGE(Recon!C{RCR_OM27}:D{RCR_OM27})-1)")
vrow("g3_flag", "合理性關卡是否觸發（1＝觸發）", f"=IF(OR(B{VR['g3_dcf']}<0.25,B{VR['g3_dcf']}>4,B{VR['g3_ev']}<0,B{VR['g3_rate']}=1,B{VR['g3_rev']}>0.1,B{VR['g3_om']}>0.1),1,0)")
V.cell(VR["scenh"], 1, "情境權重與加權").font = BOLD
vrow("wa", "情境權重 A／B／C", None, f"={wr('A')}", f"={wr('B')}", f"={wr('C')}")
vrow("wb", "情境加權目標價（各情境 DCF 與 Comparable 加權；US$）"); vrow("wc", "各情境上檔空間")
vrow("chk_w", "檢查：情境加權目標價與加權目標價之差（應為 0）", f"=ABS({wr('A')}*C{VR['wb']}+{wr('B')}*D{VR['wb']}+{wr('C')}*E{VR['wb']}-B{VR['tp']})")
V.cell(VR["peerh"], 1, "同業倍數（StockAnalysis；EV／EBIT，分母＝TTM 營業利益；日期不一致；家數 3，少於 8 不做迴歸）").font = BOLD
for j, h in enumerate(["同業", "EV／EBIT（TTM）", "資料日期", "會計準則與租賃處理", "逾期旗標（與現價日相差超過 2 個交易日）"], 1): V.cell(VR["peerh"] + 0, j + 0, V.cell(VR["peerh"], j).value) if False else None
PEERS = [("Cognizant（CTSH）", 9.17, datetime.date(2026, 10, 27), "US GAAP；營業租賃費用在營業利益內；EV 含租賃負債與否未核對"),
         ("Infosys（INFY）", 11.16, datetime.date(2026, 10, 21), "IFRS（ADR）；租賃折舊與利息分列，EBIT 不含租賃利息；EV 含租賃負債與否未核對"),
         ("CGI（GIB）", 10.28, datetime.date(2026, 10, 26), "IFRS（加拿大）；租賃折舊與利息分列；EV 含租賃負債與否未核對")]
for i, (nm, mv, dt, note) in enumerate(PEERS):
    rw = VR[f"pe{i+1}"]
    V.cell(rw, 1, nm); V.cell(rw, 2, mv).font = BLUE; V.cell(rw, 3, dt).number_format = "yyyy-mm-dd"; V.cell(rw, 4, note)
    V.cell(rw, 5, f'=IF(NETWORKDAYS(MIN(C{rw},{I("px_date")}),MAX(C{rw},{I("px_date")}))-1>2,"逾期","")')
vrow("pe_med", "同業中位數 EV／EBIT（TTM）", f"=MEDIAN(B{VR['pe1']}:B{VR['pe3']})")
vrow("pe_n", "家數", f"=COUNT(B{VR['pe1']}:B{VR['pe3']})")
V.cell(VR["pe_own"], 1, "Accenture 自身 EV／EBIT（StockAnalysis 2026-10-31；僅對照）"); V.cell(VR["pe_own"], 2, 12.57).font = BLUE
V.cell(VR["stale_rf"], 1, "rf 基準日與現價日相差（交易日；超過 2 個交易日標逾期）")
V[f"B{VR['stale_rf']}"] = f"=NETWORKDAYS(MIN({I('rf_date')},{I('px_date')}),MAX({I('rf_date')},{I('px_date')}))-1"
V[f"C{VR['stale_rf']}"] = f'=IF(B{VR["stale_rf"]}>2,"逾期","未逾期")'
V.cell(VR["pe_note"], 1, "同業倍數基準日全部早於現價日 2 個交易日以上者標「逾期」，使用前須更新；本版未更新（同業日期 2026-10-21 至 10-27，現價日 2026-11-13）。EPAM 頁無日期且資料為 2025 年，已排除。")
for s in "ABC":
    for y in YEARS[3:]:
        c = YC[y]
        Dp[f"{c}{DR[s]['df']}"] = f"=1/(1+Val!$B${VR['wacc']})^{c}{DR[s]['t']}"

# ---------------- RevDCF 頁
RD = SH["RevDCF"]
RV = {}
RD["A1"] = "反向 DCF：現價隱含什麼。以情境 B 路徑為底，兩個封閉解（DCF 對營業利益率與營收規模都是線性）；主導變數的整本重算反解在 Recon 第 4 節。"
RD["A2"] = "讀法：隱含值是「DCF 每股價值（評價日）＝現價」的點，不是評等門檻。"
rv_rows = [("lab", "項目"), ("v0", "情境 B 每股價值（評價日；US$）"), ("px", "現價（US$）"), ("gap", "現價與 DCF 每股價值之差（股權價值；US$M）"),
           ("k_m", "營業利益率全期平移 1 個百分點對股權價值的影響（US$M）"), ("d_m", "現價隱含的營業利益率平移（百分點）"), ("m_fy32", "情境 B FY2032E 營業利益率"), ("m_imp", "現價隱含的 FY2032 營業利益率"),
           ("ev0", "情境 B 企業價值（US$M）"), ("ev_imp", "現價隱含的企業價值（US$M）"), ("scale", "現價隱含的營收規模倍數（相對情境 B）"),
           ("mult_imp", "現價隱含的 EV／TTM 營業利益倍數（情境 B；目標日 TTM）"), ("mult_peer", "同業中位數 EV／EBIT"), ("ps", "分部加總：三分部共用同一同業倍數，隱含倍數即上列")]
for i, (k, lab) in enumerate(rv_rows): RV[k] = 4 + i; RD.cell(4 + i, 1, lab)
b_ = PR["B"]
RD[f"B{RV['v0']}"] = f"=Val!D{VR['vps0']}"; RD[f"B{RV['px']}"] = f"=Val!B{VR['px']}"
RD[f"B{RV['gap']}"] = f"=(B{RV['px']}-B{RV['v0']})*Val!B{VR['sh']}"
terms = "+".join(f"PL!{YC[y]}{b_['rev']}*DCF!{YC[y]}{DR['B']['frac']}*DCF!{YC[y]}{DR['B']['df']}" for y in YEARS[3:])
RD[f"B{RV['k_m']}"] = f"=0.01*(1-{I('tax')})*(({terms})+PL!K{b_['rev']}*(1+Val!B{VR['g']})*(1-Val!B{VR['g']}/Val!B{VR['ronic']})/(Val!B{VR['wacc']}-Val!B{VR['g']})/(1+Val!B{VR['wacc']})^DCF!K{DR['B']['pv']+1})"
RD[f"B{RV['d_m']}"] = f"=B{RV['gap']}/B{RV['k_m']}"
RD[f"B{RV['m_fy32']}"] = f"=PL!K{STR['營業利益率']}"
RD[f"B{RV['m_imp']}"] = f"=B{RV['m_fy32']}+B{RV['d_m']}/100"
RD[f"B{RV['ev0']}"] = f"=Val!D{VR['ev']}"
RD[f"B{RV['ev_imp']}"] = f"=B{RV['px']}*Val!B{VR['sh']}-Val!B{VR['netcash']}"
RD[f"B{RV['scale']}"] = f"=B{RV['ev_imp']}/B{RV['ev0']}"
RD[f"B{RV['mult_imp']}"] = f"=B{RV['ev_imp']}/Val!D{VR['ttm_tot']}"
RD[f"B{RV['mult_peer']}"] = f"=Val!B{VR['pe_med']}"
RD.cell(RV["ps"], 2, "—")

# ---------------- Recon 頁
Rc = SH["Recon"]
Rc["A1"] = "對帳：本模型 vs 公司指引與市場共識（共識與指引為 Interested-party；找不到者寫「未找到」）。口徑不同者不判定差異；說明欄的數字由公式組字。"
RCR = {}
fy27 = lambda k: f"PL!F{b_[k]}"
rec_rows = [
 ("h1", "1. 合併層：FY2027（情境 B）", None),
 ("g27", "營收年增率（本模型為美元；公司指引為當地幣）", (f"=PL!F{STR['營收年增率']}", 0.03, 0.06, "未找到", "")),
 ("rev27", "營收（US$M）", (f"=PL!F{STR['營收']}", "=PL!E%d*(1+C%d)" % (STR["營收"], 0), "=PL!E%d*(1+D%d)" % (STR["營收"], 0), 78030, "指引由 FY2026 營收乘當地幣增幅 3%–6% 推算（匯率未建模）；共識 78.03B 取自 tickflow 單一來源（Interested-party，品質低）")),
 ("om27", "GAAP 營業利益率", (f"=PL!F{STR['營業利益率']}", 0.159, 0.161, "未找到", "")),
 ("eps27", "GAAP 稀釋 EPS（US$）", (f"=PL!F{STR['EPS（GAAP 稀釋，美元）']}", 14.39, 14.81, 14.66, "共識 14.66（28 位，TIKR 2026-11-04）；tickflow 2026-11-16 為 14.95，兩者不同，範圍並列不擇一")),
 ("capex27", "資本支出（US$M）", (f"=PL!F{b_['capex']}", 900, 900, "未找到", "公司指引約 0.9B（Q4 FY26 法說）")),
 ("h2", "2. FY2028（情境 B）", None),
 ("eps28", "GAAP 稀釋 EPS（US$）", (f"=PL!G{STR['EPS（GAAP 稀釋，美元）']}", None, None, 15.89, "共識 15.89（26 位，TIKR）；tickflow 16.25（9 家）")),
 ("rev28", "營收（US$M）", (f"=PL!G{STR['營收']}", None, None, 82930, "共識 82.93B（tickflow，17 家；單一來源，Interested-party）")),
 ("h3", "3. 賣方目標價（個別，Interested-party）", None),
 ("tp1", "個別賣方目標價（US$）", ("=Val!B%d" % VR["tp"], None, None, 252, "Wells Fargo 252（2026-10-26）、Baird 252（2026-10-21，降評 Hold）、UBS 300（2026-10-26）；平均目標價與家數未找到")),
]
r = 3
for k, lab, spec in rec_rows:
    RCR[k] = r; Rc.cell(r, 1, lab)
    if spec is None:
        Rc.cell(r, 1).font = BOLD
        for j, h in enumerate(["本模型", "指引低", "指引高", "共識", "說明"], 2): Rc.cell(r, j, h).font = BOLD
    else:
        for j, v in enumerate(spec, 2):
            if v is not None and v != "": Rc.cell(r, j, v)
    r += 1
Rc[f"C{RCR['rev27']}"] = f"=PL!E{STR['營收']}*(1+C{RCR['g27']})"; Rc[f"D{RCR['rev27']}"] = f"=PL!E{STR['營收']}*(1+D{RCR['g27']})"
Rc[f"F{RCR['g27']}"] = (f'="本模型 FY2027 營收 "&TEXT(B{RCR["rev27"]}/1000,"0.0")&" 十億美元（年增 "&TEXT(B{RCR["g27"]},"0.0%")&"）；指引當地幣 +3% 至 +6%，其中併購貢獻約 2.5 點（CFO 口述）。本模型以人力成長 "&TEXT({I("hg")},"0.0%")&"＋價格 "&TEXT({I("rg")},"0.0%")&"為 AI 前成長，減 AI 淨效果 "&TEXT(Seg!F{SR["B"]["net"]}/1000,"0.0")&" 十億美元。"')
Rc[f"F{RCR['om27']}"] = (f'="本模型 "&TEXT(B{RCR["om27"]},"0.0%")&"，指引 15.9% 至 16.1%；FY2026 調整後 15.8%（排除 business optimization 307.5M）。落差＝"&TEXT((B{RCR["om27"]}-AVERAGE(C{RCR["om27"]}:D{RCR["om27"]}))*100,"+0.00;-0.00")&" 點，來自營業成本率與費用率持平於 FY2026 實績（規則：無證據不假設改善），公司指引隱含 AI 與定價帶來的改善。"')
Rc[f"F{RCR['eps27']}"] = (f'="本模型 EPS "&TEXT(B{RCR["eps27"]},"0.00")&"，指引 14.39–14.81；落差由營業利益率、稅率 "&TEXT({I("tax")},"0.0%")&"（FY2026 實績，公司未給 FY2027 稅率）與稀釋股數 "&TEXT({I("sh27")},"0")&"M（Assumed）共同造成。共識 14.66 取自 TIKR。"')
assert RCR["rev27"] == RCR_REV27 and RCR["om27"] == RCR_OM27
r += 1
Rc.cell(r, 1, "4. 反解（整本重算）：支持現價所需的主導變數值（原則第 4 條）").font = BOLD
RC_REV = r + 1
revs = json.load(open(REVS, encoding="utf-8")) if REVS else None
Rc.cell(RC_REV, 1, "主導變數（Sens 頁 AI 命題參數中對加權目標價影響最大者）")
Rc.cell(RC_REV + 1, 1, "基準值"); Rc.cell(RC_REV + 2, 1, "支持現價所需值（整本重算二分／割線法；加權目標價＝現價）"); Rc.cell(RC_REV + 3, 1, "說明")
if revs:
    Rc.cell(RC_REV, 2, revs["label"]); Rc.cell(RC_REV + 1, 2, revs["base"]); Rc.cell(RC_REV + 2, 2, revs["solved"]); Rc.cell(RC_REV + 3, 2, revs["note"])
else:
    Rc.cell(RC_REV, 2, "未解（尚未跑反解）")
r = RC_REV + 5
Rc.cell(r, 1, "5. 口徑與未找到").font = BOLD
Rc.cell(r + 1, 1, "同業 EV／EBIT 取自 StockAnalysis，日期 2026-10-21 至 10-27，與現價日不一致、已標逾期；家數 3；準則與租賃處理未核對；未取得 NTM 口徑倍數（原則第 10 節第 7 條），以 TTM 倍數乘目標日 TTM 營業利益。")
Rc.cell(r + 2, 1, "FY2027 Q1（2026-09 至 11）實績預計 2026-12 下旬公布；本版以 FY2026（截至 2026-08-31）為最新年度。")
Rc.cell(r + 3, 1, "未找到：Blume β 第二來源、公司債有效利率、融資租賃餘額、FY2024 GenAI 數字官方 8-K 原文、利用率定義、計費工時與時薪（公司不揭露）、平均賣方目標價與家數；詳見 research_accenture.md 第 9 節。")
RC_END = r + 3

# ---------------- Sens 頁
Se = SH["Sens"]
Se["A1"] = "敏感度：In 頁每個有「敏感度低／高」區間的輸入，一次動一個、整本 LibreOffice 重算；表中為加權目標價相對基準。情境專屬列與情境權重不列（三情境本身就是情境分析）。"
Se["A2"] = "token 價格高案（cps 高）只動成本一邊，標「上界」；另列價格與自動化量聯動、模型路由（改用便宜模型）的組合案（原則第 44 條）。群組 IF_HoldEcon 不適用。"
fp = hashlib.sha256(json.dumps([(k, str(wb["In"].cell(IR[k], 5).value)) for k in sorted(IR)], ensure_ascii=False).encode()).hexdigest()[:12]
BUILD_ID = f"accenture-v{VER}-{fp}"
sens = json.load(open(SENS, encoding="utf-8")) if SENS else None
Se.cell(3, 1, "build_id"); Se.cell(3, 4, BUILD_ID)
Se.cell(4, 1, "模型狀態"); Se.cell(4, 4, "CURRENT" if (sens and sens.get("fingerprint") == fp) else "STALE")
Se.cell(5, 1, "基準加權目標價（本檔）"); Se.cell(5, 4, f"=Val!B{VR['tp']}")
Se.cell(6, 1, "敏感度重算時的基準加權目標價"); Se.cell(6, 4, sens["base_tp"] if sens else None)
Se.cell(7, 1, "敏感度重算時的輸入指紋"); Se.cell(7, 4, sens["fingerprint"] if sens else None)
hdr_r = 9
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（美元）", "投資建議"], 1): Se.cell(hdr_r, j, h).font = BOLD
SE_ROWS = []
if sens:
    rr = hdr_r + 1
    for case in sens["cases"]:
        Se.cell(rr, 1, case["label"]); Se.cell(rr, 2, case["setting"]); Se.cell(rr, 3, case["side"])
        Se.cell(rr, 4, case["tp"]); Se.cell(rr, 5, f"=D{rr}/$D$5-1"); Se.cell(rr, 6, case["key"]); Se.cell(rr, 7, case["dcf"]); Se.cell(rr, 8, f"=G{rr}-Val!$B${VR['dcf_tp']}"); Se.cell(rr, 9, case["rec"])
        SE_ROWS.append(rr); rr += 1
    Se.cell(rr + 1, 1, "註：tp_w（情境權重）與 scn_*（情境驅動）不列：三情境並列已涵蓋；he_*（Tokenomics 持有成本）不適用：Accenture 不持有算力。")
SE_END = hdr_r + len(SE_ROWS)

# ---------------- Sources 頁
So = SH["Sources"]
So["A1"] = "來源與版本"; So.cell(2, 1, "項目"); So.cell(2, 2, "內容")
src_rows = [("Tokenomics 快照", f"{TK['source']['version']}（{TK['source']['file']}；commit {TK['source']['commit'][:7]}）；名稱：IF_TaskLen、IF_TaskTokFresh／Cached／Dec、IF_PriceFresh／Cached／Out_Terra／Sol／Astra、IF_TaskSucc_Terra／Sol／Astra（data/tokenomics_snapshot_v22.0.json）"),
            ("研究包", RES), ("公司申報", "10-K FY2026（SEC 0001467373-26-000049，2026-10-16）；10-K FY2025；10-Q FY2026 Q1–Q3；8-K 2026-10-01 Q4 FY26 新聞稿"),
            ("同業倍數", "StockAnalysis statistics 頁（CTSH 2026-10-27、INFY 2026-10-21、CGI（GIB）2026-10-26、ACN 2026-10-31）"), ("市場資料", "tradingeconomics 10 年美債 2026-11-16；StockAnalysis ACN 股價與 Beta"),
            ("本版先驗", "未做（Andy 2026-11-16 23:27）")]
for i, (a, b2) in enumerate(src_rows, 3): So.cell(i, 1, a); So.cell(i, 2, b2)

# ---------------- README 頁
Rm = SH["README"]
VB = lambda k, col="B": f"Val!{col}{VR[k]}"
Rm["A1"] = (f'="Accenture 個股模型 v{VER}｜資料截止 2026-11-16（最新年度 FY2026，截至 "&TEXT({I("bs_date")},"yyyy-mm-dd")&"）｜成品日 {VDATE}｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&"｜build_id {BUILD_ID}｜本版未做先驗"')
Rm["A2"] = "命題、結論與重點觀察的數字由公式自本檔其他頁組字（不手寫）。"
sb = SR["B"]
Rm["A3"] = (f'="命題：AI 進入按人力計價的 IT 服務後，每單位營收＝AI 新專案增量 − 定價侵蝕 − 推論成本。本模型以交付人力 × 利用率 × 可用工時 × 時薪推 AI 前營收，再依自動化比例、價格保留率與新專案占比推 AI 後營收。情境 B 在 FY2032 的 AI 淨效果為 "&TEXT(Seg!K{sb["net"]}/1000,"0.0")&" 十億美元（新專案 "&TEXT(Seg!K{sb["incr"]}/1000,"0.0")&" − 定價侵蝕 "&TEXT(Seg!K{sb["eros"]}/1000,"0.0")&" − 推論成本 "&TEXT(Seg!K{sb["inf"]}/1000,"0.00")&"），AI 後所需交付人力 "&TEXT(Seg!K{sb["hc_ai"]},"0")&" 千人（AI 前 "&TEXT(Seg!K{sb["hc0"]},"0")&" 千人）。"')
Rm["A4"] = "命題的可證偽指標：每員工營收年增是否明顯高於歷史（約 2%）、員工數下降而營收持平、利用率與合約改按成果計價的比例；任何一項偏離基準區間即須重估。"
Rm["A5"] = (f'="投資建議（在自動化、價格保留與新專案三項落在情境 A 到 C 區間的基準假設下）："&{VB("rec")}&"；加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"（現價 "&TEXT({VB("px")},"0.00")&" 美元，"&TEXT({VB("px_date")},"yyyy-mm-dd")&"）"')
Rm["A6"] = (f'="一、DCF 目標價 "&TEXT({VB("dcf_tp")},"0.00")&"、Comparable 目標價 "&TEXT({VB("comp_tp")},"0.00")&" 美元；情境 B 落差 "&TEXT({VB("gap")},"+0.00;-0.00")&" 美元，其中時點 "&TEXT({VB("gap_t")},"+0.00;-0.00")&"、倍數 "&TEXT({VB("gap_m")},"+0.00;-0.00")&"、股票薪酬 0、槓桿 0。"')
Rm["A7"] = (f'="二、情境 B FY2027 營收 "&TEXT(PL!F{STR["營收"]}/1000,"0.0")&" 十億美元（年增 "&TEXT(PL!F{STR["營收年增率"]},"0.0%")&"），對照公司指引當地幣 +3% 至 +6%；GAAP 營業利益率 "&TEXT(PL!F{STR["營業利益率"]},"0.0%")&"，指引 15.9% 至 16.1%（Recon）。"')
Rm["A8"] = (f'="三、每任務推論成本（Tokenomics Sol 層級）"&TEXT({I("cps")},"0.00")&" 美元，FY2032 推論成本僅占營收 "&TEXT(Seg!K{sb["inf"]}/Seg!K{sb["rev"]},"0.00%")&"；AI 對 Accenture 的影響主要在定價侵蝕與新專案，不在推論成本。"')
Rm["A9"] = (f'="四、現價隱含 FY2032 營業利益率 "&TEXT(RevDCF!B{RV["m_imp"]},"0.0%")&"（情境 B 為 "&TEXT(RevDCF!B{RV["m_fy32"]},"0.0%")&"）；現行 EV／EBIT "&TEXT({VB("own_mult")},"0.0")&" 倍，為同業中位數 "&TEXT({VB("pe_med")},"0.0")&" 倍的 "&TEXT({VB("own_ratio")},"0.00")&" 倍；永續價值占 EV 約 "&TEXT(Val!D{VR["tvsh"]},"0%")&"。"')
Rm["A10"] = (f'="五、情境 A／B／C 加權目標價 "&TEXT(Val!C{VR["wb"]},"0.00")&"／"&TEXT(Val!D{VR["wb"]},"0.00")&"／"&TEXT(Val!E{VR["wb"]},"0.00")&" 美元；分化主要來自自動化比例、價格保留率與新專案占比（FY2032 AI 淨效果 A／B／C 為 "&TEXT(Seg!K{SR["A"]["net"]}/1000,"0.0")&"／"&TEXT(Seg!K{SR["B"]["net"]}/1000,"0.0")&"／"&TEXT(Seg!K{SR["C"]["net"]}/1000,"0.0")&" 十億美元）。"')
Rm["A12"] = "v1.3 展示層"
Rm["A13"] = (f'="Accenture："&{VB("rec")}&"，加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元（上檔 "&TEXT({VB("up")},"+0.0%;-0.0%")&"）；本版未做先驗。"')
Rm["A14"] = "本版要點"
Rm["A15"] = "官方報告分部三個（Americas、EMEA、Asia Pacific；FY2024 為 10-K FY2026 新口徑重編值）；分部營收與營業利益皆取公司法定附註，三分部合計與合併值差 0。"
Rm["A16"] = "命題模組為原型 L；公司不揭露計費工時、時薪、價量拆分，時薪由 FY2026 營收與員工數、利用率反推（Derived），自動化比例、價格保留率、新專案占比為 Assumed，情境分化依據見 In 頁。"
Rm["A17"] = "推論成本只讀 Tokenomics v22.0 快照任務 5（類比）；公司未揭露 token 量與代理任務量，不以聊天基數乘倍數；token 價格路徑待 AI供應鏈 登錄。"
Rm["A18"] = "未做事項：先驗（Andy 指示本版跳過）；三層基準協調（Coord、C0、兩讀數）不適用：非 Siliconomics 覆蓋；Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照不適用：Accenture 不自建算力；v0.2 改讀 unit_economics。"
Rm["A19"] = (f'="合理性關卡（工作單 G3 第 3 條）："&IF({VB("g3_flag")}=1,"觸發，須診斷：","未觸發；")&"DCF 目標價÷現價 "&TEXT({VB("g3_dcf")},"0.00")&"、EV 最小值 "&TEXT({VB("g3_ev")}/1000,"0.0")&" 十億美元、FY2027 營收對指引中點差 "&TEXT({VB("g3_rev")},"0.0%")&"、營業利益率相對差 "&TEXT({VB("g3_om")},"0.0%")&"、評等 "&{VB("valr")}&"。診斷結論：本模型評等由 AI 淨效果之成本節省比例 cs（Assumed，Sens 最敏感）與同業倍數（三家、已逾期）決定；DCF 與 Comparable 彼此接近、與現價落差來自市場對 Accenture 的倍數溢價（現行倍數÷同業中位數見 Val），屬經濟判斷而非建模錯誤；恆等式核對皆為 0。"')
Rm["A20"] = "市場資料日期"
Rm["A21"] = (f'="現價 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤；無風險利率 "&TEXT({I("rf_date")},"yyyy-mm-dd")&"（與現價日相差 "&TEXT({VB("stale_rf")},"0")&" 個交易日，"&{VB("stale_rf","C")}&"）；同業倍數 2026-10-21 至 2026-10-27，已標逾期（Val 頁）。"')
Rm["A22"] = "主要失效條件（可觀察）：FY2027 Q1（2026-12 公布）起每員工營收年增明顯高於 2%、員工數下降而營收持平、新訂單 book-to-bill 跌破 1.0、定價持續「lower pricing in many areas」。"
Rm["A23"] = "風險登錄"
Rm["A24"] = "風險"; Rm["B24"] = "是否建模"; Rm["C24"] = "理由"; Rm["D24"] = "觀察指標"
risks = [("AI 取代交付工時（定價侵蝕）", "是", "a32 與 r 三情境", "每員工營收、利用率、員工數"), ("新專案不足以填補", "是", "n32 三情境", "新訂單、GenAI 相關營收（公司已停止揭露）"), ("推論成本", "是", "Tokenomics 任務 5 成本，Sens 上界", "毛利率"),
         ("併購依賴", "部分", "acq27、acq_run 進 FCFF；人力成長含併入", "併購現金、商譽"), ("匯率與地緣", "否", "美元口徑，未建模匯率", "當地幣成長 vs 美元成長")]
for i, rk in enumerate(risks, 25):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm.cell(31, 1, "工作表：README、In、Q、Seg、PL、DCF、Val、RevDCF、Recon、Sens、Sources")
Rm["A33"] = "頁"; Rm["B33"] = "輸入"; Rm["C33"] = "推導量"; Rm["D33"] = "因果方向"
mp = [("Seg", "hc_y26、util、hrs、hg、rg、a27、a32_A/B/C、r_A/B/C、n27、n32_A/B/C、tlen、cps、px_dec、d_em、d_ap", "AI 前營收、定價侵蝕、新專案增量、推論成本、AI 後人力、三分部營收", "人力 × 利用率 × 工時 × 時薪 → R0；自動化與價格保留 → 侵蝕；任務 × 成本 → 推論"),
      ("PL", "cs、c0_dr、opx_dr、da_pct、capex_pct、wc_k、acq27、acq_run、tax、oth_net、nci_pct", "毛利、營業利益、分部營業利益、FCFF、淨利、EPS", "營收 − 成本 − 費用 → 營業利益 → NOPAT → FCFF"),
      ("Val", "rf、beta_raw、erp、kd_sp、g_tv、ronic_sp、w_dcf、cm_adj、股數與淨現金", "WACC、DCF、分部加總 Comparable、加權目標價、評等", "FCFF 折現＋同業倍數 × TTM 營業利益 → 加權"),
      ("Sens", "In 頁敏感度低／高", "加權目標價、DCF 目標價、投資建議", "一次改一個輸入、整本重算")]
for i, rk in enumerate(mp, 34):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm["A39"] = "版本沿革：v0.1（2026-11-16）首版；未做先驗；本版與其他版本無對照。"
for ws_ in wb.worksheets:
    ws_.column_dimensions["A"].width = 52
    for col in "BCDEFGHIJK": ws_.column_dimensions[col].width = 14
wb["In"].column_dimensions["B"].width = 60; wb["In"].column_dimensions["H"].width = 70
wb.defined_names["PL_hdr"] = DefinedName("PL_hdr", attr_text="PL!$C$5:$K$5")
wb.save(OUT)
rowmap = {"IR": IR, "SR": SR, "PR": PR, "STR": STR, "SPR": SPR, "VR": VR, "DR": DR, "RV": RV, "RCR": RCR, "CK": CK, "QR": QR, "SE_ROWS": SE_ROWS, "SE_END": SE_END, "SCN_H": SCN_H, "SCN_ROWS": SCN_ROWS,
          "BUILD_ID": BUILD_ID, "FP": fp, "HDR_R": hdr_r, "GATE": GATE, "RC_REV": RC_REV, "RC_END": RC_END, "CPS": CPS}
if ROWMAP: json.dump(rowmap, open(ROWMAP, "w"), ensure_ascii=False, indent=1)
print("built", OUT, BUILD_ID)
