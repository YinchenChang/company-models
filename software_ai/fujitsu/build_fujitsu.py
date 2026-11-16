#!/usr/bin/env python3
"""富士通（6702）個股模型 v0.1（原型 L：按人力計價，套用於 Service Solutions；命題＋官方分部營收 → 損益表 → FCFF DCF＋分部加總 Comparable）。
用法：python3 build_fujitsu.py <raw.xlsx> [--sens data/sens_v01_20261116.json] [--rowmap rowmap.json] [--rev data/rev_solve_v01.json] [--set key=value ...]
所有計算為活公式；輸入在 In 頁（藍字）。LibreOffice 重算後才有快取值。本版未做先驗。金額 ¥百萬；會計年度 4 月制（FY2026＝2026-04-01～2027-03-31）。
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


def task_cost(tier):  # 每成功任務推論成本（US$；Task 5：Coding agent 長程，人類 8 小時）
    g = lambda n: TKI[n]["values"]
    f, c, d = (g(k)[4] for k in ("IF_TaskTokFresh", "IF_TaskTokCached", "IF_TaskTokDec"))
    return (f * g(f"IF_PriceFresh_{tier}") + c * g(f"IF_PriceCached_{tier}") + d * g(f"IF_PriceOut_{tier}")) / 1e6 / g(f"IF_TaskSucc_{tier}")[4]


CPS_USD = {t: task_cost(t) for t in ("Terra", "Sol", "Astra")}
TLEN = TKI["IF_TaskLen"]["values"][4]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
SH = {n: wb.active if n == "README" else wb.create_sheet(n) for n in ["README", "In", "Q", "Seg", "PL", "DCF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
SH["README"].title = "README"

D0 = "Verified"; IP = "Interested-party"; AN = "Analogy"; AS = "Assumed"; DE = "Derived"
R2 = "決算短信 2026Q2（2026-10-29）https://resources.global.fujitsu/prir/2026/10/2026Q2-Financial-Results-en.pdf"
RF25 = "決算短信 FY2025（2026-04-28）https://global.fujitsu/-/media/Project/Fujitsu/Fujitsu-HQ/ir/documents/Financial_Results/FY2025/FY/2025FY-Financial-Results-en.pdf"
RES = "software_ai/fujitsu/data/research_fujitsu.md、facts_fujitsu.csv"
INP = []


def add(key, label, unit, lo, base, hi, tag, src, inc=""):
    if tag == IP and not inc: raise SystemExit(f"{key}: Interested-party 須填來源誘因")
    INP.append((key, label, unit, lo, base, hi, tag, src, inc))


add("px_date", "股價基準日（收盤）", "日期", None, datetime.date(2026, 11, 13), None, D0, "traders.co.jp 6702 歷史股價頁（單一來源）")
add("px", "現價（2026-11-13 收盤）", "¥/股", None, 3455, None, D0, "https://www.traders.co.jp/stocks/61_6702/historical（單一來源；StockAnalysis 2026-10-23 頁為 4,477，下跌約 23%，期間含 2026-10-29 決算，未找到第二來源）")
add("bs_date", "資產負債表日（2026 上半年末）", "日期", None, datetime.date(2026, 9, 30), None, D0, R2 + " 財政狀態表")
add("sh_issued", "已發行股數（2026-09-30）", "M股", None, 1739.778, None, D0, R2 + " 封面：1,739,778,265 股")
add("sh_tr", "庫藏股（2026-09-30）", "M股", None, 13.160, None, D0, R2 + " 封面：13,160,155 股")
add("sh_dil", "潛在稀釋股（股票報酬等）", "M股", 0, 0, 10, AS, "未找到稀釋證券明細（決算短信封面無；稀釋 EPS 與基本 EPS 差異極小）；基準 0，高端 10M 作敏感度")
add("cash", "現金及約當現金（2026-09-30）", "¥百萬", None, 413389, None, D0, R2 + " 財政狀態表")
add("debt", "債務與租賃負債（短期借款與一年內到期長期債務及租賃 40,487＋長期債務及租賃 86,050）", "¥百萬", None, 126537, None, D0, R2 + " 財政狀態表；IFRS 16 下租賃負債與借款合併列示、無法分開（口徑：含租賃，同業 IFRS 公司同樣含租賃）")
add("inv_eq", "權益法投資（2026-09-30）", "¥百萬", None, 87949, None, D0, R2 + " 財政狀態表")
add("inv_oth", "其他投資（2026-09-30）", "¥百萬", None, 164880, None, D0, R2 + " 財政狀態表")
add("inv_hc", "非營運投資計入企業價值外資產的比例", "比例", 0.5, 0.8, 1.0, AS, "帳面價值的折扣；持股市值與流動性未核對；退休金淨負債未找到，未扣除")
add("nci", "非控制權益（2026-09-30）", "¥百萬", None, 19840, None, D0, R2 + " 財政狀態表；自股權價值扣除")
add("dps", "每股股利（年）", "¥/股", None, 55, None, IP, "決算短信 2026Q2：FY2026 股利預測 55 日圓（FY2025 50 日圓）", "公司股利預測，非實績")
add("rf_date", "無風險利率基準日", "日期", None, datetime.date(2026, 11, 16), None, D0, "https://tradingeconomics.com/japan/government-bond-yield（2026-11-16 讀取）")
add("rf", "無風險利率（10 年日債）", "比例", 0.025, 0.0289, 0.035, D0, "tradingeconomics 日本 10 年殖利率 2.89%（2026-11-16；單一來源）")
add("beta_raw", "Beta（5 年，單一來源；Blume 調整後用）", "—", 0.6, 0.74, 1.0, D0, "StockAnalysis 6702 statistics 頁 Beta (5Y) 0.74（2026-10-23 版）；規格要兩來源，未找到第二來源")
add("erp", "股權風險溢酬", "比例", 0.04, 0.05, 0.06, AS, "DCF 標準規格預設 5%（4–6%）；未做日本市場專屬估計")
add("kd_sp", "稅前債務信用利差（相對 rf）", "比例", 0.003, 0.006, 0.015, AS, "未找到債券有效利率；以利差假設")
add("g_tv", "永續成長率（日圓名目）", "比例", 0.010, 0.015, 0.025, AS, "低於預設 3%：日圓無風險利率 2.89%、永續成長不宜超過 rf；區間 1–2.5%")
add("ronic_sp", "RONIC 減 WACC 的利差", "比例", -0.02, 0.0, 0.02, AS, "DCF 標準規格預設 RONIC＝WACC")
add("tax", "稅率（NOPAT 與損益）", "比例", 0.25, 0.287, 0.31, DE, "2026 上半年所得稅 33,869÷稅前利益 118,011＝28.7%（決算短信 2026Q2；Derived）；預測期持平")
add("w_dcf", "DCF 權重（其餘為 Comparable）", "比例", 0.3, 0.5, 0.7, AS, "原則第 4 節預設 50／50")
add("cm_adj", "Comparable 倍數調整係數（對同業中位數）", "倍", 0.8, 1.0, 1.2, AS, "同業僅 4 家、日期與準則不一致，不做迴歸；本公司歷史相對倍數未自算（G3 第 4 條：未做），1.00 為承重輸入")
add("hw_disc", "Hardware／Ubiquitous 倍數相對 Service 倍數", "倍", 0.4, 0.6, 0.8, AS, "找不到硬體與終端同業倍數；以折扣假設（硬體利益率 5.7% 對 Service 15.6%，LTM，Derived）")
# ---- 命題（原型 L，套用於 Service Solutions）
add("hc_y25", "期末連結員工數 FY2025（含全部分部）", "千人", None, 99.203, None, D0, RF25 + " 主要經營指標；FY2024 的 112,743 含已出售的 Device Solutions，口徑不可比")
add("svc_sh", "Service Solutions 占連結員工比例", "比例", 0.60, 0.75, 0.85, AS, "公司不揭露分部員工數；Service 占營收 66%（Derived）；比例為假設，只影響時薪反推值，不影響營收水準")
add("util", "利用率（Service 交付人員）", "比例", 0.70, 0.80, 0.88, AS, "公司不揭露利用率；取 unit_economics L_util 基準 0.80")
add("hrs", "每人年可用工時", "小時/年", 1800, 1880, 1960, AS, "unit_economics L_hrs 基準 1,880；公司未揭露")
add("hg", "AI 前 Service 淨人力成長", "比例/年", 0.000, 0.015, 0.030, AS, "連結員工 FY2024→FY2025 因出售 Device Solutions 不可比；Service H1 營收 +5.9%（Derived）、日本 +6.8%／海外 +3.2%；基準取 1.5%")
add("rg", "AI 前每工時價格與組合成長", "比例/年", 0.010, 0.025, 0.040, AS, "價值導向定價占比 20%（FY2025）→60%（FY2030）目標（IR Day 2026，Interested-party）；Uvance 占 Service 約 30%；基準 2.5%")
add("a26", "AI 自動化工時占比 FY2026（相對 FY2025 基期的增量；情境共同）", "比例", 0.0, 0.02, 0.05, AS, "公司稱 AI 驅動交付 H1 效益約 70 億日圓、H2 預期為 2–3 倍（Q&A 口述，Interested-party）；約占 Service 營業成本 0.1%，故起點取 2%（假設）")
add("n26", "AI 新專案營收占 AI 前 Service 營收 FY2026（情境共同）", "比例", 0.0, 0.015, 0.04, AS, "Uvance 營收 FY2025 709.3B（Interested-party）為公司定義的服務組合、含舊工作轉型，不得當作 AI 增量；起點假設")
add("tlen", "代理任務長度（人類完成時間）", "小時", None, TLEN, None, D0, "Tokenomics v22.0 IF_TaskLen 第 5 項（8 小時；長程 Coding agent）；快照 data/tokenomics_snapshot_v22.0.json")
add("fx", "匯率（¥/US$；僅用於把 Tokenomics 美元成本換成日圓）", "¥/US$", 140, 150, 165, AS, "未找到來源（不影響營收；只影響推論成本換算，且推論成本占營收極小）；假設")
add("cps", "每成功任務推論成本（US$；Sol 層級市場價；Terra 為低、Astra 為高）", "US$/件", CPS_USD["Terra"], CPS_USD["Sol"], CPS_USD["Astra"], AN,
    "Tokenomics v22.0：任務 5 每次嘗試 token（新鮮 200,400／快取 1,263,600／decode 30,000；IF_TaskTok*）× 層級有效牌價（IF_PriceFresh／Cached／Out）÷ 成功率（IF_TaskSucc）；公司未揭露 token 量與代理任務量；選哪一層屬類比。高端＝只動成本一邊的上界")
add("px_dec", "每件推論成本逐年變動（FY2027 起）", "比例/年", -0.25, 0.0, 0.0, AS, "價格路徑待 AI供應鏈 登錄；基準持平（保守）；低端每年降 25%")
add("cs", "自動化工時省下的 Service 成本比例（成本含銷管費）", "比例", 0.25, 0.5, 0.8, AS, "無揭露；Service 成本率 LTM 84.4%（含銷管費，Derived），只有部分為可自動化的交付人力；公司稱 H2 Service 毛利率目標 43% 部分仰賴 AI 驅動生產力（Interested-party）")
add("ghw26", "Hardware Solutions 營收 FY2026 年增", "比例", -0.08, -0.03, 0.03, AS, "H1 +5.7%（449,107 對 424,813，Derived）；公司 FY2026 指引 960B（−4.9%，Interested-party）；基準取兩者之間")
add("ghw", "Hardware Solutions 營收 FY2027 起年增", "比例/年", -0.03, 0.0, 0.02, AS, "FY2024→FY2025 −9.8%（Derived）；無證據可延伸，持平")
add("gub26", "Ubiquitous Solutions 營收 FY2026 年增", "比例", -0.35, -0.25, -0.10, AS, "H1 −20.6%（89,886 對 113,189，Derived）；公司指引 160B（−30%，Interested-party）")
add("gub", "Ubiquitous Solutions 營收 FY2027 起年增", "比例/年", -0.05, 0.0, 0.03, AS, "無證據可延伸，持平")
add("corp_dx", "全社費用（Corporate）相對 LTM 的調整", "¥百萬", 10000, 0, -13000, AS, "規則：持平於 LTM −81,869（Derived）；公司 FY2026 指引 −95,000（Interested-party）為高端（更負）")
add("oneoff", "非常態與調整項目（調整後營業利益 − IFRS 營業利益）", "¥百萬/年", 0, 12700, 30000, DE, "2026 上半年 113,961−107,602＝6,359 的兩倍；FY2025 為 42,260（390,589−348,329）；公司 FY2026 指引 425,000−415,000＝10,000")
add("da_pct", "折舊攤銷與減損占營收", "比例", 0.035, 0.0403, 0.045, DE, "FY2025 141,227÷3,502,971＝4.03%（現金流量表；含減損；持平）")
add("capex_pct", "資本支出占營收", "比例", 0.030, 0.035, 0.045, AS, "FY2024 4.5%、FY2025 3.1%、2026 上半年 4.8%（78,036÷1,626,358）（Derived）；取中間")
add("wc_k", "營運資金增加占營收增量比例", "比例", 0.0, 0.05, 0.10, AS, "FY2025 淨利含非繼續營業的出售利益，無法由淨利對營業現金流回推歷史比率（G3 第 2 條：歷史比率未找到）；假設")
add("oth_net", "業外淨額（金融損益＋持分法損益）", "¥百萬/年", 0, 20800, 40000, DE, "2026 上半年稅前 118,011−IFRS 營業利益 107,602＝10,409 的兩倍（Derived）")
add("nci_pct", "非控制權益占淨利比例", "比例", 0.010, 0.0186, 0.030, DE, "2026 上半年 (84,142−82,581)÷84,142（決算短信 2026Q2；Derived）")
add("gm", "毛利率（標準損益表用；持平於 LTM）", "比例", 0.32, 0.3533, 0.37, DE, "LTM（FY2025＋2026 上半年−2025 上半年）毛利 1,258,750÷營收 3,562,787（決算短信；Derived）；Q2 FY2026 較前年同期低約 2 點（Interested-party 口述）")
add("sh26", "FY2026 加權稀釋股數（僅 EPS 顯示）", "M股", 1715, 1725, 1735, AS, "2026 上半年平均 1,731.9M、期末流通 1,726.6M（Verified）；僅影響 EPS 顯示")
add("bb", "FY2027 起每年股數淨減少（僅 EPS 顯示）", "比例/年", 0.0, 0.02, 0.04, IP, "FY2026 庫藏股買回上限 150B 日圓（最多 1 億股）；已發行股數 2,071M→1,740M（2025-03→2026-03）", "公司買回授權，實際執行視股價")
SC = {
 "a32": ("AI 自動化工時占比 FY2032（占 AI 前 Service 工時；線性自 a26）", "比例", 0.15, 0.20, 0.30, AS, "三情境分化依據：自動化強度。unit_economics L_a32 區間 0.15–0.60（Assumed）；公司稱 AI 驅動交付為生產力最強驅動（Interested-party）", ""),
 "r": ("價格保留率（被自動化工時中仍向客戶收取的比例）", "比例", 0.75, 0.55, 0.30, AS, "分化依據：定價模式。IR Day 2026 目標價值導向定價占比 20%→60%（FY2030）→80%（FY2035），毛利率貢獻 3 點以上（Interested-party，公司策略宣示）；高占比＝保留高；unit_economics L_r 0.10–0.70", ""),
 "n32": ("AI 新專案營收占 AI 前 Service 營收 FY2032（線性自 n26）", "比例", 0.12, 0.08, 0.04, AS, "分化依據：新工作量。Fujitsu Uvance、Kozuchi、前線部署工程師（FDE）投入（Q&A 口述，Interested-party）；unit_economics L_n32 0.03–0.25", ""),
 "tp_w": ("情境權重", "比例", 0.25, 0.5, 0.25, AS, "原則第 4 節預設 25／50／25", ""),
}
for k, v in SETS.items():
    for i, row in enumerate(INP):
        if row[0] == k: INP[i] = row[:4] + (v,) + row[5:]

ws = SH["In"]
ws["A1"] = "富士通 個股模型 v0.1 輸入頁（藍字＝輸入；黑字＝公式）。金額 ¥百萬，會計年度 4 月制（FY2026＝2026-04-01～2027-03-31）。"
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
HIST = ["FY2024", "FY2025"]
FC = YEARS[2:]          # 預測年 FY2026–FY2032（欄 E–K）
Q = SH["Q"]
Q["A1"] = "Q 頁：公司法定申報歷史數字（¥百萬；FY2024–FY2025 為會計年度，4 月制；FY2025 上半年、FY2026 上半年為半年度；LTM＝FY2025＋FY2026 上半年−FY2025 上半年，至 2026-09-30）。每列單一標記；Interested-party 列有來源誘因。"
Q["A2"] = "分部數字：FY2024、FY2025 取決算短信 p.8（¥百萬）；半年度為 Q1、Q2 分部數字相加（Q2 為 H1 減 Q1，Derived，標記於各列來源）。FY2024 稅前利益以下、FY2025 所得稅、FY2025 上半年稅前未找到。"
for j, h in enumerate(["指標", "FY2024", "FY2025", "FY2025 上半年", "FY2026 上半年", "LTM（至 2026-09）", "單位", "來源", "標記", "來源誘因"], 1): Q.cell(3, j, h).font = BOLD
QR = {}
SRC_SEG = RF25 + " p.8；半年度＝決算短信 2026Q1（2026-07-30）p.8 與 2026Q2 p.8–9 相加（Q2＝H1 減 Q1，Derived）"
qdata = [
 ("seg_sv", "Service Solutions 營收", 2245996, 2346932, 1066506, 1129445, SRC_SEG, D0, ""),
 ("seg_hw", "Hardware Solutions 營收", 1119938, 1009857, 424813, 449107, SRC_SEG, D0, ""),
 ("seg_ub", "Ubiquitous Solutions 營收", 251746, 229805, 113189, 89886, SRC_SEG, D0, ""),
 ("seg_el", "Elimination/Corporate 營收", -67564, -83623, -37966, -42080, SRC_SEG, D0, ""),
 ("rev", "合併營收", 3550116, 3502971, 1566542, 1626358, RF25 + " p.4；2026Q2 決算短信 p.4（H1）", D0, ""),
 ("oi_sv", "Service Solutions 分部調整後營業利益", 289967, 361464, 119622, 133769, SRC_SEG, IP, "公司自訂的分部利益指標（調整後營業利益，排除非常態項目）"),
 ("oi_hw", "Hardware Solutions 分部調整後營業利益", 61300, 67012, 12525, 4503, SRC_SEG, IP, "同上"),
 ("oi_ub", "Ubiquitous Solutions 分部調整後營業利益", 31372, 38825, 21744, 13410, SRC_SEG, IP, "同上"),
 ("oi_co", "Elimination/Corporate 調整後營業利益", -75374, -76712, -32564, -37721, SRC_SEG, IP, "同上"),
 ("adj", "調整後營業利益（合併）", 307265, 390589, 121327, 113961, RF25 + " p.8；2026Q2 p.8", IP, "公司自訂指標"),
 ("ebit", "營業利益（IFRS）", 265089, 348329, 105345, 107602, RF25 + " p.4；2026Q1／Q2 決算短信（H1 25＝Q1 33,486＋Q2 71,859）", D0, ""),
 ("cogs", "銷貨成本", 2382138, 2256104, 1030261, 1078194, RF25 + " p.4；2026Q2 p.4；FY2025 上半年＝營收 1,566,542−毛利 536,281（Derived）", D0, ""),
 ("gp", "毛利", 1167978, 1246867, 536281, 548164, RF25 + " p.4；2026Q1／Q2", D0, ""),
 ("sga", "銷管費", 887146, 886714, None, 457652, RF25 + " p.4；2026Q2 p.4", D0, ""),
 ("ebt", "稅前利益（繼續營業）", 273445, 409034, None, 118011, RF25 + " p.4；2026Q2 p.4", D0, ""),
 ("ni", "歸屬母公司當期利益", 219807, 449408, None, 82581, RF25 + " p.4（FY2025 含 Device Solutions 出售等非繼續營業利益）；2026Q2", D0, ""),
 ("eps", "稀釋 EPS（¥）", 120.66, 254.05, None, 47.55, RF25 + " p.4；2026Q2 p.4", D0, ""),
 ("da", "折舊攤銷與減損", 160894, 141227, None, 69286, RF25 + " p.6；2026Q2 現金流量表", D0, ""),
 ("capex", "資本支出（PP&E 與無形資產取得）", 160898, 108977, None, 78036, RF25 + " p.6；2026Q2 現金流量表", D0, ""),
 ("hc", "期末連結員工數（千人）", 112.743, 99.203, None, None, RF25 + " p.2 主要經營指標（FY2024 含 Device Solutions）", D0, ""),
]
rr = 4
for key, lab, a, b, c_, d, src, tag, inc in qdata:
    QR[key] = rr
    Q.cell(rr, 1, lab)
    for j, v in zip((2, 3, 4, 5), (a, b, c_, d)):
        Q.cell(rr, j, v if v is not None else "未找到").font = BLUE
    if key not in ("eps", "hc") and a is not None and c_ is not None and d is not None: Q.cell(rr, 6, f"=C{rr}+E{rr}-D{rr}")
    elif key in ("hc", "eps"): Q.cell(rr, 6, "—")
    else: Q.cell(rr, 6, "未找到")
    Q.cell(rr, 7, "¥" if key == "eps" else ("千人" if key == "hc" else "¥百萬")); Q.cell(rr, 8, src); Q.cell(rr, 9, tag); Q.cell(rr, 10, inc)
    rr += 1
CHKQ = []
for lab, f in (("核對：四列分部營收合計－合併營收（應為 0）", "{c}%d+{c}%d+{c}%d+{c}%d-{c}%d" % (QR["seg_sv"], QR["seg_hw"], QR["seg_ub"], QR["seg_el"], QR["rev"])),
               ("核對：四列分部調整後營業利益合計－調整後營業利益（應為 0）", "{c}%d+{c}%d+{c}%d+{c}%d-{c}%d" % (QR["oi_sv"], QR["oi_hw"], QR["oi_ub"], QR["oi_co"], QR["adj"])),
               ("核對：營收－銷貨成本－毛利（應為 0）", "{c}%d-{c}%d-{c}%d" % (QR["rev"], QR["cogs"], QR["gp"]))):
    Q.cell(rr, 1, lab)
    for j, col in zip((2, 3, 4, 5), "BCDE"): Q.cell(rr, j, "=" + f.format(c=col))
    CHKQ.append(rr); rr += 1
Q.cell(rr, 1, "FY2024 與 FY2025 營收為 IFRS 持續營業（Device Solutions 已出售，FY2024 為重編值）；FY2025 稅前利益以下 FY2024 不可與 FY2025 的出售利益直接比較。")
LTM = "F"


def QH(key, y): return f"Q!${'BC'[HIST.index(y)]}${QR[key]}"
def QL(key): return f"Q!$F${QR[key]}"


# ---------------- Seg 頁（命題模組；原型 L，Service Solutions）
S = SH["Seg"]
S["A1"] = ("命題模組（原型 L：按人力計價，套用於 Service Solutions）：Service 營收＝AI 前營收 R0（Service 人力 × 利用率 × 可用工時 × 時薪）−（被自動化工時 × 時薪 ×（1−價格保留率））＋ R0 × 新專案占比；"
           "推論成本＝被自動化工時 ÷ 任務長度 × 每成功任務成本（Tokenomics v22.0 快照，換算日圓）。Hardware 與 Ubiquitous 為不按人力計價的獨立分部，以年增率預測。")
S["A2"] = ("變數沿用 software_ai/unit_economics：L_util、L_hrs、L_rate、a、r、n、tier；該模型 v0.1 完成前在本模型 In 頁自建（v0.2 改讀 unit_economics）。FY2025 基期營收已含既有 AI 影響，a 與 n 為相對基期的增量；"
           "公司不揭露分部員工、利用率、計費工時與時薪，時薪由 FY2025 Service 營收反推（Derived）。")
SEGLAB = ["t", "hc0", "rate", "r0", "a", "n", "hauto", "tasks", "cps", "eros", "incr", "inf", "net", "hc_ai", "sv", "hw", "ub", "el", "rev", "chk1"]
LABS = {"t": "年序（FY2025＝0）", "hc0": "AI 前 Service 有效人力（千人）", "rate": "混合計費時薪（FY2025 Service 營收反推；¥/小時）", "r0": "AI 前 Service 營收 R0（人力 × 利用率 × 可用工時 × 時薪）",
        "a": "自動化工時占比 a（相對 FY2025 基期）", "n": "AI 新專案營收占 R0 比例 n", "hauto": "被自動化工時（百萬小時）", "tasks": "代理任務數（百萬件）", "cps": "每成功任務推論成本（¥；Tokenomics 美元成本×匯率）",
        "eros": "AI 淨效果之一：定價侵蝕（R0 × a ×（1−價格保留率））", "incr": "AI 淨效果之二：新專案營收增量（R0 × n）", "inf": "AI 淨效果之三：推論成本（Tokenomics 市場價）",
        "net": "AI 淨效果（增量 − 侵蝕 − 推論成本）", "hc_ai": "AI 後所需 Service 交付人力（千人）", "sv": "Service Solutions", "hw": "Hardware Solutions", "ub": "Ubiquitous Solutions", "el": "Elimination/Corporate",
        "rev": "合併營收（四列合計）", "chk1": "檢查：四列合計 − 合併營收（歷史年度對 Q 頁）"}
SR = {}
for si, s in enumerate("ABC"):
    r0 = 4 + si * (len(SEGLAB) + 4)
    S.cell(r0, 1, f"情境 {s}（{ {'A':'樂觀：自動化低、價格保留高、新專案多','B':'基準','C':'保守：自動化高、價格保留低、新專案少'}[s] }）").font = BOLD
    for y, c in YC.items():
        S[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); S[f"{c}{r0+1}"].font = BOLD
    S.cell(r0 + 1, 1, "¥百萬 除另註")
    rows = {k: r0 + 2 + i for i, k in enumerate(SEGLAB)}
    SR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items(): S.cell(rw, 1, LABS[k])
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        P = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            S[R("sv")] = f"={QH('seg_sv', y)}"; S[R("hw")] = f"={QH('seg_hw', y)}"; S[R("ub")] = f"={QH('seg_ub', y)}"; S[R("el")] = f"={QH('seg_el', y)}"
            S[R("rev")] = f"={QH('rev', y)}"
            S[R("chk1")] = f"={R('sv')}+{R('hw')}+{R('ub')}+{R('el')}-{R('rev')}"
            if y == "FY2025":
                S[R("t")] = 0
                S[R("hc0")] = f"={I('hc_y25')}*{I('svc_sh')}"
                S[R("rate")] = f"={R('sv')}*1000000/({R('hc0')}*1000*{I('util')}*{I('hrs')})"
                S[R("r0")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('rate')}/1000000"
            continue
        S[R("t")] = f"={P('t')}+1"
        S[R("hc0")] = f"={P('hc0')}*(1+{I('hg')})"
        S[R("rate")] = f"={P('rate')}*(1+{I('rg')})"
        S[R("r0")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('rate')}/1000000"
        S[R("a")] = f"={I('a26')}+({I('a32_'+s)}-{I('a26')})*({R('t')}-1)/6"
        S[R("n")] = f"={I('n26')}+({I('n32_'+s)}-{I('n26')})*({R('t')}-1)/6"
        S[R("hauto")] = f"={R('hc0')}*1000*{I('util')}*{I('hrs')}*{R('a')}/1000000"
        S[R("tasks")] = f"={R('hauto')}/{I('tlen')}"
        S[R("cps")] = f"={I('cps')}*{I('fx')}*(1+{I('px_dec')})^MAX({R('t')}-1,0)"
        S[R("eros")] = f"={R('r0')}*{R('a')}*(1-{I('r_'+s)})"
        S[R("incr")] = f"={R('r0')}*{R('n')}"
        S[R("inf")] = f"={R('tasks')}*{R('cps')}"
        S[R("net")] = f"={R('incr')}-{R('eros')}-{R('inf')}"
        S[R("hc_ai")] = f"={R('hc0')}*(1-{R('a')})+{R('incr')}*1000000/({R('rate')}*{I('util')}*{I('hrs')})/1000"
        S[R("sv")] = f"={R('r0')}-{R('eros')}+{R('incr')}"
        S[R("hw")] = f"={P('hw')}*(1+{I('ghw26') if y == 'FY2026' else I('ghw')})"
        S[R("ub")] = f"={P('ub')}*(1+{I('gub26') if y == 'FY2026' else I('gub')})"
        S[R("el")] = f"=({R('sv')}+{R('hw')}+{R('ub')})*{QL('seg_el')}/({QL('seg_sv')}+{QL('seg_hw')}+{QL('seg_ub')})"
        S[R("rev")] = f"={R('sv')}+{R('hw')}+{R('ub')}+{R('el')}"
        S[R("chk1")] = f"={R('sv')}+{R('hw')}+{R('ub')}+{R('el')}-{R('rev')}"
SEG_END = 4 + 3 * (len(SEGLAB) + 4)
S.cell(SEG_END, 1, "註：推論成本取 Tokenomics 市場價（有效牌價）；公司未揭露 token 量與代理任務量，故每任務 token 取 Tokenomics 任務 5 值（類比），不以聊天基數乘倍數。全成本對照本版未列。Elimination 營收以 LTM 占三分部營收比例持平。").font = BOLD

# ---------------- PL 頁
P = SH["PL"]
P["A1"] = "損益表：官方報告分部四列（Service Solutions、Hardware Solutions、Ubiquitous Solutions、Elimination/Corporate；決算短信 p.8），分部利益為公司自訂的調整後營業利益；合併營業利益為 IFRS（GAAP；調整後減非常態項目），調整後營業利益為 Non-GAAP 指標。¥百萬。"
P["A2"] = "FY2024A–FY2025A 取自 Q 頁；FY2026E–FY2032E 為情境預測（FY2026 為首個預測年，含 2026 上半年實績）。毛利率持平於 LTM，營業費用為殘差（含其他損益）。分部利益率持平於 LTM，Elimination/Corporate 持平於 LTM 加調整。"
PLB = ["rev", "oi_sv", "oi_hw", "oi_ub", "oi_co", "adj", "oneoff", "ebit", "inf_c", "gp", "cogs", "opex", "da", "capex", "nopat", "dwc", "fcff", "oth", "ebt", "taxx", "ni", "nci", "ni_own", "shs", "eps", "gate", "chk_oi"]
PLL = {"rev": "營收", "oi_sv": "Service Solutions 調整後營業利益", "oi_hw": "Hardware Solutions 調整後營業利益", "oi_ub": "Ubiquitous Solutions 調整後營業利益", "oi_co": "Elimination/Corporate 調整後營業利益",
       "adj": "調整後營業利益（四列合計）", "oneoff": "非常態與調整項目", "ebit": "營業利益（IFRS）", "inf_c": "其中：代理推論成本（Service 成本內）", "gp": "毛利（持平 LTM 毛利率，扣代理推論成本）", "cogs": "營業成本", "opex": "營業費用（殘差，含其他損益）",
       "da": "折舊攤銷與減損", "capex": "資本支出", "nopat": "NOPAT（IFRS 營業利益×(1−稅率)）", "dwc": "營運資金增加", "fcff": "無槓桿自由現金流 FCFF", "oth": "業外淨額", "ebt": "稅前利益", "taxx": "所得稅", "ni": "淨利（含非控制權益）",
       "nci": "非控制權益", "ni_own": "歸屬母公司淨利", "shs": "稀釋股數（M）", "eps": "EPS（IFRS 稀釋，日圓）", "gate": "收支（FCFF＋稅後業外淨額−股利）", "chk_oi": "檢查：四列分部調整後營業利益合計 − 調整後營業利益"}
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
            P[R("rev")] = f"={QH('rev', y)}"
            for k in ("oi_sv", "oi_hw", "oi_ub", "oi_co"): P[R(k)] = f"={QH(k, y)}"
            P[R("adj")] = f"={QH('adj', y)}"; P[R("ebit")] = f"={QH('ebit', y)}"; P[R("oneoff")] = f"={R('adj')}-{R('ebit')}"
            P[R("cogs")] = f"={QH('cogs', y)}"; P[R("gp")] = f"={QH('gp', y)}"; P[R("opex")] = f"={R('gp')}-{R('ebit')}"
            P[R("da")] = f"={QH('da', y)}"; P[R("capex")] = f"={QH('capex', y)}"
            P[R("ebt")] = f"={QH('ebt', y)}"; P[R("oth")] = f"={R('ebt')}-{R('ebit')}"
            P[R("ni_own")] = f"={QH('ni', y)}"; P[R("eps")] = f"={QH('eps', y)}"
            P[R("chk_oi")] = f"={R('oi_sv')}+{R('oi_hw')}+{R('oi_ub')}+{R('oi_co')}-{R('adj')}"
            continue
        P[R("rev")] = f"=Seg!{c}{sr['rev']}"
        P[R("inf_c")] = f"=Seg!{c}{sr['inf']}"
        # Service：成本率（1−LTM 利益率）× 自動化後工作量；利益＝營收−成本−推論
        P[R("oi_sv")] = f"=Seg!{c}{sr['sv']}-(1-{QL('oi_sv')}/{QL('seg_sv')})*(Seg!{c}{sr['r0']}*(1-{I('cs')}*Seg!{c}{sr['a']})+Seg!{c}{sr['incr']})-{R('inf_c')}"
        P[R("oi_hw")] = f"=Seg!{c}{sr['hw']}*{QL('oi_hw')}/{QL('seg_hw')}"
        P[R("oi_ub")] = f"=Seg!{c}{sr['ub']}*{QL('oi_ub')}/{QL('seg_ub')}"
        P[R("oi_co")] = f"={QL('oi_co')}+{I('corp_dx')}"
        P[R("adj")] = f"={R('oi_sv')}+{R('oi_hw')}+{R('oi_ub')}+{R('oi_co')}"
        P[R("oneoff")] = f"={I('oneoff')}"
        P[R("ebit")] = f"={R('adj')}-{R('oneoff')}"
        P[R("gp")] = f"={R('rev')}*{I('gm')}-{R('inf_c')}"
        P[R("cogs")] = f"={R('rev')}-{R('gp')}"
        P[R("opex")] = f"={R('gp')}-{R('ebit')}"
        P[R("da")] = f"={R('rev')}*{I('da_pct')}"; P[R("capex")] = f"={R('rev')}*{I('capex_pct')}"
        P[R("nopat")] = f"={R('ebit')}*(1-{I('tax')})"
        P[R("dwc")] = f"={I('wc_k')}*({R('rev')}-{Pv('rev')})"
        P[R("fcff")] = f"={R('nopat')}+{R('da')}-{R('capex')}-{R('dwc')}"
        P[R("oth")] = f"={I('oth_net')}"
        P[R("ebt")] = f"={R('ebit')}+{R('oth')}"; P[R("taxx")] = f"={R('ebt')}*{I('tax')}"; P[R("ni")] = f"={R('ebt')}-{R('taxx')}"
        P[R("nci")] = f"={R('ni')}*{I('nci_pct')}"; P[R("ni_own")] = f"={R('ni')}-{R('nci')}"
        P[R("shs")] = f"={I('sh26')}" if y == "FY2026" else f"={Pv('shs')}*(1-{I('bb')})"
        P[R("eps")] = f"={R('ni_own')}/{R('shs')}"
        P[R("gate")] = f"={R('fcff')}+{R('oth')}*(1-{I('tax')})-{I('dps')}*{R('shs')}"
        P[R("chk_oi")] = f"={R('oi_sv')}+{R('oi_hw')}+{R('oi_ub')}+{R('oi_co')}-{R('adj')}"
b = PR["B"]
STD = ["營收", "營收年增率", "營業成本", "毛利", "毛利率", "營業費用", "營業利益", "營業利益率", "其他收入", "稅前利益", "所得稅", "稅後淨利", "淨利率", "EPS（IFRS 稀釋，日圓）", "淨利年增率"]
P.cell(4, 1, "標準損益表（IFRS；¥百萬，EPS 為日圓；情境 B）：FY2024A–FY2025A 實績、FY2026E–FY2032E 預測；營業利益為 IFRS（調整後營業利益另見分部段）；稅後淨利＝歸屬母公司").font = BOLD
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
    P[f"{c}{STR['其他收入']}"] = f"={ref('oth')}"; P[f"{c}{STR['稅前利益']}"] = f"={ref('ebt')}"
    P[f"{c}{STR['所得稅']}"] = f"={ref('taxx')}" if y not in HIST else "未找到"
    P[f"{c}{STR['稅後淨利']}"] = f"={ref('ni_own')}"
    P[f"{c}{STR['淨利率']}"] = f"={c}{STR['稅後淨利']}/{c}{STR['營收']}"
    P[f"{c}{STR['EPS（IFRS 稀釋，日圓）']}"] = f"={ref('eps')}"
    P[f"{c}{STR['淨利年增率']}"] = f"={c}{STR['稅後淨利']}/{p}{STR['稅後淨利']}-1" if y != "FY2024" else "未找到"
P.cell(STR["淨利年增率"] + 1, 1, "FY2025 稅後淨利含 Device Solutions 出售等非繼續營業利益（歸屬母公司 449,408 對稅前持續營業 409,034），FY2026E 年增率因此為負，不代表獲利下滑；FY2024、FY2025 所得稅未找到。")
P.cell(23, 1, "分部營收與調整後營業利益（情境 B；官方報告分部）").font = BOLD
SEGP = {"Service Solutions": ("sv", "rev"), "Hardware Solutions": ("hw", "rev"), "Ubiquitous Solutions": ("ub", "rev"), "Elimination/Corporate": ("el", "rev"), "營收合計（四列）": (None, "rev"),
        "Service Solutions 調整後營業利益": ("oi_sv", "oi"), "Hardware Solutions 調整後營業利益": ("oi_hw", "oi"), "Ubiquitous Solutions 調整後營業利益": ("oi_ub", "oi"), "Elimination/Corporate 調整後營業利益": ("oi_co", "oi"), "調整後營業利益合計（四列）": (None, "oi")}
SPR = {}
for i, (lab, (k, kind)) in enumerate(SEGP.items()):
    rw = 24 + i; SPR[lab] = rw; P.cell(rw, 1, lab)
    for y, c in YC.items():
        if k and kind == "rev": P[f"{c}{rw}"] = f"=Seg!{c}{SR['B'][k]}"
        elif k: P[f"{c}{rw}"] = f"={c}{b[k]}"
        else: P[f"{c}{rw}"] = f"=SUM({c}{rw-4}:{c}{rw-1})"
CK = 24 + len(SEGP) + 1
P.cell(CK - 1, 1, "對帳（應為 0）").font = BOLD
CHKL = ["檢查：分部營收合計 − 合併營收", "檢查：分部調整後營業利益合計 − 調整後營業利益", "檢查：分部營收合計 − 公司揭露合併營收（歷史年度）", "檢查：分部調整後營業利益合計 − 公司揭露調整後營業利益（歷史年度）", "檢查：三情境分部營收合計 − 合併營收"]
for i, l in enumerate(CHKL): P.cell(CK + i, 1, l)
for y, c in YC.items():
    P[f"{c}{CK}"] = f"={c}{SPR['營收合計（四列）']}-{c}{STR['營收']}"
    P[f"{c}{CK+1}"] = f"={c}{SPR['調整後營業利益合計（四列）']}-{c}{b['adj']}"
    P[f"{c}{CK+2}"] = f"={c}{SPR['營收合計（四列）']}-{QH('rev', y)}" if y in HIST else 0
    P[f"{c}{CK+3}"] = f"={c}{SPR['調整後營業利益合計（四列）']}-{QH('adj', y)}" if y in HIST else 0
    P[f"{c}{CK+4}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['chk1']})" for s in "ABC") + ")"

# ---------------- DCF 頁
Dp = SH["DCF"]
Dp["A1"] = "DCF（FCFF 標準規格）：FY2026E～FY2032E；永續價值＝NOPAT(FY2033)×(1−g÷RONIC)÷(WACC−g)；期中折現。現金流自資產負債表日（2026-09-30）起算（FY2026 計入下半年），折現到股價基準日；三情境各算一次。"
Dp["A2"] = "WACC 由 Val 頁推出（讀 In 頁基準欄）。年度起訖日：FY(y)＝y-04-01～(y+1)-03-31。FY2026 的 FCFF 以全年預測乘計入比例近似下半年。"
DR = {}
KEYS = ["start", "end", "fcff", "frac", "t", "df", "pv"]
KLAB = ["年度起日", "年度迄日", "FCFF", "計入比例（資產負債表日之後）", "折現時點（年；自股價基準日到期中）", "折現因子", "現值"]
for si, s in enumerate("ABC"):
    r0 = 4 + si * 11
    DR[s] = {k: r0 + 1 + i for i, k in enumerate(KEYS)}
    Dp.cell(r0, 1, f"情境 {s}").font = BOLD
    for y in FC: Dp[f"{YC[y]}{r0}"] = y + "E"
    for k, lab in zip(KEYS, KLAB): Dp.cell(DR[s][k], 1, f"情境 {s}｜{lab}")
    for y in FC:
        c = YC[y]; yy = int(y[2:])
        Dp[f"{c}{DR[s]['start']}"] = f"=DATE({yy},4,1)"; Dp[f"{c}{DR[s]['end']}"] = f"=DATE({yy+1},3,31)"
        for k in ("start", "end"): Dp[f"{c}{DR[s][k]}"].number_format = "yyyy-mm-dd"
        Dp[f"{c}{DR[s]['fcff']}"] = f"=PL!{c}{PR[s]['fcff']}"
        st, en = f"{c}{DR[s]['start']}", f"{c}{DR[s]['end']}"
        Dp[f"{c}{DR[s]['frac']}"] = f"=MAX(0,MIN(1,({en}-MAX({st}-1,{I('bs_date')}))/({en}-{st}+1)))"
        Dp[f"{c}{DR[s]['t']}"] = f"=MAX(0,((MAX({st},{I('bs_date')}+1)+{en})/2-{I('px_date')})/365)"
        Dp[f"{c}{DR[s]['pv']}"] = f"={c}{DR[s]['fcff']}*{c}{DR[s]['frac']}*{c}{DR[s]['df']}"
    Dp.cell(DR[s]["pv"] + 1, 1, f"情境 {s}｜永續價值折現時點（年；FY2032 年底）"); Dp[f"K{DR[s]['pv']+1}"] = f"=(K{DR[s]['end']}-{I('px_date')})/365"
    Dp.cell(DR[s]["pv"] + 2, 1, f"情境 {s}｜檢查：FCFF 各列加總（NOPAT＋折舊−資本支出−營運資金增加）−FCFF（應為 0）")
    for y in FC:
        c = YC[y]
        Dp[f"{c}{DR[s]['pv']+2}"] = f"={c}{DR[s]['fcff']}-(PL!{c}{PR[s]['nopat']}+PL!{c}{PR[s]['da']}-PL!{c}{PR[s]['capex']}-PL!{c}{PR[s]['dwc']})"
GATE = 4 + 3 * 11 + 1
Dp.cell(GATE, 1, "收支關卡：自由現金流扣股利（FY2026–FY2030 五年合計；¥百萬；欄 B／C／D＝情境 A／B／C）")
for col, s in zip("BCD", "ABC"): Dp[f"{col}{GATE}"] = f"=SUM(PL!E{PR[s]['gate']}:I{PR[s]['gate']})"
Dp.cell(GATE + 1, 1, "檢查：收支關卡 B − 逐年加總（應為 0）"); Dp[f"B{GATE+1}"] = f"=C{GATE}-(PL!E{PR['B']['gate']}+PL!F{PR['B']['gate']}+PL!G{PR['B']['gate']}+PL!H{PR['B']['gate']}+PL!I{PR['B']['gate']})"

# ---------------- Val 頁
V = SH["Val"]
VR = {}
vrows = ["intro", "hint", "blank0", "cap", "px", "px_date", "tp_date", "dcf_tp", "comp_tp", "tp", "up", "valr", "gate", "rec", "tp_ref", "sp1", "wh", "ke", "kd", "kdat", "wacc", "ronic", "g", "emkt", "dmkt", "ewt",
         "sp2", "ndh", "cashx", "ndebt", "invx", "nciv", "netcash", "sh", "dps", "sp3", "sumh", "pvx", "tvpv", "ev", "eq", "vps0", "dcf12", "tvsh", "sp4", "ch", "ttm_sv", "ttm_hw", "ttm_ub", "ttm_co", "ttm_tot", "mult", "ev_c", "comp12",
         "sp8", "g3h", "g3_dcf", "g3_ev", "g3_rate", "g3_rev", "g3_om", "g3_flag", "sp7", "gaph", "gap", "gap_t", "gap_m", "gap_sbc", "gap_lev", "ownh", "own_ev", "own_mult", "own_ratio", "rerate", "sp5", "scenh", "wa", "wb", "wc", "chk_w",
         "sp6", "peerh", "pe1", "pe2", "pe3", "pe4", "pe_med", "pe_n", "pe_own", "stale_rf", "pe_note"]
for i, k in enumerate(vrows): VR[k] = 1 + i
RCR_REV26, RCR_OM26 = 5, 6
V.cell(VR["intro"], 1, "估值：DCF（FCFF）與 Comparable（分部加總：各分部目標時點 TTM 營業利益 × 同業 EV／EBIT）加權；目標價為 12 個月後（股價基準日＋12 個月）。三情境 25／50／25。上檔空間不是年化報酬。")
V.cell(VR["hint"], 1, "Comparable 同業只有 4 家（少於 8 家，不做迴歸）、日期與準則不一致；倍數分母為 TTM 營業利益（EV／EBIT，無槓桿對無槓桿；本公司以 IFRS 營業利益，非調整後），故乘目標日的 TTM 營業利益；Hardware 與 Ubiquitous 取 Service 倍數的折扣。")
V.cell(VR["cap"], 1, "結論（情境加權）").font = BOLD


def vrow(k, label, b=None, c=None, d=None, e=None):
    r_ = VR[k]; V.cell(r_, 1, label)
    for col, v in zip("BCDE", (b, c, d, e)):
        if v is not None: V[f"{col}{r_}"] = v


wr = lambda s: I("tp_w" + s)
vrow("px", "現價（¥）", f"={I('px')}")
vrow("px_date", "股價基準日", f"={I('px_date')}"); V[f"B{VR['px_date']}"].number_format = "yyyy-mm-dd"
vrow("tp_date", "目標時點（股價基準日＋12 個月）", f"=EDATE({I('px_date')},12)"); V[f"B{VR['tp_date']}"].number_format = "yyyy-mm-dd"
vrow("dcf_tp", "DCF 目標價（情境加權；¥）", f"={wr('A')}*C{VR['dcf12']}+{wr('B')}*D{VR['dcf12']}+{wr('C')}*E{VR['dcf12']}")
vrow("comp_tp", "Comparable 目標價（分部加總；情境加權；¥）", f"={wr('A')}*C{VR['comp12']}+{wr('B')}*D{VR['comp12']}+{wr('C')}*E{VR['comp12']}")
vrow("tp", "加權目標價（DCF 與 Comparable 加權；¥）", f"={I('w_dcf')}*B{VR['dcf_tp']}+(1-{I('w_dcf')})*B{VR['comp_tp']}")
vrow("up", "上檔空間（相對現價）", f"=B{VR['tp']}/B{VR['px']}-1")
tier = lambda u: f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'
vrow("valr", "估值評等（DCF 與 Comparable 加權，只看上檔空間）", "=" + tier(f"B{VR['up']}"))
vrow("gate", "入不敷出判定用：基準情境 FY2026–FY2030 收支合計（¥百萬）", f"=DCF!C{GATE}")
vrow("rec", "投資建議", f'=IF(AND(B{VR["gate"]}<0,OR(B{VR["valr"]}="買進",B{VR["valr"]}="強力買進")),"收支不足",B{VR["valr"]})')
vrow("tp_ref", "只看 Comparable 的目標價（上檔空間見右）", f"=B{VR['comp_tp']}", f"=B{VR['comp_tp']}/B{VR['px']}-1")
V.cell(VR["wh"], 1, "折現率（DCF）").font = BOLD
vrow("ke", "股權成本 ke（rf＋β×ERP；β＝Blume 調整＝0.67×原始＋0.33）", f"={I('rf')}+(0.67*{I('beta_raw')}+0.33)*{I('erp')}")
vrow("kd", "稅前債務成本（rf＋利差）", f"={I('rf')}+{I('kd_sp')}")
vrow("kdat", "稅後債務成本", f"=B{VR['kd']}*(1-{I('tax')})")
vrow("emkt", "股權市值（¥百萬；評價股數 × 現價）", f"={I('px')}*B{VR['sh']}")
vrow("dmkt", "債務與租賃負債（¥百萬）", f"={I('debt')}")
vrow("ewt", "股權權重", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})")
vrow("wacc", "WACC", f"=B{VR['ewt']}*B{VR['ke']}+(1-B{VR['ewt']})*B{VR['kdat']}")
vrow("ronic", "RONIC（＝WACC＋利差）", f"=B{VR['wacc']}+{I('ronic_sp')}")
vrow("g", "永續成長率", f"={I('g_tv')}")
V.cell(VR["ndh"], 1, "淨現金與非營運資產（2026-09-30 財政狀態表；¥百萬）").font = BOLD
vrow("cashx", "現金及約當現金", f"={I('cash')}")
vrow("ndebt", "債務與租賃負債（IFRS 16 合併列示）", f"={I('debt')}")
vrow("invx", "非營運投資（權益法投資＋其他投資）× 計入比例", f"=({I('inv_eq')}+{I('inv_oth')})*{I('inv_hc')}")
vrow("nciv", "非控制權益（自股權價值扣除）", f"={I('nci')}")
vrow("netcash", "淨現金加非營運投資減非控制權益（負值＝淨債務）", f"=B{VR['cashx']}-B{VR['ndebt']}+B{VR['invx']}-B{VR['nciv']}")
vrow("sh", "評價股數（M；已發行−庫藏股＋潛在稀釋）", f"={I('sh_issued')}-{I('sh_tr')}+{I('sh_dil')}")
vrow("dps", "期間股利（¥/股；持有期間）", f"={I('dps')}")
V.cell(VR["sumh"], 1, "DCF 彙總（三情境）").font = BOLD
for col, lab in zip("CDE", ("情境 A", "情境 B", "情境 C")): V[f"{col}{VR['sumh']}"] = lab
for si, s in enumerate("ABC"):
    col = "CDE"[si]; pvr = DR[s]["pv"]
    V[f"{col}{VR['pvx']}"] = f"=SUM(DCF!E{pvr}:K{pvr})"
    V[f"{col}{VR['tvpv']}"] = f"=PL!K{PR[s]['nopat']}*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^DCF!K{pvr+1}"
    V[f"{col}{VR['ev']}"] = f"={col}{VR['pvx']}+{col}{VR['tvpv']}"
    V[f"{col}{VR['eq']}"] = f"={col}{VR['ev']}+$B${VR['netcash']}"
    V[f"{col}{VR['vps0']}"] = f"={col}{VR['eq']}/$B${VR['sh']}"
    V[f"{col}{VR['dcf12']}"] = f"={col}{VR['vps0']}*(1+$B${VR['ke']})-$B${VR['dps']}"
    V[f"{col}{VR['tvsh']}"] = f"={col}{VR['tvpv']}/{col}{VR['ev']}"
    w = f"MIN(1,(DCF!$E${DR[s]['end']}-{I('px_date')})/365)"
    for rowk, k in (("ttm_sv", "oi_sv"), ("ttm_hw", "oi_hw"), ("ttm_ub", "oi_ub"), ("ttm_co", "oi_co")):
        V[f"{col}{VR[rowk]}"] = f"={w}*PL!E{PR[s][k]}+(1-{w})*PL!F{PR[s][k]}"
    V[f"{col}{VR['ttm_co']}"] = f"={w}*(PL!E{PR[s]['oi_co']}-PL!E{PR[s]['oneoff']})+(1-{w})*(PL!F{PR[s]['oi_co']}-PL!F{PR[s]['oneoff']})"
    V[f"{col}{VR['ttm_tot']}"] = f"=SUM({col}{VR['ttm_sv']}:{col}{VR['ttm_co']})"
    V[f"{col}{VR['mult']}"] = f"=$B${VR['pe_med']}*{I('cm_adj')}"
    V[f"{col}{VR['ev_c']}"] = f"=({col}{VR['ttm_sv']}+{col}{VR['ttm_co']}+({col}{VR['ttm_hw']}+{col}{VR['ttm_ub']})*{I('hw_disc')})*{col}{VR['mult']}"
    V[f"{col}{VR['comp12']}"] = f"=({col}{VR['ev_c']}+$B${VR['netcash']})/$B${VR['sh']}"
    V[f"{col}{VR['wb']}"] = f"={I('w_dcf')}*{col}{VR['dcf12']}+(1-{I('w_dcf')})*{col}{VR['comp12']}"
    V[f"{col}{VR['wc']}"] = f"={col}{VR['wb']}/$B${VR['px']}-1"
vrow("pvx", "明確期 FCFF 現值（¥百萬）"); vrow("tvpv", "永續價值現值（¥百萬）"); vrow("ev", "企業價值 EV（¥百萬）"); vrow("eq", "股權價值（EV＋淨現金與非營運資產減非控制權益；¥百萬）")
vrow("vps0", "每股價值（評價日；¥）"); vrow("dcf12", "DCF 情境目標價（12 個月後＝每股價值×(1+ke)−股利；¥）"); vrow("tvsh", "永續價值占 EV 比例")
V.cell(VR["ch"], 1, "Comparable（分部加總：目標日 TTM 營業利益 × 同業 EV／EBIT〔分母＝TTM 營業利益〕）").font = BOLD
vrow("ttm_sv", "Service Solutions 目標日 TTM 調整後營業利益（¥百萬）"); vrow("ttm_hw", "Hardware Solutions 目標日 TTM 調整後營業利益（¥百萬）"); vrow("ttm_ub", "Ubiquitous Solutions 目標日 TTM 調整後營業利益（¥百萬）")
vrow("ttm_co", "Elimination/Corporate 扣非常態項目後目標日 TTM（¥百萬）"); vrow("ttm_tot", "合計＝TTM IFRS 營業利益（¥百萬）")
vrow("mult", "套用倍數（同業 EV／EBIT 中位數×調整係數；Service 與全社用此倍數，Hardware／Ubiquitous 再乘折扣）"); vrow("ev_c", "Comparable 企業價值（¥百萬）"); vrow("comp12", "Comparable 情境目標價（¥；淨現金與非營運資產取資產負債表日值）")
V.cell(VR["g3h"], 1, "合理性關卡（工作單 G3 第 3 條；發布前必查）").font = BOLD
vrow("g3_dcf", "DCF 目標價 ÷ 現價（小於 0.25 或大於 4 觸發）", f"=B{VR['dcf_tp']}/B{VR['px']}")
vrow("g3_ev", "三情境企業價值最小值（¥百萬；負值觸發）", f"=MIN(C{VR['ev']}:E{VR['ev']})")
vrow("g3_rate", "估值評等為強力買進或強力放空（1＝觸發）", f'=IF(OR(B{VR["valr"]}="強力買進",B{VR["valr"]}="強力放空"),1,0)')
vrow("g3_rev", "FY2026 營收（情境 B）對指引差距（超過 10% 觸發）", f"=ABS(Recon!B{RCR_REV26}/AVERAGE(Recon!C{RCR_REV26}:D{RCR_REV26})-1)")
vrow("g3_om", "FY2026 調整後營業利益率對指引差距（相對值，超過 10% 觸發）", f"=ABS(Recon!B{RCR_OM26}/AVERAGE(Recon!C{RCR_OM26}:D{RCR_OM26})-1)")
vrow("g3_flag", "合理性關卡是否觸發（1＝觸發）", f"=IF(OR(B{VR['g3_dcf']}<0.25,B{VR['g3_dcf']}>4,B{VR['g3_ev']}<0,B{VR['g3_rate']}=1,B{VR['g3_rev']}>0.1,B{VR['g3_om']}>0.1),1,0)")
V.cell(VR["gaph"], 1, "DCF 與 Comparable 落差分解（情境 B；每股 ¥；公式組成）").font = BOLD
vrow("gap", "落差：DCF 情境 B 目標價 − Comparable 情境 B 目標價", f"=D{VR['dcf12']}-D{VR['comp12']}")
vrow("gap_t", "其中時點：DCF 自評價日滾到 12 個月後（每股價值×ke−股利）", f"=D{VR['dcf12']}-D{VR['vps0']}")
vrow("gap_m", "其中倍數：評價日 DCF 每股價值 − Comparable 目標價（DCF 隱含的 EV／EBIT 對同業倍數）", f"=D{VR['vps0']}-D{VR['comp12']}")
vrow("gap_sbc", "其中股票薪酬：兩法皆以 IFRS 營業利益（已扣股票報酬）為基礎", 0)
vrow("gap_lev", "其中槓桿：兩法皆為企業價值（無槓桿）加同一淨現金", 0)
V.cell(VR["ownh"], 1, "本公司現行倍數 ÷ 同業中位數與隱含重估（原則第 47 條）").font = BOLD
vrow("own_ev", "富士通現行企業價值（現價 × 評價股數 ＋ 債務 ＋ 非控制權益 − 現金 − 非營運投資；¥百萬）", f"=B{VR['emkt']}+B{VR['dmkt']}+B{VR['nciv']}-B{VR['cashx']}-B{VR['invx']}")
vrow("own_mult", "富士通現行 EV／LTM IFRS 營業利益（LTM 至 2026-09：Q 頁）", f"=B{VR['own_ev']}/{QL('ebit')}")
vrow("own_ratio", "現行倍數 ÷ 同業中位數", f"=B{VR['own_mult']}/B{VR['pe_med']}")
vrow("rerate", "若回到同業中位數（調整前、四分部同倍數）的隱含股價變動", f"=(B{VR['pe_med']}*{QL('ebit')}+B{VR['netcash']})/B{VR['sh']}/B{VR['px']}-1")
V.cell(VR["scenh"], 1, "情境權重與加權").font = BOLD
vrow("wa", "情境權重 A／B／C", None, f"={wr('A')}", f"={wr('B')}", f"={wr('C')}")
vrow("wb", "情境加權目標價（各情境 DCF 與 Comparable 加權；¥）"); vrow("wc", "各情境上檔空間")
vrow("chk_w", "檢查：情境加權目標價與加權目標價之差（應為 0）", f"=ABS({wr('A')}*C{VR['wb']}+{wr('B')}*D{VR['wb']}+{wr('C')}*E{VR['wb']}-B{VR['tp']})")
V.cell(VR["peerh"], 1, "同業倍數（StockAnalysis；EV／EBIT，分母＝TTM 營業利益；日期不一致；家數 4，少於 8 不做迴歸）").font = BOLD
PEERS = [("日立製作所（6501）", 14.34, datetime.date(2026, 11, 6), "IFRS；租賃折舊與利息分列；EV 含租賃負債與否未核對"),
         ("NEC（6701）", 15.77, datetime.date(2026, 11, 13), "IFRS；同上"),
         ("野村総合研究所（4307）", 20.93, datetime.date(2026, 11, 9), "IFRS；同上；高成長 IT 顧問，倍數偏高"),
         ("BIPROGY（8056）", 11.21, datetime.date(2026, 10, 5), "日本會計準則；營業租賃費用在營業利益內，與 IFRS 口徑不同")]
for i, (nm, mv, dt, note) in enumerate(PEERS):
    rw = VR[f"pe{i+1}"]
    V.cell(rw, 1, nm); V.cell(rw, 2, mv).font = BLUE; V.cell(rw, 3, dt).number_format = "yyyy-mm-dd"; V.cell(rw, 4, note)
    V.cell(rw, 5, f'=IF(NETWORKDAYS(MIN(C{rw},{I("px_date")}),MAX(C{rw},{I("px_date")}))-1>2,"逾期","")')
vrow("pe_med", "同業中位數 EV／EBIT（TTM）", f"=MEDIAN(B{VR['pe1']}:B{VR['pe4']})")
vrow("pe_n", "家數", f"=COUNT(B{VR['pe1']}:B{VR['pe4']})")
V.cell(VR["pe_own"], 1, "富士通自身 EV／EBIT（StockAnalysis 2026-10-23，股價 4,477 日圓時；僅對照，已過時）"); V.cell(VR["pe_own"], 2, 17.55).font = BLUE
V.cell(VR["stale_rf"], 1, "rf 基準日與現價日相差（交易日；超過 2 個交易日標逾期）")
V[f"B{VR['stale_rf']}"] = f"=NETWORKDAYS(MIN({I('rf_date')},{I('px_date')}),MAX({I('rf_date')},{I('px_date')}))-1"
V[f"C{VR['stale_rf']}"] = f'=IF(B{VR["stale_rf"]}>2,"逾期","未逾期")'
V.cell(VR["pe_note"], 1, "同業倍數基準日早於現價日 2 個交易日以上者標「逾期」，使用前須更新；本版未更新。交易日以週一至週五計，未扣日本假日。")
for s in "ABC":
    for y in FC:
        c = YC[y]
        Dp[f"{c}{DR[s]['df']}"] = f"=1/(1+Val!$B${VR['wacc']})^{c}{DR[s]['t']}"

# ---------------- RevDCF 頁
RD = SH["RevDCF"]
RV = {}
RD["A1"] = "反向 DCF：現價隱含什麼。以情境 B 路徑為底，兩個封閉解（DCF 對營業利益率與營收規模都是線性）；主導變數的整本重算反解在 Recon 第 4 節。"
RD["A2"] = "讀法：隱含值是「DCF 每股價值（評價日）＝現價」的點，不是評等門檻。"
rv_rows = [("lab", "項目"), ("v0", "情境 B 每股價值（評價日；¥）"), ("px", "現價（¥）"), ("gap", "現價與 DCF 每股價值之差（股權價值；¥百萬）"),
           ("k_m", "營業利益率全期平移 1 個百分點對股權價值的影響（¥百萬）"), ("d_m", "現價隱含的營業利益率平移（百分點）"), ("m_fy32", "情境 B FY2032E 營業利益率（IFRS）"), ("m_imp", "現價隱含的 FY2032 營業利益率"),
           ("ev0", "情境 B 企業價值（¥百萬）"), ("ev_imp", "現價隱含的企業價值（¥百萬）"), ("scale", "現價隱含的營收規模倍數（相對情境 B）"),
           ("mult_imp", "現價隱含的 EV／TTM 營業利益倍數（情境 B；目標日 TTM）"), ("mult_peer", "同業中位數 EV／EBIT"), ("ps", "分部加總：Service 與全社共用同業倍數、Hardware 與 Ubiquitous 為折扣，隱含倍數為合計口徑")]
for i, (k, lab) in enumerate(rv_rows): RV[k] = 4 + i; RD.cell(4 + i, 1, lab)
b_ = PR["B"]
RD[f"B{RV['v0']}"] = f"=Val!D{VR['vps0']}"; RD[f"B{RV['px']}"] = f"=Val!B{VR['px']}"
RD[f"B{RV['gap']}"] = f"=(B{RV['px']}-B{RV['v0']})*Val!B{VR['sh']}"
terms = "+".join(f"PL!{YC[y]}{b_['rev']}*DCF!{YC[y]}{DR['B']['frac']}*DCF!{YC[y]}{DR['B']['df']}" for y in FC)
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
rec_rows = [
 ("h1", "1. 合併層：FY2026（情境 B）", None),
 ("g26", "營收年增率", (f"=PL!E{STR['營收年增率']}", f"=C{{rev}}/PL!D{STR['營收']}-1", f"=D{{rev}}/PL!D{STR['營收']}-1", "未找到", "")),
 ("rev26", "營收（¥百萬）", (f"=PL!E{STR['營收']}", 3510000, 3510000, 3560000, "")),
 ("om26", "調整後營業利益率", (f"=PL!E{b_['adj']}/PL!E{STR['營收']}", "=425000/3510000", "=425000/3510000", "未找到", "")),
 ("adj26", "調整後營業利益（¥百萬）", (f"=PL!E{b_['adj']}", 425000, 425000, "未找到", "")),
 ("ebit26", "營業利益（IFRS；¥百萬）", (f"=PL!E{STR['營業利益']}", 415000, 415000, "未找到", "公司指引 IFRS 營業利益 415,000（決算短信 2026Q2 封面）")),
 ("ni26", "歸屬母公司當期利益（¥百萬）", (f"=PL!E{STR['稅後淨利']}", 310000, 310000, "未找到", "公司指引 310,000（決算短信 2026Q2）；本模型稅率與業外淨額為假設")),
 ("sv26", "Service Solutions 營收（¥百萬）", (f"=Seg!E{SR['B']['sv']}", 2470000, 2470000, "未找到", "公司分部指引 2,470,000／調整後營業利益 430,000（決算短信 FY2025，2026-04-28；2026Q2 是否修正未逐項核對）")),
 ("svo26", "Service Solutions 調整後營業利益（¥百萬）", (f"=PL!E{b_['oi_sv']}", 430000, 430000, "未找到", "")),
 ("hw26", "Hardware Solutions 營收（¥百萬）", (f"=Seg!E{SR['B']['hw']}", 960000, 960000, "未找到", "分部指引 960,000／62,000")),
 ("ub26", "Ubiquitous Solutions 營收（¥百萬）", (f"=Seg!E{SR['B']['ub']}", 160000, 160000, "未找到", "分部指引 160,000／28,000")),
 ("h2", "2. 2026 上半年進度（實績對模型全年）", None),
 ("h1p", "2026 上半年營收占模型 FY2026 全年比例（對 2025 上半年占 FY2025 比例）", (f"={QL('rev')}*0+Q!E{QR['rev']}/PL!E{STR['營收']}", f"=Q!D{QR['rev']}/Q!C{QR['rev']}", "", "", "2025 上半年占全年 44.7%（Derived）；下半年占比偏高為季節性")),
 ("h1o", "2026 上半年調整後營業利益占模型 FY2026 全年比例（對 2025 上半年占比）", (f"=Q!E{QR['adj']}/PL!E{b_['adj']}", f"=Q!D{QR['adj']}/Q!C{QR['adj']}", "", "", "下半年利益占比高為季節性")),
 ("h3", "3. 市場共識與賣方目標價（Interested-party）", None),
 ("c1", "FY2026 營收共識（¥百萬）", (f"=PL!E{STR['營收']}", None, None, 3560000, "共識 3.56 兆日圓（範圍 3.5–3.6 兆；StockAnalysis 預測頁，2026-09）")),
 ("tp1", "賣方平均目標價（¥）", ("=Val!B%d" % VR["tp"], None, None, 4802, "平均 4,802 日圓（13 位；2026-09）；2026-10-29 決算後的更新未找到；現價 2026-11-13 較 2026-10-23 低約 23%")),
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
assert RCR["rev26"] == RCR_REV26 and RCR["om26"] == RCR_OM26
Rc[f"C{RCR['g26']}"] = f"=C{RCR['rev26']}/PL!D{STR['營收']}-1"; Rc[f"D{RCR['g26']}"] = f"=D{RCR['rev26']}/PL!D{STR['營收']}-1"
Rc[f"B{RCR['h1p']}"] = f"=Q!E{QR['rev']}/PL!E{STR['營收']}"
Rc[f"F{RCR['g26']}"] = (f'="本模型 FY2026 營收 "&TEXT(B{RCR["rev26"]}/1000000,"0.000")&" 兆日圓（年增 "&TEXT(B{RCR["g26"]},"0.0%")&"），公司指引 "&TEXT(C{RCR["rev26"]}/1000000,"0.00")&" 兆；分部差異：Service "&TEXT(B{RCR["sv26"]}/1000,"#,##0")&" 十億（指引 "&TEXT(C{RCR["sv26"]}/1000,"#,##0")&"）、Hardware "&TEXT(B{RCR["hw26"]}/1000,"#,##0")&"（"&TEXT(C{RCR["hw26"]}/1000,"#,##0")&"）、Ubiquitous "&TEXT(B{RCR["ub26"]}/1000,"#,##0")&"（"&TEXT(C{RCR["ub26"]}/1000,"#,##0")&"）。"')
Rc[f"F{RCR['om26']}"] = (f'="本模型 "&TEXT(B{RCR["om26"]},"0.0%")&"，指引 "&TEXT(C{RCR["om26"]},"0.0%")&"；落差 "&TEXT((B{RCR["om26"]}-C{RCR["om26"]})*100,"+0.0;-0.0")&" 點，主因是分部利益率持平於 LTM（Service "&TEXT({QL("oi_sv")}/{QL("seg_sv")},"0.0%")&"，指引約 17.4%）與 Corporate 持平 LTM，公司指引隱含下半年毛利率回升與 AI 驅動生產力（口述 43% 目標）。"')
Rc[f"F{RCR['adj26']}"] = (f'="本模型調整後營業利益 "&TEXT(B{RCR["adj26"]}/1000,"#,##0")&" 十億，指引 "&TEXT(C{RCR["adj26"]}/1000,"#,##0")&" 十億；2026 上半年實績 "&TEXT(Q!E{QR["adj"]}/1000,"#,##0")&" 十億，需下半年 "&TEXT(C{RCR["adj26"]}/1000-Q!E{QR["adj"]}/1000,"#,##0")&" 十億才達指引。"')
r += 1
Rc.cell(r, 1, "4. 反解（整本重算）：支持現價所需的主導變數值（原則第 4 條）").font = BOLD
RC_REV = r + 1
revs = json.load(open(REVS, encoding="utf-8")) if REVS else None
Rc.cell(RC_REV, 1, "主導變數（Sens 頁 AI 命題與營運參數中對加權目標價影響最大者）")
Rc.cell(RC_REV + 1, 1, "基準值"); Rc.cell(RC_REV + 2, 1, "支持現價所需值（整本重算割線法；加權目標價＝現價）"); Rc.cell(RC_REV + 3, 1, "說明")
if revs:
    Rc.cell(RC_REV, 2, revs["label"]); Rc.cell(RC_REV + 1, 2, revs["base"]); Rc.cell(RC_REV + 2, 2, revs["solved"]); Rc.cell(RC_REV + 3, 2, revs["note"])
else:
    Rc.cell(RC_REV, 2, "未解（尚未跑反解）")
r = RC_REV + 5
Rc.cell(r, 1, "5. 口徑與未找到").font = BOLD
Rc.cell(r + 1, 1, "同業 EV／EBIT 取自 StockAnalysis，日期 2026-10-05 至 11-13，與現價日不一致、逾期者已標；家數 4；準則與租賃處理未逐家核對；未取得 NTM 口徑倍數，以 TTM 倍數乘目標日 TTM 營業利益；本公司歷史相對倍數（5 年）未自算（G3 第 4 條未做）。")
Rc.cell(r + 2, 1, "2026 下半年實績（至 2027-03）預計 2027-04 下旬公布；本版以 2026 上半年（至 2026-09-30）為最新已公布期。FY2026 DCF 的下半年現金流以全年 FCFF 乘 0.5 近似。")
Rc.cell(r + 3, 1, "未找到：分部員工數與利用率、計費工時與時薪、token 量與代理任務量（公司不揭露）、Blume β 第二來源、債券有效利率、退休金淨負債、稀釋證券明細、FY2025 所得稅與歷史營運資金比率（淨利含非繼續營業）、決算後的賣方平均目標價與共識 EPS、匯率來源；詳見 research_fujitsu.md 第 9 節。")
RC_END = r + 3

# ---------------- Sens 頁
Se = SH["Sens"]
Se["A1"] = "敏感度：In 頁每個有「敏感度低／高」區間的輸入，一次動一個、整本 LibreOffice 重算；表中為加權目標價相對基準。情境專屬列與情境權重不列（三情境本身就是情境分析）。"
Se["A2"] = "token 價格高案（cps 高）只動成本一邊，標「上界」；另列價格與自動化量聯動、模型路由（改用便宜模型）的組合案（原則第 44 條）。群組 IF_HoldEcon 不適用。"
fp = hashlib.sha256(json.dumps([(k, str(wb["In"].cell(IR[k], 5).value)) for k in sorted(IR)], ensure_ascii=False).encode()).hexdigest()[:12]
BUILD_ID = f"fujitsu-v{VER}-{fp}"
sens = json.load(open(SENS, encoding="utf-8")) if SENS else None
Se.cell(3, 1, "build_id"); Se.cell(3, 4, BUILD_ID)
Se.cell(4, 1, "模型狀態"); Se.cell(4, 4, "CURRENT" if (sens and sens.get("fingerprint") == fp) else "STALE")
Se.cell(5, 1, "基準加權目標價（本檔）"); Se.cell(5, 4, f"=Val!B{VR['tp']}")
Se.cell(6, 1, "敏感度重算時的基準加權目標價"); Se.cell(6, 4, sens["base_tp"] if sens else None)
Se.cell(7, 1, "敏感度重算時的輸入指紋"); Se.cell(7, 4, sens["fingerprint"] if sens else None)
hdr_r = 9
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（日圓）", "投資建議"], 1): Se.cell(hdr_r, j, h).font = BOLD
SE_ROWS = []
if sens:
    rr = hdr_r + 1
    for case in sens["cases"]:
        Se.cell(rr, 1, case["label"]); Se.cell(rr, 2, case["setting"]); Se.cell(rr, 3, case["side"])
        Se.cell(rr, 4, case["tp"]); Se.cell(rr, 5, f"=D{rr}/$D$5-1"); Se.cell(rr, 6, case["key"]); Se.cell(rr, 7, case["dcf"]); Se.cell(rr, 8, f"=G{rr}-Val!$B${VR['dcf_tp']}"); Se.cell(rr, 9, case["rec"])
        SE_ROWS.append(rr); rr += 1
    Se.cell(rr + 1, 1, "註：tp_w（情境權重）與 scn_*（情境驅動）不列：三情境並列已涵蓋；he_*（Tokenomics 持有成本）不適用：富士通不以持有算力為命題。")
SE_END = hdr_r + len(SE_ROWS)

# ---------------- Sources 頁
So = SH["Sources"]
So["A1"] = "來源與版本"; So.cell(2, 1, "項目"); So.cell(2, 2, "內容")
src_rows = [("Tokenomics 快照", f"{TK['source']['version']}（{TK['source']['file']}；commit {TK['source']['commit'][:7]}）；名稱：IF_TaskLen、IF_TaskTokFresh／Cached／Dec、IF_PriceFresh／Cached／Out_Terra／Sol／Astra、IF_TaskSucc_Terra／Sol／Astra（data/tokenomics_snapshot_v22.0.json）"),
            ("研究包", RES), ("公司申報", "有價證券報告書 FY2025（EDINET S100YM3K，2026-06-26）；決算短信 FY2025（2026-04-28）、2026Q1（2026-07-30）、2026Q2（2026-10-29）；決算說明資料；IR Day 2026（2026-09-16）"),
            ("同業倍數", "StockAnalysis statistics 頁（6501 2026-11-06、6701 2026-11-13、4307 2026-11-09、8056 2026-10-05、6702 2026-10-23）"), ("市場資料", "tradingeconomics 10 年日債 2026-11-16；StockAnalysis 6702 Beta；traders.co.jp 股價"),
            ("本版先驗", "未做（Andy 2026-11-16 23:27）")]
for i, (a, b2) in enumerate(src_rows, 3): So.cell(i, 1, a); So.cell(i, 2, b2)

# ---------------- README 頁
Rm = SH["README"]
VB = lambda k, col="B": f"Val!{col}{VR[k]}"
Rm["A1"] = (f'="富士通 個股模型 v{VER}｜資料截止 2026-11-16（最新已公布期 2026 上半年，截至 "&TEXT({I("bs_date")},"yyyy-mm-dd")&"）｜成品日 {VDATE}｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&"｜build_id {BUILD_ID}｜本版未做先驗"')
Rm["A2"] = "命題、結論與重點觀察的數字由公式自本檔其他頁組字（不手寫）。"
sb = SR["B"]
Rm["A3"] = (f'="命題：AI 進入按人力計價的 IT 服務後，每單位營收＝AI 新專案增量 − 定價侵蝕 − 推論成本。本模型把原型 L 套用於 Service Solutions（人力 × 利用率 × 可用工時 × 時薪），Hardware 與 Ubiquitous 為獨立分部。情境 B 在 FY2032 的 Service AI 淨效果為 "&TEXT(Seg!K{sb["net"]}/1000,"#,##0")&" 十億日圓（新專案 "&TEXT(Seg!K{sb["incr"]}/1000,"#,##0")&" − 定價侵蝕 "&TEXT(Seg!K{sb["eros"]}/1000,"#,##0")&" − 推論成本 "&TEXT(Seg!K{sb["inf"]}/1000,"0.00")&"），AI 後所需 Service 交付人力 "&TEXT(Seg!K{sb["hc_ai"]},"0.0")&" 千人（AI 前 "&TEXT(Seg!K{sb["hc0"]},"0.0")&" 千人）。"')
Rm["A4"] = "命題的可證偽指標：Service 營收年增與員工數的差距、價值導向定價占比是否按目標上升、Service 毛利率是否回到 43% 目標；任何一項偏離基準區間即須重估。"
Rm["A5"] = (f'="投資建議（在自動化、價格保留與新專案三項落在情境 A 到 C 區間的基準假設下）："&{VB("rec")}&"；加權目標價 "&TEXT({VB("tp")},"0.00")&" 日圓，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"（現價 "&TEXT({VB("px")},"#,##0")&" 日圓，"&TEXT({VB("px_date")},"yyyy-mm-dd")&"）"')
Rm["A6"] = (f'="一、DCF 目標價 "&TEXT({VB("dcf_tp")},"0.00")&"、Comparable 目標價 "&TEXT({VB("comp_tp")},"0.00")&" 日圓；情境 B 落差 "&TEXT({VB("gap")},"+#,##0;-#,##0")&" 日圓，其中時點 "&TEXT({VB("gap_t")},"+#,##0;-#,##0")&"、倍數 "&TEXT({VB("gap_m")},"+#,##0;-#,##0")&"、股票薪酬 0、槓桿 0。"')
Rm["A7"] = (f'="二、情境 B FY2026 營收 "&TEXT(PL!E{STR["營收"]}/1000000,"0.000")&" 兆日圓（年增 "&TEXT(PL!E{STR["營收年增率"]},"0.0%")&"），對照公司指引 "&TEXT(Recon!C{RCR["rev26"]}/1000000,"0.00")&" 兆；調整後營業利益 "&TEXT(PL!E{b["adj"]}/1000,"#,##0")&" 十億，指引 "&TEXT(Recon!C{RCR["adj26"]}/1000,"#,##0")&" 十億（Recon）。"')
Rm["A8"] = (f'="三、每任務推論成本（Tokenomics Sol 層級）"&TEXT({I("cps")},"0.00")&" 美元，FY2032 推論成本僅占 Service 營收 "&TEXT(Seg!K{sb["inf"]}/Seg!K{sb["sv"]},"0.00%")&"；AI 對富士通的影響主要在定價侵蝕與新專案，不在推論成本。"')
Rm["A9"] = (f'="四、現價隱含 FY2032 IFRS 營業利益率 "&TEXT(RevDCF!B{RV["m_imp"]},"0.0%")&"（情境 B 為 "&TEXT(RevDCF!B{RV["m_fy32"]},"0.0%")&"）；現行 EV／LTM 營業利益 "&TEXT({VB("own_mult")},"0.0")&" 倍，為同業中位數 "&TEXT({VB("pe_med")},"0.0")&" 倍的 "&TEXT({VB("own_ratio")},"0.00")&" 倍；永續價值占 EV 約 "&TEXT(Val!D{VR["tvsh"]},"0%")&"。"')
Rm["A10"] = (f'="五、情境 A／B／C 加權目標價 "&TEXT(Val!C{VR["wb"]},"#,##0")&"／"&TEXT(Val!D{VR["wb"]},"#,##0")&"／"&TEXT(Val!E{VR["wb"]},"#,##0")&" 日圓；分化來自自動化比例、價格保留率（價值導向定價）與新專案占比（FY2032 Service AI 淨效果 A／B／C 為 "&TEXT(Seg!K{SR["A"]["net"]}/1000,"#,##0")&"／"&TEXT(Seg!K{SR["B"]["net"]}/1000,"#,##0")&"／"&TEXT(Seg!K{SR["C"]["net"]}/1000,"#,##0")&" 十億日圓）。"')
Rm["A12"] = "v1.3 展示層"
Rm["A13"] = (f'="富士通："&{VB("rec")}&"，加權目標價 "&TEXT({VB("tp")},"0.00")&" 日圓（上檔 "&TEXT({VB("up")},"+0.0%;-0.0%")&"）；本版未做先驗。"')
Rm["A14"] = "本版要點"
Rm["A15"] = "官方報告分部四列（Service Solutions、Hardware Solutions、Ubiquitous Solutions、Elimination/Corporate）；分部營收與分部利益（調整後營業利益）合計與合併值在 FY2024、FY2025 與 2025、2026 上半年差 0（決算短信 ¥百萬）。"
Rm["A16"] = "命題模組為原型 L，只套用於 Service Solutions；公司不揭露分部員工、利用率、計費工時、時薪，時薪由 FY2025 Service 營收反推（Derived），自動化比例、價格保留率、新專案占比為 Assumed，分化依據見 In 頁。"
Rm["A17"] = "推論成本只讀 Tokenomics v22.0 快照任務 5（類比）；公司未揭露 token 量與代理任務量，不以聊天基數乘倍數；token 價格路徑待 AI供應鏈 登錄。分部利益為公司自訂的調整後營業利益，合併營業利益取 IFRS。"
Rm["A18"] = "未做事項：先驗（Andy 指示本版跳過）；三層基準協調（Coord、C0、兩讀數）不適用：非 Siliconomics 覆蓋；Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照不適用：富士通非以自建算力為命題；v0.2 改讀 unit_economics。"
Rm["A19"] = (f'="合理性關卡（工作單 G3 第 3 條）："&IF({VB("g3_flag")}=1,"觸發，須診斷。","未觸發，無須診斷。")&"DCF 目標價÷現價 "&TEXT({VB("g3_dcf")},"0.00")&"（門檻 0.25 至 4）、EV 最小值 "&TEXT({VB("g3_ev")}/1000,"#,##0")&" 十億日圓、FY2026 營收對指引差 "&TEXT({VB("g3_rev")},"0.0%")&"、調整後營業利益率相對差 "&TEXT({VB("g3_om")},"0.0%")&"（門檻 10%，接近門檻，原因：分部利益率持平於 LTM，公司指引隱含下半年回升）、評等 "&{VB("valr")}&"。現價 2026-11-13 較 2026-10-23 低約 23%（決算後），股價為單一來源（traders.co.jp），待核對。"')
Rm["A20"] = "市場資料日期"
Rm["A21"] = (f'="現價 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤；無風險利率 "&TEXT({I("rf_date")},"yyyy-mm-dd")&"（與現價日相差 "&TEXT({VB("stale_rf")},"0")&" 個交易日，"&{VB("stale_rf","C")}&"）；同業倍數 2026-10-05 至 2026-11-13，逾期者已標（Val 頁）。"')
Rm["A22"] = "主要失效條件（可觀察）：2026 下半年（2027-04 公布）Service 毛利率未回到 43% 目標、價值導向定價占比未依 IR Day 路徑上升、Service 營收年增與員工數反向、日本以外 Service 營收持續下滑。"
Rm["A23"] = "風險登錄"
Rm["A24"] = "風險"; Rm["B24"] = "是否建模"; Rm["C24"] = "理由"; Rm["D24"] = "觀察指標"
risks = [("AI 取代交付工時（定價侵蝕）", "是", "a32 與 r 三情境", "Service 營收／員工、毛利率"), ("新專案不足以填補", "是", "n32 三情境", "Uvance 營收、訂單"), ("推論成本", "是", "Tokenomics 任務 5 成本，Sens 上界", "Service 毛利率"),
         ("Hardware 與 Ubiquitous 縮減", "部分", "ghw26、gub26 年增假設", "分部營收與利益率"), ("匯率與日圓利率", "否", "日圓口徑；rf 與 g 以 Sens 涵蓋", "日債殖利率")]
for i, rk in enumerate(risks, 25):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm.cell(31, 1, "工作表：README、In、Q、Seg、PL、DCF、Val、RevDCF、Recon、Sens、Sources")
Rm["A33"] = "頁"; Rm["B33"] = "輸入"; Rm["C33"] = "推導量"; Rm["D33"] = "因果方向"
mp = [("Seg", "hc_y25、svc_sh、util、hrs、hg、rg、a26、a32_A/B/C、r_A/B/C、n26、n32_A/B/C、tlen、cps、fx、px_dec、ghw26、ghw、gub26、gub", "Service AI 前營收、定價侵蝕、新專案增量、推論成本、AI 後人力、四列分部營收", "人力 × 利用率 × 工時 × 時薪 → R0；自動化與價格保留 → 侵蝕；任務 × 成本 → 推論"),
      ("PL", "cs、corp_dx、oneoff、gm、da_pct、capex_pct、wc_k、tax、oth_net、nci_pct", "分部調整後營業利益、IFRS 營業利益、FCFF、淨利、EPS", "營收 − 成本 → 分部利益 → IFRS 營業利益 → NOPAT → FCFF"),
      ("Val", "rf、beta_raw、erp、kd_sp、g_tv、ronic_sp、w_dcf、cm_adj、hw_disc、股數與淨現金", "WACC、DCF、分部加總 Comparable、加權目標價、評等", "FCFF 折現＋同業倍數 × TTM 營業利益 → 加權"),
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
          "BUILD_ID": BUILD_ID, "FP": fp, "HDR_R": hdr_r, "GATE": GATE, "RC_REV": RC_REV, "RC_END": RC_END, "CPS": CPS_USD}
if ROWMAP: json.dump(rowmap, open(ROWMAP, "w"), ensure_ascii=False, indent=1)
print("built", OUT, BUILD_ID)
