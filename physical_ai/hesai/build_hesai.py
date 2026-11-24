# 禾賽 Hesai（HSAI，Nasdaq ADS；港股 2525.HK）個股模型 v0.1 — AI實體應用 Project（線 1 Robotaxi、線 2 智慧車；LiDAR 讀線 6 共用零組件）
# 用法：python3 physical_ai/hesai/build_hesai.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json，並跑核對
#       python3 physical_ai/hesai/build_hesai.py --raw OUT  → 只寫一本（讀環境變數 HESAI_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度＝曆年（12 月決算，US GAAP，報表幣別人民幣）。金額：百萬人民幣（RMB M）；股數：百萬股（2026-07-10 起 1 股拆 8 股，歷史股數已 ×8 還原）；
# 1 ADS＝8 股 B 類普通股。EPS：人民幣／股；股價：美元／ADS；估值換算用 2026-11-20 的人民幣兌美元（收盤檔 CNY=X）。
# 結構：README、In、Hist、Q、PL（三情境：線需求 × 份額 → 出貨量 × ASP → 光達營收；策略成長業務與服務 → 標準損益 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數讀線 6 共用零組件成品（「下游取數」頁；逐年需求與 ASP 路徑在「計算」頁、降價幅度在「輸入」頁，對帳頁註明）與線 2 智慧車成品（只作交叉檢查）。
# 不讀 Siliconomics／Tokenomics／Wattonomics。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md（明確期 FY2027–FY2032，與線模型年限一致）。
import sys, os, json, subprocess, shutil, hashlib, tempfile, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_禾賽個股模型_{VER}.xlsx"
LINE_REL = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
SC_REL = "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx"
LINE = os.path.join(ROOT, LINE_REL); SCL = os.path.join(ROOT, SC_REL)
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
FC = [y for y in YRS if y >= 2027]                                                       # DCF 明確期 FY2027–FY2032

# ---------------- 出處 ----------------
Q326 = "公司新聞稿：2026 年第三季未經審計財務報告（2026-11-10；經 finviz 轉載）https://finviz.com/news/419744/hesai-group-reports-third-quarter-2026-unaudited-financial-results"
Q425 = "公司 6-K 附件 99.1：2025 年第四季及全年未經審計財務報告（2026-03-24）https://www.sec.gov/Archives/edgar/data/1861737/000110465926033591/tm269592d1_ex99-1.htm"
AR25 = "公司 2025 年報財務摘要（港交所公告，2026-04-24）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0424/2026042402278.pdf（經搜尋摘要讀取，未回原文逐頁核對）"
SAF = "stockanalysis.com 財務與統計頁（資料商；讀取 2026-11-24；各頁報價日見各列）"
SPL = "公司 6-K 與存託銀行公告：2026-07-10 起 1 股拆 8 股、1 ADS 由 1 股改為 8 股 B 類普通股（ADS 數不變）https://www.sec.gov/Archives/edgar/data/0001861737/000110465926082432/tm2620203d1_6k.htm"
YCL = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
UST = "美國 10 年期公債殖利率 5.00%（2026-10-23；與 physical_ai/tesla、pony 同一取值，跨模型一致）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，數字來源為公司財報或賣方）"
LN = f"線 6 成品 {LINE_REL}"
PEERX = "peers：2026-11-20 收盤取自收盤檔；資料商頁面的淨現金與營收未更新"

# ---------------- 讀線模型 ----------------
LW = openpyxl.load_workbook(LINE, data_only=True)
_DS = LW["下游取數"]
DSR = {_DS.cell(r, 1).value: r for r in range(2, _DS.max_row + 1) if _DS.cell(r, 1).value}
def lv(name, col="C"):
    r = DSR[name]; return _DS[f"{col}{r}"].value, f"下游取數!{col}{r}"
_CA = LW["計算"]
def calc_row(label_a, nth=1, start=1):
    k = 0
    for r in range(start, _CA.max_row + 1):
        if _CA.cell(r, 1).value == label_a:
            k += 1
            if k == nth: return r
    raise KeyError(label_a)
R_ASP = calc_row("LDR", 1)          # 單價路徑（美元/顆）2026–2032
R_DEM = 132                          # 基準需求（萬顆）2026–2032；下方以 2030 對照下游取數
assert _CA.cell(R_DEM, 1).value == "LDR"
LYC = {2026: "C", 2027: "D", 2028: "E", 2029: "F", 2030: "G", 2031: "H", 2032: "I"}
assert abs(_CA[f"G{R_DEM}"].value - lv("LiDAR 需求量 2030")[0]) < 1e-6
_IN = LW["輸入"]
R_IN = [r for r in range(1, _IN.max_row + 1) if _IN.cell(r, 1).value == "LDR" and _IN.cell(r, 3).value == "美元/顆"][0]
ASP_LINE = {"lo": _IN.cell(R_IN, 4).value, "b": _IN.cell(R_IN, 5).value, "hi": _IN.cell(R_IN, 6).value}
DEC_LINE = {"lo": _IN.cell(R_IN, 7).value, "b": _IN.cell(R_IN, 8).value, "hi": _IN.cell(R_IN, 9).value}   # 年降幅 6%／10%／15%
SW = openpyxl.load_workbook(SCL, data_only=True); _SD = SW["下游取數"]
SDR = {}
for r in range(2, _SD.max_row + 1):
    if _SD.cell(r, 1).value: SDR[(_SD.cell(r, 1).value, _SD.cell(r, 2).value)] = r
def sv(name, y, col="D"):
    r = SDR[(name, y)]; return _SD[f"{col}{r}"].value, f"下游取數!{col}{r}"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點（9M 實績＋第四季指引與假設）")
inp("q4rev", "FY2026 第四季營收（公司指引 11.00–11.50 億人民幣）", "RMB M", 1100, 1125, 1150, "Interested-party", Q326 + "：Q4 營收指引 1,100–1,150（US$164–172M）；基準取中點", "9M 實績 2,609.7；Q3 單季 1,068.3", CO)
inp("q4gm", "FY2026 第四季毛利率", "%", 0.38, 0.40, 0.42, "Assumed", "Q3 單季 41.0%（438.2÷1,068.3）；管理層預期全年約 40%（Business News Today 轉述財報電話會，2026-11）", "9M 毛利率 40.2%（1,049.5÷2,609.7；Derived）")
inp("q4ship", "FY2026 第四季光達出貨量", "百萬顆", 0.70, 0.80, 0.90, "Assumed", "Q3 73.0 萬顆（Verified，" + Q326 + "）、Q4'25 63.1 萬顆；Q3 低於公司 80–85 萬顆指引，Q4 未給出貨指引（未找到）", "全年約 263 萬顆，低於公司年初 300–350 萬顆目標（公司 Q3 未再重申）")
inp("ship9m", "FY2026 前三季光達出貨量", "百萬顆", None, 1.829671, None, "Verified", Q326 + "：ADAS 1,477,400＋機器人 352,271＝1,829,671 顆", "")
inp("sgi26", "FY2026 策略成長業務（機器人致動模組等）營收", "RMB M", 200, 250, 300, "Interested-party", Q326 + "：全年策略成長業務營收指引 200–300 重申；基準取中點", "Q3 單季交付致動模組 2.34 萬套；累計 3.53 萬套", CO)
inp("svc26", "FY2026 服務營收（含工程服務）", "RMB M", 40, 55, 70, "Assumed", "Q3 服務營收 16.8（Verified，" + Q326 + "）；9M 服務營收未取得", "FY2025 44.6、FY2024 110.8（工程服務一次性收入減少）")
inp("rd_q4", "FY2026 第四季研發費用", "RMB M", None, 265, None, "Assumed", "Q3 單季 254.0；Q4 略增", "9M 實績 689.9")
inp("sm_q4", "FY2026 第四季銷售費用", "RMB M", None, 55, None, "Assumed", "Q3 單季 53.0", "9M 實績 144.1")
inp("ga_q4", "FY2026 第四季管理費用", "RMB M", None, 95, None, "Assumed", "Q3 單季 93.6", "9M 實績 213.3")
inp("oi_q4", "FY2026 第四季其他營業收益（政府補助等，淨額）", "RMB M", None, 20, None, "Assumed", "Q3 25.0、9M 54.0（Verified）；趨勢下降", "")
inp("nonop_q4", "FY2026 第四季業外收支", "RMB M", None, 20, None, "Assumed", "Q3 稅前 80.8 − 營業利益 62.6＝18.2（Derived）", "9M：稅前 169.7 − 營業利益 56.2＝113.5（含 Q2 以前的利息收入與公允價值項目）")
inp("sh26", "FY2026 加權平均稀釋股數（拆股後）", "百萬股", None, 1304.2, None, "Verified", Q326 + "：9M 稀釋加權 1,304.2 百萬股（基本 1,252.5）", "")
inp("cx_p26", "FY2026 資本支出 ÷ 營收", "%", 0.08, 0.10, 0.12, "Assumed", "財報新聞稿未揭露資本支出；FY2025 約 11.3%（資料商，" + SAF + "）；2026 擴產能到 400 萬顆以上", "FY2024 約 13.1%、FY2023 約 22%（資料商）", DV)

sec("B. 命題模組一：線 6 光達需求 × 禾賽份額 → 出貨量（讀線 6「計算」頁逐年基準需求，2030 與「下游取數」頁一致）")
for y in range(2026, 2033):
    v = _CA[f"{LYC[y]}{R_DEM}"].value / 100
    inp(f"dmd{y % 100}", f"線 6 光達需求量 {y}（基準；五線歸因，不含割草機等）", "百萬顆", None, round(v, 6), None, "Assumed",
        f"{LN} 計算!{LYC[y]}{R_DEM}（線標 Assumed／Analogy；2030 同「下游取數」C 欄 1,749 萬顆）", "線低／高情境 2030：716／3,984 萬顆（下游取數!B／D）；本模型以需求乘數 dm 表達")
dm_lo = lv("LiDAR 需求量 2030", "B")[0] / lv("LiDAR 需求量 2030")[0]; dm_hi = lv("LiDAR 需求量 2030", "D")[0] / lv("LiDAR 需求量 2030")[0]
for s, v, lo, hi, note in (("A", 1.5, None, None, "樂觀：車用搭載率與機器人需求高於線基準（線高情境為基準的約 2.3 倍，取其中間偏高）"),
                           ("B", 1.0, round(dm_lo, 4), round(dm_hi, 4), "基準＝線基準；低／高＝線低／高情境 2030 ÷ 基準"),
                           ("C", 0.7, None, None, "保守：搭載率低於基準（線低情境約 0.41 倍，取其中間偏高）")):
    inp(f"dm_{s}", f"情境 {s}：線需求乘數（FY2030 達到；FY2026＝1，線性）", "x", lo, v, hi, "Assumed", note, "線 6 需求為承重變數（搭載率 8–25%）")
sh_ratio = "=(ship9m+q4ship)/dmd26"
inp("shr26", "校準：禾賽 FY2026 出貨量 ÷ 線需求（涵蓋率）", "x", None, sh_ratio, None, "Derived", "＝(前三季出貨＋Q4 假設) ÷ 線 6 FY2026 需求 5.055 百萬顆", "大於線口徑份額的部分＝線未含的割草機等機器人需求（線 6 說明頁自述低估）；非市占率")
for s, v, lo, hi, note in (("A", 0.50, None, None, "樂觀：禾賽守住約一半線需求（含機器人需求）"),
                           ("B", 0.40, 0.25, 0.50, "基準：FY2026 約 52% 降到 FY2030 40%（速騰、華為、圖達通等競爭；Assumed）"),
                           ("C", 0.30, None, None, "保守：份額流失、價格戰")):
    inp(f"shr_{s}", f"情境 {s}：涵蓋率 FY2030（FY2026 起線性；FY2030 起持平）", "x", lo, v, hi, "Assumed", note, "可證偽：季度出貨量、車廠定點（2025 前五大客戶占 55.8%）")

sec("C. 命題模組二：ASP 與毛利（ASP 降幅讀線 6「輸入」頁；FY2026 ASP 以公司數字校準）")
for s, v, lo, hi, note in (("A", DEC_LINE["lo"], None, None, "樂觀：降價慢（線低案 6%）"),
                           ("B", DEC_LINE["b"], DEC_LINE["lo"], DEC_LINE["hi"], f"基準＝線 6 年降價基準 10%（輸入!H{R_IN}）；低／高＝6%／15%"),
                           ("C", DEC_LINE["hi"], None, None, "保守：降價快（線高案 15%）")):
    inp(f"asd_{s}", f"情境 {s}：ASP 年降幅 FY2027 起", "%/年", lo, v, hi, "Assumed", note, f"公司實績：Q2'26 ASP 1,297 人民幣、年減 35%（線 6 追蹤指標，經長橋海豚研究；Interested-party）高於基準；本模型 FY2026 ASP 以 9M 出貨與營收校準（見 PL）")
inp("gm_sgi", "策略成長業務毛利率", "%", 0.15, 0.25, 0.35, "Assumed", "公司未揭露分業務毛利（未找到）", "硬體模組，低於光達")
inp("gm_sv", "服務營收毛利率", "%", 0.0, 0.20, 0.50, "Assumed", "未揭露（未找到）", "規模小，對結論影響很小")
for s, v, lo, hi, note in (("A", 0.43, None, None, "樂觀：規模與成本下降快於降價"),
                           ("B", 0.39, 0.33, 0.43, "基準：略低於 FY2026 全年約 40%（管理層稱明年毛利率相對穩定；降價壓力 Assumed）"),
                           ("C", 0.31, None, None, "保守：降價快於成本下降")):
    inp(f"gmt_{s}", f"情境 {s}：光達毛利率 FY2030（FY2026 校準值起線性；FY2030 起持平）", "%", lo, v, hi, "Assumed", note, "FY2024 42.6%、FY2025 41.8%（Verified，" + Q425 + "）")

sec("D. 策略成長業務、服務與費用")
for s, v, lo, hi in (("A", 1.00, None, None), ("B", 0.60, 0.20, 1.00), ("C", 0.25, None, None)):
    inp(f"gsgi_{s}", f"情境 {s}：策略成長業務營收年增 FY2027–FY2028", "%/年", lo, v, hi, "Assumed", "公司 Q3 稱該業務營收倍增（Business News Today）；致動模組單季 2.34 萬套；客戶集中（Sharpa 為關係人交易）", "未找到公司 2027 年指引")
for s, v, lo, hi in (("A", 0.40, None, None), ("B", 0.25, 0.10, 0.40), ("C", 0.10, None, None)):
    inp(f"gsgl_{s}", f"情境 {s}：策略成長業務營收年增 FY2029–FY2032", "%/年", lo, v, hi, "Assumed", "機器人量產速度不確定（線 4 人形、線 3 倉儲）", "")
inp("svc_g", "服務營收年增", "%/年", 0.0, 0.05, 0.10, "Assumed", "規模小", "")
for s, v, lo, hi in (("A", 0.12, None, None), ("B", 0.07, 0.03, 0.12), ("C", 0.03, None, None)):
    inp(f"grd_{s}", f"情境 {s}：研發費用年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "FY2025 −6.9%（796.9 vs 855.6）、9M'26 +18.7%（689.9 vs 581.4）（Derived）", "")
inp("sm_p", "銷售費用 ÷ 營收（FY2027 起）", "%", 0.04, 0.053, 0.07, "Assumed", "FY2025 6.3%（192.0÷3,027.6）、9M'26 5.5%（Derived）", "")
inp("gga", "管理費用年增 FY2027 起", "%/年", 0.03, 0.06, 0.10, "Assumed", "FY2025 −8.9%、9M'26 +11.6%（Derived）；FY2026 約 308", "")
inp("goi", "其他營業收益（政府補助等）年增 FY2027 起", "%/年", -0.25, -0.10, 0.0, "Assumed", "FY2024 276.1、FY2025 181.4、9M'26 54.0 持續下降（Verified，" + Q425 + "；" + Q326 + "）", "屬營業利益內項目，品質低於本業；基準續降")
inp("yld", "流動資金收益率（業外收支 ÷ 期初流動資金）", "%", 0.01, 0.02, 0.03, "Assumed", "FY2025 利息收入 130.2 ÷ 約 6,000–7,500 現金準備 ≈ 2%（Derived）", "9M'26 業外 113.5（含非利息項目）")
inp("tax_w", "有效稅率（FY2027 起，稅前為正時）", "%", 0.05, 0.12, 0.25, "Assumed", "FY2025 稅負 30.8 ÷ 稅前 466.8＝6.6%；9M'26 為 0；中國高新技術企業 15%，累積虧損抵稅逐步用完", "FY2026 假設 0（9M 實績 0）")
inp("dil_g", "股數年增（股權報酬稀釋）", "%/年", 0.0, 0.01, 0.03, "Assumed", "FY2025 加權稀釋 146.4（×8＝1,171.5）→ 9M'26 1,304.2，含 2025 年增資（Derived）", "未計新股融資")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.0, 0.15, 0.30, "Assumed", "硬體製造商，應收與存貨占比較高；資產負債表明細未取得", "")
inp("cx_p", "資本支出 ÷ 營收（FY2027 起）", "%", 0.06, 0.08, 0.12, "Assumed", "FY2025 約 11.3%、FY2024 約 13.1%（資料商）；產能 >400 萬顆後降低", "", DV)
inp("min_cash", "最低流動資金（資金缺口門檻）", "RMB M", None, 1500, None, "Assumed", "約一年研發＋管銷", "只用於資金缺口，不進目標價")

sec("E. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤，HSAI ADS）", "US$", None, 11.83, None, "Verified", YCL + "：HSAI 2026-11-20 收盤 11.83", "同檔 2525.HK 收盤 12.58 港元；×8 股÷7.8426＝12.83 美元（港股收盤早於美股，當日下跌未反映）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("px_hk", "對照：港股 2525.HK 收盤（只作對照，不進模型）", "HK$", None, 12.58, None, "Verified", YCL + "：2525.HK 2026-11-20", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("ads_r", "每 ADS 股數", "股", None, 8, None, "Verified", SPL, "")
inp("fx", "人民幣兌美元（2026-11-20；人民幣／美元）", "RMB/US$", None, 6.6665, None, "Verified", YCL + "：CNY=X 2026-11-20", "財報用 6.7045（2026-09-30）；估值一律用收盤檔當日值")
inp("sh_x", "稀釋股數額外調整（選擇權與 RSU 淨稀釋，在 Q3 稀釋加權之外）", "%", 0.0, 0.0, 0.03, "Assumed", "Q3 稀釋加權已含稀釋證券（1,305.2 vs 基本 1,255.3）", "")
inp("sh_v", "評價日股數（稀釋）", "百萬股", None, "=1305.2*(1+sh_x)", None, "Derived", Q326 + "：Q3 稀釋加權 1,305.2 百萬股", "資料商流通 ADS 156.44 百萬（×8＝1,251.5 基本股），差在稀釋")
inp("ads_n", "評價日 ADS 數", "百萬 ADS", None, "=sh_v/ads_r", None, "Derived", "＝股數 ÷ 8", "")
inp("liq_v", "現金準備（現金、受限現金、短期投資、長期定存；2026-09-30）", "RMB M", None, 6977.1, None, "Verified", Q326 + "：公司定義 cash reserve 6,977.1（不含股票投資）", "現金＋短期投資 6,993.2；2025-12-31 現金準備 7,511.0（含長期定存）")
inp("debt_v", "債務與租賃（由資料商淨現金反推）", "RMB M", 500, 978, 1300, "Derived", SAF + "：淨現金 US$897.2M（×6.7045＝6,015）→ 現金 6,993.2 − 6,015 ≈ 978", "資產負債表債務明細未取得；FY2025 利息費用 18.9 顯示確有借款", DV)
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債）", "%", None, 0.05, None, "Verified", UST, "2026-10-23 值，距股價日逾 2 個交易日：逾期（跨模型共用，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", 0.006, 0.009, 0.012, "Assumed", "Damodaran 中國 CRP 約 0.8–1.0%（本輪未回原文核對：不作 Verified）", "營收以中國為主；歐美車廠訂單另有")
inp("beta", "β（Blume 調整後）", "x", 1.0, "=0.67*1.47+0.33", 1.7, "Derived", SAF + "：HSAI 5 年月 β 1.47（讀取 2026-11-24）；Blume 0.67β＋0.33", "只有一個來源，區間 1.0–1.7（Assumed）；跨來源 β 未取得", DV)
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業 EV／營收）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺禾賽對同業的歷史相對倍數）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("F. 同業 EV／營收（2026-11-20 收盤換算；資料商頁面日期各異，市值依收盤檔收盤價重算、淨現金與營收不變）")
PEERS = [("ev_rs", "速騰聚創 RoboSense 2498.HK（車用與機器人光達；EV／TTM 營收）", "=(471.24*13.32-2300)/2510",
          "2026-11-24 頁面：股數 471.24 百萬、淨現金 23.0 億港元、TTM 營收 25.1 億港元、無 Forward PS；11-20 收盤 13.32 港元（頁面價 13.53）；口徑為 TTM"),
         ("ev_oust", "Ouster OUST（光達；EV／NTM 營收）", "=(73.0*23.799999-404.43)/(2040/7.28)",
          "2026-11-23 頁面：股數 73.0 百萬、淨現金 4.044 億美元、市值 20.4 億（價 27.90）、Forward PS 7.28 → NTM 營收約 2.80 億；11-20 收盤 23.80"),
         ("ev_mbly", "Mobileye MBLY（自駕感知方案；EV／NTM 營收）", "=(5660*4.57/6.66-1360)/(5660/2.86)",
          "2026-10-26 頁面（讀自 physical_ai/pony 模型同一輸入，本輪未重讀）：市值 56.6 億、淨現金 13.6 億、Forward PS 2.86、股價 6.66；11-20 收盤 4.57")]
for k, lab, f, note in PEERS:
    inp(k, f"EV／營收：{lab}", "x", None, f, None, "Derived", SAF + "：" + note, "NTM 營收＝市值 ÷ Forward PS（資料商共識，Interested-party）；會計口徑：US GAAP（速騰為 IFRS）；" + PEERX, DV)
inp("ev_med", "同業 EV／營收中位數（3 家；2 家 NTM、1 家 TTM 混合）", "x", None, "=MEDIAN(ev_rs,ev_oust,ev_mbly)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）；混合口徑，套用於 FY2027 營收（目標時點的 TTM）以免高估", "地平線 9660.HK 為晶片與軟體、TTM 倍數約 8.5 倍，口徑差太大，只作對照不進中位數")
inp("ev_hzn", "對照：地平線 9660.HK EV／TTM 營收（只作對照，不進模型）", "x", None, "=(64290*3.12/4.05-7590)/4930", None, "Derived", "資料來源同 physical_ai/pony 模型的同一輸入（2026-10-02 頁面市值 642.9 億港元、淨現金 75.9 億、TTM 營收 49.3 億；11-20 收盤 3.12）", "不進中位數", DV)
inp("rev_ttm", "禾賽 TTM 營收（FY2025 ＋ 9M'26 − 9M'25）", "RMB M", None, "=h_rev_25+h_rev_9m-2027.1", None, "Derived", Q326 + "：9M'25 營收 2,027.1", "")
inp("ev_self", "禾賽現行 EV／TTM 營收（只作對照，不進模型）", "x", None, "=(px*ads_n*fx-liq_v+debt_v)/rev_ttm", None, "Derived", "EV＝市值（美元×匯率）− 現金準備 ＋ 債務", "")


def build(path, ov=None, sens=None, rev=None, bid="（基準檔）"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 RMB M（人民幣百萬）、股數百萬股（拆股後）、股價美元／ADS"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy／Assumed（給區間）／Derived（寫明換算）；查不到寫「未找到」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    last_in = r
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("dm", "線需求乘數 FY2030"), ("shr", "涵蓋率 FY2030"), ("asd", "ASP 年降幅 FY2027 起"), ("gmt", "光達毛利率 FY2030"),
           ("gsgi", "策略成長業務年增 FY2027–28"), ("gsgl", "策略成長業務年增 FY2029–32"), ("grd", "研發費用年增")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 70

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度，US GAAP，RMB M；FY2023 只取得部分科目）。公式只引用本頁儲存格（定義名稱 h_*）。股數為拆股後（×8）還原"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "9M'26", "標記", "來源"])
    NA = "未取得"
    HR = [("h_prod", "產品營收", NA, 1966.3, 2982.9, NA, "Verified", Q425),
          ("h_svc", "服務營收", NA, 110.8, 44.6, NA, "Verified", Q425 + "（9M'26 服務營收未取得；Q3 單季 16.8）"),
          ("h_rev", "營收", 1876.989, 2077.157, 3027.573, 2609.7, "Verified", AR25 + "（FY2023–FY2025）；" + Q425 + "；" + Q326),
          ("h_cogs", "營業成本", "=1876.989-661.38", 1192.6, 1762.5, 1560.2, "Verified", "FY2024–25：" + Q425 + "；FY2023＝營收 − 毛利 661.38（" + SAF + "，Derived）；9M'26：" + Q326),
          ("h_rd", "研發費用", NA, 855.6, 796.9, 689.9, "Verified", Q425 + "；" + Q326),
          ("h_sm", "銷售費用", NA, 193.0, 192.0, 144.1, "Verified", "同上"),
          ("h_ga", "管理費用", NA, 316.9, 288.8, 213.3, "Verified", "同上"),
          ("h_oi", "其他營業收益（政府補助等，淨額）", NA, 276.1, 181.4, 54.0, "Verified", "同上"),
          ("h_op", "營業利益（虧損；含其他營業收益）", NA, -204.9, 168.8, 56.2, "Verified", "同上（資料商 FY2023–25 營業利益不含其他營業收益，口徑不同，不採用）"),
          ("h_pbt", "稅前利益（虧損）", NA, -101.2, 466.8, 169.7, "Verified", "同上（FY2025 含其他收益 184.6，主要為公允價值與一次性項目）"),
          ("h_tax", "所得稅", NA, 1.1, 30.8, 0.0, "Verified", "同上"),
          ("h_ni", "淨利（虧損）", -475.97, -102.4, 435.9, 169.7, "Verified", "FY2024–25 與 9M：同上；FY2023：" + SAF + "（資料商）"),
          ("h_sh", "加權平均股數（稀釋；FY2024 虧損＝基本；拆股後）", NA, "=129.188125*8", "=146.437135*8", 1304.2, "Verified", Q425 + "（原文 ×8 還原，" + SPL + "）；" + Q326),
          ("h_cx", "資本支出（購置固定資產與無形資產）", 414.7, 271.4, 342.2, NA, "Interested-party", SAF + "（FY2023 406.75＋7.93；FY2024 259.54＋11.82；FY2025 309.37＋32.83）；公司新聞稿未揭露", ),
          ("h_ocf", "營業現金流", 57.3, 63.5, 117.0, NA, "Interested-party", SAF + "；CFO 稱 2025 年營業現金流約 117（財報電話會，經搜尋摘要）")]
    for i, row in enumerate(HR, 3):
        k, lab, a, b_, c_, d_, tag, src = row[:8]
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b_, c_, d_)):
            x = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")): x.font = BLUE
        H.cell(i, 6, tag); H.cell(i, 7, src + ("；" + CO if False else ""))
        for y, L in (("23", "B"), ("24", "C"), ("25", "D"), ("9m", "E")): dn(f"{k}_{y}", f"Hist!${L}${i}")
    i = 3 + len(HR)
    H.cell(i, 1, "檢查：產品＋服務 − 營收（FY2024–FY2025；原文四捨五入容差 1）")
    for L in "CD": H.cell(i, "BCDE".index(L) + 2, f"=ROUND({L}3+{L}4-{L}5,0)")
    H.cell(i + 1, 1, "FY2023 除營收、毛利、淨利、資本支出、營業現金流外未取得（年報財務摘要只讀到營收；其餘科目未回原文）；9M'26 產品與服務拆分、9M'25 之外的單季科目未取得")
    H.cell(i + 2, 1, "FY2025 之前的股數：原文為拆股前股數（FY2025 加權稀釋 146.437 百萬），本頁 ×8 還原為拆股後；EPS 因此不等於原文 EPS（原文 2.98 元／舊股＝0.3725 元／新股）")
    H.column_dimensions["A"].width = 44; H.column_dimensions["G"].width = 80

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度與前三季對照（RMB M；出貨量為顆數）。全為公司新聞稿原文（Verified）；Q4'25 與 Q2'26 單季科目只取得出貨量與 ASP"
    hdr(Q, 2, ["項目", "Q3'25", "Q3'26", "9M'25", "9M'26", "標記", "來源", "來源誘因"])
    qrows = [("營收", 795.4, 1068.3, 2027.1, 2609.7), ("　產品營收", 790.1, 1051.5, None, None), ("　服務營收", 5.3, 16.8, None, None),
             ("營業成本", 460.5, 630.1, 1172.5, 1560.2), ("毛利", 334.9, 438.2, 854.6, 1049.5), ("研發費用", 198.9, 254.0, 581.4, 689.9),
             ("營業利益", 77.4, 62.6, 66.8, 56.2), ("淨利", 256.2, 80.8, 282.7, 169.7),
             ("光達出貨量（顆）", 441398, 729673, None, 1829671), ("　其中 ADAS（顆）", 380759, 638055, None, 1477400), ("　其中機器人（顆）", 60639, 91618, None, 352271)]
    for i, row in enumerate(qrows, 3):
        Q.cell(i, 1, row[0])
        for j, v in enumerate(row[1:]):
            if v is not None: Q.cell(i, 2 + j, v).font = BLUE
        Q.cell(i, 6, "Verified"); Q.cell(i, 7, Q326); Q.cell(i, 8, CO)
    q0 = 3 + len(qrows)
    Q.cell(q0, 1, "毛利率"); [Q.cell(q0, 2 + j, f"={L}7/{L}3") for j, L in enumerate("BCDE")]
    Q.cell(q0 + 1, 1, "營收年增（Q3；9M）"); Q.cell(q0 + 1, 3, "=C3/B3-1"); Q.cell(q0 + 1, 5, "=E3/D3-1")
    Q.cell(q0 + 2, 1, "Q3 光達 ASP 近似（RMB；(產品營收 − 策略成長業務 Q3 假設 90) ÷ 出貨量；Derived）"); Q.cell(q0 + 2, 3, "=(C4-90)/C11*1000000")
    Q.cell(q0 + 3, 1, "FY2026 校準 ASP（RMB；PL 頁情境 B FY2026）"); Q.cell(q0 + 3, 3, "=0")
    Q.cell(q0 + 4, 1, "對照：公司 Q2'26 綜合 ASP 1,297 人民幣、年減 35%（線 6 追蹤指標，經長橋海豚研究；Interested-party）；Q3 策略成長業務營收未取得，90 為假設")
    Q.column_dimensions["A"].width = 60
    QASP = q0 + 3

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境（RMB M）：線 6 需求 × 涵蓋率 → 出貨量 × ASP → 光達營收；＋策略成長業務、服務 → 營收 → 標準損益表（US GAAP）→ 現金流。FY2023–FY2025 實績、FY2026 為 9M 實績＋第四季指引與假設、FY2027 起模型"
    PR = {}; rr = 3
    KEYS = [("dem", "線 6 光達需求量（百萬顆；基準）"), ("dm", "情境需求乘數"), ("shr", "禾賽涵蓋率（出貨量 ÷ 線需求）"), ("ship", "禾賽光達出貨量（百萬顆）"),
            ("asp", "光達 ASP（RMB／顆）"), ("aspu", "光達 ASP（US$／顆；按 2026-11-20 匯率）"),
            ("lid", "光達營收（出貨量 × ASP）"), ("sgi", "策略成長業務營收（致動模組等）"), ("prod", "產品營收"), ("svc", "服務營收"), ("rev", "營收"),
            ("chk1", "檢查：光達＋策略成長業務 − 產品營收（應為 0）"), ("chk2", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"),
            ("cgl", "營業成本：光達"), ("cgs", "營業成本：策略成長業務"), ("cgv", "營業成本：服務"), ("cogs", "營業成本"),
            ("gp", "毛利"), ("gm", "毛利率"), ("lgm", "光達毛利率（模型推算；分業務毛利未揭露）"),
            ("rd", "研發費用"), ("sm", "銷售費用"), ("ga", "管理費用"), ("oi", "其他營業收益（政府補助等，淨額）"), ("opex", "營業費用合計（淨額）"),
            ("op", "營業利益"), ("opm", "營業利益率"), ("opf", "營業利益轉正年份旗標"),
            ("nonop", "業外收支（利息收入等）"), ("pbt", "稅前利益"), ("tax", "所得稅"), ("ni", "淨利"),
            ("nim", "淨利率"), ("nig", "淨利年增率"), ("sh", "加權平均稀釋股數（百萬股；拆股後）"), ("eps", "EPS（人民幣／股；稀釋）"), ("epsa", "EPS（人民幣／ADS；＝每股 × 8）"),
            ("nopat", "NOPAT（營業利益為正時 ×(1 − 稅率)）"), ("da", "折舊攤銷（資本支出五年平均攤提；未揭露，推算）"), ("cx", "資本支出"), ("dwc", "營運資金增加"),
            ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）"),
            ("liq", "期末現金準備（未計新融資）"), ("gap", "資金缺口（最低流動資金 − 期末現金準備，正值＝需融資）"), ("gapf", "資金缺口年份旗標")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(KEYS)}
        for k, l in KEYS: P.cell(R[k], 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k, c=c: f"{c}{R[k]}"; prev = lambda k, p=p: f"{p}{R[k]}"
            def w(k, f, c=c): P[f"{c}{R[k]}"] = f
            if y <= 2025:
                hk = str(y)[2:]
                w("rev", f"=h_rev_{hk}"); w("cogs", f"=h_cogs_{hk}"); w("ni", f"=h_ni_{hk}"); w("cx", f"=h_cx_{hk}")
                w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
                if y >= 2024:
                    for k, h in (("prod", "h_prod"), ("svc", "h_svc"), ("rd", "h_rd"), ("sm", "h_sm"), ("ga", "h_ga"), ("oi", "h_oi"), ("op", "h_op"), ("pbt", "h_pbt"), ("tax", "h_tax"), ("sh", "h_sh")):
                        w(k, f"={h}_{hk}")
                    w("opex", f"={cell('gp')}-{cell('op')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nonop", f"={cell('pbt')}-{cell('op')}")
                    w("chk2", f"=ROUND({cell('prod')}+{cell('svc')}-{cell('rev')},0)")
                    w("eps", f"={cell('ni')}/{cell('sh')}"); w("epsa", f"={cell('eps')}*ads_r"); w("gr", f"={cell('rev')}/{prev('rev')}-1")
                    w("nig", f"=IF({prev('ni')}<0,IF({cell('ni')}<0,\"虧損\",\"轉盈\"),{cell('ni')}/{prev('ni')}-1)")
                    w("opf", 9999 if y == 2024 else f"=IF({cell('op')}>0,{y},9999)")
                else:
                    for k in ("prod", "svc", "rd", "sm", "ga", "oi", "op", "opm", "opex", "nonop", "pbt", "tax", "sh", "eps", "epsa"): w(k, "未取得")
                    w("opf", 9999)
                continue
            n = y - 2026
            ys = str(y)[2:]
            if y == 2026:
                w("dem", "=dmd26"); w("dm", "=1"); w("shr", "=shr26"); w("ship", "=ship9m+q4ship")
                w("rev", "=h_rev_9m+q4rev"); w("svc", "=svc26"); w("sgi", "=sgi26"); w("lid", f"={cell('rev')}-{cell('svc')}-{cell('sgi')}")
                w("asp", f"={cell('lid')}/{cell('ship')}")
                w("cogs", "=h_cogs_9m+q4rev*(1-q4gm)")
                w("cgs", f"={cell('sgi')}*(1-gm_sgi)"); w("cgv", f"={cell('svc')}*(1-gm_sv)"); w("cgl", f"={cell('cogs')}-{cell('cgs')}-{cell('cgv')}")
                w("lgm", f"=1-{cell('cgl')}/{cell('lid')}")
                w("rd", "=h_rd_9m+rd_q4"); w("sm", "=h_sm_9m+sm_q4"); w("ga", "=h_ga_9m+ga_q4"); w("oi", "=h_oi_9m+oi_q4")
                w("nonop", "=h_pbt_9m-h_op_9m+nonop_q4"); w("tax", "=0"); w("sh", "=sh26")
                w("cx", f"=cx_p26*{cell('rev')}")
                w("liq", f"=liq_v+{cell('fcff')}*stub_frac")
            else:
                w("dem", f"=dmd{ys}"); w("dm", f"=1+(dm_{s}-1)*MIN(1,{n}/4)"); w("shr", f"=shr26+(shr_{s}-shr26)*MIN(1,{n}/4)")
                w("ship", f"={cell('dem')}*{cell('dm')}*{cell('shr')}")
                w("asp", f"={prev('asp')}*(1-asd_{s})"); w("lid", f"={cell('ship')}*{cell('asp')}")
                g = f"gsgi_{s}" if y <= 2028 else f"gsgl_{s}"
                w("sgi", f"={prev('sgi')}*(1+{g})"); w("svc", f"={prev('svc')}*(1+svc_g)"); w("rev", f"={cell('prod')}+{cell('svc')}")
                w("lgm", f"=$F{R['lgm']}+(gmt_{s}-$F{R['lgm']})*MIN(1,{n}/4)")
                w("cgl", f"={cell('lid')}*(1-{cell('lgm')})"); w("cgs", f"={cell('sgi')}*(1-gm_sgi)"); w("cgv", f"={cell('svc')}*(1-gm_sv)")
                w("cogs", f"={cell('cgl')}+{cell('cgs')}+{cell('cgv')}")
                w("rd", f"={prev('rd')}*(1+grd_{s})"); w("sm", f"=sm_p*{cell('rev')}"); w("ga", f"={prev('ga')}*(1+gga)"); w("oi", f"={prev('oi')}*(1+goi)")
                w("nonop", f"=yld*MAX(0,{prev('liq')})"); w("tax", f"=IF({cell('pbt')}>0,{cell('pbt')}*tax_w,0)"); w("sh", f"=sh26*(1+dil_g)^{n}")
                w("cx", f"=cx_p*{cell('rev')}")
                w("liq", f"={prev('liq')}+{cell('fcff')}+{cell('nonop')}*(1-tax_w)-{cell('div')}")
            w("aspu", f"={cell('asp')}/fx")
            w("prod", f"={cell('lid')}+{cell('sgi')}") if y > 2026 or True else None
            if y == 2026: w("rev", "=h_rev_9m+q4rev")
            w("chk1", f"=ROUND({cell('lid')}+{cell('sgi')}-{cell('prod')},6)"); w("chk2", f"=ROUND({cell('prod')}+{cell('svc')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opex", f"={cell('rd')}+{cell('sm')}+{cell('ga')}-{cell('oi')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("opf", f"=IF({cell('op')}>0,{y},9999)")
            w("pbt", f"={cell('op')}+{cell('nonop')}"); w("ni", f"={cell('pbt')}-{cell('tax')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("nig", f"=IF({prev('ni')}<0,IF({cell('ni')}<0,\"虧損\",\"轉盈\"),{cell('ni')}/{prev('ni')}-1)")
            w("eps", f"={cell('ni')}/{cell('sh')}"); w("epsa", f"={cell('eps')}*ads_r")
            w("nopat", f"=IF({cell('op')}>0,{cell('op')}*(1-tax_w),{cell('op')})")
            ys5 = [x for x in range(max(2023, y - 4), y + 1)]
            w("da", "=SUM(" + ",".join(f"{COL[x]}{R['cx']}" for x in ys5) + f")/{len(ys5)}")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('nopat')}+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", "=0"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
            w("gap", f"=MAX(0,min_cash-{cell('liq')})"); w("gapf", f"=IF({cell('gap')}>0,{y},9999)")
        P.cell(R["dm"], 2, "1＋(情境乘數 − 1) × 線性爬升到 FY2030")
        P.cell(R["shr"], 2, "FY2026 校準值 → FY2030 情境值（線性），之後持平")
        P.cell(R["asp"], 2, "FY2026＝(營收 − 服務 − 策略成長業務) ÷ 出貨量；其後 × (1 − 年降幅)")
        P.cell(R["lgm"], 2, "FY2026＝殘差；其後線性走向情境 FY2030 值")
        P.cell(R["cogs"], 2, "FY2026＝9M 實績＋Q4 營收 ×(1 − Q4 毛利率)")
        P.cell(R["op"], 2, "含其他營業收益（政府補助），公司報表口徑")
        P.cell(R["fcff"], 2, "股權報酬不加回；虧損年不計稅盾")
        P.cell(R["liq"], 2, "FY2026＝9 月底 ＋ Q4 FCFF（stub）；其後＋FCFF＋稅後業外 − 股利")
        for k in ("gr", "gm", "lgm", "opm", "nim", "nig", "shr"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(KEYS) + 2
    Q.cell(QASP, 3, f"=PL!F{PR['B']['asp']}")
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 36

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加；人民幣百萬），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032（與線 6 年限一致），期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；每 ADS 價值以 2026-11-20 匯率換算美元（匯率持平假設）"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（RMB M）"; D["B4"] = "=px*ads_n*fx"
    D["A5"] = "負債 D（RMB M）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=(rf+0.02)*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-09-30）＝現金準備 − 債務與租賃（RMB M）"; D["B10"] = "=liq_v-debt_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*ads_n*fx)*(rf+beta*(erp+crp))+debt_v*(rf+0.02)*(1-tax_w))/((px*ads_n*fx)+debt_v)+wacc_adj-wacc,9)"
    for r_ in range(3, 11): D[f"B{r_}"].number_format = "0.00%" if r_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13; NC = len(FC); DCOL = "BCDEFGH"
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in FC]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(FC):
            c = DCOL[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"=PL!{pc}{R['nopat']}"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11; L = DCOL[NC - 1]
        items = [("明確期現值合計", f"=SUM(B{b+9}:{L}{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"={L}{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^{NC}"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值（RMB M）", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每 ADS 價值（評價日；US$）", f"=MAX(0,B{t+7})/ads_n/fx"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t；US$／ADS）", f"=B{t+8}*(1+ke)^t_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+10}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF({L}{b+1}>0,B{t+2}/{L}{b+1},\"N/M\")"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每 ADS 價值（推到 12 個月後）", f"=MAX(0,B{t}+{L}{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^{NC}+B{t+5}+B{t+6})/ads_n/fx*(1+ke)^t_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 9, ps=t + 8, ev=t + 4, tvs=t + 10, alt=t + 13, eq=t + 7)
        dn(f"dcf_{s}", f"DCF!$B${t+9}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價（美元／ADS）：DCF（權重 w_dcf）＋ Comparable（同業 EV／營收中位數 × FY2027 營收（目標時點的 TTM）＋ 目標時點淨現金，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}"); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁；US$／ADS")
    def cf(s, k, y): return f"PL!{COL[y]}{PR[s][k]}"
    r_nct = vrow("目標時點淨現金（評價日淨現金 ＋ Stub FCFF ＋ FY2027 FCFF 與稅後業外；RMB M）", [f"=dcf_nc+{cf(s,'fcff',2026)}*stub_frac+{cf(s,'fcff',2027)}+{cf(s,'nonop',2027)}*(1-tax_w)" for s in SC], "RMB M")
    r_cu = vrow("Comparable（未調整）：EV／營收中位數 × FY2027 營收 ＋ 目標時點淨現金（每 ADS；US$）", [f"=(ev_med*{cf(s,'rev',2027)}+{L}{r_nct})/ads_n/fx" for s, L in zip(SC, "BCD")],
                "目標時點（2027-11）的過去十二個月營收≈FY2027；同一倍數套用整體營收（同業不足以分業務）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（美元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（RMB M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"0.00")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_gapmax = vr; V.cell(vr, 1, "資金缺口：FY2026–FY2032 最大缺口（RMB M；情境 A／B／C；未計新融資）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=MAX(PL!F{PR[s]['gap']}:PL!L{PR[s]['gap']})")
    vr += 1
    r_gapyr = vr; V.cell(vr, 1, "資金缺口：現金準備首次低於最低流動資金的年份（情境 A／B／C）")
    for j, s in enumerate(SC):
        g = PR[s]['gapf']
        V.cell(vr, 2 + j, f"=IF(MIN(PL!F{g}:PL!L{g})=9999,\"FY2032 前無缺口\",\"FY\"&MIN(PL!F{g}:PL!L{g}))")
    vr += 1
    r_liqmin = vr; V.cell(vr, 1, "資金缺口：FY2026–FY2032 最低期末現金準備（RMB M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=MIN(PL!F{PR[s]['liq']}:PL!L{PR[s]['liq']})")
    vr += 1
    r_opy = vr; V.cell(vr, 1, "營業利益轉正年份（FY2026 起首次；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=IF(MIN(PL!F{PR[s]['opf']}:PL!L{PR[s]['opf']})=9999,\"FY2032 前未轉正\",MIN(PL!F{PR[s]['opf']}:PL!L{PR[s]['opf']}))")
    vr += 1
    r_pen30 = vr; V.cell(vr, 1, "光達出貨量 FY2030（百萬顆；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!J{PR[s]['ship']}")
    vr += 1
    r_pen35 = vr; V.cell(vr, 1, "光達出貨量 FY2032（百萬顆；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!L{PR[s]['ship']}")
    vr += 1
    r_gpm30 = vr; V.cell(vr, 1, "營業利益率 FY2030（情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!J{PR[s]['opm']}")
    vr += 2
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_cash = one("現價隱含：每 ADS 淨現金（評價日淨現金 ÷ ADS 數 ÷ 匯率；US$）", "=dcf_nc/ads_n/fx", "現價與每 ADS 淨現金的差＝市場給營運業務的價值")
    r_mkt = one("現價隱含：營運業務 EV（市值 − 淨現金，RMB M）", "=px*ads_n*fx-dcf_nc")
    one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_evs = one("隱含 EV／營收：加權目標價隱含 EV ÷ 情境 B FY2027 營收", f"=(B{r_tp}*ads_n*fx-dcf_nc)/PL!G{PR['B']['rev']}", "分母：模型情境 B FY2027 營收")
    r_evs2 = one("現行 EV／TTM 營收（禾賽）÷ 同業中位數", "=ev_self/ev_med", "只作對照")
    for r_ in (r_up, r_a1, r_a2, r_a3): V[f"B{r_}"].number_format = "0.0%"
    for L in "BCD": V[f"{L}{r_gpm30}"].number_format = "0.0%"
    for r_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{r_}"].number_format = "0.00"
    V.column_dimensions["A"].width = 76; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, nct=r_nct, pd2=r_pd2, gj=r_gj, gapmax=r_gapmax, gapyr=r_gapyr, liqmin=r_liqmin, opy=r_opy,
              pen30=r_pen30, pen35=r_pen35, gpm30=r_gpm30, cash=r_cash, mkt=r_mkt, wk=r_wk, evs=r_evs, evs2=r_evs2, vg=r_vg)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、線 6 對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司、共識或線", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收（RMB M）：模型 vs 9M 實績＋Q4 指引中點", f"=PL!F{B_['rev']}", "=2609.7+1125", "Q4 指引 1,100–1,150（Interested-party）"),
          ("FY2026 光達出貨量（百萬顆）：模型 vs 公司年初指引下緣 3.0", f"=PL!F{B_['ship']}", 3.0, "公司 2026-03 指引 300–350 萬顆；Q3 財報未再重申（Interested-party）"),
          ("FY2026 光達 ASP（US$）：模型校準 vs 線 6 基準 180", f"=PL!F{B_['aspu']}", f"={ASP_LINE['b']}", "線 6 基準 180 美元（輸入!E112，以 7.1 匯率由 1,297 人民幣換算）；本模型 ASP 為 9M 實績校準，含不同產品組合與 6.67 匯率"),
          ("FY2030 光達 ASP（US$）：模型 vs 線 6 單價路徑", f"=PL!J{B_['aspu']}", f"=計算!G{R_ASP}", "線 6 路徑＝180 × 0.9^4；差來自 FY2026 起點與匯率"),
          ("FY2030 出貨量（百萬顆）：模型 vs 線 6 基準需求", f"=PL!J{B_['ship']}", "=dmd30", "比值＝涵蓋率 × 需求乘數；線口徑不含割草機等"),
          ("FY2030 光達營收（US$M）：模型 vs 線 6 LiDAR 市場規模 2030（基準）", f"=PL!J{B_['lid']}/fx", lv("LiDAR 市場規模 2030")[0], "線市場規模 2,141 百萬美元為五線歸因；禾賽營收含線外需求（割草機等），故比值可高於涵蓋率")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    RC["C6"] = LW["計算"][f"G{R_ASP}"].value   # 線 6 單價路徑 2030（建置時讀值）
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_hesai.py 格點＋線性內插後驗證；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_hesai.py 以 LibreOffice 重算每一案例）；結果為常數快照。估值評等欄只看上檔空間"
    SE["A2"] = "說明：min_cash 只用於資金缺口，無區間；rel_k 基準即 1.00；dmd26–dmd32 為線 6 讀值，以需求乘數 dm 表達不確定性"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "In 頁輸入指紋（敏感度重算時）"; SE["D4"] = sens["fpsum"] if sens else 0
    SE["A5"] = "In 頁輸入指紋（本檔即時）"; SE["D5"] = f"=SUM(In!E4:E{last_in - 1})"
    SE["A6"] = "基準目標價（敏感度重算時）"; SE["D6"] = sens["base"]["tp"] if sens else 0
    SE["A7"] = "模型狀態"; SE["D7"] = f'=IF(AND(ABS(D5-D4)<=1E-6*MAX(1,ABS(D4)),ABS(tp-D6)<1E-6),"CURRENT","STALE")'
    hdr(SE, 9, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等", "讀法"])
    sr = 10
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]).number_format = "0.00"; SE.cell(sr, 5, f"=D{sr}/tp-1").number_format = "+0.0%;-0.0%"
            SE.cell(sr, 6, cs["dcf"]).number_format = "0.00"; SE.cell(sr, 7, f"=F{sr}/Val!B{r_dcft}-1").number_format = "+0.0%;-0.0%"
            SE.cell(sr, 8, f'=IF(D{sr}/px-1>0.5,"強力買進",IF(D{sr}/px-1>0.25,"買進",IF(D{sr}/px-1>=-0.25,"中立",IF(D{sr}/px-1>=-0.5,"放空","強力放空"))))')
            SE.cell(sr, 9, cs.get("note", "")); sr += 1
    SE.cell(sr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_hesai.py")
    for col, wd in zip("ABCDEFGHI", [50, 40, 24, 12, 10, 12, 10, 10, 50]): SE.column_dimensions[col].width = wd

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線變數（線 6：{LINE_REL}；線 2：{SC_REL}；建置時讀值，不手打）與不一致處"
    hdr(RE, 2, ["變數", "線 ID 或儲存格", "線基準值", "線低／高", "本模型用法", "一致性", "標記", "影響本模型哪些儲存格"])
    rec = []
    v, ref = lv("LiDAR 需求量 2030"); lo, _ = lv("LiDAR 需求量 2030", "B"); hi, _ = lv("LiDAR 需求量 2030", "D")
    rec.append(("LiDAR 需求量 2030（萬顆）", f"LiDAR 需求量 2030（{ref}）", v, f"{lo}／{hi}", "FY2030 基準需求；低／高轉為需求乘數 dm 的敏感度區間", "一致（直接讀線）", "Assumed", "In dmd30、dm_B 區間"))
    v, ref = lv("LiDAR 市場規模 2030"); lo, _ = lv("LiDAR 市場規模 2030", "B"); hi, _ = lv("LiDAR 市場規模 2030", "D")
    rec.append(("LiDAR 市場規模 2030（百萬美元）", f"LiDAR 市場規模 2030（{ref}）", round(v, 1), f"{lo:.0f}／{hi:.0f}", "只作 Recon 對照（禾賽光達營收是否超過線市場）", "口徑不同：線為五線歸因，禾賽含割草機等線外需求", "Assumed", "Recon 頁"))
    for y in range(2026, 2033):
        rec.append((f"線 6 逐年基準需求 {y}（萬顆）", f"計算!{LYC[y]}{R_DEM}（不在下游取數頁；2030 與下游取數一致）", _CA[f"{LYC[y]}{R_DEM}"].value, "—", "出貨量＝需求 × 需求乘數 × 涵蓋率", "一致；但逐年值不在下游取數頁，下次線改版時請列入", "Assumed", f"In dmd{y % 100} → PL 線需求"))
    rec.append(("線 6 光達年降價幅度", f"輸入!G{R_IN}:I{R_IN}", DEC_LINE["b"], f"{DEC_LINE['lo']}／{DEC_LINE['hi']}", "情境 A／B／C 的 ASP 年降幅", "一致；但公司 Q2'26 ASP 年減 35%，顯著快於線基準（見 Recon）", "Assumed", "In asd_A／B／C"))
    rec.append(("線 6 光達單價 2026（美元）", f"輸入!E{R_IN}", ASP_LINE["b"], f"{ASP_LINE['lo']}／{ASP_LINE['hi']}", "只作對照；本模型 ASP 用 9M 實績校準", "不一致：禾賽 FY2026 校準 ASP 約 196 美元（含 FY2026 產品組合與 6.67 匯率）高於線 180；差異理由見 Recon", "Verified／Interested-party", "Recon 頁"))
    v, ref = sv("中國城市 NOA 新車數", 2026); lo, _ = sv("中國城市 NOA 新車數", 2026, "C"); hi, _ = sv("中國城市 NOA 新車數", 2026, "E")
    rec.append(("線 2 中國城市 NOA 新車數 2026（萬台）", f"中國城市 NOA 新車數 2026（{ref}）", v, f"{lo}／{hi}", "交叉檢查：禾賽 9M'26 ADAS 光達 1.48 百萬顆，全年約 2.1 百萬顆，約等於每台 NOA 新車 0.5 顆", "未用於計算；線 6 需求已含智慧車拉動", "Assumed", "對照"))
    v, ref = sv("中國城市 NOA 新車數", 2030); lo, _ = sv("中國城市 NOA 新車數", 2030, "C"); hi, _ = sv("中國城市 NOA 新車數", 2030, "E")
    rec.append(("線 2 中國城市 NOA 新車數 2030（萬台）", f"中國城市 NOA 新車數 2030（{ref}）", v, f"{lo}／{hi}", "交叉檢查：模型 FY2030 情境 B 出貨 7.0 百萬顆（含機器人），若 ADAS 占 75% 約 5.2 百萬顆＝每台 NOA 新車約 0.5 顆", "未用於計算", "Assumed", "對照"))
    for i_, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i_, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["B"].width = 44; RE.column_dimensions["E"].width = 50; RE.column_dimensions["F"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("車用光達降價速度（ASP Q2'26 年減 35%）", "是（asd_B 低／高、asd_C）", "情境 C 與 Sens", "季度 ASP、出貨量與營收比"),
             ("市占與價格競爭（速騰聚創、華為、圖達通、Ouster）", "是（shr_C）", "情境 C 與 Sens", "季度出貨量、車廠定點"),
             ("客戶集中（2025 年前五大客戶 55.8%）", "否", "事件型；profile 檔數字未回原文", "20-F 客戶集中度、定點公告"),
             ("機器人需求（割草機客戶出狀況，Q3 JT 系列出貨低於預期）", "部分（gsgi、shr）", "割草機需求不在線 6 口徑內", "機器人出貨、致動模組出貨"),
             ("策略成長業務客戶為關係人（Sharpa）", "否", "營收品質", "關聯交易揭露"),
             ("美中監管（美國國防部 1260H 清單、關稅、ADR 風險）", "否", "事件型", "美國商務部與國防部名單"),
             ("全年指引收縮（Q3 未重申 300–350 萬顆）", "是（q4ship、shr）", "FY2026 出貨量 263 萬顆", "Q4 財報與 2027 指引"),
             ("人民幣匯率（成本人民幣、部分營收美元）", "否", "估值以 2026-11-20 匯率持平換算", "匯率")]
    for i_, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i_, 1 + j, v)
    i_ = 3 + len(risks) + 1
    RK.cell(i_, 1, "未找到／未取得清單（「未找到」＝查過沒有公開資料；「未取得」＝有資料但本輪未取得原文）").font = BOLD; i_ += 1
    hdr(RK, i_, ["項目", "狀態", "替代做法", "下次怎麼補"]); i_ += 1
    miss = [("分業務（光達、策略成長）毛利與營收", "未找到（公司只揭露產品與服務、出貨量）", "以出貨量 × ASP 推導光達營收；策略成長業務用公司 200–300 指引", "20-F 分部附註、法說簡報"),
            ("FY2023 完整損益（費用、營業利益、稅前、股數）", "未取得", "PL 頁 FY2023 只列營收、毛利、淨利、資本支出", "FY2024 年報比較欄"),
            ("資本支出、折舊攤銷、營業現金流（公司新聞稿未載）", "未取得原文", "資料商值（FY2023–25）；折舊以資本支出五年平均推算", "20-F 現金流量表"),
            ("債務與租賃、股票投資餘額", "未取得原文", "債務由資料商淨現金反推（978，區間 500–1,300）；股票投資不計入", "20-F 資產負債表附註"),
            ("2027 年以後指引、賣方共識與目標價", "未取得", "全為 Assumed 三情境；不放賣方目標價對照", "法說會、資料商"),
            ("選擇權與 RSU 明細", "未取得", "用 Q3 稀釋加權股數 ＋ 0–3% 區間", "20-F 股權報酬附註"),
            ("Q3 策略成長業務單季營收、9M 服務營收", "未取得", "Q3 以 90 假設，僅用於 Q 頁 ASP 近似", "財報電話會逐字稿"),
            ("同業倍數的當日完整資料", "部分未取得", "Ouster 與速騰以收盤檔收盤價重算市值；淨現金與營收取頁面值；Mobileye 沿用 pony 模型 2026-10-26 頁面", "資料商"),
            ("速騰聚創 NTM 營收", "未取得（無 Forward PS）", "用 TTM 營收，混合口徑", "資料商"),
            ("中國國家風險溢酬原文", "未取得", "Assumed 0.9%（0.6–1.2%）", "Damodaran 國家風險表")]
    for row in miss:
        for j, v in enumerate(row): RK.cell(i_, 1 + j, v)
        i_ += 1
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i_, s_ in enumerate([Q326, Q425, AR25, SPL, SAF, YCL, UST, LN + "（下游取數、計算、輸入）", f"線 2 成品 {SC_REL}（下游取數；交叉檢查）",
                             "Business News Today：Hesai targets robotics growth as lidar shipments surge faster than revenue（2026-11-11；轉述財報電話會：全年毛利率約 40%、策略成長業務倍增）https://business-news-today.com/hesai-targets-robotics-growth-as-lidar-shipments-surge-faster-than-revenue/",
                             "Earnings Whispers 經 finviz：Hesai's physical AI bet outruns plan as core lidar misses its own guide（2026-11-10）https://finviz.com/news/420130/hesais-physical-ai-bet-outruns-plan-as-core-lidar-misses-its-own-guide-and-drops-its-shipment-target"], 1):
        SO.cell(i_, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"禾賽 Hesai（HSAI；港股 2525.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（Q3 財報 2026-11-10）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 美元／ADS）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：禾賽的價值取決於光達出貨量（線 6 光達需求 × 禾賽涵蓋率）與單顆 ASP 的降幅，能否讓營收成長快過降價、毛利率守在約 40%；機器人（策略成長業務）是第二成長曲線，但目前占營收很小"
    RD["A4"] = (f'="在情境 B 線 6 需求乘數 "&TEXT(dm_B,"0.0")&"、FY2030 涵蓋率 "&TEXT(shr_B,"0%")&"、ASP 年降 "&TEXT(asd_B,"0%")&"、光達毛利率 FY2030 "&TEXT(gmt_B,"0%")&" 的假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬人民幣）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = (f'="現金（情境 B，未計新融資）：FY2026–FY2032 最低期末現金準備 "&TEXT(Val!C{r_liqmin},"#,##0")&" 百萬人民幣、資金缺口 "&Val!C{r_gapyr}&"；營業利益轉正 "&Val!C{r_opy}&"；光達出貨量 FY2030 "&TEXT(Val!C{r_pen30},"0.0")&" 百萬顆、FY2032 "&TEXT(Val!C{r_pen35},"0.0")&" 百萬顆"')
    RD["A7"] = "失效條件（可觀察）：季度出貨量連兩季低於情境 B 路徑；ASP 年減幅連續超過 15%；毛利率跌破 35%；策略成長業務營收全年低於 200 百萬人民幣"
    RD["A8"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics；LiDAR 需求與降價幅度讀線 6，線 2 只作交叉檢查。分業務毛利、FY2023 完整損益、債務明細、2027 年以後指引與共識未取得，詳見 Risk 頁"
    wb.save(path)
    return VR, PR, DR

# ---------------- 重算與讀值 ----------------
LO = "file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), "lo_profile_hesai")   # LibreOffice 隔離設定檔
def recalc_many(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 20):
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + paths[i:i + 20],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=170)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile_hesai"], stderr=subprocess.DEVNULL)
                if attempt == 1: raise
    return [os.path.join(outdir, os.path.basename(p)) for p in paths]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_cases(ovs):
    """ovs：[(tag, override dict)] → {tag: vals}；一次批次重算"""
    d = tempfile.mkdtemp(prefix="hesai_"); paths = []; meta = {}
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p); meta[p] = (tag, VR, DR)
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = {}
    for p, o in zip(paths, outs):
        tag, VR, DR = meta[p]; res[tag] = read_vals(o, VR, DR if tag == "base" else None)
    shutil.rmtree(d, ignore_errors=True)
    return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k: return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("q4rev", "FY2026 第四季營收（指引區間）"), ("q4gm", "FY2026 第四季毛利率"), ("q4ship", "FY2026 第四季出貨量"), ("sgi26", "FY2026 策略成長業務營收"),
         ("svc26", "FY2026 服務營收"), ("cx_p26", "FY2026 資本支出 ÷ 營收"), ("dm_B", "線 6 需求乘數 FY2030（情境 B）"), ("shr_B", "涵蓋率 FY2030（情境 B）"),
         ("asd_B", "ASP 年降幅（情境 B）"), ("gmt_B", "光達毛利率 FY2030（情境 B）"), ("gm_sgi", "策略成長業務毛利率"), ("gm_sv", "服務毛利率"),
         ("gsgi_B", "策略成長業務年增 FY2027–28（情境 B）"), ("gsgl_B", "策略成長業務年增 FY2029–32（情境 B）"), ("svc_g", "服務營收年增"),
         ("grd_B", "研發費用年增（情境 B）"), ("sm_p", "銷售費用 ÷ 營收"), ("gga", "管理費用年增"), ("goi", "其他營業收益年增"), ("yld", "現金收益率"),
         ("tax_w", "有效稅率"), ("dil_g", "股數年增"), ("fc_wc", "營運資金 ÷ 營收增量"), ("cx_p", "資本支出 ÷ 營收"), ("sh_x", "稀釋股數額外調整"),
         ("debt_v", "債務與租賃"), ("crp", "中國國家風險溢酬"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"),
         ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重")]

def readme(path, VR, PR, DR, bid, sens, revs, checks):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, 400) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    def rvf(r):
        return r["val"] if isinstance(r["val"], str) else f"{r['val']:.4g}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {rvf(r)} | {r['base']} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=YRS):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else ("—" if v is None else str(v)))
        return " | ".join(out)
    YS = list(range(2024, 2033))
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {f0(g(chr(66+i)+str(VR['gate'])))} | {f0(g(chr(66+i)+str(VR['liqmin'])))} | {g(chr(66+i)+str(VR['gapyr']))} | {g(chr(66+i)+str(VR['opy']))} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    chk = "\n".join(f"- `{c[0]}`：{c[1]}" for c in checks)
    txt = f"""# physical_ai/hesai — 禾賽 Hesai（HSAI；港股 2525.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

Nasdaq ADS（2026-07-10 起 1 ADS＝8 股 B 類普通股；美元）；港交所 2525.HK。線 1 Robotaxi、線 2 智慧車（成品 `{SC_REL}`，只作交叉檢查）；LiDAR 需求與降價幅度讀線 6（`{LINE_REL}`）。報表幣別人民幣（百萬），本 README 的數字由 `build_hesai.py` 從重算後的成品讀出產生，不手寫。

**範圍限制（高不確定性原則）：** 公司只揭露產品與服務兩類營收（外加出貨量），分業務毛利與營收未揭露。本模型做到：光達營收由「線 6 需求 × 涵蓋率 → 出貨量 × ASP」推導、策略成長業務與服務以成長路徑建模、標準損益表、FCFF DCF 與同業 EV／營收。**沒有做到**：分業務毛利（以 FY2026 合併毛利殘差校準）、FY2023 完整損益（只有營收、毛利、淨利、資本支出）、債務明細（由資料商淨現金反推）、股票投資、選擇權與 RSU 明細、2027 年以後公司指引與賣方共識（未取得，不放賣方目標價）、新股融資。**所有同業倍數的市值皆用 2026-11-20 收盤檔收盤價重算；淨現金與營收取頁面值，未更新。**

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A6'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月；美元／ADS）
| 情境 | DCF（推到 12 個月後） | Comparable | 情境目標價 | 收支 FY26–30（RMB M） | 最低現金準備（RMB M） | 資金缺口 | 營業利益轉正 |
|---|---|---|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B {f0(g(f"C{VR['gate']}"))} 百萬人民幣）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC；明確期 FY2027–FY2032）50%＋Comparable 50%（同業 3 家 EV／營收中位數 {inv('ev_med'):.2f} 倍 × FY2027 營收 ＋ 目標時點淨現金）。相對倍數係數 1.00（未調整）。同業只有 3 家且 NTM／TTM 混合，Comparable 是低信心輸入。
- 現價隱含：每 ADS 淨現金 {f2(g(f"B{VR['cash']}"))} 美元（現價的約 {g(f"B{VR['cash']}")/inv('px')*100:.0f}%）；營運業務 EV {f0(g(f"B{VR['mkt']}"))} 百萬人民幣；禾賽現行 EV／TTM 營收 ÷ 同業中位數 {g(f"B{VR['evs2']}"):.2f}。
- 現價檢核：csv 2026-11-20 收盤 11.83 美元／ADS；同檔 2525.HK 收盤 12.58 港元（×8÷7.8426＝12.83 美元，港股收盤早於美股）；2026-11-23 資料商頁面 12.63 美元（收盤檔之後，不採用）。照 csv 建模。

## 營收分部與標準損益（情境 B；RMB M）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 線 6 光達需求（百萬顆） | {yr('B','dem','{:,.2f}',[y for y in YS if y>=2026]) if False else ' | '.join(['—','—']+[P[f"{COL[y]}{PR['B']['dem']}"].value and f"{P[f'{COL[y]}{PR['B']['dem']}'].value:,.2f}" or '—' for y in YS if y>=2026])} |
| 光達出貨量（百萬顆） | {' | '.join(['—','—']+[f"{P[f'{COL[y]}{PR['B']['ship']}'].value:,.2f}" for y in YS if y>=2026])} |
| 涵蓋率 | {' | '.join(['—','—']+[f"{P[f'{COL[y]}{PR['B']['shr']}'].value:.0%}" for y in YS if y>=2026])} |
| 光達 ASP（RMB／顆） | {' | '.join(['—','—']+[f"{P[f'{COL[y]}{PR['B']['asp']}'].value:,.0f}" for y in YS if y>=2026])} |
| 光達營收 | {' | '.join(['—','—']+[f"{P[f'{COL[y]}{PR['B']['lid']}'].value:,.0f}" for y in YS if y>=2026])} |
| 策略成長業務營收 | {' | '.join(['—','—']+[f"{P[f'{COL[y]}{PR['B']['sgi']}'].value:,.0f}" for y in YS if y>=2026])} |
| 營收 | {yr('B','rev','{:,.0f}',YS)} |
| 營收年增 | {yr('B','gr','{:.1%}',YS)} |
| 毛利率 | {yr('B','gm','{:.1%}',YS)} |
| 營業利益 | {yr('B','op','{:,.0f}',YS)} |
| 營業利益率 | {yr('B','opm','{:.1%}',YS)} |
| 淨利 | {yr('B','ni','{:,.0f}',YS)} |
| EPS（人民幣／ADS） | {yr('B','epsa','{:.2f}',YS)} |
| FCFF | {yr('B','fcff','{:,.0f}',[y for y in YS])} |
| 期末現金準備（未計新融資） | {yr('B','liq','{:,.0f}',YS)} |

- FY2024–FY2025 實績；FY2026＝9M 實績＋第四季（營收取公司指引中點 1,125，其餘科目為假設）；FY2027 起為模型。FY2023 只取得營收、毛利、淨利（見 xlsx PL 頁）。
- 命題模組：光達營收＝出貨量 × ASP；出貨量＝線 6 基準需求 × 需求乘數 dm（FY2030 達到）× 涵蓋率（FY2026 校準 {inv('shr26'):.2f}，FY2030 降到情境值）；ASP 以 FY2026 實績校準（情境 B 約 {P[f"F{PR['B']['asp']}"].value:,.0f} 人民幣／顆）後每年降 asd。
- 毛利：光達毛利率 FY2026 殘差校準後線性走向情境 FY2030 值；策略成長業務 25%、服務 20%（皆 Assumed）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 線 6 需求（讀線，標 Assumed）：FY2026 5.06 → FY2030 17.49 → FY2032 24.43 百萬顆；需求乘數 1.5／1.0／0.7（情境 A／B／C；B 區間為線低／高情境 0.41–2.28）。
- 涵蓋率 FY2030：50%／40%／30%（B 區間 25–50%）；ASP 年降 6%／10%／15%（讀線 6 輸入頁）；光達毛利率 FY2030 43%／39%／31%。
- 策略成長業務：FY2026 250（公司指引 200–300）；年增 100%／60%／25%（FY2027–28）、40%／25%／10%（FY2029–32）。
- 費用：研發年增 12%／7%／3%；銷售 5.3% 營收；管理年增 6%；其他營業收益（政府補助）年減 10%；稅率 12%（FY2026 為 0）。
- 折現：rf 5.00%（逾期，跨模型一致）、β {inv('beta'):.3f}（單一來源 Blume；區間 1.0–1.7）、ERP 5%、中國 CRP 0.9%（Assumed）、g＝3%、RONIC＝WACC；人民幣現金流以 2026-11-20 匯率 6.6665 持平換算。
- 股數：Q3'26 稀釋加權 1,305.2 百萬股（÷8＝163.2 百萬 ADS）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 線需求乘數 dm | 中國城市 NOA 搭載率、車廠定點、線 6 追蹤指標 |
| 涵蓋率（shr） | 季度出貨量（Q3 73.0 萬顆，低於公司指引 80–85 萬）、速騰聚創出貨 |
| ASP 年降幅（asd） | 季度 ASP（Q2'26 1,297 人民幣、年減 35%） |
| 光達毛利率（gmt） | 季度毛利率（Q3 41.0%）；管理層稱明年相對穩定 |
| 策略成長業務（gsgi、gsgl） | 致動模組單季出貨（Q3 2.34 萬套）、全年營收是否達 200–300 |

失效條件：{RD['A7'].value[len('失效條件（可觀察）：'):]}。評等翻轉條件見下方反解（Sens 各案例評等：{'、'.join(rts)}）。

**結論主要建立在 Assumed 輸入上：** 需求乘數、涵蓋率、降價幅度、毛利率路徑與策略成長業務成長皆為 Assumed；線 6 需求與降價幅度本身標 Assumed／Analogy。Verified 的只有歷史財報、9M'26 實績、Q4 指引（Interested-party）、現金與股數。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算，格點內插後驗證；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線 6、線 2；xlsx「對帳」頁）
讀入：線 6 `LiDAR 需求量 2030`、`LiDAR 市場規模 2030`（下游取數）；逐年基準需求（計算!C132:I132）、年降價幅度與單價基準（輸入!E112:I112）不在下游取數頁，已在對帳頁註明；線 2 `中國城市 NOA 新車數` 只作交叉檢查（每台 NOA 新車約 0.5 顆光達）。
不一致處：(1) 禾賽 FY2026 校準 ASP（約 {P[f"F{PR['B']['aspu']}"].value:,.0f} 美元）高於線 6 基準 180 美元；(2) 公司 Q2'26 ASP 年減 35%，快於線基準年降 10%；(3) 禾賽營收含割草機等線外需求，涵蓋率因此可高於市占（FY2026 約 {inv('shr26'):.0%}）；(4) 線 6 需求不含機器人致動模組。以上不在本模型靜默修改線值。

## 限制與未找到資料
- 分業務毛利與營收：未找到；FY2023 完整損益：未取得；資本支出、折舊、營業現金流：新聞稿未載，採資料商值（Interested-party）。
- 債務與租賃：由資料商淨現金反推（978 百萬人民幣，區間 500–1,300）；股票投資餘額未取得（不計入）。
- 2027 年以後指引、賣方共識與目標價：未取得；公司 Q3 未再重申 2026 年 300–350 萬顆出貨目標，本模型 FY2026 出貨約 263 萬顆（Assumed）。
- 同業：速騰聚創只有 TTM（無 Forward PS）、Mobileye 沿用 2026-10-26 頁面；中位數只有 3 家，不做迴歸；地平線口徑差大，不進中位數。
- 中國 CRP 未回原文核對（Assumed）。rf 取 2026-10-23 值（逾期，跨模型一致）。β 只有單一來源。
- profile 檔所列 2025 年客戶集中度（前五大 55.8%）未回 20-F 原文核對，本模型未使用。
- 未做：未參與製作的代理獨立核對（待外層流程安排）；display_check.json 未建。

## 核對（建置時執行）
{chk}

## 檔案
- `build_hesai.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`，並跑核對。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)

def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("HESAI_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    px0 = in_val("px", "base")
    ovs = [("base", {})]
    beta_b = 0.67 * 1.47 + 0.33
    meta = {}
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if v is None: continue
            b_ = in_val(k, "base")
            if k == "beta": b_ = beta_b
            if k == "veh_life": b_ = None
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; ovs.append((tag, {k: v}))
            meta[tag] = dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k)
    for s in SC:
        tag = f"only_{s}"; ovs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
        meta[tag] = dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}")
    res = run_cases(ovs); base = res["base"]; print("基準", base)
    cases = []
    for tag, _ in ovs[1:]:
        r = res[tag]; m = meta[tag]
        note = "整本重算；cr_k、k_rt 隨錨點重算" if m["key"] in ("fl25", "fl26", "g_anc", "gm_rk", "gm_ls") else ""
        cases.append(dict(m, tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"], up=r["up"], rt=rating(r["tp"] / px0 - 1), note=note))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2035)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_), note=""))
    # ---- 反解：格點＋內插＋驗證 ----
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)]
    if upper is not None: targets.append((f"升一級邊界（上檔 {upper:+.0%}）", upper))
    if lower is not None: targets.append((f"降一級邊界（上檔 {lower:+.0%}）", lower))
    RV = [("shr_B", 0.40, [0.10 + 0.05 * i for i in range(17)], "情境 B 的 FY2030 涵蓋率（出貨量 ÷ 線 6 需求；其他不變，整本重算；情境 A／C 不動）"),
          ("asd_B", 0.10, [0.0 + 0.02 * i for i in range(14)], "情境 B 的 ASP 年降幅（FY2027 起；情境 A／C 不動）"),
          ("gmt_B", 0.39, [0.15 + 0.03 * i for i in range(14)], "情境 B 的光達毛利率 FY2030（FY2026 校準值起線性；情境 A／C 不動）")]
    revs = []
    for key, b0, grid, note in RV:
        gr = run_cases([(f"g{i}", {key: x}) for i, x in enumerate(grid)])
        ups = [gr[f"g{i}"]["up"] for i in range(len(grid))]
        sol = {}
        for lab, tgt in targets:
            x = None
            for i in range(len(grid) - 1):
                a, b = ups[i] - tgt, ups[i + 1] - tgt
                if a == 0: x = grid[i]; break
                if a * b < 0: x = grid[i] + (grid[i + 1] - grid[i]) * a / (a - b); break
            sol[lab] = x
        okl = [(j, lab, x) for j, (lab, x) in enumerate(sol.items()) if x is not None]
        ver = run_cases([(f"v{j}", {key: x}) for j, lab, x in okl]) if okl else {}
        vmap = {lab: ver[f"v{j}"] for j, lab, x in okl}
        for lab, tgt in targets:
            x = sol[lab]
            if x is None:
                revs.append(dict(label=lab, key=key, val=f"區間 {grid[0]:.2f}～{grid[-1]:.2f} 內無解", base=b0, note=note))
            else:
                u = vmap[lab]["up"]
                revs.append(dict(label=lab, key=key, val=x, base=b0, note=note + f"；驗證上檔 {u:+.1%}"))
            print("反解", key, lab, x)
    fpsum = None
    sens = dict(base=base, cases=cases, fp=fingerprint(), fpsum=0)
    # 成品
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    bid = f"hesai-{VER}-{fingerprint()}"
    tmp = tempfile.mkdtemp(prefix="hesaif_"); raw = os.path.join(tmp, NAME)
    # 先建一次取得 In 指紋（SUM(In!E)）
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o0"))[0]
    fpsum = openpyxl.load_workbook(out, data_only=True)["Sens"]["D5"].value
    sens["fpsum"] = fpsum
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    X = os.path.join(dist, NAME)
    wbv = openpyxl.load_workbook(X, data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("模型狀態", wbv["Sens"]["D7"].value, "tp", wbv["Val"][f"B{VR['tp']}"].value)
    json.dump(dict(build_id=bid, base=base, cases=cases, revs=revs, fp=sens["fp"], fpsum=fpsum), open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1, default=str)
    json.dump(dict(company="禾賽 Hesai", ticker="HSAI", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="RMB M", source=Q425)
                                      for n, l, v in (("Product revenues", "產品營收", 2982.9), ("Service revenues", "服務營收", 44.6))],
                   sublines=[dict(parent="Product revenues", labels=["光達營收（出貨量 × ASP）", "策略成長業務營收（致動模組等）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：光達＋策略成長業務 − 產品營收（應為 0）"],
                   note="公司只揭露產品與服務兩類營收（與出貨量）；光達與策略成長業務為模型子列（預測期 FY2026 起），策略成長業務營收指引 200–300 百萬人民幣。FY2023 分部未取得"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    # 核對
    CHECKS = [["tools/in_layout/check_in_layout.py", X], ["tools/sens_coverage/check_sens_coverage.py", X, "--require-keycol"],
              ["tools/model_governance/check_governance.py", X], ["tools/model_governance/check_rating_gate.py", X, "--strict"],
              ["tools/model_governance/check_zero_rows.py", X], ["tools/model_governance/check_cached_values.py", X],
              ["tools/segment_check/check_segments.py", HERE]]
    results = []
    for cmd in CHECKS:
        rp = subprocess.run([sys.executable] + cmd, capture_output=True, text=True, cwd=ROOT)
        last = [l for l in (rp.stdout + rp.stderr).strip().splitlines() if l.strip()]
        msg = (("通過" if rp.returncode == 0 else "失敗") + "：" + (last[0] if last else "")).replace(ROOT + "/", "")
        results.append((os.path.basename(cmd[0]), msg[:300])); print(cmd[0], rp.returncode, "\n".join(last[:8]))
    HT = X[:-5] + ".html"
    for cmd in (["physical_ai/hesai/gen_html_json.py", X], ["tools/company_html/build_company_html.py", HERE],
                ["tools/company_html/check_company_html.py", HERE], ["tools/company_html/check_layout_v13.py", HT, "--xlsx", X],
                ["tools/model_governance/check_governance.py", X, "--html", HT]):
        rp = subprocess.run([sys.executable] + cmd, capture_output=True, text=True, cwd=ROOT)
        lines = [l for l in (rp.stdout + rp.stderr).strip().splitlines() if l.strip()]
        if cmd[0].endswith(("gen_html_json.py", "build_company_html.py")):
            print(cmd[0], rp.returncode, lines[-1:] if lines else ""); assert rp.returncode == 0, lines[-5:]
            continue
        head = lines[-1] if cmd[0].endswith("check_company_html.py") else (lines[0] if lines else "")
        nm = os.path.basename(cmd[0]) + ("（--html）" if "--html" in cmd else "")
        results.append((nm, (("通過" if rp.returncode == 0 else "失敗") + "：" + head[:300]).replace(ROOT + "/", ""))); print(cmd[0], rp.returncode, "\n".join(lines[:10]))
    readme(X, VR, PR, DR, bid, sens, revs, results)

if __name__ == "__main__":
    main()
if __name__ == "__main__":
    main()
