# 舜宇光學（2382.HK）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：車載與機器人相機模組；並讀線 4 人形）
# 用法：python3 physical_ai/sunny/build_sunny.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算，香港財務報告準則）。金額：百萬人民幣；股數：百萬股；EPS：人民幣元；股價：港元（以收盤檔匯率換算）。
# 公司 2025 年報起把分部改為手機、車載、XR、其他四類；本模型照新分部建模（FY2024 為公司重編值，FY2023 未重編）。
# 結構：README、In、Hist、Q（官方分部對帳）、PL（三情境：分部營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_舜宇光學個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
H126 = "公司 2026 中期業績公告〔港交所，2026-08-26〕https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0826/2026082601115.pdf"
FY25 = "公司 2025 全年業績公告〔港交所，2026-03-30；含 2024 重編分部〕https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0330/2026033003334.pdf"
FY24 = "公司 2024 全年業績公告〔港交所，2025-03-24；2023 比較值〕https://www1.hkexnews.hk/listedco/listconews/sehk/2025/0324/2025032400848.pdf"
QTR = "Quartr 舜宇 H1 2026 業績摘要（資料商，2026-08-26）https://quartr.com/events/sunny-optical-technology-group-company-limited-2382-h1-2026_FYd9Pb9H"
SA = "stockanalysis.com 舜宇統計頁（資料商；頁面 2026-11-24 15:00 HKT，價 63.00 港元）https://stockanalysis.com/quote/hkg/2382/statistics/"
SA3008 = "stockanalysis.com 大立光統計頁（頁面 2026-11-24，價 3,980 新台幣）https://stockanalysis.com/quote/tpe/3008/statistics/"
SA3406 = "stockanalysis.com 玉晶光統計頁（頁面 2026-11-24，價 734 新台幣）https://stockanalysis.com/quote/tpe/3406/statistics/"
SA011 = "stockanalysis.com LG Innotek 統計頁（頁面 2026-11-24，價 284,500 韓元）https://stockanalysis.com/quote/krx/011070/statistics/"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
CNRF = "人民幣 10 年期國債殖利率 1.61%（與 physical_ai/geekplus In 頁 rf 同值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
GKX = "physical_ai/geekplus/dist/20261123_極智嘉個股模型_v0.1.xlsx"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

SEGS = [("hs", "手機產品"), ("veh", "車載產品"), ("xr", "XR 產品"), ("oth", "其他產品（機器人、智慧影像、泛物聯網）")]
sec("A. FY2025 實績：分部營收（2025 年報新分部）")
for (k, n), v in zip(SEGS, (27324.044, 7327.752, 2393.466, 6183.864)):
    inp(f"a25_{k}", f"營收 FY2025：{n}", "百萬人民幣", None, v, None, "Verified", FY25 + " 分部資訊", "四分部合計 43,229.126 ＝ 合併營收")

sec("B. 命題模組一：四個分部各自成長（FY2026 以上半年實績為錨、FY2027–28、FY2029–32 兩段；情境 A／B／C）")
# 鍵族 → {分部: (A, B, C, B 低, B 高, 說明)}
GRO = {
 "g26": ("FY2026", {"hs": (0.02, -0.01, -0.05, -0.04, 0.02, "上半年 −0.8%（出貨量減、均價升，" + H126 + "）；全球手機出貨下滑"),
                    "veh": (0.14, 0.11, 0.07, 0.07, 0.14, "上半年 +10.6%；2025 年 +21.3%（" + FY25 + "）"),
                    "xr": (0.14, 0.08, 0.0, 0.0, 0.14, "上半年 +11.0%（智慧眼鏡、AR 眼鏡）；2025 年 −7.1%"),
                    "oth": (0.55, 0.45, 0.35, 0.35, 0.55, "上半年 +88.5%（機器人、智慧影像、泛物聯網）；下半年基期 2025H2 約 4,041 已高，故全年低於上半年增速")}),
 "g1": ("FY2027–FY2028 每年", {"hs": (0.01, 0.0, -0.04, -0.04, 0.01, "手機相機模組均價升、量降；無證據支持成長"),
                    "veh": (0.15, 0.10, 0.04, 0.04, 0.15, "車載近三年 +21%、+10.6%；每車相機顆數增加 vs 車市放緩"),
                    "xr": (0.25, 0.12, 0.0, 0.0, 0.25, "XR 2025 年下滑、2026 上半年回升；基期小、波動大"),
                    "oth": (0.25, 0.12, 0.0, 0.0, 0.25, "基期 2026 已 +45%；機器人視覺為主要動能，持續性無證據")}),
 "g2": ("FY2029–FY2032 每年", {"hs": (0.01, 0.0, -0.03, -0.03, 0.01, ""), "veh": (0.12, 0.08, 0.03, 0.03, 0.12, ""),
                    "xr": (0.20, 0.10, 0.02, 0.02, 0.20, ""), "oth": (0.15, 0.08, 0.02, 0.02, 0.15, "")}),
}
for fam, (yl, d) in GRO.items():
    for k, n in SEGS:
        a, b, c, lo, hi, note = d[k]
        for s, v in zip(SC, (a, b, c)):
            inp(f"{fam}{k}_{s}", f"情境 {s}：{n}營收年增 {yl}", "%/年", lo if s == "B" else None, v, hi if s == "B" else None, "Assumed",
                note or "A／C 以 B 上下設定", "歷史與上半年實績為依據，未來成長無公司量化指引" if s == "B" else "")
inp("gd_floor", "公司全年指引：營收與淨利（不含換股收益）年增不低於", "%", None, 0.07, None, "Interested-party", QTR + "：全年營收與淨利（不含換股）年增不低於 7%", "只取得資料商摘要，未取得公告原文；作 Recon 對照下限，不進公式", CO)

sec("C. 命題模組二：機器人與 Robotaxi 相機模組＝線 6 市場 × 舜宇份額 × 爬坡 × 匯率（命題子列；與四分部成長重疊的部分不另扣，見 README）")
for yy, ci in ((2026, "C"), (2030, "G"), (2032, "I")):
    vals = {("A", 2026): 36.3125, ("B", 2026): 16.125, ("C", 2026): 6.55, ("A", 2030): 339.63245568, ("B", 2030): 82.38661632, ("C", 2030): 15.71291136,
            ("A", 2032): 819.93878470656, ("B", 2032): 196.863584108544, ("C", 2032): 25.831007059968}
    rown = {"A": 222, "B": 155, "C": 88}
    for s in SC:
        inp(f"m{str(yy)[2:]}_{s}", f"情境 {s}：機器人與 Robotaxi 相機模組市場 {yy}（線 6）", "百萬美元", None, vals[(s, yy)], None, "Derived",
            f"{L6} 計算!{ci}{rown[s]}（CAM 列；A＝高、B＝基準、C＝低需求情境）", "「下游取數」頁未列相機模組，故讀「計算」頁；年間以幾何插值（Derived）")
for s, v, lo, hi in (("A", 0.15, None, None), ("B", 0.08, 0.03, 0.15), ("C", 0.03, None, None)):
    inp(f"shr_{s}", f"情境 {s}：舜宇在機器人與 Robotaxi 相機模組的份額", "%", lo, v, hi, "Assumed",
        f"未取得：線 6 份額頁把舜宇列為相機模組主要廠、份額「未取得」（{L6} 份額!F21）", "區間 3–15% 無證據支撐；子列規模小（見 Val），結論不靠它")
inp("rp27", "子列爬坡：FY2027 占穩態比例（線性升到 FY2030＝100%）", "%", 0.0, 0.25, 0.5, "Assumed", "比照 physical_ai/thk、hiwin 爬坡設定", "")
inp("fx_usd", "人民幣兌美元", "RMB/US$", 6.4, 6.6665, 7.0, "Verified", YCL + "：CNY=X 2026-11-20 收盤 6.6665", "低／高為 Assumed")

sec("D. 損益表結構（標準損益表；香港財務報告準則）")
GMI = {"hs": (0.10, 0.13, 0.15, "FY2024 12.9%、FY2025 14.7%、上半年 12.3%（Derived，" + H126 + "、" + FY25 + "）；手機模組高階占比升但成本壓力"),
       "veh": (0.28, 0.305, 0.33, "FY2024 35.3%、FY2025 31.9%、上半年 30.3%；價格競爭與規格組合"),
       "xr": (0.11, 0.135, 0.19, "FY2024 11.8%、FY2025 19.6%、上半年 13.0%；波動大"),
       "oth": (0.28, 0.32, 0.35, "FY2024 29.3%、FY2025 27.5%、上半年 34.4%（機器人與智慧影像組合）")}
for k, n in SEGS:
    lo, b, hi, note = GMI[k]
    inp(f"gm_{k}", f"分部毛利率：{n}（FY2026 下半年起）", "%", lo, b, hi, "Assumed", note, "分部毛利＝分部營收 − 分部銷貨成本（公司分部資訊「分部利潤」口徑）；子列以其他產品毛利率計")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.01, 0.005), ("C", -0.01, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整（價格、組合）", "pt", lo, v, hi, "Assumed", "合併毛利率：FY2023 14.5%、FY2024 18.3%、FY2025 19.7%、上半年 19.5%（Derived）", "")
inp("sga_el", "營業費用（研發＋銷售＋管理）對營收的彈性（費用年增 ÷ 營收年增）", "x", 0.1, 0.2, 0.5, "Derived",
    "上半年費用 −0.3% ÷ 營收 +11.5% ≈ 0；FY2025 費用 +5.4% ÷ 營收 +12.9% ＝ 0.42（Derived，" + H126 + "、" + FY25 + "）", "研發 FY2025 3,258.6（營收 7.5%）；基準 0.2")
inp("oth_op", "其他營業損益淨額（其他收入＋其他損益＋減損＋持分法）FY2026 起", "百萬人民幣", 800.0, 1100.0, 1450.0, "Assumed",
    "FY2024 1,399.6；FY2025 2,357.8（含換股收益 918.9，不重複）；上半年 502.0（Derived，" + H126 + "、" + FY25 + "）", "主要為利息收入、理財收益、政府補助與聯營；上半年年化約 1,000")
inp("nonop", "業外：財務費用（負值，FY2026 起）", "百萬人民幣", -450.0, -330.0, -280.0, "Assumed", "FY2024 −499.4、FY2025 −436.2、上半年 −163.7（Derived）", "")
inp("tax_w", "有效稅率（FY2026 起）", "%", 0.10, 0.115, 0.15, "Assumed", "FY2024 11.7%、FY2025 11.2%、上半年 11.5%（Derived，高新技術企業優惠稅率）", "稅率優惠到期或海外比重上升屬風險，見 Risk")
inp("nci_pct", "非支配股東損益 ÷ 稅後淨利（FY2026 起）", "%", 0.04, 0.06, 0.09, "Derived", "FY2025 3.5%、上半年 7.3%（141.6 ÷ 1,950.0，Derived，" + H126 + "）", "")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.05, 0.057, 0.065, "Derived", "FY2025 2,305.5 ÷ 43,229 ＝ 5.3%；上半年折舊 1,348.3（年化約 2,700）（Derived）", "")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.055, 0.07, 0.085, "Derived", "FY2025 2,931.0 ÷ 43,229 ＝ 6.8%；上半年 1,636.9（Derived，" + H126 + "）", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.05, 0.10, 0.15, "Assumed", "2025-12-31 存貨 7,179.7＋應收 10,012.7＋票據 1,508.9，應付拆分未取得；以約 10% 營收為上限代理", "")
inp("payout", "股利支付率（占前一年歸屬母公司淨利）", "%", 0.20, 0.25, 0.30, "Interested-party", FY25 + "：目標派息率由 20% 提高到 25%；FY2025 期末息 120.6 港仙（2026-06 付）", "只用於收支關卡與 12 個月後每股價值的股利扣除；不進 DCF", CO)

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30）")
inp("px", "股價（2026-11-20 收盤，港元）", "港元", None, 58.5, None, "Verified", YCL + "：2382.HK 2026-11-20 收盤 58.5", "未還原股利。11-13 收盤 65.85 → 11-20 58.5（−11.2%），當週 AI 與半導體股全面下挫")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("fx_hk", "港元兌人民幣（1 人民幣＝幾港元；＝HKD=X ÷ CNY=X）", "HKD/RMB", None, "=7.8426/6.6665", None, "Derived", YCL + "：HKD=X 7.8426、CNY=X 6.6665（2026-11-20）", "模型以人民幣計算，每股價值乘此值轉港元")
inp("tr_sh", "股份獎勵信託持有股數（估）", "百萬股", 0.0, 5.5, 8.6, "Derived", "2025H1 加權基本股數 1,086.161 與已發行 1,094.805 差 8.6；2026H1 加權 1,076.196 反推約 5.5（假設買回股數按期間加權）", "未取得信託持股期末值")
inp("sh_out", "評價日股數（已發行 1,094.805 − 庫存股 15.840 − 信託持股）", "百萬股", None, "=1094.8048-15.84-tr_sh", None, "Derived", H126 + "：已發行 1,094,804,800 股；2026-01 買回 15,840,000 股作庫存股（Verified）", "僱員股權激勵未建模")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "＝上列", "")
inp("cash_v", "現金及銀行存款（現金 3,499.1＋定期 1,687.3＋短期定存 1,685.3＋質押 14.8；2026-06-30）", "百萬人民幣", None, "=3499.067+1687.299+1685.304+14.774", None, "Verified", H126 + " 資產負債表", "")
inp("fv_v", "以公允價值計入損益的金融資產（理財產品等；2026-06-30）", "百萬人民幣", None, 13326.538, None, "Verified", H126 + "：13,326.5（2025 年底 11,407.4）", "內容為理財與國債類，流動性假設未逐項查核")
inp("fv_k", "理財資產計入比例", "x", 0.7, 1.0, None, "Assumed", "基準帳面全額；低案 70%", "")
inp("debt_v", "借款及債券（銀行借款流動 2,589.9＋非流動 500.0＋應付債券 2,798.0）", "百萬人民幣", None, "=2589.877+500+2797.993", None, "Verified", H126, "")
inp("lease_v", "租賃負債（流動 50.2＋非流動 113.4）", "百萬人民幣", None, "=50.249+113.423", None, "Verified", H126, "")
inp("mi_v", "非支配股東權益（帳面）", "百萬人民幣", None, 697.015, None, "Verified", H126, "帳面值；公允價值未取得")
inp("stub_frac", "Stub 期（2026-06-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,6,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（人民幣 10 年期國債）", "%", None, 0.0161, None, "Verified", CNRF, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬", "%", None, 0.0091, None, "Verified", "沿用 " + GKX + " In 頁 crp（0.91%，跨模型一致）", "")
inp("beta", "β（舜宇與同業 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.22+0.33,0.67*MEDIAN(1.16,1.04,2.63)+0.33)", "=E{r}+0.3", "Derived",
    SA + "：舜宇 5 年 β 1.22；" + SA3008 + "：大立光 1.16；" + SA3406 + "：玉晶光 1.04；" + SA011 + "：LG Innotek 2.63", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, 0.04, None, "Assumed", "FY2025 財務費用 436.2（含匯兌與租賃利息）÷ 平均借款債券約 7,000 ≈ 6% 為上限；美元借款為主", "")
inp("g_cap", "永續成長率上限", "%", 0.01, 0.03, 0.04, "Assumed", "DCF 標準規格：3%，且不超過 rf → g＝MIN(g_cap, rf)＝1.61%", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺舜宇對同業歷史前瞻 P/E 比值）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（資料商頁面 2026-11-24 的前瞻 P/E 以「報價日價 ÷ 收盤檔 2026-11-20 收盤」換算到同一天，分母 EPS 不變）")
PEERS = [("pe_lar", "大立光 3008（手機鏡頭）", "=19.9*3670/3980", SA3008 + "：前瞻 19.90；收盤檔 11-20 收盤 3,670 ÷ 頁面價 3,980"),
         ("pe_gyc", "玉晶光 3406（手機鏡頭）", "=16.8*648/734", SA3406 + "：前瞻 16.80；收盤檔 11-20 收盤 648 ÷ 頁面價 734"),
         ("pe_lgi", "LG Innotek 011070（相機模組）", "=7.75*267500/284500", SA011 + "：前瞻 7.75；收盤檔 11-20 收盤 267,500 ÷ 頁面價 284,500")]
for k, lab, v, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Derived", src, "會計口徑：股權口徑 P/E 不受租賃處理影響；賣方共識 EPS（逾期之處已換算）", DV)
inp("pe_med", "同業前瞻 P/E 中位數（3 家；進 Comparable）", "x", "=pe_lgi", "=MEDIAN(pe_lar,pe_gyc,pe_lgi)", "=pe_lar", "Derived",
    "3 家中位數；同業數少、系統性 R² 無從檢驗（少於約 8 家不做迴歸）", "低＝LG Innotek（Analogy：相機模組、車載）；高＝大立光（Analogy：手機鏡頭龍頭）。豪威 603501（上游 CIS，報價日未取得：統計頁 404）、歌爾、丘鈦：收盤檔無，不納入")
inp("pe_self", "舜宇現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=13.16*58.5/63", None, "Derived", SA + "：前瞻 13.16 × 收盤檔 58.5 ÷ 頁面價 63.00", "只作對照，不進模型", DV)
inp("cons_tp", "賣方平均目標價（21 位；資料商頁面 2026-11-24，港元）", "港元", None, 83.53, None, "Interested-party", SA + "：平均 83.53、評等 Buy", "只作對照，不進模型", DV)

inp("fin_inc", "利息收入與理財投資收益（含在其他營業損益內；不進 FCFF，因理財與存款已以淨現金計價）", "百萬人民幣", 400.0, 560.0, 700.0, "Derived",
    "上半年：存款利息 158.8＋理財投資收益 121.3（FY2025H1 142.0＋168.1），年化約 560（Derived，" + H126 + "）", "避免淨現金與其利息收入重複計價")
def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額百萬人民幣、股數百萬股、EPS 人民幣元、股價港元"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            if isinstance(v, str): v = v.replace("{r}", str(r))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = ([(f"{fam}{k}", f"{n}營收年增 {GRO[fam][0]}") for fam in GRO for k, n in SEGS] +
           [("m26", "機器人與 Robotaxi 相機模組市場 2026（線 6）"), ("m30", "機器人與 Robotaxi 相機模組市場 2030（線 6）"), ("m32", "機器人與 Robotaxi 相機模組市場 2032（線 6）"),
            ("shr", "舜宇機器人與 Robotaxi 相機模組份額"), ("gmadj", "FY2027 起毛利率調整")])
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；百萬人民幣）。FY2024 分部為公司 2025 年報重編值；FY2023 為舊分部（光學零件、光電產品、光學儀器），不可比，分部與費用拆分未取得。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    NA = "未取得"
    S25 = "公司 2025 全年業績公告（網址見 Sources 頁）"; S24 = "公司 2024 全年業績公告（網址見 Sources 頁）"
    HR = [("h_rev", "營收", 31681.3, 38294.491, 43229.126, "Verified", S24 + "（FY2023、FY2024）；" + S25),
          ("h_cogs", "營業成本", "=B3-B5", 31288.458, 34713.238, "Verified", "FY2023＝營收 − 毛利（Derived）；" + S25),
          ("h_gp", "毛利", 4590.441, 7006.033, 8515.888, "Verified", S24 + "；" + S25),
          ("h_sga", "研發、銷售及管理費用", NA, "=2924.111+439.329+1399.231", "=3258.580+342.945+1419.957", "Verified", S25 + "（研發＋銷售及分銷＋行政；FY2023 拆分未取得）"),
          ("h_oth", "其他營業損益淨額（其他收入＋其他損益＋減損＋聯營）", NA, "=1333.065-51.792-0.099+118.381", "=1100.398+1091.976-6.228+171.591", "Verified",
           S25 + "；FY2025 含換股收益 918.920（一次性）"),
          ("h_op", "營業利益（本模型口徑：毛利 − 費用 ＋ 其他營業損益）", NA, "=C5-C6+C7", "=D5-D6+D7", "Derived", "公司不揭露營業利益，本模型以損益表各列重組；加財務費用後等於稅前利益（見 PL 驗算）"),
          ("h_fin", "業外：財務費用", NA, -499.383, -436.190, "Verified", S25),
          ("h_pbt", "稅前利益", 1358.153, 3143.534, 5415.953, "Verified", S24 + "；" + S25),
          ("h_tax", "所得稅", "=B10-B12", 366.510, 608.963, "Verified", S25 + "；FY2023＝稅前 − 稅後（Derived）"),
          ("h_cont", "稅後淨利（合併）", 1150.323, 2777.024, 4806.990, "Verified", S24 + "；" + S25),
          ("h_nci", "非支配股東損益", NA, 77.865, 167.883, "Verified", S25),
          ("h_ni", "歸屬母公司淨利", NA, 2699.159, 4639.107, "Verified", S25 + "（FY2023 未取得）"),
          ("h_eps", "基本 EPS（人民幣元）", NA, 2.482, 4.2655, "Verified", S25 + "：FY2025 426.55 分、FY2024 248.2 分（年增 71.9%、146.4%）"),
          ("h_sh", "加權基本股數（百萬股）", NA, 1087.683, 1087.584, "Verified", S25),
          ("h_da", "折舊（不動產廠房設備＋投資性不動產＋使用權資產）", NA, "=2100.639+4.876+64.553", "=2231.153+4.839+69.543", "Verified", S25),
          ("h_cx", "資本支出", NA, NA, 2931.0, "Verified", S25),
          ("h_hs", "分部營收：手機產品", NA, 25155.480, 27324.044, "Verified", S25),
          ("h_veh", "分部營收：車載產品", NA, 6039.062, 7327.752, "Verified", S25),
          ("h_xr", "分部營收：XR 產品", NA, 2577.280, 2393.466, "Verified", S25),
          ("h_oth_s", "分部營收：其他產品", NA, 4522.669, 6183.864, "Verified", S25),
          ("h_ig", "其中：換股收益（一次性，含在其他營業損益內）", NA, 0, 918.920, "Verified", S25 + "（Goertek Optical 換股）")]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 52

    # ---------------- Q（官方分部對帳） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（手機、車載、XR、其他；2025 年報起）營收與分部毛利（百萬人民幣）。分部對帳：分部合計 − 合併數（應為 0）。2025H1 為公司在 2026 中期公告的重編值"
    hdr(Q, 2, ["項目", "FY2024", "FY2025", "H1'25", "H1'26", "標記", "來源", "來源誘因"])
    segv = {"手機產品": (25155.480, 27324.044, 13248.009, 13148.628, 3244.600, 4004.411, 1899.622, 1615.297),
            "車載產品": (6039.062, 7327.752, 3400.448, 3761.928, 2134.026, 2340.844, 1137.357, 1138.644),
            "XR 產品": (2577.280, 2393.466, 860.643, 955.720, 303.316, 468.922, 148.879, 124.503),
            "其他產品": (4522.669, 6183.864, 2142.819, 4039.428, 1324.091, 1701.711, 708.590, 1390.222)}
    for i, (n, v) in enumerate(segv.items(), 3):
        Q.cell(i, 1, f"{n}營收")
        for j in range(4): Q.cell(i, 2 + j, v[j]).font = BLUE
        Q.cell(i, 6, "Verified"); Q.cell(i, 7, "公司 2025 全年與 2026 中期業績公告 分部資訊（網址見 Sources 頁）"); Q.cell(i, 8, "發行公司（立場：自家財報）")
    Q.cell(7, 1, "分部營收合計")
    for j, L in enumerate("BCDE"): Q.cell(7, 2 + j, f"=SUM({L}3:{L}6)")
    Q.cell(8, 1, "合併營收")
    for j, v in enumerate((38294.491, 43229.126, 19651.919, 21905.704)): Q.cell(8, 2 + j, v).font = BLUE
    Q.cell(8, 6, "Verified"); Q.cell(8, 7, "損益表")
    Q.cell(9, 1, "檢查：分部營收合計 − 合併營收（應為 0）")
    for j, L in enumerate("BCDE"): Q.cell(9, 2 + j, f"=ROUND({L}7-{L}8,3)")
    for i, (n, v) in enumerate(segv.items(), 10):
        Q.cell(i, 1, f"{n}分部毛利")
        for j in range(4): Q.cell(i, 2 + j, v[4 + j]).font = BLUE
        Q.cell(i, 6, "Verified"); Q.cell(i, 7, "同上 分部利潤（＝分部毛利）"); Q.cell(i, 8, "發行公司（立場：自家財報）")
    Q.cell(14, 1, "分部毛利合計")
    for j, L in enumerate("BCDE"): Q.cell(14, 2 + j, f"=SUM({L}10:{L}13)")
    Q.cell(15, 1, "合併毛利")
    for j, v in enumerate((7006.033, 8515.888, 3894.448, 4268.666)): Q.cell(15, 2 + j, v).font = BLUE
    Q.cell(15, 6, "Verified"); Q.cell(15, 7, "損益表")
    Q.cell(16, 1, "檢查：分部毛利合計 − 合併毛利（應為 0）")
    for j, L in enumerate("BCDE"): Q.cell(16, 2 + j, f"=ROUND({L}14-{L}15,3)")
    for i, n in enumerate(segv, 17):
        Q.cell(i, 1, f"{n}毛利率")
        for j, L in enumerate("BCDE"):
            Q.cell(i, 2 + j, f"={L}{i-7}/{L}{i-14}").number_format = "0.0%"
        Q.cell(i, 6, "Derived")
    Q.cell(21, 1, "合併毛利率")
    for j, L in enumerate("BCDE"): Q.cell(21, 2 + j, f"={L}15/{L}8").number_format = "0.0%"
    Q.cell(21, 6, "Derived")
    Q.cell(22, 1, "註：2025 年報起分部改為四類；FY2023 為舊分部（光學零件、光電產品、光學儀器），未重編，不列。H1'26 車載收入年增 10.6%、其他產品 +88.5%（公司 MD&A）")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：四分部營收＋機器人與 Robotaxi 相機模組子列 → 標準損益表（香港財務報告準則）→ 現金流（百萬人民幣；FY2023–25 實績、FY2026 起模型，FY2026 以上半年實績為錨）"
    PR = {}
    rr = 3
    KEYS = ([f"c_{k}" for k, _ in SEGS] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont",
             "nci", "ni", "nim", "sh", "eps", "nig", "da", "cx", "dwc", "finc", "fcff", "div", "fcfad", "mkt", "rp"])
    LAB = ([f"營收：{n}" for _, n in SEGS] +
           ["四分部營收合計", "機器人與 Robotaxi 相機模組（命題子列；線 6 市場 × 份額）", "營收（合併）", "檢查：分部合計 − 四分部營收合計", "檢查：四分部＋子列 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "研發、銷售及管理費用", "其他營業損益淨額（其他收入＋其他損益＋減損＋聯營）", "營業費用（費用 − 其他營業損益淨額）",
            "營業利益", "營業利益率", "業外：財務費用", "稅前利益", "所得稅", "稅後淨利（合併）",
            "非支配股東損益", "稅後淨利（歸屬母公司）", "淨利率", "平均股數（百萬股）", "EPS（基本，人民幣元）", "淨利年增率",
            "折舊攤銷", "資本支出", "營運資金增加", "其中利息與理財收益（不進 FCFF）", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加；NOPAT 不含利息與理財收益）", "股利（前一年淨利 × 支付率）",
            "自由現金流扣股利（收支關卡用）", "參考：機器人與 Robotaxi 相機模組市場（百萬美元，線 6）", "參考：子列爬坡（占穩態）"])
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                yy = str(y)[2:]
                for k, hk in (("rev", "rev"), ("cogs", "cogs"), ("gp", "gp"), ("sga", "sga"), ("oth", "oth"), ("op", "op"), ("fin", "fin"), ("pbt", "pbt"), ("tax", "tax"),
                              ("cont", "cont"), ("nci", "nci"), ("ni", "ni"), ("eps", "eps"), ("sh", "sh"), ("da", "da"), ("cx", "cx")):
                    w(k, f"=h_{hk}_{yy}")
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("nim", f'=IF(ISNUMBER({cell("ni")}),{cell("ni")}/{cell("rev")},"未取得")')
                w("opex", f'=IF(ISNUMBER({cell("sga")}),{cell("sga")}-{cell("oth")},"未取得")'); w("opm", f'=IF(ISNUMBER({cell("op")}),{cell("op")}/{cell("rev")},"未取得")')
                if y >= 2024:
                    for (k, _), hk in zip(SEGS, ("hs", "veh", "xr", "oth_s")): w(f"c_{k}", f"=h_{hk}_{yy}")
                    w("core", f"=SUM({c}{R['c_hs']}:{c}{R['c_oth']})"); w("hum", 0); w("rp", 0)
                    w("chk1", f"=ROUND(SUM({c}{R['c_hs']}:{c}{R['c_oth']})-{cell('core')},6)"); w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},3)")
                else:
                    for k, _ in SEGS: w(f"c_{k}", "未取得（舊分部）")
                    for k in ("core", "hum", "chk1", "chk2"): w(k, "未取得（舊分部）")
                w("gr", f"={cell('rev')}/{prev('rev')}-1" if y > 2023 else '="未取得"')
                w("nig", f"={cell('ni')}/{prev('ni')}-1" if y == 2025 else '="未取得"')
                if y == 2023:
                    for k in ("sga", "oth", "op", "fin", "ni", "eps", "sh", "nci", "da", "cx"): w(k, "未取得")
                continue
            if y == 2026:
                for k, _ in SEGS: w(f"c_{k}", f"=a25_{k}*(1+g26{k}_{s})")
                w("hum", 0); w("rp", 0); w("mkt", f"=m26_{s}")
            else:
                gk = "g1" if y <= 2028 else "g2"
                for k, _ in SEGS: w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}{k}_{s})")
                if y <= 2029: mk = f"=m26_{s}*(m30_{s}/m26_{s})^(({y}-2026)/4)"
                elif y == 2030: mk = f"=m30_{s}"
                elif y == 2031: mk = f"=m30_{s}*(m32_{s}/m30_{s})^0.5"
                else: mk = f"=m32_{s}"
                w("mkt", mk); w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)")
                w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}*fx_usd")
            w("core", f"=SUM({c}{R['c_hs']}:{c}{R['c_oth']})"); w("rev", f"={cell('core')}+{cell('hum')}")
            w("chk1", f"=ROUND(SUM({c}{R['c_hs']}:{c}{R['c_oth']})-{cell('core')},6)"); w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            adj = f"+gmadj_{s}*{cell('rev')}" if y >= 2027 else ""
            w("gp", f"={cell('c_hs')}*gm_hs+{cell('c_veh')}*gm_veh+{cell('c_xr')}*gm_xr+({cell('c_oth')}+{cell('hum')})*gm_oth{adj}")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("sga", f"={prev('sga')}*(1+sga_el*({cell('rev')}/{prev('rev')}-1))"); w("oth", "=oth_op")
            w("opex", f"={cell('sga')}-{cell('oth')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("fin", "=nonop"); w("pbt", f"={cell('op')}+{cell('fin')}"); w("tax", f"={cell('pbt')}*tax_w"); w("cont", f"={cell('pbt')}-{cell('tax')}")
            w("nci", f"={cell('cont')}*nci_pct"); w("ni", f"={cell('cont')}-{cell('nci')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("sh", "=sh_v"); w("eps", f"={cell('ni')}/{cell('sh')}")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("finc", "=fin_inc"); w("fcff", f"=({cell('op')}-{cell('finc')})*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", f"=payout*{prev('ni')}"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(R["c_hs"], 2, "FY2026＝FY2025 ×(1＋g26)；FY2027–28 ×(1＋g1)；FY2029 起 ×(1＋g2)")
        P.cell(R["hum"], 2, "市場 × 份額 × 爬坡 × 匯率；FY2026 以前視為 0；與分部成長可能重疊，不另扣")
        P.cell(R["gp"], 2, "Σ 分部營收 × 分部毛利率（子列以其他產品毛利率）＋ FY2027 起毛利率調整 × 營收")
        P.cell(R["sga"], 2, "＝前一年 ×(1＋彈性 × 營收年增)")
        P.cell(R["op"], 2, "毛利 − 費用 ＋ 其他營業損益；FY2025 含換股收益 918.9（一次性）")
        P.cell(R["fcff"], 2, "NOPAT＝(營業利益 − 利息與理財收益)×(1 − tax_w)；股票報酬未取得")
        P.cell(R["div"], 2, "FY2026 實付約 1,140（FY2025 期末息 105.9 分 × 股數）")
        for k in ("gr", "gm", "opm", "nim", "nig", "rp"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad", "hum"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（百萬人民幣）"; D["B4"] = "=px/fx_hk*sh_v"
    D["A5"] = "負債 D（借款債券＋租賃負債）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-06-30）＝現金 − 借款債券 − 租賃負債 − 非支配 ＋ 理財資產 × fv_k"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v+fv_k*fv_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px/fx_hk*sh_v)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px/fx_hk*sh_v)+debt_v+lease_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益 − 利息與理財收益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}-PL!{pc}{R['finc']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-06-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日，港元）", f"=MAX(0,B{t+8})/sh_v*fx_hk"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-payout*PL!F{R['eps']}*fx_hk"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*fx_hk*(1+ke)^t_tgt-payout*PL!F{R['eps']}*fx_hk")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 FY2028 EPS × 匯率，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS（人民幣，× 港元匯率）", [f"=pe_med*PL!H{PR[s]['eps']}*fx_hk" for s in SC],
                "目標時點的 NTM 以 FY2028 代（2027-11 起算 NTM 約 1/12 FY2027＋11/12 FY2028）；分母：模型各情境 FY2028 EPS")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（港元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 21 位平均；資料商 2026-11-24 頁面）", "=cons_tp", "只作對照（Interested-party）")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（百萬人民幣；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"0.00"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷（情境 B FY2028 EPS × 匯率）", f"=B{r_tp}/(PL!H{PR['B']['eps']}*fx_hk)", "分母：模型情境 B FY2028 EPS（港元）")
    r_pe1 = one("現價 ÷（情境 B FY2027 EPS × 匯率）", f"=px/(PL!G{PR['B']['eps']}*fx_hk)", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
    r_hum = one("機器人與 Robotaxi 相機模組子列占 FY2030 合併營收（情境 B）", f"=PL!J{PR['B']['hum']}/PL!J{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, cons=r_cons, wk=r_wk, hum=r_hum)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收 vs 指引下限（FY2025 ×(1＋7%)）", f"=PL!F{PR['B']['rev']}", "=h_rev_25*(1+gd_floor)", "公司全年指引「不低於 7%」（資料商摘要，Interested-party；未取得公告原文）"),
          ("FY2026 歸屬母公司淨利 vs 指引下限（（FY2025 淨利 − 換股收益）×(1＋7%)）", f"=PL!F{PR['B']['ni']}", "=(h_ni_25-h_ig_25)*(1+gd_floor)", "換股收益全數歸屬母公司為假設（歸屬拆分未取得）"),
          ("FY2026 隱含下半年營收 vs 2025 下半年（合併 − 上半年）", f"=PL!F{PR['B']['rev']}-Q!E8", "=Q!C8-Q!D8", "上半年實績 21,905.7（+11.5%）；差為隱含下半年年增額"),
          ("FY2026 毛利率 vs 2026 上半年實績", f"=PL!F{PR['B']['gm']}", "=Q!E15/Q!E8", "2025 年下半年毛利率 19.6%（Derived）"),
          ("FY2027 EPS（模型 B，港元）vs 現價 ÷ 現行前瞻 P/E 隱含 NTM EPS", f"=PL!G{PR['B']['eps']}*fx_hk", "=px/pe_self", "資料商 NTM（約 FY2026 末至 FY2027 初起算，賣方共識）vs 模型 FY2027；口徑不同，只看方向"),
          ("加權目標價 vs 賣方平均目標價（港元）", "=tp", "=cons_tp", "賣方 21 位平均（2026-11-24 頁面，只作對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_sunny.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_sunny.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，港元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，港元）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.005,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.00"; SE[f"F{sr}"].number_format = "#,##0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整）。FY2026 以前的成長鍵（g26*）只列情境 B；情境 A／C 以「只看情境」案例呈現。無區間的鍵不列案例。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線模型成品；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("機器人與 Robotaxi 相機模組市場 2026（低／基準／高）", "6", "6.55／16.13／36.31", "百萬美元", "機器人與 Robotaxi（不含既有車用相機）；A＝高、C＝低需求情境", L6 + " 計算!C88、C155、C222", "Derived", "In m26_C／m26_B／m26_A → PL 子列", "一致（照取）"),
           ("機器人與 Robotaxi 相機模組市場 2030（低／基準／高）", "6", "15.71／82.39／339.63", "百萬美元", "同上", L6 + " 計算!G88、G155、G222", "Derived", "In m30_*", "一致"),
           ("機器人與 Robotaxi 相機模組市場 2032（低／基準／高）", "6", "25.83／196.86／819.94", "百萬美元", "同上", L6 + " 計算!I88、I155、I222", "Derived", "In m32_*", "一致"),
           ("相機模組每台用量（人形／Robotaxi 每車）", "6", "人形 4 顆；Robotaxi 每車 11 顆（區間 8–13）", "顆", "輸入頁", L6 + " 輸入!F61、L61、L39、L83", "Assumed（線 6）", "經線 6 市場進子列；本模型不另設用量", "一致（經線 6）"),
           ("相機模組單價", "6", "25（15–40）", "美元/顆", "線 6 假設", L6 + " 輸入!E113", "Assumed（線 6）", "經線 6 市場；舜宇實際單價未取得", "一致（經線 6）；舜宇手機與車載模組均價未揭露"),
           ("全部輕型車相機市場 2026", "6", "6,750（既有需求，線 6 不計入實體 AI 合計）", "百萬美元", "既有車用相機", L6 + " 計算!C156", "Derived", "不進本模型；車載分部以公司實績為基礎成長", "自設：車載分部不以線 6 全車相機市場推導（份額未取得），改以歷史成長與區間"),
           ("舜宇份額", "6", "未取得", "%", "份額頁列舜宇為相機模組主要廠", L6 + " 份額!B21、F21", "—", "shr_*（Assumed 3–15%）", "一致（線模型亦未取得）"),
           ("人形出貨量與滲透", "4", "未直接取用", "台", "線 6 相機市場已含人形出貨 × 每台 4 顆", L4 + "（下游取數）", "Derived", "經線 6", "一致（經線 6）；人形約占線 6 相機市場一半（2030 基準，取自公司檔案引用，未逐格複核）")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("手機出貨量續降與手機模組毛利率下探（上半年 12.3%）", "是（g*hs_C、gm_hs 低案）", "情境 C、Sens", "全球手機出貨、公司月度出貨公告、手機分部毛利率"),
             ("車載價格競爭與規格組合使毛利率下滑（上半年 30.3%，FY2024 35.3%）", "是（gm_veh、gmadj_C）", "情境 C", "車載分部毛利率與營收年增"),
             ("其他產品（機器人、智慧影像）上半年 +88.5% 能否持續", "是（g1oth_*）", "FY2026 以上半年為錨，之後給 0–25% 區間", "其他產品營收與毛利率（34.4%）、大客戶集中度"),
             ("客戶集中（最大客戶約 11.3%、前五約 45.1%；profile 引 2025 年報）", "否", "客戶未具名，無法建模", "年報客戶集中度"),
             ("聯營公司與換股（Goertek Optical）會計與非支配股東比重變動", "部分（nci_pct）", "上半年非支配占淨利 7.3%", "半年報非支配股東損益"),
             ("稅率優惠到期或海外比重上升", "部分（tax_w 高案 15%）", "有效稅率 11–12%", "有效稅率"),
             ("人民幣與港元匯率", "否", "每股價值以 2026-11-20 匯率換算，未逐年建模", "HKD／CNY"),
             ("理財資產（13,327）的流動性與減損", "部分（fv_k 低案 70%）", "淨現金占估值比重高", "公允價值損益、到期結構")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([H126, FY25, FY24, QTR, SA, SA3008, SA3406, SA011, YCL, CNRF, L6 + "（計算、份額、輸入頁）", L4 + "（下游取數；經線 6）", GKX + "（In 頁 rf、crp）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"舜宇光學（2382.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（財報為 2026 中期業績，2026-08-26）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：舜宇光學的股價由手機模組的低毛利底部（上半年 12.3%）能否被車載（毛利率 30%）與其他產品（機器人、智慧影像）的成長抵銷決定；機器人與 Robotaxi 相機模組（線 6）目前規模極小，是選擇權而非承重項"
    RD["A4"] = (f'="在 FY2026 營收情境 B 年增 "&TEXT(PL!F{PR["B"]["gr"]},"0.0%")&"（公司指引不低於 7%）、FY2027–28 車載年增 "&TEXT(g1veh_B,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬人民幣）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：手機分部毛利率連續兩個半年低於 10%；車載分部營收年增低於 5%；其他產品營收年增低於 20%；全年營收年增低於公司不低於 7% 的指引"
    RD["A7"] = "範圍：只讀線 6（相機模組市場）與線 4（經線 6）；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：車載與其他產品成長能否抵銷手機低毛利並支持高於現價的目標價？答：情境 B FY2028 EPS "&TEXT(PL!H{PR["B"]["eps"]},"0.00")&" 人民幣（"&TEXT(PL!H{PR["B"]["eps"]}*fx_hk,"0.00")&" 港元），現價只有其 "&TEXT(px/(PL!H{PR["B"]["eps"]}*fx_hk),"0.0")&" 倍，同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元。問：實體 AI 有多重要？答：情境 B FY2030 機器人與 Robotaxi 相機模組子列只占合併營收 "&TEXT(Val!B{r_hum},"0.00%")&"，結論不靠它（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR
# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_sunny")
    os.makedirs(d, exist_ok=True)
    return "file://" + d

def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 40):
        chunk = srcs[i:i + 40]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={lo_profile()}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "soffice"]); continue
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

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
    """ovs：[(tag, ov)] → {tag: vals}；一次 LibreOffice 批次重算。"""
    d = tempfile.mkdtemp(prefix="sunny_"); paths = []; VR = DR = None
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = {tag: read_vals(o, VR, DR if tag == "base" else None) for (tag, _), o in zip(ovs, outs)}
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [  # (鍵, 驅動白話)
    ("g26hs_B", "手機營收年增 FY2026（B）"), ("g26veh_B", "車載營收年增 FY2026（B）"), ("g26xr_B", "XR 營收年增 FY2026（B）"), ("g26oth_B", "其他產品營收年增 FY2026（B）"),
    ("g1hs_B", "手機營收年增 FY2027–28（B）"), ("g1veh_B", "車載營收年增 FY2027–28（B）"), ("g1xr_B", "XR 營收年增 FY2027–28（B）"), ("g1oth_B", "其他產品營收年增 FY2027–28（B）"),
    ("g2hs_B", "手機營收年增 FY2029–32（B）"), ("g2veh_B", "車載營收年增 FY2029–32（B）"), ("g2xr_B", "XR 營收年增 FY2029–32（B）"), ("g2oth_B", "其他產品營收年增 FY2029–32（B）"),
    ("shr_B", "機器人與 Robotaxi 相機模組份額（B）"), ("rp27", "子列爬坡 FY2027"), ("fx_usd", "人民幣兌美元"),
    ("gm_hs", "手機分部毛利率"), ("gm_veh", "車載分部毛利率"), ("gm_xr", "XR 分部毛利率"), ("gm_oth", "其他產品毛利率"), ("gmadj_B", "毛利率調整（B）"),
    ("sga_el", "費用彈性"), ("oth_op", "其他營業損益淨額"), ("nonop", "財務費用"), ("tax_w", "有效稅率"), ("nci_pct", "非支配股東比重"), ("fin_inc", "利息與理財收益（不進 FCFF）"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("fv_k", "理財資產計入比例"), ("tr_sh", "股份獎勵信託持股"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數"), ("payout", "股利支付率")]
BETA_B = statistics.median([0.67 * 1.22 + 0.33, 0.67 * statistics.median([1.16, 1.04, 2.63]) + 0.33])
FXD = {"pe_med": statistics.median([19.9 * 3670 / 3980, 16.8 * 648 / 734, 7.75 * 267500 / 284500])}


def solve_batch(key, lo, hi, targets, px0):
    """格點批次重算（13 點）→ 對每個目標上檔找相鄰兩點線性內插；再以 7 點細格點精修一次。"""
    import numpy as np
    xs = list(np.linspace(lo, hi, 13))
    res = run_cases([(f"g{i}", {key: x}) for i, x in enumerate(xs)])
    ys = [res[f"g{i}"]["up"] for i in range(len(xs))]
    out = {}
    for tgt in targets:
        if tgt is None: continue
        br = next((i for i in range(len(xs) - 1) if (ys[i] - tgt) * (ys[i + 1] - tgt) <= 0), None)
        if br is None: out[tgt] = None; continue
        a, b_ = xs[br], xs[br + 1]
        xs2 = list(np.linspace(a, b_, 7)); r2 = run_cases([(f"f{i}", {key: x}) for i, x in enumerate(xs2)]); y2 = [r2[f"f{i}"]["up"] for i in range(7)]
        j = next(i for i in range(6) if (y2[i] - tgt) * (y2[i + 1] - tgt) <= 0)
        x0, x1, y0, y1 = xs2[j], xs2[j + 1], y2[j], y2[j + 1]
        out[tgt] = x0 if y1 == y0 else x0 + (tgt - y0) * (x1 - x0) / (y1 - y0)
    return out


def main():
    px0 = in_val("px", "base")
    jobs = [("base", {})]; meta = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = BETA_B + (-0.3 if side == "lo" else 0.3)
            if k == "pe_med": v = 7.75 * 267500 / 284500 if side == "lo" else 19.9 * 3670 / 3980
            if v is None: continue
            b_ = in_val(k, "base")
            if not isinstance(b_, (int, float)): b_ = FXD.get(k, BETA_B if k == "beta" else None)
            if b_ is not None and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v})); meta.append((tag, drv + ("（低）" if side == "lo" else "（高）"), f"{k}＝{v:.4g}", k))
    for s in SC:
        tag = f"only_{s}"; jobs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC})); meta.append((tag, f"只看情境 {s}", f"w_{s}＝1，其餘 0", f"w_{s}"))
    res = run_cases(jobs)
    base = res["base"]; print("基準", base)
    cases = [dict(drv=d, set=st, key=k, **res[t]) for t, d, st, k in meta]
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    for c in cases: print(c["key"], c["drv"], round(c["tp"]), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("g1oth_B", 0.12, -0.50, 1.00, "情境 B 其他產品 FY2027–28 年增（其他不變，整本重算；情境 A／C 不動）"),
                                    ("gm_hs", 0.13, -0.10, 0.40, "手機分部毛利率（三情境同動）"),
                                    ("pe_med", FXD["pe_med"], 0.0, 40.0, "同業前瞻 P/E 中位數（Comparable；DCF 不受影響）")):
        sol = solve_batch(key, lo_, hi_, [0.0, upper, lower], px0)
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
            if lab is None: continue
            x = sol.get(tgt)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="sunnyf_"); raw = os.path.join(tmp, NAME)
    bid = f"{VER}-{fingerprint()}"
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("Sens 狀態", wbv["Sens"]["D4"].value, "tp", wbv["Val"][f"B{VR['tp']}"].value)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    segs = [("手機", "手機產品營收", 13148.628), ("車載", "車載產品營收", 3761.928), ("XR", "XR 產品營收", 955.720), ("其他", "其他產品營收", 4039.428)]
    json.dump(dict(company="sunny", ticker="2382.HK", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="2026 上半年", value=v, unit="百萬人民幣", source=H126 + " 分部資訊") for n, l, v in segs],
                   sublines=[], check_labels=["檢查：分部營收合計 − 合併營收（應為 0）"],
                   note="官方報告分部為手機、車載、XR、其他四類（2025 年報起；FY2024 與 2025H1 為公司重編值）。PL 營收列直接用四分部；另有「機器人與 Robotaxi 相機模組」命題子列（線 6 市場 × 份額，FY2026 以前為 0）與檢查列。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)



def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=range(2023, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.0f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2023, 2033)
    txt = f"""# physical_ai/sunny — 舜宇光學（2382.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

港交所主板上市；報表人民幣、股價港元（以收盤檔 2026-11-20 的 HKD=X 與 CNY=X 換算，1 人民幣＝{inv('fx_hk'):.4f} 港元）。本 README 的數字由 `build_sunny.py` 從重算後的成品讀出產生，不手寫。

**本版沒有做的部分（高不確定性原則：證據到哪裡判斷到哪裡）：** 沒有出貨量與均價拆分（公司月度出貨公告的數字未取得，營收不拆量價）；FY2023 只有營收、毛利、稅前與稅後（舊分部不可比，費用拆分與歸屬母公司淨利未取得）；全年指引只取得資料商摘要的「營收與淨利（不含換股收益）年增不低於 7%」，未取得公告原文；同業前瞻 P/E 只有 3 家（豪威 603501 統計頁讀取失敗、歌爾與丘鈦收盤檔沒有，列「報價日未取得」不進中位數）；機器人與 Robotaxi 相機模組份額沒有證據，用區間；股份獎勵信託持股、僱員股權激勵、聯營投資公允價值未取得。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 現價 {inv('px'):.2f} 港元只有模型情境 B FY2027 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍，資料商現行前瞻 P/E 約 {inv('pe_self'):.1f} 倍、同業 3 家中位數 {inv('pe_med'):.1f} 倍（Analogy，家數少）；DCF 另有淨現金 約 {wb['DCF']['B10'].value:,.0f} 百萬人民幣（含理財資產 13,327，約占市值四分之一）。結論主要建立在 Assumed 的分部成長與毛利率、Analogy 的同業倍數，以及理財資產可全額變現的假設上；機器人與 Robotaxi 相機模組子列（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.2f}%）不是承重項。

**證據邊界：** 上檔主要來自 DCF（單獨 {f2(dcf)} 港元，{pc(dcf/inv('px')-1)}），而 DCF 的折現率偏低（rf 1.61%、β {inv('beta'):.2f}、WACC {g(f"B{VR['wk']}").split('／')[0]}）；ERP 取 6% 時加權目標價 {f2(next(c for c in sens['cases'] if c['key']=='erp' and c['set'].endswith(str(inv('erp')+0.01)) or c['key']=='erp' and c['drv'].endswith('（高）'))['tp'])} 港元。只看情境 C：{f2(next(c for c in sens['cases'] if c['key']=='w_C')['tp'])} 港元（{pc(next(c for c in sens['cases'] if c['key']=='w_C')['tp']/inv('px')-1)}，{next(c for c in sens['cases'] if c['key']=='w_C')['rt']}）。同業 P/E 取低端（LG Innotek）：{f2(next(c for c in sens['cases'] if c['key']=='pe_med' and c['drv'].endswith('（低）'))['tp'])} 港元（{next(c for c in sens['cases'] if c['key']=='pe_med' and c['drv'].endswith('（低）'))['rt']}）。基準評等「強力買進」在情境 C 與同業低端下不成立。

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，百萬人民幣） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 港元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方平均目標價 {f2(inv('cons_tp'))} 港元（21 位，2026-11-24，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 3 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 EPS × 匯率）。相對倍數係數 1.00（未調整）。FCFF 不含利息與理財收益（已以淨現金計價，避免重複）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS 港元）{g(f"B{VR['pe']}"):.1f} 倍；舜宇現行 {inv('pe_self'):.2f} 倍＝同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，百萬人民幣；FY2023–25 實績、FY2026 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 手機產品 | {yr('B','c_hs')} |
| 車載產品 | {yr('B','c_veh')} |
| XR 產品 | {yr('B','c_xr')} |
| 其他產品 | {yr('B','c_oth')} |
| 機器人與 Robotaxi 相機模組（命題子列） | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益（本模型口徑） | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（人民幣元） | {yr('B','eps','{:,.2f}')} |

- FY2023 為舊分部（光學零件、光電產品、光學儀器），不可比，故分部列空白。FY2025 營業利益含換股收益 918.9（一次性），FY2025 淨利年增因此不可直接外推；FY2026 淨利成長用不含換股的 FY2025 對照（見 Recon）。
- 營業利益為本模型口徑（毛利 − 研發、銷售、行政 ＋ 其他營業損益）；公司不揭露營業利益。其他營業損益含利息與理財收益（約 560／年），FCFF 已扣除。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 分部營收年增（情境 B）：FY2026 手機 −1%、車載 +11%、XR +8%、其他 +45%（上半年實績 −0.8%、+10.6%、+11.0%、+88.5%）；FY2027–28 為 0%、10%、12%、12%；FY2029–32 為 0%、8%、10%、8%。全部 Assumed，A／C 另設。
- 分部毛利率（B，下半年起）：手機 13%、車載 30.5%、XR 13.5%、其他 32%；合併毛利率由此推導。費用彈性 0.2；稅率 11.5%；非支配股東 6%；其他營業損益 1,100、財務費用 −330。
- 子列：線 6 相機模組市場（C／B／A＝低／基準／高）× 份額 3%／8%／15% × 爬坡 × {inv('fx_usd'):.4f}。
- 折現：rf 1.61%、β {inv('beta'):.3f}、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。股數 {inv('sh_v'):,.1f} 百萬股；淨現金與投資 {wb['DCF']['B10'].value:,.0f} 百萬人民幣（2026-06-30）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 車載與其他產品成長（g1veh、g1oth） | 2026 全年業績（約 2027-03）、月度出貨公告、其他產品營收年增 |
| 手機與車載分部毛利率（gm_hs、gm_veh） | 2026 全年分部毛利率（上半年 12.3%、30.3%） |
| 理財資產計入比例（fv_k） | 公允價值損益、理財到期結構 |
| Comparable 倍數（pe_med，Analogy） | 同業前瞻 P/E；舜宇自身倍數是否向同業收斂 |

失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 6：機器人與 Robotaxi 相機模組市場 2026／2030／2032 低／基準／高（計算!C/G/I 的 88、155、222 列）→ m26／m30／m32 的情境 C／B／A，照取（「下游取數」頁無相機模組列，故讀「計算」頁）。份額（份額!F21）線模型未取得；每台用量與單價為線 6 假設，經線 6 市場進子列，本模型不另設。
- 線 6 的「全部輕型車相機 2026 約 6,750 百萬美元」不進本模型（既有需求；舜宇份額未取得），車載分部改以公司實績與成長區間建模——與線模型的差異，已在對帳頁標明。
- 線 4（人形）未直接取用：人形相機需求已含在線 6 相機市場。

## 限制與未取得資料
- 出貨量與均價（公司月度公告）、FY2023 費用拆分與歸屬母公司淨利、車載與手機分部以外的客戶資訊、全年指引原文：未取得。
- 同業前瞻 P/E 只有大立光、玉晶光、LG Innotek 3 家（已換算到 2026-11-20）；豪威（上游 CIS）報價日未取得；歌爾、丘鈦收盤檔沒有，不納入。3 家 P/E 差距大（約 7–18 倍），中位數不穩。
- 資料商頁面價格為 2026-11-24（較 11-20 高），已以收盤檔換算；賣方共識目標價為 2026-11-24 頁面。
- 2026-11-16 至 11-20 全球 AI 與半導體股大跌屬實，收盤檔的下跌不是資料錯誤；舜宇單週下跌原因未逐項查核。
- rf、CRP 沿用 geekplus 模型取值（逾期，跨模型一致）。DCF 為本模型自寫，公式與 `docs/plan/個股DCF標準規格.md` 相同。

## 檔案
- `build_sunny.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))} 港元（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()
