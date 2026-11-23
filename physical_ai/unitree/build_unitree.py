# 宇樹科技 Unitree（688836.SS，上交所科創板）個股模型 v0.1 — AI實體應用 Project（線 4 人形機器人整機；線模型 physical_ai/lines/humanoid/）
# 用法：python3 physical_ai/unitree/build_unitree.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
#       python3 physical_ai/unitree/build_unitree.py --raw OUT  → 只寫一本（讀環境變數 UNI_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度＝曆年（12 月決算，中國企業會計準則）。金額：億元人民幣；股數：百萬股；EPS、股價：元；單價：萬元／台。
# 結構：README、In、Line（線 2b 下游取數讀值）、Hist、Q、PL（三情境：命題模組→官方產品分部→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 本題適用《高不確定性預估工作原則》（docs/plan/高不確定性預估工作原則.md）：上市僅約三個月、2026 年分產品營收未揭露，未揭露的輸入一律給區間並標 Assumed／Analogy。
# DCF 為本檔自寫（未用 tools/fcff_dcf.py），公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
from concurrent.futures import ThreadPoolExecutor
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_宇樹個股模型_{VER}.xlsx"
LINE_X = os.path.join(ROOT, "physical_ai", "lines", "humanoid", "dist", "20261116_人形機器人單位經濟_v0.1.xlsx")
LINE_REL = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
CSJ = "中國證券報 2026-06-02 A04 版（上市委審議通過報導，引招股書）https://epaper.cs.com.cn/zgzqb/images/2026-06/02/A04/zqDB1102.pdf"
CSJ8 = "中國證券報 2026-08-06 A06 版（招股意向書與發行安排）https://epaper.cs.com.cn/zgzqb/images/2026-08/06/A06/zqDB1106.pdf"
HUX = "虎嗅 2026-10-23（AD財經，引上交所招股書：分產品營收、銷量、單價、毛利率、費用、存貨）https://www.huxiu.com/article/4900054.html"
TMT = "鈦媒體 2026-08-19（錦緞，引招股書）https://www.tmtpost.com/8108156.html"
NBD = "每日經濟新聞 2026-03-20（招股書申報稿）https://www.nbd.com.cn/articles/2026-03-20/4301964.html"
EM3 = "東方財富 2026-03-20（招股書申報稿：2025 年前三季營收 11.67 億、研發費用率）https://finance.eastmoney.com/a/202603203679500125.html"
SINA = "新浪財經 2026-08-20（上市首日：發行價、發行股數、募資額）https://finance.sina.com.cn/roll/2026-08-20/doc-ininxkrz8880098.shtml"
EMH1 = "東方財富 2026-08-18（引上市公告書：2026 上半年營收、淨利、扣非、總資產）https://finance.eastmoney.com/a/202608183844373066.html"
EMQ3 = "東方財富 2026-10-28（2026 年第三季報：單季與前三季營收、淨利、扣非、營業現金流、存貨、貨幣資金、總資產、淨資產）https://finance.eastmoney.com/a/202610283912365052.html"
IFQ3 = "鳳凰網財經 2026-10-28（第三季報：前三季研發費用 2.63 億、第三季 1.27 億）https://finance.ifeng.com/c/8xc5qdaZmEa"
SSQ3 = "證券之星 2026-10-28（第三季報：第三季研發投入占營收 19.10%）https://wap.stockstar.com/detail/IG2026102800044398"
ETN = "etnet 經濟通 滬股 688836 報價頁（2026-11-23 16:30 讀取：「前收市」590.000、當日 708.000）https://etnet.com.hk/www/sc/ashares/quote.php?code=688836"
SAH = "stockanalysis.com 688836 歷史股價頁（讀取 2026-11-23，資料只到 2026-11-13：11-10 761.01、11-13 696.00）https://stockanalysis.com/quote/sha/688836/history/"
SA = "stockanalysis.com statistics 頁（資料商；頁面日期見各列）"
CNY10 = "中國 10 年期國債殖利率 2026-11-05 1.61%（tradingeconomics；與 AI半導體 勝宏／中際旭創／新易盛 模型同一取值，跨模型一致）"
YCD = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整；EV、營收為資料商計算）"

# ---------------- 線 2b 下游取數（建置時從線模型讀值，不手打） ----------------
def read_line():
    wb = openpyxl.load_workbook(LINE_X, data_only=True); ws = wb["下游取數"]
    out = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str) and (k.startswith("SHIP_") or k.startswith("HUM_PRICE_")):
            out[k] = dict(row=r, lo=ws.cell(r, 3).value, base=ws.cell(r, 4).value, hi=ws.cell(r, 5).value, name=ws.cell(r, 2).value)
    return out
LN = read_line()

# ---------------- 同業 EV／營收（資料商頁面值；以頁面股價與 2026-11-20 收盤等比例調整市值，Derived） ----------------
# (鍵, 名稱, 頁面日期, 頁面股價, 頁面市值, 頁面 EV, TTM 營收, 2026-11-20 收盤, 幣別, 業務)
PEERS = [("ev_ubt", "優必選 9880.HK（人形整機、服務機器人）", "2026-10-16", 64.55, 32.44, 31.62, 3.06, 57.900002, "HKD", "人形整機"),
         ("ev_dobot", "越疆 2432.HK（協作機器人、人形）", "2026-09-21", 21.10, 9.28, 6.97, 0.75669, 14.4, "HKD", "機器人整機"),
         ("ev_geek", "極智嘉 2590.HK（倉儲 AMR）", "2026-09-04", 9.22, 12.30, 9.85, 3.96, 5.405, "HKD", "機器人整機"),
         ("ev_ldr", "綠的諧波 688017.SS（諧波減速機）", "2026-10-09", 263.76, 48.36, 46.19, 0.66785, 230.800003, "CNY", "零組件"),
         ("ev_orb", "奧比中光 688322.SS（3D 視覺感測）", "2026-09-18", 92.38, 37.80, 35.86, 0.94288, 85.089996, "CNY", "零組件"),
         ("ev_rbs", "速騰聚創 2498.HK（光達）", "2026-10-23", 14.10, 6.64, 4.37, 2.51, 13.32, "HKD", "零組件")]
def peer_evs(p):
    _, _, _, pxp, mc, ev, rv, px1, _, _ = p
    return (ev + mc * (px1 / pxp - 1)) / rv
PEER_V = [peer_evs(p) for p in PEERS]
_q = statistics.quantiles(PEER_V, n=4, method="inclusive")
EV_LO, EV_MED, EV_HI = _q[0], statistics.median(PEER_V), _q[2]

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點：前三季實績（2026-10-28 第三季報）與第四季假設")
inp("q9_rev", "2026 年前三季營收", "億元", None, 18.17, None, "Verified", EMQ3 + "：18.17 億（+55.71%）", "上半年 11.52 億（" + EMH1 + "）；第一季 4.23 億（" + CSJ8 + "）")
inp("q3_rev", "2026 年第三季營收", "億元", None, 6.65, None, "Verified", EMQ3 + "：6.65 億（+69.93%）", "＝18.17 − 11.52（一致）")
inp("q9_ni", "2026 年前三季歸屬母公司淨利", "億元", None, 3.87, None, "Verified", EMQ3 + "：3.87 億（+270.54%；上年同期含 3.49 億股份支付）", "前三季扣非淨利 3.49 億（−18.78%）；第三季淨利 1.13 億")
inp("q9_rd", "2026 年前三季研發費用", "億元", None, 2.63, None, "Verified", IFQ3 + "：2.63 億（+191.48%）", "第三季 1.27 億，占營收 19.10%（" + SSQ3 + "）")
inp("q4k", "第四季營收 ÷ 第三季營收（2026）", "x", 1.0, 1.20, 1.36, "Analogy", "類比 2025：第四季 5.32（16.99 − 前三季 11.67）÷ 第三季 3.91（11.67 − 上半年 7.76）＝ 1.36（Derived，" + EM3 + "；" + EMH1 + "）", "低 1.0＝第四季持平第三季；公司未給全年指引（未找到）")
inp("hs26", "FY2026 人形機器人營收占合併營收", "%", 0.50, 0.55, 0.62, "Assumed", "未揭露：2026 年分產品營收未公布（上半年公告只寫人形「保持相對較快增長」，" + EMH1 + "）", "FY2025 51.09%（" + TMT + "）；FY2024 27.2%（Derived）")
inp("cs26", "FY2026 組件及其他營收占合併營收", "%", 0.06, "=h_crev_25/h_rev_25", 0.10, "Derived", "＝FY2025 組件及其他 1.33 ÷ 合併 16.99（持平；Derived）", "FY2024 14.1%、FY2023 23.3%（Derived）；低／高 Assumed")
inp("dasp_h26", "FY2026 人形平均售價年變化", "%", -0.30, -0.15, -0.05, "Assumed", "FY2023 59.34 → FY2024 26.04 → FY2025 16.64 萬元（" + HUX + "）：年降 56%、36%（產品組合轉向 G1）", "2026 起 R1 等更低價機型放量（定性），幅度未揭露")
inp("dasp_q26", "FY2026 四足平均售價年變化", "%", -0.10, -0.05, 0.0, "Assumed", "FY2023 3.83 → FY2024 3.23 → FY2025 3.03 萬元（" + HUX + "）：年降 16%、6%", "")
inp("dwc26", "FY2026 營運資金增加（以存貨增加代）", "億元", 4.0, 6.51, 8.0, "Derived", "存貨 2026-09-30 10.19 億 − 2025 年底 3.68 億 ＝ 6.51（" + EMQ3 + "；" + HUX + "）", "應收、應付未取得；第四季變化未知（低／高 Assumed）")
inp("nonop26", "FY2026 業外（利息收入、政府補助等）", "億元", 0.2, 0.35, 0.5, "Assumed", "IPO 募資淨額約 61 億於 2026-08 到帳（" + SINA + "）；前三季非經常性損益約 0.38（＝淨利 3.87 − 扣非 3.49，Derived）", "FY2026 營業利益以前三季淨利率校準後扣除本列（見 PL 頁）")
inp("cx_26", "FY2026 資本支出", "億元", 0.5, 1.0, 2.0, "Assumed", "FY2025 購建長期資產付現 0.50 億（" + TMT + "）；募資項目含製造基地 6.24 億（" + CSJ + "）", "")

sec("B. 命題模組：人形機器人＝全球人形出貨（線 2b）× 宇樹份額 × 平均售價（萬元）")
for s, v, lo, hi, note in (("A", 3, None, None, "樂觀：讀線 2b「高」欄（全球出貨高、學習曲線售價降幅大）"),
                           ("B", 2, 1, 3, "基準：讀線 2b「基準」欄；低＝讀「低」欄、高＝讀「高」欄（只動情境 B）"),
                           ("C", 1, None, None, "保守：讀線 2b「低」欄")):
    inp(f"lnsel_{s}", f"情境 {s}：讀取線 2b 欄位（1＝低、2＝基準、3＝高）", "選擇", lo, v, hi, "Assumed", note, "線模型 " + LINE_REL + " 下游取數頁 C／D／E 欄（Line 頁列出讀入值）")
for s, v, lo, hi in (("A", 0.15, None, None), ("B", 0.10, 0.06, 0.15), ("C", 0.06, None, None)):
    inp(f"shr32_{s}", f"情境 {s}：宇樹人形出貨占全球 FY2032（FY2026 起線性收斂）", "%", lo, v, hi, "Assumed",
        "FY2025 宇樹 5,215 台 ÷ 全球約 1.3 萬（Omdia）／1.8 萬（IDC）＝ 29%–40%（Derived；全球值經線 2b 追蹤指標頁，Interested-party）",
        "FY2026 份額由營收推得（Derived，見 PL 頁）；2026 上半年中國廠出貨智元居首、宇樹次之（同花順／中國基金報 2026-08-15，未查核數量）")
inp("asp_k", "宇樹人形售價年變化 ÷ 線 2b 全尺寸售價學習曲線年變化", "x", 0.5, 1.0, 1.5, "Analogy",
    "類比對象：線 2b 全尺寸（Optimus Gen 2 類）出廠價學習曲線；宇樹以小型機為主（G1、R1），口徑不同，只借用降幅", "FY2024→FY2025 宇樹人形 ASP −36%（含產品組合）")
inp("hgm26", "人形機器人毛利率 FY2026", "%", 0.55, 0.60, 0.63, "Assumed", "FY2023 87.67%、FY2024 69.26%、FY2025 63.18%（" + HUX + "）", "2026 年分產品毛利率未揭露（未找到）")
for s, v, lo, hi in (("A", 0.50, None, None), ("B", 0.42, 0.32, 0.50), ("C", 0.32, None, None)):
    inp(f"hgm32_{s}", f"情境 {s}：人形機器人毛利率 FY2032（FY2026 起線性收斂）", "%", lo, v, hi, "Assumed",
        "線 2b 整機毛利率基準 25%（15–45%，全尺寸）；宇樹小型機科研市場毛利高於整機均值", "優必選 2026H1 毛利率 44.7%（線 2b 追蹤指標頁，分部口徑待核）")

sec("C. 四足機器人、組件及其他（官方產品分部；量 × 價或成長路徑）")
for s, v, lo, hi in (("A", 0.30, None, None), ("B", 0.20, 0.10, 0.30), ("C", 0.08, None, None)):
    inp(f"qg_{s}", f"情境 {s}：四足出貨年增 FY2027–FY2029", "%/年", lo, v, hi, "Assumed", "FY2023 3,121 → FY2024 7,136 → FY2025 23,037 台（" + HUX + "）", "海外營收占四足營收逾 40%（田豐，經" + EMQ3 + "，Interested-party）")
for s, v, lo, hi in (("A", 0.15, None, None), ("B", 0.10, 0.04, 0.15), ("C", 0.03, None, None)):
    inp(f"qlt_{s}", f"情境 {s}：四足出貨年增 FY2030–FY2032", "%/年", lo, v, hi, "Assumed", "成熟期假設；無公開長期預測（未找到）", "")
inp("qasp", "四足平均售價年變化 FY2027 起", "%/年", -0.10, -0.05, 0.0, "Assumed", "FY2023–FY2025 年降 16%、6%（" + HUX + "）", "")
inp("qgm26", "四足機器人毛利率 FY2026", "%", 0.50, 0.55, 0.58, "Assumed", "FY2023 43.71%、FY2024 51.65%、FY2025 56.72%（" + HUX + "）", "")
for s, v, lo, hi in (("A", 0.56, None, None), ("B", 0.52, 0.46, 0.56), ("C", 0.46, None, None)):
    inp(f"qgm_{s}", f"情境 {s}：四足機器人毛利率 FY2027 起", "%", lo, v, hi, "Assumed", "FY2025 56.72%；消費四足價格競爭（定性）", "")
for s, v, lo, hi in (("A", 0.35, None, None), ("B", 0.25, 0.10, 0.35), ("C", 0.10, None, None)):
    inp(f"cg_{s}", f"情境 {s}：組件及其他營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "FY2023 0.37 → FY2024 0.55 → FY2025 1.33 億（Derived：合併 − 人形 − 四足）", "組件含關節馬達、光達等對外銷售（FY2025 組件 1.04 億、其他 0.23 億，" + HUX + "）")
inp("cgm", "組件及其他毛利率", "%", 0.50, 0.55, 0.60, "Assumed", "FY2025 組件毛利率 59.36%（" + HUX + "）", "")

sec("D. 營業費用與損益表下半段")
inp("rd_r", "研發費用率 FY2027 起", "%", 0.10, 0.13, 0.18, "Assumed", "2026 前三季 14.5%（Derived：2.63 ÷ 18.17）、第三季 19.10%（" + SSQ3 + "）", "FY2025 8.5%、FY2024 17.8%、FY2023 31.4%（Derived，" + CSJ + "）；募資 20.22 億投大模型研發")
inp("sm_r", "銷售費用率 FY2027 起", "%", 0.07, 0.10, 0.13, "Assumed", "FY2025 8.31%（同業平均 23.88%，" + TMT + "）；2026 第一季銷售費用 1.4 億（+133%，含春晚推廣）", "FY2026 由前三季淨利率校準推得（PL 頁）")
inp("ga_r", "管理費用率（不含股份支付）FY2026 起", "%", 0.03, 0.04, 0.06, "Assumed", "FY2025 管理費用 4.00 億 − 股份支付 3.49 億 ＝ 0.51 億，占營收 3.0%（Derived，" + HUX + "；" + TMT + "）", "")
inp("sbc_r", "股份支付費用率 FY2026 起", "%", 0.0, 0.01, 0.03, "Assumed", "FY2025 一次性股份支付 3.49 億（列非經常性，" + NBD + "）；上市後新激勵計畫未揭露（未找到）", "NOPAT 扣除股份支付（不加回）")
inp("tax_w", "有效稅率", "%", 0.12, 0.15, 0.25, "Analogy", "類比：中國高新技術企業優惠稅率 15%（宇樹是否適用、實際有效稅率未取得）", "高 25%＝法定稅率")
inp("int_r", "業外收益率（利息收入與補助 ÷ 2026-09-30 貨幣資金）FY2027 起", "%", 0.008, 0.012, 0.016, "Assumed", "人民幣定存與理財約 1–1.6%（類比；公司理財配置未揭露）", "以 cash_v 不變計（未計現金累積）")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.01, 0.02, 0.04, "Assumed", "2025 年底固定資產僅 0.35 億（" + CSJ + "）；製造基地投產後上升", "")
inp("fc_cx", "資本支出 ÷ 營收 FY2027 起", "%", 0.02, 0.05, 0.08, "Assumed", "募資項目：製造基地 6.24 億、本體研發 11.10 億（" + CSJ + "）", "FY2025 付現 0.50 億（0.3%）")
inp("fc_wc", "營運資金增加 ÷ 營收增量 FY2027 起", "%", 0.20, 0.30, 0.50, "Assumed", "FY2025 存貨增加 2.27 ÷ 營收增量 13.06 ＝ 17%；2026 前三季存貨增加 6.51（" + EMQ3 + "）", "應收與應付未取得")
inp("pay_r", "配息率（股利 ÷ 淨利）FY2027 起", "%", 0.0, 0.0, 0.2, "Assumed", "上市前未配息；股利政策未揭露（未找到）", "只用於收支關卡；不進 DCF")

sec("E. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤；單一來源）", "元", None, 590.0, None, "Verified", ETN, "單一來源（待交易所資料複核）；Yahoo 日線檔無本檔。另見 " + SAH + "；2026-11-23 漲停收 708.00（同 etnet 頁，不進模型）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "總股本（發行後）", "百萬股", None, "=40.4464/0.1", None, "Derived", "發行 4,044.64 萬股、占發行後總股本 10%（" + CSJ8 + "；" + SINA + "）", "員工持股平台上海宇翼持股已發行在外，無選擇權揭露 → 評價日股數不另加")
inp("sh_pre", "發行前總股本", "百萬股", None, "=sh_out-40.4464", None, "Derived", "＝總股本 − 新發行股數", "FY2025 EPS 分母（Derived）")
inp("sh_26", "FY2026 加權平均股數", "百萬股", None, "=sh_pre+40.4464*(DATE(2026,12,31)-DATE(2026,8,19))/365", None, "Derived", "2026-08-19 上市（" + SINA + "）", "")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "DCF 標準規格第 3 節：無 RSU／選擇權揭露", "")
inp("cash_v", "貨幣資金（2026-09-30）", "億元", None, 73.54, None, "Verified", EMQ3 + "：73.54 億", "2025 年底 14.19 億（" + CSJ + "）")
inp("debt_v", "有息負債（2026-09-30）", "億元", None, 0.0, None, "Assumed", "2025 年底無短期借款（" + CSJ + "）；2026-09-30 借款未取得", "以 0 計")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", CNY10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：中國 0.91%（與 AI半導體 中國個股同值）", "")
inp("beta", "β", "x", 1.0, 1.3, 1.6, "Analogy", "上市僅約三個月，自身 β 無統計意義（未計算）；類比高成長機器人／硬體股 Blume 調整後約 1.2–1.4", "未取得兩來源 Blume 中位數（違反標準規格第 2 節，列為限制）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（無有息負債，權重為 0）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "元", None, 0.0, None, "Assumed", "未配息", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺宇樹對同業的歷史相對倍數；上市僅約三個月）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("F. 同業 EV／營收（TTM；資料商頁面值，市值以頁面股價 → 2026-11-20 收盤等比例調整，Derived）")
for p, v in zip(PEERS, PEER_V):
    k, lab, d, pxp, mc, ev, rv, px1, cur, grp = p
    inp(k, f"EV／營收 TTM：{lab}", "x", None, f"=({ev}+{mc}*({px1}/{pxp}-1))/{rv}", None, "Derived",
        f"{SA}：{d} 頁面 EV {ev}、市值 {mc}、TTM 營收 {rv}（十億 {cur}）、股價 {pxp}；× 2026-11-20 收盤 {px1:g}（{YCD}）",
        f"群組 {grp}；淨現金不變（Derived）；TTM 營收為頁面日之前最近一期，未更新到同一估值日（限制）", DV)
inp("ev_med", "同業 EV／營收 TTM 中位數（6 家）", "x", round(EV_LO, 4), "=MEDIAN(" + ",".join(p[0] for p in PEERS) + ")", round(EV_HI, 4), "Derived",
    "少於 8 家，不做迴歸（第 10 節第 7 條）；低／高＝6 家第 1／第 3 四分位數", "離散大（約 1–60 倍）：獲利同業（綠的諧波、奧比中光）倍數高，虧損整機同業低")
inp("evs_self", "宇樹現行 EV／營收 TTM（只作對照，不進模型）", "x", None, "=(px*sh_out/100-cash_v)/(h_rev_25+q9_rev-11.67)", None, "Derived",
    "＝（2026-11-20 市值 − 貨幣資金）÷ TTM 營收（FY2025 16.99 ＋ 2026 前三季 18.17 − 2025 前三季 11.67）", "2025 前三季營收 11.67 為申報稿審閱數（" + EM3 + "）")
inp("cons", "市場共識 EPS、目標價（只作對照，不進模型）", "—", None, "未找到", None, "Assumed", "本輪搜尋未取得賣方共識（未找到）", "")


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額億元人民幣、股數百萬股、EPS 與股價元、單價萬元／台"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未找到」"
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
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("lnsel", "讀取線 2b 欄位（1 低、2 基準、3 高）"), ("shr32", "宇樹人形份額 FY2032"), ("hgm32", "人形毛利率 FY2032"), ("qg", "四足出貨年增 FY2027–FY2029"),
           ("qlt", "四足出貨年增 FY2030–FY2032"), ("qgm", "四足毛利率 FY2027 起"), ("cg", "組件及其他年增 FY2027 起")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 52; IN.column_dimensions["H"].width = 60

    # ---------------- Line（線 2b 下游取數） ----------------
    L = wb.create_sheet("Line")
    L["A1"] = f"Line — 線 2b 人形機器人單位經濟（{LINE_REL}）下游取數頁的讀入值（建置時讀取，常數；不在本模型修改口徑）"
    hdr(L, 2, ["代號", "名稱", "低", "基準", "高", "單位", "線模型儲存格"])
    LR = {}
    for i, k in enumerate([f"SHIP_{y}" for y in range(2026, 2033)] + ["HUM_PRICE_2026", "HUM_PRICE_2030", "HUM_PRICE_2035"], 3):
        d = LN[k]
        L.cell(i, 1, k); L.cell(i, 2, d["name"])
        for j, kk in enumerate(("lo", "base", "hi")):
            v = d[kk] if d[kk] is not None else d["base"]
            L.cell(i, 3 + j, v).font = BLUE
        L.cell(i, 6, "台/年" if k.startswith("SHIP") else "USD/台")
        L.cell(i, 7, f"下游取數!C{d['row']}:E{d['row']}" + ("（只有基準值，三欄同值）" if d["lo"] is None else ""))
        LR[k] = i
    L.column_dimensions["B"].width = 30; L.column_dimensions["G"].width = 40
    def lsel(k, s): return f"CHOOSE(lnsel_{s},Line!$C${LR[k]},Line!$D${LR[k]},Line!$E${LR[k]})"

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；億元人民幣，單價萬元／台）。公式只引用本頁儲存格（定義名稱 h_*_年）。招股書（上交所）經媒體轉述；原文 PDF 未逐頁核對"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營業收入（合併）", 1.59, 3.93, 16.99, "Verified", CSJ + "；FY2025 亦見 " + SINA),
          ("h_hrev", "人形機器人營收", 0.0296671, 1.07, 8.68, "Verified", CSJ + "（FY2023 296.71 萬元）"),
          ("h_qrev", "四足機器人營收", 1.19, 2.31, 6.98, "Verified", CSJ),
          ("h_hu", "人形機器人銷量（台）", 5, 412, 5215, "Verified", HUX + "；FY2025 另有「出貨逾 5,500 台」之說（含未確認收入者，" + NBD + "）"),
          ("h_hasp", "人形機器人平均售價（萬元）", 59.34, 26.04, 16.64, "Verified", HUX),
          ("h_qu", "四足機器人銷量（台）", 3121, 7136, 23037, "Verified", HUX),
          ("h_qasp", "四足機器人平均售價（萬元）", 3.83, 3.23, 3.03, "Verified", HUX),
          ("h_hgm", "人形機器人毛利率", 0.8767, 0.6926, 0.6318, "Verified", HUX),
          ("h_qgm", "四足機器人毛利率", 0.4371, 0.5165, 0.5672, "Verified", HUX),
          ("h_mgm", "主營業務毛利率", 0.4422, 0.5674, 0.6013, "Verified", HUX + "（FY2025 綜合毛利率 60.44%）"),
          ("h_rd", "研發費用", 0.4995, 0.7002, 1.45, "Verified", CSJ),
          ("h_sm", "銷售費用", 0.3772, 0.5916, 1.41, "Verified", CSJ + "；FY2025 " + HUX),
          ("h_ni", "歸屬母公司淨利", -0.1115, 0.9547, 2.78, "Verified", CSJ),
          ("h_inv", "存貨（年底）", 0.7896, 1.41, 3.68, "Verified", HUX)]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            H.cell(i, 2 + j, v).font = BLUE
            dn(f"{k}_{23 + j}", f"Hist!${'BCD'[j]}${i}")
        H.cell(i, 5, tag); H.cell(i, 6, src)
    i = 3 + len(HR)
    H25 = [("h_cogs", "營業成本 FY2025", 6.72, "Verified", HUX), ("h_gp", "毛利 FY2025", 10.27, "Verified", HUX),
           ("h_ga", "管理費用 FY2025（含股份支付）", 4.00, "Verified", HUX), ("h_sbc", "股份支付 FY2025（計入管理費用，非經常性）", 3.49, "Verified", TMT + "；" + NBD),
           ("h_crev", "組件及其他營收 FY2025（合併 − 人形 − 四足）", "=h_rev_25-h_hrev_25-h_qrev_25", "Derived", "組件 1.04 億、其他（租賃、技術服務）0.23 億（" + HUX + "）；差額為主營／其他業務口徑"),
           ("h_cash", "貨幣資金 2025 年底", 14.19, "Verified", CSJ)]
    hdr(H, i, ["FY2025 補充", "", "", "FY2025", "標記", "來源"]); i += 1
    for k, lab, v, tag, src in H25:
        H.cell(i, 1, lab); c = H.cell(i, 4, v)
        if not str(v).startswith("="): c.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src); dn(f"{k}_25", f"Hist!$D${i}"); i += 1
    H.cell(i, 1, "FY2023、FY2024 管理費用、營業成本、稅前利益、所得稅：未找到（媒體轉述只有部分科目）；FY2022 分產品：未找到"); i += 1
    H.cell(i, 1, "組件及其他營收 FY2023、FY2024（Derived）"); H.cell(i, 2, "=h_rev_23-h_hrev_23-h_qrev_23"); H.cell(i, 3, "=h_rev_24-h_hrev_24-h_qrev_24")
    dn("h_crev_23", f"Hist!$B${i}"); dn("h_crev_24", f"Hist!$C${i}")
    H.column_dimensions["A"].width = 44; H.column_dimensions["F"].width = 60

    # ---------------- Q（季度） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度合併營收與淨利（億元）。累計值為原文（Verified）；單季為相減（Derived）。2026 年分產品營收未揭露，季度分部對帳不適用"
    hdr(Q, 2, ["項目", "Q1'26", "H1'26", "9M'26", "H1'25", "9M'25", "FY2025", "標記", "來源", "來源誘因"])
    qrows = [("營業收入（累計）", 4.23, 11.52, "=q9_rev", "=11.52/1.4854", 11.67, "=h_rev_25", "Verified", CSJ8 + "；" + EMH1 + "；" + EMQ3 + "；H1'25＝H1'26 ÷（1＋48.54%）為 Derived；9M'25 為申報稿審閱數（" + EM3 + "）"),
             ("歸屬母公司淨利（累計）", "未找到", 2.74, "=q9_ni", -0.3202, "=3.87/3.7054", "=h_ni_25", "Verified", EMH1 + "；" + EMQ3 + "；9M'25＝9M'26 ÷（1＋270.54%）為 Derived"),
             ("扣非淨利（累計）", 0.4025, 2.44, 3.49, "=2.44/(1-0.1934)", "=3.49/(1-0.1878)", 5.91, "Verified", CSJ8 + "；" + EMH1 + "；" + EMQ3 + "；上年同期由年增率反推（Derived）"),
             ("研發費用（累計）", "未找到", "=q9_rd-1.27", "=q9_rd", "未找到", "=q9_rd/2.9148", "=h_rd_25", "Verified", IFQ3 + "；H1'26＝9M − 第三季 1.27（Derived）")]
    for i, row in enumerate(qrows, 3):
        Q.cell(i, 1, row[0])
        for j, v in enumerate(row[1:7]):
            c = Q.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): c.font = BLUE
        Q.cell(i, 8, row[7]); Q.cell(i, 9, row[8])
    hdr(Q, 8, ["單季（Derived）", "Q1'26", "Q2'26", "Q3'26", "Q4'25", "", "", "標記", "說明"])
    Q.cell(9, 1, "營業收入（單季）"); Q["B9"] = "=B3"; Q["C9"] = "=C3-B3"; Q["D9"] = "=D3-C3"; Q["E9"] = "=G3-F3"; Q["H9"] = "Derived"
    Q.cell(10, 1, "扣非淨利（單季）"); Q["B10"] = "=B5"; Q["C10"] = "=C5-B5"; Q["D10"] = "=D5-C5"; Q["E10"] = "=G5-F5"; Q["H10"] = "Derived"
    Q.cell(11, 1, "檢查：第三季營收（累計相減 − 原文 6.65）"); Q["D11"] = "=ROUND(D9-q3_rev,6)"
    Q.cell(12, 1, "註：第三季歸屬母公司淨利 1.13 億（年減 13.37%，另一報導 −17.37%；環比 −49.55%）；第三季研發占營收 19.10%。2026 年第一季淨利未找到（扣非 0.4025 億）")
    Q.column_dimensions["A"].width = 40; Q.column_dimensions["I"].width = 70

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（人形＝全球出貨 × 份額 × 售價）→ 官方產品分部（人形、四足、組件及其他）→ 標準損益表 → 現金流（億元；FY2023–FY2025 實績、FY2026 前三季實績＋第四季假設、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = [("lship", "全球人形出貨（線 2b；台）"), ("shr", "宇樹人形份額（出貨 ÷ 全球）"), ("hu", "人形機器人出貨（台）"), ("lpr", "線 2b 全尺寸售價學習曲線年變化"),
            ("hasp", "人形機器人平均售價（萬元）"), ("hrev", "人形機器人營收"), ("qu", "四足機器人出貨（台）"), ("qasp", "四足機器人平均售價（萬元）"), ("qrev", "四足機器人營收"),
            ("crev", "組件及其他營收"), ("rev", "合併營收"), ("chk", "檢查：分部合計 − 合併營收"), ("gr", "營收年增率"),
            ("hgm", "人形機器人毛利率"), ("qgm", "四足機器人毛利率"), ("cgm", "組件及其他毛利率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gpm", "毛利率"),
            ("rd", "研發費用"), ("sm", "銷售費用"), ("ga", "管理費用（不含股份支付）"), ("sbc", "股份支付費用"), ("opex", "營業費用合計"), ("op", "營業利益"), ("opm", "營業利益率"),
            ("nonop", "業外（利息收入、補助等）"), ("resid", "歷史：營業利益至淨利殘差（業外、稅金及附加、所得稅合計）"), ("pbt", "稅前利益"), ("tax", "所得稅"),
            ("ni", "淨利（歸屬母公司）"), ("nim", "淨利率"), ("sh", "平均股數（百萬股）"), ("eps", "EPS（元）"), ("nig", "淨利年增率"),
            ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加"), ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）")]
    NOTE = {"lship": "FY2026 起讀 Line 頁（依 lnsel 選欄）", "shr": "FY2026＝推得出貨 ÷ 全球；之後線性收斂到 shr32", "hu": "FY2026＝人形營收 ÷ 售價；FY2027 起＝全球 × 份額",
            "lpr": "FY2027–30＝(P2030 ÷ P2026)^(1/4) − 1；FY2031–32＝(P2035 ÷ P2030)^(1/5) − 1", "hasp": "FY2027 起＝前一年 ×(1＋asp_k × 線年變化)",
            "hrev": "FY2026＝合併 × hs26；FY2027 起＝出貨 × 售價 ÷ 10,000", "qrev": "FY2026＝合併 − 人形 − 組件（殘差）；FY2027 起＝出貨 × 售價",
            "rev": "FY2026＝前三季 18.17 ＋ 第三季 6.65 × q4k；FY2027 起＝分部合計", "sm": "FY2026＝毛利 − 其他費用 − 營業利益（校準：淨利率＝前三季實績）",
            "op": "FY2026＝營收 × 前三季淨利率 ÷(1 − 稅率) − 業外", "resid": "只有歷史年：淨利 − 營業利益（Derived；稅前、所得稅未揭露）",
            "fcff": "NOPAT＝營業利益 ×(1 − tax_w)；股份支付不加回"}
    UND = "未揭露"
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(KEYS)}
        for k, l in KEYS:
            P.cell(R[k], 1, l)
            if k in NOTE: P.cell(R[k], 2, NOTE[k])
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"; yy = y - 2000
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                for k, nm in (("hu", "h_hu"), ("hasp", "h_hasp"), ("hrev", "h_hrev"), ("qu", "h_qu"), ("qasp", "h_qasp"), ("qrev", "h_qrev"), ("crev", "h_crev"),
                              ("rev", "h_rev"), ("hgm", "h_hgm"), ("qgm", "h_qgm"), ("rd", "h_rd"), ("sm", "h_sm"), ("ni", "h_ni")):
                    w(k, f"={nm}_{yy}")
                w("chk", f"=ROUND({cell('hrev')}+{cell('qrev')}+{cell('crev')}-{cell('rev')},6)")
                if y > 2023:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1")
                    w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年虧損）\")")
                else:
                    w("gr", UND); w("nig", UND)
                w("nim", f"={cell('ni')}/{cell('rev')}")
                for k in ("lship", "shr", "lpr", "nonop", "pbt", "tax", "da", "cx", "dwc", "fcff", "div", "fcfad"): w(k, UND)
                if y == 2025:
                    w("cgm", "=0.5936"); w("cogs", "=h_cogs_25"); w("gp", "=h_gp_25"); w("gpm", f"={cell('gp')}/{cell('rev')}")
                    w("ga", "=h_ga_25-h_sbc_25"); w("sbc", "=h_sbc_25"); w("opex", f"={cell('rd')}+{cell('sm')}+{cell('ga')}+{cell('sbc')}")
                    w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("resid", f"={cell('ni')}-{cell('op')}")
                    w("sh", "=sh_pre"); w("eps", f"={cell('ni')}/{cell('sh')}*100")
                else:
                    w("gpm", f"=h_mgm_{yy}")
                    for k in ("cgm", "cogs", "gp", "ga", "sbc", "opex", "op", "opm", "resid", "sh", "eps"): w(k, UND)
                continue
            w("resid", "—")
            w("lship", "=" + lsel(f"SHIP_{y}", s))
            if y == 2026:
                w("rev", "=q9_rev+q3_rev*q4k")
                w("hrev", f"={cell('rev')}*hs26"); w("crev", f"={cell('rev')}*cs26"); w("qrev", f"={cell('rev')}-{cell('hrev')}-{cell('crev')}")
                w("hasp", "=h_hasp_25*(1+dasp_h26)"); w("hu", f"={cell('hrev')}/{cell('hasp')}*10000")
                w("qasp", "=h_qasp_25*(1+dasp_q26)"); w("qu", f"={cell('qrev')}/{cell('qasp')}*10000")
                w("shr", f"={cell('hu')}/{cell('lship')}"); w("lpr", "—")
                w("hgm", "=hgm26"); w("qgm", "=qgm26"); w("cgm", "=cgm")
                w("rd", f"=q9_rd/q9_rev*{cell('rev')}")
                w("op", f"={cell('rev')}*(q9_ni/q9_rev)/(1-tax_w)-nonop26")
                w("nonop", "=nonop26"); w("sh", "=sh_26")
                w("cx", "=cx_26"); w("dwc", "=dwc26"); w("div", "=0")
            else:
                w("shr", f"=$F{R['shr']}+(shr32_{s}-$F{R['shr']})*({y}-2026)/6")
                w("hu", f"={cell('lship')}*{cell('shr')}")
                if y <= 2030: w("lpr", f"=({lsel('HUM_PRICE_2030', s)}/Line!$D${LR['HUM_PRICE_2026']})^(1/4)-1")
                else: w("lpr", f"=({lsel('HUM_PRICE_2035', s)}/{lsel('HUM_PRICE_2030', s)})^(1/5)-1")
                w("hasp", f"={prev('hasp')}*(1+asp_k*{cell('lpr')})")
                w("hrev", f"={cell('hu')}*{cell('hasp')}/10000")
                w("qu", f"={prev('qu')}*(1+{'qg' if y <= 2029 else 'qlt'}_{s})"); w("qasp", f"={prev('qasp')}*(1+qasp)")
                w("qrev", f"={cell('qu')}*{cell('qasp')}/10000")
                w("crev", f"={prev('crev')}*(1+cg_{s})")
                w("rev", f"={cell('hrev')}+{cell('qrev')}+{cell('crev')}")
                w("hgm", f"=hgm26+(hgm32_{s}-hgm26)*({y}-2026)/6"); w("qgm", f"=qgm_{s}"); w("cgm", "=cgm")
                w("rd", f"=rd_r*{cell('rev')}"); w("sm", f"=sm_r*{cell('rev')}")
                w("op", f"={cell('gp')}-{cell('opex')}")
                w("nonop", "=int_r*cash_v"); w("sh", "=sh_v")
                w("cx", f"=fc_cx*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})"); w("div", f"=pay_r*MAX(0,{cell('ni')})")
            w("chk", f"=ROUND({cell('hrev')}+{cell('qrev')}+{cell('crev')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('hrev')}*(1-{cell('hgm')})+{cell('qrev')}*(1-{cell('qgm')})+{cell('crev')}*(1-{cell('cgm')})")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gpm", f"={cell('gp')}/{cell('rev')}")
            w("ga", f"=ga_r*{cell('rev')}"); w("sbc", f"=sbc_r*{cell('rev')}")
            if y == 2026: w("sm", f"={cell('gp')}-{cell('rd')}-{cell('ga')}-{cell('sbc')}-{cell('op')}")
            w("opex", f"={cell('rd')}+{cell('sm')}+{cell('ga')}+{cell('sbc')}")
            w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('nonop')}"); w("tax", f"={cell('pbt')}*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"={cell('ni')}/{cell('sh')}*100"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
            w("da", f"=fc_da*{cell('rev')}")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        for k in ("shr", "lpr", "gr", "hgm", "qgm", "cgm", "gpm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        for k in ("lship", "hu", "qu"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "#,##0"
        PR[s] = R; rr = b + len(KEYS) + 2
    P.column_dimensions["A"].width = 46; P.column_dimensions["B"].width = 40
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "hrev", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（億元）"; D["B4"] = "=px*sh_v/100"
    D["A5"] = "負債 D（有息負債）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（資產負債表 2026-09-30）＝貨幣資金 − 有息負債"; D["B10"] = "=cash_v-debt_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC"; D["B11"] = "=ROUND(((px*sh_v/100)*(rf+beta*(erp+crp))+debt_v*kd*(1-tax_w))/((px*sh_v/100)+debt_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+7})/sh_v*100"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+8}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+10}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=MAX(0,(B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6})/sh_v*100*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 9, ps=t + 8, ev=t + 4, tvs=t + 10, alt=t + 13, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+9}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業 EV／營收 TTM 中位數 × FY2027 營收＝目標時點的 TTM 營收，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業 EV／營收 TTM 中位數 × FY2027 營收 ＋ 淨現金，÷ 股數", [f"=(ev_med*PL!G{PR[s]['rev']}+dcf_nc)/sh_v*100" for s in SC],
                "6 家同業（人形整機 1、機器人整機 2、零組件 3）；目標時點 2027-11 的 TTM 營收以 FY2027 代；淨現金以 2026-09-30 計（未加 FY2027 累積現金）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權；同業 EV／營收）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px", "2026-11-20 收盤，單一來源（etnet），待交易所資料複核")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（億元；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0")&"／"&IF(B{r_up}<-0.5,"無（已是最低級）",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"#,##0"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含本益比：加權目標價 ÷ 情境 B FY2028 EPS", f"=B{r_tp}/PL!H{PR['B']['eps']}", "分母：模型情境 B FY2028 EPS")
    r_pex = one("現價本益比：現價 ÷ 情境 B FY2027 EPS", f"=px/PL!G{PR['B']['eps']}", "分母：模型情境 B FY2027 EPS（共識未找到）")
    r_evr = one("現行 EV／營收 TTM ÷ 同業中位數", "=evs_self/ev_med", "只作對照")
    for rr_ in (r_up, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, pex=r_pex, evr=r_evr, gj=r_gj)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司實績、線模型對照；反解（整本重算）。公司未給全年指引、市場共識未找到"
    hdr(RC, 2, ["項目", "模型（情境 B）", "實績或對照", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收 vs 前三季實績", f"=PL!F{B_['rev']}", "=q9_rev", "差＝模型第四季營收（＝第三季 × q4k）"),
          ("FY2026 淨利率 vs 前三季實績淨利率", f"=PL!F{B_['nim']}", "=q9_ni/q9_rev", "模型以前三季淨利率校準，差應約為 0（業外與稅率經營業利益回推）"),
          ("FY2026 銷售費用率（校準推得） vs FY2025 實績", f"=PL!F{B_['sm']}/PL!F{B_['rev']}", "=h_sm_25/h_rev_25", "校準推得的費用率若遠離實績，表示毛利率或研發假設與前三季淨利率不一致"),
          ("FY2026 人形出貨（推得） vs FY2025 銷量", f"=PL!F{B_['hu']}", "=h_hu_25", "推得＝人形營收 ÷ 售價；2026 年銷量未揭露"),
          ("FY2026 宇樹人形份額（推得） vs FY2025（以 IDC 1.8 萬計）", f"=PL!F{B_['shr']}", "=h_hu_25/18000", "全球分母：2026 年讀線 2b 基準 5.5 萬台、2025 年為 IDC 約 1.8 萬（線 2b 追蹤指標頁，Interested-party）"),
          ("FY2027 EPS（模型 B） vs 共識", f"=PL!G{B_['eps']}", "未找到", "市場共識未找到"),
          ("宇樹現行 EV／營收 TTM vs 同業中位數", "=evs_self", "=ev_med", "現價隱含倍數約為同業中位數的數倍；同業 TTM 營收日期不一")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_)
        RC.cell(i, 4, f'=IF(ISNUMBER(C{i}),B{i}-C{i},"—")'); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_unitree.py 二分法；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_unitree.py 以 LibreOffice 重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "build_id"; SE["D2"] = "{BID}"
    SE["A3"] = "模型狀態"; SE["D3"] = f'=IF(ABS(D4-B4)>0.05,"STALE","CURRENT")'; SE["E3"] = '=IF(D3="CURRENT","快照與本檔一致","In 頁已改，Sens 須重跑")'
    SE["A4"] = "快照基準值（加權，元）"; SE["B4"] = base_tp; SE["C4"] = "本檔即時值（加權，元）"; SE["D4"] = "=tp"
    SE["F4"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.0"; SE[f"F{sr}"].number_format = "#,##0.0"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。無區間的鍵不列案例。lnsel_B 低／高＝情境 B 改讀線 2b 低／高欄（出貨與售價學習曲線同時換）。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線層級變數（線 2b：{LINE_REL} 下游取數頁）；與線模型不一致處逐項寫明，不在本模型靜默修改線口徑"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（線模型儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    sh = lambda k: f"{LN[k]['lo']:,.0f}／{LN[k]['base']:,.0f}／{LN[k]['hi']:,.0f}"
    rec_ = [("全球人形出貨 2026", "4（2b）", sh("SHIP_2026") + "（低／基準／高）", "台/年", "全球、含小型機", f"下游取數!C{LN['SHIP_2026']['row']}:E{LN['SHIP_2026']['row']}", "Derived（線模型）", "PL 全球人形出貨列 FY2026；份額分母", "一致（直接讀）"),
            ("全球人形出貨 2027–2032", "4（2b）", "2030：" + sh("SHIP_2030") + "；2032：" + sh("SHIP_2032"), "台/年", "全球、含小型機", f"下游取數!C{LN['SHIP_2027']['row']}:E{LN['SHIP_2032']['row']}", "Derived（線模型）", "PL 全球人形出貨列 FY2027–FY2032（A 讀高、B 讀基準、C 讀低）", "一致（直接讀）"),
            ("全尺寸出廠價 2026／2030／2035", "4（2b）", f"{LN['HUM_PRICE_2026']['base']:,.0f}；2030 " + sh("HUM_PRICE_2030") + "；2035 " + sh("HUM_PRICE_2035"), "USD/台", "全尺寸 Optimus Gen 2 類", f"下游取數!C{LN['HUM_PRICE_2030']['row']}:E{LN['HUM_PRICE_2035']['row']}", "Derived（線模型）",
             "PL 線 2b 全尺寸售價學習曲線年變化 → 宇樹人形售價年變化（× asp_k）", "不一致（口徑）：宇樹以小型機為主，FY2025 平均售價 16.64 萬元約 2.5 萬美元，遠低於線全尺寸 10.5 萬美元；本模型只借用降幅（Analogy，asp_k 0.5–1.5）"),
            ("整機毛利率", "4（2b）", "線基準 25%（15–45%）", "%", "全尺寸整機", "輸入頁（經公司檔案 UNITREE.md 引 輸入!F35）", "Assumed（線模型）", "未直接讀；hgm32 區間 32–50% 的下緣參考", "不一致（自設）：宇樹 FY2025 人形毛利率 63.18% 為實績，小型機科研市場；本模型 FY2032 基準 42%，高於線基準，理由：實績錨點＋產品組合不同"),
            ("宇樹人形份額", "4（2b）", "FY2025 29–40%（5,215 ÷ 1.3–1.8 萬）；FY2032 基準 10%（6–15%）", "%", "出貨台數", "線 2b 追蹤指標頁（Omdia、IDC，Interested-party）", "Assumed", "In shr32_*", "線模型無廠商份額（本模型自設）"),
            ("小型機旋轉關節多用行星減速（非諧波）", "4（2b）", "物理層註記", "—", "零組件結構", "線 2b 情境頁（經公司檔案 UNITREE.md 引 情境!F10）", "Analogy", "說明用；本模型不建 BOM（宇樹自製關節馬達、模組，自研國產化率逾 90%，公司說法）", "一致（說明）")]
    for i, row in enumerate(rec_, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 34; RE.column_dimensions["C"].width = 44; RE.column_dimensions["I"].width = 60

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("美國市場銷售受限（招股書揭露新機型進入美國受限）", "否", "海外營收占比 FY2025 約 44%（Derived）；美國占比未取得", "美國出口管制、海關、公司公告"),
             ("人形付費需求仍以科研、教育、展示為主（2025 人形營收 73.6% 來自科研教育）", "部分（shr32、hgm32 區間）", "工廠端付費採用無公開訂單；需求轉向前的庫存風險", "人形營收用途別、前五大客戶、工廠部署揭露"),
             ("存貨快速累積（2026-09-30 10.19 億，較年初 +177%）", "部分（dwc26）", "可能為備貨或滯銷，無法分辨（分不出來）", "第四季與年報存貨、存貨跌價準備"),
             ("中國人形廠價格戰與產能過剩（發改委警示）", "部分（情境 C、hgm32_C）", "", "競品售價、人形平均售價季度變化"),
             ("限售股解禁（上市後一年起）", "否", "解禁日與股數未查核（未取得）", "交易所解禁公告"),
             ("股價波動極大（上市首日 +460%，11-20 至 11-23 單日漲停 +20%）", "否", "現價只取單一來源收盤；評等對現價敏感", "交易所收盤資料"),
             ("上市後新股權激勵（股份支付）", "部分（sbc_r）", "未揭露", "董事會公告")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([CSJ, CSJ8, HUX, TMT, NBD, EM3, SINA, EMH1, EMQ3, IFQ3, SSQ3, ETN, SAH, SA + "：9880、2432、2590（港股）、688017、688322（滬）、2498（港股）", CNY10, YCD, LINE_REL], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"宇樹科技 Unitree（688836.SS）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-23（財報與資料商）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0.00")&" 元，單一來源）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：宇樹的股價由人形機器人能否從科研、教育與展示市場走到大量付費採用決定——人形營收＝全球人形出貨（線 2b）× 宇樹份額 × 平均售價；四足、組件及其他兩個官方產品分部以量 × 價或成長路徑建模"
    RD["A4"] = (f'="在情境 B 讀線 2b 基準出貨、宇樹人形份額 FY2032 收斂到 "&TEXT(shr32_B,"0%")&"、人形毛利率 FY2032 收斂到 "&TEXT(hgm32_B,"0%")&"、同業 EV／營收中位數 "&TEXT(ev_med,"0.0")&" 倍的假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 元（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 億元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"；現價本益比（情境 B FY2027 EPS）"&TEXT(Val!B{r_pex},"#,##0")&" 倍"'
    ev_rows = [13 + j for j, it in enumerate(rev or []) if it["key"] == "ev_med"]
    if len(ev_rows) >= 2:
        RD["A6"] = (f'="評等翻轉與失效條件：其他輸入不變下，同業 EV／營收中位數需達約 "&TEXT(Recon!C{ev_rows[1]},"0")&" 倍才升一級、約 "&TEXT(Recon!C{ev_rows[0]},"0")&" 倍才使目標價等於現價；'
                    f'人形份額、人形毛利率單獨調整在搜尋區間內都無法翻轉評等。往有利方向的可觀察指標：年報人形銷量與平均售價、人形毛利率、工廠或商用場景千台級付費訂單、同業倍數重估"')
    else:
        RD["A6"] = "評等翻轉與失效條件：見 Recon 頁反解（本檔為敏感度重算用，未寫入反解）"
    RD["A7"] = "範圍與限制：上市約三個月，2026 年分產品營收、毛利率未揭露（以區間假設）；現價只有單一來源；市場共識未找到。不讀 Siliconomics／Tokenomics／Wattonomics（算力線）"
    RD["A8"] = "高不確定性原則：證據到哪裡判斷到哪裡——FY2023–FY2025 分產品量價為招股書實績（經媒體轉述），2026 年只有合併營收與淨利；FY2027 起的出貨、份額、毛利率皆為 Assumed 或讀線模型，結論是條件式"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def recalc(src, outdir, prof="0"):
    os.makedirs(outdir, exist_ok=True)
    cmd = ["soffice", "-env:UserInstallation=file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), f"lo_{prof}"), "--headless", "--convert-to", "xlsx", "--outdir", outdir, src]
    for attempt in range(2):
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
            break
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", f"lo_{prof}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if attempt: raise
    return os.path.join(outdir, os.path.basename(src))

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

import threading
_tl = threading.local(); _ids = iter(range(1000)); _lk = threading.Lock()
def prof_id():
    if not hasattr(_tl, "p"):
        with _lk: _tl.p = str(next(_ids))
    return _tl.p

def run_case(ov, tag):
    d = tempfile.mkdtemp(prefix="uni_"); p = os.path.join(d, f"c_{tag}.xlsx")
    VR, _, DR = build(p, ov)
    out = recalc(p, os.path.join(d, "o"), prof_id())
    v = read_vals(out, VR, DR if tag == "base" else None); shutil.rmtree(d, ignore_errors=True); return v

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS] + [LN], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

CASES = [("q4k", "第四季 ÷ 第三季營收"), ("hs26", "FY2026 人形營收占比"), ("cs26", "FY2026 組件及其他占比"), ("dasp_h26", "FY2026 人形售價年變化"), ("dasp_q26", "FY2026 四足售價年變化"),
         ("dwc26", "FY2026 營運資金增加"), ("nonop26", "FY2026 業外"), ("cx_26", "FY2026 資本支出"), ("lnsel_B", "情境 B 讀線 2b 欄位"), ("shr32_B", "人形份額 FY2032（情境 B）"),
         ("asp_k", "售價降幅 ÷ 線學習曲線"), ("hgm26", "人形毛利率 FY2026"), ("hgm32_B", "人形毛利率 FY2032（情境 B）"), ("qg_B", "四足出貨年增 FY2027–29（情境 B）"),
         ("qlt_B", "四足出貨年增 FY2030–32（情境 B）"), ("qasp", "四足售價年變化"), ("qgm26", "四足毛利率 FY2026"), ("qgm_B", "四足毛利率（情境 B）"), ("cg_B", "組件及其他年增（情境 B）"),
         ("cgm", "組件及其他毛利率"), ("rd_r", "研發費用率"), ("sm_r", "銷售費用率"), ("ga_r", "管理費用率"), ("sbc_r", "股份支付費用率"), ("tax_w", "有效稅率"), ("int_r", "業外收益率"),
         ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("pay_r", "配息率"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"),
         ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("ev_med", "同業 EV／營收中位數")]
TIERS = [0.5, 0.25, -0.25, -0.5]
def rt_of(u): return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; f1 = lambda x: f"{x:,.1f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else f'{r[chr(118)+chr(97)+chr(108)]:.4f}'} | {r['base']} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else ("—" if v in (None, "—") else str(v)))
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f1(g(chr(66+i)+str(VR['dcfr'])))} | {f1(g(chr(66+i)+str(VR['cu'])))} | {f1(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    txt = f"""# physical_ai/unitree — 宇樹科技 Unitree（688836.SS，上交所科創板）個股模型 {VER}

> **上市狀態與現價（先讀）：** 已於 2026-08-19 在上交所科創板掛牌，代號 688836（發行價 150.80 元、發行 4,044.64 萬股、發行後總股本約 404.46 百萬股）。現價用 **2026-11-20 收盤 {inv('px'):,.2f} 元**，**單一來源**：etnet 經濟通 688836 報價頁的「前收市」（2026-11-23 讀取；同頁 11-23 漲停收 708.00）。Yahoo 日線檔 `physical_ai/data/closes2_20261109_20261120.csv` 沒有本檔；**須以交易所資料複核**，複核不符時只需改 In 頁 `px` 重跑本程式。
>
> **財報期數：** 招股書 FY2023–FY2025 三個完整年度（分產品營收、銷量、單價、毛利率；經媒體轉述，原文 PDF 未逐頁核對）；上市後公告 2026 年第一季（招股書內）、上半年（上市公告書）、第三季報（2026-10-28，上市後首份定期報告）。**2026 年分產品營收、毛利率、費用明細未揭露**，FY2026 只錨定合併營收與淨利；分產品以區間假設。市場共識與公司全年指引未找到。

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

本 README 的數字由 `build_unitree.py` 從重算後的成品讀出產生，不手寫。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | Comparable（同業 EV／營收） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f1(tp)} 元**（DCF {f1(dcf)}、Comparable {f1(cmp_)}），現價 {inv('px'):,.2f}（2026-11-20 收盤，單一來源），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–30 自由現金流扣股利 {f1(g(f"C{VR['gate']}"))} 億元）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g('B' + str(VR['pe'] - 1))}，為 WACC／ke／g／RONIC；無有息負債，WACC＝ke）50%＋Comparable 50%（6 家同業 EV／營收 TTM 中位數 {inv('ev_med'):.2f} 倍 × FY2027 營收＋淨現金）。相對倍數係數 1.00（未調整：缺歷史相對倍數）。
- 對照：宇樹現行 EV／營收 TTM {inv('evs_self'):.1f} 倍，為同業中位數的 {g(f"B{VR['evr']}"):.1f} 倍；現價 ÷ 情境 B FY2027 EPS＝{g(f"B{VR['pex']}"):,.0f} 倍；隱含本益比（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍。
- 終值占 EV 高（見 DCF 頁）：FY2032 後以 g＝MIN(3%, rf)＝1.61% 收斂，對仍在成長期的公司偏保守；Sens 有終值替代法與 g 案。

## 營收分部（官方產品分部；PL 頁三情境，情境 B 摘要，億元）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 全球人形出貨（線 2b，台） | {yr('B','lship','{:,.0f}')} |
| 宇樹人形出貨（台） | {yr('B','hu','{:,.0f}')} |
| 人形平均售價（萬元） | {yr('B','hasp','{:,.2f}')} |
| 人形機器人營收 | {yr('B','hrev')} |
| 四足機器人營收 | {yr('B','qrev')} |
| 組件及其他營收 | {yr('B','crev')} |
| 合併營收 | {yr('B','rev')} |
| 毛利率 | {yr('B','gpm','{:.1%}')} |
| 營業利益 | {yr('B','op')} |
| 淨利（歸屬母公司） | {yr('B','ni')} |
| EPS（元） | {yr('B','eps','{:,.2f}')} |

- FY2023–FY2025 為招股書實績（FY2023、FY2024 管理費用、營業成本、稅前、所得稅未找到，表中標「未揭露」）；FY2025 營業利益為 Derived（毛利 − 四項費用，不含財務費用）。
- FY2026＝前三季實績 18.17 億＋第三季 6.65 億 × 第四季係數；人形與組件占比為 Assumed，四足為殘差；淨利率以前三季實績校準（銷售費用為推得值）。
- FY2027 起：人形＝讀線 2b 全球出貨 × 宇樹份額（FY2026 推得值線性收斂到 FY2032 假設值）× 平均售價（前一年 × 線 2b 全尺寸學習曲線年變化 × asp_k）；四足＝出貨 × 售價；組件及其他＝成長路徑。

## 主要驅動因子與標記
| 驅動 | 基準（區間） | 標記 | 可證偽指標 |
|---|---|---|---|
| 宇樹人形份額 FY2032（shr32） | A 15%／B 10%（6–15%）／C 6% | Assumed | 年報人形銷量 ÷ IDC／Omdia 全球出貨 |
| 全球人形出貨（線 2b） | 2030 年 {LN['SHIP_2030']['lo']:,.0f}／{LN['SHIP_2030']['base']:,.0f}／{LN['SHIP_2030']['hi']:,.0f} 台 | Derived（線模型；其輸入 Assumed／Interested-party） | IDC、Omdia、TrendForce 季度出貨 |
| 人形毛利率 FY2032（hgm32） | A 50%／B 42%（32–50%）／C 32% | Assumed | 年報、半年報分產品毛利率 |
| 人形售價降幅 ÷ 線學習曲線（asp_k） | 1.0（0.5–1.5） | Analogy | 年報人形平均售價 |
| 同業 EV／營收中位數（ev_med） | {inv('ev_med'):.2f} 倍（四分位 {EV_LO:.2f}–{EV_HI:.2f}） | Derived（資料商頁面，日期不一） | 同業股價與 TTM 營收 |

結論主要建立在 Assumed／Analogy／Derived 輸入上：FY2027 起的出貨、份額、毛利率、費用率皆無公司指引；Verified 的只有 FY2023–FY2025 分產品量價與 2026 年前三季合併數。評等翻轉與失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f1(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算，二分法；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 讀線 2b（`{LINE_REL}` 下游取數頁）：全球人形出貨 SHIP_2026–SHIP_2032 低／基準／高（情境 A 讀高、B 讀基準、C 讀低），全尺寸出廠價 HUM_PRICE_2026／2030／2035（只取年變化率）。讀入值列在 Line 頁。
- 與線模型不一致：(1) 售價口徑——宇樹以小型機為主，FY2025 人形平均售價 16.64 萬元（約 2.5 萬美元），線模型是全尺寸 10.5 萬美元；只借用學習曲線降幅並乘 asp_k（Analogy）。(2) 毛利率——線整機基準 25%，宇樹 FY2025 實績 63.18%；本模型 FY2032 基準 42%（自設，理由：實績錨點與產品組合）。(3) 廠商份額——線模型沒有，本模型自設。

## 限制與未取得資料
- **2026 年分產品營收、銷量、毛利率、費用明細未揭露**（第三季報只有合併數）；FY2026 分產品以 Assumed 占比拆分。
- 招股書數字經媒體轉述（中國證券報、虎嗅、鈦媒體等），上交所原文 PDF 未逐頁核對；FY2025 人形「出貨逾 5,500 台」與「銷量 5,215 台」口徑不同，模型用銷量。
- **現價單一來源**（etnet）；stockanalysis 歷史頁只到 2026-11-13（696.00），無法交叉核對 11-20。
- β 未用兩來源 Blume 中位數（上市僅三個月，Analogy 1.3）；有息負債 2026-09-30 未取得（以 0 計）；稅率、業外為 Analogy／Assumed。
- 同業 EV／營收取自資料商頁面，頁面日期 2026-09-04 到 2026-10-23 不一，TTM 營收未更新到同一估值日；市值以 2026-11-20 收盤等比例調整（Derived）；6 家少於 8 家，不做迴歸；多數同業虧損，無法用本益比。
- 市場共識 EPS、目標價、公司全年指引：未找到。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。

## 檔案
- `build_unitree.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Line、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("UNI_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    quick = "--quick" in sys.argv
    if "--reuse" in sys.argv:   # 只改展示文字、In 頁未改時：沿用 sens.json 的敏感度與反解快照（指紋須相同）
        sens = json.load(open(os.path.join(HERE, "sens.json")))
        assert sens["fp"] == fingerprint(), "In 頁已改，不可 --reuse"
        return finish(sens, sens["revs"])
    base = run_case({}, "base"); print("基準", base); px0 = in_val("px", "base")
    jobs = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            jobs.append((k, drv, side, v))
    for s in SC:
        jobs.append((f"w_{s}", f"只看情境 {s}", "only", s))
    cases = []
    if not quick:
        def work(j):
            k, drv, side, v = j
            if side == "only":
                r = run_case({f"w_{x}": (1.0 if x == v else 0.0) for x in SC}, f"only_{v}")
                return dict(drv=drv, set=f"w_{v}＝1，其餘 0", key=f"w_{v}", **r)
            r = run_case({k: v}, f"{k}_{side}")
            return dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, **r)
        with ThreadPoolExecutor(4) as ex:
            cases = list(ex.map(work, jobs))
        for c in cases: print(c["key"], round(c["tp"], 1), c["rt"])
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rt_of(up_)))
    # 反解
    up0 = base["up"]
    upper = min([t for t in TIERS if t > up0], default=None); lower = max([t for t in TIERS if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)]
    if upper is not None: targets.append((f"升一級邊界（上檔 {upper:+.0%}）", upper))
    if lower is not None: targets.append((f"降一級邊界（上檔 {lower:+.0%}）", lower))
    SOLVE = [("shr32_B", 0.10, 0.0, 1.0, "情境 B 宇樹人形份額 FY2032（FY2026 推得值線性收斂；情境 A／C 不動）"),
             ("hgm32_B", 0.42, 0.0, 0.95, "情境 B 人形毛利率 FY2032（情境 A／C 不動）"),
             ("ev_med", round(EV_MED, 4), 0.5, 200.0, "同業 EV／營收 TTM 中位數（三情境同動；只影響 Comparable）")]
    def solve_one(arg):
        key, b0, lo, hi, note, lab, tgt = arg
        if quick: return dict(label=lab, key=key, val="（快速模式未解）", base=b0, note=note)
        f = lambda x: run_case({key: x}, "rev")["up"] - tgt
        flo, fhi = f(lo), f(hi)
        if flo * fhi > 0:
            return dict(label=lab, key=key, val=f"區間 {lo:g}～{hi:g} 內無解", base=b0, note=note)
        a, bb = lo, hi
        for _ in range(13):
            mid = (a + bb) / 2; fm = f(mid)
            if flo * fm <= 0: bb, fhi = mid, fm
            else: a, flo = mid, fm
        return dict(label=lab, key=key, val=(a + bb) / 2, base=b0, note=note)
    args = [(k, b0, lo, hi, note, lab, tgt) for (k, b0, lo, hi, note) in SOLVE for (lab, tgt) in targets]
    with ThreadPoolExecutor(4) as ex:
        revs = list(ex.map(solve_one, args))
    for r in revs: print("反解", r["key"], r["label"], r["val"])
    sens = dict(base=base, cases=cases, fp=fingerprint(), revs=revs)
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    finish(sens, revs)


def finish(sens, revs):
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="unif_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    bid = f"{VER}-{fingerprint()}"
    wb = openpyxl.load_workbook(raw)
    wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid); wb["Sens"]["D2"] = bid; wb.save(raw)
    out = recalc(raw, os.path.join(tmp, "o"), "final")
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10])
    assert not errs, errs
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="unitree", ticker="688836.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="億元", source=CSJ + "；" + HUX)
                                      for n, l, v in (("人形機器人", "人形機器人營收", 8.68), ("四足機器人", "四足機器人營收", 6.98), ("組件及其他", "組件及其他營收", 1.33))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收"],
                   note="招股書主營業務按產品劃分：人形機器人、四足機器人、組件（關節馬達、光達等）；本模型將組件與其他業務（租賃、技術服務）合併為「組件及其他」＝合併營收 − 人形 − 四足（FY2025 組件 1.04、其他 0.23，差額為主營／其他業務口徑）。2026 年分產品營收未揭露，PL 頁 FY2026 起為模型值。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()
