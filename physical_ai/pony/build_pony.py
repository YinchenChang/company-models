# 小馬智行 Pony.ai（PONY，Nasdaq ADS；港股 2026.HK）個股模型 v0.1 — AI實體應用 Project（線 1 Robotaxi）
# 用法：python3 physical_ai/pony/build_pony.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json，並跑核對
#       python3 physical_ai/pony/build_pony.py --raw OUT  → 只寫一本（讀環境變數 PONY_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度＝曆年（12 月決算，US GAAP，報表幣別美元）。金額：百萬美元（US$M）；股數：百萬股；EPS、股價：美元。1 ADS＝1 股 A 類普通股。
# 結構：README、In、Hist、Q、PL（三情境：車隊 × 每英里經濟 → Robotaxi 分部 → 三個官方分部 → 標準損益 → 現金流與資金缺口）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數一律讀線 1a 成品 physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx（「下游取數」頁；車價例外讀「成本輸入」頁，對帳頁註明）。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md（明確期延長到 FY2035，見 README）。
import sys, os, json, subprocess, shutil, hashlib, tempfile, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_小馬智行個股模型_{VER}.xlsx"
LINE_REL = "physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx"
LINE = os.path.join(ROOT, LINE_REL)
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2036)); COL = {y: "CDEFGHIJKLMNO"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … O＝FY2035
FC = [y for y in YRS if y >= 2027]                                                       # DCF 明確期 FY2027–FY2035

# ---------------- 出處 ----------------
Q326 = "公司 6-K 附件 99.1：2026 年第三季未經審計財務報告（2026-11-18）https://www.sec.gov/Archives/edgar/data/0001969302/000110465926129308/tm2629888d1_ex99-1.htm"
H126 = "公司 6-K 附件 99.1：2026 年中期報告（截至 2026-06-30 六個月；港股中報）https://www.sec.gov/Archives/edgar/data/0001969302/000110465926109772/tm2625374d1_ex99-1.pdf"
Q425 = "公司新聞稿：2025 年第四季及全年未經審計財務報告（2026-03）https://ir.pony.ai/news-releases/news-release-details/pony-ai-inc-scales-160-robotaxi-revenues-growth-yoy-and-500-fare"
F20 = "公司 20-F（FY2025，2026-04 申報）https://www.sec.gov/Archives/edgar/data/0001969302/000110465926046406/pony-20251231x20f.htm"
SA = "stockanalysis.com 統計與預測頁（資料商；讀取 2026-11-23，各頁報價日見各列）"
INV = "investing.com Pony AI（HKEX: 2026）共識頁（讀取 2026-11-23）"
TT = "Transport Topics 2026-11-18（Q3 財報報導；ADR 年初至今 −72%）https://ttnews.com/articles/pony-ai-earnings-q3-2026"
RH = "網約車監管信息交互系統 2026 年 4 月行業數據（全國訂單 9.24 億單；經同花順轉載）https://news.10jqka.com.cn/20260514/c676701943.shtml"
YCL = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
UST = "美國 10 年期公債殖利率 5.00%（2026-10-23；與 physical_ai/tesla 同一取值，跨模型一致）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
LN = f"線 1a 成品 {LINE_REL}"

# ---------------- 讀線模型 ----------------
LW = openpyxl.load_workbook(LINE, data_only=True)
_DS = LW["下游取數"]
DSR = {_DS.cell(r, 1).value: r for r in range(2, _DS.max_row + 1) if _DS.cell(r, 1).value}
def lv(i, col="D"):
    r = DSR[i]; v = _DS[f"{col}{r}"].value
    return v, f"下游取數!{col}{r}"
_CI = LW["成本輸入"]
CIR = {_CI.cell(r, 1).value: r for r in range(1, _CI.max_row + 1) if _CI.cell(r, 1).value}
CIY = {2026: "F", 2028: "I", 2030: "L", 2032: "O"}     # 基準欄（低成本／基準／高成本三欄中間）
CIL = {2026: "E", 2028: "H", 2030: "K", 2032: "N"}; CIH = {2026: "G", 2028: "J", 2030: "M", 2032: "P"}
def cv(i, y, cols=CIY):
    r = CIR[i]; return _CI[f"{cols[y]}{r}"].value, f"成本輸入!{cols[y]}{r}"

ANC = (2026, 2028, 2030, 2032)
def ratio_avg(i, a, b):
    return sum(lv(f"{i}_{y}", a)[0] / lv(f"{i}_{y}", b)[0] for y in ANC) / len(ANC)

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點（9M 實績＋第四季；公司目標為 Interested-party）")
inp("rt26", "FY2026 Robotaxi 服務營收（公司目標：超過 FY2025 的 3.5 倍）", "US$M", None, "=3.5*h_rt_25", None, "Interested-party",
    Q326 + "：全年 Robotaxi 收入目標為 2025 年的 3.5 倍；FY2025 16.607 → 58.1", "9M 實績 38.406 → 第四季隱含 19.7（Q3 單季 17.763）；Derived", CO)
inp("rk_q4", "FY2026 第四季 Robotruck 服務營收", "US$M", None, 17.952, None, "Assumed", "＝Q3 單季 17.952（" + Q326 + "）持平", "9M 實績 41.476；Q4'25 13.118")
inp("ls_q4", "FY2026 第四季智慧方案營收", "US$M", None, 9.0, None, "Assumed", "Q3 單季 8.859、Q4'25 9.350 的約略中間值", "9M 實績 35.162（" + Q326 + "）")
inp("g_anc", "FY2026 實績相對錨點（三分部同比例）", "%", -0.05, 0.0, 0.05, "Assumed", "基準＝9M 實績＋第四季錨點", "±5%：第四季為唯一未實現的一季（約占全年 29%）")
inp("gm26", "FY2026 毛利率（9M 實績）", "%", None, "=19.824/115.044", None, "Derived", Q326 + "：9M 毛利 19.824 ÷ 營收 115.044 ＝ 17.2%", "Q3 單季 17.8%")
inp("rd26", "FY2026 研發費用（GAAP，含股權報酬）", "US$M", None, "=164.178+60.064", None, "Derived", Q326 + "：9M 164.178＋第四季假設＝Q3 單季 60.064", "non-GAAP 9M 148.261；股權報酬不加回（DCF 標準規格）")
inp("sga26", "FY2026 銷售及管理費用（GAAP）", "US$M", None, "=48.455+16.579", None, "Derived", Q326 + "：9M 48.455＋第四季假設＝Q3 單季 16.579", "")
inp("nonop26", "FY2026 業外（投資收益、交易性證券公允價值、其他）", "US$M", None, "=32.579-51.355-23.612+10.379", None, "Derived",
    Q326 + "：9M 投資收益 32.579、交易性證券公允價值變動 −51.355、其他 −23.612（含預付長期投資減損 25.0）＋第四季投資收益＝Q3 10.379", "第四季公允價值變動假設 0（不可預測）")
inp("nci26", "FY2026 少數股東損益", "US$M", None, -19.388, None, "Verified", Q326 + "：9M −19.388（第四季假設 0）", "FY2027 起假設 0（少數股東所屬子公司未揭露）")
inp("cx26", "FY2026 資本支出", "US$M", None, "=62.325+17.707", None, "Derived", Q326 + "：9M 62.325＋第四季假設＝Q3 17.707", "")
inp("sh26", "FY2026 加權平均股數", "百萬股", None, 433.841452, None, "Verified", Q326 + "：9M 加權平均 433,841,452（基本＝稀釋，虧損）", "")

sec("B. 命題模組一：車隊（滲透）— 期末營運車隊 × 情境成長")
inp("fl25", "期末 Robotaxi 車隊 FY2025", "台", 1100, 1400, 1446, "Interested-party", Q425 + "：車隊超過 1,400 台；至 2026-03-25 累計生產 1,446 台", "2025 年底精確值未取得；低案 1,100（Assumed）。只用於 FY2026 平均車隊與校準", CO)
inp("fl26", "期末 Robotaxi 車隊 FY2026（公司目標 3,500 台、20 城）", "台", 3000, 3500, None, "Interested-party", Q326 + "：2026-09-30 車隊 2,768 台；年底目標 3,500 台", "低案 3,000（第四季淨增約 230 台；Assumed）", CO)
for s, v, note in (("A", 8000, "樂觀：海外協議中 4,000 台以上（" + Q326 + "）逐步落地＋中國 20 城擴張"),
                   ("B", 6500, "基準：FY2026 淨增約 2,100 台 → FY2027 淨增約 3,000 台（Assumed）"),
                   ("C", 4500, "保守：擴張放緩（價格戰、法規、融資約束）")):
    inp(f"f27_{s}", f"情境 {s}：期末車隊 FY2027", "台", 5000 if s == "B" else None, v, 8000 if s == "B" else None, "Assumed", note, "公司未給 FY2027 目標（未找到）")
for s, v, lo, hi, note in (("A", 0.70, None, None, "樂觀"), ("B", 0.50, 0.30, 0.70, "基準：FY2026 車隊年增約 150%、FY2027 約 86% 之後減速"), ("C", 0.25, None, None, "保守")):
    inp(f"fg_{s}", f"情境 {s}：車隊年增 FY2028–FY2030", "%/年", lo, v, hi, "Assumed", note, "可證偽：季度車隊揭露、城市數")
for s, v, lo, hi in (("A", 0.30, None, None), ("B", 0.20, 0.10, 0.30), ("C", 0.08, None, None)):
    inp(f"fglt_{s}", f"情境 {s}：車隊年增 FY2031–FY2035", "%/年", lo, v, hi, "Assumed", "長期擴張；以滲透率（中國網約車趟數占比）檢查合理性", "")
inp("tpvd", "每車每日付費趟數（滲透率換算用）", "趟/日", None, 23, None, "Interested-party", "公司揭露 23 趟（線 1a 輸入!K25 註記）", "只用於滲透率，不進營收（營收用每英里口徑）", CO)
inp("mkt26", "中國網約車年趟數 FY2026（滲透率分母）", "百萬趟", None, "=924*12", None, "Derived", RH + "：2026 年 4 月 9.24 億單 × 12", "季節性未調整")
inp("mkt_g", "中國網約車趟數年增", "%/年", None, 0.03, None, "Assumed", "只用於滲透率分母", "")

sec("C. 命題模組二：每英里經濟（讀線 1a「下游取數」中國營運商＝百度／小馬平均；本模型以情境乘數與 FY2026 校準套用到小馬）")
for y in ANC:
    v, ref = lv(f"RT_PMD_CN_{y}"); lo, rlo = lv(f"RT_PMD_CN_{y}", "C"); hi, rhi = lv(f"RT_PMD_CN_{y}", "E")
    inp(f"pmd{y % 100}", f"每車每日付費英里 {y}（線基準）", "英里/日", None, v, None, "Assumed", f"{LN} {ref}（線標 Assumed，承重）", f"線低成本／高成本欄 {lo}／{hi}（{rlo}、{rhi}）；奇數年取相鄰錨點算術平均，FY2033 起持平")
for y in ANC:
    v, ref = lv(f"RT_OPPRICE_CN_{y}"); lo, rlo = lv(f"RT_OPPRICE_CN_{y}", "C"); hi, rhi = lv(f"RT_OPPRICE_CN_{y}", "E")
    inp(f"prc{y % 100}", f"每付費英里價格 {y}（線基準）", "US$/英里", None, v, None, "Assumed", f"{LN} {ref}（線標 Assumed）", f"線低價／高價欄 {lo}／{hi}；奇數年算術平均，FY2033 起持平")
for y in range(2026, 2033):
    v, ref = lv(f"RT_OPEX_CN_{y}"); lo, _ = lv(f"RT_OPEX_CN_{y}", "C"); hi, _ = lv(f"RT_OPEX_CN_{y}", "E")
    inp(f"ox{y % 100}", f"每付費英里單位營運成本 {y}（線基準，不含研發、平台）", "US$/英里", None, v, None, "Derived", f"{LN} {ref}（線每英里成本堆疊第 1–8 項）", f"線低／高成本 {lo:.3f}／{hi:.3f}；FY2033 起持平")
for y in ANC:
    v, ref = cv("CN_VEH", y)
    inp(f"veh{y % 100}", f"單車取得成本 {y}（線基準，含自駕套件）", "US$/台", None, v, None, "Interested-party", f"線 1a 成本輸入頁 {ref}（不在下游取數頁；對帳頁註明）", "小馬 2027 車型整車 <23 萬元（Gasgoo 2026-05-21，經線 1a）", "營運商與媒體（立場：降本敘事）")
v_life, r_life = cv("CN_LIFE", 2026); v_paid, r_paid = cv("CN_PAID", 2026)
inp("veh_life", "車輛經濟壽命（年）", "年", 4, f"={v_life}/(pmd26/{v_paid}*365)", 7, "Derived", f"＝線壽命里程 {v_life:,}（{r_life}）÷（每車日付費英里 ÷ 付費占比 {v_paid}（{r_paid}）× 365）", "低 4、高 7（Assumed）；只用於汰換車輛的資本支出")
um_a = ratio_avg("RT_PMD_CN", "C", "D"); um_c = ratio_avg("RT_PMD_CN", "E", "D")
for s, v, lo, hi, tag, note in (("A", round(um_a, 4), None, None, "Derived", "線低成本欄（高稼動）÷ 基準欄，四個錨點平均"),
                                ("B", 1.0, round(um_c, 4), round(um_a, 4), "Assumed", "基準＝線基準；低／高＝線高成本欄／低成本欄 ÷ 基準"),
                                ("C", round(um_c, 4), None, None, "Derived", "線高成本欄（低稼動）÷ 基準欄，四個錨點平均")):
    inp(f"um_{s}", f"情境 {s}：稼動率乘數（× 線每車每日付費英里）", "x", lo, v, hi, tag, note, "稼動率為線 1a 承重變數")
pm_lo = ratio_avg("RT_OPPRICE_CN", "C", "D"); pm_hi = ratio_avg("RT_OPPRICE_CN", "E", "D")
for s, v, lo, hi, tag, note in (("A", 1.0, None, None, "Assumed", "樂觀：價格照線基準（擴張不靠降價）"),
                                ("B", 1.0, round(pm_lo, 4), round(pm_hi, 4), "Assumed", "基準＝線基準；低／高＝線低價欄／高價欄 ÷ 基準"),
                                ("C", 0.85, None, None, "Assumed", "保守：價格戰（線低價欄約 0.71 倍與基準的中間）")):
    inp(f"pm_{s}", f"情境 {s}：價格乘數（× 線每英里價格）", "x", lo, v, hi, tag, note, "線 1a 中國價格以 7.2 匯率換算美元；含補貼")
for s, v, lo, hi, note in (("A", 0.75, None, None, "樂觀：遠端監看比與保險降幅快於線基準"),
                           ("B", 1.0, 0.75, 1.4, "基準＝線基準；低 0.75、高 1.4（Assumed；線高成本欄約 4.7 倍屬聯合極端，未採用）"),
                           ("C", 1.4, None, None, "保守：成本下降慢於線基準")):
    inp(f"om_{s}", f"情境 {s}：營運成本乘數（× 線每英里單位營運成本）", "x", lo, v, hi, "Assumed", note, "可證偽：公司保費（Gen-7 保費為人駕計程車約 40%）、遠端監看比揭露")
inp("k_rt", "校準：Robotaxi 營收 ÷ 車資推算（FY2026）", "x", None, "=rt26/(((fl25+fl26)/2)*pmd26*365*prc26/1000000)", None, "Derived",
    "＝FY2026 公司目標營收 ÷（平均車隊 × 每車日付費英里 × 365 × 每英里價格）", "大於 1 的部分＝聯合部署的車輛銷售與合作收入（產品收入 Q3 年增 90.9%，" + Q326 + "）；FY2027 起沿用（持續）")
inp("cr_k", "校準：小馬 Robotaxi 營業成本 ÷ 線成本推算（FY2026）", "x", None,
    "=((rt26+h_rk_9m+rk_q4+h_ls_9m+ls_q4)*(1+g_anc)*(1-gm26)-(h_rk_9m+rk_q4)*(1+g_anc)*(1-gm_rk)-(h_ls_9m+ls_q4)*(1+g_anc)*(1-gm_ls))/(((fl25+fl26)/2)*pmd26*365*ox26/1000000)", None, "Derived",
    "＝（FY2026 合併營業成本 − Robotruck 與智慧方案營業成本）÷（付費英里 × 線單位營運成本）", "小於 1：聯合部署車輛由夥伴持有、承擔部分營運成本；分部毛利未揭露，故以殘差推得")
inp("kdec", "校準係數衰減（0＝持續；1＝FY2027 起線性衰減、FY2029 回到 1）", "x", None, 0, 1, "Assumed", "個股模型原則第 10 節第 48 條：基準持續", "持續的理由：聯合部署模式（夥伴持有車輛、小馬賣車與收服務費）預期擴大（Uber 歐洲 2,000 台以上）")

sec("D. 其他官方分部與費用（成長路徑）")
for s, v, lo, hi in (("A", 0.35, None, None), ("B", 0.25, 0.10, 0.35), ("C", 0.10, None, None)):
    inp(f"grk_{s}", f"情境 {s}：Robotruck 服務營收年增 FY2027–FY2030", "%/年", lo, v, hi, "Assumed", "FY2025 +0.6%、9M'26 +50.9%（Derived）", "第四代自駕卡車部署（公司新聞稿）")
for s, v, lo, hi in (("A", 0.20, None, None), ("B", 0.12, 0.05, 0.20), ("C", 0.05, None, None)):
    inp(f"grkl_{s}", f"情境 {s}：Robotruck 服務營收年增 FY2031–FY2035", "%/年", lo, v, hi, "Assumed", "", "")
for s, v, lo, hi in (("A", 0.25, None, None), ("B", 0.15, 0.05, 0.25), ("C", 0.05, None, None)):
    inp(f"gls_{s}", f"情境 {s}：智慧方案營收年增 FY2027–FY2030", "%/年", lo, v, hi, "Assumed", "FY2025 +19.7%、9M'26 +50.0%（Derived；含域控制器出貨）", "")
for s, v, lo, hi in (("A", 0.12, None, None), ("B", 0.08, 0.03, 0.12), ("C", 0.03, None, None)):
    inp(f"glsl_{s}", f"情境 {s}：智慧方案營收年增 FY2031–FY2035", "%/年", lo, v, hi, "Assumed", "", "")
inp("gm_rk", "Robotruck 毛利率", "%", 0.05, 0.10, 0.15, "Assumed", "分部毛利未揭露（未找到）", "貨運服務；全公司毛利率 FY2025 15.7%、9M'26 17.2%")
inp("gm_ls", "智慧方案毛利率", "%", 0.15, 0.25, 0.35, "Assumed", "分部毛利未揭露（未找到）", "授權與域控制器硬體混合")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", 0.03, 0.0, 0.06), ("C", 0.06, None, None)):
    inp(f"grd_{s}", f"情境 {s}：研發費用年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "9M'26 GAAP 研發 +4.6%、Q3 non-GAAP −5.4%（Derived）", "FY2024 240.2（含上市股權報酬）、FY2025 217.4")
inp("sga_v", "銷售及管理費用：變動部分 ÷ 營收", "%", 0.03, 0.05, 0.08, "Assumed", "其餘為固定部分（FY2026 減去變動部分後年增 3%）", "FY2025 SG&A 57.6 ÷ 營收 90.0 ＝ 64%（規模小、固定費用為主）")
inp("yld", "流動資金收益率（投資收益 ÷ 期初流動資金）", "%", 0.02, 0.03, 0.04, "Derived", "9M'26 投資收益 32.579 ÷ 約 1,200 流動資金 × 4/3 ≈ 3.6%", "FY2025 42.985 ÷ 約 1,300 ≈ 3.3%；基準取整 3%")
inp("tax_w", "有效稅率（稅前為正時；虧損年不計稅盾）", "%", 0.15, 0.15, 0.25, "Assumed", "中國高新技術企業 15%；高案法定 25%", "累積虧損抵稅未建模（保守）")
inp("dil_g", "股數年增（股權報酬稀釋）", "%/年", 0.0, 0.01, 0.03, "Assumed", "FY2025 底 433.54 → 2026 Q3 加權 434.42（+0.2%，" + F20 + "；Derived）", "未計新股融資（另見資金缺口）")

sec("E. 現金流與資金缺口")
inp("own_r", "自有車隊比例（新增與汰換車輛中由小馬出資者）", "%", 0.3, 0.6, 1.0, "Assumed", "未揭露（未找到）；聯合部署模式由夥伴持有車輛（Uber 歐洲 2,000 台以上）", "FY2026 資本支出約 80 對淨增約 2,100 台，若全數自有約 3.8 萬美元／台")
inp("cx_oth", "其他資本支出（研發設備、場站、Robotruck）", "US$M/年", 10, 20, 35, "Assumed", "FY2025 資本支出 43.9、FY2024 11.4（" + Q425 + "）", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.0, 0.10, 0.20, "Assumed", "Q3'26 營業現金流出增加主因營運資金（公司說法）", "資產負債表營運資金明細未取得")
inp("min_cash", "最低流動資金（資金缺口門檻）", "US$M", None, 300, None, "Assumed", "約一年研發＋管銷", "只用於資金缺口，不進目標價")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤，PONY ADS）", "US$", None, 3.31, None, "Verified", YCL + "：PONY 2026-11-20 收盤 3.31", "同檔 2026.HK 收盤 26.78 港元 ÷ 7.8426 ＝ 3.41 美元（1 ADS＝1 股）；" + TT)
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("px_hk", "對照：港股 2026.HK 收盤（只作對照，不進模型）", "HK$", None, 26.780001, None, "Verified", YCL + "：2026.HK 2026-11-20", "HKD=X 7.8426（同檔）")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（A 類 352.45＋B 類 81.09；2025-12-31）", "百萬股", None, 433.541553, None, "Verified", F20 + "：A 類 352,452,783、B 類 81,088,770", "Q3'26 加權 434.42（" + Q326 + "）")
inp("sh_dil", "選擇權與 RSU 淨稀釋（庫藏股法）", "%", 0.0, 0.02, 0.05, "Assumed", "未取得（20-F 未讀到選擇權與 RSU 明細）", "DCF 標準規格第 3 節要求計入；以區間代替")
inp("sh_v", "評價日股數", "百萬股", None, "=434.42199*(1+sh_dil)", None, "Derived", "＝Q3'26 加權股數 434.42 ×（1＋淨稀釋）", "")
inp("liq_v", "流動資金（現金、受限現金、短期與長期理財債務工具；2026-09-30）", "US$M", None, 1154.0, None, "Verified", Q326 + "：US$1,154.0M（RMB7,737.3M）", "2026-06-30 為 1,206.1；季減 52.1")
inp("inv_v", "其他投資（資產負債表短期＋長期投資＋現金 − 流動資金）", "US$M", None, "=354.39+3.619+722.9+287.605-liq_v", None, "Derived", Q326 + "：現金 354.390、受限 3.619、短期投資 722.900、長期投資 287.605", "含交易性證券與股權投資（Q3 公允價值損失 77.0）")
inp("inv_k", "其他投資計入比例", "x", 0.0, 0.5, 1.0, "Assumed", "股權投資流動性與價值不確定；基準 50%", "")
inp("lease_v", "債務與租賃（資料商總債務）", "US$M", None, 19.03, None, "Interested-party", SA + "：Total Debt 19.03M（多為租賃）", "資產負債表無借款科目", DV)
inp("mi_v", "非控制權益（2026-09-30）", "US$M", None, 40.057, None, "Verified", Q326, "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債）", "%", None, 0.05, None, "Verified", UST, "2026-10-23 值，距股價日逾 2 個交易日：逾期（跨模型共用，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", 0.006, 0.009, 0.012, "Assumed", "Damodaran 中國 CRP 約 0.8–1.0%（本輪未回原文核對：未查核，不作 Verified）", "營收以中國為主")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", 1.1425, "=MEDIAN(0.67*2.01+0.33,0.67*1.31+0.33)", 1.7425, "Derived",
    SA + "：來源一＝中國自駕同業 5 年 β 中位數 2.01（文遠 2.01、地平線 2.35、禾賽 1.47）；來源二＝全球自駕與叫車 1.31（Mobileye 1.24、Tesla 1.74、Uber 1.31）", "小馬上市未滿兩年，資料商 β 為 n/a；Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業 EV／營收）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺小馬對同業的歷史相對倍數；上市未滿兩年）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業 EV／NTM 營收（2026-11-20 收盤換算；資料商頁面日期各異，市值依 csv 收盤等比例換算、淨現金與 NTM 營收不變；Interested-party：賣方共識營收）")
PEERS = [("ev_wrd", "文遠知行 WRD（中國 Robotaxi）", "=(1360*2.985/4.14-700.93)/(1360/6.32)", "2026-11-13 頁面：市值 13.6 億、淨現金 7.009 億、Forward PS 6.32、股價 4.14；11-20 收盤 2.985"),
         ("ev_bidu", "百度 BIDU（Apollo Go；集團口徑）", "=(27030*61.9/81.87-7850)/(27030/1.38)", "2026-11-06 頁面：市值 270.3 億、淨現金 78.5 億、Forward PS 1.38、股價 81.87；11-20 收盤 61.90"),
         ("ev_uber", "Uber UBER（叫車平台）", "=133650/(129600/2.00)", "2026-11-20 頁面：EV 1,336.5 億、市值 1,296 億、Forward PS 2.00（同日，不換算）"),
         ("ev_tsla", "Tesla TSLA（Robotaxi、整車）", "=1340000/(1370000/11.55)", "2026-11-20 頁面：EV 1.34 兆、市值 1.37 兆、Forward PS 11.55（同日）"),
         ("ev_mbly", "Mobileye MBLY（自駕方案）", "=(5660*4.57/6.66-1360)/(5660/2.86)", "2026-10-26 頁面：市值 56.6 億、淨現金 13.6 億、Forward PS 2.86、股價 6.66；11-20 收盤 4.57"),
         ("ev_hsai", "禾賽 HSAI（光達）", "=953.45/(1850/2.44)", "2026-11-20 頁面：EV 9.5345 億、市值 18.5 億、Forward PS 2.44（同日）")]
for k, lab, f, note in PEERS:
    inp(k, f"EV／NTM 營收：{lab}", "x", None, f, None, "Derived", SA + "：" + note, "NTM 營收＝市值 ÷ Forward PS（資料商共識，Interested-party）；會計口徑：美國上市公司 US GAAP（百度、文遠、禾賽為 US GAAP 申報），租賃負債不計入 EV", DV)
inp("ev_med", "同業 EV／NTM 營收中位數（6 家）", "x", None, "=MEDIAN(ev_wrd,ev_bidu,ev_uber,ev_tsla,ev_mbly,ev_hsai)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）", "Tesla 11.3 倍為離群值（中位數不受影響）")
inp("ev_hzn", "對照：地平線 9660.HK EV／TTM 營收（只作對照，不進模型）", "x", None, "=(64290*3.12/4.05-7590)/4930", None, "Derived", SA + "：2026-10-02 頁面市值 642.9 億港元、淨現金 75.9 億、TTM 營收 49.3 億；11-20 收盤 3.12", "無 Forward PS，口徑為 TTM，不進中位數", DV)
inp("ev_self", "小馬現行 EV／NTM 營收（只作對照，不進模型）", "x", None, "=(px*434.42199-liq_v-inv_v*0.5+lease_v+mi_v)/(2230/8.54)", None, "Derived", SA + "：2026-10-23 頁面市值 22.3 億、Forward PS 8.54 → NTM 營收約 2.61 億", "EV 以本模型淨現金口徑（其他投資計 50%）", DV)
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "US$", None, 19.86, None, "Interested-party", SA + "：17 位分析師平均 19.86（中位 18.65、低 10、高 32.8；2026-08-19 更新，Q3 財報前）", INV + "：港股平均目標 148.89 港元（10 位）", DV)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "US$M", None, 150.0, None, "Interested-party", SA + "：平均 150.0（區間 138.0–165.4；2026-08-20）", "財報前（Q3 前），只作歷史對照", DV)


def build(path, ov=None, sens=None, rev=None, bid="（基準檔）"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$M、股數百萬股、EPS 與股價美元"
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
    FAM = [("f27", "期末車隊 FY2027"), ("fg", "車隊年增 FY2028–FY2030"), ("fglt", "車隊年增 FY2031–FY2035"), ("um", "稼動率乘數"), ("pm", "價格乘數"),
           ("om", "營運成本乘數"), ("grk", "Robotruck 年增 FY2027–30"), ("grkl", "Robotruck 年增 FY2031–35"), ("gls", "智慧方案年增 FY2027–30"),
           ("glsl", "智慧方案年增 FY2031–35"), ("grd", "研發費用年增")]
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
    H["A1"] = "Hist — 歷史財務（年度，US GAAP，US$M；FY2023 分部未取得）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "9M'26", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rt", "Robotaxi 服務營收", NA, 7.266, 16.607, 38.406, "Verified", Q425 + "；" + Q326),
          ("h_rk", "Robotruck 服務營收", NA, 40.365, 40.601, 41.476, "Verified", Q425 + "；" + Q326),
          ("h_ls", "智慧方案（原授權與應用）營收", NA, 27.394, 32.793, 35.162, "Verified", Q425 + "；" + Q326 + "（2026 年起更名，口徑相同：9M'25 23.443＋Q4'25 9.350＝FY2025 32.793）"),
          ("h_rev", "營收", 71.899, 75.025, 90.001, 115.044, "Verified", F20 + "（FY2023）；" + Q425 + "；" + Q326),
          ("h_cogs", "營業成本", 55.015, 63.622, 75.840, 95.220, "Verified", "同上"),
          ("h_rd", "研發費用", 122.707, 240.179, 217.419, 164.178, "Verified", "同上（FY2024 含上市時一次性股權報酬）"),
          ("h_sga", "銷售及管理費用", 37.417, 56.747, 57.599, 48.455, "Verified", "同上"),
          ("h_op", "營業利益（虧損）", -143.240, -285.523, -260.857, -192.809, "Verified", "同上"),
          ("h_pbt", "稅前利益（虧損）", NA, -275.005, -76.758, -235.197, "Verified", "同上（FY2025 含交易性證券公允價值利益 128.031）"),
          ("h_tax", "所得稅", NA, -0.001, 0.0, -0.001, "Verified", "同上"),
          ("h_ni", "淨利（虧損）", -125.328, -275.006, -76.758, -235.198, "Verified", "同上"),
          ("h_nci", "少數股東損益", NA, -0.885, 57.211, -19.388, "Verified", "同上"),
          ("h_nia", "歸屬母公司淨利（虧損）", -124.812, -274.121, -133.969, -215.810, "Verified", "同上"),
          ("h_sh", "加權平均股數（基本）", NA, 114.318765, 379.914317, 433.841452, "Verified", "同上（2024-11 上市前股數少，FY2024 EPS 不具比較性）"),
          ("h_cx", "資本支出", NA, 11.397, 43.875, 62.325, "Verified", "同上"),
          ("h_ocf", "營業現金流", NA, -110.758, -164.955, -178.747, "Verified", "同上"),
          ("h_sbc", "股權報酬", NA, 127.003, 30.798, NA, "Verified", Q425 + "（9M'26 未取得總額；非 GAAP 調節差額約 24）")]
    HROW = {}
    for i, (k, lab, a, b, c, d, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c, d)):
            x = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): x.font = BLUE
        H.cell(i, 6, tag); H.cell(i, 7, src)
        for y, L in (("23", "B"), ("24", "C"), ("25", "D"), ("9m", "E")): dn(f"{k}_{y}", f"Hist!${L}${i}")
        HROW[k] = i
    i = 3 + len(HR)
    H.cell(i, 1, "檢查：分部合計 − 營收（FY2024–9M'26）")
    for L in "CDE": H.cell(i, "BCDE".index(L) + 2, f"=ROUND({L}3+{L}4+{L}5-{L}6,6)")
    H.cell(i + 1, 1, "FY2023 分部營收、稅前、股數、資本支出：未取得（20-F 摘要只取得合併數；上市前股數不具比較性）")
    H.column_dimensions["A"].width = 40; H.column_dimensions["G"].width = 80

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度與半年度分部營收（US$M）。累計值為原文（Verified）；單季為相減（Derived）。Q1、Q2 單季分部未取得（中報只揭露上半年）"
    hdr(Q, 2, ["累計（輸入）", "H1'25", "9M'25", "FY'25", "H1'26", "9M'26", "標記", "來源", "來源誘因"])
    cum = [("Robotaxi 服務（累計）", 3.256, 9.950, 16.607, 20.643, 38.406), ("Robotruck 服務（累計）", 17.300, 27.483, 40.601, 23.524, 41.476),
           ("智慧方案（累計）", 14.878, 23.443, 32.793, 26.303, 35.162), ("營收（累計）", 35.434, 60.876, 90.001, 70.470, 115.044)]
    for i, row in enumerate(cum, 3):
        Q.cell(i, 1, row[0])
        for j, v in enumerate(row[1:]): Q.cell(i, 2 + j, v).font = BLUE
        Q.cell(i, 7, "Verified"); Q.cell(i, 8, H126 + "；" + Q326 + "；" + Q425)
    q0 = 9
    hdr(Q, q0, ["單季或半年（Derived）", "H1'25", "Q3'25", "Q4'25", "H1'26", "Q3'26", "標記", "來源", "來源誘因"])
    for i in range(4):
        rr = q0 + 1 + i; src = 3 + i
        Q.cell(rr, 1, ["Robotaxi 服務", "Robotruck 服務", "智慧方案", "營收"][i])
        for j, f in enumerate([f"=B{src}", f"=C{src}-B{src}", f"=D{src}-C{src}", f"=E{src}", f"=F{src}-E{src}"]): Q.cell(rr, 2 + j, f)
        Q.cell(rr, 7, "Derived"); Q.cell(rr, 8, "累計原文相減")
    Q.cell(q0 + 5, 1, "檢查：分部合計 − 營收（應為 0）")
    for j in range(5):
        L = "BCDEF"[j]; Q.cell(q0 + 5, 2 + j, f"=ROUND(SUM({L}{q0+1}:{L}{q0+3})-{L}{q0+4},6)")
    Q.cell(q0 + 6, 1, "Robotaxi 占營收"); [Q.cell(q0 + 6, 2 + j, f"={'BCDEF'[j]}{q0+1}/{'BCDEF'[j]}{q0+4}") for j in range(5)]
    Q.cell(q0 + 7, 1, "期末 Robotaxi 車隊（台）"); Q.cell(q0 + 7, 4, "=fl25"); Q.cell(q0 + 7, 5, 1975); Q.cell(q0 + 7, 6, 2768)
    Q.cell(q0 + 7, 7, "Interested-party"); Q.cell(q0 + 7, 8, H126 + "（1,975）；" + Q326 + "（2,768）"); Q.cell(q0 + 7, 9, CO)
    Q.cell(q0 + 8, 1, "Q3'26 每車每日 Robotaxi 營收（US$；平均車隊，Derived）"); Q.cell(q0 + 8, 6, f"=F{q0+1}*1000000/((E{q0+7}+F{q0+7})/2)/92")
    Q.cell(q0 + 9, 1, "線 1a 基準推算每車每日車資（US$；每車日付費英里 × 每英里價格）"); Q.cell(q0 + 9, 6, "=pmd26*prc26")
    Q.cell(q0 + 10, 1, "Q3'26 毛利率"); Q.cell(q0 + 10, 6, "=7.913/44.574")
    Q.cell(q0 + 11, 1, "車資收入年增：Q1'26 +456.5%、Q2'26 +849.3%、Q3'26 近 400%（公司揭露，Interested-party）；Gen-7 單車保費年減 19%、約為人駕計程車 40%")
    Q.column_dimensions["A"].width = 52

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境（US$M）：車隊 × 每車日付費英里 × 每英里價格（讀線 1a）→ Robotaxi 分部 → 三個官方分部 → 標準損益表（US GAAP）→ 現金流與資金缺口。FY2023–FY2025 實績、FY2026 為 9M 實績＋第四季錨點、FY2027 起模型"
    PR = {}; rr = 3
    KEYS = [("fl", "期末 Robotaxi 車隊（台）"), ("fla", "平均 Robotaxi 車隊（台）"), ("pmd", "每車每日付費英里（英里）"), ("mi", "年付費英里（百萬英里）"),
            ("prc", "每付費英里價格（US$）"), ("trips", "Robotaxi 年趟數（百萬趟）"), ("mkt", "中國網約車年趟數（百萬趟）"), ("pen", "滲透率：小馬趟數 ÷ 中國網約車趟數"),
            ("rtf", "Robotaxi：車資（車隊 × 每車日付費英里 × 每英里價格）"), ("rto", "Robotaxi：其他（聯合部署車輛銷售、合作收入；校準）"),
            ("rt", "Robotaxi 服務營收"), ("rk", "Robotruck 服務營收"), ("ls", "智慧方案營收"), ("rev", "營收"),
            ("chk1", "檢查：Robotaxi 子列合計 − Robotaxi（應為 0）"), ("chk2", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"),
            ("opx", "每付費英里單位營運成本（線 × 情境乘數 × 校準；US$）"), ("rpm", "每付費英里 Robotaxi 營收（US$）"), ("gpm", "每付費英里 Robotaxi 毛利（US$）"),
            ("cgrt", "營業成本：Robotaxi"), ("cgrk", "營業成本：Robotruck"), ("cgls", "營業成本：智慧方案"), ("cogs", "營業成本"),
            ("gp", "毛利"), ("gm", "毛利率"), ("rtgm", "Robotaxi 分部毛利率（模型推算；未揭露）"),
            ("rd", "研發費用"), ("sga", "銷售及管理費用"), ("opex", "營業費用合計"), ("op", "營業利益"), ("opm", "營業利益率"), ("opf", "營業利益轉正年份旗標"),
            ("nonop", "業外收支（投資收益等）"), ("pbt", "稅前利益"), ("tax", "所得稅"), ("nit", "淨利（含少數股東）"), ("nci", "少數股東損益"), ("ni", "歸屬母公司淨利"),
            ("nim", "淨利率"), ("nig", "淨利年增率"), ("sh", "加權平均股數（百萬股）"), ("eps", "EPS（美元；基本＝稀釋）"),
            ("nopat", "NOPAT（營業利益為正時 ×(1 − 稅率)）"), ("da", "折舊攤銷（資本支出五年平均攤提；未揭露，推算）"), ("cx", "資本支出"), ("dwc", "營運資金增加"),
            ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）"),
            ("liq", "期末流動資金（未計新融資）"), ("gap", "資金缺口（最低流動資金 − 期末流動資金，正值＝需融資）"), ("gapf", "資金缺口年份旗標")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(KEYS)}
        for k, l in KEYS: P.cell(R[k], 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k, c=c: f"{c}{R[k]}"; prev = lambda k, p=p: f"{p}{R[k]}"
            def w(k, f, c=c): P[f"{c}{R[k]}"] = f
            yy = str(y)[2:]
            if y <= 2025:
                hk = {2023: "23", 2024: "24", 2025: "25"}[y]
                for k, h in (("rt", "h_rt"), ("rk", "h_rk"), ("ls", "h_ls"), ("rev", "h_rev"), ("cogs", "h_cogs"), ("rd", "h_rd"), ("sga", "h_sga"), ("op", "h_op"),
                             ("nit", "h_ni"), ("ni", "h_nia"), ("cx", "h_cx")):
                    w(k, f"=h_{h[2:]}_{hk}")
                if y >= 2024:
                    for k, h in (("pbt", "h_pbt"), ("tax", "h_tax"), ("nci", "h_nci"), ("sh", "h_sh")): w(k, f"=h_{h[2:]}_{hk}")
                    w("nonop", f"={cell('pbt')}-{cell('op')}"); w("eps", f"={cell('ni')}/{cell('sh')}")
                    w("chk2", f"=ROUND({cell('rt')}+{cell('rk')}+{cell('ls')}-{cell('rev')},6)")
                else:
                    for k in ("pbt", "tax", "nci", "sh", "eps"): w(k, "未取得")
                    w("nonop", f"={cell('nit')}-{cell('op')}")
                if y == 2025: w("fl", "=fl25")
                if y >= 2024: w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
                w("opex", f"={cell('rd')}+{cell('sga')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('nit')}/{cell('rev')}")
                if y >= 2024: w("nig", f"=IF({prev('nit')}<0,\"虧損\",{cell('nit')}/{prev('nit')}-1)")
                w("opf", 9999)
                continue
            # ---- 預測（FY2026 起；FY2026 三情境相同）----
            n = y - 2026
            def interp(pref, yv=y):
                if yv in ANC: return f"{pref}{yv % 100}"
                if yv > 2032: return f"{pref}32"
                return f"({pref}{(yv - 1) % 100}+{pref}{(yv + 1) % 100})/2"
            if y == 2026:
                w("fl", "=fl26"); um = "1"; pm = "1"; om = "1"; kk = "k_rt"; ck = "cr_k"
            else:
                if y == 2027: w("fl", f"=f27_{s}")
                elif y <= 2030: w("fl", f"={prev('fl')}*(1+fg_{s})")
                else: w("fl", f"={prev('fl')}*(1+fglt_{s})")
                um, pm, om = f"um_{s}", f"pm_{s}", f"om_{s}"
                kk = f"(1+(k_rt-1)*MAX(0,1-kdec*{n}/3))"; ck = f"(1+(cr_k-1)*MAX(0,1-kdec*{n}/3))"
            w("fla", f"=({prev('fl')}+{cell('fl')})/2")
            w("pmd", f"={interp('pmd')}*{um}")
            w("mi", f"={cell('fla')}*{cell('pmd')}*365/1000000")
            w("prc", f"={interp('prc')}*{pm}")
            w("trips", f"={cell('fla')}*tpvd*365/1000000")
            w("mkt", f"=mkt26*(1+mkt_g)^{n}")
            w("pen", f"={cell('trips')}/{cell('mkt')}")
            oxk = f"ox{min(y, 2032) % 100}"
            if y == 2026:
                w("rt", "=rt26*(1+g_anc)"); w("rtf", f"={cell('mi')}*{cell('prc')}"); w("rto", f"={cell('rt')}-{cell('rtf')}")
                w("rk", "=(h_rk_9m+rk_q4)*(1+g_anc)"); w("ls", "=(h_ls_9m+ls_q4)*(1+g_anc)")
                w("cgrk", f"={cell('rk')}*(1-gm_rk)"); w("cgls", f"={cell('ls')}*(1-gm_ls)")
                w("cogs", f"={cell('rev')}*(1-gm26)"); w("cgrt", f"={cell('cogs')}-{cell('cgrk')}-{cell('cgls')}")
                w("opx", f"={cell('cgrt')}/{cell('mi')}")
                w("rd", "=rd26"); w("sga", "=sga26"); w("nonop", "=nonop26"); w("nci", "=nci26"); w("sh", "=sh26")
                w("cx", "=cx26")
                w("liq", f"=liq_v+{cell('fcff')}*stub_frac")
            else:
                w("rtf", f"={cell('mi')}*{cell('prc')}"); w("rto", f"={cell('rtf')}*({kk}-1)"); w("rt", f"={cell('rtf')}+{cell('rto')}")
                g1, g2 = (f"grk_{s}", f"gls_{s}") if y <= 2030 else (f"grkl_{s}", f"glsl_{s}")
                w("rk", f"={prev('rk')}*(1+{g1})"); w("ls", f"={prev('ls')}*(1+{g2})")
                w("opx", f"={oxk}*{om}*{ck}")
                w("cgrt", f"={cell('mi')}*{cell('opx')}"); w("cgrk", f"={cell('rk')}*(1-gm_rk)"); w("cgls", f"={cell('ls')}*(1-gm_ls)")
                w("cogs", f"={cell('cgrt')}+{cell('cgrk')}+{cell('cgls')}")
                w("rd", f"={prev('rd')}*(1+grd_{s})"); w("sga", f"=(sga26-sga_v*$F{R['rev']})*1.03^{n}+sga_v*{cell('rev')}")
                w("nonop", f"=yld*MAX(0,{prev('liq')})"); w("nci", "=0"); w("sh", f"=sh_v*(1+dil_g)^{n}")
                vk = f"veh{min(y, 2032) % 100}" if (y in ANC or y > 2032) else f"(veh{(y - 1) % 100}+veh{(y + 1) % 100})/2"
                w("cx", f"=own_r*(MAX(0,{cell('fl')}-{prev('fl')})+{prev('fl')}/veh_life)*{vk}/1000000+cx_oth")
                w("liq", f"={prev('liq')}+{cell('fcff')}+{cell('nonop')}-{cell('div')}")
            w("rev", f"={cell('rt')}+{cell('rk')}+{cell('ls')}")
            w("chk1", f"=ROUND({cell('rtf')}+{cell('rto')}-{cell('rt')},6)")
            w("chk2", f"=ROUND({cell('rt')}+{cell('rk')}+{cell('ls')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("rpm", f"={cell('rt')}/{cell('mi')}"); w("gpm", f"=({cell('rt')}-{cell('cgrt')})/{cell('mi')}")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}"); w("rtgm", f"=1-{cell('cgrt')}/{cell('rt')}")
            w("opex", f"={cell('rd')}+{cell('sga')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("opf", f"=IF({cell('op')}>0,{y},9999)")
            w("pbt", f"={cell('op')}+{cell('nonop')}"); w("tax", f"=IF({cell('pbt')}>0,{cell('pbt')}*tax_w,0)"); w("nit", f"={cell('pbt')}-{cell('tax')}")
            w("ni", f"={cell('nit')}-{cell('nci')}"); w("nim", f"={cell('nit')}/{cell('rev')}")
            w("nig", f"=IF({prev('nit')}<0,IF({cell('nit')}<0,\"虧損\",\"轉盈\"),{cell('nit')}/{prev('nit')}-1)")
            w("eps", f"={cell('ni')}/{cell('sh')}")
            w("nopat", f"=IF({cell('op')}>0,{cell('op')}*(1-tax_w),{cell('op')})")
            cols5 = [COL[x] for x in range(max(2023, y - 4), y + 1)]
            w("da", "=SUM(" + ",".join(f"{L}{R['cx']}" for L in cols5) + ")/5")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('nopat')}+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", "=0"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
            w("gap", f"=MAX(0,min_cash-{cell('liq')})"); w("gapf", f"=IF({cell('gap')}>0,{y},9999)")
        P.cell(R["fl"], 2, "FY2027＝情境值；FY2028–30 ×(1＋fg)；FY2031–35 ×(1＋fglt)")
        P.cell(R["pmd"], 2, "線 1a 中國基準 × 稼動率乘數")
        P.cell(R["prc"], 2, "線 1a 中國基準 × 價格乘數")
        P.cell(R["rto"], 2, "＝車資 ×(校準 k_rt − 1)；FY2026＝公司目標 − 車資")
        P.cell(R["opx"], 2, "FY2026＝殘差推得；FY2027 起＝線基準 × 情境乘數 × cr_k")
        P.cell(R["cogs"], 2, "FY2026＝營收 ×(1 − 9M 毛利率)")
        P.cell(R["sga"], 2, "固定部分 ×1.03^n ＋ sga_v × 營收")
        P.cell(R["fcff"], 2, "股權報酬不加回；虧損年不計稅盾")
        P.cell(R["liq"], 2, "FY2026＝9 月底 ＋ Q4 FCFF（stub）；其後＋FCFF＋業外 − 股利")
        for k in ("gr", "gm", "rtgm", "opm", "nim", "nig", "pen"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%" if k != "pen" else "0.00%"
        PR[s] = R; rr = b + len(KEYS) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 36

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2035（車隊仍在擴張，延長到 FY2035），期中慣例；終值＝NOPAT(FY2036) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M）"; D["B4"] = "=px*sh_v"
    D["A5"] = "負債 D（租賃）"; D["B5"] = "=lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=(rf+0.02)*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-09-30）＝流動資金 ＋ 其他投資 × inv_k − 租賃 − 非控制權益"; D["B10"] = "=liq_v+inv_k*inv_v-lease_v-mi_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v)*(rf+beta*(erp+crp))+lease_v*(rf+0.02)*(1-tax_w))/((px*sh_v)+lease_v)+wacc_adj-wacc,9)"
    for r_ in range(3, 11): D[f"B{r_}"].number_format = "0.00%" if r_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13; NC = len(FC); DCOL = "BCDEFGHIJ"
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
                 ("NOPAT FY2036＝NOPAT FY2035 ×(1＋g)", f"={L}{b+1}*(1+g_tv)"),
                 ("終值（FY2035 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^{NC}"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+7})/sh_v"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t）", f"=B{t+8}*(1+ke)^t_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+10}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF({L}{b+1}>0,B{t+2}/{L}{b+1},\"N/M\")"),
                 ("終值替代法：FCFF(FY2035)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=MAX(0,B{t}+{L}{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^{NC}+B{t+5}+B{t+6})/sh_v*(1+ke)^t_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 9, ps=t + 8, ev=t + 4, tvs=t + 10, alt=t + 13, eq=t + 7)
        dn(f"dcf_{s}", f"DCF!$B${t+9}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業 EV／NTM 營收中位數 × FY2028 營收 ＋ 目標時點淨現金，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}"); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    def cf(s, k, y): return f"PL!{COL[y]}{PR[s][k]}"
    r_nct = vrow("目標時點淨現金（評價日淨現金 ＋ Stub FCFF ＋ FY2027 FCFF 與業外）", [f"=dcf_nc+{cf(s,'fcff',2026)}*stub_frac+{cf(s,'fcff',2027)}+{cf(s,'nonop',2027)}" for s in SC], "US$M")
    r_cu = vrow("Comparable（未調整）：EV／營收中位數 × FY2028 營收 ＋ 目標時點淨現金（每股）", [f"=(ev_med*{cf(s,'rev',2028)}+{L}{r_nct})/sh_v" for s, L in zip(SC, "BCD")],
                "目標時點的 NTM 以 FY2028 代；同一倍數套用三個分部（分部同業不足）")
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
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"0.00")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    # 資金缺口
    r_gapmax = vr; V.cell(vr, 1, "資金缺口：FY2026–FY2035 最大缺口（US$M；情境 A／B／C；未計新融資）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=MAX(PL!F{PR[s]['gap']}:PL!O{PR[s]['gap']})")
    vr += 1
    r_gapyr = vr; V.cell(vr, 1, "資金缺口：流動資金首次低於最低流動資金的年份（情境 A／B／C）")
    for j, s in enumerate(SC):
        g = PR[s]['gapf']
        V.cell(vr, 2 + j, f"=IF(MIN(PL!F{g}:PL!O{g})=9999,\"FY2035 前無缺口\",\"FY\"&MIN(PL!F{g}:PL!O{g}))")
    vr += 1
    r_liqmin = vr; V.cell(vr, 1, "資金缺口：FY2026–FY2035 最低期末流動資金（US$M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=MIN(PL!F{PR[s]['liq']}:PL!O{PR[s]['liq']})")
    vr += 1
    r_opy = vr; V.cell(vr, 1, "營業利益轉正年份（情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=IF(MIN(PL!F{PR[s]['opf']}:PL!O{PR[s]['opf']})=9999,\"FY2035 前未轉正\",MIN(PL!F{PR[s]['opf']}:PL!O{PR[s]['opf']}))")
    vr += 1
    r_pen30 = vr; V.cell(vr, 1, "滲透率 FY2030（小馬趟數 ÷ 中國網約車趟數；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!J{PR[s]['pen']}")
    vr += 1
    r_pen35 = vr; V.cell(vr, 1, "滲透率 FY2035（情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!O{PR[s]['pen']}")
    vr += 1
    r_gpm30 = vr; V.cell(vr, 1, "每付費英里 Robotaxi 毛利 FY2030（US$；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=PL!J{PR[s]['gpm']}")
    vr += 2
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_cash = one("現價隱含：每股淨現金（評價日淨現金 ÷ 評價日股數）", "=dcf_nc/sh_v", "現價與每股淨現金的差＝市場給營運業務的價值")
    r_mkt = one("現價隱含：營運業務 EV（市值 − 淨現金，US$M）", "=px*sh_v-dcf_nc")
    one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_evs = one("隱含 EV／營收：加權目標價隱含 EV ÷ 情境 B FY2028 營收", f"=(B{r_tp}*sh_v-dcf_nc)/PL!H{PR['B']['rev']}", "分母：模型情境 B FY2028 營收")
    r_evs2 = one("現行 EV／NTM 營收（小馬）÷ 同業中位數", "=ev_self/ev_med", "只作對照")
    for r_ in (r_up, r_a1, r_a2, r_a3): V[f"B{r_}"].number_format = "0.0%"
    for r_ in (r_pen30, r_pen35):
        for L in "BCD": V[f"{L}{r_}"].number_format = "0.00%"
    for r_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{r_}"].number_format = "0.00"
    V.column_dimensions["A"].width = 76; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, nct=r_nct, pd2=r_pd2, gj=r_gj, gapmax=r_gapmax, gapyr=r_gapyr, liqmin=r_liqmin, opy=r_opy,
              pen30=r_pen30, pen35=r_pen35, gpm30=r_gpm30, cash=r_cash, mkt=r_mkt, wk=r_wk, evs=r_evs, evs2=r_evs2, vg=r_vg)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司目標、市場共識、線 1a 對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司、共識或線", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 Robotaxi 營收 vs 公司目標（3.5 倍）", f"=PL!F{B_['rt']}", "=rt26", "公司目標（Interested-party）"),
          ("FY2026 期末車隊 vs 公司目標", f"=PL!F{B_['fl']}", 3500, "公司目標 3,500 台（Interested-party）"),
          ("FY2026 營收 vs 共識", f"=PL!F{B_['rev']}", "=cons_rev26", "資料商共識（2026-08-20，財報前，只作歷史對照）"),
          ("每車每日 Robotaxi 營收：Q3'26 實績 vs 線基準車資推算（US$）", "=Q!F17", "=Q!F18", "實績含聯合部署車輛銷售與合作收入；比值即 k_rt 的季度對照（口徑不同，只看方向）"),
          ("每付費英里單位營運成本 FY2026：模型殘差推得 vs 線基準（US$）", f"=PL!F{B_['opx']}", "=ox26", "比值＝cr_k；分部毛利未揭露，模型值為殘差"),
          ("每付費英里單位營運成本 FY2030：模型（線 × 校準）vs 線基準（US$）", f"=PL!J{B_['opx']}", "=ox30", "線為百度／小馬平均"),
          ("加權目標價 vs 賣方平均目標價（US$）", "=tp", "=cons_tp", "賣方 17 位平均（2026-08-19，Q3 前；Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_pony.py 格點＋線性內插後驗證；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_pony.py 以 LibreOffice 重算每一案例）；結果為常數快照。估值評等欄只看上檔空間"
    SE["A2"] = "說明：tpvd、mkt26、mkt_g 只用於滲透率顯示，min_cash 只用於資金缺口，皆無區間；rel_k 基準即 1.00，故「相對倍數係數＝1」案例＝基準"
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
    SE.cell(sr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_pony.py")
    for col, wd in zip("ABCDEFGHI", [50, 40, 24, 12, 10, 12, 10, 10, 50]): SE.column_dimensions[col].width = wd

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線 1a 變數（{LINE_REL}；建置時讀值，不手打）與不一致處"
    hdr(RE, 2, ["變數", "線 ID 或儲存格", "線基準值", "線低／高", "本模型用法", "一致性", "標記", "影響本模型哪些儲存格"])
    rec = []
    for y in ANC:
        v, ref = lv(f"RT_PMD_CN_{y}"); lo, _ = lv(f"RT_PMD_CN_{y}", "C"); hi, _ = lv(f"RT_PMD_CN_{y}", "E")
        rec.append((f"中國每車每日付費英里 {y}", f"RT_PMD_CN_{y}（{ref}）", v, f"{lo}／{hi}", "基準 × 情境稼動率乘數（A 線低成本欄比、C 線高成本欄比）", "一致（基準直接讀線）", "Assumed", f"In pmd{y % 100} → PL 每車每日付費英里"))
    for y in ANC:
        v, ref = lv(f"RT_OPPRICE_CN_{y}"); lo, _ = lv(f"RT_OPPRICE_CN_{y}", "C"); hi, _ = lv(f"RT_OPPRICE_CN_{y}", "E")
        rec.append((f"中國每付費英里價格 {y}", f"RT_OPPRICE_CN_{y}（{ref}）", v, f"{lo}／{hi}", "基準 × 價格乘數 × 校準 k_rt（含聯合部署車輛銷售）", "不一致：小馬每英里營收＝線價格 × k_rt（FY2026 校準，見 Recon）", "Assumed", f"In prc{y % 100} → PL 車資列"))
    for y in range(2026, 2033):
        v, ref = lv(f"RT_OPEX_CN_{y}"); lo, _ = lv(f"RT_OPEX_CN_{y}", "C"); hi, _ = lv(f"RT_OPEX_CN_{y}", "E")
        rec.append((f"中國每付費英里單位營運成本 {y}", f"RT_OPEX_CN_{y}（{ref}）", round(v, 4), f"{lo:.3f}／{hi:.3f}", "線基準 × 情境成本乘數 × 校準 cr_k", "不一致：小馬成本水準以 FY2026 殘差校準（cr_k），只沿用線的下降斜率", "Derived", f"In ox{y % 100} → PL 營業成本：Robotaxi"))
    for y in ANC:
        v, ref = cv("CN_VEH", y)
        rec.append((f"中國單車取得成本 {y}", f"CN_VEH（{ref}）", v, "—", "自有車隊資本支出", "一致；但不在下游取數頁（成本輸入頁），下次線改版時請列入下游取數", "Interested-party", f"In veh{y % 100} → PL 資本支出"))
    for i_, lab in (("RT_CPM_CN_2026", "中國每付費英里全成本 2026（含研發攤提、平台）"), ("RT_CPM_CN_2030", "中國每付費英里全成本 2030"), ("RT_CPM_CN_2032", "中國每付費英里全成本 2032")):
        v, ref = lv(i_); lo, _ = lv(i_, "C"); hi, _ = lv(i_, "E")
        rec.append((lab, f"{i_}（{ref}）", round(v, 4), f"{lo:.3f}／{hi:.3f}", "未用：本模型研發與管銷用小馬實績（FY2026 GAAP 研發約 224），不用線的分攤（百度／小馬平均 5 億美元／年、車隊 5,000 台）", "口徑不同：線全成本假設的研發分攤與車隊規模不是小馬自己的", "Derived", "只作對照"))
    v, ref = lv("RT_YEAR_1USD_CN"); rec.append(("中國全成本跨 1 美元年份", f"RT_YEAR_1USD_CN（{ref}）", v, "—", "未用；本模型對應指標＝營業利益轉正年份（Val 頁）", "口徑不同", "Derived", "只作對照"))
    rec.append(("中國滲透率", "線 1a 無中國滲透率（只有美國 RT_PEN_*）", "—", "—", "本模型自建：小馬趟數 ÷ 中國網約車趟數（網約車監管信息交互系統）", "線未提供（缺）", "Derived", "PL 滲透率列"))
    rec.append(("每車每日付費趟數", "線 1a 輸入!K25 註記", 23, "—", "滲透率換算用", "一致", "Interested-party", "In tpvd"))
    for i_, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i_, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["B"].width = 36; RE.column_dimensions["E"].width = 50; RE.column_dimensions["F"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("中國 Robotaxi 價格戰（百度蘿蔔快跑、文遠、滴滴）", "是（pm_C、pm_B 低案）", "情境 C 與 Sens", "城市牌價、每英里實收、補貼"),
             ("法規與事故（單一事故導致城市停運；2026-03 武漢系統故障停運為類比）", "否", "事件型、不可量化", "城市許可、事故通報"),
             ("融資：流動資金不足以支撐擴張到營業利益轉正", "是（資金缺口列）", "Val 頁資金缺口；新股融資稀釋未建模", "季度流動資金、增資公告"),
             ("美中監管（美國對中國自駕技術限制、ADR 退市風險）", "否", "事件型", "美國商務部規則、HFCAA 名單"),
             ("交易性證券與股權投資的公允價值波動", "部分（inv_k）", "Q3'26 公允價值損失 77.0", "季報業外科目"),
             ("海外合作（Uber 歐洲 2,000 台、ComfortDelGro、Bolt）落地時點", "部分（f27_A、fg_A）", "情境 A", "海外城市營運公告"),
             ("雙重股權結構（B 類高投票權）", "否", "治理", "股東會")]
    for i_, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i_, 1 + j, v)
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i_, s_ in enumerate([Q326, H126, Q425, F20, SA, INV, TT, RH, YCL, UST, LN + "（下游取數、成本輸入、輸入頁）"], 1):
        SO.cell(i_, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"小馬智行 Pony.ai（PONY；港股 2026.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-23（Q3 財報 2026-11-18）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：小馬的價值取決於 Robotaxi 在中國（及海外）的滲透速度——車隊規模 × 每車每日付費英里 × 每英里價格——以及每英里單位營運成本能降到多低；短期收支平衡不是主軸，但資金缺口決定能否撐到規模化"
    RD["A4"] = (f'="在情境 B 車隊 FY2027 "&TEXT(f27_B,"#,##0")&" 台、FY2028–30 年增 "&TEXT(fg_B,"0%")&"、每英里價格與成本照線 1a 基準（FY2026 校準）的假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = (f'="資金缺口（情境 B，未計新融資）：最低期末流動資金 "&TEXT(Val!C{r_liqmin},"#,##0")&" 百萬美元、最大缺口 "&TEXT(Val!C{r_gapmax},"#,##0")&"（首次缺口 "&Val!C{r_gapyr}&"）；營業利益轉正 "&Val!C{r_opy}&"；滲透率 FY2030 "&TEXT(Val!C{r_pen30},"0.00%")&"、FY2035 "&TEXT(Val!C{r_pen35},"0.00%")')
    RD["A7"] = "失效條件（可觀察）：期末車隊連兩季低於情境 B 路徑；Robotaxi 每車每日營收（Q3'26 實績見 Q 頁）連兩季下降；毛利率跌破 15%；流動資金季減幅擴大到每季 1 億美元以上"
    RD["A8"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics；線 1a 只提供中國每英里價格、稼動率與成本（百度／小馬平均），不提供中國滲透率，本模型自建。Robotruck、智慧方案以成長路徑建模"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file:///tmp/h_pony/lo_profile"
def recalc_many(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 20):
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + paths[i:i + 20],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=170)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile"], stderr=subprocess.DEVNULL)
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
    d = tempfile.mkdtemp(prefix="pony_"); paths = []; meta = {}
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

CASES = [("g_anc", "FY2026 實績相對錨點"), ("fl25", "FY2025 期末車隊（校準錨點）"), ("fl26", "FY2026 期末車隊"),
         ("f27_B", "FY2027 期末車隊（情境 B）"), ("fg_B", "車隊年增 FY2028–30（情境 B）"), ("fglt_B", "車隊年增 FY2031–35（情境 B）"),
         ("um_B", "稼動率乘數（情境 B）"), ("pm_B", "價格乘數（情境 B）"), ("om_B", "營運成本乘數（情境 B）"), ("kdec", "校準係數衰減（FY2029 回到 1）"),
         ("veh_life", "車輛經濟壽命"), ("grk_B", "Robotruck 年增 FY2027–30（情境 B）"), ("grkl_B", "Robotruck 年增 FY2031–35（情境 B）"),
         ("gls_B", "智慧方案年增 FY2027–30（情境 B）"), ("glsl_B", "智慧方案年增 FY2031–35（情境 B）"), ("gm_rk", "Robotruck 毛利率（同時重算 cr_k）"),
         ("gm_ls", "智慧方案毛利率（同時重算 cr_k）"), ("grd_B", "研發費用年增（情境 B）"), ("sga_v", "管銷變動部分 ÷ 營收"), ("yld", "流動資金收益率"),
         ("tax_w", "有效稅率"), ("dil_g", "股數年增"), ("own_r", "自有車隊比例"), ("cx_oth", "其他資本支出"), ("fc_wc", "營運資金 ÷ 營收增量"),
         ("sh_dil", "選擇權與 RSU 淨稀釋"), ("inv_k", "其他投資計入比例"), ("crp", "中國國家風險溢酬"), ("beta", "β"), ("erp", "ERP"),
         ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重")]

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
    YS = list(range(2025, 2036))
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {f0(g(chr(66+i)+str(VR['gate'])))} | {f0(g(chr(66+i)+str(VR['liqmin'])))} | {g(chr(66+i)+str(VR['gapyr']))} | {g(chr(66+i)+str(VR['opy']))} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    chk = "\n".join(f"- `{c[0]}`：{c[1]}" for c in checks)
    txt = f"""# physical_ai/pony — 小馬智行 Pony.ai（PONY；港股 2026.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

Nasdaq ADS（1 ADS＝1 股 A 類普通股，美元）；港交所 2026.HK（2025-11 上市）。線 1 Robotaxi（線 1a 成品 `{LINE_REL}`）。本 README 的數字由 `build_pony.py` 從重算後的成品讀出產生，不手寫。

**範圍限制（高不確定性原則）：** 小馬仍在大幅虧損、分部毛利與營業利益未揭露、2026-11 前只有約兩年上市財報。本模型做到：分部營收（三個官方分部）由驅動因子推導、標準損益表、FCFF DCF 與同業 EV／營收；**沒有做到**：分部別毛利（公司未揭露，以 FY2026 合併毛利殘差校準）、FY2023 分部營收（未取得）、選擇權與 RSU 明細（未取得，以 0–5% 區間代替）、新股融資與稀釋（只列資金缺口，不假設融資條件）、累積虧損抵稅。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A6'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月；美元）
| 情境 | DCF（推到 12 個月後） | Comparable | 情境目標價 | 收支 FY26–30（US$M） | 最低流動資金 | 首次資金缺口 | 營業利益轉正 |
|---|---|---|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B {f0(g(f"C{VR['gate']}"))} 百萬美元）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC；明確期 FY2027–FY2035）50%＋Comparable 50%（同業 6 家 EV／NTM 營收中位數 {inv('ev_med'):.2f} 倍 × FY2028 營收 ＋ 目標時點淨現金）。相對倍數係數 1.00（未調整：缺歷史相對倍數）。
- 現價隱含：每股淨現金 {f2(g(f"B{VR['cash']}"))} 美元；營運業務 EV {f0(g(f"B{VR['mkt']}"))} 百萬美元（市值 − 淨現金）；小馬現行 EV／NTM 營收 ÷ 同業中位數 {g(f"B{VR['evs2']}"):.2f}。
- 現價檢核：csv 2026-11-20 收盤 3.31 美元；同檔 2026.HK 收盤 26.78 港元（約 3.41 美元，1 ADS＝1 股）；investing.com 港股 52 週區間 25.96–136.90 港元；Transport Topics 稱 ADR 年初至今 −72%。量級一致，照 csv 建模。

## 營收分部與標準損益（情境 B；US$M）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 期末車隊（台） | {yr('B','fl','{:,.0f}',YS)} |
| 每車每日付費英里 | {yr('B','pmd','{:,.0f}',YS)} |
| 每付費英里價格（US$） | {yr('B','prc','{:.3f}',YS)} |
| 每付費英里單位營運成本（US$） | {yr('B','opx','{:.3f}',YS)} |
| 滲透率（中國網約車趟數） | {yr('B','pen','{:.2%}',YS)} |
| Robotaxi 服務 | {yr('B','rt','{:,.1f}',YS)} |
| Robotruck 服務 | {yr('B','rk','{:,.1f}',YS)} |
| 智慧方案 | {yr('B','ls','{:,.1f}',YS)} |
| 營收 | {yr('B','rev','{:,.1f}',YS)} |
| 毛利率 | {yr('B','gm','{:.1%}',YS)} |
| 營業利益 | {yr('B','op','{:,.1f}',YS)} |
| 歸屬母公司淨利 | {yr('B','ni','{:,.1f}',YS)} |
| EPS（美元） | {yr('B','eps','{:.2f}',YS)} |
| FCFF | {yr('B','fcff','{:,.1f}',YS)} |
| 期末流動資金（未計新融資） | {yr('B','liq','{:,.0f}',YS)} |

- FY2025 實績；FY2026＝9M 實績＋第四季錨點（Robotaxi＝公司目標 3.5 倍，其餘分部第四季持平 Q3）；FY2027 起為模型。FY2023 只取得合併數（見 xlsx PL 頁）。
- 命題模組：Robotaxi 車資＝平均車隊 × 每車日付費英里（線 1a 中國基準 × 稼動率乘數）× 365 × 每英里價格（線基準 × 價格乘數）；其他（聯合部署車輛銷售、合作收入）＝車資 ×(k_rt − 1)，k_rt 以 FY2026 公司目標校準（{inv('k_rt'):.3f}）。
- 營業成本：Robotaxi＝付費英里 × 線單位營運成本 × 情境成本乘數 × cr_k（FY2026 合併毛利殘差校準 {inv('cr_k'):.3f}）；Robotruck、智慧方案用假設毛利率（分部毛利未揭露）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 車隊：FY2026 期末 3,500（公司目標）；FY2027 8,000／6,500／4,500；FY2028–30 年增 70%／50%／25%；FY2031–35 年增 30%／20%／8%（情境 A／B／C；全為 Assumed，公司未給 FY2027 以後目標）。
- 每英里經濟（線 1a 中國＝百度／小馬平均）：稼動率乘數 {inv('um_A'):.2f}／1.00／{inv('um_C'):.2f}；價格乘數 1.00／1.00／0.85；成本乘數 0.75／1.00／1.40。
- 費用：研發年增 0%／3%／6%；管銷 5% 變動＋固定 3% 年增；稅率 15%（只在稅前為正時）。
- 資本支出：自有車隊比例 60%（30–100%，未揭露）× 新增與汰換車輛 × 線單車成本＋其他 20 百萬美元／年。
- 折現：rf 5.00%（美國 10 年期，與 Tesla 模型同值，逾期）、β {inv('beta'):.3f}（中國自駕同業與全球自駕叫車兩來源 Blume 中位數）、ERP 5%、中國 CRP 0.9%（Assumed）、g＝3%、RONIC＝WACC。
- 評價日股數：Q3'26 加權 434.42 百萬股 ×（1＋選擇權與 RSU 淨稀釋 2%，未取得明細）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 車隊路徑（f27、fg、fglt） | 公司季度車隊揭露（2026-06 1,975 → 09 2,768 → 年底目標 3,500）、城市數、海外合作台數 |
| 價格乘數（pm） | 廣州、深圳 Robotaxi 牌價與補貼；每車每日營收（Q 頁 Q3'26 實績） |
| 營運成本乘數（om）與 cr_k | 合併毛利率（Q3'26 17.8%）、Gen-7 單車保費（人駕計程車約 40%）、遠端監看比 |
| 稼動率乘數（um） | 每車每日付費趟數（公司 23 趟）、全無人營運時數 |
| 自有車隊比例（own_r） | 資本支出 ÷ 車隊淨增；產品收入（聯合部署車輛銷售） |

失效條件：{RD['A7'].value[len('失效條件（可觀察）：'):]}。評等翻轉需要的條件見下方反解（Sens 各案例評等：{'、'.join(rts)}）。

**結論主要建立在 Assumed 輸入上：** 車隊路徑、價格與成本乘數、分部毛利、資本支出比例皆為 Assumed；每英里價格、稼動率與成本讀線 1a（線本身標 Assumed／Derived）。Verified 的只有歷史財報、流動資金與股數。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算，格點內插後驗證；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線 1a；xlsx「對帳」頁）
讀入：`RT_PMD_CN_*`（每車每日付費英里）、`RT_OPPRICE_CN_*`（每英里價格）、`RT_OPEX_CN_2026–2032`（單位營運成本）、成本輸入頁 `CN_VEH`（車價，不在下游取數頁）、`CN_LIFE`／`CN_PAID`（車輛壽命換算）；對照未用：`RT_CPM_CN_*`（全成本含線的研發分攤）、`RT_YEAR_1USD_CN`。
不一致處：(1) 小馬每英里營收＝線價格 × k_rt（FY2026 校準 {inv('k_rt'):.3f}；Q3'26 每車每日 Robotaxi 營收 {wb['Q']['F17'].value:.1f} 美元 vs 線基準車資 {wb['Q']['F18'].value:.1f} 美元），差額歸為聯合部署車輛銷售與合作收入；(2) 小馬營運成本水準以 cr_k {inv('cr_k'):.3f} 校準，只沿用線的下降斜率；(3) 線全成本的研發分攤（百度／小馬平均）不用，改用小馬自己的研發與管銷實績；(4) 線沒有中國滲透率，本模型自建。以上不在本模型靜默修改線值。

## 限制與未找到資料
- 分部毛利與營業利益：未揭露（未找到）；以 FY2026 合併毛利率殘差推 Robotaxi 成本。
- FY2023 分部營收、Q1／Q2'26 單季分部：未取得（中報只揭露上半年）。2025 年底精確車隊：未取得（公司稱超過 1,400 台）。
- 選擇權與 RSU 數量、履約價：未取得；折舊攤銷：財報新聞稿未揭露，以資本支出五年平均推算。
- 公司未給 FY2027 以後車隊或營收目標（未找到）；共識逐年營收只取得 FY2026（財報前）。
- 同業倍數頁面日期由 2026-10-02 到 2026-11-20 不一，市值依 csv 收盤等比例換算到 2026-11-20（Derived），淨現金與 NTM 營收未更新；地平線只有 TTM，不進中位數。
- 中國 CRP 未回原文核對（Assumed）。rf 取 2026-10-23 值（逾期，跨模型一致）。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）；display_check.json 未建。

## 核對（建置時執行）
{chk}

## 檔案
- `build_pony.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`，並跑核對。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("PONY_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    px0 = in_val("px", "base")
    ovs = [("base", {})]
    beta_b = sorted([0.67 * 2.01 + 0.33, 0.67 * 1.31 + 0.33]); beta_b = sum(beta_b) / 2
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
    RV = [("fg_B", 0.50, [i / 10 for i in range(-2, 14)], "情境 B 的 FY2028–FY2030 車隊年增（其他不變，整本重算；情境 A／C 不動）"),
          ("pm_B", 1.0, [0.4 + 0.1 * i for i in range(17)], "情境 B 價格乘數（× 線每英里價格；FY2027 起；情境 A／C 不動）"),
          ("om_B", 1.0, [0.2 + 0.15 * i for i in range(16)], "情境 B 營運成本乘數（× 線單位營運成本；FY2027 起；情境 A／C 不動）")]
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
    bid = f"pony-{VER}-{fingerprint()}"
    tmp = tempfile.mkdtemp(prefix="ponyf_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="小馬智行 Pony.ai", ticker="PONY", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="US$M", source=Q425 + "；" + Q326)
                                      for n, l, v in (("Robotaxi services", "Robotaxi 服務營收", 16.607), ("Robotruck services", "Robotruck 服務營收", 40.601),
                                                      ("Intelligent solutions（原 Licensing and applications）", "智慧方案營收", 32.793))],
                   sublines=[dict(parent="Robotaxi services", labels=["Robotaxi：車資（車隊 × 每車日付費英里 × 每英里價格）", "Robotaxi：其他（聯合部署車輛銷售、合作收入；校準）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：Robotaxi 子列合計 − Robotaxi（應為 0）"],
                   note="PL 頁三情境各有分部核對列；FY2023 分部未取得（核對列空白）。Q 頁另有半年度與單季分部對帳。子列只在預測期（FY2026 起）"),
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
        msg = ("通過" if rp.returncode == 0 else "失敗") + "：" + (last[0] if last else "")
        results.append((os.path.basename(cmd[0]), msg[:300])); print(cmd[0], rp.returncode, "\n".join(last[:8]))
    readme(X, VR, PR, DR, bid, sens, revs, results)

if __name__ == "__main__":
    main()
