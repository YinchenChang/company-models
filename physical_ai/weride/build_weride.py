# 文遠知行 WeRide（WRD，Nasdaq ADS；港股 0800.HK）個股模型 v0.1 — AI實體應用 Project（線 1 Robotaxi）
# 用法：python3 physical_ai/weride/build_weride.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json，並跑核對
#       python3 physical_ai/weride/build_weride.py --raw OUT  → 只寫一本（讀環境變數 WRD_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度＝曆年（12 月決算，IFRS，報表幣別人民幣）。金額：以 6.7045 換算為百萬美元（US$M）；股數：百萬 ADS 等值；EPS、股價：美元。1 ADS＝3 股普通股。
# 結構：README、In、Hist、Q、PL（三情境：車隊 × 每英里經濟 → Robotaxi 分部 → Robotaxi 與其餘營收（公司只揭露 Robotaxi，其餘為殘差）→ 標準損益 → 現金流與資金缺口）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數一律讀線 1a 成品 physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx（「下游取數」頁；車價例外讀「成本輸入」頁，對帳頁註明）。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md（明確期延長到 FY2035，見 README）。
import sys, os, json, subprocess, shutil, hashlib, tempfile, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_文遠知行個股模型_{VER}.xlsx"
LINE_REL = "physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx"
LINE = os.path.join(ROOT, LINE_REL)
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2036)); COL = {y: "CDEFGHIJKLMNO"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … O＝FY2035
FC = [y for y in YRS if y >= 2027]                                                       # DCF 明確期 FY2027–FY2035

# ---------------- 出處 ----------------
Q326 = "公司 6-K 附件 99.1：2026 年第三季未經審計財務報告（IFRS，人民幣；2026-11-20）https://www.sec.gov/Archives/edgar/data/0001867729/000110465926130438/tm2630019d1_ex99-1.htm"
Q425 = "公司新聞稿：2025 年第四季及全年未經審計財務報告（2026-03-23）https://ir.weride.ai/news-releases/news-release-details/weride-reports-record-full-year-2025-revenue-rmb6846-million-90"
H126 = "公司 2026 年第二季新聞稿（2026-08）https://weride.gcs-web.com/news-releases/news-release-details/accelerating-european-expansion-through-proven-asset-light-model"
F20 = Q425
SA = "stockanalysis.com 統計與預測頁（資料商；讀取 2026-11-24，各頁報價日見各列）"
INV = "（未用）"
TT = "（未用）"
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

sec("A. FY2026 錨點（9M 實績＋第四季假設；公司只給車隊目標，未給營收指引）")
inp("fx", "匯率：人民幣兌美元（2026-09-30，聯準會 H.10；全部人民幣數字統一以此換算）", "RMB/US$", 6.55, 6.7045, 7.0, "Verified", Q326 + "：RMB6.7045＝US$1.00", "統一用單一匯率換算歷史與預估（年增率不受影響；FY2025 公司用 6.9931，絕對美元數與公司報的美元數差約 4%）")
inp("rt26", "FY2026 Robotaxi 營收（上半年＋Q3＋Q4；人民幣 90＋107.3＋125）", "US$M", "=(70+107.3+110)/fx", "=(90+107.3+125)/fx", "=(110+107.3+150)/fx", "Assumed",
    "Q3'26 Robotaxi 營收「逾 1 億人民幣、年增 204.0%」（" + Q326 + "）；Q3'25＝35.3 → Q3'26≈107.3（Derived）；上半年與 Q4 公司未揭露單季 Robotaxi", "上半年 90（70–110）依 H1 L4 營收 179.5 與 Robotaxi 占比推估；Q4 125（110–150）依車隊 +13% 與 Q3 動能；全為 Assumed")
inp("rk_q4", "FY2026 第四季其餘營收（總營收假設 330 − Robotaxi 125；人民幣）", "US$M", "=280/fx-110/fx", "=205/fx", "=400/fx-150/fx", "Assumed", "Q3'26 總營收 321.5、Q4'25 314.0（" + Q326 + "；" + Q425 + "）；Q4 總營收假設 330（280–400）", "其餘＝Robobus、Robosweeper、L2++／L3、AI 基礎設施等，公司未分部揭露，殘差")
inp("ls_q4", "（無第三分部；恆為 0）", "US$M", None, 0, None, "Derived", "公司未揭露 Robotaxi 以外分部營收，本模型只拆 Robotaxi 與殘差", "保留此列只為沿用範本結構")
inp("g_anc", "FY2026 實績相對錨點（兩分部同比例）", "%", -0.05, 0.0, 0.05, "Assumed", "基準＝9M 實績＋第四季假設", "±5%：第四季為唯一未實現的一季")
inp("gm26", "FY2026 毛利率（9M 實績）", "%", None, "=264.576/667.356", None, "Derived", Q326 + "：9M 毛利 264.576 ÷ 營收 667.356（人民幣百萬）＝ 39.6%", "Q3 單季 43.0%、FY2025 30.2%（毛利率高度依產品組合，波動大）")
inp("rd26", "FY2026 研發費用（IFRS，含股權報酬）", "US$M", None, "=(1262.916+450)/fx", None, "Derived", Q326 + "：9M 1,262.9（人民幣百萬）＋第四季假設 450（400–500；Q3 單季 465.3）", "Q3 股權報酬轉回 20.4，Q4 不再假設轉回；股權報酬不加回（DCF 標準規格）")
inp("sga26", "FY2026 銷售、管理及其他淨額（銷售＋管理 − 其他淨收益＋應收減損）", "US$M", None, "=(187.42+62.5)/fx", None, "Derived", Q326 + "：9M 銷售 78.379＋管理 165.998 − 其他淨收益 60.269＋減損 3.312＝187.42；第四季假設＝9M 季平均 62.5", "Q3 單季只有 5.7（股權報酬轉回、其他收益 37.7 所致），不具代表性，故 Q4 不沿用")
inp("nonop26", "FY2026 業外（匯兌、利息收入、公允價值、財務費用）", "US$M", None, "=(89.189+26.17)/fx", None, "Derived", Q326 + "：9M 稅前 −1,096.571 減營業損失 −1,185.760＝＋89.189；第四季假設＝Q3 單季＋26.17", "")
inp("nci26", "FY2026 少數股東損益", "US$M", None, 0, None, "Verified", Q326 + "：歸屬普通股股東淨損益＝期間淨損益", "")
inp("cx26", "FY2026 資本支出（人民幣 600；9M 450＋Q4 150）", "US$M", "=450/fx", "=600/fx", "=750/fx", "Assumed", "公司未揭露資本支出（" + Q326 + "）；不動產及設備 2025-12 的 378.8 → 2026-09 的 712.3，加折舊攤銷 174.1（9M）推得 9M 約 507 上限", "取 9M 約 450、Q4 約 150；區間 450–750")
inp("sh26", "FY2026 加權平均股數（ADS 等值；9M 淨損失 1,099.765 ÷ 每 ADS 損失 3.33）", "百萬 ADS", None, "=1099.765/3.33", None, "Derived", Q326 + "：每 ADS 損失人民幣 3.33；1 ADS＝3 股普通股", "每股損失 1.11 ×3＝3.33；四捨五入誤差約 ±0.3%")

sec("B. 命題模組一：車隊（滲透）— 期末營運 Robotaxi 車隊 × 情境成長")
inp("fl25", "期末 Robotaxi 車隊 FY2025", "台", 900, 1000, 1125, "Assumed", "2026-03-23 公告時全球 Robotaxi 車隊 1,125 台（" + Q425 + "）；2025 年底精確值未取得", "取 1,000（900–1,125）；只用於 FY2026 平均車隊與校準", "")
inp("fl26", "期末 Robotaxi 車隊 FY2026（公司 2026-03 目標 2,600 台）", "台", 2400, 2600, 3000, "Interested-party", Q425 + "：2026 年底預計 2,600 台（視市場與法規）；" + Q326 + "：2026-10-31 逾 2,300 台（Robotaxi 車隊季增 26%）", "低 2,400、高 3,000（Assumed）", CO)
for s, v, note in (("A", 6000, "樂觀：中東與歐洲（Uber 合作）落地加速"),
                   ("B", 4500, "基準：FY2026 淨增約 1,600 台 → FY2027 淨增約 1,900 台（Assumed）"),
                   ("C", 3200, "保守：擴張放緩（許可、融資、美中監管）")):
    inp(f"f27_{s}", f"情境 {s}：期末車隊 FY2027", "台", 3500 if s == "B" else None, v, 6000 if s == "B" else None, "Assumed", note, "公司未給 FY2027 目標（未找到）")
for s, v, lo, hi, note in (("A", 0.60, None, None, "樂觀"), ("B", 0.45, 0.25, 0.65, "基準：FY2027 年增約 73% 之後減速"), ("C", 0.22, None, None, "保守")):
    inp(f"fg_{s}", f"情境 {s}：車隊年增 FY2028–FY2030", "%/年", lo, v, hi, "Assumed", note, "可證偽：季度車隊揭露、城市數、海外城市")
for s, v, lo, hi in (("A", 0.28, None, None), ("B", 0.18, 0.08, 0.28), ("C", 0.07, None, None)):
    inp(f"fglt_{s}", f"情境 {s}：車隊年增 FY2031–FY2035", "%/年", lo, v, hi, "Assumed", "長期擴張；以滲透率檢查合理性", "")
inp("tpvd", "每車每日付費趟數（滲透率換算用）", "趟/日", None, 21, None, "Interested-party", Q326 + "：廣州 Robotaxi 每車每日訂單 Q3 平均 21、尖峰 33", "只用於滲透率，不進營收（營收用每英里口徑）", CO)
inp("mkt26", "中國網約車年趟數 FY2026（滲透率分母）", "百萬趟", None, "=924*12", None, "Derived", RH + "：2026 年 4 月 9.24 億單 × 12", "季節性未調整。注意：文遠 Robotaxi 營收逾九成在海外（Q3'26），本滲透率只是規模檢查，不是市占")
inp("mkt_g", "中國網約車趟數年增", "%/年", None, 0.03, None, "Assumed", "只用於滲透率分母", "")

sec("C. 命題模組二：每英里經濟（讀線 1a「下游取數」中國營運商＝百度／小馬平均（線 1a 成品原文）；本模型以情境乘數與 FY2026 校準套用到文遠）")
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
    inp(f"veh{y % 100}", f"單車取得成本 {y}（線基準，含自駕套件）", "US$/台", None, v, None, "Interested-party", f"線 1a 成本輸入頁 {ref}（不在下游取數頁；對帳頁註明）", "文遠 2027 車型整車 <23 萬元（Gasgoo 2026-05-21，經線 1a）", "營運商與媒體（立場：降本敘事）")
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
inp("cr_k", "校準：文遠 Robotaxi 營業成本 ÷ 線成本推算（FY2026）", "x", None,
    "=((rt26+h_rk_9m+rk_q4+h_ls_9m+ls_q4)*(1+g_anc)*(1-gm26)-(h_rk_9m+rk_q4)*(1+g_anc)*(1-gm_rk)-(h_ls_9m+ls_q4)*(1+g_anc)*(1-gm_ls))/(((fl25+fl26)/2)*pmd26*365*ox26/1000000)", None, "Derived",
    "＝（FY2026 合併營業成本 − 其餘營收 與（無第三分部）營業成本）÷（付費英里 × 線單位營運成本）", "小於 1：聯合部署車輛由夥伴持有、承擔部分營運成本；分部毛利未揭露，故以殘差推得")
inp("kdec", "校準係數衰減（0＝持續；1＝FY2027 起線性衰減、FY2029 回到 1）", "x", None, 0, 1, "Assumed", "個股模型原則第 10 節第 48 條：基準持續", "持續的理由：聯合部署模式（夥伴持有車輛、文遠賣車與收服務費）預期擴大（Uber 歐洲 2,000 台以上）")

sec("D. 其餘營收與費用（成長路徑）")
for s, v, lo, hi in (("A", 0.40, None, None), ("B", 0.25, 0.10, 0.40), ("C", 0.10, None, None)):
    inp(f"grk_{s}", f"情境 {s}：其餘營收年增 FY2027–FY2030", "%/年", lo, v, hi, "Assumed", "FY2025 +38.8%（536.6 對 386.0 人民幣百萬；Derived）、Q3'26 其餘營收約 214 對 135.7（+58%；Derived）", "其餘＝Robobus、Robosweeper、L2++／L3（累計交付逾 10 萬套）、AI 基礎設施；未分部揭露")
for s, v, lo, hi in (("A", 0.20, None, None), ("B", 0.12, 0.05, 0.20), ("C", 0.05, None, None)):
    inp(f"grkl_{s}", f"情境 {s}：其餘營收年增 FY2031–FY2035", "%/年", lo, v, hi, "Assumed", "", "")
for s in SC:
    inp(f"gls_{s}", f"情境 {s}：（無第三分部）年增 FY2027–FY2030", "%/年", None, 0, None, "Derived", "恆為 0；沿用範本結構", "")
for s in SC:
    inp(f"glsl_{s}", f"情境 {s}：（無第三分部）年增 FY2031–FY2035", "%/年", None, 0, None, "Derived", "恆為 0；沿用範本結構", "")
inp("gm_rk", "其餘營收毛利率", "%", 0.30, 0.40, 0.50, "Assumed", "分部毛利未揭露（未找到）", "合併毛利率 FY2025 30.2%、9M'26 39.6%、Q3'26 43.0%；Robotaxi 營運毛利率低於合併，其餘偏高")
inp("gm_ls", "（無第三分部）毛利率", "%", None, 0.40, None, "Derived", "恆不用", "")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", 0.03, 0.0, 0.06), ("C", 0.06, None, None)):
    inp(f"grd_{s}", f"情境 {s}：研發費用年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "9M'26 IFRS 研發 +31.4%（1,262.9 對 961.0 人民幣百萬）、FY2025 +25.7%；已放緩至低個位數屬 Assumed", "FY2026 起公司強調成本效率；若仍維持兩位數成長，目標價大幅下修（見 Sens）")
inp("sga_v", "銷售、管理及其他：變動部分 ÷ 營收", "%", 0.03, 0.05, 0.08, "Assumed", "其餘為固定部分（FY2026 減去變動部分後年增 3%）", "FY2025 ÷ 營收＝99.5%（含股權報酬 450）")
inp("yld", "流動資金收益率（利息收入 ÷ 期初流動資金）", "%", 0.02, 0.03, 0.04, "Derived", "9M'26 利息收入 137.1 人民幣百萬 ÷ 約 5,200 平均現金部位 × 4/3 ≈ 3.5%", "FY2025 172.3 ÷ 約 7,000 ≈ 2.5%；基準取整 3%")
inp("tax_w", "有效稅率（稅前為正時；虧損年不計稅盾）", "%", 0.15, 0.15, 0.25, "Assumed", "中國高新技術企業 15%；高案法定 25%", "累積虧損抵稅未建模（保守）")
inp("dil_g", "股數年增（股權報酬稀釋）", "%/年", 0.0, 0.01, 0.03, "Assumed", "股數變動未取得完整揭露", "未計新股融資（另見資金缺口）；公司 2026-03 宣布 1 億美元回購（Caixin），未計入")

sec("E. 現金流與資金缺口")
inp("own_r", "自有車隊比例（新增與汰換車輛中由文遠出資者）", "%", 0.3, 0.5, 1.0, "Assumed", "未揭露（未找到）；公司稱歐洲採輕資產模式（夥伴出車）", "不動產及設備 9M 淨增約 333 人民幣百萬對 Robotaxi 淨增約 1,300 台")
inp("cx_oth", "其他資本支出（研發設備、場站、其他車型）", "US$M/年", 10, 20, 35, "Assumed", "公司未揭露資本支出", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.0, 0.10, 0.20, "Assumed", "存貨與應收帳款 2026-09 為 501.3、588.5 人民幣百萬（2025-12：321.0、462.1）", "營運資金隨營收擴張")
inp("min_cash", "最低流動資金（資金缺口門檻）", "US$M", None, 250, None, "Assumed", "約一年淨現金消耗", "只用於資金缺口，不進目標價；公司稱現金可支撐逾四年（Interested-party，CFO 說法）")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤，WRD ADS）", "US$", None, 2.985, None, "Verified", YCL + "：WRD 2026-11-20 收盤 2.985", "同檔 0800.HK 收盤 8.44 港元（每股；1 ADS＝3 股；折約 3.2 美元，與 ADS 差約 8%，兩地報價時差與流動性所致）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("px_hk", "對照：港股 0800.HK 收盤（只作對照，不進模型）", "HK$", None, 8.44, None, "Verified", YCL + "：0800.HK 2026-11-20", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（ADS 等值）", "百萬 ADS", None, 330.04, None, "Interested-party", SA + "：WRD 統計頁 330.04 百萬（資料商；其他來源 310.7–342.4，來源分歧）", "與 9M 加權 330.2 吻合", DV)
inp("sh_dil", "選擇權與 RSU 淨稀釋（庫藏股法）", "%", 0.0, 0.02, 0.05, "Assumed", "未取得明細", "DCF 標準規格第 3 節要求計入；以區間代替")
inp("sh_v", "評價日股數（ADS 等值）", "百萬股", None, "=sh26*(1+sh_dil)", None, "Derived", "＝9M 加權 ADS 數 ×（1＋淨稀釋）", "")
inp("liq_v", "流動資金（現金與約當現金 573.7－含定期存款、FVTPL 短期理財 119.0、受限現金 2.8；2026-09-30）", "US$M", None, 695.5, None, "Verified", Q326 + "：現金、定期存款、理財產品與受限現金合計人民幣 4,663.0 百萬（US$695.5M）", "2025-12-31 為人民幣 7,432 百萬量級（現金與定期存款 6,967.7＋理財 144.3＋受限 19.4）；9M 流出約 2,770 人民幣百萬")
inp("inv_v", "其他投資（長期股權投資等）", "US$M", None, 0.0, None, "Assumed", "未取得資產負債表長期投資明細（未找到於本輪讀取的摘要）", "保守起見不計；影響上限＝每股數十美分")
inp("inv_k", "其他投資計入比例", "x", 0.0, 0.5, 1.0, "Assumed", "inv_v＝0，此係數無作用", "")
inp("lease_v", "債務與租賃（短期銀行借款 91.5＋租賃負債 23.1）", "US$M", None, 114.6, None, "Verified", Q326 + "：短期銀行借款人民幣 613.7 百萬、租賃負債 154.8 百萬（US$91.5M、23.1M）", "")
inp("mi_v", "非控制權益（2026-09-30）", "US$M", None, 0.0, None, "Verified", Q326 + "：歸屬普通股股東淨損益＝期間淨損益，無非控制權益損益", "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債）", "%", None, 0.05, None, "Verified", UST, "2026-10-23 值，距股價日逾 2 個交易日：逾期（跨模型共用，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", 0.006, 0.009, 0.012, "Assumed", "Damodaran 中國 CRP 約 0.8–1.0%（本輪未回原文核對：未查核，不作 Verified）", "營收逾半仍在中國；海外占比上升")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", 1.1425, "=MEDIAN(0.67*2.01+0.33,0.67*1.31+0.33)", 1.7425, "Derived",
    SA + "：來源一＝中國自駕同業 5 年 β 中位數 2.01（文遠 2.01、地平線 2.35、禾賽 1.47）；來源二＝全球自駕與叫車 1.31（Mobileye 1.24、Tesla 1.74、Uber 1.31）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）；與文遠模型同值，已隨文遠股價資料更新前沿用")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業 EV／營收）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺文遠對同業的歷史相對倍數）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業 EV／NTM 營收（2026-11-20 收盤換算；資料商頁面日期各異，市值依 csv 收盤等比例換算、淨現金與 NTM 營收不變；Interested-party：賣方共識營收）")
PEERS = [("ev_pony", "小馬智行 PONY（中國 Robotaxi）", "=(3.31*434.42199-1154-0.5*214.514+19.03+40.057)/(2230/8.54)", "2026-10-23 頁面：市值 22.3 億、Forward PS 8.54；淨現金口徑同本專案小馬智行個股模型（2026-09-30）；11-20 收盤 3.31（csv）"),
         ("ev_bidu", "百度 BIDU（Apollo Go；集團口徑）", "=(27030*61.9/81.87-7850)/(27030/1.38)", "2026-11-06 頁面：市值 270.3 億、淨現金 78.5 億、Forward PS 1.38、股價 81.87；11-20 收盤 61.90"),
         ("ev_uber", "Uber UBER（叫車平台）", "=133650/(129600/2.00)", "2026-11-20 頁面：EV 1,336.5 億、市值 1,296 億、Forward PS 2.00（同日，不換算）"),
         ("ev_tsla", "Tesla TSLA（Robotaxi、整車）", "=1340000/(1370000/11.55)", "2026-11-20 頁面：EV 1.34 兆、市值 1.37 兆、Forward PS 11.55（同日）"),
         ("ev_mbly", "Mobileye MBLY（自駕方案）", "=(5660*4.57/6.66-1360)/(5660/2.86)", "2026-10-26 頁面：市值 56.6 億、淨現金 13.6 億、Forward PS 2.86、股價 6.66；11-20 收盤 4.57"),
         ("ev_hsai", "禾賽 HSAI（光達）", "=953.45/(1850/2.44)", "2026-11-20 頁面：EV 9.5345 億、市值 18.5 億、Forward PS 2.44（同日）")]
for k, lab, f, note in PEERS:
    inp(k, f"EV／NTM 營收：{lab}", "x", None, f, None, "Derived", SA + "：" + note, "NTM 營收＝市值 ÷ Forward PS（資料商共識，Interested-party）；會計口徑差異（IFRS 與 US GAAP）未調整，租賃負債不計入 EV", DV)
inp("ev_med", "同業 EV／NTM 營收中位數（6 家）", "x", None, "=MEDIAN(ev_pony,ev_bidu,ev_uber,ev_tsla,ev_mbly,ev_hsai)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）", "Tesla 約 11 倍為離群值（中位數不受影響）。地平線 9660.HK 收盤檔有報價但無 Forward 營收，未納入；其餘同業報價日皆為 2026-11-20 收盤或依該收盤換算")
inp("ev_self", "文遠現行 EV／NTM 營收（只作對照，不進模型）", "x", None, "=(px*sh_v-liq_v-inv_v*inv_k+lease_v+mi_v)/(1360/6.32)", None, "Derived", SA + "：2026-11-13 頁面市值 13.6 億、Forward PS 6.32 → NTM 營收約 2.15 億（Q3 前共識）", "EV 以本模型淨現金口徑", DV)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "US$M", None, "=1060/fx", None, "Interested-party", "Yahoo Finance 分析頁（2026-11 搜尋摘要）：FY2026 營收共識約人民幣 10.6 億；FY2027 約 20.2 億", "未回原頁核對，只作對照", DV)


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
           ("om", "營運成本乘數"), ("grk", "其餘營收 年增 FY2027–30"), ("grkl", "其餘營收 年增 FY2031–35"), ("gls", "（無第三分部）年增 FY2027–30"),
           ("glsl", "（無第三分部）年增 FY2031–35"), ("grd", "研發費用年增")]
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
    H["A1"] = "Hist — 歷史財務（年度，IFRS，原幣人民幣百萬 ÷ 6.7045 換算為 US$M；FY2023 未取得）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "9M'26", "標記", "來源"])
    NA = "未取得"
    X = lambda v: round(v / 6.7045, 3)
    HR = [("h_rt", "Robotaxi 營收", NA, X(47.8), X(148.0), X(197.3), "FY2025 Verified；FY2024 Derived（148.0 ÷ 3.096）；9M'26 Assumed（上半年 90＋Q3 107.3）", Q425 + "；" + Q326),
          ("h_rk", "其餘營收（Robobus、Robosweeper、L2++／L3、AI 基礎設施等；總營收減 Robotaxi）", NA, X(361.1 - 47.8), X(684.6 - 148.0), X(667.356 - 197.3), "Derived", "總營收減 Robotaxi；公司未分部揭露"),
          ("h_ls", "（無第三分部；恆為 0）", NA, 0, 0, 0, "Derived", "沿用範本結構"),
          ("h_rev", "營收", NA, X(361.1), X(684.6), X(667.356), "Verified", Q425 + "；" + Q326),
          ("h_cogs", "營業成本", NA, X(250.4), X(477.8), X(402.780), "Verified", "同上"),
          ("h_rd", "研發費用", NA, X(1091.4), X(1372.2), X(1262.916), "Verified", "同上（FY2024 含上市前後股權報酬 1,187.9 人民幣百萬）"),
          ("h_sga", "銷售、管理及其他淨額（銷售＋管理 − 其他淨收益＋應收減損）", NA, X(53.6 + 1138.8 - 16.5 + 28.7), X(73.6 + 596.1 - 4.9 + 16.6), X(187.420), "Derived", "同上拆項合計"),
          ("h_op", "營業利益（虧損）", NA, X(-2185.2), X(-1846.8), X(-1185.760), "Verified", "同上"),
          ("h_pbt", "稅前利益（虧損）", NA, X(-2510.9), X(-1650.8), X(-1096.571), "Verified", "同上（FY2024 含優先股公允價值變動 −465.3 人民幣百萬）"),
          ("h_tax", "所得稅", NA, X(-5.9), X(-4.1), X(-3.194), "Verified", "同上"),
          ("h_ni", "淨利（虧損）", NA, X(-2516.8), X(-1654.9), X(-1099.765), "Verified", "同上"),
          ("h_nci", "少數股東損益", NA, 0, 0, 0, "Verified", "歸屬普通股股東淨損益＝期間淨損益"),
          ("h_nia", "歸屬母公司淨利（虧損）", NA, X(-2516.8), X(-1654.9), X(-1099.765), "Verified", "同上"),
          ("h_sh", "加權平均股數（ADS 等值，淨損失 ÷ 每 ADS 損失）", NA, round(2516.8 / 25.62, 2), round(1654.9 / 5.37, 2), round(1099.765 / 3.33, 2), "Derived", "每 ADS 損失 25.62／5.37／3.33 人民幣；1 ADS＝3 股；FY2024 受 2025 年上市前股數變動影響，不具比較性"),
          ("h_cx", "資本支出", NA, NA, NA, NA, "未取得", "公司未揭露（見 In 頁 cx26 推估）"),
          ("h_ocf", "營業現金流", NA, NA, NA, NA, "未取得", "摘要未列現金流量表"),
          ("h_sbc", "股權報酬", NA, X(1187.9), X(450.0), X(98.848), "Verified", Q425 + "；" + Q326)]
    HROW = {}
    for i, (k, lab, a_, b_, c_, d_, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a_, b_, c_, d_)):
            x = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): x.font = BLUE
        H.cell(i, 6, tag); H.cell(i, 7, src)
        for y, L in (("23", "B"), ("24", "C"), ("25", "D"), ("9m", "E")): dn(f"{k}_{y}", f"Hist!${L}${i}")
        HROW[k] = i
    i = 3 + len(HR)
    H.cell(i, 1, "檢查：分部合計 − 營收（FY2024–9M'26）")
    for L in "CDE": H.cell(i, "BCDE".index(L) + 2, f"=ROUND({L}3+{L}4+{L}5-{L}6,6)")
    H.cell(i + 1, 1, "FY2023 全部未取得（本輪只取得 FY2024 起的比較欄）。FY2024 FY2025 數字為公司以 2025-12-31 或 2026 年新聞稿列示之原幣數，以單一匯率換算")
    H.column_dimensions["A"].width = 40; H.column_dimensions["G"].width = 80

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度營收（人民幣百萬；原文 Verified）與 Robotaxi 營運指標。Robotaxi 單季只取得 Q4'25、Q3'25 與 Q3'26（年增 204.0%，Derived）"
    hdr(Q, 2, ["項目", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26", "標記", "來源", "來源誘因"])
    rowsq = [("總營收（人民幣百萬）", 170.982, 314.0, 114.2, 231.7, 321.499, "Verified", Q326 + "；" + Q425 + "；" + H126),
             ("Robotaxi 營收（人民幣百萬）", 35.3, 50.6, None, None, 107.3, "Derived", Q326 + "；" + Q425),
             ("毛利率", 0.329, 0.285, None, 0.375, 0.43, "Verified", Q326 + "；" + Q425 + "；" + H126)]
    for i, row in enumerate(rowsq, 3):
        Q.cell(i, 1, row[0])
        for j, v in enumerate(row[1:6]):
            if v is not None: Q.cell(i, 2 + j, v).font = BLUE
        Q.cell(i, 7, row[6]); Q.cell(i, 8, row[7] + ("（Q1＝H1 345.857 減 Q2；Q3'26 Robotaxi＝35.3×3.04 推得；Robotaxi Q1、Q2 未取得）" if i < 5 else "")); Q.cell(i, 9, CO)
    Q.cell(6, 1, "Robotaxi 占營收"); [Q.cell(6, 2 + j, f"={'BCDEF'[j]}4/{'BCDEF'[j]}3") for j in (0, 1, 4)]
    Q.cell(7, 1, "期末 Robotaxi 車隊（台；Q3'26 以 10-31 逾 2,300 近似）"); Q.cell(7, 3, "=fl25"); Q.cell(7, 6, 2300)
    Q.cell(7, 7, "Interested-party"); Q.cell(7, 8, Q326); Q.cell(7, 9, CO)
    Q.cell(8, 1, "Q3'26 每車每日 Robotaxi 營收（US$；Q3 車隊以 2,300 與 Q2 末約 1,830 平均，Derived）"); Q.cell(8, 6, "=F4/fx*1000000/((1830+F7)/2)/92")
    Q.cell(9, 1, "線 1a 基準推算每車每日車資（US$；每車日付費英里 × 每英里價格）"); Q.cell(9, 6, "=pmd26*prc26")
    Q.cell(10, 1, "Q3'26 毛利率（核對）"); Q.cell(10, 6, "=138.103/321.499")
    Q.cell(11, 1, "Q3'26：海外營收 146.9（占 45.7%）；Robotaxi 營收逾九成在海外；L4 車隊約 4,100 台（Robotaxi 逾 2,300）；中東約 600 台；Q3 股權報酬轉回 20.4 人民幣百萬")
    Q.column_dimensions["A"].width = 52

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境（US$M）：車隊 × 每車日付費英里 × 每英里價格（讀線 1a）→ Robotaxi 分部 → Robotaxi 與其餘營收（公司只揭露 Robotaxi，其餘為殘差）→ 標準損益表（IFRS，人民幣以單一匯率換為美元）→ 現金流與資金缺口。FY2023–FY2025 實績、FY2026 為 9M 實績＋第四季錨點、FY2027 起模型"
    PR = {}; rr = 3
    KEYS = [("fl", "期末 Robotaxi 車隊（台）"), ("fla", "平均 Robotaxi 車隊（台）"), ("pmd", "每車每日付費英里（英里）"), ("mi", "年付費英里（百萬英里）"),
            ("prc", "每付費英里價格（US$）"), ("trips", "Robotaxi 年趟數（百萬趟）"), ("mkt", "中國網約車年趟數（百萬趟）"), ("pen", "滲透率：文遠趟數 ÷ 中國網約車趟數"),
            ("rtf", "Robotaxi：車資（車隊 × 每車日付費英里 × 每英里價格）"), ("rto", "Robotaxi：其他（聯合部署車輛銷售、合作收入；校準）"),
            ("rt", "Robotaxi 服務營收"), ("rk", "其餘營收（殘差）"), ("ls", "（無第三分部）營收"), ("rev", "營收"),
            ("chk1", "檢查：Robotaxi 子列合計 − Robotaxi（應為 0）"), ("chk2", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"),
            ("opx", "每付費英里單位營運成本（線 × 情境乘數 × 校準；US$）"), ("rpm", "每付費英里 Robotaxi 營收（US$）"), ("gpm", "每付費英里 Robotaxi 毛利（US$）"),
            ("cgrt", "營業成本：Robotaxi"), ("cgrk", "營業成本：其餘營收"), ("cgls", "營業成本：（無第三分部）"), ("cogs", "營業成本"),
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
            if y == 2023:
                for k, _l in KEYS:
                    if k not in ("opf",): w(k, "未取得")
                w("opf", 9999)
                continue
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
                if y >= 2025: w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
                w("opex", f"={cell('rd')}+{cell('sga')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('nit')}/{cell('rev')}")
                if y >= 2025: w("nig", f"=IF({prev('nit')}<0,\"虧損\",{cell('nit')}/{prev('nit')}-1)")
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
    r_pen30 = vr; V.cell(vr, 1, "滲透率 FY2030（文遠趟數 ÷ 中國網約車趟數；情境 A／B／C）")
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
    r_evs2 = one("現行 EV／NTM 營收（文遠）÷ 同業中位數", "=ev_self/ev_med", "只作對照")
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
    rc = [("FY2026 Robotaxi 營收 vs 模型錨點", f"=PL!F{B_['rt']}", "=rt26", "錨點為 Assumed（公司未給營收指引）"),
          ("FY2026 期末車隊 vs 公司目標", f"=PL!F{B_['fl']}", 2600, "公司 2026-03 目標 2,600 台（Interested-party）"),
          ("FY2026 營收 vs 共識", f"=PL!F{B_['rev']}", "=cons_rev26", "Yahoo 共識（搜尋摘要，只作對照）"),
          ("每車每日 Robotaxi 營收：Q3'26 推估 vs 線基準車資推算（US$）", "=Q!F8", "=Q!F9", "實績含聯合部署車輛銷售與合作收入；比值即 k_rt 的季度對照（口徑不同，只看方向）"),
          ("每付費英里單位營運成本 FY2026：模型殘差推得 vs 線基準（US$）", f"=PL!F{B_['opx']}", "=ox26", "比值＝cr_k；分部毛利未揭露，模型值為殘差"),
          ("每付費英里單位營運成本 FY2030：模型（線 × 校準）vs 線基準（US$）", f"=PL!J{B_['opx']}", "=ox30", "線為百度／小馬平均（線 1a 成品原文）"),
          ("加權目標價 vs 現價（US$）", "=tp", "=px", "上檔空間見 Val 頁；賣方目標價本輪未查")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_weride.py 格點＋線性內插後驗證；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_weride.py 以 LibreOffice 重算每一案例）；結果為常數快照。估值評等欄只看上檔空間"
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
    SE.cell(sr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_weride.py")
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
        rec.append((f"中國每付費英里價格 {y}", f"RT_OPPRICE_CN_{y}（{ref}）", v, f"{lo}／{hi}", "基準 × 價格乘數 × 校準 k_rt（含聯合部署車輛銷售）", "不一致：文遠每英里營收＝線價格 × k_rt（FY2026 校準，見 Recon）", "Assumed", f"In prc{y % 100} → PL 車資列"))
    for y in range(2026, 2033):
        v, ref = lv(f"RT_OPEX_CN_{y}"); lo, _ = lv(f"RT_OPEX_CN_{y}", "C"); hi, _ = lv(f"RT_OPEX_CN_{y}", "E")
        rec.append((f"中國每付費英里單位營運成本 {y}", f"RT_OPEX_CN_{y}（{ref}）", round(v, 4), f"{lo:.3f}／{hi:.3f}", "線基準 × 情境成本乘數 × 校準 cr_k", "不一致：文遠成本水準以 FY2026 殘差校準（cr_k），只沿用線的下降斜率", "Derived", f"In ox{y % 100} → PL 營業成本：Robotaxi"))
    for y in ANC:
        v, ref = cv("CN_VEH", y)
        rec.append((f"中國單車取得成本 {y}", f"CN_VEH（{ref}）", v, "—", "自有車隊資本支出", "一致；但不在下游取數頁（成本輸入頁），下次線改版時請列入下游取數", "Interested-party", f"In veh{y % 100} → PL 資本支出"))
    for i_, lab in (("RT_CPM_CN_2026", "中國每付費英里全成本 2026（含研發攤提、平台）"), ("RT_CPM_CN_2030", "中國每付費英里全成本 2030"), ("RT_CPM_CN_2032", "中國每付費英里全成本 2032")):
        v, ref = lv(i_); lo, _ = lv(i_, "C"); hi, _ = lv(i_, "E")
        rec.append((lab, f"{i_}（{ref}）", round(v, 4), f"{lo:.3f}／{hi:.3f}", "未用：本模型研發與管銷用文遠實績（FY2026 研發約 256），不用線的分攤（百度／小馬平均（線 1a 成品原文） 5 億美元／年、車隊 5,000 台）", "口徑不同：線全成本假設的研發分攤與車隊規模不是文遠自己的", "Derived", "只作對照"))
    v, ref = lv("RT_YEAR_1USD_CN"); rec.append(("中國全成本跨 1 美元年份", f"RT_YEAR_1USD_CN（{ref}）", v, "—", "未用；本模型對應指標＝營業利益轉正年份（Val 頁）", "口徑不同", "Derived", "只作對照"))
    rec.append(("中國滲透率", "線 1a 無中國滲透率（只有美國 RT_PEN_*）", "—", "—", "本模型自建：文遠趟數 ÷ 中國網約車趟數（網約車監管信息交互系統）", "線未提供（缺）", "Derived", "PL 滲透率列"))
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
             ("匯率（人民幣收入與美元股價；海外收入占比上升）", "部分（fx 區間）", "In 頁 fx 敏感度", "人民幣兌美元"),
             ("海外合作（Uber 中東與歐洲、GreenMobility、新加坡、香港）落地時點", "部分（f27_A、fg_A）", "情境 A", "海外城市營運公告"),
             ("雙重股權結構與創辦人投票權；Uber、宇通、雷諾日產三菱聯盟為主要股東", "否", "治理", "股東會")]
    for i_, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i_, 1 + j, v)
    i_ = 3 + len(risks) + 1
    RK.cell(i_, 1, "未找到／未取得清單（「未找到」＝查過沒有公開資料；「未取得」＝有資料但本輪未取得原文）").font = BOLD; i_ += 1
    hdr(RK, i_, ["項目", "狀態", "替代做法", "下次怎麼補"]); i_ += 1
    miss = [("分部毛利與分部營業利益", "未找到（公司未揭露）", "FY2026 合併毛利率殘差校準 cr_k；其餘營收、（無第三分部）毛利率 Assumed", "20-F 分部附註（公司未見揭露）"),
            ("FY2023 全部財務數字、資本支出、現金流量", "未取得", "PL 頁 FY2023 欄標未取得", "FY2024 20-F 或 F-1"),
            ("Q1、Q2'26 單季 Robotaxi 營收", "未取得", "上半年合計 Assumed 90 人民幣百萬（70–110）", "Q1、Q2 新聞稿"),
            ("2025 年底精確車隊", "未取得", "2026-03 公告 1,125 台；取 1,000（900–1,125）", "20-F 營運數據"),
            ("選擇權與 RSU 數量、加權履約價", "未取得", "淨稀釋 0–5% 區間（基準 2%）", "20-F 股權報酬附註"),
            ("折舊攤銷", "未取得", "資本支出五年平均攤提推算", "20-F 現金流量表"),
            ("FY2027 以後車隊或營收目標", "未找到", "車隊路徑 Assumed（三情境）", "法說會、投資人日"),
            ("賣方目標價與逐年共識營收", "未取得", "只有 Yahoo 摘要的 FY2026、FY2027 共識", "資料商"),
            ("中國國家風險溢酬原文", "未取得", "Assumed 0.9%（0.6–1.2%）", "Damodaran 國家風險表"),
            ("自有車隊比例（聯合部署夥伴持有車輛占比）", "未找到", "Assumed 50%（30–100%）", "公司揭露車隊持有結構")]
    for row in miss:
        for j, v in enumerate(row): RK.cell(i_, 1 + j, v)
        i_ += 1
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i_, s_ in enumerate([Q326, H126, Q425, F20, SA, INV, TT, RH, YCL, UST, LN + "（下游取數、成本輸入、輸入頁）"], 1):
        SO.cell(i_, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"文遠知行 WeRide（WRD；港股 0800.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（Q3 財報 2026-11-20）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：文遠的價值取決於 Robotaxi 車隊的擴張速度（營收逾九成在海外，以中東、新加坡為主）——車隊規模 × 每車每日付費英里 × 每英里價格——以及每英里單位營運成本能降到多低；短期收支平衡不是主軸，但資金缺口決定能否撐到規模化"
    RD["A4"] = (f'="在情境 B 車隊 FY2027 "&TEXT(f27_B,"#,##0")&" 台、FY2028–30 年增 "&TEXT(fg_B,"0%")&"、每英里價格與成本照線 1a 基準（FY2026 校準）的假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = (f'="資金缺口（情境 B，未計新融資）：最低期末流動資金 "&TEXT(Val!C{r_liqmin},"#,##0")&" 百萬美元、最大缺口 "&TEXT(Val!C{r_gapmax},"#,##0")&"（首次缺口 "&Val!C{r_gapyr}&"）；營業利益轉正 "&Val!C{r_opy}&"；滲透率 FY2030 "&TEXT(Val!C{r_pen30},"0.00%")&"、FY2035 "&TEXT(Val!C{r_pen35},"0.00%")')
    RD["A7"] = "失效條件（可觀察）：期末車隊連兩季低於情境 B 路徑；Robotaxi 每車每日營收（Q3'26 實績見 Q 頁）連兩季下降；毛利率跌破 30%；流動資金季減幅擴大到每季 1 億美元以上"
    RD["A8"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics；線 1a 只提供中國每英里價格、稼動率與成本（百度／小馬平均（線 1a 成品原文）），不提供中國滲透率，本模型自建。其餘營收、（無第三分部）以成長路徑建模"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), "lo_profile_weride")   # LibreOffice 隔離設定檔
def recalc_many(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 20):
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + paths[i:i + 20],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=170)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile_weride"], stderr=subprocess.DEVNULL)
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
    d = tempfile.mkdtemp(prefix="weride_"); paths = []; meta = {}
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
        if row[0] == k:
            v = row[{"lo": 3, "base": 4, "hi": 5}[col]]
            if isinstance(v, str) and v.startswith("=") and k in ("fx", "rt26", "rk_q4", "cx26"): v = eval(v[1:], {"fx": 6.7045})
            return v
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("fx", "匯率（人民幣兌美元）"), ("rt26", "FY2026 Robotaxi 營收錨點"), ("rk_q4", "FY2026 第四季其餘營收"), ("cx26", "FY2026 資本支出"), ("g_anc", "FY2026 實績相對錨點"), ("fl25", "FY2025 期末車隊（校準錨點）"), ("fl26", "FY2026 期末車隊"),
         ("f27_B", "FY2027 期末車隊（情境 B）"), ("fg_B", "車隊年增 FY2028–30（情境 B）"), ("fglt_B", "車隊年增 FY2031–35（情境 B）"),
         ("um_B", "稼動率乘數（情境 B）"), ("pm_B", "價格乘數（情境 B）"), ("om_B", "營運成本乘數（情境 B）"), ("kdec", "校準係數衰減（FY2029 回到 1）"),
         ("veh_life", "車輛經濟壽命"), ("grk_B", "其餘營收 年增 FY2027–30（情境 B）"), ("grkl_B", "其餘營收 年增 FY2031–35（情境 B）"),
         ("gls_B", "（無第三分部）年增 FY2027–30（情境 B）"), ("glsl_B", "（無第三分部）年增 FY2031–35（情境 B）"), ("gm_rk", "其餘營收 毛利率（同時重算 cr_k）"),
         ("gm_ls", "（無第三分部）毛利率（同時重算 cr_k）"), ("grd_B", "研發費用年增（情境 B）"), ("sga_v", "管銷變動部分 ÷ 營收"), ("yld", "流動資金收益率"),
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
    txt = f"""# physical_ai/weride — 文遠知行 WeRide（WRD；港股 0800.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

Nasdaq ADS（1 ADS＝3 股普通股，美元）；港交所 0800.HK（2025-11 雙重主要上市）；財報為 IFRS、人民幣，本模型以 6.7045 一律換算成美元。線 1 Robotaxi（線 1a 成品 `{LINE_REL}`）。本 README 的數字由 `build_weride.py` 從重算後的成品讀出產生，不手寫。

**範圍限制（高不確定性原則）：** 文遠仍在大幅虧損、分部毛利與營業利益未揭露、只有 FY2024 起的比較欄。本模型做到：分部營收（Robotaxi 與殘差其餘營收兩部分；公司只揭露 Robotaxi 單項）由驅動因子推導、標準損益表、FCFF DCF 與同業 EV／營收；**沒有做到**：分部別毛利（公司未揭露，以 FY2026 合併毛利殘差校準）、FY2023 全部財務數字（未取得）、Robobus／Robosweeper／L2++／AI 基礎設施各自營收（公司未揭露）、FY2026 上半年與 Q4 的 Robotaxi 單季營收（Assumed）、資本支出（未揭露，Assumed）、選擇權與 RSU 明細（未取得，以 0–5% 區間代替）、新股融資與稀釋（只列資金缺口，不假設融資條件）、累積虧損抵稅。

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
- 現價隱含：每股淨現金 {f2(g(f"B{VR['cash']}"))} 美元；營運業務 EV {f0(g(f"B{VR['mkt']}"))} 百萬美元（市值 − 淨現金）；文遠現行 EV／NTM 營收 ÷ 同業中位數 {g(f"B{VR['evs2']}"):.2f}。
- 現價檢核：csv 2026-11-20 收盤 2.985 美元；同檔 0800.HK 收盤 8.44 港元（每股；1 ADS＝3 股，約 3.2 美元），兩地差約 8%，照 ADS 的 csv 建模。

## 營收分部與標準損益（情境 B；US$M）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 期末車隊（台） | {yr('B','fl','{:,.0f}',YS)} |
| 每車每日付費英里 | {yr('B','pmd','{:,.0f}',YS)} |
| 每付費英里價格（US$） | {yr('B','prc','{:.3f}',YS)} |
| 每付費英里單位營運成本（US$） | {yr('B','opx','{:.3f}',YS)} |
| 滲透率（中國網約車趟數） | {yr('B','pen','{:.2%}',YS)} |
| Robotaxi 服務 | {yr('B','rt','{:,.1f}',YS)} |
| 其餘營收（殘差） | {yr('B','rk','{:,.1f}',YS)} |
| 營收 | {yr('B','rev','{:,.1f}',YS)} |
| 毛利率 | {yr('B','gm','{:.1%}',YS)} |
| 營業利益 | {yr('B','op','{:,.1f}',YS)} |
| 歸屬母公司淨利 | {yr('B','ni','{:,.1f}',YS)} |
| EPS（美元） | {yr('B','eps','{:.2f}',YS)} |
| FCFF | {yr('B','fcff','{:,.1f}',YS)} |
| 期末流動資金（未計新融資） | {yr('B','liq','{:,.0f}',YS)} |

- FY2025 實績；FY2026＝9M 實績＋第四季假設（Robotaxi 全年人民幣約 322 百萬，其中上半年與 Q4 為 Assumed；其餘營收 Q4 假設 205）；FY2027 起為模型。FY2023 未取得；PL 頁 FY2023 欄標「未取得」。PL 頁另有一列「（無第三分部）」恆為 0，只為沿用範本結構。
- 命題模組：Robotaxi 車資＝平均車隊 × 每車日付費英里（線 1a 中國基準 × 稼動率乘數）× 365 × 每英里價格（線基準 × 價格乘數）；其他（車輛銷售、合作與海外夥伴收入）＝車資 ×(k_rt − 1)，k_rt 以 FY2026 Robotaxi 營收假設校準；**注意：Q3'26 Robotaxi 營收逾九成在海外（中東、新加坡等），線 1a 中國每英里價格與成本只是代理，k_rt 與 cr_k 吸收差異，結論對此高度敏感**（{inv('k_rt'):.3f}）。
- 營業成本：Robotaxi＝付費英里 × 線單位營運成本 × 情境成本乘數 × cr_k（FY2026 合併毛利殘差校準 {inv('cr_k'):.3f}）；其餘營收、（無第三分部）用假設毛利率（分部毛利未揭露）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 車隊：FY2026 期末 2,600（公司 2026-03 目標；2026-10-31 逾 2,300）；FY2027 6,000／4,500／3,200；FY2028–30 年增 60%／45%／22%；FY2031–35 年增 28%／18%／7%（情境 A／B／C；全為 Assumed，公司未給 FY2027 以後目標）。
- 每英里經濟（線 1a 中國＝百度／小馬平均（線 1a 成品原文））：稼動率乘數 {inv('um_A'):.2f}／1.00／{inv('um_C'):.2f}；價格乘數 1.00／1.00／0.85；成本乘數 0.75／1.00／1.40。
- 費用：研發年增 0%／3%／6%（9M'26 實際 +31%，此處假設急速放緩，屬承重 Assumed）；管銷 5% 變動＋固定 3% 年增；稅率 15%（只在稅前為正時）。其餘營收年增 40%／25%／10%（FY2027–30）、毛利率 40%（30–50%）。
- 資本支出：自有車隊比例 50%（30–100%，未揭露）× 新增與汰換車輛 × 線單車成本＋其他 20 百萬美元／年。
- 折現：rf 5.00%（美國 10 年期，與 Tesla 模型同值，逾期）、β {inv('beta'):.3f}（中國自駕同業與全球自駕叫車兩來源 Blume 中位數）、ERP 5%、中國 CRP 0.9%（Assumed）、g＝3%、RONIC＝WACC。
- 評價日股數：9M'26 加權 330.2 百萬 ADS（淨損失 ÷ 每 ADS 損失推得）×（1＋淨稀釋 2%，未取得明細）。流動資金 695.5 百萬美元（Verified）、借款與租賃 114.6；其他長期投資未取得，不計。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 車隊路徑（f27、fg、fglt） | 公司季度車隊揭露（2026-07-31 逾 1,800 → 10-31 逾 2,300 → 年底目標 2,600）、城市數、海外合作台數 |
| 價格乘數（pm） | 中國與海外 Robotaxi 每趟價格與補貼；每車每日營收（Q 頁 Q3'26 推估） |
| 營運成本乘數（om）與 cr_k | 合併毛利率（Q3'26 43.0%，產品組合波動大）、保險成本、遠端監看比 |
| 稼動率乘數（um） | 每車每日付費趟數（廣州 Q3 平均 21 趟）、全無人營運時數 |
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
不一致處：(1) 文遠每英里營收＝線價格 × k_rt（FY2026 校準 {inv('k_rt'):.3f}；Q3'26 每車每日 Robotaxi 營收約 {wb['Q']['F8'].value:.1f} 美元（推估）vs 線基準車資 {wb['Q']['F9'].value:.1f} 美元），差額歸為車輛銷售、海外合作收入與海外價格差；(2) 文遠營運成本水準以 cr_k {inv('cr_k'):.3f} 校準，只沿用線的下降斜率；(3) 線全成本的研發分攤（百度／小馬平均（線 1a 成品原文））不用，改用文遠自己的研發與管銷實績；(4) 線沒有中國滲透率，本模型自建。以上不在本模型靜默修改線值。

## 限制與未找到資料
- 分部毛利與營業利益：未揭露（未找到）；以 FY2026 合併毛利率殘差推 Robotaxi 成本。
- FY2023 全部數字、Q1／Q2'26 單季 Robotaxi 營收：未取得。2025 年底精確車隊：未取得（區間 900–1,125）。資本支出、長期投資明細：未取得。賣方目標價：未查。
- 選擇權與 RSU 數量、履約價：未取得；折舊攤銷：9M'26 為 174.1 人民幣百萬，但模型用資本支出五年平均推算（資本支出未揭露）。
- 公司未給 FY2027 以後車隊或營收目標（未找到）；共識營收只取得 FY2026 與 FY2027 的搜尋摘要值（Yahoo，未回原頁）。
- 同業倍數頁面日期由 2026-10-02 到 2026-11-20 不一，市值依 csv 收盤等比例換算到 2026-11-20（Derived），淨現金與 NTM 營收未更新；地平線 9660.HK 有收盤價但無 Forward 營收，不進中位數；小馬以本專案其個股模型的淨現金口徑納入。
- 中國 CRP 未回原文核對（Assumed）。rf 取 2026-10-23 值（逾期，跨模型一致）。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）；display_check.json 未建。

## 核對（建置時執行）
{chk}

## 檔案
- `build_weride.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`，並跑核對。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("WRD_OVERRIDE", "{}"))
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
    RV = [("fg_B", 0.45, [i / 10 for i in range(-2, 14)], "情境 B 的 FY2028–FY2030 車隊年增（其他不變，整本重算；情境 A／C 不動）"),
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
    bid = f"weride-{VER}-{fingerprint()}"
    tmp = tempfile.mkdtemp(prefix="weridef_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="文遠知行 WeRide", ticker="WRD", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="US$M（人民幣 ÷ 6.7045）", source=Q425 + "；" + Q326)
                                      for n, l, v in (("Robotaxi services", "Robotaxi 服務營收", 22.075), ("Other revenue (residual)", "其餘營收（殘差）", 80.034),
                                                      ("Unused slot", "（無第三分部）營收", 0.0))],
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
        msg = (("通過" if rp.returncode == 0 else "失敗") + "：" + (last[0] if last else "")).replace(ROOT + "/", "")
        results.append((os.path.basename(cmd[0]), msg[:300])); print(cmd[0], rp.returncode, "\n".join(last[:8]))
    HT = X[:-5] + ".html"
    for cmd in (["physical_ai/weride/gen_html_json.py", X], ["tools/company_html/build_company_html.py", HERE],
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
