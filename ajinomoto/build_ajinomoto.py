# 味之素（Ajinomoto，2802.T）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業——ABF 增層膜材料；樣板 ibiden/、dell/）
# 用法：python3 ajinomoto/build_ajinomoto.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月至隔年 3 月：FY(t)＝t/4–(t+1)/3（公司稱「(t+1) 年 3 月期」），對照 CY(t) 的加速器顆數（重疊 9 個月，不做平移；與 Ibiden 模型相同）。
# 金額：十億日圓（¥B）；EPS：日圓（2025-04-01 一拆二之後的股數口徑）。口徑：IFRS「事業利益」（公司主要獲利指標）。
# AI 曝險＝ファンクショナルマテリアルズ（電子材料，以 ABF 為主）；其餘事業（調味料・食品、冷凍食品、生技醫藥與胺基酸等）以成長率與利益率推，不隨 AI 情境。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import valuation_rule as VR
import fcff_dcf as FD   # v0.3（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC）共用模組（docs/plan/個股DCF標準規格.md）
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx"); IB_F = only("ibiden/dist/*.xlsx"); UM_F = only("unimicron/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("um") or str(k).startswith("u_") or str(k).startswith("g29_")):
        TS[k] = (r[4].value, r[0].row)
NV = ("bw", "rb", "ru", "hp"); OT = ("amd", "tpu", "trn", "oth")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
assert all(f"u_{p}_{y}" in TS for p in NV + OT for y in (26, 27, 28)) and all(f"g29_{p}" in TS for p in NV + OT), TS.keys()
def units(plats):
    u = {y: sum(TS[f"u_{p}_{y}"][0] for p in plats) for y in (26, 27, 28)}
    u[29] = sum(TS[f"u_{p}_28"][0] * (1 + TS[f"g29_{p}"][0]) for p in plats)
    return u
UNV = units(NV); UOT = units(OT)
U_SRC = f"{rel(TSMC_F)}｜In!E{TS['u_bw_26'][1]}:E{TS['g29_oth'][1]}"
ibw = openpyxl.load_workbook(IB_F, data_only=True)
IB = {}
for r in ibw["In"].iter_rows(min_row=4):
    if r[0].value in ("g_asp", "pp_27", "pp_28", "pp_29"):
        IB[r[0].value] = (r[3].value, r[4].value, r[5].value, f"{rel(IB_F)}｜In!E{r[0].row}")
assert len(IB) == 4, IB.keys()
IBU = {}
for r in ibw["Units"].iter_rows(min_row=4, max_row=20):
    lab = str(r[0].value or "")
    if lab.startswith("每顆 NVIDIA 載板金額"):
        IBU["v"] = ([r[2 + j].value for j in range(4)], f"{rel(IB_F)}｜Units!C{r[0].row}:F{r[0].row}")
    if lab == "AI 伺服器載板營收":
        IBU["ai"] = ([r[2 + j].value for j in range(4)], f"{rel(IB_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert len(IBU) == 2 and all(v is not None for v in IBU["v"][0] + IBU["ai"][0]), IBU
umw = openpyxl.load_workbook(UM_F, data_only=True)
UM = None
for r in umw["In"].iter_rows(min_row=4):
    if r[0].value == "v_ot":
        UM = (r[3].value, r[4].value, r[5].value, f"{rel(UM_F)}｜In!E{r[0].row}")
assert UM and all(v is not None for v in UM[:3]), UM

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；顆數：百萬顆；比例：%。口徑：IFRS 事業利益。每個輸入盡量附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝味之素決算短信、決算概要與 IR 簡報；"
            "公司預測為 Interested-party；股東行動派（Palliser）、券商、媒體、資料商皆為 Interested-party。加速器顆數讀 tsmc 成品、載板面積與層數年增讀 ibiden 成品、非 NVIDIA 每顆載板面積讀 unimicron 成品（唯讀）")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
TAN25 = "https://www.ajinomoto.com/cms_wp_ajnmt_global/wp-content/uploads/pdf/2026_05_07_03E.pdf"
TANQ1 = "https://news.ajinomoto.co.jp/2026/08/2026_08_06_03.pdf"
SUMQ1 = "https://www.ajinomoto.co.jp/company/jp/ir/event/presentation/main/011111114/teaserItems1/00/linkList/01/link/FY25Q4_Results_J.pdf"
PRQ1 = "https://www.ajinomoto.co.jp/company/jp/ir/event/presentation/main/011111114/teaserItems1/00/linkList/03/link/FY25Q3_Presentation_J.pdf"
ASV = "https://www.ajinomoto.com/sustainability/pdf/2026/ASV_Report_2026_A4_en_014-015.pdf"
TF26 = "https://www.trendforce.com/news/2026/05/08/news-ajinomoto-ramps-chip-packaging-push-with-%C2%A51-2b-land-buy-for-new-plant-in-2032-abf-margins-top-50-on-ai-boom/"
TF25 = "https://www.trendforce.com/news/2025/04/01/news-japanese-firm-ajinomoto-to-invest-jpy-25-billion-by-2030-to-expand-abf-production/"
EDGX = "https://pro.edgex.exchange/en-US/news/article/ajinomoto-ai-chip-film-price-hike"
PAL = "https://mms.businesswire.com/media/20260331071448/ja/2761329/1/JPN_Palliser_-_Ajinomoto_Value_Enhancement_Plan_March_2026_-_vSent.pdf"
DBS = "https://www.dbs.com.sg/treasures/aics/templatedata/article/generic/data/en/GR/Japan/2802_JP.xml"
TN = "https://cdnfinance.technews.tw/2026/06/23/abf-drives-strong-performance-brokers-raise-target-prices-ajinomoto-consecutive-record-highs/"
BGF = "https://finance.biggo.com/news/rYUmMZwBy4YEFZDUYyRK"
SA = "https://stockanalysis.com/quote/tyo/2802/statistics/"
SAF = "https://stockanalysis.com/quote/tyo/2802/forecast/"
SAR = "https://stockanalysis.com/quote/tyo/2802/financials/ratios/"
SAI = "https://stockanalysis.com/quote/tyo/2802/financials/"
SAH = "https://stockanalysis.com/quote/tyo/2802/history/"
CMC = "https://companiesmarketcap.com/ajinomoto/pe-ratio/"
MAT = "https://finance.matsui.co.jp/stock/2802/index"
YJ = "https://finance.yahoo.co.jp/quote/2802.T"
SWS = "https://www.simplywall.st/stocks/jp/food-beverage-tobacco/tse-2802/ajinomoto-shares/valuation"
STO = "https://www.stockopedia.com/share-prices/ajinomoto-co-TYO:2802/"
KY = "https://kabuyoho.jp/report?bcode=2802"
SAP = "https://stockanalysis.com/quote/{}/statistics/"
URLNAME = {TAN25: "公司：2026 年 3 月期決算短信（2026-05-07）", TANQ1: "公司：2027 年 3 月期 Q1 決算短信（2026-08-06）", SUMQ1: "公司：Q1 決算概要（分部門、子部門；2026-08-06）",
           PRQ1: "公司：2026-08-06 IR 簡報（業績預想、AFT 實績、中東影響）", ASV: "公司：ASV Report 2026 分部門回顧", TF26: "TrendForce 2026-05-08（引日經、朝日電視）",
           TF25: "TrendForce 2025-04-01（引日經）", EDGX: "EdgeX 2026-08-08", PAL: "Palliser Capital 價值提升計畫 2026-03（股東行動派）", DBS: "DBS 2026-08-12",
           TN: "TechNews 2026-06-23（引 MoneyDJ／日經）", BGF: "BigGo Finance 2026-02-06", SA: "stockanalysis.com（統計）", SAF: "stockanalysis.com（預估）", SAR: "stockanalysis.com（比率）",
           SAI: "stockanalysis.com（損益）", SAH: "stockanalysis.com（歷史股價）", CMC: "companiesmarketcap", MAT: "松井證券", YJ: "Yahoo!ファイナンス", SWS: "Simply Wall St",
           STO: "Stockopedia", KY: "株予報"}

sec("A. AI 加速器顆數（讀 tsmc 成品；CY2026–CY2028 為輸入、CY2029＝CY2028 ×（1＋各類成長）；情境 A（原始顆數），B／C 依同字母顆數係數縮放）")
for y in (26, 27, 28, 29):
    inp(f"unv_{y}", f"NVIDIA 加速器顆數 CY20{y}（Blackwell＋Rubin＋Rubin Ultra＋Hopper）", "百萬顆", None, round(UNV[y], 6), None, "Derived", U_SRC if y == 26 else "同上",
        "唯讀上游；與台積電、Ibiden、欣興模型同一份顆數" if y == 26 else "同上")
for y in (26, 27, 28, 29):
    inp(f"uot_{y}", f"非 NVIDIA 加速器顆數 CY20{y}（AMD、TPU、Trainium、其他 ASIC）", "百萬顆", None, round(UOT[y], 6), None, "Derived", U_SRC if y == 26 else "同上",
        "ABF 由全部載板廠採用（市占逾九成），所以非 NVIDIA 加速器全數計入，不乘 Ibiden 份額權重" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：CY2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um28_{s}'][1]}", SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：CY2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um29_{s}'][1]}", SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
inp("v_ot", "非 NVIDIA 加速器每顆載板面積（金額）÷ NVIDIA（欣興模型）", "x", UM[0], UM[1], UM[2], "Assumed", UM[3], "TPU、Trainium、MI 系列封裝面積接近 GPU、層數不同；量化：未找到（欣興模型同值）")
inp("g_asp", "每顆載板面積 × 層數年增（Ibiden 模型的每顆載板金額年增）", "%", IB["g_asp"][0], IB["g_asp"][1], IB["g_asp"][2], "Analogy", IB["g_asp"][3],
    "ABF 用量與載板面積 × 增層數成正比：Palliser（Interested-party）Rubin Ultra 155×75 mm、18 層，為 PC CPU 的 13.3 倍面積；EdgeX：每顆 AI 加速器 8–16 層")

sec("B. ABF（ファンクショナルマテリアルズ，FM）需求、產能與價格")
inp("ai_sh", "FY2026 AI 用 ÷ FM 營收", "%", 0.30, 0.40, 0.50, "Assumed", PAL + "：ABF 營收中 AI 占比 FY2024 20%、FY2025E 29%、FY2026E 45%、FY2027E 55%（Interested-party）",
    "公司未揭露 AI 占比：未找到；" + DBS + "：FM 事業利益 1Q FY26 年增 77%（AI 伺服器帶動）")
inp("g_nai", "非 AI ABF 用量年增（PC、一般伺服器、網通）", "%", -0.02, 0.03, 0.06, "Assumed", "PC 與一般伺服器緩步成長（與欣興模型 g_nai 同值）", "量化：未找到")
for y, (lo, b, hi) in ((27, (0.95, 1.00, 1.00)), (28, (0.88, 0.97, 1.00)), (29, (0.80, 0.95, 1.00))):
    inp(f"rel_{y}", f"味之素 ABF 份額 20{y} ÷ 2026（相對值）", "x", lo, b, hi, "Assumed",
        EDGX + "：味之素逾 95%、積水化學約 5%（媒體，Interested-party）" if y == 27 else "同上；" + TF26 + "：市占逾 95%",
        PAL + "：高階 AI 用約 100%（Interested-party）" if y == 27 else "中國廠與積水化學認證進度：未找到")
for y in (27, 28, 29):
    inp(f"cap_{y}", *SIL.row9(f"abffilm_{y}")[1:])
for y, (lo, b, hi) in ((27, (1.00, 1.12, 1.20)), (28, (0.95, 1.15, 1.30)), (29, (0.90, 1.12, 1.30))):
    inp(f"pp_{y}", f"ABF 價格指數 20{y}（FY2026 平均＝1）", "x", lo, b, hi, "Assumed",
        EDGX + "：2026-05 宣布約 30% 漲價，逐客戶適用（FY2026 只反映部分）" if y == 27 else ("同上；供需缺口 2027 約 21%、2028 約 42%（" + EDGX + "引用的供應鏈分析）" if y == 28 else "新產能到位、競爭者認證後溢價回落"),
        PAL + "：要求漲價 30% 以上（Interested-party）" if y == 27 else f"Ibiden 模型載板溢價指數 {IB[f'pp_{y}'][1]}（{IB[f'pp_{y}'][3]}）")
for s, v in (("A", 1.00), ("B", 0.95), ("C", 0.88)):
    inp(f"sp_{s}", f"情境 {s}：FY2028 起 ABF 價格乘數", "x", None, v, None, "Assumed", "公司特有變數依情境調整，方向 A 最有利：需求放緩時漲價難以延續", "—")
inp("c_g", "ABF 每單位成本年增（新產能折舊、原料）", "%", 0.0, 0.03, 0.06, "Assumed", "FM 事業利益率 FY2024 約 52%（40 ÷ 77，Palliser）→ FY2025 53%、1Q FY26 58%", "價格指數上升的部分全部落在利潤（同 Ibiden 模型寫法）")

sec("C. 其他事業（不隨 AI 情境）")
inp("g_sf", "調味料・食品營收年增", "%", 0.02, 0.04, 0.06, "Assumed", TAN25 + "：FY2025 +4.6%；FY2026 預測 +6.6%（含匯率）", "—")
inp("m_sf", "調味料・食品事業利益率（FY2027 起）", "%", 0.135, "=f_sfbp/f_sf", 0.155, "Derived", "FY2026 預測 145.9 ÷ 998.6（" + SUMQ1 + "）", "FY2025 15.3%（143.0 ÷ 936.9）")
inp("g_ff", "冷凍食品營收年增", "%", 0.0, 0.03, 0.05, "Assumed", "FY2025 +0.3%；FY2026 預測 +7.0%（含匯率）", "—")
inp("m_ff", "冷凍食品事業利益率（FY2027 起）", "%", 0.03, 0.045, 0.06, "Assumed", "FY2024 4.5%、FY2025 2.9%、FY2026 預測 3.9%（公司）", PAL + "：要求冷凍食品重整，ROIC 由約 3% 提高到 8%")
inp("g_ho", "ヘルスケア等（扣除 FM）營收年增：生技醫藥 CDMO、胺基酸、其他", "%", 0.02, 0.05, 0.08, "Assumed", "1Q FY26：生技醫藥＆原料 +11%（胺基酸 +21%、CDMO +2%）、其他 +12%（" + SUMQ1 + "）", "—")
inp("m_ho", "ヘルスケア等（扣除 FM）事業利益率（FY2027 起）", "%", 0.05, "=(f_hobp-f_fmbp)/(f_ho-f_fm)", 0.09, "Derived", "FY2026 預測 (85.0 − 65.5) ÷ (406.8 − 120.6)", "FY2025 (66.2 − 54.6) ÷ (341.5 − 102.5)＝4.9%")
inp("g_cc", "全社共通費年增（負值，費用）", "%", 0.02, 0.03, 0.05, "Assumed", "FY2025 +6.9%（39.9 → 42.6）；FY2026 預測 +8.5%（46.2）", "—")

sec("D. 淨利與股數（FY2027 起）")
inp("ni_bp", "淨利（歸屬母公司）÷ 事業利益", "x", 0.58, 0.62, 0.66, "Derived", "FY2026 預測 123.5 ÷ 202.0＝0.611（" + TANQ1 + "）", "FY2025 0.743（含本社土地出售利益）、FY2024 0.441（減損）；2027-04 吸收合併味之素ファインテクノ可能有稅務正面影響（金額未定，" + PRQ1 + "）")
inp("sh_g", "股數年變動（自社株買回）", "%", -0.03, -0.015, 0.0, "Assumed", "期中平均股數 FY2024 10.07 億 → FY2025 9.73 億 → 1Q FY26 9.56 億（" + TANQ1 + "）", "買回計畫金額：未找到")

sec("E. FY2026 錨點（公司 2026-08-06 上修後預測；Q2 決算後更新）")
for k, lab, v in (("f_rev", "營收", 1732.0), ("f_sf", "調味料・食品營收", 998.6), ("f_ff", "冷凍食品營收", 310.6), ("f_ho", "ヘルスケア等營收", 406.8), ("f_oth", "その他營收", 15.8),
                  ("f_fm", "  其中 ファンクショナルマテリアルズ（電子材料等）營收", 120.6), ("f_bp", "事業利益", 202.0), ("f_sfbp", "調味料・食品事業利益", 145.9),
                  ("f_ffbp", "冷凍食品事業利益", 12.1), ("f_hobp", "ヘルスケア等事業利益", 85.0), ("f_othbp", "その他事業利益", 5.1), ("f_cc", "全社共通費", -46.2),
                  ("f_fmbp", "  其中 ファンクショナルマテリアルズ事業利益", 65.5), ("f_oi", "營業利益", 184.2), ("f_ni", "淨利（歸屬母公司）", 123.5)):
    inp(k, f"FY2026 {lab}預測", "¥B", None, v, None, "Interested-party", SUMQ1,
        "期初 1,723.0（" + TAN25 + "）" if k == "f_rev" else ("期初 197.0；中東情勢（原油 110 美元、158 日圓）可能 −25 十億（" + PRQ1 + "）" if k == "f_bp" else
        ("期初 111.6（＋9.0 全數來自電子材料，" + EDGX + "）" if k == "f_fm" else ("期初 60.5（＋5.0）" if k == "f_fmbp" else ("期初 120.0" if k == "f_ni" else "—")))))
inp("f_eps", "FY2026 EPS 預測（基本）", "¥", None, 129.84, None, "Interested-party", TANQ1, "期初 126.16")
for k, lab, v, f in (("q1_rev", "營收", 412.105, "#,##0.0"), ("q1_bp", "事業利益", 60.116, "#,##0.0"), ("q1_fm", "FM 營收", 33.1, "#,##0.0"), ("q1_fmbp", "FM 事業利益", 19.1, "#,##0.0"),
                     ("q1_ni", "淨利（歸屬母公司）", 36.452, "#,##0.0"), ("q1_eps", "基本 EPS", 38.11, "0.00")):
    inp(k, f"1Q FY26（2026-04～06）{lab}實績", "¥" if k == "q1_eps" else "¥B", None, v, None, "Verified", TANQ1 if k in ("q1_rev", "q1_bp", "q1_ni", "q1_eps") else SUMQ1,
        "年增：營收 +13%、事業利益 +27%、淨利 +13%" if k == "q1_rev" else ("年增：FM 營收 +54%、事業利益 +77%（" + SUMQ1 + "）" if k == "q1_fm" else "—"))

sec("F. FY2025 實績（2026 年 3 月期）")
for k, lab, v in (("a_rev", "營收", 1583.719), ("a_sf", "調味料・食品營收", 936.926), ("a_ff", "冷凍食品營收", 290.308), ("a_ho", "ヘルスケア等營收", 341.504), ("a_oth", "その他營收", 14.979),
                  ("a_bp", "事業利益", 181.163), ("a_sfbp", "調味料・食品事業利益", 143.036), ("a_ffbp", "冷凍食品事業利益", 8.457), ("a_hobp", "ヘルスケア等事業利益", 66.202),
                  ("a_othbp", "その他事業利益", 6.064), ("a_cc", "全社共通費", -42.597), ("a_oi", "營業利益", 199.412), ("a_ni", "淨利（歸屬母公司）", 134.675),
                  ("a_da", "折舊攤銷", 88.914), ("a_cx", "資本支出（有形＋無形）", 104.550)):
    inp(k, f"FY2025 {lab}", "¥B", None, v, None, "Verified", TAN25,
        "淨利含本社大樓土地出售利益（日經 2026-04-28）" if k == "a_ni" else ("有形 96.446＋無形 8.104" if k == "a_cx" else ("ASV Report 2026 同值" if k in ("a_rev", "a_bp") else "—")))
inp("a_fm", "FY2025 ファンクショナルマテリアルズ營收（推估）", "¥B", 100.7, "=0.30*a_ho", 104.0, "Derived", ASV + "：ヘルスケア等營收中 Functional Materials 占 30%（四捨五入）",
    TF26 + "：電子材料營收 100.7（+31%）；1Q FY25 21.5（" + SUMQ1 + "）；子公司味之素ファインテクノ營收 98.383（" + PRQ1 + "）")
inp("a_fmbp", "FY2025 ファンクショナルマテリアルズ事業利益", "¥B", None, 54.6, None, "Interested-party", TF26 + "：546 億（+35%），約占全公司三成",
    DBS + "：FM 占 FY25 事業利益 30%；味之素ファインテクノ營業利益 53.012（" + PRQ1 + "）")
inp("a_eps", "FY2025 基本 EPS（分割後）", "¥", None, 138.36, None, "Verified", TAN25, "2025-04-01 一拆二；期中平均股數 9.734 億")

sec("G. 共識（只作 Recon 對帳與差距拆解，不作校準）")
inp("c_ni26", "共識 FY2026 淨利（2027-03 期）", "¥B", None, 138.34, None, "Interested-party", SAF + "（S&P Global 彙整；14 位）", "株予報：共識經常利益 196.8 十億（2026-08）")
inp("c_eps26", "共識 FY2026 EPS（＝共識淨利 ÷ 錨點股數）", "¥", None, "=c_ni26/(f_ni/f_eps)", None, "Derived", "共識淨利 ÷ 公司預測隱含股數 9.512 億", "共識 EPS 的直接數字：未找到")
inp("c_rev26", "共識 FY2026 營收", "¥B", None, 1730.0, None, "Interested-party", SAF + "（1.73 兆）", "—")
inp("c_g27", "共識 FY2027 EPS 年增（2028-03 期）", "%", None, 0.2363, None, "Interested-party", STO + "（2026-10-26：FY2027E +24.37%、FY2028E +23.63%，以公司年度計）", "單一來源；基期口徑未明示")
inp("c_eps27", "共識 FY2027 EPS（＝共識 FY2026 EPS ×（1＋年增））", "¥", None, "=c_eps26*(1+c_g27)", None, "Derived", "上兩列推導", "直接數字：未找到")
inp("c_rev27", "共識 FY2027 營收", "¥B", None, 1840.0, None, "Interested-party", SAF + "：Revenue Next Year 1.84 兆（+6.41%）", "—")

sec("H. 價格與市場倍數")
inp("px", "股價（2026-10-23 收盤）", "¥", None, 5064, None, "Interested-party", MAT + "：2026-10-23 15:30 現值 5,064、前日 5,046", YJ + "：10-22 收盤 5,046（與松井前日相同）；" + STO + "：前收 5,064")
inp("pe_cf", "P/E（公司預測，Yahoo!ファイナンス 10-22）", "x", None, 39.08, None, "Interested-party", YJ + "：PER（會社予想）39.08 倍、EPS 129.12", "—")
inp("fpe", "前瞻 P/E（stockanalysis，2026-09-25）", "x", None, 32.76, None, "Interested-party", SA + "：價格 5,051、Forward PE 32.76", STO + "：12 月滾動預估 P/E 30.66（10-26）")
inp("cons_feps", "共識前瞻 12 月 EPS（＝取得時價格 ÷ 前瞻 P/E）", "¥", None, "=5051/fpe", None, "Derived", "stockanalysis 2026-09-25 價格 5,051 ÷ 前瞻 P/E 32.76", "DBS：FY26F 正常化 EPS 155")
inp("cons_tp", "市場共識目標價（賣方平均）", "¥", 5800, 6321.43, 7300, "Interested-party", SWS + "：14 位平均 6,321（最高 7,300、最低 5,800；2026-10-26）",
    SAF + "：14 位平均 6,104、中位數 6,200（2026-08-07）；" + KY + "：5,906（2026-08）；高盛 6,500（" + TN + "）、DBS 6,200")
inp("so", "期末發行股數（含庫藏）", "十億股", None, 0.977736, None, "Verified", TANQ1 + "：977,735,616 股；庫藏 23,173,875 股", "—")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_g", "收斂法 B 的 FY2032 營收年增（v0.3 起不再用於 DCF；DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.3 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑 fade_gA／B／C")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定各半）", "—")
inp("nm10_k", "十年平均淨利率替代：最後一年（FY2029）淨利率 × 係數", "x", None, 0.8, None, "Assumed", "十年（2016–2025）淨利率：未找到（免費資料商只到近五年：5.6–8.5%，平均 6.5%；" + SAI + "）",
    "工作單：找不到十年平均用最後一年 × 0.8")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與替代十年平均的中點）", "%", None, 0, None, "Derived", "PL 頁情境 B FY2029 × (1＋nm10_k) ÷ 2", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（替代十年平均＝FY2029 × nm10_k）", "%", None, 0, None, "Assumed", "PL 頁情境 C FY2029 × nm10_k", "—")
TP_KEYS += ["tp_g", "tp_w_dcf", "nm10_k", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

sec("J. 自由現金流列用（算力收支地圖讀取；不進 EPS 與目標價）")
inp("da_rev", "折舊 ÷ 營收（自由現金流列用）", "%", None, "=a_da/a_rev", None, "Derived", "FY2025 實績：折舊攤銷 88.914 ÷ 營收 1,583.719＝5.6%（" + TAN25 + "）", "FY2026 起各年比率不變（Assumed）；1Q FY26 23.4 ÷ 412.1＝5.7%")
inp("cx_rev", "資本支出 ÷ 營收（自由現金流列用）", "%", None, "=a_cx/a_rev", None, "Derived", "FY2025 實績：資本支出 104.550 ÷ 營收＝6.6%（" + TAN25 + "）", "FY2026 起各年比率不變（Assumed）；1Q FY26 33.4 ÷ 412.1＝8.1%")

# ---------------- v0.3（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0） ----------------
# 評價日＝2026-12-31（v0.3.1 起統一評價日；原 2026-06-30）；FY2026 只計評價日後一季（frac0＝0.25）；淨現金與投資沿用最近資產負債表 2026-06-30（1Q FY26）；目標時點＝評價日後 12 個月。
SA_BS = "stockanalysis 資產負債表（百萬日圓；https://stockanalysis.com/quote/tyo/2802/financials/balance-sheet/；2026-11-05 讀取；該站年度以結束年標示，其 FY2026＝本模型 FY2025）"
SA_CF = "stockanalysis 現金流量表（百萬日圓；https://stockanalysis.com/quote/tyo/2802/financials/cash-flow-statement/；2026-11-05 讀取）"
SA_IS = "stockanalysis 損益表（百萬日圓；https://stockanalysis.com/quote/tyo/2802/financials/；2026-11-05 讀取）"
BETA_F = "自算產業 β：JP 食品 同業 5 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 日經 225 ^N225 迴歸，Hamada 去槓桿、現金修正）"
BETA_C = "自算產業 β：JP 電子材料／載板 同業 5 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 日經 225 ^N225 迴歸，Hamada 去槓桿、現金修正）"
YF_S = "Yahoo Finance 月報酬 2021-11～2026-10，對 日經 225 ^N225 迴歸，Hamada 去槓桿、現金修正"
_US10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10 同序列；與股價同日（2026-10-23）；只用於 Sens「美元口徑折現」（工作單第二輪修正，Andy 2026-11-05 23:16）")
for _row in FD.in_rows(dict(
        date="2026-12-31（統一評價日；FY2026 只計評價日後一季；v0.3.2 起資產負債表日 2026-06-30 的淨現金與投資以 stub 期滾到評價日）", amt="¥B", shu="十億股", ccy="¥",
        rf_us=(0.05, _US10[0], _US10[1]),
        stub=dict(bs_date="2026-06-30", val_date="2026-12-31", bs_src=SA_BS + "（最新已公布一季：2027 年 3 月期第 1 季；第 2 季決算短信尚未公布）",
                  div=(0, "Assumed", "2026-06-30 → 2026-12-31 無已宣告股利／買回：2026-09 基準日中間股利 ¥25（12 月支付）為公司預想，董事會決議隨第 2 季決算（11 月）",
                       "只扣已宣告者；若計入約 ¥25 × 股數 sh_v＝¥24B（每股 ¥25）")),
        rf=(0.02817, "財務省 國債金利情報（https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv）：2026-11-02 10 年期 2.817%（本月檔最早一日；2026-11-04 為 2.855%）",
            "Trading Economics 2026-11-05 為 2.76%（https://tradingeconomics.com/japan/government-bond-yield）；10 月逐日值本次無法讀取，取與股價日（10-23）最近可得值；同 Kioxia v0.5、Ibiden v0.5", "無風險利率（日本 10 年期公債殖利率，2026-11-02）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)",
                    "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：JP 食品與 JP 電子材料／載板同業中位數依 FM 事業利益占比加權（" + YF_S + "）；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：βL＝βU ×（1＋(1 − tax_w) × D/E）；Andy 2026-11-05 21:21）",
                    "Sens 低／高（Analogy；v0.3.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（tools/fcff_dcf.py v2 write_sens；原「加權中位數 × 0.8 再 Blume／未調整」的高案例低於基準）；個股迴歸 β 未取得"),
        crp=(0.0091, "Verified", "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）：Japan（A1）國家風險溢酬 0.91%",
             "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋CRP，不設股權成本下限"),
        kd=("=rf+0.005", "Assumed", "日本 10 年期公債＋信用利差 0.5pt（A 級附近；借款利率未逐筆查）", "負債占資本約 10%"),
        tax_w=(0.306, "Assumed", "日本法定實效稅率約 30.6%（公司實際有效稅率：資料商未列，未找到）", "NOPAT 各年同用本值；淨利（PL）用 ni_bp 比例，不受本值影響"),
        debt=("=561.919-52.498", "Derived", SA_BS + "：2026-06-30 Total Debt 561,919（含租賃）− 2026-03-31 長期租賃 52,498", "2026-06-30 租賃負債未單獨列示，以 03-31 值拆分（Assumed）"),
        lease=(52.498, "Verified", SA_BS + "：2026-03-31 Long-Term Leases 52,498", "—"),
        cash=(196.979, "Verified", SA_BS + "：2026-06-30 Cash & Short-Term Investments 196,979", "2026-03-31：106,693"),
        inv=(189.162, "Verified", SA_BS + "：2026-06-30 Long-Term Investments 189,162（投資有價證券與權益法投資，政策持股為主）", "以帳面計入；Sens：不計"),
        mi=(73.845, "Verified", SA_BS + "：2026-06-30 Minority Interest 73,845", "—"),
        sh=("=so-0.023173875", "Derived", "期末發行股數 so 0.977736 − 庫藏股 0.023174 十億股（" + TANQ1 + "）", "未來買回由 FCFF 支應，不另推股數"),
        dps=(50, "Derived", "stockanalysis 股利頁（https://stockanalysis.com/quote/tyo/2802/dividend/）：2026-09 基準日中間股利 ¥25（2026-12 支付）；2027-03 期末、2027-09 中間股利假設同 ¥25（2027-06、2027-12 支付）", "評價日 2026-12-31 到目標時點 2027-12-31 之間兩次，共 ¥50（Assumed 股利不變）"),
        t_tgt=(1.0, "目標價為評價日後 12 個月（2027-12-31）"),
        fc_da=(None, "=da_rev", None, "Derived", "＝In 頁 da_rev（FY2025 實績：折舊攤銷 88.914 ÷ 營收 1,583.719＝5.6%）", "FY2027 起線性收斂到長期穩態（FY2032）"),
        fc_da_ss=("=(78.298/1439.231+86.461/1530.556+88.914/1583.719)/3", "Derived",
                  SA_CF + "：近三年（本模型 FY2023–FY2025）折舊攤銷 ÷ 營收平均＝5.4%／5.6%／5.6%（D&A 78,298／86,461／88,914；營收 1,439,231／1,530,556／1,583,719）", "—"),
        fc_cx=(None, "=cx_rev", None, "Derived", "＝In 頁 cx_rev（FY2025 實績：資本支出 104.550〔有形 96.446＋無形 8.104〕÷ 營收＝6.6%）", "Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=MAX((65.786/1439.231+88.104/1530.556+96.446/1583.719)/3+8.104/1583.719,fc_da_ss*(1+g_tv))", "Derived",
                  SA_CF + "：近三年有形資本支出 ÷ 營收平均 5.5%（65,786／88,104／96,446）＋FY2025 無形資本支出比例 0.5%；且不低於折舊 ×（1＋g）（規格一致性）", "—"),
        fc_wc=(0.10, "=((208.069+287.122+42.904-231.979-45.916)/1439.231+(186.669+286.952+62.901-240.614-47.217)/1530.556+(201.877+318.632+76.968-303.96-51.585)/1583.719)/3", 0.25, "Derived",
               SA_BS + "：近三年（2024–2026 年 3 月底）營運資金 ÷ 營收平均＝18.1%／16.2%／15.3%；營運資金＝應收＋存貨＋其他流動資產 − 應付 − 應計費用（不含其他流動負債）",
               "Sens 低／高 10%／25%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_u_food", "產業無槓桿 β 中位數：自算 JP 食品（Blume 調整前；Blume 後 0.437）", "x", None, 0.16, None, "Derived",
    BETA_F + "：2802 2801 2809 2897 2269", "調味料・食品、冷凍食品、生技醫藥等本業")
inp("beta_u_chem", "產業無槓桿 β 中位數：自算 JP 電子材料／載板（Blume 調整前；Blume 後 1.389）", "x", None, 1.58, None, "Derived",
    BETA_C + "：4062 4004 4063 4186 3110", "ファンクショナルマテリアルズ（ABF 增層膜等電子材料）")
inp("beta_w_fm", "β 加權用：FM 事業利益 ÷ 全公司事業利益（FY2026 錨點）", "%", None, "=f_fmbp/f_bp", None, "Derived", "公司 FY2026 預測：FM 65.5 ÷ 事業利益 202.0（Q 頁）", "本業與 ABF 風險不同，以事業利益占比加權產業 β（Assumed 方法）")
inp("beta_u", "產業無槓桿 β（採用值）＝Blume 調整（0.67 ×〔（1 − FM 占比）× 食品 ＋ FM 占比 × 電子材料〕＋ 0.33）", "x", None, "=0.67*((1-beta_w_fm)*beta_u_food+beta_w_fm*beta_u_chem)+0.33", None, "Derived",
    "自算產業 β（Andy 2026-11-05 22:55 對照表：JP 食品 0.437 × 68% ＋ JP 電子材料 1.389 × 32%＝0.742；本模型 FM 占比取 Q 頁 FY2026 錨點 32.4%）；用現金修正值：淨現金與投資另計，β 只對營業資產", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；β 低／高見 Sens")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代 v0.3.1 的加權中位數 × 0.8 再 Blume／中位數未調整")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑")
TP_KEYS += [k for k in FD.JUDG]

r = 4; INROW = {}
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    for u_, nm_ in URLNAME.items():
        s1 = s1.replace(u_, nm_); s2 = s2.replace(u_, nm_)
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — FY2025 實績、1Q FY26 實績與 FY2026 錨點（¥B；EPS 日圓）。FY2026 錨點＝公司 2026-08-06 上修後的全年預測（Q1 公布時）；上半年決算（約 2026-11 上旬）後更新"
Q["A2"] = ("公司不單獨揭露 ABF 營收；ファンクショナルマテリアルズ（FM，電子材料等）為ヘルスケア等的子部門，季度決算概要揭露其營收與事業利益。FY2025 FM 營收以 ASV Report 的 30% 占比推估（Derived），"
           "電子材料營收 100.7（媒體引公司）與子公司味之素ファインテクノ營收 98.4 為對照。FM 內 AI 用比例公司未揭露，以 In 頁 ai_sh（Assumed）切分")
hdr(Q, 3, ["項目", "FY2025A", "1Q FY26（實績）", "FY2026 錨點", "年增（錨點 ÷ FY2025）", "1Q ÷ 錨點（進度）", "標記", "說明"])
QR = {}
QROWS = (("rev", "營收", "=a_rev", "=q1_rev", "=f_rev"), ("sf", "  調味料・食品", "=a_sf", 238.7, "=f_sf"), ("ff", "  冷凍食品", "=a_ff", 72.5, "=f_ff"),
         ("ho", "  ヘルスケア等", "=a_ho", 97.0, "=f_ho"), ("fm", "    其中 ファンクショナルマテリアルズ（電子材料等）", "=a_fm", "=q1_fm", "=f_fm"),
         ("hox", "    其中 FM 以外（生技醫藥＆原料、其他）", "=a_ho-a_fm", 63.9, "=f_ho-f_fm"), ("oth", "  その他", "=a_oth", 3.7, "=f_oth"),
         ("bp", "事業利益", "=a_bp", "=q1_bp", "=f_bp"), ("sfbp", "  調味料・食品", "=a_sfbp", 41.0, "=f_sfbp"), ("ffbp", "  冷凍食品", "=a_ffbp", 2.1, "=f_ffbp"),
         ("hobp", "  ヘルスケア等", "=a_hobp", 25.1, "=f_hobp"), ("fmbp", "    其中 ファンクショナルマテリアルズ", "=a_fmbp", "=q1_fmbp", "=f_fmbp"),
         ("hoxbp", "    其中 FM 以外", "=a_hobp-a_fmbp", 6.0, "=f_hobp-f_fmbp"), ("othbp", "  その他", "=a_othbp", 1.6, "=f_othbp"), ("cc", "  全社共通費", "=a_cc", -9.9, "=f_cc"),
         ("oi", "營業利益", "=a_oi", 55.895, "=f_oi"), ("ni", "淨利（歸屬母公司）", "=a_ni", "=q1_ni", "=f_ni"), ("eps", "EPS（基本，日圓）", "=a_eps", "=q1_eps", "=f_eps"))
for i, (k, lab, a, q1, f) in enumerate(QROWS, 4):
    fmt = "0.00" if k == "eps" else "#,##0.0"
    Q.cell(i, 1, lab); Q[f"B{i}"] = a; Q[f"C{i}"] = q1; Q[f"D{i}"] = f
    if not isinstance(q1, str): Q[f"C{i}"].font = BLUE
    Q[f"E{i}"] = f"=D{i}/B{i}-1"; Q[f"F{i}"] = f"=C{i}/D{i}"
    for c in "BCD": Q[f"{c}{i}"].number_format = fmt
    Q[f"E{i}"].number_format = "0.0%"; Q[f"F{i}"].number_format = "0.0%"
    Q.cell(i, 7, "Derived" if k in ("fm", "hox", "hoxbp") else ("Verified／Interested-party" if k != "fmbp" else "Interested-party／Verified"))
    QR[k] = i
Q.cell(QR["fm"], 8, "FY2025＝0.30 × ヘルスケア等（ASV Report 占比）；1Q 與錨點為公司決算概要揭露值")
Q.cell(QR["ni"], 8, "FY2025 含本社大樓土地出售利益（營業利益 199.4 > 事業利益 181.2），FY2026 預測淨利因此年減")
Q.cell(QR["sf"], 8, "1Q 子部門值取自決算概要（億日圓 ÷ 100）")
Q.cell(QR["cc"], 8, "費用（負值）")
r = QR["eps"] + 2
Q.cell(r, 1, "比率").font = BOLD; r += 1
for k, lab, fb, fc, fd in (("fmm", "FM 事業利益率", f"=B{QR['fmbp']}/B{QR['fm']}", f"=C{QR['fmbp']}/C{QR['fm']}", f"=D{QR['fmbp']}/D{QR['fm']}"),
                           ("fmsh", "FM 營收 ÷ 全公司營收", f"=B{QR['fm']}/B{QR['rev']}", f"=C{QR['fm']}/C{QR['rev']}", f"=D{QR['fm']}/D{QR['rev']}"),
                           ("fmbpsh", "FM 事業利益 ÷ 全公司事業利益", f"=B{QR['fmbp']}/B{QR['bp']}", f"=C{QR['fmbp']}/C{QR['bp']}", f"=D{QR['fmbp']}/D{QR['bp']}"),
                           ("bpm", "事業利益率", f"=B{QR['bp']}/B{QR['rev']}", f"=C{QR['bp']}/C{QR['rev']}", f"=D{QR['bp']}/D{QR['rev']}"),
                           ("nibp", "淨利 ÷ 事業利益", f"=B{QR['ni']}/B{QR['bp']}", f"=C{QR['ni']}/C{QR['bp']}", f"=D{QR['ni']}/D{QR['bp']}"),
                           ("sh", "股數（十億股＝淨利 ÷ EPS）", f"=B{QR['ni']}/B{QR['eps']}", f"=C{QR['ni']}/C{QR['eps']}", f"=D{QR['ni']}/D{QR['eps']}")):
    Q.cell(r, 1, lab); Q[f"B{r}"] = fb; Q[f"C{r}"] = fc; Q[f"D{r}"] = fd
    for c in "BCD": Q[f"{c}{r}"].number_format = "0.0%" if k != "sh" else "0.0000"
    Q.cell(r, 7, "Derived"); QR[k] = r; r += 1
Q.cell(QR["fmbpsh"], 8, "DBS：FY25 30%、FY26 修正預測 32%（Interested-party）")
Q.cell(QR["fmm"], 8, "1Q 達 58%；公司全年預測隱含下半年 FM 事業利益率約 53%（保守）")
Q.cell(r + 1, 1, "讀法：FM 只占營收約 7%，卻占事業利益約三分之一；1Q 的 FM 營收、事業利益分別達全年預測的 27%、29%，公司全年預測隱含下半年放緩，屬保守。").font = BOLD
Q_NOTE = r + 1
for nm, k in (("rev", "rev"), ("sf", "sf"), ("ff", "ff"), ("ho", "ho"), ("fm", "fm"), ("hox", "hox"), ("oth", "oth"), ("bp", "bp"), ("sfbp", "sfbp"), ("ffbp", "ffbp"),
              ("fmbp", "fmbp"), ("hoxbp", "hoxbp"), ("othbp", "othbp"), ("cc", "cc"), ("ni", "ni"), ("eps", "eps"), ("sh", "sh")):
    defname(f"q25_{nm}", f"Q!$B${QR[k]}"); defname(f"q26_{nm}", f"Q!$D${QR[k]}")
Q.column_dimensions["A"].width = 48
for c in "BCDEF": Q.column_dimensions[c].width = 14
Q.column_dimensions["G"].width = 24; Q.column_dimensions["H"].width = 80
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", aish="ai_sh", ga="g_asp", vot="v_ot", rel="rel_{y}", cap="cap_{y}", pp="pp_{y}", cg="c_g", gnai="g_nai",
            gsf="g_sf", msf="m_sf", gff="g_ff", mff="m_ff", gho="g_ho", mho="m_ho", gcc="g_cc", nibp="ni_bp", shg="sh_g", wm=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    p = lambda k, y=None: P[k].format(y=(y % 100 if y else ""))
    um = lambda y: {2026: "", 2027: "", 2028: f"*um28_{s}", 2029: f"*um29_{s}"}[y] + (f"*{P['wm'][y]}" if P["wm"] and y >= 2027 else "")   # v0.13：wm＝Wattonomics 錨定年別乘數（預設 None＝不乘）
    spx = lambda y: "" if y < 2028 else f"*sp_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    sec_("A. ABF 需求（D 欄＝FY2026 校準年；FY(t) 對照 CY(t) 顆數）")
    row("unv", "NVIDIA 加速器顆數 × 情境係數", "百萬顆", {c_(y): f"=unv_{y % 100}{um(y)}" for y in BU}, note="tsmc 成品：Blackwell、Rubin、Rubin Ultra、Hopper")
    row("uot", "非 NVIDIA 加速器顆數 × 情境係數", "百萬顆", {c_(y): f"=uot_{y % 100}{um(y)}" for y in BU}, note="AMD、TPU、Trainium、其他 ASIC")
    row("ueff", "有效顆數 ＝ NVIDIA ＋ 面積係數 × 非 NVIDIA", "百萬顆", {c_(y): f"={c_(y)}{R['unv']}+{P['vot']}*{c_(y)}{R['uot']}" for y in BU})
    row("aidx", "每顆載板面積 × 層數指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['ga']})" for y in F3}}, fmt="0.000")
    row("rel", "味之素 ABF 份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={p('rel', y)}" for y in F3}}, fmt="0.00")
    row("dai", "AI 用 ABF 需求量指數（FY2026＝1）", "x", {c_(y): f"={c_(y)}{R['ueff']}/$D{R['ueff']}*{c_(y)}{R['aidx']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.000",
        note="有效顆數成長 × 面積層數 × 份額")
    row("dnai", "非 AI ABF 需求量指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gnai']})" for y in F3}}, fmt="0.000")
    row("dtot", "ABF 需求量指數（以 FY2026 營收組合加權）", "x", {c_(y): f"={P['aish']}*{c_(y)}{R['dai']}+(1-{P['aish']})*{c_(y)}{R['dnai']}" for y in BU}, fmt="0.000")
    row("cap", "可出貨量上限（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={p('cap', y)}" for y in F3}}, fmt="0.000", note="產能（Assumed）；需求高於上限時出貨＝上限")
    row("ship", "出貨量指數 ＝ MIN（需求, 上限）", "x", {c_(y): f"=MIN({c_(y)}{R['dtot']},{c_(y)}{R['cap']})" for y in BU}, fmt="0.000")
    row("fill", "需求滿足率（出貨 ÷ 需求）", "%", {c_(y): f"={c_(y)}{R['ship']}/{c_(y)}{R['dtot']}" for y in BU}, fmt="0.0%", note="小於 100%＝產能受限，缺口由價格與競爭者填補")
    row("pidx", "ABF 價格指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={p('pp', y)}{spx(y)}" for y in F3}}, fmt="0.000", note="FY2028 起乘情境價格乘數 sp_A／B／C")
    row("fmai", "FM 營收：AI 用", "¥B", {"D": f"=q26_fm*{P['aish']}", **{c_(y): f"=q26_fm*{P['aish']}*{c_(y)}{R['dai']}*{c_(y)}{R['fill']}*{c_(y)}{R['pidx']}" for y in F3}}, fmt="#,##0.0",
        note="FY2026 錨點 × AI 占比 × 需求量 × 滿足率 × 價格")
    row("fmn", "FM 營收：非 AI（PC、一般伺服器等）", "¥B", {"D": f"=q26_fm*(1-{P['aish']})", **{c_(y): f"=q26_fm*(1-{P['aish']})*{c_(y)}{R['dnai']}*{c_(y)}{R['fill']}*{c_(y)}{R['pidx']}" for y in F3}}, fmt="#,##0.0")
    row("fm", "ファンクショナルマテリアルズ（FM）營收", "¥B", {"C": "=q25_fm", **{c_(y): f"={c_(y)}{R['fmai']}+{c_(y)}{R['fmn']}" for y in BU}}, fmt="#,##0.0")
    row("fmm", "FM 事業利益率 ＝ 1 −（1 − 校準值）×（1＋c_g）^年數 ÷ 價格指數", "%", {"C": "=q25_fmbp/q25_fm", "D": "=q26_fmbp/q26_fm",
        **{c_(y): f"=1-(1-$D{r})*(1+{P['cg']})^{y - 2026}/{c_(y)}{R['pidx']}" for y in F3}}, fmt="0.0%", note="Derived：每單位成本按 c_g 上升，漲價差額全部落在利潤")
    row("fmbp", "FM 事業利益", "¥B", {"C": "=q25_fmbp", **{c_(y): f"={c_(y)}{R['fm']}*{c_(y)}{R['fmm']}" for y in BU}}, fmt="#,##0.0")
    sec_("B. 其他事業（不隨 AI 情境）")
    row("sf", "調味料・食品營收", "¥B", {"C": "=q25_sf", "D": "=q26_sf", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gsf']})" for y in F3}}, fmt="#,##0.0")
    row("sfbp", "調味料・食品事業利益", "¥B", {"C": "=q25_sfbp", "D": "=q26_sfbp", **{c_(y): f"={c_(y)}{R['sf']}*{P['msf']}" for y in F3}}, fmt="#,##0.0")
    row("ff", "冷凍食品營收", "¥B", {"C": "=q25_ff", "D": "=q26_ff", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gff']})" for y in F3}}, fmt="#,##0.0")
    row("ffbp", "冷凍食品事業利益", "¥B", {"C": "=q25_ffbp", "D": "=q26_ffbp", **{c_(y): f"={c_(y)}{R['ff']}*{P['mff']}" for y in F3}}, fmt="#,##0.0")
    row("hox", "ヘルスケア等（扣除 FM）營收", "¥B", {"C": "=q25_hox", "D": "=q26_hox", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gho']})" for y in F3}}, fmt="#,##0.0")
    row("hoxbp", "ヘルスケア等（扣除 FM）事業利益", "¥B", {"C": "=q25_hoxbp", "D": "=q26_hoxbp", **{c_(y): f"={c_(y)}{R['hox']}*{P['mho']}" for y in F3}}, fmt="#,##0.0")
    row("oth", "その他營收", "¥B", {"C": "=q25_oth", "D": "=q26_oth", **{c_(y): f"={c_(y - 1)}{r}" for y in F3}}, fmt="#,##0.0", note="FY2027 起持平（Assumed）")
    row("othbp", "その他事業利益", "¥B", {"C": "=q25_othbp", "D": "=q26_othbp", **{c_(y): "=q26_othbp" for y in F3}}, fmt="#,##0.0")
    row("cc", "全社共通費", "¥B", {"C": "=q25_cc", "D": "=q26_cc", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gcc']})" for y in F3}}, fmt="#,##0.0")
    row("adjr", "營收調整（公司合計 − 各事業加總；四捨五入）", "¥B", {"C": "=a_rev-(q25_sf+q25_ff+q25_hox+q25_fm+q25_oth)", "D": "=f_rev-(q26_sf+q26_ff+q26_hox+q26_fm+q26_oth)",
        **{c_(y): f"=$D{r}" for y in F3}}, fmt="#,##0.000", note="FY2027 起固定為 FY2026 值")
    row("adjb", "事業利益調整（公司合計 − 各事業加總；四捨五入）", "¥B", {"C": "=a_bp-(q25_sfbp+q25_ffbp+q25_hoxbp+q25_fmbp+q25_othbp+q25_cc)",
        "D": "=f_bp-(q26_sfbp+q26_ffbp+q26_hoxbp+q26_fmbp+q26_othbp+q26_cc)", **{c_(y): f"=$D{r}" for y in F3}}, fmt="#,##0.000")
    sec_("C. 全公司")
    row("rev", "營收", "¥B", {c_(y): f"={c_(y)}{R['sf']}+{c_(y)}{R['ff']}+{c_(y)}{R['hox']}+{c_(y)}{R['fm']}+{c_(y)}{R['oth']}+{c_(y)}{R['adjr']}" for y in FYS}, fmt="#,##0.0")
    row("hc", "ヘルスケア等營收（FM＋其他ヘルスケア等）", "¥B", {c_(y): f"={c_(y)}{R['fm']}+{c_(y)}{R['hox']}" for y in FYS}, fmt="#,##0.0", note="官方報告分部：ヘルスケア等（含 FM／ABF）")
    row("oth2", "その他營收（含公司合計與各事業加總的四捨五入差）", "¥B", {c_(y): f"={c_(y)}{R['oth']}+{c_(y)}{R['adjr']}" for y in FYS}, fmt="#,##0.0", note="FY2026 公司計畫以十億日圓為單位，各事業加總與合計差 0.2")
    row("segchk", "檢查：分部合計 − 合併營收", "¥B", {c_(y): f"={c_(y)}{R['sf']}+{c_(y)}{R['ff']}+{c_(y)}{R['hc']}+{c_(y)}{R['oth2']}-{c_(y)}{R['rev']}" for y in FYS}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("fmsh", "FM ÷ 營收", "%", {c_(y): f"={c_(y)}{R['fm']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("bp", "事業利益", "¥B", {c_(y): f"={c_(y)}{R['sfbp']}+{c_(y)}{R['ffbp']}+{c_(y)}{R['hoxbp']}+{c_(y)}{R['fmbp']}+{c_(y)}{R['othbp']}+{c_(y)}{R['cc']}+{c_(y)}{R['adjb']}" for y in FYS}, fmt="#,##0.0")
    row("bpm", "事業利益率", "%", {c_(y): f"={c_(y)}{R['bp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("fmbpsh", "FM 事業利益 ÷ 事業利益", "%", {c_(y): f"={c_(y)}{R['fmbp']}/{c_(y)}{R['bp']}" for y in FYS}, fmt="0.0%")
    row("nibp", "淨利 ÷ 事業利益", "x", {"C": "=q25_ni/q25_bp", "D": "=q26_ni/q26_bp", **{c_(y): f"={P['nibp']}" for y in F3}}, fmt="0.000")
    row("ni", "淨利（歸屬母公司）", "¥B", {"C": "=q25_ni", "D": "=q26_ni", **{c_(y): f"={c_(y)}{R['bp']}*{c_(y)}{R['nibp']}" for y in F3}}, fmt="#,##0.0")
    row("sh_", "股數（期中平均）", "十億股", {"C": "=q25_sh", "D": "=q26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['shg']})" for y in F3}}, fmt="0.0000")
    row("eps", "EPS（日圓）", "¥", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("nim", "淨利率", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("fmeps", "FM 每股貢獻 ＝ FM 事業利益 × 淨利 ÷ 事業利益 ÷ 股數", "¥", {c_(y): f"={c_(y)}{R['fmbp']}*{c_(y)}{R['nibp']}/{c_(y)}{R['sh_']}" for y in BU}, note="ABF 對集團 EPS 的貢獻（共通費按事業利益比例不分攤）")
    row("coreeps", "FM 以外每股 ＝ EPS − FM 每股貢獻", "¥", {c_(y): f"={c_(y)}{R['eps']}-{c_(y)}{R['fmeps']}" for y in BU})
    row("aiusd", "AI 用 ABF 營收 ÷ 有效顆數（美元／顆；¥150／US$）", "US$", {c_(y): f"={c_(y)}{R['fmai']}/{c_(y)}{R['ueff']}*1000/150" for y in BU}, fmt="#,##0.0",
        note="¥B ÷ 百萬顆＝千日圓／顆；公司 FY2026 匯率前提 ¥150")
    # FY2030–32：收斂法（上游顆數只到 CY2029＝FY2029；非驅動因子推導）
    for y in FADE:
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+(fade_g{s}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+(fade_m{s}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['sh_']}"] = f"={pc}{R['sh_']}*(1+{P['shg']})"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh_']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.0"), ("nim", "0.00%"), ("ni", "#,##0.0"), ("sh_", "0.0000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = ("PL — 味之素損益（¥B；EPS 日圓；IFRS 事業利益口徑）。FY2025A 實績、FY2026 錨點（公司上修後預測；三情境相同），FY2027 起為模型。FY(t)＝t/4–(t+1)/3，對照 CY(t) 顆數。"
            "三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數與情境價格乘數從 CY2028＝FY2028 開始）")
PB = {}; r0 = 70
SCN_T = {"A": "情境 A（順利）：顆數係數 1.0；ABF 價格乘數 1.0", "B": "情境 B（平平）：CY2028 顆數 ×0.8、CY2029 ×0.85；ABF 價格乘數 0.95", "C": "情境 C（不如預期）：CY2028–29 顆數 ×0.6；ABF 價格乘數 0.88"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "¥B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("fm", "  ファンクショナルマテリアルズ（ABF 等）", "¥B", "#,##0.0"),
        ("fmai", "    其中 AI 用", "¥B", "#,##0.0"), ("fmm", "FM 事業利益率", "%", "0.0%"), ("fmbp", "FM 事業利益", "¥B", "#,##0.0"),
        ("bp", "事業利益", "¥B", "#,##0.0"), ("bpm", "  事業利益率", "%", "0.0%"), ("fmbpsh", "  FM 事業利益 ÷ 事業利益", "%", "0.0%"),
        ("ni", "淨利", "¥B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"), ("eps", "EPS", "¥", "#,##0.00"), ("epsg", "  EPS 年增", "%", "0.0%"), ("fmeps", "  FM 每股貢獻", "¥", "#,##0.00"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("fmai", "fmeps") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=PL!G{SUMR[('B', 'nim')]}*(1+nm10_k)/2")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*nm10_k")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點（校準恆等式）").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（¥B）", "+".join(f"ABS(D{PB[s]['rev']}-f_rev)" for s in SC)),
               ("  三情境：事業利益差（¥B）", "+".join(f"ABS(D{PB[s]['bp']}-f_bp)" for s in SC)),
               ("  三情境：EPS 差（日圓）", "+".join(f"ABS(D{PB[s]['eps']}-f_eps)" for s in SC))):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 68, sr
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 CY2029＝FY2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
# 自由現金流（算力收支 2026-10-25 請求；地圖依列標籤讀取）：表頭列＋一列（三情境 × FY2026–FY2032）
_fh = max(PB[s]["end"] for s in SC) + 2
_FC = {}; _ci = 2
for _s in SC:
    for _y in range(2026, 2033):
        _FC[(_s, _y)] = get_column_letter(_ci); _ci += 1
hdr(PL, _fh, ["項目"] + [f"FY{_y} {_s}" + ("（收斂）" if _y >= 2030 else "") for (_s, _y) in _FC])
PL.cell(_fh + 1, 1, "自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）")
for (_s, _y), _c in _FC.items():
    _e = FC[_y]; _R2 = PB[_s]
    PL[f"{_c}{_fh + 1}"] = f"=PL!{_e}{_R2['ni']}+PL!{_e}{_R2['rev']}*(da_rev-cx_rev)"; PL[f"{_c}{_fh + 1}"].number_format = "#,##0.0"
    if _y >= 2030: PL[f"{_c}{_fh + 1}"].fill = FADEFILL
PL.cell(_fh + 2, 1, "註：上一列＝淨利 ＋ 營收 ×（折舊 ÷ 營收 − 資本支出 ÷ 營收），淨利與營收取本頁各情境推導區塊（FY2030 起為收斂法）；比率見 In 頁 da_rev、cx_rev（FY2025 實績）。¥B。不進 EPS 與目標價。")
FCF_ROW = _fh + 1

# ---------------- Units：ABF 需求推導（情境 A） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 加速器顆數 → ABF 需求量 → 出貨（產能上限）→ FM 營收（情境 A）；每顆加速器的 ABF 金額與 Ibiden 載板金額對照。FY(t) 對照 CY(t)"
A_ = PB["A"]
ucol = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "FY2026（CY2026）", "FY2027", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("unv", "NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['unv']}", "#,##0.00", "tsmc 成品"),
        ("uot", "非 NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uot']}", "#,##0.00", "tsmc 成品"),
        ("ueff", "有效顆數（非 NVIDIA × 面積係數 v_ot）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['ueff']}", "#,##0.00", "v_ot 讀欣興模型"),
        ("aidx", "每顆載板面積 × 層數指數", "x", lambda y: f"=PL!{FC[y]}{A_['aidx']}", "0.000", "g_asp 讀 Ibiden 模型"),
        ("dai", "AI 用 ABF 需求量指數", "x", lambda y: f"=PL!{FC[y]}{A_['dai']}", "0.000", "FY2026＝1"),
        ("dtot", "ABF 需求量指數（含非 AI）", "x", lambda y: f"=PL!{FC[y]}{A_['dtot']}", "0.000", ""),
        ("cap", "可出貨量上限", "x", lambda y: f"=PL!{FC[y]}{A_['cap']}", "0.000", "Assumed"),
        ("fill", "需求滿足率", "%", lambda y: f"=PL!{FC[y]}{A_['fill']}", "0.0%", "小於 100%＝產能受限"),
        ("pidx", "ABF 價格指數", "x", lambda y: f"=PL!{FC[y]}{A_['pidx']}", "0.000", "Assumed"),
        ("fmai", "FM 營收：AI 用", "¥B", lambda y: f"=PL!{FC[y]}{A_['fmai']}", "#,##0.0", ""),
        ("fm", "FM 營收", "¥B", lambda y: f"=PL!{FC[y]}{A_['fm']}", "#,##0.0", ""),
        ("aiusd", "每顆有效加速器的味之素 AI 用 ABF 營收", "US$", lambda y: f"=PL!{FC[y]}{A_['aiusd']}", "#,##0.0", "¥150／US$"),
        ("ibv", "對照：每顆 NVIDIA 載板金額（Ibiden 模型）", "US$", None, "#,##0", IBU["v"][1]),
        ("ratio", "ABF 金額 ÷ 載板金額", "%", lambda y: f"={ucol(y)}{UR['aiusd']}/{ucol(y)}{UR['ibv']}", "0.0%", "Palliser：ABF 約占載板物料成本 10%（Interested-party）"),
        ("ibai", "對照：Ibiden AI 伺服器載板營收（Ibiden 模型）", "¥B", None, "#,##0.0", IBU["ai"][1]),
        ("ibg", "味之素 AI 用 ABF 營收年增 − Ibiden AI 載板營收年增", "pt", lambda y: (f"=({ucol(y)}{UR['fmai']}/{ucol(y - 1)}{UR['fmai']}-1)-({ucol(y)}{UR['ibai']}/{ucol(y - 1)}{UR['ibai']}-1)" if y > 2026 else None), "0.0%", "正值＝ABF 成長快於 Ibiden 載板（漲價、ASIC 全數計入）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit); UR[key] = ur
    for j, y in enumerate(BU):
        if f:
            v = f(y)
            if v is None: continue
        else:
            v = (IBU["v"] if key == "ibv" else IBU["ai"])[0][j]
        UN[f"{ucol(y)}{ur}"] = v; UN[f"{ucol(y)}{ur}"].number_format = fmt
        if not f: UN[f"{ucol(y)}{ur}"].font = BLUE
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：每顆 AI 加速器只用到約二、三十美元的 ABF（約為 NVIDIA 載板金額的一成），所以 ABF 不是成本瓶頸、而是產能與認證瓶頸：需求量（顆數 × 面積層數）每年成長三到八成，"
              "味之素的可出貨量（Assumed）跟不上時，出貨受限、成長改由價格貢獻。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 62; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 16
UN.column_dimensions["G"].width = 70

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（日圓）。目標價＝DCF 與 Comparable 加權（第 0 節）；結論另列「現價 ÷ FY2027、FY2028 EPS（模型）」；十年 P/E 規則值（第 6 節，年底落後 P/E）只作參考；第 7 節為分部加總（SOTP）與現價隱含 ABF 價值（參考）"
V["A2"] = "循環位置：味之素不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 Ibiden、Dell）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# v0.3（2026-11-05）：0.1 改為標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；明細在 DCF 頁）。摘要 18 列寫在舊 0.1 節位置（第 5–22 列），0.2 節以後列號不變。
EPSPL = lambda s_, y: f"PL!{FC[y]}{SUMR[(s_, 'eps')]}"   # 情境 EPS 直接讀 PL 摘要（Val 舊 EPS 列刪除）
EPSR = {}
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（v0.3）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2026 只計評價日後一季）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md")
DCFW["A2"] = ("EBIT：FY2026–FY2029 取 PL 推導區塊「事業利益」（IFRS 公司主要獲利指標，近似營業利益、不含一次性損益）；FY2030–FY2032（收斂法只有淨利）由淨利還原：EBIT＝淨利 ÷ FY2029 淨利 ÷ 事業利益（與 PL 同一口徑）。"
              "稅率各年同用 tax_w（日本法定實效稅率）。折舊、資本支出：FY2025 實績占營收比例（da_rev、cx_rev），FY2027 起線性收斂到長期穩態；營運資金：近三年平均占營收 × 營收增量。")
FY_X = list(range(2026, 2033))
_eb = lambda s_, y: f"PL!{FC[y]}{PB[s_]['bp']}" if y <= 2029 else f"=PL!{FC[y]}{PB[s_]['ni']}/PL!G{PB[s_]['nibp']}"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{ {'A': '順利', 'B': '平平', 'C': '不如預期'}[s_] }）", rev={y: f"PL!{FC[y]}{PB[s_]['rev']}" for y in FY_X}, rev0=f"PL!C{PB[s_]['rev']}",
              ebit={y: _eb(s_, y) for y in FY_X}, tax={y: "tax_w" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷ FY2029（淨利 ÷ 事業利益）") for s_ in SC]
RDCF = FD.write_detail(DCFW, 4, DPATHS, FY_X, 0.25, title="三情境 FCFF 明細（¥B；FY2026＝2026-04～2027-03，評價日 2026-12-31 後只計一季）", stub=True)   # v0.3.2：stub 期淨現金滾到評價日（模型無逐年資本支出／折舊列，維持占營收比例）
RDCF["end"] = FD.write_checks(DCFW, RDCF["end"], nc_bs="196.979+189.162-(561.919-52.498)-52.498-73.845",
                              nc_note="2026-06-30 資產負債表（¥B；stockanalysis；最近一季，v0.3.2 起以 stub 期滾到評價日 2026-12-31）：現金 196.979＋長期投資 189.162 − 債務 509.421 − 租賃 52.498 − 少數股權 73.845") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGHI": DCFW.column_dimensions[_c].width = 13
assert FD.summary_rows(stub=True) <= 20   # 摘要寫在第 5–24 列，0.2 節第 26 列不動
SDCF = FD.write_summary(wb, V, 5, RDCF, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0", beta_note="自算產業 β 重新加槓桿",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.3.2；FY2026 評價日後一季至 FY2032 明確期＋永續；明細見 DCF 頁）")
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"], tvsh=SDCF["tvsh"], mult=SDCF["mult"])
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E：食品同業）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("龜甲萬（Kikkoman）", "2801.T", 22.86, "2026-09-25", "tyo/2801"), ("日清食品 HD（Nissin Foods）", "2897.T", 16.97, "2026-08-07", "tyo/2897"),
         ("明治 HD（Meiji）", "2269.T", 17.79, "2026-09-18", "tyo/2269"), ("Kewpie（丘比）", "2809.T", 23.40, "2026-08-14", "tyo/2809"),
         ("McCormick（味好美）", "MKC", 15.56, "2026-10-23", "stocks/mkc"), ("Nestlé（雀巢）", "NESN.SW", 17.74, "2026-10-23", "swx/NESN")]
for i, (nm, tk, pe, dt, _u) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(34, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0",
     "食品與調味料同業：日本四家（東證食料品）＋全球調味料兩家（不含 A 股）。理由：東證與 GICS 皆歸食品，賣方由食品分析師覆蓋，事業利益約三分之二來自食品本業；電子材料同業與 SOTP 見第 7 節（參考，不進目標價）")
tset(35, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s_}*(5/12*{EPSPL(s_, 2027)}+7/12*{EPSPL(s_, 2028)})" for s_ in SC), None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月（2027-11～2028-10）EPS ≈ 5／12 × FY2027＋7／12 × FY2028（會計年度 3 月底結束，依月份調整）")
tset(36, "comptp", "Comparable 目標價", ["=B34*B35", None, None], "#,##0", "同業中位數 × 前瞻 EPS")
V.cell(38, 1, "0.3 結論").font = BOLD
hdr(V, 39, ["項目", "數值", "對照", "對照", "說明"])
tset(40, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B36", None, None], "#,##0", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(41, "comp2", "Comparable 目標價（市場主流同業法）", ["=B36", None, None], "#,##0")
tset(42, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "Simply Wall St：14 位分析師平均 6,321（最高 7,300、最低 5,800），2026-10-26；stockanalysis 14 位 6,104（2026-08-07）")
tset(43, "upw", "上檔空間：加權目標價", ["=B40/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-23 收盤 5,064）")
tset(44, "upc", "上檔空間：Comparable 目標價", ["=B41/px-1", None, None], "0.0%")
tset(45, "geps", "與市場共識目標價的差距：EPS 因素", ["=B35/C45-1", "=cons_feps", None], "0.0%", "B＝本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；C＝共識前瞻 EPS（stockanalysis 價格 5,051 ÷ 前瞻 P/E 32.76）；兩者期間差約一年")
V["C45"].number_format = "#,##0.00"
tset(46, "gmul", "與市場共識目標價的差距：倍數因素", ["=C46/D46-1", "=cons_tp/cons_feps", "=B34"], "0.0%")
V["C46"].number_format = "0.0"; V["D46"].number_format = "0.0"
V["E46"] = "C＝共識目標價 ÷ 共識前瞻 EPS 隱含的倍數；D＝同業中位數；B＝C ÷ D − 1"
tset(47, "band", "價格帶（DCF 不如預期～順利）", [f"=D{T['dcf12']}", f"=B{T['dcf12']}", None], "#,##0", "B＝C 情境 DCF 價值、C＝A 情境 DCF 價值（皆推到 12 個月後）")
V["C47"].number_format = "#,##0"
O = 60
tset(48, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [None, None, None], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節 companiesmarketcap）的分位；低於最小記 0%、高於最大記 100%；只作檢查，不進目標價")
tset(49, "grade", "上檔空間級距", ['=IF(B43>0.5,"大於 50%",IF(B43>0.25,"25%–50%",IF(B43>=0,"0%–25%",IF(B43>=-0.25,"0%～−25%",IF(B43>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
tset(50, "rec", "投資建議", ['=IF(B43>0.5,"強力買進",IF(B43>0.25,"買進",IF(B43>=-0.25,"中立",IF(B43>=-0.5,"放空","強力放空"))))', None, None], "@",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
V.cell(51, 1, "目標價方法")
V["B51"] = ("DCF（v0.3 標準版）：無槓桿自由現金流 FCFF＝NOPAT（事業利益 ×（1 − 30.6%））＋折舊 − 資本支出 − 營運資金增加，以 WACC 折到 2026-12-31（期中慣例；FY2026 只計評價日後一季；淨現金沿用 2026-06-30 資產負債表），"
            "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 日本 10 年期公債殖利率)、RONIC＝WACC；β＝自算 JP 食品與 JP 電子材料產業無槓桿 β 依 FY2026 事業利益占比加權、Blume 調整後重新加槓桿，ERP 5%＋日本 CRP；EV 加淨現金與投資、減少數股權 ÷ 股數，以 ke 推到 12 個月後並扣期間股利（DCF 頁、tools/fcff_dcf.py）。"
            "Comparable：龜甲萬、日清食品、明治、Kewpie、McCormick、Nestlé 前瞻 P/E 中位數（食品同業，不含 A 股）× 2027-11～2028-10 的 EPS（5/12 × FY2027＋7/12 × FY2028）。"
            "偏誤：食品同業倍數未反映 ABF 的獨占與成長（Comparable 偏低）；食品 β 低使 WACC 約 7%、終值占 DCF 企業價值約七成，DCF 對 WACC 與永續成長率敏感（Sens 頁）。"
            "情境對應：A＝順利（原始顆數）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6），ABF 價格乘數 1.0／0.95／0.88")
for r_ in (T["dcftp"], 36, 40):
    V[f"B{r_}"].fill = KEY
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-23 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0"
for k, y, lab in ((3, 2025, "現價 ÷ FY2025 EPS（實績）"), (4, 2026, "現價 ÷ FY2026 EPS（錨點）"), (5, 2027, "現價 ÷ FY2027 EPS（模型）"), (6, 2028, "現價 ÷ FY2028 EPS（模型）"), (7, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 8, 1, "未來 12 月 EPS（模型 B）≈ 5/12 × FY2026＋7/12 × FY2027（2026-11～2027-10）"); V[f"B{O + 8}"] = f"=5/12*PL!D{SUMR[('B', 'eps')]}+7/12*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 8}"].number_format = "#,##0.00"
V.cell(O + 9, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 9}"] = f"=px/B{O + 8}"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "P/E（公司預測，Yahoo!ファイナンス 10-22）"); V[f"B{O + 10}"] = "=pe_cf"; V[f"B{O + 10}"].number_format = "0.0"
V.cell(O + 11, 1, "前瞻 P/E（stockanalysis，09-25）"); V[f"B{O + 11}"] = "=fpe"; V[f"B{O + 11}"].number_format = "0.0"
V[f"E{O + 4}"] = "FY2025 含土地出售利益；FY2026 三情境相同（錨點）"; V[f"E{O + 5}"] = "FY2027 三情境相同（顆數係數從 CY2028＝FY2028 開始）"
CY0 = O + 13
V.cell(CY0, 1, "2. 循環位置（本模型營收最高年）").font = BOLD
hdr(V, CY0 + 1, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
CYR = {}
for i, s in enumerate(SC, CY0 + 2):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s); CYR[s] = i
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
R30 = CY0 + 6
V.cell(R30, 1, "3. 規則值隱含股價（日圓；約 FY2027 年底＝2028-03，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "#,##0.00"
    V.cell(i, 3, f"=B{i}").number_format = "#,##0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底（2028-03）價 ＝ 現價 ×（1＋報酬）所需的倍數（÷ FY2028 EPS；約一年後的價格看下一年 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (0.6, 0.8, 1.0, 1.21, 1.5):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝約 21%。所需倍數高於歷史中位數時，現價依賴 ABF 成長延續到 FY2029 以後。歷史倍數為 IFRS 淨利口徑（含一次性損益）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
PE_CMC = [18.6, 21.1, 16.4, 30.2, 54.9, 20.3, 24.4, 25.5, 66.4, 42.5]
PE_SA = [None, None, None, None, None, 24.62, 25.92, 33.45, 41.40, 31.34]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；分位排名用 companiesmarketcap 日曆年底；stockanalysis 會計年度年底〔次年 3 月底〕只作核對）").font = BOLD
hdr(V, T10 + 1, ["年底", "採用 P/E（companiesmarketcap）", "companiesmarketcap", "stockanalysis（次年 3 月底）"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f"=C{rr}").number_format = "0.00"
    V.cell(rr, 3, f"=B{RB0 + 5 + i}").number_format = "0.00"
    V.cell(rr, 4, f'=IF(C{RB0 + 5 + i}="","",C{RB0 + 5 + i})').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
for k_, c in enumerate("BCD"):
    _e28 = EPSPL("ABC"["BCD".index(c)], 2028)
    V[f"{c}48"] = f"=IF(px/{_e28}<=MIN({PE10}),0,IF(px/{_e28}>=MAX({PE10}),1,PERCENTRANK({PE10},px/{_e28})))"
    V[f"{c}48"].number_format = "0%"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；IFRS 淨利；2016–2025 共 10 年）"},
    {"label": "stockanalysis（會計年度年底 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；次年 3 月底股價 ÷ 當年度 EPS；只有 2021–2025）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用 companiesmarketcap（日曆年底，10 年齊全）；stockanalysis 只有會計年度年底 2021–2025，只列統計作核對，不採用。"
                "2020、2024 年 P/E 54.9、66.4 倍來自減損使淨利偏低（2024 年度冷凍食品等減損），分位統計受影響。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
# 7. 參考：分部加總（SOTP）與現價隱含 ABF 價值（不進目標價）
S0 = block_end + 1
V.cell(S0, 1, "7. 參考：分部加總（SOTP）與現價隱含的 ABF 價值（不進目標價；Comparable 目標價仍為第 0.2 節食品同業法）").font = BOLD
hdr(V, S0 + 1, ["電子材料同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
EPEERS = [("Resonac（力森諾科）", "4004.T", 20.86, "2026-09-25"), ("東京應化（Tokyo Ohka Kogyo）", "4186.T", 20.85, "2026-10-19"),
          ("日東紡（Nitto Boseki）", "3110.T", 24.44, "2026-09-18"), ("信越化學（Shin-Etsu）", "4063.T", 18.85, "2026-10-21"), ("Ibiden", "4062.T", 46.82, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(EPEERS, S0 + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
EP1, EP2 = S0 + 2, S0 + 1 + len(EPEERS)
SR = {}
srow_ = EP2 + 1
def sotp(key, label, f, fmt, note=""):
    global srow_
    V.cell(srow_, 1, label); V[f"B{srow_}"] = f; V[f"B{srow_}"].number_format = fmt
    if note: V[f"E{srow_}"] = note
    SR[key] = srow_; srow_ += 1
sotp("emed", "電子材料同業前瞻 P/E 中位數", f"=MEDIAN(C{EP1}:C{EP2})", "0.0", "半導體材料與 AI 載板：Resonac、東京應化、日東紡（T-glass）、信越化學、Ibiden（Palliser 的比較組另含欣興、南電，屬台股載板）")
FW = lambda key: "+".join(f"tp_w{s}*(5/12*PL!E{PB[s][key]}+7/12*PL!F{PB[s][key]})" for s in SC)
sotp("ffeps", "12 個月後前瞻 EPS：FM（ABF 等）部分（情境加權）", "=" + FW("fmeps"), "#,##0.00", "FM 每股貢獻（PL）；5/12 × FY2027＋7/12 × FY2028")
sotp("fceps", "12 個月後前瞻 EPS：FM 以外（食品等）部分（情境加權）", "=" + FW("coreeps"), "#,##0.00", "EPS − FM 每股貢獻；全社共通費全部歸本業")
sotp("sotp", "SOTP 價值 ＝ 本業 × 食品同業中位數 ＋ FM × 電子材料同業中位數", f"=B{SR['fceps']}*B{T['pmed']}+B{SR['ffeps']}*B{SR['emed']}", "#,##0", "參考，不進目標價")
sotp("sotpup", "  SOTP 相對現價", f"=B{SR['sotp']}/px-1", "0.0%")
sotp("ceps", "現在的未來 12 月 EPS：FM 以外（模型 B；5/12 × FY2026＋7/12 × FY2027）", f"=5/12*PL!D{PB['B']['coreeps']}+7/12*PL!E{PB['B']['coreeps']}", "#,##0.00")
sotp("feps0", "現在的未來 12 月 EPS：FM（模型 B）", f"=5/12*PL!D{PB['B']['fmeps']}+7/12*PL!E{PB['B']['fmeps']}", "#,##0.00")
sotp("cval", "本業價值（× 食品同業中位數）", f"=B{SR['ceps']}*B{T['pmed']}", "#,##0")
sotp("ival", "現價隱含的 ABF（FM）每股價值 ＝ 現價 − 本業價值", f"=px-B{SR['cval']}", "#,##0")
sotp("ish", "  ABF 占現價比例", f"=B{SR['ival']}/px", "0.0%")
sotp("ipe", "  現價隱含的 ABF 前瞻 P/E（÷ FM 未來 12 月 EPS）", f"=B{SR['ival']}/B{SR['feps0']}", "0.0", "對照：電子材料同業中位數（上方）、Ibiden 前瞻 P/E")
V.column_dimensions["A"].width = 62
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點）：對 FY2027、FY2028 營收、事業利益率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 事業利益率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 事業利益率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 占比 30%", "FY2026 FM 營收中 AI 用 30%", dict(aish="ai_sh_lo")), ("AI 占比 50%", "50%", dict(aish="ai_sh_hi")),
         ("面積層數 +10%／年", "Ibiden 模型低值", dict(ga="g_asp_lo")), ("面積層數 +22%／年", "Ibiden 模型高值", dict(ga="g_asp_hi")),
         ("非 NVIDIA 面積係數 0.6", "欣興模型低值", dict(vot="v_ot_lo")), ("非 NVIDIA 面積係數 1.0", "欣興模型高值", dict(vot="v_ot_hi")),
         ("產能上限 低", "1.15／1.30／1.45（公司 +50%／五年的速度）", dict(cap="cap_{y}_lo")), ("產能上限 高", "1.45／1.85／2.25（不設限）", dict(cap="cap_{y}_hi")),
         ("ABF 價格 低", "1.00／0.95／0.90", dict(pp="pp_{y}_lo")), ("ABF 價格 高", "1.20／1.30／1.30", dict(pp="pp_{y}_hi")),
         ("份額流失", "份額相對值 0.95／0.88／0.80", dict(rel="rel_{y}_lo")), ("單位成本 +6%／年", "c_g 高值", dict(cg="c_g_hi")),
         ("非 AI ABF −2%／年", "g_nai 低值", dict(gnai="g_nai_lo")), ("調味料・食品利益率 13.5%", "m_sf 低值", dict(msf="m_sf_lo")),
         ("調味料・食品利益率 15.5%", "m_sf 高值", dict(msf="m_sf_hi")), ("調味料・食品成長 2%", "g_sf 低值", dict(gsf="g_sf_lo")),
         ("冷凍食品利益率 6%", "m_ff 高值（重整成功）", dict(mff="m_ff_hi")), ("淨利 ÷ 事業利益 0.58", "ni_bp 低值", dict(nibp="ni_bp_lo")),
         ("淨利 ÷ 事業利益 0.66", "ni_bp 高值（合併稅務利益）", dict(nibp="ni_bp_hi")), ("買回 多（−3%）", "股數年變動 −3%", dict(shg="sh_g_lo"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "bpm", 2027, "0.0%"), ("E", "eps", 2027, "#,##0.00"),
                           ("G", "rev", 2028, "#,##0.0"), ("H", "bpm", 2028, "0.0%"), ("I", "eps", 2028, "#,##0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"F{i}"] = f"=E{i}/E$4-1"; SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c in "FJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 50; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 B", "CY2028 顆數 ×0.8、CY2029 ×0.85；ABF 價格 ×0.95", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "CY2028–29 顆數 ×0.6；ABF 價格 ×0.88", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 起需求量高於可出貨量上限，AI 占比、面積層數與顆數假設的變動大多被產能上限吸收；EPS 對 ABF 價格與產能上限最敏感，其次是調味料・食品利益率（本業事業利益約三分之二）。情境 B、C 的差異從 FY2028 起。")
SENS_NOTE = nn + 3
TPS0 = nn + 5
# v0.3：目標價參數敏感度改為標準版 DCF 的標準案例（tools/fcff_dcf.py；每格為閉式公式，與 DCF 頁明細同式）；舊 tp_r／tp_g 案例刪除
_TS = FD.write_sens(SE, TPS0, RDCF, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref="Val!$B$36", base_ref="Val!$B$40", fmt="#,##0",
                    title="目標價參數敏感度（v0.3 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）；基準＝三情境加權（Val 頁第 0.3 節）",
                    beta_du="beta_du", usd=True)   # v0.3.2：β＝β_u ∓ beta_du 重新加槓桿；Sens「美元口徑折現」（非美國公司）
TPS_LAST = _TS[1]
assert TPS_LAST < 50, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 48
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司預測、共識與上游載板模型。只作事後對照，不為對上數字而調參數"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
RC.cell(rr_, 1, "A. 共識（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=c_rev26", pl("rev", 2026), "stockanalysis（14 位）", "錨點＝公司預測 1,732")
rec("FY2026 EPS vs 共識", "=c_eps26", pl("eps", 2026), "stockanalysis 共識淨利 ÷ 股數", "錨點＝公司預測 129.84；共識高於公司預測（公司 FM 預測保守）", fmt="#,##0.00")
rec("FY2027 營收 vs 共識", "=c_rev27", pl("rev", 2027), "stockanalysis（Revenue Next Year）", "")
rec("FY2027 EPS vs 共識", "=c_eps27", pl("eps", 2027), "Stockopedia 年增率推導", "單一來源；基期口徑未明示", fmt="#,##0.00")
rec("FY2028 EPS vs 共識", None, pl("eps", 2028), "未找到", "共識 FY2028 EPS：未找到", fmt="#,##0.00")
RC.cell(rr_, 1, "B. 公司預測與 ABF 說法").font = BOLD; rr_ += 1
rec("FY2026 FM 營收 vs 1Q 年化（× 4）", "=q1_fm*4", pl("fm", 2026), "公司 1Q 決算概要", "公司全年預測低於 1Q 年化：下半年放緩或保守")
rec("FY2026 FM 事業利益 vs 1Q 年化", "=q1_fmbp*4", pl("fmbp", 2026), "同上", "")
rec("FY2027 FM 事業利益率 vs Palliser 預估 66%", 0.66, pl("fmm", 2027), "Palliser 2026-03（Interested-party）", "Palliser：FY2025E 55%、FY2026E 63%、FY2027E 66%；共識（Palliser 引用）62%", fmt="0.0%")
rec("FY2027 FM 營收 ÷ FY2025（兩年成長）", None, lambda s: f"=PL!E{PB[s]['fm']}/PL!C{PB[s]['fm']}-1", "Palliser：未來兩年 ABF 營收年複合 30% 以上（約 +69%）", "", fmt="0.0%", gap=False)
rec("FY2029 FM 營收 vs FY2024 實績 77（Palliser）", 77, pl("fm", 2029), "Palliser FY2024 77；TrendForce 2025-04-01（引日經：2030 年前年增 10% 以上）", "公司說法：年增 10% 以上 × 5 年 ≈ 124 以上")
RC.cell(rr_, 1, "C. 上游載板模型（Ibiden；唯讀）").font = BOLD; rr_ += 1
rec("FY2027 每顆 AI 加速器 ABF 金額 ÷ NVIDIA 載板金額", 0.10, lambda s: f"=Units!D{UR['ratio']}", "Palliser：ABF 約占載板物料成本 10%", "Units 頁（情境 A）", fmt="0.0%")
rec("FY2027 AI 用 ABF 營收年增 − Ibiden AI 載板營收年增", None, lambda s: f"=Units!D{UR['ibg']}", "ibiden/dist Units", "負值＝味之素出貨受產能上限", fmt="0.0%")
RC.cell(rr_, 1, "D. 校準").font = BOLD; rr_ += 1
rec("FY2026 FM 事業利益率（錨點）", None, pl("fmm", 2026), "公司預測 65.5 ÷ 120.6", "", fmt="0.0%")
rec("FY2026 AI 用 ABF 營收（Assumed 切分）", None, pl("fmai", 2026), "ai_sh 40%", "公司未揭露")
rec("FY2026 每顆有效加速器 AI 用 ABF 營收（US$）", None, pl("aiusd", 2026), "—", "")
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 60
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文為決算短信、決算概要、IR 簡報；公司預測為 Interested-party；股東行動派、券商、媒體與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：2026 年 3 月期決算短信（FY2025 實績、分部門、FY2026 期初預測、折舊、資本支出、股數；2026-05-07）", TAN25),
        ("公司：2027 年 3 月期 Q1 決算短信（1Q 實績、上修後全年預測、EPS 129.84、股數；2026-08-06）", TANQ1),
        ("公司：2027 年 3 月期 Q1 決算概要（分部門與子部門：ファンクショナルマテリアルズ 1Q 33.1／19.1、全年 120.6／65.5；匯率前提 ¥150）", SUMQ1),
        ("公司：2026-08-06 IR 簡報（味之素ファインテクノ FY2025 營收 98.383、營業利益 53.012；2027-04 吸收合併；中東影響 −25 十億）", PRQ1),
        ("公司：ASV Report 2026 分部門回顧（Functional Materials 占ヘルスケア等營收 30%）", ASV),
        ("媒體：TrendForce 2026-05-08（FM 事業利益 546 億、電子材料營收 1,007 億、市占逾 95%、岐阜新廠 2032）", TF26),
        ("媒體：TrendForce 2025-04-01（2030 年前投資 250 億、ABF 產能 +50%、FM 年增 10% 以上）", TF25),
        ("媒體：EdgeX 2026-08-08（ABF 漲價約 30%、積水化學約 5%、供需缺口、股價 5,196）", EDGX),
        ("媒體：BigGo Finance 2026-02-06（FY2025 預測上修，電子材料帶動）", BGF),
        ("股東行動派：Palliser Capital 價值提升計畫 2026-03（ABF 歷年營收與利益、AI 占比、漲價建議、同業 P/E；Interested-party）", PAL),
        ("券商：DBS 2026-08-12（BUY、目標價 6,200、40 倍 × FY26F 正常化 EPS 155）", DBS),
        ("媒體：TechNews 2026-06-23（高盛目標價 6,500）", TN),
        ("股價：松井證券（2026-10-23 收盤 5,064）", MAT), ("股價：Yahoo!ファイナンス（10-22 收盤 5,046、PER 會社予想 39.08）", YJ),
        ("股價：stockanalysis 歷史股價（10 月）", SAH), ("資料商：stockanalysis 統計（前瞻 P/E 32.76，09-25）", SA),
        ("資料商：stockanalysis 預估（目標價 14 位 6,104、FY2026 共識營收、淨利）", SAF), ("資料商：stockanalysis 比率（ROE、會計年度年底 P/E）", SAR),
        ("資料商：stockanalysis 損益（五年淨利率）", SAI), ("資料商：Simply Wall St（目標價 14 位 6,321，2026-10-26）", SWS),
        ("資料商：Stockopedia（共識 EPS 年增率、滾動預估 P/E）", STO), ("資料商：株予報（目標價 5,906、共識經常利益）", KY),
        ("資料商：companiesmarketcap 年底 P/E", CMC),
        ("資料商：stockanalysis 同業前瞻 P/E（食品：2801、2897、2269、2809、MKC、NESN；電子材料：4004、4186、3110、4063、4062）",
         "；".join(SAP.format(t) for t in ("tyo/2801", "tyo/2897", "tyo/2269", "tyo/2809", "stocks/mkc", "swx/NESN", "tyo/4004", "tyo/4186", "tyo/3110", "tyo/4063", "tyo/4062"))),
        ("市場層與公司模型：tsmc、ibiden、unimicron 成品（本 repo，唯讀）", f"{rel(TSMC_F)}；{rel(IB_F)}；{rel(UM_F)}"),
        ("[Verified] 自由現金流列：FY2025 折舊攤銷 88.914、資本支出 104.550（有形 96.446＋無形 8.104）、營收 1,583.719（¥B；決算短信）", TAN25),
        ("[Assumed] 折舊 ÷ 營收、資本支出 ÷ 營收 FY2026 起各年維持 FY2025 實績比率", "本模型假設"),
        ("待補（人工）：上半年決算（約 2026-11 上旬，日期未找到）；ABF 單獨營收與 AI 占比（公司未揭露）；ABF 產能數字；共識 FY2028 EPS；十年淨利率；積水化學與中國廠認證進度", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
import json as _json
OBS = ["目標價 ¥2,537（DCF 2,110、Comparable 2,965 各半），比現價 5,064（2026-10-23）低 49.9%，級距 −25%～−50%；Comparable 單獨為 −41.5%。主因是倍數不是 EPS：本模型 12 個月後前瞻 EPS 166.9 反而比共識前瞻 EPS 154.2 高 8.2%，"
       "但食品同業前瞻 P/E 中位數 17.8 倍、共識目標價 6,321 隱含 41.0 倍（倍數因素 +131%）；（v0.2 舊法 DCF 的倍數說明已不適用：v0.3 起 DCF 改為標準版，見 Val 第 0.1 節與展示層。）",
       "現價隱含的 ABF 價值：FM 以外的未來 12 月 EPS 92.2 × 食品同業 17.8 倍＝1,638，現價其餘 3,426（68%）都是 ABF（FM）的價值，等於 FM 未來 12 月 EPS 56.8 的 60.3 倍，"
       "高於電子材料同業中位數 20.9 倍與 Ibiden 46.8 倍；以電子材料同業倍數計 FM 的 SOTP 為 3,176（−37.3%）。",
       "ABF 成長：FM 營收由 FY2026 公司預測 120.6 十億日圓升到 FY2027 175.6（+46%；AI 用 48.2 → 94.6），FM 事業利益 101.8 占全公司 41%、每股貢獻 67.4（EPS 162.8 的 41%）；"
       "前提是約 30% 漲價年度化（價格指數 1.12）與可出貨量 +30%（需求 1.34 倍、滿足率 97%）。ABF 價格取低值時 FY2028 EPS −13.6%、產能取公司 +50%／五年的速度時 −7.3%；"
       "追蹤上半年決算（約 2026-11 上旬）的 FM 營收與事業利益（1Q 33.1／19.1 已達全年預測的 27%／29%）。"]   # 重點觀察（與 README.md 相同文字）
lines = [("味之素（Ajinomoto，2802.T）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業——ABF 增層膜材料；樣板＝ibiden/、dell/）　2026-10-26", True),
 ("版本：v0.1（初版）。味之素以調味料、冷凍食品、胺基酸與生技醫藥為主；AI 曝險在ファンクショナルマテリアルズ（FM）的 ABF（Ajinomoto Build-up Film）——高階 ABF 載板的增層絕緣材料，市占逾九成（媒體與股東行動派說法，Interested-party）。", False),
 ("模型命題：味之素的 FM 營收＝FY2026 錨點 ×〔AI 用：AI 加速器有效顆數（tsmc 成品，同字母情境係數）× 每顆載板面積層數（Ibiden 模型）× 份額；非 AI：成長率〕× 需求滿足率（可出貨量上限）× ABF 價格指數；FM 事業利益率隨價格與單位成本變動；其他事業以成長率與利益率推，不隨 AI 情境。", False),
 ("核心問題：FM 只占營收約 7%、卻占事業利益約三分之一；AI 載板需求能讓 ABF 成長多少、對集團 EPS 貢獻多少？現價約 31 倍 FY2027 EPS、遠高於食品同業約 18 倍，隱含多少 ABF 價值？", False),
 ("重點觀察 1：" + OBS[0], False), ("重點觀察 2：" + OBS[1], False), ("重點觀察 3：" + OBS[2], False),
 ("會計年度：FY(t)＝t 年 4 月～t＋1 年 3 月（公司稱「t＋1 年 3 月期」），對照 CY(t) 顆數（與 Ibiden 模型相同）。金額：十億日圓。口徑：IFRS 事業利益（公司主要指標）；EPS＝歸屬母公司淨利 ÷ 期中平均股數（2025-04-01 一拆二後）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 用 ABF：tsmc 成品的 NVIDIA 與非 NVIDIA 加速器顆數（非 NVIDIA × 欣興模型的面積係數 0.8）× 同字母情境係數 × Ibiden 模型的每顆載板面積層數年增 18% × 份額相對值 → 需求量；與非 AI 需求（年增 3%）加權後，受可出貨量上限（Assumed）限制。", False),
 ("價格：2026-05 宣布的約 30% 漲價在 FY2026 只反映部分，FY2027 價格指數 1.12、FY2028 1.15、FY2029 1.12（Assumed），FY2028 起再乘情境乘數 1.0／0.95／0.88；FM 事業利益率＝1 −（1 − 校準值）× 單位成本指數 ÷ 價格指數。", False),
 ("校準：FY2026 ＝ 公司 2026-08-06 上修後預測（營收 1,732、事業利益 202.0、淨利 123.5、EPS 129.84；FM 120.6／65.5）。AI 占比 40%（Assumed）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（v0.3 起為標準版：無槓桿自由現金流 FCFF 以 WACC 折現，FY2026 評價日後一季至 FY2032＋永續，明細在 DCF 頁；三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（食品同業：龜甲萬、日清食品、明治、Kewpie、McCormick、Nestlé 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價、差距拆解與上檔空間級距。第 7 節 SOTP（FM 用電子材料同業倍數）與現價隱含 ABF 價值只作參考。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（上游顆數到 CY2029＝FY2029；營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("", False),
 ("已知限制：(1) 公司不單獨揭露 ABF 與 AI 占比，FM 以子部門揭露值、AI 占比 40% 為 Assumed；(2) 可出貨量上限與價格指數是 Assumed，FY2027 起產能受限，EPS 對兩者最敏感；(3) FY2026 錨點用公司預測，1Q FM 已達全年 27–29%，預測偏保守；(4) 本業以成長率推，中東情勢（公司估最多 −25 十億）未納入；(5) 共識 FY2027 EPS 單一來源推導、FY2028 未找到；(6) 十年 P/E 2020、2024 年受減損影響；(7) 上半年決算（約 2026-11 上旬）後更新錨點。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "ajinomoto")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.1", "2026-10-30")  # v0.2.1：In 頁版面（個股模型原則第 8 節；數值不變）
# ---------------- v0.13（2026-11-04）：Sens「Wattonomics 錨定」案例（三層校準原則；決策 S37）。In 頁最後三列、Sens 頁最下方小表與案例區塊；既有格不動 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
WTA_MF = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
def wta_in_rows():
    r = IN.max_row + 2
    IN.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = BOLD; r += 1
    def _inp(key, item, unit, lo, base, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((key, item, unit, lo, base, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    SIL.wta_rows(_inp)
def wta_sens():
    t0 = SE.max_row + 3; b0 = t0 + 11
    RW = engine(SE, b0, "案例區塊：Wattonomics 錨定（情境 B；FY2027 起顆數 × wta_27／wta_28／wta_29）", scn="B", wm=WTA_MF)
    SE.cell(t0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    for c in range(1, 10 + 1): SE.cell(t0 + 1, c, SE.cell(3, c).value).font = BOLD
    put_case(t0 + 2, "基準（情境 A；同第 4 列）", "PL 區塊 A", PB["A"], "PL!")
    put_case(t0 + 3, "情境 B（本案例的底）", "CY2028 顆數 ×0.8、CY2029 ×0.85；ABF 價格 ×0.95", PB["B"], "PL!")
    put_case(t0 + 4, "Wattonomics 錨定", "情境 B；FY2027 起 NVIDIA 與非 NVIDIA 顆數 × wta_27／wta_28／wta_29（FY(t) 對照 CY(t) 顆數）", RW, "")
    SE.cell(t0 + 5, 1, "差異%：Wattonomics 錨定 ÷ 基準 − 1"); SE.cell(t0 + 6, 1, "差異%：Wattonomics 錨定 ÷ 情境 B − 1")
    for c in "CDEGHI":
        for rr, rb in ((t0 + 5, t0 + 2), (t0 + 6, t0 + 3)):
            SE[f"{c}{rr}"] = f"=IFERROR({c}{t0 + 4}/{c}{rb}-1,\"\")"; SE[f"{c}{rr}"].number_format = "0.0%"
    SE.cell(t0 + 7, 1, WTA_NOTE)
    SE.cell(t0 + 8, 1, "FY2030–FY2032 收斂欄沿用本模型延伸表收斂法（由 FY2029 收斂），不另乘係數；需求高於產能上限時，量的下修先被產能吸收。")
    SE.cell(t0 + 9, 1, "目標價：Sens 頁目標價參數表只動 DCF 參數（r、g、權重），不逐案例重算 DCF，本案例不列目標價。")
    return RW
wta_in_rows(); WTA_R = wta_sens()   # v0.13：放在 In 版面套用之後，確保在 In 頁最後
# ---------------- v0.3（2026-11-05）：HTML 新版面 v1.0（個股 HTML 版面規格 v1.0）＋標準版 DCF 的展示層 ----------------
# README 頁最後新增「展示層」「命題回覆」兩段（HTML 讀；數字一律由公式帶出）；原 README 各列保留不動。文字中的 {鍵|格式} 轉成 TEXT(儲存格,"格式")。
import re as _re  # noqa: E402
def _srow(prefix):
    for r_ in range(4, SE.max_row + 1):
        if str(SE.cell(r_, 1).value or "").startswith(prefix):
            return r_
    raise KeyError(prefix)
_s = {k: _srow(p) for k, p in (("plo", "ABF 價格 低"), ("phi", "ABF 價格 高"), ("clo", "產能上限 低"), ("chi", "產能上限 高"), ("msf", "調味料・食品利益率 13.5%"))}
_V = lambda k: f"Val!B{T[k]}"
_P = lambda s_, k, y: f"PL!{FC[y]}{SUMR[(s_, k)]}"
_fe = "(" + "+".join(f"tp_w{s_}*(5/12*{_P(s_, 'eps', 2027)}+7/12*{_P(s_, 'eps', 2028)})" for s_ in SC) + ")"
REF = {"tpw": _V("tpw"), "dcftp": _V("dcftp"), "comp": _V("comptp"), "compadj": "tp_comp_adj", "px": "px", "upw": _V("upw"), "upc": _V("upc"), "rec": _V("rec"),
       "grade": _V("grade"), "wdcf": "tp_w_dcf", "pmed": _V("pmed"), "feps": _V("feps"), "cons": _V("cons"), "geps": _V("geps"), "rmf": "rm_factor",
       "pmadj": f"({_V('pmed')}*rm_factor)", "pxpe": f"(px/{_V('feps')})", "dcfup": f"({_V('dcftp')}/px-1)", "upcadj": "(tp_comp_adj/px-1)",
       "ival": f"Val!B{SR['ival']}", "ish": f"Val!B{SR['ish']}", "ipe": f"Val!B{SR['ipe']}", "emed": f"Val!B{SR['emed']}", "sotp": f"Val!B{SR['sotp']}", "sotpup": f"Val!B{SR['sotpup']}",
       "cval": f"Val!B{SR['cval']}", "fm26": _P("A", "fm", 2026), "fm27": _P("A", "fm", 2027), "fmai26": _P("A", "fmai", 2026), "fmai27": _P("A", "fmai", 2027),
       "fmbp27": _P("A", "fmbp", 2027), "fmbpsh27": _P("A", "fmbpsh", 2027), "fmeps27": _P("A", "fmeps", 2027), "eps27": _P("A", "eps", 2027), "fmbpsh26": _P("A", "fmbpsh", 2026),
       "eps28A": _P("A", "eps", 2028), "eps28B": _P("B", "eps", 2028), "eps28C": _P("C", "eps", 2028),
       "wacc": "wacc", "ke": "ke", "beta": "beta", "betau": "beta_u", "wfm": "beta_w_fm", "rfv": "rf", "erp": "erp", "crp": "crp", "gtv": "g_tv"}
for s_, c in zip(SC, "BCD"):
    REF[f"dcf{s_}"] = f"Val!{c}{T['dcf12']}"; REF[f"tvsh{s_}"] = f"Val!{c}{T['tvsh']}"; REF[f"dcfpx{s_}"] = f"(Val!{c}{T['dcf12']}/px)"
REF["multB"] = f"Val!C{T['mult']}"
for k, r_ in _s.items():
    REF[f"l_{k}"] = f"Sens!A{r_}"; REF[f"e_{k}"] = f"Sens!J{r_}"
_t0 = _TS[0]
REF.update(s_wlo=f"Sens!D{_t0}", s_whi=f"Sens!F{_t0}", s_blo=f"Sens!D{_t0 + 1}", s_bhi=f"Sens!F{_t0 + 1}", s_glo=f"Sens!D{_t0 + 3}", s_ghi=f"Sens!F{_t0 + 3}", s_alt=f"Sens!F{_t0 + 8}",
           dcfnc="dcf_nc", stubB=f"Val!C{SDCF['stub']}", ncrB=f"Val!C{SDCF['ncr']}", rfus="rf_us")
_usd = next(r_ for r_ in range(_TS[0], _TS[1] + 1) if SE.cell(r_, 1).value == "美元口徑折現")   # v0.3.2：Sens 美元口徑列
REF.update(usd=f"Sens!C{_usd}", usdp=f"Sens!D{_usd}")
_tv = [f"Val!{c}{T['tvsh']}" for c in "BCD"]   # v0.3.2：終值占 EV 超過 80% 的情境（規格第 4 節；展示層標示）
REF["tvflag"] = ('IF(MAX(' + ",".join(_tv) + ')>0.8,"終值占 EV 超過 80% 的情境："&' + '&'.join(f'IF({t_}>0.8,"{s_} "&TEXT({t_},"0%")&" ","")' for s_, t_ in zip("ABC", _tv))
                 + '&"（DCF 價值多半來自終值，對 g、RONIC 與收斂期假設敏感）。","")')
_dc, _cc = f"(tp_w_dcf*({REF['dcftp']}-px))", f"((1-tp_w_dcf)*(tp_comp_adj-px))"
REF["dcfsh"], REF["compsh"] = f"({_dc}/({_dc}+{_cc}))", f"({_cc}/({_dc}+{_cc}))"
REF["src"] = f'IF({REF["dcfsh"]}>=0.6,"差距主要來自 DCF",IF({REF["dcfsh"]}<=0.4,"差距主要來自 Comparable","差距由 DCF 與 Comparable 約各半貢獻"))'
REF["mulg"] = f"({REF['pmadj']}/{REF['pxpe']}-1)"
REF["updesc"] = f'IF({REF["upw"]}>0.25,"遠高於",IF({REF["upw"]}>=0,"略高於",IF({REF["upw"]}>=-0.25,"低於","遠低於")))'
REF["verdict"] = f'IF({REF["upw"]}>0.25,"現價低於估值、有明顯上檔",IF({REF["upw"]}>=0,"現價大致合理",IF({REF["upw"]}>=-0.25,"現價高於估值","現價明顯高於估值")))'
_np = "(" + "+".join(f"({REF[f'dcfpx{s_}']}>1)" for s_ in SC) + ")"
REF["dcfwho"] = f'IF({_np}=3,"三種情境的 DCF 價值都高於現價",IF({_np}=0,"三種情境的 DCF 價值都低於現價","DCF 價值有的情境高於、有的低於現價"))'
_SP = "+0.0%;−0.0%"

def _fx(t):
    def lit(x):
        x = x.replace('"', '""')
        return [f'"{x[i:i + 200]}"' for i in range(0, len(x), 200)]
    parts, pos = [], 0
    for m in _re.finditer(r"\{([A-Za-z0-9_]+)(?:\|([^}]*))?\}", t):
        parts += lit(t[pos:m.start()])
        e = REF[m.group(1)]
        parts.append(e if m.group(2) is None else f'TEXT({e},"{m.group(2).replace("SP", _SP)}")')
        pos = m.end()
    parts += lit(t[pos:])
    parts = [x for x in parts if x != '""']
    return t if len(parts) == 1 and parts[0].startswith('"') else "=" + "&".join(parts)

PRES = [
 "命題：味之素的本業是調味料、冷凍食品與生技醫藥；AI 曝險集中在ファンクショナルマテリアルズ（FM）的 ABF 增層膜——營收約 7%、事業利益約三分之一。股價取決於 ABF 能長多大，以及現價已經替 ABF 付了多少。三個問題如下。",
 "命題 1：ABF 能長多大：AI 加速器顆數、每顆載板面積層數、份額與漲價，讓 FM 營收與事業利益在 FY2027 成長多少、對集團 EPS 貢獻多少。",
 "命題 2：現價替 ABF 付了多少：本業以食品同業倍數計價後，現價其餘部分隱含的 ABF 價值與倍數是多少。",
 "命題 3：現價合不合理：以標準版 DCF（自由現金流折現）與食品同業倍數估出的每股價值，和現價相比是高還是低。",
 "重點 1：合理價值{updesc}現價，{src}。加權目標價約 ¥{tpw|#,##0}（DCF ¥{dcftp|#,##0} 占 {wdcf|0%}、Comparable〔相對倍數調整後〕¥{compadj|#,##0} 占其餘），比現價 ¥{px|#,##0} {upw|SP}，投資建議{rec}（級距 {grade}）。把加權目標價與現價的差距拆成兩個來源（各自權重 ×（目標價 − 現價））：DCF 占 {dcfsh|0%}、Comparable 占 {compsh|0%}。Comparable 的差距來自倍數、不是盈餘：12 個月後前瞻 EPS ¥{feps|#,##0.0} 比共識 {geps|SP}，但食品同業倍數（相對倍數調整後 {pmadj|0.0} 倍）比現價隱含的 {pxpe|0.0} 倍低 {mulg|0%;0%}。DCF 比現價 {dcfup|SP}：{dcfwho}。",
 "重點 2：ABF 是成長來源，但仍只占集團一部分。情境 A 的 FM 營收由 FY2026 公司預測 ¥{fm26|#,##0.0}B 升到 FY2027 ¥{fm27|#,##0.0}B（其中 AI 用 ¥{fmai26|#,##0.0}B → ¥{fmai27|#,##0.0}B），FM 事業利益 ¥{fmbp27|#,##0.0}B 占全公司 {fmbpsh27|0%}、每股貢獻 ¥{fmeps27|#,##0.0}（EPS ¥{eps27|#,##0.0}）。漲價、產能與 AI 占比為假設（Assumed）。",
 "重點 3：EPS 對 ABF 價格與產能上限最敏感（情境 A 的 FY2028 EPS；敏感度測試，不是門檻）：「{l_plo}」「{l_phi}」為 {e_plo|SP}／{e_phi|SP}，「{l_clo}」「{l_chi}」為 {e_clo|SP}／{e_chi|SP}；本業「{l_msf}」為 {e_msf|SP}。",
 "重點 4：v0.3 起 DCF 改為標準版（無槓桿自由現金流折現）：DCF 目標價 ¥{dcftp|#,##0}，三種情境推到 12 個月後為 A ¥{dcfA|#,##0}、B ¥{dcfB|#,##0}、C ¥{dcfC|#,##0}（{dcfwho}）。折現率 WACC {wacc|0.0%}：股權成本 {ke|0.0%}＝rf {rfv|0.00%}（日本 10 年期公債）＋ β {beta|0.00} ×（ERP {erp|0%} ＋ 日本國家風險溢酬 {crp|0.00%}）；β 為自算 JP 食品與 JP 電子材料產業無槓桿 β 依 FM 事業利益占比 {wfm|0%} 加權、Blume 調整（{betau|0.000}）後重新加槓桿，ERP 5%、RONIC＝WACC 與加權方法為假設；永續成長率 g＝{gtv|0.00%}。食品 β 低使 WACC 低於舊法的 10%，DCF 目標價因此上升；情境 B 的終值現值占企業價值 {tvshB|0%}、隱含終值 EV／EBITDA 約 {multB|0.0} 倍。以三情境加權目標價為基準，WACC −1pt／＋1pt 使加權目標價變動 {s_wlo|SP}／{s_whi|SP}，g −1pt／＋1pt 為 {s_glo|SP}／{s_ghi|SP}，β 低／高案例（β_u ∓ 0.3 各自重新加槓桿）為 {s_blo|SP}／{s_bhi|SP}（敏感度測試，不是門檻）。"
 "淨現金與投資以 2026-06-30 資產負債表 ¥{dcfnc|#,##0}B 加 stub 期 FCFF（FY2026 FCFF × stub_frac；情境 B ¥{stubB|#,##0}B）滾到評價日為 ¥{ncrB|#,##0}B。"
 "美元口徑折現（Sens 案例：rf 改用美國 10 年期 {rfus|0.00%}、CRP 照基準、g＝MIN(3%, {rfus|0.00%})，現金流與匯率不變）下加權目標價 ¥{usd|#,##0}，相對基準 {usdp|SP}；基準仍採日本利率＋固定匯率。{tvflag}"]
ANS = [
 "ABF 能長多大？答：依本模型情境 A，FM 營收 FY2027 為 ¥{fm27|#,##0.0}B（FY2026 公司預測 ¥{fm26|#,##0.0}B），AI 用 ¥{fmai27|#,##0.0}B；FM 事業利益占全公司由 FY2026 的 {fmbpsh26|0%} 升到 {fmbpsh27|0%}，每股貢獻 ¥{fmeps27|#,##0.0}。FY2028 EPS 情境 A／B／C 為 ¥{eps28A|#,##0.0}／¥{eps28B|#,##0.0}／¥{eps28C|#,##0.0}（顆數與 ABF 價格乘數依情境；漲價、產能、AI 占比為假設）。",
 "現價替 ABF 付了多少？答：本業（FM 以外）的未來 12 月 EPS 以食品同業倍數計價約 ¥{cval|#,##0}，現價其餘 ¥{ival|#,##0}（占現價 {ish|0%}）都是 ABF（FM）的價值，等於 FM 未來 12 月 EPS 的 {ipe|0.0} 倍，電子材料同業中位數 {emed|0.0} 倍；以電子材料同業倍數計 FM 的分部加總為 ¥{sotp|#,##0}（{sotpup|SP}）。分部加總只作參考，不進目標價。",
 "現價合不合理？答：依本模型{verdict}。標準版 DCF 目標價 ¥{dcftp|#,##0}（{dcfup|SP}），{dcfwho}；Comparable（食品同業倍數，相對倍數調整後）¥{compadj|#,##0}（{upcadj|SP}）。加權目標價 ¥{tpw|#,##0}，比現價 {upw|SP}，投資建議{rec}。市場共識目標價 ¥{cons|#,##0}，{cgap}：本模型 12 個月後前瞻 EPS 比共識 {geps|SP}，共識目標價隱含的前瞻 P/E 比食品同業中位數高 {gmul|SP}，等於市場替 ABF 付了遠高於食品同業的倍數。"]
RD["A1"] = RD["A1"].value + "；v0.3（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數）＋HTML 新版面 v1.0（README 新增展示層與命題回覆）；Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；v0.3.1（2026-11-05）：β 改自算產業 β（JP 食品＋JP 電子材料依 FM 占比加權，Blume 調整），評價日統一為 2026-12-31；v0.3.2（2026-11-05）：第二輪修正（6-30 淨現金以 stub 期滾到評價日、β 敏感度 ±0.3、Sens 美元口徑折現）"
_r = RD.max_row + 2
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.3 新增，上方原列保留給下游）").font = BOLD; _r += 1
# v0.3.2：與共識目標價差距的來源以公式判定（EPS 因素 vs 倍數因素，取對數後比大小）；README 上方「重點觀察 1」改為公式（原 2026-10-29 調整前撰寫的「主因」說明刪除）
REF["gmul"] = _V("gmul")
REF["cgap"] = f'IF(ABS(LN(1+{_V("gmul")}))>ABS(LN(1+{_V("geps")})),"與共識目標價的差距主要來自倍數","與共識目標價的差距主要來自 EPS")'
REF["dcfsh"] = f"((tp_w_dcf*({_V('dcftp')}-px))/(tp_w_dcf*({_V('dcftp')}-px)+(1-tp_w_dcf)*(tp_comp_adj-px)))"
REF["upsrc"] = f'IF({REF["dcfsh"]}>=0.6,"差距主要來自 DCF",IF({REF["dcfsh"]}<=0.4,"差距主要來自 Comparable","差距由 DCF 與 Comparable 約各半貢獻"))'
_o1 = next(r_ for r_ in range(1, RD.max_row + 1) if str(RD.cell(r_, 1).value or "").startswith(("重點觀察 1：", '="重點觀察 1：')))
RD.cell(_o1, 1, _fx("重點觀察 1：目標價 ¥{tpw|#,##0}（DCF {dcftp|#,##0}、Comparable（相對倍數調整後）{compadj|#,##0}；DCF 權重 {wdcf|0%}），相對現價 ¥{px|#,##0}（2026-10-23）為 {upw|SP}，投資建議{rec}。"
                    "{upsrc}（DCF 占 {dcfsh|0.0%}；各自權重 ×（目標價 − 現價））。{cgap}：本模型前瞻 EPS 比共識 {geps|SP}、共識目標價隱含倍數比食品同業高 {gmul|SP}。"))
for t_ in PRES:
    RD.cell(_r, 1, _fx(t_)); _r += 1
_r += 1
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上方三個命題）").font = BOLD; _r += 1
for t_ in ANS:
    RD.cell(_r, 1, _fx(t_)); _r += 1
_r += 1
RD.cell(_r, 1, "v0.3（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。新增 DCF 頁（三情境 FCFF 明細與檢查列；評價日 2026-06-30，FY2026 只計評價日後三季）；Val 0.1 節改為折現率與 DCF 彙總（0.2 節以後列號不變）；Sens 目標價參數改為標準案例；In 頁新增 M2 節（日本 rf、Damodaran 食品與特用化學無槓桿 β 依事業利益占比加權後重新加槓桿、ERP＋日本 CRP、資本結構、穩態比例），刪除 tp_r、tp_roe。HTML 改用個股 HTML 版面規格 v1.0。PL、Q、Units、Recon 頁與 v0.2.1 相同。"); _r += 1
_SPV = '"+0.0%;−0.0%"'
RD.cell(_r, 1, "=" + "&".join([
    '"v0.3.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：JP 食品同業 5 家無槓桿 β 中位數 "', 'TEXT(beta_u_food,"0.00")', '"、JP 電子材料／載板同業 5 家 "', 'TEXT(beta_u_chem,"0.00")',
    '"（Yahoo Finance 月報酬 2021-11～2026-10 對日經 225 迴歸，Hamada 去槓桿、現金修正），依 FM 事業利益占比 "', 'TEXT(beta_w_fm,"0.0%")', '" 加權後 Blume 調整為 "', 'TEXT(beta_u,"0.000")',
    '"，依市值 D/E 重新加槓桿為 "', 'TEXT(beta,"0.00")', '"，取代 Damodaran Global 食品與特用化學表；β 低／高案例改為加權中位數 × 0.8 再 Blume／中位數未調整（原低案例＝只用食品 β）。"',
    '"評價日由 2026-06-30 統一為 2026-12-31（FY2026 只計評價日後一季；目標時點 2027-12-31；淨現金與投資沿用 2026-06-30 資產負債表）。"',
    '"新舊對照（v0.3 → v0.3.1）：β 0.77 → "', 'TEXT(beta,"0.00")', '"，WACC 6.9% → "', 'TEXT(wacc,"0.0%")',
    '"，DCF 目標價 2,692 → 2,663，加權目標價 3,096 → 3,082，上檔空間 −38.9% → −39.1%，投資建議 放空 → 放空。其他模型內容不變。"'])); _r += 1   # v0.3.2：v0.3.1 當時的結果改為固定文字（原公式指到 Val!B22 等列，stub 期多兩列後已移位）
RD.cell(_r, 1, _fx("v0.3.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2）。β 組別不變（JP 食品 0.437 與 JP 電子材料 1.389 依 FM 事業利益占比 {wfm|0.0%} 加權＝{betau|0.000}，重新加槓桿 {beta|0.00}）；β 敏感度改 β_u ∓ 0.3 重新加槓桿（原高案例低於基準）。"
                   "淨現金與投資：資產負債表日 2026-06-30（最新已公布一季）¥{dcfnc|#,##0}B，以 stub 期滾到評價日 2026-12-31＝＋stub 期 FCFF（FY2026 FCFF × stub_frac；情境 B ¥{stubB|#,##0}B）− 已宣告股利／買回 0＝¥{ncrB|#,##0}B（v0.3.1 直接沿用 6-30 值）。"
                   "模型沒有逐年資本支出／折舊列，維持 FY2025 實績占營收比例。Sens 新增「美元口徑折現」：加權目標價 ¥{usd|#,##0}（{usdp|SP}）。"
                   "新舊對照（v0.3.1 → v0.3.2）：β 0.81 → {beta|0.00}，WACC 7.0% → {wacc|0.0%}，DCF 目標價 2,663 → {dcftp|#,##0}，加權目標價 3,082 → {tpw|#,##0}，上檔空間 −39.1% → {upw|SP}，投資建議 放空 → {rec}。")); _r += 1
RD.cell(_r, 1, "HTML 橫幅輔助值（只引用 Val 頁；不進任何計算）").font = BOLD; _r += 1
RD.cell(_r, 1, "上檔空間：Comparable 目標價（相對倍數調整後）"); RD.cell(_r, 2, "=tp_comp_adj/px-1").number_format = "0.0%"; _r += 1
wb.save(sys.argv[1])
if os.environ.get("AJI_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    _json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW, "QR": QR, "Q_NOTE": Q_NOTE,
                "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE,
                "T": T, "T10": T10, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "NCASES": len(cases),
                "S0": S0, "SR": SR, "EP1": EP1, "EP2": EP2, "FCF_ROW": FCF_ROW},
               open(os.environ["AJI_ROWMAP"], "w"), ensure_ascii=False, indent=1)
