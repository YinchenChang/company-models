# KLA（KLAC）收支與估值模型 v0.4.1（AI 半導體 Project 步驟 5：設備與材料——量測與檢測）
# 用法：python3 klac/build_klac.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 7 月到隔年 6 月：FY2027＝2026-07～2027-06（一半在 CY2026、一半在 CY2027）。金額：百萬美元（US$M）；EPS 為美元（2026-06 的 10 比 1 分割調整後）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import fcff_dcf as FD   # v0.4：標準版 DCF（無槓桿自由現金流＋WACC）共用模組

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- 讀入設備市場模型成品（唯讀；不修改） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
EQK = {"先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕": "L1", "其他代工與邏輯": "L2", "其他前段": "L3",
       "DRAM（含 HBM）〔AI〕": "M1", "NAND": "M2", "先進封裝〔AI〕": "P"}
EQ = {}; blk = None; SEMI_WFE = None
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    for s in SC:
        if lab.startswith(f"5{s}. "): blk = s
    if lab.startswith("6. "): blk = None
    if blk and lab in EQK:
        EQ[(blk, EQK[lab])] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if blk and lab == f"設備總額（情境 {blk}）":
        EQ[(blk, "T")] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if lab == "SEMI 晶圓前段設備（WFE）":
        SEMI_WFE = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert len(EQ) == 21 and SEMI_WFE, sorted(EQ)
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：百萬美元（US$M）；設備市場：十億美元（US$B）。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司原文（SEC 8-K／10-K）算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計用。SemiAnalysis 未使用。共識只放 Recon，不作校準"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
CALL = "KLA FQ4-26 法說（2026-07-28；MarketBeat 逐字稿）"
sec("A. CY2026 錨點（上半年實績＋公司對下半年的說法）")
inp("h2up", "CY2026 下半年營收較上半年增幅（FQ1-27 實績＋FQ2-27 營收指引）", "%", "=(Q!I4+g_q2_rev_lo)/(Q!E4+Q!F4)-1", "=(Q!I4+g_q2_rev)/(Q!E4+Q!F4)-1", "=(Q!I4+g_q2_rev_hi)/(Q!E4+Q!F4)-1", "Derived", "Q 頁：（FQ1-27 實績 4,136.0＋FQ2-27 指引 4,700±200）÷ 上半年實績 7,072.7 − 1；KLA FQ1-27 新聞稿（SEC 8-K，2026-10-28）", "v0.3 前為公司口述「下半年約比上半年高 20%」（FQ4-26 法說，Interested-party）；低／高＝FQ2-27 指引區間")
inp("g_svc_h2", "CY2026 下半年服務營收較上半年增幅", "%", 0.03, 0.06, 0.09, "Assumed", "無：公司稱服務 2026 年成長在長期目標 13–15% 的下緣", "6%≈年增 13% 的半年版；只影響錨點中系統與服務的拆分")
inp("g_oth26", "SPTS＋PCB 兩事業 CY2026 年增", "%", 0.15, 0.25, 0.30, "Interested-party", CALL + "：Orbotech 合計產品 CY2026 成長 25% 以上", "第二來源：未找到。CY2025 基期取 Q 頁四季合計")
inp("ap26", "先進封裝製程控制營收 CY2026", "US$M", 1000, 1100, 1400, "Interested-party", CALL + "：約 US$1.1B，年增 70% 以上", "handsoff substack（2026-05-31，引公司 FQ3-26）：FY2026 1.3–1.4B（會計年度口徑，與 1.1B 的日曆年口徑不同，未解）；高＝1.4B")
inp("ap_g26", "先進封裝製程控制營收 CY2026 年增", "%", None, 0.70, None, "Interested-party", CALL + "：年增 70% 以上（取 70%）", "用於推 CY2025 基期＝ap26÷1.7")
inp("k_ap", "先進封裝製程控制營收成長 ÷ 先進封裝設備市場成長（倍數）", "x", 1.0, 2.0, 2.5, "Interested-party", CALL + "：成長率接近封裝市場的 2 倍", "低＝與市場同速、高＝2.5 倍（Assumed）；下滑時同樣放大（β 解讀，Assumed）")
sec("B. 製程控制強度（系統營收 ÷ 加權前段 WFE；以 CY2026 錨點校準）")
inp("fl26", "CY2026 半導體製程控制系統營收中代工邏輯占比", "%", 0.60, 0.63, 0.66, "Interested-party", CALL + "：2026 年代工邏輯「low to mid-60s」；FQ1-27 為 73%／記憶體 27%", "第二來源：未找到。取 63%；用來推記憶體相對強度")
inp("m_mem", "記憶體 WFE 的製程控制強度 ÷ 代工邏輯（相對倍數）", "x", "=(1-fl26_hi)/fl26_hi*Units!D8/Units!D11", "=(1-fl26)/fl26*Units!D8/Units!D11", "=(1-fl26_lo)/fl26_lo*Units!D8/Units!D11", "Derived", "（1−代工邏輯占比）÷ 代工邏輯占比 × CY2026 代工邏輯 WFE ÷ 記憶體 WFE（設備模型情境 A）", "代工邏輯 WFE＝先進邏輯＋其他代工邏輯＋其他前段（含光罩，KLA 光罩檢測歸此）；低欄用 fl26 高值")
inp("kla_wfe25", "CY2025 WFE（公司口徑，含先進封裝）", "US$B", None, 120, None, "Interested-party", CALL + "：2025 年約 US$120B", "設備模型 2025 前段＋先進封裝＝" + f"{EQ[('A','L1')][0][0]+EQ[('A','L2')][0][0]+EQ[('A','L3')][0][0]+EQ[('A','M1')][0][0]+EQ[('A','M2')][0][0]+EQ[('A','P')][0][0]:.1f}（相近）")
inp("share_up", "KLA 占 WFE 比重 2025→2030 的提升", "pt", 0.0, 0.015, 0.03, "Interested-party", CALL + "：公司 2030 模型稱市占較 2025 提升 150bp（±25）", "第二來源：未找到。低＝不提升、高＝300bp（Assumed）；公司未說明分母，本模型解讀為「半導體製程控制營收 ÷ WFE」")
inp("pc_u", "製程控制強度年增（先進製程與先進封裝檢測強度上升）", "%/年", "=((s25+share_up_lo)/s25)^(1/5)-1", "=((s25+share_up)/s25)^(1/5)-1", "=((s25+share_up_hi)/s25)^(1/5)-1", "Derived", "＝[(CY2025 占比＋提升)÷CY2025 占比]^(1/5)−1；CY2025 占比＝半導體製程控制營收 ÷ 公司 WFE（Units）", "自 CY2027 起套用於 CY2026 校準的強度")
inp("kla_sh_pc", "KLA 製程控制市占（顯示用：換算製程控制占 WFE）", "%", 0.53, 0.58, 0.63, "Interested-party", "handsoff substack 2026-05-31（引 BofA CY2025）：製程控制市占約 58%、晶圓檢測 81.9%（賣方）", "第二來源：未找到。只用於顯示，不影響營收")
sec("C. 中國（重要變數；基準不另扣，風險放 Sens）")
inp("chn_sh26", "中國占營收 FY2026", "%", None, 0.30, None, "Verified", "KLA FY2026 10-K：中國 30%（FY2025 33%、FY2024 43%）", "handsoff substack：FQ3-26 單季 24%（約 US$820M）")
inp("chn_sh25", "中國占營收 FY2025", "%", None, 0.33, None, "Verified", "KLA FY2026 10-K", "—")
inp("chn_sh24", "中國占營收 FY2024", "%", None, 0.43, None, "Verified", "KLA FY2026 10-K", "—")
inp("chn_cut", "CY2027 起中國系統營收下降比例（出口管制）", "%", 0.25, 0.0, 0.50, "Assumed", "無：10-K 稱出口許可未獲准已傷及訂單；BIS 子公司規則暫停一年至 2026-11", "基準 0（設備模型的 WFE 已含中國）；Sens 列 25%／50%")
sec("D. 服務與 SPTS＋PCB（會計年度，前一年 × 成長率）")
for s, v, o in (("A", 0.14, 0.12), ("B", 0.12, 0.08), ("C", 0.10, 0.00)):
    inp(f"g_svc_{s}", f"服務營收 FY2027–29 年增（{s}）", "%", v - 0.02, v, v + 0.02, "Assumed", "公司長期目標 13–15%，2027 年預期在上緣；約 80% 為合約制（" + CALL + "）", "FY2026 實績 +16.5%（Q 頁）；C 取裝機量成長放緩")
    inp(f"g_oth_{s}", f"SPTS＋PCB 營收 FY2027–29 年增（{s}）", "%", o - 0.04, o, o + 0.04, "Assumed", "公司：EPC／PCB 長期成長 high single digit；CY2026 Orbotech 合計 >25%", "FY2026 實績 +11.5%（Q 頁）")
sec("E. 毛利與營業費用（不寫死利益率）")
inp("gm_inc", "增量毛利率（毛利變動 ÷ 營收變動；相對 FY2026）", "%", 0.60, "=(a26_gp-a25_gp)/(a26_rev-a25_rev)", 0.65, "Derived", "兩點：FY2025、FY2026 實績（Q 頁）", "公司：增量毛利率 60–65% 至 2027（非 GAAP；" + CALL + "）＝低／高；記憶體定價逆風約 100bp")
inp("opex_q1", "營業費用 FQ1-27（R&D＋SG&A）", "US$M", None, 690, None, "Interested-party", CALL + "：約 US$690M", "第二來源：未找到")
inp("opex_step", "之後每季營業費用增加", "US$M／季", 15, 17.5, 20, "Interested-party", CALL + "：每季增加約 US$15–20M", "取中點")
inp("opex27", "營業費用 FY2027（4×FQ1＋6×每季增額）", "US$M", None, "=4*opex_q1+6*opex_step", None, "Derived", "690＋707.5＋725＋742.5", "三情境相同（FY2027 已大致決定）")
for s, g in (("A", 0.08), ("B", 0.05), ("C", 0.02)):
    inp(f"g_opex_{s}", f"營業費用 FY2028–29 年增（{s}）", "%", g - 0.03, g, g + 0.03, "Assumed", "無：FY2026 +11.4%、FY2027 指引隱含 +7.6%", "—")
sec("F. 營業外、稅、股數")
inp("oth_exp", "營業外淨額（利息與其他）每年", "US$M", -200, -100, 0, "Interested-party", CALL + "：其他收支約每季 US$25M 費用", "FY2026 實績 −54.9（利息 −284.4＋其他 +229.6）；低／高 Assumed")
inp("tax", "有效稅率", "%", 0.13, 0.145, 0.17, "Interested-party", CALL + "：規劃稅率 14.5%", "FY2026 實績 13.8%（Q 頁）")
inp("shares27", "稀釋股數 FY2027", "百萬股", None, 1311.1, None, "Verified", "KLA FQ1-27 新聞稿（SEC 8-K，2026-10-28）：FQ1-27 稀釋加權股數 1,311.1M", "FY2026 稀釋 1,319.6M、FQ4-26 1,315.0M；FQ1-27 回購 935.8")
inp("bb", "稀釋股數年減（買回）FY2028–29", "%", 0.0, "=1-a26_sh/a25_sh", 0.025, "Derived", "FY2025→FY2026 稀釋股數變化（Q 頁）", "低／高 Assumed")
sec("G. 公司指引、共識（只供 Recon）、股價、循環位置、目標價權重")
inp("g_q2_rev", "FQ2-27 營收指引", "US$M", 4500, 4700, 4900, "Interested-party", "KLA FQ1-27 新聞稿（SEC 8-K，2026-10-28）：US$4.7B±0.2B", "FQ1-27 實績 4,136.0（指引 4.0B±0.2B）")
inp("g_q2_gm", "FQ2-27 GAAP 毛利率指引", "%", 0.608, 0.618, 0.628, "Interested-party", "同上", "非 GAAP 62.5%±1.0%")
inp("g_q2_eps", "FQ2-27 GAAP 稀釋 EPS 指引", "US$", 1.30, 1.41, 1.52, "Interested-party", "同上", "非 GAAP 1.44±0.11")
inp("cons_q2_rev", "共識 FQ2-27 營收", "US$M", None, 4490, None, "Interested-party", "Yahoo Finance（2026-10-25 讀取；21 位分析師）", "第二來源：未找到")
inp("cons_rev27", "共識營收 FY2027", "US$M", None, 18140, None, "Interested-party", "Yahoo Finance（24 位）：18.14B", "stockanalysis（2026-10-20）：18.1B")
inp("cons_eps27", "共識 EPS FY2027（非 GAAP）", "US$", None, 5.44, None, "Interested-party", "Yahoo Finance（25 位）：5.44", "stockanalysis（2026-10-20，24 位）：5.46")
inp("cons_rev28", "共識營收 FY2028", "US$M", None, 21570, None, "Interested-party", "Yahoo Finance（21 位）：21.57B", "第二來源：未找到（stockanalysis 需付費）")
inp("cons_eps28", "共識 EPS FY2028（非 GAAP）", "US$", None, 6.74, None, "Interested-party", "Yahoo Finance（22 位）：6.74", "第二來源：未找到")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；KLA 的 FY(t) 對照日曆年 t（理由見 Val 第 4 節）" if s == "A" else "同上")
inp("px", "股價（stockanalysis 2026-10-23 收盤）", "US$", None, 145.61, None, "Verified", "stockanalysis.com 歷史股價：2026-10-23 收盤 145.61", "digrin：145.61；tipranks：2026-10-22 收盤 147.83（10-22 單日 −9.5%，原因未查明）")
inp("tp_wA", "目標價情境權重 A", "%", None, 0.3, None, "Assumed", "Andy 可改", "三個權重合計應為 100%（Val 第 0 節檢查）")
inp("tp_wB", "目標價情境權重 B", "%", None, 0.5, None, "Assumed", "Andy 可改", "—")
inp("tp_wC", "目標價情境權重 C", "%", None, 0.2, None, "Assumed", "Andy 可改", "—")

# v0.4：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py）的 In 參數。評價日 v0.4.1 起改為 2026-12-31（原 FY2026 年底 2026-06-30）；股價 2026-10-23；資產負債為 Q1 FY2027（2026-09-30）
UST10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10（同一序列，https://fred.stlouisfed.org/series/DGS10）：10-28～10-30 5.05%／4.98%／4.94% 與財政部表一致；與股價同日（2026-10-23）；同 NVIDIA v0.3", "無風險利率（美國 10 年期公債殖利率，2026-10-23）")
DAM_US = ("自算產業 β：US 半導體設備 同業 7 家中位數（AMAT、LRCX、KLAC、TER、FORM、COHU、KLIC；Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）無槓桿 β 1.75；以公司 D/E（市值）與稅率依 Hamada 重新加槓桿：β＝βU ×（1＋(1−t) × D/E）",
          "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）：採用無槓桿 β 1.503。低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3），各自依公司 D/E 與稅率重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）；取代 Damodaran 美國產業表")
CRP_US = (0.0, "Assumed", "美國等成熟市場：0（同 NVIDIA v0.3；Damodaran 2026-01-05 表美國 Aa1、CRP 0.23%，本模型不加）", "非美國市場：Damodaran 最新 Country Risk Premium 表（Andy 2026-11-05 20:45）")
SA = "stockanalysis.com KLAC 財報（US$M，2026-11-05 讀取）"
SAQ = "stockanalysis.com KLAC 季度資產負債表（Q1 FY2027，2026-09-30；US$M；2026-11-05 讀取）https://stockanalysis.com/stocks/klac/financials/balance-sheet/?p=quarterly"
RELEV = "*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))"
for _row in FD.in_rows(dict(
        date="2026-12-31（v0.4.1 起；FY2027 只計評價日後 6 個月；淨現金為 Q1 FY2027 資產負債表 2026-09-30）", amt="US$M", shu="百萬股", ccy="US$",
        rf=(0.0500,) + UST10,
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=SAQ + "；KLA 10-Q（Q1 FY2027，最新已公布一季）",
                  div=(0, "Verified", "KLA 新聞稿：最近一次宣告 2026-08-06（每股 0.23 美元，2026-09-01 發放，在資產負債表日之前）；2026-12 發放的下一次建置日 2026-11-05 尚未宣告（去年同期 2025-11-06 宣告），不扣（https://ir.kla.com/news-events/press-releases/detail/520/kla-declares-regular-cash-dividend）",
                       "買回：無已宣告的 stub 期金額，不扣（股數沿用評價用稀釋股數，與 FCFF 一致）")),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived") + DAM_US,
        crp=CRP_US,
        kd=("=rf+0.008", "Assumed", "rf ＋ 0.8 個百分點信用利差（科磊長期債評等 A 級）", "負債占資本約 3%，對 WACC 影響可忽略"),
        tax_w=("=tax", "Interested-party", "＝In 頁 tax（公司規劃稅率 14.5%）", "NOPAT 各年也用 tax（與 PL 淨利＝（營業利益＋營業外）×（1 − tax）同一稅率）"),
        debt=(5723, "Verified", SAQ + "：長期債務 5,723（無一年內到期）", "KLA 10-Q（Q1 FY2027）"),
        lease=("=54.01+215.34", "Verified", SAQ + "：租賃負債 流動 54.01＋非流動 215.34", "KLA 10-Q（Q1 FY2027）"),
        cash=("=1532+3209", "Verified", SAQ + "：現金及約當現金 1,532＋短期投資 3,209", "KLA 10-Q（Q1 FY2027）"),
        inv=(0, "Verified", SAQ + "：無長期投資列", "—"),
        mi=(0, "Verified", SAQ + "：無少數股權", "—"),
        dps=("=1058/shares27", "Derived", SA + "：FY2026 已付股利 1,058 ÷ FY2027 稀釋股數 1,311.1M（每股約 0.81 美元）", "評價日到目標時點之間假設同此水準（Assumed；公司每年調升股利，偏低）"),
        sh=("=shares27", "Verified", "＝In 頁 shares27（FQ1-27 稀釋加權股數 1,311.1M，SEC 8-K）", "回購由 FCFF 支應，評價用最近一季稀釋股數，不另推未來股數"),
        fc_da=(None, "=(401.73/9812+394.09/a25_rev+393.98/a26_rev)/3", None, "Derived", SA + "：近三年折舊攤銷 ÷ 營收：FY2024 401.73／9,812、FY2025 394.09／FY2025 營收、FY2026 393.98／FY2026 營收（Q 頁）", "KLA 10-K 現金流量表"),
        fc_cx=(None, "=(277.38/9812+335.26/a25_rev+375.95/a26_rev)/3", None, "Derived", SA + "：近三年資本支出 ÷ 營收：FY2024 277.38、FY2025 335.26、FY2026 375.95（US$M）", "Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=fc_da*1.1", "Assumed", "長期穩態＝折舊 ÷ 營收 × 1.1（近三年資本支出約為折舊的 0.85–0.95 倍；穩態需 ≥ 折舊 ×（1＋g））", "一致性檢查：FY2032 資本支出 ≥ 折舊 ×（1＋g）（DCF 頁）"),
        fc_wc=(0.15, "=((1902+3035+124.97+465.1-359.49-803.53-2134-394.25)/9812+(2369+3212+201.05+421.97-458.51-879-2001-355.67)/a25_rev+(3021+3649+224.52+585.45-623.67-1017-1967-241.93)/a26_rev)/3", 0.35, "Derived",
               "近三年營運資金 ÷ 營收平均：營運資金＝應收＋存貨＋預付＋其他流動資產 − 應付 − 應計費用 − 流動與長期預收款（stockanalysis.com KLAC 季度資產負債表 6 月底欄）", "Sens 低／高 15%／35%（Assumed）"),
        )):
    rows.append(_row)
# 第二輪修正（2026-11-05）：β 以 β_u 重新加槓桿寫成公式；β 敏感度＝β_u ± 0.3（tools/fcff_dcf.py v2 write_sens 讀 beta_u、beta_de、beta_du）
inp("beta_u", "產業無槓桿 β（採用）：自算 US 半導體設備 同業中位數 Blume 調整", "x", None, 1.503, None, "Derived", "自算產業 β：US 半導體設備 同業中位數（Yahoo Finance 月報酬 2021-11～2026-10，對本地大盤迴歸，Hamada 去槓桿、現金修正）；見 In 頁 beta 列來源", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 稀釋股數 sh_v)", "與 WACC 權重同一口徑")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代上一版的「中位數 × 0.8 再 Blume／中位數未調整」")

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
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    defname(key, f"In!$E${r}"); INROW[key] = r
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 10, 11, 10, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Q：實績（百萬美元；EPS 為分割調整後美元） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（百萬美元；EPS 為美元，2026-06-12 除權的 10 比 1 分割調整後）。會計年度 7–6 月：FY2026＝2025-07～2026-06。藍字＝輸入（SEC 8-K 新聞稿）；黑字＝公式（Derived）"
Q["A2"] = "v0.3（2026-10-31）：I 欄滾入 FY2027 第一季（2026-07～09；KLA FQ1-27 新聞稿（SEC 8-K，2026-10-28））；下方指引改為 FQ2-27（2026-10～12）"
hdr(Q, 3, ["項目", "FQ3-25（2025-01～03）", "FQ4-25（2025-04～06）", "FY26 上半（2025-07～12，FQ1＋FQ2）", "FQ3-26（2026-01～03）", "FQ4-26（2026-04～06）", "FY2025 全年", "FY2026 全年", "FQ1-27（2026-07～09）", "標記", "來源1（公司）", "來源2／說明"])
S_Q326 = "FQ3-26 新聞稿（SEC 8-K，2026-04-29）"; S_Q426 = "FQ4-26 新聞稿（SEC 8-K，2026-07-28）"
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0.0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGHI", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 10, tag); Q.cell(r, 11, s1); Q.cell(r, 12, s2)
H1 = lambda a, b: f"={a}-{b}"   # 九個月 − FQ3
qrow(4, "rev", "營收", [3063.0, 3174.741, H1(9921.9, 3415.1), 3415.1, 3657.556, 12156.162, 13579.476, 4136.015], "Verified", "FQ3-25：" + S_Q326 + "前年同期欄；FQ4-25、全年：" + S_Q426 + "；上半＝九個月 9,921.9 − FQ3-26", "stockanalysis 年度：FY2025 12,156、FY2026 13,579；FQ2-26 單季 3,297（FQ3-26 新聞稿）")
qrow(5, "prod", "產品營收", [2393.8, 2472.182, H1(7616.4, 2640.3), 2640.3, 2837.151, 9472.854, 10453.537, 3285.469], "Verified", "同上", "第二來源：未找到（年度服務占比 23%，10-K）")
qrow(6, "svc", "服務營收", [669.2, 702.559, H1(2305.5, 774.8), 774.8, 820.405, 2683.308, 3125.939, 850.546], "Verified", "同上", "FY2026 10-K：服務約占營收 23%（3,125.9÷13,579.5＝23.0%）")
qrow(7, "spc", "  事業別：半導體製程控制（含服務）", [2738.8, 2877.647, H1(8988.0, 3083.9), 3083.9, 3256.781, 10947.359, 12244.733, 3754.254], "Verified", "同上（事業別表）", "—")
qrow(8, "spts", "  事業別：特殊半導體製程（SPTS）", [156.5, 141.866, H1(424.4, 164.0), 164.0, 159.704, 587.107, 584.064, 139.487], "Verified", "同上", "—")
qrow(9, "pcb", "  事業別：PCB 與元件檢測", [168.6, 154.106, H1(509.3, 167.6), 167.6, 241.110, 621.721, 750.415, 242.151], "Verified", "同上", "—")
qrow(10, "fx", "  匯率影響等（營收 − 三事業，Derived）", [f"={c}4-{c}7-{c}8-{c}9" for c in "BCDEFGHI"], "Derived", "—", "公司事業別表另列 FX；四捨五入差")
qrow(11, "sys", "半導體製程控制系統（產品營收 − SPTS − PCB；Derived）", [f"={c}5-{c}8-{c}9" for c in "BCDEFGHI"], "Derived", "產品營收扣兩事業全部營收（其服務部分未揭露，故系統略為低估、與服務合計不變）", "系統＋服務＋兩事業＝營收（恆等）")
qrow(12, "oth", "SPTS＋PCB", [f"={c}8+{c}9" for c in "BCDEFGHI"], "Derived", "—", "—")
qrow(13, "cost", "營業成本", [1175.7, 1207.286, H1(3842.0, 1327.7), 1327.7, 1413.108, 4751.867, 5255.060, 1577.791], "Verified", "同上", "—")
qrow(14, "gp", "毛利", [f"={c}4-{c}13" for c in "BCDEFGHI"], "Derived", "—", "stockanalysis 年度毛利：FY2025 7,404、FY2026 8,324")
qrow(15, "gmr", "毛利率（GAAP）", [f"={c}14/{c}4" for c in "BCDEFGHI"], "Derived", "—", "公司非 GAAP：FQ4-26 62.4%", "0.0%")
qrow(16, "rd", "研發費用", [338.0, 352.989, H1(1133.1, 388.8), 388.8, 399.023, 1360.334, 1532.118, 395.096], "Verified", "同上", "—")
qrow(17, "sga", "銷管費用", [248.9, 262.706, H1(840.0, 291.1), 291.1, 291.477, 1029.734, 1131.518, 301.458], "Verified", "同上", "法說稱 FQ4-26 SG&A 283（非 GAAP）")
qrow(18, "imp", "商譽與無形資產減損", [0, 0, 0, 0, 0, 239.1, 0, 0], "Verified", "FY2025 減損 239.1（FQ3-26 新聞稿九個月欄、FQ4-26 新聞稿）", "—")
qrow(19, "opx", "營業費用（研發＋銷管；不含減損）", [f"={c}16+{c}17" for c in "BCDEFGHI"], "Derived", "—", "—")
qrow(20, "oi", "營業利益（毛利 − 營業費用 − 減損）", [f"={c}14-{c}19-{c}18" for c in "BCDEFGHI"], "Derived", "—", "stockanalysis 年度營業利益：FY2026 5,661（相符）；FY2025 5,014（不含減損）")
qrow(21, "nonop", "營業外淨額（其他收入 − 利息）", [None, None, None, None, None, "=171.487-302.166", "=229.585-284.44", "=30.339-67.927"], "Verified", S_Q426 + "：利息 FY25 302.2／FY26 284.4；其他收入 171.5／229.6", "—")
qrow(22, "taxv", "所得稅", [None, None, None, None, None, 582.805, 775.154, 275.494], "Verified", "同上", "—")
qrow(23, "taxr", "有效稅率", [None, None, None, None, None, "=G22/(G20+G21)", "=H22/(H20+H21)", "=I22/(I20+I21)"], "Derived", "—", "—", "0.0%")
qrow(24, "ni", "淨利（GAAP）", [1088.4, 1202.849, H1(3467.7, 1201.0), 1201.0, 1363.059, 4061.643, 4830.771, 1548.588], "Verified", "同上", "stockanalysis：FY2025 4,062、FY2026 4,831")
qrow(25, "ck_ni", "檢查：稅前 − 稅 − 淨利（年度，應約 0）", [None, None, None, None, None, "=G20+G21-G22-G24", "=H20+H21-H22-H24", "=I20+I21-I22-I24"], "—", "—", "—", "0.0")
qrow(26, "eps", "EPS（GAAP 稀釋，分割調整後）", [0.816, 0.91, "=(26.26-9.12)/10", 0.912, 1.04, 3.04, 3.66, 1.18], "Verified／Derived", "FQ4-26 新聞稿（分割後）：FQ4-25 0.91、FY2025 3.04、FY2026 3.66、FQ4-26 1.04；FQ3-26 新聞稿（分割前 ÷10）：FQ3-25 8.16、FQ3-26 9.12、九個月 26.26", "SEC XBRL（10-K）：FY2025 3.04、FY2026 3.66；stockanalysis 同", "0.000")
qrow(27, "sh", "稀釋加權股數（百萬股，分割調整後）", [1333.0, 1327.341, None, 1317.5, 1314.986, 1337.502, 1319.633, 1311.064], "Verified", S_Q426 + "；FQ3：分割前 133.30／131.75 ×10", "—", "#,##0.0")
qrow(28, "ck", "檢查：FY2026 − (上半＋FQ3＋FQ4)（營收、產品、服務、淨利依序 B–E 欄）", ["=H4-D4-E4-F4", "=H5-D5-E5-F5", "=H6-D6-E6-F6", "=H24-D24-E24-F24"], "—", "九個月與單季皆取新聞稿四捨五入值，差額為 0.1 以內的捨入差", "—", "0.00")
for nm in ("rev", "prod", "svc", "spc", "spts", "pcb", "sys", "oth", "gp", "rd", "sga", "imp", "opx", "oi", "nonop", "taxv", "ni", "eps", "sh"):
    defname(f"a25_{nm}", f"Q!$G${QR[nm]}"); defname(f"a26_{nm}", f"Q!$H${QR[nm]}")
Q.cell(30, 1, "FQ2-27 公司指引（2026-10-28）").font = BOLD
hdr(Q, 31, ["項目", "指引", "低", "高", "", "", "", "", "", "標記", "來源1（公司）", "來源2／說明"])
for i, (nm, k, f) in enumerate((("營收（US$M）", "g_q2_rev", "#,##0"), ("GAAP 毛利率", "g_q2_gm", "0.0%"), ("GAAP 稀釋 EPS（US$）", "g_q2_eps", "0.00")), 32):
    Q.cell(i, 1, nm); Q.cell(i, 2, f"={k}").number_format = f; Q.cell(i, 3, f"={k}_lo").number_format = f; Q.cell(i, 4, f"={k}_hi").number_format = f
    Q.cell(i, 10, "Interested-party"); Q.cell(i, 11, "KLA FQ1-27 新聞稿（SEC 8-K，2026-10-28）"); Q.cell(i, 12, "非 GAAP：毛利率 62.5%±1.0%、EPS 1.44±0.11")
Q.column_dimensions["A"].width = 56
for c in "BCDEFGHI": Q.column_dimensions[c].width = 15
Q.column_dimensions["J"].width = 18; Q.column_dimensions["K"].width = 80; Q.column_dimensions["L"].width = 60
Q.freeze_panes = "B4"

# ---------------- Units：設備市場（設備模型成品）與 KLA 日曆年實績 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 設備市場（十億美元，日曆年；讀 markets/equipment 成品三情境）與 KLA 日曆年實績（百萬美元；由 Q 頁各季加總）"
hdr(UN, 3, ["項目", "單位", "CY2025", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ccy = lambda y: chr(ord("C") + y - 2025)   # C＝2025 … G＝2029
UR = {}; ur = 4
NM = {"L1": "先進邏輯〔AI〕", "L2": "其他代工與邏輯", "L3": "其他前段（類比、功率、光罩等）", "M1": "DRAM（含 HBM）〔AI〕", "M2": "NAND", "P": "先進封裝〔AI〕", "T": "設備總額"}
for s in SC:
    UN.cell(ur, 1, f"情境 {s}：設備市場（markets/equipment 5{s} 區塊）").font = BOLD; ur += 1
    for k in ("L1", "L2", "L3", "M1", "M2", "P", "T"):
        if k == "L3":
            pass
        vals, src = EQ[(s, k)]
        UN.cell(ur, 1, "  " + NM[k]); UN.cell(ur, 2, "US$B")
        for j, v in enumerate(vals):
            c = UN.cell(ur, 3 + j, v); c.font = BLUE; c.number_format = "0.00"
        UN.cell(ur, 8, src); UR[(s, k)] = ur; ur += 1
        if k == "L3":
            UN.cell(ur, 1, "  代工邏輯 WFE 小計（L）"); UN.cell(ur, 2, "US$B")
            for y in range(2025, 2030): UN[f"{ccy(y)}{ur}"] = f"={ccy(y)}{ur-3}+{ccy(y)}{ur-2}+{ccy(y)}{ur-1}"; UN[f"{ccy(y)}{ur}"].number_format = "0.00"
            UR[(s, "L")] = ur; ur += 1
        if k == "M2":
            UN.cell(ur, 1, "  記憶體 WFE 小計（M）"); UN.cell(ur, 2, "US$B")
            for y in range(2025, 2030): UN[f"{ccy(y)}{ur}"] = f"={ccy(y)}{ur-2}+{ccy(y)}{ur-1}"; UN[f"{ccy(y)}{ur}"].number_format = "0.00"
            UR[(s, "M")] = ur; ur += 1
    ur += 1
# m_mem 公式引用 Units!D8（A 的 L，2026）與 D11（A 的 M，2026）——確認列號
assert UR[("A", "L")] == 8 and UR[("A", "M")] == 11, (UR[("A", "L")], UR[("A", "M")])
UN.cell(ur, 1, "SEMI 晶圓前段設備（WFE；對帳用）"); UN.cell(ur, 2, "US$B")
for j, v in enumerate(SEMI_WFE[0]): c = UN.cell(ur, 3 + j, v); c.font = BLUE; c.number_format = "0.00"
UN.cell(ur, 8, SEMI_WFE[1] + "（SEMI 2026-07；設備商協會，傾向樂觀）"); UR["semi"] = ur; ur += 2
UN.cell(ur, 1, "KLA 日曆年實績（百萬美元；CY2025＝FQ3-25＋FQ4-25＋FY26 上半；CY2026 上半＝FQ3-26＋FQ4-26）").font = BOLD; ur += 1
hdr(UN, ur, ["項目", "單位", "CY2025", "CY2026 上半", "", "", "", "說明"]); ur += 1
for nm, lab in (("rev", "營收"), ("prod", "產品"), ("svc", "服務"), ("spc", "半導體製程控制（事業別，含服務）"), ("oth", "SPTS＋PCB"), ("sys", "半導體製程控制系統（產品 − SPTS − PCB）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, "US$M")
    UN[f"C{ur}"] = f"=Q!B{QR[nm]}+Q!C{QR[nm]}+Q!D{QR[nm]}"; UN[f"D{ur}"] = f"=Q!E{QR[nm]}+Q!F{QR[nm]}"
    for c in "CD": UN[f"{c}{ur}"].number_format = "#,##0.0"
    defname(f"cy25_{nm}", f"Units!$C${ur}"); defname(f"h1_{nm}", f"Units!$D${ur}"); ur += 1
UN.cell(ur, 1, "CY2025 半導體製程控制營收 ÷ 公司 WFE（s25）"); UN.cell(ur, 2, "%"); UN[f"C{ur}"] = "=cy25_spc/(kla_wfe25*1000)"; UN[f"C{ur}"].number_format = "0.00%"
UN.cell(ur, 8, "pc_u（強度年增）的基期；公司稱 2030 年較 2025 提升 150bp"); defname("s25", f"Units!$C${ur}"); ur += 2
UN.cell(ur, 1, "KLA 會計年度與日曆年：FY(t)＝CY(t−1) 下半＋CY(t) 上半。CY2027 起每年上下半各 1/2（Assumed，忽略年內斜率）；CY2026 的下半＝錨點 − 上半實績。")
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10; UN.column_dimensions["H"].width = 80
for c in "CDEFG": UN.column_dimensions[c].width = 12

# ---------------- 引擎（PL 與 Sens 共用） ----------------
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # FY2025＝C … FY2029＝G
DEFP = dict(m="m_mem", u="pc_u", kap="k_ap", ap26="ap26", h2="h2up", cut="chn_cut", gmi="gm_inc", tax="tax", bb="bb", oexp="oth_exp",
            gsh2="g_svc_h2", go26="g_oth26", opq="opex27", scn="A", gsvc=None, goth=None, gopx=None, wf=None)   # v0.13：Wattonomics 錨定案例用；預設不變（wf＝年→AI 分項係數儲存格）
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    gsvc = P["gsvc"] or f"g_svc_{s}"; goth = P["goth"] or f"g_oth_{s}"; gopx = P["gopx"] or f"g_opex_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2025／FY2025A", "CY2026／FY2026A", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "說明"]); r += 1
    ws.cell(r, 1, "A. 日曆年（C＝CY2025 … G＝CY2029）").font = BOLD; r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    CYS = range(2025, 2030); c = ccy
    row("W", "加權前段 WFE（代工邏輯＋m × 記憶體）", "US$M", {c(y): f"=(Units!{c(y)}{UR[(s,'L')]}+{P['m']}*Units!{c(y)}{UR[(s,'M')]})*1000" + (f"-(Units!{c(y)}{UR[(s,'L1')]}+{P['m']}*Units!{c(y)}{UR[(s,'M1')]})*(1-{P['wf'][y]})*1000" if P['wf'] and y in P['wf'] else "") for y in CYS}, note=f"設備模型情境 {s}；m＝記憶體相對強度")
    row("Peq", "先進封裝設備市場", "US$M", {c(y): f"=Units!{c(y)}{UR[(s,'P')]}*1000" + (f"*{P['wf'][y]}" if P['wf'] and y in P['wf'] else "") for y in CYS})
    row("tot", "營收 CY2026 錨點（上半實績 ×（2＋下半增幅））", "US$M", {"D": f"=h1_rev*(2+{P['h2']})"}, note="上半實績＋公司說法（下半約高 20%）")
    row("svc", "服務 CY2026 錨點", "US$M", {"D": f"=h1_svc*(2+{P['gsh2']})"})
    row("oth", "SPTS＋PCB CY2026 錨點", "US$M", {"D": f"=cy25_oth*(1+{P['go26']})"})
    rs = r + 3  # 系統列位置（ap、fe、r 之後）
    row("ap", "先進封裝製程控制系統營收", "US$M", {"C": f"={P['ap26']}/(1+ap_g26)", "D": f"={P['ap26']}", **{c(y): f"={c(y-1)}{r}*(1+{P['kap']}*({c(y)}{R['Peq']}/{c(y-1)}{R['Peq']}-1))" for y in range(2027, 2030)}}, note="CY2027 起：前一年 ×（1＋k × 封裝設備市場成長）")
    row("fe", "前段製程控制系統營收", "US$M", {"C": f"=C{rs}-C{r-1}", "D": f"=D{rs}-D{r-1}", **{c(y): f"={c(y)}{r+1}*{c(y)}{R['W']}" for y in range(2027, 2030)}}, note="CY2025／2026＝系統 − 先進封裝；之後＝強度 × 加權 WFE")
    row("ri", "製程控制強度（前段系統 ÷ 加權前段 WFE）", "%", {"C": f"=C{r-1}/C{R['W']}", "D": f"=D{r-1}/D{R['W']}", **{c(y): f"=$D{r}*(1+{P['u']})^{y-2026}" for y in range(2027, 2030)}}, fmt="0.00%", note="CY2026 校準；之後每年 ×（1＋pc_u）")
    assert r == rs
    row("sys", "半導體製程控制系統營收 CY", "US$M", {"C": "=cy25_sys", "D": f"=D{R['tot']}-D{R['svc']}-D{R['oth']}", **{c(y): f"=({c(y)}{R['fe']}+{c(y)}{R['ap']})*(1-chn_sh26*{P['cut']})" for y in range(2027, 2030)}}, note="CY2026＝錨點 − 服務 − 兩事業；CY2027 起 ×（1 − 中國占比 × 中國下降比例）")
    row("pcs", "顯示：製程控制占前段 WFE（強度 ÷ KLA 市占）", "%", {c(y): f"={c(y)}{R['ri']}/kla_sh_pc" for y in CYS}, fmt="0.0%", note="KLA 市占 58% 為賣方數字，只供顯示")
    ws.cell(r, 1, "B. 會計年度（C＝FY2025A … G＝FY2029；FY(t)＝CY(t−1) 下半＋CY(t) 上半）").font = BOLD; r += 1
    fy = FYC
    row("fsys", "半導體製程控制系統營收 FY", "US$M", {"C": "=a25_sys", "D": "=a26_sys", "E": f"=(D{R['sys']}-h1_sys)+0.5*E{R['sys']}", "F": f"=0.5*E{R['sys']}+0.5*F{R['sys']}", "G": f"=0.5*F{R['sys']}+0.5*G{R['sys']}"}, note="FY2027＝CY2026 下半（錨點 − 上半實績）＋1/2 CY2027")
    row("fsvc", "服務營收", "US$M", {"C": "=a25_svc", "D": "=a26_svc", **{fy[y]: f"={fy[y-1]}{r}*(1+{gsvc})" for y in range(2027, 2030)}})
    row("foth", "SPTS＋PCB 營收", "US$M", {"C": "=a25_oth", "D": "=a26_oth", **{fy[y]: f"={fy[y-1]}{r}*(1+{goth})" for y in range(2027, 2030)}})
    row("rev", "營收", "US$M", {fy[y]: f"={fy[y]}{R['fsys']}+{fy[y]}{R['fsvc']}+{fy[y]}{R['foth']}" for y in range(2025, 2030)}, note="FY2025、FY2026＝實績（恆等）")
    row("gp", "毛利", "US$M", {"C": "=a25_gp", "D": "=a26_gp", **{fy[y]: f"=a26_gp+{P['gmi']}*({fy[y]}{R['rev']}-a26_rev)" for y in range(2027, 2030)}}, note="FY2026 毛利＋增量毛利率 ×（營收 − FY2026 營收）")
    row("gm", "毛利率", "%", {fy[y]: f"={fy[y]}{R['gp']}/{fy[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("opx", "營業費用（研發＋銷管）", "US$M", {"C": "=a25_opx", "D": "=a26_opx", "E": f"={P['opq']}", **{fy[y]: f"={fy[y-1]}{r}*(1+{gopx})" for y in range(2028, 2030)}})
    row("oi", "營業利益", "US$M", {"C": "=a25_oi", **{fy[y]: f"={fy[y]}{R['gp']}-{fy[y]}{R['opx']}" for y in range(2026, 2030)}}, note="FY2025 含減損 239.1")
    row("oim", "營業利益率", "%", {fy[y]: f"={fy[y]}{R['oi']}/{fy[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("noi", "營業外淨額", "US$M", {"C": "=a25_nonop", "D": "=a26_nonop", **{fy[y]: f"={P['oexp']}" for y in range(2027, 2030)}})
    row("ni", "淨利", "US$M", {"C": "=a25_ni", "D": "=a26_ni", **{fy[y]: f"=({fy[y]}{R['oi']}+{fy[y]}{R['noi']})*(1-{P['tax']})" for y in range(2027, 2030)}})
    row("sh", "稀釋股數", "百萬股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=shares27", **{fy[y]: f"={fy[y-1]}{r}*(1-{P['bb']})" for y in range(2028, 2030)}})
    row("eps", "EPS（GAAP 稀釋）", "US$", {"C": "=a25_eps", "D": "=a26_eps", **{fy[y]: f"={fy[y]}{R['ni']}/{fy[y]}{R['sh']}" for y in range(2027, 2030)}}, fmt="0.00")
    row("chn", "顯示：中國營收（× FY2026 中國占比；CY2027 起扣下降）", "US$M", {"D": "=a26_rev*chn_sh26", **{fy[y]: f"={fy[y]}{R['rev']}*chn_sh26*(1-{P['cut']}*IF({y}=2027,0.5,1))" for y in range(2027, 2030)}}, fmt="#,##0", note="只供顯示")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — KLA 損益（US$M；EPS 為美元）。FY2025、FY2026 為實績，FY2027–FY2029 為模型。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 45
TITLE = {"A": "情境 A：設備市場續增（2029 見頂）", "B": "情境 B：設備市場 2028 持平、2029 −5%", "C": "情境 C：設備市場 2028 −20%（2027 見頂）"}
for s in SC:
    PB[s] = engine(PL, r0, TITLE[s], scn=s); r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; PL.cell(sr, 8, TITLE[s]); sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$M", "#,##0"), ("fsys", "  半導體製程控制系統", "US$M", "#,##0"), ("fsvc", "  服務", "US$M", "#,##0"), ("foth", "  SPTS＋PCB", "US$M", "#,##0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "US$M", "#,##0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "US$M", "#,##0"), ("eps", "EPS（GAAP 稀釋）", "US$", "0.00")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：FY2025、FY2026 模型營收 − 實績（應為 0）").font = BOLD; sr += 1
for s in SC:
    PL.cell(sr, 1, f"  情境 {s}"); PL[f"C{sr}"] = f"=ABS(C{SUMR[(s,'rev')]}-a25_rev)+ABS(D{SUMR[(s,'rev')]}-a26_rev)"; PL[f"C{sr}"].number_format = "0.000"; sr += 1
PL.column_dimensions["A"].width = 54; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 60
for c in "CDEFG": PL.column_dimensions[c].width = 13

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（第 4 節）。EPS 一律 GAAP 稀釋"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」GAAP 稀釋 EPS（年底所在會計年度，例如 2025-12 對 FY2026＝2026-06 期）；另列「前期」（上一個已結束會計年度）。價格與 EPS 皆為 2026-06 的 10 比 1 分割調整後"
# 0. 目標價（列號固定，引用第 5 節）
V.cell(4, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, 5, ["項目", "數值", "說明"])
TP0 = 6
# 第 1 節以後
V.cell(16, 1, "1. 公司公布 EPS（會計年度，GAAP 稀釋，美元，分割調整後）— 兩個來源").font = BOLD
hdr(V, 17, ["會計年度（6 月結束）", "來源1：SEC XBRL（10-K；分割前 ÷10）", "來源2：macrotrends（6 月 TTM，÷10）", "兩者差距", "標記", "說明"])
EPS1 = {2016: 0.449, 2017: 0.588, 2018: 0.510, 2019: 0.749, 2020: 0.770, 2021: 1.337, 2022: 2.192, 2023: 2.415, 2024: 2.03, 2025: 3.04, 2026: 3.66}
EPS2 = {2016: 0.449, 2017: 0.588, 2018: 0.510, 2019: 0.749, 2020: 0.770, 2021: 1.337, 2022: 2.192, 2023: 2.415, 2024: 2.028, 2025: 3.037, 2026: 3.66}
NOTE = {2018: "美國稅改一次性費用使 FY2018 偏低", 2024: "SEC 為 FY2026 10-K 重編的分割後 2.03（分割前 20.28）", 2025: "分割前 30.37", 2026: "來源2 為 stockanalysis（macrotrends 2026-06 列異常，未用）；Q 頁"}
EP = {}
for i, y in enumerate(range(2016, 2027), 18):
    V.cell(i, 1, f"FY{y}（{y}-06）")
    V.cell(i, 2, EPS1[y]).font = BLUE; V.cell(i, 3, EPS2[y]).font = BLUE
    V.cell(i, 4, f"=B{i}/C{i}-1").number_format = "0.0%"
    for c in "BC": V[f"{c}{i}"].number_format = "0.000"
    V.cell(i, 5, "Verified"); V.cell(i, 6, NOTE.get(y, "SEC companyconcept EarningsPerShareDiluted（10-K 年度值）")); EP[y] = i
V.cell(29, 1, "FY2026 EPS 改由 Q 頁實績帶入第 2 節（＝a26_eps）；其餘年度用來源1。兩來源差距皆來自四捨五入").font = BOLD
V.cell(30, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 31, ["年底", "收盤價 US$（分割調整）", "當期 FY EPS", "前期 FY EPS", "P/E 當期（自算）", "P/E 前期（自算）", "macrotrends TTM P/E", "自算當期 − macrotrends", "收盤價來源2：macrotrends（÷10，含股利調整）", "標記", "來源／說明"])
PX = {2016: 78.68, 2017: 105.07, 2018: 89.49, 2019: 178.17, 2020: 258.91, 2021: 430.11, 2022: 377.03, 2023: 581.30, 2024: 630.12, 2025: 1215.08}
MT_PX = {2016: 77.51, 2017: 103.76, 2018: 88.61, 2019: 176.85, 2020: 257.45, 2021: 428.18, 2022: 375.83, 2023: 580.14, 2024: 629.42, 2025: 1214.77}
MT_PE = {2016: 14.09, 2017: 25.00, 2018: 9.69, 2019: 24.94, 2020: 29.36, 2021: 22.06, 2022: 15.44, 2023: 29.37, 2024: 26.49, 2025: 35.36}
for i, y in enumerate(range(2016, 2026), 32):
    V.cell(i, 1, f"{y}-12")
    V.cell(i, 2, f"={PX[y]}/10").font = BLUE
    V.cell(i, 3, "=a26_eps" if y == 2025 else f"=B{EP[y+1]}"); V.cell(i, 4, f"=B{EP[y]}")
    V.cell(i, 5, f"=B{i}/C{i}"); V.cell(i, 6, f"=B{i}/D{i}")
    V.cell(i, 7, MT_PE[y]).font = BLUE; V.cell(i, 8, f"=E{i}-G{i}"); V.cell(i, 9, f"={MT_PX[y]}/10").font = BLUE
    for c in "BCDI": V[f"{c}{i}"].number_format = "#,##0.000"
    for c in "EFGH": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, {2018: "當期 P/E 11.9 為十年最低（2018Q4 記憶體循環下修）", 2025: "當期 EPS＝FY2026 實績 3.66；最高 33.2"}.get(y, "收盤價：digrin「Real price」（未調整）12 月收盤 ÷10"))
V.cell(43, 1, "收盤價來源1：digrin（Real price，年底收盤，未做股利調整）；來源2：macrotrends（含股利調整，故略低，I 欄）。macrotrends P/E 為 TTM 口徑（年底已公布的近四季 EPS），與自算的當期口徑不同，只作對照。")
V.cell(45, 1, "3. 分位").font = BOLD
hdr(V, 46, ["統計", "P/E 當期（自算）", "P/E 前期（自算）", "macrotrends P/E（參考）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 47):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E32:E41,{p})"); V.cell(i, 3, f"=PERCENTILE(F32:F41,{p})"); V.cell(i, 4, f"=PERCENTILE(G32:G41,{p})")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(50, 1, "年數"); V["B50"] = "=COUNT(E32:E41)"; V["C50"] = "=COUNT(F32:F41)"; V["D50"] = "=COUNT(G32:G41)"
defname("pe_peak", "Val!$B$47"); defname("pe_mid", "Val!$B$48"); defname("pe_trough", "Val!$B$49")
V.cell(51, 1, "偏誤方向：2016–2018 年（AI 需求出現前）當期 P/E 13.4、20.6、11.9 偏低，拉低 25 分位與中位數；2023、2025 年（28.7、33.2）為 AI 期高倍數。照規則計算，不另換倍數。")
# 4. 循環位置
V.cell(53, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(54, 1, CP.rule_text(PEAK) + "。KLA 的 FY(t)＝t−1 年 7 月～t 年 6 月，上下半各一半；以 FY(t) 對照日曆年 t（結束年）。驗證：以 FY 加權的設備總額（1/2 CY(t−1)＋1/2 CY(t)）的見頂會計年度與此判定一致（右表）")
hdr(V, 55, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "FY 加權設備總額 FY2027", "FY 加權設備總額 FY2028", "FY 加權見頂（FY2027–29）"])
for i, s in enumerate(SC, 56):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    t = UR[(s, "T")]
    V.cell(i, 6, f"=0.5*Units!D{t}+0.5*Units!E{t}").number_format = "0.0"; V.cell(i, 7, f"=0.5*Units!E{t}+0.5*Units!F{t}").number_format = "0.0"
    V.cell(i, 8, f'=IF(AND(F{i}>=G{i},F{i}>=0.5*Units!F{t}+0.5*Units!G{t}),"FY2027",IF(G{i}>=0.5*Units!F{t}+0.5*Units!G{t},"FY2028","FY2029 或之後"))')
# 5. 隱含股價
V.cell(60, 1, "5. 隱含股價（美元）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 61, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 62; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{56 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = "0.00"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{56 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0.0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for cc in range(1, 9): V.cell(vr, cc).fill = KEY
            VAL[s] = vr
        vr += 1
assert vr == 83
V.cell(84, 1, "6. 目前遠期 P/E").font = BOLD
hdr(V, 85, ["項目", "A", "B", "C", "備註"])
fr = 86
for nm, f, fm, note in (("現價（stockanalysis 2026-10-23 收盤）", ["=px"], "0.00", "2026-10 月內由 206.89（10-02）跌至 145.61"),
                        ("遠期 P/E：現價 ÷ 共識 FY2027 EPS（非 GAAP）", ["=px/cons_eps27"], "0.0", "stockanalysis 顯示 26.65x"),
                        ("現價 ÷ FY2026 EPS（實績）", ["=px/a26_eps"], "0.0", ""),
                        ("現價 ÷ FY2027 EPS（模型）", [f"=px/PL!E{SUMR[(s,'eps')]}" for s in SC], "0.0", "A／B／C"),
                        ("現價 ÷ FY2028 EPS（模型）", [f"=px/PL!F{SUMR[(s,'eps')]}" for s in SC], "0.0", "A／B／C"),
                        ("現價隱含倍數 ÷ 規則倍數（情境 A）— 以 FY2027 A EPS", [f"=B{fr+3}/E56"], "0.00", ">1＝現價高於規則給的倍數")):
    V.cell(fr, 1, nm)
    for j, ff in enumerate(f): V.cell(fr, 2 + j, ff).number_format = fm
    V.cell(fr, 5, note); fr += 1
# 0. 目標價內容
TPR = [("目標價（12 個月，情境加權）", f"=tp_wA*B{TP0+2}+tp_wB*B{TP0+3}+tp_wC*B{TP0+4}", "0.00", "權重 × 各情境隱含價"),
       ("上檔空間（目標價 ÷ 現價 − 1）", f"=B{TP0}/px-1", "0.0%", "現價見第 6 節"),
       ("情境 A 隱含價（樂觀）", f"=F{VAL['A']}", "0.00", "第 5 節規則值（設備市場續增）"),
       ("情境 B 隱含價（基準）", f"=F{VAL['B']}", "0.00", "第 5 節規則值（設備市場 2028 持平）"),
       ("情境 C 隱含價（下檔）", f"=F{VAL['C']}", "0.00", "第 5 節規則值（設備市場 2028 −20%）"),
       ("情境權重 A／B／C", "=tp_wA", "0%", None),
       ("目標價方法", "P/E 規則值：FY2027（2027-06 期）GAAP 稀釋 EPS × 2016–2025 年底當期 P/E 依循環位置取分位（A、B 中段＝中位數，C 見頂＝25 分位；設備股共用規則）。十年倍數含 AI 前（2016–2018）低倍數期，偏誤方向為偏低（壓低目標價）；照規則計算，未另換倍數", None, None),
       ("檢查：權重合計（應為 100%）", "=tp_wA+tp_wB+tp_wC", "0%", None)]
for i, (lab, f, fm, note) in enumerate(TPR, TP0):
    V.cell(i, 1, lab); V.cell(i, 2, f)
    if fm: V.cell(i, 2).number_format = fm
    if note: V.cell(i, 3, note)
    if i == TP0:
        for cc in (1, 2, 3): V.cell(i, cc).fill = KEY
V.cell(TP0 + 5, 3, "=tp_wB").number_format = "0%"; V.cell(TP0 + 5, 4, "=tp_wC").number_format = "0%"
V.column_dimensions["A"].width = 50
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["K"].width = 70

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 US$M", "FY2027 EPS US$", "相對基準", "FY2028 營收 US$M", "FY2028 EPS US$", "相對基準"])
cases = [("中國系統營收 −25%（CY2027 起）", "出口管制擴大（Assumed）", dict(cut="chn_cut_lo")), ("中國系統營收 −50%（CY2027 起）", "", dict(cut="chn_cut_hi")),
         ("製程控制強度不再提升", "pc_u＝0", dict(u="pc_u_lo")), ("製程控制強度提升加倍", "2030 較 2025 +300bp", dict(u="pc_u_hi")),
         ("FQ2-27 營收指引下緣 4.5B", "CY2026 錨點較低", dict(h2="h2up_lo")), ("FQ2-27 營收指引上緣 4.9B", "", dict(h2="h2up_hi")),
         ("記憶體相對強度 低（代工邏輯占 66%）", "", dict(m="m_mem_lo")), ("記憶體相對強度 高（代工邏輯占 60%）", "", dict(m="m_mem_hi")),
         ("先進封裝倍數 1.0x", "與封裝市場同速", dict(kap="k_ap_lo")), ("先進封裝倍數 2.5x", "", dict(kap="k_ap_hi")),
         ("先進封裝 CY2026 1.4B", "取 substack 會計年度口徑上緣", dict(ap26="ap26_hi")),
         ("增量毛利率 60%", "公司 60–65% 下緣", dict(gmi="gm_inc_lo")), ("增量毛利率 65%", "上緣", dict(gmi="gm_inc_hi")),
         ("營業費用 FY2028–29 年增 5%", "", dict(gopx="g_opex_A_lo")), ("營業費用 FY2028–29 年增 11%", "", dict(gopx="g_opex_A_hi")),
         ("服務年增 12%", "", dict(gsvc="g_svc_A_lo")), ("服務年增 16%", "", dict(gsvc="g_svc_A_hi")),
         ("稅率 13%", "", dict(tax="tax_lo")), ("稅率 17%", "", dict(tax="tax_hi")),
         ("營業外 −200", "", dict(oexp="oth_exp_lo")), ("股數不減", "買回停止", dict(bb="bb_lo")), ("股數年減 2.5%", "", dict(bb="bb_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}E{R['rev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"F{i}"] = f"={pre}F{R['rev']}"; SE[f"G{i}"] = f"={pre}F{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "0.00"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, 50 + n * 36, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "設備市場情境 B", "2028 持平、2029 −5%；服務、兩事業、費用成長 B", PB["B"], "PL!")
put_case(nn + 1, "設備市場情境 C", "2028 −20%；服務、兩事業、費用成長 C", PB["C"], "PL!")
SE.cell(nn + 3, 1, "讀法：FY2027 的一半是 CY2026 下半（錨點，三情境相同），另一半是 CY2027（三情境的設備市場相同），所以情境差異到 FY2028 才明顯；中國與強度假設自 CY2027 起作用，FY2027 只吃一半。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識、公司的 WFE 說法、SEMI。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L = chr(64 + cv)
        RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
eg = lambda key, col: (lambda s: f"=PL!{col}{PB[s][key]}")
RC.cell(rr_, 1, "A. CY2026 下半錨點 vs 公司指引＋共識").font = BOLD; rr_ += 1
rec("CY2026 下半營收（錨點 − 上半實績）vs FQ1-27 實績＋共識 FQ2-27（財報前）", "=Q!I4+cons_q2_rev", lambda s: f"=PL!D{PB[s]['tot']}-h1_rev", "公司 8-K（FQ1 4,136.0）＋Yahoo 共識（FQ2 4.49B，2026-10-25）", "錨點用公司 FQ2-27 指引 4.7B，高於財報前共識")
RC.cell(rr_, 1, "B. FY2027、FY2028 vs 共識（Yahoo Finance；共識 EPS 為非 GAAP，模型為 GAAP）").font = BOLD; rr_ += 1
rec("FY2027 營收", "=cons_rev27", pl("rev", 2027), "Yahoo（18.14B）；stockanalysis 18.1B")
rec("FY2027 EPS", "=cons_eps27", pl("eps", 2027), "Yahoo 5.44；stockanalysis 5.46", "GAAP 比非 GAAP 低約 2–3%（FY2026：3.66 vs 3.76）", fmt="0.00")
rec("FY2028 營收", "=cons_rev28", pl("rev", 2028), "Yahoo（21.57B）")
rec("FY2028 EPS", "=cons_eps28", pl("eps", 2028), "Yahoo 6.74", "", fmt="0.00")
RC.cell(rr_, 1, "C. WFE：設備模型 vs 公司說法、SEMI（US$B）").font = BOLD; rr_ += 1
rec("CY2025 WFE（前段＋先進封裝）vs 公司約 120", "=kla_wfe25", lambda s: f"=Units!C{UR[(s,'L')]}+Units!C{UR[(s,'M')]}+Units!C{UR[(s,'P')]}", CALL, "設備模型 2025 以 SEMI 實績校準", fmt="0.0")
rec("CY2026 WFE（前段＋先進封裝）vs 公司「low 150s」（取 152）", 152, lambda s: f"=Units!D{UR[(s,'L')]}+Units!D{UR[(s,'M')]}+Units!D{UR[(s,'P')]}", CALL, "設備模型 2026 比公司說法高；本模型以 KLA 自身的 CY2026 錨點校準強度，所以這個差異不進入營收", fmt="0.0")
RC.cell(rr_ - 1, 2).font = BLUE
rec("CY2026 前段 WFE vs SEMI WFE", f"=Units!D{UR['semi']}", lambda s: f"=Units!D{UR[(s,'L')]}+Units!D{UR[(s,'M')]}", "SEMI 2026-07（設備模型成品）", "", fmt="0.0")
rec("CY2027 前段 WFE vs SEMI WFE（年增連乘）", f"=Units!E{UR['semi']}", lambda s: f"=Units!E{UR[(s,'L')]}+Units!E{UR[(s,'M')]}", "同上", "公司引述 2027 共識約 190B（含封裝；未確認）", fmt="0.0")
RC.cell(rr_, 1, "D. 強度與市占（顯示；非獨立檢驗）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "製程控制強度 CY2025（前段系統 ÷ 加權前段 WFE）"); RC.cell(rr_, 3, f"=PL!C{PB['A']['ri']}").number_format = "0.00%"; RC.cell(rr_, 10, "CY2026 校準值見下一列；兩者差距反映設備模型 2026 WFE 高於公司說法"); rr_ += 1
RC.cell(rr_, 1, "製程控制強度 CY2026（校準）"); RC.cell(rr_, 3, f"=PL!D{PB['A']['ri']}").number_format = "0.00%"; rr_ += 1
RC.cell(rr_, 1, "製程控制占前段 WFE CY2026（強度 ÷ KLA 市占 58%）"); RC.cell(rr_, 3, f"=PL!D{PB['A']['pcs']}").number_format = "0.0%"; RC.cell(rr_, 10, "業界常見說法約 11–13%（未找到可引用來源）"); rr_ += 1
RC.cell(rr_, 1, "E. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FQ1-27 實績營收 ÷ 模型 FY2027（A）"); RC.cell(rr_, 3, f"=Q!I4/PL!E{SUMR[('A','rev')]}").number_format = "0.0%"; RC.cell(rr_, 10, "營收逐季上升，Q1 低於 25% 屬正常"); rr_ += 1
RC.cell(rr_, 1, "模型 FY2027 中國營收（A；× FY2026 占比 30%）"); RC.cell(rr_, 3, f"=PL!E{PB['A']['chn']}").number_format = "#,##0"; rr_ += 1
RC.column_dimensions["A"].width = 60
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 80

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（SEC 原文可開啟；法說數字取自逐字稿網站）"; SO["A1"].font = BOLD
srcs = [("KLA FQ4-26 與 FY2026 新聞稿（SEC 8-K Exhibit 99.1，2026-07-28）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000024/exhibit991earningsrelease7.htm"),
 ("KLA FQ3-26 新聞稿（SEC 8-K Exhibit 99.1，2026-04-29）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000014/exhibit991earningsrelease3.htm"),
 ("KLA FY2026 10-K（中國占比、服務占比、訂單、出口管制）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000027/klac-20260630.htm"),
 ("SEC XBRL companyconcept：EarningsPerShareDiluted（歷年 EPS）", "https://data.sec.gov/api/xbrl/companyconcept/CIK0000319201/us-gaap/EarningsPerShareDiluted.json"),
 ("KLA FQ4-26 法說逐字稿（MarketBeat）", "https://www.marketbeat.com/earnings/reports/2026-7-28-kla-co-stock/"),
 ("TIKR：KLA FQ4-26 摘要（2026-07-29）", "https://www.tikr.com/blog/kla-corporations-q4-earnings-hit-a-record-3-66-billion-the-guide-points-higher-still"),
 ("MIAX：KLAC 10 比 1 分割通知（除權 2026-06-12）", "https://www.miaxglobal.com/sites/default/files/alert-files/KLAC_Split_58941.pdf"),
 ("tipranks：財報日 2026-10-28（FQ1-27）、10 月收盤價", "https://www.tipranks.com/stocks/klac/earnings"),
 ("stockanalysis：股價歷史（2026-10-23 收盤 145.61）", "https://stockanalysis.com/stocks/klac/history/"),
 ("stockanalysis：年度財報（FY2022–26）與共識（2026-10-20）", "https://stockanalysis.com/stocks/klac/forecast/"),
 ("Yahoo Finance：共識（FY2027、FY2028、FQ1、FQ2）", "https://finance.yahoo.com/quote/KLAC/analysis/"),
 ("digrin：年底收盤價（Real price）", "https://www.digrin.com/stocks/detail/KLAC/price"),
 ("macrotrends：P/E 與 TTM EPS 歷史", "https://www.macrotrends.net/stocks/charts/KLAC/kla/pe-ratio"),
 ("handsoff substack（2026-05-31；引 BofA）：製程控制市占 58%、FQ3-26 中國 24%", "https://handsoff.substack.com/p/kla-beat-across-the-board-process"),
 ("市場層：markets/equipment 成品（本 repo）", rel(EQ_F)),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 120

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("KLA（KLAC）收支與估值模型 v0.2（AI 半導體 Project 步驟 5：設備與材料——量測與檢測）　2026-10-25", True),
 ("模型命題：KLA 的系統營收跟著晶圓廠設備支出（WFE）走，再乘上製程控制強度——先進製程與先進封裝每片晶圓需要更多量測與檢測。股價取決於設備市場在 2027 年之後放緩時，強度上升與服務能否撐住盈餘，以及中國（FY2026 占營收 30%）的出口管制風險。", False), ("", False),
 ("會計年度：7 月到隔年 6 月（FY2027＝2026-07～2027-06，一半在 CY2026、一半在 CY2027）。FY(t)＝CY(t−1) 下半＋CY(t) 上半。金額百萬美元；EPS 為 GAAP 稀釋、2026-06-12 除權的 10 比 1 分割調整後。FY2027 第一季預定 2026-10-28 公布，建置日（2026-10-25）尚未公布，故不納入。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("系統營收（日曆年）＝製程控制強度 × 加權前段 WFE（設備模型三情境；代工邏輯＋m × 記憶體）＋先進封裝製程控制（封裝設備成長 × 2 倍）；強度以 CY2026 錨點（上半實績＋公司「下半約高 20%」）校準，之後每年 +pc_u（公司 2030 模型：占 WFE 比重 +150bp）", False),
 ("服務、SPTS＋PCB：前一年 × 成長率（Assumed，錨在公司長期目標）。毛利＝FY2026 毛利＋增量毛利率（兩點推導 64.6%，公司說 60–65%）× 營收增量；營業費用 FY2027 取公司說法，之後 × 成長率；不寫死利益率", False),
 ("情境 A／B／C：設備市場模型的三情境（A 續增、B 2028 持平、C 2028 −20%），服務、兩事業與費用成長隨情境；中國出口管制只在 Sens", False),
 ("Val：P/E 規則——10 年（2016–2025 年底）當期 P/E 分位，見頂年 25 分位、中段中位數、谷底 75 分位；循環位置依設備股共用規則（各情境見頂年度讀設備市場模型：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。第 0 節為情境加權目標價。本版未做先驗，無 Andy 合理 P/E 對照", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("已知限制：(1) 製程控制強度只有一個校準點（CY2026 錨點），強度上升率來自公司 2030 模型（Interested-party）；(2) 記憶體相對強度由公司「代工邏輯 low-to-mid 60s」一句推得；(3) 系統營收＝產品 − 兩事業（兩事業的服務未揭露）；(4) CY2027 起上下半各 1/2，忽略年內斜率；(5) 中國不在基準內另扣（設備模型 WFE 已含），出口管制只在 Sens；(6) 共識 EPS 為非 GAAP，模型為 GAAP；(7) 2026-10 股價由約 207 跌到 146，原因未查明。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('應用材料（Applied Materials）', 'AMAT', 21.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('科林研發（Lam Research）', 'LRCX', 22.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('ASML（美國 ADR）', 'ASML', 24.74, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Lasertec', '6920.T', 34.02, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2028: "8/12", 2029: "4/12"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 8／12 × FY2028（2027-07～2028-06）＋4／12 × FY2029（會計年度 7–6 月，依月份調整）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='US$',
    unit_rev='US$M',
    unit_ni='US$M',
    fmt='0.00',
    fmt_rev='#,##0',
    band_fmt='0.00',
    roe=('FY2024–FY2026', [0.8784, 1.0077, 0.875], 'stockanalysis 比率頁（2026-10-25 讀取）：FY2024 87.84%、FY2025 100.77%、FY2026 87.50%', '單一來源；高 ROE 來自大量回購使股東權益偏低'),
    nim=('FY2016–FY2025，GAAP', [(2.98, 0.70442), (3.48, 0.92608), (4.04, 0.80226), (4.57, 1.18), (5.81, 1.22), (6.92, 2.08), (9.21, 3.32), (10.5, 3.39), (9.81, 2.76), (12.16, 4.06)], 'vcpscanner.com 年度損益（US$B）', '第二來源：FY2025 與 Q 頁實績核對；其餘年度單一來源（FY2026 淨利率 35.6%，未納入）'),
    cons_tp=(232.38, 'stockanalysis：30 位分析師平均 US$232.38（中位數 228.50；區間 175–325），資料更新 2026-10-20（10 比 1 分割後）', '第二來源：未找到'),
    cons_feps=('=cons_eps28', '共識前瞻 EPS＝共識 FY2028 非 GAAP EPS（In 頁 cons_eps28，Yahoo 6.74）；本模型前瞻 EPS 為 8／12 FY2028＋4／12 FY2029（GAAP），約晚四個月', 'Derived'),
    peer_note='前段設備與量測檢測同業：應用材料、科林研發、ASML、東京威力科創、Lasertec',
    dcf_note='；會計年度 7–6 月，FY2027＝2026-07～2027-06 以 t＝1 折現（略低估）',
    pe10='E32:E41',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（當期口徑自算）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF（v0.4 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（股權報酬視為費用），以 WACC 折現到 2026-12-31（期中慣例；FY2027 只計評價日後 6 個月），FY2027–32 明確期（FY2030 起收斂法），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金與投資 ÷ 稀釋股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁、tools/fcff_dcf.py）；β 為自算產業 β（美國半導體設備同業中位數，Blume 調整）重新加槓桿（股權成本與 WACC 見 0.1 節）。Comparable：應用材料、科林研發、ASML、東京威力科創、Lasertec 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：科磊製程控制份額高、歷史上享有溢價，用同業中位數偏低估。',
    readme=['v0.2（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。'])

# ---------------- 估值 v2（2026-10-25，Andy 11:18／11:20／11:28／11:53）：PL 延伸到 FY2032（收斂法）＋ Val「0. 目標價」（DCF 與 Comparable 加權） ----------------
# 為不位移既有儲存格（FY2026–FY2029 既有 EPS 逐格不變），新內容放在各頁既有內容之後；HTML 估值章第一個區塊讀「0.3 結論」
def valuation_v2(CFG):
    cols = CFG["cols"]; MF = CFG["fmt"]; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
    FADEFILL = PatternFill("solid", fgColor="E7E6E6")
    # --- In：目標價與收斂法參數（黃底；Assumed，Andy 可改） ---
    r = IN.max_row + 2
    IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（v2 新增；黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
    def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
        nonlocal r
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        c = IN.cell(r, 5, base)
        if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
        if fmt: c.number_format = fmt
        if yellow: c.fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    # v0.4：刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；DCF 折現率見 In 頁 M2 節（tools/fcff_dcf.py）
    inrow("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.4 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
    if CFG["nim"]:
        inrow("nim10", f"十年平均淨利率（{CFG['nim'][0]}，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(" + ",".join(f"{n}/{d}" for d, n in CFG["nim"][1]) + ")", "Derived", CFG["nim"][2], CFG["nim"][3], "0.0%")
    else:
        IN.cell(r, 1, "nim10"); IN.cell(r, 2, "十年平均淨利率"); IN.cell(r, 3, "%"); IN.cell(r, 5, "未找到"); IN.cell(r, 7, "Assumed")
        IN.cell(r, 8, CFG["nim_missing"]); IN.cell(r, 9, "依工作單：找不到十年平均時用最後一年（FY2029）淨利率 × 0.8；PL 延伸區塊逐情境計算"); r += 1
    inrow("cons_tp", "市場共識目標價（賣方平均）", CFG["cur"], CFG["cons_tp"][0], "Interested-party", CFG["cons_tp"][1], CFG["cons_tp"][2], MF)
    inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", CFG["cur"], CFG["cons_feps"][0], "Interested-party", CFG["cons_feps"][1], CFG["cons_feps"][2], MF)

    # --- PL：延伸到 FY2032（收斂法） ---
    p0 = PL.max_row + 3
    PL.cell(p0, 1, "延伸到 FY2032（v2 新增）：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
    PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均" + ("" if CFG["nim"] else "；十年平均未找到，以 FY2029 × 0.8 代替（Assumed）") + "）；股數維持 FY2029（EPS 與淨利同比例）")
    ext_cols = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
    hdr(PL, p0 + 2, ["情境（延伸）", "單位"] + [str(PL[f"{cols[y]}3"].value) for y in range(2025, 2030)] + ["FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    EXT = {}; pr = p0 + 3
    gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
    for s in SC:
        PL.cell(pr, 1, f"情境 {s}（{CFG['scn'][s]}）").font = BOLD; pr += 1
        R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
        for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
            PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": CFG["unit_rev"], "ni": CFG["unit_ni"], "eps": CFG["cur"]}.get(k, "%"))
        for y in range(2025, 2033):
            c = ext_cols[y]; pc = ext_cols.get(y - 1)
            if y <= 2029:
                PL[f"{c}{R['rev']}"] = f"={cols[y]}{SUMR[(s, 'rev')]}"
                PL[f"{c}{R['ni']}"] = f"={cols[y]}{SUMR[(s, 'ni')]}"
                PL[f"{c}{R['eps']}"] = f"={cols[y]}{SUMR[(s, 'eps')]}"
                PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
                if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            else:
                k = y - 2029
                mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+{'nim10' if CFG['nim'] else '0.8*$G' + str(R['nim'])})/2", "C": "nim10" if CFG["nim"] else f"0.8*$G{R['nim']}"}[s]
                PL[f"{c}{R['g']}"] = f"=$G{R['g']}+((tp_g{gadd[s]})-$G{R['g']})*{k}/3"
                PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
                PL[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({mT}-$G{R['nim']})*{k}/3"
                PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
                PL[f"{c}{R['eps']}"] = f"=$G{R['eps']}*{c}{R['ni']}/$G{R['ni']}"
                for kk in R: PL[f"{c}{R[kk]}"].fill = FADEFILL
            for kk, fm in (("rev", CFG["fmt_rev"]), ("g", "0.0%"), ("nim", "0.0%"), ("ni", CFG["fmt_rev"]), ("eps", MF)):
                if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
        PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
        EXT[s] = R; pr += 6
    for c in "IJ": PL.column_dimensions[c].width = 14
    PL.column_dimensions["K"].width = 40

    # --- Val：舊法改名（不再叫目標價） ---
    REN = (("0. 目標價與上檔空間", "舊法：規則值情境加權（參考）"), ("目標價（12 個月，情境加權）", "規則值情境加權價（12 個月；參考）"),
           ("上檔空間（目標價 ÷ 現價 − 1）", "相對現價（規則值情境加權價 ÷ 現價 − 1）"), ("目標價方法", "規則值方法"))
    for row in V.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            for a, b in REN:
                if v.startswith(a): row[0].value = b + v[len(a):]; break
    # --- Val：0. 目標價（列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期） ---
    t0 = V.max_row + 3
    T = {}
    def tset(rr, key, label, fs, fmt, note=""):
        V.cell(rr, 1, label); T[key] = rr
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
        if note: V[f"E{rr}"] = note
    V.cell(t0, 1, "0. 目標價").font = BOLD; T["top"] = t0
    V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["h01"] = t0 + 1
    hdr(V, t0 + 2, ["項目", f"A（{CFG['scn']['A']}）", f"B（{CFG['scn']['B']}）", f"C（{CFG['scn']['C']}）", "說明"]); T["hd1"] = t0 + 2
    V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"; T["eps_h"] = t0 + 3
    EPSR = {}
    for i, y in enumerate(YRS6):
        tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{ext_cols[y]}{EXT[s]['eps']}" for s in SC], MF); EPSR[y] = t0 + 4 + i
    rr = t0 + 10
    V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；設備股資本支出低（輕資產），此式偏誤小"; T["cf_h"] = rr
    tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
    RV = rr + 1; CFR = {}
    for i, y in enumerate(YRS6):
        tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], MF); CFR[y] = rr + 2 + i
    c1, c6 = CFR[2027], CFR[2032]
    tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], MF,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027" + CFG["dcf_note"])
    tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], MF)
    tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], MF, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
    q = rr + 12
    V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; T["h02"] = q
    hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["phd"] = q + 1
    for i, (nm, tk, pe, dt, src) in enumerate(CFG["peers"], q + 2):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    pl_ = q + 1 + len(CFG["peers"]); T["p1"] = q + 2; T["pl"] = pl_
    q = pl_ + 1
    tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{T['p1']}:C{pl_})", None, None], "0.0", CFG["peer_note"])
    fw = CFG["fwd"]
    def feps(c): return "(" + "+".join(f"{w}*{c}{EPSR[y]}" for y, w in fw.items()) + ")"
    tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("B") + "+tp_wB*" + feps("C") + "+tp_wC*" + feps("D"), None, None], MF, CFG["fwd_note"])
    tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 上列 EPS")
    q += 4
    V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
    hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
    q += 2
    tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
    tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2")
    tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, CFG["cons_tp"][1])
    tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", CFG["px_note"])
    tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
    tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；" + CFG["cons_feps"][1])
    tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
    V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
    tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"{CFG["band_fmt"]}")&"～"&TEXT(B{T["dcf12"]},"{CFG["band_fmt"]}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    PE = CFG["pe10"]
    tset(q + 8, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%", CFG["pe10_note"])
    V.cell(q + 9, 1, "目標價方法"); V[f"B{q + 9}"] = CFG["method"]; T["meth"] = q + 9
    for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
    if CFG.get("a3"): V["A3"] = CFG["a3"].format(row=t0)

    # --- Sens：目標價參數 ---
    s0 = SE.max_row + 3
    SE.cell(s0, 1, "目標價參數（v2 新增；只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
    hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_formula(c, r_="tp_r", g="tp_g"):
        cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
        return f"({pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.4", "Comparable 0.6", dict(w="0.4")), ("DCF 權重 0.8", "Comparable 0.2", dict(w="0.8"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf")
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        rr_ = kw.get("r_", "tp_r")
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{rr_})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = MF
        SE[f"F{i}"].number_format = "0.0%"
    T["sens0"] = s0; T["sens_last"] = s0 + 1 + len(cases)

    # --- README 頁 ---
    rr = RD.max_row + 1
    for t in CFG["readme"]:
        RD.cell(rr, 1, t); rr += 1
    return T, EXT

V2T, V2EXT = valuation_v2(CFG_V2)

# ---------------- v0.4（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py）取代股東現金流法 ----------------
# 依據：docs/plan/個股DCF標準規格.md v1.0、工作單「AI半導體 個股 DCF 改為標準版」（Andy 2026-11-05 20:42、20:45、21:21）。
# 做法：Val 0.1 節（原股東現金流）整段改寫為 FD 摘要（列數不變，0.2、0.3 節位置不動）；原 0.1 節的情境 EPS 表移到新 DCF 頁，
#   Comparable 的 12 個月後前瞻 EPS 與歷史倍數檢查改讀 DCF 頁（數值不變）；Sens 頁目標價參數段改為 FD 標準案例。
import re as _re3
def dcf_v03(T, paths, years, frac0, ver, a1, a2, MF, band_fmt, conv_from=None, nc_bs=None, nc_note="", detail_title=None, usd=False):
    DW = wb.create_sheet("DCF")
    DW["A1"] = a1; DW["A2"] = a2
    DW.cell(4, 1, "情境 EPS（FY2027–FY2032；v0.2 起在 Val 0.1 節，" + ver + " 移到本頁；Val 的 12 個月後前瞻 EPS 與歷史倍數檢查讀本表，數值不變）").font = BOLD
    hdr(DW, 5, ["項目", "A", "B", "C", "說明"])
    emap = {}
    for i, y in enumerate((2027, 2028, 2029, 2030, 2031, 2032)):
        rv = T[f"eps{y}"]; r = 6 + i; emap[rv] = r
        DW.cell(r, 1, f"  FY{y}")
        for c in "BCD":
            DW[f"{c}{r}"] = V[f"{c}{rv}"].value; DW[f"{c}{r}"].number_format = MF
    DW.cell(6, 5, "PL 頁延伸區塊；FY2030 起為收斂法")
    rx = _re3.compile(r"(?<![A-Za-z_!$\d])\$?([BCD])\$?(\d+)(?![\d(])")
    for k in ("feps", "pr"):
        for c in "BCD":
            v = V[f"{c}{T[k]}"].value
            if isinstance(v, str):
                V[f"{c}{T[k]}"].value = rx.sub(lambda m: f"DCF!${m.group(1)}${emap[int(m.group(2))]}" if int(m.group(2)) in emap else m.group(0), v)
    R = FD.write_detail(DW, 14, paths, years, frac0, conv_from=conv_from, fmt=MF, stub=True,
                        title=detail_title or "三情境 FCFF 明細（金額單位同 PL 頁；年度一欄；期中慣例）")
    R["end"] = FD.write_checks(DW, R["end"], nc_bs=nc_bs, nc_note=nc_note) + 1
    DW.column_dimensions["A"].width = 60; DW.column_dimensions["B"].width = 40
    for c in "CDEFGH": DW.column_dimensions[c].width = 13
    r1, r2 = T["h01"], T["dcftp"]
    assert str(V.cell(r1, 1).value).startswith("0.1 DCF"), V.cell(r1, 1).value
    _ns = FD.summary_rows(stub=True)   # 第二輪修正：模組 v2 stub 期在「＋淨現金與投資」後多兩列，用原 0.1 節之後的空白列
    for rr in range(r2 + 1, r1 + _ns):
        assert all(V.cell(rr, cc).value is None for cc in range(1, 8)), (rr, V.cell(rr, 1).value)
    assert V.cell(r1 + _ns, 1).value is None, V.cell(r1 + _ns, 1).value
    r2 = max(r2, r1 + _ns - 1)
    for rr in range(r1, r2 + 1):
        for cc in range(1, 8):
            V.cell(rr, cc).value = None; V.cell(rr, cc).fill = PatternFill(); V.cell(rr, cc).font = Font()
    S = FD.write_summary(wb, V, r1, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt=MF,
                         title=f"0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；{ver}；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
    V[f"B{T['tpw']}"] = f"=tp_w_dcf*B{S['dcf']}+(1-tp_w_dcf)*B{T['comptp']}"
    V[f"B{T['band']}"] = f'=TEXT(D{S["v12"]},"{band_fmt}")&"～"&TEXT(B{S["v12"]},"{band_fmt}")'
    s0, s1 = T["sens0"], T["sens_last"]
    for rr in range(s0, s0 + 16):
        for cc in range(1, 10):
            if rr > s1: assert SE.cell(rr, cc).value is None, (rr, cc, SE.cell(rr, cc).value)
            SE.cell(rr, cc).value = None; SE.cell(rr, cc).fill = PatternFill(); SE.cell(rr, cc).font = Font()
    ts = FD.write_sens(SE, s0, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt=MF,
                       title=f"目標價參數敏感度（{ver} 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=usd)
    SE.column_dimensions["G"].width = 14
    T.update(dcf=S["dcf"], v12=S["v12"], S=S, R=R, ts=ts)
    for k in ("eps_h", "cf_h", "reinv", "dcfv", "dcf12", "dcftp") + tuple(f"eps{y}" for y in range(2027, 2033)) + tuple(f"cf{y}" for y in range(2027, 2033)):
        T.pop(k, None)
    T["dcftp"] = S["dcf"]; T["sens_last"] = ts[1]
    # 第二輪修正：模組 v2 的 In 頁 M2 標題不寫版號（原本此處改寫版號的段落刪除）
    return T
_EXTC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}   # PL 延伸區塊欄（valuation_v2 同式）
FY_X = list(range(2027, 2033))
def _ebit(s_, y):
    return f"PL!{FYC[y]}{PB[s_]['oi']}" if y <= 2029 else f"=PL!{_EXTC[y]}{V2EXT[s_]['ni']}/(1-tax)-oth_exp"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{CFG_V2['scn'][s_]}）", rev={y: f"PL!{_EXTC[y]}{V2EXT[s_]['rev']}" for y in FY_X}, rev0=f"PL!{_EXTC[2026]}{V2EXT[s_]['rev']}",
              ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "tax" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷（1 − tax）− oth_exp（PL 延伸區塊只有淨利）") for s_ in SC]
V2T = dcf_v03(V2T, DPATHS, FY_X, 0.5, "v0.4",
              "DCF — 標準版（v0.4）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（v0.4.1 起；原 FY2026 年底 2026-06-30；期中慣例）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py",
              "EBIT：FY2027–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸區塊只有淨利）由淨利還原：EBIT＝淨利 ÷（1 − tax）− oth_exp（與 PL 淨利＝（營業利益＋營業外）×（1 − tax）同式）。稅率：tax（14.5%）。"
              "營業外淨額不進 NOPAT，現金另在評價時加入。折舊、資本支出：近三年占營收比例（In 頁 fc_da、fc_cx），FY2028 起資本支出線性收斂到長期穩態 fc_cx_ss；營運資金：近三年平均占營收 × 營收增量。",
              CFG_V2["fmt"], CFG_V2["band_fmt"], nc_bs="1532+3209-5723-(54.01+215.34)",
              nc_note="Q1 FY2027 資產負債表（US$M）：現金 1,532＋短期投資 3,209 − 長期債務 5,723 − 租賃（54.01＋215.34）",
              detail_title="三情境 FCFF 明細（US$M；FY2027＝2026-07～2027-06，評價日 2026-12-31 後只計 6 個月）")
if os.environ.get("V2_ROWS"):
    print("V2", V2T, V2EXT)

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.2', 'v0.2.1'
OBS = [
 "目標價 US$119（DCF 95、Comparable 相對倍數調整後 143 各半），比現價 145.61（2026-10-23）低 18.3%，級距 0%～−25%，投資建議維持中立。v0.3 滾入 FQ1-27（營收 4,136.0、GAAP EPS 1.18，高於指引中點 4.0B／1.14）與 FQ2-27 指引（營收 4.7B±0.2B、GAAP EPS 1.41±0.11）；加權目標價 115.2 → 118.9（+3.2%）。",
 "CY2026 下半年較上半年增幅由公司口述的 20% 改為實績＋指引推得的 24.9%（區間 22.1%～27.8%），FY2027 EPS A／B／C 5.24／5.19／5.12 → 5.48／5.43／5.36。DCF 仍比現價低 34.7%，主因情境 B、C 的 EPS 在 FY2029 後依收斂法下滑；Comparable 單獨 +3.0%。",
 "中國出口管制仍是最敏感的變數（Sens 頁）。股價仍為 2026-10-23 收盤，財報後未更新。下一個檢驗點是 FQ2-27 財報（2027-01 下旬）：CY2027 上半年是否延續成長，以及公司對 2027 年 WFE 的看法。"
]   # 重點觀察（數字取自重算後的 Val／PL／Sens；與 README.md 相同文字）
UPW_ROW = next(r_ for r_ in range(1, V.max_row + 1) if V.cell(r_, 1).value == "上檔空間：加權目標價")
GRADE_ROW = V.max_row + 1; REC_ROW = GRADE_ROW + 1
_u = f"B{UPW_ROW}"
V.cell(GRADE_ROW, 1, "上檔空間級距")
V[f"B{GRADE_ROW}"] = f'=IF({_u}>0.5,"大於 50%",IF({_u}>0.25,"25%–50%",IF({_u}>=0,"0%–25%",IF({_u}>=-0.25,"0%～−25%",IF({_u}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{GRADE_ROW}"].number_format = "@"
V[f"E{GRADE_ROW}"] = f"以「上檔空間：加權目標價」（第 {UPW_ROW} 列）分六級（Andy 2026-10-25 14:00）"
V.cell(REC_ROW, 1, "投資建議")
V[f"B{REC_ROW}"] = f'=IF({_u}>0.5,"強力買進",IF({_u}>0.25,"買進",IF({_u}>=-0.25,"中立",IF({_u}>=-0.5,"放空","強力放空"))))'
V[f"B{REC_ROW}"].number_format = "@"
V[f"E{REC_ROW}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）"
_t = RD["A1"].value
assert VER_PREV in _t and VER_NOW not in _t, _t
RD["A1"] = _t.replace(VER_PREV, VER_NOW, 1)
for r_ in range(2, RD.max_row + 1):
    _v = RD.cell(r_, 1).value
    if isinstance(_v, str) and _v.startswith("版本："):
        RD.cell(r_, 1, "版本：" + VER_NOW + "（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）。前版：" + _v[3:])
        break
rn = RD.max_row + 1
RD.cell(rn, 1, VER_NOW + "：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（第 " + str(GRADE_ROW) + "、" + str(REC_ROW) + " 列，依「上檔空間：加權目標價」的固定規則）；重點觀察見下方三列。")
RD.cell(rn + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"')
for k_, o_ in enumerate(OBS, 1):
    RD.cell(rn + 1 + k_, 1, f"重點觀察 {k_}：" + o_)


# ---------------- 分部原則第 3 條（2026-10-28）：核對列與宣告（追加於工作表最後，不位移既有列，數值不變） ----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(_os2.path.dirname(_os2.path.dirname(_os2.path.abspath(__file__))), "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[7, 8, 9, 10], total_row=4, cols="BCDEFGHI", rnd=3, note="分部原則第 3 條（2026-10-28 追加）：含第 10 列匯率影響等（營收 − 三事業，Derived）；分部宣告見 segments.json")
RD.cell(RD.max_row + 1, 1, "v0.2.2（2026-10-28）：分部原則重建（數值不變）。Q 頁最後追加核對列「檢查：分部合計 − 合併營收」；官方分部與子列宣告見 segments.json。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.2.1", "v0.2.2")

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402

# ---------------- Wattonomics 錨定（三層校準原則；Siliconomics v0.13，決策 S37）----------------
# In：三列 wta_27／28／29 接在模型輸入最後（apply_in_layout 的「情境驅動對照」展示節接在其後）；Sens：案例區塊放在既有內容最下方，摘要列接在案例清單之後
sys.path.insert(0, os.path.join(ROOT, "siliconomics")); import reader as SIL  # noqa: E402
from openpyxl.workbook.defined_name import DefinedName as _DN  # noqa: E402
from openpyxl.styles import Font as _F  # noqa: E402
def _wta_in(wb):
    ws = wb["In"]; r = ws.max_row + 1
    ws.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = _F(bold=True); r += 1
    rows_ = []; SIL.wta_rows(lambda *a: rows_.append(a))
    for a in rows_:
        for j, v in enumerate(a, 1):
            if v is not None: ws.cell(r, j, v)
        ws.cell(r, 5).font = _F(color="0000FF")
        wb.defined_names[a[0]] = _DN(a[0], attr_text=f"In!$E${r}")
        r += 1
_wta_in(wb)
WTA_READ = "讀法：Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"

def _wta_f(SE, w0, CM, desc):
    SE.cell(w0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    SE.cell(w0 + 1, 1, WTA_READ)
    SE.cell(w0 + 2, 1, desc)
    hdr(SE, w0 + 3, ["日曆年", "單位"] + [f"CY{y}" for y in range(2025, 2030)] + ["", "來源"])
    fr = w0 + 4
    SE.cell(fr, 1, "Wattonomics 錨定係數（In 頁 wta_27／28／29）"); SE.cell(fr, 2, "x")
    for y, f in zip(range(2025, 2030), (1, 1, "=wta_27", "=wta_28", "=wta_29")): SE[f"{CM[y]}{fr}"] = f; SE[f"{CM[y]}{fr}"].number_format = "0.000"
    return fr
WTA_AIWFE = "產業量＝設備市場模型（markets/equipment，情境 B）中標〔AI〕的分項：先進邏輯〔AI〕、DRAM（含 HBM）〔AI〕、先進封裝〔AI〕；CY2027 起 × wta，其他分項不變。注意：設備市場由公司資本支出推得，與部署 GW 只是間接連動（資本支出領先產能），此處依工作單以同年係數近似。"

fr = _wta_f(SE, SE.max_row + 3, {y: ccy(y) for y in range(2025, 2030)}, WTA_AIWFE + " 本模型：加權前段 WFE 扣除（先進邏輯〔AI〕＋m × DRAM〔AI〕）×（1 − wta）；先進封裝設備市場 × wta（Units 頁情境 B 列）。服務、SPTS＋PCB、強度、中國照情境 B。")
RW = engine(SE, fr + 3, "案例區塊：Wattonomics 錨定（情境 B 為底；設備市場 AI 分項 CY2027 起 × wta）", scn="B", wf={y: f"${ccy(y)}${fr}" for y in (2027, 2028, 2029)})
put_case(nn + 2, "Wattonomics 錨定（情境 B 為底）", "設備市場 AI 分項 CY2027 起 × wta；相對欄對情境 B 列", RW)
for a_, b_ in (("E", "D"), ("H", "G")): SE[f"{a_}{nn + 2}"] = f"={b_}{nn + 2}/{b_}{nn}-1"
RMX.apply(wb, "klac")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
RD.cell(RD.max_row + 2, 1, "v0.3（2026-10-31）：滾入 FY2027 第一季（2026-10-28：營收 4,136.0、產品 3,285.5、服務 850.5、淨利 1,548.6、GAAP EPS 1.18；Q 頁 I 欄）與 FQ2-27 指引（營收 4.7B±0.2B、GAAP 毛利率 61.8%、GAAP EPS 1.41）。CY2026 下半年增幅 h2up 改為 Derived（FQ1-27 實績＋FQ2-27 指引 ÷ 上半年 − 1＝24.9%；原公司口述 20%）；稀釋股數 1,311.1。FY2027 EPS（A）5.24 → 5.48；加權目標價 115.2 → 118.9，投資建議維持中立。")
apply_in_layout.apply(wb, "v0.2.3", "2026-10-30")  # v0.2.3：In 頁版面（個股模型原則第 8 節；數值不變）
# ---------------- v0.3.1（2026-11-05）：只改展示層——README 新增「展示層」「命題回覆」兩段（個股 HTML 版面規格 v1.0；公式帶數字）----------------
# 新列放在 README 既有內容之後，不位移既有列；README 以外的工作表不動
def _rw(ws, text, exact=False, col=1, start=1):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, col).value
        if isinstance(v_, str) and ((v_.strip() == text) if exact else v_.strip().startswith(text)):
            return r_
    raise KeyError((ws.title, text))
def _F(*parts):   # 文字公式：字串常數每段 ≤ 120 字（Excel 公式內字串常數上限 255 字元）；("=", 運算式) 為公式片段
    out = []
    for p_ in parts:
        if isinstance(p_, tuple):
            out.append(p_[1])
        else:
            s_ = str(p_)
            out += ['"' + s_[i_:i_ + 120].replace('"', '""') + '"' for i_ in range(0, len(s_), 120)]
    return "=" + "&".join(out)
_E = lambda e: ("=", e)
_TX = lambda c, f: _E(f'TEXT({c},"{f}")')
_SG = lambda c: _TX(c, "+0.0%;-0.0%")
_PC = lambda c: _TX(c, "0.0%")
_UP = lambda c: _E(f'IF({c}>=0,"高 ","低 ")&TEXT(ABS({c}),"0.0%")')
_ws_in = wb["In"]
_IN = {_ws_in.cell(r_, 1).value: r_ for r_ in range(4, _ws_in.max_row + 1) if _ws_in.cell(r_, 1).value}
I_ = lambda k: f"In!E{_IN[k]}"
I_lo = lambda k: f"In!D{_IN[k]}"
I_hi = lambda k: f"In!F{_IN[k]}"
_V, _RC, _SE, _RD = wb["Val"], wb["Recon"], wb["Sens"], wb["README"]
vw, vd, vc = _rw(_V, "目標價（DCF 與 Comparable 加權）", exact=True), _rw(_V, "DCF 目標價（情境加權）"), _rw(_V, "Comparable 目標價", exact=True)
vca, vk, vm = _rw(_V, "Comparable 目標價（相對倍數調整後）"), _rw(_V, "相對倍數調整係數"), _rw(_V, "同業前瞻 P/E 中位數")
vup, vuc, vcons, vrec = _rw(_V, "上檔空間：加權目標價"), _rw(_V, "上檔空間：Comparable"), _rw(_V, "市場共識目標價"), _rw(_V, "投資建議", exact=True)
vpx = _rw(_V, "現價（")
vfe = _rw(_V, "12 個月後的前瞻 EPS")
vd12 = _rw(_V, "DCF 情境價值推到 12 個月後")
e0 = _rw(wb["DCF"], "情境 EPS（FY2027") + 1   # v0.4：情境 EPS 表移到 DCF 頁
E27, E28, E29, E32 = e0 + 1, e0 + 2, e0 + 3, e0 + 6
W_ = I_("tp_w_dcf")
# 加權目標價相對現價的差距拆解（補充規則 B）：DCF 貢獻＝w ×（DCF − 現價），Comparable 貢獻＝(1 − w) ×（調整後 Comparable − 現價）
SH_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/(Val!B{vw}-Val!B{vpx})"
def _block(lead, props, obs, answers):
    rr = _RD.max_row + 2
    _RD.cell(rr, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；" + VER_HTML + " 新增，上方各列保留）").font = Font(bold=True); rr += 1
    for t_ in [lead] + props + obs:
        _RD.cell(rr, 1, t_); rr += 1
    rr += 1
    _RD.cell(rr, 1, "命題回覆（HTML 用；依本模型計算回答開頭的命題）").font = Font(bold=True); rr += 1
    for t_ in answers:
        _RD.cell(rr, 1, t_); rr += 1
    rr += 1
    _RD.cell(rr, 1, "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）的上檔空間（＝Val 相對倍數調整後 Comparable 目標價 ÷ 現價 − 1；進加權目標價的是該值，橫幅中間格則放未調整的只看 Comparable 目標價）")
    _RD.cell(rr, 2, f"=Val!B{vca}/Val!B{vpx}-1").number_format = "0.0%"
    _RD["A1"] = str(_RD["A1"].value).replace("v0.2.2", "v0.4", 1) + "；" + VER_HTML + "（2026-11-05）只改展示層（README 展示層與命題回覆、HTML 新版面 v1.0），數值不變；v0.4（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
    return rr

CT_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/Val!B{vpx}"          # DCF 對加權上檔空間的貢獻（占現價）
CT_CMP = f"(1-{W_})*(Val!B{vca}-Val!B{vpx})/Val!B{vpx}"     # 調整後 Comparable 對加權上檔空間的貢獻（占現價）
T2 = lambda c: _TX(c, "0.00")
T1 = lambda c: _TX(c, "0.0")
T0 = lambda c: _TX(c, "#,##0")
VER_HTML = "v0.3.1"
r_e27, r_e28 = _rw(_RC, "FY2027 EPS"), _rw(_RC, "FY2028 EPS")
s_cn25, s_cn50, s_c, s_pc0 = _rw(_SE, "中國系統營收 −25%"), _rw(_SE, "中國系統營收 −50%"), _rw(_SE, "設備市場情境 C"), _rw(_SE, "製程控制強度不再提升")
s_w8 = _rw(_SE, "DCF 權重 0.3／0.7")
_PLs = wb["PL"]
p_e = _rw(_PLs, "EPS（GAAP 稀釋）")
TB, TC = _E(f"PL!H{_rw(_PLs, '情境 B', exact=True)}"), _E(f"PL!H{_rw(_PLs, '情境 C', exact=True)}")
# v0.4：舊版（v0.3.1，股東現金流法 DCF）目標價對照列——常數取自 v0.3.1 重算後 xlsx（klac/dist/_待刪除/）Val 頁，只作對照、不進任何計算
_ro = _RD.max_row + 2
_RD.cell(_ro, 1, "v0.3.1 → v0.4 對照：舊法（股東現金流 DCF，折現率 10%）的目標價（常數，取自 v0.3.1 重算後 xlsx 的 Val 頁；只作對照）").font = Font(bold=True)
OLD = {}
for k_, (key_, lab_, v_, f_) in enumerate((("dcf", "v0.3.1 DCF 目標價（舊法：股東現金流）", 95.0396149502467, "#,##0.00"),
                                          ("tp", "v0.3.1 目標價（DCF 與 Comparable 加權）", 118.897052166628, "#,##0.00"),
                                          ("up", "v0.3.1 上檔空間：加權目標價", -0.183455448344015, "0.0%"),
                                          ("rec", "v0.3.1 投資建議", "中立", "@")), 1):
    _RD.cell(_ro + k_, 1, lab_); _c = _RD.cell(_ro + k_, 2, v_); _c.font = BLUE; _c.number_format = f_; OLD[key_] = _ro + k_
lead = "命題：科磊的系統營收跟著晶圓廠設備支出走，再乘上製程控制強度（先進製程與先進封裝每片晶圓需要更多量測與檢測），中國約占營收三成；問題是設備支出在 2027 年後放緩時，強度上升與服務能否撐住盈餘、撐住現價。三個問題如下。"
props = [
    "命題 1：設備支出放緩時盈餘還會成長嗎——2028 年設備支出持平或下滑時，製程控制強度上升與服務成長能否讓 FY2028、FY2029 每股盈餘仍高於 FY2027。",
    "命題 2：中國出口管制會讓盈餘少多少——中國系統營收被砍四分之一或一半時，每股盈餘的影響有多大。",
    "命題 3：每股值多少——以 DCF 與同業倍數加權的目標價，相對現價的上檔空間是多少。",
]
obs = [
    _F("重點 1：加權目標價低於現價，投資建議", _E(f"Val!B{vrec}"), "：加權目標價約 ", T2(f"Val!B{vw}"), " 美元，比現價 ", T2(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"),
       _E(f'IF(ABS(Val!B{vup})<=0.25,"（在正負 25% 之間為中立）",IF(Val!B{vup}<0,"（−25%～−50% 為放空）","（+25%～+50% 為買進）"))'), "。拆解：DCF（", T2(f"Val!B{vd}"), "，權重 ", _TX(W_, "0%"), "）拉低 ", _PC(f"-{CT_DCF}"), "，相對倍數調整後的 Comparable（", T2(f"Val!B{vca}"),
       "）拉低 ", _PC(f"-{CT_CMP}"), "，低於現價主要來自 DCF。"),
    _F("重點 2：FY2027 與共識相近，FY2028 三情境都低於共識：FY2027 每股盈餘順利／平平／不如預期 ", T2(f"Recon!C{r_e27}"), "／", T2(f"Recon!E{r_e27}"), "／", T2(f"Recon!G{r_e27}"),
       " 美元，共識 ", T2(f"Recon!B{r_e27}"), "；FY2028 為 ", T2(f"Recon!C{r_e28}"), "／", T2(f"Recon!E{r_e28}"), "／", T2(f"Recon!G{r_e28}"), "，比共識 ", T2(f"Recon!B{r_e28}"), " 低 ",
       _PC(f"-Recon!D{r_e28}"), "／", _PC(f"-Recon!F{r_e28}"), "／", _PC(f"-Recon!H{r_e28}"), "（同一年度比較；注意本模型為 GAAP 每股盈餘、共識為非 GAAP，口徑不同，非 GAAP 通常較高）。"),
    _F("重點 3：設備市場情境與中國最敏感（以順利情境為基準，Sens 頁）：改用不如預期情境（", TC, "；服務與兩事業成長放緩）使 FY2028 每股盈餘 ", _SG(f"Sens!H{s_c}"),
       "；中國系統營收自 CY2027 起 −50% 使 ", _SG(f"Sens!H{s_cn50}"), "；製程控制強度不再提升使 ", _SG(f"Sens!H{s_pc0}"), "。"),
    _F("重點 4：DCF 改為標準版（無槓桿自由現金流折現）：推到 12 個月後的 DCF 情境價值順利／平平／不如預期為 ", T2(f"Val!B{vd12}"), "／", T2(f"Val!C{vd12}"), "／", T2(f"Val!D{vd12}"),
       "，DCF 目標價 ", T2(f"Val!B{vd}"), "（舊法股東現金流 ", T2(f"README!B{OLD['dcf']}"), "）。WACC ", _TX("wacc", "0.0%"), "（股權成本 ", _TX("ke", "0.0%"), "＝rf ", _TX("rf", "0.00%"),
       " ＋ β ", _TX("beta", "0.00"), " × ", _E('IF(crp>0,"（ERP "&TEXT(erp,"0%")&" ＋ CRP "&TEXT(crp,"0.00%")&"）","ERP "&TEXT(erp,"0%"))'), "；β 為自算產業 β（美國半導體設備同業中位數，Blume 調整）重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設）；",
       "平平與不如預期情境的盈餘在 FY2029 後依收斂法變化（FY2032 每股盈餘 ", T2(f"DCF!C{E32}"), "、", T2(f"DCF!D{E32}"), "；FY2030 起為收斂法、非驅動因子推導，是假設）。終值現值占平平情境企業價值 ",
       _TX(f"Val!C{_rw(_V, '其中：終值現值占 EV')}", "0%"), "、隱含終值 EV／EBITDA ", _TX(f"Val!C{_rw(_V, '隱含終值 EV／EBITDA')}", "0.0"), " 倍。以情境加權目標價為基準，WACC −1pt／＋1pt 時加權目標價 ",
       _SG(f"Sens!D{_rw(_SE, 'WACC ±1pt')}"), "／", _SG(f"Sens!F{_rw(_SE, 'WACC ±1pt')}"), "；「", _E(f"Sens!A{s_w8}"), "」時 ", _SG(f"Sens!D{s_w8}"), "／", _SG(f"Sens!F{s_w8}"), "。"),
]
answers = [
    _F("設備支出放緩時盈餘還會成長嗎？答：設備支出持平時仍小幅成長，大幅下滑時不行。FY2027 三情境每股盈餘 ", T2(f"DCF!B{E27}"), "／", T2(f"DCF!C{E27}"), "／", T2(f"DCF!D{E27}"),
       " 美元（順利／平平／不如預期；FY2026 實績 ", T2(f"PL!D{p_e}"), "）。平平情境（", TB, "）FY2028／FY2029 為 ", T2(f"DCF!C{E28}"), "／", T2(f"DCF!C{E29}"),
       "，靠製程控制強度每年提升 ", _PC(I_("pc_u")), "（由公司 2030 年模型推得）與服務年增 ", _PC(I_("g_svc_B")), "（假設）撐住；不如預期情境（", TC, "）降到 ", T2(f"DCF!D{E28}"), "／", T2(f"DCF!D{E29}"),
       "，低於 FY2027。以順利情境為基準，強度不再提升使 FY2028 每股盈餘 ", _SG(f"Sens!H{s_pc0}"), "。"),
    _F("中國出口管制會讓盈餘少多少？答：中國占 FY2026 營收 ", _PC(I_("chn_sh26")), "，基準不另扣。以順利情境為基準（Sens 頁），中國系統營收自 CY2027 起 −25% 使 FY2028 每股盈餘 ",
       _SG(f"Sens!H{s_cn25}"), "、−50% 使 ", _SG(f"Sens!H{s_cn50}"), "；下降比例是敏感度假設。"),
    _F("每股值多少？答：加權目標價約 ", T2(f"Val!B{vw}"), " 美元，比現價 ", T2(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"), "，投資建議", _E(f"Val!B{vrec}"),
       "。加權目標價＝DCF ", T2(f"Val!B{vd}"), " 與 Comparable 各占 ", _TX(W_, "0%"), "／", _TX(f"1-{W_}", "0%"), "；進加權的 Comparable 是相對倍數調整後的 ", T2(f"Val!B{vca}"),
       "（同業前瞻本益比中位數 ", T1(f"Val!B{vm}"), " 倍 × 相對倍數調整係數 ", T2(f"Val!B{vk}"), " × 12 個月後前瞻每股盈餘 ", T2(f"Val!B{vfe}"), "，上檔空間 ", _SG(f"Val!B{vca}/Val!B{vpx}-1"),
       "）。只看 Comparable 目標價（市場主流同業法，不做相對倍數調整）為 ", T2(f"Val!B{vc}"), "（上檔空間 ", _SG(f"Val!B{vuc}"), "），頁首橫幅中間格放這個值，用來對照市場主流的目標價。"),
]
# 第二輪修正（2026-11-05）：展示層重點——淨現金滾到評價日（stub）、美元口徑折現（非美國公司）、終值占比 > 80% 標示（公式判斷；Sens 以加權目標價為基準）
def _r2pt(n, usd, ccy_fmt="#,##0.00", amt_fmt="#,##0.0", unit=""):
    _Vw, _Sw = wb["Val"], wb["Sens"]
    def _f(ws, lab, start=1):
        return next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(lab))
    h = _f(_Vw, "0.1 DCF"); nc, st, ncr, tv = (_f(_Vw, l, h) for l in ("＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "其中：終值現值占 EV"))
    t = lambda c, f: f'TEXT({c},"{f}")'
    s = (f'="重點 {n}：淨現金由資產負債表日 "&bs_date&" 滾到評價日 2026-12-31：資產負債表淨現金與投資 "&{t(f"Val!C{nc}", amt_fmt)}&" ＋ stub 期 FCFF（情境 B 第一年 FCFF × "&TEXT(stub_frac,"0.000")&" 年）"&{t(f"Val!C{st}", amt_fmt)}'
         f'&" − 已宣告股利 "&{t("stub_div", amt_fmt)}&" ＝ "&{t(f"Val!C{ncr}", amt_fmt)}&"{unit}。"')
    if usd:
        u = _f(_Sw, "美元口徑折現")
        s += (f'&"美元口徑折現（Sens 案例：rf 改用美國 10 年期 "&TEXT(rf_us,"0.00%")&"、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")＝"&TEXT(MIN(g_cap,rf_us),"0.0%")&"，CRP、匯率與現金流不變）時，加權目標價 "'
              f'&{t(f"Sens!C{u}", ccy_fmt)}&"，相對基準加權目標價 "&TEXT(Sens!D{u},"+0.0%;-0.0%")&"；基準採當地 10 年期利率 "&TEXT(rf,"0.00%")&"＋固定匯率。"')
    s += (f'&IF(MAX(Val!B{tv}:D{tv})>0.8,"終值現值占 EV 超過 80% 的情境："&IF(Val!B{tv}>0.8,"A "&TEXT(Val!B{tv},"0%")&" ","")&IF(Val!C{tv}>0.8,"B "&TEXT(Val!C{tv},"0%")&" ","")&IF(Val!D{tv}>0.8,"C "&TEXT(Val!D{tv},"0%"),"")&"（DCF 價值大多來自終值，對 WACC 與 g 較敏感）。",'
          f'"三情境終值現值占 EV "&TEXT(Val!B{tv},"0%")&"／"&TEXT(Val!C{tv},"0%")&"／"&TEXT(Val!D{tv},"0%")&"（A／B／C），皆未超過 80%。")')
    return s
obs.append(_r2pt(len(obs) + 1, False, amt_fmt="#,##0", unit="（US$M）"))
_block(lead, props, obs, answers)
_RD.cell(_RD.max_row + 2, 1, "v0.4（2026-11-05）：DCF 改為標準版（docs/plan/個股DCF標準規格.md v1.0；tools/fcff_dcf.py）。(1) Val 0.1 節改為 FCFF＋WACC 摘要（折現率、明確期現值、終值現值、EV、淨現金、每股價值、終值占比、隱含終值倍數），列數不變；(2) 新增 DCF 頁：三情境 FCFF 明細與檢查列；原 Val 0.1 節的情境 EPS 表移到 DCF 頁（Comparable 讀此表，數值不變）；"
                "(3) In 頁新增 M2 節（rf 美國 10 年期公債、Damodaran 美國半導體設備產業 β 重新加槓桿、ERP 5%、g＝MIN(3%, rf)、RONIC＝WACC、資本結構、折舊／資本支出／營運資金比例）；刪除 tp_r、tp_roe；(4) Sens 頁目標價參數段改為標準案例。新舊目標價對照見本頁「v0.3.1 → v0.4 對照」列。")
# v0.4.1（2026-11-05）：β 改自算產業 β（US 半導體設備同業中位數、Blume 調整；Andy 2026-11-05 22:55）、評價日改為 2026-12-31，其他模型不變
_RD["A1"] = str(_RD["A1"].value) + "；v0.4.1（2026-11-05）β 改自算產業 β（US 半導體設備同業中位數、Blume 調整；Andy 2026-11-05 22:55）、評價日改為 2026-12-31，其他模型不變"
_RD.cell(_RD.max_row + 2, 1, "v0.4.1（2026-11-05）：(1) β 由 Damodaran 美國產業表改為自算產業 β——US 半導體設備同業 7 家（AMAT、LRCX、KLAC、TER、FORM、COHU、KLIC）Yahoo Finance 月報酬 2021-11～2026-10 對 ^GSPC 迴歸、Hamada 去槓桿、現金修正後的中位數，Blume 調整 0.67β＋0.33，再依本公司 D/E 與稅率重新加槓桿（In 頁 beta）；Sens β 低／高改為中位數 × 0.8 再 Blume／中位數未調整。(2) 評價日由 FY2026 年底（2026-06-30）改為 2026-12-31（全模型統一），FY2027 只計評價日後 6 個月，目標時點 2027-12-31；淨現金沿用 Q1 FY2027 資產負債表（2026-09-30）。其他模型不變")
# 第二輪修正（2026-11-05；工作單 2026-11-05 23:20，Andy 23:16 定案）：共用模組 tools/fcff_dcf.py v2
_RD["A1"] = str(_RD["A1"].value) + "；v0.4.2（2026-11-05）第二輪修正：換用 tools/fcff_dcf.py v2——淨現金由 Q1 FY2027 資產負債表日 2026-09-30 滾到評價日 2026-12-31（stub 期 FCFF；stub 期無已宣告股利）、β 敏感度改 β_u ± 0.3 重新加槓桿"
_RD.cell(_RD.max_row + 2, 1, "v0.4.2（2026-11-05）：第二輪修正（工作單 2026-11-05 23:20；Andy 2026-11-05 23:16 定案），換用共用模組 tools/fcff_dcf.py v2。(1) Stub 期：資產負債表日 2026-09-30（Q1 FY2027 10-Q，最新已公布一季）≠ 評價日 2026-12-31，淨現金滾到評價日＝資產負債表淨現金＋各情境 FY2027 FCFF × stub_frac − 已宣告股利（0：2026-12 發放的下一次建置日尚未宣告；買回無已宣告金額，不扣）；In 頁新增 bs_date、stub_frac、stub_div，Val 0.1 節多「＋Stub 期 FCFF」「淨現金滾到評價日」兩列（0.2 節以後位置不動）。(2) β：In 頁新增 beta_u（1.503）、beta_de、beta_du（0.3），beta＝beta_u 重新加槓桿；Sens β 低／高改為 β_u ∓ 0.3 各自重新加槓桿。(3) ke 說明格由模組參數提供；In 頁 M2 標題不寫版號。美國公司不加「美元口徑折現」案例（＝基準）。資本支出、折舊：PL 沒有逐年列，維持占營收比例。展示層新增重點 5（stub 淨現金、終值占比是否超過 80%，公式帶數字）。PL、Comparable、情境權重不變。")

# ---------------- 第三輪修正（2026-11-06；合併前）：展示與說明字串（fixes_round3 第 14、16、22、23、24 項）＋各公司專屬修正 ----------------
# 只改說明字串、列名與 README 舊列；數值變動只來自各公司專屬修正（見呼叫處註解）。共用模組 tools/fcff_dcf.py v2 不改，說明字串在此覆寫。
def _r3_fix(wb, ver_new, a1_note, rd_note, stub_note=None, stub_frac_note=None, stub_pt=None, rd_clear=(), rd_set=None):
    RDw, INw, Vw, SEw = wb["README"], wb["In"], wb["Val"], wb["Sens"]
    # (22) README 舊「重點觀察」列（v0.3 前的過時數字）：清空內容（不刪列，其他公式引用 README 列號），首列留一句說明
    cl = [r for r in range(1, RDw.max_row + 1) if isinstance(RDw.cell(r, 1).value, str) and RDw.cell(r, 1).value.startswith(("重點觀察 ", '="重點觀察 '))]
    cl += [r for r in rd_clear if r not in cl]
    for r in sorted(cl):
        RDw.cell(r, 1).value = None
        for c in range(2, 6):
            RDw.cell(r, c).value = None
    if cl:
        RDw.cell(min(cl), 1, f"（{ver_new} 第三輪修正：本處原為 v0.2.1 起的「重點觀察」與相關舊說明，數字為 v0.3 標準版 DCF 之前的舊法、已過時，刪除；現行重點見本頁「展示層」重點各列，數字皆為公式）")
    for r, t in (rd_set or {}).items():
        RDw.cell(r, 1).value = t
    # (23) 情境名稱統一：情境 A（順利）／B（平平）／C（不如預期）（Andy 10-25 命名）；只改文字格，不改公式格
    MAP = (("情境 A（樂觀）", "情境 A（順利）"), ("情境 B（基準）", "情境 B（平平）"), ("情境 C（保守）", "情境 C（不如預期）"), ("情境 C（下檔）", "情境 C（不如預期）"),
           ("情境 A 隱含價（樂觀）", "情境 A 隱含價（順利）"), ("情境 B 隱含價（基準）", "情境 B 隱含價（平平）"), ("情境 C 隱含價（下檔）", "情境 C 隱含價（不如預期）"))
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("="):
                    w = v
                    for a, b in MAP:
                        w = w.replace(a, b)
                    if w != v:
                        c.value = w
    # (24) 上檔空間級距列名：六級僅供分組，投資建議另依五級規則
    for r in range(1, Vw.max_row + 1):
        if Vw.cell(r, 1).value == "上檔空間級距":
            Vw.cell(r, 1).value = "上檔空間級距（六級，供分組）"
            e = str(Vw.cell(r, 5).value or "")
            Vw.cell(r, 5).value = (e + "；" if e else "") + "六級只供分組比較；投資建議依下一列的五級固定規則"
    # (16) Sens「g ±1pt」：g 受 rf 上限時列名改「g −1pt／上限 rf」（公式判斷）
    for r in range(1, SEw.max_row + 1):
        if SEw.cell(r, 1).value == "永續成長率 g ±1pt":
            SEw.cell(r, 1).value = '=IF(g_tv_hi<g_tv+0.01-0.000001,"永續成長率 g −1pt／上限 rf","永續成長率 g ±1pt")'
    # (14) stub 說明字串依實際口徑（模組預設字樣為「第一年 FCFF × stub_frac」）
    if stub_note:
        h = next(r for r in range(1, Vw.max_row + 1) if str(Vw.cell(r, 1).value or "").startswith("0.1 DCF"))
        r = next(r for r in range(h, Vw.max_row + 1) if str(Vw.cell(r, 1).value or "").startswith("＋Stub 期 FCFF"))
        Vw.cell(r, 5).value = '="資產負債表日 "&bs_date&"；stub_frac "&TEXT(stub_frac,"0.000")&" 年（' + stub_note + '）"'
    if stub_frac_note:
        r = next(r for r in range(1, INw.max_row + 1) if INw.cell(r, 1).value == "stub_frac")
        INw.cell(r, 9).value = stub_frac_note
    if stub_pt:
        old = '（情境 B 第一年 FCFF × "&TEXT(stub_frac,"0.000")&" 年）'
        hit = 0
        for r in range(1, RDw.max_row + 1):
            v = RDw.cell(r, 1).value
            if isinstance(v, str) and old in v:
                RDw.cell(r, 1).value = v.replace(old, stub_pt); hit += 1
        assert hit == 1, hit
    # 版本
    RDw["A1"] = RDw["A1"].value + "；" + a1_note
    RDw.cell(RDw.max_row + 2, 1, rd_note)
# 第三輪修正（獨立核對發現）：In「情境驅動對照」補前段 WFE 三情境列（本模型最主要的情境驅動；只引用 Units 頁設備市場三情境，純展示），情境權重列移到最後
from copy import copy as _cp
_I3 = wb["In"]
_rw3 = next(r for r in range(1, _I3.max_row + 1) if _I3.cell(r, 1).value == "scn_w")
_nw = 4
for cc in range(1, 10):
    src, dst = _I3.cell(_rw3, cc), _I3.cell(_rw3 + _nw, cc)
    dst.value = src.value; dst._style = _cp(src._style)
for k_, y_ in enumerate(range(2026, 2030)):
    r_ = _rw3 + k_; cy_ = chr(ord("C") + y_ - 2025)
    for cc in range(1, 10):
        _I3.cell(r_, cc)._style = _cp(_I3.cell(_rw3 + _nw, cc)._style)
    _I3.cell(r_, 1, f"scn_wfe_{y_ % 100}"); _I3.cell(r_, 2, f"前段 WFE（代工邏輯＋記憶體；設備市場三情境）CY{y_}"); _I3.cell(r_, 3, "US$B")
    for cc, s_ in zip((4, 5, 6), SC):
        _I3.cell(r_, cc, f"=Units!{cy_}{UR[(s_, 'L')]}+Units!{cy_}{UR[(s_, 'M')]}").number_format = "0.0"
    _I3.cell(r_, 7, "Derived"); _I3.cell(r_, 8, "Units 頁：markets/equipment 三情境的代工邏輯＋記憶體 WFE 小計（純展示；系統營收＝製程控制強度 × 加權前段 WFE）")
    _I3.cell(r_, 9, None)# 第三輪修正（2026-11-06）：現價統一更新（fixes_round3 第 35 項；主控決定）——本公司現價與 Comparable 同業前瞻 P/E 一律用 2026-11-04 收盤（Yahoo Finance 日線收盤 2026-11-04）。
#   同業前瞻 P/E 換算到 2026-11-04：原 P/E × 11-04 收盤 ÷ 原報價日收盤（前瞻 EPS 共識視為不變；原報價日收盤取自 stockanalysis 歷史股價／statistics 頁）。
R3_PX_NEW = {"ASML.AS": 1337.6, "ASML": 1496.3, "AMAT": 369.4, "LRCX": 221.43, "KLAC": 149.79, "8035.T": 9663, "ASM.AS": 683.6, "6857.T": 34570, "TER": 250.99,
             "6146.T": 47370, "7729.T": 14920, "6920.T": 39130, "6525.T": 6400, "COHU": 43.98, "FORM": 85.24}
R3_PX_OLD = {   # 原報價日 → 收盤（同業 P/E 的價格日期）
    ("AMAT", "2026-10-23"): 377.24, ("LRCX", "2026-10-23"): 230.1, ("KLAC", "2026-10-23"): 145.61, ("ASML", "2026-10-23"): 1571.80,
    ("8035.T", "2026-10-21"): 10160, ("ASM.AS", "2026-10-23"): 699.80, ("6920.T", "2026-10-23"): 38040, ("TER", "2026-10-23"): 308.85,
    ("6857.T", "2026-10-21"): 34440, ("6146.T", "2026-10-23"): 49290, ("6525.T", "2026-10-02"): 10050, ("7729.T", "2026-09-29"): 18405,
    ("COHU", "2026-10-23"): 47.96, ("FORM", "2026-10-23"): 98.51}
def _r3_px(wb, tk, px_label, px_src):
    INw, Vw = wb["In"], wb["Val"]
    r = next(r for r in range(1, INw.max_row + 1) if INw.cell(r, 1).value == "px")
    old_px = INw.cell(r, 5).value
    INw.cell(r, 2).value = re.sub(r"\d{4}-\d{2}-\d{2} 收盤", "2026-11-04 收盤", str(INw.cell(r, 2).value)).replace("stockanalysis ", "")
    INw.cell(r, 5).value = R3_PX_NEW[tk]
    INw.cell(r, 8).value = f"Yahoo Finance 日線收盤 2026-11-04：{R3_PX_NEW[tk]:,}（{tk}）"
    INw.cell(r, 9).value = f"第三輪修正（2026-11-06）：現價統一為 2026-11-04 收盤（原 {old_px:,}，{px_src}）"
    for rr in range(1, Vw.max_row + 1):
        a = Vw.cell(rr, 1).value
        if isinstance(a, str) and a.startswith("現價（") and not str(Vw.cell(rr, 2).value or "").startswith("現價"):
            _old_lab = a
            Vw.cell(rr, 1).value = px_label
    for rr in range(1, Vw.max_row + 1):   # 其他說明格引用的現價列名
        for cc in range(2, 8):
            e = Vw.cell(rr, cc).value
            if isinstance(e, str) and not e.startswith("=") and _old_lab in e:
                Vw.cell(rr, cc).value = e.replace(_old_lab, px_label)
    for rr in range(1, Vw.max_row + 1):   # 說明欄的現價日期
        for cc in range(2, 8):
            e = Vw.cell(rr, cc).value
            if isinstance(e, str) and not e.startswith("=") and ("px" in e or "現價" in e) and re.search(r"\d{4}-\d{2}-\d{2} 收盤", e):
                Vw.cell(rr, cc).value = re.sub(r"(stockanalysis )?\d{4}-\d{2}-\d{2} 收盤", "2026-11-04 收盤", e)
    h = next(r for r in range(1, Vw.max_row + 1) if str(Vw.cell(r, 1).value or "").startswith("0.2 Comparable"))
    out = []
    for rr in range(h + 1, h + 15):
        t, pe, d = Vw.cell(rr, 2).value, Vw.cell(rr, 3).value, Vw.cell(rr, 4).value
        if isinstance(t, str) and isinstance(pe, (int, float)) and isinstance(d, str) and re.match(r"\d{4}-\d{2}-\d{2}$", d):
            p0 = R3_PX_OLD[(t, d)]; p1 = R3_PX_NEW[t]
            Vw.cell(rr, 3).value = round(pe * p1 / p0, 2)
            Vw.cell(rr, 4).value = "2026-11-04"
            Vw.cell(rr, 5).value = (str(Vw.cell(rr, 5).value or "") + f"；第三輪修正：換算到 2026-11-04＝原 {pe}（{d}）× 11-04 收盤 {p1:,} ÷ {d} 收盤 {p0:,}"
                                    f"（前瞻 EPS 視為不變；Yahoo Finance 日線收盤 2026-11-04；原日期收盤 stockanalysis）")
            out.append((t, pe, Vw.cell(rr, 3).value))
    assert out
    return out
import re
R3_PE = _r3_px(wb, "KLAC", "現價（2026-11-04 收盤）", "2026-10-23 收盤，stockanalysis")

_r3_fix(wb, "v0.4.3",
        "v0.4.3（2026-11-06）第三輪修正：情境驅動對照補前段 WFE 三情境列、README 舊重點觀察列刪除、情境名稱統一、列名與 HTML 展示調整、現價統一為 2026-11-04 收盤",
        "v0.4.3（2026-11-06）：第三輪修正（合併前）。(1) In「情境驅動對照」補前段 WFE（代工邏輯＋記憶體）CY2026–2029 三情境列（設備市場模型；本模型最主要的情境驅動），情境權重列移到最後；HTML 情境驅動表同步。"
        "(2) Stub 期：資產負債表日 2026-09-30 → 評價日 2026-12-31 落在 FY2027（7–6 月）內，維持「含 stub 期的會計年度（FY2027）FCFF × stub_frac」（fixes_round3 第 2 項的口徑；Q 頁 FQ2-27 指引沒有營業費用，不另推季度 FCFF），說明字串改寫實際口徑。"
        "12 月發放的股利：2026-11-05 前未宣告（最近一次 2026-08-06 宣告、09-01 發放），stub 期不扣。"
        "(3) README 舊「重點觀察」列（v0.3 前舊法數字）刪除。(4) 情境名稱統一為 情境 A（順利）／B（平平）／C（不如預期）。(5) Val「上檔空間級距」列名改「上檔空間級距（六級，供分組）」，註明投資建議依五級規則。"
        "(6) Sens「永續成長率 g ±1pt」在 g＋1pt 受 rf 上限時列名改「g −1pt／上限 rf」（公式判斷）。(7) 現價與 Comparable 同業前瞻 P/E 統一為 2026-11-04 收盤（同業 P/E＝原 P/E × 11-04 收盤 ÷ 原報價日收盤）。"
        "(8) HTML：DCF 彙總主表直接顯示「＋Stub 期 FCFF」「淨現金滾到評價日」；折現率主表直接顯示股權成本、稅後債務成本、WACC、永續成長率、RONIC，rf／β／ERP／CRP 與權益占資本摺疊；橫幅標示幣別（美元）。模組 tools/fcff_dcf.py v2 不變。",
        stub_note="含 stub 期的會計年度 FY2027 FCFF × stub_frac",
        stub_frac_note="stub FCFF＝含 stub 期的會計年度（FY2027，2026-07～2027-06）FCFF × 本值（年度比例）",
        stub_pt='（含 stub 期的會計年度 FY2027 FCFF × "&TEXT(stub_frac,"0.000")&" 年）')
FD.yellow(IN)   # v0.4：M2 節判斷參數黃底
wb.save(sys.argv[1])
