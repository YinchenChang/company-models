# Applied Materials（AMAT）收支與估值模型 v0.2.1（AI 半導體 Project 步驟 5：設備與材料——第三家公司）
# 用法：python3 amat/build_amat.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度約 11 月到隔年 10 月底（52／53 週，結束於 10 月最後一個星期日）：FY2026＝2025-10-27～2026-10-25。
# 金額：百萬美元（US$M）；EPS 為美元，採公司 non-GAAP 稀釋 EPS（指引與共識同口徑），GAAP 另列。
# 結構比照 advantest/build_advantest.py：README、In、Q、Units、PL、Val、Recon、Sens、Sources。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

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
SEG_LAB = {"la": "先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕", "lo": "其他代工與邏輯", "of": "其他前段",
           "dr": "DRAM（含 HBM）〔AI〕", "nd": "NAND"}
SEG = {}; _cnt = {k: 0 for k in SEG_LAB}
SEMI = {}; AIINC = None
for r in ws_.iter_rows():
    lab = r[0].value
    for k, l in SEG_LAB.items():
        if lab == l:
            s = SC[_cnt[k]]; _cnt[k] += 1
            SEG[(k, s)] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    for k, l in (("wfe", "SEMI 晶圓前段設備（WFE）"), ("fl", "SEMI 代工與邏輯設備"), ("dr", "SEMI DRAM 設備"), ("nd", "SEMI NAND 設備")):
        if lab == l:
            SEMI[k] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if AIINC is None and isinstance(lab, str) and lab.strip() == "AI 增量占總增量（相對 2025 累計）":
        AIINC = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}（情境 A）")
assert all(v == 3 for v in _cnt.values()) and len(SEMI) == 4 and AIINC, (_cnt, SEMI)
for k in SEG_LAB:   # 2025–2027 三情境相同（設備市場模型的設計）
    for j in range(3):
        assert abs(SEG[(k, "A")][0][j] - SEG[(k, "C")][0][j]) < 1e-9, k
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則：各情境見頂年度（markets/equipment）

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：百萬美元（US$M），市場為十億美元（US$B）。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司新聞稿與 SEC 8-K 附件以網頁擷取工具讀取（同一份文件的兩個轉載不算兩個獨立來源）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)

sec("A. 晶圓廠設備市場（日曆年，US$B；讀 markets/equipment 成品，三情境；唯讀，不自行重估）")
NM = {"la": "先進邏輯設備", "lo": "其他代工與邏輯設備", "of": "其他前段設備（類比、功率、光罩、廠務等）", "dr": "DRAM（含 HBM）設備", "nd": "NAND 設備"}
for s in SC:
    for k in SEG_LAB:
        for j, y in enumerate(range(2025, 2030)):
            inp(f"{k}{s}_{y % 100}", f"{NM[k]} CY{y}（{s}）", "US$B", None, SEG[(k, s)][0][j], None, "Derived", SEG[(k, s)][1] if j == 0 else "同上",
                "markets/equipment Mkt 第 5 節（公司資本支出 × 設備占比，2025 以 SEMI 實績校準）；2028–29 隨情境係數" if (j == 0 and k == "la" and s == "A") else "")
inp("fyw", "會計年度落在前一個日曆年的比例（FY(t)＝fyw×CY(t−1)＋(1−fyw)×CY(t)）", "%", 1 / 12, 2 / 12, 3 / 12, "Derived",
    "公司會計年度結束於 10 月最後一個星期日（FY2026：2025-10-27～2026-10-25），約 11–12 月兩個月落在前一年 → 2/12", "低 1/12、高 3/12（Assumed）：設備出貨與客戶資本支出約差一季，高值等於把市場往前移一季")
for k, l in (("wfe", "SEMI 晶圓前段設備（WFE）"), ("fl", "SEMI 代工與邏輯設備"), ("dr", "SEMI DRAM 設備"), ("nd", "SEMI NAND 設備")):
    inp(f"semi_{k}26", f"{l} CY2026（對帳用）", "US$B", None, SEMI[k][0][1], None, "Interested-party", SEMI[k][1], "SEMI 2026-07 年中預測（設備商協會，傾向樂觀）；只用於 Recon")
inp("ai_inc27", "設備市場模型：AI 相關設備占 2025→2027 增量（情境 A）", "%", None, AIINC[0][2], None, "Derived", AIINC[1], "先進邏輯＋DRAM／HBM＋先進封裝；只用於 Recon")
inp("co_ai27", "公司說法：先進代工邏輯、DRAM、先進封裝占 CY2026–27 晶圓廠設備成長約 80%", "%", None, 0.80, None, "Interested-party", "AMAT Q3 FY26 法說（2026-08-13；MarketBeat 法說重點）", "Nasdaq 法說重點（可能同一來源轉載）；只用於 Recon")

sec("B. 半導體系統（Semiconductor Systems）：FY2026 錨點 × 各終端市場設備指數 × 份額係數")
inp("q4_dr", "FY2026 第四季半導體系統中 DRAM 占比（公司未給）", "%", 0.26, 0.30, 0.34, "Assumed", "公司：DRAM 營收 2H CY2026「very significant increase」（Q3 FY26 法說，MarketBeat）；Q3 26%、Q1 34%", "低＝Q3 的 26%、高＝Q1 的 34%")
inp("q4_nd", "FY2026 第四季半導體系統中 NAND（Flash）占比", "%", 0.04, 0.06, 0.08, "Assumed", "Q1 4%、Q2 4%、Q3 7%（公司新聞稿）", "—")
inp("m_dr", "FY2026 半導體系統 DRAM 權重（前三季實績＋第四季假設，金額加權）", "%", "=(SUMPRODUCT(Q!G19:I19,Q!G4:I4)+q4_dr_lo*Q!J4)/an_ss", "=(SUMPRODUCT(Q!G19:I19,Q!G4:I4)+q4_dr*Q!J4)/an_ss", "=(SUMPRODUCT(Q!G19:I19,Q!G4:I4)+q4_dr_hi*Q!J4)/an_ss", "Derived", "Q 頁：各季營收 × 終端市場占比", "低／高對應 q4_dr 低／高")
inp("m_nd", "FY2026 半導體系統 NAND 權重", "%", None, "=(SUMPRODUCT(Q!G20:I20,Q!G4:I4)+q4_nd*Q!J4)/an_ss", None, "Derived", "同上", "—")
inp("m_fl", "FY2026 半導體系統 晶圓代工、邏輯與其他 權重（＝1 − DRAM − NAND）", "%", "=1-m_dr_lo-m_nd", "=1-m_dr-m_nd", "=1-m_dr_hi-m_nd", "Derived", "同上；公司「Foundry, logic and other」含 ICAPS（成熟製程、功率、類比等），故對應設備市場的先進邏輯＋其他代工邏輯＋其他前段", "低／高與 m_dr 對沖")
inp("sf", "份額係數（FY2027 起 AMAT 半導體系統營收 ÷ 依市場指數推得的營收）", "x", 0.95, 1.00, 1.05, "Assumed", "無：基準假設 AMAT 在三個終端市場的份額維持 FY2026 水準", "上方＝環繞式閘極、背面供電、先進封裝的份額提升；下方＝中國國產替代、競爭（Lam、TEL、KLA）")

sec("C. 出口管制與中國")
inp("aff_hit", "BIS「關係企業規則」恢復時的年營收影響", "US$M", 300, 600, 900, "Interested-party", "AMAT 8-K（2025-09-29）：該規則將使 FY2026 營收減少約 US$600M（Q4 FY25 約 US$110M）", "Reuters（Yahoo Finance 2025-11-14）：規則暫停後約 US$600M 營收可恢復；Skadden：暫停至 2026-11-09，除非 BIS 延長。低／高 Assumed（±50%）")
for s, v in (("A", 0), ("B", 0), ("C", 1)):
    inp(f"aff_{s}", f"情境 {s}：關係企業規則於 FY2027 起恢復（1＝是）", "0/1", None, v, None, "Assumed", "暫停期至 2026-11-09（Skadden 2025-11）；是否延長未知" if s == "A" else "同上", "A、B 假設延長；C 假設恢復（與設備市場下修同向的保守組合）")
inp("cn_sh", "中國占營收（FY2026 前三季，Derived）", "%", None, "=SUM(Q!G21:I21)/SUM(Q!G7:I7)", None, "Derived", "Q 頁：前三季中國營收 ÷ 營收（公司新聞稿；Q3 為 28%）", "公司：Q3 中國占半導體系統＋AGS 的 26%；預期 CY2026 中國營收增加（28nm 代工邏輯帶動）")
inp("cn_cut", "敏感度：中國營收下修比例（國產替代或新管制）", "%", 0.0, 0.0, 0.40, "Assumed", "無；基準 0（市場層的「其他代工邏輯」已含中國擴產的成長假設）", "Sens 用高值 20%、40%（只對半導體系統＋AGS）")

sec("D. AGS 服務與其他（含顯示器）：前一年 × 成長率")
for s, g in (("A", 0.15), ("B", 0.11), ("C", 0.07)):
    inp(f"g_ags_{s}", f"AGS FY2027–29 年成長（{s}）", "%", g - 0.04, g, g + 0.04, "Interested-party" if s == "A" else "Assumed",
        "AMAT Q3 FY26 法說：AGS 長期年成長 mid-teens；CY2026 >20%（MarketBeat、Nasdaq 法說重點）" if s == "A" else "無：裝機量與稼動隨設備市場放緩（B／C 為 Assumed）", "區間 ±4 個百分點（Assumed）")
for s, g in (("A", 0.0), ("B", -0.05), ("C", -0.10)):
    inp(f"g_oth_{s}", f"其他（主要為顯示器）FY2027–29 年成長（{s}）", "%", g - 0.10, g, g + 0.10, "Assumed", "無：顯示器設備隨面板廠投資起伏；Q4 FY26 指引約 US$510M 高於前三季（280–312）", "FY2026 為 1.4% 營收以下的小項；區間 ±10 個百分點")

sec("E. 毛利、營業費用、淨利、股數（兩點推導：FY2025 實績與 FY2026 錨點；non-GAAP；不寫死利益率）")
inp("cogs25", "營業成本 FY2025（營收 − non-GAAP 毛利）", "US$M", None, "=a25_rev-a25_gp", None, "Derived", "Q 頁", "—")
inp("cogs26", "營業成本 FY2026 錨點", "US$M", None, "=an_rev-an_gp", None, "Derived", "Q 頁（前三季實績＋第四季指引毛利率 50.4%）", "—")
inp("cv", "營業成本變動部分占營收（兩點斜率）", "%", None, "=(cogs26-cogs25)/(an_rev-a25_rev)", None, "Derived", "（FY2026 − FY2025 營業成本）÷ 營收差", "隱含毛利率隨營收上升而遞增、但趨緩")
inp("cF", "營業成本固定部分", "US$M", None, "=cogs25-cv*a25_rev", None, "Derived", "FY2025 營業成本 − 變動部分", "—")
inp("phic", "營業成本中變動部分占比（FY2025；Sens 用）", "%", 0.75, "=cv*a25_rev/cogs25", 0.95, "Derived", "基準＝兩點推導值；低／高 Assumed", "低＝固定成本較多（營收增加時毛利率升得更快）")
inp("opex25", "營業費用 FY2025（non-GAAP 毛利 − non-GAAP 營業利益）", "US$M", None, "=a25_gp-a25_oi", None, "Derived", "Q 頁", "—")
inp("opex26", "營業費用 FY2026 錨點（前三季＋第四季指引 US$1,580M）", "US$M", None, "=an_opx", None, "Derived", "Q 頁", "公司：FY2027 Q1 為 14 週，營業費用增幅大於往常（Q3 FY26 法說）")
inp("ov", "營業費用變動部分占營收（兩點斜率）", "%", None, "=(opex26-opex25)/(an_rev-a25_rev)", None, "Derived", "（FY2026 − FY2025 營業費用）÷ 營收差", "—")
inp("oF", "營業費用固定部分", "US$M", None, "=opex25-ov*a25_rev", None, "Derived", "FY2025 營業費用 − 變動部分", "—")
inp("phi", "營業費用中變動部分占比（FY2025；Sens 用）", "%", 0.35, "=ov*a25_rev/opex25", 0.75, "Derived", "基準＝兩點推導值；低／高 Assumed", "公司：產能 2028 年前倍增、Q3 增聘逾 1,500 人（費用前置的風險）")
inp("nioi", "non-GAAP 淨利 ÷ non-GAAP 營業利益", "x", 0.85, "=a25_ni/a25_oi", "=an_ni/an_oi", "Derived", "基準＝FY2025 實績 0.888；高＝FY2026 錨點隱含（含公司內部無形資產移轉的稅務利益，Q1／Q2 指引各 US$0.04／股）", "低 0.85（Assumed：全球最低稅負等稅率上升）")
inp("sh_q4", "FY2026 第四季稀釋股數（公司未給）", "百萬股", 793, 796, 800, "Assumed", "Q3 稀釋 800M；10-Q 封面流通股 793.6M（2026-07-26）；Q3 回購 US$440M", "只影響 FY2026 錨點淨利（＝第四季指引 EPS × 股數）")
inp("bb", "FY2027 起稀釋股數年減幅（淨回購）", "%", 0.0, 0.01, 0.025, "Assumed", "FY2025 稀釋股數 808M → FY2026 前三季 799M（約 −1.1%）；回購授權餘額 US$12.8B（Q3 FY26 法說）", "—")

sec("F. 公司指引、共識、股價、循環位置")
inp("g_rev_q4", "指引營收 FY2026 第四季（中點，±US$500M）", "US$M", 9750, 10250, 10750, "Interested-party", "AMAT Q3 FY26 新聞稿（2026-08-13）", "MarketBeat 法說重點；低／高＝公司指引區間（未進 Sens：錨點只影響 FY2026 水準）")
inp("g_eps_q4", "指引 non-GAAP 稀釋 EPS FY2026 第四季（中點，±US$0.20）", "US$", 3.82, 4.02, 4.22, "Interested-party", "同上", "同上（低／高＝指引區間，未進 Sens）")
inp("g_ss_q4", "指引 半導體系統 FY2026 第四季（約，年增 62%）", "US$M", None, 7900, None, "Interested-party", "AMAT Q3 FY26 法說（MarketBeat）", "Nasdaq 法說重點（可能同源）")
inp("g_ags_q4", "指引 AGS FY2026 第四季（約，年增 22%）", "US$M", None, 1840, None, "Interested-party", "同上", "同上")
inp("g_oth_q4", "指引 其他（主要為顯示器）FY2026 第四季（約）", "US$M", None, 510, None, "Interested-party", "同上", "同上")
inp("g_gm_q4", "指引 non-GAAP 毛利率 FY2026 第四季", "%", None, 0.504, None, "Interested-party", "同上", "同上")
inp("g_opx_q4", "指引 non-GAAP 營業費用 FY2026 第四季", "US$M", None, 1580, None, "Interested-party", "同上", "同上")
inp("q4g_ss", "第四季半導體系統指引年增（用於反推 FY2025 Q4 重編後數字）", "%", None, 0.62, None, "Interested-party", "同上（約 62%）", "四捨五入到整數百分比：反推值誤差約 ±0.3%")
inp("q4g_ags", "第四季 AGS 指引年增", "%", None, 0.22, None, "Interested-party", "同上（約 22%）", "同上")
inp("cons_rev26", "共識營收 FY2026", "US$M", None, 34260, None, "Interested-party", "stockanalysis.com（資料更新 2026-10-16；34 位分析師）", "第二來源：tickflow 31,430（明顯過時，低於前三季實績＋指引，未採）")
inp("cons_eps26", "共識 non-GAAP EPS FY2026", "US$", None, 12.79, None, "Interested-party", "同上", "同上：tickflow 11.05（過時，未採）")
inp("cons_rev27", "共識營收 FY2027", "US$M", None, 46370, None, "Interested-party", "stockanalysis.com（Next Year 摘要框）", "第二來源：未找到（tickflow 37,670 過時，未採）")
inp("cons_eps27", "共識 non-GAAP EPS FY2027", "US$", None, 18.49, None, "Interested-party", "同上", "第二來源：未找到。FY2028 共識：未找到（stockanalysis 需付費）")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；AMAT 的 FY(t) 有 10/12 落在日曆年 t，直接對照 t（不平移）" if s == "A" else "同上")
inp("px", "股價（stockanalysis 2026-10-23 收盤）", "US$", None, 377.24, None, "Interested-party", "stockanalysis.com：2026-10-23 收盤 377.24、市值 299.38B、流通股 793.60M", "statmuse 最新價 377.24（未標日期）")
inp("sa_fpe", "遠期 P/E（stockanalysis）", "x", None, 21.39, None, "Interested-party", "stockanalysis.com（2026-10-23）", "自算：現價 ÷ 共識 FY2027 EPS，見 Val")

sec("G. 目標價情境權重（Val 第 0 節用；Andy 可改）")
for s_, w_, nm_ in (("A", 0.3, "樂觀"), ("B", 0.5, "基準"), ("C", 0.2, "下檔")):
    inp(f"tp_w{s_}", f"目標價情境權重：情境 {s_}（{nm_}）", "權重", None, w_, None, "Assumed",
        "Andy 可改（2026-10-25 目標價規格預設 A／B／C＝0.3／0.5／0.2）", "三個權重合計應為 1；目標價＝Σ 權重 × 情境隱含價（Val 第 0 節）")

r = 4
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
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 64, 10, 10, 11, 10, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Q：實績（FY2025 四季重編、FY2026 前三季）與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（百萬美元；EPS 為美元）。會計年度約 11 月～隔年 10 月：FY2026＝2025-10-27～2026-10-25。藍字＝輸入；黑字＝公式。損益為 non-GAAP（公司揭露），GAAP EPS 另列。分部自 FY2026 起重編（200mm 設備移入半導體系統、公司費用全數分攤），FY2025 Q1–Q3 用 FY2026 新聞稿的重編比較數"
Q["A2"] = "FY2026 第四季（2026-07-27～2026-10-25）已結束、尚未公布：預定 2026-11-12 公布（stockanalysis）。本版 FY2026＝前三季實績＋第四季公司指引（錨點）。公布後更新 J 欄"
hdr(Q, 3, ["項目", "FY25 Q1（11–1月）", "FY25 Q2（2–4月）", "FY25 Q3（5–7月）", "FY25 Q4（8–10月）", "FY2025 全年", "FY26 Q1", "FY26 Q2", "FY26 Q3", "FY26 Q4（指引）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGHIJK", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 12, tag); Q.cell(r, 13, s1); Q.cell(r, 14, s2)
R1 = "AMAT 各季新聞稿（ir.appliedmaterials.com；SEC 8-K 附件 99.1）"
qrow(4, "ss", "半導體系統（重編後）", [5597, 5401, 5564, "=J4/(1+q4g_ss)", "=SUM(B4:E4)", 5141, 5965, 7040, "=g_ss_q4", "=SUM(G4:J4)"], "Verified／Derived",
     R1 + "：FY25 Q1–Q3 為 FY26 各季新聞稿的重編比較數；FY25 Q4 由 FY26 Q4 指引 ÷ (1＋62%) 反推（Derived）", "FY2025 舊口徑 20,798（FY25 Q4 新聞稿）；重編差額主要為 200mm 設備由 AGS 移入")
qrow(5, "ags", "AGS 服務（重編後）", [1353, 1420, 1463, "=J5/(1+q4g_ags)", "=SUM(B5:E5)", 1559, 1665, 1781, "=g_ags_q4", "=SUM(G5:J5)"], "Verified／Derived",
     R1 + "：同上；FY25 Q4 由指引 ÷ (1＋22%) 反推", "FY2025 舊口徑 6,385")
qrow(6, "oth", "其他（含顯示器）", [216, 279, 275, "=E7-E4-E5", "=SUM(B6:E6)", 312, 280, 294, "=g_oth_q4", "=SUM(G6:J6)"], "Verified／Derived",
     R1 + "；FY25 Q4＝營收 − 半導體系統 − AGS（Derived）", "FY2025 舊口徑顯示器 1,060（另含於公司與其他）")
qrow(7, "rev", "營收", [7166, 7100, 7302, 6800, 28368, 7012, 7910, 9115, "=g_rev_q4", "=SUM(G7:J7)"], "Verified",
     R1 + "：FY2025 28,368；FY26 前三季 24,037", "stockanalysis 季報表（逐季相同）")
qrow(8, "chk1", "檢查：分部合計 − 營收（應為 0）", [f"={c}4+{c}5+{c}6-{c}7" for c in "BCDEFGHIJK"], "—", "—", "FY25 Q1–Q3 與 FY26 各季：公司分部表加總＝營收", "0")
qrow(9, "gmr", "non-GAAP 毛利率", [0.489, 0.492, 0.489, 0.481, "=F10/F7", 0.491, 0.500, 0.504, "=g_gm_q4", "=K10/K7"], "Verified",
     R1 + "：FY2025 全年 48.8%（四季加權 48.8%，見 F 欄）", "GAAP 毛利率：stockanalysis 季報表毛利 ÷ 營收（差 0.1 個百分點內）", "0.0%")
qrow(10, "gp", "non-GAAP 毛利", [f"={c}7*{c}9" for c in "BCDE"] + ["=SUM(B10:E10)"] + [f"={c}7*{c}9" for c in "GHIJ"] + ["=SUM(G10:J10)"], "Derived", "營收 × 毛利率（毛利率為公司公布的一位小數）", "—")
qrow(11, "oi", "non-GAAP 營業利益", [2190, 2180, 2245, 1947, "=SUM(B11:E11)", 2107, 2536, 3096, "=J10-J12", "=SUM(G11:J11)"], "Verified／Derived",
     R1 + "：FY2025 8,562；FY26 Q4＝指引毛利 − 指引營業費用", "GAAP 營業利益（stockanalysis）：FY26 Q1 2,096 不含 BIS 和解金 253")
qrow(12, "opx", "non-GAAP 營業費用（毛利 − 營業利益）", [f"={c}10-{c}11" for c in "BCDEFGHI"] + ["=g_opx_q4", "=SUM(G12:J12)"], "Derived", "—", "FY26 Q4 為公司指引約 1,580")
qrow(13, "ni", "non-GAAP 淨利", [1946, 1940, 1989, 1732, "=SUM(B13:E13)", 1899, 2286, 2795, "=J15*J14", "=SUM(G13:J13)"], "Verified／Derived",
     R1 + "：FY2025 7,607；FY26 Q4＝指引 EPS × 假設股數", "—")
qrow(14, "sh", "稀釋加權股數（百萬股）", [819, 812, 802, 798, 808, 799, 799, 800, "=sh_q4", "=K13/K15"], "Verified／Derived", R1, "FY2026 錨點＝淨利 ÷ EPS（Derived）")
qrow(15, "eps", "non-GAAP 稀釋 EPS（美元）", [2.38, 2.39, 2.48, 2.17, 9.42, 2.38, 2.86, 3.50, "=g_eps_q4", "=SUM(G15:J15)"], "Verified",
     R1 + "：FY2025 9.42（四季合計 9.42）", "共識 FY2026 12.79（stockanalysis）", "0.00")
qrow(16, "geps", "GAAP 稀釋 EPS（美元；參考）", [1.45, 2.63, 2.22, 2.38, 8.66, 2.54, 3.51, 3.17, None, None], "Verified",
     R1 + "：FY2025 8.66（FY2024 8.61）", "stockanalysis 季報表（逐季相同）；FY26 Q1 含 BIS 和解金 253（每股 0.32）、Q2 含投資評價益", "0.00")
qrow(17, "gni", "GAAP 淨利（參考）", [1185, 2137, 1779, 1897, 6998, 2026, 2806, 2538, None, None], "Verified", R1, "stockanalysis 季報表")
qrow(18, "mfl", "半導體系統：晶圓代工、邏輯與其他 占比", [0.69, 0.66, 0.69, 0.65, None, 0.62, 0.67, 0.67, "=1-J19-J20", "=SUMPRODUCT(G18:J18,G4:J4)/K4"], "Verified／Assumed",
     R1 + "（FY25 Q4 為舊口徑：FY25 Q4 新聞稿；FY2025 全年 67%）", "FY26 Q4 為 Assumed（In 頁 q4_dr、q4_nd）", "0%")
qrow(19, "mdr", "半導體系統：DRAM 占比", [0.27, 0.27, 0.22, 0.29, None, 0.34, 0.29, 0.26, "=q4_dr", "=SUMPRODUCT(G19:J19,G4:J4)/K4"], "Verified／Assumed", R1 + "（FY2025 全年 26%）", "Q3 FY26 DRAM（含 HBM 封裝）年增 52%（法說）", "0%")
qrow(20, "mnd", "半導體系統：Flash（NAND）占比", [0.04, 0.07, 0.09, 0.06, None, 0.04, 0.04, 0.07, "=q4_nd", "=SUMPRODUCT(G20:J20,G4:J4)/K4"], "Verified／Assumed", R1 + "（FY2025 全年 7%）", "—", "0%")
qrow(21, "cn", "中國營收", [2243, 1774, 2548, 1964, "=SUM(B21:E21)", 2095, 2087, 2506, None, None], "Verified", R1 + "：FY2025 8,529（30%；FY2024 10,117、37%）", "Reuters：中國占比降至約 25%（Yahoo Finance 2025-11-14）")
qrow(22, "cnp", "中國占營收", [f"={c}21/{c}7" for c in "BCDEFGHI"], "Derived", "公司：Q1 30%、Q2 27%、Q3 28%", "—", "0.0%")
qrow(23, "ck2", "檢查：四季合計 − 全年公布值（營收、毛利率、營業利益、淨利、EPS、中國 依序在 B–G 欄）", ["=SUM(B7:E7)-F7", "=F9-0.488", "=F11-8562", "=F13-7607", "=SUM(B15:E15)-F15", "=F21-8529"], "—", "—", "毛利率差為四捨五入（公司公布 48.8%）", "0.000")
for nm, rr in (("ss", 4), ("ags", 5), ("oth", 6), ("rev", 7), ("gp", 10), ("oi", 11), ("ni", 13), ("sh", 14), ("eps", 15)):
    defname(f"a25_{nm}", f"Q!$F${rr}")
for nm, rr in (("ss", 4), ("ags", 5), ("oth", 6), ("rev", 7), ("gp", 10), ("oi", 11), ("opx", 12), ("ni", 13), ("sh", 14), ("eps", 15)):
    defname(f"an_{nm}", f"Q!$K${rr}")
Q.cell(26, 1, "FY2026 錨點的組成（前三季實績＋第四季指引）").font = BOLD
hdr(Q, 27, ["項目", "前三季實績", "第四季指引", "FY2026 錨點", "FY2025", "年增", "第四季 × 4（年化）"])
for i, (lab, rr, fmt) in enumerate((("營收", 7, "#,##0"), ("半導體系統", 4, "#,##0"), ("AGS", 5, "#,##0"), ("其他", 6, "#,##0"), ("non-GAAP 營業利益", 11, "#,##0"), ("non-GAAP EPS", 15, "0.00")), 28):
    Q.cell(i, 1, lab)
    Q.cell(i, 2, f"=SUM(G{rr}:I{rr})").number_format = fmt; Q.cell(i, 3, f"=J{rr}").number_format = fmt; Q.cell(i, 4, f"=K{rr}").number_format = fmt
    Q.cell(i, 5, f"=F{rr}").number_format = fmt; Q.cell(i, 6, f"=D{i}/E{i}-1").number_format = "0.0%"; Q.cell(i, 7, f"=C{i}*4").number_format = fmt
Q.cell(35, 1, "non-GAAP 營業利益率：FY2025"); Q.cell(35, 2, "=a25_oi/a25_rev").number_format = "0.0%"
Q.cell(36, 1, "non-GAAP 營業利益率：FY2026 錨點"); Q.cell(36, 2, "=an_oi/an_rev").number_format = "0.0%"
Q.cell(37, 1, "共識 FY2026 營收、EPS（stockanalysis；只作對照）"); Q.cell(37, 2, "=cons_rev26").number_format = "#,##0"; Q.cell(37, 3, "=cons_eps26").number_format = "0.00"
Q.column_dimensions["A"].width = 56
for c in "BCDEFGHIJK": Q.column_dimensions[c].width = 13
Q.column_dimensions["L"].width = 18; Q.column_dimensions["M"].width = 90; Q.column_dimensions["N"].width = 70
Q.freeze_panes = "B4"

# ---------------- Units：晶圓廠設備市場（日曆年 → 會計年度）與 AMAT 隱含份額 ----------------
YRS = range(2025, 2030)
col = {y: chr(ord("C") + y - 2025) for y in YRS}   # C＝2025 … G＝2029
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 晶圓廠設備市場（US$B；markets/equipment 成品，三情境）：日曆年 → AMAT 會計年度（FY(t)＝fyw×CY(t−1)＋(1−fyw)×CY(t)），以及 AMAT 半導體系統占市場的隱含份額（只作檢驗，不校準）"
hdr(UN, 3, ["項目", "單位"] + [str(y) for y in YRS] + ["說明"])
UR = {}; ur = 4
for s in SC:
    UN.cell(ur, 1, f"情境 {s}：日曆年（CY）").font = BOLD; ur += 1
    for key, lab, f in (("fl", "晶圓代工、邏輯與其他（先進邏輯＋其他代工邏輯＋其他前段）", lambda y: f"=la{s}_{y%100}+lo{s}_{y%100}+of{s}_{y%100}"),
                        ("dr", "DRAM（含 HBM）", lambda y: f"=dr{s}_{y%100}"), ("nd", "NAND", lambda y: f"=nd{s}_{y%100}")):
        UN.cell(ur, 1, "  " + lab); UN.cell(ur, 2, "US$B")
        for y in YRS: UN[f"{col[y]}{ur}"] = f(y); UN[f"{col[y]}{ur}"].number_format = "0.0"
        UR[(s, key, "cy")] = ur; ur += 1
    UN.cell(ur, 1, "  合計（晶圓前段）"); UN.cell(ur, 2, "US$B")
    for y in YRS: UN[f"{col[y]}{ur}"] = f"=SUM({col[y]}{ur-3}:{col[y]}{ur-1})"; UN[f"{col[y]}{ur}"].number_format = "0.0"
    UR[(s, "tot", "cy")] = ur; ur += 1
    UN.cell(ur, 1, f"情境 {s}：會計年度（FY；FY2025 需 CY2024，市場層沒有，留空）").font = BOLD; ur += 1
    for key, lab in (("fl", "晶圓代工、邏輯與其他"), ("dr", "DRAM（含 HBM）"), ("nd", "NAND"), ("tot", "合計（晶圓前段）")):
        UN.cell(ur, 1, "  " + lab); UN.cell(ur, 2, "US$B")
        for y in range(2026, 2030):
            UN[f"{col[y]}{ur}"] = f"=fyw*{col[y-1]}{UR[(s, key, 'cy')]}+(1-fyw)*{col[y]}{UR[(s, key, 'cy')]}"; UN[f"{col[y]}{ur}"].number_format = "0.0"
        UR[(s, key, "fy")] = ur; ur += 1
    ur += 1
UN.cell(ur, 1, "AMAT 半導體系統 ÷ 會計年度晶圓前段市場（隱含份額；情境 A；FY2027 起份額係數 sf 決定變化）").font = BOLD; ur += 1
UN_SHARE_ROW = ur   # 由 PL 填（PL 列號在後面才知道）
ur += 2
UN.cell(ur, 1, "說明：AMAT「Foundry, logic and other」含 ICAPS（成熟製程、功率、類比、物聯網等），故以設備市場的先進邏輯＋其他代工邏輯＋其他前段為對應；設備市場模型 2026 年以後高於 SEMI 約 7–10%（markets/equipment Recon），本模型只用各終端市場的成長指數，水準由 AMAT FY2026 錨點決定。")
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 10; UN.column_dimensions["H"].width = 60
for c in "CDEFG": UN.column_dimensions[c].width = 11

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYC = col   # FY 欄位同 C＝FY2025 … G＝FY2029
DEFP = dict(sf="sf", fyw="fyw", mdr="m_dr", mfl="m_fl", mnd="m_nd", cF="cF", cv="cv", oF="oF", ov="ov", nioi="nioi", bb="bb",
            aff=None, affv="aff_hit", cn="0", gags=None, goth=None, scn="A")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    aff = P["aff"] or f"aff_{s}"; gags = P["gags"] or f"g_ags_{s}"; goth = P["goth"] or f"g_oth_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "2025", "2026", "2027", "2028", "2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for c_, f in fs.items():
            ws[f"{c_}{r}"] = f; ws[f"{c_}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    ws.cell(r, 1, "A. 市場（日曆年 CY，US$B；讀 In 頁的設備市場成品）").font = BOLD; r += 1
    row("fl_c", "晶圓代工、邏輯與其他 CY", "US$B", {col[y]: f"=la{s}_{y%100}+lo{s}_{y%100}+of{s}_{y%100}" for y in YRS}, "0.0")
    row("dr_c", "DRAM（含 HBM）CY", "US$B", {col[y]: f"=dr{s}_{y%100}" for y in YRS}, "0.0")
    row("nd_c", "NAND CY", "US$B", {col[y]: f"=nd{s}_{y%100}" for y in YRS}, "0.0")
    ws.cell(r, 1, "B. 會計年度（FY；FY(t)＝fyw×CY(t−1)＋(1−fyw)×CY(t)）").font = BOLD; r += 1
    for k in ("fl", "dr", "nd"):
        row(f"{k}_f", {"fl": "晶圓代工、邏輯與其他 FY", "dr": "DRAM（含 HBM）FY", "nd": "NAND FY"}[k], "US$B",
            {col[y]: f"={P['fyw']}*{col[y-1]}{R[k+'_c']}+(1-{P['fyw']})*{col[y]}{R[k+'_c']}" for y in range(2026, 2030)}, "0.0")
    row("idx", "半導體系統需求指數（FY2026＝1；以 FY2026 終端市場權重加權）", "x",
        {col[y]: f"={P['mfl']}*{col[y]}{R['fl_f']}/$D{R['fl_f']}+{P['mdr']}*{col[y]}{R['dr_f']}/$D{R['dr_f']}+{P['mnd']}*{col[y]}{R['nd_f']}/$D{R['nd_f']}" for y in range(2026, 2030)},
        "0.000", "Σ 權重 × FY 市場 ÷ FY2026 市場")
    ws.cell(r, 1, "C. 損益（US$M；FY2025 實績、FY2026 錨點＝前三季實績＋第四季指引、FY2027 起模型）").font = BOLD; r += 1
    row("ss0", "半導體系統：需求推導（調整前）", "US$M", {"C": "=a25_ss", "D": "=an_ss", **{col[y]: f"=$D{r}*{col[y]}{R['idx']}*{P['sf']}" for y in range(2027, 2030)}}, note="FY2027 起＝FY2026 錨點 × 需求指數 × 份額係數")
    row("ags0", "AGS：成長路徑（調整前）", "US$M", {"C": "=a25_ags", "D": "=an_ags", **{col[y]: f"={col[y-1]}{r}*(1+{gags})" for y in range(2027, 2030)}})
    row("adj", "出口管制與中國調整（依半導體系統、AGS 營收比例分攤到兩分部）", "US$M", {"C": 0, "D": 0, **{col[y]: f"=-{P['affv']}*{aff}-{P['cn']}*cn_sh*({col[y]}{R['ss0']}+{col[y]}{R['ags0']})" for y in range(2027, 2030)}},
        note="關係企業規則恢復（情境 C 或 Sens）；中國下修只在 Sens")
    row("ss", "半導體系統營收", "US$M", {col[y]: f"={col[y]}{R['ss0']}+{col[y]}{R['adj']}*{col[y]}{R['ss0']}/({col[y]}{R['ss0']}+{col[y]}{R['ags0']})" for y in YRS}, note="官方分部 Semiconductor Systems")
    row("ags", "AGS 服務營收", "US$M", {col[y]: f"={col[y]}{R['ags0']}+{col[y]}{R['adj']}*{col[y]}{R['ags0']}/({col[y]}{R['ss0']}+{col[y]}{R['ags0']})" for y in YRS}, note="官方分部 Applied Global Services")
    row("oth", "其他（含顯示器）營收", "US$M", {"C": "=a25_oth", "D": "=an_oth", **{col[y]: f"={col[y-1]}{r}*(1+{goth})" for y in range(2027, 2030)}}, note="官方：Corporate and Other（含顯示器）")
    row("rev", "營收", "US$M", {col[y]: f"={col[y]}{R['ss']}+{col[y]}{R['ags']}+{col[y]}{R['oth']}" for y in YRS})
    row("chk", "檢查：分部合計 − 合併營收（FY2025 實績、FY2026 錨點）", "US$M", {"C": f"=C{R['ss']}+C{R['ags']}+C{R['oth']}-a25_rev", "D": f"=D{R['ss']}+D{R['ags']}+D{R['oth']}-an_rev"}, "#,##0.0")
    row("gm", "non-GAAP 毛利率", "%", {"C": "=a25_gp/a25_rev", "D": "=an_gp/an_rev", **{col[y]: f"=1-({P['cF']}+{P['cv']}*{col[y]}{R['rev']})/{col[y]}{R['rev']}" for y in range(2027, 2030)}}, "0.0%", "營業成本＝固定＋變動×營收（兩點推導）")
    row("gp", "non-GAAP 毛利", "US$M", {col[y]: f"={col[y]}{R['rev']}*{col[y]}{R['gm']}" for y in YRS})
    row("opx", "non-GAAP 營業費用", "US$M", {"C": "=opex25", "D": "=opex26", **{col[y]: f"={P['oF']}+{P['ov']}*{col[y]}{R['rev']}" for y in range(2027, 2030)}}, note="固定＋變動×營收（兩點推導）")
    row("oi", "non-GAAP 營業利益", "US$M", {col[y]: f"={col[y]}{R['gp']}-{col[y]}{R['opx']}" for y in YRS})
    row("oim", "non-GAAP 營業利益率", "%", {col[y]: f"={col[y]}{R['oi']}/{col[y]}{R['rev']}" for y in YRS}, "0.0%")
    row("ni", "non-GAAP 淨利", "US$M", {"C": "=a25_ni", "D": "=an_ni", **{col[y]: f"={col[y]}{R['oi']}*{P['nioi']}" for y in range(2027, 2030)}})
    row("sh", "稀釋股數", "百萬股", {"C": "=a25_sh", "D": "=an_sh", **{col[y]: f"=$D{r}*(1-{P['bb']})^{y-2026}" for y in range(2027, 2030)}}, "#,##0.0")
    row("eps", "non-GAAP 稀釋 EPS", "US$", {"C": "=a25_eps", "D": "=an_eps", **{col[y]: f"={col[y]}{R['ni']}/{col[y]}{R['sh']}" for y in range(2027, 2030)}}, "0.00",
        "FY2026＝前三季實績 EPS 合計＋第四季指引 EPS")
    row("shr", "半導體系統 ÷ FY 晶圓前段市場（隱含份額）", "%", {col[y]: f"={col[y]}{R['ss0']}/1000/({col[y]}{R['fl_f']}+{col[y]}{R['dr_f']}+{col[y]}{R['nd_f']})" for y in range(2026, 2030)}, "0.0%",
        "檢驗用；FY2027 起等於 FY2026 份額 × 份額係數 × 組合效果")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Applied Materials 損益（US$M；non-GAAP；FY2025 為實績，FY2026 為錨點（前三季實績＋第四季指引），FY2027–FY2029 為模型）。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 40
TITLES = {"A": "情境 A：設備市場續增（2028 +12.7%、2029 +5.9%；markets/equipment）",
          "B": "情境 B：設備市場 2028 持平、2029 −5%；AGS 成長放緩",
          "C": "情境 C：設備市場 2028 −20%、2029 持平；關係企業規則恢復；AGS、顯示器放緩"}
for s in SC:
    PB[s] = engine(PL, r0, TITLES[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$M", "#,##0"), ("ss", "  半導體系統", "US$M", "#,##0"), ("ags", "  AGS 服務", "US$M", "#,##0"),
                                ("gm", "non-GAAP 毛利率", "%", "0.0%"), ("oi", "non-GAAP 營業利益", "US$M", "#,##0"), ("oim", "non-GAAP 營業利益率", "%", "0.0%"),
                                ("ni", "non-GAAP 淨利", "US$M", "#,##0"), ("eps", "non-GAAP EPS", "US$", "0.00")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in YRS:
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "讀法：FY2026 三情境相同（錨點）；FY2027 的 1/6 落在 CY2026、5/6 在 CY2027，設備市場三情境在 CY2027 以前相同，所以 FY2027 只有情境 C 的出口管制調整造成差異；FY2028 起情境差異來自設備市場。").font = BOLD
PL_NOTE_ROW = sr
PL.column_dimensions["A"].width = 58; PL.column_dimensions["B"].width = 10; PL.column_dimensions["H"].width = 60
for c in "CDEFG": PL.column_dimensions[c].width = 14
# Units 頁的隱含份額列（引用 PL 情境 A 區塊）
UN.cell(UN_SHARE_ROW, 1, "  情境 A"); UN.cell(UN_SHARE_ROW, 2, "%")
for y in range(2026, 2030):
    UN[f"{col[y]}{UN_SHARE_ROW}"] = f"=PL!{col[y]}{PB['A']['shr']}"; UN[f"{col[y]}{UN_SHARE_ROW}"].number_format = "0.0%"
UN.cell(UN_SHARE_ROW, 8, "AMAT 半導體系統也含先進封裝等後段設備，故份額略高於純前段口徑")

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元；non-GAAP EPS）。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（見第 4 節）。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價（未調整股利）÷ 當年 10 月結束的會計年度 FY(t) 的 non-GAAP 稀釋 EPS（FY(t) 有 10/12 落在日曆年 t，與模型對 FY2027 EPS 套用倍數的口徑一致）；GAAP 口徑另列對照"
V.cell(4, 1, "1. 公司公布 EPS（會計年度，美元）— 兩個口徑、兩個來源").font = BOLD
hdr(V, 5, ["會計年度（10 月底結束）", "non-GAAP 稀釋 EPS（公司新聞稿）", "GAAP 稀釋 EPS（公司新聞稿）", "GAAP EPS（macrotrends）", "GAAP 兩來源差距", "標記", "說明"])
EPSD = [(2016, 1.75, 1.54, 1.54, "公司 FY2016 Q4 新聞稿（SEC 8-K 2016-11）"), (2017, 3.25, 3.17, 3.25, "公司 FY2018 Q4 新聞稿（FY2017 比較數）；macrotrends 為 ASC 606 追溯重編後"),
        (2018, 4.45, 3.23, 2.96, "公司 FY2018 Q4 新聞稿；macrotrends 為 ASC 606 追溯重編後（FY2019 採用）"), (2019, 3.04, 2.86, 2.86, "公司 FY2020 Q4 新聞稿（FY2019 比較數）"),
        (2020, 4.17, 3.92, 3.92, "公司 FY2020 Q4 新聞稿"), (2021, 6.84, 6.40, 6.40, "公司 FY2021 Q4 新聞稿"),
        (2022, 7.70, 7.44, 7.44, "公司 FY2023 Q4 新聞稿（FY2022 比較數）"), (2023, 8.05, 8.11, 8.11, "公司 FY2023 Q4 新聞稿"),
        (2024, 8.65, 8.61, 8.61, "公司 FY2025 Q4 新聞稿（FY2024 比較數）"), (2025, 9.42, 8.66, 8.66, "公司 FY2025 Q4 新聞稿；stockanalysis 8.66")]
EP = {}
for i, (fy_, ng, ga, mt, note) in enumerate(EPSD, 6):
    V.cell(i, 1, f"FY{fy_}")
    for j, v in ((2, ng), (3, ga), (4, mt)): V.cell(i, j, v).font = BLUE; V.cell(i, j).number_format = "0.00"
    V.cell(i, 5, f"=D{i}/C{i}-1").number_format = "0.0%"
    V.cell(i, 6, "Verified"); V.cell(i, 7, note); EP[fy_] = i
V.cell(16, 1, "規則用 non-GAAP（模型 EPS、公司指引、共識同口徑）。non-GAAP 的第二來源：未找到（各資料商多列 GAAP）；GAAP 兩來源除 FY2017、FY2018（ASC 606 重編）外相同")
V.cell(18, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 19, ["年底", "收盤價（未調整）", "non-GAAP EPS FY(t)", "GAAP EPS FY(t)", "P/E non-GAAP（自算）", "P/E GAAP（自算）", "companiesmarketcap P/E", "自算 GAAP − cmc", "股利調整後收盤（statmuse）", "標記", "來源／說明"])
PX = {2016: 32.27, 2017: 51.12, 2018: 32.74, 2019: 61.04, 2020: 86.30, 2021: 157.36, 2022: 97.38, 2023: 162.07, 2024: 162.63, 2025: 256.99}
ADJ = {2016: 29.10, 2017: 46.52, 2018: 30.27, 2019: 57.47, 2020: 82.31, 2021: None, 2022: None, 2023: None, 2024: None, 2025: 256.08}
CMC = {2016: 18.6, 2017: 14.7, 2018: 9.20, 2019: 20.0, 2020: 21.0, 2021: 23.5, 2022: 12.7, 2023: 19.6, 2024: 18.6, 2025: 29.5}
for i, y in enumerate(range(2016, 2026), 20):
    V.cell(i, 1, f"{y}-12")
    V.cell(i, 2, PX[y]).font = BLUE
    V.cell(i, 3, f"=B{EP[y]}"); V.cell(i, 4, f"=C{EP[y]}")
    V.cell(i, 5, f"=B{i}/C{i}"); V.cell(i, 6, f"=B{i}/D{i}")
    V.cell(i, 7, CMC[y]).font = BLUE; V.cell(i, 8, f"=F{i}-G{i}")
    if ADJ[y] is not None: V.cell(i, 9, ADJ[y]).font = BLUE
    for c in "BCDI": V[f"{c}{i}"].number_format = "0.00"
    for c in "EFGH": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, {2018: "FY2018 non-GAAP EPS 4.45 為當年高峰（稅改後）→ P/E 7.4 為極端低值", 2025: "statmuse 2025-12-31 收盤 256.08（股利調整後）；digrin 未調整 256.99"}.get(y, "收盤價：digrin（Real price，未調整）"))
V.cell(31, 1, "收盤價來源1：digrin（AMAT Price History 的 Real price，未調整股利）；來源2：statmuse 年底收盤（股利調整後，2016–2020、2025）——兩者比值 0.90–1.00 隨年份遞增，符合股利調整，故未調整價可信；2021–2024 第二來源：未找到。companiesmarketcap 的年底 P/E 為 GAAP TTM 口徑，與自算 GAAP 差距見 H 欄（−1.0 至 +2.0）。")
V.cell(33, 1, "3. 分位（10 年）").font = BOLD
hdr(V, 34, ["統計", "P/E non-GAAP（自算；規則用）", "P/E GAAP（自算；參考）", "cmc P/E（參考）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 35):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E20:E29,{p})"); V.cell(i, 3, f"=PERCENTILE(F20:F29,{p})"); V.cell(i, 4, f"=PERCENTILE(G20:G29,{p})")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(38, 1, "年數"); V["B38"] = "=COUNT(E20:E29)"; V["C38"] = "=COUNT(F20:F29)"; V["D38"] = "=COUNT(G20:G29)"
defname("pe_peak", "Val!$B$35"); defname("pe_mid", "Val!$B$36"); defname("pe_trough", "Val!$B$37")
V.cell(40, 1, "敏感度：去掉 2018 年（non-GAAP 極端低值 7.4）後的分位（9 年）").font = BOLD
hdr(V, 41, ["統計", "P/E non-GAAP（不含 2018）"])
for i, (nm, p) in enumerate((("25 分位", 0.25), ("中位數", 0.5), ("75 分位", 0.75)), 42):
    V.cell(i, 1, nm); V.cell(i, 2, f"=PERCENTILE(L20:L29,{p})").number_format = "0.0"
for i in range(20, 30):
    V.cell(i, 12, f'=IF(A{i}="2018-12","",E{i})').number_format = "0.0"
V.cell(19, 12, "輔助：不含 2018")

V.cell(59, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(60, 1, CP.rule_text(PEAK) + "。AMAT 的 FY(t)（約 t−1 年 11 月～t 年 10 月）有 10/12 落在日曆年 t，直接對照 t，不平移")
hdr(V, 61, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "規則"])
for i, s in enumerate(SC, 62):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "tools/cycle_position.py")
V.cell(66, 1, "5. 隱含股價（美元；約 FY2027 年底，未折現）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 67, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 68; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{62 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = "0.00"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{62 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V.cell(vr + 1, 1, "6. 目前倍數").font = BOLD
hdr(V, vr + 2, ["項目", "A", "B", "C", "備註"])
fr = vr + 3; FR = {}
for key, nm, f, fm, note in (("px", "現價（stockanalysis 2026-10-23 收盤）", "=px", "0.00", ""),
                             ("sa", "遠期 P/E：stockanalysis", "=sa_fpe", "0.0", "stockanalysis 顯示 21.39x"),
                             ("cons", "現價 ÷ 共識 FY2027 EPS（18.49）", "=px/cons_eps27", "0.0", "與 stockanalysis 遠期 P/E 接近（其遠期 EPS 為未來 12 月或 FY2027 共識）"),
                             ("ttm", "現價 ÷ 近 12 月 non-GAAP EPS（FY25 Q4＋FY26 Q1–Q3）", "=px/(Q!E15+SUM(Q!G15:I15))", "0.0", "stockanalysis TTM P/E 32.54 為 GAAP 口徑"),
                             ("e26", "現價 ÷ FY2026 EPS（錨點）", None, "0.0", "三情境相同"),
                             ("e27", "現價 ÷ FY2027 EPS（模型）", None, "0.0", ""),
                             ("e28", "現價 ÷ FY2028 EPS（模型）", None, "0.0", ""),
                             ("rat", "現價 ÷ FY2027 EPS（A）÷ 規則倍數（A）", None, "0.00", ">1＝現價高於規則給的倍數")):
    V.cell(fr, 1, nm); V.cell(fr, 5, note); FR[key] = fr
    if f: V.cell(fr, 2, f).number_format = fm
    elif key in ("e26", "e27", "e28"):
        y = {"e26": 2026, "e27": 2027, "e28": 2028}[key]
        for j, s in enumerate(SC): V.cell(fr, 2 + j, f"=px/PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = fm
    else:
        V.cell(fr, 2, f"=B{FR['e27']}/E62").number_format = fm
    fr += 1
V_END = fr

# ---------------- Val 第 0 節：目標價與上檔空間（v0.1.1 新增；放在既有空白列，不位移既有列） ----------------
TP0 = 46; TPR = {}
assert all(V.cell(rw, c).value is None for rw in range(TP0 - 1, TP0 + 9) for c in range(1, 12)), TP0
assert V["A3"].value is None
V["A3"] = f"目標價與上檔空間：見第 0 節（本頁第 {TP0} 列；v0.1.1 新增，為了不位移既有列而放在空白列）"
V.cell(TP0, 1, "0. 目標價與上檔空間（12 個月，情境加權；權重為 In 頁 tp_wA／tp_wB／tp_wC，Andy 可改）").font = BOLD
for k_, (key_, lab_) in enumerate((("tp", "目標價（12 個月，情境加權）"), ("up", "上檔空間（目標價 ÷ 現價 − 1）"),
        ("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"),
        ("w", "情境權重 A／B／C"), ("m", "目標價方法")), TP0 + 1):
    V.cell(k_, 1, lab_); TPR[key_] = k_
V.cell(TPR["tp"], 2, f"=tp_wA*B{TPR['A']}+tp_wB*B{TPR['B']}+tp_wC*B{TPR['C']}").number_format = "#,##0"
V.cell(TPR["tp"], 5, "＝權重 A × 情境 A 隱含價＋權重 B × 情境 B 隱含價＋權重 C × 情境 C 隱含價（US$）")
V.cell(TPR["up"], 2, f"=B{TPR['tp']}/px-1").number_format = "0.0%"
V.cell(TPR["up"], 5, "相對「" + V.cell(FR["px"], 1).value + "」")
for s_ in SC:
    V.cell(TPR[s_], 2, f"=F{VAL[s_]}").number_format = "#,##0"
    V.cell(TPR[s_], 5, f"＝規則值隱含股價 F{VAL[s_]}（{TITLES[s_]}）")
for j_, s_ in enumerate(SC):
    V.cell(TPR["w"], 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(TPR["w"], 5, "In 頁輸入（Assumed，Andy 可改）：B 欄＝情境 A、C 欄＝情境 B、D 欄＝情境 C")
V.cell(TPR["m"], 2, "P/E 規則值：各情境 FY2027 non-GAAP EPS × 2016–2025 年底 non-GAAP P/E（自算）依設備股共用循環規則取分位（見頂＝25 分位、中段＝中位數、谷底＝75 分位；見頂年讀 markets/equipment）。隱含價約為 FY2027 年底價、未折現，作 12 個月目標價的近似；情境 A／B／C 即本模型三個設備市場情境（A 最高、C 最低）。偏誤：十年倍數含 2018 年極端低值（7.4x），去掉後分位略高（第 3 節敏感度），目標價可能略偏低。")
for c_ in range(1, 6):
    V.cell(TPR["tp"], c_).fill = KEY; V.cell(TPR["up"], c_).fill = KEY
V.cell(TPR["tp"], 1).font = BOLD
defname("tp_price", f"Val!$B${TPR['tp']}"); defname("tp_upside", f"Val!$B${TPR['up']}")
V.column_dimensions["A"].width = 50
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["K"].width = 80

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 US$M", "FY2027 EPS US$", "相對基準", "FY2028 營收 US$M", "FY2028 EPS US$", "相對基準"])
cases = [("份額係數 0.95", "AMAT 份額 −5%（中國國產替代、競爭）", dict(sf="sf_lo")), ("份額係數 1.05", "份額 +5%（環繞式閘極、背面供電、先進封裝）", dict(sf="sf_hi")),
         ("第四季 DRAM 占比 26%", "FY2026 終端權重偏向代工邏輯", dict(mdr="m_dr_lo", mfl="m_fl_lo")), ("第四季 DRAM 占比 34%", "權重偏向 DRAM", dict(mdr="m_dr_hi", mfl="m_fl_hi")),
         ("會計年度權重 1/12", "FY 幾乎等於日曆年", dict(fyw="fyw_lo")), ("會計年度權重 3/12", "設備出貨領先一季", dict(fyw="fyw_hi")),
         ("營業成本變動比例 75%", "固定成本較多 → 毛利率對營收更敏感", dict(cF="(1-phic_lo)*cogs25", cv="phic_lo*cogs25/a25_rev")),
         ("營業成本變動比例 95%", "幾乎全變動 → 毛利率接近持平", dict(cF="(1-phic_hi)*cogs25", cv="phic_hi*cogs25/a25_rev")),
         ("營業費用變動比例 35%", "FY2025 營業費用中變動部分 35%", dict(oF="(1-phi_lo)*opex25", ov="phi_lo*opex25/a25_rev")),
         ("營業費用變動比例 75%", "變動 75%（費用隨營收擴張）", dict(oF="(1-phi_hi)*opex25", ov="phi_hi*opex25/a25_rev")),
         ("AGS 成長 11%", "情境 A −4 個百分點", dict(gags="g_ags_A_lo")), ("AGS 成長 19%", "情境 A +4 個百分點", dict(gags="g_ags_A_hi")),
         ("其他 −10%／年", "顯示器下滑", dict(goth="g_oth_A_lo")), ("其他 +10%／年", "", dict(goth="g_oth_A_hi")),
         ("淨利÷營業利益 0.85", "稅率上升", dict(nioi="nioi_lo")), ("淨利÷營業利益 FY2026 錨點值", "含稅務利益", dict(nioi="nioi_hi")),
         ("淨回購 0%／年", "", dict(bb="bb_lo")), ("淨回購 2.5%／年", "", dict(bb="bb_hi")),
         ("關係企業規則恢復（US$600M／年）", "情境 A 加上 C 的出口管制假設", dict(aff="1")),
         ("關係企業規則恢復（US$900M／年）", "影響 +50%", dict(aff="1", affv="aff_hit_hi")),
         ("中國營收 −20%", "半導體系統＋AGS 中的中國部分 −20%", dict(cn="0.2")), ("中國營收 −40%", "", dict(cn="cn_cut_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "0.00"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
base_row = 60
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, base_row + n * 32, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "設備市場情境 B", "2028 持平、2029 −5%；AGS 11%、其他 −5%", PB["B"], "PL!")
put_case(nn + 1, "設備市場情境 C", "2028 −20%、2029 持平；關係企業規則恢復；AGS 7%、其他 −10%", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 有 5/6 落在 CY2027，而設備市場三情境在 CY2027 以前相同，所以市場情境對 FY2027 影響小、對 FY2028 才明顯；對 FY2027 最敏感的是份額、營業成本與營業費用的固定／變動拆分、出口管制。")
SENS_NOTE = nn + 3
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 46
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、公司說法、SEMI、第四季年化。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv_, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv_, getter(s)).number_format = fmt
        L = chr(64 + cv_)
        RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. 共識（stockanalysis，資料更新 2026-10-16；只作對照）").font = BOLD; rr_ += 1
rec("FY2026 營收（錨點）", "=cons_rev26", pl("rev", 2026), "stockanalysis", "錨點＝前三季實績＋第四季指引中點，非模型預測")
rec("FY2026 non-GAAP EPS（錨點）", "=cons_eps26", pl("eps", 2026), "stockanalysis", "", "0.00")
rec("FY2027 營收", "=cons_rev27", pl("rev", 2027), "stockanalysis", "")
rec("FY2027 non-GAAP EPS", "=cons_eps27", pl("eps", 2027), "stockanalysis", "", "0.00")
RC.cell(rr_, 1, "B. 第四季年化（FY2026 Q4 指引 × 4）vs FY2027 模型").font = BOLD; rr_ += 1
rec("半導體系統", "=g_ss_q4*4", pl("ss", 2027), "公司 Q4 指引", "FY2027 模型若接近第四季年化，代表 FY2027 各季大致持平於第四季水準")
rec("營收", "=g_rev_q4*4", pl("rev", 2027), "公司 Q4 指引", "")
RC.cell(rr_, 1, "C. 市場與公司說法").font = BOLD; rr_ += 1
rec("AGS FY2026 年增（重編口徑）vs 公司 CY2026「>20%」", 0.20, lambda s: f"=PL!D{PB[s]['ags']}/PL!C{PB[s]['ags']}-1", "公司 Q3 FY26 法說", "FY 與 CY 口徑差一季", fmt="0.0%", pt=True)
rec("AI 相關占 2025→2027 設備增量：設備市場模型 vs 公司「約 80%」", "=co_ai27", lambda s: "=ai_inc27", "公司 Q3 FY26 法說；markets/equipment",
    "公司口徑含先進封裝與先進代工邏輯、DRAM；模型 63% 較低，代表公司看的成長更集中在 AI 相關（只作對照，不改上游）", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "D. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "AMAT 半導體系統 FY2026 ÷ SEMI 晶圓前段 CY2026（隱含份額）"); RC.cell(rr_, 2, "=semi_wfe26").number_format = "0.0"; RC.cell(rr_, 3, f"=PL!D{PB['A']['ss']}/1000/semi_wfe26").number_format = "0.0%"; RC.cell(rr_, 9, "SEMI 2026-07（equipment 成品）"); RC.cell(rr_, 10, "對設備市場模型的 FY 口徑為 17.4%（Units 頁）；SEMI 較低，故對 SEMI 的份額較高；AMAT 半導體系統另含先進封裝等"); rr_ += 1
RC.cell(rr_, 1, "FY2026 前三季營收 ÷ FY2026 錨點"); RC.cell(rr_, 3, "=SUM(Q!G7:I7)/an_rev").number_format = "0.0%"; RC.cell(rr_, 10, "第四季指引占全年 30%：AI 需求逐季上升"); rr_ += 1
RC.cell(rr_, 1, "FY2027 營收年增（情境 A）"); RC.cell(rr_, 3, f"=PL!E{SUMR[('A','rev')]}/PL!D{SUMR[('A','rev')]}-1").number_format = "0.0%"; RC.cell(rr_, 10, "共識隱含 FY2027 年增見下一列"); rr_ += 1
RC.cell(rr_, 1, "共識隱含 FY2027 營收年增"); RC.cell(rr_, 3, "=cons_rev27/cons_rev26-1").number_format = "0.0%"; rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 64
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 34; RC.column_dimensions["J"].width = 80

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司新聞稿與 SEC 文件以網頁擷取工具讀取；10-Q 全文擷取在附註 5 前截斷，地區別與分部重編說明未能讀到）"; SO["A1"].font = BOLD
srcs = [("公司：Q3 FY2026 新聞稿（2026-08-13）", "https://ir.appliedmaterials.com/news-releases/news-release-details/applied-materials-announces-third-quarter-2026-results"),
 ("公司：Q2 FY2026 新聞稿（2026-05-14）", "https://ir.appliedmaterials.com/news-releases/news-release-details/applied-materials-announces-second-quarter-2026-results"),
 ("公司：Q1 FY2026 新聞稿（2026-02-12；BIS 和解金 US$253M、分部重編）", "https://ir.appliedmaterials.com/news-releases/news-release-details/applied-materials-announces-first-quarter-2026-results/"),
 ("公司：Q4 FY2025 新聞稿（SEC 8-K 附件，2025-11-13）", "https://www.sec.gov/Archives/edgar/data/6951/000162828025051998/exhibit991q42025earningsre.htm"),
 ("公司：8-K（2025-09-29；BIS 關係企業規則影響 FY2026 約 US$600M）", "https://www.sec.gov/Archives/edgar/data/6951/000162828025043753/amat-20250929.htm"),
 ("公司：10-Q Q3 FY2026（封面流通股 793,597,443）", "https://www.sec.gov/Archives/edgar/data/0000006951/000162828026058235/amat-20260726.htm"),
 ("公司：Q1 FY2026 10-Q 分部附註（XBRL R21）", "https://www.sec.gov/Archives/edgar/data/6951/000162828026009694/R21.htm"),
 ("公司：FY2023 Q4 新聞稿（FY2023、FY2022 EPS）", "https://www.sec.gov/Archives/edgar/data/6951/000000695123000035/exhibit991q42023earningsre.htm"),
 ("公司：FY2021 Q4 新聞稿（FY2021、FY2020 EPS）", "https://investor.appliedmaterials.com/news-releases/news-release-details/applied-materials-announces-fourth-quarter-and-fiscal-year-2021"),
 ("公司：FY2020 Q4 新聞稿（FY2020、FY2019 EPS）", "https://www.sec.gov/Archives/edgar/data/6951/000000695120000041/exhibit991q42020earnin.htm"),
 ("公司：FY2018 Q4 新聞稿（FY2018、FY2017 EPS）", "https://www.sec.gov/Archives/edgar/data/0000006951/000000695118000032/exhibit991q42018earningsre.htm"),
 ("公司：FY2016 Q4 新聞稿（FY2016 EPS）", "https://www.sec.gov/Archives/edgar/data/6951/000000695116000063/exhibit991earningsreleaseq.htm"),
 ("MarketBeat：Q3 FY26 法說重點（第四季分部指引、AGS、80% 說法）", "https://www.marketbeat.com/instant-alerts/applied-materials-q3-earnings-call-highlights-2026-08-13/"),
 ("Nasdaq：Q3 FY26 法說重點（可能與 MarketBeat 同源）", "https://www.nasdaq.com/articles/applied-materials-q3-earnings-call-highlights"),
 ("Skadden：BIS 暫停關係企業規則一年（2025-11-10～2026-11-09）", "https://www.skadden.com/insights/publications/2025/11/bis-suspends-affiliates-rule-for-one-year-as-part-of-the-us-china-trade-deal"),
 ("Reuters（Yahoo Finance 2025-11-14）：中國占比降至約 25%、暫停後約 US$600M 可恢復", "https://finance.yahoo.com/news/applied-materials-shares-drop-stringent-111350305.html"),
 ("stockanalysis：股價、市值、遠期 P/E（2026-10-23）", "https://stockanalysis.com/stocks/amat/"),
 ("stockanalysis：共識（2026-10-16）", "https://stockanalysis.com/stocks/amat/forecast/"),
 ("stockanalysis：年度與季度財報（GAAP 對照）", "https://stockanalysis.com/stocks/amat/financials/?p=quarterly"),
 ("tickflow：共識（過時，未採）", "https://www.tickflow.io/stock/AMAT/forecast"),
 ("macrotrends：GAAP EPS 歷史（第二來源）", "https://www.macrotrends.net/stocks/charts/AMAT/applied-materials/eps-earnings-per-share-diluted"),
 ("digrin：年底收盤價（未調整與調整後）", "https://www.digrin.com/stocks/detail/AMAT/price"),
 ("statmuse：年底收盤價（股利調整後）", "https://www.statmuse.com/money/ask/applied-materials-closing-price-december-31-2019"),
 ("companiesmarketcap：年底 P/E（GAAP TTM）", "https://companiesmarketcap.com/applied-materials/pe-ratio/"),
 ("市場層：markets/equipment 成品（本 repo）；tools/cycle_position.py", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Applied Materials（AMAT）收支與估值模型 v0.2（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("模型命題：AMAT 的營收約四分之三來自半導體系統，跟著晶圓廠設備支出走；成長集中在先進代工邏輯、DRAM／HBM 與先進封裝，中國約占三成。股價取決於 2028 年設備支出會不會見頂，以及出口管制會不會再收緊。", False), ("", False),
 ("會計年度：約 11 月到隔年 10 月底（52／53 週；FY2026＝2025-10-27～2026-10-25）。FY(t) 有 10/12 落在日曆年 t：日曆年轉會計年度用 FY(t)＝2/12×CY(t−1)＋10/12×CY(t)；循環位置把 FY(t) 直接對照日曆年 t。FY2026 已結束但第四季尚未公布（預定 2026-11-12），本版 FY2026＝前三季實績＋第四季公司指引（錨點）。金額為百萬美元；EPS 為 non-GAAP 稀釋（與指引、共識同口徑）。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("半導體系統＝FY2026 錨點 × Σ（FY2026 終端市場權重 × 各終端市場設備支出的會計年度指數）× 份額係數。設備支出讀 markets/equipment 的三情境（晶圓代工邏輯＝先進邏輯＋其他代工邏輯＋其他前段；DRAM 含 HBM；NAND）。", False),
 ("AGS 服務、其他（含顯示器）：前一年 × 成長率（A 取公司長期 mid-teens）。出口管制：情境 C 假設 BIS 關係企業規則在暫停期（至 2026-11-09）後恢復，每年少 US$600M（公司 2025-09 估計）。", False),
 ("毛利率：營業成本＝固定＋變動 × 營收；營業費用＝固定＋變動 × 營收（兩者都用 FY2025 實績與 FY2026 錨點兩點推導）；淨利＝營業利益 × FY2025 的 non-GAAP 淨利 ÷ 營業利益。不寫死利益率。", False),
 ("Val：P/E 規則——10 年（2016–2025 年底）non-GAAP P/E 分位，見頂年 25 分位、中段中位數、谷底 75 分位；循環位置依設備股共用規則（各情境見頂年度讀設備市場模型：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。本版未做先驗，無 Andy 合理 P/E 對照。Recon：共識、第四季年化、SEMI、公司說法。Sens：份額、終端權重、會計年度權重、成本與費用的固定／變動拆分、AGS、出口管制、中國", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("已知限制：(1) 半導體系統只用終端市場設備指數 × 份額係數，份額變化（環繞式閘極、背面供電、先進封裝、中國國產替代）只在 Sens；(2) FY2025 Q4 的重編分部由 FY2026 Q4 指引年增率（整數百分比）反推；(3) 毛利率與營業費用只有兩點；(4) 第四季終端市場組合、股數為 Assumed；(5) 中國約三成營收，出口管制只以單一金額（公司 2025-09 估計）建模；(6) 設備市場模型 2026 年以後高於 SEMI 約 7–10%，本模型只用成長指數；(7) FY2028 共識未找到。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('科林研發（Lam Research）', 'LRCX', 22.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('科磊（KLA）', 'KLAC', 26.65, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('ASML（美國 ADR）', 'ASML', 24.74, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2028: "1"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ FY2028（會計年度 11–10 月）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='US$',
    unit_rev='US$M',
    unit_ni='US$M',
    fmt='0.00',
    fmt_rev='#,##0',
    band_fmt='0.00',
    roe=('FY2023–FY2025', [0.4804, 0.4061, 0.3551], 'stockanalysis 比率頁（2026-10-25 讀取）：FY2023 48.04%、FY2024 40.61%、FY2025 35.51%', '單一來源；近 12 月 41.07%（同站）'),
    nim=('FY2016–FY2025，GAAP', [(10.82, 1.72), (14.7, 3.52), (16.7, 3.04), (14.61, 2.71), (17.2, 3.62), (23.06, 5.89), (25.79, 6.53), (26.52, 6.86), (27.18, 7.18), (28.37, 7.0)], 'vcpscanner.com 年度損益（US$B）', '第二來源：stockanalysis 財報頁 FY2021–25 營收與淨利相同；模型淨利為 non-GAAP，十年平均為 GAAP（略低）'),
    cons_tp=(627.83, 'stockanalysis：40 位分析師平均 US$627.83（中位數 607.50；區間 358–900），資料更新 2026-10-16', '第二來源：未找到'),
    cons_feps=('=cons_eps27', '共識前瞻 EPS＝共識 FY2027 non-GAAP EPS（In 頁 cons_eps27，stockanalysis 18.49）；本模型前瞻 EPS 為 FY2028，晚一年，EPS 因素偏大', 'Derived'),
    peer_note='前段設備同業：科林研發、科磊、ASML、東京威力科創',
    dcf_note='；會計年度 11–10 月，FY2027 大致對齊 t＝1',
    pe10='E20:E29',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（non-GAAP 自算）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝non-GAAP EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均，FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：科林研發、科磊、ASML、東京威力科創前瞻 P/E 中位數 × FY2028 EPS。偏誤：non-GAAP EPS 略高於 GAAP，偏高估；十年平均淨利率為 GAAP，情境 B、C 收斂偏保守。',
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
    inrow("tp_r", "折現率（股東權益要求報酬；美、台、日、韓同一值）", "%", 0.10, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%）", CFG.get("r_note", "—"), "0.0%", True)
    inrow("tp_g", "長期成長率（永續成長；收斂法 FY2032 營收年增的基準）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_roe", f"ROE（近三年平均，{CFG['roe'][0]}）", "%", "=AVERAGE(" + ",".join(str(x) for x in CFG["roe"][1]) + ")", "Derived", CFG["roe"][2], CFG["roe"][3], "0.0%")
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
if os.environ.get("V2_ROWS"):
    print("V2", V2T, V2EXT)

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.2', 'v0.2.1'
OBS = [
 "目標價 US$352（DCF 253、Comparable 451 各半），比現價 377.24（2026-10-23）低 6.7%，級距 0%～−25%；Comparable 單獨為 +19.6%。兩法差距大：DCF 比現價低 33.0%，因情境 B、C 的 EPS 在 FY2028 後依收斂法持平或下滑（B 的 FY2032 EPS 16.01 低於 FY2027 的 16.48），不是 ROE（近三年 41.4%）；本模型前瞻 EPS 17.55 比共識 FY2027 EPS 18.49 低 5.1%（期間晚一年）。",
 "中國是最敏感的單一輸入：中國營收 −20% 使 FY2028 EPS −6.9%、−40% 使 −13.7%；設備市場情境 C（2028 年 −20%）使 FY2028 EPS 由 19.85 降到 13.95（−29.7%），DCF 價格帶 US$161～355。",
 "下一個檢驗點是 FY2026 第四季財報（預定 2026-11-12）的 FY2027 第一季指引與中國占比；與共識目標價 627.83 的差距主要是倍數：共識隱含 34.0 倍，比同業前瞻 P/E 中位數 25.7 倍高 32.1%。"
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

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "amat")
wb.save(sys.argv[1])
print("PB", {s: (PB[s]["ss"], PB[s]["rev"], PB[s]["eps"], PB[s]["idx"]) for s in SC}, "SUMR", SUMR[("A", "rev")], SUMR[("C", "eps")], "VAL", VAL, "FR", FR, "SENS_LAST", SENS_LAST, "RECON_LAST", RECON_LAST)
