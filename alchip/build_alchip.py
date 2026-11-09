# 世芯-KY（3661.TW）收支與估值模型 v0.1.1（AI 半導體 Project：運算層，ASIC 設計服務＋量產；樣板 mediatek/、mrvl/）
# 用法：python3 alchip/build_alchip.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；每顆價格：千美元（US$K）；EPS：新台幣（基本）。
# 營收＝NRE（委託設計）＋量產（turnkey：晶圓、封裝、測試轉售，毛利率低）。
# Trainium 量產營收＝AWS Trainium 顆數（tsmc 成品）× 世芯份額 × 每顆營收（Broadcom 模型的 pa_trn 扣 HBM，與 Marvell 模型同口徑）× k_x × 匯率。
import sys, os, glob, re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import cycle_position as CP
import fcff_dcf as FD   # v0.2：標準版 DCF（無槓桿自由現金流＋WACC）共用模組（tools/fcff_dcf.py，原樣取自 NVIDIA v0.3）
import valuation_rule as VR
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "markets", "scen")); import scen as SCN   # noqa: E402  v0.16：產業情境 A／C 參考值（市場層；B 讀 Siliconomics）

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
TSMC_F = only("tsmc/dist/*.xlsx"); AVGO_F = only("avgo/dist/*.xlsx"); MRVL_F = only("mrvl/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith(("u_trn", "g29_trn", "um"))):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"u_trn_{y}" in TS for y in (26, 27, 28)) and "g29_trn" in TS
for s in SC:      # 情境與行業一致：tsmc 顆數係數＝markets/memory SCEN_UNITS（同一字母）
    assert abs(TS[f"um28_{s}"][0] - MEM.SCEN_UNITS[s][0]) < 1e-12 and abs(TS[f"um29_{s}"][0] - MEM.SCEN_UNITS[s][1]) < 1e-12, s
GBT = [g for k, nm, u_, g, t, sr in MEM.UNITS if k == "trn"][0]
UT = [u_ for k, nm, u_, g, t, sr in MEM.UNITS if k == "trn"][0]
for i, y in enumerate((26, 27, 28)):
    assert abs(TS[f"u_trn_{y}"][0] - UT[i]) < 1e-12, y          # tsmc 與 markets/memory 同一份顆數
AV = {}
for r in openpyxl.load_workbook(AVGO_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("hbm_") or k in ("f_hbm", "m_hbm", "pa_trn")):
        AV[k] = ((r[3].value, r[4].value, r[5].value), f"{rel(AVGO_F)}｜In!E{r[0].row}")
assert "pa_trn" in AV and "f_hbm" in AV
MV = {}
for r in openpyxl.load_workbook(MRVL_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith(("msh_trn", "core_trn")):
        MV[k] = ((r[3].value, r[4].value, r[5].value), f"{rel(MRVL_F)}｜In!E{r[0].row}")
assert len(MV) == 4, MV.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；每顆價格：千美元（US$K）；顆數：百萬顆；匯率：新台幣／美元。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝世芯新聞稿與法說（alchip.com）；公司說法為 Interested-party；券商、資料商、媒體為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR2 = "https://www.alchip.com/en/Newsroom/Alchip_2026_Q2_financial_results"
CT2 = "https://www.ctee.com.tw/news/20260814701424-430201"
INV2 = "https://ca.investing.com/news/transcripts/earnings-call-transcript-alchip-q2-2026-tops-eps-view-as-ai-ramp-drives-growth-93CH-4800784"
CNY2 = "https://news.cnyes.com/news/id/6577160"
VOC1 = "https://vocus.cc/article/6a01c1ccfd89780001c384b6"
VOC2 = "https://vocus.cc/article/6a805de6fd8978000134e926"
UDN25 = "https://money.udn.com/money/amp/story/5607/9363922"
SEP26 = "https://finance.technews.tw/2026/10/12/alchip-3661-202609-financial-report/"
SEP25 = "https://finance.technews.tw/2025/10/13/alchip-3661-202509-financial-report/"
SA = "https://stockanalysis.com/quote/tpe/3661/"
SAF = "https://stockanalysis.com/quote/tpe/3661/forecast/"
SAI = "https://stockanalysis.com/quote/tpe/3661/financials/"
CTB = "https://www.chinatimes.com/realtimenews/20260814001557-260410"
WR3 = "https://wantrich.chinatimes.com/news/20260817900143-420101"
EU9 = "https://www.ctee.com.tw/news/20260923700937-430201"
MW9 = "https://udn.com/news/story/6848/9732499"
GDR = "https://money.udn.com/money/story/5607/9613262"
URLNAME = {PR2: "公司 2Q26 新聞稿", CT2: "工商時報 2026-08-14（2Q26 財報）", INV2: "Investing.com 2Q26 法說逐字稿（二手）", CNY2: "鉅亨網 2026-08-14（2Q26 法說）",
           VOC1: "方格子 1Q26 法說摘要（二手）", VOC2: "方格子 2Q26 財報摘要（二手）", UDN25: "經濟日報 2026-03-06（FY2025 財報）",
           SEP26: "科技新報（9 月營收，2026-10-12）", SEP25: "科技新報（2025 年 9 月營收）", SA: "stockanalysis.com", SAF: "stockanalysis.com 共識頁", SAI: "stockanalysis.com 財報頁",
           CTB: "時報資訊 2026-08-14（美系外資 5,088）", WR3: "旺得富 2026-08-17（三家美系外資）", EU9: "工商時報 2026-09-23（歐系外資 6,784）",
           MW9: "理財周刊 2026-09-06（作者估計）", GDR: "經濟日報 2026-07-07（GDR 掛牌）"}

sec("A. AWS Trainium 顆數（tsmc 成品；與 HBM、台積電、Broadcom、Marvell 模型同一份；情境係數讀台積電模型＝markets/memory SCEN_UNITS，不另訂）")
for i, y in enumerate((26, 27, 28)):
    inp(f"u_trn_{y}", f"AWS Trainium 顆數 CY20{y}", "百萬顆", None, TS[f"u_trn_{y}"][0], None, "Assumed", TS[f"u_trn_{y}"][1], "markets/memory/market.py UNITS 同值（建置時檢查）" if i == 0 else "同上")
inp("g29_trn", "AWS Trainium CY2029 顆數成長", "%", None, TS["g29_trn"][0], None, "Assumed", TS["g29_trn"][1], "markets/memory 無 2029；沿用台積電模型")
for s in SC:
    for y in (28, 29):
        inp(f"um{y}_{s}", f"情境 {s}：CY20{y} 顆數 ÷ 基準", "x", None, TS[f"um{y}_{s}"][0], None, "Assumed", TS[f"um{y}_{s}"][1],
            "A＝原始顆數；B＝AI 資本支出消化；C＝明顯下修（與台積電、Broadcom、Marvell 模型相同；markets/memory SCEN_UNITS，建置時檢查）" if (s == "A" and y == 28) else "同上")
for i, y in enumerate((26, 27, 28)):
    inp(f"gb_trn_{y}", f"AWS Trainium 每顆 HBM CY20{y}", "GB", None, GBT[i], None, "Derived／Assumed", "markets/memory/market.py UNITS", "CY2029 沿用 CY2028；Marvell 模型同值" if i == 0 else "同上")

sec("B. 每顆規格（Broadcom 模型 pa_trn，唯讀；扣 HBM 的口徑與 Marvell 模型相同）與世芯份額（Marvell 份額讀 Marvell 成品，合計 ≤ 100%）")
inp("pa_trn", "AWS Trainium 每顆（含 HBM，模組層級）", "US$K", None, AV["pa_trn"][0][1], None, "Derived", AV["pa_trn"][1], "Broadcom 模型讀自 NVIDIA 模型 In 頁")
inp("core_trn", "Trainium 每顆核心（不含 HBM）＝ 每顆 − CY2026 HBM 金額", "US$K", None, "=pa_trn-gb_trn_26*hbm_A_26/1000", None, "Derived", "pa_trn − 110 GB × CY2026 HBM 單價（Marvell 模型 core_trn 同式）", MV["core_trn"][1] + " 同值（建置時檢查）")
lo, b_, hi = AV["f_hbm"][0]
inp("f_hbm_ref", "對照：Broadcom 模型的 HBM 代購比例", "x", lo, b_, hi, "Assumed", AV["f_hbm"][1], "世芯不沿用（見下列）；只作 Sens 高值")
inp("f_al", "世芯代購 HBM 並計入營收的比例", "x", 0.0, 0.0, b_, "Assumed", "美系外資：Trainium 4 生命週期可服務市場「不含 HBM」（旺得富 2026-08-17，Interested-party）；3Q26 毛利率約 18–25% 仍高於代購 HBM 情形", "公司未揭露 HBM 由誰採購：未找到；高值＝Broadcom 模型 0.6，只放 Sens")
lo, b_, hi = AV["m_hbm"][0]
inp("m_hbm", "代購 HBM 的毛利率（只在 f_al > 0 時有作用）", "x", lo, b_, hi, "Assumed", AV["m_hbm"][1], "沿用 Broadcom 模型")
for y in (26, 27, 28):
    inp(f"mv_trn_{y}", f"Marvell 占 AWS Trainium 顆數 CY20{y}（份額加總檢查用）", "%", None, MV[f"msh_trn_{y}"][0][1], None, "Assumed", MV[f"msh_trn_{y}"][1], "Marvell 模型；Marvell（基準）＋世芯（高）≤ 100%（建置時檢查）")
# 世芯占 Trainium 份額：Siliconomics In 頁 M7（v0.3 起單一來源；數值與文字由本檔原寫死值逐字搬入）
for key in ["al_trn_26"] + [f"al{y}_{s}" for y in (27, 28) for s in ("A", "B", "C")]:
    y = 26 if key == "al_trn_26" else int(key[2:4])
    inp(*SCN.row9(key))   # v0.16：A、C 讀市場層 markets/scen，B 與 al_trn_26 讀 Siliconomics
    assert MV[f"msh_trn_{y}"][0][1] + SCN.V[key]["hi"] <= 1.0 + 1e-9, key   # Marvell（基準）＋世芯（高）≤ 100%
PFC = {"A": 1.5, "B": 1.2, "C": 1.0}
for s in SC:
    inp(f"pfc_{s}", f"CY2028 起每顆營收係數（2 奈米 Trainium 4；情境 {s}）", "x", 0.8 if s == "A" else None, PFC[s], 2.5 if s == "A" else None, "Assumed",
        "公司：2 奈米產品價格「遠高於」3 奈米、出貨量高於 3 奈米（2Q26 法說，鉅亨網）" if s == "A" else "同上",
        "美系外資：Trainium 4 每顆超過 1 萬美元、2028 年 40 萬顆（時報資訊 2026-08-14，Interested-party）；主要客戶採 COT 模式（法說），每顆營收可能不升反降" if s == "A" else "A 最有利、C 最不利")
inp("uf", "世芯 Trainium 顆數係數（Sens 用；CY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "無", "±20%")
inp("m_trn", "Trainium 量產（turnkey）毛利率（FY2026 校準與 4Q26 用）", "x", 0.10, 0.13, 0.16, "Assumed",
    "公司：3Q、4Q26 綜合毛利率「20 幾」、全年低至中 20%（2Q26 法說，Investing.com 逐字稿）", "Q 頁核對：以 13% 推得非 Trainium 毛利率 1Q–3Q26 約 48–50%，與 1Q26 50.2% 一致")
GTA = {"A": 0.01, "B": 0.0, "C": -0.02}
for s in SC:
    inp(f"gta_{s}", f"Trainium 量產毛利率調整（FY2027 起；情境 {s}）", "pt", None, GTA[s], None, "Assumed",
        "公司：2 奈米專案毛利率至少與 3 奈米相當、可能更高（2Q26 法說）" if s == "A" else "同上", "A 最有利、C 最不利")

sec("C. 記憶體價格路徑（讀 Broadcom 模型成品；只在 f_al > 0 的 Sens 案例有作用；情境別）")
for s in SC:
    for y in (26, 27, 28, 29):
        inp(f"hbm_{s}_{y}", f"HBM 混合單價 CY20{y}（{s}）", "US$/GB", None, AV[f"hbm_{s}_{y}"][0][1], None, "Derived", AV[f"hbm_{s}_{y}"][1], "markets/memory：2026 BofA $14.6/GB；2027 +80%" if (s == "A" and y == 26) else "同上")

sec("D. 其他營收（NRE＋非 Trainium 量產：中國 ADAS、Arm CPU、日本客戶等）：簡化成長率（FY2027 起；情境別）")
GO = {"A": (0.15, 0.12), "B": (0.05, 0.05), "C": (-0.10, -0.05)}
for s in SC:
    inp(f"g_o27_{s}", f"其他營收年增 FY2027（情境 {s}）", "%", -0.15, GO[s][0], 0.25, "Assumed",
        "公司：ADAS 為 2026 年第二大營收來源、兩個 Arm CPU 專案量產、2 奈米 NRE 開始認列；有信心再取得一家大型雲端客戶（2Q26 法說）" if s == "A" else "同上",
        "FY2025 營收 −40.1%（舊量產專案結束；經濟日報）" if s == "A" else "—")
    inp(f"g_oL_{s}", f"其他營收年增 FY2028–29（情境 {s}）", "%", -0.15, GO[s][1], 0.25, "Assumed", "同上", "—")
inp("gmadj", "其他營收毛利率調整（Sens 用；基準 0）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "其他營收毛利率為 FY2026 殘差（Derived），之後沿用")

sec("E. FY2026 錨點（1H26 實績＋3Q26 營收實績＋3Q26 指引＋4Q26 假設）")
inp("dc26", "公司 FY2026 3 奈米 AI 加速器（Trainium）營收", "US$B", 1.3, 1.5, 1.8, "Interested-party", CTB + "：公司 2026 年 3 奈米營收目標超過 15 億美元（轉述）",
    "2Q26 法說：2026 年 3 奈米營收預估不變、訂單幾乎都已下（鉅亨網）；美系外資 13 億、16 億美元（旺得富 2026-08-17）")
inp("q2_trn", "其中 2Q26（6 月起出貨）", "NT$B", 2.0, 3.0, 3.9, "Assumed", "公司：5 月底量產、6 月出貨；2Q26 北美占營收 51%、3／2 奈米 47%（2Q26 新聞稿）", "高值＝北美 51% × 2Q26 營收")
inp("fx_q1", "1Q26 平均匯率（營收新台幣 ÷ 美元）", "NT$/US$", None, "=4.187/0.1324", None, "Derived", PR2 + "：1Q26 US$132.4M", CT2 + "：1H26 營收 118.27 億 − 2Q26 76.40 億")
inp("fx_q2", "2Q26 平均匯率（營收新台幣 ÷ 美元）", "NT$/US$", None, "=7.640/0.2417", None, "Derived", PR2 + "：2Q26 US$241.7M", CT2 + "：2Q26 營收 76.40 億")
inp("fx26", "3Q、4Q26 匯率（FY2026 Trainium 換算用）", "NT$/US$", 30.5, 31.6, 33.0, "Assumed", "2Q26 平均（Derived）", "實際 3Q26 平均匯率：未找到")
inp("q3_rev", "3Q26 營收（月營收實績；7–9 月合計）", "NT$B", None, "=37.878-11.827", None, "Verified", SEP26 + "：前 9 月 378.78 億（9 月 99.79 億）", CT2 + "：1H26 118.27 億；理財周刊：7 月 74.33 億")
inp("q3_trnsh", "3Q26 營收中 Trainium 量產占比", "%", 0.70, 0.80, 0.90, "Assumed", MW9 + "：Trainium 3 占 3Q、4Q 營收 80–90%（Interested-party）", "公司：未來營收 NRE 10–20%、量產 80–90%（2Q26 法說）")
inp("q3_gm", "3Q26 毛利率（公司：低至中 20%）", "%", 0.18, 0.21, 0.25, "Interested-party", INV2, CTB + "：美系外資估 3Q26 約 18%")
inp("q3_opx", "3Q26 營業費用（公司：與 2Q26 相近）", "NT$B", 0.95, 1.05, 1.20, "Interested-party", INV2, "2Q26 1.015（工商時報：毛利 26.62 億 − 營業利益 16.47 億）")
inp("q4_g", "4Q26 營收季增（公司：4Q 再成長）", "%", 0.0, 0.15, 0.30, "Assumed", CNY2, MW9 + "：作者估 4Q26 約 365 億（+40%，Interested-party）")
inp("q4_gmo", "4Q26 非 Trainium（NRE＋其他量產）毛利率", "%", 0.42, 0.48, 0.52, "Assumed", "1Q26 綜合毛利率 50.2%（NRE 比重高；方格子）", "Q 頁核對：2Q、3Q26 推得約 48–49%")
inp("q4_opx", "4Q26 營業費用", "NT$B", 0.95, 1.05, 1.25, "Assumed", INV2 + "：3Q、4Q 與 2Q 相近", "—")
inp("q_nonop", "3Q、4Q26 營業外淨額（每季）", "NT$B", 0.25, 0.40, 0.55, "Assumed", "1Q26 0.417、2Q26 0.388（工商時報推得）", "7 月 GDR 增資（經濟日報 2026-07-07），利息收入可能增加；匯兌損益：未找到")
inp("q_tax", "3Q、4Q26 有效稅率", "%", 0.17, 0.20, 0.23, "Assumed", "1Q26 20.0%、2Q26 19.5%（工商時報推得）", "FY2025 16.7%（stockanalysis 稅前淨利率推得）")
inp("q_sh", "3Q、4Q26 加權股數（GDR 後）", "十億股", None, 0.08621, None, "Interested-party", SA + "：流通股數 86.21M（2026-10-23 擷取）", "2Q26 加權 81.87M（淨利 ÷ EPS）；GDR 7 月 7 日掛牌（經濟日報）")

sec("F. FY2027 起：營業費用、營業外、稅、股數、匯率")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.2, 0.3, 0.6, "Assumed", "公司：營業費用年增「在控制內」、長期營業利益率目標十幾 %（2Q26 法說）", "FY2026 營業費用約 +23%、營收約 +120%（Q 頁）")
inp("g_opx_min", "營業費用年增下限", "%", 0.05, 0.08, 0.12, "Assumed", "2 奈米專案人力、EDA 與伺服器費用增加（工商時報 2Q26）", "—")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 +2 個百分點")
inp("nonop_f", "營業外淨額（每年）", "NT$B", 1.0, 1.6, 2.2, "Assumed", "FY2025 1.704（Derived）、1H26 0.805", "GDR 增資後現金增加：金額未找到")
inp("tax", "有效稅率（FY2027 起）", "%", 0.17, 0.20, 0.25, "Assumed", "FY2025 16.7%、1H26 19.7%（Derived）", "—")
inp("sh_g", "加權股數年變動", "%", -0.005, 0.0, 0.01, "Assumed", "GDR 後 86.21M（stockanalysis）", "員工酬勞股票：小")
inp("fx", "匯率（FY2027 起；Trainium 美元營收換算）", "NT$/US$", 30.0, SIL.base("fx_twd"), 34.0, "Assumed", "與聯發科、台積電模型相同假設", "其他營收以新台幣成長率推，未另換算")

sec("G. 公司說法、共識與券商（只作 Recon 對帳，不作校準）")
inp("co_nre", "公司：未來營收中 NRE 比重（10–20% 中點）", "%", 0.10, 0.15, 0.20, "Interested-party", INV2, "—")
inp("co_gm26", "公司：FY2026 毛利率（低至中 20%，取中點）", "%", 0.21, 0.23, 0.25, "Interested-party", INV2, "—")
inp("cons_rev26", "共識營收 FY2026", "NT$B", 35.4, 70.7, 86.1, "Interested-party", SAF + "（14 位分析師；資料日 2026-08-13）", "FY2027 起共識為付費資料：未找到")
inp("cons_eps26", "共識 EPS FY2026", "NT$", 103.06, 135.68, 163.01, "Interested-party", SAF + "（資料日 2026-08-13）", "美系外資甲 132.94（旺得富 2026-08-17）")
inp("b1_eps27", "美系外資甲 FY2027 EPS", "NT$", None, 204.17, None, "Interested-party", WR3, "理財周刊作者估 195（Interested-party）")
inp("aut_eps27", "理財周刊作者 FY2027 EPS", "NT$", None, 195.0, None, "Interested-party", MW9, "—")
inp("aut_rev27", "理財周刊作者 FY2027 營收", "NT$B", None, 110.4, None, "Interested-party", MW9, "Trainium 3 占 77%")
inp("b1_trn26", "美系外資甲 FY2026 Trainium 3 營收", "US$B", None, 1.3, None, "Interested-party", WR3, "約占全年營收 60%")
inp("b1_trn27", "美系外資甲 FY2027 Trainium 3 營收", "US$B", None, 2.2, None, "Interested-party", WR3, "約占營收 68%")
inp("b2_ai27", "美系外資乙 FY2027 AI ASIC 營收", "US$B", None, 3.8, None, "Interested-party", WR3, "FY2026 16 億美元")
inp("b1_trn4", "美系外資 FY2028 Trainium 4 營收（不含 3 奈米尾量）", "US$B", None, 4.0, None, "Interested-party", WR3 + "；" + CTB, "美系外資丙情境：2 奈米 2028 年 90–120 億美元")
inp("ubs_trn27", "UBS：CY2027 Trainium 3 出貨", "百萬顆", None, 2.8, None, "Interested-party", MW9 + "（轉述 UBS：2026 180 萬、2027 280 萬、2028 460 萬顆）", "市場層 CY2027 全部 Trainium 2.5 百萬顆（tsmc），口徑不同")

sec("H. 價格與市場倍數")
inp("px", "股價（2026-11-05 收盤）", "NT$", None, 2960, None, "Verified", "Yahoo Finance 日線收盤 2026-11-05（3661.TW）", "v0.2.2 為 2026-11-04 收盤 3180；v0.2.1 為 2026-10-22 收盤 3,495（" + SA + "：10-22 收盤 3,495（−5.80%）" + "）")
inp("pe_ttm", "近 12 月 P/E（stockanalysis）", "x", None, 50.74, None, "Interested-party", SA, "companiesmarketcap 47.2")
inp("fpe", "前瞻 P/E（2026-11-05 收盤 ÷ 資料商前瞻 EPS）", "x", None, 14.52, None, "Interested-party", SA + "：stockanalysis.com 統計頁 Forward PE 17.15（2026-10-22 頁面，價格 3495）→ 前瞻 EPS 203.8；P/E＝2026-11-04 基準 15.6 × Yahoo Finance 日線收盤 2026-11-05 2960 ÷ 2026-11-04 收盤 3180（共識 EPS 不變）", "v0.2.2 為 15.6（2026-11-04）；v0.2.1 為 stockanalysis 頁面值（價格日 2026-10-22）")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "NT$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 17.15", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", 3755, 5803.25, 6800, "Interested-party", SA + "：16 位分析師平均 5,803（2026-10-23 擷取）", "共識頁（2026-08-13）：15 位平均 5,666、中位數 6,000；歐系外資 6,784（2026-09-23）")
inp("nim5", "近五年平均淨利率（FY2021–25；對照用）", "%", None, "=AVERAGE(1490/10428,1834/13725,3325/30482,6446/51969,5598/30926)", None, "Derived",
    "stockanalysis 財報頁：FY2021–25 營收與淨利", "十年平均（2016–2025）：未找到（2016–2020 未取得）；收斂法依工作單改用最後一年 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.2 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.2 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與「FY2029 × 0.8」的中點）", "%", None, 0, None, "Assumed", "十年平均未找到，依工作單以最後一年 × 0.8 代替", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上", "—")
# v0.2（2026-11-05）：標準版 DCF（FCFF＋WACC）參數（tools/fcff_dcf.in_rows；docs/plan/個股DCF標準規格.md）。資產負債 2026-06-30（stockanalysis 轉載公司季報）
SABS = "stockanalysis.com 世芯-KY 季度資產負債表（https://stockanalysis.com/quote/tpe/3661/financials/balance-sheet/?p=quarterly；2026-06-30 欄；轉載公司季報；2026-11-05 讀取）"
SAY = "stockanalysis.com 世芯-KY 年度資產負債表（https://stockanalysis.com/quote/tpe/3661/financials/balance-sheet/；2026-11-05 讀取）"
SACF = "stockanalysis.com 世芯-KY 年度現金流量表（https://stockanalysis.com/quote/tpe/3661/financials/cash-flow-statement/；2026-11-05 讀取）"
SADV = "stockanalysis.com 世芯-KY 股利頁（https://stockanalysis.com/quote/tpe/3661/dividend/；2026-11-05 讀取）"
REV3 = "營收 FY2023 30.482、FY2024 51.969、FY2025 30.926（NT$B；stockanalysis 財報頁）"
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.1 DCF 基準日）", ccy="NT$", amt="NT$B", shu="十億股",
        rf=(0.0198, "Trading Economics「Taiwan 10Y Bond Yield」：2026-11-04 1.98%（https://tradingeconomics.com/taiwan/government-bond-yield；2026-11-05 讀取）",
            "同頁：2026-09-24 1.96%；與台積電、聯電、世界先進 DCF 同一值", "無風險利率（新台幣 10 年期公債殖利率，2026-11-04）"),
        rf_us=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取；同 NVIDIA v0.3）",
               "工作單第二輪修正（Andy 2026-11-05 23:16）：基準維持當地利率＋固定匯率；本列只用於 Sens「美元口徑折現」（rf＝美國 10 年期＋CRP、g＝MIN(3%, 5%)）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：TW 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 ^TWII（加權指數） 迴歸，Hamada 去槓桿、現金修正），以本公司市值 D/E 與稅率重新加槓桿（Hamada）", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。Sens 低／高（Analogy；v0.2.1 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）"),
        crp=(0.0078, "Interested-party", "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新）：Taiwan Moody's Aa3、國家風險溢酬 0.78%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）",
             "非美國市場：ERP＝5%＋CRP（Andy 2026-11-05 20:45）；不設股權成本下限；開曼註冊、台灣上市與營運，取台灣"),
        kd=("=rf+0.01", "Assumed", "新台幣 10 年期公債＋信用利差 1%（Assumed；公司無金融借款）", "負債只有租賃（約 0.29 十億新台幣），占資本約 0.1%，對 WACC 影響可忽略"),
        tax_w=("=tax", "Derived", "＝In 頁 tax（FY2027 起有效稅率 20%）", "NOPAT 用同一稅率"),
        debt=(0.0, "Verified", SABS + "：無銀行借款或公司債（Total Debt 289.79 全為租賃負債，NT$M）", "—"),
        lease=(0.28979, "Verified", SABS + "：租賃負債流動 100.36＋非流動 189.43（NT$M）", "—"),
        cash=("=39.405+13.682", "Verified", SABS + "：現金及約當現金 39,405＋短期投資 13,682（NT$M）",
              "2026-10-08 發放的股利改在 stub 期扣除（stub_div；v0.2 自 6 月底現金扣除）；3Q26 資產負債表尚未公布（2026-11-05 確認 stockanalysis 最新為 2026-06-30）"),
        inv=(2.872, "Verified", SABS + "：長期投資 2,872（NT$M；帳面）", "帳面計入；Sens 測不計"),
        mi=(0.01769, "Verified", SABS + "：非控制權益 17.69（NT$M）", "—"),
        stub=dict(bs_date="2026-06-30", val_date="2026-12-31", bs_src=SABS + "（最新已公布一季；3Q26 財報約 11 月上旬公布）",
                  div=("=32.57285*0.08621", "Derived", "2026-10-08 發放的現金股利每股 32.57285 元（除息 2026-09-03）× 86.21M 股（" + SADV + "）；落在 2026-06-30～2026-12-31",
                       "買回無已宣告金額，不扣")),
        dps=(32.57285, "Assumed", "2026 年發放每股 32.57285 元（除息 2026-09-03、發放 2026-10-08；" + SADV + "）",
             "評價日到 2027-12-31 之間假設發放一次、金額與 2026 年相同（Assumed；FY2026 EPS 高於 FY2025，實際可能較高）"),
        sh=("=q_sh", "Interested-party", "＝In 頁 q_sh（GDR 後流通股數 86.21M；stockanalysis）", "2026-06-30 資產負債表股數同為 86.21M"),
        fc_da=(None, "=AVERAGE(2.017/30.482,2.639/51.969,3.052/30.926)", None, "Derived",
               SACF + "：折舊攤銷 FY2023 2,017、FY2024 2,639、FY2025 3,052（NT$M）；" + REV3 + "；比例 6.6%、5.1%、9.9%，平均 7.2%",
               "多為 EDA／IP 授權與設備攤提，近三年營收以 NRE 與舊量產為主；轉為 Trainium turnkey 量產後占比可能下降（README 已知限制）"),
        fc_cx=(None, "=AVERAGE(2.321/30.482,1.998/51.969,2.287/30.926)", None, "Derived",
               SACF + "：資本支出 FY2023 2,321、FY2024 1,998、FY2025 2,287（NT$M）；比例 7.6%、3.8%、7.4%，平均 6.3%", "Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=MAX(fc_da_ss*(1+g_tv),fc_cx)", "Assumed", "FY2029 起線性收斂，FY2032 達到：不低於維持水準（折舊 ×（1＋g））", "近三年資本支出低於折舊攤銷，穩態取維持水準；Andy 可改"),
        fc_wc=(0.0, "=AVERAGE((2.643+12.219+0.1235+0.29359-1.925-0.84346-10.408-0.28254)/30.482,(6.138+9.95+0.11766+0.55073-1.826-1.46-6.114-0.29745)/51.969,(3.254+9.677+0.62172+0.65385-1.986-1.486-12.021-0.60836)/30.926)", 0.15, "Derived",
               "近三年年底營運資金 ÷ 營收平均：應收＋存貨＋預付＋其他流動資產 − 應付 − 應計費用 − 合約負債（客戶預收）− 其他流動負債（" + SAY + "）；比例 6.0%、13.6%、−6.1%，平均 4.5%",
               "2026-06-30：5,552＋24,725＋579＋648 − 7,784 − 1,425 − 15,297 − 3,886＝3,112（NT$M；Trainium 量產備料）÷ 2Q26 年化營收 30.6B＝10.2%；Sens 低／高 0%／15%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β：自算 TW 半導體（同業 13 家中位數，現金修正、Blume 調整）", "x", None, 1.429, None, "Analogy",
    "自算產業 β：TW 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 ^TWII（加權指數） 迴歸，Hamada 去槓桿、現金修正）：中位數 1.64；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）：0.67 × 1.64 ＋ 0.33＝1.429（同業表與方法：docs/plan/自算產業β_20261105.md）",
    "Sens 低／高改為 β_u ∓ beta_du（v0.2.1）；取代 Damodaran 表（Andy 2026-11-05 22:55）")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代上一版 beta_u 低／高欄（中位數 × 0.8 再 Blume 1.209／中位數未調整 1.64）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑")
TP_KEYS += ["tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

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
    for u_, nm_ in sorted(URLNAME.items(), key=lambda kv: -len(kv[0])):   # 網址只列在 Sources 頁；In 頁寫來源名稱（第三輪清單 18：只換完整網址，長者先換，避免截斷成「stockanalysis.comfinancials/」）
        s1 = re.sub(re.escape(u_) + r"(?![\w/%?=&.~-])", nm_, s1); s2 = re.sub(re.escape(u_) + r"(?![\w/%?=&.~-])", nm_, s2)
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
FD.yellow(IN, FD.JUDG + ("beta_u",))
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
# core_trn 與 Marvell 模型同值（建置時檢查）
assert abs(AV["pa_trn"][0][1] - GBT[0] * AV["hbm_A_26"][0][1] / 1000 - MV["core_trn"][0][1]) < 1e-9

# ---------------- Q：實績與 FY2026 錨點（NT$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣、基本；股數為十億股）。藍字＝輸入（公司新聞稿與媒體轉述財報）；黑字＝公式。Trainium 與其他營收的拆分為 Assumed（公司未揭露）"
Q["A2"] = "FY2026 錨點＝1Q、2Q26 實績＋3Q26 營收實績（月營收）與指引（毛利率、費用）＋4Q26 假設；Trainium 全年取公司「3 奈米超過 15 億美元」，4Q26 為全年減 2Q、3Q。3Q26 財報與法說約 11 月上旬（確切日期未找到），屆時以實績取代 E 欄"
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（營收實績＋指引）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "FY2025：經濟日報 2026-03-06、stockanalysis；1Q26、2Q26：工商時報 2026-08-14（公司財報）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 30.926, "C": 4.187, "D": 7.640, "E": "=q3_rev", "F": "=E4*(1+q4_g)"}, tag="Verified／Assumed（4Q）", s2="3Q26：月營收實績（科技新報）；FY2025 經濟日報 309.3 億")
qrow(5, "  Trainium 量產（3 奈米 AI 加速器，北美雲端客戶）", {"B": 0, "C": 0, "D": "=q2_trn", "E": "=E4*q3_trnsh", "F": "=dc26*fx26-D5-E5"}, tag="Interested-party／Assumed",
     s1="公司：5 月底量產、6 月出貨（2Q26 新聞稿）", s2="FY2025 視為 0（3 奈米 Trainium 量產自 2026-05）；4Q＝全年 15 億美元 × 匯率 − 2Q − 3Q")
qrow(6, "  其他（NRE＋非 Trainium 量產）", {c: f"={c}4-{c}5" for c in "BCDEF"}, tag="Derived", s1="—")
qrow(7, "毛利", {"B": 8.154, "C": 2.101, "D": 2.662, "E": "=E4*q3_gm", "F": "=m_trn*F5+q4_gmo*F6"}, tag="Verified／Interested-party（3Q）／Assumed（4Q）",
     s2="1Q26＝1H26 47.63 億 − 2Q26 26.62 億；4Q：Trainium × m_trn ＋ 其他 × q4_gmo")
qrow(8, "毛利率", {c: f"={c}7/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="公司：FY2025 26.4%、1Q26 50.2%、2Q26 34.8%")
qrow(9, "  核對：非 Trainium 毛利率（以 m_trn 推得）", {c: f"=({c}7-m_trn*{c}5)/{c}6" for c in "CDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="1Q–3Q 應接近 4Q 假設 q4_gmo")
qrow(10, "營業費用", {"B": "=8.154-5.016", "C": "=2.101-1.367", "D": "=2.662-1.647", "E": "=q3_opx", "F": "=q4_opx"}, tag="Derived／Interested-party／Assumed", s2="毛利 − 營業利益")
qrow(11, "營業利益", {"B": 5.016, "C": 1.367, "D": 1.647, "E": "=E7-E10", "F": "=F7-F10"}, tag="Verified／Derived", s2="1Q26＝1H26 30.14 億 − 2Q26 16.47 億")
qrow(12, "營業外淨額", {c: f"={c}13-{c}11" for c in "BCD"} | {"E": "=q_nonop", "F": "=q_nonop"}, tag="Derived／Assumed", s2="稅前 − 營業利益")
qrow(13, "稅前利益", {"B": "=30.926*0.2173", "C": 1.784, "D": 2.035, "E": "=E11+E12", "F": "=F11+F12"}, tag="Verified／Derived", s2="FY2025＝營收 × 稅前淨利率 21.73%（stockanalysis）；1Q26＝1H26 38.19 億 − 2Q26 20.35 億")
qrow(14, "所得稅", {c: f"={c}13-{c}16" for c in "BCD"} | {"E": "=E13*q_tax", "F": "=F13*q_tax"}, tag="Derived／Assumed", s2="稅前 − 淨利")
qrow(15, "  有效稅率", {c: f"={c}14/{c}13" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—")
qrow(16, "歸屬母公司淨利", {"B": 5.598, "C": 1.427, "D": 1.639, "E": "=E13-E14", "F": "=F13-F14"}, tag="Verified／Derived", s2="1Q26＝1H26 30.66 億 − 2Q26 16.39 億；無重大非控制權益（假設）")
qrow(17, "加權股數（十億股）", {"B": "=5.598/69.18", "C": "=1.427/17.55", "D": "=1.639/20.02", "E": "=q_sh", "F": "=q_sh"}, fmt="0.00000", ann="=AVERAGE(C17:F17)", tag="Derived", s1="淨利 ÷ EPS（公司）")
qrow(18, "EPS（基本）", {c: f"={c}16/{c}17" for c in "BCDEF"}, fmt="#,##0.00", ann="=G16/G17", tag="Derived", s1="淨利 ÷ 股數")
qrow(19, "  公司揭露 EPS", {"B": 69.18, "C": 17.55, "D": 20.02}, fmt="#,##0.00", ann="", tag="Verified", s2="HiStock：FY2025 69.18、1Q26 17.55、2Q26 20.01")
Q.cell(21, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("FY2026 Trainium ÷ 營收", "=G5/G4", "0.0%"), ("FY2026 其他營收年增（FY2025 全部視為其他）", "=G6/B6-1", "0.0%"),
          ("3Q26 營收季增", "=E4/D4-1", "0.0%"), ("4Q26 Trainium ÷ 4Q26 營收", "=F5/F4", "0.0%"), ("FY2026 營業利益率", "=G11/G4", "0.0%"),
          ("檢查：1H26 營收 自算 − 公司 11.827（應約為 0）", "=C4+D4-11.827", "0.000"), ("檢查：1H26 歸屬淨利 自算 − 公司 3.066（應約為 0）", "=C16+D16-3.066", "0.000"),
          ("檢查：FY2025 營業費用 ÷ 營收", "=B10/B4", "0.0%")]
for i, (lab, f, fm) in enumerate(checks, 22):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("trn", 5), ("oth", 6), ("gp", 7), ("opx", 10), ("oi", 11), ("nonop", 12), ("tax", 14), ("ni", 16), ("sh", 17), ("eps", 18)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 56
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 50; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", uf="1", shs="", shp="0", pfc="pfc_{s}", fh="f_al", mh="m_hbm", mc="m_trn", gta="gta_{s}", fx="fx",
            go27="g_o27_{s}", goL="g_oL_{s}", gma="0", phi="phi", opxa="0", nonop="nonop_f", tax="tax", fg="fade_g{s}", fm="fade_m{s}", wta=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    _wx = lambda y: (f"*{P['wta'][y]}" if P["wta"] and y in P["wta"] else "")   # Wattonomics 錨定（Sens 案例；預設 None＝不乘）
    p = lambda k: P[k].format(s=s)
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
    def uexp(y):
        if y <= 2027: return f"u_trn_{y % 100}"
        if y == 2028: return f"u_trn_28*um28_{s}"
        return f"u_trn_28*(1+g29_trn)*um29_{s}"
    def shr_(y):
        if y == 2026: return "=al_trn_26"
        k = f"al{min(y, 2028) % 100}_{s}{P['shs']}"
        return f"=MIN(1,MAX(0,{k}+{P['shp']}))"
    sec_("A. Trainium 量產（由下而上；D 欄＝FY2026 校準年；FY(t)＝CY(t)）")
    row("u", "AWS Trainium 市場顆數（含情境係數）", "百萬顆", {FC[y]: "=" + uexp(y) for y in BU}, "0.000")
    row("sh", "世芯份額：AWS Trainium", "%", {FC[y]: shr_(y) for y in BU}, "0.0%")
    row("b", "世芯顆數：AWS Trainium", "百萬顆", {FC[y]: f"={FC[y]}{R['u']}*{FC[y]}{R['sh']}" + ("" if y == 2026 else f"*{P['uf']}") + _wx(y) for y in BU}, "0.000")
    row("gb", "每顆 HBM", "GB", {FC[y]: f"=gb_trn_{min(y, 2028) % 100}" for y in BU}, "0")
    row("hpx", "HBM 單價", "US$/GB", {FC[y]: f"=hbm_{s}_{y % 100}" for y in BU})
    row("c", "每顆核心（不含 HBM；CY2028 起 × 2 奈米係數）", "US$K", {FC[y]: "=core_trn" + (f"*{p('pfc')}" if y >= 2028 else "") for y in BU})
    row("h", "每顆 HBM 金額（全部）", "US$K", {FC[y]: f"={FC[y]}{R['gb']}*{FC[y]}{R['hpx']}/1000" for y in BU})
    row("a", "每顆世芯營收（核心＋代購 HBM；未乘 k_x）", "US$K", {FC[y]: f"={FC[y]}{R['c']}+{P['fh']}*{FC[y]}{R['h']}" for y in BU}, note="核心 ＋ 代購比例 × HBM 金額；基準代購比例 0")
    row("xraw", "Trainium 量產營收（未校準）", "US$B", {FC[y]: f"={FC[y]}{R['b']}*{FC[y]}{R['a']}" for y in BU})
    row("kx", "校準係數 k_x ＝ FY2026 公司 3 奈米營收 ÷ 未校準", "x", {"D": f"=dc26/D{r - 1}"}, "0.000",
        note="Derived：以公司 FY2026「3 奈米超過 15 億美元」校準；FY2027 起沿用。k_x < 1＝世芯營收只含晶片與封裝轉售，低於模組價值；也含 CY2026 只有 7 個月量產")
    row("asu", "Trainium 量產營收（美元）", "US$B", {FC[y]: f"={FC[y]}{R['xraw']}*$D${R['kx']}" for y in BU})
    row("fxr", "匯率", "NT$/US$", {"D": "=fx26", **{FC[y]: f"={P['fx']}" for y in F3}})
    row("trn", "Trainium 量產營收", "NT$B", {"C": "=a25_trn", **{FC[y]: f"={FC[y]}{R['asu']}*{FC[y]}{R['fxr']}" for y in BU}})
    row("xh", "  其中代購 HBM", "NT$B", {FC[y]: f"=$D${R['kx']}*{P['fh']}*{FC[y]}{R['b']}*{FC[y]}{R['h']}*{FC[y]}{R['fxr']}" for y in BU})
    row("xc", "  其中晶片與封裝轉售＋服務", "NT$B", {FC[y]: f"={FC[y]}{R['trn']}-{FC[y]}{R['xh']}" for y in BU})
    row("ppu", "每顆世芯營收（校準後）", "US$K", {FC[y]: f"=IF({FC[y]}{R['b']}>0,{FC[y]}{R['asu']}/{FC[y]}{R['b']},0)" for y in BU})
    sec_("B. 其他營收（NRE＋非 Trainium 量產；FY2027 起簡化成長率）")
    row("oth", "其他（NRE＋非 Trainium 量產）", "NT$B", {"C": "=a25_oth", "D": "=a26_oth", "E": f"=D{r}*(1+{p('go27')})", "F": f"=E{r}*(1+{p('goL')})", "G": f"=F{r}*(1+{p('goL')})"})
    row("rev", "營收", "NT$B", {FC[y]: f"={FC[y]}{R['oth']}+{FC[y]}{R['trn']}" for y in FYS})
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("tsh", "Trainium ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['trn']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("osh", "其他 ÷ 營收（對照公司 NRE 10–20%）", "%", {FC[y]: f"={FC[y]}{R['oth']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    sec_("C. 毛利與損益")
    row("gpa", "Trainium 毛利 ＝（m_trn＋情境調整）× 晶片 ＋ m_hbm × 代購 HBM", "NT$B",
        {"D": f"={P['mc']}*D{R['xc']}+{P['mh']}*D{R['xh']}", **{FC[y]: f"=({P['mc']}+{p('gta')})*{FC[y]}{R['xc']}+{P['mh']}*{FC[y]}{R['xh']}" for y in F3}})
    row("gmx", "其他營收毛利率（FY2026 殘差，之後沿用＋調整）", "%", {"C": f"=a25_gp/C{R['oth']}", "D": f"=(a26_gp-D{R['gpa']})/D{R['oth']}", **{FC[y]: f"=$D{r}+{P['gma']}" for y in F3}}, "0.00%",
        note="Derived：讓 FY2026 毛利等於錨點")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{FC[y]: f"={FC[y]}{R['gmx']}*{FC[y]}{R['oth']}+{FC[y]}{R['gpa']}" for y in F3}})
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("opx", "營業費用", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+MAX(g_opx_min,{P['phi']}*{FC[y]}{R['g']}))+{P['opxa']}*{FC[y]}{R['rev']}" for y in F3}})
    row("oi", "營業利益", "NT$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "NT$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={P['nonop']}" for y in F3}})
    row("pti", "稅前利益", "NT$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS})
    row("tax", "所得稅", "NT$B", {"C": "=a25_tax", "D": "=a26_tax", **{FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}})
    row("ni", "歸屬母公司淨利", "NT$B", {FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}" for y in FYS})
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: (f"=q_sh*(1+sh_g)" if y == 2027 else f"={FC[y - 1]}{r}*(1+sh_g)") for y in F3}}, "0.00000",
        note="FY2027 起以 GDR 後股數 86.21M 為基礎")
    row("eps", "EPS（基本）", "NT$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    for y in FADE:      # FY2030–32：收斂法（上游只到 2029；非驅動因子推導）
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({p('fg')}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({p('fm')}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['shr']}"] = f"={pc}{R['shr']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['shr']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.00"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("shr", "0.00000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 世芯-KY 損益（十億新台幣；EPS 新台幣、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（三情境相同），FY2027–FY2029 由驅動因子推導，FY2030–FY2032 為收斂法。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數基準（tsmc 情境 A）；世芯 Trainium 份額 55%；2 奈米每顆營收 ×1.5；Trainium 毛利率 +1pt；其他營收 +15%／+12%",
         "B": "情境 B（平平）：CY2028 顆數 ×0.8、CY2029 ×0.85（tsmc 情境 B）；份額 50%／45%；2 奈米 ×1.2；其他營收 +5%",
         "C": "情境 C（不如預期）：CY2028–29 顆數 ×0.6（tsmc 情境 C）；份額 40%／35%；2 奈米 ×1.0；Trainium 毛利率 −2pt；其他營收 −10%／−5%"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("trn", "  Trainium 量產", "NT$B", "#,##0.0"), ("oth", "  其他（NRE＋非 Trainium 量產）", "NT$B", "#,##0.0"),
        ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"),
        ("eps", "EPS", "NT$", "#,##0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=PL!G{SUMR[('B', 'nim')]}*0.9")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：EPS 差（NT$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC)),
               ("  情境 A：Trainium 由下而上 − Q 頁 Trainium（NT$B）", f"D{PB['A']['trn']}-a26_trn")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 CY2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 世芯 Trainium 量產規格（情境 A）。每顆核心＝Broadcom 模型 pa_trn 扣 CY2026 HBM（與 Marvell 模型同口徑，唯讀）；顆數：百萬顆（日曆年；與 HBM、台積電、Broadcom、Marvell 模型同一份）；每顆價：US$K"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
UN.cell(ur, 1, "1. AWS Trainium（情境 A；世芯營收口徑＝晶片與封裝轉售，基準不含 HBM，乘 k_x）").font = BOLD; ur += 1
for key, lab, unit, f, fmt, note in (
        ("u", "AWS Trainium 市場顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['u']}", "0.00", "tsmc 成品 In 頁"),
        ("s", "世芯份額：AWS Trainium", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "In 頁 B 節（Assumed）"),
        ("b", "世芯顆數：AWS Trainium", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['b']}", "0.000", "市場層顆數 × 份額"),
        ("c", "每顆核心（不含 HBM）", "US$K", lambda y: f"=PL!{FC[y]}{A_['c']}", "0.00", "CY2028 起 × 2 奈米係數（情境 A 1.5）"),
        ("p", "每顆世芯營收（校準後）", "US$K", lambda y: f"=PL!{FC[y]}{A_['ppu']}", "0.00", "（核心＋代購比例 × HBM）× k_x"),
        ("asu", "世芯 Trainium 量產營收（美元）", "US$B", lambda y: f"=PL!{FC[y]}{A_['asu']}", "0.00", "—")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit); UR[key] = ur
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "2. 份額加總檢查（Marvell＋世芯 ≤ 100%；Marvell 份額讀 Marvell 成品 In 頁）").font = BOLD; ur += 1
for s in SC:
    UN.cell(ur, 1, f"Marvell＋世芯：AWS Trainium（情境 {s}）"); UN.cell(ur, 2, "%")
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f"=mv_trn_{min(y, 2028) % 100}+PL!{FC[y]}{PB[s]['sh']}"; UN[f"{cyc(y)}{ur}"].number_format = "0.0%"
    UN.cell(ur, 7, "其餘為 Annapurna 自行設計（COT）或其他夥伴" if s == "A" else "同上"); UR["sum_" + s] = ur; ur += 1
UN.cell(ur, 1, "世芯 Trainium 份額上限（1 − Marvell 基準）"); UN.cell(ur, 2, "%")
for y in BU:
    UN[f"{cyc(y)}{ur}"] = f"=1-mv_trn_{min(y, 2028) % 100}"; UN[f"{cyc(y)}{ur}"].number_format = "0.0%"
UN.cell(ur, 7, "In 頁份額高值受此上限限制"); UR["cap"] = ur; ur += 2
UN.cell(ur, 1, "3. 校準值（只用 FY2026 已發生的實績、當季指引與公司全年 3 奈米說法；不用共識、不用 FY2027 起的目標）").font = BOLD; ur += 1
for key, lab, f, fmt, note in (("kx", "校準係數 k_x", f"=PL!D{A_['kx']}", "0.000", "公司 FY2026 3 奈米營收 ÷ 未校準由下而上"),
                               ("gmx", "其他營收毛利率（FY2026 殘差）", f"=PL!D{A_['gmx']}", "0.00%", "讓 FY2026 毛利等於錨點"),
                               ("dc26", "FY2026 3 奈米營收（公司）", "=dc26", "0.00", "US$B")):
    UN.cell(ur, 1, lab); UN[f"C{ur}"] = f; UN[f"C{ur}"].number_format = fmt; UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：世芯每顆營收只有 Trainium 模組價值的一部分（k_x 約 0.4），顆數與份額決定 FY2027 規模；FY2028 為 2 奈米換代，每顆營收係數與份額是 A、C 情境的主要差異。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：世芯不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照聯發科、Amkor；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-11-05 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差在世芯份額、Trainium 毛利率與其他營收成長（顆數情境係數從 CY2028 開始）"
CY0 = O + 12
V.cell(CY0, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
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
V.cell(R30, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：companiesmarketcap 年底 P/E（2016 為負值，剔除）；自算＝digrin 12 月未調整股價 ÷ HiStock EPS（2016–2017 EPS 未找到）
PE_CMC = [None, 16.5, 15.4, 30.6, 43.3, 46.0, 29.9, 70.6, 39.1, 50.4]
PX_DEC = [31.25, 90.5, 69.0, 234.5, 621.0, 1020.0, 788.0, 3275.0, 3280.0, 3510.0]
EPS_H = [None, None, 4.22, 7.20, 13.61, 21.34, 25.69, 45.47, 81.34, 69.18]
PE_SELF = [f"={p_}/{e_}" if e_ else None for p_, e_ in zip(PX_DEC, EPS_H)]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；負值剔除）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "companiesmarketcap", "自算"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f'=IF(COUNT(C{rr}:D{rr})=0,"",AVERAGE(C{rr}:D{rr}))').number_format = "0.00"
    V.cell(rr, 3, f'=IF(B{RB0 + 5 + i}="","",B{RB0 + 5 + i})').number_format = "0.00"
    V.cell(rr, 4, f'=IF(C{RB0 + 5 + i}="","",C{RB0 + 5 + i})').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2017–2025 共 9 年；2016 為 −8.41 負值，剔除）"},
    {"label": "自算（digrin 12 月未調整股價 ÷ HiStock EPS）", "vals": PE_SELF, "tag": "Derived（價格為月資料的 12 月值；EPS：HiStock 2018–2025；2016–2017 EPS 未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。stockanalysis 2021–2025 年底 P/E（48.36、30.84、73.13、41.01、50.79）與自算相近，只作核對。十年 P/E 多為 NRE 比重高、營收規模小的時期（中位數約 40 倍），世芯轉為 turnkey 量產為主後，倍數未必適用。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22：DCF 與 Comparable 各半；DCF 負值取 0；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# v0.2：標準版 DCF（FCFF＋WACC；tools/fcff_dcf.py）。DCF 明細在 DCF 頁；EBIT：FY2027–29＝推導區塊營業利益；FY2030–32（收斂法只有淨利）由淨利還原
SCN_NAME = {"A": "順利", "B": "平平", "C": "不如預期"}
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（v0.2）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；金額：十億新台幣")
DCFW["A2"] = ("EBIT：FY2027–FY2029 取各推導區塊「營業利益」（TIFRS；員工酬勞與股票薪酬已在營業費用內，不另扣）；FY2030–FY2032（收斂法只有淨利）＝歸屬母公司淨利 ÷（1 − 稅率）− 營業外淨額 nonop_f。"
              "稅率＝In 頁 tax。營業外收支（利息收入為主）不進 NOPAT，現金與投資另在評價時加入。折舊、資本支出、營運資金：近三年（FY2023–FY2025）占營收平均，資本支出 FY2029 起線性收斂到 fc_cx_ss。")
DCF_NEXT = [4]
def dpaths(blocks, pre, tax="tax", nonop="nonop_f", lab=None, names=SC):
    out = []
    for s_ in names:
        bl = blocks[s_]
        ebit = {y: (f"{pre}{FC[y]}{bl['oi']}" if y <= 2029 else f"={pre}{FC[y]}{bl['ni']}/(1-{tax})-{nonop}") for y in YRS6}
        _r26, _r25 = f"{pre}D{bl['rev']}", f"{pre}C{bl['rev']}"   # 第三輪修正：stub 期用含 stub 期會計年度 FY2026（錨點）FCFF × stub_frac，不用 FY2027 全年
        out.append(dict(name=s_, label=(lab or "情境") + f" {s_}（{SCN_NAME[s_]}）", rev={y: f"{pre}{FC[y]}{bl['rev']}" for y in YRS6}, rev0=f"{pre}D{bl['rev']}",
                        ebit=ebit, tax={y: tax for y in YRS6}, ebit_note="FY2030 起由淨利還原",
                        stub_fcff=f"({pre}D{bl['oi']}*(1-{tax})+{_r26}*(fc_da-fc_cx-fc_cx_d)-fc_wc*({_r26}-{_r25}))*stub_frac"))
    return out
def ddetail(paths, title):
    R_ = FD.write_detail(DCFW, DCF_NEXT[0], paths, list(YRS6), 1.0, conv_from=2029, title=title, fmt="#,##0.00", stub=True)
    DCF_NEXT[0] = R_["end"] + 1
    return R_
RD_ = ddetail(dpaths(PB, "PL!"), "三情境 FCFF 明細（十億新台幣；會計年度＝日曆年；評價日 2026-12-31）")
DCF_NEXT[0] = FD.write_checks(DCFW, DCF_NEXT[0], nc_bs="39.405+13.682+2.872-0.28979-0.01769",
                              nc_note="2026-06-30 資產負債表（NT$B；stockanalysis 轉載）：現金 39.405＋短期投資 13.682＋長期投資 2.872 − 租賃 0.290 − 非控制權益 0.018（2026-10-08 發放股利改在 stub 期扣除）") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 34
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 12
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β，Blume 調整，重新加槓桿")
for _j, _h in enumerate(["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"], 1): V.cell(SDCF["h"], _j, _h)
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"])
assert SDCF["dcf"] <= 24
_PE = lambda c, y: f"PL!{FC[y]}{SUMR[(dict(zip('BCD', SC))[c], 'eps')]}"
EPSR = {}   # v0.2：舊 0.1 節 EPS 列已移除（ROWMAP 相容）
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("創意（GUC）", "3443.TW", 78.8, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 102.85（2026-10-08 頁面，價格 8360）→ 前瞻 EPS 81.28；P/E＝2026-11-04 基準 87.53 × Yahoo Finance 日線收盤 2026-11-05 6405 ÷ 2026-11-04 收盤 7115（共識 EPS 不變）"),
         ("智原（Faraday）", "3035.TW", 36.13, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 41.6（2026-10-16 頁面，價格 197.5）→ 前瞻 EPS 4.748；P/E＝2026-11-04 基準 40.13 × Yahoo Finance 日線收盤 2026-11-05 171.5 ÷ 2026-11-04 收盤 190.5（共識 EPS 不變）"),
         ("聯發科（MediaTek）", "2454.TW", 33.44, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 38.75（2026-10-22 頁面，價格 4010）→ 前瞻 EPS 103.5；P/E＝2026-11-04 基準 35.9 × Yahoo Finance 日線收盤 2026-11-05 3460 ÷ 2026-11-04 收盤 3715（共識 EPS 不變）"),
         ("博通（Broadcom）", "AVGO", 16.76, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 18.74（2026-10-30 頁面，價格 326.08）→ 前瞻 EPS 17.4；P/E＝2026-11-04 基準 18.71 × Yahoo Finance 日線收盤 2026-11-05 291.53 ÷ 2026-11-04 收盤 325.48（共識 EPS 不變）"),
         ("邁威爾（Marvell）", "MRVL", 23.66, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 42.09（2026-10-27 頁面，價格 237.6）→ 前瞻 EPS 5.645；P/E＝2026-11-04 基準 28.31 × Yahoo Finance 日線收盤 2026-11-05 133.55 ÷ 2026-11-04 收盤 159.8（共識 EPS 不變）")]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "台灣 ASIC 設計服務（創意、智原）＋ASIC 對手（聯發科、博通、邁威爾）；不含 A 股（非 A 股公司，Andy 2026-10-25 11:53）")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{dict(zip('BCD', SC))[c]}*(2/12*{_PE(c, 2027)}+10/12*{_PE(c, 2028)})" for c in "BCD"), None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "#,##0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B35", None, None], "#,##0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "#,##0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", "stockanalysis：16 位分析師平均 5,803，2026-10-23 擷取")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-11-05 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{_PE(c, 2028)}>=MAX({PE10}),1,IF(px/{_PE(c, 2028)}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{_PE(c, 2028)})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；低於十年最小值記為 0%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF（v0.2 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 新台幣 10 年期公債殖利率)、RONIC＝WACC；"
            "ke＝rf＋β ×（ERP 5%＋台灣國家風險溢酬），β 為自算台灣半導體產業無槓桿 β（同業 13 家中位數、Blume 調整）依市值 D/E 重新加槓桿；EV 加淨現金與長期投資 ÷ 股數，以 ke 推到 12 個月後並扣期間股利（DCF 頁、tools/fcff_dcf.py）。"
            "折舊、資本支出、營運資金（含客戶預收）取 FY2023–25 占營收平均（資本支出 FY2029 起收斂到維持水準）。情境價值負值取 0。"
            "Comparable：創意、智原、聯發科、博通、邁威爾前瞻 P/E 中位數；創意 103x 為極端值，中位數不受影響；世芯本身前瞻 P/E 約 17x，遠低於同業，反映單一客戶集中與 turnkey 低毛利，Comparable 偏向高估。"
            "主要偏誤：Trainium 顆數受市場層 CY2027 2.5 百萬顆限制（UBS 估 Trainium 3 即 2.8 百萬顆），FY2030 起只用收斂法，DCF 偏低估；2 奈米 COT 模式若使每顆營收下降，則偏高估。"
            "情境對應：A＝順利（tsmc 顆數 A、份額 55%）、B＝平平（顆數 B、份額 50%／45%）、C＝不如預期（顆數 C、份額 40%／35%）")
for r_ in (T["dcftp"], 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，k_x 與其他營收毛利率在每個案例內重算）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("Trainium 份額 −10pt", "CY2027 起世芯占 Trainium −10 個百分點", dict(shp="(-0.10)")), ("Trainium 份額取高值", "CY2027 75%、CY2028 起 85%（上限 1 − Marvell）", dict(shs="_hi")),
         ("顆數 ×0.8", "CY2027 起世芯顆數 ×0.8", dict(uf="uf_lo")), ("顆數 ×1.2", "×1.2（市場層 Trainium 顆數低於 UBS 估計）", dict(uf="uf_hi")),
         ("2 奈米每顆營收 ×2.5", "CY2028 起（美系外資：Trainium 4 每顆超過 1 萬美元）", dict(pfc="pfc_A_hi")), ("2 奈米每顆營收 ×0.8", "CY2028 起（COT 模式、客戶自購晶圓）", dict(pfc="pfc_A_lo")),
         ("代購 HBM 0.6", "比照 Broadcom 模型（k_x 重算）", dict(fh="f_al_hi")),
         ("Trainium 毛利率 10%", "基準 13%（＋情境 A 調整 1pt）", dict(mc="m_trn_lo")), ("Trainium 毛利率 16%", "", dict(mc="m_trn_hi")),
         ("其他營收 情境 C 路徑", "FY2027 −10%、之後 −5%", dict(go27="g_o27_C", goL="g_oL_C")),
         ("其他營收毛利率 −3pt", "FY2027 起", dict(gma="gmadj_lo")), ("營業費用彈性 0.6", "基準 0.3", dict(phi="phi_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi")), ("匯率 30", "基準 32（Trainium 美元營收換算）", dict(fx="fx_lo")),
         ("稅率 25%", "基準 20%", dict(tax="tax_hi")), ("營業外 1.0", "基準 1.6", dict(nonop="nonop_f_lo"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0"), ("D", "gm", 2027, "0.00%"), ("E", "eps", 2027, "#,##0.00"),
                           ("G", "rev", 2028, "#,##0"), ("H", "gm", 2028, "0.00%"), ("I", "eps", 2028, "#,##0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"F{i}"] = f"=E{i}/E$4-1"; SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c in "FJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 40; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 B", "顆數 B；份額 50%／45%；2 奈米 ×1.2；其他營收 +5%", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "顆數 C；份額 40%／35%；Trainium 毛利率 −2pt；其他營收 −10%", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 三情境差在世芯份額、Trainium 毛利率與其他營收；顆數情境 B、C 的差異從 FY2028 起。FY2028 最大的單一變數是 2 奈米每顆營收係數（價格遠高於 3 奈米 vs COT 模式），其次是份額與顆數。")
SENS_NOTE = nn + 3
TPS0 = nn + 5
# v0.2：舊 tp_r／tp_g／tp_w_dcf 七案改為標準版 DCF 參數案例（tools/fcff_dcf.write_sens：閉式公式與 DCF 頁明細同式；加權目標價）
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", beta_du="beta_du", usd=True,
                     title="目標價參數敏感度（v0.2 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）")
TPS_LAST = _tvs[1]
assert _tvs[2] <= 40, _tvs
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 44
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司說法、共識、券商。只作事後對照，不為對上數字而調參數。券商 FY2027–28 估計與 UBS 顆數為 Interested-party，不用於校準"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
RC.cell(rr_, 1, "A. FY2026 錨點 vs 公司說法與共識").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "stockanalysis（14 位；2026-08-13）", "共識資料日在 3Q26 營收公布前（2026-08-13）；區間 35.4–86.1", fmt="#,##0.0")
rec("FY2026 EPS vs 共識", "=cons_eps26", pl("eps", 2026), "stockanalysis（2026-08-13）", "共識區間 103.06–163.01；美系外資甲 132.94")
rec("FY2026 毛利率 vs 公司低至中 20%（百分點）", "=co_gm26", pl("gm", 2026), "2Q26 法說", "", fmt="0.0%", pt=True)
rec("FY2026 Trainium（美元）vs 美系外資甲 13 億美元", "=b1_trn26", pl("asu", 2026), "旺得富 2026-08-17", "模型取公司「超過 15 億美元」", fmt="0.00")
RC.cell(rr_, 1, "B. Trainium vs 賣方（市場層顆數 × 份額口徑）").font = BOLD; rr_ += 1
rec("FY2027 Trainium（美元）vs 美系外資甲 22 億美元", "=b1_trn27", pl("asu", 2027), "旺得富 2026-08-17", "", fmt="0.00")
rec("FY2027 Trainium（美元）vs 美系外資乙 AI ASIC 38 億美元", "=b2_ai27", pl("asu", 2027), "旺得富 2026-08-17", "美系外資乙口徑含其他 AI ASIC", fmt="0.00")
rec("FY2028 Trainium（美元）vs 美系外資 Trainium 4 40 億美元", "=b1_trn4", pl("asu", 2028), "旺得富、時報資訊", "對照值不含 3 奈米尾量；美系外資丙情境 90–120 億美元", fmt="0.00")
rec("CY2027 市場層 Trainium 顆數 vs UBS Trainium 3 出貨", "=ubs_trn27", pl("u", 2027), "理財周刊轉述 UBS", "市場層含全部 Trainium 世代仍低於 UBS 的 Trainium 3；修訂應由市場層提出", fmt="0.00")
rec("CY2027 世芯顆數 ÷ UBS Trainium 3 出貨", None, lambda s: f"=PL!E{PB[s]['b']}/ubs_trn27", "推導", "世芯為 Trainium 3 主供應商時，此值應接近 100%", fmt="0.0%")
rec("FY2027 每顆世芯營收（校準後）", None, pl("ppu", 2027), "推導", "US$K；k_x × 核心", fmt="0.00")
RC.cell(rr_, 1, "C. EPS 與營收 vs 券商（校準未使用）").font = BOLD; rr_ += 1
rec("FY2027 EPS vs 美系外資甲", "=b1_eps27", pl("eps", 2027), "旺得富 2026-08-17", "")
rec("FY2027 EPS vs 理財周刊作者", "=aut_eps27", pl("eps", 2027), "理財周刊 2026-09-06", "")
rec("FY2027 營收 vs 理財周刊作者", "=aut_rev27", pl("rev", 2027), "理財周刊 2026-09-06", "", fmt="#,##0.0")
rec("共識前瞻 12 月 EPS vs 模型（情境加權 2／12 × FY2027＋10／12 × FY2028）", "=cons_feps", lambda s: f"=Val!B{T['feps']}", "stockanalysis 前瞻 P/E 推得", "三欄同值（情境加權）")
RC.cell(rr_, 1, "D. 校準與份額").font = BOLD; rr_ += 1
rec("k_x（校準係數）", None, lambda s: f"=PL!D{PB[s]['kx']}", "公司 FY2026 3 奈米營收 ÷ 未校準", "< 1：世芯營收為晶片與封裝轉售，不含模組其他價值；CY2026 只有約 7 個月量產", fmt="0.000")
rec("其他營收毛利率（FY2026 殘差）", None, lambda s: f"=PL!D{PB[s]['gmx']}", "FY2026 毛利殘差", "1Q26 綜合毛利率 50.2%（NRE 比重高）", fmt="0.00%")
rec("CY2028 Marvell＋世芯 Trainium 份額", 1.0, lambda s: f"=Units!E{UR['sum_' + s]}", "mrvl 成品", "應 ≤ 100%", fmt="0.0%", pt=True)
rec("FY2027 其他營收 ÷ 營收 vs 公司 NRE 10–20%（百分點）", "=co_nre", pl("osh", 2027), "2Q26 法說", "其他營收含非 Trainium 量產，應高於 NRE 比重", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以世芯 Trainium 營收 ÷ 部署 GW 對照 Tokenomics 非 NVIDIA 加速器比例"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 80
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝世芯新聞稿；法說內容多為媒體與逐字稿轉述；券商、資料商、媒體皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：2Q26 財報新聞稿（美元營收、毛利、稅前、淨利；製程、應用、地區占比）", PR2),
        ("媒體：工商時報 2026-08-14（2Q26、1H26 新台幣損益：營收、毛利、營業利益、稅前、淨利、EPS）", CT2),
        ("逐字稿：Investing.com 2Q26 法說（毛利率指引、NRE 10–20%、營業費用、2 奈米）", INV2),
        ("媒體：鉅亨網 2026-08-14（2Q26 法說：2 奈米價量、COT 模式、2027 年 3 奈米上修空間）", CNY2),
        ("二手：方格子 1Q26 法說摘要（1Q26 營收、應用、製程、地區）", VOC1), ("二手：方格子 2Q26 財報摘要（毛利率 50.2% → 34.8%、製程占比）", VOC2),
        ("媒體：經濟日報 2026-03-06（FY2025 營收、毛利率、營業利益率、淨利、EPS；4Q25）", UDN25),
        ("媒體：科技新報（2026 年 9 月營收、前 9 月累計）", SEP26), ("媒體：科技新報（2025 年 9 月營收）", SEP25),
        ("媒體：經濟日報 2026-07-07（GDR 掛牌）", GDR),
        ("券商轉述：時報資訊 2026-08-14（美系外資 5,088；3Q26 毛利率約 18%；Trainium 4 2028 年 40 萬顆、每顆超過 1 萬美元）", CTB),
        ("券商轉述：旺得富 2026-08-17（三家美系外資：EPS、Trainium 3／4 營收、目標價）", WR3),
        ("券商轉述：工商時報 2026-09-23（歐系外資 6,784、EPS 上修幅度）", EU9),
        ("媒體：理財周刊 2026-09-06（作者 FY2026–27 估計；轉述 UBS Trainium 3 出貨）", MW9),
        ("資料商：stockanalysis（現價 10-22、股數、近 12 月與前瞻 P/E、共識目標價）", SA), ("資料商：stockanalysis 共識頁（FY2026 共識）", SAF),
        ("資料商：stockanalysis 財報與比率頁（FY2021–25 損益、ROE、年底 P/E）", SAI + "；https://stockanalysis.com/quote/tpe/3661/financials/ratios/"),
        ("資料商：stockanalysis 同業前瞻 P/E（3443、3035、2454、AVGO、MRVL）", "https://stockanalysis.com/quote/tpe/3443/；https://stockanalysis.com/quote/tpe/3035/；https://stockanalysis.com/quote/tpe/2454/；https://stockanalysis.com/stocks/avgo/；https://stockanalysis.com/stocks/mrvl/"),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/alchip-technologies/pe-ratio/"),
        ("資料商：digrin 月股價（12 月未調整值）", "https://www.digrin.com/stocks/detail/3661.TW/price"),
        ("資料商：HiStock 歷年 EPS（2018–2025）與季 EPS", "https://histock.tw/stock/3661/%E6%AF%8F%E8%82%A1%E7%9B%88%E9%A4%98"),
        ("市場層：tsmc、avgo、mrvl 成品與 markets/memory UNITS、SCEN_UNITS（本 repo，唯讀）", f"{rel(TSMC_F)}；{rel(AVGO_F)}；{rel(MRVL_F)}；markets/memory/market.py"),
        ("待補（人工）：3Q26 財報與法說（約 11 月上旬）；NRE 與量產營收金額；HBM 採購方式；2 奈米每顆營收與 COT 分工；2016–2017 EPS 與 2016–2020 淨利率（十年平均）；GDR 募資金額；前瞻 P/E 十年歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("世芯-KY（3661.TW）收支與估值模型 v0.1.1（AI 半導體 Project：運算層，ASIC 設計服務＋量產）　2026-10-25", True),
 ("版本：v0.1（初版）。Trainium 顆數讀台積電模型（與 markets/memory 同一份），每顆規格讀 Broadcom 模型 pa_trn（扣 HBM 口徑與 Marvell 模型相同），份額與 Marvell 合計不超過 100%；其他營收（NRE＋非 Trainium 量產）用簡化成長率。本版未做先驗。", False),
 ("模型命題：世芯的營收由 NRE 轉為 Trainium 量產（turnkey，含晶圓與封裝轉售、毛利率低）。FY2026 3 奈米約 15 億美元、3Q26 單季營收 260 億元；以市場層 Trainium 顆數 × 世芯份額 × 每顆營收往前推，FY2027–28 的營收與 EPS 有多少，2 奈米換代的份額與每顆營收差異有多大？", False),
 ("核心問題：在市場層 Trainium 顆數（CY2027 2.5 百萬顆）與 Marvell 份額口徑下，世芯 FY2027–28 的 EPS 與賣方差多少；FY2028 的 2 奈米換代靠份額還是每顆營收；現價約 17 倍前瞻 P/E 隱含哪一條路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。FY2025A 實績、FY2026 錨點、FY2027–FY2029 驅動因子推導、FY2030–FY2032 收斂法。金額：十億新台幣；EPS：新台幣（基本）。口徑：歸屬母公司淨利（TIFRS）。3Q26 財報約 11 月上旬公布，本版只納入 3Q26 營收實績。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("Trainium 量產：tsmc 成品的 AWS Trainium 顆數 × 情境係數（CY2028–29；tsmc／markets/memory 同字母）× 世芯份額（Assumed；CY2026 30%，CY2027 A／B／C 55%／50%／40%，CY2028 起 55%／45%／35%）×（每顆核心 6.4K × 2 奈米係數 1.5／1.2／1.0＋代購比例 0 × HBM 金額）× k_x（FY2026 以公司「3 奈米超過 15 億美元」校準）× 匯率 32。", False),
 ("其他營收：NRE＋非 Trainium 量產（中國 ADAS、Arm CPU、日本客戶等）由 FY2026 錨點依情境成長率推。", False),
 ("毛利：Trainium 量產毛利率 13%（Assumed；情境 A +1pt、C −2pt）；其他營收毛利率為 FY2026 殘差（Derived）後沿用。營業費用彈性 0.3、年增下限 8%；營業外 1.6、稅率 20%、股數 86.21M（GDR 後）。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績（月營收）＋3Q26 指引（毛利率低至中 20%、費用與 2Q 相近）＋4Q26 假設（營收季增 15%、Trainium 為全年 15 億美元餘額）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（v0.2 起為標準版：無槓桿自由現金流 FCFF＋WACC，FY2027–32 明確期＋永續，三情境權重 0.3／0.5／0.2，情境價值負值取 0；明細見 DCF 頁）＋0.5 × Comparable 目標價（創意、智原、聯發科、博通、邁威爾前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到，B、C 依工作單以 FY2029 × 0.8 代替。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自世芯新聞稿與媒體轉述財報；法說內容多為二手轉述；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 公司未揭露客戶名稱、NRE 與量產金額、HBM 採購方式，Trainium 份額、拆分與毛利率為 Assumed；(2) 市場層 CY2027 Trainium 2.5 百萬顆低於 UBS 的 Trainium 3 2.8 百萬顆，修訂應由市場層提出，不在本模型自改；(3) 2 奈米每顆營收係數為 Assumed（公司：價格遠高於 3 奈米；但主要客戶採 COT 模式），放在情境與 Sens；(4) 單一客戶集中度高（分析師估 60–70% 以上），情境 C 只反映份額與顆數下修，未含失去客戶；(5) 3Q26 毛利率與費用為指引值；(6) 十年 P/E 多為 NRE 主導期；2016 為負值、2016–2017 EPS 未找到；(7) R4 產業部署路徑未建（Recon 預留列）；(8) v0.2 標準版 DCF 的折舊、資本支出、營運資金以 FY2023–25 占營收平均推估，當時營收以 NRE 與舊量產為主，Trainium turnkey 量產比重提高後，這三個比例可能偏離（Sens 測資本支出 ±1pt、營運資金 0%／15%）。", False), ("", False),
 ("情境對照表（A／B／C 讀同字母的行業情境；公司特有變數 A 最有利、C 最不利）", True),
 ("行業：Trainium 顆數係數＝tsmc In 頁 um28／um29（＝markets/memory SCEN_UNITS：A 1.0／1.0、B 0.8／0.85、C 0.6／0.6）；HBM 單價＝avgo In 頁 hbm_A／B／C（基準代購比例 0，不影響營收）。不讀 markets/equipment、markets/network（世芯無設備、網通營收）。", False),
 ("公司特有：世芯 Trainium 份額 CY2027 55%／50%／40%、CY2028 起 55%／45%／35%；2 奈米每顆營收係數 1.5／1.2／1.0；Trainium 毛利率調整 +1／0／−2pt；其他營收年增 FY2027 +15%／+5%／−10%、FY2028–29 +12%／+5%／−5%。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170


# ---------------- v0.1.1：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；各頁只在現有最後一列之後追加，不位移既有列 ----------------
_up = next(r for r in range(1, V.max_row + 1) if V.cell(r, 1).value == "上檔空間：加權目標價")
GRADE_ROW = V.max_row + 2; REC_ROW = GRADE_ROW + 1
V.cell(GRADE_ROW, 1, "上檔空間級距")
V.cell(GRADE_ROW, 2, f'=IF(B{_up}>0.5,"大於 50%",IF(B{_up}>0.25,"25%–50%",IF(B{_up}>=0,"0%–25%",IF(B{_up}>=-0.25,"0%～−25%",IF(B{_up}>=-0.5,"−25%～−50%","小於 −50%")))))')
V.cell(GRADE_ROW, 5, f"以「上檔空間：加權目標價」（B{_up}）分六級（Andy 2026-10-25 14:00）")
V.cell(REC_ROW, 1, "投資建議")
V.cell(REC_ROW, 2, f'=IF(B{_up}>0.5,"強力買進",IF(B{_up}>0.25,"買進",IF(B{_up}>=-0.25,"中立",IF(B{_up}>=-0.5,"放空","強力放空"))))')
V.cell(REC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = ['目標價 NT$5,926（DCF 3,503、Comparable 8,348 各半），比現價 3,495（2026-10-22）高 69.6%，級距 大於 50%；Comparable 單獨為 +138.9%。主因是倍數：本模型 12 個月後前瞻 EPS 215.44 只比共識前瞻 EPS 203.79 高 5.7%，但同業（創意、智原、聯發科、博通、邁威爾）前瞻 P/E 中位數 38.8 倍，現價只隱含 17.2 倍、共識目標價 5,803 隱含 28.5 倍（比同業低 26.5%）；DCF 3,503 與現價幾乎相同（+0.2%），上檔全部來自 Comparable，而 Comparable 未扣世芯單一客戶集中與 turnkey 低毛利的折價，偏高估。', '情境分歧很大：FY2028 EPS A／B／C 為 338.6／196.6／103.5，DCF 價格帶 1,226～5,691，差在 Trainium 顆數與世芯份額（CY2028 起 55%／45%／35%）；目標價對 DCF 權重也敏感，權重 0.7 時降 16.4%（約 4,957，上檔 +41.8%）。', 'Trainium 份額是 FY2027 EPS 最敏感的變數（取高值 +20.7%、−10pt 時 −10.4%）；FY2028 起再加上 2 奈米每顆營收（×2.5 時 FY2028 營收 +58.9%、COT 模式 ×0.8 時 −41.2%）（Sens 頁）。下一個檢驗點是 3Q26 財報與法說（約 2026-11 上旬）的 Trainium 3 量產進度。']   # 重點觀察（與 README.md 相同文字；數字取自重算後的 Val／PL／Sens 頁）
_rr = RD.max_row + 1
RD.cell(_rr, 1, 'v0.1.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁末新增「上檔空間級距」「投資建議」兩列（Andy 2026-10-25 14:00、14:17 固定規則），README 頁末新增投資建議與重點觀察三點；既有儲存格位置與數值不變。')
RD.cell(_rr + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"').font = Font(bold=True)
for _j, _t in enumerate(OBS, 1): RD.cell(_rr + 1 + _j, 1, f"重點觀察 {_j}：" + _t)
if os.environ.get("RATING_ROWMAP"):   # 除錯用：輸出新增列號（供撰寫 html.json）；CI 不設此變數
    print("ROWMAP", GRADE_ROW, REC_ROW, _rr, _rr + 1, _rr + 2)


# ---------------- v0.1.2（2026-10-28）：分部原則重建（Q 頁第 33 列起追加；不位移既有列，數值不變） ----------------
Q.cell(33, 1, "官方報告分部與應用別（FY2025；分部原則第 3 條）").font = BOLD
_S = "世芯 2025 第四季暨全年財報新聞稿（semiwiki 轉載，2026-03）"
for _r, _lab, _v, _tag, _s1, _s2 in (
        (34, "單一營運部門（ASIC 設計服務）營收", 30.926, "Verified", "公司 FY2025 財報：單一應報導部門（309.3 億元，經濟日報 2026-03-06）", "與第 4 列 FY2025 營收同一數字"),
        (35, "  HPC 與 AI ASIC", "=B34*0.83", "Verified", _S + "：83%", "比重為公司揭露；金額為推得"),
        (36, "  消費性", "=B34*0.12", "Verified", _S + "：12%", "同上"),
        (37, "  利基與網通", "=B34*0.04", "Verified", _S + "：4%（合計）", "同上"),
        (38, "  其他（差額，四捨五入）", "=B34-B35-B36-B37", "Derived", "—", "比重合計 99%"),
        (39, "檢查：分部合計 − 合併營收", "=B34-B4", "Derived", "—", "應為 0")):
    Q.cell(_r, 1, _lab); Q[f"B{_r}"] = _v; Q[f"B{_r}"].number_format = "#,##0.000"
    if not (isinstance(_v, str) and _v.startswith("=")): Q[f"B{_r}"].font = BLUE
    Q.cell(_r, 8, _tag); Q.cell(_r, 9, _s1); Q.cell(_r, 10, _s2)
_rr = RD.max_row + 1
RD.cell(_rr, 1, "v0.1.2（2026-10-28）：分部原則重建（數值不變）。世芯為單一應報導部門；Q 頁第 33 列起追加 FY2025 部門營收、應用別子列（HPC 與 AI ASIC 83%、消費性 12%、利基與網通 4%，公司財報新聞稿）與核對列「分部合計 − 合併營收」。公司未揭露 NRE／量產金額；推估年仍以 Trainium 量產／其他拆分。")
RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.1.1", "v0.1.2")

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "alchip")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.3", "2026-10-30")  # v0.1.3：In 頁版面（個股模型原則第 8 節；數值不變）

# ---------------- Wattonomics 錨定（決策 S37；Siliconomics v0.13）：In 頁最後三列（wta_27／28／29，數值取 SIL.WTA）＋Sens 頁新案例；既有儲存格不變 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
def _wta_in(wb):
    from openpyxl.styles import Font as _F
    from openpyxl.workbook.defined_name import DefinedName as _DN
    ws = wb["In"]; r = ws.max_row + 2
    ws.cell(r, 1, "Z. Wattonomics 錨定係數（敏感度用；Siliconomics v0.13）").font = _F(bold=True); r += 1
    rows = []; SIL.wta_rows(lambda *a: rows.append(a))
    for a in rows:
        for j, v in enumerate(a, 1):
            if v is not None: ws.cell(r, j, v)
        ws.cell(r, 5).font = _F(color="0000FF")
        wb.defined_names[a[0]] = _DN(a[0], attr_text=f"In!$E${r}"); r += 1
_wta_in(wb)
def _wta_tp_t(SE, r0, ew, eb, T, fwd, fmt, dcf12="dcf12", feps="feps", tpw="tpw", yrs=range(2027, 2033), floor=False):
    """Wattonomics 錨定案例的單一路徑目標價（Sens 頁）：DCF 列先寫 v0.1 股東現金流式，v0.2 起由呼叫端覆寫為標準版 FCFF 明細（DCF 頁）的推到 12 個月後價值；
    Comparable＝tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS。
    ew／eb：年 → 儲存格（案例／情境 B）。回傳 {鍵: 儲存格}。"""
    from openpyxl.styles import Font as _F
    ys = list(yrs)
    def dcf(e):
        cf = lambda y: f"{e(y)}*(1-MIN(tp_g/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+tp_r)^{t}" for t, y in enumerate(ys, 1))
        v = f"{pv}+{cf(ys[-1])}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^{len(ys)}"
        return (f"MAX(0,{v})" if floor else f"({v})") + "*(1+tp_r)"
    for j, lab in enumerate(["目標價（單一路徑）", "Wattonomics 錨定", "情境 B", "基準（情境加權，Val 第 0 節）"]): SE.cell(r0, 1 + j, lab).font = _F(bold=True)
    R = {"dcf": r0 + 1, "fwd": r0 + 2, "comp": r0 + 3, "tp": r0 + 4, "chk": r0 + 5}
    SE.cell(R["dcf"], 1, "DCF 價值推到 12 個月後（×(1＋tp_r)）"); SE.cell(R["dcf"], 2, "=" + dcf(ew)); SE.cell(R["dcf"], 3, "=" + dcf(eb)); SE.cell(R["dcf"], 4, f"=Val!B{T['dcftp']}")
    SE.cell(R["fwd"], 1, "12 個月後前瞻 EPS（權重同 Val 第 0 節）")
    SE.cell(R["fwd"], 2, "=" + "+".join(f"{w}*{ew(y)}" for y, w in fwd)); SE.cell(R["fwd"], 3, "=" + "+".join(f"{w}*{eb(y)}" for y, w in fwd)); SE.cell(R["fwd"], 4, f"=Val!B{T[feps]}")
    SE.cell(R["comp"], 1, "Comparable（相對倍數調整後；tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS）")
    for c in "BC": SE[f"{c}{R['comp']}"] = f"=tp_comp_adj*{c}{R['fwd']}/Val!B{T[feps]}"
    SE[f"D{R['comp']}"] = "=tp_comp_adj"
    SE.cell(R["tp"], 1, "目標價（tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable）")
    for c in "BC": SE[f"{c}{R['tp']}"] = f"=tp_w_dcf*{c}{R['dcf']}+(1-tp_w_dcf)*{c}{R['comp']}"
    SE[f"D{R['tp']}"] = f"=Val!B{T[tpw]}"
    for k in ("dcf", "fwd", "comp", "tp"):
        for c in "BCD": SE[f"{c}{R[k]}"].number_format = "#,##0.00" if k == "fwd" else fmt
    SE.cell(R["chk"], 1, "核對：情境 B 欄 DCF − Val 第 0 節情境 B（應為 0）"); SE.cell(R["chk"], 3, f"=C{R['dcf']}-Val!C{T[dcf12]}").number_format = "0.000000"
    return {"tp_w": f"B{R['tp']}", "tp_b": f"C{R['tp']}", "tp_base": f"D{R['tp']}", "end": R["chk"]}
def _wta_table(SE, r0, cols, rows):
    """小表：cols＝[(欄名, fmt)]；rows＝[(名稱, [儲存格…])]，第一列＝基準、第二列＝情境 B；每個指標後接「相對基準」「相對情境 B」。"""
    from openpyxl.styles import Font as _F
    from openpyxl.utils import get_column_letter as _gl
    SE.cell(r0, 1, "案例").font = _F(bold=True)
    for j, (nm, fmt) in enumerate(cols):
        for k, t in enumerate((nm, "相對基準", "相對情境 B")): SE.cell(r0, 2 + 3 * j + k, t).font = _F(bold=True)
    for i, (nm, cells) in enumerate(rows):
        r = r0 + 1 + i; SE.cell(r, 1, nm)
        for j, (cell, (cn, fmt)) in enumerate(zip(cells, cols)):
            v, d1, d2 = (_gl(2 + 3 * j + k) for k in range(3))
            SE[f"{v}{r}"] = f"={cell}"; SE[f"{v}{r}"].number_format = fmt
            SE[f"{d1}{r}"] = f"={v}{r}/{v}${r0 + 1}-1"; SE[f"{d2}{r}"] = f"={v}{r}/{v}${r0 + 2}-1"
            SE[f"{d1}{r}"].number_format = SE[f"{d2}{r}"].number_format = "0.0%"
    return r0 + len(rows)
WTA_D = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
_w0 = SE.max_row + 3
SE.cell(_w0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
SE.cell(_w0 + 1, 1, WTA_NOTE)
SE.cell(_w0 + 2, 1, '做法：FY(t)＝CY(t)。FY2027 起世芯 Trainium 顆數（AWS Trainium 市場顆數 × 份額）再乘 wta（FY2027＝wta_27 … FY2029＝wta_29）；k_x 由 FY2026 校準不受影響；其他營收（NRE＋非 Trainium 量產）不乘。FY2030–FY2032 沿用 PL 收斂法（營收年增收斂到 fade_gB；淨利率收斂目標＝本案例 FY2029 淨利率 × 0.9，與 PL 情境 B 的 fade_mB 同規則）。')
_t0 = _w0 + 4
SB_W = engine(SE, _t0 + 6, "案例區塊：Wattonomics 錨定（情境 B × wta）", scn="B", wta=WTA_D)
for _y in FADE:   # 淨利率收斂目標：情境 B 規則（FY2029 淨利率 × 0.9）套用在本案例自己的 FY2029
    _c = SE[f"{FC[_y]}{SB_W['nim']}"]; assert "fade_mB" in _c.value; _c.value = _c.value.replace("fade_mB", f"($G{SB_W['nim']}*0.9)")
_ew = lambda y: f"{FC[y]}{SB_W['eps']}"; _eb = lambda y: f"PL!{FC[y]}{SUMR[('B', 'eps')]}"
_W = _wta_tp_t(SE, SB_W["end"] + 1, _ew, _eb, T, [(2027, "2/12"), (2028, "10/12")], '#,##0.00')
# v0.2：DCF 改為標準版（FCFF＋WACC）：本路徑在 DCF 頁另寫一個明細區塊（與三情境同式），覆寫 _wta_tp_t 的舊股東現金流 DCF 列
_RW = ddetail(dpaths({"B": SB_W}, "Sens!", lab="Wattonomics 錨定 × 情境", names=("B",)), "Sens 案例路徑：Wattonomics 錨定（情境 B × wta）FCFF 明細（未進基準）")
_wr = int(_W["tp_w"][1:]) - 3
SE.cell(_wr, 1, "DCF 價值推到 12 個月後（標準版 FCFF＋WACC；×（1＋ke）^t_tgt − 期間股利）")
SE.cell(_wr, 2, f"=DCF!$C${_RW['paths']['B']['v12']}"); SE.cell(_wr, 3, f"=DCF!$C${RD_['paths']['B']['v12']}")
SE.cell(_wr + 4, 1, "核對：情境 B 欄 DCF − Val 第 0.1 節情境 B 價值推到 12 個月後（應為 0）")
_e = lambda s_, y: f"PL!{FC[y]}{SUMR[(s_, 'eps')]}"
_wta_table(SE, _t0, [("FY2027 EPS", "#,##0.00"), ("FY2028 EPS", "#,##0.00"), ("目標價", '#,##0.00')],
           [("基準（Sens 主表基準＝情境 A；目標價＝情境加權）", [_e("A", 2027), _e("A", 2028), _W["tp_base"]]),
            ("情境 B（案例底）", [_e("B", 2027), _e("B", 2028), _W["tp_b"]]),
            ("Wattonomics 錨定（情境 B × wta）", [_ew(2027), _ew(2028), _W["tp_w"]])])
# ---------------- v0.2（2026-11-05）：README 重點觀察改公式、「展示層」「命題回覆」兩段（個股 HTML 版面規格 v1.0；數字皆為公式）、版本說明 ----------------
_tvsh = SDCF["tvsh"]
_TVF = f'IF(MAX(Val!B{_tvsh}:D{_tvsh})>0.8,"終值現值占 EV 超過 80% 的情境："&IF(Val!B{_tvsh}>0.8,"A ","")&IF(Val!C{_tvsh}>0.8,"B ","")&IF(Val!D{_tvsh}>0.8,"C ","")&"（DCF 價值高度依賴終值假設，需留意）。","")'   # 第二輪修正：終值占 EV > 80% 的情境標示（公式判斷）
_USD = next(r for r in range(1, SE.max_row + 1) if SE.cell(r, 1).value == "美元口徑折現")
_DWK = next(r for r in range(_tvs[0], SE.max_row + 1) if str(SE.cell(r, 1).value).startswith("DCF 權重")) - _tvs[0]   # DCF 權重案例（模組 v2 加美元口徑列後下移一列；依標籤定位）   # Sens「美元口徑折現」列（C＝加權目標價、D＝相對基準）
def XF(*parts):
    out = []
    for p_ in parts:
        if isinstance(p_, tuple): out.append(p_[0]); continue
        t_ = p_.replace('"', '""')
        out += [f'"{t_[i:i + 200]}"' for i in range(0, len(t_), 200)]
    return "=" + "&".join(out)
_vr = lambda p_: next(r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").strip().startswith(p_))
_B = lambda k, c="B": f"Val!{c}{T[k]}"
_E = lambda s_, y: f"PL!{FC[y]}{SUMR[(s_, 'eps')]}"
_P = lambda k, s_, y: f"PL!{FC[y]}{PB[s_][k]}"
_CN = [c[0] for c in cases]
_SF = lambda nm, col="F": f"Sens!{col}{5 + _CN.index(nm)}"           # Sens 主表（情境 A 為基準）：F＝FY2027 EPS、J＝FY2028 EPS 相對基準
_TS = lambda k, col="D": f"Sens!{col}{_tvs[0] + k}"                   # 目標價參數敏感度（加權目標價為基準）：D＝低、F＝高（相對基準）
_own, _npeer, _pmed10, _raw = (f"Val!B{_vr(p_)}" for p_ in ("本公司 10 年中位數", "納入的同業家數", "同業 10 年中位數的中位數", "原始比值"))
_rec, _grd = f"Val!B{REC_ROW}", f"Val!B{GRADE_ROW}"
_T = lambda x, f="#,##0": (f'TEXT({x},"{f}")',)
_SP = '"+0.0%;-0.0%"'
_S = lambda x: (f"TEXT({x},{_SP})",)
_dP = f"tp_w_dcf*({_B('dcftp')}-px)"; _cP = f"(1-tp_w_dcf)*(tp_comp_adj-px)"
_EN = f"Val!B{O + 7}"                                                # 現在的未來 12 個月 EPS（模型情境 B；與共識前瞻 EPS 同一時間窗）
_M = f"({_B('pmed')}*rm_factor)"                                     # Comparable 用的倍數
_HL = lambda x: (f'IF({x}>=0,"高 ","低 ")&TEXT(ABS({x}),"0.0%")',)
_WIN = ("Comparable 的上檔主要是倍數：以同一時間窗（現在的未來 12 個月）比較，本模型 EPS（情境 B）約 ", _T(_EN, "0.00"), " 元、比共識前瞻 EPS ", _T("cons_feps", "0.00"), " 元", _HL(_EN + "/cons_feps-1"),
        "，現價相當於 ", _T("px/" + _EN, "0.0"), " 倍（以共識 EPS 計 ", _T("px/cons_feps", "0.0"), " 倍）；Comparable 用的倍數是同業（創意、智原、聯發科、博通、邁威爾）前瞻 P/E 中位數 ", _T(_B("pmed"), "0.0"),
        " 倍 × 相對倍數調整係數 ", _T("rm_factor", "0.00"), "＝", _T(_M, "0.0"), " 倍，再乘一年後的前瞻 EPS（情境加權 ", _T(_B("feps"), "0.00"), " 元，比現在", _HL(_B("feps") + "/" + _EN + "-1"),
        "）。兩者相乘：倍數差（Comparable 倍數 ÷ 現價倍數 − 1）", _S(f"{_M}*{_EN}/px-1"), "、一年 EPS 成長 ", _S(_B("feps") + "/" + _EN + "-1"), "，",
        (f'IF(ABS(LN({_M}*{_EN}/px))>=ABS(LN({_B("feps")}/{_EN})),"倍數差是主因","盈餘成長是主因")',), "。同業倍數未扣世芯單一客戶集中與 turnkey 低毛利的折價，Comparable 可能偏高。")
_RRD = {}
for _r in range(1, RD.max_row + 1):
    _v = str(RD.cell(_r, 1).value or "")
    for _k in ("1", "2", "3"):
        if _v.startswith((f"重點觀察 {_k}：", f'="重點觀察 {_k}：')): _RRD.setdefault(_k, _r)
assert len(_RRD) == 3, _RRD
RD.cell(_RRD["1"], 1, XF("重點觀察 1：目標價 NT$", _T(_B("tpw")), "（DCF ", _T(_B("dcftp")), "、Comparable（相對倍數調整後）", _T("tp_comp_adj"), "；DCF 權重 ", _T("tp_w_dcf", "0%"),
    "），比現價 ", _T("px"), "（2026-10-22）", (f'IF({_B("upw")}>=0,"高 ","低 ")&TEXT(ABS({_B("upw")}),"0.0%")',), "，級距 ", (_grd,), "、投資建議 ", (_rec,),
    "。拆到估值來源：加權目標價相對現價的差距（", (f'TEXT({_B("tpw")}-px,"+#,##0;-#,##0")',), " 元）中，Comparable 部分 ", (f'TEXT({_cP},"+#,##0;-#,##0")',),
    "、DCF 部分 ", (f'TEXT({_dP},"+#,##0;-#,##0")',), "（標準版 DCF ", _T(_B("dcftp")), (f'IF({_B("dcftp")}<px,"，比現價低 "&TEXT(1-{_B("dcftp")}/px,"0.0%"),"，比現價高 "&TEXT({_B("dcftp")}/px-1,"0.0%"))',),
    "；WACC ", _T("wacc", "0.0%"), "）。", *_WIN, "只看 Comparable（市場主流同業法、未調整）為 ", _S(_B("upc")), "。"))
RD.cell(_RRD["2"], 1, XF("重點觀察 2：情境分歧很大：FY2028 EPS A／B／C 為 ", _T(_E("A", 2028), "0.0"), "／", _T(_E("B", 2028), "0.0"), "／", _T(_E("C", 2028), "0.0"),
    "，DCF 價格帶 ", (_B("band"),), "（v0.2 標準版 DCF），差在 Trainium 顆數與世芯份額（CY2028 起 ", _T("al28_A", "0%"), "／", _T("al28_B", "0%"), "／", _T("al28_C", "0%"),
    "，假設）；目標價對 DCF 權重也敏感，權重 0.7 時 ", _S(_TS(_DWK, "F")), "（約 ", _T(_TS(_DWK, "E")), "，上檔 ", _S(_TS(_DWK, "E") + "/px-1"), "）。"))
_g4 = lambda nm: f"Sens!G{5 + _CN.index(nm)}/Sens!G$4-1"
RD.cell(_RRD["3"], 1, XF("重點觀察 3：Trainium 份額是 FY2027 EPS 最敏感的變數（情境 A 為基準：取高值 ", _S(_SF("Trainium 份額取高值")), "、−10pt 時 ", _S(_SF("Trainium 份額 −10pt")),
    "）；FY2028 起再加上 2 奈米每顆營收（×2.5 時 FY2028 營收 ", _S(_g4("2 奈米每顆營收 ×2.5")), "、COT 模式 ×0.8 時 ", _S(_g4("2 奈米每顆營收 ×0.8")), "）（Sens 頁）。下一個檢驗點是 3Q26 財報與法說（約 2026-11 上旬）的 Trainium 3 量產進度。"))
_r = RD.max_row + 2
RD.cell(_r, 1, XF("v0.2.2（2026-11-06）：第三輪修正。(1) Stub 期 FCFF 改用含 stub 期會計年度 FY2026（錨點）FCFF × stub_frac（原誤用 FY2027 全年）：情境 B ", (f'TEXT(Val!C{SDCF["stub"]},"0.00")',),
    "，淨現金滾到評價日 ", (f'TEXT(Val!C{SDCF["ncr"]},"0.00")',), "（NT$B）。(2) 現價改 2026-11-04 收盤 ", _T("px"), "（Yahoo Finance；原 2026-10-22）；Comparable 同業前瞻 P/E 與本公司前瞻 P/E 改為 11-04 收盤 ÷ 資料商前瞻 EPS，同業中位數 ", _T(_B("pmed"), "0.00"),
    "。加權目標價 ", _T(_B("tpw"), "#,##0.00"), "（上檔 ", _S(_B("upw")), "，", (_rec,), "）。(3) 展示文字：β 來源改引用 docs/plan/自算產業β_20261105.md（刪本機路徑）；重點觀察 1 刪除調整前舊說明；級距列名「六級，供分組」；情境驅動收斂法列名中性；Sens「g ±1pt」改「g −1pt／上限 rf」。",
    "(4) HTML：DCF 彙總主表顯示「＋Stub 期 FCFF」「淨現金滾到評價日」；情境驅動表百分比格式化；橫幅標明新台幣。")); _r += 1
RD.cell(_r, 1, XF("v0.2.1（2026-11-05）：換用共用模組 tools/fcff_dcf.py v2（工作單第二輪修正；Andy 2026-11-05 23:16 定案）。(1) 資產負債表：最新已公布一季仍為 2026-06-30（3Q26 財報尚未公布），淨現金與投資 ",
    (f'TEXT(dcf_nc,"0.00")',), "（不再先扣 10 月股利）。(2) Stub 期：評價日 2026-12-31，淨現金滾到評價日＝上述＋stub 期 FCFF（各情境 FY2027 FCFF × stub_frac ", (f'TEXT(stub_frac,"0.000")',), " 年；情境 B ",
    (f'TEXT(Val!C{SDCF["stub"]},"0.00")',), "）− 已宣告股利 ", (f'TEXT(stub_div,"0.00")',), "（每股 32.57285 元，2026-10-08 發放）＝", (f'TEXT(Val!C{SDCF["ncr"]},"0.00")',), "（十億新台幣）。",
    "DCF 目標價 3,534.28 → ", _T(_B("dcftp"), "#,##0.00"), "、加權目標價 6,128.51 → ", _T(_B("tpw"), "#,##0.00"), "（上檔 ", _S(_B("upw")), "，", (_rec,), "）。(3) β 敏感度：低／高改為 β_u ∓ 0.3 各自重新加槓桿（In 頁 beta_du）。",
    "(4) 新增 Sens「美元口徑折現」：rf＝美國 10 年期 ", _T("rf_us", "0.00%"), "＋CRP、g＝MIN(3%, ", _T("rf_us", "0.00%"), ")，其他不變；加權目標價 ", _T(f"Sens!C{_USD}", "#,##0"), " 元（", _S(f"Sens!D{_USD}"), "）。",
    "(5) ke 說明格改由模組參數 beta_note 提供；In 頁 M2 標題不寫版號。資本支出、折舊：本模型沒有逐年資本支出／折舊列，維持占營收比例（資本支出 FY2029 起收斂到 fc_cx_ss）。PL、Comparable、情境權重不變。")); _r += 1
RD.cell(_r, 1, "v0.2（2026-11-05）：模型改版——DCF 改為標準版（Andy 2026-11-05 20:42、20:45；docs/plan/個股DCF標準規格.md）：FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（含客戶預收），以 WACC 折現，終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）；"
    "ke＝rf＋β ×（ERP＋CRP），rf 1.98%（新台幣 10 年期公債）、β＝自算台灣半導體產業無槓桿 β 1.429（同業 13 家中位數 1.64，Blume 調整；Andy 2026-11-05 22:55）依市值 D/E 重新加槓桿、ERP 5%、台灣 CRP 0.78%，g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金與長期投資 ÷ 股數，以 ke 推到 12 個月後並扣期間股利（每股 32.57 元，Assumed）。"
    "新增 DCF 頁（tools/fcff_dcf.py），Val 0.1 節改為 WACC 與 DCF 彙總，Sens 目標價參數改為 WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法與 DCF 權重；Wattonomics 錨定案例改用同一 FCFF 式；刪除 tp_r、tp_roe，tp_g 只留給 PL 延伸期收斂。"
    "同時改用個股 HTML 版面規格 v1.0（README 新增展示層、命題回覆；重點觀察 1–3 改為公式）。PL、Q、Units、Comparable、情境權重不變。"
    "新舊對照（v0.1.3 → v0.2）：DCF 目標價 3,503.25 → 3,534.28，加權目標價 6,113.00 → 6,128.51，上檔空間 +74.9% → +75.4%，投資建議 強力買進 → 強力買進（不變）。"
    "DCF 差異拆解（建置時以 DCF 頁明細依序替換計算，元／股）：現金流口徑改 FCFF（仍 r 10%、g 4%、終值 FCFF ×（1＋g））+1,300.5；折現率改 WACC 10.2%（並以 ke 推到 12 個月後）−148.0；g 改 1.98% −735.2；終值改價值驅動式（RONIC＝WACC）−353.7；扣期間股利 −32.6。"
    "FCFF 每股合計（FY2027–32 未折現）A 2,192.5／B 1,276.3／C 575.8，舊股東現金流 1,862.2／1,098.9／513.2（FCFF 不扣再投資、只扣資本支出與營運資金）。"); _r += 2
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.2 起；上方各列原文保留給下游）").font = BOLD; _r += 1
RD.cell(_r, 1, "命題：世芯的營收由 NRE 轉為 AWS Trainium 量產（turnkey，含晶圓與封裝轉售、毛利率低）。以市場層 Trainium 顆數 × 世芯份額 × 每顆營收往前推，依序回答以下三個問題。"); _r += 1
RD.cell(_r, 1, "命題 1：FY2027–28 能賺多少。在市場層 Trainium 顆數（台積電模型）與 Marvell 份額口徑下，世芯 FY2027–28 的營收與 EPS 有多少，和賣方估計差多少。"); _r += 1
RD.cell(_r, 1, "命題 2：2 奈米換代靠什麼。FY2028 起 Trainium 4（2 奈米）換代，世芯的營收靠份額還是每顆營收，兩者各有多大影響。"); _r += 1
RD.cell(_r, 1, "命題 3：每股值多少。用標準版 DCF 與同業倍數算出的目標價，和現價相比是高是低、差在哪裡。"); _r += 1
RD.cell(_r, 1, XF("重點 1：加權目標價", (f'IF({_B("upw")}>=0,"高於","低於")',), "現價，投資建議", (_rec,), "。加權目標價約 ", _T(_B("tpw")), " 元，比現價 ", _T("px"), " 元",
    (f'IF({_B("upw")}>=0,"高 ","低 ")&TEXT(ABS({_B("upw")}),"0.0%")',), "。拆解這段差距：Comparable（相對倍數調整後約 ", _T("tp_comp_adj"), " 元）貢獻 ", (f'TEXT({_cP}/({_B("tpw")}-px),"0%")',),
    "、DCF（約 ", _T(_B("dcftp")), " 元）貢獻 ", (f'TEXT({_dP}/({_B("tpw")}-px),"0%")',), "（DCF 低於現價時為負），", (f'IF(ABS({_cP})>=ABS({_dP}),"主因是 Comparable","主因是 DCF")',), "。", *_WIN, "只看 Comparable（市場主流同業法、未調整）目標價 ", _T(_B("comptp")), " 元，比現價高 ", (f'TEXT({_B("upc")},"0.0%")',), "。"))
_r += 1
RD.cell(_r, 1, XF("重點 2：三情境差很大，差在 Trainium 顆數與世芯份額。FY2028 EPS 情境 A／B／C 約 ", _T(_E("A", 2028), "0.0"), "／", _T(_E("B", 2028), "0.0"), "／", _T(_E("C", 2028), "0.0"),
    " 元，DCF 價值（推到 12 個月後）", _T(_B("dcf12", "D")), "～", _T(_B("dcf12", "B")), " 元。世芯 Trainium 份額 CY2028 起 ", _T("al28_A", "0%"), "／", _T("al28_B", "0%"), "／", _T("al28_C", "0%"),
    "、2 奈米每顆營收係數 ", _T("pfc_A", "0.0"), "／", _T("pfc_B", "0.0"), "／", _T("pfc_C", "0.0"), " 都是假設；顆數情境係數讀台積電模型（CY2028 ", _T("um28_A", "0.0"), "／", _T("um28_B", "0.0"), "／", _T("um28_C", "0.0"),
    "）。目標價對 DCF 權重也敏感：權重 0.7 時加權目標價 ", _S(_TS(_DWK, "F")), "、0.3 時 ", _S(_TS(_DWK, "D")), "。"))
_r += 1
RD.cell(_r, 1, XF("重點 3：Trainium 份額是 FY2027 EPS 最敏感的變數；FY2028 起再加上 2 奈米每顆營收。以情境 A 為基準（Sens 頁），份額取高值時 FY2027 EPS ", _S(_SF("Trainium 份額取高值")),
    "、−10 個百分點時 ", _S(_SF("Trainium 份額 −10pt")), "；2 奈米每顆營收 ×2.5（美系外資說法）時 FY2028 EPS ", _S(_SF("2 奈米每顆營收 ×2.5", "J")), "、×0.8（COT 模式）時 ", _S(_SF("2 奈米每顆營收 ×0.8", "J")),
    "。下一個檢驗點是 3Q26 財報與法說（約 2026-11 上旬）的 Trainium 3 量產進度。"))
_r += 1
RD.cell(_r, 1, XF("重點 4：v0.2 起 DCF 改為標準版（無槓桿自由現金流折現），DCF 目標價 ", _T(_B("dcftp")), " 元（舊法股東現金流、折現率 10%、成長率 4% 為 3,503 元；差異拆解見上方 v0.2 版本說明）。WACC ",
    _T("wacc", "0.0%"), "（股權成本 ", _T("ke", "0.0%"), "＝rf ", _T("rf", "0.00%"), " ＋ β ", _T("beta", "0.00"), " ×（ERP ", _T("erp", "0%"), " ＋ 台灣國家風險溢酬 ", _T("crp", "0.00%"),
    "）；β 是自算台灣半導體產業無槓桿 β（Blume 調整）依市值 D/E 重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設），永續成長率 ", _T("g_tv", "0.00%"), "（＝新台幣 10 年期公債殖利率，低於上限 3%）。終值現值占情境 B 企業價值 ",
    (f'TEXT(Val!C{SDCF["tvsh"]},"0%")',), "，隱含終值 EV／EBITDA 約 ", (f'TEXT(Val!C{SDCF["mult"]},"0.0")',), " 倍，低於同業交易倍數：WACC 高、永續成長率低、RONIC＝WACC（永續期新投資不創造價值），且折舊攤銷占營收 ",
    _T("fc_da", "0.0%"), "（EBITDA 比 EBIT 大）。以情境加權目標價為基準：WACC −1pt／＋1pt 時 ", _S(_TS(0)), "／", _S(_TS(0, "F")), "，β 低／高 ", _S(_TS(1)), "／", _S(_TS(1, "F")),
    "，終值改用 FCFF ×（1＋g） ", _S(_TS(8, "F")), "。",
    "美元口徑折現（敏感度案例，未進基準；rf 改美國 10 年期 ", _T("rf_us", "0.00%"), " 並保留台灣 CRP ", _T("crp", "0.00%"), "、g＝MIN(", _T("g_cap", "0%"), ", ", _T("rf_us", "0.00%"), ")，現金流與匯率不變）時加權目標價 ",
    _T(f"Sens!C{_USD}"), " 元（相對基準 ", _S(f"Sens!D{_USD}"), "）；基準採當地利率＋固定匯率（Andy 2026-11-05 23:16）。", (_TVF,)))
_r += 2
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上面三個命題）").font = BOLD; _r += 1
RD.cell(_r, 1, XF("FY2027–28 能賺多少？答：FY2027 營收三情境約 ", _T(_P("rev", "A", 2027), "0"), "／", _T(_P("rev", "B", 2027), "0"), "／", _T(_P("rev", "C", 2027), "0"),
    " 十億新台幣、EPS 約 ", _T(_E("A", 2027), "0"), "／", _T(_E("B", 2027), "0"), "／", _T(_E("C", 2027), "0"), " 元；FY2028 EPS 約 ", _T(_E("A", 2028), "0"), "／", _T(_E("B", 2028), "0"), "／", _T(_E("C", 2028), "0"),
    " 元（情境 A／B／C）。對照賣方：三情境 FY2027 EPS 相對美系外資甲 ", _T("b1_eps27", "0"), " 元分別 ", _S(_E("A", 2027) + "/b1_eps27-1"), "／", _S(_E("B", 2027) + "/b1_eps27-1"), "／", _S(_E("C", 2027) + "/b1_eps27-1"),
    "；FY2028 Trainium 營收三情境約 ", _T(_P("asu", "A", 2028), "0.0"), "／", _T(_P("asu", "B", 2028), "0.0"), "／", _T(_P("asu", "C", 2028), "0.0"), " 十億美元，美系外資估 Trainium 4 約 ", _T("b1_trn4", "0.0"),
    " 十億美元（不含 3 奈米尾量）。差異主要來自市場層 Trainium 顆數（CY2027 ", _T("u_trn_27", "0.0"), " 百萬顆，低於 UBS 估的 Trainium 3 ", _T("ubs_trn27", "0.0"), " 百萬顆）與世芯份額假設；市場層顆數的修訂應由市場層提出。"))
_r += 1
RD.cell(_r, 1, XF("2 奈米換代靠份額還是每顆營收？答：兩者都重要，FY2028 最大的單一變數是", (f'IF(ABS({_SF("2 奈米每顆營收 ×2.5", "J")})>=ABS({_SF("Trainium 份額取高值", "J")}),"每顆營收","份額")',), "。以情境 A 為基準（Sens 頁 FY2028 EPS），2 奈米每顆營收 ×2.5 時 ", _S(_SF("2 奈米每顆營收 ×2.5", "J")),
    "、×0.8（COT 模式，客戶自購晶圓）時 ", _S(_SF("2 奈米每顆營收 ×0.8", "J")), "；份額取高值時 ", _S(_SF("Trainium 份額取高值", "J")), "、−10 個百分點時 ", _S(_SF("Trainium 份額 −10pt", "J")),
    "。每顆世芯營收（校準後）情境 A 由 FY2027 約 ", _T(_P("ppu", "A", 2027), "0.0"), " 千美元到 FY2028 約 ", _T(_P("ppu", "A", 2028), "0.0"), " 千美元；基準的 2 奈米係數 ", _T("pfc_A", "0.0"), "／", _T("pfc_B", "0.0"), "／", _T("pfc_C", "0.0"),
    " 與份額 ", _T("al28_A", "0%"), "／", _T("al28_B", "0%"), "／", _T("al28_C", "0%"), " 都是假設（公司只說 2 奈米價格遠高於 3 奈米，但主要客戶採 COT 模式）。"))
_r += 1
RD.cell(_r, 1, XF("每股值多少？答：加權目標價約 ", _T(_B("tpw")), " 元（DCF ", _T(_B("dcftp")), "、Comparable 相對倍數調整後 ", _T("tp_comp_adj"), "，權重各半），比現價 ", _T("px"), " 元",
    (f'IF({_B("upw")}>=0,"高 ","低 ")&TEXT(ABS({_B("upw")}),"0.0%")',), "，投資建議", (_rec,), "。只看 Comparable（市場主流同業法，不做相對倍數調整）的目標價 ", _T(_B("comptp")), " 元、上檔 ", _S(_B("upc")),
    "，高於市場共識目標價 ", _T(_B("cons")), " 元；進加權的是調整後的 ", _T("tp_comp_adj"), " 元。標準版 DCF 三情境價值（推到 12 個月後）順利、平平、不如預期約 ", _T(_B("dcf12", "B")), "、", _T(_B("dcf12", "C")), "、", _T(_B("dcf12", "D")),
    " 元，", (f'IF(AND({_B("dcf12", "B")}>px,{_B("dcf12", "C")}<px),"只有情境 A 高於現價",IF({_B("dcf12", "D")}>px,"三情境都高於現價",IF({_B("dcf12", "B")}<px,"三情境都低於現價","情境 A、B 高於現價")))',),
    "。以現在的未來 12 個月 EPS 計，現價約為本模型（情境 B）的 ", _T("px/" + _EN, "0.0"), " 倍、共識的 ", _T("px/cons_feps", "0.0"), " 倍（共識目標價 ÷ 共識前瞻 EPS 為 ", _T("cons_tp/cons_feps", "0.0"), " 倍），同業前瞻 P/E 中位數 ", _T(_B("pmed"), "0.0"), " 倍；上檔主要看市場是否給世芯接近同業的倍數。"))
_r += 2
RD.cell(_r, 1, "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間（HTML 估值表用；加權目標價使用的 Comparable 即此值；投資建議橫幅中間格放只看 Comparable〔市場主流同業法〕，見 Val 頁）")
RD.cell(_r, 2, "=tp_comp_adj/px-1").number_format = "0.0%"
RD["A1"] = (str(RD["A1"].value).replace("v0.1.2", "v0.2.2", 1) + "；v0.2（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），"
            "Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；README 新增展示層與命題回覆（HTML 新版面 v1.0）"
            "；v0.2.1（2026-11-05）換用共用模組 tools/fcff_dcf.py v2：淨現金由資產負債表日 2026-06-30 滾到評價日 2026-12-31（stub 期 FCFF − 已宣告股利）、β 敏感度改 β_u ± 0.3、新增 Sens「美元口徑折現」"
            "；v0.2.2（2026-11-06）第三輪修正：stub 期改 FY2026 錨點 FCFF、現價與同業報價統一 2026-11-04、展示文字與 HTML")
# ---------------- 第三輪修正（2026-11-06）：展示文字統一（情境名稱、來源字樣、舊重點觀察、級距列名、stub 口徑說明）；只改文字，不改數值 ----------------
import re as _re3
def _r3_fix(wb, ver, stub_fy, extra=()):
    """ver：本版版號；stub_fy：stub FCFF 所用會計年度的說明；extra：公司自訂 (regex, 取代) 清單（先套用）"""
    subs = list(extra) + [
        (r"(情境 ?A[^（）\n\"]{0,8})（樂觀）", r"\1（順利）"), (r"(情境 ?B[^（）\n\"]{0,8})（基準）", r"\1（平平）"),
        (r"(情境 ?C[^（）\n\"]{0,8})（(?:保守|下檔|悲觀)）", r"\1（不如預期）"),
        (r"／B（基準）", "／B（平平）"), (r"／C（(?:保守|下檔|悲觀)）", "／C（不如預期）"),
        (r"在?(?:本|建置)環境(?:中)?(?:無法開啟|找不到|取不到)(?:（未找到）)?", "未取得"),
        (r"無法開啟：未找到", "未取得"), (r"px（2026-10-23 收盤）", "px（2026-11-05 收盤）"), (r"無法開啟", "未取得"), (r"取不到", "未取得"), (r"，已抓取", ""), (r"已逐頁抓取", "已逐頁讀取"), (r"已抓取", "已讀取"), (r"抓取", "取得"),
        (r"第一年 FCFF × stub_frac", f"含 stub 期會計年度（{stub_fy}）FCFF × stub_frac"),
        (r"重點觀察 1 的第一句已改為公式；重點觀察 2、3 與其他說明中的目標價與上檔空間數字為調整前撰寫，以 Val 頁為準。",
         f"重點觀察 1 的目標價等數字為公式；調整前撰寫的舊說明已於 {ver} 刪除（第三輪修正）。"),
    ]
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                w = v
                for a, b in subs:
                    w = _re3.sub(a, b, w)
                if w.startswith(("重點觀察 1：", '="重點觀察 1：')) and '&"以下為 2026-10-29 相對倍數調整前撰寫的說明' in w:
                    w = w[:w.index('&"以下為 2026-10-29 相對倍數調整前撰寫的說明')]
                if ws.title == "In" and not w.startswith("=") and "自算產業 β：" in w and "docs/plan/自算產業β_20261105.md" not in w:
                    w = w + "（同業表與方法：docs/plan/自算產業β_20261105.md）"
                if w.startswith("上檔空間級距") and ws.title == "Val" and c.column == 1 and "供分組" not in w:
                    w = "上檔空間級距（六級，供分組）"
                    e = ws.cell(c.row, 5)
                    if isinstance(e.value, str) and "投資建議" not in e.value:
                        e.value = e.value + "；投資建議另依五級固定規則（下一列「投資建議」）"
                if w != v:
                    c.value = w; n += 1
    return n
def _r3_static(wb, prefix, text):
    """README「本版要點」前一版列：公式改為靜態文字（避免顯示本版數字）"""
    ws = wb["README"]
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith('="' + prefix):
            ws.cell(r, 1).value = text; return 1
    raise KeyError(prefix)
def _r3_gcap(wb):
    """第三輪清單 16：g 受 rf 上限（g＝rf）時，Sens「g ±1pt」列名改「g −1pt／上限 rf」（高案例被上限夾住＝基準）"""
    ws = wb["Sens"]; n = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("永續成長率 g ±1pt"):
            ws.cell(r, 1).value = v.replace("永續成長率 g ±1pt", "永續成長率 g −1pt／上限 rf（g 已等於 rf，＋1pt 被上限夾住）"); n += 1
    return n
# 第三輪清單 15：情境驅動對照的收斂法列名改中性
_wsI = wb["In"]; _s0 = next(r for r in range(1, _wsI.max_row + 1) if str(_wsI.cell(r, 1).value or "").startswith("情境驅動對照"))
for _rr in range(_s0 + 2, _wsI.max_row + 1):
    if not _wsI.cell(_rr, 1).value: break
    _lb = str(_wsI.cell(_rr, 2).value or "")
    if _lb.startswith("FY2032 ") and _lb.endswith("）"): _wsI.cell(_rr, 2).value = _lb[:_lb.rindex("（")]
_r3_static(wb, "v0.2.1（2026-11-05）", "v0.2.1（2026-11-05）：換用共用模組 tools/fcff_dcf.py v2（工作單第二輪修正；Andy 2026-11-05 23:16 定案）。(1) 資產負債表：最新已公布一季仍為 2026-06-30（3Q26 財報尚未公布），淨現金與投資 55.65（不再先扣 10 月股利）。(2) Stub 期：評價日 2026-12-31，淨現金滾到評價日＝上述＋stub 期 FCFF（各情境 FY2027 FCFF × stub_frac 0.504 年；情境 B 6.99）− 已宣告股利 2.81（每股 32.57285 元，2026-10-08 發放）＝59.83（十億新台幣）。DCF 目標價 3,534.28 → 3,623.34、加權目標價 6,128.51 → 6,173.04（上檔 +76.6%，強力買進）。(3) β 敏感度：低／高改為 β_u ∓ 0.3 各自重新加槓桿（In 頁 beta_du）。(4) 新增 Sens「美元口徑折現」：rf＝美國 10 年期 5.00%＋CRP、g＝MIN(3%, 5.00%)，其他不變；加權目標價 5,876 元（-4.8%）。(5) ke 說明格改由模組參數 beta_note 提供；In 頁 M2 標題不寫版號。資本支出、折舊：本模型沒有逐年資本支出／折舊列，維持占營收比例（資本支出 FY2029 起收斂到 fc_cx_ss）。PL、Comparable、情境權重不變。")
_r3_gcap(wb)   # rf 1.98% < g 上限 3%：g＝rf
_r3_fix(wb, "v0.2.2", "FY2026＝日曆年 2026，錨點", extra=[('"（2026-10-22）"', '"（2026-11-05）"'), ('與股價同日（2026-10-23）', 'rf 沿用 2026-10-23'), ('同業倍數未扣世芯單一客戶集中與 turnkey 低毛利的折價，Comparable 可能偏高。', '同業倍數未扣世芯單一客戶集中與 turnkey 低毛利的折價，Comparable 可能偏高。')])
# ---------------- v0.2.3（2026-11-06）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守（本資料夾 px1105_post.py；只改文字） ----------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import px1105_post  # noqa: E402
print(px1105_post.apply(wb, os.path.dirname(os.path.abspath(__file__)), "v0.2.3", "v0.2.2", "v0.2.2（2026-11-06）"))
wb.save(sys.argv[1])
if os.environ.get("ALCHIP_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10,
               "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0},
              open(os.environ["ALCHIP_ROWMAP"], "w"), ensure_ascii=False, indent=1)
