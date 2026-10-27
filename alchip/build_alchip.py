# 世芯-KY（3661.TW）收支與估值模型 v0.1.1（AI 半導體 Project：運算層，ASIC 設計服務＋量產；樣板 mediatek/、mrvl/）
# 用法：python3 alchip/build_alchip.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；每顆價格：千美元（US$K）；EPS：新台幣（基本）。
# 營收＝NRE（委託設計）＋量產（turnkey：晶圓、封裝、測試轉售，毛利率低）。
# Trainium 量產營收＝AWS Trainium 顆數（tsmc 成品）× 世芯份額 × 每顆營收（Broadcom 模型的 pa_trn 扣 HBM，與 Marvell 模型同口徑）× k_x × 匯率。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import cycle_position as CP
import valuation_rule as VR
import market as MEM
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
            "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、Broadcom、Marvell 模型相同；markets/memory SCEN_UNITS，建置時檢查）" if (s == "A" and y == 28) else "同上")
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
    inp(*SIL.row9(key))
    assert MV[f"msh_trn_{y}"][0][1] + SIL.V[key]["hi"] <= 1.0 + 1e-9, key   # Marvell（基準）＋世芯（高）≤ 100%
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
inp("fx", "匯率（FY2027 起；Trainium 美元營收換算）", "NT$/US$", 30.0, 32.0, 34.0, "Assumed", "與聯發科、台積電模型相同假設", "其他營收以新台幣成長率推，未另換算")

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
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 3495.0, None, "Interested-party", SA + "：10-22 收盤 3,495（−5.80%）", "10-23 收盤：未找到")
inp("pe_ttm", "近 12 月 P/E（stockanalysis）", "x", None, 50.74, None, "Interested-party", SA, "companiesmarketcap 47.2")
inp("fpe", "前瞻 P/E（stockanalysis）", "x", None, 17.15, None, "Interested-party", SA, "—")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "NT$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 17.15", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", 3755, 5803.25, 6800, "Interested-party", SA + "：16 位分析師平均 5,803（2026-10-23 擷取）", "共識頁（2026-08-13）：15 位平均 5,666、中位數 6,000；歐系外資 6,784（2026-09-23）")
inp("nim5", "近五年平均淨利率（FY2021–25；對照用）", "%", None, "=AVERAGE(1490/10428,1834/13725,3325/30482,6446/51969,5598/30926)", None, "Derived",
    "stockanalysis 財報頁：FY2021–25 營收與淨利", "十年平均（2016–2025）：未找到（2016–2020 未取得）；收斂法依工作單改用最後一年 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%）", "美、台、日、韓同一值；單一客戶集中度高的風險不改折現率，放在情境 C")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.2140,0.2242,0.1394)", None, "Derived", "stockanalysis 比率頁：FY2023 21.40%、FY2024 22.42%、FY2025 13.94%", "單一來源；FY2025 偏低（營收 −40%）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與「FY2029 × 0.8」的中點）", "%", None, 0, None, "Assumed", "十年平均未找到，依工作單以最後一年 × 0.8 代替", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

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
    for u_, nm_ in URLNAME.items():          # 網址只列在 Sources 頁；In 頁寫來源名稱
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
            go27="g_o27_{s}", goL="g_oL_{s}", gma="0", phi="phi", opxa="0", nonop="nonop_f", tax="tax", fg="fade_g{s}", fm="fade_m{s}")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
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
    row("b", "世芯顆數：AWS Trainium", "百萬顆", {FC[y]: f"={FC[y]}{R['u']}*{FC[y]}{R['sh']}" + ("" if y == 2026 else f"*{P['uf']}") for y in BU}, "0.000")
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
V.cell(O + 2, 1, "現價（2026-10-22 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0"
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
V.cell(5, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
hdr(V, 6, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境；FY2030 起為收斂法"; T["eps_h"] = 7
EPSR = {}
for i, y in enumerate(YRS6):
    tset(8 + i, f"eps{y}", f"  FY{y}", [f"=PL!{FC[y]}{SUMR[(s, 'eps')]}" for s in SC], "#,##0.00"); EPSR[y] = 8 + i
V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = 14
tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B$15)" for c in "BCD"], "#,##0.00")
def dcf_formula(c, r="tp_r", g="tp_g", reinv="$B$15"):
    """c＝情境欄；r、g 可替換（Sens 用）。基準日 2026 年底，t＝1 對 FY2027；負值取 0（Andy 12:22）。"""
    cf = lambda y: f"Val!{c}{EPSR[y]}*(1-{reinv})"
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"MAX(0,{pv}+{cf(2032)}*(1+{g})/({r}-{g})/(1+{r})^6)"
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "#,##0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；負值取 0")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "#,##0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "#,##0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("創意（GUC）", "3443.TW", 102.85, "2026-10-08"), ("智原（Faraday）", "3035.TW", 37.81, "2026-09-04"), ("聯發科（MediaTek）", "2454.TW", 38.75, "2026-10-22"),
         ("博通（Broadcom）", "AVGO", 18.52, "2026-10-23"), ("邁威爾（Marvell）", "MRVL", 40.44, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "台灣 ASIC 設計服務（創意、智原）＋ASIC 對手（聯發科、博通、邁威爾）；不含 A 股（非 A 股公司，Andy 2026-10-25 11:53）")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "#,##0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "#,##0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "#,##0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", "stockanalysis：16 位分析師平均 5,803，2026-10-23 擷取")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；低於十年最小值記為 0%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 19%（再投資率約 21%），FY2027–32 明確期＋永續；世芯為輕資產設計服務公司（量產為轉售，資本支出小），不改用淨利＋折舊－資本支出；情境價值負值取 0。"
            "Comparable：創意、智原、聯發科、博通、邁威爾前瞻 P/E 中位數；創意 103x 為極端值，中位數不受影響；世芯本身前瞻 P/E 約 17x，遠低於同業，反映單一客戶集中與 turnkey 低毛利，Comparable 偏向高估。"
            "主要偏誤：Trainium 顆數受市場層 CY2027 2.5 百萬顆限制（UBS 估 Trainium 3 即 2.8 百萬顆），FY2030 起只用收斂法，DCF 偏低估；2 奈米 COT 模式若使每顆營收下降，則偏高估。"
            "情境對應：A＝順利（tsmc 顆數 A、份額 55%）、B＝平平（顆數 B、份額 50%／45%）、C＝不如預期（顆數 C、份額 40%／35%）")
for r_ in (24, 35, 39):
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
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；PL 收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def tp_case(i, name, setting, r="tp_r", g="tp_g", w="tp_w_dcf"):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    reinv = f"MIN({g}/tp_roe,0.8)"
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, r, g, reinv)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
    SE[f"F{i}"] = f"=E{i}/E${TPS0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = "#,##0.00"
    SE[f"F{i}"].number_format = "0.0%"
TPCASES = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.3", "", dict(w="0.3")), ("DCF 權重 0.7", "", dict(w="0.7"))]
for k, (nm, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + 2 + k, nm, st, **kw)
TPS_LAST = TPS0 + 1 + len(TPCASES)
assert TPS_LAST < 40, TPS_LAST
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
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，情境價值負值取 0）＋0.5 × Comparable 目標價（創意、智原、聯發科、博通、邁威爾前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到，B、C 依工作單以 FY2029 × 0.8 代替。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自世芯新聞稿與媒體轉述財報；法說內容多為二手轉述；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 公司未揭露客戶名稱、NRE 與量產金額、HBM 採購方式，Trainium 份額、拆分與毛利率為 Assumed；(2) 市場層 CY2027 Trainium 2.5 百萬顆低於 UBS 的 Trainium 3 2.8 百萬顆，修訂應由市場層提出，不在本模型自改；(3) 2 奈米每顆營收係數為 Assumed（公司：價格遠高於 3 奈米；但主要客戶採 COT 模式），放在情境與 Sens；(4) 單一客戶集中度高（分析師估 60–70% 以上），情境 C 只反映份額與顆數下修，未含失去客戶；(5) 3Q26 毛利率與費用為指引值；(6) 十年 P/E 多為 NRE 主導期；2016 為負值、2016–2017 EPS 未找到；(7) R4 產業部署路徑未建（Recon 預留列）。", False), ("", False),
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


order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("ALCHIP_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10,
               "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0},
              open(os.environ["ALCHIP_ROWMAP"], "w"), ensure_ascii=False, indent=1)
