# 鴻海精密（2317）個股模型 v0.3（AI 半導體 Project：伺服器組裝次產業，樣板＝quanta/）
# 用法：python3 foxconn/build_foxconn.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；整櫃價格與市場以美元輸入、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利。集團公司：比照三星只拆「雲端網路產品」中的 AI 伺服器做細；雲端網路的非 AI 部分（網通、一般伺服器）、
# 消費智能（iPhone）、電腦終端、元件及其他以簡化成長率推。業外、稅與非控制權益以「淨利 ÷ 營業利益」一個係數涵蓋。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import valuation_rule as VR

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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); BR_F = only("markets/chip_bridge/dist/*.xlsx")
QU_F = only("quanta/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ITC = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v is not None for v in ITC[0]), ITC
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
BR = {}
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("racks", "capit", "rack"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 3, BR.keys()
# 廣達模型（組裝線權威，規則第 13 條）：ODM 每櫃價格、Rubin 占比、匯率、每 GW 價值變動；有效份額只作合計檢查
QWB = openpyxl.load_workbook(QU_F, data_only=True)
QU = {}
for r in QWB["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("odm_gb", "odm_vr", "vr26", "fx", "vpg_28", "vpg_29"):
        QU[k] = (r[3].value, r[4].value, r[5].value, f"{rel(QU_F)}｜In!D{r[0].row}:F{r[0].row}")
assert len(QU) == 6, QU.keys()
QSH = None
for r in QWB["Units"].iter_rows(min_row=4):
    if r[0].value == "廣達有效份額":
        QSH = ([r[2 + j].value for j in range(4)], f"{rel(QU_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert QSH and all(v is not None for v in QSH[0]), QSH

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），整櫃價格與市場為美元；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝鴻海法說（經富果、鉅亨、優分析摘要）、月營收公告；公司的產品占比與展望為 Interested-party；券商估計皆為 Interested-party。組裝市場與 ODM 每櫃價格讀廣達模型（唯讀），口徑一致"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
FUG2 = "https://blog.fugle.tw/post/earnings-call-2317-2026-08-12"
FUG4 = "https://blog.fugle.tw/post/earnings-call-2317-2026-03-16"
UAN25 = "https://uanalyze.com.tw/articles/4369349056"
CNY1 = "https://news.cnyes.com/news/id/6458468"
KGI = "https://www.kgi.com.hk/zh-cn/-/media/files/kgishk/research-reports/tw-reports/hon-hai-precision-2317-tt20260812en.pdf"
TN09 = "https://finance.technews.tw/2026/10/05/2026-sep-foxconn/"
TN08 = "https://finance.technews.tw/2026/09/05/hon-hai-2317-202608-financial-report/"
UAN09 = "https://uanalyze.com.tw/articles/4830556496"
INSIDE = "https://www.inside.com.tw/article/42074-foxconn-q2-2026-ai-server-revenue-passes-half"
UAN2Q25 = "https://uanalyze.com.tw/articles/5938328761"
MS0515 = "https://finance.technews.tw/2026/05/15/ai-server-800g-switch-foxconn/"
BNEXT = "https://www.bnext.com.tw/article/90315/2026-foxconn-earnings-call"
UBS0722 = "https://money.udn.com/money/amp/story/5607/9644763"
BT0818 = "https://www.businesstoday.com.tw/article/category/183008/post/202608170022/"
UDN0910 = "https://money.udn.com/money/amp/story/5607/9745501"
UDN0827 = "https://money.udn.com/money/story/5739/9716181"

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A 基準，B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；除以每 GW IT 資本支出得到部署 GW（與廣達、台達電模型同口徑）" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合用）", "%", QU["vr26"][0], QU["vr26"][1], QU["vr26"][2], "Derived", QU["vr26"][3], "廣達模型輸入（TrendForce 2026 組合 Rubin 22%）；鴻海：LPU 與 Vera 相關產品 4Q26 逐步放量（2Q26 法說）")
inp("racks_gb", "每 GW 機架數（GB300）", "架", None, BR["racks"][0], None, "Derived", BR["racks"][2], "Tokenomics IF_RacksPerGW（經橋接表）")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "同上")
inp("rack_gb", "整櫃買方價格（GB300，Tokenomics）", "US$M/架", None, BR["rack"][0], None, "Derived", BR["rack"][2], "參考：ODM 價 ÷ 買方價")
inp("rack_vr", "整櫃買方價格（VR200，Tokenomics）", "US$M/架", None, BR["rack"][1], None, "Derived", BR["rack"][2], "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、NVIDIA、廣達模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", QU["fx"][0], QU["fx"][1], QU["fx"][2], "Derived", QU["fx"][3], "與廣達、台達電模型相同")

sec("B. 組裝價格（讀廣達模型，唯讀）、鴻海份額與客供")
inp("odm_gb", "ODM 每櫃價格 GB300（廣達模型）", "US$M/架", QU["odm_gb"][0], QU["odm_gb"][1], QU["odm_gb"][2], "Derived", QU["odm_gb"][3], "廣達模型標記 Interested-party（工商時報、FTNN 引述法人）；本模型不另訂")
inp("odm_vr", "ODM 每櫃價格 VR200（廣達模型）", "US$M/架", QU["odm_vr"][0], QU["odm_vr"][1], QU["odm_vr"][2], "Derived", QU["odm_vr"][3], "同上")
inp("vpg_28", "每 GW ODM 整櫃價值 2028 ÷ 2027（Rubin Ultra；廣達模型）", "x", QU["vpg_28"][0], QU["vpg_28"][1], QU["vpg_28"][2], "Assumed", QU["vpg_28"][3], "廣達模型的 Assumed：Rubin Ultra 每櫃價格與每 GW 架數未找到")
inp("vpg_29", "每 GW ODM 整櫃價值 2029 ÷ 2027（廣達模型）", "x", QU["vpg_29"][0], QU["vpg_29"][1], QU["vpg_29"][2], "Assumed", QU["vpg_29"][3], "同上")
inp("rel_27", "鴻海 AI 伺服器有效份額 2027 ÷ 2026（相對值）", "x", *SIL.vals("asm_fxrel_27"), "Assumed", "公司：Vera Rubin 市占目標五成、ASIC 市佔目標四成以上（2Q26 法說，富果）", UDN0827 + "：法人預期鴻海 VR200 整體市占 40% 以上、部分客戶超過 50%（Interested-party）")
inp("rel_28", "鴻海份額 2028 ÷ 2026", "x", *SIL.vals("asm_fxrel_28"), "Assumed", "同上；廣達、緯創同時擴產（廣達模型假設廣達份額 0.95）", "—")
inp("rel_29", "鴻海份額 2029 ÷ 2026", "x", *SIL.vals("asm_fxrel_29"), "Assumed", "同上", "—")
inp("cs_27", "客供比例 2027（AI 毛額價值中改為客戶供料、不計營收的比例）", "%", 0.0, 0.10, 0.25, "Assumed", FUG2 + "：客供目前主要用於 ASIC 專案，未來 GPU 與 ASIC 都可能採用、占比逐步提升（公司；未給比例）", CNY1 + "：1Q26 雲端網路因營收認列模式改變低於預期、不影響獲利（公司）；與廣達模型同值")
inp("cs_28", "客供比例 2028", "%", 0.0, 0.20, 0.40, "Assumed", "同上", "—")
inp("cs_29", "客供比例 2029", "%", 0.0, 0.25, 0.50, "Assumed", "同上", "—")
inp("keep", "客供部分保留的毛利比例（相對買賣模式）", "x", 0.70, 0.90, 1.0, "Assumed", "公司：營收認列模式改變不影響獲利（1Q26 法說，鉅亨）", "與廣達模型同值；量化：未找到")

sec("C. 營收結構錨點（公司揭露產品占比 × 營收）")
inp("cl_25", "2025 年雲端網路產品 ÷ 營收", "%", None, 0.40, None, "Verified", FUG4 + "：2025 雲端網路 40%、消費智能 38%、電腦終端 15%、元件及其他 7%（公司）", UAN25 + "：同")
inp("ce_25", "2025 年消費智能產品 ÷ 營收", "%", None, 0.38, None, "Verified", "同上", "—")
inp("pc_25", "2025 年電腦終端產品 ÷ 營收", "%", None, 0.15, None, "Verified", "同上（元件及其他＝殘差 7%）", "—")
inp("ais_25", "2025 年 AI 伺服器 ÷ 雲端網路產品", "%", 0.45, 0.50, 0.55, "Interested-party", UAN2Q25 + "：2Q25 AI 伺服器占雲端網路超過五成（公司）", "2023 約 22%、2024 約三成（數位時代 2026-03-16）；前三季 AI 伺服器營收已破兆元（世界新聞網 2025-11-12）")
inp("cl_26", "2026 年雲端網路產品 ÷ 營收", "%", 0.51, 0.53, 0.56, "Interested-party", FUG2 + "：1Q26 48%、2Q26 51%；3Q26 雲端網路季增、年增強勁成長（公司）", KGI + "：2026F 55%（凱基）")
inp("ais_26", "2026 年 AI 伺服器 ÷ 雲端網路產品", "%", 0.52, 0.60, 0.68, "Interested-party", MS0515 + "：AI 伺服器已占 1Q26 雲端網路營收超過 50%（摩根士丹利引述）", "3Q26 AI 機櫃營收季增 30% 以上（公司，經凱基）；全年 AI 機櫃出貨倍數以上成長（公司）")
inp("ce_26", "2026 年消費智能產品 ÷ 營收", "%", 0.28, 0.30, 0.32, "Interested-party", FUG2 + "：1Q26 33%、2Q26 29%；3Q26 季增、年增顯著成長（公司）", KGI + "：2026F 29.4%")
inp("pc_26", "2026 年電腦終端產品 ÷ 營收", "%", 0.11, 0.12, 0.13, "Interested-party", FUG2 + "：1Q26 14%、2Q26 15%；3Q26 季減（公司）", KGI + "：2026F 11.2%；元件及其他＝殘差")
inp("g_cn27", "雲端網路非 AI（網通、一般伺服器）年增 2027", "%", 0.0, 0.12, 0.25, "Assumed", "公司：800G 以上交換機 2026 營收倍增；通用伺服器 2H26 雙位數以上成長（2Q26 法說）", "AI 以外的雲端網路拆分未揭露：未找到")
inp("g_cn28", "雲端網路非 AI 年增 2028", "%", -0.05, 0.08, 0.20, "Assumed", "同上", "—")
inp("g_cn29", "雲端網路非 AI 年增 2029", "%", -0.05, 0.06, 0.15, "Assumed", "同上", "—")
inp("g_ce27", "消費智能（iPhone 等）年增 2027", "%", -0.08, 0.03, 0.08, "Assumed", "公司：2026 消費智能顯著成長（2025 年報法說）", KGI + "：2027F 消費電子 +2%")
inp("g_ce28", "消費智能年增 2028", "%", -0.08, 0.02, 0.06, "Assumed", "同上", "—")
inp("g_ce29", "消費智能年增 2029", "%", -0.08, 0.02, 0.06, "Assumed", "同上", "—")
inp("g_pc27", "電腦終端年增 2027", "%", -0.10, -0.02, 0.05, "Assumed", "公司：2026 略微衰退（記憶體漲價）", KGI + "：2027F 電腦 −5%")
inp("g_pc28", "電腦終端年增 2028", "%", -0.10, 0.0, 0.05, "Assumed", "同上", "—")
inp("g_pc29", "電腦終端年增 2029", "%", -0.10, 0.0, 0.05, "Assumed", "同上", "—")
inp("g_co27", "元件及其他年增 2027", "%", -0.05, 0.08, 0.30, "Assumed", "公司：2026 大致持平至成長；3Q26 年增強勁成長（月營收公告）", KGI + "：2027F 元件及其他 +93%（含電動車與半導體，Interested-party；本版不採用）")
inp("g_co28", "元件及其他年增 2028", "%", -0.05, 0.06, 0.20, "Assumed", "同上", "—")
inp("g_co29", "元件及其他年增 2029", "%", -0.05, 0.05, 0.15, "Assumed", "同上", "—")

sec("D. 毛利率")
inp("d_ai", "AI 伺服器毛利率（毛額基礎）減其他營收毛利率", "pt", -0.015, -0.005, 0.005, "Assumed", "公司毛利率 FY2025 6.15%（雲端占 40%）、1H26 6.15%（雲端約 50%）：AI 占比上升而毛利率大致持平，差距小", "部門別毛利率未揭露：未找到；4Q25 毛利率 5.88% 受半成品模式伺服器比重影響（公司）")
inp("vrd", "VR 世代起 AI 毛利率（毛額基礎）變動（FY2027 起）", "pt", -0.005, -0.002, 0.0, "Assumed", "凱基 2026-08-12：毛利率 2026F 5.9% → 2027F 5.7%", "瑞銀 2026-07-22：2026、2027 皆 6.1%；廣達模型 −0.3 個百分點")

sec("E. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收合計）", "NT$B", None, 3026.9, None, "Verified", TN09 + "：3Q26 3 兆 269 億元、季增 20.44%、年增 47.12%；前 9 月 7 兆 6,698 億元（公司月營收）", UAN09 + "：3Q26 3.02 兆、9 月 1.15 兆（7 月 9,465 億、8 月 9,217.66 億，TechNews 2026-09-05）")
inp("q4g", "4Q26 營收季增", "%", 0.0, 0.15, 0.25, "Interested-party", "公司：AI 持續成長加上 ICT 旺季，4Q 整體營運可望符合目前市場預期（10-05 月營收公告）", KGI + "：4Q26F 3,536（以 3Q26F 2,829 為基礎，季增 25%）；4Q25 季增 27%（公司）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.057, 0.060, 0.063, "Interested-party", "1Q26 6.18%、2Q26 6.12%（公司）；4Q25 5.88%（半成品模式影響）", KGI + "：2026F 5.9%；瑞銀 2026F 6.1%")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.020, 0.0215, 0.024, "Assumed", "1Q26 約 2.61%、2Q26 約 2.36%（公司毛利 − 營業利益推得）；營收放大後費用率下降", "公司：2026 研發近 200 億元、年增 40%（2Q26 法說）")
inp("q_nio", "3Q、4Q26 淨利 ÷ 營業利益", "%", 0.60, 0.66, 0.72, "Derived", "1H26 109.89 ÷ 170.45＝0.645；FY2025 189.35 ÷ 259.22＝0.730（公司）", "公司：2026 平均稅率 24–26%（2025 約 20%）；1Q26 業外損失 16.2 億（匯兌、利息）；凱基 2026F 0.658")
inp("q_sh", "股數（十億股；EPS 換算）", "十億股", None, "=(49.919+59.974)/7.84", None, "Derived", "1H26 淨利 499.19＋599.74 億 ÷ EPS 7.84（公司）", "FY2025 189.354 ÷ 13.61＝13.913；凱基流通股數 14,076 百萬股")

sec("F. 營業費用、淨利係數、股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.2, 0.4, 0.7, "Derived", "FY2026 錨點實測（Q 頁 B21）", "營業費用率 FY2025 2.95% → 2Q26 2.36%；研發年增 40%，基準取高於實測")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.002, 0.0, 0.002, "Assumed", "無", "營益率約 3.8%，±0.2 個百分點約 ±5% 的營業利益")
inp("nio_f", "淨利 ÷ 營業利益", "%", 0.60, 0.66, 0.73, "Derived", "同 q_nio（1H26 0.645、FY2025 0.730）", "凱基隱含 2026F 0.658、2027F 0.663（稅率 25%、少數股權 380–440 億）")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.01, "Assumed", "FY2025 13.913 → 1H26 14.017（可轉債轉換）", "—")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("G. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("b_rev26", "券商 FY2026 營收（凱基 08-12，3Q 前）", "NT$B", None, 11010.682, None, "Interested-party", KGI, "3Q26 實績 3,026.9 高於凱基 3Q26F 2,829")
inp("b_eps26", "券商 FY2026 EPS（凱基 08-12）", "NT$", None, 18.67, None, "Interested-party", KGI, "瑞銀 07-22：18.0；摩根士丹利 05-15：16.38")
inp("b_rev27", "券商 FY2027 營收（凱基 08-12）", "NT$B", None, 14019.401, None, "Interested-party", KGI, "瑞銀：2027 年增 28%")
inp("b_eps27a", "券商 FY2027 EPS（凱基 08-12）", "NT$", None, 22.64, None, "Interested-party", KGI, "—")
inp("b_eps27b", "券商 FY2027 EPS（瑞銀 07-22）", "NT$", None, 22.0, None, "Interested-party", UBS0722, "目標價 320＝15x 2027 EPS")
inp("b_eps27c", "券商 FY2027 EPS（法人，08-18）", "NT$", None, 25.0, None, "Interested-party", BT0818 + "（今周刊；未具名法人：2026 19、2027 25、2028 28）", "—")
inp("b_eps28", "券商 FY2028 EPS（法人，08-18）", "NT$", None, 28.0, None, "Interested-party", BT0818, "摩根士丹利 05-15：23.03")
inp("b_gm27", "券商 FY2027 毛利率（凱基）", "%", None, 0.057, None, "Interested-party", KGI, "營益率 3.4%")
inp("rk_sh", "三家 ODM 中鴻海的 NVL72 櫃數份額（8 月調查）", "%", None, "=2600/(2600+2525+2350)", None, "Interested-party", "經濟日報 2026-09-10 引述美系券商：鴻海約 2,600 櫃、廣達 2,500–2,550、緯創 2,300–2,400（網址見 Sources）", "櫃數口徑；單月、三家合計（不含 Dell、Supermicro 等）")
inp("rk_sh2", "鴻海占 NVIDIA AI 伺服器機櫃的全球份額（法人估）", "%", None, 0.44, None, "Interested-party", INSIDE + "：法人估約 44%（公司說 40% 以上）", "櫃數口徑（推測；原文未寫明櫃數或金額）")

sec("H. 價格與估值")
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 245, None, "Interested-party", "聚財網 2317 日成交（10-22 收盤 245、10-21 250）", "Yahoo 股市：10-16 收盤 244")
inp("pe_ttm", "近 12 月 P/E（Yahoo，10-16）", "x", None, 16.08, None, "Interested-party", "Yahoo 股市 2317（10-16）", "Goodinfo 目前 P/E 16.14；companiesmarketcap TTM 15.3")

sec("I. 目標價情境權重（Val 第 0 節用；Andy 可改）")
for s_, w_, nm_ in (("A", 0.3, "順利"), ("B", 0.5, "平平"), ("C", 0.2, "不如預期")):
    inp(f"tp_w{s_}", f"目標價情境權重：情境 {s_}（{nm_}）", "權重", None, w_, None, "Assumed",
        "Andy 可改（2026-10-25 目標價規格預設 A／B／C＝0.3／0.5／0.2）", "三個權重合計應為 1；v0.2 起用於 DCF 與 Comparable 的情境加權（Val 第 0 節）")

# ---------------- v0.2 設定與新增輸入（附加在 In 既有輸入之後，不位移既有列） ----------------
V2 = {'name': '鴻海',
 'unit': 'NT$B',
 'shunit': '十億股',
 'cur': 'NT$',
 'shg': 'sh_g',
 'epsfmt': '0.00',
 'pxfmt': '#,##0',
 'keys': {'rev': 'rev', 'ni': 'ni', 'sh': 'sh_', 'eps': 'eps'},
 'roe': ('=AVERAGE(0.0928,0.0971,0.1125)', 'stockanalysis 比率頁（Return on Equity）：FY2023 9.28%、FY2024 9.71%、FY2025 11.25%', '單一來源；近 12 月 13.15%（同站，2026-09-18）'),
 'cons_tp': (347.45, 'stockanalysis forecast 頁（S&P Global）21 位分析師平均 347.45（中位數 338.50），頁面時間 2026-10-12', '第二來源：未找到'),
 'cons_feps': ('=239.5/11.37', 'stockanalysis 個股頁 2026-10-19：價格 239.5 ÷ Forward PE 11.37'),
 'peers': [('廣達（Quanta）', '2382.TW', 11.31, '2026-10-14'),
           ('緯創（Wistron）', '3231.TW', 9.93, '2026-09-29'),
           ('和碩（Pegatron）', '4938.TW', 14.88, '2026-10-22'),
           ('Jabil', 'JBL', 15.49, '2026-10-23'),
           ('Flex', 'FLEX', 14.74, '2026-10-23')],
 'peer_note': '電子代工與 AI 伺服器組裝同業：廣達、緯創、和碩、Jabil、Flex（券商常用比較組）。工業富聯（601138.SS，鴻海子公司）為 A 股，依規則不列。非 A 股公司的同業名單不放中國 A 股（A 股有結構性溢價；Andy 2026-10-25 11:53）',
 'fwd_w': (0.16666666666666666, 0.8333333333333334),
 'fwd_note': '約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）',
 'pe10': None,
 'px_note': '現價＝In 頁 px（2026-10-22 收盤 245；10-23 收盤未找到）',
 'method': 'DCF：股東現金流＝EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE＝近三年平均約 10%（再投資率約 40%），FY2027–29 驅動因子、FY2030–32 收斂法＋永續。Comparable：廣達、緯創、和碩、Jabil、Flex 前瞻 P/E 中位數；鴻海含消費電子與電動車等低倍數事業，同業含 '
           'EMS，偏誤方向不明確。情境對應：A＝順利、B＝平平、C＝不如預期（原顆數情境）。',
 'readme_val': '估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − tp_g ÷ tp_roe），FY2027–FY2032 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，推到 12 個月後）＋0.5 × Comparable 目標價（廣達、緯創、和碩、Jabil、Flex 前瞻 P/E 中位數 × 12 '
               '個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解（EPS 因素、倍數因素）、DCF 價格帶、十年 P/E 分位檢查（只作檢查）。',
 'readme_add': ['v0.2（2026-10-25）：(1) PL 頁最下方新增延伸表，三情境延伸到 FY2032；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt；淨利率由 FY2029 收斂到 A＝FY2029、B＝FY2029 '
                '與十年平均的中點、C＝十年平均；十年平均淨利率未找到，依工作單以 FY2029 × 0.8 替代，Assumed；PL 延伸表灰底）。(2) Val 第 0 節目標價＝DCF 與 Comparable 加權（tp_w_dcf＝0.5）；v0.1.1 的「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留，不再是目標價。(3) Sens '
                '新增目標價參數（tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.4／0.8）。(4) Recon 新增目標價對帳。(5) 現價：10-23 收盤未找到（Yahoo 股市、財報狗於 2026-10-25 仍為 10-22 或更早），依工作單沿用 2026-10-22 收盤 245。上游廣達模型改為 v0.2，本模型讀取的 In、Units '
                '數值不變。 新增內容一律放在各頁既有內容之後，FY2025–FY2029 既有數值不變。',
                '情境對應：A＝順利（顆數基準）、B＝平平（CY2028 顆數 ×0.8、CY2029 ×0.85）、C＝不如預期（CY2028–29 ×0.6），即原本的三個顆數情境。三情境在 FY2028 前完全相同（只差顆數係數，從 CY2028 起），DCF 三情境仍照算，差異來自 FY2028 之後（含 FY2030 起的收斂法）。',
                '先驗：本版未做先驗，無 Andy 合理 P/E 對照。']}
sec("J. 目標價（DCF 與 Comparable 加權）與收斂法參數（v0.2 新增；黃底＝Assumed，Andy 可改）")
inp("tp_r", "折現率／要求報酬（DCF）", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%；美、台、日、韓同一值）", "Sens 頁 ±1pt")
inp("tp_g", "長期成長率（DCF 永續；情境 B 的 FY2032 營收年增收斂目標）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "Sens 頁 ±1pt（只動 DCF，PL 收斂目標不隨之變動）")
inp("tp_roe", "ROE（近三年平均）", "%", None, V2["roe"][0], None, "Derived", V2["roe"][1], V2["roe"][2])
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 改為 0.5：DCF／Comparable 各半）", "Sens 頁 0.4／0.8")
inp("fade_gA", "情境 A（順利）：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂，FY2032 達到")
inp("fade_gB", "情境 B（平平）：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C（不如預期）：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("nim10_f", "十年平均淨利率的替代：最後一年（FY2029）淨利率 ×（本值）", "x", None, 0.8, None, "Assumed", "十年（2016–2025）平均淨利率：未找到（免費資料商只到近五年）；依工作單以最後一年 × 0.8 替代", "淨利率收斂目標：A＝FY2029、B＝FY2029 與替代值的中點、C＝替代值（PL 延伸表）")
inp("cons_tp", "市場共識目標價（賣方平均）", V2["cur"], None, V2["cons_tp"][0], None, "Interested-party", V2["cons_tp"][1], V2["cons_tp"][2])
inp("cons_feps", "共識前瞻 12 月 EPS（＝價格 ÷ 前瞻 P/E）", V2["cur"], None, V2["cons_feps"][0], None, "Interested-party", V2["cons_feps"][1], "賣方共識，只作對帳與差距拆解，不作校準")

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
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自鴻海法說摘要、凱基報告（實績欄）與月營收"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（月營收）＋3Q、4Q 假設（毛利率、營業費用率、淨利 ÷ 營業利益；4Q 營收＝3Q × (1＋季增)）。1H26 營收＝前 9 月 − 3Q；1Q26 毛利由公司毛利率 6.18% × 營收推得。業外與稅前：只有 1Q26 業外（損失 16.2 億）與凱基 FY2025 稅前 2,934 億，季別未找到"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：" + UAN25 + "；1H26：" + FUG2 + "、" + CNY1
S2 = "3Q26 營收：月營收公告（TechNews 2026-10-05）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 8103.105, "C": "=7669.8-q3_rev", "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Interested-party（4Q 季增）", s2="FY2025 8,103.105（凱基實績欄；公司 8.1 兆）；1H26＝前 9 月 7,669.8 − 3Q（2Q26 2,525.894，凱基）")
qrow(5, "毛利", {"B": 498.161, "C": "=(C4-2525.894)*0.0618+154.533", "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Derived（1Q26 毛利率 × 營收）", s2="FY2025 毛利率 6.15%；1Q26 6.18%；2Q26 毛利 1,545.33 億（公司）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived（毛利 − 營業利益）／Assumed")
qrow(8, "營業利益", {"B": 259.223, "C": "=75.649+94.803", "D": "=D5-D7", "E": "=E5-E7"}, tag="Verified／Derived", s2="FY2025 2,592.23 億、營益率 3.20%；1Q26 756.49 億、2Q26 948.03 億（公司）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "歸屬母公司淨利", {"B": 189.354, "C": "=49.919+59.974", "D": "=D8*q_nio", "E": "=E8*q_nio"}, tag="Verified／Derived（3Q、4Q）", s2="FY2025 1,893.54 億；1Q26 499.19 億、2Q26 599.74 億（公司稅後淨利；凱基列為母公司淨利，少數股權另列）")
qrow(11, "淨利 ÷ 營業利益（含業外、稅、非控制權益）", {c: f"={c}10/{c}8" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(12, "股數（十億股）", {"B": "=189.354/13.61", "C": "=q_sh", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s2="淨利 ÷ EPS")
qrow(13, "EPS", {c: f"={c}10/{c}12" for c in "BCDE"}, fmt="0.00", ann="=F10/F12", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(14, "  公司揭露 EPS", {"B": 13.61, "C": 7.84}, fmt="0.00", ann="", tag="Verified", s2="1Q26 3.56、2Q26 4.27")
qrow(15, "AI 伺服器營收（＝營收 × 雲端網路占比 × AI 占雲端網路）", {"B": "=B4*cl_25*ais_25"}, ann="=F4*cl_26*ais_26", tag="Derived／Interested-party", s1="公司產品占比（In 頁 C 節）", s2="各季與年度 AI 伺服器營收金額未揭露：未找到")
qrow(16, "雲端網路非 AI（網通、一般伺服器）", {"B": "=B4*cl_25-B15"}, ann="=F4*cl_26-F15", tag="Derived", s1="—", s2="—")
qrow(17, "消費智能（iPhone 等）", {"B": "=B4*ce_25"}, ann="=F4*ce_26", tag="Derived／Interested-party", s1="—", s2="—")
qrow(18, "電腦終端", {"B": "=B4*pc_25"}, ann="=F4*pc_26", tag="Derived／Interested-party", s1="—", s2="—")
qrow(19, "元件及其他（殘差）", {"B": "=B4-B15-B16-B17-B18"}, ann="=F4-F15-F16-F17-F18", tag="Derived", s1="—", s2="FY2025 公司揭露 7%")
qrow(20, "非 AI 營收合計", {"B": "=B4-B15"}, ann="=F4-F15", tag="Derived", s1="—", s2="—")
Q.cell(22, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(23, 1, "FY2026 營收年增"); Q["B23"] = "=F4/B4-1"
Q.cell(24, 1, "FY2026 營業費用年增"); Q["B24"] = "=F7/B7-1"
Q.cell(25, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B25"] = "=B24/B23"
Q.cell(26, 1, "FY2026 營收 ÷ 凱基（08-12，3Q 前）− 1"); Q["B26"] = "=F4/b_rev26-1"
Q.cell(27, 1, "FY2026 EPS ÷ 凱基 − 1"); Q["B27"] = "=F13/b_eps26-1"
Q.cell(28, 1, "AI 伺服器營收 FY2026 年增（隱含；公司說 AI 機櫃出貨倍數以上成長）"); Q["B28"] = "=F15/B15-1"
for rr in range(23, 29): Q[f"B{rr}"].number_format = "0.0%" if rr != 25 else "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("ni", 10), ("sh", 12), ("eps", 13), ("ai", 15), ("cn", 16), ("ce", 17), ("pc", 18), ("co", 19), ("nai", 20)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 26; Q.column_dimensions["H"].width = 90; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
DEFP = dict(scn="A", rel="rel", cs="cs", keep="keep", vpg="vpg", mf="1", gcn="g_cn", gce="g_ce", gpc="g_pc", gco="g_co",
            d="d_ai", vrd="vrd", opxa="0", nio="nio_f")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    sec_("A. AI 部署與整櫃組裝市場（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029；與廣達模型同口徑）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in BU[1:]}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    row("vpg", "每 GW ODM 整櫃價值", "US$B/GW", {"D": "=(racks_gb*odm_gb*(1-vr26)+racks_vr*odm_vr*vr26)/1000", "E": "=racks_vr*odm_vr/1000",
                                                  **{c_(y): f"=$E{r}*{P['vpg']}_{y % 100}" for y in BU[2:]}}, fmt="0.00",
        note="架數 × ODM 每櫃價格；含 ASIC 機架（Assumed：每 GW 價值相同）")
    row("mku", "AI 伺服器組裝市場（ODM 價，美元）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['vpg']}" for y in BU}, fmt="#,##0.0")
    row("mkn", "AI 伺服器組裝市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['mku']}*fx" for y in BU}, fmt="#,##0.0")
    row("sh0", "鴻海有效份額（FY2026 校準）＝ AI 伺服器營收 ÷ 市場", "%", {"D": f"=a26_ai/D{R['mkn']}"}, fmt="0.0%",
        note="Derived：分子為公司雲端網路占比 × AI 占雲端網路 × 營收（金額口徑）；分母含 ASIC、HGX 機架，所以低於整櫃櫃數份額")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "鴻海有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("aiv", "AI 伺服器毛額價值（買賣模式下的營收）", "NT$B", {c_(y): f"={c_(y)}{R['mkn']}*{c_(y)}{R['sh']}" for y in BU}, fmt="#,##0.0")
    row("cs", "客供比例（FY2027 起新增部分）", "%", {"D": "=0", **{c_(y): f"={P['cs']}_{y % 100}" for y in BU[1:]}}, fmt="0.0%",
        note="FY2026 已有的半成品／客供已含在錨點與校準份額內，記為 0")
    row("ai", "AI 伺服器營收", "NT$B", {"C": "=a25_ai", **{c_(y): f"={c_(y)}{R['aiv']}*(1-{c_(y)}{R['cs']})" for y in BU}}, fmt="#,##0.0")
    sec_("B. 非 AI 事業（簡化成長率）與合計")
    for k, lab, g in (("cn", "雲端網路非 AI（網通、一般伺服器）", "gcn"), ("ce", "消費智能（iPhone 等）", "gce"),
                      ("pc", "電腦終端", "gpc"), ("co", "元件及其他", "gco")):
        row(k, lab, "NT$B", {"C": f"=a25_{k}", "D": f"=a26_{k}", **{c_(y): f"={c_(y - 1)}{r}*(1+{P[g]}{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("nai", "非 AI 營收合計", "NT$B", {c_(y): f"=SUM({c_(y)}{R['cn']}:{c_(y)}{R['co']})" for y in FYS}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in FYS}, fmt="#,##0.0")
    row("cl", "雲端網路產品（官方產品類別＝AI 伺服器＋雲端網路非 AI）", "NT$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['cn']}" for y in FYS}, fmt="#,##0.0",
        note="官方四大產品類別：雲端網路產品、消費智能產品（上列「消費智能」）、電腦終端產品（「電腦終端」）、元件及其他")
    row("chk", "檢查：四大產品類別合計 − 合併營收（FY2025 實績、FY2026 錨點）", "NT$B",
        {"C": f"=C{R['cl']}+C{R['ce']}+C{R['pc']}+C{R['co']}-a25_rev", "D": f"=D{R['cl']}+D{R['ce']}+D{R['pc']}+D{R['co']}-a26_rev",
         **{c_(y): f"={c_(y)}{R['cl']}+{c_(y)}{R['ce']}+{c_(y)}{R['pc']}+{c_(y)}{R['co']}-{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 伺服器 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("clsh", "雲端網路產品 ÷ 營收", "%", {c_(y): f"=({c_(y)}{R['ai']}+{c_(y)}{R['cn']})/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pgw", "每 GW 鴻海 AI 伺服器營收", "US$B/GW", {c_(y): f"={c_(y)}{R['ai']}/fx/{c_(y)}{R['gw']}" for y in BU}, fmt="0.00")
    row("rk", "隱含鴻海出貨櫃數（毛額價值 ÷ ODM 每櫃價格；NVL72 等值）", "千櫃",
        {"D": f"=D{R['aiv']}/fx/(odm_gb*(1-vr26)+odm_vr*vr26)", **{c_(y): f"={c_(y)}{R['aiv']}/fx/odm_vr" for y in BU[1:]}}, fmt="#,##0.0",
        note="十億美元 ÷ 百萬美元 ＝ 千櫃；FY2028 起若每 GW 價值改變，每櫃價格仍以 VR200 換算（參考）")
    sec_("C. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × AI 占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['aish']}"}, fmt="0.00%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 毛利率（毛額基礎）＝其他 + 差距")
    row("gai", "AI 伺服器毛利 ＝ AI 毛利率 × 毛額價值 ×（1 − 客供 ×（1 − 保留比例））", "NT$B",
        {c_(y): f"=($D${R['gmn']}+{P['d']}" + ("" if y == 2026 else f"+{P['vrd']}") + f")*{c_(y)}{R['aiv']}*(1-{c_(y)}{R['cs']}*(1-{P['keep']}))" for y in BU}, fmt="#,##0.00",
        note="客供部分不計營收，但大部分毛利金額保留（keep）；FY2027 起加 VR 稀釋 vrd")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"=$D${R['gmn']}*{c_(y)}{R['nai']}+{c_(y)}{R['gai']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("gma", "AI 伺服器毛利率（營收基礎）", "%", {c_(y): f"={c_(y)}{R['gai']}/{c_(y)}{R['ai']}" for y in BU}, fmt="0.00%",
        note="客供比例上升 → 營收基礎毛利率上升")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,phi*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("nio", "淨利 ÷ 營業利益", "%", {"C": "=a25_ni/a25_oi", "D": "=a26_ni/a26_oi", **{c_(y): f"={P['nio']}" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"={c_(y)}{R['oi']}*{c_(y)}{R['nio']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.000")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 鴻海損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 75
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 伺服器營收", "NT$B", "#,##0.0"),
                                ("cl", "  雲端網路產品", "NT$B", "#,##0.0"), ("ce", "  消費智能產品", "NT$B", "#,##0.0"), ("pc", "  電腦終端產品", "NT$B", "#,##0.0"), ("co", "  元件及其他", "NT$B", "#,##0.0"),
                                ("nai", "  非 AI 營收", "NT$B", "#,##0.0"), ("aish", "  AI 伺服器 ÷ 營收", "%", "0.0%"), ("gp", "毛利", "NT$B", "#,##0.0"),
                                ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"),
                                ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收 − 錨點；FY2026 模型 AI 伺服器 − 錨點；FY2026 毛利 − 錨點").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：AI 伺服器差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['ai']}-a26_ai"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：毛利差（NT$B；以 FY2027 公式套用 FY2026）"); PL[f"D{sr}"] = f"=D{PB['A']['gmn']}*D{PB['A']['nai']}+D{PB['A']['gai']}-a26_gp"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 75, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：組裝市場、份額、客供 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 伺服器組裝市場（與廣達模型同口徑）、鴻海份額、客供比例。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）；ODM 每櫃價格讀廣達模型"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
QSH_V = dict(zip(BU, QSH[0]))
for key, lab, unit, f, fmt, note in (
        ("gw", "部署 GW（情境 A）", "GW", lambda y: f"=PL!{FC[y]}{A_['gw']}", "0.0", "AI IT 資本支出 ÷ 每 GW IT 資本支出"),
        ("vpg", "每 GW ODM 整櫃價值", "US$B/GW", lambda y: f"=PL!{FC[y]}{A_['vpg']}", "0.00", "架數 × ODM 每櫃價格（讀廣達模型）"),
        ("mku", "AI 伺服器組裝市場（ODM 價）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mku']}", "#,##0.0", "含 ASIC 機架（假設每 GW 價值相同）"),
        ("sh", "鴻海有效份額（金額口徑）", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "FY2026 以公司雲端網路占比與 AI 占比校準"),
        ("aiv", "鴻海 AI 伺服器毛額價值", "US$B", lambda y: f"=PL!{FC[y]}{A_['aiv']}/fx", "#,##0.0", "買賣模式下的營收"),
        ("cs", "客供比例（FY2027 起新增）", "%", lambda y: f"=PL!{FC[y]}{A_['cs']}", "0.0%", "Assumed"),
        ("ai", "鴻海 AI 伺服器營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx", "#,##0.0", ""),
        ("rk", "隱含鴻海出貨櫃數（NVL72 等值）", "千櫃", lambda y: f"=PL!{FC[y]}{A_['rk']}", "#,##0.0", "毛額價值 ÷ ODM 每櫃價格"),
        ("pgw", "每 GW 鴻海 AI 伺服器營收", "US$B/GW", lambda y: f"=PL!{FC[y]}{A_['pgw']}", "0.00", ""),
        ("gma", "鴻海 AI 伺服器毛利率（營收基礎）", "%", lambda y: f"=PL!{FC[y]}{A_['gma']}", "0.00%", "客供上升 → 毛利率上升、毛利金額大致不變"),
        ("qsh", "廣達有效份額（廣達模型 Units，唯讀）", "%", lambda y: QSH_V[y], "0.0%", "來源：" + QSH[1]),
        ("tsh", "鴻海＋廣達合計有效份額", "%", lambda y: f"={cyc(y)}7+{cyc(y)}14", "0.0%", "其餘為緯創、英業達、Dell、Supermicro、ASIC 自有組裝等"),
        ("oth", "其他廠商（緯創、英業達、Dell、Supermicro 等）＝ 市場 ×（1 − 合計份額）", "US$B", lambda y: f"={cyc(y)}6*(1-{cyc(y)}15)", "#,##0.0", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
        if key == "qsh": UN[f"{cyc(y)}{ur}"].font = BLUE
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
assert UR["sh"] == 7 and UR["qsh"] == 14 and UR["tsh"] == 15 and UR["mku"] == 6, UR
ur += 1
UN.cell(ur, 1, "讀法：鴻海有效份額（金額口徑，分母為全部 AI IT 資本支出換算的組裝市場，含 ASIC、HGX 與其他組裝廠）低於櫃數份額（三家 ODM 中約 35%、法人估 NVIDIA 機櫃約 44%），口徑不同。鴻海＋廣達合計約三成；份額相對值與客供比例是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」與十年 P/E 規則值（年底落後 P/E，基準為 FY2027 EPS）"
V["A2"] = "循環位置：鴻海不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照廣達、NVIDIA、台達電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-10-22 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "#,##0"
V.cell(7, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(8, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(9, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(10, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((7, 2026), (8, 2027), (9, 2028), (10, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(11, 1, "未來 12 月 EPS（模型 A）≈ 0.2 × FY2026＋0.8 × FY2027"); V["B11"] = f"=0.2*PL!D{SUMR[('A', 'eps')]}+0.8*PL!E{SUMR[('A', 'eps')]}"; V["B11"].number_format = "0.00"
V.cell(12, 1, "  現價 ÷ 未來 12 月 EPS（模型 A）"); V["B12"] = "=px/B11"; V["B12"].number_format = "0.0"
V.cell(13, 1, "近 12 月 P/E（Yahoo，10-16）"); V["B13"] = "=pe_ttm"; V["B13"].number_format = "0.0"
V["E8"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"
V.cell(15, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, 16, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 17):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(21, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, 22, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = 23 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{17 + j}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
V.cell(26, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, 27, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = 28
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。")
vr += 2
V.cell(vr, 1, "5. 先驗對照：本版未做先驗，無 Andy 合理 P/E 對照").font = BOLD
PE_CMC = [6.85, 10.8, 6.50, 8.55, 10.3, 8.83, 8.78, 9.59, 16.2, 16.8]       # companiesmarketcap 年底 P/E 2016–2025
PE_GI = ["=84.2/8.6", "=95.2/8.01", "=70.8/8.03", "=90.8/8.32", "=92/7.34", "=104/10.05", "=99.9/10.21", "=104.5/10.25", "=184/11.01", "=230.5/13.61"]  # Goodinfo 年底收盤 ÷ EPS
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2016–2025 共 10 年；價格可能為還原股價）"},
    {"label": "Goodinfo（年底收盤 ÷ EPS）", "vals": PE_GI, "tag": "Derived（資料商收盤價與 EPS；未還原除權息）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。兩來源 2016–2020 年相差約 1–3 個倍數點（companiesmarketcap 疑用還原股價，數字較低；2016–2017 年有股票股利）。2016–2023 年鴻海以 iPhone 組裝為主（約 7–12x），2024 年起 AI 伺服器重新定價（約 16–17x），規則值可能低估 AI 期的合理倍數。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# ---------------- Val 第 0 節：目標價與上檔空間（v0.1.1 新增；放在既有空白列，不位移既有列） ----------------
TP0 = 37; TPR = {}
assert all(V.cell(rw, c).value is None for rw in range(TP0 - 1, TP0 + 9) for c in range(1, 9)), TP0
assert V["A3"].value is None
V["A3"] = f"目標價與上檔空間：見第 0 節（本頁第 {TP0} 列；v0.1.1 新增，為了不位移既有列而放在空白列）"
V.cell(TP0, 1, "舊法：規則值情境加權（參考；v0.1.1 的「0. 目標價與上檔空間」改名，不再是目標價；權重為 In 頁 tp_wA／tp_wB／tp_wC）").font = BOLD
for k_, (key_, lab_) in enumerate((("tp", "舊法隱含價（情境加權，參考）"), ("up", "舊法隱含價相對現價（÷ 現價 − 1）"),
        ("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"),
        ("w", "情境權重 A／B／C"), ("m", "舊法說明")), TP0 + 1):
    V.cell(k_, 1, lab_); TPR[key_] = k_
V.cell(TPR["tp"], 2, f"=tp_wA*B{TPR['A']}+tp_wB*B{TPR['B']}+tp_wC*B{TPR['C']}").number_format = "#,##0"
V.cell(TPR["tp"], 5, "＝權重 A × 情境 A 隱含價＋權重 B × 情境 B 隱含價＋權重 C × 情境 C 隱含價（NT$）")
V.cell(TPR["up"], 2, f"=B{TPR['tp']}/px-1").number_format = "0.0%"
V.cell(TPR["up"], 5, "相對「" + V["A6"].value + "」")
for s_ in SC:
    V.cell(TPR[s_], 2, f"=E{VAL[s_]}").number_format = "#,##0"
    V.cell(TPR[s_], 5, f"＝規則值隱含股價 E{VAL[s_]}（{SCN_T[s_]}）")
for j_, s_ in enumerate(SC):
    V.cell(TPR["w"], 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(TPR["w"], 5, "In 頁輸入（Assumed，Andy 可改）：B 欄＝情境 A、C 欄＝情境 B、D 欄＝情境 C")
V.cell(TPR["m"], 2, "P/E 規則值：各情境 FY2027 EPS × 2016–2025 年底 P/E 依循環位置取分位（見頂＝25 分位、中段＝中位數、谷底＝75 分位；companiesmarketcap 與 Goodinfo 兩來源分別取分位再平均）。隱含價約為 FY2027 年底價、未折現，作 12 個月目標價的近似；情境 A／B／C 即本模型三個顆數情境（A 最高、C 最低）。偏誤：十年倍數含 2016–2023 年 iPhone 組裝期低倍數（約 7–12x），目標價可能偏低。")
for c_ in range(1, 6):
    V.cell(TPR["tp"], c_).fill = KEY; V.cell(TPR["up"], c_).fill = KEY
V.cell(TPR["tp"], 1).font = BOLD
defname("old_rule_px", f"Val!$B${TPR['tp']}"); defname("old_rule_up", f"Val!$B${TPR['up']}")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 伺服器營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 伺服器營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 伺服器營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.85、2028 0.75、2029 0.70（廣達、緯創搶單）", dict(rel="RLO")),
         ("份額上升（高）", "份額相對值 1.15（Vera Rubin 市占目標五成）", dict(rel="RHI")),
         ("客供 0（維持買賣）", "客供比例 2027–29 皆 0", dict(cs="CSLO")), ("客供 高", "客供比例 25%／40%／50%", dict(cs="CSHI")),
         ("客供保留毛利 0.7", "客供部分只保留 70% 毛利", dict(keep="keep_lo")), ("客供保留毛利 1.0", "毛利金額完全保留", dict(keep="keep_hi")),
         ("每 GW 價值 低", "2028–29 每 GW ODM 價值 ×0.85（Rubin Ultra 降價或客供更深）", dict(vpg="VPLO")), ("每 GW 價值 高", "×1.2", dict(vpg="VPHI")),
         ("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("消費智能 低", "2027–29 −8%（iPhone 衰退）", dict(gce="GCELO")), ("消費智能 高", "2027 +8%、2028–29 +6%", dict(gce="GCEHI")),
         ("雲端非 AI 低", "網通與一般伺服器 2027 0%、2028–29 −5%", dict(gcn="GCNLO")), ("雲端非 AI 高", "2027 +25%、2028 +20%、2029 +15%", dict(gcn="GCNHI")),
         ("AI 毛利率差距 −1.5pt", "AI 毛利率（毛額基礎）比其他低 1.5 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 +0.5pt", "高 0.5 個百分點", dict(d="d_ai_hi")),
         ("VR 稀釋 −0.5pt", "FY2027 起 AI 毛利率（毛額基礎）−0.5 個百分點", dict(vrd="vrd_lo")), ("VR 不稀釋", "FY2027 起不變", dict(vrd="vrd_hi")),
         ("營業費用率 +0.2pt", "營業費用額外 +0.2% × 營收", dict(opxa="opx_adj_hi")),
         ("淨利 ÷ 營業利益 0.60", "FY2027 起 0.60（業外損失擴大、稅率上升）", dict(nio="nio_f_lo")),
         ("淨利 ÷ 營業利益 0.73", "FY2027 起 0.73（回到 FY2025）", dict(nio="nio_f_hi"))]
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("CSLO", "cs", "_lo"), ("CSHI", "cs", "_hi"), ("VPLO", "vpg", "_lo"), ("VPHI", "vpg", "_hi"),
                      ("GCELO", "g_ce", "_lo"), ("GCEHI", "g_ce", "_hi"), ("GCNLO", "g_cn", "_lo"), ("GCNHI", "g_cn", "_hi")):
    for y in (27, 28, 29):
        if base == "vpg" and y == 27: continue
        src = f"{base}{y}{suf}" if base.startswith("g_") else f"{base}_{y}{suf}"
        target = f"{nm}{y}" if base.startswith("g_") else f"{nm}_{y}"
        wb.defined_names[target] = DefinedName(target, attr_text=wb.defined_names[src].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ai", 2027, "#,##0.0"), ("E", "gm", 2027, "0.00%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ai", 2028, "#,##0.0"), ("J", "gm", 2028, "0.00%"), ("K", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "GL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 40; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6", PB["C"], "PL!")
assert nn + 3 < 40
SENS_N = len(cases)
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。每個案例的 FY2026 有效份額與其他營收毛利率依該案例重算，所以 FY2026 恆等於錨點。鴻海 AI 伺服器約占營收三成，非 AI 事業（iPhone、網通、電腦）的成長率對營收影響大、對 EPS 影響小於 AI 部署與份額；淨利 ÷ 營業利益（業外、稅、少數股權）是單一最大的非營運變數。")
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 券商估計與櫃數份額。只作事後對照，不為對上數字而調參數。券商多為 3Q 營收公布前、部分未具名；FY2027 付費共識未找到"
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
RC.cell(rr_, 1, "A. 券商估計（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 凱基（08-12，3Q 前）", "=b_rev26", pl("rev", 2026), "凱基 2026-08-12", "3Q26 實績高於凱基估 7%", fmt="#,##0.0")
rec("FY2026 EPS vs 凱基", "=b_eps26", pl("eps", 2026), "凱基 2026-08-12", "瑞銀 18.0、今周刊法人 19")
rec("FY2027 營收 vs 凱基", "=b_rev27", pl("rev", 2027), "凱基 2026-08-12", "凱基元件及其他 2027F +93%", fmt="#,##0.0")
rec("FY2027 EPS vs 凱基", "=b_eps27a", pl("eps", 2027), "凱基 2026-08-12", "")
rec("FY2027 EPS vs 瑞銀（07-22）", "=b_eps27b", pl("eps", 2027), "經濟日報 2026-07-22", "")
rec("FY2027 EPS vs 法人（08-18）", "=b_eps27c", pl("eps", 2027), "今周刊 2026-08-18", "三家區間 22.0–25.0")
rec("FY2028 EPS vs 法人（08-18）", "=b_eps28", pl("eps", 2028), "今周刊 2026-08-18", "摩根士丹利 05-15：23.03")
rec("FY2027 毛利率 vs 凱基（百分點）", "=b_gm27", pl("gm", 2027), "凱基 2026-08-12", "", fmt="0.00%", pt=True)
RC.cell(rr_, 1, "B. 份額（櫃數口徑 vs 本模型金額口徑）").font = BOLD; rr_ += 1
rec("FY2026 有效份額 vs 三家 ODM 中的 NVL72 櫃數份額", "=rk_sh", lambda s: f"=PL!D{PB[s]['sh0']}", "經濟日報 2026-09-10（美系券商；櫃數）", "口徑不同：本模型為金額、分母為全部 AI 組裝市場（含 ASIC、HGX、其他廠）", fmt="0.0%", pt=True)
rec("FY2026 有效份額 vs NVIDIA AI 機櫃全球份額（法人估）", "=rk_sh2", lambda s: f"=PL!D{PB[s]['sh0']}", "Inside 2026-08（法人；公司說 40% 以上）", "同上", fmt="0.0%", pt=True)
rec("FY2026 隱含鴻海出貨櫃數（千櫃，NVL72 等值）", None, pl("rk", 2026), "—", "櫃數調查單月約 2.6 千櫃 × 12 ≈ 31 千櫃（Interested-party；含運算托盤口徑不明）", fmt="#,##0.0")
rec("FY2026 每 GW 鴻海 AI 伺服器營收（US$B）", None, pl("pgw", 2026), "—", "", fmt="0.00")
rec("FY2026 鴻海 ÷ 廣達 AI 伺服器營收（倍）", None, lambda s: "=Units!C10/(Units!C14*Units!C6)", "廣達模型 Units（唯讀）", "櫃數調查 2,600 ÷ 2,525 ≈ 1.03；鴻海含較多 ASIC 與交換機以外的 AI 產品", fmt="0.00")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 毛利率（毛額基礎）＝此值＋d_ai", fmt="0.00%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "", fmt="0.00%", pt=True)
rec("FY2026 AI 伺服器 ÷ 營收", None, pl("aish", 2026), "—", "", fmt="0.0%")
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E13", "—", "凱基 4Q26F 5.92", fmt="0.00")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文只經法說摘要與新聞轉述，原始簡報待人工複核；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司（摘要）：富果 2Q26 法說摘要 2026-08-12（2Q26 與 1H26 損益、產品組合、AI 機櫃、客供、市占目標）", FUG2),
        ("公司（摘要）：鉅亨 1Q26 法說 2026-05-14（1Q26 營業利益、淨利、EPS、業外、產品組合、營收認列模式改變）", CNY1),
        ("公司（摘要）：富果 2025 年法說 2026-03-16（2025 損益、產品組合、2026 展望、稅率）", FUG4),
        ("公司（摘要）：優分析 2025 年財報（營業利益 2,592.23 億、淨利 1,893.54 億、EPS 13.61、4Q25 業外）", UAN25),
        ("公司（摘要）：優分析 2Q25 財報（AI 伺服器占雲端網路超過五成）", UAN2Q25),
        ("公司（摘要）：數位時代 2026-03-16（AI 伺服器占雲端 2023 約 22%、2024 約三成）", BNEXT),
        ("月營收：TechNews 2026-10-05（9 月 1 兆 1,586.33 億、3Q26 3 兆 269 億、前 9 月 7 兆 6,698 億）", TN09),
        ("月營收：TechNews 2026-09-05（8 月 9,217.66 億）", TN08),
        ("月營收：優分析 2026-10-05（3Q26 3.02 兆）", UAN09),
        ("券商：凱基 2026-08-12（2Q26 實績 vs 預估、2025A–2027F 損益、產品別、業外、少數股權）", KGI),
        ("券商：經濟日報 2026-07-22（瑞銀 2026–27 EPS、毛利率、營益率）", UBS0722),
        ("券商：TechNews 2026-05-15（摩根士丹利 2026–28 EPS；AI 伺服器占 1Q26 雲端網路過半）", MS0515),
        ("券商：今周刊 2026-08-18（法人 2026–28 EPS、目標價）", BT0818),
        ("券商：Inside 2026-08（法人估 NVIDIA AI 機櫃份額約 44%）", INSIDE),
        ("券商：經濟日報 2026-09-10（美系券商 NVL72 櫃數調查）", UDN0910),
        ("券商：經濟日報 2026-08-27（VR200 市占 40% 以上）", UDN0827),
        ("資料商：聚財網 2317 日成交（10 月收盤）", "https://stock.wearn.com/cdata.asp?kind=2317"),
        ("資料商：Yahoo 股市（10-16 收盤、近 12 月 P/E）", "https://tw.stock.yahoo.com/quote/2317.TW"),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/foxconn/pe-ratio/"),
        ("資料商：Goodinfo 歷年經營績效（年底收盤、EPS）", "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=2317&RPT_CAT=M_YEAR"),
        ("市場層：markets/network、markets/chip_bridge、tsmc、quanta 成品（本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 法說（11 月中）；AI 伺服器營收金額與占雲端網路比例；雲端網路非 AI 拆分；客供比例量化；季別業外與稅前利益；客戶占比（年報）；前瞻 P/E 十年歷史（Bloomberg）；券商共識（付費）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("鴻海精密（2317）個股模型 v0.3（AI 半導體 Project：伺服器組裝次產業，樣板＝廣達）　2026-10-28", True),
 ("版本：v0.3（2026-10-28：依個股模型原則補上官方四大產品類別——新增「雲端網路產品」＝AI 伺服器＋雲端網路非 AI 與四類合計核對列，數值不變）；v0.2.x（2026-10-25～27）；v0.1.1（2026-10-25：Val 新增第 0 節「目標價與上檔空間」、In 新增情境權重 tp_wA／tp_wB／tp_wC；其餘數值與 v0.1 相同）。v0.1（初版）：AI 伺服器組裝市場與 ODM 每櫃價格讀廣達模型（本專案權威，規則第 13 條），與廣達同口徑；本模型只決定鴻海份額、客供與其他事業。", False),
 ("模型命題：鴻海是集團公司，AI 伺服器營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 每 GW ODM 整櫃價值 × 鴻海有效份額 ×（1 − 客供比例），2026 年約占營收三成；其餘七成（iPhone、網通與一般伺服器、電腦、元件）以簡化成長率推。AI 伺服器拉高營收與營業利益，但業外損失、稅率上升與少數股權使淨利 ÷ 營業利益從 0.73 降到約 0.66。市場約 10 倍未來 12 月 EPS。", False),
 ("核心問題：AI 伺服器份額（NVIDIA 整櫃最大組裝廠）能不能在客供上升下撐住營業利益成長，以及淨利 ÷ 營業利益能不能穩在 0.66 左右；約 10 倍未來 12 月 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029（三年）。金額：十億新台幣；整櫃價格與市場以美元計、以匯率換算。口徑：歸屬母公司淨利。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 伺服器營收：markets/network 的 AI IT 資本支出 × 情境顆數係數 ÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 GW 架數（Tokenomics）× ODM 每櫃價格（廣達模型）＝組裝市場；× 鴻海有效份額（FY2026 以公司雲端網路占比 × AI 占雲端網路校準，金額口徑）× 份額相對值（Assumed）×（1 − 客供比例，Assumed）。", False),
 ("非 AI：雲端網路非 AI、消費智能、電腦終端、元件及其他各自以 FY2026 錨點 × 成長率（Assumed）。毛利率：AI 伺服器（毛額基礎）比其他營收低 d_ai（Assumed），FY2027 起再加 VR 稀釋 vrd（Assumed），其他毛利率為 FY2026 殘差（Derived）；客供部分不計營收、保留 keep 比例的毛利。營業費用：彈性 phi（Derived）。淨利：營業利益 × 淨利係數。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績＋假設（4Q 季增、毛利率、營業費用率、淨利 ÷ 營業利益）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：現價 ÷ FY2027、FY2028 EPS；十年 P/E 規則值（兩個資料商來源）× FY2027 EPS；循環位置以本模型營收最高年判定（不套設備股見頂年）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文只經法說摘要與新聞；券商多為 3Q 營收公布前。", False),
 ("", False),
 ("已知限制：(1) AI 伺服器營收金額與占雲端網路比例公司未揭露，FY2026 以「1Q26 超過五成」與 2H 機櫃放量假設 60%，是份額校準的最大不確定；(2) 雲端網路非 AI（網通、一般伺服器）未拆，以單一成長率推；(3) 客供比例與保留毛利比例是 Assumed，鴻海已有半成品模式（1Q26 雲端網路營收受影響），FY2026 已含在錨點；(4) 份額為金額口徑，與櫃數份額（三家 ODM 中約 35%、法人估 NVIDIA 機櫃約 44%）不同；(5) 業外、稅與少數股權以一個係數涵蓋，季別業外與稅前未找到；(6) 十年 P/E 2016–2023 為 iPhone 組裝期低倍數，規則值可能低估 AI 期倍數；(7) 3Q26 法說（11 月中）後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

V2["pe10"] = (list(zip(range(2016, 2026), [f"=({a}+{str(b).lstrip('=')})/2" for a, b in zip(PE_CMC, PE_GI)])), "companiesmarketcap 與 Goodinfo 兩來源平均；同規則值")
V2["old_row"] = TP0
# ================= v0.2（2026-10-25）：PL 延伸到 FY2032、Val 第 0 節目標價（DCF 與 Comparable 加權）、Sens 目標價參數 =================
# 全部放在各頁既有內容之後，不位移既有儲存格（FY2025–FY2029 數值與 v0.1.1 相同）
FADEFILL = PatternFill("solid", fgColor="E7E6E6")
XY = (2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032); XC = {y: "CDEFGHIJ"[i] for i, y in enumerate(XY)}; FADE = (2030, 2031, 2032)
K2 = V2["keys"]
for rw in range(4, IN.max_row + 1):
    if IN.cell(rw, 1).value in ("tp_wA", "tp_wB", "tp_wC", "tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "nim10_f"):
        IN.cell(rw, 5).fill = KEY
# --- PL 延伸表（三情境到 FY2032） ---
xr = PL.max_row + 3; XT = xr
PL.cell(xr, 1, "延伸：三情境到 FY2032（v0.2 新增）。FY2025–FY2029 引用上方推導區塊；FY2030 起為收斂法（灰底），非驅動因子推導：營收年增由 FY2029 線性收斂到 fade_g、淨利率由 FY2029 線性收斂到收斂目標，FY2032 達到").font = BOLD
xr += 1; XR = {}; XH = None
for s in SC:
    E = PB[s]; R = {}
    PL.cell(xr, 1, f"延伸｜{SCN_T[s]}").font = BOLD; xr += 1
    hdr(PL, xr, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    XH = XH or xr; xr += 1
    for key, lab, unit in (("rev", "營收（延伸）", V2["unit"]), ("g", "  營收年增", "%"), ("ni", "淨利（延伸）", V2["unit"]), ("nim", "  淨利率", "%"),
                           ("sh", "股數（延伸）", V2["shunit"]), ("eps", "EPS（延伸）", V2["cur"]), ("epsg", "  EPS 年增", "%"),
                           ("tg", "收斂目標：FY2032 營收年增", "%"), ("tm", "收斂目標：FY2032 淨利率", "%")):
        PL.cell(xr, 1, lab); PL.cell(xr, 2, unit); R[key] = xr; xr += 1
    G = lambda k: f"$G${R[k]}"
    for y in XY:
        c = XC[y]; pc = XC.get(y - 1)
        if y <= 2029:
            PL[f"{c}{R['rev']}"] = f"={c}{E[K2['rev']]}"
            if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            PL[f"{c}{R['ni']}"] = f"={c}{E[K2['ni']]}"
            if y > 2025 or K2.get("sh25", True): PL[f"{c}{R['sh']}"] = f"={c}{E[K2['sh']]}"
            PL[f"{c}{R['eps']}"] = f"={c}{E[K2['eps']]}"
        else:
            k = y - 2029
            PL[f"{c}{R['g']}"] = f"={G('g')}+($C${R['tg']}-{G('g')})*{k}/3"
            PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
            PL[f"{c}{R['nim']}"] = f"={G('nim')}+($C${R['tm']}-{G('nim')})*{k}/3"
            PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
            PL[f"{c}{R['sh']}"] = f"={pc}{R['sh']}*(1+{V2['shg']})"
            PL[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh']}"
        if y <= 2029: PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
        if y > 2025: PL[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("rev", "#,##0.0"), ("g", "0.0%"), ("ni", "#,##0.0"), ("nim", "0.0%"), ("sh", "0.0000"), ("eps", V2["epsfmt"]), ("epsg", "0.0%")):
            PL[f"{c}{R[key]}"].number_format = fm
            if y in FADE: PL[f"{c}{R[key]}"].fill = FADEFILL
    PL[f"C{R['tg']}"] = f"=fade_g{s}"; PL[f"C{R['tg']}"].number_format = "0.0%"
    PL[f"C{R['tm']}"] = {"A": f"=G{R['nim']}", "B": f"=G{R['nim']}*(1+nim10_f)/2", "C": f"=G{R['nim']}*nim10_f"}[s]; PL[f"C{R['tm']}"].number_format = "0.0%"
    PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導（上游市場層只到 2029）")
    PL.cell(R["tg"], 11, {"A": "In 頁 fade_gA＝tp_g＋2pt", "B": "In 頁 fade_gB＝tp_g", "C": "In 頁 fade_gC＝tp_g−2pt"}[s])
    PL.cell(R["tm"], 11, {"A": "＝本情境 FY2029 淨利率", "B": "＝FY2029 淨利率與十年平均替代值（FY2029 × nim10_f）的中點", "C": "＝十年平均替代值（FY2029 × nim10_f；十年平均未找到）"}[s])
    PL.cell(R["sh"], 11, f"FY2030 起股數年變動＝{V2['shg']}")
    XR[s] = R; xr += 1
PL.column_dimensions["I"].width = 14; PL.column_dimensions["J"].width = 14; PL.column_dimensions["K"].width = 60
# --- Val 第 0 節：目標價 ---
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
t0 = V.max_row + 3; T0 = t0
V.cell(t0, 1, "0. 目標價（DCF 與 Comparable 加權；v0.2 新增，放在本頁既有內容之後以免位移既有儲存格；三情境欄 B／C／D＝A 順利／B 平平／C 不如預期）").font = BOLD
V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
hdr(V, t0 + 2, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"]); T["h1"] = t0 + 2
r_ = t0 + 3
V.cell(r_, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{r_}"] = "PL 頁延伸表；FY2030 起為收斂法"; T["eps_h"] = r_; r_ += 1
EPSR = {}
for y in YRS6:
    tset(r_, f"eps{y}", f"  FY{y}", [f"=PL!{XC[y]}{XR[s]['eps']}" for s in SC], V2["epsfmt"]); EPSR[y] = r_; r_ += 1
V.cell(r_, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = r_; r_ += 1
tset(r_, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）"); RV = r_; r_ += 1
CFR = {}
for y in YRS6:
    tset(r_, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], V2["epsfmt"]); CFR[y] = r_; r_ += 1
def dcf_formula(c, r="tp_r", g="tp_g", reinv=None):
    """c＝情境欄；r、g 可替換（Sens 用）。基準日 2026 年底，t＝1 對 FY2027。"""
    reinv = reinv or f"Val!$B${RV}"
    cf = lambda y: f"Val!{c}{EPSR[y]}*(1-{reinv})"
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"({pv}+{cf(2032)}*(1+{g})/({r}-{g})/(1+{r})^6)"
c1, c6 = CFR[2027], CFR[2032]
tset(r_, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], V2["pxfmt"],
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027"); r_ += 1
tset(r_, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{T['dcfv']}*(1+tp_r)" for c in "BCD"], V2["pxfmt"]); r_ += 1
tset(r_, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{T['dcf12']}+tp_wB*C{T['dcf12']}+tp_wC*D{T['dcf12']}", None, None], V2["pxfmt"], "Σ tp_wA／wB／wC × 情境價值 ×(1＋tp_r)；權重見 In 頁"); r_ += 2
V.cell(r_, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; r_ += 1
hdr(V, r_, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
P1 = r_
for nm, tk, pe, dt in V2["peers"]:
    V.cell(r_, 1, nm); V.cell(r_, 2, tk); cc = V.cell(r_, 3, pe); cc.font = BLUE; cc.number_format = "0.0"
    V.cell(r_, 4, dt); V.cell(r_, 5, "stockanalysis.com 個股頁 Forward PE（分析師共識；價格日期如左）"); V.cell(r_, 6, "Interested-party（賣方共識）"); r_ += 1
P2 = r_ - 1; T["p1"], T["p2"] = P1, P2
tset(r_, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{P1}:C{P2})", None, None], "0.0", V2["peer_note"]); r_ += 1
w1, w2 = V2["fwd_w"]
tset(r_, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*({w1}*{c}{EPSR[2027]}+{w2}*{c}{EPSR[2028]})" for s, c in zip(SC, "BCD")), None, None], V2["epsfmt"], V2["fwd_note"]); r_ += 1
tset(r_, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{T['feps']}", None, None], V2["pxfmt"], "同業中位數 × 上列 EPS"); r_ += 2
V.cell(r_, 1, "0.3 結論").font = BOLD; r_ += 1
hdr(V, r_, ["項目", "數值", "", "", "說明"]); T["h3"] = r_; r_ += 1
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], V2["pxfmt"], "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁"); r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], V2["pxfmt"], "引用 0.2"); r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", None, None], V2["pxfmt"], V2["cons_tp"][1]); r_ += 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", V2["px_note"]); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS 見 In 頁 cons_feps 與 Recon"); r_ += 1
tset(r_, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%"); r_ += 1
V[f"E{T['gmul']}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
tset(r_, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）"); r_ += 1
PEROW = r_ + 3
PE10 = f"$B${PEROW + 1}:$K${PEROW + 1}"
tset(r_, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
     [f"=IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E 分布（下方列）的分位；低於最小值記 0%、高於最大值記 100%；只作檢查，不進目標價"); r_ += 1
V.cell(r_, 1, "目標價方法"); V[f"B{r_}"] = V2["method"]; T["m"] = r_; r_ += 1
V.cell(PEROW, 1, "十年年底 P/E（歷史分位檢查用；" + V2["pe10"][1] + "）").font = BOLD
for j, (yy, vv) in enumerate(V2["pe10"][0]):
    V.cell(PEROW, 2 + j, str(yy))
    cc = V.cell(PEROW + 1, 2 + j, vv); cc.font = BLUE; cc.number_format = "0.0"
V.cell(PEROW + 1, 1, "年底 P/E"); T["pe10h"] = PEROW; T["pe10"] = PEROW + 1
for k_ in ("dcftp", "comptp", "tpw"):
    V[f"B{T[k_]}"].fill = KEY
V.cell(T["tpw"], 1).font = BOLD
defname("tp_final", f"Val!$B${T['tpw']}"); defname("tp_dcf", f"Val!$B${T['dcftp']}"); defname("tp_comp", f"Val!$B${T['comptp']}")
V["A3"] = f"目標價：見第 0 節（本頁第 {T0} 列；v0.2 新增，DCF 與 Comparable 加權）。第 {V2['old_row']} 列「舊法：規則值情境加權（參考）」不再是目標價，只作參考"
for c in "IJK": V.column_dimensions[c].width = 10
# --- Sens：目標價參數 ---
TPS0 = SE.max_row + 3
SE.cell(TPS0, 1, "目標價參數（v0.2）：一次動一個，對加權目標價的影響（只改 DCF 參數；PL 收斂目標 fade_g 不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def tp_case(i, name, setting, r="tp_r", g="tp_g", w="tp_w_dcf"):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    reinv = f"MIN({g}/tp_roe,0.8)"
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, r, g, reinv)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
    SE[f"F{i}"] = f"=E{i}/E${TPS0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = V2["pxfmt"]
    SE[f"F{i}"].number_format = "0.0%"
TPCASES = [("基準", "tp_r 10%、tp_g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "tp_r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "tp_r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "tp_g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "tp_g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.4", "tp_w_dcf 0.4", dict(w="0.4")), ("DCF 權重 0.8", "tp_w_dcf 0.8", dict(w="0.8"))]
for k_, (nm, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + 2 + k_, nm, st, **kw)
TPS_LAST = TPS0 + 1 + len(TPCASES)
# --- Recon：目標價對帳 ---
rq = RC.max_row + 2
RC.cell(rq, 1, "E. 目標價對帳（v0.2；模型欄為情境加權值，非情境 A）").font = BOLD; rq += 1
RECV2 = rq
for lab, ref, mdl, src in (("12 個月後的前瞻 EPS（情境加權）vs 共識前瞻 EPS", "=cons_feps", f"=Val!B{T['feps']}", V2["cons_feps"][1]),
                           ("目標價（DCF 與 Comparable 加權）vs 市場共識目標價", "=cons_tp", f"=Val!B{T['tpw']}", V2["cons_tp"][1]),
                           ("Comparable 目標價 vs 市場共識目標價", "=cons_tp", f"=Val!B{T['comptp']}", "同上")):
    RC.cell(rq, 1, lab); RC.cell(rq, 2, ref).number_format = V2["epsfmt"]; RC.cell(rq, 3, mdl).number_format = V2["epsfmt"]
    RC.cell(rq, 4, f"=C{rq}/B{rq}-1").number_format = "0.0%"; RC.cell(rq, 9, src); RC.cell(rq, 10, "共識只作對帳，不作校準"); rq += 1
# --- README ---
for rw in range(1, RD.max_row + 1):
    v_ = RD.cell(rw, 1).value
    if isinstance(v_, str) and rw == 1: RD.cell(rw, 1, v_.replace("v0.1.1", "v0.2.1"))
    elif isinstance(v_, str) and v_.startswith("估值："): RD.cell(rw, 1, V2["readme_val"] + "（v0.1.1 原估值說明，保留作參考：" + v_[3:] + "）")
    elif isinstance(v_, str) and v_.startswith("版本："): RD.cell(rw, 1, "版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察（數值不變））。v0.2（2026-10-25：目標價改為 DCF 與 Comparable 加權、三情境延伸到 FY2032；FY2025–FY2029 數值不變）。前版：" + v_[3:])
rn = RD.max_row + 1
for t_ in V2["readme_add"]:
    RD.cell(rn, 1, t_); rn += 1
if os.environ.get("V2_ROWMAP"):
    import json as _json
    _json.dump({"XT": XT, "XH": XH, "XR": XR, "T": T, "EPSR": EPSR, "CFR": CFR, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "RECV2": RECV2},
               open(os.environ["V2_ROWMAP"], "w"), ensure_ascii=False, indent=1)

# ================= v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（工作單；Andy 2026-10-25 14:00、14:17） =================
# 全部放在 Val、README 現有最後一列之後，不位移既有儲存格
OBS = [
       "目標價 NT$349.9（DCF 290.3、Comparable 409.5 各半），比現價 245（2026-10-22）高 42.8%，級距 25%–50%；Comparable 單獨為 67.2%。主因是 EPS 與倍數兩者：本模型 12 個月後前瞻 EPS 27.78 比共識前瞻 EPS 21.06 高 31.9%（兩者期間差約一年），同業（廣達、緯創、和碩、Jabil、Flex）前瞻 P/E 中位數 14.7 倍也高於現價隱含 11.6 倍；DCF 只高 18.5%，因 ROE 10.1% 接近折現率 10%（再投資率 39.7%）。",
       "FY2027 EPS 26.71 比法人 22.0–25.0 高 6.9%–21.4%；淨利 ÷ 營業利益是顆數以外最敏感的變數（0.60 時 FY2027 EPS −9.1%，回到 FY2025 的 0.73 時 +10.6%）。3Q26 財報與法說（約 11 月中）的業外與稅率是第一個檢驗點。",
       "FY2028 EPS A／B／C＝31.42／27.71／23.60，法人 28.0 落在 A 與 B 之間；份額相對值 0.75（廣達、緯創搶單）時 FY2028 EPS −14.3%，份額 1.15（Vera Rubin 市占目標五成）時 +8.5%。追蹤 Vera Rubin 整櫃出貨份額。"
      ]   # 重點觀察（與 README.md 相同文字；數字取自重算後的 xlsx）
_U = f"B{T['upw']}"
GR = V.max_row + 3
V.cell(GR - 1, 1, "0.4 上檔空間級距與投資建議（v0.2.1 新增；放在本頁最後以免位移既有儲存格；依「上檔空間：加權目標價」第 " + str(T["upw"]) + " 列）").font = BOLD
tset(GR, "grade", "上檔空間級距", [f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "General",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）；負值區間同樣重要")
tset(GR + 1, "rec", "投資建議", [f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))', None, None], "General",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）")
rn = RD.max_row + 1
RD.cell(rn, 1, f'="投資建議："&Val!B{GR + 1}&"（依上檔空間固定規則）"').font = BOLD
OBS_ROW = rn + 1
for i_, t_ in enumerate(OBS, 1):
    RD.cell(rn + i_, 1, f"重點觀察 {i_}：" + t_)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "foxconn")
wb.save(sys.argv[1])
if os.environ.get("FOXCONN_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": SENS_N, "VAL": VAL, "SB": {k: v["end"] for k, v in SB.items()}},
              open(os.environ["FOXCONN_ROWMAP"], "w"), ensure_ascii=False, indent=1)
