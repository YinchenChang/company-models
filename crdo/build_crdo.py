# Credo（CRDO）收支與估值模型 v0.2（AI 半導體 Project：網通次產業內複製——主動電纜 AEC、SerDes、光 DSP）
# 用法：python3 crdo/build_crdo.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度：FY(t) 於 t 年 4 月底／5 月初結束（52／53 週；FY2026：2025-05-04～2026-05-02；FY2027：2026-05-03～2027-05-01）。
#   FY(t) 有 8 個月落在日曆年 t−1（5–12 月）、4 個月落在 t（1–4 月）。上游市場層是日曆年，FY(t) 市場＝8/12 × CY(t−1)＋4/12 × CY(t)（月數加權）。
# 金額：十億美元（US$B）；EPS：美元；口徑：公司 non-GAAP（排除股票薪酬、併購攤銷與整合費用）；十年 P/E 規則值用 GAAP 換算 EPS。
# FY2026 為完整實績；FY2027 錨點＝1Q27 實績＋2Q27 指引＋2H27（全年指引「年增逾 85%」減 1H）。2Q27 財報約 2026-12 初公布。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
NET_LABS = {"銅（scale-up 背板與線纜＋scale-out DAC／AEC）": "cu", "scale-out 銅纜（DAC／AEC）": "so",
            "  其中可插拔": "pl", "AI IT 資本支出（加速器 × 每顆 IT capex）": "itc"}
NETR = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    lab = r[0].value
    if lab in NET_LABS:
        NETR[NET_LABS[lab]] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert len(NETR) == 4 and all(v is not None for k in NETR for v in NETR[k][0]), NETR.keys()
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝SEC 8-K 新聞稿與 10-K；公司不揭露產品線營收（AEC、光學、IC、IP）與客戶名稱；客戶名稱只有券商與媒體來源"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR4 = "https://www.sec.gov/Archives/edgar/data/0001807794/000162828026039474/credoq42026ex-991.htm"
PR1 = "https://www.sec.gov/Archives/edgar/data/0001807794/000162828026059795/credoq12027ex-991.htm"
K10 = "https://www.sec.gov/Archives/edgar/data/0001807794/000162828026043303/crdo-20260502.htm"
LDS = "https://letsdatascience.com/news/credo-raises-guidance-as-optical-ramp-accelerates-30d77185"
TIG = "https://www.itiger.com/news/1122940235"
ZAC = "https://www.yourwyominglink.com/credo-q1-earnings-call-highlights-optics-as-the-next-growth-engine/article_d01cfee6-9667-5ab6-8b08-fcf8fffa13d0.html"
JEF = "https://invezz.com/en-ae/news/2026/04/13/credo-jumps-11percent-as-jefferies-backs-copper-tech-in-ai-boom-debate/"
BZ = "https://www.benzinga.com/25/03/44149586/credo-expands-hyperscaler-partnerships-with-amazon-microsoft-xai-amid-nvidia-powered-ai-boom"
YH = "https://finance.yahoo.com/quote/CRDO/analysis/"
SA = "https://stockanalysis.com/stocks/crdo/"
SAF = "https://stockanalysis.com/stocks/crdo/forecast/"
TK = "https://www.tikr.com/blog/credo-stock-fell-22-last-week-as-it-launched-a-new-optical-cable-heres-where-it-could-go-by-2031"
CMC = "https://companiesmarketcap.com/credo-technology/pe-ratio/"

sec("A. AI 銅纜與光模組市場（讀 markets/network 成品；日曆年 CY2026–CY2029；情境 A 基準，B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mcu_{y}", f"AI 銅纜市場 CY20{y}（scale-up 背板與線纜＋scale-out DAC／AEC）", "US$B", None, NETR["cu"][0][j], None, "Derived", NETR["cu"][1],
        "加速器顆數 × 每顆銅內容；scale-up 銅＝互連 × 銅占比、2029 年 scale-up 光學占 20%（網通模型）" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mso_{y}", f"其中 scale-out 銅纜（DAC／AEC）CY20{y}", "US$B", None, NETR["so"][0][j], None, "Derived", NETR["so"][1], "Units、Recon 用（只作拆分參考）" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mpl_{y}", f"AI 可插拔光模組市場 CY20{y}", "US$B", None, NETR["pl"][0][j], None, "Derived", NETR["pl"][1], "光 DSP、矽光 PIC、ZeroFlap 光學的市場基底" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 CY20{y}", "US$B", None, NETR["itc"][0][j], None, "Derived", NETR["itc"][1], "Recon 用：Credo 營收 ÷ AI IT 資本支出" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：CY2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、NVIDIA、Arista、中際旭創模型相同）")
    inp(f"um29_{s}", f"情境 {s}：CY2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("w8", "FY(t) 落在日曆年 t−1 的比重（5–12 月）", "x", None, "=8/12", None, "Derived", K10 + "：FY2026 截至 2026-05-02、FY2025 截至 2025-05-03（52／53 週）", "FY(t) 市場＝w8 × CY(t−1)＋（1 − w8）× CY(t)")

sec("B. AEC 份額（本專案為權威：Credo 在 AI 銅纜的份額，規則第 13 條）與客戶集中")
inp("rel_a_28", "AEC 有效份額 FY2028 ÷ FY2027（相對值；＝AEC 滲透率變化 × Credo 份額變化）", "x", 0.75, 1.00, 1.15, "Assumed", JEF + "：Jefferies（2026-04）AEC 2026 年 860 萬條、2027 年 1,230 萬條；Credo 份額樂觀情境約 70%（Interested-party）", "1.6T 使被動 DAC 距離更短（AEC 滲透率上推）vs Astera Labs、Marvell、Broadcom 與線纜廠競爭（10-K 列名競爭者）")
inp("rel_a_29", "AEC 有效份額 FY2029 ÷ FY2027", "x", 0.60, 0.95, 1.20, "Assumed", "同上；ZeroFlap 主動光纜（ALC，2027 年出貨）可能部分自我替代 AEC", TK + "：ALC 50 公尺、每端約 12 瓦（公司在 OCP 2026-10 發表）")
inp("cr_aec_sh", "Credo 占 AEC 市場（Units 拆分用；不進損益）", "%", 0.60, 0.75, 0.90, "Assumed", JEF + "：樂觀情境 2027 年約 70%", "第一手 AEC 市場份額：未找到")
inp("cust_top", "最大終端客戶占營收（Sens「最大客戶流失」用）", "%", None, 0.33, None, "Interested-party", ZAC + "：1Q27 前四大終端客戶 33%、28%、13%、10%（Zacks 轉述法說）", K10 + "：FY2026 終端客戶 D 33%（FY2025 63%）、B 32%、E 19%（公司，未具名）")
inp("cl_28", "最大客戶流失情境：AEC 有效份額 FY2028 ÷ FY2027（Sens 用）", "x", None, "=rel_a_28*(1-cust_top)", None, "Derived", "AEC 份額相對值 ×（1 − 最大客戶占比）；假設流失的客戶全是 AEC", "客戶名稱：Amazon、xAI、Microsoft（Needham 等券商，Benzinga 2025-03；Interested-party）；公司未具名")
inp("cl_29", "最大客戶流失情境：AEC 有效份額 FY2029 ÷ FY2027（Sens 用）", "x", None, "=rel_a_29*(1-cust_top)", None, "Derived", "同上", BZ + "：Amazon 曾占單季營收 86%（FY2025 Q3）")

sec("C. 光學（光 DSP、矽光 PIC〔DustPhotonics〕、ZeroFlap 光學；讀可插拔光模組市場）")
inp("opt_26", "FY2026 光學營收（未揭露）", "US$B", 0.08, 0.15, 0.25, "Assumed", "公司不揭露產品線營收：未找到", LDS + "：FY2027 光學三條產品線各逾 1 億美元、合計逾 6 億美元（公司法說，Interested-party）")
inp("opt_27", "FY2027 光學營收（公司目標「逾 6 億美元」）", "US$B", 0.50, 0.60, 0.70, "Interested-party", TIG + "：FY2027 光學逾 6 億美元（1Q27 法說重申）", "https://finance.yahoo.com/markets/stocks/articles/credo-makes-875-million-move-013700129.html：2026-04 宣布併購時為「逾 5 億美元」")
inp("rel_o_28", "光學有效份額 FY2028 ÷ FY2027", "x", 0.90, 1.20, 1.60, "Assumed", ZAC + "：兩個次世代大型設計案預計 FY2028 量產；1.6T DSP 於 FY2027 後段開始貢獻（公司）", "DSP 競爭：Marvell、Broadcom 為主要供應商（10-K 列名競爭者）")
inp("rel_o_29", "光學有效份額 FY2029 ÷ FY2027", "x", 0.90, 1.35, 2.00, "Assumed", "同上", "—")

sec("D. 其他營收（線卡 PHY、PCIe 重定時器、SerDes IP 授權與小晶片；網通模型未涵蓋）")
inp("oth_s", "其他 IC 與 IP ÷ 總營收（FY2026）", "%", 0.04, 0.07, 0.10, "Assumed", "公司不揭露產品與 IP 授權拆分：未找到", K10 + "：未滿足履約義務約 3,190 萬美元，預計下一會計年度認列（IP 授權）")
inp("g_oth", "其他 IC 與 IP 年增（FY2027 起）", "%", 0.0, 0.15, 0.30, "Assumed", "無", "—")

sec("E. FY2027 錨點（1Q27 實績＋2Q27 指引＋2H27 依全年指引推導）")
inp("q1_rev", "1Q27 營收（截至 2026-08-01）", "US$B", None, 0.479003, None, "Verified", PR1 + "：479,003 千美元（年增 114.7%）", "https://pulse2.com/credo-reports-q1-fiscal-2027-revenue-of-479-million-up-115-year-over-year/：同值")
inp("q1_gp", "1Q27 non-GAAP 毛利（＝營業利益＋營業費用）", "US$B", None, "=0.230641+0.095154", None, "Verified", PR1 + "：non-GAAP 營業利益 230,641、營業費用 95,154 千美元；毛利率 68.0%", "GAAP 毛利率 64.5%（併購攤銷）")
inp("q1_opx", "1Q27 non-GAAP 營業費用", "US$B", None, 0.095154, None, "Verified", PR1, "GAAP 188.4（含股票薪酬 82.3）")
inp("q1_ninc", "1Q27 其他收益（淨）", "US$B", None, 0.00814, None, "Verified", PR1 + "：8,140 千美元", "現金與短期投資 7.64 億美元（DustPhotonics 現金對價後）")
inp("q1_ni", "1Q27 non-GAAP 淨利", "US$B", None, 0.236262, None, "Verified", PR1 + "：236,262 千美元", TIG + "：同值（淨利率 49.3%）")
inp("q1_sh", "1Q27 non-GAAP 稀釋股數", "十億股", None, 0.197505, None, "Verified", PR1 + "：197,505 千股", "GAAP 稀釋 194,378 千股")
inp("q1_eps", "1Q27 non-GAAP 稀釋 EPS", "US$", None, 1.20, None, "Verified", PR1 + "：1.20（GAAP 0.67）", ZAC + "：同值；共識 1.17")
inp("q2_rev", "2Q27 營收指引（中點；截至 2026-10-31）", "US$B", 0.525, 0.530, 0.535, "Interested-party", PR1 + "：5.25–5.35 億美元", TIG + "：同值")
inp("q2_gm", "2Q27 non-GAAP 毛利率指引（中點）", "%", 0.67, 0.68, 0.69, "Interested-party", PR1 + "：67.0–69.0%", "GAAP 62.9–64.9%")
inp("q2_opx", "2Q27 non-GAAP 營業費用指引（中點）", "US$B", 0.100, 0.1025, 0.105, "Interested-party", PR1 + "：1.00–1.05 億美元", "GAAP 1.99–2.04 億美元")
inp("g27", "FY2027 營收年增（公司「逾 85%」；取下緣）", "%", 0.85, 0.85, 0.95, "Interested-party", TIG + "：1Q27 法說：全年年增逾 85%（4Q26 法說時為逾 80%）", "https://www.trefis.com/articles/618478/does-the-drop-in-credo-stock-change-anything/2026-10-13：同（逾 85%）")
inp("gm_h2", "2H27 non-GAAP 毛利率", "%", 0.66, 0.68, 0.69, "Assumed", TIG + "：FY2027 non-GAAP 毛利率約與 FY2026 持平（68.1%）（公司）", "光學占比上升可能下壓：未找到量化")
inp("g_opx27", "FY2027 non-GAAP 營業費用年增（公司「約 50%」）", "%", 0.45, 0.50, 0.60, "Interested-party", LDS + "：FY2027 non-GAAP 營業費用年增約 50%（4Q26 法說）", TK + "：同（約 50%）")
inp("q_oth", "2Q27 起每季其他收益", "US$B", 0.005, 0.007, 0.009, "Assumed", "1Q27 0.0081；現金自 14.4 億降至 7.64 億美元（DustPhotonics）", "—")
inp("tx27", "FY2027 non-GAAP 有效稅率（2Q27 起）", "%", 0.01, 0.015, 0.04, "Assumed", "1Q27：1 − 236.262 ÷（230.641＋8.140）＝1.05%；FY2026 1.15%（推導）", K10 + "：開曼群島法定稅率 0；FY2026 GAAP 有效稅率 1%")
inp("sh_27", "FY2027 non-GAAP 稀釋股數（全年平均）", "十億股", 0.198, 0.199, 0.201, "Assumed", "1Q27 197.5 百萬股；4Q26→1Q27 季增約 1.6 百萬股（含 DustPhotonics 約 80 萬股）", K10 + "：DustPhotonics 或有對價最多約 280 萬股")

sec("F. FY2025、FY2026 實績（non-GAAP）")
for y, rv, gp, opx, ninc, ni, sh, eps, src2 in (
        (25, 0.436775, "=0.115352+0.168751", 0.168751, 0.017746, 0.129926, 0.184644, 0.70, "10-K FY2026（三年損益表）：營收 436,775 千美元"),
        (26, 1.335116, "=0.638822+0.270945", 0.270945, 0.030430, 0.661543, 0.191256, 3.46, K10 + "：營收 1,335,116 千美元、GAAP 淨利 472,279")):
    inp(f"a{y}_rev_in", f"FY20{y} 營收", "US$B", None, rv, None, "Verified", PR4, src2)
    inp(f"a{y}_gp_in", f"FY20{y} non-GAAP 毛利（＝營業利益＋營業費用）", "US$B", None, gp, None, "Verified", PR4 + ("：non-GAAP 毛利率 65.0%" if y == 25 else "：non-GAAP 毛利率 68.1%"), "—")
    inp(f"a{y}_opx_in", f"FY20{y} non-GAAP 營業費用", "US$B", None, opx, None, "Verified", PR4, "—")
    inp(f"a{y}_ninc_in", f"FY20{y} 其他收益（淨）", "US$B", None, ninc, None, "Verified", PR4, "—")
    inp(f"a{y}_ni_in", f"FY20{y} non-GAAP 淨利", "US$B", None, ni, None, "Verified", PR4, LDS + ("：FY2026 non-GAAP 淨利 6.62 億美元" if y == 26 else "—"))
    inp(f"a{y}_sh_in", f"FY20{y} non-GAAP 稀釋股數", "十億股", None, sh, None, "Verified", PR4, "—")
    inp(f"a{y}_eps_in", f"FY20{y} non-GAAP 稀釋 EPS", "US$", None, eps, None, "Verified", PR4 + ("：0.70（GAAP 0.29）" if y == 25 else "：3.46（GAAP 2.51）"), "—")

sec("G. 營業費用、稅率、股數（FY2028 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.30, "=Q!B28", 0.90, "Derived", "FY2027 錨點營業費用年增 ÷ 營收年增（Q 頁）", "FY2026：60.6% ÷ 205.7% ＝ 0.29（推導；費用落後營收）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.02, "Assumed", "無", "Sens 用 +2 個百分點")
inp("gm_28", "non-GAAP 毛利率 FY2028", "%", 0.63, 0.67, 0.69, "Assumed", "FY2025 65.0%、FY2026 68.1%、1Q27 68.0%（公司）", "光學（PIC、ZeroFlap 光學）毛利率可能低於 AEC：未找到量化")
inp("gm_29", "non-GAAP 毛利率 FY2029", "%", 0.60, 0.66, 0.69, "Assumed", "同上；AEC 競爭者增加時價格下壓", "—")
inp("oth_f", "其他收益（每年，FY2028 起）", "US$B", 0.02, 0.03, 0.06, "Assumed", "現金與短期投資 7.64 億美元 × 約 4%", "—")
inp("tx_28", "non-GAAP 有效稅率 FY2028", "%", 0.05, 0.13, 0.17, "Assumed", K10 + "：預期 FY2028 起適用 OECD 15% 全球最低稅（公司）", "FY2026、1Q27 約 1%；共識是否已含：未找到")
inp("tx_29", "non-GAAP 有效稅率 FY2029", "%", 0.05, 0.14, 0.17, "Assumed", "同上", "—")
inp("sh_g", "股數年變動（FY2028 起）", "%", 0.005, 0.015, 0.03, "Assumed", "FY2025→FY2026 non-GAAP 稀釋股數 +3.6%（含 Amazon 認股權證）；FY2026 股票薪酬 1.83 億美元", "無買回計畫：未找到")
inp("mf", "AI 銅纜市場係數（Sens 用；FY2028 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 CY2027 銅 +97%、CY2028 +16%、CY2029 +2.5%", "±20%（銅轉光提前、或 scale-up 銅擴大）")
inp("g2n", "GAAP EPS ÷ non-GAAP EPS（規則值換算用）", "x", 0.56, 0.70, 0.75, "Assumed", PR4 + "：FY2026 GAAP 2.51 ÷ non-GAAP 3.46 ＝ 0.725；1Q27 0.67 ÷ 1.20 ＝ 0.56（含 DustPhotonics 攤銷與整合費用）", TK + "：FY2027 共識 GAAP 約 4.40 ÷ 調整後約 6.30 ＝ 0.70（Interested-party）")

sec("H. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("cons_q2", "共識 2Q27 營收", "US$B", None, 0.53165, None, "Interested-party", YH + "（17 家，2026-10-01 收盤時）", "—")
inp("cons_rev27", "共識 FY2027 營收", "US$B", None, 2.50, None, "Interested-party", YH + "（19 家；2.47–2.58）", SAF + "：2.5（19 家）")
inp("cons_eps27", "共識 FY2027 non-GAAP EPS", "US$", None, 6.31, None, "Interested-party", YH + "（20 家；5.98–6.64）", SAF + "：6.31（19 家）")
inp("cons_rev28", "共識 FY2028 營收", "US$B", None, 3.87, None, "Interested-party", YH + "（19 家；3.59–4.41）", "第二來源：未找到（StockAnalysis 需付費）")
inp("cons_eps28", "共識 FY2028 non-GAAP EPS", "US$", None, 9.69, None, "Interested-party", YH + "（20 家；8.69–11.31）", "第二來源：未找到")
inp("jef_aec", "Jefferies 引述 AEC 營收共識（期間未明示）", "US$B", None, 2.3, None, "Interested-party", JEF + "：共識 23 億美元、Jefferies 樂觀情境約 30 億美元（2026-04）", "期間：未找到（推測為 CY2027 或 FY2028）")
inp("jef_u26", "Jefferies：AEC 市場 2026 年（條）", "百萬條", None, 8.6, None, "Interested-party", JEF, "樂觀 9.5")
inp("jef_u27", "Jefferies：AEC 市場 2027 年（條）", "百萬條", None, 12.3, None, "Interested-party", JEF, "樂觀 14.6")
inp("tp_street", "賣方平均目標價", "US$", None, 279.09, None, "Interested-party", SAF + "（20 家，2026-10-05；185–315）", TK + "：約 279")
inp("tp_wA", "目標價情境權重 A", "%", None, 0.3, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
inp("tp_wB", "目標價情境權重 B", "%", None, 0.5, None, "Assumed", "Andy 可改", "—")
inp("tp_wC", "目標價情境權重 C", "%", None, 0.2, None, "Assumed", "Andy 可改", "—")

sec("I. 價格與估值")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 147.29, None, "Interested-party", SA + "：147.29（前收 162.77；盤後 148.75）", CMC + "：147.29（市值 276.8 億美元）")
inp("px_hi", "52 週高點", "US$", None, 308.67, None, "Interested-party", SA + "：52 週區間 86.49–308.67", TK + "：2026-06-22 收盤 302.52")
inp("sh_out", "流通股數（2026-10）", "十億股", None, 0.18795, None, "Interested-party", SA + "：187.95 百萬股", K10 + "：186,477,898 股（2026-06-08，公司）")

import re
NAMES = {PR4: "公司 4Q26 新聞稿", PR1: "公司 1Q27 新聞稿", K10: "公司 FY2026 10-K", LDS: "Let's Data Science", TIG: "老虎證券新聞", ZAC: "Zacks 轉載",
         JEF: "Invezz 轉述 Jefferies", BZ: "Benzinga", YH: "Yahoo Finance", SA: "StockAnalysis", SAF: "StockAnalysis 預估", TK: "TIKR", CMC: "companiesmarketcap"}
def strip_url(t):   # 含「未找到」的列會被 HTML 掃描列出：網址改寫為來源名稱（網址見 Sources）
    if not isinstance(t, str): return t
    for u, nm in NAMES.items(): t = t.replace(u, nm + "（網址見 Sources）")
    return re.sub(r"https?://[^\s：（）]+", "（網址見 Sources）", t)
sec("K. 目標價（DCF＋Comparable）與收斂法參數（v0.2 新增；黃底，Assumed，Andy 可改）")
inp("tp_r", "折現率／要求報酬（DCF）", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%；各市場同一值）", "DCF 與 Sens")
inp("tp_g", "長期成長率（DCF 永續；收斂法情境 B 的 FY2032 營收年增）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "—")
inp("tp_roe", "ROE（與 EPS 同口徑：non-GAAP 淨利 ÷ 平均股東權益，近三年平均；DCF 再投資率＝tp_g ÷ tp_roe）", "%", None, 0.2, None, "Assumed", "本模型沒有股東權益（資產負債表未建），non-GAAP ROE 算不出；依主串規則（2026-10-25）取 0.20", '參考（不採）：GAAP ROE stockanalysis FY2024 −6.39%、FY2025 8.54%、FY2026 34.41%（與 non-GAAP EPS 口徑不一致）')
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "收斂法：情境 A 於 FY2032 的營收年增（＝tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法；FY2030 起由 FY2029 年增線性收斂", "FY2030 起為收斂法，非驅動因子推導")
inp("fade_gB", "收斂法：情境 B 於 FY2032 的營收年增（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "收斂法：情境 C 於 FY2032 的營收年增（＝tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "收斂法：情境 A 的中段淨利率（＝情境 A FY2029 淨利率）", "%", None, None, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "收斂法：情境 B 的中段淨利率（＝FY2029 與十年平均的中點；十年平均以 FY2029 × 0.8 代替）", "%", None, None, None, "Assumed", "十年平均淨利率（2016–2025，與本模型同口徑）：未找到，依工作單以 FY2029 × 0.8 代替；本列＝（FY2029＋FY2029 × 0.8）÷ 2", "Andy 可改")
inp("fade_mC", "收斂法：情境 C 的中段淨利率（＝十年平均；以情境 C FY2029 × 0.8 代替）", "%", None, None, None, "Assumed", "十年平均淨利率：未找到，依工作單以情境 C FY2029 淨利率 × 0.8 代替", "Andy 可改")
inp("cons_tp", "市場共識目標價（賣方平均）", 'US$', None, '=tp_street', None, "Interested-party", '賣方平均目標價（In 頁 tp_street：stockanalysis 20 家 279.09，2026-10-05）', 'stockanalysis 統計頁 2026-10-23：20 位平均 284.05')
inp("cons_feps", '共識 12 個月後前瞻 EPS（≈ FY2028；FY2029 共識未找到）', 'US$', None, '=cons_eps28', None, "Derived", 'Yahoo FY2028 non-GAAP 共識 9.69（In 頁 cons_eps28）；本模型前瞻 EPS 含 6／12 的 FY2029', "共識只放 Recon 與本列，不作校準")

r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    if any("未找到" in str(x) for x in (item, s1, s2)): s1, s2 = strip_url(s1), strip_url(s2)
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
for col, w in zip("ABCDEFGHI", [12, 64, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2027 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2027 錨點（十億美元；EPS 為美元；股數為十億股；non-GAAP）。藍字＝輸入；黑字＝公式。FY2027＝2026-05-03～2027-05-01"
Q["A2"] = "FY2027 錨點＝1Q27 實績＋2Q27 指引中點＋2H27（全年營收＝FY2026 ×（1＋85%）減 1H；毛利率 gm_h2；營業費用＝FY2026 ×（1＋50%）減 1H）。產品線拆分（光學、其他 IC 與 IP）為假設，AEC 為殘差。2Q27 財報（約 2026-12 初）公布後以實績取代 2Q"
hdr(Q, 3, ["項目", "FY2025A", "FY2026A", "1Q27（實績）", "2Q27（指引）", "2H27（推導）", "FY2027 錨點", "標記", "說明"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", note=""):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, note)
qrow(4, "營收", {"B": "=a25_rev_in", "C": "=a26_rev_in", "D": "=q1_rev", "E": "=q2_rev", "F": "=G4-D4-E4", "G": "=C4*(1+g27)"}, tag="Verified／Interested-party（指引）", note="2H27＝全年指引 − 1Q27 − 2Q27")
qrow(5, "毛利（non-GAAP）", {"B": "=a25_gp_in", "C": "=a26_gp_in", "D": "=q1_gp", "E": "=E4*q2_gm", "F": "=F4*gm_h2", "G": "=D5+E5+F5"}, tag="Verified／Derived")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(7, "營業費用（non-GAAP）", {"B": "=a25_opx_in", "C": "=a26_opx_in", "D": "=q1_opx", "E": "=q2_opx", "F": "=G7-D7-E7", "G": "=C7*(1+g_opx27)"}, tag="Verified／Interested-party（指引）", note="2H27＝全年（年增約 50%）− 1H")
qrow(8, "營業費用率", {c: f"={c}7/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(9, "營業利益（non-GAAP）", {c: f"={c}5-{c}7" for c in "BCDEFG"}, tag="Derived")
qrow(10, "其他收益（淨）", {"B": "=a25_ninc_in", "C": "=a26_ninc_in", "D": "=q1_ninc", "E": "=q_oth", "F": "=2*q_oth", "G": "=D10+E10+F10"}, tag="Verified／Assumed")
qrow(11, "稅前利益", {c: f"={c}9+{c}10" for c in "BCDEFG"}, tag="Derived")
qrow(12, "所得稅（non-GAAP）", {"B": "=B11-a25_ni_in", "C": "=C11-a26_ni_in", "D": "=D11-q1_ni", "E": "=E11*tx27", "F": "=F11*tx27", "G": "=D12+E12+F12"}, tag="Derived", note="實績＝稅前 − 淨利（推導）")
qrow(13, "有效稅率", {c: f"={c}12/{c}11" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(14, "淨利（non-GAAP）", {"B": "=a25_ni_in", "C": "=a26_ni_in", "D": "=q1_ni", "E": "=E11-E12", "F": "=F11-F12", "G": "=D14+E14+F14"}, tag="Verified／Derived")
qrow(15, "淨利率", {c: f"={c}14/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived", note="公司 FY2027 non-GAAP 淨利率目標約 50%")
qrow(16, "稀釋股數（non-GAAP）", {"B": "=a25_sh_in", "C": "=a26_sh_in", "D": "=q1_sh", "G": "=sh_27"}, tag="Verified／Assumed")
qrow(17, "EPS（non-GAAP）", {"B": "=B14/B16", "C": "=C14/C16", "D": "=D14/D16", "G": "=G14/G16"}, fmt="0.00", tag="Derived", note="公司揭露：FY2025 0.70、FY2026 3.46、1Q27 1.20")
qrow(18, "光學營收（光 DSP、矽光 PIC、ZeroFlap 光學）", {"C": "=opt_26", "G": "=opt_27"}, tag="Assumed／Interested-party")
qrow(19, "其他 IC 與 IP（＝FY2026 營收 × oth_s × 成長）", {"C": "=C4*oth_s", "G": "=C19*(1+g_oth)"}, tag="Assumed")
qrow(20, "AEC（＝營收 − 光學 − 其他；殘差）", {"C": "=C4-C18-C19", "G": "=G4-G18-G19"}, tag="Derived")
qrow(21, "AEC ÷ 營收", {"C": "=C20/C4", "G": "=G20/G4"}, fmt="0.0%", tag="Derived", note="公司：AEC 仍是最大業務；1Q27 營收年增額逾九成來自 AEC")
Q.cell(23, 1, "FY2027 錨點的檢查（參考；校準未使用共識）").font = BOLD
Q.cell(24, 1, "FY2027 營收 ÷ 共識 − 1"); Q["B24"] = "=G4/cons_rev27-1"
Q.cell(25, 1, "FY2027 EPS ÷ 共識 − 1"); Q["B25"] = "=G17/cons_eps27-1"
Q.cell(26, 1, "FY2027 營收年增"); Q["B26"] = "=G4/C4-1"
Q.cell(27, 1, "FY2027 營業費用年增"); Q["B27"] = "=G7/C7-1"
Q.cell(28, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B28"] = "=B27/B26"
Q.cell(29, 1, "1Q27 EPS（模型）vs 公司揭露 1.20"); Q["B29"] = "=D17/q1_eps-1"
Q.cell(30, 1, "FY2027 AEC 年增"); Q["B30"] = "=G20/C20-1"
Q.cell(31, 1, "2H27 每季平均營收 ÷ 2Q27 − 1"); Q["B31"] = "=F4/2/E4-1"
for rr in range(24, 32): Q[f"B{rr}"].number_format = "0.0%" if rr != 28 else "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("ninc", 10), ("tx", 13), ("ni", 14), ("sh", 16), ("eps", 17)):
    for y, col in ((25, "B"), (26, "C"), (27, "G")):
        defname(f"a{y}_{nm}", f"Q!${col}${rr}")
for nm, rr in (("opt", 18), ("ic", 19), ("aec", 20)):
    for y, col in ((26, "C"), (27, "G")):
        defname(f"a{y}_{nm}", f"Q!${col}${rr}")
Q.column_dimensions["A"].width = 52
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 80
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2027, 2028, 2029); F_ = (2028, 2029)
umx = lambda cy, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[cy]
def fymkt(base, y, s, mf):
    """FY(y) 市場＝w8 × CY(y−1)＋（1 − w8）× CY(y)；FY2028 起乘市場係數 mf"""
    a, b = y - 1, y
    f = f"(w8*{base}_{a % 100}*{umx(a, s)}+(1-w8)*{base}_{b % 100}*{umx(b, s)})"
    return f"={f}" + ("" if y == 2027 or mf == "1" else f"*{mf}")
DEFP = dict(scn="A", rela="rel_a", relo="rel_o", gm="gm", tx="tx", mf="1", gn="g_oth", opxa="0", phi="phi")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026A", "FY2027（錨點）", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.000", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    sec_("A. AEC（E 欄＝FY2027 校準年；F–G＝FY2028–FY2029；市場為日曆年月數加權）")
    row("mcu", "AI 銅纜市場（FY 加權）× 情境顆數係數", "US$B", {c_(y): fymkt("mcu", y, s, P["mf"]) for y in BU}, fmt="#,##0.00",
        note="FY(t)＝8/12 × CY(t−1)＋4/12 × CY(t)；CY2028 起乘顆數係數")
    row("sha0", "Credo AEC 有效份額（FY2027 校準）＝ AEC 營收 ÷ 銅纜市場", "%", {"E": f"=a27_aec/E{R['mcu']}"}, fmt="0.00%",
        note="Derived：＝AEC 滲透率 × Credo 占 AEC（拆分見 Units）")
    row("rela", "AEC 有效份額相對值（÷ FY2027）", "x", {"E": "=1", **{c_(y): f"={P['rela']}_{y % 100}" for y in F_}}, fmt="0.00")
    row("sha", "AEC 有效份額", "%", {c_(y): f"=$E${R['sha0']}*{c_(y)}{R['rela']}" for y in BU}, fmt="0.00%")
    row("aec", "AEC 營收", "US$B", {"D": "=a26_aec", "E": "=a27_aec", **{c_(y): f"={c_(y)}{R['mcu']}*{c_(y)}{R['sha']}" for y in F_}})
    sec_("B. 光學（光 DSP、矽光 PIC、ZeroFlap 光學）")
    row("mpl", "AI 可插拔光模組市場（FY 加權）× 情境顆數係數", "US$B", {c_(y): fymkt("mpl", y, s, "1") for y in BU}, fmt="#,##0.00")
    row("sho0", "Credo 光學有效份額（FY2027 校準）＝ 光學營收 ÷ 可插拔市場", "%", {"E": f"=a27_opt/E{R['mpl']}"}, fmt="0.00%")
    row("relo", "光學有效份額相對值（÷ FY2027）", "x", {"E": "=1", **{c_(y): f"={P['relo']}_{y % 100}" for y in F_}}, fmt="0.00")
    row("opt", "光學營收", "US$B", {"D": "=a26_opt", "E": "=a27_opt", **{c_(y): f"={c_(y)}{R['mpl']}*$E${R['sho0']}*{c_(y)}{R['relo']}" for y in F_}})
    sec_("C. 其他營收與合計")
    row("ic", "其他 IC 與 IP（線卡 PHY、重定時器、SerDes 授權）", "US$B", {"D": "=a26_ic", "E": "=a27_ic", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']})" for y in F_}})
    row("rev", "營收", "US$B", {"C": "=a25_rev", "D": "=a26_rev", **{c_(y): f"={c_(y)}{R['aec']}+{c_(y)}{R['opt']}+{c_(y)}{R['ic']}" for y in BU}})
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aecs", "AEC ÷ 營收", "%", {c_(y): f"={c_(y)}{R['aec']}/{c_(y)}{R['rev']}" for y in (2026, 2027, 2028, 2029)}, fmt="0.0%")
    row("opts", "光學 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['opt']}/{c_(y)}{R['rev']}" for y in (2026, 2027, 2028, 2029)}, fmt="0.0%")
    sec_("D. 損益（non-GAAP）")
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", "D": "=a26_gp/a26_rev", "E": "=a27_gp/a27_rev", **{c_(y): f"={P['gm']}_{y % 100}" for y in F_}}, fmt="0.0%",
        note="FY2028 起為輸入路徑（光學占比上升、AEC 競爭）")
    row("gp", "毛利", "US$B", {c_(y): f"={c_(y)}{R['rev']}*{c_(y)}{R['gm']}" for y in FYS})
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", "E": "=a27_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in F_}})
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", "E": f"=E{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in F_}})
    row("oi", "營業利益", "US$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("ninc", "其他收益（淨）", "US$B", {"C": "=a25_ninc", "D": "=a26_ninc", "E": "=a27_ninc", **{c_(y): "=oth_f" for y in F_}})
    row("tr", "有效稅率", "%", {"C": "=a25_tx", "D": "=a26_tx", "E": "=a27_tx", **{c_(y): f"={P['tx']}_{y % 100}" for y in F_}}, fmt="0.0%",
        note="FY2028 起 OECD 15% 全球最低稅（10-K）")
    row("ni", "淨利", "US$B", {"C": "=a25_ni", "D": "=a26_ni", "E": "=a27_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['ninc']})*(1-{c_(y)}{R['tr']})" for y in F_}})
    row("nim", "淨利率", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("sh_", "稀釋股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=a27_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in F_}}, fmt="0.0000")
    row("eps", "EPS（non-GAAP）", "US$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS}, fmt="0.00")
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Credo 損益（US$B；non-GAAP）。FY2025A、FY2026A 實績、FY2027 錨點（1Q27 實績＋2Q27 指引＋2H27 依全年指引），FY2028 起為模型。FY(t) 截至 t 年 4 月底／5 月初。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2027 三情境相同；FY2028 只有 4/12 落在 CY2028，B、C 的差異主要在 FY2029"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A", "FY2027（錨點）", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.00"), ("g", "  營收年增", "%", "0.0%"), ("aec", "  AEC 營收", "US$B", "#,##0.00"), ("opt", "  光學營收", "US$B", "#,##0.00"),
                                ("ic", "  其他 IC 與 IP", "US$B", "#,##0.00"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"), ("nim", "淨利率", "%", "0.0%"),
                                ("ni", "淨利", "US$B", "#,##0.00"), ("eps", "EPS（non-GAAP）", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("aec", "opt", "ic") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2027 模型營收、AEC 營收、淨利 − 錨點（校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（US$B）"); PL[f"E{sr}"] = f"=E{PB['A']['rev']}-a27_rev"; PL[f"E{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：AEC 營收 − 銅纜市場 × 有效份額（US$B）"); PL[f"E{sr}"] = f"=E{PB['A']['aec']}-E{PB['A']['mcu']}*E{PB['A']['sha']}"; PL[f"E{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：（營業利益＋其他收益）×（1 − 稅率）− 淨利錨點"); PL[f"E{sr}"] = f"=(E{PB['A']['oi']}+E{PB['A']['ninc']})*(1-E{PB['A']['tr']})-a27_ni"; PL[f"E{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：市場與份額（權威） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — Credo 在 AI 銅纜（AEC）與光學的份額（本專案權威，規則第 13 條）。市場取自 markets/network（加速器顆數 × 每顆內容；日曆年 → 會計年度月數加權）"
A_ = PB["A"]
uc = {2027: "C", 2028: "D", 2029: "E"}
hdr(UN, 3, ["項目", "單位", "FY2027（錨點）", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("mcu", "AI 銅纜市場（情境 A；FY 加權）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mcu']}", "#,##0.0", "markets/network Mkt 頁「銅」"),
        ("mso", "  其中 scale-out 銅纜（DAC／AEC）", "US$B", lambda y: f"=w8*mso_{(y - 1) % 100}+(1-w8)*mso_{y % 100}", "#,##0.0", "其餘為 scale-up 背板與線纜"),
        ("aec", "Credo AEC 營收（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['aec']}", "#,##0.00", "FY2027 為錨點殘差"),
        ("sha", "Credo AEC 有效份額（÷ 銅纜市場）", "%", lambda y: f"=PL!{FC[y]}{A_['sha']}", "0.00%", "FY2027 校準"),
        ("shs", "Credo AEC ÷ scale-out 銅纜", "%", lambda y: f"=PL!{FC[y]}{A_['aec']}/{uc[y]}5", "0.0%", "若 AEC 全在 scale-out"),
        ("pen", "隱含 AEC 滲透率（＝有效份額 ÷ Credo 占 AEC）", "%", lambda y: f"=PL!{FC[y]}{A_['sha']}/cr_aec_sh", "0.0%", "Credo 占 AEC 75%（Assumed）"),
        ("mpl", "AI 可插拔光模組市場（情境 A；FY 加權）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mpl']}", "#,##0.0", "markets/network Mkt 頁「其中可插拔」"),
        ("opt", "Credo 光學營收（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['opt']}", "#,##0.00", ""),
        ("sho", "Credo 光學有效份額（÷ 可插拔市場）", "%", lambda y: f"=PL!{FC[y]}{A_['opt']}/PL!{FC[y]}{A_['mpl']}", "0.00%", "光 DSP 約占模組成本兩成（次級來源），換算 DSP 份額約為本列 5 倍"),
        ("itr", "Credo 營收 ÷ AI IT 資本支出（FY 加權）", "%", lambda y: f"=PL!{FC[y]}{A_['rev']}/(w8*itc_{(y - 1) % 100}+(1-w8)*itc_{y % 100})", "0.00%", "markets/network 的 AI IT 資本支出")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{uc[y]}{ur}"] = f(y); UN[f"{uc[y]}{ur}"].number_format = fmt
    UN.cell(ur, 6, note); UR[key] = ur; ur += 1
assert UR["mso"] == 5, UR
ur += 1
UN.cell(ur, 1, "讀法：有效份額＝AEC 滲透率 × Credo 占 AEC；兩者都沒有第一手資料，只校準乘積（FY2027 錨點）。AEC 份額相對值、光學份額相對值是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 10
for c in "CDE": UN.column_dimensions[c].width = 14
UN.column_dimensions["F"].width = 70

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。目標價＝DCF 與 Comparable 加權（第 0 節）；舊法規則值情境加權保留參考；另列「現價 ÷ FY2027、FY2028 EPS（模型，non-GAAP）」、市場隱含 EPS 與十年 P/E 規則值。Credo 2022-01 上市，年底 P/E 只有 2023–2025，且 2023、2024 為虧損（負值，不用）：規則值只有 2025 年一個觀測值"
V["A2"] = "循環位置：Credo 不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 Arista、中際旭創）。營收到 FY2029 仍在增加時，見頂年未知，判為中段。規則值基準＝2027 年底的近 12 月 GAAP 換算 EPS（≈ ½ FY2027＋½ FY2028，× g2n），與年底 P/E 的計算口徑一致"
VAL = {}
# 1. 價格與市場隱含倍數（第 14 列起；第 0 節在最後寫入第 4–12 列）
V.cell(14, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 15, ["項目", "A", "B", "C", "說明"])
V.cell(16, 1, "現價（2026-10-23 收盤）"); V["B16"] = "=px"; V["B16"].number_format = "0.00"
V.cell(17, 1, "現價 ÷ FY2026 EPS（實績）"); V.cell(18, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(19, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(20, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((17, 2026), (18, 2027), (19, 2028), (20, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(21, 1, "現價 ÷ 共識 FY2027 EPS"); V["B21"] = "=px/cons_eps27"; V["B21"].number_format = "0.0"
V.cell(22, 1, "現價 ÷ 共識 FY2028 EPS"); V["B22"] = "=px/cons_eps28"; V["B22"].number_format = "0.0"
V.cell(23, 1, "現價 ÷ 52 週高點 − 1"); V["B23"] = "=px/px_hi-1"; V["B23"].number_format = "0%"
V.cell(24, 1, "現價 ÷ 賣方平均目標價 − 1"); V["B24"] = "=px/tp_street-1"; V["B24"].number_format = "0%"
V["E18"] = "FY2027 三情境相同（錨點）"; V["E19"] = "FY2028 只有 4/12 落在 CY2028；B、C 的差異主要在 FY2029"
V.cell(26, 1, "2. 循環位置（本模型營收最高年）").font = BOLD
hdr(V, 27, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 28):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(32, 1, "3. 規則值隱含股價（美元；約 2027 年底，未折現）＝ 2027 年底近 12 月 GAAP 換算 EPS × 規則倍數").font = BOLD
hdr(V, 33, ["情境", "FY2027 EPS", "FY2028 EPS", "2027 年底近 12 月 GAAP 換算 EPS", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
for j, s in enumerate(SC):
    i = 34 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=PL!F{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 4, f"=(B{i}+C{i})/2*g2n").number_format = "0.00"
    V.cell(i, 5, f"=D{28 + j}").number_format = "0.0"
    V.cell(i, 6, f"=D{i}*E{i}").number_format = "#,##0"
    V.cell(i, 7, f"=F{i}/px-1").number_format = "0%"
    V.cell(i, 8, f"=px/B{i}").number_format = "0.0"
    for c in range(1, 9): V.cell(i, c).fill = KEY
    VAL[s] = i
V.cell(37, 1, "近 12 月＝2026-11～2027-10（3Q27–2Q28）≈ ½ FY2027＋½ FY2028；年底 P/E 資料商以近 12 月 GAAP EPS 計算，故以 g2n 換算")
V.cell(39, 1, "4. 核心問題：市場隱含的 FY2028 EPS（＝現價 ÷ 假設倍數；non-GAAP）與模型的距離").font = BOLD
hdr(V, 40, ["假設倍數（÷ FY2028 EPS）", "市場隱含 FY2028 EPS", "模型 FY2028 EPS（A）", "隱含 ÷ 模型 A − 1", "模型 FY2028 EPS（C）", "隱含 ÷ 模型 C − 1"])
vr = 41; IMPL = {}
for m in (10, 15, 20, 30):
    V.cell(vr, 1, m).number_format = "0.0"; V.cell(vr, 1).font = BLUE
    V.cell(vr, 2, f"=px/A{vr}").number_format = "0.00"
    V.cell(vr, 3, f"=PL!F{SUMR[('A', 'eps')]}").number_format = "0.00"
    V.cell(vr, 4, f"=B{vr}/C{vr}-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[('C', 'eps')]}").number_format = "0.00"
    V.cell(vr, 6, f"=B{vr}/E{vr}-1").number_format = "0%"
    IMPL[m] = vr; vr += 1
V.cell(vr, 1, "讀法：以 15 倍計（約等於現價 ÷ 共識 FY2028 EPS），現價隱含的 FY2028 EPS 與模型 A 相近；以 20 倍計約等於情境 C；以 30 倍（高成長股常見）計約為模型 A 的一半，需要 AEC 份額流失、最大客戶流失、銅纜市場 ×0.8 其中兩到三項同時發生（見 Sens）。"); READ_ROW = vr
vr += 1
V.cell(vr, 1, "本版未做先驗，無 Andy 合理 P/E 對照。"); NOPRIOR_ROW = vr
assert vr < 50, vr
PE_CMC = [None, None, None, None, None, None, None, None, None, 118.0]   # companiesmarketcap 年底 P/E；2023 −92.7、2024 −456（虧損），2022 年以前未上市或無資料
PE_SA = [None, None, None, None, None, None, None, None, None, 157.04]   # StockAnalysis 會計年度末 P/E：FY2025（2025-05-03）；FY2026（2026-05-02）72.01 不在 2016–2025 範圍
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；只有 2025 年一個正值；2023 −92.7x、2024 −456x 為虧損，不用）"},
    {"label": "StockAnalysis（會計年度末 P/E；參考）", "vals": PE_SA, "tag": "Interested-party（資料商；2025 列＝FY2025 年底 2025-05-03；FY2026 年底 2026-05-02 72.0x；日期口徑不同，只作參考、不採用）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "2027 年底近 12 月 GAAP 換算 EPS（A）", "basis": f"=D{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "2027 年底近 12 月 GAAP 換算 EPS（A）", "basis": f"=D{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "2027 年底近 12 月 GAAP 換算 EPS（A）", "basis": f"=D{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝年底 P/E 依循環位置取分位。Credo 只有 2025 年一個正值（118x，營收年增約兩倍的高成長期），三個分位相同；歷史倍數為 GAAP 落後 P/E，規則值偏高的方向明確。第二來源：未找到（StockAnalysis 只有會計年度末，列為參考）；前瞻 P/E 歷史待 Bloomberg。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
# 舊法：規則值情境加權（v0.1 稱「0. 目標價與上檔空間」；v0.2 改名保留，只作參考；建置後與其他內容一起下移）
V.cell(4, 1, "舊法：規則值情境加權（參考）").font = BOLD
hdr(V, 5, ["項目", "數值", "", "", "說明"])
V.cell(6, 1, "規則值情境加權價（參考，非目標價）"); V["B6"] = f"=tp_wA*F{VAL['A']}+tp_wB*F{VAL['B']}+tp_wC*F{VAL['C']}"; V["B6"].number_format = "#,##0"
V["E6"] = "權重 × 各情境隱含價（第 3 節規則值）"
V.cell(7, 1, "規則值情境加權價 ÷ 現價 − 1（參考）"); V["B7"] = "=B6/px-1"; V["B7"].number_format = "0.0%"; V["E7"] = "現價＝In 頁 px（2026-10-23 收盤）"
for rr, s, lab in ((8, "A", "情境 A 隱含價（樂觀）"), (9, "B", "情境 B 隱含價（基準）"), (10, "C", "情境 C 隱含價（下檔）")):
    V.cell(rr, 1, lab); V[f"B{rr}"] = f"=F{VAL[s]}"; V[f"B{rr}"].number_format = "#,##0"
    V[f"E{rr}"] = f'="近 12 月 GAAP 換算 EPS "&TEXT(D{VAL[s]},"0.00")&" × "&TEXT(E{VAL[s]},"0.0")&" 倍（第 3 節）"'
V.cell(11, 1, "情境權重 A／B／C"); V["B11"] = "=tp_wA"; V["C11"] = "=tp_wB"; V["D11"] = "=tp_wC"
for c in "BCD": V[f"{c}11"].number_format = "0%"
V["E11"] = '="合計 "&TEXT(B11+C11+D11,"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
V.cell(12, 1, "舊法說明")
V["B12"] = ("P/E：各情境 2027 年底近 12 月 GAAP 換算 EPS（≈ ½ FY2027＋½ FY2028 non-GAAP EPS × g2n 0.70）× 十年年底 GAAP P/E 依循環位置取分位（非設備股：以本模型營收最高年判定，三情境皆為中段＝中位數）。"
            "Credo 2022 年上市，十年倍數只有 2025 年一個正值（118x，FY2026 營收年增 206% 的高成長期；2023、2024 年虧損為負值不用），偏誤方向為明顯高估；隱含價約為 2027 年底價格、未折現。"
            "對照：現價只約 23 倍共識 FY2027 non-GAAP EPS、15 倍 FY2028（第 1 節）")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額在每個案例內依 FY2027 錨點重算，FY2027 恆等於錨點）：對 FY2028、FY2029 的營收、AEC 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2028 營收", "FY2028 AEC 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準", "FY2029 營收", "FY2029 AEC 營收", "FY2029 毛利率", "FY2029 EPS", "FY2029 EPS 相對基準"])
cases = [("AEC 份額流失（低）", "AEC 有效份額相對值 FY2028 0.75、FY2029 0.60（Astera、Marvell、Broadcom 與線纜廠搶份額；或 ALC 自我替代）", dict(rela="RLO")),
         ("AEC 份額上升（高）", "FY2028 1.15、FY2029 1.20（1.6T 使 DAC 退場，AEC 滲透率上升）", dict(rela="RHI")),
         ("最大客戶流失", "FY2028 起 AEC × (1 − 33%)（最大終端客戶轉用其他供應商或光學）", dict(rela="cl")),
         ("光學份額 低", "光學有效份額相對值 0.90（設計案延後）", dict(relo="OLO")), ("光學份額 高", "FY2028 1.60、FY2029 2.00", dict(relo="OHI")),
         ("毛利率 低（價格戰）", "毛利率 FY2028 63%、FY2029 60%", dict(gm="GMLO")), ("毛利率 高", "毛利率 69%", dict(gm="GMHI")),
         ("稅率 17%", "FY2028 起 non-GAAP 稅率 17%", dict(tx="TXHI")), ("稅率 5%", "全球最低稅延後或抵減", dict(tx="TXLO")),
         ("AI 銅纜市場 ×0.8", "FY2028 起銅纜市場 ×0.8（銅轉光提前）", dict(mf="mf_lo")), ("AI 銅纜市場 ×1.2", "×1.2（scale-up 銅擴大）", dict(mf="mf_hi")),
         ("營業費用彈性 0.9", "營業費用隨營收成長較多", dict(phi="phi_hi")), ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi"))]
for nm, base, suf in (("RLO", "rel_a", "_lo"), ("RHI", "rel_a", "_hi"), ("OLO", "rel_o", "_lo"), ("OHI", "rel_o", "_hi"), ("GMLO", "gm", "_lo"), ("GMHI", "gm", "_hi"), ("TXHI", "tx", "_hi"), ("TXLO", "tx", "_lo")):
    for y in (28, 29):
        wb.defined_names[f"{nm}_{y}"] = DefinedName(f"{nm}_{y}", attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2028, "#,##0.00"), ("D", "aec", 2028, "#,##0.00"), ("E", "gm", 2028, "0.0%"), ("F", "eps", 2028, "0.00"),
                           ("H", "rev", 2029, "#,##0.00"), ("I", "aec", 2029, "#,##0.00"), ("J", "gm", 2029, "0.0%"), ("K", "eps", 2029, "0.00")):
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
SE.cell(nn + 3, 1, "讀法：FY2027 為錨點、各案例相同；顆數情境只影響 FY2028 的 4/12（CY2028），差異主要在 FY2029。「AEC 份額流失」「最大客戶流失」「AI 銅纜市場 ×0.8」是 FY2028 EPS 最大的三條不利路徑；FY2028 起稅率 17%（OECD 最低稅）影響較小；顆數情境 C 主要打在 FY2029。光學在 FY2029 前仍小於 AEC。")
SENS_READ = nn + 3
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 60
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、券商、研調。只作事後對照，不為對上數字而調參數。共識 FY2028 EPS 區間 8.69–11.31"
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
RC.cell(rr_, 1, "A. 共識與指引（校準未使用共識）").font = BOLD; rr_ += 1
rec("2Q27 營收（指引中點）vs 共識", "=cons_q2", lambda s: "=Q!E4", "Yahoo（17 家）", "指引 5.25–5.35")
rec("FY2027 營收 vs 共識", "=cons_rev27", pl("rev", 2027), "Yahoo（19 家；2.47–2.58）", "錨點＝公司「年增逾 85%」下緣")
rec("FY2027 EPS vs 共識", "=cons_eps27", pl("eps", 2027), "Yahoo（20 家；5.98–6.64）", "")
rec("FY2027 淨利率 vs 公司目標約 50%", 0.5, pl("nim", 2027), "公司 1Q27 法說", "", fmt="0.0%", pt=True)
rec("FY2028 營收 vs 共識", "=cons_rev28", pl("rev", 2028), "Yahoo（19 家；3.59–4.41）", "")
rec("FY2028 EPS vs 共識", "=cons_eps28", pl("eps", 2028), "Yahoo（20 家；8.69–11.31）", "模型 FY2028 稅率 13%（OECD 最低稅）；共識是否已含：未找到")
rec("FY2028 AEC 營收 vs Jefferies 引述共識", "=jef_aec", pl("aec", 2028), "Jefferies（Invezz 轉述，2026-04）", "共識期間未明示")
RC.cell(rr_, 1, "B. 份額與市場（機制檢查）").font = BOLD; rr_ += 1
rec("FY2027 隱含 AEC 單價（＝AEC 營收 ÷〔Jefferies 條數 × Credo 占 AEC〕；美元／條）", None,
    lambda s: f"=PL!E{PB[s]['aec']}*1000/((w8*jef_u26+(1-w8)*jef_u27)*cr_aec_sh)", "Jefferies 條數（FY 加權）", "800G AEC 單價一般認知為數百美元（第一手：未找到）", fmt="#,##0")
rec("FY2027 Credo AEC ÷ scale-out 銅纜（網通模型）", None, lambda s: f"=Units!C{UR['shs']}", "markets/network", "", fmt="0.0%")
rec("FY2028 AEC 有效份額（÷ 銅纜市場）", None, pl("sha", 2028), "—", "", fmt="0.00%")
rec("FY2028 Credo 營收 ÷ AI IT 資本支出", None, lambda s: f"=PL!F{PB[s]['rev']}/(w8*itc_27+(1-w8)*itc_28)", "markets/network", "每 1 美元 AI IT 資本支出分到的 Credo 營收", fmt="0.00%")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("1Q27 EPS（模型）vs 公司揭露 1.20", "=q1_eps", lambda s: "=Q!D17", "公司", "淨利 ÷ 稀釋股數", fmt="0.00")
rec("FY2026 EPS（模型）vs 公司揭露 3.46", "=a26_eps_in", pl("eps", 2026), "公司", "", fmt="0.00")
rec("FY2027 營業費用彈性（Derived）", None, lambda s: "=Q!B28", "Q 頁", "", fmt="0.00")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以 Credo 營收 ÷ 部署 GW 計算每 GW 營收"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "2Q27 財報（約 2026-12 初）：公布後以實績取代 Q 頁 2Q 指引，並更新全年指引"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_END = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝SEC 8-K 新聞稿與 10-K）"; SO["A1"].font = BOLD
srcs = [("公司：4Q26 與 FY2026 實績、1Q27 指引（8-K 附件 99.1）", PR4),
        ("公司：1Q27 實績、2Q27 指引（8-K 附件 99.1）", PR1),
        ("公司：FY2026 10-K（客戶集中、DustPhotonics、OECD 最低稅、股數、地區營收）", K10),
        ("二手：4Q26 法說摘要（FY2027 年增逾 80%、光學逾 6 億美元、營業費用約 +50%）", LDS),
        ("二手：1Q27 摘要（年增逾 85%、淨利率約 50%、AEC 占營收年增九成以上）", TIG),
        ("二手：1Q27 法說重點（Zacks；前四大終端客戶 33%、28%、13%、10%；五家超大規模客戶）", ZAC),
        ("券商：Jefferies AEC 條數與份額（Invezz 轉述）", JEF),
        ("券商：Needham 等（客戶名稱 Amazon、Microsoft、xAI；Benzinga 2025-03）", BZ),
        ("共識：Yahoo Finance 分析師預估（2026-10-01）", YH),
        ("共識：StockAnalysis 預估與目標價（2026-10-05）", SAF),
        ("價格：StockAnalysis（2026-10-23 收盤、52 週區間、股數）", SA),
        ("價格與評論：TIKR（10 月跌幅、OCP 發表 ALC、FY2027 共識 GAAP／調整後 EPS）", TK),
        ("資料商：companiesmarketcap 年底 P/E", CMC),
        ("資料商：StockAnalysis 會計年度末 P/E", "https://stockanalysis.com/stocks/crdo/financials/ratios/"),
        ("市場層：markets/network、tsmc 成品（本 repo，唯讀）", ""),
        ("待補（人工）：2Q27 財報；產品線營收拆分（AEC、光學、IC、IP）；客戶名稱第一手出處；AEC 市場份額第一手來源；共識 FY2028 是否含 OECD 最低稅；P/E 第二來源；前瞻 P/E 歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Credo（CRDO）收支與估值模型 v0.2（AI 半導體 Project：網通次產業內複製——主動電纜 AEC、SerDes、光 DSP）　2026-10-25", True),
 ("版本：v0.2（2026-10-25：三情境延伸到 FY2032、目標價＝DCF 與 Comparable 加權；v0.1 初版）。Credo 在 AI 銅纜（AEC）與光學的份額在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）。", False),
 ("模型命題：Credo 的 AEC 營收＝AI 銅纜市場（本專案網通模型：scale-up 背板與線纜＋scale-out DAC／AEC）× Credo 有效份額（AEC 滲透率 × Credo 占 AEC）；光學（光 DSP、矽光 PIC、ZeroFlap 光學）＝可插拔光模組市場 × 光學份額；線卡 PHY、重定時器與 SerDes IP 另依成長率推。股價自 52 週高點跌約 52%、只約 15 倍共識 FY2028 EPS：市場定價像是 AEC 成長在 FY2028 前後見頂。", False),
 ("核心問題：市場隱含的 FY2028 EPS 是否低於模型——要成立，AEC 份額（競爭、主動光纜自我替代）、客戶集中（前四大終端客戶占 84%）、銅轉光（AI 銅纜市場）三條路徑各要不利到什麼程度？", False), ("", False),
 ("會計年度：FY(t) 於 t 年 4 月底／5 月初結束（FY2027＝2026-05-03～2027-05-01），8 個月落在日曆年 t−1；上游市場層為日曆年，FY(t) 市場＝8/12 × CY(t−1)＋4/12 × CY(t)。預測期 FY2028–FY2029（FY2027 為錨點）。口徑：公司 non-GAAP；十年 P/E 規則值用 GAAP 換算 EPS。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AEC：markets/network 的 AI 銅纜市場（FY 加權）× 情境顆數係數 × Credo AEC 有效份額（FY2027 錨點校準）× 份額相對值（Assumed）。光學：可插拔光模組市場 × 光學有效份額（FY2027 以公司目標逾 6 億美元校準）× 相對值（Assumed）。", False),
 ("其他 IC 與 IP：FY2026 營收 × 7%（Assumed）× 成長率。毛利率：FY2028 起為輸入路徑（Assumed）。營業費用：彈性 phi（FY2027 錨點 Derived）。稅率：FY2028 起 13–14%（Assumed，OECD 最低稅）。", False),
 ("校準：FY2027 ＝ 1Q27 實績（營收 4.79 億美元、non-GAAP EPS 1.20）＋ 2Q27 指引中點（5.30 億美元）＋ 2H27（全年年增 85% 下緣減 1H）。產品線拆分公司未揭露：光學 FY2026 為假設、FY2027 用公司目標，AEC 為殘差。不用共識（只在 Recon 對帳）。", False),
 ("估值：目標價（Val 第 0 節）＝0.5 × DCF（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 0.3／0.5／0.2）＋0.5 × Comparable（美股高速連接晶片同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）。舊法規則值情境加權保留參考（不是目標價）；市場隱含 FY2028 EPS（Val 第 4 節）；十年 P/E 規則值（只有 2025 年一個正值）× 2027 年底近 12 月 GAAP 換算 EPS；循環位置以本模型營收最高年判定。Val 第 0 節為情境加權目標價。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。客戶名稱（Amazon、xAI、Microsoft）只有券商與媒體來源；公司 10-K 只揭露未具名的客戶占比。", False),
 ("", False),
 ("已知限制：(1) AEC 與光學份額相對值、毛利率路徑、FY2028 起稅率皆為 Assumed，是 FY2028–FY2029 EPS 最敏感的輸入（Sens）；(2) 產品線營收未揭露，AEC 為殘差；(3) 網通模型的銅纜市場以 NVIDIA 平台為主，Credo 客戶以 ASIC 與超大規模業者為主，份額校準吸收了這個口徑差；(4) 十年 P/E 只有 2025 年一個觀測值、含高成長期溢價，規則值與目標價明顯偏高；(5) 會計年度與日曆年以月數加權近似；(6) R4 產業部署路徑未建（Recon 預留列）。", False)]
lines.append(("v0.2（2026-10-25）：Val 頁最前面新增「0. 目標價」（0.1 DCF、0.2 Comparable、0.3 結論），既有內容整體下移；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留。PL 三情境延伸到 FY2032（FY2030 起為收斂法，非驅動因子推導；灰底），說明欄移到 K 欄。In 頁新增 K 節（tp_r、tp_g、tp_roe、tp_w_dcf、收斂法參數、共識目標價）。Sens 加目標價參數列。FY2026–FY2029 既有數字不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170


# ================= v0.2（2026-10-25）：FY2030–FY2032 收斂法；目標價＝DCF 與 Comparable 加權 =================
# 工作單：Andy 2026-10-25 11:18／11:20（三情境延伸到 FY2032、收斂法）、11:28（目標價＝DCF＋Comparable）、11:53（tp_w_dcf 0.5；A 股與其他市場同業分開）
# 本節在其他工作表都建好之後執行：(1) PL 說明欄由 H 移到 K，H–J 放 FY2030–32 收斂年度；(2) Val 既有內容整體下移，最前面放「0. 目標價」；(3) Sens 加目標價參數列
import re as _re
from copy import copy as _copy
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
FCX = dict(FC); FCX.update({2030: "H", 2031: "I", 2032: "J"})
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
INR = {IN.cell(i, 1).value: i for i in range(4, IN.max_row + 1) if IN.cell(i, 1).value}
for _k, _i in INR.items():
    if str(_k).startswith(("tp_", "fade_")) or _k in ("cons_tp", "cons_feps"):
        for _j in range(1, 10): IN.cell(_i, _j).fill = KEY

def _move_notes(ws):
    """說明欄（第 8 欄，H）移到第 11 欄（K），讓出 H–J；表頭列補上收斂年度。"""
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 8)
        if c.value is None: continue
        assert not (isinstance(c.value, str) and c.value.startswith("=")), (ws.title, r, c.value)
        k = ws.cell(r, 11); k.value = c.value; k.font = _copy(c.font); k.fill = _copy(c.fill)
        c.value = None; c.font = Font(); c.fill = PatternFill()
        if str(ws.cell(r, 1).value or "") in ("項目", "情境／項目"):
            for j, y in zip((8, 9, 10), FADE):
                h = ws.cell(r, j, f"FY{y}（收斂）"); h.font = BOLD; h.fill = HDR
    ws.column_dimensions["K"].width = ws.column_dimensions["H"].width or 70
    for col in "HIJ": ws.column_dimensions[col].width = 14

def _fade(ws, R, s, eps_f):
    """FY2030–32：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到）；非驅動因子推導。"""
    for y in FADE:
        c, pc, k = FCX[y], FCX[y - 1], y - 2029
        fs = {"g": f"=$G{R['g']}+(fade_g{s}-$G{R['g']})*{k}/3", "rev": f"={pc}{R['rev']}*(1+{c}{R['g']})",
              "nim": f"=$G{R['nim']}+(fade_m{s}-$G{R['nim']})*{k}/3", "ni": f"={c}{R['rev']}*{c}{R['nim']}",
              "sh_": f"={pc}{R['sh_']}*(1+sh_g)", "eps": "=" + eps_f(c, R), "epsg": f"={c}{R['eps']}/{pc}{R['eps']}-1"}
        for key, f in fs.items():
            ws[f"{c}{R[key]}"] = f; ws[f"{c}{R[key]}"].number_format = ws[f"G{R[key]}"].number_format; ws[f"{c}{R[key]}"].fill = FADEFILL

def _shift_rows(wb, sheet, start, n):
    """把 sheet 第 start 列起的內容整體下移 n 列，並同步改寫所有公式與定義名稱中指向這些列的參照（等同插入列）。"""
    ws = wb[sheet]
    cell_ = r"\$?[A-Z]{1,3}\$?\d+"
    pat = _re.compile(r"(?<![A-Za-z0-9_.!$:])((?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?(" + cell_ + r")(?::(" + cell_ + r"))?(?![A-Za-z0-9_(!])")
    one = _re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")
    def fix(f, own):
        parts = f.split('"')
        for i in range(0, len(parts), 2):
            def rep(m):
                pre = m.group(1)
                tgt = (pre is None and own) or (pre is not None and pre.strip("'!") == sheet)
                def sh(x):
                    a, rr = one.fullmatch(x).groups()
                    rr = int(rr)
                    return f"{a}{rr + n if (tgt and rr >= start) else rr}"
                return (pre or "") + sh(m.group(2)) + (":" + sh(m.group(3)) if m.group(3) else "")
            parts[i] = pat.sub(rep, parts[i])
        return '"'.join(parts)
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = fix(c.value, w.title == sheet)
    for nm in list(wb.defined_names.keys()):
        d = wb.defined_names[nm]
        if sheet + "!" in d.attr_text: d.attr_text = fix(d.attr_text, False)
    cells = [(c.row, c.column) for row in ws.iter_rows(min_row=start) for c in row if c.value is not None or c.has_style]
    for r, col in sorted(cells, reverse=True):
        src = ws.cell(r, col); dst = ws.cell(r + n, col)
        dst.value = src.value; dst._style = _copy(src._style); dst.number_format = src.number_format
        src.value = None; src._style = _copy(ws.cell(1, 30)._style)

def _tp_section(V, eps_ref, fwd, fwd_note, peers, peer_note, cons_note, method, pe_rng, fmt, cons_feps_note):
    """Val 第 0 節（第 4–48 列）：0.1 DCF、0.2 Comparable、0.3 結論。列標籤依工作單逐字；B／C／D＝情境 A 順利／B 平平／C 不如預期。"""
    T = {}
    def tset(r, key, label, fs, nf, note=""):
        V.cell(r, 1, label); T[key] = r
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = nf
        if note: V[f"E{r}"] = note
    V.cell(4, 1, "0. 目標價").font = BOLD
    V.cell(5, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
    hdr(V, 6, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
    V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境；FY2030 起為收斂法（PL 頁灰底）"
    for i, y in enumerate(YRS6):
        tset(8 + i, f"eps{y}", f"  FY{y}", [f"={eps_ref(s, y)}" for s in SC], "0.00")
    V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）")
    tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=IF(tp_roe<=tp_g,0.8,MIN(tp_g/tp_roe,0.8))", None, None], "0.0%", "ROE＝In 頁 tp_roe；tp_roe ≤ tp_g 時取上限 0.8")
    for i, y in enumerate(YRS6):
        tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{8 + i}*(1-$B$15)" for c in "BCD"], "0.00")
    tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], fmt,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027")
    tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], fmt)
    tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], fmt, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
    V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
    hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
    assert 3 <= len(peers) <= 5, peers
    for i, (nm, tk, pe, dt, src) in enumerate(peers, 28):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{27 + len(peers)})", None, None], "0.0", peer_note)
    tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）",
         ["=" + "+".join(f"tp_w{s}*(" + "+".join(f"{w}*{c}{8 + YRS6.index(y)}" for y, w in fwd) + ")" for s, c in zip(SC, "BCD")), None, None], "0.00", fwd_note)
    tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], fmt, "同業中位數 × 前瞻 EPS")
    V.cell(37, 1, "0.3 結論").font = BOLD
    hdr(V, 38, ["項目", "數值", "", "", "說明"])
    tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], fmt, "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
    tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], fmt)
    tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], fmt, cons_note)
    tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px")
    tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
    tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", cons_feps_note)
    tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
    V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
    tset(46, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D23,"{fmt}")&"～"&TEXT(B23,"{fmt}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}9>MAX({pe_rng}),1,IF(px/{c}9<MIN({pe_rng}),0,PERCENTRANK({pe_rng},px/{c}9)))" for c in "BCD"], "0%",
         "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 6 節採用來源）的分位；高於十年最大值＝100%；只作檢查，不進目標價")
    V.cell(48, 1, "目標價方法"); V["B48"] = method
    for r_ in (24, 35, 39):
        V[f"B{r_}"].fill = KEY
    return T

def _tp_sens(SE, T):
    """Sens：目標價參數（tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.3／0.7）；放在案例摘要與第 40 列推導區塊之間。"""
    last = max(r for r in range(1, 40) if any(SE.cell(r, j).value is not None for j in range(1, 13)))
    t0 = last + 2
    SE.cell(t0, 1, "目標價參數（只改 DCF 參數與權重；PL 收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
    hdr(SE, t0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_f(c, r, g, reinv):
        cf = lambda i: f"Val!{c}{8 + i}*(1-{reinv})"
        pv = "+".join(f"{cf(i)}/(1+{r})^{i + 1}" for i in range(6))
        return f"({pv}+{cf(5)}*(1+{g})/({r}-{g})/(1+{r})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = t0 + 2 + k; r = kw.get("r", "tp_r"); g = kw.get("g", "tp_g"); w = kw.get("w", "tp_w_dcf")
        reinv = f"IF(tp_roe<={g},0.8,MIN({g}/tp_roe,0.8))"
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_f(c, r, g, reinv)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${t0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = "0.00"
        SE[f"F{i}"].number_format = "0.0%"
    assert t0 + 1 + len(cases) < 40, t0
    return t0, t0 + 1 + len(cases)

def _fade_all(eps_f, nim10):
    """PL：三情境推導區塊與頂端摘要延伸到 FY2032；In 頁 fade_m 連到 PL 頁 FY2029 淨利率。nim10＝十年平均淨利率的名稱；None＝未找到，依工作單用 FY2029 × 0.8。"""
    _move_notes(PL)
    for s in SC:
        _fade(PL, PB[s], s, eps_f)
        for key in ("rev", "g", "nim", "ni", "eps", "epsg"):
            if (s, key) not in SUMR: continue
            for y in FADE:
                cc = PL[f"{FCX[y]}{SUMR[(s, key)]}"]; cc.value = f"={FCX[y]}{PB[s][key]}"
                cc.number_format = PL[f"G{SUMR[(s, key)]}"].number_format; cc.fill = FADEFILL
    nimr = {s: f"PL!$G${PB[s]['nim']}" for s in SC}
    IN.cell(INR["fade_mA"], 5, f"={nimr['A']}")
    IN.cell(INR["fade_mB"], 5, f"=({nimr['B']}+{nim10})/2" if nim10 else f"=({nimr['B']}+{nimr['B']}*0.8)/2")
    IN.cell(INR["fade_mC"], 5, f"={nim10}" if nim10 else f"={nimr['C']}*0.8")
    PL["A2"] = ("灰底＝FY2030–FY2032 收斂法，非驅動因子推導（上游市場層只到 2029）：營收年增由 FY2029 線性收斂到 fade_g（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt），"
                "淨利率由 FY2029 線性收斂到 fade_m（In 頁 K 節），FY2032 達到；股數依 sh_g")

def _rule_pe_range(V, ncols=1):
    """第 6 節十年倍數表（2016–2025 列）中採用來源的欄位範圍（PERCENTRANK 用；ncols＞1 時各來源合併成一個分布）。"""
    yr = [r for r in range(1, V.max_row + 1) if V.cell(r, 1).value in [str(y) for y in range(2016, 2026)]]
    yr = [r for r in yr if r > 50][:10]
    assert len(yr) == 10 and yr[-1] - yr[0] == 9, yr
    return f"$B${yr[0]}:${chr(65 + ncols)}${yr[-1]}"

_fade_all(lambda c, R: f"{c}{R['ni']}/{c}{R['sh_']}", None)
_shift_rows(wb, "Val", 4, 47)
TPT = _tp_section(V, lambda s_, y: f"PL!{FCX[y]}{SUMR[(s_, 'eps')]}", [(2028, '6/12'), (2029, '6/12')],
    '約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 6／12 × FY2028＋6／12 × FY2029（會計年度 4 月底／5 月初結束）',
    [('Astera Labs', 'ALAB', 45.77, '2026-10-23', 'stockanalysis forward PE（分析師共識；價格日期如左）'), ('Marvell', 'MRVL', 40.44, '2026-10-23', 'stockanalysis forward PE（分析師共識；價格日期如左）'), ('Broadcom', 'AVGO', 18.52, '2026-10-23', 'stockanalysis forward PE（分析師共識；價格日期如左）'), ('MACOM', 'MTSI', 28.59, '2026-10-23', 'stockanalysis forward PE（分析師共識；價格日期如左）')],
    '美股高速連接晶片同業（Astera Labs、Marvell、Broadcom、MACOM；不放 A 股）',
    'In 頁 tp_street：20 家平均 279.09（2026-10-05）；10-23 為 284.05',
    'DCF：股東現金流＝EPS（non-GAAP）×（1 − g／ROE），r＝10%、g＝4%、ROE 需與 non-GAAP EPS 同口徑，本模型無股東權益、算不出，取 0.20（Assumed；再投資率 20%），FY2027–32 明確期＋永續，FY2030 起為收斂法。Comparable：美股高速連接晶片同業（Astera Labs、Marvell、Broadcom、MACOM）前瞻 P/E 中位數 × 12 個月後前瞻 EPS（non-GAAP）；Broadcom 規模與業務組合不同，拉低中位數。會計年度 4 月底結束：FY2027 為錨點年（2026-05～2027-04），DCF 的 t＝1 對 FY2027，偏高估。十年 P/E 只有 2025 年一個觀測值（GAAP），分位檢查參考性低。情境對應：顆數情境 A＝順利、B＝平平、C＝不如預期。',
    _rule_pe_range(V, 1), '#,##0.00',
    "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝In 頁 cons_feps")
TPS0, TPS_LAST = _tp_sens(SE, TPT)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("CRDO_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "UN_NOTE": UN_NOTE, "CHK": CHK_ROW, "SENSN": SENS_N, "SENS_READ": SENS_READ,
               "VAL": VAL, "IMPL": {str(k): v for k, v in IMPL.items()}, "READ": READ_ROW, "NOPRIOR": NOPRIOR_ROW, "RECON_END": RECON_END, "rule_rows": rule_rows, "SB": SB},
              open(os.environ["CRDO_ROWMAP"], "w"), ensure_ascii=False, indent=1)
