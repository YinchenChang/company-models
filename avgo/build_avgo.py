# Broadcom（AVGO）收支與估值模型 v0.2.1（AI 半導體 Project：運算層第二個樣板，ASIC＋網通）
# 用法：python3 avgo/build_avgo.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 FY(t) 於 t 年 10 月底／11 月初結束（FY2026：2025-11-03～2026-11-01），10/12 落在日曆年 t，
# 所以 FY(t) 直接對照 CY(t)，不做平移（NVIDIA 需平移一年，AVGO 不需）。金額：十億美元（US$B）；EPS：美元；口徑：公司 non-GAAP。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "nvidia"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # 本資料夾的 tp_v2.py 優先
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import tp_v2 as TP
import fcff_dcf as FD   # v0.5：標準版 DCF（無槓桿自由現金流＋WACC）共用模組（tools/fcff_dcf.py，原樣取自 NVIDIA v0.3）
import accel_share as AS          # 加速器金額份額的可重用函式（NVIDIA 模型建立；AVGO 為 ASIC 側）

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
XP = ("tpu", "oth")                                   # Broadcom 的 XPU 來自這兩個市場層平台
PNAME = {k: nm for k, nm, *_ in MEM.UNITS}
PLAT = [k for k, *_ in MEM.UNITS]

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx"); SKH_F = only("skhynix/dist/*.xlsx"); NET_F = only("markets/network/dist/*.xlsx")
NV_F = only("nvidia/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

ws_ = openpyxl.load_workbook(TSMC_F, data_only=True)["In"]
TS = {}
for r in ws_.iter_rows(min_row=4):
    k = r[0].value
    if k and (k.startswith("um2") or k.startswith("g29_") or k.startswith("u_")):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
for k, nm, u_, gb, tag, src in MEM.UNITS:     # 與 markets/memory 的 UNITS 必須一致（同一份）
    for i, y in enumerate((26, 27, 28)):
        assert abs(TS[f"u_{k}_{y}"][0] - u_[i]) < 1e-9, (k, y)
for s_ in SC:
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9

wsn = openpyxl.load_workbook(NET_F, data_only=True)
NETB = None; REPB = None
for r in wsn["Mkt"].iter_rows(min_row=4):
    if r[0].value == "Broadcom 網通（本模型歸屬）":
        NETB = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
for r in wsn["In"].iter_rows(min_row=4):
    if r[0].value == "rep_avgo26":
        REPB = ((r[3].value, r[4].value, r[5].value), f"{rel(NET_F)}｜In!D{r[0].row}:F{r[0].row}")
assert NETB and REPB and all(v is not None for v in NETB[0])

wsk = openpyxl.load_workbook(SKH_F, data_only=True)
HBMP = {}
for r in wsk["Mkt"].iter_rows(min_row=4):
    lab = str(r[0].value or "")
    if lab.startswith("HBM 單價（"):
        HBMP[lab[-2]] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert sorted(HBMP) == list(SC), HBMP.keys()

wsv = openpyxl.load_workbook(NV_F, data_only=True)["In"]
NVP = {}
for r in wsv.iter_rows(min_row=4):
    if r[0].value in ("pm_bw", "pm_rb", "pm_ru", "pm_hp", "pa_amd", "pa_trn", "pa_tpu", "pa_oth"):
        NVP[r[0].value] = (r[4].value, f"{rel(NV_F)}｜In!E{r[0].row}")
assert len(NVP) == 8, NVP

PEAK, PEAK_SRC = CP.peak_years(ROOT)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；每顆價格：千美元（US$K）；顆數：百萬顆；HBM：GB、US$/GB。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝Broadcom 8-K 新聞稿（SEC EDGAR、investors.broadcom.com）；公司對 FY2027、FY2028 的 AI 營收展望為 Interested-party，只放 Recon 對帳，不作校準"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
EDG = "https://www.sec.gov/Archives/edgar/data/1730168/"
Q3PR = "https://investors.broadcom.com/news-releases/news-release-details/broadcom-inc-announces-third-quarter-fiscal-year-2026-financial"
CALL = "法說轉述：Investing.com 逐字稿、BigGo Finance 摘要（2026-09-02；二手）"

sec("A. 加速器顆數（markets/memory 與 HBM、台積電、網通、NVIDIA 模型同一份；情境係數讀台積電模型成品，不另訂）")
for k, nm, u_, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u_[i], None, tag, "markets/memory/market.py UNITS：" + src if i == 0 else "同上", f"台積電模型 {TS[f'u_{k}_{y}'][1]}（值相同，建置時檢查）" if i == 0 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", None, TS[f"g29_{k}"][0], None, "Assumed", TS[f"g29_{k}"][1], "markets/memory 無 2029；沿用台積電模型")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
for k in XP:
    gb = [g for kk, nm, u_, g, t, sr in MEM.UNITS if kk == k][0]
    for i, y in enumerate((26, 27, 28)):
        inp(f"gb_{k}_{y}", f"{PNAME[k]} 每顆 HBM 20{y}", "GB", None, gb[i], None, "Interested-party" if k == "tpu" else "Derived／Assumed", "markets/memory/market.py UNITS", "2029 沿用 2028" if i == 0 else "同上")

sec("B. ASIC 每顆規格（本專案為權威，規則第 13 條）：Broadcom 在各平台的顆數份額、每顆核心售價（不含 HBM）、HBM 由 Broadcom 代購的比例")
for _k in ("sh_tpu_26", "sh_tpu_27", "sh_tpu_28", "sh_oth_26", "sh_oth_27", "sh_oth_28", "core_tpu", "core_oth", "f_hbm"):   # Siliconomics In 頁 M6、M7（v0.3 起）
    inp(*SIL.row9(_k))
inp("m_hbm", "代購 HBM 的毛利率", "x", 0.0, 0.10, 0.20, "Assumed", "無：未揭露；轉售料件通常低加成", "第二來源：未找到")
inp("pfc", "2027 年起 XPU 核心售價係數（Sens 用；基準 1）", "x", 0.8, 1.0, 1.2, "Assumed", "無：新一代（TPU v8、MTIA 新世代）核心單價變化未揭露", "Sens 用 ±20%")

sec("C. 記憶體價格路徑（讀 skhynix 成品；唯讀；情境別）")
for s in SC:
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"hbm_{s}_{y}", f"HBM 混合單價 20{y}（{s}）", "US$/GB", None, HBMP[s][0][j + 1], None, "Derived", HBMP[s][1], "markets/memory：2026 BofA $14.6/GB；2027 +80% 三情境相同；2028–29 依情境" if (s == "A" and y == 26) else "同上")

sec("D. AI 網通（讀 markets/network 成品；情境 A（原始顆數），B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"net_{y}", f"Broadcom 網通（本模型歸屬）20{y}", "US$B", None, NETB[0][j], None, "Derived", NETB[1], "廠商歸屬比例為 Assumed（markets/network In 頁 avgo_su、avgo_so、avgo_nic、avgo_opt）；2026 較實際推估低約 28%（見 Recon）" if y == 26 else "同上")
inp("rep_avgo26", "Broadcom AI 網通營收 FY2026（實際＋推估）", "US$B", REPB[0][0], REPB[0][1], REPB[0][2], "Derived", REPB[1], "Q3 FY26：XPU 占 AI 營收 73%（法說，二手）→ 網通約 4.5；FY2026 AI 57.6 × 約 30%")
inp("gm_net", "AI 網通毛利率", "x", 0.65, 0.72, 0.80, "Assumed", "無：公司不揭露產品別毛利率", "第二來源：未找到")

sec("E. 非 AI 半導體與基礎架構軟體（VMware 等）")
inp("gm_nai", "非 AI 半導體毛利率", "x", 0.60, 0.65, 0.70, "Assumed", "無：公司不揭露；半導體部門整體毛利率約 67%（Q3 FY26 法說 Q&A 更正值，二手）", "第二來源：未找到")
inp("gm_sw", "軟體毛利率（non-GAAP）", "x", 0.90, 0.93, 0.95, "Derived", "Q3 FY26：(non-GAAP 毛利 22.191 − 半導體 20.839 × 67%) ÷ 軟體 8.752 ≈ 94%（67% 為二手）", "低／高 Assumed")
inp("g_nai", "非 AI 半導體營收年增（FY2027 起）", "%", -0.05, 0.03, 0.08, "Assumed", "FY2026 Q3 約 4.1／季（半導體 20.8 − AI 16.7；公司）", "第二來源：未找到")
inp("g_sw27", "軟體營收年增 FY2027", "%", 0.08, 0.12, 0.18, "Assumed", "公司：Q3 FY26 軟體年增 29%、ARR 年增 15%；Q4 預估約 8.7（約 +25%）（法說，二手）", "第二來源：未找到")
inp("g_sw", "軟體營收年增 FY2028 起", "%", 0.04, 0.08, 0.12, "Assumed", "VMware 漲價效果遞減（Assumed）", "—")

sec("F. FY2026 錨點（Q1–Q3 實績＋Q4 指引）")
inp("q4_rev", "Q4 FY26 營收指引", "US$B", None, 34.8, None, "Interested-party", Q3PR + "：約 $34.8B（年增 93%）", "top1markets、BigGo（二手）同值")
inp("q4_ai", "Q4 FY26 AI 半導體營收", "US$B", None, 21.7, None, "Interested-party", "Q3 FY26 新聞稿 CEO 說明：$21.7B（年增 236%）", "BigGo、top1markets（二手）同值")
inp("q4_sw", "Q4 FY26 軟體營收", "US$B", 8.5, 8.7, 8.9, "Interested-party", CALL + "：約 $8.7B", "第二來源：未找到")
inp("q4_gm", "Q4 FY26 non-GAAP 毛利率", "%", 0.72, 0.73, 0.74, "Interested-party", CALL + "：約 73%（去年同期 78%）", "第二來源：未找到")
inp("q4_oim", "Q4 FY26 non-GAAP 營業利益率指引", "%", None, 0.66, None, "Interested-party", Q3PR + "：約 66%", "top1markets（二手）同值")
inp("q4_int", "Q4 FY26 non-GAAP 利息費用", "US$B", 0.65, 0.70, 0.75, "Derived", "Q3 FY26 0.703（公司）", "—")
inp("q4_oth", "Q4 FY26 non-GAAP 其他收益", "US$B", 0.05, 0.10, 0.15, "Derived", "Q1–Q3 0.118／0.118／0.098（公司）", "—")
inp("q4_tx", "Q4 FY26 non-GAAP 有效稅率", "%", 0.155, 0.16, 0.165, "Derived", "Q3 FY26 3.118 ÷ 19.490＝16.0%（公司）", "—")
inp("q4_sh", "Q4 FY26 non-GAAP 稀釋股數", "十億股", None, 4.94, None, "Derived", "Q3 FY26 4.937（公司）", "—")

sec("G. 營業費用、利息、稅、股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.05, 0.10, 0.30, "Derived", "FY2026 錨點營業費用年增 ÷ 營收年增（Q 頁，約 0.09）", "低／高 Assumed")
inp("g_opx_min", "營業費用年增下限", "%", 0.0, 0.03, 0.06, "Assumed", "無", "—")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("int_f", "non-GAAP 利息費用（每年）", "US$B", 2.5, 2.8, 4.0, "Assumed", "FY2026 錨點約 2.9（Q 頁）；總債務 59.4（Q3 FY26，公司）", "高端：客戶融資若改由 Broadcom 舉債承擔（見 Recon F 節）")
inp("oth_f", "non-GAAP 其他收益（每年）", "US$B", 0.2, 0.4, 0.6, "Assumed", "FY2026 約 0.43（Q 頁）", "—")
inp("tx_f", "non-GAAP 有效稅率", "%", 0.15, 0.16, 0.18, "Derived", "FY2026 Q1–Q3 約 15.6–16.5%（公司）", "—")
inp("sh_g", "稀釋股數年變動", "%", -0.015, -0.005, 0.0, "Derived", "FY2025 4.943 → Q3 FY26 4.937（公司）", "低／高 Assumed")
inp("uk", "Broadcom XPU 顆數係數（Sens 用；FY2027 起）", "x", 0.9, 1.0, 1.1, "Assumed", "無", "±10%")

sec("H. 公司展望與共識（只作 Recon 對帳，不作校準）")
inp("co_ai26", "公司 FY2026 AI 營收展望", "US$B", None, 58.0, None, "Interested-party", CALL, "Q1–Q3 實績＋Q4 指引合計 57.6（Q 頁）")
inp("co_ai27", "公司 FY2027 AI 營收展望", "US$B", None, 115.0, None, "Interested-party", CALL + "：約 $115B；公司稱不會逐季更新", "JPMorgan：FY2027 超過 $150B（KuCoin 轉述，Interested-party）")
inp("co_ai28", "公司 FY2028 AI 營收展望", "US$B", None, 230.0, None, "Interested-party", CALL + "：約 $230B", "top1markets 2026-10-22（二手）同值")
inp("jpm_ai27", "JPMorgan FY2027 AI 營收估計", "US$B", None, 150.0, None, "Interested-party", "JPMorgan（KuCoin 2026-06-22 轉述）：超過 $150B", "—")
inp("cons_rev26", "共識 FY2026 營收", "US$B", None, 105.98, None, "Interested-party", "StockAnalysis（S&P Global，40 位分析師，2026-10-22）", "—")
inp("cons_eps26", "共識 FY2026 non-GAAP EPS", "US$", None, 11.66, None, "Interested-party", "同上（區間 11.35–12.38）", "—")
inp("cons_rev27", "共識 FY2027 營收", "US$B", None, 173.97, None, "Interested-party", "同上（「明年」欄，頁面未標年度）", "—")
inp("cons_eps27", "共識 FY2027 non-GAAP EPS", "US$", None, 19.40, None, "Interested-party", "同上", "—")
inp("jpm_u27", "2027 年 ASIC 占加速器出貨顆數（JPMorgan）", "%", None, 0.536, None, "Interested-party", "JPMorgan（KuCoin 轉述）：ASIC 12.5M ÷ 23.3M", "本專案市場層 ASIC 2027 12.5M 同一來源")
inp("cp_sh27", "2027 年 Broadcom 占 ASIC 設計夥伴份額（Counterpoint）", "%", None, 0.60, None, "Interested-party", "Counterpoint（Benzinga／Yahoo 2026-01-28 轉述；口徑未說明）", "JPMorgan：高階 ASIC 80–85%")

sec("I. 加速器每顆價格（金額份額計算用；讀 NVIDIA 模型成品，模組層級）")
for k, lab in (("pm_bw", "Blackwell 模組"), ("pm_rb", "Rubin 模組"), ("pm_ru", "Rubin Ultra 模組"), ("pm_hp", "Hopper 模組"), ("pa_amd", "AMD Instinct"), ("pa_trn", "AWS Trainium"), ("pa_tpu", "Google TPU"), ("pa_oth", "其他 ASIC")):   # CF13 (a)（Andy 2026-11-05）：金額份額一律用整顆售價
    inp(k, f"{lab} 每顆", "US$K", None, NVP[k][0], None, "Derived", NVP[k][1], "NVIDIA 模型 In 頁（同一套價格；唯讀）")

sec("J. 價格、估值與客戶融資")
inp("px", "股價（2026-11-05 收盤）", "US$", None, 291.53, None, "Verified", "Yahoo Finance 日線收盤 2026-11-05", "v0.5.3 為 2026-11-04 收盤 325.48；v0.5.2 為 2026-10-23 收盤 321.68（markets/invest_map/data/prices.csv）；v0.1.1 為 10-22 收盤 323.24（StockAnalysis、YCharts）")
inp("pe_ttm_gaap", "近 12 月 GAAP P/E", "x", None, 41.26, None, "Interested-party", "StockAnalysis", "top1markets：近 12 月約 45x（10 月報導）")
inp("pe_fwd_sa", "前瞻 P/E（StockAnalysis）", "x", None, 18.61, None, "Interested-party", "StockAnalysis（口徑未說明）", "GuruFocus 17.39；valueinvesting.io 49.62（2026-10-22；口徑不明，可能為 GAAP 或舊預估，不採）")
inp("eps_q4_25", "Q4 FY25 non-GAAP EPS（近 12 月計算用）", "US$", None, 1.95, None, "Verified", EDG + "000173016825000116/avgo-11022025x8kxex99.htm", "—")
inp("pe_fair_andy", "Andy 先驗：合理 P/E", "x", None, 25, None, "Assumed", "先驗_Andy_AVGO_20261023 B9b（Drive；不在本 repo）", "只作對照")
inp("fin_anth", "Broadcom 對 Anthropic 放款上限（可轉股）", "US$B", None, 42, None, "Interested-party", "top1markets 2026-10-22 引 Reuters（Anthropic IPO 文件）", "Bloomberg：銀行團 $60B 中 $42B 優先擔保部分由 Broadcom 殘值支撐（二手）")
inp("fin_oai", "Broadcom 為 OpenAI 安排的融資", "US$B", None, 50, None, "Interested-party", "top1markets 引 WSJ：超過 $50B", "—")
inp("fin_lease", "一家客戶 AI 機架租賃擔保", "US$B", None, 29, None, "Interested-party", "top1markets 引 Broadcom 季報：最高約 $29B", "—")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "共用規則（tools/cycle_position.py）：AVGO 的 FY(t) 對照日曆年 t（不平移）" if s == "A" else "同上")
sec("目標價情境權重（v0.1.1 新增；Val 頁第 0 節用）")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
sec("目標價參數（v0.2 新增：DCF 與 Comparable 加權、FY2029 後收斂法；v0.5 起 DCF 改為標準版，折現率見 M2 節；黃底＝Andy 可改）")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.5 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單 2026-10-25）", "v0.5 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_gd", "延伸期 A／C 情境長期成長相對 tp_g 的差（A＋、C−）", "pt", None, 0.02, None, "Assumed", "工作單 2026-10-25：A＝tp_g＋2pt、C＝tp_g−2pt", "—")
inp("tp_w_dcf", "目標價中 DCF 的權重（其餘為 Comparable）", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）", "Sens 頁 0.3／0.7")
inp("nm10_k", "延伸期 C 情境淨利率目標＝FY2029 淨利率 × 本係數（同口徑十年平均未找到）", "x", None, 0.8, None, "Assumed", "工作單：找不到十年平均用最後一年 × 0.8。本模型為 non-GAAP 口徑；同口徑（non-GAAP）十年淨利率：未找到", "參考：GAAP 十年平均 22.58%（valueinvesting.io FY2016–FY2025：−13.13%、9.59%、58.80%、12.05%、12.39%、24.54%、34.62%、39.31%、11.43%、36.20%），含攤銷與一次性稅務，口徑不同不採")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, 529.10, None, "Interested-party", "stockanalysis.com（51 位分析師平均，2026-10-25 頁面；現價為 10-23 收盤）", "賣方有目標價誘因；只作對照，不進目標價")

# v0.5：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py）。評價日沿用 v0.2 的 2026 年底；股價、rf 為 2026-10-23；資產負債為 10-Q 2026-08-02（Q3 FY26）
Q10 = "Broadcom 10-Q（季度結束 2026-08-02）https://www.sec.gov/Archives/edgar/data/1730168/000173016826000080/avgo-20260802.htm"
XB = "SEC XBRL companyfacts（CIK 0001730168；https://data.sec.gov/api/xbrl/companyfacts/CIK0001730168.json；2026-11-05 讀取）"
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.2 DCF 基準日）", amt="US$B", shu="十億股",
        rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
            "FRED DGS10（同一序列，https://fred.stlouisfed.org/series/DGS10）；與股價同日（2026-10-23）；同 NVIDIA v0.3", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：US 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 迴歸，Hamada 去槓桿、現金修正），以本公司市值 D/E 與稅率重新加槓桿（βL＝βU ×（1＋(1 − tax_w) × D/E））",
                    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。Sens 低／高（Analogy；v0.5.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）"),
        kd=("=2.355/0.75/((67.12+61.079)/2)", "Derived",
            Q10 + "：前三季利息費用 2,355（US$M）年化 ÷ 期初期末平均債務本金（67,120、61,079）", "各批票據有效利率見同 10-Q 債務附註（2.6%–5.2% 區間）"),
        tax_w=("=tx_f", "Derived", "＝In 頁 tx_f（non-GAAP 有效稅率 16%，公司口徑）", "NOPAT 用 PL 頁各年稅率（FY2027 起同為 tx_f）"),
        debt=(61.079, "Verified", Q10 + "：Total debt principal outstanding 61,079（US$M；帳面 59,419＝短期 2,252＋長期 57,167）", XB + "：DebtInstrumentCarryingAmount 2026-08-02 61,079"),
        lease=(1.325, "Verified", XB + "：OperatingLeaseLiability 2025-11-02 為 1,325（流動 144＋非流動 1,181；季報未標記，取最近年報）", "Broadcom FY2025 10-K 租賃附註"),
        cash=(23.975, "Verified", Q10 + " 資產負債表：現金及約當現金 23,975（2026-08-02，US$M）", XB + "：CashAndCashEquivalentsAtCarryingValue 同值"),
        inv=(0.0, "Assumed", "非營業投資未單獨揭露（其他長期資產 8,705 含遞延稅與其他，不計）", "—"),
        mi=(0.0, "Verified", Q10 + "：無少數股權", "—"),
        stub=dict(bs_date="2026-08-02", val_date="2026-12-31", bs_src="Broadcom 10-Q（季度結束 2026-08-02，Q3 FY26，最新已公布一季；Q4 FY26 財報 2026-12 公布）；" + Q10,
                  div=("=0.65*sh_v", "Derived", Q10 + " 附註 11：2026-09-01 宣告每股季度股利 0.65 美元，2026-09-30 發放（登記日 09-21）（落在 2026-08-02～2026-12-31）；× 評價用股數 sh_v",
                       "下一次（Q4 FY26）股利尚未宣告（Q4 財報約 2026-12 中旬），不扣；買回無已宣告金額，不扣（股數沿用 Q3 稀釋股數，與 FCFF 一致）")),
        dps=("=0.65*4", "Derived", Q10 + " 附註 11：2026-09-01 宣告每股季度股利 0.65 美元；評價日 2026-12-31 到 2027-12-31 之間假設四次、金額不變", "公司每年 12 月調整股利（歷年慣例），本格不預估調升"),
        sh=("=Q!E24", "Verified", "Q3 FY26 non-GAAP 稀釋股數 4.937 十億股（Q 頁 E24；公司新聞稿）", "GAAP 稀釋 4.887（10-Q）；回購由 FCFF 支應，不另推未來股數"),
        fc_da=(None, "=(0.593/51.574+0.574/63.887+0.484/71.089)/3", None, "Derived",
               "折舊（不含併購無形資產攤銷；EBIT 為 non-GAAP、已排除該攤銷）÷ 營收：FY2024 593／51,574、FY2025 574／63,887、FY2026 前三季 484／71,089；三期平均", XB + "：Depreciation、Revenue"),
        fc_cx=(None, "=(0.548/51.574+0.623/63.887+1.013/71.089)/3", None, "Derived",
               "資本支出 ÷ 營收：FY2024 548、FY2025 623、FY2026 前三季 1,013（" + XB + "：PaymentsToAcquirePropertyPlantAndEquipment）", "Sens 用 fc_cx_d ±1pt"),
        fc_wc=(0.0, "=((7.145+2.27+5.98-1.56-2.129-11.673)/63.887+(13.707+4.523+9.968-4-1.506-13.08)/(29.591*4))/2", 0.10, "Derived",
               "營運資金 ÷ 營收：營運資金＝應收＋存貨＋其他流動資產 − 應付 − 員工薪酬 − 其他流動負債；FY2025 年底 33 ÷ 63,887＝0.1%、2026-08-02 9,612 ÷ Q3 年化營收 118.4B＝8.1%，兩期平均（" + XB + "）",
               "FY2023（併購 VMware 前）2.4%、FY2024（VMware 購買價格分攤後，合約負債與其他流動負債大增）−10.0%，不納入平均（判斷，見 PR）；Sens 低／高 0%／10%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（採用）：自算 US 半導體 同業中位數 Blume 調整（0.67 × 中位數＋0.33）", "x", None, 1.884, None, "Analogy",
    "自算產業 β：US 半導體 同業 13 家（NVDA AMD AVGO MRVL QCOM INTC MU ARM CRDO ALAB RMBS MTSI GFS）無槓桿 β 中位數（Yahoo Finance 含股利調整月收盤價 2021-11～2026-10 共 60 個月，對 S&P 500 迴歸，Hamada 去槓桿、現金修正）2.32 → Blume 調整 1.884（Andy 2026-11-05 22:55）",
    "Sens 低／高改為 β_u ∓ beta_du（v0.5.2）；取代 Damodaran 美國 Semiconductor 1.50（v0.5）")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代上一版 beta_u 低／高欄（中位數 × 0.8 再 Blume／中位數未調整）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 評價股數 sh_v)", "與 WACC 權重同一口徑")
inp("sbc_rev", "股票薪酬 ÷ 營收（EBIT 扣除；non-GAAP 營業利益排除股票薪酬，標準版 DCF 視為費用）", "%", None, "=(5.67/51.574+7.568/63.887+6.287/71.089)/3", None, "Derived",
    "股票薪酬 ÷ 營收：FY2024 5,670／51,574、FY2025 7,568／63,887、FY2026 前三季 6,287／71,089（" + XB + "：AllocatedShareBasedCompensationExpense）；三期平均 10.6%",
    "規格第 1 節：股票薪酬不加回（預設）")
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
Q["A1"] = "Q — 實績（十億美元；EPS 為美元；股數為十億股）。口徑：公司 non-GAAP（排除股票薪酬、無形資產攤銷、重組）。藍字＝輸入；黑字＝公式。實績取自 Broadcom 8-K 新聞稿（SEC EDGAR）"
Q["A2"] = "FY2026 錨點＝Q1–Q3 實績＋Q4 指引（F 欄；假設見 In 頁 F 節）。AI 營收：Q1 8.4、Q2 10.8、Q3 16.7（公司）、Q4 21.7（公司指引）；AI 網通取 markets/network 的實際推估 rep_avgo26，XPU＝AI − 網通"
hdr(Q, 3, ["項目", "FY2025A", "FY26 Q1（～2026-02-01）", "FY26 Q2（～05-03）", "FY26 Q3（～08-02）", "FY26 Q4（指引）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = ("FY2025、Q4 FY25：" + EDG + "000173016825000116/avgo-11022025x8kxex99.htm；Q1 FY26：" + EDG + "000173016826000011/avgo-02012026x8kxex99.htm；"
      "Q2 FY26：" + EDG + "000173016826000051/avgo-05032026x8kxex99.htm；Q3 FY26：" + Q3PR)
S2 = "交叉：Q3 FY26 稿的九個月合計（營收 71.089、non-GAAP 毛利 54.168、營業利益 47.849、淨利 38.631）＝Q1＋Q2＋Q3（下方檢查列）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 63.887, "C": 19.311, "D": 22.187, "E": 29.591, "F": "=q4_rev"}, tag="Verified／Interested-party（Q4 指引）")
qrow(5, "  半導體解決方案", {"B": 36.858, "C": 12.515, "D": 15.009, "E": 20.839, "F": "=F4-F6"}, tag="Verified／Derived（Q4）")
qrow(6, "  基礎架構軟體", {"B": 27.029, "C": 6.796, "D": 7.178, "E": 8.752, "F": "=q4_sw"}, tag="Verified／Interested-party（Q4）")
qrow(7, "AI 半導體營收", {"B": "=co_ai26/(1+1.86)", "C": 8.4, "D": 10.8, "E": 16.7, "F": "=q4_ai"}, tag="Verified／Interested-party（Q4）／Derived（FY2025）", s2="FY2025＝公司 FY2026 展望 58 ÷（1＋186%）≈ 20.3（JPMorgan：約 20，二手）")
qrow(8, "  AI 網通（實際推估；markets/network rep_avgo26）", {"G": "=rep_avgo26"}, ann="", tag="Derived", s1=REPB[1], s2="Q3 FY26：XPU 占 AI 73%（法說，二手，單一來源）")
qrow(9, "  XPU（＝AI − 網通）", {"G": "=G7-G8"}, ann="", tag="Derived", s1="—", s2="FY2026 XPU 占 AI 營收（下一列）")
qrow(10, "  XPU ÷ AI", {"G": "=G9/G7"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="Q3 FY26 73%（法說，二手）")
qrow(11, "非 AI 半導體（＝半導體 − AI）", {c: f"={c}5-{c}7" for c in "BCDEF"}, tag="Derived")
qrow(12, "non-GAAP 毛利", {"B": 50.245, "C": 14.868, "D": 17.109, "E": 22.191, "F": "=F4*q4_gm"}, tag="Verified／Interested-party（Q4）")
qrow(13, "毛利率", {c: f"={c}12/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 78.6%；Q4 FY26 指引約 73%")
qrow(14, "non-GAAP 營業費用", {"B": 8.248, "C": 2.042, "D": 2.181, "E": 2.096, "F": "=F12-F15"}, tag="Verified／Derived（Q4＝毛利 − 營業利益）")
qrow(15, "non-GAAP 營業利益", {"B": 41.997, "C": 12.826, "D": 14.928, "E": 20.095, "F": "=F4*q4_oim"}, tag="Verified／Interested-party（Q4）")
qrow(16, "營業利益率", {c: f"={c}15/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(17, "non-GAAP 利息費用", {"B": 3.072, "C": 0.746, "D": 0.745, "E": 0.703, "F": "=q4_int"}, tag="Verified／Derived（Q4）")
qrow(18, "non-GAAP 其他收益", {"B": 0.294, "C": 0.118, "D": 0.118, "E": 0.098, "F": "=q4_oth"}, tag="Verified／Derived（Q4）")
qrow(19, "稅前利益", {c: f"={c}15-{c}17+{c}18" for c in "BCDEF"}, tag="Derived")
qrow(20, "non-GAAP 所得稅", {"B": 5.491, "C": 2.013, "D": 2.227, "E": 3.118, "F": "=F19*q4_tx"}, tag="Verified／Derived（Q4）")
qrow(21, "有效稅率", {c: f"={c}20/{c}19" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(22, "non-GAAP 淨利", {c: f"={c}19-{c}20" for c in "BCDEF"}, tag="Derived")
qrow(23, "  公司揭露 non-GAAP 淨利", {"B": 33.728, "C": 10.185, "D": 12.074, "E": 16.372}, ann="", tag="Verified")
qrow(24, "non-GAAP 稀釋股數（十億股）", {"B": 4.943, "C": 4.957, "D": 4.940, "E": 4.937, "F": "=q4_sh"}, ann="=AVERAGE(C24:F24)", tag="Verified／Derived（Q4）")
qrow(25, "non-GAAP EPS", {c: f"={c}22/{c}24" for c in "BCDEF"}, fmt="0.000", ann="=G22/G24", tag="Derived", s1="淨利 ÷ 稀釋股數", s2="—")
qrow(26, "  公司揭露 non-GAAP EPS", {"B": 6.82, "C": 2.05, "D": 2.44, "E": 3.32}, fmt="0.00", ann="", tag="Verified")
qrow(27, "檢查：淨利 − 公司揭露（應約為 0；四捨五入）", {c: f"={c}22-{c}23" for c in "BCDE"}, fmt="0.000", ann="", tag="—", s1="—", s2="—")
qrow(28, "檢查：Q1＋Q2＋Q3 − 九個月合計（營收、毛利、營業利益、淨利；應為 0）", {"G": "=(C4+D4+E4-71.089)+(C12+D12+E12-54.168)+(C15+D15+E15-47.849)+(C22+D22+E22-38.631)"}, fmt="0.000", ann="", tag="—", s1=Q3PR + "（九個月欄）", s2="—")
Q.cell(30, 1, "FY2026 錨點與共識、公司展望的距離（參考；校準未使用共識）").font = BOLD
Q.cell(31, 1, "FY2026 營收 ÷ 共識 − 1"); Q["B31"] = "=G4/cons_rev26-1"
Q.cell(32, 1, "FY2026 EPS ÷ 共識 − 1"); Q["B32"] = "=G25/cons_eps26-1"
Q.cell(33, 1, "FY2026 AI 營收 ÷ 公司展望 58 − 1"); Q["B33"] = "=G7/co_ai26-1"
Q.cell(34, 1, "FY2026 營收年增"); Q["B34"] = "=G4/B4-1"
Q.cell(35, 1, "FY2026 營業費用年增"); Q["B35"] = "=G14/B14-1"
Q.cell(36, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B36"] = "=B35/B34"
for rr in range(31, 37): Q[f"B{rr}"].number_format = "0.0%" if rr < 36 else "0.00"
for nm, rr in (("rev", 4), ("semi", 5), ("sw", 6), ("ai", 7), ("net", 8), ("xpu", 9), ("nai", 11), ("gp", 12), ("opx", 14), ("oi", 15), ("int", 17), ("oth", 18), ("tax", 20), ("ni", 22), ("sh", 24), ("eps", 25)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 90; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)          # 有由下而上推導的年度（FY＝CY）；FY2026 用來校準
def uexp(k, y, s):
    if y <= 2027: return f"u_{k}_{y % 100}"
    if y == 2028: return f"u_{k}_28*um28_{s}"
    return f"u_{k}_28*(1+g29_{k})*um29_{s}"
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
DEFP = dict(scn="A", uf="1", sht="0", sho="0", pfc="1", fh="f_hbm", mh="m_hbm", nf="1", gs27="g_sw27", gs="g_sw", opxa="0", flat29=False, wta=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    _wx = lambda y: (f"*{P['wta'][y]}" if P["wta"] and y in P["wta"] else "")   # Wattonomics 錨定（Sens 案例；預設 None＝不乘）
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    yu = lambda y: 2028 if (P["flat29"] and y == 2029) else y
    sec_("A. XPU（由下而上；D 欄＝FY2026 校準年，E–G＝FY2027–FY2029；FY(t)＝CY(t)）")
    for k in XP:
        row("u_" + k, f"{PNAME[k]} 市場顆數", "百萬顆", {c_(y): "=" + uexp(k, yu(y), s) for y in BU}, fmt="0.00")
    shx = {"tpu": P["sht"], "oth": P["sho"]}
    for k in XP:
        row("s_" + k, f"Broadcom 份額：{PNAME[k]}", "%", {c_(y): (f"=sh_{k}_{min(y, 2028) % 100}" if y == 2026 else f"=MIN(1,MAX(0,sh_{k}_{min(y, 2028) % 100}+{shx[k]}))") for y in BU}, fmt="0.0%")
    for k in XP:
        row("b_" + k, f"Broadcom 顆數：{PNAME[k]}", "百萬顆", {c_(y): f"={c_(y)}{R['u_' + k]}*{c_(y)}{R['s_' + k]}" + ("" if y == 2026 else f"*{P['uf']}") + _wx(y) for y in BU}, fmt="0.00")
    for k in XP:
        row("g_" + k, f"每顆 HBM：{PNAME[k]}", "GB", {c_(y): f"=gb_{k}_{min(y, 2028) % 100}" for y in BU}, fmt="0")
    row("hpx", "HBM 單價", "US$/GB", {c_(y): f"=hbm_{s}_{y % 100}" for y in BU}, fmt="0.00")
    for k in XP:
        row("c_" + k, f"每顆核心售價：{PNAME[k]}", "US$K", {c_(y): f"=core_{k}" + ("" if y == 2026 else f"*{P['pfc']}") for y in BU}, fmt="0.00")
    for k in XP:
        row("h_" + k, f"每顆 HBM 金額（全部）：{PNAME[k]}", "US$K", {c_(y): f"={c_(y)}{R['g_' + k]}*{c_(y)}{R['hpx']}/1000" for y in BU}, fmt="0.00")
    for k in XP:
        row("a_" + k, f"每顆 Broadcom 售價（核心＋代購 HBM）：{PNAME[k]}", "US$K", {c_(y): f"={c_(y)}{R['c_' + k]}+{P['fh']}*{c_(y)}{R['h_' + k]}" for y in BU}, fmt="0.00",
            note="核心售價 ＋ 代購比例 × HBM 金額（未乘 k_x）" if k == "tpu" else "")
    row("xraw", "XPU 營收（未校準）", "US$B", {c_(y): "=" + "+".join(f"{c_(y)}{R['b_' + k]}*{c_(y)}{R['a_' + k]}" for k in XP) for y in BU})
    row("kx", "XPU 校準係數 k_x ＝ FY2026 錨點 XPU ÷ 未校準", "x", {"D": f"=a26_xpu/D{r - 1}"}, fmt="0.000", note="Derived：以已發生的實績（Q1–Q3＋Q4 指引）校準；FY2027 起沿用。顆數層、份額與核心單價無法分開識別，k_x 是合併係數")
    row("xpu", "XPU 營收", "US$B", {c_(y): f"={c_(y)}{R['xraw']}*$D${R['kx']}" for y in BU})
    row("xh", "  其中代購 HBM", "US$B", {c_(y): f"=$D${R['kx']}*{P['fh']}*(" + "+".join(f"{c_(y)}{R['b_' + k]}*{c_(y)}{R['h_' + k]}" for k in XP) + ")" for y in BU})
    row("xc", "  其中核心（設計＋邏輯晶片＋封裝）", "US$B", {c_(y): f"={c_(y)}{R['xpu']}-{c_(y)}{R['xh']}" for y in BU})
    row("xu", "Broadcom XPU 顆數（未乘 k_x）", "百萬顆", {c_(y): f"={c_(y)}{R['b_tpu']}+{c_(y)}{R['b_oth']}" for y in BU}, fmt="0.00")
    sec_("B. AI 網通、非 AI 半導體、軟體")
    row("nraw", "AI 網通（markets/network 歸屬 × 情境顆數係數）", "US$B", {c_(y): f"=net_{yu(y) % 100}*{umx(yu(y), s)}" + ("" if y == 2026 else f"*{P['nf']}") + _wx(y) for y in BU})
    row("kn", "網通校準係數 k_n ＝ 實際推估 ÷ 歸屬（FY2026）", "x", {"D": f"=a26_net/D{r - 1}"}, fmt="0.000", note="Derived：markets/network 的 Broadcom 歸屬比例為 Assumed，2026 較實際推估低；FY2027 起沿用")
    row("net", "AI 網通營收", "US$B", {c_(y): f"={c_(y)}{R['nraw']}*$D${R['kn']}" for y in BU})
    row("ai", "AI 半導體營收", "US$B", {"C": "=a25_ai", **{c_(y): f"={c_(y)}{R['xpu']}+{c_(y)}{R['net']}" for y in BU}})
    row("xs", "XPU ÷ AI", "%", {c_(y): f"={c_(y)}{R['xpu']}/{c_(y)}{R['ai']}" for y in BU}, fmt="0.0%")
    row("nai", "非 AI 半導體", "US$B", {"C": "=a25_nai", "D": "=a26_nai", **{c_(y): f"={c_(y - 1)}{r}*(1+g_nai)" for y in BU[1:]}})
    row("sw", "基礎架構軟體", "US$B", {"C": "=a25_sw", "D": "=a26_sw", "E": f"=D{r}*(1+{P['gs27']})", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gs']})" for y in BU[2:]}})
    sec_("C. 報告值（FY2025A、FY2026 為實績與錨點；FY2027 起為由下而上）")
    row("rev", "營收", "US$B", {"C": "=a25_rev", "D": "=a26_rev", **{c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}+{c_(y)}{R['sw']}" for y in BU[1:]}})
    row("bur", "  由下而上營收（FY2026 應等於錨點）", "US$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}+{c_(y)}{R['sw']}" for y in BU})
    row("semi", "半導體解決方案營收（AI 半導體＋非 AI 半導體）", "US$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in FYS}, note="官方分部 Semiconductor solutions")
    row("segchk", "檢查：分部合計 − 合併營收（FY2025 實績、FY2026 錨點起）", "US$B", {c_(y): f"={c_(y)}{R['semi']}+{c_(y)}{R['sw']}-{c_(y)}{R['rev']}" for y in FYS}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("gpo", "毛利：網通＋非 AI＋軟體＋代購 HBM", "US$B", {c_(y): f"=gm_net*{c_(y)}{R['net']}+gm_nai*{c_(y)}{R['nai']}+gm_sw*{c_(y)}{R['sw']}+{P['mh']}*{c_(y)}{R['xh']}" for y in BU})
    row("mc", "XPU 核心毛利率 m_core ＝（錨點毛利 − 其他毛利）÷ 核心營收（FY2026）", "x", {"D": f"=(a26_gp-D{r - 1})/D{R['xc']}"}, fmt="0.000", note="Derived：讓 FY2026 毛利率等於錨點；FY2027 起的毛利率變化來自營收組合（代購 HBM 隨 HBM 漲價與容量上升）")
    row("gp", "毛利", "US$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"={c_(y)}{R['gpo']}+$D${R['mc']}*{c_(y)}{R['xc']}" for y in BU[1:]}})
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("hsh", "代購 HBM ÷ 營收", "%", {c_(y): f"={c_(y)}{R['xh']}/{c_(y)}{R['rev']}" for y in BU}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(g_opx_min,phi*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.00")
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.00")
    row("oi", "營業利益", "US$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("int", "利息費用", "US$B", {"C": "=a25_int", "D": "=a26_int", **{c_(y): "=int_f" for y in BU[1:]}}, fmt="0.00")
    row("oth", "其他收益", "US$B", {"C": "=a25_oth", "D": "=a26_oth", **{c_(y): "=oth_f" for y in BU[1:]}}, fmt="0.00")
    row("etr", "有效稅率", "%", {"C": "=a25_tax/(a25_oi-a25_int+a25_oth)", "D": "=a26_tax/(a26_oi-a26_int+a26_oth)", **{c_(y): "=tx_f" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "淨利（non-GAAP）", "US$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}-{c_(y)}{R['int']}+{c_(y)}{R['oth']})*(1-{c_(y)}{R['etr']})" for y in BU[1:]}})
    row("sh", "稀釋股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.000")
    row("eps", "EPS（non-GAAP）", "US$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh']}" for y in FYS}, fmt="0.00")
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Broadcom 損益（US$B；non-GAAP）。FY2025A 實績、FY2026 錨點（Q1–Q3 實績＋Q4 指引），FY2027 起為由下而上。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始；HBM 價格 2027 三情境相同）"
PB = {}; r0 = 70
SCN_T = {"A": "情境 A：原始顆數（顆數係數 1.0；v0.12 起本庫基準＝情境 B）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 半導體", "US$B", "#,##0.0"), ("xpu", "    XPU", "US$B", "#,##0.0"),
                                ("net", "    AI 網通", "US$B", "#,##0.0"), ("sw", "  基礎架構軟體", "US$B", "#,##0.0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利（non-GAAP）", "US$B", "#,##0.0"),
                                ("eps", "EPS（non-GAAP）", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("xpu", "net") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 由下而上營收 − 錨點；FY2026 由下而上 XPU − 錨點 XPU（k_x、k_n 校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['bur']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：XPU 差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['xpu']}-a26_xpu"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 70, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：ASIC 每顆規格（權威）、金額份額（accel_share）、校準值 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — ASIC 每顆規格（本專案權威，規則第 13 條）與加速器金額份額。顆數：百萬顆（日曆年；與 HBM、台積電、網通、NVIDIA 模型同一份）；每顆價：US$K"
A_ = PB["A"]
CYS = (2026, 2027, 2028, 2029); cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026（FY2026）", "CY2027（FY2027）", "CY2028（FY2028）", "CY2029（FY2029）", "說明"])
ur = 4
UN.cell(ur, 1, "1. ASIC 每顆規格（情境 A；Broadcom 營收口徑＝核心售價＋代購 HBM，乘 k_x）").font = BOLD; ur += 1
SPEC = {}
for k in XP:
    for key, lab, unit, f, fmt, note in (
        ("s", f"Broadcom 份額：{PNAME[k]}", "%", lambda y: f"=PL!{FC[y]}{A_['s_' + k]}", "0.0%", "In 頁 B 節（Assumed）"),
        ("b", f"Broadcom 顆數：{PNAME[k]}", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['b_' + k]}", "0.00", "市場層顆數 × 份額"),
        ("g", f"每顆 HBM：{PNAME[k]}", "GB", lambda y: f"=PL!{FC[y]}{A_['g_' + k]}", "0", "markets/memory UNITS"),
        ("h", f"每顆 HBM 金額：{PNAME[k]}", "US$K", lambda y: f"=PL!{FC[y]}{A_['h_' + k]}", "0.00", "容量 × HBM 單價（情境 A）"),
        ("p", f"每顆 Broadcom 營收（校準後）：{PNAME[k]}", "US$K", lambda y: f"=PL!{FC[y]}{A_['a_' + k]}*PL!$D${A_['kx']}", "0.00", "（核心＋代購比例 × HBM）× k_x"),
        ("q", f"每顆總價（核心 × k_x ＋ 全部 HBM；金額份額用）：{PNAME[k]}", "US$K", lambda y: f"=PL!{FC[y]}{A_['c_' + k]}*PL!$D${A_['kx']}+PL!{FC[y]}{A_['h_' + k]}", "0.00", "與 NVIDIA 模組價同口徑（含全部 HBM，不論誰採購）")):
        UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
        for y in CYS:
            UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, note); SPEC[(key, k)] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "2. 校準值（只用 FY2026 已發生的實績與當季指引；不用共識、不用公司對 FY2027–28 的展望）").font = BOLD; ur += 1
CAL = {}
for key, lab, f, fmt, note in (("kx", "XPU 校準係數 k_x", f"=PL!D{A_['kx']}", "0.000", "FY2026 錨點 XPU ÷ 未校準由下而上；顆數、份額、核心單價的合併係數"),
                               ("kn", "網通校準係數 k_n", f"=PL!D{A_['kn']}", "0.000", "rep_avgo26 ÷ markets/network 歸屬"),
                               ("mc", "XPU 核心毛利率 m_core", f"=PL!D{A_['mc']}", "0.0%", "讓 FY2026 毛利率等於錨點的殘差（Derived）"),
                               ("x26", "FY2026 錨點 XPU", "=a26_xpu", "0.0", "AI 57.6 − 網通 17.6"),
                               ("xr26", "FY2026 未校準 XPU", f"=PL!D{A_['xraw']}", "0.0", "")):
    UN.cell(ur, 1, lab); UN[f"C{ur}"] = f; UN[f"C{ur}"].number_format = fmt; UN.cell(ur, 7, note); CAL[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "3. 加速器金額份額（accel_share；各平台整顆售價 × 顆數；整顆售價讀 NVIDIA 模型〔Siliconomics M6〕，CF13 (a)）").font = BOLD; ur += 1
# CF13 (a)（Andy 2026-11-05 15:24）：市場金額份額一律用整顆售價 pa_*（與 NVIDIA、AMD 模型同口徑）；本頁第 1 節「核心 × k_x＋HBM」只用於 Broadcom 自身營收
PRICE = {"bw": "pm_bw", "rb": "pm_rb", "ru": "pm_ru", "hp": "pm_hp", "amd": "pa_amd", "trn": "pa_trn", "tpu": "pa_tpu", "oth": "pa_oth"}
ycols = {y: cyc(y) for y in CYS}
SHR = {}
for s in SC:
    ttl = {"A": "情境 A：基準", "B": "情境 B：2028 顆數 ×0.8、2029 ×0.85", "C": "情境 C：2028–29 顆數 ×0.6"}[s]
    units = {(p, y): uexp(p, y, s) for p in PLAT for y in CYS}
    ur, SHR[s] = AS.write_share_block(UN, ur, ttl, CYS, ycols, units, PRICE, PLAT, PNAME)
    UN.cell(ur - 1, 1, "  ASIC 合計（TPU＋Trainium＋其他）金額份額").font = BOLD
    for y in CYS:
        c = cyc(y)
        UN[f"{c}{ur - 1}"] = f"={c}{SHR[s][('ds', 'tpu')]}+{c}{SHR[s][('ds', 'trn')]}+{c}{SHR[s][('ds', 'oth')]}"; UN[f"{c}{ur - 1}"].number_format = "0.0%"
    SHR[s]["asic_d"] = ur - 1
    UN.cell(ur, 1, "  ASIC 合計顆數份額").font = BOLD
    for y in CYS:
        c = cyc(y)
        UN[f"{c}{ur}"] = f"={c}{SHR[s][('us', 'tpu')]}+{c}{SHR[s][('us', 'trn')]}+{c}{SHR[s][('us', 'oth')]}"; UN[f"{c}{ur}"].number_format = "0.0%"
    SHR[s]["asic_u"] = ur
    UN.cell(ur + 1, 1, "  Broadcom ÷ ASIC 金額（整顆售價口徑；CF13 (a)）").font = BOLD
    for y in CYS:
        c = cyc(y)
        UN[f"{c}{ur + 1}"] = (f"=(PL!{FC[y]}{PB[s]['b_tpu']}*pa_tpu+PL!{FC[y]}{PB[s]['b_oth']}*pa_oth)"
                              f"/({c}{SHR[s][('d', 'tpu')]}+{c}{SHR[s][('d', 'trn')]}+{c}{SHR[s][('d', 'oth')]})")
        UN[f"{c}{ur + 1}"].number_format = "0.0%"
    SHR[s]["avgo_d"] = ur + 1
    UN.cell(ur + 2, 1, "  Broadcom ÷ ASIC 顆數").font = BOLD
    for y in CYS:
        c = cyc(y)
        UN[f"{c}{ur + 2}"] = f"=PL!{FC[y]}{PB[s]['xu']}/({c}{SHR[s][('u', 'tpu')]}+{c}{SHR[s][('u', 'trn')]}+{c}{SHR[s][('u', 'oth')]})"
        UN[f"{c}{ur + 2}"].number_format = "0.0%"
    SHR[s]["avgo_u"] = ur + 2
    ur += 4
UN.cell(ur, 1, "口徑：每顆價為整顆售價（模組層級；NVIDIA＝GPU 模組價，AMD 與 ASIC＝Siliconomics M6 pa_*，經 NVIDIA 模型），不含系統加價與網通；CF13 (a)（Andy 2026-11-05）起與 NVIDIA、AMD 模型同口徑。本頁第 1 節「核心 × k_x＋全部 HBM」只用於 Broadcom 營收與「Broadcom ÷ ASIC 金額」列。").font = BOLD
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 17
UN.column_dimensions["G"].width = 70

# ---------------- Val：P/E ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元；non-GAAP EPS）。結論用「現價 ÷ FY2027、FY2028 EPS（模型與共識）」。十年 P/E 規則值本版不適用：GAAP 落後 P/E 受無形資產攤銷扭曲（2016 年為負、2018 年 6.5x、2024 年 137x），前瞻 P/E 歷史待 Bloomberg（10-27 抓取單）"
V["A2"] = "循環位置：共用規則（tools/cycle_position.py classify；AVGO 的 FY(t) 對照日曆年 t，不平移），見頂年度讀設備市場模型；本版只記錄判定，倍數待前瞻 P/E 歷史"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-11-05 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "0.00"
V.cell(7, 1, "現價 ÷ FY2026 EPS（錨點）")
V.cell(8, 1, "現價 ÷ FY2027 EPS（模型）")
V.cell(9, 1, "現價 ÷ FY2028 EPS（模型）")
V.cell(10, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((7, 2026), (8, 2027), (9, 2028), (10, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(11, 1, "現價 ÷ 共識 FY2027 EPS"); V["B11"] = "=px/cons_eps27"; V["B11"].number_format = "0.0"
V.cell(12, 1, "近 12 月 non-GAAP P/E（Q4 FY25＋Q1–Q3 FY26 EPS）"); V["B12"] = "=px/(eps_q4_25+Q!C26+Q!D26+Q!E26)"; V["B12"].number_format = "0.0"
V.cell(13, 1, "近 12 月 GAAP P/E（StockAnalysis）"); V["B13"] = "=pe_ttm_gaap"; V["B13"].number_format = "0.0"
V.cell(14, 1, "前瞻 P/E（StockAnalysis；口徑未說明）"); V["B14"] = "=pe_fwd_sa"; V["B14"].number_format = "0.0"
V["E8"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"; V["E12"] = "Andy 先驗 B9a 答「前瞻 35x」：接近近 12 月 non-GAAP 倍數，而非前瞻倍數（見 Recon D 節）"
V.cell(16, 1, "2. 循環位置（共用規則，FY(t) 對照 CY t）——只記錄判定；倍數待前瞻 P/E 歷史").font = BOLD
hdr(V, 17, ["情境", "設備市場見頂年度（CY）", "FY2027 判定", "FY2028 判定", "FY2029 判定", "規則"])
for i, s in enumerate(SC, 18):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}")
    for j, y in enumerate((2027, 2028, 2029)):
        V.cell(i, 3 + j, "=" + CP.excel_rule(str(y), f"pk_{s}"))
    V.cell(i, 6, "tools/cycle_position.py classify（不平移）")
V.cell(22, 1, "3. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（情境 A 的 FY2028 EPS；約 2027-10 的價格看下一年 EPS）").font = BOLD
hdr(V, 23, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：Andy 合理 P/E", "對照：現價 ÷ 共識 FY2027 EPS"])
vr = 24
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_fair_andy").number_format = "0.0"; V.cell(vr, 5, "=px/cons_eps27").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數愈低，現價愈便宜；倍數是否合理，待前瞻 P/E 十年歷史（Bloomberg）。")
vr += 2
V.cell(vr, 1, "4. 對照：以 Andy 合理 P/E 25x 計算的隱含價（只作對照，非規則值）").font = BOLD; vr += 1
hdr(V, vr, ["情境", "FY2027 EPS", "× 25x", "相對現價", "FY2028 EPS", "× 25x", "相對現價"]); vr += 1
IMP = {}
for s in SC:
    V.cell(vr, 1, s)
    V.cell(vr, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 3, f"=B{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 4, f"=C{vr}/px-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 6, f"=E{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
    IMP[s] = vr; vr += 1
vr += 1
V.cell(vr, 1, "5. 十年 GAAP P/E（Macrotrends，10 月 31 日；不適用——GAAP EPS 受攤銷扭曲，只作紀錄）").font = BOLD; vr += 1
hdr(V, vr, ["年底", "P/E（GAAP 落後）", "說明"]); vr += 1
for y, v in ((2016, None), (2017, 49.57), (2018, 6.50), (2019, 38.87), (2020, 49.77), (2021, 32.75), (2022, 16.88), (2023, 24.95), (2024, 137.36), (2025, None)):
    V.cell(vr, 1, str(y))
    if v is not None: V.cell(vr, 2, v).font = BLUE; V.cell(vr, 2).number_format = "0.00"
    V.cell(vr, 3, "GAAP EPS 為負，P/E 無意義" if y == 2016 else ("頁面資料只到 2025-08：未找到" if y == 2025 else "")); vr += 1
V.cell(vr, 1, "來源：macrotrends.net/stocks/charts/AVGO/broadcom/pe-ratio（Interested-party，資料商）；第二來源：未找到。前瞻 P/E 歷史待 Bloomberg 抓取單（10-27）").font = Font(italic=True)
V.column_dimensions["A"].width = 58
for c in "BCDEFG": V.column_dimensions[c].width = 18
V.column_dimensions["E"].width = 40

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；k_x、k_n、m_core 在每個案例內依 FY2026 實績重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("XPU 顆數係數 ×0.9", "FY2027 起 Broadcom XPU 顆數 ×0.9", dict(uf="uk_lo")), ("XPU 顆數係數 ×1.1", "×1.1", dict(uf="uk_hi")),
         ("TPU 份額 −15pt", "FY2027 起 Broadcom 占 TPU 顆數 −15 個百分點（MediaTek 版放量）", dict(sht="(-0.15)")), ("TPU 份額 +10pt", "+10 個百分點（上限 100%）", dict(sht="0.10")),
         ("其他 ASIC 份額 −15pt", "FY2027 起 −15 個百分點", dict(sho="(-0.15)")), ("其他 ASIC 份額 +15pt", "+15 個百分點", dict(sho="0.15")),
         ("核心售價 ×0.8", "FY2027 起 XPU 核心售價 ×0.8", dict(pfc="pfc_lo")), ("核心售價 ×1.2", "×1.2", dict(pfc="pfc_hi")),
         ("HBM 代購比例 0.3", "代購 30%（k_x、m_core 重算）", dict(fh="f_hbm_lo")), ("HBM 代購比例 1.0", "全部代購", dict(fh="f_hbm_hi")),
         ("代購 HBM 毛利率 0%", "零加成", dict(mh="m_hbm_lo")), ("代購 HBM 毛利率 20%", "20%", dict(mh="m_hbm_hi")),
         ("AI 網通 ×0.8", "FY2027 起網通 ×0.8", dict(nf="0.8")), ("AI 網通 ×1.2", "×1.2", dict(nf="1.2")),
         ("軟體成長 低", "FY2027 +8%、之後 +4%", dict(gs27="g_sw27_lo", gs="g_sw_lo")), ("軟體成長 高", "FY2027 +18%、之後 +12%", dict(gs27="g_sw27_hi", gs="g_sw_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ai", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ai", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
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
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85；記憶體 B 路徑", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6；記憶體 C 路徑", PB["C"], "PL!")
assert nn + 3 < 40
SENS_N = len(cases)
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。k_x 隨案例重算，所以 HBM 代購比例只改變營收組合與毛利率，不改 FY2026 水準。")
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司展望、共識、市場層。只作事後對照，不為對上數字而調參數。公司 FY2027、FY2028 AI 營收展望（115、230）為 Interested-party，不用於校準"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False, gap=True):
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
RC.cell(rr_, 1, "A. AI 營收 vs 公司展望（Interested-party；校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 AI 營收 vs 公司展望 58", "=co_ai26", pl("ai", 2026), "公司（法說）", "FY2026 為錨點（Q1–Q3 實績＋Q4 指引 57.6）")
rec("FY2027 AI 營收 vs 公司展望 115", "=co_ai27", pl("ai", 2027), "公司（法說）", "")
rec("FY2027 AI 營收 vs JPMorgan 150", "=jpm_ai27", pl("ai", 2027), "JPMorgan（KuCoin 轉述）", "")
rec("FY2028 AI 營收 vs 公司展望 230", "=co_ai28", pl("ai", 2028), "公司（法說）", "差距主要在顆數：公司說法含 Anthropic 2028 年再加約 10 GW、OpenAI 5 GW 以上")
RC.cell(rr_, 1, "B. What must be true：要達到公司展望，Broadcom XPU 顆數需要是模型的幾倍（網通、核心售價、HBM 不變）").font = BOLD; rr_ += 1
rec("FY2027：所需 XPU 顆數倍數", None, lambda s: f"=(co_ai27-PL!E{PB[s]['net']})/PL!E{PB[s]['xpu']}", "—", "倍數 >1＝公司展望需要比市場層更多的 Broadcom 顆數", fmt="0.00")
rec("FY2028：所需 XPU 顆數倍數", None, lambda s: f"=(co_ai28-PL!F{PB[s]['net']})/PL!F{PB[s]['xpu']}", "—", "", fmt="0.00")
rec("FY2028：所需 Broadcom XPU 顆數（百萬顆）vs 模型", f"=PL!F{PB['A']['xu']}", lambda s: f"=(co_ai28-PL!F{PB[s]['net']})/PL!F{PB[s]['xpu']}*PL!F{PB[s]['xu']}", "—", "對照值＝模型 A 的 Broadcom XPU 顆數", fmt="0.00", gap=False)
RC.cell(rr_, 1, "C. 共識（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "StockAnalysis", "錨點")
rec("FY2026 EPS vs 共識", "=cons_eps26", pl("eps", 2026), "StockAnalysis", "錨點（Q4 依指引與 In 頁假設）", fmt="0.00")
rec("FY2027 營收 vs 共識", "=cons_rev27", pl("rev", 2027), "StockAnalysis", "")
rec("FY2027 EPS vs 共識", "=cons_eps27", pl("eps", 2027), "StockAnalysis", "", fmt="0.00")
rec("FY2027 共識隱含 AI 營收（共識營收 − 模型非 AI − 模型軟體）", None, lambda s: f"=cons_rev27-PL!E{PB[s]['nai']}-PL!E{PB[s]['sw']}", "推導", "與公司展望 115 比較", gap=False)
RC.cell(rr_, 1, "D. 份額與顆數（先驗比對的機制檢查）").font = BOLD; rr_ += 1
rec("2027 年 ASIC 金額份額（模組層級）vs 先驗比對推導 14–19% 中點", 0.165, lambda s: f"=Units!D{SHR[s]['asic_d']}", "先驗比對_AVGO（Drive）", "", fmt="0.0%", pt=True)
rec("2027 年 ASIC 顆數份額 vs JPMorgan 53.6%", "=jpm_u27", lambda s: f"=Units!D{SHR[s]['asic_u']}", "JPMorgan（二手）；市場層同一來源", "非獨立檢驗（市場層顆數即取自 JPMorgan）", fmt="0.0%", pt=True)
rec("2027 年 Broadcom ÷ ASIC 顆數 vs Counterpoint 60%", "=cp_sh27", lambda s: f"=Units!D{SHR[s]['avgo_u']}", "Counterpoint（二手；口徑未說明）", "", fmt="0.0%", pt=True)
rec("2027 年 Broadcom ÷ ASIC 金額 vs Counterpoint 60%", "=cp_sh27", lambda s: f"=Units!D{SHR[s]['avgo_d']}", "同上", "JPMorgan：高階 ASIC 80–85%", fmt="0.0%", pt=True)
rec("2027 年 NVIDIA 金額份額", None, lambda s: f"=Units!D{SHR[s][('ds', 'nv')]}", "accel_share", "NVIDIA 模型同一套價格", fmt="0.0%")
rec("每顆價差：NVIDIA Rubin 模組 ÷ TPU 總價（CY2027）", None, lambda s: f"=pm_rb/Units!D{SPEC[('q', 'tpu')]}", "—", "顆數份額約為金額份額數倍的原因", fmt="0.0")
rec("近 12 月 non-GAAP P/E vs Andy 先驗「前瞻 35x」", 35, lambda s: "=Val!B12", "先驗_Andy_AVGO B9a（Drive）", "35x 接近近 12 月倍數；前瞻（FY2027 EPS）倍數見 Val 頁", fmt="0.0")
RC.cell(rr_, 1, "E. 校準與網通").font = BOLD; rr_ += 1
rec("k_x（XPU 校準係數）", None, lambda s: f"=PL!D{PB[s]['kx']}", "FY2026 錨點 ÷ 未校準", "接近 1＝市場層顆數 × 份額 × 假設單價與實績一致", fmt="0.000")
rec("k_n（網通校準係數）", None, lambda s: f"=PL!D{PB[s]['kn']}", "rep_avgo26 ÷ markets/network 歸屬", "markets/network 的 Broadcom 歸屬比例偏低（Assumed）", fmt="0.000")
rec("m_core（XPU 核心毛利率，Derived）", None, lambda s: f"=PL!D{PB[s]['mc']}", "FY2026 毛利殘差", "高＝設計與邏輯晶片加價高；代購 HBM 稀釋整體", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "公司 Q4 FY26 指引約 73%，續降", fmt="0.0%", pt=True)
rec("FY2027 毛利率 vs 共識隱含（約 72%：EPS 19.40 × 股數 ÷（1−稅）＋利息＋營業費用，÷ 共識營收）", "=(cons_eps27*a26_sh/(1-tx_f)+int_f-oth_f+PL!E" + str(PB['A']['opx']) + ")/cons_rev27", pl("gm", 2027), "推導（共識無毛利率）", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "F. 客戶融資曝險（資訊；不進基準 EPS）").font = BOLD; rr_ += 1
rec("Anthropic 放款＋OpenAI 融資安排（標題值）vs FY2027 AI 營收（模型）", "=fin_anth+fin_oai", pl("ai", 2027), "top1markets 引 Reuters、WSJ、Bloomberg", "媒體：約 92B ≈ 公司 FY2027 AI 目標的 80%；非資產負債表曝險，部分在銀行團；另有一家客戶機架租賃擔保約 29B（季報），可能與上列重疊，不加總", gap=False)
rec("兩家合計 ÷ FY2027 AI 營收（模型）", None, lambda s: f"=(fin_anth+fin_oai)/PL!E{PB[s]['ai']}", "—", "", fmt="0%")
rec("兩家合計 ÷ FY2027 淨利（模型）", None, lambda s: f"=(fin_anth+fin_oai)/PL!E{PB[s]['ni']}", "—", "若全數損失的量級（不含回收與銀行團分擔）", fmt="0.0")
RC.cell(rr_, 1, "G. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以 Broadcom AI 營收 ÷ 部署 GW 對照 Tokenomics IF_NonNV_Hold_TPU 等比例"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 80
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝Broadcom 8-K 新聞稿，已逐份抓取；法說逐字稿只經二手轉述）"; SO["A1"].font = BOLD
srcs = [("公司：FY2025 與 Q4 FY25 新聞稿（8-K）", EDG + "000173016825000116/avgo-11022025x8kxex99.htm"),
        ("公司：Q1 FY26 新聞稿（8-K）", EDG + "000173016826000011/avgo-02012026x8kxex99.htm"),
        ("公司：Q2 FY26 新聞稿（8-K）", EDG + "000173016826000051/avgo-05032026x8kxex99.htm"),
        ("公司：Q3 FY26 新聞稿（含 Q4 指引、九個月合計）", Q3PR),
        ("二手：Q3 FY26 法說逐字稿（XPU 73%、客戶、毛利率路徑、FY2027–28 展望）", "https://www.investing.com/news/transcripts/earnings-call-transcript-broadcom-tops-q3-2026-estimates-as-ai-sales-surge-93CH-4886849"),
        ("二手：BigGo Finance 法說摘要", "https://finance.biggo.com/news/US_AVGO_2026-09-02"),
        ("二手：StockAnalysis 共識、現價、前瞻 P/E", "https://stockanalysis.com/stocks/avgo/forecast/"),
        ("二手：GuruFocus 前瞻 P/E", "https://www.gurufocus.com/term/forward-pe-ratio/AVGO"),
        ("二手：valueinvesting.io 前瞻 P/E 49.62（口徑不明，不採）", "https://valueinvesting.io/AVGO/metric/forward-pe"),
        ("二手：top1markets 客戶融資與股價（2026-10-22）", "https://www.top1markets.com/news/broadcom-stock-ai-revenue-230-billion-financing-risk"),
        ("二手：JPMorgan（KuCoin 轉述）：ASIC 顆數、Broadcom AI 營收", "https://www.kucoin.com/news/flash/jpmorgan-ai-custom-chip-shipments-to-surpass-gpus-by-2027"),
        ("二手：Counterpoint（Yahoo 轉述）：Broadcom 2027 ASIC 份額約 60%", "https://finance.yahoo.com/news/broadcom-set-dominate-custom-ai-163116560.html"),
        ("二手：Macrotrends GAAP P/E", "https://www.macrotrends.net/stocks/charts/AVGO/broadcom/pe-ratio"),
        ("市場層：markets/memory/market.py（UNITS）；tsmc、skhynix、markets/network、nvidia 成品（本 repo，唯讀）", ""),
        ("待補（人工）：前瞻 P/E 十年歷史（Bloomberg，10-27）；各客戶營收；XPU 產品別毛利率；HBM 代購比例", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Broadcom（AVGO）收支與估值模型 v0.2.1（AI 半導體 Project：運算層第二個樣板，ASIC＋網通）　2026-10-25", True),
 ("版本：v0.4（2026-11-05：CF13 (a)〔Andy 2026-11-05 15:24〕——Units 頁加速器金額份額改用整顆售價 pa_*〔與 NVIDIA、AMD 模型同口徑〕；只影響份額展示，營收、EPS、目標價不變）。v0.1（初版）。ASIC 每顆規格（Broadcom 份額、核心售價、HBM 代購）在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）。", False),
 ("模型命題：Broadcom 的 AI 營收＝各 XPU 平台的顆數 × Broadcom 份額 × 每顆售價（核心＋代購 HBM）＋網通附掛；2027–28 年的增量集中在 Anthropic、OpenAI 兩家靠融資部署的客戶。股價約 17 倍 FY2027 EPS：這個折價是在定價客戶信用風險，還是低估了營收？", False),
 ("核心問題：以市場層顆數往前推，Broadcom 的 AI 營收與 EPS 能否支撐公司展望（FY2027 約 115B、FY2028 約 230B）；差距要靠多少額外顆數補上。", False), ("", False),
 ("會計年度：FY(t) 於 t 年 10 月底／11 月初結束，10/12 落在日曆年 t，FY(t) 直接對照 CY(t)，不做平移。預測期 FY2027–FY2029（三年）。口徑：公司 non-GAAP（排除股票薪酬、攤銷、重組）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("XPU：Google TPU 與「其他 ASIC」顆數（markets/memory UNITS × 情境係數）× Broadcom 份額（Assumed）×（核心售價＋代購比例 × 每顆 HBM 容量 × HBM 單價）× k_x（FY2026 實績校準）→ XPU 營收。", False),
 ("AI 網通：markets/network 的「Broadcom 網通（本模型歸屬）」× 情境顆數係數 × k_n（FY2026 實際推估校準）。非 AI 半導體、軟體：FY2026 錨點 × Assumed 成長率。", False),
 ("毛利率：網通、非 AI、軟體、代購 HBM 各用 Assumed 毛利率；XPU 核心毛利率 m_core 為 FY2026 殘差（Derived）。FY2027 起毛利率下滑來自代購 HBM 占比上升（HBM 2027 年 +80%、每顆容量上升）。", False),
 ("校準：FY2026 ＝ Q1–Q3 實績＋Q4 指引（營收 34.8、AI 21.7、營業利益率 66%）。不用共識、不用公司對 FY2027–28 的展望（只在 Recon 對帳）。", False),
 ("估值：結論用現價 ÷ FY2027、FY2028 EPS；十年 GAAP P/E 規則值不適用（攤銷扭曲），前瞻 P/E 歷史待 Bloomberg。循環位置（共用規則，不平移）只作紀錄：設備市場見頂年度 " + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + "。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文＝SEC EDGAR 8-K；法說內容只經二手轉述（標二手）。", False),
 ("已知限制：(1) k_x 是顆數、份額、核心單價的合併係數，無法分開識別；(2) Broadcom 各平台份額、核心售價、HBM 代購比例皆為 Assumed，FY2027 起的毛利率對代購比例高度敏感（Sens）；(3) 「其他 ASIC」含昇騰、Maia 等非 Broadcom 晶片，份額為粗估；(4) 客戶融資曝險不進基準 EPS（Recon F）；(5) 十年倍數待補；(6) R4 產業部署路徑未建（Recon 預留列）。", False)]
lines.append(("v0.1.1（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有儲存格）與 In 頁目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。本模型無規則值，目標價用十年 GAAP P/E 中位數 × 各情境 FY2027 non-GAAP EPS（偏誤方向為高估，見 Val 頁目標價方法）；其餘儲存格不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- 舊法：規則值情境加權（v0.1.1 的「0. 目標價與上檔空間」；v0.2 改名保留，不再叫目標價） ----------------
def add_target_price(V, imp, note, fmt, method, px_note):
    t0 = V.max_row + 2
    V.cell(t0, 1, "舊法：規則值情境加權（參考）").font = BOLD
    hdr(V, t0 + 1, ["項目", "數值", "", "", "說明"])
    V.cell(t0 + 2, 1, "規則值情境加權價（參考，非目標價）"); V[f"B{t0 + 2}"] = f"=tp_wA*B{t0 + 4}+tp_wB*B{t0 + 5}+tp_wC*B{t0 + 6}"; V[f"B{t0 + 2}"].number_format = fmt
    V[f"E{t0 + 2}"] = "權重 × 各情境隱含價"
    V.cell(t0 + 3, 1, "相對現價（參考）"); V[f"B{t0 + 3}"] = f"=B{t0 + 2}/px-1"; V[f"B{t0 + 3}"].number_format = "0.0%"; V[f"E{t0 + 3}"] = px_note
    for i, (s, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
        rr = t0 + 4 + i
        V.cell(rr, 1, lab); V[f"B{rr}"] = imp[s]; V[f"B{rr}"].number_format = fmt; V[f"E{rr}"] = note[s]
    rr = t0 + 7
    V.cell(rr, 1, "情境權重 A／B／C"); V[f"B{rr}"] = "=tp_wA"; V[f"C{rr}"] = "=tp_wB"; V[f"D{rr}"] = "=tp_wC"
    for c in "BCD": V[f"{c}{rr}"].number_format = "0%"
    V[f"E{rr}"] = f'="合計 "&TEXT(B{rr}+C{rr}+D{rr},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
    V.cell(t0 + 8, 1, "舊法說明"); V[f"B{t0 + 8}"] = method
    return t0
PE_ROWS = [rw for rw in range(1, V.max_row + 1) if V.cell(rw, 1).value in [str(y) for y in range(2016, 2026)]]
assert len(PE_ROWS) == 10 and PE_ROWS[-1] - PE_ROWS[0] == 9, PE_ROWS
PE_MED = f"MEDIAN($B${PE_ROWS[0]}:$B${PE_ROWS[-1]})"
TP0 = add_target_price(V, {s: f"={PE_MED}*PL!E{SUMR[(s, 'eps')]}" for s in SC},
    {s: f"情境 {s} FY2027 non-GAAP EPS（PL 頁）× 十年 GAAP P/E 中位數（第 5 節，有值的 8 年）" for s in SC}, "#,##0.00",
    "P/E：十年（2016–2025，10 月底會計年度年底）GAAP P/E 中位數（Macrotrends，第 5 節；2016 GAAP EPS 為負、2025 未找到，取有值的 8 年）× 各情境 FY2027 non-GAAP EPS（FY2027＝2026-11～2027-10）。"
    "本模型無規則值（前瞻 P/E 歷史待補），依工作單改用此備案。偏誤方向為高估：GAAP EPS 受併購無形資產攤銷壓低，GAAP P/E 高於 non-GAAP 口徑（2024 年併購 VMware 後約 137x），乘在 non-GAAP EPS 上偏高；"
    "FY2027 三情境 EPS 相同（情境差從 FY2028 起），故三情境隱含價相同；隱含價約為 FY2027 年底（2027-10）的價格、未折現。",
    "現價＝In 頁 px（2026-10-23 收盤）")

# ---------------- v0.2：三情境延伸到 FY2032、Val 第 0 節（DCF 與 Comparable 加權）、Sens 目標價參數 ----------------
XH, XR, XC = TP.extend_pl(PL, PB, FC, FYS, 2029, {"rev": "rev", "ni": "ni", "sh": "sh", "eps": "eps"}, "US$",
    "三情境延伸到 FY2032（v0.2；FY2025A–FY2029 引用上方推導區塊，FY2030 起為收斂法，非驅動因子推導；收斂年度為橘底）",
    "上游市場層（markets/*、台積電顆數）只到 CY2029＝FY2029；FY2030 起：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋tp_gd、B＝tp_g、C＝tp_g−tp_gd），"
    "淨利率由 FY2029 線性收斂到中段利潤率（A＝FY2029、B＝FY2029 與 C 目標的中點、C＝FY2029 × nm10_k；同口徑十年平均未找到），股數沿用 FY2029 年變動")
TPL, TPN = TP.add_tp_section(wb, V,
    eps={s: [f"PL!{XC[y]}{XR[(s, 'eps')]}" for y in range(2027, 2033)] for s in SC}, e_mon=10,
    fwd={s: [(f"PL!{XC[2028]}{XR[(s, 'eps')]}", 1)] for s in SC},
    fwd_note="會計年度 FY(t) 約為 t−1 年 11 月～t 年 10 月：約 12 個月後（2027-10 底）的未來 12 個月＝2027-11～2028-10＝FY2028 全年（1 × FY2028）",
    peers=[("NVIDIA", "NVDA", 15.02, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 16.92（2026-10-30 頁面，價格 204.64）→ 前瞻 EPS 12.09；P/E＝2026-11-04 基準 16.55 × Yahoo Finance 日線收盤 2026-11-05 181.6 ÷ 2026-11-04 收盤 200.13（共識 EPS 不變）", "Interested-party（資料商彙整賣方預估）"),
           ("AMD", "AMD", 36.12, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 45.09（2026-10-30 頁面，價格 504.8）→ 前瞻 EPS 11.2；P/E＝2026-11-04 基準 43.66 × Yahoo Finance 日線收盤 2026-11-05 404.4 ÷ 2026-11-04 收盤 488.78（共識 EPS 不變）", "Interested-party（資料商彙整賣方預估）"),
           ("Marvell", "MRVL", 23.66, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 42.09（2026-10-27 頁面，價格 237.6）→ 前瞻 EPS 5.645；P/E＝2026-11-04 基準 28.31 × Yahoo Finance 日線收盤 2026-11-05 133.55 ÷ 2026-11-04 收盤 159.8（共識 EPS 不變）", "Interested-party（資料商彙整賣方預估）"),
           ("Qualcomm", "QCOM", 16.93, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 20.03（2026-10-29 頁面，價格 185.41）→ 前瞻 EPS 9.257；P/E＝2026-11-04 基準 18.71 × Yahoo Finance 日線收盤 2026-11-05 156.75 ÷ 2026-11-04 收盤 173.19（共識 EPS 不變）", "Interested-party（資料商彙整賣方預估）"),
           ("台積電 ADR", "TSM", 16.58, "2026-11-05", "stockanalysis.com 統計頁 Forward PE 19.07（2026-10-30 頁面，價格 406.13）→ 前瞻 EPS 21.3；P/E＝2026-11-04 基準 18.58 × Yahoo Finance 日線收盤 2026-11-05 353.1 ÷ 2026-11-04 收盤 395.75（共識 EPS 不變）", "Interested-party（資料商彙整賣方預估）")],
    peer_note="券商常用的 AI 半導體大型股同業（GPU、客製 ASIC、手機 SoC、代工）；不含 A 股（Andy 11:53）。前瞻 P/E 口徑為資料商的未來 12 個月共識 EPS，各家 GAAP／non-GAAP 口徑不一",
    cons_tp="cons_tp", cons_note="stockanalysis.com：51 位分析師平均目標價（2026-10-25 頁面）；Interested-party",
    cons_fwd="cons_eps27", cons_fwd_note="共識 FY2027 non-GAAP EPS（In 頁 cons_eps27；FY2028 共識未找到，故 EPS 因素以 FY2027 同年度比較：B＝情境加權 FY2027 EPS ÷ 共識 − 1）",
    ours=lambda L: f"(tp_wA*PL!{XC[2027]}{XR[('A', 'eps')]}+tp_wB*PL!{XC[2027]}{XR[('B', 'eps')]}+tp_wC*PL!{XC[2027]}{XR[('C', 'eps')]})",
    eps28={s: f"PL!{XC[2028]}{XR[(s, 'eps')]}" for s in SC}, hist=f"B{PE_ROWS[0]}:B{PE_ROWS[-1]}", fmt="#,##0.00",
    dcf_note="；FY2027 期內占比 10/12（FY2027 約 2026-11～2027-10，基準日 2026 年底後 10 個月）",
    method="DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；EBIT＝non-GAAP 營業利益 − 股票薪酬（股權報酬視為費用；併購無形資產攤銷不計，折舊也不含該攤銷），以 WACC 折現到 2026 年底（期中慣例；FY2027 計 10/12），"
           "終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金 ÷ 稀釋股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁、tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md）。"
           "Comparable：NVIDIA、AMD、Marvell、Qualcomm、台積電 ADR 前瞻 P/E 中位數 × FY2028 EPS。偏誤方向：non-GAAP EPS 排除股票薪酬與攤銷（偏高）；C 情境 FY2030 起淨利率收斂到 FY2029 × 0.8（Assumed）；"
           "歷史分位用十年 GAAP 落後 P/E（攤銷扭曲，只作檢查）；β 取自算產業 β（US 半導體同業 13 家無槓桿 β 中位數 Blume 調整 1.884）以本公司市值 D/E 重新加槓桿（Andy 2026-11-05 22:55）；DCF 對 WACC、股票薪酬口徑敏感（Sens 頁）。")
# ---------------- v0.5：標準版 DCF（FCFF＋WACC）——DCF 明細頁＋Val 0.1 節摘要（取代 v0.2 股東現金流法；tools/fcff_dcf.py） ----------------
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = "DCF — 標準版（v0.5）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py"
DCFW["A2"] = ("EBIT＝non-GAAP 營業利益 − 股票薪酬（sbc_rev × 營收）：FY2027–FY2029 營業利益取 PL 推導區塊；FY2030–FY2032（PL 延伸表只有淨利）由淨利還原：淨利 ÷（1 − tx_f）＋利息 int_f − 其他收益 oth_f（與 PL 同一稅率、同一口徑）。"
              "併購無形資產攤銷為非現金且已排除於 non-GAAP，故折舊只取不含該攤銷的折舊攤銷（fc_da）。利息與其他收益不進 NOPAT，現金與投資另在評價時加入。資本支出、折舊、營運資金比例見 In 頁 M2 節。")
FY_X = list(range(2027, 2033))
def _ebit(s_, y):
    oi = f"PL!{FC[y]}{PB[s_]['oi']}" if y <= 2029 else f"(PL!{XC[y]}{XR[(s_, 'ni')]}/(1-tx_f)+int_f-oth_f)"
    return f"={oi}-sbc_rev*PL!{XC[y]}{XR[(s_, 'rev')]}"
def _tax(s_, y):
    return f"PL!{FC[y]}{PB[s_]['etr']}" if y <= 2029 else "tx_f"
def _fcff_y(s_, y):   # 第三輪修正：年度 FCFF（與 DCF 頁第一年同式：NOPAT＋營收 ×（fc_da − fc_cx − fc_cx_d）− fc_wc × 營收增量）
    rv, rv0 = f"PL!{XC[y]}{XR[(s_, 'rev')]}", f"PL!{XC[y - 1]}{XR[(s_, 'rev')]}"
    return f"(({_ebit(s_, y)[1:]})*(1-{_tax(s_, y)})+{rv}*(fc_da-fc_cx-fc_cx_d)-fc_wc*({rv}-{rv0}))"
def _stub(s_):   # 第三輪修正：stub 期 2026-08-02 → 12-31 跨兩個會計年度：FY2026 Q4（08-03～11-01）＋FY2027 前兩個月（11-02～12-31），各按天數比例
    return f"{_fcff_y(s_, 2026)}*(DATE(2026,11,1)-DATE(2026,8,2))/365+{_fcff_y(s_, 2027)}*(DATE(2026,12,31)-DATE(2026,11,1))/365"
DPATHS = [dict(name=s, label=f"情境 {s}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s] }）", rev={y: f"PL!{XC[y]}{XR[(s, 'rev')]}" for y in FY_X}, rev0=f"PL!{XC[2026]}{XR[(s, 'rev')]}",
              ebit={y: _ebit(s, y) for y in FY_X}, tax={y: _tax(s, y) for y in FY_X}, ebit_note="non-GAAP 營業利益 − sbc_rev × 營收；FY2030 起營業利益＝淨利 ÷（1 − tx_f）＋int_f − oth_f",
              stub_fcff=_stub(s)) for s in SC]
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 10 / 12, title="三情境 FCFF 明細（US$B；FY2027＝2026-11～2027-10，評價日後只計 2027-01～10）", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="23.975-(2.252+57.167)-1.66-1.325",
                             nc_note="2026-08-02 資產負債表與 10-Q（US$B）：現金 23.975 − 債務帳面（短期 2.252＋長期 57.167）− 未攤銷折價與發行成本 1.660 − 營業租賃 1.325（FY2025 年報）") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
_r1, _r2 = TPL["eps_h"] - 2, TPL["dcf"]            # 舊 0.1 節（標題列～DCF 目標價列）整段改寫；列數不變，0.2、0.3 節位置不動
assert V.cell(_r1, 1).value.startswith("0.1 DCF"), V.cell(_r1, 1).value
for _rr in range(_r1, _r2 + 1):
    for _cc in range(1, 8):
        V.cell(_rr, _cc).value = None; V.cell(_rr, _cc).fill = PatternFill(); V.cell(_rr, _cc).font = Font()
assert _r2 - _r1 + 1 >= FD.summary_rows(stub=True)
SDCF = FD.write_summary(wb, V, _r1, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")   # 第二輪修正：模組 v2 beta_note（取代覆寫說明格）
for _k in [k for k in TPL if k.startswith(("eps", "cf")) or k in ("po", "tv", "h1")]:
    del TPL[_k]
TPL.update(dcf=SDCF["dcf"], dcfv=SDCF["v12"], tvsh=SDCF["tvsh"])
V.cell(TPL["tp"], 2).value = f"=tp_w_dcf*B{SDCF['dcf']}+(1-tp_w_dcf)*B{TPL['comp']}"
V.cell(TPL["band"], 2).value = f"=D{SDCF['v12']}"; V.cell(TPL["band"], 3).value = f"=B{SDCF['v12']}"
V.cell(TPL["band"], 5).value = "B＝C 情境 DCF 價值、C＝A 情境 DCF 價值（皆已推到 12 個月後）"
_tvs = FD.write_sens(SE, nn + 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${TPL['comp']}", base_ref=f"Val!$B${TPL['tp']}", beta_du="beta_du")   # β＝β_u ∓ beta_du 重新加槓桿
assert _tvs[2] <= 40, _tvs
TP.dump_rowmap(TPL, TPN, XH, XR, XC, nn + 5)
TP.yellow_inputs(IN)
FD.yellow(IN, FD.JUDG + ("beta_u",))
RD.cell(len(lines) + 1, 1, "v0.2（2026-10-25）：(1) 目標價改為 DCF 與 Comparable 加權（Val 頁最前面「0. 目標價」：0.1 DCF 股東權益現金流、0.2 Comparable 同業前瞻 P/E、0.3 結論；Val 既有內容整體下移，公式一併改寫）；"
        "v0.1.1 的「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留在 Val 頁末。(2) PL 頁末新增三情境延伸到 FY2032：FY2030 起為收斂法，非驅動因子推導（上游市場層只到 CY2029＝FY2029）。"
        "(3) 現價更新到 2026-10-23 收盤 321.68（markets/invest_map/data/prices.csv）。(4) 目標價參數放 In 頁末（黃底，Andy 可改）；Sens 頁新增 tp_r、tp_g、tp_w_dcf 對加權目標價的影響。FY2025–FY2029 既有數字不變。")
RD.cell(len(lines) + 2, 1, "情境對應（v0.2 延伸期）：A＝順利（基準顆數）、B＝平平（CY2028 顆數 ×0.8、CY2029 ×0.85）、C＝不如預期（CY2028–29 ×0.6）；FY2026、FY2027 三情境相同，情境差從 FY2028 起。")


# ---------------- v0.2.1：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；各頁只在現有最後一列之後追加，不位移既有列 ----------------

# (A) 自由現金流（算力收支 2026-10-25 請求；地圖依列標籤「自由現金流」抽數）：In、PL、Sources 只在現有最後一列之後追加
_ir = IN.max_row + 2
IN.cell(_ir, 1, "N. 自由現金流假設（v0.2.1 新增；FY2025 實績為錨，各年營收 × 比率）").font = BOLD
for _j, (_k, _it, _v, _s1) in enumerate((("da_rev", '折舊 ÷ 營收（只含不動產、廠房設備折舊；non-GAAP 淨利已排除無形資產攤銷，不再加回；自由現金流用）', '=0.574/63.887', 'Broadcom FY2025 年報數字（會計年度至 2025-11-02；8-K Exhibit 99.1 現金流量表，與 10-K 相同）：Depreciation 574 ÷ 營收 63,887（US$M）'),
                                        ("cx_rev", '資本支出 ÷ 營收（購置不動產、廠房設備；自由現金流用）', '=0.623/63.887', '同上：Purchases of property, plant and equipment 623 ÷ 營收 63,887（US$M）')), 1):
    IN.cell(_ir + _j, 1, _k); IN.cell(_ir + _j, 2, _it); IN.cell(_ir + _j, 3, "%"); IN.cell(_ir + _j, 5, _v); IN.cell(_ir + _j, 5).number_format = "0.00%"
    IN.cell(_ir + _j, 7, "Verified"); IN.cell(_ir + _j, 8, _s1); IN.cell(_ir + _j, 9, '各年沿用 FY2025 比率（Assumed：比率不變）；FY2024 為 1.15%／1.06%（593、548 ÷ 51,574）')
    wb.defined_names[_k] = DefinedName(_k, attr_text=f"In!$E${_ir + _j}")
FCY = list(range(2026, 2033))
_pr = PL.max_row + 2
PL.cell(_pr, 1, "v0.2.1 新增：每年自由現金流＝淨利（延伸表）＋營收 × da_rev − 營收 × cx_rev（In 頁 N 節；FY2025 比率沿用）；不含營運資金變動。欄位：各情境 × 各年度").font = BOLD
hdr(PL, _pr + 1, ["項目", "單位"] + [f"FY{y} {s}" for s in SC for y in FCY] + ["說明"])
FCF_ROW = _pr + 2
PL.cell(FCF_ROW, 1, "自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）"); PL.cell(FCF_ROW, 2, "US$B")
for _j, (_s, _y) in enumerate([(s, y) for s in SC for y in FCY]):
    _c = XC[_y]; _ni = XR[(_s, "ni")]; _rv = XR[(_s, "rev")]
    PL.cell(FCF_ROW, 3 + _j, f"={_c}{_ni}+{_c}{_rv}*da_rev-{_c}{_rv}*cx_rev").number_format = "0.0"
PL.cell(FCF_ROW, 3 + len(SC) * len(FCY), '淨利＝PL 延伸表「淨利（延伸表）」，non-GAAP（排除股票薪酬、無形資產攤銷等）：股票薪酬為非現金費用，排除後接近現金口徑；無形資產攤銷已不在淨利內，故折舊只加回不動產、廠房設備折舊')
_sr = SO.max_row + 1
for _j, (_a, _b) in enumerate([('自由現金流假設 da_rev、cx_rev：Broadcom FY2025 現金流量表 [Verified]', 'https://investors.broadcom.com/static-files/dd25b8a9-a65a-4125-b02b-b1a6eb154cde（8-K，2025-12-11：Depreciation 574、Amortization of intangible and right-of-use assets 8,201、資本支出 623、營收 63,887，US$M）；10-K：https://www.sec.gov/Archives/edgar/data/1730168/000173016825000121/avgo-20251102.htm'), ('自由現金流各年比率沿用 FY2025 [Assumed]', '折舊、資本支出占營收比例 FY2026 起不變；未含營運資金變動；淨利為 non-GAAP')]):
    SO.cell(_sr + _j, 1, _a); SO.cell(_sr + _j, 2, _b)
_up = next(r for r in range(1, V.max_row + 1) if V.cell(r, 1).value == "上檔空間：加權目標價")
GRADE_ROW = V.max_row + 2; REC_ROW = GRADE_ROW + 1
V.cell(GRADE_ROW, 1, "上檔空間級距")
V.cell(GRADE_ROW, 2, f'=IF(B{_up}>0.5,"大於 50%",IF(B{_up}>0.25,"25%–50%",IF(B{_up}>=0,"0%–25%",IF(B{_up}>=-0.25,"0%～−25%",IF(B{_up}>=-0.5,"−25%～−50%","小於 −50%")))))')
V.cell(GRADE_ROW, 5, f"以「上檔空間：加權目標價」（B{_up}）分六級（Andy 2026-10-25 14:00）")
V.cell(REC_ROW, 1, "投資建議")
V.cell(REC_ROW, 2, f'=IF(B{_up}>0.5,"強力買進",IF(B{_up}>0.25,"買進",IF(B{_up}>=-0.25,"中立",IF(B{_up}>=-0.5,"放空","強力放空"))))')
V.cell(REC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = ['目標價 US$341（DCF 325、Comparable 357 各半），比現價 321.68（2026-10-23）高 5.9%，級距 0%–25%；Comparable 單獨為 +10.9%。與共識目標價 529.1 的差距主要在倍數：本模型 FY2027 EPS 18.08 只比共識 19.40 低 6.8%，但共識目標價隱含 27.3 倍共識 EPS，比同業（NVIDIA、AMD、Marvell、Qualcomm、台積電 ADR）前瞻 P/E 中位數 19.6 倍高 39.1%；現價隱含 16.6 倍。', '情境分歧從 FY2028 起：FY2027 EPS 三情境同為 18.08，FY2028 A／B／C 為 21.03／17.84／14.82（B、C 低於 FY2027），DCF 價格帶 222～394；現價約為 B 情境 FY2028 EPS 的 18.0 倍。PL 頁末自由現金流（non-GAAP 淨利＋折舊 − 資本支出）B 情境 FY2027 約 88.8、FY2028 約 87.2 十億美元，與淨利幾乎相同（資本支出約為營收 1%）。', 'FY2027 EPS 最敏感的是 XPU 核心售價（±20% 對 EPS ±8.2%），其次是 TPU 份額（−15pt 時 −5.7%，聯發科版 TPU 放量）與 AI 網通（×0.8 時 −5.1%）（Sens 頁）。下一個檢驗點是 FY2026 第四季財報（約 2026-12 中旬）的 AI 半導體營收與下一季指引。']   # 重點觀察（與 README.md 相同文字；數字取自重算後的 Val／PL／Sens 頁）
_rr = RD.max_row + 1
RD.cell(_rr, 1, 'v0.2.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁末新增「上檔空間級距」「投資建議」兩列（Andy 2026-10-25 14:00、14:17 固定規則），README 頁末新增投資建議與重點觀察三點；既有儲存格位置與數值不變。另依算力收支 2026-10-25 請求，PL 頁末新增「自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）」一列（三情境 × 各年度），In 頁末新增 da_rev、cx_rev 兩個比率（最近完整會計年度實績）。')
RD.cell(_rr + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"').font = Font(bold=True)
for _j, _t in enumerate(OBS, 1): RD.cell(_rr + 1 + _j, 1, f"重點觀察 {_j}：" + _t)
# v0.5：重點觀察 2 的 DCF 價格帶改用公式（下游 tools/target_summary 讀本列文字；DCF 改為標準版）
_o2 = OBS[1].split("DCF 價格帶 222～394")
assert len(_o2) == 2
RD.cell(_rr + 3, 1, '="重點觀察 2：' + _o2[0] + 'DCF 價格帶 "&TEXT(Val!B' + str(TPL["band"]) + ',"0")&"～"&TEXT(Val!C' + str(TPL["band"]) + ',"0")&"（v0.5 標準版 DCF）' + _o2[1] + '"')
if os.environ.get("RATING_ROWMAP"):   # 除錯用：輸出新增列號（供撰寫 html.json）；CI 不設此變數
    print("ROWMAP", GRADE_ROW, REC_ROW, _rr, _rr + 1, _rr + 2, FCF_ROW)


order =["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "avgo")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.4", "2026-11-05")  # v0.4：CF13 (a) 金額份額改整顆售價（只影響份額展示）；v0.3.1：In 頁版面

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
def _wta_tp(SE, r0, Rw, years, last_up, keys, L, XC, XR, fwd, m10, fmt, unit):
    """Wattonomics 錨定案例的單一路徑目標價（Sens 頁）：延伸到 FY2032 的收斂法與 PL 延伸表情境 B 同式（tp_v2.extend_pl），
    DCF 與 Val 第 0 節同式（折現年數、期內占比取 Val F／G 欄），Comparable＝tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS。
    同時以 PL 延伸表情境 B 列算情境 B 單一路徑目標價供對照。回傳 {鍵: 儲存格}。"""
    from openpyxl.styles import Font as _F
    allys = list(years) + list(range(last_up + 1, 2033)); conv = [y for y in allys if y > last_up]; N = len(conv)
    SE.cell(r0, 1, "Wattonomics 錨定：延伸到 FY2032（收斂法同 PL 延伸表情境 B）與單一路徑目標價").font = _F(bold=True)
    SE.cell(r0 + 1, 1, "項目"); SE.cell(r0 + 1, 2, "單位")
    for y in allys: SE[f"{XC[y]}{r0 + 1}"] = f"FY{y}"
    rr = {k: r0 + 2 + i for i, k in enumerate(("rev", "g", "nm", "ni", "sh", "eps"))}
    for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nm", "淨利率"), ("ni", "淨利"), ("sh", "稀釋股數"), ("eps", f"EPS（{unit}）")): SE.cell(rr[k], 1, lab)
    lc = XC[last_up]
    M10 = "nm10" if m10 == "nm10" else f"${lc}${rr['nm']}*nm10_k"
    mT = f"(${lc}${rr['nm']}+{M10})/2"
    for y in allys:
        c = XC[y]; p = XC.get(y - 1)
        if y <= last_up:
            for k in ("rev", "ni", "sh", "eps"): SE[f"{c}{rr[k]}"] = f"={c}{Rw[keys[k]]}"
            if p: SE[f"{c}{rr['g']}"] = f"={c}{rr['rev']}/{p}{rr['rev']}-1"
            SE[f"{c}{rr['nm']}"] = f"={c}{rr['ni']}/{c}{rr['rev']}"
        else:
            k_ = y - last_up
            SE[f"{c}{rr['g']}"] = f"=${lc}${rr['g']}+(tp_g-${lc}${rr['g']})*{k_}/{N}"
            SE[f"{c}{rr['rev']}"] = f"={p}{rr['rev']}*(1+{c}{rr['g']})"
            SE[f"{c}{rr['nm']}"] = f"=${lc}${rr['nm']}+({mT}-${lc}${rr['nm']})*{k_}/{N}"
            SE[f"{c}{rr['ni']}"] = f"={c}{rr['rev']}*{c}{rr['nm']}"
            SE[f"{c}{rr['sh']}"] = f"={p}{rr['sh']}*(1+${lc}${rr['sh']}/${XC[last_up - 1]}${rr['sh']}-1)"
            SE[f"{c}{rr['eps']}"] = f"={c}{rr['ni']}/{c}{rr['sh']}"
        for k, f_ in (("rev", "#,##0.0"), ("g", "0.0%"), ("nm", "0.0%"), ("ni", "#,##0.0"), ("sh", "0.000"), ("eps", "#,##0.00")): SE[f"{c}{rr[k]}"].number_format = f_
    # v0.4：DCF 改為標準版（FCFF＋WACC）：本路徑在 DCF 頁另寫一個明細區塊（tools/fcff_dcf.py，與三情境同式）
    _eb = lambda y: (f"=Sens!{XC[y]}{Rw['oi']}" if y <= last_up else f"=(Sens!{XC[y]}{rr['ni']}/(1-tx_f)+int_f-oth_f)") + f"-sbc_rev*Sens!{XC[y]}{rr['rev']}"
    _P = dict(name="W", label="Wattonomics 錨定（情境 B × wta；Sens 案例，未進基準）", rev={y: f"Sens!{XC[y]}{rr['rev']}" for y in range(2027, 2033)}, rev0=f"Sens!{XC[2026]}{rr['rev']}",
              ebit={y: _eb(y) for y in range(2027, 2033)},
              tax={y: (f"Sens!{XC[y]}{Rw['etr']}" if y <= last_up else "tx_f") for y in range(2027, 2033)}, ebit_note="non-GAAP 營業利益 − sbc_rev × 營收；FY2030 起營業利益＝淨利 ÷（1 − tx_f）＋int_f − oth_f")
    RW = FD.write_detail(wb["DCF"], RD_["end"] + 1, [_P], list(range(2027, 2033)), 1.0, title="Sens 案例路徑：Wattonomics 錨定（情境 B × wta）FCFF 明細", stub=True)
    v12 = lambda R_, n: f"DCF!$C${R_['paths'][n]['v12']}"
    ew = lambda y: f"{XC[y]}{rr['eps']}"; eb = lambda y: f"PL!{XC[y]}{XR[('B', 'eps')]}"
    h = rr["eps"] + 2
    for j, lab in enumerate(["目標價（單一路徑）", "Wattonomics 錨定", "情境 B（PL 延伸表）", "基準（情境加權，Val 第 0 節）"]): SE.cell(h, 1 + j, lab).font = _F(bold=True)
    T = {"dcf": h + 1, "fwd": h + 2, "comp": h + 3, "tp": h + 4, "chk": h + 5}
    SE.cell(T["dcf"], 1, "DCF 價值推到 12 個月後（標準版 FCFF＋WACC；×（1＋ke）^t_tgt）"); SE.cell(T["dcf"], 2, "=" + v12(RW, "W")); SE.cell(T["dcf"], 3, "=" + v12(RD_, "B")); SE.cell(T["dcf"], 4, f"=Val!B{L['dcf']}")
    SE.cell(T["fwd"], 1, "12 個月後前瞻 EPS（權重同 Val 第 0 節）")
    SE.cell(T["fwd"], 2, "=" + "+".join(f"{w}*{ew(y)}" for y, w in fwd)); SE.cell(T["fwd"], 3, "=" + "+".join(f"{w}*{eb(y)}" for y, w in fwd)); SE.cell(T["fwd"], 4, f"=Val!B{L['fwd']}")
    SE.cell(T["comp"], 1, "Comparable（相對倍數調整後；tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS）")
    for c in "BC": SE[f"{c}{T['comp']}"] = f"=tp_comp_adj*{c}{T['fwd']}/Val!B{L['fwd']}"
    SE[f"D{T['comp']}"] = "=tp_comp_adj"
    SE.cell(T["tp"], 1, "目標價（tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable）")
    for c in "BC": SE[f"{c}{T['tp']}"] = f"=tp_w_dcf*{c}{T['dcf']}+(1-tp_w_dcf)*{c}{T['comp']}"
    SE[f"D{T['tp']}"] = f"=Val!B{L['tp']}"
    for k in ("dcf", "fwd", "comp", "tp"):
        for c in "BCD": SE[f"{c}{T[k]}"].number_format = "#,##0.00" if k == "fwd" else fmt
    SE.cell(T["chk"], 1, "核對：情境 B 欄 DCF − Val 第 0.1 節情境 B 價值推到 12 個月後（應為 0）"); SE.cell(T["chk"], 3, f"=C{T['dcf']}-Val!C{L['dcfv']}").number_format = "0.000000"
    return {"tp_w": f"B{T['tp']}", "tp_b": f"C{T['tp']}", "tp_base": f"D{T['tp']}", "eps": ew, "end": T["chk"]}


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
SE.cell(_w0 + 2, 1, '做法：FY(t)＝CY(t)。FY2027 起 Broadcom XPU 顆數（市場顆數 × 份額）與 AI 網通（markets/network 歸屬 × 情境顆數係數）再乘 wta（FY2027＝wta_27 … FY2029＝wta_29）；k_x、k_n、m_core 由 FY2026 校準不受影響；非 AI 半導體、軟體不乘。FY2030 起沿用 PL 延伸表情境 B 的收斂法。')
_t0 = _w0 + 4
SB_W = engine(SE, _t0 + 6, "案例區塊：Wattonomics 錨定（情境 B × wta）", scn="B", wta=WTA_D)
_W = _wta_tp(SE, SB_W["end"] + 1, SB_W, FYS, 2029, {'rev': 'rev', 'ni': 'ni', 'sh': 'sh', 'eps': 'eps'}, TPL, XC, XR, [(2028, 1)], 'nm10_k', '#,##0.00', 'US$')
_e = lambda s_, y: f"PL!{XC[y]}{XR[(s_, 'eps')]}"
_wta_table(SE, _t0, [("FY2027 EPS", "0.00"), ("FY2028 EPS", "0.00"), ("目標價", '#,##0.00')],
           [("基準（Sens 主表基準＝情境 A；目標價＝情境加權）", [_e("A", 2027), _e("A", 2028), _W["tp_base"]]),
            ("情境 B（案例底）", [_e("B", 2027), _e("B", 2028), _W["tp_b"]]),
            ("Wattonomics 錨定（情境 B × wta）", [_W["eps"](2027), _W["eps"](2028), _W["tp_w"]])])
# ---------------- v0.4.1（2026-11-05）：只改展示層——README 頁「展示層」「命題回覆」「本版要點」三段（個股 HTML 版面規格 v1.0；公式帶數字，其他頁不變） ----------------
def _fr(ws, prefix, start=1):
    for _r in range(start, ws.max_row + 1):
        _v = ws.cell(_r, 1).value
        if isinstance(_v, str) and _v.strip().startswith(prefix): return _r
    raise KeyError((ws.title, prefix))
_VB = lambda k: f"Val!B{TPL[k]}"
_E = lambda s_, y: f"PL!{XC[y]}{XR[(s_, 'eps')]}"
_T = lambda x, f: f'TEXT({x},"{f}")'
_PP = lambda x: f'TEXT(({x})*100,"0.0")'
_SG = '"+0.0%;-0.0%"'
_vadj = _fr(V, "相對倍數調整係數"); _vown = _fr(V, "本公司 10 年中位數"); _vpeer = _fr(V, "同業 10 年中位數的中位數"); _vband = TPL["band"]
_vpe27 = _fr(V, "現價 ÷ FY2027 EPS（模型）"); _vpec = _fr(V, "現價 ÷ 共識 FY2027 EPS")
_r_ai27 = _fr(RC, "FY2027 AI 營收 vs 公司展望"); _r_jpm = _fr(RC, "FY2027 AI 營收 vs JPMorgan"); _r_ai28 = _fr(RC, "FY2028 AI 營收 vs 公司展望")
_r_m27 = _fr(RC, "FY2027：所需 XPU 顆數倍數"); _r_m28 = _fr(RC, "FY2028：所需 XPU 顆數倍數"); _r_u28 = _fr(RC, "FY2028：所需 Broadcom XPU 顆數")
_r_rev27 = _fr(RC, "FY2027 營收 vs 共識"); _r_eps27 = _fr(RC, "FY2027 EPS vs 共識"); _r_gmc = _fr(RC, "FY2027 毛利率 vs 共識隱含")
_r_fin = _fr(RC, "Anthropic 放款＋OpenAI"); _r_fai = _fr(RC, "兩家合計 ÷ FY2027 AI 營收"); _r_fni = _fr(RC, "兩家合計 ÷ FY2027 淨利")
_s_c1 = _fr(SE, "核心售價 ×1.2"); _s_c0 = _fr(SE, "核心售價 ×0.8"); _s_tpu = _fr(SE, "TPU 份額 −15pt"); _s_net = _fr(SE, "AI 網通 ×0.8")
_s_wta = _fr(SE, "Wattonomics 錨定（情境 B × wta）", _w0 + 1)
_ADJM = f"Val!B{TPL['med']}*Val!B{_vadj}"   # 相對倍數調整後的目標倍數
RD.cell(1, 1).value = (RD.cell(1, 1).value.replace("模型 v0.2.1（", "模型 v0.5.3（", 1) + "；v0.4.1（2026-11-05）只改展示層（README 展示層、命題回覆與本版要點三段，HTML 新版面 v1.0）"
                       "；v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
                       "；v0.5.1（2026-11-05）β 改自算產業 β（US 半導體同業中位數 Blume 調整 1.884 重新加槓桿；Andy 2026-11-05 22:55），其他輸入不變"
                       "；v0.5.2（2026-11-05）換用共用模組 tools/fcff_dcf.py v2：淨現金由資產負債表日 2026-08-02 滾到評價日 2026-12-31（stub 期 FCFF − 已宣告股利）、β 敏感度改 β_u ± 0.3 重新加槓桿"
                       "；v0.5.3（2026-11-06）第三輪修正：stub 期拆成 FY2026 Q4＋FY2027 前兩個月；現價與同業報價統一到 2026-11-04 收盤；展示文字與 HTML")
_r = RD.max_row + 2; _r0 = _r
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.4.1 新增，A3、A4、A20–A23 原文保留給下游）").font = BOLD; _r += 1
RD.cell(_r, 1, "命題：Broadcom 的 AI 營收＝各 XPU 平台的顆數 × Broadcom 份額 × 每顆售價（核心＋代購 HBM）＋網通附掛；2027–28 年的增量集中在 Anthropic、OpenAI 兩家靠融資部署的客戶。問題是以市場層顆數推得的營收撐不撐得住公司展望，以及現價的低倍數在反映什麼，四個問題如下。"); _r += 1
RD.cell(_r, 1, f'="命題 1：以市場層顆數往前推，Broadcom 的 AI 營收能否達到公司展望（FY2027 約 "&TEXT(co_ai27,"0")&"、FY2028 約 "&TEXT(co_ai28,"0")&" 十億美元）。"'); _r += 1
RD.cell(_r, 1, "命題 2：達不到的話，差距要靠多少額外的 Broadcom XPU 顆數補上。"); _r += 1
RD.cell(_r, 1, "命題 3：增量集中在兩家靠融資部署的客戶，Broadcom 自身承擔的客戶融資曝險有多大。"); _r += 1
RD.cell(_r, 1, f'="命題 4：股價約為共識 FY2027 EPS 的 "&TEXT(Val!B{_vpec},"0")&" 倍：這個折價是在定價客戶信用風險，還是低估了營收。"'); _r += 1
RD.cell(_r, 1, f'="重點 1：加權目標價約 "&{_T(_VB("tp"), "0.00")}&" 美元，比現價 "&TEXT(px,"0.00")&" 低 "&TEXT(-{_VB("up")},"0.0%")&"，投資建議「"&Val!B{REC_ROW}&"」。拆到估值來源：加權目標價相對現價的差距（"&TEXT({_VB("tp")}-px,"+0.00;-0.00")&" 美元）中，DCF 部分 "&TEXT(tp_w_dcf*({_VB("dcf")}-px),"+0.00;-0.00")&"、Comparable 部分 "&TEXT((1-tp_w_dcf)*(tp_comp_adj-px),"+0.00;-0.00")&"（DCF 約占 "&TEXT(tp_w_dcf*({_VB("dcf")}-px)/({_VB("tp")}-px),"0%")&"）。標準版 DCF "&{_T(_VB("dcf"), "0.00")}&" 比現價"&IF({_VB("dcf")}<px,"低 "&TEXT(1-{_VB("dcf")}/px,"0.0%"),"高 "&TEXT({_VB("dcf")}/px-1,"0.0%"))&"：EBIT 扣股票薪酬（營收的 "&TEXT(sbc_rev,"0.0%")&"，non-GAAP EPS 不扣）後以 WACC "&TEXT(wacc,"0.0%")&" 折現（見重點 4）。進加權的是相對倍數調整後的 Comparable "&TEXT(tp_comp_adj,"0.00")&"：目標倍數＝同業前瞻本益比中位數 "&{_T(_VB("med"), "0.0")}&" 倍 × 調整係數 "&TEXT(Val!B{_vadj},"0.00")&"（Broadcom 過去 10 年前瞻本益比中位數 "&TEXT(Val!B{_vown},"0.0")&" 倍 ÷ 同業 "&TEXT(Val!B{_vpeer},"0.0")&" 倍）＝"&TEXT({_ADJM},"0.0")&" 倍。以同一個 12 個月後前瞻 EPS（情境加權 "&{_T(_VB("fwd"), "0.00")}&"）計，現價相當於 "&TEXT(px/{_VB("fwd")},"0.0")&" 倍，所以 Comparable 與現價的差距全部是倍數差，不是盈餘差。只看 Comparable（市場主流同業法、未調整）為 "&{_T(_VB("comp"), "0.00")}&"（"&TEXT({_VB("upc")},{_SG})&"）。"'); _r += 1
RD.cell(_r, 1, f'="重點 2：FY2027 三情境相同（EPS "&{_T(_E("B", 2027), "0.00")}&" 美元），分歧從 FY2028 開始：FY2028 EPS 情境 A（順利）／B（平平）／C（不如預期）為 "&{_T(_E("A", 2028), "0.00")}&"／"&{_T(_E("B", 2028), "0.00")}&"／"&{_T(_E("C", 2028), "0.00")}&"，B、C 低於 FY2027。DCF 價格帶（情境 C～A，推到約 12 個月後）"&TEXT(Val!B{_vband},"0.0")&"～"&TEXT(Val!C{_vband},"0.0")&"。FY2028 AI 營收三情境為 "&TEXT(Recon!C{_r_ai28},"0.0")&"／"&TEXT(Recon!E{_r_ai28},"0.0")&"／"&TEXT(Recon!G{_r_ai28},"0.0")&"（十億美元），都低於公司展望 "&TEXT(Recon!B{_r_ai28},"0")&"。"'); _r += 1
RD.cell(_r, 1, f'="重點 3：單一輸入（Sens 頁一次動一個，以情境 A 為基準）中，FY2028 EPS 對 XPU 核心售價最敏感：×0.8／×1.2 時 "&TEXT(Sens!L{_s_c0},{_SG})&"／"&TEXT(Sens!L{_s_c1},{_SG})&"（每顆核心售價 TPU "&TEXT(core_tpu,"0")&"、其他 ASIC "&TEXT(core_oth,"0")&" 千美元為假設）；其次是 TPU 份額少 15 個百分點 "&TEXT(Sens!L{_s_tpu},{_SG})&"（份額為假設）與 AI 網通 ×0.8 "&TEXT(Sens!L{_s_net},{_SG})&"，都未達兩位數。改用 Wattonomics 錨定的產業量（情境 B × 錨定係數，敏感度案例、未進基準）時，單一路徑目標價 "&TEXT(Sens!H{_s_wta},"0.00")&"（相對加權目標價 "&TEXT(Sens!I{_s_wta},{_SG})&"）。"'); _r += 1
RD.cell(_r, 1, f'="重點 4：v0.5 起 DCF 改為標準版（無槓桿自由現金流折現），DCF 目標價 "&{_T(_VB("dcf"), "0.00")}&"（舊法為 non-GAAP EPS 的股東現金流、折現率 10%；新舊數字對照見版本沿革），投資建議因此由中立轉為「"&Val!B{REC_ROW}&"」。WACC "&TEXT(wacc,"0.0%")&"（股權成本 "&TEXT(ke,"0.0%")&"＝rf "&TEXT(rf,"0.0%")&" ＋ β "&TEXT(beta,"0.00")&" × ERP "&TEXT(erp,"0%")&"；β 為自算產業 β（US 半導體同業中位數 Blume 調整）依本公司市值 D/E 重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設）；EBIT 為 non-GAAP 營業利益扣股票薪酬（營收的 "&TEXT(sbc_rev,"0.0%")&"）。終值現值占情境 B 企業價值 "&TEXT(Val!C{SDCF["tvsh"]},"0%")&"，隱含終值 EV／EBITDA 約 "&TEXT(Val!C{SDCF["mult"]},"0.0")&" 倍。以情境加權目標價為基準：WACC ±1pt 時 "&TEXT(Sens!D{_tvs[0]},{_SG})&"／"&TEXT(Sens!F{_tvs[0]},{_SG})&"，β 低／高 "&TEXT(Sens!D{_tvs[0] + 1},{_SG})&"／"&TEXT(Sens!F{_tvs[0] + 1},{_SG})&"，終值改用 FCFF ×（1＋g） "&TEXT(Sens!F{_tvs[0] + 8},{_SG})&"。"&IF(MAX(Val!B{SDCF["tvsh"]}:D{SDCF["tvsh"]})>0.8,"終值現值占 EV 超過 80% 的情境："&IF(Val!B{SDCF["tvsh"]}>0.8,"A ","")&IF(Val!C{SDCF["tvsh"]}>0.8,"B ","")&IF(Val!D{SDCF["tvsh"]}>0.8,"C ","")&"（DCF 價值高度依賴終值假設，需留意）。","")'); _r += 1
_r += 1
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算回答開頭四個命題）").font = BOLD; _r += 1
RD.cell(_r, 1, f'="AI 營收能否達到公司展望？答：FY2027 可以，FY2028 三情境都達不到。FY2027 AI 營收 "&TEXT(Recon!C{_r_ai27},"0.0")&"（十億美元，三情境相同），比公司展望 "&TEXT(Recon!B{_r_ai27},"0")&" 高 "&TEXT(Recon!D{_r_ai27},"0.0%")&"，但比 JPMorgan 的 "&TEXT(Recon!B{_r_jpm},"0")&" 低 "&TEXT(-Recon!D{_r_jpm},"0.0%")&"；FY2028 情境 A／B／C 為 "&TEXT(Recon!C{_r_ai28},"0.0")&"／"&TEXT(Recon!E{_r_ai28},"0.0")&"／"&TEXT(Recon!G{_r_ai28},"0.0")&"，分別比展望 "&TEXT(Recon!B{_r_ai28},"0")&" 低 "&TEXT(-Recon!D{_r_ai28},"0.0%")&"／"&TEXT(-Recon!F{_r_ai28},"0.0%")&"／"&TEXT(-Recon!H{_r_ai28},"0.0%")&"。顆數取自市場層；Broadcom 在各平台的份額與每顆核心售價為假設（Assumed），公司展望為 Interested-party、未用於校準。"'); _r += 1
RD.cell(_r, 1, f'="差距要靠多少額外顆數補上？答：在網通、核心售價、HBM 不變的前提下反推，FY2028 要達到展望，Broadcom XPU 顆數須為模型的 "&TEXT(Recon!C{_r_m28},"0.00")&" 倍（情境 A）／"&TEXT(Recon!E{_r_m28},"0.00")&" 倍（B）／"&TEXT(Recon!G{_r_m28},"0.00")&" 倍（C），約 "&TEXT(Recon!C{_r_u28},"0.0")&"／"&TEXT(Recon!E{_r_u28},"0.0")&"／"&TEXT(Recon!G{_r_u28},"0.0")&" 百萬顆（情境 A 的模型值為 "&TEXT(Recon!B{_r_u28},"0.0")&" 百萬顆）。FY2027 只需模型的 "&TEXT(Recon!C{_r_m27},"0.00")&" 倍，模型已超過展望。這是其他條件不變下的反推，不是預測。"'); _r += 1
RD.cell(_r, 1, f'="客戶融資曝險有多大？答：媒體報導的 Anthropic 放款上限 "&TEXT(fin_anth,"0")&"（十億美元，下同）與 OpenAI 融資安排 "&TEXT(fin_oai,"0")&" 合計 "&TEXT(Recon!B{_r_fin},"0")&"（標題值；非資產負債表曝險，部分由銀行團分擔），約為本模型 FY2027 AI 營收的 "&TEXT(Recon!C{_r_fai},"0%")&"、FY2027 淨利的 "&TEXT(Recon!C{_r_fni},"0%")&"（全數損失時的量級，不含回收）。另有一家客戶機架租賃擔保約 "&TEXT(fin_lease,"0")&"，可能與上列重疊、不加總。本模型基準 EPS 不計入這項風險。"'); _r += 1
RD.cell(_r, 1, f'="折價是在定價信用風險，還是低估了營收？答：本模型找不到營收被低估的證據，也無法單獨辨識信用風險。FY2027 營收 "&TEXT(Recon!C{_r_rev27},"0.0")&" 與共識 "&TEXT(Recon!B{_r_rev27},"0.0")&" 只差 "&TEXT(Recon!D{_r_rev27},{_SG})&"，EPS "&TEXT(Recon!C{_r_eps27},"0.00")&" 比共識 "&TEXT(Recon!B{_r_eps27},"0.00")&" 低 "&TEXT(-Recon!D{_r_eps27},"0.0%")&"（FY2027 毛利率 "&TEXT(Recon!C{_r_gmc},"0.0%")&"，比共識隱含約 "&TEXT(Recon!B{_r_gmc},"0.0%")&" 低 "&{_PP(f"-Recon!D{_r_gmc}")}&" 個百分點；本模型的毛利率下滑來自代購 HBM 占比上升，代購比例 "&TEXT(f_hbm,"0%")&"、代購毛利率 "&TEXT(m_hbm,"0%")&" 為假設）。同窗比較，現價約為本模型 FY2027 EPS（FY2027 約為 2026-11～2027-10，接近未來 12 個月）的 "&TEXT(Val!B{_vpe27},"0.0")&" 倍，低於同業前瞻本益比中位數（未來 12 個月）"&{_T(_VB("med"), "0.0")}&" 倍；但 Broadcom 過去 10 年前瞻本益比中位數本來就只有同業的 "&TEXT(Val!B{_vadj},"0.00")&" 倍，依這個相對倍數調整後的 Comparable 為 "&TEXT(tp_comp_adj,"0.00")&"，低於現價。所以現價的低倍數與 Broadcom 自身的歷史折價一致；加權目標價 "&{_T(_VB("tp"), "0.00")}&"（"&TEXT({_VB("up")},{_SG})&"）。"'); _r += 1
_r += 1
RD.cell(_r, 1, "本版要點與版本沿革（HTML 附錄用）").font = BOLD; _r += 1
RD.cell(_r, 1, f'="v0.5.3（2026-11-06）：第三輪修正。(1) Stub 期（2026-08-02 → 12-31）跨兩個會計年度，FCFF 改為 FY2026 FCFF × 91／365（Q4 FY26：08-03～11-01）＋FY2027 FCFF × 60／365（11-02～12-31），原用 FY2027 全年 × stub_frac：情境 B "&TEXT(Val!C{_fr(V, "＋Stub 期 FCFF")},"0.00")&"（US$B），淨現金滾到評價日 "&TEXT(Val!C{_fr(V, "淨現金滾到評價日")},"0.00")&"。(2) 現價改 2026-11-04 收盤 "&TEXT(px,"0.00")&"（Yahoo Finance；原 2026-10-23 321.68）；Comparable 同業前瞻 P/E 改為 11-04 收盤 ÷ 資料商前瞻 EPS（stockanalysis 統計頁 10-27～10-30 價格 ÷ Forward PE），中位數 "&TEXT(Val!B{_fr(V, "同業前瞻 P/E 中位數")},"0.00")&"。加權目標價 "&TEXT(Val!B{TPL["tp"]},"0.00")&"（上檔 "&TEXT(Val!B{_up},"+0.0%;-0.0%")&"，"&Val!B{REC_ROW}&"）。(3) 展示文字：情境名稱 A（順利）／B（平平）／C（不如預期）；來源工具字樣改「未取得」；β 來源引用 docs/plan/自算產業β_20261105.md；重點觀察 1 刪除 2026-10-29 調整前舊說明；級距列名「六級，供分組」。(4) HTML：DCF 彙總主表顯示「＋Stub 期 FCFF」「淨現金滾到評價日」；情境驅動表年度無千分位、百分比格式化；橫幅標明美元。"'); _r += 1
RD.cell(_r, 1, '="v0.5.2（2026-11-05）：換用共用模組 tools/fcff_dcf.py v2（工作單第二輪修正；Andy 2026-11-05 23:16 定案）。(1) Stub 期：資產負債表日 2026-08-02（10-Q，Q3 FY26，最新已公布一季）與評價日 2026-12-31 不同，淨現金滾到評價日＝資產負債表淨現金 "&TEXT(dcf_nc,"0.00")&"＋stub 期 FCFF（各情境 FY2027 FCFF × stub_frac "&TEXT(stub_frac,"0.000")&" 年；情境 B "&TEXT(Val!C' + str(SDCF["stub"]) + ',"0.00")&"）− 已宣告股利 "&TEXT(stub_div,"0.00")&"（每股 0.65 美元，2026-09-30 發放；Q4 FY26 股利尚未宣告、買回無已宣告金額，皆不扣）＝"&TEXT(Val!C' + str(SDCF["ncr"]) + ',"0.00")&"（US$B）。DCF 目標價 144.06 → "&TEXT(Val!B' + str(SDCF["dcf"]) + ',"0.00")&"、加權目標價 182.86 → "&TEXT(Val!B' + str(TPL["tp"]) + ',"0.00")&"（上檔 "&TEXT(Val!B' + str(TPL["up"]) + ',"+0.0%;-0.0%")&"，"&Val!B' + str(REC_ROW) + '&"）。(2) β 敏感度：低／高改為 β_u ∓ 0.3 各自重新加槓桿（In 頁 beta_du；beta_u 低／高欄刪除）。(3) ke 說明格改由模組參數 beta_note 提供；In 頁 M2 標題不寫版號。資本支出、折舊：本模型沒有逐年資本支出／折舊列（只有近三期比率），維持占營收比例。美國公司不加「美元口徑折現」案例（＝基準）。PL、Comparable、情境權重不變。"'); _r += 1
RD.cell(_r, 1, "v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55，取代 Damodaran 表）：US 半導體同業 13 家（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 迴歸，Hamada 去槓桿、現金修正）無槓桿 β 中位數 2.32，Blume 調整 0.67β＋0.33＝1.884，以本公司市值 D/E 與稅率重新加槓桿；Sens 低／高＝中位數 × 0.8 再 Blume／中位數未調整（同式重新加槓桿）。評價日 2026-12-31 不變；其他輸入不變。"); _r += 1
RD.cell(_r, 1, "v0.5（2026-11-05）：模型改版——DCF 改為標準版（Andy 2026-11-05 20:42、20:45；docs/plan/個股DCF標準規格.md）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；EBIT＝non-GAAP 營業利益 − 股票薪酬（sbc_rev），折舊不含併購無形資產攤銷；"
        "以 WACC 折現（ke＝rf＋β ×（ERP＋CRP），rf 5.00%、β＝Damodaran 美國半導體產業無槓桿 β（現金修正 1.50）以市值 D/E 重新加槓桿（Hamada；Andy 2026-11-05 21:21）、ERP 5%、CRP 0），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；"
        "EV 加淨現金（現金 − 債務本金 − 租賃）÷ 稀釋股數，以 ke 推到 12 個月後並扣期間股利（每季 0.65 美元 × 4）。新增 DCF 頁（共用模組 tools/fcff_dcf.py），Val 0.1 節改為 WACC 與 DCF 彙總，Sens 目標價參數改為標準版；Wattonomics 錨定案例路徑改用同一 FCFF 式；"
        "刪除 tp_r、tp_roe，tp_g 只留給 PL 延伸期收斂。PL、Q、Units、Comparable、情境權重不變。"); _r += 1
RD.cell(_r, 1, "v0.4.1（2026-11-05）：只改展示層——README 頁新增「展示層」「命題回覆」「本版要點」三段（公式帶數字），HTML 改用個股 HTML 版面規格 v1.0（tools/company_html/relayout.py）；README 以外逐格與 v0.4 相同，敏感度沿用。"); _r += 1
for _p in ("版本：v0.4", "2026-10-29", "v0.2.1", "v0.2（", "v0.1.1"):
    RD.cell(_r, 1, f"=A{_fr(RD, _p)}"); _r += 1
from openpyxl.styles import Alignment as _AL
for _rr in range(_r0, _r): RD.cell(_rr, 1).alignment = _AL(wrap_text=True, vertical="top")
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
_r3_fix(wb, "v0.5.3", "FY2026 Q4＋FY2027 前兩個月", extra=[(r"自訂（路徑 stub_fcff）", "FY2026 FCFF × 91／365＋FY2027 FCFF × 60／365"), (r"（第一年 FCFF × stub_frac）", "（FY2026 Q4＋FY2027 前兩個月 FCFF，按天數）"), (r"樣板標籤 A 樂觀／B 基準／C 下檔沿用", "情境名稱 A 順利／B 平平／C 不如預期"), (r"與股價同日（2026-10-23）", "rf 沿用 2026-10-23（股價與同業報價已換成 2026-11-05 收盤，rf 不變）"), (r"；現價約為 B 情境 FY2028 EPS 的 18\.0 倍", ""), (r'"（2026-10-23）"', '"（2026-11-05）"')])
_r3_static(wb, "v0.5.2（2026-11-05）", "v0.5.2（2026-11-05）：換用共用模組 tools/fcff_dcf.py v2（工作單第二輪修正；Andy 2026-11-05 23:16 定案）。(1) Stub 期：資產負債表日 2026-08-02（10-Q，Q3 FY26，最新已公布一季）與評價日 2026-12-31 不同，淨現金滾到評價日＝資產負債表淨現金 -38.43＋stub 期 FCFF（各情境 FY2027 FCFF × stub_frac 0.414 年；情境 B 29.93）− 已宣告股利 3.21（每股 0.65 美元，2026-09-30 發放；Q4 FY26 股利尚未宣告、買回無已宣告金額，皆不扣）＝-11.70（US$B）。DCF 目標價 144.06 → 150.27、加權目標價 182.86 → 185.96（上檔 -42.2%，放空）。(2) β 敏感度：低／高改為 β_u ∓ 0.3 各自重新加槓桿（In 頁 beta_du；beta_u 低／高欄刪除）。(3) ke 說明格改由模組參數 beta_note 提供；In 頁 M2 標題不寫版號。資本支出、折舊：本模型沒有逐年資本支出／折舊列（只有近三期比率），維持占營收比例。美國公司不加「美元口徑折現」案例（＝基準）。PL、Comparable、情境權重不變。")
# ---------------- v0.5.4（2026-11-06）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守（本資料夾 px1105_post.py；只改文字） ----------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import px1105_post  # noqa: E402
print(px1105_post.apply(wb, os.path.dirname(os.path.abspath(__file__)), "v0.5.4", "v0.5.3", "v0.5.3（2026-11-06）"))
wb.save(sys.argv[1])
if os.environ.get("AVGO_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "SHR": {s: {str(k): v for k, v in R.items()} for s, R in SHR.items()},
               "SPEC": {f"{k[0]}|{k[1]}": v for k, v in SPEC.items()}, "CAL": CAL, "CHK": CHK_ROW, "SENSN": SENS_N, "IMP": IMP, "SB": SB},
              open(os.environ["AVGO_ROWMAP"], "w"), ensure_ascii=False, indent=1)
