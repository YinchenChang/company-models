# 台積電模型 v0.2.1（AI 半導體 Project 步驟 4）：CoWoS 供需上限 × AI 營收由下而上 × 損益與估值
# 用法：python3 tsmc/build_tsmc.py raw.xlsx ；再以 LibreOffice headless 重算
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "tools"))
import valuation_rule as VR

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
YRS = (26, 27, 28, 29); SC = ("A", "B", "C")
PLAT = {"bw": {y: "GB" for y in YRS}, "rb": {y: "VR" for y in YRS}, "ru": {y: "RU" for y in YRS}, "hp": {y: "HP" for y in YRS},
        "amd": {26: "AMD8", 27: "AMDR", 28: "AMDR", 29: "AMDR"}, "tpu": {26: "TPU7", 27: "TPU8", 28: "TPU8", 29: "TPU8"},
        "trn": {y: "TRN" for y in YRS}, "oth": {y: "ASIC" for y in YRS}}
# 每片 CoWoS 晶圓可封裝的加速器數（良品），台積電內容（千美元／顆）
PP = {"GB": (14, 2.3, "Derived", "每片封裝數：36Kr 2026 NVIDIA 約 66 萬片 ÷ NVIDIA 約 7.6M 顆 ≈ 11.5（組合）；台積電內容：橋接表 v0.2（台積電內容與 v0.1 相同）（Silicon Analysts、Epoch：邏輯＋CoWoS 約 $1.9K，B300 封裝溢價）"),
      "VR": (8, 2.8, "Derived", "Morgan Stanley 2027 NVIDIA 122.2 萬片 ÷ 約 9.7M 顆 ≈ 7.9（組合）；內容：橋接表 v0.2（台積電內容與 v0.1 相同）"),
      "RU": (4, 6.0, "Assumed", "4 顆光罩尺寸晶粒＋16 疊 HBM4E"),
      "HP": (29, 1.1, "Analogy", "CoWoS-S 單晶粒；Silicon Analysts H100 邏輯 $300＋封裝 $750"),
      "AMD8": (8.5, 2.0, "Derived／Analogy", "36Kr 2026 AMD 約 9 萬片 ÷ 約 0.75M 顆"),
      "AMDR": (6, 3.0, "Assumed", "MI455X 12 疊 HBM4"),
      "TPU7": (28, 1.5, "Derived／Assumed", "36Kr：Broadcom 約 20 萬片、TPU 60–65% → 約 12.5 萬片 ÷ 約 3.85M 顆"),
      "TPU8": (24, 1.8, "Assumed", "TPU 8 HBM 288GB"),
      "TRN": (30, 1.2, "Derived／Assumed", "36Kr：世芯約 6 萬片（Trainium3）"),
      "ASIC": (20, 1.2, "Assumed", "")}

IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E）。金額：十億美元；台積電內容：千美元／顆；CoWoS：千片"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
sec("A. 加速器顆數（markets/memory，與 HBM、網通同一份）")
for k, nm, u, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u[i], None, tag, ("markets/memory：" + src) if i == 0 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", MEM.G29[k]["lo"], MEM.G29[k]["base"], MEM.G29[k]["hi"], MEM.G29[k]["tag"],
        f"Siliconomics SIF_g29_{k}（單一來源；2029 無平台別顆數預測，區間 ±10pt）" if k == MEM.UNITS[0][0] else "同上")
for s, (v28, v29) in zip(SC, [MEM.SCEN_UNITS[s_] for s_ in SC]):   # 單一來源：markets/memory/market.py（HBM 需求位元用同一組）
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, v28, None, "Assumed", "A＝基準；B＝AI 資本支出消化；C＝明顯下修" if s == "A" else "")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, v29, None, "Assumed", "")
sec("B. 每片 CoWoS 晶圓封裝數、每顆台積電內容")
for p, (pp, ct, tag, src) in PP.items():
    inp(f"pp_{p}", *SIL.row8(f"cw_pp_{p}")[1:])
    inp(f"ct_{p}", f"{p}：每顆台積電內容（邏輯晶圓＋CoWoS）", "$K／顆", round(ct * 0.8, 2), ct, round(ct * 1.2, 2), tag, "同上")
inp("cow_oth26", *SIL.row8("cw_oth26")[1:])
inp("cow_oth_g", *SIL.row8("cw_oth_g")[1:])
sec("C. CoWoS 產能（台積電，月產能年底值，千片）")
for y in (25, 26, 27, 28, 29):
    inp(f"cap_end{y}", *SIL.row8(f"cw_capend{y}")[1:])
inp("osat", *SIL.row8("cw_osat")[1:])
inp("ms_dem26", *SIL.row8("cw_msdem26")[1:])
inp("ms_dem27", *SIL.row8("cw_msdem27")[1:])
_hg = SIL.NAMES["SIF_HBM_GapEB_A"][2] / SIL.NAMES["SIF_HBM_SupplyEB"][2]   # 2027（欄序 2025..2029）
inp("hbm_gap27", "HBM 2027 需求缺口 ÷ 供給（Siliconomics）", "%", None, _hg, None, "Derived",
    f"Siliconomics SIF_HBM_GapEB_A ÷ SIF_HBM_SupplyEB（2027）：需求 {SIL.NAMES['SIF_HBM_DemandGrossEB'][2]:.2f} EB vs 供給 {SIL.NAMES['SIF_HBM_SupplyEB'][2]:.2f} EB")
sec("D. 台積電營收")
inp("fx", "美元兌新台幣（2027 起）", "NT$/US$", 30, SIL.base("fx_twd"), 33, "Assumed", "v0.1.1 標 Verified（當時 Q3 指引匯率）；4Q26 指引匯率為 31.8，2027 起未獲公司指引，數值不變")
inp("rev25", "營收 2025", "$B", None, 122.43, None, "Verified", "四季加總")
PR3 = "公司：3Q26 新聞稿 pr.tsmc.com/english/news/3346（6-K a3q26e）｜媒體：見 Sources 頁"
inp("rev26_1q", "營收 1Q26", "$B", None, 35.90, None, "Verified", "公司：1Q26 新聞稿；NT$1,134.10B")
inp("rev26_2q", "營收 2Q26", "$B", None, 40.20, None, "Verified", "公司：2Q26 新聞稿；NT$1,270.38B")
inp("rev26_3q", "營收 3Q26", "$B", None, 46.72, None, "Verified", PR3 + "｜NT$1,494.24B；多數媒體 46.7（個別媒體 46.13、46.87 與公司數不符，不採）")
inp("rev26_4q", "營收 4Q26（指引低／中／高）", "$B", 52.3, 53.3, 54.3, "Interested-party", "公司 4Q26 指引 US$52.3–54.3B（匯率 31.8）｜媒體：BigGo、民視財經；中點為本模型取值，公司歷來常高於指引上緣")
inp("fx26_4q", "4Q26 指引匯率", "NT$/US$", None, 31.8, None, "Interested-party", "公司 4Q26 指引假設")
inp("ntd26_1q", "營收 1Q26（新台幣）", "NT$B", None, 1134.10, None, "Verified", "公司：1Q26 新聞稿")
inp("ntd26_2q", "營收 2Q26（新台幣）", "NT$B", None, 1270.38, None, "Verified", "公司：2Q26 新聞稿")
inp("ntd26_3q", "營收 3Q26（新台幣）", "NT$B", None, 1494.24, None, "Verified", PR3)
inp("ntd26_4q", "營收 4Q26（新台幣；指引美元 × 指引匯率）", "NT$B", "=rev26_4q_lo*fx26_4q", "=rev26_4q*fx26_4q", "=rev26_4q_hi*fx26_4q", "Derived", "")
inp("gm26_1q", "毛利率 1Q26", "%", None, 0.662, None, "Verified", "公司：1Q26 新聞稿")
inp("gm26_2q", "毛利率 2Q26", "%", None, 0.677, None, "Verified", "公司：2Q26 新聞稿")
inp("gm26_3q", "毛利率 3Q26", "%", None, 0.664, None, "Verified", PR3 + "｜2Q26 的 67.7% 下滑 1.3 個百分點，公司歸因 2 奈米初期爬坡稀釋")
inp("gm26_4q", "毛利率 4Q26（指引低／中／高）", "%", 0.65, 0.66, 0.67, "Interested-party", "公司 4Q26 指引 65–67%｜媒體：民視財經（中值 66%）")
inp("opm26_1q", "營業利益率 1Q26", "%", None, 0.581, None, "Verified", "公司：1Q26 新聞稿")
inp("opm26_2q", "營業利益率 2Q26", "%", None, 0.603, None, "Verified", "公司：2Q26 新聞稿")
inp("opm26_3q", "營業利益率 3Q26", "%", None, 0.587, None, "Verified", PR3)
inp("opm26_4q", "營業利益率 4Q26（指引低／中／高）", "%", 0.565, 0.575, 0.585, "Interested-party", "公司 4Q26 指引 56.5–58.5%｜媒體：BigGo")
inp("ni26_1q", "淨利 1Q26", "NT$B", None, 572.48, None, "Verified", "公司：1Q26 新聞稿（EPS NT$22.08）")
inp("ni26_2q", "淨利 2Q26", "NT$B", None, 706.56, None, "Verified", "公司：2Q26 新聞稿（EPS NT$27.25）")
inp("ni26_3q", "淨利 3Q26", "NT$B", None, 763.06, None, "Verified", PR3 + "（EPS NT$29.43）")
inp("eps26_1to3q", "EPS 1–3Q26 實績合計", "NT$", None, "=22.08+27.25+29.43", None, "Verified", "公司：各季新聞稿 22.08、27.25、29.43")
inp("hpc_3q", "HPC 占營收 3Q26（備忘，未進入公式）", "%", None, 0.68, None, "Verified", "公司：3Q26 簡報／逐字稿，2Q26 為 66%｜媒體：vestedfinance、Quest Review")
inp("g26_guid", "公司 2026 全年美元營收成長指引", "%", None, 0.44, None, "Interested-party", "公司：由「略高於 40%」上修至約 44%｜媒體：StockFeel、民視財經；僅作 Recon 對照，不驅動營收")
inp("rev26", "營收 2026（1–3Q 實績＋4Q 指引中點；低／高＝4Q 指引下／上緣）", "$B", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q_lo", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q_hi", "Derived", "v0.1.1 為 172（全年指引「略高於 40%」推算）")
inp("ntd26", "營收 2026（新台幣）", "NT$B", None, "=ntd26_1q+ntd26_2q+ntd26_3q+ntd26_4q", None, "Derived", "")
inp("fx_26", "2026 平均匯率（新台幣營收 ÷ 美元營收）", "NT$/US$", None, "=ntd26/rev26", None, "Derived", "僅用於 2026 EPS；2027 起用 fx")
inp("ai_oth", "其他 AI 相關 ÷ 加速器內容（網通 ASIC、HBM 基底晶粒、CPU）", "%", 0.0, 0.2, 0.4, "Assumed", "使 2026 AI 營收約等於台積電口徑推算（2025 十幾%後段 × CAGR 約 57%）")
for y, v in ((27, 0.12), (28, 0.08), (29, 0.05)):
    inp(f"px{y}", f"AI 內容價格調升 20{y}", "%", None, v, None, "Interested-party／Derived", "2027 先進製程 +5–10%（Nikkei、MS）、CoWoS 混合單價 +23%（JPMorgan）" if y == 27 else "Assumed")
for s, vals in zip(SC, ((0.12, 0.10, 0.08), (0.12, 0.05, 0.03), (0.12, 0.0, -0.03))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"na{y}_{s}", f"非 AI 營收成長 20{y}（{s}）", "%", None, v, None, "Assumed", "公司 2024–29 全公司 CAGR 約 25%、AI 約 55–59% → 非 AI 約 10–15%" if (y == 27 and s == "A") else "")
sec("D2. 依平台營收（官方報告口徑：HPC、智慧型手機、物聯網、車用、DCE、其他；v0.4 依個股模型原則）")
MR = "公司 Management Report（investor.tsmc.com）：Net Revenue by Platform"
PLAT_S = (("hpc", "高效能運算（HPC）"), ("sp", "智慧型手機"), ("iot", "物聯網"), ("au", "車用"), ("dce", "消費性電子（DCE）"))
for k, nm in PLAT_S:
    v25 = {"hpc": 0.58, "sp": 0.29, "iot": 0.05, "au": 0.05, "dce": 0.01}[k]
    inp(f"pf25_{k}", f"{nm} 占營收 2025", "%", None, v25, None, "Verified", "公司 4Q25 法說逐字稿（CFO）：2025 年 HPC 58%、智慧型手機 29%、IoT 5%、車用 5%、DCE 1%")
    for q, vals in ((1, {"hpc": 0.61, "sp": 0.26, "iot": 0.06, "au": 0.04, "dce": 0.01}), (2, {"hpc": 0.66, "sp": 0.22, "iot": 0.05, "au": 0.04, "dce": 0.01}),
                    (3, {"hpc": 0.68, "sp": 0.21, "iot": 0.04, "au": 0.04, "dce": 0.01})):
        inp(f"pf26_{q}q_{k}", f"{nm} 占營收 {q}Q26", "%", None, vals[k], None, "Verified", MR + f"（{q}Q26）")
inp("pf25_oth", "其他 占營收 2025", "%", None, "=1-pf25_hpc-pf25_sp-pf25_iot-pf25_au-pf25_dce", None, "Derived", "1 − 其餘五平台（公司 2%）")
for q in (1, 2, 3):
    inp(f"pf26_{q}q_oth", f"其他 占營收 {q}Q26", "%", None, f"=1-pf26_{q}q_hpc-pf26_{q}q_sp-pf26_{q}q_iot-pf26_{q}q_au-pf26_{q}q_dce", None, "Derived", f"1 − 其餘五平台（公司 2%）")
for k, nm, v, src in (("hpc", "HPC：非 AI（CPU、PC、一般伺服器、其他 HPC）", 0.01, "2025 HPC +48%（含 AI）；PC、CPU 轉 N2／A16（Assumed）"),
                      ("sp", "智慧型手機", -0.02, "2025 +11%（HPC +48%），占比 2025 29% → 3Q26 21%（公司）"),
                      ("iot", "物聯網", 0.0, "2025 +15%（公司）"), ("au", "車用", 0.02, "2025 +34%（公司）；3Q26 季增 23%"),
                      ("dce", "消費性電子（DCE）", -0.03, "2025 持平（公司）；3Q26 季減 12%"), ("oth", "其他", 0.0, "—")):
    inp(f"dp_{k}", f"{nm} 年增相對非 AI 基準成長的差距（2027–29）", "pt", round(v - 0.02, 3), v, round(v + 0.02, 3), "Assumed", src + "；各平台年增＝非 AI 基準成長 na（情境別）＋差距")
sec("E. 利潤率、資本支出、股本、估值")
inp("gm26", "毛利率 2026（四季以新台幣營收加權）", "%", None, "=(ntd26_1q*gm26_1q+ntd26_2q*gm26_2q+ntd26_3q*gm26_3q+ntd26_4q*gm26_4q)/ntd26", None, "Derived", "v0.1.1 為 66%")
inp("opm26", "營業利益率 2026（四季以新台幣營收加權）", "%", None, "=(ntd26_1q*opm26_1q+ntd26_2q*opm26_2q+ntd26_3q*opm26_3q+ntd26_4q*opm26_4q)/ntd26", None, "Derived", "")
inp("opex_26", "營業費用 ÷ 營收 2026", "%", None, "=gm26-opm26", None, "Derived", "僅用於 2026；2027 起用 opex")
inp("ni_op_26", "淨利 ÷ 營業利益 2026（1–3Q 實績比，4Q 沿用）", "x", None, "=(ni26_1q+ni26_2q+ni26_3q)/(ntd26_1q*opm26_1q+ntd26_2q*opm26_2q+ntd26_3q*opm26_3q)", None, "Derived", "1Q 0.869、2Q 0.922、3Q 0.870；4Q 公司未指引淨利率，Assumed 沿用 1–3Q 比")
for s, vals in zip(SC, ((0.64, 0.63, 0.62), (0.64, 0.58, 0.56), (0.64, 0.54, 0.52))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"gm{y}_{s}", f"毛利率 20{y}（{s}）", "%", None, v, None, "Derived／Assumed", "N2 稀釋 3–4 個百分點、海外廠 2–3 個百分點（公司）對上漲價；長期目標 56% 以上" if (y == 27 and s == "A") else "")
inp("opex", "營業費用 ÷ 營收", "%", None, 0.075, None, "Derived", "2Q26 毛利率 67.7% − 營業利益率 60.3%")
inp("ni_op", "淨利 ÷ 營業利益（2027 起；低／高＝區間）", "x", "=ni_op_underlying", "=(1+nonop_hist)*(1-etr_new)", "=ni_op_2023", "Derived", "基準＝(1＋歷史營業外比率)×(1－新稅率)，推導見 NI_OI 頁；低 0.864＝2026 年扣除一次性利益後的內含水準；高 0.909＝2023 年年度值（稅率上升前、2022–2025 年度最高）。v0.1.1 為 0.92（含一次性利益的 2Q26 單季值，不再當上限）")
inp("shares", "股數", "十億股", None, 25.93, None, "Derived", "市值 ÷ 股價；3Q26 淨利 NT$763.06B ÷ EPS 29.43 ＝ 25.93，一致")
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 2385, None, "Verified", "markets/invest_map/data/prices.csv（Yahoo 股市、永豐、鉅亨；10-23 收盤未取得，依工作單用 10-22）；MarketScreener 共識頁 10-22 收盤同；v0.2.2 為 10-07 的 2,585（Wantgoo）")
for y, v, tag, src in ((26, 62, "Interested-party", "公司指引 US$60–64B 中點（7 月上修；10 月維持不變）｜媒體：民視財經、vestedfinance"), (27, 77, "Interested-party", "Citi；券商 73–81"), (28, 90, "Interested-party", "JPMorgan"), (29, 95, "Assumed", "")):
    inp(f"capex{y}", f"資本支出 20{y}", "$B", None, v, None, tag, src)
for i, (v, src) in enumerate(((15, "近谷底（2022 年約 11.9x）"), (20.5, "ADR 10 年中位數（GuruFocus）"), (32.7, "10 年最高（GuruFocus）")), 1):
    inp(f"pe_{i}", f"P/E {v}x", "x", None, v, None, "Interested-party", src)
inp("cons_rev27", "共識營收 2027", "$B", None, 229.0, None, "Interested-party", "MarketScreener NT$7,329B ÷ 32")
inp("cons_eps26", "共識 EPS 2026", "NT$", None, 107.6, None, "Interested-party", "MarketScreener")
inp("cons_eps27", "共識 EPS 2027", "NT$", None, 142.0, None, "Interested-party", "MarketScreener")
inp("ai25_share", "AI 加速器 ÷ 營收 2025", "%", None, 0.18, None, "Verified", "公司：十幾%後段（high-teens）")
inp("ai_cagr", "AI 加速器營收 CAGR（2024–29）", "%", None, 0.57, None, "Verified", "公司：50% 中後段，往高 50% 靠攏")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期；Val 頁第 0 節用）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）；三者合計應為 100%（Val 頁第 0 節列示）")
# v0.3（2026-10-25）：DCF＋Comparable 目標價與收斂法參數（接在 In 頁最後，不位移既有列與 um28／um29 等名稱；黃底）
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；美、台、日、韓同一值）")
inp("tp_g", "長期成長率（永續、收斂法 B 的 2032 營收年增）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）")
inp("tp_roe", "ROE（近三年平均，2023–2025）", "%", None, "=AVERAGE(0.2671,0.2994,0.3504)", None, "Derived", "stockanalysis 比率頁：2023 26.71%、2024 29.94%、2025 35.04%（單一來源）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）")
inp("nim10", "十年平均淨利率（2016–2025；歸屬母公司淨利 ÷ 營收）", "%", None,
    "=AVERAGE(334.25/947.94,343.15/977.45,351.13/1031.47,345.26/1069.99,517.89/1339.26,0.3758,0.449,0.3879,0.4054,0.451)", None, "Derived",
    "2021–2025：MarketScreener 歷史淨利率（37.58%、44.9%、38.79%、40.54%、45.1%；stockanalysis 37.32%–44.57% 口徑略異）；2016–2020：公司年報營收與歸屬母公司淨利（NT$B，本次未重新開啟原文，單一來源）")
for s, f, t in (("A", "=tp_g+0.02", "tp_g＋2pt"), ("B", "=tp_g", "＝tp_g"), ("C", "=tp_g-0.02", "tp_g−2pt")):
    inp(f"fade_g{s}", f"情境 {s}：2032 營收年增收斂目標（{t}）", "%", None, f, None, "Assumed", "工作單 2026-10-25 11:20 收斂法：2030 起由 2029 年增線性收斂，2032 達到")
for s, t in (("A", "＝情境 A 2029 淨利率"), ("B", "情境 B 2029 淨利率與十年平均的中點"), ("C", "＝十年平均淨利率")):
    inp(f"fade_m{s}", f"情境 {s}：2032 淨利率收斂目標（{t}）", "%", None, 0, None, "Derived", "PL 頁 2029、In 頁 nim10（公式於建置時連結）")
inp("cons_tp", "市場共識目標價（賣方平均；低／高＝最低／最高）", "NT$", 2650, 3390.43, 5000, "Interested-party", "MarketScreener 共識頁：35 位分析師平均 3,390.43（最低 2,650、最高 5,000），10-22 收盤頁面，2026-10-25 擷取")
inp("cons_feps27", "共識 EPS 2027（2026-10-25 擷取；第 0 節差距拆解用）", "NT$", None, 148.1, None, "Interested-party", "MarketScreener 預估頁（cons_eps27 的 142.0 為較早擷取，保留不動）")
inp("cons_feps28", "共識 EPS 2028（2026-10-25 擷取）", "NT$", None, 188.2, None, "Interested-party", "MarketScreener 預估頁")
TPKEYS = ["tp_wA", "tp_wB", "tp_wC", "tp_r", "tp_g", "tp_roe", "tp_w_dcf", "nim10", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC", "cons_tp", "cons_feps27", "cons_feps28"]
YEL = PatternFill("solid", fgColor="FFF2CC")
INROW = {}
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, src)
    INROW[key] = r
    if key in TPKEYS:
        for j in range(1, 9): IN.cell(r, j).fill = YEL
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGH", [12, 50, 10, 8, 10, 8, 20, 110]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

def u(k, y, s="A"):
    base = f"u_{k}_{y}" if y < 29 else f"u_{k}_28*(1+g29_{k})"
    if y == 28: return f"{base}*um28_{s}"
    if y == 29: return f"{base}*um29_{s}"
    return base

# ---------------- CoWoS：供需 ----------------
CW = wb.create_sheet("CoWoS")
CW["A1"] = "CoWoS — 需求（加速器顆數 ÷ 每片封裝數＋其他用途）vs 產能（台積電月產能年均 × 12 × (1＋委外)）；千片。2028–29 用情境 A 顆數"
hdr(CW, 3, ["項目", "2026", "2027", "2028", "2029", "說明"])
YC = {26: "B", 27: "C", 28: "D", 29: "E"}; CR = {}; cr = [4]
def cput(label, fs, key=None, fmt="#,##0", note="", bold=False):
    rr = cr[0]; CW.cell(rr, 1, label).font = Font(bold=bold)
    for j, f in enumerate(fs):
        if f in (None, ""): continue
        CW.cell(rr, 2 + j, f).number_format = fmt
    CW.cell(rr, 6, note)
    if key: CR[key] = rr
    cr[0] += 1
for k, nm, *_ in MEM.UNITS:
    cput(f"  {nm}", [f"={u(k, y)}/pp_{PLAT[k][y]}*1000" for y in YRS], key="w_" + k)
cput("  其他 CoWoS 用途", ["=cow_oth26"] + [f"={YC[y-1]}{{r}}*(1+cow_oth_g)" for y in (27, 28, 29)], key="w_o")
for y in (27, 28, 29): CW.cell(CR["w_o"], 2 + YRS.index(y), f"={YC[y-1]}{CR['w_o']}*(1+cow_oth_g)").number_format = "#,##0"
cput("需求合計", [f"=SUM({YC[y]}4:{YC[y]}{CR['w_o']})" for y in YRS], key="dem", bold=True)
cput("  其中 NVIDIA", [f"={YC[y]}{CR['w_bw']}+{YC[y]}{CR['w_rb']}+{YC[y]}{CR['w_ru']}+{YC[y]}{CR['w_hp']}" for y in YRS], key="nv")
cput("台積電產能（月產能年均 × 12）", [f"=(cap_end{y-1}+cap_end{y})/2*12" for y in YRS], key="cap")
cput("可用產能（含委外）", [f"={YC[y]}{CR['cap']}*(1+osat)" for y in YRS], key="capt")
cput("需求 ÷ 可用產能", [f"={YC[y]}{CR['dem']}/{YC[y]}{CR['capt']}" for y in YRS], key="util", fmt="0%", note=">100%＝CoWoS 為瓶頸")
cput("對照：HBM 需求缺口 ÷ 供給（2027）", ["", "=hbm_gap27"], fmt="0%", note="markets/memory；與上一列比，數字大者為較緊的瓶頸")
cput("對照：Morgan Stanley 全球需求", ["=ms_dem26", "=ms_dem27"], key="ms")
cput("  本模型 ÷ MS", [f"=B{CR['dem']}/B{CR['ms']}-1", f"=C{CR['dem']}/C{CR['ms']}-1"], fmt="0%")
cput("  NVIDIA：本模型 vs 36Kr 66 萬片（2026）、MS 122.2 萬片（2027）", [f"=B{CR['nv']}/660-1", f"=C{CR['nv']}/1222-1"], fmt="0%")
CW.column_dimensions["A"].width = 52; CW.column_dimensions["F"].width = 40

# ---------------- P&L：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 台積電損益（十億美元；EPS 新台幣）。AI 營收＝Σ 顆數 × 每顆內容 × 價格調升 × (1＋其他 AI)；非 AI＝前一年 × (1＋成長)"
cols = [("B", "2026", None, 26)]
L = iter("CDEFGHIJK")
for s in SC:
    for y in (27, 28, 29):
        cols.append((next(L), f"20{y} {s}", s, y))
hdr(PL, 3, ["項目"] + [c[1] for c in cols])
PR = {k: 4 + i for i, k in enumerate(["ai_acc", "ai", "nonai", "rev", "aish", "gm", "gp", "op", "ni", "eps", "pe_px", "capex", "fcf"])}
LAB = {"ai_acc": "AI 加速器內容（Σ 顆數 × 每顆內容 × 價格）", "ai": "AI 相關營收（含其他 AI）", "nonai": "非 AI 營收", "rev": "營收",
       "aish": "AI ÷ 營收", "gm": "毛利率", "gp": "毛利", "op": "營業利益", "ni": "淨利", "eps": "EPS（新台幣）", "pe_px": "目前股價 ÷ EPS",
       "capex": "資本支出", "fcf": "淨利 − 資本支出（不含折舊，粗估）"}
for k, rr in PR.items(): PL.cell(rr, 1, LAB[k])
def pxf(y):
    return {26: "1", 27: "(1+px27)", 28: "(1+px27)*(1+px28)", 29: "(1+px27)*(1+px28)*(1+px29)"}[y]
for col, lab, s, y in cols:
    s_ = s or "A"
    acc = "(" + "+".join(f"{u(k, y, s_)}*ct_{PLAT[k][y]}" for k, *_ in MEM.UNITS) + f")*{pxf(y)}"
    PL[f"{col}{PR['ai_acc']}"] = "=" + acc
    PL[f"{col}{PR['ai']}"] = f"={col}{PR['ai_acc']}*(1+ai_oth)"
    if y == 26:
        PL[f"{col}{PR['rev']}"] = "=rev26"
        PL[f"{col}{PR['nonai']}"] = f"={col}{PR['rev']}-{col}{PR['ai']}"
        PL[f"{col}{PR['gm']}"] = "=gm26"
    else:
        prev = "B" if y == 27 else [c[0] for c in cols if c[2] == s and c[3] == y - 1][0]
        PL[f"{col}{PR['nonai']}"] = f"={prev}{PR['nonai']}*(1+na{y}_{s})"
        PL[f"{col}{PR['rev']}"] = f"={col}{PR['ai']}+{col}{PR['nonai']}"
        PL[f"{col}{PR['gm']}"] = f"=gm{y}_{s}"
    PL[f"{col}{PR['aish']}"] = f"={col}{PR['ai']}/{col}{PR['rev']}"
    PL[f"{col}{PR['gp']}"] = f"={col}{PR['rev']}*{col}{PR['gm']}"
    sx = "_26" if y == 26 else ""
    PL[f"{col}{PR['op']}"] = f"={col}{PR['rev']}*({col}{PR['gm']}-opex{sx})"
    PL[f"{col}{PR['ni']}"] = f"={col}{PR['op']}*ni_op{sx}"
    PL[f"{col}{PR['eps']}"] = f"={col}{PR['ni']}*fx{sx}/shares"
    PL[f"{col}{PR['pe_px']}"] = f"=px/{col}{PR['eps']}"
    PL[f"{col}{PR['capex']}"] = f"=capex{y}"
    PL[f"{col}{PR['fcf']}"] = f"={col}{PR['ni']}-{col}{PR['capex']}"
    for k, rr in PR.items():
        PL[f"{col}{rr}"].number_format = "0.0%" if k in ("aish", "gm") else ("0.0" if k != "pe_px" else "0.0\"x\"")
PL.column_dimensions["A"].width = 44
for c in "BCDEFGHIJK": PL.column_dimensions[c].width = 10
# v0.4：依平台營收（第 26 列起；接在既有列之後，不移動既有儲存格）
PFR = {k: 27 + i for i, k in enumerate(["hpc", "hpc_ai", "hpc_n", "sp", "iot", "au", "dce", "oth", "chk"])}
PFL = {"hpc": "高效能運算（HPC）", "hpc_ai": "  HPC：AI 相關（＝AI 相關營收列）", "hpc_n": "  HPC：非 AI", "sp": "智慧型手機", "iot": "物聯網", "au": "車用",
       "dce": "消費性電子（DCE）", "oth": "其他", "chk": "檢查：平台合計 − 營收"}
PL.cell(26, 1, "依平台營收（官方報告口徑；十億美元）。2026＝各季營收 × 平台比重（4Q26 用指引中點 × 3Q26 比重）；2027 起各平台＝前一年 ×（1＋非 AI 基準成長＋平台差距）；AI 相關全部歸 HPC").font = BOLD
for k, rr in PFR.items(): PL.cell(rr, 1, PFL[k])
def pf26(k):
    return "=" + "+".join(f"rev26_{q}q*pf26_{min(q, 3)}q_{k}" for q in (1, 2, 3, 4))
for col, lab, s, y in cols:
    if y == 26:
        PL[f"{col}{PFR['hpc']}"] = pf26("hpc")
        PL[f"{col}{PFR['hpc_n']}"] = f"={col}{PFR['hpc']}-{col}{PFR['hpc_ai']}"
        for k in ("sp", "iot", "au", "dce", "oth"): PL[f"{col}{PFR[k]}"] = pf26(k)
    else:
        prev = "B" if y == 27 else [c[0] for c in cols if c[2] == s and c[3] == y - 1][0]
        PL[f"{col}{PFR['hpc_n']}"] = f"={prev}{PFR['hpc_n']}*(1+na{y}_{s}+dp_hpc)"
        PL[f"{col}{PFR['hpc']}"] = f"={col}{PFR['hpc_ai']}+{col}{PFR['hpc_n']}"
        for k in ("sp", "iot", "au", "dce", "oth"): PL[f"{col}{PFR[k]}"] = f"={prev}{PFR[k]}*(1+na{y}_{s}+dp_{k})"
        PL[f"{col}{PR['nonai']}"] = "=" + "+".join(f"{col}{PFR[k]}" for k in ("hpc_n", "sp", "iot", "au", "dce", "oth"))
    PL[f"{col}{PFR['hpc_ai']}"] = f"={col}{PR['ai']}"
    PL[f"{col}{PFR['chk']}"] = f"={col}{PFR['hpc']}+" + "+".join(f"{col}{PFR[k]}" for k in ("sp", "iot", "au", "dce", "oth")) + f"-{col}{PR['rev']}"
    for k, rr in PFR.items(): PL[f"{col}{rr}"].number_format = "0.00"
PL.cell(37, 1, "2025 實績（依平台；十億美元＝2025 營收 × 公司比重）").font = BOLD
for i, (k, nm) in enumerate(PLAT_S + (("oth", "其他"),)):
    PL.cell(38 + i, 1, f"2025：{nm}"); PL[f"B{38 + i}"] = f"=rev25*pf25_{k}"; PL[f"B{38 + i}"].number_format = "0.00"
PL.cell(44, 1, "檢查：2025 平台合計 − 2025 營收"); PL["B44"] = "=SUM(B38:B43)-rev25"; PL["B44"].number_format = "0.000"
PL.cell(18, 1, "註：B 欄 2026 ＝1–3Q 實績＋4Q 指引中點（匯率、營業費用、淨利比用 2026 專屬參數）；非 AI 營收＝營收 − AI（由下而上）。2027 起兩者分開推。")
# v0.3：2030–2032 收斂法欄（L–T，接在既有欄之後；上游顆數只到 2029，非驅動因子推導）＋營收年增、淨利率兩列（第 20、21 列）
FADEFILL = PatternFill("solid", fgColor="E7E6E6")
COLOF = {(c[2] or "-", c[3]): c[0] for c in cols}          # (情境, 年) → 欄
FCOL = {}
L2 = iter("LMNOPQRST")
for s in SC:
    for y in (30, 31, 32):
        FCOL[(s, y)] = next(L2); PL[f"{FCOL[(s, y)]}3"] = f"20{y} {s} 收斂"
        PL[f"{FCOL[(s, y)]}3"].font = BOLD; PL[f"{FCOL[(s, y)]}3"].fill = HDR
PR["g"], PR["nim"] = 20, 21
PL.cell(PR["g"], 1, "營收年增"); PL.cell(PR["nim"], 1, "淨利率（淨利 ÷ 營收）")
PL["B20"] = "=B7/rev25-1"; PL["B21"] = "=B12/B7"
for col, lab, s, y in cols:
    if y == 26: continue
    prev = "B" if y == 27 else COLOF[(s, y - 1)]
    PL[f"{col}20"] = f"={col}7/{prev}7-1"; PL[f"{col}21"] = f"={col}12/{col}7"
for col, *_ in cols:
    PL[f"{col}20"].number_format = "0.0%"; PL[f"{col}21"].number_format = "0.0%"
for s in SC:
    L9 = COLOF[(s, 29)]
    for k, y in enumerate((30, 31, 32), 1):
        c = FCOL[(s, y)]; p = L9 if k == 1 else FCOL[(s, y - 1)]
        PL[f"{c}20"] = f"=${L9}20+(fade_g{s}-${L9}20)*{k}/3"
        PL[f"{c}{PR['rev']}"] = f"={p}{PR['rev']}*(1+{c}20)"
        PL[f"{c}21"] = f"=${L9}21+(fade_m{s}-${L9}21)*{k}/3"
        PL[f"{c}{PR['ni']}"] = f"={c}{PR['rev']}*{c}21"
        PL[f"{c}{PR['eps']}"] = f"={c}{PR['ni']}*fx/shares"
        PL[f"{c}{PR['pe_px']}"] = f"=px/{c}{PR['eps']}"
        for rr, fm in ((20, "0.0%"), (PR["rev"], "0.0"), (21, "0.0%"), (PR["ni"], "0.0"), (PR["eps"], "0.0"), (PR["pe_px"], "0.0\"x\"")):
            PL[f"{c}{rr}"].number_format = fm; PL[f"{c}{rr}"].fill = FADEFILL
for c in "LMNOPQRST": PL.column_dimensions[c].width = 11
PL.cell(19, 1, "灰底＝2030–2032 收斂法（上游顆數只到 2029）：營收年增由 2029 線性收斂到 fade_g、淨利率收斂到 fade_m（In 頁目標價參數，2032 達到），非驅動因子推導；股數、匯率固定")
IN.cell(INROW["fade_mA"], 5, f"=PL!{COLOF[('A', 29)]}21")
IN.cell(INROW["fade_mB"], 5, f"=(PL!{COLOF[('B', 29)]}21+nim10)/2")
IN.cell(INROW["fade_mC"], 5, "=nim10")

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。隱含股價 ÷ 目前股價 − 1"
hdr(V, 3, ["方法", "倍數", "EPS 年度", "情境", "EPS", "隱含股價", "相對目前"])
vr = 4
for col, lab, s, y in cols:
    if y not in (27, 28): continue
    for i in (1, 2, 3):
        V.cell(vr, 1, f"P/E"); V.cell(vr, 2, f"=pe_{i}"); V.cell(vr, 3, f"20{y}"); V.cell(vr, 4, s)
        V.cell(vr, 5, f"=PL!{col}{PR['eps']}").number_format = "0.0"
        V.cell(vr, 6, f"=B{vr}*E{vr}").number_format = "#,##0"
        V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        vr += 1
V.cell(vr + 1, 1, "讀法：目前股價 ÷ 2027 EPS 即市場給的遠期 P/E；與 10 年中位 20.5x 比，看市場是否已提前反映 2027–28 的成長與情境 B／C 風險")
for c, w in zip("ABCDEFG", [8, 8, 10, 8, 10, 12, 10]): V.column_dimensions[c].width = w

# 估值規則區塊：十年 P/E 歷史與規則值（2027＝中段取中位數、區間 25–75 分位；另列 2028 情境 C × 75 分位）
# 歷史值原放在 markets/overview/extract.py（2026-10-23 移入）
VR.add_rule_block(V, vr + 4, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "FullRatio", "vals": [14.5, 17.71, 16.13, 25.62, 31.11, 29.18, 11.94, 19.46, 28.72, 29.03], "tag": "Interested-party"},
    {"label": "companiesmarketcap", "vals": [11.6, 14.8, 14.2, 23.5, 28.3, 27.2, 11.2, 19.3, 28.8, 28.3], "tag": "Verified"}],
    "cases": [
        {"case": "中段", "pos": "中段 → 中位數", "q": 50, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "低", "pos": "25 分位（區間下緣）", "q": 25, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "高", "pos": "75 分位（區間上緣）", "q": 75, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "2028 C", "pos": "參考：2028 情境 C × 75 分位", "q": 75, "basis_label": "2028 EPS（C）", "basis": f"=PL!J{PR['eps']}"}]}],
    "0.0", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位；台積電 2027 為中段，取中位數，區間為 25 與 75 分位。兩個來源各自取分位再平均、四捨五入到兩位。倍數歷史為輸入（藍字）。")
V.column_dimensions["A"].width = 36

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司口徑、共識（只作事後對照）"
hdr(RC, 3, ["項目", "本模型", "對照值", "差距", "對照來源"])
rec = [("2026 AI 相關營收 vs 公司口徑推算", f"=PL!B{PR['ai']}", "=rev25*ai25_share*(1+ai_cagr)", "2025 AI 加速器十幾%後段 × CAGR 57%（公司）"),
       ("2026 AI 加速器內容（不含其他 AI）vs 同上", f"=PL!B{PR['ai_acc']}", "=rev25*ai25_share*(1+ai_cagr)", "其他 AI 相關（ai_oth）為 Assumed"),
       ("2026 EPS vs 共識", f"=PL!B{PR['eps']}", "=cons_eps26", "MarketScreener"),
       ("2026 營收 vs 公司全年成長指引（2025 × (1＋44%)）", f"=PL!B{PR['rev']}", "=rev25*(1+g26_guid)", "公司 3Q26 法說：全年美元成長約 44%"),
       ("2026 EPS vs 1–3Q 實績＋4Q 指引推算", f"=PL!B{PR['eps']}", "=eps26_1to3q+ntd26_4q*opm26_4q*ni_op_26/shares", "公司各季 EPS 22.08／27.25／29.43＋4Q 指引中點"),
       ("2027 營收（A）vs 共識", f"=PL!C{PR['rev']}", "=cons_rev27", "MarketScreener"),
       ("2027 EPS（A）vs 共識", f"=PL!C{PR['eps']}", "=cons_eps27", "MarketScreener"),
       ("2026 CoWoS 需求 vs Morgan Stanley", f"=CoWoS!B{CR['dem']}", "=ms_dem26", "千片"),
       ("2027 CoWoS 需求 vs Morgan Stanley", f"=CoWoS!C{CR['dem']}", "=ms_dem27", "千片")]
for i, (nm, a, b, src) in enumerate(rec, 4):
    RC.cell(i, 1, nm); RC.cell(i, 2, a).number_format = "#,##0.0"; RC.cell(i, 3, b).number_format = "#,##0.0"
    RC.cell(i, 4, f"=B{i}/C{i}-1").number_format = "0%"; RC.cell(i, 5, src)
RC.column_dimensions["A"].width = 44; RC.column_dimensions["E"].width = 50

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 2027 情境 A：一次動一個輸入"
hdr(SE, 3, ["輸入", "2027 EPS", "相對基準", "2027 CoWoS 需求 ÷ 產能"])
base_eps = f"PL!C{PR['eps']}"; base_ai = f"PL!C{PR['ai']}"; base_rev = f"PL!C{PR['rev']}"
def eps_from(rev, gm="gm27_A"):
    return f"=({rev})*({gm}-opex)*ni_op*fx/shares"
items = [("基準", eps_from(base_rev), f"=CoWoS!C{CR['util']}"),
         ("Rubin 顆數 −30%", eps_from(f"{base_rev}-0.3*u_rb_27*ct_VR*(1+px27)*(1+ai_oth)"), f"=(CoWoS!C{CR['dem']}-0.3*CoWoS!C{CR['w_rb']})/CoWoS!C{CR['capt']}"),
         ("AI 內容價格 2027 +0%（不漲）", eps_from(f"{base_rev}-{base_ai}*px27/(1+px27)"), f"=CoWoS!C{CR['util']}"),
         ("毛利率 2027 60%", eps_from(base_rev, "0.60"), f"=CoWoS!C{CR['util']}"),
         ("毛利率 2027 67%", eps_from(base_rev, "0.67"), f"=CoWoS!C{CR['util']}"),
         ("非 AI 營收 2027 +0%", eps_from(f"{base_rev}-PL!B{PR['nonai']}*na27_A"), f"=CoWoS!C{CR['util']}"),
         ("CoWoS 2027 年底產能 17 萬片（KGI）", eps_from(base_rev), f"=CoWoS!C{CR['dem']}/((cap_end26+170)/2*12*(1+osat))")]
for i, (nm, e, ut) in enumerate(items, 4):
    SE.cell(i, 1, nm); SE.cell(i, 2, e).number_format = "0.0"; SE.cell(i, 3, f"=B{i}/B$4-1").number_format = "0%"; SE.cell(i, 4, ut).number_format = "0%"
SE.column_dimensions["A"].width = 36

SE.cell(13, 1, "淨利 ÷ 營業利益 對 EPS 的影響（情境 A）").font = BOLD
hdr(SE, 14, ["設定", "比值", "2027 EPS", "相對基準", "2028 EPS", "相對基準"])
for k, (nm, rr_) in enumerate((("基準（推導值）", "ni_op"), ("淨利 ÷ 營業利益 0.864（低）", "ni_op_lo"), ("淨利 ÷ 營業利益 0.909（高）", "ni_op_hi")), 15):
    SE.cell(k, 1, nm); SE.cell(k, 2, f"={rr_}").number_format = "0.000"
    SE.cell(k, 3, f"=PL!C{PR['op']}*B{k}*fx/shares").number_format = "0.0"
    SE.cell(k, 4, f"=C{k}/C$15-1").number_format = "0.0%"
    SE.cell(k, 5, f"=PL!D{PR['op']}*B{k}*fx/shares").number_format = "0.0"
    SE.cell(k, 6, f"=E{k}/E$15-1").number_format = "0.0%"
for c, w in zip("BCDEF", [10, 10, 10, 10, 10]): SE.column_dimensions[c].width = w



# ---------------- NI_OI：淨利 ÷ 營業利益拆解 ----------------
NO = wb.create_sheet("NI_OI")
NO["A1"] = "NI_OI — 淨利 ÷ 營業利益拆解（新台幣十億）。淨利÷營業利益 ＝ (1＋營業外÷營業利益) × (1－有效稅率) × 少數股權等調整。藍字＝輸入"
hdr(NO, 3, ["期間", "營業利益", "營業外收支", "稅前淨利", "所得稅", "淨利", "利息收入", "匯兌損益", "一次性項目", "有效稅率", "營業外÷營業利益", "少數股權等調整", "淨利÷營業利益", "獨立來源數", "標記", "來源／說明"])
AR22 = "https://investor.tsmc.com/static/annualReports/2022/english/index.html"
PER = [  # 期間, OI, nonop, tax, NI, interest, fx, oneoff, 來源數, 標記, 說明
 ("2022", 1121.28, 22.91, None, 1016.90, None, None, 0, 1, "Verified（單一來源）", "公司年報（經搜尋摘錄）；稅額＝稅前淨利－淨利（Derived）；媒體僅能核對淨利（歸屬母公司 1,016.53，股東會資料）；利息、匯兌明細未找到。待複核：另一搜尋結果的稅前淨利 1,140.7 疑為還原表欄位誤讀"),
 ("2023", 921.47, 57.71, 141.40, 837.77, 60.29, None, 0, 2, "Verified", "公司：2023 年報（6-K 2023-02）｜媒體：MoneyDJ 113 年損益表的前期欄；利息 60.29 僅 2025 20-F 一個來源；匯兌未找到"),
 ("2024", 1322.05, 83.79, 233.41, 1172.43, 87.21, 10.00, 0, 2, "Verified", "公司：2024 年報、2025 20-F｜媒體：MoneyDJ、中央社；匯兌 10.00 僅 20-F／財報一個來源"),
 ("2025", 1936.09, 105.56, 326.25, 1715.40, 105.74, 13.83, 0, 2, "Verified", "公司：2025 年報、2025 20-F｜媒體：工商時報（營業外約 1,056 億、稅前 2 兆 417 億、稅 3,263 億）；匯兌 13.83 僅 20-F／財報一個來源；其他利益及損失未拆"),
 ("1Q26", "=ntd26_1q*opm26_1q", 28.83, None, 572.48, None, None, 0, 1, "Derived", "營業利益＝新台幣營收×營益率（Derived）；營業外 28.83 取自 2Q26 簡報比較欄，欄位標籤未能直接確認；稅額＝稅前淨利－淨利（Derived）"),
 ("2Q26", "=ntd26_2q*opm26_2q", 95.83, 155.65, 706.56, None, None, 63.20, 2, "Verified／Derived", "公司：2Q26 簡報、3Q26 Management Report（稅 155.65、有效稅率 18.1%）；SEC 3Q26 6-K（含 2Q26 比較欄）核對營業利益、稅前淨利、淨利相符｜媒體：鉅亨（營業外 958.27 億）；營業利益 NT$ 欄為 Derived。一次性項目 63.20 的出處見下方「一次性利益出處」"),
 ("3Q26", 876.45, 34.59, 147.70, 763.06, None, None, 0, 2, "Verified", "公司：3Q26 Management Report（營業外 34.59、稅 147.70、有效稅率 16.2%；較 2Q26 降，主因無 VIS 處分與評價利益）｜SEC 3Q26 6-K a3q26e（營業利益、稅前淨利、淨利核對相符；使用者提供的核對）"),
]
r0 = 4
for i, (per, oi, no, tx, ni, it, fxg, one, nsrc, tag, note) in enumerate(PER):
    r = r0 + i
    NO.cell(r, 1, per)
    for j, v in ((2, oi), (3, no), (5, tx), (6, ni), (7, it), (8, fxg), (9, one)):
        if v is None: continue
        c = NO.cell(r, j, v)
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    NO.cell(r, 4, f"=B{r}+C{r}")
    if tx is None: NO.cell(r, 5, f"=D{r}-F{r}")
    NO.cell(r, 14, nsrc); NO.cell(r, 15, tag); NO.cell(r, 16, note)
rT = r0 + len(PER)          # 1–3Q26 合計列
NO.cell(rT, 1, "1–3Q26 合計")
for j in "BCDEFI": NO[f"{j}{rT}"] = f"=SUM({j}{rT-3}:{j}{rT-1})"
NO.cell(rT, 15, "Derived"); NO.cell(rT, 16, "三季加總；一次性項目＝2Q26 VIS 處分利益（含股權評價利益未另拆）")
for r in range(r0, rT + 1):
    NO[f"J{r}"] = f"=E{r}/D{r}"; NO[f"K{r}"] = f"=C{r}/B{r}"; NO[f"L{r}"] = f"=F{r}/(D{r}-E{r})"; NO[f"M{r}"] = f"=F{r}/B{r}"
    for c in "BCDEFGHI": NO[f"{c}{r}"].number_format = "#,##0.0"
    for c in "JK": NO[f"{c}{r}"].number_format = "0.0%"
    NO[f"L{r}"].number_format = "0.000"; NO[f"M{r}"].number_format = "0.000"
p0 = rT + 2
NO.cell(p0, 1, "基準值推導（In!ni_op）").font = BOLD
rows_ = [
 ("新稅率＝2024 起各期所得稅合計 ÷ 稅前淨利合計（2024、2025、1–3Q26）", f"=(E{r0+2}+E{r0+3}+E{rT})/(D{r0+2}+D{r0+3}+D{rT})", "0.0%", "etr_new",
  "取 2024 起：2024 年起全球最低稅負生效，有效稅率由 2022 年約 11% 升到 14–18%（2024 16.6%、2025 16.0%、2Q26 18.1%、3Q26 16.2%）；視為長期性變化"),
 ("歷史營業外比率＝2022–2025 加 1–3Q26（扣除一次性項目）的營業外合計 ÷ 營業利益合計", f"=(SUM(C{r0}:C{r0+3})+C{rT}-I{rT})/(SUM(B{r0}:B{r0+3})+B{rT})", "0.0%", "nonop_hist",
  "營業外收支主體是利息收入（2024 87.2、2025 105.7，占營業外 100% 以上）；匯兌 2024 +10.0、2025 +13.8，量級小且為收益；2Q26 VIS 處分利益 63.2 為一次性，已扣除"),
 ("基準＝(1＋歷史營業外比率) × (1－新稅率)", f"=(1+B{p0+2})*(1-B{p0+1})", "0.000", "ni_op_calc", "少數股權等調整各期約 1.000，略去"),
 ("對照：2024、2025 實績比（前版低值 0.887）", f"=(F{r0+2}+F{r0+3})/(B{r0+2}+B{r0+3})", "0.000", None, ""),
 ("對照：1–3Q26 實績比", f"=M{rT}", "0.000", None, "含 2Q26 一次性利益"),
 ("區間低值＝1–3Q26 扣除一次性利益（按新稅率稅後）", f"=(F{rT}-I{rT}*(1-B{p0+1}))/B{rT}", "0.000", "ni_op_underlying", "2026 年的內含水準；1Q26、3Q26 單季分別 0.869、0.871"),
 ("區間高值＝2023 年年度值（稅率上升前、2022–2025 年度最高）", f"=M{r0+1}", "0.000", "ni_op_2023", "2022 0.907、2023 0.909、2024 0.887、2025 0.886"),
 ("敏感：63.2 若非一次性（不扣除）→ 歷史營業外比率", f"=(SUM(C{r0}:C{r0+3})+C{rT})/(SUM(B{r0}:B{r0+3})+B{rT})", "0.0%", "nonop_hist_alt", "分子不扣 63.2"),
 ("敏感：63.2 若非一次性 → 基準", f"=(1+B{p0+8})*(1-B{p0+1})", "0.000", "ni_op_alt", "相對基準 0.875 的變化見 README"),
]
for k, (lab, f, fmt, nm, note) in enumerate(rows_, 1):
    rr = p0 + k
    NO.cell(rr, 1, lab); NO.cell(rr, 2, f).number_format = fmt; NO.cell(rr, 16, note)
    if nm: wb.defined_names[nm] = DefinedName(nm, attr_text=f"NI_OI!$B${rr}")
NO.cell(p0 + 14, 1, "一次性利益出處（2Q26 VIS 處分與評價利益 63.20）：公司 2Q26 合併財務報告（SEC 6-K）載明處分利益 NT$63,202,285 千元，其中含對剩餘持股改按公允價值衡量的非現金重新衡量利益 56,398,783 千元（喪失重大影響力）；公司 2026-05-15 公告表示近期無進一步出售計畫；3Q26 Management Report 說明營業外由 95.83 降至 34.59 主因無該項處分與評價利益；媒體：中央社、東森財經、鉅亨。頁碼：搜尋摘錄未含頁碼，未取得（文件與段落見 Sources 頁）。SEC 3Q26 6-K 本身未標註一次性項目。")
NO.cell(p0 + 12, 1, "判定：差距兩個來源都有——(a) 稅率由 2022 年 11% 升至 2024 年起約 16.6%，屬長期；(b) 營業外比率 2022 年 2.0% → 2023–25 年 5.5–6.3%（利息收入墊高），2Q26 另有一次性處分利益使 0.92 偏高。依規則取「新稅率 × 歷史平均營業外比率」。")
NO.column_dimensions["A"].width = 60; NO.column_dimensions["P"].width = 120; NO.column_dimensions["O"].width = 22
for c in "BCDEFGHIJKLMN": NO.column_dimensions[c].width = 11

# ---------------- Sources、README ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源"; SO["A1"].font = BOLD
srcs = [("台積電 2Q26 新聞稿", "https://pr.tsmc.com/english/news/3326"),
 ("台積電 2Q26 法說逐字稿", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-08/3e494f0c14dd0890f897aa044415e21d93486cc4/TSMC%202Q26%20Transcript.pdf"),
 ("台積電 4Q25 法說逐字稿", "https://investor.tsmc.com/japanese/encrypt/files/encrypt_file/reports/2026-01/51d09df96cd89ac19d65af39032b038dc2896a24/TSMC%204Q25%20Transcript.pdf"),
 ("台積電 2026-09 營收（6-K）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000680/tsm-revenue20261008.htm"),
 ("TrendForce CoWoS 2025-12-08", "https://www.trendforce.com/news/2025/12/08/news-tsmcs-cowos-l-s-reportedly-fully-booked-osat-partners-step-up-with-ases-cowop-in-focus"),
 ("TrendForce CoWoS 2026-06-15", "https://www.trendforce.com/news/2026/06/15/news-tsmc-cowos-supply-demand-gap-reportedly-seen-narrowing-from-20-to-10-by-end-2026-as-capacity-expands/"),
 ("經濟日報／Morgan Stanley 2026-09-13", "https://money.udn.com/money/story/5607/9752236"),
 ("36Kr CoWoS 分配 2026-01-09", "https://eu.36kr.com/en/p/3631956080985095"),
 ("經濟日報／JPMorgan 2026-09-29", "https://money.udn.com/money/amp/story/5607/9783752"),
 ("Tom's Hardware／Nikkei 2027 漲價", "https://www.tomshardware.com/tech-industry/semiconductors/tsmc-eyes-price-hikes-of-up-to-25-percent-on-chip-production-services-in-2027-report-claims-plans-to-raise-baseline-prices-by-5-percent-to-10-percent-on-advanced-nodes"),
 ("MarketScreener 共識", "https://in.marketscreener.com/quote/stock/TSMC-TAIWAN-SEMICONDUCTOR-6492349/valuation/"),
 ("GuruFocus P/E", "https://www.gurufocus.com/term/pettm/TSM"),
 ("【v0.2】台積電 3Q26 新聞稿（公司）", "https://pr.tsmc.com/english/news/3346"),
 ("【v0.2】台積電 3Q26 EarningsRelease（公司）", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-10/526b65af7305df2dde6b9ece6812b1fdf1e49f9b/3Q26%20EarningsRelease.pdf"),
 ("【v0.2】台積電 3Q26 6-K（含指引）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000726/a3q26e_withguidancexfinal.htm"),
 ("【v0.2】台積電 3Q26 逐字稿（公司）", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-10/eba4fac72873bf546b669bd2b5d973df5c93215d/TSMC%203Q26%20Transcript.pdf"),
 ("【v0.2】台積電 1Q26 新聞稿", "https://pr.tsmc.com/english/news/3297"),
 ("【v0.2】台積電 2Q26 6-K（含指引）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000451/a2q26e_withguidancexfinal.htm"),
 ("【v0.2】媒體：StockFeel 3Q26 法說重點", "https://www.stockfeel.com.tw/%E5%8F%B0%E7%A9%8D%E9%9B%BB-2330-2026q3-%E6%B3%95%E8%AA%AA%E6%9C%83/"),
 ("【v0.2】媒體：民視財經 3Q26 法說", "https://finance.ftvnews.com.tw/news/detail/2026A15W0572"),
 ("【v0.2】媒體：Futurum 3Q26", "https://futurumgroup.com/insights/tsmc-q3-fy-2026-earnings-beat-and-raise-on-ai-and-hpc-demand/"),
 ("【v0.2】媒體：BigGo Finance 3Q26", "https://finance.biggo.com/news/US_TSM_2026-10-15"),
 ("【v0.2】媒體：vestedfinance 3Q26", "https://vestedfinance.com/blog/us-stocks/tsmc-q3-earnings-2026/"),
 ("【v0.2】媒體：TrendForce 券商解讀 3Q26", "https://www.trendforce.com/news/2026/10/16/news-after-tsmcs-3q26-earnings-key-brokerage-takeaways-as-2028-capex-forecasts-reach-us105b/"),
 ("【v0.2】NI_OI：台積電 2024 年報", "https://investor.tsmc.com/static/annualReports/2024/english/index.html"),
 ("【v0.2】NI_OI：台積電 2025 年報", "https://investor.tsmc.com/static/annualReports/2025/english/index.html"),
 ("【v0.2】NI_OI：台積電 2025 20-F（利息、匯兌）", "https://investor.tsmc.com/sites/ir/sec-filings/2025_20F%20Report.pdf"),
 ("【v0.2】NI_OI：台積電 3Q26 Management Report", "https://investor.tsmc.com/schinese/encrypt/files/encrypt_file/qr/phase4_reports/2026-10/f53a64fd40fc0b453f94e077e5aba8cad506f27b/3Q26%20ManagementReport.pdf"),
 ("【v0.2】NI_OI：MoneyDJ 113 年（2024）損益表", "https://www.moneydj.com/kmdj/news/newsviewer.aspx?a=0c9c7a58-321f-4b5c-83c9-7f4837872969"),
 ("【v0.2】NI_OI：工商時報 2025 年報", "https://www.ctee.com.tw/news/20260416702066-430502"),
 ("【v0.2】NI_OI：台積電 2Q26 合併財務報告（SEC 6-K；VIS 處分利益 63,202,285 千元）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000541/a2026q2consolidatedreport-.htm"),
 ("【v0.2】NI_OI：台積電出售 VIS 股權公告（SEC 6-K，2026-05）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000280/tsmctosell81ofvanguardinte.htm"),
 ("【v0.2】NI_OI：台積電 3Q26 6-K（核對 2Q26、3Q26 營業利益、稅前淨利、淨利）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000726/a3q26e_withguidancexfinal.htm"),
 ("【v0.2】NI_OI：中央社（處分 VIS 1.52 億股）", "https://www.cna.com.tw/news/afe/202605180297.aspx"),
 ("【v0.2】NI_OI：東森財經（估獲利 632.02 億）", "https://fnc.ebc.net.tw/fncnews/stock/210975"),
 ("【v0.2】NI_OI：鏡報（2Q26 處分 VIS 利益）", "https://www.mirrordaily.news/story/79601"),
 ("【v0.2】NI_OI：鉅亨 2Q26 財報分析", "https://hao.cnyes.com/post/259466"),
 ("【v0.2】查證方式說明", "公司原文頁面在本次建置環境無法直接開啟（網路限制），3Q26 數字取自搜尋引擎對公司新聞稿的摘錄，並與至少兩則獨立媒體交叉；公司逐字稿原文待人工複核"),
 ("markets/chip_bridge v0.2、markets/memory（本 repo）", "")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 44; SO.column_dimensions["B"].width = 130
RD = wb.create_sheet("README", 0)
lines = [("台積電個股模型 v0.4（AI 半導體 Project 步驟 4）　2026-10-28　營收層改依官方平台（HPC、智慧型手機、IoT、車用、DCE、其他），AI 為 HPC 子列", True),
 ("模型命題：台積電在 AI 算力中的意義是供給上限而非分到的金額比例——AI 加速器只占其營收約兩成；2027 年加速器出貨受 HBM 而非 CoWoS 限制，台積電的獲利成長由「AI 顆數 × 每顆內容 × 漲價」與「非 AI 八成的穩定成長」共同決定。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("CoWoS：加速器顆數（markets/memory）÷ 每片封裝數 → CoWoS 需求；月產能年均 × 12 × (1＋委外) → 可用產能；需求 ÷ 產能 → 是否為瓶頸；與 HBM 缺口比較", False),
 ("PL：顆數 × 每顆台積電內容（橋接表）× 價格調升 × (1＋其他 AI) → AI 營收（HPC 子列）；非 AI 分為 HPC 非 AI、智慧型手機、IoT、車用、DCE、其他（PL 第 26 列起；2026 依各季平台比重，2027 起前一年 ×（1＋非 AI 基準成長＋平台差距））；× 毛利率 − 營業費用 → 營業利益 → 淨利 → EPS", False),
 ("情境：A 基準；B 2028 顆數 −20%、毛利率降至 58%；C 2028 顆數 −40%、毛利率 54%（AI 資本支出明顯下修）", False),
 ("Val：EPS × P/E（15x、10 年中位 20.5x、10 年最高 32.7x）", False), ("Recon、Sens：見各頁", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("v0.2.1 變更（2026-10-23）：Val 頁新增「十年倍數與規則值」區塊（十年 P/E 歷史、分位、規則值與隱含股價；原放在 markets/overview/extract.py），總覽與 HTML 展示層改讀此處；損益、EPS 與原有儲存格不變。", False),
 ("v0.2 變更（2026-10-22，3Q26 財報後）：2026 營收 172→176.1 $B（1–3Q 實績＋4Q 指引中點 53.3）；毛利率 66.0%→66.5%（四季加權）；EPS 114.2→112.1（新舊對照見 tsmc/README.md）。新增 2026 專屬參數 opex_26、ni_op_26、fx_26（來自實績）；2027 起淨利÷營業利益由 0.92 改為推導值（NI_OI 頁；區間 0.864–0.909）；毛利率、價格調升、非 AI 成長、CoWoS 產能全部未改（公司未給新的明確指引）。", False),
 ("公司 3Q26 說法未取得：AI 加速器營收占比、CoWoS 產能的新數字（搜尋未找到逐字稿原文），故 CoWoS 頁不動。", False),
 ("已知限制：(1) 台積電不揭露 CoWoS 片數與先進封裝營收，產能為券商估計；(2) 先進製程（N3／N2）晶圓產能未找到，未建晶圓瓶頸；(3) 其他 AI 相關營收比例為 Assumed；(4) 淨利 ÷ 營業利益 2027 起為單一推導值（見 NI_OI），2022 年與 1Q26 的明細為單一來源。", False)]
lines.append(("v0.2.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有儲存格）與 In 頁目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。各情境隱含價＝十年 P/E 中位數（規則值：2027 中段）× 各情境 2027 EPS；其餘儲存格不變。", False))
lines.append(("v0.3（2026-10-25）：Val 頁新增「0. 目標價」＝0.5 × DCF（股東現金流＝EPS ×（1 − g／ROE），FY2027–FY2032 明確期＋永續，r 10%、g 4%）＋0.5 × Comparable（聯電、格芯、英特爾、中芯港股、三星前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」。PL 頁三情境延伸到 2032（L–T 欄，2030 起為收斂法，非驅動因子推導）；新增營收年增、淨利率兩列。現價更新為 2026-10-22 收盤 2,385（10-23 未取得）。2026–2029 EPS 與 v0.2.2 相同；In 頁既有列與 um28／um29 等名稱未位移。", False))
lines.append(("情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。2027 三情境 EPS 完全相同（情境差從 2028 起，只差顆數係數與毛利率），DCF 情境照算，差異來自 2028 之後。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- Val 第 0 節：目標價（v0.3；Andy 2026-10-25 11:28 定版：DCF 與 Comparable 加權，11:53 權重各半；列標籤依工作單逐字） ----------------
# 為不位移既有儲存格，本節放在 Val 既有內容之後（原 v0.2.2「0. 目標價與上檔空間」的位置）；HTML 估值章第一個區塊即為本節。B／C／D 欄＝A 順利／B 平平／C 不如預期
RB0 = vr + 4                                   # 第 6 節起始列
PE0 = RB0 + 5                                  # 十年 P/E 表 2016 列
assert V.cell(PE0, 1).value == "2016" and V.cell(PE0 + 9, 1).value == "2025", (V.cell(PE0, 1).value, V.cell(PE0 + 9, 1).value)
PE10 = f"$C${PE0}:$C${PE0 + 9}"                # companiesmarketcap（Verified）
YRS6 = (27, 28, 29, 30, 31, 32)
def plc(s, y): return COLOF[(s, y)] if y <= 29 else FCOL[(s, y)]
T = {}
KEYF = PatternFill("solid", fgColor="FFF2CC")
def tset(r_, key, label, fs, fmt, note=""):
    V.cell(r_, 1, label); T[key] = r_
    for c_, f_ in zip("BCD", fs):
        if f_ is None: continue
        V[f"{c_}{r_}"] = f_; V[f"{c_}{r_}"].number_format = fmt
    if note: V[f"E{r_}"] = note
r_ = V.max_row + 2
V.cell(r_, 1, "0. 目標價").font = BOLD; T["sec0"] = r_; r_ += 1
V.cell(r_, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; r_ += 1
hdr(V, r_, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"]); T["h01"] = r_; r_ += 1
V.cell(r_, 1, "情境 EPS（FY2027–FY2032）"); V.cell(r_, 5, "PL 頁各情境（新台幣）；2030 起為收斂法（PL 頁灰底）；2027 三情境相同（情境差從 2028 起）"); r_ += 1
EPSR = {}
for y in YRS6:
    tset(r_, f"eps{y}", f"  FY20{y}", [f"=PL!{plc(s, y)}{PR['eps']}" for s in SC], "0.00"); EPSR[y] = r_; r_ += 1
V.cell(r_, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）")
V.cell(r_, 5, "台積電資本密集，但本模型沒有折舊推導（PL 頁只有資本支出），故用工作單通式 EPS ×（1 − g／ROE），未改用淨利＋折舊－資本支出"); r_ += 1
tset(r_, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）"); REINV = r_; r_ += 1
CFR = {}
for y in YRS6:
    tset(r_, f"cf{y}", f"  FY20{y}", [f"={c_}{EPSR[y]}*(1-$B${REINV})" for c_ in "BCD"], "0.00"); CFR[y] = r_; r_ += 1
c0, c9 = CFR[27], CFR[32]
tset(r_, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c_}{c0}:{c_}{c9},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c_}{c9}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c_ in "BCD"], "#,##0",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027"); r_ += 1
tset(r_, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c_}{T['dcfv']}*(1+tp_r)" for c_ in "BCD"], "#,##0"); r_ += 1
tset(r_, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{T['dcf12']}+tp_wB*C{T['dcf12']}+tp_wC*D{T['dcf12']}", None, None], "#,##0", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁"); r_ += 2
V.cell(r_, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; r_ += 1
hdr(V, r_, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
PEERS = [("聯電（UMC）", "UMC／2303.TW", 20.46, "2026-10-23"), ("格芯（GlobalFoundries）", "GFS", 20.04, "2026-10-22"),
         ("英特爾（Intel）", "INTC", 50.64, "2026-10-23"), ("中芯國際（SMIC，港股）", "0981.HK", 37.14, "2026-10-16"),
         ("三星電子", "005930.KS", 3.35, "2026-10-21")]
p0 = r_
for nm_, tk_, pe_, dt_ in PEERS:
    V.cell(r_, 1, nm_); V.cell(r_, 2, tk_); c = V.cell(r_, 3, pe_); c.font = BLUE; c.number_format = "0.0"
    V.cell(r_, 4, dt_); V.cell(r_, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(r_, 6, "Interested-party"); r_ += 1
p9 = r_ - 1; T["p0"], T["p9"] = p0, p9
tset(r_, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{p0}:C{p9})", None, None], "0.0",
     "晶圓代工與整合元件同業（聯電、格芯、英特爾、中芯港股、三星）；無 A 股。中芯以港股計（A 股 688981 不列）；三星 P/E 由記憶體盈餘主導、英特爾盈餘低谷，中位數受兩端影響小"); r_ += 1
tset(r_, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*(2/12*{c_}{EPSR[27]}+10/12*{c_}{EPSR[28]})" for s, c_ in zip(SC, "BCD")), None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）"); r_ += 1
tset(r_, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{T['feps']}", None, None], "#,##0", "同業中位數 × 前瞻 EPS"); r_ += 2
V.cell(r_, 1, "0.3 結論").font = BOLD; r_ += 1
hdr(V, r_, ["項目", "數值", "", "", "說明"]); T["h03"] = r_; r_ += 1
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], "#,##0", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁"); r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], "#,##0"); r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "MarketScreener：35 位分析師平均 3,390.43（最低 2,650、最高 5,000），10-22 收盤頁面，2026-10-25 擷取"); r_ += 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤；10-23 未取得）"); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/(2/12*cons_feps27+10/12*cons_feps28)-1", None, None], "0.0%",
     "本模型前瞻 EPS ÷ 共識前瞻 EPS（2／12 × 共識 2027＋10／12 × 共識 2028，In 頁 cons_feps27／28）− 1；Recon 頁另列"); r_ += 1
tset(r_, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/(2/12*cons_feps27+10/12*cons_feps28))/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{r_}"] = f'="共識目標價隱含 "&TEXT(cons_tp/(2/12*cons_feps27+10/12*cons_feps28),"0.0")&"x（÷ 共識前瞻 EPS），同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'; r_ += 1
tset(r_, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）"); r_ += 1
PRF = lambda c_: f"px/{c_}{EPSR[28]}"
tset(r_, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
     [f"=IF({PRF(c_)}>=MAX({PE10}),1,IF({PRF(c_)}<=MIN({PE10}),0,PERCENTRANK({PE10},{PRF(c_)})))" for c_ in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS，在 2016–2025 年底 P/E（第 6 節 companiesmarketcap）分布的分位（高於最大值記 100%、低於最小值記 0%）；只作檢查，不進目標價"); r_ += 1
V.cell(r_, 1, "目標價方法"); T["meth"] = r_
V[f"B{r_}"] = ("DCF：股東現金流＝EPS ×（1 − g／ROE），r 10%、g 4%、ROE 為 2023–2025 平均約 31%（再投資率約 13%），FY2027–FY2032 明確期＋永續；"
               "本模型無折舊推導，未改用淨利＋折舊－資本支出（2027–2029 資本支出 77–95 十億美元，現金流偏高估）；2030 起收斂法，情境 A 維持 2029 淨利率，偏高估。"
               "Comparable：聯電、格芯、英特爾、中芯（港股）、三星前瞻 P/E 中位數（聯電）× 12 個月後前瞻 EPS；同業成長與技術地位低於台積電，偏低估。"
               "情境對應：A＝順利（顆數基準）、B＝平平（2028 顆數 ×0.8、毛利率 58%）、C＝不如預期（顆數 ×0.6、毛利率 54%）；2027 三情境 EPS 相同，DCF 情境差來自 2028 起。")
r_ += 2
for k_ in ("dcftp", "comptp", "tpw"): V[f"B{T[k_]}"].fill = KEYF

# 舊法：規則值情境加權（參考；v0.3 起不再叫目標價）
MID = [rw for rw in range(1, V.max_row + 1) if V.cell(rw, 1).value == "規則值：中段（P/E）"]
assert len(MID) == 1, MID
EPS27 = {"A": "C", "B": "F", "C": "I"}   # PL 頁 2027 欄（表頭：2027 A／B／C）
assert all(PL[f"{c}3"].value == f"2027 {s}" for s, c in EPS27.items())
t0 = r_; T["old"] = t0
V.cell(t0, 1, "舊法：規則值情境加權（參考）").font = BOLD
hdr(V, t0 + 1, ["項目", "數值", "", "", "說明"])
V.cell(t0 + 2, 1, "規則值情境加權價（參考，非目標價）"); V[f"B{t0 + 2}"] = f"=tp_wA*B{t0 + 4}+tp_wB*B{t0 + 5}+tp_wC*B{t0 + 6}"; V[f"B{t0 + 2}"].number_format = "#,##0"
V[f"E{t0 + 2}"] = "權重 × 各情境隱含價；約為 2027 年底價格、未折現。v0.3 起目標價改為第 0 節（DCF 與 Comparable 加權）"
V.cell(t0 + 3, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）"); V[f"B{t0 + 3}"] = f"=B{t0 + 2}/px-1"; V[f"B{t0 + 3}"].number_format = "0.0%"
for i, (s, lab) in enumerate((("A", "情境 A（順利）隱含價"), ("B", "情境 B（平平）隱含價"), ("C", "情境 C（不如預期）隱含價"))):
    rr = t0 + 4 + i
    V.cell(rr, 1, lab); V[f"B{rr}"] = f"=$C${MID[0]}*PL!{EPS27[s]}{PR['eps']}"; V[f"B{rr}"].number_format = "#,##0"
    V[f"E{rr}"] = f"情境 {s} 2027 EPS（PL 頁）× 十年 P/E 中位數（第 6 節規則值：中段）"
rr = t0 + 7
V.cell(rr, 1, "情境權重 A／B／C"); V[f"B{rr}"] = "=tp_wA"; V[f"C{rr}"] = "=tp_wB"; V[f"D{rr}"] = "=tp_wC"
for c in "BCD": V[f"{c}{rr}"].number_format = "0%"
V[f"E{rr}"] = f'="合計 "&TEXT(B{rr}+C{rr}+D{rr},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
V.cell(t0 + 8, 1, "舊法說明")
V[f"B{t0 + 8}"] = ("各情境 2027 EPS × 十年（2016–2025 年底）P/E 中位數（FullRatio 與 companiesmarketcap 兩來源平均；規則值：2027 為中段）。"
                   "2027 三情境 EPS 相同，故三情境隱含價相同；隱含價約為 2027 年底的價格、未折現。")

# ---------------- Sens：目標價參數（tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.3／0.7） ----------------
TPS0 = 21
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；PL 頁收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def dcf_expr(c_, r="tp_r", g="tp_g"):
    reinv = f"MIN({g}/tp_roe,0.8)"
    cf = lambda y: f"Val!{c_}{EPSR[y]}*(1-{reinv})"
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"({pv}+{cf(32)}*(1+{g})/({r}-{g})/(1+{r})^6)"
TPCASES = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
for k_, (nm_, st_, kw) in enumerate(TPCASES):
    i_ = TPS0 + 2 + k_
    rr_, gg_, ww_ = kw.get("r", "tp_r"), kw.get("g", "tp_g"), kw.get("w", "tp_w_dcf")
    SE.cell(i_, 1, nm_); SE.cell(i_, 2, st_)
    SE[f"C{i_}"] = "=(" + "+".join(f"tp_w{s}*{dcf_expr(c_, rr_, gg_)}" for s, c_ in zip(SC, "BCD")) + f")*(1+{rr_})"
    SE[f"D{i_}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i_}"] = f"={ww_}*C{i_}+(1-{ww_})*D{i_}"
    SE[f"F{i_}"] = f"=E{i_}/E${TPS0 + 2}-1"
    for c_ in "CDE": SE[f"{c_}{i_}"].number_format = "#,##0"
    SE[f"F{i_}"].number_format = "0.0%"
TPS_LAST = TPS0 + 1 + len(TPCASES)
SE.cell(TPS_LAST + 1, 1, "核對：基準列 DCF 目標價 − Val 第 0 節（應為 0）"); SE[f"C{TPS_LAST + 1}"] = f"=C{TPS0 + 2}-Val!B{T['dcftp']}"; SE[f"C{TPS_LAST + 1}"].number_format = "0.000000"

# Recon：共識前瞻 EPS（第 0 節差距拆解用；接在既有列之後）
rn = RC.max_row + 1
for i, (nm, a, b, src) in enumerate([("2027 EPS（A）vs 共識（2026-10-25 擷取）", f"=PL!C{PR['eps']}", "=cons_feps27", "MarketScreener 預估頁；第 0 節差距拆解用"),
                                      ("2028 EPS（A）vs 共識（2026-10-25 擷取）", f"=PL!D{PR['eps']}", "=cons_feps28", "MarketScreener 預估頁")], rn):
    RC.cell(i, 1, nm); RC.cell(i, 2, a).number_format = "#,##0.0"; RC.cell(i, 3, b).number_format = "#,##0.0"
    RC.cell(i, 4, f"=B{i}/C{i}-1").number_format = "0%"; RC.cell(i, 5, src)

# ---------------- v0.3.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；新增列放在各頁現有最後一列之後，既有儲存格不變 ----------------
# 自由現金流（算力收支 2026-10-25 請求；地圖依列標籤讀取）：In 頁最後新增 da_rev、cx_rev；PL 頁最後新增一列（B–T 欄，三情境、2026–2032）；原「淨利 − 資本支出（粗估）」列保留不動
_ir = IN.max_row + 1
for _k, _nm, _v, _src in (("da_rev", "折舊 ÷ 營收（自由現金流列用）", "=681.649/3809.054", "2025 實績：折舊攤銷 681,649 百萬新台幣 ÷ 營收 3,809,054 百萬＝17.9%（stockanalysis 現金流量表、損益表）"),
                          ("cx_rev", "資本支出 ÷ 營收（只用於 2030–2032 收斂欄）", "=1272.411/3809.054", "2025 實績：資本支出 1,272,411 百萬新台幣 ÷ 營收＝33.4%（stockanalysis 現金流量表；約 US$40.9B，與公司公布一致）；2026–2029 用 PL 頁資本支出列（capex26–capex29）")):
    IN.cell(_ir, 1, _k); IN.cell(_ir, 2, _nm); IN.cell(_ir, 3, "%"); IN.cell(_ir, 5, _v).number_format = "0.0%"
    IN.cell(_ir, 7, "Assumed"); IN.cell(_ir, 8, _src); IN.cell(_ir, 9, "錨＝最近完整年度 2025 實績 [Verified]；以後各年比率不變 [Assumed]")
    wb.defined_names[_k] = DefinedName(_k, attr_text=f"In!$E${_ir}"); _ir += 1
_fr = PL.max_row + 1
PL.cell(_fr, 1, "自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）")
for col, lab, s, y in cols:
    PL[f"{col}{_fr}"] = f"={col}{PR['ni']}+{col}{PR['rev']}*da_rev-{col}{PR['capex']}"; PL[f"{col}{_fr}"].number_format = "0.0"
for (s, y), c in FCOL.items():
    PL[f"{c}{_fr}"] = f"={c}{PR['ni']}+{c}{PR['rev']}*(da_rev-cx_rev)"; PL[f"{c}{_fr}"].number_format = "0.0"; PL[f"{c}{_fr}"].fill = FADEFILL
PL.cell(_fr + 1, 1, "註：上一列折舊＝營收 × da_rev（In 頁，2025 實績 17.9%）；資本支出 2026–2029＝PL 頁資本支出列（In 頁 capex26–29），2030–2032 收斂欄＝營收 × cx_rev（2025 實績 33.4%）。十億美元。")
_sr = SO.max_row + 1
for _a, _b in (("[Verified] 自由現金流列：台積電 2025 折舊攤銷 681,649、資本支出 1,272,411、營收 3,809,054（百萬新台幣；stockanalysis）", "https://stockanalysis.com/quote/tpe/2330/financials/cash-flow-statement/ ；https://stockanalysis.com/quote/tpe/2330/financials/"),
               ("[Assumed] 折舊 ÷ 營收各年、資本支出 ÷ 營收（2030–2032 收斂欄）維持 2025 實績比率", "本模型假設")):
    SO.cell(_sr, 1, _a); SO.cell(_sr, 2, _b); _sr += 1
_U = f"B{T['upw']}"
_G = V.max_row + 2
V.cell(_G, 1, "上檔空間級距")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）"
V.cell(_G + 1, 1, "投資建議")
V[f"B{_G + 1}"] = f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))'
V[f"B{_G + 1}"].number_format = "@"
V[f"E{_G + 1}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空"
OBS = ["目標價 NT$2,446（DCF 2,150、Comparable 2,742 各半），比現價 2,385（2026-10-22）高 2.6%，級距 0%–25%；Comparable 單獨為 +15.0%。與共識目標價 3,390 的差距主要是 EPS：本模型 12 個月後前瞻 EPS 134.0 比共識前瞻 EPS 低 26.2%，倍數因素只 −8.7%（共識目標價隱含 18.7 倍，同業中位數 20.5 倍）。",
       "2027 三情境 EPS 相同（137.6），分歧從 2028 開始：A 157.6、B 129.5、C 106.3（顆數 ×1.0／0.8／0.6、B／C 毛利率 58%／54%），DCF 情境價值 2,769／2,041／1,496。2027 EPS 對非 AI 營收（持平 → −7.5%）與毛利率（60% → −7.1%、67% → +5.3%）最敏感。",
       "資本支出不隨情境下修（2028、2029 皆為 90、95 十億美元），情境 B、C 的自由現金流降得比淨利快：2028 自由現金流 A 84.8、B 57.4、C 34.1 十億美元（淨利 127.7／104.9／86.2）；DCF 用 EPS ×（1 − g／ROE）未扣資本支出，偏高估。下一個檢驗事件是 4Q26 財報（約 2027-01 中旬）的 2027 資本支出與毛利率指引；折現率 ±1pt 使加權目標價 −6.0%／+8.3%。"]   # 重點觀察（與 tsmc/README.md 相同文字）
_R = RD.max_row + 1
RD.cell(_R, 1, "v0.3.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。另依算力收支 2026-10-25 請求，PL 頁最後新增「自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）」一列（三情境、2026–2032；折舊＝營收 × 2025 實績 17.9%，資本支出 2026–2029 用既有資本支出列、2030–2032 用營收 × 2025 實績 33.4%），In 頁最後新增「折舊 ÷ 營收」（da_rev）與「資本支出 ÷ 營收」（cx_rev）；原「淨利 − 資本支出（不含折舊，粗估）」列保留不動，既有數值不變。")
RD.cell(_R + 1, 1, f'="投資建議："&Val!B{_G + 1}&"（依上檔空間固定規則）"')
for _i, _t in enumerate(OBS, 1): RD.cell(_R + 1 + _i, 1, f"重點觀察 {_i}：{_t}")
if os.environ.get("G1_ROWS"): print("G1ROWS", "tsmc", "grade", _G, "rec", _G + 1, "readme_obs", _R + 2, _R + 4)

# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "tsmc")
wb.save(sys.argv[1]); print("saved")
