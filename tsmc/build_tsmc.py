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
import fcff_dcf as FD   # v0.5：標準版 DCF（無槓桿自由現金流＋WACC）共用模組

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
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, v28, None, "Assumed", SIL.scen_note(s, 28, v28))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, v29, None, "Assumed", SIL.scen_note(s, 29, v29))
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
_hg = SIL.NAMES["SIF_HBM_GapEB_Base"][2] / SIL.NAMES["SIF_HBM_SupplyEB"][2]   # 2027（欄序 2025..2029）
inp("hbm_gap27", "HBM 2027 需求缺口 ÷ 供給（Siliconomics）", "%", None, _hg, None, "Derived",
    f"Siliconomics SIF_HBM_GapEB_Base ÷ SIF_HBM_SupplyEB（2027）：需求 {SIL.NAMES['SIF_HBM_DemandGrossEB'][2]:.2f} EB vs 供給 {SIL.NAMES['SIF_HBM_SupplyEB'][2]:.2f} EB")
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
inp("px", "股價（2026-11-05 收盤）", "NT$", None, 2240, None, "Verified", "Yahoo Finance 日線收盤 2026-11-05（2330.TW；v0.5.4 統一價格日）；v0.5.3 為 2026-11-04 的 2,260，v0.5.1 為 2026-10-22 的 2,385，v0.2.2 為 10-07 的 2,585（Wantgoo）")
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
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 樂觀／B 基準／C 保守；Val 頁第 0 節用）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）；三者合計應為 100%（Val 頁第 0 節列示）")
# v0.3（2026-10-25）：DCF＋Comparable 目標價與收斂法參數（接在 In 頁最後，不位移既有列與 um28／um29 等名稱；黃底）
inp("tp_g", "收斂法 B 的 2032 營收年增目標（PL 延伸期；v0.5 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）；v0.5 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數），本值只決定 2030–2032 營收收斂路徑")
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
# v0.5（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md）。評價日（折現基準日）沿用 v0.3 的 2026 年底；
# 金額單位：新台幣十億元（PL 頁美元 × fx 換算）；資產負債為 2026-06-30（stockanalysis 轉載公司季報）
SA_BS = "stockanalysis.com 台積電季度資產負債表（https://stockanalysis.com/quote/tpe/2330/financials/balance-sheet/?p=quarterly；2026-06-30 欄；轉載公司季報；2026-11-05 讀取）"
SA_CF = "stockanalysis.com 台積電年度現金流量表（https://stockanalysis.com/quote/tpe/2330/financials/cash-flow-statement/；2026-11-05 讀取）"
MR3 = "台積電 3Q26 Management Report（2026-10；investor.tsmc.com 季報專區；2026-11-05 讀取）"
FS3 = "台積電 2026 年第三季合併財務報表（investor.tsmc.com 中文財報 FS.pdf；2026-09-30 欄；2026-11-05 讀取）"
REV3 = "營收 2023 2,161.736、2024 2,894.308、2025 3,809.054（NT$B；公司年報）"
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.3 DCF 基準日）", ccy="NT$", amt="NT$B", shu="十億股",
        rf=(0.0198, "Trading Economics「Taiwan 10Y Bond Yield」：2026-11-04 1.98%（https://tradingeconomics.com/taiwan/government-bond-yield；2026-11-05 讀取）",
            "同頁：2026-09-24 1.96%（2008-11 以來最高）；股價日 2026-10-22 附近同水準", "無風險利率（新台幣 10 年期公債殖利率，2026-11-04）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：TW 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 臺灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：βL＝βU ×（1＋(1 − tax_w) × D/E）；Andy 2026-11-05 21:21）",
                    "Sens 低／高（Analogy；v0.5.1 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得"),
        rf_us=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；與 NVIDIA 模型同一來源）",
               "工作單第二輪修正（Andy 2026-11-05 23:16）：只用於 Sens「美元口徑折現」（rf＝美國 10 年期＋CRP 照基準、g＝MIN(3%, rf_us)）；基準維持當地利率＋固定匯率"),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=MR3 + "；" + FS3,
                  div=("=7*sh_v", "Derived", "台積電 1Q26 季度現金股利每股 7 元（2026-10 發放，落在 2026-09-30～2026-12-31）× 股數 sh_v；" + FS3 + "：應付現金股利 363,053 百萬元＝1Q26、2Q26 各每股 7 元 × 25.93 十億股",
                       "2Q26 股利（每股 7 元，董事會 2026-08-11 核准）於 2027-01 發放，已含在 dps_tgt（評價日後四次）；3Q26 股利尚未宣告；未扣買回（無已宣告金額）")),
        crp=(0.0078, "Interested-party", "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新）：Taiwan Moody's Aa3、國家風險溢酬 0.78%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）",
             "非美國市場：ERP＝5%＋CRP（Andy 2026-11-05 20:45）；不設股權成本下限"),
        kd=("=rf+0.005", "Assumed", "新台幣 10 年期公債＋信用利差 0.5%（台積電信用評等 AA 級；Assumed）", "債務占資本約 2%，對 WACC 影響可忽略"),
        tax_w=("=etr_new", "Derived", "＝NI_OI 頁 etr_new（2024、2025、1–3Q26 所得稅合計 ÷ 稅前淨利合計；全球最低稅負生效後）", "NOPAT 用同一稅率"),
        debt=(1031.58, "Verified", MR3 + "：Interest-bearing debts 1,031.58（NT$B；2026-09-30）", FS3 + "：一年內到期公司債及銀行借款 213,198＋應付公司債 770,315（NT$M），其餘為長期銀行借款；不含租賃"),
        lease=(36.882, "Verified", SA_BS + "：租賃負債流動 3,606＋非流動 33,276（NT$M；2026-06-30）", "3Q26 財報摘要未揭露租賃負債，沿用 2026-06-30（占淨現金不到 1%）"),
        cash=(4016.98, "Verified", MR3 + "：Cash & marketable securities 4,016.98（NT$B；2026-09-30。2026-06-30 為 3,518.01，與 v0.5 現金＋短期投資 3,517.787 同口徑）", FS3 + "：現金及約當現金 3,621,894（NT$M）"),
        inv=(224.633, "Verified", FS3 + "：長期投資 224,633（NT$M；含權益法投資與金融資產）", "帳面計入；Sens 測不計"),
        mi=("=nci_bk-nci_bk_l+nci_mv_guc+nci_mv_vis", "Derived", "＝非控制權益帳面 nci_bk − 上市子公司（創意 3443、采鈺 6789）的非控制權益帳面 nci_bk_l ＋ 兩家非控制權益市值（2026-11-05 收盤）",
            "v0.5.2（第三輪修正：上市子公司一律市值）；其餘子公司（JASM、ESMC 等未上市）用帳面；明細見 In 頁最後 nci_* 各列"),
        dps=("=7*4", "Derived", "台積電董事會 2026-08 核准 2Q26 季度現金股利每股 7 元（pr.tsmc.com/chinese/news/3332；經濟日報 money.udn.com/money/story/5607/9684950）",
             "評價日 2026-12-31 到目標時點 2027-12-31 之間假設四次、金額不變（Assumed）"),
        sh=("=shares", "Derived", "＝In 頁 shares（3Q26 淨利 ÷ EPS）", "評價用股數，不另推未來股數"),
        fc_da=(None, "=AVERAGE(526.883/2161.736,657.326/2894.308,681.649/3809.054)", None, "Derived",
               SA_CF + "：折舊攤銷 2023 526,883、2024 657,326、2025 681,649（NT$M）；" + REV3 + "；比例 24.4%、22.7%、17.9%，平均 21.7%", "In 頁 da_rev（2025 單年 17.9%）只用於 PL 自由現金流列"),
        fc_cx=(None, "=AVERAGE(949.817/2161.736,956.007/2894.308,1272.411/3809.054)", None, "Derived",
               SA_CF + "：資本支出 2023 949,817、2024 956,007、2025 1,272,411（NT$M）；比例 43.9%、33.0%、33.4%，平均 36.8%", "v0.5.1 起 FY2027–FY2029 改讀 PL 頁資本支出列（capex27–29 × fx）；本值只用於下一列穩態 fc_cx_ss 的定義；Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=MAX(fc_da_ss*(1+g_tv),(fc_cx+fc_da_ss*(1+g_tv))/2)", "Assumed", "FY2030 起由 FY2029（PL 頁資本支出列最後一年）的比例線性收斂，FY2032 達到：近期比例與維持水準（折舊 ×（1＋g））的中點，下限為維持水準", "營收成長在 2030–2032 收斂到 tp_g±2pt，資本密集度隨之下降（Andy 可改）"),
        fc_wc=(0.10, "=AVERAGE((201.938+250.997-57.204)/2161.736,(272.088+287.869-74.097)/2894.308,(281.791+288.11-84.224)/3809.054)", 0.20, "Derived",
               "近三年年底營運資金 ÷ 營收平均：營運資金＝應收帳款＋存貨 − 應付帳款（" + SA_BS.replace("季度", "") + "；2023 201,938＋250,997 − 57,204、2024 272,088＋287,869 − 74,097、2025 281,791＋288,110 − 84,224，NT$M）；" + REV3,
               "比例 18.3%、16.8%、12.8%；Sens 低／高 10%／20%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（自算 TW 半導體，Blume 調整後；採用值）", "x", None, "=0.67*beta_u_med+0.33", None, "Derived",
    "自算產業 β：TW 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 臺灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）；＝0.67 × 中位數 beta_u_med ＋ 0.33", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產")
inp("beta_u_med", "產業無槓桿 β 中位數（自算 TW 半導體，Blume 調整前）", "x", None, 1.64, None, "Derived",
    "自算產業 β：TW 半導體 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 臺灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）", "只用於計算 beta_u；v0.5.1 起 β 敏感度改 β_u ± beta_du（v0.5 為中位數 × 0.8 再 Blume／中位數未調整）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)；與 WACC 權重同一口徑")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代 v0.5 的低＝中位數 × 0.8 再 Blume、高＝中位數未調整")
TPKEYS = ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_w_dcf", "nim10", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC", "cons_tp", "cons_feps27", "cons_feps28"]
YEL = PatternFill("solid", fgColor="FFF2CC")
INROW = {}
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = srow[:8]
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    if len(srow) > 8 and srow[8]: IN.cell(r, 9, srow[8])
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
FD.yellow(IN, FD.JUDG + ("beta_u",))
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
 ("【v0.2】查證方式說明", "公司原文頁面當時未取得，3Q26 數字取自搜尋引擎對公司新聞稿的摘錄，並與至少兩則獨立媒體交叉；公司逐字稿原文待人工複核"),
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
lines.append(("情境對應：A＝樂觀、B＝基準、C＝保守（原為顆數情境）。2027 三情境 EPS 完全相同（情境差從 2028 起，只差顆數係數與毛利率），DCF 情境照算，差異來自 2028 之後。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- Val 第 0 節：目標價（v0.3；Andy 2026-10-25 11:28 定版：DCF 與 Comparable 加權，11:53 權重各半；列標籤依工作單逐字） ----------------
# 為不位移既有儲存格，本節放在 Val 既有內容之後（原 v0.2.2「0. 目標價與上檔空間」的位置）；HTML 估值章第一個區塊即為本節。B／C／D 欄＝A 樂觀／B 基準／C 保守
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
# v0.5：0.1 節改為標準版 DCF（FCFF＋WACC；tools/fcff_dcf.py）。DCF 明細另開 DCF 頁；本節列數與舊版相同（20 列＋空列），0.2 節以後位置不動
X01 = r_
DCFW = wb.create_sheet("DCF", wb.sheetnames.index("Val") + 1)
DCFW["A1"] = ("DCF — 標準版（v0.5.4；tools/fcff_dcf.py v2）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；金額為新台幣十億元（PL 頁美元 × fx）")
DCFW["A2"] = ("EBIT：2027–2029 取 PL 頁「營業利益」（營收 ×（毛利率 − 營業費用率））；2030–2032（PL 收斂欄只有淨利）由淨利還原：EBIT＝淨利 ÷ ni_op（PL 頁同一淨利 ÷ 營業利益比）。"
              "稅率＝NI_OI 頁 etr_new。營業外收支（主要是利息收入）不進 NOPAT，現金與投資另在評價時加入。折舊：近三年（2023–2025）占營收平均；資本支出：FY2027–FY2029 讀 PL 頁資本支出列（capex27–29，美元 × fx；三情境相同），FY2030 起由 FY2029 比例線性收斂到 fc_cx_ss（且 ≥ 折舊比例 ×（1＋g））；"
              "營運資金：近三年平均占營收 × 營收增量。淨現金：資產負債表日 2026-09-30，以 stub 期（含 stub 期會計年度 FY2026 FCFF × stub_frac − 已宣告 1Q26 股利；DCF 頁下方）滾到評價日；非控制權益：上市子公司（創意、采鈺）用 2026-11-05 市值，其餘帳面。")
FY_X = list(range(2027, 2033))
def _ebit(s_, y):
    c_ = plc(s_, y - 2000)
    return f"PL!{c_}{PR['op']}*fx" if y <= 2029 else f"=PL!{c_}{PR['ni']}/ni_op*fx"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s_] }）", rev={y: f"PL!{plc(s_, y - 2000)}{PR['rev']}*fx" for y in FY_X},
               rev0=f"(PL!B{PR['rev']}*fx)", ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "etr_new" for y in FY_X}, cx={y: f"capex{y - 2000}*fx" for y in (2027, 2028, 2029)},
               ebit_note="2030 起＝淨利 ÷ ni_op", stub_fcff="stub_fcff26*stub_frac") for s_ in SC]   # v0.5.2：stub＝含 stub 期會計年度 FY2026 FCFF × stub_frac（DCF 頁下方）
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 1.0, conv_from=2029, title="三情境 FCFF 明細（新台幣十億元；會計年度＝日曆年；評價日 2026-12-31）", fmt="#,##0.0", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="4016.98+224.633-1031.58-36.882-(45.326-(13.465*(1-0.35)+18.453*(1-0.6725))+6405*0.13401*(1-0.35)+360*0.3183*(1-0.6725))",
                             nc_note="2026-09-30 資產負債表（NT$B；3Q26 Management Report 與合併財報）：現金與有價證券 4,016.98＋長期投資 224.633 − 有息負債 1,031.58 − 租賃 36.882（2026-06-30）− 非控制權益（帳面 45.326，其中創意、采鈺改 2026-11-05 市值）") + 1
# v0.5.2（第三輪修正）：stub 期 FCFF＝含 stub 期的會計年度 FY2026 FCFF × stub_frac（v0.5.1 誤用 FY2027）；三情境相同（FY2026＝1–3Q 實績＋4Q 指引）
_sr0 = RD_["end"] + 1
DCFW.cell(_sr0, 1, "Stub 期 FCFF 錨點：含 stub 期的會計年度 FY2026 FCFF（新台幣十億元；三情境相同；× stub_frac 進淨現金滾到評價日）").font = BOLD
_s26 = [("營收 FY2026", f"=PL!B{PR['rev']}*fx_26", "PL 頁 2026 欄 × fx_26（1–3Q 實績＋4Q 指引）"),
        ("營業利益 EBIT FY2026", f"=PL!B{PR['op']}*fx_26", "PL 頁 2026 欄 × fx_26"),
        ("NOPAT＝EBIT ×（1 − etr_new）", f"=C{_sr0 + 2}*(1-etr_new)", "稅率同明細"),
        ("＋折舊攤銷（營收 × fc_da）", f"=C{_sr0 + 1}*fc_da", "同明細 FY2027 比例"),
        ("− 資本支出（capex26 × fx_26）", f"=capex26*fx_26", "In 頁 capex26（公司指引中點）"),
        ("− 營運資金增加（fc_wc ×（營收 − 2025 營收 3,809.054））", f"=fc_wc*(C{_sr0 + 1}-3809.054)", REV3),
        ("FY2026 FCFF", f"=C{_sr0 + 3}+C{_sr0 + 4}-C{_sr0 + 5}-C{_sr0 + 6}", "stub FCFF＝本值 × stub_frac")]
for _i, (_a, _f, _n) in enumerate(_s26, 1):
    DCFW.cell(_sr0 + _i, 1, _a); DCFW.cell(_sr0 + _i, 3, _f).number_format = "#,##0.0"; DCFW.cell(_sr0 + _i, 2, _n)
wb.defined_names["stub_fcff26"] = DefinedName("stub_fcff26", attr_text=f"DCF!$C${_sr0 + 7}")
RD_["end"] = _sr0 + 9
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 36
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 12
SDCF = FD.write_summary(wb, V, X01, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
for _j, _h in enumerate(["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], 1): V.cell(SDCF["h"], _j, _h)
T.update(h01=SDCF["h"], dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"])
# v0.5.3（第三輪修正第 12 點）：合併 EV 含子公司 100% → WACC 權重與 β 重新加槓桿的 E＝母公司市值＋非控制權益 mi_v（上市子公司市值、其餘帳面）
V.cell(SDCF["w_e"], 1).value = "權益占資本 E ÷（E＋D）（市值；E＝母公司市值＋非控制權益 mi_v；D＝債務＋租賃）"
V.cell(SDCF["w_e"], 2).value = "=(px*sh_v+mi_v)/(px*sh_v+mi_v+debt_v+lease_v)"
V.cell(SDCF["w_e"], 5).value = '="E "&TEXT(px*sh_v+mi_v,"#,##0")&"（含非控制權益 "&TEXT(mi_v,"#,##0.0")&"）、D "&TEXT(debt_v+lease_v,"#,##0.0")'
for _r in range(1, DCFW.max_row + 1):
    if str(DCFW.cell(_r, 1).value or "").startswith("檢查：WACC 由 In 頁輸入重算"):
        DCFW.cell(_r, 3).value = "=((px*sh_v+mi_v)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/(px*sh_v+mi_v+debt_v+lease_v)+wacc_adj-wacc"
        DCFW.cell(_r, 2).value = "E＝母公司市值＋非控制權益（第三輪修正第 12 點）"
IN.cell(INROW["beta_de"], 5).value = "=(debt_v+lease_v)/(px*sh_v+mi_v)"
IN.cell(INROW["beta_de"], 8).value = "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v ＋ 非控制權益 mi_v)；與 WACC 權重同一口徑（第三輪修正第 12 點：E 含非控制權益，上市子公司市值、其餘帳面）"
V.cell(SDCF["stub"], 5, '="資產負債表日 "&bs_date&"；stub_frac "&TEXT(stub_frac,"0.000")&" 年（含 stub 期會計年度 FY2026 FCFF × stub_frac；DCF 頁下方）"')
r_ = X01 + 21
V.cell(r_, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; r_ += 1
hdr(V, r_, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
# v0.5.2（第三輪修正）：同業前瞻 P/E 統一換算到 2026-11-04 收盤＝原 P/E × 11-04 收盤 ÷ 原報價日收盤（共識 EPS 不變）；原報價日收盤未取得者沿用原值並註明
# v0.5.4（2026-11-06，工作單 px1105）：同業前瞻 P/E 再換算到 2026-11-05 收盤＝現行 P/E（11-04 基準）× 11-05 收盤 ÷ 11-04 收盤（共識 EPS 不變）；
# 現行 P/E 不是 11-04 基準者（英特爾、中芯：原報價日收盤未取得）沿用原值並註明「未取得」
PEERS = [("聯電（UMC）", "UMC／2303.TW", 20.22, "2026-10-22", 125.5, 105.5, 98.4), ("格芯（GlobalFoundries）", "GFS", 20.04, "2026-10-22", 43.37, 40.34, 36.36),
         ("英特爾（Intel）", "INTC", 50.64, "2026-10-23", None, 58.27, 52.25), ("中芯國際（SMIC，港股）", "0981.HK", 37.14, "2026-10-16", None, 46.44, 47.5),
         ("三星電子", "005930.KS", 3.35, "2026-10-21", 223500, 218000, 201500)]
_YF5 = "Yahoo Finance 日線收盤 2026-11-05"
p0 = r_
for nm_, tk_, pe_, dt_, po_, pn_, p5_ in PEERS:
    V.cell(r_, 1, nm_); V.cell(r_, 2, tk_); c = V.cell(r_, 3, f"={pe_}*{pn_}/{po_}*{p5_}/{pn_}" if po_ else pe_); c.font = BLUE; c.number_format = "0.0"
    V.cell(r_, 4, "2026-11-05" if po_ else dt_)
    V.cell(r_, 5, (f"現行 P/E（11-04 基準）＝stockanalysis.com forward PE {pe_}（{dt_}，收盤 {po_:,}；stockanalysis 股價歷史）× 2026-11-04 收盤 {pn_:,} ÷ {po_:,}；再 × 2026-11-05 收盤 {p5_:,} ÷ 11-04 收盤 {pn_:,}（{_YF5}；共識 EPS 不變）" if po_ else
                   f"stockanalysis.com forward PE（分析師共識；{dt_}）；{dt_} 收盤未取得，未換算到 2026-11-05（2026-11-05 收盤 {p5_:,}，{_YF5}），沿用原值（未取得）")); V.cell(r_, 6, "Interested-party"); r_ += 1
p9 = r_ - 1; T["p0"], T["p9"] = p0, p9
tset(r_, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{p0}:C{p9})", None, None], "0.0",
     "晶圓代工與整合元件同業（聯電、格芯、英特爾、中芯港股、三星）；無 A 股。中芯以港股計（A 股 688981 不列）；三星 P/E 由記憶體盈餘主導、英特爾盈餘低谷，中位數受兩端影響小"); r_ += 1
tset(r_, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*(2/12*PL!{plc(s, 27)}{PR['eps']}+10/12*PL!{plc(s, 28)}{PR['eps']})" for s in SC), None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）"); r_ += 1
tset(r_, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{T['feps']}", None, None], "#,##0", "同業中位數 × 前瞻 EPS"); r_ += 2
V.cell(r_, 1, "0.3 結論").font = BOLD; r_ += 1
hdr(V, r_, ["項目", "數值", "", "", "說明"]); T["h03"] = r_; r_ += 1
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], "#,##0", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁"); r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], "#,##0"); r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "MarketScreener：35 位分析師平均 3,390.43（最低 2,650、最高 5,000），10-22 收盤頁面，2026-10-25 擷取"); r_ += 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-11-05 收盤）"); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/(2/12*cons_feps27+10/12*cons_feps28)-1", None, None], "0.0%",
     "本模型前瞻 EPS ÷ 共識前瞻 EPS（2／12 × 共識 2027＋10／12 × 共識 2028，In 頁 cons_feps27／28）− 1；Recon 頁另列"); r_ += 1
tset(r_, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/(2/12*cons_feps27+10/12*cons_feps28))/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{r_}"] = f'="共識目標價隱含 "&TEXT(cons_tp/(2/12*cons_feps27+10/12*cons_feps28),"0.0")&"x（÷ 共識前瞻 EPS），同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'; r_ += 1
tset(r_, "band", "價格帶（DCF 保守～樂觀）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）"); r_ += 1
PRF = lambda c_: f"px/PL!{plc(dict(zip('BCD', SC))[c_], 28)}{PR['eps']}"
tset(r_, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
     [f"=IF({PRF(c_)}>=MAX({PE10}),1,IF({PRF(c_)}<=MIN({PE10}),0,PERCENTRANK({PE10},{PRF(c_)})))" for c_ in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS，在 2016–2025 年底 P/E（第 6 節 companiesmarketcap）分布的分位（高於最大值記 100%、低於最小值記 0%）；只作檢查，不進目標價"); r_ += 1
V.cell(r_, 1, "目標價方法"); T["meth"] = r_
V[f"B{r_}"] = ("DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT(2033) ×（1 − g ÷ RONIC）÷（WACC − g），"
               "g＝MIN(3%, 新台幣 10 年期公債殖利率)、RONIC＝WACC；ke＝rf＋β ×（ERP 5%＋台灣國家風險溢酬），β 為 自算 TW 半導體產業無槓桿 β 依市值 D/E 重新加槓桿；EV 加淨現金與長期投資、減非控制權益 ÷ 股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁、tools/fcff_dcf.py）。"
               "折舊取 2023–2025 占營收平均；資本支出 FY2027–2029 用 PL 頁資本支出列（capex27–29），FY2030 起收斂；淨現金由 2026-09-30 以 stub 期滾到評價日；2030 起收斂法，情境 A 維持 2029 淨利率，偏高估。"
               "Comparable：聯電、格芯、英特爾、中芯（港股）、三星前瞻 P/E 中位數（2026-11-05 收盤換算）× 12 個月後前瞻 EPS；同業成長與技術地位低於台積電，偏低估。"
               "情境對應：A＝樂觀（原始顆數）、B＝基準（2028 顆數 ×0.8、毛利率 58%）、C＝保守（顆數 ×0.6、毛利率 54%）；2027 三情境 EPS 相同，DCF 情境差來自 2028 起。")
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
for i, (s, lab) in enumerate((("A", "情境 A（樂觀）隱含價"), ("B", "情境 B（基準）隱含價"), ("C", "情境 C（保守）隱含價"))):
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

# ---------------- Sens：目標價參數（v0.5 標準版 DCF：WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法、DCF 權重；tools/fcff_dcf.py） ----------------
TPS0 = 21
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt="#,##0",
                    title="目標價參數敏感度（標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=True)
assert SE.cell(TPS0 + 2 + 9, 1).value == "美元口徑折現", SE.cell(TPS0 + 2 + 9, 1).value
TPS_LAST = _tvs[1]
for _r in range(_tvs[0], _tvs[1] + 1):   # v0.5.2（第三輪修正第 16 點）：g 受 rf 上限（rf 1.98% < 3%），高案例＝rf
    if SE.cell(_r, 8).value == "g_tv": SE.cell(_r, 1).value = "永續成長率 g −1pt／上限 rf"

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
# v0.5.2（第三輪修正）：非控制權益——上市子公司（創意 3443、采鈺 6789）改市值，其餘帳面；股價統一 2026-11-04 收盤
SAQ = "stockanalysis.com 季度資產負債表（2026-06-30 欄；轉載公司季報；2026-11-06 讀取）"
for _k, _nm, _u, _v, _tag, _src in (
        ("nci_bk", "非控制權益（帳面，合併資產負債表）", "NT$B", 45.326, "Verified", FS3 + "：非控制權益 45,326（NT$M；2026-09-30）"),
        ("own_guc", "台積電持有創意（3443）股權比例", "%", 0.35, "Interested-party", "優分析（uanalyze.com.tw/articles/1626552582）：「創意為台積電持股約 35% 的子公司」；台積電具控制力而合併"),
        ("sh_guc", "創意流通股數", "十億股", 0.13401, "Verified", "stockanalysis.com/quote/tpe/3443/financials/balance-sheet/?p=quarterly：134.01M 股（2026-06-30）"),
        ("eq_guc", "創意股東權益（帳面）", "NT$B", 13.465, "Verified", SAQ.replace("季度", "創意季度") + "：13,465（NT$M）"),
        ("px_guc", "創意股價（2026-11-05 收盤）", "NT$", 6405, "Verified", "Yahoo Finance 日線收盤 2026-11-05（3443.TW）；v0.5.3 為 2026-11-04 的 7,115"),
        ("own_vis", "台積電持有采鈺（6789）股權比例", "%", 0.6725, "Verified", "台積電 2025 年報第八章關係企業（investor.tsmc.com/static/annualReports/2025/chinese/pdf/2025_tsmc_ar_c_ch8.pdf）：持有 213,619,000 股、67.25%"),
        ("sh_vis", "采鈺流通股數", "十億股", 0.3183, "Verified", "stockanalysis.com/quote/tpe/6789/financials/balance-sheet/?p=quarterly：318.3M 股（2026-06-30）；年報實收資本 3,182,991 千元 ÷ 10"),
        ("eq_vis", "采鈺股東權益（帳面）", "NT$B", 18.453, "Verified", SAQ.replace("季度", "采鈺季度") + "：18,453（NT$M）"),
        ("px_vis", "采鈺股價（2026-11-05 收盤）", "NT$", 360, "Verified", "Yahoo Finance 日線收盤 2026-11-05（6789.TW）；v0.5.3 為 2026-11-04 的 390"),
        ("nci_bk_l", "上市子公司非控制權益帳面（創意＋采鈺；自帳面總額扣除）", "NT$B", "=eq_guc*(1-own_guc)+eq_vis*(1-own_vis)", "Derived", "＝子公司股東權益 ×（1 − 台積電持股），2026-06-30 帳面（3Q26 子公司明細未取得）"),
        ("nci_mv_guc", "創意非控制權益市值", "NT$B", "=px_guc*sh_guc*(1-own_guc)", "Derived", "＝股價 × 股數 ×（1 − 台積電持股）"),
        ("nci_mv_vis", "采鈺非控制權益市值", "NT$B", "=px_vis*sh_vis*(1-own_vis)", "Derived", "＝股價 × 股數 ×（1 − 台積電持股）")):
    IN.cell(_ir, 1, _k); IN.cell(_ir, 2, _nm); IN.cell(_ir, 3, _u); IN.cell(_ir, 5, _v).number_format = "0.00%" if _u == "%" else "#,##0.000"
    if not (isinstance(_v, str) and _v.startswith("=")): IN.cell(_ir, 5).font = BLUE
    IN.cell(_ir, 7, _tag); IN.cell(_ir, 8, _src); IN.cell(_ir, 9, "v0.5.2 第三輪修正：上市子公司非控制權益一律市值")
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
V.cell(_G, 1, "上檔空間級距（六級，供分組）")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00），只供分組；投資建議用下一列五級規則"
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
# v0.5：重點觀察 2、3 改為公式（DCF 數字隨標準版 DCF 更新；下游 tools/target_summary 讀本列文字）
_q = lambda x: x.replace('"', '""')
_pe = lambda s_, y: f"PL!{COLOF[(s_, y)]}{PR['eps']}"
_d12 = lambda c_: f'TEXT(Val!{c_}{T["dcf12"]},"#,##0")'
RD.cell(_R + 3, 1, '="重點觀察 2：2027 三情境 EPS 相同（"&TEXT(' + _pe("A", 27) + ',"0.0")&"），分歧從 2028 開始：A "&TEXT(' + _pe("A", 28) + ',"0.0")&"、B "&TEXT(' + _pe("B", 28) + ',"0.0")&"、C "&TEXT(' + _pe("C", 28) + ',"0.0")'
                '&"（顆數 ×1.0／0.8／0.6、B／C 毛利率 58%／54%），標準版 DCF 情境價值（推到 12 個月後）"&' + _d12("B") + '&"／"&' + _d12("C") + '&"／"&' + _d12("D") +
                '&"。2027 EPS 對非 AI 營收（持平 → "&TEXT(Sens!C9,"+0.0%;-0.0%")&"）與毛利率（60% → "&TEXT(Sens!C7,"+0.0%;-0.0%")&"、67% → "&TEXT(Sens!C8,"+0.0%;-0.0%")&"）最敏感。"')
_fc = lambda s_: f'TEXT(PL!{COLOF[(s_, 28)]}{_fr},"0.0")'
_ni = lambda s_: f'TEXT(PL!{COLOF[(s_, 28)]}{PR["ni"]},"0.0")'
RD.cell(_R + 4, 1, '="重點觀察 3：資本支出不隨情境下修（2028、2029 皆為 "&TEXT(capex28,"0")&"、"&TEXT(capex29,"0")&" 十億美元），情境 B、C 的自由現金流降得比淨利快：2028 自由現金流 A "&' + _fc("A") + '&"、B "&' + _fc("B") + '&"、C "&' + _fc("C") +
                '&" 十億美元（淨利 "&' + _ni("A") + '&"／"&' + _ni("B") + '&"／"&' + _ni("C") + '&"）。v0.5.1 起標準版 DCF 的資本支出 FY2027–FY2029 直接用 PL 頁資本支出列（同上），FY2030 起收斂到 "&TEXT(fc_cx_ss,"0.0%")&"。'
                '下一個檢驗事件是 4Q26 財報（約 2027-01 中旬）的 2027 資本支出與毛利率指引；WACC −1pt／＋1pt 使加權目標價 "&TEXT(Sens!D' + str(TPS0 + 2) + ',"+0.0%;-0.0%")&"／"&TEXT(Sens!F' + str(TPS0 + 2) + ',"+0.0%;-0.0%")&"（以情境加權目標價為基準）。"')
if os.environ.get("G1_ROWS"): print("G1ROWS", "tsmc", "grade", _G, "rec", _G + 1, "readme_obs", _R + 2, _R + 4)

# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "tsmc")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.4.1", "2026-10-30")  # v0.4.1：In 頁版面（個股模型原則第 8 節；數值不變）
# v0.5.2（第三輪修正第 15、23 點）：情境驅動對照表頭統一情境名稱（v0.5.4 起改回「樂觀／基準／保守」，Andy 2026-11-06 03:14）；收斂法列名改中性
for _r in range(4, IN.max_row + 1):
    if str(IN.cell(_r, 1).value or "").startswith("情境驅動對照"):
        for _j, _h in zip((4, 5, 6), ("情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）")): IN.cell(_r + 1, _j).value = _h
    if IN.cell(_r, 1).value == "scn_fade_g": IN.cell(_r, 2).value = "2032 營收年增收斂目標"
    if IN.cell(_r, 1).value == "scn_fade_m": IN.cell(_r, 2).value = "2032 淨利率收斂目標"

# ---------------- Wattonomics 錨定（決策 S37；Siliconomics v0.13）：In 頁最後三列＋Sens 頁最下方小表與案例區塊；既有儲存格不變 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
_wr = IN.max_row + 2
IN.cell(_wr, 1, "Z. Wattonomics 錨定係數（敏感度用；Siliconomics v0.13）").font = BOLD; _wr += 1
_wrows = []; SIL.wta_rows(lambda *a: _wrows.append(a))
for _a in _wrows:
    for _j, _v in enumerate(_a, 1):
        if _v is not None: IN.cell(_wr, _j, _v)
    IN.cell(_wr, 5).font = BLUE
    wb.defined_names[_a[0]] = DefinedName(_a[0], attr_text=f"In!$E${_wr}"); _wr += 1
# Sens：AI 加速器內容與「其他 AI」皆與顆數成正比（PL：AI 相關＝Σ 顆數 × 內容 × 價格 ×（1＋ai_oth）），非 AI 不受顆數影響；故案例營收＝AI 相關（情境 B）× wta＋非 AI（情境 B）
W0 = SE.max_row + 3
SE.cell(W0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
SE.cell(W0 + 1, 1, WTA_NOTE)
SE.cell(W0 + 2, 1, "2030–2032：沿用 PL 頁延伸表收斂法（營收年增由案例 2029 年增線性收斂到 fade_gB、淨利率收斂目標＝（本案例 2029 淨利率＋nim10）÷ 2，與 In 頁 fade_mB 同規則；2032 達到）。目標價＝單一路徑（不做情境加權）：tp_w_dcf × DCF（Val 第 0 節同式）＋（1 − tp_w_dcf）× Comparable（tp_comp_adj × 本路徑前瞻 EPS ÷ 加權前瞻 EPS）")
hdr(SE, W0 + 3, ["案例", "2027 EPS", "相對基準", "相對情境 B", "2028 EPS", "相對基準", "相對情境 B", "目標價", "相對基準", "相對情境 B"])
WB0 = W0 + 9   # 案例區塊
WC = dict(zip((27, 28, 29, 30, 31, 32), "BCDEFG"))
SE.cell(WB0, 1, "案例區塊：Wattonomics 錨定（情境 B × wta；十億美元，EPS 新台幣）").font = BOLD
hdr(SE, WB0 + 1, ["項目", "2027", "2028", "2029", "2030 收斂", "2031 收斂", "2032 收斂", "說明"])
WR = {k: WB0 + 2 + i for i, k in enumerate(["wta", "ai_b", "ai", "nai", "rev", "gm", "ni", "g", "nim", "eps", "eps_b", "cf", "cf_b"])}
WLAB = {"wta": "Wattonomics 錨定係數", "ai_b": "AI 相關營收（情境 B，PL 頁）", "ai": "AI 相關營收 × wta", "nai": "非 AI 營收（情境 B，不變）", "rev": "營收",
        "gm": "毛利率（情境 B）", "ni": "淨利", "g": "營收年增", "nim": "淨利率", "eps": "EPS（新台幣）", "eps_b": "對照：情境 B EPS（PL 頁）",
        "cf": "營業利益 EBIT（v0.5 標準版 DCF 用；2030 起＝淨利 ÷ ni_op）", "cf_b": "對照：情境 B 營業利益 EBIT（PL 頁）"}
for k, rr in WR.items(): SE.cell(rr, 1, WLAB[k])
for y in (27, 28, 29):
    c, pc = WC[y], COLOF[("B", y)]
    SE[f"{c}{WR['wta']}"] = f"=wta_{y}"
    SE[f"{c}{WR['ai_b']}"] = f"=PL!{pc}{PR['ai']}"
    SE[f"{c}{WR['ai']}"] = f"={c}{WR['ai_b']}*{c}{WR['wta']}"
    SE[f"{c}{WR['nai']}"] = f"=PL!{pc}{PR['nonai']}"
    SE[f"{c}{WR['rev']}"] = f"={c}{WR['ai']}+{c}{WR['nai']}"
    SE[f"{c}{WR['gm']}"] = f"=gm{y}_B"
    SE[f"{c}{WR['ni']}"] = f"={c}{WR['rev']}*({c}{WR['gm']}-opex)*ni_op"
    SE[f"{c}{WR['g']}"] = f"={c}{WR['rev']}/" + ("PL!B" + str(PR['rev']) if y == 27 else f"{WC[y - 1]}{WR['rev']}") + "-1"
    SE[f"{c}{WR['nim']}"] = f"={c}{WR['ni']}/{c}{WR['rev']}"
for k_, y in enumerate((30, 31, 32), 1):
    c, p = WC[y], WC[y - 1]
    SE[f"{c}{WR['g']}"] = f"=$D{WR['g']}+(fade_gB-$D{WR['g']})*{k_}/3"
    SE[f"{c}{WR['rev']}"] = f"={p}{WR['rev']}*(1+{c}{WR['g']})"
    SE[f"{c}{WR['nim']}"] = f"=$D{WR['nim']}+(($D{WR['nim']}+nim10)/2-$D{WR['nim']})*{k_}/3"
    SE[f"{c}{WR['ni']}"] = f"={c}{WR['rev']}*{c}{WR['nim']}"
for y in YRS6:
    c = WC[y]
    SE[f"{c}{WR['eps']}"] = f"={c}{WR['ni']}*fx/shares"
    SE[f"{c}{WR['eps_b']}"] = f"=PL!{plc('B', y)}{PR['eps']}"
    SE[f"{c}{WR['cf']}"] = f"={c}{WR['rev']}*({c}{WR['gm']}-opex)" if y <= 29 else f"={c}{WR['ni']}/ni_op"
    SE[f"{c}{WR['cf_b']}"] = f"=PL!{plc('B', y)}{PR['op']}" if y <= 29 else f"=PL!{plc('B', y)}{PR['ni']}/ni_op"
for k, rr in WR.items():
    for c in "BCDEFG":
        if SE[f"{c}{rr}"].value is not None: SE[f"{c}{rr}"].number_format = "0.0%" if k in ("gm", "g", "nim") else ("0.000" if k == "wta" else "0.0")
SE.cell(WR["eps"], 8, "2027–2029：營收 ×（毛利率 − opex）× ni_op × fx ÷ shares（PL 頁同式）；2030–2032：淨利 × fx ÷ shares（PL 頁收斂欄同式）")
# v0.5：DCF 改為標準版（FCFF＋WACC）：本路徑在 DCF 頁另寫一個明細區塊（tools/fcff_dcf.py，與三情境同式）
_PW = dict(name="W", label="Wattonomics 錨定（情境 B × wta；Sens 案例，未進基準）", rev={2000 + y: f"Sens!{WC[y]}{WR['rev']}*fx" for y in YRS6}, rev0=f"(PL!B{PR['rev']}*fx)",
           ebit={2000 + y: f"Sens!{WC[y]}{WR['cf']}*fx" for y in YRS6}, tax={2000 + y: "etr_new" for y in YRS6}, ebit_note="Sens 頁案例區塊 EBIT × fx")
_PW["cx"] = {y: f"capex{y - 2000}*fx" for y in (2027, 2028, 2029)}   # v0.5.1：與三情境同（PL 頁資本支出列）
RW = FD.write_detail(DCFW, RD_["end"] + 1, [_PW], FY_X, 1.0, conv_from=2029, title="Sens 案例路徑：Wattonomics 錨定（情境 B × wta）FCFF 明細（新台幣十億元）", fmt="#,##0.0", stub=True)
_wdcf = lambda R_, n: f"DCF!$C${R_['paths'][n]['v12']}"
WT = WR["cf_b"] + 2
SE.cell(WT - 1, 1, "目標價（單一路徑）").font = BOLD
hdr(SE, WT, ["項目", "Wattonomics 錨定", "情境 B", "基準（情境加權，Val 第 0 節）", "", "", "", "說明"])
WTR = {"dcf": WT + 1, "feps": WT + 2, "comp": WT + 3, "tp": WT + 4}
SE.cell(WTR["dcf"], 1, "DCF 價值推到 12 個月後（標準版 FCFF＋WACC；×（1＋ke）^t_tgt − 期間股利）"); SE[f"B{WTR['dcf']}"] = "=" + _wdcf(RW, "W"); SE[f"C{WTR['dcf']}"] = "=" + _wdcf(RD_, "B"); SE[f"D{WTR['dcf']}"] = f"=Val!B{T['dcftp']}"
SE.cell(WTR["feps"], 1, "12 個月後前瞻 EPS（2／12 × FY2027＋10／12 × FY2028）")
SE[f"B{WTR['feps']}"] = f"=2/12*B{WR['eps']}+10/12*C{WR['eps']}"; SE[f"C{WTR['feps']}"] = f"=2/12*B{WR['eps_b']}+10/12*C{WR['eps_b']}"; SE[f"D{WTR['feps']}"] = f"=Val!B{T['feps']}"
SE.cell(WTR["comp"], 1, "Comparable（相對倍數調整後）")
for c in "BC": SE[f"{c}{WTR['comp']}"] = f"=tp_comp_adj*{c}{WTR['feps']}/Val!B{T['feps']}"
SE[f"D{WTR['comp']}"] = "=tp_comp_adj"
SE.cell(WTR["tp"], 1, "目標價（tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable）")
for c in "BC": SE[f"{c}{WTR['tp']}"] = f"=tp_w_dcf*{c}{WTR['dcf']}+(1-tp_w_dcf)*{c}{WTR['comp']}"
SE[f"D{WTR['tp']}"] = f"=Val!B{T['tpw']}"
for rr in WTR.values():
    for c in "BCD": SE[f"{c}{rr}"].number_format = "0.00" if rr == WTR["feps"] else "#,##0"
SE.cell(WTR["tp"] + 1, 1, "核對：情境 B 欄 DCF − Val 第 0.1 節情境 B 價值推到 12 個月後（應為 0）"); SE[f"C{WTR['tp'] + 1}"] = f"=C{WTR['dcf']}-Val!C{T['dcf12']}"; SE[f"C{WTR['tp'] + 1}"].number_format = "0.000000"
# 小表三列
for i, (nm, e27, e28, tp) in enumerate([("基準（Sens 主表基準＝情境 A；目標價＝情境加權）", "B4", f"PL!D{PR['eps']}", f"D{WTR['tp']}"),
                                        ("情境 B（案例底）", f"PL!F{PR['eps']}", f"PL!G{PR['eps']}", f"C{WTR['tp']}"),
                                        ("Wattonomics 錨定（情境 B × wta）", f"B{WR['eps']}", f"C{WR['eps']}", f"B{WTR['tp']}")]):
    rr = W0 + 4 + i; b, bb = W0 + 4, W0 + 5
    SE.cell(rr, 1, nm); SE[f"B{rr}"] = f"={e27}"; SE[f"E{rr}"] = f"={e28}"; SE[f"H{rr}"] = f"={tp}"
    for v, d1, d2 in (("B", "C", "D"), ("E", "F", "G"), ("H", "I", "J")):
        SE[f"{d1}{rr}"] = f"={v}{rr}/{v}${b}-1"; SE[f"{d2}{rr}"] = f"={v}{rr}/{v}${bb}-1"
        SE[f"{v}{rr}"].number_format = "#,##0" if v == "H" else "0.0"
        for d in (d1, d2): SE[f"{d}{rr}"].number_format = "0.0%"

# ---------------- v0.4.2（2026-11-05）：README 展示層與命題回覆（個股 HTML 版面規格 v1.0；只改 README，其他工作表逐格不變） ----------------
RD["A1"] = (str(RD["A1"].value) + "；v0.4.2（2026-11-05）只改展示層（README 新增展示層與命題回覆兩段、HTML 改用新版面 v1.0），其他工作表數值與 v0.4.1 逐格相同"
            "；v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
            "；v0.5.1（2026-11-05）第二輪修正：共用模組 tools/fcff_dcf.py v2（淨現金由 2026-09-30 滾到評價日、β 敏感度 β_u ± 0.3、Sens 美元口徑折現）、資產負債表更新到 3Q26、DCF 資本支出 FY2027–2029 改讀 PL 頁資本支出列"
            "；v0.5.2（2026-11-06）第三輪修正：stub 期 FCFF 改 FY2026、上市子公司非控制權益改市值、現價與同業報價統一 2026-11-04 收盤"
            "；v0.5.3（2026-11-06）WACC 權重 E 含非控制權益"
            "；v0.5.4（2026-11-06）現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守")
def XF(*parts):
    """公式字串：str＝文字（每段 ≤ 200 字，避免 Excel 255 字上限）；tuple(expr,)＝公式運算式。"""
    out = []
    for p_ in parts:
        if isinstance(p_, tuple): out.append(p_[0]); continue
        t_ = p_.replace('"', '""')
        out += [f'"{t_[i:i + 200]}"' for i in range(0, len(t_), 200)]
    return "=" + "&".join(out)
_SP = '"+0.0%;-0.0%"'
_vr = RD.max_row + 1
RD.cell(_vr, 1, "v0.5.4：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守")
_vr += 1
RD.cell(_vr, 1, "v0.5.3（2026-11-06）：第三輪修正第 12 點。WACC 權重與 β 重新加槓桿的 E 改為母公司市值＋非控制權益 mi_v（691；上市子公司市值、其餘帳面）：權益占資本 → 98.23%、WACC → 10.22%。新舊對照（v0.5.2 → v0.5.3）：DCF 目標價 1,532 → 1,532，加權目標價 2,338 → 2,338，上檔空間 +3.5% → +3.5%，投資建議 中立 → 中立。")
_vr += 1
RD.cell(_vr, 1, "v0.5.2（2026-11-06）：第三輪修正（合併前）。(1) Stub 期 FCFF 改用含 stub 期的會計年度 FY2026 FCFF × stub_frac（1,692 × 0.252＝426；v0.5.1 誤用 FY2027，568）。(2) 非控制權益：創意（3443）、采鈺（6789）改 2026-11-04 市值（660.4，取代其帳面 14.8），mi_v 45.3 → 690.9。"
                "(3) 現價與同業報價統一 2026-11-04 收盤：現價 2,385 → 2,260；同業 P/E 中位數 20.5 → 18.6。(4) README 重點觀察 1、重點 1、命題回覆 3 改公式；情境名稱統一；HTML 情境驅動表依節標題定位。新舊對照（v0.5.1 → v0.5.2，當時數值）：DCF 目標價 1,566 → 1,532，加權目標價 2,509 → 2,338，上檔空間 +5.2% → +3.5%，投資建議 中立 → 中立。")
_vr += 1
RD.cell(_vr, 1, "v0.5.1（2026-11-05）：第二輪修正（工作單 2026-11-05 23:20；Andy 23:16 定案）。(1) 共用模組 tools/fcff_dcf.py v2。(2) 資產負債表更新到 2026-09-30（3Q26 財報）：現金與有價證券 4,016.98、有息負債 1,031.58、長期投資 224.633、非控制權益 45.326（NT$B；租賃 36.882 沿用 2026-06-30）；"
                "淨現金以 stub 期滾到評價日 2026-12-31：淨現金與投資 3,128＋stub 期 FCFF（當時用 FY2027 FCFF × stub_frac 0.252 年；情境 B 568；v0.5.2 改 FY2026）− 已宣告 1Q26 股利 182＝3,514。(3) DCF 資本支出 FY2027–FY2029 改讀 PL 頁資本支出列（77／90／95 十億美元 × fx），FY2030 起收斂到 29.4%。"
                "(4) β 敏感度改 β_u ∓ 0.3 各自重新加槓桿。(5) Sens 新增「美元口徑折現」。新舊對照（v0.5 → v0.5.1，當時數值）：WACC 10.22% → 10.22%，DCF 目標價 1,528 → 1,566，加權目標價 2,490 → 2,509，上檔空間 +4.4% → +5.2%，投資建議 中立 → 中立。")
_vr += 1
RD.cell(_vr, 1, XF("v0.5（2026-11-05）：模型改版——DCF 改為標準版（Andy 2026-11-05 20:42、20:45；docs/plan/個股DCF標準規格.md）：FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現，"
                   "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）；ke＝rf＋β ×（ERP＋CRP），rf ", ('TEXT(rf,"0.00%")',), "（新台幣 10 年期公債）、β ", ('TEXT(beta,"0.00")',),
                   "（自算 TW 半導體產業無槓桿 β ", ('TEXT(beta_u,"0.00")',), "（Blume 調整）依市值 D/E 重新加槓桿）、ERP 5%、台灣 CRP ", ('TEXT(crp,"0.00%")',), "，g＝MIN(3%, rf)＝", ('TEXT(g_tv,"0.00%")',),
                   "。新增 DCF 頁（tools/fcff_dcf.py）；Sens 目標價參數改為 WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法與 DCF 權重；刪除 tp_r、tp_roe。"
                   "新舊對照（v0.4.2 → v0.5；v0.5 當時數值，本版數值見上一列 v0.5.1）：DCF 目標價 2,155 → 1,528，加權目標價 2,803 → 2,490"
                   "，上檔空間 +17.5% → +4.4%，投資建議 中立 → 中立。PL、Comparable 不變。"
                   "DCF 差異拆解（建置時以 DCF 頁明細依序替換計算，元／股）：現金流口徑改 FCFF（仍用 r 10%、g 4%）−223；折現率改 WACC 10.2%、g 改 2.0% −393；終值改價值驅動式（RONIC＝WACC）+17；扣期間股利（每季 7 元 × 4）−28。"
                   "FCFF 每股合計（2027–2032，未折現）比舊股東現金流低約 22%–26%（資本支出占營收 36.8% 高於折舊 21.7%）。"))
def _serow(prefix):
    return next(r for r in range(1, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith(prefix))
_kgi = _serow("CoWoS 2027 年底產能")
_V = lambda k, c="B": f"Val!{c}{T[k]}"
_P = lambda s, y, r: f"PL!{COLOF[(s, y)] if y != 26 else 'B'}{r}"
_F = lambda s, y: f"PL!{COLOF[(s, y)]}{_fr}"
_UTL, _HBM = f"CoWoS!C{CR['util']}", f"CoWoS!C{CR['util'] + 1}"
_rec = f"Val!B{_G + 1}"
r = RD.max_row + 2
RD.cell(r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.4.2 起；上方各列原文保留給下游）").font = Font(bold=True); r += 1
RD.cell(r, 1, "命題：台積電在 AI 算力中扮演的是「供給上限」，而不是分到多少金額比例；以下三個問題依序回答。"); r += 1
RD.cell(r, 1, "命題 1：AI 只占營收一部分。AI 加速器相關營收占台積電營收的比重有多大，其餘非 AI 業務能否穩定成長。"); r += 1
RD.cell(r, 1, "命題 2：2027 年誰先卡住。加速器出貨先受 HBM 供給限制，還是先受台積電 CoWoS 先進封裝產能限制。"); r += 1
RD.cell(r, 1, "命題 3：每股值多少。AI 顆數 × 每顆台積電內容 × 漲價，加上非 AI 成長，推到 EPS 與目標價之後，是否高於現價。"); r += 1
_gap = f"({_V('tpw')}-cons_tp)"; _gd = f"(tp_w_dcf*({_V('dcftp')}-cons_tp))"; _gc = f"((1-tp_w_dcf)*(tp_comp_adj-cons_tp))"
RD.cell(r, 1, XF("重點 1：合理價值", (f'IF({_V("upw")}>=0,"高於","低於")',), "現價，投資建議", (_rec,), "。加權目標價約 ", (f'TEXT({_V("tpw")},"#,##0")',), " 元，比現價 ", ('TEXT(px,"#,##0")',), " 元（2026-11-05）",
                 (f'IF({_V("upw")}>=0,"高 ","低 ")&TEXT(ABS({_V("upw")}),"0.0%")',), "。兩種算法差距大：未來現金流折現（DCF）約 ", (f'TEXT({_V("dcftp")},"#,##0")',), " 元、同業倍數法（Comparable，已乘本公司相對同業的長期倍數比 ", ('TEXT(rm_factor,"0.00")',),
                 "）約 ", ('TEXT(tp_comp_adj,"#,##0")',), " 元，各占一半。市場分析師平均目標價約 ", ('TEXT(cons_tp,"#,##0")',), " 元；加權目標價與共識的差距 ", (f'TEXT({_gap},"+#,##0;-#,##0")',),
                 " 元，拆成兩部分：DCF 部分（權重 × (DCF − 共識)）", (f'TEXT({_gd},"+#,##0;-#,##0")',), " 元、占 ", (f'TEXT({_gd}/{_gap},"0%")',), "；Comparable 部分 ", (f'TEXT({_gc},"+#,##0;-#,##0")',), " 元、占 ", (f'TEXT({_gc}/{_gap},"0%")',),
                 "。", (f'IF({_gd}/{_gap}>=0.7,"差距主要來自 DCF","差距由 DCF 與 Comparable 共同造成")',), "（標準版 DCF 的折現率 WACC ", ('TEXT(wacc,"0.0%")',), "、永續成長率 ", ('TEXT(g_tv,"0.0%")',), "）。Comparable（相對倍數調整後）",
                 ('IF(tp_comp_adj>=cons_tp,"高於","低於")&"共識 "&TEXT(ABS(tp_comp_adj/cons_tp-1),"0.0%")',), "：本模型 12 個月後前瞻 EPS 比共識低 ", (f'TEXT(-{_V("geps")},"0.0%")',),
                 "，相對倍數調整（", ('TEXT(rm_factor,"0.00")',), " 倍）抵銷大部分。")); r += 1
RD.cell(r, 1, f'="重點 2：2027 年 HBM 比 CoWoS 更緊。本模型 2027 年 CoWoS 需求約為可用產能（含委外）的 "&TEXT({_UTL},"0.0%")&"，"&IF({_UTL}>1,"已超過","還沒超過")&"滿載；HBM 需求缺口則約為供給的 "&TEXT({_HBM},"0.0%")&"。但這個結論對 CoWoS 產能假設很敏感——敏感度案例「"&Sens!A{_kgi}&"」下，比值升到 "&TEXT(Sens!D{_kgi},"0.0%")&"，CoWoS 會與 HBM 同時受限（這是測試案例，不是門檻）。"'); r += 1
RD.cell(r, 1, f'="重點 3：情境分歧從 2028 年才開始，資本支出不隨情境下修。2027 年三情境 EPS 都是 "&TEXT({_P("A", 27, PR["eps"])},"0.0")&" 元；2028 年樂觀、基準、保守分別為 "&TEXT({_P("A", 28, PR["eps"])},"0.0")&"、"&TEXT({_P("B", 28, PR["eps"])},"0.0")&"、"&TEXT({_P("C", 28, PR["eps"])},"0.0")&" 元。2028 年資本支出三情境都是 "&TEXT(capex28,"0")&" 十億美元，自由現金流因此降得比淨利快（三情境約 "&TEXT({_F("A", 28)},"0.0")&"、"&TEXT({_F("B", 28)},"0.0")&"、"&TEXT({_F("C", 28)},"0.0")&" 十億美元）。標準版 DCF 的資本支出 FY2027–FY2029 直接用這個資本支出列（情境 B FY2028 占營收 "&TEXT(DCF!{RD_["col"][2028]}{RD_["paths"]["B"]["cxr"]},"0.0%")&"，高於折舊占營收 "&TEXT(fc_da,"0.0%")&"），FY2030 起收斂到 "&TEXT(fc_cx_ss,"0.0%")&"。"'); r += 1
_t0 = TPS0 + 2   # Sens 目標價參數表第一個案例列（WACC ±1pt）；之後依序 β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法、DCF 權重
RD.cell(r, 1, XF("重點 4：v0.5 起 DCF 改為標準版（無槓桿自由現金流折現），DCF 目標價為 ", (f'TEXT({_V("dcftp")},"#,##0")',), " 元（舊法股東現金流、折現率 10% 為 2,155 元；差異拆解見 README 頁 v0.5 版本說明；v0.5.1 的變動見其版本說明）。主因是折現率：WACC ",
                 ('TEXT(wacc,"0.0%")',), "（股權成本 ", ('TEXT(ke,"0.0%")',), "＝rf ", ('TEXT(rf,"0.00%")',), " ＋ β ", ('TEXT(beta,"0.00")',), " ×（ERP 5% ＋ 台灣國家風險溢酬 ", ('TEXT(crp,"0.00%")',),
                 "）；β 是 自算 TW 半導體產業無槓桿 β 依本公司市值 D/E 重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設），高於舊法 10%，永續成長率也由 4% 降為 ", ('TEXT(g_tv,"0.0%")',),
                 "（不超過新台幣 10 年期公債殖利率）。終值現值占情境 B 企業價值 ", (f'TEXT(Val!C{SDCF["tvsh"]},"0%")',), "，隱含終值 EV／EBITDA 約 ", (f'TEXT(Val!C{SDCF["mult"]},"0.0")',),
                 " 倍。DCF 參數案例（Sens 頁）：β 低／高 ", (f'TEXT(Sens!D{_t0 + 1},{_SP})',), "／", (f'TEXT(Sens!F{_t0 + 1},{_SP})',), "；ERP 4%／6% ", (f'TEXT(Sens!D{_t0 + 2},{_SP})',), "／", (f'TEXT(Sens!F{_t0 + 2},{_SP})',),
                 "；WACC −1pt／＋1pt 時 ", (f'TEXT(Sens!D{_t0},{_SP})',), "／", (f'TEXT(Sens!F{_t0},{_SP})',), "（以情境加權目標價為基準；β 低／高案例＝β_u ∓ 0.3 各自重新加槓桿，為測試案例）。")); r += 1
_tv = SDCF["tvsh"]; _us = _t0 + 9
RD.cell(r, 1, XF("重點 5：美元口徑折現（Sens 案例，不是基準）：無風險利率改美國 10 年期 ", ('TEXT(rf_us,"0.00%")',), "（基準新台幣 ", ('TEXT(rf,"0.00%")',), "）＋β ×（ERP＋台灣 CRP），股權成本 ", ('TEXT(rf_us+beta*(erp+crp),"0.0%")',),
                 "、永續成長率＝MIN(3%, 美國 10 年期)＝", ('TEXT(MIN(g_cap,rf_us),"0.0%")',), "，現金流與匯率不變；加權目標價 ", (f'TEXT(Sens!C{_us},"#,##0")',), "（相對基準 ", (f'TEXT(Sens!D{_us},{_SP})',), "）。",
                 (f'IF(OR(Val!B{_tv}>0.8,Val!C{_tv}>0.8,Val!D{_tv}>0.8),"終值現值占企業價值超過 80% 的情境："&IF(Val!B{_tv}>0.8,"A "&TEXT(Val!B{_tv},"0%")&" ","")&IF(Val!C{_tv}>0.8,"B "&TEXT(Val!C{_tv},"0%")&" ","")&IF(Val!D{_tv}>0.8,"C "&TEXT(Val!D{_tv},"0%"),"")&"（DCF 價值大部分來自終值）。","三情境終值現值占企業價值都未超過 80%（A／B／C "&TEXT(Val!B{_tv},"0%")&"／"&TEXT(Val!C{_tv},"0%")&"／"&TEXT(Val!D{_tv},"0%")&"）。")',))); r += 1
r += 1
RD.cell(r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上面三個命題）").font = Font(bold=True); r += 1
RD.cell(r, 1, f'="AI 只占營收一部分嗎？答：是。2026 年 AI 相關營收約 "&TEXT({_P("A", 26, PR["ai"])},"0.0")&" 十億美元，占營收 "&TEXT({_P("A", 26, PR["aish"])},"0.0%")&"；到 2028 年，基準情境占 "&TEXT({_P("B", 28, PR["aish"])},"0.0%")&"（樂觀 "&TEXT({_P("A", 28, PR["aish"])},"0.0%")&"、保守 "&TEXT({_P("C", 28, PR["aish"])},"0.0%")&"），非 AI 仍占大多數。非 AI 營收由 2026 年 "&TEXT({_P("A", 26, PR["nonai"])},"0.0")&" 十億美元到 2028 年基準情境 "&TEXT({_P("B", 28, PR["nonai"])},"0.0")&" 十億美元；非 AI 成長率是假設（In 頁 na 系列），不是公司指引。"'); r += 1
RD.cell(r, 1, f'="2027 年誰先卡住？答："&IF({_UTL}<1,"HBM。","CoWoS 與 HBM 都卡。")&"本模型 2027 年 CoWoS 需求約 "&TEXT(CoWoS!C{CR["dem"]},"#,##0")&" 千片、可用產能（含委外）約 "&TEXT(CoWoS!C{CR["capt"]},"#,##0")&" 千片，需求 ÷ 產能 "&TEXT({_UTL},"0.0%")&"；HBM 需求缺口約為供給的 "&TEXT({_HBM},"0.0%")&"。CoWoS 產能是券商估計（公司不揭露）：敏感度案例「"&Sens!A{_kgi}&"」下比值為 "&TEXT(Sens!D{_kgi},"0.0%")&"，屆時兩者同時受限。"'); r += 1
RD.cell(r, 1, XF("每股值多少？答：加權目標價約 ", (f'TEXT({_V("tpw")},"#,##0")',), " 元（DCF ", (f'TEXT({_V("dcftp")},"#,##0")',), "、Comparable 相對倍數調整後 ", ('TEXT(tp_comp_adj,"#,##0")',), "，權重各半），比現價 ", ('TEXT(px,"#,##0")',), " 元",
                 (f'IF({_V("upw")}>=0,"高 ","低 ")&TEXT(ABS({_V("upw")}),"0.0%")',), "，投資建議", (_rec,), "。DCF 三情境價值（推到 12 個月後）樂觀、基準、保守約 ", (f'TEXT({_V("dcf12")},"#,##0")',), "、", (f'TEXT({_V("dcf12", "C")},"#,##0")',), "、", (f'TEXT({_V("dcf12", "D")},"#,##0")',), " 元，",
                 (f'IF(Val!B{T["dcf12"]}<px,"三情境都低於現價",IF(Val!D{T["dcf12"]}<px,"部分情境低於現價","三情境都高於現價"))',), "（標準版 DCF：WACC ", ('TEXT(wacc,"0.0%")',), "、永續成長率 ", ('TEXT(g_tv,"0.0%")',), "）。比市場平均目標價 ", ('TEXT(cons_tp,"#,##0")',),
                 " 元", (f'IF({_gap}<0,"低 ","高 ")&TEXT(ABS({_gap}),"#,##0")',), " 元：其中 DCF 部分 ", (f'TEXT({_gd},"+#,##0;-#,##0")',), " 元（占 ", (f'TEXT({_gd}/{_gap},"0%")',), "）、Comparable 部分 ", (f'TEXT({_gc},"+#,##0;-#,##0")',), " 元（占 ", (f'TEXT({_gc}/{_gap},"0%")',),
                 "），", (f'IF({_gd}/{_gap}>=0.7,"差距主要來自 DCF 的折現率與永續成長率","差距由 DCF 與 Comparable 共同造成")',), "。")); r += 1
r += 1
RD.cell(r, 1, "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間（HTML 投資建議橫幅用；加權目標價使用的 Comparable 即此值）")
RD.cell(r, 2, "=tp_comp_adj/px-1").number_format = "0.0%"
# v0.5.2（第三輪修正第 22 點）：README「重點觀察 1」整句改公式（刪除 v0.3 前寫死的 134.0／−26.2% 等過時數字；下游 tools/target_summary 讀本列）
RD.cell(_R + 2, 1, XF("重點觀察 1：目標價 NT$", (f'TEXT({_V("tpw")},"#,##0")',), "（DCF ", (f'TEXT({_V("dcftp")},"#,##0")',), "、Comparable（相對倍數調整後）", ('TEXT(tp_comp_adj,"#,##0")',), "；DCF 權重 ", ('TEXT(tp_w_dcf,"0%")',), "），比現價 ", ('TEXT(px,"#,##0")',), "（2026-11-05）",
                      (f'IF({_V("upw")}>=0,"高 "&TEXT({_V("upw")},"0.0%"),"低 "&TEXT(-{_V("upw")},"0.0%"))',), "，級距 ", (f"Val!B{_G}",), "、投資建議 ", (_rec,), "。相對倍數調整係數 ", ('TEXT(rm_factor,"0.00")',), "（Val 第 0.5 節）；Comparable 單獨（市場主流同業法、未調整）為 ", (f'TEXT({_V("upc")},"+0.0%;-0.0%")',),
                      "。與共識目標價 ", ('TEXT(cons_tp,"#,##0")',), " 的差距 ", (f'TEXT({_gap},"+#,##0;-#,##0")',), " 元中，DCF 部分 ", (f'TEXT({_gd},"+#,##0;-#,##0")',), "、Comparable 部分 ", (f'TEXT({_gc},"+#,##0;-#,##0")',),
                      "；Comparable 內本模型前瞻 EPS ", (f'TEXT({_V("feps")},"0.0")',), " 比共識前瞻 EPS 低 ", (f'TEXT(-{_V("geps")},"0.0%")',), "、倍數因素 ", (f'TEXT({_V("gmul")},"+0.0%;-0.0%")',), "（共識目標價隱含 ", ('TEXT(cons_tp/(2/12*cons_feps27+10/12*cons_feps28),"0.0")',), " 倍，同業中位數 ", (f'TEXT({_V("pmed")},"0.0")',), " 倍）。"))
wb.save(sys.argv[1]); print("saved")
