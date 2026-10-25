# NVIDIA（NVDA）收支與估值模型 v0.2（AI 半導體 Project：運算層第一個樣板）
# 用法：python3 nvidia/build_nvidia.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 FY(t)＝t−1 年 2 月～t 年 1 月，約等於日曆年 t−1（FY(t) 直接取 CY(t−1)，不跨年加權；11/12 落在同一日曆年）。金額：十億美元（US$B）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import market as MEM
import cycle_position as CP
import valuation_rule as VR
import accel_share as AS
import tp_v2 as TP

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
NVP = ("bw", "rb", "ru", "hp")                       # NVIDIA 四個世代
PNAME = {k: nm for k, nm, *_ in MEM.UNITS}
PLAT = [k for k, *_ in MEM.UNITS]

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx"); SKH_F = only("skhynix/dist/*.xlsx"); NET_F = only("markets/network/dist/*.xlsx")
BRG_F = only("markets/chip_bridge/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

ws_ = openpyxl.load_workbook(TSMC_F, data_only=True)["In"]
TS = {}   # 台積電模型 In 頁：2029 顆數成長與情境顆數係數
for r in ws_.iter_rows(min_row=4):
    k = r[0].value
    if k and (k.startswith("um2") or k.startswith("g29_") or k.startswith("u_")):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
for k, nm, u_, gb, tag, src in MEM.UNITS:     # 與 markets/memory 的 UNITS 必須一致（同一份）
    for i, y in enumerate((26, 27, 28)):
        assert abs(TS[f"u_{k}_{y}"][0] - u_[i]) < 1e-9, (k, y)
for s_ in SC:                                  # 情境係數與 markets/memory 的 SCEN_UNITS 一致
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9

wsn = openpyxl.load_workbook(NET_F, data_only=True)
NETU = {}   # 網通模型 Units 頁 2029 顆數：須等於 UNITS 28 年 ×(1+g29)（同一份成長率）
for r in wsn["Units"].iter_rows(min_row=4, max_row=11):
    NETU[r[0].value] = r[4].value
for k, nm, u_, *_ in MEM.UNITS:
    assert abs(NETU[nm] - u_[2] * (1 + TS[f"g29_{k}"][0])) < 1e-9, nm
NETM = None; REP = None
for r in wsn["Mkt"].iter_rows(min_row=4):
    if r[0].value == "NVIDIA 網通（本模型歸屬）":
        NETM = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
for r in wsn["In"].iter_rows(min_row=4):
    if r[0].value == "rep_nv26":
        REP = ((r[3].value, r[4].value, r[5].value), f"{rel(NET_F)}｜In!D{r[0].row}:F{r[0].row}")
assert NETM and REP and all(v is not None for v in NETM[0])
for r in wsn["In"].iter_rows(min_row=4):     # 網通模型的 2029 成長率須等於台積電模型的
    if r[0].value and r[0].value.startswith("g29_"):
        assert abs(r[4].value - TS[r[0].value][0]) < 1e-9, r[0].value

wsb = openpyxl.load_workbook(BRG_F, data_only=True)["Rack"]
BR = {}   # 橋接表 Rack 頁：欄 B＝GB300、C＝VR200（2026 HBM 價）
for r in wsb.iter_rows(min_row=4, max_row=32):
    lab = str(r[0].value or "")
    for key, pre in (("hbm_px", "HBM 單價"), ("module", "GPU 模組售價"), ("cpu", "CPU（"), ("lpddr", "非 HBM DRAM"), ("resid", "殘差：匣級"),
                     ("rack", "機架售價"), ("gm_chip", "  NVIDIA 晶片層級毛利率")):
        if lab.startswith(pre) and key not in BR:
            BR[key] = (r[1].value, r[2].value, r[3].value, r[4].value, r[0].row)
assert len(BR) == 7, BR
wsb_in = openpyxl.load_workbook(BRG_F, data_only=True)["In"]
GPUP = [c.value for c in wsb_in[15][3:9]]   # In!D15:I15：GB300 低／基準／高、VR200 低／基準／高
assert abs(GPUP[1] - BR["module"][0]) < 1e-9 and abs(GPUP[4] - BR["module"][1]) < 1e-9, GPUP
BRG_REF = lambda key, col: f"{rel(BRG_F)}｜Rack!{col}{BR[key][4]}"

wsk = openpyxl.load_workbook(SKH_F, data_only=True)
HBMP = {}; CDX = {}
for r in wsk["Mkt"].iter_rows(min_row=4):
    lab = str(r[0].value or "")
    if lab.startswith("HBM 單價（"):
        HBMP[lab[-2]] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
for r in wsk["Path"].iter_rows(min_row=10, max_row=20):
    if r[0].value == "一般 DRAM" and r[1].value in SC and isinstance(r[2].value, (int, float)) and r[2].value > 1.1 and r[3].value > 4:
        CDX[r[1].value] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Path!C{r[0].row}:G{r[0].row}")
assert sorted(HBMP) == list(SC) and sorted(CDX) == list(SC), (HBMP.keys(), CDX.keys())

PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 見頂年度（日曆年）：設備市場模型
assert tuple(CP.classify_lag(2028, PEAK[s_]) for s_ in ("A", "B", "C")) == ("中段", "中段", "見頂"), PEAK   # 擁有者 2026-10-23 核准的結果：FY2028 在 A、B 為中段、C 為見頂

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；每顆價格：千美元（US$K）；顆數：百萬顆。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未單獨使用。公司原文＝NVIDIA 8-K 新聞稿（nvidianews.nvidia.com）；CFO 書面說明與 10-Q 在建置環境無法開啟，標「未找到」"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR = "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-"
N1 = "NVIDIA 8-K 新聞稿 "

sec("A. 加速器顆數（markets/memory 與 HBM、台積電、網通模型同一份；情境係數讀台積電模型成品，不另訂）")
for k, nm, u_, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u_[i], None, tag, "markets/memory/market.py UNITS：" + src if i == 0 else "同上", f"台積電模型 {TS[f'u_{k}_{y}'][1]}（值相同，建置時檢查）" if i == 0 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", None, TS[f"g29_{k}"][0], None, "Assumed", TS[f"g29_{k}"][1], "markets/memory 無 2029；沿用台積電模型（網通模型相同，建置時檢查）")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電模型相同；markets/memory SCEN_UNITS 相同，建置時檢查）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
for k in NVP:
    gb = [g for kk, nm, u_, g, t, sr in MEM.UNITS if kk == k][0]
    for i, y in enumerate((26, 27, 28)):
        inp(f"gb_{k}_{y}", f"{PNAME[k]} 每顆 HBM 20{y}", "GB", None, gb[i], None, "Interested-party", "markets/memory/market.py UNITS（GB300：B300 288；Rubin 288；Rubin Ultra：NVIDIA 1 TB／封裝）", "2029 沿用 2028" if i == 0 else "同上")

sec("B. 每顆售價（運算部分）：GPU 模組＋CPU＋LPDDR＋匣級加價中 NVIDIA 保留的部分（US$K／顆；橋接表 v0.2.1 Rack 頁）")
inp("pm_bw", "GPU 模組售價：Blackwell（GB300，NVIDIA 對買方）", "US$K", GPUP[0], BR["module"][0], GPUP[2], "Derived", BRG_REF("module", "B"), f"橋接表 In!D15:F15（低／基準／高）；Tokenomics v20.1 機架基準價 $4.3M。B200 較 B300 便宜，本模型不分（在 Recon 列為顆數＋價格差額的候選來源）")
inp("pm_rb", "GPU 模組售價：Rubin（VR200，2026 HBM 價）", "US$K", GPUP[3], BR["module"][1], GPUP[5], "Derived", BRG_REF("module", "C"), "橋接表 In!G15:I15；Bernstein 估 VR200 機架 $9.1M、Fubon 估 Rubin GPU $78–80K（二手，Interested-party）高於此基準，未採")
inp("pm_ru", "GPU 模組售價：Rubin Ultra", "US$K", 60, 75, 100, "Assumed", "無：未發布。基準＝VR200 55 ×（VR200／GB300 的 1.36 倍外推）＝75", "BofA 估 Rubin Ultra 機架約 $21M（二手；每封裝約 $146K 是本模型算術，封裝數有爭議；wccftech）。擁有者待定區間")
inp("pm_hp", "GPU 模組售價：Hopper（H200 等）", "US$K", 25, 30, 35, "Assumed", "無：2026 出貨僅 0.5M，影響小", "第二來源：未找到")
for k, col, nm in (("bw", "B", "GB300"), ("rb", "C", "VR200")):
    inp(f"pcpu_{k}", f"CPU 每顆 GPU 分攤（{nm}）", "US$K", None, round(BR["cpu"][0 if col == "B" else 1] * 1000 / 72, 6), None, "Derived", BRG_REF("cpu", col) + "（$M／架 × 1000 ÷ 72）", "NVIDIA 自家 Grace／Vera")
    inp(f"plp_{k}", f"LPDDR5X 每顆 GPU 分攤（{nm}）", "US$K", None, round(BR["lpddr"][0 if col == "B" else 1] * 1000 / 72, 6), None, "Derived", BRG_REF("lpddr", col) + "（÷ 72）", "隨一般 DRAM 價格路徑變動（見 D 節）")
    inp(f"pres_{k}", f"匣級加價（殘差）每顆 GPU（{nm}）", "US$K", None, round(BR["resid"][0 if col == "B" else 1] * 1000 / 72, 6), None, "Derived", BRG_REF("resid", col) + "（÷ 72）", "匣級電路板、供電與 NVIDIA／ODM 匣級加價；NVIDIA 保留比例見 k_tray")
inp("pcpu_ru", "CPU 每顆 GPU 分攤（Rubin Ultra）", "US$K", None, "=pcpu_rb", None, "Analogy", "比照 VR200", "—")
inp("plp_ru", "LPDDR5X 每顆 GPU 分攤（Rubin Ultra）", "US$K", None, "=plp_rb", None, "Analogy", "比照 VR200", "—")
inp("pres_ru", "匣級加價每顆 GPU（Rubin Ultra）", "US$K", None, "=pres_rb", None, "Analogy", "比照 VR200", "—")
inp("pcpu_hp", "CPU 每顆 GPU 分攤（Hopper）", "US$K", None, 0, None, "Assumed", "Hopper 以 HGX 板出貨，價格已含在模組價", "—")
inp("plp_hp", "LPDDR5X 每顆 GPU 分攤（Hopper）", "US$K", None, 0, None, "Assumed", "同上", "—")
inp("pres_hp", "匣級加價每顆 GPU（Hopper）", "US$K", None, 0, None, "Assumed", "同上", "—")
inp("k_tray", "匣級加價中 NVIDIA 保留的比例", "x", 0.0, 0.5, 1.0, "Assumed", "無：橋接表殘差由 NVIDIA 與 ODM 共享，公司未揭露", "第二來源：未找到。影響運算營收水準（Blackwell 每顆 ±1.5K）；FY2027 的水準差額見 Recon")
inp("hpx_bw", "HBM 單價 2026 合約：Blackwell（GB300）", "US$/GB", None, BR["hbm_px"][0], None, "Derived", BRG_REF("hbm_px", "B"), "橋接表 In!hbm_p（Tokenomics）；Hopper 比照")
inp("hpx_rb", "HBM 單價 2026 合約：Rubin（VR200）", "US$/GB", None, BR["hbm_px"][1], None, "Derived", BRG_REF("hbm_px", "C"), "Rubin Ultra 比照；HBM4 估計 $15–53/GB（Silicon Analysts／Fubon／Bernstein，二手，見 Recon）")
inp("hpx_ru", "HBM 單價 2026 合約：Rubin Ultra", "US$/GB", None, "=hpx_rb", None, "Analogy", "比照 VR200", "—")
inp("hpx_hp", "HBM 單價 2026 合約：Hopper", "US$/GB", None, "=hpx_bw", None, "Analogy", "比照 GB300", "—")
inp("abs_mem", "HBM／LPDDR 漲價由 NVIDIA 吸收的比例（其餘轉嫁到售價）", "x", 0.0, 0.5, 1.0, "Assumed", "無：橋接表 Rack 頁列兩端（D 欄吸收＝1、E 欄全額轉嫁＝0）；公司稱 FY2028 Q1 起調價（二手：Yahoo Finance、247wallst）", "第二來源：未找到。擁有者待定基準值；基準取兩端中點，不為對上公司毛利率指引而調（指引只在 Recon 對帳）")

sec("C. 記憶體價格路徑（讀 skhynix 成品；唯讀；情境別）")
for s in SC:
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"hbm_{s}_{y}", f"HBM 混合單價 20{y}（{s}）", "US$/GB", None, HBMP[s][0][j + 1], None, "Derived", HBMP[s][1], "markets/memory：2026 BofA $14.6/GB；2027 +80% 三情境相同；2028–29 依情境" if (s == "A" and y == 26) else "同上")
for s in SC:
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"cdx_{s}_{y}", f"一般 DRAM 合約價年均指數 20{y}（{s}；2025Q1＝1）", "x", None, CDX[s][0][j + 1], None, "Derived", CDX[s][1], "LPDDR 成本隨一般 DRAM 合約價（Analogy）" if (s == "A" and y == 26) else "同上")

sec("D. 網通（讀 markets/network 成品；情境 A 基準，B／C 依 NVIDIA 顆數係數縮放——網通金額與顆數成正比、組合不變，故縮放精確）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"net_{y}", f"NVIDIA 網通（本模型歸屬）20{y}", "US$B", None, NETM[0][j], None, "Derived", NETM[1], "廠商歸屬比例為 Assumed（markets/network In 頁 E 節）；2026 較實際推估低 16.6%（見 Recon）" if y == 26 else "同上")
inp("rep_nv26", "NVIDIA 網通營收 2026（實際＋推估）", "US$B", REP[0][0], REP[0][1], REP[0][2], "Verified／Derived", REP[1], "Q1 FY27 14.8（公司，8-K 新聞稿）、Q2 約 17.5（法說「季增 18%」，二手）、Q3–Q4 Assumed")
inp("net_q1", "Data Center 網通營收 Q1 FY27", "US$B", None, 14.8, None, "Verified", PR + "first-quarter-fiscal-2027（公司最後一次揭露 compute／networking 拆分）", "Q2 FY27 起不再揭露：未找到（CFO 書面說明無法開啟）")
inp("c_net", "網通非 HBM 成本占營收", "x", 0.2, 0.3, 0.4, "Assumed", "無：公司不揭露網通毛利率", "第二來源：未找到。決定網通毛利率 60–80%")

sec("E. Edge Computing（PC、遊戲機、工作站、機器人、車用；自 Q1 FY27 起公司只揭露 Data Center 與 Edge Computing 兩個平台）")
inp("gm_edge", "Edge Computing 毛利率", "x", 0.5, 0.6, 0.7, "Assumed", "無：公司不揭露平台毛利率（舊 Gaming／ProViz／Auto 的毛利率亦未揭露）", "第二來源：未找到")
inp("eg", "Edge Computing 營收年增（FY2028 起）", "%", 0.0, 0.08, 0.15, "Assumed", "公司：Q2 FY27 年增 27%、季增 13%；記憶體成本壓抑消費端需求（二手）", "第二來源：未找到")
inp("edge_q", "Edge Computing 季增（Q3、Q4 FY27 假設）", "%", 0.0, 0.05, 0.10, "Assumed", "Q1 FY27 季增 10%、Q2 季增 13%（公司）", "—")

sec("F. 中國資料中心運算（基準為 0＝公司 Q3 FY27 指引假設；只在 Sens 列「恢復」）")
inp("cn_amt", "中國恢復後的資料中心運算營收（FY2028 起，每年）", "US$B", 10, 20, 40, "Assumed", "公司 Q3 FY27 指引不假設中國資料中心運算營收；Q2 FY27 的 H200 對中國出貨 < 資料中心營收的 1%（二手：Yahoo／CNBC 摘要）", "第二來源：未找到。FY2026 Q1 H20 費用 $4.5B（公司）作為量級參照")

sec("G. 營業費用、稅、股數（FY2028 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.4, 0.6, 0.8, "Derived", "FY2027 錨點營業費用年增 ÷ 營收年增（Q 頁：約 0.6）", "低／高 Assumed")
inp("g_opx_min", "營業費用年增下限", "%", 0.0, 0.05, 0.10, "Assumed", "無", "營收下滑時費用僵固")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準為 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("tx_f", "有效稅率（FY2028 起）", "%", 0.16, 0.17, 0.18, "Interested-party", "公司：FY2027 全年 GAAP／non-GAAP 稅率 16.0%–18.0%（Q2 FY27 新聞稿）", "FY2026 15.1%、Q1 FY27 16.6%、Q2 16.5%（公司，Q 頁）")
inp("oth_f", "其他收益淨額（不含權益證券損益；FY2028 起每年）", "US$B", 0.5, 1.2, 2.4, "Assumed", "FY2026 2.1、Q1 FY27 0.43、Q2 FY27 0.00（GAAP 扣權益證券損益；Q 頁）", "權益證券損益屬一次性，一律排除")
inp("sh_g", "稀釋股數年變動（FY2028 起）", "%", -0.025, -0.012, 0.0, "Derived", "FY2025→FY2026 −1.2%（公司，Q 頁）；Q2 FY27 回購 $19.7B＋股息 $6.0B、剩餘授權約 $99B（公司）", "低／高 Assumed")

sec("H. FY2027 錨點假設（Q1、Q2 實績＋Q3 指引＋Q4 假設）")
inp("q3_rev", "Q3 FY27 營收指引（±2%）", "US$B", 105.84, 108.0, 110.16, "Interested-party", PR + "second-quarter-fiscal-2027：$108.0B ±2%；不假設中國資料中心運算營收", "Webull、Yahoo Finance 摘要（二手）")
inp("q3_gm", "Q3 FY27 毛利率指引（GAAP＝non-GAAP）", "%", 0.735, 0.74, 0.745, "Interested-party", PR + "second-quarter-fiscal-2027：74.0% ±50bp", "implicator.ai（二手）")
inp("q3_opx", "Q3 FY27 GAAP 營業費用指引", "US$B", None, 9.2, None, "Interested-party", PR + "second-quarter-fiscal-2027：約 $9.2B（non-GAAP $9.0B）", "—")
inp("q4_g", "Q4 FY27 營收季增（Q4 假設）", "%", 0.05, 0.10, 0.15, "Assumed", "無：Q1→Q2 +17.9%、Q2→Q3 指引 +12.2%（公司）；共識 FY2027 營收約 $411.7B（stockanalysis 摘要，二手）", "第二來源：未找到")
inp("q4_gm", "Q4 FY27 毛利率（Q4 假設）", "%", 0.71, 0.715, 0.74, "Assumed", "二手報導：Q4 FY27 谷底 71–72%（implicator.ai、siliconanalysts；逐字稿未取得）", "高＝與 Q3 指引持平")
inp("q4_opx_g", "Q4 FY27 營業費用季增（Q4 假設）", "%", 0.05, 0.08, 0.12, "Assumed", "Q1→Q2 +10.3%、Q2→Q3 指引 +9.4%（公司）", "—")
inp("q_oth", "Q3、Q4 FY27 其他收益淨額（不含權益證券損益；每季）", "US$B", 0.0, 0.35, 0.45, "Assumed", "Q1 FY27 0.43、Q2 0.00（GAAP 扣權益證券損益）；non-GAAP 0.46、0.30（公司）", "公司 Q3 指引未列：未找到")
inp("tx_27", "Q3、Q4 FY27 有效稅率", "%", 0.16, 0.17, 0.18, "Interested-party", "公司：FY2027 全年 16.0%–18.0%", "—")
inp("q_sh", "Q3、Q4 FY27 稀釋股數季變動", "%", -0.006, -0.003, 0.0, "Derived", "FY2026 季均 −0.2%（公司，Q 頁）", "—")

sec("I. 公司指引與共識（只作 Recon 對帳，不作校準）")
inp("g28_gr", "公司 FY2028 營收年增", "%", None, 0.70, None, "Interested-party", "Q2 FY27 法說（二手：Yahoo Finance、CNBC 標題確認「約 70%」）；假設供給受限；逐字稿未取得", "工作單記約 US$690B：未找到公司直接陳述的美元金額，Recon 以 FY2027 錨點 ×1.7 推算")
inp("g28_gm", "公司 FY2028 毛利率指引", "%", 0.72, 0.725, 0.73, "Interested-party", "二手報導：FY2028 全年 72–73%（Q1 FY28 起調價生效）；Q4 FY27 谷底 71–72%（moomoo、biggo；逐字稿未取得）", "公司 Q3 FY27 指引 74.0%（8-K）")
inp("cons_g28", "共識 FY2028 營收年增（指引後）", "%", None, 0.63, None, "Interested-party", "Yahoo Finance 摘要（二手）；指引前 LSEG 約 44%", "第二來源：未找到")
inp("cons_rev27", "共識 FY2027 營收", "US$B", None, 411.7, None, "Interested-party", "stockanalysis.com（搜尋摘要，頁面無法開啟）；另一來源 396（可能較舊）", "—")
inp("cons_eps27", "共識 FY2027 EPS", "US$", None, 9.31, None, "Interested-party", "stockanalysis.com（搜尋摘要）；9.28–9.31 的其他摘要", "新口徑 non-GAAP（含 SBC）")
inp("cons_eps28", "共識 FY2028 EPS", "US$", 15.70, 15.91, 16.13, "Interested-party", "247wallst 2026-10-19：15.70–15.91；Wallstreetzen 平均 16.13（68 位分析師，區間 13.13–19.65）", "新口徑 non-GAAP（含 SBC）。FY2029 共識：未找到")
inp("th_mult", "命題中現價對「2027 年 EPS」的倍數（日曆年 2027＝FY2028）", "x", None, 13.5, None, "Interested-party", "擁有者命題（工作單）", "對照：現價 ÷ 共識 FY2028 EPS 約 13.4–13.5x（Val 頁）")
inp("th_sh_lo", "命題：2027 年 NVIDIA 金額份額（下緣）", "%", None, 0.75, None, "Interested-party", "擁有者命題；可歸屬的外部來源：未找到（Kaiso：2028 前 60–75%；Silicon Analysts：2025 約 80%）", "—")
inp("th_sh_hi", "命題：2027 年 NVIDIA 金額份額（上緣）", "%", None, 0.80, None, "Interested-party", "同上", "—")

sec("J. 各家每顆加速器金額（份額計算用；模組層級，與 NVIDIA 的 GPU 模組價同一口徑）")
inp("pa_amd", "AMD Instinct 每顆", "US$K", 20, 30, 45, "Interested-party", "MI355X：HSBC 約 $23K（自 $25K 下修）、其他報導 $20–25K；MI455X 無官方價，部落格 $15–50K（finviz、gpusmith、MikeLongTerm；賣方有目標價誘因）", "第二來源：以上三者互相一致於 $20–25K（MI355X）；MI455X 無獨立來源")
inp("pa_tpu", "Google TPU 每顆", "US$K", 8, 12, 18, "Assumed", "無：每顆 ASP 未找到（只有租用價代理：Ironwood $1.03–1.60／小時，SemiAnalysis，不可單獨使用）", "未找到")
inp("pa_trn", "AWS Trainium 每顆", "US$K", 5, 8, 12, "Assumed", "無：每顆 ASP 未找到", "Maia 200 約為 NVIDIA（逾 $30K）的 30–40%＝$9–12K（二手，cryptobriefing）作量級參照")
inp("pa_oth", "其他 ASIC 每顆", "US$K", 5, 8, 12, "Assumed", "無：未找到（Maia、MTIA 皆無 ASP）", "同上")
inp("dsh", "ASIC 金額份額變動（Sens 用；NVIDIA 份額 ±）", "pt", None, 0.05, None, "Assumed", "無", "Sens 用 ±5 個百分點")
inp("uk", "NVIDIA 顆數係數（Sens 用；FY2028 起乘以所有 NVIDIA 世代顆數）", "x", 0.9, 1.0, 1.1, "Assumed", "無", "±10%")

sec("K. 價格與估值")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 210.96, None, "Interested-party", "markets/invest_map/data/prices.csv（tipranks；bullfincher 兩來源核對）；stockanalysis.com 統計頁同值（2026-10-23 收盤 210.96）。v0.1.2 為 10-22 收盤 212.21", "股數量級核對（v0.1.2）：市值約 $5.12T ÷ 212.21 ≈ 24.1B 股，與 Q2 FY27 稀釋股數 24.285B 同量級")
inp("r_disc", "折現率（把 FY2028 年底隱含價折回現在；參考用）", "%", 0.08, 0.10, 0.12, "Assumed", "無", "FY2028 年底（2028-01）距 2026-10-22 約 1.27 年")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "共用規則（tools/cycle_position.py）：NVIDIA 的 FY(t) 對照日曆年 t−1，用 classify_lag；設備市場模型的情境係數為 Assumed" if s == "A" else "同上")
sec("L. 目標價情境權重（v0.1.2 新增；Val 頁第 0 節用）")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
sec("M. 目標價參數（v0.2 新增：DCF 與 Comparable 加權、FY2030 後收斂法；黃底＝Andy 可改）")
inp("tp_r", "DCF 折現率（股東權益成本）", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單 2026-10-25：美、台、日、韓同一值）", "Sens 頁 ±1pt")
inp("tp_g", "長期成長率（FY2032 後永續；PL 延伸期 B 情境的收斂目標）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單 2026-10-25）", "Sens 頁 ±1pt（只改 DCF 再投資率與永續成長，不改 PL 收斂路徑）")
inp("tp_gd", "延伸期 A／C 情境長期成長相對 tp_g 的差（A＋、C−）", "pt", None, 0.02, None, "Assumed", "工作單 2026-10-25：A＝tp_g＋2pt、C＝tp_g−2pt", "—")
inp("tp_roe", "近三年平均 ROE（再投資率＝tp_g ÷ tp_roe）", "%", None, 0.7915, None, "Derived", "valueinvesting.io NVDA ROE（GAAP，年度）：FY2024 69.24%、FY2025 91.87%、FY2026 76.33%，平均 79.15%（資料商，Interested-party）", "stockanalysis 近 12 月 ROE 117.21%（2026-10-25 頁面）；再投資率上限 0.8")
inp("tp_w_dcf", "目標價中 DCF 的權重（其餘為 Comparable）", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）", "Sens 頁 0.3／0.7")
inp("nm10", "十年平均淨利率（FY2016–FY2025，GAAP）：延伸期 C 情境的收斂目標", "%", None, 0.3118, None, "Derived",
    "finboard.net NVDA 損益表（營收、淨利，US$M）：FY2016 5,010／614、FY2017 6,910／1,666、FY2018 9,714／3,047、FY2019 11,716／4,141、FY2020 10,918／2,796、FY2021 16,675／4,332、FY2022 26,914／9,752、FY2023 26,974／4,368、FY2024 60,922／29,760、FY2025 130,497／72,880；十年淨利率平均 31.18%",
    "valueinvesting.io 淨利率 FY2017–FY2025 逐年相同（24.11%…55.85%）；本模型 EPS 口徑（GAAP 扣權益證券損益）與 GAAP 淨利相近")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, 328.95, None, "Interested-party", "stockanalysis.com（62 位分析師平均，2026-10-25 頁面；現價為 10-23 收盤）", "賣方有目標價誘因；只作對照，不進目標價")

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


# ---------------- Q：實績與 FY2027 錨點（FY2026 四季、FY2027 Q1、Q2 實績；Q3 指引；Q4 假設） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億美元；EPS 為美元；股數為十億股）。會計年度 FY2026＝2025-02～2026-01。藍字＝輸入；黑字＝公式。實績取自 NVIDIA 8-K 新聞稿；第二來源＝後續季度新聞稿的比較欄（同一公司不同文件）"
Q["A2"] = "FY2027 錨點＝Q1、Q2 實績＋Q3 指引中點＋Q4 假設（I、J 欄，假設見 In 頁 H 節）。EPS 口徑：GAAP 扣除權益證券損益（一次性；權益證券損益 Q1 FY27 為 $15.9B、Q2 為 $7.8B），含股票薪酬；與公司自 Q1 FY27 起的新 non-GAAP 口徑（含 SBC）相近，差額見下方檢查列"
hdr(Q, 3, ["項目", "FY26 Q1（2025-02～04）", "FY26 Q2（～07）", "FY26 Q3（～10）", "FY26 Q4（～2026-01）", "FY2026 全年", "FY27 Q1（2026-02～04）", "FY27 Q2（～07）", "FY27 Q3（指引中點）", "FY27 Q4（假設）", "FY2027 錨點", "標記", "來源1（公司）", "來源2／說明"])
QS = {"B": "first-quarter-fiscal-2026", "C": "second-quarter-fiscal-2026", "D": "third-quarter-fiscal-2026", "E": "fourth-quarter-and-fiscal-2026", "G": "first-quarter-fiscal-2027", "H": "second-quarter-fiscal-2027"}
S1 = "NVIDIA 8-K 新聞稿（nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-）：" + "；".join(f"{k}欄＝{v}" for k, v in QS.items())
S2 = "後續季度新聞稿比較欄交叉：Q1、Q4 FY26 見 Q1 FY27 稿；Q2 FY26 見 Q2 FY27 稿；Q3 FY26 見 Q4 FY26 稿"
ACT = "BCDEGH"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None, ann27=None):
    """vals：{欄: 值或公式}；ann／ann27：F／K 欄公式（None＝SUM）。"""
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(B{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    if ann27 != "":
        Q[f"K{r}"] = ann27 or f"=SUM(G{r}:J{r})"; Q[f"K{r}"].number_format = fmt
    Q.cell(r, 12, tag); Q.cell(r, 13, s1); Q.cell(r, 14, s2)
def act(vs): return dict(zip(ACT, vs))
qrow(4, "營收", {**act([44.062, 46.743, 57.006, 68.127, 81.615, 96.221]), "I": "=q3_rev", "J": "=I4*(1+q4_g)"}, tag="Verified／Interested-party（Q3 指引）／Assumed（Q4）")
qrow(5, "Data Center 營收（公司揭露，四捨五入到 0.1）", {**act([39.1, 41.1, 51.2, 62.3, 75.2, 89.0]), "I": "=I4-I6", "J": "=J4-J6"}, fmt="#,##0.0", tag="Verified／Derived（I、J）", s2=S2 + "；FY2026 合計 193.7＝公司全年值")
qrow(6, "Edge Computing 營收（＝營收 − Data Center）", {**{c: f"={c}4-{c}5" for c in "BCDEGH"}, "I": "=H6*(1+edge_q)", "J": "=I6*(1+edge_q)"}, fmt="#,##0.000", tag="Derived（實績）／Assumed（I、J）", s1="公司 Q1 FY27 起只揭露 Data Center 與 Edge Computing 兩個平台；FY26 各季 Edge＝營收 − Data Center（舊 Gaming＋ProViz＋Auto＋OEM 合計）", s2="Q1、Q2 FY27 公司官方 Edge 為 6.4、7.2（下一列核對）")
qrow(7, "  公司官方揭露的 Edge Computing（Q1、Q2 FY27）", {"G": 6.4, "H": 7.2}, fmt="#,##0.0", ann="", ann27="", tag="Verified", s1=PR + "first-quarter-fiscal-2027；" + PR + "second-quarter-fiscal-2027", s2="與上列推得值差 < 0.05（四捨五入）")
qrow(8, "毛利（GAAP）", {**act([26.668, 33.853, 41.849, 51.093, 61.157, 72.142]), "I": "=I4*q3_gm", "J": "=J4*q4_gm"}, tag="Verified／Interested-party（I）／Assumed（J）")
qrow(9, "毛利率", {c: f"={c}8/{c}4" for c in "BCDEFGHIJK"}, fmt="0.0%", ann="", ann27="", tag="Derived", s1="—", s2="公司：Q1 FY26 60.5%（含 H20 費用 $4.5B）、FY2026 71.1%")
qrow(10, "營業費用（GAAP，R&D＋SG&A）", {**act([5.030, 5.413, 5.839, 6.794, 7.621, 8.408]), "I": "=q3_opx", "J": "=I10*(1+q4_opx_g)"}, tag="Verified／Interested-party（I）／Assumed（J）")
qrow(11, "營業利益", {**act([21.638, 28.440, 36.010, 44.299, 53.536, 63.734]), "I": "=I8-I10", "J": "=J8-J10"}, tag="Verified／Derived（I、J）")
qrow(12, "檢查：毛利 − 營業費用 − 營業利益（實績應為 0）", {c: f"={c}8-{c}10-{c}11" for c in "BCDEGH"}, fmt="0.000", tag="—", s1="—", s2="—")
qrow(13, "其他收益淨額（GAAP 合計，含權益證券損益）", act([0.272, 2.766, 1.926, 6.098, 16.367, 7.773]), ann27="", tag="Verified")
qrow(14, "  其中：權益證券淨損益（一次性，排除）", act([-0.175, 2.247, 1.354, 5.491, 15.936, 7.771]), ann27="", tag="Verified", s1=S1 + "（non-GAAP 調節表「Gains from equity securities, net」）")
qrow(15, "其他收益淨額（不含權益證券損益）", {**{c: f"={c}13-{c}14" for c in "BCDEGH"}, "I": "=q_oth", "J": "=q_oth"}, tag="Derived／Assumed（I、J）", s1="GAAP 合計 − 權益證券損益；利息收入、利息費用與其他", s2="公司 non-GAAP 其他收益：Q1 FY26 0.45、Q2 0.52、Q4 0.62、Q1 FY27 0.46、Q2 FY27 0.30（另含股息收入等調節）")
qrow(16, "稅前利益（不含權益證券損益）", {c: f"={c}11+{c}15" for c in "BCDEGHIJ"}, tag="Derived")
qrow(17, "所得稅（GAAP）", act([3.135, 4.784, 6.026, 7.437, 11.582, 11.819]), ann27="", tag="Verified")
qrow(18, "稅前利益（GAAP）", {c: f"={c}11+{c}13" for c in "BCDEGH"}, ann27="", tag="Derived", s1="營業利益＋其他收益淨額", s2="公司 Q1 FY27 69.903、Q2 71.507")
qrow(19, "有效稅率", {**{c: f"={c}17/{c}18" for c in "BCDEGH"}, "I": "=tx_27", "J": "=tx_27", "F": "=F17/F18", "K": "=1-K20/K16"}, fmt="0.0%", ann="", ann27="", tag="Derived／Interested-party（FY2027 全年 16.0–18.0%）", s1="GAAP 稅 ÷ GAAP 稅前；K 欄為扣除權益證券損益後的全年有效稅率（實績季用 GAAP 有效稅率套用在不含損益的稅前利益）", s2="公司：FY2027 全年 GAAP／non-GAAP 稅率 16.0%–18.0%")
qrow(20, "淨利（不含權益證券損益）＝稅前 ×（1 − 有效稅率）", {c: f"={c}16*(1-{c}19)" for c in "BCDEGHIJ"}, tag="Derived")
qrow(21, "稀釋股數（十億股）", {**act([24.611, 24.532, 24.483, 24.432, 24.391, 24.285]), "I": "=H21*(1+q_sh)", "J": "=I21*(1+q_sh)"}, fmt="#,##0.000", ann="=AVERAGE(B21:E21)", ann27="=AVERAGE(G21:J21)", tag="Verified／Derived（I、J）")
qrow(22, "EPS（不含權益證券損益，美元）", {**{c: f"={c}20/{c}21" for c in "BCDEGHIJ"}, "F": "=F20/F21", "K": "=K20/K21"}, fmt="0.000", ann="", ann27="", tag="Derived", s1="淨利（不含權益證券損益）÷ 稀釋股數", s2="FY2027 Q3、Q4 為指引與假設推得")
qrow(23, "淨利（GAAP，公司）", act([18.775, 26.422, 31.910, 42.960, 58.321, 59.688]), ann27="", tag="Verified")
qrow(24, "EPS（GAAP 稀釋，公司）", act([0.76, 1.08, 1.30, 1.76, 2.39, 2.46]), fmt="0.00", ann=4.90, ann27="", tag="Verified", s2=S2 + "；FY2026 全年 4.90（公司）")
Q["F24"].font = BLUE
qrow(25, "EPS（non-GAAP 稀釋，新口徑含 SBC；Q3 FY26 官方重編值：未找到）", {"B": 0.78, "C": 1.01, "E": 1.59, "G": 1.87, "H": 2.22}, fmt="0.00", ann="", ann27="", tag="Verified", s1="Q1 FY27 稿對帳表（Q1、Q4 FY26、Q1 FY27）；Q2 FY27 稿對帳表（Q2 FY26、Q2 FY27）", s2="舊口徑（排除 SBC）原值：0.81、1.05、1.30、1.62；FY2026 4.77")
qrow(26, "股票薪酬 SBC（現金流量表）", act([1.474, 1.624, 1.655, 1.633, 1.928, 2.027]), ann27="", tag="Verified")
qrow(27, "檢查：GAAP 淨利 − (GAAP 稅前 − 稅)（應為 0）", {c: f"={c}23-({c}18-{c}17)" for c in "BCDEGH"}, fmt="0.000", ann="", ann27="", tag="—", s1="—", s2="—")
qrow(28, "檢查：EPS（不含權益證券損益）− non-GAAP EPS（新口徑）", {c: f"={c}22-{c}25" for c in "BCEGH"}, fmt="0.000", ann="", ann27="", tag="—", s1="差額來自收購相關成本、其他調節項與稅", s2="Q1、Q2 FY27 差 −0.03～−0.02：本模型 EPS 口徑可近似公司新 non-GAAP")
qrow(29, "檢查：FY2026 四季合計 − 公司全年值（營收 215.938、營業利益 130.387，應為 0）", {"F": "=(F4-215.938)+(F11-130.387)"}, fmt="0.000", ann="", ann27="", tag="—", s1="—", s2="—")
for nm, rr in (("rev", 4), ("dc", 5), ("edge", 6), ("gp", 8), ("opx", 10), ("oi", 11), ("oth", 15), ("ni", 20), ("sh", 21), ("eps", 22)):
    defname(f"a26_{nm}", f"Q!$F${rr}"); defname(f"a27_{nm}", f"Q!$K${rr}")
defname("a27_etr", "Q!$K$19")
Q.cell(31, 1, "FY2027 錨點與共識、指引的距離（參考；校準時未使用共識）").font = BOLD
Q.cell(32, 1, "FY2027 營收 ÷ 共識營收 − 1"); Q["B32"] = "=K4/cons_rev27-1"; Q["B32"].number_format = "0.0%"
Q.cell(33, 1, "FY2027 EPS（不含權益證券損益）÷ 共識 EPS − 1"); Q["B33"] = "=K22/cons_eps27-1"; Q["B33"].number_format = "0.0%"
Q.cell(34, 1, "FY2027 營收年增"); Q["B34"] = "=K4/F4-1"; Q["B34"].number_format = "0.0%"
Q.cell(35, 1, "FY2027 營業費用年增"); Q["B35"] = "=K10/F10-1"; Q["B35"].number_format = "0.0%"
Q.cell(36, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B36"] = "=B35/B34"; Q["B36"].number_format = "0.00"
Q.column_dimensions["A"].width = 62
for c in "BCDEFGHIJK": Q.column_dimensions[c].width = 14
Q.column_dimensions["L"].width = 24; Q.column_dimensions["M"].width = 90; Q.column_dimensions["N"].width = 70
Q.freeze_panes = "B4"

# ---------------- Units：加速器顆數（三情境）、金額份額（可重用函式 accel_share）、售價、校準 ----------------
CYS = (2026, 2027, 2028, 2029)
cyc = lambda y: chr(ord("C") + y - 2026)          # Units 的日曆年欄：C＝2026 … F＝2029
def uexp(k, y, s):
    if y <= 2027: return f"u_{k}_{y % 100}"
    if y == 2028: return f"u_{k}_28*um28_{s}"
    return f"u_{k}_28*(1+g29_{k})*um29_{s}"
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 加速器顆數（百萬顆；日曆年；與 HBM、台積電、網通模型同一份）與金額份額（模組層級每顆價 × 顆數）。NVIDIA 會計年度 FY(t) 對照日曆年 t−1（FY2027＝CY2026 … FY2030＝CY2029）"
hdr(UN, 3, ["項目", "單位", "CY2026（FY2027）", "CY2027（FY2028）", "CY2028（FY2029）", "CY2029（FY2030）", "說明"])
PRICE = {"bw": "pm_bw", "rb": "pm_rb", "ru": "pm_ru", "hp": "pm_hp", "amd": "pa_amd", "tpu": "pa_tpu", "trn": "pa_trn", "oth": "pa_oth"}
ycols = {y: cyc(y) for y in CYS}
SHR = {}; ur = 4
for s in SC:
    ttl = {"A": "情境 A：基準", "B": "情境 B：2028 顆數 ×0.8、2029 ×0.85", "C": "情境 C：2028–29 顆數 ×0.6"}[s]
    units = {(p, y): uexp(p, y, s) for p in PLAT for y in CYS}
    ur, SHR[s] = AS.write_share_block(UN, ur, ttl, CYS, ycols, units, PRICE, PLAT, PNAME)
UN.cell(ur, 1, "金額份額的口徑：各平台「模組層級」每顆價（NVIDIA＝GPU 模組價；AMD、ASIC＝每顆晶片價；皆 In 頁 B、J 節），不含系統加價與網通，故為加速器晶片的金額份額。ASIC 單價為 Assumed／Interested-party，份額的區間見 Sens（ASIC 金額份額 ±5 個百分點）。").font = BOLD
ur += 2
UN.cell(ur, 1, "每顆運算售價（基準，2026 記憶體價）＝GPU 模組＋CPU＋LPDDR＋k_tray × 匣級加價（US$K）").font = BOLD; ur += 1
hdr(UN, ur, ["平台", "單位", "模組", "CPU", "LPDDR", "匣級加價 × k_tray", "合計 pb"]); ur += 1
PBR = {}
for p in NVP:
    UN.cell(ur, 1, PNAME[p]); UN.cell(ur, 2, "US$K")
    for col, f in zip("CDEF", (f"=pm_{p}", f"=pcpu_{p}", f"=plp_{p}", f"=k_tray*pres_{p}")):
        UN[f"{col}{ur}"] = f; UN[f"{col}{ur}"].number_format = "0.00"
    UN[f"G{ur}"] = f"=SUM(C{ur}:F{ur})"; UN[f"G{ur}"].number_format = "0.00"
    PBR[p] = ur; ur += 1
ur += 1
CAL_ROW = ur   # 校準區塊（引擎建好後填）
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 16
for c in "CDEF": UN.column_dimensions[c].width = 18
UN.column_dimensions["G"].width = 70

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用，確保一致） ----------------
FYS = (2026, 2027, 2028, 2029, 2030); FC = {2026: "C", 2027: "D", 2028: "E", 2029: "F", 2030: "G"}
BU = (2027, 2028, 2029, 2030)   # 有「由下而上」推導的會計年度（CY＝FY−1）
DEFP = dict(scn="A", flat29=False, krec="k_rec", cnh="c_nh", uf="1", nvm=None, abs="abs_mem", cn="0", opxa="0", pf="1", pm_bw="pm_bw", pm_rb="pm_rb", pm_ru="pm_ru", pm_hp="pm_hp")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2026A", "FY2027", "FY2028", "FY2029", "FY2030", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    cy = lambda fy: fy - 1
    c_ = lambda fy: FC[fy]
    cyu = lambda fy: 2028 if (P["flat29"] and cy(fy) == 2029) else cy(fy)   # flat29：CY2029 顆數與 CY2028 持平（FY2030 顆數持平）
    sec_("A. 由下而上（C 欄空白；D＝FY2027≈CY2026 … G＝FY2030≈CY2029；顆數為「認列顆數」＝市場層顆數 × k_rec，FY2027 的 Data Center 營收與毛利率因此等於錨點，FY2028 起沿用 k_rec）")
    def ufx(fy):
        parts = [P["krec"]]
        if cy(fy) >= 2027:
            parts.append(P["uf"])
            if P["nvm"]: parts.append(P["nvm"](cy(fy)))
        return "*".join(parts) or "1"
    row("ufac", "NVIDIA 認列顆數乘數（k_rec × Sens 乘數）", "x", {c_(fy): "=" + ufx(fy) for fy in BU}, fmt="0.000")
    for p in NVP:
        row("u_" + p, f"{PNAME[p]} 認列顆數", "百萬顆", {c_(fy): f"=({uexp(p, cyu(fy), s)})*{c_(fy)}{R['ufac']}" for fy in BU}, fmt="0.00")
    row("nvu", "NVIDIA 認列顆數合計", "百萬顆", {c_(fy): f"=SUM({c_(fy)}{R['u_bw']}:{c_(fy)}{R['u_hp']})" for fy in BU}, fmt="0.00", note="2026＝7.6、2027＝9.7、2028＝9.5（A）")
    pb = lambda p: f"({P['pm_' + p]}+pcpu_{p}+plp_{p}+k_tray*pres_{p})"
    hb = lambda p, fy: f"hbm_{s}_{cy(fy) % 100}/hbm_{s}_26"
    cx = lambda fy: f"cdx_{s}_{cy(fy) % 100}/cdx_{s}_26"
    gbx = lambda p, fy: f"gb_{p}_{min(cy(fy), 2028) % 100}"
    for p in NVP:
        row("px_" + p, f"{PNAME[p]} 每顆售價", "US$K", {c_(fy): f"={P['pf']}*{pb(p)}+(1-{P['abs']})*({gbx(p, fy)}*hpx_{p}*({hb(p, fy)}-1)/1000+plp_{p}*({cx(fy)}-1))" for fy in BU}, fmt="0.00",
            note="基準價 ＋ (1−吸收比例) ×（HBM 與 LPDDR 相對 2026 的成本增加）" if p == "bw" else "")
    for p in NVP:
        row("hc_" + p, f"{PNAME[p]} 每顆 HBM 成本", "US$K", {c_(fy): f"={gbx(p, fy)}*hpx_{p}*{hb(p, fy)}/1000" for fy in BU}, fmt="0.00",
            note="容量 × 2026 合約價 × 情境別 HBM 價格指數（markets/memory）" if p == "bw" else "")
    for p in NVP:
        row("cg_" + p, f"{PNAME[p]} 每顆成本", "US$K", {c_(fy): f"={P['pf']}*{P['cnh']}*{pb(p)}+{c_(fy)}{R['hc_' + p]}+plp_{p}*({cx(fy)}-1)" for fy in BU}, fmt="0.00",
            note="非 HBM 成本（基準價 × c_nh，FY2027 校準）＋ HBM 成本 ＋ LPDDR 成本增加（每顆；顆數另乘 k_rec）" if p == "bw" else "")
    for p in NVP:
        row("r_" + p, f"{PNAME[p]} 運算營收", "US$B", {c_(fy): f"={c_(fy)}{R['u_' + p]}*{c_(fy)}{R['px_' + p]}" for fy in BU})
    row("comp", "運算營收（Σ 顆數 × 每顆售價）", "US$B", {c_(fy): f"=SUM({c_(fy)}{R['r_bw']}:{c_(fy)}{R['r_hp']})" for fy in BU})
    row("compc", "運算成本", "US$B", {c_(fy): "=" + "+".join(f"{c_(fy)}{R['u_' + p]}*{c_(fy)}{R['cg_' + p]}" for p in NVP) for fy in BU})
    row("hbmc", "其中 HBM 成本", "US$B", {c_(fy): "=" + "+".join(f"{c_(fy)}{R['u_' + p]}*{c_(fy)}{R['hc_' + p]}" for p in NVP) for fy in BU})
    row("cn", "中國資料中心運算營收（基準為 0＝公司指引假設；Sens 列恢復）", "US$B", {c_(fy): ("=" + P["cn"] if fy >= 2028 else "=0") for fy in BU})
    row("cnc", "中國部分成本（毛利率同運算）", "US$B", {c_(fy): f"={c_(fy)}{R['cn']}*{c_(fy)}{R['compc']}/{c_(fy)}{R['comp']}" for fy in BU})
    umx = lambda fy: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[cyu(fy)]
    row("net", "NVIDIA 網通營收（本模型歸屬 × 情境顆數係數 × 乘數）", "US$B", {c_(fy): f"=net_{cyu(fy) % 100}*{umx(fy)}*{c_(fy)}{R['ufac']}" for fy in BU}, note="markets/network；網通金額與 NVIDIA 顆數成正比")
    row("netc", "網通成本（c_net × 營收）", "US$B", {c_(fy): f"=c_net*{c_(fy)}{R['net']}" for fy in BU})
    row("edge", "Edge Computing 營收", "US$B", {"D": "=a27_edge", **{c_(fy): f"={c_(fy - 1)}{r}*(1+eg)" for fy in BU[1:]}})
    row("edgec", "Edge 成本", "US$B", {c_(fy): f"=(1-gm_edge)*{c_(fy)}{R['edge']}" for fy in BU})
    row("bur", "由下而上營收合計", "US$B", {c_(fy): f"={c_(fy)}{R['comp']}+{c_(fy)}{R['cn']}+{c_(fy)}{R['net']}+{c_(fy)}{R['edge']}" for fy in BU})
    row("buc", "由下而上成本合計", "US$B", {c_(fy): f"={c_(fy)}{R['compc']}+{c_(fy)}{R['cnc']}+{c_(fy)}{R['netc']}+{c_(fy)}{R['edgec']}" for fy in BU})
    row("bugm", "由下而上毛利率", "%", {c_(fy): f"=1-{c_(fy)}{R['buc']}/{c_(fy)}{R['bur']}" for fy in BU}, fmt="0.0%", note="FY2027 ＝ 錨點毛利率（c_nh 校準的恆等式，非獨立檢驗；對照區塊 k_rec＝1 時用 c_nh_p）")
    row("hbmr", "HBM 成本占由下而上營收", "%", {c_(fy): f"={c_(fy)}{R['hbmc']}/{c_(fy)}{R['bur']}" for fy in BU}, fmt="0.0%")
    sec_("B. 報告值（FY2026A、FY2027 為實績與錨點；FY2028 起為由下而上）")
    row("rev", "營收", "US$B", {"C": "=a26_rev", "D": "=a27_rev", **{c_(fy): f"={c_(fy)}{R['bur']}" for fy in BU[1:]}})
    row("dc", "  Data Center（運算＋網通＋中國）", "US$B", {"C": "=a26_dc", "D": "=a27_dc", **{c_(fy): f"={c_(fy)}{R['comp']}+{c_(fy)}{R['cn']}+{c_(fy)}{R['net']}" for fy in BU[1:]}})
    row("edger", "  Edge Computing", "US$B", {"C": "=a26_edge", "D": "=a27_edge", **{c_(fy): f"={c_(fy)}{R['edge']}" for fy in BU[1:]}})
    row("g", "營收年增", "%", {c_(fy): f"={c_(fy)}{R['rev']}/{c_(fy - 1)}{R['rev']}-1" for fy in FYS[1:]}, fmt="0.0%")
    row("gp", "毛利", "US$B", {"C": "=a26_gp", "D": "=a27_gp", **{c_(fy): f"={c_(fy)}{R['rev']}-{c_(fy)}{R['buc']}" for fy in BU[1:]}})
    row("gm", "毛利率", "%", {c_(fy): f"={c_(fy)}{R['gp']}/{c_(fy)}{R['rev']}" for fy in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a26_opx", "D": "=a27_opx", **{c_(fy): f"={c_(fy - 1)}{r}*(1+MAX(g_opx_min,phi*{c_(fy)}{R['g']}))" for fy in BU[1:]}}, note="前一年 ×（1＋MAX(下限, 彈性 × 營收年增)）")
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(fy): f"={c_(fy)}{R['opb']}+{P['opxa']}*{c_(fy)}{R['rev']}" for fy in BU[1:]}})
    row("oi", "營業利益", "US$B", {c_(fy): f"={c_(fy)}{R['gp']}-{c_(fy)}{R['opx']}" for fy in FYS})
    row("oim", "營業利益率", "%", {c_(fy): f"={c_(fy)}{R['oi']}/{c_(fy)}{R['rev']}" for fy in FYS}, fmt="0.0%")
    row("oth", "其他收益淨額（不含權益證券損益）", "US$B", {"C": "=a26_oth", "D": "=a27_oth", **{c_(fy): "=oth_f" for fy in BU[1:]}}, fmt="0.00")
    row("etr", "有效稅率", "%", {"C": f"=1-a26_ni/(a26_oi+a26_oth)", "D": "=a27_etr", **{c_(fy): "=tx_f" for fy in BU[1:]}}, fmt="0.0%")
    row("ni", "淨利（不含權益證券損益）", "US$B", {"C": "=a26_ni", "D": "=a27_ni", **{c_(fy): f"=({c_(fy)}{R['oi']}+{c_(fy)}{R['oth']})*(1-{c_(fy)}{R['etr']})" for fy in BU[1:]}})
    row("sh", "稀釋股數", "十億股", {"C": "=a26_sh", "D": "=a27_sh", **{c_(fy): f"={c_(fy - 1)}{r}*(1+sh_g)" for fy in BU[1:]}}, fmt="0.000")
    row("eps", "EPS（不含權益證券損益）", "US$", {c_(fy): f"={c_(fy)}{R['ni']}/{c_(fy)}{R['sh']}" for fy in FYS}, fmt="0.00")
    row("epsg", "EPS 年增", "%", {c_(fy): f"={c_(fy)}{R['eps']}/{c_(fy - 1)}{R['eps']}-1" for fy in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — NVIDIA 損益（US$B；FY2026A、FY2027 為實績與錨點，FY2028 起為模型，顆數用認列顆數＝市場層顆數 × k_rec）。FY(t)≈CY(t−1)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2027、FY2028 三情境幾乎相同（情境係數從 CY2028＝FY2029 才開始；FY2028 的差在 B、C 的 LPDDR 成本）"
PB = {}; r0 = 70
for s in SC:
    PB[s] = engine(PL, r0, {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}[s], scn=s)
    r0 = PB[s]["end"] + 2
PB0 = engine(PL, r0, "對照：v0.1 基準（k_rec＝1，未校準；FY2027 由下而上高於錨點；c_nh_p 為 v0.1 的毛利率校準）", scn="A", krec="1", cnh="c_nh_p")
hdr(PL, 3, ["情境／項目", "單位", "FY2026A", "FY2027", "FY2028", "FY2029", "FY2030", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("dc", "  Data Center", "US$B", "#,##0.0"), ("edger", "  Edge Computing", "US$B", "#,##0.0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "US$B", "#,##0.0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利（不含權益證券損益）", "US$B", "#,##0.0"),
                                ("eps", "EPS（不含權益證券損益）", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for fy in FYS:
            if key in ("g", "epsg") and fy == 2026: continue
            PL[f"{FC[fy]}{sr}"] = f"={FC[fy]}{PB[s][key]}"; PL[f"{FC[fy]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2027 由下而上毛利率 − 錨點毛利率；FY2027 由下而上營收 − 錨點營收（k_rec、c_nh 校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：毛利率差"); PL[f"D{sr}"] = f"=D{PB['A']['bugm']}-D{PB['A']['gm']}"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['bur']}-a27_rev"; PL[f"D{sr}"].number_format = "0.0000"
PL.column_dimensions["A"].width = 60; PL.column_dimensions["B"].width = 14; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 13

# ---------------- Units 的校準區塊（引擎建好後填）：k_rec、c_nh、c_nh_p、pf_cal ----------------
A_ = PB["A"]; Z_ = PB0
UN.cell(CAL_ROW, 1, "校準（只用公司已發生的實績與當年指引：FY2027 錨點的 Data Center 營收與毛利率；不用共識、不用公司對 FY2028 的預測）").font = BOLD
cal = CAL_ROW + 1
pbraw = "+".join(f"u_{p}_26*(pm_{p}+pcpu_{p}+plp_{p}+k_tray*pres_{p})" for p in NVP)
UN.cell(cal, 1, "NVIDIA 認列顆數係數 k_rec ＝ 錨點 Data Center 營收 ÷ （市場層顆數 × 基準售價的運算營收 ＋ 網通歸屬）（FY2027；營收與成本同比例縮放）"); UN.cell(cal, 2, "x")
UN[f"C{cal}"] = f"=a27_dc/({pbraw}+net_26)"; UN[f"C{cal}"].number_format = "0.0000"
UN.cell(cal, 7, "Derived：以已發生的實績校準，FY2028 起沿用。不改 markets/memory 的共用顆數（市場層顆數維持原值，只在本模型內縮放）。顆數層與售價層無法分開識別，k_rec 是兩者的合併係數；k_tray 變動時 k_rec 隨之重算，營收水準不變")
defname("k_rec", f"Units!$C${cal}"); UN[f"D{cal}"] = "Derived"; K_ROW = cal
pbu = "+".join(f"PL!D{A_['u_' + p]}*(pm_{p}+pcpu_{p}+plp_{p}+k_tray*pres_{p})" for p in NVP)
UN.cell(cal + 1, 1, "非 HBM 成本占基準售價 c_nh ＝ [(1−錨點毛利率) × 由下而上營收 − HBM 成本 − 網通成本 − Edge 成本] ÷ Σ 認列顆數 × 基準售價（FY2027）"); UN.cell(cal + 1, 2, "x")
UN[f"C{cal+1}"] = f"=((1-a27_gp/a27_rev)*PL!D{A_['bur']}-PL!D{A_['hbmc']}-PL!D{A_['netc']}-PL!D{A_['edgec']})/({pbu})"
UN[f"C{cal+1}"].number_format = "0.0000"; UN.cell(cal + 1, 7, "Derived：讓 FY2027 由下而上的毛利率等於錨點毛利率；FY2028 起的毛利率變化來自 HBM／LPDDR 成本、世代與顆數組合、吸收比例")
defname("c_nh", f"Units!$C${cal+1}")
pbz = "+".join(f"PL!D{Z_['u_' + p]}*(pm_{p}+pcpu_{p}+plp_{p}+k_tray*pres_{p})" for p in NVP)
UN.cell(cal + 2, 1, "c_nh_p（v0.1 的毛利率校準，k_rec＝1；只供 Sens 對照案例用）"); UN.cell(cal + 2, 2, "x")
UN[f"C{cal+2}"] = f"=((1-a27_gp/a27_rev)*PL!D{Z_['bur']}-PL!D{Z_['hbmc']}-PL!D{Z_['netc']}-PL!D{Z_['edgec']})/({pbz})"
UN[f"C{cal+2}"].number_format = "0.0000"; UN.cell(cal + 2, 7, "Derived；公式同 c_nh，但取對照區塊（未校準顆數）")
defname("c_nh_p", f"Units!$C${cal+2}")
UN.cell(cal + 3, 1, "FY2027 未校準由下而上 Data Center 營收（k_rec＝1；運算＋網通）"); UN.cell(cal + 3, 2, "US$B")
UN[f"C{cal+3}"] = f"=PL!D{Z_['comp']}+PL!D{Z_['net']}"; UN[f"C{cal+3}"].number_format = "0.0"
UN.cell(cal + 4, 1, "FY2027 錨點 Data Center 營收（公司）"); UN.cell(cal + 4, 2, "US$B"); UN[f"C{cal+4}"] = "=a27_dc"; UN[f"C{cal+4}"].number_format = "0.0"
UN.cell(cal + 5, 1, "未校準差額（未校準 ÷ 錨點 − 1）"); UN.cell(cal + 5, 2, "%"); UN[f"C{cal+5}"] = f"=C{cal+3}/C{cal+4}-1"; UN[f"C{cal+5}"].number_format = "0.0%"
UN.cell(cal + 5, 7, "差額來源待查（顆數層或售價層）：見 Recon B2 節")
UN.cell(cal + 6, 1, "FY2027 認列後 Data Center 營收（k_rec 校準後；應等於錨點）"); UN.cell(cal + 6, 2, "US$B")
UN[f"C{cal+6}"] = f"=PL!D{A_['comp']}+PL!D{A_['net']}"; UN[f"C{cal+6}"].number_format = "0.0"
UN.cell(cal + 7, 1, "售價整體係數 pf_cal ＝ (錨點 Data Center − 未校準網通) ÷ 未校準運算（Derived；v0.1 的作法，只在 Sens 當對照案例，不進基準）"); UN.cell(cal + 7, 2, "x")
UN[f"C{cal+7}"] = f"=(a27_dc-PL!D{Z_['net']})/PL!D{Z_['comp']}"; UN[f"C{cal+7}"].number_format = "0.000"
defname("pf_cal", f"Units!$C${cal+7}")

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。結論用「現價 ÷ FY2028 EPS（模型與共識）」；10 年落後 P/E 規則值不適用（保留計算，不進結論）。跨公司規則：合理倍數＝10 年（FY2016–FY2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。NVIDIA 的 FY(t) 對照日曆年 t−1，循環位置用共用規則的 lag 包裝（tools/cycle_position.py classify_lag），見頂年度讀設備市場模型"
V["A2"] = "不適用的理由：歷史倍數是「會計年度年底價 ÷ 該年度（剛結束）GAAP 稀釋 EPS」＝落後（trailing）倍數；規則把它乘在 FY2028 EPS 上，得到的是約 2028-01 的價格（距 2026-10-22 約 1.27 年，未折現；折現值另列）。歷史倍數含高成長期的高估值（FY2023 因 EPS 低谷達 114x），規則值遠高於現價——見下方「What must be true」"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-10-23 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "0.00"
V.cell(7, 1, "現價 ÷ FY2028 EPS（模型）")
V.cell(8, 1, "現價 ÷ FY2029 EPS（模型）")
V.cell(9, 1, "現價 ÷ 共識 FY2028 EPS")
V.cell(10, 1, "命題：現價約為「2027 年 EPS」的倍數（工作單：13.5x；2027 年＝FY2028）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    V[f"{c}7"] = f"=px/PL!E{SUMR[(s, 'eps')]}"; V[f"{c}8"] = f"=px/PL!F{SUMR[(s, 'eps')]}"
    for rr in (7, 8): V[f"{c}{rr}"].number_format = "0.0"
V["B9"] = "=px/cons_eps28"; V["B9"].number_format = "0.0"; V["B10"] = "=th_mult"; V["B10"].number_format = "0.0"
V["E7"] = "FY2028 EPS 三情境幾乎相同（情境係數從 FY2029 才開始；差在 B、C 的 LPDDR 成本），故 A、B、C 的 FY2028 倍數相近"
V["E9"] = "共識 EPS 為新 non-GAAP 口徑（含 SBC），與本模型口徑近似（Q 頁檢查列）"

V.cell(12, 1, "2. 循環位置（共用規則的 lag 版）與規則倍數").font = BOLD
V.cell(13, 1, "規則：各情境見頂年度（日曆年）取自設備市場模型（Mkt 頁「見頂年度」：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + "）；NVIDIA 的 FY(t) 對照日曆年 t−1：t−1 早於見頂年＝中段（中位數）、等於＝見頂（25 分位）、晚於＝谷底（75 分位）。擁有者 2026-10-23 核准：FY2028 在 A、B 為中段，C 為見頂。")
hdr(V, 14, ["情境", "設備市場見頂年度（CY）", "FY2028 對照 CY", "FY2028 判定", "FY2028 倍數", "FY2029 對照 CY", "FY2029 判定", "FY2029 倍數", "規則"])
for i, s in enumerate(SC, 15):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}")
    V.cell(i, 3, "=2028-1")
    V.cell(i, 4, "=" + CP.excel_rule_lag("2028", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "=2029-1")
    V.cell(i, 7, "=" + CP.excel_rule_lag("2029", f"pk_{s}"))
    V.cell(i, 8, f'=IF(G{i}="見頂",pe_peak,IF(G{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 9, "tools/cycle_position.py classify_lag（lag＝1）")
V.cell(19, 1, "對照欄（只作對照，不當主值）：以 NVIDIA 本模型營收路徑自己的見頂年判斷循環位置").font = BOLD
hdr(V, 20, ["情境", "營收見頂 FY（FY2027–FY2030 內）", "對照 CY", "FY2028 判定", "FY2028 倍數", "觀測期是否截尾", "說明"])
for i, s in enumerate(SC, 21):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2027,2028,2029,2030}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, f"=B{i}-1")
    V.cell(i, 4, "=" + CP.excel_rule("(2028-1)", f"C{i}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, f'=IF(B{i}=2030,"是：營收到觀測期末仍在增加，見頂年未知","否")')
    V.cell(i, 7, "FY(t) 對照 CY t−1；以本模型營收最大的會計年度為見頂" if s == "A" else "同上")
V.cell(25, 1, "3. 規則值隱含股價（美元）——不適用：落後倍數含高成長期溢價，保留計算但不進結論；規則值＝FY2028 EPS × 依循環位置的倍數").font = BOLD
hdr(V, 26, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價（不適用；約 FY2028 年底，未折現）", "相對現價", "折現到現在"])
vr = 27; VAL = {}
for j, s in enumerate(SC):
    for y, typ, nm in ((2028, f"E{15 + j}", "rule"), (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位"),
                       (2029, f"H{15 + j}", "rule29")):
        col = FC[y]
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{col}{SUMR[(s, 'eps')]}").number_format = "0.00"
        V.cell(vr, 4, f'="規則值（不適用）："&D{15 + j}' if nm == "rule" else (f'="規則值（FY2029；不適用）："&G{15 + j}' if nm == "rule29" else nm))
        V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        V.cell(vr, 8, f"=F{vr}/(1+r_disc)^(1.27+{y - 2028})").number_format = "#,##0"
        if nm in ("rule", "rule29"):
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[(s, y)] = vr
        vr += 1
V.cell(vr, 1, "FY2029 折現年數＝1.27＋1（FY2029 年底約 2029-01）；折現率為 In 頁 r_disc（Assumed 8–12%）。").font = Font(italic=True)
vr += 2
V.cell(vr, 1, "4. What must be true：要讓 FY2028 年底價等於現價 ×（1＋報酬），需要的倍數（情境 A 的 FY2028 EPS）").font = BOLD; vr += 1
hdr(V, vr, ["FY2028 年底隱含價 ÷ 現價", "所需倍數", "對照：歷史 25 分位", "對照：歷史中位數", "對照：歷史 75 分位"]); vr += 1
WMT0 = vr
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!E{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, "=pe_peak").number_format = "0.0"; V.cell(vr, 4, "=pe_mid").number_format = "0.0"; V.cell(vr, 5, "=pe_trough").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝兩年約 10% 年報酬。現價 ÷ FY2028 EPS 約 15x，遠低於歷史年底倍數的 25 分位；規則值反映的是 NVIDIA 過去的高成長溢價，不是它在 EPS 成長放緩時仍會被給的倍數。")
vr += 2
assert vr < 70, vr
# 十年倍數歷史與規則值（tools/valuation_rule.py）
PE_MT = [None, 42.09, 50.33, 21.47, 51.97, 74.74, 63.40, 114.67, 51.61, 40.78]   # Macrotrends Jan-31 P/E，FY2016–FY2025
PE_FR = [24.4, 34.93, 46.79, 23.55, 52.18, 72.17, 58.56, None, None, None]        # fullratio 日曆年年底 P/E（2016–2022；日期口徑不同，僅參考）
block_end = VR.add_rule_block(V, 70, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "Macrotrends（1 月 31 日 ÷ TTM GAAP EPS）", "vals": PE_MT, "tag": "Interested-party（資料商；搜尋摘要，頁面無法開啟）；FY2016 未找到；FY2017 取自另一摘要；FY2023 114.67（區間 112.19–114.67）；與 Wisesheets 對照：FY2019 20.29、FY2023 106.44、FY2024 51.63"},
    {"label": "fullratio（日曆年年底 P/E，僅參考）", "vals": PE_FR, "tag": "Interested-party（日期口徑為 12 月底而非會計年度年底；2016–2022 共 7 年；不採用）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2028 EPS（A）", "basis": f"=PL!E{SUMR[('A', 'eps')]}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2028 EPS（A）", "basis": f"=PL!E{SUMR[('A', 'eps')]}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2028 EPS（A）", "basis": f"=PL!E{SUMR[('A', 'eps')]}"}]}],
    "0.0", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。本模型來源：只有 Macrotrends 一個序列可用（9 年，缺 FY2016）；第二個來源（fullratio）為日曆年年底口徑，只列參考。自算（會計年度年底收盤價 ÷ GAAP 稀釋 EPS）的收盤價在建置環境找不到（未找到），見 Sources 頁待補。")
rule_rows = [rw for rw in range(70, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
V.cell(block_end + 1, 1, "自算 P/E 的缺口：會計年度年底收盤價（FY2016–FY2025）未找到；已找到公司 GAAP 稀釋 EPS（as reported／分割調整：見下表），待補收盤價後可自算第二來源").font = BOLD
hdr(V, block_end + 2, ["會計年度", "GAAP 稀釋 EPS（原報）", "分割除數", "分割調整後 EPS", "non-GAAP EPS（調整後）", "來源"])
EPSH = [("FY2016", 1.08, 40, 1.67), ("FY2017", 2.57, 40, 3.06), ("FY2018", 4.82, 40, 4.92), ("FY2019", 6.63, 40, 6.64), ("FY2020", 4.52, 40, 5.79),
        ("FY2021", 1.73, 10, 2.50), ("FY2022", 3.85, 10, 4.44), ("FY2023", 1.74, 10, 3.34), ("FY2024", 11.93, 10, 12.96), ("FY2025", 2.94, 1, 2.99)]
for i, (fy_, e, dv, ng) in enumerate(EPSH):
    rw = block_end + 3 + i
    V.cell(rw, 1, fy_); V.cell(rw, 2, e).font = BLUE; V.cell(rw, 3, dv).font = BLUE
    V.cell(rw, 4, f"=B{rw}/C{rw}"); V.cell(rw, 5, f"={ng}/C{rw}")
    V.cell(rw, 6, "NVIDIA 第四季新聞稿（nvidianews.nvidia.com）；FY2016–FY2020 原報為 4:1 與 10:1 分割前（除數 40）；FY2021–FY2024 原報已含 4:1、未含 10:1（除數 10）；FY2025 已含兩次分割（除數 1）。non-GAAP 為舊口徑（排除 SBC）。FY2016 取自搜尋摘要")
    for c in "BDE": V[f"{c}{rw}"].number_format = "0.000"
V.column_dimensions["A"].width = 52
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 18
V.column_dimensions["K"].width = 60

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A，k_rec 校準後的基準）：對 FY2028、FY2029、FY2030 的營收、EPS 與隱含股價（規則值，不適用，僅供對照）的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）。隱含股價＝EPS × 情境 A 的規則倍數（FY2028：Val!E15；FY2029：Val!H15）"
hdr(SE, 3, ["案例", "設定", "FY2028 營收", "FY2028 EPS", "FY2028 EPS 相對基準", "FY2028 隱含價（規則值，不適用，僅供對照）", "隱含價相對基準", "FY2029 營收", "FY2029 EPS", "FY2029 EPS 相對基準", "FY2029 隱含價（規則值，不適用）", "隱含價相對基準", "FY2030 EPS", "FY2030 EPS 年增（對 FY2029）"])
sh_nv = lambda cy: f"((Units!{cyc(cy)}{SHR['A'][('ds', 'nv')]}+{{d}})/Units!{cyc(cy)}{SHR['A'][('ds', 'nv')]})"
PFCASE = "對照：v0.1 的作法——只調售價（pf_cal；k_rec＝1）"
cases = [("NVIDIA 顆數係數 低（×0.9）", "FY2028 起所有 NVIDIA 世代顆數 ×0.9", dict(uf="uk_lo")),
         ("NVIDIA 顆數係數 高（×1.1）", "×1.1", dict(uf="uk_hi")),
         ("Rubin Ultra 售價 低", "GPU 模組價 60K", dict(pm_ru="pm_ru_lo")), ("Rubin Ultra 售價 高", "GPU 模組價 100K", dict(pm_ru="pm_ru_hi")),
         ("ASIC 金額份額 +5pt（NVIDIA 份額 −5pt）", "NVIDIA 顆數 ×(份額−5pt)÷份額（Units 頁金額份額），總市場金額不變", dict(nvm=lambda cy: sh_nv(cy).format(d="-dsh"))),
         ("ASIC 金額份額 −5pt（NVIDIA 份額 +5pt）", "NVIDIA 顆數 ×(份額+5pt)÷份額", dict(nvm=lambda cy: sh_nv(cy).format(d="dsh"))),
         ("HBM／LPDDR 吸收比例 0（全額轉嫁）", "吸收比例 0.0", dict(abs="abs_mem_lo")), ("HBM／LPDDR 吸收比例 1（全額吸收）", "吸收比例 1.0", dict(abs="abs_mem_hi")),
         ("中國恢復 低（10）", "FY2028 起每年加 10", dict(cn="cn_amt_lo")), ("中國恢復 基準（20）", "每年加 20", dict(cn="cn_amt")), ("中國恢復 高（40）", "每年加 40", dict(cn="cn_amt_hi")),
         ("營業費用率 −2pt", "營業費用額外 −2% × 營收", dict(opxa="opx_adj_lo")), ("營業費用率 +2pt", "+2% × 營收", dict(opxa="opx_adj_hi")),
         ("CY2029 顆數與 CY2028 持平（FY2030 顆數持平）", "核心問題：顆數持平以後 EPS 還能不能成長；網通同步持平", dict(flat29=True)),
         (PFCASE, "全部運算售價（與非 HBM 成本）× pf_cal，顆數不縮放：讓 FY2027 的 Data Center 等於錨點；FY2028 毛利率與公司指引不一致（見 Recon 對應列）；不進基準", dict(pf="pf_cal", krec="1", cnh="c_nh_p"))]
SE.cell(4, 1, "基準（情境 A）"); SE.cell(4, 2, "PL 區塊 A")
base_row = 40
SB = {}
for n, (nm, setting, ov) in enumerate(cases):
    r0_ = base_row
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}E{R['rev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"H{i}"] = f"={pre}F{R['rev']}"; SE[f"I{i}"] = f"={pre}F{R['eps']}"
    SE[f"M{i}"] = f"={pre}G{R['eps']}"; SE[f"N{i}"] = f"=M{i}/I{i}-1"; SE[f"M{i}"].number_format = "0.00"; SE[f"N{i}"].number_format = "0.0%"
    SE[f"F{i}"] = f"=D{i}*Val!$E$15"; SE[f"K{i}"] = f"=I{i}*Val!$H$15"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"J{i}"] = f"=I{i}/I$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "CH": SE[f"{c}{i}"].number_format = "#,##0.0"
    for c in "DI": SE[f"{c}{i}"].number_format = "0.00"
    for c in "FK": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "EGJL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85；記憶體 B 路徑", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6；記憶體 C 路徑", PB["C"], "PL!")
put_case(nn + 2, "對照：v0.1 基準（k_rec＝1，未校準）", "FY2027 由下而上高於錨點；FY2028 起沿用未校準水準", PB0, "PL!")
SE.cell(nn + 4, 1, "讀法：FY2028（≈CY2027）三情境幾乎相同；顆數係數、售價與吸收比例的影響從 FY2028 起；情境 B、C 的差異從 FY2029 起。隱含價欄為規則值（10 年落後 P/E 分位，不適用，僅供對照），只反映 EPS 變動，不反映循環位置變動。")
SENS_N = len(cases)
SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKLMN": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識、市場層。只作事後對照，不為對上數字而調整參數。公司指引（FY2028 營收年增約 70%、毛利率 72–73%）與共識只放在這裡，不用於校準"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
RFMT = {}
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False, fmts=None):
    global rr_
    RFMT[rr_] = (fmt, ref is not None)
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2028 vs 公司指引與共識（校準時未使用）").font = BOLD; rr_ += 1
rec("FY2028 營收 vs 指引推算（FY2027 錨點 ×(1＋70%)）", "=a27_rev*(1+g28_gr)", pl("rev", 2028), "公司 Q2 FY27 法說（二手）；美元金額：未找到公司直接陳述", "工作單記約 US$690B；本列用錨點 × 1.7 推算")
rec("FY2028 營收年增 vs 指引 70%（差距以百分點計）", "=g28_gr", pl("g", 2028), "同上", "", fmt="0.0%", pt=True)
rec("FY2028 營收 vs 共識推算（共識 FY2027 營收 ×(1＋63%)）", "=cons_rev27*(1+cons_g28)", pl("rev", 2028), "stockanalysis／Yahoo（二手）", "共識 FY2028 營收直接值：未找到")
rec("FY2028 毛利率 vs 指引 72–73% 中點（百分點）", "=g28_gm", pl("gm", 2028), "二手報導；逐字稿未取得", "", fmt="0.0%", pt=True)
rec("FY2027 錨點 EPS vs 共識 FY2027 EPS", "=cons_eps27", pl("eps", 2027), "stockanalysis（二手）", "")
rec("FY2028 EPS vs 共識 FY2028 EPS", "=cons_eps28", pl("eps", 2028), "247wallst／Wallstreetzen（二手）", "共識區間 15.70–16.13（In 頁）", fmt="0.00")
rec("FY2028 毛利率：v0.1 對照案例（只調售價 pf_cal）vs 指引 72–73% 中點（百分點）", "=g28_gm", lambda s: f"=Sens!E{SB[PFCASE]['gm']}", "Sens 對照案例（只調售價，顆數不縮放）", "與公司指引不一致：只調售價時 HBM 成本相對營收被放大，FY2028 毛利率偏低", fmt="0.0%", pt=True)
rec("FY2029 EPS（共識：未找到）", None, pl("eps", 2029), "—", "WallStreetSmart 自有模型 FY2029 EPS 20.70（非共識，不採）", fmt="0.00")
RC.cell(rr_, 1, "B. FY2027 水準差與 k_rec（v0.1.1：基準已用 k_rec 校準到錨點；未校準的差額保留供查）").font = BOLD; rr_ += 1
ABU = PB["A"]; ZBU = PB0
rec("Data Center 營收：未校準由下而上（k_rec＝1；運算＋網通）vs 公司錨點", "=a27_dc", lambda s: f"=PL!D{ZBU['comp']}+PL!D{ZBU['net']}", "公司 Q1、Q2 實績＋Q3 指引＋Q4 假設（v0.1 基準的差額）", "FY2027 三情境相同（情境係數從 CY2028 起）")
rec("認列顆數係數 k_rec（Derived）", None, lambda s: "=k_rec", "錨點 Data Center ÷ 未校準由下而上", "以已發生的實績校準，FY2028 起沿用；營收與成本同比例縮放；顆數層與售價層無法分開識別", fmt="0.000")
rec("認列後 FY2027 Data Center 營收 vs 錨點（應為 0）", "=a27_dc", lambda s: f"=PL!D{ABU['comp']}+PL!D{ABU['net']}", "校準恆等式，非獨立檢驗", "")
rec("認列後 FY2027 毛利率 vs 錨點毛利率（應為 0；百分點）", "=a27_gp/a27_rev", lambda s: f"=PL!D{ABU['bugm']}", "c_nh 校準恆等式，非獨立檢驗", "", fmt="0.0%", pt=True)
rec("認列顆數（百萬顆）：k_rec × 市場層 vs 市場層 CY2026 NVIDIA 顆數", "=u_bw_26+u_rb_26+u_ru_26+u_hp_26", lambda s: f"=PL!D{ABU['nvu']}", "市場層 UNITS（markets/memory，未修改）", "認列顆數低於市場層＝顆數層與售價層的合併下調", fmt="0.00")
rec("  網通：本模型歸屬（未縮放）vs 實際推估 rep_nv26（markets/network 的對帳）", "=rep_nv26", lambda s: f"=PL!D{ZBU['net']}", "markets/network v0.2.2 Recon：−16.6%", "歸屬比例為 Assumed")
rec("  網通：本模型歸屬（未縮放）vs Q1 FY27 實際年化（14.8 × 4）", "=net_q1*4", lambda s: f"=PL!D{ZBU['net']}", "公司 Q1 FY27 稿", "網通季增中，年化為下限")
rec("  運算：未校準由下而上 vs 錨點 Data Center − 網通推估（rep_nv26）", "=a27_dc-rep_nv26", lambda s: f"=PL!D{ZBU['comp']}", "網通推估 rep_nv26＝74（Q1 實際 14.8＋推估）", "把網通拿掉後的運算差距（B2 節各候選來源以此為缺口）")
RC.cell(rr_, 1, "B2. 差額來源待查（以上一列的運算缺口為準；各列為候選來源，未驗證）").font = BOLD; rr_ += 1
hdr(RC, rr_, ["候選來源", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "目前狀態", "可查的公開資料"]); rr_ += 1
GAP = lambda: f"(PL!D{ZBU['comp']}-(a27_dc-rep_nv26))"
rec("① Rubin CY2026 出貨時點（FY↔CY）：缺口相當於 Rubin 顆數（百萬）；對照值＝Rubin CY2026 顆數", "=u_rb_26", lambda s: f"={GAP()}/PL!D{ZBU['px_rb']}",
    "未驗證。二手報導：Rubin 量產出貨自日曆年 2026 第三季開始、第四季放量（digitalcitizen、winbuzzer、經濟日報經 wccftech）；日曆年第三季多數落在 FY2027 第二、三季，只有 2027 年 1 月落在 FY2027 之外，時點本身應只解釋小部分缺口（本模型推論，未驗證）",
    "可查：公司法說與 10-Q 的 Rubin 營收認列時點與出貨顆數（財務長書面說明、10-Q 本環境無法開啟）；TrendForce 的 Rubin 季別出貨；ODM（鴻海、廣達）月營收與機櫃出貨", fmt="0.00")
rec("① 同上：缺口占 Rubin CY2026 顆數的比例（需落在 FY2027 之外的比例）", None, lambda s: f"={GAP()}/PL!D{ZBU['px_rb']}/u_rb_26", "—", "比例愈高，愈難只用出貨時點解釋", fmt="0.0%")
rec("② B200／B300 價差：僅靠 Blackwell 均價下調解釋缺口，所需降幅", None, lambda s: f"={GAP()}/PL!D{ZBU['r_bw']}",
    "未驗證。公司不分 B200／B300 揭露；橋接表 Rack 頁只有 GB300（模組價 40.4K）。二手：HSBC 稱 B200→B300（GB200→GB300）沒有明顯 ASP 提升（wccftech 轉述，原報告未取得）→ 價差本身不足以解釋缺口",
    "可查：賣方 ASP 估計（HSBC、KeyBanc）、B200 HGX 8 卡系統報價（經銷商）、GB200 機架價與 GB300 機架價對照；目前只找到機架層級 $3.7–4.0M（blog，無原始出處）", fmt="0.0%")
rec("③ 匣級加價保留比例 k_tray 由基準降到 0：未校準運算營收變動占缺口的比例", None, lambda s: f"=k_tray*(u_bw_26*pres_bw+u_rb_26*pres_rb+u_ru_26*pres_ru+u_hp_26*pres_hp)/{GAP()}",
    "Assumed（0–1）。公司不揭露匣級加價與 ODM 分成", "可查：NVIDIA 與 ODM（鴻海、廣達、緯創）的毛利率與營收認列方式（是否以機櫃或 GPU 板認列）", fmt="0.0%")
rec("④ 顆數層（市場層 UNITS 為賣方與研調估計）：全歸顆數層時，CY2026 NVIDIA 顆數需為（百萬顆）；對照值＝市場層", "=u_bw_26+u_rb_26+u_ru_26+u_hp_26", lambda s: f"=(a27_dc-rep_nv26)/(PL!D{ZBU['comp']}/PL!D{ZBU['nvu']})",
    "未驗證。市場層顆數為 Morgan Stanley、TrendForce 等估計（Interested-party）", "可查：台積電 CoWoS 出貨與 NVIDIA 分配、各研調的 2026 出貨更新；公司不揭露顆數", fmt="0.00")
rec("⑤ 全歸售價層時，每顆平均運算售價（US$K）；對照值＝模型（未校準）", f"=PL!D{ZBU['comp']}/PL!D{ZBU['nvu']}", lambda s: f"=(a27_dc-rep_nv26)/PL!D{ZBU['nvu']}", "未驗證", "同②③", fmt="0.0")
RC.cell(rr_, 1, "C. 網通與記憶體").font = BOLD; rr_ += 1
rec("HBM 價格指數 CY2027 ÷ CY2026：本模型（markets/memory）vs 橋接表 VR200 2027 合約（36 ÷ 16.7）", f"={BR['hbm_px'][3]}/{BR['hbm_px'][1]}", lambda s: f"=hbm_{s}_27/hbm_{s}_26", "橋接表 Rack 頁 D4／C4（Tokenomics v20.1）", "市場層混合價 +80%；橋接表 HBM4 合約 +116%：吸收時毛利率壓力的上限", fmt="0.00")
rec("HBM 價格 CY2027 ($/GB)：市場層 vs 二手 HBM4 估計（Bernstein 約 $53、Fubon $31–32、TrendForce 約 $36）", 36, lambda s: f"=hbm_{s}_27", "Silicon Analysts／TrendForce／Bernstein（二手；賣方與記憶體廠有價格方向誘因）", "市場層為混合單價（含 HBM3E）", fmt="0.0")
RC.cell(rr_, 1, "D. 命題檢查（不為對上而調參數）").font = BOLD; rr_ += 1
rec("2027 年 NVIDIA 金額份額（加速器晶片，模組層級）vs 命題下緣 75%", "=th_sh_lo", lambda s: f"=Units!{cyc(2027)}{SHR[s][('ds', 'nv')]}", "擁有者命題；外部來源：未找到", "區間見 Sens（ASIC 單價 Assumed）；命題上緣 80%", fmt="0.0%", pt=True)
rec("2027 年 NVIDIA 金額份額 vs 命題上緣 80%", "=th_sh_hi", lambda s: f"=Units!{cyc(2027)}{SHR[s][('ds', 'nv')]}", "同上", "", fmt="0.0%", pt=True)
rec("2027 年 NVIDIA 顆數份額（對照 DIGITIMES：約 35% 的 2027 顆數）", 0.351, lambda s: f"=Units!{cyc(2027)}{SHR[s][('us', 'nv')]}", "DIGITIMES Research 2026-10-13（二手）", "顆數份額與金額份額差距大＝NVIDIA 單價遠高於 ASIC", fmt="0.0%", pt=True)
rec("FY2028 毛利率 − FY2027 毛利率（百分點；命題：HBM 漲價壓毛利 2–3 個百分點）", -0.025, lambda s: f"=PL!E{SUMR[(s, 'gm')]}-PL!D{SUMR[(s, 'gm')]}", "擁有者命題（中點 −2.5pt）", "含吸收比例、世代組合與 Rubin 較高售價；HBM 成本占營收的變化見下一列", fmt="0.0%", pt=True)
rec("HBM 成本占由下而上營收：FY2028 − FY2027（百分點，吸收前）", None, lambda s: f"=PL!E{PB[s]['hbmr']}-PL!D{PB[s]['hbmr']}", "—", "吸收比例 0.5 時約一半轉嫁到售價", fmt="0.0%")
rec("現價 ÷ FY2028 EPS vs 命題 13.5x", "=th_mult", lambda s: f"=px/PL!E{SUMR[(s, 'eps')]}", "擁有者命題；共識 EPS 下約 13.4x", "EPS 越高，現價越便宜", fmt="0.0")
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以 NVIDIA 資料中心營收 ÷ 部署 GW 對照 IT capex"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 78
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 46; RC.column_dimensions["J"].width = 90

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝NVIDIA 8-K 新聞稿，已逐頁抓取；CFO 書面說明、10-Q、法說逐字稿在建置環境無法開啟，其內容只經搜尋摘要）"; SO["A1"].font = BOLD
srcs = [("公司：Q2 FY27 新聞稿（2026-08-26）", PR + "second-quarter-fiscal-2027"), ("公司：Q1 FY27 新聞稿（2026-05-20）", PR + "first-quarter-fiscal-2027"),
        ("公司：Q4 FY26 新聞稿", PR + "fourth-quarter-and-fiscal-2026"), ("公司：Q3 FY26 新聞稿", PR + "third-quarter-fiscal-2026"),
        ("公司：Q2 FY26 新聞稿", PR + "second-quarter-fiscal-2026"), ("公司：Q1 FY26 新聞稿", PR + "first-quarter-fiscal-2026"),
        ("公司：Q4 FY25 新聞稿（FY2025 全年；EPS 歷史）", PR + "fourth-quarter-and-fiscal-2025"),
        ("公司：FY2016–FY2024 第四季新聞稿（EPS 歷史）", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-fourth-quarter-and-fiscal-2024 等"),
        ("公司：Q1 FY27 CFO 書面說明（網通 14.8 之外的拆分；無法開啟：未找到）", "https://s201.q4cdn.com/141608511/files/doc_financials/2027/Q127/Q1FY27-CFO-Commentary.pdf"),
        ("二手：Q2 FY27 法說摘要（FY2028 營收約 +70%、毛利率路徑）", "https://finance.yahoo.com/technology/ai/articles/nvidia-q2-earnings-call-highlights-230417656.html"),
        ("二手：CNBC（70% 成長）", "https://www.cnbc.com/2026/08/26/nvidia-70percent-growth-forecast-puts-it-on-track-to-be-tech-no-2-company.html"),
        ("二手：implicator.ai（毛利率路徑 Q4 FY27 71–72%、FY2028 72–73%）", "https://www.implicator.ai/nvidia-70-percent-growth-forecast-margin-outlook/"),
        ("二手：Webull（Q3 指引 108B）", "https://www.webull.com/blog/304-Nvidia-Q2-FY2027-Earnings-Beats-Revenue-EPS-Estimates-Guides-Q3-to-108B"),
        ("二手：247wallst（共識 FY2028 EPS 15.70–15.91）", "https://247wallst.com/investing/2026/10/19/nvidia-stock-could-be-headed-for-a-significant-revaluation/"),
        ("二手：Wallstreetzen（FY2028 EPS 平均 16.13）", "https://www.wallstreetzen.com/stocks/us/nasdaq/nvda/stock-forecast"),
        ("二手：stockanalysis（共識 FY2027、現價、遠期 P/E）", "https://stockanalysis.com/stocks/nvda/forecast/"),
        ("二手：Macrotrends（會計年度年底 P/E）", "https://www.macrotrends.net/stocks/charts/NVDA/nvidia/pe-ratio"),
        ("二手：Wisesheets（P/E 對照）", "https://www.wisesheets.io/pe-ratio/NVDA"), ("二手：fullratio（日曆年年底 P/E）", "https://fullratio.com"),
        ("二手：HSBC MI355X $23K（finviz）", "https://finviz.com/news/165555/hsbc-trims-amd-target-amid-revised-pricing-assumptions"),
        ("二手：BofA Rubin Ultra 機架約 $21M（wccftech）", "https://wccftech.com/nvidia-rubin-ultra-rack-estimated-to-cost-21-million-hbm4e-swelling-to-1-5m-per-unit/"),
        ("二手：Bernstein VR200 $9.1M、HBM4 $53/GB（wccftech）", "https://wccftech.com/bernstein-warns-nvidias-vera-rubin-racks-will-hit-9-1-million-as-hbm4-prices-triple-to-53-per-gigabyte/"),
        ("二手：TrendForce／Silicon Analysts（HBM4 堆疊 $600→$1,300）", "https://siliconanalysts.com/market/trendforce-lifts-2027-hbm-asp-forecast-to-121-hbm4-stack-seen-doubling-to-1-300-2026-10-01"),
        ("二手：DIGITIMES（2027 ASIC 15.26M 顆 vs NVIDIA 約 9.1M）", "https://www.digitimes.com/reports/item.php?id=20261013RS400"),
        ("二手：Kaiso（2028 前 NVIDIA 份額 60–75%）", "https://www.kaisoresearch.com/blog/ai-accelerators-market-431b-by"),
        ("市場層：markets/memory/market.py（UNITS）；tsmc、skhynix、markets/network、markets/chip_bridge、markets/equipment 成品（本 repo）", ""),
        ("待補（人工）：FY2016–FY2025 會計年度年底收盤價（Yahoo／Nasdaq 歷史價）；CFO 書面說明；Q2 FY27 法說逐字稿", "未找到"),
        ("SemiAnalysis", "未單獨使用（僅作為 Ironwood 租用價的二手出處，未採）")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("NVIDIA（NVDA）收支與估值模型 v0.2（AI 半導體 Project：運算層第一個樣板）　2026-10-25", True),
 ("v0.1.1（PR 審查後）：(1) 新增 NVIDIA 認列顆數係數 k_rec（Derived，只在本模型內縮放，不改 markets/memory 的共用顆數；營收與成本同比例縮放），使 FY2027 由下而上的 Data Center 營收與毛利率等於錨點，FY2028 起沿用；v0.1 的未校準基準保留在 PL 對照區塊與 Sens。(2) 原 pf_cal（只調售價）留在 Sens 當對照。(3) Recon 新增 B2「差額來源待查」。(4) Val：10 年落後 P/E 規則值標「不適用」（保留計算，不進結論）。", False),
 ("模型命題：NVIDIA 的營收＝加速器顆數 × 每顆售價＋網通附掛。2027 年 ASIC 把 NVIDIA 的金額份額壓到約 75–80%，HBM 漲價壓毛利 2–3 個百分點。股價取決於 2028 年顆數持平之後，EPS 還能不能成長：現價約為 2027 年 EPS 的 13.5 倍，是低估，還是見頂前的折價。", False),
 ("核心問題：加速器顆數在 2028 年持平以後，NVIDIA 的 EPS 還能不能成長。（命題的「2027 年」＝日曆年 2027＝FY2028；「2028 年顆數持平」＝CY2028＝FY2029，故問題是 FY2030 EPS 相對 FY2029。）", False), ("", False),
 ("會計年度：FY(t)＝t−1 年 2 月～t 年 1 月，約等於日曆年 t−1（FY2027＝CY2026 … FY2030＝CY2029）。上游市場層是日曆年，FY(t) 直接取 CY(t−1)，不做跨年加權（11/12 落在同一日曆年）——這是近似。預測期 FY2027–FY2030。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("各世代顆數（markets/memory UNITS × 情境係數）× 每顆售價（橋接表 Rack 頁：GPU 模組＋CPU＋LPDDR＋匣級加價中 NVIDIA 保留的部分；Rubin Ultra 為 Assumed 區間）→ 運算營收；＋網通附掛（markets/network 的「NVIDIA 網通（本模型歸屬）」列，B／C 依顆數係數縮放）＋Edge Computing（公司實績＋成長率）", False),
 ("毛利率：非 HBM 成本占基準售價的比例（c_nh）在 FY2027 校準到公司錨點毛利率；FY2028 起毛利率的變化來自 HBM 與 LPDDR 成本（markets/memory 的情境別價格路徑 × 每顆容量）、NVIDIA 吸收比例（In 頁 abs_mem，Assumed 0–1，基準 0.5）、世代與顆數組合。中國資料中心運算基準為 0（公司 Q3 FY27 指引假設），Sens 列恢復。", False),
 ("EPS 口徑：GAAP 扣除權益證券損益（一次性），含股票薪酬；近似公司自 Q1 FY27 起的新 non-GAAP（含 SBC；Q 頁檢查列差 −0.03～−0.02）。營業費用、稅率、股數：公司實績（Q 頁）＋In 頁假設。", False),
 ("校準：FY2027 ＝ Q1、Q2 實績＋Q3 指引（$108B ±2%）＋Q4 假設。k_rec（Derived）＝錨點 Data Center 營收 ÷ 未校準的由下而上（市場層顆數 × 橋接表售價＋網通歸屬）；以已發生的實績校準，不用公司對 FY2028 的預測；顆數層與售價層無法分開識別，k_rec 是合併係數，FY2028 起沿用。c_nh 另讓 FY2027 毛利率等於錨點毛利率。驅動因子從 FY2028 起作用。公司指引（FY2028 營收約 +70%、毛利率 72–73%）與共識只在 Recon 對帳，不作校準。", False),
 ("估值：P/E 規則——FY2016–FY2025 年底 P/E 分位（見頂 25 分位、中段中位數、谷底 75 分位）；循環位置用 tools/cycle_position.py 的共用規則（lag 版：FY(t) 對照 CY t−1；見頂年度讀設備市場模型 " + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2028：" + "、".join(f"{s_} {CP.classify_lag(2028, PEAK[s_])}" for s_ in SC) + "）；另列以本模型營收路徑自己的見頂年判斷的對照欄（只作對照）。10 年落後 P/E 規則值標「不適用」（歷史倍數含高成長期溢價；保留計算，不進結論）；結論用現價 ÷ FY2028 EPS（模型與共識）。前瞻 P/E 歷史資料由擁有者另行提供，本版不做。", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未單獨使用。公司原文＝8-K 新聞稿（nvidianews.nvidia.com，已抓取）；CFO 書面說明、10-Q、法說逐字稿無法開啟，相關數字只經搜尋摘要（標二手），原文待人工複核。", False),
 ("已知限制：(1) k_rec 是顆數層與售價層的合併係數（無法分開識別，Recon B2 列各候選來源與公開資料，皆未驗證）；FY2028 起沿用 k_rec 假設該缺口持續；Sens 對照案例顯示 v0.1 未校準基準與只調售價（pf_cal）的結果，後者 FY2028 毛利率與公司指引不一致。(2) 吸收比例、Rubin Ultra 售價、ASIC 單價、Edge 與網通毛利率為 Assumed。(3) P/E 歷史只有一個可用序列（Macrotrends，9 年）；自算需要會計年度年底收盤價（未找到）。(4) 10 年落後 P/E 規則值不適用：歷史倍數含高成長期溢價，乘在 FY2028 EPS 上得到的價格遠高於現價（Val 頁 What must be true）。(5) FY↔CY 近似：Rubin 在 CY2026 的 1.7M 顆多數出貨在 FY2027 Q4 之後。(6) 網通的 B／C 縮放假設組合不變。(7) 公司 FY2028 美元指引（約 $690B）未找到直接陳述。", False)]
lines.append(("v0.1.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有儲存格）與 In 頁目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。主估值法標「不適用」，目標價改用十年 P/E 中位數 × 各情境 FY2027 EPS；其餘儲存格不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- 舊法：規則值情境加權（v0.1.2 的「0. 目標價與上檔空間」；v0.2 改名保留，不再叫目標價） ----------------
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
TP0 = add_target_price(V, {s: f"=pe_mid*PL!D{SUMR[(s, 'eps')]}" for s in SC},
    {s: f"情境 {s} FY2027 EPS（PL 頁）× 十年 P/E 中位數（第 6 節 pe_mid）" for s in SC}, "#,##0.00",
    "P/E：十年（FY2016–FY2025 會計年度年底，Macrotrends 9 年）落後 P/E 中位數 × 各情境 FY2027 EPS（不含權益證券損益；FY2027＝2026-02～2027-01）。"
    "本模型主估值法（FY2028 EPS × 循環位置分位）標「不適用」，依工作單改用此備案。十年倍數含高成長期溢價（FY2023 約 115x、FY2021 約 75x），偏誤方向為高估；"
    "FY2027 三情境 EPS 相同（情境差從 FY2028 起），故三情境隱含價相同；隱含價約為 FY2027 年底（2027-01）的價格、未折現。",
    "現價＝In 頁 px（2026-10-23 收盤）")

# ---------------- v0.2：三情境延伸到 FY2032、Val 第 0 節（DCF 與 Comparable 加權）、Sens 目標價參數 ----------------
XH, XR, XC = TP.extend_pl(PL, PB, FC, FYS, 2030, {"rev": "rev", "ni": "ni", "sh": "sh", "eps": "eps"}, "US$",
    "三情境延伸到 FY2032（v0.2；FY2026A–FY2030 引用上方推導區塊，FY2031 起為收斂法，非驅動因子推導；收斂年度為橘底）",
    "上游市場層（markets/*、台積電顆數）只到 CY2029＝FY2030；FY2031 起：營收年增由 FY2030 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋tp_gd、B＝tp_g、C＝tp_g−tp_gd），"
    "淨利率由 FY2030 線性收斂到中段利潤率（A＝FY2030、B＝FY2030 與十年平均的中點、C＝十年平均 nm10），股數沿用 FY2030 年變動", m10="nm10")
_h = next(r for r in range(1, V.max_row + 1) if V.cell(r, 1).value == "P/E：十年歷史倍數")
TPL, TPN = TP.add_tp_section(wb, V,
    eps={s: [f"PL!{XC[y]}{XR[(s, 'eps')]}" for y in range(2027, 2033)] for s in SC}, e_mon=1,
    fwd={s: [(f"PL!{XC[2028]}{XR[(s, 'eps')]}", 0.25), (f"PL!{XC[2029]}{XR[(s, 'eps')]}", 0.75)] for s in SC},
    fwd_note="會計年度 FY(t)＝t−1 年 2 月～t 年 1 月：約 12 個月後（2027-10 底）的未來 12 個月＝2027-11～2028-10＝FY2028 的 3 個月＋FY2029 的 9 個月（3/12 × FY2028＋9/12 × FY2029）",
    peers=[("AMD", "AMD", 46.49, "2026-10-23", "stockanalysis.com 統計頁（Forward PE）", "Interested-party（資料商彙整賣方預估）"),
           ("Broadcom", "AVGO", 18.52, "2026-10-23", "同上", "Interested-party（同上）"),
           ("Marvell", "MRVL", 40.44, "2026-10-23", "同上", "Interested-party（同上）"),
           ("台積電 ADR", "TSM", 19.61, "2026-10-23", "同上", "Interested-party（同上）"),
           ("Qualcomm", "QCOM", 18.98, "2026-10-23", "同上", "Interested-party（同上）")],
    peer_note="券商常用的 AI 運算／半導體大型股同業（GPU、客製 ASIC、代工、手機 SoC）；不含 A 股（Andy 11:53）。前瞻 P/E 口徑為資料商的未來 12 個月共識 EPS，各家 GAAP／non-GAAP 口徑不一",
    cons_tp="cons_tp", cons_note="stockanalysis.com：62 位分析師平均目標價（2026-10-25 頁面）；Interested-party",
    cons_fwd="cons_eps28", cons_fwd_note="共識 FY2028 EPS（In 頁 cons_eps28；FY2029 共識未找到，故 EPS 因素以 FY2028 同年度比較：B＝情境加權 FY2028 EPS ÷ 共識 − 1）",
    ours=lambda L: f"(tp_wA*PL!{XC[2028]}{XR[('A', 'eps')]}+tp_wB*PL!{XC[2028]}{XR[('B', 'eps')]}+tp_wC*PL!{XC[2028]}{XR[('C', 'eps')]})",
    eps28={s: f"PL!{XC[2028]}{XR[(s, 'eps')]}" for s in SC}, hist=f"B{_h + 2}:B{_h + 11}", fmt="#,##0.00",
    method="DCF：股東現金流＝EPS ×（1 − tp_g ÷ tp_roe），tp_r 10%、tp_g 4%、tp_roe 近三年平均 79%（再投資率約 5%）；基準日 2026 年底，FY2027 只計 2027-01 一個月，FY2032 後永續。"
           "Comparable：AMD、Broadcom、Marvell、台積電 ADR、Qualcomm 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤方向：C 情境 FY2031 起淨利率收斂到十年平均 31%（低於近年 55% 以上），偏保守；"
           "FY2028 三情境 EPS 幾乎相同，情境差主要來自 FY2029 起；永續價值占 DCF 價值過半，對 tp_r、tp_g 敏感（Sens 頁）。")
TP.add_tp_sens(SE, 26, TPL, "#,##0.00")
TP.dump_rowmap(TPL, TPN, XH, XR, XC, 26)
TP.yellow_inputs(IN)
RD.cell(len(lines) + 1, 1, "v0.2（2026-10-25）：(1) 目標價改為 DCF 與 Comparable 加權（Val 頁最前面「0. 目標價」：0.1 DCF 股東權益現金流、0.2 Comparable 同業前瞻 P/E、0.3 結論；Val 既有內容整體下移，公式一併改寫）；"
        "v0.1.2 的「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留在 Val 頁末。(2) PL 頁末新增三情境延伸到 FY2032：FY2031 起為收斂法，非驅動因子推導（上游市場層只到 CY2029＝FY2030）。"
        "(3) 現價更新到 2026-10-23 收盤 210.96（markets/invest_map/data/prices.csv）。(4) In 頁 M 節新增目標價參數（黃底，Andy 可改）；Sens 頁新增 tp_r、tp_g、tp_w_dcf 對加權目標價的影響。FY2026–FY2030 既有數字不變。")
RD.cell(len(lines) + 2, 1, "情境對應（v0.2 延伸期）：A＝順利（基準顆數）、B＝平平（CY2028 顆數 ×0.8、CY2029 ×0.85）、C＝不如預期（CY2028–29 ×0.6）。三情境在 FY2027 相同、FY2028 幾乎相同（只差 LPDDR 成本），DCF 的情境差異主要來自 FY2029 之後。本版未做先驗，無 Andy 合理 P/E 對照。")

order =["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("NV_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "VAL": {f"{k[0]}|{k[1]}": v for k, v in VAL.items()}, "SHR": {s: {str(k): v for k, v in R.items()} for s, R in SHR.items()}, "CHK": CHK_ROW, "KROW": K_ROW, "SENSN": SENS_N, "PB0": PB0, "RFMT": {str(k): v for k, v in RFMT.items()}, "PFROW": SB[PFCASE]["gm"], "RC_END": rr_}, open(os.environ["NV_ROWMAP"], "w"))
