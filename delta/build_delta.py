# 台達電（2308）個股模型 v0.2（AI 半導體 Project：電源與散熱次產業樣板）
# 用法：python3 delta/build_delta.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；市場與每 kW 內容以美元輸入、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（公司揭露的「稅後淨利」＝稅前 − 所得稅 − 非控制權益）。
import sys, os, glob
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import valuation_rule as VR
import fcff_dcf as FD   # v0.5：標準版 DCF（無槓桿自由現金流＋WACC）共用模組

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
    if k in ("racks", "capit", "cool", "rack"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 4, BR.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），市場與每 kW 內容為美元；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝台達電法說簡報與逐字稿（filecenter.deltaww.com）、月營收公告；公司的 AI 占比與毛利率說法為 Interested-party；券商估計皆為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
AM4Q25 = "https://filecenter.deltaww.com/ir/download/calendar/4Q25_Analyst Meeting.pdf"
TR2Q26 = "https://filecenter.deltaww.com/IR/download/calendar/2Q26_Transcript.pdf"
TN1012 = "https://finance.technews.tw/2026/10/12/delta-electronics-q3-revenue-nears-200-billion-record-high-q4-double-digit-growth"
KGI = "https://www.kgi.com.hk/en/-/media/files/kgishk/research-reports/tw-reports/power-supply-industry_01072025.pdf"

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A（原始顆數），B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；本模型以此除以每 GW IT 資本支出得到部署 GW" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合每 GW IT 資本支出用）", "%", *SIL.vals("vr26"), "Interested-party", "TrendForce 2026 組合 Rubin 22%（經 markets/memory）", "公司 2Q26 法說：2026 年 Blackwell 出貨遠大於 Rubin")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "Recon 用：每架內容換算")
inp("cool_vr", "VR200 每架機架內液冷（冷板等，橋接表）", "US$K", None, BR["cool"][1], None, "Interested-party", BR["cool"][2], "Morgan Stanley（經 Tom's Hardware 2025-11-05）：VR NVL144 機架內冷卻 BOM US$55.7K；不含 CDU、分歧管")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", *SIL.vals("fx_twd"), "Derived", "公司 2Q26 法說換算 NT$183.2B ≈ US$5.6B（隱含約 32.7）", "2026 年初 31.4（〈台幣〉開紅盤，uanalyze）；10-23 收盤：未找到")

sec("B. AI 電源內容與台達電份額（本專案為權威：每架電源與液冷內容、台達電份額，規則第 13 條）")
inp("pkw_26", *SIL.row9("pwr_pkw_26")[1:])   # Siliconomics M17
inp("pkw_27", *SIL.row9("pwr_pkw_27")[1:])   # Siliconomics M17
inp("pkw_28", *SIL.row9("pwr_pkw_28")[1:])   # Siliconomics M17
inp("pkw_29", *SIL.row9("pwr_pkw_29")[1:])   # Siliconomics M17
inp("rel_27", *SIL.row9("pwr_drel_27")[1:])   # Siliconomics M17
inp("rel_28", *SIL.row9("pwr_drel_28")[1:])   # Siliconomics M17
inp("rel_29", *SIL.row9("pwr_drel_29")[1:])   # Siliconomics M17
inp("lcr_27", "台達電液冷每 GW 營收 2027 ÷ 2026", "x", 0.9, 1.10, 1.3, "Assumed", "橋接表：機架內液冷 GB300 US$50K → VR200 US$56K／架（每 kW 下降：架數 6,764 → 4,052）", "公司：液冷占營收 2025 約 10% → 2026 超過 12%；Liquid-to-Air 使用期比市場預期長（2Q26 法說）")
inp("lcr_28", "台達電液冷每 GW 營收 2028 ÷ 2026", "x", 0.85, 1.15, 1.4, "Assumed", "同上", "—")
inp("lcr_29", "台達電液冷每 GW 營收 2029 ÷ 2026", "x", 0.8, 1.20, 1.5, "Assumed", "同上", "—")

sec("C. 營收結構錨點（公司揭露比例 × 營收）")
inp("ai_25", "2025 年 AI 相關營收 ÷ 營收（讀為不含液冷）", "%", None, 0.18, None, "Interested-party", TN1012 + "：AI 相關營收 2025 年 18%、2026 年超過 25%（公司）", "數位時代 2026-02-26：AI 電源超過 20%、液冷約 9%、合計約 30%（口徑不同，未採用）")
inp("ai_26", "2026 年 AI 相關營收 ÷ 營收（讀為不含液冷）", "%", 0.25, 0.26, 0.28, "Interested-party", TR2Q26 + "：CFO「一定超過 25%」；BigGo 摘要：朝 28% 目標", "讀法（Assumed）：AI 相關＝AI 電源，不含液冷；另一讀法見 Sens「AI 相關含液冷」")
inp("lc_25", "2025 年液冷營收 ÷ 營收", "%", None, 0.10, None, "Interested-party", TR2Q26 + "：液冷約占 2025 年營收 10%", "數位時代：約 9%")
inp("lc_26", "2026 年液冷營收 ÷ 營收", "%", 0.12, 0.125, 0.13, "Interested-party", TR2Q26 + "：2026 年超過 12%", TN1012 + "：液冷產品可望超過 12%")
sec("C2. 營收分部（官方四大事業群：電源及零組件、交通、自動化、基礎設施；v0.3 依個股模型原則重建）")
AR25 = "台達電 2025 年報（英文版 p.136，2026-03-30）：電源及零組件 50%、交通 7%、自動化 10%、基礎設施 33%（經 BESSmanufacturers 摘錄）"
FT926 = "FTNN 2026-10-12（台達電 9 月營收公告）：電源及零組件 51%、基礎設施 35%、自動化 8%、交通 6%"
for k, nm, v25, v26, lo26, hi26 in (("pe", "電源及零組件", 0.50, 0.51, 0.50, 0.53), ("mo", "交通", 0.07, 0.06, 0.05, 0.07), ("au", "自動化", 0.10, 0.08, 0.07, 0.09)):
    inp(f"s25_{k}", f"2025 年{nm} ÷ 營收", "%", None, v25, None, "Verified", AR25, "四捨五入值；基礎設施＝1 − 其餘")
    inp(f"s26_{k}", f"2026 年{nm} ÷ 營收", "%", lo26, v26, hi26, "Interested-party", FT926 + "（9 月或前三季比重，文中未區分；全年 Assumed 同此）",
        "理財周刊 2026-08-05：2Q26 電源及零組件 962.96 億（約 52.6%）、基礎設施 629.38 億（約 34.3%）" if k == "pe" else "—")
inp("s25_inf", "2025 年基礎設施 ÷ 營收", "%", None, "=1-s25_pe-s25_mo-s25_au", None, "Derived", "1 − 其餘三事業群（公司 33%）", "—")
inp("s26_inf", "2026 年基礎設施 ÷ 營收", "%", None, "=1-s26_pe-s26_mo-s26_au", None, "Derived", "1 − 其餘三事業群（公司 35%）", "2Q26 基礎設施年增約 78%、營業利益年增約 121%（理財周刊 2026-08-05）")
for k, nm, vals, s1, s2 in (
        ("pe", "電源及零組件：非 AI（一般伺服器電源、消費與工業電源、零組件）", ((0.0, 0.05, 0.10), (0.0, 0.04, 0.08), (0.0, 0.03, 0.06)), "2Q26 電源及零組件年增約 50%（含 AI 電源；BigGo 法說摘要）", "AI 電源（命題子列）另由部署 GW × 每 kW × 份額推導"),
        ("inf", "基礎設施：非液冷（資料中心 UPS 與電力、通訊電源、樓宇自動化、儲能）", ((0.03, 0.09, 0.15), (0.0, 0.07, 0.12), (0.0, 0.06, 0.10)), "2Q26 基礎設施年增約 78%（含液冷；理財周刊 2026-08-05）；資料中心與液冷需求、擴廠（BigGo）", "液冷（命題子列）另由部署 GW × 每 GW 液冷營收推導"),
        ("au", "自動化", ((0.0, 0.05, 0.10), (0.0, 0.05, 0.10), (0.0, 0.04, 0.08)), "2Q26 自動化年增約 15% 但虧損；記憶體供給改善後下半年轉盈（BigGo 法說摘要）", "2025 自動化 +5%（4Q25 法說簡報）"),
        ("mo", "交通（電動車）", ((-0.10, 0.0, 0.10), (-0.05, 0.02, 0.10), (-0.05, 0.02, 0.10)), "交通 2026 年衰退、持續虧損，全年表現不佳；長期趨勢不變（BigGo 法說摘要）", "2025 交通 −16%（4Q25 法說簡報）")):
    for y, v in zip((27, 28, 29), vals):
        inp(f"g_{k}{y}", f"{nm}年增 20{y}", "%", *v, "Assumed", s1 if y == 27 else "同上", s2 if y == 27 else "—")

sec("D. 毛利率")
inp("d_ai", "AI 電源毛利率高於其他營收的幅度", "pt", 0.04, 0.08, 0.12, "Assumed", "券商：18kW PSU 與 HVDC 電源櫃毛利率約 50%（旺得富 2026-10-01，Interested-party）；公司整體 1H26 36.3%", "部門別毛利率未揭露：未找到；電源及零組件 2025 部門利益率 20.1% vs 基礎設施 15.6%（4Q25 法說簡報）")

sec("E. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收合計）", "NT$B", None, 197.059, None, "Verified", TN1012 + "：3Q26 1,970.59 億元、前三季 5,396.68 億元", "時報資訊 2026-10-12 同值（含更正稿）")
inp("q4g", "4Q26 營收季增", "%", 0.03, 0.10, 0.20, "Interested-party", "3Q26 法說（2026-10-29）：第四季營收「一定比第三季高」；法人：第四季雙位數季增（TechNews 2026-10-12）", "券商 2026-10-01：4Q26 2,511 億（季增 20%，以 3Q 2,093 億為基礎；3Q 實際 1,971 億）")
inp("q_gm", "4Q26 毛利率（3Q26 已為實績 34.1%）", "%", 0.341, 0.345, 0.36, "Interested-party", "3Q26 法說（2026-10-29，股感轉述）：第四季毛利率不保證，但應不會比第三季 34.1% 差；原 2Q26 法說：董事長「35% 左右合理」", "券商 2026-10-01：4Q26 38.7%（新高）")
inp("q_opx", "4Q26 營業費用 ÷ 營收", "%", 0.165, 0.175, 0.185, "Assumed", "3Q26 17.5%（公司）；1H26 19.1%、4Q25 18.3%（公司）", "研發 +33%、銷管 +40%（2Q26 年增）")
inp("q_nop", "4Q26 業外收益", "NT$B", 0.0, 1.1, 4.4, "Derived", "2Q26 4.4（部分為未實現投資利益）、4Q25 −0.8（公司）", "FY2025 3.9（公司）")
inp("q_nir", "4Q26 歸屬母公司淨利 ÷ 稅前（3Q26 實績 0.697）", "%", 0.67, 0.69, 0.71, "Derived", "1H26 45.7 ÷ 66.1＝0.691；FY2025 60.1 ÷ 87.9＝0.684（公司）", "含稅率約 22% 與非控制權益（台達泰國等）")
inp("q_sh", "股數（十億股；EPS 換算）", "十億股", None, "=45.7/17.59", None, "Derived", "1H26 淨利 45.7 ÷ EPS 17.59（公司）", "已發行 2,597,543,329 股（HiStock）")

sec("F. 營業費用、業外、稅與非控制權益、股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.6, 0.8, 1.0, "Derived", "FY2026 錨點實測（Q 頁 B26）", "FY2025：營業費用 +20% ÷ 營收 +32%＝0.62（4Q25 法說簡報）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("nop_f", "業外收益（每年）", "NT$B", 4.0, 8.0, 12.0, "Assumed", "FY2025 3.9、1H26 7.2（公司）", "—")
inp("nir_f", "歸屬母公司淨利 ÷ 稅前", "%", 0.66, 0.69, 0.71, "Derived", "同 q_nir", "—")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.005, "Assumed", "FY2025 推導 2.598 → 1H26 2.598（公司 EPS 與淨利）", "—")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("G. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("b_rev26", "券商 FY2026 營收（2026-10-01，3Q 前）", "NT$B", None, 803.0, None, "Interested-party", "旺得富 2026-10-01（本土投顧，未具名；由 826.4 下修）", "—")
inp("b_eps26a", "券商 FY2026 EPS（本土，2026-10-01）", "NT$", None, 43.1, None, "Interested-party", "旺得富 2026-10-01", "—")
inp("b_eps26b", "券商 FY2026 EPS（美系，2026-05-04）", "NT$", None, 41.5, None, "Interested-party", "CMoney 2026-05-04（美系大行，未具名）", "—")
inp("b_eps27a", "券商 FY2027 EPS（本土，2026-10-01）", "NT$", None, 80.8, None, "Interested-party", "旺得富 2026-10-01；2028 EPS 121.6", "—")
inp("b_eps27b", "券商 FY2027 EPS（美系，2026-05-04）", "NT$", None, 65.1, None, "Interested-party", "CMoney 2026-05-04；2028 EPS 98.5", "—")
inp("b_eps27c", "券商 FY2027 EPS（投顧研究團隊，2026-10-12）", "NT$", None, 62.68, None, "Interested-party", "Yahoo 新聞 2026-10-12（台達電 Q3 營收與 337 調查）：2027 營收 1.05 兆、EPS 62.68、38x", "—")
inp("b_rev27", "券商 FY2027 營收", "NT$B", None, 1050.0, None, "Interested-party", "同上", "鉅亨：「2027 上看兆元」")
inp("ubs_26", *SIL.row9("pwr_ubs_26")[1:])   # Siliconomics M17
inp("ubs_27", *SIL.row9("pwr_ubs_27")[1:])   # Siliconomics M17
inp("kgi_sh", *SIL.row9("pwr_kgi_0701")[1:])   # Siliconomics M17（CF11 待 Andy：與 pwr_kgi_0926 並列，只作 Recon）
inp("pe_fair_andy", "Andy 先驗：合理 P/E", "x", None, 25, None, "Assumed", "先驗_Andy_台達電_20261024 D10b（Drive；不在本 repo）", "只作對照")

sec("H. 價格與估值")
inp("px", "股價（2026-10-30 收盤）", "NT$", None, 1335, None, "Interested-party", "Taiwan News、cnyes（markets/invest_map/data/prices.csv）：1,335（10-30 跌停）", "v0.3.1 前標 10-23 的 1,605 實為 10-24 頁面昨收；10-23、10-26 台股未交易，10-22 收盤 1,560")
inp("pe_ttm", "近 12 月 P/E（10-22）", "x", None, 49.65, None, "Interested-party", "Goodinfo（10-22）", "HiStock 49.65")

sec("I. 目標價情境權重（Val 第 0 節用；Andy 可改）")
for s_, w_, nm_ in (("A", 0.3, "順利"), ("B", 0.5, "平平"), ("C", 0.2, "不如預期")):
    inp(f"tp_w{s_}", f"目標價情境權重：情境 {s_}（{nm_}）", "權重", None, w_, None, "Assumed",
        "Andy 可改（2026-10-25 目標價規格預設 A／B／C＝0.3／0.5／0.2）", "三個權重合計應為 1；v0.2 起用於 DCF 與 Comparable 的情境加權（Val 第 0 節）")

# ---------------- v0.2 設定與新增輸入（附加在 In 既有輸入之後，不位移既有列） ----------------
V2 = {'name': '台達電',
 'unit': 'NT$B',
 'shunit': '十億股',
 'cur': 'NT$',
 'shg': 'sh_g',
 'epsfmt': '0.00',
 'pxfmt': '#,##0',
 'keys': {'rev': 'rev', 'ni': 'ni', 'sh': 'sh_', 'eps': 'eps'},
 'roe': ('=AVERAGE(0.1661,0.1547,0.2246)', 'stockanalysis 比率頁（Return on Equity）：FY2023 16.61%、FY2024 15.47%、FY2025 22.46%', '單一來源；近 12 月 29.97%（同站）'),
 'cons_tp': (2502.36, 'stockanalysis 個股與 forecast 頁（S&P Global）22 位分析師平均 2,502（中位數 2,500），頁面時間 2026-10-22', '第二來源：未找到'),
 'cons_feps': ('=1560/30.51', 'stockanalysis 個股頁 2026-10-22：價格 1,560 ÷ Forward PE 30.51'),
 'peers': [('Vertiv', 'VRT', 24.53, '2026-10-23'), ('Eaton', 'ETN', 24.17, '2026-10-23'), ('Schneider Electric', 'SU.PA', 19.13, '2026-10-23'), ('光寶科（Lite-On）', '2301.TW', 19.7, '2026-10-22')],
 'peer_note': 'AI 資料中心電源與電氣同業：Vertiv、Eaton、Schneider Electric、光寶科（券商常用比較組）。麥格米特（002851.SZ）為 A 股，依規則不列。非 A 股公司的同業名單不放中國 A 股（A 股有結構性溢價；Andy 2026-10-25 11:53）',
 'fwd_w': (0.16666666666666666, 0.8333333333333334),
 'fwd_note': '約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）',
 'pe10': None,
 'px_note': '現價＝In 頁 px（2026-10-30 收盤 1,335）',
 'method': 'DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（股權報酬視為費用），以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 台灣 10 年期公債殖利率)、RONIC＝WACC；ke＝rf＋β ×（ERP 5%＋台灣國家風險溢酬），β 為自算產業 β（TW 散熱／電源同業中位數，Blume 調整）依台達電 D／E 重新加槓桿；EV 加淨現金與長期投資、扣非控制權益 ÷ 股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁；tools/fcff_dcf.py）。FY2027–29 驅動因子、FY2030–32 收斂法。Comparable：Vertiv、Eaton、Schneider、光寶科前瞻 P/E '
           '中位數；歐美電氣大廠業務組合較成熟、成長較低，台達電歷來倍數較高，偏低估。情境對應：A＝順利、B＝平平、C＝不如預期（原顆數情境）。',
 'readme_val': '估值：Val 第 0 節目標價＝0.5 × DCF 目標價（v0.5 起為標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，FY2027–FY2032 明確期＋終值，以 WACC 折現；ke＝rf＋β ×（ERP＋國家風險溢酬）、g＝MIN(3%, rf)、RONIC＝WACC，三情境權重 0.3／0.5／0.2，以 ke 推到 12 個月後並扣期間股利；明細見 DCF 頁）＋0.5 × Comparable 目標價（Vertiv、Eaton、Schneider Electric、光寶科 前瞻 '
               'P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解（EPS 因素、倍數因素）、DCF 價格帶、十年 P/E 分位檢查（只作檢查）。',
 'readme_add': ['v0.2（2026-10-25）：(1) PL 頁最下方新增延伸表，三情境延伸到 FY2032；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt；淨利率由 FY2029 收斂到 A＝FY2029、B＝FY2029 '
                '與十年平均的中點、C＝十年平均；十年平均淨利率未找到，依工作單以 FY2029 × 0.8 替代，Assumed；PL 延伸表灰底）。(2) Val 第 0 節目標價＝DCF 與 Comparable 加權（tp_w_dcf＝0.5）；v0.1.1 的「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留，不再是目標價。(3) Sens '
                '新增目標價參數（tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.4／0.8）。(4) Recon 新增目標價對帳。(5) 現價：2026-10-23 收盤 1,605 已是最新（markets/invest_map/data/prices.csv 同值，單一來源），未變動。 新增內容一律放在各頁既有內容之後，FY2025–FY2029 '
                '既有數值不變。',
                '情境對應：A＝順利（原始顆數）、B＝平平（CY2028 顆數 ×0.8、CY2029 ×0.85）、C＝不如預期（CY2028–29 ×0.6），即原本的三個顆數情境。三情境在 FY2028 前完全相同（只差顆數係數，從 CY2028 起），DCF 三情境仍照算，差異來自 FY2028 之後（含 FY2030 起的收斂法）。',
                '先驗：v0.1 的 Andy 合理 P/E 25x 對照（Val 第 5 節）保留不動，不進目標價；本版未新增先驗。']}
sec("J. 目標價（DCF 與 Comparable 加權）與收斂法參數（v0.2 新增；黃底＝Assumed，Andy 可改）")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.5 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.5 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑（fade_gA／B／C）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 改為 0.5：DCF／Comparable 各半）", "Sens 頁 0.4／0.8")
inp("fade_gA", "情境 A（順利）：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂，FY2032 達到")
inp("fade_gB", "情境 B（平平）：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C（不如預期）：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("nim10_f", "十年平均淨利率的替代：最後一年（FY2029）淨利率 ×（本值）", "x", None, 0.8, None, "Assumed", "十年（2016–2025）平均淨利率：未找到（免費資料商只到近五年）；依工作單以最後一年 × 0.8 替代", "淨利率收斂目標：A＝FY2029、B＝FY2029 與替代值的中點、C＝替代值（PL 延伸表）")
inp("cons_tp", "市場共識目標價（賣方平均）", V2["cur"], None, V2["cons_tp"][0], None, "Interested-party", V2["cons_tp"][1], V2["cons_tp"][2])
inp("cons_feps", "共識前瞻 12 月 EPS（＝價格 ÷ 前瞻 P/E）", V2["cur"], None, V2["cons_feps"][0], None, "Interested-party", V2["cons_feps"][1], "賣方共識，只作對帳與差距拆解，不作校準")
# v0.5：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。評價日沿用 v0.2 DCF 基準日 2026-12-31；股價 2026-10-30；資產負債為 2026-09-30
BETA_SRC = "自算產業 β 表（2026-11-05 22:50 計算；Yahoo Finance 含股利調整月收盤價 2021-11～2026-10，60 個月；Andy 2026-11-05 22:55 選定，取代 Damodaran 表）"
DAMO_C = "Damodaran「Country Default Spreads and Risk Premiums」2026-01-05 版（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）"
SA_BS = "stockanalysis.com 台達電（TPE:2308）資產負債表，2026-09-30 欄（https://stockanalysis.com/quote/tpe/2308/financials/balance-sheet/；2026-11-05 讀取；資料商轉載合併報表）"
SA_CF = "stockanalysis.com 台達電現金流量表（https://stockanalysis.com/quote/tpe/2308/financials/cash-flow-statement/；2026-11-05 讀取）"
inp("beta_u", "產業無槓桿 β：自算 TW 散熱／電源同業中位數，Blume 調整（0.67β＋0.33）", "x", "=beta_u-beta_du", "=0.67*1.30+0.33", "=beta_u+beta_du", "Derived",
    BETA_SRC + "：TW 散熱／電源同業 6 家（3017 3324 3653 2308 2301 6223），對台灣加權指數 ^TWII 迴歸、Hamada 去槓桿、現金修正後中位數 1.30；Blume 調整 0.67 × 1.30＋0.33＝1.201",
    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3）")
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.2 DCF 基準日）", ccy="NT$", amt="NT$B", shu="十億股",
        rf_us=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
               "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）；與 NVIDIA 模型同值。只用於 Sens「美元口徑折現」（Andy 2026-11-05 23:16：基準維持當地利率＋固定匯率）"),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=SA_BS + "；富邦證券 e-Broker 資產負債表（季）2026.3Q 欄同值（2026-11-05 讀取；網址見 Sources 頁）",
                  div=(0, "Verified", "FY2025 現金股利每股 11.6 元已於 2026-07-22 發放（除息 2026-06-17；玩股網除權息表，2026-11-05 讀取；網址見 Sources 頁），早於資產負債表日 2026-09-30", "台達電一年配息一次；2026-10-01～12-31 無已宣告股利或買回（未找到），不扣")),
        rf=(0.0198, "Trading Economics「Taiwan 10Y Bond Yield」：2026-11-04 1.98%（https://tradingeconomics.com/taiwan/government-bond-yield；2026-11-05 讀取）",
            "經濟日報 2026-08-21（https://money.udn.com/money/story/5613/9704659）：Bloomberg 經濟學家預期下半年接近 2%；股價日 10-30 前後一個月變動約 0.03pt", "無風險利率（台灣 10 年期公債殖利率，2026-11-04）"),
        beta_fixed=("=beta_u_lo*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "=beta_u*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "=beta_u_hi*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "Derived",
                    "自算產業 β：TW 散熱／電源 同業 6 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對台灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：β＝beta_u ×（1＋(1 − tax_w) × D ÷ E），D＝債務＋租賃、E＝現價 × 股數）",
                    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。Sens 低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）"),
        crp=(0.0078, "Verified", DAMO_C + "：Taiwan（Moody's Aa3）Country Risk Premium 0.78%", "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋國家風險溢酬、不設股權成本下限"),
        kd=("=rf+0.01", "Assumed", "台灣 10 年期公債＋1 個百分點（公司借款利率未揭露：未找到）", "負債占資本約 3%，對 WACC 影響可忽略"),
        tax_w=("=(10.924/51.316+19.93/87.866)/2", "Derived", "近年有效稅率平均：FY2024 所得稅 109.24 億 ÷ 稅前 513.16 億＝21.3%、FY2025 199.30 億 ÷ 878.66 億＝22.7%（台達電 4Q25 法說簡報，filecenter.deltaww.com）",
               "FY2023：未找到（只用兩年）；NOPAT 各年亦用本稅率（本模型損益只有「淨利 ÷ 稅前」合併稅與非控制權益，無單獨稅率列）"),
        debt=("=12.779+26.462+71.385", "Verified", SA_BS + "：短期借款 12,779＋一年內到期長期負債 26,462＋長期負債 71,385（NT$M）", "2025-12-31：2,364＋14,961＋56,703（同站）"),
        lease=(4.454, "Verified", SA_BS + "：租賃負債（流動＋非流動）4,454（NT$M）", "2025-12-31：3,288（同站）"),
        cash=(213.212, "Verified", SA_BS + "：現金及短期投資 213,212（NT$M）", "2025-12-31：154,408（同站）"),
        inv=(8.252, "Verified", SA_BS + "：長期投資 8,252（NT$M；權益法與金融資產）", "帳面計入評價；Sens 測不計"),
        mi=(66.233, "Verified", SA_BS + "：非控制權益 66,233（NT$M；台達泰國等）", "FCFF 為合併口徑（含非控制權益），故評價時扣除非控制權益帳面值"),
        dps=("=a26_eps*11.6/23.14", "Derived", "評價日 2026-12-31 到目標時點 2027-12-31 之間發放 FY2026 盈餘的現金股利：假設發放率同 FY2025（每股 11.6 元 ÷ EPS 23.14 元＝50.1%）× FY2026 錨點 EPS",
             "uanalyze 2026-02-25：FY2025 EPS 23.14 元、擬配息 11.6 元（https://uanalyze.com.tw/articles/6260448112）"),
        sh=("=q_sh", "Derived", "＝In 頁 q_sh（1H26 淨利 ÷ EPS，約 25.98 億股）", "已發行 2,597,543,329 股（HiStock）；不另推未來股數"),
        fc_da=(None, "=(21.59/401.227+25.016/421.148+27.839/554.885)/3", None, "Derived", SA_CF + "：折舊攤銷 FY2023 21,590、FY2024 25,016、FY2025 27,839（NT$M）÷ 營收 401,227／421,148／554,885，近三年平均 5.45%",
               "FY2025 值與 In 頁 da_r 相同（5.02%）；近 12 月（2026-06）30,482 ÷ 654,539＝4.7%"),
        fc_cx=(None, "=(27.83/401.227+33.485/421.148+46.091/554.885)/3", None, "Derived", SA_CF + "：資本支出 FY2023 27,830、FY2024 33,485、FY2025 46,091（NT$M）÷ 營收，近三年平均 7.73%",
               "近 12 月（2026-06）54,170 ÷ 654,539＝8.3%；Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=fc_da_ss*(1+5*g_tv)", "Assumed", "長期穩態（FY2032）資本支出 ÷ 營收＝折舊 ÷ 營收 ×（1＋5 × g）：設備平均年限約 10 年、以 g 成長時，資本支出 ÷ 折舊約 1＋g × 年限 ÷ 2", "滿足規格「資本支出 ≥ 折舊 ×（1＋g）」；FY2029 起由近期比例線性收斂"),
        fc_wc=(0.20, "=((84.909+76.227-53.539)/401.227+(95.77+83.856-69.223)/421.148+(128.446+101.478-94.452)/554.885)/3", 0.30, "Derived",
               SA_BS.replace("，2026-09-30 欄", "") + "：營運資金＝應收帳款＋存貨 − 應付帳款；FY2023 107,597、FY2024 110,403、FY2025 135,472（NT$M）÷ 營收＝26.8%、26.2%、24.4%，平均 25.8%",
               "2026-09-30 194,135 ÷ 近 12 月營收約 6,545 億＝29.7%（存貨上升）；Sens 低／高 20%／30%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代第一輪的「中位數 × 0.8 再 Blume／中位數未調整」")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑；Sens β 案例用")

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
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自台達電法說簡報、逐字稿與月營收"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（月營收）＋3Q、4Q 假設（毛利率、營業費用率、業外、淨利 ÷ 稅前；4Q 營收＝3Q × (1＋季增)）。1H26 營業利益＝2Q26 30.6＋1Q26（2Q26 ÷ 1.08；公司稱季增 8%）。部門別損益只有年度簡報數字，季度部門數字只在簡報畫面：未找到"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：" + AM4Q25 + "；1H26：" + TR2Q26
S2 = "3Q26 營收：月營收公告（TechNews 2026-10-12）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 554.885, "C": 342.6, "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Interested-party（4Q 季增）")
qrow(5, "毛利", {"B": 190.2, "C": 124.3, "D": "=q3_gp", "E": "=E4*q_gm"}, tag="Verified／Interested-party")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 34.3%；1H26 36.3%")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D5-D8", "E": "=E4*q_opx"}, tag="Derived（毛利 − 營業利益）／Assumed")
qrow(8, "營業利益", {"B": 83.9, "C": "=30.6+30.6/1.08", "D": "=q3_op", "E": "=E5-E7"}, tag="Verified／Derived")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 15.1%；1H26 17.2%（公司）")
qrow(10, "業外收益", {"B": 3.934, "C": "=C11-C8", "D": "=D11-D8", "E": "=q_nop"}, tag="Verified／Derived")
qrow(11, "稅前利益", {"B": "=B8+B10", "C": 66.1, "D": "=q3_pti", "E": "=E8+E10"}, tag="Verified／Derived", s2="FY2025 公司揭露 87.9（四捨五入差）")
qrow(12, "歸屬母公司淨利", {"B": 60.108, "C": 45.7, "D": "=q3_ni", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）")
qrow(13, "淨利 ÷ 稅前（含稅與非控制權益）", {c: f"={c}12/{c}11" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股）", {"B": "=60.108/23.14", "C": "=q_sh", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s2="淨利 ÷ EPS")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(16, "  公司揭露 EPS", {"B": 23.14, "C": 17.59, "D": "=q3_eps"}, fmt="0.00", ann="", tag="Verified", s2="2Q26 9.68")
qrow(17, "AI 電源營收（＝營收 × AI 相關占比）", {"B": "=B4*ai_25"}, ann="=F4*ai_26", tag="Derived／Interested-party", s1="公司 AI 相關占比（In 頁 C 節）", s2="各季 AI 營收未揭露：未找到")
qrow(18, "液冷營收（＝營收 × 液冷占比）", {"B": "=B4*lc_25"}, ann="=F4*lc_26", tag="Derived／Interested-party", s1="公司液冷占比", s2="—")
qrow(19, "非 AI 營收（＝營收 − AI 電源 − 液冷）", {"B": "=B4-B17-B18"}, ann="=F4-F17-F18", tag="Derived", s1="—", s2="含非 AI 資料中心、電源零組件、自動化、交通")
qrow(20, "AI 電源（另一讀法：AI 相關含液冷 → AI 電源＝營收 ×（AI 相關 − 液冷））", {"B": "=B4*(ai_25-lc_25)"}, ann="=F4*(ai_26-lc_26)", tag="Derived", s1="Sens 用", s2="—")
qrow(21, "非 AI（另一讀法）", {"B": "=B4-B20-B18"}, ann="=F4-F20-F18", tag="Derived", s1="Sens 用", s2="—")
Q.cell(23, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(24, 1, "FY2026 營收年增"); Q["B24"] = "=F4/B4-1"
Q.cell(25, 1, "FY2026 營業費用年增"); Q["B25"] = "=F7/B7-1"
Q.cell(26, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B26"] = "=B25/B24"
Q.cell(27, 1, "FY2026 營收 ÷ 券商（10-01，3Q 前）− 1"); Q["B27"] = "=F4/b_rev26-1"
Q.cell(28, 1, "FY2026 EPS ÷ 券商（本土 10-01）− 1"); Q["B28"] = "=F15/b_eps26a-1"
Q.cell(29, 1, "非 AI 營收 FY2026 年增"); Q["B29"] = "=F19/B19-1"
for rr in range(24, 30): Q[f"B{rr}"].number_format = "0.0%" if rr != 26 else "0.00"
Q.cell(31, 1, "營收分部（官方四大事業群；v0.3）").font = BOLD
for rr, k, nm in ((32, "pe", "電源及零組件"), (33, "mo", "交通"), (34, "au", "自動化"), (35, "inf", "基礎設施")):
    qrow(rr, nm, {"B": f"=B4*s25_{k}"}, ann=f"=F4*s26_{k}", tag="Verified（FY2025）／Interested-party（FY2026）", s1="營收 × 事業群比重（In 頁 C2 節）", s2="1H26、3Q26 各事業群金額：未全取得（2Q26 電源及零組件、基礎設施見 In 頁）")
qrow(36, "檢查：分部合計 − 合併營收", {"B": "=B32+B33+B34+B35-B4"}, ann="=F32+F33+F34+F35-F4", tag="Derived", s1="應為 0", s2="—")
for nm, rr in (("pe", 32), ("mo", 33), ("au", 34), ("inf", 35)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nop", 10), ("pt", 11), ("ni", 12), ("sh", 14), ("eps", 15), ("aip", 17), ("lc", 18), ("nai", 19), ("aipx", 20), ("naix", 21)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 26; Q.column_dimensions["H"].width = 90; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
_mfy = lambda m, y: m[y] if isinstance(m, dict) else m   # Siliconomics v0.13：mf 可為年別 dict（Sens「Wattonomics 錨定」案例）；預設字串不變
DEFP = dict(scn="A", rel="rel", pkw="pkw", lcr="lcr", mf="1", gsuf="", d="d_ai", opxa="0", aip="aip", nai="nai")
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
    sec_("A. AI 部署與電源市場（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{_mfy(P['mf'], y)}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in BU[1:]}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    row("pkw", "AI 電源內容", "US$/kW", {c_(y): f"={P['pkw']}_{y % 100}" for y in BU}, fmt="#,##0")
    row("pmk", "AI 電源市場（美元）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['pkw']}/1000" for y in BU})
    row("pmn", "AI 電源市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['pmk']}*fx" for y in BU}, fmt="#,##0.0")
    row("sh0", "台達電有效份額（FY2026 校準）＝ AI 電源營收 ÷ 市場", "%", {"D": f"=a26_{P['aip']}/D{R['pmn']}"}, fmt="0.0%",
        note="Derived：分子為公司 AI 相關占比 × 營收；分母為本模型的 AI 電源市場（含 ASIC 機架）")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "台達電有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("aip", "AI 電源營收", "NT$B", {"C": f"=a25_{P['aip']}", **{c_(y): f"={c_(y)}{R['pmn']}*{c_(y)}{R['sh']}" for y in BU}}, fmt="#,##0.0")
    sec_("B. 液冷、非 AI 與合計")
    row("lpg0", "液冷每 GW 營收（FY2026 校準）", "NT$B/GW", {"D": f"=a26_lc/D{R['gw']}"}, fmt="0.000")
    row("lrel", "液冷每 GW 相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['lcr']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("lc", "液冷營收", "NT$B", {"C": "=a25_lc", **{c_(y): f"={c_(y)}{R['gw']}*$D${R['lpg0']}*{c_(y)}{R['lrel']}" for y in BU}}, fmt="#,##0.0")
    gs = lambda k, y: f"g_{k}{P['gsuf']}{y % 100}"
    row("pen", "電源及零組件：非 AI", "NT$B", {"C": f"=a25_pe-a25_{P['aip']}", "D": f"=a26_pe-a26_{P['aip']}", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('pe', y)})" for y in BU[1:]}}, fmt="#,##0.0",
        note="AI 電源（上列）為本分部的命題子列；FY2025、FY2026＝分部 − AI 電源")
    row("pe", "電源及零組件", "NT$B", {c_(y): f"={c_(y)}{R['aip']}+{c_(y)}{R['pen']}" for y in FYS}, fmt="#,##0.0", note="＝AI 電源＋非 AI")
    row("infn", "基礎設施：非液冷", "NT$B", {"C": "=a25_inf-a25_lc", "D": "=a26_inf-a26_lc", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('inf', y)})" for y in BU[1:]}}, fmt="#,##0.0",
        note="液冷（上列）為本分部的命題子列（Assumed：液冷全部歸基礎設施）")
    row("inf", "基礎設施", "NT$B", {c_(y): f"={c_(y)}{R['lc']}+{c_(y)}{R['infn']}" for y in FYS}, fmt="#,##0.0", note="＝液冷＋非液冷")
    row("au", "自動化", "NT$B", {"C": "=a25_au", "D": "=a26_au", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('au', y)})" for y in BU[1:]}}, fmt="#,##0.0")
    row("mo", "交通", "NT$B", {"C": "=a25_mo", "D": "=a26_mo", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('mo', y)})" for y in BU[1:]}}, fmt="#,##0.0")
    row("nai", "非 AI 營收（電源非 AI＋基礎設施非液冷＋自動化＋交通）", "NT$B", {c_(y): f"={c_(y)}{R['pen']}+{c_(y)}{R['infn']}+{c_(y)}{R['au']}+{c_(y)}{R['mo']}" for y in FYS}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['pe']}+{c_(y)}{R['inf']}+{c_(y)}{R['au']}+{c_(y)}{R['mo']}" for y in FYS}, fmt="#,##0.0", note="＝四大事業群合計")
    row("chk", "檢查：分部合計 − 合併營收（FY2025 實績、FY2026 錨點）", "NT$B", {"C": f"=C{R['rev']}-a25_rev", "D": f"=D{R['rev']}-a26_rev"}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 電源 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['aip']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("aall", "AI 電源＋液冷 ÷ 營收", "%", {c_(y): f"=({c_(y)}{R['aip']}+{c_(y)}{R['lc']})/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pgw", "每 GW 台達電 AI 營收（電源＋液冷）", "US$M/GW", {c_(y): f"=({c_(y)}{R['aip']}+{c_(y)}{R['lc']})/fx/{c_(y)}{R['gw']}*1000" for y in BU}, fmt="#,##0")
    sec_("C. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × AI 電源占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['aish']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 電源毛利率＝其他 + 差距。FY2027 起毛利率隨 AI 電源占比變動")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"=$D${R['gmn']}*{c_(y)}{R['rev']}+{P['d']}*{c_(y)}{R['aip']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,phi*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nop", "業外收益", "NT$B", {"C": "=a25_nop", "D": "=a26_nop", **{c_(y): "=nop_f" for y in BU[1:]}}, fmt="#,##0.0")
    row("nir", "歸屬母公司淨利 ÷ 稅前", "%", {"C": "=a25_ni/(a25_oi+a25_nop)", "D": "=a26_ni/(a26_oi+a26_nop)", **{c_(y): "=nir_f" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['nop']})*{c_(y)}{R['nir']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.000")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 台達電損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 80
SCN_T = {"A": "情境 A：原始顆數（顆數係數 1.0；v0.12 起本庫基準＝情境 B）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("aip", "  AI 電源營收", "NT$B", "#,##0.0"), ("lc", "  液冷營收", "NT$B", "#,##0.0"), ("pe", "  電源及零組件", "NT$B", "#,##0.0"), ("inf", "  基礎設施", "NT$B", "#,##0.0"), ("au", "  自動化", "NT$B", "#,##0.0"), ("mo", "  交通", "NT$B", "#,##0.0"),
                                ("nai", "  非 AI 營收", "NT$B", "#,##0.0"), ("aall", "  AI 電源＋液冷 ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"),
                                ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收 − 錨點；FY2026 模型 AI 電源 − 錨點；FY2026 毛利 − 錨點").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：AI 電源差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['aip']}-a26_aip"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：毛利差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['gmn']}*D{PB['A']['rev']}+d_ai*D{PB['A']['aip']}-a26_gp"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 80, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：電源與液冷內容、份額（權威） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 電源市場、台達電份額、每架內容（本專案權威，規則第 13 條）。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("gw", "部署 GW（情境 A）", "GW", lambda y: f"=PL!{FC[y]}{A_['gw']}", "0.0", "AI IT 資本支出 ÷ 每 GW IT 資本支出"),
        ("pkw", "AI 電源內容", "US$/kW", lambda y: f"=PL!{FC[y]}{A_['pkw']}", "#,##0", "In 頁 B 節（2026–27 Derived，2028 起 Assumed）"),
        ("prk", "每架 AI 電源內容（以 VR200 每 GW 架數換算）", "US$K/架", lambda y: f"=PL!{FC[y]}{A_['pkw']}*1000000/racks_vr/1000", "#,##0", "1 GW ÷ 4,052 架 ≈ 247 kW／架"),
        ("pmk", "AI 電源市場", "US$B", lambda y: f"=PL!{FC[y]}{A_['pmk']}", "#,##0.0", "含 ASIC 機架（假設每 kW 內容相同）"),
        ("sh", "台達電有效份額", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "FY2026 以公司 AI 相關占比校準"),
        ("aip", "台達電 AI 電源營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['aip']}/fx", "#,##0.00", ""),
        ("lc", "台達電液冷營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['lc']}/fx", "#,##0.00", "FY2026 以公司液冷占比校準"),
        ("lrk", "台達電液冷每架營收（以 VR200 架數換算）", "US$K/架", lambda y: f"=PL!{FC[y]}{A_['lc']}/fx/PL!{FC[y]}{A_['gw']}/racks_vr*1000000", "#,##0", "對照：橋接表機架內液冷 VR200 US$56K（台達電多做機架外 CDU、Liquid-to-Air）"),
        ("pgw", "每 GW 台達電 AI 營收（電源＋液冷）", "US$M/GW", lambda y: f"=PL!{FC[y]}{A_['pgw']}", "#,##0", ""),
        ("oth", "其他廠商 AI 電源（光寶、Flex、麥格米特、Vertiv 等）＝ 市場 − 台達電", "US$B", lambda y: f"=PL!{FC[y]}{A_['pmk']}-PL!{FC[y]}{A_['aip']}/fx", "#,##0.0", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：FY2026 有效份額低於凱基的 Blackwell PSU 約 70%，因為本模型分母含 ASIC 機架與 HVDC、BBU；份額路徑（相對值）與 2028 年起的每 kW 內容是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」與十年 P/E 規則值（年底落後 P/E，基準為 FY2027 EPS）"
V["A2"] = "循環位置：台達電不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 NVIDIA、Arista；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-10-30 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "#,##0"
V.cell(7, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(8, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(9, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(10, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((7, 2026), (8, 2027), (9, 2028), (10, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(11, 1, "未來 12 月 EPS（模型 A）≈ ¼ FY2026＋¾ FY2027"); V["B11"] = f"=0.25*PL!D{SUMR[('A', 'eps')]}+0.75*PL!E{SUMR[('A', 'eps')]}"; V["B11"].number_format = "0.00"
V.cell(12, 1, "  現價 ÷ 未來 12 月 EPS（模型 A）"); V["B12"] = "=px/B11"; V["B12"].number_format = "0.0"
V.cell(13, 1, "近 12 月 P/E（Goodinfo，10-22）"); V["B13"] = "=pe_ttm"; V["B13"].number_format = "0.0"
V["E8"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"; V["E12"] = "Andy 先驗 D10a 答 30x（未來 12 月）"
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
hdr(V, 27, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：Andy 合理 P/E", "對照：十年中位數"])
vr = 28
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_fair_andy").number_format = "0.0"; V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。")
vr += 2
V.cell(vr, 1, "5. 對照：以 Andy 合理 P/E 25x 計算的隱含價（只作對照，非規則值）").font = BOLD; vr += 1
hdr(V, vr, ["情境", "FY2027 EPS", "× 25x", "相對現價", "FY2028 EPS", "× 25x", "相對現價"]); vr += 1
IMP = {}
for s in SC:
    V.cell(vr, 1, s)
    V.cell(vr, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 3, f"=B{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 4, f"=C{vr}/px-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 6, f"=E{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
    IMP[s] = vr; vr += 1
assert vr < 50, vr
PE_CMC = [17.2, 16.4, 15.3, 14.9, 23.8, 24.4, 21.3, 23.5, 31.5, 41.5]          # companiesmarketcap 年底 P/E 2016–2025
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2016–2025 共 10 年）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。只有一個來源（第二來源：未找到；Goodinfo 歷年加權平均 P/E 區間 20.96–44.37 只作量級核對）。2016–2019 年台達電為低成長電源零組件公司（約 15–17x），AI 前後的倍數口徑不同，規則值可能低估 AI 期的合理倍數。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# ---------------- Val 第 0 節：目標價與上檔空間（v0.1.1 新增；放在既有空白列，不位移既有列） ----------------
TP0 = 41; TPR = {}
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
V.cell(TPR["m"], 2, "P/E 規則值：各情境 FY2027 EPS × 2016–2025 年底 P/E 依循環位置取分位（見頂＝25 分位、中段＝中位數、谷底＝75 分位；companiesmarketcap 一個來源）。隱含價約為 FY2027 年底價、未折現，作 12 個月目標價的近似；情境 A／B／C 即本模型三個顆數情境（A 最高、C 最低）。偏誤：十年倍數含 2016–2019 年低成長電源期低倍數（約 15–17x），目標價可能偏低。")
for c_ in range(1, 6):
    V.cell(TPR["tp"], c_).fill = KEY; V.cell(TPR["up"], c_).fill = KEY
V.cell(TPR["tp"], 1).font = BOLD
defname("old_rule_px", f"Val!$B${TPR['tp']}"); defname("old_rule_up", f"Val!$B${TPR['up']}")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額、液冷每 GW、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 電源營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 電源營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 電源營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.85、2028 0.70、2029 0.65（光寶、Flex、中國廠追趕）", dict(rel="RLO")),
         ("份額持平（高）", "份額相對值 1.0", dict(rel="RHI")),
         ("電源內容 低", "每 kW：2026 380、2027 500、2028–29 550", dict(pkw="PKLO")), ("電源內容 高", "每 kW：2026 520、2027 700、2028 800、2029 850", dict(pkw="PKHI")),
         ("液冷每 GW 低", "相對值 0.9／0.85／0.8", dict(lcr="LCLO")), ("液冷每 GW 高", "相對值 1.3／1.4／1.5", dict(lcr="LCHI")),
         ("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("非 AI 分部成長 低", "四條非 AI 分部成長率取低值", dict(gsuf="LO")), ("非 AI 分部成長 高", "四條非 AI 分部成長率取高值", dict(gsuf="HI")),
         ("AI 毛利率差距 4pt", "AI 電源毛利率只高 4 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 12pt", "高 12 個百分點", dict(d="d_ai_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi")),
         ("AI 相關含液冷（讀法）", "FY2026 AI 電源＝營收 ×（26% − 12.5%）；份額分子變小", dict(aip="aipx", nai="naix"))]
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("PKLO", "pkw", "_lo"), ("PKHI", "pkw", "_hi"), ("LCLO", "lcr", "_lo"), ("LCHI", "lcr", "_hi")):
    for y in (26, 27, 28, 29):
        if base in ("rel", "lcr") and y == 26: continue
        src = f"{base}_{y}{suf}"
        target = f"{nm}_{y}"
        wb.defined_names[target] = DefinedName(target, attr_text=wb.defined_names[src].attr_text)
for k in ("pe", "inf", "au", "mo"):   # 分部成長的低／高別名（Sens 用）：g_peLO27 等
    for nm, suf in (("LO", "_lo"), ("HI", "_hi")):
        for y in (27, 28, 29):
            wb.defined_names[f"g_{k}{nm}{y}"] = DefinedName(f"g_{k}{nm}{y}", attr_text=wb.defined_names[f"g_{k}{y}{suf}"].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "aip", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "aip", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
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
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。每個案例的 FY2026 有效份額與液冷每 GW 依該案例重算，所以 FY2026 恆等於錨點。「電源內容」案例的 FY2026 份額會反向調整（內容低 → 份額高），所以只影響 FY2027 起的成長路徑。")
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 券商估計、研調市場與份額。只作事後對照，不為對上數字而調參數。券商皆未具名或單一來源；2027 共識（付費）未找到"
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
rec("FY2026 營收 vs 券商（10-01，3Q 前）", "=b_rev26", pl("rev", 2026), "旺得富 2026-10-01", "券商 3Q 估 2,093 億，實際 1,971 億", fmt="#,##0.0")
rec("FY2026 EPS vs 本土券商（10-01）", "=b_eps26a", pl("eps", 2026), "旺得富 2026-10-01", "")
rec("FY2026 EPS vs 美系券商（05-04）", "=b_eps26b", pl("eps", 2026), "CMoney 2026-05-04", "")
rec("FY2027 營收 vs 投顧研究團隊", "=b_rev27", pl("rev", 2027), "Yahoo 新聞 2026-10-12", "", fmt="#,##0.0")
rec("FY2027 EPS vs 投顧研究團隊（10-12）", "=b_eps27c", pl("eps", 2027), "Yahoo 新聞 2026-10-12", "")
rec("FY2027 EPS vs 美系券商（05-04）", "=b_eps27b", pl("eps", 2027), "CMoney 2026-05-04", "")
rec("FY2027 EPS vs 本土券商（10-01）", "=b_eps27a", pl("eps", 2027), "旺得富 2026-10-01", "三家區間 62.7–80.8")
rec("FY2027 營收年增（模型）", None, pl("g", 2027), "—", "", fmt="0.0%")
RC.cell(rr_, 1, "B. 市場與份額（先驗比對的機制檢查）").font = BOLD; rr_ += 1
rec("CY2026 AI 電源市場（US$B）vs UBS", "=ubs_26", pl("pmk", 2026), "UBS（經 TechNews 2026-10-12）", "口徑未知（是否含 BBU、HVDC、非 NVIDIA）", fmt="#,##0.0")
rec("CY2027 AI 電源市場（US$B）vs UBS", "=ubs_27", pl("pmk", 2027), "同上", "", fmt="#,##0.0")
rec("FY2026 台達電有效份額 vs 凱基 Blackwell PSU 份額", "=kgi_sh", lambda s: f"=PL!D{PB[s]['sh0']}", "凱基 2025-07", "口徑不同：本模型分母含 ASIC 機架、HVDC、BBU", fmt="0.0%", pt=True)
rec("FY2027 每 GW 台達電 AI 營收（US$M）", None, pl("pgw", 2027), "—", "先驗比對推導約 300–550", fmt="#,##0")
rec("FY2027 台達電液冷每架（US$K，VR200 架數）vs 橋接表機架內液冷", "=cool_vr", lambda s: f"=Units!D{UR['lrk']}", "橋接表 v0.2.1（Morgan Stanley）", "台達電液冷多在機架外（CDU、Liquid-to-Air），所以兩者不是同一個東西", fmt="#,##0")
rec("FY2027 AI 電源＋液冷 ÷ 營收", None, pl("aall", 2027), "—", "", fmt="0.0%")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 電源毛利率＝此值＋d_ai", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "AI 電源占比上升使毛利率上升", fmt="0.0%", pt=True)
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E15", "—", "券商 10-01 估 14.7（以 4Q 營收 2,511 億為基礎）", fmt="0.00")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝台達電法說簡報與逐字稿，已逐份抓取；券商與研調皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：4Q25 法說簡報（FY2025 損益、部門營收與利益）", AM4Q25),
        ("公司：2Q26 法說逐字稿（1H26 損益、AI 占比、液冷占比、800V 時程、資本支出、零組件短缺）", TR2Q26),
        ("二手：BigGo 2Q26 法說摘要（AI 相關朝 28%；對手名單）", "https://finance.biggo.com/news/TW_2308.TW_2026-07-30"),
        ("二手：數位時代 2026-02-26（2025 部門營收比重 50／33／10／7）", "https://www.bnext.com.tw/article/90178/2026-delta-earnings-call"),
        ("二手：TechNews 2026-10-12（3Q26 營收、AI 相關 2025 年 18%、UBS AI 伺服器電源市場）", TN1012),
        ("券商：凱基 Power supply sector 2025-07-01（每架 PSU 內容、Blackwell 份額約 70%）", KGI),
        ("券商：旺得富 2026-10-01（本土投顧 EPS、4Q26 估計、HVDC 占 2027 營收 5–10%、18kW 與 HVDC 毛利率約 50%）", "https://wantrich.chinatimes.com/news/20261001900208-420101"),
        ("券商：CMoney 2026-05-04（美系 EPS 2026–28）", "https://www.cmoney.tw/notes/note-detail.aspx?nid=1180676"),
        ("券商：Yahoo 新聞 2026-10-12（投顧 2027 營收 1.05 兆、EPS 62.68；Vicor 337 調查）", "https://tw.news.yahoo.com/%E5%8F%B0%E9%81%94%E9%9B%BBq3%E7%87%9F%E6%94%B6%E5%89%B5%E6%96%B0%E9%AB%98-337%E8%AA%BF%E6%9F%A5%E8%A1%9D%E6%93%8A-%E6%B3%95%E4%BA%BA%E9%80%99%E6%A8%A3%E7%9C%8B-035400801.html"),
        ("研調：Morgan Stanley NVL72 冷卻 BOM（經 Tom's Hardware 2025-11-05）", "https://finance.yahoo.com/news/cooling-system-single-nvidia-blackwell-154500100.html"),
        ("資料商：Yahoo 股市（10-23 收盤）、Goodinfo（10-22 收盤、近 12 月 P/E）", "https://tw.stock.yahoo.com/quote/2308.TW"),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/delta-electronics/pe-ratio/"),
        ("市場層：markets/network、markets/chip_bridge、tsmc 成品（本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 法說（10 月底）；部門季度損益；AI 電源與液冷的季度金額；P/E 第二來源；前瞻 P/E 十年歷史（Bloomberg）；券商共識（付費）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("台達電（2308）個股模型 v0.3（AI 半導體 Project：電源與散熱次產業樣板）　2026-10-28", True),
 ("版本：v0.3（2026-10-28：依個股模型原則重建營收層——營收列改為官方四大事業群〔電源及零組件、交通、自動化、基礎設施〕，AI 電源為電源及零組件子列、液冷為基礎設施子列，分部合計＝合併營收；非 AI 單一成長率改為四條分部成長率）；v0.2.x（2026-10-25～27：目標價、級距、Siliconomics 改讀）；v0.1.1（2026-10-25：Val 新增第 0 節「目標價與上檔空間」、In 新增情境權重 tp_wA／tp_wB／tp_wC；其餘數值與 v0.1 相同）。v0.1（初版）：AI 電源每 kW 內容、台達電份額、液冷每 GW 營收在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）；奇鋐等散熱廠之後讀這條液冷線。", False),
 ("模型命題：台達電的 AI 營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）×（AI 電源每 kW 內容 × 台達電份額＋液冷每 GW 營收）。800V HVDC 把交流轉直流移到機架外，每 GW 可做的金額變大；同時光寶、Flex、中國廠追趕份額。股價兩週跌約兩成後，約 22–28 倍未來 12 月 EPS。", False),
 ("核心問題：每 kW 電源內容上升（800V）能不能抵銷份額流失與零組件成本，讓 AI 營收在 2027–28 年維持高成長；現價隱含的路徑合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029（三年）。金額：十億新台幣；市場以美元計、以匯率換算。口徑：歸屬母公司淨利。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 電源營收：markets/network 的 AI IT 資本支出 × 情境顆數係數 ÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 kW 電源內容（凱基推導、2028 起 Assumed）× 匯率 × 台達電有效份額（FY2026 以公司 AI 相關占比校準）× 份額相對值（Assumed）。", False),
 ("液冷營收：FY2026 以公司液冷占比校準每 GW 營收 × 相對值（Assumed）。營收分部：電源及零組件＝AI 電源＋非 AI（年增 5／4／3%）；基礎設施＝液冷＋非液冷（年增 9／7／6%）；自動化（5／5／4%）；交通（0／2／2%）；皆 Assumed，FY2025、FY2026 依公司事業群比重（50／7／10／33%、51／6／8／35%）。毛利率：AI 電源比其他營收高 d_ai（Assumed），其他毛利率為 FY2026 殘差（Derived）。營業費用：彈性 phi（Derived）。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績＋假設（4Q 季增、毛利率、營業費用率）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：現價 ÷ FY2027、FY2028 EPS；十年 P/E 規則值（單一資料商來源）× FY2027 EPS；循環位置以本模型營收最高年判定（台達電不套設備股見頂年）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文＝法說簡報與逐字稿；券商皆未具名或單一來源。", False),
 ("", False),
 ("已知限制：(1)「AI 相關」是否含液冷，公司說法不一致（Sens「AI 相關含液冷」）；(2) 2028 年起每 kW 內容（800V）與份額相對值為 Assumed，是 FY2028 EPS 最敏感的輸入；(3) 部署 GW 以 VR200 每 GW IT 資本支出換算，ASIC 機架的每 GW 資本支出較低，GW 可能低估；(4) 十年 P/E 只有一個來源，且 2016–2019 年為 AI 前的低倍數；(5) 未計 Vicor 337 調查（垂直供電）與零組件短缺的下檔；(6) 3Q26 法說（10 月底）後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

V2["pe10"] = (list(zip(range(2016, 2026), PE_CMC)), "companiesmarketcap 年底 P/E，Interested-party，單一來源；同規則值")
V2["old_row"] = TP0
# ================= v0.2（2026-10-25）：PL 延伸到 FY2032、Val 第 0 節目標價（DCF 與 Comparable 加權）、Sens 目標價參數 =================
# 全部放在各頁既有內容之後，不位移既有儲存格（FY2025–FY2029 數值與 v0.1.1 相同）
FADEFILL = PatternFill("solid", fgColor="E7E6E6")
XY = (2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032); XC = {y: "CDEFGHIJ"[i] for i, y in enumerate(XY)}; FADE = (2030, 2031, 2032)
K2 = V2["keys"]
for rw in range(4, IN.max_row + 1):
    if IN.cell(rw, 1).value in ("tp_wA", "tp_wB", "tp_wC", "tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "nim10_f"):
        IN.cell(rw, 5).fill = KEY
FD.yellow(IN)   # v0.5：標準版 DCF 的判斷參數黃底（Andy 可改）
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
# v0.5：0.1 節改為標準版 DCF（FCFF＋WACC）摘要（tools/fcff_dcf.py），明細在 DCF 頁；0.1 節原占 20 列，摘要 18 列，0.2 節起位置不變
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（v0.5）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2027 起整年）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0")
DCFW["A2"] = ("EBIT：FY2027–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸表只有淨利）由淨利還原：EBIT＝淨利 ÷ nir_f − nop_f（nir_f＝歸屬母公司淨利 ÷ 稅前，與延伸表同口徑）。"
              "NOPAT＝EBIT ×（1 − tax_w），為合併口徑（含非控制權益），評價時扣非控制權益帳面值；業外收益不進 NOPAT，現金與長期投資另在評價時加入。"
              "折舊攤銷、資本支出、營運資金：近三年實績占營收比例（In 頁 fc_da、fc_cx、fc_wc），FY2029 起資本支出比例線性收斂到 FY2032 穩態 fc_cx_ss。v0.5.2：淨現金由資產負債表日 2026-09-30 滾到評價日（＋各情境 FY2027 FCFF × stub_frac − 已宣告股利；Val 0.1 節）。")
FY_X = list(range(2027, 2033))
_ebit = lambda s_, y: f"PL!{XC[y]}{PB[s_]['oi']}" if y <= 2029 else f"=PL!{XC[y]}{XR[s_]['ni']}/nir_f-nop_f"
DPATHS = [dict(name=s_, label=SCN_T[s_].replace("情境 A：", "情境 A（順利）：").replace("情境 B：", "情境 B（平平）：").replace("情境 C：", "情境 C（不如預期）："),
               rev={y: f"PL!{XC[y]}{XR[s_]['rev']}" for y in FY_X}, rev0=f"PL!{XC[2026]}{XR[s_]['rev']}",
               ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "tax_w" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷ nir_f − nop_f") for s_ in SC]
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 1.0, conv_from=2029, title="三情境 FCFF 明細（NT$B；會計年度＝日曆年，FY2027 起整年）", fmt="#,##0.0", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="(213.212)+8.252-(12.779+26.462+71.385)-4.454-66.233",
                             nc_note="2026-09-30 資產負債表（NT$B）：現金及短期投資 213.212＋長期投資 8.252 − 借款（12.779＋26.462＋71.385）− 租賃 4.454 − 非控制權益 66.233") + 1
DCFW.column_dimensions["A"].width = 60; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
SDCF = FD.write_summary(wb, V, t0 + 1, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt=V2["pxfmt"],
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.5；明細見 DCF 頁；B／C／D＝A 順利／B 平平／C 不如預期）", beta_note="自算產業 β 重新加槓桿")
for _k in ("pvx", "pvtv", "ev", "nc"):
    for _c in "BCD": V[f"{_c}{SDCF[_k]}"].number_format = "#,##0.0"
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"], h1=SDCF["h"])
r_ = t0 + 22
V.cell(r_, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; r_ += 1
hdr(V, r_, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
P1 = r_
for nm, tk, pe, dt in V2["peers"]:
    V.cell(r_, 1, nm); V.cell(r_, 2, tk); cc = V.cell(r_, 3, pe); cc.font = BLUE; cc.number_format = "0.0"
    V.cell(r_, 4, dt); V.cell(r_, 5, "stockanalysis.com 個股頁 Forward PE（分析師共識；價格日期如左）"); V.cell(r_, 6, "Interested-party（賣方共識）"); r_ += 1
P2 = r_ - 1; T["p1"], T["p2"] = P1, P2
tset(r_, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{P1}:C{P2})", None, None], "0.0", V2["peer_note"]); r_ += 1
w1, w2 = V2["fwd_w"]
tset(r_, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*({w1}*PL!{XC[2027]}{XR[s]['eps']}+{w2}*PL!{XC[2028]}{XR[s]['eps']})" for s in SC), None, None], V2["epsfmt"], V2["fwd_note"]); r_ += 1
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
     [(lambda e: f"=IF(px/{e}<=MIN({PE10}),0,IF(px/{e}>=MAX({PE10}),1,PERCENTRANK({PE10},px/{e})))")(f"PL!{XC[2028]}{XR[s]['eps']}") for s in SC], "0%",
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
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt=V2["pxfmt"],
                    title="目標價參數敏感度（v0.5 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=True)
TPS_LAST = _tvs[1]
USD_ROW = next(r for r in range(_tvs[0], _tvs[1] + 1) if SE.cell(r, 1).value == "美元口徑折現")   # 第二輪修正：Sens「美元口徑折現」列
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
    elif isinstance(v_, str) and v_.startswith("版本："): RD.cell(rw, 1, "版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察；PL 頁最後新增三情境自由現金流列（淨利＋折舊 − 資本支出，不含營運資金變動）、In 頁最後新增折舊 ÷ 營收、資本支出 ÷ 營收（數值不變））。v0.2（2026-10-25：目標價改為 DCF 與 Comparable 加權、三情境延伸到 FY2032；FY2025–FY2029 數值不變）。前版：" + v_[3:])
rn = RD.max_row + 1
for t_ in V2["readme_add"]:
    RD.cell(rn, 1, t_); rn += 1
if os.environ.get("V2_ROWMAP"):
    import json as _json
    _json.dump({"XT": XT, "XH": XH, "XR": XR, "T": T, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "RECV2": RECV2},
               open(os.environ["V2_ROWMAP"], "w"), ensure_ascii=False, indent=1)

# ================= v0.2.1（2026-10-25）：自由現金流（算力收支 2026-10-25 請求；投資地圖依列標籤「自由現金流」讀取） =================
# In 頁最後新增折舊 ÷ 營收、資本支出 ÷ 營收（FY2025 實績為錨，各年沿用）；PL 頁最後新增三情境自由現金流列。不位移既有儲存格。
ir = IN.max_row + 2
IN.cell(ir, 1, "K. 自由現金流參數（v0.2.1 新增；以 FY2025 年報實績為錨，FY2026 起各年沿用；只供 PL 頁自由現金流列，不進 EPS 與目標價）").font = BOLD; ir += 1
for key_, item_, val_, s1_ in (("da_r", "折舊與攤銷 ÷ 營收", "=27.839/a25_rev", "FY2025 折舊與攤銷 278.39 億元 ÷ FY2025 營收（Q 頁 B4）；stockanalysis（S&P Global）現金流量表，2026-10-25 擷取；營業現金流與 Yahoo 股市四季合計 984.74 億元一致"),
                               ("cx_r", "資本支出 ÷ 營收", "=46.091/a25_rev", "FY2025 資本支出（購置不動產、廠房及設備）460.91 億元 ÷ FY2025 營收（Q 頁 B4）；stockanalysis（S&P Global）現金流量表，2026-10-25 擷取；營業現金流與 Yahoo 股市四季合計 984.74 億元一致")):
    IN.cell(ir, 1, key_); IN.cell(ir, 2, item_); IN.cell(ir, 3, "%")
    IN.cell(ir, 5, val_).number_format = "0.00%"
    IN.cell(ir, 7, "Verified（FY2025 實績）／Assumed（各年沿用）"); IN.cell(ir, 8, s1_); IN.cell(ir, 9, "FY2026 起各年＝營收 × 本比率（Assumed：資本密集度不變）")
    wb.defined_names[key_] = DefinedName(key_, attr_text=f"In!$E${ir}"); ir += 1
fr = PL.max_row + 3
PL.cell(fr, 1, "自由現金流（v0.2.1 新增；算力收支 2026-10-25 請求）：淨利（歸屬母公司，延伸表）＋折舊與攤銷（營收 × In 頁 da_r）− 資本支出（營收 × In 頁 cx_r）；不含營運資金變動。FY2030 起營收、淨利為收斂法").font = BOLD; fr += 1
hdr(PL, fr, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); FCF_H = fr; fr += 1
FCFR = {}
for s in SC:
    PL.cell(fr, 1, f"自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）｜{SCN_T[s]}"); PL.cell(fr, 2, V2["unit"])
    for y in XY[1:]:
        c = XC[y]
        PL[f"{c}{fr}"] = f"={c}{XR[s]['ni']}+{c}{XR[s]['rev']}*da_r-{c}{XR[s]['rev']}*cx_r"; PL[f"{c}{fr}"].number_format = "#,##0.0"
        if y in FADE: PL[f"{c}{fr}"].fill = FADEFILL
    PL.cell(fr, 11, "淨利、營收引用本頁延伸表同情境列；折舊、資本支出以 FY2025 實績比率 × 營收")
    FCFR[s] = fr; fr += 1
so = SO.max_row + 1
SO.cell(so, 1, "[Verified] 台達電 FY2025 年報：折舊與攤銷 278.39 億元、資本支出 460.91 億元、營業現金流 984.74 億元（資料商轉載合併現金流量表）"); SO.cell(so, 2, "https://stockanalysis.com/quote/tpe/2308/financials/cash-flow-statement/")
SO.cell(so + 1, 1, "[Assumed] FY2026 起折舊 ÷ 營收、資本支出 ÷ 營收沿用 FY2025 實績比率（In 頁 da_r、cx_r）；自由現金流不含營運資金變動"); SO.cell(so + 1, 2, "本模型假設")

# ================= v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（工作單；Andy 2026-10-25 14:00、14:17） =================
# 全部放在 Val、README 現有最後一列之後，不位移既有儲存格
OBS = ["目標價 NT$1,211（DCF 980、Comparable 相對倍數調整後 1,442 各半），比現價 1,335（2026-10-30）低 9.3%，級距 0%～−25%，投資建議維持中立。v0.4 滾入 3Q26：毛利率 34.1%（原假設 36%；含存貨跌價損失約 39 億元）、營業利益 326 億、EPS 9.04；加權目標價 1,237 → 1,211（−2.1%）。上檔空間由 −22.9%（10-23 標示價 1,605）縮小到 −9.3%，主要來自股價下跌 16.8%，不是模型上修。", "FY2026 EPS 37.89 → 36.71（−3.1%）、FY2027 EPS 60.71 → 59.44（−2.1%）；4Q26 假設季增 10%、毛利率 34.5%（公司：4Q 營收一定高於 3Q、毛利率不會比 3Q 差）。存貨跌價集中在電源及零組件、預計 1Q27 消化，若毛利率回到 36% 以上，FY2027 EPS 有上修空間。", "待查：股感轉述法說稱「AI 相關營收已約占一半、液冷約 13%」，與本模型 FY2026 AI 相關 26%（不含液冷）口徑差很大；若公司口徑改為含資料中心基礎設施，需重新校準 AI 電源份額（Siliconomics M17 pwr_dsh_26，屬產業變數），會影響光寶與 Recon。下一個檢驗點是 4Q26 營收（2027-01 月營收）與 HVDC ±400V 出貨。"]   # 重點觀察（與 README.md 相同文字；數字取自重算後的 xlsx）
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

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "delta")
# ---------------- v0.4（2026-11-02）：3Q26 實績（In 頁最後追加；Q 頁 D 欄改讀） ----------------
_ri = IN.max_row + 2
IN.cell(_ri, 1, "v0.4：3Q26 季報實績（2026-10-29 公布、法說；Q 頁 D 欄）").font = BOLD; _ri += 1
for _k, _lab, _v, _tag, _src, _src2 in [
        ("q3_gp", "3Q26 毛利", 67.2, "Verified", "3Q26 法說簡報（股感 2026-10-29 轉述）：毛利 672 億元、毛利率 34.1%（2Q26 35.6%）", "存貨跌價損失約 39 億元，集中在電源及零組件，預計 1Q27 陸續消化"),
        ("q3_op", "3Q26 營業利益", 32.6, "Verified", "同上：營業淨利 326 億元、營業利益率 16.6%、營業費用率 17.5%", "口袋學堂 2026-10-29：約 326 億元，季增 7%"),
        ("q3_pti", "3Q26 稅前利益", 33.7, "Verified", "同上：稅前淨利 337 億元", "—"),
        ("q3_ni", "3Q26 歸屬母公司淨利", 23.495, "Verified", "優分析／鉅亨 2026-10-29：稅後純益 234.95 億元（季減 6.53%）", "前三季 691.86 億元"),
        ("q3_eps", "3Q26 EPS（公司）", 9.04, "Verified", "同上：每股純益 9.04 元；前三季 26.64 元", "只作核對")]:
    IN.cell(_ri, 1, _k); IN.cell(_ri, 2, _lab); IN.cell(_ri, 3, "NT$" if _k == "q3_eps" else "NT$B")
    IN.cell(_ri, 5, _v).font = BLUE; IN.cell(_ri, 7, _tag); IN.cell(_ri, 8, _src); IN.cell(_ri, 10, _src2)
    wb.defined_names[_k] = DefinedName(_k, attr_text=f"In!$E${_ri}"); _ri += 1
RD.cell(RD.max_row + 1, 1, "v0.4（2026-11-02）：滾入 3Q26 實績（毛利率 34.1%、營業利益 326 億、歸屬母公司淨利 234.95 億、EPS 9.04；Q 頁 D 欄改讀 In 頁最後新增的 q3_*）；4Q26 毛利率 34.5%（公司：不會比 3Q 差）、營業費用率 17.5%、季增 10%；股價改 2026-10-30 收盤 1,335（Andy 11-01：上檔空間用最新收盤）。")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.4", "2026-11-02")  # v0.3.1：In 頁版面（個股模型原則第 8 節；數值不變）
# ---------------- Siliconomics v0.13（2026-11-04；決策 S37）：Sens「Wattonomics 錨定」案例 ----------------
# In 頁最後追加 Z 節（wta_27／28／29，數值取 SIL.WTA）；Sens 頁既有內容最下方另起小表與案例區塊。既有儲存格不位移、數值不變。
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
def _wta_in():
    r = IN.max_row + 2
    IN.cell(r, 1, "Z. Wattonomics 錨定係數（敏感度用；Siliconomics v0.13）").font = BOLD; r += 1
    def w(k, lab, unit, lo, b, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((k, lab, unit, lo, b, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).font = BLUE; IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[k] = DefinedName(k, attr_text=f"In!$E${r}"); r += 1
    SIL.wta_rows(w)
def _wta_sens(mf, setting, how):
    r0 = SE.max_row + 3
    SE.cell(r0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    SE.cell(r0 + 1, 1, WTA_NOTE)
    SE.cell(r0 + 2, 1, how)
    ncol = max(j for j in range(1, 30) if SE.cell(3, j).value is not None)
    H = [SE.cell(3, j).value for j in range(1, ncol + 1)]
    hdr(SE, r0 + 3, H)
    base_b = "情境 B" in str(SE.cell(4, 1).value)
    rb, rB = r0 + 4, r0 + 5
    rw = rB + (0 if base_b else 1)
    rd = rw + 1
    WR = engine(SE, rd + 4, "案例區塊：Wattonomics 錨定（情境 B × wta）", scn="B", mf=mf)
    SE.cell(rb, 1, "基準（主表第 4 列）"); SE.cell(rb, 2, SE.cell(4, 1).value)
    for j in range(3, ncol + 1):
        L = SE.cell(4, j).column_letter
        SE[f"{L}{rb}"] = f"={L}4"; SE[f"{L}{rb}"].number_format = SE[f"{L}4"].number_format
    if not base_b:
        put_case(rB, "情境 B（PL 區塊 B）", "Siliconomics 基準係數（未乘 wta）", PB["B"], "PL!")
    put_case(rw, "Wattonomics 錨定", setting, WR, "")
    for k, (rr, lab) in enumerate(((rb, "錨定 − 基準（主表第 4 列）"), (rB, "錨定 − 情境 B")) if not base_b else ((rb, "錨定 − 基準（主表第 4 列）＝情境 B"),)):
        ri = rd + k
        SE.cell(ri, 1, lab); SE.cell(ri, 2, "差異%（毛利率、營益率等百分比欄為百分點差）")
        for j in range(3, ncol + 1):
            if "相對基準" in str(H[j - 1]): continue
            L = SE.cell(4, j).column_letter
            pct = "%" in SE[f"{L}4"].number_format
            SE[f"{L}{ri}"] = f"={L}{rw}-{L}{rr}" if pct else f"={L}{rw}/{L}{rr}-1"
            SE[f"{L}{ri}"].number_format = "0.0%"
    return dict(r0=r0, base=rb, case=rw, diff=rd, block=WR)
_wta_in()
WTA_SENS = _wta_sens({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}, '情境 B；AI IT 資本支出（加速器顆數 × 每顆 IT capex；→ 部署 GW × 每 kW 電源內容） CY2027 起 × wta_27／28／29', '做法：引擎的 AI 部署係數 mf（原只乘在 AI IT 資本支出列、FY2027 起）改為年別 dict（wta_27／28／29）；FY2026 不乘。FY2030 以後無延伸表（Sens 引擎只到 FY2029）。目標價：Sens 目標價表只有 DCF 參數案例，本案例不重算目標價。')
# ---------------- v0.4.1（2026-11-05）：展示層（個股 HTML 版面規格 v1.0；docs/plan/個股HTML版面規格.md） ----------------
# 只寫 README 頁（既有列之後）；其他工作表不動。數字一律以公式取自各頁。
def _row(ws, label, start=1, exact=False):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, 1).value
        if isinstance(v_, str) and ((v_.strip() == label.strip()) if exact else v_.strip().startswith(label.strip())): return r_
    raise KeyError((ws.title, label))
_V = lambda lab, c="B": f"Val!{c}{_row(V, lab)}"
_I = lambda k: f"In!E{_row(IN, k, exact=True)}"
_pa, _pb, _pc = (_row(PL, f"情境 {s}", exact=True) for s in SC)
_P = lambda lab, c, st=_pa: f"PL!{c}{_row(PL, lab, st, exact=True)}"
_U = lambda lab, c: f"Units!{c}{_row(UN, lab)}"
_S = lambda lab: (f"Sens!A{_row(SE, lab)}", f"Sens!L{_row(SE, lab)}")
TP, DCF, CMP, PX, UP, REC = _V("目標價（DCF 與 Comparable 加權）"), _V("DCF 目標價（情境加權）"), _V("Comparable 目標價（相對倍數調整後）"), _V("現價（"), _V("上檔空間：加權目標價"), _V("投資建議")
_dv = _row(V, "DCF 情境價值推到 12 個月後")
T_ = lambda x, f="#,##0": f'TEXT({x},"{f}")'
B_ = lambda x: f'TEXT({x}*10,"#,##0")'   # 十億新台幣 → 億元
_dir = lambda x: f'IF({x}<0,"低 ","高 ")&TEXT(ABS({x}),"0.0%")'
_sm, _ss, _sc = _S("AI 部署 ×0.8"), _S("份額流失（低）"), _S("顆數情境 C")
rn = RD.max_row + 2
RD.cell(rn, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；A3、A19–A22 原文保留）").font = BOLD; rn += 1
for t_ in [
    "命題：台達電能否靠 AI 電源與液冷撐起 2027–28 年的高成長，以及現價是否已反映這條路徑；分成三個問題。",
    "命題 1｜每 kW 內容能否抵銷份額流失：800V 高壓直流讓每 kW 電源內容上升，但光寶、Flex 與中國廠追趕份額；兩者相抵後，AI 電源營收能否在 2027–28 年維持高成長。",
    "命題 2｜獲利能否跟上營收：AI 電源毛利率較高，但存貨跌價與零組件成本壓低毛利率；EPS 能否隨營收成長。",
    "命題 3｜現價合不合理：把 DCF 與同業倍數法加權後，每股價值與現價相比如何。",
    f'="重點 1｜加權目標價比現價"&{_dir(UP)}&"，投資建議"&{REC}&"：加權目標價約 "&{T_(TP)}&" 元，由 DCF 約 "&{T_(DCF)}&" 元與同業倍數法（相對倍數調整後）約 "&{T_(CMP)}&" 元加權（DCF 權重 "&TEXT({_I("tp_w_dcf")},"0%")&"）；現價 "&{T_(PX)}&" 元。"',
    f'="重點 2｜AI 電源是成長主力，但已假設份額逐年下滑：情境 A（順利）的 AI 電源營收由 FY2026 約 "&{B_(_P("  AI 電源營收", "D"))}&" 億元增至 FY2028 約 "&{B_(_P("  AI 電源營收", "F"))}&" 億元，AI 電源加液冷占營收由 "&TEXT({_P("  AI 電源＋液冷 ÷ 營收", "D")},"0%")&" 升到 "&TEXT({_P("  AI 電源＋液冷 ÷ 營收", "F")},"0%")&"；同期台達電有效份額由 "&TEXT({_U("台達電有效份額", "C")},"0.0%")&" 降到 "&TEXT({_U("台達電有效份額", "E")},"0.0%")&"，靠每 kW 電源內容由 "&{T_(_U("AI 電源內容", "C"))}&" 美元升到 "&{T_(_U("AI 電源內容", "E"))}&" 美元補回。"',
    f'="重點 3｜最大的變數是 AI 部署量與份額：一次只改一個輸入時，「"&{_sm[0]}&"」使情境 A 的 FY2028 EPS 變動 "&TEXT({_sm[1]},"+0.0%;-0.0%")&"，「"&{_ss[0]}&"」變動 "&TEXT({_ss[1]},"+0.0%;-0.0%")&"，「"&{_sc[0]}&"」變動 "&TEXT({_sc[1]},"+0.0%;-0.0%")&"。這些是敏感度測試的案例，不是門檻。"',
    f'="重點 4｜v0.5 起 DCF 改為標準版（無槓桿自由現金流折現），DCF 目標價約 "&{T_(DCF)}&" 元：WACC "&TEXT(wacc,"0.0%")&"（股權成本 "&TEXT(ke,"0.0%")&"＝台灣 10 年期公債 "&TEXT(rf,"0.00%")&" ＋ β "&TEXT(beta,"0.00")&" ×（ERP "&TEXT(erp,"0%")&" ＋ 國家風險溢酬 "&TEXT(crp,"0.00%")&"）；β 為自算產業 β（TW 散熱／電源同業中位數，Blume 調整）依台達電負債比重新加槓桿，ERP 與 RONIC＝WACC 為假設），永續成長率 "&TEXT(g_tv,"0.00%")&"（不超過台灣 10 年期公債殖利率）。情境 B 的終值現值占企業價值 "&TEXT(Val!C{SDCF["tvsh"]},"0%")&"、隱含終值 EV／EBITDA 約 "&TEXT(Val!C{SDCF["mult"]},"0.0")&" 倍；DCF 參數中加權目標價對 "&IF(MAX(ABS(Sens!D{_tvs[0] + 1}),ABS(Sens!F{_tvs[0] + 1}))>MAX(ABS(Sens!D{_tvs[0] + 2}),ABS(Sens!F{_tvs[0] + 2})),"β","ERP")&" 最敏感（β 低／高（β_u ∓ 0.3）時 "&TEXT(Sens!D{_tvs[0] + 1},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0] + 1},"+0.0%;-0.0%")&"；ERP 4%／6% 時 "&TEXT(Sens!D{_tvs[0] + 2},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0] + 2},"+0.0%;-0.0%")&"；WACC ±1pt 時 "&TEXT(Sens!D{_tvs[0]},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0]},"+0.0%;-0.0%")&"；以三情境加權目標價為基準）。"',
    f'="重點 5｜美元口徑折現（敏感度案例，非基準）：基準以台灣 10 年期公債 "&TEXT(rf,"0.00%")&" 折現並假設匯率固定；若改用美國 10 年期公債 "&TEXT(rf_us,"0.00%")&"＋ β ×（ERP＋國家風險溢酬）、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")＝"&TEXT(MIN(g_cap,rf_us),"0.00%")&"，股權成本 "&TEXT(rf_us+beta*(erp+crp),"0.0%")&"，加權目標價 "&TEXT(Sens!C{USD_ROW},"#,##0")&" 元（相對基準 "&TEXT(Sens!D{USD_ROW},"+0.0%;-0.0%")&"）。"',
    f'="重點 6｜終值占比："&IF(MAX(Val!B{SDCF["tvsh"]}:D{SDCF["tvsh"]})>0.8,"終值現值占企業價值超過 80% 的情境："&IF(Val!B{SDCF["tvsh"]}>0.8,"A ","")&IF(Val!C{SDCF["tvsh"]}>0.8,"B ","")&IF(Val!D{SDCF["tvsh"]}>0.8,"C ","")&"（DCF 價值主要來自終值，對 g 與 WACC 較敏感）；","三情境終值現值占企業價值皆未超過 80%；")&"情境 A／B／C 分別為 "&TEXT(Val!B{SDCF["tvsh"]},"0%")&"／"&TEXT(Val!C{SDCF["tvsh"]},"0%")&"／"&TEXT(Val!D{SDCF["tvsh"]},"0%")&"。"']:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "命題回覆（HTML 用；依本模型計算回答開頭三個命題）").font = BOLD; rn += 1
for t_ in [
    f'="每 kW 內容能否抵銷份額流失？答：在情境 A（順利）下"&IF({_U("AI 電源內容", "E")}/{_U("AI 電源內容", "C")}*{_U("台達電有效份額", "E")}/{_U("台達電有效份額", "C")}>1,"可以","不行")&"（前提是假設：2028 每 kW 電源內容 "&{T_(_I("pkw_28"))}&" 美元（In!pkw_28）與台達電份額相對值 "&TEXT({_I("rel_27")},"0.00")&"／"&TEXT({_I("rel_28")},"0.00")&"／"&TEXT({_I("rel_29")},"0.00")&" 倍（2027／2028／2029 ÷ 2026，In!rel_27–29）皆為 Assumed）。每 kW 電源內容由 CY2026 的 "&{T_(_U("AI 電源內容", "C"))}&" 美元升到 CY2028 的 "&{T_(_U("AI 電源內容", "E"))}&" 美元（"&TEXT({_U("AI 電源內容", "E")}/{_U("AI 電源內容", "C")}-1,"+0%;-0%")&"），台達電有效份額由 "&TEXT({_U("台達電有效份額", "C")},"0.0%")&" 降到 "&TEXT({_U("台達電有效份額", "E")},"0.0%")&"（相對 "&TEXT({_U("台達電有效份額", "E")}/{_U("台達電有效份額", "C")}-1,"+0%;-0%")&"）。營收成長的主因則是部署量：部署 GW 由 "&TEXT({_U("部署 GW", "C")},"0.0")&" 增至 "&TEXT({_U("部署 GW", "E")},"0.0")&"；合計使台達電 AI 電源營收由 CY2026 約 "&TEXT({_U("台達電 AI 電源營收", "C")},"0.0")&" 十億美元增至 CY2028 約 "&TEXT({_U("台達電 AI 電源營收", "E")},"0.0")&" 十億美元（FY2026 約 "&{B_(_P("  AI 電源營收", "D"))}&" 億元、FY2028 約 "&{B_(_P("  AI 電源營收", "F"))}&" 億元）。部署量較低的情境 B、C，FY2028 AI 電源營收約 "&{B_(_P("  AI 電源營收", "F", _pb))}&"、"&{B_(_P("  AI 電源營收", "F", _pc))}&" 億元。"',
    f'="獲利能否跟上營收？答：情境 A "&IF(AND({_P("EPS", "F")}>{_P("EPS", "D")},{_P("營業利益率", "F")}>{_P("營業利益率", "D")}),"可以","跟不上")&"。毛利率由 FY2026 的 "&TEXT({_P("毛利率", "D")},"0.0%")&" 升到 FY2028 的 "&TEXT({_P("毛利率", "F")},"0.0%")&"、營業利益率由 "&TEXT({_P("營業利益率", "D")},"0.0%")&" 升到 "&TEXT({_P("營業利益率", "F")},"0.0%")&"，EPS 由 FY2026 的 "&TEXT({_P("EPS", "D")},"0.00")&" 元增至 FY2027 的 "&TEXT({_P("EPS", "E")},"0.00")&" 元、FY2028 的 "&TEXT({_P("EPS", "F")},"0.00")&" 元。毛利率改善來自 AI 電源占比上升，前提是 AI 電源毛利率比其他營收高 "&TEXT({_I("d_ai")}*100,"0")&" 個百分點（假設）。情境 B、C 的 FY2028 EPS 約 "&TEXT({_P("EPS", "F", _pb)},"0.00")&"、"&TEXT({_P("EPS", "F", _pc)},"0.00")&" 元。"',
    f'="現價合不合理？答："&IF({UP}<0,"現價高於本模型的合理價值。","現價低於本模型的合理價值。")&"加權目標價約 "&{T_(TP)}&" 元（標準版 DCF "&{T_(DCF)}&"、同業倍數法 "&{T_(CMP)}&"），比現價 "&{T_(PX)}&" 元"&{_dir(UP)}&"，投資建議"&{REC}&"。DCF（無槓桿自由現金流、WACC "&TEXT(wacc,"0.0%")&"）三情境推到 12 個月後的每股價值約 "&{T_(f"Val!B{_dv}")}&"／"&{T_(f"Val!C{_dv}")}&"／"&{T_(f"Val!D{_dv}")}&" 元（情境 A／B／C）。市場共識目標價約 "&{T_(_V("市場共識目標價"))}&" 元，"&IF({_V("市場共識目標價")}>{TP},"高於","低於")&"本模型；本模型 12 個月後的前瞻 EPS 比共識"&IF({_V("與市場共識目標價的差距：EPS 因素")}<0,"低 ","高 ")&TEXT(ABS({_V("與市場共識目標價的差距：EPS 因素")}),"0%")&"（兩者時間窗不同：本模型為約 2027-10 起的未來 12 個月 EPS，共識為目前的前瞻 12 個月 EPS，部分差距來自時間差）；與共識目標價的差距主要來自倍數：共識目標價隱含約 "&TEXT({_V("市場共識目標價")}/{_I("cons_feps")},"0.0")&" 倍前瞻 EPS，同業中位數約 "&TEXT({_V("同業前瞻 P/E 中位數")},"0.0")&" 倍。"']:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "本版要點（HTML 附錄用）").font = BOLD; rn += 1
for t_ in [f'="v0.5.2（2026-11-05）第二輪修正（換用共用模組 tools/fcff_dcf.py v2；Andy 2026-11-05 23:16 定案）：(1) 資產負債表維持最新已公布一季 2026-09-30（3Q26），淨現金以 stub 期滾到評價日 2026-12-31＝淨現金與投資 "&TEXT(dcf_nc,"#,##0.0")&"＋stub 期 FCFF（各情境 FY2027 FCFF × "&TEXT(stub_frac,"0.000")&" 年；情境 B "&TEXT(Val!C{SDCF["stub"]},"#,##0.0")&"）− 已宣告股利 "&TEXT(stub_div,"0.0")&"（FY2025 股利已於 7 月發放，該段無已宣告股利）＝"&TEXT(Val!C{SDCF["ncr"]},"#,##0.0")&"（十億新台幣）；(2) β 組更名「TW 散熱／電源」，β_u "&TEXT(beta_u,"0.000")&" 不變，Sens β 低／高改 β_u ∓ 0.3 各自重新加槓桿；(3) 新增 Sens「美元口徑折現」：加權目標價 "&TEXT(Sens!C{USD_ROW},"#,##0")&" 元（"&TEXT(Sens!D{USD_ROW},"+0.0%;-0.0%")&"）；(4) 本模型無逐年資本支出／折舊列，維持占營收比例。DCF 目標價 981 → "&{T_(DCF)}&" 元、加權目標價 1,211 → "&{T_(TP)}&" 元、上檔空間 −9.3% → "&TEXT({UP},"+0.0%;-0.0%")&"、投資建議中立 → "&{REC}&"。"',
           "v0.5.1（2026-11-05）β 改為自算產業 β（Andy 2026-11-05 22:55）：TW 散熱／電源／探針卡同業 6 家無槓桿 β 中位數 1.30，Blume 調整後 1.201，依台達電 D／E 重新加槓桿為 1.23；WACC 8.9%、DCF 目標價約 981 元、加權目標價約 1,211 元、上檔空間 -9.3%、投資建議中立。評價日維持 2026-12-31；其他模型輸入不變。",
           "v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；新增 DCF 頁，Val 0.1 節改為折現率與 DCF 彙總，Sens 目標價參數改為 WACC、β、ERP、g、RONIC、資本支出、營運資金等案例；In 頁新增 M2 節、刪除 tp_r、tp_roe）。Comparable、情境權重、DCF 權重 50% 與投資建議規則不變。",
           "v0.4.1（2026-11-05）只改展示層：README 新增「展示層」「命題回覆」兩段（數字以公式取自各頁），HTML 改用個股 HTML 版面規格 v1.0；README 以外逐格與 v0.4 相同。",
           "v0.4（2026-11-02）：滾入 3Q26 實績（Q 頁 D 欄改讀 In 頁 q3_*）、4Q26 假設改為季增 10%、毛利率 34.5%；股價改 2026-10-30 收盤。"]:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "HTML 橫幅附註用：進加權的 Comparable 目標價（相對倍數調整後）相對現價的上檔空間（橫幅中間格為只看 Comparable 目標價（市場主流同業法），其上檔空間見 Val 頁）")
RD.cell(rn, 2, f"={CMP}/{PX}-1").number_format = "0.0%"
RD["A1"] = RD["A1"].value + "；v0.4.1（2026-11-05）只改展示層：README 新增展示層與命題回覆段、HTML 改用版面規格 v1.0，數值不變；v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重與投資建議規則不變；v0.5.1（2026-11-05）β 改為自算產業 β（TW 散熱／電源同業中位數，Blume 調整；Andy 2026-11-05 22:55），其他輸入不變；v0.5.2（2026-11-05）第二輪修正：換用 tools/fcff_dcf.py v2（淨現金由 2026-09-30 滾到評價日、β 敏感度 β_u ± 0.3、新增 Sens 美元口徑折現）"
_so = SO.max_row + 1   # v0.5.2 第二輪修正：資產負債表日與 stub 期股利的來源網址
for _t, _u in (("[Verified] 第二輪修正：資產負債表（季）2026.3Q 欄（富邦證券 e-Broker；資料商轉載合併報表；與 stockanalysis 2026-09-30 欄同值）", "https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcp/zcpa/zcpa_2308.djhtm"),
               ("[Verified] 第二輪修正：除權息表（FY2025 現金股利 11.6 元，除息 2026-06-17、發放 2026-07-22）", "https://www.wantgoo.com/stock/2308/dividend-policy/ex-dividend")):
    SO.cell(_so, 1, _t); SO.cell(_so, 2, _u); _so += 1
wb.save(sys.argv[1])
if os.environ.get("DELTA_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": SENS_N, "IMP": IMP, "VAL": VAL, "SB": SB},
              open(os.environ["DELTA_ROWMAP"], "w"), ensure_ascii=False, indent=1)
