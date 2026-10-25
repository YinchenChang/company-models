# 台達電（2308）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業樣板）
# 用法：python3 delta/build_delta.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；市場與每 kW 內容以美元輸入、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（公司揭露的「稅後淨利」＝稅前 − 所得稅 − 非控制權益）。
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

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A 基準，B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；本模型以此除以每 GW IT 資本支出得到部署 GW" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合每 GW IT 資本支出用）", "%", 0.15, 0.22, 0.30, "Interested-party", "TrendForce 2026 組合 Rubin 22%（經 markets/memory）", "公司 2Q26 法說：2026 年 Blackwell 出貨遠大於 Rubin")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "Recon 用：每架內容換算")
inp("cool_vr", "VR200 每架機架內液冷（冷板等，橋接表）", "US$K", None, BR["cool"][1], None, "Interested-party", BR["cool"][2], "Morgan Stanley（經 Tom's Hardware 2025-11-05）：VR NVL144 機架內冷卻 BOM US$55.7K；不含 CDU、分歧管")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、NVIDIA、Arista 模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", 30.5, 32.0, 33.0, "Derived", "公司 2Q26 法說換算 NT$183.2B ≈ US$5.6B（隱含約 32.7）", "2026 年初 31.4（〈台幣〉開紅盤，uanalyze）；10-23 收盤：未找到")

sec("B. AI 電源內容與台達電份額（本專案為權威：每架電源與液冷內容、台達電份額，規則第 13 條）")
inp("pkw_26", "AI 電源內容（機架電源櫃、PSU、BBU、HVDC 電源櫃）2026", "US$/kW", 380, 450, 520, "Derived", KGI + "：GB300 每架 PSU US$60–70K ÷ 136kW＝440–515；GB200 35–40K ÷ 120kW＝290–330（Interested-party）", "2026 年以 GB300 為主（公司：Blackwell 遠大於 Rubin）")
inp("pkw_27", "AI 電源內容 2027（VR200）", "US$/kW", 500, 600, 700, "Derived", KGI + "：VR200 每架 PSU 超過 US$150K ÷ 約 250kW＝600 以上", "每 GW 4,052 架（Tokenomics）→ 每架約 247kW，與 250kW 相符")
inp("pkw_28", "AI 電源內容 2028（800V HVDC 放量）", "US$/kW", 550, 680, 800, "Assumed", "公司：±400V、800V 3Q26 量產、2027 才有量；券商估 ±400V、800V 電源櫃占 2027 營收 5–10%（旺得富 2026-10-01，Interested-party）", "800V 把交流轉直流移到機架外電源櫃與 SST：每 kW 內容上升（Assumed）")
inp("pkw_29", "AI 電源內容 2029", "US$/kW", 550, 720, 850, "Assumed", "同上", "—")
inp("rel_27", "台達電 AI 電源份額 2027 ÷ 2026（相對值）", "x", 0.85, 0.95, 1.0, "Assumed", KGI + "：Blackwell 平台 AI 伺服器 PSU 台達電約 70%；光寶、Flex 追趕（Interested-party）", "公司 2Q26 法說：Rubin 110kW PSU「領先」；對手：麥格米特、歐陸通（中國）、Vertiv、Flex（BigGo 摘要）")
inp("rel_28", "台達電 AI 電源份額 2028 ÷ 2026", "x", 0.70, 0.90, 1.0, "Assumed", "同上", "—")
inp("rel_29", "台達電 AI 電源份額 2029 ÷ 2026", "x", 0.65, 0.85, 1.0, "Assumed", "同上", "—")
inp("lcr_27", "台達電液冷每 GW 營收 2027 ÷ 2026", "x", 0.9, 1.10, 1.3, "Assumed", "橋接表：機架內液冷 GB300 US$50K → VR200 US$56K／架（每 kW 下降：架數 6,764 → 4,052）", "公司：液冷占營收 2025 約 10% → 2026 超過 12%；Liquid-to-Air 使用期比市場預期長（2Q26 法說）")
inp("lcr_28", "台達電液冷每 GW 營收 2028 ÷ 2026", "x", 0.85, 1.15, 1.4, "Assumed", "同上", "—")
inp("lcr_29", "台達電液冷每 GW 營收 2029 ÷ 2026", "x", 0.8, 1.20, 1.5, "Assumed", "同上", "—")

sec("C. 營收結構錨點（公司揭露比例 × 營收）")
inp("ai_25", "2025 年 AI 相關營收 ÷ 營收（讀為不含液冷）", "%", None, 0.18, None, "Interested-party", TN1012 + "：AI 相關營收 2025 年 18%、2026 年超過 25%（公司）", "數位時代 2026-02-26：AI 電源超過 20%、液冷約 9%、合計約 30%（口徑不同，未採用）")
inp("ai_26", "2026 年 AI 相關營收 ÷ 營收（讀為不含液冷）", "%", 0.25, 0.26, 0.28, "Interested-party", TR2Q26 + "：CFO「一定超過 25%」；BigGo 摘要：朝 28% 目標", "讀法（Assumed）：AI 相關＝AI 電源，不含液冷；另一讀法見 Sens「AI 相關含液冷」")
inp("lc_25", "2025 年液冷營收 ÷ 營收", "%", None, 0.10, None, "Interested-party", TR2Q26 + "：液冷約占 2025 年營收 10%", "數位時代：約 9%")
inp("lc_26", "2026 年液冷營收 ÷ 營收", "%", 0.12, 0.125, 0.13, "Interested-party", TR2Q26 + "：2026 年超過 12%", TN1012 + "：液冷產品可望超過 12%")
inp("g_nai27", "非 AI 營收年增 2027", "%", 0.0, 0.06, 0.10, "Assumed", "2026 錨點非 AI 約 +15%（Q 頁；含非 AI 資料中心、電源零組件）；自動化 2Q26 +15% 但虧損、交通衰退（2Q26 法說）", "2025：自動化 +5%、交通 −16%（4Q25 法說簡報）")
inp("g_nai28", "非 AI 營收年增 2028", "%", 0.0, 0.05, 0.10, "Assumed", "同上", "—")
inp("g_nai29", "非 AI 營收年增 2029", "%", 0.0, 0.04, 0.08, "Assumed", "同上", "—")

sec("D. 毛利率")
inp("d_ai", "AI 電源毛利率高於其他營收的幅度", "pt", 0.04, 0.08, 0.12, "Assumed", "券商：18kW PSU 與 HVDC 電源櫃毛利率約 50%（旺得富 2026-10-01，Interested-party）；公司整體 1H26 36.3%", "部門別毛利率未揭露：未找到；電源及零組件 2025 部門利益率 20.1% vs 基礎設施 15.6%（4Q25 法說簡報）")

sec("E. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收合計）", "NT$B", None, 197.059, None, "Verified", TN1012 + "：3Q26 1,970.59 億元、前三季 5,396.68 億元", "時報資訊 2026-10-12 同值（含更正稿）")
inp("q4g", "4Q26 營收季增", "%", 0.05, 0.10, 0.20, "Interested-party", "法人：第四季雙位數季增（TechNews 2026-10-12）", "券商 2026-10-01：4Q26 2,511 億（季增 20%，以 3Q 2,093 億為基礎；3Q 實際 1,971 億）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.35, 0.36, 0.38, "Interested-party", TR2Q26 + "：董事長「35% 左右合理」、下半年優於上半年；2Q26 35.6%", "券商 2026-10-01：4Q26 38.7%（新高）")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.175, 0.185, 0.195, "Assumed", "1H26 19.1%、4Q25 18.3%（公司）", "研發 +33%、銷管 +40%（2Q26 年增）")
inp("q_nop", "3Q、4Q26 業外收益（每季）", "NT$B", 0.0, 2.0, 4.4, "Derived", "2Q26 4.4（部分為未實現投資利益）、4Q25 −0.8（公司）", "FY2025 3.9（公司）")
inp("q_nir", "3Q、4Q26 歸屬母公司淨利 ÷ 稅前", "%", 0.67, 0.69, 0.71, "Derived", "1H26 45.7 ÷ 66.1＝0.691；FY2025 60.1 ÷ 87.9＝0.684（公司）", "含稅率約 22% 與非控制權益（台達泰國等）")
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
inp("ubs_26", "UBS：全球 AI 伺服器電源市場 2026", "US$B", None, 17.7, None, "Interested-party", "TechNews 2026-10-12（引述 UBS；2027 38.2、2025–2030 年複合成長約 69%；網址見 Sources）", "口徑（是否含 BBU、HVDC、非 NVIDIA）：未找到")
inp("ubs_27", "UBS：全球 AI 伺服器電源市場 2027", "US$B", None, 38.2, None, "Interested-party", "同上", "—")
inp("kgi_sh", "凱基：Blackwell 平台 AI 伺服器 PSU 台達電份額", "%", None, 0.70, None, "Interested-party", "凱基 Power supply sector（2025-07-01；網址見 Sources）", "第二個量化來源：未找到")
inp("pe_fair_andy", "Andy 先驗：合理 P/E", "x", None, 25, None, "Assumed", "先驗_Andy_台達電_20261024 D10b（Drive；不在本 repo）", "只作對照")

sec("H. 價格與估值")
inp("px", "股價（2026-10-23 收盤）", "NT$", None, 1605, None, "Interested-party", "Yahoo 股市 2308（10-24 頁面「昨收」1,605）", "HiStock、Goodinfo：10-22 收盤 1,560（−4.0%）；兩週 −20.6%")
inp("pe_ttm", "近 12 月 P/E（10-22）", "x", None, 49.65, None, "Interested-party", "Goodinfo（10-22）", "HiStock 49.65")

sec("I. 目標價情境權重（Val 第 0 節用；Andy 可改）")
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
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自台達電法說簡報、逐字稿與月營收"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（月營收）＋3Q、4Q 假設（毛利率、營業費用率、業外、淨利 ÷ 稅前；4Q 營收＝3Q × (1＋季增)）。1H26 營業利益＝2Q26 30.6＋1Q26（2Q26 ÷ 1.08；公司稱季增 8%）。部門別損益只有年度簡報數字，季度部門數字只在簡報畫面：未找到"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
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
qrow(5, "毛利", {"B": 190.2, "C": 124.3, "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Interested-party")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 34.3%；1H26 36.3%")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived（毛利 − 營業利益）／Assumed")
qrow(8, "營業利益", {"B": 83.9, "C": "=30.6+30.6/1.08", "D": "=D5-D7", "E": "=E5-E7"}, tag="Verified／Derived")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 15.1%；1H26 17.2%（公司）")
qrow(10, "業外收益", {"B": 3.934, "C": "=C11-C8", "D": "=q_nop", "E": "=q_nop"}, tag="Verified／Derived")
qrow(11, "稅前利益", {"B": "=B8+B10", "C": 66.1, "D": "=D8+D10", "E": "=E8+E10"}, tag="Verified／Derived", s2="FY2025 公司揭露 87.9（四捨五入差）")
qrow(12, "歸屬母公司淨利", {"B": 60.108, "C": 45.7, "D": "=D11*q_nir", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）")
qrow(13, "淨利 ÷ 稅前（含稅與非控制權益）", {c: f"={c}12/{c}11" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股）", {"B": "=60.108/23.14", "C": "=q_sh", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s2="淨利 ÷ EPS")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(16, "  公司揭露 EPS", {"B": 23.14, "C": 17.59}, fmt="0.00", ann="", tag="Verified", s2="2Q26 9.68")
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
DEFP = dict(scn="A", rel="rel", pkw="pkw", lcr="lcr", mf="1", gn="g_nai", d="d_ai", opxa="0", aip="aip", nai="nai")
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
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU})
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
    row("nai", "非 AI 營收", "NT$B", {"C": f"=a25_{P['nai']}", "D": f"=a26_{P['nai']}", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']}{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['aip']}+{c_(y)}{R['lc']}+{c_(y)}{R['nai']}" for y in FYS}, fmt="#,##0.0")
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
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("aip", "  AI 電源營收", "NT$B", "#,##0.0"), ("lc", "  液冷營收", "NT$B", "#,##0.0"),
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
assert sr < 60, sr
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
V.cell(6, 1, "現價（2026-10-23 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "#,##0"
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
V.cell(TP0, 1, "0. 目標價與上檔空間（12 個月，情境加權；權重為 In 頁 tp_wA／tp_wB／tp_wC，Andy 可改）").font = BOLD
for k_, (key_, lab_) in enumerate((("tp", "目標價（12 個月，情境加權）"), ("up", "上檔空間（目標價 ÷ 現價 − 1）"),
        ("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"),
        ("w", "情境權重 A／B／C"), ("m", "目標價方法")), TP0 + 1):
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
defname("tp_price", f"Val!$B${TPR['tp']}"); defname("tp_upside", f"Val!$B${TPR['up']}")
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
         ("非 AI 成長 低", "2027–29 0%", dict(gn="GNLO")), ("非 AI 成長 高", "2027 +10%、2028 +10%、2029 +8%", dict(gn="GNHI")),
         ("AI 毛利率差距 4pt", "AI 電源毛利率只高 4 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 12pt", "高 12 個百分點", dict(d="d_ai_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi")),
         ("AI 相關含液冷（讀法）", "FY2026 AI 電源＝營收 ×（26% − 12.5%）；份額分子變小", dict(aip="aipx", nai="naix"))]
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("PKLO", "pkw", "_lo"), ("PKHI", "pkw", "_hi"), ("LCLO", "lcr", "_lo"), ("LCHI", "lcr", "_hi"), ("GNLO", "g_nai", "_lo"), ("GNHI", "g_nai", "_hi")):
    for y in (26, 27, 28, 29):
        if base in ("rel", "lcr", "g_nai") and y == 26: continue
        src = f"{base}{y}{suf}" if base == "g_nai" else f"{base}_{y}{suf}"
        target = f"{nm}{y}" if base == "g_nai" else f"{nm}_{y}"
        wb.defined_names[target] = DefinedName(target, attr_text=wb.defined_names[src].attr_text)
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
lines = [("台達電（2308）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業樣板）　2026-10-25", True),
 ("版本：v0.1.1（2026-10-25：Val 新增第 0 節「目標價與上檔空間」、In 新增情境權重 tp_wA／tp_wB／tp_wC；其餘數值與 v0.1 相同）。v0.1（初版）：AI 電源每 kW 內容、台達電份額、液冷每 GW 營收在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）；奇鋐等散熱廠之後讀這條液冷線。", False),
 ("模型命題：台達電的 AI 營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）×（AI 電源每 kW 內容 × 台達電份額＋液冷每 GW 營收）。800V HVDC 把交流轉直流移到機架外，每 GW 可做的金額變大；同時光寶、Flex、中國廠追趕份額。股價兩週跌約兩成後，約 22–28 倍未來 12 月 EPS。", False),
 ("核心問題：每 kW 電源內容上升（800V）能不能抵銷份額流失與零組件成本，讓 AI 營收在 2027–28 年維持高成長；現價隱含的路徑合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029（三年）。金額：十億新台幣；市場以美元計、以匯率換算。口徑：歸屬母公司淨利。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 電源營收：markets/network 的 AI IT 資本支出 × 情境顆數係數 ÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 kW 電源內容（凱基推導、2028 起 Assumed）× 匯率 × 台達電有效份額（FY2026 以公司 AI 相關占比校準）× 份額相對值（Assumed）。", False),
 ("液冷營收：FY2026 以公司液冷占比校準每 GW 營收 × 相對值（Assumed）。非 AI：FY2026 殘差 × 成長率（Assumed）。毛利率：AI 電源比其他營收高 d_ai（Assumed），其他毛利率為 FY2026 殘差（Derived）。營業費用：彈性 phi（Derived）。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績＋假設（4Q 季增、毛利率、營業費用率）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：現價 ÷ FY2027、FY2028 EPS；十年 P/E 規則值（單一資料商來源）× FY2027 EPS；循環位置以本模型營收最高年判定（台達電不套設備股見頂年）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文＝法說簡報與逐字稿；券商皆未具名或單一來源。", False),
 ("", False),
 ("已知限制：(1)「AI 相關」是否含液冷，公司說法不一致（Sens「AI 相關含液冷」）；(2) 2028 年起每 kW 內容（800V）與份額相對值為 Assumed，是 FY2028 EPS 最敏感的輸入；(3) 部署 GW 以 VR200 每 GW IT 資本支出換算，ASIC 機架的每 GW 資本支出較低，GW 可能低估；(4) 十年 P/E 只有一個來源，且 2016–2019 年為 AI 前的低倍數；(5) 未計 Vicor 337 調查（垂直供電）與零組件短缺的下檔；(6) 3Q26 法說（10 月底）後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("DELTA_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": SENS_N, "IMP": IMP, "VAL": VAL, "SB": SB},
              open(os.environ["DELTA_ROWMAP"], "w"), ensure_ascii=False, indent=1)
