# 雙鴻（3324）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業，樣板＝avc/ 奇鋐）
# 用法：python3 auras/build_auras.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；市場與每架內容以美元輸入、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（合併淨利 − 非控制權益）。
# 情境與行業一致（Andy 2026-10-25 13:13）：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS）；
# 液冷滲透率（技術替代：氣冷 → 液冷）屬行業情境，A＝替代快（奇鋐模型 In 頁高值）、B＝基準、C＝慢（低值），FY2026 三情境相同；
# 公司特有變數（雙鴻份額相對值）A 取高值、B 基準、C 低值。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（液冷行業變數 v0.6.1 起單一來源）
import cycle_position as CP
import valuation_rule as VR
import fcff_dcf as FD   # v0.2：標準版 DCF（無槓桿自由現金流＋WACC）共用模組
import market as MEM

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN_T = {"A": "情境 A（順利）：顆數係數 A＋液冷滲透率高值＋份額相對值高值",
         "B": "情境 B（平平）：顆數係數 B（CY2028 ×0.8、CY2029 ×0.85）＋基準輸入",
         "C": "情境 C（不如預期）：顆數係數 C（CY2028–29 ×0.6）＋液冷滲透率低值＋份額相對值低值"}

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); BR_F = only("markets/chip_bridge/dist/*.xlsx")
AVC_F = only("avc/dist/*.xlsx")
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
for s_ in SC:   # 情境與行業一致：tsmc 顆數係數必須等於 markets/memory 的 SCEN_UNITS（同一字母）
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9, s_
BR = {}
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("racks", "capit", "cool"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 3, BR.keys()
AW = openpyxl.load_workbook(AVC_F, data_only=True)
AV = {}   # 奇鋐模型 In 頁的液冷內容、滲透率、份額相對值（低／基準／高）；兩家散熱廠用同一液冷市場口徑
for r in AW["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith(("vr26", "fx")):
        AV[k] = (r[3].value, r[4].value, r[5].value, f"{rel(AVC_F)}｜In!D{r[0].row}:F{r[0].row}")
assert all(k in AV for k in ("vr26", "fx")), AV.keys()
# 液冷內容、滲透率、份額相對值：v0.6.1 起讀 Siliconomics M19（與奇鋐模型同一值）
for k, sk in ([("mani", "lc_mani")] + [(f"lcadj_{y}", f"lc_adj_{y}") for y in (27, 28, 29)] + [(f"lpen_{y}", f"lc_pen_{y}") for y in (26, 27, 28, 29)]
              + [(f"rel_{y}", f"lc_avrel_{y}") for y in (27, 28, 29)]):
    AV[k] = (*SIL.vals(sk), f"siliconomics｜SIF_{sk}")
AVSH = None
for r in AW["Units"].iter_rows(min_row=4):
    if r[0].value == "奇鋐有效份額":
        AVSH = ([r[2 + j].value for j in range(4)], f"{rel(AVC_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert AVSH and all(v is not None for v in AVSH[0]), AVSH

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），市場與每架內容為美元；比例：%。每個輸入盡量附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝雙鴻財報與月營收（經 HiStock、玩股網、時報轉載）；公司的產品比重說法為 Interested-party；券商估計皆為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
KGI08 = "https://www.kgi.com.hk/zh-cn/-/media/files/kgishk/research-reports/tw-reports/auras-technology-3324-tt20260807en.pdf"
WR03 = "https://wantrich.chinatimes.com/news/20260304900286-420101"
WR05 = "https://wantrich.chinatimes.com/news/20260521900232-420101"
WG = "https://www.wantgoo.com/stock/3324/financial-statements/income-statement"
HIS = "https://histock.tw/stock/financial.aspx?no=3324"
HISEPS = "https://histock.tw/stock/3324/每股盈餘"
TN09 = "https://finance.technews.tw/2026/10/08/auras-3324-202609-financial-report/"
WV = "https://winvest.tw/News/Detail/95436"
UDN0918 = "https://money.udn.com/money/story/5607/9762508"
UDN0919 = "https://money.udn.com/money/story/5607/9764013"
SAF = "https://stockanalysis.com/quote/tpex/3324/forecast/"
SAQ = "https://stockanalysis.com/quote/tpex/3324/"
SAR = "https://stockanalysis.com/quote/tpex/3324/financials/ratios/"
GI = "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=3324"
YH = "https://tw.stock.yahoo.com/quote/3324.TWO"

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A（原始顆數），B／C 依同字母顆數係數縮放；與奇鋐模型相同上游）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；本模型以此除以每 GW IT 資本支出得到部署 GW" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合每 GW 數值用）", "%", AV["vr26"][0], AV["vr26"][1], AV["vr26"][2], "Interested-party", AV["vr26"][3], "TrendForce 2026 組合 Rubin 22%（經 markets/memory；與奇鋐、台達電模型相同）")
inp("racks_gb", "每 GW 機架數（GB300）", "架", None, BR["racks"][0], None, "Derived", BR["racks"][2], "Tokenomics（經橋接表）")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "同上")
inp("cool_gb", "GB300 每架機架內液冷（運算匣＋交換匣，橋接表）", "US$K", None, BR["cool"][0], None, "Interested-party", BR["cool"][2], "Morgan Stanley（經 Tom's Hardware 2025-11-05）：GB300 NVL72 冷卻 BOM；不含 CDU、分歧管")
inp("cool_vr", "VR200 每架機架內液冷（橋接表）", "US$K", None, BR["cool"][1], None, "Interested-party", BR["cool"][2], "同上：VR NVL144")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", AV["fx"][0], AV["fx"][1], AV["fx"][2], "Derived", AV["fx"][3], "與奇鋐、台達電模型相同")

sec("B. 液冷內容、滲透率（讀奇鋐模型 In 頁：兩家用同一液冷市場口徑）與雙鴻份額相對值")
inp("mani", "每架分歧管與快接頭（橋接表未含）", "US$K", *AV["mani"][:3], "Assumed", AV["mani"][3], "Siliconomics M19（與奇鋐模型同一值）；雙鴻另做機架分歧管（Vera Rubin inner／rack manifold，經濟日報 2026-09-18）")
for y in (27, 28, 29):
    inp(f"lcadj_{y}", f"每架液冷內容 ÷ 橋接表 20{y}（Rubin 全液冷、ASIC 液冷；2026＝1）", "x", *AV[f"lcadj_{y}"][:3], "Assumed", AV[f"lcadj_{y}"][3],
        "Siliconomics M19（與奇鋐模型同一值）：TrendForce 稱 Vera Rubin 全液冷提高單架散熱價值（Interested-party）；本模型三情境用基準值，區間只在 Sens" if y == 27 else "同上")
for y in (26, 27, 28, 29):
    inp(f"lpen_{y}", f"液冷滲透率 20{y}（部署 GW 中液冷占比）", "%", *AV[f"lpen_{y}"][:3], "Assumed", AV[f"lpen_{y}"][3],
        "Siliconomics M19（與奇鋐模型同一值；TrendForce 以顆數計 2026 53%、2027 近 60%）。技術替代屬行業情境：FY2027 起 A 取高值、B 基準、C 低值；FY2026 三情境皆基準" if y == 26 else "同上")
for y in (27, 28, 29):
    inp(f"rel_{y}", f"雙鴻液冷份額 20{y} ÷ 2026（相對值）", "x", *AV[f"rel_{y}"][:3], "Assumed", AV[f"rel_{y}"][3] + "（區間與奇鋐相同）",
        "公司特有：A 取高值（三個 ASIC 液冷專案 3Q／4Q26 量產、Vera Rubin 分歧管）、B 基準（與奇鋐相同的份額流失斜率：酷碼、台達電、富世達等追趕）、C 低值；凱基 2026-08-07：雙鴻為台灣前三大散熱廠（未給份額，Interested-party）；份額數字：未找到" if y == 27 else "同上")

sec("C. 營收結構錨點（公司揭露比例 × 營收）")
inp("srv_25", "2025 年伺服器 ÷ 營收", "%", None, 0.57, None, "Interested-party", WR05 + "：2025 年伺服器 57%、PC 27%、VGA 14%、其他 2%（公司，股東會）", KGI08 + "：2025 年伺服器 57%、筆電與桌機 27%、電競筆電顯卡主機板 14%、其他 2%")
inp("nb_25", "2025 年筆電與桌機 ÷ 營收", "%", None, 0.27, None, "Interested-party", "同上（公司 PC 27%）", "同上（凱基 27%）")
inp("lc_25", "2025 年液冷 ÷ 營收", "%", 0.30, 0.34, 0.38, "Interested-party", "凱基 2026-08-07（網址見 Sources）：2025 年液冷占營收 34%", "第二來源：未找到")
inp("srv_26", "2026 年伺服器 ÷ 營收", "%", 0.72, 0.74, 0.76, "Interested-party", KGI08 + "：2026F 伺服器 74%（2Q26 78%）", "公司：2026 年伺服器營收年增 110–120%（凱基轉述）→ 2025 年 57% × 2.15 ÷ 營收年增約 64% ≈ 75%（Derived）")
inp("nb_26", "2026 年筆電與桌機 ÷ 營收", "%", 0.13, 0.15, 0.17, "Interested-party", "凱基 2026-08-07（網址見 Sources）：2026F 筆電與桌機 15%（2Q26 17%）", "第二來源：未找到")
inp("lc_26", "2026 年液冷 ÷ 營收", "%", 0.55, 0.60, 0.63, "Interested-party", KGI08 + "：1H26 液冷 55%、2026F 60%", "公司：2026 年液冷營收年增 190–200%（凱基轉述）→ 2025 年 34% × 2.95 ÷ 1.64 ≈ 61%（Derived）")
for y, lo, b, hi in ((27, -0.20, -0.05, 0.05), (28, -0.20, -0.05, 0.05), (29, -0.20, -0.05, 0.05)):
    inp(f"g_sair_{y}", f"伺服器氣冷營收年增 20{y}（含 AI 氣冷與一般伺服器）", "%", lo, b, hi, "Assumed",
        "氣冷轉液冷；凱基 2027F 伺服器 82%、液冷 72% → 伺服器氣冷約 10%（2026F 14%），隱含約 −7%" if y == 27 else "同上", "非 AI 另列（工作單）：不隨部署 GW" if y == 27 else "—")
for y, lo, b, hi in ((27, -0.10, -0.03, 0.05), (28, -0.08, 0.0, 0.05), (29, -0.08, 0.0, 0.05)):
    inp(f"g_nb_{y}", f"筆電與桌機營收年增 20{y}", "%", lo, b, hi, "Assumed", "凱基 2027F 筆電與桌機 11%（2026F 15%），隱含約 −5%" if y == 27 else "同上", "—")
for y, lo, b, hi in ((27, -0.15, 0.0, 0.08), (28, -0.10, 0.0, 0.08), (29, -0.10, 0.0, 0.08)):
    inp(f"g_oth_{y}", f"電競筆電、顯卡、主機板與其他營收年增 20{y}", "%", lo, b, hi, "Assumed", "凱基 2027F 電競與顯卡 6%、其他 1%（2026F 8%、3%）" if y == 27 else "同上", "—")

sec("D. 毛利率")
inp("d_lc", "液冷毛利率高於其他營收的幅度", "pt", 0.02, 0.05, 0.08, "Assumed", "毛利率 FY2025 27.4% → 1H26 29.0%，同期液冷占比 34% → 55%（隱含差距約 7–8pt，含泰國廠試產成本等其他因素）", "產品別毛利率：未找到")

sec("E. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（7–9 月月營收合計）", "NT$B", None, "=3.765466+3.141249+2.884061", None, "Verified", HIS + "：7 月 37.65、8 月 31.41、9 月 28.84 億元", TN09 + "：前 9 月 270.40 億元（− 1H26 172.52 ≈ 97.88）")
inp("q4g", "4Q26 營收季增", "%", 0.08, 0.14, 0.18, "Interested-party", UDN0919 + "：花旗估 3Q 季增 17%、4Q 季增 16%（Vera Rubin 出貨、ASIC 液冷量產）", KGI08 + "：4Q26 114.98 億元（以 3Q 102.66 億為基礎，季增 12%；3Q 實際 97.9 億）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.28, 0.29, 0.30, "Interested-party", KGI08 + "：3Q26 毛利率 29.0%", "2Q26 28.4%、1Q26 29.7%（公司）；花旗：3Q、4Q 毛利率回升（經濟日報 2026-09-19）")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.10, 0.11, 0.122, "Interested-party", KGI08 + "：3Q26 營益率 18.0%、毛利率 29.0%（營業費用率 11.0%）", "1Q26 11.9%、2Q26 12.4%（公司；毛利 − 營業利益）")
inp("q_nop", "3Q、4Q26 業外收益（每季）", "NT$B", -0.15, 0.0, 0.15, "Derived", "1Q26 +0.05、2Q26 −0.07（稅前 − 營業利益）", WG)
inp("q_nir", "3Q、4Q26 歸屬母公司淨利 ÷ 稅前", "%", 0.72, 0.74, 0.76, "Derived", "1Q26 11.64 ÷ 15.70＝0.741；2Q26 9.69 ÷ 13.17＝0.736", "含稅率約 19–21% 與非控制權益")
inp("q_sh", "股數（十億股；3Q 起與 FY2026 EPS 換算）", "十億股", None, 0.093076837, None, "Verified", "HiStock：已發行普通股 93,076,837 股", SAQ + "：Shares outstanding 93.08M；1H26 淨利 21.33 億 ÷ EPS 23.02 ＝ 加權 0.0927")

sec("F. 營業費用、業外、稅與非控制權益、股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.5, 0.65, 0.85, "Derived", "FY2026 錨點實測（Q 頁 B26）", "FY2025 營業費用年增 47%、營收年增 48%（彈性約 1.0，泰國廠擴產期）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("nop_f", "業外收益（每年）", "NT$B", -0.5, 0.0, 0.5, "Assumed", "FY2025 約 0.02（2Q25 匯損 −0.42）、1H26 約 −0.02", "匯兌損益波動大")
inp("nir_f", "歸屬母公司淨利 ÷ 稅前", "%", 0.72, 0.74, 0.78, "Derived", "同 q_nir；FY2025 25.72 ÷ 32.77＝0.785（2Q25 稅負低）", "非控制權益比例變動：未建模")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.01, "Assumed", "FY2025 加權 0.0910 → 2026 年已發行 0.0931 十億股", "可轉債、員工認股：未找到")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("G. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("b_rev26", "凱基 FY2026 營收（2026-08-07）", "NT$B", None, 39.016, None, "Interested-party", KGI08, "—")
inp("b_eps26", "凱基 FY2026 EPS（2026-08-07）", "NT$", None, 55.69, None, "Interested-party", KGI08, "—")
inp("b_rev27", "凱基 FY2027 營收（2026-08-07）", "NT$B", None, 50.486, None, "Interested-party", KGI08, "—")
inp("b_eps27", "凱基 FY2027 EPS（2026-08-07）", "NT$", None, 81.55, None, "Interested-party", KGI08, "目標價 1,410（22 倍 2026–27 平均 EPS）")
inp("b_lc27", "凱基：2027 年液冷 ÷ 營收", "%", None, 0.72, None, "Interested-party", KGI08, "內文 70–75%")
inp("u_eps26", "美系券商 FY2026 EPS（2026-09-18）", "NT$", None, 58.8, None, "Interested-party", UDN0918 + "：美系券商 2026 58.8、2027 89、2028 110（20 倍 2027 EPS）", "券商名稱未載明；花旗 09-19 目標價 1,800（經濟日報）")
inp("u_eps27", "美系券商 FY2027 EPS（2026-09-18）", "NT$", None, 89, None, "Interested-party", UDN0918, "—")
inp("u_eps28", "美系券商 FY2028 EPS（2026-09-18）", "NT$", None, 110, None, "Interested-party", UDN0918, "—")
inp("c_rev26", "共識 FY2026 營收（stockanalysis／S&P Global，2026-10-19）", "NT$B", None, 38.7, None, "Interested-party", SAF, "區間 37.5–39.9")
inp("c_eps26", "共識 FY2026 EPS（同上）", "NT$", None, 56.18, None, "Interested-party", SAF, "區間 52.19–58.82")
inp("c_rev27", "共識 FY2027 營收（同上，Revenue Next Year）", "NT$B", None, 49.10, None, "Interested-party", SAF, "—")
inp("c_eps27", "共識 FY2027 EPS（同上，EPS Next Year）", "NT$", None, 77.97, None, "Interested-party", "stockanalysis forecast 頁（網址見 Sources）", "FY2028 共識：未找到（付費）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"avsh_{y}", f"奇鋐有效份額 20{y}（奇鋐模型 Units，情境不變；Recon 用）", "%", None, AVSH[0][j], None, "Derived", AVSH[1],
        "奇鋐模型三情境的份額相對值相同；只作兩家份額合計檢查，不進本模型計算" if y == 26 else "同上")

sec("H. 價格與估值")
inp("px", "股價（2026-11-04 收盤）", "NT$", None, 1300, None, "Verified", "Yahoo Finance 日線收盤 2026-11-04（第三輪修正：全批統一價格日）", "markets/invest_map/data/prices.csv 2026-11-04 同值 1,300（statementdog；cmoney 昨收 1,320 一致）；v0.2.2 為 2026-10-22 收盤 1,360（Yahoo 股市）")
inp("eps_ttm", "近 12 月 EPS（3Q25–2Q26）", "NT$", None, "=10.66+10.19+12.61+10.41", None, "Verified", HISEPS + "：3Q25 10.66、4Q25 10.19、1Q26 12.61、2Q26 10.41", "—")

sec("I. 目標價（DCF 與 Comparable 加權）與收斂法參數（黃底＝Assumed，Andy 可改）")
for s_, w_, nm_ in (("A", 0.3, "順利"), ("B", 0.5, "平平"), ("C", 0.2, "不如預期")):
    inp(f"tp_w{s_}", f"目標價情境權重：情境 {s_}（{nm_}）", "權重", None, w_, None, "Assumed",
        "Andy 可改（2026-10-25 目標價規格預設 A／B／C＝0.3／0.5／0.2）", "三個權重合計應為 1；用於 DCF 與 Comparable 的情境加權（Val 第 0 節）")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.2 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.2 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑（fade_gA／B／C）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 0.5：DCF／Comparable 各半）", "Sens 頁 0.3／0.7")
inp("fade_gA", "情境 A（順利）：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂，FY2032 達到")
inp("fade_gB", "情境 B（平平）：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C（不如預期）：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("nim10", "歷史平均淨利率（2017–2025 九年；2016 年未找到）", "%", None, "=AVERAGE(0.0424,0.0302,0.0946,0.0923,0.0809,0.0929,0.0975,0.122,0.116)", None, "Derived",
    "Goodinfo（網址見 Sources）：稅後淨利率 2017 4.24%、2018 3.02%、2019 9.46%、2020 9.23%、2021 8.09%、2022 9.29%、2023 9.75%、2024 12.2%、2025 11.6%（合併淨利口徑）",
    "2016 年：未找到（以九年平均代替十年平均）；本模型淨利為歸屬母公司，FY2025 歸母淨利率 11.05%，合併口徑略高")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 1523, None, "Interested-party", "stockanalysis forecast 頁（網址見 Sources）：17 位分析師平均 1,523（中位數 1,550，區間 1,025–2,000），頁面時間 2026-10-19", "第二來源：未找到（凱基 1,410、花旗 1,800 為個別券商）")
inp("cons_feps", "共識前瞻 12 月 EPS（＝2／12 × FY2026＋10／12 × FY2027 共識）", "NT$", None, "=2/12*c_eps26+10/12*c_eps27", None, "Derived",
    "stockanalysis 共識 FY2026 56.18、FY2027 77.97（2026-10-19；Interested-party）", "期間為 2026-10 起算的 12 個月，本模型前瞻 EPS 為 12 個月後起算，兩者期間不同；個股頁 Forward PE 17.27（價格 1,140，2026-08-28）隱含 66.0")

# v0.2：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。評價日沿用 DCF 基準日 2026-12-31；股價 2026-10-22；資產負債為 2026-06-30
BETA_SRC = "自算產業 β 表（2026-11-05 22:50 計算；Yahoo Finance 含股利調整月收盤價 2021-11～2026-10，60 個月；Andy 2026-11-05 22:55 選定，取代 Damodaran 表）"
DAMO_C = "Damodaran「Country Default Spreads and Risk Premiums」2026-01-05 版（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）"
SA_BS = "stockanalysis.com 雙鴻（TPEX:3324）資產負債表，2026-06-30 欄（https://stockanalysis.com/quote/tpex/3324/financials/balance-sheet/；2026-11-05 讀取；資料商轉載合併報表）"
SA_CF = "stockanalysis.com 雙鴻現金流量表（https://stockanalysis.com/quote/tpex/3324/financials/cash-flow-statement/；2026-11-05 讀取）"
SA_IS = "stockanalysis.com 雙鴻財務摘要（https://stockanalysis.com/quote/tpex/3324/financials/；2026-11-05 讀取）"
NC_BS = "4.676+0.95021-(1.976+0.15725+6.498)-(0.1419+0.18019)-0.58824"
NC_NOTE = "2026-06-30 資產負債表（NT$B）：現金及短期投資 4.676＋長期投資 0.950 − 借款（1.976＋0.157＋6.498）− 租賃（0.142＋0.180）− 非控制權益 0.588"
inp("beta_u", "產業無槓桿 β：自算 TW 散熱／電源同業中位數，Blume 調整（0.67β＋0.33）", "x", "=beta_u-beta_du", "=0.67*1.30+0.33", "=beta_u+beta_du", "Derived",
    BETA_SRC + "：TW 散熱／電源同業 6 家（3017 3324 3653 2308 2301 6223），對台灣加權指數 ^TWII 迴歸、Hamada 去槓桿、現金修正後中位數 1.30；Blume 調整 0.67 × 1.30＋0.33＝1.201",
    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3）")
for _row in FD.in_rows(dict(
        date="2026-12-31（同 DCF 基準日）", ccy="NT$", amt="NT$B", shu="十億股",
        rf_us=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
               "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）；與 NVIDIA 模型同值。只用於 Sens「美元口徑折現」（Andy 2026-11-05 23:16：基準維持當地利率＋固定匯率）"),
        stub=dict(bs_date="2026-06-30", val_date="2026-12-31", bs_src=SA_BS + "；富邦證券 e-Broker 資產負債表（季）2026.2Q 欄同值（2026-11-05 讀取；3Q26 季報尚未公布；網址見 Sources 頁）",
                  div=("=11.903*q_sh", "Verified", "FY2025 現金股利每股 11.903 元（除息 2026-06-17、發放 2026-07-16；玩股網除權息表，2026-11-05 讀取；網址見 Sources 頁）× 已發行股數 q_sh", "發放日落在資產負債表日 2026-06-30 之後、評價日之前（6-30 帳上為應付股利），從淨現金扣除；買回無已宣告金額，不扣")),
        rf=(0.0198, "Trading Economics「Taiwan 10Y Bond Yield」：2026-11-04 1.98%（https://tradingeconomics.com/taiwan/government-bond-yield；2026-11-05 讀取）",
            "經濟日報 2026-08-21（https://money.udn.com/money/story/5613/9704659）：Bloomberg 經濟學家預期下半年接近 2%；與台達電模型同值", "無風險利率（台灣 10 年期公債殖利率，2026-11-04）"),
        beta_fixed=("=beta_u_lo*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "=beta_u*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "=beta_u_hi*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))", "Derived",
                    "自算產業 β：TW 散熱／電源 同業 6 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對台灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：β＝beta_u ×（1＋(1 − tax_w) × D ÷ E），D＝債務＋租賃、E＝現價 × 股數）",
                    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。Sens 低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）"),
        crp=(0.0078, "Verified", DAMO_C + "：Taiwan（Moody's Aa3）Country Risk Premium 0.78%", "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋國家風險溢酬、不設股權成本下限"),
        kd=("=rf+0.01", "Assumed", "台灣 10 年期公債＋1 個百分點（公司借款利率未揭露：未找到）", "負債占資本約 7%"),
        tax_w=("=1-(1.232/(12.713*0.1233)+1.893/(15.779*0.1499)+2.572/(23.276*0.1408))/3", "Derived",
               SA_IS + "：近三年 1 −（淨利 ÷ 稅前）平均：稅前＝營收 × 稅前淨利率（FY2023 12,713 × 12.33%、FY2024 15,779 × 14.99%、FY2025 23,276 × 14.08%）、淨利 1,232／1,893／2,572（NT$M）",
               "所得稅費用直接數字未取得；非控制權益小（2026-06-30 帳面 5.9 億元），以 1 − 淨利 ÷ 稅前近似有效稅率；NOPAT 各年亦用本稅率"),
        debt=("=1.976+0.15725+6.498", "Verified", SA_BS + "：短期借款 1,976＋一年內到期長期負債 157.25＋長期負債 6,498（NT$M）", "2025-12-31：2,578＋99.39＋2,021（同站）"),
        lease=("=0.1419+0.18019", "Verified", SA_BS + "：租賃負債流動 141.9＋非流動 180.19（NT$M）", "2025-12-31：72.82＋78.83（同站）"),
        cash=(4.676, "Verified", SA_BS + "：現金及短期投資 4,676（NT$M）", "2025-12-31：1,539（同站）"),
        inv=(0.95021, "Verified", SA_BS + "：長期投資 950.21（NT$M）", "帳面計入評價；Sens 測不計"),
        mi=(0.58824, "Verified", SA_BS + "：非控制權益 588.24（NT$M）", "—"),
        dps=("=a26_eps*12/27.6", "Derived", "評價日 2026-12-31 到目標時點 2027-12-31 之間發放 FY2026 盈餘的現金股利：假設發放率同 FY2025（每股 12.0 元 ÷ EPS 27.60 元＝43.5%）× FY2026 錨點 EPS",
             SA_IS + "：每股股利 FY2023 6.35、FY2024 10.0、FY2025 12.0 元"),
        sh=("=q_sh", "Derived", "＝In 頁 q_sh（已發行 93,076,837 股）", "不另推未來股數"),
        fc_da=(None, "=(0.44039/12.713+0.47427/15.779+0.6266/23.276)/3", None, "Derived", SA_CF + "：折舊攤銷 FY2023 440.39、FY2024 474.27、FY2025 626.6（NT$M）÷ 營收 12,713／15,779／23,276，近三年平均 3.05%",
               "近 12 月（2026-06）780.96"),
        fc_cx=(None, "=(0.93489/12.713+1.436/15.779+2.234/23.276)/3", None, "Derived", SA_CF + "：資本支出 FY2023 934.89、FY2024 1,436、FY2025 2,234（NT$M）÷ 營收，近三年平均 8.68%（泰國、越南擴產）",
               "近 12 月（2026-06）3,076；Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=fc_da_ss*(1+5*g_tv)", "Assumed", "長期穩態（FY2032）資本支出 ÷ 營收＝折舊 ÷ 營收 ×（1＋5 × g）：設備平均年限約 10 年、以 g 成長時，資本支出 ÷ 折舊約 1＋g × 年限 ÷ 2", "滿足規格「資本支出 ≥ 折舊 ×（1＋g）」；FY2029 起由近期比例線性收斂"),
        fc_wc=(0.25, "=((4.419+2.307-3.157)/12.713+(5.429+2.81-3.775)/15.779+(10.068+6.057-7.594)/23.276)/3", 0.37, "Derived",
               SA_BS.replace("，2026-06-30 欄", "") + "：營運資金＝應收帳款＋存貨 − 應付帳款；FY2023 3,569、FY2024 4,464、FY2025 8,531（NT$M）÷ 營收＝28.1%、28.3%、36.7%，平均 31.0%",
               "2026-06-30 11,486；Sens 低／高 25%／37%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代第一輪的「中位數 × 0.8 再 Blume／中位數未調整」")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑；Sens β 案例用")

r = 4; IN_ROWS = {}
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
    if key.startswith(("tp_", "fade_g", "nim10")): IN.cell(r, 5).fill = KEY
    IN_ROWS[key] = r
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
FD.yellow(IN)   # v0.2：標準版 DCF 的判斷參數黃底（Andy 可改）
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
# 情境別名稱：份額相對值與液冷滲透率（FY2027 起 A＝高、B＝基準、C＝低；FY2026 滲透率三情境皆基準）
for y in (27, 28, 29):
    rr = IN_ROWS[f"rel_{y}"]
    defname(f"relA_{y}", f"In!$F${rr}"); defname(f"relB_{y}", f"In!$E${rr}"); defname(f"relC_{y}", f"In!$D${rr}")
for y in (26, 27, 28, 29):
    rr = IN_ROWS[f"lpen_{y}"]
    for s_, col in (("A", "F"), ("B", "E"), ("C", "D")):
        defname(f"lpen{s_}_{y}", f"In!$E${rr}" if y == 26 else f"In!${col}${rr}")

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自雙鴻財報（經玩股網、時報、HiStock 轉載）與月營收"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（月營收）＋3Q、4Q 假設（毛利率、營業費用率、業外、淨利 ÷ 稅前；4Q 營收＝3Q × (1＋季增)）。3Q26 財報與法說（2025 年為 11 月中）尚未公布。各季毛利、營業利益、稅前取自玩股網各季損益（合併）；歸屬母公司淨利與 EPS 取自公司公告（經時報、凱基、HiStock）"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：" + WR03 + "；各季：" + WG
S2 = "凱基 2026-08-07（" + KGI08 + "）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 23.276, "C": "=8.551657+8.700121", "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Interested-party（4Q 季增）", s2="FY2025 232.76 億元（時報 2026-03-04）；1Q26 85.52、2Q26 87.00 億元（玩股網；凱基 2Q 87.00）")
qrow(5, "毛利", {"B": 6.372, "C": "=2.536523+2.472471", "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Interested-party（3Q、4Q）", s2="FY2025 毛利率 27.38%（時報）；2Q26 28.4%（凱基）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 27.4%；1H26 29.0%")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived（毛利 − 營業利益）／Interested-party")
qrow(8, "營業利益", {"B": 3.259, "C": "=1.519644+1.390279", "D": "=D5-D7", "E": "=E5-E7"}, tag="Verified／Derived", s2="FY2025 營益率 14.0%（時報）；2Q26 13.90 億元、營益率 16.0%（凱基）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "業外收益", {"B": "=B11-B8", "C": "=C11-C8", "D": "=q_nop", "E": "=q_nop"}, tag="Derived")
qrow(11, "稅前利益", {"B": 3.277436, "C": "=1.570048+1.317191", "D": "=D8+D10", "E": "=E8+E10"}, tag="Verified（資料商轉載）／Derived", s1=WG, s2="FY2025 為四季加總")
qrow(12, "歸屬母公司淨利", {"B": 2.572, "C": "=1.164+0.969", "D": "=D11*q_nir", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）",
     s2="FY2025 25.72 億（時報 2026-03-04）；1Q26 11.64 億（時報 2026-05-21）；2Q26 9.69 億（凱基）；合併淨利 FY2025 26.98、1Q26 12.65、2Q26 10.39（玩股網）")
qrow(13, "淨利 ÷ 稅前（含稅與非控制權益）", {c: f"={c}12/{c}11" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股）", {"B": "=2.572/28.26", "C": "=2.133/23.02", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived／Verified", s2="FY2025、1H26＝淨利 ÷ EPS（加權）；3Q 起＝已發行股數")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="FY2026 錨點以已發行股數換算（1H26 加權股數較少，差約 0.4%）")
qrow(16, "  公司揭露 EPS", {"B": 28.26, "C": 23.02}, fmt="0.00", ann="", tag="Verified", s1=HISEPS, s2="1Q26 12.61、2Q26 10.41；3Q25 10.66、4Q25 10.19（" + WV + "）")
SPL = [  # (列, 標籤, FY2025 公式, FY2026 公式)
    (17, "液冷營收（＝營收 × 液冷占比）", "=B4*lc_25", "=F4*lc_26"),
    (18, "伺服器氣冷（＝營收 ×（伺服器 − 液冷））", "=B4*(srv_25-lc_25)", "=F4*(srv_26-lc_26)"),
    (19, "筆電與桌機（＝營收 × 筆電與桌機占比）", "=B4*nb_25", "=F4*nb_26"),
    (20, "電競筆電、顯卡、主機板與其他（＝營收 ×（1 − 伺服器 − 筆電與桌機））", "=B4*(1-srv_25-nb_25)", "=F4*(1-srv_26-nb_26)"),
    (21, "液冷（Sens：2026 液冷占比 低）", None, "=F4*lc_26_lo"),
    (22, "伺服器氣冷（Sens：液冷占比 低）", None, "=F4*(srv_26-lc_26_lo)"),
    (23, "液冷（Sens：2026 液冷占比 高）", None, "=F4*lc_26_hi"),
    (24, "伺服器氣冷（Sens：液冷占比 高）", None, "=F4*(srv_26-lc_26_hi)")]
for rr, lab, fb, ff in SPL:
    qrow(rr, lab, {"B": fb} if fb else {}, ann=ff, tag="Derived／Interested-party" if rr < 21 else "Derived", s1="公司與凱基產品比重（In 頁 C 節）" if rr < 21 else "Sens 用",
         s2="液冷、伺服器氣冷的季度金額未揭露：未找到" if rr == 17 else "—")
Q.cell(26, 1, "營業費用彈性（FY2026 營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B26"] = "=(F7/B7-1)/(F4/B4-1)"; Q["B26"].number_format = "0.00"
Q.cell(28, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(29, 1, "FY2026 營收年增"); Q["B29"] = "=F4/B4-1"
Q.cell(30, 1, "FY2026 液冷營收年增（公司：+190–200%）"); Q["B30"] = "=F17/B17-1"
Q.cell(31, 1, "FY2026 伺服器營收年增（公司：+110–120%）"); Q["B31"] = "=(F17+F18)/(B17+B18)-1"
Q.cell(32, 1, "FY2026 營收 ÷ 共識（stockanalysis）− 1"); Q["B32"] = "=F4/c_rev26-1"
Q.cell(33, 1, "FY2026 EPS ÷ 共識（stockanalysis）− 1"); Q["B33"] = "=F15/c_eps26-1"
for rr in range(29, 34): Q[f"B{rr}"].number_format = "0.0%"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nop", 10), ("pt", 11), ("ni", 12), ("sh", 14), ("eps", 15),
               ("lc", 17), ("sair", 18), ("nb", 19), ("oth", 20), ("lcLL", 21), ("sairLL", 22), ("lcLH", 23), ("sairLH", 24)):
    if rr <= 20: defname(f"a25_{nm}", f"Q!$B${rr}")
    defname(f"a26_{nm}", f"Q!$F${rr}")
Q.column_dimensions["A"].width = 64
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 26; Q.column_dimensions["H"].width = 90; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
_mfy = lambda m, y: m[y] if isinstance(m, dict) else m   # Siliconomics v0.13：mf 可為年別 dict（Sens「Wattonomics 錨定」案例）；預設字串不變
DEFP = dict(scn="B", rel=None, lcadj="lcadj", lpen=None, mf="1", gs="g_sair", gb="g_nb", go="g_oth", d="d_lc", opxa="0", mani="mani",
            lc="lc", sair="sair", nir="nir_f")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    P["rel"] = P["rel"] or f"rel{s}"; P["lpen"] = P["lpen"] or f"lpen{s}"
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
    m = P["mani"]
    sec_("A. AI 部署與液冷市場（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{_mfy(P['mf'], y)}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in BU[1:]}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    row("lpen", "液冷滲透率（部署 GW 中液冷占比）", "%", {c_(y): f"={P['lpen']}_{y % 100}" for y in BU}, fmt="0.0%")
    row("cpg", "每 GW 液冷內容（機架內液冷＋分歧管；與奇鋐模型同口徑）", "US$M/GW",
        {"D": f"=((1-vr26)*racks_gb*(cool_gb+{m})+vr26*racks_vr*(cool_vr+{m}))/1000",
         **{c_(y): f"=racks_vr*(cool_vr+{m})*{P['lcadj']}_{y % 100}/1000" for y in BU[1:]}}, fmt="#,##0",
        note="橋接表每架機架內液冷 × 每 GW 機架數；FY2027 起 × 內容調整（Rubin 全液冷、ASIC 液冷）")
    row("lmk", "AI 液冷市場（美元）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['lpen']}*{c_(y)}{R['cpg']}/1000" for y in BU})
    row("lmn", "AI 液冷市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['lmk']}*fx" for y in BU}, fmt="#,##0.0")
    row("sh0", "雙鴻有效份額（FY2026 校準）＝ 液冷營收 ÷ 市場", "%", {"D": f"=a26_{P['lc']}/D{R['lmn']}"}, fmt="0.0%",
        note="Derived：分子為券商液冷占比 × 營收（含 CDU、機架分歧管）；分母為機架內液冷＋分歧管市場（不含機架外 CDU）")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "雙鴻有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("lc", "液冷營收", "NT$B", {"C": "=a25_lc", **{c_(y): f"={c_(y)}{R['lmn']}*{c_(y)}{R['sh']}" for y in BU}}, fmt="#,##0.0")
    sec_("B. 非 AI（伺服器氣冷、筆電與桌機、電競顯卡與其他）與合計")
    row("sair", "伺服器氣冷營收（含 AI 氣冷與一般伺服器）", "NT$B", {"C": "=a25_sair", "D": f"=a26_{P['sair']}", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gs']}_{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("nb", "筆電與桌機營收", "NT$B", {"C": "=a25_nb", "D": "=a26_nb", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gb']}_{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("oth", "電競筆電、顯卡、主機板與其他營收", "NT$B", {"C": "=a25_oth", "D": "=a26_oth", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['go']}_{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['lc']}+{c_(y)}{R['sair']}+{c_(y)}{R['nb']}+{c_(y)}{R['oth']}" for y in FYS}, fmt="#,##0.0")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("lcsh", "液冷 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['lc']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pgw", "每 GW 雙鴻液冷營收", "US$M/GW", {c_(y): f"={c_(y)}{R['lc']}/fx/{c_(y)}{R['gw']}*1000" for y in BU}, fmt="#,##0")
    sec_("C. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × 液冷占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['lcsh']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；液冷毛利率＝其他 + 差距。FY2027 起毛利率隨液冷占比變動")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"=$D${R['gmn']}*{c_(y)}{R['rev']}+{P['d']}*{c_(y)}{R['lc']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,phi*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nop", "業外收益", "NT$B", {"C": "=a25_nop", "D": "=a26_nop", **{c_(y): "=nop_f" for y in BU[1:]}}, fmt="#,##0.0")
    row("nir", "歸屬母公司淨利 ÷ 稅前", "%", {"C": "=a25_ni/(a25_oi+a25_nop)", "D": "=a26_ni/(a26_oi+a26_nop)", **{c_(y): f"={P['nir']}" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['nop']})*{c_(y)}{R['nir']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.0000")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 雙鴻損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設），FY2027–FY2029 為驅動因子推導，FY2030–FY2032 為收斂法（本頁延伸表，灰底）。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("lc", "  液冷營收", "NT$B", "#,##0.0"),
                                ("sair", "  伺服器氣冷營收", "NT$B", "#,##0.0"), ("nb", "  筆電與桌機營收", "NT$B", "#,##0.0"), ("oth", "  電競、顯卡與其他營收", "NT$B", "#,##0.0"),
                                ("lcsh", "  液冷 ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"),
                                ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收 − 錨點；FY2026 模型液冷 − 錨點；FY2026 毛利 − 錨點").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 B：營收差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['B']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 B：液冷差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['B']['lc']}-a26_lc"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 B：毛利差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['B']['gmn']}*D{PB['B']['rev']}+d_lc*D{PB['B']['lc']}-a26_gp"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14
# --- PL 延伸表（三情境到 FY2032；FY2030 起為收斂法） ---
XY = (2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032); XC = {y: "CDEFGHIJ"[i] for i, y in enumerate(XY)}; FADE = (2030, 2031, 2032)
xr = PL.max_row + 3; XT = xr
PL.cell(xr, 1, "延伸：三情境到 FY2032。FY2025–FY2029 引用上方推導區塊；FY2030 起為收斂法（灰底），非驅動因子推導：營收年增由 FY2029 線性收斂到 fade_g、淨利率由 FY2029 線性收斂到收斂目標，FY2032 達到").font = BOLD
xr += 1; XR = {}; XH = None
for s in SC:
    E = PB[s]; R = {}
    PL.cell(xr, 1, f"延伸｜{SCN_T[s]}").font = BOLD; xr += 1
    hdr(PL, xr, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    XH = XH or xr; xr += 1
    for key, lab, unit in (("rev", "營收（延伸）", "NT$B"), ("g", "  營收年增", "%"), ("ni", "淨利（延伸）", "NT$B"), ("nim", "  淨利率", "%"),
                           ("sh", "股數（延伸）", "十億股"), ("eps", "EPS（延伸）", "NT$"), ("epsg", "  EPS 年增", "%"),
                           ("tg", "收斂目標：FY2032 營收年增", "%"), ("tm", "收斂目標：FY2032 淨利率", "%")):
        PL.cell(xr, 1, lab); PL.cell(xr, 2, unit); R[key] = xr; xr += 1
    G = lambda k: f"$G${R[k]}"
    for y in XY:
        c = XC[y]; pc = XC.get(y - 1)
        if y <= 2029:
            PL[f"{c}{R['rev']}"] = f"={c}{E['rev']}"
            if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            PL[f"{c}{R['ni']}"] = f"={c}{E['ni']}"
            PL[f"{c}{R['sh']}"] = f"={c}{E['sh_']}"
            PL[f"{c}{R['eps']}"] = f"={c}{E['eps']}"
            PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
        else:
            k = y - 2029
            PL[f"{c}{R['g']}"] = f"={G('g')}+($C${R['tg']}-{G('g')})*{k}/3"
            PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
            PL[f"{c}{R['nim']}"] = f"={G('nim')}+($C${R['tm']}-{G('nim')})*{k}/3"
            PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
            PL[f"{c}{R['sh']}"] = f"={pc}{R['sh']}*(1+sh_g)"
            PL[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh']}"
        if y > 2025: PL[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("rev", "#,##0.0"), ("g", "0.0%"), ("ni", "#,##0.0"), ("nim", "0.0%"), ("sh", "0.0000"), ("eps", "0.00"), ("epsg", "0.0%")):
            PL[f"{c}{R[key]}"].number_format = fm
            if y in FADE: PL[f"{c}{R[key]}"].fill = FADEFILL
    PL[f"C{R['tg']}"] = f"=fade_g{s}"; PL[f"C{R['tg']}"].number_format = "0.0%"
    PL[f"C{R['tm']}"] = {"A": f"=G{R['nim']}", "B": f"=(G{R['nim']}+nim10)/2", "C": "=nim10"}[s]; PL[f"C{R['tm']}"].number_format = "0.0%"
    PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導（上游市場層只到 2029）")
    PL.cell(R["tg"], 11, {"A": "In 頁 fade_gA＝tp_g＋2pt", "B": "In 頁 fade_gB＝tp_g", "C": "In 頁 fade_gC＝tp_g−2pt"}[s])
    PL.cell(R["tm"], 11, {"A": "＝本情境 FY2029 淨利率", "B": "＝FY2029 淨利率與歷史平均淨利率（In 頁 nim10，2017–2025 九年）的中點", "C": "＝歷史平均淨利率（In 頁 nim10，2017–2025 九年；2016 年未找到）"}[s])
    PL.cell(R["sh"], 11, "FY2030 起股數年變動＝sh_g")
    XR[s] = R; xr += 1
PL.column_dimensions["I"].width = 14; PL.column_dimensions["J"].width = 14; PL.column_dimensions["K"].width = 60
XEPS = lambda s, y: f"PL!{XC[y]}{XR[s]['eps']}"

# ---------------- Units：液冷市場、份額與兩家份額合計 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 液冷市場（機架內液冷＋分歧管，與奇鋐模型同口徑）、雙鴻份額與每 GW 營收（情境 B）。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）；每架液冷取自橋接表"
B_ = PB["B"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("gw", "部署 GW（情境 B）", "GW", lambda y: f"=PL!{FC[y]}{B_['gw']}", "0.0", "AI IT 資本支出 × 顆數係數 B ÷ 每 GW IT 資本支出"),
        ("lpen", "液冷滲透率（部署 GW）", "%", lambda y: f"=PL!{FC[y]}{B_['lpen']}", "0.0%", "奇鋐模型的假設（In 頁 B 節）"),
        ("cpg", "每 GW 液冷內容", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['cpg']}", "#,##0", "橋接表機架內液冷＋分歧管 × 每 GW 機架數 × 內容調整"),
        ("lmk", "AI 液冷市場（機架內液冷＋分歧管）", "US$B", lambda y: f"=PL!{FC[y]}{B_['lmk']}", "#,##0.0", "不含機架外 CDU、液對氣"),
        ("sh", "雙鴻有效份額", "%", lambda y: f"=PL!{FC[y]}{B_['sh']}", "0.0%", "FY2026 以凱基液冷占比校準"),
        ("avsh", "奇鋐有效份額（奇鋐模型）", "%", lambda y: f"=avsh_{y % 100}", "0.0%", "奇鋐模型 Units（唯讀）；奇鋐三情境份額相同"),
        ("sum", "兩家有效份額合計（奇鋐＋雙鴻）", "%", lambda y: f"={cyc(y)}{UR['sh']}+{cyc(y)}{UR['avsh']}", "0.0%", "兩家分子都含 CDU 與機架外產品，合計偏高估實際冷板份額"),
        ("oth", "其他廠商（酷碼、台達電、富世達、Boyd、CoolIT 等）＝ 1 − 兩家合計", "%", lambda y: f"=1-{cyc(y)}{UR['sum']}", "0.0%", ""),
        ("lc", "雙鴻液冷營收", "US$B", lambda y: f"=PL!{FC[y]}{B_['lc']}/fx", "#,##0.00", ""),
        ("pgw", "每 GW 雙鴻液冷營收", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['pgw']}", "#,##0", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit); UR[key] = ur
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：液冷市場口徑、滲透率、每架內容調整與奇鋐模型相同；雙鴻份額以 FY2026 液冷占比（凱基）校準。兩家有效份額合計約六成，其餘為酷碼、台達電、富世達、Boyd、CoolIT 等；合計若超過八成即不合理，應回頭檢查份額相對值或市場口徑。").font = BOLD
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。目標價＝DCF 與 Comparable 加權（第 0 節）；另列「現價 ÷ FY2027、FY2028 EPS（模型）」與十年 P/E 規則值（年底落後 P/E，基準為 FY2027 EPS；參考，不是目標價）"
V["A2"] = "循環位置：雙鴻不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照奇鋐、台達電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
V1 = 60   # 第 1–4 節與規則區塊的起始列（第 0 節放在本頁最前面）
V.cell(V1, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, V1 + 1, ["項目", "A", "B", "C", "說明"])
V.cell(V1 + 2, 1, "現價（2026-11-04 收盤）"); V[f"B{V1 + 2}"] = "=px"; V[f"B{V1 + 2}"].number_format = "#,##0"
for k_, (lab, y) in enumerate((("現價 ÷ FY2026 EPS（錨點）", 2026), ("現價 ÷ FY2027 EPS（模型）", 2027), ("現價 ÷ FY2028 EPS（模型）", 2028), ("現價 ÷ FY2029 EPS（模型）", 2029))):
    rr = V1 + 3 + k_; V.cell(rr, 1, lab)
    for j, s in enumerate(SC):
        V[f"{'BCD'[j]}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{'BCD'[j]}{rr}"].number_format = "0.0"
PXR = V1 + 2
V.cell(V1 + 7, 1, "未來 12 月 EPS（模型 B）≈ 2／12 × FY2026＋10／12 × FY2027"); V[f"B{V1 + 7}"] = f"=2/12*PL!D{SUMR[('B', 'eps')]}+10/12*PL!E{SUMR[('B', 'eps')]}"; V[f"B{V1 + 7}"].number_format = "0.00"
V.cell(V1 + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{V1 + 8}"] = f"=px/B{V1 + 7}"; V[f"B{V1 + 8}"].number_format = "0.0"
V.cell(V1 + 9, 1, "近 12 月 P/E（3Q25–2Q26 EPS，現價 11-04）"); V[f"B{V1 + 9}"] = "=px/eps_ttm"; V[f"B{V1 + 9}"].number_format = "0.0"
V.cell(V1 + 10, 1, "前瞻 P/E（stockanalysis，2026-08-28；價格 1,140）"); V[f"B{V1 + 10}"] = 17.27; V[f"B{V1 + 10}"].font = BLUE; V[f"B{V1 + 10}"].number_format = "0.0"
V[f"E{V1 + 10}"] = "Interested-party（賣方共識）；10 月價格下的前瞻 P/E：未找到"
C2 = V1 + 12
V.cell(C2, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, C2 + 1, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, C2 + 2):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
C3 = C2 + 6
V.cell(C3, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數（參考，不是目標價）").font = BOLD
hdr(V, C3 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = C3 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{C2 + 2 + j}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
C4 = C3 + 6
V.cell(C4, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, C4 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數", "對照：十年 75 分位"])
vr = C4 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_mid").number_format = "0.0"; V.cell(vr, 5, "=pe_trough").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。本版未做先驗，无 Andy 合理 P/E 對照。".replace("无", "無"))
WMBT_NOTE = vr; vr += 1
RB0 = vr + 2
# 十年年底 P/E：Goodinfo 年底收盤 ÷ 當年 EPS（2017–2025；2016 年未找到）；stockanalysis 年底 P/E（2021–2025，只作核對）
GI_PX = [None, 73, 97.9, 220, 214, 202, 148, 352.5, 672, 1010]
GI_EPS = [None, 3.66, 2.90, 11.71, 12.88, 13.12, 14.68, 14.28, 21.23, 28.26]
PE_GI = [None if p is None else f"={p}/{e}" for p, e in zip(GI_PX, GI_EPS)]
PE_SA = [None] * 5 + [15.35, 10.08, 24.65, 31.96, 35.83]
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "Goodinfo（年底收盤 ÷ 當年 EPS）", "vals": PE_GI, "tag": "Derived（資料商年底收盤與 EPS；2017–2025 共 9 年，2016 年未找到）"},
    {"label": "stockanalysis（年底 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；只有 2021–2025，只作核對、不採用）", "use": False,
     "note": "stockanalysis 2021–2025 與 Goodinfo 推導差距在 0.4x 內；HiStock 12 月 P/E（近 4 季 EPS）2021–2025：17.2、9.1、25.7、37.2、44.0"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用 Goodinfo 年底收盤 ÷ 當年 EPS（9 年；2016 年未找到）；stockanalysis 只有 2021–2025，只作核對。2017–2022 年雙鴻為筆電、顯卡與伺服器氣冷散熱廠（約 10–34x，2018 年 EPS 低使倍數偏高），AI 前後的倍數口徑不同。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
PE10 = f"$B${RB0 + 5}:$B${RB0 + 14}"
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22：DCF 與 Comparable 加權；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# v0.2：0.1 節改為標準版 DCF（FCFF＋WACC）摘要（tools/fcff_dcf.py），明細在 DCF 頁；原 0.1 節占第 5–24 列，摘要 18 列（第 5–22 列），0.2 節起位置不變
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（v0.2）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2027 起整年）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0")
DCFW["A2"] = ("EBIT：FY2027–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸表只有淨利）由淨利還原：EBIT＝淨利 ÷ nir_f − nop_f（nir_f＝歸屬母公司淨利 ÷ 稅前）。"
              "NOPAT＝EBIT ×（1 − tax_w），為合併口徑，評價時扣非控制權益帳面值；業外收益不進 NOPAT，現金與長期投資另在評價時加入。"
              "折舊攤銷、資本支出、營運資金：近三年實績占營收比例（In 頁 fc_da、fc_cx、fc_wc），FY2029 起資本支出比例線性收斂到 FY2032 穩態 fc_cx_ss。v0.2.3：淨現金由資產負債表日 2026-06-30 滾到評價日（＋含 stub 期的會計年度 FY2026 FCFF × stub_frac − 已宣告股利；Val 0.1 節；v0.2.2 誤用 FY2027 FCFF）。")
FY_X = list(range(2027, 2033))
_ebit = lambda s_, y: f"PL!{XC[y]}{PB[s_]['oi']}" if y <= 2029 else f"=PL!{XC[y]}{XR[s_]['ni']}/nir_f-nop_f"
DPATHS = [dict(name=s_, label=SCN_T[s_], rev={y: f"PL!{XC[y]}{XR[s_]['rev']}" for y in FY_X}, rev0=f"PL!{XC[2026]}{XR[s_]['rev']}",
               ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "tax_w" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷ nir_f − nop_f",
               stub_fcff=f"(PL!{XC[2026]}{PB[s_]['oi']}*(1-tax_w)+PL!{XC[2026]}{XR[s_]['rev']}*(fc_da-fc_cx-fc_cx_d)-fc_wc*(PL!{XC[2026]}{XR[s_]['rev']}-PL!{XC[2025]}{XR[s_]['rev']}))*stub_frac") for s_ in SC]
# 第三輪修正（fixes_round3 第 2、25 項）：stub 期（2026-07-01～12-31）FCFF＝含 stub 期的會計年度 FY2026（日曆年）FCFF × stub_frac，不用 FY2027 全年
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 1.0, conv_from=2029, title="三情境 FCFF 明細（NT$B；會計年度＝日曆年，FY2027 起整年）", fmt="#,##0.00", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs=NC_BS, nc_note=NC_NOTE) + 1
DCFW.column_dimensions["A"].width = 60; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.2；明細見 DCF 頁；B／C／D＝A 順利／B 平平／C 不如預期）", beta_note="自算產業 β 重新加槓桿")
for _k in ("pvx", "pvtv", "ev", "nc"):
    for _c in "BCD": V[f"{_c}{SDCF[_k]}"].number_format = "#,##0.00"
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"], h1=SDCF["h"])
EPSR = {}; CFR = {}
_D12 = SDCF["v12"]
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E；非 A 股公司，同業不放 A 股）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("奇鋐（AVC）", "3017.TW", "=2820/(3585/27.04)", "2026-11-04"), ("健策（Jentech）", "3653.TW", "=6450/(6920/69.92)", "2026-11-04"),
         ("Vertiv", "VRT", "=205.55/(207.93/26.21)", "2026-11-04"), ("台達電（Delta）", "2308.TW", "=1220/(1560/30.51)", "2026-11-04")]
PEER_SRC = {"3017.TW": "stockanalysis 個股頁 2026-10-05：價格 3,585 ÷ Forward PE 27.04（奇鋐模型 cons_feps 同一筆）", "3653.TW": "stockanalysis 統計頁 2026-09-24：價格 6,920 ÷ Forward PE 69.92（可取得的最新頁面，較舊）",
            "VRT": "stockanalysis 統計頁 2026-10-27：價格 207.93 ÷ Forward PE 26.21", "2308.TW": "stockanalysis 個股頁 2026-10-22：價格 1,560 ÷ Forward PE 30.51（台達電模型 cons_feps 同一筆）"}
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, f"第三輪修正：P/E＝Yahoo Finance 日線收盤 2026-11-04 ÷ 資料商前瞻 EPS（{PEER_SRC[tk]}）"); V.cell(i, 6, "Interested-party（賣方共識）")
PEER_LAST = 27 + len(PEERS)
V.cell(PEER_LAST + 1, 1, "並列參考（不進目標價）：無。英維克（002837.SZ）為 A 股，依規則不列；酷碼（Cooler Master）、Boyd、CoolIT 未上市")
tset(PEER_LAST + 2, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0",
     "散熱與液冷同業：奇鋐、健策、Vertiv、台達電（與奇鋐模型同一比較組，雙鴻換成奇鋐）；非 A 股公司的同業名單不放中國 A 股（Andy 2026-10-25 11:53）")
FE = PEER_LAST + 3
tset(FE, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*(2/12*{XEPS(s, 2027)}+10/12*{XEPS(s, 2028)})" for s in SC), None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(FE + 1, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{FE}", None, None], "#,##0", "同業中位數 × 上列 EPS")
C03 = FE + 3
V.cell(C03, 1, "0.3 結論").font = BOLD
hdr(V, C03 + 1, ["項目", "數值", "", "", "說明"])
r_ = C03 + 2
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], "#,##0", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁"); r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], "#,##0", "引用 0.2"); r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "stockanalysis（S&P Global）：17 位分析師平均 1,523、中位數 1,550（頁面時間 2026-10-19）"); r_ += 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-11-04 收盤 1,300；Yahoo Finance 日線收盤）"); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%",
     "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝2／12 × FY2026＋10／12 × FY2027 共識（In 頁 cons_feps；期間為現在起算 12 個月，與本模型不同）"); r_ += 1
tset(r_, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%"); r_ += 1
V[f"E{T['gmul']}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
tset(r_, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{_D12},"#,##0")&"～"&TEXT(B{_D12},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）"); r_ += 1
tset(r_, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
     [(lambda e: f"=IF(px/{e}<MIN({PE10}),0,IF(px/{e}>MAX({PE10}),1,PERCENTRANK({PE10},px/{e})))")(XEPS(s, 2028)) for s in SC], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2017–2025 年底 P/E（第 6 節 Goodinfo；2016 年未找到）的分位；低於最小值記 0%、高於最大值記 100%；只作檢查，不進目標價"); r_ += 1
V.cell(r_, 1, "目標價方法"); T["m"] = r_
V[f"B{r_}"] = ("DCF（v0.2 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），"
               "g＝MIN(3%, 台灣 10 年期公債殖利率)、RONIC＝WACC；ke＝rf＋β ×（ERP 5%＋台灣國家風險溢酬），β 為自算產業 β（TW 散熱／電源同業中位數，Blume 調整）依雙鴻 D／E 重新加槓桿；"
               "資本支出（泰國、越南擴產期）與營運資金為近三年實績比例，FY2029 起資本支出收斂到穩態（DCF 頁）。"
               "Comparable：奇鋐、健策、Vertiv、台達電前瞻 P/E 中位數（臺美同業，不放 A 股）；健策倍數高（約 59x）、雙鴻自身前瞻 P/E 約 17x，同業倍數分散，中位數可能高估雙鴻的市場倍數。"
               "情境對應：A＝順利（顆數基準、滲透率與份額高值）、B＝平平（顆數 ×0.8／0.85、基準）、C＝不如預期（顆數 ×0.6、低值）")
assert r_ < V1 - 2, r_
for rr in (T["dcftp"], T["comptp"], T["tpw"]):
    V[f"B{rr}"].fill = KEY
V.cell(T["tpw"], 1).font = BOLD
defname("tp_final", f"Val!$B${T['tpw']}"); defname("tp_dcf", f"Val!$B${T['dcftp']}"); defname("tp_comp", f"Val!$B${T['comptp']}")
V["A3"] = f"目標價：見第 0 節（本頁第 4 列起）。第 3 節與第 6 節「規則值」保留給 AI 投資地圖抽數，不是目標價"
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 B＝基準輸入；有效份額、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、液冷營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 液冷營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 液冷營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.85、2028 0.75、2029 0.70（＝情境 C 的份額）", dict(rel="RLO")),
         ("份額上升（高）", "份額相對值 1.05（＝情境 A 的份額）", dict(rel="RHI")),
         ("每架液冷內容 低", "內容調整 1.0（只有橋接表的機架內液冷）", dict(lcadj="CALO")), ("每架液冷內容 高", "內容調整 2027 1.6、2028–29 1.8", dict(lcadj="CAHI")),
         ("液冷滲透率 低", "2026 45%、2027 60%、2028 70%、2029 75%（FY2026 份額反向調整）", dict(lpen="LPLO")), ("液冷滲透率 高", "2026 65%、2027 85%、2028–29 95%", dict(lpen="LPHI")),
         ("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("伺服器氣冷 低", "2027–29 每年 −20%", dict(gs="GSLO")), ("伺服器氣冷 高", "2027–29 每年 +5%", dict(gs="GSHI")),
         ("筆電與其他 低", "筆電與桌機 −10／−8／−8%、電競顯卡與其他 −15／−10／−10%", dict(gb="GBLO", go="GOLO")), ("筆電與其他 高", "筆電與桌機 +5%、電競顯卡與其他 +8%", dict(gb="GBHI", go="GOHI")),
         ("液冷毛利率差距 2pt", "液冷毛利率只高 2 個百分點", dict(d="d_lc_lo")), ("液冷毛利率差距 8pt", "高 8 個百分點", dict(d="d_lc_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi")),
         ("淨利 ÷ 稅前 0.72", "基準 0.74", dict(nir="nir_f_lo")), ("淨利 ÷ 稅前 0.78", "", dict(nir="nir_f_hi")),
         ("2026 液冷占比 55%（讀法）", "FY2026 液冷＝營收 × 55%（1H26 實績比例）；份額分子變小、氣冷變大", dict(lc="lcLL", sair="sairLL")),
         ("2026 液冷占比 63%（讀法）", "FY2026 液冷＝營收 × 63%", dict(lc="lcLH", sair="sairLH"))]
for nm, base, suf, yrs in (("RLO", "rel", "_lo", (27, 28, 29)), ("RHI", "rel", "_hi", (27, 28, 29)), ("CALO", "lcadj", "_lo", (27, 28, 29)), ("CAHI", "lcadj", "_hi", (27, 28, 29)),
                           ("LPLO", "lpen", "_lo", (26, 27, 28, 29)), ("LPHI", "lpen", "_hi", (26, 27, 28, 29)),
                           ("GSLO", "g_sair", "_lo", (27, 28, 29)), ("GSHI", "g_sair", "_hi", (27, 28, 29)),
                           ("GBLO", "g_nb", "_lo", (27, 28, 29)), ("GBHI", "g_nb", "_hi", (27, 28, 29)),
                           ("GOLO", "g_oth", "_lo", (27, 28, 29)), ("GOHI", "g_oth", "_hi", (27, 28, 29))):
    for y in yrs:
        target = f"{nm}_{y}"
        wb.defined_names[target] = DefinedName(target, attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "lc", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "lc", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "GL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 B）", "PL 區塊 B", PB["B"], "PL!")
base_row = 50; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 B）", scn="B", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 A（全部）", "顆數係數 A＋液冷滲透率、份額相對值高值", PB["A"], "PL!")
put_case(nn + 1, "情境 C（全部）", "顆數係數 C＋液冷滲透率、份額相對值低值", PB["C"], "PL!")
SENS_LAST = nn + 1; SENS_NOTE = nn + 3
SE.cell(SENS_NOTE, 1, "讀法：基準為情境 B。每個案例的 FY2026 有效份額與其他營收毛利率依該案例重算，所以 FY2026 恆等於錨點。「液冷滲透率」案例的 FY2026 份額會反向調整（滲透率低 → 份額高），只影響 FY2027 起的成長路徑；情境 A／C 的滲透率只在 FY2027 起取高／低值，所以影響較大。")
TPS0 = nn + 5
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt="#,##0",
                    title="目標價參數敏感度（v0.2 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=True)
TPS_LAST = _tvs[1]
USD_ROW = next(r for r in range(_tvs[0], _tvs[1] + 1) if SE.cell(r, 1).value == "美元口徑折現")   # 第二輪修正：Sens「美元口徑折現」列
assert TPS_LAST < 50, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、券商估計與奇鋐模型的份額（兩家份額合計檢查）。只作事後對照，不為對上數字而調參數。共識取自 stockanalysis（S&P Global）；券商多經媒體轉述"
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
RC.cell(rr_, 1, "A. 共識與券商估計（校準未使用）").font = BOLD; rr_ += 1
RECON_A = rr_
rec("FY2026 營收 vs 共識（stockanalysis）", "=c_rev26", pl("rev", 2026), "stockanalysis（2026-10-19）", "區間 375–399 億", fmt="#,##0.0")
rec("FY2026 營收 vs 凱基（08-07，3Q 前）", "=b_rev26", pl("rev", 2026), "凱基 2026-08-07", "凱基 3Q 估 102.7 億，實際 97.9 億", fmt="#,##0.0")
rec("FY2026 EPS vs 共識（stockanalysis）", "=c_eps26", pl("eps", 2026), "stockanalysis（2026-10-19）", "3Q26 財報前")
rec("FY2026 EPS vs 凱基（08-07）", "=b_eps26", pl("eps", 2026), "凱基 2026-08-07", "")
rec("FY2026 EPS vs 美系券商（09-18）", "=u_eps26", pl("eps", 2026), "經濟日報 2026-09-18", "")
rec("FY2027 營收 vs 共識（stockanalysis）", "=c_rev27", pl("rev", 2027), "stockanalysis（2026-10-19）", "", fmt="#,##0.0")
rec("FY2027 營收 vs 凱基（08-07）", "=b_rev27", pl("rev", 2027), "凱基 2026-08-07", "", fmt="#,##0.0")
rec("FY2027 EPS vs 共識（stockanalysis）", "=c_eps27", pl("eps", 2027), "stockanalysis（2026-10-19）", "")
rec("FY2027 EPS vs 凱基（08-07）", "=b_eps27", pl("eps", 2027), "凱基 2026-08-07", "")
rec("FY2027 EPS vs 美系券商（09-18）", "=u_eps27", pl("eps", 2027), "經濟日報 2026-09-18", "")
rec("FY2028 EPS vs 美系券商（09-18）", "=u_eps28", pl("eps", 2028), "經濟日報 2026-09-18", "FY2028 共識：未找到")
RC.cell(rr_, 1, "B. 液冷結構與兩家份額合計（機制檢查）").font = BOLD; rr_ += 1
RECON_B = rr_
rec("FY2027 液冷 ÷ 營收 vs 凱基（72%）", "=b_lc27", pl("lcsh", 2027), "凱基 2026-08-07", "", fmt="0.0%", pt=True)
rec("FY2026 雙鴻有效份額（本模型市場口徑）", None, lambda s: f"=PL!D{PB[s]['sh0']}", "—", "份額第二來源：未找到", fmt="0.0%")
for y in BU:
    rec(f"CY{y} 兩家有效份額合計（奇鋐＋雙鴻）", None, lambda s, y=y: f"=avsh_{y % 100}+PL!{FC[y]}{PB[s]['sh']}", "奇鋐模型 Units（唯讀）＋本模型",
        "奇鋐三情境份額相同（只讀其值）；兩家分子都含 CDU 等機架外產品，合計偏高估冷板份額。超過八成應視為不合理" if y == 2026 else "—", fmt="0.0%")
rec("FY2027 每 GW 雙鴻液冷營收（US$M）", None, pl("pgw", 2027), "—", "", fmt="#,##0")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "液冷毛利率＝此值＋d_lc", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "液冷占比上升使毛利率上升", fmt="0.0%", pt=True)
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E15", "—", "凱基 08-07 估 17.83（以 4Q 營收 115.0 億為基礎）", fmt="0.00")
RC.cell(rr_, 1, "D. 目標價對帳（模型欄為情境加權值，非情境 A）").font = BOLD; rr_ += 1
RECV2 = rr_
for lab, ref, mdl, src in (("12 個月後的前瞻 EPS（情境加權）vs 共識前瞻 EPS", "=cons_feps", f"=Val!B{T['feps']}", "stockanalysis 共識（期間不同）"),
                           ("目標價（DCF 與 Comparable 加權）vs 市場共識目標價", "=cons_tp", f"=Val!B{T['tpw']}", "stockanalysis（2026-10-19）"),
                           ("Comparable 目標價 vs 市場共識目標價", "=cons_tp", f"=Val!B{T['comptp']}", "同上")):
    RC.cell(rr_, 1, lab); RC.cell(rr_, 2, ref).number_format = "#,##0.00"; RC.cell(rr_, 3, mdl).number_format = "#,##0.00"
    RC.cell(rr_, 4, f"=C{rr_}/B{rr_}-1").number_format = "0.0%"; RC.cell(rr_, 9, src); RC.cell(rr_, 10, "共識只作對帳，不作校準"); rr_ += 1
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); RECON_END = rr_; rr_ += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文經媒體與資料商轉載：法說簡報原檔未逐份抓取；券商與研調皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司（經時報）：FY2025 財報 2026-03-04（營收、毛利率、營業利益、淨利、EPS）", WR03),
        ("公司（經時報）：股東會 2026-05-21（1Q26 淨利、EPS；2025 產品比重）", WR05),
        ("資料商：玩股網各季損益（1Q25–2Q26 營收、毛利、營業利益、稅前、合併淨利）", WG),
        ("資料商：HiStock 月營收（2025-01–2026-09）、各季 EPS、已發行股數、股價", HIS),
        ("公司（經 TechNews）：9 月營收與前 9 月累計 2026-10-08", TN09),
        ("公司（經 Win 投資）：2Q26 EPS、1H26 累計 EPS 2026-10-08", WV),
        ("券商：凱基 2026-08-07（2Q26 實績、產品比重、液冷占比、3Q／4Q 與 2026–27 估計、目標價）", KGI08),
        ("券商：美系券商（經濟日報 2026-09-18；EPS 2026–28）", UDN0918),
        ("券商：花旗（經濟日報 2026-09-19；目標價 1,800、3Q／4Q 季增）", UDN0919),
        ("資料商：stockanalysis forecast 頁（共識目標價、2026–27 營收與 EPS，2026-10-19）", SAF),
        ("資料商：stockanalysis 個股頁（股數、Forward PE）與比率頁（年底 P/E、ROE）", SAR),
        ("資料商：Goodinfo 歷年年底收盤、EPS、淨利率、ROE（2017–2025；2016 年未取得）", GI),
        ("資料商：Yahoo 股市（2026-10-22 收盤）", YH),
        ("同業前瞻 P/E：stockanalysis 個股頁（3017、3653、VRT、2308）", "https://stockanalysis.com/quote/tpe/3017/"),
        ("市場層：markets/network、markets/chip_bridge、tsmc 成品；奇鋐模型 In 頁（液冷內容、滲透率、份額相對值）與 Units 頁（份額）（本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 財報與法說（預計 11 月中）；法說簡報原檔；液冷、CDU、分歧管的金額拆分；雙鴻液冷份額；2016 年 EPS 與年底股價；前瞻 P/E 十年歷史（Bloomberg）；FY2028 共識（付費）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("雙鴻（3324）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業；樣板＝奇鋐）　2026-10-25", True),
 ("版本：v0.1.1（2026-10-25：補上檔空間級距、投資建議、重點觀察；上游奇鋐 v0.2.1 液冷滲透率改依情境字母（奇鋐），雙鴻本已依情境字母讀奇鋐 In 頁同一組數值，數值不變）。v0.1（初版）。液冷市場口徑（每架機架內液冷取自橋接表、分歧管、內容調整、滲透率）與奇鋐模型相同（唯讀奇鋐模型 In 頁）；雙鴻份額以 FY2026 液冷占比校準，份額相對值區間與奇鋐相同。目標價＝DCF 與 Comparable 加權；三情境延伸到 FY2032（FY2030 起為收斂法）。", False),
 ("模型命題：雙鴻的液冷營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 液冷滲透率 × 每 GW 液冷內容 × 雙鴻有效份額；非 AI（伺服器氣冷、筆電與桌機、電競顯卡與其他）另以成長率推估。液冷已占 2026 年營收約六成，EPS 幾乎由液冷市場成長與份額決定。", False),
 ("核心問題：液冷市場隨部署 GW 與 Rubin 全液冷擴大時，雙鴻能不能守住相對奇鋐的份額（三個 ASIC 液冷專案、Vera Rubin 分歧管），讓 FY2027–28 EPS 達到券商預期的高成長；現價隱含的倍數合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 為驅動因子推導，FY2030–FY2032 為收斂法，非驅動因子推導。金額：十億新台幣；市場以美元計、以匯率換算。口徑：歸屬母公司淨利。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("液冷營收：markets/network 的 AI IT 資本支出 × 同字母顆數係數（tsmc；＝markets/memory SCEN_UNITS）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 液冷滲透率 × 每 GW 液冷內容（橋接表每架機架內液冷＋分歧管 × 每 GW 機架數 × 內容調整；皆讀奇鋐模型）× 匯率 × 雙鴻有效份額（FY2026 以凱基液冷占比校準）× 份額相對值。", False),
 ("非 AI：伺服器氣冷（含 AI 氣冷）、筆電與桌機、電競筆電顯卡主機板與其他，FY2026 以產品比重拆分 × 成長率（Assumed）。毛利率：液冷比其他營收高 d_lc（Assumed），其他毛利率為 FY2026 殘差（Derived）。營業費用：彈性 phi（Derived）。", False),
 ("情境：A＝順利（顆數係數 A、FY2027 起液冷滲透率高值、份額相對值高值）；B＝平平（顆數係數 B、基準）；C＝不如預期（顆數係數 C、低值）。液冷滲透率屬行業情境（技術替代），份額屬公司特有變數；方向一致：A 最有利、C 最不利。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績（月營收）＋假設（4Q 季增、毛利率、營業費用率）。不用共識或券商數字校準（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（v0.2 起為標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，FY2027–FY2032 明確期＋終值，以 WACC 折現；ke＝rf＋β ×（ERP＋國家風險溢酬）、g＝MIN(3%, rf)、RONIC＝WACC，三情境權重 0.3／0.5／0.2，以 ke 推到 12 個月後並扣期間股利；明細見 DCF 頁；負值取 0）＋0.5 × Comparable 目標價（奇鋐、健策、Vertiv、台達電前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、DCF 價格帶、十年 P/E 分位檢查。第 3、6 節十年 P/E 規則值保留給 AI 投資地圖，不是目標價。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文經媒體與資料商轉載；券商多經媒體轉述。", False),
 ("", False),
 ("已知限制：(1) 液冷占比只有凱基說法（2025 年 34%、1H26 55%、2026F 60%），公司只給液冷營收年增；(2) 液冷滲透率、每架內容調整、份額相對值為 Assumed（與奇鋐模型同值），是 FY2028 EPS 最敏感的輸入；(3) 本模型的液冷市場只含機架內液冷與分歧管，雙鴻的 CDU 營收計入有效份額的分子，份額偏高估；(4) 十年 P/E 只有 9 年（2016 年未找到），2017–2022 年為 AI 前的倍數；(5) 未建資本支出與折舊；(6) 3Q26 財報（預計 11 月中）後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ================= v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（工作單；Andy 2026-10-25 14:00、14:17） =================
# 全部放在 Val、README 現有最後一列之後，不位移既有儲存格
OBS = [
       "目標價 NT$2,769（DCF 1,979、Comparable 3,559 各半），比現價 1,360（2026-10-22）高 103.6%，級距大於 50%；Comparable 單獨為 161.7%。主因是 EPS 與倍數兩者：本模型 12 個月後前瞻 EPS 123.7 比共識前瞻 EPS 74.3 高 66.4%（兩者期間差約一年；FY2027 B 101.6 比共識 77.97 高 30.4%），同業（奇鋐、健策、Vertiv、台達電）前瞻 P/E 中位數 28.8 倍也高於現價隱含 18.3 倍與共識目標價 1,523 隱含的 20.5 倍；Comparable 未扣雙鴻相對奇鋐的份額折價（雙鴻自身前瞻 P/E 17.3 倍）。",
       "每架液冷內容調整是最敏感的變數：調整 1.0 時 FY2027 EPS −20.9%，2027 1.6 時 +20.9%；營業費用率 +2pt 時 −10.0%。液冷滲透率與內容調整讀奇鋐模型 In 頁同一組數值（奇鋐 v0.2.1 改依情境字母後，本模型數值不變）。",
       "情境分歧大：FY2028 EPS A／B／C＝188.5／117.8／61.3（A 約為 C 的三倍），DCF 情境價值 3,223／1,744／700；美系券商 FY2028 EPS 110 接近情境 B。10 月營收（約 11-10）與 3Q26 財報與法說（預計 11 月中）可檢驗 2026 年液冷占比約六成（凱基）。"
      ]   # 重點觀察（與 README.md 相同文字；數字取自重算後的 xlsx）
_U = f"B{T['upw']}"
GR = V.max_row + 3
V.cell(GR - 1, 1, "0.4 上檔空間級距與投資建議（v0.1.1 新增；放在本頁最後以免位移既有儲存格；依「上檔空間：加權目標價」第 " + str(T["upw"]) + " 列）").font = BOLD
tset(GR, "grade", "上檔空間級距", [f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "General",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）；負值區間同樣重要")
tset(GR + 1, "rec", "投資建議", [f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))', None, None], "General",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）")
rn = RD.max_row + 1
RD.cell(rn, 1, f'="投資建議："&Val!B{GR + 1}&"（依上檔空間固定規則）"').font = BOLD
OBS_ROW = rn + 1
for i_, t_ in enumerate(OBS, 1):
    RD.cell(rn + i_, 1, f"重點觀察 {i_}：" + t_)
# v0.2：重點觀察 3 的 DCF 情境價值改用公式（v0.2 標準版 DCF；其餘文字不變）
_o3 = OBS[2].split("DCF 情境價值 3,223／1,744／700")
assert len(_o3) == 2
RD.cell(rn + 3, 1, '="重點觀察 3：' + _o3[0] + 'DCF 情境價值（v0.2 標準版）"&TEXT(Val!B' + str(_D12) + ',"#,##0")&"／"&TEXT(Val!C' + str(_D12) + ',"#,##0")&"／"&TEXT(Val!D' + str(_D12) + ',"#,##0")&"' + _o3[1] + '"')


# ---------------- 分部原則第 3 條（2026-10-28）：核對列與宣告（追加於工作表最後，不位移既有列，數值不變） ----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(_os2.path.dirname(_os2.path.dirname(_os2.path.abspath(__file__))), "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[17, 18, 19, 20], total_row=4, cols="B", rnd=6, note="分部原則第 3 條（2026-10-28 追加）；分部宣告見 segments.json")
RD.cell(RD.max_row + 1, 1, "v0.1.2（2026-10-28）：分部原則重建（數值不變）。Q 頁最後追加核對列「檢查：分部合計 − 合併營收」；官方分部與子列宣告見 segments.json。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.1.1", "v0.1.2")

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "auras")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.3", "2026-10-30", split=True, split_skip=['Q'])  # v0.1.3：In 頁版面；v0.1.4（2026-10-31）：第 3 點情境專屬列拆分（數值不變）
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
WTA_SENS = _wta_sens({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}, '情境 B；AI IT 資本支出（→ 液冷市場） CY2027 起 × wta_27／28／29', '做法：引擎的 AI 部署係數 mf（原只乘在 AI IT 資本支出列、FY2027 起）改為年別 dict（wta_27／28／29）；FY2026 不乘。FY2030 以後無延伸表（Sens 引擎只到 FY2029）。目標價：Sens 目標價表只有 DCF 參數案例，本案例不重算目標價。')
# ---------------- v0.1.5（2026-11-05）：展示層（個股 HTML 版面規格 v1.0；docs/plan/個股HTML版面規格.md） ----------------
# 只寫 README 頁（既有列之後）；其他工作表不動。數字一律以公式取自各頁。
def _row(ws, label, start=1, exact=False):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, 1).value
        if isinstance(v_, str) and ((v_.strip() == label.strip()) if exact else v_.strip().startswith(label.strip())): return r_
    raise KeyError((ws.title, label))
_V = lambda lab, c="B": f"Val!{c}{_row(V, lab)}"
_I = lambda k: f"In!E{_row(IN, k, exact=True)}"
_pa, _pb, _pc = (_row(PL, f"情境 {s}", exact=True) for s in ("A", "B", "C"))
_PS = {"A": _pa, "B": _pb, "C": _pc}
_P = lambda lab, c, s="A": f"PL!{c}{_row(PL, lab, _PS[s], exact=True)}"
_BLK = {s: next(r_ for r_ in range(_PS["C"] + 1, PL.max_row + 1) if str(PL.cell(r_, 1).value or "").startswith(f"情境 {s}（")) for s in ("A", "B", "C")}   # 各情境推導區塊
_PD = lambda lab, c, s: f"PL!{c}{_row(PL, lab, _BLK[s])}"
_U = lambda lab, c: f"Units!{c}{_row(UN, lab)}"
_S = lambda lab, c: (f"Sens!A{_row(SE, lab, 4)}", f"Sens!{c}{_row(SE, lab, 4)}")
_RC = lambda lab, c="B": f"Recon!{c}{_row(RC, lab)}"
TP, DCF, CMP, CMK, PX, UP, UPC, REC = (_V(x) for x in ("目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（相對倍數調整後）",
                                               "Comparable 目標價（市場主流同業法）", "現價（", "上檔空間：加權目標價", "上檔空間：Comparable 目標價", "投資建議"))
CONS, EPSF, PEER, FEPS, RMK = _V("市場共識目標價"), _V("與市場共識目標價的差距：EPS 因素"), _V("同業前瞻 P/E 中位數"), _V("12 個月後的前瞻 EPS"), _V("相對倍數調整係數")
_dv = _row(V, "DCF 情境價值推到 12 個月後")
DV = lambda c: f"Val!{c}{_dv}"
T_ = lambda x, f="#,##0": f'TEXT({x},"{f}")'
P_ = lambda x, f="0.0%": f'TEXT({x},"{f}")'
SG = lambda x: f'TEXT({x},"+0.0%;-0.0%")'
_dir = lambda x: f'IF({x}<0,"低 ","高 ")&TEXT(ABS({x}),"0.0%")'
_hl = lambda a, b: f'IF({a}>{b},"高於","低於")'
_c27, _u27, _u28 = _I("c_eps27"), _I("u_eps27"), _I("u_eps28")
_s_ct, _s_mf, _s_sh = _S("每架液冷內容 低", "G"), _S("AI 部署 ×0.8", "G"), _S("份額流失（低）", "G")
_w = _I("tp_w_dcf")
_tm = f"{PEER}*{RMK}"            # 目標倍數（同業中位數 × 相對倍數調整係數）
_pm = f"({PX}/{_I('cons_feps')})"   # 現價對目前共識前瞻 EPS 的倍數
LCR = lambda c, s: _P("  液冷營收", c, s) + "*10"   # 十億 → 億元
rn = RD.max_row + 2
RD.cell(rn, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；原 README 各列保留）").font = BOLD; rn += 1
for t_ in [
    "命題：液冷市場隨部署 GW 與 Rubin 全液冷擴大時，雙鴻能否守住相對奇鋐的份額，讓 FY2027–28 EPS 達到券商預期的高成長，以及現價隱含的倍數是否合理；分成三個問題。",
    "命題 1｜液冷份額能否守住：ASIC 液冷專案與 Vera Rubin 分歧管是雙鴻的機會，但奇鋐、酷碼、台達電也在搶冷板份額；雙鴻的有效份額能否維持。",
    "命題 2｜EPS 能否達到券商預期：液冷已占營收約六成，EPS 幾乎由液冷市場決定；本模型三情境的 EPS 與共識、券商相比如何。",
    "命題 3｜現價合不合理：把 DCF 與同業倍數法加權後，每股價值與現價相比如何。",
    f'="重點 1｜加權目標價比現價"&{_dir(UP)}&"，投資建議"&{REC}&"：加權目標價約 "&{T_(TP)}&" 元，由 DCF 約 "&{T_(DCF)}&" 元與同業倍數法（相對倍數調整後）約 "&{T_(CMP)}&" 元加權（DCF 權重 "&TEXT({_w},"0%")&"）；現價 "&{T_(PX)}&" 元。只看 Comparable（市場主流同業法、未調整）約 "&{T_(CMK)}&" 元、上檔 "&{SG(UPC)}&"；雙鴻 10 年前瞻倍數長期低於同業，相對倍數調整係數 "&{T_(RMK, "0.00")}&"。"',
    f'="重點 2｜大上檔的主因是 Comparable，而 Comparable 的上檔主要來自盈餘差：加權上檔 "&{SG(UP)}&" 中，DCF 貢獻 "&TEXT({_w}*({DCF}/{PX}-1)*100,"+0.0;-0.0")&" 個百分點（占 "&TEXT({_w}*({DCF}-{PX})/({TP}-{PX}),"0%")&"），Comparable（相對倍數調整後）貢獻 "&TEXT((1-{_w})*({CMP}/{PX}-1)*100,"+0.0;-0.0")&" 個百分點（占 "&TEXT((1-{_w})*({CMP}-{PX})/({TP}-{PX}),"0%")&"）。Comparable 內，本模型 12 個月後的前瞻 EPS 約 "&{T_(FEPS, "0.0")}&" 元，比目前的共識前瞻 12 個月 EPS "&{T_(_I("cons_feps"), "0.0")}&" 元高 "&TEXT({EPSF},"0%")&"（含約 12 個月的時間差；同年度比較，情境 A／B／C 的 FY2027 EPS 約 "&{T_(_P("EPS", "E", "A"), "0.0")}&"／"&{T_(_P("EPS", "E", "B"), "0.0")}&"／"&{T_(_P("EPS", "E", "C"), "0.0")}&" 元，共識 "&{T_(_c27, "0.0")}&" 元）；目標倍數（同業中位數 "&TEXT({PEER},"0.0")&" 倍 × 調整係數 "&{T_(RMK, "0.00")}&"）約 "&TEXT({_tm},"0.0")&" 倍，比現價對共識前瞻 EPS 的 "&TEXT({_pm},"0.0")&" 倍"&IF({_tm}>{_pm},"高 ","低 ")&TEXT(ABS({_tm}/{_pm}-1),"0%")&"。盈餘差來自液冷營收：情境 B 由 FY2026 約 "&{T_(LCR("D", "B"))}&" 億元增至 FY2027 約 "&{T_(LCR("E", "B"))}&" 億元，靠部署 GW（"&{T_(_PD("部署 GW", "D", "B"), "0.0")}&" → "&{T_(_PD("部署 GW", "E", "B"), "0.0")}&"）與液冷滲透率（"&{P_(_PD("液冷滲透率", "D", "B"), "0%")}&" → "&{P_(_PD("液冷滲透率", "E", "B"), "0%")}&"，假設）。"',
    f'="重點 3｜情境分歧大，每架液冷內容與部署量最敏感：FY2028 EPS 情境 A／B／C 約 "&{T_(_P("EPS", "F", "A"), "0.0")}&"／"&{T_(_P("EPS", "F", "B"), "0.0")}&"／"&{T_(_P("EPS", "F", "C"), "0.0")}&" 元。以情境 B 為基準一次只改一個輸入時，「"&{_s_ct[0]}&"」使 FY2027 EPS 變動 "&{SG(_s_ct[1])}&"，「"&{_s_mf[0]}&"」變動 "&{SG(_s_mf[1])}&"，「"&{_s_sh[0]}&"」變動 "&{SG(_s_sh[1])}&"。這些是敏感度測試的案例，不是門檻。"',
    f'="重點 4｜v0.2 起 DCF 改為標準版（無槓桿自由現金流折現），DCF 目標價約 "&{T_(DCF)}&" 元：WACC "&TEXT(wacc,"0.0%")&"（股權成本 "&TEXT(ke,"0.0%")&"＝台灣 10 年期公債 "&TEXT(rf,"0.00%")&" ＋ β "&TEXT(beta,"0.00")&" ×（ERP "&TEXT(erp,"0%")&" ＋ 國家風險溢酬 "&TEXT(crp,"0.00%")&"）；β 為自算產業 β（TW 散熱／電源同業中位數，Blume 調整）依雙鴻負債比重新加槓桿，ERP 與 RONIC＝WACC 為假設），永續成長率 "&TEXT(g_tv,"0.00%")&"（不超過台灣 10 年期公債殖利率）。情境 B 的終值現值占企業價值 "&TEXT(Val!C{SDCF["tvsh"]},"0%")&"、隱含終值 EV／EBITDA 約 "&TEXT(Val!C{SDCF["mult"]},"0.0")&" 倍；DCF 參數中加權目標價對 "&IF(MAX(ABS(Sens!D{_tvs[0] + 1}),ABS(Sens!F{_tvs[0] + 1}))>MAX(ABS(Sens!D{_tvs[0] + 2}),ABS(Sens!F{_tvs[0] + 2})),"β","ERP")&" 最敏感（β 低／高（β_u ∓ 0.3）時 "&TEXT(Sens!D{_tvs[0] + 1},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0] + 1},"+0.0%;-0.0%")&"；ERP 4%／6% 時 "&TEXT(Sens!D{_tvs[0] + 2},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0] + 2},"+0.0%;-0.0%")&"；WACC ±1pt 時 "&TEXT(Sens!D{_tvs[0]},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{_tvs[0]},"+0.0%;-0.0%")&"；以三情境加權目標價為基準）。"',
    f'="重點 5｜美元口徑折現（敏感度案例，非基準）：基準以台灣 10 年期公債 "&TEXT(rf,"0.00%")&" 折現並假設匯率固定；若改用美國 10 年期公債 "&TEXT(rf_us,"0.00%")&"＋ β ×（ERP＋國家風險溢酬）、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")＝"&TEXT(MIN(g_cap,rf_us),"0.00%")&"，股權成本 "&TEXT(rf_us+beta*(erp+crp),"0.0%")&"，加權目標價 "&TEXT(Sens!C{USD_ROW},"#,##0")&" 元（相對基準 "&TEXT(Sens!D{USD_ROW},"+0.0%;-0.0%")&"）。"',
    f'="重點 6｜終值占比："&IF(MAX(Val!B{SDCF["tvsh"]}:D{SDCF["tvsh"]})>0.8,"終值現值占企業價值超過 80% 的情境："&IF(Val!B{SDCF["tvsh"]}>0.8,"A ","")&IF(Val!C{SDCF["tvsh"]}>0.8,"B ","")&IF(Val!D{SDCF["tvsh"]}>0.8,"C ","")&"（DCF 價值主要來自終值，對 g 與 WACC 較敏感）；","三情境終值現值占企業價值皆未超過 80%；")&"情境 A／B／C 分別為 "&TEXT(Val!B{SDCF["tvsh"]},"0%")&"／"&TEXT(Val!C{SDCF["tvsh"]},"0%")&"／"&TEXT(Val!D{SDCF["tvsh"]},"0%")&"。"']:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "命題回覆（HTML 用；依本模型計算回答開頭三個命題）").font = BOLD; rn += 1
for t_ in [
    f'="液冷份額能否守住？答：本模型的情境 B 假設不能完全守住。份額相對值 2027／2028／2029 為 "&{T_(_I("rel_27"), "0.00")}&"／"&{T_(_I("rel_28"), "0.00")}&"／"&{T_(_I("rel_29"), "0.00")}&"（假設），雙鴻有效份額由 FY2026 校準的 "&{P_(_U("雙鴻有效份額", "C"))}&" 降到 CY2028 的 "&{P_(_U("雙鴻有效份額", "E"))}&"；情境 A 假設升到 "&{T_(_I("rel_28_A"), "0.00")}&" 倍、情境 C 降到 "&{T_(_I("rel_28_C"), "0.00")}&" 倍（2028，假設）。奇鋐加雙鴻的有效份額合計由 "&{P_(_U("兩家有效份額合計", "C"), "0%")}&" 降到 CY2028 的 "&{P_(_U("兩家有效份額合計", "E"), "0%")}&"（情境 B）。即使份額下滑，液冷營收在情境 B 仍由 FY2026 約 "&{T_(LCR("D", "B"))}&" 億元增至 FY2028 約 "&{T_(LCR("F", "B"))}&" 億元，靠部署 GW 與液冷滲透率（假設）；情境 A／C 的 FY2028 液冷營收約 "&{T_(LCR("F", "A"))}&"／"&{T_(LCR("F", "C"))}&" 億元。"',
    f'="EPS 能否達到券商預期？答：要看情境。FY2027 EPS 情境 A／B／C 約 "&{T_(_P("EPS", "E", "A"), "0.0")}&"／"&{T_(_P("EPS", "E", "B"), "0.0")}&"／"&{T_(_P("EPS", "E", "C"), "0.0")}&" 元，對共識 "&{T_(_c27, "0.0")}&" 元分別"&{_hl(_P("EPS", "E", "A"), _c27)}&"、"&{_hl(_P("EPS", "E", "B"), _c27)}&"、"&{_hl(_P("EPS", "E", "C"), _c27)}&"，對美系券商 "&{T_(_u27, "0.0")}&" 元分別"&{_hl(_P("EPS", "E", "A"), _u27)}&"、"&{_hl(_P("EPS", "E", "B"), _u27)}&"、"&{_hl(_P("EPS", "E", "C"), _u27)}&"；FY2028 EPS 約 "&{T_(_P("EPS", "F", "A"), "0.0")}&"／"&{T_(_P("EPS", "F", "B"), "0.0")}&"／"&{T_(_P("EPS", "F", "C"), "0.0")}&" 元，對美系券商 "&{T_(_u28, "0.0")}&" 元分別"&{_hl(_P("EPS", "F", "A"), _u28)}&"、"&{_hl(_P("EPS", "F", "B"), _u28)}&"、"&{_hl(_P("EPS", "F", "C"), _u28)}&"。情境 B 毛利率由 FY2026 的 "&{P_(_P("毛利率", "D", "B"))}&" 升到 FY2028 的 "&{P_(_P("毛利率", "F", "B"))}&"，前提是液冷毛利率比其他營收高 "&TEXT({_I("d_lc")}*100,"0")&" 個百分點（假設）。"',
    f'="現價合不合理？答："&IF({UP}<0,"現價高於本模型的合理價值。","現價低於本模型的合理價值。")&"加權目標價約 "&{T_(TP)}&" 元（標準版 DCF "&{T_(DCF)}&"、同業倍數法相對倍數調整後 "&{T_(CMP)}&"），比現價 "&{T_(PX)}&" 元"&{_dir(UP)}&"，投資建議"&{REC}&"。但三情境差距很大：DCF 推到 12 個月後約 "&{T_(DV("B"))}&"／"&{T_(DV("C"))}&"／"&{T_(DV("D"))}&" 元（情境 A／B／C），情境 C "&IF({DV("D")}<{PX},"低於","高於")&"現價。只看 Comparable（市場主流同業法）約 "&{T_(CMK)}&" 元；市場共識目標價約 "&{T_(CONS)}&" 元，隱含約 "&TEXT({CONS}/{_I("cons_feps")},"0.0")&" 倍目前共識前瞻 EPS，低於同業中位數約 "&TEXT({PEER},"0.0")&" 倍。本模型加權目標價比共識"&{_dir(f"{TP}/{CONS}-1")}&"，主要因為本模型 EPS 較高；倍數"&IF(ABS({_tm}/({CONS}/{_I("cons_feps")})-1)<0.1,"相近","不同")&"（本模型目標倍數約 "&TEXT({_tm},"0.0")&" 倍，以 12 個月後（約 2027-10 起）的前瞻 EPS 計；共識約 "&TEXT({CONS}/{_I("cons_feps")},"0.0")&" 倍，以目前的前瞻 EPS（2／12 × FY2026＋10／12 × FY2027 共識）計；兩者時間窗不同，倍數不完全可比；EPS 比較的時間窗見重點 2）。"']:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "本版要點（HTML 附錄用）").font = BOLD; rn += 1
for t_ in [f'="v0.2.3（2026-11-06）第三輪修正（合併前）：(1) 現價改 2026-11-04 收盤 "&TEXT(px,"#,##0")&" 元（Yahoo Finance；原 2026-10-22 收盤 1,360）；Comparable 同業前瞻 P/E 改為 11-04 收盤 ÷ 資料商前瞻 EPS（健策原為 08-27 過期報價；中位數 "&TEXT(Val!B{PEER_LAST + 2},"0.0")&" 倍，原 28.8 倍）；"&"(2) stub 期 FCFF 改用含 stub 期的會計年度 FY2026 FCFF × stub_frac（情境 B "&TEXT(Val!C{SDCF["stub"]},"#,##0.00")&"；原 FY2027 FCFF × 0.504＝-0.77）；(3) 命題回覆 3 註明目標倍數與共識倍數的時間窗不同；"&"(4) 情境名稱統一為 A（順利）／B（平平）／C（不如預期）、級距列名、Sens g 列名、stub 口徑說明。DCF 目標價 1,588 → "&{T_(DCF)}&" 元、加權目標價 2,095 → "&{T_(TP)}&" 元、上檔空間 +54.0% → "&TEXT({UP},"+0.0%;-0.0%")&"、投資建議強力買進 → "&{REC}&"。"',
           'v0.2.2（2026-11-05）第二輪修正（換用共用模組 tools/fcff_dcf.py v2；Andy 2026-11-05 23:16 定案）：(1) 資產負債表維持最新已公布一季 2026-06-30（3Q26 季報尚未公布），淨現金以 stub 期滾到評價日 2026-12-31＝淨現金與投資 -3.92＋stub 期 FCFF（各情境 FY2027 FCFF × 0.504 年；情境 B -0.77）− 已宣告股利 1.11（FY2025 股利每股 11.903 元，2026-07-16 發放）＝-5.79（十億新台幣）；(2) β 組更名「TW 散熱／電源」，β_u 1.201 不變，Sens β 低／高改 β_u ∓ 0.3 各自重新加槓桿；(3) 新增 Sens「美元口徑折現」：加權目標價 1,870 元（-10.7%）；(4) 本模型無逐年資本支出／折舊列，維持占營收比例。DCF 目標價 1,611 → 1,588 元、加權目標價 2,106 → 2,095 元、上檔空間 +54.9% → +54.0%、投資建議強力買進 → 強力買進。',
           "v0.2.1（2026-11-05）β 改為自算產業 β（Andy 2026-11-05 22:55）：TW 散熱／電源／探針卡同業 6 家無槓桿 β 中位數 1.30，Blume 調整後 1.201，依雙鴻 D／E 重新加槓桿為 1.27；WACC 8.9%、DCF 目標價約 1,611 元、加權目標價約 2,106 元、上檔空間 +54.9%、投資建議強力買進。評價日 2026-12-31；其他模型輸入不變。",
           "v0.2（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；新增 DCF 頁，Val 0.1 節改為折現率與 DCF 彙總，Sens 目標價參數改為 WACC、β、ERP、g、RONIC、資本支出、營運資金等案例；In 頁新增 M2 節、刪除 tp_r、tp_roe）。Comparable、情境權重、DCF 權重 50% 與投資建議規則不變。",
           "v0.1.5（2026-11-05）只改展示層：README 新增「展示層」「命題回覆」兩段（數字以公式取自各頁），HTML 改用個股 HTML 版面規格 v1.0；README 以外逐格與 v0.1.4 相同。",
           "v0.1.4（2026-10-31）：情境專屬列由敏感度欄拆出（數值不變）。"]:
    RD.cell(rn, 1, t_); rn += 1
rn += 1
RD.cell(rn, 1, "HTML 橫幅附註用：進加權的 Comparable 目標價（相對倍數調整後）相對現價的上檔空間（橫幅中間格為只看 Comparable 目標價（市場主流同業法），其上檔空間見 Val 頁）")
RD.cell(rn, 2, f"={CMP}/{PX}-1").number_format = "0.0%"
RD["A1"] = RD["A1"].value + "；v0.1.5（2026-11-05）只改展示層：README 新增展示層與命題回覆段、HTML 改用版面規格 v1.0，數值不變；v0.2（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重與投資建議規則不變；v0.2.1（2026-11-05）β 改為自算產業 β（TW 散熱／電源同業中位數，Blume 調整；Andy 2026-11-05 22:55），其他輸入不變；v0.2.2（2026-11-05）第二輪修正：換用 tools/fcff_dcf.py v2（淨現金由 2026-06-30 滾到評價日並扣已宣告股利、β 敏感度 β_u ± 0.3、新增 Sens 美元口徑折現）"
_so = SO.max_row + 1   # 第二輪修正：資產負債表日與 stub 期股利等來源網址
for _t, _u in (("[Verified] 第二輪修正：資產負債表（季）2026.2Q 欄（富邦證券 e-Broker；資料商轉載合併報表；與 stockanalysis 2026-06-30 欄同值）", "https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcp/zcpa/zcpa_3324.djhtm"),
               ("[Verified] 第二輪修正：除權息表（FY2025 現金股利 11.903 元，除息 2026-06-17、發放 2026-07-16）", "https://www.wantgoo.com/stock/3324/dividend-policy/ex-dividend"),):
    SO.cell(_so, 1, _t); SO.cell(_so, 2, _u); _so += 1
# ---------------- 第三輪修正（2026-11-06；合併前）：展示與說明字串（fixes_round3 第 14、16、22、23、24、30 項） ----------------
# 只改說明字串、列名與 README 舊說明；數值變動只來自各公司專屬修正（見呼叫處）。共用模組 tools/fcff_dcf.py v2 不改，說明字串在此覆寫。
def _r3_text(wb, stub_lab, px_old_dates=(), px_new_date="2026-11-04", extra=()):
    """stub_lab：stub FCFF 實際口徑（取代模組預設字樣「第一年 FCFF × stub_frac」）；px_old_dates：README／Val 公式內寫死的舊現價日期（改為 px_new_date）"""
    import re as _re
    subs = list(extra) + [
        ("情境 A（樂觀）", "情境 A（順利）"), ("情境 B（基準）", "情境 B（平平）"), ("情境 C（保守）", "情境 C（不如預期）"), ("情境 C（下檔）", "情境 C（不如預期）"),
        ("隱含價（樂觀）", "隱含價（順利）"), ("隱含價（基準）", "隱含價（平平）"), ("隱含價（保守）", "隱含價（不如預期）"), ("隱含價（下檔）", "隱含價（不如預期）"),
        ("已逐份抓取", "已逐份讀取"), ("未逐份抓取", "未逐份讀取"), ("已抓取", "已讀取"), ("抓取", "讀取"), ("無法開啟", "未取得"), ("取不到", "未取得"),
        ("原樣取自 " + "claude" + "/html-v1-" + "nvidia", "tools/fcff_dcf.py v2"),   # 舊說明字樣（字串拆開，避免本檔含該字樣）
        ("stub FCFF＝第一年 FCFF × stub_frac", "stub FCFF＝" + stub_lab), ("（第一年 FCFF × stub_frac）", "（" + stub_lab + "）"),
        ("stub FCFF＝各路徑第一年 FCFF × 本值（年度比例）", "stub FCFF＝" + stub_lab),
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
                    w = _re.sub(a, b, w) if a.startswith("(?") else w.replace(a, b)
                if ws.title in ("README", "Val") and w.startswith("="):
                    for d in px_old_dates:
                        w = w.replace(f"（{d}）", f"（{px_new_date}）")
                # (22) README「重點觀察 1」：刪除 2026-10-29 相對倍數調整前撰寫的舊說明（數字已過時）
                if ws.title == "README" and w.startswith('="重點觀察 1：') and '&"以下為 2026-10-29 相對倍數調整前撰寫的說明' in w:
                    w = w[:w.index('&"以下為 2026-10-29 相對倍數調整前撰寫的說明')]
                if ws.title == "README" and "重點觀察 1 的第一句已改為公式；重點觀察 2、3 與其他說明中的目標價與上檔空間數字為調整前撰寫，以 Val 頁為準。" in w:
                    w = w.replace("重點觀察 1 的第一句已改為公式；重點觀察 2、3 與其他說明中的目標價與上檔空間數字為調整前撰寫，以 Val 頁為準。",
                                  "重點觀察 1 的第一句已改為公式；第三輪修正（2026-11-06）刪除重點觀察 1 中調整前撰寫的舊說明；重點觀察 2、3 為當時撰寫的說明，目標價與上檔空間以 Val 頁為準。")
                # (15) 情境驅動對照：收斂法列名改中性（apply_in_layout 取情境 C 列名，會帶「（不如預期）」）
                if ws.title == "In" and c.column == 2 and ws.cell(c.row, 1).value == "scn_fade_g" and "收斂目標" in w:
                    w = "FY2032 營收年增收斂目標（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）"
                # (24) 上檔空間級距列名
                if ws.title == "Val" and c.column == 1 and w == "上檔空間級距":
                    w = "上檔空間級距（六級，供分組）"
                    e = ws.cell(c.row, 5)
                    e.value = (str(e.value) + "；" if e.value else "") + "六級只供分組比較；投資建議依下一列的五級固定規則"
                # (16) Sens「g ±1pt」：g 受 rf 上限時列名改「g −1pt／上限 rf」（公式判斷）
                if ws.title == "Sens" and c.column == 1 and w == "永續成長率 g ±1pt":
                    w = '=IF(g_tv_hi<g_tv+0.01-0.000001,"永續成長率 g −1pt／上限 rf","永續成長率 g ±1pt")'
                if w != v:
                    c.value = w; n += 1
    return n
RD["A1"] = RD["A1"].value + "；v0.2.3（2026-11-06）第三輪修正（合併前）：現價與同業報價統一到 2026-11-04 收盤、stub 期 FCFF 改 FY2026 錨點、展示字串統一"
_r3_text(wb, "含 stub 期會計年度 FY2026（日曆年）FCFF × stub_frac", px_old_dates=("2026-10-22",))
wb.save(sys.argv[1])
if os.environ.get("AURAS_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": len(cases), "SENS_LAST": SENS_LAST, "SENSNOTE": SENS_NOTE,
               "VAL": VAL, "IN": IN_ROWS, "WMBT_NOTE": WMBT_NOTE, "RB0": RB0, "BLOCK_END": block_end, "RECON_END": RECON_END, "RECON_B": RECON_B, "UN_NOTE": ur,
               "XT": XT, "XH": XH, "XR": XR, "T": T, "EPSR": EPSR, "CFR": CFR, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "RECV2": RECV2, "V1": V1, "C2": C2, "C3": C3, "C4": C4,
               "PEER_LAST": PEER_LAST, "PXR": PXR, "C03": C03},
              open(os.environ["AURAS_ROWMAP"], "w"), ensure_ascii=False, indent=1)
