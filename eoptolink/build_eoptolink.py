# 新易盛（300502.SZ）收支與估值模型 v0.2.2（AI 半導體 Project：網通次產業內複製——光模組，比照 innolight/）
# 用法：python3 eoptolink/build_eoptolink.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不做平移。金額：十億人民幣（RMB B）；市場：十億美元（US$B，讀 markets/network）；EPS：人民幣。
# 口徑：中國會計準則；淨利＝歸屬母公司股東淨利。公司不給營收或獲利指引；FY2026 錨點＝1H26 實績＋2H26 假設（3Q26 季報預計 2026-10-30 公布，截至 10-25 尚未公布）。
# 股數：FY2025 年度分配 10 轉 4 股（約 9.94 億股 → 13.94 億股），全部期間以轉增後股數表示（FY2025 EPS 9.61 ÷ 1.4 重述）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import fcff_dcf as FD   # v0.4：標準版 DCF（無槓桿自由現金流＋WACC）共用模組
import valuation_rule as VR
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); INL_F = only("innolight/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
NETR = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    lab = r[0].value
    if lab in ("  其中 CPO", "  其中可插拔", "AI IT 資本支出（加速器 × 每顆 IT capex）"):
        NETR[lab.strip()] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert len(NETR) == 3 and all(v is not None for k in NETR for v in NETR[k][0]), NETR.keys()
# 網通 CPO 三情境（markets/network v0.2.3；Andy 2026-10-25 13:13：技術替代屬行業情境，A＝替代快、C＝慢，公司讀同字母）
NETS = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    for s_ in ("A", "B", "C"):
        for k_, l_ in (("mpl", f"  其中可插拔（情境 {s_}）"), ("mcp", f"  其中 CPO（情境 {s_}）")):
            if r[0].value == l_:
                NETS[(k_, s_)] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert len(NETS) == 6 and all(v is not None for k in NETS for v in NETS[k][0]), NETS.keys()
NETR["其中可插拔"] = NETS[("mpl", "A")]; NETR["其中 CPO"] = NETS[("mcp", "A")]   # 情境 A 列（數值＝主表「其中可插拔」「其中 CPO」）
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
INL = {}
for r in openpyxl.load_workbook(INL_F, data_only=True)["Units"].iter_rows(min_row=4):
    if r[0].value in ("中際旭創可插拔有效份額", "中際旭創 AI 營收（情境 A；美元）"):
        INL[r[0].value] = ([r[2 + j].value for j in range(4)], f"{rel(INL_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert len(INL) == 2 and all(v is not None for k in INL for v in INL[k][0]), INL.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。公司金額：十億人民幣（RMB B）；市場：十億美元（US$B）；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝深交所定期報告與公告（經財經媒體轉載）；公司不給營收指引；客戶名稱、客戶集中度、800G／1.6T 份額：未找到第一手來源"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
FY25E = "https://finance.eastmoney.com/a/202604243717729349.html"
FY25N = "https://www.nbd.com.cn/articles/2026-04-24/4356482.html"
HX = "https://www.huxiu.com/article/4853326.html"
H1C = "https://m.c114.com.cn/w123-1316262.html"
H1E = "https://finance.eastmoney.com/a/202608243851382445.html"
XN = "https://stock.10jqka.com.cn/20260915/c679905127.shtml"
CJ = "https://m.gelonghui.com/news/5306856"
KY = "https://pdf.dfcfw.com/pdf/H3_AP202608251828397684_1.pdf"
TF = "https://stock.10jqka.com.cn/20261010/c680549851.shtml"
INC = "https://stock.10jqka.com.cn/20260820/c679144690.shtml"
CONS = "https://basic.10jqka.com.cn/300502/worth.html"
GS = "https://wallstreetcn.com/articles/3777393"
GS25 = "https://www.itiger.com/news/1150980067"
PX1 = "https://quote.cfi.cn/300502.html"
PX2 = "https://www.etnet.com.hk/www/sc/ashares/quote.php?code=300502"
INV = "https://pl.investing.com/equities/eoptolink-tech-earnings"
GF = "https://www.google.com/finance/quote/300502:SHE?hl=zh"
N928 = "https://stock.10jqka.com.cn/20260928/c680329640.shtml"
LC = "https://lightcounting.com/newsletter/en/may-2026-optical-vendor-landscape-376"
CITI = "https://hao.cnyes.com/post/262773"

sec("A. AI 光模組市場（讀 markets/network 成品；情境 A（原始顆數），B／C 依顆數係數縮放）與中際旭創份額（讀 innolight 成品）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mpl_{y}", f"AI 可插拔光模組市場 20{y}", "US$B", None, NETR["其中可插拔"][0][j], None, "Derived", NETR["其中可插拔"][1],
        "加速器顆數 × 每顆光學內容（scale-out 為主；2029 含 scale-up 光學）− CPO" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mcp_{y}", f"AI CPO 光學市場 20{y}", "US$B", None, NETR["其中 CPO"][0][j], None, "Derived", NETR["其中 CPO"][1], "共同封裝光學；網通模型的 CPO 滲透路徑" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][0][j], None, "Derived", NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][1], "Recon 用：新易盛 AI 營收 ÷ AI IT 資本支出" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"inl_{y}", f"中際旭創可插拔有效份額 20{y}（情境 A）", "%", None, INL["中際旭創可插拔有效份額"][0][j], None, "Derived", INL["中際旭創可插拔有效份額"][1],
        "中際旭創份額的權威在 innolight/ Units 頁；本模型只用來檢查兩家合計是否合理（Units 頁）" if y == 26 else "同上")
inp("inl_aiu26", "中際旭創 AI 營收 2026（美元）", "US$B", None, INL["中際旭創 AI 營收（情境 A；美元）"][0][0], None, "Derived", INL["中際旭創 AI 營收（情境 A；美元）"][1], "Recon 用：新易盛 ÷ 中際旭創")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))

sec("B. 新易盛份額與單價（本專案為權威：新易盛在 AI 光模組的份額，規則第 13 條）")
inp("rel_27", "可插拔份額 2027 ÷ 2026（相對值）", "x", 0.80, 1.00, 1.15, "Assumed", LC + "：LightCounting 2025 排名新易盛第二（US$3.5B，+189%，首次超過 Coherent）（Interested-party）",
    "中際旭創模型同年相對值 0.95（讓出份額給新易盛、Coherent）；新易盛持平時兩家合計仍在 65% 左右（Units 頁）")
inp("rel_28", "可插拔份額 2028 ÷ 2026", "x", 0.65, 1.00, 1.20, "Assumed", "同上；1.6T 份額仍在爬坡（西南證券 9-15，Interested-party）", "主要客戶轉單或 1.6T 落後為低值；美國限制另列")
inp("rel_29", "可插拔份額 2029 ÷ 2026", "x", 0.55, 0.95, 1.20, "Assumed", "同上；2029 起 scale-up 光學與 CPO 比重上升", "—")
inp("us_27", "美國限制情境：份額相對值 2027（Sens 用）", "x", None, 0.90, None, "Assumed", CITI + "：花旗認為短期嚴格執行機率低；中國廠供美國大型雲端高速光模組約 60–70%（Interested-party）",
    "新易盛境外營收占 97.92%（1H26，西南證券轉引公司）；泰國產能不必然豁免（美國價值含量門檻）")
inp("us_28", "美國限制情境：份額相對值 2028（Sens 用）", "x", None, 0.55, None, "Assumed", "同上；與中際旭創模型相同", "—")
inp("us_29", "美國限制情境：份額相對值 2029（Sens 用）", "x", None, 0.50, None, "Assumed", "同上", "—")
inp("cpo_27", *SIL.row9("cposh_eoptolink_27")[1:])
inp("cpo_28", *SIL.row9("cposh_eoptolink_28")[1:])
inp("cpo_29", *SIL.row9("cposh_eoptolink_29")[1:])
inp("pxr_27", "單價係數 2027（相對網通模型每顆內容；基準 1）", "x", 0.90, 1.0, 1.05, "Assumed", "網通模型每顆光學內容已含世代單價；本係數只在 Sens 動（降價快於模型）", "與中際旭創模型相同")
inp("pxr_28", "單價係數 2028", "x", 0.80, 1.0, 1.05, "Assumed", "同上", "—")
inp("pxr_29", "單價係數 2029", "x", 0.72, 1.0, 1.05, "Assumed", "同上", "—")
inp("fx", "人民幣兌美元（RMB／USD）", "RMB", *SIL.vals("fx_rmb"), "Verified", "Siliconomics SIF_fx_rmb（中國經濟網 2026-10-16 中間價 6.7220；v0.8.2 前 7.1）", "1Q26 財務費用 5.22 億元，主因匯兌損失（每日經濟新聞轉載公司一季報）")

sec("C. 其他營收（國內、電信、非 AI 數通；網通模型未涵蓋）")
inp("oth_s", "其他營收 ÷ 總營收（FY2025、FY2026）", "%", 0.03, 0.05, 0.10, "Assumed", "西南證券 9-15（網址見 Sources）：境外收入占比約 97.92%（1H26）", "光互聯產品占營收 99.9%（1H26 208.83 ÷ 209.10）；非 AI 部分比重：未找到")
inp("g_oth", "其他營收年增（FY2027 起）", "%", 0.0, 0.10, 0.20, "Assumed", "與中際旭創模型相同", "—")

sec("D. 毛利率")
inp("gm_h2", "4Q26 毛利率（3Q26 已為實績，約 46.3%）", "%", 0.44, 0.465, 0.48, "Assumed", CJ + "：2Q26 綜合毛利率 47.96%（季減 1.20pt）→ 1Q26 49.16%（推導）", XN + "：1H26 綜合毛利率 48.44%")
inp("gm_27", "毛利率 2027", "%", 0.43, 0.47, 0.49, "Assumed", "1.6T 占比上升（上推）vs 降價（下壓）；FY2024 光互聯 44.85%、FY2025 47.81%（公司）", "開源證券 2026E–2028E 47.0%（Interested-party）")
inp("gm_28", "毛利率 2028", "%", 0.40, 0.46, 0.48, "Assumed", "同上；與中際旭創模型路徑相同", "—")
inp("gm_29", "毛利率 2029", "%", 0.36, 0.45, 0.47, "Assumed", "同上", "—")

sec("E. FY2026 錨點（1H26 實績＋2H26 假設）")
inp("h1_rev", "1H26 營收", "RMB B", None, 20.910, None, "Verified", H1C + "：209.1 億元（+100.34%）", H1E + "：同值")
inp("h1_gm", "1H26 綜合毛利率", "%", None, 0.4844, None, "Verified", XN + "：綜合毛利率 48.44%（西南證券轉引公司半年報）", H1C + "：光互聯產品 48.46%（占營收 99.9%）")
inp("h1_ni", "1H26 歸母淨利", "RMB B", None, 7.529, None, "Verified", H1C + "：75.29 億元（+90.98%）", H1E + "：同值")
inp("h1_eps", "1H26 基本 EPS（轉增後）", "RMB", None, 5.41, None, "Verified", H1E + "：5.41 元（+90.49%）", "https://quartr.com/events/eoptolink-technology-inc-ltd-300502-q2-2026_FgrWDhx6：5.41（去年同期 2.84）")
inp("q1_rev", "1Q26 營收", "RMB B", None, 8.338, None, "Verified", FY25N + "：約 83.4 億元（+105.76%）", "1H26 209.10 − 2Q26 125.72 ＝ 83.38（長江證券轉引公司）")
inp("q_g", "4Q26 營收季增（3Q26 已為實績）", "%", 0.10, 0.20, 0.40, "Assumed", "2Q26 ÷ 1Q26 − 1 ＝ +50.8%（公司）；1.6T 下半年放量節奏加快（公司說法，同花順 9-28）", INV + "：3Q26 共識營收 177.6 億元（隱含季增約 +41%，Interested-party）；季報預計 2026-10-30 公布")
inp("opx_h2", "2H26 營業費用、減值與其他 ÷ 營收", "%", 0.04, 0.05, 0.07, "Assumed", "2Q26 殘差約 4.9%（Q 頁推導；1Q26 含匯兌損失 5.22 億元，約 9.8%）", KY + "：1H26 銷售 0.48%、管理 0.62%、研發 2.1%（開源證券）")
inp("tx", "有效稅率", "%", 0.08, 0.12, 0.15, "Assumed", "成都母公司高新技術企業 15%；泰國子公司稅率與獲利占比：未找到", "低於 15% 的依據：FY2025 若用 15%，殘差營業費用（2.6%）低於研發費用率 2.83%，不合理（Q 頁）")
inp("attr", "歸母淨利 ÷ 淨利（含少數股權）", "x", 0.99, "=9.532/(9.532+0.021)", 1.0, "Derived", KY + "：FY2025 少數股東損益 0.21 億元（開源證券財務摘要）", "歸母 95.32 億元（公司）")
inp("sh_post", "轉增後總股數", "十億股", None, 1.394, None, "Verified", KY + "：總股本 13.94 億股（開源證券 8-25）", PX1 + "：總股本 13.95 億股；10 轉 4 基數約 9.94 億股（每日經濟新聞）")
inp("q_sh_h1", "1H26 加權股數", "十億股", None, "=h1_ni/h1_eps", None, "Derived", "歸母淨利 ÷ 基本 EPS", "—")

sec("F. FY2025 實績")
inp("a25_rev_in", "FY2025 營收", "RMB B", None, 24.842, None, "Verified", FY25E + "：248.42 億元（+187.29%）", FY25N + "：約 248.4 億元")
inp("a25_gm_in", "FY2025 毛利率（光互聯產品代用，占營收 99.7%）", "%", None, 0.4781, None, "Verified", "每日經濟新聞轉載公司年報（網址見 Sources）：光互聯產品毛利率 47.81%（2024 年 44.85%）", "虎嗅（網址見 Sources）：同值；全公司毛利率：未找到")
inp("a25_ni_in", "FY2025 歸母淨利", "RMB B", None, 9.532, None, "Verified", FY25E + "：95.32 億元（+235.89%）", FY25N + "：同值")
inp("a25_eps_in", "FY2025 基本 EPS（轉增前）", "RMB", None, 9.61, None, "Verified", FY25E + "：9.61 元", FY25N + "：同值")
inp("a25_sh_in", "FY2025 股數（轉增後重述）", "十億股", None, "=a25_ni_in/(a25_eps_in/1.4)", None, "Derived", "歸母淨利 ÷（EPS ÷ 1.4）；10 轉 4（FY2025 年度分配）", "重述 EPS ＝ 9.61 ÷ 1.4 ＝ 6.86")

sec("G. 營業費用、股數（FY2027 起）")
inp("phi", "營業費用（含減值與其他）成長對營收成長的彈性", "x", 0.30, 0.60, 1.00, "Analogy", "中際旭創模型 0.61（Derived）；本公司 FY2026 推導值（Q 頁 B26）因 1Q26 匯兌損失失真，不採用", "FY2025 研發 7.02 億元、占營收 2.83%（公司）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.02, "Assumed", "無", "Sens 用 +2 個百分點")
inp("sh_g", "股數年變動", "%", -0.005, 0.0, 0.01, "Assumed", INC + "：2026 年限制性股票 648.12 萬股（0.46%），授予價 212 元", "考核：2026 年營收 ≥ 500 億元、2026–27 累計 ≥ 1,400 億、2026–28 累計 ≥ 2,900 億（公司公告，Interested-party）")
inp("mf", "AI 市場係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年可插拔光模組 +84%", "±20%")

sec("H. 共識、券商、市場資料（只作 Recon 對帳，不作校準）")
for y, rv, ni, eps in ((26, 55.189, 20.135, 15.05), (27, 99.385, 36.358, 27.29), (28, 143.815, 53.635, 39.93)):
    inp(f"cons_rev{y}", f"共識 FY20{y} 營收", "RMB B", None, rv, None, "Interested-party", CONS + "（2026-10-23；6 個月內 19 家，明細只列 10 家）", "區間見同頁")
    inp(f"cons_ni{y}", f"共識 FY20{y} 歸母淨利", "RMB B", None, ni, None, "Interested-party", CONS + "（19 家）", "天風（10-10）、國金（8-25）兩家特別高，拉高均值")
    inp(f"cons_eps{y}", f"共識 FY20{y} EPS", "RMB", None, eps, None, "Interested-party", CONS + "（19 家）", "—")
for y, ni in ((26, 20.683), (27, 30.080), (28, 34.576)):
    inp(f"gs_ni{y}", f"高盛 FY20{y} 歸母淨利（單一券商）", "RMB B", None, ni, None, "Interested-party", GS + "：高盛（2026-07）2026–28 年 206.83、300.80、345.76 億元；目標價 633 元（29.3x 2027 年獲利）", "共識均值明顯高於高盛（FY2027 +21%、FY2028 +55%）")
inp("gs_tp", "高盛 12 個月目標價（2026-07）", "RMB", None, 633.0, None, "Interested-party", GS + "：600.71 → 633.0 元", GS25 + "：高盛稱 2018 年以來平均前瞻 P/E 約 29x")
inp("q3_cons_rev", "3Q26 共識營收", "RMB B", None, 17.76, None, "Interested-party", INV + "：17.76bn（季報 2026-10-30）", "—")
inp("lc_eo25", "LightCounting 2025 新易盛光模組營收", "US$B", None, 3.5, None, "Interested-party", LC + "：US$3.5B（+189%，排名第二）", "中際旭創 US$5.3B")
inp("tp_avg90", "同花順 90 天內機構目標均價", "RMB", None, 460.89, None, "Interested-party", XN + "：15 家（買入 13、增持 2），目標均價 460.89 元", "多為 9 月前股價較高時給出")

sec("I. 價格與估值")
inp("px", "股價（2026-10-30 收盤，A 股）", "RMB", None, 196.00, None, "Interested-party", "同花順、stockanalysis（markets/invest_map/data/prices.csv）：196.00（10-23 為 261.00）", PX2 + "：同值")
inp("px_hi", "52 週高點", "RMB", None, 818.38, None, "Interested-party", PX2 + "：52 週區間上緣 818.38", GF + "：另一資料商顯示 618.87（9-30），兩者不一致：未核實；2026-06-10 收盤 772.50（融中財經）")
inp("px_0930", "股價（2026-09-30 收盤）", "RMB", None, 389.00, None, "Interested-party", "Google 財經（網址見 Sources）：389.00", "近一個月跌幅原因：未找到")
for s, w in zip(SC, (0.3, 0.5, 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

sec("K. 目標價（DCF＋Comparable）與收斂法參數（v0.2 新增；黃底，Assumed，Andy 可改）")
inp("tp_g", "收斂法情境 B 的 FY2032 營收年增（v0.4 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.4 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 收斂路徑（fade_g）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "收斂法：情境 A 於 FY2032 的營收年增（＝tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法；FY2030 起由 FY2029 年增線性收斂", "FY2030 起為收斂法，非驅動因子推導")
inp("fade_gB", "收斂法：情境 B 於 FY2032 的營收年增（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "收斂法：情境 C 於 FY2032 的營收年增（＝tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "收斂法：情境 A 的中段淨利率（＝情境 A FY2029 淨利率）", "%", None, None, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "收斂法：情境 B 的中段淨利率（＝FY2029 與十年平均的中點；十年平均以 FY2029 × 0.8 代替）", "%", None, None, None, "Assumed", "十年平均淨利率（2016–2025，與本模型同口徑）：未找到，依工作單以 FY2029 × 0.8 代替；本列＝（FY2029＋FY2029 × 0.8）÷ 2", "Andy 可改")
inp("fade_mC", "收斂法：情境 C 的中段淨利率（＝十年平均；以情境 C FY2029 × 0.8 代替）", "%", None, None, None, "Assumed", "十年平均淨利率：未找到，依工作單以情境 C FY2029 淨利率 × 0.8 代替", "Andy 可改")
inp("cons_tp", "市場共識目標價（賣方平均）", 'RMB', None, '=tp_avg90', None, "Interested-party", '同花順 90 天內機構目標均價（In 頁 tp_avg90；15 家，10 轉 4 後口徑）', 'stockanalysis 預測頁平均 652.48（9 位；口徑疑為轉增前，不採）；多數目標價訂於 9 月前股價較高時')
inp("cons_feps", '共識 12 個月後前瞻 EPS（2/12 × FY2027＋10/12 × FY2028）', 'RMB', None, '=2/12*cons_eps27+10/12*cons_eps28', None, "Derived", 'In 頁同花順共識 FY2027、FY2028 EPS（與本模型前瞻 EPS 同期間）', "共識只放 Recon 與本列，不作校準")
# v0.4：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py）。評價日沿用 v0.2 的 2026 年底；股價 2026-10-30、rf 2026-11-05；資產負債為 2026-06-30（合併）
Q3 = "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）"
SAQ = "stockanalysis.com 300502（SHE）財報頁（資料商整理公司中國會計準則合併財報；2026-11-05 讀取）"
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.2 DCF 基準日）", ccy="RMB", amt="RMB B", shu="十億股",
        rf=(0.0161, "Trading Economics 中國 10 年期國債殖利率：2026-11-05 1.61%（1.6080；https://tradingeconomics.com/china/government-bond-yield；2026-11-05 讀取）",
            "同頁：11-02 約 1.60%、10-19～20 約 1.66%（價格日 10-30 約 1.6%）", "無風險利率（人民幣 10 年期國債殖利率，2026-11-05）"),
        rf_us=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）", "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）：10-28～10-30 5.05%／4.98%／4.94%；只用於 Sens「美元口徑折現」（Andy 2026-11-05 23:16 定案）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：CN 光通訊 同業 6 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 上證綜指 000001.SS 迴歸，Hamada 去槓桿、現金修正）；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：βL＝βU ×（1＋(1 − tax_w) × D/E）；Andy 2026-11-05 21:21）",
                    "低／高＝β_u ∓ beta_du（0.3），各自同式重新加槓桿（Analogy；工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得"),
        erp=(0.04, 0.05, 0.06, "工作單 2026-11-05 預設 5%（成熟市場）；中國另加 CRP"),
        crp=(0.0091, "Interested-party", "Damodaran Country Risk Premium（2026-01-05 更新）：China Moody's A1、違約利差 0.60%、CRP 0.91%、總 ERP 5.14%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）",
             "2026-07 更新檔（ctrypremJuly26.xlsx）在建置環境無法讀取；Andy 2026-11-05 20:45：非美國 ERP＝5%＋CRP、不設股權成本下限"),
        kd=("=rf+0.01", "Assumed", "無金融借款（資料商頁面 Total Debt 只有租賃）；kd 只用於租賃", "負債占資本 < 0.1%，對 WACC 影響可忽略"),
        tax_w=("=tx", "Assumed", "＝In 頁 tx（12%）", "NOPAT 亦用 tx"),
        debt=(0.0, "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 短期借款、長期借款皆未列示", "一年內到期的非流動負債 8.08（RMB M）為租賃，列入下一列"),
        lease=("=0.008081+0.030741", "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 一年內到期的非流動負債 8.08＋租賃負債 30.74（RMB M）", "2026-06-30 為 5.72＋18.11"),
        cash=("=6.576626+0.272152", "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 貨幣資金 6,576.63＋交易性金融資產 272.15（RMB M）", "2026-06-30 貨幣資金 7,240；第三季付 FY2025 股利 995.9 已反映"),
        inv=("=0.013766+0.028499", "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 長期股權投資 13.77＋其他非流動金融資產 28.50（RMB M；帳面）", "Sens 測不計"),
        mi=(0.225555, "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 少數股東權益 225.55（RMB M；帳面）", "子公司未上市，用帳面"),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=Q3,
                  div=(0, "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：未宣告新的利潤分配或回購（第三季已付 FY2025 股利 995.9 RMB M，已反映在 9-30 現金）", "—")),
        dps=("=0.714*a26_ni/a25_ni", "Assumed", "FY2025 每股股利 0.714 元（stockanalysis，轉增後口徑；2026 年發放）", "評價日到目標時點之間發 FY2026 股利一次；假設配發率不變：0.714 × FY2026 歸母淨利 ÷ FY2025"),
        sh=(1.394737, "Verified", "新易盛 2026 年第三季度報告（2026-10-29 董事會；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF）：2026-09-30 股本 1,394,737,276 股（無庫存股）", "v0.4.2 前＝sh_post（13.94 億股）"),
        fc_da=(None, "=(129.5/3098+190.22/8647+369.09/24842)/3", None, "Interested-party",
               SAQ + "：折舊攤銷 ÷ 營收，FY2023 129.5 ÷ 3,098＝4.18%、FY2024 190.2 ÷ 8,647＝2.20%、FY2025 369.1 ÷ 24,842＝1.49%，平均", "—"),
        fc_cx=(None, "=(554.31/3098+1476/8647+1320/24842)/3", None, "Interested-party",
               SAQ + "：資本支出 ÷ 營收，FY2023 554.3 ÷ 3,098＝17.9%、FY2024 1,476 ÷ 8,647＝17.1%、FY2025 1,320 ÷ 24,842＝5.31%，平均", "Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=1320/24842", "Assumed", "長期穩態：最近一年（FY2025）5.31%（營收三年成長 8 倍，三年平均高估往後比例；FY2028 起線性收斂）", "≥ 折舊 ×（1＋g）（DCF 頁一致性檢查）"),
        fc_wc=(0.25, "=((753.1+963.35+4.08+99.57-632.56-81.47-13.75-71.94)/3098+(2620+4132+0.03+572.82-2844-557.34-9.74-49.36)/8647+(4597+7234+0.03+528.65-5275-657.81-90.7-87.12)/24842)/3", 0.45, "Interested-party",
               SAQ + "：營運資金＝應收＋存貨＋預付＋其他流動資產 − 應付 − 應計費用 − 預收 − 其他流動負債；÷ 營收 FY2023 32.9%、FY2024 44.7%、FY2025 25.2%，平均",
               "Sens 低／高 25%／45%（Assumed）"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（自算 CN 光通訊，Blume 調整後；採用值）", "x", None, "=0.67*beta_u_med+0.33", None, "Derived",
    "自算產業 β：CN 光通訊 同業 6 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 上證綜指 000001.SS 迴歸，Hamada 去槓桿、現金修正）；＝0.67 × 中位數 beta_u_med ＋ 0.33", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產")
inp("beta_u_med", "產業無槓桿 β 中位數（自算 CN 光通訊，Blume 調整前；β 低／高案例用）", "x", None, 2.69, None, "Derived",
    "自算產業 β：CN 光通訊 同業 6 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 上證綜指 000001.SS 迴歸，Hamada 去槓桿、現金修正）", "β_u＝0.67 × 本值＋0.33；β 低／高案例改用 β_u ∓ beta_du（v0.4.2 起）")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代原規則（Blume 前中位數 × 0.8 再 Blume／中位數未調整）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑")

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
INROW = {}
for rr in range(4, r):
    if IN.cell(rr, 3).value is not None: INROW[IN.cell(rr, 1).value] = rr

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億人民幣；EPS 為人民幣；股數為十億股，轉增後）。藍字＝輸入；黑字＝公式。口徑：歸母淨利；FY2025 毛利率以光互聯產品（占營收 99.7%）代用"
Q["A2"] = "FY2026 錨點＝1H26 實績＋2H26（3Q、4Q 營收依季增假設；毛利率 gm_h2；營業費用率 opx_h2）。營業費用、減值與其他為殘差：毛利 − 歸母淨利 ÷［（1 − 稅率）× 歸母比例］。3Q26 季報預計 2026-10-30 公布（截至 10-25 尚未公布），公布後以實績取代 3Q"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（實績）", "4Q26（假設）", "2H26 合計", "FY2026 錨點", "標記", "說明"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", note=""):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, note)
qrow(4, "營收", {"B": "=a25_rev_in", "C": "=h1_rev", "D": "=q3_rev", "E": "=D4*(1+q_g)", "F": "=D4+E4", "G": "=C4+F4"}, tag="Verified／Assumed（4Q）", note="3Q26＝季報實績（v0.3）；4Q＝3Q ×（1＋季增）。原：2Q26＝1H26 − 1Q26（125.72）；3Q＝2Q ×（1＋季增）")
qrow(5, "毛利", {"B": "=B4*a25_gm_in", "C": "=C4*h1_gm", "D": "=q3_gp", "E": "=E4*gm_h2", "F": "=D5+E5", "G": "=C5+F5"}, tag="Derived")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(7, "稅後歸母係數 ＝（1 − 稅率）× 歸母比例", {c: "=(1-tx)*attr" for c in "BCDE"}, fmt="0.000", tag="Assumed／Derived")
qrow(8, "營業費用、減值與其他（淨）", {"B": "=B5-B10/B7", "C": "=C5-C10/C7", "D": "=D5-D10/D7", "E": "=E4*opx_h2", "F": "=D8+E8", "G": "=C8+F8"}, tag="Derived（殘差）／Assumed（2H）",
     note="1H26 殘差含 1Q26 匯兌損失（財務費用 5.22 億元）")
qrow(9, "營業費用率", {c: f"={c}8/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(10, "歸母淨利", {"B": "=a25_ni_in", "C": "=h1_ni", "D": "=q3_ni", "E": "=(E5-E8)*E7", "F": "=D10+E10", "G": "=C10+F10"}, tag="Verified／Derived（3Q、4Q）")
qrow(11, "淨利率", {c: f"={c}10/{c}4" for c in "BCDEFG"}, fmt="0.0%", tag="Derived")
qrow(12, "股數（十億股，轉增後）", {"B": "=a25_sh_in", "C": "=q_sh_h1", "D": "=sh_post", "E": "=sh_post", "G": "=(C12*2+D12+E12)/4"}, tag="Derived", note="FY2026＝1H 加權 × ½ ＋ 期末股數 × ½（近似）")
qrow(13, "EPS", {"B": "=B10/B12", "C": "=C10/C12", "D": "=D10/D12", "E": "=E10/E12", "G": "=G10/G12"}, fmt="0.00", tag="Derived", note="FY2025 轉增後重述 ＝ 9.61 ÷ 1.4；1H26 公司揭露 5.41")
qrow(14, "其他營收（國內、電信等；＝營收 × oth_s）", {"B": "=B4*oth_s", "G": "=G4*oth_s"}, tag="Assumed")
qrow(15, "AI 營收（＝營收 − 其他）", {"B": "=B4-B14", "G": "=G4-G14"}, tag="Derived")
qrow(16, "AI 營收（US$B ＝ ÷ fx）", {"B": "=B15/fx", "G": "=G15/fx"}, tag="Derived")
Q.cell(18, 1, "季度拆分（推導）").font = BOLD
Q.cell(19, 1, "2Q26 營業費用、減值與其他 ÷ 營收（殘差；2Q26 毛利率 47.96%、歸母淨利 47.49 億元）"); Q["B19"] = "=(12.572*0.4796-4.749/((1-tx)*attr))/12.572"; Q["B19"].number_format = "0.0%"
Q.cell(20, 1, "1Q26 營業費用、減值與其他 ÷ 營收（殘差；1Q26 毛利率 49.16%、歸母淨利 27.80 億元）"); Q["B20"] = "=(q1_rev*0.4916-(h1_ni-4.749)/((1-tx)*attr))/q1_rev"; Q["B20"].number_format = "0.0%"
Q.cell(21, 1, "FY2026 錨點與共識的距離（參考；校準未使用共識）").font = BOLD
Q.cell(22, 1, "FY2026 營收 ÷ 共識 − 1"); Q["B22"] = "=G4/cons_rev26-1"
Q.cell(23, 1, "FY2026 歸母淨利 ÷ 共識 − 1"); Q["B23"] = "=G10/cons_ni26-1"
Q.cell(24, 1, "FY2026 營收年增"); Q["B24"] = "=G4/B4-1"
Q.cell(25, 1, "FY2026 營業費用年增"); Q["B25"] = "=G8/B8-1"
Q.cell(26, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；受 1Q26 匯兌損失影響，In 頁 phi 不採用）"); Q["B26"] = "=B25/B24"
Q.cell(27, 1, "1H26 EPS（模型）vs 公司揭露 5.41"); Q["B27"] = "=C13/h1_eps-1"
Q.cell(28, 1, "FY2026 營收 vs 股權激勵第一期考核 500 億元（÷ 50 − 1）"); Q["B28"] = "=G4/50-1"
for rr in range(22, 29): Q[f"B{rr}"].number_format = "0.0%" if rr != 26 else "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 8), ("ni", 10), ("sh", 12), ("eps", 13), ("oth", 14), ("ai", 15), ("aiu", 16)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 80
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
mk = lambda b, y, s: f"{b}_{y % 100}" if s == "A" else f"{b}{s}_{y % 100}"   # 網通 CPO 三情境：情境 A 讀 A 節，B、C 讀同字母列
_mfy = lambda m, y: m[y] if isinstance(m, dict) else m   # Siliconomics v0.13：mf 可為年別 dict（Sens「Wattonomics 錨定」案例）；預設字串不變
DEFP = dict(scn="A", rel="rel", cpo="cpo", pxr="pxr", gm="gm", mf="1", fx="fx", gn="g_oth", opxa="0", phi="phi")
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
    F = BU[1:]
    sec_("A. AI 營收（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029）")
    row("mpl", "AI 可插拔光模組市場 × 情境顆數係數", "US$B", {c_(y): f"={mk('mpl', y, s)}*{umx(y, s)}" + ("" if y == 2026 else f"*{_mfy(P['mf'], y)}") for y in BU})
    row("mcp", "AI CPO 光學市場 × 情境顆數係數", "US$B", {c_(y): f"={mk('mcp', y, s)}*{umx(y, s)}" + ("" if y == 2026 or not isinstance(P['mf'], dict) else f"*{P['mf'][y]}") for y in BU})
    row("fxr", "人民幣兌美元", "RMB", {c_(y): ("=fx" if y == 2026 else f"={P['fx']}") for y in BU}, fmt="0.00",
        note="FY2026 校準用基準匯率；Sens 只動 FY2027 起")
    row("sh0", "新易盛可插拔有效份額（FY2026 校準）＝ AI 營收（美元）÷ 可插拔市場", "%", {"D": f"=a26_aiu/D{R['mpl']}"}, fmt="0.0%",
        note="Derived：FY2026 CPO 營收視為 0；FY2027 起乘相對值")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in F}}, fmt="0.00")
    row("sh", "可插拔有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("pxr", "單價係數（相對網通模型）", "x", {"D": "=1", **{c_(y): f"={P['pxr']}_{y % 100}" for y in F}}, fmt="0.00")
    row("cps", "新易盛占 CPO 光學金額", "%", {"D": "=0", **{c_(y): f"={P['cpo']}_{y % 100}" for y in F}}, fmt="0.0%")
    row("aiu", "AI 營收（美元）", "US$B", {"D": "=a26_aiu", **{c_(y): f"={c_(y)}{R['mpl']}*{c_(y)}{R['sh']}*{c_(y)}{R['pxr']}+{c_(y)}{R['mcp']}*{c_(y)}{R['cps']}" for y in F}})
    row("ai", "AI 營收", "RMB B", {"C": "=a25_ai", **{c_(y): f"={c_(y)}{R['aiu']}*{c_(y)}{R['fxr']}" for y in BU}})
    sec_("B. 其他營收與合計")
    row("oth", "其他營收（國內、電信等）", "RMB B", {"C": "=a25_oth", "D": "=a26_oth", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']})" for y in F}})
    row("rev", "營收", "RMB B", {"C": "=a25_rev", **{c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['oth']}" for y in BU}})
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    sec_("C. 損益")
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", "D": "=a26_gp/a26_rev", **{c_(y): f"={P['gm']}_{y % 100}" for y in F}}, fmt="0.0%",
        note="FY2027 起為輸入路徑（1.6T 占比上推 vs 降價下壓）")
    row("gp", "毛利", "RMB B", {c_(y): f"={c_(y)}{R['rev']}*{c_(y)}{R['gm']}" for y in FYS})
    row("opb", "營業費用、減值與其他（基準路徑）", "RMB B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in F}})
    row("opx", "營業費用、減值與其他", "RMB B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in F}})
    row("pre", "稅前利益（＝毛利 − 營業費用等）", "RMB B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS})
    row("prem", "稅前利益率", "%", {c_(y): f"={c_(y)}{R['pre']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("ni", "歸母淨利", "RMB B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"={c_(y)}{R['pre']}*(1-tx)*attr" for y in F}})
    row("nim", "淨利率", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("sh_", "股數（轉增後）", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=sh_post*(1+sh_g)", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in (2028, 2029)}}, fmt="0.000")
    row("eps", "EPS", "RMB", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 新易盛損益（RMB B）。FY2025A 實績、FY2026 錨點（1H26 實績＋2H26 假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026 三情境相同；FY2027 起三情境讀不同的網通 CPO 情境（A 替代快、C 慢），顆數係數從 CY2028 開始"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：原始顆數（顆數係數 1.0；v0.12 起本庫基準＝情境 B）；網通 CPO 情境 A（替代快）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85；網通 CPO 情境 B", "C": "情境 C：CY2028–29 顆數 ×0.6；網通 CPO 情境 C（替代慢）"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "RMB B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 營收", "RMB B", "#,##0.0"), ("oth", "  其他營收", "RMB B", "#,##0.0"),
                                ("aish", "  AI ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("nim", "淨利率", "%", "0.0%"), ("ni", "歸母淨利", "RMB B", "#,##0.0"),
                                ("eps", "EPS", "RMB", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收、AI 營收、歸母淨利 − 錨點（校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（RMB B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：AI 營收差（RMB B）"); PL[f"D{sr}"] = f"=D{PB['A']['ai']}-a26_ai"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：稅前利益 ×（1 − 稅率）× 歸母比例 − 歸母淨利錨點"); PL[f"D{sr}"] = f"=D{PB['A']['pre']}*(1-tx)*attr-a26_ni"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：市場與份額（權威） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 新易盛在 AI 光模組的份額（本專案權威，規則第 13 條）。市場取自 markets/network（加速器顆數 × 每顆光學內容；美元）；中際旭創份額取自 innolight 成品（唯讀）"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("mpl", "AI 可插拔光模組市場（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mpl']}", "#,##0.0", "markets/network Mkt 頁"),
        ("mcp", "AI CPO 光學市場（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mcp']}", "#,##0.0", "同上"),
        ("sh", "新易盛可插拔有效份額", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "FY2026 以 1H26 實績＋2H26 假設校準"),
        ("inl", "中際旭創可插拔有效份額（innolight 成品，情境 A）", "%", lambda y: f"=inl_{y % 100}", "0.0%", "唯讀；中際旭創份額的權威在 innolight/"),
        ("tot", "兩家合計可插拔份額", "%", lambda y: f"={cyc(y)}{ur_sh}+{cyc(y)}{ur_inl}", "0.0%", "其餘為 Coherent、光迅、Lumentum、華工等"),
        ("aiu", "新易盛 AI 營收（情境 A；美元）", "US$B", lambda y: f"=PL!{FC[y]}{A_['aiu']}", "#,##0.00", ""),
        ("oth", "其他廠商（Coherent 等）＝ 可插拔 ×（1 − 兩家合計份額）", "US$B", lambda y: f"={cyc(y)}{ur_mpl}*(1-{cyc(y)}{ur_tot})", "#,##0.0", "不含 CPO"),
        ("itr", "新易盛 AI 營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['aiu']}/itc_{y % 100}", "0.00%", "markets/network 的 AI IT 資本支出（加速器 × 每顆 IT capex）")):
    UR[key] = ur
    ur_mpl, ur_sh, ur_inl, ur_tot = UR.get("mpl"), UR.get("sh"), UR.get("inl"), UR.get("tot")
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：FY2026 有效份額由 1H26 實績＋2H26 假設校準；與中際旭創合計約三分之二，其餘由 Coherent 等分食。份額相對值與 CPO 份額是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（人民幣；A 股）。目標價＝DCF 與 Comparable 加權（第 0 節）；舊法規則值情境加權保留參考；另列「現價 ÷ FY2027、FY2028 EPS（模型）」、市場隱含 EPS。十年 P/E 用 2016–2025（兩個資料商來源）"
V["A2"] = "循環位置：新易盛不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照中際旭創、Arista；FY(t)＝CY(t)，不平移）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
# 1. 價格（第 0 節目標價放最上方，見本段之後）
V.cell(14, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 15, ["項目", "A", "B", "C", "說明"])
V.cell(16, 1, "現價（2026-10-30 收盤）"); V["B16"] = "=px"; V["B16"].number_format = "0.00"
V.cell(17, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(18, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(19, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(20, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((17, 2026), (18, 2027), (19, 2028), (20, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(21, 1, "現價 ÷ 共識未來 12 月 EPS（2/12 × FY2026＋10/12 × FY2027）"); V["B21"] = "=px/(cons_eps26*2/12+cons_eps27*10/12)"; V["B21"].number_format = "0.0"
V.cell(22, 1, "現價 ÷ 52 週高點 − 1"); V["B22"] = "=px/px_hi-1"; V["B22"].number_format = "0%"
V.cell(23, 1, "現價 ÷ 近 12 月 EPS（2H25＋1H26；推導）"); V["B23"] = "=px/((a25_ni-h1_ni/1.9098+h1_ni)/sh_post)"; V["B23"].number_format = "0.0"
V.cell(24, 1, "現價 ÷ 2026-09-30 收盤 − 1"); V["B24"] = "=px/px_0930-1"; V["B24"].number_format = "0%"
V["E18"] = "FY2027 三情境只差在網通 CPO 情境（顆數係數從 CY2028 開始）"; V["E22"] = "52 週高點兩個資料商不一致（In 頁 px_hi）"; V["E23"] = "1H25 歸母淨利＝1H26 ÷（1＋90.98%）"
V.cell(26, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, 27, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 28):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(32, 1, "3. 規則值隱含股價（人民幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, 33, ["情境", "FY2027 EPS", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = 34 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=D{28 + j}").number_format = "0.0"
    V.cell(i, 4, f"=B{i}*C{i}").number_format = "#,##0"
    V.cell(i, 5, f"=D{i}/px-1").number_format = "0%"
    V.cell(i, 6, f"=px/B{i}").number_format = "0.0"
    for c in range(1, 7): V.cell(i, c).fill = KEY
    VAL[s] = i
V.cell(38, 1, "4. 核心問題：市場隱含的 FY2028 EPS（＝現價 ÷ 假設倍數；約 2027 年底的價格看下一年 EPS）與模型的距離").font = BOLD
hdr(V, 39, ["假設倍數（÷ FY2028 EPS）", "市場隱含 FY2028 EPS", "模型 FY2028 EPS（A）", "隱含 ÷ 模型 A − 1", "模型 FY2028 EPS（C）", "隱含 ÷ 模型 C − 1"])
vr = 40; IMPL = {}
for m in (8, 10, 15, 20, "=pe_peak"):
    V.cell(vr, 1, m).number_format = "0.0"
    if not isinstance(m, str): V.cell(vr, 1).font = BLUE
    V.cell(vr, 2, f"=px/A{vr}").number_format = "0.00"
    V.cell(vr, 3, f"=PL!F{SUMR[('A', 'eps')]}").number_format = "0.00"
    V.cell(vr, 4, f"=B{vr}/C{vr}-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[('C', 'eps')]}").number_format = "0.00"
    V.cell(vr, 6, f"=B{vr}/E{vr}-1").number_format = "0%"
    IMPL[m] = vr; vr += 1
V.cell(vr, 1, "讀法：第五列＝十年 P/E 25 分位（見頂倍數）。現價以 10 倍看 FY2028，隱含 EPS 約比模型 A 低一成；以 15 倍看則約等於情境 C（顆數 ×0.6）——要成立，需要份額、單價、美國限制其中至少一項大幅不利（見 Sens），或 CY2028 光模組需求明顯下修。")
READ_ROW = vr
vr += 2
V.cell(vr, 1, "本版未做先驗：無 Andy 合理 P/E 對照。"); NOPRIOR_ROW = vr
assert vr < 50, vr
PE_CMC = [110, 82.4, 194, 59.2, 54.5, 40.8, 18.6, 70.8, 41.9, 65.0]   # companiesmarketcap 年底 P/E（2016–2025）
PE_SA = [None, None, None, None, None, 30.00, 13.29, 50.86, 28.87, 44.93]   # stockanalysis 會計年底 P/E（2021–2025）
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2016–2025 共 10 年）"},
    {"label": "stockanalysis（會計年底 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；只有 2021–2025 共 5 年）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝年底 P/E 依循環位置取分位，兩個來源各取分位再平均。兩來源差距大（2021–2025 年 companiesmarketcap 系統性較高；stockanalysis 無 2016–2020），且 2016–2020 年為小型股、獲利低基期的高倍數，2023、2025 年為高成長期落後 P/E；規則值很可能高估成長放緩後的倍數（高盛稱 2018 年以來平均前瞻 P/E 約 29x，In 頁 gs_tp）。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
# 舊法：規則值情境加權（v0.1 稱「0. 目標價與上檔空間」；v0.2 改名保留，只作參考；建置後與其他內容一起下移）
V.cell(4, 1, "舊法：規則值情境加權（參考）").font = BOLD
hdr(V, 5, ["項目", "數值", "", "", "說明"])
V.cell(6, 1, "規則值情境加權價（參考，非目標價）"); V["B6"] = f"=tp_wA*D{VAL['A']}+tp_wB*D{VAL['B']}+tp_wC*D{VAL['C']}"; V["B6"].number_format = "#,##0"
V["E6"] = "權重 × 各情境隱含價（第 3 節規則值）"
V.cell(7, 1, "規則值情境加權價 ÷ 現價 − 1（參考）"); V["B7"] = "=B6/px-1"; V["B7"].number_format = "0.0%"; V["E7"] = "現價＝In 頁 px（2026-10-30 收盤）"
for rr, s, lab in ((8, "A", "情境 A 隱含價（樂觀）"), (9, "B", "情境 B 隱含價（基準）"), (10, "C", "情境 C 隱含價（下檔）")):
    V.cell(rr, 1, lab); V[f"B{rr}"] = f"=D{VAL[s]}"; V[f"B{rr}"].number_format = "#,##0"
    V[f"E{rr}"] = f'="FY2027 EPS × "&TEXT(C{VAL[s]},"0.0")&"x（"&C{28 + SC.index(s)}&"，第 3 節）"'
V.cell(11, 1, "情境權重 A／B／C"); V["B11"] = "=tp_wA"; V["C11"] = "=tp_wB"; V["D11"] = "=tp_wC"
for c in "BCD": V[f"{c}11"].number_format = "0%"
V["E11"] = '="合計 "&TEXT(B11+C11+D11,"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
V.cell(12, 1, "舊法說明")
V["B12"] = ("P/E：各情境 FY2027 EPS × 十年（2016–2025 年底）P/E 依循環位置取分位（本模型營收最高年判定：營收到 FY2029 仍增加者為中段＝中位數，見頂＝25 分位）。"
            "十年倍數含 2016–2020 年小型股低獲利基期與 2023、2025 年高成長期的落後高倍數，偏誤方向為高估；隱含價約為 FY2027 年底的價格、未折現，當作 12 個月目標價時偏高。對照：高盛以 29.3x 前瞻 P/E 給目標價 633 元（In 頁 gs_tp）")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.80、2028 0.65、2029 0.55（主要客戶轉單、1.6T 落後）", dict(rel="RLO")),
         ("份額上升（高）", "份額相對值 2027 1.15、2028–29 1.20（搶中際旭創、Coherent 份額）", dict(rel="RHI")),
         ("美國限制", "份額相對值 2027 0.90、2028 0.55、2029 0.50（Covered List 擴及光模組）", dict(rel="us")),
         ("降價快於模型", "單價係數 2027 0.90、2028 0.80、2029 0.72", dict(pxr="PXLO")),
         ("毛利率 低（價格戰）", "毛利率 2027 43%、2028 40%、2029 36%", dict(gm="GMLO")), ("毛利率 高", "毛利率 2027 49%、2028 48%、2029 47%", dict(gm="GMHI")),
         ("CPO 份額 0", "新易盛不參與 CPO 光學", dict(cpo="CPLO")), ("CPO 份額 高", "2027 20%、2028 25%、2029 30%", dict(cpo="CPHI")),
         ("AI 市場 ×0.8", "FY2027 起可插拔光模組市場 ×0.8", dict(mf="mf_lo")), ("AI 市場 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("人民幣升值（6.9）", "FY2027 起 RMB／USD 6.9", dict(fx="fx_lo")), ("人民幣貶值（7.3）", "FY2027 起 7.3", dict(fx="fx_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi")),
         ("營業費用彈性 1.0", "phi＝1.0（費用與營收同步成長）", dict(phi="phi_hi"))]
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("PXLO", "pxr", "_lo"), ("GMLO", "gm", "_lo"), ("GMHI", "gm", "_hi"), ("CPLO", "cpo", "_lo"), ("CPHI", "cpo", "_hi")):
    for y in (27, 28, 29):
        wb.defined_names[f"{nm}_{y}"] = DefinedName(f"{nm}_{y}", attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
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
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85；CPO 情境 B", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6；CPO 情境 C", PB["C"], "PL!")
assert nn + 3 < 40
SENS_N = len(cases)
SE.cell(nn + 3, 1, "讀法：FY2027 三情境只差在網通 CPO 情境（A 替代快、C 慢）；顆數差異從 FY2028 起。「美國限制」「份額流失」「降價」「毛利率」四案例對應市場隱含衰退的路徑；新易盛境外營收占九成七以上，美國限制與主要客戶轉單是最大的單一風險。")
SENS_READ = nn + 3
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、券商、研調。只作事後對照，不為對上數字而調參數。共識（同花順 19 家，明細只列 10 家）區間很寬（FY2027 EPS 16.29–54.49），均值明顯高於高盛（FY2027 歸母淨利 +21%）：天風、國金兩家特別高"
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
RC.cell(rr_, 1, "A. 共識與券商（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "同花順（19 家；440.5–684.8）", "錨點＝1H26 實績＋2H26 季增假設")
rec("FY2026 歸母淨利 vs 共識", "=cons_ni26", pl("ni", 2026), "同花順（19 家；126.9–272.9）", "")
rec("FY2026 歸母淨利 vs 高盛", "=gs_ni26", pl("ni", 2026), "華爾街見聞轉引高盛（2026-07）", "")
rec("FY2027 營收 vs 共識", "=cons_rev27", pl("rev", 2027), "同花順（570.5–1,566.6）", "")
rec("FY2027 歸母淨利 vs 共識", "=cons_ni27", pl("ni", 2027), "同花順（227.1–541.6）", "")
rec("FY2027 歸母淨利 vs 高盛", "=gs_ni27", pl("ni", 2027), "華爾街見聞轉引高盛（2026-07）", "高盛低於共識均值約兩成")
rec("FY2027 EPS vs 共識", "=cons_eps27", pl("eps", 2027), "同花順（16.29–54.49）", "")
rec("FY2028 營收 vs 共識", "=cons_rev28", pl("rev", 2028), "同花順（711.4–2,371.1）", "")
rec("FY2028 歸母淨利 vs 高盛", "=gs_ni28", pl("ni", 2028), "華爾街見聞轉引高盛（2026-07）", "")
rec("FY2028 EPS vs 共識", "=cons_eps28", pl("eps", 2028), "同花順（21.94–65.32）", "")
rec("3Q26 營收（模型假設）vs 共識", "=q3_cons_rev", lambda s: "=Q!D4", "investing.com", "季報 2026-10-30 公布後取代")
RC.cell(rr_, 1, "B. 份額與市場（機制檢查）").font = BOLD; rr_ += 1
rec("FY2025 AI 營收（美元）vs LightCounting 2025 新易盛光模組營收", "=lc_eo25", lambda s: "=a25_aiu", "LightCounting", "口徑不同：LightCounting 含全部光模組", fmt="#,##0.0")
rec("FY2026 可插拔有效份額＋中際旭創份額（兩家合計）", None, lambda s: f"=PL!D{PB[s]['sh0']}+inl_26", "innolight 成品", "合計應明顯低於 100%（另有 Coherent 等）", fmt="0.0%")
rec("FY2028 兩家合計可插拔份額", None, lambda s: f"=PL!F{PB[s]['sh']}+inl_28", "innolight 成品（情境 A）", "", fmt="0.0%")
rec("FY2026 新易盛 AI 營收 ÷ 中際旭創 AI 營收（美元）", None, lambda s: f"=PL!D{PB[s]['aiu']}/inl_aiu26", "innolight 成品 Units 頁", "西南證券：1H26 營收約為中際旭創的 50.1%", fmt="0.0%")
rec("FY2027 新易盛 AI 營收 ÷ AI IT 資本支出", None, lambda s: f"=PL!E{PB[s]['aiu']}/itc_27", "markets/network", "每 1 美元 AI IT 資本支出分到的光模組營收", fmt="0.00%")
RC.cell(rr_, 1, "C. 校準與價格").font = BOLD; rr_ += 1
rec("1H26 EPS（模型）vs 公司揭露 5.41", "=h1_eps", lambda s: "=Q!C13", "公司", "股數以淨利 ÷ EPS 推導，恆等", fmt="0.00")
rec("FY2026 營業費用率（Derived）", None, lambda s: "=Q!G9", "Q 頁殘差", "", fmt="0.0%")
rec("規則值情境加權價（舊法，參考）vs 高盛目標價", "=gs_tp", lambda s: "=Val!B6", "高盛（2026-07）", "本模型用十年落後 P/E，高盛用前瞻 29.3x", fmt="#,##0")
rec("規則值情境加權價（舊法，參考）vs 同花順 90 天機構目標均價", "=tp_avg90", lambda s: "=Val!B6", "同花順（15 家）", "", fmt="#,##0")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以新易盛 AI 營收 ÷ 部署 GW 計算每 GW 營收"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "3Q26 季報（2026-10-29）：v0.3 已以實績取代 Q 頁 3Q（營收 109.40 億，低於共識 177.6 億）"); RC.cell(rr_, 10, "見 In 頁最後 q3_*"); rr_ += 1
RC_END = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝深交所定期報告與公告，經財經媒體轉載；原文 PDF 待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司（轉載）：FY2025 營收、歸母淨利、EPS、10 轉 4 派 10（東方財富／中證網）", FY25E),
        ("公司（轉載）：FY2025 光互聯毛利率、4Q25、1Q26 實績與匯兌損失（每日經濟新聞）", FY25N),
        ("公司（轉載）：FY2025 存貨、減值、1Q26 季減（虎嗅）", HX),
        ("公司（轉載）：1H26 實績、光互聯毛利率 48.46%、境外收入、研發（C114）", H1C),
        ("公司（轉載）：1H26 基本 EPS 5.41（東方財富）", H1E),
        ("券商：西南證券 9-15（1H26 綜合毛利率 48.44%、境外 97.92%、與中際旭創比較）", XN),
        ("券商：長江證券（2Q26 毛利率 47.96%、淨利率 37.77%）", CJ),
        ("券商：開源證券 8-25（總股本 13.94 億、少數股東損益、費用率）", KY),
        ("券商：天風證券 10-10（2Q26、產能、盈利預測）", TF),
        ("公司（轉載）：2026 年限制性股票激勵（同花順）", INC),
        ("公司說法（轉載）：NPO 2027 批量、1.6T 下半年放量（同花順 9-28）", N928),
        ("共識：同花順 F10 盈利預測（2026-10-23）", CONS),
        ("券商：高盛 2026-07（盈利預測、目標價 633）（華爾街見聞）", GS),
        ("券商：高盛 2025-12（前瞻 P/E 歷史平均約 29x）（老虎證券）", GS25),
        ("研調：LightCounting 2025 廠商排名", LC),
        ("券商：花旗對 FCC 26-50 的解讀（鉅亨）", CITI),
        ("價格：中財網（10-23 收盤）、etnet（同值、52 週區間）、Google 財經（9-30 收盤）", PX1),
        ("財報日：investing.com（3Q26 2026-10-30、3Q 共識）", INV),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/eoptolink-technology/pe-ratio/"),
        ("資料商：stockanalysis 會計年底 P/E", "https://stockanalysis.com/quote/she/300502/financials/ratios/"),
        ("市場層：markets/network、tsmc、innolight 成品（本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 季報（10-30）；全公司 FY2025 毛利率；客戶名稱與集中度；800G／1.6T 份額第一手來源；泰國子公司稅率；52 週高點（資料商不一致）；近一個月跌幅原因；前瞻 P/E 歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("新易盛（300502.SZ）收支與估值模型 v0.2.2（AI 半導體 Project：網通次產業內複製——光模組）　2026-10-26", True),
 ("版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）；v0.2（2026-10-25：三情境延伸到 FY2032、目標價＝DCF 與 Comparable 加權；v0.1 初版）。新易盛在 AI 光模組的份額在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）；中際旭創份額讀 innolight 成品，兩家合計在 Units 頁檢查。", False),
 ("模型命題：新易盛的 AI 營收＝AI 可插拔光模組市場（本專案網通模型）× 新易盛份額 × 單價係數＋CPO 光學 × 新易盛 CPO 份額；其他營收另依成長率推。股價 261 元（10-23）自 9 月底下跌約三成、只約 10 倍模型 FY2027 EPS：市場定價像是 2027 年就是獲利高點，或份額與美國限制風險遠大於模型。", False),
 ("核心問題：以 15 倍看 FY2028，現價隱含的 EPS 約等於模型情境 C、比情境 A 低四成——要成立，份額（主要客戶轉單、1.6T 落後）、美國限制、單價三條路徑各要不利到什麼程度？", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導；FY2030–FY2032 為收斂法。口徑：中國會計準則；淨利＝歸母淨利；股數與 EPS 全部以 10 轉 4 後表示（FY2025 EPS 9.61 ÷ 1.4 重述）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 營收：markets/network 的 AI 可插拔光模組市場 × 情境顆數係數 × 新易盛有效份額（FY2026 以 1H26 實績＋2H26 假設校準）× 份額相對值（Assumed）× 單價係數（基準 1）＋ CPO 光學市場 × CPO 份額（Assumed）；乘匯率換成人民幣。", False),
 ("其他營收：FY2026 營收 × 5%（Assumed）× 成長率。毛利率：FY2027 起為輸入路徑（Assumed）。營業費用、減值與其他：彈性 phi 0.6（Analogy：中際旭創）。稅率 12%（Assumed）、歸母比例 0.998（Derived）。", False),
 ("校準：FY2026 ＝ 1H26 實績（營收 20.91、綜合毛利率 48.44%、歸母淨利 7.53）＋ 2H26（季增 20%、毛利率 48%、營業費用率 5%，Assumed）。公司不給指引；不用共識（只在 Recon 對帳）。3Q26 季報預計 10-30 公布，本版未納入。", False),
 ("估值：目標價（Val 第 0 節）＝0.5 × DCF（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 0.3／0.5／0.2）＋0.5 × Comparable（A 股光通訊同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）。舊法規則值情境加權保留參考；現價 ÷ FY2027、FY2028 EPS；市場隱含 FY2028 EPS（Val 第 4 節）；十年 P/E 規則值（2016–2025，兩個資料商）× FY2027 EPS；循環位置以本模型營收最高年判定。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文經財經媒體轉載；客戶名稱、客戶集中度與 800G／1.6T 份額未找到第一手來源。", False),
 ("", False),
 ("已知限制：(1) 份額相對值、CPO 份額、毛利率路徑皆為 Assumed，是 FY2028 EPS 最敏感的輸入（Sens）；(2) FY2026 錨點的 2H26 為假設，3Q26 季報 10-30 公布；(3) 網通模型 2027 年可插拔光模組 +84%，未設雷射、DSP 供給上限（Sens 用市場 ×0.8）；(4) 十年 P/E 兩個來源差距大，且含低基期與高成長期溢價，規則值偏高；(5) 稅率與泰國子公司獲利占比未找到；(6) 主要客戶與美國限制風險只有次級來源；(7) R4 產業部署路徑未建（Recon 預留列）。", False)]
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
TPT = _tp_section(V, lambda s_, y: f"PL!{FCX[y]}{SUMR[(s_, 'eps')]}", [(2027, "2/12"), (2028, "10/12")],
    "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）",
    [("中際旭創（Innolight）", "300308.SZ", 14.11, "2026-10-12", "stockanalysis forward PE（S&P Global 共識；價格日期如左）"),
     ("天孚通信（TFC）", "300394.SZ", 46.53, "2026-09-30", "同上（頁面價格日期較舊）"),
     ("光迅科技（Accelink）", "002281.SZ", 77.09, "2026-08-18", "同上（頁面價格日期較舊）"),
     ("華工科技（HGTECH）", "000988.SZ", 45.08, "2026-08-07", "同上（頁面價格日期較舊）"),
     ("劍橋科技（CIG）", "603083.SS", 31.74, "2026-10-23", "同上")],
    "A 股光通訊同業（新易盛是 A 股公司，同業只用 A 股；工作單 Andy 11:53）",
    "同花順 90 天機構目標均價 460.89（15 家；In 頁 tp_avg90）；多數訂於 9 月前股價較高時",
    ("DCF（v0.4 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（EBIT＝稅前利益，中國會計準則，股份支付已列費用），以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 人民幣 10 年期國債)、RONIC＝WACC；ke＝rf＋β ×（ERP 5%＋中國 CRP）；EV 加淨現金減少數股權 ÷ 股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁、tools/fcff_dcf.py）。"
     "Comparable：A 股光通訊同業（中際旭創、天孚通信、光迅科技、華工科技、劍橋科技）前瞻 P/E 中位數；同業規模較小、部分價格日期為 8–9 月，"
     "而中際旭創只有約 14 倍，中位數偏高估。情境對應：顆數情境 A＝順利、B＝平平、C＝不如預期；FY2027 三情境 EPS 只差在網通 CPO 情境（A＝CPO 替代快），顆數差異從 CY2028 起。"),
    _rule_pe_range(V, 2), "#,##0.00",
    "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝In 頁 cons_feps（同花順 FY2027、FY2028 共識，與本模型同期間加權）")
# ---------------- v0.4：標準版 DCF（FCFF＋WACC）——DCF 明細頁＋Val 0.1 節摘要（取代 v0.2 股東現金流法；tools/fcff_dcf.py） ----------------
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = "DCF — 標準版（v0.4）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py"
DCFW["A2"] = ("EBIT：FY2027–FY2029 取 PL 推導區塊「稅前利益（＝毛利 − 營業費用、減值與其他）」（含少量利息與投資收益，口徑同 PL）；FY2030–FY2032（收斂法只有歸母淨利）由淨利還原：稅前利益＝歸母淨利 ÷（（1 − tx）× attr）。"
              "稅率＝tx（15%）。少數股權在評價時以帳面扣除。折舊、資本支出、營運資金：近三年（FY2023–FY2025）占營收平均；資本支出 FY2028 起收斂到最近一年（In 頁 M2 節）。資產負債為 2026-09-30（第三季報），淨現金以 stub 期（第一年 FCFF × stub_frac − 已宣告未付股利）滾到評價日。")
FY_X = list(range(2027, 2033))
def _rev(s_, y):
    return f"PL!{FCX[y]}{PB[s_]['rev']}"
def _ebit(s_, y):
    return f"PL!{FCX[y]}{PB[s_]['pre']}" if y <= 2029 else f"=PL!{FCX[y]}{PB[s_]['ni']}/((1-tx)*attr)"
DPATHS = [dict(name=s, label=f"情境 {s}（{ {'A': '順利', 'B': '平平', 'C': '不如預期'}[s] }）", rev={y: _rev(s, y) for y in FY_X}, rev0=_rev(s, 2026),
              ebit={y: _ebit(s, y) for y in FY_X}, tax={y: "tx" for y in FY_X}, ebit_note="稅前利益；FY2030 起＝歸母淨利 ÷（（1 − tx）× attr）") for s in SC]
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 1.0, title="三情境 FCFF 明細（RMB B；FY(t)＝CY(t)；評價日 2026-12-31，FY2027 全年在評價日後）", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="6.576626+0.272152+0.013766+0.028499-0-(0.008081+0.030741)-0.225555",
                             nc_note="2026-09-30 資產負債表（RMB B；第三季報）：貨幣資金 6.577＋交易性金融資產 0.272＋長期股權投資 0.014＋其他非流動金融資產 0.028 − 租賃（0.008＋0.031）− 少數股東權益 0.226") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
_r1, _r2 = 5, 5 + FD.summary_rows(stub=True) - 1   # 舊 0.1 節（第 5–22 列）＋原空白第 23–24 列整段改寫（v0.4.2 stub 多兩列）；0.2、0.3 節位置不動
assert all(V.cell(_rr, _cc).value is None for _rr in range(TPT["dcftp"] + 1, _r2 + 1) for _cc in range(1, 8))
assert V.cell(_r1, 1).value.startswith("0.1 DCF"), V.cell(_r1, 1).value
for _rr in range(_r1, _r2 + 1):
    for _cc in range(1, 8):
        V.cell(_rr, _cc).value = None; V.cell(_rr, _cc).fill = PatternFill(); V.cell(_rr, _cc).font = Font()
assert _r2 - _r1 + 1 >= FD.summary_rows(stub=True)
SDCF = FD.write_summary(wb, V, _r1, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算 CN 光通訊產業 β 依 D/E 重新加槓桿")   # v0.4.2：模組 v2 beta_note
for _k in [k for k in TPT if k.startswith(("eps", "cf")) or k in ("reinv", "dcfv")]:
    del TPT[_k]
TPT.update(dcftp=SDCF["dcf"], dcf12=SDCF["v12"])
V[f"B{TPT['dcftp']}"].fill = KEY
# Comparable 前瞻 EPS 與歷史分位檢查原本讀 0.1 節的 EPS 列（第 8–13 列，已改寫）；改為直接讀 PL 頁（同一儲存格來源，數值不變）
_epl = lambda s_, y: f"PL!{FCX[y]}{SUMR[(s_, 'eps')]}"
_fw = [(2027, "2/12"), (2028, "10/12")]
V.cell(TPT["feps"], 2).value = "=" + "+".join(f"tp_w{s_}*(" + "+".join(f"{w}*{_epl(s_, y)}" for y, w in _fw) + ")" for s_ in SC)
_pe_rng = _rule_pe_range(V, 2)
for _c, s_ in zip("BCD", SC):
    V[f"{_c}{TPT['pr']}"] = f"=IF(px/{_epl(s_, 2028)}>MAX({_pe_rng}),1,IF(px/{_epl(s_, 2028)}<MIN({_pe_rng}),0,PERCENTRANK({_pe_rng},px/{_epl(s_, 2028)})))"
V.cell(TPT["tpw"], 2).value = f"=tp_w_dcf*B{SDCF['dcf']}+(1-tp_w_dcf)*B{TPT['comptp']}"
V.cell(TPT["band"], 2).value = f'=TEXT(D{SDCF["v12"]},"#,##0.00")&"～"&TEXT(B{SDCF["v12"]},"#,##0.00")'
_last = max(r for r in range(1, 40) if any(SE.cell(r, j).value is not None for j in range(1, 13)))
TPS0 = _last + 2 if _last + 2 + 12 < 40 else _last + 1
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${TPT['comptp']}", base_ref=f"Val!$B${TPT['tpw']}", beta_du="beta_du", usd=True,
                     title="目標價參數敏感度（標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）")
assert _tvs[2] - 1 < 40, _tvs   # 檢查列（_tvs[2] − 1）須在第 40 列（案例推導區塊起點）之前
TPS_LAST = _tvs[1]
FD.yellow(IN, FD.JUDG + ("beta_u_med", "beta_du"))

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 14:00、14:17；既有數值與儲存格位置不變） ----------------
_UPW = [r_ for r_ in range(1, 60) if V.cell(r_, 1).value == "上檔空間：加權目標價"]
assert _UPW == [42], _UPW
GR_ROW, RC_ROW = 49, 50   # 0.3 結論後的空白預留列（不移動其他列）
for _r in (GR_ROW, RC_ROW):
    assert all(V.cell(_r, _c).value is None for _c in range(1, V.max_column + 1)), _r
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B42>0.5,"大於 50%",IF(B42>0.25,"25%–50%",IF(B42>=0,"0%–25%",IF(B42>=-0.25,"0%～−25%",IF(B42>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B42>0.5,"強力買進",IF(B42>0.25,"買進",IF(B42>=-0.25,"中立",IF(B42>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = ["目標價 ¥586（DCF 427、Comparable 相對倍數調整後 745 各半），比現價 196.00（2026-10-30）高 199.0%，級距 大於 50%，投資建議維持強力買進。v0.3 滾入 3Q26：營收 109.40 億、季減 13.0%（原假設季增 20%，共識 177.6 億），歸母淨利 38.43 億；加權目標價 706 → 586（−17.0%）。上檔空間仍大，主因是股價：10-23 的 261.00 跌到 10-30 的 196.00（−24.9%），現價只約 9.3 倍模型 FY2027 EPS（A）；DCF 單獨也有 +118%。", "3Q 營收不及預期使 FY2026 校準份額下移：FY2026 EPS 14.40 → 11.60（−19.5%），FY2027 EPS（A）25.58 → 21.10（−17.5%）。FY2026 營收 449.8 億比共識低 18.5%，也低於股權激勵第一期考核 500 億（−10.0%）。情境 C FY2028 EPS 13.89，比情境 A 的 23.29 低 40.4%，DCF 價格帶 253～537。", "關鍵是 3Q 的落差是時點還是份額：國盛證券（10-30，Interested-party）稱部分 3Q 推遲的發貨將在 4Q 確認；本版 4Q 季增 20%、毛利率 46.5%（3Q 推得約 46.3%）。若 4Q 營收回到 2Q 的 125.7 億以上，表示是時點，FY2027 份額應上修；若 4Q 仍在 110–130 億，代表份額被中際旭創或其他廠拿走（同季中際旭創季增 24.6%）。下一個檢驗點是 FY2026 年報（2027 年 3–4 月）。"]   # 重點觀察（與 README.md 相同文字）
_n0 = RD.max_row
RD.cell(_n0 + 1, 1, "v0.2.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（Val 頁 0.3 結論加「上檔空間級距」「投資建議」兩列、README 頁加重點觀察三點；數值不變，既有儲存格位置不變）。")
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
# ---------------- v0.2.3（2026-10-28）：分部原則重建（Q 頁第 30 列起追加；不位移既有列，數值不變） ----------------
Q.cell(30, 1, "官方報告產品別（FY2025 年報；分部原則第 3 條）").font = BOLD
for _r, _lab, _v, _tag, _note in (
        (31, "光互聯產品營收（含 AI 與國內、電信）", "=B4*0.997", "Verified", "每日經濟新聞轉載公司 FY2025 年報：光互聯產品占營收 99.7%（金額為推得）"),
        (32, "其他產品營收（差額）", "=B4-B31", "Derived", "合併營收 − 光互聯產品"),
        (33, "檢查：分部合計 − 合併營收", "=ROUND(B31+B32-B4,6)", "Derived", "應為 0；PL 的 AI 營收屬光互聯產品，推估年以 AI／其他拆分")):
    _v2 = float(_v) if not _v.startswith("=") else _v
    Q.cell(_r, 1, _lab); Q[f"B{_r}"] = _v2; Q[f"B{_r}"].number_format = "#,##0.000"
    if not isinstance(_v2, str): Q[f"B{_r}"].font = BLUE
    Q.cell(_r, 8, _tag); Q.cell(_r, 9, _note)
RD.cell(RD.max_row + 1, 1, "v0.2.3（2026-10-28）：分部原則重建（數值不變）。Q 頁第 30 列起追加 FY2025 年報產品別：光互聯產品（約 99.7%）、其他（差額）與核對列「分部合計 − 合併營收」；AI 營收屬光互聯產品，推估年仍以 AI／其他拆分。")
RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.2.2", "v0.2.3")
wb._sheets = [wb[n] for n in order]
# ---------------- v0.2.2（2026-10-26）：網通 CPO 三情境（同字母）；In 頁在現有最後一列之後新增，不移動既有列 ----------------
_ri = IN.max_row + 2
IN.cell(_ri, 1, "網通 CPO 三情境（v0.2.2；markets/network v0.2.3：A＝替代快、B＝平平、C＝慢；情境 A 為上方 A 節，B、C 為下列，各再乘同字母情境顆數係數）").font = BOLD; _ri += 1
for _kb, _lab, _kind in [('mpl', 'AI 可插拔光模組市場', 'mpl'), ('mcp', 'AI CPO 光學市場', 'mcp')]:
    for _s in ("B", "C"):
        for _j, _y in enumerate((26, 27, 28, 29)):
            _k = f"{_kb}{_s}_{_y}"
            IN.cell(_ri, 1, _k); IN.cell(_ri, 2, f"情境 {_s}：{_lab} CY20{_y}"); IN.cell(_ri, 3, "US$B")
            IN.cell(_ri, 5, NETS[(_kind, _s)][0][_j]).font = BLUE
            IN.cell(_ri, 7, "Derived"); IN.cell(_ri, 8, NETS[(_kind, _s)][1])
            IN.cell(_ri, 9, f"網通模型 CPO 情境 {_s}（技術替代屬行業情境，公司讀同字母）" if _y == 26 else "同上")
            wb.defined_names[_k] = DefinedName(_k, attr_text=f"In!$E${_ri}"); _ri += 1
_n1 = RD.max_row
RD.cell(_n1 + 1, 1, 'v0.2.2（2026-10-26）：網通模型 v0.2.3 的 CPO 滲透改三情境（A＝替代快、B＝平平、C＝慢），本模型 A／B／C 改讀同字母的可插拔與 CPO 市場（In 頁末新增 B、C 兩組；情境 A 沿用 A 節 mpl_／mcp_，數值改為網通情境 A）；FY2027 起三情境不再相同。其他輸入不變。')
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "eoptolink")
Q3ROWS = [["q3_rev", "3Q26 營收", 10.94, "Verified", "公司 3Q26 報告（四川經濟網、IT之家 2026-10-29 轉載）：109.40 億元（+80.30%；季減 13.0%）", "前三季 318.50 億元"], ["q3_gp", "3Q26 毛利（前三季毛利 − 1H26 毛利）", "=15.191-h1_rev*h1_gm", "Derived", "IT之家轉載：前三季毛利 151.91 億元、毛利率 47.69%", "1H26 綜合毛利率 48.44% → 3Q 推得約 46.3%"], ["q3_ni", "3Q26 歸母淨利", 3.843, "Verified", "公司 3Q26 報告：38.43 億元（+61.15%；季減 19.1%）", "國盛證券 10-30：部分 3Q 推遲的發貨將在 4Q 確認（Interested-party）"], ["q3_eps", "3Q26 基本 EPS（前三季 8.16 − 1H26 5.41）", 2.75, "Derived", "公司 3Q26 報告：前三季基本 EPS 8.16 元", "只作核對"]]
V03_NOTE = "v0.3（2026-11-02）：滾入 3Q26 實績（營收 109.40 億、季減 13%，原假設季增 20%；歸母淨利 38.43 億；毛利率約 46.3%），4Q26 營收＝3Q ×（1＋季增 20%）、毛利率 46.5%；股價改 2026-10-30 收盤 196.00（Andy 11-01：上檔空間用最新收盤）。"
# ---------------- v0.3（2026-11-02）：3Q26 實績（In 頁最後追加；Q 頁 D 欄改讀） ----------------
_ri = IN.max_row + 2
IN.cell(_ri, 1, "v0.3：3Q26 季報實績（2026-10-29 公布；Q 頁 D 欄）").font = BOLD; _ri += 1
for _k, _lab, _v, _tag, _src, _src2 in Q3ROWS:
    IN.cell(_ri, 1, _k); IN.cell(_ri, 2, _lab); IN.cell(_ri, 3, "RMB" if _k == "q3_eps" else "RMB B")
    IN.cell(_ri, 5, _v).font = BLUE; IN.cell(_ri, 7, _tag); IN.cell(_ri, 8, _src); IN.cell(_ri, 10, _src2)
    wb.defined_names[_k] = DefinedName(_k, attr_text=f"In!$E${_ri}"); _ri += 1
RD.cell(RD.max_row + 1, 1, V03_NOTE)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.3", "2026-11-02")  # v0.2.4：In 頁版面（個股模型原則第 8 節；數值不變）
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
WTA_SENS = _wta_sens({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}, '情境 B；可插拔與 CPO 光學市場（加速器顆數 × 每顆光學內容）CY2027 起 × wta_27／28／29', '做法：引擎的 AI 市場係數 mf 改為年別 dict（wta_27／28／29），同時乘在可插拔與 CPO 兩個市場列；FY2026 不乘。FY2030 以後無延伸表（Sens 引擎只到 FY2029）。目標價：Sens 目標價表只有 DCF 參數案例，本案例不重算目標價。')
# ---------------- v0.4（2026-11-05）：展示層——README 頁「展示層」「命題回覆」「本版要點」三段（個股 HTML 版面規格 v1.0；公式帶數字，原列保留） ----------------
RD["A1"] = "新易盛（300502.SZ）收支與估值模型 v0.4.2（AI 半導體 Project：網通次產業內複製——光模組）　2026-11-05；v0.4.2：換用共用模組 tools/fcff_dcf.py v2（資產負債表更新到 2026-09-30 並以 stub 期滾到評價日、β 敏感度 β_u ± 0.3、新增 Sens「美元口徑折現」）；v0.4.1：β 改用自算產業 β（CN 光通訊；Andy 2026-11-05 22:55），其他不變；v0.4：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），HTML 改用個股 HTML 版面規格 v1.0（README 新增展示層、命題回覆）；Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
def _lr(ws, prefix, col=1):
    """A 欄（或指定欄）以 prefix 開頭的第一列。"""
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))
_ik = lambda k: next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == k)
_V = {k: _lr(V, p) for k, p in (("tp", "目標價（DCF 與 Comparable 加權）"), ("dcf", "DCF 目標價（情境加權）"), ("comp", "Comparable 目標價"),
                                ("compadj", "Comparable 目標價（相對倍數調整後）"), ("up", "上檔空間：加權目標價"), ("upc", "上檔空間：Comparable"),
                                ("rec", "投資建議"), ("px", "現價（"), ("med", "同業前瞻 P/E 中位數"), ("fwd", "12 個月後的前瞻 EPS（情境加權）"),
                                ("epsgap", "與市場共識目標價的差距：EPS 因素"), ("coef", "相對倍數調整係數"), ("curpe", "現價 ÷ 共識未來 12 月 EPS"),
                                ("hi", "現價 ÷ 52 週高點"), ("q4", "4. 核心問題"), ("pe28", "現價 ÷ FY2028 EPS（模型）"),
                                ("tvsh", "  其中：終值現值占 EV"), ("mult", "  隱含終值 EV／EBITDA"))}
assert V.cell(_V["comp"], 1).value == "Comparable 目標價"
_v = lambda k, c="B": f"Val!{c}{_V[k]}"
_q = lambda i, c: f"Val!{c}{_V['q4'] + 1 + i}"   # 第 4 節：i＝1 起為假設倍數各列
_S = {k: _lr(SE, p) for k, p in (("sh", "份額流失（低）"), ("shh", "份額上升（高）"), ("us", "美國限制"), ("px", "降價快於模型"), ("gm", "毛利率 低"),
                                ("mkt", "AI 市場 ×0.8"), ("scB", "顆數情境 B"), ("scC", "顆數情境 C"), ("wdcf", "DCF 權重 0.3／0.7"), ("swacc", "WACC ±1pt"), ("sbeta", "β 低／高"), ("usd", "美元口徑折現"))}
_s = lambda k, c: f"Sens!{c}{_S[k]}"
_pl = lambda s_, k, y: f"PL!{FCX[y]}{SUMR[(s_, k)]}"
_u_sh = _lr(UN, "新易盛可插拔有效份額")
_P = lambda x: f'TEXT({x},"+0.0%;-0.0%")'
_N = lambda x, f="#,##0.00": f'TEXT({x},"{f}")'
_in = lambda c, k, f="0.00": _N(f"In!{c}{_ik(k)}", f)
_upa = f"({_v('tp')}-px)"
_sh = lambda num: _N(f"({num})/{_upa}", "0%")
_r = RD.max_row + 2
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.4 新增，A3、A4、重點觀察原列保留）").font = BOLD; _r += 1
RD.cell(_r, 1, "命題：新易盛的 AI 營收＝AI 可插拔光模組市場 × 新易盛份額 × 單價，再加 CPO 光學；境外營收占九成七以上、客戶集中，股價自高點大跌、約十倍的共識未來 12 月 EPS，市場定價像是 2027 年就是獲利高點。三個問題如下。"); _r += 1
RD.cell(_r, 1, "命題 1｜份額撐不撐得住：中際旭創、Coherent 搶 1.6T 份額、主要客戶轉單，或美國限制擴及光模組時，FY2028 EPS 會少多少。"); _r += 1
RD.cell(_r, 1, "命題 2｜單價與毛利率：光模組降價快於模型、價格戰壓低毛利率時，FY2028 EPS 會少多少。"); _r += 1
RD.cell(_r, 1, f'="命題 3｜現價隱含什麼：股價比 52 週高點低 "&{_N(f"-{_v('hi')}", "0%")}&"、約 "&{_N(_v("curpe"), "0.0")}&" 倍共識未來 12 月 EPS；市場隱含的 FY2028 EPS 要多低才說得通、每股值多少。"'); _r += 1
RD.cell(_r, 1, f'="重點 1｜合理價值"&IF({_v("up")}<0,"低於","高於")&"現價，投資建議"&{_v("rec")}&"：加權目標價每股約 "&{_N(_v("tp"))}&" 元（標準版 DCF "&{_N(_v("dcf"))}&"、Comparable 經相對倍數調整後 "&{_N(_v("compadj"))}&"，各占一半），相對現價 "&{_N(_v("px"))}&" 為 "&{_P(_v("up"))}&"。只看同業倍數、不做相對倍數調整的 Comparable 目標價是 "&{_N(_v("comp"))}&"（"&{_P(_v("upc"))}&"），這個數不進加權目標價。"'); _r += 1
RD.cell(_r, 1, f'="重點 2｜上檔主要來自 "&IF((1-tp_w_dcf)*({_v('compadj')}-px)>tp_w_dcf*({_v('dcf')}-px),"Comparable","標準版 DCF")&"：加權目標價比現價"&IF({_upa}<0,"低 ","高 ")&{_N(f"ABS({_upa})")}&" 元，其中 Comparable 貢獻 "&{_sh(f"(1-tp_w_dcf)*({_v('compadj')}-px)")}&"、標準版 DCF 貢獻 "&{_sh(f"tp_w_dcf*({_v('dcf')}-px)")}&"（DCF 每股 "&{_N(_v("dcf"))}&" 元，相對現價 "&{_P(_v("dcf") + "/px-1")}&"；WACC "&{_N("wacc", "0.0%")}&"＝rf "&{_N("rf", "0.00%")}&" ＋ β "&{_N("beta", "0.00")}&" ×（ERP "&{_N("erp", "0%")}&" ＋ 中國 CRP "&{_N("crp", "0.00%")}&"），β 取自算 CN 光通訊產業 β（同業中位數 Blume 調整後依 D/E 重新加槓桿），ERP 與 RONIC＝WACC 為假設；永續成長率受人民幣 rf 限制為 "&{_N("g_tv", "0.00%")}&"，終值現值占情境 B 企業價值 "&{_N(_v("tvsh", "C"), "0%")}&"）。Comparable 內部用同一時間窗拆：倍數差——A 股光通訊同業前瞻 P/E 中位數 "&{_N(_v("med"), "0.0")}&" 倍，是新易盛目前前瞻 P/E "&{_N(_v("curpe"), "0.0")}&" 倍（現價 ÷ 共識未來 12 月 EPS，同窗）的 "&{_N(_v("med") + "/" + _v("curpe"), "0.0")}&" 倍；盈餘差——本模型 12 個月後的前瞻 EPS "&{_N(_v("fwd"))}&" 比共識同期"&IF({_v('epsgap')}<0,"低 ","高 ")&{_N(f"ABS({_v('epsgap')})", "0%")}&"，相對倍數調整係數 "&{_N(_v("coef"))}&IF({_v("coef")}<1,"（小於 1，向下）","（不小於 1）")&"。"&IF(MAX({_v("tvsh")},{_v("tvsh", "C")},{_v("tvsh", "D")})>0.8,"注意：終值現值占企業價值超過 80% 的情境——A "&{_N(_v("tvsh"), "0%")}&"、B "&{_N(_v("tvsh", "C"), "0%")}&"、C "&{_N(_v("tvsh", "D"), "0%")}&"。","三情境終值現值占企業價值 "&{_N(_v("tvsh"), "0%")}&"／"&{_N(_v("tvsh", "C"), "0%")}&"／"&{_N(_v("tvsh", "D"), "0%")}&"，皆未超過 80%。")'); _r += 1
RD.cell(_r, 1, f'="重點 3｜EPS 最敏感的是份額與加速器顆數：以情境 A 為基準一次改一個輸入，FY2028 EPS 在美國限制案例變動 "&{_P(_s("us", "L"))}&"、份額流失 "&{_P(_s("sh", "L"))}&"、降價快於模型 "&{_P(_s("px", "L"))}&"、AI 市場 ×0.8 "&{_P(_s("mkt", "L"))}&"；顆數情境 C 為 "&{_P(_s("scC", "L"))}&"。目標價參數（以情境加權目標價為基準，一次動一個）：DCF 權重 0.3／0.7 時加權目標價 "&{_N(_s("wdcf", "C"))}&"／"&{_N(_s("wdcf", "E"))}&"；WACC ±1pt "&{_P(_s("swacc", "D"))}&"／"&{_P(_s("swacc", "F"))}&"、β 低／高 "&{_P(_s("sbeta", "D"))}&"／"&{_P(_s("sbeta", "F"))}&"。美元口徑折現（rf 改美國 10 年期 "&{_N("rf_us", "0.00%")}&"＋β ×（ERP＋CRP）、g＝MIN("&{_N("g_cap", "0%")}&", "&{_N("rf_us", "0.00%")}&")＝"&{_N("MIN(g_cap,rf_us)", "0.00%")}&"，其他不變；基準為人民幣利率 "&{_N("rf", "0.00%")}&"＋固定匯率）時加權目標價 "&{_N(_s("usd", "C"))}&"（"&{_P(_s("usd", "D"))}&"）。"'); _r += 1
_r += 1
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算回答開頭三個命題）").font = BOLD; _r += 1
RD.cell(_r, 1, f'="份額撐不撐得住？答：影響大；基準份額路徑本身是假設。基準份額相對值（相對 CY2026）為 "&{_in("E", "rel_27")}&"／"&{_in("E", "rel_28")}&"／"&{_in("E", "rel_29")}&"（CY2027／28／29，Assumed），情境 A 的可插拔有效份額由 CY2026 的 "&{_N(f"Units!C{_u_sh}", "0.0%")}&IF(Units!E{_u_sh}<Units!C{_u_sh}-0.0005," 降到"," 大致維持在")&" CY2028 的 "&{_N(f"Units!E{_u_sh}", "0.0%")}&"。以情境 A 為基準，份額相對值降為 "&{_in("D", "rel_27")}&"／"&{_in("D", "rel_28")}&"／"&{_in("D", "rel_29")}&" 的測試案例 FY2028 EPS 變動 "&{_P(_s("sh", "L"))}&"；美國限制案例（"&{_in("E", "us_27")}&"／"&{_in("E", "us_28")}&"／"&{_in("E", "us_29")}&"，Assumed）變動 "&{_P(_s("us", "L"))}&"；份額上升（"&{_in("F", "rel_27")}&"）時 "&{_P(_s("shh", "L"))}&"。這些是測試案例，不是門檻。"'); _r += 1
RD.cell(_r, 1, f'="單價與毛利率的影響多大？答：兩者相近。單價係數降為 "&{_in("D", "pxr_27")}&"／"&{_in("D", "pxr_28")}&"／"&{_in("D", "pxr_29")}&"（基準 1，Assumed）時 FY2028 EPS 變動 "&{_P(_s("px", "L"))}&"；毛利率由基準 "&{_in("E", "gm_27", "0%")}&"／"&{_in("E", "gm_28", "0%")}&"／"&{_in("E", "gm_29", "0%")}&"（Assumed）降到 "&{_in("D", "gm_27", "0%")}&"／"&{_in("D", "gm_28", "0%")}&"／"&{_in("D", "gm_29", "0%")}&" 時變動 "&{_P(_s("gm", "L"))}&"（皆以情境 A 為基準）。"'); _r += 1
RD.cell(_r, 1, f'="現價隱含什麼、每股值多少？答：現價是本模型 FY2028 EPS 的 "&{_N(_v("pe28"), "0.0")}&"／"&{_N(_v("pe28", "C"), "0.0")}&"／"&{_N(_v("pe28", "D"), "0.0")}&" 倍（情境 A／B／C）。若市場給 "&{_N(_q(2, "A"), "0")}&" 倍，隱含 FY2028 EPS "&{_N(_q(2, "B"))}&"，比情境 A 的 "&{_N(_q(2, "C"))}&" 低 "&{_N(f"-{_q(2, 'D')}", "0%")}&"、比情境 C 的 "&{_N(_q(2, "E"))}&" 高 "&{_N(_q(2, "F"), "0%")}&"；用十年 P/E 25 分位（見頂倍數 "&{_N(_q(4, "A"), "0.0")}&" 倍）評價，現價只需要 FY2028 EPS "&{_N(_q(4, "B"))}&"。依本模型每股約 "&{_N(_v("tp"))}&" 元（加權目標價；標準版 DCF "&{_N(_v("dcf"))}&"、Comparable 經相對倍數調整後 "&{_N(_v("compadj"))}&" 各半），相對現價 "&{_P(_v("up"))}&"，投資建議"&{_v("rec")}&"；只看 DCF 則相對現價 "&{_P(_v("dcf") + "/px-1")}&"。"'); _r += 1
_r += 1
RD.cell(_r, 1, "本版要點與版本沿革（HTML 附錄用）").font = BOLD; _r += 1
RD.cell(_r, 1, f'="v0.4.2（2026-11-05）：換用共用模組 tools/fcff_dcf.py v2（工作單第二輪修正，Andy 2026-11-05 23:16 定案）。(1) 資產負債表更新到 2026-09-30（第三季報）：淨現金與投資 "&TEXT(dcf_nc,"0.00")&"（RMB B；2026-06-30 為 7.09），以 stub 期滾到評價日 2026-12-31＝＋stub FCFF（各情境 FY2027 FCFF × stub_frac "&TEXT(stub_frac,"0.000")&" 年；情境 B "&TEXT(Val!C{SDCF["stub"]},"0.00")&"）（無已宣告股利）＝"&TEXT(Val!C{SDCF["ncr"]},"0.00")&"。(2) β 敏感度改 β_u ± 0.3 各自重新加槓桿。(3) 新增 Sens「美元口徑折現」：加權目標價 "&TEXT(Sens!C{_S["usd"]},"#,##0.00")&"（"&TEXT(Sens!D{_S["usd"]},"+0.0%;-0.0%")&"）。DCF 目標價 200.96 → "&TEXT(Val!B{SDCF["dcf"]},"#,##0.00")&"、加權目標價 472.78 → "&TEXT(Val!B{TPT["tpw"]},"#,##0.00")&"（上檔 "&TEXT({_v("up")},"+0.0%")&"，"&{_v("rec")}&"）。PL、Comparable、情境權重不變；本模型沒有逐年資本支出列，資本支出、折舊維持占營收比例。"'); _r += 1
RD.cell(_r, 1, '="v0.4.1（2026-11-05）：β 改用自算產業 β（Andy 2026-11-05 22:55）：CN 光通訊同業 6 家（300308、300502、300394、002281、000988、603083）無槓桿 β 中位數 "&TEXT(beta_u_med,"0.00")&"（Yahoo Finance 月報酬 2021-11～2026-10 對上證綜指迴歸，Hamada 去槓桿、現金修正），Blume 調整後 "&TEXT(beta_u,"0.000")&"，依市值 D/E 重新加槓桿 β＝"&TEXT(beta,"0.000")&"，WACC "&TEXT(wacc,"0.00%")&"；取代 Damodaran Global Telecom. Equipment 1.30。β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日維持 2026-12-31。其他輸入不變。"'); _r += 1
RD.cell(_r, 1, "v0.4（2026-11-05）：(1) DCF 改為標準版（Andy 2026-11-05 20:42、20:45；docs/plan/個股DCF標準規格.md v1.0）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（EBIT＝稅前利益），以 WACC 折現（ke＝rf＋β ×（ERP＋CRP），rf＝人民幣 10 年期國債 1.61%、β＝Damodaran Global Telecom. Equipment 現金修正無槓桿 β 1.30 依市值 D/E 重新加槓桿、ERP 5%、中國 CRP 0.91%），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)＝1.61%、RONIC＝WACC；EV 加淨現金減少數股權 ÷ 股數，以 ke 推到 12 個月後並扣除期間股利。新增 DCF 頁（共用模組 tools/fcff_dcf.py），Val 0.1 節改為 WACC 與 DCF 彙總，Sens 目標價參數改為十個標準案例；刪除 tp_r、tp_roe。(2) HTML 改用個股 HTML 版面規格 v1.0（README 頁新增展示層、命題回覆與本段；html.json 改由 gen_html_json.py 產生）。PL、Comparable、情境權重不變。"); _r += 1
RD.cell(_r, 1, f"=A{_lr(RD, 'v0.3（2026-11-02）')}"); _r += 1
SO.cell(SO.max_row + 1, 1, "[Analogy] v0.4 標準版 DCF（v0.4.1 起 β 不再用 Damodaran 表，僅 CRP 沿用）：Damodaran Betas by Sector（Global、US，2026 年 1 月）Telecom. Equipment 現金修正無槓桿 β 1.30（Global）、0.89（US）；Country Risk Premium（2026-01）China 0.91%"); SO.cell(SO.max_row, 2, "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html")
SO.cell(SO.max_row + 1, 1, "[Interested-party] v0.4 標準版 DCF：資料商整理的合併財報（2026-06-30 資產負債、FY2023–FY2025 折舊、資本支出、營運資金科目、每股股利）；人民幣 10 年期國債 1.61%（Trading Economics）"); SO.cell(SO.max_row, 2, "https://stockanalysis.com/quote/she/300502/financials/；https://tradingeconomics.com/china/government-bond-yield")
from openpyxl.styles import Alignment as _AL
SO.cell(SO.max_row + 1, 1, "[Derived] 自算產業 β（Andy 2026-11-05 22:55）：CN 光通訊同業 6 家（300308、300502、300394、002281、000988、603083）Yahoo Finance 含股利調整月收盤價 2021-11～2026-10，對 上證綜指 000001.SS 迴歸；Hamada 去槓桿、現金修正，中位數 2.69；Blume 調整 0.67β＋0.33＝2.132"); SO.cell(SO.max_row, 2, "https://finance.yahoo.com（chart API 月資料；2026-11-05 抓取）")
SO.cell(SO.max_row + 1, 1, "[Verified] v0.4.2 資產負債表（2026-09-30）：新易盛 2026 年第三季度報告 貨幣資金、交易性金融資產、租賃、長期股權投資、其他非流動金融資產、少數股東權益、股本"); SO.cell(SO.max_row, 2, "https://static.cninfo.com.cn/finalpage/2026-10-30/1225642947.PDF")
SO.cell(SO.max_row + 1, 1, "[Verified] v0.4.2 美國 10 年期公債殖利率 5.00%（2026-10-23；只用於 Sens「美元口徑折現」）"); SO.cell(SO.max_row, 2, "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；https://fred.stlouisfed.org/series/DGS10")
wb.save(sys.argv[1])
if os.environ.get("EOPTOLINK_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "UN_NOTE": UN_NOTE, "CHK": CHK_ROW, "SENSN": SENS_N, "SENS_READ": SENS_READ, "VAL": VAL,
               "IMPL": {str(k): v for k, v in IMPL.items()}, "READ_ROW": READ_ROW, "NOPRIOR": NOPRIOR_ROW, "RC_END": RC_END, "INROW": INROW, "SB": SB, "RULE": rule_rows},
              open(os.environ["EOPTOLINK_ROWMAP"], "w"), ensure_ascii=False, indent=1)
