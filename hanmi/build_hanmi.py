# 韓美半導體 Hanmi Semiconductor（042700.KS）收支與估值模型 v0.2.2（AI 半導體 Project 步驟 5：設備與材料）
# 用法：python3 hanmi/build_hanmi.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年（FY2026＝2026-01～12）。金額：億韓元（₩億）；EPS：韓元。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
YRS = (2025, 2026, 2027, 2028, 2029)

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
SKH_F = only("skhynix/dist/*.xlsx"); SS_F = only("samsung/dist/*.xlsx"); MU_F = only("micron/dist/*.xlsx")
DSC_F = only("disco/dist/*.xlsx"); EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

ws_ = openpyxl.load_workbook(SKH_F, data_only=True)["Mkt"]
SHIP = {}   # HBM 出貨位元（EB），三情境各一列（markets/memory v0.4 起隨情境顆數係數變動）
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    if lab.startswith("出貨＝MIN"):
        SHIP[lab[-2]] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert sorted(SHIP) == ["A", "B", "C"] and all(v is not None for s in SHIP for v in SHIP[s][0])

def in_keys(f, prefixes):
    out = {}
    for r in openpyxl.load_workbook(f, data_only=True)["In"].iter_rows(min_row=4):
        k = r[0].value
        if isinstance(k, str) and k.startswith(prefixes):
            out[k] = ((r[3].value, r[4].value, r[5].value), f"{rel(f)}｜In!D{r[0].row}:F{r[0].row}")
    return out
MS = {m: in_keys(f, ("sh2",)) for m, f in (("sk", SKH_F), ("ss", SS_F), ("mu", MU_F))}
for m in MS: assert all(f"sh{y % 100}" in MS[m] for y in YRS), m
DS = in_keys(DSC_F, ("den_", "stk_", "hbp_"))
# v0.7.1：晶粒容量、堆疊層數、混合鍵合占比改讀 Siliconomics M22（與 Disco 模型同一值）
for _k in [k for k in DS if k.startswith(("den_", "stk_", "hbp_"))]:
    DS[_k] = (SIL.vals("hbm_" + _k), f"siliconomics｜SIF_hbm_{_k}")
assert all(f"{p}_{y}" in DS for p in ("den", "stk", "hbp") for y in range(25, 31))

ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
PKG = None
for r in ws_.iter_rows():
    if r[0].value == "SEMI 組裝與封裝設備":
        PKG = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert PKG
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則：各情境見頂年度（markets/equipment）

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：億韓元（₩億）；HBM 位元：EB；晶粒：百萬顆。每個輸入附兩個來源（來源2 找不到者寫「未找到」）"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司（hanmisemi.com）不出具年度指引；DART 原文在建置環境無法開啟，實績取自 DART 轉載的財務資料頁（valueline、stockanalysis）與新聞稿報導，原文待人工複核"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)

sec("A. HBM 出貨位元（markets/memory，經 SK 海力士模型成品 Mkt 頁；情境 A／B／C 各一列；唯讀）")
for s in SC:
    for j, y in enumerate(YRS):
        inp(f"hs{s}_{y % 100}", f"HBM 出貨位元 {y}（EB；{s}）", "EB", None, SHIP[s][0][j], None, "Derived", SHIP[s][1],
            "markets/memory v0.4：出貨＝MIN(需求 × 情境顆數係數, 供給)；2028–29 隨情境，2025–27 三情境相同" if j == 0 else "同上")
inp("g_tail", "CY2030 HBM 出貨位元相對 2029 的年增（市場層只到 2029；只用於 CY2029 的鍵合機購買）", "%", 0.0, 0.10, 0.20, "Assumed", "無：市場層只到 2029；鍵合機在產能需要前一年購入，CY2029 的購買要看 CY2030", "與 Disco 模型的 g_tail 同值（Assumed）")
sec("B. 晶粒、堆疊層數與混合鍵合（與 Disco 模型同一組假設，讀其成品；唯讀）")
for y in range(25, 31):
    v, src = DS[f"den_{y}"]
    inp(f"den_{y}", f"DRAM 核心晶粒容量（組合平均）20{y}", "Gb／顆", v[0], v[1], v[2], "Assumed", src, "JEDEC HBM4：24Gb 或 32Gb 晶粒；容量愈大，同樣位元需要的晶粒（＝TC 鍵合次數）愈少" if y == 25 else "同上")
for y in range(25, 31):
    v, src = DS[f"stk_{y}"]
    inp(f"stk_{y}", f"HBM 平均堆疊層數 20{y}", "層", v[0], v[1], v[2], "Assumed", src, "層數愈高，每顆晶粒的鍵合精度要求與時間愈高（見 beta）" if y == 25 else "同上")
inp("beta", "堆疊層數對每顆晶粒鍵合時間的影響（每多 4 層，鍵合時間 +beta）", "x", 0.0, 0.10, 0.25, "Assumed", "無實證：公司稱 12、16 層需新一代 TC 鍵合機（TC BONDER 4／4.5，2026-07 新聞稿），未給時間倍數", "第二來源：未找到。層數加權＝1＋beta×(層數 − 2025 層數)÷4")
for y in range(25, 31):
    v, src = DS[f"hbp_{y}"]
    inp(f"hbp_{y}", f"混合鍵合占 HBM 鍵合產能比例 20{y}（Disco 模型：低／基準／高）", "%", v[0], v[1], v[2], "Assumed", src,
        "TrendForce（2025-10-07，引 News1）：HBM4E 時代混合鍵合比重將大增；公司：16 層以上 HBM 約 2029 年量產時推出第二代混合鍵合機（2026-07-14 新聞稿）。本模型（v0.2.1）：混合鍵合為行業技術變數，與 Disco 同字母——CY2028 起 A 取高值（導入快）、B 取基準、C 取低值（導入慢），CY2027 以前三情境皆取基準；Sens 另列不納入與基準" if y == 25 else "同上")
inp("dlt", "已裝機鍵合產能每年汰換／升級比例（新機需求＝產能增量＋汰換）", "%/年", 0.05, 0.10, 0.15, "Assumed", "無：公司不揭露裝機量與機齡；HBM3E→HBM4→HBM4E 換代需新機（TC BONDER 4、4.5）", "與 Disco 模型的 dlt 同值（Assumed）")
sec("C. 三家記憶體廠的 HBM 營收市占（讀 SK 海力士、三星、美光模型成品 In 頁；以營收市占代替產能增量占比，Analogy）")
MN = {"sk": "SK 海力士", "ss": "三星", "mu": "美光"}
for m in ("sk", "ss", "mu"):
    for y in YRS:
        v, src = MS[m][f"sh{y % 100}"]
        inp(f"ms_{m}_{y % 100}", f"{MN[m]} HBM 營收市占 {y}", "%", None, v[1], None, "Interested-party", src, "各記憶體模型的 HBM 市占（賣方與公司說法）；本模型以 t＋1 年的市占（三家正規化）當 t 年鍵合產能增量的分配（Analogy：營收市占≈位元市占≈產能增量占比）" if y == 2025 else "同上")
sec("D. 韓美在各記憶體廠的 TC 鍵合機份額（份額是本模型的關鍵，皆為 Assumed／Interested-party，給區間）")
inp("hsk_25", "韓美占 SK 海力士 TC 鍵合機採購 2025", "%", 0.40, 0.50, 0.65, "Interested-party", "TrendForce（2025-12-12，引 The Elec，匿名來源）：韓美 2024 年前為 SK 海力士 HBM 獨家供應商；ASMPT 2024 年底加入，HBM4 約 50 套中約半數為 ASMPT；韓華 Semitech 2025 年初加入", "LS 證券（Interested-party，dealsite 2025-07-25）：2025 年韓美 TC 鍵合機營收中 SK 海力士占 40%、海外 60%")
for s, b in (("A", 0.45), ("B", 0.45), ("C", 0.35)):
    inp(f"hsk_{s}", f"韓美占 SK 海力士 TC 鍵合機採購（{s}；v0.2.1：2026–27 三情境共用 A 值，2028 起分情境）", "%", b - 0.10, b, b + 0.10, "Assumed", "韓美 2025-05（₩428 億）、2026-06（₩442 億，約 15 台）接 SK 海力士 HBM4 訂單；SK 海力士 2026-03 預計下約 100 台 HBM4 鍵合機，三家分配未公開（TrendForce 2025-12-12）", "C：韓華 Semitech 與 ASMPT 擴大份額（Assumed）" if s == "C" else "A、B：ASMPT 約四成、韓華約一成（Assumed）")
inp("hmu_25", "韓美占美光 TC 鍵合機採購 2025", "%", 0.70, 0.80, 0.90, "Interested-party", "美光 2025 年選韓美為「Top Supplier」（公司新聞稿 2026-02-09）；美光為 2024 年第二大客戶 ₩1,111 億（事業報告，newdaily 2025-04-02）", "TechInvestments（2026，引 CGSI）：美光以韓美為主，測試 ASMPT 與 BESI；BofA（dealsite 2026-07-22）：Q2 2026 主要來自美光加單")
for s, b in (("A", 0.75), ("B", 0.70), ("C", 0.55)):
    inp(f"hmu_{s}", f"韓美占美光 TC 鍵合機採購（{s}；v0.2.1：2026–27 三情境共用 A 值，2028 起分情境）", "%", b - 0.10, b, b + 0.10, "Assumed", "同上；美光擴大 HBM 後段產能（買下 AUO、力積電廠房；2026-07）", "C：ASMPT、BESI 進入美光（Assumed）" if s == "C" else "—")
inp("hss", "韓美占三星 TC 鍵合機採購（各年）", "%", 0.0, 0.0, 0.10, "Assumed", "三星主要採用自家 SEMES（TechInvestments 2026，引 CGSI）；韓美被視為可能但份額有限的供應商", "Sens 列 10%")
sec("E. FY2025 構成（校準錨）與其他業務")
inp("tcb_sh25", "TC 鍵合機（含中國）占 FY2025 營收", "%", 0.50, 0.60, 0.70, "Interested-party", "CGSI（TechInvestments 引用）：FY25F 中國出口 ₩589 億＝TC 鍵合機營收 17% → TC 鍵合機約 ₩3,460 億（Derived）÷ 營收 5,767 ≈ 60%", "公司不揭露產品別營收；第二來源：未找到")
inp("cn_sh25", "中國客戶占 FY2025 TC 鍵合機營收", "%", 0.10, 0.17, 0.25, "Interested-party", "CGSI（TechInvestments 引用）：FY25F 17%、FY26F 38%、FY27F 40%", "第二來源：未找到")
inp("a25_hbm", "FY2025 對三大記憶體廠的 HBM TC 鍵合機營收（Derived）", "₩億", None, "=a25_rev*tcb_sh25*(1-cn_sh25)", None, "Derived", "營收 × TC 鍵合機占比 ×（1 − 中國占比）", "其餘（中國 TC 鍵合機、MSVP、EMI Shield、2.5D／FC 鍵合機、零件服務）歸「其他」")
inp("g_o26", "其他業務 FY2026 成長", "%", 0.0, 0.15, 0.35, "Assumed", "無：CGSI 預估中國 TC 鍵合機 FY26 約 2.8 倍（Interested-party，不採作點估計）；MSVP 訂單增加（dealsite 2026-07-14，無數字）", "AI 前（2019–2022）公司總營收 1,204–3,732（Analogy：其他業務的量級）")
for s, b in (("A", 0.12), ("B", 0.08), ("C", -0.05)):
    inp(f"g_o_{s}", f"其他業務年成長（{s}；v0.2.1：FY2027 三情境共用 A 值，FY2028–29 分情境）", "%", b - 0.05, b, b + 0.05, "Assumed", "無：2.5D 封裝機（2026 推出三款）、HBF 用鍵合機、北美客戶", "C：中國出口受管制（Assumed）" if s == "C" else "—")
sec("F. FY2026 錨點（上半年實績＋下半年假設；第三季 2026-11 才公布）")
inp("h2f", "FY2026 下半年營收 ÷（2 × 第二季營收）", "x", 0.70, 0.90, 1.10, "Assumed", "第二季含美光超常加單（BofA，dealsite 2026-07-22）；SK 海力士 2026-03 預計約 100 台 HBM4 訂單（TrendForce）", "Interested-party 對照（不校準）：BofA 下半年約 ₩5,200 億＝1.04；共識全年 ₩7,280 億隱含 0.85")
inp("m_h2", "FY2026 下半年營業利益率", "%", 0.42, 0.48, 0.52, "Assumed", "第二季 51.9%（創高）；FY2025 43.6%", "基準取兩者之間")
inp("rev26", "FY2026 營收錨點（Derived）", "₩億", None, "=q1_rev+q2_rev+h2f*2*q2_rev", None, "Derived", "上半年實績 ＋ h2f × 2 × 第二季", "—")
inp("oi26", "FY2026 營業利益錨點（Derived）", "₩億", None, "=q1_oi+q2_oi+h2f*2*q2_rev*m_h2", None, "Derived", "上半年實績 ＋ 下半年營收 × m_h2", "—")
sec("G. 毛利、營業費用、淨利、股數")
inp("phi", "FY2025 營業費用中變動部分占比", "%", 0.30, 0.50, 0.70, "Assumed", "無：公司不揭露費用結構；FY2024 營業費用 589、FY2025 807，營收只增 3%（兩點推導不可用：費用增加非由營收驅動）", "營業費用＝固定（1−phi）× FY2025 ＋ 變動 phi × FY2025 ÷ FY2025 營收 × 營收")
inp("opex25", "營業費用 FY2025（毛利 − 營業利益）", "₩億", None, "=a25_gp-a25_oi", None, "Derived", "Q 頁", "—")
inp("opex_fix", "營業費用固定部分", "₩億", None, "=(1-phi)*opex25", None, "Derived", "—", "—")
inp("opex_var", "營業費用變動部分占營收", "%", None, "=phi*opex25/a25_rev", None, "Derived", "—", "—")
inp("gm_adj_C28", "情境 C：FY2028 毛利率額外調整（韓華、ASMPT 價格競爭）", "pt", -0.06, -0.03, 0.0, "Assumed", "無：韓美 2025-04 起 TC 鍵合機漲價 25%（newdaily 2025-04-02）；競爭加劇時可能回吐", "—")
inp("gm_adj_C29", "情境 C：FY2029 毛利率額外調整", "pt", -0.08, -0.05, 0.0, "Assumed", "無", "—")
inp("nioi", "淨利 ÷ 營業利益", "x", 0.72, "=a25_ni/a25_oi", 0.95, "Derived", "基準＝FY2025 實績 0.851；低＝FY2024–25 平均 0.72（FY2024 有金融資產評價損失）；高＝近 12 月（2026-06）0.95（含 SpaceX 等投資評價益，一次性）", "valueline 近 12 月：淨利 2,220、營業利益 2,342")
inp("shares", "股數（FY2025 淨利 ÷ FY2025 EPS，加權平均）", "百萬股", None, "=a25_ni/a25_eps*100", None, "Derived", "2,140.0 億 ÷ 2,231.7 ≈ 95.9M", "流通股：stockanalysis 94.85M（2026-10-02）、etfshopping 95.31M（2025）；庫藏股註銷使股數略降，本版不建模")
inp("fx", "匯率（只用於對帳的美元換算）", "₩/US$", *SIL.vals("fx_krw"), "Assumed", "無：本模型以韓元建模，匯率不進損益", "—")
sec("H. 估值、共識與對帳來源")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；本公司會計年度＝日曆年" if s == "A" else "同上")
inp("px", "股價（etfshopping 2026-10-23 收盤）", "₩", None, 188500, None, "Verified", "etfshopping.com/stock/042700：2026-10-23 收盤 188,500", "Daum 新聞頁股價欄 188,500（−1.41%）；valueline 10/23 16:55 顯示 189,500（可能含盤後），差 0.5%")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重：情境 {s}", "%", None, w, None, "Assumed", "Andy 可改（AI 半導體 Project 共用設定 0.3／0.5／0.2）", "—")
inp("cons_rev26", "共識營收 FY2026", "₩億", None, 7280, None, "Interested-party", "etfshopping 證券商共識（2026-10-19 更新頁面）：₩7,280 億", "上上人證券 7,850（heraldcorp 2026-06-09）；BofA 上半年＋下半年 ≈ 8,221")
inp("cons_oi26", "賣方營業利益 FY2026（上上人證券）", "₩億", None, 3694, None, "Interested-party", "上上人證券（heraldcorp 2026-06-09）", "共識營業利益：未找到")
inp("cons_eps26", "第三方預估 EPS FY2026（alphadistill）", "₩", None, 3400, None, "Interested-party", "alphadistill 2026-10-12（自家估計，非券商共識）", "stockanalysis 遠期 P/E 67.90 @ ₩272,000（2026-10-02）隱含遠期 EPS 約 4,006；FnGuide／Naver 共識頁無法開啟：未找到")
inp("cons_rev27", "第三方預估營收 FY2027（alphadistill）", "₩億", None, 9900, None, "Interested-party", "alphadistill 2026-10-12", "共識：未找到")
inp("cons_eps27", "第三方預估 EPS FY2027（alphadistill）", "₩", None, 4650, None, "Interested-party", "alphadistill 2026-10-12", "共識：未找到")
inp("tam25", "TC 鍵合機市場 2025（ASMPT 估，含邏輯用）", "US$M", None, 759, None, "Interested-party", "ASMPT（競爭者，TechInvestments 引用）：2025 約 $759M、2028 約 $1.6B", "第二來源：TechInsights 2025–30 年增約 13%（公司新聞稿引用，無金額）")
inp("tam28", "TC 鍵合機市場 2028（ASMPT 估）", "US$M", None, 1600, None, "Interested-party", "同上", "—")
inp("sh_claim", "韓美 HBM TC 鍵合機全球市占（公司引 TechInsights）", "%", None, 0.712, None, "Interested-party", "公司新聞稿（亞洲經濟 2026-02-09）：71.2%，第一", "年份未註明；第二來源：未找到")
inp("semi_pkg25", "SEMI 組裝與封裝設備 2025", "US$B", None, PKG[0][0], None, "Derived", PKG[1], "markets/equipment（SEMI 2025 由年增回推）")
inp("semi_pkg28", "SEMI 組裝與封裝設備 2028", "US$B", None, PKG[0][3], None, "Interested-party", PKG[1], "SEMI 2026-07 年中預測（設備商協會，傾向樂觀）")
inp("asp", "TC 鍵合機單價", "₩億／台", 25, 30, 40, "Interested-party", "韓經（heraldcorp 2026-06-09 引）：約 ₩30 億／台（₩442 億約 15 台）", "TrendForce（2025-12-12）：ASMPT 雙頭機約 ₩40 億／台；只用於台數換算顯示")

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
for col, w in zip("ABCDEFGHI", [12, 60, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（億韓元；EPS 為韓元）。會計年度＝日曆年。藍字＝輸入；黑字＝公式（Derived）。公司只公布初步營收與營業利益（DART 公平揭露）；淨利與 EPS 取自財務資料頁的近 12 月數字"
Q["A2"] = "FY2026 第三季（2026-07～09）：建置日 2026-10-25 尚未公布（去年第三季於 2025-10-31 公布；mydailybyte 列 2026-11-06，stockanalysis 估 2026-11-13）。公布後更新本頁 J 欄與 In 頁 h2f、m_h2"
hdr(Q, 3, ["項目", "FY2024", "25Q1", "25Q2", "25Q3", "25Q4", "FY2025", "26Q1", "26Q2", "26Q3（未公布）", "標記", "來源1", "來源2／說明"])
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGHI", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 11, tag); Q.cell(r, 12, s1); Q.cell(r, 13, s2)
qrow(4, "rev", "營收", [5589, "=G4-D4-E4-F4", 1800, 1662, 830, 5767, 509, 2511], "Verified（25Q1 Derived）",
     "公司初步實績（DART）：FY2025 5,767（亞洲經濟 2026-02-09）；25Q2 1,800（dealsite 2025-07-25）；26Q1 509（−65.5%）、26Q2 2,511（+39.5%；ddaily 2026-07-14）",
     "etfshopping 各季（DART 初步）：25Q3 1,662、25Q4 830、26Q2 2,512；butler.works：FY2024 5,589、FY2025 5,767；25Q1＝全年 − 其他三季（＝上半年 3,274 − 1,800）")
qrow(5, "gp", "毛利", [3143, None, None, None, None, 3320.6, None, None], "Verified", "valueline（DART 轉載）FY2025：營收 5,766.8、銷貨成本 2,446.1、毛利 3,320.6", "stockanalysis：FY2025 毛利 3,319.4、FY2024 3,143.0（季度毛利未找到）", "#,##0.0")
qrow(6, "oi", "營業利益", [2554, "=G6-D6-E6-F6", 863, 678, 276, 2513.9, 85, 1303], "Verified（25Q1 Derived）",
     "公司初步實績：FY2025 營益率 43.6%；26Q1 85（−87.9%）、26Q2 1,303（+51.0%，營益率 51.9%）", "valueline FY2025 2,513.9；stockanalysis 2,511.8；etfshopping 各季；25Q1＝上半年 1,559 − 863", "#,##0.0")
qrow(7, "oim", "營業利益率", [f"={c}6/{c}4" for c in "BCDEFGHI"], "Derived", "—", "公司：FY2025 43.6%、26Q2 51.9%", "0.0%")
qrow(8, "opx", "營業費用（毛利 − 營業利益）", ["=B5-B6", None, None, None, None, "=G5-G6", None, None], "Derived", "—", "valueline FY2025 銷管費 806.7", "#,##0.0")
qrow(9, "ni", "淨利（歸屬母公司）", [1526, None, None, None, None, 2140.0, 190, None], "Verified", "valueline（近 12 月表）FY2025 2,140.0；26Q1 190（etfshopping）", "butler.works：FY2024 1,526、FY2025 2,140；stockanalysis FY2025 2,140.1；25 各季、26Q2 單季未找到（近 12 月 2026-06 為 2,220.3）", "#,##0.0")
qrow(10, "eps", "EPS（韓元）", [1584, None, None, None, None, 2231.7, None, None], "Verified（來源不一致）", "valueline FY2025 2,231.7（DART 轉載）", "stockanalysis 2,248、etfshopping 2,256（差 0.7–1.1%，口徑未解：基本／稀釋或加權股數）；FY2024 stockanalysis 1,584、etfshopping 1,589", "#,##0.0")
qrow(11, "ck", "檢查：四季合計 − 全年（營收、營業利益；B、C 欄）", ["=SUM(C4:F4)-G4", "=SUM(C6:F6)-G6"], "—", "—", "—", "0.0")
Q["J4"] = "未公布"; Q["J6"] = "未公布"
for nm, rr in (("rev", 4), ("gp", 5), ("oi", 6), ("ni", 9), ("eps", 10)):
    defname(f"a25_{nm}", f"Q!$G${rr}")
for nm, rr in (("rev", 4), ("oi", 6)):
    defname(f"q1_{nm}", f"Q!$H${rr}"); defname(f"q2_{nm}", f"Q!$I${rr}")
Q.cell(13, 1, "客戶別營收（事業報告）").font = BOLD
hdr(Q, 14, ["客戶", "FY2024", "占營收", "", "", "", "FY2025", "", "", "", "標記", "來源1", "來源2／說明"])
for i, (nm, v, n25) in enumerate((("SK 海力士", 2996, "LS 證券估：TC 鍵合機營收的 40%（Interested-party）"), ("美光", 1111, "未找到"), ("其他（中國、OSAT 等）", "=B4-B15-B16", "CGSI 估中國 ₩589 億（Interested-party）")), 15):
    Q.cell(i, 1, nm); Q.cell(i, 2, v).number_format = "#,##0"
    if not isinstance(v, str): Q.cell(i, 2).font = BLUE
    Q.cell(i, 3, f"=B{i}/B4").number_format = "0.0%"; Q.cell(i, 7, n25)
    Q.cell(i, 11, "Verified" if i < 17 else "Derived"); Q.cell(i, 12, "公司 2024 事業報告（newdaily 2025-04-02 報導）" if i < 17 else "營收 − 上兩列")
Q.column_dimensions["A"].width = 46
for c in "BCDEFGHIJ": Q.column_dimensions[c].width = 11
Q.column_dimensions["K"].width = 22; Q.column_dimensions["L"].width = 80; Q.column_dimensions["M"].width = 80
Q.freeze_panes = "B4"

# ---------------- Units：鍵合產能需求（情境 A，示範；PL 引擎每個區塊自行計算） ----------------
COL = {y: chr(ord("C") + y - 2025) for y in range(2025, 2031)}   # C＝2025 … H＝2030
HB = {"A": "hbp_{y}_hi", "B": "hbp_{y}", "C": "hbp_{y}_lo"}   # v0.2.1：混合鍵合為行業技術變數，與 Disco 同字母（A 導入快＝高值、B 基準、C 慢＝低值）；CY2027 以前三情境皆讀基準
DEFP = dict(scn="A", den="", stk="", beta="beta", dlt="dlt", gt="g_tail", hb=None, hsk=None, hmu=None, hss="hss",
            hsk25="hsk_25", hmu25="hmu_25", a25h="a25_hbm", go26="g_o26", go=None, rev26="rev26", oi26="oi26",
            F="opex_fix", v="opex_var", nioi="nioi", gmadj=True)

def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    hb = P["hb"] if P["hb"] is not None else HB[s]
    hby = lambda y: (P["hb"] if P["hb"] is not None else ("hbp_{y}" if y <= 2027 else HB[s])).format(y=y % 100)   # 情境分歧自 CY2028
    hb28p = (P["hb"] if P["hb"] is not None else "hbp_{y}").format(y=28)   # CY2027 購機的共同前提：CY2028 混合鍵合取基準
    hsk = P["hsk"] or f"hsk_{s}"; hmu = P["hmu"] or f"hmu_{s}"; go = P["go"] or f"g_o_{s}"
    hsk0 = P["hsk"] or "hsk_A"; hmu0 = P["hmu"] or "hmu_A"; go0 = P["go"] or "g_o_A"   # v0.2.1：CY2026–27 三情境共用情境 A 值，公司特有變數自 CY2028 起分歧
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "2025", "2026", "2027", "2028", "2029", "2030", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for y, f in fs.items():
            ws[f"{COL[y]}{r}"] = f; ws[f"{COL[y]}{r}"].number_format = fmt
        ws.cell(r, 9, note); R[key] = r; r += 1
    ws.cell(r, 1, "A. 日曆年鍵合產能需求（本公司會計年度＝日曆年）").font = BOLD; r += 1
    row("ship", "HBM 出貨位元", "EB", {**{y: f"=hs{s}_{y % 100}" for y in YRS}, 2030: f"=G{r}*(1+{P['gt']})"}, note="markets/memory；2030＝2029 ×（1＋g_tail）")
    row("den", "DRAM 晶粒容量", "Gb", {y: f"=den_{y % 100}{P['den']}" for y in range(2025, 2031)}, fmt="0.0")
    row("dies", "核心晶粒數＝位元 × 8000 ÷ 容量", "百萬顆", {y: f"={COL[y]}{R['ship']}*8000/{COL[y]}{R['den']}" for y in range(2025, 2031)}, fmt="#,##0", note="1 EB＝8×10⁹ Gb；每顆核心晶粒鍵合一次")
    row("stk", "平均堆疊層數", "層", {y: f"=stk_{y % 100}{P['stk']}" for y in range(2025, 2031)}, fmt="0.0")
    row("lf", "層數加權（每顆晶粒鍵合時間）", "x", {y: f"=1+{P['beta']}*({COL[y]}{R['stk']}-$C{R['stk']})/4" for y in range(2025, 2031)}, fmt="0.000")
    row("hb", "混合鍵合占比", "%", {y: "=" + hby(y) for y in range(2025, 2031)}, fmt="0.0%", note={"A": "情境 A：CY2028 起 Disco 模型高值（導入快）", "B": "情境 B：Disco 模型基準值", "C": "情境 C：CY2028 起 Disco 模型低值（導入慢）"}.get(s, "") if P["hb"] is None else "")
    row("cap", "TC 鍵合產能需求（晶粒 × 層數加權 ×（1 − 混合鍵合））", "百萬顆當量", {y: f"={COL[y]}{R['dies']}*{COL[y]}{R['lf']}*(1-{COL[y]}{R['hb']})" for y in range(2025, 2031)}, fmt="#,##0")
    nxt = lambda y: (f"hsA_28*8000/{COL[2028]}{R['den']}*{COL[2028]}{R['lf']}*(1-{hb28p})" if y == 2027 else f"{COL[y+1]}{R['cap']}")
    row("new", "新機需求＝MAX(0, 下一年產能 − 本年)＋汰換 × 本年（CY2027 的下一年產能用三情境共同前提）", "百萬顆當量", {y: f"=MAX(0,{nxt(y)}-{COL[y]}{R['cap']})+{P['dlt']}*{COL[y]}{R['cap']}" for y in YRS}, fmt="#,##0", note="鍵合機在產能需要前一年購入；v0.2.1：CY2027 購機的 CY2028 產能＝情境 A 的 HBM 位元（hsA_28）× 混合鍵合基準，三情境相同（與 markets/equipment 2027 三情境相同、記憶體供給三情境相同一致）；CY2028 起照讀本情境")
    tot = lambda y: "+".join(f"ms_{m}_{min(y + 1, 2029) % 100}" for m in ("sk", "ss", "mu"))
    for m, nm in (("sk", "SK 海力士"), ("ss", "三星"), ("mu", "美光")):
        row("w" + m, f"{nm} 占產能增量（t＋1 年 HBM 市占，三家正規化）", "%", {y: f"=ms_{m}_{min(y + 1, 2029) % 100}/({tot(y)})" for y in YRS}, fmt="0.0%")
    shk = {2025: P["hsk25"], 2026: hsk0, 2027: hsk0, 2028: hsk, 2029: hsk}; shm = {2025: P["hmu25"], 2026: hmu0, 2027: hmu0, 2028: hmu, 2029: hmu}
    row("hw", "韓美加權份額＝Σ 占比 × 韓美在該廠份額", "%", {y: f"={COL[y]}{R['wsk']}*{shk[y]}+{COL[y]}{R['wss']}*{P['hss']}+{COL[y]}{R['wmu']}*{shm[y]}" for y in YRS}, fmt="0.0%")
    row("addr", "韓美可得新機需求＝新機需求 × 加權份額", "百萬顆當量", {y: f"={COL[y]}{R['new']}*{COL[y]}{R['hw']}" for y in YRS}, fmt="#,##0")
    ws.cell(r, 1, "B. 損益（會計年度；FY2026＝錨點）").font = BOLD; r += 1
    row("oth", "其他業務營收（中國 TC 鍵合機、MSVP、EMI、2.5D 等）", "₩億", {2025: f"=a25_rev-{P['a25h']}", 2026: f"=C{r}*(1+{P['go26']})", **{y: f"={COL[y-1]}{r}*(1+{go0 if y == 2027 else go})" for y in (2027, 2028, 2029)}}, fmt="#,##0", note="v0.2.1：FY2027 三情境共用 g_o_A，FY2028 起讀 g_o_{情境}")
    row("hbm", "HBM TC 鍵合機營收（三大記憶體廠）", "₩億", {2025: f"={P['a25h']}", 2026: f"={P['rev26']}-D{R['oth']}", **{y: f"=$C{r+1}*{COL[y]}{R['addr']}" for y in (2027, 2028, 2029)}}, fmt="#,##0", note="FY2026＝錨點營收 − 其他；FY2027 起＝k × 可得新機需求")
    row("k", "校準係數 k＝(FY2025＋FY2026 HBM 營收) ÷ (兩年可得新機需求)", "₩億／百萬顆", {2025: f"=(C{R['hbm']}+D{R['hbm']})/(C{R['addr']}+D{R['addr']})"}, fmt="0.000", note="兩年平均，平滑訂單的季節與交期波動；不使用共識")
    row("rev", "營收", "₩億", {2025: "=a25_rev", **{y: f"={COL[y]}{R['oth']}+{COL[y]}{R['hbm']}" for y in range(2026, 2030)}}, fmt="#,##0")
    row("hsh", "HBM TC 鍵合機占營收", "%", {y: f"={COL[y]}{R['hbm']}/{COL[y]}{R['rev']}" for y in YRS}, fmt="0.0%")
    row("opx", "營業費用", "₩億", {2025: "=opex25", **{y: f"={P['F']}+{P['v']}*{COL[y]}{R['rev']}" for y in range(2026, 2030)}}, fmt="#,##0", note="固定＋變動 × 營收")
    adj = {2027: "0", 2028: "gm_adj_C28" if (s == "C" and P["gmadj"]) else "0", 2029: "gm_adj_C29" if (s == "C" and P["gmadj"]) else "0"}
    row("gm", "毛利率", "%", {2025: "=a25_gp/a25_rev", 2026: f"=({P['oi26']}+D{R['opx']})/D{R['rev']}", **{y: f"=$D{r}+{adj[y]}" for y in (2027, 2028, 2029)}}, fmt="0.0%", note="FY2026＝錨點營業利益反推（Derived）；FY2027 起沿用（情境 C 另扣價格競爭）")
    row("gp", "毛利", "₩億", {y: f"={COL[y]}{R['rev']}*{COL[y]}{R['gm']}" for y in YRS}, fmt="#,##0")
    row("oi", "營業利益", "₩億", {2025: "=a25_oi", 2026: f"={P['oi26']}", **{y: f"={COL[y]}{R['gp']}-{COL[y]}{R['opx']}" for y in (2027, 2028, 2029)}}, fmt="#,##0")
    row("oim", "營業利益率", "%", {y: f"={COL[y]}{R['oi']}/{COL[y]}{R['rev']}" for y in YRS}, fmt="0.0%")
    row("ni", "淨利", "₩億", {2025: "=a25_ni", **{y: f"={COL[y]}{R['oi']}*{P['nioi']}" for y in range(2026, 2030)}}, fmt="#,##0")
    row("eps", "EPS", "₩", {2025: "=a25_eps", **{y: f"={COL[y]}{R['ni']}*100/shares" for y in range(2026, 2030)}}, fmt="#,##0")
    row("mkt", "三大廠 HBM TC 鍵合機市場（全部供應商）＝k × 新機需求", "₩億", {y: f"=$C{R['k']}*{COL[y]}{R['new']}" for y in YRS}, fmt="#,##0", note="對帳用")
    row("usd", "同上（美元）", "US$M", {y: f"={COL[y]}{R['mkt']}/fx*100" for y in YRS}, fmt="#,##0")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 韓美半導體損益（₩億；FY2025 實績，FY2026 錨點＝上半年實績＋下半年假設，FY2027–FY2029 為模型）。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 40
TITLES = {"A": "情境 A：基準顆數；混合鍵合導入快（CY2028 起 Disco 高值）", "B": "情境 B：2028 顆數 ×0.8、2029 ×0.85；混合鍵合取 Disco 基準", "C": "情境 C：2028–29 顆數 ×0.6；混合鍵合導入慢（CY2028 起 Disco 低值）；韓華、ASMPT 搶份額與價格壓力"}
for s in SC:
    PB[s] = engine(PL, r0, TITLES[s], scn=s); r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "₩億", "#,##0"), ("hbm", "  HBM TC 鍵合機", "₩億", "#,##0"), ("oth", "  其他業務", "₩億", "#,##0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "₩億", "#,##0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "₩億", "#,##0"), ("eps", "EPS", "₩", "#,##0")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in YRS:
            PL[f"{COL[y]}{sr}"] = f"={COL[y]}{PB[s][key]}"; PL[f"{COL[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：FY2026 三情境相同（營收、EPS 差額合計，應為 0）").font = BOLD
PL[f"C{sr}"] = f"=ABS(D{SUMR[('A','rev')]}-D{SUMR[('B','rev')]})+ABS(D{SUMR[('A','rev')]}-D{SUMR[('C','rev')]})+ABS(D{SUMR[('A','eps')]}-D{SUMR[('C','eps')]})"
PL.column_dimensions["A"].width = 56; PL.column_dimensions["B"].width = 14; PL.column_dimensions["I"].width = 50
for c in "CDEFGH": PL.column_dimensions[c].width = 12
PL["H3"] = "說明"; PL["H3"].font = BOLD; PL["H3"].fill = HDR
PL["H5"] = "FY2026 為錨點（上半年實績＋下半年 h2f、m_h2），三情境相同；情境差異從 FY2027 起"

# ---------------- Units 頁：情境 A 的產能需求摘要（引用 PL 區塊） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — HBM 位元 → 核心晶粒 → TC 鍵合產能需求 → 新機需求 → 韓美可得需求（日曆年；三情境，引用 PL 區塊）"
hdr(UN, 3, ["項目", "單位", "2025", "2026", "2027", "2028", "2029", "2030", "說明"])
ur = 4; UR = {}
for s in SC:
    UN.cell(ur, 1, f"情境 {s}").font = BOLD; ur += 1
    for key, lab, unit, fmt in (("ship", "HBM 出貨位元", "EB", "0.00"), ("dies", "核心晶粒數", "百萬顆", "#,##0"), ("cap", "TC 鍵合產能需求", "百萬顆當量", "#,##0"),
                                ("new", "新機需求", "百萬顆當量", "#,##0"), ("hw", "韓美加權份額", "%", "0.0%"), ("addr", "韓美可得新機需求", "百萬顆當量", "#,##0"), ("mkt", "三大廠 HBM TC 鍵合機市場", "₩億", "#,##0")):
        UN.cell(ur, 1, "  " + lab); UN.cell(ur, 2, unit)
        last = 2031 if key in ("ship", "dies", "cap") else 2030
        for y in range(2025, last):
            UN[f"{COL[y]}{ur}"] = f"=PL!{COL[y]}{PB[s][key]}"; UN[f"{COL[y]}{ur}"].number_format = fmt
        UR[(s, key)] = ur; ur += 1
    ur += 1
UN.cell(ur, 1, "台數換算（顯示用；單價為 Interested-party，只影響台數，不影響營收）").font = BOLD; ur += 1
for s in SC:
    UN.cell(ur, 1, f"  情境 {s}：三大廠 TC 鍵合機台數＝市場 ÷ 單價"); UN.cell(ur, 2, "台")
    for y in YRS: UN[f"{COL[y]}{ur}"] = f"={COL[y]}{UR[(s, 'mkt')]}/asp"; UN[f"{COL[y]}{ur}"].number_format = "#,##0"
    UR[(s, "units")] = ur; ur += 1
UN.cell(ur, 9, "對照：SK 海力士 2026-03 預計下約 100 台 HBM4 鍵合機（TrendForce 引 The Elec）")
UN.column_dimensions["A"].width = 52; UN.column_dimensions["I"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（第 4 節）。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "P/E 定義：日曆年底收盤價 ÷ 當年 EPS（本公司會計年度＝日曆年）。兩個來源各自取分位再平均（tools/valuation_rule.py 的多來源慣例）"
V.cell(4, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, 5, ["項目", "數值", "", "", "說明"])
VAL = {}  # 第 5 節規則值列（之後填）
V.cell(6, 1, "目標價（12 個月，情境加權）"); V.cell(7, 1, "上檔空間（目標價 ÷ 現價 − 1）")
V.cell(8, 1, "情境 A 隱含價（樂觀）"); V.cell(9, 1, "情境 B 隱含價（基準）"); V.cell(10, 1, "情境 C 隱含價（下檔）")
V.cell(11, 1, "情境權重 A／B／C")
for j, s in enumerate(SC): V.cell(11, 2 + j, f"=tp_w{s}").number_format = "0%"
V.cell(11, 5, "In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）")
V.cell(12, 1, "目標價方法")
V.cell(12, 2, "P/E：各情境 FY2027 EPS（模型）× 依循環位置的十年（2016–2025 年底）P/E 分位（設備股共用規則：A、B 中段＝中位數，C 見頂＝25 分位），再以 0.3／0.5／0.2 加權。偏誤：十年樣本有 7 年在 HBM 需求出現前（2016–2022，倍數 7–35 倍），會把分位往下拉；2024–25 年 50 倍以上的高成長期只占兩年。照規則計算，不另換倍數")
V.cell(4, 5, "")

V.cell(14, 1, "1. 十年 P/E（年底）— 兩個來源").font = BOLD
hdr(V, 15, ["年底", "來源1：companiesmarketcap", "來源2：stockanalysis／券商報告", "兩者差距", "標記", "來源2 出處"])
CMC = {2016: 23.2, 2017: 55.0, 2018: 7.16, 2019: 19.7, 2020: 17.2, 2021: 34.8, 2022: 12.1, 2023: 21.8, 2024: 51.5, 2025: 56.5}
S2 = {2016: (None, "未找到（2016–2017 年的券商報告與財務頁未取得）"), 2017: (None, "未找到"),
      2018: (9.7, "SK 證券 2021-12-22 報告（年底 P/E 9.7）"), 2019: (24.1, "SK 證券 2021-12-22 報告（24.1）"),
      2020: (18.78, "Hana 證券 2023-02-13 報告（18.78；SK 證券 18.8）"), 2021: (17.93, "stockanalysis（年底市值 ÷ 淨利；Hana 18.06）"),
      2022: (12.13, "stockanalysis（Hana 11.59）"), 2023: (22.48, "stockanalysis"), 2024: (51.76, "stockanalysis"), 2025: (56.47, "stockanalysis")}
for i, y in enumerate(range(2016, 2026), 16):
    V.cell(i, 1, f"{y}-12"); V.cell(i, 2, CMC[y]).font = BLUE
    if S2[y][0] is not None: V.cell(i, 3, S2[y][0]).font = BLUE
    V.cell(i, 4, f'=IF(C{i}="","",B{i}-C{i})')
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 5, "Interested-party（資料商／券商）"); V.cell(i, 6, S2[y][1])
V.cell(26, 1, "說明：companiesmarketcap 以美元計價的年底股價 ÷ EPS；2021 年 34.8 與另兩個來源（17.9、18.1）差異大，原因未解（該站 2021 年 EPS 疑有口徑問題），照列不剔除。2022 年 1 比 1 無償配股，各來源 EPS 皆為調整後")
V.cell(28, 1, "2. 分位").font = BOLD
hdr(V, 29, ["統計", "採用（兩來源平均）", "來源1", "來源2", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 30):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 3, f"=PERCENTILE(B16:B25,{p})"); V.cell(i, 4, f"=PERCENTILE(C16:C25,{p})"); V.cell(i, 2, f"=ROUND(AVERAGE(C{i},D{i}),2)")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(33, 1, "年數"); V["C33"] = "=COUNT(B16:B25)"; V["D33"] = "=COUNT(C16:C25)"
defname("pe_peak", "Val!$B$30"); defname("pe_mid", "Val!$B$31"); defname("pe_trough", "Val!$B$32")

V.cell(35, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(36, 1, CP.rule_text(PEAK) + "。本公司會計年度＝日曆年，FY(t) 直接對照日曆年 t")
hdr(V, 37, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "規則"])
for i, s in enumerate(SC, 38):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "tools/cycle_position.py")
V.cell(42, 1, "5. 隱含股價（韓元）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 43, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 44
for s in SC:
    for y, typ, nm in ((2027, f"E{38 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{COL[y]}{SUMR[(s,'eps')]}").number_format = "#,##0"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{38 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
# 第 0 節公式
for i, s in zip((8, 9, 10), SC):
    V.cell(i, 2, f"=F{VAL[s]}").number_format = "#,##0"; V.cell(i, 5, f"Val!F{VAL[s]}：FY2027 EPS × 規則倍數")
V.cell(6, 2, "=tp_wA*B8+tp_wB*B9+tp_wC*B10").number_format = "#,##0"; V.cell(6, 5, "權重 × 三情境隱含價")
V.cell(7, 2, "=B6/px-1").number_format = "0.0%"; V.cell(7, 5, "現價見第 6 節")
for c in range(1, 6): V.cell(6, c).fill = KEY
V.cell(vr + 1, 1, "6. 目前倍數").font = BOLD
hdr(V, vr + 2, ["項目", "A", "B", "C", "備註"])
fr = vr + 3
V.cell(fr, 1, "現價（etfshopping 2026-10-23 收盤）"); V.cell(fr, 2, "=px").number_format = "#,##0"; fr += 1
for y in (2026, 2027, 2028):
    V.cell(fr, 1, f"現價 ÷ FY{y} EPS（模型）")
    for j, s in enumerate(SC): V.cell(fr, 2 + j, f"=px/PL!{COL[y]}{SUMR[(s,'eps')]}").number_format = "0.0"
    fr += 1
V.cell(fr, 1, "近 12 月 P/E（stockanalysis，2026-10-02 @ ₩272,000）"); V.cell(fr, 2, 116.68).font = BLUE; V.cell(fr, 5, "現價 ÷ 近 12 月 EPS 2,331＝"); V.cell(fr, 6, "=px/2331.18").number_format = "0.0"; fr += 1
V.cell(fr, 1, "現價 ÷ 第三方預估 FY2026 EPS（alphadistill 3,400）"); V.cell(fr, 2, "=px/cons_eps26").number_format = "0.0"; fr += 1
V.cell(fr, 1, "現價隱含倍數 ÷ 規則倍數（情境 A，FY2027）"); V.cell(fr, 2, f"=B{fr-4}/E38").number_format = "0.00"; V.cell(fr, 5, ">1＝現價高於規則給的倍數")
V.column_dimensions["A"].width = 50
for c in "BCDEFG": V.column_dimensions[c].width = 16
V.column_dimensions["F"].width = 60

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎；校準係數 k 在各區塊內重算）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ₩億", "FY2027 EPS ₩", "相對基準", "FY2028 營收 ₩億", "FY2028 EPS ₩", "相對基準"])
cases = [("SK 海力士份額 −10pt", "35%", dict(hsk="hsk_A_lo")), ("SK 海力士份額 +10pt", "55%", dict(hsk="hsk_A_hi")),
         ("美光份額 −10pt", "65%", dict(hmu="hmu_A_lo")), ("美光份額 +10pt", "85%", dict(hmu="hmu_A_hi")),
         ("三星份額 10%", "韓美打進三星", dict(hss="hss_hi")),
         ("混合鍵合 不納入", "各年 0%", dict(hb="0")), ("混合鍵合 Disco 基準", "2028 5%、2029 15%、2030 25%", dict(hb="hbp_{y}")),
         ("晶粒容量 低", "各年取低值（晶粒較多）", dict(den="_lo")), ("晶粒容量 高", "各年取高值（晶粒較少）", dict(den="_hi")),
         ("堆疊層數 低", "", dict(stk="_lo")), ("堆疊層數 高", "", dict(stk="_hi")),
         ("層數對鍵合時間 0", "beta＝0", dict(beta="beta_lo")), ("層數對鍵合時間 0.25", "beta＝0.25", dict(beta="beta_hi")),
         ("汰換比例 5%", "", dict(dlt="dlt_lo")), ("汰換比例 15%", "", dict(dlt="dlt_hi")),
         ("FY2026 下半年 ×0.7", "錨點營收較低 → k 較低", dict(rev26="(q1_rev+q2_rev+h2f_lo*2*q2_rev)", oi26="(q1_oi+q2_oi+h2f_lo*2*q2_rev*m_h2)")),
         ("FY2026 下半年 ×1.1", "錨點營收較高 → k 較高", dict(rev26="(q1_rev+q2_rev+h2f_hi*2*q2_rev)", oi26="(q1_oi+q2_oi+h2f_hi*2*q2_rev*m_h2)")),
         ("下半年營益率 42%", "毛利率基準隨之下降", dict(oi26="(q1_oi+q2_oi+h2f*2*q2_rev*m_h2_lo)")), ("下半年營益率 52%", "", dict(oi26="(q1_oi+q2_oi+h2f*2*q2_rev*m_h2_hi)")),
         ("TC 鍵合機占 FY2025 營收 50%", "HBM 基期較低、其他較高", dict(a25h="(a25_rev*tcb_sh25_lo*(1-cn_sh25))")), ("TC 鍵合機占 FY2025 營收 70%", "", dict(a25h="(a25_rev*tcb_sh25_hi*(1-cn_sh25))")),
         ("其他業務成長 低", "7%", dict(go="g_o_A_lo")), ("其他業務成長 高", "17%", dict(go="g_o_A_hi")),
         ("營業費用變動比例 30%", "", dict(F="(1-phi_lo)*opex25", v="phi_lo*opex25/a25_rev")), ("營業費用變動比例 70%", "", dict(F="(1-phi_hi)*opex25", v="phi_hi*opex25/a25_rev")),
         ("淨利÷營業利益 0.72", "", dict(nioi="nioi_lo")), ("淨利÷營業利益 0.95", "含一次性投資評價益，不當基準", dict(nioi="nioi_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}E{R['rev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"F{i}"] = f"={pre}F{R['rev']}"; SE[f"G{i}"] = f"={pre}F{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CDFG": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, 60 + n * 40, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "2028 ×0.8、2029 ×0.85；混合鍵合基準；份額與其他成長 B", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "2028–29 ×0.6；混合鍵合低值（導入慢）；份額回落、毛利率下調", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：鍵合機比產能早一年買，FY2027 的需求看 CY2027→2028 的產能增量，所以顆數情境（2028 起）直接影響 FY2027。晶粒容量的影響大，是因為同樣位元下，容量提高會減少需要鍵合的晶粒數；校準係數 k 由 FY2025–26 兩年營收推得，各案例內重算。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGH": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、賣方、競爭者估的市場規模、SEMI。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L = chr(64 + cv)
        RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y=2026: (lambda s: f"=PL!{COL[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2026 vs 共識與賣方（校準未使用）").font = BOLD; rr_ += 1
rec("營收", "=cons_rev26", pl("rev"), "etfshopping 共識 ₩7,280 億", "錨點＝上半年 3,020 ＋ 0.9 × 2 × 第二季")
rec("營業利益", "=cons_oi26", pl("oi"), "上上人證券（2026-06-09）")
rec("EPS（韓元）", "=cons_eps26", pl("eps"), "alphadistill（2026-10-12，非券商共識）", "stockanalysis 遠期 EPS 約 4,006（隱含）")
RC.cell(rr_, 1, "B. FY2027 vs 第三方預估").font = BOLD; rr_ += 1
rec("營收", "=cons_rev27", pl("rev", 2027), "alphadistill")
rec("EPS（韓元）", "=cons_eps27", pl("eps", 2027), "alphadistill")
RC.cell(rr_, 1, "C. 市場規模與市占").font = BOLD; rr_ += 1
rec("TC 鍵合機市場 CY2025（US$M；模型＝三大廠 HBM，不含邏輯與中國）", "=tam25", lambda s: f"=PL!C{PB[s]['usd']}", "ASMPT 估（含邏輯用 TCB）", "口徑：模型只含三大記憶體廠的 HBM 鍵合")
rec("TC 鍵合機市場 CY2028（US$M）", "=tam28", lambda s: f"=PL!F{PB[s]['usd']}", "ASMPT 估")
rec("韓美 HBM TC 鍵合機市占 CY2025（模型＝加權份額）", "=sh_claim", lambda s: f"=PL!C{PB[s]['hw']}", "公司引 TechInsights 71.2%", "差距以百分點計", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "D. 三大廠 HBM TC 鍵合機市場 vs SEMI 組裝與封裝設備（markets/equipment）").font = BOLD; rr_ += 1
rec("CY2025：模型市場（US$B）vs SEMI 組裝與封裝", "=semi_pkg25", lambda s: f"=PL!C{PB[s]['usd']}/1000", "SEMI（equipment 成品）", "差距＝占 SEMI 組裝封裝的比例 − 1", fmt="0.00")
rec("CY2028：同上", "=semi_pkg28", lambda s: f"=PL!F{PB[s]['usd']}/1000", "SEMI 2026-07", "", fmt="0.00")
RC.cell(rr_, 1, "E. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FY2026 上半年營收 ÷ 錨點全年"); RC.cell(rr_, 3, f"=(q1_rev+q2_rev)/PL!D{SUMR[('A','rev')]}").number_format = "0.0%"; rr_ += 1
RC.cell(rr_, 1, "CY2026 三大廠 TC 鍵合機台數（模型 A，市場 ÷ 單價）"); RC.cell(rr_, 3, f"=Units!D{UR[('A','units')]}").number_format = "#,##0"; RC.cell(rr_, 10, "對照：SK 海力士 2026-03 預計約 100 台（The Elec）；HBM4 現有約 50 套"); rr_ += 1
RC.column_dimensions["A"].width = 60
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（DART 原文在建置環境無法開啟；Naver、FnGuide 頁面被擋；數字取自新聞稿報導與財務資料頁，原文待人工複核）"; SO["A1"].font = BOLD
srcs = [("亞洲經濟：2025 年創社最大營收 5,767、營益率 43.6%、TC 鍵合機市占 71.2%（2026-02-09）", "https://view.asiae.co.kr/en/article/2026020908430125470"),
 ("ddaily：2026 第二季營收 2,511、營業利益 1,303（2026-07-14）", "https://www.ddaily.co.kr/page/view/2026071411214300568"),
 ("THE FACT：第一季 509／85、目標營收撤回（2026-08）", "https://news.tf.co.kr/read/economy/2350287.htm"),
 ("dealsite：BofA 下半年約 ₩5,200 億、美光加單（2026-07-22）", "https://dealsite.co.kr/articles/165650"),
 ("dealsite：2025 第二季、LS 證券客戶比重（2025-07-25）", "https://dealsite.co.kr/articles/145326"),
 ("heraldcorp：SK 海力士 ₩442 億訂單、上上人證券預估（2026-06-09）", "https://biz.heraldcorp.com/article/10778245"),
 ("newdaily：2024 客戶別營收（SK 海力士 2,996、美光 1,111）、漲價 25%（2025-04-02）", "https://biz.newdaily.co.kr/site/data/html/2025/04/02/2025040200057.html"),
 ("etfshopping：各季初步實績、共識營收、2026-10-23 收盤", "https://etfshopping.com/stock/042700/earnings"),
 ("valueline：近 12 月損益（毛利、銷管費、淨利、EPS）", "https://www.valueline.co.kr/finance/income/042700"),
 ("stockanalysis：年度損益、年底市值 P/E（2021–2025）", "https://stockanalysis.com/quote/krx/042700/financials/ratios/"),
 ("butler.works：年度營收、營業利益、淨利", "https://www.butler.works/ko/companies/00161383"),
 ("companiesmarketcap：年底 P/E（2016–2025）", "https://companiesmarketcap.com/hanmi-semiconductor/pe-ratio/"),
 ("SK 證券 2021-12-22 報告（2018–2020 P/E）", "https://file.alphasquare.co.kr/media/pdfs/company-report/SK20211222한미반도체.pdf"),
 ("Hana 證券 2023-02-13 報告（2020–2022 P/E）", "https://files-scs.pstatic.net/2023/02/13/k2cjFZmqGF/한미반도체［042700］_20230213_Hana_840814.pdf"),
 ("TrendForce：SK 海力士向 ASMPT 下 HBM4 TC 鍵合機訂單、專利訴訟（2025-12-12）", "https://www.trendforce.com/news/2025/12/12/news-sk-hynix-reportedly-places-hbm4-tc-bonder-order-with-asmpt-as-hanmi-hanwha-patent-clash-heats-up/"),
 ("TrendForce：混合鍵合機市場 2028 約 $2B（2025-10-07）", "https://www.trendforce.com/news/2025/10/07/news-hybrid-bonder-equipment-market-reportedly-near-2b-by-2028-as-korean-makers-race-to-next-gen-hbm-tools/"),
 ("TechInvestments：HBM 鍵合設備展望（ASMPT TAM、CGSI 中國比重）", "https://www.techinvestments.io/p/hbmdie-bonding-semicap-outlook"),
 ("alphadistill：FY26E／FY27E 預估（2026-10-12）", "https://alphadistill.com/report/3887"),
 ("mydailybyte：週報（第三季 2026-11-06 公布）", "https://www.mydailybyte.com/littlebit/plus/weekly-report-042700-2026-10-19"),
 ("市場層與上游成品：skhynix／samsung／micron／disco／markets/equipment 的 dist（本 repo）", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 120

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("韓美半導體 Hanmi Semiconductor（042700.KS）收支與估值模型 v0.2.2（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("v0.1（2026-10-25）：首版。樣板 advantest/（設備股，統一循環規則）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("模型命題：韓美的盈餘由 HBM 的 TC 鍵合需求決定——需求來自記憶體廠下一年的鍵合產能增量（HBM 位元 ÷ 晶粒容量 × 層數加權），韓美的份額集中在 SK 海力士與美光。股價取決於 HBM 位元成長在 2028 之後放緩、晶粒容量提高與混合鍵合導入時，TC 鍵合機的新機需求與韓美在 SK 海力士的份額能否守住。", False), ("", False),
 ("會計年度＝日曆年。金額以億韓元（₩億）為主；EPS 為韓元。FY2026 為錨點：上半年實績（509＋2,511）＋下半年假設（h2f、m_h2）；第三季預定 2026-11 公布，建置日未納入。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("HBM 出貨位元（markets/memory，情境別）÷ 晶粒容量 → 核心晶粒數 × 層數加權 ×（1 − 混合鍵合占比）→ TC 鍵合產能需求；新機需求＝下一年產能增量＋汰換", False),
 ("× 各記憶體廠占產能增量（t＋1 年 HBM 市占，讀三家記憶體模型）× 韓美在該廠份額 → 韓美可得需求 × 校準係數 k（FY2025–26 兩年營收）→ HBM TC 鍵合機營收；其他業務＝前一年 × 成長率", False),
 ("毛利率：FY2026 由錨點營業利益反推，之後沿用（C 扣價格競爭）；營業費用＝固定＋變動 × 營收；不寫死利益率", False),
 ("情境 A／B／C：HBM 位元讀 markets/memory（2028 ×1.0／0.8／0.6，2029 ×1.0／0.85／0.6）；混合鍵合為行業技術變數、與 Disco 同字母（v0.2.1：CY2028 起 A 取高值、B 基準、C 低值）；公司特有變數（份額、其他業務成長）CY2028 起分情境、A 最有利；C 另加韓華、ASMPT 搶份額與毛利率下調。Val：十年 P/E 分位，循環位置依設備股共用規則（見頂年 " + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）；目標價＝三情境 0.3／0.5／0.2 加權", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。DART 原文未開啟，實績取自新聞稿報導與財務資料頁，原文待人工複核。", False),
 ("已知限制：(1) 公司不揭露產品別與 2025 客戶別營收，TC 鍵合機占比（60%）與中國比重（17%）取自 CGSI（賣方）；(2) 韓美在 SK 海力士、美光的份額為 Assumed，模型對它們最敏感；(3) 以 HBM 營收市占代替各廠產能增量占比（Analogy）；(4) 晶粒容量、層數、混合鍵合比例沿用 Disco 模型的 Assumed；(5) 鍵合機假設比產能早一年購入，實際交期可能更短，會讓模型的轉折早於實際；(6) 韓美自家混合鍵合機（2027 量產工廠）的營收不建模；(7) 十年 P/E 2016–2017 只有一個來源，companiesmarketcap 2021 年值與他源差異大。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=COL, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('BE Semiconductor（BESI）', 'BESI.AS', 25.82, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('ASMPT', '0522.HK', 30.26, '2026-09-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Kulicke & Soffa', 'KLIC', 12.61, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Disco', '6146.T', 26.01, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2027: "2/12", 2028: "10/12"}, fwd_note='約 2027-10 時的未來 12 月 ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='₩',
    unit_rev='₩億',
    unit_ni='₩億',
    fmt='#,##0',
    fmt_rev='#,##0',
    band_fmt='#,##0',
    roe=('FY2023–FY2025', [0.5554, 0.2743, 0.3476], 'stockanalysis 比率頁（2026-10-25 讀取）：FY2023 55.54%、FY2024 27.43%、FY2025 34.76%', '單一來源；近 12 月 33.24%（同站）'),
    nim=None,
    cons_tp=(271429, 'stockanalysis：7 位分析師平均 ₩271,429（中位數 230,000；區間 120,000–500,000），頁面價格日 2026-09-21', '第二來源：未找到；FnGuide／Naver 共識頁無法開啟'),
    cons_feps=(5000, '共識前瞻 EPS＝stockanalysis 摘要「EPS Next Year」FY2027 約 ₩5,000（只顯示到千位）；本模型前瞻 EPS 以 FY2028 為主，約晚一年，EPS 因素偏大', 'alphadistill FY2027 4,650（In 頁 cons_eps27，非券商共識）'),
    peer_note='後段封裝設備同業：BESI（混合鍵合）、ASMPT（TC 鍵合）、Kulicke & Soffa、Disco；韓華 Semitech 未上市，無法列入',
    dcf_note='',
    pe10='B16:B25',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 2 節來源1 companiesmarketcap）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均，FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：BESI、ASMPT、Kulicke & Soffa、Disco 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：十年平均淨利率未找到，以 FY2029 × 0.8 代替（Assumed）；Kulicke & Soffa 前瞻倍數偏低、ASMPT 價格日期為 9 月，中位數受兩者拉扯；韓國市場折價未調整。',
    readme=['v0.2（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。十年平均淨利率未找到，收斂目標以 FY2029 淨利率 × 0.8 代替（Assumed）。'],
    nim_missing='十年淨利資料未找到（韓文券商報告為 PDF 片段，未能逐年核對）')

# ---------------- 估值 v2（2026-10-25，Andy 11:18／11:20／11:28／11:53）：PL 延伸到 FY2032（收斂法）＋ Val「0. 目標價」（DCF 與 Comparable 加權） ----------------
# 為不位移既有儲存格（FY2026–FY2029 既有 EPS 逐格不變），新內容放在各頁既有內容之後；HTML 估值章第一個區塊讀「0.3 結論」
def valuation_v2(CFG):
    cols = CFG["cols"]; MF = CFG["fmt"]; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
    FADEFILL = PatternFill("solid", fgColor="E7E6E6")
    # --- In：目標價與收斂法參數（黃底；Assumed，Andy 可改） ---
    r = IN.max_row + 2
    IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（v2 新增；黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
    def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
        nonlocal r
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        c = IN.cell(r, 5, base)
        if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
        if fmt: c.number_format = fmt
        if yellow: c.fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    inrow("tp_r", "折現率（股東權益要求報酬；美、台、日、韓同一值）", "%", 0.10, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%）", CFG.get("r_note", "—"), "0.0%", True)
    inrow("tp_g", "長期成長率（永續成長；收斂法 FY2032 營收年增的基準）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_roe", f"ROE（近三年平均，{CFG['roe'][0]}）", "%", "=AVERAGE(" + ",".join(str(x) for x in CFG["roe"][1]) + ")", "Derived", CFG["roe"][2], CFG["roe"][3], "0.0%")
    inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
    if CFG["nim"]:
        inrow("nim10", f"十年平均淨利率（{CFG['nim'][0]}，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(" + ",".join(f"{n}/{d}" for d, n in CFG["nim"][1]) + ")", "Derived", CFG["nim"][2], CFG["nim"][3], "0.0%")
    else:
        IN.cell(r, 1, "nim10"); IN.cell(r, 2, "十年平均淨利率"); IN.cell(r, 3, "%"); IN.cell(r, 5, "未找到"); IN.cell(r, 7, "Assumed")
        IN.cell(r, 8, CFG["nim_missing"]); IN.cell(r, 9, "依工作單：找不到十年平均時用最後一年（FY2029）淨利率 × 0.8；PL 延伸區塊逐情境計算"); r += 1
    inrow("cons_tp", "市場共識目標價（賣方平均）", CFG["cur"], CFG["cons_tp"][0], "Interested-party", CFG["cons_tp"][1], CFG["cons_tp"][2], MF)
    inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", CFG["cur"], CFG["cons_feps"][0], "Interested-party", CFG["cons_feps"][1], CFG["cons_feps"][2], MF)

    # --- PL：延伸到 FY2032（收斂法） ---
    p0 = PL.max_row + 3
    PL.cell(p0, 1, "延伸到 FY2032（v2 新增）：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
    PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均" + ("" if CFG["nim"] else "；十年平均未找到，以 FY2029 × 0.8 代替（Assumed）") + "）；股數維持 FY2029（EPS 與淨利同比例）")
    ext_cols = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
    hdr(PL, p0 + 2, ["情境（延伸）", "單位"] + [str(PL[f"{cols[y]}3"].value) for y in range(2025, 2030)] + ["FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    EXT = {}; pr = p0 + 3
    gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
    for s in SC:
        PL.cell(pr, 1, f"情境 {s}（{CFG['scn'][s]}）").font = BOLD; pr += 1
        R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
        for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
            PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": CFG["unit_rev"], "ni": CFG["unit_ni"], "eps": CFG["cur"]}.get(k, "%"))
        for y in range(2025, 2033):
            c = ext_cols[y]; pc = ext_cols.get(y - 1)
            if y <= 2029:
                PL[f"{c}{R['rev']}"] = f"={cols[y]}{SUMR[(s, 'rev')]}"
                PL[f"{c}{R['ni']}"] = f"={cols[y]}{SUMR[(s, 'ni')]}"
                PL[f"{c}{R['eps']}"] = f"={cols[y]}{SUMR[(s, 'eps')]}"
                PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
                if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            else:
                k = y - 2029
                mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+{'nim10' if CFG['nim'] else '0.8*$G' + str(R['nim'])})/2", "C": "nim10" if CFG["nim"] else f"0.8*$G{R['nim']}"}[s]
                PL[f"{c}{R['g']}"] = f"=$G{R['g']}+((tp_g{gadd[s]})-$G{R['g']})*{k}/3"
                PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
                PL[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({mT}-$G{R['nim']})*{k}/3"
                PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
                PL[f"{c}{R['eps']}"] = f"=$G{R['eps']}*{c}{R['ni']}/$G{R['ni']}"
                for kk in R: PL[f"{c}{R[kk]}"].fill = FADEFILL
            for kk, fm in (("rev", CFG["fmt_rev"]), ("g", "0.0%"), ("nim", "0.0%"), ("ni", CFG["fmt_rev"]), ("eps", MF)):
                if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
        PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
        EXT[s] = R; pr += 6
    for c in "IJ": PL.column_dimensions[c].width = 14
    PL.column_dimensions["K"].width = 40

    # --- Val：舊法改名（不再叫目標價） ---
    REN = (("0. 目標價與上檔空間", "舊法：規則值情境加權（參考）"), ("目標價（12 個月，情境加權）", "規則值情境加權價（12 個月；參考）"),
           ("上檔空間（目標價 ÷ 現價 − 1）", "相對現價（規則值情境加權價 ÷ 現價 − 1）"), ("目標價方法", "規則值方法"))
    for row in V.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            for a, b in REN:
                if v.startswith(a): row[0].value = b + v[len(a):]; break
    # --- Val：0. 目標價（列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期） ---
    t0 = V.max_row + 3
    T = {}
    def tset(rr, key, label, fs, fmt, note=""):
        V.cell(rr, 1, label); T[key] = rr
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
        if note: V[f"E{rr}"] = note
    V.cell(t0, 1, "0. 目標價").font = BOLD; T["top"] = t0
    V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["h01"] = t0 + 1
    hdr(V, t0 + 2, ["項目", f"A（{CFG['scn']['A']}）", f"B（{CFG['scn']['B']}）", f"C（{CFG['scn']['C']}）", "說明"]); T["hd1"] = t0 + 2
    V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"; T["eps_h"] = t0 + 3
    EPSR = {}
    for i, y in enumerate(YRS6):
        tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{ext_cols[y]}{EXT[s]['eps']}" for s in SC], MF); EPSR[y] = t0 + 4 + i
    rr = t0 + 10
    V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；設備股資本支出低（輕資產），此式偏誤小"; T["cf_h"] = rr
    tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
    RV = rr + 1; CFR = {}
    for i, y in enumerate(YRS6):
        tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], MF); CFR[y] = rr + 2 + i
    c1, c6 = CFR[2027], CFR[2032]
    tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], MF,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027" + CFG["dcf_note"])
    tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], MF)
    tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], MF, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
    q = rr + 12
    V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; T["h02"] = q
    hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["phd"] = q + 1
    for i, (nm, tk, pe, dt, src) in enumerate(CFG["peers"], q + 2):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    pl_ = q + 1 + len(CFG["peers"]); T["p1"] = q + 2; T["pl"] = pl_
    q = pl_ + 1
    tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{T['p1']}:C{pl_})", None, None], "0.0", CFG["peer_note"])
    fw = CFG["fwd"]
    def feps(c): return "(" + "+".join(f"{w}*{c}{EPSR[y]}" for y, w in fw.items()) + ")"
    tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("B") + "+tp_wB*" + feps("C") + "+tp_wC*" + feps("D"), None, None], MF, CFG["fwd_note"])
    tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 上列 EPS")
    q += 4
    V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
    hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
    q += 2
    tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
    tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2")
    tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, CFG["cons_tp"][1])
    tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", CFG["px_note"])
    tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
    tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；" + CFG["cons_feps"][1])
    tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
    V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
    tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"{CFG["band_fmt"]}")&"～"&TEXT(B{T["dcf12"]},"{CFG["band_fmt"]}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    PE = CFG["pe10"]
    tset(q + 8, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%", CFG["pe10_note"])
    V.cell(q + 9, 1, "目標價方法"); V[f"B{q + 9}"] = CFG["method"]; T["meth"] = q + 9
    for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
    if CFG.get("a3"): V["A3"] = CFG["a3"].format(row=t0)

    # --- Sens：目標價參數 ---
    s0 = SE.max_row + 3
    SE.cell(s0, 1, "目標價參數（v2 新增；只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
    hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_formula(c, r_="tp_r", g="tp_g"):
        cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
        return f"({pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.4", "Comparable 0.6", dict(w="0.4")), ("DCF 權重 0.8", "Comparable 0.2", dict(w="0.8"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf")
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        rr_ = kw.get("r_", "tp_r")
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{rr_})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = MF
        SE[f"F{i}"].number_format = "0.0%"
    T["sens0"] = s0; T["sens_last"] = s0 + 1 + len(cases)

    # --- README 頁 ---
    rr = RD.max_row + 1
    for t in CFG["readme"]:
        RD.cell(rr, 1, t); rr += 1
    return T, EXT

V2T, V2EXT = valuation_v2(CFG_V2)
RD.cell(RD.max_row + 1, 1, "v0.2.1（2026-10-25）情境一致性修正：(1) CY2027 購機的「下一年（CY2028）產能」三情境用共同前提＝情境 A 的 HBM 位元（hsA_28）× 混合鍵合基準，與 markets/equipment 2027 三情境相同、記憶體供給三情境相同一致；(2) 韓美在 SK 海力士／美光的份額與其他業務成長，CY2026–27 三情境共用情境 A 值，CY2028 起分情境（A 最有利）；(3) 混合鍵合為行業技術變數，與 Disco 同字母：CY2028 起 A＝導入快（Disco 高值）、B＝基準、C＝導入慢（低值）。A 情境 TC 鍵合需求被混合鍵合侵蝕最多，FY2028 EPS A＜B 屬預期結果。FY2027 三情境相同。")

# ---------------- v0.2.2（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；新增列放在各頁現有最後一列之後，既有儲存格不變 ----------------
_U = f"B{V2T['upw']}"
_G = V.max_row + 2
V.cell(_G, 1, "上檔空間級距")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）"
V.cell(_G + 1, 1, "投資建議")
V[f"B{_G + 1}"] = f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))'
V[f"B{_G + 1}"].number_format = "@"
V[f"E{_G + 1}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空"
OBS = ["目標價 ₩48,379（DCF 27,780、Comparable 68,979 各半），比現價 188,500（2026-10-23）低 74.3%，級距 小於 −50%；Comparable 單獨為 −63.4%。主因 EPS 與倍數都有：本模型 12 個月後前瞻 EPS 2,662 比共識前瞻 EPS 約 5,000 低 46.8%；現價隱含 70.8 倍前瞻 EPS，同業（BESI、ASMPT、Kulicke & Soffa、Disco）中位數 25.9 倍、共識目標價 271,429 隱含 54.3 倍（比同業高 109%）。DCF 低不是 ROE（39.2%）造成，而是 FY2028 起 TC 鍵合需求回落（FY2027 EPS 3,098 → A 情境 FY2028 2,363）。",
       "混合鍵合替代是最大變數：不納入混合鍵合時 FY2028 EPS 3,782（基準 2,363，+60.0%）、FY2027 +9.1%；晶粒容量取低值 FY2027 +15.1%，淨利 ÷ 營業利益 0.72 則 −15.4%。",
       "三情境的現價隱含 P/E 都高於十年最大值（分位 100%）；共識前瞻 EPS 只顯示到千位（約 5,000），差距待第三季財報（預定 2026-11-06 前後）的 TC 鍵合出貨與 2027 年訂單檢驗。"]   # 重點觀察（與 hanmi/README.md 相同文字）
_R = RD.max_row + 1
RD.cell(_R, 1, "v0.2.2（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。")
RD.cell(_R + 1, 1, f'="投資建議："&Val!B{_G + 1}&"（依上檔空間固定規則）"')
for _i, _t in enumerate(OBS, 1): RD.cell(_R + 1 + _i, 1, f"重點觀察 {_i}：{_t}")
if os.environ.get("G1_ROWS"): print("G1ROWS", "hanmi", "grade", _G, "rec", _G + 1, "readme_obs", _R + 2, _R + 4)

if os.environ.get("V2_ROWS"):
    print("V2", V2T, V2EXT)


# ---------------- 分部原則第 3 條（2026-10-28）：核對列與宣告（追加於工作表最後，不位移既有列，數值不變） ----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(_os2.path.dirname(_os2.path.dirname(_os2.path.abspath(__file__))), "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[15, 16, 17], total_row=4, cols="B", rnd=3, note="分部原則第 3 條（2026-10-28 追加）：客戶別營收（FY2024 事業報告）；分部宣告見 segments.json")
RD.cell(RD.max_row + 1, 1, "v0.2.3（2026-10-28）：分部原則重建（數值不變）。Q 頁最後追加核對列「檢查：分部合計 − 合併營收」；官方分部與子列宣告見 segments.json。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.2.2", "v0.2.3")

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "hanmi")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.4", "2026-10-31", split=True, split_skip=[])  # v0.2.4：In 頁版面；v0.2.5（2026-10-31）：第 3 點情境專屬列拆分（數值不變）
wb.save(sys.argv[1])
print("PB", {s: PB[s] for s in SC}); print("SUMR", SUMR); print("VAL", VAL, "fwd rows to", fr, "SENS_LAST", SENS_LAST, "Recon last", rr_)
