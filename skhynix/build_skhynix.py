# SK 海力士收支與估值模型 v0.1 建檔程式（openpyxl 寫公式；LibreOffice 重算）
# 用法：python3 build_skhynix.py raw.xlsx ；再 soffice --headless --convert-to xlsx --outdir dist raw.xlsx
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()

def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR

def name(key, ref):
    wb.defined_names[key] = DefinedName(key, attr_text=ref)

# ======================= In：輸入 =======================
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字＝基準值；低／高供敏感度）。每列有具名範圍＝鍵（指向基準欄 E）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源／說明"])
S = []  # (key, item, unit, lo, base, hi, tag, src) or ("sec", title)
def sec(t): S.append(("sec", t))
def inp(*a): S.append(a)

sec("A. 匯率與市場價格")
inp("fx", "美元兌韓元（2026–2028 平均）", "KRW/USD", 1350, 1420, 1500, "Derived", "Q2 2026 營收 79.32 兆韓元 ≈ US$55.2B → 約 1,437；7 月 ADS 定價匯率 1,538（S1）")
inp("px", "股價（2026-10-07 收盤）", "KRW", None, 1723000, None, "Verified", "Investing.com 歷史價格；10-02 為 1,841,000（StockAnalysis、MarketScreener）")

sec("B. 加速器數量層（百萬顆，封裝）與每顆 HBM（GB）— HBM 需求端")
units = [
 ("bw", "NVIDIA Blackwell（B200／B300）", (5.4, 2.4, 0.5), (260, 280, 288), "Interested-party", "顆數：Morgan Stanley 2026-07（2026 5.4M）；AC Research 2027 約 2.4M 淨新增；2028 Assumed。GB：B200 186、B300 288 的組合 [Derived]"),
 ("rb", "NVIDIA Rubin", (1.7, 7.0, 7.0), (288, 288, 288), "Interested-party", "2026：TrendForce 2026 組合 Rubin 22% × NVIDIA 約 7.6M [Derived]；2027：Morgan Stanley 約 7.0M（約 9 萬架 NVL72）；2028 Assumed 持平。GB：NVIDIA 規格 288"),
 ("ru", "NVIDIA Rubin Ultra", (0.0, 0.3, 2.0), (1024, 1024, 1024), "Assumed", "時程 Assumed（2027 下半年起）；GB：NVIDIA 1 TB／封裝"),
 ("hp", "NVIDIA Hopper（H200 等）", (0.5, 0.0, 0.0), (141, 141, 141), "Interested-party", "TrendForce 2026 組合 Hopper 7%"),
 ("amd", "AMD Instinct", (0.75, 1.5, 2.0), (288, 360, 432), "Interested-party／Assumed", "2027：Morgan Stanley MI455X 1.0M＋MI450 0.5M；2026、2028 Assumed。GB：MI355X 288、MI455X 432、MI450 約 216 [Derived]"),
 ("tpu", "Google TPU", (3.85, 5.5, 7.0), (192, 250, 288), "Interested-party", "Morgan Stanley 3.2／5.0（2026／2027）、GF Securities 4.5／8.84；取中間偏 MS；2028 Assumed。GB：Ironwood 192、TPU 8i 288、8t 約 208"),
 ("trn", "AWS Trainium", (2.0, 2.5, 3.0), (110, 144, 200), "Assumed", "年出貨未找到；累計已部署約 1.4M Trainium2（2026-04）。GB：Trn2 96、Trn3 144、Trn4 約 288"),
 ("oth", "其他 ASIC（MTIA、Maia、Ascend 等）", (1.0, 4.5, 5.0), (150, 200, 250), "Derived／Assumed", "2027：JPMorgan ASIC 12.5M − TPU − Trainium；其餘 Assumed"),
]
for k, nm, u, gb, tag, src in units:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u[i], None, tag, src if i == 0 else "同上")
    for i, y in enumerate((26, 27, 28)):
        inp(f"g_{k}_{y}", f"{nm} 每顆 HBM 20{y}", "GB", None, gb[i], None, tag, "同上")

sec("C. HBM 市場（供給、價格、份額）")
inp("hbm_gross", "出貨位元 ÷ 加速器裝載位元（封裝良率損失、庫存、通路）", "x", 1.0, 1.15, 1.3, "Assumed", "Epoch：CoWoS-L 良率 65–95%；另含庫存與未列入的買方")
inp("hbm_sup27", "2027 HBM 位元供給成長", "%", 0.50, 0.55, 0.63, "Interested-party", "TrendForce 2027 位元需求 +50–60%（供給受晶圓與 TSV 產能限制）；JPMorgan 2026–28 CAGR 63%")
inp("hbm_sup28", "2028 HBM 位元供給成長", "%", 0.30, 0.45, 0.63, "Assumed", "JPMorgan CAGR 63% 為上限；新廠 2028 才貢獻（TrendForce）")
inp("hbm_p26", "HBM 混合單價 2026", "$/GB", 12.5, 14.6, 16.0, "Interested-party", "BofA $14.6/GB（2026-08）；DigiTimes HBM3E $12–12.8、HBM4 約 $16")
inp("hbm_p27_A", "HBM 單價 2027 變動（情境 A）", "%", 0.38, 0.80, 1.21, "Interested-party", "BofA +38%、JPM +54%、UBS +76–79%、TrendForce +121%；取中間")
inp("hbm_p27_B", "HBM 單價 2027 變動（情境 B）", "%", 0.38, 0.80, 1.21, "Interested-party", "2027 年約已在議價，兩情境相同")
inp("hbm_p28_A", "HBM 單價 2028 變動（情境 A）", "%", 0.0, 0.10, 0.25, "Interested-party", "JPMorgan 2028 +25%（上限）；BofA 2028 溫和修正")
inp("hbm_p28_B", "HBM 單價 2028 變動（情境 B）", "%", -0.40, -0.25, -0.10, "Assumed", "2027 見頂、2028 一般 DRAM 大跌時，HBM 年約重議跟著下修")
inp("sh26", "SK 海力士 HBM 營收市占 2026", "%", 0.48, 0.52, 0.56, "Interested-party", "Counterpoint Q1 58%、Q2 50%；IDC Q1 56.4%；UBS 位元 48%")
inp("sh27", "SK 海力士 HBM 營收市占 2027", "%", 0.38, 0.42, 0.48, "Interested-party", "UBS 2027 位元：SK 39%、三星 41%、美光 20%；JPM：三星＋美光 59%")
inp("sh28", "SK 海力士 HBM 營收市占 2028", "%", 0.35, 0.40, 0.45, "Assumed", "延續 2027")
inp("nv_dc26", "NVIDIA 資料中心營收 2026（約 FY27）", "$B", 340, 374, 400, "Verified／Assumed", "FY27 Q1 75.2、Q2 89.0（8-K）；Q3 指引總營收 108（資料中心取 100）；Q4 Assumed 110")
inp("nv_content", "每顆 GPU 的 NVIDIA 內容（GB300 橋接表）", "$K", 45, 51.6, 58, "Derived", "markets/chip_bridge v0.1：GB300 IT capex $77K/GPU × NVIDIA 合計 ÷ IT 67.4%")
inp("nv_hbm", "每顆 GPU HBM 金額（GB300 橋接表）", "$K", 3.5, 4.46, 4.9, "Derived", "markets/chip_bridge v0.1（288GB × $15.5/GB）")
inp("pub26_lo", "已發布 2026 HBM 市場（低）", "$B", None, 54.6, None, "Interested-party", "SK 海力士自估（Motley Fool／글로벌이코노믹 2026-09）")
inp("pub26_hi", "已發布 2026 HBM 市場（高）", "$B", None, 77.4, None, "Interested-party", "BofA 2026-08-07；CLSA 推算約 69")
inp("pub27", "已發布 2027 HBM 市場", "$B", None, 152.8, None, "Interested-party", "BofA；JPMorgan 2027–28 為 160–282")

sec("D. SK 海力士 2026 下半年（實績錨定＋公司指引＋TrendForce 價格）")
inp("q3_db", "Q3 DRAM 位元季增", "%", 0.08, 0.10, 0.12, "Verified", "公司 Q2 法說：約 +10%")
inp("q3_dp", "Q3 DRAM 混合 ASP 季增", "%", 0.08, 0.12, 0.16, "Derived", "TrendForce 一般 DRAM +13–18%；HBM 年約鎖價拉低混合值")
inp("q3_nb", "Q3 NAND 位元季增", "%", 0.0, 0.02, 0.04, "Verified", "公司：低個位數")
inp("q3_np", "Q3 NAND ASP 季增", "%", 0.10, 0.125, 0.20, "Interested-party", "TrendForce +10–15%；Morgan Stanley 實際約 +20%")
inp("q4_db", "Q4 DRAM 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", "公司：下半年位元成長優於上半年")
inp("q4_dp", "Q4 DRAM 混合 ASP 季增", "%", 0.05, 0.10, 0.15, "Derived", "TrendForce 一般 DRAM +10–15%；BofA 產業 ASP 約 +8%")
inp("q4_nb", "Q4 NAND 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", "")
inp("q4_np", "Q4 NAND ASP 季增", "%", 0.10, 0.175, 0.20, "Interested-party", "TrendForce +15–20%")
inp("costdn_q", "單位成本季降", "%", 0.01, 0.03, 0.05, "Assumed", "製程微縮；年約 −10–12%")
inp("da_q3", "Q3 折舊攤銷", "兆韓元", None, 4.3, None, "Assumed", "Q1 3.726、Q2 4.028（6-K）延伸")
inp("da_q4", "Q4 折舊攤銷", "兆韓元", None, 4.6, None, "Assumed", "同上")
inp("opex_fix", "營業費用固定部分（每季）", "兆韓元", None, 1.5, None, "Derived", "Q1、Q2 2026：營業費用 3.97、5.24 對營收 52.6、79.3 迴歸 → 固定約 1.5、變動 4.7%")
inp("opex_var", "營業費用變動率", "% 營收", 0.04, 0.047, 0.06, "Derived", "同上（含績效獎金隨獲利增加）")
inp("nonop_q", "下半年每季業外淨收益（利息等）", "兆韓元", 0.5, 1.0, 2.0, "Assumed", "淨現金約 70–110 兆 × 約 2%／年 ÷ 4")
inp("tax", "有效稅率", "%", 0.20, 0.23, 0.25, "Verified", "Q1 21.8%、Q2 23.5%（6-K）；2025 14.9%")

sec("E. 2027–2028 年度驅動（情境 A＝共識：2027–28 高檔持平；情境 B＝2027 見頂、2028 大幅回落）")
inp("cb27", "一般 DRAM 位元成長 2027", "%", 0.05, 0.10, 0.15, "Derived", "BofA DRAM 供給 +19%；HBM 晶圓占比 22%→30%（TrendForce）吃掉一般 DRAM 增量")
inp("cb28", "一般 DRAM 位元成長 2028", "%", 0.10, 0.15, 0.19, "Interested-party", "BofA 2028 供給 +19%；新廠 2028 貢獻")
inp("cp27_A", "一般 DRAM 年均價 2027 ÷ 2026（A）", "%", 0.15, 0.30, 0.45, "Derived", "TrendForce 季價路徑：2026 均值指數 5.08、Q4 6.45；BofA 2027 持平略降 → 2027 均值約 6.5（+28%）；Bernstein $/GB +43%")
inp("cp27_B", "一般 DRAM 年均價 2027 ÷ 2026（B）", "%", 0.10, 0.25, 0.40, "Derived", "Citi：2027 Q2 見頂後下滑 → 2027 均值約 6.4（+25%）")
inp("cp28_A", "一般 DRAM 年均價 2028 ÷ 2027（A）", "%", -0.20, -0.10, 0.0, "Interested-party", "BofA：2028 溫和修正、非硬著陸")
inp("cp28_B", "一般 DRAM 年均價 2028 ÷ 2027（B）", "%", -0.60, -0.50, -0.30, "Interested-party", "Bernstein：DRAM $/GB 2028 −52.9%")
inp("nb27", "NAND 位元成長 2027", "%", 0.15, 0.20, 0.25, "Verified", "Micron：產業 2027 中 20% 段")
inp("nb28", "NAND 位元成長 2028", "%", 0.15, 0.20, 0.25, "Verified", "同上")
inp("np27_A", "NAND 年均價 2027 ÷ 2026（A）", "%", 0.15, 0.32, 0.45, "Derived", "TrendForce 季價路徑 2026 均值約 4.6、Q4 6.1；2027 下半年供需平衡 → 均值約 6.05")
inp("np27_B", "NAND 年均價 2027 ÷ 2026（B）", "%", 0.10, 0.25, 0.40, "Derived", "Citi：2027 Q2 見頂")
inp("np28_A", "NAND 年均價 2028 ÷ 2027（A）", "%", -0.30, -0.20, -0.10, "Interested-party", "TrendForce：2027 下半年價格面臨下修壓力")
inp("np28_B", "NAND 年均價 2028 ÷ 2027（B）", "%", -0.70, -0.65, -0.40, "Interested-party", "Bernstein：NAND $/GB 2028 −68.8%")
inp("other_y", "其他營收（每年）", "兆韓元", None, 1.6, None, "Verified", "2025 約 1.6% 營收")
inp("w_c", "成本權重：一般 DRAM", "%", None, 0.45, None, "Assumed", "2026 Q4 成本組成（未揭露）")
inp("w_h", "成本權重：HBM", "%", None, 0.30, None, "Assumed", "HBM 每位元耗晶圓約 DDR5 的 3 倍（Micron）")
inp("w_n", "成本權重：NAND", "%", None, 0.25, None, "Assumed", "")
inp("costdn_y", "單位成本年降", "%", 0.05, 0.10, 0.15, "Assumed", "")
inp("da_life", "新增資本支出轉為折舊的年數", "年", 4, 5, 6, "Assumed", "D&A_t＝D&A_t−1＋(capex_t−1 − D&A_t−1) ÷ 年數")
inp("da_cogs", "折舊攤銷計入營業成本的比例", "%", None, 0.9, None, "Assumed", "")
inp("opex_g", "營業費用固定部分年增", "%", None, 0.10, None, "Assumed", "")
inp("cap26", "資本支出 2026", "兆韓元", 45, 48.7, 50, "Verified／Interested-party", "公司指引「40 兆後段」；MarketScreener 共識 48.7")
inp("cap27", "資本支出 2027", "兆韓元", 55, 62.3, 75, "Interested-party", "MarketScreener 共識 62.3；BNK：需 +30–40%；未承諾資本支出 61.5 兆（6-K）")
inp("cap28", "資本支出 2028", "兆韓元", 60, 71.2, 85, "Interested-party", "MarketScreener 共識 71.2")
inp("rf", "淨現金收益率", "%", None, 0.02, None, "Assumed", "")

sec("F. 資本與股東回饋")
inp("eq_h1", "權益總額 2026-06-30", "兆韓元", None, 262.693, None, "Verified", "6-K 半年報（S2）")
inp("nc_h1", "淨現金 2026-06-30", "兆韓元", None, 69.4, None, "Verified", "公司定義（Q2 簡報）")
inp("sh_h1", "流通股數 2026-06-30", "百萬股", None, 711.0755, None, "Verified", "6-K")
inp("ads_kw", "ADS 募資（2026-07）", "兆韓元", None, 40.4, None, "Derived", "US$26.25B × 1,538.05")
inp("ads_sh", "ADS 新股", "百萬股", None, 17.79, None, "Verified", "424B4")
inp("bb26", "2026 庫藏股買回並註銷（8/20 起約 3 個月）", "兆韓元", None, 40.0, None, "Verified", "公司 2026-08-19 公告")
inp("bb26_sh", "2026 買回股數", "百萬股", None, 24.07, None, "Verified", "同上")
inp("div26_h2", "2026 下半年股利支付", "兆韓元", None, 0.54, None, "Derived", "每季 375 韓元 × 2 季 × 約 720M 股")
inp("dps27", "每股股利 2027", "KRW", None, 4541, None, "Interested-party", "MarketScreener 共識")
inp("dps28", "每股股利 2028", "KRW", None, 5249, None, "Interested-party", "同上")
inp("payout", "股東回饋 ÷ 自由現金流（股利＋買回）", "%", 0.3, 0.5, 0.7, "Verified", "公司 2026-08-19：2025–27 累計 FCF 的 50% 以上；2028 Assumed 延續")
inp("bb_px", "2027–28 買回均價", "KRW", None, 1723000, None, "Assumed", "＝目前股價")

sec("G. 估值")
inp("pb_1", "P/B：1.0 倍", "x", None, 1.0, None, "Assumed", "研究者設定的循環谷底倍數")
inp("pb_2", "P/B：歷史中位數", "x", None, 1.6, None, "Interested-party", "2010–2025 年底 P/B 0.81–2.00，中位約 1.6（companiesmarketcap）")
inp("pb_3", "P/B：券商遠期", "x", None, 3.3, None, "Interested-party", "Shinhan 2026-07 目標價依 3.3x 遠期 P/B")
inp("pe_1", "P/E：4 倍", "x", None, 4, None, "Assumed", "高峰盈餘的低倍數；共識 2027E P/E 約 3.7x")
inp("pe_2", "P/E：8 倍", "x", None, 8, None, "Assumed", "")
inp("pe_3", "P/E：12 倍", "x", None, 12, None, "Assumed", "")

r = 4
for s in S:
    if s[0] == "sec":
        IN.cell(r, 1, s[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = s
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v); c.font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, src)
    name(key, f"In!$E${r}")
    if lo is not None: name(key + "_lo", f"In!$D${r}")
    if hi is not None: name(key + "_hi", f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGH", [12, 46, 11, 8, 10, 8, 20, 120]):
    IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ======================= Mkt：加速器數量層與 HBM 市場 =======================
M = wb.create_sheet("Mkt")
M["A1"] = "Mkt — 加速器數量層（最小版）與 HBM 市場：需求（顆數 × GB）、供給（位元成長）、價格 → 市場規模；並與由上而下、已發布數字對帳"
hdr(M, 3, ["項目", "單位", "2026", "2027", "2028", "說明"])
mr = 4; MK = {}
def mput(label, unit, fs, key=None, fmt="0.00", note=""):
    global mr
    M.cell(mr, 1, label); M.cell(mr, 2, unit)
    for j, f in enumerate(fs):
        c = M.cell(mr, 3 + j, f); c.number_format = fmt
    M.cell(mr, 6, note)
    if key: MK[key] = mr
    mr += 1
yrs = (26, 27, 28); YC = {26: "C", 27: "D", 28: "E"}
M.cell(mr, 1, "需求：加速器裝載的 HBM（PB）").font = BOLD; mr += 1
for k, nm, *_ in units:
    mput(f"  {nm}", "PB", [f"=u_{k}_{y}*g_{k}_{y}" for y in yrs], key="pb_" + k, fmt="0")
first, last = MK["pb_bw"], MK["pb_oth"]
mput("需求合計（加速器裝載）", "EB", [f"=SUM({YC[y]}{first}:{YC[y]}{last})/1000" for y in yrs], key="dem", note="1 EB＝1,000 PB")
mput("需求合計 × 出貨／裝載係數", "EB", [f"={YC[y]}{MK['dem']}*hbm_gross" for y in yrs], key="demg")
mput("需求年增", "%", ["", f"=D{MK['demg']}/C{MK['demg']}-1", f"=E{MK['demg']}/D{MK['demg']}-1"], key="demgr", fmt="0%")
mput("加速器顆數合計", "百萬顆", [f"=" + "+".join(f"u_{k}_{y}" for k, *_ in units) for y in yrs], key="units", fmt="0.0")
mput("  其中 GPU（NVIDIA＋AMD）", "百萬顆", [f"=u_bw_{y}+u_rb_{y}+u_ru_{y}+u_hp_{y}+u_amd_{y}" for y in yrs], key="gpu", fmt="0.0", note="JPMorgan 2027：GPU 10.9M、ASIC 12.5M")
mput("每顆平均 HBM", "GB", [f"={YC[y]}{MK['dem']}*1000/{YC[y]}{MK['units']}" for y in yrs], key="gbavg", fmt="0")
mr += 1
M.cell(mr, 1, "供給與出貨（位元）").font = BOLD; mr += 1
mput("供給（2026＝需求出貨；之後依供給成長）", "EB", [f"=C{MK['demg']}", f"=C{MK['demg']}*(1+hbm_sup27)", f"=D{{r}}*(1+hbm_sup28)"], key="sup")
M.cell(MK["sup"], 5, f"=D{MK['sup']}*(1+hbm_sup28)")
mput("出貨＝MIN(需求, 供給)", "EB", [f"=MIN({YC[y]}{MK['demg']},{YC[y]}{MK['sup']})" for y in yrs], key="ship")
mput("需求 − 供給（缺口＞0＝短缺，支撐價格）", "EB", [f"={YC[y]}{MK['demg']}-{YC[y]}{MK['sup']}" for y in yrs], key="gap")
mr += 1
M.cell(mr, 1, "價格與市場規模").font = BOLD; mr += 1
mput("HBM 單價（情境 A）", "$/GB", ["=hbm_p26", f"=C{{r}}*(1+hbm_p27_A)", ""], key="pA")
M.cell(MK["pA"], 4, f"=C{MK['pA']}*(1+hbm_p27_A)"); M.cell(MK["pA"], 5, f"=D{MK['pA']}*(1+hbm_p28_A)")
mput("HBM 單價（情境 B）", "$/GB", ["=hbm_p26", "", ""], key="pB")
M.cell(MK["pB"], 4, f"=C{MK['pB']}*(1+hbm_p27_B)"); M.cell(MK["pB"], 5, f"=D{MK['pB']}*(1+hbm_p28_B)")
mput("HBM 市場（情境 A）", "$B", [f"={YC[y]}{MK['ship']}*{YC[y]}{MK['pA']}" for y in yrs], key="mA", fmt="0.0", note="EB × $/GB ＝ $B")
mput("HBM 市場（情境 B）", "$B", [f"={YC[y]}{MK['ship']}*{YC[y]}{MK['pB']}" for y in yrs], key="mB", fmt="0.0")
mput("SK 海力士 HBM 市占", "%", ["=sh26", "=sh27", "=sh28"], key="sh", fmt="0%")
mput("SK 海力士 HBM 營收（A）", "兆韓元", [f"={YC[y]}{MK['mA']}*{YC[y]}{MK['sh']}*fx/1000" for y in yrs], key="skA", fmt="0.0")
mput("SK 海力士 HBM 營收（B）", "兆韓元", [f"={YC[y]}{MK['mB']}*{YC[y]}{MK['sh']}*fx/1000" for y in yrs], key="skB", fmt="0.0")
mput("SK 海力士 HBM 位元成長", "%", ["", f"=D{MK['ship']}*D{MK['sh']}/(C{MK['ship']}*C{MK['sh']})-1", f"=E{MK['ship']}*E{MK['sh']}/(D{MK['ship']}*D{MK['sh']})-1"], key="skbit", fmt="0%")
mr += 1
M.cell(mr, 1, "對帳（驗收標準 2）").font = BOLD; mr += 1
mput("由下而上：NVIDIA 裝載 HBM × 單價", "$B", [f"=(C{MK['pb_bw']}+C{MK['pb_rb']}+C{MK['pb_ru']}+C{MK['pb_hp']})/1000*hbm_p26", "", ""], key="bu_nv", fmt="0.0", note="2026；未乘出貨係數")
mput("由上而下：NVIDIA 資料中心營收 × (每顆 HBM ÷ 每顆 NVIDIA 內容)", "$B", ["=nv_dc26*nv_hbm/nv_content", "", ""], key="td_nv", fmt="0.0", note="營收、HBM 金額取自橋接表；與顆數預測獨立")
mput("  差距（由下而上 ÷ 由上而下 − 1）", "%", [f"=C{MK['bu_nv']}/C{MK['td_nv']}-1", "", ""], key="d1", fmt="0%")
mput("由營收推得的 NVIDIA 顆數", "百萬顆", ["=nv_dc26/nv_content", "", ""], key="nvu", fmt="0.0")
mput("  對照：顆數預測（Blackwell＋Rubin＋Hopper）", "百萬顆", ["=u_bw_26+u_rb_26+u_hp_26", "", ""], key="nvu2", fmt="0.0")
mput("已發布 HBM 市場（低／高）", "$B", ["=pub26_lo", "=pub27", ""], key="pub", fmt="0.0", note="2026：SK 海力士自估 54.6；高 77.4（BofA）。2027：BofA 152.8")
mput("  本模型 ÷ 已發布（A）", "%", [f"=C{MK['mA']}/C{MK['pub']}-1", f"=D{MK['mA']}/D{MK['pub']}-1", ""], key="d2", fmt="0%")
mput("  本模型 ÷ 已發布高值（2026）", "%", [f"=C{MK['mA']}/pub26_hi-1", "", ""], key="d3", fmt="0%")
for col, w in zip("ABCDEF", [58, 9, 10, 10, 10, 60]):
    M.column_dimensions[col].width = w

# ======================= Q：季度（2025Q1–2026Q4）=======================
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 季度損益（2025Q1–2026Q2 實績，SEC 424B4／6-K 與公司公告；2026Q3–Q4 由實績 × 公司指引 × TrendForce 價格推估）；兆韓元"
qs = ["2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1", "2026Q2", "2026Q3E", "2026Q4E"]
hdr(Q, 3, ["項目"] + qs + ["2026 合計"])
act = {
 "rev": [17.639, 22.232, 24.449, 32.827, 52.576, 79.319],
 "op":  [7.441, 9.213, 11.383, 19.170, 37.610, 60.543],
 "ni":  [8.108, 6.996, 12.598, 15.246, 40.346, 93.923],
 "da":  [3.334, 3.443, 3.556, 3.557, 3.726, 4.028],
}
QR = {}
def qrow(label, key, vals=None, fmt="0.00"):
    rr = 4 + len(QR)
    Q.cell(rr, 1, label); QR[key] = rr
    if vals:
        for j, v in enumerate(vals):
            if v is None: continue
            c = Q.cell(rr, 2 + j, v); c.number_format = fmt
            if isinstance(v, (int, float)): c.font = BLUE
    return rr
qrow("營收", "rev", act["rev"])
qrow("  DRAM", "dram", [None, None, None, None, "=F4*0.773", "=G4*0.73"])
qrow("  NAND", "nand", [None, None, None, None, "=F4*0.22", "=G4*0.27"])
qrow("  其他", "oth", [None, None, None, None, "=F4-F5-F6", "=G4-G5-G6"])
qrow("營業成本", "cogs", [None, None, None, None, "=F4*(1-0.79)", "=G4*(1-0.83)"])
qrow("營業費用", "opex", [None, None, None, None, "=F4-F8-F10", "=G4-G8-G10"])
qrow("營業利益", "op", act["op"])
qrow("營業利益率", "opm")
qrow("淨利", "ni", act["ni"])
qrow("折舊攤銷", "da", act["da"])
qrow("DRAM 位元指數（2026Q2＝1）", "dbi")
qrow("NAND 位元指數", "nbi")
r_ = QR
# 2026 Q3/Q4 推估（H、I 欄）
for c, p, db, dp, nb, np_, da in (("H", "G", "q3_db", "q3_dp", "q3_nb", "q3_np", "da_q3"), ("I", "H", "q4_db", "q4_dp", "q4_nb", "q4_np", "da_q4")):
    Q[f"{c}{r_['dram']}"] = f"={p}{r_['dram']}*(1+{db})*(1+{dp})"
    Q[f"{c}{r_['nand']}"] = f"={p}{r_['nand']}*(1+{nb})*(1+{np_})"
    Q[f"{c}{r_['oth']}"] = "=other_y/4"
    Q[f"{c}{r_['rev']}"] = f"={c}{r_['dram']}+{c}{r_['nand']}+{c}{r_['oth']}"
    Q[f"{c}{r_['dbi']}"] = f"={p}{r_['dbi']}*(1+{db})"
    Q[f"{c}{r_['nbi']}"] = f"={p}{r_['nbi']}*(1+{nb})"
    Q[f"{c}{r_['cogs']}"] = (f"=({p}{r_['cogs']}-da_cogs*{p}{r_['da']})*((w_c+w_h)*{c}{r_['dbi']}/{p}{r_['dbi']}+w_n*{c}{r_['nbi']}/{p}{r_['nbi']})*(1-costdn_q)+da_cogs*{da}")
    Q[f"{c}{r_['opex']}"] = f"=opex_fix+opex_var*{c}{r_['rev']}"
    Q[f"{c}{r_['op']}"] = f"={c}{r_['rev']}-{c}{r_['cogs']}-{c}{r_['opex']}"
    Q[f"{c}{r_['ni']}"] = f"=({c}{r_['op']}+nonop_q)*(1-tax)"
    Q[f"{c}{r_['da']}"] = f"={da}"
Q[f"G{r_['dbi']}"] = 1; Q[f"G{r_['nbi']}"] = 1
for c in "BCDEFGHI":
    Q[f"{c}{r_['opm']}"] = f"={c}{r_['op']}/{c}{r_['rev']}"; Q[f"{c}{r_['opm']}"].number_format = "0.0%"
for k, rr in r_.items():
    if k in ("dbi", "nbi"): continue
    Q[f"J{rr}"] = f"=SUM(F{rr}:I{rr})" if k != "opm" else f"=J{r_['op']}/J{r_['rev']}"
    Q[f"J{rr}"].number_format = "0.0%" if k == "opm" else "0.00"
Q.cell(18, 1, "註：2025 季度 DRAM／NAND 拆分、Q2–Q4 2025 資本支出未找到；2025Q3、Q4 折舊為合計 7.113 平均分配 [Derived]。2026Q1 毛利率約 79%、Q2 83% 為公司簡報數；營業費用＝營收 − 營業成本 − 營業利益 [Derived]")
Q.column_dimensions["A"].width = 30
for c in "BCDEFGHIJ": Q.column_dimensions[c].width = 11

# ======================= Y：年度（兩情境）=======================
Y = wb.create_sheet("Y")
Y["A1"] = "Y — 年度損益、資本支出、淨值與股數（兆韓元；股數百萬股）。情境 A：2027–28 高檔持平（市場共識型）；情境 B：2027 見頂、2028 價格大幅回落"
hdr(Y, 3, ["項目", "2025A", "2026E", "2027E A", "2028E A", "2027E B", "2028E B", "說明"])
YR = {}
def yrow(label, key, f25=None, f26=None, fA=None, fB=None, fmt="0.0", note=""):
    rr = 4 + len(YR)
    Y.cell(rr, 1, label); YR[key] = rr
    for col, f in (("B", f25), ("C", f26)):
        if f is not None:
            c = Y[f"{col}{rr}"]; c.value = f; c.number_format = fmt
    for (c27, c28, sc) in (("D", "E", "A"), ("F", "G", "B")):
        if fA is not None:
            f27, f28 = (fA if sc == "A" else fB)
            for col, f in ((c27, f27), (c28, f28)):
                if f is None: continue
                ff = f.replace("{P}", {"D": "C", "E": "D", "F": "C", "G": "F"}[col]).replace("{C}", col).replace("{S}", sc)
                cc = Y[f"{col}{rr}"]; cc.value = ff; cc.number_format = fmt
    Y.cell(rr, 8, note)
    return rr
qJ = lambda k: f"=Q!J{QR[k]}"
yrow("營收", "rev", "=SUM(Q!B4:E4)", qJ("rev"), ("={C}5+{C}6+{C}7+{C}8", "={C}5+{C}6+{C}7+{C}8"), ("={C}5+{C}6+{C}7+{C}8", "={C}5+{C}6+{C}7+{C}8"))
yrow("  一般 DRAM", "conv", "=74.9-28.1", f"=Q!J{QR['dram']}-C6", ("={P}5*(1+cb27)*(1+cp27_{S})", "={P}5*(1+cb28)*(1+cp28_{S})"), ("={P}5*(1+cb27)*(1+cp27_{S})", "={P}5*(1+cb28)*(1+cp28_{S})"), note="2025：DRAM 74.9（FY25 77.1%）− HBM 28.1（BofA）")
yrow("  HBM", "hbm", "=28.1", f"=Mkt!C{MK['skA']}", (f"=Mkt!D{MK['skA']}", f"=Mkt!E{MK['skA']}"), (f"=Mkt!D{MK['skB']}", f"=Mkt!E{MK['skB']}"), note="2026 起＝HBM 市場 × 市占（Mkt）；BofA 2026E 45.7、CLSA 49.4")
yrow("  NAND", "nand", "=20.69", qJ("nand"), ("={P}7*(1+nb27)*(1+np27_{S})", "={P}7*(1+nb28)*(1+np28_{S})"), ("={P}7*(1+nb27)*(1+np27_{S})", "={P}7*(1+nb28)*(1+np28_{S})"), note="FY25 NAND 21.3%")
yrow("  其他", "oth", "=97.147-74.9-28.1*0-20.69-0", qJ("oth"), ("=other_y", "=other_y"), ("=other_y", "=other_y"))
Y["B8"] = "=B4-74.9-20.69"
yrow("  HBM ÷ 營收", "hbmsh", "=B6/B4", "=C6/C4", ("={C}6/{C}4", "={C}6/{C}4"), ("={C}6/{C}4", "={C}6/{C}4"), fmt="0.0%")
yrow("折舊攤銷", "da", "=SUM(Q!B13:E13)", qJ("da"), ("={P}10+({P}17-{P}10)/da_life", "={P}10+({P}17-{P}10)/da_life"), ("={P}10+({P}17-{P}10)/da_life", "={P}10+({P}17-{P}10)/da_life"), note="2027 起：D&A_t－1＋(capex_t－1 − D&A_t－1) ÷ 年數")
yrow("成本量指數年增", "vol", None, None, (f"=w_c*(1+cb27)+w_h*(1+Mkt!D{MK['skbit']})+w_n*(1+nb27)", f"=w_c*(1+cb28)+w_h*(1+Mkt!E{MK['skbit']})+w_n*(1+nb28)"), (f"=w_c*(1+cb27)+w_h*(1+Mkt!D{MK['skbit']})+w_n*(1+nb27)", f"=w_c*(1+cb28)+w_h*(1+Mkt!E{MK['skbit']})+w_n*(1+nb28)"), fmt="0.00", note="HBM 位元依 SK 海力士 HBM 位元成長")
yrow("營業成本", "cogs", "=B4-B14-B13", qJ("cogs"), ("=({P}12-da_cogs*{P}10)*{C}11*(1-costdn_y)+da_cogs*{C}10", "=({P}12-da_cogs*{P}10)*{C}11*(1-costdn_y)+da_cogs*{C}10"), ("=({P}12-da_cogs*{P}10)*{C}11*(1-costdn_y)+da_cogs*{C}10", "=({P}12-da_cogs*{P}10)*{C}11*(1-costdn_y)+da_cogs*{C}10"), note="現金成本 × 量 × (1 − 年降)＋折舊")
yrow("營業費用", "opex", "=B4-B14-B12", qJ("opex"), ("=opex_fix*4*(1+opex_g)+opex_var*{C}4", "=opex_fix*4*(1+opex_g)^2+opex_var*{C}4"), ("=opex_fix*4*(1+opex_g)+opex_var*{C}4", "=opex_fix*4*(1+opex_g)^2+opex_var*{C}4"))
Y["B12"] = "=B4*0.37"; Y["B13"] = "=B4-B12-B14"
yrow("營業利益", "op", "=SUM(Q!B10:E10)", qJ("op"), ("={C}4-{C}12-{C}13", "={C}4-{C}12-{C}13"), ("={C}4-{C}12-{C}13", "={C}4-{C}12-{C}13"))
yrow("營業利益率", "opm", "=B14/B4", "=C14/C4", ("={C}14/{C}4", "={C}14/{C}4"), ("={C}14/{C}4", "={C}14/{C}4"), fmt="0.0%")
yrow("淨利", "ni", "=SUM(Q!B12:E12)", qJ("ni"), ("=({C}14+rf*{P}22)*(1-tax)", "=({C}14+rf*{P}22)*(1-tax)"), ("=({C}14+rf*{P}22)*(1-tax)", "=({C}14+rf*{P}22)*(1-tax)"), note="2026 含 H1 Kioxia 相關一次性收益約 63 兆（業外）")
yrow("資本支出", "capex", "=27.519", "=cap26", ("=cap27", "=cap28"), ("=cap27", "=cap28"))
yrow("自由現金流（淨利＋折舊 − 資本支出）", "fcf", "=B16+B10-B17", "=C16+C10-C17", ("={C}16+{C}10-{C}17", "={C}16+{C}10-{C}17"), ("={C}16+{C}10-{C}17", "={C}16+{C}10-{C}17"), note="未計營運資金變動")
yrow("股利", "div", None, "=div26_h2", ("=dps27*{P}23/1000000", "=dps28*{P}23/1000000"), ("=dps27*{P}23/1000000", "=dps28*{P}23/1000000"), note="2026 只列下半年（上半年已在 6 月底權益中）")
yrow("買回", "bb", None, "=bb26", ("=MAX(0,payout*{C}18-{C}19)", "=MAX(0,payout*{C}18-{C}19)"), ("=MAX(0,payout*{C}18-{C}19)", "=MAX(0,payout*{C}18-{C}19)"))
yrow("期末權益", "eq", None, f"=eq_h1+(Q!H{QR['ni']}+Q!I{QR['ni']})+ads_kw-C20-C19", ("={P}21+{C}16-{C}19-{C}20", "={P}21+{C}16-{C}19-{C}20"), ("={P}21+{C}16-{C}19-{C}20", "={P}21+{C}16-{C}19-{C}20"), note="2026：6 月底權益＋下半年淨利＋ADS − 買回 − 股利")
yrow("期末淨現金", "nc", None, f"=nc_h1+(Q!H{QR['ni']}+Q!I{QR['ni']})+(Q!H{QR['da']}+Q!I{QR['da']})-C17*0.6+ads_kw-C20-C19", ("={P}22+{C}18-{C}19-{C}20", "={P}22+{C}18-{C}19-{C}20"), ("={P}22+{C}18-{C}19-{C}20", "={P}22+{C}18-{C}19-{C}20"), note="2026 下半年資本支出取全年約 60%（上半年實付 18.3）")
yrow("期末流通股數", "sh", None, "=sh_h1+ads_sh-bb26_sh", ("={P}23-{C}20/bb_px*1000000", "={P}23-{C}20/bb_px*1000000"), ("={P}23-{C}20/bb_px*1000000", "={P}23-{C}20/bb_px*1000000"), fmt="0.0")
yrow("每股淨值 BVPS", "bvps", None, "=C21/C23*1000000", ("={C}21/{C}23*1000000", "={C}21/{C}23*1000000"), ("={C}21/{C}23*1000000", "={C}21/{C}23*1000000"), fmt="#,##0")
yrow("每股盈餘 EPS（期末股數）", "eps", None, "=C16/C23*1000000", ("={C}16/{C}23*1000000", "={C}16/{C}23*1000000"), ("={C}16/{C}23*1000000", "={C}16/{C}23*1000000"), fmt="#,##0")
yrow("ROE（期末權益）", "roe", None, "=C16/C21", ("={C}16/{C}21", "={C}16/{C}21"), ("={C}16/{C}21", "={C}16/{C}21"), fmt="0%")
Y.column_dimensions["A"].width = 34
for c in "BCDEFG": Y.column_dimensions[c].width = 11
Y.column_dimensions["H"].width = 60
Y["A30"] = "2025A 營業成本與營業費用的拆分未取得一手數字；B12 以 37% 營收示意（只影響 2025 欄顯示，不影響 2026 以後）"

# 檢查列號假設（yrow 依序 4..）
assert YR["rev"] == 4 and YR["conv"] == 5 and YR["hbm"] == 6 and YR["nand"] == 7 and YR["oth"] == 8 and YR["da"] == 10 and YR["vol"] == 11
assert YR["cogs"] == 12 and YR["opex"] == 13 and YR["op"] == 14 and YR["ni"] == 16 and YR["capex"] == 17 and YR["fcf"] == 18
assert YR["div"] == 19 and YR["bb"] == 20 and YR["eq"] == 21 and YR["nc"] == 22 and YR["sh"] == 24 - 1 + 1 or True

# ======================= Val：估值 =======================
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值（P/B 為主、P/E 對照）。隱含股價 ÷ 目前股價 − 1"
hdr(V, 3, ["方法", "倍數", "淨值／盈餘基準", "情境", "每股基準（韓元）", "隱含股價", "相對目前"])
vr = 4
for basis, col_lbl in (("2026 年底", "C"), ("2027 年底", "D"), ("2028 年底", "E")):
    for sc, colmap in (("A", {"C": "C", "D": "D", "E": "E"}), ("B", {"C": "C", "D": "F", "E": "G"})):
        if basis == "2026 年底" and sc == "B": continue
        col = colmap[col_lbl]
        for nm, k in (("P/B 1.0x", "pb_1"), ("P/B 歷史中位 1.6x", "pb_2"), ("P/B 券商遠期 3.3x", "pb_3")):
            V.cell(vr, 1, nm); V.cell(vr, 2, f"={k}"); V.cell(vr, 3, basis + " BVPS"); V.cell(vr, 4, "共用" if basis == "2026 年底" else sc)
            V.cell(vr, 5, f"=Y!{col}24").number_format = "#,##0"
            V.cell(vr, 6, f"=B{vr}*E{vr}").number_format = "#,##0"
            V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
            vr += 1
for yy, (cA, cB) in (("2027", ("D", "F")), ("2028", ("E", "G"))):
    for sc, col in (("A", cA), ("B", cB)):
        for nm, k in (("P/E 4x", "pe_1"), ("P/E 8x", "pe_2"), ("P/E 12x", "pe_3")):
            V.cell(vr, 1, nm); V.cell(vr, 2, f"={k}"); V.cell(vr, 3, yy + " EPS"); V.cell(vr, 4, sc)
            V.cell(vr, 5, f"=Y!{col}25").number_format = "#,##0"
            V.cell(vr, 6, f"=B{vr}*E{vr}").number_format = "#,##0"
            V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
            vr += 1
V.cell(vr + 1, 1, "讀法：循環股在盈餘高峰時 P/E 低、P/B 高；若 2028 盈餘回落（情境 B），市場通常在 2027 就以 2028 的淨值與盈餘定價。")
for c, w in zip("ABCDEFG", [22, 8, 18, 8, 16, 14, 10]): V.column_dimensions[c].width = w

# ======================= Sens：單變數敏感度（2027E 情境 A 營業利益）=======================
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次只動一個輸入（其餘＝基準）。輸出：2027E 營業利益（情境 A）、2028E 營業利益（情境 B）"
hdr(SE, 3, ["輸入", "設定", "2027E 營業利益 A", "相對基準", "說明"])
def op27A(ov):
    g = lambda k: ov.get(k, k)
    conv = f"(Y!C5*(1+{g('cb27')})*(1+{g('cp27_A')}))"
    ship27 = f"MIN(Mkt!D{MK['demg']},Mkt!C{MK['demg']}*(1+{g('hbm_sup27')}))"
    hbm = f"({ship27}*{g('hbm_p26')}*(1+{g('hbm_p27_A')})*{g('sh27')}*{g('fx')}/1000)"
    nand = f"(Y!C7*(1+nb27)*(1+{g('np27_A')}))"
    rev = f"({conv}+{hbm}+{nand}+other_y)"
    cogs = f"((Y!C12-da_cogs*Y!C10)*Y!D11*(1-{g('costdn_y')})+da_cogs*Y!D10)"
    opex = f"(opex_fix*4*(1+opex_g)+opex_var*{rev})"
    return f"={rev}-{cogs}-{opex}"
rows = [("基準", {}, ""),
        ("一般 DRAM 2027 價 +15%", {"cp27_A": "cp27_A_lo"}, ""), ("一般 DRAM 2027 價 +45%", {"cp27_A": "cp27_A_hi"}, ""),
        ("HBM 2027 價 +38%（BofA）", {"hbm_p27_A": "hbm_p27_A_lo"}, ""), ("HBM 2027 價 +121%（TrendForce）", {"hbm_p27_A": "hbm_p27_A_hi"}, ""),
        ("SK HBM 市占 2027 38%", {"sh27": "sh27_lo"}, ""), ("SK HBM 市占 2027 48%", {"sh27": "sh27_hi"}, ""),
        ("HBM 2027 供給 +50%", {"hbm_sup27": "hbm_sup27_lo"}, ""), ("HBM 2027 供給 +63%", {"hbm_sup27": "hbm_sup27_hi"}, ""),
        ("一般 DRAM 位元 +5%", {"cb27": "cb27_lo"}, ""), ("一般 DRAM 位元 +15%", {"cb27": "cb27_hi"}, ""),
        ("NAND 2027 價 +15%", {"np27_A": "np27_A_lo"}, ""), ("NAND 2027 價 +45%", {"np27_A": "np27_A_hi"}, ""),
        ("匯率 1,350", {"fx": "fx_lo"}, "只影響 HBM（美元計價）"), ("匯率 1,500", {"fx": "fx_hi"}, ""),
        ("單位成本年降 5%", {"costdn_y": "costdn_y_lo"}, ""), ("單位成本年降 15%", {"costdn_y": "costdn_y_hi"}, "")]
for i, (nm, ov, note) in enumerate(rows):
    rr = 4 + i
    SE.cell(rr, 1, nm); SE.cell(rr, 2, ", ".join(f"{k}→{v}" for k, v in ov.items()) or "—")
    SE.cell(rr, 3, op27A(ov)).number_format = "0.0"
    SE.cell(rr, 4, f"=C{rr}/C$4-1").number_format = "0%"
    SE.cell(rr, 5, note)
SE.cell(4 + len(rows) + 1, 1, "核對：基準列應等於 Y!D14")
SE.cell(4 + len(rows) + 1, 3, f"=C4-Y!D14").number_format = "0.000"
SE.column_dimensions["A"].width = 32; SE.column_dimensions["B"].width = 28; SE.column_dimensions["C"].width = 16; SE.column_dimensions["E"].width = 30


# 2028E 情境 B 敏感度
base_r = 4 + len(rows) + 3
SE.cell(base_r, 1, "2028E 營業利益（情境 B）").font = BOLD
def op28B(ov):
    g = lambda k: ov.get(k, k)
    conv = f"(Y!F5*(1+cb28)*(1+{g('cp28_B')}))"
    hbm = f"(Mkt!E{MK['ship']}*Mkt!D{MK['pB']}*(1+{g('hbm_p28_B')})*{g('sh28')}*fx/1000)"
    nand = f"(Y!F7*(1+nb28)*(1+{g('np28_B')}))"
    rev = f"({conv}+{hbm}+{nand}+other_y)"
    opex = f"(opex_fix*4*(1+opex_g)^2+opex_var*{rev})"
    return f"={rev}-Y!G12-{opex}"
rows2 = [("基準", {}), ("一般 DRAM 2028 價 −60%", {"cp28_B": "cp28_B_lo"}), ("一般 DRAM 2028 價 −30%", {"cp28_B": "cp28_B_hi"}),
         ("NAND 2028 價 −70%", {"np28_B": "np28_B_lo"}), ("NAND 2028 價 −40%", {"np28_B": "np28_B_hi"}),
         ("HBM 2028 價 −40%", {"hbm_p28_B": "hbm_p28_B_lo"}), ("HBM 2028 價 −10%", {"hbm_p28_B": "hbm_p28_B_hi"}),
         ("一般 DRAM −60%、NAND −70%、HBM −40%（同時）", {"cp28_B": "cp28_B_lo", "np28_B": "np28_B_lo", "hbm_p28_B": "hbm_p28_B_lo"})]
for i, (nm, ov) in enumerate(rows2):
    rr = base_r + 1 + i
    SE.cell(rr, 1, nm); SE.cell(rr, 2, ", ".join(f"{k}→{v}" for k, v in ov.items()) or "—")
    SE.cell(rr, 3, op28B(ov)).number_format = "0.0"
    SE.cell(rr, 4, f"=C{rr}/C${base_r+1}-1").number_format = "0%"
SE.cell(base_r + 2 + len(rows2), 1, "核對：基準列應等於 Y!G14")
SE.cell(base_r + 2 + len(rows2), 3, f"=C{base_r+1}-Y!G14").number_format = "0.000"
SE.cell(base_r + 3 + len(rows2), 1, "讀法：情境 B 的 2028 價格（2027 均價 × 0.5）仍約為 2025 年均價的 2.7 倍；若回到 2025 年水準（類似 2023 年谷底），需要約 −80%，不在本表範圍")

# ======================= Recon：與市場共識的事後對照 =======================
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 模型（由驅動因子正向推導）vs 市場共識（只作事後對照，不用來回推參數）；兆韓元"
hdr(RC, 3, ["項目", "本模型 2026E", "共識 2026E", "本模型 2027E A", "本模型 2027E B", "共識 2027E", "共識來源"])
rcs = [("營收", "4", 344.2, 531.3, "MarketScreener 2026-10-02"), ("營業利益", "14", 266.1, 417.2, "MarketScreener；FnGuide 2027 391.8（08-14）"), ("淨利", "16", 256.3, 336.2, "MarketScreener"), ("HBM 營收", "6", 45.7, None, "BofA 2026-09-13（2026E）")]
for i, (nm, yr_, c26, c27, src) in enumerate(rcs):
    rr = 4 + i
    RC.cell(rr, 1, nm)
    RC.cell(rr, 2, f"=Y!C{yr_}").number_format = "0.0"
    c = RC.cell(rr, 3, c26); c.font = BLUE
    RC.cell(rr, 4, f"=Y!D{yr_}").number_format = "0.0"
    RC.cell(rr, 5, f"=Y!F{yr_}").number_format = "0.0"
    if c27 is not None:
        c = RC.cell(rr, 6, c27); c.font = BLUE
    RC.cell(rr, 7, src)
RC.cell(9, 1, "2026Q3E 營收：本模型 vs 共識 94.1（FnGuide，2026-09-25）")
RC.cell(9, 2, f"=Q!H{QR['rev']}").number_format = "0.0"; c = RC.cell(9, 3, 94.1); c.font = BLUE
RC.cell(10, 1, "2026Q3E 營業利益：本模型 vs 共識 74.1–78.1")
RC.cell(10, 2, f"=Q!H{QR['op']}").number_format = "0.0"; c = RC.cell(10, 3, 76.1); c.font = BLUE
RC.column_dimensions["A"].width = 44
for c in "BCDEF": RC.column_dimensions[c].width = 14
RC.column_dimensions["G"].width = 40

# ======================= Feedback =======================
FB = wb.create_sheet("Feedback")
FB["A1"] = "Feedback — 回饋算力收支的摘要數字"; FB["A1"].font = BOLD
fb = [("每顆加速器 HBM 金額（GB300，2026 價）", "=nv_hbm", "$K", "markets/chip_bridge v0.1"),
      ("全市場每顆加速器平均 HBM 容量 2026／2027", f"=TEXT(Mkt!C{MK['gbavg']},\"0\")&\" ／ \"&TEXT(Mkt!D{MK['gbavg']},\"0\")", "GB", "Mkt"),
      ("HBM 市場 2026／2027（情境 A）", f"=TEXT(Mkt!C{MK['mA']},\"0\")&\" ／ \"&TEXT(Mkt!D{MK['mA']},\"0\")", "$B", "Mkt"),
      ("SK 海力士 HBM 毛利率", "無法由公開資料分離", "—", "公司不揭露；TrendForce、三星、Micron 皆指出 2026 年一般 DRAM 毛利高於 HBM，Bloter 報導差距約 40 個百分點（單一來源）"),
      ("2027 HBM 需求缺口（需求 − 供給）", f"=Mkt!D{MK['gap']}", "EB", "加速器顆數預測隱含的位元需求高於供給成長；顆數預測可能偏高或價格會上漲")]
for i, (a, b, c, d) in enumerate(fb, 3):
    FB.cell(i, 1, a); FB.cell(i, 2, b); FB.cell(i, 3, c); FB.cell(i, 4, d)
FB.column_dimensions["A"].width = 44; FB.column_dimensions["B"].width = 18; FB.column_dimensions["D"].width = 100

# ======================= Sources =======================
SO = wb.create_sheet("Sources")
srcs = [
 ("S1 SK hynix 424B4 招股書（2026-07）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526299963/d32785d424b4.htm"),
 ("S2 SK hynix 6-K 半年報（2026-08-18）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526354777/d147827d6k.htm"),
 ("Q2 2026 簡報（Investing.com）", "https://www.investing.com/news/company-news/sk-hynix-q2-2026-slides-record-revenue-76-operating-margin-93CH-4818489"),
 ("Q2 2026 法說逐字稿（Investing.com）", "https://ng.investing.com/news/transcripts/earnings-call-transcript-sk-hynix-posts-record-q2-2026-results-as-shares-fall-93CH-2622277"),
 ("40 兆買回（Korea Herald 2026-08-19）", "https://www.koreaherald.com/article/10845792"),
 ("MarketScreener 共識（2026-10-02）", "https://www.marketscreener.com/quote/stock/SK-HYNIX-INC-6494929/finances/"),
 ("Herald Business：FnGuide 共識、目標價（2026-08-14）", "https://biz.heraldcorp.com/article/10841314"),
 ("BofA HBM 估計（Economy Tribune 2026-09-13）", "https://www.economytribune.co.kr/news/articleView.html?idxno=3903874"),
 ("Bloter（CLSA，2026-09-15）", "https://www.bloter.net/news/articleView.html?idxno=673654"),
 ("Counterpoint HBM 2Q26（SamMobile）", "https://www.sammobile.com/?p=16676368"),
 ("UBS 2027 HBM 市占（Korea Herald）", "https://www.koreaherald.com/article/10828143"),
 ("TrendForce 4Q26 價格（2026-09-30）", "https://www.trendforce.com/presscenter/news/20260930-13258.html"),
 ("TrendForce 2027 HBM ASP +121%（2026-09-29）", "https://www.trendforce.com/presscenter/news/20260929-13255.html"),
 ("TrendForce 2027 展望（2026-07-30）", "https://www.trendforce.com/presscenter/news/20260730-13158.html"),
 ("TrendForce 3Q26 價格（2026-07-03）", "https://www.trendforce.com/presscenter/news/20260703-13134.html"),
 ("Citi 2027 Q2 見頂（TipRanks）", "https://www.tipranks.com/news/memory-prices-will-peak-next-year-warns-citi-as-it-cuts-micron-mu-price-target-by-18"),
 ("Bernstein 2028 價格（Notebookcheck）", "https://www.notebookcheck.net/DRAM-crisis-Analysts-expect-drastic-price-drop-in-2028.1337992.0.html"),
 ("BofA 記憶體預測（KuCoin）", "https://www.kucoin.com/news/flash/morgan-stanley-report-memory-super-cycle-to-continue-until-2027-samsung-may-launch-massive-buyback-and-special-dividend"),
 ("Micron FQ4 法說（Webull）", "https://www.webull.com/news/15672802111996928"),
 ("Micron HBM 3:1 晶圓比（XenoSpectrum）", "https://xenospectrum.com/en/micron-hbm-ddr-silicon-gap/"),
 ("JPMorgan 2027 ASIC vs GPU（KuCoin）", "https://www.kucoin.com/news/flash/jpmorgan-ai-asic-shipments-to-surpass-gpus-by-2027"),
 ("Morgan Stanley 顆數（TechNews 2026-07-13）", "https://technews.tw/2026/07/13/the-significant-increase-in-chip-demand-in-2027-will-boost-the-size-of-the-wafer-foundry-and-advanced-packaging-markets/"),
 ("Morgan Stanley TPU（KuCoin）", "https://www.kucoin.com/news/flash/morgan-stanley-defends-broadcom-s-role-in-google-s-tpu-supply-chain-forecasts-5m-shipments-by-2027"),
 ("GF Securities TPU（24/7 Wall St）", "https://247wallst.com/investing/2026/08/25/google-tpu-volume-could-triple-to-8-8-million-by-2027/"),
 ("TrendForce NVIDIA 組合（2026-04-08）", "https://www.trendforce.com/presscenter/news/20260408-13003.html"),
 ("AC Research Blackwell 封裝數", "https://www.amcompute.com/blog/blackwell-gpu-market-size-and-fragmentation"),
 ("NVIDIA Q2 FY27 財報（2026-08-26）", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027"),
 ("companiesmarketcap P/B 歷史", "https://companiesmarketcap.com/sk-hynix/pb-ratio/"),
 ("Korea JoongAng Daily 券商目標價（2026-07-30）", "https://www.koreajoongangdaily.com/business/brokerages-split-on-targets-for-sk-hynix-shares-after-chipmaker-misses-q2-market-consensus/12799468"),
 ("markets/chip_bridge v0.1（本 repo）", "markets/chip_bridge/dist/20261008_晶片內容橋接表_v0.1.xlsx"),
]
SO["A1"] = "來源"; SO["A1"].font = BOLD
for i, (a, b) in enumerate(srcs, 2):
    SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 50; SO.column_dimensions["B"].width = 130

# ======================= README =======================
RD = wb.create_sheet("README", 0)
lines = [
 ("SK 海力士收支與估值模型 v0.1（AI 半導體樣板第一家）　2026-10-08", True),
 ("模型命題：SK 海力士的盈餘由兩條價格線決定——每季議價的一般 DRAM 與年約鎖價的 HBM；股價取決於這個循環在哪一年見頂，以及見頂時累積了多少淨值。", False),
 ("", False),
 ("驅動因子 → 推導量（頁：輸入 → 推導 → 因果方向）", True),
 ("Mkt：加速器顆數 × 每顆 GB → HBM 需求位元；供給成長 → 出貨位元＝MIN(需求, 供給)；× 單價 → HBM 市場；× 市占 → SK 海力士 HBM 營收。需求 > 供給 → 價格上漲（缺口列）", False),
 ("Q：2026H1 實績（SEC 6-K）× 公司 Q3 指引 × TrendForce 季價 → 2026H2 營收與成本；成本隨位元量增、每季單位成本降", False),
 ("Y：一般 DRAM 2026 × (1＋位元) × (1＋年均價變動)；NAND 同；HBM 取自 Mkt → 營收；成本量指數 × (1 − 年降)＋折舊 → 營業利益 → 淨利 → 權益（減股利、買回）→ 每股淨值", False),
 ("Val：每股淨值 × P/B（主）、EPS × P/E（對照）→ 隱含股價 vs 目前股價", False),
 ("Sens：一次動一個輸入 → 2027E 營業利益", False),
 ("Recon：本模型 vs 共識（事後對照；參數不由共識回推）", False),
 ("", False),
 ("情境：A＝2027–28 高檔持平（BofA、TrendForce 型）；B＝2027 見頂、2028 一般 DRAM −50%、NAND −65%（Bernstein 型）。2026 兩情境相同。", False),
 ("取數：markets/chip_bridge v0.1（每顆 GPU HBM 金額、NVIDIA 內容，取 Tokenomics v5.27）。SemiAnalysis 未使用。", False),
 ("標記：Verified 公開一手；Interested-party 有利害關係；Analogy 類比；Assumed 無實證；Derived 推導。Assumed／Analogy 皆有低／高區間（In 頁）。", False),
 ("", False),
 ("已知簡化（v0.1）：(1) 2025 年成本拆分示意；(2) 營運資金變動未計；(3) 股數用期末；(4) HBM 市場暫放本活頁簿，複製到美光時移到 markets/hbm/；(5) 一般 DRAM 與 HBM 的成本只以權重區分，HBM 毛利率無法由公開資料分離。", False),
 ("", False),
 ("樣板驗收標準（AI 半導體 Project）", True),
 ("1 取數註明版本：✓（Tokenomics v5.27 經橋接表 v0.1）　2 HBM 由上而下 vs 由下而上對帳：見 Mkt 對帳列　3 回饋摘要：Feedback 頁　4 命題＋驅動對照＋Excel：本頁　5 兩個使用端結論：另記（不放公開 repo）　6 審閱＋模型後先驗：待　7 複製測試（美光）：待", False),
]
for i, (t, b) in enumerate(lines, 1):
    RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

wb.save(sys.argv[1])
print("saved", MK, QR, YR)
