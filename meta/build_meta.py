# Meta（META）個股模型 v0.1：命題模組＋官方營收分部 → 損益表（docs/plan/個股模型原則_命題加營收分部.md）
# 工作單：meta/docs/workorders/20261030_Meta_v0.1_個股模型.md（B 段）；架構範本 alphabet/build_alphabet.py（v0.3.2）。
# 用法：python3 meta/build_meta.py raw.xlsx [--sens sens.json] [--set 鍵=值 ...]；再以 LibreOffice headless 重算（meta/make_v01.sh 一次跑完）。
# 金額：十億美元（US$B）；每股：美元（Class C 股票股利 1:1 之前的口徑，見 README）；股數：十億股；MW＝IT 關鍵電力 MW。日曆年＝會計年度。
# 欄位（模型頁共用）：C FY2024A、D FY2025A、E 9M26A（內部鍵 "1H"）、F Q4 26E（內部鍵 "2H"）、G FY2026E（＝E＋F）、H–M FY2027–FY2032。評價日 2026-09-30。
import sys, os, json, argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument("out")
ap.add_argument("--sens", help="敏感度重算結果（make_v01.sh 產生）；省略時 Sens 頁結果欄留空")
ap.add_argument("--set", action="append", default=[], help="覆寫 In 頁輸入：鍵=值（敏感度重算用）")
A = ap.parse_args()
OVR = {}
for s in A.set:
    k, v = s.split("=", 1); OVR[k] = float(v)

BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); SUB = PatternFill("solid", fgColor="E2EFDA")
wb = openpyxl.Workbook()

FD = json.load(open(os.path.join(HERE, "data", "meta_facts_2026q3.json"), encoding="utf-8"))
F = {it["id"]: it for it in FD["items"]}
TKV = "v21.7"
TK = json.load(open(os.path.join(HERE, "data", f"tokenomics_snapshot_{TKV}.json"), encoding="utf-8"))
GEN = {x["code"]: x["name"] for x in TK["source"]["generations"]}
CONS = json.load(open(os.path.join(HERE, "data", "consensus_meta_20261030.json"), encoding="utf-8"))
PEER = json.load(open(os.path.join(HERE, "data", "peers_ads_20261030.json"), encoding="utf-8"))
def tk(name, code=None):
    v = TK["items"][name]["values"]
    return v if code is None else v[GEN[code]]["基準"]
def fv(i): return F[i]["value"]
def fs(i): it = F[i]; return f"{it['source']}（{it.get('docDate') or '推導'}；事實總帳 {i}）"

def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

YK = ["24", "25", "1H", "2H", "26", 27, 28, 29, 30, 31, 32]
COL = dict(zip(YK, "CDEFGHIJKLM"))
YH = {"24": "FY2024A", "25": "FY2025A", "1H": "9M26A", "2H": "Q4 26E", "26": "FY2026E", 27: "FY2027E", 28: "FY2028E", 29: "FY2029E",
      30: "FY2030E", 31: "FY2031E", 32: "FY2032E"}
FC = ["2H", 27, 28, 29, 30, 31, 32]
ANN = [27, 28, 29, 30, 31, 32]
LEN = {p: {"1H": 0.75, "2H": 0.25}.get(p, 1) for p in YK}
PY = {"2H": 2026, 27: 2027, 28: 2028, 29: 2029, 30: 2030, 31: 2031, 32: 2032}
def prevp(p): return {"2H": "1H", 27: "2H", 28: 27, 29: 28, 30: 29, 31: 30, 32: 31}[p]
def prevfy(p): return {27: "26", 28: 27, 29: 28, 30: 29, 31: 30, 32: 31}[p]
MODEL_HDR = ["項目", "單位"] + [YH[k] for k in YK] + ["說明"]
NOTE_C = 14

def put(ws, r, label, unit, vals, fmt="#,##0.00", note=None, bold=False, fill=None):
    ws.cell(r, 1, label); ws.cell(r, 2, unit)
    if bold: ws.cell(r, 1).font = BOLD
    for p, v in vals.items():
        c = ws[f"{COL[p]}{r}"] if p in COL else ws[f"{p}{r}"]
        c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")) and isinstance(v, (int, float)): c.font = BLUE
    if note: ws.cell(r, NOTE_C, note)
    if fill:
        for cc in range(1, NOTE_C + 1): ws.cell(r, cc).fill = fill

def widths(ws, a=58, b=10, mid=11, note=90, ncols=11):
    ws.column_dimensions["A"].width = a; ws.column_dimensions["B"].width = b
    for i in range(ncols): ws.column_dimensions[chr(ord("C") + i)].width = mid
    ws.column_dimensions["N"].width = note

RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); Q = wb.create_sheet("Q"); TH = wb.create_sheet("Thesis"); SG = wb.create_sheet("Seg")
PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF"); V = wb.create_sheet("Val"); RC = wb.create_sheet("Recon")
SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")

# =====================================================================================================
# In
IN["A1"] = "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$B；每股 US$（Class C 股票股利前口徑）；股數十億股；MW＝IT MW。標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 一律給低／高區間）"
IN["A2"] = "事實總帳：data/meta_facts_2026q3.json（facts_2026q3.py 兩讀核對，610 筆）；Tokenomics v21.7 快照（data/tokenomics_snapshot_v21.7.json）；共識 data/consensus_meta_20261030.json；同業 data/peers_ads_20261030.json（五家 2026-10-07 基準＋GOOGL）。評價日 2026-09-30，目標時點 2027-10-30。黃底＝判斷題（Andy 可改）"
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明"])
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", judg=False):
    ROWS.append((key, item, unit, lo, base, hi, tag, src, note, judg))
CR_ = CONS["priceReference"]; PT_ = CONS["priceTarget"]; AE = CONS["annualEstimates"]

sec("A. 現價、股數與資產負債（評價日 2026-09-30）")
inp("px", "現價（2026-10-28 收盤；10-29 收盤只找到一個來源）", "US$", None, CR_["close"], None, "Verified",
    "StockScan https://stockscan.io/stocks/META/price-history（10-28 −1.24%）；Digrin、TipRanks（10-27 620.19 − 7.70）相符（consensus_meta_20261030.json priceReference）",
    "工作單 M8：10-29 收盤需兩個來源一致；2026-10-30 再查只有 ChartRow「as of close, Oct 29, 2026」624.55（+1.97%）一個來源（https://chartrow.com/quote/meta），故用 10-28。只影響上檔空間與評等")
inp("px_val", "評價日股價（2026-09-30 收盤；WACC 權益權重用）", "US$", None, 725.18, None, "Verified",
    "TipRanks 歷史股價 https://www.tipranks.com/stocks/meta/historical-prices（2026-10-30 讀取：9/29 738.79、9/30 725.18）；StockScan 9/30 −1.84%（738.79 × 0.9816＝725.20，相符）", "")
inp("sh_out", "流通股數（2026-10-16，10-Q 封面 Class A＋B）", "十億股", None, round(fv("sh.out.20261016") / 1000, 6), None, "Verified", fs("sh.out.20261016"))
inp("sh_dil_q3", "2026Q3 稀釋加權股數", "十億股", None, fv("is.sh_d.2026q3") / 1000, None, "Verified", fs("is.sh_d.2026q3"))
inp("sh_basic_q3", "2026Q3 基本加權股數", "十億股", None, fv("is.sh_b.2026q3") / 1000, None, "Verified", fs("is.sh_b.2026q3"), "期末稀釋股數＝Q3 稀釋加權＋(流通股 − Q3 基本加權)")
inp("cash_0", "現金＋有價證券（2026-09-30）", "US$B", None, (fv("bs.cash.20260930") + fv("bs.mktsec.20260930")) / 1000, None, "Verified", "新聞稿資產負債表：現金 11.222＋有價證券 71.227（事實總帳 bs.cash／bs.mktsec.20260930）")
inp("cash_ye25", "現金＋有價證券（2025-12-31）", "US$B", None, (fv("bs.cash.20251231") + fv("bs.mktsec.20251231")) / 1000, None, "Verified", "新聞稿資產負債表 35.873＋45.719")
inp("debt_0", "票據面額（2026-09-30）", "US$B", None, fv("debt.face.20260930") / 1000, None, "Verified", fs("debt.face.20260930"), "帳面 83.669＝流動 2.747＋非流動 80.922")
inp("debt_ye25", "長期債務帳面（2025-12-31，非流動）", "US$B", None, fv("bs.ltd.20251231") / 1000, None, "Verified", fs("bs.ltd.20251231"), "只作 CF 頁期初顯示；9M26 期末直接取 debt_0")
inp("fin_lease", "融資租賃負債（類債）", "US$B", None, 0.0, None, "Assumed", "未找到：10-Q 讀取未回傳融資租賃負債（A 段報告「已搜尋，未找到」）",
    "暫以 0 計；預測期資本支出含融資租賃本金（公司口徑），已付本金在現金流內。資料缺口：請 Andy 下載 10-Q 完整 PDF", judg=True)
inp("legal_oth", "其他法律相關應計（類債）＝法律應計 19.881 − 協議費用前三季合計 11.89", "US$B", 0.0, round(fv("legal.accrual.20260930") / 1000 - fv("legal.agCharge.ytd9m2026"), 3), None, "Derived",
    fs("legal.accrual.20260930") + "；協議費用 " + fs("legal.agCharge.ytd9m2026"),
    "比照 Alphabet 應計罰款（類債）；低＝0（不扣）。v0.1.1 更正：v0.1 只扣 Q3 的 9.89（得 9.991），完整 10-Q 載協議費用前三季合計 11.89（Q3 以前已認列約 2.00），協議應計已由 ag_y1／ag_rest 現金付款計入", judg=True)
inp("mcps", "6.50% 強制轉換特別股募得現金淨額（2026-10-29 交割，評價日後；毛額 12.5 − 承銷折扣 0.0416）", "US$B", None, fv("eq.mcps.net") / 1000, None, "Verified", fs("eq.mcps.net") + "；" + fs("eq.mcps.uwDisc"),
    "v0.1.1：改用淨額（v0.1 用毛額 12.5）；發行費用未揭露、不扣；超額配售（約 1.8）未知是否行使，不計")
inp("conv_sh", "強制轉換特別股 2029-08 轉換股數（上限）", "十億股", 0.017025, 0.020, 0.020, "Derived",
    "FWP：每股特別股（1,000 美元）轉 1.3620–1.6000 股 Class A；12.5B ÷ 1,000 × 1.6＝20.0M 股（上限）", "取稀釋最大（保守）；轉換前股利照付")
for p, v in (("2H", 0.0), (27, 0.8125), (28, 0.8125), (29, 0.609)):
    inp(f"pref_{p}", f"強制轉換特別股股利 {YH[p]}", "US$B", None, v, None, "Assumed",
        "6.50% × 12.5B＝0.8125／年；首次付息推定 2027-02，2029-08 轉換前最後一期（FY2029 約三季）" if p == "2H" else "")
inp("fwd_on", "Class A 遠期出售於 FY2027 交割（1＝計入；0＝不計）", "x", 0, 1, None, "Assumed",
    fs("eq.fwd.shares") + "：40M 股、初始遠期價 622.91875、2027-11-19 前由公司選擇交割", "交割前公司不收款；目標時點 2027-10-30 前推定已交割（判斷）", judg=True)
inp("fwd_sh", "遠期出售股數", "十億股", None, fv("eq.fwd.shares") / 1000, None, "Verified", fs("eq.fwd.shares"))
inp("fwd_cash", "遠期出售淨額（初始遠期價 622.91875 × 40M 股）", "US$B", None, round(622.91875 * 0.04, 4), None, "Derived", fs("eq.fwd.shares"), "超額配售 6M 股不計")

sec("B. 折現率與目標價規則（同 Alphabet v0.3.2）")
inp("rf", "無風險利率（10 年美債，2026-10-27）", "%", None, 0.0499, None, "Verified",
    "美國財政部每日殖利率曲線（10 Yr）2026-10-27 4.99%（Alphabet v0.3.1 In 頁 rf，2026-10-29 讀取）", "本次重讀財政部網頁被擋（權限逾時），沿用同一讀值")
inp("beta_cr", "原始 β：ChartRow（期間未載）", "x", None, 1.24, None, "Interested-party", "https://chartrow.com/quote/meta（2026-10-30 讀取，as of close Oct 29, 2026）")
inp("beta_cd", "原始 β：CoinDataFlow「Beta (5Y Monthly)」", "x", None, 1.185, None, "Interested-party", "https://coindataflow.com/en/stock/META（Updated 2026-10-30 00:37 UTC）",
    "StockAnalysis 統計頁讀取被擋（權限逾時）")
inp("blume_a", "Blume 調整係數", "x", None, 0.67, None, "Derived", "Blume（1971）；同 Alphabet")
inp("beta", "β（Blume 調整後兩個來源的中位數）", "x", round(0.67 * 1.185 + 0.33, 4), "=MEDIAN(blume_a*beta_cr+1-blume_a,blume_a*beta_cd+1-blume_a)", round(0.67 * 1.24 + 0.33, 4), "Derived",
    "同 Alphabet 修正一 F4", "低／高＝兩個來源各自的調整值")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "同 Alphabet（MAG v0.2.2 val.erp）")
inp("kd_sp", "約 10 年期 AA− 公司債利差", "%", 0.004, 0.007, 0.010, "Assumed", "判斷：Meta 評等 AA−／Aa3 級（未入事實總帳）；10 年期利差約 50–100bp")
inp("kd", "稅前債務成本＝rf＋利差", "%", None, "=rf+kd_sp", None, "Derived", "同 Alphabet")
inp("wacc_adj", "WACC 調整（Sens 格：±1pt）", "%", -0.01, 0.0, 0.01, "Assumed", "同 Alphabet")
inp("g_tv", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "同 Alphabet", judg=True)
inp("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", None, 0.0, 0.05, "Assumed", "同 Alphabet 修正一 F5：RONIC＝WACC", "Sens：WACC＋5pt", judg=True)
inp("mid", "期中慣例：年度現金流距年初", "年", None, 0.5, None, "Derived", "同 Alphabet")
inp("t_val", "評價日（2026-09-30）的年份小數", "年", None, 2026.75, None, "Derived", "同 Alphabet")
inp("t_q4", "Q4 26 現金流折現年數（期中）", "年", None, 0.125, None, "Derived", "2026-10-01～12-31 中點")
inp("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "同 Alphabet（A 25%／B 50%／C 25%）", judg=True)
inp("tp_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "同上", judg=True)
inp("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", judg=True)
inp("w_dcf", "目標價中 DCF 的權重（其餘＝分部加總）", "%", 0.3, 0.5, 0.7, "Assumed", "同 Alphabet（DCF 50%＋分部加總 50%）", judg=True)
inp("t_tgt", "目標時點距評價日（2026-09-30 → 2027-10-30）", "年", None, round(395 / 365, 4), None, "Derived", "395 天 ÷ 365")
inp("w_ntm27", "目標時點 NTM 中 FY2027 的權重", "%", None, round(62 / 365, 4), None, "Derived", "(2027-12-31 − 2027-10-30) ÷ 365＝62 天")
inp("cons_tp", "市場共識目標價（平均）", "US$", None, PT_["mean"], None, "Interested-party",
    f"MarketBeat {PT_['url']}（{PT_['analysts']} 位，頁面 {PT_['pageUpdated']}；consensus_meta_20261030.json）", "TipRanks 851.15（45 位，10-28）、valueinvesting 841.37；賣方有承銷關係（2026-10-26 股權融資），只作對照")
inp("dps_tgt", "評價日到目標時點的每股股利（推到目標價時扣除）", "US$", None, "=dps/4+dps*(1+dps_g)*(1-w_ntm27)", None, "Derived", "Q4 26 一季＋2027 年至 10-30")
inp("internal_ai_reading", "自用 AI 讀法（2＝主值：自用 AI 價值中性，計影子效益；1＝敏感度：只計成本）", "x", 1, 2, None, "Assumed",
    "原則第 7 節（Andy 2026-10-27 15:03）：自用增量 MW × Tokenomics IF_HoldEcon（k＝1）＝影子效益（非財報營收），只進評價口徑", "Meta 全部 AI 算力皆自用（無對外雲端）", judg=True)
inp("sbc_add", "DCF 加回股權報酬（1＝是；Sens 用）", "x", None, 0, 1, "Assumed", "同 Alphabet：股權報酬視為真實成本、不加回", "", judg=True)
inp("tv_alt", "終值改用 FY2032 無槓桿自由現金流 ×(1＋g)（1＝是；Sens 用）", "x", None, 0, 1, "Assumed", "同 Alphabet", "", judg=True)

sec("C. 分部加總倍數：數位廣告同業 NTM EV/EBITDA（M7：五家 2026-10-07 基準＋GOOGL；中位數在 Val 頁計算）")
for p in PEER["peers"]:
    inp(f"pm_ads_{p['ticker'].lower()}", f"數位廣告：{p['name']}（{p['ticker']}）NTM EV/EBITDA", "x", None, p["ntmEvEbitda"], None, "Derived",
        "data/peers_ads_20261030.json（" + ("GOOGL 由 repo 內 Alphabet 資料推算，2026-10-07" if p["ticker"] == "GOOGL" else "2026-10-07 基準，未更新") + "）",
        "EV ÷ NTM EBITDA（MarketScreener 共識 FY1／FY2 時間加權）")
inp("pm_meta_own", "Meta 自身 NTM EV/EBITDA（2026-10-28；只作對照與 Sens）", "x", None, PEER["selfReference"]["ntmEvEbitda"], None, "Derived", "data/peers_ads_20261030.json selfReference（10-07 為 10.79）")
inp("ads_mode", "FoA 與 RL 的倍數（0＝六家中位數；1＝GOOGL；2＝Meta 自身）", "x", None, 0, 2, "Assumed",
    "M7：同業照 A 段檔，預設中位數；Alphabet v0.3.2 對 Alphabet 用大型同業（META）而非小型廣告科技，對應到 Meta 為 GOOGL 14.71×，列 Sens 與判斷題", "Sens：GOOGL 14.71×（高）／Meta 自身 8.99×", judg=True)
inp("rl_mode", "Reality Labs 在分部加總的計法（0＝NTM EBITDA × 廣告倍數，虧損為負值；1＝以 0 計）", "x", None, 0, 1, "Assumed",
    "判斷：同 Alphabet 對 Alphabet 層級負 EBITDA 的處理（× 倍數）；DCF 已逐年扣 RL 虧損", "Sens：RL 以 0 計", judg=True)

sec("D. 命題模組：自用 AI 容量（M4：由資本支出經 Tokenomics 每 MW 資本支出反推，同 Alphabet C10；MTIA 不套 NonNV 係數）")
inp("cx_na", "非 AI 資本支出占營收（辦公室、一般伺服器、網路；其餘＝AI）", "%", 0.02, 0.04, 0.07, "Assumed",
    "判斷：Meta 2019（AI 前）資本支出約占營收兩成含一般資料中心；2025 起公司稱資本支出主要為 AI 基礎設施（法說）", "決定歷年資本支出中 AI 的部分與預測期非 AI 資本支出", judg=True)
inp("cx_23", "FY2023 資本支出（含融資租賃本金）", "US$B", None, 28.10, None, "Interested-party", "Meta 2023 年第四季財報新聞稿（全年 28.10B）；未入本版事實總帳（只用於 2029 年汰換 MW）")
inp("ai_23", "FY2023 及以前機隊：AI 占資本支出比例（換算 2023 年底以前的 AI MW）", "%", 0.5, 0.7, 0.9, "Assumed", "判斷：2023 年 H100 採購開始放量", "只影響 FY2029 汰換資本支出")
inp("cx_24", "FY2024 資本支出（購置不動產與設備＋融資租賃本金）", "US$B", None, -(fv("cf.capex_ppe.fy2024") + fv("cf.fl_principal.fy2024")) / 1000, None, "Verified", "事實總帳 cf.capex_ppe.fy2024＋cf.fl_principal.fy2024")
inp("cx_25", "FY2025 資本支出（含融資租賃本金）", "US$B", None, fv("cf.capexInclFL.fy2025") / 1000, None, "Verified", fs("cf.capexInclFL.fy2025"), "新聞稿文字 72.22 相符")
inp("capex_9m", "2026 年 9M 資本支出（含融資租賃本金）", "US$B", None, fv("cf.capexInclFL.ytd9m2026") / 1000, None, "Verified", fs("cf.capexInclFL.ytd9m2026"), "購置不動產與設備 83.055＋融資租賃本金 2.769")
inp("capex_g", "2026 資本支出指引中點（含融資租賃本金）", "US$B", 135, 140, 145, "Interested-party", fs("txt.guid_capex2026.2026q3"), "C10：FY2026＝9M 實績＋Q4 公式＝本值（Q4 新增 MW 以公式反推，不需迭代）")
inp("cx27_k", "M6：FY2027 資本支出基準值＝2026 指引中點 ×(1＋2026 年增率 × 本係數)", "x", 0.25, 0.5, 0.75, "Assumed",
    "工作單 M6：法說 2027 無數字（逐字稿「until we get to the Q4 call」，事實總帳 call.capex2027）→ 2026 年增率的一半", "Sens：1/4／3/4", judg=True)
inp("cx27_up", "M6：情境 A／C 的 FY2027 資本支出相對 B 的幅度", "%", None, 0.15, None, "Assumed", "工作單 M6：A／C 各 ±15%", "", judg=True)
inp("cx27_B", "情境 B：FY2027 資本支出（含融資租賃本金）", "US$B", None, "=capex_g*(1+(capex_g/cx_25-1)*cx27_k)", None, "Assumed", "M6；共識 FY2027 購置不動產與設備 193.06（MarketScreener，財報前、不含租賃本金）")
inp("cx27_A", "情境 A：FY2027 資本支出（AI 變現提升 → 加碼投資）", "US$B", None, "=cx27_B*(1+cx27_up)", None, "Assumed", "M6：B × 1.15", judg=True)
inp("cx27_C", "情境 C：FY2027 資本支出（變現提升消退 → 放緩）", "US$B", None, "=cx27_B*(1-cx27_up)", None, "Assumed", "M6：B × 0.85", judg=True)
inp("a28", "FY2028 新增 MW ÷ FY2027 新增 MW", "x", 0.8, 1.0, 1.2, "Assumed", "同 Alphabet（2027、2028 同速）；FY2029 起收斂", judg=True)
inp("conv_mw", "FY2029 起 MW 新增速度收斂年數（到 FY2032 只剩 g × 期初 MW）", "年", None, 4, None, "Assumed", "同 Alphabet 修正一 F5", judg=True)
inp("selfb", "自建機房比例（租用部分不計機房資本支出）", "%", 0.75, 0.85, 0.95, "Assumed", "同 Alphabet（MAG v0.2.2 capexModel.selfBuild）；Meta 合資（Hyperion）與租賃在表外", "")
inp("am_26", "2026 新增中 GB300 錨占比（其餘 VR200 錨）", "%", None, 1.0, None, "Assumed", "同 Alphabet；MTIA 無揭露，全部按 GPU 世代（M4）")
inp("am_27", "2027 新增中 GB300 錨占比（其餘 VR200 錨）", "%", None, 0.5, None, "Assumed", "同 Alphabet")
inp("am_28", "2028 起新增中 VR200 錨占比（其餘 GB300 錨）", "%", None, 1.0, None, "Assumed", "同 Alphabet")
inp("cip_0", "在建工程（2026-09-30）", "US$B", None, fv("ppe.cip.20260930") / 1000, None, "Verified", fs("ppe.cip.20260930"), "Q4 26 起轉入在役：折舊同 Alphabet；MW 處理不同——本模型把在建工程對應 MW 計入增量在役 MW（有影子效益與營運成本），Alphabet 修正三不另增 MW（待判斷）")
inp("cip_ai", "在建工程中 AI 占比", "%", 0.8, 0.9, 1.0, "Assumed", "判斷：與 2026 年 AI 占資本支出比例（約 9 成，見 Thesis）一致", "在建工程對應的 MW 在 Q4 26 轉入在役 → 自用增量（影子效益與營運成本）", judg=True)
inp("cip_it", "在建工程中 IT（伺服器與網通）占比", "%", 0.5, 0.6, 0.7, "Assumed", "判斷：Meta 未揭露；Alphabet 2026Q2 約 60%", "IT 依 life_it、其餘依 life_fac 折舊")
inp("life_it", "AI IT 折舊年限（年；汰換週期）", "年", 5, 6, 7, "Assumed", "Tokenomics IF_DeprLifeIT 6 年；Meta 自 2025 起伺服器耐用年限 5.5 年（公開資訊，未入事實總帳）", "整數年（汰換排程用）")
inp("life_fac", "機房折舊年限", "年", 15, 20, 25, "Derived", "Tokenomics IF_CapexFacility 12.67 ÷ IF_DeprFac 0.634＝20 年")
inp("life_na", "非 AI 資本支出折舊年限", "年", 8, 10, 15, "Assumed", "同 Alphabet")
inp("phi_0", "資本支出當期折舊比例（年度批次）", "x", 0.0, 0.25, 0.5, "Assumed", "同 Alphabet")
inp("phi_q4", "Q4 26 批次（在建工程轉入與 Q4 資本支出）當期折舊比例（年）", "x", 0.0625, 0.125, 0.25, "Assumed",
    "判斷：季中上線 → 一季的一半（0.125 年）；最後一年補足 1 − 本值（全期合計＝金額）", "Alphabet 以 phi_0（0.25）計 Q4 批次，使 Q4 折舊跳升；高＝Alphabet 口徑", judg=True)
inp("da_base", "評價日在役資產年折舊（2026Q3 × 4）", "US$B/年", None, fv("cf.da.2026q3") * 4 / 1000, None, "Derived", fs("cf.da.2026q3") + " × 4")
inp("rent_l", "租用算力年租金滿載值（尚未起租租賃＋第三方雲端；只進 Sens）", "US$B/年", None, 0.0, round(fv("lease.notCommenced.20260930") / 1000 / 20, 2), "Assumed",
    fs("lease.notCommenced.20260930") + "：355.73 ÷ 約 20 年租期（假設）＝17.79／年；第三方雲端與 AI token 成本法說未給金額",
    "M4（chat 端審查第 4 條）：不進基準，只放 Recon 與 Sens；FY2027 起 1/3、FY2028 2/3、FY2029 起全額；讀法 2 下與影子效益同樣價值中性（只影響現金流與收支關卡）", judg=True)
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"he_{nm}", f"IF_HoldEcon｜{code}（每 MW 年持有成本，經濟口徑）", "US$m/MW-年", tk("IF_HoldEcon")[GEN[code]]["低成本"], tk("IF_HoldEcon", code), tk("IF_HoldEcon")[GEN[code]]["高成本"], "Derived",
        f"data/tokenomics_snapshot_{TKV}.json（{TK['items']['IF_HoldEcon']['ref']}）", "讀法 2 影子效益的每 MW 值（k＝1）")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"opx_{nm}", f"IF_OpexGW｜{code}（每 MW 年營運費用，不含折舊）", "US$m/MW-年", None, tk("IF_OpexGW", code), None, "Derived", f"data/tokenomics_snapshot_{TKV}.json")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"cit_{nm}", f"IF_CapexIT｜{code}（每 MW IT 資本支出）", "US$m/MW", None, tk("IF_CapexIT", code), None, "Derived", f"data/tokenomics_snapshot_{TKV}.json")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3")):
    inp(f"dpi_{nm}", f"IF_DeprIT｜{code}（每 MW IT 年折舊）", "US$m/MW-年", None, tk("IF_DeprIT", code), None, "Derived", f"data/tokenomics_snapshot_{TKV}.json")
inp("cfac", "IF_CapexFacility（每 MW 機房資本支出）", "US$m/MW", None, tk("IF_CapexFacility", "VR200"), None, "Derived", f"data/tokenomics_snapshot_{TKV}.json")

sec("E. 營收分部（Seg 頁）：Advertising＝曝光 × 價格；情境軸＝AI 對每則廣告價格的提升")
inp("conv_n", "成長率收斂年數（FY2026 起始 → 長期值）", "年", None, 4, None, "Assumed", "同 Alphabet：近四季年增在 FY2030 收斂到長期值")
inp("g0_q", "起始年增改用最新一季（1＝是；0＝近四季）", "x", None, 0, 1, "Assumed",
    "工作單 B 段第 3 項：近四季年增；2026Q3 曝光年增 +10%（近四季平均 +15.25%），法說未給 2027 營收展望", "Sens 用", judg=True)
inp("gl_impr", "廣告曝光數長期年增", "%", 0.03, 0.05, 0.07, "Assumed", "判斷：DAP 年增 3%（2026Q3 新聞稿）＋使用時間與新版位（Threads、WhatsApp）", "")
inp("gl_price", "每則廣告平均價格長期年增", "%", 0.01, 0.03, 0.05, "Assumed", "判斷：約等於名目 GDP 與廣告通膨", "")
inp("pA", "情境 A：AI 對每則廣告價格的提升（相對 B 路徑）FY2030 達到", "%", 0.05, 0.10, 0.15, "Assumed",
    "AI 推薦與廣告工具（Advantage+、Muse）提高轉換率 → 價格；2026 年每則價格年增 +12%／+12%／+14%（事實總帳 txt.price_yoy）", "線性：FY2026 為 0 → FY2030 全額，之後持平", judg=True)
inp("pC", "情境 C：AI 價格提升消退／競爭（相對 B 路徑）FY2030 達到", "%", -0.15, -0.10, -0.05, "Assumed",
    "AI 答案與代理分流廣告曝光、同業（Alphabet、TikTok）AI 工具追上 → 價格溢價消退", "", judg=True)
inp("inc_m", "價格情境差額的增量 EBITDA 率", "%", 0.75, 0.85, 0.95, "Assumed", "判斷：價格提升幾乎不增加成本（Meta 無 TAC）", "", judg=True)
inp("gl_other", "FoA 其他營收長期年增", "%", 0.08, 0.15, 0.25, "Assumed", "近四季 +64.6%（WhatsApp 商業訊息、Meta One 訂閱）", "")
inp("gl_rl", "Reality Labs 營收長期年增", "%", -0.05, 0.05, 0.15, "Assumed", "近四季 −4.5%", "")
inp("rl_loss26", "2026 Reality Labs 虧損（指引中點）", "US$B", 18.0, 18.5, 19.0, "Interested-party", fs("txt.guid_rl_loss2026.2026q3"), "Q4 26 虧損＝本值 − 9M 實績")
inp("rl_loss27", "2027 Reality Labs 虧損", "US$B", 12.0, 15.0, 18.0, "Interested-party", "法說 CEO／CFO：約 15B（事實總帳 call.rlLoss2027）", "")
inp("rl_loss32", "FY2032 Reality Labs 虧損（FY2028 起線性）", "US$B", 0.0, 5.0, 10.0, "Assumed", "法說：之後「trending towards profitability」（事實總帳 A 段報告第 5 節）", "", judg=True)
inp("rl_da_sh", "折舊中歸 Reality Labs 的比例（基礎折舊）", "%", 0.02, 0.05, 0.08, "Assumed", "判斷：公司不揭露分部折舊", "只影響分部 EBITDA 拆分")
inp("mb_q3", "FoA EBITDA 率起點改用 2026Q3 單季（1＝是；0＝9M26）", "x", None, 0, 1, "Assumed", "同 Alphabet 預設 9M26（最新年初至今）；兩者皆排除一次性項目", "Sens 用", judg=True)
inp("ot_legal_q2", "一次性：2026Q2 法律費用", "US$B", None, fv("txt.legal_charge.2026q2"), None, "Verified", fs("txt.legal_charge.2026q2"), "M5 擴大（chat 端審查第 3 條）")
inp("ot_sev_q2", "一次性：2026Q2 遣散費用", "US$B", None, fv("txt.severance.2026q2"), None, "Verified", fs("txt.severance.2026q2"), "未揭露分部；推定 FoA")
inp("ot_legal_q3", "一次性：2026Q3 法律費用（多州檢察長協議）", "US$B", None, fv("txt.legal_charge.2026q3"), None, "Verified", fs("txt.legal_charge.2026q3"), "新聞稿未寫分部；FoA 營業利益跌幅顯示在 FoA（推論）")
inp("ot_tax_q1", "一次性：2026Q1 CAMT 所得稅利益", "US$B", None, fv("txt.tax_benefit_camt.2026q1"), None, "Verified", fs("txt.tax_benefit_camt.2026q1"))
inp("eps_x_q3", "2026Q3 排除法律費用後稀釋 EPS（公司揭露）", "US$", None, 7.40, None, "Interested-party",
    "Q3 8-K 附件 99.1：「Diluted EPS of $3.83 would have increased by $3.57 to $7.40」；稅率 24% 排除後為 17%（2026-10-30 讀取；獨立核對同讀）", "公司自行計算的排除後數字")
inp("ot_legal_at", "法律費用稅後保留比例（＝(7.40 − 3.83) × Q3 稀釋股數 ÷ 9.89）", "%", None, f"=(eps_x_q3-{fv('is.eps_d.2026q3')})*sh_dil_q3/ot_legal_q3", None, "Derived",
    "Q3 稅後影響約 9.18、稅盾約 0.70（法說：「non-deductible portion」）", "Q2 法律 2.40 套同一比例（Assumed）；v0.1 初版假設全額不可抵稅，已更正")

sec("F. 損益與現金流")
inp("tax", "有效稅率（預測期）", "%", 0.13, 0.16, 0.19, "Interested-party", fs("txt.guid_tax_q4.2026q3"), "2026Q4 指引 15–17% 中點；延用到預測期")
inp("r_cash", "現金收益率", "%", None, 0.04, None, "Assumed", "判斷：短期公債約 4%")
inp("r_debt", "債務利率", "%", None, 0.048, None, "Assumed", "判斷：Meta 票據票面約 4.3–5.75%（未入事實總帳）")
inp("sbc_r", "股權報酬占營收（近四季）", "%", None, round(fv("ttm.sbc.2026q3") / fv("ttm.total.2026q3"), 4), None, "Derived", "近四季 SBC 26.943 ÷ 營收 241.606（事實總帳 ttm.*）", "營運現金加回；DCF 不加回")
inp("dil", "股數年稀釋（股權報酬）", "%", 0.0, 0.014, 0.02, "Derived", "2026 年 9M 基本加權股數 2,525 → 2,552（三季 +1.07%，年化約 1.4%；暫停買回）")
inp("wc_k", "營運資金變動 ÷ 營收增量（流入）", "%", -0.02, 0.0, 0.02, "Assumed", "判斷：預設 0")
inp("dps", "普通股每股年股利", "US$", None, round(-4 * fv("cf.div.2026q3") / fv("is.sh_b.2026q3"), 4), None, "Derived", "2026Q3 股利支付 1.361 ÷ 基本加權 2.552 × 4（含股利約當）")
inp("dps_g", "每股股利年成長", "%", None, 0.05, None, "Assumed", "判斷（2025 年 0.50 → 0.525）")
inp("bb_t", "回購目標（FY2027 起每年，現金足夠時）", "US$B", 0, 0.0, -fv("cf.buyback.fy2025") / 1000, "Assumed", "2026 年 9M 回購 0、10-26 起對外募股 → 預設 0", "高＝2025 全年 26.248", judg=True)
inp("min_cash", "最低現金", "US$B", None, 40, None, "Assumed", "判斷")
inp("mat_27", "債務到期 FY2027（2026-09-30 一年內到期）", "US$B", None, fv("debt.current.20260930") / 1000, None, "Verified", fs("debt.current.20260930"), "其他年度到期即再融資（不影響淨現金），未逐年列")
inp("ag_pay", "多州檢察長協議年付款（12.7B ÷ 10 年）", "US$B", None, round(fv("legal.agSettle.total") / 1000 / 10, 3), None, "Derived", fs("legal.agSettle.total"),
    "chat 端審查第 3 條：留在現金流與 DCF；推定 FY2027–FY2036 每年等額（Assumed）；FY2033 以後 4 期以 WACC 折現列類債。資料缺口：流動法律應計 6.87 → 11.58（+4.7），第一年付款可能大於 1.27（前高後低），影響 FY2027 收支關卡、不影響目標價，待 10-Q 完整 PDF")
inp("contg", "協議或有金額（尚未認列；只進 Sens）", "US$B", None, 0.0, fv("legal.agContingent") / 1000, "Verified", fs("legal.agContingent"), "chat 端審查第 3 條：列 Sens 一列，不進基準", judg=True)

sec("H. 共識（Recon；MarketScreener 頁面報價日 2026-10-08，Q3 財報前）")
for k, lab in (("revenue", "營收"), ("ebit", "營業利益"), ("eps", "EPS"), ("capex", "購置不動產與設備（不含租賃本金）"), ("fcf", "自由現金流"), ("ebitda", "EBITDA（含 SBC 加回）")):
    for i, y in enumerate((26, 27, 28)):
        inp(f"c_{k}_{y}", f"共識 {lab} FY20{y}", "US$" if k == "eps" else "US$B", None, AE[f"20{y}"][k], None, "Interested-party", AE["url"] + "（頁面 2026-10-08）" if i == 0 else "同上")
inp("c_q4rev", "共識 2026Q4 營收（TipRanks）", "US$B", None, CONS["quarterlyEstimates"]["2026Q4"]["revenue"], None, "Interested-party", CONS["quarterlyEstimates"]["url"])
inp("c_eps26_vi", "共識 FY2026 EPS（valueinvesting，推測排除一次性）", "US$", None, CONS["annualEstimatesAlt"]["2026"]["eps"], None, "Interested-party", CONS["annualEstimatesAlt"]["url"], "未載日期與口徑")

IN_ROW = {}
r = 4
for row in ROWS:
    if row[0] == "sec":
        IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src, note, judg = row
    if key in OVR: base = OVR[key]
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            c.number_format = "0.0%" if unit == "%" else ("0.0000" if unit == "x" else "#,##0.000")
    IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, note)
    if judg:
        for cc in range(1, 10): IN.cell(r, cc).fill = KEY
    dn(key, f"In!$E${r}")
    IN_ROW[key] = r
    r += 1
r += 1
IN.cell(r, 1, "J. 情境驅動對照（三情境並排，進主模型的加權目標價；本節只引用上方各列，修改請改上方藍字。上方「敏感度低／高」欄只用於 Sens 頁，不進目標價）").font = BOLD; r += 1
hdr(IN, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]); r += 1
for key, item, unit, ks, fmt, note in (
        ("scn_p", "AI 對每則廣告價格的提升 FY2030 達到（相對 B 路徑）", "%", ("pA", None, "pC"), "0.0%", "情境 B＝近四季成長收斂，不另加調整（0）"),
        ("scn_cx27", "FY2027 資本支出（含融資租賃本金）", "US$B", ("cx27_A", "cx27_B", "cx27_C"), "#,##0.0", "M6：B＝2026 指引中點 ×(1＋2026 年增率的一半)，A／C 各 ±15%"),
        ("scn_w", "情境權重", "%", ("tp_wA", "tp_wB", "tp_wC"), "0%", "[判斷] 黃底，Andy 可改")):
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, k_ in zip((4, 5, 6), ks):
        IN.cell(r, j, f"=E{IN_ROW[k_]}" if k_ else 0).number_format = fmt
    IN.cell(r, 7, "Derived"); IN.cell(r, 8, "引用 In 頁 " + "／".join(k_ for k_ in ks if k_) + "；" + note)
    r += 1
for col, w in zip("ABCDEFGHI", [14, 64, 12, 10, 11, 10, 16, 90, 80]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# =====================================================================================================
# Q：實績（8 季＋FY2024、FY2025）
Q["A1"] = "Q — 實績（US$B；EPS 美元；股數十億股）。資料取自 data/meta_facts_2026q3.json（四季財報新聞稿兩讀核對：SEC 8-K 附件 99.1 與 IR 網站 PDF；比較欄取自後一季新聞稿）[Verified]"
Q["A2"] = "核對列（實績必須為 0）：Advertising＋Other revenue＝Family of Apps、分部合計＝合併營收、分部營業利益合計＝合併營業利益、營收 − 成本費用＝營業利益、FCF＝營運現金 − 資本支出（含租賃本金）；FY2025 四季合計對年報取超出四捨五入容差（0.002）的部分"
QCOLS = ["2024q4", "2025q1", "2025q2", "2025q3", "2025q4", "2026q1", "2026q2", "2026q3", "fy2024", "fy2025"]
QC = dict(zip(QCOLS, "BCDEFGHIJK"))
hdr(Q, 3, ["項目"] + [c.upper() for c in QCOLS] + ["9M26", "近四季 TTM", "前四季", "標記", "來源"])
QR = {}
SRCQ = "Meta 各季財報新聞稿（8-K 附件 99.1；事實總帳 data/meta_facts_2026q3.json）"
def qrow(r, key, label, fid=None, scale=1000, neg=False, formula=None, fmt="#,##0.000", h1="sum", ttm=True, tag="Verified", src=SRCQ, cols=None):
    Q.cell(r, 1, label); QR[key] = r
    for c in (cols or QCOLS):
        col = QC[c]
        if formula:
            Q[f"{col}{r}"] = "=" + formula(col)
        else:
            iid = f"{fid}.{c}"
            if iid not in F: continue
            v = F[iid]["value"]
            Q[f"{col}{r}"] = (-v if neg else v) / scale
            Q[f"{col}{r}"].font = BLUE
        Q[f"{col}{r}"].number_format = fmt
    if h1 == "sum": Q[f"L{r}"] = f"=G{r}+H{r}+I{r}"
    elif h1: Q[f"L{r}"] = h1
    if ttm:
        Q[f"M{r}"] = f"=SUM(F{r}:I{r})"; Q[f"N{r}"] = f"=SUM(B{r}:E{r})"
    for cc in "LMN": Q[f"{cc}{r}"].number_format = fmt
    Q.cell(r, 15, tag); Q.cell(r, 16, src)

r = 4
Q.cell(r, 1, "營收（官方分部與子列）").font = BOLD; r += 1
qrow(r, "ads", "Advertising", "seg.ads"); r += 1
qrow(r, "other", "Other revenue", "seg.other"); r += 1
qrow(r, "foa", "Family of Apps（分部）", "seg.foa"); r += 1
qrow(r, "chk_foa", "檢查：Advertising＋Other revenue－Family of Apps", formula=lambda c: f"{c}{QR['ads']}+{c}{QR['other']}-{c}{QR['foa']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "rl", "Reality Labs（分部）", "seg.rl"); r += 1
qrow(r, "rev", "合併營收", "is.rev"); r += 1
qrow(r, "chk_rev", "檢查：分部合計－合併營收", formula=lambda c: f"{c}{QR['foa']}+{c}{QR['rl']}-{c}{QR['rev']}", fmt="0.000", tag="—", src="—"); r += 1
r += 1
Q.cell(r, 1, "分部營業利益").font = BOLD; r += 1
qrow(r, "oi_foa", "Family of Apps 營業利益", "seg.oi_foa"); r += 1
qrow(r, "oi_rl", "Reality Labs 營業損益", "seg.oi_rl"); r += 1
qrow(r, "oi", "合併營業利益", "is.ebit"); r += 1
qrow(r, "chk_oi", "檢查：分部營業利益合計－合併營業利益", formula=lambda c: f"{c}{QR['oi_foa']}+{c}{QR['oi_rl']}-{c}{QR['oi']}", fmt="0.000", tag="—", src="—"); r += 1
r += 1
Q.cell(r, 1, "合併損益").font = BOLD; r += 1
qrow(r, "cogs", "營業成本", "is.cogs"); r += 1
qrow(r, "rd", "研發費用", "is.rd"); r += 1
qrow(r, "ms", "行銷費用", "is.ms"); r += 1
qrow(r, "ga", "管理費用（含法律費用）", "is.ga"); r += 1
qrow(r, "chk_cost", "檢查：營收 − 成本費用 − 營業利益", formula=lambda c: f"{c}{QR['rev']}-{c}{QR['cogs']}-{c}{QR['rd']}-{c}{QR['ms']}-{c}{QR['ga']}-{c}{QR['oi']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "oie", "利息與其他收益淨額", "is.oie"); r += 1
qrow(r, "pt", "稅前利益", "is.pretax"); r += 1
qrow(r, "chk_pt", "檢查：營業利益＋業外 − 稅前利益", formula=lambda c: f"{c}{QR['oi']}+{c}{QR['oie']}-{c}{QR['pt']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "taxa", "所得稅（負值＝利益）", "is.tax"); r += 1
qrow(r, "ni", "淨利", "is.ni"); r += 1
qrow(r, "chk_ni", "檢查：稅前 − 所得稅 − 淨利", formula=lambda c: f"{c}{QR['pt']}-{c}{QR['taxa']}-{c}{QR['ni']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "sh", "稀釋加權股數（十億股）", "is.sh_d", h1="={0}".format(fv("is.sh_d.ytd9m2026") / 1000), ttm=False); r += 1
qrow(r, "eps", "EPS（GAAP 稀釋，美元）", "is.eps_d", scale=1, fmt="0.00", h1="={0}".format(fv("is.eps_d.ytd9m2026"))); r += 1
r += 1
Q.cell(r, 1, "現金流與股東回饋").font = BOLD; r += 1
qrow(r, "da", "折舊與攤銷", "cf.da"); r += 1
qrow(r, "sbc", "股權報酬（SBC）", "cf.sbc"); r += 1
qrow(r, "cfo", "營運現金流", "cf.cfo"); r += 1
qrow(r, "cxp", "購置不動產與設備", "cf.capex_ppe", neg=True); r += 1
qrow(r, "flp", "融資租賃本金支付", "cf.fl_principal", neg=True); r += 1
qrow(r, "capex", "資本支出（含融資租賃本金；公司口徑）", formula=lambda c: f"{c}{QR['cxp']}+{c}{QR['flp']}", tag="Derived"); r += 1
qrow(r, "fcf", "自由現金流（公司定義）", "cf.fcf"); r += 1
qrow(r, "chk_fcf", "檢查：營運現金流 − 資本支出（含租賃本金）− 自由現金流", formula=lambda c: f"{c}{QR['cfo']}-{c}{QR['capex']}-{c}{QR['fcf']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "bb", "買回", "cf.buyback", neg=True); r += 1
qrow(r, "div", "股利", "cf.div", neg=True); r += 1
r += 1
Q.cell(r, 1, "命題指標與一次性項目（新聞稿文字段；%＝年增）").font = BOLD; r += 1
qrow(r, "impr", "廣告曝光數年增", "txt.impr_yoy", scale=100, fmt="0.0%", h1=None, ttm=False, cols=["2025q4", "2026q1", "2026q2", "2026q3"]); r += 1
qrow(r, "price", "每則廣告平均價格年增", "txt.price_yoy", scale=100, fmt="0.0%", h1=None, ttm=False, cols=["2025q4", "2026q1", "2026q2", "2026q3"]); r += 1
qrow(r, "dap", "每日活躍人數 DAP（十億人）", "txt.dap", scale=1, fmt="0.00", h1=None, ttm=False, cols=["2025q4", "2026q1", "2026q2", "2026q3"]); r += 1
qrow(r, "legal", "一次性：法律費用（營業利益內）", "txt.legal_charge", scale=1, h1="sum", ttm=False, cols=["2026q2", "2026q3"]); r += 1
qrow(r, "sev", "一次性：遣散費用（營業利益內）", "txt.severance", scale=1, h1="sum", ttm=False, cols=["2026q2"]); r += 1
qrow(r, "taxb", "一次性：CAMT 所得稅利益", "txt.tax_benefit_camt", scale=1, h1="sum", ttm=False, cols=["2026q1"]); r += 1
for k in ("legal", "sev", "taxb"):
    Q[f"L{QR[k]}"] = f"=SUM(G{QR[k]}:I{QR[k]})"
r += 1
Q.cell(r, 1, "核對彙總（實績必須為 0）").font = BOLD; r += 1
QCHK = ["chk_foa", "chk_rev", "chk_oi", "chk_cost", "chk_pt", "chk_ni", "chk_fcf"]
Q.cell(r, 1, "檢查：上列核對列絕對值合計（8 季＋FY2024＋FY2025）"); QR["chk_all"] = r
Q[f"B{r}"] = "=" + "+".join(f"SUMPRODUCT(ABS(B{QR[k]}:K{QR[k]}))" for k in QCHK); Q[f"B{r}"].number_format = "0.000"; r += 1
FYK = ["ads", "other", "foa", "rl", "oi_foa", "oi_rl", "cogs", "rd", "ms", "ga", "cfo", "da", "sbc", "cxp", "flp", "ni"]
Q.cell(r, 1, "檢查：FY2025 四季合計 − 年報（超出四捨五入容差 0.002 的部分）"); QR["chk_fy"] = r
Q[f"B{r}"] = "=" + "+".join(f"MAX(0,ABS(SUM(C{QR[k]}:F{QR[k]})-K{QR[k]})-0.002)" for k in FYK); Q[f"B{r}"].number_format = "0.000"; r += 1
YTF = {"ads": "seg.ads", "other": "seg.other", "rl": "seg.rl", "rev": "is.rev", "oi_foa": "seg.oi_foa", "oi_rl": "seg.oi_rl", "oi": "is.ebit", "ni": "is.ni", "da": "cf.da", "sbc": "cf.sbc", "cfo": "cf.cfo", "cxp": "cf.capex_ppe", "flp": "cf.fl_principal"}
Q.cell(r, 1, "檢查：9M26（三季合計）− 2026Q3 新聞稿年初至今欄（超出四捨五入容差 0.002 的部分）"); QR["chk_ytd"] = r
Q[f"B{r}"] = "=" + "+".join(f"MAX(0,ABS(L{QR[k]}-({abs(fv(YTF[k] + '.ytd9m2026')) / 1000 if k in ('cxp', 'flp') else fv(YTF[k] + '.ytd9m2026') / 1000}))-0.002)" for k in YTF); Q[f"B{r}"].number_format = "0.000"; r += 2

Q.cell(r, 1, "推導量（Seg、PL 頁的起始值；[Derived]）").font = BOLD; r += 1
hdr(Q, r, ["項目", "數值", "說明"]); r += 1
def qd(key, label, formula, fmt="0.0%", note=""):
    global r
    Q.cell(r, 1, label); Q[f"B{r}"] = formula; Q[f"B{r}"].number_format = fmt; Q.cell(r, 3, note)
    dn(key, f"Q!$B${r}"); QR[key] = r; r += 1
qd("g0_ads_ttm", "近四季年增：Advertising（金額）", f"=M{QR['ads']}/N{QR['ads']}-1", note="(2025Q4–2026Q3) ÷ (2024Q4–2025Q3) − 1")
qd("g0_ads_q", "最新一季年增：Advertising（2026Q3）", f"=I{QR['ads']}/E{QR['ads']}-1")
qd("g0_impr_ttm", "近四季平均：廣告曝光數年增", f"=AVERAGE(F{QR['impr']}:I{QR['impr']})", note="公司只揭露年增率；四季簡單平均")
qd("g0_impr", "起始：廣告曝光數年增（g0_q＝1 時用 2026Q3）", f"=IF(g0_q=1,I{QR['impr']},g0_impr_ttm)")
qd("g0_price", "起始：每則廣告價格年增（＝(1＋廣告金額年增) ÷ (1＋曝光年增) − 1）", f"=(1+IF(g0_q=1,g0_ads_q,g0_ads_ttm))/(1+g0_impr)-1",
   note="以金額年增反推，使曝光 × 價格＝廣告金額年增；近四季公司揭露價格年增平均 +11.0%")
qd("g0_other", "近四季年增：Other revenue", f"=IF(g0_q=1,I{QR['other']}/E{QR['other']}-1,M{QR['other']}/N{QR['other']}-1)")
qd("g0_rl", "近四季年增：Reality Labs 營收", f"=IF(g0_q=1,I{QR['rl']}/E{QR['rl']}-1,M{QR['rl']}/N{QR['rl']}-1)")
qd("ot_oi_9m", "9M26 一次性項目（營業利益內：法律 2.40＋遣散 1.18＋法律 9.89）", "=ot_legal_q2+ot_sev_q2+ot_legal_q3", "#,##0.000")
qd("m_foa_9m", "FoA EBITDA 率（9M26；排除一次性；營業利益＋一次性＋折舊 ×(1 − RL 占比)）", f"=(L{QR['oi_foa']}+ot_oi_9m+L{QR['da']}*(1-rl_da_sh))/L{QR['foa']}")
qd("m_foa_q3", "FoA EBITDA 率（2026Q3 單季；排除一次性）", f"=(I{QR['oi_foa']}+ot_legal_q3+I{QR['da']}*(1-rl_da_sh))/I{QR['foa']}")
qd("m_foa_25", "對照：FoA EBITDA 率（FY2025）", f"=(K{QR['oi_foa']}+K{QR['da']}*(1-rl_da_sh))/K{QR['foa']}")
qd("m_foa", "FoA EBITDA 率（預測期採用；自用 AI 增量成本另計）", "=IF(mb_q3=1,m_foa_q3,m_foa_9m)")
qd("rev_9m", "9M26 合併營收", f"=L{QR['rev']}", "#,##0.000")
for col in "BCDEFGHIJKLMN": Q.column_dimensions[col].width = 11
Q.column_dimensions["A"].width = 60; Q.column_dimensions["O"].width = 14; Q.column_dimensions["P"].width = 70
Q.freeze_panes = "B4"
def QA(k, p):
    return {"24": f"Q!J{QR[k]}", "25": f"Q!K{QR[k]}", "1H": f"Q!L{QR[k]}"}[p]

# =====================================================================================================
# Seg：營收分部（三情境）
SG["A1"] = "Seg — 營收分部層：官方分部 Family of Apps（子列 Advertising、Other revenue）、Reality Labs。命題模組：Advertising＝廣告曝光數 × 每則廣告平均價格；情境差異只在每則價格（A：AI 提升／B：近四季成長收斂／C：提升消退）"
SG["A2"] = "成長路徑：Q4 26 年增＝近四季年增（Q 頁），之後線性收斂到長期值（conv_n 年，FY2030 到達）。曝光與價格各自收斂；價格情境＝B 路徑 ×(1＋情境調整)，調整自 FY2026 的 0 線性到 FY2030 的 In 頁值。非 AI 資本支出＝營收 × cx_na"
SCN = {"A": ("pA", "情境 A（AI 提高每則廣告價格）"), "B": ("0", "情境 B（近四季成長收斂；基準）"), "C": ("pC", "情境 C（AI 價格提升消退）")}
SR = {s: {} for s in "ABC"}
r = 4
def gpath(g0, gl):
    return {**{"2H": f"={g0}"}, **{p: f"={g0}+({gl}-{g0})*MIN(1,({PY[p]}-2026)/conv_n)" for p in ANN}}
def sput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(SG, r, label, unit, vals, fmt, note, **kw); SR[s][key] = r; r += 1
def depf(amt, age, life, f0="phi_0"):
    a = f"({age})"
    return f"{amt}/{life}*IF({a}<0,0,IF({a}=0,{f0},IF({a}<{life},1,IF({a}={life},1-{f0},0))))"
def dep_row(srcs, life):
    out = {}
    for p in FC:
        terms = [depf(a, PY[p] - vy, life, "phi_q4" if vy == 2026 else "phi_0") for a, vy in srcs if PY[p] >= vy]   # Q4 26 批次當期 phi_q4、最後一年 1 − phi_q4；全期合計＝金額
        out[p] = "=" + "+".join(terms)
    return out
for s in "ABC":
    padj, title = SCN[s]
    SG.cell(r, 1, title).font = BOLD; SG.cell(r, 1).fill = SUB; r += 1
    hdr(SG, r, MODEL_HDR); r += 1
    R = SR[s]
    sput(s, "g_impr", "  廣告曝光數年增（B 路徑）", "%", gpath("g0_impr", "gl_impr"), "0.0%", "Q4 26＝近四季平均（Q 頁），FY2030 收斂到 gl_impr")
    sput(s, "g_price", "  每則廣告價格年增（B 路徑）", "%", gpath("g0_price", "gl_price"), "0.0%", "Q4 26＝由近四季廣告金額年增反推，FY2030 收斂到 gl_price")
    v = {"24": "=" + QA("ads", "24"), "25": "=" + QA("ads", "25"), "1H": "=" + QA("ads", "1H"),
         "2H": f"=Q!F{QR['ads']}*(1+F{R['g_impr']})*(1+F{R['g_price']})", "26": f"=E{r}+F{r}"}
    for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{R['g_impr']})*(1+{COL[p]}{R['g_price']})"
    sput(s, "adsB", "Advertising：B 路徑（曝光 × 價格）", "US$B", v, note="Q4 26＝2025Q4 × (1＋曝光年增) × (1＋價格年增)；年度＝前一年 × 同式")
    sput(s, "ramp", "  情境爬坡（FY2026 0 → FY2030 1）", "x", {**{"2H": 0}, **{p: f"=MIN(1,({PY[p]}-2026)/conv_n)" for p in ANN}}, "0.00")
    sput(s, "padj", "  每則價格情境調整（相對 B 路徑）", "%", {p: f"={padj}*{COL[p]}{R['ramp']}" for p in FC}, "0.0%",
         {"A": "In 頁 pA", "B": "基準不另調整", "C": "In 頁 pC"}[s])
    v = {p: f"={COL[p]}{R['adsB']}" for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['adsB']}*(1+{COL[p]}{R['padj']})"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "ads", "Advertising", "US$B", v, bold=True)
    sput(s, "g_ads", "  Advertising 年增", "%", {p: f"={COL[p]}{r - 1}/{COL[q]}{r - 1}-1" for p, q in (("25", "24"), ("26", "25"), (27, "26"), (28, 27), (29, 28), (30, 29), (31, 30), (32, 31))}, "0.0%")
    sput(s, "g_other", "  Other revenue 年增", "%", gpath("g0_other", "gl_other"), "0.0%", "三情境相同")
    v = {"24": "=" + QA("other", "24"), "25": "=" + QA("other", "25"), "1H": "=" + QA("other", "1H"), "2H": f"=Q!F{QR['other']}*(1+F{R['g_other']})", "26": f"=E{r}+F{r}"}
    for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{R['g_other']})"
    sput(s, "other", "Other revenue", "US$B", v, bold=True)
    v = {p: "=" + QA("foa", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['ads']}+{COL[p]}{R['other']}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "foa", "Family of Apps", "US$B", v, bold=True, fill=SUB)
    sput(s, "g_rl", "  Reality Labs 營收年增", "%", gpath("g0_rl", "gl_rl"), "0.0%", "三情境相同")
    v = {"24": "=" + QA("rl", "24"), "25": "=" + QA("rl", "25"), "1H": "=" + QA("rl", "1H"), "2H": f"=Q!F{QR['rl']}*(1+F{R['g_rl']})", "26": f"=E{r}+F{r}"}
    for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{R['g_rl']})"
    sput(s, "rl", "Reality Labs", "US$B", v, bold=True, fill=SUB)
    v = {p: "=" + QA("rev", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['foa']}+{COL[p]}{R['rl']}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "rev", "合併營收", "US$B", v, bold=True)
    sput(s, "chk1", "檢查：Advertising＋Other revenue－Family of Apps", "US$B", {p: f"={COL[p]}{R['ads']}+{COL[p]}{R['other']}-{COL[p]}{R['foa']}" for p in YK}, "0.000")
    sput(s, "chk2", "檢查：分部合計－合併營收", "US$B", {p: f"={COL[p]}{R['foa']}+{COL[p]}{R['rl']}-{COL[p]}{R['rev']}" for p in YK}, "0.000")
    sput(s, "chk3", "檢查：實績欄合併營收 − Q 頁（FY2024、FY2025、9M26）", "US$B", {p: f"={COL[p]}{R['rev']}-{QA('rev', p)}" for p in ("24", "25", "1H")}, "0.000")
    sput(s, "cxn", "非 AI 資本支出（營收 × cx_na）", "US$B", {**{p: f"={COL[p]}{R['rev']}*cx_na" for p in FC}, "26": f"=F{r}"})
    d = dep_row([(f"${COL[p]}${R['cxn']}", PY[p]) for p in FC], "life_na"); d["26"] = f"=F{r}"
    sput(s, "dnt", "新增非 AI 折舊（年限 life_na）", "US$B", d)
    r += 1
widths(SG)
SG.freeze_panes = "C4"

# =====================================================================================================
# Thesis：自用 AI 容量（資本支出 → MW → 折舊、營運成本、影子效益）
TH["A1"] = "Thesis — 命題模組：Meta 的 AI 算力全部自用。資本支出（指引與 M6）÷ Tokenomics 每 MW 資本支出 → 自用 AI MW（同 Alphabet C10）；MW → 新增 AI 折舊、自用 AI 增量營運成本（相對評價日在役機隊）→ FoA 營業利益；讀法 2：自用增量 MW × IF_HoldEcon（k＝1）＝影子效益（只進評價口徑）"
TH["A2"] = "期初機隊：FY2023（Assumed）、FY2024、FY2025、2026 年 9M 的 AI 資本支出（資本支出 − 營收 × cx_na）各以當年錨世代的每 MW 資本支出（IT＋自建比例 × 機房）換算；在建工程對應的 MW 於 Q4 26 轉入在役。MTIA 無揭露 → 全部按 GPU 世代（M4，不套 NonNV）。不計尚未起租租賃與第三方雲端（M4；Recon、Sens）→ 可能低估自用算力"
T0 = {}
r = 4
TH.cell(r, 1, "A. 期初機隊（單一數值，C 欄；三情境共用）").font = BOLD; r += 1
hdr(TH, r, ["項目", "單位", "數值"] + [""] * 10 + ["說明"]); r += 1
def tscal(key, label, unit, formula, fmt="#,##0.0", note=""):
    global r
    TH.cell(r, 1, label); TH.cell(r, 2, unit); TH[f"C{r}"] = formula; TH[f"C{r}"].number_format = fmt; TH.cell(r, NOTE_C, note)
    dn(key, f"Thesis!$C${r}"); T0[key] = r; r += 1
tscal("pmw_h", "每 MW 資本支出：H100 錨（IT＋自建比例 × 機房）", "US$m/MW", "=cit_h+selfb*cfac", "0.00")
tscal("pmw_b2", "每 MW 資本支出：GB200 錨", "US$m/MW", "=cit_b2+selfb*cfac", "0.00")
tscal("pmw_b3", "每 MW 資本支出：GB300 錨", "US$m/MW", "=cit_b3+selfb*cfac", "0.00")
tscal("pmw_26", "每 MW 資本支出：2026 新增組合", "US$m/MW", "=am_26*cit_b3+(1-am_26)*cit_vr+selfb*cfac", "0.00")
tscal("pmw_27", "每 MW 資本支出：2027 新增組合", "US$m/MW", "=am_27*cit_b3+(1-am_27)*cit_vr+selfb*cfac", "0.00")
tscal("pmw_28", "每 MW 資本支出：2028 起新增組合", "US$m/MW", "=(1-am_28)*cit_b3+am_28*cit_vr+selfb*cfac", "0.00")
tscal("aicx_24", "FY2024 AI 資本支出（資本支出 − 營收 × cx_na）", "US$B", f"=cx_24-{QA('rev', '24')}*cx_na", "#,##0.00")
tscal("aicx_25", "FY2025 AI 資本支出", "US$B", f"=cx_25-{QA('rev', '25')}*cx_na", "#,##0.00")
tscal("aicx_9m", "2026 年 9M AI 資本支出", "US$B", f"=capex_9m-{QA('rev', '1H')}*cx_na", "#,##0.00")
tscal("ai_sh26", "  2026 年 9M AI 占資本支出", "%", "=aicx_9m/capex_9m", "0.0%", "對照 cip_ai")
tscal("mw_pre24", "2023 年底以前 AI MW（FY2023 × AI 比例 ÷ H100 錨）", "MW", "=cx_23*ai_23/pmw_h*1000", note="只用於 FY2029 汰換")
tscal("mw_24", "FY2024 新增 AI MW（H100 錨）", "MW", "=aicx_24/pmw_h*1000")
tscal("mw_25", "FY2025 新增 AI MW（GB200 錨）", "MW", "=aicx_25/pmw_b2*1000")
tscal("mw_9m", "2026 年 9M 新增 AI MW（GB300 錨）", "MW", "=aicx_9m/pmw_b3*1000")
tscal("mw_tot0", "評價日 AI MW（支出口徑，含在建）", "MW", "=mw_pre24+mw_24+mw_25+mw_9m", note="Wattonomics 對照：本值為資本支出換算，非實際上線 MW")
tscal("mw_cip", "在建工程對應 MW（cip_0 × cip_ai ÷ GB300 錨）", "MW", "=cip_0*cip_ai/pmw_b3*1000", note="Q4 26 轉入在役")
tscal("base_h", "評價日在役 MW：H100 錨", "MW", "=mw_pre24+mw_24")
tscal("base_b2", "評價日在役 MW：GB200 錨（在建工程超過 9M 新增的部分自此扣）", "MW", "=mw_25-MAX(0,mw_cip-mw_9m)")
tscal("base_b3", "評價日在役 MW：GB300 錨", "MW", "=MAX(0,mw_9m-mw_cip)")
tscal("base_tot", "評價日在役 AI MW（基準機隊；其成本已在 FoA 利益率內）", "MW", "=base_h+base_b2+base_b3")
tscal("dpi0", "基準機隊每 MW IT 年折舊（IF_DeprIT 加權）", "US$m/MW-年", "=(base_h*dpi_h+base_b2*dpi_b2+base_b3*dpi_b3)/base_tot", "0.000")
tscal("he0", "基準機隊每 MW 持有成本（IF_HoldEcon 加權）", "US$m/MW-年", "=(base_h*he_h+base_b2*he_b2+base_b3*he_b3)/base_tot", "0.000")
tscal("opx0", "基準機隊每 MW 營運費用（IF_OpexGW 加權）", "US$m/MW-年", "=(base_h*opx_h+base_b2*opx_b2+base_b3*opx_b3)/base_tot", "0.000")
r += 1
T = {s: {} for s in "ABC"}
BK = [("h", "H100 錨"), ("b2", "GB200 錨"), ("b3", "GB300 錨"), ("vr", "VR200 錨")]
def addmix(p, b):
    y = PY[p]
    if y == 2026: return {"b3": "am_26", "vr": "(1-am_26)"}.get(b, "0")
    if y == 2027: return {"b3": "am_27", "vr": "(1-am_27)"}.get(b, "0")
    return {"vr": "am_28", "b3": "(1-am_28)"}.get(b, "0")
PMW = {"2H": "pmw_26", 27: "pmw_27"}
for s in "ABC":
    TS = T[s]; R = SR[s]
    def tput(key, label, unit, vals, fmt="#,##0.0", note=None, **kw):
        global r
        put(TH, r, label, unit, vals, fmt, note, **kw); TS[key] = r; r += 1
    TH.cell(r, 1, SCN[s][1]).font = BOLD; TH.cell(r, 1).fill = SUB; r += 1
    hdr(TH, r, MODEL_HDR); r += 1
    TH.cell(r, 1, "B. 自用 AI 容量（MW；期末）").font = BOLD; r += 1
    ADDS = r; END = r + 5
    cx27 = f"cx27_{s}"
    _v = {"2H": f"=(capex_g-capex_9m-Seg!F{R['cxn']})/pmw_26*1000", 27: f"=({cx27}-Seg!H{R['cxn']})/pmw_27*1000", 28: f"=H{ADDS}*a28", "26": f"=F{ADDS}"}
    for p in (29, 30, 31, 32):
        _v[p] = f"=$I${ADDS}+(g_tv*{COL[prevp(p)]}{END}-$I${ADDS})*MIN(1,({PY[p]}-2028)/conv_mw)"
    tput("adds", "自用 AI MW 新增（Q4 26、FY2027 由資本支出反推；FY2028＝FY2027 × a28；FY2029 起收斂）", "MW", _v,
         note="C10：Q4 26＝(2026 指引中點 − 9M 實績 − Q4 非 AI) ÷ 每 MW 資本支出；FY2027＝(情境 FY2027 資本支出 − 非 AI) ÷ 每 MW（汰換 0）")
    for b, nm in BK:
        v = {"1H": {"h": "=mw_pre24+mw_24", "b2": "=mw_25", "b3": "=mw_9m", "vr": 0}[b]}
        for p in FC: v[p] = f"={COL[prevp(p)]}{r}+{COL[p]}{ADDS}*{addmix(p, b)}"
        v["26"] = f"=F{r}"
        tput("end_" + b, f"期末 MW：{nm}", "MW", v)
    assert r == END
    tput("end", "期末自用 AI MW（在役；9M26 欄為支出口徑含在建）", "MW", {p: "=" + "+".join(f"{COL[p]}{TS['end_' + b]}" for b, _ in BK) for p in ["1H"] + FC + ["26"]})
    BASE = {"h": "base_h", "b2": "base_b2", "b3": "base_b3", "vr": "0"}
    for b, nm in BK:
        v = {"2H": f"=({BASE[b]}+F{TS['end_' + b]})/2"}
        for p in ANN: v[p] = f"=({COL[prevp(p)]}{TS['end_' + b]}+{COL[p]}{TS['end_' + b]})/2"
        tput("avg_" + b, f"平均在役 MW：{nm}", "MW", v, note="Q4 26：期初＝評價日在役（不含在建），期末含在建轉入" if b == "h" else None)
    tput("avg", "平均在役自用 AI MW", "MW", {p: "=" + "+".join(f"{COL[p]}{TS['avg_' + b]}" for b, _ in BK) for p in FC})
    tput("avg_inc", "  其中：相對評價日在役機隊的增量 MW", "MW", {p: f"={COL[p]}{TS['avg']}-base_tot" for p in FC})
    TH.cell(r, 1, "C. AI 資本支出（新增 MW × 每 MW；汰換只換 IT，VR200 錨）").font = BOLD; r += 1
    SUMADD = lambda p: "+".join(f"{addmix(p, b)}*cit_{b}" for b, _ in BK)
    tput("cx_it", "AI IT 資本支出（成長）", "US$B", {**{p: f"={COL[p]}{ADDS}*({SUMADD(p)})/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
    tput("cx_fac", "AI 機房資本支出（自建比例 × IF_CapexFacility）", "US$B", {**{p: f"={COL[p]}{ADDS}*selfb*cfac/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
    tput("retire", "汰換 MW（life_it 年前的 MW，換成 VR200 錨 IT）", "MW",
         {"2H": 0, 27: 0, 28: 0, 29: "=mw_pre24", 30: "=mw_24", 31: "=mw_25", 32: f"=mw_9m+F{ADDS}", "26": f"=F{r}"},
         note="IT 6 年：2023 年底以前 → 2029、FY2024 → 2030、FY2025 → 2031、2026 → 2032；汰換後 MW 照常在役")
    tput("cx_ref", "AI 汰換資本支出（只含 IT）", "US$B", {**{p: f"={COL[p]}{TS['retire']}*cit_vr/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
    tput("cx_ai", "AI 資本支出合計", "US$B", {p: f"={COL[p]}{TS['cx_it']}+{COL[p]}{TS['cx_fac']}+{COL[p]}{TS['cx_ref']}" for p in FC + ["26"]}, "#,##0.00", bold=True)
    tput("cx_tot", "資本支出合計（含融資租賃本金）＝AI＋非 AI（Seg）", "US$B",
         {**{p: f"={COL[p]}{TS['cx_ai']}+Seg!{COL[p]}{R['cxn']}" for p in FC}, "1H": "=capex_9m", "26": f"=E{r}+F{r}"}, "#,##0.00")
    tput("chk_c10", "檢查：FY2026 資本支出 − 2026 指引中點（C10）", "US$B", {"26": f"=G{TS['cx_tot']}-capex_g"}, "0.000")
    tput("chk_m6", "檢查：FY2027 資本支出 − 情境 FY2027 資本支出（M6）", "US$B", {27: f"=H{TS['cx_tot']}-{cx27}"}, "0.000")
    TH.cell(r, 1, "D. 新增 AI 折舊（在建工程 Q4 26 轉入；新增資本支出依年限直線攤提；當期 × phi_0、最後一年 1 − phi_0；Q4 26 欄 × 0.25）").font = BOLD; r += 1
    tput("cip_itv", "在建工程轉入：IT", "US$B", {"2H": "=cip_0*cip_it"}, "#,##0.00")
    tput("cip_facv", "在建工程轉入：機房", "US$B", {"2H": "=cip_0*(1-cip_it)"}, "#,##0.00")
    tput("v_it", "IT 批次金額（成長＋汰換）", "US$B", {p: f"={COL[p]}{TS['cx_it']}+{COL[p]}{TS['cx_ref']}" for p in FC}, "#,##0.00")
    tput("v_fac", "機房批次金額", "US$B", {p: f"={COL[p]}{TS['cx_fac']}" for p in FC}, "#,##0.00")
    it_src = [(f"$F${TS['cip_itv']}", 2026)] + [(f"${COL[p]}${TS['v_it']}", PY[p]) for p in FC]
    fac_src = [(f"$F${TS['cip_facv']}", 2026)] + [(f"${COL[p]}${TS['v_fac']}", PY[p]) for p in FC]
    d = dep_row(it_src, "life_it"); d["26"] = f"=F{r}"; tput("da_it", "新增 AI 折舊：IT（含在建工程）", "US$B", d, "#,##0.00")
    d = dep_row(fac_src, "life_fac"); d["26"] = f"=F{r}"; tput("da_fac", "新增 AI 折舊：機房（含在建工程）", "US$B", d, "#,##0.00")
    tput("da_ai", "新增 AI 折舊合計（含在建工程轉入）", "US$B", {p: f"={COL[p]}{TS['da_it']}+{COL[p]}{TS['da_fac']}" for p in FC + ["26"]}, "#,##0.00", bold=True)
    TH.cell(r, 1, "E. 自用 AI 增量營運成本（相對評價日在役機隊；基準機隊成本已在 FoA 利益率內）").font = BOLD; r += 1
    SUMOP = lambda c: "+".join(f"{c}{TS['avg_' + b]}*opx_{b}" for b, _ in BK)
    tput("int_opex", "自用 AI 增量營運費用（Tokenomics IF_OpexGW）", "US$B", {**{p: f"=(({SUMOP(COL[p])})-base_tot*opx0)/1000*{LEN[p]}" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
    tput("rentl", "租用算力租金（Sens 用；基準 0）", "US$B", {"2H": 0, 27: "=rent_l/3", 28: "=rent_l*2/3", 29: "=rent_l", 30: "=rent_l", 31: "=rent_l", 32: "=rent_l", "26": f"=F{r}"}, "#,##0.00",
         note="M4：尚未起租租賃與第三方雲端只進 Sens；讀法 2 下與影子效益同樣價值中性（DCF、分部加總皆加回）")
    tput("int_cost", "自用 AI 增量成本合計（營運＋租金＋新增 AI 折舊）", "US$B", {p: f"={COL[p]}{TS['int_opex']}+{COL[p]}{TS['rentl']}+{COL[p]}{TS['da_ai']}" for p in FC + ["26"]}, "#,##0.00", bold=True)
    TH.cell(r, 1, "F. 評價日在役資產的基礎折舊（隨期初機隊汰換遞減）").font = BOLD; r += 1
    _ret = TS["retire"]
    _cv = {"2H": 0}
    for p in ANN:
        prev_cols = [COL[q] for q in ANN if PY[q] < PY[p] and PY[q] <= 2031]
        cur = f"+mid*{COL[p]}{_ret}" if PY[p] <= 2031 else ""
        _cv[p] = "=" + ("+".join(f"{c}{_ret}" for c in prev_cols) if prev_cols else "0") + cur
    _cv["26"] = f"=F{r}"
    tput("cumret", "累計已汰換的期初機隊 MW（2025 年以前；當年以半年計）", "MW", _cv, note="2026 年批次（2032 汰換）的折舊在在建工程與新增批次內，不從基礎折舊扣")
    tput("da_b", "基礎折舊＝da_base − MIN(da_base × IT 占比, 累計汰換 MW × 每 MW IT 折舊)", "US$B",
         {**{p: f"=(da_base-MIN(da_base*cip_it,{COL[p]}{TS['cumret']}*dpi0/1000))*{LEN[p]}" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
    TH.cell(r, 1, "G. 自用 AI 影子效益（讀法 2：增量在役 MW × IF_HoldEcon，k＝1；非財報營收）").font = BOLD; r += 1
    SUMHE = lambda c: "+".join(f"{c}{TS['avg_' + b]}*he_{b}" for b, _ in BK)
    tput("shadow", "自用 AI 影子效益", "US$B", {**{p: f"=(({SUMHE(COL[p])})-base_tot*he0)/1000*{LEN[p]}" for p in FC}, "26": f"=F{r}"}, "#,##0.00",
         note="原則第 7 節：只進評價口徑（PL「營業利益（評價口徑）」、Val）；營收列與分部對帳不動")
    tput("sh_vs", "  影子效益 ÷ 自用 AI 增量成本", "x", {p: f"={COL[p]}{TS['shadow']}/{COL[p]}{TS['int_cost']}" for p in FC}, "0.00", note="接近 1＝價值中性")
    r += 1
widths(TH)
TH.freeze_panes = "C4"

# =====================================================================================================
# PL
PL["A1"] = "PL — 分部營收 → 分部營業利益 → 合併營業利益 → 業外 → 稅 → 淨利 → EPS（三情境；US$B，EPS 美元）。實績欄（FY2024A、FY2025A、9M26A）取自 Q 頁；FY2026E＝9M 實績＋Q4 模型"
PL["A2"] = ("FoA EBITDA＝(FoA 營收 − 價格情境差額) × FoA EBITDA 率（9M26，排除一次性）＋差額 × inc_m − 自用 AI 增量營運成本與租金；FoA 折舊＝基礎折舊 ×(1 − RL 占比)＋新增 AI＋新增非 AI。"
            "RL 營業損益＝指引與法說路徑。財報口徑保留 2026 年一次性項目；評價口徑排除（M5 擴大）並加讀法 2 影子效益")
PR_ = {s: {} for s in "ABC"}
r = 4
PL.cell(r, 1, "分部對帳彙總（三情境 × 全部年度取絕對值最大者；實績欄同列）").font = BOLD; r += 1
hdr(PL, r, MODEL_HDR); r += 1
SUMROW = r; r += 4
def pput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(PL, r, label, unit, vals, fmt, note, **kw); PR_[s][key] = r; r += 1
for s in "ABC":
    R = SR[s]; P = PR_[s]; TS = T[s]
    PL.cell(r, 1, SCN[s][1]).font = BOLD; PL.cell(r, 1).fill = SUB; r += 1
    hdr(PL, r, MODEL_HDR); r += 1
    ref = lambda k: {p: f"=Seg!{COL[p]}{R[k]}" for p in YK if SG[f"{COL[p]}{R[k]}"].value is not None}
    pput(s, "ads", "Advertising", "US$B", ref("ads"))
    pput(s, "other", "Other revenue", "US$B", ref("other"))
    pput(s, "foa", "Family of Apps", "US$B", ref("foa"), bold=True, fill=SUB)
    pput(s, "rl", "Reality Labs", "US$B", ref("rl"), bold=True, fill=SUB)
    pput(s, "rev", "合併營收", "US$B", ref("rev"), bold=True)
    pput(s, "g", "  合併營收年增", "%", {p: f"={COL[p]}{P['rev']}/{COL[q]}{P['rev']}-1" for p, q in (("25", "24"), ("26", "25"), (27, "26"), (28, 27), (29, 28), (30, 29), (31, 30), (32, 31))}, "0.0%")
    pput(s, "chk1", "檢查：Advertising＋Other revenue－Family of Apps", "US$B", {p: f"={COL[p]}{P['ads']}+{COL[p]}{P['other']}-{COL[p]}{P['foa']}" for p in YK}, "0.000")
    pput(s, "chk2", "檢查：分部合計－合併營收", "US$B", {p: f"={COL[p]}{P['foa']}+{COL[p]}{P['rl']}-{COL[p]}{P['rev']}" for p in YK}, "0.000")
    pput(s, "dpx", "  價格情境差額（相對 B 路徑）", "US$B", {**{p: f"=Seg!{COL[p]}{R['ads']}-Seg!{COL[p]}{R['adsB']}" for p in FC}, "26": f"=F{r}"})
    pput(s, "e_foa", "Family of Apps EBITDA（不含一次性）", "US$B",
         {p: f"=m_foa*({COL[p]}{P['foa']}-{COL[p]}{P['dpx']})+inc_m*{COL[p]}{P['dpx']}-Thesis!{COL[p]}{TS['int_opex']}-Thesis!{COL[p]}{TS['rentl']}" for p in FC})
    pput(s, "d_foa", "Family of Apps 折舊（基礎 ×(1 − RL 占比)＋新增 AI＋新增非 AI）", "US$B",
         {p: f"=Thesis!{COL[p]}{TS['da_b']}*(1-rl_da_sh)+Thesis!{COL[p]}{TS['da_ai']}+Seg!{COL[p]}{R['dnt']}" for p in FC})
    v = {p: "=" + QA("oi_foa", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['e_foa']}-{COL[p]}{P['d_foa']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_foa", "Family of Apps 營業利益", "US$B", v, bold=True, note="9M26 實績含一次性法律 12.29 與遣散 1.18（M5）")
    v = {p: "=" + QA("oi_rl", p) for p in ("24", "25", "1H")}
    v["2H"] = f"=-(rl_loss26+{QA('oi_rl', '1H')})"
    v[27] = "=-rl_loss27"
    for p in (28, 29, 30, 31, 32): v[p] = f"=-(rl_loss27+(rl_loss32-rl_loss27)*({PY[p]}-2027)/5)"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_rl", "Reality Labs 營業損益", "US$B", v, bold=True, note="Q4 26＝2026 指引中點 − 9M 實績；FY2027 約 −15（法說）；之後線性到 FY2032 rl_loss32")
    pput(s, "d_rl", "Reality Labs 折舊（基礎 × RL 占比）", "US$B", {p: f"=Thesis!{COL[p]}{TS['da_b']}*rl_da_sh" for p in FC})
    pput(s, "e_rl", "Reality Labs EBITDA", "US$B", {p: f"={COL[p]}{P['oi_rl']}+{COL[p]}{P['d_rl']}" for p in FC})
    pput(s, "oi_seg", "分部營業利益合計", "US$B", {p: f"={COL[p]}{P['oi_foa']}+{COL[p]}{P['oi_rl']}" for p in YK})
    pput(s, "ebitda", "合併 EBITDA（分部 EBITDA 合計）", "US$B", {p: f"={COL[p]}{P['e_foa']}+{COL[p]}{P['e_rl']}" for p in FC})
    v = {p: "=" + QA("da", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"=Thesis!{COL[p]}{TS['da_b']}+Thesis!{COL[p]}{TS['da_ai']}+Seg!{COL[p]}{R['dnt']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "da", "折舊與攤銷合計（基礎＋新增 AI＋新增非 AI）", "US$B", v)
    pput(s, "da_ai", "  其中：新增 AI 折舊", "US$B", {**{p: f"=Thesis!{COL[p]}{TS['da_ai']}" for p in FC}, "26": f"=F{r}"})
    v = {p: "=" + QA("oi", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['ebitda']}-{COL[p]}{P['da']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi", "合併營業利益（財報口徑）", "US$B", v, bold=True, fill=SUB, note="預測期＝合併 EBITDA − 折舊合計（與分部合計獨立計算，下一列核對）")
    pput(s, "chk4", "檢查：分部營業利益合計－合併營業利益", "US$B", {p: f"={COL[p]}{P['oi_seg']}-{COL[p]}{P['oi']}" for p in YK}, "0.000")
    pput(s, "ot", "  一次性項目（營業利益內；負值＝費用）", "US$B", {"1H": "=-ot_oi_9m", **{p: 0 for p in FC}, "26": f"=E{r}+F{r}"}, note="Q2 法律 2.40＋遣散 1.18、Q3 法律 9.89（新聞稿）；預測期 0")
    pput(s, "oix", "營業利益（排除一次性）", "US$B", {p: f"={COL[p]}{P['oi']}-{COL[p]}{P['ot']}" for p in ["1H"] + FC + ["26"]})
    pput(s, "shadow", "自用 AI 影子效益（讀法 2；非財報營收）", "US$B", {**{p: f"=Thesis!{COL[p]}{TS['shadow']}" for p in FC}, "26": f"=F{r}"})
    pput(s, "oiv", "營業利益（評價口徑：排除一次性＋讀法 2 影子效益）", "US$B",
         {**{p: f"={COL[p]}{P['oix']}+(internal_ai_reading=2)*{COL[p]}{P['shadow']}" for p in FC + ["26"]}, "1H": f"=E{P['oix']}"}, bold=True)
    pput(s, "om", "  營業利益率（財報口徑）", "%", {p: f"={COL[p]}{P['oi']}/{COL[p]}{P['rev']}" for p in YK}, "0.0%")
    pput(s, "omx", "  營業利益率（排除一次性）", "%", {p: f"={COL[p]}{P['oix']}/{COL[p]}{P['rev']}" for p in ["1H"] + FC + ["26"]}, "0.0%")
    pput(s, "om_foa", "  Family of Apps 營業利益率", "%", {p: f"={COL[p]}{P['oi_foa']}/{COL[p]}{P['foa']}" for p in YK}, "0.0%")
    P["ii"] = r; r += 1; P["ie"] = r; r += 1
    v = {p: "=" + QA("oie", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['ii']}-{COL[p]}{P['ie']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oie", "業外收益合計", "US$B", v)
    pput(s, "pt", "稅前利益", "US$B", {p: f"={COL[p]}{P['oi']}+{COL[p]}{P['oie']}" for p in YK})
    v = {p: "=" + QA("taxa", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['pt']}*tax"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "tx", "所得稅", "US$B", v, note="9M26 含 Q1 CAMT 稅務利益 −8.03")
    pput(s, "ni", "淨利（財報口徑）", "US$B", {p: f"={COL[p]}{P['pt']}-{COL[p]}{P['tx']}" for p in YK}, bold=True)
    pput(s, "chk5", "檢查：實績淨利 − Q 頁", "US$B", {p: f"={COL[p]}{P['ni']}-{QA('ni', p)}" for p in ("24", "25", "1H")}, "0.000")
    pput(s, "ot_ni", "  一次性項目稅後（法律依 8-K 排除後 EPS 隱含稅盾、遣散按 tax、稅務利益全額）", "US$B",
         {"1H": "=-(ot_legal_q2+ot_legal_q3)*ot_legal_at-ot_sev_q2*(1-tax)+ot_tax_q1", **{p: 0 for p in FC}, "26": f"=E{r}+F{r}"})
    pput(s, "pref", "特別股股利", "US$B", {"1H": 0, "2H": "=pref_2H", "26": f"=E{r}+F{r}", 27: "=pref_27", 28: "=pref_28", 29: "=pref_29", 30: 0, 31: 0, 32: 0})
    pput(s, "nic", "歸屬普通股淨利", "US$B", {p: f"={COL[p]}{P['ni']}-{COL[p]}{P['pref']}" for p in ["1H"] + FC + ["26"]})
    P["she"] = r; r += 1; P["sh"] = r; r += 1
    pput(s, "eps", "EPS（GAAP 稀釋，財報口徑，美元）", "US$", {**{p: f"={COL[p]}{P['ni']}/{COL[p]}{P['sh']}" for p in ("24", "25")},
                                                       **{p: f"={COL[p]}{P['nic']}/{COL[p]}{P['sh']}" for p in ["1H"] + FC + ["26"]}}, "0.00", bold=True)
    pput(s, "epsx", "EPS（排除一次性，美元）", "US$", {p: f"=({COL[p]}{P['nic']}-{COL[p]}{P['ot_ni']})/{COL[p]}{P['sh']}" for p in ["1H"] + FC + ["26"]}, "0.00",
         note="M5 擴大：排除 Q1 稅務利益、Q2 法律與遣散、Q3 法律")
    pput(s, "epsv", "EPS（評價口徑：排除一次性＋讀法 2 稅後影子效益，美元）", "US$",
         {p: f"=({COL[p]}{P['nic']}-{COL[p]}{P['ot_ni']}+(internal_ai_reading=2)*{COL[p]}{P['shadow']}*(1-tax))/{COL[p]}{P['sh']}" for p in FC + ["26"]}, "0.00")
    v = {p: "=" + QA("sbc", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['rev']}*sbc_r"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "sbc", "股權報酬（已含在分部營業利益內）", "US$B", v)
    r += 1
widths(PL)
PL.freeze_panes = "C4"
CHKS = [("chk1", "檢查：Advertising＋Other revenue－Family of Apps"), ("chk2", "檢查：分部合計－合併營收"), ("chk4", "檢查：分部營業利益合計－合併營業利益")]
for i, (k, lab) in enumerate(CHKS):
    rr = SUMROW + i
    PL.cell(rr, 1, lab); PL.cell(rr, 2, "US$B")
    for p in YK:
        c = COL[p]
        if PL[f"{c}{PR_['A'][k]}"].value is None: continue
        PL[f"{c}{rr}"] = "=MAX(" + ",".join(f"ABS({c}{PR_[s][k]})" for s in "ABC") + ")"
        PL[f"{c}{rr}"].number_format = "0.000"
    PL.cell(rr, NOTE_C, "三情境各區塊同名核對列的最大絕對值；必須為 0（tools/segment_check 讀本列）")

# =====================================================================================================
# CF
CF["A1"] = "CF — 現金流（US$B）。營運現金＝淨利＋折舊＋股權報酬＋營運資金 − 協議付款；資本支出（含融資租賃本金，公司口徑）＝AI（Thesis）＋非 AI（Seg）；自由現金流＝營運現金 − 資本支出。融資順序：現金 → 減少回購 → 新債（最低現金 40）"
CF["A2"] = "2026-10-29 強制轉換特別股 12.5（Q4 26 流入）；Class A 遠期出售 24.9 推定 FY2027 交割（fwd_on）；回購 2026 暫停、預設 FY2027 起 0；多州檢察長協議 12.7 分 10 年（FY2027 起每年 1.27，Assumed）"
CR = {s: {} for s in "ABC"}
r = 4
CF.cell(r, 1, "收支關卡（地圖讀這兩列；模型期間 FY2026–FY2030，FY2026 含 9M 實績）").font = BOLD; r += 1
hdr(CF, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
GATE = r; r += 5
def cput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(CF, r, label, unit, vals, fmt, note, **kw); CR[s][key] = r; r += 1
for s in "ABC":
    R = SR[s]; P = PR_[s]; C = CR[s]; TS = T[s]
    CF.cell(r, 1, SCN[s][1]).font = BOLD; CF.cell(r, 1).fill = SUB; r += 1
    hdr(CF, r, MODEL_HDR); r += 1
    cput(s, "ni", "淨利", "US$B", {p: f"=PL!{COL[p]}{P['ni']}" for p in YK})
    cput(s, "da", "折舊與攤銷", "US$B", {p: f"=PL!{COL[p]}{P['da']}" for p in YK})
    cput(s, "sbc", "股權報酬（非現金）", "US$B", {p: f"=PL!{COL[p]}{P['sbc']}" for p in YK})
    cput(s, "wc", "營運資金變動（流入＋）", "US$B", {"2H": 0, **{p: f"=wc_k*(PL!{COL[p]}{P['rev']}-PL!{COL[prevfy(p)]}{P['rev']})" for p in ANN}, "26": f"=F{r}"})
    cput(s, "ag", "多州檢察長協議付款（費用已於 2026Q3 認列）", "US$B", {"2H": 0, **{p: "=-ag_pay" for p in ANN}, "26": f"=F{r}"}, note="chat 端審查第 3 條：留在現金流與 DCF")
    v = {p: f"={QA('cfo', p)}-{COL[p]}{C['ni']}-{COL[p]}{C['da']}-{COL[p]}{C['sbc']}" for p in ("24", "25", "1H")}
    for p in FC: v[p] = 0
    v["26"] = f"=E{r}+F{r}"
    cput(s, "oth", "其他非現金調整（實績差額：應計法律費用、遞延稅、營運資金等）", "US$B", v)
    cput(s, "cfo", "營運現金流", "US$B", {p: "=" + "+".join(f"N({COL[p]}{C[k]})" for k in ("ni", "da", "sbc", "wc", "ag", "oth")) for p in YK}, bold=True)
    cput(s, "cx_ai", "資本支出：AI（Thesis）", "US$B", {**{p: f"=Thesis!{COL[p]}{TS['cx_ai']}" for p in FC}, "26": f"=F{r}"})
    cput(s, "cx_na", "資本支出：非 AI（Seg）", "US$B", {**{p: f"=Seg!{COL[p]}{R['cxn']}" for p in FC}, "26": f"=F{r}"})
    v = {p: "=" + QA("capex", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{C['cx_ai']}+{COL[p]}{C['cx_na']}"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "cx", "資本支出合計（含融資租賃本金）", "US$B", v, bold=True)
    cput(s, "fcf", "自由現金流（營運現金 − 資本支出；公司定義）", "US$B", {p: f"={COL[p]}{C['cfo']}-{COL[p]}{C['cx']}" for p in YK}, bold=True, fill=SUB)
    cput(s, "chkf", "檢查：實績自由現金流 − Q 頁", "US$B", {p: f"={COL[p]}{C['fcf']}-{QA('fcf', p)}" for p in ("24", "25", "1H")}, "0.000")
    v = {p: "=" + QA("div", p) for p in ("24", "25", "1H")}
    v["2H"] = "=dps/4*sh_out"
    for p in ANN: v[p] = f"=dps*(1+dps_g)^({PY[p]}-2026)*PL!{COL[prevp(p)]}{P['she']}"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "div", "普通股股利", "US$B", v, note="Q4 26＝每股年股利 ÷ 4 × 流通股；之後每股 +5%／年 × 期初稀釋股數")
    cput(s, "pref", "特別股股利", "US$B", {p: f"=PL!{COL[p]}{P['pref']}" for p in ["1H"] + FC + ["26"]})
    cput(s, "fcfd", "自由現金流扣股利", "US$B", {p: f"={COL[p]}{C['fcf']}-{COL[p]}{C['div']}-N({COL[p]}{C['pref']})" for p in YK}, bold=True)
    cput(s, "mat", "債務到期償還", "US$B", {"1H": 0, "2H": 0, 27: "=mat_27", **{p: 0 for p in (28, 29, 30, 31, 32)}, "26": f"=E{r}+F{r}"}, note="9M26 無償還；其他年度到期即再融資（不建模）")
    cput(s, "newd", "已發行新債", "US$B", {"1H": fv('cf.debt_proceeds.ytd9m2026') / 1000, "2H": 0, **{p: 0 for p in ANN}, "26": f"=E{r}+F{r}"}, note="9M26：發債淨額 24.910（新聞稿）")
    cput(s, "eq", "股權發行（Q4 26 強制轉換特別股；FY2027 遠期出售）", "US$B", {"1H": 0, "2H": "=mcps", 27: "=fwd_cash*fwd_on", **{p: 0 for p in (28, 29, 30, 31, 32)}, "26": f"=E{r}+F{r}"})
    cput(s, "copen", "期初現金（現金＋有價證券）", "US$B", {"1H": "=cash_ye25", **{p: f"={COL[prevp(p)]}{r + 5}" for p in FC}})
    C["oth2"] = r; r += 1
    cput(s, "pre", "回購前現金", "US$B", {p: f"={COL[p]}{C['copen']}+{COL[p]}{C['fcf']}-{COL[p]}{C['div']}-{COL[p]}{C['pref']}-{COL[p]}{C['mat']}+{COL[p]}{C['newd']}+{COL[p]}{C['eq']}+{COL[p]}{C['oth2']}" for p in ["1H"] + FC})
    v = {"1H": "=" + QA("bb", "1H"), "2H": 0}
    for p in ANN: v[p] = f"=MAX(0,MIN(bb_t,{COL[p]}{C['pre']}-min_cash))"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "bb", "回購", "US$B", v, note="Q4 26＝0（暫停）；之後＝MIN(回購目標, 回購前現金 − 最低現金)")
    cput(s, "plug", "新增舉債（現金低於最低現金時）", "US$B", {"1H": 0, **{p: f"=MAX(0,min_cash-({COL[p]}{C['pre']}-{COL[p]}{C['bb']}))" for p in FC}, "26": f"=E{r}+F{r}"})
    cput(s, "cend", "期末現金（現金＋有價證券）", "US$B", {p: f"={COL[p]}{C['pre']}-{COL[p]}{C['bb']}+{COL[p]}{C['plug']}" for p in ["1H"] + FC})
    assert C["cend"] == C["copen"] + 5
    put(CF, C["oth2"], "其他（9M26 實績差額：投資活動其他、代扣稅、匯率；預測期 0）", "US$B",
        {"1H": f"=cash_0-(E{C['copen']}+E{C['fcf']}-E{C['div']}-E{C['pref']}-E{C['mat']}+E{C['newd']}+E{C['eq']})", **{p: 0 for p in FC}})
    cput(s, "chk", "檢查：9M26 期末現金 − 2026-09-30（應為 0）", "US$B", {"1H": f"=E{C['cend']}-cash_0"}, "0.000")
    cput(s, "dopen", "期初債務", "US$B", {"1H": "=debt_ye25", **{p: f"={COL[prevp(p)]}{r + 1}" for p in FC}})
    cput(s, "dend", "期末債務", "US$B", {"1H": "=debt_0", **{p: f"={COL[p]}{C['dopen']}-{COL[p]}{C['mat']}+{COL[p]}{C['newd']}+{COL[p]}{C['plug']}" for p in FC}}, note="9M26＝票據面額 84.0（10-Q）")
    cput(s, "nc", "淨現金（現金 − 債務）", "US$B", {p: f"={COL[p]}{C['cend']}-{COL[p]}{C['dend']}" for p in ["1H"] + FC})
    cput(s, "bpx", "回購均價（評價日股價 × (1＋ke)^年數）", "US$", {**{p: f"=px_val*(1+ke)^({PY[p]}+mid-t_val)" for p in ANN}, "2H": "=px_val"}, "0.00")
    r += 1
widths(CF)
CF.freeze_panes = "C4"
CF.cell(GATE, 1, "收支關卡：自由現金流扣股利（五年合計）")
CF.cell(GATE + 1, 1, "  其中：五年新增舉債")
CF.cell(GATE + 2, 1, "  其中：五年回購合計")
CF.cell(GATE + 3, 1, "是否需要舉債或停止回購")
for j, s in enumerate("ABC"):
    c = "BCD"[j]; C = CR[s]
    CF[f"{c}{GATE}"] = f"=SUM(G{C['fcfd']}:K{C['fcfd']})"; CF[f"{c}{GATE}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 1}"] = f"=SUM(G{C['plug']}:K{C['plug']})"; CF[f"{c}{GATE + 1}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 2}"] = f"=SUM(G{C['bb']}:K{C['bb']})"; CF[f"{c}{GATE + 2}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 3}"] = (f'=IF({c}{GATE + 1}>0.001,"需要舉債"&IF({c}{GATE + 2}<4*bb_t-0.001,"且減少回購",""),'
                            f'IF({c}{GATE + 2}<4*bb_t-0.001,"不需舉債，但須減少回購","不需舉債，回購照目標"))')
CF.cell(GATE + 4, 1, "檢查：收支關卡 − 首年起五年逐欄加總（FY2026＋…＋FY2030）")
for j, s_ in enumerate("ABC"):
    c = "BCD"[j]; C_ = CR[s_]
    CF[f"{c}{GATE + 4}"] = f"={c}{GATE}-(G{C_['fcfd']}+H{C_['fcfd']}+I{C_['fcfd']}+J{C_['fcfd']}+K{C_['fcfd']})"; CF[f"{c}{GATE + 4}"].number_format = "0.000"
CF.cell(GATE, 5, "FY2026–FY2030 自由現金流（公司定義，已扣融資租賃本金與協議付款）− 普通股與特別股股利；FY2026 含 9M 實績。負值＝入不敷出")
CF.cell(GATE + 3, 5, "依現金瀑布：現金 → 減少回購 → 新債（最低現金 40）；強制轉換特別股與遠期出售不計入關卡（融資）")
for s in "ABC":
    P = PR_[s]; C = CR[s]
    put(PL, P["ii"], "利息收入（期初現金 × r_cash）", "US$B", {**{p: f"=r_cash*CF!{COL[p]}{C['copen']}*{LEN[p]}" for p in FC}, "26": f"=F{P['ii']}"})
    put(PL, P["ie"], "利息費用（期初債務 × r_debt）", "US$B", {**{p: f"=r_debt*CF!{COL[p]}{C['dopen']}*{LEN[p]}" for p in FC}, "26": f"=F{P['ie']}"})
    v = {"1H": "=sh_dil_q3+(sh_out-sh_basic_q3)"}
    for p in FC:
        q = COL[prevp(p)]
        v[p] = f"={q}{P['she']}*(1+dil*{LEN[p]})-CF!{COL[p]}{C['bb']}/CF!{COL[p]}{C['bpx']}" + ("+fwd_sh*fwd_on" if p == 27 else "") + ("+conv_sh" if p == 29 else "")
    put(PL, P["she"], "期末稀釋股數（十億股）", "十億股", v, "0.000", "9M26 期末＝Q3 稀釋加權＋(10-16 流通股 − Q3 基本加權)；之後 ×(1＋dil × 期間長度) − 回購 ÷ 回購均價；FY2027 遠期出售 40M（fwd_on）；2029 強制轉換特別股 +conv_sh")
    v = {"24": "=" + QA("sh", "24"), "25": "=" + QA("sh", "25"), "1H": "=" + QA("sh", "1H"), "26": f"=0.75*E{P['sh']}+0.25*F{P['sh']}"}
    for p in FC: v[p] = f"=({COL[prevp(p)]}{P['she']}+{COL[p]}{P['she']})/2"
    v["2H"] = f"=E{P['she']}*(1+dil*0.125)"
    put(PL, P["sh"], "平均稀釋股數（十億股）", "十億股", v, "0.000")

# =====================================================================================================
# Val
V["A1"] = "Val — 目標價（約 12 個月後，2027-10-30）＝DCF 50%＋分部加總（SOTP）50%；情境權重 A 25%／B 50%／C 25%。主值＝自用 AI 讀法 2，讀法 1 並列。每股為 Class C 股票股利（2027-01，1:1）前的口徑；股票股利後每股數字約 ÷ 2"
V["A2"] = ("DCF：無槓桿自由現金流（NOPAT＋折舊 − 資本支出＋營運資金 − 協議付款；股權報酬不加回）以 WACC 折現到評價日 2026-09-30（期中慣例），終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)，RONIC＝WACC；"
           "以 ke 推到目標時點並扣期間股利。SOTP：目標時點 NTM EBITDA（股權報酬加回）× 數位廣告同業 NTM EV/EBITDA；讀法 2 加回自用 AI 增量營運成本並按帳面加自用 AI 新增資產")
VR_ = {}
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}
r = 4
COLS3 = "BCD"
def vput(key, label, vals, fmt="#,##0.00", note=None, bold=False, fill=None):
    global r
    V.cell(r, 1, label)
    if bold: V.cell(r, 1).font = BOLD
    for j, v in enumerate(vals):
        if v is None: continue
        c = V.cell(r, 2 + j, v); c.number_format = fmt[j] if isinstance(fmt, (list, tuple)) else fmt
    if note: V.cell(r, 5, note)
    if fill:
        for cc in range(1, 6): V.cell(r, cc).fill = fill
    if key: VR_[key] = r
    r += 1
S3 = lambda f: [f(s) for s in "ABC"]
X3 = lambda s: COLS3["ABC".index(s)]
V.cell(r, 1, "0.3 結論").font = BOLD; r += 1
hdr(V, r, ["項目", "數值", "對照", "對照", "說明"]); r += 1
CONC = r; r += 26
V.cell(r, 1, "0.0 折現率（CAPM；Blume β、約 10 年期 kd）").font = BOLD; r += 1
hdr(V, r, ["項目", "數值", "", "", "說明"]); r += 1
vput("emkt", "股權市值 E＝評價日股價 × 流通股", ["=px_val*sh_out"], "#,##0.0", note="評價日 2026-09-30 收盤 725.18")
vput("dmkt", "負債 D（票據面額＋融資租賃）", ["=debt_0+fin_lease"], "#,##0.0")
vput("ke", "股東權益成本 ke＝rf＋β × ERP", ["=rf+beta*erp"], "0.00%", note='="β（Blume 調整、兩來源中位數）"&TEXT(beta,"0.000")')
dn("ke", f"Val!$B${VR_['ke']}")
vput("kdat", "稅後債務成本", ["=kd*(1-tax)"], "0.00%")
vput("wacc", "WACC（＋wacc_adj，Sens 用）", [f"=(B{VR_['emkt']}*ke+B{VR_['dmkt']}*B{VR_['kdat']})/(B{VR_['emkt']}+B{VR_['dmkt']})+wacc_adj"], "0.00%", bold=True)
dn("wacc", f"Val!$B${VR_['wacc']}")
vput("ronic", "永續期 RONIC＝WACC＋ronic_sp", ["=wacc+ronic_sp"], "0.00%")
dn("ronic", f"Val!$B${VR_['ronic']}")
r += 1
V.cell(r, 1, "0.1 DCF（無槓桿自由現金流；三情境各一區塊；欄 F＝Q4 26、H–M＝FY2027–FY2032）").font = BOLD; r += 1
DC = {s: {} for s in "ABC"}
VC = {"2H": "F", 27: "H", 28: "I", 29: "J", 30: "K", 31: "L", 32: "M"}
for s in "ABC":
    P = PR_[s]; C = CR[s]; D = DC[s]; TS = T[s]
    V.cell(r, 1, SCN[s][1]).font = BOLD; V.cell(r, 1).fill = SUB; r += 1
    hdr(V, r, ["項目", "", "", "", "說明", YH["2H"], "", YH[27], YH[28], YH[29], YH[30], YH[31], YH[32]]); r += 1
    def drow(key, label, f, fmt="#,##0.00", note=None):
        global r
        V.cell(r, 1, label)
        for p in FC:
            c = V[f"{VC[p]}{r}"]; c.value = f(p); c.number_format = fmt
        if note: V.cell(r, 5, note)
        D[key] = r; r += 1
    drow("ebit", "營業利益（財報口徑 EBIT；預測期無一次性）", lambda p: f"=PL!{COL[p]}{P['oi']}")
    drow("nopat", "NOPAT＝(EBIT＋sbc_add × 股權報酬) ×(1 − 稅率)", lambda p: f"=({VC[p]}{D['ebit']}+sbc_add*PL!{COL[p]}{P['sbc']})*(1-tax)")
    drow("da", "＋折舊", lambda p: f"=PL!{COL[p]}{P['da']}")
    drow("cx", "− 資本支出（含融資租賃本金）", lambda p: f"=CF!{COL[p]}{C['cx']}")
    drow("wc", "＋營運資金變動", lambda p: f"=CF!{COL[p]}{C['wc']}")
    drow("ag", "＋協議付款（負值；不可抵稅）", lambda p: f"=CF!{COL[p]}{C['ag']}")
    drow("ufcf", "無槓桿自由現金流（財報口徑）", lambda p: f"={VC[p]}{D['nopat']}+{VC[p]}{D['da']}-{VC[p]}{D['cx']}+{VC[p]}{D['wc']}+{VC[p]}{D['ag']}")
    drow("shd", "自用 AI 影子效益＋租用算力租金加回（稅後；讀法 2 加入）", lambda p: f"=(PL!{COL[p]}{P['shadow']}+Thesis!{COL[p]}{TS['rentl']})*(1-tax)")
    drow("t", "折現年數（評價日 2026-09-30 起，期中）", lambda p: "=t_q4" if p == "2H" else f"={PY[p]}+mid-t_val", "0.00")
    drow("df", "折現因子 1 ÷ (1＋WACC)^年數", lambda p: f"=1/(1+wacc)^{VC[p]}{D['t']}", "0.0000")
    drow("pv", "現值（財報口徑）", lambda p: f"={VC[p]}{D['ufcf']}*{VC[p]}{D['df']}")
    drow("pvs", "現值（影子效益）", lambda p: f"={VC[p]}{D['shd']}*{VC[p]}{D['df']}")
    r += 1
V.cell(r, 1, "DCF 彙總").font = BOLD; r += 1
hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
TV_T = lambda s: f"($M${DC[s]['t']}+mid)"
vput("pvx", "明確期現值合計（財報口徑）", S3(lambda s: f"=SUM(F{DC[s]['pv']}:M{DC[s]['pv']})"))
vput("tvv", "終值（FY2032 年底，財報口徑）", S3(lambda s: f"=IF(tv_alt=1,(M{DC[s]['ufcf']}-M{DC[s]['ag']})*(1+g_tv)/(wacc-g_tv),M{DC[s]['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))"),
     note="價值驅動：NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；協議付款 FY2033 以後列類債")
vput("pvtv", "終值現值", S3(lambda s: f"={X3(s)}{VR_['tvv']}/(1+wacc)^{TV_T(s)}"))
vput("ev1", "企業價值 EV（讀法 1：只計自用 AI 成本）", S3(lambda s: f"={X3(s)}{VR_['pvx']}+{X3(s)}{VR_['pvtv']}"))
vput("pvsx", "＋自用 AI 影子效益現值：明確期", S3(lambda s: f"=SUM(F{DC[s]['pvs']}:M{DC[s]['pvs']})"))
vput("pvstv", "＋自用 AI 影子效益現值：終值", S3(lambda s: f"=IF(tv_alt=1,M{DC[s]['shd']}*(1+g_tv)/(wacc-g_tv),M{DC[s]['shd']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))/(1+wacc)^{TV_T(s)}"))
vput("ev2", "企業價值 EV（讀法 2：自用 AI 價值中性）", S3(lambda s: f"={X3(s)}{VR_['ev1']}+{X3(s)}{VR_['pvsx']}+{X3(s)}{VR_['pvstv']}"))
vput("ev", "企業價值 EV（主值，依 internal_ai_reading）", S3(lambda s: f"=IF(internal_ai_reading=2,{X3(s)}{VR_['ev2']},{X3(s)}{VR_['ev1']})"), bold=True)
vput("tvsh", "  其中：終值占 EV（主值）", S3(lambda s: f"=({X3(s)}{VR_['pvtv']}+(internal_ai_reading=2)*{X3(s)}{VR_['pvstv']})/{X3(s)}{VR_['ev']}"), "0.0%")
vput("nc0", "＋淨現金（評價日：現金 − 票據面額 − 融資租賃）", ["=cash_0-debt_0-fin_lease"] * 3)
_tt = DC["A"]["t"]
vput("post", "＋評價日後股權融資（強制轉換特別股淨額 12.46＋遠期出售 × fwd_on，以 ke 自 FY2027 期中折現）", [f"=mcps+fwd_cash*fwd_on/(1+ke)^$H${_tt}"] * 3,
     note="v0.1.1：遠期出售現金於 FY2027 交割（期中，距評價日 0.75 年），以 ke 折現到評價日（同特別股股利）；遠期價每日按隔夜利率減利差增值、按股利調降，未計（保守）")
vput("agtv", "− 協議付款 FY2033–FY2036（4 期，以 WACC 折現到評價日）", ["=ag_pay*((1+wacc)^-6.75+(1+wacc)^-7.75+(1+wacc)^-8.75+(1+wacc)^-9.75)"] * 3, note="FY2027–FY2032 已在無槓桿自由現金流內")
vput("dl", "− 類債（其他法律應計＋或有金額 contg）", ["=legal_oth+contg"] * 3)
vput("ppref", "− 特別股股利現值（轉換前，以 ke 折現）", [f"=pref_2H/(1+ke)^$F${_tt}+pref_27/(1+ke)^$H${_tt}+pref_28/(1+ke)^$I${_tt}+pref_29/(1+ke)^$J${_tt}"] * 3)
vput("shv", "股數（評價日期末稀釋＋遠期出售 × fwd_on＋強制轉換特別股）", S3(lambda s: f"=PL!E{PR_[s]['she']}+fwd_sh*fwd_on+conv_sh"), "0.000")
BR = lambda ev: (lambda s: "=MAX(0,({0}{1}+{0}{2}+{0}{3}-{0}{4}-{0}{5}-{0}{6})/{0}{7}*(1+ke)^t_tgt-dps_tgt)".format(X3(s), VR_[ev], VR_["nc0"], VR_["post"], VR_["agtv"], VR_["dl"], VR_["ppref"], VR_["shv"]))
vput("dcf1", "DCF 每股價值推到目標時點（讀法 1）", S3(BR("ev1")), note="(EV＋淨現金＋期後融資 − 協議尾款 − 類債 − 特別股股利) ÷ 股數 ×(1＋ke)^t_tgt − 期間每股股利")
vput("dcf2", "DCF 每股價值推到目標時點（讀法 2）", S3(BR("ev2")))
vput("dcf12", "DCF 情境價值推到 12 個月後（主值）", S3(lambda s: f"=IF(internal_ai_reading=2,{X3(s)}{VR_['dcf2']},{X3(s)}{VR_['dcf1']})"), bold=True, fill=KEY)
r += 1
V.cell(r, 1, "0.2 分部加總（SOTP；目標時點 NTM EBITDA，股權報酬加回 × 數位廣告同業 NTM EV/EBITDA；＝本模型的 Comparable）").font = BOLD; r += 1
hdr(V, r, ["同業組／倍數", "倍數", "家數", "", "說明"]); r += 1
pkeys = [f"pm_ads_{p['ticker'].lower()}" for p in PEER["peers"]]
vput("mult_ads", "數位廣告 NTM EV/EBITDA 中位數（GOOGL、PINS、TTD、APP、RDDT、SNAP）", ["=MEDIAN(" + ",".join(pkeys) + ")", len(pkeys)], ["0.00", "0"], note="data/peers_ads_20261030.json：五家 2026-10-07 基準（未更新）＋GOOGL（同基準日）")
dn("mult_ads", f"Val!$B${VR_['mult_ads']}")
vput("mult_x", "FoA 與 RL 採用倍數（ads_mode：0 中位數／1 GOOGL／2 Meta 自身）", ["=CHOOSE(ads_mode+1,mult_ads,pm_ads_googl,pm_meta_own)"], "0.00")
dn("mult_x", f"Val!$B${VR_['mult_x']}")
r += 1
hdr(V, r, ["分部（NTM EBITDA，SBC 加回）", "情境 A", "情境 B", "情境 C", "倍數與說明", "EV：A", "EV：B", "EV：C"]); r += 1
NT = lambda s, k: f"(w_ntm27*PL!H{PR_[s][k]}+(1-w_ntm27)*PL!I{PR_[s][k]})"
NTT = lambda s, k: f"(w_ntm27*Thesis!H{T[s][k]}+(1-w_ntm27)*Thesis!I{T[s][k]})"
SOTP0 = r
V.cell(r, 1, "Family of Apps（EBITDA＋股權報酬）")
for j, s in enumerate("ABC"):
    V.cell(r, 2 + j, f"={NT(s, 'e_foa')}+sbc_r*{NT(s, 'rev')}").number_format = "#,##0.00"
    V.cell(r, 6 + j, f"={COLS3[j]}{r}*mult_x").number_format = "#,##0.0"
V.cell(r, 5, '="數位廣告 "&TEXT(mult_x,"0.00")&"×；股權報酬全數歸 FoA"'); VR_["s_foa"] = r; r += 1
V.cell(r, 1, "Reality Labs（EBITDA，負值；rl_mode＝1 時以 0 計）")
for j, s in enumerate("ABC"):
    V.cell(r, 2 + j, f"={NT(s, 'e_rl')}").number_format = "#,##0.00"
    V.cell(r, 6 + j, f"=IF(rl_mode=1,0,{COLS3[j]}{r}*mult_x)").number_format = "#,##0.0"
V.cell(r, 5, "同 FoA 倍數（判斷：同 Alphabet 層級負 EBITDA 處理）"); VR_["s_rl"] = r; r += 1
SOTP1 = r - 1
vput("sev1", "分部加總 EV（讀法 1）", [f"=SUM(F{SOTP0}:F{SOTP1})", f"=SUM(G{SOTP0}:G{SOTP1})", f"=SUM(H{SOTP0}:H{SOTP1})"], "#,##0.0")
V.cell(r, 1, "＋自用 AI 增量營運成本與租金加回（讀法 2；價值中性）")
for j, s in enumerate("ABC"):
    V.cell(r, 2 + j, f"={NTT(s, 'int_opex')}+{NTT(s, 'rentl')}").number_format = "#,##0.00"
    V.cell(r, 6 + j, f"={COLS3[j]}{r}*mult_x").number_format = "#,##0.0"
V.cell(r, 5, "同 Alphabet 修正二（自用 AI 的資本成本不在 EBITDA，不加影子效益倍數）"); VR_["s_sh"] = r; r += 1
V.cell(r, 1, "＋自用 AI 新增資產帳面淨值（FY2027 年底；在建工程 AI 部分＋Q4 26、FY2027 AI 資本支出 − 期間新增 AI 折舊）")
for j, s in enumerate("ABC"):
    TS = T[s]
    V.cell(r, 6 + j, f"=cip_0*cip_ai+Thesis!F{TS['cx_ai']}+Thesis!H{TS['cx_ai']}-(Thesis!F{TS['da_ai']}+Thesis!H{TS['da_ai']})*(cip_0*cip_ai+Thesis!F{TS['cx_ai']}+Thesis!H{TS['cx_ai']})/(cip_0+Thesis!F{TS['cx_ai']}+Thesis!H{TS['cx_ai']})").number_format = "#,##0.0"
V.cell(r, 5, "同 Alphabet 修正三：FY2027 年底淨現金已扣這些支出，讀法 2 下按帳面計；折舊按 AI 比例扣"); VR_["s_nbv"] = r; r += 1
vput("sev2", "分部加總 EV（讀法 2）", [f"=B{VR_['sev1']}+F{VR_['s_sh']}+F{VR_['s_nbv']}", f"=C{VR_['sev1']}+G{VR_['s_sh']}+G{VR_['s_nbv']}", f"=D{VR_['sev1']}+H{VR_['s_sh']}+H{VR_['s_nbv']}"], "#,##0.0")
vput("sev", "分部加總 EV（主值）", S3(lambda s: f"=IF(internal_ai_reading=2,{X3(s)}{VR_['sev2']},{X3(s)}{VR_['sev1']})"), "#,##0.0", bold=True)
vput("sebt", "  分部 NTM EBITDA 合計（主值）", S3(lambda s: f"=SUM({X3(s)}{SOTP0}:{X3(s)}{SOTP1})+(internal_ai_reading=2)*{X3(s)}{VR_['s_sh']}"), "#,##0.0")
vput("snc", "＋淨現金（FY2027 年底，CF 頁；減融資租賃）", S3(lambda s: f"=CF!H{CR[s]['nc']}-fin_lease"), note="已含強制轉換特別股、遠期出售（fwd_on）與 FY2027 協議付款")
vput("sag", "− 協議付款 FY2028–FY2036（9 期，以 WACC 折現到 FY2027 年底）", ["=ag_pay*(1-(1+wacc)^-9)/wacc"] * 3)
vput("sdl", "− 類債（其他法律應計＋或有金額）", [f"=B{VR_['dl']}"] * 3)
vput("ssh", "股數（FY2027 年底稀釋＋強制轉換特別股）", S3(lambda s: f"=PL!H{PR_[s]['she']}+conv_sh"), "0.000")
SB2 = lambda ev: (lambda s: "=MAX(0,({0}{1}+{0}{2}-{0}{3}-{0}{4})/{0}{5})".format(X3(s), VR_[ev], VR_["snc"], VR_["sag"], VR_["sdl"], VR_["ssh"]))
vput("sotp1", "分部加總每股價值（讀法 1）", S3(SB2("sev1")))
vput("sotp2", "分部加總每股價值（讀法 2）", S3(SB2("sev2")))
vput("sotp", "分部加總每股價值（主值；目標時點）", S3(lambda s: f"=IF(internal_ai_reading=2,{X3(s)}{VR_['sotp2']},{X3(s)}{VR_['sotp1']})"), bold=True, fill=KEY)
vput("simp", "  隱含 EV ÷ NTM EBITDA（主值）", S3(lambda s: "={0}{1}/{0}{2}".format(X3(s), VR_["sev"], VR_["sebt"])), "0.00")
r_end = r; r = CONC
W3 = lambda row: f"tp_wA*B{row}+tp_wB*C{row}+tp_wC*D{row}"
vput("tpdcf", "DCF 目標價（情境加權）", [f"={W3(VR_['dcf12'])}"], "#,##0.00", fill=KEY)
vput("tpcomp", "Comparable 目標價（市場主流同業法）", [f"={W3(VR_['sotp'])}"], "#,##0.00", note="＝分部加總（SOTP）情境加權")
vput("tp", "目標價（DCF 與 Comparable 加權）", ["=w_dcf*B{0}+(1-w_dcf)*B{1}".format(r - 2, r - 1)], "#,##0.00", bold=True, fill=KEY,
     note='="DCF 權重 "&TEXT(w_dcf,"0%")&"、分部加總 "&TEXT(1-w_dcf,"0%")&"；自用 AI 讀法 "&internal_ai_reading')
vput("px", "現價（2026-10-28 收盤）", ["=px"], "#,##0.00", note="StockScan、Digrin、TipRanks（In 頁 px）；10-29 收盤只有 ChartRow 一個來源，未採用")
vput("cons", "市場共識目標價", ["=cons_tp"], "#,##0.00", note="MarketBeat 58 位平均（2026-10-22）；只作對照")
vput("up", "上檔空間：加權目標價", [f"=B{VR_['tp']}/px-1"], "0.0%", bold=True, fill=KEY)
vput("upc", "上檔空間：Comparable 目標價", [f"=B{VR_['tpcomp']}/px-1"], "0.0%")
def tier(u): return f'IF({u}>=0.5,"大於 50%",IF({u}>=0.25,"25%–50%",IF({u}>=0,"0%–25%",IF({u}>-0.25,"0%～−25%",IF({u}>-0.5,"−25%～−50%","小於 −50%")))))'
def recf(u): return f'IF({u}>=0.5,"強力買進",IF({u}>=0.25,"買進",IF({u}>-0.25,"中立",IF({u}>-0.5,"放空","強力放空"))))'
U = f"B{VR_['up']}"
vput("tier", "上檔空間級距", ["=" + tier(U)], "@", note="以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
vput("rec", "投資建議", ["=" + recf(U)], "@", bold=True, fill=KEY, note="五級固定規則：≥ +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、≤ −50% 強力放空")
vput("band", "價格帶（DCF 不如預期～順利）", [f"=D{VR_['dcf12']}", f"=B{VR_['dcf12']}"], "#,##0.00", note="B＝C 情境 DCF、C＝A 情境 DCF")
vput("pess", "悲觀情境價（C 情境 DCF）", [f"=D{VR_['dcf12']}"], "#,##0.00", note="地圖「悲觀情境空間」＝本列 ÷ 現價 − 1")
vput("scen", "情境目標價（A／B／C；DCF 與分部加總加權）", S3(lambda s: "=w_dcf*{0}{1}+(1-w_dcf)*{0}{2}".format(X3(s), VR_["dcf12"], VR_["sotp"])), "#,##0.00")
vput("scenup", "情境上檔空間（A／B／C）", S3(lambda s: "={0}{1}/px-1".format(X3(s), VR_["scen"])), "0.0%")
vput("tvsh2", "DCF 終值占 EV（A／B／C）", S3(lambda s: "={0}{1}".format(X3(s), VR_["tvsh"])), "0.0%")
TPR = lambda d, sp: f"=w_dcf*({W3(VR_[d])})+(1-w_dcf)*({W3(VR_[sp])})"
vput("tp_r2", "讀法 2（自用 AI 價值中性）：目標價／上檔空間／投資建議", [TPR("dcf2", "sotp2"), f"=B{r}/px-1", "=" + recf(f"C{r}")], ["#,##0.00", "0.0%", "@"], note="活公式（主值）")
vput("tp_r1", "讀法 1（自用 AI 只計成本）：目標價／上檔空間／投資建議", [TPR("dcf1", "sotp1"), f"=B{r}/px-1", "=" + recf(f"C{r}")], ["#,##0.00", "0.0%", "@"], note="活公式（下緣敏感度）")
vput("tp_dcf12", "讀法 1／2 的 DCF 目標價（情境加權）", [f"={W3(VR_['dcf1'])}", f"={W3(VR_['dcf2'])}"], "#,##0.00")
vput("tp_sotp12", "讀法 1／2 的分部加總目標價（情境加權）", [f"={W3(VR_['sotp1'])}", f"={W3(VR_['sotp2'])}"], "#,##0.00")
vput("meth", "目標價方法", ["DCF（無槓桿 FCF、WACC、RONIC＝WACC）50%＋分部加總（數位廣告同業 NTM EV/EBITDA）50%；情境 25／50／25；自用 AI 讀法 2"], "@")
vput("wsum", "檢查：情境權重合計（應為 1）", ["=ROUND(tp_wA+tp_wB+tp_wC,6)"], "0.000")
vput("chk_rd", "檢查：主值目標價 − 所選讀法的目標價（應為 0）", [f"=B{VR_['tp']}-IF(internal_ai_reading=2,B{VR_['tp_r2']},B{VR_['tp_r1']})"], "0.000")
assert r <= CONC + 26, r
r = r_end
V.column_dimensions["A"].width = 62
for c in "BCD": V.column_dimensions[c].width = 13
V.column_dimensions["E"].width = 60
for c in "FGHIJKLM": V.column_dimensions[c].width = 12

# =====================================================================================================
# Recon
RC["A1"] = "Recon — 口徑對照：合併層對市場共識（MarketScreener，頁面 2026-10-08，Q3 財報前）、公司指引、M4 未入基準的算力（尚未起租租賃、第三方雲端）、一次性項目兩種口徑。只作對照，不作校準"
r = 3
RC.cell(r, 1, "1. 合併層：本模型 vs 市場共識").font = BOLD; r += 1
hdr(RC, r, ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B ÷ 共識 − 1", "說明"]); r += 1
RCR = {}
for k, lab, src in (("revenue", "營收", "rev"), ("ebit", "營業利益（財報口徑）", "oi"), ("eps", "EPS（GAAP）", "eps"), ("capex", "資本支出（本模型含融資租賃本金）", "cx"), ("fcf", "自由現金流", "fcf"), ("ebitda", "EBITDA（本模型＝財報營業利益＋折舊＋SBC）", "ebitda")):
    for y in (26, 27, 28):
        RC.cell(r, 1, lab); RC.cell(r, 2, f"FY20{y}")
        for j, s in enumerate("ABC"):
            col = COL["26" if y == 26 else y]
            if src in ("rev", "oi", "eps"): ref = f"PL!{col}{PR_[s][src]}"
            elif src == "ebitda": ref = f"PL!{col}{PR_[s]['oi']}+PL!{col}{PR_[s]['da']}+PL!{col}{PR_[s]['sbc']}"
            else: ref = f"CF!{col}{CR[s][src]}"
            RC.cell(r, 3 + j, f"={ref}").number_format = "#,##0.00"
        RC.cell(r, 6, f"=c_{k}_{y}").number_format = "#,##0.00"
        RC.cell(r, 7, f"=D{r}/F{r}-1").number_format = "0.0%"
        if k == "capex": RC.cell(r, 8, "共識＝購置不動產與設備（不含融資租賃本金，近四季約 3.5）")
        if k in ("ebit", "eps") and y == 26: RC.cell(r, 8, "共識在 Q3 財報前，未含 Q3 法律費用 9.89")
        RCR[f"{k}{y}"] = r; r += 1
RC.cell(r, 1, "EPS（排除一次性）FY2026 vs valueinvesting 共識（推測排除一次性）"); RC.cell(r, 2, "FY2026")
for j, s in enumerate("ABC"): RC.cell(r, 3 + j, f"=PL!G{PR_[s]['epsx']}").number_format = "0.00"
RC.cell(r, 6, "=c_eps26_vi").number_format = "0.00"; RC.cell(r, 7, f"=D{r}/F{r}-1").number_format = "0.0%"; r += 2
RC.cell(r, 1, "2. 逐分部與公司指引（情境 B）").font = BOLD; r += 1
hdr(RC, r, ["項目", "本模型", "指引低", "指引高", "共識", "說明"]); r += 1
PB = PR_["B"]; CB = CR["B"]; TB = T["B"]
def rc2(lab, model, lo, hi, cons, note):
    global r
    RC.cell(r, 1, lab)
    for j, v in enumerate((model, lo, hi, cons)):
        if v is not None: RC.cell(r, 2 + j, v).number_format = "#,##0.00"
    RC.cell(r, 6, note); r += 1
g = fv("txt.guid_q4rev.2026q3")
rc2("2026Q4 營收", f"=PL!F{PB['rev']}", g["low"], g["high"], "=c_q4rev", "公司指引（匯率約 1% 逆風）；共識 TipRanks；本模型 Q4 年增＝近四季年增")
g = fv("txt.guid_exp2026.2026q3")
rc2("2026 總費用（營收 − 財報營業利益）", f"=PL!G{PB['rev']}-PL!G{PB['oi']}", g["low"], g["high"], None, "含 Q3 法律費用 9.89")
g = fv("txt.guid_capex2026.2026q3")
rc2("2026 資本支出（含融資租賃本金）", f"=CF!G{CB['cx']}", g["low"], g["high"], "=c_capex_26", "C10：本模型＝指引中點；共識不含租賃本金")
g = fv("txt.guid_rl_loss2026.2026q3")
rc2("2026 Reality Labs 虧損", f"=-PL!G{PB['oi_rl']}", g["low"], g["high"], None, "Q4 26 依指引中點倒推")
rc2("2027 Reality Labs 虧損", f"=-PL!H{PB['oi_rl']}", None, None, None, "法說：約 15（事實總帳 call.rlLoss2027）")
rc2("2027 資本支出（情境 B，M6）", f"=CF!H{CB['cx']}", None, None, "=c_capex_27", "法說未給數字（Q4 法說提供）；共識財報前、不含租賃本金")
rc2("Family of Apps：分部營收 FY2026", f"=PL!G{PB['foa']}", None, None, None, "分部共識：未找到（公開來源只有合併層）")
rc2("Reality Labs：分部營收 FY2026", f"=PL!G{PB['rl']}", None, None, None, "分部共識：未找到")
r += 1
RC.cell(r, 1, "3. M4：未入基準的算力（只作對照與 Sens；本模型可能低估自用算力）").font = BOLD; r += 1
hdr(RC, r, ["項目", "數值", "說明"]); r += 1
RCM = {}
def rc3(key, lab, f, fmt="#,##0.0", note=""):
    global r
    RC.cell(r, 1, lab); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, note); RCM[key] = r; r += 1
rc3("ncl", "尚未起租之租賃義務（2026-09-30）", f"={fv('lease.notCommenced.20260930') / 1000}", note=fs("lease.notCommenced.20260930") + "；含 Hyperion、El Paso 合資")
rc3("ncm", "不可取消合約承諾（2026-09-30）", f"={fv('commit.noncancel.20260930') / 1000}", note=fs("commit.noncancel.20260930"))
rc3("rentfull", "租用算力年租金滿載值（假設 20 年租期；In 頁 rent_l 高值）", f"=In!F{IN_ROW['rent_l']}", note="Sens「租用算力」案例使用")
rc3("rentmw", "  換算 MW（÷ VR200 每 MW 持有成本）", f"=B{r - 1}/he_vr*1000", "#,##0", note="粗估：租金全部視為算力持有成本")
rc3("mw27", "本模型 FY2027 年底自用 AI MW（情境 B，資本支出口徑）", f"=Thesis!H{TB['end']}", "#,##0")
rc3("mw0", "本模型評價日 AI MW（支出口徑，含在建）", "=mw_tot0", "#,##0")
rc3("tpc", "第三方雲端與第三方 AI token 成本", "=0", note="法說稱費用成長來自「third-party cloud spend」與第三方 AI token 成本，未給金額：未找到")
rc3("mtia", "MTIA（自研晶片）占比", "=0", note="逐字稿、新聞稿、10-Q 讀取部分皆未提及：未找到；不套 NonNV 係數（M4）")
r += 1
RC.cell(r, 1, "4. 一次性項目：財報口徑 vs 評價口徑（M5 擴大）").font = BOLD; r += 1
hdr(RC, r, ["項目", "財報口徑", "排除一次性", "差額", "說明"]); r += 1
for lab, a, b in (("FY2026 營業利益（情境 B）", f"PL!G{PB['oi']}", f"PL!G{PB['oix']}"), ("FY2026 EPS（情境 B）", f"PL!G{PB['eps']}", f"PL!G{PB['epsx']}"), ("9M26 EPS", f"PL!E{PB['eps']}", f"PL!E{PB['epsx']}")):
    RC.cell(r, 1, lab); RC.cell(r, 2, "=" + a).number_format = "#,##0.00"; RC.cell(r, 3, "=" + b).number_format = "#,##0.00"; RC.cell(r, 4, f"=C{r}-B{r}").number_format = "#,##0.00"
    RC.cell(r, 5, "排除：Q1 CAMT 稅務利益 8.03、Q2 法律 2.40＋遣散 1.18、Q3 法律 9.89；協議現金付款仍在現金流與 DCF"); r += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFG": RC.column_dimensions[c].width = 13
RC.column_dimensions["H"].width = 60

# =====================================================================================================
# Sens
SENS_CASES = [
    ("internal_ai_reading", "自用 AI 讀法 1（只計成本）", "lo", None, "讀法 2 → 讀法 1"),
    ("ads_mode", "FoA 倍數改用 GOOGL 14.71×", None, "hi1", "中位數 9.41× → GOOGL 14.71×"),
    ("ads_mode", "FoA 倍數改用 Meta 自身 8.99×", None, "hi", "中位數 9.41× → 8.99×"),
    ("rl_mode", "Reality Labs 在分部加總以 0 計", None, "hi", "負 EBITDA × 倍數 → 0"),
    ("gl_impr", "曝光長期年增", "lo", "hi", "3%／7%"),
    ("gl_price", "每則價格長期年增", "lo", "hi", "1%／5%"),
    ("g0_q", "起始年增改用 2026Q3 單季", None, "hi", "近四季 → 2026Q3"),
    ("pA", "情境 A 價格提升幅度", "lo", "hi", "5%／15%（只影響 A）"),
    ("pC", "情境 C 價格消退幅度", "lo", "hi", "−15%／−5%（只影響 C）"),
    ("mb_q3", "FoA EBITDA 率改用 2026Q3 單季", None, "hi", "9M26 → 2026Q3（皆排除一次性）"),
    ("gl_other", "Other revenue 長期年增", "lo", "hi", "8%／25%"),
    ("rl_loss32", "FY2032 Reality Labs 虧損", "lo", "hi", "0／10"),
    ("rl_loss27", "2027 Reality Labs 虧損", "lo", "hi", "12／18"),
    ("rl_loss26", "2026 Reality Labs 虧損", "lo", "hi", "18／19"),
    ("rl_da_sh", "折舊中 RL 占比", "lo", "hi", "2%／8%"),
    ("he_all", "IF_HoldEcon 四世代同時取低／高（讀法 2 影子效益的每 MW 值）", "lo", "hi", "Tokenomics 低成本／高成本欄（GB300 10.89 → 8.43／18.53）"),
    ("capex_g", "2026 資本支出指引取區間兩端", "lo", "hi", "135／145"),
    ("cx27_k", "M6：FY2027 資本支出係數", "lo", "hi", "2026 年增率 × 1/4／3/4"),
    ("a28", "FY2028 新增 MW ÷ FY2027", "lo", "hi", "0.8／1.2"),
    ("cx_na", "非 AI 資本支出占營收", "lo", "hi", "2%／7%"),
    ("cip_ai", "在建工程中 AI 占比", "lo", "hi", "80%／100%"),
    ("cip_ai", "在建工程不另增 MW（cip_ai＝0；對齊 Alphabet 修正三的 MW 處理）", "zero", None, "90% → 0（在建工程仍照時程折舊）"),
    ("cip_it", "在建工程中 IT 占比", "lo", "hi", "50%／70%"),
    ("ai_23", "FY2023 AI 占資本支出", "lo", "hi", "50%／90%（只影響 2029 汰換）"),
    ("phi_q4", "Q4 26 批次當期折舊比例", "lo", "hi", "0.0625／0.25（高＝Alphabet 口徑）"),
    ("phi_0", "年度批次當期折舊比例", "lo", "hi", "0／0.5"),
    ("life_fac", "機房折舊年限", "lo", "hi", "15／25 年"),
    ("life_na", "非 AI 折舊年限", "lo", "hi", "8／15 年"),
    ("selfb", "自建機房比例", "lo", "hi", "75%／95%"),
    ("life_it", "AI IT 折舊年限", "lo", "hi", "5／7 年"),
    ("rent_l", "租用算力（尚未起租租賃＋第三方雲端；M4）", None, "hi", "0 → 17.8／年（FY2029 滿載）"),
    ("contg", "協議或有金額 5.3 認列為類債", None, "hi", "0 → 5.3"),
    ("legal_oth", "其他法律應計不扣", "lo", None, "10.0 → 0"),
    ("fwd_on", "遠期出售不交割", "lo", None, "計入 → 不計"),
    ("tax", "有效稅率", "lo", "hi", "13%／19%"),
    ("wc_k", "營運資金 ÷ 營收增量", "lo", "hi", "−2%／＋2%"),
    ("conv_sh", "強制轉換特別股轉換股數", "lo", None, "20.0M → 17.0M 股"),
    ("kd_sp", "公司債利差", "lo", "hi", "0.4%／1.0%"),
    ("dil", "股數年稀釋", "lo", "hi", "0%／2%"),
    ("bb_t", "回購目標", None, "hi", "0 → 26.2／年"),
    ("wacc_adj", "WACC", "lo", "hi", "−1pt／＋1pt"),
    ("beta", "β（Blume 調整後的單一來源）", "lo", "hi", "1.124／1.161"),
    ("erp", "股權風險溢酬", "lo", "hi", "4%／6%"),
    ("g_tv", "永續成長 g", "lo", "hi", "2%／4%"),
    ("ronic_sp", "RONIC＝WACC＋5pt", None, "hi", "RONIC＝WACC → WACC＋5pt"),
    ("tv_alt", "終值改用 FY2032 自由現金流", None, "hi", "價值驅動 → UFCF ×(1＋g)"),
    ("w_dcf", "DCF 權重", "lo", "hi", "30%／70%"),
    ("sbc_add", "DCF 加回股權報酬", None, "hi", "不加回 → 加回（稅後）"),
]
SE["A1"] = "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價（Val 頁）；由 meta/make_v01.sh（SENS=1）產生，建置時寫入（數值）"
SE["A2"] = "相對基準＝案例目標價 ÷ 基準加權目標價 − 1；龍捲風依絕對值排序（HTML）。H、I 欄＝收支關卡 B（五年）與 FY2027 資本支出（B）。inc_m（價格情境差額的增量 EBITDA 率）不列：A、C 對稱且權重相同，對加權目標價結構上恆為 0，只影響情境間距。gl_rl（RL 營收長期年增）亦不列：RL 營業損益依指引路徑設定，營收只經非 AI 資本支出影響目標價，重算變動小於 0.001%（−5%／15%：843.94／843.95）"
hdr(SE, 3, ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "收支關卡 B（五年）", "FY2027 資本支出（B）"])
r = 4
SE_ROWS = []
for key, lab, lo, hi, note in SENS_CASES:
    for side, val in (("低", lo), ("高", hi)):
        if val is None: continue
        SE.cell(r, 1, f"{lab}：{side}"); SE.cell(r, 2, note); SE.cell(r, 3, side); SE.cell(r, 6, key)
        res = SENS.get(f"{key}|{side}|{val}") or SENS.get(f"{key}|{side}")
        if res is not None:
            c = SE.cell(r, 4, res["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
            SE.cell(r, 5, f"=D{r}/Val!$B${VR_['tp']}-1").number_format = "+0.0%;-0.0%;0.0%"
            for cc, kk in ((8, "gateB"), (9, "cx27B")):
                c = SE.cell(r, cc, res[kk]); c.number_format = "#,##0.0"; c.font = BLUE
        SE.cell(r, 7, "重算值（make_v01.sh）")
        SE_ROWS.append((key, side, val, r)); r += 1
SE_END = r - 1
r += 1
SE.cell(r, 1, "基準加權目標價（Val 頁）"); SE.cell(r, 4, f"=Val!B{VR_['tp']}").number_format = "#,##0.00"; SE_BASE = r
SE.cell(r, 8, f"=CF!C{GATE}").number_format = "#,##0.0"; SE.cell(r, 9, f"=CF!H{CR['B']['cx']}").number_format = "#,##0.0"; r += 1
SE.cell(r, 1, "敏感度重算時的基準值（應與上一列相同；不同表示敏感度已過期）")
if "base|base" in SENS:
    c = SE.cell(r, 4, SENS["base|base"]["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
    SE.cell(r, 5, f"=D{r}-D{r - 1}").number_format = "0.000"
SE_STALE = r; r += 1
SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 36; SE.column_dimensions["G"].width = 24

# =====================================================================================================
# Sources
SO["A1"] = "Sources — 來源（編號、文件、日期、網址、標記）"
hdr(SO, 3, ["編號", "文件", "日期", "網址", "標記", "用途"])
SRCS = [
    ("Meta 2026Q3 財報新聞稿（8-K 附件 99.1）", "2026-10-21", "https://www.sec.gov/Archives/edgar/data/0001326801/000162828026066700/meta-09302026xexhibit991.htm", "Verified", "2026Q3、9M 實績、指引、一次性項目"),
    ("Meta 2026Q2、Q1、2025Q4 財報新聞稿（8-K 附件 99.1；IR PDF 第二讀）", "2026-07／04／01", "見事實總帳各筆 url", "Verified", "各季實績與比較欄（FY2024、2024Q4–2025Q3）"),
    ("Meta 2026Q3 10-Q", "2026-10-29", "https://www.sec.gov/Archives/edgar/data/0001326801/000162828026066773/meta-20260930.htm", "Verified", "債務、在建工程、租賃、承諾、法律應計、多州檢察長協議 12.7B／5.3B、股數"),
    ("Meta FWP（Class A 遠期出售、強制轉換特別股）", "2026-10-26", "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526432549/d382225dfwp.htm", "Verified", "遠期出售 40M 股、特別股 12.5B 與轉換比例"),
    ("Meta 8-K（特別股交割、Class C）", "2026-10-29", "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526438017/d183592d8k.htm", "Verified", "特別股 10-29 交割；Class C 股票股利"),
    ("Meta 2026Q3 法說逐字稿（Andy 提供，不入庫；推定 2026-10-21）", "2026-10-21", "Drive 02-算力收支/MAG財報下載/Transcript_META_3Q26.docx", "Interested-party", "2027 資本支出無數字、RL 虧損路徑、無融資計畫（通話時）"),
    ("事實總帳 meta/data/meta_facts_2026q3.json（facts_2026q3.py）", "2026-10-30", "repo 內", "各筆標記", "Q 頁全部實績與多數輸入"),
    ("Tokenomics 快照 meta/data/tokenomics_snapshot_v21.7.json（自 alphabet/data/ 複製）", "2026-10-29", "repo 內（YinchenChang/Tokenomics）", "Derived", "IF_HoldEcon、IF_OpexGW、IF_CapexIT／Facility、IF_DeprIT"),
    ("共識檔 meta/data/consensus_meta_20261030.json（MarketBeat、TipRanks、MarketScreener、valueinvesting）", "2026-10-30", "repo 內", "Interested-party", "共識目標價、年度與季度共識、現價 612.49"),
    ("同業檔 meta/data/peers_ads_20261030.json", "2026-10-07 基準", "repo 內", "Derived", "SOTP 倍數（五家 10-07＋GOOGL）"),
    ("TipRanks META 歷史股價", "2026-10-30", "https://www.tipranks.com/stocks/meta/historical-prices", "Verified", "評價日股價 725.18（9/30）"),
    ("StockScan META 歷史股價", "2026-10-30", "https://stockscan.io/stocks/META/price-history", "Verified", "10-28 收盤 612.49；9/30 −1.84%"),
    ("ChartRow META", "2026-10-30", "https://chartrow.com/quote/meta", "Interested-party", "β 1.24；10-29 收盤 624.55（單一來源，未採用）"),
    ("CoinDataFlow META", "2026-10-30", "https://coindataflow.com/en/stock/META", "Interested-party", "β（5Y Monthly）1.185"),
    ("美國財政部每日殖利率曲線（經 Alphabet v0.3.1）", "2026-10-27", "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610", "Verified", "rf 4.99%"),
]
for i, (doc, dt, url, tag, use) in enumerate(SRCS, 1):
    for j, v in enumerate((f"S{i:02d}", doc, dt, url, tag, use), 1): SO.cell(3 + i, j, v)
for col, w in zip("ABCDEF", [6, 80, 14, 70, 18, 60]): SO.column_dimensions[col].width = w

# =====================================================================================================
# README
RD["A1"] = "Meta（META）個股模型 v0.1（2026-10-30）— 命題模組＋官方營收分部 → 損益表"
RD["A3"] = "命題：Meta 的 AI 資本支出（2026 指引 1,350–1,450 億美元，含融資租賃本金）能否由廣告現金流支撐，且自用算力（GPU＋自研 MTIA）是否帶來足以回本的廣告變現提升。"
RD["A4"] = "核心問題：FY2028 AI 增量營業利益（讀法 2：自用 AI 影子效益＝自用在役 MW × Tokenomics 每 MW 持有成本）能否覆蓋 AI 資本支出的資金成本；收支關卡＝FY2026–FY2030 自由現金流扣股利合計。"
VB = lambda k, c="B": f"Val!{c}{VR_[k]}"
PA_, PC_ = PR_["A"], PR_["C"]
RD["A5"] = f'="投資建議："&{VB("rec")}&"（加權目標價 US$"&TEXT({VB("tp")},"0.00")&"，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"；現價 "&TEXT(px,"0.00")&"，2026-10-28 收盤；每股為 Class C 股票股利前口徑）"'
RD["A6"] = (f'="重點觀察 1：目標價 "&TEXT({VB("tp")},"0")&"＝DCF "&TEXT({VB("tpdcf")},"0")&"、分部加總 "&TEXT({VB("tpcomp")},"0")&" 各半；情境 A／B／C 為 "'
            f'&TEXT({VB("scen")},"0")&"／"&TEXT({VB("scen", "C")},"0")&"／"&TEXT({VB("scen", "D")},"0")&"；讀法 1 目標價 "&TEXT({VB("tp_r1")},"0")&"。市場共識目標價 "&TEXT(cons_tp,"0")&"。"')
RD["A7"] = (f'="重點觀察 2：FY2028 自用 AI 影子效益（B）"&TEXT(Thesis!I{T["B"]["shadow"]},"0.0")&"，對自用 AI 增量成本（營運＋新增 AI 折舊）"&TEXT(Thesis!I{T["B"]["int_cost"]},"0.0")'
            f'&"；FY2028 Advertising A／B／C 為 "&TEXT(PL!I{PA_["ads"]},"0")&"／"&TEXT(PL!I{PB["ads"]},"0")&"／"&TEXT(PL!I{PC_["ads"]},"0")&"，合併營業利益率（財報口徑，B）由 FY2025 "&TEXT(PL!D{PB["om"]},"0.0%")&" 變為 FY2028 "&TEXT(PL!I{PB["om"]},"0.0%")&"。"')
RD["A8"] = (f'="重點觀察 3：收支關卡（FY2026–2030 自由現金流扣股利合計）A／B／C＝"&TEXT(CF!B{GATE},"0.0")&"／"&TEXT(CF!C{GATE},"0.0")&"／"&TEXT(CF!D{GATE},"0.0")&"；B 情境 "&CF!C{GATE + 3}'
            f'&"。FY2027 資本支出（B）"&TEXT(CF!H{CR["B"]["cx"]},"0")&"，自用 AI MW 由評價日 "&TEXT(mw_tot0,"#,##0")&" 增至 FY2028 年底 "&TEXT(Thesis!I{T["B"]["end"]},"#,##0")&"。"')
r = 10
RD.cell(r, 1, "驅動 → 推導對照表（每頁輸入、推導量、因果方向）").font = BOLD; r += 1
hdr(RD, r, ["頁", "輸入（In 頁鍵）", "推導量", "因果方向"]); r += 1
DRV = [("Thesis", "capex_g、cx27_A／B／C、cx_na、selfb、a28、cip_0、cip_ai、Tokenomics IF_CapexIT／Facility／OpexGW／HoldEcon", "自用 AI MW、AI 資本支出、新增 AI 折舊、自用 AI 增量營運成本、影子效益", "資本支出 → MW（C10）→ 折舊與營運成本 → FoA 營業利益；MW → 影子效益（讀法 2）"),
       ("Seg", "近四季年增（Q 頁）、gl_impr、gl_price、pA／pC、gl_other、gl_rl、conv_n", "Advertising（曝光 × 價格）、Other revenue、FoA、RL、合併營收；非 AI 資本支出與折舊", "近四季年增 → 長期值；Advertising＝B 路徑 ×(1＋價格情境調整)"),
       ("PL", "m_foa（Q 頁）、inc_m、rl_loss26／27／32、rl_da_sh、tax、一次性項目", "分部 EBITDA、分部營業利益、合併營業利益、EPS（財報／排除一次性／評價口徑）", "分部營收 → EBITDA − 折舊 → 營業利益 → 業外、稅 → EPS"),
       ("CF", "sbc_r、wc_k、ag_pay、dps、bb_t、min_cash、mcps、fwd_cash", "自由現金流、股利、回購、新債、淨現金、收支關卡", "淨利＋折舊 − 協議付款 − 資本支出 → 自由現金流 → 股利 → 回購／舉債"),
       ("Val", "rf、beta、erp、kd_sp、g_tv、ronic_sp、同業倍數、ads_mode、rl_mode、tp_w*、w_dcf", "DCF、分部加總、目標價、上檔空間、投資建議", "無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50"),
       ("Sens", "In 頁敏感度低／高", "加權目標價變動", "一次改一個輸入、整本重算")]
for row in DRV:
    for j, v in enumerate(row, 1): RD.cell(r, j, v)
    r += 1
r += 1
RD.cell(r, 1, "本版要點（v0.1）").font = BOLD; r += 1
for t_ in ("架構同 Alphabet v0.3.2：命題（自用 AI MW）＋官方分部（Family of Apps〔Advertising、Other revenue〕、Reality Labs）→ 損益表 → 現金流 → DCF 50%＋分部加總 50%；情境權重 25／50／25；讀法 2 主值。",
           "Advertising＝曝光 × 價格：公司每季揭露兩者年增率；起始價格年增由近四季廣告金額年增反推，使兩者相乘等於金額年增。",
           "自用 MW 由資本支出經 Tokenomics 每 MW 資本支出反推（C10），Q4 26 與 FY2027 為公式解（不需迭代）；尚未起租租賃、第三方雲端只列 Recon 與 Sens（可能低估自用算力）；MTIA 不套 NonNV 係數。",
           "一次性項目（M5 擴大）：財報口徑保留、評價口徑排除；多州檢察長協議付款 12.7 分 10 年留在現金流與 DCF；或有 5.3 只列 Sens。",
           "法說日期推定 2026-10-21（新聞稿與 8-K）；現價用 2026-10-28 收盤 612.49（10-29 只找到一個來源）；同業倍數五家為 2026-10-07 基準＋GOOGL。"):
    RD.cell(r, 1, t_); r += 1
r += 1
RD.cell(r, 1, "工作表：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。藍字＝輸入、黑字＝公式；黃底＝判斷題。建置：bash meta/make_v01.sh").font = BOLD; r += 1
RD.cell(r, 1, "數字標記：Verified（公司正式財報或申報文件）／Interested-party（公司指引、賣方、第三方網站）／Analogy／Assumed（給區間）／Derived（換算）。每股數字為 Class C 股票股利（2027-01 下旬，1:1）前的口徑。")
RD.column_dimensions["A"].width = 120; RD.column_dimensions["B"].width = 60; RD.column_dimensions["C"].width = 60; RD.column_dimensions["D"].width = 60
for rr in range(3, 9): RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")

if os.environ.get("ROWMAP"):
    json.dump({"T": T, "T0": T0, "SR": SR, "PR": PR_, "CR": CR, "VR": VR_, "DC": DC, "GATE": GATE, "SUMROW": SUMROW, "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE,
               "SE_STALE": SE_STALE, "IN_ROW": IN_ROW, "RCR": RCR, "RCM": RCM, "QR": QR},
              open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=0)
wb.save(A.out)
print("saved", A.out)
