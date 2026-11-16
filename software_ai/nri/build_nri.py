# NRI（4307）個股模型 v0.1（原型 L 按人力計價；命題＋官方營收分部 → 損益表）。金額十億日圓（¥B）；每股日圓；股數百萬股；會計年度 4 月制（FY2027＝2026-04～2027-03）。
# 用法：python3 build_nri.py raw.xlsx [--sens sens.json] [--set 鍵=值 ...]；再以 LibreOffice 重算（make_nri.sh）。本版未做先驗。
import sys, os, json, argparse, hashlib, datetime, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE); sys.path.insert(0, os.path.join(HERE, "..", "..", "tools", "in_layout"))
import facts_nri as FN
ap = argparse.ArgumentParser(); ap.add_argument("out"); ap.add_argument("--sens"); ap.add_argument("--kreq"); ap.add_argument("--set", action="append", default=[])
A = ap.parse_args(); OVR = {k: float(v) for k, v in (s.split("=", 1) for s in A.set)}
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
VER = "v0.1"; VDATE = "2026-11-16"
def _bid():
    h = hashlib.sha256()
    for fn in ["build_nri.py", "facts_nri.py", os.path.join("data", "tokenomics_snapshot_v22.0.json")]: h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"nri-{VER}-{h.hexdigest()[:10]}"
BUILD_ID = _bid()
_TKJ = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8")); TK = _TKJ["items"]; TK_VER = _TKJ["source"]["version"]
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else None
wb = openpyxl.Workbook()
def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0): c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
YK = ["24", "25", "26", "1H", "2H", 27, 28, 29, 30, 31, 32]
class _C(dict):
    def __missing__(self, k): return k
COL = _C(zip(YK, "CDEFGHIJKLM"))
YH = {"24": "FY2024A", "25": "FY2025A", "26": "FY2026A", "1H": "1H27A", "2H": "2H27E", 27: "FY2027E", 28: "FY2028E", 29: "FY2029E", 30: "FY2030E", 31: "FY2031E", 32: "FY2032E"}
FCOL = "IJKLM"; NOTE_C = 14
def put(ws, r, label, unit, vals, fmt="#,##0.0", note=None, bold=False):
    ws.cell(r, 1, label); ws.cell(r, 2, unit)
    if bold: ws.cell(r, 1).font = BOLD
    for p, v in vals.items():
        c = ws[f"{COL.get(p, p)}{r}"]; c.value = v; c.number_format = fmt
        if isinstance(v, (int, float)): c.font = BLUE
    if note: ws.cell(r, NOTE_C, note)
def head(ws, r, label="項目"):
    hdr(ws, r, [label, "單位"] + [YH[k] for k in YK] + ["說明"])
def widths(ws, a=58):
    ws.column_dimensions["A"].width = a; ws.column_dimensions["B"].width = 10
    for c in "CDEFGHIJKLM": ws.column_dimensions[c].width = 11
    ws.column_dimensions["N"].width = 80
RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); HS = wb.create_sheet("Hist"); Q = wb.create_sheet("Q"); TH = wb.create_sheet("Thesis"); PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF")
V = wb.create_sheet("Val"); RV = wb.create_sheet("RevDCF"); RC = wb.create_sheet("Recon"); SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")
G = lambda x: x / 1000.0
SEGS = ["Consulting", "Financial IT Solutions", "Industrial IT Solutions", "IT Platform Services", "Others／Adjustments"]
SK = {"Consulting": "cons", "Financial IT Solutions": "fin", "Industrial IT Solutions": "ind", "IT Platform Services": "itp", "Others／Adjustments": "oth"}
SUBF = ["Securities", "Insurance", "Banking", "Other financial"]; SUBI = ["Distribution", "Manufacturing, service & other"]
# ===================================== Hist（歷史財務；原則第 10 節第 14 條）
HS["A1"] = "Hist — 歷史財務（¥B；每股日圓；股數百萬股）。年度 C–G＝FY2022A–FY2026A（4 月制，FY2026＝2025-04～2026-03）；H＝1H FY2026（2025-04～09，新口徑重編）；I＝1H FY2027（2026-04～09）。出處與擷取方式見 data/facts_nri.csv（WebFetch 摘要模型轉錄，非逐字人工核對）"
hdr(HS, 3, ["項目", "單位", "FY2022A", "FY2023A", "FY2024A", "FY2025A", "FY2026A", "1H26A", "1H27A", "標記", "來源", "來源誘因"])
HR = {}; r = 4
def hrow(key, label, unit, vals, fmt="#,##0.000", tag="Verified", src="", inc=""):
    global r
    HR[key] = r; HS.cell(r, 1, label); HS.cell(r, 2, unit)
    for c, v in zip("CDEFGHI", vals):
        if v is not None: x = HS[f"{c}{r}"]; x.value = v; x.number_format = fmt; x.font = BLUE
    HS.cell(r, 10, tag); HS.cell(r, 11, src); HS.cell(r, 12, inc); r += 1
def hsec(t):
    global r
    HS.cell(r, 1, t).font = BOLD; r += 1
hsec("A. 合併損益、現金流與資產負債（決算短信；IFRS）")
SRC_A = "決算短信 FY3/23、FY3/24、FY3/25、FY3/26（見 facts_nri.csv）；1H＝2026-10-29 與 2025-10-30 決算短信"
for k, nm, fmt in [("rev", "營收", "#,##0.000"), ("cogs", "營業成本", "#,##0.000"), ("gp", "毛利", "#,##0.000"), ("sga", "販管費", "#,##0.000"), ("oi", "營業利益", "#,##0.000"), ("pbt", "稅前利益", "#,##0.000"),
                   ("tax", "所得稅", "#,##0.000"), ("ni", "歸屬母公司淨利", "#,##0.000"), ("eps", "基本 EPS", "0.00"), ("ocf", "營業活動現金流", "#,##0.000"), ("cx_p", "資本支出：有形資產", "#,##0.000"),
                   ("cx_i", "資本支出：無形資產", "#,##0.000"), ("da", "折舊攤銷（含使用權資產）", "#,##0.000"), ("div", "股利支付", "#,##0.000"), ("bb", "庫藏股取得", "#,##0.000"),
                   ("cash", "現金及約當現金（期末）", "#,##0.000"), ("debt", "有利子負債（期末）", "#,##0.000"), ("eqo", "歸屬母公司權益（期末）", "#,##0.000"), ("dps", "每股股利", "0"),
                   ("sh_iss", "期末發行股數（含庫藏）", "#,##0.000000"), ("sh_tre", "期末庫藏股", "#,##0.000000")]:
    vals, _, _ = FN.ANN[k]; sc = (lambda x: x) if k in ("eps", "dps") or k.startswith("sh_") else G
    h1 = FN.H1.get(k); hv = [None, None]
    if h1: hv = [sc(h1[0]), sc(h1[1])]
    hrow(k, nm, "日圓" if k in ("eps", "dps") else ("百萬股" if k.startswith("sh_") else "¥B"), [sc(x) for x in vals] + hv, fmt, "Verified", SRC_A)
hrow("fin_i", "財務收益（半期）", "¥B", [None] * 5 + [G(FN.H1["fin_i"][0]), G(FN.H1["fin_i"][1])]); hrow("fin_c", "財務費用（半期）", "¥B", [None] * 5 + [G(FN.H1["fin_c"][0]), G(FN.H1["fin_c"][1])])
hrow("pl", "當期利益（含非控制）（半期）", "¥B", [None] * 5 + [G(FN.H1["pl"][0]), G(FN.H1["pl"][1])]); hrow("sh_avg", "加權平均股數（半期）", "百萬股", [None] * 5 + FN.H1["sh_avg"], "#,##0.000000")
hsec("B. 分部外部營收（舊口徑：FY2022–FY2026 當時揭露；FY2025、FY2026 為 FY3/26 決算短信口徑）")
HOLD = {}
for s in SEGS:
    HOLD[s] = r; hrow("old_" + SK[s], s, "¥B", [G(x) for x in FN.SEGX["old"][s]], tag="Verified", src="各年度決算短信分部資訊（FY2022 外部營收合計 611,632，與合併差 2，四捨五入）")
hsec("C. 分部外部營收（新口徑，2026-04 起：澳洲移入 Financial IT）：G＝FY2026 重編；H＝1H26 重編；I＝1H27")
HNEW = {}
for s in SEGS:
    HNEW[s] = r; hrow("new_" + SK[s], s, "¥B", [None, None, None, None, G(FN.SEGX["new26"][s]), G(FN.H1SEG_EXT[s][0]), G(FN.H1SEG_EXT[s][1])], src="2026-10-29 決算短信 p.9、p.5–6、p.19")
hsec("D. 分部營業利益（FY2022–FY2026 舊口徑，含內部交易口徑，FY2026 Industrial 含減損 97.060；H、I＝新口徑半期）")
HOI = {}
for s in SEGS:
    HOI[s] = r; hrow("oi_" + SK[s], s, "¥B", [G(x) for x in FN.SEGO[s]] + [G(FN.H1SEG_OI[s][0]), G(FN.H1SEG_OI[s][1])], src="各年度決算短信分部資訊；半期 2026-10-29 決算短信 p.5–6")
hsec("E. 分部總營收（含內部交易；半期新口徑）")
HTOT = {}
for s in SEGS[:4]:
    HTOT[s] = r; hrow("tot_" + SK[s], s, "¥B", [None] * 5 + [G(FN.H1SEG_TOT[s][0]), G(FN.H1SEG_TOT[s][1])], src="2026-10-29 決算短信 p.5–6、p.19")
hsec("F. 分部下產品線（公司固定揭露的產業別外部營收）：G＝FY2026（新口徑）；I＝1H27")
HSUB = {}
for s in SUBF + SUBI:
    HSUB[s] = r; hrow("sub_" + s[:4], s, "¥B", [None] * 4 + [G(FN.SUBL26[s]), None, G(FN.H1SUB27[s])], src="2026-10-29 決算短信 p.9；決算說明資料 p.6–7（1H27）")
hsec("G. 2026-09-30 資產負債表與股數（2026-10-29 決算短信 p.7、p.11–12）")
HB = {}
for k, nm, v, u in [("cash", "現金及約當現金", G(FN.BS27["cash"]), "¥B"), ("ofa", "其他流動金融資產", G(FN.BS27["ofa"]), "¥B"), ("debt", "有利子負債", G(FN.BS27["debt"]), "¥B"), ("mi", "非控制權益", G(FN.BS27["mi"]), "¥B"),
                    ("iss", "發行股數", FN.BS27["sh_iss"], "百萬股"), ("tre", "庫藏股（含員工持股信託）", FN.BS27["sh_tre"], "百萬股"), ("lease", "租賃負債（2026-03）", G(FN.BS27["lease_mar26"]), "¥B")]:
    HB[k] = r; hrow("b_" + k, nm, u, [None] * 6 + [v], "#,##0.000000" if u == "百萬股" else "#,##0.000")
hsec("H. 其他公司揭露（公司 IR 資料；Interested-party）")
HO = {}
for k, nm, v, u in [("emp", "NRI 集團員工數（2026-03）", FN.OTHER["emp"], "人"), ("sub", "外包人力（日本約 9,500＋海外約 7,000；欄位標示含糊）", FN.OTHER["sub_jp"] + FN.OTHER["sub_off"], "人"),
                    ("fx", "美元兌日圓平均匯率 1H27", FN.OTHER["fx_usd_h1"], "日圓"), ("bb27", "FY2027 上半期庫藏股取得（市場 70.000＋員工持股信託 23.860）", 69.999 + 23.860, "¥B")]:
    HO[k] = r; hrow("o_" + k, nm, u, [None] * 6 + [v], "#,##0.00", "Interested-party" if k in ("emp", "sub") else "Verified", "決算說明資料 p.44；決算短信 p.8、p.17", "公司自揭 KPI：呈現規模的誘因，口徑由公司定義；兩數相加為 Derived" if k in ("emp", "sub") else "")
HS.column_dimensions["A"].width = 60
for c in "CDEFGHI": HS.column_dimensions[c].width = 11
def H_(key, col): return f"Hist!${col}${HR[key]}"
# ===================================== In
IN["A1"] = "In — 輸入（藍字；基準欄 E 進主模型）。金額 ¥B；每股日圓；股數百萬股；4 月制。標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 給低／高區間）"
IN["A2"] = '="資料截止日 "&TEXT(data_date,"yyyy-mm-dd")&"；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"；評價日 "&TEXT(val_date,"yyyy-mm-dd")&"；目標時點 "&TEXT(tp_date,"yyyy-mm-dd")&"（股價基準日＋12 個月）。黃底＝判斷題。"'
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列）"])
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", judg=False): ROWS.append((key, item, unit, lo, base, hi, tag, src, note, judg))
P = FN.OTHER; peers_pe = [p[2] for p in FN.PEERS]
q = statistics.quantiles(peers_pe, n=4, method="inclusive")
sec("A. 日期、現價、股數與資產負債")
inp("data_date", "資料截止日（最晚輸入：10 年期公債 2026-11-11 報導；財務省入札 11-05）", "日期", None, datetime.date(2026, 11, 11), None, "Derived", "facts_nri.csv 最晚擷取日期")
inp("px_date", "股價基準日（現價日）", "日期", None, datetime.date(2026, 11, 9), None, "Verified", "StockAnalysis 4307 統計頁（資料日 2026-11-09）", "距工作日 2026-11-16 已逾 2 個交易日：標「逾期」（G2 第 10 條）；2026-11-05 收盤未找到")
inp("val_date", "評價日（最近一季資產負債表日）", "日期", None, datetime.date(2026, 9, 30), None, "Verified", "2026-10-29 決算短信")
inp("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條")
inp("px", "現價（2026-11-09 收盤，日圓）", "日圓", None, P["px"], None, "Verified", "StockAnalysis 4307（2026-11-09 15:30 JST 延遲報價）", "單一來源；Yahoo!ファイナンス 最新可讀日為 2026-10-22（6,195），第二來源同日未找到")
inp("px_val", "評價日收盤（2026-09-30，日圓；WACC 權益權重用現價 px，本列只作對照）", "日圓", None, P["px_val"], None, "Verified", "Yahoo!ファイナンス 4307 時系列")
inp("sh_out", "流通股數（發行股數 − 庫藏股，2026-09-30；法定申報封面口徑）", "百萬股", None, f"={H_('b_iss','I')}-{H_('b_tre','I')}", None, "Derived", "決算短信 p.1（發行 561.197409、庫藏 5.811620；庫藏含員工持股信託）", "G2-4：NRI 無 RSU／選擇權揭露，稀釋後 EPS＝基本 EPS（1H27 皆 106.61），稀釋股數不另加；ASR、可轉債無")
inp("cash_0", "現金及約當現金（2026-09-30）", "¥B", None, f"={H_('b_cash','I')}", None, "Verified", "決算短信 p.7")
inp("ofa_k", "其他流動金融資產計入淨現金的比例（14.040；性質為定存與有價證券，未逐項核對）", "比例", 0, 1, 1, "Assumed", "決算短信 p.7；比例為判斷", "低＝不計入", judg=True)
inp("debt_0", "有利子負債（2026-09-30；公司口徑，推定含租賃負債）", "¥B", None, f"={H_('b_debt','I')}", None, "Verified", "決算短信 p.7、p.11", "G2-5：只計債務與融資租賃；營業租賃不另計。有利子負債是否已含租賃負債未能由摘要確認")
inp("lease_v", "租賃負債另加（若有利子負債不含租賃；2026-03 為 31.736）", "¥B", 0, 0, G(FN.BS27["lease_mar26"]), "Assumed", "決算短信 FY3/26 p.9–11", "基準 0（推定已含）；高＝另加 31.736", judg=True)
inp("mi_0", "非控制權益", "¥B", None, f"={H_('b_mi','I')}", None, "Verified", "決算短信 p.7")
inp("fx", "美元兌日圓（1H27 公司平均匯率）", "日圓", 130, P["fx_usd_h1"], 160, "Verified", "決算說明資料 p.22", "只用於把 Tokenomics 美元推論成本換成日圓")
sec("B. 營收：AI 前基準（量 × 價）與分部")
inp("h2_pers", "2H27 年增相對 1H27 年增的延續比例（各分部）", "比例", 0.4, 0.75, 1.0, "Assumed", "判斷；1H27 年增 +8.3% 含大型專案與匯率", "中點；Recon 顯示 FY2027E 與指引、共識的差距，不用指引反推", judg=True)
inp("pg", "每小時單價年增（工資與組合；AI 前）", "%", 0.0, 0.02, 0.04, "Assumed", "判斷；公司未揭露單價趨勢（決算說明資料「未揭露」）")
inp("gv_28", "AI 前需求量年增 FY2028", "%", 0.02, 0.035, 0.05, "Assumed", "判斷；FY2022–FY2026 營收年增 12.5%／6.4%／3.8%／6.5%（Hist Derived）", "", judg=True)
inp("gv_32", "AI 前需求量年增 FY2032（線性收斂）", "%", 0.01, 0.02, 0.03, "Assumed", "判斷")
inp("dg_cons", "Consulting 需求量年增相對全體", "pt", -0.01, 0.0, 0.01, "Assumed", "判斷")
inp("dg_fin", "Financial IT 需求量年增相對全體（FY2026 營收年增 +8.7%，高於全體）", "pt", 0.0, 0.01, 0.02, "Assumed", "FY3/26 決算說明資料 p.8；差距為判斷")
inp("dg_ind", "Industrial IT 需求量年增相對全體", "pt", -0.02, -0.01, 0.0, "Assumed", "判斷")
inp("gi_28", "IT Platform Services 外部營收年增 FY2028（FY2026 +14.9%、1H27 +17.4%）", "%", 0.05, 0.09, 0.13, "Assumed", "判斷；訂單餘額 1H27 +33.5%（Interested-party）")
inp("gi_32", "IT Platform Services 外部營收年增 FY2032", "%", 0.03, 0.05, 0.07, "Assumed", "判斷")
sec("C. 命題：交付人員恆等式與 AI（原型 L；變數同 unit_economics）")
inp("util", "利用率（計費工時 ÷ 工時）", "比例", 0.72, 0.80, 0.85, "Assumed", "公司未揭露數值利用率（決算說明資料）；類比日本 SI 慣例", "只影響隱含人數與節省人力的換算", judg=True)
inp("hrs_y", "每人年工時", "小時", 1800, 1900, 2000, "Assumed", "判斷")
inp("rate0", "FY2027 每計費小時單價（AI 前）", "日圓", 13000, 16000, 19000, "Analogy", "類比日本 SI 人月單價；NRI 未揭露。由恆等式推隱含人數，並與公司揭露人數（Hist 第 H 節）對照於 Recon", "G2-1：NRI 未揭露 token 量或代理任務量，無法以揭露校準每任務 token", judg=True)
inp("c_head", "每名交付人員年成本（員工＋外包，¥M）", "¥M", 9.0, 11.0, 13.0, "Assumed", "判斷；勞務費拆分未於決算短信揭露（未找到）")
inp("sav_r", "自動化省下人力的成本實現率（外包可彈性、員工難削減）", "比例", 0.5, 0.7, 0.9, "Assumed", "判斷")
inp("proj_m", "AI 新專案營業利益率", "%", 0.18, 0.25, 0.30, "Assumed", "判斷；公司整體 20.5%（1H27）")
inp("auto_A", "情境 A：AI 增量自動化比例（FY2032 相對 FY2027，計費工時被取代的比例）", "比例", None, 0.30, None, "Assumed", "判斷；公司：AI 開發工具 2026 年約 25% 適用、目標 FY2028 全面導入（Interested-party，決算說明資料 p.11）")
inp("auto_B", "情境 B：AI 增量自動化比例", "比例", None, 0.25, None, "Assumed", "同上")
inp("auto_C", "情境 C：AI 增量自動化比例", "比例", None, 0.25, None, "Assumed", "同上")
inp("ret_A", "情境 A：價格保留率（客戶不分享生產力的比例）", "比例", None, 0.85, None, "Assumed", "判斷；1H27 Financial IT 營收 +8.8%、營業利益 +17.7%（公司稱 AI 生產力革新貢獻，Interested-party）")
inp("ret_B", "情境 B：價格保留率", "比例", None, 0.70, None, "Assumed", "判斷")
inp("ret_C", "情境 C：價格保留率（客戶要求分享，改採結果計價）", "比例", None, 0.50, None, "Assumed", "判斷")
inp("proj_A", "情境 A：AI 新專案營收占 AI 前基準的比例（FY2032）", "比例", None, 0.05, None, "Assumed", "判斷；AX 顧問營收 1H27 約倍增（Interested-party，決算說明資料 p.13）")
inp("proj_B", "情境 B：AI 新專案營收比例", "比例", None, 0.03, None, "Assumed", "判斷")
inp("proj_C", "情境 C：AI 新專案營收比例", "比例", None, 0.015, None, "Assumed", "判斷")
inp("inf_mult", "代理額外推論倍數（重試、工具呼叫、審查）", "倍", 1, 5, 50, "Assumed", "判斷；Tokenomics 為單次任務嘗試的 token", "G2-1：未揭露，標 Assumed")
inp("tk_tier", "模型層級（1＝Terra、2＝Sol、3＝Astra；選哪層價格屬 Analogy）", "層級", 1, 2, 3, "Analogy", "Tokenomics v22.0 IF_Price*／IF_TaskSucc*（Thesis 頁）；市場價，非全成本", "G2-1、2：價格層級為類比")
sec("D. 分部利潤（最近半年實績起點，全期持平；原則第 10 節第 35 條、G2-3）")
for k, nm, s in [("mg_cons", "Consulting", "Consulting"), ("mg_fin", "Financial IT", "Financial IT Solutions"), ("mg_ind", "Industrial IT", "Industrial IT Solutions"), ("mg_itp", "IT Platform Services", "IT Platform Services")]:
    lo = {"mg_cons": 0.20, "mg_fin": 0.16, "mg_ind": 0.11, "mg_itp": 0.17}[k]; hi = {"mg_cons": 0.28, "mg_fin": 0.185, "mg_ind": 0.145, "mg_itp": 0.205}[k]
    inp(k, f"{nm} 營業利益率（占分部總營收含內部；1H27 實績）", "%", lo, f"={H_('oi_'+SK[s],'I')}/{H_('tot_'+SK[s],'I')}", hi, "Derived", "決算短信 p.5–6（1H27）；低高為 FY2026 與 1H27 區間", "低高＝Assumed")
for s in SEGS[:4]:
    inp("rt_" + SK[s], f"{s} 總營收 ÷ 外部營收（含內部交易）", "倍", None, f"={H_('tot_'+SK[s],'I')}/{H_('new_'+SK[s],'I')}", None, "Derived", "決算短信 p.5–6（1H27），全期持平")
inp("adj_oi", "調整額營業利益（年）", "¥B", None, f"=2*{H_('oi_oth','I')}", None, "Derived", "1H27 調整額 0.533 年化")
inp("seas", "2H 營業利益相對 1H（季節性）", "比例", 0.93, 0.978, 1.03, "Derived", "FY2026 事業利益（156.673−79.213）÷ 79.213＝0.978；FY2026 決算說明資料 p.6", "低高為 Assumed")
inp("sga_r", "販管費率（1H27 實績；只用於標準損益表拆分，不影響營業利益）", "%", None, f"={H_('sga','I')}/{H_('rev','I')}", None, "Derived", "決算短信 p.3", "G2-3：基準持平")
sec("E. 損益以下、現金流與股東回饋")
inp("fin_net", "財務收益淨額（年）", "¥B", 0.0, 2.0, 5.0, "Assumed", "1H27 1.432；FY2026 0.578；FY2025 −0.746（Hist）")
inp("tax", "稅率（1H27 實績 稅 ÷ 稅前）", "%", 0.297, f"={H_('tax','I')}/{H_('pbt','I')}", 0.335, "Derived", "決算短信 p.3；低＝FY2025 實績", "低高＝Assumed")
inp("mi_r", "非控制權益占當期利益比", "%", None, f"=1-{H_('ni','I')}/{H_('pl','I')}", None, "Derived", "決算短信 p.3")
inp("payout", "配息率（占歸屬母公司淨利；公司目標約 40%，FY2027 指引 40.1%）", "%", 0.35, 0.40, 0.50, "Interested-party", "決算短信 p.9–10", "公司政策；", judg=True)
inp("bb_27", "FY2027 上半期已執行庫藏股（市場購入＋員工持股信託）", "¥B", None, f"={H_('o_bb27','I')}", None, "Verified", "決算短信 p.8、p.17")
inp("bb_t", "FY2028 起每年庫藏股", "¥B", 0, 40, 80, "Assumed", "公司：彈性、ROE 目標 25%（Interested-party）；FY2024 68.4／FY2025 30.3／FY2026 0（Hist）", "", judg=True)
inp("bb_pg", "庫藏股買進價年漲幅（自現價）", "%", 0.0, 0.04, 0.08, "Assumed", "判斷")
inp("capex_r", "資本支出（現金，有形＋無形）÷ 營收（FY2026 實績）", "%", 0.050, f"=({H_('cx_p','G')}+{H_('cx_i','G')})/{H_('rev','G')}", 0.065, "Derived", "決算短信 FY3/26 p.14–16", "低高＝Assumed")
inp("da_r", "折舊攤銷（不含使用權資產）÷ 營收（FY2026：43.0 ÷ 營收）", "%", 0.048, 0.043 / 0.814708, 0.058, "Derived", "決算說明資料 FY3/26 p.47（FY3/26 欄 43.0；欄位標示含糊）", "使用權資產折舊視為營業費用（租賃作營業成本口徑），與淨債務不另加租賃一致；低高＝Assumed")
inp("wc_k", "營運資金增加 ÷ 營收增量（代理：OCF 與淨利＋折舊＋減損差額÷營收增量，約 0.3）", "比例", 0.10, 0.30, 0.50, "Assumed", "Hist：FY2026 OCF 147.6 − (15.8＋51.3＋97.6)＝−17.0；FY2025 −13.4；摘要無現金流量表營運項目細節（未找到）", "規格要求歷史比率；此為代理", judg=True)
sec("F. 折現率與目標價規則")
inp("rf", "無風險利率（10 年期日本公債，財務省入札平均殖利率 2026-11-05）", "%", 0.022768, P["rf"], 0.032768, "Verified", "財務省 10 年債入札結果；2026-11-11 報導約 2.77%（TradingEconomics）", "低高＝±1pt（Assumed）")
inp("beta_raw", "NRI 原始 β（5 年，資料商）", "倍", None, P["beta_raw"], None, "Verified", "StockAnalysis 2026-11-09（0.34；09-09 頁為 0.32）", "單一資料商，非自算產業 β（自算產業β文件無資訊服務組）；同業 β 同源 −0.09～0.74")
inp("beta", "β（Blume：0.67×原始＋0.33）", "倍", "=(0.34-0.3)*0.67+0.33" if False else 0.258, "=0.67*beta_raw+0.33", 0.858, "Derived", "docs/plan/自算產業β_20261105.md 的 Blume 規則；低高＝±0.3（Analogy）", "β 偏低使 ke、WACC 偏低；DCF 對此敏感（Sens）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "個股 DCF 標準規格第 2 節預設")
inp("kd", "稅前負債成本（財務費用年化 ÷ 平均有利子負債＝1.983×2÷211）", "%", 0.015, 0.0188, 0.03, "Derived", "決算短信 p.3、p.7", "低高＝Assumed")
inp("g_tv", "永續成長率上限（實際取 MIN(本值, 10 年期公債殖利率)）", "%", 0.02, 0.03, 0.04, "Assumed", "個股 DCF 標準規格 第 3 節")
inp("ronic_sp", "RONIC − WACC", "pt", 0.0, 0.0, 0.05, "Assumed", "規格預設 RONIC＝WACC")
inp("wacc_adj", "WACC 調整", "pt", None, 0.0, None, "Assumed", "無")
inp("w_dcf", "目標價中 DCF 權重（其餘為 Comparable）", "比例", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設各半")
inp("tp_wA", "情境權重 A", "比例", None, 0.25, None, "Assumed", "預設 25／50／25（X4）", judg=True)
inp("tp_wB", "情境權重 B", "比例", None, 0.50, None, "Assumed", "同上", judg=True)
inp("tp_wC", "情境權重 C", "比例", None, 0.25, None, "Assumed", "同上", judg=True)
inp("pe_mult", "同業前瞻 P/E 中位數（8 家；NTM 口徑，Recon 第 4 節）", "倍", q[0], "=MEDIAN(Recon!$C$60:$C$67)", q[2], "Interested-party", "StockAnalysis 各公司統計頁（估值日 09-11～11-16 不一致，逾期標示）", "低高＝同業 25／75 分位；資料商前瞻倍數含賣方預估")
inp("rm_adj", "相對倍數調整係數", "倍", None, 1.0, None, "Assumed", "1.00（未調整：缺 NRI 與同業十年前瞻倍數；tools/data 無 4307）", "原則第 10 節第 18 條")
inp("dps_int", "FY2027 中間股利（2026-09-29 除權，現價已不含）", "日圓", None, 42, None, "Verified", "決算短信 p.9")
IN_ROW = {}; r = 4
def isnum(v): return isinstance(v, (int, float)) and not isinstance(v, bool)
for row in ROWS:
    if row[0] == "sec": IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src, note, judg = row
    if key in OVR: base = OVR[key]
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            c.number_format = "yyyy-mm-dd" if unit == "日期" else ("0.0%" if unit in ("%", "pt") else "#,##0.000")
    IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, note)
    if tag == "Interested-party": IN.cell(r, 10, "公司管理層（指引、政策、IR 資料）：維持市場信心的誘因，未經查核" if key == "payout" else "資料商前瞻倍數含賣方預估：賣方評等與交易業務誘因；估值日不一致")
    if judg:
        for cc in range(1, 10): IN.cell(r, cc).fill = KEY
    dn(key, f"In!$E${r}"); IN_ROW[key] = r; r += 1
IN_LAST = r - 1
for col, w in zip("ABCDEFGHIJ", [14, 70, 10, 10, 11, 10, 16, 80, 70, 50]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
# ===================================== Q（四季）
Q["A1"] = "Q — 最近四季（¥B；EPS 日圓）。2Q FY2027 與 3Q、4Q FY2026 為由累計值相減的 Derived；分部季度只有 1Q FY2027 揭露（新口徑）。"
Q["A2"] = "核對：4 季合計＝TTM（營收 FY2026＋1H27−1H26）。3Q、4Q FY2026 分部為舊口徑（Q 頁 H 之後欄）；新口徑 2H FY2026 見 Thesis 頁。"
hdr(Q, 3, ["項目", "單位", "3Q FY2026", "4Q FY2026", "1Q FY2027", "2Q FY2027", "標記", "來源", "來源誘因"])
QR = {}
qdata = [("營收", "rev", [G(FN.M9_26["rev"] - FN.H1["rev"][0]), G(FN.ANN["rev"][0][4] - FN.M9_26["rev"]), G(FN.Q1["rev"]), G(FN.H1["rev"][1] - FN.Q1["rev"])]),
         ("營業利益（4Q FY2026 含減損 97.586）", "oi", [G(FN.M9_26["oi"] - FN.H1["oi"][0]), G(FN.ANN["oi"][0][4] - FN.M9_26["oi"]), G(FN.Q1["oi"]), G(FN.H1["oi"][1] - FN.Q1["oi"])]),
         ("歸屬母公司淨利", "ni", [G(FN.M9_26["ni"] - FN.H1["ni"][0]), G(FN.ANN["ni"][0][4] - FN.M9_26["ni"]), G(FN.Q1["ni"]), G(FN.H1["ni"][1] - FN.Q1["ni"])]),
         ("EPS（日圓）", "eps", [round(FN.M9_26["eps"] - FN.H1["eps"][0], 2), round(FN.ANN["eps"][0][4] if False else 26.62 - FN.M9_26["eps"], 2), FN.Q1["eps"], round(FN.H1["eps"][1] - FN.Q1["eps"], 2)])]
for i, (nm, k, v) in enumerate(qdata):
    rr = 4 + i; QR[k] = rr; Q.cell(rr, 1, nm); Q.cell(rr, 2, "日圓" if k == "eps" else "¥B")
    for j, x in enumerate(v): c = Q.cell(rr, 3 + j, x); c.font = BLUE; c.number_format = "#,##0.000"
    Q.cell(rr, 7, "Derived"); Q.cell(rr, 8, "決算短信累計值相減（9M、FY、1Q、1H；facts_nri.csv）")
rr = 9; Q.cell(rr, 1, "核對：四季營收合計 − (FY2026 營收 + 1H27 − 1H26)")
Q.cell(rr, 3, f"=SUM(C4:F4)-({H_('rev','G')}+{H_('rev','I')}-{H_('rev','H')})").number_format = "0.000"; QR["chk"] = rr
Q.column_dimensions["A"].width = 60; Q.column_dimensions["H"].width = 70
# ===================================== Thesis（命題：L 原型單位經濟引擎）
widths(TH)
TH["A1"] = "Thesis — 命題：AI 進入 IT 服務後，每名交付人員年的營收變化＝AI 新專案增量 − 價格侵蝕（客戶分享生產力）− 推論成本；利潤再加自動化省下的人力成本。恆等式：營收（不含 AI 新專案）＝人數 × 利用率 × 每人年工時 × 實現小時單價。變數名稱同 unit_economics（v0.2 改讀）；推論成本只讀 Tokenomics v22.0 快照（市場價，不用全成本）。"
TH["A2"] = "範圍：恆等式驅動三個人力計價分部（Consulting、Financial IT Solutions、Industrial IT Solutions）；IT Platform Services 以成長路徑。AI 效果以 FY2027E 為零點（FY2027E 已含 1H27 實績的 AI 影響），FY2032 到位、線性。"
r = 4; head(TH, r, "Tokenomics 讀值（v22.0 快照）"); r += 1
TKR = {}
def tkrow(key, label, v, unit, name, cell):
    global r
    TKR[key] = r; TH.cell(r, 1, label); TH.cell(r, 2, unit); c = TH.cell(r, 3, v); c.font = BLUE; c.number_format = "#,##0.0000"; TH.cell(r, NOTE_C, f"{name}；{cell}；Derived（Tokenomics 權威值，本頁只讀）"); r += 1
for t_i, t in enumerate(["Terra", "Sol", "Astra"]):
    for k, nm, lab in [("pf", "IF_PriceFresh_", "新鮮輸入"), ("pc", "IF_PriceCached_", "快取輸入"), ("po", "IF_PriceOut_", "輸出")]:
        it = TK[nm + t]; tkrow(f"{k}{t_i+1}", f"{t}｜{lab}市場價（有效）", it["values"], "$/M", nm + t, it.get("cell") or it.get("ref"))
    tkrow(f"s{t_i+1}", f"{t}｜成功率（2 小時任務）", TK["IF_TaskSucc_" + t]["values"][3], "比例", "IF_TaskSucc_" + t + "[4]", "Interface 列（任務長度 2 小時欄）")
tkrow("len", "任務長度（人類完成時間）", TK["IF_TaskLen"]["values"][3], "小時", "IF_TaskLen[4]", "Interface")
tkrow("tf", "每次嘗試新鮮輸入 token（2 小時任務）", TK["IF_TaskTokFresh"]["values"][3], "tok", "IF_TaskTokFresh[4]", "Interface")
tkrow("tc", "每次嘗試快取輸入 token", TK["IF_TaskTokCached"]["values"][3], "tok", "IF_TaskTokCached[4]", "Interface")
tkrow("td", "每次嘗試 decode token", TK["IF_TaskTokDec"]["values"][3], "tok", "IF_TaskTokDec[4]", "Interface")
T = lambda k: f"$C${TKR[k]}"
TH.cell(r, 1, "每次嘗試推論成本（美元；選定層級）"); TH.cell(r, 2, "$"); TH.cell(r, 3, f"=({T('tf')}*CHOOSE(tk_tier,{T('pf1')},{T('pf2')},{T('pf3')})+{T('tc')}*CHOOSE(tk_tier,{T('pc1')},{T('pc2')},{T('pc3')})+{T('td')}*CHOOSE(tk_tier,{T('po1')},{T('po2')},{T('po3')}))/1000000").number_format = "0.00000"; TKR["att"] = r; r += 1
TH.cell(r, 1, "每個成功任務推論成本（日圓；÷成功率 × 代理額外倍數 × 匯率）"); TH.cell(r, 2, "日圓"); TH.cell(r, 3, f"=C{TKR['att']}/CHOOSE(tk_tier,{T('s1')},{T('s2')},{T('s3')})*inf_mult*fx").number_format = "#,##0.00"; TKR["task"] = r; r += 2
# 共用：AI 前分部營收
head(TH, r, "分部外部營收：AI 前基準（三情境共用；¥B）"); r += 1
TR = {}
def rowkey(k): return TR[k]
TR["t"] = r; put(TH, r, "t（FY2027E＝0）", "年", {27: 0, 28: 1, 29: 2, 30: 3, 31: 4, 32: 5}, "0"); r += 1
TR["gv"] = r; put(TH, r, "AI 前需求量年增（線性 FY2028→FY2032）", "%", {c: f"=gv_28+(gv_32-gv_28)*({c}{TR['t']}-1)/4" for c in FCOL}, "0.0%"); r += 1
TR["gi"] = r; put(TH, r, "IT Platform Services 年增", "%", {c: f"=gi_28+(gi_32-gi_28)*({c}{TR['t']}-1)/4" for c in FCOL}, "0.0%"); r += 1
dgk = {"cons": "dg_cons", "fin": "dg_fin", "ind": "dg_ind"}
for s in SEGS:
    k = SK[s]; TR["S" + k] = r; vals = {}
    hn, h1, h27 = H_("new_" + k, "G"), H_("new_" + k, "H"), H_("new_" + k, "I")
    vals["26"] = f"={hn}"; vals["1H"] = f"={h27}"; vals["2H"] = f"=({hn}-{h1})*(1+h2_pers*({h27}/{h1}-1))"; vals[27] = f"=F{r}+G{r}"
    if s in HOLD:
        vals["24"] = f"=Hist!E{HOLD[s]}"; vals["25"] = f"=Hist!F{HOLD[s]}"
    for i, c in enumerate(FCOL):
        pc = "H" if i == 0 else FCOL[i - 1]
        if k in dgk: vals[c] = f"={pc}{r}*(1+{c}{TR['gv']}+{dgk[k]})*(1+pg)"
        elif k == "itp": vals[c] = f"={pc}{r}*(1+{c}{TR['gi']})"
        else: vals[c] = f"={pc}{r}"
    put(TH, r, f"AI 前外部營收：{s}", "¥B", vals, "#,##0.000", "FY2024A、FY2025A 為舊口徑（澳洲在 Industrial IT）；FY2026A 起新口徑；2H27E＝2H26 重編 ×(1＋延續比例 × 1H 年增)" if k == "cons" else None); r += 1
TR["pool"] = r; put(TH, r, "人力計價分部合計（Consulting＋Financial IT＋Industrial IT；AI 前）", "¥B", {c: f"={c}{TR['Scons']}+{c}{TR['Sfin']}+{c}{TR['Sind']}" for c in ["26", 27] + list(FCOL)}, "#,##0.000", bold=True); r += 1
TR["P"] = r; put(TH, r, "單價指數（AI 前）", "倍", {27: 1, **{c: f"=(1+pg)^{c}{TR['t']}" for c in FCOL}}, "0.0000"); TH[f"H{r}"] = "=1"; r += 1
TR["pp"] = r; put(TH, r, "AI 前實現小時單價", "日圓", {c: f"=rate0*{c}{TR['P']}" for c in [27] + list(FCOL)}, "#,##0"); r += 1
TR["hp"] = r; put(TH, r, "AI 前需求工時", "百萬小時", {c: f"={c}{TR['pool']}*1000/{c}{TR['pp']}" for c in [27] + list(FCOL)}, "#,##0.00"); r += 1
TR["Np"] = r; put(TH, r, "AI 前所需交付人員（隱含）", "人", {c: f"={c}{TR['hp']}*1000000/(util*hrs_y)" for c in [27] + list(FCOL)}, "#,##0"); r += 2
SC = {}
for X in "ABC":
    head(TH, r, f"情境 {X}：AI 單位經濟引擎（人力計價分部合計）"); r += 1; E = {}
    def rw(k, label, unit, f, fmt="#,##0.000", note=None, first=27):
        global r
        E[k] = r; put(TH, r, label, unit, {c: f(c) for c in ([27] if first == 27 else []) + list(FCOL)}, fmt, note); r += 1
    pool, tt = TR["pool"], TR["t"]
    rw("auto", "AI 增量自動化比例", "比例", lambda c: f"=auto_{X}*{c}{tt}/5", "0.0%")
    rw("prj", "AI 新專案營收（增量）", "¥B", lambda c: f"=proj_{X}*{c}{tt}/5*{c}{pool}")
    rw("ero", "價格侵蝕（客戶分享生產力＝AI 前營收 × 自動化 ×（1 − 價格保留率））", "¥B", lambda c: f"={c}{pool}*{c}{E['auto']}*(1-ret_{X})")
    rw("rev", "人力計價分部營收（含 AI 新專案）", "¥B", lambda c: f"={c}{pool}-{c}{E['ero']}+{c}{E['prj']}", note="＝AI 前營收 − 侵蝕 ＋ AI 新專案")
    rw("f", "營收因子（÷ AI 前）", "倍", lambda c: f"={c}{E['rev']}/{c}{pool}", "0.0000")
    rw("bh", "計費工時（AI 後；不含 AI 新專案）", "百萬小時", lambda c: f"={c}{TR['hp']}*(1-{c}{E['auto']})", "#,##0.00")
    rw("pe", "實現每計費小時單價（AI 後）", "日圓", lambda c: f"=({c}{pool}-{c}{E['ero']})*1000/{c}{E['bh']}", "#,##0")
    rw("N", "所需交付人員（AI 後）", "人", lambda c: f"={c}{E['bh']}*1000000/(util*hrs_y)", "#,##0")
    rw("id", "恆等式核對：營收 − 人數 × 利用率 × 工時 × 單價 − AI 新專案（應為 0）", "¥B", lambda c: f"={c}{E['rev']}-{c}{E['prj']}-{c}{E['N']}*util*hrs_y*{c}{E['pe']}/1000000000", "0.000000")
    rw("sav", "省下人力成本（省下人數 × 每人年成本 × 單價指數 × 實現率）", "¥B", lambda c: f"=({c}{TR['Np']}-{c}{E['N']})*c_head*{c}{TR['P']}*sav_r/1000")
    rw("inf", "推論成本（取代工時 ÷ 任務長度 × 每任務成本）", "¥B", lambda c: f"={c}{TR['hp']}*{c}{E['auto']}/{T('len')}*$C${TKR['task']}/1000", "#,##0.0000", "G2-1：NRI 未揭露任務量，任務數由取代工時推得（Derived）；每任務成本 Assumed 倍數")
    rw("net", "AI 淨效果（營收口徑）＝AI 新專案增量 − 價格侵蝕 − 推論成本", "¥B", lambda c: f"={c}{E['prj']}-{c}{E['ero']}-{c}{E['inf']}")
    rw("aiop", "AI 營業利益效果＝−侵蝕 ＋ AI 新專案 × 利益率 ＋ 省下人力 − 推論成本", "¥B", lambda c: f"=-{c}{E['ero']}+{c}{E['prj']}*proj_m+{c}{E['sav']}-{c}{E['inf']}", note="進 PL 分部營業利益（依 AI 前營收比分配）")
    rw("be", "損益兩平所需價格保留率（營業利益效果＝0；公式反解）", "比例", lambda c: f'=IF({c}{E["auto"]}>0,1-({c}{E["prj"]}*proj_m+{c}{E["sav"]}-{c}{E["inf"]})/({c}{pool}*{c}{E["auto"]}),"—")', "0.0%")
    rw("ph", "每交付人員：營收（不含 AI 新專案）", "¥M／人", lambda c: f"=({c}{pool}-{c}{E['ero']})*1000/{c}{E['N']}", "#,##0.00")
    rw("pnet", "每交付人員：AI 淨效果（營收口徑）", "¥M／人", lambda c: f"={c}{E['net']}*1000/{c}{E['N']}", "#,##0.000")
    rw("pinf", "每交付人員：推論成本", "¥M／人", lambda c: f"={c}{E['inf']}*1000/{c}{E['N']}", "#,##0.0000")
    SC[X] = E; r += 1
TH.column_dimensions["A"].width = 70
# ===================================== PL
widths(PL)
PL["A1"] = "PL — 分部營收（外部客戶）→ 分部營業利益 → 合併損益（¥B；三情境）。分部＝公司官方報告分部（2026-04 起新口徑）；歷史 FY2024A、FY2025A 為舊口徑。營業利益由分部加總；FY2027E＝1H27A＋2H27E。"
PL["A2"] = "歷史期營收合計核對對 Hist 頁公司揭露合併營收（差額為外部營收摘要的四捨五入，≤ 0.005）；FY2026A 分部營業利益為舊口徑（新口徑重編值未找到）。GAAP 口徑（IFRS）。"
PR = {}; r = 4
for X in "ABC":
    pre = "" if X == "B" else f"情境 {X}｜"
    head(PL, r, f"情境 {X}" + ("（基準）" if X == "B" else "")); hrow_ = r; r += 1; Rr = {"hdr": hrow_}; E = SC[X]
    def Hh(s, col): return f"Hist!{col}{HNEW[s]}"
    seg_r = {}
    order = [("Consulting", None), ("Financial IT Solutions", SUBF), ("Industrial IT Solutions", SUBI), ("IT Platform Services", None), ("Others／Adjustments", None)]
    for s, subs in order:
        k = SK[s]; seg_r[s] = r; vals = {}
        for p, hc in (("24", "E"), ("25", "F")):
            if s in HOLD: vals[p] = f"=Hist!{hc}{HOLD[s]}"
        vals["26"] = f"=Hist!G{HNEW[s]}"
        for p in ["1H", "2H", 27]: vals[p] = f"=Thesis!{COL[p]}{TR['S'+k]}"
        for c in FCOL: vals[c] = f"=Thesis!{c}{TR['S'+k]}" + (f"*Thesis!{c}{E['f']}" if k in dgk else "")
        put(PL, r, pre + s if pre else s, "¥B", vals, "#,##0.000", "外部客戶營收" if s == "Consulting" else None); r += 1
        if subs:
            tot = sum(FN.H1SUB27[x] for x in subs)
            for x in subs:
                sv = {"26": f"=Hist!G{HSUB[x]}", "1H": f"=Hist!I{HSUB[x]}"}
                for p in ["2H", 27] + list(FCOL): sv[p] = f"={COL[p]}{seg_r[s]}*Hist!$I${HSUB[x]}/SUM(Hist!$I${HSUB[subs[0]]}:$I${HSUB[subs[-1]]})"
                sv[27] = f"=F{r}+G{r}"
                put(PL, r, pre + x if pre else x, "¥B", sv, "#,##0.000", "產品線（分部下產業別）：占比＝1H27 揭露占比，全期持平"); r += 1
            Rr["sub_" + k] = (r - len(subs), r - 1)
    Rr["seg"] = seg_r
    Rr["rev"] = r
    vals = {"24": f"=Hist!E{HR['rev']}", "25": f"=Hist!F{HR['rev']}", "26": f"=Hist!G{HR['rev']}", "1H": f"=Hist!I{HR['rev']}"}
    for p in ["2H", 27] + list(FCOL): vals[p] = "=" + "+".join(f"{COL[p]}{seg_r[s]}" for s in SEGS)
    put(PL, r, pre + "合併營收", "¥B", vals, "#,##0.000", bold=True); r += 1
    Rr["chk"] = r; put(PL, r, f"檢查：分部合計 − 合併營收（情境 {X}）", "¥B", {p: "=" + "+".join(f"{COL[p]}{seg_r[s]}" for s in SEGS) + f"-{COL[p]}{Rr['rev']}" for p in YK}, "0.000"); r += 1
    chk2 = []
    for s, subs in order:
        if subs: a, b = Rr["sub_" + SK[s]]; chk2.append(f"SUM({{c}}{a}:{{c}}{b})-{{c}}{seg_r[s]}")
    Rr["chk2"] = r; put(PL, r, f"檢查：產品線合計 − 分部（情境 {X}）", "¥B", {p: "=" + "+".join(x.format(c=COL[p]) for x in chk2) for p in ["26", "1H", "2H", 27] + list(FCOL)}, "0.000", "FY2026A 對公司揭露的產品線外部營收（四捨五入差 ≤ 0.005）"); r += 1
    put(PL, r, pre + "營收年增率", "%", {p: f"={COL[p]}{Rr['rev']}/{COL[q_]}{Rr['rev']}-1" for p, q_ in [("25", "24"), ("26", "25"), (27, "26"), (28, 27), (29, 28), (30, 29), (31, 30), (32, 31)]}, "0.0%"); Rr["yoy"] = r; r += 1
    oi_r = {}
    for s, subs in order:
        k = SK[s]; oi_r[s] = r; vals = {}
        for p, hc in (("24", "E"), ("25", "F"), ("26", "G"), ("1H", "I")): vals[p] = f"=Hist!{hc}{HOI[s]}"
        e27 = seg_r[s]
        if s == "Others／Adjustments":
            vals["2H"] = "=adj_oi/2"
        else:
            vals["2H"] = f"=rt_{k}*G{e27}*mg_{k}*seas"
        vals[27] = f"=F{r}+G{r}"
        for c in FCOL:
            if s == "Others／Adjustments": vals[c] = "=adj_oi"
            elif k in dgk: vals[c] = f"=rt_{k}*Thesis!{c}{TR['S'+k]}*mg_{k}+Thesis!{c}{E['aiop']}*Thesis!{c}{TR['S'+k]}/Thesis!{c}{TR['pool']}"
            else: vals[c] = f"=rt_{k}*{c}{e27}*mg_{k}"
        put(PL, r, pre + "營業利益：" + s, "¥B", vals, "#,##0.000", "舊口徑含內部交易口徑；FY2026A Industrial 含減損 97.060" if s == "Consulting" else None); r += 1
    Rr["oi_s"] = oi_r
    Rr["oi"] = r; put(PL, r, pre + "合併營業利益", "¥B", {p: "=" + "+".join(f"{COL[p]}{oi_r[s]}" for s in SEGS) for p in YK}, "#,##0.000", bold=True); r += 1
    Rr["chk3"] = r; put(PL, r, f"檢查：分部營業利益合計 − 合併營業利益（情境 {X}；歷史對 Hist 公司值）", "¥B", {"24": f"={COL['24']}{Rr['oi']}-Hist!E{HR['oi']}", "25": f"=D{Rr['oi']}-Hist!F{HR['oi']}", "26": f"=E{Rr['oi']}-Hist!G{HR['oi']}", "1H": f"=F{Rr['oi']}-Hist!I{HR['oi']}"}, "0.000"); r += 1
    Rr["opm"] = r; put(PL, r, pre + "營業利益率", "%", {p: f"={COL[p]}{Rr['oi']}/{COL[p]}{Rr['rev']}" for p in YK}, "0.0%"); r += 1
    Rr["fin"] = r; put(PL, r, pre + "財務收益淨額", "¥B", {"24": f"=Hist!E{HR['pbt']}-Hist!E{HR['oi']}", "25": f"=Hist!F{HR['pbt']}-Hist!F{HR['oi']}", "26": f"=Hist!G{HR['pbt']}-Hist!G{HR['oi']}", "1H": f"=Hist!I{HR['fin_i']}-Hist!I{HR['fin_c']}", "2H": "=fin_net/2", 27: f"=F{r}+G{r}", **{c: "=fin_net" for c in FCOL}}, "#,##0.000"); r += 1
    Rr["pbt"] = r; put(PL, r, pre + "稅前利益", "¥B", {p: f"={COL[p]}{Rr['oi']}+{COL[p]}{Rr['fin']}" for p in YK}, "#,##0.000"); r += 1
    Rr["tax"] = r; put(PL, r, pre + "所得稅", "¥B", {"24": f"=Hist!E{HR['tax']}", "25": f"=Hist!F{HR['tax']}", "26": f"=Hist!G{HR['tax']}", "1H": f"=Hist!I{HR['tax']}", "2H": f"=tax*G{Rr['pbt']}", 27: f"=F{r}+G{r}", **{c: f"=tax*{c}{Rr['pbt']}" for c in FCOL}}, "#,##0.000"); r += 1
    Rr["pf"] = r; put(PL, r, pre + "當期利益", "¥B", {p: f"={COL[p]}{Rr['pbt']}-{COL[p]}{Rr['tax']}" for p in YK}, "#,##0.000"); r += 1
    Rr["mi"] = r; put(PL, r, pre + "非控制權益損益", "¥B", {"24": f"=C{Rr['pf']}-Hist!E{HR['ni']}", "25": f"=D{Rr['pf']}-Hist!F{HR['ni']}", "26": f"=E{Rr['pf']}-Hist!G{HR['ni']}", "1H": f"=F{Rr['pf']}-Hist!I{HR['ni']}", "2H": f"=mi_r*G{Rr['pf']}", 27: f"=F{r}+G{r}", **{c: f"=mi_r*{c}{Rr['pf']}" for c in FCOL}}, "#,##0.000"); r += 1
    Rr["ni"] = r; put(PL, r, pre + "歸屬母公司淨利", "¥B", {p: f"={COL[p]}{Rr['pf']}-{COL[p]}{Rr['mi']}" for p in YK}, "#,##0.000", bold=True); r += 1
    Rr["bbt"] = r; put(PL, r, pre + "庫藏股取得", "¥B", {"24": f"=Hist!E{HR['bb']}", "25": f"=Hist!F{HR['bb']}", "26": f"=Hist!G{HR['bb']}", 27: "=bb_27", **{c: "=bb_t" for c in FCOL}}, "#,##0.0"); r += 1
    Rr["she"] = r; put(PL, r, pre + "期末流通股數", "百萬股", {"26": f"=Hist!G{HR['sh_iss']}-Hist!G{HR['sh_tre']}", 27: "=sh_out", **{c: f"={'H' if c=='I' else chr(ord(c)-1)}{r}-{c}{Rr['bbt']}*1000/(px*(1+bb_pg)^{c}{TR['t']+0 if False else 0}+0)" for c in FCOL}}, "#,##0.000"); r += 1
    for i, c in enumerate(FCOL):   # 買進價＝px*(1+bb_pg)^t
        PL[f"{c}{Rr['she']}"] = f"={'H' if c=='I' else FCOL[i-1]}{Rr['she']}-{c}{Rr['bbt']}*1000/(px*(1+bb_pg)^Thesis!{c}{TR['t']})"
    Rr["sha"] = r; put(PL, r, pre + "加權平均股數", "百萬股", {"24": "=Hist!E%d" % HR["sh_iss"] + f"-Hist!E{HR['sh_tre']}", "1H": f"=Hist!I{HR['sh_avg']}", "2H": "=sh_out", 27: f"=(F{r}+G{r})/2", **{c: f"=({'H' if c=='I' else FCOL[FCOL.index(c)-1]}{Rr['she']}+{c}{Rr['she']})/2" for c in FCOL}}, "#,##0.000", "歷史 EPS 取公司揭露值"); r += 1
    PL[f"C{Rr['sha']}"] = None
    Rr["eps"] = r; put(PL, r, pre + "EPS（日圓）", "日圓", {"24": f"=Hist!E{HR['eps']}", "25": f"=Hist!F{HR['eps']}", "26": f"=Hist!G{HR['eps']}", "1H": f"=Hist!I{HR['eps']}", "2H": f"=G{Rr['ni']}*1000/G{Rr['sha']}", 27: f"=H{Rr['ni']}*1000/H{Rr['sha']}", **{c: f"={c}{Rr['ni']}*1000/{c}{Rr['sha']}" for c in FCOL}}, "#,##0.00", bold=True); r += 1
    Rr["dps"] = r; put(PL, r, pre + "每股股利（配息率 × EPS）", "日圓", {"24": f"=Hist!E{HR['dps']}", "25": f"=Hist!F{HR['dps']}", "26": f"=Hist!G{HR['dps']}", 27: f"=payout*H{Rr['eps']}", **{c: f"=payout*{c}{Rr['eps']}" for c in FCOL}}, "#,##0.0"); r += 1
    PR[X] = Rr; r += 1
# 標準損益表（情境 B）
B = PR["B"]; SC_ = {"24": "C", "25": "D", "26": "E", 27: "H", 28: "I", 29: "J", 30: "K", 31: "L", 32: "M"}
IS = {}; ISC = dict(zip(["24", "25", "26", 27, 28, 29, 30, 31, 32], "CDEFGHIJK"))
PL.cell(r, 1, "標準損益表（GAAP；IFRS 口徑；¥B，EPS 日圓）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B（基準）。採 IFRS（GAAP 口徑）因 NRI 只揭露 IFRS 合併數；預測期營業成本、毛利由營業費用率（1H27 實績持平）推算：營業利益來自分部加總，營業費用＝費用率 × 營收，毛利＝營業利益＋營業費用，營業成本為殘差（Assumed；不改營業利益與淨利）。FY2026A 營業費用含減損 97.586。").font = BOLD; IS["note"] = r; r += 1
IS["hdr"] = r; hdr(PL, r, ["項目", "單位", "FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]); r += 1
def isr(key, label, unit, f, fmt="#,##0.0"):
    global r
    IS[key] = r; PL.cell(r, 1, label); PL.cell(r, 2, unit)
    for p, c in ISC.items(): x = PL[f"{c}{r}"]; x.value = f(p, c, SC_[p]); x.number_format = fmt
    r += 1
hist_ = lambda k: (lambda p, c, m: f"=Hist!{ {'24':'E','25':'F','26':'G'}[p] }{HR[k]}" if p in ("24", "25", "26") else None)
isr("rev", "營收", "¥B", lambda p, c, m: f"={m}{B['rev']}")
isr("yoy", "營收年增率", "%", lambda p, c, m: (f"={c}{IS['rev']}/Hist!{ {'24':'D','25':'E','26':'F'}[p] }{HR['rev']}-1" if p == "24" else f"={c}{IS['rev']}/{chr(ord(c)-1)}{IS['rev']}-1"), "0.0%")
isr("cogs", "營業成本", "¥B", lambda p, c, m: (f"=Hist!{ {'24':'E','25':'F','26':'G'}[p] }{HR['cogs']}" if p in ("24", "25", "26") else f"={c}{IS['rev']}-{c}{IS['rev']+3}"))
isr("gp", "毛利", "¥B", lambda p, c, m: (f"=Hist!{ {'24':'E','25':'F','26':'G'}[p] }{HR['gp']}" if p in ("24", "25", "26") else f"={m}{B['oi']}+{c}{IS['rev']+5}"))
isr("gm", "毛利率", "%", lambda p, c, m: f"={c}{IS['gp']}/{c}{IS['rev']}", "0.0%")
isr("opx", "營業費用（含其他損益淨額）", "¥B", lambda p, c, m: (f"={c}{IS['gp']}-{c}{IS['gp']+3}" if p in ("24", "25", "26") else f"=sga_r*{c}{IS['rev']}"))
isr("oi", "營業利益", "¥B", lambda p, c, m: f"={m}{B['oi']}")
isr("opm", "營業利益率", "%", lambda p, c, m: f"={c}{IS['oi']}/{c}{IS['rev']}", "0.0%")
isr("oth", "其他收入（淨額；財務收益淨額）", "¥B", lambda p, c, m: f"={m}{B['fin']}")
isr("pbt", "稅前利益", "¥B", lambda p, c, m: f"={m}{B['pbt']}")
isr("tax", "所得稅", "¥B", lambda p, c, m: f"={m}{B['tax']}")
isr("ni", "稅後淨利（歸屬母公司）", "¥B", lambda p, c, m: f"={m}{B['ni']}")
isr("nm", "淨利率", "%", lambda p, c, m: f"={c}{IS['ni']}/{c}{IS['rev']}", "0.0%")
isr("eps", "EPS（完全稀釋，日圓；NRI 稀釋與基本相同）", "日圓", lambda p, c, m: f"={m}{B['eps']}", "#,##0.00")
PL.cell(r, 1, "比率區").font = BOLD; IS["ratio"] = r; r += 1
isr("r_yoy", "營收年增率", "%", lambda p, c, m: f"={c}{IS['yoy']}", "0.0%"); isr("r_gm", "毛利率", "%", lambda p, c, m: f"={c}{IS['gm']}", "0.0%"); isr("r_opm", "營業利益率", "%", lambda p, c, m: f"={c}{IS['opm']}", "0.0%")
isr("r_nm", "淨利率", "%", lambda p, c, m: f"={c}{IS['nm']}", "0.0%")
isr("r_niy", "淨利年增率", "%", lambda p, c, m: (f"={c}{IS['ni']}/Hist!E{HR['ni']}-1" if False else (f"={c}{IS['ni']}/Hist!D{HR['ni']}-1" if p == "24" else f"={c}{IS['ni']}/{chr(ord(c)-1)}{IS['ni']}-1")), "0.0%")
PL.cell(r + 1, 1, "核對（應為零）").font = BOLD; IS["chk"] = r + 1; r += 2
isr("k1", "營收 − 營業成本 − 毛利", "¥B", lambda p, c, m: f"={c}{IS['rev']}-{c}{IS['cogs']}-{c}{IS['gp']}", "0.000")
isr("k2", "毛利 − 營業費用 − 營業利益", "¥B", lambda p, c, m: f"={c}{IS['gp']}-{c}{IS['opx']}-{c}{IS['oi']}", "0.000")
isr("k3", "稅前 − 所得稅 − 少數股東 − 淨利（FY2024A–FY2026A 對公司揭露）", "¥B", lambda p, c, m: f"={c}{IS['pbt']}-{c}{IS['tax']}-{m}{B['mi']}-{c}{IS['ni']}", "0.000")
PR["IS"] = IS
# ===================================== CF
widths(CF)
CF["A1"] = "CF — 現金流與收支關卡（¥B；三情境）。自由現金流＝當期利益＋折舊攤銷（不含使用權資產）− 資本支出（現金）− 營運資金增加；使用權資產折舊視為營業費用（租賃作營業成本，淨債務不另加租賃，G2-5）。"
r = 4; CR = {}
for X in "ABC":
    head(CF, r, f"情境 {X}" + ("（基準）" if X == "B" else "")); r += 1; Rr = {}; p_ = PR[X]
    cols = [27] + list(FCOL)
    def cr(k, label, f, fmt="#,##0.0", note=None, bold=False):
        global r
        Rr[k] = r; put(CF, r, label, "¥B", {c: f(c) for c in cols}, fmt, note, bold); r += 1
    prv = lambda c: ("H" if c == "I" else chr(ord(c) - 1))
    cr("ni", "當期利益", lambda c: f"=PL!{COL[c]}{p_['pf']}")
    cr("da", "折舊攤銷（不含使用權資產）", lambda c: f"=da_r*PL!{COL[c]}{p_['rev']}")
    cr("cx", "資本支出（現金，有形＋無形）", lambda c: f"=capex_r*PL!{COL[c]}{p_['rev']}")
    cr("wc", "營運資金增加", lambda c: f"=wc_k*(PL!{COL[c]}{p_['rev']}-" + (f"{H_('rev','G')})" if c == 27 else f"PL!{prv(c)}{p_['rev']})"))
    cr("fcf", "自由現金流（租賃作營業成本口徑）", lambda c: f"={c}{Rr['ni']}+{c}{Rr['da']}-{c}{Rr['cx']}-{c}{Rr['wc']}", bold=True)
    cr("div", "股利", lambda c: f"=payout*PL!{COL[c]}{p_['ni']}")
    cr("fcfd", "自由現金流扣股利", lambda c: f"={c}{Rr['fcf']}-{c}{Rr['div']}", bold=True)
    cr("bb", "庫藏股", lambda c: f"=PL!{COL[c]}{p_['bbt']}")
    cr("cend", "期末現金（FY2026 實績起滾動；債務持平）", lambda c: (f"={H_('cash','G')}" if c == 27 else f"={prv(c)}{r}") + f"+{c}{Rr['fcfd']}-{c}{Rr['bb']}")
    CF[f"H{Rr['cend']}"] = f"={H_('cash','G')}+H{Rr['fcfd']}-H{Rr['bb']}"
    cr("nc", "淨現金（現金 − 有利子負債 2026-09-30）", lambda c: f"={c}{Rr['cend']}-debt_0")
    CR[X] = Rr; r += 1
GATE = r
CF.cell(r, 1, "收支關卡：FY2027E–FY2031E 自由現金流扣股利合計（A／B／C，¥B）").font = BOLD
for j, X in enumerate("ABC"): c = CF.cell(r, 2 + j, f"=SUM(H{CR[X]['fcfd']}:L{CR[X]['fcfd']})"); c.number_format = "#,##0.0"
r += 1
CF.cell(r, 1, "最低年末現金（FY2027E–FY2031E，A／B／C）")
for j, X in enumerate("ABC"): CF.cell(r, 2 + j, f"=MIN(H{CR[X]['cend']}:L{CR[X]['cend']})").number_format = "#,##0.0"
GATE_MIN = r; r += 1
CF.cell(r, 1, "檢查：資本支出 ≥ 折舊 ×(1＋g)（稅後永續年）"); CF.cell(r, 2, '=IF(capex_r>=da_r*(1+MIN(g_tv,rf)),"通過","不通過")'); CHK_CX = r
# ===================================== Val
V["A1"] = "Val — 估值：DCF（FCFF，明確期到 FY2032，期中折現，評價日 2026-09-30）與 Comparable（同業前瞻 P/E 中位數 × NTM EPS）各半；情境權重 25／50／25（In 頁）。目標價為股價基準日起 12 個月。"
V["A2"] = "Comparable 不做分部加總：NRI 無分部 EBITDA 揭露，且同業無分部可比倍數（原則第 4 節，README 寫明）；本版以整家公司前瞻 P/E，分子分母同為股權價值與歸屬母公司淨利（G2-6）。"
VR = {}; r = 4
def vrow(k, label, b=None, c=None, d=None, fmt="#,##0.00", note=None, bold=False):
    global r
    VR[k] = r; V.cell(r, 1, label)
    if bold: V.cell(r, 1).font = BOLD
    for col, v in (("B", b), ("C", c), ("D", d)):
        if v is not None: V[f"{col}{r}"] = v; V[f"{col}{r}"].number_format = fmt
    if note: V.cell(r, 5, note)
    r += 1
hdr(V, r, ["折現率與資本結構", "數值", "", "", "說明"]); r += 1
vrow("emkt", "股權市值（現價 × 流通股數，¥B）", "=px*sh_out/1000", fmt="#,##0.0")
vrow("dmkt", "債務（有利子負債＋另加租賃，¥B）", "=debt_0+lease_v", fmt="#,##0.0")
vrow("ke", "股權成本 ke＝rf＋β×ERP", "=rf+beta*erp", fmt="0.00%")
vrow("kdat", "稅後負債成本", "=kd*(1-tax)", fmt="0.00%")
vrow("g", "永續成長率 g＝MIN(g 上限, rf)", "=MIN(g_tv,rf)", fmt="0.00%")
vrow("wacc", "WACC", f"=(B{VR['emkt']}*B{VR['ke']}+B{VR['dmkt']}*B{VR['kdat']})/(B{VR['emkt']}+B{VR['dmkt']})+wacc_adj", fmt="0.00%")
vrow("ronic", "RONIC", f"=B{VR['wacc']}+ronic_sp", fmt="0.00%")
vrow("stub", "2H27E 期間（年）＝2027-03-31 − 評價日", "=(DATE(2027,3,31)-val_date)/365", fmt="0.000")
vrow("tau", "評價日至目標時點（年）", "=(tp_date-val_date)/365", fmt="0.000")
vrow("dps_t", "持有期間每股股利（FY2027 期末＋FY2028 中間，日圓；FY2027 DPS − 已除權中間股利＋½ FY2028 DPS）", f"=PL!H{PR['B']['dps']}-dps_int+0.5*PL!I{PR['B']['dps']}", fmt="#,##0.0", note="情境 B 的 DPS")
r += 1
FC_ = {}
for X in "ABC":
    hdr(V, r, [f"DCF 情境 {X}（¥B）", "", "", "", "", "2H27E", "", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "終值"]); r += 1; p_ = PR[X]; Rr = {}
    cols = ["G", "I", "J", "K", "L", "M"]
    def fr(k, label, f, fmt="#,##0.0", tv=None):
        global r
        Rr[k] = r; V.cell(r, 1, label)
        for c in cols: V[f"{c}{r}"] = f(c); V[f"{c}{r}"].number_format = fmt
        if tv: V[f"N{r}"] = tv; V[f"N{r}"].number_format = fmt
        r += 1
    prv = lambda c: ("H" if c == "I" else chr(ord(c) - 1))
    fr("rev", "營收", lambda c: f"=PL!{c}{p_['rev']}")
    fr("ebit", "營業利益（EBIT）", lambda c: f"=PL!{c}{p_['oi']}")
    fr("nopat", "NOPAT＝EBIT×(1−稅率)", lambda c: f"={c}{Rr['ebit']}*(1-tax)", tv=f"=M{r}*(1+$B${VR['g']})")
    fr("da", "＋折舊攤銷（不含使用權資產）", lambda c: f"=da_r*{c}{Rr['rev']}")
    fr("cx", "−資本支出", lambda c: f"=capex_r*{c}{Rr['rev']}")
    fr("wc", "−營運資金增加", lambda c: f"=wc_k*({c}{Rr['rev']}-" + (f"({H_('rev','G')}-{H_('rev','H')}))" if c == "G" else f"{'PL!H'+str(p_['rev']) if c=='I' else prv(c)+str(Rr['rev'])})"))
    fr("fcff", "FCFF", lambda c: f"={c}{Rr['nopat']}+{c}{Rr['da']}-{c}{Rr['cx']}-{c}{Rr['wc']}", tv=f"=N{Rr['nopat']}*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})")
    fr("t", "折現年數（期中）", lambda c: {"G": f"=$B${VR['stub']}/2", "I": f"=$B${VR['stub']}+0.5", "J": f"=$B${VR['stub']}+1.5", "K": f"=$B${VR['stub']}+2.5", "L": f"=$B${VR['stub']}+3.5", "M": f"=$B${VR['stub']}+4.5"}[c], "0.000", tv=f"=$B${VR['stub']}+5")
    fr("pv", "現值", lambda c: f"={c}{Rr['fcff']}/(1+$B${VR['wacc']})^{c}{Rr['t']}", tv=f"=N{Rr['fcff']}/(1+$B${VR['wacc']})^N{Rr['t']}")
    fr("pvn", "其中 NOPAT 部分現值（反向 DCF 用）", lambda c: f"={c}{Rr['nopat']}/(1+$B${VR['wacc']})^{c}{Rr['t']}", tv=f"=N{Rr['nopat']}*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^N{Rr['t']}")
    Rr["ev"] = r; V.cell(r, 1, f"企業價值 EV（情境 {X}）"); V[f"B{r}"] = f"=SUM(G{Rr['pv']}:N{Rr['pv']})"; V[f"B{r}"].number_format = "#,##0.0"; r += 1
    Rr["tvs"] = r; V.cell(r, 1, "終值占 EV"); V[f"B{r}"] = f"=N{Rr['pv']}/B{Rr['ev']}"; V[f"B{r}"].number_format = "0.0%"; r += 1
    Rr["eq"] = r; V.cell(r, 1, "股權價值＝EV＋現金＋其他流動金融資產×比例−有利子負債−另加租賃−非控制權益"); V[f"B{r}"] = f"=B{Rr['ev']}+cash_0+ofa_k*{H_('b_ofa','I')}-debt_0-lease_v-mi_0"; V[f"B{r}"].number_format = "#,##0.0"; r += 1
    Rr["v0"] = r; V.cell(r, 1, "每股價值（評價日，日圓）"); V[f"B{r}"] = f"=B{Rr['eq']}*1000/sh_out"; V[f"B{r}"].number_format = "#,##0"; r += 1
    Rr["vt"] = r; V.cell(r, 1, f"DCF 情境 {X} 每股價值推到目標時點（×(1+ke)^期間 − 股利；0 為下限）"); V[f"B{r}"] = f"=MAX(0,B{Rr['v0']}*(1+$B${VR['ke']})^$B${VR['tau']}-$B${VR['dps_t']})"; V[f"B{r}"].number_format = "#,##0"; r += 1
    FC_[X] = Rr; r += 1
hdr(V, r, ["結論", "數值", "情境 A", "情境 B", "說明"]); r += 1
vrow("px", "現價（日圓）", "=px", fmt="#,##0", note=None)
V.cell(VR["px"], 5, '="現價日 "&TEXT(px_date,"yyyy-mm-dd")&"（逾期標示見 README 市場資料日期）"')
vrow("tpdcf", "DCF 目標價（情境加權）", f"=tp_wA*B{FC_['A']['vt']}+tp_wB*B{FC_['B']['vt']}+tp_wC*B{FC_['C']['vt']}", f"=B{FC_['A']['vt']}", f"=B{FC_['C']['vt']}", fmt="#,##0", note="情境 A（C 欄）、情境 C（D 欄）；情境 B 見上方")
vrow("om", "NTM EPS 權重 ω（FY2028 占比：FY2028 年底 − 目標時點）", "=(DATE(2028,3,31)-tp_date)/365", fmt="0.000")
for X in "ABC": vrow("ntm" + X, f"NTM EPS（目標時點起 12 個月；情境 {X}）", f"=$B${VR['om']}*PL!I{PR[X]['eps']}+(1-$B${VR['om']})*PL!J{PR[X]['eps']}", fmt="#,##0.00", note="分母：NTM EPS＝FY2028E 與 FY2029E 依月數加權（歸屬母公司淨利 ÷ 加權平均股數）")
vrow("comp_raw", "Comparable 目標價（市場主流同業法；同業前瞻 P/E 中位數 × NTM EPS，情境加權）", f"=pe_mult*(tp_wA*B{VR['ntmA']}+tp_wB*B{VR['ntmB']}+tp_wC*B{VR['ntmC']})", fmt="#,##0", note="倍數分母：同業前瞻（NTM）EPS，資料商；估值日 09-11～11-16 不一致")
vrow("tpcomp", "Comparable 目標價（相對倍數調整後）", f"=B{VR['comp_raw']}*rm_adj", fmt="#,##0", note="係數 1.00（未調整：缺 NRI 與同業十年前瞻倍數）")
vrow("tp", "目標價（DCF 與 Comparable 加權）", f"=w_dcf*B{VR['tpdcf']}+(1-w_dcf)*B{VR['tpcomp']}", fmt="#,##0", bold=True)
vrow("up", "上檔空間：加權目標價", f"=B{VR['tp']}/px-1", fmt="+0.0%;-0.0%", note="目標價為股價基準日起 12 個月；折現期間自評價日起（見上方），上檔空間不是年化報酬")
vrow("upd", "上檔空間：DCF", f"=B{VR['tpdcf']}/px-1", fmt="+0.0%;-0.0%")
vrow("upc", "上檔空間：Comparable", f"=B{VR['tpcomp']}/px-1", fmt="+0.0%;-0.0%")
vrow("tier", "估值評等（只看上檔空間）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))', fmt="@", note="級距：> +50%、> +25%、≥ −25%、≥ −50%（原則第 10 節第 24 條）")
vrow("gate", "入不敷出判定用：收支關卡合計（情境 B，¥B）", f"=CF!C{GATE}", fmt="#,##0.0")
vrow("gate_t", "入不敷出判定", f'=IF(B{VR["gate"]}<0,"不通過","通過")', fmt="@")
vrow("rec", "投資建議", f'=IF(AND(B{VR["gate"]}<0,OR(B{VR["tier"]}="買進",B{VR["tier"]}="強力買進")),"收支不足",B{VR["tier"]})', fmt="@", bold=True)
vrow("cons", "市場共識目標價（14 家平均，S&P Global 經 StockAnalysis；2026-09-11，財報前，只作歷史對照）", P["tp_cons_avg"], fmt="#,##0")
V.cell(VR["cons"], 3, "家數"); V.cell(VR["cons"], 4, P["tp_cons_n"])
vrow("scen_up", "情境目標價 A／B／C（加權＝w×DCF＋(1−w)×Comparable）", None)
for j, X in enumerate("ABC"): V.cell(VR["scen_up"], 2 + j, f"=w_dcf*B{FC_[X]['vt']}+(1-w_dcf)*rm_adj*pe_mult*B{VR['ntm'+X]}").number_format = "#,##0"
vrow("pe_now", "本公司現行前瞻 P/E ÷ 同業中位數（現價 ÷ 情境 B NTM EPS ÷ 中位數；G2-6、第 47 條）", f"=px/B{VR['ntmB']}/pe_mult", fmt="0.00", note="以模型 NTM EPS 計；資料商 NRI 前瞻 P/E 25.78（2026-11-09）÷ 同業中位數另見 Recon 第 4 節")
vrow("gap_dcf_cmp", "DCF 與 Comparable 落差（日圓）＝DCF − Comparable", f"=B{VR['tpdcf']}-B{VR['tpcomp']}", fmt="#,##0")
vrow("gap_pe", "  其中：倍數差（DCF 隱含 NTM P/E − 同業中位數）× 情境 B NTM EPS", f"=(B{VR['tpdcf']}/B{VR['ntmB']}-pe_mult)*B{VR['ntmB']}", fmt="#,##0", note="DCF 隱含 NTM P/E 以目標價 ÷ NTM EPS；落差來源：折現率（β 偏低）、終值、倍數口徑；股票薪酬 NRI 影響小（未單獨揭露）、槓桿（淨負債 ÷ 股權市值小）、時點（DCF 自評價日、Comparable 為目標時點）")
vrow("impl_pe", "DCF 隱含 NTM P/E", f"=B{VR['tpdcf']}/B{VR['ntmB']}", fmt="0.0")
vrow("wdcf_sh", "加權目標價中 DCF 貢獻／Comparable 貢獻（¥/股）", f"=w_dcf*B{VR['tpdcf']}", f"=(1-w_dcf)*B{VR['tpcomp']}", fmt="#,##0")
for k in ("up", "upd", "upc"): pass
V.column_dimensions["A"].width = 70; V.column_dimensions["E"].width = 60
# ===================================== RevDCF
RV["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（情境 B；封閉解：FCFF 對營業利益線性，解營業利益縮放倍數使目標價＝現價）"
RV["A2"] = "解法：每股目標價＝MAX(0, [(λ×NOPAT 現值＋其餘現值＋淨現金)×1000÷股數]×(1+ke)^τ − 股利)；令其＝現價，λ 為所有預測年營業利益的共同縮放倍數（含終值）。"
RVR = {}; Bf = FC_["B"]; r = 4
def rv(k, label, f, fmt="#,##0.00"):
    global r
    RVR[k] = r; RV.cell(r, 1, label); RV[f"B{r}"] = f; RV[f"B{r}"].number_format = fmt; r += 1
rv("pn", "NOPAT 部分現值（含終值，¥B）", f"=SUM(Val!G{Bf['pvn']}:N{Bf['pvn']})", "#,##0.0")
rv("po", "其餘現值（折舊 − 資本支出 − 營運資金，¥B）", f"=SUM(Val!G{Bf['pv']}:M{Bf['pv']})-SUM(Val!G{Bf['pvn']}:M{Bf['pvn']})", "#,##0.0")
rv("nc", "淨現金與非營業項目（¥B）", f"=cash_0+ofa_k*{H_('b_ofa','I')}-debt_0-lease_v-mi_0", "#,##0.0")
rv("v0", "現價隱含的評價日每股價值（日圓）＝(現價＋股利)÷(1+ke)^τ", f"=(px+Val!B{VR['dps_t']})/(1+Val!B{VR['ke']})^Val!B{VR['tau']}", "#,##0")
rv("lam", "隱含營業利益縮放倍數 λ", f"=(B{RVR['v0']}*sh_out/1000-B{RVR['nc']}-B{RVR['po']})/B{RVR['pn']}", "0.000")
rv("m0", "模型 FY2032E 營業利益率（情境 B）", f"=PL!M{PR['B']['opm']}", "0.0%")
rv("m1", "現價隱含 FY2032E 營業利益率（＝模型 × λ）", f"=B{RVR['m0']}*B{RVR['lam']}", "0.0%")
rv("oi0", "模型 FY2032E 營業利益（¥B）", f"=PL!M{PR['B']['oi']}", "#,##0.0")
rv("oi1", "現價隱含 FY2032E 營業利益（¥B）", f"=B{RVR['oi0']}*B{RVR['lam']}", "#,##0.0")
rv("pe1", "現價 ÷ 情境 B NTM EPS（現行前瞻 P/E，模型口徑）", f"=px/Val!B{VR['ntmB']}", "0.0")
rv("pe2", "同業前瞻 P/E 中位數（資料商）", "=pe_mult", "0.0")
rv("pe3", "同業中位數隱含 NRI 股價 ÷ 現價（Comparable 單法）", f"=Val!B{VR['tpcomp']}/px", "0.00")
RV["A20"] = f'="現價隱含 FY2032E 營業利益率 "&TEXT(B{RVR["m1"]},"0.0%")&"，模型 "&TEXT(B{RVR["m0"]},"0.0%")&"；現價"&IF(B{RVR["lam"]}>1,"高於","低於")&"模型 DCF 所需的利潤（λ＝"&TEXT(B{RVR["lam"]},"0.00")&"）。"'
RV["A21"] = f'="現金流折現法：支持現價需要 FY2032E 營業利益 ¥"&TEXT(B{RVR["oi1"]},"#,##0")&"B（模型 ¥"&TEXT(B{RVR["oi0"]},"#,##0")&"B）；Comparable 以同業 P/E 中位數 "&TEXT(B{RVR["pe2"]},"0.0")&" 倍，現行前瞻 P/E "&TEXT(B{RVR["pe1"]},"0.0")&" 倍。"'
RV.column_dimensions["A"].width = 80
# ===================================== Val：同業（同口徑；G2-6）
r += 1
V.cell(r, 1, "同業前瞻 P/E（資料商；分母＝NTM 前瞻 EPS；分子分母同為股權價值對淨利，與本公司同口徑；逐家標會計準則與租賃處理）").font = BOLD; r += 1
hdr(V, r, ["同業", "代號", "前瞻 P/E", "EV/EBITDA", "資料日", "逾期", "會計準則", "分母期間與租賃處理"]); r += 1
PE_A = r
LEASE = {"IFRS": "IFRS 16：租賃資本化，EBITDA 已加回使用權資產折舊（EV 含租賃負債）；P/E 不受影響", "J-GAAP": "J-GAAP：營業租賃表外，EV／EBITDA 與 IFRS 同業不完全可比；P/E 不受影響", "US GAAP": "US GAAP：營業租賃費用於營業利益內，EV／EBITDA 與 IFRS 不完全可比；P/E 不受影響"}
for tk_, nm_, pe_, ev_, dt_, gaap_ in FN.PEERS:
    V.cell(r, 1, nm_); V.cell(r, 2, tk_); V.cell(r, 3, pe_).number_format = "0.00"; V.cell(r, 4, ev_).number_format = "0.00"
    V.cell(r, 5, datetime.date.fromisoformat(dt_)).number_format = "yyyy-mm-dd"
    V.cell(r, 6, f'=IF(ABS(NETWORKDAYS(E{r},px_date))-1>2,"逾期","")'); V.cell(r, 7, gaap_); V.cell(r, 8, "NTM 前瞻 EPS（資料商，S&P Global 共識）；" + LEASE[gaap_])
    r += 1
PE_B = r - 1
VR["pe_a"], VR["pe_b"] = PE_A, PE_B
IN[f"E{IN_ROW['pe_mult']}"] = f"=MEDIAN(Val!C{PE_A}:C{PE_B})"
V.cell(r, 1, "同業 P/E 中位數"); V.cell(r, 3, f"=MEDIAN(C{PE_A}:C{PE_B})").number_format = "0.00"; VR["pe_med"] = r; r += 1
V.cell(r, 1, "同業逾期家數（資料日與股價基準日相差 > 2 個交易日）"); V.cell(r, 3, f'=COUNTIF(F{PE_A}:F{PE_B},"逾期")'); VR["pe_old"] = r; r += 1
V.cell(r, 1, "NRI 現行前瞻 P/E（資料商，2026-11-09；Interested-party）"); V.cell(r, 3, 25.78).number_format = "0.00"; VR["pe_own"] = r; r += 1
V.cell(r, 1, "NRI 現行倍數 ÷ 同業中位數（隱含需相對重估 = 比值 − 1；10 年比值未找到）"); V.cell(r, 3, f"=C{VR['pe_own']}/C{VR['pe_med']}").number_format = "0.00"; VR["pe_ratio"] = r
V.cell(r, 4, f"=C{r}-1").number_format = "+0%;-0%"; VR["pe_reval"] = r; r += 1
# ===================================== Recon
RC["A1"] = "Recon — 對帳：本模型 vs 公司指引與市場共識（指引與共識為 Interested-party；找不到者寫「未找到」；口徑不同者不判定差異；¥B）。"
RCR = {}; r = 3
def rcrow(k, lab, vals):
    global r
    RCR[k] = r; RC.cell(r, 1, lab)
    for j, v in enumerate(vals, 2):
        if v is not None: RC.cell(r, j, v)
    r += 1
def rch(k, lab):
    global r
    RCR[k] = r; RC.cell(r, 1, lab).font = BOLD
    for j, h in enumerate(["本模型", "指引", "共識", "共識口徑", "說明"], 2): RC.cell(r, j, h).font = BOLD
    r += 1
rch("h1", "1. 合併層：FY2027E（情境 B；3 月期）")
Bm = PR["B"]
rcrow("rev27", "營收", [f"=PL!H{Bm['rev']}", G(FN.OTHER["guid_rev27"]), G(FN.OTHER["rev27_cons"]), "S&P Global 14 家（2026-09-11，財報前）", "公司 2026-10-29 上修後指引 870.0；共識為財報前，Interested-party；株予報 858.1（2026-07-24）。模型 2H27E 以 1H27 年增 × 延續比例推，不以指引反推"])
rcrow("oi27", "營業利益", [f"=PL!H{Bm['oi']}", G(FN.OTHER["guid_oi27"]), G(FN.OTHER["oi27_kb"]), "株予報（2026-07-24，財報前）", "指引 180.0（2026-10-29 自 150.0 上修）；差距來自 2H 季節性與 AI 效果（Thesis）假設，列示不判定"])
rcrow("ni27", "歸屬母公司淨利", [f"=PL!H{Bm['ni']}", G(FN.OTHER["guid_ni27"]), G(FN.OTHER["ni27_kb"]), "株予報（2026-07-24，財報前）", "指引 123.0；本模型財務收益、稅率、非控制權益為 Assumed／Derived"])
rcrow("eps27", "EPS（日圓）", [f"=PL!H{Bm['eps']}", FN.OTHER["guid_eps27"], "未找到", "—", "指引 218.50；本模型股數由庫藏股假設決定"])
rcrow("dps27", "每股股利（日圓）", [f"=PL!H{Bm['dps']}", FN.OTHER["guid_dps27"], "未找到", "—", "指引 88（中間 42＋期末 46）；本模型＝配息率 × EPS"])
rcrow("gap", "營業利益：本模型 − 指引（¥B；由公式組字）", [f"=B{RCR['oi27']}-C{RCR['oi27']}", None, None, None, "差距的來源見上列說明"])
rch("h2", "2. 共識與目標價")
rcrow("c1", "賣方平均目標價與家數（日圓）", [f"=Val!B{VR['tp']}", None, FN.OTHER["tp_cons_avg"], "14 家，高 6,700／低 5,190（2026-09-11，財報前；Interested-party：賣方）", "只作歷史對照，不作為模型輸入；財報後共識未找到"])
rcrow("c2", "FY2028 營收與利潤共識", [f"=PL!I{Bm['rev']}", None, "未找到", "—", "FY2028 共識未找到（資料商僅列至 FY2027 或財報前舊值）"])
rch("h3", "3. 同業與日期（G2-6、G2-10）")
rcrow("p1", "同業前瞻 P/E 中位數／現行 NRI ÷ 中位數", [f"=Val!C{VR['pe_med']}", None, f"=Val!C{VR['pe_ratio']}", "資料商 NTM", "同業資料日 2026-09-11 至 2026-11-16，與股價基準日不一致者標逾期（Val 頁）；8 家，不做迴歸；以係數 1.00 不調整"])
RC.cell(r + 1, 1, '="同業逾期 "&Val!C%d&" 家（日期相差超過 2 個交易日）；10 年期公債殖利率資料日 "&TEXT(In!E%d,"yyyy-mm-dd")&IF(NETWORKDAYS(In!E%d,px_date)-1>2,"（逾期）","")&"；現價日 "&TEXT(px_date,"yyyy-mm-dd")&"（距工作日 2026-11-16 逾 2 個交易日：逾期）。"' % (VR["pe_old"], IN_ROW["data_date"], IN_ROW["data_date"]))
RC.cell(r + 2, 1, "4. 口徑：資料商前瞻 P/E 為 NTM 共識 EPS；本模型 Comparable 以同業中位數 × 本公司 NTM EPS（目標時點起 12 個月，歸屬母公司淨利 ÷ 加權平均股數），分子（股價）分母（EPS）同為股權口徑，不涉及 FCF 或 EV 槓桿差異。")
RC.cell(r + 3, 1, "5. NRI 為 IFRS；同業 8 家含 IFRS 4 家（TIS、BIPROGY、Fujitsu、Infosys）、J-GAAP 2 家、US GAAP 2 家；P/E 對租賃處理不敏感，EV／EBITDA 對租賃敏感，本模型未用 EV／EBITDA。")
RC.cell(r + 4, 1, "6. 股數：流通股數＝發行 561.197409M − 庫藏 5.811620M（2026-09-30，決算短信封面）；NRI 未揭露 RSU／選擇權使稀釋 EPS 與基本 EPS 不同（1H27 皆 106.61），稀釋股數不另加；無 ASR、可轉債。")
RC.cell(r + 5, 1, "7. 淨債務口徑：有利子負債（公司口徑）＋ 另加租賃（基準 0；高端另加 2026-03 租賃負債 31.736）；營業租賃使用權資產折舊視為營業費用，不另計入淨債務；有利子負債是否已含租賃負債，由摘要不能確認（記為未找到）。")
RC.cell(r + 6, 1, "8. 未找到：見 data/research_nri.md 第 9 節（勞務費拆分、數值利用率、每任務 token 或代理任務量、FY2026 分部營業利益新口徑重編、財報後共識、4307 十年前瞻倍數）。")
RCR["dom_h"] = r + 8
RC.cell(r + 8, 1, "9. 主導變數反解：使加權目標價＝現價所需的值（整本重算割線法；原則第 4 條；不是評等門檻）").font = BOLD
for j, h in enumerate(["基準值", "反解值", "單位", "說明"], 2): RC.cell(r + 9, j, h).font = BOLD
KREQ = json.load(open(A.kreq, encoding="utf-8")) if A.kreq and os.path.exists(A.kreq) else None
for i_, row_ in enumerate(KREQ["rows"] if KREQ else []):
    RC.cell(r + 10 + i_, 1, row_["label"])
    for j, v in enumerate([row_["base"], row_["solved"], row_["unit"], row_["note"]], 2): RC.cell(r + 10 + i_, j, v)
RCR["kreq_a"] = r + 10; RCR["kreq_n"] = len(KREQ["rows"]) if KREQ else 0
# 指紋（輸入＋建置碼）：In 頁基準欄內容
FP = hashlib.sha256((BUILD_ID + json.dumps([(k, str(IN.cell(rr, 5).value)) for k, rr in sorted(IN_ROW.items())], ensure_ascii=False)).encode()).hexdigest()[:12]
BUILD_ID = f"nri-{VER}-{FP}"
if KREQ: RC.cell(r + 10 + RCR["kreq_n"], 1, "反解結果對應輸入指紋 " + str(KREQ["fingerprint"]) + "；" + ("與本檔一致" if KREQ["fingerprint"] == FP else "與本檔不一致（過期，須重跑）"))
# ===================================== Sens
SE["A1"] = "Sens — 敏感度：In 頁每個有「敏感度低／高」區間的輸入，一次動一個、整本 LibreOffice 重算；表中為加權目標價相對基準。情境專屬列與情境權重不列（三情境本身就是情境分析）。"
SE["A2"] = "token 價格層級與額外推論倍數的「高」案只動成本一邊，標「上界」；另列價格與成本聯動、模型組合（路由到便宜模型）案例（原則第 44 條）。"
SE.cell(3, 1, "build_id"); SE.cell(3, 4, BUILD_ID)
SE.cell(4, 1, "模型狀態"); SE.cell(4, 4, "CURRENT" if (SENS and SENS.get("fingerprint") == FP) else "STALE")
SE.cell(5, 1, "基準加權目標價（本檔）"); SE.cell(5, 4, f"=Val!B{VR['tp']}")
SE.cell(6, 1, "敏感度重算時的基準加權目標價"); SE.cell(6, 4, SENS["base_tp"] if SENS else None)
SE.cell(7, 1, "敏感度重算時的輸入指紋"); SE.cell(7, 4, SENS["fingerprint"] if SENS else None)
HDR_R = 9
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（日圓）", "投資建議"], 1): SE.cell(HDR_R, j, h).font = BOLD
SE_ROWS = []
if SENS:
    rr = HDR_R + 1
    for case in SENS["cases"]:
        SE.cell(rr, 1, case["label"]); SE.cell(rr, 2, case["setting"]); SE.cell(rr, 3, case["side"]); SE.cell(rr, 4, case["tp"]); SE.cell(rr, 5, f"=D{rr}/$D$5-1")
        SE.cell(rr, 6, case["key"]); SE.cell(rr, 7, case["dcf"]); SE.cell(rr, 8, f"=G{rr}-Val!$B${VR['tpdcf']}"); SE.cell(rr, 9, case["rec"]); SE_ROWS.append(rr); rr += 1
    SE.cell(rr + 1, 1, "註：tp_wA／B／C（情境權重）與 auto／ret／proj 各情境列不列：三情境並列已涵蓋。")
SE_END = HDR_R + len(SE_ROWS)
# ===================================== Sources
SO["A1"] = "來源與版本"; SO.cell(2, 1, "項目"); SO.cell(2, 2, "內容")
for i, (a, b2) in enumerate([("Tokenomics 快照", f"{TK_VER}（data/tokenomics_snapshot_v22.0.json）；名稱：IF_TaskLen、IF_TaskTokFresh／Cached／Dec、IF_Price{{Fresh,Cached,Out}}_{{Terra,Sol,Astra}}、IF_TaskSucc_{{Terra,Sol,Astra}}"),
                              ("研究包", "software_ai/nri/data/research_nri.md、facts_nri.csv、segments.json"),
                              ("公司申報", "NRI 決算短信（FY3/23–FY3/26、2026-10-29 中間期）、決算說明資料、IR 資料室 https://www.nri.com/jp/ir；WebFetch 摘要模型轉錄（非逐字人工核對）"),
                              ("同業倍數", "StockAnalysis 各公司統計頁（TIS、BIPROGY、Fujitsu、NS Solutions、OBIC、ACN、CTSH、INFY）"),
                              ("本版先驗", "未做")], 3):
    SO.cell(i, 1, a); SO.cell(i, 2, b2)
# ===================================== README
RD["A1"] = f'="野村總合研究所 NRI（4307）個股模型 {VER}｜資料截止 "&TEXT(data_date,"yyyy-mm-dd")&"｜成品日 {VDATE}｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"｜build_id {BUILD_ID}｜本版未做先驗"'
RD["A2"] = "命題、結論與重點觀察的數字由公式自本檔其他頁組字（不手寫）。"
VB = lambda k, col="B": f"Val!{col}{VR[k]}"
EB = SC["B"]
RD["A3"] = (f'="命題：AI 進入 IT 服務後，按人力計價的營收＝AI 新專案增量 − 價格侵蝕（客戶分享生產力）− 推論成本；利潤再加省下的人力成本。情境 B 在 FY2032 的 AI 淨效果（營收口徑）為 "&TEXT(Thesis!M{EB["net"]},"0.0")&" 十億日圓（AI 新專案 "&TEXT(Thesis!M{EB["prj"]},"0.0")&" − 價格侵蝕 "&TEXT(Thesis!M{EB["ero"]},"0.0")&" − 推論成本 "&TEXT(Thesis!M{EB["inf"]},"0.00")&"），營業利益效果 "&TEXT(Thesis!M{EB["aiop"]},"0.0")&" 十億日圓；損益兩平所需價格保留率 "&TEXT(Thesis!M{EB["be"]},"0%")&"（基準 "&TEXT(ret_B,"0%")&"）。"')
RD["A4"] = "命題的可證偽指標：Financial IT 營業利益率（1H27 17.3%）、每名交付人員的營收、AX 顧問與 AI 新專案營收占比；任何一項偏離基準區間即須重估。"
RD["A5"] = (f'="投資建議（在價格保留率落在情境 A 到 C 區間、推論成本取 Tokenomics 市場價且不降價的基準假設下）："&{VB("rec")}&"；估值評等 "&{VB("tier")}&"；加權目標價 "&TEXT({VB("tp")},"#,##0")&" 日圓，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"（現價 "&TEXT(px,"#,##0")&" 日圓，"&TEXT(px_date,"yyyy-mm-dd")&"）；收支關卡（FY2027–FY2031 自由現金流扣股利合計，情境 B）"&TEXT({VB("gate")},"0.0")&" 十億日圓"')
RD["A6"] = (f'="一、DCF 目標價 "&TEXT({VB("tpdcf")},"#,##0")&"、Comparable 目標價 "&TEXT({VB("tpcomp")},"#,##0")&" 日圓；落差 "&TEXT({VB("gap_dcf_cmp")},"#,##0")&" 日圓；DCF 隱含 NTM P/E "&TEXT({VB("impl_pe")},"0.0")&" 倍對同業中位數 "&TEXT(pe_mult,"0.0")&" 倍，倍數差貢獻 "&TEXT({VB("gap_pe")},"#,##0")&" 日圓；其餘來自股票薪酬（NRI 未揭露，影響小）、槓桿（淨現金）、時點（DCF 自評價日、Comparable 為目標時點）。"')
RD["A7"] = (f'="二、情境 B FY2027 營收 "&TEXT(Recon!B{RCR["rev27"]},"0.0")&" 十億日圓，公司指引 "&TEXT(Recon!C{RCR["rev27"]},"0.0")&"；營業利益 "&TEXT(Recon!B{RCR["oi27"]},"0.0")&"，指引 "&TEXT(Recon!C{RCR["oi27"]},"0.0")&"（2026-10-29 上修；Interested-party）；差 "&TEXT(Recon!B{RCR["gap"]},"0.0")&" 十億日圓。"')
RD["A8"] = (f'="三、每任務推論成本（Tokenomics 市場價，Analogy）約 "&TEXT(Thesis!C{TKR["task"]},"#,##0")&" 日圓，情境 B FY2032 推論成本僅 "&TEXT(Thesis!M{EB["inf"]},"0.00")&" 十億日圓；NRI 未揭露 token 量與代理任務量，每任務 token 取 Tokenomics 2 小時任務值（Analogy），額外倍數為 Assumed。"')
RD["A9"] = (f'="四、現價隱含 FY2032 營業利益率 "&TEXT(RevDCF!B{RVR["m1"]},"0.0%")&"（情境 B 為 "&TEXT(RevDCF!B{RVR["m0"]},"0.0%")&"）；永續價值占 EV 約 "&TEXT(Val!B{FC_["B"]["tvs"]},"0%")&"。"')
RD["A10"] = (f'="五、上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"（DCF "&TEXT({VB("upd")},"+0.0%;-0.0%")&"、Comparable "&TEXT({VB("upc")},"+0.0%;-0.0%")&"）；本公司現行前瞻 P/E 為同業中位數的 "&TEXT({VB("pe_ratio","C")},"0.00")&" 倍，Comparable 需相對重估 "&TEXT({VB("pe_reval","D")},"+0%;-0%")&"（10 年比值未找到）。"')
RD["A12"] = "v1.3 展示層"
RD["A13"] = (f'="野村總研："&{VB("rec")}&"，加權目標價 "&TEXT({VB("tp")},"#,##0")&" 日圓（上檔 "&TEXT({VB("up")},"+0.0%;-0.0%")&"）；本版未做先驗。"')
RD["A14"] = "本版要點"
RD["A15"] = "官方營收分部取 2026-04 新口徑（Consulting／Financial IT／Industrial IT／IT Platform Services／Others）；FY2024A、FY2025A 為舊口徑（澳洲在 Industrial IT），FY2026A 起新口徑重編。"
RD["A16"] = "命題以交付人員恆等式驅動三個人力計價分部；IT Platform Services 以成長路徑；子列（產業別產品線）占比依 1H27 揭露持平。"
RD["A17"] = "推論成本只讀 Tokenomics v22.0 快照；NRI 無 token 或代理任務量揭露，每任務 token 與價格層級為 Analogy，額外倍數為 Assumed，待 AI供應鏈登錄。"
RD["A18"] = "未做事項：先驗；三層基準協調、Wattonomics 對照不適用；分部營業利益新口徑 FY2026 重編值、財報後共識、勞務費拆分、數值利用率：未找到。"
RD["A20"] = "市場資料日期"
RD["A21"] = (f'="現價 "&TEXT(px_date,"yyyy-mm-dd")&" 收盤（距工作日 2026-11-16 逾 2 個交易日：逾期）；無風險利率資料日 2026-11-05（財務省入札）"&IF(NETWORKDAYS(DATE(2026,11,5),px_date)-1>2,"（逾期）","（與現價日相差不逾 2 個交易日）")&"；同業倍數日期 "&TEXT(MIN(Val!E{PE_A}:E{PE_B}),"yyyy-mm-dd")&" 至 "&TEXT(MAX(Val!E{PE_A}:E{PE_B}),"yyyy-mm-dd")&"，逾期 "&Val!C{VR["pe_old"]}&" 家；股數取決算短信封面（2026-09-30）。"')
RD["A22"] = (f'="主要失效條件（可觀察）：Financial IT 營業利益率跌破 "&TEXT(mg_fin-0.02,"0.0%")&"；訂單餘額年增轉負（1H27 +33.5%，Interested-party）；價格保留率實際低於損益兩平 "&TEXT(Thesis!M{EB["be"]},"0%")&"；AI 新專案占營收低於情境 C 的 "&TEXT(proj_C,"0.0%")&"。"')
RD["A23"] = "風險登錄"
for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"], 1): RD.cell(24, j, h)
for i, rk in enumerate([("客戶要求分享 AI 生產力（價格保留率下降）", "是", "ret 三情境", "Financial IT 營收／營業利益年增差"), ("自動化取代工時", "是", "auto 三情境", "交付人員數、營收／人"), ("推論成本侵蝕", "是", "Tokenomics 市場價與倍數", "推論成本占營收"), ("澳洲減損後續與海外併購", "否", "FY2026 已列 76.998 減損，預測期不再認列", "Financial IT 海外營收、減損公告"), ("庫藏股與配息", "部分", "只影響股數與現金", "庫藏股公告")], 25):
    for j, v in enumerate(rk, 1): RD.cell(i, j, v)
RD.cell(31, 1, "工作表：README、In、Hist、Q、Thesis、PL、CF、Val、RevDCF、Recon、Sens、Sources")
for j, h in enumerate(["頁", "輸入", "推導量", "因果方向"], 1): RD.cell(33, j, h)
for i, rk in enumerate([("Thesis", "util、hrs_y、rate0、c_head、sav_r、auto／ret／proj_A／B／C、inf_mult、tk_tier", "隱含人數、AI 淨效果、損益兩平價格保留率、推論成本", "人數 × 利用率 × 工時 × 單價 → 營收；侵蝕、新專案、推論成本 → 淨效果"),
                        ("PL", "gv_28／32、dg_*、gi_*、mg_*、seas、tax、fin_net", "分部營收、營業利益、淨利、EPS", "分部營收 × 利益率 ＋ AI 效果 → 營業利益 → 淨利"),
                        ("Val", "rf、beta_raw、erp、g_tv、w_dcf、pe_mult", "WACC、DCF、Comparable、加權目標價、評等", "FCFF 折現＋同業 P/E × NTM EPS → 加權"),
                        ("Sens", "In 頁敏感度低／高", "加權目標價、DCF 目標價、投資建議", "一次改一個輸入、整本重算")], 34):
    for j, v in enumerate(rk, 1): RD.cell(i, j, v)
RD["A39"] = "版本沿革：v0.1（2026-11-16）首版；未做先驗；本版與其他版本無對照。"
# ===================================== 套用 In 版面、樣式與輸出
import apply_in_layout
apply_in_layout.apply(wb, VER, "2026-11-16")
for name, rng in (("PL_hdr", f"PL!$C${PR['B']['hdr']}:$M${PR['B']['hdr']}"),): wb.defined_names[name] = DefinedName(name, attr_text=rng)
wb.save(A.out)
rowmap = {"IR": IN_ROW, "TR": TR, "SC": SC, "PR": {k: v for k, v in PR.items() if k != "IS"}, "IS": PR["IS"], "VR": VR, "FC": FC_, "RVR": RVR, "RCR": RCR, "CR": CR, "TKR": TKR, "SE_ROWS": SE_ROWS, "SE_END": SE_END, "HDR_R": HDR_R, "BUILD_ID": BUILD_ID, "FP": FP, "GATE": GATE, "HR": HR, "PE": [PE_A, PE_B]}
if os.environ.get("ROWMAP"): json.dump(rowmap, open(os.environ["ROWMAP"], "w"), ensure_ascii=False, indent=1, default=str)
print("built", A.out, BUILD_ID)
