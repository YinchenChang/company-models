# Datadog（DDOG）個股模型 v0.1（命題＋官方營收分部 → 損益表；原型 U 按用量計價；docs/plan/個股模型原則_命題加營收分部.md、個股DCF標準規格.md、個股HTML版面規格.md v1.3）
# 工作單：software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md 的 G 與 G2 節。本版未做先驗。
# 用法：python3 software_ai/datadog/build_datadog.py raw.xlsx [--sens sens.json] [--kreq kreq.json] [--set 鍵=值 ...]；再以 LibreOffice headless 重算（make_datadog.sh 一次跑完）。
# 金額：百萬美元（US$M）；股數：百萬股；每股：美元。會計年度＝日曆年（12 月 31 日止）。
# 欄位（模型頁共用）：C FY2023A、D FY2024A、E FY2025A、F 9M26A、G Q4 26E、H FY2026E（＝F＋G）、I–N FY2027–FY2032。評價日 2026-09-30。
import sys, os, json, argparse, re, hashlib, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument("out")
ap.add_argument("--sens"); ap.add_argument("--kreq")
ap.add_argument("--set", action="append", default=[])
A = ap.parse_args()
OVR = {}
for s_ in A.set:
    k_, v_ = s_.split("=", 1); OVR[k_] = float(v_)
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); SUB = PatternFill("solid", fgColor="E2EFDA")
wb = openpyxl.Workbook()
VER = "v0.1"; VDATE = "2026-11-16"


def _build_id():
    h = hashlib.sha256()
    for fn in sorted(["build_datadog.py", "build_datadog_p2.py", "build_datadog_p3.py", "build_datadog_p4.py"] + [os.path.join("data", f) for f in os.listdir(os.path.join(HERE, "data")) if f.endswith(".json") and not f.startswith(("sens_", "kreq_"))]):
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"datadog-{VER}-{h.hexdigest()[:10]}"


BUILD_ID = _build_id()
FD = json.load(open(os.path.join(HERE, "data", "datadog_facts.json"), encoding="utf-8"))
F = {it["id"]: it for it in FD["items"]}
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}
KREQ = json.load(open(A.kreq, encoding="utf-8")) if A.kreq and os.path.exists(A.kreq) else {}
def fv(i): return F[i]["value"]
def fs(i): it = F[i]; return f"{it['source']}（擷取 {FD['asOf']}；事實 {i}）"
def tkv(n): return TK["items"][n]["values"]
TKS = f"Tokenomics v22.0 快照（data/tokenomics_snapshot_v22.0.json；commit {TK['source']['commit'][:7]}）"


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
def T2(x): return f'TEXT({x},"#,##0.00")'
def T0(x): return f'TEXT({x},"#,##0")'
def T1(x): return f'TEXT({x},"#,##0.0")'
def P1(x): return f'TEXT({x},"0.0%")'
def SP1(x): return f'TEXT({x},"+0.0%;-0.0%;0.0%")'
def D10(x): return f'TEXT({x},"yyyy-mm-dd")'


YK = ["23", "24", "25", "1H", "2H", "26", 27, 28, 29, 30, 31, 32]
COL = dict(zip(YK, "CDEFGHIJKLMN"))
YH = {"23": "FY2023A", "24": "FY2024A", "25": "FY2025A", "1H": "9M26A", "2H": "Q4 26E", "26": "FY2026E", 27: "FY2027E", 28: "FY2028E", 29: "FY2029E", 30: "FY2030E", 31: "FY2031E", 32: "FY2032E"}
ANN = [27, 28, 29, 30, 31, 32]
ANNC = [COL[y] for y in ANN]
NOTE_C = 15
MODEL_HDR = ["項目", "單位"] + [YH[k] for k in YK] + ["說明"]
def prevc(c): return chr(ord(c) - 1)


def widths(ws, a=58, b=10, mid=11, note=90):
    ws.column_dimensions["A"].width = a; ws.column_dimensions["B"].width = b
    for i in range(12): ws.column_dimensions[chr(ord("C") + i)].width = mid
    ws.column_dimensions["O"].width = note


def put(ws, r, label, unit, vals, fmt="#,##0.0", note=None, bold=False, fill=None):
    ws.cell(r, 1, label); ws.cell(r, 2, unit)
    if bold: ws.cell(r, 1).font = BOLD
    for c, v in vals.items():
        cell = ws[f"{c}{r}"]; cell.value = v; cell.number_format = fmt
    if note: ws.cell(r, NOTE_C, note)
    if fill:
        for cc in range(1, NOTE_C + 1): ws.cell(r, cc).fill = fill


RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); HS = wb.create_sheet("Hist"); QQ = wb.create_sheet("Q"); TH = wb.create_sheet("Thesis"); SG = wb.create_sheet("Seg")
PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF"); DC = wb.create_sheet("DCF"); V = wb.create_sheet("Val"); RV = wb.create_sheet("RevDCF")
RC = wb.create_sheet("Recon"); SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")

# ====================================================================================================== Hist（歷史財務；原則第 10 節第 14 條）
HS["A1"] = "Hist — 歷史財務（US$M；會計年度＝日曆年；公式只引用本頁儲存格，定義名稱 h_<項目>_<期間>）。來源：財報新聞稿與 10-K／10-Q（見 K 欄）；FY2021–22 部分項目未取得（留空＝未找到）"
HS["A2"] = "核對列：FY2025 四季合計＝全年（見 Q 頁）；9M26＝三季合計（見 Q 頁）。「營運資金與其他變動」＝營運現金流 − 淨利 − 折舊攤銷 − 股票薪酬（含遞延合約成本、遞延收入等；原則第 23 條，逐年與營運現金流勾稽）。"
hdr(HS, 3, ["項目", "單位", "FY2021A", "FY2022A", "FY2023A", "FY2024A", "FY2025A", "9M25A", "9M26A", "標記", "來源", "來源誘因"])
HP = {"21": "C", "22": "D", "23": "E", "24": "F", "25": "G", "9m25": "H", "9m26": "I"}
HROWS = [("rev", "營收"), ("cogs", "營業成本（GAAP）"), ("gp", "毛利（GAAP）"), ("rd", "研發費用（GAAP）"), ("sm", "銷售與行銷費用（GAAP）"), ("ga", "一般及行政費用（GAAP）"), ("oi", "營業利益（GAAP）"),
         ("onet", "其他收入淨額"), ("pti", "稅前淨利"), ("tax", "所得稅費用"), ("ni", "淨利（GAAP）"), ("dsh", "稀釋加權股數（百萬股）"), ("eps", "稀釋 EPS（GAAP，US$）"), ("sbc", "股票薪酬（總額）"),
         ("ptx", "員工股票交易薪資稅"), ("amort", "取得無形資產攤銷"), ("mna", "併購交易費用"), ("ngp", "毛利（non-GAAP）"), ("ngrd", "研發費用（non-GAAP）"), ("ngsm", "銷售與行銷費用（non-GAAP）"),
         ("ngga", "一般及行政費用（non-GAAP）"), ("ngoi", "營業利益（non-GAAP）"), ("ngeps", "稀釋 EPS（non-GAAP，US$）"), ("da", "折舊與攤銷"), ("ocf", "營運現金流"), ("capex", "購置不動產與設備"),
         ("capsw", "資本化軟體開發成本"), ("fcf", "自由現金流（公司定義）"), ("wcx", "營運資金與其他變動（營運現金流 − 淨利 − 折舊攤銷 − 股票薪酬）")]
H_ROW = {k_: 4 + i_ for i_, (k_, _) in enumerate(HROWS)}
r = 4
for k, lab in HROWS:
    HS.cell(r, 1, lab); HS.cell(r, 2, "百萬股" if k == "dsh" else ("US$" if "eps" in k else "US$M")); H_ROW[k] = r
    derived = k in ("gp", "wcx")
    for p, c in HP.items():
        cell = HS[f"{c}{r}"]
        if k == "gp" and f"is.cogs.{'fy20' + p if p.isdigit() else p}" in F or (k == "gp" and p in ("9m25", "9m26")):
            cell.value = f"={c}{H_ROW['rev']}-{c}{H_ROW['cogs']}"
        elif k == "onet" and p == "23":
            cell.value = f"={c}{H_ROW['pti']}-{c}{H_ROW['oi']}"
        elif k == "wcx":
            if all(x in F for x in [f"is.ocf.{'fy20' + p if p.isdigit() else p}", f"is.ni.{'fy20' + p if p.isdigit() else p}", f"is.da.{'fy20' + p if p.isdigit() else p}"]):
                cell.value = f"={c}{H_ROW['ocf']}-{c}{H_ROW['ni']}-{c}{H_ROW['da']}-{c}{H_ROW['sbc']}"
        else:
            fid = f"is.{k}.{'fy20' + p if p.isdigit() else p}"
            if fid in F: cell.value = F[fid]["value"]; cell.font = BLUE
        if cell.value is not None:
            cell.number_format = "#,##0.000"; dn(f"h_{k}_{p}", f"Hist!${c}${r}")
    HS.cell(r, 10, "Derived" if derived or k == "onet" else "Verified")
    HS.cell(r, 11, "財報新聞稿與 10-K／10-Q（事實總帳 data/facts_datadog.csv，id＝is.%s.*）；稀釋 EPS 與 non-GAAP 為公司揭露" % k)
    if k in ("ngp", "ngrd", "ngsm", "ngga", "ngoi", "ngeps"):
        HS.cell(r, 10, "Interested-party"); HS.cell(r, 12, "公司自訂 non-GAAP 口徑：排除股票薪酬等，呈現較高獲利的誘因")
    r += 1
HS_LAST = r - 1
r += 1
HS.cell(r, 1, "資產負債表與其他（2026-09-30）").font = BOLD; r += 1
hdr(HS, r, ["項目", "單位", "數值", "", "", "", "", "", "", "標記", "來源", "來源誘因"]); r += 1
for fid, lab, nm in (("bs.cash.20260930", "現金及約當現金", "h_cash"), ("bs.mkt.20260930", "有價證券", "h_mkt"), ("debt.face.20260930", "可轉換票據本金（0.00%，2029-12-01 到期）", "h_debt"),
                     ("rpo.20260930", "剩餘履約義務（RPO）", "h_rpo26"), ("rpo.20251231", "剩餘履約義務（RPO）2025-12-31", "h_rpo25"), ("rpo.20241231", "剩餘履約義務（RPO）2024-12-31", "h_rpo24")):
    HS.cell(r, 1, lab); HS.cell(r, 2, "US$M"); c = HS.cell(r, 3, fv(fid)); c.font = BLUE; c.number_format = "#,##0.0"; dn(nm, f"Hist!$C${r}")
    HS.cell(r, 10, F[fid]["tag"]); HS.cell(r, 11, fs(fid)); r += 1
widths(HS, a=60, b=9, mid=12, note=40)
HS.column_dimensions["K"].width = 90

# ====================================================================================================== Q（季；單一主要標記）
QQ["A1"] = "Q — 季度實績（US$M；單一主要標記；公司自揭 KPI 標 Interested-party 並寫來源誘因）。核對列：四季合計＝FY2025；三季合計＝9M26"
hdr(QQ, 3, ["項目", "單位", "4Q24", "1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26", "3Q26", "標記", "來源", "來源誘因"])
QK = ["4q24", "1q25", "2q25", "3q25", "4q25", "1q26", "2q26", "3q26"]; QC = dict(zip(QK, "CDEFGHIJ"))
Q_ROW = {}
r = 4
for k, lab in (("rev", "營收"), ("cogs", "營業成本（GAAP）"), ("oi", "營業利益（GAAP）"), ("ngoi", "營業利益（non-GAAP）"), ("sbc", "股票薪酬"), ("ni", "淨利（GAAP）")):
    QQ.cell(r, 1, lab); QQ.cell(r, 2, "US$M"); Q_ROW[k] = r
    for q in QK:
        c = QQ[f"{QC[q]}{r}"]; c.value = fv(f"q.{k}.{q}"); c.font = BLUE; c.number_format = "#,##0.000"; dn(f"q_{k}_{q}", f"Q!${QC[q]}${r}")
    QQ.cell(r, 11, "Interested-party" if k == "ngoi" else "Verified"); QQ.cell(r, 12, "各季財報新聞稿（事實 q.%s.*；見 data/facts_datadog.csv）" % k)
    if k == "ngoi": QQ.cell(r, 13, "公司自訂 non-GAAP 口徑（排除股票薪酬等）")
    r += 1
QQ.cell(r, 1, "稀釋加權股數 3Q26（百萬股）"); QQ.cell(r, 2, "百萬股"); c = QQ.cell(r, 10, fv("is.dsh.3q26")); c.font = BLUE; dn("q_dsh_3q26", f"Q!$J${r}"); QQ.cell(r, 11, "Verified"); QQ.cell(r, 12, fs("is.dsh.3q26")); r += 1
QQ.cell(r, 1, "營收年增率"); QQ.cell(r, 2, "%"); Q_ROW["yoy"] = r
for q, pq in (("1q26", "1q25"), ("2q26", "2q25"), ("3q26", "3q25"), ("4q25", "4q24")):
    QQ[f"{QC[q]}{r}"] = f"={QC[q]}{Q_ROW['rev']}/{QC[pq]}{Q_ROW['rev']}-1"; QQ[f"{QC[q]}{r}"].number_format = "0.0%"
QQ.cell(r, 11, "Derived"); QQ.cell(r, 12, "由本頁營收相除"); r += 1
QQ.cell(r, 1, "營收季增率"); QQ.cell(r, 2, "%")
for i, q in enumerate(QK[1:], 1):
    QQ[f"{QC[q]}{r}"] = f"={QC[q]}{Q_ROW['rev']}/{QC[QK[i - 1]]}{Q_ROW['rev']}-1"; QQ[f"{QC[q]}{r}"].number_format = "0.0%"
QQ.cell(r, 11, "Derived"); QQ.cell(r, 12, "由本頁營收相除"); r += 1
QQ.cell(r, 1, "non-GAAP 營業利益率"); QQ.cell(r, 2, "%")
for q in QK:
    QQ[f"{QC[q]}{r}"] = f"={QC[q]}{Q_ROW['ngoi']}/{QC[q]}{Q_ROW['rev']}"; QQ[f"{QC[q]}{r}"].number_format = "0.0%"
QQ.cell(r, 11, "Derived"); QQ.cell(r, 12, "由本頁相除"); r += 1
for fid, lab in (("q.cust100k.3q25", "ARR ≥ $100K 客戶數（約）"),):
    pass
QQ.cell(r, 1, "ARR ≥ $100K 客戶數（約）"); QQ.cell(r, 2, "戶")
for q, fid in (("4q24", "q.cust100k.4q24"), ("3q25", "q.cust100k.3q25"), ("4q25", "q.cust100k.4q25"), ("1q26", "q.cust100k.1q26"), ("2q26", "q.cust100k.2q26"), ("3q26", "q.cust100k.3q26")):
    c = QQ[f"{QC[q]}{r}"]; c.value = fv(fid); c.font = BLUE
QQ.cell(r, 11, "Interested-party"); QQ.cell(r, 12, "各季財報新聞稿（事實 q.cust100k.*）"); QQ.cell(r, 13, F["q.cust100k.3q26"]["incentive"]); r += 2
QQ.cell(r, 1, "檢查：FY2025 四季營收合計 − 年度營收（應為 0）"); QQ.cell(r, 3, f"=SUM(D{Q_ROW['rev']}:G{Q_ROW['rev']})-h_rev_25").number_format = "0.000"; Q_CHK1 = r; r += 1
QQ.cell(r, 1, "檢查：9M26 三季營收合計 − 9M26 營收（應為 0）"); QQ.cell(r, 3, f"=SUM(H{Q_ROW['rev']}:J{Q_ROW['rev']})-h_rev_9m26").number_format = "0.000"; Q_CHK2 = r; r += 1
QQ.cell(r, 1, "檢查：9M26 三季 non-GAAP 營業利益合計 − 9M26 non-GAAP 營業利益（應為 0）"); QQ.cell(r, 3, f"=SUM(H{Q_ROW['ngoi']}:J{Q_ROW['ngoi']})-h_ngoi_9m26").number_format = "0.000"; Q_CHK3 = r; r += 1
widths(QQ, a=62, b=9, mid=11, note=30); QQ.column_dimensions["L"].width = 60

# ====================================================================================================== In
IN["A1"] = "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$M；股數百萬股"
IN["A2"] = ('="事實總帳 data/facts_datadog.csv（345 列，facts_datadog.py）；Tokenomics v22.0 快照 data/tokenomics_snapshot_v22.0.json；評價日 "&TEXT(val_date,"yyyy-mm-dd")&"；現價日 "&TEXT(px_date,"yyyy-mm-dd")&"；目標時點 "&TEXT(tp_date,"yyyy-mm-dd")')
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"])
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", judg=False, inc=None): ROWS.append((key, item, unit, lo, base, hi, tag, src, note, judg, inc))
INC_CO = "公司管理層（KPI、法說口頭、指引）：呈現成長與維持市場信心的誘因，口徑由公司定義、未經查核"
INC_SELL = "賣方分析師：評等與投行／交易業務誘因，目標價傾向跟隨股價"
INC_DATA = "資料商計算值（以賣方預估為分母）：無直接利益，但分母口徑與日期未經查核"


def ttm(k): return fv(f"is.{k}.4q25") + fv(f"is.{k}.9m26")


sec("A. 日期、現價、股數與資產負債（評價日 2026-09-30）")
inp("data_date", "資料截止日（本版最晚一筆輸入：Tokenomics v22.0 快照擷取日）", "日期", None, datetime.date(2026, 11, 16), None, "Derived", "data/tokenomics_snapshot_v22.0.json source.extractedAt", "原則第 10 節第 41 條")
inp("px_date", "股價基準日（現價日）", "日期", None, datetime.date(2026, 11, 13), None, "Verified", fs("mkt.px.20261113"), "原則第 26 條：目標時點＝本值＋12 個月")
inp("val_date", "評價日（最近一季資產負債表日）", "日期", None, datetime.date(2026, 9, 30), None, "Verified", fs("bs.cash.20260930"))
inp("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條")
inp("px", "現價（2026-11-13 收盤）", "US$", None, fv("mkt.px.20261113"), None, "Verified", fs("mkt.px.20261113") + "；" + fs("mkt.px.20261112"), "兩來源前收價相符（11-12 收盤 184.46）；財報後（11-04 單日 −13%～−14%）")
inp("sh_out", "流通股數（10-Q 封面 2026-10-28，Class A＋B）", "百萬股", None, fv("sh.out.20261028"), None, "Verified", fs("sh.out.20261028"), "法定申報封面為主來源（G2 第 4 條）；stockanalysis 361.04 相符")
inp("rsu_unv", "未歸屬 RSU／PSU（2026-09-30）", "百萬股", None, fv("rsu.20260930"), None, "Verified", fs("rsu.20260930"), "DCF 規格第 3 節：全額計入，扣預期沒收")
inp("rsu_fo", "RSU 預期沒收率（10-Q 未揭露；未找到）", "%", None, 0.0, 0.10, "Assumed", "未找到：10-Q 未載沒收率；Sens 高＝10%", "G2 第 4 條；基準 0＝不扣（保守地多計股數）", judg=True)
inp("opt_n", "流通選擇權（2026-09-30）", "百萬股", None, fv("opt.n.20260930"), None, "Verified", fs("opt.n.20260930"))
inp("opt_k", "選擇權加權平均履約價", "US$", None, fv("opt.k.20260930"), None, "Verified", fs("opt.k.20260930"), "庫藏股法：價格用現價")
inp("dil_inc", "評價股數計入 RSU 與選擇權（1＝是；0＝只用流通股）", "x", 0, 1, None, "Assumed", "DCF 規格第 3 節（基準 1）", "Sens 低＝0（原口徑）", judg=True)
inp("conv_if", "2029 票據視為股份（1＝以 if-converted 計入股數並移除債務；0＝視為債務）", "x", None, 0, 1, "Assumed", "票據本金 1,000、轉換價約 217.60（" + fs("conv.px") + "）；現價與基準每股價值皆低於轉換價 → 視為債務", "Sens 高＝1", judg=True)
inp("conv_sh", "票據轉換股數（百萬股）＝1,000 ÷ 217.60", "百萬股", None, "=h_debt/217.6", None, "Derived", "4.5955 股／千美元（" + fs("conv.px") + "）")
inp("cash_0", "現金＋有價證券（2026-09-30）", "US$M", None, "=h_cash+h_mkt", None, "Derived", fs("bs.cash.20260930") + "；" + fs("bs.mkt.20260930"), "881.230＋4,352.357")
inp("debt_0", "債務本金（2029 年到期 0.00% 可轉換票據）", "US$M", None, "=h_debt", None, "Verified", fs("debt.face.20260930"), "淨債務只計債務本金＋融資租賃，營業租賃不計（G2 第 5 條）")
inp("fin_lease", "融資租賃負債（10-Q 未見融資租賃；未找到）", "US$M", None, 0.0, None, "Assumed", "未找到：資產負債表只列營業租賃負債（流動 46.177、非流動 274.462；不計入，因營業租賃費用已在營業利益）", "G2 第 5 條")
sec("B. 營收：命題模組（原型 U 按用量計價：營收＝用量單位 × 單價；v0.2 改讀 unit_economics，本版先在此自建）")
inp("q4_g", "Q4 26E 營收季增率（相對 3Q26）", "%", -0.01, 0.01, 0.03, "Assumed", "指引中點隱含 −0.2%（" + fs("guid.g_q4lo") + "）；近三季實績皆超出指引上緣 3.8%–4.7%（Recon）", "指引不入基準（X3）；基準＝略高於指引上緣的 +1.0%", judg=True)
inp("ai_26", "AI 產品 credits 營收 FY2026E（Bits AI 等；公司未揭露）", "US$M", 20, 45, 90, "Assumed", "未找到：公司不揭露 AI 產品營收；約三分之一客戶使用 AI 產品（" + fs("ai.pen.3q26") + "）", "基準＝FY2026E 營收約 1%", judg=True)
inp("sa_26", "AI 原生客戶占 FY2026E 平台營收比（扣除 AI credits）", "%", 0.12, 0.15, 0.18, "Assumed", "公司最後揭露 2025Q3 12%（" + fs("ai.sh.3q25") + "）；2026Q3 由非 AI 成長 28% 反推約 16.7%（Recon 第 3 節，Derived）", "全年平均介於兩者", judg=True)
inp("vn_27_A", "非 AI 客戶用量成長 FY2027（情境 A）", "%", None, 0.27, None, "Assumed", "2026Q3 非 AI 客戶營收成長 high 20s、連續五季加速（" + fs("ai.nonai.3q26") + "）", "樂觀：維持高 20s 動能")
inp("vn_27_B", "非 AI 客戶用量成長 FY2027（情境 B）", "%", None, 0.24, None, "Assumed", "同上；共識 FY2027 總成長 23%（" + fs("cons.rev27g") + "）", "基準：介於近季 28% 與保守情境之間，反映 Q4 指引隱含的減速")
inp("vn_27_C", "非 AI 客戶用量成長 FY2027（情境 C）", "%", None, 0.18, None, "Assumed", "同上", "保守：雲端最佳化、預算趨緊（The Stack 2026-11-04）")
inp("vn_32", "非 AI 客戶用量成長 FY2032（線性收斂目標）", "%", 0.06, 0.09, 0.12, "Assumed", "判斷：大型 SaaS 十年期收斂", "FY2027 至 FY2032 線性", judg=True)
inp("pn_g", "非 AI 客戶單價年變動（量大折扣、承諾用量折價）", "%", -0.04, -0.02, 0.0, "Assumed", "公司稱預期類似折扣結構（" + fs("ai.exlargest.fy26") + "）；金額未揭露", "價格保留率＝1＋本值", judg=True)
inp("va_27_A", "AI 原生客戶用量成長 FY2027（情境 A）", "%", None, 0.45, None, "Assumed", "代理工作負載：MCP 工具呼叫年增 50 倍、代理可觀測 span 年增 9 倍（" + fs("ai.cust.20260930") + "）", "樂觀：代理工作負載持續放大")
inp("va_27_B", "AI 原生客戶用量成長 FY2027（情境 B）", "%", None, 0.30, None, "Assumed", "AI 客戶成長「明顯快於其他業務」；2026Q3 貢獻成長高個位數百分點（10-Q）", "基準")
inp("va_27_C", "AI 原生客戶用量成長 FY2027（情境 C）", "%", None, 0.05, None, "Assumed", "最大客戶 2026Q3 用量下降，Q4 指引以合約最低承諾計（" + fs("ai.largest.minq4") + "）", "保守：最大客戶停留於最低承諾，群體其餘成長被抵銷")
inp("va_32", "AI 原生客戶用量成長 FY2032（線性收斂目標）", "%", 0.06, 0.10, 0.14, "Assumed", "判斷", "FY2027 至 FY2032 線性", judg=True)
inp("pa_A", "AI 原生客戶單價 FY2027 變動（情境 A）", "%", None, -0.05, None, "Assumed", "多家大型 AI 原生客戶以較低單價提前續約（" + fs("ai.cust.20260930") + "；Q3 法說：early renewals at better pricing）", "樂觀：讓利較小；讓利幅度自 FY2027 起三年線性消退")
inp("pa_B", "AI 原生客戶單價 FY2027 變動（情境 B）", "%", None, -0.10, None, "Assumed", "同上；幅度公司未揭露", "基準")
inp("pa_C", "AI 原生客戶單價 FY2027 變動（情境 C）", "%", None, -0.15, None, "Assumed", "同上", "保守")
inp("lc_cut", "事件型尾端：最大客戶與同群體 FY2027 用量再下降比例（基準 0）", "%", None, 0.0, 0.5, "Assumed", "最大客戶 2026Q3 已降用量、占比未揭露（" + fs("ai.largest.minq4") + "）", "原則第 19 條：單一大客戶事件型尾端 Sens，不進基準", judg=True)
inp("ai_27_A", "AI 產品 credits 營收 FY2027（情境 A）", "US$M", None, 180, None, "Assumed", "未找到：公司不揭露", "樂觀：Bits 為公司史上採用最快產品（法說）")
inp("ai_27_B", "AI 產品 credits 營收 FY2027（情境 B）", "US$M", None, 110, None, "Assumed", "同上", "基準")
inp("ai_27_C", "AI 產品 credits 營收 FY2027（情境 C）", "US$M", None, 60, None, "Assumed", "同上", "保守")
inp("gai28", "AI 產品 credits 營收 FY2028 年增", "%", 0.4, 0.6, 0.8, "Assumed", "判斷", "之後每年下降 gai_step", judg=True)
inp("gai_step", "AI 產品 credits 年增每年下降（百分點）", "%", None, 0.10, None, "Assumed", "判斷")
inp("gai_floor", "AI 產品 credits 年增下限", "%", 0.1, 0.2, 0.3, "Assumed", "判斷", "", judg=True)
sec("C. 推論成本（只讀 Tokenomics；每任務 token 先用公司揭露校準，G2 第 1 條）")
inp("cr_task", "每次 Bits Investigation 耗用 credits（官網估算）", "credits", 3, fv("px.credit.inv"), 6, "Interested-party", fs("px.credit.inv"), "官網稱僅供參考", inc="公司定價頁估算：呈現產品親民的誘因；實際依任務")
inp("cr_px", "每 credit 價格（年約牌價）", "US$", 0.80, fv("px.credit.annual"), fv("px.credit.od"), "Verified", fs("px.credit.annual") + "；隨需 1.30（" + fs("px.credit.od") + "）", "低＝0.80 為 Assumed 大客戶折扣")
inp("cost_sh", "每任務推論成本 ÷ 每任務 credit 收入（公司：credits 定價貼近推論成本）", "%", 0.6, 1.0, 1.2, "Interested-party", "Q3 法說：「we stay fairly close to inference costs there」；「pass on the savings as the models become more efficient to our customers」（" + F["q.rev.3q26"]["source"] + "；" + fs("ai.pen.3q26") + "）", "以揭露校準每任務 token；比例取點為 Assumed", inc=INC_CO, judg=True)
inp("tk_px0", "校準用 token 價格（IF_PriceRef_Sol，中高層參考請求混合）", "US$/M", None, tkv("IF_PriceRef_Sol"), None, "Analogy", TKS + "；Interface!C108；IFC_Conf＝B", "選哪一層價格屬 Analogy（G2 第 1 條）：Datadog 實際模型組合（含自家 Scout 與外購）未揭露")
inp("tk_tokens", "每任務 token（百萬）＝成本占比 × 每任務 credit 收入 ÷ 校準價格（由揭露推得，非聊天基數 × IF_TaskTokMult）", "百萬 token", None, "=cost_sh*cr_task*cr_px/tk_px0", None, "Derived", "公司揭露（credits 價格、耗用、成本貼近定價）× Tokenomics 價格；對照：IF_TaskTokMult 8 小時任務 ≈1.49 百萬 token（Recon 第 6 節）", "反映：公司揭露＞聊天基數類推")
inp("tk_px", "實際採用的 token 價格（Terra／Sol／Astra 層；Sens 低＝Terra、高＝Astra）", "US$/M", tkv("IF_PriceRef_Terra"), tkv("IF_PriceRef_Sol"), tkv("IF_PriceRef_Astra"), "Analogy", TKS + "；IF_PriceRef_Terra／Sol／Astra", "只動成本一邊；高案標「上界」（G2 第 2 條）", judg=True)
inp("tk_cheap", "路由到便宜模型時的 token 價格（IF_PriceRef_Terra）", "US$/M", None, tkv("IF_PriceRef_Terra"), None, "Analogy", TKS)
inp("route_sh", "路由到便宜模型的任務占比（模型組合）", "%", None, 0.0, 0.5, "Assumed", "公司自家 Scout 推理模型已用於 Bits Investigations（Q3 法說），占比未揭露", "G2 第 2 條：模型組合案例", judg=True)
inp("tp_dec", "token 價格年變動（市場價；待 AI供應鏈 登錄價格路徑）", "%", -0.5, -0.25, 0.0, "Assumed", "待 AI供應鏈 登錄：Tokenomics 快照只有 2026 單點價格", "不自行重估 token 價格；僅作敏感度區間", judg=True)
inp("pt_k", "token 價格變動轉嫁到 credit 價格的比例（聯動）", "%", None, 0.5, 1.0, "Assumed", "法說：把模型效率的節省轉給客戶", "高＝完全轉嫁（價格與成本聯動）", judg=True)
sec("D. 成本與損益（比率參數以最近四季實績持平；收斂只放 Sens，G2 第 3 條）")
rv_t = ttm("rev") if False else fv("q.rev.4q25") + fv("is.rev.9m26")
inp("gm_core", "non-GAAP 毛利率（不含 AI 推論成本；近四季）", "%", 0.79, round(ttm("ngp") / rv_t, 6), 0.815, "Derived", "近四季（4Q25＋9M26）non-GAAP 毛利 ÷ 營收（" + fs("is.ngp.9m26") + "；" + fs("is.ngp.4q25") + "）", "全期持平；單位成本隨用量（價格讓利會壓低毛利）", judg=True)
inp("uc_g", "單位成本年變動（工程效率；基準持平）", "%", 0.0, 0.0, -0.02, "Assumed", "基準無證據→持平；高＝每年 −2%", "G2 第 3 條", judg=True)
inp("rd_ng", "non-GAAP 研發費用率（近四季）", "%", None, round(ttm("ngrd") / rv_t, 6), None, "Derived", "近四季 non-GAAP 研發 ÷ 營收", "全期持平")
inp("sm_ng", "non-GAAP 銷售與行銷費用率（近四季）", "%", None, round(ttm("ngsm") / rv_t, 6), None, "Derived", "近四季", "全期持平")
inp("ga_ng", "non-GAAP 一般及行政費用率（近四季）", "%", None, round(ttm("ngga") / rv_t, 6), None, "Derived", "近四季", "全期持平")
inp("opx_st", "non-GAAP 營業費用率每年變動（三項合計）", "%", None, 0.0, -0.01, "Assumed", "歷史檢視（G3 第 1 條）：三項合計占營收 FY2024 56.9%、FY2025 58.6%、近四季 56.2%，無一致下降趨勢（研發上升、銷售與行銷下降）→基準持平；高＝每年 −1.0 百分點", "歷史規模經濟主要在股票薪酬（見 sbc_st），非 non-GAAP 費用率", judg=True)
inp("opx_floor", "non-GAAP 營業費用率（三項合計）下限", "%", 0.40, 0.45, 0.50, "Assumed", "判斷：僅在 opx_st 為負的 Sens 案例生效；基準持平不觸發", "", judg=True)
inp("sbc_0", "股票薪酬占營收（近四季）", "%", None, round(ttm("sbc") / rv_t, 6), None, "Derived", "近四季 SBC ÷ 營收", "全期持平；DCF 不加回（預設）")
inp("sbc_st", "股票薪酬占營收每年變動（歷史規模經濟：FY2023 22.7% → 近四季 20.1%，斜率 ÷ 2.5 年；G3 第 1 條）", "%", 0.0, round((ttm("sbc") / rv_t - fv("is.sbc.fy2023") / fv("is.rev.fy2023")) / 2.5, 6), -0.02, "Derived", "近四季與 FY2023 的 SBC 占營收差 ÷ 2.5 年（事實 is.sbc.*、is.rev.*）", "基準＝歷史斜率延伸（固定＋變動成本）；下限 sbc_floor；Sens 低＝全期持平（下緣）", judg=True)
inp("sbc_floor", "股票薪酬占營收下限（成熟大型軟體公司區間下緣）", "%", 0.08, 0.12, 0.16, "Assumed", "判斷：延伸斜率的下限，避免無限下降", "G3 第 1 條", judg=True)
inp("ptx_sh", "員工股票交易薪資稅 ÷ 股票薪酬（近四季）", "%", None, round(ttm("ptx") / ttm("sbc"), 6), None, "Derived", "近四季")
inp("oth_adj", "取得無形資產攤銷＋併購費用占營收（近四季）", "%", None, round((ttm("amort") + ttm("mna")) / rv_t, 6), None, "Derived", "近四季")
inp("gp_adj", "GAAP 與 non-GAAP 毛利差占營收（近四季）", "%", None, round((ttm("ngp") - (fv("q.rev.4q25") - fv("q.cogs.4q25")) - fv("is.gp.9m26")) / rv_t, 6), None, "Derived", "近四季 non-GAAP 毛利 − GAAP 毛利 ÷ 營收", "只影響營業成本與營業費用的分類，不影響營業利益")
inp("y_cash", "現金收益率（期初現金＋有價證券）", "%", 0.03, 0.043, 0.05, "Derived", "9M26 利息收入及其他 157.363 年化 ÷ 平均現金約 4,854（" + fs("is.intinc.9m26") + "）", "進現金流 Assumed，給低／高（原則第 22 條）")
inp("int_exp", "利息費用（可轉換票據發行成本攤銷；票面 0%）", "US$M", None, round(fv("is.iss.9m26") * 4 / 3, 3), None, "Derived", fs("is.iss.9m26"), "年化")
inp("tax_q4", "GAAP 所得稅率 Q4 26E", "%", None, 0.10, None, "Assumed", "9M26 有效稅率 8.4%（19.794 ÷ 235.242）", "")
inp("tax_27", "GAAP 所得稅率 FY2027", "%", 0.10, 0.15, 0.21, "Assumed", "FY2025 15.2%、FY2024 9.9%；美國遞延所得稅資產全額備抵（" + fs("tax.paid.fy2025") + "）", "", judg=True)
inp("tax_lr", "GAAP 所得稅率長期（FY2030 起；自 FY2027 線性）", "%", 0.18, 0.21, 0.25, "Assumed", "公司 non-GAAP 長期稅率 21%（" + fs("fx.taxrate.ng") + "）", "原則第 40 條：現金稅率與模型稅率並列見 Recon", judg=True, inc=None)
inp("tax_ng", "non-GAAP 稅率", "%", None, 0.21, None, "Interested-party", fs("fx.taxrate.ng"), "", inc=INC_CO)
inp("dil", "稀釋股數年增（股票薪酬；無回購）", "%", 0.015, 0.025, 0.035, "Assumed", "基本加權股數 9M26 較 9M25 +2.9%、2025 年較 2024 年 +3.3%（事實 is.bsh.*）", "", judg=True)
inp("da_r", "折舊攤銷占營收", "%", 0.015, 0.017, 0.020, "Derived", "9M26 58.202 ÷ 3,324.970＝1.75%（" + fs("is.da.9m26") + "）", "", judg=True)
inp("cx_r", "資本支出＋資本化軟體占營收", "%", 0.03, 0.035, 0.04, "Interested-party", fs("guid.cx"), "9M26 實績 3.67%（（32.119＋89.897）÷ 3,324.970）", inc=INC_CO, judg=True)
inp("wc_k", "營運資金增加 ÷ 營收增量（負＝流入；歷史：FY2024 −11.1%、FY2025 −18.3%、9M26 −1.5%，三期平均）", "%", -0.15, -0.103, 0.0, "Derived", "Hist 頁「營運資金與其他變動」÷ 營收增量，三期平均；遞延收入增加為現金流入（G3 第 2 條）", "原則第 23 條", judg=True)
sec("E. 折現率與目標價規則")
inp("rf", "無風險利率（10 年期美國公債，2026-11-13）", "%", 0.0465, fv("mkt.rf.20261113") / 100, 0.0565, "Verified", fs("mkt.rf.20261113"), "與現價同日", judg=True)
inp("rf_date", "無風險利率基準日", "日期", None, datetime.date(2026, 11, 13), None, "Verified", fs("mkt.rf.20261113"), "與現價同日")
inp("beta_si", "β 來源一：stockanalysis（5Y 月）", "x", None, fv("mkt.beta.si"), None, "Interested-party", fs("mkt.beta.si"), "", inc=INC_DATA)
inp("beta_yh", "β 來源二：Yahoo Finance（5Y 月）", "x", None, fv("mkt.beta.yh"), None, "Interested-party", fs("mkt.beta.yh"), "兩來源相同、可能同源", inc=INC_DATA)
inp("blume_a", "Blume 調整權重", "x", None, 0.67, None, "Assumed", "DCF 規格第 2 節（0.67×原始＋0.33×1）")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", 1.2, "=blume_a*MEDIAN(beta_si,beta_yh)+(1-blume_a)", 1.6, "Derived", "DCF 規格第 2 節", "", judg=True)
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 規格第 2 節預設 5%（4–6%）", "", judg=True)
inp("kd_sp", "稅前債務利差（相對 rf）", "%", None, 0.015, None, "Assumed", "票面 0%，採 rf＋利差作經濟成本（DCF 規格第 2 節）")
inp("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 規格第 3 節預設 3%（2–4%）", "", judg=True)
inp("ronic_sp", "RONIC − WACC（基準 0＝新投資價值中性）", "%", None, 0.0, 0.05, "Assumed", "DCF 規格第 3 節預設 RONIC＝WACC", "高＝+5 百分點", judg=True)
inp("sbc_add", "DCF 加回股票薪酬（稅後；0＝不加回，預設）", "x", None, 0, 1, "Assumed", "DCF 規格第 1 節預設不加回", "加回只作敏感度", judg=True)
inp("w_dcf", "DCF 權重（Comparable＝1−本值）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設各半", "", judg=True)
inp("w_pe", "Comparable 內 Forward P/E 腿的權重（其餘為 Forward P/S）", "%", 0.0, 0.5, 1.0, "Assumed", "判斷：兩個倍數各半", "", judg=True)
inp("rel_m", "相對倍數調整係數（1.00＝未調整：缺同業成長與利潤率同口徑資料）", "x", 0.8, 1.0, 1.3, "Assumed", "原則第 18、47 條：缺資料不調整並列入承重輸入", "Sens 含 0.8／1.3；DDOG 資料商三年營收成長預估 25.6% 高於同業中位數 20.8%（Interested-party）", judg=True)
sec("F. 同業倍數（stockanalysis 統計頁；Forward 分母口徑未載明；日期不一→逾期標示見 Recon）")
PEER = [("snow", "Snowflake（SNOW）", "2026-11-13"), ("net", "Cloudflare（NET）", "2026-11-13"), ("mdb", "MongoDB（MDB）", "2026-11-06"), ("dt", "Dynatrace（DT）", "2026-11-06"), ("estc", "Elastic（ESTC）", "2026-11-06")]
for k, nm, dt in PEER:
    inp(f"pm_ps_{k}", f"{nm} Forward P/S（{dt}）", "x", None, fv(f"peer.fps.{k}"), None, "Interested-party", fs(f"peer.fps.{k}"), "分母＝Forward 營收（FY 或 NTM 未載明）", inc=INC_DATA)
    inp(f"pm_pe_{k}", f"{nm} Forward P/E（{dt}）", "x", None, fv(f"peer.fpe.{k}"), None, "Interested-party", fs(f"peer.fpe.{k}"), "分母＝Forward 非 GAAP EPS（口徑未載明）", inc=INC_DATA)
    inp(f"pm_dt_{k}", f"{nm} 倍數日期", "日期", None, datetime.date.fromisoformat(dt), None, "Verified", fs(f"peer.fps.{k}"))
inp("pm_ps_ddog", "Datadog 現行 Forward P/S（資料商，2026-11-13）", "x", None, fv("peer.fps.ddog"), None, "Interested-party", fs("peer.fps.ddog"), "只作對照", inc=INC_DATA)
inp("pm_pe_ddog", "Datadog 現行 Forward P/E（資料商，2026-11-13）", "x", None, fv("peer.fpe.ddog"), None, "Interested-party", fs("peer.fpe.ddog"), "只作對照", inc=INC_DATA)
sec("G. 市場共識與情境權重")
inp("cons_tp", "賣方平均目標價（TipRanks，2026-11-12；家數頁面未載；低 171、高 335）", "US$", 171.0, fv("cons.tp.tipranks"), 335.0, "Interested-party", fs("cons.tp.tipranks"), "只作對照，不進模型", inc=INC_SELL + "；只作對照，不進模型")
inp("cons_date", "共識目標價基準日", "日期", None, datetime.date(2026, 11, 12), None, "Verified", fs("cons.tp.tipranks"), "只作對照，不進模型")
inp("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "原則第 4 節預設 25／50／25")
inp("tp_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "同上")
inp("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上")

IN_ROW = {}
r = 4
for row in ROWS:
    if row[0] == "sec":
        IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src, note, judg, inc = row
    if "未找到" in src and "未找到" not in note: note = "未找到（見來源欄）；" + note
    if key in OVR: base = OVR[key]
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            c.number_format = "yyyy-mm-dd" if unit == "日期" else ("0.0%" if unit == "%" else "#,##0.0000")
    IN.cell(r, 7, tag); IN.cell(r, 8, __import__('re').sub(r'https?://([^/\s（）]+)\S*', r'\1', src) if '未找到' in src else src); IN.cell(r, 9, note)
    if tag == "Interested-party": IN.cell(r, 10, inc or INC_CO)
    if judg:
        for cc in range(1, 10): IN.cell(r, cc).fill = KEY
    dn(key, f"In!$E${r}"); IN_ROW[key] = r; r += 1
IN_LAST = r - 1
r += 1
IN.cell(r, 1, "J. 情境驅動對照（三情境並排，進主模型的加權目標價；本節只引用上方各列，修改請改上方藍字。上方「敏感度低／高」欄只用於 Sens 頁，不進目標價）").font = BOLD; r += 1
hdr(IN, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]); SCN_H = r; r += 1
SCN = {}
for k, lab, unit, base_key, why in (("scn_vn", "非 AI 客戶用量成長 FY2027", "%", "vn_27", "引用 In 頁 vn_27_A／vn_27_B／vn_27_C；分化依據：近五季非 AI 成長 high 20s 的延續或減速（Q3 法說）"),
                                  ("scn_va", "AI 原生客戶用量成長 FY2027", "%", "va_27", "引用 In 頁 va_27_A／va_27_B／va_27_C；分化依據：代理工作負載放大 vs 最大客戶停留最低承諾"),
                                  ("scn_pa", "AI 原生客戶單價 FY2027 變動", "%", "pa", "引用 In 頁 pa_A／pa_B／pa_C；分化依據：提前續約讓利幅度（公司未揭露）"),
                                  ("scn_ai", "AI 產品 credits 營收 FY2027", "US$M", "ai_27", "引用 In 頁 ai_27_A／ai_27_B／ai_27_C；分化依據：AI 產品採用速度（公司未揭露營收）")):
    IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, unit)
    for j, s_ in zip((4, 5, 6), "ABC"):
        IN.cell(r, j, f"=E{IN_ROW[base_key + '_' + s_]}").number_format = "0.0%" if unit == "%" else "#,##0.0"
    IN.cell(r, 7, "Assumed"); IN.cell(r, 8, why); SCN[k] = r; r += 1
IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重"); IN.cell(r, 3, "%")
for j, s_ in zip((4, 5, 6), "ABC"): IN.cell(r, j, f"=E{IN_ROW['tp_w' + s_]}").number_format = "0%"
IN.cell(r, 7, "Assumed"); IN.cell(r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；原則第 4 節預設"); SCN["scn_w"] = r; r += 1
for k, w_ in (("A", 8), ("B", 14), ("C", 16)): pass
IN.column_dimensions["A"].width = 12; IN.column_dimensions["B"].width = 62
for c_ in "CDEFG": IN.column_dimensions[c_].width = 12
IN.column_dimensions["H"].width = 70; IN.column_dimensions["I"].width = 50; IN.column_dimensions["J"].width = 40

exec(open(os.path.join(HERE, "build_datadog_p2.py"), encoding="utf-8").read())
