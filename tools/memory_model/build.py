# 記憶體公司收支與估值樣板 v0.2（AI 半導體 Project）
# 用法：python3 tools/memory_model/build.py <公司設定檔 company.py> <輸出 raw.xlsx>
# 公司設定檔只放參數（幣別、會計年度、季度實績、輸入值）；結構全部在本檔與 markets/memory/market.py。
import sys, os, importlib.util
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MKT

def load(path):
    spec = importlib.util.spec_from_file_location("company", path)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m.CONFIG

C = load(sys.argv[1])
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
U = C["unit"]            # 例：兆韓元、十億美元
CCY = C["ccy_per_usdB"]  # 1 十億美元 換算成公司單位的公式
SS = C["share_scale"]    # 每股金額＝金額 ÷ 股數(百萬) × SS
BASE = C["base_fy"]; OFF = C["fy_offset"]
FYL = C["fy_label"]      # 例："{}"、"FY{}"
YRS = [BASE + 1, BASE + 2]

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = f"In — 輸入（藍字＝基準值；低／高供敏感度）。每列具名範圍＝鍵（指向 E 欄）。單位：{U}"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源／說明"])
rows = []
sec = lambda t: rows.append(("sec", t))
inp = lambda *a: rows.append(a)
for s in C["inputs"]:
    if s[0] == "sec": sec(s[1])
    else: inp(*s)
MKT.inputs(sec, inp)
MKT.ref_inputs(sec, inp)
r = 4
for s in rows:
    if s[0] == "sec":
        IN.cell(r, 1, s[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = s
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None: IN.cell(r, j, v).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, src)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGH", [12, 46, 11, 8, 10, 8, 20, 120]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

ref = MKT.write(wb, hdr, Font)
PATH, MK = ref["path"], ref["mk"]
MYC = {2025: "C", 2026: "D", 2027: "E", 2028: "F"}

def fy_quarters(y):
    """會計年度 y 的四個日曆季（0..15 索引）"""
    s = (y - 2025) * 4 - OFF
    return list(range(s, s + 4))
def cy_of(i): return 2025 + i // 4
def path_avg(kind, s, y):
    q = fy_quarters(y)
    return f"AVERAGE(Path!{MKT.qcol(q[0])}{PATH[kind][s]}:Path!{MKT.qcol(q[-1])}{PATH[kind][s]})"
def hbm_fy(y, s):
    terms = [f"Mkt!{MYC[cy_of(i)]}{MK['m' + s]}/4*{C['hbm_share'][cy_of(i)]}" for i in fy_quarters(y)]
    return "(" + "+".join(terms) + f")*{CCY}"
def hbmbits_fy(y):
    return "(" + "+".join(f"Mkt!{MYC[cy_of(i)]}{MK['ship']}/4*{C['hbm_share'][cy_of(i)]}" for i in fy_quarters(y)) + ")"

# ---------------- Q：季度 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = f"Q — 季度損益（{U}）。實績：{C['q_source']}；推估季：實績 × 公司指引 × TrendForce 價格"
qs = C["quarters"]
hdr(Q, 3, ["項目"] + [q["label"] for q in qs] + [f"{FYL.format(BASE)} 合計"])
QROWS = ["rev", "dram", "nand", "oth", "cogs", "opex", "op", "opm", "ni", "da", "dbi", "nbi"]
QLAB = {"rev": "營收", "dram": "  DRAM（含 HBM）", "nand": "  NAND", "oth": "  其他", "cogs": "營業成本", "opex": "營業費用", "op": "營業利益",
        "opm": "營業利益率", "ni": "淨利", "da": "折舊攤銷", "dbi": "DRAM 位元指數（最後實績季＝1）", "nbi": "NAND 位元指數"}
QR = {k: 4 + i for i, k in enumerate(QROWS)}
for k in QROWS: Q.cell(QR[k], 1, QLAB[k])
QC = [chr(ord("B") + i) for i in range(len(qs))]
last_act = max(i for i, q in enumerate(qs) if q["type"] == "A")
for i, q in enumerate(qs):
    c = QC[i]
    if q["type"] == "A":
        for k in ("rev", "dram", "nand", "cogs", "op", "ni", "da"):
            v = q.get(k)
            if v is None: continue
            cell = Q[f"{c}{QR[k]}"]; cell.value = v; cell.number_format = "0.00"
            if isinstance(v, (int, float)): cell.font = BLUE
        if q.get("dram") is not None:
            Q[f"{c}{QR['oth']}"] = f"={c}{QR['rev']}-{c}{QR['dram']}-{c}{QR['nand']}"
        if q.get("cogs") is not None:
            Q[f"{c}{QR['opex']}"] = f"={c}{QR['rev']}-{c}{QR['cogs']}-{c}{QR['op']}"
    else:
        p = QC[i - 1]; g = q["keys"]
        Q[f"{c}{QR['dram']}"] = f"={p}{QR['dram']}*(1+{g['db']})*(1+{g['dp']})"
        Q[f"{c}{QR['nand']}"] = f"={p}{QR['nand']}*(1+{g['nb']})*(1+{g['np']})"
        Q[f"{c}{QR['oth']}"] = "=other_y/4"
        Q[f"{c}{QR['rev']}"] = f"={c}{QR['dram']}+{c}{QR['nand']}+{c}{QR['oth']}"
        Q[f"{c}{QR['dbi']}"] = f"={p}{QR['dbi']}*(1+{g['db']})"
        Q[f"{c}{QR['nbi']}"] = f"={p}{QR['nbi']}*(1+{g['nb']})"
        Q[f"{c}{QR['cogs']}"] = (f"=({p}{QR['cogs']}-da_cogs*{p}{QR['da']})*((w_c+w_h)*{c}{QR['dbi']}/{p}{QR['dbi']}+w_n*{c}{QR['nbi']}/{p}{QR['nbi']})*(1-costdn_q)+da_cogs*{g['da']}")
        Q[f"{c}{QR['opex']}"] = f"=opex_fix_q+opex_var*{c}{QR['rev']}"
        Q[f"{c}{QR['op']}"] = f"={c}{QR['rev']}-{c}{QR['cogs']}-{c}{QR['opex']}"
        Q[f"{c}{QR['ni']}"] = f"=({c}{QR['op']}+nonop_q)*(1-tax)"
        Q[f"{c}{QR['da']}"] = f"={g['da']}"
Q[f"{QC[last_act]}{QR['dbi']}"] = 1; Q[f"{QC[last_act]}{QR['nbi']}"] = 1
for c in QC:
    Q[f"{c}{QR['opm']}"] = f"=IF({c}{QR['rev']}=0,\"\",{c}{QR['op']}/{c}{QR['rev']})"; Q[f"{c}{QR['opm']}"].number_format = "0.0%"
TOT = chr(ord("B") + len(qs))
b0, b1 = QC[-4], QC[-1]   # 基準年＝最後四季
pb0, pb1 = (QC[-8], QC[-5]) if len(qs) >= 8 else (None, None)
for k in QROWS:
    if k in ("dbi", "nbi"): continue
    Q[f"{TOT}{QR[k]}"] = f"=SUM({b0}{QR[k]}:{b1}{QR[k]})" if k != "opm" else f"={TOT}{QR['op']}/{TOT}{QR['rev']}"
    Q[f"{TOT}{QR[k]}"].number_format = "0.0%" if k == "opm" else "0.00"
Q.cell(18, 1, C.get("q_note", ""))
Q.column_dimensions["A"].width = 30
for c in QC + [TOT]: Q.column_dimensions[c].width = 11

# ---------------- Y：年度（基準年＋2 年 × 3 情境）----------------
Y = wb.create_sheet("Y")
Y["A1"] = f"Y — 年度（{U}；股數百萬股）。情境 A：{MKT.SCEN_DESC['A']}；B：{MKT.SCEN_DESC['B']}；C：{MKT.SCEN_DESC['C']}"
cols = [("B", FYL.format(BASE - 1) + " 實績", None, None), ("C", FYL.format(BASE) + "（基準）", None, None)]
FC = {}
letters = iter("DEFGHI")
for s in MKT.SCEN:
    for n, y in enumerate(YRS, 1):
        L = next(letters); FC[(s, n)] = L
        cols.append((L, f"{FYL.format(y)} {s}", s, n))
hdr(Y, 3, ["項目"] + [c[1] for c in cols] + ["說明"])
YR = {k: 4 + i for i, k in enumerate(["rev", "conv", "hbm", "nand", "oth", "hbmsh", "da", "vol", "cogs", "opex", "op", "opm", "ni", "capex", "fcf", "div", "bb", "eq", "nc", "sh", "bvps", "eps", "roe", "hbmbits"])}
YLAB = {"rev": "營收", "conv": "  一般 DRAM", "hbm": "  HBM", "nand": "  NAND", "oth": "  其他", "hbmsh": "  HBM ÷ 營收", "da": "折舊攤銷",
        "vol": "成本量指數（÷ 前一年）", "cogs": "營業成本", "opex": "營業費用", "op": "營業利益", "opm": "營業利益率", "ni": "淨利", "capex": "資本支出",
        "fcf": "自由現金流（淨利＋折舊 − 資本支出）", "div": "股利", "bb": "買回", "eq": "期末權益", "nc": "期末淨現金", "sh": "期末流通股數",
        "bvps": "每股淨值 BVPS", "eps": "每股盈餘 EPS（期末股數）", "roe": "ROE（期末權益）", "hbmbits": "HBM 出貨位元（EB，公司）"}
for k, rr in YR.items(): Y.cell(rr, 1, YLAB[k])
def setc(col, key, f, fmt="0.0"):
    c = Y[f"{col}{YR[key]}"]; c.value = f; c.number_format = fmt
R = YR
# 前一年實績（只有營收、營業利益、淨利、折舊）
if pb0:
    for k in ("rev", "op", "ni", "da"):
        setc("B", k, f"=SUM(Q!{pb0}{QR[k]}:{pb1}{QR[k]})")
    setc("B", "opm", f"=B{R['op']}/B{R['rev']}", "0.0%")
# 基準年
qs_ = lambda k: f"=Q!{TOT}{QR[k]}"
setc("C", "rev", qs_("rev")); setc("C", "hbm", "=" + hbm_fy(BASE, "A")); setc("C", "conv", f"=Q!{TOT}{QR['dram']}-C{R['hbm']}")
setc("C", "nand", qs_("nand")); setc("C", "oth", qs_("oth")); setc("C", "da", qs_("da")); setc("C", "cogs", qs_("cogs"))
setc("C", "opex", qs_("opex")); setc("C", "op", qs_("op")); setc("C", "ni", qs_("ni")); setc("C", "capex", "=cap_0")
setc("C", "hbmbits", "=" + hbmbits_fy(BASE), "0.00")
fq = [i for i, q in enumerate(qs[-4:]) if q["type"] == "F"]
fni = "+".join(f"Q!{qs and QC[-4 + i]}{QR['ni']}" for i in fq) or "0"
fda = "+".join(f"Q!{QC[-4 + i]}{QR['da']}" for i in fq) or "0"
setc("C", "div", "=div_0"); setc("C", "bb", "=bb_0")
setc("C", "eq", f"=eq_last+({fni})+extra_eq-C{R['bb']}-C{R['div']}")
setc("C", "nc", f"=nc_last+({fni})+({fda})-capex_rem+extra_eq-C{R['bb']}-C{R['div']}")
setc("C", "sh", "=sh_last+extra_sh-bb_sh_0", "0.0")
for k in ("hbmsh", "opm", "fcf", "bvps", "eps", "roe"):
    pass
def common(col, P):
    setc(col, "hbmsh", f"={col}{R['hbm']}/{col}{R['rev']}", "0.0%")
    setc(col, "opm", f"={col}{R['op']}/{col}{R['rev']}", "0.0%")
    setc(col, "fcf", f"={col}{R['ni']}+{col}{R['da']}-{col}{R['capex']}")
    setc(col, "bvps", f"={col}{R['eq']}/{col}{R['sh']}*{SS}", "#,##0.00" if SS < 1e5 else "#,##0")
    setc(col, "eps", f"={col}{R['ni']}/{col}{R['sh']}*{SS}", "#,##0.00" if SS < 1e5 else "#,##0")
    setc(col, "roe", f"={col}{R['ni']}/{col}{R['eq']}", "0%")
common("C", None)
for (s, n), col in FC.items():
    y = YRS[n - 1]
    P = "C" if n == 1 else FC[(s, 1)]
    setc(col, "conv", f"={P}{R['conv']}*(1+cb_{n})*({path_avg('conv', s, y)}/{path_avg('conv', s, y - 1)})")
    setc(col, "hbm", "=" + hbm_fy(y, s))
    setc(col, "nand", f"={P}{R['nand']}*(1+nb_{n})*({path_avg('nand', s, y)}/{path_avg('nand', s, y - 1)})")
    setc(col, "oth", "=other_y")
    setc(col, "rev", f"={col}{R['conv']}+{col}{R['hbm']}+{col}{R['nand']}+{col}{R['oth']}")
    setc(col, "hbmbits", "=" + hbmbits_fy(y), "0.00")
    setc(col, "da", f"={P}{R['da']}+({P}{R['capex']}-{P}{R['da']})/da_life")
    setc(col, "vol", f"=w_c*(1+cb_{n})+w_h*{col}{R['hbmbits']}/{P}{R['hbmbits']}+w_n*(1+nb_{n})", "0.00")
    setc(col, "cogs", f"=({P}{R['cogs']}-da_cogs*{P}{R['da']})*{col}{R['vol']}*(1-costdn_y)+da_cogs*{col}{R['da']}")
    setc(col, "opex", f"=opex_fix_y*(1+opex_g)^{n}+opex_var*{col}{R['rev']}")
    setc(col, "op", f"={col}{R['rev']}-{col}{R['cogs']}-{col}{R['opex']}")
    setc(col, "ni", f"=({col}{R['op']}+rf*{P}{R['nc']})*(1-tax)")
    setc(col, "capex", f"=cap_{n}")
    setc(col, "div", f"=dps_{n}*{P}{R['sh']}/{SS}")
    setc(col, "bb", f"=MAX(0,payout*{col}{R['fcf']}-{col}{R['div']})")
    setc(col, "eq", f"={P}{R['eq']}+{col}{R['ni']}-{col}{R['div']}-{col}{R['bb']}")
    setc(col, "nc", f"={P}{R['nc']}+{col}{R['fcf']}-{col}{R['div']}-{col}{R['bb']}")
    setc(col, "sh", f"={P}{R['sh']}-{col}{R['bb']}/bb_px*{SS}", "0.0")
    common(col, P)
Y.column_dimensions["A"].width = 34
for c in "BCDEFGHI": Y.column_dimensions[c].width = 12
notes = {"conv": "前一年 × (1＋位元) × 會計年度價格指數平均比（Path）", "hbm": "Σ 會計年度各季：HBM 市場 ÷ 4 × 市占 × 匯率（Mkt）",
         "cogs": "現金成本 × 量 × (1 − 年降)＋折舊", "ni": "（營業利益＋淨現金 × 收益率）× (1 − 稅率)", "bb": "回饋比例 × FCF − 股利"}
for k, t in notes.items(): Y.cell(R[k], 10, t)
Y.column_dimensions["J"].width = 50

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值（P/B 主、P/E 對照）。隱含股價 ÷ 目前股價 − 1"
hdr(V, 3, ["方法", "倍數", "基準", "情境", "每股基準", "隱含股價", "相對目前"])
vr = 4
pbs = [(C["pb_names"][i], f"pb_{i+1}") for i in range(3)]
pes = [(f"P/E {C['pe_vals'][i]}x", f"pe_{i+1}") for i in range(3)]
def vrow(nm, k, basis, sc, ref, fmtp):
    global vr
    V.cell(vr, 1, nm); V.cell(vr, 2, f"={k}"); V.cell(vr, 3, basis); V.cell(vr, 4, sc)
    V.cell(vr, 5, f"=Y!{ref}").number_format = fmtp
    V.cell(vr, 6, f"=B{vr}*E{vr}").number_format = fmtp
    V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
    vr += 1
fmtp = "#,##0" if SS > 1e5 else "#,##0.00"
for nm, k in pbs: vrow(nm, k, FYL.format(BASE) + " 期末 BVPS", "共用", f"C{R['bvps']}", fmtp)
for n, y in enumerate(YRS, 1):
    for s in MKT.SCEN:
        for nm, k in pbs: vrow(nm, k, FYL.format(y) + " 期末 BVPS", s, f"{FC[(s, n)]}{R['bvps']}", fmtp)
for n, y in enumerate(YRS, 1):
    for s in MKT.SCEN:
        for nm, k in pes: vrow(nm, k, FYL.format(y) + " EPS", s, f"{FC[(s, n)]}{R['eps']}", fmtp)
for c, w in zip("ABCDEFG", [24, 8, 20, 8, 16, 14, 10]): V.column_dimensions[c].width = w

# ---------------- Sens（營收變動 × (1 − 變動營業費用率)；成本不隨價格變動，故為精確值）----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次只動一個價格或份額輸入；Δ營業利益＝Δ營收 × (1 − 營業費用變動率)（成本只隨量變動，所以這是精確值，不是近似）"
hdr(SE, 3, ["欄", "輸入變動", "Δ營收", "營業利益", "相對基準"])
sr = 4
for (s, n) in (("A", 1), ("B", 2), ("C", 2)):
    col = FC[(s, n)]; lab = f"{FYL.format(YRS[n-1])} {s}"
    SE.cell(sr, 1, lab).font = BOLD; SE.cell(sr, 2, "基準"); SE.cell(sr, 4, f"=Y!{col}{R['op']}").number_format = "0.0"; base_r = sr; sr += 1
    items = [("一般 DRAM 收入 −20%", f"=-0.2*Y!{col}{R['conv']}"), ("一般 DRAM 收入 +20%", f"=0.2*Y!{col}{R['conv']}"),
             ("HBM 收入 −30%（價格或份額）", f"=-0.3*Y!{col}{R['hbm']}"), ("HBM 收入 +30%", f"=0.3*Y!{col}{R['hbm']}"),
             ("NAND 收入 −30%", f"=-0.3*Y!{col}{R['nand']}"), ("NAND 收入 +30%", f"=0.3*Y!{col}{R['nand']}")]
    for nm, f in items:
        SE.cell(sr, 1, lab); SE.cell(sr, 2, nm); SE.cell(sr, 3, f).number_format = "0.0"
        SE.cell(sr, 4, f"=D{base_r}+C{sr}*(1-opex_var)").number_format = "0.0"
        SE.cell(sr, 5, f"=D{sr}/D{base_r}-1").number_format = "0%"
        sr += 1
    sr += 1
SE.cell(sr, 1, "讀法：哪一條收入線的同幅變動對營業利益影響最大，就是該年度的主要驅動。價格路徑本身的情境差異見 Y 頁 A／B／C 欄。")
SE.column_dimensions["A"].width = 14; SE.column_dimensions["B"].width = 30

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = f"Recon — 本模型（正向推導）vs 共識或公司指引（只作事後對照，不用來回推參數）；{U}"
hdr(RC, 3, ["項目", "本模型", "對照值", "差距", "對照來源"])
for i, (nm, f, v, src) in enumerate(C["recon"]):
    rr = 4 + i
    RC.cell(rr, 1, nm); RC.cell(rr, 2, f).number_format = "0.0"
    RC.cell(rr, 3, v).font = BLUE
    RC.cell(rr, 4, f"=B{rr}/C{rr}-1").number_format = "0%"; RC.cell(rr, 5, src)
RC.column_dimensions["A"].width = 46; RC.column_dimensions["E"].width = 50

# ---------------- Feedback、Sources、README ----------------
FB = wb.create_sheet("Feedback"); FB["A1"] = "Feedback — 回饋算力收支"; FB["A1"].font = BOLD
fb = [("每顆加速器 HBM 金額（GB300，2026 價）", "=nv_hbm", "$K"),
      ("全市場每顆加速器平均 HBM 2026／2027", f"=TEXT(Mkt!D{MK['gbavg']},\"0\")&\" ／ \"&TEXT(Mkt!E{MK['gbavg']},\"0\")", "GB"),
      ("HBM 市場 2026／2027（A）", f"=TEXT(Mkt!D{MK['mA']},\"0\")&\" ／ \"&TEXT(Mkt!E{MK['mA']},\"0\")", "$B"),
      ("2027 HBM 需求缺口", f"=Mkt!E{MK['gap']}", "EB"),
      (f"{C['name']} HBM 毛利率", "無法由公開資料分離", "—")]
for i, row in enumerate(fb, 3):
    for j, v in enumerate(row, 1): FB.cell(i, j, v)
FB.column_dimensions["A"].width = 40; FB.column_dimensions["B"].width = 20
SO = wb.create_sheet("Sources"); SO["A1"] = "來源"; SO["A1"].font = BOLD
for i, (a, b) in enumerate(C["sources"] + [("markets/memory/market.py（共用市場層）", "本 repo"), ("markets/chip_bridge v0.1", "本 repo")], 2):
    SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 60; SO.column_dimensions["B"].width = 120
RD = wb.create_sheet("README", 0)
lines = [(f"{C['name']} 收支與估值模型 {C['version']}（記憶體樣板 v0.2）　{C['date']}", True), ("模型命題：" + C["thesis"], False), ("", False),
 ("驅動因子 → 推導量", True),
 ("Path（共用市場層）：一般 DRAM、NAND 合約價季增 → 季度指數（情境 A／B／C）；各公司依會計年度取平均，算年均價變動", False),
 ("Mkt（共用市場層）：加速器顆數 × 每顆 GB → 需求位元；MIN(需求, 供給) → 出貨；× 單價 → HBM 市場（日曆年）", False),
 ("Q：最後實績季 × 公司指引 × TrendForce 季價 → 基準年剩餘季度", False),
 ("Y：前一年 × (1＋位元) × 價格指數比 → 一般 DRAM、NAND；Σ 各季 HBM 市場 × 市占 → HBM；成本量 × (1 − 年降)＋折舊 → 營業利益 → 淨利 → 權益 → 每股淨值", False),
 ("Val：每股淨值 × P/B（主）、EPS × P/E（對照）", False), ("Sens：各收入線同幅變動對營業利益的影響", False), ("Recon：與共識、公司指引事後對照", False), ("", False),
 (f"情境：A＝{MKT.SCEN_DESC['A']}；B＝{MKT.SCEN_DESC['B']}；C＝{MKT.SCEN_DESC['C']}", False),
 (f"幣別單位：{U}；會計年度：{C['fy_desc']}", False), ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False), ("", False)] + [(t, False) for t in C["readme_extra"]]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170
wb.save(sys.argv[2])
print("saved", sys.argv[2])
