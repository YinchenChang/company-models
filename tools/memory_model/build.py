# 記憶體公司收支與估值樣板 v0.4（v0.4：DCF＋Comparable 目標價、三情境延伸到 FY2032 收斂法；v0.3：預測 3 年、非記憶體事業列）（AI 半導體 Project）
# 用法：python3 tools/memory_model/build.py <公司設定檔 company.py> <輸出 raw.xlsx>
# 公司設定檔只放參數（幣別、會計年度、季度實績、輸入值）；結構全部在本檔與 markets/memory/market.py。
import sys, os, importlib.util
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MKT
sys.path.insert(0, os.path.join(ROOT, "tools"))
import valuation_rule as VR

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
YRS = [BASE + 1, BASE + 2, BASE + 3]

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
# 目標價參數（Andy 2026-10-25）：放在 In 頁最後，不位移既有列（v0.4 樣板：DCF＋Comparable；黃底）
TPC = C["tp"]
sec("目標價參數（黃底；Val 頁第 0 節 DCF＋Comparable 與收斂法用；Assumed 者 Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}（A 順利／B 平平／C 不如預期）", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）；三者合計應為 100%（Val 頁第 0 節列示）")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；美、台、日、韓同一值）")
inp("tp_g", "長期成長率（永續、收斂法 B 的 FY2032 營收年增）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）")
inp("tp_roe", TPC["roe_label"], "%", None, TPC["roe"], None, "Derived", TPC["roe_src"])
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）")
inp("nim10", "十年平均淨利率（2016–2025；淨利 ÷ 合併營收）", "%", None, TPC["nim10"], None, "Derived", TPC["nim10_src"])
for s_, f_, t_ in (("A", "=tp_g+0.02", "tp_g＋2pt"), ("B", "=tp_g", "＝tp_g"), ("C", "=tp_g-0.02", "tp_g−2pt")):
    inp(f"fade_g{s_}", f"情境 {s_}：FY2032 合併營收年增收斂目標（{t_}）", "%", None, f_, None, "Assumed", "工作單 2026-10-25 11:20 收斂法：FY2030 起由 FY2029 年增線性收斂，FY2032 達到")
# 記憶體為循環股：收斂期改收斂到「循環中段」淨利率（主串 2026-10-25 定預設；Andy 可改），不延續高峰利潤率
inp("cyc_mA", "循環中段係數：情境 A 淨利率收斂目標 ÷ 十年平均", "x", None, 1.2, None, "Assumed", "Andy 可改（預設 1.2；記憶體循環股，收斂期不延續 FY2029 高峰利潤率）")
inp("cyc_mC", "循環中段係數：情境 C 淨利率收斂目標 ÷ 十年平均", "x", None, 0.8, None, "Assumed", "Andy 可改（預設 0.8）")
for s_, f_, t_ in (("A", "=nim10*cyc_mA", "十年平均 × cyc_mA"), ("B", "=nim10", "＝十年平均淨利率"), ("C", "=nim10*cyc_mC", "十年平均 × cyc_mC")):
    inp(f"fade_m{s_}", f"情境 {s_}：FY2032 淨利率收斂目標（循環中段：{t_}）", "%", None, f_, None, "Derived", "In 頁 nim10 × 循環中段係數；FY2030 起由 FY2029 淨利率線性收斂，FY2032 達到")
PXU = next(x[2] for x in C["inputs"] if x[0] == "px")
inp("cons_tp", "市場共識目標價（賣方平均；低／高＝最低／最高）", PXU, TPC["cons_tp"][1], TPC["cons_tp"][0], TPC["cons_tp"][2], "Interested-party", TPC["cons_tp"][3])
for k_, lab_, v_, s_ in TPC["cons"]:
    inp(k_, lab_, PXU, None, v_, None, "Interested-party", s_)
TPKEYS = ["tp_wA", "tp_wB", "tp_wC", "tp_r", "tp_g", "tp_roe", "tp_w_dcf", "nim10", "fade_gA", "fade_gB", "fade_gC", "cyc_mA", "cyc_mC", "fade_mA", "fade_mB", "fade_mC", "cons_tp"] + [x[0] for x in TPC["cons"]]
YEL = PatternFill("solid", fgColor="FFF2CC")
INROW = {}
r = 4
for s in rows:
    if s[0] == "sec":
        IN.cell(r, 1, s[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = s
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None: IN.cell(r, j, v).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, src)
    INROW[key] = r
    if key in TPKEYS:
        for j in range(1, 9): IN.cell(r, j).fill = YEL
        if isinstance(base, str) and base.startswith("="): IN.cell(r, 5).font = Font(color="000000")
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGH", [12, 46, 11, 8, 10, 8, 20, 120]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

ref = MKT.write(wb, hdr, Font)
PATH, MK = ref["path"], ref["mk"]
MYC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}

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
def hbmbits_fy(y, s):
    return "(" + "+".join(f"Mkt!{MYC[cy_of(i)]}{MK['ship' + s]}/4*{C['hbm_share'][cy_of(i)]}" for i in fy_quarters(y)) + ")"

# ---------------- Q：季度 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = f"Q — 季度損益（{U}）。實績：{C['q_source']}；推估季：實績 × 公司指引 × TrendForce 價格"
qs = C["quarters"]
hdr(Q, 3, ["項目"] + [q["label"] for q in qs] + [f"{FYL.format(BASE)} 合計"])
QROWS = ["rev", "dram", "nand", "oth", "cogs", "opex", "op", "opm", "ni", "da", "dbi", "nbi"]
QLAB = {"rev": "記憶體營收", "dram": "  DRAM（含 HBM）", "nand": "  NAND", "oth": "  其他", "cogs": "營業成本", "opex": "營業費用", "op": "記憶體營業利益",
        "opm": "記憶體營業利益率", "ni": "淨利（合併，含非記憶體事業）", "da": "折舊攤銷", "dbi": "DRAM 位元指數（最後實績季＝1）", "nbi": "NAND 位元指數"}
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
        Q[f"{c}{QR['ni']}"] = f"=({c}{QR['op']}+nonop_q+nm_op_q)*(1-tax)"
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
# 官方報告分部（v0.4.1 樣板，選用）：實績季的分部營收與核對列「分部合計 − 合併營收」（分部原則第 3 條）。
# 推估季不拆分部（模型以技術別 DRAM／NAND 推估）；分部列只放公司已揭露的實績。
SEG = C.get("segments")
if SEG:
    r0 = 20
    Q.cell(r0, 1, SEG["title"]).font = BOLD
    rows_ = []
    for j, (lab, vals) in enumerate(SEG["rows"]):
        r = r0 + 1 + j; rows_.append(r); Q.cell(r, 1, lab)
        for i, v in enumerate(vals):
            if v is None: continue
            cell = Q[f"{QC[i]}{r}"]; cell.value = v; cell.number_format = "0.000"
            if isinstance(v, (int, float)): cell.font = BLUE
        if all(vals[i] is not None for i in range(len(qs) - 4, len(qs))):
            Q[f"{TOT}{r}"] = f"=SUM({b0}{r}:{b1}{r})"; Q[f"{TOT}{r}"].number_format = "0.000"
    rc = r0 + 1 + len(SEG["rows"])
    tot_ok = lambda i: True
    if SEG.get("total"):   # 合併營收另列（記憶體營收≠合併營收的集團公司，如三星）
        lab, vals = SEG["total"]; Q.cell(rc, 1, lab)
        for i, v in enumerate(vals):
            if v is not None:
                Q[f"{QC[i]}{rc}"] = v; Q[f"{QC[i]}{rc}"].font = BLUE; Q[f"{QC[i]}{rc}"].number_format = "0.000"
        tot_row = rc; tot_ok = lambda i: vals[i] is not None; rc += 1
    else:
        tot_row = QR["rev"]
    Q.cell(rc, 1, SEG["check_label"])
    full = [i for i in range(len(qs)) if tot_ok(i) and all(SEG["rows"][j][1][i] is not None for j in range(len(rows_)))]
    for i in full:
        c = QC[i]
        Q[f"{c}{rc}"] = "=" + "+".join(f"{c}{r}" for r in rows_) + f"-{c}{tot_row}"; Q[f"{c}{rc}"].number_format = "0.000"
    if not SEG.get("total"):
        Q[f"{TOT}{rc}"] = "=" + "+".join(f"{TOT}{r}" for r in rows_) + f"-{TOT}{tot_row}"; Q[f"{TOT}{rc}"].number_format = "0.000"
    for j, line in enumerate(SEG.get("notes", [])):
        Q.cell(rc + 1 + j, 1, line)
Q.column_dimensions["A"].width = 30
for c in QC + [TOT]: Q.column_dimensions[c].width = 11

# ---------------- Y：年度（基準年＋2 年 × 3 情境）----------------
Y = wb.create_sheet("Y")
Y["A1"] = f"Y — 年度（{U}；股數百萬股）。情境 A：{MKT.SCEN_DESC['A']}；B：{MKT.SCEN_DESC['B']}；C：{MKT.SCEN_DESC['C']}"
cols = [("B", FYL.format(BASE - 1) + " 實績", None, None), ("C", FYL.format(BASE) + "（基準）", None, None)]
FC = {}
letters = iter("DEFGHIJKL")
for s in MKT.SCEN:
    for n, y in enumerate(YRS, 1):
        L = next(letters); FC[(s, n)] = L
        cols.append((L, f"{FYL.format(y)} {s}", s, n))
# v0.4 樣板：FY2030–FY2032 收斂法欄（M–U，接在既有欄之後；說明欄移到 V）
FADEY = [BASE + 4, BASE + 5, BASE + 6]
FCF_ = {}
letters2 = iter("MNOPQRSTU")
for s in MKT.SCEN:
    for y in FADEY:
        L = next(letters2); FCF_[(s, y)] = L
        cols.append((L, f"{FYL.format(y)} {s} 收斂", s, y))
hdr(Y, 3, ["項目"] + [c[1] for c in cols] + ["說明"])
NOTEC = 2 + len(cols)
YR = {k: 4 + i for i, k in enumerate(["rev", "conv", "hbm", "nand", "oth", "hbmsh", "da", "vol", "cogs", "opex", "op", "opm", "nmrev", "nmop", "crev", "cop", "ni", "capex", "fcf", "div", "bb", "eq", "nc", "sh", "bvps", "eps", "roe", "hbmbits", "cg", "nim"])}
YLAB = {"nmrev": "非記憶體事業營收", "nmop": "非記憶體事業營業利益", "crev": "合併營收", "cop": "合併營業利益", "rev": "記憶體營收", "conv": "  一般 DRAM", "hbm": "  HBM", "nand": "  NAND", "oth": "  其他", "hbmsh": "  HBM ÷ 營收", "da": "折舊攤銷",
        "vol": "成本量指數（÷ 前一年）", "cogs": "營業成本", "opex": "營業費用", "op": "記憶體營業利益", "opm": "記憶體營業利益率", "ni": "淨利（合併）", "capex": "資本支出",
        "fcf": "自由現金流（淨利＋折舊 − 資本支出）", "div": "股利", "bb": "買回", "eq": "期末權益", "nc": "期末淨現金", "sh": "期末流通股數",
        "bvps": "每股淨值 BVPS", "eps": "每股盈餘 EPS（期末股數）", "roe": "ROE（期末權益）", "hbmbits": "HBM 出貨位元（EB，公司）",
        "cg": "合併營收年增", "nim": "淨利率（淨利 ÷ 合併營收）"}
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
setc("C", "nmrev", "=nm_rev_0"); setc("C", "nmop", "=nm_op_0")
setc("C", "hbmbits", "=" + hbmbits_fy(BASE, "A"), "0.00")  # 基準年：情境係數 2028 才開始，三情境相同
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
    setc(col, "crev", f"={col}{R['rev']}+{col}{R['nmrev']}")
    setc(col, "cop", f"={col}{R['op']}+{col}{R['nmop']}")
    setc(col, "hbmsh", f"={col}{R['hbm']}/{col}{R['rev']}", "0.0%")
    setc(col, "opm", f"={col}{R['op']}/{col}{R['rev']}", "0.0%")
    setc(col, "fcf", f"={col}{R['ni']}+{col}{R['da']}-{col}{R['capex']}")
    setc(col, "bvps", f"={col}{R['eq']}/{col}{R['sh']}*{SS}", "#,##0.00" if SS < 1e5 else "#,##0")
    setc(col, "eps", f"={col}{R['ni']}/{col}{R['sh']}*{SS}", "#,##0.00" if SS < 1e5 else "#,##0")
    setc(col, "roe", f"={col}{R['ni']}/{col}{R['eq']}", "0%")
common("C", None)
for (s, n), col in FC.items():
    y = YRS[n - 1]
    P = "C" if n == 1 else FC[(s, n - 1)]   # v0.4 修正：前一年＝第 n−1 年（原誤取第 1 年，FY2029 以 FY2027 為基期）
    setc(col, "conv", f"={P}{R['conv']}*(1+cb_{n})*({path_avg('conv', s, y)}/{path_avg('conv', s, y - 1)})")
    setc(col, "hbm", "=" + hbm_fy(y, s))
    setc(col, "nand", f"={P}{R['nand']}*(1+nb_{n})*({path_avg('nand', s, y)}/{path_avg('nand', s, y - 1)})")
    setc(col, "oth", "=other_y")
    setc(col, "rev", f"={col}{R['conv']}+{col}{R['hbm']}+{col}{R['nand']}+{col}{R['oth']}")
    setc(col, "hbmbits", "=" + hbmbits_fy(y, s), "0.00")
    setc(col, "da", f"={P}{R['da']}+({P}{R['capex']}-{P}{R['da']})/da_life")
    setc(col, "vol", f"=w_c*(1+cb_{n})+w_h*{col}{R['hbmbits']}/{P}{R['hbmbits']}+w_n*(1+nb_{n})", "0.00")
    setc(col, "cogs", f"=({P}{R['cogs']}-da_cogs*{P}{R['da']})*{col}{R['vol']}*(1-costdn_y)+da_cogs*{col}{R['da']}")
    setc(col, "opex", f"=opex_fix_y*(1+opex_g)^{n}+opex_var*{col}{R['rev']}")
    setc(col, "op", f"={col}{R['rev']}-{col}{R['cogs']}-{col}{R['opex']}")
    setc(col, "nmrev", f"=nm_rev_{n}"); setc(col, "nmop", f"=nm_op_{n}")
    setc(col, "ni", f"=({col}{R['cop']}+rf*{P}{R['nc']})*(1-tax)")
    setc(col, "capex", f"=cap_{n}")
    setc(col, "div", f"=dps_{n}*{P}{R['sh']}/{SS}")
    setc(col, "bb", f"=MAX(0,payout*{col}{R['fcf']}-{col}{R['div']})")
    setc(col, "eq", f"={P}{R['eq']}+{col}{R['ni']}-{col}{R['div']}-{col}{R['bb']}")
    setc(col, "nc", f"={P}{R['nc']}+{col}{R['fcf']}-{col}{R['div']}-{col}{R['bb']}")
    setc(col, "sh", f"={P}{R['sh']}-{col}{R['bb']}/bb_px*{SS}", "0.0")
    common(col, P)
# 合併營收年增、淨利率（基準年起；收斂法的起點）
setc("C", "nim", f"=C{R['ni']}/C{R['crev']}", "0.0%")
for (s, n), col in FC.items():
    P = "C" if n == 1 else FC[(s, n - 1)]
    setc(col, "cg", f"={col}{R['crev']}/{P}{R['crev']}-1", "0.0%")
    setc(col, "nim", f"={col}{R['ni']}/{col}{R['crev']}", "0.0%")
# FY2030–FY2032：收斂法（上游市場層只到 2029；非驅動因子推導）。合併營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到）；股數固定為 FY2029 期末
FADEFILL = PatternFill("solid", fgColor="E7E6E6")
for (s, y), col in FCF_.items():
    k = y - FADEY[0] + 1
    L9 = FC[(s, 3)]; P = L9 if k == 1 else FCF_[(s, y - 1)]
    setc(col, "cg", f"=${L9}{R['cg']}+(fade_g{s}-${L9}{R['cg']})*{k}/3", "0.0%")
    setc(col, "crev", f"={P}{R['crev']}*(1+{col}{R['cg']})")
    setc(col, "nim", f"=${L9}{R['nim']}+(fade_m{s}-${L9}{R['nim']})*{k}/3", "0.0%")
    setc(col, "ni", f"={col}{R['crev']}*{col}{R['nim']}")
    setc(col, "sh", f"={P}{R['sh']}", "0.0")
    setc(col, "eps", f"={col}{R['ni']}/{col}{R['sh']}*{SS}", "#,##0.00" if SS < 1e5 else "#,##0")
    for key in ("cg", "crev", "nim", "ni", "sh", "eps"): Y[f"{col}{R[key]}"].fill = FADEFILL
Y["A2"] = (f"灰底＝{FYL.format(FADEY[0])}–{FYL.format(FADEY[-1])} 收斂法（上游市場層只到 2029）：合併營收年增由 {FYL.format(BASE + 3)} 線性收斂到 fade_g、淨利率收斂到 fade_m（In 頁目標價參數），"
           "股數固定；非驅動因子推導，只算營收、淨利、EPS")
Y.column_dimensions["A"].width = 34
for c in "BCDEFGHIJKLMNOPQRSTU": Y.column_dimensions[c].width = 12
notes = {"conv": "前一年 × (1＋位元) × 會計年度價格指數平均比（Path）", "hbm": "Σ 會計年度各季：HBM 市場 ÷ 4 × 市占 × 匯率（Mkt）",
         "cogs": "現金成本 × 量 × (1 − 年降)＋折舊", "ni": "（營業利益＋淨現金 × 收益率）× (1 − 稅率)", "bb": "回饋比例 × FCF − 股利"}
notes["crev"] = f"{FYL.format(FADEY[0])} 起為收斂法：合併營收年增由 {FYL.format(BASE + 3)} 線性收斂到 fade_g、淨利率收斂到 fade_m（{FYL.format(FADEY[-1])} 達到），非驅動因子推導"
for k, t in notes.items(): Y.cell(R[k], NOTEC, t)
Y.column_dimensions["V"].width = 50

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

# 估值規則區塊：十年倍數歷史與規則值（見頂 25 分位、中段中位數、谷底 75 分位；2027 A 為見頂與中段、2028 C 為谷底）
RULE = C["rule"]
def _y(s, n, k): return f"Y!{FC[(s, n)]}{R[k]}"
def _cases(basis_peak, basis_trough, lab_peak, lab_trough):
    return [{"case": "見頂", "pos": "見頂年 → 25 分位", "q": 25, "basis_label": lab_peak, "basis": basis_peak},
            {"case": "中段", "pos": "中段 → 中位數", "q": 50, "basis_label": lab_peak, "basis": basis_peak},
            {"case": "谷底", "pos": "谷底 → 75 分位", "q": 75, "basis_label": lab_trough, "basis": basis_trough}]
y1, y2 = FYL.format(YRS[0]), FYL.format(YRS[1])
vseries = [
    {"id": "P/B", "name": "P/B（主）", "sources": [dict(x) for x in RULE["pb"]["sources"]],
     "cases": _cases("=" + _y("A", 1, "bvps"), "=" + _y("C", 2, "bvps"), f"{y1} 年底每股淨值（A）", f"{y2} 年底每股淨值（C）")}]
pe_src = [dict(x) for x in RULE["pe"]["sources"]] if "pe" in RULE else None
nr = vr + 2
if "sec" in RULE:
    nr, SECW = VR.add_sec_table(V, nr, "5. 美光歷史財務（SEC 10-K XBRL；P/B 為 macrotrends 會計年度底）",
        ["會計年度", "P/B", "權益 $B", "庫藏股（成本）$B", "歸屬淨利 $B"], [(str(y), *v) for y, *v in RULE["sec"]],
        "P/E ＝ P/B × 權益 ÷ 淨利（虧損年不計）；買回加回 P/B ＝ P/B × 權益 ÷（權益＋庫藏股）。庫藏股為成本；FY2026 年底尚未公布，調整每股淨值用 FY2025 值（偏低）")
    sr = [SECW[str(y)] for y, *_ in RULE["sec"]]
    pe_src = [{"label": "SEC 10-K 換算", "tag": "Verified／Derived", "note": "虧損年不計",
               "vals": [f'=IF(E{r_}>0,B{r_}*C{r_}/E{r_},"")' for r_ in sr]}]
    pba_src = [{"label": "SEC 10-K 換算", "tag": "Verified／Derived",
                "vals": [f"=B{r_}*C{r_}/(C{r_}+D{r_})" for r_ in sr]}]
    t_last = f"$D${SECW['2025']}"
    adj27 = f"=(Y!{FC[('A', 1)]}{R['eq']}+{t_last}+Y!{FC[('A', 1)]}{R['bb']})/Y!{FC[('A', 1)]}{R['sh']}*1000"
    adj28 = f"=(Y!{FC[('C', 2)]}{R['eq']}+{t_last}+Y!{FC[('C', 1)]}{R['bb']}+Y!{FC[('C', 2)]}{R['bb']})/Y!{FC[('C', 2)]}{R['sh']}*1000"
vseries.append({"id": "P/E", "name": "P/E（對照）", "sources": pe_src,
    "cases": _cases("=" + _y("A", 1, "eps"), "=" + _y("C", 2, "eps"), f"{y1} EPS（A）", f"{y2} EPS（C）")})
if "sec" in RULE:
    vseries.append({"id": "買回加回 P/B", "name": "買回加回 P/B（對照）", "sources": pba_src,
        "cases": _cases(adj27, adj28, f"{y1} 年底調整每股淨值（A）", f"{y2} 年底調整每股淨值（C）")})
tr = VR.add_rule_block(V, nr, vseries, fmtp, "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底倍數依循環位置取分位。多個來源時各自取分位再平均、四捨五入到兩位。倍數歷史為輸入（藍字），來源與標記見各表；原放在 markets/overview/extract.py（2026-10-23 移入）。")

# 0. 目標價（v0.4 樣板；Andy 2026-10-25 11:28 定版：DCF 與 Comparable 加權，11:53 權重各半；列標籤依工作單逐字，記憶體 Comparable 用 P/B）
# 放在 Val 既有內容之後（不位移既有列；HTML 估值章第一個區塊讀這裡）；B／C／D 欄＝情境 A 順利／B 平平／C 不如預期
YRS6 = YRS + FADEY
def ycol(s, y): return FC[(s, YRS.index(y) + 1)] if y in YRS else FCF_[(s, y)]
FW = C["tp"]["fwd_w"]            # 12 個月後的未來 12 月 EPS 權重（第一年、第二年）
PB0 = nr + 5                    # 第 6 節 P/B 十年歷史第一個來源（B 欄）
PB10 = f"$B${PB0}:$B${PB0 + 9}"
assert V.cell(PB0, 1).value == "2016" and V.cell(PB0 + 9, 1).value == "2025", (V.cell(PB0, 1).value, V.cell(PB0 + 9, 1).value)
T = {}
def tset(r_, key, label, fs, fmt, note=""):
    V.cell(r_, 1, label); T[key] = r_
    for c_, f_ in zip("BCD", fs):
        if f_ is None: continue
        V[f"{c_}{r_}"] = f_; V[f"{c_}{r_}"].number_format = fmt
    if note: V[f"E{r_}"] = note
KEYF = PatternFill("solid", fgColor="FFF2CC")
r_ = tr
V.cell(r_, 1, "0. 目標價").font = BOLD; T["sec0"] = r_; r_ += 1
V.cell(r_, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["sec01"] = r_; r_ += 1
hdr(V, r_, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"]); T["h01"] = r_; r_ += 1
V.cell(r_, 1, "情境 EPS（FY2027–FY2032）"); V.cell(r_, 5, f"Y 頁各情境；{FYL.format(FADEY[0])} 起為收斂法（Y 頁灰底）"); T["eps_h"] = r_; r_ += 1
EPSR = {}
for y in YRS6:
    tset(r_, f"eps{y}", f"  {FYL.format(y)}", [f"=Y!{ycol(s, y)}{R['eps']}" for s in MKT.SCEN], fmtp); EPSR[y] = r_; r_ += 1
V.cell(r_, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = r_
V.cell(r_, 5, f"記憶體資本密集：{FYL.format(YRS[0])}–{FYL.format(YRS[-1])} 改用每股（淨利＋折舊－資本支出）＝ Y 頁自由現金流 ÷ 期末股數；收斂期（{FYL.format(FADEY[0])} 起）沒有資本支出與折舊的推導，用 EPS ×（1 − 再投資率）"); r_ += 1
tset(r_, "reinv", f"  再投資率（tp_g ÷ tp_roe，上限 0.8；{FYL.format(FADEY[0])} 起用）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）"); REINV = r_; r_ += 1
CFR = {}
for y in YRS6:
    fs = [(f"=Y!{ycol(s, y)}{R['fcf']}/Y!{ycol(s, y)}{R['sh']}*{SS}" if y in YRS else f"={c_}{EPSR[y]}*(1-$B${REINV})") for s, c_ in zip(MKT.SCEN, "BCD")]
    tset(r_, f"cf{y}", f"  {FYL.format(y)}" + ("（淨利＋折舊－資本支出，每股）" if y in YRS else "（EPS ×（1 − 再投資率））"), fs, fmtp); CFR[y] = r_; r_ += 1
c0, c9 = CFR[YRS6[0]], CFR[YRS6[-1]]
tset(r_, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c_}{c0}:{c_}{c9},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c_}{c9}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c_ in "BCD"], fmtp,
     f"Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝{FYL.format(FADEY[-1])} 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 {FYL.format(YRS[0])}" + C["tp"].get("dcf_note", "")); r_ += 1
tset(r_, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c_}{T['dcfv']}*(1+tp_r)" for c_ in "BCD"], fmtp); r_ += 1
tset(r_, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{T['dcf12']}+tp_wB*C{T['dcf12']}+tp_wC*D{T['dcf12']}", None, None], fmtp, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁"); r_ += 2
V.cell(r_, 1, "0.2 Comparable（市場主流同業前瞻 P/B）").font = BOLD; T["sec02"] = r_; r_ += 1
hdr(V, r_, ["同業與前瞻 P/B", "代號", "前瞻 P/B", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
p0 = r_
for nm_, tk_, pb_, dt_, src_ in C["tp"]["peers"]:
    V.cell(r_, 1, nm_); V.cell(r_, 2, tk_); c = V.cell(r_, 3, pb_); c.font = BLUE; c.number_format = "0.00"
    V.cell(r_, 4, dt_); V.cell(r_, 5, src_); V.cell(r_, 6, "Interested-party"); r_ += 1
p9 = r_ - 1
T["p0"], T["p9"] = p0, p9
tset(r_, "pmed", "同業前瞻 P/B 中位數", [f"=MEDIAN(C{p0}:C{p9})", None, None], "0.00", C["tp"]["peer_note"]); r_ += 1
y28 = YRS[1]
tset(r_, "fbv", "12 個月後的前瞻每股淨值（情境加權）", ["=" + "+".join(f"tp_w{s}*Y!{ycol(s, y28)}{R['bvps']}" for s in MKT.SCEN), None, None], fmtp,
     f"同業倍數＝現價 ÷ 下一個會計年度末共識每股淨值（約 14 個月後）；對應 12 個月後（約 2027-10）用 {FYL.format(y28)} 年底每股淨值（情境加權）"); r_ += 1
tset(r_, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{T['fbv']}", None, None], fmtp, "同業中位數 × 前瞻每股淨值"); r_ += 1
feps = "+".join(f"tp_w{s}*({FW[0]}*Y!{ycol(s, YRS[0])}{R['eps']}+{FW[1]}*Y!{ycol(s, YRS[1])}{R['eps']})" for s in MKT.SCEN)
tset(r_, "feps", "  對照：12 個月後的前瞻 EPS（情境加權）", ["=" + feps, None, None], fmtp, C["tp"]["fwd_note"]); r_ += 2
V.cell(r_, 1, "0.3 結論").font = BOLD; T["sec03"] = r_; r_ += 1
hdr(V, r_, ["項目", "數值", "", "", "說明"]); T["h03"] = r_; r_ += 1
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], fmtp, "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁"); r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], fmtp); r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", None, None], fmtp, C["tp"]["cons_tp"][3]); r_ += 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", "現價＝In 頁 px"); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "dcfpx", "DCF ÷ 現價", [f"={c_}{T['dcf12']}/px" for c_ in "BCD"], '0.00"x"', "各情境 DCF 價值（推到 12 個月後）÷ 現價（A／B／C）；任一情境高於 3 倍時，目標價方法加註"); r_ += 1
tset(r_, "gbv", "與市場共識目標價的差距：EPS 因素（記憶體以每股淨值計）", [f"=B{T['fbv']}/cons_bvps28-1", None, None], "0.0%",
     f"本模型 {FYL.format(y28)} 年底每股淨值（情境加權）÷ 共識 {FYL.format(y28)} 每股淨值 − 1（Comparable 用 P/B，故以每股淨值代 EPS）"); r_ += 1
tset(r_, "geps", "  對照：前瞻 EPS ÷ 共識前瞻 EPS − 1", [f"=B{T['feps']}/({FW[0]}*cons_eps27+{FW[1]}*cons_eps28)-1", None, None], "0.0%", "共識 EPS 見 In 頁與 Recon 頁"); r_ += 1
tset(r_, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_bvps28)/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{r_}"] = f'="共識目標價隱含 P/B "&TEXT(cons_tp/cons_bvps28,"0.00")&"x（÷ 共識 {FYL.format(y28)} 每股淨值），同業中位數 "&TEXT(B{T["pmed"]},"0.00")&"x"'; r_ += 1
tset(r_, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）"); r_ += 1
tset(r_, "pr", "歷史倍數檢查：現價隱含 P/B 的十年分位（A／B／C）", [f"=IF(px/Y!{ycol(s, y28)}{R['bvps']}>=MAX({PB10}),1,IF(px/Y!{ycol(s, y28)}{R['bvps']}<=MIN({PB10}),0,PERCENTRANK({PB10},px/Y!{ycol(s, y28)}{R['bvps']})))" for s in MKT.SCEN], "0%",
     f"現價 ÷ 各情境 {FYL.format(y28)} 年底每股淨值，在 2016–2025 年底 P/B（第 6 節第一個來源）分布的分位（高於十年最大值記 100%、低於最小值記 0%）；只作檢查，不進目標價（記憶體主估值為 P/B，故以 P/B 代 P/E）"); r_ += 1
V.cell(r_, 1, "目標價方法"); T["meth"] = r_
_m = C["tp"]["method"]
_lit = "&".join('"' + _m[i:i + 200] + '"' for i in range(0, len(_m), 200))      # 公式字串常數上限 255 字
V[f"B{r_}"] = (f'={_lit}&IF(MAX(B{T["dcfpx"]}:D{T["dcfpx"]})>3,"注意：DCF ÷ 現價有情境高於 3 倍（見上方「DCF ÷ 現價」列）；DCF 對收斂期淨利率與永續假設敏感，宜與 Comparable、市場共識並讀。","")'); r_ += 2
for k_ in ("dcftp", "comptp", "tpw"): V[f"B{T[k_]}"].fill = KEYF

# 舊法：規則值情境加權（參考；v0.4 起不再叫目標價）。情境對應（Andy 2026-10-25）：A 順利＝P/B 75 分位、B 平平＝中位數、C 不如預期＝25 分位，× 各情境 y1 年底每股淨值；C 一定最低
AD = {V.cell(i, 1).value: i for i in range(1, tr) if isinstance(V.cell(i, 1).value, str) and V.cell(i, 1).value.startswith("採用 ")}
TPMAP = (("A", f"情境 A（順利）：75 分位 × {y1} 年底每股淨值（A）", "75 分位"), ("B", f"情境 B（平平）：中位數 × {y1} 年底每股淨值（B）", "中位數"), ("C", f"情境 C（不如預期）：25 分位 × {y1} 年底每股淨值（C）", "25 分位"))
V.cell(r_, 1, "舊法：規則值情境加權（參考）").font = BOLD; T["old"] = r_
hdr(V, r_ + 1, ["項目", "數值", "", "", "說明"])
t0 = r_ + 2; T["old0"] = t0
V.cell(t0, 1, "規則值情境加權價（參考，非目標價）")
V.cell(t0, 2, f"=tp_wA*B{t0 + 2}+tp_wB*B{t0 + 3}+tp_wC*B{t0 + 4}").number_format = fmtp
V.cell(t0, 5, "權重 × 各情境隱含價；約為年底價格、未折現。v0.4 樣板起目標價改為第 0 節（DCF 與 Comparable 加權）")
V.cell(t0 + 1, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）")
V.cell(t0 + 1, 2, f"=B{t0}/px-1").number_format = "0.0%"
for k_, (s_, lab, q_) in enumerate(TPMAP):
    rr_ = t0 + 2 + k_; src = AD[f"採用 {q_}（P/B）"]
    V.cell(rr_, 1, lab); V.cell(rr_, 2, f"=B{src}*Y!{FC[(s_, 1)]}{R['bvps']}").number_format = fmtp
    V.cell(rr_, 5, f'="P/B 採用 {q_} "&TEXT(B{src},"0.00")&"x（第 6 節 B{src}）× Y 頁 {FC[(s_, 1)]}{R["bvps"]}"')
V.cell(t0 + 5, 1, "情境權重 A／B／C")
for j_, s_ in enumerate("ABC"): V.cell(t0 + 5, 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(t0 + 5, 5, f'="合計 "&TEXT(B{t0 + 5}+C{t0 + 5}+D{t0 + 5},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"')
V.cell(t0 + 6, 1, "舊法說明")
V.cell(t0 + 6, 2, (f"十年（2016–2025 年底）P/B 分位 × 各情境 {y1} 年底每股淨值：A 順利＝75 分位、B 平平＝中位數、C 不如預期＝25 分位（Andy 2026-10-25；C 一定最低）。"
                   "第 6 節的見頂／中段／谷底規則值不變（AI 投資地圖抽數用）。" + C.get("tp_bias", "")))
T["old_end"] = t0 + 6
# 樣板 v0.4.1：上檔空間級距、投資建議（Andy 2026-10-25 14:00、14:17）；放在 Val 現有最後一列之後（不位移既有列）
_u = f"B{T['upw']}"
T["grade"] = T["old_end"] + 2; T["rec"] = T["old_end"] + 3
tset(T["grade"], "grade", "上檔空間級距", [f'=IF({_u}>0.5,"大於 50%",IF({_u}>0.25,"25%–50%",IF({_u}>=0,"0%–25%",IF({_u}>=-0.25,"0%～−25%",IF({_u}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
tset(T["rec"], "rec", "投資建議", [f'=IF({_u}>0.5,"強力買進",IF({_u}>0.25,"買進",IF({_u}>=-0.25,"中立",IF({_u}>=-0.5,"放空","強力放空"))))', None, None], "@",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
V.column_dimensions["A"].width = 36

# ---------------- Sens（營收變動 × (1 − 變動營業費用率)；成本不隨價格變動，故為精確值）----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次只動一個價格或份額輸入；Δ營業利益＝Δ營收 × (1 − 營業費用變動率)（成本只隨量變動，所以這是精確值，不是近似）"
hdr(SE, 3, ["欄", "輸入變動", "Δ營收", "營業利益", "相對基準"])
sr = 4
for (s, n) in (("A", 1), ("B", 2), ("C", 2), ("C", 3)):
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
# 目標價參數（v0.4 樣板）：tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.3／0.7 對加權目標價的影響（Comparable 不隨 r、g 變動；收斂法的 fade_g 不隨 tp_g 變動）
TPS0 = sr + 2
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；Y 頁收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def dcf_expr(c_, r="tp_r", g="tp_g"):
    reinv = f"MIN({g}/tp_roe,0.8)"
    cf = lambda y: (f"Val!{c_}{CFR[y]}" if y in YRS else f"Val!{c_}{EPSR[y]}*(1-{reinv})")
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"({pv}+{cf(YRS6[-1])}*(1+{g})/({r}-{g})/(1+{r})^6)"
TPCASES = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
for k_, (nm_, st_, kw) in enumerate(TPCASES):
    i_ = TPS0 + 2 + k_
    r_, g_, w_ = kw.get("r", "tp_r"), kw.get("g", "tp_g"), kw.get("w", "tp_w_dcf")
    SE.cell(i_, 1, nm_); SE.cell(i_, 2, st_)
    SE[f"C{i_}"] = "=(" + "+".join(f"tp_w{s}*{dcf_expr(c_, r_, g_)}" for s, c_ in zip(MKT.SCEN, "BCD")) + f")*(1+{r_})"
    SE[f"D{i_}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i_}"] = f"={w_}*C{i_}+(1-{w_})*D{i_}"
    SE[f"F{i_}"] = f"=E{i_}/E${TPS0 + 2}-1"
    for c_ in "CDE": SE[f"{c_}{i_}"].number_format = fmtp
    SE[f"F{i_}"].number_format = "0.0%"
TPS_LAST = TPS0 + 1 + len(TPCASES)
SE.cell(TPS_LAST + 1, 1, "核對：基準列 DCF 目標價應等於 Val 頁第 0 節"); SE[f"C{TPS_LAST + 1}"] = f"=C{TPS0 + 2}-Val!B{T['dcftp']}"; SE[f"C{TPS_LAST + 1}"].number_format = "0.000000"
SE.column_dimensions["A"].width = 14; SE.column_dimensions["B"].width = 30

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = f"Recon — 本模型（正向推導）vs 共識或公司指引（只作事後對照，不用來回推參數）；{U}"
hdr(RC, 3, ["項目", "本模型", "對照值", "差距", "對照來源"])
RECON = C["recon"] + [(f"{FYL.format(YRS[0])}E EPS（A）vs 共識（第 0 節差距拆解用）", "=Y!" + FC[("A", 1)] + "{eps}", "=cons_eps27", C["tp"]["cons_src"]),
                      (f"{FYL.format(YRS[1])}E EPS（A）vs 共識（第 0 節差距拆解用）", "=Y!" + FC[("A", 2)] + "{eps}", "=cons_eps28", C["tp"]["cons_src"]),
                      (f"{FYL.format(YRS[1])}E 年底每股淨值（A）vs 共識", "=Y!" + FC[("A", 2)] + "{bvps}", "=cons_bvps28", C["tp"]["cons_src"])]
for i, (nm, f, v, src) in enumerate(RECON):
    rr = 4 + i
    RC.cell(rr, 1, nm); RC.cell(rr, 2, f.format(**R)).number_format = "0.0"
    RC.cell(rr, 3, v).font = BLUE
    RC.cell(rr, 4, f"=B{rr}/C{rr}-1").number_format = "0%"; RC.cell(rr, 5, src)
RC.column_dimensions["A"].width = 46; RC.column_dimensions["E"].width = 50

# ---------------- Feedback、Sources、README ----------------
FB = wb.create_sheet("Feedback"); FB["A1"] = "Feedback — 回饋算力收支"; FB["A1"].font = BOLD
fb = [("每顆加速器 HBM 金額（GB300，2026 價）", "=nv_hbm", "$K"),
      ("全市場每顆加速器平均 HBM 2026／2027", f"=TEXT(Mkt!D{MK['gbavg']},\"0\")&\" ／ \"&TEXT(Mkt!E{MK['gbavg']},\"0\")", "GB"),
      ("HBM 市場 2026／2027（A）", f"=TEXT(Mkt!D{MK['mA']},\"0\")&\" ／ \"&TEXT(Mkt!E{MK['mA']},\"0\")", "$B"),
      ("2027 HBM 需求缺口（三情境相同）", f"=Mkt!E{MK['gap']}", "EB"),
      (f"{C['name']} HBM 毛利率", "無法由公開資料分離", "—")]
for i, row in enumerate(fb, 3):
    for j, v in enumerate(row, 1): FB.cell(i, j, v)
FB.column_dimensions["A"].width = 40; FB.column_dimensions["B"].width = 20
SO = wb.create_sheet("Sources"); SO["A1"] = "來源"; SO["A1"].font = BOLD
for i, (a, b) in enumerate(C["sources"] + [("markets/memory/market.py（共用市場層）", "本 repo"), ("markets/chip_bridge v0.2", "本 repo")], 2):
    SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 60; SO.column_dimensions["B"].width = 120
RD = wb.create_sheet("README", 0)
lines = [(f"{C['name']} 收支與估值模型 {C['version']}（記憶體樣板 v0.4.1）　{C['date']}", True), ("模型命題：" + C["thesis"], False), ("", False),
 ("驅動因子 → 推導量", True),
 ("Path（共用市場層）：一般 DRAM、NAND 合約價季增 → 季度指數（情境 A／B／C）；各公司依會計年度取平均，算年均價變動", False),
 ("Mkt（共用市場層）：加速器顆數 × 每顆 GB × 情境顆數係數（2028–29，與台積電、Advantest 同一套）→ 需求位元；MIN(需求, 供給（三情境相同）) → 出貨；× 單價（依情境）→ HBM 市場（日曆年）", False),
 ("Q：最後實績季 × 公司指引 × TrendForce 季價 → 基準年剩餘季度", False),
 ("Y：前一年 × (1＋位元) × 價格指數比 → 一般 DRAM、NAND；Σ 各季 HBM 市場 × 市占 → HBM；成本量 × (1 − 年降)＋折舊 → 營業利益 → 淨利 → 權益 → 每股淨值", False),
 ("Val：第 0 節目標價＝DCF（股東現金流，FY2027–FY2032 明確期＋永續）與 Comparable（同業前瞻 P/B × 前瞻每股淨值）加權；另有每股淨值 × P/B（主）、EPS × P/E（對照）與十年倍數規則值", False), ("Sens：各收入線同幅變動對營業利益的影響", False), ("非記憶體事業（只有三星）：營收與營業利益為輸入，不建模", False), ("Recon：與共識、公司指引事後對照", False), ("", False),
 (f"情境：A＝{MKT.SCEN_DESC['A']}；B＝{MKT.SCEN_DESC['B']}；C＝{MKT.SCEN_DESC['C']}（目標價對應：A＝順利、B＝平平、C＝不如預期；{FYL.format(FADEY[0])}–{FYL.format(FADEY[-1])} 為收斂法）", False),
 (f"幣別單位：{U}；會計年度：{C['fy_desc']}", False), ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False), ("", False)] + [(t, False) for t in C["readme_extra"]]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
# 樣板 v0.4.1：投資建議（公式，跟著 Val 數字變）與重點觀察三點（各公司 company.py 的 obs）；接在 README 現有最後一列之後
assert len(C["obs"]) >= 3, "company.py 須有 obs（重點觀察至少三點）"
_rr = len(lines) + 1
RD.cell(_rr, 1, f'="投資建議："&Val!B{T["rec"]}&"（依上檔空間固定規則）"')
for i, t in enumerate(C["obs"], 1): RD.cell(_rr + i, 1, f"重點觀察 {i}：{t}")
RD.column_dimensions["A"].width = 170
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, os.path.basename(os.path.dirname(os.path.abspath(sys.argv[1]))))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, C.get("in_layout_version", C["version"]), "2026-10-30")  # In 頁版面（個股模型原則第 8 節；數值不變）
wb.save(sys.argv[2])
print("saved", sys.argv[2])
