# R4 對帳：Wattonomics 側殘差歸因（Compare 1–4 點＋「2026 七家改甲2」＋NVIDIA 營收對照）。
# 用法：python3 wattonomics/tools/r4_attrib.py <輸出 xlsx>
#   讀 wattonomics/model/CURRENT 的重算值（data_only）與 Siliconomics 側 R4 橋接活頁簿（r1），
#   寫成「輸入格（藍字）＋公式格（黑字）」的活頁簿；之後以 LibreOffice 重算。本工具不改任何模型。
#   公式一律在 Excel；Python 只搬值與排版。
import os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WT = os.path.join(ROOT, "wattonomics", "model")
wt_f = open(os.path.join(WT, "CURRENT")).read().strip()
W = openpyxl.load_workbook(os.path.join(WT, wt_f), data_only=True)
SIL_R4 = "siliconomics/docs/reports/r4_bridge/20261031_R4對帳橋接_Siliconomics側_r1.xlsx"
S = openpyxl.load_workbook(os.path.join(ROOT, SIL_R4), data_only=True)

YRS = [2024, 2025, 2026, 2027, 2028, 2029]   # 計算欄
OUT = [2026, 2027, 2028, 2029]               # 輸出年
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()

# ---------------------------------------------------------------- In
IN = wb.active; IN.title = "In"
IN["A1"] = (f"In — R4 Wattonomics 側歸因的輸入。藍字＝輸入：取自 {wt_f}（Wattonomics v0.1.7）與 {os.path.basename(SIL_R4)}"
            "（Siliconomics 側 r1）；低／高為敏感度。標記沿用專案規則。")
IN["A1"].font = BOLD
for j, h in enumerate(["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源"], 1):
    c = IN.cell(3, j, h); c.font = BOLD; c.fill = HDR
names = {}
r = 4


def inp(key, label, unit, lo, base, hi, tag, src):
    global r
    for j, v in enumerate([key, label, unit, lo, base, hi, tag, src], 1):
        c = IN.cell(r, j, v)
        if j in (4, 5, 6) and v is not None:
            c.font = BLUE
    names[key] = f"In!$E${r}"
    if lo is not None:
        names[key + "_lo"] = f"In!$D${r}"
    if hi is not None:
        names[key + "_hi"] = f"In!$F${r}"
    r += 1


def sec(t):
    global r
    IN.cell(r, 1, t).font = BOLD; r += 1


def ws_row(ws, key, col=2):
    for row in ws.iter_rows():
        if row[col - 1].value == key:
            return row[0].row
    raise KeyError(key)


def v(ws, row, col):
    x = ws.cell(row, col).value
    return x if isinstance(x, (int, float)) else None


WCOL = {2023: 4, 2024: 5, 2025: 6, 2026: 7, 2027: 8, 2028: 9, 2029: 10}   # Wattonomics D..J
INYC = {2024: 9, 2025: 10, 2026: 11, 2027: 12, 2028: 13, 2029: 14, 2030: 15}  # In Y2／Y9 欄 I..O＝財年

# A. Siliconomics 側（r1 Bridge 頁）
sec("A. Siliconomics 側 R4 橋接 r1（情境 A；Bridge 頁值）")
SB = S["Bridge"]; SI = S["In"]
SBC = {2025: 4, 2026: 5, 2027: 6, 2028: 7, 2029: 8}
srow = {k: ws_row(SB, k, 1) for k in ["b1", "b2", "b3", "b4", "b5", "b6", "b7m", "b8", "b10"]}
labels = {"b1": "Siliconomics 部署 GW（成本等值；出貨年）", "b2": "① 成本等值 → 實體", "b3": "實體出貨 GW",
          "b4": "② 出貨年 → 上線年", "b5": "③ 未上線庫存增加", "b6": "④ 汰換（毛 → 淨）",
          "b7m": "（揭露）Siliconomics 中國 H200 成本等值 GW", "b8": "Siliconomics 推得淨新增上線 IT GW", "b10": "⑥ 殘差（r1）"}
for k in srow:
    for y in ([2025] if k == "b3" else []) + OUT:
        val = v(SB, srow[k], SBC[y])
        inp(f"s_{k}_{y}", f"{labels[k]} {y}", "GW", None, val, None, "Derived", f"{os.path.basename(SIL_R4)} Bridge!{CL(SBC[y])}{srow[k]}")
for key in ["lag", "inv_k", "retire"]:
    rr = ws_row(SI, key, 1)
    inp(key, SI.cell(rr, 2).value, SI.cell(rr, 3).value, v(SI, rr, 4), v(SI, rr, 5), v(SI, rr, 6), "Assumed",
        f"{os.path.basename(SIL_R4)} In!D{rr}:F{rr}（與 Siliconomics 側同參數，對稱套用）")
rr = ws_row(SI, "cap_oth_2026", 1)
inp("s_cap_oth_2026", "Siliconomics 其他 ASIC（MTIA、Maia、昇騰等）AI IT capex 2026", "$B", None, v(SI, rr, 5), None, "Derived",
    f"{os.path.basename(SIL_R4)} In!E{rr}")
for key in ["pgw_vr", "hold_oth", "nv_dc_25", "nv_dc_26", "nvc_gb_tk", "nvc_gb_jh"]:
    rr = ws_row(SI, key, 1)
    inp(key, SI.cell(rr, 2).value, SI.cell(rr, 3).value, v(SI, rr, 4), v(SI, rr, 5), v(SI, rr, 6), SI.cell(rr, 7).value,
        f"{os.path.basename(SIL_R4)} In!E{rr}；原出處：{SI.cell(rr, 8).value}")

# B. Wattonomics 市場層
sec(f"B. Wattonomics 市場層（{wt_f}；需求 A、D0）")
DP = W["Deploy"]; PL = W["Players"]; CG = W["CapexGW"]; WIN = W["In"]
for y in [2025] + OUT:
    inp(f"wt_new_{y}", f"Wattonomics 年新增 IT GW（主路徑）{y}", "GW", None, v(DP, 19, WCOL[y]), None, "Derived", f"{wt_f} Deploy!{CL(WCOL[y])}19")
for y in [2025] + OUT:
    inp(f"wt_ye_{y}", f"Wattonomics 年底上線 IT GW（主路徑）{y}", "GW", None, v(DP, 13, WCOL[y]), None, "Derived", f"{wt_f} Deploy!{CL(WCOL[y])}13")
for y in OUT:
    inp(f"nnv_mkt_{y}", f"市場新增 GW 非 NVIDIA 占比 {y}", "%", None, v(PL, 12, WCOL[y]), None, "Derived", f"{wt_f} Players!{CL(WCOL[y])}12")
cnr = ws_row(PL, "cn_add"); othr = ws_row(PL, "oth_tk")
for y in OUT:
    inp(f"cn_add_{y}", f"Wattonomics 中國年新增 IT GW {y}", "GW", None, v(PL, cnr, WCOL[y]), None, "Derived", f"{wt_f} Players!{CL(WCOL[y])}{cnr}")
inp("oth_ye_2025", "Wattonomics「其他（主權、企業、學研）」2025 年底 IT GW（之後持平）", "GW", None, v(PL, othr, 6), None, "Derived",
    f"{wt_f} Players!F{othr}（2026 起無園區列，年新增＝0）")
rr = ws_row(WIN, "tolband", 1)
inp("tolband", "對帳容許帶（±）", "%", v(WIN, rr, 8), v(WIN, rr, 9), v(WIN, rr, 10), "Assumed", f"{wt_f} In!H{rr}:J{rr}")
rr = ws_row(WIN, "lag_bldg", 1)
inp("wb1", "廠房提前入帳：前一年權重 w1（lag_bldg 18 個月→0.5；12→1；24→0）", "x", 1, 0.5, 0, "Assumed",
    f"{wt_f} CapexGW!E4:E6（lag_bldg 低／基準／高；In!H{rr}:J{rr}）")
inp("wb2", "廠房提前入帳：前兩年權重 w2（lag_bldg 18 個月→0.5；12→0；24→1）", "x", 0, 0.5, 1, "Assumed",
    f"{wt_f} CapexGW!F4:F6（當年權重 w0＝0）")
rr = ws_row(WIN, "ai_sh", 1)
inp("ai_sh", "資本支出 AI 占比（未覆寫者）", "%", v(WIN, rr, 8), v(WIN, rr, 9), v(WIN, rr, 10), "Assumed", f"{wt_f} In!H{rr}:J{rr}")
rr = ws_row(WIN, "ai_amzn", 1)
inp("ai_amzn", "Amazon 資本支出 AI 占比覆寫", "%", v(WIN, rr, 8), v(WIN, rr, 9), v(WIN, rr, 10), "Derived", f"{wt_f} In!H{rr}:J{rr}")

# C. 七家（CapexGW 頁與 In Y9）
P7 = [("msft", 6, 10), ("googl", 12, 29), ("amzn", 12, 48), ("meta", 12, 67), ("orcl", 5, 86), ("crwv", 12, 105), ("nbis", 12, 124)]
PNAME = {"msft": "Microsoft", "googl": "Alphabet", "amzn": "Amazon", "meta": "Meta", "orcl": "Oracle", "crwv": "CoreWeave",
         "nbis": "Nebius", "xai": "SpaceX／xAI", "oai": "OpenAI／Stargate 自建", "ant": "Anthropic 自建"}
sec("C. 七家：每 IT GW 成本、站點年新增、In Y2 曆年 AI 資本支出（CapexGW 頁）、In Y9 個股模型 v0.3 資本支出（財年）")
for p, fye, base in P7:
    inp(f"fye_{p}", f"{PNAME[p]} 財年結束月", "月", None, fye, None, "Verified", f"{wt_f} In fye_{p}")
    for y in YRS:
        inp(f"pit_{p}_{y}", f"{PNAME[p]} 每 IT GW IT capex {y}", "$B/GW", None, v(CG, base + 7, WCOL[y]), None, "Derived",
            f"{wt_f} CapexGW!{CL(WCOL[y])}{base + 7}")
        inp(f"pfac_{p}_{y}", f"{PNAME[p]} 每 IT GW 廠房 capex {y}", "$B/GW", None, v(CG, base + 8, WCOL[y]), None, "Derived",
            f"{wt_f} CapexGW!{CL(WCOL[y])}{base + 8}")
        inp(f"ai2_{p}_{y}", f"{PNAME[p]} 曆年 AI 資本支出（In Y2 讀法）{y}", "$B", None, v(CG, base + 6, WCOL[y]), None, "Interested-party",
            f"{wt_f} CapexGW!{CL(WCOL[y])}{base + 6}（空白＝In Y2 無值）")
        inp(f"site_{p}_{y}", f"{PNAME[p]} 甲（站點）年新增 IT GW {y}", "GW", None, v(CG, base + 11, WCOL[y]), None, "Derived",
            f"{wt_f} CapexGW!{CL(WCOL[y])}{base + 11}")
        inp(f"main_{p}_{y}", f"{PNAME[p]} 主路徑年新增 IT GW {y}", "GW", None, v(CG, base + 12, WCOL[y]), None, "Derived",
            f"{wt_f} CapexGW!{CL(WCOL[y])}{base + 12}")
    gr = ws_row(WIN, f"gcap_{p}", 1)
    for fy in range(2024, 2031):
        inp(f"gcap_{p}_{fy}", f"{PNAME[p]} 資本支出合計（個股模型 v0.3 基準）FY{fy}", "$B", None, v(WIN, gr, INYC[fy]), None, "Derived",
            f"{wt_f} In!{CL(INYC[fy])}{gr}（In Y9；{str(WIN.cell(gr, 7).value)[:60]}…）")
    if p == "msft":
        cr = ws_row(WIN, "cap_msft", 1)
        inp("cap2_msft_2025", "Microsoft 資本支出 FY2025（In Y2；Y9 無 FY2025，曆年換算補用）", "$B", None, v(WIN, cr, 10), None, "Verified",
            f"{wt_f} In!J{cr}（10-K 現金 PP&E 增添，不含融資租賃；與 Y9 口徑不同，標 Derived 混合）")
    if p == "orcl":
        cr = ws_row(WIN, "cap_orcl", 1)
        inp("cap2_orcl_2025", "Oracle 資本支出 FY2025（In Y2；Y9 無 FY2025，曆年換算補用）", "$B", None, v(WIN, cr, 10), None, "Verified",
            f"{wt_f} In!J{cr}")

# D. 非七家（11 家中有個股模型資本支出者）
sec("D. 非七家：xAI、OpenAI 自建、Anthropic 自建（In Y9 AI 資本支出；Players 站點年新增）；每 IT GW 用 NVIDIA 世代加權")
for y in YRS:
    inp(f"pitN_{y}", f"NVIDIA 每 IT GW IT capex（世代加權）{y}", "$B/GW", None, v(CG, 7, WCOL[y]), None, "Derived", f"{wt_f} CapexGW!{CL(WCOL[y])}7")
    inp(f"pfacN_{y}", f"NVIDIA 每 IT GW 廠房 capex {y}", "$B/GW", None, v(CG, 8, WCOL[y]), None, "Derived", f"{wt_f} CapexGW!{CL(WCOL[y])}8")
for p in ["xai", "oai", "ant"]:
    ar = ws_row(PL, f"{p}_add")
    for y in YRS:
        inp(f"site_{p}_{y}", f"{PNAME[p]} 甲（站點）年新增 IT GW {y}", "GW", None, v(PL, ar, WCOL[y]), None, "Derived", f"{wt_f} Players!{CL(WCOL[y])}{ar}")
    gr = ws_row(WIN, f"gcapai_{p}", 1)
    for fy in range(2024, 2030):
        inp(f"gai_{p}_{fy}", f"{PNAME[p]} AI 資本支出（個股模型）FY{fy}", "$B", None, v(WIN, gr, INYC[fy]), None, "Derived",
            f"{wt_f} In!{CL(INYC[fy])}{gr}（In Y9）")

for j, w in enumerate([18, 62, 8, 9, 10, 9, 16, 90], 1):
    IN.column_dimensions[CL(j)].width = w
for k, ref in names.items():
    wb.defined_names[k] = DefinedName(k, attr_text=ref)


def n(key):
    """名稱存在才用；否則回傳空字串常數（讓 IF 判斷為缺值）。"""
    return key if key in names else '""'


# ---------------------------------------------------------------- Calc（每個敏感度情況一塊）
CA = wb.create_sheet("Calc")
CA["A1"] = ("Calc — 甲2 重算（與 Wattonomics CapexGW 同式）：當年資金＝IT 占比 × AI_t ＋ 廠房占比 ×（w1·AI_t−1 ＋ w2·AI_t−2）；"
            "CapexGW 在 t−1、t−2 缺值時留空，本頁改為：缺 t−2 用 t−1、兩者皆缺用當年（只影響 2024–2026 與無前期值的玩家，屬近似）；GW_付款年＝資金 ÷（IT＋廠房）；資金缺值時以站點年新增代入（同 CapexGW「缺值以甲代入」）。"
            "上線年＝(1−lag)·GW_t ＋ lag·GW_t−1 − inv_k·(GW_t − GW_t−1)（與 Siliconomics ②③ 同參數）。")
CA["A1"].font = BOLD
CC = {y: 4 + i for i, y in enumerate(YRS)}  # D..I
cr = 3
SUM = {}  # (case, kind, y) -> cell


def C(y):
    return CL(CC[y])


def put(label, key, fmap, bold=False):
    global cr
    CA.cell(cr, 1, key); CA.cell(cr, 2, label)
    for y in YRS:
        f = fmap(y)
        if f is not None:
            c = CA.cell(cr, CC[y], f); c.number_format = "0.000"
            if bold:
                c.font = BOLD
    row = cr; cr += 1
    return row


CASES = [("base", "基準", {}),
         ("ailo", "AI 占比 低", dict(ai="_lo")), ("aihi", "AI 占比 高", dict(ai="_hi")),
         ("lblo", "lag_bldg 低（12 月）", dict(wb="_lo")), ("lbhi", "lag_bldg 高（24 月）", dict(wb="_hi")),
         ("lglo", "上線延遲 低（3 個月）", dict(lag="lag_lo")), ("lghi", "上線延遲 高（6 個月）", dict(lag="lag_hi")),
         ("ivlo", "庫存 低", dict(inv="inv_k_lo")), ("ivhi", "庫存 高", dict(inv="inv_k_hi"))]


def cy_y9(p, y, ai):
    """In Y9 財年 → 曆年 AI 資本支出（公式字串）。"""
    share = f"ai_amzn{ai}" if p == "amzn" else f"ai_sh{ai}"
    if p in ("msft", "orcl"):
        m = 6 if p == "msft" else 5
        a = f"gcap_{p}_{y}" if f"gcap_{p}_{y}" in names and IN[names[f'gcap_{p}_{y}'].split('!')[1].replace('$', '')].value is not None else None
        if a is None and f"cap2_{p}_{y}" in names:
            a = f"cap2_{p}_{y}"
        b = f"gcap_{p}_{y+1}"
        if a is None or IN[names[b].split('!')[1].replace('$', '')].value is None:
            return None
        return f"({m}/12*{a}+{12-m}/12*{b})*{share}"
    k = f"gcap_{p}_{y}"
    if IN[names[k].split('!')[1].replace('$', '')].value is None:
        return None
    return f"{k}*{share}"


for ck, cl, kw in CASES:
    ai = kw.get("ai", ""); wb_ = kw.get("wb", ""); lag = kw.get("lag", "lag"); inv = kw.get("inv", "inv_k")
    CA.cell(cr, 1, f"■ 情況：{cl}").font = BOLD; cr += 1
    for j, h in enumerate(["鍵", "項目"] + [str(y) for y in YRS], 1):
        c = CA.cell(cr, j, h); c.font = BOLD; c.fill = HDR
    cr += 1
    rows = {"x2": [], "on2": [], "x9": [], "on9": [], "site": [], "xn": [], "onn": [], "siten": []}
    for p, fye, base in P7:
        for src in ("2", "9"):
            if src == "2":
                if ai == "":
                    aif = lambda y, p=p: f"=IF(ISNUMBER(ai2_{p}_{y}),ai2_{p}_{y},\"\")"
                else:
                    share = "ai_amzn" if p == "amzn" else "ai_sh"
                    aif = lambda y, p=p, share=share: f"=IF(ISNUMBER(ai2_{p}_{y}),ai2_{p}_{y}/{share}*{share}{ai},\"\")"
            else:
                aif = lambda y, p=p: (lambda s: f"={s}" if s else '=""')(cy_y9(p, y, ai))
            ra = put(f"{PNAME[p]} 曆年 AI 資本支出（{'In Y2' if src == '2' else 'Y9 個股模型 v0.3'}）", f"{ck}_ai{src}_{p}", aif)
            def gwf(y, p=p, ra=ra):
                cur = f"{C(y)}{ra}"
                p1 = f"{CL(CC[y]-1)}{ra}" if y > YRS[0] else '""'
                p2 = f"{CL(CC[y]-2)}{ra}" if y > YRS[1] else '""'
                pit, pf = f"pit_{p}_{y}", f"pfac_{p}_{y}"
                fac = f"IF(AND(ISNUMBER({p1}),ISNUMBER({p2})),wb1{wb_}*{p1}+wb2{wb_}*{p2},IF(ISNUMBER({p1}),{p1},{cur}))"
                fund = f"({pit}/({pit}+{pf})*{cur}+{pf}/({pit}+{pf})*{fac})"
                return f"=IF(ISNUMBER({cur}),{fund}/({pit}+{pf}),site_{p}_{y})"
            rx = put(f"{PNAME[p]} 甲2 年新增 IT GW（付款年；缺值以站點代入）", f"{ck}_x{src}_{p}", gwf)
            def onf(y, rx=rx):
                if y == YRS[0]:
                    return None
                cur = f"{C(y)}{rx}"; prv = f"{CL(CC[y]-1)}{rx}"
                return f"=(1-{lag})*{cur}+{lag}*{prv}-{inv}*({cur}-{prv})"
            ro = put(f"{PNAME[p]} 甲2 年新增 IT GW（上線年）", f"{ck}_on{src}_{p}", onf)
            rows["x" + src].append(rx); rows["on" + src].append(ro)
        rows["site"].append(put(f"{PNAME[p]} 甲（站點）年新增", f"{ck}_site_{p}", lambda y, p=p: f"=site_{p}_{y}"))
    for p in ["xai", "oai", "ant"]:
        ra = put(f"{PNAME[p]} AI 資本支出（個股模型，曆年＝財年）", f"{ck}_ain_{p}",
                 lambda y, p=p: f"=IF(ISNUMBER(gai_{p}_{y}),gai_{p}_{y},\"\")")
        def gwn(y, p=p, ra=ra):
            cur = f"{C(y)}{ra}"
            p1 = f"{CL(CC[y]-1)}{ra}" if y > YRS[0] else '""'
            p2 = f"{CL(CC[y]-2)}{ra}" if y > YRS[1] else '""'
            pit, pf = f"pitN_{y}", f"pfacN_{y}"
            fac = f"IF(AND(ISNUMBER({p1}),ISNUMBER({p2})),wb1{wb_}*{p1}+wb2{wb_}*{p2},IF(ISNUMBER({p1}),{p1},{cur}))"
            fund = f"({pit}/({pit}+{pf})*{cur}+{pf}/({pit}+{pf})*{fac})"
            return f"=IF(ISNUMBER({cur}),{fund}/({pit}+{pf}),site_{p}_{y})"
        rx = put(f"{PNAME[p]} 模型資本支出推得年新增 IT GW（付款年；缺值以站點代入）", f"{ck}_xn_{p}", gwn)
        def onn(y, rx=rx):
            if y == YRS[0]:
                return None
            cur = f"{C(y)}{rx}"; prv = f"{CL(CC[y]-1)}{rx}"
            return f"=IF(ISNUMBER(gai_{p}_{y-1}),(1-{lag})*{cur}+{lag}*{prv}-{inv}*({cur}-{prv}),{cur})"
        ro = put(f"{PNAME[p]} 模型資本支出推得年新增 IT GW（上線年；前一年無模型值則不平移）", f"{ck}_onn_{p}", onn)
        rows["xn"].append(rx); rows["onn"].append(ro)
        rows["siten"].append(put(f"{PNAME[p]} 甲（站點）年新增", f"{ck}_siten_{p}", lambda y, p=p: f"=site_{p}_{y}"))
    for kind, lab in [("x2", "七家 甲2 付款年（In Y2）"), ("on2", "七家 甲2 上線年（In Y2）"), ("x9", "七家 甲2 付款年（Y9）"),
                      ("on9", "七家 甲2 上線年（Y9）"), ("site", "七家 甲（站點）"), ("xn", "非七家 模型推得 付款年"),
                      ("onn", "非七家 模型推得 上線年"), ("siten", "非七家 甲（站點）")]:
        rr_ = put(f"合計：{lab}", f"{ck}_S_{kind}", lambda y, kind=kind: "=" + "+".join(f"{C(y)}{x}" for x in rows[kind]), True)
        for y in YRS:
            SUM[(ck, kind, y)] = f"Calc!${C(y)}${rr_}"
    cr += 1
CA.column_dimensions["A"].width = 20; CA.column_dimensions["B"].width = 58
for y in YRS:
    CA.column_dimensions[C(y)].width = 10

# ---------------------------------------------------------------- Bridge
B = wb.create_sheet("Bridge")
B["A1"] = ("Bridge — Wattonomics 側歸因：Siliconomics 推得上線 GW（r1）vs Wattonomics 主路徑（v0.1.7，A、D0），逐項列 Wattonomics 側可量化的差距。"
           "正值＝Wattonomics 應加的 GW。")
B["A1"].font = BOLD
for j, h in enumerate(["列", "項目", "說明"] + [str(y) for y in OUT] + ["標記"], 1):
    c = B.cell(3, j, h); c.font = BOLD; c.fill = HDR
BC = {y: 4 + i for i, y in enumerate(OUT)}  # D..G
br = 4
R = {}


def BCL(y):
    return CL(BC[y])


def brow(key, label, note, fmap, tag="", bold=False, fmt="0.00"):
    global br
    B.cell(br, 1, key); B.cell(br, 2, label); B.cell(br, 3, note); B.cell(br, 8, tag)
    for y in OUT:
        f = fmap(y)
        if f is not None:
            c = B.cell(br, BC[y], f); c.number_format = fmt
            if bold:
                c.font = BOLD
    R[key] = br; br += 1


def S_(kind, y, ck="base"):
    return SUM[(ck, kind, y)]


def ref(k, y):
    return f"{BCL(y)}{R[k]}"


B.cell(br, 1, "一、起點").font = BOLD; br += 1
brow("s8", "Siliconomics 推得淨新增上線 IT GW（r1）", "Siliconomics Bridge b8（已扣 ②③④）", lambda y: f"=s_b8_{y}", "Derived", True)
brow("w0", "Wattonomics 主路徑年新增 IT GW（v0.1.7，A、D0）", "Deploy 第 19 列", lambda y: f"=wt_new_{y}", "Derived", True)
brow("r0", "殘差 r1（Siliconomics 側交來）", "s8 − w0", lambda y: f"={ref('s8', y)}-{ref('w0', y)}", "", True)
br += 1
B.cell(br, 1, "二、Wattonomics 側逐項（Compare 第 1–4 點＋覆蓋）").font = BOLD; br += 1
brow("d1", "【覆蓋】2026 七家改甲2（付款年，In Y2 讀法）＝AI半導體 10-31 觀察 2 的量",
     "Σ七家 甲2（付款年）− Σ七家 站點；2027 起主路徑已是甲2 → 0",
     lambda y: f"={S_('x2', y)}-{S_('site', y)}" if y == 2026 else "=0", "Derived")
brow("d2", "【第 1 點 出貨年 vs 上線年】七家甲2 付款年 → 上線年（Wattonomics 自己的口徑修正）",
     "甲2 的 IT 資本支出以付款年當上線年；以 Siliconomics ②③ 同參數平移：Σ上線年 − Σ付款年（In Y2）",
     lambda y: f"={S_('on2', y)}-{S_('x2', y)}", "Derived／Assumed")
brow("d3", "【覆蓋】七家資本支出改讀個股模型 v0.3（v0.1.8 規劃）",
     "Σ上線年（Y9）− Σ上線年（In Y2）", lambda y: f"={S_('on9', y)}-{S_('on2', y)}", "Derived")
brow("d4", "【覆蓋】非七家：xAI、OpenAI 自建、Anthropic 自建改讀個股模型資本支出",
     "Σ模型推得（上線年）− Σ站點；每 GW 用 NVIDIA 世代加權成本（OpenAI 含客製晶片 → 偏低估 GW）",
     lambda y: f"={S_('onn', y)}-{S_('siten', y)}", "Derived")
brow("d5", "【覆蓋】「其他（企業、學研）」2026 起持平 → 改隨市場成長",
     "其他_t−1 ×（主路徑年底 t ÷ 年底 t−1 − 1），逐年累積", lambda y: None, "Analogy")
# d5 公式需逐年累積：先算其他年底
for y in OUT:
    prev = "oth_ye_2025" if y == 2026 else f"(oth_ye_2025+SUM($D${R['d5']}:{CL(BC[y]-1)}{R['d5']}))"
    c = B.cell(R["d5"], BC[y], f"={prev}*(wt_ye_{y}/wt_ye_{y-1}-1)"); c.number_format = "0.00"
brow("d6", "【第 3 點 成本等值 vs 實體】Wattonomics 側", "Wattonomics 甲2 已用 NonNV 持有比換算實體 GW（與 Siliconomics ① 同法），無需調整",
     lambda y: "=0", "—")
brow("d7", "【第 4 點 中國】Wattonomics 中國年新增（揭露，不調整）", "Rystad 容量 × AI 占比；見下方「中國」節",
     lambda y: f"=cn_add_{y}", "Derived")
brow("wadj", "Wattonomics 調整後年新增 IT GW", "w0 + d1 + d2 + d3 + d4 + d5（d7 不加）",
     lambda y: "=" + "+".join(ref(k, y) for k in ["w0", "d1", "d2", "d3", "d4", "d5"]), "Derived", True)
br += 1
B.cell(br, 1, "三、比較口徑對齊（兩邊都不改模型）").font = BOLD; br += 1
brow("e1", "【第 1 點附帶】汰換：Wattonomics 不建模汰換（主路徑實為毛額）→ 加回 Siliconomics ④",
     "−④；兩邊改以毛額比較", lambda y: f"=-s_b6_{y}", "Assumed")
brow("sg", "Siliconomics 推得上線 IT GW（毛額，與 Wattonomics 同口徑）", "s8 + e1", lambda y: f"={ref('s8', y)}+{ref('e1', y)}", "Derived", True)
br += 1
B.cell(br, 1, "四、殘差與判讀").font = BOLD; br += 1
brow("rf", "殘差（調整後）", "sg − wadj", lambda y: f"={ref('sg', y)}-{ref('wadj', y)}", "", True)
brow("ra", "比值：Wattonomics 調整後 ÷ Siliconomics（毛額）", "", lambda y: f"={ref('wadj', y)}/{ref('sg', y)}", "", True, "0.000")
brow("rb", "比值：Wattonomics v0.1.7 ÷ Siliconomics（毛額）", "", lambda y: f"={ref('w0', y)}/{ref('sg', y)}", "", False, "0.000")
brow("lb", "容許帶判定（調整後）", "|1 − 比值| ≤ tolband → 帶內",
     lambda y: f"=IF(ABS(1-{ref('ra', y)})<=tolband,\"帶內\",IF({ref('ra', y)}<1,\"帶外（Wattonomics 偏低）\",\"帶外（Wattonomics 偏高）\"))", "", True, "General")
brow("km", "七家上線 GW 需放大倍數（其餘不變，使殘差歸零）", "(sg −（wadj − 七家上線年 Y9））÷ 七家上線年 Y9",
     lambda y: f"=({ref('sg', y)}-({ref('wadj', y)}-{S_('on9', y)}))/{S_('on9', y)}", "Derived", False, "0.00")
br += 1
B.cell(br, 1, "五、敏感度（一次動一個輸入；殘差＝sg − wadj）").font = BOLD; br += 1
for j, h in enumerate(["情況", "", ""] + [str(y) for y in OUT], 1):
    c = B.cell(br, j, h); c.font = BOLD; c.fill = HDR
br += 1
for ck, cl, kw in CASES:
    lagn = kw.get("lag", "lag"); invn = kw.get("inv", "inv_k")
    B.cell(br, 1, ck); B.cell(br, 2, cl)
    B.cell(br, 3, "Siliconomics 側同參數時，s8 亦隨之變（以 r1 的 ②③ 比例近似重算）" if ck[:2] in ("lg", "iv") else "")
    for y in OUT:
        # Siliconomics s8 隨 lag／inv 變：s8' = s8 − ②③ + (−(lag'+inv')·Δ出貨)
        if ck[:2] in ("lg", "iv"):
            dship = f"(s_b3_{y}-{'s_b3_' + str(y-1)})"
            sg = f"(s_b8_{y}-s_b4_{y}-s_b5_{y}-({lagn}+{invn})*{dship}-s_b6_{y})"
        else:
            sg = f"{ref('sg', y)}"
        d1 = f"({S_('x2', y, ck)}-{S_('site', y, ck)})" if y == 2026 else "0"
        wadj = (f"(wt_new_{y}+{d1}+({S_('on2', y, ck)}-{S_('x2', y, ck)})+({S_('on9', y, ck)}-{S_('on2', y, ck)})"
                f"+({S_('onn', y, ck)}-{S_('siten', y, ck)})+{ref('d5', y)})")
        c = B.cell(br, BC[y], f"={sg}-{wadj}"); c.number_format = "0.00"
    br += 1
for lab, fm in [("只做 d1（不平移、不改讀 v0.3、不補非七家）", lambda y: f"={ref('sg', y)}-({ref('w0', y)}+{ref('d1', y)})"),
                ("d1＋d2（改甲2 並平移；仍讀 In Y2）", lambda y: f"={ref('sg', y)}-({ref('w0', y)}+{ref('d1', y)}+{ref('d2', y)})"),
                ("不含 d4、d5（只動七家）", lambda y: f"={ref('sg', y)}-({ref('w0', y)}+{ref('d1', y)}+{ref('d2', y)}+{ref('d3', y)})")]:
    B.cell(br, 2, lab)
    for y in OUT:
        c = B.cell(br, BC[y], fm(y)); c.number_format = "0.00"
    br += 1
br += 1
B.cell(br, 1, "六、NVIDIA 營收對照（2026；只作佐證）").font = BOLD; br += 1


def one(key, label, note, f, tag="", fmt="0.00"):
    global br
    B.cell(br, 1, key); B.cell(br, 2, label); B.cell(br, 3, note); B.cell(br, 8, tag)
    c = B.cell(br, BC[2026], f); c.number_format = fmt
    R[key] = br; br += 1


one("n1", "NVIDIA 出貨 GW 2026（TK 每 GW 內容 20.8）", "nv_dc_26 ÷ nvc_gb_tk", "=nv_dc_26/nvc_gb_tk", "Derived")
one("n1j", "NVIDIA 出貨 GW 2026（黃仁勳 GB 說法 25）", "nv_dc_26 ÷ nvc_gb_jh", "=nv_dc_26/nvc_gb_jh", "Interested-party")
one("n0", "NVIDIA 出貨 GW 2025（以 GB300 內容換算；2025 多 GB200／Hopper、每 GW 內容較低 → 偏低估）", "nv_dc_25 ÷ nvc_gb_tk",
    "=nv_dc_25/nvc_gb_tk", "Derived")
one("n2", "NVIDIA 推得上線 GW 2026（同 ②③ 參數）", "(1−lag)·n1 + lag·n0 − inv_k·(n1 − n0)",
    f"=(1-lag)*D{R['n1']}+lag*D{R['n0']}-inv_k*(D{R['n1']}-D{R['n0']})", "Derived")
one("n3", "Wattonomics v0.1.7 2026 NVIDIA 上線 GW", "w0 ×（1 − nnv_mkt）", f"=D{R['w0']}*(1-nnv_mkt_2026)", "Derived")
one("n4", "Wattonomics 調整後 2026 NVIDIA 上線 GW（以市場 NonNV 占比近似）", "wadj ×（1 − nnv_mkt）", f"=D{R['wadj']}*(1-nnv_mkt_2026)", "Derived")
one("n5", "比值 v0.1.7 ÷ NVIDIA 推得", "n3 ÷ n2", f"=D{R['n3']}/D{R['n2']}", "", "0.000")
one("n6", "比值 調整後 ÷ NVIDIA 推得", "n4 ÷ n2", f"=D{R['n4']}/D{R['n2']}", "", "0.000")
br += 1
B.cell(br, 1, "七、中國（第 4 點；2026）").font = BOLD; br += 1
one("c1", "Wattonomics 中國年新增 IT GW 2026", "Players cn_add", "=cn_add_2026", "Derived")
one("c2", "Siliconomics 可辨識中國上限：H200 成本等值 ＋ 其他 ASIC 全部視為昇騰的實體 GW", "s_b7m_2026 + s_cap_oth_2026 ÷（pgw_vr × hold_oth）",
    "=s_b7m_2026+s_cap_oth_2026/(pgw_vr*hold_oth)", "Derived（上限）")
one("c3", "差（Wattonomics 中國 − Siliconomics 可辨識上限）", "正值＝不含中國比較時殘差反而更大", f"=D{R['c1']}-D{R['c2']}", "Derived")
B.column_dimensions["A"].width = 6; B.column_dimensions["B"].width = 66; B.column_dimensions["C"].width = 62
for y in OUT:
    B.column_dimensions[BCL(y)].width = 11
B.column_dimensions["H"].width = 16

out = sys.argv[1]
wb.save(out)
print("wrote", out)
