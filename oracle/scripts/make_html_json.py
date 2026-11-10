#!/usr/bin/env python3
"""產生 oracle/html.json（tools/company_html 的設定檔）：列號取自 build_oracle.py 的列號對照（ORCL_ROWMAP），
In 頁列號取自成品 xlsx 的鍵欄。build_oracle.py 改版後重跑本檔，再跑 tools/company_html/build_company_html.py oracle。

用法：ORCL_ROWMAP=rowmap.json python3 oracle/build_oracle.py raw.xlsx ；python3 oracle/scripts/make_html_json.py rowmap.json
"""
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = json.load(open(sys.argv[1], encoding="utf-8"))
X = sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
INR = {wb["In"].cell(r, 1).value: r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
SCN_H = 1 + next(r for r in range(4, wb["In"].max_row + 1) if "情境驅動對照" in str(wb["In"].cell(r, 1).value or ""))
TR, PR, CR, VAL = R["TR"], R["PR"], R["CR"], R["VAL"]
DC, SO, CON = VAL["DC"], VAL["SO"], VAL["CON"]
TB, PB, CB = TR["B"], PR["B"], CR["B"]
G = R["GATE"]
FU = R["FUND"]["Fund"]; FS = FU["SUM"]; FR = FU["R"]; RPB = R["FUND"]["RPO"]["R"]["B"]; DB = R["FUND"]["Debt"]; LSB = R["FUND"]["Lease"]["R"]["B"]; ST_ = R["FUND"]["Sites"]
RVN = max(v for k, v in R["FUND"]["RevDCF"]["R"].items() if not k.startswith("_"))


def drv(key, fmt, out_lab, out_ref):
    r = INR[key]
    return [f"@In!B{r}", [f"@In!E{r}|{fmt}", " ", f"@In!C{r}"], f"@In!G{r}|tag", [out_lab, out_ref]]


def met(name, key, fmt, chart=False, src="PL", cols="D{r}:I{r}", rows=None):
    rows = rows or {s: f"@{src}!" + cols.format(r=PR[s][key]) for s in "ABC"}
    m = {"name": name, "fmt": fmt, "rows": rows}
    if chart: m["chart"] = True
    return m



# ─── v0.5.1：HTML 版面 v1.0（docs/plan/個股HTML版面規格.md；範例 spacex/gen_html_json.py） ───
RDW = wb["README"]
def rfind(text):
    return next(r for r in range(1, RDW.max_row + 1) if str(RDW.cell(r, 1).value or "").startswith(text))
rd_show, rd_ans = rfind("展示層（HTML"), rfind("命題回覆（HTML")
rd_fail, rd_risk, rd_lim = rfind("失效條件（可觀察"), rfind("風險登錄"), rfind("已知限制")
rd_drv = rfind("驅動 → 推導對照")
IW = wb["In"]
# In 參數依業務分塊：In 頁節字母 → 業務、小標（小標不得含數字）；A–J 節（含 I2）每個鍵須剛好歸入一組
SEC_SUB = {"B": "已連網 MW 路徑與延誤", "C": "世代組合與汰換", "D": "每 MW 經濟性（Tokenomics 快照）與定價", "E": "機房租賃（未起租承諾）",
           "F": "OCI 其他（非 AI）營收路徑", "G": "其他營收線的長期年增", "H": "分部利潤率、未分攤費用與折舊",
           "I": "現金流與融資瀑布", "I2": "已發生的股債增資", "A": "股價與股數", "J": "折現率、目標價規則與同業倍數"}
BIZ = {"ai": ["B", "C", "D", "E"], "other": ["F", "G", "H"], "co": ["A", "J", "I", "I2"]}   # 全公司參數（股價股數、折現率、目標價規則、稅與融資）放最後一塊
SECKEYS, _sec = {}, None
for r in range(4, IW.max_row + 1):
    a_, b_ = IW.cell(r, 1).value, IW.cell(r, 2).value
    if a_ and b_ in (None, ""):
        _sec = str(a_).split(".")[0].strip(); continue
    if a_ and _sec in SEC_SUB: SECKEYS.setdefault(_sec, []).append(a_)
_got = [k for secs in BIZ.values() for x in secs for k in SECKEYS.get(x, [])]
assert set(SECKEYS) == set(SEC_SUB), ("In 頁節別與分塊設定不符", sorted(SECKEYS), sorted(SEC_SUB))
assert len(_got) == len(set(_got)) == sum(len(v) for v in SECKEYS.values()), "In 參數歸組不完整或重複"
has_rng = lambda k: IW.cell(INR[k], 4).value not in (None, "") or IW.cell(INR[k], 6).value not in (None, "")
IN_HEAD = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"]
def in_fmt(k):
    v, u = IW.cell(INR[k], 5).value, str(IW.cell(INR[k], 3).value or "")
    if hasattr(v, "year"): return "d"
    if u.startswith("%") or u == "pt": return "p1"
    if u in ("年", "月", "季", "次", "MW", "MW-IT") or u.startswith("MW-IT"): return "n0"
    if isinstance(v, (int, float)) and abs(v) >= 100: return "n1"
    return "n3"
def in_row(k):
    f = in_fmt(k); r = INR[k]
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]
def in_tables(biz, label):
    rng_rows, fix_rows = [], []
    for x in BIZ[biz]:
        a_ = [in_row(k) for k in SECKEYS[x] if has_rng(k)]; b_ = [in_row(k) for k in SECKEYS[x] if not has_rng(k)]
        if a_: rng_rows += [{"sub": SEC_SUB[x]}] + a_
        if b_: fix_rows += [{"sub": SEC_SUB[x]}] + b_
    out = []
    if rng_rows: out.append({"type": "table", "caption": f"{label}：有敏感度區間的參數（基準進目標價；低／高只用於敏感度頁）", "head": IN_HEAD, "rows": rng_rows})
    if fix_rows: out.append({"type": "table", "collapse": True, "summary": f"展開：{label}的其他固定輸入（實績、公司揭露、換算值）", "head": IN_HEAD, "rows": fix_rows})
    return out
# Val：主要項目與摺疊
VW = wb["Val"]
def vtab(keys, src, cols, fmts, rf=None):
    rf = rf or {}
    return [[f"@Val!{c}{src[k]}|{rf.get(k, {}).get(c, f)}" if f else f"@Val!{c}{src[k]}" for c, f in zip(cols, fmts)] for k in keys]
CON_MAIN = ["tp", "dcf", "comp", "px", "up", "upc", "rec", "cons"]
CON_REST = [k for k in sorted(CON, key=lambda k: CON[k]) if k not in CON_MAIN]
CON_RF = {k: {"C": "sp1"} for k in ("up", "upc", "gcons")} | {k: {"C": ""} for k in ("rec", "tier", "meth", "wchk")} | {"tvsh": {"C": "p1"}, "aiw": {"C": "p1"}}
DC_MAIN = ["ev", "tvsh", "eqv", "ps", "v12"]
DC_REST = ["2027", "2028", "2029", "2030", "2031", "2032", "tvb", "tv"]
DC_RF = {"tvsh": {c: "p1" for c in "CDE"}, "ps": {c: "n2" for c in "CDE"}, "v12": {c: "n2" for c in "CDE"}}
SO_MAIN = ["ev_oci", "ev_g", "ev_l", "ev", "v12"]
SO_REST = [k for k in sorted(SO, key=lambda k: SO[k]) if k not in SO_MAIN and k != "tps"]
SO_RF = {"sh": {c: "n3" for c in "CDE"}, "ps": {c: "n2" for c in "CDE"}, "v12": {c: "n2" for c in "CDE"}, "v12m": {c: "n2" for c in "CDE"},
         "oci_sh": {c: "p1" for c in "CDE"}, "g_sh": {c: "p1" for c in "CDE"}, "mg": {c: "p1" for c in "CDE"}, "ml": {c: "p1" for c in "CDE"}}
def vrow_lab(text):
    return next(r for r in range(1, VW.max_row + 1) if str(VW.cell(r, 1).value or "").startswith(text))
V_KE, V_KD, V_E, V_D, V_W, V_RON = (vrow_lab(t) for t in ("股權成本 ke", "稅後債務成本", "股權市值", "負債（借款", "WACC", "永續期新投資報酬率"))
# Recon：每節主要要點，其餘摺疊
RW = wb["Recon"]
rlast = max(r for r in range(1, RW.max_row + 1) if RW.cell(r, 1).value)
RSEC = [r for r in range(4, rlast + 1) if RW.cell(r, 1).value and RW.cell(r, 2).value is None and RW.cell(r, 3).value is None and str(RW.cell(r, 1).value)[:2] in ("A.", "B.", "B2", "C.", "D.", "E.")]
def rrow(lab, a, b):
    return next(r for r in range(a, b + 1) if str(RW.cell(r, 1).value or "").startswith(lab))
REC_MAIN_LAB = {"A.": ["合併營收 FY2029", "調整後 EBITDA FY2029", "GAAP EPS FY2029", "自由現金流 FY2029", "FY2027 營收 vs 公司指引"],
                "B.": ["Cloud Infrastructure（OCI）FY2027", "Cloud Infrastructure（OCI）FY2029"],
                "B2": ["FY2029 自由現金流", "FY2029 淨借款", "FY2030 合併營收"],
                "C.": ["FY2029 OCI AI 營業利益率", "FY2029 AI 增量 ROIC"], "D.": ["情境目標價 vs v0.3.1", "v0.3.1 情境目標價"], "E.": ["支持現價所需", "評等由放空翻為中立"]}
REC_HEAD = ["@Recon!A3", "@Recon!B3", "@Recon!C3", "@Recon!D3", "@Recon!E3", "@Recon!F3", "@Recon!G3", "@Recon!H3", "@Recon!I3"]
REC_COLS, REC_FMT = ["A", "B", "C", "D", "E", "F", "G", "H", "I"], ["", "n1", "n1", "sp1", "n1", "sp1", "n1", "sp1", ""]
def rec_fmt(r):
    lab = str(RW.cell(r, 1).value or ""); f = dict(zip(REC_COLS, REC_FMT))
    if any(k in lab for k in ("EPS", "利潤率", "ROIC", "殘差", "m_sw")):
        for c in "BCEG": f[c] = "n2" if "EPS" in lab else "p1"
    if any(k in lab for k in ("金額計", "殘差", "m_sw", "ROIC", "毛利率")):
        for c in "DFH": f[c] = "n1" if "金額計" in lab else "sp1"
    if "p_long" in lab:
        f.update({"B": "n2", "C": "n2", "D": "sp1"})
    for c in "BCDEFGH":
        v = RW[f"{c}{r}"].value
        if isinstance(v, str) and not v.replace(".", "").replace("-", "").isdigit(): f[c] = ""
    return f
def rec_rows(rows):
    return [[f"@Recon!{c}{r}|{rec_fmt(r)[c]}" if rec_fmt(r)[c] else f"@Recon!{c}{r}" for c in REC_COLS] for r in rows]
rec_blocks = []
for i, sr in enumerate(RSEC):
    end = (RSEC[i + 1] - 1) if i + 1 < len(RSEC) else rlast
    allr = [r for r in range(sr + 1, end + 1) if RW.cell(r, 1).value]
    key = str(RW.cell(sr, 1).value)[:2]
    main = [rrow(l, sr + 1, end) for l in REC_MAIN_LAB[key]]
    rest = [r for r in allr if r not in main]
    if main: rec_blocks.append({"type": "table", "caption": f"@Recon!A{sr}", "head": REC_HEAD, "rows": rec_rows(main)})
    if rest: rec_blocks.append({"type": "table", "collapse": True, "summary": "展開：" + ("本節全部項目" if not main else "本節其他項目"),
                                "caption": f"@Recon!A{sr}" if not main else None, "head": REC_HEAD, "rows": rec_rows(rest)})
for blk in rec_blocks:
    if blk.get("caption") is None: blk.pop("caption", None)
# Sens：龍捲風圖只顯示 |變動| ≥ 10%，其餘摺疊
SW_ = wb["Sens"]
def srows(col, big):
    out = []
    for r in range(R["SENS"][0], R["SENS"][1] + 1):
        v = SW_[f"{col}{r}"].value
        v = float(v) if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace(".", "").replace("-", "").isdigit()) else None
        if v is None: continue
        if (abs(v) >= 0.1) == big and (big or abs(v) > 0): out.append(str(r))
    return ",".join(out)
def tor(col, title):
    b, sm = srows(col, True), srows(col, False)
    out = [{"type": "tornado", "title": title + "：變動達兩位數百分比的案例", "sheet": "Sens", "rows": b, "label_col": "A", "value_col": col, "fmt": "sp1"}] if b else []
    if sm: out.append({"type": "tornado", "collapse": True, "summary": "展開：" + title + "，變動未達兩位數百分比的案例", "title": title, "sheet": "Sens", "rows": sm, "label_col": "A", "value_col": col, "fmt": "sp1"})
    return out
FD = R["FUND"]

cfg = {
    "company": "Oracle",
    "ticker": "ORCL",
    "wide_text": True,
    "hero": {"rec": f"@Val!C{CON['rec']}",
             "tiles": [[f"@Val!A{CON['tp']}", f"@Val!C{CON['tp']}|n2", ["上檔空間 ", f"@Val!C{CON['up']}|sp1"]],
                       ["Comparable 目標價（分部加總混合版）", f"@Val!C{CON['comp']}|n2", ["上檔空間 ", f"@Val!C{CON['upc']}|sp1"]],
                       [f"@Val!A{CON['px']}", f"@Val!C{CON['px']}|n2", ""]],
             "note": [f"@README!A{rd_show + 8}", "。投資建議依加權目標價判定：", f"@Val!F{CON['rec']}"]},
    "thesis": f"@README!A{rd_show + 1}",
    "question": "FY2029 的 OCI 營收與營業利益率，相對於資本支出與租約的擴張，以及資金缺口（自由現金流為負）要如何補",
    "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
    "sections": [
        {"id": "thesis", "title": "命題與核心問題", "blocks": [
            {"type": "h3", "text": "命題"},
            {"type": "list", "ordered": True, "items": [f"@README!A{rd_show + k}" for k in (2, 3, 4)]},
            {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）",
             "items": [f"@README!A{rd_ans + k}" for k in (1, 2, 3)]},
            {"type": "list", "collapse": True, "summary": "展開：主要失效條件（出現任一可觀察證據，結論須重估）",
             "items": [f"@README!A{rd_fail + k}" for k in (1, 2, 3)]},
            {"type": "h3", "text": "重點觀察"},
            {"type": "list", "ordered": True, "items": [f"@README!A{rd_show + k}" for k in (5, 6, 7)]},
            {"type": "table", "collapse": True, "summary": "展開：命題答案摘要（三情境）", "caption": f"@Thesis!A{R['SUM'][0]-1}",
             "head": [f"@Thesis!A{R['SUM'][0]}", f"@Thesis!B{R['SUM'][0]}", f"@Thesis!C{R['SUM'][0]}", f"@Thesis!D{R['SUM'][0]}", f"@Thesis!E{R['SUM'][0]}"],
             "rowgen": {"sheet": "Thesis", "rows": f"{R['SUM'][0]+1}-{R['SUM'][1]}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n1", "n1", "n1"],
                        "row_fmt": {str(R['SUM'][0] + k): {"C": "p1", "D": "p1", "E": "p1"} for k in (3, 4, 7)}}}]},
        {"id": "funding", "title": "資金缺口與融資（收支模組）", "intro": "@Fund!A2", "blocks": [
            {"type": "table", "caption": "三情境摘要", "head": ["@Fund!A3", "@Fund!C3", "@Fund!D3", "@Fund!E3"],
             "rowgen": {"sheet": "Fund", "rows": f"{FS['trough']}-{FS['chk']}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(FS["lev"]): {"C": "n2", "D": "n2", "E": "n2"}, str(FS["tyear"]): {"C": "", "D": "", "E": ""}}}},
            {"type": "pl", "years": f"@Fund!D{FU['hdr']}:I{FU['hdr']}", "metrics": [
                met("融資前累積現金：不做任何新融資（十億美元）", None, "n1", chart=True, rows={s: f"@Fund!D{FR[s]['cumx']}:I{FR[s]['cumx']}" for s in "ABC"}),
                met("融資前現金流：營運來源 − 用途（十億美元）", None, "n1", chart=True, rows={s: f"@Fund!D{FR[s]['flow']}:I{FR[s]['flow']}" for s in "ABC"}),
                met("外部融資需求（十億美元）", None, "n1", rows={s: f"@Fund!D{FR[s]['need']}:I{FR[s]['need']}" for s in "ABC"})]},
            {"type": "table", "caption": "資金缺口壓力表：瀑布另需外部融資（情境 B；預先算好的組合）", "head": [f"@Fund!A{FU['STRESS']['wf'][0] - 1}", f"@Fund!D{FU['STRESS']['wf'][0] - 1}", f"@Fund!E{FU['STRESS']['wf'][0] - 1}", f"@Fund!F{FU['STRESS']['wf'][0] - 1}"],
             "rowgen": {"sheet": "Fund", "rows": f"{FU['STRESS']['wf'][0]}-{FU['STRESS']['wf'][1]}", "cols": ["A", "D", "E", "F"], "fmts": ["", "n1", "n1", "n1"]}},
        ] + [{"type": "table", "collapse": True, "summary": sm_, "head": [f"@Fund!A{a - 1}", f"@Fund!D{a - 1}", f"@Fund!E{a - 1}", f"@Fund!F{a - 1}"],
              "rowgen": {"sheet": "Fund", "rows": f"{a}-{b}", "cols": ["A", "D", "E", "F"], "fmts": ["", f_, f_, f_]}}
             for (a, b), f_, sm_ in ((FU["STRESS"]["eqhy"], "n1", "展開：壓力表：其中新股＋高息債"), (FU["STRESS"]["trough"], "n1", "展開：壓力表：融資前累積現金谷底"),
                                     (FU["STRESS"]["lev"], "n2", "展開：壓力表：調整後槓桿峰值"), (FU["STRESS"]["tp"], "n1", "展開：壓力表：加權目標價"))] + [
            {"type": "table", "collapse": True, "summary": "展開：來源、用途、缺口與融資瀑布（情境 B，逐年）", "head": [f"@Fund!A{FU['hdr']}"] + [f"@Fund!{c}{FU['hdr']}" for c in "DEFGHI"],
             "rowgen": {"sheet": "Fund", "rows": f"{FR['B']['L']}-{FR['B']['debt']}", "cols": ["A", "D", "E", "F", "G", "H", "I"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
                        "row_fmt": {str(FR["B"][k_]): {c: "n2" for c in "DEFGHI"} for k_ in ("L", "lev")}}},
            {"type": "table", "collapse": True, "summary": "展開：合約排程對模型交付（RPO；情境 B）", "head": ["@RPO!A4"] + [f"@RPO!{c}4" for c in "DEFGH"],
             "rowgen": {"sheet": "RPO", "rows": f"6,{RPB['rev']}-{RPB['dbl']}", "cols": ["A", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1"],
                        "row_fmt": {str(RPB[k_]): {c: "n2" for c in "DEFGH"} for k_ in ("cov", "dbl")}}},
            {"type": "table", "collapse": True, "summary": "展開：債務到期表與再融資備忘", "head": [f"@Debt!A{DB['hdr']}"] + [f"@Debt!{c}{DB['hdr']}" for c in "DEFGHIJ"],
             "rowgen": {"sheet": "Debt", "rows": f"{DB['L']['old']}-{DB['L']['cum']}", "cols": ["A", "D", "E", "F", "G", "H", "I", "J"], "fmts": ["", "n3", "n3", "n3", "n3", "n3", "n3", "n3"],
                        "row_fmt": {str(DB["L"]["cpn"]): {c: "p2" for c in "DEFGHI"}}}},
            {"type": "table", "collapse": True, "summary": "展開：債務逐檔與新票據", "head": ["@Debt!A4", "@Debt!C4", "@Debt!D4", "@Debt!E4", "@Debt!F4"],
             "rowgen": {"sheet": "Debt", "rows": f"{DB['r0']}-{DB['tot']},{DB['n0']}-{DB['n1'] + 1}", "cols": ["A", "C", "D", "E", "F"], "fmts": ["", "d", "p3", "n3", "n3"]}},
            {"type": "table", "collapse": True, "summary": "展開：未起租租約的租金路徑（情境 B）", "head": ["@Lease!A14"] + [f"@Lease!{c}14" for c in "DEFGHI"],
             "rowgen": {"sheet": "Lease", "rows": f"{LSB['rent']}-{LSB['permw']}", "cols": ["A", "D", "E", "F", "G", "H", "I"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2"],
                        "row_fmt": {str(LSB["ratio"]): {"D": "p1"}}}},
            {"type": "table", "collapse": True, "summary": "展開：具名站點與容量上限", "head": ["@Sites!A4", "@Sites!B4", "@Sites!C4", "@Sites!D4", "@Sites!E4", "@Sites!F4"],
             "rowgen": {"sheet": "Sites", "rows": f"{ST_['s0']}-{ST_['R']['res']},{ST_['R']['capc']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n0", "n0", "", ""]}}]},
        {"id": "drivers", "title": "驅動因子 → 推導量", "blocks": [
            {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": [f"@In!{c}{SCN_H}" for c in "BDEFCG"],
             "rows": [[f"@In!B{INR[k]}", f"@In!D{INR[k]}|{f}", f"@In!E{INR[k]}|{f}", f"@In!F{INR[k]}|{f}", f"@In!C{INR[k]}", f"@In!G{INR[k]}|tag"]
                      for k, f in (("scn_pace", "n0"), ("scn_cap", "n0"), ("scn_delay", "n0"), ("scn_w", "p0"))]},
            {"type": "h3", "text": "一、OCI AI 算力（命題模組）"},
            {"type": "note", "parts": ["@Thesis!A2"]}]
            + in_tables("ai", "OCI AI 算力")
            + [{"type": "table", "caption": "OCI AI 算力：已連網 MW → 營收 → 費用、折舊 → 利潤與報酬（情境 B）",
             "head": ["@Thesis!A3", "@Thesis!B3", "@Thesis!C3", "@Thesis!D3", "@Thesis!E3", "@Thesis!F3", "@Thesis!G3", "@Thesis!H3", "@Thesis!I3", "@Thesis!J3"],
             "rowgen": {"sheet": "Thesis", "rows": f"{TB['conn']}-{TB['bill']},{TB['anc']}-{TB['rent']},{TB['cx']},{TB['dep']}-{TB['om']},{TB['oth']},{TB['oci']},{TB['roic']}",
                        "cols": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"], "fmts": ["", "", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1"],
                        "row_fmt": {**{str(TB[k]): {c: "n0" for c in "CDEFGHIJ"} for k in ("conn", "add", "avg", "lag", "bill")},
                                    **{str(TB[k]): {c: "n2" for c in "CDEFGHIJ"} for k in ("anc", "pm", "oxm")},
                                    **{str(TB[k]): {c: "p1" for c in "CDEFGHIJ"} for k in ("om", "roic")}}}},
            {"type": "table", "collapse": True, "summary": "展開：RPO 轉換檢查（命題的對照列）", "caption": f"@Thesis!A{R['RPO'][0]-2}",
             "head": [f"@Thesis!A{R['RPO'][0]-1}", f"@Thesis!B{R['RPO'][0]-1}", f"@Thesis!C{R['RPO'][0]-1}", f"@Thesis!D{R['RPO'][0]-1}", f"@Thesis!E{R['RPO'][0]-1}"],
             "rowgen": {"sheet": "Thesis", "rows": f"{R['RPO'][0]}-{R['RPO'][1]}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n2", "n2", "n2"],
                        "row_fmt": {str(R['RPO'][0]): {"C": "n1", "D": "n1", "E": "n1"}}}},
            {"type": "h3", "text": "二、OCI 其他、SaaS、授權、支援、硬體、服務"}]
            + in_tables("other", "其他營收線與利潤率")
            + [{"type": "table", "caption": "校準：雲端與軟體分部中「OCI AI 以外」的利潤率（以最近一季推得）", "head": ["項目", "數值"],
                "rows": [[f"@Thesis!A{R['CAL'] + k}", f"@Thesis!C{R['CAL'] + k}|{f}"] for k, f in ((0, "n2"), (1, "n2"), (2, "p1"), (3, "p1"), (4, "p1"))]},
               {"type": "h3", "text": "三、全公司：股價、股數、折現率、目標價規則、稅與融資"}]
            + in_tables("co", "全公司")},
        {"id": "pl", "title": "三情境損益（分部 → 合併）", "blocks": [
            {"type": "pl", "years": "@PL!D3:I3", "metrics": [
                met("合併營收（十億美元）", "rev", "n1", chart=True),
                met("Cloud Infrastructure（十億美元）", "oci", "n1", chart=True),
                met("其中 OCI AI 算力（命題子列）", "ai", "n1"),
                met("OCI AI 營業利益率", "ociom", "p1", chart=True),
                met("營業利益率（GAAP）", "oim", "p1"),
                met("EPS（GAAP 稀釋，美元）", "eps", "n2", chart=True)]},
            {"type": "table", "collapse": True, "summary": "展開：情境 B 的營收線、分部與合併損益", "head": ["@PL!A3", "@PL!C3", "@PL!D3", "@PL!E3", "@PL!F3", "@PL!G3", "@PL!H3", "@PL!I3"],
             "rowgen": {"sheet": "PL", "rows": f"{PB['oci']}-{PB['rev']}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2"]}},
            {"type": "table", "caption": "分部核對（三情境合計絕對差，應為零）", "head": ["@PL!A3", "@PL!C3", "@PL!D3", "@PL!E3", "@PL!F3", "@PL!G3", "@PL!H3", "@PL!I3"],
             "rowgen": {"sheet": "PL", "rows": f"{R['CHK'][0]}-{R['CHK'][2]}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I"], "fmts": ["", "n3", "n3", "n3", "n3", "n3", "n3", "n3"]}},
            {"type": "h3", "text": "現金流與收支關卡"},
            {"type": "table", "head": ["@CF!A3", "@CF!C3", "@CF!D3", "@CF!E3"],
             "rowgen": {"sheet": "CF", "rows": f"{G}-{G+6},{G+8}-{G+10}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(G + 4): {"C": "n2", "D": "n2", "E": "n2"}}}},
            {"type": "pl", "years": f"@CF!D{R['CFH']}:I{R['CFH']}", "metrics": [
                met("自由現金流（十億美元）", None, "n1", chart=True, rows={s: f"@CF!D{CR[s]['fcf']}:I{CR[s]['fcf']}" for s in "ABC"}),
                met("自由現金流扣股利（十億美元）", None, "n1", rows={s: f"@CF!D{CR[s]['fcfd']}:I{CR[s]['fcfd']}" for s in "ABC"}),
                met("資本支出（十億美元）", None, "n1", chart=True, rows={s: f"@CF!D{CR[s]['cx']}:I{CR[s]['cx']}" for s in "ABC"}),
                met("調整後槓桿（倍）", None, "n2", rows={s: f"@CF!D{CR[s]['lev']}:I{CR[s]['lev']}" for s in "ABC"})]},
            {"type": "note", "parts": ["@CF!A1"]}]},
        {"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": [
            {"type": "bars", "title": "目標價與現價", "fmt": "n1", "items": [
                {"label": f"@Val!A{CON['dcf']}", "value": f"@Val!C{CON['dcf']}", "series": 0},
                {"label": f"@Val!A{CON['comp']}", "value": f"@Val!C{CON['comp']}", "series": 1},
                {"label": f"@Val!A{CON['tp']}", "value": f"@Val!C{CON['tp']}", "series": 2},
                {"label": f"@Val!A{CON['cons']}", "value": f"@Val!C{CON['cons']}", "series": 0}],
             "ref": {"label": f"@Val!A{CON['px']}", "value": f"@Val!C{CON['px']}"}},
            {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值", "說明"], "rows": vtab(CON_MAIN, CON, ["A", "C", "F"], ["", "n2", ""], CON_RF)},
            {"type": "kv", "items": [[f"@Val!A{SO['tps']}", [f"@Val!C{SO['tps']}|n2", "／", f"@Val!D{SO['tps']}|n2", "／", f"@Val!E{SO['tps']}|n2", "（情境 A／B／C）"]]]},
            {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（級距、價格帶、共識差距、對照口徑、檢查列、方法）", "head": ["項目", "數值", "對照", "說明"],
             "rows": vtab(CON_REST, CON, ["A", "C", "D", "F"], ["", "n2", "n2", ""], CON_RF)},
            {"type": "table", "caption": "DCF 折現率（主要項目）", "head": ["項目", "數值"],
             "rows": [[f"@Val!A{r}", f"@Val!C{r}|p2"] for r in (V_W, V_KE, V_KD, V_RON)] + [[f"@In!B{INR['tp_g']}", f"@In!E{INR['tp_g']}|p1"]]},
            {"type": "table", "collapse": True, "summary": "展開：折現率的組成（股權市值、負債、無風險利率、β、股權風險溢酬、稅前債務成本、稅率）", "head": ["項目", "數值"],
             "rows": [[f"@Val!A{V_E}", f"@Val!C{V_E}|n1"], [f"@Val!A{V_D}", f"@Val!C{V_D}|n1"]]
                     + [[f"@In!B{INR[k]}", f"@In!E{INR[k]}|{f}"] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("kd", "p2"), ("etr", "p1"))]},
            {"type": "table", "caption": "DCF 彙總（三情境；主要項目）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": vtab(DC_MAIN, DC, ["A", "C", "D", "E"], ["", "n1", "n1", "n1"], DC_RF)},
            {"type": "table", "collapse": True, "summary": "展開：DCF 的逐年無槓桿自由現金流與永續價值", "head": ["項目", "情境 A", "情境 B", "情境 C"],
             "rows": vtab(DC_REST, DC, ["A", "C", "D", "E"], ["", "n1", "n1", "n1"], DC_RF)},
            {"type": "table", "caption": "分部加總（三情境；主要項目）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": vtab(SO_MAIN, SO, ["A", "C", "D", "E"], ["", "n1", "n1", "n1"], SO_RF)},
            {"type": "table", "collapse": True, "summary": "展開：分部加總的其他項目（OCI AI 逐年自由現金流、其餘業務 EBITDA 拆分、淨借款、股數、對照）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
             "rows": vtab(SO_REST, SO, ["A", "C", "D", "E"], ["", "n1", "n1", "n1"], SO_RF)},
            {"type": "table", "collapse": True, "summary": "展開：同業倍數", "head": [f"@Val!{c}{VAL['PEERS'][0][0] - 2}" for c in "ABCDE"],
             "rowgen": {"sheet": "Val", "rows": f"{VAL['PEERS'][0][0]}-{VAL['PEERS'][0][1]},{VAL['PEERS'][1][0]}-{VAL['PEERS'][1][1]}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n1", "x1", ""]}}]},
        {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
            {"type": "table", "head": [f"@RevDCF!{c}4" for c in "ACDEGHIJ"],
             "rowgen": {"sheet": "RevDCF", "rows": f"5-{RVN}", "cols": ["A", "C", "D", "E", "G", "H", "I", "J"], "fmts": ["", "", "n3", "n3", "n3", "n3", "", "n3"]}},
            {"type": "note", "parts": ["@RevDCF!A3"]}]},
        {"id": "recon", "title": "對帳：共識、公司指引、命題檢查", "intro": "@Recon!A1", "blocks": rec_blocks},
        {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks":
            tor("D", "加權目標價相對基準：輸入取低值") + tor("F", "加權目標價相對基準：輸入取高值") + [
            {"type": "table", "collapse": True, "summary": "展開：全部案例",
             "head": ["@Sens!A3", "@Sens!B3", "@Sens!C3", "@Sens!D3", "@Sens!E3", "@Sens!F3", "@Sens!H3", "@Sens!I3", "@Sens!J3", "@Sens!K3"],
             "rowgen": {"sheet": "Sens", "rows": f"{R['SENS'][0]}-{R['SENS'][1]}", "cols": ["A", "B", "C", "D", "E", "F", "H", "I", "J", "K"],
                        "fmts": ["", "", "n2", "sp1", "n2", "sp1", "p1", "p1", "n1", "n1"]}},
            {"type": "table", "collapse": True, "summary": "展開：WACC × 永續成長率的加權目標價", "caption": f"@Sens!A{R['GRID'][0]}", "head": [f"@Sens!A{R['GRID'][0]+1}", f"@Sens!C{R['GRID'][0]+1}", f"@Sens!D{R['GRID'][0]+1}", f"@Sens!E{R['GRID'][0]+1}"],
             "rowgen": {"sheet": "Sens", "rows": f"{R['GRID'][0]+2}-{R['GRID'][1]}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"]}}]},
        {"id": "status", "title": "資料狀態", "blocks": [
            {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"],
                                     ["Tokenomics 快照版本", "@In!H" + str(INR["he_gb300"])]]},
            {"type": "h3", "text": "資料與口徑"},
            {"type": "list", "items": ["@README!A9", "@README!A10", "@Q!A2", "@In!A2"]},
            {"type": "h3", "text": "已知限制"},
            {"type": "note", "parts": [f"@README!A{rd_lim}"]},
            {"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"},
            {"type": "list", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "items": [f"@README!A{rd_drv + k}" for k in range(1, 6)]},
            {"type": "h3", "text": "附錄：版本要點"},
            {"type": "list", "collapse": True, "summary": "展開：近幾版的修改要點", "items": ["@README!A1", f"@README!A{rfind('v0.5.0（')}", f"@README!A{rfind('v0.4.7 變更')}", f"@README!A{rfind('與 v0.3.1 的差異')}"]},
            {"type": "h3", "text": "風險登錄（未進主模型的尾部風險）"},
            {"type": "list", "collapse": True, "summary": "展開：風險、是否建模、理由、觀察指標", "items": [f"@README!A{rd_risk + k}" for k in range(1, 6)]},
            {"type": "h3", "text": "「未找到」清單"},
            {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源",
             "sheets": {"In": ["B"], "Q": ["A"], "Val": ["A"], "Recon": ["A"]}, "empty": "沒有標為未找到的項目。"}]},
    ],
}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json written", X)
