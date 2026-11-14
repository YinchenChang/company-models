#!/usr/bin/env python3
"""產生 coreweave/html.json（tools/company_html 的設定檔；版面規格 docs/plan/個股HTML版面規格.md v1.3（v5.0.2 起；章節：命題→估值→敏感度→反向 DCF→損益→驅動→資金缺口與融資→對帳→資料狀態）；樣板 oracle/scripts/make_html_json.py）。
列號取自 build_coreweave.py 的列號對照（CRWV_ROWMAP），In 頁列號取自成品 xlsx 的鍵欄。設定檔文字不含數字（產生器規則）。
用法：python3 coreweave/scripts/make_html_json.py rowmap.json"""
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = json.load(open(sys.argv[1], encoding="utf-8"))
X = sorted(f for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
IW = wb["In"]
INR = {IW.cell(r, 1).value: r for r in range(4, IW.max_row + 1) if IW.cell(r, 1).value}
SCN_H = 1 + next(r for r in range(4, IW.max_row + 1) if "情境驅動對照" in str(IW.cell(r, 1).value or ""))
TR, PR, CR, CON, DC, CO, VR, FS, FR = R["TR"], R["PR"], R["CR"], R["CON"], R["DC"], R["CO"], R["VR"], R["FS"], R["FR"]
TB, PB, CB = TR["B"], PR["B"], CR["B"]
RL, RCR, RVN = R["RL"], R["RCR"], R["RVN"]

SEC_SUB = {"A": "股價、股數、日期", "B": "在役 MW 路徑", "C": "世代組合與汰換", "D": "每 MW 經濟性（Tokenomics 快照）與定價", "E": "實績錨點（季報）",
           "F": "現金流與融資瀑布", "G": "折現率、目標價規則與同業", "K": "公司指引、宣稱與對照（不進模型）"}
BIZ = {"ai": ["B", "C", "D"], "co": ["A", "E", "F", "G", "K"]}
SECKEYS, _sec = {}, None
for r in range(4, SCN_H - 2):
    a_, b_ = IW.cell(r, 1).value, IW.cell(r, 2).value
    if a_ and b_ in (None, ""):
        _sec = str(a_).split(".")[0].strip(); continue
    if a_ and _sec in SEC_SUB: SECKEYS.setdefault(_sec, []).append(a_)
assert set(SECKEYS) == set(SEC_SUB), sorted(SECKEYS)
_got = [k for v in BIZ.values() for x in v for k in SECKEYS[x]]
assert len(_got) == len(set(_got)) == sum(len(v) for v in SECKEYS.values()), "In 參數歸組不完整或重複"
has_rng = lambda k: IW.cell(INR[k], 4).value not in (None, "") or IW.cell(INR[k], 6).value not in (None, "")
IN_HEAD = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"]
def in_fmt(k):
    v, u = IW.cell(INR[k], 5).value, str(IW.cell(INR[k], 3).value or "")
    if hasattr(v, "year"): return "d"
    if u in ("%", "pt"): return "p1"
    if u in ("年", "MW", "MW／年"): return "n0"
    if isinstance(v, (int, float)) and abs(v) >= 100: return "n1"
    return "n3"
def in_row(k):
    f = in_fmt(k); r = INR[k]
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]
def in_tables(biz, label):
    a, b = [], []
    for x in BIZ[biz]:
        ra = [in_row(k) for k in SECKEYS[x] if has_rng(k)]; rb = [in_row(k) for k in SECKEYS[x] if not has_rng(k)]
        if ra: a += [{"sub": SEC_SUB[x]}] + ra
        if rb: b += [{"sub": SEC_SUB[x]}] + rb
    out = []
    if a: out.append({"type": "table", "collapse": True, "summary": f"展開：{label}的參數（有敏感度區間；基準進目標價，低／高只用於敏感度頁）", "head": IN_HEAD, "rows": a})
    if b: out.append({"type": "table", "collapse": True, "summary": f"展開：{label}的其他固定輸入（實績、公司揭露、換算值）", "head": IN_HEAD, "rows": b})
    return out
def met(name, rows, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": rows}
    if chart: m["chart"] = True
    return m
V3 = lambda keys, src, f: [[f"@Val!A{src[k]}", f"@Val!C{src[k]}|{f.get(k, 'n1')}", f"@Val!D{src[k]}|{f.get(k, 'n1')}", f"@Val!E{src[k]}|{f.get(k, 'n1')}"] for k in keys]
CONF = {"mup": "x1", "mdn": "x1", "up": "sp1", "upc": "sp1", "dist": "p1", "dcfup": "sp1", "cmpup": "sp1", "adjup": "sp1", "tvsh": "p1", "rec": "", "inr": "", "dvalid": "", "bn_scn": "", "bn_edge": "", "bn_dcf": "", "upchk": "p1"}
def ctab(keys):
    return [[f"@Val!A{CON[k]}", f"@Val!C{CON[k]}|{CONF.get(k, 'n2')}" if CONF.get(k, 'n2') else f"@Val!C{CON[k]}"] for k in keys]
RW = wb["Recon"]
def rrows(keys, fmts=None):
    out = []
    for k in keys:
        r = RCR[k]; f = (fmts or {}).get(k, "n2")
        row = [f"@Recon!A{r}"]
        for c, ff in zip("BCDEFGH", (f, f, "sp1", f, "sp1", f, "sp1")):
            v = RW[f"{c}{r}"].value
            row.append(f"@Recon!{c}{r}|{ff}" if isinstance(v, (int, float)) else (f"@Recon!{c}{r}" if v is not None else ""))
        out.append(row)
    return out
REC_HEAD = [f"@Recon!{c}3" for c in "ABCDEFGH"]
SW = wb["Sens"]
def srows(col, big):
    out = []
    for r in range(R["SENS"][0], R["SENS"][1] + 1):
        v = SW[f"{col}{r}"].value
        if not isinstance(v, (int, float)): continue
        if (abs(v) >= 0.1) == big and (big or abs(v) > 0): out.append(str(r))
    return ",".join(out)
def tor(col, title):
    b, sm = srows(col, True), srows(col, False)
    out = [{"type": "tornado", "title": title + "：變動達兩位數百分比的案例", "sheet": "Sens", "rows": b, "label_col": "A", "value_col": col, "fmt": "sp1"}] if b else []
    if sm: out.append({"type": "tornado", "collapse": True, "summary": "展開：" + title + "，變動未達兩位數百分比的案例", "title": title, "sheet": "Sens", "rows": sm, "label_col": "A", "value_col": col, "fmt": "sp1"})
    return out
YR = "D{r}:J{r}"
pl3 = lambda key, src=PR, sh="PL": {s: f"@{sh}!" + YR.format(r=src[s][key]) for s in "ABC"}
ch = RCR["chart_h"]
theme1 = {"type": "theme", "title": "每 MW 年收入：模型 vs Tokenomics 錨 vs 公司宣稱新約價格（情境 B，百萬美元／MW·年）", "x": f"@Recon!C{ch}:H{ch}", "fmt": "n1", "stack": False,
          "bars": [{"name": f"@Recon!A{RCR['ch_m']}", "y": f"@Recon!C{RCR['ch_m']}:H{RCR['ch_m']}"}],
          "lines": [{"name": "Tokenomics 錨：在役世代加權每 MW 年持有成本", "y": f"@Recon!C{RCR['ch_tk']}:H{RCR['ch_tk']}", "color": "#3b3838"},
                    {"name": f"@Recon!A{RCR['ch_cl']}", "y": f"@Recon!C{RCR['ch_cl']}:H{RCR['ch_cl']}", "color": "#b4472f", "dash": True, "labels": True}]
          + ([{"name": "支持現價所需的每 MW 年收入（反向 DCF）", "y": f"@Recon!C{RCR['ch_req']}:H{RCR['ch_req']}", "color": "#c29a2b", "dash": True}] if "ch_req" in RCR else []),
          "caption": "長條＝本模型每 MW 年收入（Tokenomics 持有成本 × 定價倍數）；黑線＝Tokenomics 錨；紅虛線＝公司法說宣稱的新約價格（利害關係人說法，只對照）；金虛線＝支持現價所需的每 MW 年收入（整本重算反解）。"}
fs = "CDE"
theme2 = {"type": "theme", "title": "資金缺口與融資來源（第四季至二〇三〇年合計，十億美元）× 三情境", "x": "@Fund!C3:E3", "fmt": "n0", "stack": True, "total": f"@Fund!C{FS['cum']}:E{FS['cum']}",
          "bars": [{"name": f"@Fund!A{FS[k]}", "y": f"@Fund!C{FS[k]}:E{FS[k]}"} for k in ("ig", "eq", "hy")],
          "lines": [{"name": f"@Fund!A{FS['gate']}", "y": f"@Fund!C{FS['gate']}:E{FS['gate']}", "color": "#b4472f", "labels": "below"}],
          "caption": "堆疊＝外部融資的三種來源（投資級債、新股、高息債），標籤為合計；紅線＝收支關卡（自由現金流合計，負值＝缺口）。容量越大，缺口越大。"}
rp = R["RPB"]
theme3 = {"type": "theme", "title": "Backlog 轉營收：RPO 排程 vs 模型營收（十億美元）", "x": "@RPO!D4:J4", "fmt": "n0", "stack": False,
          "bars": [{"name": f"@RPO!A{rp['sched']}", "y": f"@RPO!D{rp['sched']}:J{rp['sched']}", "color": "#9aa5ad"}],
          "lines": [{"name": f"@RPO!A{rp['rev' + s]}", "y": f"@RPO!D{rp['rev' + s]}:J{rp['rev' + s]}", "color": c, **({"dash": True} if s != "B" else {"labels": True})} for s, c in (("A", "#0b7a80"), ("B", "#3b3838"), ("C", "#b4472f"))],
          "caption": "灰長條＝RPO 依公司揭露的時間分桶平均攤到各年（利害關係人揭露）；三條線＝三情境模型營收。線高於長條的部分要靠新簽合約或隨需。"}

# ---- v1.3 新增圖（數字取自 xlsx；列號依 rowmap）
SEGD, PLS = R["SEGD"], R["PLS"]
SCOL = {"A": "#0b7a80", "B": "#3b3838", "C": "#b4472f"}
SCOL_L = {"A": "#6fc7cb", "B": "#9aa5ad", "C": "#e3a593"}
SNAME = {"A": "情境 A（樂觀）", "B": "情境 B（基準）", "C": "情境 C（保守）"}
theme_gate = {"type": "theme", "title": "收支關卡：累計自由現金流 × 三情境（十億美元；FY2030 的累計值＝收支關卡，負值＝入不敷出）", "x": "@CF!D4:H4", "fmt": "n0", "stack": False,
              "bars": [{"name": ["累計：", SNAME[s]], "y": f"@CF!D{R['CFC'][s]}:H{R['CFC'][s]}", "color": SCOL_L[s]} for s in "ABC"],
              "lines": [{"name": "情境 B 每年自由現金流", "y": f"@CF!D{CR['B']['fcf']}:H{CR['B']['fcf']}", "color": "#3b3838", "labels": True}],
              "caption": ["長條＝自 FY2026 起累計的自由現金流（營運現金流 − 資本支出）；黑線＝情境 B 每年的自由現金流。五年合計（收支關卡）：情境 A ", f"@Fund!C{FS['gate']}|n1", "、情境 B ", f"@Fund!D{FS['gate']}|n1",
                          "、情境 C ", f"@Fund!E{FS['gate']}|n1", "；合計為負即入不敷出，投資建議不得為買進。"]}
theme_pl = {"type": "theme", "title": "三情境營收（長條，左軸）與營業利益（折線，右軸）（十億美元）", "x": "@PL!D3:H3", "fmt": "n1", "fmt_r": "n1", "stack": False,
            "bars": [{"name": ["營收：", SNAME[s]], "y": f"@PL!D{PR[s]['rev']}:H{PR[s]['rev']}", "color": SCOL_L[s]} for s in "ABC"],
            "lines": [{"name": ["營業利益：", SNAME[s]], "y": f"@PL!D{PR[s]['oi']}:H{PR[s]['oi']}", "axis": "r", "color": SCOL[s], **({"labels": True} if s == "B" else {"dash": True})} for s in "ABC"],
            "caption": "長條＝三情境營收（左軸）；折線＝三情境 GAAP 營業利益（右軸，刻度與左軸不同）。"}
seg_years = f"C{SEGD['hdr']}:H{SEGD['hdr']}"
theme_seg = {"type": "theme", "title": "分部營收推導（情境 B）：計費 MW × 每 MW 年收入，拆成長約與隨需（十億美元）", "x": f"@Seg!{seg_years}", "fmt": "n1", "stack": True,
             "total": f"@Seg!C{SEGD['tot']}:H{SEGD['tot']}", "pad": 1.15,
             "bars": [{"name": f"@Seg!A{SEGD['act']}", "y": f"@Seg!C{SEGD['act']}:H{SEGD['act']}", "color": "#9aa5ad"},
                      {"name": "長約營收", "y": f"@Seg!C{SEGD['long']}:H{SEGD['long']}", "color": "#0b7a80"},
                      {"name": "隨需營收", "y": f"@Seg!C{SEGD['spot']}:H{SEGD['spot']}", "color": "#c29a2b"}],
             "lines": [{"name": "計費 MW", "y": f"@Seg!C{SEGD['bill']}:H{SEGD['bill']}", "axis": "r", "color": "#6f6bae", "labels": True}],
             "fmt_r": "n0",
             "caption": "灰＝實績（公司未揭露長約／隨需拆分；FY2026 為前三季）；青綠＝長約營收；金＝隨需營收；標籤為合計＝合併營收。紫線＝計費 MW（右軸；FY2026 為第四季值）。"}
SEG_ROWS = [SEGD[k] for k in ("bill", "anc", "pl", "ps", "od", "act", "long", "spot", "ovr", "tot", "lsh", "chk")]
SEG_FMT = {SEGD["bill"]: "n0", SEGD["anc"]: "n2", SEGD["pl"]: "n2", SEGD["ps"]: "n2", SEGD["od"]: "p0", SEGD["lsh"]: "p1", SEGD["chk"]: "n3"}
def segrow(r_):
    f = SEG_FMT.get(r_, "n2")
    return [f"@Seg!A{r_}"] + [f"@Seg!{c}{r_}|{f}" if wb["Seg"][f"{c}{r_}"].value is not None else "" for c in "CDEFGHIJ"]
PLC = "CDEFGHIJK"
def plsrow(k, f):
    r_ = PLS[k]
    return [f"@PL!A{r_}"] + [f"@PL!{c}{r_}|{f}" if wb["PL"][f"{c}{r_}"].value is not None else "" for c in PLC]
IS_ROWS = [plsrow(k, f) for k, f in (("rev", "n2"), ("revg", "p1"), ("cogs", "n2"), ("gp", "n2"), ("gpm", "p1"), ("opx", "n2"), ("oi", "n2"), ("oim", "p1"), ("int", "n2"), ("oth", "n2"),
                                    ("pti", "n2"), ("tax", "n2"), ("ni", "n2"), ("nim", "p1"), ("eps", "n2"))] + [{"sub": f"@PL!A{PLS['ratio']}"}] + \
          [plsrow(k, "p1") for k in ("revg", "gpm", "oim", "nim", "nig")]
VAL_COMP = V3(["eb", "ev", "nd27", "sh27", "v12"], CO, {"sh27": "n3", "v12": "n2"})
VAL_COMP[1][0] = "企業價值"
DCF_RATE = [[f"@Val!A{VR[k]}" if k != "wacc" else "WACC（加權平均資金成本）", f"@Val!C{VR[k]}|p2"] for k in ("wacc", "ke", "kdat", "ronic")] + [[f"@In!B{INR['tp_g']}", f"@In!E{INR['tp_g']}|p1"]]
REC_J4 = rrows(["q9rev", "q3pm", "q3eb", "j4m", "j4tk", "j4s"] + (["j4req"] if "j4req" in RCR else []), {"q3eb": "p1"})
for row in REC_J4:
    if row[0] == f"@Recon!A{RCR['j4tk']}": row[0] = "  Tokenomics 錨（新約主力機型每 MW 年持有成本）vs 公司宣稱中點"

cfg = {
    "company": "CoreWeave", "ticker": "CRWV", "wide_text": True, "layout": "v1.3",
    "hero": {"rec": f"@Val!C{CON['rec']}",
             "tiles": [[f"@Val!A{CON['tp']}", f"@Val!C{CON['tp']}|n2", ["上檔空間 ", f"@Val!C{CON['up']}|sp1"]],
                       [f"@Val!A{CON['comp']}", f"@Val!C{CON['comp']}|n2", ["上檔空間 ", f"@Val!C{CON['upc']}|sp1"]],
                       [f"@Val!A{CON['px']}", f"@Val!C{CON['px']}|n2", ""]],
             "note": f"@README!A{RL['hero13']}"},
    "thesis": f"@README!A{RL['thesis']}",
    "question": "每增加一 GW 的產能，是否創造高於資金成本的價值；資金缺口要靠多少新債與新股補、股東被稀釋多少",
    "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
    "sections": [
        {"id": "thesis", "title": "命題與核心問題", "blocks": [
            {"type": "h3", "text": "命題"},
            {"type": "list", "items": [f"@README!A{RL['show'] + k}" for k in (1, 2, 3)]},
            {"type": "list", "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案）", "items": [f"@README!A{RL['ans'] + k}" for k in (1, 2, 3)]},
            {"type": "list", "collapse": True, "summary": "展開：主要失效條件（出現任一可觀察證據，結論須重估）", "items": [f"@README!A{RL['fail'] + k}" for k in (1, 2, 3)]},
            theme_gate,
            {"type": "h3", "text": "重點觀察"},
            {"type": "list", "ordered": True, "items": [f"@README!A{RL[k]}" for k in ("ob1", "ob2", "ob3", "ob4", "ob5")]},
            {"type": "list", "collapse": True, "summary": "展開：重點觀察的模型細節（含參數與指引對照）", "items": [f"@README!A{RL[k]}" for k in ("rec", "o1", "o2", "o3", "o4", "o6")]},
            {"type": "table", "collapse": True, "summary": "展開：命題答案摘要（三情境）", "head": [f"@Thesis!{c}{R['SUM'][0]}" for c in "ABCDE"],
             "rowgen": {"sheet": "Thesis", "rows": f"{R['SUM'][0] + 1}-{R['SUM'][1]}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n1", "n1", "n1"],
                        "row_fmt": {str(R['SUM'][0] + k): {"C": "p1", "D": "p1", "E": "p1"} for k in (3, 4, 5)} | {str(R['SUM'][0] + 1): {"C": "n0", "D": "n0", "E": "n0"}}}}]},
        {"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": [
            {"type": "bars", "title": "目標價與現價", "fmt": "n1", "items": [
                *([{"label": f"@Val!A{CON['dcf']}", "value": f"@Val!C{CON['dcf']}", "series": 0}] if (wb["Val"][f"C{CON['dcf']}"].value or 0) > 0 else []), {"label": f"@Val!A{CON['comp']}", "value": f"@Val!C{CON['comp']}", "series": 1},
                {"label": f"@Val!A{CON['tp']}", "value": f"@Val!C{CON['tp']}", "series": 2}, {"label": f"@Val!A{CON['cons']}", "value": f"@Val!C{CON['cons']}", "series": 0}],
             "ref": {"label": f"@Val!A{CON['px']}", "value": f"@Val!C{CON['px']}"}},
            {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值"], "rows": ctab(["tp", "dcf", "comp", "px", "up", "upc", "rec", "dvalid", "cons"])},
            {"type": "kv", "items": [[f"@Val!A{CON['tps']}", [f"@Val!C{CON['tps']}|n2", "／", f"@Val!D{CON['tps']}|n2", "／", f"@Val!E{CON['tps']}|n2", "（情境 A／B／C）"]]]},
            {"type": "list", "collapse": True, "summary": "展開：評等標示、級距邊界、評等脆弱點與方法註記（原橫幅說明）",
             "items": [f"@Val!C{CON['rec_lab']}", f"@Val!C{CON['bn_edge']}", f"@Val!C{CON['bn_frag']}", f"@Val!C{CON['bn_scn']}", f"@Val!C{CON['bn_dcf']}", f"@Val!F{CON['rec']}"]},
            {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（上檔空間歸因、區間、增量價值、級距距離、舊版對照）", "head": ["項目", "數值"],
             "rows": ctab(["mup", "mdn", "comp0", "dcfup", "cmpup", "adjup", "upchk", "lo", "hi", "inr", "npvAB", "npvBC", "npvps", "dist", "v413", "chg", "tvsh"])},
            {"type": "table", "caption": "DCF 折現率（主要項目）", "head": ["項目", "數值"], "rows": DCF_RATE},
            {"type": "table", "collapse": True, "summary": "展開：折現率組成（股權市值、負債、無風險利率、β、股權風險溢酬、稅前債務成本、稅率）", "head": ["項目", "數值"],
             "rows": [[f"@Val!A{VR['E']}", f"@Val!C{VR['E']}|n1"], [f"@Val!A{VR['D']}", f"@Val!C{VR['D']}|n1"]] + [[f"@In!B{INR[k]}", f"@In!E{INR[k]}|{f}"] for k, f in (("rf", "p2"), ("beta_u", "n3"), ("beta", "n2"), ("erp", "p1"), ("kd", "p2"), ("tax_lt", "p1"))]},
            {"type": "table", "caption": "DCF 彙總（三情境；主要項目）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
             "rows": V3(["ev", "tvsh", "valid", "eqv", "ps", "v12"], DC, {"tvsh": "p1", "valid": "", "ps": "n2", "v12": "n2"})},
            {"type": "table", "collapse": True, "summary": "展開：DCF 逐年 FCFF 與永續價值", "head": ["項目", "情境 A", "情境 B", "情境 C"],
             "rows": V3(["2026", "2027", "2028", "2029", "2030", "2031", "2032", "tvb", "tv", "nol"], DC, {})},
            {"type": "table", "caption": "Comparable（三情境；主要項目）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": VAL_COMP},
            {"type": "table", "collapse": True, "summary": "展開：同業倍數", "head": [f"@Val!{c}{R['PEERH']}" for c in "ABCDEGHIJ"],
             "rowgen": {"sheet": "Val", "rows": f"{R['PEERS'][0]}-{VR['mult']}", "cols": list("ABCDEGHIJ"), "fmts": ["", "", "n1", "n2", "", "p1", "n2", "p1", "p1"],
                        "row_fmt": {str(VR["altm"]): {"C": "n2", "D": "n2", "E": "n2"}}}}]},
        {"id": "sens", "title": "敏感度", "intro": "一次只動一個輸入（取低值或高值），整本模型重算，看加權目標價相對基準變動多少；長條只列變動達兩位數百分比的案例，其餘收在摺疊。",
         "blocks": [dict(b, label_col="O") for b in tor("D", "加權目標價相對基準：輸入取低值") + tor("F", "加權目標價相對基準：輸入取高值")] + [
            {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法與收支關卡）", "head": [f"@Sens!{c}3" for c in "ABCDEFHIJKMN"],
             "rowgen": {"sheet": "Sens", "rows": f"{R['SENS'][0]}-{R['SENS'][1]}", "cols": list("ABCDEFHIJKMN"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2", "n2", "n1", "n1", "", ""]}},
            {"type": "table", "collapse": True, "summary": "展開：WACC × 永續成長率的加權目標價", "caption": f"@Sens!A{R['GRID'][0]}", "head": [f"@Sens!{c}{R['GRID'][0] + 1}" for c in "ACDE"],
             "rowgen": {"sheet": "Sens", "rows": f"{R['GRID'][0] + 2}-{R['GRID'][1]}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2"]}}]},
        {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": f"@README!A{RL['rev13']}", "blocks": [
            {"type": "table", "head": [f"@RevDCF!{c}4" for c in "ACGHIJ"],
             "rowgen": {"sheet": "RevDCF", "rows": f"5-{RVN['dn'] if 'dn' in RVN else RVN['px']}", "cols": list("ACGHIJ"), "fmts": ["", "n2", "n2", "n2", "sp1", "n1"]}},
            {"type": "table", "collapse": True, "summary": "展開：反解明細（驅動與反解係數；整本重算）", "head": [f"@RevDCF!{c}4" for c in "ABCDGHIJ"],
             "rowgen": {"sheet": "RevDCF", "rows": f"5-{RVN['dn'] if 'dn' in RVN else RVN['px']}", "cols": list("ABCDGHIJ"), "fmts": ["", "", "n2", "n3", "n2", "n2", "sp1", "n1"]}},
            {"type": "note", "collapse": True, "summary": "展開：方法說明", "parts": ["@RevDCF!A2"]},
            {"type": "note", "parts": ["@RevDCF!A3"]}]},
        {"id": "pl", "title": "損益", "intro": f"@README!A{RL['pl13']}", "blocks": [
            {"type": "table", "caption": "標準損益表（GAAP；十億美元，EPS 美元；FY2026 起為情境 B）", "head": [f"@PL!{c}{PLS['hdr']}" for c in "A" + PLC], "rows": IS_ROWS},
            {"type": "table", "collapse": True, "summary": "展開：完全稀釋加權股數與核對列", "grid": True, "head": [f"@PL!{c}{PLS['hdr']}" for c in "A" + PLC],
             "rows": [plsrow("sh", "n3"), plsrow("chk1", "n3"), plsrow("chk2", "n3"), plsrow("chk3", "n3")]},
            theme_pl,
            {"type": "pl", "collapse": True, "summary": "展開：三情境損益（營收、調整後 EBITDA、營業利益率、EPS）", "years": "@PL!D3:J3", "metrics": [
                met("營收（十億美元）", pl3("rev"), "n1", True), met("調整後 EBITDA（十億美元）", pl3("aeb"), "n1", True),
                met("調整後 EBITDA 率", pl3("aebm"), "p1"), met("營業利益率（GAAP）", pl3("oim"), "p1", True), met("EPS（GAAP 稀釋，美元）", pl3("eps"), "n2")]},
            {"type": "table", "collapse": True, "summary": "展開：情境 B 的損益明細（模型列，含調整後 EBITDA 與 NOL）", "caption": f"@README!A{RL['plb13']}", "head": [f"@PL!{c}3" for c in "ACDEFGHIJ"],
             "rowgen": {"sheet": "PL", "rows": f"{PB['ai']}-{PB['eps']}", "cols": list("ACDEFGHIJ"), "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2", "n2"],
                        "row_fmt": {str(PB[k]): {c: "p1" for c in "CDEFGHIJ"} for k in ("revg", "aebm", "oim")}}},
            {"type": "table", "collapse": True, "summary": "展開：分部核對（三情境合計絕對差，應為零）", "head": [f"@PL!{c}3" for c in "ACDEFGHIJ"],
             "rowgen": {"sheet": "PL", "rows": f"{R['CHK'][0]}-{R['CHK'][1]}", "cols": list("ACDEFGHIJ"), "fmts": ["", "n3", "n3", "n3", "n3", "n3", "n3", "n3", "n3"]}}]},
        {"id": "drivers", "title": "驅動因子 → 推導量", "intro": f"@README!A{RL['drv13']}", "blocks": [
            {"type": "h3", "text": "一、分部營收推導（AI 雲端平台：長約＋隨需）"},
            {"type": "note", "parts": [f"@README!A{RL['seg13']}"]},
            theme_seg,
            {"type": "table", "caption": "分部營收推導（情境 B；FY2026 欄的 MW 與每 MW 年收入為第四季）", "head": [f"@Seg!{c}{SEGD['hdr']}" for c in "ACDEFGHIJ"], "rows": [segrow(x) for x in SEG_ROWS]},
            {"type": "h3", "text": "二、每 MW 年收入與 backlog"},
            theme1,
            theme3,
            {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": [f"@In!{c}{SCN_H}" for c in "BDEFCG"],
             "rows": [[f"@In!B{SCN_H + i}", f"@In!D{SCN_H + i}|{f}", f"@In!E{SCN_H + i}|{f}", f"@In!F{SCN_H + i}|{f}", f"@In!C{SCN_H + i}", f"@In!G{SCN_H + i}|tag"]
                      for i, f in zip(range(1, 7), ("n0", "n0", "n0", "n0", "n0", "p0"))]},
            {"type": "h3", "text": "三、參數與推導明細（查核用，展開閱讀）"},
            {"type": "note", "collapse": True, "summary": "展開：命題模組的公式說明", "parts": ["@Thesis!A2"]}] + in_tables("ai", "AI 雲端") + [
            {"type": "table", "collapse": True, "summary": "展開：AI 雲端：在役 MW → 營收 → 費用、資本支出、折舊 → 利潤與報酬（情境 B）", "head": [f"@Thesis!{c}3" for c in "ABCDEFGHIJK"],
             "rowgen": {"sheet": "Thesis", "rows": f"{TB['conn']}-{TB['bill']},{TB['anc']},{TB['pm']},{TB['rev']}-{TB['ebm']},{TB['cx']}-{TB['roic']}",
                        "cols": list("ABCDEFGHIJK"), "fmts": ["", "", "n2", "n2", "n2", "n2", "n2", "n2", "n2", "n2", "n2"],
                        "row_fmt": {**{str(TB[k]): {c: "n0" for c in "CDEFGHIJK"} for k in ("conn", "add", "ret", "avg", "bill")},
                                    **{str(TB[k]): {c: "p1" for c in "CDEFGHIJK"} for k in ("ebm", "om", "roic")}}}}] + in_tables("co", "全公司")},
        {"id": "funding", "title": "資金缺口與融資（收支模組）", "intro": f"@README!A{RL['fund13']}", "blocks": [
            theme2,
            {"type": "table", "caption": "三情境摘要", "head": ["@Fund!A3", "@Fund!C3", "@Fund!D3", "@Fund!E3"],
             "rowgen": {"sheet": "Fund", "rows": f"{FS['gate']}-{FS['chk']}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"],
                        "row_fmt": {str(FS["ndmax"]): {c: "n2" for c in "CDE"}, str(FS["icmin"]): {c: "n2" for c in "CDE"}, str(FS["warn"]): {c: "" for c in "CDE"},
                                    str(FS["eqmax"]): {c: "p1" for c in "CDE"}, str(FS["dil"]): {c: "p1" for c in "CDE"}, str(FS["chk"]): {c: "n3" for c in "CDE"}}}},
            {"type": "pl", "years": "@CF!D4:J4", "metrics": [
                met("自由現金流（十億美元）", {s: f"@CF!D{CR[s]['fcf']}:J{CR[s]['fcf']}" for s in "ABC"}, "n1", True),
                met("外部融資需求（十億美元）", {s: f"@CF!D{CR[s]['need']}:J{CR[s]['need']}" for s in "ABC"}, "n1", True),
                met("借款合計（期末，十億美元）", {s: f"@CF!D{CR[s]['debt']}:J{CR[s]['debt']}" for s in "ABC"}, "n1"),
                met("累計稀釋", {s: f"@CF!D{CR[s]['dil']}:J{CR[s]['dil']}" for s in "ABC"}, "p1")]},
            {"type": "note", "collapse": True, "summary": "展開：收支模組的計算順序", "parts": ["@Fund!A2"]},
            {"type": "table", "collapse": True, "summary": "展開：情境 B 逐年來源與用途", "head": [f"@Fund!A{R['FUH']}"] + [f"@Fund!{c}{R['FUH']}" for c in "DEFGHIJ"],
             "rowgen": {"sheet": "Fund", "rows": f"{FR['flow']}-{FR['dil']}", "cols": ["A", "D", "E", "F", "G", "H", "I", "J"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2"],
                        "row_fmt": {str(FR[k]): {c: "p1" for c in "DEFGHIJ"} for k in ("eqm", "dil")}}},
            {"type": "table", "collapse": True, "summary": "展開：債務逐檔（第二季十季報＋第三季新增）", "head": ["@Debt!A4", "@Debt!B4", "@Debt!C4", "@Debt!D4", "@Debt!E4"],
             "rowgen": {"sheet": "Debt", "rows": f"5-{R['D']['diff']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "", "p2", "n3"]}},
            {"type": "table", "collapse": True, "summary": "展開：到期與攤還表、再融資備忘", "head": [f"@Debt!A{R['D']['hdr']}"] + [f"@Debt!{c}{R['D']['hdr']}" for c in "DEFGHIJ"],
             "rowgen": {"sheet": "Debt", "rows": f"{R['D']['old']}-{R['D']['refi']}", "cols": ["A", "D", "E", "F", "G", "H", "I", "J"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2"]}},
            {"type": "table", "collapse": True, "summary": "展開：機房租金路徑", "head": ["@Lease!A4"] + [f"@Lease!{c}4" for c in "DEFGHIJ"],
             "rowgen": {"sheet": "Lease", "rows": f"5-{R['LSB']['ratio']}", "cols": ["A", "D", "E", "F", "G", "H", "I", "J"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2"]}},
            {"type": "table", "collapse": True, "summary": "展開：具名站點與容量上限", "head": ["@Sites!A4", "@Sites!B4", "@Sites!C4", "@Sites!D4", "@Sites!E4", "@Sites!F4"],
             "rowgen": {"sheet": "Sites", "rows": f"5-{R['ST']['res']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n0", "n0", "", ""]}}]},
        {"id": "recon", "title": "對帳：公司指引、公司實況、公司宣稱、容量", "intro": "@Recon!A1", "blocks": [
            {"type": "table", "caption": "@Recon!A4", "head": REC_HEAD, "rows": rrows(["rev26", "rev26pos", "aoi26", "aoi26pos", "cx26", "cxq4", "cxq4pos", "rev4req"])},
            {"type": "table", "collapse": True, "summary": "展開：指引的其他項目", "head": REC_HEAD, "rows": rrows(["arr", "int4", "int4pos", "intnew", "int4r", "aoiq4", "depq4", "rev4mw"], {"int4r": "p2", "rev4mw": "n0", "depq4": "p1"})},
            {"type": "table", "caption": "公司實況驗證與公司宣稱的新約價格", "head": REC_HEAD, "rows": REC_J4},
            {"type": "table", "collapse": True, "summary": "展開：RPO、類比、Wattonomics、舊法", "head": REC_HEAD, "rows": rrows(["rpo27", "rpo29", "iren", "wat26", "wat30", "old", "cons"], {"wat26": "n0", "wat30": "n0"})}]},
        {"id": "status", "title": "資料狀態", "blocks": [
            {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"],
                                     ["建置識別與模型狀態", [f"@Sens!D{R['GOV']['bid']}", "／", f"@Sens!D{R['GOV']['st']}"]], ["版本與日期", "@README!A1"]]},
            {"type": "h3", "text": "資料與口徑"},
            {"type": "list", "items": [f"@README!A{RL['cal']}", f"@README!A{RL['seg']}", "@Q!A2", "@In!A2", f"@PL!A{PLS['note']}"]},
            {"type": "h3", "text": "已知限制"},
            {"type": "note", "parts": [f"@README!A{RL['lim']}"]},
            {"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"},
            {"type": "list", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "items": [f"@README!A{RL['drv'] + k}" for k in range(1, 6)]},
            {"type": "h3", "text": "附錄：版本要點與目標價變動拆解"},
            {"type": "list", "collapse": True, "summary": "展開：版本要點", "items": [f"@README!A{RL['v502']}", f"@README!A{RL['v501']}", f"@README!A{RL['v50']}"] + ([f"@README!A{RL['dec']}"] if "dec" in RL else [])},
            {"type": "h3", "text": "風險登錄（未進主模型的尾部風險）"},
            {"type": "list", "collapse": True, "summary": "展開：風險、是否建模、理由、觀察指標", "items": [f"@README!A{RL['risk'] + k}" for k in range(1, 7)]},
            {"type": "h3", "text": "「未取得」與「找不到」清單"},
            {"type": "scan", "needle": "找不到", "collapse": True, "summary": "展開：標為「找不到」的輸入與來源", "sheets": {"In": ["H", "I"], "Val": ["A"]}, "empty": "沒有標為找不到的項目。"}]},
    ],
}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json written", X)
