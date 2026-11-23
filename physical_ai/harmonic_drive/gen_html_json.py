# 由成品 xlsx 依列標籤找列號，產生 physical_ai/harmonic_drive/html.json（版面規格 v1.3；列號隨模型變動，勿手改 html.json）。
# 用法：python3 physical_ai/harmonic_drive/gen_html_json.py   （讀 dist/ 內唯一的 xlsx）
import glob
import json
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XS = sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))
assert len(XS) == 1, XS
wb = openpyxl.load_workbook(XS[0], data_only=True)
YRS = "EFGHIJK"          # PL：E＝FY2026 … K＝FY2032
PLC = "CDEFGHIJK"        # PL：C＝FY2024 … K＝FY2032


def find(sh, text, col="A", start=1, end=None, exact=True):
    ws = wb[sh]
    for r in range(start, (end or ws.max_row) + 1):
        v = ws[f"{col}{r}"].value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(f"{sh}!{col}：找不到「{text}」")


# ---- Val ----
VL = {k: find("Val", t, exact=False) for k, t in dict(
    ke="股權成本 ke", kd_at="稅後債務成本", e_mv="股權市值", d_bv="負債（借款", wacc="WACC", g_show="永續成長率 g", ronic="RONIC",
    ev="企業價值 EV", tvs="終值占 EV", dcf="DCF 情境價值推到 12 個月後", eps27="FY2027 EPS（模型）", cmpu="Comparable 未調整（同業",
    cmp="Comparable（× 相對倍數係數", wt="情境加權前", w="情境權重", dcf_tp="DCF 目標價（情境加權）", cmp_tp="Comparable 目標價（情境加權",
    cmpu_tp="Comparable 未調整（情境加權）", tp="目標價（DCF 與 Comparable 加權）", px="現價（股價基準日）", up="上檔空間：加權目標價",
    rating="估值評等", gate="收支關卡：", gate_b="入不敷出判定用", gate_t="入不敷出判定", rec="投資建議", dist="距最近級距邊界",
    up_b="升一級所需目標價", dn_b="降一級門檻目標價", att_d="DCF 貢獻", att_c="未調整 Comparable 貢獻", att_r="相對倍數調整貢獻",
    att_chk="檢查：三段合計", relnow="HDS 現行前瞻倍數", relreq="支持現價所需的相對倍數係數", relflip="評等升一級所需的相對倍數係數",
    t_model="模型折現期間", t_hold="持有期間", tp_date="目標時點").items()}
VL["gate_t"] = find("Val", "入不敷出判定")
VL["wacc"] = find("Val", "WACC")
VL["rec"] = find("Val", "投資建議")
VL["tp_date"] = find("Val", "目標時點")
VL["w"] = find("Val", "情境權重")


def v(k, f="", col="B"):
    return f"@Val!{col}{VL[k]}" + (f"|{f}" if f else "")


def va(k):
    return f"@Val!A{VL[k]}"


# ---- In ----
SCN0 = find("In", "情境驅動對照", exact=False) + 2
IN = {}
for r in range(4, SCN0 - 2):
    k = wb["In"][f"A{r}"].value
    if isinstance(k, str) and wb["In"][f"E{r}"].value is not None and k not in IN:
        IN[k] = r
SCN_END = find("In", "scn_w")

# ---- PL（情境 B 區塊在最上方）----
def pl_rows(s):
    suf = "" if s == "B" else f"（情境 {s}）"
    st = find("PL", f"情境 {s}" + ("（基準）" if s == "B" else ""))
    labs = dict(rob="用途：產業機器人（含協作）", semi="用途：半導體設備", hum="用途：人形／AI 機器人", oth="用途：其他", rev="合併營收",
                jp="日本 營收", cn="中國 營收", na="北美 營收", eu="歐洲 營收", op="營業利益", opm="營業利益率", nonop="業外損益", ord="經常利益",
                ni="歸屬母公司淨利", eps="EPS", da="折舊攤銷", cx="資本支出", dnwc="營運資金增加", nopat="NOPAT（＝營業利益 ×（1 − 稅率））",
                fcff="FCFF", div="股利（淨利 × 配當性向）", gate="FCFF 扣股利（收支）")
    out = {k: find("PL", t + suf, start=st) for k, t in labs.items()}
    out["chk"] = out["eu"] + 1; out["chk2"] = out["eu"] + 2
    return out


PR = {s: pl_rows(s) for s in "ABC"}
PB = PR["B"]

# ---- Thesis ----
def th_rows(s):
    st = find("Thesis", f"情境 {s}")
    labs = dict(inst="全球工業機器人裝機（CY，千台）", instfy="FY 加權裝機（千台）", rob="產業機器人（含協作）營收", semi="半導體設備 營收",
                hu="全球人形出貨（CY，千台）", hufy="FY 加權人形出貨（千台）", asp="人形用諧波單價（千日圓／個）", hum="人形／AI 機器人 營收",
                oth="其他用途（醫療、工具機、航太等）營收", tot="用途合計（＝合併營收）", humsh="人形占合併營收")
    return {k: find("Thesis", t, start=st) for k, t in labs.items()}


TR = {s: th_rows(s) for s in "ABC"}

# ---- Sens ----
SE = wb["Sens"]
case_rows = [r for r in range(8, SE.max_row + 1) if isinstance(SE[f"F{r}"].value, (int, float)) and isinstance(SE[f"A{r}"].value, str)]
SE_END = max(case_rows)
low = [r for r in case_rows if "（敏感度低）" in SE[f"A{r}"].value]
high = [r for r in case_rows if "（敏感度高）" in SE[f"A{r}"].value]
other = [r for r in case_rows if r not in low and r not in high and abs(SE[f"F{r}"].value) >= 0.0005]
SE_NOTE = find("Sens", "讀法：", exact=False)

# ---- Recon ----
RC = {k: find("Recon", t, exact=False) for k, t in dict(
    rev26="FY2026 營收（情境 B）vs 指引", cons_tp="加權目標價 vs 賣方平均目標價", solve_h="反解（整本重算", s1="支持現價（目標價＝現價）所需",
    s2="評等升一級所需的情境 B", s3="支持現價所需的情境 B HDS 人形份額", comp_h="競爭者對照", comp_end="綠的諧波 前瞻 P/E ÷").items()}

# ---- 敏感度 ----
def torn(rows, title, collapse=None):
    b = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "F", "fmt": "sp1", "title": title, "rows": ",".join(map(str, rows))}
    if collapse:
        b.update(collapse=True, summary=collapse)
    return b


sens_blocks = []
for side, rs in (("低", low), ("高", high)):
    big = [r for r in rs if abs(SE[f"F{r}"].value) >= 0.1]
    sens_blocks.append(torn(big or sorted(rs, key=lambda r: -abs(SE[f"F{r}"].value))[:1], f"加權目標價相對基準：輸入取敏感度{side}值，變動達兩位數百分比的案例"))
rest = []
for side, rs in (("低", low), ("高", high)):
    small = [r for r in rs if 0.0005 <= abs(SE[f"F{r}"].value) < 0.1]
    if small:
        rest.append(torn(small, f"加權目標價相對基準：輸入取敏感度{side}值，其餘案例", f"展開：輸入取敏感度{side}值，變動未達兩位數百分比的案例"))
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks + [
    {"type": "note", "parts": ["@Sens!A1"]},
    torn(other, "加權目標價相對基準：情境驅動、人形出貨與份額、情境權重案例"),
    *rest,
    {"type": "table", "collapse": True, "summary": "展開：全部案例（設定、加權目標價、DCF 目標價與估值評等）",
     "head": ["@Sens!A7", "@Sens!B7", "@Sens!E7", "@Sens!F7", "@Sens!G7", "@Sens!H7", "@Sens!I7"],
     "rowgen": {"sheet": "Sens", "rows": f"8-{SE_END}", "cols": ["A", "B", "E", "F", "G", "H", "I"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", ""]}},
    {"type": "note", "parts": [f"@Sens!A{SE_NOTE}"]},
    {"type": "kv", "items": [["@Sens!A2", "@Sens!D2"], ["@Sens!A3", "@Sens!D3"]]}]}

# ---- 章節 ----
cfg = {
    "company": "哈默納科 Harmonic Drive Systems", "ticker": "6324.T", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A2",
    "question": "人形機器人的選擇權能否支撐現價：其他輸入不變時，情境 B 的人形出貨與份額要多高，目標價才等於現價",
    "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
    "hero": {"rec": v("rec"),
             "tiles": [[va("tp"), v("tp", "n0"), ["上檔空間 ", v("up", "sp1")]],
                       [va("cmp_tp"), v("cmp_tp", "n0"), ""],
                       [f"@In!B{IN['px']}", f"@In!E{IN['px']}|n0", ""]],
             "note": ["加權目標價相對現價 ", v("up", "sp1"), "，估值評等", v("rating"), "；現價對 FY2026 指引 EPS 的倍數是核心同業中位數的 ", v("relnow", "n1"),
                      " 倍。其他輸入不變時，情境 B 的 CY2032 全球人形出貨須達 ", f"@Recon!B{RC['s1']}|n0", " 千台（基準 ", f"@In!E{IN['hu32_B']}|n0",
                      " 千台）目標價才等於現價。收支關卡（情境 B，FY2026–FY2030 自由現金流扣股利）", v("gate_b", "n1"), " 十億日圓，", v("gate_t"), "。"]},
    "sections": [],
}
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "compare", "title": "CY2032 全球人形出貨（千台）：模型基準 vs 支持現價所需 vs 評等升一級所需",
     "groups": [{"name": "情境 B 的 CY2032 全球人形出貨（千台）", "fmt": "n0",
                 "items": [{"name": "模型情境 B 基準", "v": f"@In!E{IN['hu32_B']}"}, {"name": "模型情境 A", "v": f"@In!E{IN['hu32_A']}", "ref": True},
                           {"name": "評等升一級所需", "v": f"@Recon!B{RC['s2']}"}, {"name": "支持現價所需", "v": f"@Recon!B{RC['s1']}"}]},
                {"name": "HDS 占全球人形諧波用量份額", "fmt": "p1",
                 "items": [{"name": "模型情境 B 基準", "v": f"@In!E{IN['shr_hum_B']}"}, {"name": "模型情境 A", "v": f"@In!E{IN['shr_hum_A']}", "ref": True},
                           {"name": "支持現價所需", "v": f"@Recon!B{RC['s3']}"}]}],
     "caption": "反解為整本重算二分搜尋（其他輸入不變）；灰色＝情境 A 的設定值，作為樂觀端參照。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": ["@README!A2", "@README!A7"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": ["@README!A3", "@README!A4", "@README!A5", "@README!A6"]},
    {"type": "h3", "text": "方法"},
    {"type": "note", "parts": ["@README!A8"]}]}

rows_main = [("dcf_tp", "n0"), ("cmp_tp", "n0"), ("cmpu_tp", "n0"), ("tp", "n0"), ("px", "n0"), ("up", "sp1"), ("rating", ""), ("rec", ""),
             ("dist", "p1"), ("up_b", "n0"), ("dn_b", "n0")]
valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": va("dcf_tp"), "value": v("dcf_tp"), "series": 1}, {"label": va("cmp_tp"), "value": v("cmp_tp"), "series": 2},
               {"label": va("tp"), "value": v("tp"), "series": 0}, {"label": "市場共識目標價（賣方平均，見對帳章）", "value": f"@Recon!C{RC['cons_tp']}", "series": 1}],
     "ref": {"label": va("px"), "value": v("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值", "說明"],
     "rows": [[va(k), v(k, f), f"@Val!E{VL[k]}"] for k, f in rows_main] + [[f"@Recon!A{RC['cons_tp']}", f"@Recon!C{RC['cons_tp']}|n0", "來源與區間見對帳頁（只作對照，不進模型）"]]},
    {"type": "table", "caption": "收支關卡：FY2026–FY2030 FCFF 扣股利合計（十億日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C", "說明"],
     "rows": [[va("gate"), v("gate", "n1"), v("gate", "n1", "C"), v("gate", "n1", "D"), f"@Val!E{VL['gate']}"]]},
    {"type": "h3", "text": "折現率"},
    {"type": "kv", "items": [[f"@In!B{IN['rf']}", f"@In!E{IN['rf']}|p2"], [f"@In!B{IN['beta']}", f"@In!E{IN['beta']}|n2"], [f"@In!B{IN['erp']}", f"@In!E{IN['erp']}|p1"],
                             [f"@In!B{IN['crp']}", f"@In!E{IN['crp']}|p2"]] + [[va(k), v(k, "p2")] for k in ("ke", "kd_at", "wacc", "ronic", "g_show")]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[va(k)] + [v(k, f, c) for c in "BCD"] for k, f in (("ev", "n1"), ("tvs", "p1"), ("dcf", "n0"))]},
    {"type": "table", "collapse": True, "summary": "展開：股權市值、負債與 DCF 明細",
     "head": ["項目", "數值"], "rows": [[va(k), v(k, "n1")] for k in ("e_mv", "d_bv")]},
    {"type": "table", "caption": "Comparable（同業前瞻 P/E × 模型 FY2027 EPS）與情境加權", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[va(k)] + [v(k, f, c) for c in "BCD"] for k, f in (("eps27", "n1"), ("cmpu", "n0"), ("cmp", "n0"), ("wt", "n0"), ("w", "p0"))]},
    {"type": "table", "caption": "上檔空間三段歸因與相對倍數", "head": ["項目", "數值", "說明"],
     "rows": [[va(k), v(k, f), f"@Val!E{VL[k]}"] for k, f in (("att_d", "sp1"), ("att_c", "sp1"), ("att_r", "sp1"), ("relnow", "n2"), ("relreq", "n2"), ("relflip", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數（Comp 頁）",
     "head": ["@Comp!A4", "@Comp!B4", "@Comp!C4", "@Comp!D4", "@Comp!F4", "@Comp!G4", "@Comp!H4", "@Comp!J4", "@Comp!L4", "@Comp!M4", "@Comp!O4"],
     "rowgen": {"sheet": "Comp", "rows": "5-10", "cols": ["A", "B", "C", "D", "F", "G", "H", "J", "L", "M", "O"],
                "fmts": ["", "", "n2", "", "n2", "n2", "n2", "", "n2", "", "n2"]}},
    {"type": "table", "collapse": True, "summary": "展開：同業中位數與本公司現行倍數",
     "head": ["項目", "數值"], "rows": [["@Comp!A12", "@Comp!F12|n2"], ["@Comp!A13", "@Comp!G13|n2"], ["@Comp!A14", "@Comp!F14|n2"], ["@Comp!A15", "@Comp!F15|n2"], ["@Comp!A16", "@Comp!F16|n2"]]},
    {"type": "note", "parts": ["@Comp!A1"]},
    {"type": "table", "caption": "期間", "head": ["項目", "數值"], "rows": [[va("t_model"), v("t_model", "n2")], [va("t_hold"), v("t_hold", "n2")], [va("tp_date"), v("tp_date")]]}]}

revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": [f"@Recon!A{RC['solve_h']}"]},
    {"type": "table", "caption": "整本重算反解：支持現價與評等升一級所需的主導變數", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Recon!A{RC['s1']}", f"@Recon!B{RC['s1']}|n0", f"@Recon!E{RC['s1']}"], [f"@Recon!A{RC['s2']}", f"@Recon!B{RC['s2']}|n0", f"@Recon!E{RC['s2']}"],
              [f"@Recon!A{RC['s3']}", f"@Recon!B{RC['s3']}|p1", f"@Recon!E{RC['s3']}"]]},
    {"type": "table", "caption": "同業倍數法：現價隱含的相對倍數", "head": ["項目", "數值"],
     "rows": [["@Comp!A14", "@Comp!F14|n2"], ["@Comp!A12", "@Comp!F12|n2"], [va("relnow"), v("relnow", "n2")], [va("relreq"), v("relreq", "n2")], [va("relflip"), v("relflip", "n2")]]}]}

pl_keys = [("rev", "n2"), ("jp", "n2"), ("cn", "n2"), ("na", "n2"), ("eu", "n2"), ("op", "n2"), ("opm", "p1"), ("nonop", "n2"), ("ord", "n2"), ("ni", "n2"), ("eps", "n1"),
           ("da", "n2"), ("cx", "n2"), ("dnwc", "n2"), ("fcff", "n2"), ("div", "n2"), ("gate", "n2")]


def met(name, key, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": {s: f"@PL!E{PR[s][key]}:K{PR[s][key]}" for s in "ABC"}}
    if chart:
        m["chart"] = True
    return m


pl = {"id": "pl", "title": "損益", "blocks": [
    {"type": "note", "parts": ["損益採日本會計基準（J-GAAP）口徑：公司決算短信與指引皆為此口徑。模型只建營收、營業利益、業外、經常利益與淨利，未建營業成本、毛利、營業費用、稅前利益與所得稅列（版面規格的標準損益表列未補：不為版面改模型）。"]},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "theme", "title": "營收（十億日圓，依用途堆疊）與營業利益率：FY2026–FY2032（情境 B）", "x": "@PL!E3:K3", "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": f"@PL!E{PB['rev']}:K{PB['rev']}",
     "bars": [{"name": [f"@PL!A{PB[k]}"], "y": f"@PL!E{PB[k]}:K{PB[k]}", "color": c} for k, c in (("rob", "#0b7a80"), ("semi", "#6fc7cb"), ("hum", "#c29a2b"), ("oth", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{PB['opm']}"], "y": f"@PL!E{PB['opm']}:K{PB['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝命題模組的四個用途（非官方分部，合計等於合併營收）；折線＝營業利益率（右軸），以公司指引為錨、增量利益率外推。"},
    {"type": "table", "caption": "損益表（J-GAAP；十億日圓，EPS 為日圓）：FY2024、FY2025 實績，FY2026 起情境 B", "head": ["@PL!A3"] + [f"@PL!{c}3" for c in PLC],
     "rows": [[("營收（合併營收）" if k == "rev" else f"@PL!A{PB[k]}")] + [f"@PL!{c}{PB[k]}|{f}" for c in PLC] for k, f in pl_keys]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": "@PL!E3:K3",
     "metrics": [met("合併營收（十億日圓）", "rev", "n1", True), met("營業利益（十億日圓）", "op", "n1", True), met("營業利益率", "opm", "p1"),
                 met("EPS（日圓）", "eps", "n1", True), met("FCFF 扣股利（十億日圓）", "gate", "n2")]},
    {"type": "table", "collapse": True, "summary": "展開：分部與用途合計核對（應為零）", "head": ["項目"] + [f"@PL!{c}3" for c in PLC],
     "rows": [[f"@PL!A{PR[s][k]}"] + [f"@PL!{c}{PR[s][k]}|n3" for c in PLC] for s in "BAC" for k in ("chk", "chk2")]}]}

TB = TR["B"]
THC = "DEFGHIJK"
in_grp = [
    ("一、產業機器人（含協作）", ["mix_rob", "ifr_25", "ifr_26", "ifr_29"]),
    ("二、半導體設備與其他用途", ["mix_semi"]),
    ("三、人形機器人", ["hum_26", "hu_25", "n_harm", "asp_26", "asp_lr"]),
    ("四、官方分部（地區）占比", ["cn_mix", "na_mix", "eu_mix"]),
    ("五、獲利與現金流", ["inc_m", "nonop", "tax", "da_26", "da_ss", "cx_r", "nwc_r"]),
    ("六、折現率與估值", ["rf", "crp", "erp", "b_own", "b_peer", "beta", "kd", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "tv_alt", "pe_peer", "rel_k", "w_dcf"]),
]


def in_row(k):
    r = IN[k]
    u = wb["In"][f"C{r}"].value
    f = "p1" if u in ("%", "%／年") else "n2"
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!G{r}|tag"]


rng = [(g, [k for k in ks if wb["In"][f"D{IN[k]}"].value is not None or wb["In"][f"F{IN[k]}"].value is not None]) for g, ks in in_grp]
fix = [(g, [k for k in ks if k not in dict(rng)[g]]) for g, ks in in_grp]
used = {k for _, ks in in_grp for k in ks}
scn_keys = {k for k in IN if k[-2:] in ("_A", "_B", "_C") or k in ("tp_wA", "tp_wB", "tp_wC")}
mkt = {"px_date", "px", "val_date", "tp_date", "bs_date", "stub_frac", "frac_0", "sh_v", "cash_v", "invs_v", "debt_v", "lease_v", "int_dps", "stub_div", "dps_tgt",
       "g_rev26", "g_op26", "g_ord26", "g_ni26", "payout"}
missing = [k for k in IN if k not in used | scn_keys | mkt]
assert not missing, f"In 參數未歸組：{missing}"

drivers = {"id": "drivers", "title": "驅動因子 → 推導量", "blocks": [
    {"type": "theme", "title": "人形／AI 機器人營收（十億日圓）× 情境 A／B／C：FY2026–FY2032", "x": "@PL!E3:K3", "fmt": "n1", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：人形營收"], "y": f"@Thesis!E{TR[s]['hum']}:K{TR[s]['hum']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "lines": [{"name": ["情境 B：人形占合併營收（右軸）"], "y": f"@Thesis!E{TB['humsh']}:K{TB['humsh']}", "axis": "r", "color": "#c29a2b", "labels": True}],
     "fmt_r": "p1",
     "caption": "人形營收＝FY 加權全球人形出貨 × 每台諧波數 × HDS 份額 × 單價（單價逐年下降）；三情境在出貨與份額分化。"},
    {"type": "note", "parts": ["@Thesis!A1"]},
    {"type": "table", "caption": "命題模組推導（情境 B；十億日圓，除另註）", "head": ["項目"] + [f"@Thesis!{c}3" for c in THC],
     "rows": [[f"@Thesis!A{TB[k]}"] + [f"@Thesis!{c}{TB[k]}|{'p1' if k == 'humsh' else 'n1' if k in ('inst', 'instfy', 'hu', 'hufy', 'asp') else 'n2'}" for c in THC]
              for k in ("inst", "instfy", "rob", "semi", "hu", "hufy", "asp", "hum", "oth", "tot", "humsh")]},
    {"type": "table", "caption": "官方分部推導（情境 B）：合併營收 × 地區占比（FY2026 第一季實績占比固定）", "head": ["@PL!A3"] + [f"@PL!{c}3" for c in PLC],
     "rows": [[f"@PL!A{PB[k]}"] + [f"@PL!{c}{PB[k]}|n2" for c in PLC] for k in ("rev", "jp", "cn", "na", "eu")]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照）", "head": ["項目", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）"],
     "rows": [[f"@In!B{r}"] + [f"@In!{c}{r}|{'n0' if wb['In'][f'A{r}'].value == 'hu32' else 'p1'}" for c in "DEF"] for r in range(SCN0, SCN_END + 1)]},
    {"type": "table", "caption": "單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [x for g, ks in rng if ks for x in ([{"sub": g}] + [in_row(k) for k in ks])]},
    {"type": "table", "collapse": True, "summary": "展開：固定輸入（實績、公司揭露、換算值）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [x for g, ks in fix if ks for x in ([{"sub": g}] + [in_row(k) for k in ks])]
             + [{"sub": "七、市場資料、日期與公司指引"}] + [in_row(k) for k in IN if k in mkt and k not in ("px_date", "val_date", "tp_date", "bs_date")]}]}

recon = {"id": "recon", "title": "對帳：公司指引、受注、共識與競爭者", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "模型 vs 公司指引、受注與賣方共識", "head": ["@Recon!A3", "@Recon!B3", "@Recon!C3", "@Recon!D3", "@Recon!E3"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|n2", f"@Recon!C{r}|n2", f"@Recon!D{r}|n2", f"@Recon!E{r}"] for r in range(4, RC["cons_tp"])]
             + [[f"@Recon!A{RC['cons_tp']}", f"@Recon!B{RC['cons_tp']}|n0", f"@Recon!C{RC['cons_tp']}|n0", f"@Recon!D{RC['cons_tp']}|n0", "TipRanks 賣方平均（近三個月）；只作對照，不進模型；網址見 xlsx 對帳頁"]]},
    {"type": "table", "caption": "競爭者對照：綠的諧波", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|{'p1' if r == RC['comp_end'] - 1 else 'n1' if r == RC['comp_end'] else 'n2'}",
               (f"@Recon!E{r}" if r != RC["comp_h"] + 2 else "公司下半年匯率假設（業績予想修正公告，網址見 xlsx 對帳頁）")] for r in range(RC["comp_h"] + 1, RC["comp_end"] + 1)]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（對帳頁）", "head": ["@對帳!A3", "@對帳!B3", "@對帳!C3", "@對帳!D3", "@對帳!G3", "@對帳!H3"],
     "rowgen": {"sheet": "對帳", "rows": f"4-{wb['對帳'].max_row}", "cols": ["A", "B", "C", "D", "G", "H"], "fmts": ["", "", "n2", "", "tag", ""]}}]}

status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "限制"},
    {"type": "list", "items": ["@README!A9", "@Comp!A1", f"@In!J{IN['rf']}"]},
    {"type": "table", "caption": "風險登錄", "head": ["@Risk!A3", "@Risk!B3", "@Risk!C3", "@Risk!D3"], "rowgen": {"sheet": "Risk", "rows": f"4-{wb['Risk'].max_row}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為未取得的項目", "sheets": {"In": ["B"], "Q": ["A"]}, "empty": "沒有標為未取得的項目。"}]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", os.path.basename(XS[0]))
