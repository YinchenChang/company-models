# 聯電：把 html_src.json（v1.0 之前的頁面設定，內容不動）依「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）轉成 html.json。
# 用法：python3 umc/gen_html_json.py <成品 xlsx>（讀 umc/html_src.json，寫 umc/html.json；勿手改 html.json）
import copy, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
import relayout as RL  # noqa: E402

X = sys.argv[1]
WB = openpyxl.load_workbook(X, data_only=True)
cfg = copy.deepcopy(json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8")))
SEC = {s["id"]: s for s in cfg["sections"]}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, need_num=False, exact=False):
    ws = WB[sh]
    for r in range(1, ws.max_row + 1):
        t = str(ws.cell(r, 1).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


I = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_tok(prefix):
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    raise KeyError(prefix)


# ─── 驅動因子：舊「輸入 → 推導量」表移到附錄（轉換器依 caption「驅動 → 推導」辨認） ───
for b in SEC["drivers"]["blocks"]:
    if b["type"] == "table" and not b.get("caption") and b.get("head", [None])[0] == "輸入":
        b["caption"] = "驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"

# ─── 估值：長條圖 → 結論主要項目 → 投資建議 → DCF 折現率 → DCF 彙總 → Comparable → 其餘摺疊 ───
V = lambda k: row_of("Val", k, need_num=True)
tpw, dcft, comp, compadj, cons = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（市場主流同業法）"), V("Comparable 目標價（相對倍數調整後）"), V("市場共識目標價")
upw, upc, dcf12, px = V("上檔空間：加權目標價"), V("上檔空間：Comparable"), V("DCF 情境價值推到 12 個月後"), V("現價")
rec, grade = row_of("Val", "投資建議", exact=True), row_of("Val", "上檔空間級距")
rk = row_of("Val", "股權成本 ke")                       # v0.2：0.1 節 ke、kd_at、w_e、WACC、g、RONIC，接著 DCF 彙總 10 列
r_pvx, r_dcf = row_of("Val", "明確期 FCFF 現值"), row_of("Val", "DCF 目標價（情境加權）")   # 第二輪修正：依標籤定位（模組 v2 stub 期多兩列）
assert r_pvx == rk + 6, (r_pvx, rk)
_fm = lambda lab: "n3" if lab.startswith("評價用") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1"
fb = next(r for r in range(row_of("DCF", "情境 B（基準）"), WB["DCF"].max_row + 1) if WB["DCF"].cell(r, 1).value == "營收")
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
old = SEC["valuation"]["blocks"]
pick = lambda f: next(b for b in old if f(b))
bars = pick(lambda b: b["type"] == "bars")
for _it in bars["items"]:   # v0.2：0.1 節改版後 DCF 目標價列移到 dcft
    if _it["value"] == "@Val!B24": _it.update(label=f"@Val!A{dcft}", value=f"@Val!B{dcft}")
dcf_tbl = pick(lambda b: b.get("caption") == "@Val!A5")
peer_tbl = pick(lambda b: b.get("caption") == "@Val!A26")
comp_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == "@Val!A33")
px_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == f"@Val!A{px}")
rm_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0].startswith("@Val!A14"))
rule_tbls = [b for b in old if b["type"] == "table" and not b.get("caption") and not b.get("collapse")]
hist_tbl = pick(lambda b: str(b.get("summary", "")).startswith("展開：十年年底 P/E"))
notes = {b["parts"][0]: b for b in old if b["type"] == "note"}
H5 = ["項目", "數值", "", "", "說明"]
inr = lambda k, f: [f"@In!B{I[k]}", f"@In!E{I[k]}|{f}"]
kv2tbl = lambda kv, cap, **kw: dict({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [[k, v] for k, v in kv["items"]]}, **kw)
SEC["valuation"]["blocks"] = [
 bars,
 {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
   [f"@Val!A{tpw}", f"@Val!B{tpw}|n2", "", "", f"@Val!E{tpw}"],
   [f"@Val!A{dcft}", f"@Val!B{dcft}|n2", "", "", f"@Val!E{dcft}"],
   [f"@Val!A{compadj}", f"@Val!B{compadj}|n2", "", "", f"@Val!E{compadj}"],
   [f"@Val!A{comp}", f"@Val!B{comp}|n2", "", "", f"@Val!E{comp}"],
   [f"@Val!A{px}", f"@Val!B{px}|n2", "", "", ""],
   [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"],
   [f"@Val!A{upc}", f"@Val!B{upc}|sp1", "", "", f"@Val!E{upc}"],
   ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""],
   [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
   [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n1", f"@Val!C{dcf12}|n1", f"@Val!D{dcf12}|n1", ""],
   [f"@Val!A{cons}", f"@Val!B{cons}|n2", "", "", f"@Val!E{cons}"]]},
 {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（與共識的差距拆解、價格帶、十年分位、上檔空間級距）", "head": H5, "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}" if r == c0 + 10 else "", f"@Val!D{r}|{f}" if r == c0 + 10 else "", f"@Val!E{r}"]
   for r, f in ((c0 + 7, "sp1"), (c0 + 8, "sp1"), (c0 + 9, ""), (c0 + 10, "p0"))] + [[f"@Val!A{grade}", f"@Val!B{grade}", "", "", f"@Val!E{grade}"]]},
 {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != r_dcf else "", f"@Val!D{r}|{fm}" if r != r_dcf else "", f"@Val!E{r}"]
   for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, r_dcf + 1))]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B（基準）FCFF 逐年明細（DCF 頁；十億新台幣）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
  "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
             "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
 kv2tbl(comp_kv, "Comparable（同業前瞻 P/E 中位數 × 一年後的前瞻 EPS）"),
 dict(peer_tbl, collapse=True, summary="展開：同業前瞻倍數"),
 kv2tbl(rm_kv, "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
 notes["@Val!A48"],
 kv2tbl(px_kv, "價格與市場隱含倍數", collapse=True, summary="展開：現價與市場隱含倍數"),
 {"type": "h3", "text": "十年 P/E 規則值（舊法參考，不進目標價）"}] + [dict(b, collapse=True, summary=s) for b, s in zip(rule_tbls, ("展開：規則值隱含股價（FY2027 EPS × 依循環位置的倍數）", "展開：循環位置（本模型營收最高年）"))] + [
 notes["@Val!A2"], hist_tbl, notes["@Val!A101"]]

SEC["recon"]["blocks"][0]["caption"] = "本模型與公司說法、共識、券商的對照"

# 敏感度：第二張龍捲風是 FY2028 EPS（情境 A），不是加權目標價，轉換器之後還原標題
TOR = {(b["sheet"], b["value_col"]): b["title"] for b in SEC["sens"]["blocks"] if b["type"] == "tornado"}
# v0.2：舊折現率三案改為標準版 DCF 參數案例（同一張表），案例列與讀法列位置改變
_sl = next(r for r in range(5, 80) if WB["Sens"].cell(r, 1).value is None) - 1          # 最後一個案例列
_sn = row_of("Sens", "讀法：")

for b in SEC["sens"]["blocks"]:
    if b["type"] == "tornado": b["rows"] = f"5-{_sl}"
    if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Sens": b["rowgen"]["rows"] = f"4-{_sl}"
    if b["type"] == "note" and b["parts"][0].startswith("@Sens!A") and b["parts"][0] != "@Sens!A2": b["parts"] = [f"@Sens!A{_sn}"]
pass


def scn_table_fix(sec):
    """第三輪修正第 32 點：情境驅動表依 In 頁「情境驅動對照」節標題找列範圍（不寫死列號）；百分比列以 % 顯示。"""
    ws = WB["In"]
    t = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("情境驅動對照"))
    h = t + 1; rows = []; r = h + 1
    while ws.cell(r, 1).value not in (None, ""):
        rows.append(r); r += 1
    assert ws.cell(rows[-1], 1).value == "scn_w", ws.cell(rows[-1], 1).value
    def fm(rr):
        u = str(ws.cell(rr, 3).value or "")
        vals = [ws.cell(rr, c).value for c in (4, 5, 6) if isinstance(ws.cell(rr, c).value, (int, float))]
        if u in ("%", "pt") or "%" in u: return "p1"
        return "n0" if vals and max(abs(v) for v in vals) >= 100 else "n2"
    for b in sec["blocks"]:
        if b.get("type") == "table" and str(b.get("caption", "")).startswith("情境驅動"):
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{rr}", f"@In!D{rr}|{fm(rr)}", f"@In!E{rr}|{fm(rr)}", f"@In!F{rr}|{fm(rr)}", f"@In!C{rr}"] for rr in rows]
            b.pop("rowgen", None)
            return
    raise KeyError("情境驅動表")


scn_table_fix(SEC["drivers"])

SPEC = {
 "in_skip": [in_tok("G."), in_tok("情境驅動對照")],
 "in_groups": [
  ["成熟製程晶圓（產能 × 稼動率 × 平均售價）", [
    ["產能（十二吋約當）", ["cap_26", "capg27", "capg28", "capg29", "cap_adj"]],
    ["平均稼動率（三情境）", [k for k in I if re.match(r"^ut2\d_[ABC]$", k)] + ["ut_adj"]],
    ["平均售價（美元，每片十二吋約當）", ["asp_26"] + [k for k in I if re.match(r"^asp2\d_[ABC]$", k)] + ["asp_adj"]]]],
  ["新業務（AI 相關：先進封裝矽中介層、矽光子；Intel 十二奈米合作）", [
    ["AI 相關營收", ["ai_26", "aig27_A", "aig27_B", "aig27_C", "aig_L", "ai_adj"]],
    ["AI 加速器顆數係數（Siliconomics，FY2028 起乘在 AI 相關營收）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["Intel 十二奈米營收", [k for k in I if re.match(r"^int(_27|2\d_[ABC])$", k)]],
    ["新業務毛利", ["m_new"]]]],
  ["全公司（匯率、成本、FY2026 錨點、股價與目標價規則）", [
    ["匯率", ["fx", "fx_26"]],
    ["成本：折舊、每片現金成本、營業費用", ["da_tot26", "da_sh", "dag_27", "dag_28", "dag_29", "dag_adj", "cc_g", "phi", "g_opx_min"]],
    ["營業外、稅與股數", ["nonop_f", "tax", "sh_g"]],
    ["FY2026 錨點（第四季指引或假設）", ["q4_ship", "q4_asp", "q4_ut", "q4_gm", "q4_opx", "q4_nonop", "q4_tax", "q_sh"]],
    ["股價與市場倍數", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim_5"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_u_med", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["Wattonomics 錨定（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"1. 成熟製程晶圓": "成熟製程晶圓（產能 × 稼動率 × 平均售價）", "@Units!A34": "新業務（AI 相關：先進封裝矽中介層、矽光子；Intel 十二奈米合作）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "Comparable 目標價（市場", "現價", "上檔空間：加權", "上檔空間：Comparable", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司說法、共識、券商的對照": ["FY2026 營收 vs 共識", "FY2026 EPS vs 共識", "FY2027 平均售價年增", "FY2029 AI 相關營收", "共識前瞻 12 月 EPS"]},
 "n_recon": 4,
}

out = RL.relayout(cfg, X, SPEC)

# 補充規則 A：橫幅中間格＝只看 Comparable（市場主流同業法，未經相對倍數調整）＋自己的上檔空間；附註寫進加權的調整後值
h = out["hero"]
h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp}|n2", ["上檔空間 ", f"@Val!B{upc}|sp1"]]
if abs(WB["Val"].cell(compadj, 2).value - WB["Val"].cell(comp, 2).value) > 1e-9:
    h["note"] = (h.get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{compadj}|n2", " 元。"]

for s in out["sections"]:
    for b in s["blocks"]:
        if b["type"] == "tornado":
            t0 = TOR.get((b["sheet"], b["value_col"]))
            if t0 and not t0.startswith("加權目標價"):
                b["title"] = t0 + ("（未達兩位數百分比的案例）" if b.get("collapse") else "（變動達兩位數百分比的案例）")
    if s["id"] == "status":   # 附錄「版本要點」只收版本說明列
        for b in s["blocks"]:
            if b.get("summary") == "展開：本版與前版的修改要點":
                ws = WB["README"]
                b["rows"] = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1) if re.match(r"^v\d+\.\d", str(ws.cell(r, 1).value or ""))]

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(out, X)
cfg = out


# ================= v0.2.4（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 台積電 v0.5.5、NVIDIA v0.3.10；docs/plan/個股HTML版面規格.md「v1.3 規格」）＋三層協調 C0 =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含三層協調 C0）→ 資料狀態；聯電無個別模組。
# 數字與白話文字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表與分部營收推導、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁；umc/v024_post.py 產生）。列號依標籤找；本段文字不含年份與日期以外的數字。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RVW, CW, SEW = WB["README"], WB["PL"], WB["RevDCF"], WB["Coord"], WB["Sens"]
    r13 = row_of("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.2.4"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    # ---- 橫幅：白話與對帳中揭露；規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    assert old_note[-2] == " " and str(old_note[-1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[-2:]
    for t in cfg["hero"]["tiles"]:
        t[0] = [t[0], "（新台幣）"] if isinstance(t[0], str) else list(t[0]) + ["（新台幣）"]
    # ---- 1. 命題：主圖＝分部營收推導（情境 B；成熟製程晶圓、AI 相關、Intel 十二奈米堆疊）----
    sg = row_of("PL", "分部營收推導（情境 B")
    hS = sg + 1
    ex = lambda label, start=sg: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == label)
    r_cap, r_ut, r_ship, r_asp, r_w, r_ai, r_int, r_rs, r_sh = (ex(t) for t in ("產能（12 吋約當）", "平均稼動率", "晶圓出貨（12 吋約當）", "平均售價（成熟製程晶圓）", "成熟製程晶圓營收", "AI 相關營收",
                                                                               "Intel 12 奈米營收", "營收", "AI 相關與 Intel 12 奈米 ÷ 營收"))
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B，十億新台幣）與 AI 相關、Intel 十二奈米占比：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.35,
              "total": f"@PL!F{r_rs}:I{r_rs}",
              "bars": [{"name": [f"@PL!A{r_w}"], "y": f"@PL!F{r_w}:I{r_w}", "color": "#6fc7cb"}, {"name": [f"@PL!A{r_ai}"], "y": f"@PL!F{r_ai}:I{r_ai}", "color": "#0b7a80"},
                       {"name": [f"@PL!A{r_int}"], "y": f"@PL!F{r_int}:I{r_int}", "color": "#c29a2b"}],
              "lines": [{"name": [f"@PL!A{r_sh}"], "y": f"@PL!F{r_sh}:I{r_sh}", "axis": "r", "color": "#3b3838", "labels": True}],
              "caption": "堆疊＝情境 B（基準）的成熟製程晶圓、AI 相關（先進封裝矽中介層、矽光子）與 Intel 十二奈米營收；頂端數字＝營收；折線＝AI 相關與 Intel 十二奈米占營收（右軸）。FY2026 為前三季實績加第四季指引的錨點。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    S["thesis"]["blocks"] = [chart1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}] + th
    # ---- 2. 估值：四個目標價圖＋表置頂（G）；章說明、評等規則、對帳中兩讀數區其後 ----
    va = S["valuation"]["blocks"]
    g0 = row_of("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣；豎線＝現價）", "fmt": "n1",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
    rules = {"type": "note", "parts": old_note[:-2]}
    h_i = next(i for i, b in enumerate(va) if b.get("type") == "h3" and str(b.get("text", "")).startswith("對帳中"))
    coord_blocks = va[h_i:h_i + 4]
    assert [b["type"] for b in coord_blocks] == ["h3", "note", "table", "bars"], [b["type"] for b in coord_blocks]
    old_bars = next(b for b in va if b.get("type") == "bars" and b.get("title") == "目標價與現價")
    rest = [b for b in va if b not in coord_blocks and b is not old_bars]
    top = [bars4]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    S["valuation"]["blocks"] = top + [tbl4, rules] + coord_blocks + rest
    # ---- 3. 敏感度：加權目標價相對基準的低案、高案主圖置頂（H）----
    t0 = row_of("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SEW.cell(r2 + 1, 1).value and not str(SEW.cell(r2 + 1, 1).value).startswith("核對"):
        r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and abs(SEW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SEW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append((dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))),
                    dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                         summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）")))
    all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案與設定）",
               "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCD"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}"] for r in range(r1, r2 + 1)]}
    others = []
    for b in S["sens"]["blocks"]:
        if b.get("type") == "tornado":
            if b.get("value_col") == "J":   # 舊「加權目標價相對基準」龍捲風（同一組案例），由上方主圖取代
                continue
            b = dict(b, collapse=True, summary=("展開：對照：" + str(b.get("title") or "EPS 敏感度")))
        elif b.get("type") == "table" and not b.get("collapse"):
            b = dict(b, collapse=True, summary="展開：" + (str(b.get("caption") or "表")[:40]))
        others.append(b)
    intro = S["sens"].pop("intro", None)
    lead = [tor[0][0], tor[1][0], {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入，三情境同時改、整本逐案重算；變動達兩位數百分比者直接顯示，其餘摺疊）。業務案例取 Sens 頁逐案推導區塊，目標價參數取 Sens 頁 DCF 參數案例。下方摺疊為全部案例與 EPS 敏感度對照。"]}]
    if intro: lead.append({"type": "note", "parts": [intro], "collapse": True, "summary": "展開：敏感度頁的說明"})
    S["sens"]["blocks"] = lead + [tor[0][1], tor[1][1], all_tbl] + others
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RVW.max_row + 1) if str(RVW.cell(r, 1).value or "").startswith(p))
    ra, rb_, rbb, rc, rcmp, rsum, yh = rv("A. DCF"), rv("(a)"), rv("(b)"), rv("(c)"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0_ = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收（", rbb), rv("模型：FY2028 營收（", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
    kvr = [rv(p, ra) for p in ("現價（", "支持現價所需的評價日", "淨現金滾到評價日", "評價用股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日")]
    kvf = ["n2", "n2", "n1", "n2", "n1", "n1", "n2", "n2"]
    cm = [rv(p, rcmp) for p in ("12 個月後前瞻 EPS", "只看 Comparable：支持現價所需的前瞻", "只看 Comparable：所需的相對", "加權（DCF 不變）", "同業前瞻本益比中位數")]
    yrr = [rv(p, yh) for p in ("營收（", "營業利益 EBIT", "營業利益率", "稅率", "NOPAT＋折舊", "營運資金增加", "現值（")]
    yfm = ["n1", "n1", "p1", "p1", "n1", "n1", "n1"]
    S["revdcf"] = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億新台幣）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(kvr, kvf)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{cm[0]}", "—", f"@RevDCF!C{cm[0]}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(cm[1:4])]
                     + [[f"@RevDCF!A{cm[4]}", "—", f"@RevDCF!C{cm[4]}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "collapse": True, "summary": "展開：情境 B 各年營收、營業利益與現值（反解的計算輸入；十億新台幣）",
             "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"],
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrr, yfm)]}]}
    # ---- 5. 損益（標準損益表）----
    p0 = row_of("PL", "標準損益表（v0.2.4")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含其他營業損益）", "營業利益", "營業利益率", "其他收入（淨額；營業外）", "稅前利益", "所得稅",
            "非控制權益損益", "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股；基本加權平均）", "EPS（新台幣；基本，歷史為公司公告值）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "十億股": "n2"}.get(str(u).strip(), "n1")
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC]
    is_tbl = {"type": "table", "caption": "標準損益表（十億新台幣；EPS 為新台幣）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）；TIFRS 合併（GAAP 口徑；實績三年與 FY2022 基期皆取公司年度合併財報〔台灣版 IFRS〕，同一準則，不混用美國申報的 IASB 版 IFRS）", "head": head,
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": head, "collapse": True, "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與 DCF 頁、FY2030E–FY2032E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億新台幣）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸；TIFRS 合併，淨利為歸屬母公司）。FY2026E 為前三季實績加第四季指引的錨點，含第二季一次性投資收益；FY2030E–FY2032E 為收斂法延伸期（營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    legacy = []
    for b in old_pl:
        if b["type"] == "pl":
            legacy.append(dict(b, collapse=True, summary="展開：三情境損益摘要與折線"))
        elif b["type"] in ("note", "kv", "table"):
            legacy.append(dict(b, collapse=True, summary="展開：" + ("延伸期（FY2030 起）說明" if b["type"] == "note" else "三情境錨點核對")))
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + legacy)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    sf = {"千片": "n0", "%": "p1", "US$/片": "n0", "NT$/US$": "n2", "NT$B": "n1", "US$B": "n3"}
    r_chk2 = ex("核對：晶圓＋AI 相關＋Intel 12 奈米 − 營收（應為 0）")
    seg_rows = [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{'n3' if r == r_chk2 else sf.get(str(PLW.cell(r, 2).value), 'n2')}" for c in "FGHI"] for r in range(r_cap, r_sh + 1)]
    seg_tbl = {"type": "table", "caption": f"@PL!A{sg}", "head": [f"@PL!A{hS}", f"@PL!B{hS}"] + [f"@PL!{c}{hS}" for c in "FGHI"], "rows": seg_rows}
    chart2 = {"type": "theme", "title": "成熟製程晶圓出貨（情境 B，千片十二吋約當）與平均售價（美元／片，右軸）：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.3,
              "bars": [{"name": [f"@PL!A{r_ship}"], "y": f"@PL!F{r_ship}:I{r_ship}", "color": "#6fc7cb"}],
              "lines": [{"name": [f"@PL!A{r_asp}"], "y": f"@PL!F{r_asp}:I{r_asp}", "axis": "r", "color": "#0b7a80", "labels": True}],
              "caption": "長條＝晶圓出貨（產能 × 稼動率；產能依公司已核准擴產）；折線＝成熟製程晶圓平均售價（美元，FY2026 為校準值）。稼動率與售價是公司層級假設，Siliconomics 尚無成熟製程行業層。"}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, chart2, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 7. 對帳：三層協調 C0（Coord 頁）----
    c1, c2, c3 = cr("1. 讀哪一層"), cr("2. C0 讀取層級核對"), cr("3. 對帳中年份")
    h1 = c1 + 1; h2 = next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = cr("對帳中項目數（讀取層級核對")
    ucf = lambda r: {"%": "p1", "x": "n3", "NT$/US$": "n2"}.get(str(CW.cell(r, 2).value), "n1")
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上方公司實績表）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）",
              "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ucf(r)}", f"@Coord!E{r}|{ucf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]}
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(CW.cell(r, 2).value, (int, float)) and not isinstance(CW.cell(r, 2).value, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
                       for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = cr("層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    a4 = cr("4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    jt = cr("C0（實績）判定", ih); jr0 = cr("C0（實績）判定：合併營收"); a4e = cr("對帳中項目數（C0 實績表", jr0)
    ifm = lambda r: {"%": "p1", "%/年": "p1", "千片": "n0", "US$B": "n2"}.get(str(CW.cell(r, 2).value), "n3")
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ifm(r)}", f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, jt) if CW.cell(r, 1).value]
    jr_rows = [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(jr0, a4e)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
                  {"type": "table", "caption": "公司實績（月營收、季營收、出貨、產能利用率、資本支出）與參考值：期間、等級與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
                  {"type": "table", "caption": "公司實績對層級的可比性判定", "head": ["項目", "判定", "結論與缺漏揭露"], "rows": jr_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e + 1}"]}, {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "本模型與公司說法、共識、券商的對照"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與共識對照"
    # ---- 8. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]


_v13()
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
