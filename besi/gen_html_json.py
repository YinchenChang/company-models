# BESI：把 html_src.json（v1.0 之前的頁面設定，內容不動）依「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）轉成 html.json。
# 用法：python3 besi/gen_html_json.py <成品 xlsx>（讀 besi/html_src.json，寫 besi/html.json；勿手改 html.json）
# ─── 共通段（各公司 gen_html_json.py 相同；只依賴 tools/company_html/relayout.py，不修改它） ───
import copy, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
import relayout as RL  # noqa: E402

X = sys.argv[1]
WB = openpyxl.load_workbook(X, data_only=True)
SRC = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
cfg = copy.deepcopy(SRC)
SEC = {s["id"]: s for s in cfg["sections"]}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, col=1, need_num=False, exact=False, start=1):
    """第一個 A 欄（或 col 欄）以 prefix 開頭的列；need_num＝B 欄須為數字。"""
    ws = WB[sh]
    for r in range(start, ws.max_row + 1):
        t = str(ws.cell(r, col).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


IN_ROW = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_skip_token(prefix="情境驅動對照"):
    """In 頁「情境驅動對照」節標題在轉換器內的節代號（依 relayout._check_in_coverage 的切法）。"""
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    return None


def make_hero(px_ref=None):
    """投資建議橫幅：加權目標價與「加權目標價實際使用的 Comparable」（相對倍數調整後）各附上檔空間。"""
    tp = row_of("Val", "目標價（DCF 與 Comparable 加權）", need_num=True)
    up = row_of("Val", "上檔空間：加權目標價", need_num=True)
    rec = row_of("Val", "投資建議", exact=True)
    try:
        tpc = row_of("Val", "Comparable 目標價（相對倍數調整後）", need_num=True)
        upc_ref = f"@README!B{row_of('README', '展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間')}"
    except KeyError:
        tpc = row_of("Val", "Comparable 目標價", need_num=True)
        upc_ref = f"@Val!B{row_of('Val', '上檔空間：Comparable', need_num=True)}"
    if px_ref is None:
        try:
            p = row_of("Val", "現價", need_num=True)
            px_ref = (f"@Val!A{p}", f"@Val!B{p}")
        except KeyError:
            p = IN_ROW["px"]
            px_ref = (f"@In!B{p}", f"@In!E{p}")
    h = {"rec": f"@Val!B{rec}",
         "tiles": [[f"@Val!A{tp}", f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                   [f"@Val!A{tpc}", f"@Val!B{tpc}|n2", ["上檔空間 ", upc_ref + "|sp1"]],
                   [px_ref[0], px_ref[1] + "|n2", ""]]}
    if WB["Val"].cell(rec, 5).value:
        h["note"] = ["投資建議依加權目標價判定：", f"@Val!E{rec}"]
    return h


def run(spec, hero=None):
    """套用轉換器；橫幅改用 make_hero()（轉換器以第一個「Comparable 目標價」列為準，與加權目標價實際使用的相對倍數調整後 Comparable 不同）。"""
    tor_titles = {(b.get("sheet"), b.get("value_col"), b.get("label_col")): b.get("title") for s in cfg["sections"] for b in s["blocks"] if b["type"] == "tornado"}
    RL._hero = lambda w: hero or make_hero()
    out = RL.relayout(cfg, X, spec)
    # 龍捲風標題：轉換器一律改寫為「加權目標價…」；原圖若不是加權目標價，還原原標題（避免把 EPS 敏感度寫成目標價敏感度）
    for s in out["sections"]:
        for b in s["blocks"]:
            if b["type"] != "tornado":
                continue
            t0 = tor_titles.get((b.get("sheet"), b.get("value_col"), b.get("label_col")))
            if t0 and not t0.startswith("加權目標價"):
                b["title"] = t0 + ("（未達兩位數百分比的案例）" if b.get("collapse") else "（變動達兩位數百分比的案例）")
    # 附錄「版本要點」只收版本說明列（v0.／日期開頭），不收舊的投資建議與重點觀察列
    st = next(s for s in out["sections"] if s["id"] == "status")
    for b in st["blocks"]:
        if b.get("summary") == "展開：本版與前版的修改要點":
            ws = WB["README"]
            b["rows"] = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1)
                         if re.match(r"^(v\d+\.\d|\d{4}-\d{2}-\d{2})", str(ws.cell(r, 1).value or ""))]
    return out


def hero_rule_a(out):
    """補充規則 A（2026-11-05 20:20）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）＋自己的上檔空間；
    有相對倍數調整且數值不同時，附註寫出進加權的調整後 Comparable。"""
    try:
        tpm = row_of("Val", "Comparable 目標價（市場主流同業法）", need_num=True)
    except KeyError:
        tpm = row_of("Val", "Comparable 目標價", need_num=True, exact=True)
    upm = row_of("Val", "上檔空間：Comparable", need_num=True)
    h = out["hero"]
    h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{tpm}|n2", ["上檔空間 ", f"@Val!B{upm}|sp1"]]
    try:
        adj = row_of("Val", "Comparable 目標價（相對倍數調整後）", need_num=True)
    except KeyError:
        adj = None
    if adj and abs(WB["Val"].cell(adj, 2).value - WB["Val"].cell(tpm, 2).value) > 1e-9:
        note = h.get("note") or []
        note = (note if isinstance(note, list) else [note]) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{adj}|n2", " 歐元。"]
        h["note"] = note
    return out


def dump(out):
    json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok")
# ─── 共通段結束 ───

NFMT = ["n1", "n1", "n1", "n1", "n2", "n2", "n2", "p1", "x1", "n2"]
SKIP_EXTRA = []
GROUPS = lambda SK: [
  ["HBM 混合鍵合（HBM 位元 → 晶粒 → 混合鍵合機台）", [
    ["HBM 出貨位元（markets/memory，唯讀）", SK["A"]],
    ["晶粒、堆疊層數與混合鍵合比例（與 Disco、韓美同一組假設）", SK["B"]],
    ["混合鍵合機台（單價、產能、份額）", SK["D"]]]],
  ["邏輯混合鍵合（SoIC 等；AI 先進封裝晶圓）", [
    ["AI 先進封裝晶圓需求與導入比例", SK["C"]]]],
  ["主流業務（傳統貼合機、封裝、電鍍、TC 鍵合）", [
    ["FY2026 錨點與混合鍵合營收拆分", SK["E"]],
    ["主流業務成長（隨設備市場）", SK["F"]]]]]

# ─── BESI／韓美共用：估值與敏感度區塊依標準版 DCF 改寫；驅動因子依 In 頁各節分塊 ───
I = IN_ROW
va = SEC["valuation"]
rk = row_of("Val", "股權成本 ke")
r_pvx, dcft = row_of("Val", "明確期 FCFF 現值"), row_of("Val", "DCF 目標價（情境加權）")   # 第二輪修正：依標籤定位（模組 v2 stub 期多兩列）
assert r_pvx == rk + 6, (r_pvx, rk)
_fm = lambda lab: "n3" if lab.startswith("評價用") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1"
fb = row_of("DCF", "營收", start=row_of("DCF", "情境 B"))
cap01 = f"@Val!A{rk - 2}"          # 0.1 節標題列（舊 DCF 表的 caption）
nb = []
for b in va["blocks"]:
    if b["type"] == "bars":
        b = copy.deepcopy(b)
        for it in b["items"]:
            if it["value"] == f"@Val!B{rk + 17}": it.update(label=f"@Val!A{dcft}", value=f"@Val!B{dcft}")
        nb.append(b); continue
    if b["type"] == "table" and b.get("caption") == cap01:
        nb += [{"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"],
                "rows": [[f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
               {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"],
                "rows": [[f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != dcft else "", f"@Val!D{r}|{fm}" if r != dcft else "", f"@Val!E{r}"]
                         for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, dcft + 1))]},
               {"type": "table", "collapse": True, "summary": "展開：情境 B FCFF 逐年明細（DCF 頁）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
                "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
                           "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}}]
        continue
    if b["type"] == "h3" and b.get("text") == "重點觀察": continue
    if b["type"] == "list" and str(b.get("items", [""])[0]).startswith("@README!A") and len(b["items"]) == 3: continue
    nb.append(b)
va["blocks"] = nb
# 敏感度：目標價參數表（tools/fcff_dcf.write_sens；8 欄）
_th = row_of("Sens", "目標價參數敏感度"); _te = row_of("Sens", "DCF 權重 0.3／0.7", start=_th)
for b in SEC["sens"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and b.get("caption") != "@Sens!A3":
        b.update(caption=f"@Sens!A{_th}", head=[f"@Sens!{c}{_th + 1}" for c in "ABCDEFG"],
                 rowgen={"sheet": "Sens", "rows": f"{_th + 2}-{_te}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]})
SEC["sens"]["blocks"].insert(0, {"type": "tornado", "title": "加權目標價相對基準的變動（DCF 參數案例）", "sheet": "Sens", "rows": f"{_th + 2}-{_te}", "label_col": "A", "value_col": "D", "fmt": "sp1"})
for b in SEC["sens"]["blocks"]:
    if b["type"] == "table" and b.get("head", [None])[0] == "@Sens!A3": b.update(collapse=True, summary="展開：全部案例（情境 A 的 FY2027、FY2028 營收與 EPS）")
# 驅動因子：「情境驅動」表依 In 頁「情境驅動對照」一節重建（html_src 的列號已過時）
_hs = next(r for r in range(4, WB["In"].max_row + 1) if str(WB["In"].cell(r, 1).value or "").startswith("情境驅動對照")) + 1
_he = next((r for r in range(_hs + 1, WB["In"].max_row + 2) if not WB["In"].cell(r, 1).value), WB["In"].max_row + 1)
for b in SEC["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_hs}" for c in "BDEFC"]
        b["rows"] = [[f"@In!B{r}", f"@In!D{r}", f"@In!E{r}", f"@In!F{r}", f"@In!C{r}"] for r in range(_hs + 1, _he)]
# In 頁各節的鍵
_ws = WB["In"]; SK = {}; _cur = None
for _r in range(4, _ws.max_row + 1):
    _a, _b = _ws.cell(_r, 1).value, _ws.cell(_r, 2).value
    if _a and _b in (None, ""): _cur = re.split(r"[.\s（]", str(_a))[0]; SK[_cur] = []; continue
    if _a and _cur and re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", str(_a)): SK[_cur].append(str(_a))
_DCFK = ["rf", "rf_us", "beta_u", "beta_u_med", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]
_CAPK = ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]
_FCK = ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]
_J = [k for k in SK.get("J", []) if k not in _DCFK + _CAPK + _FCK]

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
 "in_skip": [in_skip_token(), in_skip_token("情境專屬列")] + SKIP_EXTRA,
 "in_groups": GROUPS(SK) + [
  ["全公司（毛利、費用、稅、股數、估值與目標價規則）", [
    ["毛利、營業費用、稅與股數", SK["G"]],
    ["估值、共識與對帳來源", SK["H"]],
    ["目標價規則與收斂法", _J + [k for k in SK.get("M2", []) if k not in _DCFK + _CAPK + _FCK]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", [k for k in _DCFK if k in I]],
    ["資本結構、淨現金與股數（評價用）", _CAPK],
    ["自由現金流（折舊、資本支出、營運資金）", _FCK],
    ["Wattonomics 錨定（只進敏感度）", SK["Z"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"]},
 "n_recon": 4,
}
# ================= v0.2.4（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5、Advantest v0.7.5、ASML v0.3.5；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（營收推導）→ 對帳（含 C0 實績）→ 資料狀態。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 C0；besi/v024_post.py 產生）。列號依標籤找；設定文字不含數字（年份與日期除外）。
def _v13(cfg):
    S = {x["id"]: x for x in cfg["sections"]}
    wb = WB
    RD, PLW, RV, SW, CW, UN = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"], wb["Units"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    fx_ = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip() == p)
    r13 = fnd(RD, "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.2.4"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "US$B": "n1", "€M": "n1", "百萬股": "n1", "€": "n2", "x": "n3", "EB": "n2", "Gb": "n1", "層": "n1", "百萬顆": "n0", "百萬顆當量": "n0", "台": "n0", "千片": "n0", "千片當量": "n1", "€M／千片當量": "n3"}.get(str(u).strip(), "n2")
    # ---- 橫幅：白話說明＋對帳中揭露句；評等規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2   # 末兩項＝空格、Coord 橫幅揭露句
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 混合鍵合與主流業務（堆疊）與三情境營收 ----
    hdr = 3
    assert PLW.cell(hdr, 1).value == "情境／項目" and PLW.cell(hdr, 3).value == "FY2025A", PLW.cell(hdr, 3).value
    pa, pb, pc = fx_(PLW, "情境 A"), fx_(PLW, "情境 B"), fx_(PLW, "情境 C")
    revA, revB, revC = pa + 1, pb + 1, pc + 1
    assert all(PLW.cell(x, 1).value == "營收" for x in (revA, revB, revC)) and PLW.cell(revB + 1, 1).value.strip() == "混合鍵合" and PLW.cell(revB + 2, 1).value.strip() == "主流業務"
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收組成（情境 B）與三情境營收（百萬歐元）", "x": f"@PL!C{hdr}:G{hdr}", "fmt": "n0", "stack": True, "pad": 1.3,
              "total": f"@PL!C{revB}:G{revB}",
              "bars": [{"name": f"@PL!A{revB + 1}", "y": f"@PL!C{revB + 1}:G{revB + 1}", "color": "#0b7a80"},
                       {"name": f"@PL!A{revB + 2}", "y": f"@PL!C{revB + 2}:G{revB + 2}", "color": "#c29a2b"}],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!C{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!C{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的混合鍵合營收與主流業務營收，頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為前三季實績加第四季指引的錨點；三情境的差異來自 HBM 位元、CoWoS 晶圓、混合鍵合導入比例與設備市場。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）----
    va = S["valuation"]["blocks"]
    g0 = fnd(wb["Val"], "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（歐元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    v_intro = S["valuation"].pop("intro", None)
    if v_intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [v_intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = fnd(SW, "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value and not str(SW.cell(r2 + 1, 1).value).startswith("快照狀態"):
        r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        vals = {r: SW[f"{col}{r}"].value for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float))}
        big = [r for r, v in vals.items() if abs(v) >= 0.1]
        small = [r for r, v in vals.items() if 0.0005 <= abs(v) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    snap = fnd(SW, "快照狀態")
    se = [b for b in se if not (b.get("type") == "tornado" and str(b.get("title", "")).startswith("加權目標價相對基準"))]
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = ("展開：對照：FY2027 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：FY2027 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}
    sec_intro = S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + ([{"type": "note", "parts": [sec_intro], "collapse": True, "summary": "展開：EPS 敏感度（情境 A）的說明"}] if sec_intro else []) + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rcc), rv("模型 WACC", rcc)
    a_rows = [rv(p, ra) for p in ("支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n1", "n2", "n1", "n1", "n2", "n2"]
    pe1, rc1, pe2 = rv("只看 Comparable：支持現價所需的前瞻本益比", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（百萬歐元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = fnd(PLW, "標準損益表（v0.2.4")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（銷管＋研發）", "營業利益", "營業利益率", "其他收入（淨額；＝ − 財務費用淨額）", "稅前利益", "所得稅",
            "稅後淨利", "淨利率", "基本加權平均股數（百萬股）", "EPS（IFRS 基本）"]
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    is_tbl = {"type": "table", "caption": "標準損益表（百萬歐元；EPS 為歐元）：FY2023A–FY2025A 實績（FY2023 只有營收）、FY2026E–FY2032E 情境 B（基準；IFRS，GAAP 口徑，無 Non-GAAP）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史、預測期對 PL 情境 B、延伸表、DCF 頁與稅率）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（百萬歐元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，IFRS）。FY2023A 只有營收；FY2026E 為錨點、FY2027E 起為情境推導；FY2030E–FY2032E 為收斂法延伸期（只推營收與淨利，營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要（FY2025–FY2029；營收、混合鍵合、主流業務、毛利率、EPS）")]
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                          + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + old_blocks)
    # ---- 6. 驅動因子：營收推導（情境 B）----
    rb = next(r for r in range(1, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("情境 B（基準）："))
    bh = rb + 1
    assert str(PLW.cell(bh, 3).value) == "2025" and str(PLW.cell(bh, 7).value) == "2029", (PLW.cell(bh, 3).value, PLW.cell(bh, 7).value)
    prow = lambda lab_, start=rb: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "CDEFG"] for r in rs]
    hrows = [prow(x) for x in ("HBM 出貨位元", "核心晶粒數", "混合鍵合占比", "所需機台", "新機需求＝MAX")]
    lrows = [prow(x) for x in ("AI 先進封裝晶圓需求", "混合鍵合占 AI 先進封裝晶圓比例", "混合鍵合邏輯晶圓", "新機需求指數")]
    prows = [prow(x) for x in ("HBM 混合鍵合營收", "邏輯與其他混合鍵合營收", "校準係數 k", "混合鍵合營收合計", "主流業務營收")]
    r_rev = next(r for r in range(prows[-1], PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_hsh, r_gm, r_oi, r_ni, r_eps = prow("混合鍵合占營收", r_rev), prow("毛利率", r_rev), prow("營業利益", r_rev), prow("淨利", r_rev), prow("EPS", r_rev)
    seg_tbl = {"type": "table", "caption": "營收推導（情境 B；日曆年＝會計年度）：HBM 位元與 CoWoS 晶圓 → 混合鍵合機台需求 → 混合鍵合營收；主流業務隨設備市場", "head": [f"@PL!A{bh}", f"@PL!B{bh}"] + [f"@PL!{c}{bh}" for c in "CDEFG"],
               "rows": [{"sub": "一、HBM 混合鍵合（HBM 位元 → 晶粒 → 混合鍵合機台）"}] + seg(hrows)
                       + [{"sub": "二、邏輯混合鍵合（CoWoS 晶圓 × 導入比例 → 新機需求指數）"}] + seg(lrows)
                       + [{"sub": "三、營收、毛利率與每股盈餘"}] + seg(prows + [r_rev, r_hsh, r_gm, r_oi, r_ni, r_eps])}
    hbr, mn = prows[3], prows[4]
    seg_chart = {"type": "theme", "title": "混合鍵合與主流業務營收（百萬歐元；情境 B）與混合鍵合占營收（右軸）", "x": f"@PL!C{bh}:G{bh}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25,
                 "bars": [{"name": f"@PL!A{hbr}", "y": f"@PL!C{hbr}:G{hbr}", "color": "#0b7a80"}, {"name": f"@PL!A{mn}", "y": f"@PL!C{mn}:G{mn}", "color": "#c29a2b"}],
                 "lines": [{"name": f"@PL!A{r_hsh}", "y": f"@PL!C{r_hsh}:G{r_hsh}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "長條＝混合鍵合營收（HBM 段＋邏輯段）與主流業務營收；折線＝混合鍵合占營收。FY2025 為實績（混合鍵合營收為假設拆分）、FY2026 為錨點、FY2027 起為情境 B 推導。"}
    dr = S["drivers"]["blocks"]
    d_intro = S["drivers"].pop("intro", None)
    S["drivers"].update(title="驅動因子：營收推導")
    S["drivers"]["blocks"] = ([{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}]
                              + ([{"type": "note", "parts": [d_intro]}] if d_intro else []) + dr)
    # ---- 7. 對帳：C0（讀取層級與實績）----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "對帳中項目數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "p1" if cunit(r) in ("%",) else "n3"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p2" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) in ("US$B",) else "n1"
    isprog = lambda r: str(CW.cell(r, 1).value or "").startswith("參考：已公布期間實績占")
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", (f"@Coord!F{r}|p1" if isprog(r) else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"]
                for r in range(mh + 1, a4e - 1)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + ("n0" if CW.cell(r, 2).value == "月" else fm(r)), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    ot = fnd(CW, "其他公司實績"); oh = next(r for r in range(ot, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); oe = next(r for r in range(oh, CW.max_row + 2) if not CW.cell(r + 1, 1).value)
    oth_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + ("n0" if CW.cell(r, 2).value == "台" else "n1"), f"@Coord!E{r}|" + ("n0" if CW.cell(r, 2).value == "台" else "n1"), f"@Coord!G{r}"] for r in range(oh + 1, oe + 1)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
                  {"type": "table", "caption": "營收：公司季報實績 vs 層級隱含值（層級沒有 BESI 份額；份額沿用 CY2025 實績的兩列為假設，只列參考、不判定）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e - 1}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
                  {"type": "table", "caption": f"@Coord!A{ot}", "collapse": True, "summary": "展開：訂單、毛利、營業利益、淨利與混合鍵合累計訂單台數的實績（層級無可比量，列示供對照）", "head": [f"@Coord!{c}{oh}" for c in "ABCDE"] + [f"@Coord!G{oh}"], "rows": oth_rows}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl, {"type": "note", "parts": [f"@Coord!A{cnt + 1}"]}]
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}

    def _rd(r):
        v = CW.cell(r, 2).value
        return [f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(v, (int, float)) and not isinstance(v, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [_rd(r) for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = fnd(CW, "層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"],
                   "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "公司目標、先驗比對與校準"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與先驗比對"
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]



_OUT = hero_rule_a(run(SPEC))
# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(_OUT, X)
_v13(_OUT)
dump(_OUT)

