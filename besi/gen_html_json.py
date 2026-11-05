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
               {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"],
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
dump(hero_rule_a(run(SPEC)))
