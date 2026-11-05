# 光寶科 html.json 產生器（v0.4.2：DCF 彙總加 stub 期兩列、In 分組加第二輪修正新參數；v0.3.1：個股 HTML 版面規格 v1.0，docs/plan/個股HTML版面規格.md；v0.4：DCF 改標準版（FCFF＋WACC），估值章改放折現率、DCF 彙總與各情境 FCFF 明細）
# 讀 liteon/html_src.json（原版面設定，內容不動）→ 依列標籤重定位與整理 → relayout() → 寫 liteon/html.json（勿手改 html.json）。
# 用法：python3 liteon/gen_html_json.py [成品 xlsx]（省略時用 liteon/dist/ 內唯一的 xlsx）
# ───────── 以下為共通整理步驟（各公司 gen_html_json.py 相同；公司差異在檔尾 SPEC 與 CO 設定） ─────────
import copy, glob, json, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

X = sys.argv[1] if len(sys.argv) > 1 else (lambda fs: fs[0] if len(fs) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{fs}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
wb = openpyxl.load_workbook(X, data_only=True)
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
sec = {s["id"]: s for s in cfg["sections"]}
IW = wb["In"]
INK = {str(IW.cell(r, 1).value): r for r in range(4, IW.max_row + 1) if IW.cell(r, 1).value}


def row(sh, label, start=1, exact=False, must=True):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == label.strip()) if exact else v.startswith(label.strip()):
            return r
    if must:
        raise KeyError((sh, label))
    return None


def until_blank(sh, r0):
    ws, r = wb[sh], r0
    while ws.cell(r, 1).value not in (None, ""):
        r += 1
    return r - 1


def text(spec):
    m = re.match(r"^@([^!]+)!([A-Z]+)(\d+)", spec) if isinstance(spec, str) else None
    return str(wb[m.group(1)][f"{m.group(2)}{m.group(3)}"].value or "") if m else str(spec or "")


def plain(t, fallback):
    """摘要文字：去節號與括號內容；仍含數字（年度除外）就用通用字樣（產生器禁止手寫數字）。"""
    c = re.sub(r"^[\d.\s]+", "", str(t or "")).strip()
    c = re.sub(r"（.*?）", "", c).strip()
    return c if c and not re.search(r"\d", re.sub(r"(FY|CY)\d{4}", "", c)) else fallback


V = lambda lab, c="B", f="": f"@Val!{c}{row('Val', lab)}" + (f"|{f}" if f else "")
VA = lambda lab: f"@Val!A{row('Val', lab)}"
SCN0 = row("In", "情境驅動對照")
SCN_TOKEN = re.split(r"[.\s]", str(IW.cell(SCN0, 1).value))[0]

# ── 驅動因子：「單一參數」對照表 → 附錄的驅動 → 推導對照表；情境驅動表依 In 頁現行列號重建 ──
for b in sec["drivers"]["blocks"]:
    cap = str(b.get("caption", ""))
    if cap.startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照表"
    elif cap.startswith("情境驅動"):
        h = SCN0 + 1
        b["head"] = [f"@In!{c}{h}" for c in "BDEFCG"]
        b["rows"] = [[f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]
                     for r in range(h + 1, until_blank("In", h + 1) + 1)
                     for f in [{"%": "p1", "x": "n2", "權重": "p0", "pt": "p1"}.get(str(IW.cell(r, 3).value), "n2")]]

# ── 估值：目標價與現價長條圖 → 結論主要項目 → DCF 折現率 → DCF → Comparable → 相對倍數調整；舊法與規則值集中摺疊 ──
VW = wb["Val"]
_hdr = []
for r in range(1, VW.max_row + 1):
    a = VW.cell(r, 1).value
    m = re.match(r"^(\d+(?:\.\d+)?)[\s.]", a) if isinstance(a, str) else None
    if m and all(VW.cell(r, c).value in (None, "") for c in range(2, 6)):
        _hdr.append((r, m.group(1)))


def val_kind(r):
    k = None
    for hr, num in _hdr:
        if hr <= r:
            k = num
    return "tp" if k is not None and k.split(".")[0] == "0" else "ref"


def val_rows(b):
    rs = [int(x) for x in re.findall(r"@Val!(?:[A-Z]+)(\d+)", json.dumps({k: v for k, v in b.items() if k not in ("caption",)}, ensure_ascii=False))]
    g = b.get("rowgen")
    if g and g.get("sheet") == "Val":
        rs.append(int(str(g["rows"]).split("-")[0].split(",")[0]))
    return min(rs) if rs else None


def pfmt(lab):
    v = VW.cell(row("Val", lab), 2).value
    return "n0" if isinstance(v, (int, float)) and abs(v) >= 1000 else "n2"


DROP_ROWS = {row("Val", p, must=False) for p in ("上檔空間級距", "投資建議", "目標價方法")} - {None}
CONC = [("目標價（DCF 與 Comparable 加權）", "px"), ("DCF 目標價（情境加權）", "px"), ("Comparable 目標價（相對倍數調整後）", "px"),
        ("Comparable 目標價（市場主流同業法）", "px"), ("現價（", "px"), ("上檔空間：加權目標價", "sp1"), ("上檔空間：Comparable 目標價", "sp1"),
        ("上檔空間級距", ""), ("投資建議", ""), ("市場共識目標價", "px"), ("與市場共識目標價的差距：EPS 因素", "sp1"),
        ("與市場共識目標價的差距：倍數因素", "sp1"), ("價格帶", ""), ("目標價方法", "")]
conc = {"type": "table", "caption": "目標價、現價與投資建議", "head": ["項目", "數值"],
        "rows": [[VA(l), V(l, f=pfmt(l) if f == "px" else f)] for l, f in CONC if row("Val", l, must=False)]}
hr_ = row("Val", "歷史倍數檢查", must=False)
hist = [{"type": "table", "collapse": True, "summary": "展開：歷史倍數檢查（現價隱含市盈率的十年分位；只作檢查）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
         "rows": [[f"@Val!A{hr_}", f"@Val!B{hr_}|p0", f"@Val!C{hr_}|p0", f"@Val!D{hr_}|p0"]]}] if hr_ else []
bars = {"type": "bars", "title": "目標價與現價", "fmt": pfmt("目標價（DCF 與 Comparable 加權）"), "items": [
    {"label": VA(l), "value": V(l), "series": s} for l, s in (("DCF 目標價（情境加權）", 0), ("Comparable 目標價（相對倍數調整後）", 1),
                                                         ("目標價（DCF 與 Comparable 加權）", 2), ("市場共識目標價", 0)) if row("Val", l, must=False)],
    "ref": {"label": VA("現價（"), "value": V("現價（")}}
def dcf_blocks():
    """標準版 DCF：折現率表（緊接 DCF 彙總之前）、DCF 彙總、各情境 FCFF 明細（摺疊）、口徑說明。列號依標籤定位。"""
    vr = lambda lab: row("Val", lab)
    rate_t = {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率", "head": ["項目", "數值", "說明"],
              "rows": [[f"@Val!A{vr(l)}", f"@Val!B{vr(l)}|{f}", f"@Val!E{vr(l)}"] for l, f in
                       (("股權成本 ke", "p2"), ("稅後債務成本", "p2"), ("權益占資本", "p1"), ("WACC", "p2"), ("永續成長率 g", "p2"), ("永續期 RONIC", "p2"))]}
    h = vr("股權成本 ke") - 1
    px = pfmt("DCF 目標價（情境加權）") if "pfmt" in globals() else "n2"
    summ = [("明確期 FCFF 現值", "n1"), ("終值現值", "n1"), ("企業價值 EV", "n1"), ("＋淨現金與投資", "n1"), ("＋Stub 期 FCFF", "n1"), ("淨現金滾到評價日", "n1"), ("評價用稀釋股數", "n3"), ("DCF 情境每股價值（評價日）", px),
            ("DCF 情境價值推到 12 個月後", px), ("其中：終值現值占 EV", "p1"), ("隱含終值 EV／EBITDA", "x1")]
    sum_t = {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{h}" for c in "ABCD"],
             "rows": [[f"@Val!A{vr(l)}"] + [f"@Val!{c}{vr(l)}|{f}" for c in "BCD"] for l, f in summ] + [[f"@Val!A{vr('DCF 目標價（情境加權）')}", f"@Val!B{vr('DCF 目標價（情境加權）')}|{px}", "", ""]]}
    D = wb["DCF"]
    hd = next(r for r in range(1, 12) if D.cell(r, 1).value == "項目")
    cols = [get_column_letter(c) for c in range(3, D.max_column + 1) if D.cell(hd, c).value]
    out = [rate_t, sum_t]
    for r in range(1, D.max_row + 1):
        if D.cell(r, 1).value == "營收" and str(D.cell(r - 1, 1).value or "").startswith("情境"):
            pv = next(k for k in range(r, r + 20) if str(D.cell(k, 1).value or "").startswith("現值（FCFF"))
            rf = {str(k): {c: ("p1" if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷")) else "n3") for c in cols}
                  for k in range(r, pv + 1) if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷", "折現因子"))}
            out.append({"type": "table", "caption": f"@DCF!A{r - 1}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hd}"] + [f"@DCF!{c}{hd}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r}-{pv}", "cols": ["A"] + cols, "fmts": [""] + ["n1"] * len(cols), "row_fmt": rf}})
    out.append({"type": "note", "parts": ["@DCF!A2"]})
    return out


tp_b, ref_b, rel_tbl, pend, rate_done = [], [], None, None, False
for b in sec["valuation"]["blocks"]:
    t = b["type"]
    if t == "bars" or (t == "h3" and b.get("text") == "重點觀察") or (t == "list" and str(b.get("items", [""])[0]).startswith("@README!")):
        continue
    if t == "h3":
        pend = b["text"]
        continue
    rr = val_rows(b)
    if t == "kv":
        if pend and pend.startswith("相對倍數調整"):
            rel_tbl = {"type": "table", "caption": pend, "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}
            pend = None
            continue
        b = {"type": "table", "caption": pend or "", "head": ["項目", "數值"], "rows": [list(it) for it in b["items"]]}
        if not b["caption"]:
            b.pop("caption")
    if t == "note" and rr in DROP_ROWS:
        continue
    if t == "table" and rr is not None and val_kind(rr) == "tp" and (text(b.get("caption", "")).startswith("0.3") or b.get("caption") == "目標價（DCF 與 Comparable 加權）"):
        continue   # 原結論表：由上方 conc 取代
    kind = val_kind(rr) if rr else ("tp" if not ref_b else "ref")
    if kind == "tp":
        if b["type"] == "table" and text(b.get("caption", "")).startswith("0.1") and not rate_done:
            tp_b.extend(dcf_blocks())   # 舊 0.1 DCF（股東現金流）表：改為標準版 DCF 的折現率、彙總與明細
            rate_done = True
            pend = None
            continue
        if pend:
            tp_b.append({"type": "h3", "text": pend})
        tp_b.append(b)
    else:
        if b["type"] == "table" and not b.get("collapse"):
            first = b["rows"][0][0] if b.get("rows") and isinstance(b["rows"][0], list) else ""
            sect = [hr for hr, _ in _hdr if rr and hr <= rr]
            fb = plain(VW.cell(sect[-1], 1).value, "") if sect else ""
            b = dict(b, collapse=True, summary="展開：" + plain(pend or text(b.get("caption", "")), fb or (plain(text(first), "參考") + "等項目")))
        elif pend:
            ref_b.append({"type": "h3", "text": pend})
        ref_b.append(b)
    pend = None
if not rate_done:
    tp_b[0:0] = dcf_blocks()
# 敏感度：目標價參數表（標準版 DCF 案例）依 Sens 頁現行列號重建
for b in sec["sens"]["blocks"]:
    if b["type"] == "table" and text(b.get("caption", "")).startswith("目標價參數"):
        r0 = row("Sens", "目標價參數")
        b["caption"] = f"@Sens!A{r0}"
        b["head"] = [f"@Sens!{c}{r0 + 1}" for c in "ABCDEFG"]
        _pf = pfmt("目標價（DCF 與 Comparable 加權）") if "pfmt" in globals() else "n2"
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r0 + 2}-{row('Sens', '檢查：閉式公式', r0) - 1}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "", _pf, "sp1", _pf, "sp1", _pf]}
sec["valuation"]["blocks"] = [bars, conc] + hist + tp_b + ([rel_tbl] if rel_tbl else []) + \
    ([{"type": "h3", "text": "參考：舊法、規則值與市場隱含倍數（不進目標價）"}] + ref_b if ref_b else [])
_rv = row("README", "估值：", must=False)
if _rv:
    sec["valuation"]["intro"] = f"@README!A{_rv}"   # Val!A1 偏重舊法（規則值），改用 README 的目標價方法說明
sec["thesis"].pop("intro", None)   # 與頂層核心問題重複
old_thesis_kv = copy.deepcopy(next((b for b in sec["thesis"]["blocks"] if b["type"] == "kv"), None))


def post(cfg):
    # 1. 橫幅第二格（補充規則 A）：只看 Comparable 目標價（市場主流同業法，未經相對倍數調整）與其上檔空間；
    #    有相對倍數調整時，進加權的調整後值寫在橫幅附註（上檔空間取 README 頁 HTML 橫幅附註用列）
    cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V("Comparable 目標價（市場主流同業法）", f="n2"),
                               ["上檔空間 ", V("上檔空間：Comparable 目標價", f="sp1")]]
    hb = row("README", "HTML 橫幅附註用：", must=False)
    if hb and row("Val", "Comparable 目標價（相對倍數調整後）", must=False):
        cfg["hero"]["note"] = cfg["hero"]["note"] + ["；進加權的 Comparable 為相對倍數調整後 ", V("Comparable 目標價（相對倍數調整後）", f="n2"),
                                                     "（上檔空間 ", f"@README!B{hb}|sp1", "）"]
    # 2. 敏感度龍捲風圖：標題依 Sens 頁欄名（本模型龍捲風不是加權目標價）
    for b in sec["sens"]["blocks"]:
        if b["type"] == "tornado":
            b["title"] = [f"@Sens!{b['value_col']}3", "，一次動一個輸入（", "@Sens!A4", "）",
                          "：變動未達兩位數百分比的案例" if b.get("collapse") else "：變動達兩位數百分比的案例"]
    # 3. In 頁單位「pt」（比例差）與「NT$」（每股金額）顯示格式：轉換器預設一位小數，改為百分比與兩位小數
    for b in sec["drivers"]["blocks"]:
        for r in b.get("rows", []):
            if isinstance(r, list) and isinstance(r[0], str) and r[0].startswith("@In!B"):
                u = str(IW.cell(int(r[0][5:]), 3).value)
                to = {"pt": "|p1", "NT$": "|n2"}.get(u)
                if to:
                    r[1:4] = [c.replace("|n1", to) if isinstance(c, str) else c for c in r[1:4]]
    # 4. 第一章補情境定義與口徑（原第一章的對照表）
    if old_thesis_kv:
        sec["thesis"]["blocks"] += [{"type": "h3", "text": "情境定義與口徑"}, old_thesis_kv]
    json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok", os.path.relpath(X, HERE))

# ───────── 光寶科：依業務性質分塊 ─────────
for b in sec["recon"]["blocks"]:
    if b["type"] == "table" and "rowgen" in b and not b.get("caption"):
        b["caption"] = "本模型與共識、台達電模型、校準值的對照"
AI = "AI 電源（雲端及物聯網事業群的 AI 相關）"
SPEC = {
 "in_skip": ["G", SCN_TOKEN],
 "in_groups": [
  [AI, [
    ["AI 部署規模（讀網通市場模型、橋接表與台積電顆數係數）", ["itc_26", "itc_27", "itc_28", "itc_29", "itpg_gb", "itpg_vr", "vr26", "racks_vr", "um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "mf"]],
    ["每 kW 電源內容（台達電模型）與光寶份額", ["pkw_26", "pkw_27", "pkw_28", "pkw_29", "x_np", "rel_27", "rel_28", "rel_29", "ai_25", "ai_26"]],
    ["AI 相關毛利率", ["d_ai", "gmu_27", "gmu_28", "gmu_29"]],
    ["情境 A／C 專屬列（份額相對值與 AI 毛利率額外改善）", ["rel_27_A", "rel_28_A", "rel_29_A", "gmu_27_A", "gmu_28_A", "gmu_29_A", "rel_27_C", "rel_28_C", "rel_29_C", "gmu_27_C", "gmu_28_C", "gmu_29_C"]]]],
  ["非 AI 事業（雲端及物聯網非 AI、光電、資訊及消費性電子）", [
    ["事業群比重與第二季實績", ["s25_cl", "s25_op", "s25_ic", "q2_cl", "q2_op", "q2_ic", "g1_cl", "g1_op", "g1_ic", "s2h_cl", "s2h_op"]],
    ["非 AI 分部成長率", ["g_cn27", "g_cn28", "g_cn29", "g_op27", "g_op28", "g_op29", "g_ic27", "g_ic28", "g_ic29"]]]],
  ["全公司：錨點、費用、股價、折現率與目標價規則", [
    ["FY2026 錨點（月營收實績與第四季假設）", ["ytd9", "q3_rev", "q4g", "q_gm", "q_opx", "q_nop", "q_nir", "q_sh"]],
    ["營業費用、業外、稅與股數（FY2027 起）", ["phi", "opx_adj", "nop_f", "nir_f", "sh_g"]],
    ["匯率", ["fx"]],
    ["股價、市場倍數與共識目標價", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim_avg"]],
    ["目標價規則", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["延伸期收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"AI 電源市場": AI, "@Units!A17": AI, "標記沿用": AI},
 "main": {"目標價、現價與投資建議": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（相對倍數調整後）", "現價（", "上檔空間：加權目標價", "投資建議", "市場共識目標價"],
          "DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "相對倍數調整": ["同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
          "本模型與共識": ["FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "CY2027 光寶＋台達電 AI 電源有效份額合計", "目標價（DCF 與 Comparable 加權）vs 市場共識目標價"]},
}
cfg = relayout(cfg, X, SPEC)
post(cfg)
