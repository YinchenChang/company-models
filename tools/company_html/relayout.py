"""把既有 html.json 轉成「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）的共用轉換器。

用法（在各公司的 gen 腳本或建置腳本最後呼叫）：
    from relayout import relayout
    cfg = relayout(cfg, workbook_path, spec)

spec（每家公司自訂，dict）：
  in_groups   [[塊標題, [[小組標題, [In 鍵, …]], …]], …]   依業務性質分塊；最後一塊放全公司參數
  in_skip     [In 節字母, …]   不必歸組的節（例如共識、情境對照），預設 ["J"]
  blocks_to   {既有 drivers 區塊的 caption 或 note 參照開頭: 塊標題}   既有命題模組表／說明放到哪一塊（未列者放最後一塊之前）
  main        {表格 caption 開頭（取值後文字）: [主要列標籤開頭, …]}   估值、對帳表格的主要列；未列的表格取前 n_main 列
  n_main      預設主要列數（預設 5）
  thesis_note 補充說明（可省略）

README 頁須有兩段（由模型建置腳本寫入，公式帶數字）：
  「展示層」標題列之後依序：命題摘要、命題一、二、三（可多於三點，至「白話重點」前）、白話重點 1–3
  實作：標題列 A 欄開頭為「展示層」；其下第 1 列＝命題摘要；其下各列 A 欄開頭為「命題 n：」者為命題，「重點 n：」者為白話重點
  「命題回覆」標題列之後：每列一題答案（至空列止）
"""
import re
import openpyxl

TOR_BIG = 0.1
CONC_MAIN = ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "Comparable 目標價", "現價", "上檔空間：加權目標價", "上檔空間：Comparable", "投資建議", "情境目標價", "市場共識目標價", "WACC"]   # 敏感度：|變動| ≥ 10% 直接顯示，其餘摺疊


def _unit_fmt(u):
    u = str(u or "")
    if u == "%":
        return "p1"
    if u == "x":
        return "n2"
    if u in ("次", "年", "季", "US$/月"):
        return "n1"
    if u.startswith("US$m/") or u in ("GW", "GW/年", "US$", "MW"):
        return "n2"
    if u == "十億股":
        return "n3"
    return "n1"


def _expand(rows):
    out = []
    for part in str(rows).split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += list(range(int(a), int(b) + 1))
        elif part:
            out.append(int(part))
    return out


def _ref_parts(spec):
    """'@Val!A12|n2' → ('Val', 'A', 12, 'n2')；非單格參照回傳 None。"""
    m = re.match(r"^@([^!]+)!([A-Z]+)(\d+)(?:\|(.*))?$", spec) if isinstance(spec, str) else None
    return (m.group(1), m.group(2), int(m.group(3)), m.group(4) or "") if m else None


class _WB:
    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path, data_only=True)

    def v(self, sh, co):
        return self.wb[sh][co].value

    def text(self, spec):
        if isinstance(spec, list):
            return "".join(self.text(p) for p in spec)
        p = _ref_parts(spec)
        if p:
            return str(self.v(p[0], f"{p[1]}{p[2]}") or "")
        return str(spec or "")

    def find(self, sh, prefix, col="A", start=1, end=None):
        ws = self.wb[sh]
        for r in range(start, (end or ws.max_row) + 1):
            if str(ws[f"{col}{r}"].value or "").strip().startswith(prefix):
                return r
        return None


def _rowgen_to_rows(w, blk):
    """rowgen 表格 → 明確列（保留格式與分組小標），回傳 [(標籤文字, 列規格或 {"sub"})]。"""
    g = blk["rowgen"]
    cols, fmts = g["cols"], g.get("fmts", [""] * len(g["cols"]))
    out = []
    for r in _expand(g["rows"]):
        ov = g.get("row_fmt", {}).get(str(r), {})
        vals = [w.v(g["sheet"], f"{c}{r}") for c in cols]
        if len(cols) > 1 and vals[0] is not None and all(x is None for x in vals[1:]):
            out.append((str(vals[0]), {"sub": f"@{g['sheet']}!{cols[0]}{r}"}))
            continue
        out.append((str(vals[0] or ""), [f"@{g['sheet']}!{c}{r}|{ov.get(c, f)}" for c, f in zip(cols, fmts)]))
    return out


def _table_rows(w, blk):
    if "rowgen" in blk:
        return _rowgen_to_rows(w, blk)
    out = []
    for r in blk.get("rows", []):
        if isinstance(r, dict):
            out.append((w.text(r["sub"]), r))
        else:
            out.append((w.text(r[0]), r))
    return out


def _clean(cap, fallback):
    """摘要文字不得含手寫數字：去掉節號，仍含數字就用通用字樣。"""
    c = re.sub(r"^[\d.\s]+", "", str(cap or "")).strip()
    c = re.sub(r"（.*?）", "", c).strip()
    return c if c and not re.search(r"\d", re.sub(r"(FY|CY)\d{4}", "", c)) else fallback


def _split_table(w, blk, prefixes, n_main, summary):
    rows = _table_rows(w, blk)
    data = [(lab, r) for lab, r in rows if not isinstance(r, dict)]
    if prefixes:
        main_ids = set()
        for p in prefixes:   # 「標籤開頭|第二欄文字」可指定年度等
            lp, _, sp = p.partition("|")
            for i, (lab, r) in enumerate(data):
                if lab.strip().startswith(lp.strip()) and (not sp or w.text(r[1]).strip() == sp) and i not in main_ids:
                    main_ids.add(i)
                    break
            else:
                raise KeyError(f"主要列找不到：{p!r}（表 {blk.get('caption')!r}）")
        main = [data[i][1] for i in sorted(main_ids)]
    else:
        main = [r for _, r in data[:n_main]]
    rest_rows, cur_sub = [], None
    main_set = {id(r) for r in main}
    for lab, r in rows:
        if isinstance(r, dict):
            cur_sub = r
            continue
        if id(r) in main_set:
            continue
        if cur_sub is not None:
            rest_rows.append(cur_sub)
            cur_sub = None
        rest_rows.append(r)
    base = {k: v for k, v in blk.items() if k not in ("rowgen", "rows", "collapse", "summary")}
    out = []
    if main:
        out.append(dict(base, rows=main))
    if [r for r in rest_rows if not isinstance(r, dict)]:
        out.append({"type": "table", "collapse": True, "summary": summary, "head": blk.get("head", []), "rows": rest_rows})
    return out


def _in_tables(w, groups_for_block, label):
    ws = w.wb["In"]
    key_row = {str(ws[f"A{r}"].value): r for r in range(4, ws.max_row + 1) if ws[f"A{r}"].value}
    head = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"]

    def row(k):
        r = key_row[k]
        f = _unit_fmt(ws[f"C{r}"].value)
        return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]

    def ranged(k):
        r = key_row[k]
        return ws[f"D{r}"].value not in (None, "") or ws[f"F{r}"].value not in (None, "")

    a_rows, b_rows = [], []
    for sub, keys in groups_for_block:
        a = [row(k) for k in keys if ranged(k)]
        b = [row(k) for k in keys if not ranged(k)]
        if a:
            a_rows += [{"sub": sub}] + a
        if b:
            b_rows += [{"sub": sub}] + b
    out = []
    if a_rows:
        out.append({"type": "table", "caption": f"{label}：有敏感度區間的參數（基準進目標價；低／高只用於敏感度頁）", "head": head, "rows": a_rows})
    if b_rows:
        out.append({"type": "table", "collapse": True, "summary": f"展開：{label}的其他固定輸入（實績、公司揭露、換算值）", "head": head, "rows": b_rows})
    return out


def _check_in_coverage(w, spec):
    ws = w.wb["In"]
    skip = set(spec.get("in_skip", ["J"]))
    sec, need = None, set()
    for r in range(4, ws.max_row + 1):
        a = str(ws[f"A{r}"].value or "")
        if a and ws[f"B{r}"].value in (None, ""):
            sec = re.split(r"[.\s]", a)[0]
            continue
        if a and sec not in skip:
            need.add(a)
    got = [k for _, gs in spec["in_groups"] for _, ks in gs for k in ks]
    dup = {k for k in got if got.count(k) > 1}
    if dup:
        raise SystemExit(f"In 參數重複歸組：{sorted(dup)}")
    miss, extra = need - set(got), set(got) - need
    if miss or extra:
        raise SystemExit(f"In 參數歸組不完整：缺 {sorted(miss)}；多 {sorted(extra)}")


def _hero(w):
    f = lambda p: w.find("Val", p)
    rec, tp, tpc, px, up, upc = f("投資建議"), f("目標價（DCF 與 Comparable 加權）"), f("Comparable 目標價"), f("現價"), f("上檔空間：加權目標價"), f("上檔空間：Comparable")
    if None in (rec, tp, tpc, px, up, upc):
        raise SystemExit("Val 頁找不到投資建議橫幅所需的列（投資建議、目標價、Comparable 目標價、現價、兩個上檔空間）")
    h = {"rec": f"@Val!B{rec}",
         "tiles": [[f"@Val!A{tp}", f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                   [f"@Val!A{tpc}", f"@Val!B{tpc}|n2", ["上檔空間 ", f"@Val!B{upc}|sp1"]],
                   [f"@Val!A{px}", f"@Val!B{px}|n2", ""]]}
    if w.v("Val", f"E{rec}"):
        h["note"] = ["投資建議依加權目標價判定：", f"@Val!E{rec}"]
    return h


def _readme(w):
    r0 = w.find("README", "展示層")
    ra = w.find("README", "命題回覆")
    if not r0 or not ra:
        raise SystemExit("README 頁缺「展示層」或「命題回覆」段（見 docs/plan/個股HTML版面規格.md）")
    ws = w.wb["README"]
    lead = r0 + 1
    props, obs, r = [], [], r0 + 2
    while ws[f"A{r}"].value:
        t = str(ws[f"A{r}"].value)
        (props if t.startswith("命題") else obs).append(f"@README!A{r}")
        r += 1
    ans, r = [], ra + 1
    while ws[f"A{r}"].value:
        ans.append(f"@README!A{r}")
        r += 1
    return f"@README!A{lead}", props, obs, ans


def relayout(cfg, xlsx, spec):
    w = _WB(xlsx)
    _check_in_coverage(w, spec)
    n_main = spec.get("n_main", 5)
    mains = spec.get("main", {})
    lead, props, obs, ans = _readme(w)
    secs = {s["id"]: s for s in cfg["sections"]}

    # 1. 命題
    cfg["thesis"] = lead
    cfg["wide_text"] = True
    cfg["hero"] = _hero(w)
    old_thesis = secs["thesis"]
    appendix = [b for b in old_thesis["blocks"] if b["type"] == "list" and b is not None][-1:] if old_thesis["blocks"] else []
    secs["thesis"]["blocks"] = [
        {"type": "h3", "text": "命題"},
        {"type": "list", "ordered": True, "items": props},
        {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）", "items": ans},
        {"type": "h3", "text": "重點觀察"},
        {"type": "list", "ordered": True, "items": obs}]

    # 2. 驅動因子：情境表 → 各業務塊（參數＋既有命題模組表）→ 全公司
    dr = secs["drivers"]
    old = dr["blocks"]
    intro = dr.pop("intro", None)
    titles = [t for t, _ in spec["in_groups"]]
    placed = {t: [] for t in titles}
    scen, driver_map, others, drv_text = None, None, [], []
    for b in old:
        cap = w.text(b.get("caption", "")) if b.get("caption") else ""
        if cap.startswith("情境驅動"):
            scen = b
        elif cap.startswith("驅動 → 推導") or (b["type"] == "table" and b.get("head", [None])[:1] == ["頁"]):
            driver_map = b
        elif cap.startswith("單一參數區間"):
            continue
        elif b["type"] in ("h3", "list"):
            drv_text.append(b)
            continue
        else:
            others.append(b)
    if intro:
        others.insert(0, {"type": "note", "parts": [intro]})
    dest_default = titles[-2] if len(titles) > 1 else titles[-1]
    for b in others:
        key = w.text(b.get("caption", "")) if b.get("caption") else (b.get("parts", [""])[0] if b["type"] == "note" else "")
        dest = next((t for p, t in spec.get("blocks_to", {}).items() if str(key).startswith(p) or (b["type"] == "note" and b.get("parts", [""])[0] == p)), dest_default)
        placed[dest].append(b)
    nums = "一二三四五六七八九十"
    blocks = [scen] if scen else []
    for i, (t, gs) in enumerate(spec["in_groups"]):
        blocks.append({"type": "h3", "text": f"{nums[i]}、{t}"})
        blocks += _in_tables(w, gs, t.split("（")[0])
        blocks += placed[t]
    dr["blocks"] = blocks

    # 4. 估值：長條圖 → 主要列；折現率直接顯示
    va = secs["valuation"]
    nb, bars, rest = [], [b for b in va["blocks"] if b["type"] == "bars"], [b for b in va["blocks"] if b["type"] != "bars"]
    nb += bars
    for b in rest:
        if b["type"] == "table" and not b.get("collapse"):
            cap = w.text(b.get("caption", ""))
            pref = next((v for k, v in mains.items() if cap.startswith(k)), None)
            if pref is None:
                hit = [p for p in CONC_MAIN if any(lab.strip().startswith(p) for lab, _ in _table_rows(w, b))]
                if len(hit) >= 4:
                    pref = hit
            nb += _split_table(w, b, pref, n_main, f"展開：{_clean(cap, '本表')}的其他項目")
        elif b["type"] == "kv":
            nb.append({"type": "table", "caption": "折現率與資本結構", "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]})
        else:
            nb.append(b)
    va["blocks"] = nb

    # 5. 對帳：每節主要列
    rc = secs["recon"]
    nb = []
    for b in rc["blocks"]:
        if b["type"] == "table" and not b.get("collapse"):
            cap = w.text(b.get("caption", ""))
            pref = next((v for k, v in mains.items() if cap.startswith(k)), None)
            nb += _split_table(w, b, pref, spec.get("n_recon", 4), f"展開：{_clean(cap, '本節')[:30]}的其他對帳項目")
        else:
            nb.append(b)
    rc["blocks"] = nb

    # 6. 敏感度：|變動| ≥ 10% 直接顯示
    se = secs["sens"]
    nb = []
    for b in se["blocks"]:
        if b["type"] == "tornado":
            rows = _expand(b["rows"])
            vals = [(r, w.v(b["sheet"], f"{b['value_col']}{r}")) for r in rows]
            big = [str(r) for r, v in vals if isinstance(v, (int, float)) and abs(v) >= TOR_BIG]
            small = [str(r) for r, v in vals if isinstance(v, (int, float)) and 0 < abs(v) < TOR_BIG]
            if big:
                nb.append(dict(b, rows=",".join(big), title="加權目標價相對基準的變動（變動達兩位數百分比的案例；一次改一個輸入，整本重算）"))
            if small:
                nb.append(dict(b, rows=",".join(small), collapse=True, summary="展開：變動未達兩位數百分比的案例", title="加權目標價相對基準的變動（未達兩位數百分比）"))
        else:
            nb.append(b)
    se["blocks"] = nb

    # 7. 資料狀態：附錄
    st = secs["status"]
    app = []
    if driver_map:
        dm = dict(driver_map)
        dm.pop("caption", None)
        app += [{"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"}, dict(dm, collapse=True, summary="展開：每頁的輸入、推導量與因果方向")]
    if drv_text:
        app.append({"type": "h3", "text": "附錄：各頁因果（查核用）"})
        app += [dict(b, collapse=True, summary="展開：各頁因果") for b in drv_text if b["type"] == "list"]
    for b in old_thesis_extra(cfg, w):
        app.append(b)
    ins = next((i for i, b in enumerate(st["blocks"]) if b["type"] == "h3" and "未找到" in str(b.get("text"))), len(st["blocks"]))
    st["blocks"] = st["blocks"][:ins] + app + st["blocks"][ins:]
    return cfg


def old_thesis_extra(cfg, w):
    """版本要點：README 頁「本版要點」或「與上一版的差異」段 → 附錄（摺疊）。"""
    ws = w.wb["README"]
    r0 = None
    rv = w.find("README", "版本：")
    if rv and not w.find("README", "本版要點"):
        return [{"type": "h3", "text": "附錄：版本要點"},
                {"type": "table", "collapse": True, "summary": "展開：本版與前版的修改要點", "head": ["要點"], "rows": [[f"@README!A{rv}"]]}]
    for p in ("本版要點", "與上一版的差異", "版本要點"):
        r0 = w.find("README", p)
        if r0:
            break
    if not r0:
        return []
    rows, r = [], r0 + 1
    while ws[f"A{r}"].value and not str(ws[f"A{r}"].value).startswith(("工作表：", "展示層", "命題回覆")):
        rows.append([f"@README!A{r}"])
        r += 1
    if not rows:
        return []
    return [{"type": "h3", "text": "附錄：版本要點"},
            {"type": "table", "collapse": True, "summary": "展開：本版與前版的修改要點", "head": ["要點"], "rows": rows}]


if __name__ == "__main__":
    # 用法：python3 tools/company_html/relayout.py <公司資料夾> <成品 xlsx>
    # 讀 <公司資料夾>/layout_spec.py 的 SPEC，就地改寫 <公司資料夾>/html.json（建置腳本寫出 html.json 之後、LibreOffice 重算之後執行一次）
    import json, sys, os, importlib.util
    cdir, xlsx = sys.argv[1], sys.argv[2]
    sp = importlib.util.spec_from_file_location("layout_spec", os.path.join(cdir, "layout_spec.py"))
    mod = importlib.util.module_from_spec(sp); sp.loader.exec_module(mod)
    hp = os.path.join(cdir, "html.json")
    cfg = json.load(open(hp, encoding="utf-8"))
    if cfg.get("hero"):
        raise SystemExit("html.json 已是新版面（有 hero）；請先由建置腳本重寫 html.json 再執行")
    cfg = relayout(cfg, xlsx, mod.SPEC)
    json.dump(cfg, open(hp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("relayout：", hp)
