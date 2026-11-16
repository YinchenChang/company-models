# Kioxia v0.5.5（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Advantest v0.7.5；docs/plan/個股HTML版面規格.md「v1.3 規格」）。
# 由 kioxia/gen_html_json.py 最後呼叫（COORD.inject_html 之後）。章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子 → 對帳（含 C0 實績）→ 資料狀態。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 C0；kioxia/v055_post.py 產生）。列號依標籤找；設定文字不含數字（年份與日期除外）。


def apply(cfg, wb, ver="v0.5.5", ccy="日圓", unit_b="十億日圓", co="Kioxia", scen_lbl=("情境 A", "情境 C"), seg_note=None):
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = fnd(RD, "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert ver + "（" in str(RD[f"A{RM['ver']}"].value), RD[f"A{RM['ver']}"].value   # v0.5.6：本版要點列前置 px1113 說明，v0.5.5 說明接在其後
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "US$B": "n2", "$B": "n2", "¥B": "n0", "百萬股": "n1", "¥": "n0", "USD": "n2", "$": "n2", "JPY": "n0", "x": "n2"}.get(str(u).strip(), "n1")
    # ---- 橫幅：白話說明＋對帳中揭露句；原說明移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 分部營收（堆疊）與情境 A、C 營收 ----
    hdr = 3
    assert PLW.cell(hdr, 1).value == "情境／項目", PLW.cell(hdr, 1).value
    pa, pb, pc = fnd(PLW, "情境 A（"), fnd(PLW, "情境 B（"), fnd(PLW, "情境 C（")
    rowB = lambda lab_: next(r for r in range(pb + 1, pc) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    revA, revB, revC = pa + 1, pb + 1, pc + 1
    assert PLW.cell(revA, 1).value == "營收" and PLW.cell(revB, 1).value == "營收" and PLW.cell(revC, 1).value == "營收"
    segs = seg_note["segs"]
    seg_rows = [rowB(x) for x in segs]
    colors = ["#0b7a80", "#6fc7cb", "#c29a2b", "#8a8c8c"]
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": f"營收分部（情境 B）與三情境營收（{unit_b}）", "x": f"@PL!C{hdr}:G{hdr}", "fmt": "n0", "stack": True, "pad": 1.3,
              "total": f"@PL!C{revB}:G{revB}",
              "bars": [{"name": f"@PL!A{r}", "y": f"@PL!C{r}:G{r}", "color": c} for r, c in zip(seg_rows, colors)],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!C{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!C{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": seg_note["caption"]}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）----
    va = S["valuation"]["blocks"]
    V = wb["Val"]
    g0 = fnd(V, "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 6)); r_pxg = g0 + 6
    assert str(V.cell(r_pxg, 1).value).startswith("現價"), V.cell(r_pxg, 1).value
    four = {"type": "bars", "title": f"四個目標價與現價（{ccy}；豎線＝現價）", "fmt": "n0",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 0, 1), ("DCF 目標價", "Comparable 目標價（進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n0", "", f"@Val!D{r_pxg}"]]}
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
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = ("展開：對照：EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}
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
    afmt = ["n0", "n1", "n3", "n1", "n1", "n2", "n0"]
    pe1, rc1, pe2 = rv("只看 Comparable：支持現價所需的前瞻 P/B", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          [f"FY2028 營收（{unit_b}）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n2"), (rc1, "n2"), (pe2, "n2"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFGH"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFGH"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = fnd(PLW, f"標準損益表（{ver}")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    PC = "CDEFGHIJKL"
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    is_tbl = {"type": "table", "caption": seg_note["is_caption"],
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(r) for r in range(ph + 1, r_ratio)] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（預測期對 PL 情境 B、延伸表與 DCF 頁）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": f"營收（{unit_b}）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。歷史欄空白＝未找到；FY2026E 起為情境推導；FY2030E–FY2032E 為收斂法延伸期（只推營收與淨利，營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要與折線（FY2025–FY2029）")]
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                          + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "CDEFG"] for r in rs]
    drows = [rowB(x) for x in seg_note["drv_rows"]]
    seg_tbl = {"type": "table", "caption": seg_note["drv_caption"], "head": [f"@PL!A{hdr}", f"@PL!B{hdr}"] + [f"@PL!{c}{hdr}" for c in "CDEFG"], "rows": seg(drows)}
    r_gm, r_om = rowB("毛利率"), rowB("營業利益率")
    seg_chart = {"type": "theme", "title": f"營收（{unit_b}）與毛利率、營業利益率（右軸）：情境 B", "x": f"@PL!C{hdr}:G{hdr}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
                 "bars": [{"name": f"@PL!A{r}", "y": f"@PL!C{r}:G{r}", "color": c} for r, c in zip(seg_rows, colors)],
                 "lines": [{"name": f"@PL!A{r_gm}", "y": f"@PL!C{r_gm}:G{r_gm}", "axis": "r", "color": "#3b3838", "labels": True},
                           {"name": f"@PL!A{r_om}", "y": f"@PL!C{r_om}:G{r_om}", "axis": "r", "color": "#b4472f", "dash": True}],
                 "caption": seg_note["drv_chart_caption"]}
    dr = S["drivers"]["blocks"]
    d_intro = S["drivers"].pop("intro", None)
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"]["blocks"] = ([{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}]
                              + ([{"type": "note", "parts": [d_intro]}] if d_intro else []) + dr)
    # ---- 7. 對帳：C0（讀取層級與實績）----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "對帳中項目數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "p1" if cunit(r) == "%" else "n3" if cunit(r) in ("x", "US$B", "$B") else "n1"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) in ("US$B", "$B", "x") else "n0" if str(CW.cell(r, 2).value) in ("月", "¥/US$") else "n1"
    isprog = lambda r: str(CW.cell(r, 1).value or "").startswith("參考：已公布期間實績占")
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n2" if not isprog(r) else f"@Coord!D{r}|n1", f"@Coord!E{r}|n2" if not isprog(r) else f"@Coord!E{r}|n1",
                 (f"@Coord!F{r}|p1" if isprog(r) else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"] for r in range(mh + 1, a4e - 1)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + fm(r), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    ot = fnd(CW, "其他公司實績"); oh = next(r for r in range(ot, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); oe = next(r for r in range(oh, CW.max_row + 2) if not CW.cell(r + 1, 1).value)
    oth_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!G{r}"] for r in range(oh + 1, oe + 1)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
                  {"type": "table", "caption": seg_note["c0_caption"], "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e - 1}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
                  {"type": "table", "caption": f"@Coord!A{ot}", "collapse": True, "summary": "展開：其他分部與損益的實績（層級無可比量，列示供對照）", "head": [f"@Coord!{c}{oh}" for c in "ABCDE"] + [f"@Coord!G{oh}"], "rows": oth_rows}]
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
                   "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "公司實績、指引與共識"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與共識"
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]
