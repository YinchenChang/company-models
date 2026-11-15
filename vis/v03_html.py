# 世界先進 v0.3（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 聯電 umc/gen_html_json.py 的 _v13）＋三層協調讀取層級與 C0；兩讀數不適用（無 Coord 兩讀數區塊）。
# 章節：命題 → 估值（含 VSMC SOTP）→ 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（三層協調 C0）→ 資料狀態。
# 數字與白話一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表與分部營收推導、RevDCF、VSMC、Val 0.7、Sens 末表、Coord 頁；vis/v03_post.py 產生）；本檔文字不含年份與日期以外的數字。


def apply(cfg, WB, row_of):
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RVW, CW, SEW, VSW = WB["README"], WB["PL"], WB["RevDCF"], WB["Coord"], WB["Sens"], WB["VSMC"]
    r13 = row_of("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"] or [])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    for t in cfg["hero"]["tiles"]:
        t[0] = [t[0], "（新台幣）"] if isinstance(t[0], str) else list(t[0]) + ["（新台幣）"]
    # ---- 1. 命題 ----
    sg = row_of("PL", "分部營收推導（情境 B")
    hS = sg + 1
    ex = lambda label, start=sg: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == label)
    r_ship, r_asp, r_w, r_v, r_rs, r_sh = (ex(t) for t in ("晶圓出貨（8 吋）", "平均售價（8 吋晶圓）", "8 吋晶圓營收", "VSMC 營收（12 吋）", "營收", "12 吋（VSMC）÷ 營收"))
    r_cap, r_chk = ex("產能（8 吋）"), ex("核對：8 吋＋VSMC − 營收（應為 0）")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B，十億新台幣）與 VSMC 占營收：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.35,
              "total": f"@PL!F{r_rs}:I{r_rs}",
              "bars": [{"name": [f"@PL!A{r_w}"], "y": f"@PL!F{r_w}:I{r_w}", "color": "#6fc7cb"}, {"name": [f"@PL!A{r_v}"], "y": f"@PL!F{r_v}:I{r_v}", "color": "#0b7a80"}],
              "lines": [{"name": [f"@PL!A{r_sh}"], "y": f"@PL!F{r_sh}:I{r_sh}", "axis": "r", "color": "#3b3838", "labels": True}],
              "caption": "堆疊＝情境 B（基準）的八吋晶圓與 VSMC 十二吋營收；頂端數字＝營收；折線＝VSMC 占營收（右軸）。FY2026 為前三季實績加第四季指引的錨點。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    S["thesis"]["blocks"] = [chart1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}] + th
    # ---- 2. 估值：四個目標價置頂；VSMC SOTP 摘要 ----
    va = S["valuation"]["blocks"]
    g0 = row_of("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣；豎線＝現價）", "fmt": "n1",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價（SOTP）", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
    vs0 = next(r for r in range(VSW.max_row, 0, -1) if str(VSW.cell(r, 1).value or "").startswith("SOTP 摘要"))
    vse = next(r for r in range(vs0, VSW.max_row + 1) if str(VSW.cell(r, 1).value or "").startswith("有效 DCF 權重"))
    vfm = lambda r: "" if str(VSW.cell(r, 1).value).startswith(("VSMC 正常化", "DCF 適用")) else "n2"
    sotp = {"type": "table", "caption": f"@VSMC!A{vs0}", "head": [f"@VSMC!{c}{vs0 + 1}" for c in "ACDE"],
            "rows": [[f"@VSMC!A{r}"] + [f"@VSMC!{c}{r}" + (f"|{vfm(r)}" if vfm(r) else "|") for c in "CDE"] for r in range(vs0 + 2, vse - 1)]
                    + [[f"@VSMC!A{vse - 1}", f"@VSMC!C{vse - 1}|n0", "", ""], [f"@VSMC!A{vse}", f"@VSMC!C{vse}|p0", "", ""]]}
    vsb = next(r for r in range(1, VSW.max_row + 1) if str(VSW.cell(r, 1).value or "").startswith("三情境 FCFF 明細") and "情境 B" in str(VSW.cell(r, 1).value))
    vh = vsb + 1; vend = next(r for r in range(vh, VSW.max_row + 1) if str(VSW.cell(r, 1).value or "").startswith("SOTP 調整"))
    vcols = [chr(ord("C") + i) for i in range(14)]
    vtbl = {"type": "table", "caption": f"@VSMC!A{vsb}", "collapse": True, "summary": "展開：VSMC 情境 B 逐年營收、建廠折舊瀑布、自由現金流與企業價值",
            "head": [f"@VSMC!A{vh}"] + [f"@VSMC!{c}{vh}" for c in vcols],
            "rows": [[f"@VSMC!A{r}"] + [f"@VSMC!{c}{r}|" + ("p1" if "毛利率" in str(VSW.cell(r, 1).value) else "n2") for c in vcols] for r in range(vh + 1, vh + 11)]
                    + [[f"@VSMC!A{r}", f"@VSMC!C{r}|" + ("" if "檢查" in str(VSW.cell(r, 1).value) else "n2")] + [""] * 13 for r in range(vh + 11, vend + 1)]}
    rules = {"type": "note", "parts": old_note} if old_note else None
    old_bars = [b for b in va if b.get("type") == "bars" and b.get("title") == "目標價與現價"]
    rest = [b for b in va if b not in old_bars]
    top = [bars4]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [tbl4] + ([rules] if rules else []) + [{"type": "h3", "text": "DCF 的分部加總：既有八吋業務與 VSMC 六成"}, {"type": "note", "parts": ["@VSMC!A1"]}, sotp, {"type": "note", "parts": ["@VSMC!A2"], "collapse": True, "summary": "展開：VSMC 估值口徑"}, vtbl]
    S["valuation"]["blocks"] = top + rest
    # ---- 3. 敏感度 ----
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
                         summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例")))
    all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案與設定）",
               "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCD"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}"] for r in range(r1, r2 + 1)]}
    others = []
    for b in S["sens"]["blocks"]:
        if b.get("type") == "tornado":
            if b.get("value_col") == "J":
                continue
            b = dict(b, collapse=True, summary=("展開：對照：" + str(b.get("title") or "EPS 敏感度")))
        elif b.get("type") == "table" and not b.get("collapse"):
            b = dict(b, collapse=True, summary=("展開：敏感度明細表" if str(b.get("caption") or "").startswith("@") or not b.get("caption") else "展開：" + str(b["caption"])[:40]))
        others.append(b)
    intro = S["sens"].pop("intro", None)
    lead = [tor[0][0], tor[1][0], {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入，三情境同時改；VSMC 折舊拆分各案每案重算 PL、八吋 DCF 與 VSMC 分部估值）。"]}]
    if intro: lead.append({"type": "note", "parts": [intro], "collapse": True, "summary": "展開：敏感度頁的說明"})
    S["sens"]["blocks"] = lead + [tor[0][1], tor[1][1], all_tbl] + others
    # ---- 4. 反向 DCF ----
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
            {"type": "table", "caption": "現金流折現法（情境 B 八吋路徑；VSMC 分部價值固定）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（八吋；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["現金流放大倍數（各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 八吋營收（十億新台幣）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 八吋營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(kvr, kvf)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{cm[0]}", "—", f"@RevDCF!C{cm[0]}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(cm[1:4])]
                     + [[f"@RevDCF!A{cm[4]}", "—", f"@RevDCF!C{cm[4]}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "collapse": True, "summary": "展開：情境 B 八吋各年營收、營業利益與現值（反解的計算輸入；十億新台幣）",
             "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"],
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrr, yfm)]}]}
    # ---- 5. 損益 ----
    p0 = row_of("PL", "標準損益表（v0.3")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含其他營業損益）", "營業利益", "營業利益率", "其他收入（淨額；營業外）", "稅前利益", "所得稅",
            "非控制權益損益", "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股；基本加權平均）", "EPS（新台幣；基本，歷史為公司公告值）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "十億股": "n2"}.get(str(u).strip(), "n1")
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk2 = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk2, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC]
    is_tbl = {"type": "table", "caption": "標準損益表（十億新台幣；EPS 為新台幣）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）；IFRS 合併（GAAP 口徑）", "head": head,
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk2)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk2}", "head": head, "collapse": True, "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B、延伸期推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk2 + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億新台幣）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸；IFRS 合併，淨利為歸屬母公司）。FY2026E 為前三季實績加第四季指引；VSMC 量產後建廠折舊使營業利益率下滑；FY2030E 起為收斂法。"}
    legacy = []
    for b in S["pl"]["blocks"]:
        if b["type"] == "pl":
            legacy.append(dict(b, collapse=True, summary="展開：三情境損益摘要與折線"))
        elif b["type"] in ("note", "kv", "table"):
            legacy.append(dict(b, collapse=True, summary="展開：" + ("延伸期（FY2030 起）說明" if b["type"] == "note" else "三情境錨點核對")))
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + legacy)
    # ---- 6. 驅動因子 ----
    sf = {"千片": "n0", "%": "p1", "US$/片": "n0", "NT$/US$": "n2", "NT$B": "n1"}
    seg_rows = [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{'n3' if r == r_chk else sf.get(str(PLW.cell(r, 2).value), 'n2')}" for c in "FGHI"] for r in range(r_cap, r_sh + 1)]
    seg_tbl = {"type": "table", "caption": f"@PL!A{sg}", "head": [f"@PL!A{hS}", f"@PL!B{hS}"] + [f"@PL!{c}{hS}" for c in "FGHI"], "rows": seg_rows}
    chart2 = {"type": "theme", "title": "八吋晶圓出貨（情境 B，千片）與平均售價（美元／片，右軸）：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.3,
              "bars": [{"name": [f"@PL!A{r_ship}"], "y": f"@PL!F{r_ship}:I{r_ship}", "color": "#6fc7cb"}],
              "lines": [{"name": [f"@PL!A{r_asp}"], "y": f"@PL!F{r_asp}:I{r_asp}", "axis": "r", "color": "#0b7a80", "labels": True}],
              "caption": "長條＝八吋晶圓出貨（產能 × 稼動率）；折線＝八吋平均售價（美元，FY2026 為校準值）。稼動率與售價是公司層級假設，Siliconomics 尚無成熟製程代工的量與價。"}
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, chart2, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + S["drivers"]["blocks"]
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 7. 對帳：三層協調 C0 ----
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    c1, c2, c3 = cr("1. 讀哪一層"), cr("2. 讀取層級"), cr("3. C0（實績）")
    h1, h2 = c1 + 1, c2 + 1
    h3 = next(r for r in range(c3, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = cr("對帳中項目數（C0 實績表")
    act = {"type": "table", "caption": "公司實績（月營收、第三季損益、稼動率、資本支出指引）與層級隱含值", "head": [f"@Coord!{c}{h3}" for c in "ABCDEFG"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|" + ("p1" if CW.cell(r, 4).value == "%" else "n3"), f"@Coord!D{r}", f"@Coord!E{r}", f"@Coord!F{r}", f"@Coord!G{r}"] for r in range(h3 + 1, cnt)]}
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + ("|n2" if str(CW.cell(r, 2).value).startswith("=") else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"] for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    same = {"type": "table", "caption": f"@Coord!A{c2}", "collapse": True, "summary": "展開：讀取層級核對（匯率；同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFGH"],
            "rows": [[f"@Coord!A{h2 + 1}", f"@Coord!B{h2 + 1}", f"@Coord!C{h2 + 1}", f"@Coord!D{h2 + 1}|n2", f"@Coord!E{h2 + 1}|n2", f"@Coord!F{h2 + 1}|sp1", f"@Coord!G{h2 + 1}", f"@Coord!H{h2 + 1}"]]}
    blocks = [{"type": "h3", "text": "三層協調：讀取層級與公司實績對帳"}, {"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}, act,
              {"type": "note", "parts": ["對帳中項目數：", f"@Coord!F{cnt}|n0", "。"]}, {"type": "note", "parts": [f"@Coord!A{cr('C4／C5')}"]}, rd_tbl, same,
              {"type": "h3", "text": "本模型與公司說法、共識、券商的對照（原有）"}]
    S["recon"]["blocks"] = blocks + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與共識對照"
    # ---- 8. 資料狀態 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    # ---- 白話替換：可見文字（摺疊與資料狀態以外）不出現參數代號；xlsx 原文不改（設定檔文字不含年份與日期以外的數字）----
    import re as _re
    PLAIN = {f"VSMC!A{vse}": "有效 DCF 權重（＝設定權重 × DCF 適用）",
             f"Val!A{row_of('Val', 'DCF 情境價值推到 12 個月後')}": "DCF 情境價值推到目標時點（以股權成本複利，扣除期間股利）",
             f"Val!E{row_of('Val', 'WACC', exact=True)}": "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）",
             f"Val!A{row_of('Val', '永續成長率 g')}": "永續成長率 g（長期成長上限與無風險利率取小者）",
             f"Val!A{row_of('Val', '永續期 RONIC')}": "永續期 RONIC（＝WACC＋超額報酬差）",
             f"Val!E{row_of('Val', '＋Stub 期 FCFF')}": "資產負債表日到評價日的 stub 期自由現金流（含 stub 期的會計年度 FCFF 按期間比例；DCF 頁下方）"}
    IW = WB["In"]
    for r in range(1, IW.max_row + 1):
        k, t = str(IW.cell(r, 1).value or ""), str(IW.cell(r, 2).value or "")
        if _re.match(r"^vgm2\d_[ABC]$", k) or ("vda_adj" in t and "情境" in t):
            PLAIN[f"In!B{r}"] = _re.sub(r"（.*情境 ([ABC])）$", r"（情境 \1；現金毛利率＝本值＋八吋折舊強度）", t)
        elif k == "g_tv":
            PLAIN[f"In!B{r}"] = "永續成長率 g（最後一年後；長期成長上限與無風險利率取小者）"

    def _plain(x):
        if isinstance(x, list):
            return [_plain(v) for v in x]
        if isinstance(x, str) and x.startswith("@") and x[1:].split("|")[0] in PLAIN:
            return PLAIN[x[1:].split("|")[0]]
        return x
    for sc in cfg["sections"]:
        if sc["id"] == "status":
            continue
        for b in sc["blocks"]:
            if b.get("collapse"):
                continue
            if b.get("type") == "kv":
                b["items"] = [_plain(it) for it in b.get("items", [])]
            if b.get("type") == "table":
                b["rows"] = [({"sub": _plain(rw["sub"])} if isinstance(rw, dict) else _plain(rw)) for rw in b.get("rows", [])]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]
