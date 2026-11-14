# 記憶體樣板（SK 海力士 v0.7.6）HTML 版面 v1.3 轉換：把 html_layout.build 產生的 v1.0 版面 cfg 改成個股 HTML 版面規格 v1.3（docs/plan/個股HTML版面規格.md）。
# 章節：thesis → valuation → sens → revdcf → pl → drivers → recon → status（記憶體無個別模組）。樣板 NVIDIA（nvidia/gen_html_json.py 的 _v13、_coord、_top_first）。
# 數字與標籤一律取自 xlsx（Hist、PL、RevDCF、Val 0.7、Sens、Coord、README 新增列；由 tools/memory_model/v13_post.py 寫入）；列號依標籤找。
# 只由 html_layout.build(..., v13=True) 呼叫（美光、三星不呼叫）。


def transform(cfg, wb):
    S = {s["id"]: s for s in cfg["sections"]}

    def lab(sh, prefix, start=1, col=1, exact=False):
        ws = wb[sh]
        for r in range(start, ws.max_row + 1):
            v = ws.cell(r, col).value
            if isinstance(v, str) and ((v.strip() == prefix.strip()) if exact else v.strip().startswith(prefix.strip())):
                return r
        raise KeyError((sh, prefix))

    PLW, RV, CW, SEW, RD = wb["PL"], wb["RevDCF"], wb["Coord"], wb["Sens"], wb["README"]
    cfg["layout"] = "v1.3"
    # ---- README v1.3 展示層列 ----
    r13 = lab("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.7."), RD[f"A{RM['ver']}"].value
    # ---- 橫幅：只留白話（含對帳中揭露）----
    old_note = list(cfg["hero"]["note"])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題：圖＋重點觀察（5 點）----
    ph = lab("PL", "三情境並列（圖用）")
    scen_h = ph + 1
    row_of = lambda prefix: lab("PL", prefix, ph)
    ni_rows = [row_of(f"淨利（情境 {s}") for s in "ABC"]
    th = S["thesis"]["blocks"]
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, {"type": "theme", "title": "三情境淨利（兆韓元）：盈餘在哪一年見頂", "x": f"@PL!C{scen_h}:J{scen_h}", "fmt": "n0", "stack": False, "pad": 1.15,
                  "bars": [{"name": f"@PL!A{r}", "y": f"@PL!C{r}:J{r}", "color": c} for r, c in zip(ni_rows, ("#b7e3e4", "#10a0a7", "#8a8c8c"))],
                  "caption": "長條＝三種價格情境的稅後淨利（樂觀、基準、保守；2025 為實績，2026 為基準年，2030 起為收斂法）。回答核心問題：盈餘在哪一年到高點、之後回落多少；2027 年各情境接近，分歧從 2028 年起放大。"})
    # ---- 2. 估值：四個目標價置頂；說明放圖下；兩讀數區緊接其後 ----
    va = S["valuation"]["blocks"]
    g0 = lab("Val", "0.7 估值並列")
    gr = list(range(g0 + 2, g0 + 7))
    r_px = g0 + 7
    # COORD.inject_html 已在最前面插入：h3、note、table、bars（兩讀數）
    coord_blocks = va[0:4]
    assert coord_blocks[0]["type"] == "h3" and coord_blocks[3]["type"] == "bars", [b["type"] for b in coord_blocks]
    def _int(x):
        if isinstance(x, str):
            return x[:-3] + "|n0" if x.endswith("|n2") else x
        if isinstance(x, list):
            return [_int(i) for i in x]
        if isinstance(x, dict):
            return {k: ("n0" if k == "fmt" and v == "n2" else _int(v)) for k, v in x.items()}
        return x
    coord_blocks = _int(coord_blocks)
    r_tpc = next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("加權目標價") and isinstance(CW.cell(r, 2).value, (int, float)))
    coord_blocks.append({"type": "table", "caption": "兩讀數加權目標價的精確值（韓元，兩位小數；供與 xlsx 核對，上表四捨五入為整數）", "collapse": True,
                         "summary": "展開：兩讀數加權目標價的精確值（兩位小數）", "head": ["項目", "現行基準讀數", "另一側讀數"],
                         "rows": [[f"@Coord!A{r_tpc}", f"@Coord!B{r_tpc}|n2", f"@Coord!C{r_tpc}|n2"]]})
    for t_ in cfg["hero"]["tiles"]:
        t_[:] = _int(t_)
    rest = va[4:]
    rest = [b for b in rest if not (b["type"] == "bars" and str(b.get("title", "")).startswith("目標價與現價"))]
    bar4 = {"type": "bars", "title": "四個目標價與現價（韓元；豎線＝現價）", "fmt": "n0",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr]}
    _on = [p_ for p_ in old_note[:7] if not (str(p_).startswith("@") and "固定規則" in str(wb[str(p_)[1:].split("!")[0]][str(p_)[1:].split("!")[1].split("|")[0]].value or "")) and "投資建議依" not in str(p_)]
    if _on and isinstance(_on[0], str) and _on[0].startswith("。"):
        _on[0] = _on[0][1:]
    va_new = [bar4, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}, tbl4, {"type": "note", "parts": _on}] + coord_blocks + rest
    S["valuation"]["blocks"] = va_new
    S["valuation"].pop("intro", None)
    # ---- 3. 敏感度：加權目標價相對基準（低案、高案）置頂 ----
    s0 = lab("Sens", "加權目標價相對基準（HTML")
    r1 = s0 + 2
    r2 = lab("Sens", "快照狀態", r1) - 1
    num = lambda col, r: isinstance(SEW[f"{col}{r}"].value, (int, float))
    mags = lambda col: sorted((abs(SEW[f"{col}{r}"].value) for r in range(r1, r2 + 1) if num(col, r)), reverse=True)
    out = []
    any_big = any(max(mags(c) + [0]) >= 0.1 for c in "BC")
    th_main = 0.1 if any_big else 0.01
    for col, side in (("B", "低"), ("C", "高")):
        cand = [r for r in range(r1, r2 + 1) if num(col, r) and abs(SEW[f"{col}{r}"].value) >= 0.0005]
        cand.sort(key=lambda r: -abs(SEW[f"{col}{r}"].value))
        big = [r for r in cand if abs(SEW[f"{col}{r}"].value) >= th_main] if any_big else cand[:10]
        small = [r for r in cand if r not in big]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        ttl = (f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例" if any_big
               else f"加權目標價相對基準：輸入取{side}值（無案例達兩位數，列影響最大者）")
        out.append(dict(base, title=ttl, rows=",".join(map(str, big))))
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值（其餘案例）", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：輸入取{side}值，其餘案例（變動較小；幾乎為零者見全部案例表）"))
    out.append({"type": "table", "caption": f"@Sens!A{s0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{s0 + 1}" for c in "ABCDEF"],
                "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}|n0", f"@Sens!E{r}|n0", f"@Sens!F{r}|p1"] for r in range(r1, r2 + 1)]})
    st_r = r2 + 1
    out.append({"type": "note", "parts": ["快照狀態：", f"@Sens!B{st_r}", "。", f"@Sens!C{st_r}"]})
    sb = S["sens"]["blocks"]
    for b in sb:
        if b["type"] == "tornado":
            b["collapse"] = True
            tl = b.get("title")
            tail = tl[-1] if isinstance(tl, list) and tl and isinstance(tl[-1], str) and not tl[-1].startswith("@") else ""
            b["summary"] = "展開：對照：營業利益對各收入線同幅變動的敏感度" + tail
    main_idx = [i for i, b in enumerate(out) if b["type"] == "tornado" and not b.get("collapse")]
    top = [out[i] for i in main_idx]
    intro = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（輸入取低值、取高值各一張，一次動一個輸入，整本重算；達兩位數百分比者直接顯示）。"
                                      "沒有任何單一輸入使加權目標價變動達兩位數百分比時，主圖改列變動較大的案例。下方摺疊為其餘案例、全部案例表、"
                                      "營業利益對各收入線的敏感度與目標價參數明細。"]}
    S["sens"]["blocks"] = top + [intro] + [b for i, b in enumerate(out) if i not in main_idx] + sb
    S["sens"].pop("intro", None)
    # ---- 4. 反向 DCF ----
    A = {}
    for key, pref in (("a", "A. DCF"), ("a_v0", "支持現價所需的評價日每股價值"), ("evr", "支持現價所需的企業價值"), ("ev", "模型企業價值 EV"), ("evx", "EV 須調整"), ("vb", "模型：情境 B 評價日"),
                      ("yh", "項目（情境 B）"), ("ra", "(a)"), ("rb", "(b)"), ("rc", "(c)"), ("cmp", "B. Comparable"), ("sum", "C. 摘要")):
        A[key] = lab("RevDCF", pref)
    f_ = lambda p, start: lab("RevDCF", p, start)
    m1, m0 = f_("現價隱含的營業利益率", A["ra"]), f_("模型：基準的等效營業利益率", A["ra"])
    s1, rv1, rv0, g1, g0_ = (f_("現價隱含的營收倍數", A["rb"]), f_("現價隱含的 2028E 營收", A["rb"]), f_("模型：2028E 營收", A["rb"]),
                             f_("現價隱含的 2028E 營收年增率", A["rb"]), f_("模型：2028E 營收年增率", A["rb"]))
    w1, w0 = f_("現價隱含的 WACC", A["rc"]), f_("模型 WACC", A["rc"])
    cm = A["cmp"]
    ybody = list(range(A["yh"] + 1, A["ra"] - 1))
    yfmt = lambda r: "p1" if str(RV.cell(r, 1).value).startswith(("營業利益率", "稅率")) else "n3" if str(RV.cell(r, 1).value).startswith(("評價日後", "折現因子", "終值係數", "每 1")) else "n1"
    rvdcf = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{A['sum']+1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{A['sum']+2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（2028E–2032E 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（2028E 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["2028E 營收（兆韓元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["2028E 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{A['a']}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n0" if r in (A["a"] + 2, A["a"] + 3, A["a"] + 9) else "n6" if r == A["a"] + 5 else "x1" if r == A["a"] + 8 else "n1")]
                      for r in range(A["a"] + 2, A["a"] + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{cm}", "head": [f"@RevDCF!{c}{cm + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{cm + 2}", "—", f"@RevDCF!C{cm + 2}|n0", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|n2", f"@RevDCF!C{r}|n2", f"@RevDCF!D{r}|sp1"] for r in range(cm + 3, cm + 6)]
                     + [[f"@RevDCF!A{cm + 6}", "—", f"@RevDCF!C{cm + 6}|n2", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{A['yh']}", "head": [f"@RevDCF!{c}{A['yh']}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{yfmt(r)}" for c in "BCDEFG"] for r in ybody]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p5 = lab("PL", "項目", 1, 1, exact=True)
    plr = lambda labl, start=p5: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == labl)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（銷管＋研發）", "營業利益", "營業利益率", "其他收入（淨額）", "稅前利益", "所得稅", "稅後淨利", "淨利率", "股數（百萬股）", "EPS（完全稀釋）"]
    fmt = lambda u: {"%": "p1", "韓元": "n0", "百萬股": "n1"}.get(u, "n1")
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in "CDEFGHIJ"]
    r_ratio = plr("比率")
    r_chk = plr("核對（應為 0）")
    r_d30 = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（2030E）：毛利"))
    is_tbl = {"type": "table", "caption": "標準損益表（GAAP；兆韓元，股數百萬股，EPS 韓元）：2023A–2025A 實績、2026E–2030E 情境 B（基準）",
              "head": [f"@PL!A{p5}"] + [f"@PL!{c}{p5}" for c in "CDEFGHIJ"],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]
                      + [{"sub": "核對：2030E 兩種推法的營業利益差額（毛利率與營業費用率沿用 2029E；目標價用延伸期淨利，不受影響）"}]
                      + [rowx(r) for r in (r_d30, r_d30 + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（兆韓元）與毛利率、營業利益率、淨利率：2023A–2030E（基準情境 B）", "x": f"@PL!C{p5}:J{p5}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:J{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:J{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:J{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:J{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。2026E 為基準年；標示「對帳中」的年度依賴 Siliconomics 對帳中的 HBM 產業量；2030E 為收斂法（只推營收與淨利），毛利率與營業費用率沿用 2029E。"}
    scr = lambda base_label, prefix: next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith(f"{prefix}（情境 {base_label}"))
    rvr = {s: scr(s, "營收") for s in "ABC"}
    epr = {s: scr(s, "EPS") for s in "ABC"}
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，兆韓元）與 EPS（折線，韓元，右軸）：2026E–2032E", "x": f"@PL!D{scen_h}:J{scen_h}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.1,
                "bars": [{"name": f"@PL!A{rvr[s]}", "y": f"@PL!D{rvr[s]}:J{rvr[s]}", "color": c} for s, c in zip("ABC", ("#b7e3e4", "#10a0a7", "#8a8c8c"))],
                "lines": [{"name": f"@PL!A{epr[s]}", "y": f"@PL!D{epr[s]}:J{epr[s]}", "axis": "r", "color": c, "dash": d} for s, c, d in zip("ABC", ("#0b7a80", "#3b3838", "#b4472f"), (True, False, True))],
                "caption": "情境之間 2027E 相同（價格已議定），差異自 2028E 起放大；2030E 起為收斂法（淨利率與營收年增向長期水準收斂）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    notes = [p_ for b in old_pl if b["type"] == "note" for p_ in b["parts"]]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（2026–2029；營收、營業利益、淨利、每股淨值、EPS）"),
                  {"type": "h3", "text": "延伸期（2030–2032）的收斂法"}, dict(pls[1], collapse=True, summary="展開：三情境延伸到 2032（收斂法；營收、淨利率、淨利、EPS）")]
    old_blocks.append({"type": "note", "parts": notes, "collapse": True, "summary": "展開：延伸期說明"})
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": ["@PL!A2"]}, is_chart, is_tbl, {"type": "note", "parts": ["@PL!A3"]}, sc_chart] + old_blocks)
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    sh = lab("PL", "分部營收推導（情境 B")
    hdrB = sh + 1
    sr = lambda labl, start=sh: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == labl)
    r_tot = sr("分部加總（營收）")
    dr_, hr_, nr_, or_ = sr("一般 DRAM 營收"), sr("HBM 營收"), sr("NAND 營收"), sr("其他營收")
    hs = sr("HBM ÷ 營收")
    sg = lambda rows, f: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{f(r)}" for c in "DEFG"] for r in rows]
    ff = lambda r: {"%": "p1", "x": "n3", "指數": "n3", "EB": "n2", "US$/GB": "n2", "US$B": "n1", "韓元／美元": "n0"}.get(PLW.cell(r, 2).value, "n1")
    seg_rows = []
    sec_titles = [r for r in range(sh + 2, r_tot) if PLW.cell(r, 2).value is None and PLW.cell(r, 1).value]
    for r in range(sh + 2, sr("核對：分部加總 − 合併營收（應為 0）") + 1):
        if PLW.cell(r, 2).value is None and PLW.cell(r, 1).value:
            seg_rows.append({"sub": f"@PL!A{r}"})
        else:
            seg_rows += sg([r], ff)
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；兆韓元）：一般 DRAM＋HBM＋NAND＋其他＝合併營收", "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"], "rows": seg_rows}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，兆韓元）與 HBM 占營收比（右軸）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2, "total": f"@PL!D{r_tot}:G{r_tot}",
                 "bars": [{"name": f"@PL!A{dr_}", "y": f"@PL!D{dr_}:G{dr_}", "color": "#6fc7cb"}, {"name": f"@PL!A{hr_}", "y": f"@PL!D{hr_}:G{hr_}", "color": "#0b7a80"},
                          {"name": f"@PL!A{nr_}", "y": f"@PL!D{nr_}:G{nr_}", "color": "#c29a2b"}, {"name": f"@PL!A{or_}", "y": f"@PL!D{or_}:G{or_}", "color": "#8a8c8c"}],
                 "lines": [{"name": f"@PL!A{hs}", "y": f"@PL!D{hs}:G{hs}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝一般 DRAM、HBM、NAND、其他四個分部（非記憶體事業為零，未畫）；頂端數字＝分部加總；折線＝HBM 占營收比。2026E 為基準年（前三季實績＋第四季推估）；2027E 起由價格路徑與位元成長推導。"}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 7. 對帳：C0 在最前面 ----
    c0 = lab("Coord", "C0 由下往上對帳")
    hdr = lab("Coord", "項目", c0)
    sum_r = lab("Coord", "C0 摘要", hdr)
    lay = lab("Coord", "層級之間（參考", sum_r)
    lay_end = lab("Coord", "另一側讀數：情境 B HBM 市場", lay)
    cf = lambda r: ("sp1" if "差距" in str(CW.cell(r, 1).value) else "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) == "x" else "n2" if str(CW.cell(r, 2).value) == "EB" else "n1")

    def crows(r0, r1):
        out = []
        for r in range(r0, r1 + 1):
            a, bb = CW.cell(r, 1).value, CW.cell(r, 2).value
            if a is None:
                continue
            if bb is None and all(CW.cell(r, c).value is None for c in range(3, 7)):
                out.append({"sub": f"@Coord!A{r}"})
            else:
                out.append([f"@Coord!A{r}", f"@Coord!B{r}"] + [f"@Coord!{c}{r}|{cf(r)}" if not isinstance(CW[f"{c}{r}"].value, str) else f"@Coord!{c}{r}" for c in "CDEF"])
        return out
    note_rows = [r for r in range(hdr + 1, sum_r) if CW.cell(r, 7).value]
    import importlib.util
    import os
    _sp = importlib.util.spec_from_file_location("skh_c0_actuals", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "skhynix", "c0_actuals.py"))
    _m = importlib.util.module_from_spec(_sp)
    _sp.loader.exec_module(_m)
    act_blocks = _m.html_blocks(CW, lab)
    c0_blocks = act_blocks + [
                 {"type": "h3", "text": "讀取層級（同源，不判定）：本模型讀入的 Siliconomics 值"},
                 {"type": "note", "parts": [f"@Coord!B{sum_r}"]},
                 {"type": "table", "caption": "讀取層級對照表（本模型直接讀層，差距恆為零、無檢驗力，故不判定、不計入對帳中；公司實績對帳見上表）", "head": [f"@Coord!{c}{hdr}" for c in "ABCDEF"], "rows": crows(hdr + 1, sum_r - 1)},
                 {"type": "table", "caption": "對帳的口徑、來源與證據等級", "collapse": True, "summary": "展開：口徑、來源與證據等級", "head": ["項目", "口徑、來源與證據等級"],
                  "rows": [[f"@Coord!A{r}", f"@Coord!G{r}"] for r in note_rows]},
                 {"type": "h3", "text": "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳）"},
                 {"type": "note", "parts": [f"@Coord!B{lay_end + 1}"]},
                 {"type": "table", "caption": "另一側讀數對 HBM 出貨位元的影響（Wattonomics 與 Siliconomics 兩層之間的差距；本公司並列、不裁決）", "head": [f"@Coord!{c}{hdr}" for c in "ABCDEF"], "rows": crows(lay + 1, lay_end)},
                 {"type": "h3", "text": "與共識、公司指引的對照"}]
    S["recon"]["blocks"] = c0_blocks + S["recon"]["blocks"]
    S["recon"].pop("intro", None)
    # ---- 8. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rvdcf
    cfg["sections"] = [S[k] for k in order]
    return cfg
