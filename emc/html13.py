# 電源散熱次產業（奇鋐、雙鴻、光寶、健策）HTML 版面 v1.3＋三層協調 C0（實績）共通步驟（2026-11-15）。
# 樣板：台達電 delta/gen_html_json.py 的 _v13、_v13_post（NVIDIA v0.3.10 版面）。2026-11-15 起高速 CCL／PCB 四家共用（標準損益表 EPS 單位 RMB 也取兩位小數）。各公司資料夾放同一份本檔；公司差異（命題圖、分部推導表、部署圖、白話替換）
# 由 <資料夾>/gen_html_json.py 傳入 H 設定。數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF、Val 0.7、Sens 末表、Coord 頁）。
import os, sys


def apply(cfg, wb, X, H, root):
    def row(sh, label, start=1, exact=False):
        ws = wb[sh]
        for r in range(start, ws.max_row + 1):
            v = str(ws.cell(r, 1).value or "").strip()
            if (v == label.strip()) if exact else v.startswith(label.strip()):
                return r
        raise KeyError((sh, label))

    def until_blank(sh, r0):
        ws, r = wb[sh], r0
        while ws.cell(r, 1).value not in (None, ""):
            r += 1
        return r - 1

    S = {s["id"]: s for s in cfg["sections"]}
    RD, PLW, RV = wb["README"], wb["PL"], wb["RevDCF"]
    r13 = row("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith(H["ver"]), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    rbB = next(r for r in range(1, PLW.max_row) if str(PLW.cell(r, 1).value or "").startswith("情境 B") and PLW.cell(r + 1, 1).value == "項目")
    hB = rbB + 1
    assert PLW.cell(hB, 1).value == "項目"
    pB = lambda lab, ex=True: row("PL", lab, rbB, ex)
    # ---- 1. 命題 ----
    th = S["thesis"]["blocks"]
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    i_cmd = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "命題")
    th.insert(i_cmd + 2, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    th.insert(0, H["units_chart"](pB, hB))
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與口徑":
            b["text"] = "情境定義"
    # ---- 2. 估值：四個目標價（Val 0.7）----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars" and b.get("title") == "目標價與現價")
    g0 = row("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣元；豎線＝現價）", "fmt": "n0",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    t07 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
           "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n0", "", f"@Val!D{r_px}"]]}
    va[i_bar] = bars4
    va.insert(i_bar, t07)
    va.insert(i_bar + 2, {"type": "note", "parts": old_note})
    # ---- 4. 反向 DCF ----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2027 營收（", rbb), rv("模型：FY2027 營收（", rbb)
    g1, g0_ = rv("現價隱含的 FY2027 營收年增率", rbb), rv("模型：FY2027 營收年增率", rbb)
    w1 = rv("現價隱含的 WACC", rc)
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2027–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2027 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2027 營收（十億元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2027 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w1 + 1}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if r in (ra + 2, ra + 3, ra + 9) else "n3" if r == ra + 5 else "x1" if r == ra + 8 else "n1")]
                      for r in range(ra + 2, ra + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]
                     + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "利益率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 4)]
                     + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                     + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
        ]}
    # ---- 5. 損益（標準損益表；GAAP）----
    p0 = row("PL", "標準損益表（")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: row("PL", lab_, start, True)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率", "其他收入（淨額，營業外收益）",
            "稅前利益", "所得稅", "非控制權益損益", "稅後淨利（歸屬母公司）", "淨利率", "股數", "EPS（基本，公司公告）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "RMB": "n2", "十億股": "n3"}.get(u, "n1")
    cols = "CDEFGHIJK"
    rowx = lambda r, f=None: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{f or fmt(PLW.cell(r, 2).value)}" for c in cols]
    r_ratio, r_chk = plr("比率"), plr("核對（應為 0）")
    r_d = row("PL", "核對：延伸期毛利", ph)
    is_tbl = {"type": "table", "caption": "標準損益表（GAAP 口徑，十億元新台幣；EPS 為新台幣元）：FY2023A–FY2025A 實績、FY2026E 錨點年、FY2027E–FY2031E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]
                      + [{"sub": "核對與揭露：FY2030E–FY2031E 毛利 − 營業費用 − 營業利益（應為零；延伸期營業費用＝毛利 − 營業利益，是殘差不是假設）與隱含營業費用率"}] + [rowx(r) for r in (r_d, r_d + 1)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols], "collapse": True,
               "summary": "展開：標準損益表的核對列（恆等式與對情境 B 摘要的差額，應為零）",
               "rows": [rowx(r, "n3") for r in range(r_chk + 1, until_blank("PL", r_chk + 1) + 1) if r not in (r_d, r_d + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    xl = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
    is_chart = {"type": "theme", "title": "營收（十億元）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": xl,
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸）。FY2026E 起為模型（含已公布實績與第四季假設）；FY2030E、FY2031E 為延伸期，營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2029。FY2027E 起標「對帳中」。"}
    if H.get("ext"):   # (表頭列, {情境: (營收列, EPS 列)})
        xh, XR = H["ext"](row)
    else:
        x0 = row("PL", "延伸：三情境到 FY2032")
        xa, xb, xc = (row("PL", f"延伸｜情境 {k}", x0) for k in "ABC")
        xh = xa + 1
        XR = {k: (row("PL", "營收（延伸）", b, True), row("PL", "EPS（延伸）", b, True)) for k, b in zip("ABC", (xa, xb, xc))}
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億元）與 EPS（折線，新台幣元，右軸）", "x": f"@PL!D{xh}:J{xh}",
                "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.15,
                "bars": [{"name": ["情境 A（樂觀）營收"], "y": f"@PL!D{XR['A'][0]}:J{XR['A'][0]}", "color": "#b7e3e4"},
                         {"name": ["情境 B（基準）營收"], "y": f"@PL!D{XR['B'][0]}:J{XR['B'][0]}", "color": "#10a0a7"},
                         {"name": ["情境 C（保守）營收"], "y": f"@PL!D{XR['C'][0]}:J{XR['C'][0]}", "color": "#8a8c8c"}],
                "lines": [{"name": ["情境 A（樂觀）EPS"], "y": f"@PL!D{XR['A'][1]}:J{XR['A'][1]}", "axis": "r", "color": "#0b7a80", "dash": True},
                          {"name": ["情境 B（基準）EPS"], "y": f"@PL!D{XR['B'][1]}:J{XR['B'][1]}", "axis": "r", "color": "#3b3838", "labels": True},
                          {"name": ["情境 C（保守）EPS"], "y": f"@PL!D{XR['C'][1]}:J{XR['C'][1]}", "axis": "r", "color": "#b4472f", "dash": True}],
                "caption": H["sc_caption"]}
    old_pl = S["pl"]["blocks"]
    old_intro = S["pl"].pop("intro", None)
    pl_old = [b for b in old_pl if b["type"] == "pl"]
    assert len(pl_old) == 2, len(pl_old)
    old_blocks = [dict(pl_old[0], collapse=True, summary="展開：三情境摘要（原版面）")]
    if old_intro:
        old_blocks.append({"type": "note", "parts": [old_intro], "collapse": True, "summary": "展開：三情境摘要的會計年度說明"})
    old_blocks.append(dict(pl_old[1], collapse=True, summary="展開：三情境摘要（原版面，續）"))
    pl_blocks = [{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, chk_tbl, sc_chart] + old_blocks
    S["pl"].update(title="損益（標準損益表）", blocks=pl_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    uf = lambda u: {"US$B": "n1", "GW": "n1", "US$/kW": "n0", "%": "p1", "NT$B": "n1", "x": "n2", "NT$B/GW": "n2", "US$M/GW": "n0", "US$": "n0", "M 顆": "n2", "百萬顆": "n2"}.get(u, "n1")
    segc = "CDEFG"
    seg_rows = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{uf(PLW.cell(r, 2).value)}" for c in segc] for r in rs if PLW.cell(r, 1).value and not str(PLW.cell(r, 1).value).startswith("檢查")]
    rows_ = []
    for sub, a_lab, b_lab in H["seg"]:
        ra_, rb2 = pB(a_lab, False), pB(b_lab, False)
        rows_ += [{"sub": sub}] + seg_rows(range(ra_ + 1, rb2))
    seg_tbl = {"type": "table", "caption": H["seg_caption"], "head": [f"@PL!A{hB}", f"@PL!B{hB}"] + [f"@PL!{c}{hB}" for c in segc], "rows": rows_}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, H["gw_chart"](pB, hB), seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 9. 資料狀態 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    cfg["sections"].append(rev_)

    # 三層協調 C4：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
    sys.path.insert(0, os.path.join(root, "tools", "coord"))
    import coord as COORD  # noqa: E402
    COORD.inject_html(cfg, X)

    # ---- v1.3 後處理：估值與敏感度主圖置頂、敏感度加權目標價龍捲風圖、對帳章加 C0 ----
    S = {s["id"]: s for s in cfg["sections"]}
    SW = wb["Sens"]
    t0 = row("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    out = []
    for col, side in (("B", "低"), ("C", "高")):
        num = lambda r: isinstance(SW[f"{col}{r}"].value, (int, float)) and not isinstance(SW[f"{col}{r}"].value, bool)
        allr = sorted([r for r in range(r1, r2 + 1) if num(r) and abs(SW[f"{col}{r}"].value) >= 0.0005], key=lambda r: -abs(SW[f"{col}{r}"].value))
        big = [r for r in allr if abs(SW[f"{col}{r}"].value) >= 0.1]
        if len(big) >= 2:
            ttl = f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例"
        else:   # 沒有兩個以上達兩位數者：主圖放變動最大的五個案例
            big = allr[:5]
            ttl = f"加權目標價相對基準：輸入取{side}值，變動最大的案例"
        small = [r for r in allr if r not in big]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        out.append(dict(base, title=ttl, rows=",".join(map(str, sorted(big)))))
        if small:
            out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, sorted(small))), collapse=True,
                            summary=f"展開：加權目標價相對基準：輸入取{side}值，其餘案例"))
    out.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]})
    sb = S["sens"]["blocks"]
    for b in sb:
        if b.get("type") == "tornado":
            b["collapse"] = True
            b["summary"] = "展開：對照：FY2028 EPS 敏感度（一次動一個輸入）" + ("，變動未達兩位數百分比的案例" if "未達" in str(b.get("title")) else "")
        if b.get("type") == "table" and not b.get("collapse") and str(b.get("caption", "")).startswith("@Sens!A"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（折現率、股權風險、永續成長率、DCF 權重等）"
        if b.get("type") == "note" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：敏感度頁原文說明"
    S["sens"]["blocks"] = out + sb
    S["sens"]["intro"] = "主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入）；下方摺疊為其餘案例、全部案例、目標價參數與 FY2028 EPS 敏感度對照。"
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar - 1, i_bar + 1
    assert va[i_tbl].get("type") == "table" and str(va[i_tbl].get("caption", "")).startswith("@Val!A"), va[i_tbl]
    assert va[i_nt].get("type") == "note" and "投資建議依加權目標價判定" in str(va[i_nt].get("parts")), va[i_nt]
    top = [va[i_bar]]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [va[i_tbl], va[i_nt]]
    S["valuation"]["blocks"] = top + [b for i, b in enumerate(va) if i not in (i_tbl, i_bar, i_nt)]
    sb = S["sens"]["blocks"]
    main = [i for i, b in enumerate(sb) if b.get("type") == "tornado" and not b.get("collapse") and str(b.get("title", "")).startswith("加權目標價相對基準")]
    assert len(main) == 2, main
    top = [sb[i] for i in main]
    if S["sens"].get("intro"):
        top.append({"type": "note", "parts": [S["sens"].pop("intro")]})
    S["sens"]["blocks"] = top + [b for i, b in enumerate(sb) if i not in main]
    # 對帳章：三層協調 C0（Coord 頁）
    CW = wb["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    c0, s1, s2, s3 = cr("三層協調 C0"), cr("1. 本模型讀哪一層"), cr("2. 讀取層級"), cr("3. 讀取層級與層級之間的結果")

    def f2(r):
        lab, u = str(CW.cell(r, 1).value), str(CW.cell(r, 6).value or "")
        if "差距" in lab:
            return "sp1"
        if "判定" in lab:
            return ""
        return {"GW": "n1", "US$B": "n1", "%": "p1", "US$/kW": "n0", "US$M/GW": "n0"}.get(u, "n2")
    c0t = {"type": "table", "caption": "讀取層級（同源，不判定）：本模型的產業量讀自 Siliconomics（份額由本公司模型校準後寫入），與層級值逐年對照的差距無檢驗力，不計入對帳中；下方「層級之間」列僅供參考",
           "head": [f"@Coord!{c}{s2 + 1}" for c in "ABCDE"],
           "rows": [({"sub": f"@Coord!A{r}"} if CW.cell(r, 2).value is None else [f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{f2(r)}" for c in "BCDE"]) for r in range(s2 + 2, s3 - 1) if CW.cell(r, 1).value]}
    c0d = {"type": "table", "caption": f"@Coord!A{s2}", "collapse": True, "summary": "展開：各列的單位、口徑、來源與證據等級", "head": [f"@Coord!{c}{s2 + 1}" for c in "AFG"],
           "rows": [[f"@Coord!A{r}", f"@Coord!F{r}|", f"@Coord!G{r}|"] for r in range(s2 + 2, s3 - 1) if CW.cell(r, 6).value or CW.cell(r, 7).value]}
    lay = {"type": "table", "caption": f"@Coord!A{s1}", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版、用在哪裡", "head": [f"@Coord!{c}{s1 + 1}" for c in "ABCD"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|", f"@Coord!C{r}|", f"@Coord!D{r}|"] for r in range(s1 + 2, s2 - 1) if CW.cell(r, 1).value]}
    s4 = cr("4. C0（實績）")
    t4 = cr("C0（實績）：公司實績 vs 層級隱含值", s4)
    e4 = cr("C0（實績）結果摘要", t4)
    a4 = cr("4a. 公司已公布實績", s4)
    b4 = cr("4b. 層級隱含值", a4)
    isnum = lambda v: isinstance(v, (int, float)) or str(v).startswith("=")
    fm4 = lambda r: ("" if any(k in str(CW.cell(r, 1).value) for k in ("判定", "結果")) else "sp1" if "差距" in str(CW.cell(r, 1).value) or "敏感度" in str(CW.cell(r, 1).value) else
                     "p1" if CW.cell(r, 6).value == "%" else "n1")
    c0a = {"type": "table", "caption": f"@Coord!A{t4}", "head": [f"@Coord!{c}{t4 + 1}" for c in "ABC"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{fm4(r)}" if isnum(CW.cell(r, 2).value) else f"@Coord!B{r}|",
                     f"@Coord!C{r}|{fm4(r)}" if isnum(CW.cell(r, 3).value) else f"@Coord!C{r}|"]
                    for r in range(t4 + 2, e4 - 1) if CW.cell(r, 1).value and "：" in str(CW.cell(r, 1).value) and (CW.cell(r, 2).value is not None or CW.cell(r, 3).value is not None)
                    and not str(CW.cell(r, 1).value).startswith(("參考：本模型", "FY2026 營收錨點", "錨點隱含", "隱含月均"))]}
    c0in = {"type": "table", "caption": f"@Coord!A{a4}", "collapse": True, "summary": "展開：公司已公布實績與揭露（月營收、季報）的數值、來源與證據等級", "head": [f"@Coord!{c}{a4 + 1}" for c in "ABFG"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{'p1' if CW.cell(r, 6).value == '%' else 'n3' if CW.cell(r, 6).value != '月' else 'n0'}", f"@Coord!F{r}|", f"@Coord!G{r}|"]
                     for r in range(a4 + 2, b4 - 1) if CW.cell(r, 1).value]}
    c0lay = {"type": "table", "caption": f"@Coord!A{b4}", "collapse": True, "summary": "展開：層級隱含值（Siliconomics 現行版）的組成", "head": [f"@Coord!{c}{b4 + 1}" for c in "ABFG"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{'p1' if CW.cell(r, 6).value == '%' else 'n2'}", f"@Coord!F{r}|", f"@Coord!G{r}|"]
                      for r in range(b4 + 2, t4) if CW.cell(r, 1).value and CW.cell(r, 2).value is not None]}
    c0det = {"type": "table", "caption": "公司實績對照各列的單位、口徑、來源與證據等級", "collapse": True, "summary": "展開：公司實績對照各列的單位、口徑、來源與證據等級", "head": [f"@Coord!{c}{t4 + 1}" for c in "AFG"],
             "rows": [[f"@Coord!A{r}", f"@Coord!F{r}|", f"@Coord!G{r}|"] for r in range(t4 + 2, e4 - 1) if CW.cell(r, 7).value]}
    sums = [r for r in range(e4 + 1, CW.max_row + 1) if CW.cell(r, 1).value]
    c0b = [{"type": "h3", "text": "三層協調：公司實績對層級隱含值"}, {"type": "note", "parts": [f"@Coord!A{s4 + 1}"]}]
    c0b += [{"type": "note", "parts": [f"@Coord!A{r}"]} for r in sums]
    c0b += [c0a, c0det, c0in, c0lay, {"type": "note", "parts": [f"@Coord!A{s4 + 2}"]},
            {"type": "h3", "text": "讀取層級（同源，不判定）與層級之間（參考）"}, {"type": "note", "parts": [f"@Coord!A{c0 + 1}"]}]
    c0b += [{"type": "note", "parts": [f"@Coord!A{r}"]} for r in range(s3 + 1, s3 + 4) if CW.cell(r, 1).value]
    c0b += [c0t, c0d, lay, {"type": "h3", "text": H["recon_h3"]}]
    S["recon"]["blocks"] = c0b + S["recon"]["blocks"]
    # 全頁可見文字（摺疊與資料狀態以外）不出現參數代號：含代號的儲存格改用白話（設定檔文字，不含數字）；xlsx 原文不改
    PLAIN = H.get("plain", {})

    def _plain(spec):
        if isinstance(spec, list):
            return [_plain(x) for x in spec]
        if isinstance(spec, str) and spec.startswith("@"):
            ref = spec[1:].split("|")[0]
            if ref in PLAIN:
                return PLAIN[ref]
        return spec
    for sc in cfg["sections"]:
        if sc["id"] == "status":
            continue
        FOLD = H.get("fold", {})
        for i, b in enumerate(sc["blocks"]):
            if b.get("type") == "note" and not b.get("collapse") and len(b.get("parts", [])) == 1 and str(b["parts"][0])[1:] in FOLD:
                sc["blocks"][i] = b = dict(b, collapse=True, summary=FOLD[str(b["parts"][0])[1:]])
            if b.get("type") == "kv":
                b["items"] = [_plain(it) for it in b.get("items", [])]
            if b.get("type") == "table":
                b["rows"] = [({"sub": _plain(r["sub"])} if isinstance(r, dict) else _plain(r)) for r in b.get("rows", [])]
                if "head" in b:
                    b["head"] = _plain(b["head"])
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]
    return RM
