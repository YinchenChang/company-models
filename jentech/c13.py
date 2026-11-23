# 電源散熱次產業（奇鋐、雙鴻、光寶、健策）版面 v1.3＋三層協調 C0（實績）共通後處理（2026-11-15）。
# 樣板：台達電 delta/v056_post.py、delta/v057_post.py（NVIDIA v0.3.10 版面）；各公司資料夾放同一份本檔，公司差異在 <資料夾>/v13_post.py 的設定。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對前一版逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025；PL 頁末：標準損益表（FY2023A–FY2031E）；RevDCF 頁（新）：反向 DCF；
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準；README 頁末：v1.3 白話列；
#   Coord 頁：三層協調 C0（讀取層級、讀取層級（同源，不判定）、層級之間（參考）、C0（實績））。
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
TOL = 0.15


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


def lab(ws, prefix, start=1, col=1, exact=False, end=None):
    for r in range(start, (end or ws.max_row) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def blockB(P):
    """PL 頁情境 B 推導區塊的標題列（A 欄以「情境 B」開頭、下一列為「項目」）。"""
    return next(r for r in range(1, P.max_row) if str(P.cell(r, 1).value or "").startswith("情境 B") and P.cell(r + 1, 1).value == "項目")


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- Hist：損益表歷史 ----------
# C["hist"]：(鍵, 項目, FY2022, FY2023, FY2024, FY2025, 標記, 來源)；鍵＝rev gp oi pti tax nci ni sh eps（oi、pti 等為公司財報值）
def hist_is(wb, C):
    H = wb.create_sheet("Hist")
    H["A1"] = C["hist_title"]; H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022（只供年增率）", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    R = {}
    r = 4
    for k, labl, v22, v23, v24, v25, tag, src in C["hist"]:
        R[k] = r
        H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if v is None:
                continue
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str):
                c.font = BLUE
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        H.cell(r, 7, tag); H.cell(r, 8, src)
        r += 1
    for k, labl, f, note in (("nonop", "營業外收益淨額（＝稅前 − 營業利益）", "={c}{pti}-{c}{oi}", "由財報稅前利益與營業利益推得"),
                             ("opx", "營業費用（＝毛利 − 營業利益）", "={c}{gp}-{c}{oi}", "由財報毛利與營業利益推得（含其他營業收支淨額）")):
        R[k] = r; H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl); H.cell(r, 7, "Derived"); H.cell(r, 8, note)
        for col, yy in (("C", "22"), ("D", "23"), ("E", "24"), ("F", "25")):
            if H[f"{col}{R['pti']}"].value is None or H[f"{col}{R['oi']}"].value is None:
                continue
            H[f"{col}{r}"] = f.format(c=col, **R); H[f"{col}{r}"].number_format = "#,##0.000"
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        r += 1
    R["chk"] = r
    H.cell(r, 1, "核對"); H.cell(r, 2, "稅前 − 所得稅 − 非控制權益 − 稅後淨利（應為 0；百萬元四捨五入差 < 0.005）")
    for col in "DEF":
        H[f"{col}{r}"] = f"={col}{R['pti']}-{col}{R['tax']}-{col}{R['nci']}-{col}{R['ni']}"; H[f"{col}{r}"].number_format = "0.000"
    r += 1
    H.cell(r, 1, "核對"); H.cell(r, 2, C.get("hist_chk2", "FY2025：本表營收、毛利、營業利益、淨利 − Q 頁錨點（Q 頁部分為推導值，差額為四捨五入或推導差）"))
    for col, k, nm in (("C", "rev", "a25_rev"), ("D", "gp", "a25_gp"), ("E", "oi", "a25_oi"), ("F", "ni", "a25_ni")):
        H[f"{col}{r}"] = f"=F{R[k]}-{nm}"; H[f"{col}{r}"].number_format = "0.000"
    R["chk2"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 46
    for c in "CDEF":
        H.column_dimensions[c].width = 16
    H.column_dimensions["H"].width = 110
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"


def pl_standard(wb, CR, C):
    P = wb["PL"]
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    B = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), nop=g(C.get("nop_label", "業外收益")), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"))
    if C.get("ext_rows"):   # 延伸期（FY2030–FY2031）在別處的模型（例：光寶科摘要表 H–J 欄）
        X = C["ext_rows"](P)
    else:
        x0 = lab(P, "延伸｜情境 B")
        gx = lambda label: lab(P, label, x0, exact=True, end=x0 + 14)
        X = dict(rev=gx("營收（延伸）"), ni=gx("淨利（延伸）"), sh=gx("股數（延伸）"), eps=gx("EPS（延伸）"), hdr=x0 + 1)
    D = wb["DCF"]
    dbr = lab(D, "情境 B")
    d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（{C.get('since', C['ver'])} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2031E（情境 B 基準；{C['fy26_short']}，FY2030E–FY2031E 為延伸期）。"
                 "NT$B；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, f"口徑：GAAP（台灣採用的 IFRS 合併財報，歸屬母公司淨利）。{C['name']}不另發布 Non-GAAP，法說與市場共識 EPS 皆以此口徑，故本表不做 Non-GAAP 調整。"
                 "FY2026E 以後的所得稅＝稅前 × 模型稅率（In 頁的稅率，與 DCF 同一個稅率），非控制權益＝稅前 − 所得稅 − 歸屬母公司淨利（模型以「淨利 ÷ 稅前」合併處理兩者）；歷史年度為公司財報值（Hist 頁）")
    r_note = r; r += 1
    P.cell(r, 1, f"FY2026E＝{C['fy26_long']}；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益由淨利還原，同 DCF 頁）："
                 "毛利率沿用 FY2029E 據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    fl = lambda y: f"Coord!${'BCDE'[y - 2026]}${CR['flag']}"
    for col, fy, cy in (("G", "FY2027E", 2027), ("H", "FY2028E", 2028), ("I", "FY2029E", 2029), ("J", "FY2030E", 2029), ("K", "FY2031E", 2029)):
        P[f"{col}{r}"] = f'=IF({fl(cy)}=1,"{fy}（對帳中）","{fy}")'
    r_hdr = r; r += 1
    L = {}
    M = "#,##0.0"; PCT = "0.0%"

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None:
                continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note:
            P.cell(r, 12, note)
        L[key] = r; r += 1
    eng = lambda key: [f"=PL!{c}{B[key]}" for c in "DEFG"]
    ext = lambda key: ([f"=PL!H{X[key]}", f"=PL!I{X[key]}"] if X.get(key) else [f"=PL!H{X['ni']}/PL!H{X['eps']}", f"=PL!I{X['ni']}/PL!I{X['eps']}"])   # 無延伸股數列者＝淨利 ÷ EPS
    hs = lambda k: [f"=h_is_{k}_23", f"=h_is_{k}_24", f"=h_is_{k}_25"]
    row("rev", "營收", "NT$B", hs("rev") + eng("rev") + ext("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_is_rev_22-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 取 Hist 頁")
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "CDEFGHI"] + [f"=J{L['rev']}*(1-$I{r_c + 2})", f"=K{L['rev']}*(1-$I{r_c + 2})"], M,
        "歷史與 FY2026E–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", hs("gp") + eng("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "NT$B", hs("opx") + eng("opx") + [None] * 2, M,
        "歷史＝毛利 − 營業利益（含其他營業收支）；FY2030E–FY2031E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）")
    row("oi", "營業利益", "NT$B", hs("oi") + eng("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E–FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利還原；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額，營業外收益）", "NT$B", hs("nonop") + eng("nop") + ["=nop_f", "=nop_f"], M, "FY2030E–FY2031E 同 FY2027E 起的年度業外收益（In 頁）")
    row("pti", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "NT$B", hs("tax") + [f"={c}{L['pti']}*tax_w" for c in "FGHIJK"], M, "FY2026E 起＝稅前 × 模型稅率（In 頁的稅率）")
    row("nci", "非控制權益損益", "NT$B", hs("nci") + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['pti'] + 3}" for c in "FGHIJK"], M, "FY2026E 起＝稅前 − 所得稅 − 稅後淨利（歸屬母公司）")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", hs("ni") + eng("ni") + ext("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數", "十億股", hs("sh") + eng("sh") + ext("sh"), "0.000")
    row("eps", "EPS（基本，公司公告）", "NT$", hs("eps") + eng("eps") + ext("eps"), "0.00", C["eps_note"])
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/h_is_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 稅後淨利取 Hist 頁")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（歷史與 FY2026E–FY2029E）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None, None], "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E–FY2031E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "JK":
        P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"I{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E–FY2031E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "NT$B", [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E–FY2031E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 6 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "IJK"], "0.00%")
    row("c3", "核對：稅後淨利 − 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "NT$B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{B['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    row("c4", "核對：稅前 − 所得稅 − 非控制權益 − 稅後淨利", "NT$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000",
        "歷史年度為財報百萬元四捨五入差")
    row("c5", "核對：延伸期稅前利益 × 淨利 ÷ 稅前比率（In 頁 nir_f）− 稅後淨利（FY2030E–FY2031E）", "NT$B", [None] * 7 + [f"={c}{L['pti']}*nir_f-{c}{L['ni']}" for c in "JK"], "0.000",
        "營業利益由 DCF 頁（淨利 ÷ nir_f − nop_f）還原，故稅前 × nir_f 應等於稅後淨利")
    row("c6", "核對：FY2025A 營業利益 − Q 頁錨點（財報值對 Q 頁推導值，非 0 表示 Q 頁為推導近似）", "NT$B", [None, None, f"=E{L['oi']}-a25_oi", None, None, None, None, None, None], "0.000",
        "Q 頁 FY2025 營業利益為錨點用值；本表取財報值（Hist 頁）")
    for col in COLS:
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    P.column_dimensions["L"].width = 70
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, B=B, X=X)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb, C):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    assert str(D.cell(fr, 1).value).startswith("評價日後占該年度比例") and str(D.cell(ty, 1).value).startswith("折現年數")
    V = wb["Val"]; P = wb["PL"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    rb = blockB(P); r26 = lab(P, "營收", rb, exact=True)
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（{C.get('since', C['ver'])} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                f"Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。{C['name']} FY2027 為模型推導年（非錨點），故營業利益率與營收倍數都從 FY2027 起動；"
                "反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "說明"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note:
            RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（NT$）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（NT$）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（NT$B；含 stub 期 FCFF 與已宣告股利）", f"={ncr}", "#,##0.0", "Val 頁 DCF 彙總「淨現金滾到評價日」情境 B 欄（負值＝債務與非控制權益大於現金與投資）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（NT$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；NT$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（NT$）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, "BCDEFG"[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（NT$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（NT$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻", lambda dc, c: f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（NT$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（NT$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（NT$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    kv("rev0", "FY2026 營收（NT$B；FY2027 營運資金增量的基期，不隨倍數變動）", f"=PL!D{r26}", "#,##0.0")
    kv("ev_chk", "核對：以本表各年重組的 EV − DCF 頁 EV（應為 0）", f"=SUMPRODUCT(B{Y['cf']}:G{Y['cf']}-B{Y['wc']}:G{Y['wc']},B{Y['fr']}:G{Y['fr']},B{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}-B{A['ev']}", "0.000000",
       "獨立驗證封閉解用到的各年係數（折現因子、評價日後比例、終值係數）")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2027–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（NT$B）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（NT$B）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2027–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(B{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(B{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2027–FY2032 營收等比例放大；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（NT$B）", f"=fc_wc*B{A['rev0']}*B{Y['fr']}*B{Y['df']}", "#,##0.0", "FY2027 營運資金增量＝fc_wc ×（s × FY2027 營收 − FY2026 營收）：其中 FY2026 營收那一項不隨倍數 s 變動")
    kv("b", "隨營收倍數等比例變動的 EV（NT$B）", f"=B{A['ev']}-B{A['a']}", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2027 營收（NT$B）", f"=B{A['s1']}*B{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2027 營收（情境 B；NT$B）", f"=B{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2027 營收年增率", f"=B{A['rv1']}/B{A['rev0']}-1", "0.0%")
    kv("g0", "模型：FY2027 營收年增率（情境 B）", f"=B{Y['rev']}/B{A['rev0']}-1", "0.0%")
    kv("rv2", "現價隱含的 FY2028 營收（NT$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv2m", "模型：FY2028 營收（情境 B；NT$B）", f"=C{Y['rev']}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（NT$）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C$8+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4):
            RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "#,##0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（NT$，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
    r += 1
    RV.cell(r, 1, "B. Comparable：支持現價所需的倍數").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；NT$）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"#,##0")',
                     T(" 元；要支持現價 "), 'TEXT(px,"#,##0")', T(" 元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大","縮小")', T("為 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<wacc,"降到 ","升到 ")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2027 起營收全面乘以 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv2"]},"#,##0")', T(" 十億元，模型為 "), f'TEXT(B{A["rv2m"]},"#,##0")',
                     T("）；或營業利益率改為 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"#,##0")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(f" 倍（同業中位數 × {C['name']}過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T(f" 倍）。換言之，市場目前給{C['name']}的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    return A


# ---------- Val：估值並列四個目標價（G 節） ----------
def val_g(wb, C):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    pxr = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, f"0.7 估值並列：四個目標價（{C.get('since', C['ver'])} 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "NT$", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（賣方平均）", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{pxr}"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（日期見左欄）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（H 節） ----------
TOR_LABEL = {"WACC ±1pt": "折現率（WACC）低／高一個百分點",
             "β 低／高": "股權風險係數低／高（各自重新加槓桿）",
             "ERP 低／高": "股權風險溢酬 4%／6%",
             "RONIC＝WACC＋5pt": "永續期新投資報酬率提高到折現率加五個百分點（只有高）",
             "DCF 權重 0.3／0.7": "加權目標價中 DCF 的權重 30%／70%（基準 50%）"}


def sens_tor(wb, CR, C):
    SE = wb["Sens"]; V = wb["Val"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    assert SE.cell(hr, 1).value == "參數"
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；{C.get('since', C['ver'])} 起，版面規格 v1.3 H 節；引用上方「目標價參數」表與 Val、Coord 頁，白話標籤，不回饋計算）").font = BOLD; t0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    for sr in src:
        t = str(SE.cell(sr, 1).value)
        key = next((k for k in TOR_LABEL if t.startswith(k)), None)
        SE.cell(r, 1, TOR_LABEL[key] if key else f"=A{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—":
                continue
            SE.cell(r, col, f"={scol}{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{sr}"); r += 1
    tp = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    cu = f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    SE.cell(r, 1, "同業比較法不加歷史倍數調整（相對倍數係數＝1；Comparable 取市場主流同業法）")
    SE.cell(r, 2, f"=(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/{tp}-1").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Val 頁：DCF 與未調整 Comparable 依 DCF 權重加權"); r += 1
    SE.cell(r, 1, "產業部署 GW 改用另一側讀數（Wattonomics 換算；Siliconomics 標對帳中，整本重建）")
    SE.cell(r, 2, f"=Coord!D{CR['tp']}").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Coord 頁兩讀數目標價（另一側相對現行）"); r += 1
    return dict(t0=t0, r1=first, r2=r - 1)


# ---------- README：v1.3 白話列 ----------
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


def readme13(wb, RVA, PLA, TOR, CR, C, lines):
    """lines：公司設定提供的 dict（hero、o2、o3、drv 等）；本函式補共通的 rec、o1、o4、o5 與版本列。"""
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）")
    cadj, cu = vb("Comparable 目標價（相對倍數調整後）"), vb("Comparable 目標價（市場主流同業法）")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    cr = lambda col: f"Coord!{col}{CR['tp']}"
    dw = lab(SE, "加權目標價中 DCF 的權重", TOR["t0"])
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；{C.get('since', C['ver'])} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", *lines["hero"](tp, up, rec, dcf, cadj))
    put("rec", T("投資建議「"), rec, T(f"」成立的前提：{lines['premise']}，DCF 與同業比較法各占一半。"
                  "DCF 權重降到三成時為"), REC5(f"((0.3*{dcf}+0.7*{cadj})/px-1)"), T("，同業比較法不加歷史倍數調整時為"), REC5(f"((tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/px-1)"),
        T("，產業部署量改用另一側讀數時為"), f"Coord!C{CR['rec']}", T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，"), f'IF({up}<0,"低於","高於")', T("現價 "), f'TEXT(px,"#,##0")', T(" 元 "), f'TEXT(ABS({up}),"0.0%")',
        T("，評等"), rec, T("；DCF "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")', T(" 元各占一半。"))
    put("o2", *lines["o2"])
    put("o3", *lines["o3"])
    put("o4", T("現價隱含：要支持現價，折現率須為 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或營收全面乘以 "), f'TEXT({rv("s1")},"0.00")', T(" 倍；DCF 權重 30%／70% 時加權目標價各 "),
        f'TEXT(Sens!B{dw},"+0%;-0%")', T("／"), f'TEXT(Sens!C{dw},"+0%;-0%")', T("。"))
    put("o5", T("對帳中：CY2027–CY2029 Siliconomics 標對帳中，另一側讀數（Wattonomics 換算）目標價 "), f'TEXT({cr("C")},"#,##0")', T(" 元（"), f"Coord!C{CR['rec']}",
        T("），較現行基準 "), f'TEXT(Coord!D{CR["tp"]},"+0.0%;-0.0%")', T(f"；{lines['o5_tail']}"))
    put("drv", *lines["drv"])
    RD.cell(r, 1, C["ver_line"]); out["ver"] = r; r += 1
    return out


# ---------- Coord：三層協調 C0（讀取層級、同源對照、層級之間） ----------
class Sheet:
    """Coord 頁逐列寫入的小工具（第 4 節 C0（實績）用）。"""
    def __init__(self, ws, r):
        self.C, self.r = ws, r

    def bold(self, text):
        self.C.cell(self.r, 1, text).font = BOLD; self.r += 1; return self.r - 1

    def text(self, text):
        self.C.cell(self.r, 1, text); self.r += 1; return self.r - 1

    def hdr(self, cols):
        for j, h in enumerate(cols, 1):
            c = self.C.cell(self.r, j, h); c.font = BOLD; c.fill = HDR
        self.r += 1; return self.r - 1

    def row(self, label, b=None, c=None, unit="", note="", fmt="#,##0.000", blue=False, cfmt=None, url=""):
        r = self.r
        self.C.cell(r, 1, label)
        for col, v, f in ((2, b, fmt), (3, c, cfmt or fmt)):
            if v is None:
                continue
            cell = self.C.cell(r, col, v); cell.number_format = f
            if blue and not (isinstance(v, str) and v.startswith("=")):
                cell.font = BLUE
        if unit:
            self.C.cell(r, 6, unit)
        if note:
            self.C.cell(r, 7, note)
        if url:
            self.C.cell(r, 8, url)
        self.r += 1
        return r

    def skip(self, n=1):
        self.r += n


def judge(rg):
    return (f'=IF(AND(ABS(B{rg})>{TOL},ABS(C{rg})>{TOL},B{rg}*C{rg}>0),"對帳中",'
            f'IF(OR(ABS(B{rg})>{TOL},ABS(C{rg})>{TOL}),"時差邊界（參考）","一致"))')


def coord_c0(wb, SIL, CR, C):
    """C["layer_rows"]：讀取層級表（label, value 公式, 層級名稱, 用途, fmt）；C["same_blocks"]：同源對照（key, title, unit, PL 標籤, 層級值 list, note_m, note_l, fmt）。"""
    CW = wb["Coord"]; P = wb["PL"]
    rb = blockB(P)
    pb = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調 C0（三層基準協調規則 v1.4.2 C0；工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）：讀哪一層、哪一版；讀取層級（同源，不判定）；C0（實績）").font = BOLD; G = {"c0": r}; r += 1
    CW.cell(r, 1, "原則：同一數字只有一個權威層，下游註明版本、不自行重估（Project 規則第 3 條）；C0 只用公司實績對層級隱含值判定，兩讀數都超過 ±15% 且同號才標「對帳中」，不改上游值，由主控開對帳單；個股不裁決（C4），往上回饋只有證據（C5）。").font = Font(italic=True)
    r += 2
    CW.cell(r, 1, "1. 本模型讀哪一層、哪一版（硬體個股讀 Siliconomics；Wattonomics 只經 Siliconomics 的 R4 比值與另一側讀數進入，現行版本 v0.3.2）").font = BOLD; G["s1"] = r; r += 1
    for j, h in enumerate(["項目", "值", "層級名稱（版本）", "本模型用途"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    rows = [("Siliconomics 快照版本", SIL.VERSION, ver, "行業變數單一來源（siliconomics/reader.py）；下游不讀 In、Path、Mkt 等內部頁", None)] + C["layer_rows"](ver, CR) + [
        ("Wattonomics 錨定係數 CY2027／28／29（SIF_WTAnchor_Base；R4 比值）", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")',
         f"{ver}；Wattonomics 現行 v0.3.6（只經 Siliconomics 進入）", "只進敏感度（Sens「Wattonomics 錨定」案例）；不改基準", None),
        ("另一側讀數 ÷ 現行基準 CY2026／27／28／29（SIF_DeployGW_AltRatio_<年>）", f'=TEXT(B{CR["ratio"]},"0.000")&"／"&TEXT(C{CR["ratio"]},"0.000")&"／"&TEXT(D{CR["ratio"]},"0.000")&"／"&TEXT(E{CR["ratio"]},"0.000")',
         f"{ver}（比值列見本頁上方）", "兩讀數目標價（本頁上方）只換產業部署 GW 一個輸入", None)]
    for label, val, who, use, fmt in rows:
        CW.cell(r, 1, label); c = CW.cell(r, 2, val)
        if fmt:
            c.number_format = fmt
        CW.cell(r, 3, who); CW.cell(r, 4, use); r += 1
    G["s1_end"] = r - 1
    r += 1
    CW.cell(r, 1, f"2. 讀取層級（同源，不判定）：本模型自己的產業量與 Siliconomics 層級值逐年對照。個股的量直接讀自 Siliconomics（或份額由本公司模型校準後寫入 Siliconomics），差距恆為 0 或只反映本模型校準，無檢驗力，"
                  "不計入 C0 對帳中項目；C0 判定見下方第 4 節「C0（實績）」").font = BOLD; G["s2"] = r; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位", "口徑、來源、證據等級"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def block(key, title, unit, model, layer, note_m, note_l, fmt, verdict=SAME, tag_title="判定（同源，不判定）", ncols=4):
        nonlocal r
        CW.cell(r, 1, f"{title}：本模型（情境 B；PL 頁）"); CW.cell(r, 6, unit); CW.cell(r, 7, note_m)
        for j, f in enumerate(model[:ncols]):
            CW.cell(r, 2 + j, f).number_format = fmt
        rm = r; r += 1
        CW.cell(r, 1, f"{title}：層級值"); CW.cell(r, 6, unit); CW.cell(r, 7, note_l)
        for j, v in enumerate(layer[:ncols]):
            c = CW.cell(r, 2 + j, v); c.number_format = fmt
            if not (isinstance(v, str) and v.startswith("=")):
                c.font = BLUE
        rl = r; r += 1
        CW.cell(r, 1, f"{title}：差距（本模型 ÷ 層級值 − 1）")
        for j, col in enumerate("BCDE"[:ncols]):
            CW.cell(r, 2 + j, f'=IF({col}{rl}=0,"—",{col}{rm}/{col}{rl}-1)').number_format = "+0.0%;-0.0%"
        rg = r; r += 1
        CW.cell(r, 1, f"{title}：{tag_title}")
        for j, col in enumerate("BCDE"[:ncols]):
            CW.cell(r, 2 + j, verdict if not verdict.startswith("=") else verdict.format(col=col, rg=rg))
        rj = r; r += 1
        G[key] = (rm, rl, rg, rj)
    pc = lambda label: [f"=PL!{c}{pb(label)}" for c in "DEFG"]
    SB = C["same_blocks"](SIL, pc)
    for key, title, unit, pl_label, layer, note_m, note_l, fmt in SB:
        block(key, title, unit, pl_label if isinstance(pl_label, list) else pc(pl_label), layer, note_m, note_l, fmt)
    G["same_keys"] = [b[0] for b in SB]
    if C.get("resid"):   # C0 恆等式：已建模個股份額加總之外的殘差（須明列且不為空）
        k, note = C["resid"]
        CW.cell(r, 1, "C0 恆等式：殘差（其他廠商＝1 − 已建模個股加總；須明列且不為空）"); CW.cell(r, 6, "%"); CW.cell(r, 7, note)
        for j, col in enumerate("BCDE"):
            CW.cell(r, 2 + j, f"=1-{col}{G[k][0]}").number_format = "0.0%"
        G["resid"] = r; r += 1
    # 層級之間（參考）
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：Wattonomics 側讀數").font = BOLD; r += 1
    alt = [SIL.NAMES[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)]
    if C.get("gw_label"):
      block("alt", "部署 GW 對 Wattonomics 側讀數", "GW", pc(C.get("gw_label", "部署 GW（IT 電力）")), alt,
          "本模型情境 B（讀 Siliconomics 基準 × 情境顆數係數）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Alt_<年>（Wattonomics 換算值；需求側讀數）", "0.00",
          verdict='=IF(ABS({col}{rg})>0.15,"參考：層級之間（Siliconomics 與 Wattonomics 兩層之間的差距，已並列兩讀數）","一致")', tag_title="判定（±15%；僅參考）")
    r += 1
    CW.cell(r, 1, "3. 讀取層級與層級之間的結果（HTML 用；公式帶數字）").font = BOLD; G["s3"] = r; r += 1
    keys = G["same_keys"]
    mx = lambda k: "MAX(" + ",".join(f"IFERROR(ABS({c}{G[k][2]}),0)" for c in "BCDE") + ")"
    parts = [T(f"本模型讀 Siliconomics {SIL.VERSION}：{C['same_names']}與層級值的最大差距分別為 ")]
    for i, k in enumerate(keys):
        parts += ([T("、")] if i else []) + [f'TEXT({mx(k)},"0.0%")']
    parts += [T("（同源：本模型直接讀入或由本公司校準後寫入 Siliconomics，差距只反映讀取一致，不判定）。")]
    CW.cell(r, 1, F(*parts)); G["sum1"] = r; r += 1
    if C.get("gw_label"):
        CW.cell(r, 1, F(T("層級之間（參考）：本模型（Siliconomics 基準）的部署 GW 在 CY2027–CY2029 比 Wattonomics 換算值高 "), f'TEXT(C{G["alt"][2]},"0%")', T("／"), f'TEXT(D{G["alt"][2]},"0%")', T("／"),
                         f'TEXT(E{G["alt"][2]},"0%")', T("；這是 Siliconomics 與 Wattonomics 兩層之間的差距（Siliconomics 對帳單 r3 處理；已並列兩讀數目標價），"
                                                        "不是本模型自己的差距，不計入本公司 C0，本模型不改上游值、不裁決。")))
    else:   # 本模型不讀部署 GW（以加速器顆數推導）：層級之間只列 Siliconomics 的另一側讀數比值
        CW.cell(r, 1, F(T("層級之間（參考）：本模型不讀部署 GW（以加速器顆數推導），兩讀數只換產業量；Siliconomics 部署 GW 的 Wattonomics 換算值 ÷ 現行基準 CY2027–CY2029 為 "),
                         f'TEXT(C{CR["ratio"]},"0.00")', T("／"), f'TEXT(D{CR["ratio"]},"0.00")', T("／"), f'TEXT(E{CR["ratio"]},"0.00")',
                         T("；這是 Siliconomics 與 Wattonomics 兩層之間的差距（Siliconomics 對帳單 r3 處理；已並列兩讀數目標價），不計入本公司 C0，本模型不改上游值、不裁決。")))
    G["sum2"] = r; r += 1
    CW.cell(r, 1, "C4／C5：本模型不裁決兩層的對帳中年份；本版沒有可往上回饋的新證據（往上回饋只有實績、法定文件揭露的承諾、已簽約或已確認站點，個股預測與公司指引不往上流）。"); G["sum3"] = r; r += 1
    CW.column_dimensions["A"].width = 70; CW.column_dimensions["D"].width = 24; CW.column_dimensions["F"].width = 10; CW.column_dimensions["G"].width = 110; CW.column_dimensions["H"].width = 60
    G["end"] = r - 1
    return G


def version_a1(wb, C):
    """README A1：標題段版號更正為新版，末尾加本版說明。"""
    RD = wb["README"]
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(r"v\d+(?:\.\d+)+（", f"{C['ver']}（", head, count=1)
    assert n == 1, head
    a1 = re.sub("\u300020\\d\\d-\\d\\d-\\d\\d", "\u3000" + C["built"], head2 + sep + rest, count=1)
    RD["A1"].value = a1 + "；" + C["a1_tail"]


def apply(wb, SIL, CR, C, lines_fn, c0_fn):
    HR = hist_is(wb, C)
    PLA = pl_standard(wb, CR, C)
    RVA = revdcf(wb, C)
    VG = val_g(wb, C)
    ST = sens_tor(wb, CR, C)
    G = coord_c0(wb, SIL, CR, C)
    R4 = c0_fn(wb, G)
    RMA = readme13(wb, RVA, PLA, ST, CR, C, lines_fn(wb))
    version_a1(wb, C)
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, G=G, R4=R4, RM=RMA)
    return (f"v13_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {ST['t0']} 列；"
            f"Coord C0 第 {G['c0']} 列起、C0（實績）第 {R4['title']}–{R4['end']} 列；README v1.3 白話列 {len(RMA)} 列")
