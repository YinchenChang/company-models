# 台達電 v0.5.6（2026-11-14）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7／v0.3.8；Andy 2026-11-14 16:57 退件後全面改版）
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.5.5 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司法說簡報；標準損益表的歷史三年與年增率）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B；口徑＝GAAP〔台灣採用 IFRS，歸屬母公司〕）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末：0.7 估值並列四個目標價
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用）
#   README 頁末：v1.3 展示層白話列
#   Coord 頁：新增「三層協調 C0」一節（讀哪一層、哪一版；由下往上對帳）與 co_flag_<年> 具名範圍（鏡像 Siliconomics 對帳中旗標）
# build_delta.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, here, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。
import os
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.5.6", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
D4Q25 = "https://filecenter.deltaww.com/ir/download/calendar/4Q25_Analyst Meeting.pdf"
D4Q24 = "https://filecenter.deltaww.com/ir/download/calendar/4Q24_Analyst Meeting.pdf"
D4Q23 = "https://filecenter.deltaww.com/ir/download/calendar/4Q23_Analyst Meeting.pdf"
SA = "https://stockanalysis.com/quote/tpe/2308/financials/"


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


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- Hist：損益表歷史 ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025, 標記)；金額 NT$B、EPS NT$、股數十億股
def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.5.6 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
               "來源：台達電 4Q23／4Q24／4Q25 法說簡報（損益表：營業外收支、稅前、所得稅、非控制權益、稅後淨利、EPS 為百萬元原值；營收、毛利為十億元四捨五入）、"
               "毛利取 stockanalysis 年度毛利（百萬元；與法說簡報四捨五入值一致）。營業利益＝稅前 − 營業外（法說簡報營業利益只給四捨五入值）；營業費用＝毛利 − 營業利益")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022（只供年增率）", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    rows = [
        ("rev", "營收", 384.4, 401.227, 421.148, "=a25_rev", "Verified", f"FY2023 營收＝月營收公告累計 4,012.27 億元（TechNews 2024-01-10）；FY2022 384.4（四捨五入，{D4Q23}）；FY2024 與 FY2025 同 Q 頁／法說簡報"),
        ("gp", "毛利", None, 117.213, 136.580, 190.157, "Verified", f"{SA}（2026-11-14 讀取；百萬元）；法說簡報四捨五入值 117.2／136.6／190.2 一致（{D4Q24}；{D4Q25}）"),
        ("nonop", "營業外收益淨額", None, 7.692, 3.664, 3.934, "Verified", f"{D4Q24}（FY2023、FY2024）；{D4Q25}（FY2025）：營業外收支淨額，百萬元原值"),
        ("pti", "稅前利益", None, 48.642, 51.316, 87.866, "Verified", f"同上：Income before tax，百萬元原值"),
        ("tax", "所得稅", None, 9.762, 10.924, 19.930, "Verified", "同上：Income tax expense，百萬元原值"),
        ("nci", "非控制權益損益", None, 5.487, 5.163, 7.828, "Verified", "同上：Non-controlling interest，百萬元原值"),
        ("ni", "稅後淨利（歸屬母公司）", 32.666, 33.393, 35.229, "=a25_ni", "Verified", f"同上：Net profit after tax（歸屬母公司；FY2022 取 {D4Q23}）；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", None, 2.598, 2.598, 2.598, "Verified", "法說簡報註：EPS 以 2,598 百萬股計"),
        ("eps", "EPS（公司公告，NT$）", 12.58, 12.86, 13.56, 23.14, "Verified", "同上：EPS（FY2025 同 Q 頁 23.14）；資料商稀釋 EPS 為 12.80／13.51／23.08，差異來自股數口徑，採公司公告"),
    ]
    r = 4
    R = {}
    for k, labl, v22, v23, v24, v25, tag, src in rows:
        R[k] = r; r += 1
    for k, labl, v22, v23, v24, v25, tag, src in rows:
        rr = R[k]
        H.cell(rr, 1, f"h_is_{k}"); H.cell(rr, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if v is None:
                continue
            c = H[f"{col}{rr}"]; c.value = v; c.number_format = "0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str):
                c.font = BLUE
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${rr}")
        H.cell(rr, 7, tag); H.cell(rr, 8, src)
    # 衍生列（公式）
    for k, labl, f in (("oi", "營業利益（＝稅前 − 營業外）", "={c}{pti}-{c}{nonop}"), ("opx", "營業費用（＝毛利 − 營業利益）", "={c}{gp}-{c}{oi}")):
        R[k] = r; H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl); H.cell(r, 7, "Derived")
        H.cell(r, 8, "法說簡報營業利益為十億元四捨五入（FY2023 40.9、FY2024 47.7、FY2025 83.9）；本列由稅前與營業外推得，與之一致")
        for col, yy in (("D", "23"), ("E", "24"), ("F", "25")):
            H[f"{col}{r}"] = f.format(c=col, pti=R["pti"], nonop=R["nonop"], gp=R["gp"], oi=R.get("oi", r)); H[f"{col}{r}"].number_format = "#,##0.000"
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        r += 1
    R["chk"] = r
    H.cell(r, 1, "核對"); H.cell(r, 2, "稅前 − 所得稅 − 非控制權益 − 稅後淨利（應為 0）")
    for col in "DEF":
        H[f"{col}{r}"] = f"={col}{R['pti']}-{col}{R['tax']}-{col}{R['nci']}-{col}{R['ni']}"; H[f"{col}{r}"].number_format = "0.000"
    r += 1
    H.cell(r, 1, "核對"); H.cell(r, 2, "FY2025：本表營業利益與毛利 − Q 頁錨點（四捨五入差；Q 頁取法說簡報十億元值 83.9、190.2）")
    H[f"F{r}"] = f"=F{R['oi']}-a25_oi"; H[f"F{r}"].number_format = "0.000"
    H[f"G{r}"] = f"=F{R['gp']}-a25_gp"; H[f"G{r}"].number_format = "0.000"
    R["chk2"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 46
    for c in "CDEF":
        H.column_dimensions[c].width = 16
    H.column_dimensions["H"].width = 110
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"


def pl_standard(wb, CR):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    B = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), nop=g("業外收益"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"))
    x0 = lab(P, "延伸｜情境 B")
    gx = lambda label: lab(P, label, x0, exact=True, end=x0 + 14)
    X = dict(rev=gx("營收（延伸）"), ni=gx("淨利（延伸）"), sh=gx("股數（延伸）"), eps=gx("EPS（延伸）"), hdr=x0 + 1)
    D = wb["DCF"]
    dbr = lab(D, "情境 B")
    d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.5.6 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2031E（情境 B 基準；FY2026E 含 3Q26 實績，FY2030E–FY2031E 為延伸期）。"
                 "NT$B；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（台灣採用的 IFRS 合併財報，歸屬母公司淨利）。台達電不另發布 Non-GAAP，法說簡報與市場共識 EPS 皆以此口徑，故本表不做 Non-GAAP 調整。"
                 "FY2026E 以後的所得稅＝稅前 × 模型稅率（In 頁的稅率，與 DCF 同一個稅率），非控制權益＝稅前 − 所得稅 − 歸屬母公司淨利（模型以「淨利 ÷ 稅前」合併處理兩者）；歷史年度為法說簡報原值")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝1H26 實績＋3Q26 實績＋4Q26 假設；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益由淨利還原，同 DCF 頁）："
                 "毛利率與營業費用率沿用 FY2029E 據以算出營業成本、毛利與營業費用，與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + YEARS + ["說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    # 對帳中年度標示（FY2027E 起依 CY2027–CY2029 旗標；FY2030E–FY2031E 由 FY2029E 水位延伸，依 CY2029 旗標）
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
    eng = lambda key: [f"=PL!{c}{B[key]}" for c in "DEFG"]          # FY2026E–FY2029E（情境 B 區塊 D–G 欄）
    ext = lambda key: [f"=PL!H{X[key]}", f"=PL!I{X[key]}"]           # FY2030E、FY2031E（延伸表 H、I 欄）
    hs = lambda k: [f"=h_is_{k}_23", f"=h_is_{k}_24", f"=h_is_{k}_25"]
    row("rev", "營收", "NT$B", hs("rev") + eng("rev") + ext("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_is_rev_22-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022＝384.4（法說簡報四捨五入）")
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "CDEFGHI"] + [f"=J{L['rev']}*(1-$I{r_c + 2})", f"=K{L['rev']}*(1-$I{r_c + 2})"], M,
        "歷史與 FY2026E–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", hs("gp") + eng("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "NT$B", hs("opx") + eng("opx") + [f"=J{L['rev']}*$I{r_c + 3}/$I{L['rev']}", f"=K{L['rev']}*$I{r_c + 3}/$I{L['rev']}"], M,
        "FY2030E–FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "NT$B", hs("oi") + eng("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E–FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利還原；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額，營業外收益）", "NT$B", hs("nonop") + eng("nop") + ["=nop_f", "=nop_f"], M, "FY2030E–FY2031E 同 FY2027E 起的年度業外收益（In 頁 nop_f）")
    row("pti", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "NT$B", hs("tax") + [f"={c}{L['pti']}*tax_w" for c in "FGHIJK"], M, "FY2026E 起＝稅前 × 模型稅率（In 頁的稅率）")
    row("nci", "非控制權益損益", "NT$B", hs("nci") + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['pti'] + 3}" for c in "FGHIJK"], M, "FY2026E 起＝稅前 − 所得稅 − 稅後淨利（歸屬母公司）")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", hs("ni") + eng("ni") + ext("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數", "十億股", hs("sh") + eng("sh") + ext("sh"), "0.000")
    row("eps", "EPS（基本，公司公告）", "NT$", hs("eps") + eng("eps") + ext("eps"), "0.00", "歷史為公司公告基本 EPS（非稀釋）；預測為歸屬母公司淨利 ÷ 股數（無稀釋證券，同基本口徑）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/h_is_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 稅後淨利 32.666（法說簡報）")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（歷史與 FY2026E–FY2029E）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None, None], "0.000")
    row("c2b", "差額（FY2030E–FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利還原）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "差額（FY2030E–FY2031E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JK"], PCT)
    row("c3", "核對：稅後淨利 − 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "NT$B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{B['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    row("c4", "核對：稅前 − 所得稅 − 非控制權益 − 稅後淨利", "NT$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c5", "核對：延伸期稅前利益 × 淨利 ÷ 稅前比率（In 頁 nir_f）− 稅後淨利（FY2030E–FY2031E）", "NT$B", [None] * 7 + [f"={c}{L['pti']}*nir_f-{c}{L['ni']}" for c in "JK"], "0.000",
        "營業利益由 DCF 頁（淨利 ÷ nir_f − nop_f）還原，故稅前 × nir_f 應等於稅後淨利")
    row("c6", "核對：FY2025A 營業利益與毛利 − Q 頁錨點（四捨五入差，非 0）", "NT$B", [None, None, f"=E{L['oi']}-a25_oi", None, None, None, None, None, None], "0.000",
        "Q 頁取法說簡報十億元四捨五入值（營業利益 83.9、毛利 190.2）；本表取百萬元原值推得（83.932、190.157）")
    for col in COLS:
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    P.column_dimensions["L"].width = 70
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, B=B, X=X)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7   # DCF 頁：評價日後占比、折現年數（期中）
    assert str(D.cell(fr, 1).value).startswith("評價日後占該年度比例") and str(D.cell(ty, 1).value).startswith("折現年數")
    V = wb["Val"]; P = wb["PL"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    rb = lab(P, "情境 B："); r26 = lab(P, "營收", rb, exact=True)
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.5.6 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。台達電 FY2027 為模型推導年（非錨點），故營業利益率與營收倍數都從 FY2027 起動；"
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
                     T(" 元；要支持現價 "), 'TEXT(px,"#,##0")', T(" 元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2027 起營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv2"]},"#,##0")', T(" 十億元，模型為 "), f'TEXT(B{A["rv2m"]},"#,##0")',
                     T("）；或營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"#,##0")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 台達電過去十年相對同業的溢價 "), f'TEXT({coef},"0.00")', T(" 倍）。換言之，市場目前給台達電的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：估值並列四個目標價（G 節） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.5.6 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
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
    V.cell(r, 1, "=A6"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（日期見左欄）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（H 節） ----------
TOR_LABEL = {"WACC ±1pt": "折現率（WACC）低／高一個百分點",
             "β 低／高": "股權風險係數低／高（各自重新加槓桿）",
             "ERP 低／高": "股權風險溢酬 4%／6%",
             "RONIC＝WACC＋5pt": "永續期新投資報酬率提高到折現率加五個百分點（只有高）",
             "DCF 權重 0.3／0.7": "加權目標價中 DCF 的權重 30%／70%（基準 50%）"}


def sens_tor(wb, CR):
    SE = wb["Sens"]; V = wb["Val"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    assert SE.cell(hr, 1).value == "參數"
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.5.6 起，版面規格 v1.3 H 節；引用上方「目標價參數」表與 Val、Coord 頁，白話標籤，不回饋計算）").font = BOLD; t0 = r; r += 1
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
    SE.cell(r, 1, "同業比較法不加歷史溢價（相對倍數係數＝1；Comparable 取市場主流同業法）")
    SE.cell(r, 2, f"=(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/{tp}-1").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Val 頁：DCF 與未調整 Comparable 依 DCF 權重加權"); r += 1
    SE.cell(r, 1, "產業部署 GW 改用另一側讀數（Wattonomics 換算；Siliconomics 標對帳中，整本重建）")
    SE.cell(r, 2, f"=Coord!D{CR['tp']}").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Coord 頁兩讀數目標價（另一側相對現行）"); r += 1
    return dict(t0=t0, r1=first, r2=r - 1)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, TOR, CR):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）")
    cadj, cu, coef, pmed = vb("Comparable 目標價（相對倍數調整後）"), vb("Comparable 目標價（市場主流同業法）"), vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數")
    REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rb = lab(P, "情境 B：")
    pb = lambda label, c, ex=False: f"PL!{c}{lab(P, label, rb, exact=ex, end=rb + 70)}"
    ra = lab(P, "情境 A：")
    pa = lambda label, c, ex=False: f"PL!{c}{lab(P, label, ra, exact=ex, end=ra + 70)}"
    sm = lab(SE, "AI 部署 ×0.8"); ss = lab(SE, "份額流失（低）")
    cr = lambda col: f"Coord!{col}{CR['tp']}"
    dw = lab(SE, "加權目標價中 DCF 的權重", TOR["t0"])
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.6 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "), f'TEXT({dcf},"#,##0")',
        T(" 元偏低，同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")', T(" 元支撐；AI 電源與液冷續成長，但目標價對 AI 部署量與份額最敏感。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：產業部署量依 Siliconomics 基準、AI 電源每 kW 內容與份額相對值（皆為假設）成立，DCF 與同業比較法各占一半。"
                  "DCF 權重降到三成時為"), REC5(f"((0.3*{dcf}+0.7*{cadj})/px-1)"), T("，同業比較法不加歷史溢價時為"), REC5(f"((tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/px-1)"),
        T("，產業部署量改用另一側讀數時為"), f"Coord!C{CR['rec']}", T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，"), f'IF({up}<0,"低於","高於")', T("現價 "), f'TEXT(px,"#,##0")', T(" 元 "), f'TEXT(ABS({up}),"0.0%")',
        T("，評等"), rec, T("；DCF "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法 "), f'TEXT({cadj},"#,##0")', T(" 元各占一半。"))
    put("o2", T("AI 電源是成長主力但份額逐年下滑：情境 A AI 電源營收由 FY2026 約 "), f'TEXT({pa("AI 電源營收", "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pa("AI 電源營收", "F")}*10,"#,##0")',
        T(" 億元，同期有效份額由 "), f'TEXT({pa("台達電有效份額", "D", True)},"0.0%")', T(" 降到 "), f'TEXT({pa("台達電有效份額", "F", True)},"0.0%")', T("。"))
    put("o3", T("最大的變數是 AI 部署量與份額：AI 部署 ×0.8 使情境 A 的 FY2028 EPS "), f'TEXT(Sens!L{sm},"+0.0%;-0.0%")', T("，份額流失 "), f'TEXT(Sens!L{ss},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
        f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。"))
    put("o4", T("現價隱含：要支持現價，折現率須降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或營收全面放大 "), f'TEXT({rv("s1")},"0.00")', T(" 倍；目標價對 DCF 權重最敏感（30%／70% 各 "),
        f'TEXT(Sens!B{dw},"+0%;-0%")', T("／"), f'TEXT(Sens!C{dw},"+0%;-0%")', T("）。"))
    put("o5", T("對帳中：CY2027–CY2029 Siliconomics 標對帳中，另一側讀數（Wattonomics 換算）目標價 "), f'TEXT({cr("C")},"#,##0")', T(" 元（"), f"Coord!C{CR['rec']}",
        T("），較現行基準 "), f'TEXT(Coord!D{CR["tp"]},"+0.0%;-0.0%")', T("；本模型的部署 GW 與 AI 電源市場同 Siliconomics 基準。"))
    put("drv", T("分部營收推導（情境 B）：AI 電源營收＝部署 GW × 每 kW 電源內容 × 台達電有效份額；液冷營收＝部署 GW × 每 GW 液冷營收；電源及零組件＝AI 電源＋非 AI；基礎設施＝液冷＋非液冷；"
                 "自動化與交通各依年增率；四大事業群相加為合併營收。FY2028 AI 電源 "), f'TEXT({pb("AI 電源營收", "F")},"#,##0")', T("、液冷 "), f'TEXT({pb("液冷營收", "F")},"#,##0")',
        T("、非 AI 與其他 "), f'TEXT({pb("非 AI 營收", "F")},"#,##0")', T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0")', T(" 十億元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：Andy 16:57 退件後全面改為個股 HTML 版面規格 v1.3（樣板 NVIDIA）並補三層協調 C0：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；"
                  "新增 RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023A–FY2031E，GAAP）、Hist 頁（損益表歷史）、Val 0.7（四個目標價並列）、Sens 末「加權目標價相對基準」、"
                  "Coord 頁「三層協調 C0」（讀哪一層哪一版、由下往上對帳）、README 頁末 v1.3 白話列；README A1 標題版號更正；既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    return out


# ---------- Coord：三層協調 C0 ----------
def coord_c0(wb, SIL, CR, PLA):
    C = wb["Coord"]; P = wb["PL"]; IN = wb["In"]; V = wb["Val"]
    rb = lab(P, "情境 B：")
    pb = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    r = C.max_row + 3
    C.cell(r, 1, "三層協調 C0（三層基準協調規則 v1.4 C0；Andy 2026-11-14 16:57／16:53 指示）：讀哪一層、哪一版，與由下往上對帳").font = BOLD; r += 1
    C.cell(r, 1, "原則：同一數字只有一個權威層，下游註明版本、不自行重估（Project 規則第 3 條）；差距超過 ±15% 標「對帳中」，不改上游值，由主控開對帳單；個股不裁決（C4），往上回饋只有證據（C5）。").font = Font(italic=True)
    r += 2
    # 1. 讀哪一層、哪一版
    C.cell(r, 1, "1. 本模型讀哪一層、哪一版（硬體個股讀 Siliconomics；Wattonomics 只經 Siliconomics 的 R4 比值與另一側讀數進入，現行版本 v0.3.2）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級名稱（版本）", "本模型用途"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    inrow = lambda k: next(i for i in range(4, IN.max_row + 1) if IN.cell(i, 1).value == k)
    V_ = lambda k: f"=In!E{inrow(k)}"
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    txt = lambda *ks: "=" + '&"／"&'.join(f'TEXT(In!E{inrow(k)},"0.00")' for k in ks)
    rd = []

    def kv(label, val, who, use, fmt=None):
        nonlocal r
        C.cell(r, 1, label); c = C.cell(r, 2, val)
        if fmt:
            c.number_format = fmt
        C.cell(r, 3, who); C.cell(r, 4, use); rd.append(r); r += 1
    kv("Siliconomics 快照版本", SIL.VERSION, ver, "行業變數單一來源（siliconomics/reader.py）；下游不讀 In、Path、Mkt 等內部頁")
    kv("每 kW AI 電源內容 CY2026／27／28／29（US$/kW；pwr_pkw_26–29）", f'=TEXT(pkw_26,"0")&"／"&TEXT(pkw_27,"0")&"／"&TEXT(pkw_28,"0")&"／"&TEXT(pkw_29,"0")', ver, "AI 電源市場＝部署 GW × 每 kW 內容（In 頁 B 節）")
    kv("台達電份額相對值 CY2027／28／29（÷ CY2026；pwr_drel_27–29）", '=TEXT(rel_27,"0.00")&"／"&TEXT(rel_28,"0.00")&"／"&TEXT(rel_29,"0.00")', ver, "有效份額路徑；FY2026 有效份額由公司 AI 相關占比校準（Derived）")
    kv("匯率（NT$/US$；fx_twd）", "=fx", ver, "AI 電源市場換算新台幣", "0.00")
    kv("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "基準情境 B 的 AI IT 資本支出係數；A（原始顆數）、C（0.6）為個股自訂情境")
    kv("AI IT 資本支出 CY2026–29（US$B；markets/network）", f'=TEXT(itc_26,"#,##0")&"／"&TEXT(itc_27,"#,##0")&"／"&TEXT(itc_28,"#,##0")&"／"&TEXT(itc_29,"#,##0")',
       f"=In!H{inrow('itc_26')}", "部署 GW 的分子（原始顆數口徑）；本層成品為市場層，Tokenomics 經橋接表進入")
    kv("每 GW IT 資本支出（US$B；GB300／VR200；markets/chip_bridge）", '=TEXT(itpg_gb,"0.0")&"／"&TEXT(itpg_vr,"0.0")', f"=In!H{inrow('itpg_vr')}", "部署 GW 的分母（Tokenomics IF_CapexIT，經橋接表唯讀）")
    kv("Wattonomics 錨定係數 CY2027／28／29（SIF_WTAnchor_Base；R4 比值）", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")',
       f"{ver}；Wattonomics 現行 v0.3.2（只經 Siliconomics 進入）", "只進敏感度（Sens「Wattonomics 錨定」案例）；不改基準")
    kv("另一側讀數 ÷ 現行基準 CY2026／27／28／29（SIF_DeployGW_AltRatio_<年>）", f'=TEXT(B{CR["ratio"]},"0.000")&"／"&TEXT(C{CR["ratio"]},"0.000")&"／"&TEXT(D{CR["ratio"]},"0.000")&"／"&TEXT(E{CR["ratio"]},"0.000")',
       f"{ver}（比值列見本頁上方 B 節）", "兩讀數目標價（上方 B 節）只換產業部署 GW 一個輸入")
    r += 1
    # 2. C0 由下往上對帳
    C.cell(r, 1, "2. C0 由下往上對帳：本模型自己的產業量與層級值（逐年；差距＝本模型 ÷ 層級值 − 1；本公司自己的量超過 ±15% 標「對帳中」，不改上游值）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位", "口徑、來源、證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    C_h = r; r += 1
    N = SIL.NAMES
    G = {}

    def block(key, title, unit, model, layer, note_m, note_l, fmt, gapfmt="+0.0%;-0.0%", lay_is_formula=False, judge_extra="", tag="對帳中"):
        nonlocal r
        C.cell(r, 1, f"{title}：本模型（情境 B；PL 頁）"); C.cell(r, 6, unit); C.cell(r, 7, note_m)
        for j, f in enumerate(model):
            C.cell(r, 2 + j, f).number_format = fmt
        rm = r; r += 1
        C.cell(r, 1, f"{title}：層級值"); C.cell(r, 6, unit); C.cell(r, 7, note_l)
        for j, v in enumerate(layer):
            c = C.cell(r, 2 + j, v); c.number_format = fmt
            if not (isinstance(v, str) and v.startswith("=")):
                c.font = BLUE
        rl = r; r += 1
        C.cell(r, 1, f"{title}：差距（本模型 ÷ 層級值 − 1）")
        for j, col in enumerate("BCDE"):
            C.cell(r, 2 + j, f"={col}{rm}/{col}{rl}-1").number_format = gapfmt
        rg = r; r += 1
        C.cell(r, 1, f"{title}：判定（±15%）")
        for j, col in enumerate("BCDE"):
            C.cell(r, 2 + j, f'=IF(ABS({col}{rg})>0.15,"{tag}{judge_extra}","一致")')
        rj = r; r += 1
        G[key] = (rm, rl, rg, rj)
    pc = lambda label: [f"=PL!{c}{pb(label)}" for c in "DEFG"]
    block("gw", "部署 GW（IT 電力）", "GW", pc("部署 GW（IT 電力）"), list(N["SIF_DeployGW_Base"]),
          "Derived：AI IT 資本支出（markets/network）× 情境 B 顆數係數 ÷ 每 GW IT 資本支出（bridge）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00")
    block("mkt", "AI 電源市場（US$B）", "US$B", pc("AI 電源市場（美元）"), list(N["SIF_PwrMkt_Base"]),
          "Derived：部署 GW × 每 kW 電源內容", f"Derived：Siliconomics {SIL.VERSION} SIF_PwrMkt_Base", "0.00")
    dsh, drel = SIL.V["pwr_dsh_26"]["base"], [1.0] + [SIL.V[f"pwr_drel_{y}"]["base"] for y in (27, 28, 29)]
    block("sh", "台達電有效份額", "%", pc("台達電有效份額"), [f"={dsh}*{x}" for x in drel],
          "Derived：FY2026 以公司 AI 相關占比 × 營收 ÷ 市場校準，之後 × 份額相對值（Assumed）", f"Interested-party／Assumed：Siliconomics {SIL.VERSION} pwr_dsh_26 × pwr_drel_27–29（FY2026 快照值 {dsh:.4f}）", "0.0%")
    block("pkw", "每 kW 電源內容", "US$/kW", pc("AI 電源內容"), [SIL.V[f"pwr_pkw_{y}"]["base"] for y in (26, 27, 28, 29)],
          "讀自 Siliconomics（In 頁 pkw_26–29）", f"Interested-party（2026–27：凱基推導）／Assumed（2028 起）：Siliconomics {SIL.VERSION} pwr_pkw_26–29", "#,##0")
    # 殘差：已建模個股份額加總（台達電本模型＋光寶科層級值）vs SIF_PwrShareSum_Base
    lsh = SIL.V["pwr_lsh_26"]["base"]
    rm = r; C.cell(r, 1, "已建模個股份額加總：本模型（台達電本模型份額＋光寶科層級值）"); C.cell(r, 6, "%")
    C.cell(r, 7, "Derived：台達電取本模型；光寶科取 Siliconomics 層級值（光寶科模型不在本模型讀取範圍）")
    lrel = [1.0] + [SIL.V[f"pwr_lrel_{y}"]["base"] for y in (27, 28, 29)]
    for j, col in enumerate("BCDE"):
        C.cell(r, 2 + j, f"={col}{G['sh'][0]}+{lsh}*{lrel[j]}").number_format = "0.0%"
    r += 1
    C.cell(r, 1, "已建模個股份額加總：層級值（Siliconomics）"); C.cell(r, 6, "%"); C.cell(r, 7, f"Derived：Siliconomics {SIL.VERSION}")
    for j, v in enumerate(N["SIF_PwrShareSum_Base"]):
        c = C.cell(r, 2 + j, v); c.number_format = "0.0%"; c.font = BLUE
    rl = r; G["shsum"] = (rm, rl); r += 1
    C.cell(r, 1, "已建模個股份額加總：差距（百分點）")
    for j, col in enumerate("BCDE"):
        C.cell(r, 2 + j, f"={col}{rm}-{col}{rl}").number_format = "+0.0%;-0.0%"
    r += 1
    C.cell(r, 1, "C0 恆等式：殘差（其他廠商＝1 − 已建模個股加總；須明列且不為空）"); C.cell(r, 6, "%"); C.cell(r, 7, "Flex、麥格米特、Vertiv 等；不得歸零（C0 第 3 點）")
    for j, col in enumerate("BCDE"):
        C.cell(r, 2 + j, f"=1-{col}{rm}").number_format = "0.0%"
    G["resid"] = r; r += 1
    # 層級之間與外部估計：參考，不計入本公司 C0（主控 2026-11-14：Siliconomics 對帳單 r3 處理）
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：Wattonomics 側讀數與外部估計").font = BOLD; r += 1
    alt = [N[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)]
    block("alt", "部署 GW 對 Wattonomics 側讀數", "GW", pc("部署 GW（IT 電力）"), alt,
          "同上（本模型情境 B）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Alt_<年>（Wattonomics 換算值；需求側讀數）", "0.00", judge_extra="（Siliconomics 與 Wattonomics 兩層之間的差距，已並列兩讀數）", tag="參考：層級之間")
    ubs = [SIL.V["pwr_ubs_26"]["base"], SIL.V["pwr_ubs_27"]["base"]]
    rm = r; C.cell(r, 1, "AI 電源市場對 UBS 估計：層級值（US$B）"); C.cell(r, 6, "US$B"); C.cell(r, 7, "Derived：同上（SIF_PwrMkt_Base）")
    for j, col in enumerate("BC"):
        C.cell(r, 2 + j, f"={col}{G['mkt'][1]}").number_format = "0.0"
    r += 1
    C.cell(r, 1, "AI 電源市場對 UBS 估計：UBS（US$B）"); C.cell(r, 6, "US$B"); C.cell(r, 7, "Interested-party：UBS（經 TechNews 2026-10-12；口徑未知，是否含 BBU、HVDC、非 NVIDIA；Siliconomics pwr_ubs_26／27）")
    for j, v in enumerate(ubs):
        c = C.cell(r, 2 + j, v); c.number_format = "0.0"; c.font = BLUE
    rl = r; r += 1
    C.cell(r, 1, "AI 電源市場對 UBS 估計：差距（層級值 ÷ UBS − 1）")
    for j, col in enumerate("BC"):
        C.cell(r, 2 + j, f"={col}{rm}/{col}{rl}-1").number_format = "+0.0%;-0.0%"
    rg = r; r += 1
    C.cell(r, 1, "AI 電源市場對 UBS 估計：判定（±15%；僅參考）")
    for j, col in enumerate("BC"):
        C.cell(r, 2 + j, f'=IF(ABS({col}{rg})>0.15,"參考（外部估計，口徑未知，未改值）","一致")')
    G["ubs"] = (rm, rl, rg, r); r += 1
    # 層級值快照輸入（藍字常數只在此處；份額層級值列與份額加總列連到這裡）
    C.cell(r, 1, "層級值快照輸入：台達電有效份額（Siliconomics 快照）"); C.cell(r, 6, "%"); C.cell(r, 7, f"Interested-party／Assumed：Siliconomics {SIL.VERSION}（pwr_dsh_26 × pwr_drel_26–29 快照常數；上方份額層級值列連到此列）")
    for j, x in enumerate(drel):
        c = C.cell(r, 2 + j, f"={dsh}*{x}"); c.number_format = "0.0%"; c.font = BLUE
    r_sh_in = r; r += 1
    C.cell(r, 1, "層級值快照輸入：光寶科有效份額（Siliconomics 快照）"); C.cell(r, 6, "%"); C.cell(r, 7, f"Interested-party／Assumed：Siliconomics {SIL.VERSION}（pwr_lsh_26 × pwr_lrel 快照常數；份額加總本模型列連到此列）")
    for j in range(4):
        c = C.cell(r, 2 + j, f"={lsh}*{lrel[j]}"); c.number_format = "0.0%"; c.font = BLUE
    r_lt_in = r; r += 1
    for j, col in enumerate("BCDE"):
        C.cell(G["sh"][1], 2 + j, f"={col}{r_sh_in}").font = Font(color="000000")
        C.cell(G["shsum"][0], 2 + j, f"={col}{G['sh'][0]}+{col}{r_lt_in}")
    r += 1
    # 3. 結果摘要
    C.cell(r, 1, "3. 對帳結果（HTML 用；公式帶數字）").font = BOLD; r += 1
    mx = lambda k: f"MAX(ABS(B{G[k][2]}),ABS(C{G[k][2]}),ABS(D{G[k][2]}),ABS(E{G[k][2]}))"
    C.cell(r, 1, F(T(f"本模型讀 Siliconomics {SIL.VERSION}：部署 GW、AI 電源市場、有效份額與每 kW 內容四項層級量與本模型的最大差距分別為 "),
                   f'TEXT({mx("gw")},"0.0%")', T("、"), f'TEXT({mx("mkt")},"0.0%")', T("、"), f'TEXT({mx("sh")},"0.0%")', T("、"), f'TEXT({mx("pkw")},"0.0%")',
                   T("（±15% 內，一致）；已建模個股加總之外的殘差為 "), f'TEXT(MIN(B{G["resid"]}:E{G["resid"]}),"0%")', T("–"), f'TEXT(MAX(B{G["resid"]}:E{G["resid"]}),"0%")',
                   T("，不為空。")))
    G["sum1"] = r; r += 1
    C.cell(r, 1, F(T("層級之間（參考）：本模型（Siliconomics 基準）的部署 GW 在 CY2027–CY2029 比 Wattonomics 換算值高 "), f'TEXT(C{G["alt"][2]},"0%")', T("／"), f'TEXT(D{G["alt"][2]},"0%")', T("／"),
                   f'TEXT(E{G["alt"][2]},"0%")', T("，超過 ±15%；這是 Siliconomics 與 Wattonomics 兩層之間的差距（Siliconomics 對帳單 r3 處理；已並列兩讀數目標價），"
                                                  "不是本模型自己的差距，不計入本公司 C0，本模型不改上游值、不裁決。")))
    G["sum2"] = r; r += 1
    C.cell(r, 1, F(T("外部估計（參考）：Siliconomics 的 AI 電源市場對 UBS 估計低 "), f'TEXT(-B{G["ubs"][2]},"0%")', T("（CY2026）、"), f'TEXT(-C{G["ubs"][2]},"0%")',
                   T("（CY2027）；UBS 為 Interested-party、口徑未知，只列參考，不計入本公司 C0，本模型不改值。")))
    G["sum3"] = r; r += 1
    C.cell(r, 1, "C4／C5：本模型不裁決兩層的對帳中年份；本版沒有可往上回饋的新證據（往上回饋只有實績、法定文件揭露的承諾、已簽約或已確認站點，個股預測與公司指引不往上流）。"); G["sum4"] = r; r += 1
    C.column_dimensions["A"].width = 70; C.column_dimensions["D"].width = 24; C.column_dimensions["F"].width = 10; C.column_dimensions["G"].width = 110
    G["hdr"] = C_h; G["rd0"], G["rd1"] = rd[0], rd[-1]
    return G


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    head, _, rest = a1.partition("；")
    assert "個股模型 v0.3（" in head, head
    head = head.replace("個股模型 v0.3（", f"個股模型 v{VER}（", 1)
    head = head.replace("2026-10-28", BUILT, 1)
    RD["A1"].value = (head + "；v0.3 初版（2026-10-28）；" + rest
                      + f"；v{VER}（{BUILT}）：Andy 16:57 退件後改為個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）並補 Coord 頁三層協調 C0；A1 標題版號由舊的 v0.3 更正；既有數值不變")


def apply(wb, here, SIL, CR):
    HR = hist_is(wb)
    PLA = pl_standard(wb, CR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_tor(wb, CR)
    G = coord_c0(wb, SIL, CR, PLA)
    RMA = readme13(wb, RVA, PLA, ST, CR)
    version(wb)
    wb._v056 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, G=G, RM=RMA)
    return f"v056_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {ST['t0']} 列；Coord C0；README v1.3 白話列 {len(RMA)} 列"
