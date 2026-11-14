# 廣達 v0.4.8（2026-11-14；v0.4.7 起的後處理，v0.4.8 在 Coord 頁末新增 C0〔實績〕節、同源列改標「讀取層級」，見 coord_c0）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7／v0.3.8）＋三層協調新機制 C0（三層基準協調規則 v1.4.1）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.4.6 逐格比對）：
#   Hist 頁（新）：損益表歷史 FY2023–FY2025（S&P Global Market Intelligence，經 stockanalysis.com；FY2022 營收與淨利只供年增率）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B；Non-GAAP 不適用，廣達只有 IFRS 合併財報，標 GAAP 口徑）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末：0.7 四個目標價並列（HTML 用）
#   Sens 頁末：目標價敏感度：業務主因（整本重算快照 quanta/sens.json）與「加權目標價相對基準」白話表（HTML 敏感度主圖用）
#   README 頁末：v1.3 展示層白話列（橫幅說明、重點觀察 ≤5 點、分部營收推導說明、本版要點）；A1 標題段版號更正為 v0.4.7
#   Coord 頁末：C0 由下往上對帳（讀哪一層、哪一版；本模型組裝市場與 Siliconomics SIF_AsmMkt_Base 等逐年比較；差距 ±15% 標對帳中）
# build_quanta.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, SIL, here, coord_rows)。
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.8", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SP = "S&P Global Market Intelligence 經 stockanalysis.com 損益表頁（https://stockanalysis.com/quote/tpe/2382/financials/income-statement/；2026-11-14 讀取，頁面更新 2026-08-13；單位百萬元新台幣，此處換算為十億元）"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


def lab(ws, prefix, start=1, col=1, exact=False):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- Hist：損益表歷史 ----------
# (鍵, 項目, FY2023, FY2024, FY2025, FY2022（只供年增率）)；金額十億新台幣、EPS 新台幣、股數十億股
HIS = [
    ("rev", "營收", 1085.611, 1410.756, 2123.689, 1280.429),
    ("cogs", "營業成本", 1000.728, 1299.996, 1975.357, None),
    ("gp", "毛利", 84.883, 110.760, 148.332, None),
    ("opx", "營業費用（銷管＋研發）", 41.333, 49.138, 60.936, None),
    ("oi", "營業利益", 43.550, 61.622, 87.396, None),
    ("pti", "稅前利益", 51.731, 73.167, 95.273, None),
    ("tax", "所得稅", 11.243, 12.884, 19.487, None),
    ("nci", "少數股權損益", 0.81219, 0.58113, 0.79861, None),
    ("ni_rep", "歸屬母公司淨利（資料來源列示值）", 39.676, 59.702, 74.988, 28.957),
    ("eps_b", "EPS（基本，新台幣；以流通股數 3.855 十億股計）", 10.29, 15.49, 19.45, None),
    ("eps_d", "EPS（稀釋，新台幣；資料來源列示值）", 10.23, 15.32, 18.91, None),
    ("sh_b", "流通股數（十億股）", 3.855, 3.855, 3.855, None),
    ("sh_d", "稀釋加權股數（十億股）", 3.879, 3.904, 3.999, None),
]


def hist_is(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 廣達損益表歷史（v0.4.7 新增；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣。"
               "FY(t)＝CY(t)。來源為第三方彙整（S&P Global Market Intelligence），未逐項對公司年報原文；與模型 Q 頁 FY2025（凱基、公司揭露比例推得）有約 0.1 十億元差異，見表末對照列")
    H["A1"].font = BOLD
    r = 3
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2022（只供年增率）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v23, v24, v25, v22 in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25), ("F", "22", v22)):
            if v is None: continue
            c = H[f"{col}{r}"]; c.value = v; c.font = BLUE
            c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"
            dn(wb, f"h_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, "Verified／第三方彙整"); H.cell(r, 8, SP + ("；FY2022 取同站統計頁（營收 1,280,429、淨利 28,957 百萬元）" if k in ("rev", "ni_rep") else ""))
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    chk = {}

    def crow(key, text, f, fmt="0.000"):
        nonlocal r
        H.cell(r, 1, text)
        for col in "CDE":
            H[f"{col}{r}"] = f.format(c=col, **{k: f"{col}{v}" for k, v in rows.items()}); H[f"{col}{r}"].number_format = fmt
        chk[key] = r; r += 1
    crow("c1", "營收 − 營業成本 − 毛利", "={rev}-{cogs}-{gp}")
    crow("c2", "毛利 − 營業費用 − 營業利益", "={gp}-{opx}-{oi}")
    crow("c3", "稅前利益 − 所得稅 − 少數股權損益 − 歸屬母公司淨利（列示值；四捨五入差 ≤ 0.001）", "={pti}-{tax}-{nci}-{ni_rep}")
    crow("c4", "歸屬母公司淨利 ÷ 流通股數 − 基本 EPS", "={ni_rep}/{sh_b}-{eps_b}", "0.000")
    r += 1
    H.cell(r, 1, "對照：歷史（本頁，S&P Global）− 模型 FY2025A（Q 頁；凱基、公司揭露比例推得）。非零差異＝兩個來源口徑或四捨五入不同；模型既有 FY2025 欄不改").font = BOLD; r += 1
    for key, text, h_k, a_nm in (("d_rev", "營收", "rev", "a25_rev"), ("d_gp", "毛利", "gp", "a25_gp"), ("d_oi", "營業利益", "oi", "a25_oi"), ("d_ni", "歸屬母公司淨利", "ni_rep", "a25_ni"), ("d_eps", "EPS（基本）", "eps_b", "a25_eps")):
        H.cell(r, 1, f"差異：{text}（本頁 − Q 頁 FY2025）")
        H[f"E{r}"] = f"=E{rows[h_k]}-{a_nm}"; H[f"E{r}"].number_format = "0.000"; chk[key] = r; r += 1
    H.column_dimensions["A"].width = 52; H.column_dimensions["B"].width = 44
    for c in "CDEF": H.column_dimensions[c].width = 14
    H.column_dimensions["H"].width = 100
    return rows


# ---------- PL：標準損益表 ----------
YR = list(range(2023, 2032))
COLS = "CDEFGHIJK"


def dep_years(SIL):
    ys = [y for y in (2026, 2027, 2028, 2029) if SIL.RECO_FLAG.get(y) == 1]
    out = set(ys)
    if 2029 in ys:
        out |= {2030, 2031}   # 延伸期由 FY2029 路徑收斂而來
    return out


def pl_standard(wb, SIL):
    P = wb["PL"]; D = wb["DCF"]
    rb = lab(P, "情境 B：")
    g = lambda p: lab(P, p, rb, exact=True)
    R = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"))
    xb = lab(P, "延伸｜情境 B")
    X = {k: lab(P, p, xb, exact=True) for k, p in (("rev", "營收（延伸）"), ("ni", "淨利（延伸）"), ("sh", "股數（延伸）"), ("eps", "EPS（延伸）"))}
    db = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", db)
    dep = dep_years(SIL)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.4.7 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績，Hist 頁）＋FY2026E–FY2031E（情境 B 基準；FY2026 為錨點＝1H26 實績＋3Q 營收實績＋假設）。NT$B；EPS 新台幣。"
                 "只引用本頁上方、Hist、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, ("口徑：GAAP（IFRS 合併財報，歸屬母公司）。廣達不另編 Non-GAAP，市場共識與公司法說皆用合併財報口徑，故本表用 GAAP；模型淨利為歸屬母公司淨利。"
                  "歷史 EPS 為基本 EPS（以流通股數 3.855 十億股計，與模型股數一致）；稀釋 EPS 在表末對照列。"
                  "預測期模型只推營業利益與「淨利 ÷ 營業利益」一個係數（涵蓋業外、稅與少數股權），本表把它拆成：其他收入＝營業利益 × FY2025 實績比例、少數股權損益＝稅前利益 × FY2025 實績比例、"
                  "所得稅＝稅前利益 − 少數股權損益 − 淨利（由模型淨利倒推，等於隱含有效稅率）。這三列是展示用拆分，不回饋計算、不改淨利與目標價。"))
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益取 DCF 頁情境 B＝淨利 ÷ 淨利係數）。"
                 "延伸期毛利率與營業費用率沿用前一年，據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + [f"FY{y}{'A' if y <= 2025 else 'E'}" + ("（對帳中）" if y in dep else "") for y in YR] + ["說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r_hdr = r; r += 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1
    M = "#,##0.0"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFG"]            # FY2026E–FY2029E（PL 情境 B 區塊 D–G 欄）
    xe = lambda key: [f"=PL!H{X[key]}", f"=PL!I{X[key]}"]            # FY2030E–FY2031E（延伸表 H、I 欄）
    h3 = lambda k: [f"=h_{k}_23", f"=h_{k}_24", f"=h_{k}_25"]
    row("rev", "營收", "NT$B", h3("rev") + fc("rev") + xe("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", h3("cogs") + [f"=F{L['rev']}-F{r_c + 1}", f"=G{L['rev']}-G{r_c + 1}", f"=H{L['rev']}-H{r_c + 1}", f"=I{L['rev']}-I{r_c + 1}",
                                                          f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-J{r_c + 2})"], M,
        "FY2030E、FY2031E＝營收 ×（1 − 前一年毛利率）")
    row("gp", "毛利", "NT$B", h3("gp") + fc("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝前一年（沿用）")
    r_o = r
    row("opx", "營業費用（銷管＋研發）", "NT$B", h3("opx") + fc("opx") + [f"=J{L['rev']}*I{r_o}/I{L['rev']}", f"=K{L['rev']}*J{r_o}/J{L['rev']}"], M,
        "FY2030E、FY2031E＝營收 × 前一年營業費用率")
    row("oi", "營業利益", "NT$B", h3("oi") + fc("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    ro = "(h_pti_25-h_oi_25)/h_oi_25"
    row("oth", "其他收入（淨額：利息、匯兌、投資收益等）", "NT$B", [f"=h_pti_{y}-h_oi_{y}" for y in ("23", "24", "25")] + [f"={c}{L['oi']}*{ro}" for c in "FGHIJK"], M,
        "歷史＝稅前利益 − 營業利益；預測＝營業利益 × FY2025 實績比例（展示用拆分）")
    row("pti", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    r_t = r
    row("tax", "所得稅", "NT$B", h3("tax") + [f"={c}{L['pti']}-{c}{r_t + 1}-{c}{r_t + 2}" for c in "FGHIJK"], M, "預測＝稅前利益 − 少數股權損益 − 淨利（由模型淨利倒推）")
    row("nci", "少數股權損益", "NT$B", h3("nci") + [f"={c}{L['pti']}*h_nci_25/h_pti_25" for c in "FGHIJK"], M, "預測＝稅前利益 × FY2025 實績比例（展示用拆分）")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", h3("ni_rep") + fc("ni") + xe("ni"), M, "歷史＝資料來源列示值（Hist 頁）")
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數（十億股）", "十億股", h3("sh_b") + fc("sh") + xe("sh"), "0.000", "歷史＝流通股數；預測＝模型股數（不含稀釋）")
    row("eps", "EPS（基本；新台幣）", "NT$", h3("eps_b") + fc("eps") + xe("eps"), "0.00", "歷史＝公司揭露基本 EPS；預測＝淨利 ÷ 模型股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/h_ni_rep_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    row("etr", "有效稅率（所得稅 ÷ 稅前利益；預測為隱含值）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT)
    P.cell(r, 1, "稀釋 EPS 對照（歷史；預測不估稀釋）").font = BOLD; r_dil = r; r += 1
    row("g_epsd", "EPS（稀釋；資料來源列示值）", "NT$", h3("eps_d"), "0.00", "稀釋加權股數見 Hist 頁")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用前一年比率）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額（FY2030E、FY2031E）占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 少數股權損益 − 稅後淨利（歷史有四捨五入差，≤ 0.001）", "NT$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "NT$B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    for col in "CDEFGHIJK": P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, dil=r_dil, chk=r_chk, start=s0, R=R, X=X)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = lab(D, "評價日後占該年度比例"), lab(D, "折現年數（期中")
    tvy = lab(D, "終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    nc_cell = f"Val!$B${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.7 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「約一年後的前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。"
                "廣達 FY2027 營運資金大增（營收年增 36%，無槓桿自由現金流為負），所以 FY2027 視為錨點、不隨反解變動")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "說明"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note: RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（新台幣）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（NT$B）", f"={nc_cell}", "#,##0.0", "Val 頁 0.1 節：淨現金與投資＋stub 期 FCFF − 已宣告股利")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（NT$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；NT$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（新台幣）", f"=DCF!$C${ps}", "0.00")
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
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvy}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（NT$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（NT$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（NT$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（NT$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（NT$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}<=0,"現價低於營業利益率為零時的價值：現價甚至沒有給預測期獲利定價",IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（NT$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（NT$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（NT$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；NT$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（新台幣）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tvy}+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（新台幣，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "約一年後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 元；要支持現價 "), 'TEXT(px,"0.00")', T(" 元，企業價值須"), f'IF(B{A["evx"]}<1,"縮小為","放大為")', T("模型的 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}>B{A["w0"]},"升到 ","降到 ")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面"), f'IF(B{A["s1"]}<1,"縮小為 ","放大為 ")',
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或營業利益率"),
                     f'IF(B{A["m1"]}<B{A["m0"]},"降到 ","提高到 ")', f'TEXT(B{A["m1"]},"0.0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0.0%")', T("，終值年權重大）。"),
                     f'IF(B{A["m1"]}<=0,"現價低於預測期獲利為零時的價值。","")'))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 元相當於約一年後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 廣達過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給廣達的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，歷史相對倍數比為 "), f'TEXT({coef},"0.00")', T("。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：0.7 四個目標價並列（HTML 用） ----------
def val_g(wb):
    V = wb["Val"]
    r = V.max_row + 2
    V.cell(r, 1, "0.7 四個目標價並列（v0.4.7 新增；個股 HTML 版面規格 v1.3 G 節；只引用本頁既有列，不回饋計算）").font = BOLD; g0 = r; r += 1
    for j, h in enumerate(["項目", "目標價（NT$）", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    cons = lab(V, "市場共識目標價", exact=True)
    items = [("DCF 目標價（情境加權；進加權）", "=tp_dcf", "單純 DCF：三情境無槓桿自由現金流折現，價值推到 12 個月後，依情境權重加權"),
             ("Comparable 目標價（相對倍數調整後；進加權）", "=tp_comp_adj", "同業前瞻 P/E 中位數 × 廣達十年相對同業倍數比 × 約一年後前瞻 EPS（情境加權）"),
             ("Comparable 目標價（市場主流同業法，未調整；不進加權）", "=tp_comp", "同業前瞻 P/E 中位數 × 約一年後前瞻 EPS，未做相對倍數調整"),
             ("加權目標價（橫幅主值）", "=tp_final", "DCF 與 Comparable（調整後）依 DCF 權重加權；投資建議依它的上檔空間判定"),
             ("市場共識目標價", "=cons_tp", f"=Val!E{cons}")]
    for k, (a, f, d) in enumerate(items):
        V.cell(r, 1, a); V.cell(r, 2, f).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, d); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return g0


# ---------- Sens：業務主因快照與加權目標價相對基準 ----------
BIZ = [  # (白話標籤, 低案設定, 高案設定, 鍵) ；鍵＝QUANTA_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄）
    ("廣達份額（相對值 0.85／0.75／0.70 ～ 1.1）", "rel_27=lo;rel_28=lo;rel_29=lo", "rel_27=hi;rel_28=hi;rel_29=hi", "份額"),
    ("客供比例（全為零 ～ 25%／40%／50%）", "cs_27=lo;cs_28=lo;cs_29=lo", "cs_27=hi;cs_28=hi;cs_29=hi", "客供比例"),
    ("客供部分保留毛利比例（70% ～ 100%）", "keep=lo", "keep=hi", "客供保留毛利"),
    ("每 GW 整櫃價值 2028–29（×0.85 ～ ×1.2）", "vpg_28=lo;vpg_29=lo", "vpg_28=hi;vpg_29=hi", "每 GW 價值"),
    ("AI 部署量（FY2027 起 AI IT 資本支出 ×0.8 ～ ×1.2）", "itc_27=x0.8;itc_28=x0.8;itc_29=x0.8", "itc_27=x1.2;itc_28=x1.2;itc_29=x1.2", "AI 部署量"),
    ("非 AI 營收成長（−5% ～ +8%／+6%）", "g_nai27=lo;g_nai28=lo;g_nai29=lo", "g_nai27=hi;g_nai28=hi;g_nai29=hi", "非 AI 成長"),
    ("AI 毛利率差距（低 3.5 個百分點 ～ 0.5 個百分點）", "d_ai=lo", "d_ai=hi", "AI 毛利率差距"),
    ("VR 世代毛利率稀釋（−0.6 個百分點 ～ 不稀釋）", "vrd=lo", "vrd=hi", "VR 稀釋"),
    ("營業費用彈性（營業費用年增 ÷ 營收年增：0.2 ～ 0.7；基準 0.4）", "phi=lo", "phi=hi", "營業費用彈性"),
    ("淨利 ÷ 營業利益（0.75 ～ 0.89）", "nio_f=lo", "nio_f=hi", "淨利係數"),
]
TOR_LABEL = {"WACC ±1pt": "折現率 WACC 低一個百分點／高一個百分點", "β 低／高": "β（風險係數）取低／高", "ERP 低／高": "股權風險溢酬 4%／6%",
             "永續成長率": "永續成長率低一個百分點／上限", "RONIC": "永續期新投資報酬率提高（只有高）", "資本支出 ÷ 營收": "資本支出占營收 −1／＋1 個百分點",
             "營運資金 ÷ 營收": "營運資金占營收 10%／20%", "淨現金不含權益投資": "淨現金不含權益投資（只有低）", "終值改用 FCFF": "終值改用 FCFF ×（1＋g）替代法（只有高）",
             "美元口徑折現": "改用美元口徑折現（只有低）", "DCF 權重": "DCF 權重 0.3／0.7"}


def sens_post(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    f = os.path.join(here, "sens.json")
    S = json.load(open(f, encoding="utf-8")) if os.path.exists(f) else None
    if S is None:
        assert os.environ.get("QUANTA_OVR") or os.environ.get("QUANTA_NOSENS"), "quanta/sens.json 不存在：先跑 python3 quanta/sens_quanta.py"
    r = SE.max_row + 3
    SE.cell(r, 1, "目標價敏感度：業務主因（v0.4.7 新增；每個案例以 QUANTA_OVR 把 In 頁輸入改為敏感度低／高值後整本重建、LibreOffice 重算，讀加權目標價；quanta/sens_quanta.py → quanta/sens.json）").font = BOLD
    t0 = r; r += 1
    for j, h in enumerate(["案例", "設定（In 輸入）", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "快照狀態"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    tpc = "tp_final"
    rows = {}
    for lab_, lo, hi, key in BIZ:
        SE.cell(r, 1, lab_); SE.cell(r, 2, f"低：{lo}；高：{hi}（lo／hi＝In 頁敏感度低／高欄；x＝基準 × 倍數）")
        e = (S or {}).get("biz", {}).get(key)
        if e:
            for col, side in (("C", "低"), ("E", "高")):
                if side in e:
                    SE[f"{col}{r}"] = e[side]["tp"]; SE[f"{col}{r}"].number_format = "#,##0.00"; SE[f"{col}{r}"].font = BLUE
                    SE[f"{chr(ord(col) + 1)}{r}"] = f"={col}{r}/G{r}-1"; SE[f"{chr(ord(col) + 1)}{r}"].number_format = "+0.0%;-0.0%"
        SE[f"G{r}"] = f"={tpc}"; SE[f"G{r}"].number_format = "#,##0.00"
        SE[f"H{r}"] = f'=IF(ABS(G{r}-{(S or {}).get("_base", {}).get("tp", 0)!r})<=1E-6*ABS(G{r}),"CURRENT","STALE")'
        rows[key] = r; r += 1
    SE.cell(r, 1, ("快照狀態：快照基準加權目標價與本頁即時值相同為 CURRENT；改 In 頁輸入或模型後須重跑 python3 quanta/sens_quanta.py，再正式建置。"
                   "每個案例只改該列輸入，其他輸入維持基準；三情境與 DCF、Comparable 整本重算。")); r += 2
    # 加權目標價相對基準（HTML 敏感度主圖用）
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.4.7 起，版面規格 v1.3 H 節；只引用上方「業務主因」與「目標價參數」兩表，白話標籤，不回饋計算）").font = BOLD; tt = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源列"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for lab_, lo, hi, key in BIZ:
        sr = rows[key]
        SE.cell(r, 1, lab_)
        for col, scol in ((2, "D"), (3, "F")):
            if SE[f"{scol}{sr}"].value is not None:
                SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{sr}"); r += 1
    tp0 = lab(SE, "目標價參數敏感度"); hr = lab(SE, "參數", tp0, exact=True); q = hr + 1
    while SE.cell(q, 1).value and not str(SE.cell(q, 1).value).startswith("檢查"):
        t = str(SE.cell(q, 1).value)
        labl = next((v for k, v in TOR_LABEL.items() if t.startswith(k)), t)
        assert not re.search(r"(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+", labl), labl
        SE.cell(r, 1, labl)
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{q}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{q}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{q}"); r += 1; q += 1
    return tt


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, SIL, PLA, RVA, coord_rows):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; RC = wb["Recon"]; D = wb["DCF"]
    vb = lambda p, **k: f"Val!B{lab(V, p, **k)}"
    TPW, UP, REC = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    DCF, CMP, CMP0, PX = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("Comparable 目標價（市場主流同業法）"), vb("現價（")
    NTM12 = f"Val!B{lab(V, '現價 ÷ 未來 12 月 EPS')}"
    sv = lambda p, c: f"Sens!{c}{lab(SE, p)}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（v0.4.7 起；個股 HTML 版面規格 v1.3：橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字；公式帶數字；原展示層各列保留）").font = BOLD
    r0 = r; r += 1
    M = {}

    def put(key, f):
        nonlocal r
        RD.cell(r, 1, f); M[key] = r; r += 1
    put("hero", F(f'{REC}', T("：加權目標價 "), f'TEXT({TPW},"0")', T(" 元，比現價 "), f'TEXT({PX},"0")', T(" 元高 "), f'TEXT({UP},"0%")',
                  T("；DCF "), f'TEXT({DCF},"0")', T(" 元與 Comparable（相對倍數調整後）"), f'TEXT({CMP},"0")', T(" 元各占一半。盈餘高於法人是上檔的主因。")))
    put("rec", F(T("投資建議依加權目標價相對現價的上檔空間判定：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、低於 −50% 強力放空。目前上檔空間 "),
                 f'TEXT({UP},"+0.0%;-0.0%")', T("，投資建議 "), f'{REC}', T("。")))
    put("o1", F(T("1｜目標價 "), f'TEXT({TPW},"0")', T(" 元、上檔 "), f'TEXT({UP},"0%")', T("：DCF "), f'TEXT({DCF},"0")', T(" 元、Comparable（調整後）"), f'TEXT({CMP},"0")',
                T(" 元各占一半；現價只有本模型未來 12 個月 EPS 的 "), f'TEXT({NTM12},"0.0")', T(" 倍。只看市場主流同業法的 Comparable 為 "), f'TEXT({CMP0},"0")', T(" 元。")))
    put("o2", F(T("2｜盈餘比法人高是上檔的主因：FY2027 EPS "), f'TEXT(Recon!C{lab(RC, "FY2027 EPS vs 法人")},"0.00")', T(" 元，比法人（10-08）高 "),
                f'TEXT(Recon!D{lab(RC, "FY2027 EPS vs 法人")},"0%")', T("；淨利 ÷ 營業利益用 "), 'TEXT(nio_f,"0.000")', T("（業外、稅與少數股權），第三季法說是第一個檢驗點；若盈餘只到共識，上檔大幅縮小。")))
    put("o3", F(T("3｜三情境從 FY2028 才分開：FY2028 EPS 樂觀、基準、保守約 "), f'TEXT(PL!F{lab(wb["PL"], "EPS", lab(wb["PL"], "情境 A："), exact=True)},"0.0")', T("、"),
                f'TEXT(PL!F{PLA["R"]["eps"]},"0.0")', T("、"), f'TEXT(PL!F{lab(wb["PL"], "EPS", lab(wb["PL"], "情境 C："), exact=True)},"0.0")', T(" 元。要追蹤份額與部署量：AI 部署量 ×0.8 時 FY2028 EPS 變動 "),
                f'TEXT({sv("AI 部署 ×0.8", "L")},"+0.0%;-0.0%")', T("，份額相對值降到 0.75 時 "), f'TEXT({sv("份額流失（低）", "L")},"+0.0%;-0.0%")', T("。")))
    put("o4", F(T("4｜DCF 為標準版（無槓桿自由現金流折現）：WACC "), 'TEXT(wacc,"0.0%")', T("、永續成長率 "), 'TEXT(g_tv,"0.00%")', T("，終值現值占情境 B 企業價值 "),
                f'TEXT(DCF!C{lab(D, "終值現值占 EV", lab(D, "情境 B（基準）"))},"0%")', T("；WACC 低／高一個百分點時加權目標價 "), f'TEXT({sv("WACC ±1pt", "D")},"+0%;-0%")', T("／"),
                f'TEXT({sv("WACC ±1pt", "F")},"+0%;-0%")', T("。")))
    if coord_rows:
        CR = coord_rows
        put("o5", F(T("5｜對帳中：Siliconomics 標示 CY"), 'Coord!B' + str(CR["yrs"]), T(" 對帳中；另一側讀數（Wattonomics 換算的部署量）目標價 "), f'TEXT(Coord!C{CR["tp"]},"0.00")', T(" 元（"), f'Coord!C{CR["rec"]}', T("），現行基準讀數 "),
                    f'TEXT(Coord!B{CR["tp"]},"0.00")', T(" 元（"), f'Coord!B{CR["rec"]}', T("）；個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    else:
        put("o5", "5｜本次建置不含 Coord 頁（三層協調兩讀數快照重建用）。")
    rb = lab(wb["PL"], "情境 B：")
    pr = lambda p: lab(wb["PL"], p, rb)
    put("drv", F(T("分部營收怎麼推導（情境 B）：AI 伺服器營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 每 GW ODM 整櫃價值 × 廣達有效份額 ×（1 − 客供比例）；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("AI 伺服器營收")},"#,##0")', T(" 十億元。非 AI 營收＝一般伺服器＋筆電＋其他，由 FY2026 錨點乘各自成長率；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("非 AI 營收")},"#,##0")', T(" 十億元。兩者相加＝合併營收 "), f'TEXT(PL!E{PLA["R"]["rev"]},"#,##0")', T(" 十億元；分項加總與合併營收的檢查列為 0。")))
    put("ver", F(T("v0.4.8（2026-11-14）：三層協調 C0 補公司實績對帳（規則 v1.4.2）：Coord 頁末新增「C0〔實績〕」表，以 2026 年 1–10 月月營收與法說揭露比例換算的 AI 伺服器營收對 Siliconomics 整櫃組裝市場 × 廣達份額的同期隱含值，主讀數與第二讀數並列判定；原 C0 同源列改標「讀取層級（同源，不判定）」；其餘數值、加權目標價與評等不變。v0.4.7（2026-11-14）：版面規格 v1.3 全面改版（章節順序、橫幅 ≤200 字、重點觀察 ≤5 點、標準損益表 FY2023A–FY2031E、反向 DCF 頁、四個目標價與現價圖、敏感度改加權目標價相對基準低案／高案主圖）；"
                 "三層協調新機制 C0：Coord 頁新增「由下往上對帳」（本模型組裝市場與 Siliconomics SIF_AsmMkt_Base 等逐年比較，差距超過 ±15% 標對帳中）；"
                 "xlsx 只新增 Hist、RevDCF 頁與各頁末新增區塊，既有數值、加權目標價 512.56 與投資建議（強力買進）不變；README A1 標題段版號由 v0.2.1 更正為 v0.4.7（現行 v0.4.8）。")))
    return M


# ---------- Coord：C0（實績）——公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.4.8 新增） ----------
TN = "https://finance.technews.tw/2026/{}/qci-2382-2026{}-financial-report/"
MONTHS = [   # (月份, 營收 NT$B, 來源網址, 說明)
    ("2026-01", 230.835, "https://uanalyze.com.tw/articles/2838847558；" + TN.format("03/09", "02"), "Derived：公司公布前 2 月累計 4,464.18 億元 − 2 月 2,155.83 億元 ＝ 2,308.35 億元（優分析標題：1 月 2,308 億元）"),
    ("2026-02", 215.583, TN.format("03/09", "02"), "2 月營收 2,155.83 億元，累計前 2 月 4,464.18 億元"),
    ("2026-03", 362.803, TN.format("04/08", "03"), "3 月營收 3,628.03 億元，累計前 3 月 8,092.21 億元"),
    ("2026-04", 339.921, TN.format("05/08", "04"), "4 月營收 3,399.21 億元"),
    ("2026-05", 311.481, TN.format("06/08", "05"), "5 月營收 3,114.81 億元"),
    ("2026-06", 385.191, TN.format("07/07", "06"), "6 月營收 3,851.91 億元，累計前 6 月約 1.85 兆元"),
    ("2026-07", 366.274, "https://uanalyze.com.tw/articles/3817752990", "7 月營收 3,662.74 億元"),
    ("2026-08", 423.971, "https://statementdog.com/news/17848", "8 月營收 4,239.71 億元（歷史新高）"),
    ("2026-09", 325.3, "https://statementdog.com/news/17848；https://uanalyze.com.tw/articles/2106956697", "9 月營收 3,253 億元（億元位數；前 9 月累計 29,613.58 億元）"),
    ("2026-10", 353.25, TN.format("11/09", "10"), "10 月營收 3,532.50 億元，月增 8.6%，累計前 10 月約 3.31 兆元"),
]
def srcname(u):
    n = []
    if "technews" in u: n.append("TechNews 月營收報導")
    if "uanalyze" in u: n.append("優分析月營收報導")
    if "statementdog" in u: n.append("財報狗月營收報導")
    return "、".join(n)


VERIF = "Verified（公司月營收公告；經媒體轉載，MOPS 原文未直接讀取）"


def c0_actuals(wb, C, r, l3, rs_l, d5):
    out = {"sec": r}
    C.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.4.8 新增；個股側只用已公布的公司實績與法說揭露，不用賣方估計或模型預測）").font = BOLD; r += 1
    C.cell(r, 1, ("方法：個股側＝公司月營收（法定揭露）× 法說揭露的伺服器占營收比例 × AI 伺服器占伺服器比例＝公司實績側 AI 伺服器營收；廣達未揭露 AI 伺服器營收金額，只能由總營收乘揭露比例換算（口徑換算，等級 Interested-party／Derived）。"
                  "層級側＝Siliconomics 整櫃組裝市場（基準層，CY2026）× 廣達份額（份額表）× 匯率＝CY2026 全年隱含 AI 伺服器營收。"
                  "主讀數＝2026 年已公布月份（1–10 月）累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12（線性時差近似，Derived）；第二讀數＝最近已公布一季（3Q26，7–9 月）× 4 vs 層級 CY2026。成長年中主讀數偏低、第二讀數偏高，兩者夾住真值。"
                  "判定：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致。"))
    r += 1
    C.cell(r, 1, ("提醒：層級側廣達份額（Siliconomics 份額表）是廣達模型以 FY2026 錨點（1H26 實績＋3Q 營收實績＋4Q 假設，乘 85% × 80% 比例）校準後存入 Siliconomics 的快照，所以層級側全年值等於本模型 FY2026 AI 伺服器營收錨點（In 頁／Q 頁）。"
                  "本表的檢驗力在於：實績側只用公司揭露的比例（伺服器 ≥ 80% 取 80%；AI 占伺服器 75–80% 取中點），不用錨點採用的 85% 與 80%，因此能看出錨點比例是否高於公司揭露。"
                  "2H26 起部分高價料件改客供（公司：比重仍低）會使認列營收低於毛額價值，使實績側偏低（Assumed 影響，不調整）。"))
    r += 1
    for j, h in enumerate(["項目", "主讀數（1–10 月累計）", "第二讀數（3Q26 × 4）", "單位與口徑", "來源與證據等級", "出處網址（只在 xlsx）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    out["head"] = r; r += 1
    C.cell(r, 1, "4.1 公司實績：月營收（NT$B；Excel 儲存格單位十億新台幣）").font = BOLD; r += 1
    m0 = r
    for mo, v, src, note in MONTHS:
        C.cell(r, 1, f"　{mo} 月營收"); C[f"B{r}"] = v; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.000"
        C.cell(r, 4, "NT$B；合併營收（公司月營收公告）；" + note); C.cell(r, 5, VERIF + "；出處：" + srcname(src)); C.cell(r, 6, src)
        r += 1
    m1 = r - 1
    out["m0"], out["m1"] = m0, m1
    C.cell(r, 1, "　2026-01～2026-10 累計（主讀數用）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m1})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布前 10 月累計約 3.31 兆元（與本列相符）"); C.cell(r, 5, "Verified（加總）"); cum = r; r += 1
    C.cell(r, 1, "　3Q26（7–9 月）合計（第二讀數用）"); C[f"B{r}"] = f"=SUM(B{m0+6}:B{m0+8})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布 3Q26 營收約 1.1155 兆元"); C.cell(r, 5, "Verified（加總）"); q3 = r; r += 1
    C.cell(r, 1, "　核對：1–6 月合計 − 公司公布 1H26 營收（Q 頁）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+5})-Q!C4"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（公司以億元位數公布，容許 ±0.05）"); C.cell(r, 5, "核對列"); ck1 = r; r += 1
    C.cell(r, 1, "　核對：1–9 月合計 − 公司公布前 9 月累計 2,961.358（29,613.58 億元）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+8})-2961.358"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（9 月以億元位數列示，容許 ±0.05）"); C.cell(r, 5, "核對列（出處：財報狗月營收報導）"); C.cell(r, 6, "https://statementdog.com/news/17848"); ck2 = r; r += 1
    C.cell(r, 1, "　核對：3Q26 合計 − 3Q26 營收輸入（In 頁 q3_rev）"); C[f"B{r}"] = f"=B{q3}-q3_rev"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（In 頁 q3_rev 取 1,115.5）"); C.cell(r, 5, "核對列"); ck3 = r; r += 1
    out["cum"], out["q3"] = cum, q3
    C.cell(r, 1, "4.2 口徑換算參數（公司法說揭露比例；原簡報未直接讀取）").font = BOLD; r += 1
    C.cell(r, 1, "　伺服器占營收比例（公司：自 1Q26 起穩定超過八成；取下限）"); C[f"B{r}"] = 0.80; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司揭露為「超過八成」，只取下限 80%（保守）；In 頁伺服器占營收的 85% 來自凱基預測，不是公司揭露"); C.cell(r, 5, "Interested-party（法說摘要：富果 2Q26 法說備忘錄）"); C.cell(r, 6, "https://blog.fugle.tw/post/earnings-call-2382-2026-08-13"); srv = r; r += 1
    C.cell(r, 1, "　AI 伺服器占伺服器營收：1H26 揭露區間下限"); C[f"B{r}"] = 0.75; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司：2Q26 AI 伺服器占伺服器 75–80%"); C.cell(r, 5, "Interested-party（法說摘要：優分析 2Q26 法說報導）"); C.cell(r, 6, "https://uanalyze.com.tw/articles/5252953366"); a_lo = r; r += 1
    C.cell(r, 1, "　AI 伺服器占伺服器營收：1H26 揭露區間上限（全年看法約 80%）"); C[f"B{r}"] = 0.80; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司：下半年比重持續提升、全年維持 80% 的看法"); C.cell(r, 5, "Interested-party（法說摘要：優分析 2Q26 法說報導）"); C.cell(r, 6, "https://uanalyze.com.tw/articles/5252953366"); a_hi = r; r += 1
    C.cell(r, 1, "　AI 伺服器占營收比例（主：下限 80% × 區間中點）"); C[f"B{r}"] = f"=B{srv}*AVERAGE(B{a_lo},B{a_hi})"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "＝伺服器占營收下限 × AI 占伺服器區間中點；Derived（換算）"); C.cell(r, 5, "Derived／Interested-party"); ratio = r; r += 1
    C.cell(r, 1, "　參考：AI 伺服器占營收比例（下限：80% × 75%）"); C[f"B{r}"] = f"=B{srv}*B{a_lo}"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司揭露可得的最低換算；只作參考，不判定"); C.cell(r, 5, "Derived"); ratio_lo = r; r += 1
    C.cell(r, 1, "　參考：AI 伺服器占營收比例（本模型 In 頁：伺服器占營收 × AI 占伺服器）"); C[f"B{r}"] = "=srv_26*ais_26"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "本模型 FY2026 錨點採用的比例；只作參考，不判定"); C.cell(r, 5, "Interested-party（凱基預測 85%、公司 80%）"); ratio_md = r; r += 1
    out["ratio"] = ratio
    C.cell(r, 1, "4.3 對帳（實績側 vs 層級隱含）").font = BOLD; r += 1
    def row2(label, fb, fc, fmt, unit, src, bold=False):
        nonlocal r
        C.cell(r, 1, label)
        C[f"B{r}"] = fb; C[f"C{r}"] = fc
        for c in "BC": C[f"{c}{r}"].number_format = fmt
        C.cell(r, 4, unit); C.cell(r, 5, src)
        if bold:
            for c in "ABC": C[f"{c}{r}"].font = BOLD
        r += 1; return r - 1
    a_rev = row2("　公司實績：營收（主：1–10 月累計；第二：3Q26 × 4）", f"=B{cum}", f"=B{q3}*4", "#,##0.0", "NT$B；合併營收", VERIF)
    a_rat = row2("　× AI 伺服器占營收比例", f"=B{ratio}", f"=B{ratio}", "0.0%", "見 4.2", "Derived／Interested-party")
    a_ai = row2("　＝ 公司實績側 AI 伺服器營收", f"=B{a_rev}*B{a_rat}", f"=C{a_rev}*C{a_rat}", "#,##0.0", "NT$B；主讀數為 10 個月累計；第二讀數為年化", "Derived（口徑換算）", True)
    l_mk = row2("　層級：整櫃組裝市場 CY2026（Siliconomics 基準層）", f"=B{l3}", f"=B{l3}", "#,##0.0", "US$B；連到下方「讀取層級」第 (3) 項的 Siliconomics 層列", "Siliconomics v0.19；Derived")
    l_sh = row2("　× 廣達份額（Siliconomics 份額表）", f"=B{rs_l}", f"=B{rs_l}", "0.0%", "連到下方 (5) 的 Siliconomics 份額列", "Siliconomics v0.19；Derived（由本模型 FY2026 錨點校準，見提醒）")
    l_fx = row2("　× 匯率（新台幣 ÷ 美元；In 頁 fx）", "=fx", "=fx", "0.0", "NT$/US$", "Derived（與台達電模型相同；公司換算隱含約 32.7）")
    l_fy = row2("　＝ 層級隱含 CY2026 全年 AI 伺服器營收", f"=B{l_mk}*B{l_sh}*B{l_fx}", f"=C{l_mk}*C{l_sh}*C{l_fx}", "#,##0.0", "NT$B；全年", "Derived", True)
    l_tf = row2("　× 期間係數（主：已公布月數 ÷ 12；第二：1）", f"=COUNT(B{m0}:B{m1})/12", "=1", "0.000", "線性時差近似（層級值為日曆年）", "Derived")
    l_pd = row2("　＝ 層級側同期隱含 AI 伺服器營收", f"=B{l_fy}*B{l_tf}", f"=C{l_fy}*C{l_tf}", "#,##0.0", "NT$B", "Derived", True)
    out["fy"] = l_fy
    C.cell(r, 1, "　差距（實績側 ÷ 層級側 − 1）"); C[f"B{r}"] = f"=B{a_ai}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_ai}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"; C[f"{c}{r}"].font = BOLD
    C.cell(r, 4, "正值＝公司實績高於層級隱含"); out["d"] = r; d = r; r += 1
    C.cell(r, 1, "　判定（兩讀數都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致）")
    C[f"B{r}"] = f'=IF(AND(ABS(B{d})>0.15,ABS(C{d})>0.15,SIGN(B{d})=SIGN(C{d})),"對帳中",IF(OR(ABS(B{d})>0.15,ABS(C{d})>0.15),"時差邊界（參考）","一致"))'
    C[f"B{r}"].font = BOLD; C.cell(r, 4, "判定只看本表主比例（AI 占營收＝下限 80% × 區間中點）"); C.cell(r, 5, "規則 v1.4.2 C0 第 2 點、工作單 20261114"); out["judge"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：AI 占營收取下限（80% × 75%）時的差距"); C[f"B{r}"] = f"=B{a_rev}*B{ratio_lo}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_rev}*B{ratio_lo}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "公司揭露可得的最低換算"); out["ref_lo"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：AI 占營收取本模型 In 頁比例（85% × 80%）時的差距"); C[f"B{r}"] = f"=B{a_rev}*B{ratio_md}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_rev}*B{ratio_md}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "本模型錨點採用的比例（凱基預測 85%）；非公司揭露"); out["ref_md"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：公司實績側營收 1–10 月累計 ÷ 本模型 FY2026 營收錨點（Q 頁）"); C[f"B{r}"] = f"=B{cum}/a26_rev"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "NT$B 比；錨點＝1H26 實績＋3Q26 實績＋4Q 假設；11–12 月合計由錨點隱含"); C.cell(r, 5, "Derived"); out["ref_rev"] = r; r += 1
    out["ck"] = (ck1, ck2, ck3); out["end"] = r
    C.column_dimensions["A"].width = max(C.column_dimensions["A"].width or 0, 60)
    return out


# ---------- Coord：C0 由下往上對帳 ----------
def coord_c0(wb, SIL, R):
    C = wb["Coord"]; P = wb["PL"]
    C.column_dimensions["F"].width = 46; C.column_dimensions["G"].width = 70
    rb = lab(P, "情境 B：", exact=False); ra = lab(P, "情境 A：")
    gb = lambda p, ex=False: lab(P, p, rb, exact=ex); ga = lambda p, ex=False: lab(P, p, ra, exact=ex)
    N = SIL.NAMES
    r = C.max_row + 3
    C.cell(r, 1, "C0 由下往上對帳（三層基準協調規則 v1.4.2 C0 第 2 點；v0.4.7 新增、v0.4.8 改版）：比較基準是公司實績（第 4 節 C0〔實績〕：月營收與法說揭露比例 vs 層級隱含值）；本模型自層級讀入的同源值只列在第 2 節「讀取層級」，不判定、不計對帳中。差距超過 ±15% 標「對帳中」，不改上游值，由主控開對帳單。").font = BOLD
    c0 = r; r += 1
    C.cell(r, 1, "1. 讀哪一層、哪一版").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "版本／檔案", "本模型用途", "備註"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    for a, b, c, d, e in (
        ("硬體股讀哪一層", "Siliconomics（零組件層；基準）", ver, "產業部署量、AI IT 資本支出、整櫃組裝市場、顆數係數、份額表、對帳中旗標（reader.py）", "下游只註明版本、不自行重估（專案規則第 3 條）"),
        ("AI IT 資本支出 SIF_AIITCapex_Base（情境 A 供給側；情境 B 以顆數係數縮放）", "US$B；見下表", "markets/network v0.2.3（Mkt!C27:F27）", "In 頁 itc_26–itc_29", "數值與 Siliconomics 基準相同（核對見下表）"),
        ("每 GW IT 資本支出 SIF_ITCapexPerGW、每 GW 機架數、整櫃買方價", "33.4／37.6／37.6／37.6 US$B", "markets/chip_bridge v0.2.3（Tokenomics v21.x 經橋接表）", "In 頁 itpg_gb、itpg_vr、racks_gb、racks_vr、rack_gb、rack_vr", "唯讀"),
        ("情境顆數係數（2028、2029）", "A 1.0／1.0；B 0.8／0.71；C 0.6／0.6", "tsmc v0.5.4（In!E37:E42；Siliconomics v0.19 基準係數）", "In 頁 um28_A–um29_C", "情境 B＝Siliconomics 基準"),
        ("份額快照 asm_qush_26、asm_qurel_27–29", "0.1339／1.0／0.95／0.90", ver, "In 頁 rel_27–rel_29；FY2026 有效份額由公司揭露比例校準", "公司模型以自身營收校準，Siliconomics 只存快照"),
        ("Wattonomics 錨定 SIF_WTAnchor_Base（wta_27–29）", "0.72／0.65／0.63", ver, "只用於 Sens「Wattonomics 錨定」案例，不進目標價", "Wattonomics 現行版本 v0.3.2（本模型不直接讀）"),
        ("對帳中旗標 SIF_RecoFlag_<yyyy>、SIF_DeployGW_AltRatio_<yyyy>", "旗標 0／1／1／1；比值 0.88／0.62／0.67／0.67", ver, "C4 兩讀數（本頁上方）", "另一側讀數＝Wattonomics 換算的產業 GW")):
        for j, v in enumerate((a, b, c, d, e), 1):
            C.cell(r, j, v)
        r += 1
    r += 1
    C.cell(r, 1, "2. 讀取層級（同源，不判定；CY＝FY；情境 B 基準對 Siliconomics 基準；金額 US$B）：本模型自層級讀入的值，差距恆為 0，無檢驗力，故不進 C0 判定（規則 v1.4.2）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位與口徑", "來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    flags = []
    YC = "BCDE"; PC = "DEFG"   # Coord 欄 B–E ↔ PL 欄 D–G（FY2026–FY2029）

    def block(title, model_f, layer_vals, unit, src, fmt="#,##0.0", layer_label="Siliconomics 層", tol=0.15, ref=False):
        nonlocal r
        C.cell(r, 1, title).font = BOLD; r += 1
        C.cell(r, 1, "　本模型"); C.cell(r, 6, unit); C.cell(r, 7, src)
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = model_f(PC[k]); C[f"{c}{r}"].number_format = fmt
        m = r; r += 1
        C.cell(r, 1, f"　{layer_label}")
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = layer_vals[k]; C[f"{c}{r}"].number_format = fmt; C[f"{c}{r}"].font = BLUE
        l = r; r += 1
        C.cell(r, 1, "　差距（本模型 ÷ 層 − 1）")
        for c in YC:
            C[f"{c}{r}"] = f"={c}{m}/{c}{l}-1"; C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
        d = r; r += 1
        if ref:
            C.cell(r, 1, f"　判定（參考：|差距| > {tol:.0%}＝層級差距；不標本公司對帳中）")
            for c in YC:
                C[f"{c}{r}"] = f'=IF(ABS({c}{d})>{tol},"層級差距","一致")'
            r += 1
        else:
            C.cell(r, 1, "　判定（讀取層級，同源，不判定；不計入對帳中項目數）")
            for c in YC:
                C[f"{c}{r}"] = "同源（不判定）"
            flags.append(r); r += 1
        return m, l, d
    dg = lambda k: [N[k][i] for i in range(4)]
    block("(1) AI IT 資本支出（情境 B）", lambda c: f"=PL!{c}{gb('AI IT 資本支出 × 情境顆數係數')}", dg("SIF_AIITCapex_Base"), "US$B；加速器顆數 × 每顆 IT capex（含 ASIC）× 顆數係數", "Siliconomics 基準層（AI IT 資本支出）；Derived（markets/network v0.2.3 讀入）")
    block("(2) 部署 GW（IT 電力；情境 B）", lambda c: f"=PL!{c}{gb('部署 GW（IT 電力）')}", dg("SIF_DeployGW_Base"), "GW；AI IT 資本支出 ÷ 每 GW IT 資本支出", "Siliconomics 基準層（部署 GW）；Derived", "0.00")
    m3, l3, d3 = block("(3) 整櫃組裝市場（ODM 價；情境 B）", lambda c: f"=PL!{c}{gb('AI 伺服器組裝市場（ODM 價，美元）')}", dg("SIF_AsmMkt_Base"), "US$B；部署 GW × 每 GW ODM 整櫃價值（含 ASIC 機架，Assumed：每 GW 價值相同）",
                       "Siliconomics 基準層（整櫃組裝市場）；Derived／Assumed（每 GW 價值）")
    block("(4) 部署 GW（情境 A；供給側讀數）", lambda c: f"=PL!{c}{ga('部署 GW（IT 電力）')}", [N[f"SIF_DeployGW_Supply_{y}"] for y in (2026, 2027, 2028, 2029)], "GW；情境 A＝原始顆數（供給側）", "Siliconomics 供給側部署 GW 逐年值；Derived", "0.00")
    # 份額與未建模殘差
    C.cell(r, 1, "(5) 份額與恆等式（已建模 ODM 加總＋未建模殘差＝市場）").font = BOLD; r += 1
    qs = SIL.V["asm_qush_26"]["base"]; qrel = [1.0] + [SIL.V[f"asm_qurel_{y}"]["base"] for y in (27, 28, 29)]
    C.cell(r, 1, "　廣達有效份額（本模型，情境 B）"); C.cell(r, 6, "%；廣達 AI 伺服器毛額價值 ÷ 組裝市場"); C.cell(r, 7, "PL 有效份額；FY2026 由公司揭露比例校準（Derived），2027–29 為份額相對值（Assumed）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = f"=PL!{PC[k]}{gb('廣達有效份額', True)}"; C[f"{c}{r}"].number_format = "0.0%"
    rs_m = r; r += 1
    C.cell(r, 1, "　廣達有效份額（Siliconomics 份額表：asm_qush_26 × asm_qurel）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = qs * qrel[k]; C[f"{c}{r}"].number_format = "0.0%"; C[f"{c}{r}"].font = BLUE
    rs_l = r; r += 1
    C.cell(r, 1, "　差距（百分點）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}-{c}{rs_l}"; C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    r += 1
    C.cell(r, 1, "　已建模 ODM 份額合計（廣達、鴻海、緯創、緯穎；SIF_AsmShareSum）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = N["SIF_AsmShareSum"][k]; C[f"{c}{r}"].number_format = "0.0%"; C[f"{c}{r}"].font = BLUE
    rs_s = r; C.cell(r, 7, "Siliconomics 已建模份額合計；Derived／Assumed"); r += 1
    C.cell(r, 1, "　未建模殘差（1 − 合計；Dell、Supermicro、Celestica 等；明列，不為 0）")
    for c in YC:
        C[f"{c}{r}"] = f"=1-{c}{rs_s}"; C[f"{c}{r}"].number_format = "0.0%"
    rs_r = r; r += 1
    C.cell(r, 1, "　恆等式核對：已建模合計 ＋ 殘差 − 100%（應為 0）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_s}+{c}{rs_r}-1"; C[f"{c}{r}"].number_format = "0.0000"
    r += 1
    C.cell(r, 1, "　廣達占已建模合計")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}/{c}{rs_s}"; C[f"{c}{r}"].number_format = "0.0%"
    r += 1
    C.cell(r, 1, "　廣達 AI 伺服器毛額價值（US$B；＝組裝市場 × 有效份額）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{m3}*{c}{rs_m}"; C[f"{c}{r}"].number_format = "#,##0.0"
    C.cell(r, 7, "本模型：PL 情境 B 組裝市場 × 有效份額；上限為市場（恆等式：加總不得超過上層）"); r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; r += 1
    C.cell(r, 1, "　下列差距是 Wattonomics 與 Siliconomics 兩層之間的差距，不是廣達自己的量對其所讀 Siliconomics 值的差距；本公司依 C4 並列兩讀數，不裁決、不改上游。"); r += 1
    m5, l5, d5 = block("(6) 部署 GW（情境 B）對 Wattonomics 換算值（兩層之間的差距）", lambda c: f"=PL!{c}{gb('部署 GW（IT 電力）')}", [N[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)],
                       "GW；Siliconomics 的需求側部署 GW 逐年值＝Wattonomics 資本支出與電力路徑換算的需求側讀數", "Wattonomics v0.3.2 經 Siliconomics 換算；Derived（需求側）", "0.00", layer_label="Wattonomics 換算值（需求側）", ref=True)
    r += 1
    C.cell(r, 1, "3. 結論（C0；判定見第 4 節 C0〔實績〕表）").font = BOLD; r += 1
    C.cell(r, 1, "對帳中項目數（C0〔實績〕判定為「對帳中」者；讀取層級的同源列不判定、不計入）")
    n_other = r; r += 1
    C.cell(r, 1, "C0 結論（白話）")
    c0_concl = r; r += 2
    A4 = c0_actuals(wb, C, r, l3, rs_l, d5)
    C[f"B{n_other}"] = f'=IF(B{A4["judge"]}="對帳中",1,0)'
    C[f"B{c0_concl}"] = F(T("以公司實績對層級值（第 4 節）：2026 年 1–10 月月營收累計換算的 AI 伺服器營收，對 Siliconomics 整櫃組裝市場 × 廣達份額的同期隱含值，差距 "), f'TEXT(B{A4["d"]},"+0.0%;-0.0%")',
                          T("（主讀數）；最近一季（3Q26）× 4 對全年隱含值，差距 "), f'TEXT(C{A4["d"]},"+0.0%;-0.0%")', T("（第二讀數）；判定「"), f'B{A4["judge"]}',
                          T("」（門檻 ±15%）。本模型自層級讀入的同源值（第 2 節）不判定。與 Wattonomics 換算的需求側讀數相比，部署 GW 在 CY2027–CY2029 差距 "),
                          f'TEXT(C{d5},"+0%")', T("、"), f'TEXT(D{d5},"+0%")', T("、"), f'TEXT(E{d5},"+0%")',
                          T("，超過 ±15%：這是兩層之間的差距（見「層級之間」小節），已由 Siliconomics 對帳單 r3 處理，不計入本公司 C0；本模型依 C4 並列兩讀數目標價（上方），不改上游值，不裁決哪一側正確。"))
    return dict(c0=c0, concl=c0_concl, flags=flags, d3=d3, d5=d5, m3=m3, l3=l3, **A4)


# ---------- 版本字串 ----------
def version(wb, here, coord_rows):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v0.2.1" in head, head
    a1 = head.replace("v0.2.1", f"v{VER}", 1) + a1[len(head):]
    a1 += (f"；v{VER}（{BUILT}）：三層協調 C0 補公司實績對帳（Coord 頁末「C0〔實績〕」：1–10 月月營收與法說揭露比例換算的 AI 伺服器營收 vs Siliconomics 整櫃組裝市場 × 廣達份額；"
           "原同源列改標「讀取層級（同源，不判定）」）；v0.4.7：版面規格 v1.3 全面改版與 C0 由下往上對帳，新增 Hist、RevDCF 頁；既有數值、加權目標價與評等不變")
    RD["A1"].value = a1
    return a1


def apply(wb, SIL, here, coord_rows):
    HR = hist_is(wb)
    PLA = pl_standard(wb, SIL)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_post(wb, here)
    C0 = coord_c0(wb, SIL, coord_rows) if coord_rows else None
    RMA = readme13(wb, SIL, PLA, RVA, coord_rows)
    version(wb, here, coord_rows)
    SO = wb["Sources"]; rs = SO.max_row + 1
    SO.cell(rs, 1, "[Verified] 廣達 2026 年 1–10 月月營收（C0〔實績〕；TechNews 各月、優分析 7 月與 1 月、財報狗 8–9 月；MOPS 原文未直接讀取，網址見 Coord 頁第 4.1 節各列）")
    SO.cell(rs, 2, "https://finance.technews.tw/2026/11/09/qci-2382-202610-financial-report/")
    SO.cell(rs + 1, 1, "[Interested-party] 法說摘要：伺服器占營收自 1Q26 起超過八成（富果 2Q26 法說摘要）；AI 伺服器占伺服器 75–80%、全年約 80%（優分析 2Q26 法說）：C0〔實績〕換算比例用，原簡報未直接讀取")
    SO.cell(rs + 1, 2, "https://blog.fugle.tw/post/earnings-call-2382-2026-08-13；https://uanalyze.com.tw/articles/5252953366")
    wb._v047 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, C0=C0, RM=RMA)
    return f"v047_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 主圖表 {ST}；README v1.3 {len(RMA)} 列；C0 {'有' if C0 else '無（無 Coord 頁）'}"
