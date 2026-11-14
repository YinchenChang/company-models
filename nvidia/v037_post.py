# NVIDIA v0.3.7（2026-11-14）後處理：個股 HTML 版面規格 v1.3 試作（Andy 2026-11-14 08:11；docs/plan/個股HTML版面規格.md「v1.3 規格」）。
# 只新增，不改既有數值格：
#   Hist 頁末：損益表歷史（FY2023–FY2026，SEC XBRL companyfacts／10-K；標準損益表的歷史三年與 FY2024 年增率）
#   PL 頁末：標準損益表（FY2024–FY2026 實績＋FY2027–FY2031 情境 B；Non-GAAP＝模型口徑）、比率區、GAAP 對照列、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑，封閉解與二分法）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   README 頁末：v1.3 展示層白話列（橫幅說明 ≤200 字、重點觀察 ≤5 點 ≤750 字、投資建議條件式白話、分部營收推導說明、版本沿革）
# 版本字串（README!A1、build_id）與成品日（Hist h_built）改為 v0.3.7／2026-11-14。build_nvidia.py 在 v035_post 之後、存檔前呼叫 apply(wb, here)。
import hashlib, os
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3.7", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SEC = "https://www.sec.gov/Archives/edgar/data/1045810/"
XBRL = "SEC XBRL companyfacts（CIK 0001045810；https://data.sec.gov/api/xbrl/companyfacts/CIK0001045810.json）10-K 年度值"
K10_26 = "NVIDIA 10-K（FY2026，accession 0001045810-26-000021；FY2024、FY2025 為比較欄）"
K10_25 = "NVIDIA 10-K（FY2025，accession 0001045810-25-000023；FY2023 為比較欄）"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def lab(ws, prefix, start=1, col=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix): return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def F(*parts):
    """公式字串：parts 為 ("t", 文字) 或 Excel 運算式（字串）。"""
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


# ---------- Hist：損益表歷史（v0.3.7） ----------
# (鍵, 項目, FY2024, FY2025, FY2026, FY2023, XBRL 標籤)；金額十億美元、EPS 美元、股數十億股
HIS = [
    ("rev", "營收（GAAP）", "=h_rev_24", "=h_rev_25", "=h_rev_26", 26.974, "Revenues"),
    ("cogs", "營業成本（GAAP）", 16.621, 32.639, 62.475, 11.618, "CostOfRevenue"),
    ("gp", "毛利（GAAP）", 44.301, 97.858, 153.463, 15.356, "GrossProfit"),
    ("opx", "營業費用（GAAP，研發＋銷管）", 11.329, 16.405, 23.076, 11.132, "OperatingExpenses"),
    ("oi", "營業利益（GAAP）", 32.972, 81.453, 130.387, 4.224, "OperatingIncomeLoss"),
    ("nonop", "其他收益淨額合計（GAAP，含權益證券損益）", 0.846, 2.573, 11.063, -0.043, "NonoperatingIncomeExpense"),
    ("eqg", "其中：權益證券（非上市與上市）淨損益", 0.238, 1.030, 8.918, -0.045, "GainLossOnInvestments"),
    ("pti", "稅前利益（GAAP）", 33.818, 84.026, 141.450, 4.181, "IncomeLossFromContinuingOperationsBeforeIncomeTaxes…"),
    ("tax", "所得稅（GAAP）", 4.058, 11.146, 21.383, -0.187, "IncomeTaxExpenseBenefit"),
    ("ni", "稅後淨利（GAAP）＝稅前 − 所得稅", "=C{pti}-C{tax}", "=D{pti}-D{tax}", "=E{pti}-E{tax}", "=F{pti}-F{tax}", "Derived"),
    ("eps", "EPS（GAAP 完全稀釋，公司公告；美元，2024-06 一拆十後）", 1.19, 2.94, 4.90, 0.17, "EarningsPerShareDiluted"),
    ("dil", "稀釋加權股數（十億股）", 24.940, 24.804, 24.514, 25.070, "WeightedAverageNumberOfDilutedSharesOutstanding"),
]


def hist_is(wb):
    H = wb["Hist"]
    r = H.max_row + 2
    H.cell(r, 1, "損益表歷史（v0.3.7 起；PL 頁「標準損益表」的歷史三年與 FY2024 年增率用；FY2026 欄只作核對，標準損益表的 FY2026 取 Q 頁四季合計）").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "FY2023（只供年增率）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v24, v25, v26, v23, tag in HIS:
        H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "24", v24), ("D", "25", v25), ("E", "26", v26), ("F", "23", v23)):
            if isinstance(v, str):
                v = v.format(**rows)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"h_is_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, "Derived" if tag == "Derived" else "Verified")
        H.cell(r, 8, ("稅前利益 − 所得稅" if tag == "Derived" else f"{XBRL}：{tag}；{K10_26}；FY2023 取 {K10_25}"
                      + ("；FY2024–FY2026 引用本頁上方 h_rev" if k == "rev" else "")))
        r += 1
    H.cell(r, 1, "核對：FY2026 營業利益（本節 10-K）− Q 頁四季合計（應為 0）"); H[f"E{r}"] = "=h_is_oi_26-a26_oi"; H[f"E{r}"].number_format = "0.000"; r += 1
    H.cell(r, 1, "核對：FY2026 毛利（本節 10-K）− Q 頁四季合計（應為 0）"); H[f"E{r}"] = "=h_is_gp_26-a26_gp"; H[f"E{r}"].number_format = "0.000"; r += 1
    H.cell(r, 1, "核對：營收 − 營業成本 − 毛利（各年應為 0）")
    for col in "CDEF":
        H[f"{col}{r}"] = f"={col}{rows['rev']}-{col}{rows['cogs']}-{col}{rows['gp']}"; H[f"{col}{r}"].number_format = "0.000"
    r += 1
    # 成品日（版本字串與成品日依規格不列入逐格比對）
    rb = lab(H, "h_built")
    H.cell(rb, 3).value = "=DATE(2026,11,14)"
    return rows


# ---------- PL：標準損益表（v0.3.7） ----------
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJ"


def pl_standard(wb):
    P = wb["PL"]
    rb = lab(P, "情境 B："); g = lambda prefix: lab(P, prefix, rb)
    R = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), oth=g("其他收益淨額"), ni=g("淨利（不含權益證券損益）"),
             sh=g("稀釋股數"), eps=g("EPS（不含權益證券損益）"))
    R["opx"] = lab(P, "營業費用", R["opx"] + 1)       # 「營業費用（基準路徑）」的下一列＝營業費用
    xb = lab(P, "情境 B", lab(P, "三情境延伸到 FY2032"))
    X = dict(rev=lab(P, "營收", xb), ni=lab(P, "淨利（延伸表）", xb), sh=lab(P, "稀釋股數", xb), eps=lab(P, "EPS（延伸表）", xb))
    D = wb["DCF"]; dbr = lab(D, "情境 B"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    Q = wb["Q"]; q_eps_chk = lab(Q, "檢查：EPS（不含權益證券損益）− non-GAAP EPS")
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.3.7 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋未來五年（FY2027–FY2031，情境 B 基準）。US$B；EPS 美元。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, F(T("口徑：Non-GAAP（本模型口徑）。NVIDIA 的財測（毛利率、營業費用）與市場共識 EPS 以 Non-GAAP 為主，故本表用 Non-GAAP：營收、毛利、營業費用與 GAAP 相同，"
                    "稅前與稅後排除權益證券損益（一次性、波動大）；股票薪酬照列為費用，與公司自 FY2027 起的 Non-GAAP 新口徑相近（EPS 每季差 "),
                   'TEXT(ABS(MAX(Q!B' + str(q_eps_chk) + ':I' + str(q_eps_chk) + ')),"0.00")', T("～"),
                   'TEXT(ABS(MIN(Q!B' + str(q_eps_chk) + ':I' + str(q_eps_chk) + ')),"0.00")',
                   T(" 美元，Q 頁）。FY2024–FY2026 與公司當年公告的 Non-GAAP EPS 不同：當年公司口徑排除股票薪酬，本表照列為費用。GAAP 合計列在本表下方對照。FY2024、FY2025 的稅後淨利＝（營業利益＋不含權益證券損益的其他收益）×（1 − GAAP 有效稅率）；FY2026 為 Q 頁四季合計（各季有效稅率）。")))
    r_note = r; r += 1
    P.cell(r, 1, "FY2027E＝錨點（第一、二季實績＋第三季指引＋第四季假設）；FY2028E–FY2030E＝情境 B 由下而上；FY2031E＝延伸期（收斂法只推營收與淨利率，營業利益＝淨利 ÷（1 − 稅率）− 其他收益，同 DCF 頁；不拆營業成本、毛利與營業費用，該三列留白）")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + YEARS + ["說明"]
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
        if note: P.cell(r, 11, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "CDEFG"]       # FY2026A–FY2030E（PL 情境 B 區塊 C–G 欄）
    row("rev", "營收", "US$B", ["=h_is_rev_24", "=h_is_rev_25"] + fc("rev") + [f"=PL!H{X['rev']}"], M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_is_rev_23-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    row("cogs", "營業成本（COGS）", "US$B", ["=h_is_cogs_24", "=h_is_cogs_25"] + [f"=PL!{c}{R['rev']}-PL!{c}{R['gp']}" for c in "CDEFG"] + [None], M)
    row("gp", "毛利", "US$B", ["=h_is_gp_24", "=h_is_gp_25"] + fc("gp") + [None], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS[:7]] + [None], PCT)
    row("opx", "營業費用（研發＋銷管）", "US$B", ["=h_is_opx_24", "=h_is_opx_25"] + fc("opx") + [None], M)
    row("oi", "營業利益", "US$B", ["=h_is_oi_24", "=h_is_oi_25"] + fc("oi") + [f"=DCF!G{d_ebit}"], M, "FY2031E 取 DCF 頁情境 B 的營業利益（延伸期反推）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額，不含權益證券損益）", "US$B", ["=h_is_nonop_24-h_is_eqg_24", "=h_is_nonop_25-h_is_eqg_25"] + fc("oth") + ["=oth_f"], M, "利息收入 − 利息費用 ± 其他；權益證券損益排除")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", ["=C{pti}*h_is_tax_24/h_is_pti_24".format(**L), "=D{pti}*h_is_tax_25/h_is_pti_25".format(**L)]
        + [f"={c}{L['pti']}-{c}{L['pti'] + 2}" for c in COLS[2:]], M, "FY2024、FY2025＝稅前 × GAAP 有效稅率；其餘＝稅前 − 稅後淨利（模型的有效稅率）")
    row("ni", "稅後淨利", "US$B", [f"=C{L['pti']}-C{L['tax']}", f"=D{L['pti']}-D{L['tax']}"] + fc("ni") + [f"=PL!H{X['ni']}"], M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", ["=h_is_dil_24", "=h_is_dil_25"] + fc("sh") + [f"=PL!H{X['sh']}"], M)
    row("eps", "EPS（完全稀釋）", "US$", [f"=C{L['ni']}/C{L['sh']}", f"=D{L['ni']}/D{L['sh']}"] + fc("eps") + [f"=PL!H{X['eps']}"], "0.00")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS[:7]] + [None], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    ni23 = "((h_is_oi_23+h_is_nonop_23-h_is_eqg_23)*(1-h_is_tax_23/h_is_pti_23))"
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/{ni23}-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT,
        "FY2024 的基期＝FY2023 同口徑淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估權益證券損益，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_oth", "GAAP：其他收益淨額（含權益證券損益）", "US$B", ["=h_is_nonop_24", "=h_is_nonop_25", "=h_is_nonop_26"], M, "10-K 年度值（Hist 頁）")
    row("g_pti", "GAAP：稅前利益", "US$B", ["=h_is_pti_24", "=h_is_pti_25", "=h_is_pti_26"], M)
    row("g_tax", "GAAP：所得稅", "US$B", ["=h_is_tax_24", "=h_is_tax_25", "=h_is_tax_26"], M)
    row("g_ni", "GAAP：稅後淨利", "US$B", ["=h_is_ni_24", "=h_is_ni_25", "=h_is_ni_26"], M)
    row("g_eps", "GAAP：EPS（完全稀釋，公司公告）", "US$", ["=h_is_eps_24", "=h_is_eps_25", "=h_is_eps_26"], "0.00")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS[:7]] + [None], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None], "0.000")
    row("c3", "核對：稅後淨利 − 情境 B 摘要（FY2026A–FY2030E）／延伸表（FY2031E）", "US$B",
        [None, None] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip(COLS[2:7], "CDEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}"], "0.000")
    row("c4", "核對：GAAP 稅後淨利 − 本表稅後淨利 − 權益證券損益 ×（1 − GAAP 有效稅率）（FY2024、FY2025）", "US$B",
        [f"=C{L['g_ni']}-C{L['ni']}-h_is_eqg_24*(1-h_is_tax_24/h_is_pti_24)", f"=D{L['g_ni']}-D{L['ni']}-h_is_eqg_25*(1-h_is_tax_25/h_is_pti_25)"], "0.000")
    # v0.3.7 核對後新增（獨立核對必修 1、3）
    r += 1
    P.cell(r, 1, "FY2028 EPS 與共識差距的靜態拆解（情境 B；一次只改一項，營業費用、其他收益、稅率、股數不變；只供展示，不回饋計算）").font = BOLD; L["gap"] = r; r += 1
    eB = lambda rev, gm: f"((({rev})*({gm})-PL!E{R['opx']})+PL!E{R['oth']})*(1-PL!E{R['oth'] + 1})/PL!E{R['sh']}"
    for key, labl, f in (("eps_s", "靜態式重算的 FY2028 EPS（情境 B；應等於模型）", "=" + eB(f"PL!E{R['rev']}", f"PL!E{R['gp']}/PL!E{R['rev']}")),
                         ("eps_rv", "營收改為共識推算（Recon 頁）時的 FY2028 EPS", "=" + eB(f"Recon!B{lab(wb['Recon'], 'FY2028 營收 vs 共識推算')}", f"PL!E{R['gp']}/PL!E{R['rev']}")),
                         ("eps_gm", "毛利率改為公司指引中點時的 FY2028 EPS", "=" + eB(f"PL!E{R['rev']}", "g28_gm"))):
        P.cell(r, 1, labl); P[f"G{r}"] = f; P[f"G{r}"].number_format = "0.00"; L[key] = r; r += 1
    P.cell(r, 1, "核對：靜態式 − 模型 FY2028 EPS（應為 0）"); P[f"G{r}"] = f"=G{L['eps_s']}-PL!E{R['eps']}"; P[f"G{r}"].number_format = "0.000000"; L["gap_chk"] = r; r += 1
    P.cell(r, 1, "營收落差的 EPS 影響（營收改共識推算）"); P[f"G{r}"] = f"=G{L['eps_rv']}/PL!E{R['eps']}-1"; P[f"G{r}"].number_format = "+0.0%;-0.0%"; L["imp_rv"] = r; r += 1
    P.cell(r, 1, "毛利率落差的 EPS 影響（毛利率改指引中點）"); P[f"G{r}"] = f"=G{L['eps_gm']}/PL!E{R['eps']}-1"; P[f"G{r}"].number_format = "+0.0%;-0.0%"; L["imp_gm"] = r; r += 1
    r += 1
    P.cell(r, 1, "差額：由下而上營收合計 − 報告值營收（情境 B；FY2027 為錨點校準殘差：校準係數只以情境 A 校準，FY2028 起報告值即由下而上，差額為 0）")
    for c in "DEFG":
        P[f"{c}{r}"] = f"=PL!{c}{lab(P, '由下而上營收合計', lab(P, '情境 B：'))}-PL!{c}{R['rev']}"; P[f"{c}{r}"].number_format = "0.000"
    L["seg_gap"] = r; r += 1
    for col in "CDEFGHIJ": P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, start=s0)
    return L


# ---------- RevDCF：反向 DCF（v0.3.7） ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7   # DCF 頁：評價日後占比、折現年數（期中）
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.7 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；US$B）", f"=dcf_nc+DCF!$C${fcf}*stub_frac-stub_div", "#,##0.0", "同 DCF 頁每股價值公式")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；US$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${ps}", "0.00")
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
    yr("rev", "營收（US$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（US$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（US$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（US$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（US$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
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
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（美元，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；美元）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), f'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 美元；要支持現價 "), f'TEXT(px,"0.00")', T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或營業利益率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), f'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，低於同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × NVIDIA 過去十年相對同業的溢價 "), f'TEXT({coef},"0.00")', T(" 倍）。換言之，市場目前給 NVIDIA 的倍數只有同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，沒有反映歷史溢價。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; RC = wb["Recon"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, band = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), lab(V, "價格帶（DCF")
    cons_gap = vb("與市場共識目標價的差距：EPS 因素")
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
    rb = lab(P, "情境 B："); gmB = lab(P, "毛利率", lab(P, "B. 報告值", rb)); epsB = lab(P, "EPS（不含權益證券損益）", rb)
    rb0 = lab(P, "情境 B"); gm_s = lab(P, "毛利率", rb0); eps_s = lab(P, "EPS（不含權益證券損益）", rb0)
    k_case = lab(SE, "NVIDIA 顆數係數 低"); asic = lab(SE, "ASIC 金額份額 +5pt"); flat = lab(SE, "CY2029 顆數與 CY2028 持平")
    q3 = "q3_rev"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.7 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("上檔主要來自同業比較法：NVIDIA 本益比若回到過去十年相對同業的溢價（"), f'TEXT({coef},"0.00")', T(" 倍），同業比較法得出 "),
        f'TEXT({cadj},"0")', T(" 美元；只看現金流折現則為 "), f'TEXT({dcf},"0")', T(" 美元，低於現價。不採這項溢價時，加權目標價為 "),
        f'TEXT({noadj},"0")', T(" 美元（"), f'TEXT({noadj}/px-1,"+0%;-0%")', T("），評等為"), REC5(f"({noadj}/px-1)"), T("。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：同業比較法採用 NVIDIA 過去十年相對同業的本益比溢價、FY2027 的營收校準缺口在之後年度持續、HBM 漲價由 NVIDIA 吸收一半。"
                                    "不採溢價時評等為"), REC5(f"({noadj}/px-1)"), T("（"), f'TEXT({noadj}/px-1,"+0.0%;-0.0%")', T("）；現金流折現法權重降到三成時為"),
        f'{REC5(f"((0.3*{dcf}+0.7*{cadj})/px-1)")}', T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), f'TEXT(px,"0.00")', T(" 高 "), f'TEXT({up},"0.0%")',
        T("，評等「"), rec, T("」。上檔全部來自同業比較法：同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍，乘上 NVIDIA 過去十年相對同業的溢價 "),
        f'TEXT({coef},"0.00")', T(" 倍；不加這項溢價，加權目標價只有 "), f'TEXT({noadj},"0.00")', T(" 美元，評等為"), REC5(f"({noadj}/px-1)"), T("。"))
    put("o2", T("現金流折現偏貴：以 "), f'TEXT(wacc,"0.0%")', T(" 折現，現金流折現法目標價 "), f'TEXT({dcf},"0.00")', T(" 美元（三情境加權），三個情境（"),
        f'TEXT(Val!B{band},"0")', T("～"), f'TEXT(Val!C{band},"0")', T(" 美元）都低於現價。反向推算（情境 B），要支持現價需把折現率降到 "),
        f'TEXT({rv("w1")},"0.0%")', T("，或把 FY2028 起的營收全面放大 "), f'TEXT({rv("s1")},"0.00")', T(" 倍。"))
    put("o3", T("盈餘低於共識：情境加權的 FY2028 EPS 比市場共識低 "), f'TEXT(-{cons_gap},"0.0%")', T("。以情境 B 靜態估算，營收若達共識推算的 "),
        f'TEXT(Recon!B{lab(RC, "FY2028 營收 vs 共識推算")},"0")', T(" 十億美元（模型 "), f'TEXT(PL!E{lab(P, "營收", rb0)},"0")', T("），EPS "),
        f'TEXT(PL!G{PLA["imp_rv"]},"+0.0%;-0.0%")', T("；毛利率若從 "), f'TEXT(PL!E{gm_s},"0.0%")', T(" 升到公司指引中點 "), 'TEXT(g28_gm,"0.0%")',
        T("，EPS "), f'TEXT(PL!G{PLA["imp_gm"]},"+0.0%;-0.0%")', T("。營收落差的影響較大。"))
    put("o4", T("核心問題：顆數持平後 EPS 還能成長，但幅度取決於顆數。情境 B 的顆數在 2028 年先減，FY2029 EPS "), f'TEXT(PL!F{eps_s},"0.00")',
        T(" 美元，比 FY2028 "), f'TEXT(PL!F{eps_s}/PL!E{eps_s}-1,"+0.0%;-0.0%")', T("；2029 年顆數回升後，FY2030 為 "), f'TEXT(PL!G{eps_s},"0.00")',
        T(" 美元（"), f'TEXT(PL!G{eps_s}/PL!F{eps_s}-1,"+0.0%;-0.0%")', T("）。若 2029 年顆數與 2028 年持平（以情境 A 為底），FY2030 EPS 只成長 "),
        f'TEXT(Sens!N{flat},"+0.0%;-0.0%")', T("。"))
    put("o5", T("最敏感的是出貨顆數：顆數少一成時 FY2028 EPS "), f'TEXT(Sens!E{k_case},"+0.0%;-0.0%")', T("；ASIC 金額份額多 5 個百分點時 "),
        f'TEXT(Sens!E{asic},"+0.0%;-0.0%")', T("。下一個檢驗點是第三季（FY2027）財報：營收對照公司指引 "), f'TEXT({q3}*10,"#,##0")', T(" 億美元。"))
    put("drv", T("分部營收推導（情境 B）：資料中心運算營收＝各世代 GPU 認列顆數 × 每顆售價（GPU 模組、CPU、LPDDR 與機櫃加價中 NVIDIA 保留的部分）的加總；"
                 "網通營收＝網通市場模型中歸屬 NVIDIA 的金額，隨 NVIDIA 顆數同比例縮放；Edge Computing（遊戲、PC、工作站、車用）＝FY2027 實績錨點 × 固定年增率。"
                 "三者相加為合併營收。FY2028 運算 "), f'TEXT(PL!E{lab(P, "運算營收（Σ", rb)},"0")', T("、網通 "), f'TEXT(PL!E{lab(P, "NVIDIA 網通營收", rb)},"0")',
        T("、Edge "), f'TEXT(PL!E{lab(P, "Edge Computing 營收", rb)},"0")', T("，合計 "), f'TEXT(PL!E{lab(P, "由下而上營收合計", rb)},"0")', T(" 十億美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：只改展示層——個股 HTML 版面規格 v1.3 試作：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁（反向 DCF：現價隱含營業利益率、營收規模、WACC 與 Comparable 倍數）、"
                  "PL 頁末標準損益表（FY2024–FY2026 實績＋FY2027–FY2031 情境 B，Non-GAAP）、Hist 頁末損益表歷史（SEC XBRL）、README 頁末 v1.3 白話列；既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def build_id(here):
    h = hashlib.sha256(VER.encode())
    for fn in sorted(["build_nvidia.py", "v035_post.py", "v037_post.py", "px1105_post.py", "tp_v2.py", "accel_share.py", "sens.json", "segments.json", "hist_frozen.json"]):
        p = os.path.join(here, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    return f"nvidia-v{VER}-{h.hexdigest()[:10]}"


def version(wb, here):
    RD = wb["README"]; SE = wb["Sens"]
    old = SE.cell(lab(SE, "build_id"), 4).value
    new = build_id(here)
    n = 0
    for ws in (RD, SE):
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and old in c.value:
                    c.value = c.value.replace(old, new); n += 1
    a1 = RD["A1"].value
    assert "個股模型 v0.3.6" in a1, a1[:80]
    RD["A1"].value = a1.replace("個股模型 v0.3.6", f"個股模型 v{VER}").replace(
        "FY↔CY 月數加權）", f"FY↔CY 月數加權）；v{VER}：只改展示層（個股 HTML 版面規格 v1.3：標準損益表、反向 DCF 頁）", 1)
    return old, new, n


def apply(wb, here):
    HR = hist_is(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    RMA = readme13(wb, RVA, PLA)
    old, new, n = version(wb, here)
    wb._v037 = dict(PL=PLA, RV=RVA, RM=RMA)
    return f"v037_post：Hist 損益表歷史 {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；README v1.3 白話列 {len(RMA)} 列；build_id {old} → {new}（{n} 格）"
