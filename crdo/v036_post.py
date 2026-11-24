# Credo v0.3.6（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0 公司實績對帳、硬體股兩讀數）。
# 依據：docs/workorders/20261114_第二批13家_v1.3與三層協調.md、docs/workorders/20261114_C0實績對帳補列_七家.md、
#       docs/plan/三層基準協調規則.md v1.4.2（C0 第 2 點、C4）、docs/plan/個股HTML版面規格.md v1.3、tools/coord/README.md。
# 樣板：Broadcom v0.5.6／v0.5.7（avgo/v056_post.py、v057_post.py；同一建檔家族）、AMD v0.4.7、Marvell v0.3.5（mrvl/v035_post.py）。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史（FY2024 non-GAAP，Credo 8-K 新聞稿；FY2025–FY2026 連 Q 頁；GAAP 對照三年）
#   PL 頁末：標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B；Non-GAAP）、比率、GAAP 對照、核對列
#   RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（營運輸入＝整本重算快照 crdo/sens.json）
#   Coord 頁末 F–J 節；具名範圍 SIF_RecoFlag_<年>（指向 Coord 頁對帳中旗標列；FY(t)≈CY(t−1)，供版面檢查依會計年度對應）
#   README 頁末：v1.3 展示層白話列；A1 版號（tools/coord stamp）
# build_crdo.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL, CR)。COORD_RUN（alt_run.py）或 CRDO_SENS（sens_crdo.py）時不執行。
import json, os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, OLD, BUILT = "0.3.6", "0.3.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
E8K = "https://www.sec.gov/Archives/edgar/data/1807794/"
SRC = {"FY2025": E8K + "000162828025028827/credoq42025ex-991.htm（Credo 8-K 2025-06-02 EX-99.1：Q4 與 FY2025 新聞稿；FY2024 為比較欄）",
       "FY2024": E8K + "000162828024025607/credoq42024ex-991.htm（Credo 8-K 2024-05-29 EX-99.1：Q4 與 FY2024 新聞稿；FY2023 為比較欄，只用營收與 non-GAAP 淨利作年增率基期）",
       "FY2026": E8K + "000162828026039474/credoq42026ex-991.htm（Credo 8-K 2026-06-01 EX-99.1：Q4 與 FY2026 新聞稿）"}
SRC_Q = {"Q4": E8K + "000162828026039474/credoq42026ex-991.htm（8-K 2026-06-01 EX-99.1：Q4 FY26 營收 437,003 千美元）",
         "Q1": E8K + "000162828026059795/credoq12027ex-991.htm（8-K 2026-09-01 EX-99.1：Q1 FY27 營收 479.0 百萬美元）"}


def lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def lab(ws, prefix, start=1, col=1, exact=False):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- Hist：損益表歷史（新頁） ----------
# (鍵, 項目, FY2024, FY2025, FY2026（連 Q 頁）, FY2023（基期）, 標記)；十億美元、EPS 美元、股數十億股
HIS = [
    ("rev", "營收（GAAP＝non-GAAP）", 0.19297, "=Q!B4", "=Q!C4", 0.184194, "Verified"),
    ("gp", "毛利（non-GAAP）", 0.120562, "=Q!B5", "=Q!C5", None, "Verified"),
    ("opx", "營業費用（non-GAAP）", 0.117833, "=Q!B7", "=Q!C7", None, "Verified"),
    ("oi", "營業利益（non-GAAP）", 0.002729, "=Q!B9", "=Q!C9", None, "Verified"),
    ("oth", "其他收益（淨）", 0.014313, "=Q!B10", "=Q!C10", None, "Verified"),
    ("pti", "稅前利益（non-GAAP）＝營業利益＋其他收益", "=C{oi}+C{oth}", "=Q!B11", "=Q!C11", None, "Derived"),
    ("tax", "所得稅（non-GAAP）＝稅前利益 − 淨利", "=C{pti}-C{ni}", "=Q!B12", "=Q!C12", None, "Derived"),
    ("ni", "稅後淨利（non-GAAP）", 0.01457, "=Q!B14", "=Q!C14", 0.008687, "Verified"),
    ("sh", "non-GAAP 稀釋加權股數（十億股）", 0.170143, "=Q!B16", "=Q!C16", None, "Verified"),
    ("eps", "EPS（non-GAAP 稀釋，公司公告；美元）", 0.09, "=a25_eps_in", "=a26_eps_in", None, "Verified"),
    ("g_ni", "稅後淨利（GAAP；淨損為負）", -0.028369, 0.052183, 0.472279, None, "Verified"),
    ("g_eps", "EPS（GAAP 稀釋，公司公告；美元）", -0.18, 0.29, 2.51, None, "Verified"),
]


def hist_pl(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = (f"Hist — 損益表歷史（v{VER} 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。十億美元；EPS 美元；股數十億股。會計年度 FY(t) 於 t 年 4 月底／5 月初結束。"
               "non-GAAP 為公司口徑（排除股票薪酬、資產減損與相關稅務影響）；FY2025、FY2026 欄連 Q 頁（與公司全年值一致），GAAP 三年取自新聞稿損益表")
    H["A1"].font = BOLD
    r0 = 1; r = 3
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "FY2023（只供年增率基期）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v24, v25, v26, v23, tag in HIS:
        H.cell(r, 1, f"hp_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "24", v24), ("D", "25", v25), ("E", "26", v26), ("F", "23", v23)):
            if v is None:
                continue
            if isinstance(v, str) and "{" in v:
                v = v.format(**rows)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"hp_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, tag)
        H.cell(r, 8, f"FY2024：{SRC['FY2024']}；FY2025：" + ("Q 頁 B 欄（" + SRC["FY2025"] + "）" if isinstance(v25, str) else SRC["FY2025"]) + "；FY2026：" + ("Q 頁 C 欄（" + SRC["FY2026"] + "）" if isinstance(v26, str) else SRC["FY2026"]))
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    for labl, f in (("稅前利益 − 所得稅 − 稅後淨利（non-GAAP）", "={c}{pti}-{c}{tax}-{c}{ni}"),
                    ("稅後淨利 ÷ 稀釋股數 − 公司公告 EPS（四捨五入差）", "={c}{ni}/{c}{sh}-{c}{eps}")):
        H.cell(r, 1, labl)
        for c in "CDE":
            H[f"{c}{r}"] = f.format(c=c, **rows); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 1, "備註：FY2024 non-GAAP 所得稅＝GAAP 所得稅 5.624 − 調節表其他稅務影響 3.152＝2.472（百萬美元），與稅前 − 淨利一致。FY2024 GAAP 淨損 −28.4 百萬美元（含股票薪酬 39.0）。").font = Font(italic=True)
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 58
    for c in "CDEF": H.column_dimensions[c].width = 14
    H.column_dimensions["H"].width = 90
    return dict(start=r0, rows=rows)


# ---------- PL：標準損益表 ----------
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJK"   # 本表欄；C–E＝Hist（FY2024A–FY2026A）；F–I＝PL 情境 B 區塊 D–G（FY2027 錨點、FY2028–FY2030）；J–K＝PL 延伸表情境 B 的 H–I（FY2031–FY2032）


def pl_standard(wb):
    P = wb["PL"]
    Q = wb["Q"]
    for rr, t in ((4, "營收"), (5, "毛利（non-GAAP）"), (7, "營業費用（non-GAAP）"), (9, "營業利益（non-GAAP）"), (10, "其他收益（淨）"), (11, "稅前利益"),
                  (12, "所得稅（non-GAAP）"), (14, "淨利（non-GAAP）"), (16, "稀釋股數（non-GAAP）")):
        assert Q.cell(rr, 1).value == t, (rr, Q.cell(rr, 1).value)
    assert Q.cell(3, 2).value == "FY2025A" and Q.cell(3, 3).value == "FY2026A"
    rb = lab(P, "情境 B：")
    hb = rb + 1
    assert [P.cell(hb, c).value for c in (3, 4, 5, 8, 10)] == ["FY2025A", "FY2026A", "FY2027（錨點）", "FY2030（收斂）", "FY2032（收斂）"], [P.cell(hb, c).value for c in range(3, 11)]
    rc = lab(P, "D. 損益（non-GAAP）", rb)
    g = lambda lbl, s=rb: lab(P, lbl, s, exact=True)
    R = dict(rev=lab(P, "營收", lab(P, "C. 其他營收與合計", rb), exact=True), gp=g("毛利", rc), opx=g("營業費用", rc), oi=g("營業利益", rc), oth=g("其他收益（淨）", rc),
             tx=g("有效稅率", rc), ni=g("淨利", rc), sh=g("稀釋股數", rc), eps=g("EPS（non-GAAP）", rc))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋未來六年（FY2027E 錨點、FY2028E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "會計年度 FY(t) 於 t 年 4 月底／5 月初結束。只引用本頁上方與 Hist、Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。Credo 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股票薪酬與資產減損），故本表用 Non-GAAP；其他收入（淨額）＝其他收益（淨）。GAAP 淨利與 EPS 列在本表下方對照（只列實績）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2027E＝錨點（1Q27 實績＋2Q27 指引＋2H27 由全年說法推得）；FY2028E–FY2029E＝情境 B 由下而上；FY2030E–FY2032E＝收斂法（只推營收與淨利，營業利益＝淨利 ÷（1 − FY2029E 稅率）− FY2029E 其他收益，同 DCF 頁）。"
                 "收斂期毛利率沿用 FY2029E，據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列（目標價用收斂期淨利，不受影響）。EPS＝non-GAAP 稀釋")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r_hdr = r; r += 1
    L = {}
    NC = 3 + len(YEARS)

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, NC, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "EFG"]            # FY2027E–FY2029E（PL 情境 B 區塊 E–G）
    ex = lambda key: [f"=PL!{c}{R[key]}" for c in "HIJ"]            # FY2030E–FY2032E（同區塊 H–J，收斂法）
    hh = lambda k: [f"=hp_{k}_24", f"=hp_{k}_25", f"=hp_{k}_26"]
    LAST = "H"   # 本表 FY2029E 欄
    EXT = "IJK"
    row("rev", "營收", "US$B", hh("rev") + fc("rev") + ex("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/hp_rev_23-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "US$B", [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "CDEFGH"] + [f"={c}{L['rev']}*(1-${LAST}{r_c + 2})" for c in EXT], M,
        "FY2024A–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp") + [f"={c}{L['rev']}-{c}{L['cogs']}" for c in EXT], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx") + [None] * 3, M,
        "FY2030E–FY2032E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）（non-GAAP，不含股票薪酬）")
    r_ni = r_c + 9
    row("oi", "營業利益", "US$B", hh("oi") + fc("oi") + [f"={c}{r_ni}/(1-PL!$G${R['tx']})-PL!$G${R['oth']}" for c in EXT], M, "FY2030E–FY2032E 由淨利反推（同 DCF 頁；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝其他收益（淨）", "US$B", hh("oth") + fc("oth") + [f"=PL!$G${R['oth']}"] * 3, M, "non-GAAP；利息收入為主")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", hh("tax") + [f"={c}{L['pti']}-{c}{r_ni}" for c in "FGHIJK"], M,
        "歷史＝non-GAAP 所得稅；預測＝稅前 − 稅後淨利（FY2028 起 OECD 全球最低稅，模型有效稅率 13%–14%）")
    assert L["tax"] == r_c + 8 and r == r_ni
    row("ni", "稅後淨利", "US$B", hh("ni") + fc("ni") + ex("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", hh("sh") + fc("sh") + ex("sh"), M)
    row("eps", "EPS（non-GAAP 稀釋）", "US$", hh("eps") + fc("eps") + ex("eps"), "0.00",
        "FY2024A–FY2026A＝公司公告 non-GAAP 稀釋 EPS；預測＝淨利 ÷ 稀釋股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/hp_ni_23-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 的基期＝FY2023 non-GAAP 淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_ni", "GAAP：稅後淨利（淨損為負）", "US$B", ["=hp_g_ni_24", "=hp_g_ni_25", "=hp_g_ni_26"], M)
    row("g_eps", "GAAP：EPS（稀釋，公司公告）", "US$", ["=hp_g_eps_24", "=hp_g_eps_25", "=hp_g_eps_26"], "0.00", "含股票薪酬與資產減損")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2024A–FY2029E）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:6]] + [None] * 3, "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E–FY2032E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "IJK":
        P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"H{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E–FY2032E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "US$B", [None] * 6 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "IJK"], "0.000")
    row("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E–FY2032E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 5 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "HIJK"], "0.00%")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 區塊（FY2025A–FY2032E）", "US$B",
        [None, f"=D{L['ni']}-PL!C{R['ni']}", f"=E{L['ni']}-PL!D{R['ni']}"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJK", "EFGHIJ")], "0.000")
    row("c5", "核對：EPS − 情境 B 區塊（FY2027E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHIJK", "EFGHIJ")], "0.000")
    P.column_dimensions["A"].width = max(P.column_dimensions["A"].width or 0, 62)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = lab(D, "評價日後占該年度比例"), lab(D, "折現年數（期中")
    tvr = lab(D, "終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（v{VER} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。明確期第一年 FY2027（錨點；評價日後只含 2027-01～04）維持模型值不動，只動 FY2028–FY2032（含終值）。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("nc", "淨現金滾到評價日（情境 B；US$B）", f"={ncr}", "#,##0.0", "Val 頁：淨現金與投資＋stub 期 FCFF − 已宣告股利（與 WACC 無關）")
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
    yr("ebit", "營業利益 EBIT（GAAP 口徑：non-GAAP 營業利益 × 調整係數；US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元 EBIT 對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（US$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（US$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期 EBIT 率（FY2028–FY2032 統一一個 EBIT 率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效 EBIT 率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型 EBIT 率")
    kv("m1", "現價隱含的 EBIT 率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；EBIT 率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
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
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tvr}+$B${A['nc']})/sh_v")
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
    kv("w_ok", "判讀（隱含 WACC）", f'=IF(ABS(B{A["w_res"]})<0.01,"已收斂","搜尋下限仍不足：折現率降到下限（永續成長率＋0.5 個百分點）時每股價值仍低於支持現價所需值")', "@")
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
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 美元；要支持現價 "), 'TEXT(px,"0.00")', T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要："), f'IF(ABS(B{A["w_res"]})<0.01,"折現率降到 "&TEXT(B{A["w1"]},"0.0%"),"折現率即使降到 "&TEXT(B{A["w1"]},"0.0%")&"（搜尋下限）仍不足")', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或 EBIT 率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 相對倍數係數 "), f'TEXT({coef},"0.00")', T("；Credo 上市未滿十年，資料不足，不調整）。換言之，市場目前給 Credo 的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("模型採用的係數。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFGH": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：估值並列四個目標價（版面規格 v1.3 G 節；只新增列） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）")
    cons = vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r_px = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, f"0.7 估值並列：四個目標價（v{VER} 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "美元", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 美元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{r_px}"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px 列）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（版面規格 v1.3 H 節） ----------
def sens_tor(wb, here):
    SE, V = wb["Sens"], wb["Val"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")
    P = wb["PL"]
    eps_s = {k: lab(P, "EPS（non-GAAP）", lab(P, f"情境 {k}", exact=True), exact=True) for k in "ABC"}   # PL 頁頂端摘要；Val 前瞻 EPS＝FY2028、FY2029 各半（同 Val 頁公式）
    fwd = lambda k: f"(PL!F{eps_s[k]}+PL!G{eps_s[k]})/2"
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    par = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        par.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；v{VER} 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（crdo/sens_crdo.py：三情境同時套用案例覆寫，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用本頁「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    frame = [
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fwd('C')}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fwd('A')}",
         "Val 頁單一情境 DCF 價值；前瞻 EPS＝PL 頁該情境 FY2028、FY2029 EPS 各半（同 Val 頁）"),
        ("DCF 在加權中的權重 100%（只看現金流折現）／0%（只看同業比較法）（基準＝50%）",
         f"={dcf}", f"={cadj}", "Val 頁：DCF 與 Comparable（調整後）目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    for k, d in J["biz"].items():
        SE.cell(r, 1, d["label"])
        for side, col, tcol in (("低", 2, 4), ("高", 3, 5)):
            if d.get(side) is None:
                continue
            SE.cell(r, tcol, d[side]["tp"]).number_format = "#,##0.00"; SE.cell(r, tcol).font = BLUE
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（crdo/sens.json）；覆寫同本頁上方「一次動一個輸入」案例，三情境同時套用"); r += 1
    for sr in par:
        t = str(SE.cell(sr, 1).value)
        if t.startswith("DCF 權重"):
            continue
        SE.cell(r, 1, t)
        for col, tcol, scol in ((2, 4, "C"), (3, 5, "E")):
            v = SE[f"{scol}{sr}"].value
            if v is None or (isinstance(v, str) and not v.startswith("=")):
                continue
            SE.cell(r, tcol, f"=Sens!{scol}{sr}").number_format = "#,##0.00"
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「目標價參數敏感度」表）"); r += 1
    last = r - 1
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f'=IF(ABS({tp}-{J["_base"]["tp"]!r})<=1E-6*ABS({tp}),"CURRENT","STALE")')
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 crdo/sens_crdo.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：讀取層級、讀取層級核對、層級之間、損益表年度標示、C0（實績） ----------
def coord(wb, SIL):
    C, I, Q, P = wb["Coord"], wb["In"], wb["Q"], wb["PL"]
    N = SIL.NAMES
    r = C.max_row + 2
    ys = (2026, 2027, 2028, 2029)
    r_flag = lab(C, "對帳中旗標")
    for j, y in enumerate(ys):   # FY(t)＝8/12 × CY(t−1)＋4/12 × CY(t)：版面檢查依具名旗標以會計年度對應（同 NVIDIA、Marvell）；值連 Coord 頁旗標列
        dn(wb, f"SIF_RecoFlag_{y}", f"Coord!${'BCDE'[j]}${r_flag}")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    C.cell(r, 1, "F. 讀取層級（工作單 20261114 第二批第 2 節）：本模型讀哪一層、哪一版、哪些名稱。自層級讀入的同源值只列在此節與 G 節，不進 C0 判定").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("AI 銅纜市場 mcu_*、其中 scale-out 銅纜 mso_*（CY2026–2029）", "市場層 markets/network（加速器顆數讀 Siliconomics）", src_in("mcu_26").split("｜")[0],
         "AEC 營收＝銅纜市場（FY 加權）× Credo 有效份額", "In 頁 A 節"),
        ("AI 可插拔光模組市場 mpl_*（三情境）", "市場層 markets/network", src_in("mpl_26").split("｜")[0], "光學營收＝光模組市場 × Credo 有效份額", "In 頁 A 節、網通 CPO 三情境節"),
        ("情境顆數係數 um28／um29（情境 B 0.80／0.710）", "Siliconomics（經台積電成品）", src_in("um28_B").split("｜")[0], "情境 B、C 的市場量", "In 頁 A 節"),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中"),
        ("Wattonomics 錨定係數 wta_27–wta_29（舊敏感度案例；SIF_WTAnchor_Base）", "Wattonomics（經 Siliconomics 轉入）", vtxt, "Sens 頁「Wattonomics 錨定」案例", "舊口徑，保留供對照；另一側讀數取代其角色"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。Credo 的 AEC、光學有效份額由本模型以 FY2027 錨點校準（＝錨點營收 ÷ 層級市場），層級（markets/network、Siliconomics）沒有 Credo 份額或營收歸屬；"
                 "因此「層級市場 × 本模型有效份額」與 FY2027 錨點恆等（同源校準），見 J 節。").font = Font(italic=True); r += 2
    C.cell(r, 1, "G. 讀取層級核對（同源，不判定）：個股讀入的層級值 vs 層級值，必然相同；只證明本模型讀到現行版，不計入對帳中").font = BOLD; r += 1
    for j, h in enumerate(["項目", "本模型值", "層級對照值", "差距", "狀態", "單位與口徑", "對照來源（層、版本）", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    for k, labl, unit, fmt in (("um28_B", "情境 B：CY2028 顆數 ÷ 原始顆數", "x", "0.000"), ("um29_B", "情境 B：CY2029 顆數 ÷ 原始顆數", "x", "0.000")):
        C.cell(r, 1, labl); C.cell(r, 2, f"={k}").number_format = fmt; C.cell(r, 3, SIL.base(k)).number_format = fmt; C.cell(r, 3).font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"; C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, unit); C.cell(r, 7, f"Siliconomics {SIL.VERSION}（SIF_{k}，經台積電成品）"); C.cell(r, 8, SIL.V[k]["tag"] if k in SIL.V else "")
        C.cell(r, 9, "本模型直接讀層級值；恆等，差距應為 0"); r += 1
    G1 = r - 1
    r += 1
    C.cell(r, 1, "H. 層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：產業部署 GW，Siliconomics 現行基準 vs 另一側讀數（Wattonomics 換算）。這是兩層之間的差距，不是 Credo 自己的量；本模型依 C4 並列兩讀數目標價").font = BOLD; r += 1
    for j, h in enumerate(["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    L0 = r
    for j, y in enumerate(ys):
        C.cell(r, 1, f"產業部署 GW CY{y}")
        C.cell(r, 2, N["SIF_DeployGW_Base"][j]).number_format = "#,##0.00"; C.cell(r, 3, N[f"SIF_DeployGW_Alt_{y}"]).number_format = "#,##0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "層間參考（Siliconomics 標對帳中；本公司不計）" if SIL.RECO_FLAG[y] else "層間參考（兩層一致）")
        C.cell(r, 6, "GW；SIF_DeployGW_Base vs SIF_DeployGW_Alt_<年>")
        C.cell(r, 7, f"Siliconomics {SIL.VERSION}（旗標 SIF_RecoFlag_{y}＝{int(SIL.RECO_FLAG[y])}）；Siliconomics 對帳單 r3")
        C.cell(r, 8, "Derived（基準）／Wattonomics 換算（另一側）"); C.cell(r, 9, "兩層之間的差距，不裁決；本模型 Coord 頁 A–E 節並列兩讀數目標價")
        r += 1
    L1 = r - 1
    r += 1
    C.cell(r, 1, "I. 對帳中年份與損益表年度標示（C4；Credo FY(t)＝8/12 × CY(t−1)＋4/12 × CY(t)，FY2027 為錨點、另一側讀數不縮放；FY2028 起含對帳中年份 CY2027；FY2030E–FY2032E 為收斂期，依賴 CY2029）").font = BOLD; r += 1
    FYS_ = list(range(2027, 2033))
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in FYS_], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年（主要）")
    dep = [2026, 2027, 2028, 2029, 2029, 2029]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2027E–FY2032E）")
    for j, fy in enumerate(FYS_):
        C.cell(r, 2 + j, f'=IF(${"BCDE"[dep[j] - 2026]}${r_flag}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    # ---- J：C0（實績） ----
    C.cell(r, 1, "J. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）").font = BOLD; j_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝Credo 8-K 已公布的營收（Verified；公司不揭露 AEC、光學等產品線營收）。層級側＝層級市場（markets/network AI 銅纜、可插拔光模組，FY 加權）× Credo 有效份額＋其他 IC 與 IP，"
                 "即 PL 情境 B 區塊 FY2027 欄營收。同源校準：Credo 有效份額＝FY2027 錨點營收（1Q27 實績＋2Q27 指引＋2H27 由公司全年說法推得）÷ 層級市場，"
                 "層級市場在比式中消去，層級隱含恆等於本模型錨點（下表核對列差 0）；這張表實際只檢驗「公司實績 vs 本模型 FY2027 錨點（公司說法推得）」，不是層級值。"
                 "故判定標「參考（同源校準）」，結論「無可比實績（層級值同源）」；缺的獨立揭露：AEC 營收（或非本模型校準的 Credo 份額）。"
                 "時差：FY2027＝2026-05～2027-04，層級市場已按月數加權為同一會計年度；已公布 FY2027 期間只有 1Q27（2026-05～08-01，3 個月；2Q27 財報約 2026-12 初）。"
                 "主讀數＝1Q27 實績 vs 層級 FY2027 × 3 ÷ 12；第二讀數＝1Q27 × 4 vs 層級 FY2027（約等於 2026-06 中旬的年化速率，早於會計年度中點 2026-11，成長年偏低）。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); j_note = r; r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級", "網址／名稱"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000", url=""):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); C.cell(r, 7, url); IN[key] = r; r += 1
    assert str(Q.cell(3, 4).value).startswith("1Q27"), Q.cell(3, 4).value
    rbB = lab(P, "情境 B：")
    rmc, rsa, raec = lab(P, "AI 銅纜市場（FY 加權）", rbB), lab(P, "AEC 有效份額", rbB, exact=True), lab(P, "AEC 營收", rbB, exact=True)
    rmp, rso, ropt = lab(P, "AI 可插拔光模組市場（FY 加權）", rbB), lab(P, "Credo 光學有效份額", rbB), lab(P, "光學營收", rbB, exact=True)
    roth, rrev = lab(P, "其他 IC 與 IP", rbB), lab(P, "營收", lab(P, "C. 其他營收與合計", rbB), exact=True)
    rso2 = lab(P, "光學有效份額相對值", rbB)
    inp("q4", "參考：營收 4Q26（公司 8-K；2026-02～05-02，不在 FY2027）", 0.437003, "US$B", "2026-02～2026-05-02（4Q26）", "Credo 8-K 2026-06-01 新聞稿（Q4 FY26）", "Verified", BLUE, url=SRC_Q["Q4"])
    inp("q1", "營收：1Q27（公司 8-K）", f"=Q!D4", "US$B", "2026-05-03～2026-08-01（1Q27）", "Q 頁 D4（In 頁 q1_rev）；Credo 8-K 2026-09-01 新聞稿（Q1 FY27）", "Verified", GREEN, url=SRC_Q["Q1"])
    inp("mo", "已公布月數（FY2027 1Q27）", 3, "月", "2026-05～2026-07", "FY2027 起於 2026-05；已公布一季＝3 個月", "Derived", BLUE, fmt="0")
    inp("mc", "層級：AI 銅纜市場 FY2027（FY 加權）", f"=PL!E{rmc}", "US$B", "FY2027（8/12 CY2026＋4/12 CY2027）", f"PL 情境 B 區塊 E{rmc}（markets/network）", "Derived（層）", GREEN, url=f"PL!E{rmc}")
    inp("sa", "本模型：AEC 有效份額 FY2027（＝錨點 AEC ÷ 銅纜市場；同源校準）", f"=PL!E{rsa}", "%", "FY2027", f"PL 情境 B 區塊 E{rsa}", "Derived（本模型校準）", GREEN, fmt="0.00%", url=f"PL!E{rsa}")
    inp("mp", "層級：AI 可插拔光模組市場 FY2027（FY 加權）", f"=PL!E{rmp}", "US$B", "FY2027", f"PL 情境 B 區塊 E{rmp}（markets/network）", "Derived（層）", GREEN, url=f"PL!E{rmp}")
    inp("so", "本模型：光學有效份額 FY2027（＝錨點光學 ÷ 光模組市場；同源校準）", f"=PL!E{lab(P, 'Credo 光學有效份額（FY2027 校準）', rbB)}", "%", "FY2027", "PL 情境 B 區塊", "Derived（本模型校準）", GREEN, fmt="0.00%")
    inp("ot", "本模型：其他 IC 與 IP FY2027", f"=PL!E{roth}", "US$B", "FY2027", f"PL 情境 B 區塊 E{roth}（FY2026 營收 × 7% × 成長，Assumed）", "Assumed", GREEN)
    inp("L", "層級隱含：Credo 營收 FY2027＝銅纜市場 × AEC 份額＋光模組市場 × 光學份額＋其他", f"=B{IN['mc']}*B{IN['sa']}+B{IN['mp']}*B{IN['so']}+B{IN['ot']}", "US$B", "FY2027", "上列相乘相加", "Derived（同源）")
    inp("chk", "核對：層級隱含 − 本模型 FY2027 錨點營收（應為 0；同源校準恆等式）", f"=B{IN['L']}-Q!G4", "US$B", "FY2027", "Q 頁 G4（錨點）", "Derived", fmt="0.000000")
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    J0 = r
    b = IN; pct = "+0.0%;-0.0%"
    C.cell(r, 1, "主讀數：Credo 營收 1Q27 實績 vs 層級 FY2027 × 已公布月數 ÷ 12")
    C.cell(r, 2, f"=B{b['q1']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "參考（同源校準）")
    C.cell(r, 6, "US$B；層級＝FY2027 全年 × 3 ÷ 12（線性時差，成長年中前段季度偏低）")
    C.cell(r, 7, "公司：Credo 8-K；層級：markets/network × 本模型校準份額"); C.cell(r, 8, "Verified（公司）／Derived（同源）")
    C.cell(r, 9, "層級市場消去：實際比較 1Q27 實績 vs 本模型 FY2027 錨點的線性分攤；錨點含公司說的下半年加速（2H27 每季平均比 2Q27 高約 38%），前段季度低於平均屬預期")
    rm = r; r += 1
    C.cell(r, 1, "第二讀數：Credo 營收 1Q27 × 4 vs 層級 FY2027")
    C.cell(r, 2, f"=B{b['q1']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "參考（同源校準）")
    C.cell(r, 6, "US$B；1Q27 × 4（年化，約等於 2026-06 中旬的速率）；層級＝FY2027 全年")
    C.cell(r, 7, "公司：Credo 8-K（Q 頁 D4）；層級同上"); C.cell(r, 8, "Verified（公司）／Derived（同源）")
    C.cell(r, 9, "季中點早於會計年度中點，成長年年化值偏低；同源，不判定")
    rs = r; r += 1
    C.cell(r, 1, "C0（實績）判定：Credo 營收 FY2027").font = BOLD
    C.cell(r, 5, "無可比實績（層級值同源）").font = BOLD
    C.cell(r, 6, "缺：AEC 營收揭露、或非本模型校準的 Credo 份額（層級只有市場量，沒有 Credo 歸屬）")
    C.cell(r, 7, "J 節核對列：層級隱含 − 錨點＝0")
    C.cell(r, 8, "—")
    C.cell(r, 9, f'="兩讀數差距 "&TEXT(D{rm},"+0%;-0%")&"／"&TEXT(D{rs},"+0%;-0%")&" 只反映錨點的季節與成長形狀（同源），不開對帳單；2Q27 財報（2026-12）公布後可改以兩季實績對錨點檢查，仍屬同源"')
    rv = r; r += 1
    J1 = r - 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（J 節 C0（實績）判定；同源列與讀取層級列不計）"); C.cell(r, 2, f'=IF(E{rv}="對帳中",1,0)'); cnt = r; r += 1
    C.cell(r, 1, "說明：J 節為本公司唯一的 C0 判定，結論「無可比實績（層級值同源）」：Credo 份額由本模型以 FY2027 錨點校準，層級市場消去，表中兩讀數只是實績對本模型錨點的時差形狀。"
                 "產業部署 GW 為層與層之間的差距（H 節），本模型並列兩讀數目標價，不裁決，也不計入對帳中項目數。")
    note = r; r += 1
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, G0=G0, G1=G1, L0=L0, L1=L1, H_PL=H_PL, J=j_title, Jn=j_note, J0=J0, J1=J1, rv=rv, rm=rm, rs=rs, cnt=cnt, note=note)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed = vb("同業前瞻 P/E 中位數")
    rb = lab(P, "情境 B：")
    pr = lambda lbl, s=rb: lab(P, lbl, s, exact=True)
    aec_b, opt_b, rev_b = pr("AEC 營收"), pr("光學營收"), pr("營收", lab(P, "C. 其他營收與合計", rb))
    oth_b = lab(P, "其他 IC 與 IP", rb)
    rr28, reps28, raec, rnm = lab(RC, "FY2028 營收 vs 共識"), lab(RC, "FY2028 EPS vs 共識"), lab(RC, "FY2028 AEC 營收 vs Jefferies"), lab(RC, "FY2027 淨利率 vs 公司目標")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, SEN['first'])}"
    w30 = f"((0.3*{dcf}+0.7*{cadj})/px-1)"
    w70 = f"((0.7*{dcf}+0.3*{cadj})/px-1)"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；v{VER} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元遠低於現價 "), 'TEXT(px,"0")', T(" 美元，同業比較法 "), f'TEXT({cadj},"0")',
        T(" 美元高於現價；兩者各半加權 "), f'TEXT({tp},"0")', T(" 美元，評等「"), rec, T("」，結論取決於兩法的權重。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：現金流折現與同業比較法各占一半；AEC 與光學營收跟著層級的銅纜與光模組市場成長，份額以 FY2027 錨點校準。現金流折現權重提高到七成時評等為"),
        REC5(w70), T("（"), f'TEXT({w70},"+0.0%;-0.0%")', T("），降到三成時為"), REC5(w30), T("（"), f'TEXT({w30},"+0.0%;-0.0%")', T("）。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("），同業比較法 "), f'TEXT({cadj},"0.00")',
        T(" 美元＝同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（上市未滿十年，不做相對倍數調整）。"))
    put("o2", T("現價隱含的條件：要讓現金流折現法支持現價，"), f'IF(ABS({rv("w_res")})<0.01,"折現率須降到 "&TEXT({rv("w1")},"0.0%"),"折現率即使降到 "&TEXT({rv("w1")},"0.0%")&" 仍不足")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0.0")', T("，模型 "), f'TEXT({rv("rv0")},"0.0")', T(" 十億美元）；同業比較法只需 "), f'TEXT({rv("pe1")},"0.0")', T(" 倍本益比。"))
    put("o3", T("盈餘低於共識：FY2028 EPS 基準情境 "), f'TEXT(Recon!E{reps28},"0.00")', T(" 美元，比共識 "), f'TEXT(Recon!B{reps28},"0.00")', T(" 低 "), f'TEXT(-Recon!F{reps28},"0.0%")',
        T("；主因 FY2028 起 OECD 全球最低稅使有效稅率由 FY2027 錨點的 "), 'TEXT(Q!G13,"0.0%")', T(" 升到 "), 'TEXT(tx_28,"0%")', T("（共識是否已含未找到）。"))
    put("o4", T("核心問題：FY2028 營收基準情境 "), f'TEXT(Recon!E{rr28},"0.00")', T(" 十億美元，共識 "), f'TEXT(Recon!B{rr28},"0.00")', T("（差 "), f'TEXT(Recon!F{rr28},"+0.0%;-0.0%")',
        T("）；FY2028 AEC 營收 "), f'TEXT(Recon!E{raec},"0.00")', T("，比賣方引述的 AEC 共識 "), f'TEXT(Recon!B{raec},"0.0")', T(" 高 "), f'TEXT(Recon!F{raec},"0%")', T("；成長來自層級銅纜市場，份額沿用 FY2027 校準值。"))
    put("o5", T("框架與客戶集中的影響最大：只看現金流折現時加權目標價 "), f'TEXT({SB("DCF 在加權中的權重", "D")},"0")', T(" 美元（"), f'TEXT({SB("DCF 在加權中的權重", "B")},"+0%;-0%")',
        T("），只看同業比較法 "), f'TEXT({SB("DCF 在加權中的權重", "C")},"+0%;-0%")', T("；營運輸入中最大客戶流失 "), f'TEXT({SB("最大客戶流失", "B")},"+0.0%;-0.0%")',
        T("、AEC 份額流失 "), f'TEXT({SB("AEC 有效份額", "B")},"+0.0%;-0.0%")', T("。"))
    put("drv", T("分部營收推導（情境 B）：AEC＝AI 銅纜市場（會計年度加權）× Credo 有效份額（FY2027 校準，FY2028 起乘相對值）；光學＝AI 可插拔光模組市場 × 光學有效份額；"
                 "其他 IC 與 IP 以固定年增推算。三者相加為營收。FY2028：AEC "), f'TEXT(PL!F{aec_b},"0.00")', T("、光學 "), f'TEXT(PL!F{opt_b},"0.00")',
        T("、其他 "), f'TEXT(PL!F{oth_b},"0.00")', T("，合計 "), f'TEXT(PL!F{rev_b},"0.00")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型的 AI 銅纜與光模組市場讀市場層 markets/network，情境係數與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1'] + 3}", T("（Coord 頁 F 節）。C0 以公司實績對層級隱含（J 節）：1Q27 營收 "),
        f'TEXT(Coord!B{CO["rm"]},"0.000")', T(" 十億美元對層級 FY2027 依月數折算 "), f'TEXT(Coord!C{CO["rm"]},"0.000")', T("，差 "), f'TEXT(Coord!D{CO["rm"]},"+0.0%;-0.0%")',
        T("；但 Credo 份額由本模型以 FY2027 錨點校準，層級市場在比式中消去，結論「"), f"Coord!E{CO['rv']}", T("」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")',
        T(" 項。產業部署 GW 兩層差距另列參考；對帳中年份 CY2027–CY2029（FY2028 起）依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（工作單 20261114 第二批第 2 節）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史）、"
                  "PL 頁末標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B，Non-GAAP）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準（整本重算快照 crdo/sens.json）、"
                  "Coord 頁（兩讀數目標價、讀取層級、C0 實績：1Q27 營收 vs 層級隱含，份額同源校準，結論無可比實績）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    """README A1 標題段的舊版號（v0.2.3，舊檔標題未隨版更新）改為本版；tools/coord stamp 已在末尾加本版說明。"""
    RD = wb["README"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and "模型 v0.2.3（" in a1.split("；", 1)[0], a1[:80]
    RD["A1"].value = a1.replace("模型 v0.2.3（", f"模型 v{VER}（", 1)


def apply(wb, here, SIL, CR):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("CRDO_SENS") == "1":
        return "v036_post：略過（alt_run／sens_crdo 建置）"
    assert CR, "Coord 頁不存在：先跑 python3 tools/coord/alt_run.py crdo"
    HR = hist_pl(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb, here)
    CO = coord(wb, SIL)
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    version(wb)
    return (f"v036_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；"
            f"Coord F–J 節（C0 判定第 {CO['rv']} 列）；README v1.3 白話列 {len(RMA)} 列")
