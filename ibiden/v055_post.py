# Ibiden v0.5.5（2026-11-14）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7）與三層協調 C0（規則 v1.4.1）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.5.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司決算短信，J-GAAP；標準損益表的歷史三年與 FY2023 年增率）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列四個目標價
#   Sens 頁末：加權目標價相對基準（低案、高案；業務主因整本重算快照 ibiden/sens.json＋目標價參數表）
#   Coord 頁末：三層協調 C0（讀哪一層哪一版、由下往上對帳、對帳中年份處理、FY(t) 對 CY(t) 口徑近似）；co_flag_<yyyy> 具名範圍
#   In 頁末：N 節（標準損益表推算假設 opx_el）
#   README 頁末：v1.3 展示層白話列；A1 標題段版號更正為現行版號
import json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.5.5", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
JPX26 = "https://www2.jpx.co.jp/disc/40620/140120260505517290.pdf"
JPX24 = "https://www2.jpx.co.jp/disc/40620/140120240430579860.pdf"
SRC_H = (f"公司：2026 年 3 月期決算短信（{JPX26}；FY2025 與 FY2024〔比較欄〕）；2024 年 3 月期決算短信（{JPX24}；FY2023 與 FY2022〔比較欄〕）；"
         "日本會計基準（J-GAAP）連結損益計算書，百萬日圓換算十億日圓")


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
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
    """公式字串：parts 為 ("t", 文字) 或 Excel 運算式（字串）。"""
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
TX = lambda ref, f: f'TEXT({ref},"{f}")'
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- Hist：損益表歷史 ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)；金額十億日圓（百萬日圓 ÷ 1000）；FY2025 的營收、毛利、營業利益、淨利連到 In 頁（已核實值）
HIS = [
    ("rev", "營收（売上高）", 417.549, 370.511, 369.436, "=a_rev"),
    ("cogs", "營業成本（売上原価）", 290.033, 268.040, 256.108, 284.532),
    ("gp", "毛利（売上総利益）", 127.515, 102.471, 113.328, "=a_gp"),
    ("sga", "營業費用（販売費及び一般管理費；含研發）", 55.152, 54.902, 65.706, 69.640),
    ("oi", "營業利益", 72.362, 47.568, 47.621, "=a_oi"),
    ("noi", "營業外收益", 4.952, 4.773, 5.015, 4.759),
    ("noe", "營業外費用", 1.138, 1.201, 4.747, 5.964),
    ("ord", "經常利益", 76.176, 51.140, 47.890, 60.822),
    ("spg", "特別利益（含政策持股出售利益）", 0.604, 5.195, 54.218, 59.281),
    ("spl", "特別損失", 5.078, 8.901, 50.657, 29.011),
    ("pti", "稅前利益（税金等調整前当期純利益）", 71.702, 47.435, 51.451, 91.092),
    ("tax", "所得稅（法人税等合計）", 19.288, 15.685, 17.400, 26.929),
    ("nci", "非支配股東損益", 0.226, 0.259, 0.346, 0.449),
    ("ni", "稅後淨利（歸屬母公司）", 52.187, 31.490, 33.704, "=a_ni"),
]
# 每股盈餘：公告值（基本、潛在株式調整後）與分割倍數（2026-01 一拆二、2026-10 一拆二；FY2024 起公告值已含 2026-01 分割）
EPSR = [("eps_r", "EPS 基本（公司公告值，公告當時股數口徑）", (373.73, 225.44, 120.66, 228.16)),
        ("eps_dr", "EPS 潛在株式調整後（公司公告值；FY2022 無潛在株式）", (None, 224.74, 113.10, 214.91)),
        ("split", "換算到 2026-10 分割後的倍數（公告值 ÷ 本列）", (4, 4, 2, 2))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.5.5 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。FY(t)＝t/4–(t+1)/3：FY2022＝2023 年 3 月期、FY2023＝2024 年 3 月期、"
               "FY2024＝2025 年 3 月期、FY2025＝2026 年 3 月期。金額：十億日圓；EPS：日圓。FY2025 的營收、毛利、營業利益、淨利連到 In 頁 F 節（已核實）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H + ("；FY2025 引用 In 頁" if isinstance(vals[-1], str) else ""))
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.font = BLUE; c.number_format = "0.00" if k != "split" else "0"
        H.cell(r, 7, "Verified" if k != "split" else "Derived")
        H.cell(r, 8, SRC_H if k != "split" else "2026-01-01 一拆二（公司決算短信註記）與 2026-10-01 一拆二（In 頁 a_eps 說明）；FY2022、FY2023 公告值早於兩次分割（÷4），FY2024、FY2025 公告值已含 2026-01 分割（÷2）")
        R[k] = r; r += 1
    for k, labl, f_, fmt in (("eps", "EPS 基本（分割後換算；模型口徑）", "={c}{eps_r}/{c}{split}", "0.00"), ("epsd", "EPS 潛在株式調整後（分割後換算；只作對照）", "=IF({c}{eps_dr}=\"\",\"\",{c}{eps_dr}/{c}{split})", "0.00"),
                             ("sh", "加權平均股數（十億股，分割後；＝淨利 ÷ EPS）", "={c}{ni}/{c}{eps}", "0.0000")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = fmt
        H.cell(r, 7, "Derived"); H.cell(r, 8, "公告值 ÷ 分割倍數；FY2025 與 In 頁 a_eps（228.16 ÷ 2）相同")
        R[k] = r; r += 1
    r += 1
    H.cell(r, 1, "核對（單位：十億日圓；日本公司以百萬日圓四捨五入，各列加總可有 ±0.002 的捨入差）").font = BOLD; r += 1
    chk = [("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"), ("毛利 − 營業費用 − 營業利益", "={c}{gp}-{c}{sga}-{c}{oi}"),
           ("營業利益 ＋ 營業外收益 − 營業外費用 − 經常利益", "={c}{oi}+{c}{noi}-{c}{noe}-{c}{ord}"),
           ("經常利益 ＋ 特別利益 − 特別損失 − 稅前利益", "={c}{ord}+{c}{spg}-{c}{spl}-{c}{pti}"),
           ("稅前利益 − 所得稅 − 非支配股東損益 − 稅後淨利", "={c}{pti}-{c}{tax}-{c}{nci}-{c}{ni}")]
    R["chk0"] = r
    for labl, f_ in chk:
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "核對：FY2025 營業利益（本頁）− In 頁 a_oi（應為 0；連結取值）"); H[f"F{r}"] = f"=F{R['oi']}-a_oi"; H[f"F{r}"].number_format = "0.000"
    R["chk1"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 58
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["H"].width = 90
    return R


# ---------- In：N 節（標準損益表推算假設） ----------
def in_n(wb):
    IN = wb["In"]
    r = IN.max_row + 2
    IN.cell(r, 1, "N. 標準損益表推算假設（v0.5.5 新增；只用於 PL 頁「標準損益表」推算毛利與營業費用，不進 EPS、現金流與目標價）").font = BOLD; r += 1
    vals = ("opx_el", "營業費用彈性（營業費用年增 ÷ 營收年增；FY2026–FY2029 推算毛利用）", "x", 0.3, 0.5, 0.7, "Assumed",
            "模型只推營業利益：毛利＝營業利益＋營業費用（推算）；營業費用＝前一年 ×（1＋彈性 × 營收年增）。FY2023–FY2025 實績營業費用年增 −0.5%、＋19.7%、＋6.0%（Hist 頁）對營收年增 −11.3%、−0.3%、＋12.7%，無穩定比例",
            "Assumed；低／高不進任何目標價（只供讀者觀察毛利率對費用假設的敏感度）")
    for j, v in enumerate(vals, 1):
        c = IN.cell(r, j, v)
        if j in (4, 5, 6): c.font = BLUE
    dn(wb, "opx_el", f"In!$E${r}")
    return r


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"
HC = {"C": "D", "D": "E", "E": "F"}   # 標準損益表歷史欄 → Hist 欄（FY2023→D、FY2024→E、FY2025→F；FY2022 在 C）
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029}   # 預測欄依賴的 Siliconomics 日曆年（FY(t) 對 CY(t)；FY2030–31 延伸期沿用 FY2029 之後的路徑）


def pl_standard(wb, HR):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    g = lambda label: lab(P, label, rb, exact=True)
    R = dict(rev=g("營收"), oi=g("營業利益"), ni=g("淨利（歸屬母公司）"), sh=g("股數（分割後）"), eps=g("EPS（日圓）"))
    xb = lab(P, "延伸｜情境 B")
    X = dict(rev=lab(P, "營收（延伸）", xb), ni=lab(P, "淨利（延伸）", xb), sh=lab(P, "股數（延伸）", xb), eps=lab(P, "EPS（延伸）", xb))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.5.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026–FY2031，情境 B 基準；FY2026E 為公司上修後預測的錨點）。¥B；EPS 日圓。"
                 "只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（日本會計基準 J-GAAP，公司決算短信的連結損益）。公司決算短信、公司預測與市場共識都以此為主，沒有另一套 Non-GAAP；稅前利益含特別損益（政策持股出售利益等，FY2024–FY2025 影響大），"
                 "所以歷史淨利率與營業利益率差距大。預測期的特別損益不估（其他收入＝稅前利益 − 營業利益，由淨利反推）。")
    r_note = r; r += 1
    P.cell(r, 1, "預測期只推營業利益與淨利：FY2026–FY2029 的營業費用＝前一年 ×（1＋彈性 × 營收年增）（彈性見 In 頁 N 節），毛利＝營業利益＋營業費用，營業成本＝營收 − 毛利；"
                 "稅前利益＝淨利 ÷（1 − 稅率），所得稅＝稅前利益 × 稅率，少數股權併入其他收入。FY2030E–FY2031E（收斂法延伸期，只推營收與淨利）：毛利率與營業費用率沿用前一年據以算出營業成本、毛利與營業費用，"
                 "與延伸期營業利益（DCF 頁 EBIT＝淨利 ÷ 淨利對營業利益比）的差額列核對列，不改淨利與目標價（Andy 2026-11-14）。")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + YEARS + ["說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJK":   # 預測欄標題：依賴 Siliconomics 對帳中年份者加「（對帳中）」（co_flag_<yyyy> 在 Coord 頁）
        yr = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr}"&IF(co_flag_{FLAGY[col]}=1,"（對帳中）","")'
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
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    fc4 = lambda k: [f"=PL!{c}{R[k]}" for c in "DEFG"]
    nxt = lambda k, ec: [f"=PL!{c}{X[k]}" for c in ec]
    ri = r   # 本列起算：營收
    row("rev", "營收", "¥B", hh("rev") + fc4("rev") + nxt("rev", "HI"), M, "FY2026E＝公司預測；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2031E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/Hist!C{h}-1".format(rev=L["rev"], h=HR["rev"])] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    rc = r   # 營業成本列
    row("cogs", "營業成本（COGS）", "¥B", hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHI"] + [f"={c}{L['rev']}*(1-{p}{rc + 1}/{p}{L['rev']})" for p, c in (("I", "J"), ("J", "K"))], M,
        "FY2026E–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − 前一年毛利率）")
    ro = rc + 4   # 營業利益列（cogs=rc, gp=rc+1, gm=rc+2, opx=rc+3, oi=rc+4）
    row("gp", "毛利", "¥B", hh("gp") + [f"={c}{ro}+{c}{rc + 3}" for c in "FGHI"] + [f"={c}{L['rev']}-{c}{rc}" for c in "JK"], M, "FY2026E–FY2029E＝營業利益＋營業費用（推算）")
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "預測期為推算值（營業費用彈性 opx_el 假設）；FY2030E–FY2031E＝前一年")
    row("opx", "營業費用（含研發）", "¥B", hh("sga") + [f"={p}{rc + 3}*(1+opx_el*({c}{L['rev']}/{p}{L['rev']}-1))" for p, c in zip("EFGH", "FGHI")]
        + [f"={c}{L['rev']}*{p}{rc + 3}/{p}{L['rev']}" for p, c in (("I", "J"), ("J", "K"))], M, "FY2026E–FY2029E＝前一年 ×（1＋彈性 × 營收年增）；FY2030E–FY2031E＝營收 × 前一年營業費用率")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "¥B", hh("oi") + fc4("oi") + [f"=DCF!{c}{d_ebit}" for c in "GH"], M, "FY2030E–FY2031E 取 DCF 頁情境 B 的 EBIT（延伸期由淨利反推；不改）")
    assert L["oi"] == ro
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）", "¥B", [f"={c}{r + 1}-{c}{L['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益：實績含營業外淨額與特別損益；預測期由淨利反推（少數股權併入）")
    row("pti", "稅前利益", "¥B", hh("pti") + [f"={c}{r + 2}/(1-tax_w)" for c in "FGHIJK"], M, "預測期＝稅後淨利 ÷（1 − 稅率）")
    row("tax", "所得稅", "¥B", hh("tax") + [f"={c}{L['pti']}*tax_w" for c in "FGHIJK"], M, "預測期＝稅前利益 × 稅率（In 頁稅率，近三年平均）")
    row("ni", "稅後淨利（歸屬母公司）", "¥B", hh("ni") + fc4("ni") + nxt("ni", "HI"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("nci", "（備註）非支配股東損益", "¥B", hh("nci") + [None] * 6, M, "預測期併入其他收入，不另列")
    row("sh", "股數（十億股，分割後）", "十億股", [f"=Hist!{HC[c]}{HR['sh']}" for c in "CDE"] + fc4("sh") + nxt("sh", "HI"), "0.0000")
    row("eps", "EPS（基本；分割後）", "¥", [f"=Hist!{HC[c]}{HR['eps']}" for c in "CDE"] + fc4("eps") + nxt("eps", "HI"), "0.00", "模型口徑為基本 EPS（淨利 ÷ 股數）；完全稀釋（含 CB）見下方對照列")
    row("epsd", "（備註）EPS 完全稀釋（公司公告，分割後換算；只列實績）", "¥", [f"=Hist!{HC[c]}{HR['epsd']}" for c in "CDE"] + [None] * 6, "0.00")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/Hist!C{h}-1".format(ni=L["ni"], h=HR["ni"])] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "¥B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2031E 為差額，見下）", "¥B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None, None], "0.000")
    row("c2b", "差額（FY2030E–FY2031E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利反推）；目標價用延伸期淨利，不受影響", "¥B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "差額（FY2030E–FY2031E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JK"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利（預測期；歷史含少數股權）", "¥B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000", "歷史欄差額＝非支配股東損益")
    row("c4", "核對：稅後淨利 − PL 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "¥B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    row("c5", "核對：營收 − PL 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "¥B",
        [None] * 3 + [f"={c}{L['rev']}-PL!{s}{R['rev']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['rev']}-PL!H{X['rev']}", f"=K{L['rev']}-PL!I{X['rev']}"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型 FY(t)＝t 年 4 月至 (t+1) 年 3 月，引擎對 FY(t) 直接取 CY(t) 的加速器顆數與 Siliconomics 年別比值（重疊 9 個月，不平移）；"
                 "標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2031E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。口徑近似見 Coord 頁 C0 節。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, rows_X=X, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7   # DCF 頁：評價日後占比、折現年數（期中）
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數"), (D.cell(fr, 1).value, D.cell(ty, 1).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"(dcf_nc+DCF!$C${fcf}*stub_frac-stub_div)"   # 淨現金滾到評價日（情境 B；第一年 FCFF 為錨點，不隨反解變動）
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.5.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2026（錨點）與 FY2027（三情境相同）不動，營業利益率與營收倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("px", "現價（日圓）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（日圓）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；¥B）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（¥B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；¥B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（日圓）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2026, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}
    NC = len("BCDEFGH")

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGHI"):
            c = RV.cell(r, 2 + i, fn(dc, "BCDEFGH"[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（¥B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（¥B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "I" else "=0"), "0.0000")
    yr("k", "每 1 日圓營業利益對 EV 的貢獻（FY2026–FY2027 不動，記 0）", lambda dc, c: ("=0" if dc in "CD" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（¥B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（¥B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（¥B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（¥B）", f"=SUMPRODUCT(D{Y['rev']}:H{Y['rev']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（¥B）", f"=B{A['ev']}-SUMPRODUCT(D{Y['ebit']}:H{Y['ebit']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2026、FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2026＋FY2027 現值（¥B）", f"=B{Y['pv']}+C{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（¥B）",
       f"=SUM(D{Y['pv']}:H{Y['pv']})+DCF!I{nop}*H{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（¥B）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；¥B）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近下界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（日圓）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$I${fcf},DCF!$C${fr}:$I${fr},1/(1+{w})^DCF!$C${ty}:$I${ty})"
                f"+DCF!$I${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C$8+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.001"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "#,##0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.1 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（日圓，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；日圓）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(" 日圓；要支持現價 "), TX("px", "#,##0"), T(" 日圓，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億日圓，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(" 日圓相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，高於同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × Ibiden 過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給 Ibiden 的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，低於歷史溢價 "), TX(coef, "0.00"), T(" 倍。")))
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
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.5.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "日圓", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(" 日圓；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=Val!A6"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    snap = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    tp_cell = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    t0 = lab(SE, "目標價參數敏感度"); hr = t0 + 1
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.5.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（ibiden/sens_ibiden.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", "低案：加權目標價（日圓）", "高案：加權目標價（日圓）", "快照基準加權目標價（日圓）"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    b = snap["_base"]["tp"]
    for case in snap["cases"]:
        k = case["key"]; z = snap["biz"][k]
        SE.cell(r, 1, case["label"]); SE.cell(r, 4, case["setting"] + "（整本重算快照）")
        for col, side, tcol in ((2, "低", 5), (3, "高", 6)):
            if side in z:
                c = SE.cell(r, tcol, z[side]["tp"]); c.font = BLUE; c.number_format = "#,##0.00"
                SE.cell(r, col, f"={'EF'[tcol - 5]}{r}/$G{r}-1").number_format = "+0.0%;-0.0%"
        c = SE.cell(r, 7, b); c.font = BLUE; c.number_format = "#,##0.00"
        r += 1
    for sr in src:
        SE.cell(r, 1, f"=Sens!A{sr}"); SE.cell(r, 4, f"=\"Sens 頁目標價參數敏感度：\"&Sens!B{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
        r += 1
    last = r - 1
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 ibiden/sens_ibiden.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r)


# ---------- Coord：C0 ----------
SILC = {  # Siliconomics 層值（快照 v0.19；blue 字面值，附鍵名）
}


def coord_c0(wb, SIL, CR, HR, PLA):
    from openpyxl.worksheet.worksheet import Worksheet  # noqa: F401
    CW = wb["Coord"]
    # 具名範圍：對帳中旗標鏡像（co_flag_<yyyy>；check_layout_v13 與 PL 標題用）
    for j, y in enumerate((2026, 2027, 2028, 2029)):
        dn(wb, f"co_flag_{y}", f"Coord!${'BCDE'[j]}${CR['flag']}")
    P = wb["PL"]; U = wb["Units"]; V = wb["Val"]
    rb = lab(P, "情境 B：")
    g = lambda label: lab(P, label, rb, exact=False)
    R = dict(unv=lab(P, "NVIDIA 加速器顆數 × 情境係數", rb), uot=lab(P, "非 NVIDIA 加速器顆數 × 情境係數", rb), ueff=lab(P, "有效顆數", rb),
             asp=lab(P, "每顆載板金額指數", rb), rel=lab(P, "Ibiden 份額相對值", rb), pp=lab(P, "短缺溢價指數", rb), ai=lab(P, "AI 伺服器載板營收", rb))
    ra = lab(P, "情境 A", 1)
    RA = dict(unv=lab(P, "NVIDIA 加速器顆數 × 情境係數", ra), uot=lab(P, "非 NVIDIA 加速器顆數 × 情境係數", ra))
    ur_ai_usd = lab(U, "AI 伺服器載板營收（美元）"); ur_asp = lab(U, "每顆 NVIDIA 載板金額")
    N = SIL.NAMES; Vs = SIL.V
    base_tot = N["SIF_UnitsPlat_Total_Base"]; raw_tot = N["SIF_UnitsPlat_Total"]
    ibsh, vot = Vs["abf_ibsh26"]["base"], Vs["abf_vot"]["base"]
    abfm, g26, ai26, acc26 = Vs["abfm_25"]["base"], Vs["abfm_g26"]["base"], Vs["abf_ai26"]["base"], Vs["abf_accsh26"]["base"]
    cap = {y: Vs[f"abfcap_ib_{y % 100}"]["base"] for y in (2027, 2028, 2029)}
    nv26, ot26 = N["SIF_UnitsPlat_bw"][1] + N["SIF_UnitsPlat_rb"][1] + N["SIF_UnitsPlat_ru"][1] + N["SIF_UnitsPlat_hp"][1], \
        N["SIF_UnitsPlat_amd"][1] + N["SIF_UnitsPlat_tpu"][1] + N["SIF_UnitsPlat_trn"][1] + N["SIF_UnitsPlat_oth"][1]
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調 C0：由下往上對帳（v0.5.5；三層基準協調規則 v1.4.1 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = []
    rd.append(("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（只讀 SIF_／abf_ 名稱，不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股，讀 Siliconomics"))
    rd.append(("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁 B 節）與 Sens 敏感度案例；不改基準"))
    rd.append(("加速器顆數（NVIDIA 四世代、非 NVIDIA 四平台）", "唯讀上游", "markets/network In 頁 u_*／g29_*（與 Siliconomics SIF_UnitsPlat_* 同值，本頁 2 節逐年核對）", "AI 網通市場模型 v0.2.3", "PL 引擎 A 節；情境 A＝原始顆數"))
    rd.append(("顆數情境係數 B（CY2028）", "=um28_B", "Siliconomics SIF_UnitsCoef_Base（經 tsmc 模型 In 頁 um28_B 讀入；規則 C3 依證據狀態決定）", SIL.VERSION, "情境 B（基準）；A＝1.0、C 個股自訂"))
    rd.append(("顆數情境係數 B（CY2029）", "=um29_B", "同上（um29_B）", SIL.VERSION, "同上"))
    for key, nm, silk in (("w_ot", "非 NVIDIA 加速器權重（Ibiden 份額 × 每顆載板金額，相對 NVIDIA）", "abf_wot_ib"), ("g_asp", "每顆載板金額年增", "abf_gasp"),
                          ("rel_27", "Ibiden 份額相對值 2027", "abf_ibrel_27"), ("rel_28", "Ibiden 份額相對值 2028", "abf_ibrel_28"), ("rel_29", "Ibiden 份額相對值 2029", "abf_ibrel_29"),
                          ("pp_27", "價格溢價指數 2027", "abf_pp_27"), ("pp_28", "價格溢價指數 2028", "abf_pp_28"), ("pp_29", "價格溢價指數 2029", "abf_pp_29")):
        rd.append((nm, f"={key}", f"Siliconomics {silk}（SIL.row9 讀入 In 頁 {key}）", SIL.VERSION, "PL 引擎 A 節（份額、金額、溢價）"))
    for y in (2027, 2028, 2029):
        rd.append((f"Wattonomics 錨定係數 CY{y}（敏感度案例用）", f"=wta_{y % 100}", "Siliconomics SIF_WTAnchor_Base（R4 比值：Wattonomics ÷ 本庫推得上線 GW）", SIL.VERSION, "Sens 頁「Wattonomics 錨定」案例；基準不乘"))
    rd.append(("對帳中旗標與另一側讀數比值（CY2026–CY2029）", "見本頁上方 B 節", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION, "兩讀數目標價（C4）"))
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
        r += 1
    r += 1
    # ---- 2. C0 對帳 ----
    CW.cell(r, 1, "2. C0 由下往上對帳（模型自己的產業相關量 vs 層的對應值；差距 > ±15% 標「對帳中」，不改上游值，由主控開對帳單）").font = BOLD; r += 1
    CW.cell(r, 1, "年度口徑近似：本模型 FY(t)＝t 年 4 月–(t+1) 年 3 月，引擎對 FY(t) 直接取 CY(t) 的顆數與年別比值（重疊 9 個月，不平移、不月數加權）；層的值為日曆年值，故 FY(t) 對 CY(t) 的比較是近似。月數加權（9／12 × CY(t) ＋ 3／12 × CY(t+1)）的層值列在備註欄供對照。")
    ROW["caveat"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定（|差距| > 15% 為對帳中）", "層級名稱與版本／口徑／證據等級／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr2"] = r; r += 1
    first = r

    def line(label, unit, yr_, f_model, f_layer, fmt, note, gapfmt="+0.0%;-0.0%"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_model).number_format = fmt
        c = CW.cell(r, 5, f_layer); c.number_format = fmt
        if not (isinstance(f_layer, str) and f_layer.startswith("=")): c.font = BLUE
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = gapfmt
        CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"對帳中","一致")')
        CW.cell(r, 8, note)
        r += 1
        return r - 1
    def line_ref(label, unit, yr_, f_a, f_b, fmt, note):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_a).number_format = fmt; CW.cell(r, 5, f_b).number_format = fmt
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, note)
        r += 1
    ycol = {2026: "D", 2027: "E", 2028: "F", 2029: "G"}   # PL 引擎欄（FY2026–FY2029）
    nb = f"Siliconomics {SIL.VERSION}"
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        c = ycol[y]
        line("加速器顆數合計（情境 B，NVIDIA＋非 NVIDIA）vs Siliconomics 基準", "百萬顆", f"CY{y}", f"=PL!{c}{R['unv']}+PL!{c}{R['uot']}", base_tot[i + 1], "#,##0.000",
             f"{nb} SIF_UnitsPlat_Total_Base（原始顆數 × 基準係數 SIF_UnitsCoef_Base）；Derived；模型情境 B 係數與基準相同，故應為 0")
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        c = ycol[y]
        line("加速器顆數合計（情境 A，原始顆數）vs Siliconomics 原始顆數", "百萬顆", f"CY{y}", f"=PL!{c}{RA['unv']}+PL!{c}{RA['uot']}", raw_tot[i + 1], "#,##0.000",
             f"{nb} SIF_UnitsPlat_Total；Derived；來源 markets/network（AI 網通市場模型 v0.2.3）")
    for lbl, key, silv in (("非 NVIDIA 加速器權重", "w_ot", Vs["abf_wot_ib"]["base"]), ("每顆載板金額年增", "g_asp", Vs["abf_gasp"]["base"]), ("Ibiden 份額相對值 2027", "rel_27", Vs["abf_ibrel_27"]["base"]),
                           ("Ibiden 份額相對值 2028", "rel_28", Vs["abf_ibrel_28"]["base"]), ("Ibiden 份額相對值 2029", "rel_29", Vs["abf_ibrel_29"]["base"]),
                           ("價格溢價指數 2027", "pp_27", Vs["abf_pp_27"]["base"]), ("價格溢價指數 2028", "pp_28", Vs["abf_pp_28"]["base"]), ("價格溢價指數 2029", "pp_29", Vs["abf_pp_29"]["base"])):
        line(f"{lbl}：模型 In 頁 vs Siliconomics", "x" if key != "g_asp" else "%", "—", f"={key}", silv, "0.000",
             f"{nb} abf_* 鍵（直接讀取，鏈路完整性核對；證據等級見 Siliconomics In 頁：份額與溢價多為 Assumed／Interested-party）")
    # 份額與每顆金額（CY2026）
    mkt = f"{abfm}*(1+{g26})*{ai26}*{acc26}"
    r_mkt = r
    CW.cell(r, 1, "加速器載板市場 CY2026（層推導：全球 ABF 載板產值 × (1＋年增) × AI 用占比 × 加速器載板占比）"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, "CY2026")
    CW.cell(r, 5, f"={mkt}").number_format = "#,##0.000"
    CW.cell(r, 8, f"{nb} abfm_25 {abfm}、abfm_g26 {g26}、abf_ai26 {ai26}、abf_accsh26 {acc26}（TPCA／IEK、Palliser Capital：Interested-party；abf_accsh26 為 Derived，由本模型 Units 頁每顆金額 US$215 推得）；本列無模型對應值，供下列兩列使用")
    r += 1
    rr = nv26 and (ot26 / nv26)
    line("Ibiden 占加速器載板金額 FY2026：模型（AI 載板營收 ÷ 層的加速器載板市場）vs 層（Ibiden 份額 × 權重推導）", "%", "FY2026／CY2026",
         f"=Units!C{ur_ai_usd}/E{r_mkt}", f"={ibsh}*(1+w_ot*{ot26}/{nv26})/(1+{vot}*{ot26}/{nv26})", "0.0%",
         f"層值＝abf_ibsh26 {ibsh}（BofA 經 thebuildout，Interested-party）×(1＋權重 × 非 NVIDIA 顆數 ÷ NVIDIA 顆數)÷(1＋abf_vot {vot} × 同比)；"
         "同源註記：層的加速器載板市場占比 abf_accsh26 本身由本模型 Units 頁每顆金額 US$215 推得，兩者不獨立，一致不構成獨立證據（C5：個股預測不往上流，建議對帳單釐清）")
    line("NVIDIA 每顆載板金額 FY2026（美元）：模型 Units 頁 vs 層（加速器載板市場 ÷ 等效顆數）", "US$", "FY2026／CY2026", f"=Units!C{ur_asp}",
         f"=E{r_mkt}*1000/({nv26}+{vot}*{ot26})", "#,##0.0",
         f"層值＝市場 ÷（NVIDIA 顆數 ＋ abf_vot × 非 NVIDIA 顆數）；同源註記：同上（abf_accsh26 由本模型 US$215 推得）。先驗比對推導區間 $150–300（Recon 頁）")
    # 產能指數
    for y in (2027, 2028, 2029):
        c = ycol[y]
        wm = f"(0.75*{cap[y]}+0.25*{cap[y + 1]})" if y + 1 in cap else None
        note = (f"{nb} abfcap_ib_{y % 100}（Ibiden ABF 產能指數，2026＝1；Assumed，依公司路線圖「2028 年產能為 2024 年 2.5–2.8 倍」）；模型值＝有效顆數 × 份額相對值 × 金額指數 ÷ FY2026 有效顆數"
                "（面積與層數隨金額指數，不含短缺溢價），情境 B；口徑：層的指數涵蓋全部 ABF 產能，模型值只含 AI 載板，且 FY(t) 與 CY(t) 重疊 9 個月，可比性中")
        if wm: note += f"；月數加權層值 {wm.replace('0.75*', '0.75×').replace('0.25*', '0.25×')}＝{0.75 * cap[y] + 0.25 * cap[y + 1]:.3f}"
        line("Ibiden AI 載板實體量指數（情境 B）vs Ibiden ABF 產能指數", "x（FY2026＝1）", f"FY{y}／CY{y}", f"=PL!{c}{R['ueff']}*PL!{c}{R['rel']}*PL!{c}{R['asp']}/PL!$D${R['ueff']}", cap[y], "0.000", note)
    last = r - 1
    ROW["first"], ROW["last"] = first, last
    CW.cell(r, 1, "對帳中項目數（|差距| > 15%）"); CW.cell(r, 7, f'=COUNTIF(G{first}:G{last},"對帳中")'); ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：AI IT 資本支出（模型 itc_* 由顆數 × 每顆 IT capex 推得，與 Siliconomics 部署 GW × 每 GW IT capex 同源，等於顆數核對的重述）；每 GW Ibiden 營收（Units 頁，為輸出而非輸入，層沒有對應值）；Ibiden 實際營收（層不持有）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是 Ibiden 自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        c = ycol[y]
        line_ref("加速器顆數合計：Siliconomics 基準 vs Wattonomics 換算值", "百萬顆", f"CY{y}", f"=E{first + i}", f"=E{first + i}*{ratio(y)}", "#,##0.000",
             f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived（層推導值）；對帳中年份依 C4 並列兩讀數目標價（本頁 B 節），個股不裁決")
    ROW["ref2"] = r - 1
    r += 2
    # ---- 3. 對帳中年份 ----
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t) 對 CY(t)：FY2027–FY2029 及其後由 FY2029 收斂推得的 FY2030–FY2031 延伸期；FY2026 為公司預測錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 日圓（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 日圓（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（顆數 × 另一側比值），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("FY(t) 對 CY(t) 的口徑近似：本模型 FY2027 直接乘 CY2027 比值 "), TX(f"C{CR['ratio']}", "0.000"), T("；若依月數加權（9／12 × CY2027 ＋ 3／12 × CY2028），比值為 "),
                    TX(f"0.75*C{CR['ratio']}+0.25*D{CR['ratio']}", "0.000"), T("，差 "), TX(f"(0.75*C{CR['ratio']}+0.25*D{CR['ratio']})/C{CR['ratio']}-1", "+0.0%;-0.0%"),
                    T("；FY2028 直接乘 "), TX(f"D{CR['ratio']}", "0.000"), T("，月數加權為 "), TX(f"0.75*D{CR['ratio']}+0.25*E{CR['ratio']}", "0.000"), T("。兩者差距小於兩讀數差距，不改變評等；本模型採直接對應（與其他載板廠一致），揭露於此。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 22; CW.column_dimensions["H"].width = 110
    return ROW


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, HR, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; RV = wb["RevDCF"]; DC = wb["DCF"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    wacc_c = "wacc"
    rb = lab(P, "情境 B：")
    pa, pb = lab(P, "情境 A"), lab(P, "情境 B")
    em_a = lab(P, "電子事業營益率", pa); eps_a = lab(P, "EPS", pa, exact=True); eps_b = lab(P, "EPS", pb, exact=True); em_b = lab(P, "電子事業營益率", pb)
    tv_row = lab(DC, "終值現值占 EV", lab(DC, "情境 B（基準）"));
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_pm, s_sh, s_wts = sens_r("短缺溢價路徑"), sens_r("Ibiden 份額相對值"), sens_r("情境權重")
    s_w = next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").strip().startswith("=Sens!A") and "DCF 權重" in str(wb["Sens"].cell(int(re.search(r"A(\d+)", SE.cell(r, 1).value).group(1)), 1).value))
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    PLR = PLA["rows_B"]
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 日圓，低於現價 "), TX(f"-{up}", "0.0%"), T("，評等"), rec, T("：同業比較法 "), TX(cadj, "#,##0"), T(" 日圓高於現價，但現金流折現只有 "), TX(dcf, "#,##0"), T(" 日圓。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：Comparable 採用 Ibiden 過去十年相對同業的本益比溢價（"), TX(coef, "0.00"), T(" 倍）、擴產期資本支出照公司排程。不採溢價時加權目標價 "),
        TX(noadj, "#,##0"), T(" 日圓（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!E{s_w}" if False else f"Sens!C{int(re.search(r'A(\d+)', SE.cell(s_w, 1).value).group(1))}", "#,##0"),
        T(" 日圓（"), TX(f"Sens!C{int(re.search(r'A(\d+)', SE.cell(s_w, 1).value).group(1))}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{int(re.search(r'A(\d+)', SE.cell(s_w, 1).value).group(1))}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 日圓，比現價 "), TX("px", "#,##0"), T(" 日圓低 "), TX(f"-{up}", "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 溢價 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0"), T(" 日圓）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 日圓遠高於本模型。"))
    put("o2", T("DCF 偏低：擴產期資本支出大（FY2026 設備投資 "), TX("cxs_26*10", "#,##0"), T(" 億日圓、FY2027–FY2028 各約 "), TX("cxs_27*10", "#,##0"), T(" 億），折舊較小，明確期自由現金流薄；終值現值占 EV "),
        TX(f"DCF!C{tv_row}", "0%"), T("。WACC "), TX(wacc_c, "0.0%"), T("；要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"), T("，或 FY2028 起現金流（營收與成本同步）放大 "), TX(rv("s1"), "0.00"), T(" 倍。"))
    put("o3", T("最大變數是短缺溢價與份額（加權目標價相對基準）：溢價快速回落 "), TX(f"Sens!B{s_pm}", "+0.0%;-0.0%"), T("、延續 "), TX(f"Sens!C{s_pm}", "+0.0%;-0.0%"), T("；份額流失 "), TX(f"Sens!B{s_sh}", "+0.0%;-0.0%"),
        T("；情境權重全給保守 "), TX(f"Sens!B{s_wts}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o4", T("核心問題：溢價回落的代價。情境 A 的電子事業營益率由 FY2027 "), TX(f"PL!E{em_a}", "0.0%"), T(" 降到 FY2029 "), TX(f"PL!G{em_a}", "0.0%"), T("；顆數較少的情境 B 的 FY2028 EPS "), TX(f"PL!F{eps_b}", "0"),
        T(" 日圓，比情境 A 的 "), TX(f"PL!F{eps_a}", "0"), T(" 低 "), TX(f"-(PL!F{eps_b}/PL!F{eps_a}-1)", "0.0%"), T("。起點 AI 占電子事業 "), TX("ai_e", "0%"), T(" 是假設（公司未揭露）。"))
    put("o5", T("對帳中：CY2027–CY2029 的加速器顆數，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 日圓（"), f"Coord!C{CR['rec']}", T("），與現行同評等。個股不裁決，對帳在 Siliconomics。"))
    put("drv", T("分部營收推導（情境 B）：電子事業＝AI 伺服器載板（有效加速器顆數 × 每顆有效金額 × 金額指數 × 份額相對值 × 溢價指數）＋非 AI 部分（逐年成長）；陶瓷事業與其他事業各依成長率推算；三事業相加為合併營收。FY2028 AI 載板 "),
        TX(f"PL!F{lab(P, 'AI 伺服器載板營收', rb)}", "0"), T("、電子事業 "), TX(f"PL!F{lab(P, '電子事業營收', rb)}", "0"), T("、陶瓷 "), TX(f"PL!F{lab(P, '陶瓷事業營收', rb)}", "0"), T("、其他 "), TX(f"PL!F{lab(P, '其他事業營收', rb)}", "0"),
        T("，合計 "), TX(f"PL!F{lab(P, '營收', rb, exact=True)}", "0"), T(" 十億日圓。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA）與三層協調 C0。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2022–FY2025，公司決算短信）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B，GAAP）、"
                  "Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁 C0 對帳（讀哪一層哪一版、由下往上對帳、FY(t) 對 CY(t) 口徑近似）、In 頁 N 節（營業費用彈性，只用於推算毛利）、README 頁末 v1.3 白話列；"
                  "README A1 標題段版號由 v0.2.1 更正為現行版號。既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    assert "模型 v0.2.1（AI 半導體 Project" in a1, a1[:60]
    a1 = a1.replace("模型 v0.2.1（AI 半導體 Project", f"模型 v{VER}（AI 半導體 Project", 1)
    RD["A1"].value = a1 + (f"；v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（章節順序、標準損益表、反向 DCF 頁、四個目標價、加權目標價相對基準敏感度）與三層協調 C0（Coord 頁新增由下往上對帳）；"
                           "README A1 標題段版號由 v0.2.1 更正為現行版號；既有數值不變")


def apply(wb, here, SIL, CR):
    # check_layout_v13 的兩讀數檢查要求 Coord 頁該列標籤恰為「加權目標價」（tools/coord 的 NAMES 為較長的標籤；本列只改文字標籤，數值與公式不變）
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    in_n(wb)
    CO = coord_c0(wb, SIL, CR, HR, None)
    PLA = pl_standard(wb, HR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    version(wb)
    wb._v055 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v055_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 加權目標價相對基準 {SNS['t0']} 起；Coord C0 {CO['first']}–{CO['last']}；README v1.3 白話列 {len(RMA)} 列"
