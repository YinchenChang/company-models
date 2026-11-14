# 泰瑞達 v0.4.5（2026-11-14）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Ibiden v0.5.6）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.4.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司 10-K；US GAAP；標準損益表的歷史三年與 FY2023 年增率）
#   In 頁末：N 節（標準損益表推算假設 oth_net）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列四個目標價
#   Sens 頁末：加權目標價相對基準（低案、高案；業務主因整本重算快照 ter/sens.json＋目標價參數表）
#   Coord 頁末：三層協調——1. 讀取層級、2. 讀取層級核對（同源，不判定）、層級之間（參考）、3. 對帳中年份、4. C0（實績）（SoC 測試機營收 2026 Q1–Q3 vs 層級隱含值）
#   Sources 頁：新增 3Q26 10-Q 與 FY2025 10-K 損益表來源；README 頁末：v1.3 展示層白話列與本版說明
import glob, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.5", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
K10 = "https://www.sec.gov/Archives/edgar/data/97210/000119312526059002/ter-20251231.htm"
Q10 = "https://www.sec.gov/Archives/edgar/data/97210/000119312526441091/ter-20260927.htm"
SRC_H = (f"公司 FY2025 10-K（Item 8 合併損益表，第 41 頁；{K10}）：FY2025、FY2024、FY2023 三欄；US GAAP，百萬美元換算（千美元 ÷ 1,000）")
SRC_Q3 = (f"公司 3Q26 10-Q（2026-10-30 提交，季末 2026-09-27；Part I Item 1 附註 E 收入「Disaggregation of Revenue」，第 9 頁；網址見 Sources 頁）。"
          "欄位對照：SoC、記憶體、IST、機器人、Product Test；以 3Q25 欄（SoC 440.216、記憶體 128.062、IST 37.589、機器人 75.087、Product Test 88.256）對 Q 頁 3Q25 實績核對欄位順序")


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
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)；金額百萬美元（10-K 千美元 ÷ 1,000）；FY2022 只有營收、淨利、EPS（取 Val 頁第 1 節 SEC XBRL，供 FY2023 年增率）
HIS = [
    ("rev", "營收（Total revenues）", "=Val!F13", 2676.298, 2819.880, 3190.024),
    ("cogs", "營業成本（Total cost of revenues）", None, 1139.550, 1170.953, 1332.679),
    ("gp", "毛利（Gross profit）", None, 1536.748, 1648.927, 1857.345),
    ("sga", "銷售與管理費用（Selling and administrative）", None, 577.315, 617.047, 648.874),
    ("ed", "工程與研發費用（Engineering and development）", None, 418.089, 460.876, 504.596),
    ("amort", "取得無形資產攤銷（Acquired intangible assets amortization）", None, 18.999, 18.764, 15.270),
    ("restr", "重組與其他（Restructuring and other）", None, 21.277, 15.571, 38.554),
    ("gain", "出售事業利益（Gain on sale of business；FY2024 出售 DIS 事業，負值＝減少費用）", None, 0.0, -57.119, 0.0),
    ("opx", "營業費用合計（Total operating expenses）", None, 1035.680, 1055.139, 1207.294),
    ("oi", "營業利益（Income from operations）", None, 501.068, 593.788, 650.051),
    ("ii", "利息收入（Interest income）", None, 27.348, 24.772, 15.696),
    ("ie", "利息費用（Interest expense）", None, 3.806, 3.587, 6.846),
    ("oe", "其他費用淨額（Other (income) expense, net；負值＝收入）", None, -0.962, 5.887, 5.641),
    ("pti", "稅前利益（扣權益法前；Income before income taxes and equity in net earnings of affiliate）", None, 525.572, 609.086, 653.260),
    ("tax", "所得稅（Income tax provision）", None, 76.820, 59.503, 79.299),
    ("eq", "權益法投資損益（Equity in net earnings of affiliate；Technoprobe，負值＝損失）", None, 0.0, -7.211, -19.914),
    ("ni", "稅後淨利（Net income；10-K 無非控制權益列）", "=Val!E13", 448.752, 542.372, 554.047),
]
EPSR = [("eps_b", "EPS 基本（公司公告值）", (None, 2.91, 3.41, 3.48)),
        ("eps_d", "EPS 稀釋（公司公告值；GAAP）", ("=Val!B13", 2.73, 3.32, 3.47)),
        ("sh_d", "稀釋加權平均股數（百萬股）", (None, 164.304, 163.314, 159.719))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.4.5 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度＝日曆年（FY2022＝2022-01～12）。金額：百萬美元；EPS：美元。"
               "FY2023–FY2025 取公司 FY2025 10-K 合併損益表（三年並列）；FY2022 只取營收、淨利、EPS（Val 頁第 1 節，SEC XBRL），供 FY2023 年增率")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H + ("；FY2022 引用 Val 頁第 1 節（SEC XBRL）" if isinstance(vals[0], str) else ""))
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k != "sh_d" else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H + ("；FY2022 引用 Val 頁第 1 節" if isinstance(vals[0], str) else ""))
        R[k] = r; r += 1
    H.cell(r, 1, "h_oth"); H.cell(r, 2, "其他收入淨額（＝利息收入 − 利息費用 − 其他費用淨額；＝稅前利益 − 營業利益）")
    for c in "DEF":
        H[f"{c}{r}"] = f"={c}{R['ii']}-{c}{R['ie']}-{c}{R['oe']}"; H[f"{c}{r}"].number_format = "#,##0.000"
    H.cell(r, 7, "Derived"); H.cell(r, 8, "10-K 損益表各列加減；與稅前利益 − 營業利益相同（核對列）")
    R["oth"] = r; r += 2
    H.cell(r, 1, "核對（單位：百萬美元；10-K 以千美元列示，各列加總可有 ±0.002 的捨入差）").font = BOLD; r += 1
    chk = [("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"),
           ("營業費用各項（銷管＋研發＋攤銷＋重組＋出售利益）− 營業費用合計", "={c}{sga}+{c}{ed}+{c}{amort}+{c}{restr}+{c}{gain}-{c}{opx}"),
           ("毛利 − 營業費用合計 − 營業利益", "={c}{gp}-{c}{opx}-{c}{oi}"),
           ("營業利益 ＋ 其他收入淨額 − 稅前利益", "={c}{oi}+{c}{oth}-{c}{pti}"),
           ("稅前利益 − 所得稅 ＋ 權益法投資損益 − 稅後淨利", "={c}{pti}-{c}{tax}+{c}{eq}-{c}{ni}")]
    R["chk0"] = r
    for labl, f_ in chk:
        H.cell(r, 2, labl)
        for c in "DEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "核對：FY2025 營收（本頁，10-K）− Q 頁 FY2025 營收（四季加總）（Q 頁以百萬美元一位小數列示）"); H[f"F{r}"] = f"=F{R['rev']}-a25_rev"; H[f"F{r}"].number_format = "0.000"
    R["chk1"] = r; r += 1
    H.cell(r, 2, "核對：FY2025 稅後淨利（本頁）− Q 頁 FY2025 淨利"); H[f"F{r}"] = f"=F{R['ni']}-a25_ni"; H[f"F{r}"].number_format = "0.000"
    R["chk2"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 70
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["H"].width = 100
    return R


# ---------- In：N 節（標準損益表推算假設） ----------
def in_n(wb):
    IN = wb["In"]
    r = IN.max_row + 2
    IN.cell(r, 1, "N. 標準損益表推算假設（v0.4.5 新增；只用於 PL 頁「標準損益表」把稅前利益拆成營業利益與其他收入，不進 EPS、現金流與目標價）").font = BOLD; r += 1
    vals = ("oth_net", "其他收入淨額每年（利息收入 − 利息費用 − 其他費用；＝稅前利益 − 營業利益；預測期用）", "US$M", None,
            "=((Q!G15+Q!H15+Q!I15)-(Q!G14+Q!H14+Q!I14))*12/9", None, "Derived",
            "前三季 FY2026 實績年化：（稅前利益 − 營業利益）三季合計 × 12 ÷ 9（Q 頁 G–I 欄第 14、15 列；1Q26–3Q26 合計 0.8）",
            "FY2023–FY2025 為 24.5／15.3／3.2（Hist 頁）、上半年 FY2026 −1.5、3Q26 +2.3（Q 頁）；金額小，營業利益≈稅前利益；Derived，不設低／高")
    for j, v in enumerate(vals, 1):
        c = IN.cell(r, j, v)
        if j in (4, 5, 6) and v is not None and not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    IN.cell(r, 5).number_format = "0.0"
    dn(wb, "oth_net", f"In!$E${r}")
    return r


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "D", "D": "E", "E": "F"}   # 標準損益表歷史欄 → Hist 欄（FY2023→D、FY2024→E、FY2025→F；FY2022 在 C）
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}   # 預測欄依賴的 Siliconomics 日曆年（FY(t)＝CY(t)；FY2030–32 延伸期沿用 CY2029 旗標）
EC = {"J": "H", "K": "I", "L": "J"}   # 標準損益表 FY2030–32 欄 → PL 延伸區塊欄
DCOL = {"J": "F", "K": "G", "L": "H"}  # 標準損益表 FY2030–32 欄 → DCF 頁欄（FY2027＝C … FY2032＝H）
EC4 = {"F": "D", "G": "E", "H": "F", "I": "G"}   # 標準損益表 FY2026–29 欄 → PL 推導區塊欄


def pl_standard(wb, HR, CR, PB):
    P = wb["PL"]
    R = PB["B"]
    xb = lab(P, "情境（延伸）")
    xs = lab(P, "情境 B（基準）", xb, exact=True)
    X = dict(rev=lab(P, "營收", xs, exact=True), ni=lab(P, "淨利（收斂法延伸）", xs), eps=lab(P, "EPS（延伸）", xs))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.4.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加 4Q26 指引中點的錨點）。US$M；EPS 美元。"
                 "只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（美國會計準則，公司 10-K、10-Q 的合併損益）。公司同時公布 Non-GAAP（EPS 另加權益法投資攤銷與無形資產攤銷），共識 EPS 為 Non-GAAP；本表與模型一致用 GAAP，EPS 為稀釋 EPS。"
                 "歷史毛利、營業費用、營業利益取 10-K；預測期模型只推事業別稅前利益與淨利，營業利益＝稅前利益 − 其他收入淨額（In 頁 N 節，金額很小），毛利＝營收 × 毛利率（PL 頁各情境區塊的「毛利率（參考）」），營業費用＝毛利 − 營業利益（推算）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：稅前利益取 DCF 頁情境 B 的 EBIT（淨利還原），毛利率與營業費用率沿用前一年據以算出營業成本、毛利與營業費用，"
                 "與延伸期營業利益（稅前利益 − 其他收入淨額）的差額列核對列，不改淨利與目標價（Andy 2026-11-14）。權益法投資損益在稅後（淨利 = 稅前利益 − 所得稅 + 權益法投資損益）。")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + YEARS + ["說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":   # 預測欄標題：依賴 Siliconomics 對帳中年份者加「（對帳中）」（旗標在 Coord 頁）
        yr = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr}"&IF(Coord!${"BCDE"[FLAGY[col]-2026]}${CR["flag"]}=1,"（對帳中）","")'
    r_hdr = r; r += 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, 13, note)
        L[key] = r; r += 1

    M = "#,##0.0"; PCT = "0.0%"
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    fc4 = lambda k: [f"=PL!{EC4[c]}{R[k]}" for c in "FGHI"]
    ext = lambda k: [f"=PL!{EC[c]}{X[k]}" for c in "JKL"]
    ra = r   # 營收列
    row("rev", "營收", "US$M", hh("rev") + fc4("rev") + ext("rev"), M, "FY2026E＝前三季實績＋4Q26 指引中點；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/Hist!C{h}-1".format(rev=L["rev"], h=HR["rev"])] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    rc = r   # 營業成本列：cogs=rc, gp=rc+1, gm=rc+2, opx=rc+3, oi=rc+4
    row("cogs", "營業成本（COGS）", "US$M", hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    ro = rc + 4
    row("gp", "毛利", "US$M", hh("gp") + [f"={c}{L['rev']}*{c}{rc + 2}" for c in "FGHIJKL"], M, "預測期＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{rc + 1}/{c}{L['rev']}" for c in "CDE"] + [f"=PL!{EC4[c]}{R['gm']}" for c in "FGHI"] + [f"={p}{rc + 2}" for p in "IJK"], PCT,
        "FY2026E–FY2029E 取 PL 各情境區塊「毛利率（參考；不進 EPS）」（FY2026 錨點＋斜率 × SoC 占營收變動）；FY2030E–FY2032E＝前一年")
    row("opx", "營業費用（含研發）", "US$M", hh("opx") + [f"={c}{rc + 1}-{c}{ro}" for c in "FGHI"] + [f"={c}{L['rev']}*{p}{rc + 3}/{p}{L['rev']}" for p, c in (("I", "J"), ("J", "K"), ("K", "L"))], M,
        "FY2026E–FY2029E＝毛利 − 營業利益（推算）；FY2030E–FY2032E＝營收 × 前一年營業費用率；歷史含 FY2024 出售事業利益 57.1（減少費用）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "US$M", hh("oi") + [f"={c}{r + 3}-oth_net" for c in "FGHIJKL"], M, "預測期＝稅前利益 − 其他收入淨額（In 頁 N 節 oth_net）")
    assert L["oi"] == ro
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）", "US$M", [f"={c}{r + 1}-{c}{L['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（利息收入 − 利息費用 − 其他費用）；預測期＝In 頁 oth_net")
    row("pti", "稅前利益（扣權益法前）", "US$M", hh("pti") + fc4("pti") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M,
        "FY2026E–FY2029E 取 PL 推導區塊（事業別稅前利益合計）；FY2030E–FY2032E 取 DCF 頁情境 B 的 EBIT（淨利還原，不含權益法損益）")
    row("tax", "所得稅", "US$M", hh("tax") + fc4("tax") + [f"={c}{L['pti']}*etr" for c in "JKL"], M, "FY2030E–FY2032E＝稅前利益 × 有效稅率（In 頁 etr）")
    row("eq", "權益法投資損益（Technoprobe；稅後）", "US$M", hh("eq") + fc4("aff") + ["=aff"] * 3, M, "損失為負；FY2026 起含少數股東權益（3Q26 起，Q 頁第 17 列註）")
    row("ni", "稅後淨利（歸屬泰瑞達）", "US$M", hh("ni") + fc4("ni") + ext("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數（百萬股）", "百萬股", [f"=Hist!{HC[c]}{HR['sh_d']}" for c in "CDE"] + fc4("sh_") + [f"={c}{L['ni']}/{c}{r + 1}" for c in "JKL"], M, "FY2030E–FY2032E＝淨利 ÷ EPS（延伸區塊）")
    row("eps", "EPS（GAAP 稀釋）", "US$", [f"=Hist!{HC[c]}{HR['eps_d']}" for c in "CDE"] + fc4("eps") + ext("eps"), "0.00", "GAAP 稀釋；共識 EPS 為 Non-GAAP（每年約高 0.3，Recon 頁）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/Hist!C{h}-1".format(ni=L["ni"], h=HR["ni"])] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$M", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2032E 為差額，見下）", "US$M", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None] * 3, "0.000")
    row("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利還原）；目標價用延伸期淨利，不受影響", "US$M",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JKL"], "0.000")
    row("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JKL"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 ＋ 權益法投資損益 − 稅後淨利", "US$M", [f"={c}{L['pti']}-{c}{L['tax']}+{c}{L['eq']}-{c}{L['ni']}" for c in COLS], "0.000",
        "FY2026E 差額 0.4＝上半年實績淨利 773.4 與（稅前 909.3 − 所得稅 128.9 − 權益法 6.6 ＝ 773.8）的季報捨入／少數股東差；錨點沿用 Q 頁，不改")
    row("c4", "核對：稅後淨利 − PL 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "US$M",
        [None] * 3 + [f"={c}{L['ni']}-PL!{EC4[c]}{R['ni']}" for c in "FGHI"] + [f"={c}{L['ni']}-PL!{EC[c]}{X['ni']}" for c in "JKL"], "0.000")
    row("c5", "核對：營收 − PL 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "US$M",
        [None] * 3 + [f"={c}{L['rev']}-PL!{EC4[c]}{R['rev']}" for c in "FGHI"] + [f"={c}{L['rev']}-PL!{EC[c]}{X['rev']}" for c in "JKL"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度＝日曆年，引擎對 FY(t) 直接取 CY(t) 的測試機市場與 Siliconomics 年別比值；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2032E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, rows_X=X, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 16, 17, 18   # DCF 頁：評價日後占比、折現年數（期中）、終值折現年數
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數"), (D.cell(fr, 1).value, D.cell(ty, 1).value, D.cell(tn, 1).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"   # 淨現金滾到評價日（含 stub 期 FCFF 與已宣告股利；三情境相同，取情境 B 欄）
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（第一個明確期年度，三情境接近）不動，營業利益率與營收倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("px", "現價（美元）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；US$M）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節；同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（百萬股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（US$M）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；US$M）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}
    NC = "BCDEFG"   # 本頁 FY2027–FY2032 欄；對應 DCF 頁 C–H 欄

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$M）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$M；事業別稅前利益合計）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$M）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（US$M）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（US$M，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$M）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（US$M）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2027 現值（US$M）", f"=B{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（US$M）", f"=SUM(C{Y['pv']}:G{Y['pv']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$M）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$M）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近下界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tn}+$B${A['nc']})/sh_v")
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
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(" 美元；要支持現價 "), TX("px", "#,##0"), T(" 美元，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 百萬美元，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(" 美元相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 泰瑞達過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給泰瑞達的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("歷史相對倍數 "), TX(coef, "0.00"), T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：估值並列四個目標價（版面規格 v1.3 G 節；只新增列） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價', exact=True)}"
    px_lab = f"Val!$A${lab(V, '現價（')}"
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.4.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "美元", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(" 美元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=" + px_lab); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    snap = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    tp_cell = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）', exact=True)}"
    t0 = lab(SE, "目標價參數敏感度"); hr = t0 + 1
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.4.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（ter/sens_ter.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "快照基準加權目標價（美元）"], 1):
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
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 ter/sens_ter.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR):
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]
    N = SIL.NAMES
    advf = os.path.basename(glob.glob(os.path.join(ROOT, "advantest", "dist", "*.xlsx"))[0])
    adv_ver = re.search(r"v\d+(?:\.\d+)+", advf).group(0)
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    rec = json.load(open(os.path.join(ROOT, "siliconomics", "data", "recon_snapshot_v0.19.json"), encoding="utf-8"))
    adv_rs = next(x for x in rec["markets"]["soc"]["rows"] if x["who"] == "Advantest")["vals"][1]   # CY2026 Advantest 份額（Siliconomics recon 快照）
    mk26 = N["SIF_TestSoCMkt_Base"][1]
    oth = SIL.V["rc_soc_oth"]
    P = wb["PL"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.4.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（讀 SIF_TestSoCMkt_Base、ts_tw_lo／ts_tw_hi、SIF_WTAnchor_Base；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（測試機），讀 Siliconomics"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁敏感度案例；不改基準"),
          ("SoC、記憶體測試機市場（CY，三情境）", "唯讀上游", f"advantest/dist/{advf}（基準情境市場與 Siliconomics SIF_TestSoCMkt_Base 核對一致，建置時斷言）", adv_ver, "PL 引擎：SoC 測試機營收＝市場 × 泰瑞達市占；記憶體同式"),
          ("泰瑞達 SoC 測試機市占（FY2026 錨點隱含）", "=s_soc", "本模型自算：Q 頁 FY2026 錨點 SoC 營收 ÷ CY2026 SoC 測試機市場；非 Siliconomics 輸入", "v0.4.5", "A、B 情境維持；C 情境與 Advantest 互補（同源，不作 C0 比較基準）"),
          ("Advantest SoC 市占 FY2026（互補基準）", "=adv_sh26", "Advantest 模型（Siliconomics recon 快照 markets.soc 同值）", adv_ver, "Advantest 自己的 CY2026 推得值；C 情境互補與 C0（實績）層級隱含值使用"),
          ("顆數情境係數（CY2028、CY2029；說明用）", "=um28_B", "經台積電模型與 Advantest 模型讀入 In 頁 um28_B／um29_B（Siliconomics SIF_UnitsCoef_Base 為基準）", SIL.VERSION, "情境 B 的市場已含係數；本模型不另乘"),
          ("測試對 WFE 比長期區間（下緣、上緣）", "=tw_lo", "Siliconomics ts_tw_lo／ts_tw_hi（SIL.row9 讀入 In 頁 tw_lo、tw_hi）", SIL.VERSION, "Recon 頁 D 節（同源，不判定）")]
    for y in (2027, 2028, 2029):
        rd.append((f"Wattonomics 錨定係數 CY{y}（敏感度案例用）", f"=wta_{y % 100}", "Siliconomics SIF_WTAnchor_Base（R4 比值：Wattonomics ÷ 本庫推得上線 GW）", SIL.VERSION, "Sens 頁「Wattonomics 錨定」案例；基準不乘"))
    rd.append(("對帳中旗標與另一側讀數比值（CY2026–CY2029）", "見本頁上方", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION, "兩讀數目標價（C4）"))
    ROW["rd0"] = r
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
        r += 1
    r += 1
    # ---- 2. 讀取層級核對（同源，不判定） ----
    CW.cell(r, 1, "2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0，無檢驗力，故不判定、不計入對帳中；C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）").font = BOLD; r += 1
    CW.cell(r, 1, "本模型會計年度＝日曆年，FY(t) 對 CY(t) 直接對應，無年度口徑差。下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。")
    ROW["caveat"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定", "層級名稱與版本／口徑／證據等級／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr2"] = r; r += 1
    first = r

    def line(label, unit, yr_, f_model, f_layer, fmt, note):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_model).number_format = fmt
        c = CW.cell(r, 5, f_layer); c.number_format = fmt
        if not (isinstance(f_layer, str) and f_layer.startswith("=")): c.font = BLUE
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
        return r - 1
    nb = f"Siliconomics {SIL.VERSION}"
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        line("SoC 測試機市場（情境 B；Advantest 模型讀入）vs Siliconomics 基準", "US$B", f"CY{y}", f"=ms_B_{y}", N["SIF_TestSoCMkt_Base"][i + 1], "#,##0.000",
             f"{nb} SIF_TestSoCMkt_Base（Equip 頁 SoC 測試機市場；CY2026 為 ts_mk_soc26，Interested-party：Advantest 自估）；Derived；模型經 Advantest 模型 PL 頁讀入，建置時斷言基準情境相同")
    line("測試對 WFE 比長期區間下緣", "%", "—", "=tw_lo", SIL.V["ts_tw_lo"]["base"], "0.0%", f"{nb} ts_tw_lo（公司 2Q26 法說；Interested-party）；SIL.row9 直接讀入 In 頁")
    line("測試對 WFE 比長期區間上緣", "%", "—", "=tw_hi", SIL.V["ts_tw_hi"]["base"], "0.0%", f"{nb} ts_tw_hi（同上）")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    last = r - 1
    ROW["first"], ROW["last"] = first, last
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：記憶體測試機市占（模型記憶體市占由 FY2026 錨點反推，層級沒有泰瑞達份額）；機器人、Product Test（層級不持有）；AI 加速器顆數（模型不直接讀顆數，由 Advantest 模型的 SoC 市場間接承接）。")
    r += 1
    # ---- 層級之間（參考） ----
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是泰瑞達自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for y in (2027, 2028, 2029):
        CW.cell(r, 1, "AI SoC 測試機市場（Advantest 模型情境 B）：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, f"=adv_aiB_{y % 100}").number_format = "#,##0.000"; CW.cell(r, 5, f"=adv_aiB_{y % 100}*{ratio(y)}").number_format = "#,##0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；對帳中年份依 C4 並列兩讀數目標價（本頁上方），個股不裁決")
        r += 1
    ROW["ref2"] = r - 1
    r += 2
    # ---- 3. 對帳中年份 ----
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t)：FY2027–FY2029 及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；FY2026 為前三季實績加 4Q26 指引的錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 美元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 美元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的 AI SoC 測試機市場 × 另一側比值；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的市場縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。本模型會計年度與日曆年相同，無 FY(t) 對 CY(t) 的月數近似。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    # ---- 4. C0（實績）----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.4.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（3Q26 10-Q 附註 E 的 SoC 測試機營收，2026 年 Q1–Q3＝2026-01～2026-09 累計；泰瑞達會計年度＝日曆年，無年度口徑差；不用指引或預測）；層級側＝Siliconomics 現行基準的 CY2026 SoC 測試機市場"
                  "× 泰瑞達隱含份額，隱含份額＝1 − Advantest SoC 市占（Advantest 模型 CY2026 推得；Siliconomics recon 快照同值）− 其他廠份額（Siliconomics 其他廠份額）。層級值是日曆年 CY2026：主讀數＝已公布期間累計實績 vs 層級 × 已公布月數 ÷ 12"
                  "（線性時差近似，Derived）；第二讀數＝最近一季 × 4 vs 層級；兩讀數的差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」。SoC 營收各季不單調（Q 頁：1Q26、2Q26、3Q26），兩讀數都是時差近似，不預設其一偏高或偏低，也不當成真值的上下界。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查：層級隱含值裡，市場值取 Advantest 公司自估（Siliconomics v0.9 起；依 Advantest FY2026 Q2 決算說明）、Advantest 份額由 Advantest 自己的實績與指引推得，兩者獨立於泰瑞達的季報；"
                  "其他廠份額（Analogy）基準取 CY2025 與 CY2026 兩個推得值之間，其中 CY2025 點用泰瑞達實績，屬部分同源（影響約隱含值的七分之一）。因此判定附註並列其他廠份額低、高值的參考讀數，結論不因同源而改判。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "公司實績中無可比層級值者：記憶體測試機營收（層級沒有泰瑞達或非 Advantest 的記憶體份額；模型記憶體市占由本公司錨點反推，同源）、半導體測試合計、Product Test、機器人與合併營收（層級不持有泰瑞達營收）：無可比實績。實績本身列於下表供對照。")
    ROW["none"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; r += 1
    CW.cell(ROW["hdr_in"], 9, "鍵名（追蹤用；HTML 不顯示）").font = BOLD
    inp = {}
    KEYS = {"mk": f"{'SIF_TestSoCMkt_Base'}（In 頁 ms_B_2026 連格；Siliconomics {SIL.VERSION} 值 {mk26}；建置時核對相同）", "adv": f"In 頁 adv_sh26／as_A_2026（連格）；Siliconomics recon 快照 markets.soc Advantest CY2026 {adv_rs:.6f}（D 欄與快照一致，建置時核對）",
            "oth": f"rc_soc_oth（D 欄為寫死值，讀自 Siliconomics {SIL.VERSION} 快照 {SIL.FILE}）", "oth_lo": "rc_soc_oth（下界）", "oth_hi": "rc_soc_oth（上界）"}
    assert abs(adv_rs - 0.713710144927536) < 1e-12 and abs(mk26 - 11.5) < 1e-12, (adv_rs, mk26)

    def put(key, label, unit, per, val, grade, src, fmt):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        if key in KEYS: CW.cell(r, 9, KEYS[key])
        inp[key] = r; r += 1
    put("m9", "SoC 測試機營收：2026 年 Q1–Q3 累計", "US$M", "2026-01～2026-09", 2401.289, "Verified", SRC_Q3 + "；九個月合計欄 SoC 2,401,289 千美元", "#,##0.000")
    put("q3", "SoC 測試機營收：3Q26", "US$M", "2026-07～2026-09", 676.509, "Verified", SRC_Q3 + "；三個月欄 SoC 676,509 千美元（Q 頁 I 欄取 676.0，來自法說摘要四捨五入）", "#,##0.000")
    put("h1", "SoC 測試機營收：上半年（1Q26＋2Q26）", "US$M", "2026-01～2026-06", "=Q!J4", "Verified", "Q 頁 J4（2Q26 10-Q 收入拆分上半年 1,724.8；連格）", "#,##0.000")
    put("q1", "SoC 測試機營收：1Q26（對照）", "US$M", "2026-01～2026-03", "=Q!G4", "Verified", "Q 頁 G4（1Q26 10-Q；連格）", "#,##0.0")
    put("q2", "SoC 測試機營收：2Q26（對照）", "US$M", "2026-04～2026-06", "=Q!H4", "Verified", "Q 頁 H4（2Q26 10-Q；連格）", "#,##0.0")
    put("xchk", "核對：九個月累計 − 3Q26 − 上半年（應約 0；10-Q 千美元對 Q 頁一位小數）", "US$M", "—", f"=D{r - 5}-D{r - 4}-D{r - 3}", "Derived", "九個月累計（3Q26 10-Q）減三個月（3Q26 10-Q）減上半年（Q 頁，2Q26 10-Q）；差額為四捨五入", "0.000")
    put("mo", "已公布月數", "月", "2026-01～2026-09", 9, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("mk", "Siliconomics 層級值：SoC 測試機市場 CY2026（基準）", "US$B", "CY2026", "=ms_B_2026", "Interested-party（層級值）",
        f"Siliconomics {SIL.VERSION} ts_mk_soc26（快照檔 {SIL.FILE}）：Advantest FY2026 Q2 決算說明（2026-10-28）CY2026 SoC 測試機市場約 115 億美元；本格連 In 頁 ms_B_2026（Advantest 模型讀入，與 SIF_TestSoCMkt_Base 相同）", "0.000")
    put("adv", "Advantest SoC 市占 CY2026（層級 recon 值；Advantest 模型）", "%", "CY2026", "=adv_sh26", "Derived（Advantest 模型）",
        f"Advantest 模型 {adv_ver} PL 頁（Siliconomics recon 快照 markets.soc：CY2026＝三季實績加第四季指引推得，與泰瑞達實績無關）；本格連 In 頁 adv_sh26", "0.0%")
    put("oth", "其他廠 SoC 測試機份額（Siliconomics rc_soc_oth，基準）", "%", "CY2026", oth["base"], "Analogy（層級值）",
        f"Siliconomics {SIL.VERSION} rc_soc_oth：CY2025 點＝1 − Advantest 自估 66% − 泰瑞達 27.4%；CY2026 點由兩家合併對帳推得其他約 1.9%；基準取兩點之間（部分同源，見上方同源檢查）；致茂、Cohu 等", "0.0%")
    put("oth_lo", "其他廠 SoC 測試機份額（下界）", "%", "CY2026", oth["lo"], "Analogy（層級值）", "同上（rc_soc_oth 低欄）；只作參考，不判定", "0.0%")
    put("oth_hi", "其他廠 SoC 測試機份額（上界）", "%", "CY2026", oth["hi"], "Analogy（層級值）", "同上（rc_soc_oth 高欄）；只作參考，不判定", "0.0%")
    put("sh", "層級隱含泰瑞達 SoC 份額（基準）", "%", "CY2026", f"=1-D{inp['adv']}-D{inp['oth']}", "Derived", "＝1 − Advantest 份額 − 其他廠份額；連到上列各格", "0.0%")
    put("lay", "層級隱含泰瑞達 SoC 測試機營收（基準）", "US$M", "CY2026", f"=D{inp['mk']}*1000*D{inp['sh']}", "Derived", "市場 × 隱含份額（美元十億 → 百萬）；連到上列各格", "#,##0.0")
    put("lay_lo", "層級隱含泰瑞達 SoC 測試機營收（其他廠＝下界）", "US$M", "CY2026", f"=D{inp['mk']}*1000*(1-D{inp['adv']}-D{inp['oth_lo']})", "Derived", "其他廠份額取低值（隱含值較高）", "#,##0.0")
    put("lay_hi", "層級隱含泰瑞達 SoC 測試機營收（其他廠＝上界）", "US$M", "CY2026", f"=D{inp['mk']}*1000*(1-D{inp['adv']}-D{inp['oth_hi']})", "Derived", "其他廠份額取高值（隱含值較低）", "#,##0.0")
    put("anc", "參考：模型 FY2026 錨點 SoC 測試機營收（上半年實績＋3Q26 實績＋4Q26 推估）", "US$M", "FY2026（＝CY2026）", "=f26_soc", "Derived（模型）", "Q 頁 FY2026 錨點（個股自身預測，非實績；殘差法：4Q26 營收指引中點減非 SoC 事業）；只作參考", "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge="（讀數）", unit="US$M"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("主讀數：已公布期間累計 SoC 測試機營收 vs 層級 × 月數 ÷ 12", "2026-01～09 vs CY2026 × 9/12", f"=D{inp['m9']}", f"=D{inp['lay']}*D{inp['mo']}/12",
             "實績＝3Q26 10-Q 九個月累計；層級＝Siliconomics 市場 ×（1 − Advantest 份額 − 其他廠份額）× 9 ÷ 12（線性時差近似）。泰瑞達 SoC 營收 1Q26 最高、3Q26 較低，累計值不等於全年的 9/12")
    rb = rd_("第二讀數：最近一季（3Q26）× 4 vs 層級", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp['lay']}",
             "最近一季年化；3Q26 是 2026 年內較低的一季，年化後不代表全年水準，只作時差近似的另一端參考")
    CW.cell(r, 1, "C0（實績）判定：SoC 測試機營收")
    CW.cell(r, 7, f'=IF(AND(ABS(F{ra})>0.15,ABS(F{rb})>0.15,F{ra}*F{rb}>0),"對帳中",IF(OR(ABS(F{ra})>0.15,ABS(F{rb})>0.15),"時差邊界（參考）","一致"))')
    CW.cell(r, 8, "兩讀數的差距都在 ±15% 內：一致，不開對帳單。層級隱含值部分同源（其他廠份額含泰瑞達實績校準，見同源檢查）：其他廠份額取上界時主讀數超過 ±15%、第二讀數仍在內（時差邊界），取下界時兩讀數都在內（見下方參考列）；"
                  "不論取哪一端都不會兩讀數同號超過，對帳中維持 0")
    ROW["judge"] = r; r += 1
    for key, nm in (("lay_lo", "其他廠＝下界"), ("lay_hi", "其他廠＝上界")):
        rd_(f"參考：主讀數（{nm}）", "2026-01～09 vs CY2026 × 9/12", f"=D{inp['m9']}", f"=D{inp[key]}*D{inp['mo']}/12", "其他廠份額 lo／hi 下的同一比較；不判定", "參考（不判定）")
        rd_(f"參考：第二讀數（{nm}）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp[key]}", "同上；不判定", "參考（不判定）")
    CW.cell(r, 1, "參考：已公布期間實績占模型 FY2026 錨點 SoC 營收（個股自身預測）"); CW.cell(r, 2, "US$M"); CW.cell(r, 3, "2026-01～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['m9']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（九個月實績 ÷ 模型全年錨點），不是差距；錨點是個股預測，非公司指引也非實績")
    r += 1
    CW.cell(r, 1, "參考：模型 FY2026 錨點 SoC 營收 vs 層級隱含 CY2026"); CW.cell(r, 2, "US$M"); CW.cell(r, 3, "FY2026 vs CY2026")
    CW.cell(r, 4, f"=D{inp['anc']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['lay']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    CW.cell(r, 7, "參考（個股自身預測，不判定）"); CW.cell(r, 8, "個股預測對層級隱含值；模型錨點含 4Q26 指引推估，差距不是實績檢驗；外部估計與個股預測只列參考")
    r += 1
    CW.cell(r, 1, "讀法：公司實績是 3Q26 10-Q 原文、獨立；層級隱含值由 Advantest 自估的市場、Advantest 自己的份額與 Siliconomics 的其他廠份額推得，泰瑞達的實績只透過其他廠份額的 CY2025 點進入一小部分。"
                  "所以這一列檢驗「Advantest 說的市場與份額，加上其他廠份額，留給泰瑞達的 SoC 空間，夠不夠容納泰瑞達已公布的營收」，結論為一致。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    r += 1
    # 無可比實績項目：實績列示
    CW.cell(r, 1, "其他公司實績（無可比層級值；2026 年 Q1–Q3 與 3Q26，3Q26 10-Q 附註 E）").font = BOLD; ROW["oth_t"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "2026 Q1–Q3 累計", "3Q26", "", "判定", "出處／口徑／缺少的揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_oth"] = r; r += 1
    ROW["oth0"] = r
    others = [("記憶體測試機營收", 715.603, 300.821, "無可比實績（層級無泰瑞達記憶體份額；模型記憶體市占同源）", "層級只有市場（ts_mk_mem26，Advantest 自估）；缺泰瑞達與其他廠的記憶體份額獨立揭露"),
              ("IST（整合系統測試）營收", 187.745, 94.681, "無可比實績（層級無對應市場）", "Siliconomics 沒有 IST 市場值"),
              ("半導體測試合計（SoC＋記憶體＋IST）", "=D{a}+D{b}+D{c}", "=E{a}+E{b}+E{c}", "無可比實績（層級無對應市場）", "新聞稿 3Q26 半導體測試 1,072；本列為三項相加"),
              ("Product Test 營收", 287.003, 138.111, "無可比實績（層級無對應市場）", "層級不持有"),
              ("機器人營收", 325.794, 95.828, "無可比實績（層級無對應市場）", "層級不持有"),
              ("合併營收", 3917.434, 1305.950, "無可比實績（層級無泰瑞達營收）", "層級不持有泰瑞達營收；實績與 3Q26 新聞稿合併營收 1,305.95、前三季 3,917.43 相同")]
    base_r = r
    for i, (nm, a9, a3, judge, miss) in enumerate(others):
        CW.cell(r, 1, nm); CW.cell(r, 2, "US$M"); CW.cell(r, 3, "2026-01～09／2026-07～09")
        if isinstance(a9, str):
            a9 = a9.format(a=base_r - 0, b=base_r + 1, c=base_r - 0)   # 先佔位，下面覆寫
        for col, v in ((4, a9), (5, a3)):
            c = CW.cell(r, col, v); c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        CW.cell(r, 7, judge); CW.cell(r, 8, miss + "；" + SRC_Q3 if i != 2 else miss)
        r += 1
    # 半導體測試合計：SoC（輸入表 m9／q3）＋記憶體（本表第一列）＋IST（本表第二列）
    rs = base_r + 2
    CW.cell(rs, 4, f"=D{inp['m9']}+D{base_r}+D{base_r + 1}").number_format = "#,##0.000"
    CW.cell(rs, 5, f"=D{inp['q3']}+E{base_r}+E{base_r + 1}").number_format = "#,##0.000"
    ROW["oth1"] = r - 1
    return ROW


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, HR, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; RV = wb["RevDCF"]; DC = wb["DCF"]; RC = wb["Recon"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    rb = PLA["rows_B"]
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_mk, s_sh, s_inc, s_wts = sens_r("SoC 測試機市場"), sens_r("泰瑞達 SoC 市占"), sens_r("測試事業增量利益率"), sens_r("情境權重")
    sd = SNS["dcfw"]   # Sens 頁目標價參數表「DCF 權重 0.3／0.7」列：C＝低案（DCF 權重 0.3）加權目標價
    r_e27 = lab(RC, "EPS（共識，non-GAAP）", lab(RC, "B. FY2027"))
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.4.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 美元，高於現價 "), TX(up, "0.0%"), T("，評等"), rec, T("：DCF 只有 "), TX(dcf, "#,##0"), T(" 美元，同業比較法 "), TX(cadj, "#,##0"), T(" 美元，各占一半。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：Comparable 採用泰瑞達過去十年相對同業的本益比（"), TX(coef, "0.00"), T(" 倍）與 FY2027 後測試機市場續增。不採相對倍數調整時加權目標價 "),
        TX(noadj, "#,##0"), T(" 美元（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{sd}", "#,##0"),
        T(" 美元（"), TX(f"Sens!C{sd}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{sd}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 美元，比現價 "), TX("px", "#,##0"), T(" 美元高 "), TX(up, "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.0"), T(" 美元）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 美元更高。"))
    put("o2", T("模型盈餘高於共識：FY2027 EPS 基準情境 "), TX(f"Recon!E{r_e27}", "0.00"), T(" 美元（GAAP），比共識 "), TX(f"Recon!B{r_e27}", "0.00"), T("（Non-GAAP）高 "), TX(f"Recon!F{r_e27}", "0.0%"),
        T("；測試機市場取自 Advantest 模型，FY2027 比 FY2026 高約四成，SoC＋記憶體測試機占晶圓前段設備 "), TX("Recon!E17", "0.0%"), T("，高於公司說的長期上緣。"))
    put("o3", T("最大變數是市場規模、市占與增量利益率（加權目標價相對基準）：SoC 測試機市場少一成半 "), TX(f"Sens!B{s_mk}", "+0.0%;-0.0%"), T("；SoC 市占 ±5 個百分點 "), TX(f"Sens!B{s_sh}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_sh}", "+0.0%;-0.0%"),
        T("；增量利益率 40%／56% "), TX(f"Sens!B{s_inc}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_inc}", "+0.0%;-0.0%"), T("；情境權重全給保守 "), TX(f"Sens!B{s_wts}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o4", T("DCF 偏低的原因：基準與保守情境 FY2028 起營收近乎持平、淨利率下滑，終值年權重大。WACC "), TX("wacc", "0.0%"), T("；要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"), T(" 倍。現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 百萬美元，模型 "), TX(rv("rv0"), "#,##0"), T("。"))
    put("o5", T("對帳中：CY2027–CY2029 的 AI SoC 測試機市場，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 美元（"), f"Coord!C{CR['rec']}", T("），與現行基準讀數（"), f"Coord!B{CR['rec']}", T("）評等不同。個股不裁決。公司實績（SoC 營收 2026 年 Q1–Q3）對層級隱含值："),
        f"Coord!G{CO['judge']}", T("。"))
    put("drv", T("分部營收推導（情境 B）：SoC 測試機＝SoC 測試機市場 × 泰瑞達市占；記憶體測試機＝記憶體測試機市場 × 市占；IST 隨 SoC 市場；Product Test 與機器人逐年成長；五項相加為合併營收。FY2028 SoC "),
        TX(f"PL!F{rb['rsoc']}", "#,##0"), T("、記憶體 "), TX(f"PL!F{rb['rmem']}", "#,##0"), T("、IST "), TX(f"PL!F{rb['rist']}", "#,##0"), T("、Product Test "), TX(f"PL!F{rb['rpt']}", "#,##0"),
        T("、機器人 "), TX(f"PL!F{rb['rrob']}", "#,##0"), T("，合計 "), TX(f"PL!F{rb['rev']}", "#,##0"), T(" 百萬美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、Ibiden）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2022–FY2025，公司 10-K）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，GAAP）、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數目標價；讀取層級；C0：SoC 測試機營收 2026 年 Q1–Q3 實績對層級隱含值，判定一致、對帳中 0）、"
                  "In 頁 N 節（其他收入淨額，只用於拆分稅前利益）、README 頁末 v1.3 白話列。既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, PB):
    # check_layout_v13 的兩讀數檢查要求 Coord 頁該列標籤恰為「加權目標價」（tools/coord 的 NAMES 為較長的標籤；本列只改文字標籤，數值與公式不變）
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    in_n(wb)
    CO = coord_c0(wb, SIL, CR, HR)
    PLA = pl_standard(wb, HR, CR, PB)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    SO.cell(rr, 1, "公司：3Q26 10-Q（2026-10-30 提交；附註 E 收入拆分：SoC 2026 Q1–Q3 累計 2,401,289 千美元、3Q26 676,509 千美元；C0 實績用）"); SO.cell(rr, 2, Q10)
    SO.cell(rr + 1, 1, "公司：3Q26 新聞稿（8-K Ex.99.1，2026-10-27）：合併營收 1,305.95、半導體測試 1,072、Product Test 138、機器人 96（C0 實績交叉核對）"); SO.cell(rr + 1, 2, "https://www.sec.gov/Archives/edgar/data/97210/000119312526434981/ter-20261027.htm")
    SO.cell(rr + 2, 1, "公司：FY2025 10-K 合併損益表（FY2023–FY2025；標準損益表歷史用）"); SO.cell(rr + 2, 2, K10)
    wb._v045 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v045_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 加權目標價相對基準 {SNS['t0']} 起；Coord 讀取層級與 C0 {CO['rd0']}–{CO['last']}；README v1.3 白話列 {len(RMA)} 列"
