# Advantest v0.7.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.7.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司決算短信；IFRS；會計年度 4–3 月）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（快照 advantest/sens.json）
#   Coord 頁末：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）；Sources、README 頁末
import glob, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.7.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
FN24 = "https://www.advantest.com/en/news/2025/ildr5p0000000il7-att/E_FR_FY2024_FN.pdf"
FN25 = "https://www.advantest.com/news/2026/a81o6o0000000hgw-att/E_FR_FY2025_FN.pdf"
FR1Q = "https://www.advantest.com/en/news/2026/qnpuno0000000cr7-att/E_FR_FY2026_1Q.pdf"
FR2Q = "https://www.advantest.com/en/news/2026/mpd9kn0000000gii-att/E_FR_FY2026_2Q.pdf"
IN1Q = "https://www.advantest.com/document/ja/investors/ir-library/result/JE_BIZ_260729_note.pdf"
IN2Q = "https://www.advantest.com/document/ja/investors/ir-library/result/JE_BIZ_261028_note.pdf"
SRC_H = (f"公司決算短信 Consolidated Statement of Profit or Loss（IFRS，3(2) 合併損益表）：FY2023、FY2024 取 E_FR_FY2024_FN.pdf 第 8 頁（EPS 與加權股數第 13 頁；網址見 Sources 頁），"
         f"FY2025 取 E_FR_FY2025_FN.pdf（網址見 Sources 頁）；百萬日圓 ÷ 1,000 換成十億日圓")
SRC_A = ("公司 Investor Note（決算說明）：FY2026 Q1 SoC 259.6、記憶體 48.0（JE_BIZ_260729_note.pdf 第 6 頁；網址見 Sources 頁）；Q2 SoC 353.3、記憶體 53.4（JE_BIZ_261028_note.pdf 第 6 頁；網址見 Sources 頁）。"
         "註：決算短信本身（E_FR_FY2026_1Q／2Q）只列測試系統事業分部，不分 SoC、記憶體金額，故這兩項取自 Investor Note；以決算短信分部數做合計核對")
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
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)；十億日圓（公司決算短信百萬日圓 ÷ 1,000）；FY2022 只有營收、淨利、EPS（In 頁 nim10 列 finboard.jp、Val 頁第 1 節；供 FY2023 年增率）
HIS = [
    ("rev", "營收（Net sales）", 560.191, 486.507, 779.707, 1128.610),
    ("cogs", "營業成本（Cost of sales）", None, 240.477, 334.622, 402.503),
    ("gp", "毛利（Gross profit）", None, 246.030, 445.085, 726.107),
    ("sga", "銷售管理費用（Selling, general and administrative；含研發）", None, 158.963, 195.392, 229.628),
    ("oi_in", "其他收益（Other income）", None, 3.926, 1.366, 3.552),
    ("oe", "其他費用（Other expenses）", None, 9.365, 22.898, 0.911),
    ("oi", "營業利益（Operating income）", None, 81.628, 228.161, 499.120),
    ("fi", "財務收益（Financial income）", None, 1.244, 1.895, 20.354),
    ("fe", "財務費用（Financial expenses）", None, 4.702, 5.282, 2.754),
    ("pti", "稅前利益（Income before income taxes）", None, 78.170, 224.774, 516.720),
    ("tax", "所得稅（Income taxes）", None, 15.880, 63.597, 141.367),
    ("ni", "稅後淨利（Net income；全數歸屬母公司，無非控制權益）", 130.400, 62.290, 161.177, 375.353),
]
EPSR = [("eps_b", "EPS 基本（公司公告值）", (174.35, 84.45, 218.67, 515.15)),
        ("eps_d", "EPS 稀釋（公司公告值）", (None, 84.16, 218.01, 513.30)),
        ("sh_b", "基本加權平均股數（百萬股）", (None, 737.560501, 737.064308, 728.634777)),
        ("sh_d", "稀釋加權平均股數（百萬股）", (None, 740.120997, 739.310748, 731.260642))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.7.5 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度 4–3 月：FY2023＝2023-04～2024-03、FY2024＝2024-04～2025-03、FY2025＝2025-04～2026-03。金額：十億日圓；EPS：日圓。"
               "IFRS。FY2023–FY2025 取公司決算短信合併損益表；FY2022 只取營收、淨利、EPS（In 頁 nim10 列、Val 頁第 1 節），供 FY2023 年增率")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H + ("；FY2022 營收、淨利取 In 頁 nim10 列（finboard.jp；第二來源 the-shashi.com 營收相同）" if vals[0] is not None else ""))
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H + ("；FY2022 基本 EPS 取 Val 頁第 1 節（irbank，分割調整後）" if vals[0] is not None else ""))
        R[k] = r; r += 1
    H.cell(r, 1, "h_oth"); H.cell(r, 2, "其他收入淨額（＝財務收益 − 財務費用；＝稅前利益 − 營業利益）")
    for c in "DEF":
        H[f"{c}{r}"] = f"={c}{R['fi']}-{c}{R['fe']}"; H[f"{c}{r}"].number_format = "#,##0.000"
    H.cell(r, 7, "Derived"); H.cell(r, 8, "損益表各列加減；與稅前利益 − 營業利益相同（核對列）")
    R["oth"] = r; r += 1
    H.cell(r, 1, "h_opx"); H.cell(r, 2, "營業費用（＝毛利 − 營業利益；＝銷管費用 − 其他收益 ＋ 其他費用；Advantest 模型 opex25 同口徑）")
    for c in "DEF":
        H[f"{c}{r}"] = f"={c}{R['gp']}-{c}{R['oi']}"; H[f"{c}{r}"].number_format = "#,##0.000"
    H.cell(r, 7, "Derived"); H.cell(r, 8, "IFRS 損益表沒有營業費用合計列；本列＝毛利 − 營業利益（含其他營業收支淨額）")
    R["opx"] = r; r += 2
    H.cell(r, 1, "核對（單位：十億日圓；決算短信以百萬日圓列示，加總差為 0）").font = BOLD; r += 1
    chk = [("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"),
           ("毛利 − 銷管費用 ＋ 其他收益 − 其他費用 − 營業利益", "={c}{gp}-{c}{sga}+{c}{oi_in}-{c}{oe}-{c}{oi}"),
           ("營業利益 ＋ 其他收入淨額 − 稅前利益", "={c}{oi}+{c}{oth}-{c}{pti}"),
           ("稅前利益 − 所得稅 − 稅後淨利", "={c}{pti}-{c}{tax}-{c}{ni}"),
           ("稅後淨利 ÷ 基本加權平均股數 − 基本 EPS（日圓；公司 EPS 四捨五入至分，差 ≤ 0.005）", "={c}{ni}*1000/{c}{sh_b}-{c}{eps_b}")]
    R["chk0"] = r
    for labl, f_ in chk:
        H.cell(r, 2, labl)
        for c in "DEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "核對：FY2025 營收（本頁，決算短信）− Q 頁 FY2025 營收（Q 頁一位小數；差 ≤ 0.05）"); H[f"F{r}"] = f"=F{R['rev']}-a25_rev"; H[f"F{r}"].number_format = "0.000"; R["chk1"] = r; r += 1
    H.cell(r, 2, "核對：FY2025 稅後淨利（本頁）− Q 頁 FY2025 淨利（Q 頁一位小數）"); H[f"F{r}"] = f"=F{R['ni']}-a25_ni"; H[f"F{r}"].number_format = "0.000"; R["chk2"] = r; r += 1
    H.cell(r, 2, "核對：FY2025 毛利（本頁，決算短信 726.107）− Q 頁 FY2025 毛利（Q 頁＝營收 × 毛利率 64.3%，四捨五入；差約 0.4，不改既有格）"); H[f"F{r}"] = f"=F{R['gp']}-a25_gp"; H[f"F{r}"].number_format = "0.000"; R["chk3"] = r; r += 1
    H.cell(r, 2, "核對：基本 EPS（本頁）− Val 頁第 1 節各會計年度 EPS（FY2022＝2023-03 期 B13、FY2023＝B14、FY2024＝B15、FY2025＝B16；單位日圓）")
    for c, vr in zip("CDEF", (13, 14, 15, 16)):
        H[f"{c}{r}"] = f"={c}{R['eps_b']}-Val!B{vr}"; H[f"{c}{r}"].number_format = "0.00"
    H.cell(r, 8, "FY2024 差 −7.57（本頁 218.67 − Val 226.24）為已知未解差異：Val 頁（README v0.7 以來）寫 226.24 取自 FY2024 決算書，但 FY2024 與 FY2025 兩份決算短信的 FY2024 基本 EPS 都是 218.67，"
                 "且 161,177 ÷ 737.064 百萬股＝218.67；本版不改 Val 既有格（不變性），標示待主控裁決（影響 Val 頁 2024-12 年底 P/E 與十年分位）")
    R["chk4"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 70
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "D", "D": "E", "E": "F"}
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}   # 預測欄依賴的日曆年旗標（FY(t)＝3/4 CY(t)＋1/4 CY(t+1)；FY2030–32 沿用 CY2029）
PC4 = {"F": "D", "G": "E", "H": "F", "I": "G"}   # 標準損益表 FY2026–29 → PL 頁欄
EC = {"J": "H", "K": "I", "L": "J"}   # FY2030–32 → PL 延伸區塊欄
DCOL = {"F": "C", "G": "D", "H": "E", "I": "F", "J": "G", "K": "H", "L": "I"}   # → DCF 頁欄（FY2026＝C … FY2032＝I）


def pl_standard(wb, HR, CR, PB):
    P = wb["PL"]; R = PB["B"]
    xb = lab(P, "情境（延伸）")
    xs = lab(P, "情境 B（基準）", xb, exact=True)
    X = dict(rev=lab(P, "營收", xs, exact=True), ni=lab(P, "淨利（收斂法延伸）", xs), eps=lab(P, "EPS（延伸）", xs))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, "標準損益表（v0.7.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準）。會計年度 4–3 月；十億日圓；EPS 日圓。只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：IFRS（國際財務報導準則；公司決算短信，與美國 GAAP 分類略異；Non-GAAP 不適用，Advantest 不另公布）。EPS 為基本 EPS（與模型、公司指引同口徑；稀釋 EPS 見 Hist 頁）。"
                 "預測期模型只推毛利、營業費用（固定＋變動）、營業利益與淨利（淨利＝營業利益 × 淨利÷營業利益比）；稅前利益＝淨利 ÷（1 − WACC 用稅率），所得稅＝稅前利益 − 淨利，其他收入淨額（財務收益 − 財務費用）＝稅前利益 − 營業利益，三者皆為由模型淨利反推的 Derived 列，不進 EPS、現金流與目標價。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷ 淨利÷營業利益比），毛利率沿用 FY2029，營業費用＝毛利 − 營業利益（推算）。FY2026E 為模型值（CY2026 實績加指引錨定的 SoC 營收；非公司指引）。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":
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
    fc = lambda k: [f"=PL!{PC4[c]}{R[k]}" for c in "FGHI"]
    row("rev", "營收", "¥B", hh("rev") + fc("rev") + [f"=PL!{EC[c]}{X['rev']}" for c in "JKL"], M, "FY2026E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/Hist!C{h}-1".format(rev=L["rev"], h=HR["rev"])] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    rc = r
    row("cogs", "營業成本（COGS）", "¥B", hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "¥B", hh("gp") + [f"=PL!{PC4[c]}{R['gp']}" for c in "FGHI"] + [f"={c}{L['rev']}*{c}{rc + 2}" for c in "JKL"], M, "FY2026E–FY2029E 取 PL 情境 B；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{rc + 1}/{c}{L['rev']}" for c in "CDEFGHI"] + [f"={p}{rc + 2}" for p in "IJK"], PCT, "FY2030E–FY2032E 沿用前一年")
    row("opx", "營業費用（含研發與其他營業收支淨額）", "¥B", hh("opx") + [f"=PL!{PC4[c]}{R['opx']}" for c in "FGHI"] + [f"={c}{rc + 1}-{c}{rc + 4}" for c in "JKL"], M, "歷史＝毛利 − 營業利益；FY2026E–FY2029E 為模型固定＋變動；FY2030E 起＝毛利 − 營業利益（推算）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "¥B", hh("oi") + [f"=PL!{PC4[c]}{R['oi']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['ni']}/nioi" for c in "JKL"], M, "FY2030E 起＝淨利 ÷ nioi（同 DCF 頁 EBIT）")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）", "¥B", [f"={c}{r + 1}-{c}{L['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（財務收益 − 財務費用）；預測期為由淨利反推的 Derived 值")
    row("pti", "稅前利益", "¥B", hh("pti") + [f"={c}{r + 2}/(1-tax_w)" for c in "FGHIJKL"], M, "預測期＝淨利 ÷（1 − tax_w）")
    row("tax", "所得稅", "¥B", hh("tax") + [f"={c}{L['pti']}-{c}{r + 1}" for c in "FGHIJKL"], M, "預測期＝稅前利益 − 淨利")
    row("ni", "稅後淨利（歸屬母公司）", "¥B", hh("ni") + [f"=PL!{PC4[c]}{R['ni']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['ni']}" for c in "JKL"], M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "基本股數（百萬股）", "百萬股", [f"=Hist!{HC[c]}{HR['sh_b']}" for c in "CDE"] + ["=shares"] * 4 + [f"={c}{L['ni']}*1000/{c}{r + 1}" for c in "JKL"], M, "FY2026E–FY2029E＝In 頁 shares（指引淨利 ÷ 指引 EPS）；FY2030E 起＝淨利 ÷ EPS（延伸區塊）")
    row("eps", "EPS（IFRS 基本）", "¥", [f"=Hist!{HC[c]}{HR['eps_b']}" for c in "CDE"] + [f"=PL!{PC4[c]}{R['eps']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['eps']}" for c in "JKL"], "#,##0.00", "基本 EPS；共識與公司指引為基本 EPS")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_tax", "有效稅率（所得稅 ÷ 稅前利益）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT, "預測期＝In 頁 tax_w")
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/Hist!C{h}-1".format(ni=L["ni"], h=HR["ni"])] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；歷史欄為決算短信捨入）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "¥B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "¥B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "¥B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：營業利益 − DCF 頁情境 B 的 EBIT（FY2026E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['oi']}-DCF!{DCOL[c]}{d_ebit}" for c in "FGHIJKL"], "0.000")
    row("c5", "核對：稅後淨利 − PL 情境 B（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['ni']}-PL!{PC4[c]}{R['ni']}" for c in "FGHI"] + [f"={c}{L['ni']}-PL!{EC[c]}{X['ni']}" for c in "JKL"], "0.000")
    row("c6", "核對：營收 − PL 情境 B（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['rev']}-PL!{PC4[c]}{R['rev']}" for c in "FGHI"] + [f"={c}{L['rev']}-PL!{EC[c]}{X['rev']}" for c in "JKL"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度 4–3 月，FY(t)＝3/4×CY(t)＋1/4×CY(t+1)；標題「（對帳中）」依 Coord 頁各日曆年旗標（FY2026 含 1/4 CY2027，但該部分不乘比值，故不標；FY2027–FY2029 依賴 CY2027–CY2029；FY2030E–FY2032E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標）。")
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
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.7.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2026（只計評價日後 0.25 年）與 FY2027 不動，營業利益率與營收倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("nc", "淨現金滾到評價日（情境 B；¥B）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節；同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.0000")
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
    NC = "BCDEFGH"   # 本頁 FY2026–FY2032 欄；對應 DCF 頁 C–I 欄

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGHI"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（¥B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（¥B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "I" else "=0"), "0.0000")
    yr("k", "每 1 日圓營業利益對 EV 的貢獻（FY2026、FY2027 不動，記 0）", lambda dc, c: ("=0" if dc in "CD" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
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
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2026–FY2027 現值（¥B）", f"=B{Y['pv']}+C{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（¥B）", f"=SUM(D{Y['pv']}:H{Y['pv']})+DCF!I{nop}*H{Y['tvf']}", "#,##0.0")
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
                f"+DCF!$I${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tn}+$B${A['nc']})/sh_v")
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
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × Advantest 過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給Advantest的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("歷史相對倍數 "), TX(coef, "0.00"), T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFGH": RV.column_dimensions[c].width = 14
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
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.7.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
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
    SE.cell(r, 4, "v0.7.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（advantest/sens_advantest.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
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
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 advantest/sens_advantest.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))



# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR, PB):
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]; N = SIL.NAMES
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    rec = json.load(open(os.path.join(ROOT, "siliconomics", "data", "recon_snapshot_v0.19.json"), encoding="utf-8"))
    adv_rs = next(x for x in rec["markets"]["soc"]["rows"] if x["who"] == "Advantest")["vals"][1]
    mk26 = N["SIF_TestSoCMkt_Base"][1]
    assert abs(adv_rs - 0.713710144927536) < 1e-12 and abs(mk26 - 11.5) < 1e-12, (adv_rs, mk26)
    RB = PB["B"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.7.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（讀 ts_mk_soc25／26、ts_mk_mem25／26、ts_nonai26、ts_g_nonai、ts_eff_par、ts_tt_*、ts_w_hbm、SIF_UnitsCoef_Base、SIF_WTAnchor_Base；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（測試機），讀 Siliconomics"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁敏感度案例；不改基準"),
          ("CY2026 SoC 測試機市場", "=mk_soc26", "Siliconomics ts_mk_soc26（Advantest 公司自估，Interested-party；In 頁 mk_soc26 直接讀入）", SIL.VERSION, "校準錨：每加權顆測試支出＝（市場 − 非 AI）÷ 加權顆數"),
          ("Advantest SoC 市占（CY2026 推得）", "=sh_soc", "本模型自算（In 頁 sh_soc＝CY2026 三季實績加第四季指引一半 ÷ 匯率 ÷ 市場）；Siliconomics recon 快照 markets.soc 讀回同值", "v0.7.5", "方向是個股 → 層級（同源）；C0 因此標參考"),
          ("顆數情境係數（CY2028、CY2029；說明用）", "=um28_B", "經台積電模型讀入 In 頁 um28_B／um29_B（Siliconomics SIF_UnitsCoef_Base 為基準）", SIL.VERSION, "情境 B 的市場已含係數；本模型不另乘")]
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
    CW.cell(r, 1, "2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0，無檢驗力，故不判定、不計入對帳中；C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）").font = BOLD; r += 1
    CW.cell(r, 1, "本模型會計年度 4–3 月；下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。日曆年市場值與 FY 的換算在 PL 頁（FY(t)＝3/4×CY(t)＋1/4×CY(t+1)）。")
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
    nb = f"Siliconomics {SIL.VERSION}"
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        line("SoC 測試機市場（情境 B；本模型 PL 頁）vs Siliconomics 基準", "US$B", f"CY{y}", f"=PL!{'CDEF'[i]}{RB['soc']}", N["SIF_TestSoCMkt_Base"][i + 1], "#,##0.000",
             f"{nb} SIF_TestSoCMkt_Base；Derived；本模型情境 B 的 SoC 測試機市場（AI＋非 AI；情境 B 顆數係數＝Siliconomics 基準），建置時斷言相同")
    for k, lbl, u, f_ in (("mk_soc25", "SoC 測試機市場", "US$B", "#,##0.000"), ("mk_mem25", "記憶體測試機市場", "US$B", "#,##0.000"), ("mk_mem26", "記憶體測試機市場", "US$B", "#,##0.000"),
                          ("nonai26", "非 AI SoC 測試機市場", "US$B", "#,##0.000"), ("g_nonai", "非 AI SoC 市場年增", "%", "0.0%"), ("w_hbm", "記憶體測試機需求中隨 HBM 位元成長的比重", "%", "0.0%")):
        yy = {"mk_soc25": "CY2025", "mk_mem25": "CY2025", "mk_mem26": "CY2026", "nonai26": "CY2026", "g_nonai": "—", "w_hbm": "—"}[k]
        line(lbl, u, yy, f"={k}", SIL.V["ts_" + k]["base"], f_, f"{nb} ts_{k}（{SIL.V['ts_' + k]['tag']}）；In 頁直接讀入")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    last = r - 1
    ROW["first"], ROW["last"] = first, last
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：Advantest 記憶體測試機市占（層級只有記憶體市場，沒有 Advantest 份額；模型市占由 FY2025 實績反推）；其他系統、服務與其他（層級不持有）；加速器顆數（讀自市場層與台積電模型，非本模型輸入）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是 Advantest 自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for i, y in enumerate((2027, 2028, 2029)):
        CW.cell(r, 1, "AI SoC 測試機市場（情境 B）：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, f"=PL!{'DEF'[i]}{RB['ai']}").number_format = "#,##0.000"; CW.cell(r, 5, f"=D{r}*{ratio(y)}").number_format = "#,##0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；對帳中年份依 C4 並列兩讀數目標價（本頁上方），個股不裁決")
        r += 1
    ROW["ref2"] = r - 1
    r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝3/4×CY(t)＋1/4×CY(t+1)：FY2027–FY2029 及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；FY2026 的 1/4 CY2027 部分不乘比值、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 日圓（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 日圓（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的加權顆數 × 另一側比值；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數縮放加速器加權顆數（進 AI SoC 測試機市場）；記憶體測試機市場（HBM 位元、DRAM 設備支出）未縮放，另一側讀數可能偏高。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    # ---- 4. C0 ----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.7.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績：FY2026 Q1–Q2（2026-04～2026-09；會計年度 4 月起）SoC 測試機營收（Investor Note 第 6 頁；決算短信只列測試系統事業分部，故以分部數核對合計）；不用指引或預測。層級側＝Siliconomics 現行基準 CY2026 SoC 測試機市場 × Advantest 市占 × 匯率。"
                  "期間對應：實績 2026-04～09（6 個月）對層級 CY2026 × 6 ÷ 12（線性時差近似，Derived）＝主讀數；第二讀數＝最近一季（Q2，2026-07～09）× 4 對層級 CY2026（季中點約 8 月中旬，晚於日曆年中點，SoC 逐季上升時偏高，不得寫成「夾住真值」）。"
                  "兩讀數差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」；層級隱含值任一部分與個股實績同源時只標「參考（同源校準）」、不判定（tools/coord/README）。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（全部消去）：Advantest 市占 sh_soc＝（2026-01～03 實績＋FY2026 Q1＋Q2 實績＋（FY2026 指引 SoC − Q1 − Q2）÷ 2）÷（匯率 × CY2026 市場），即由 Advantest 自己的實績與指引除以市場得出（In 頁 sh_soc；Siliconomics recon 快照 markets.soc 讀回同值）。"
                  "層級隱含 ¥ ＝ 市場 × 匯率 × 市占，市場值（11.5）與匯率（150）互相消去，剩下 ＝ 2026-01～03 實績 ＋ 2026-04～09 實績 ＋ 2026-10～12 以下半年指引一半估計。"
                  "所以本表實際檢驗的只是「2026-04～09 實績 ÷ 該隱含值是否≈ 6/12」，即 Advantest 自己的季別分布與下半年指引分配假設，不是對 Siliconomics 層級值的獨立檢驗；依規則標「參考（同源校準）」，結論「無可比實績（層級值同源）」，對帳中維持 0。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "缺的獨立揭露：(1) 非 Advantest 自估的 CY2026 SoC 測試機市場值（11.5 是 Advantest 自估，Interested-party；第三方只有 SEMI 測試設備合計，口徑含探針台與分選機）；(2) 其他廠（泰瑞達、致茂等）CY2026 SoC 營收的獨立彙總（泰瑞達見 ter 模型 C0，其他廠份額 Siliconomics rc_soc_oth 為 Analogy 且由兩家反推）。"
                  "記憶體測試機：層級有市場值（ts_mk_mem26＝Advantest 自估）但沒有 Advantest 記憶體份額，無可比實績；其他系統、服務與其他、合併營收：層級不持有，無可比實績。實績本身列於下表供對照。")
    ROW["none"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; r += 1
    CW.cell(ROW["hdr_in"], 9, "鍵名（追蹤用；HTML 不顯示）").font = BOLD
    inp = {}

    def put(key, label, unit, per, val, grade, src, fmt, keyname=None):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        if keyname: CW.cell(r, 9, keyname)
        inp[key] = r; r += 1
    put("q1", "SoC 測試機營收：FY2026 Q1", "¥B", "2026-04～2026-06", "=Q!G4", "Verified", "Q 頁 G4（連格）；原文 JE_BIZ_260729_note.pdf 第 6 頁「SoC Test System sales were 259.6 billion JPY」（網址見 Sources 頁）", "#,##0.0")
    put("q2", "SoC 測試機營收：FY2026 Q2（最近一季）", "¥B", "2026-07～2026-09", "=Q!H4", "Verified", "Q 頁 H4（連格）；原文 JE_BIZ_261028_note.pdf 第 6 頁「SoC Test System sales were 353.3 billion JPY」（網址見 Sources 頁）", "#,##0.0")
    put("h1", "SoC 測試機營收：FY2026 Q1–Q2 累計", "¥B", "2026-04～2026-09", f"=D{r - 2}+D{r - 1}", "Verified（兩季相加）", "Q1＋Q2；決算短信 E_FR_FY2026_2Q 未列 SoC 金額（只說明 SoC 銷售增加），見下方測試系統事業分部核對", "#,##0.0")
    put("mem1", "記憶體測試機營收：FY2026 Q1–Q2 累計", "¥B", "2026-04～2026-09", "=Q!G5+Q!H5", "Verified", "Q 頁 G5＋H5；Q1 48.0（JE_BIZ_260729_note 第 6 頁）、Q2 53.4（JE_BIZ_261028_note 第 6 頁）", "#,##0.0")
    put("ts_h1", "測試系統事業營收（決算短信分部）：FY2026 上半年", "¥B", "2026-04～2026-09", 770.231, "Verified", "決算短信 E_FR_FY2026_2Q.pdf 第 3 頁分部表（770,231 百萬日圓；網址見 Sources 頁）", "#,##0.000")
    put("ts_q1", "測試系統事業營收（決算短信分部）：FY2026 Q1", "¥B", "2026-04～2026-06", 333.562, "Verified", "決算短信 E_FR_FY2026_1Q.pdf 第 3、11 頁（333,562 百萬日圓；網址見 Sources 頁）", "#,##0.000")
    put("xchk", "核對：Q 頁 SoC＋記憶體＋其他系統（Q1＋Q2）− 決算短信測試系統事業上半年（應約 0；Q 頁一位小數）", "¥B", "2026-04～2026-09", "=Q!G4+Q!H4+Q!G5+Q!H5+Q!G6+Q!H6-D" + str(r - 2), "Derived", "SoC、記憶體取自 Investor Note，其他系統為殘差（Q1）與公布值（Q2）；合計對上決算短信分部數，差額為四捨五入", "0.000")
    put("mo", "已公布月數", "月", "2026-04～2026-09", 6, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("mk", "Siliconomics 層級值：SoC 測試機市場 CY2026（基準）", "US$B", "CY2026", "=mk_soc26", "Interested-party（層級值）",
        f"Siliconomics {SIL.VERSION} ts_mk_soc26（快照 {SIL.FILE}；值 {mk26}，建置時核對）：Advantest FY2026 Q2 決算說明（2026-10-28）CY2026 SoC 測試機市場約 115 億美元（公司自估）；本格連 In 頁 mk_soc26", "0.000")
    put("fx", "匯率（公司財測前提）", "¥/US$", "FY2026", "=fx", "Interested-party", "In 頁 fx（公司第二季起假設 ¥150；Siliconomics 不共用日圓匯率）", "0")
    put("adv", "Advantest SoC 市占 CY2026（層級 recon 值；Advantest 模型自算）", "%", "CY2026", "=sh_soc", "Derived（Advantest 模型；同源）",
        f"In 頁 sh_soc；Siliconomics recon 快照 markets.soc Advantest FY2026 起值 {adv_rs:.6f}（建置時核對相同）；＝（2026-01～03 實績 237.2＋Q1＋Q2＋（指引 SoC − Q1 − Q2）÷ 2）÷（匯率 × 市場）", "0.0%")
    put("lay", "層級隱含 Advantest SoC 測試機營收（基準）", "¥B", "CY2026", f"=D{inp['mk']}*D{inp['fx']}*D{inp['adv']}", "Derived", "市場（US$B）× 匯率 × 市占；連到上列各格", "#,##0.0")
    put("rec", "核對：層級隱含 − （2026-01～03 實績＋Q1＋Q2＋（指引 − Q1 − Q2）÷ 2）（應為 0：市場與匯率消去）", "¥B", "CY2026", f"=D{inp['lay']}-(Q!E4+q1_soc+q2_soc+(g_soc-q1_soc-q2_soc)/2)", "Derived", "證明層級隱含值＝Advantest 自己的實績加指引（同源）", "0.000")
    put("anc", "參考：模型 FY2026 SoC 測試機營收（情境 B；個股預測）", "¥B", "FY2026（2026-04～2027-03）", f"=PL!D{RB['rsoc']}", "Derived（模型）", "PL 頁情境 B FY2026；個股自身預測，非實績；只作參考", "#,##0.0")
    put("gsoc", "參考：公司指引 SoC 測試機 FY2026（2026-10-28 再上修）", "¥B", "FY2026", "=g_soc", "Interested-party", "In 頁 g_soc＝1,375.0（Q 頁第 24 列；公司指引，不算實績）", "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge="（讀數）"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "¥B"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("主讀數：已公布期間累計 SoC 測試機營收 vs 層級 × 月數 ÷ 12", "2026-04～09 vs CY2026 × 6/12", f"=D{inp['h1']}", f"=D{inp['lay']}*D{inp['mo']}/12",
             "實績＝FY2026 Q1＋Q2；層級＝Siliconomics 市場 × 市占 × 匯率 × 6 ÷ 12（線性時差近似）。實績期間不含 2026-01～03 與 10～12 月，SoC 逐季上升，時差近似的偏誤方向取決於季別分布，不預設")
    rb = rd_("第二讀數：最近一季（FY2026 Q2）× 4 vs 層級", "2026-07～09 × 4 vs CY2026", f"=D{inp['q2']}*4", f"=D{inp['lay']}",
             "最近一季年化；季中點（8 月中旬）晚於日曆年中點，成長期偏高；只作時差近似的另一端參考")
    CW.cell(r, 1, "C0（實績）判定：SoC 測試機營收"); CW.cell(r, 7, "參考（同源校準）")
    CW.cell(r, 8, "結論：無可比實績（層級值同源）：層級隱含值＝Advantest 自己的實績加指引（市場與匯率消去），缺非 Advantest 自估的市場值與其他廠獨立營收揭露。數字參考：若不論同源，兩讀數差距見上列，都在 ±15% 內會是「一致」；對帳中維持 0")
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占模型 FY2026 SoC 營收（個股自身預測）"); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['h1']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（上半年實績 ÷ 模型全年），不是差距；模型值是個股預測，非公司指引也非實績")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占公司指引 SoC 測試機 FY2026"); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09 vs FY2026 指引")
    CW.cell(r, 4, f"=D{inp['h1']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['gsoc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "公司指引只列參考（不算實績）；比值是進度")
    r += 1
    CW.cell(r, 1, "讀法：公司實績是 Investor Note 原文、獨立；層級隱含值由 Advantest 自估的市場、Advantest 自己的實績與指引推得，市場與匯率消去，所以這一列檢驗的只是 Advantest 自己的季別分布，結論：無可比實績（層級值同源），數字僅供參考。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    r += 1
    CW.cell(r, 1, "其他公司實績（無可比層級值；FY2026 Q1–Q2 與 Q2）").font = BOLD; ROW["oth_t"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "FY2026 Q1–Q2 累計", "FY2026 Q2", "", "判定", "出處／口徑／缺少的揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_oth"] = r; r += 1
    ROW["oth0"] = r
    others = [("記憶體測試機營收", f"=D{inp['mem1']}", "=Q!H5", "無可比實績（層級無 Advantest 記憶體份額）", "層級有市場（ts_mk_mem26，Advantest 自估），缺 Advantest 記憶體測試機份額的獨立揭露；模型市占由 FY2025 實績反推（同源）。" + SRC_A),
              ("測試系統事業營收（決算短信分部）", f"=D{inp['ts_h1']}", f"=D{inp['ts_h1']}-D{inp['ts_q1']}", "無可比實績（層級無對應量）", "決算短信分部表（Q2 ＝ 上半年 − Q1）；層級不持有分部營收"),
              ("服務、支援與其他營收", "=Q!G7+Q!H7", "=Q!H7", "無可比實績（層級無對應量）", "Q 頁第 7 列（Q1 決算短信 33.911、Q2 為合併營收 − 分部）；層級不持有"),
              ("合併營收", "=Q!G9+Q!H9", "=Q!H9", "無可比實績（層級無 Advantest 營收）", "Q 頁第 9 列；決算短信上半年 845,200 百萬日圓（E_FR_FY2026_2Q.pdf 第 1 頁）；層級不持有 Advantest 營收")]
    for nm, a9, a3, judge, miss in others:
        CW.cell(r, 1, nm); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09／2026-07～09")
        for col, v in ((4, a9), (5, a3)):
            CW.cell(r, col, v).number_format = "#,##0.0"
        CW.cell(r, 7, judge); CW.cell(r, 8, miss)
        r += 1
    ROW["oth1"] = r - 1
    return ROW


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, HR, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_mk, s_sh, s_fx, s_wts = sens_r("SoC 測試機市場"), sens_r("Advantest SoC 市占"), sens_r("匯率"), sens_r("情境權重")
    sd = SNS["dcfw"]
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.7.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 日圓，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("：DCF 只有 "), TX(dcf, "#,##0"), T(" 日圓，同業比較法 "), TX(cadj, "#,##0"), T(" 日圓，各占一半。"))
    put("rec", T("投資建議「"), rec, T("」的由來：下檔全部來自 DCF（三情境價值都低於現價）；同業比較法採 Advantest 過去十年相對同業的本益比（"), TX(coef, "0.00"), T(" 倍）。不採相對倍數調整時加權目標價 "),
        TX(noadj, "#,##0"), T(" 日圓（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{sd}", "#,##0"),
        T(" 日圓（"), TX(f"Sens!C{sd}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{sd}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 日圓，比現價 "), TX("px", "#,##0"), T(" 日圓"), f'IF({up}<0,"低","高")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "#,##0"), T(" 日圓）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 日圓。"))
    put("o2", T("最大變數是市場規模、市占與匯率（加權目標價相對基準）：CY2026 SoC 測試機市場取 11.0／12.0 十億美元 "), TX(f"Sens!B{s_mk}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_mk}", "+0.0%;-0.0%"),
        T("；SoC 市占 ±5 個百分點 "), TX(f"Sens!B{s_sh}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_sh}", "+0.0%;-0.0%"), T("；匯率 ¥140／¥160 "), TX(f"Sens!B{s_fx}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_fx}", "+0.0%;-0.0%"),
        T("；情境權重全給保守 "), TX(f"Sens!B{s_wts}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o3", T("DCF 偏低的原因：WACC "), TX("wacc", "0.0%"), T("（股權成本 "), TX("ke", "0.0%"), T("）偏高、基準與保守情境 FY2028 起營收近乎持平。要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"), T(" 倍（現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 十億日圓，模型 "), TX(rv("rv0"), "#,##0"), T("）。"))
    put("o4", T("對帳中：CY2027–CY2029 的 AI SoC 測試機市場，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 日圓（"), f"Coord!C{CR['rec']}", T("），現行基準讀數 "), TX(f"Coord!B{CR['tp']}", "#,##0"), T("（"), f"Coord!B{CR['rec']}", T("），評等相同。個股不裁決。公司實績（SoC 營收 FY2026 Q1–Q2）對層級隱含值：無可比實績（層級值同源；標參考）。"))
    put("o5", T("進度：FY2026 上半年營收 "), TX("Q!G9+Q!H9", "#,##0"), T(" 十億日圓，占公司 10 月指引 "), TX("g_rev", "#,##0"), T(" 的 "), TX("(Q!G9+Q!H9)/g_rev", "0%"), T("；SoC 測試機 "), TX("Q!G4+Q!H4", "#,##0"), T("，占指引 "), TX("g_soc", "#,##0"),
        T(" 的 "), TX("(Q!G4+Q!H4)/g_soc", "0%"), T("。模型 FY2026 營收 "), TX(f"PL!D{PLA['rows_B']['rev']}", "#,##0"), T("，與指引差 "), TX(f"PL!D{PLA['rows_B']['rev']}/g_rev-1", "+0.0%;-0.0%"), T("。"))
    put("drv", T("分部營收推導（情境 B）：SoC 測試機＝SoC 測試機市場（FY）× Advantest 市占 × 匯率；記憶體測試機＝記憶體測試機市場 × 市占 × 匯率；其他系統與服務逐年成長；加總為合併營收。FY2028 SoC "),
        TX(f"PL!F{PLA['rows_B']['rsoc']}", "#,##0"), T("、記憶體 "), TX(f"PL!F{PLA['rows_B']['rmem']}", "#,##0"), T("、其他系統 "), TX(f"PL!F{PLA['rows_B']['roth']}", "#,##0"), T("、服務與其他 "), TX(f"PL!F{PLA['rows_B']['rsvc']}", "#,##0"),
        T("，合計 "), TX(f"PL!F{PLA['rows_B']['rev']}", "#,##0"), T(" 十億日圓。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2022–FY2025，公司決算短信 IFRS）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數；讀取層級；C0：SoC 測試機營收 FY2026 Q1–Q2 實績對層級隱含值，標參考（同源校準），結論無可比實績（層級值同源），對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, PB):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    CO = coord_c0(wb, SIL, CR, HR, PB)
    PLA = pl_standard(wb, HR, CR, PB)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((("公司：FY2026 Q1 Investor Note（2026-07-29）第 6 頁：SoC 259.6、記憶體 48.0（C0 實績）", IN1Q),
                                ("公司：FY2026 Q2 Investor Note（2026-10-28）第 6 頁：SoC 353.3、記憶體 53.4（C0 實績）", IN2Q),
                                ("公司：FY2026 Q1 決算短信（E_FR_FY2026_1Q）：合併營收 367,473、測試系統事業 333,562、服務與其他 33,911 百萬日圓（C0 核對）", FR1Q),
                                ("公司：FY2026 Q2 決算短信（E_FR_FY2026_2Q）：上半年合併營收 845,200、測試系統事業 770,231 百萬日圓（C0 核對）", FR2Q),
                                ("公司：FY2024 決算短信（E_FR_FY2024_FN）合併損益表第 8 頁：FY2023、FY2024（Hist 頁；FY2024 基本 EPS 218.67）", FN24),
                                ("公司：FY2025 決算短信（E_FR_FY2025_FN）合併損益表：FY2025、FY2024（Hist 頁）", FN25))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v075 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v075_post：Hist；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7 {VG}；Sens 相對基準 {SNS['t0']}；Coord C0 {CO['rd0']}–{CO['last']}；README {len(RMA)} 列"
