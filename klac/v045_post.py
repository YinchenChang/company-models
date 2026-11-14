# 科磊 v0.4.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5、Advantest v0.7.5）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.4.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2023–FY2026（公司 Q4 新聞稿 8-K Ex.99.1，US GAAP；會計年度 7 月到隔年 6 月；EPS 為 2026-06 十比一分割調整後）
#   PL 頁末：標準損益表（FY2024A–FY2026A GAAP 實績＋FY2027E–FY2032E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（快照 klac/sens.json）
#   Coord 頁末：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）；Sources、README 頁末
import json, os, re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
E = "https://www.sec.gov/Archives/edgar/data/319201/"
PR24 = E + "000031920124000018/exhibit991earningsrelease7.htm"
PR25 = E + "000031920125000020/exhibit991earningsrelease7.htm"
SRC_H = ("公司 Q4 新聞稿（8-K Ex.99.1）附 US GAAP 合併損益表（十二個月）：FY2024 取 FQ4-25 新聞稿（2025-07-31），FY2025 同稿並與 Q 頁一致，FY2026 取 Q 頁（FQ4-26 新聞稿 2026-07-28 全年欄）；"
         "FY2023 營收、淨利、EPS 取 FQ4-24 新聞稿（2024-07-24）。EPS 與股數為 2026-06 十比一分割調整後（FY2023、FY2024 原值 ÷ 10，Derived）；網址見 Sources 頁。千美元 ÷ 1,000 換成百萬美元")


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




# ---------- Hist：損益表歷史（GAAP） ----------
# (鍵, 項目, FY2023, FY2024, FY2025, FY2026)；百萬美元（新聞稿千美元 ÷ 1,000）；None＝未列；FY2023 只有營收、淨利、EPS（供 FY2024 年增率）
HIS = [
    ("rev", "營收（Total revenues）", 10496.056, 9812.247, 12156.162, 13579.476),
    ("prod", "  產品營收（Product）", None, 7482.679, 9472.854, 10453.537),
    ("svc", "  服務營收（Service）", None, 2329.568, 2683.308, 3125.939),
    ("cogs", "營業成本（Costs of revenues）", None, 3928.073, 4751.867, 5255.060),
    ("rd", "研發費用（Research and development）", None, 1278.981, 1360.334, 1532.118),
    ("sga", "銷管費用（Selling, general and administrative）", None, 969.509, 1029.734, 1131.518),
    ("imp", "商譽與無形資產減損（Impairment of goodwill and purchased intangible assets）", None, 289.474, 239.100, 0),
    ("ie", "利息費用（Interest expense）", None, 311.253, 302.166, 284.440),
    ("oinc", "其他收入淨額（Other expense (income), net；本列正值＝收入）", None, 155.075, 171.487, 229.585),
    ("pti", "稅前利益（Income before income taxes）", None, 3190.032, 4644.448, 5605.925),
    ("tax", "所得稅（Provision for income taxes）", None, 428.136, 582.805, 775.154),
    ("ni", "稅後淨利（Net income attributable to KLA）", 3387.277, 2761.896, 4061.643, 4830.771),
]
EPSR = [("eps_d", "EPS 稀釋（GAAP；2026-06 十比一分割調整後）", (2.415, 2.028, 3.04, 3.66)),
        ("sh_d", "稀釋加權平均股數（百萬股；分割調整後）", (None, 1361.87, 1337.502, 1319.633))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.4.5 起；PL 頁「標準損益表」的歷史三年與 FY2024 年增率用）。會計年度 7 月到隔年 6 月：FY2024＝2023-07～2024-06、FY2026＝2025-07～2026-06。"
               "金額：百萬美元；EPS：美元（2026-06 十比一分割調整後）。US GAAP；FY2023 只取營收、淨利、EPS，供 FY2024 年增率")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2026", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H)
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "0.000" if k.startswith("eps") else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified（FY2023、FY2024 為分割調整，Derived）"); H.cell(r, 8, SRC_H + "；FY2023 原 24.15、FY2024 原 20.28 與 136.187 百萬股，÷ 10")
        R[k] = r; r += 1
    for k, labl, f_ in (("gp", "毛利（＝營收 − 營業成本）", "={c}{rev}-{c}{cogs}"), ("opx", "營業費用（＝研發＋銷管＋減損）", "={c}{rd}+{c}{sga}+{c}{imp}"),
                        ("oi", "營業利益（＝毛利 − 營業費用）", "={c}{gp}-{c}{opx}"), ("oth", "其他收入淨額（＝其他收入 − 利息費用；＝稅前利益 − 營業利益）", "={c}{oinc}-{c}{ie}")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        R[k] = r
        for c in "DEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "#,##0.000"
        H.cell(r, 7, "Derived"); H.cell(r, 8, "損益表各列加減（新聞稿未列小計）")
        r += 1
    r += 1
    H.cell(r, 1, "核對（單位：百萬美元；新聞稿以千美元列示，加總差為 0）").font = BOLD; r += 1
    chk = [("產品 ＋ 服務 − 營收", "={c}{prod}+{c}{svc}-{c}{rev}"),
           ("營業利益 ＋ 其他收入淨額 − 稅前利益", "={c}{oi}+{c}{oth}-{c}{pti}"),
           ("稅前利益 − 所得稅 − 稅後淨利", "={c}{pti}-{c}{tax}-{c}{ni}"),
           ("稅後淨利 ÷ 稀釋股數 − 稀釋 EPS（美元；差 ≤ 0.005）", "={c}{ni}/{c}{sh_d}-{c}{eps_d}")]
    R["chk0"] = r
    for labl, f_ in chk:
        H.cell(r, 2, labl)
        for c in "DEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "核對：FY2025、FY2026 營收、營業利益、淨利（本頁）− Q 頁全年（三項差的合計，應為 0）")
    for c, q in (("E", "a25"), ("F", "a26")):
        H[f"{c}{r}"] = f"={c}{R['rev']}-{q}_rev+{c}{R['oi']}-{q}_oi+{c}{R['ni']}-{q}_ni"; H[f"{c}{r}"].number_format = "0.000"
    R["chk1"] = r; r += 1
    H.cell(r, 2, "核對：稀釋 EPS（本頁）− Q 頁（FY2025、FY2026）")
    for c, q in (("E", "a25"), ("F", "a26")):
        H[f"{c}{r}"] = f"={c}{R['eps_d']}-{q}_eps"; H[f"{c}{r}"].number_format = "0.00"
    R["chk2"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 66
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJK"
HC = {"C": "D", "D": "E", "E": "F"}   # 標準損益表歷史欄 → Hist 欄（FY2024→D … FY2026→F；FY2023 在 C）
FLAGY = {"F": 2027, "G": 2028, "H": 2029, "I": 2029, "J": 2029, "K": 2029}   # FY(t)＝CY(t−1) 下半＋CY(t) 上半：比值只乘 CY2027 起；FY2030–32 延伸期沿用 CY2029
PC4 = {"F": "E", "G": "F", "H": "G"}   # 標準損益表 FY2027–29 → PL 推導區塊欄
EC = {"I": "H", "J": "I", "K": "J"}   # FY2030–32 → PL 延伸區塊欄
DCOL = {"F": "C", "G": "D", "H": "E", "I": "F", "J": "G", "K": "H"}   # → DCF 頁欄（FY2027＝C … FY2032＝H）


def pl_standard(wb, HR, CR, PB):
    P = wb["PL"]; R = PB["B"]
    xb = lab(P, "情境（延伸）")
    xs = lab(P, "情境 B（基準）", xb, exact=True)
    X = dict(rev=lab(P, "營收", xs, exact=True), ni=lab(P, "淨利（收斂法延伸）", xs), eps=lab(P, "EPS（延伸）", xs))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, "標準損益表（v0.4.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋未來六年（FY2027–FY2032，情境 B 基準；FY2027 上半錨定 CY2026 下半〔第一季實績＋第二季指引〕）。會計年度 7 月到隔年 6 月；US$M；EPS 美元。"
                 "只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（美國會計準則；公司新聞稿與 10-K 合併損益）。EPS 為稀釋 EPS（2026-06 十比一分割調整後；公司另公布 Non-GAAP，共識多為 Non-GAAP）。"
                 "預測期模型推毛利（FY2026 毛利＋增量毛利率 × 營收差）、營業費用、營業利益與營業外淨額；稅前利益＝營業利益＋營業外淨額，淨利＝稅前利益 ×（1 − 稅率），所得稅＝稅前利益 − 淨利。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷（1 − 稅率）− 營業外淨額），毛利率沿用 FY2029，營業費用＝毛利 − 營業利益（推算）。FY2025 營業費用含減損 239.1、FY2024 含 289.5。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJK":
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
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1
    M = "#,##0.0"; PCT = "0.0%"
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    row("rev", "營收", "US$M", hh("rev") + [f"=PL!{PC4[c]}{R['rev']}" for c in "FGH"] + [f"=PL!{EC[c]}{X['rev']}" for c in "IJK"], M,
        "FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/Hist!C{h}-1".format(rev=L["rev"], h=HR["rev"])] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    rc = r
    row("cogs", "營業成本（COGS）", "US$M", hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHIJK"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "US$M", hh("gp") + [f"=PL!{PC4[c]}{R['gp']}" for c in "FGH"] + [f"={c}{L['rev']}*{c}{rc + 2}" for c in "IJK"], M, "FY2027E–FY2029E 取 PL 情境 B；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{rc + 1}/{c}{L['rev']}" for c in "CDEFGH"] + [f"={p}{rc + 2}" for p in "HIJ"], PCT, "FY2030E–FY2032E 沿用前一年")
    row("opx", "營業費用（研發＋銷管；歷史含減損）", "US$M", hh("opx") + [f"=PL!{PC4[c]}{R['opx']}" for c in "FGH"] + [f"={c}{rc + 1}-{c}{rc + 4}" for c in "IJK"], M,
        "FY2027E–FY2029E 為模型（FY2027 依季度指引、之後年增）；FY2030E 起＝毛利 − 營業利益（推算）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "US$M", hh("oi") + [f"=PL!{PC4[c]}{R['oi']}" for c in "FGH"] + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "IJK"], M, "FY2030E 起＝DCF 頁情境 B 的 EBIT")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（營業外淨額：其他收入 − 利息費用）", "US$M", hh("oth") + [f"=PL!{PC4[c]}{R['noi']}" for c in "FGH"] + ["=oth_exp"] * 3, M, "預測期＝In 頁營業外淨額（每年固定）")
    row("pti", "稅前利益", "US$M", hh("pti") + [f"={c}{L['oi']}+{c}{L['oth']}" for c in "FGHIJK"], M, "預測期＝營業利益＋營業外淨額")
    row("tax", "所得稅", "US$M", hh("tax") + [f"={c}{L['pti']}-{c}{r + 1}" for c in "FGHIJK"], M, "預測期＝稅前利益 − 淨利（＝稅前利益 × 稅率）")
    row("ni", "稅後淨利", "US$M", hh("ni") + [f"=PL!{PC4[c]}{R['ni']}" for c in "FGH"] + [f"=PL!{EC[c]}{X['ni']}" for c in "IJK"], M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋加權平均股數（百萬股）", "百萬股", [f"=Hist!{HC[c]}{HR['sh_d']}" for c in "CDE"] + [f"=PL!{PC4[c]}{R['sh']}" for c in "FGH"] + [f"={c}{L['ni']}/{c}{r + 1}" for c in "IJK"], "#,##0.0",
        "FY2027E–FY2029E＝PL 情境 B 股數；FY2030E 起＝淨利 ÷ EPS（延伸區塊）")
    row("eps", "EPS（GAAP 稀釋）", "US$", [f"=Hist!{HC[c]}{HR['eps_d']}" for c in "CDE"] + [f"=PL!{PC4[c]}{R['eps']}" for c in "FGH"] + [f"=PL!{EC[c]}{X['eps']}" for c in "IJK"], "0.00",
        "GAAP 稀釋 EPS（模型口徑；分割調整後）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_tax", "有效稅率（所得稅 ÷ 稅前利益）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT, "預測期＝In 頁稅率")
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/Hist!C{h}-1".format(ni=L["ni"], h=HR["ni"])] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 的基期＝FY2023 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$M", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "US$M", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS], "0.000")
    row("c3", "核對：稅前利益 ×（1 − 稅率）− 稅後淨利（預測期）", "US$M", [None] * 3 + [f"={c}{L['pti']}*(1-tax)-{c}{L['ni']}" for c in "FGHIJK"], "0.000")
    row("c4", "核對：營業利益 − DCF 頁情境 B 的 EBIT（FY2027E–FY2032E）", "US$M", [None] * 3 + [f"={c}{L['oi']}-DCF!{DCOL[c]}{d_ebit}" for c in "FGHIJK"], "0.000")
    row("c5", "核對：營收、稅後淨利 − PL 情境 B 與延伸表（FY2027E–FY2032E；兩項差的合計）", "US$M",
        [None] * 3 + [f"={c}{L['rev']}-PL!{PC4[c]}{R['rev']}+{c}{L['ni']}-PL!{PC4[c]}{R['ni']}" for c in "FGH"] + [f"={c}{L['rev']}-PL!{EC[c]}{X['rev']}+{c}{L['ni']}-PL!{EC[c]}{X['ni']}" for c in "IJK"], "0.000")
    row("c6", "核對：FY2026A（本表）− PL 情境 B 的 FY2026 實績欄（營收＋淨利差）", "US$M", [None, None, f"=E{L['rev']}-PL!D{R['rev']}+E{L['ni']}-PL!D{R['ni']}"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度 7 月到隔年 6 月，FY(t)＝CY(t−1) 下半＋CY(t) 上半；兩讀數比值只乘 CY2027 起，故 FY2027（上半為 CY2027）到 FY2029 依賴 CY2027–CY2029，標題「（對帳中）」依 Coord 頁各日曆年旗標；FY2030E–FY2032E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, rows_X=X, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF（DCF 頁 FY2027–FY2032＝C–H 欄） ----------
CCY, UB = "美元", "US$M"


def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 16, 17, 18
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數")
    assert D.cell(15, 3).value == "FY2027" and D.cell(15, 8).value == "FY2032", (D.cell(15, 3).value, D.cell(15, 8).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（第一個明確期年度）不動，營業利益率與營收倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("px", f"現價（{CCY}）", "=px", "#,##0.00")
    kv("v0", f"支持現價所需的評價日每股價值（{CCY}）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", f"淨現金滾到評價日（情境 B；{UB}）", "=" + ncr, "#,##0.00", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節；同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（百萬股）", "=sh_v", "#,##0.0")
    kv("evr", f"支持現價所需的企業價值 EV（{UB}）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.00")
    kv("ev", f"模型企業價值 EV（情境 B；{UB}）", f"=DCF!$C${ev}", "#,##0.00")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", f"模型：情境 B 評價日每股價值（{CCY}）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}
    NC = "BCDEFG"

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", f"營收（{UB}）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.00")
    yr("ebit", f"營業利益 EBIT（{UB}）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.00")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", f"NOPAT＋折舊 − 資本支出（{UB}）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.00")
    yr("wc", f"營運資金增加（{UB}）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.00")
    yr("pv", f"現值（{UB}，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.00")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", f"營收 × 貢獻係數合計（{UB}）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.00")
    kv("E0", f"與營業利益無關的 EV（{UB}）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.00")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", f"不隨倍數變動的 EV：FY2027 現值（{UB}）", f"=B{Y['pv']}", "#,##0.00")
    kv("b", f"隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（{UB}）", f"=SUM(C{Y['pv']}:G{Y['pv']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.00")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", f"現價隱含的 FY2028 營收（{UB}）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.00")
    kv("rv0", f"模型：FY2028 營收（情境 B；{UB}）", f"=C{Y['rev']}", "#,##0.00")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近下界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", f"中點的評價日每股價值（{CCY}）"], 1):
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
    kv("w_res", f"核對：隱含 WACC 下的每股價值 − 支持現價所需值（{CCY}，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, f"12 個月後前瞻 EPS（情境加權；{CCY}）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(f" {CCY}；要支持現價 "), TX("px", "#,##0"), T(f" {CCY}，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 百萬美元，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(f" {CCY}相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 科磊 過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給 科磊 的倍數是同業的 "),
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
    for j, h in enumerate(["目標價", CCY, "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(f" {CCY}；進加權目標價"))),
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
    SE.cell(r, 4, "v0.4.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（klac/sens_klac.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", f"低案：加權目標價（{CCY}）", f"高案：加權目標價（{CCY}）", f"快照基準加權目標價（{CCY}）"], 1):
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
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 klac/sens_klac.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR, PB, EQ_F):
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]; N = SIL.NAMES; UN = wb["Units"]
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    eqf = os.path.relpath(EQ_F, ROOT); eq_ver = re.search(r"v\d+(?:\.\d+)+", os.path.basename(EQ_F)).group(0)
    em = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
    wf_rows = [r for r in range(119, em.max_row + 1) if em.cell(r, 1).value == "晶圓前段設備"]
    assert len(wf_rows) == 3, wf_rows
    WFB = [em.cell(wf_rows[1], 3 + j).value for j in range(5)]
    assert abs(N["SIF_SEMI_WFE"][0] - WFB[0]) < 1e-9, (N["SIF_SEMI_WFE"][0], WFB[0])
    ub = lab(UN, "情境 B：設備市場")
    uL, uM, uP = ub + 4, ub + 7, ub + 8
    assert str(UN.cell(uL, 1).value).strip().startswith("代工邏輯 WFE 小計") and str(UN.cell(uM, 1).value).strip().startswith("記憶體 WFE 小計") and str(UN.cell(uP, 1).value).strip().startswith("先進封裝")
    RB = PB["B"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.4.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（讀 SIF_WTAnchor_Base、SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>；SEMI 晶圓前段設備 SIF_SEMI_WFE 與 Units 頁對帳列同值；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（量測與檢測設備），讀 Siliconomics"),
          ("晶圓前段與先進封裝設備市場（CY，三情境；分項先進邏輯、其他代工與邏輯、其他前段、DRAM 含 HBM、NAND、先進封裝）", "唯讀上游", f"{eqf}（Mkt 頁 5A／5B／5C；行業輸入已移到 Siliconomics M20，三大記憶體廠與台積電資本支出屬公司變數留在設備市場模型）", eq_ver, "Units 頁 → PL 引擎加權前段 WFE、先進封裝設備市場"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁敏感度案例；不改基準"),
          ("製程控制強度（前段系統 ÷ 加權前段 WFE；CY2026，模型推得）", f"=PL!D{RB['ri']}", "本模型自算：CY2026 錨點（上半實績＋下半由第一季實績與第二季指引）扣先進封裝後 ÷ 加權前段 WFE；記憶體權重 m 亦由本公司錨點推得；非 Siliconomics 輸入，層級沒有科磊份額", "v0.4.5", "PL 引擎 CY2027 起前段系統營收；C0 同源說明")]
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
    CW.cell(r, 1, "本模型會計年度 7 月到隔年 6 月；PL 引擎先算日曆年系統營收，再以 FY(t)＝CY(t−1) 下半＋CY(t) 上半換算。下列各列都是模型從上游讀入的日曆年值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。")
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
    cc = lambda y: chr(ord("C") + y - 2025)
    for i, y in enumerate(range(2025, 2030)):
        line("晶圓前段設備（情境 B；Units 頁代工邏輯＋記憶體小計）vs 設備市場模型成品", "US$B", f"CY{y}", f"=Units!{cc(y)}{uL}+Units!{cc(y)}{uM}", WFB[i], "#,##0.000",
             f"{eqf} Mkt 頁情境 B「晶圓前段設備」（Derived；CY2025＝SEMI 實績校準）；建置時直接讀成品原值")
    for i, y in enumerate(range(2025, 2030)):
        line("SEMI 晶圓前段設備（WFE；Units 頁對帳列）vs Siliconomics", "US$B", f"CY{y}", f"=Units!{cc(y)}{UR_SEMI[0]}", N["SIF_SEMI_WFE"][i], "#,##0.000",
             f"Siliconomics {SIL.VERSION} SIF_SEMI_WFE（CY2025 SEMI 實績 Verified；CY2026 起 SEMI 2026-07 年中預測，Interested-party）；只用於對帳")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"Siliconomics {SIL.VERSION} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：科磊在製程控制或晶圓前段設備的份額（層級只有市場值，沒有科磊份額）；服務、SPTS、PCB 與元件檢測、中國營收（層級不持有）；製程控制強度年增、先進封裝倍數（公司變數）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準（設備市場模型情境 B 的 AI 分項）與 Wattonomics 換算值兩層之間的差距，不是科磊自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for y in (2027, 2028, 2029):
        CW.cell(r, 1, "AI 設備（先進邏輯＋DRAM 含 HBM＋先進封裝；情境 B）：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, f"=Units!{cc(y)}{ub + 1}+Units!{cc(y)}{ub + 5}+Units!{cc(y)}{uP}").number_format = "#,##0.000"; CW.cell(r, 5, f"=D{r}*{ratio(y)}").number_format = "#,##0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；對帳中年份依 C4 並列兩讀數目標價（本頁上方），個股不裁決")
        r += 1
    ROW["ref2"] = r - 1
    r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t−1) 下半＋CY(t) 上半：FY2027 的上半起到 FY2029，及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；FY2027 下半錨定 CY2026 錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 美元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 美元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的設備市場 AI 分項 × 另一側比值；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數只縮放設備市場中標〔AI〕的先進邏輯、DRAM（含 HBM）與先進封裝，其他代工邏輯、其他前段與 NAND 不變；設備市場由晶圓廠資本支出推得、與部署 GW 只間接連動，同年比值為近似。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    # ---- 4. C0 ----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.4.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績：2026 年 1–9 月（FQ3-26、FQ4-26、FQ1-27）半導體製程控制系統營收（＝產品營收 − SPTS − PCB 與元件檢測兩事業營收，各項為 8-K 新聞稿 Verified，相減為 Derived；Q 頁第 11 列），不用指引或預測。"
                  "層級側＝設備市場（加權前段 WFE、先進封裝設備，設備市場模型情境 B；2025–2027 三情境相同）× 科磊強度。主讀數＝已公布期間累計 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；"
                  "第二讀數＝最近一季（FQ1-27，2026-07～09）× 4 vs 層級 CY2026（季中點約 8 月中旬，晚於日曆年中點，營收逐季上升時偏高；不得寫成「夾住真值」）。兩讀數差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」；"
                  "層級沒有本公司份額、或隱含值任一部分與個股同源時只標參考、不判定（tools/coord/README）。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（市場值消去）：Siliconomics 與設備市場模型只有市場值（由晶圓廠資本支出 × 設備占比推得，2025 以 SEMI 實績校準；不由本模型校準——科磊「CY2026 WFE low-$150B」等說法只列在設備市場模型 Recon 對照），沒有科磊份額。"
                  "模型的製程控制強度＝CY2026 前段系統錨點（上半實績＋下半由第一季實績與第二季指引推得）÷ 加權前段 WFE，先進封裝為公司說法（約 1.1 十億美元），代入「強度 × 加權 WFE ＋ 先進封裝」後市場消去，層級隱含值恆等於科磊自己的 CY2026 系統錨點（下方核對列為 0）。"
                  "所以這張表實際檢驗的只是「1–9 月實績 ÷ 錨點是否≈ 9/12」，即十二月季指引占全年的比重，不是對層級值的獨立檢驗；依規則標「參考（同源校準）」，結論「無可比實績（層級無本公司份額；以模型份額換算則同源）」，對帳中維持 0。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "缺的獨立揭露：(1) 非科磊自估的 CY2026 製程控制（量測與檢測）設備市場或科磊份額（模型中的市占 58% 為賣方數字，只供顯示）；(2) 同業（應用材料、Lasertec、Onto 等）同口徑營收的獨立彙總。"
                  "參考列：份額沿用 CY2025 實績（CY2025 系統營收 ÷ 晶圓前段設備 CY2025＝SEMI 實績）時，層級隱含值只隨晶圓前段設備 CY2026 年增變動——這列確實用到層級市場值，但份額不變（且不分記憶體與代工邏輯權重）是假設、不是層級值，故只列參考、不判定。")
    ROW["none"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; CW.cell(r, 9, "鍵名（追蹤用；HTML 不顯示）").font = BOLD; r += 1
    inp = {}

    def put(key, label, unit, per, val, grade, src, fmt):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        inp[key] = r; r += 1
    put("q1", "半導體製程控制系統營收：FQ3-26", "US$M", "2026-01～2026-03", "=Q!E11", "Verified 分項相減（Derived）", "Q 頁 E11（＝產品 2,640.3 − SPTS 164.0 − PCB 167.6；FQ3-26 新聞稿 8-K 2026-04-29）", "#,##0.0")
    put("q2", "半導體製程控制系統營收：FQ4-26", "US$M", "2026-04～2026-06", "=Q!F11", "Verified 分項相減（Derived）", "Q 頁 F11（FQ4-26 新聞稿 8-K 2026-07-28）", "#,##0.0")
    put("q3", "半導體製程控制系統營收：FQ1-27（最近一季）", "US$M", "2026-07～2026-09", "=Q!I11", "Verified 分項相減（Derived）", "Q 頁 I11（FQ1-27 新聞稿 8-K 2026-10-28）", "#,##0.0")
    put("ytd", "半導體製程控制系統營收：2026 年 1–9 月累計", "US$M", "2026-01～2026-09", f"=D{r - 3}+D{r - 2}+D{r - 1}", "Derived（三季相加）", "三季相加", "#,##0.0")
    put("mo", "已公布月數", "月", "2026-01～2026-09", 9, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("L", "層級值：代工邏輯 WFE CY2026（情境 B）", "US$B", "CY2026", f"=Units!D{uL}", "Derived（層級值）", f"Units 頁（{eqf} Mkt 頁 5B：先進邏輯＋其他代工與邏輯＋其他前段）", "#,##0.000")
    put("M", "層級值：記憶體 WFE CY2026（情境 B）", "US$B", "CY2026", f"=Units!D{uM}", "Derived（層級值）", "Units 頁（DRAM 含 HBM＋NAND）", "#,##0.000")
    put("P", "層級值：先進封裝設備 CY2026（情境 B）", "US$B", "CY2026", f"=Units!D{uP}", "Derived（層級值）", "Units 頁（先進封裝〔AI〕）", "#,##0.000")
    put("W", "加權前段 WFE（代工邏輯＋記憶體權重 × 記憶體；模型）", "US$M", "CY2026", f"=PL!D{RB['W']}", "Derived（記憶體權重由本公司錨點推得）", "PL 頁情境 B「加權前段 WFE」CY2026 欄", "#,##0.0")
    put("ri", "製程控制強度 CY2026（模型推得；同源）", "%", "CY2026", f"=PL!D{RB['ri']}", "Derived（本模型；同源）", "PL 頁情境 B「製程控制強度」CY2026 欄＝前段系統錨點 ÷ 加權前段 WFE", "0.00%")
    put("ap", "先進封裝製程控制系統營收 CY2026（公司說法）", "US$M", "CY2026", f"=PL!D{RB['ap']}", "Interested-party", "In 頁 ap26（公司 FQ4-26 法說：約 1.1 十億美元）", "#,##0.0")
    put("lay", "層級隱含 半導體製程控制系統營收（模型強度）", "US$M", "CY2026", f"=D{inp['ri']}*D{inp['W']}+D{inp['ap']}", "Derived", "強度 × 加權前段 WFE ＋ 先進封裝；連到上列各格", "#,##0.0")
    put("idc", "核對：層級隱含 − CY2026 系統錨點（應為 0：市場消去）", "US$M", "CY2026", f"=D{inp['lay']}-PL!D{RB['sys']}", "Derived", "證明層級隱含值＝科磊自己的上半實績加下半錨點（同源）", "0.000")
    put("a25", "參考：CY2025 半導體製程控制系統營收（實績）", "US$M", "2025-01～2025-12", "=cy25_sys", "Verified 分項相減（Derived）", "Units 頁（Q 頁 FQ3-25、FQ4-25、FY26 上半相加）", "#,##0.0")
    put("w25", "參考：晶圓前段設備 CY2025（SEMI 實績，設備市場模型校準年）", "US$B", "CY2025", f"=Units!C{uL}+Units!C{uM}", "Verified", "Units 頁代工邏輯＋記憶體小計 CY2025＝SEMI 2025 實績 116.9", "#,##0.000")
    put("ref", "參考：層級隱含 系統營收（份額沿用 CY2025 實績）", "US$M", "CY2026", f"=D{inp['a25']}*(D{inp['L']}+D{inp['M']})/D{inp['w25']}", "Derived（份額為假設）", "＝CY2025 實績 × 晶圓前段設備 CY2026 ÷ CY2025；份額不變為假設（非層級值）", "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge="（讀數）"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "US$M"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("主讀數：已公布期間累計系統營收 vs 層級 × 月數 ÷ 12（模型強度）", "2026-01～09 vs CY2026 × 9/12", f"=D{inp['ytd']}", f"=D{inp['lay']}*D{inp['mo']}/12",
             "實績＝三季相加；層級＝（強度 × 加權 WFE＋先進封裝）× 9 ÷ 12（線性時差近似）。層級隱含值＝錨點，差距只反映十二月季指引占全年比重")
    rb = rd_("第二讀數：最近一季（FQ1-27）× 4 vs 層級（模型強度）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp['lay']}",
             "最近一季年化；季中點（8 月中旬）晚於日曆年中點，成長期偏高；只作時差近似的另一端參考")
    CW.cell(r, 1, "C0（實績）判定：半導體製程控制系統營收"); CW.cell(r, 7, "參考（同源校準）")
    CW.cell(r, 8, "結論：無可比實績（層級無本公司份額；以模型份額換算則同源）：層級隱含值＝科磊自己的實績加指引（市場消去），缺非科磊自估的製程控制市場或科磊份額揭露。數字參考：若不論同源，兩讀數差距見上列；對帳中維持 0")
    ROW["judge"] = r; r += 1
    rr1 = rd_("參考：主讀數（份額沿用 CY2025 實績）", "2026-01～09 vs CY2026 × 9/12", f"=D{inp['ytd']}", f"=D{inp['ref']}*D{inp['mo']}/12",
              "層級市場值有作用（晶圓前段設備 CY2026 ÷ CY2025）；份額不變為假設，故只列參考", "參考（份額假設，不判定）")
    rr2 = rd_("參考：第二讀數（份額沿用 CY2025 實績）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp['ref']}",
              "同上；季中點晚於年中，成長期偏高", "參考（份額假設，不判定）")
    CW.cell(r, 1, "參考：已公布期間實績占模型 CY2026 系統錨點（個股自身）"); CW.cell(r, 2, "US$M"); CW.cell(r, 3, "2026-01～09 vs CY2026 錨點")
    CW.cell(r, 4, f"=D{inp['ytd']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=PL!D{RB['sys']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（1–9 月實績 ÷ 錨點），不是差距；錨點含十二月季公司指引，非實績")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, "讀法：公司實績是新聞稿分項相減、獨立；層級沒有科磊份額，用模型強度換算時市場消去（同源），用 CY2025 實績份額換算時份額不變是假設，兩者都不能判定。結論：無可比實績，數字僅供參考。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["rr1"], ROW["rr2"], ROW["last"] = ra, rb, rr1, rr2, r - 1
    r += 1
    CW.cell(r, 1, "其他公司實績（無可比層級值；2026 年 1–9 月與最近一季）").font = BOLD; ROW["oth_t"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "2026 年 1–9 月累計", "FQ1-27", "", "判定", "出處／口徑／缺少的揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_oth"] = r; r += 1
    others = [("服務營收", "=Q!E6+Q!F6+Q!I6", "=Q!I6", "無可比實績（層級無對應量）", "Q 頁第 6 列；層級不持有服務營收"),
              ("半導體製程控制事業（含服務；事業別）", "=Q!E7+Q!F7+Q!I7", "=Q!I7", "無可比實績（層級無科磊份額）", "Q 頁第 7 列（新聞稿事業別表，Verified）"),
              ("SPTS＋PCB 與元件檢測", "=Q!E12+Q!F12+Q!I12", "=Q!I12", "無可比實績（層級無對應量）", "Q 頁第 12 列"),
              ("合併營收", "=Q!E4+Q!F4+Q!I4", "=Q!I4", "無可比實績（層級無科磊營收）", "Q 頁第 4 列")]
    ROW["oth0"] = r
    for nm, a9, a3, judge, miss in others:
        CW.cell(r, 1, nm); CW.cell(r, 2, "US$M"); CW.cell(r, 3, "2026-01～09／2026-07～09")
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
    s_wf, s_pu, s_gm, s_wts = sens_r("晶圓前段設備市場"), sens_r("製程控制強度年增"), sens_r("增量毛利率"), sens_r("情境權重")
    sd = SNS["dcfw"]
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rb = PLA["rows_B"]
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    pc = lambda c: TX(c, "+0.0%;-0.0%")
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.4.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 美元，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("：DCF 只有 "), TX(dcf, "#,##0"), T(" 美元，同業比較法 "), TX(cadj, "#,##0"), T(" 美元，各占一半。"))
    put("rec", T("投資建議「"), rec, T("」的由來：DCF 遠低於現價（三情境價值都低於現價），同業比較法採科磊過去十年相對同業的本益比（"), TX(coef, "0.00"), T(" 倍）。不採相對倍數調整時加權目標價 "),
        TX(noadj, "#,##0"), T(" 美元（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{sd}", "#,##0"),
        T(" 美元（"), TX(f"Sens!C{sd}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{sd}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 美元，比現價 "), TX("px", "#,##0.00"), T(" 美元"), f'IF({up}<0,"低 ","高 ")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.00"), T(" 美元）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 美元。"))
    put("o2", T("最大變數是設備市場（加權目標價相對基準）：CY2027 起晶圓前段與先進封裝設備 ±15% "), pc(f"Sens!B{s_wf}"), T("／"), pc(f"Sens!C{s_wf}"),
        T("；製程控制強度年增取低／高 "), pc(f"Sens!B{s_pu}"), T("／"), pc(f"Sens!C{s_pu}"), T("；增量毛利率取低／高 "), pc(f"Sens!B{s_gm}"), T("／"), pc(f"Sens!C{s_gm}"),
        T("；情境權重全給保守 "), pc(f"Sens!B{s_wts}"), T("。這些是測試案例，不是門檻。"))
    put("o3", T("DCF 偏低的原因：WACC "), TX("wacc", "0.0%"), T("（股權成本 "), TX("ke", "0.0%"), T("）、基準情境 FY2028 起營收成長放緩、延伸期淨利率下滑。要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"), T(" 倍（現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 百萬美元，模型 "), TX(rv("rv0"), "#,##0"), T("）。"))
    put("o4", T("對帳中：CY2027–CY2029 的 AI 設備（先進邏輯、DRAM 含 HBM、先進封裝），Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 美元（"), f"Coord!C{CR['rec']}", T("），現行基準讀數 "), TX(f"Coord!B{CR['tp']}", "#,##0"), T("（"), f"Coord!B{CR['rec']}", T("）"),
        f'IF(Coord!B{CR["rec"]}=Coord!C{CR["rec"]},"，評等相同","，評等不同")', T("。個股不裁決。公司實績（製程控制系統營收 2026 年 1–9 月）對層級隱含值：無可比實績（層級無科磊份額；標參考）。"))
    put("o5", T("進度：2026 年 1–9 月半導體製程控制系統營收 "), TX(f"Coord!D{CO['ytd_r']}", "#,##0"), T(" 百萬美元，占模型 CY2026 錨點（含十二月季指引）"), TX(f"Coord!E{CO['prog_r']}", "#,##0"), T(" 的 "), TX(f"Coord!F{CO['prog_r']}", "0%"),
        T("；最近一季合併營收 "), TX("Q!I4", "#,##0"), T(" 百萬美元。"))
    put("drv", T("系統營收推導（情境 B）：前段系統＝製程控制強度 × 加權前段 WFE（代工邏輯＋記憶體權重 × 記憶體），強度每年上升；先進封裝系統隨封裝設備市場成長；服務、SPTS 與 PCB 逐年成長；再以上下半年換算會計年度。FY2028 系統 "),
        TX(f"PL!F{rb['fsys']}", "#,##0"), T("、服務 "), TX(f"PL!F{rb['fsvc']}", "#,##0"), T("、SPTS 與 PCB "), TX(f"PL!F{rb['foth']}", "#,##0"), T("，合計 "), TX(f"PL!F{rb['rev']}", "#,##0"), T(" 百萬美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達、Advantest）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2023–FY2026，US GAAP，分割調整後 EPS）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B，GAAP）、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數；讀取層級；C0：製程控制系統營收 2026 年 1–9 月實績對層級隱含值，層級無科磊份額、以模型強度換算則同源，結論無可比實績，對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, PB, EQ_F):
    global UR_SEMI
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    UR_SEMI = [lab(wb["Units"], "SEMI 晶圓前段設備（WFE")]
    HR = hist_is(wb)
    CO = coord_c0(wb, SIL, CR, HR, PB, EQ_F)
    CO["ytd_r"] = lab(wb["Coord"], "半導體製程控制系統營收：2026 年 1–9 月累計")
    CO["prog_r"] = lab(wb["Coord"], "參考：已公布期間實績占模型 CY2026 系統錨點")
    PLA = pl_standard(wb, HR, CR, PB)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((("公司：FQ4-25 新聞稿（8-K Ex.99.1，2025-07-31）：FY2025、FY2024 合併損益表（標準損益表歷史用；分割前 EPS）", PR25),
                                ("公司：FQ4-24 新聞稿（8-K Ex.99.1，2024-07-24）：FY2023 營收 10,496.1、淨利 3,387.3 百萬美元、稀釋 EPS 24.15（分割前；FY2024 年增率基期）", PR24))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v045 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v045_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 加權目標價相對基準 {SNS['t0']} 起；Coord 讀取層級與 C0 {CO['rd0']}–{CO['last']}；README v1.3 白話列 {len(RMA)} 列"


UR_SEMI = [None]
