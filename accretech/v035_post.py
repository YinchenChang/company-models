# 東京精密 v0.3.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5、Advantest v0.7.5）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.3.4 逐格比對驗證）：
#   HistIS 頁（新）：標準損益表歷史 FY2022–FY2025（營收、營業利益、淨利、EPS 連既有 Hist 頁；毛利、銷管費取 In 頁 a25_gp／a25_sga 與其說明；稅前與所得稅本輪未取得）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（快照 accretech/sens.json）
#   Coord 頁末：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）；Sources、README 頁末
import json, os, re
from openpyxl.styles import Font, PatternFill

VER, BUILT = "0.3.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
TSQ2 = "公司 2027年3月期 第2四半期（中間期）決算短信（2026-11-04）"
NA = "未取得（本輪無法讀取原文；待補）"
FS25 = "公司 2025年3月期 決算短信〔日本基準〕（連結）前期欄，2025-05-09（https://www2.jpx.co.jp/disc/77290/140120250508534580.pdf）"   # v0.3.6
FS26 = "公司 2026年3月期 決算短信〔日本基準〕（連結），2026-05-13（https://www.accretech.com/jp/ir/news/news-202605180003/main/0/link/FS_2026_4q_J_rev.pdf）"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def lab(ws, prefix, start=1, col=1, exact=False):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def F(*parts):
    """公式字串：parts 為 ("t", 文字) 或 Excel 運算式（字串）。"""
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
TX = lambda ref, f: f'TEXT({ref},"{f}")'
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- HistIS：標準損益表歷史 ----------
def hist_is(wb):
    Hs = wb["Hist"]
    rr = {fy: lab(Hs, f"FY{fy}（") for fy in (2022, 2023, 2024, 2025)}
    assert [rr[y] for y in (2022, 2023, 2024, 2025)] == [13, 14, 15, 16], rr
    H = wb.create_sheet("HistIS", index=wb.sheetnames.index("Hist") + 1)
    H["A1"] = ("HistIS — 標準損益表歷史（v0.3.5 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度 4–3 月：FY2023＝2023-04～2024-03。金額：十億日圓；EPS：日圓。日本會計準則。"
               "營收、營業利益、淨利、EPS 連既有 Hist 頁（公司決算短信、松井證券）；毛利 FY2025 取 In 頁 a25_gp；銷管費 FY2025 取 a25_sga。"
               "v0.3.6 起補 FY2023 毛利、FY2023–FY2024 銷管費與 FY2023–FY2025 稅前利益、所得稅（公司決算短信原文；v0.3.5 為未取得）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    hc = {"C": 2022, "D": 2023, "E": 2024, "F": 2025}
    link = lambda col: {c: f"=Hist!{col}{rr[y]}" for c, y in hc.items()}
    R = {}
    rows = [("rev", "營收（売上高）", link("B"), "Verified", "既有 Hist 頁 B 欄（公司決算短信、松井證券）"),
            ("cogs", "營業成本（売上原価；＝營收 − 毛利）", {c: f"={c}4-{c}6" for c in "DEF"}, "Derived", "營收 − 毛利"),
            ("gp", "毛利（売上総利益）", {"D": 54.762, "E": 62.453, "F": "=a25_gp"}, "Verified", "FY2025：In 頁 a25_gp（68,860）；FY2024：62,453（" + FS26 + "）；FY2023：54,762（" + FS25 + "）"),
            ("opx", "銷管費（販売費及び一般管理費，含研發）", {"D": 29.454, "E": 32.750, "F": "=a25_sga"}, "Verified", "FY2025：In 頁 a25_sga 35,122；FY2024：32,750（" + FS26 + "）；FY2023：29,454（" + FS25 + "）"),
            ("oi", "營業利益（営業利益）", {c: v for c, v in link("C").items() if c != "C"}, "Verified", "既有 Hist 頁 C 欄"),
            ("ni", "稅後淨利（歸屬母公司）", link("D"), "Verified", "既有 Hist 頁 D 欄"),
            ("eps_b", "EPS 基本（公司公告值）", link("E"), "Verified", "既有 Hist 頁 E 欄（過去十年無股票分割）"),
            ("sh_b", "加權平均股數（百萬股；＝淨利 ÷ 基本 EPS）", {c: f"={c}9*1000/{c}10" for c in "DEF"}, "Derived", "淨利 ÷ 基本 EPS（扣除員工持股信託後的公司口徑近似）")]
    r = 4
    for k, labl, cells, tag, src in rows:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c, v in cells.items():
            x = H[f"{c}{r}"]; x.value = v; x.number_format = "0.00" if k == "eps_b" else "#,##0.000"
            if not (isinstance(v, str) and v.startswith("=")): x.font = BLUE
        H.cell(r, 7, tag); H.cell(r, 8, src)
        R[k] = r; r += 1
    assert (R["rev"], R["gp"], R["oi"], R["ni"], R["eps_b"]) == (4, 6, 8, 9, 10), R
    for k, labl, cells, src in (("pti", "稅前利益（税金等調整前当期純利益）", {"D": 27.255, "E": 34.275, "F": 33.186}, "FY2023：27,255（" + FS25 + "）；FY2024：34,275、FY2025：33,186（" + FS26 + "）"),
                                ("tax", "所得稅（法人税等合計）", {"D": 7.791, "E": 8.531, "F": 8.354}, "FY2023：7,791（" + FS25 + "）；FY2024：8,531、FY2025：8,354（" + FS26 + "）")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c, v in cells.items():
            H[f"{c}{r}"] = v; H[f"{c}{r}"].number_format = "#,##0.000"; H[f"{c}{r}"].font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, src + "；v0.3.6 補（In 頁 tax_w 仍為 v0.3 的第三方近似，未改）"); R[k] = r; r += 1
    H.cell(r, 2, "稅前利益 − 所得稅 ≠ 歸屬母公司淨利：差額為非支配株主に帰属する当期純利益（金額小）"); r += 2
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    R["chk0"] = r
    H.cell(r, 2, "毛利 − 銷管費 − 營業利益（FY2023–FY2025；決算短信，百萬日圓捨入）")
    for c in "DEF": H[f"{c}{r}"] = f"={c}6-{c}7-{c}8"; H[f"{c}{r}"].number_format = "0.000"
    r += 1
    H.column_dimensions["A"].width = 10; H.column_dimensions["B"].width = 56
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["G"].width = 22; H.column_dimensions["H"].width = 100
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
    P.cell(r, 1, "標準損益表（v0.3.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準）。會計年度 4–3 月；十億日圓；EPS 日圓。只引用本頁上方、HistIS、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：日本會計準則（J-GAAP；公司決算短信〔日本基準〕；Non-GAAP 不適用，東京精密不另公布）。EPS 為基本 EPS（與模型、公司指引同口徑）。"
                 "預測期模型以部門營業利益推導（SPE、計測各自＝FY2026 錨點＋邊際利潤率 × 營收變化），不推毛利與營業費用，故預測期營業成本、毛利、營業費用留空；淨利＝營業利益 × 淨利÷營業利益比；稅前利益＝淨利 ÷（1 − WACC 用稅率），所得稅＝稅前利益 − 淨利，其他收入淨額（營業外收支＋特別損益）＝稅前利益 − 營業利益，三者皆為由模型淨利反推的 Derived 列，不進 EPS、現金流與目標價。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷ 淨利÷營業利益比），毛利率沿用 FY2029，營業費用＝毛利 − 營業利益（推算）。FY2026E 為錨點（公司 2026-11-04 修正指引；上半為實績；三情境相同）。歷史三年毛利、銷管費、稅前利益、所得稅為公司決算短信（v0.3.6 補齊）；歷史的稅前利益 − 所得稅與歸屬母公司淨利差非支配股東損益。")
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
    hh = lambda k: [f"=HistIS!{HC[c]}{HR[k]}" for c in "CDE"]
    fc = lambda k: [f"=PL!{PC4[c]}{R[k]}" for c in "FGHI"]
    row("rev", "營收", "¥B", hh("rev") + fc("rev") + [f"=PL!{EC[c]}{X['rev']}" for c in "JKL"], M, "FY2026E＝錨點；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/HistIS!C{h}-1".format(rev=L["rev"], h=HR["rev"])] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    rc = r
    row("cogs", "營業成本（COGS）", "¥B", [f"=HistIS!{HC[c]}{HR['cogs']}" for c in "CDE"], M, "歷史＝營收 − 毛利；預測期模型不推（部門營業利益法），留空")
    row("gp", "毛利", "¥B", [f"=HistIS!{HC[c]}{HR['gp']}" for c in "CDE"], M, "歷史＝決算短信；預測期模型不推，留空")
    row("gm", "毛利率", "%", [f"={c}{rc + 1}/{c}{L['rev']}" for c in "CDE"], PCT, "預測期留空")
    row("opx", "營業費用（銷管費，含研發）", "¥B", [f"=HistIS!{HC[c]}{HR['opx']}" for c in "CDE"], M, "歷史＝決算短信銷管費；預測期留空")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "¥B", hh("oi") + [f"=PL!{PC4[c]}{R['oi']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['ni']}/nioi" for c in "JKL"], M, "FY2030E 起＝淨利 ÷ nioi（同 DCF 頁 EBIT）")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）", "¥B", [f"={c}{r + 1}-{c}{L['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（營業外收支＋特別損益）；預測期為由淨利反推的 Derived 值")
    row("pti", "稅前利益", "¥B", [f"=HistIS!{HC[c]}{HR['pti']}" for c in "CDE"] + [f"={c}{r + 2}/(1-tax_w)" for c in "FGHIJKL"], M, "歷史＝決算短信；預測期＝淨利 ÷（1 − tax_w）")
    row("tax", "所得稅", "¥B", [f"=HistIS!{HC[c]}{HR['tax']}" for c in "CDE"] + [f"={c}{L['pti']}-{c}{r + 1}" for c in "FGHIJKL"], M, "歷史＝決算短信（法人税等合計；與淨利差非支配股東損益）；預測期＝稅前利益 − 淨利")
    row("ni", "稅後淨利（歸屬母公司）", "¥B", hh("ni") + [f"=PL!{PC4[c]}{R['ni']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['ni']}" for c in "JKL"], M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "基本股數（百萬股）", "百萬股", [f"=HistIS!{HC[c]}{HR['sh_b']}" for c in "CDE"] + ["=shares"] * 4 + [f"={c}{L['ni']}*1000/{c}{r + 1}" for c in "JKL"], M, "歷史＝淨利 ÷ EPS（Derived）；FY2026E–FY2029E＝In 頁 shares（中間期期中平均，扣除員工持股信託）；FY2030E 起＝淨利 ÷ EPS（延伸區塊）")
    row("eps", "EPS（J-GAAP 基本）", "¥", [f"=HistIS!{HC[c]}{HR['eps_b']}" for c in "CDE"] + [f"=PL!{PC4[c]}{R['eps']}" for c in "FGHI"] + [f"=PL!{EC[c]}{X['eps']}" for c in "JKL"], "#,##0.00", "基本 EPS；共識與公司指引為基本 EPS")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in "CDE"], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_tax", "有效稅率（所得稅 ÷ 稅前利益）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT, "歷史＝決算短信；預測期＝In 頁 tax_w")
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/HistIS!C{h}-1".format(ni=L["ni"], h=HR["ni"])] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利（FY2023A–FY2025A）", "¥B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "CDE"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2025A）", "¥B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDE"], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利（預測期）", "¥B", [None] * 3 + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in "FGHIJKL"], "0.000")
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
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
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
                     T(" 倍（同業中位數 × 東京精密過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給東京精密的倍數是同業的 "),
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
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.3.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
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
    SE.cell(r, 4, "v0.3.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（accretech/sens_accretech.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
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
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 accretech/sens_accretech.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))



# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR, PB, EQ_F):
    import openpyxl
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    eqf = os.path.relpath(EQ_F, ROOT); eq_ver = re.search(r"v\d+(?:\.\d+)+", os.path.basename(EQ_F)).group(0)
    em = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
    EQB, AIB, bk = {}, {}, None
    for row in em.iter_rows():
        a = row[0].value
        if isinstance(a, str) and a[:3] in ("5A.", "5B.", "5C."): bk = a[1]; continue
        if bk == "B" and isinstance(a, str):
            for key, lb in (("tst", "測試與其他封裝"), ("apk", "先進封裝〔AI〕"), ("wfe", "晶圓前段設備")):
                if a.startswith(lb) and key not in EQB: EQB[key] = [row[2 + j].value for j in range(5)]
            if a in ("先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕", "DRAM（含 HBM）〔AI〕") and a not in AIB: AIB[a] = [row[2 + j].value for j in range(5)]
    assert len(EQB) == 3 and len(AIB) == 2, (EQB, AIB)
    RB = PB["B"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.3.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（讀 SIF_WTAnchor_Base、SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>；設備市場的行業輸入在 Siliconomics M20；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（探針台、切割研削），讀 Siliconomics 與設備市場層"),
          ("設備市場（CY，三情境；測試與其他封裝、先進封裝〔AI〕、晶圓前段設備）", "唯讀上游", f"{eqf}（Mkt 頁 5A–5C 節）", eq_ver, "In 頁 tst／apk／wfe<情境>_<年>；探針台與加工裝置需求指數"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁敏感度案例；不改基準")]
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
    CW.cell(r, 1, "本模型會計年度 4–3 月；設備市場為日曆年，需求指數以 FY(t)＝3/4×CY(t)＋1/4×CY(t+1) 換算。下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。")
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
        CW.cell(r, 6, f'=IF(E{r}=0,IF(D{r}=0,0,"—"),D{r}/E{r}-1)').number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
    for key, nm in (("tst", "測試與其他封裝設備"), ("apk", "先進封裝設備〔AI〕"), ("wfe", "晶圓前段設備")):
        for i, y in enumerate(range(2025, 2030)):
            line(f"{nm}（情境 B；In 頁讀入）vs 設備市場模型成品", "US$B", f"CY{y}", f"={key}B_{y % 100}", EQB[key][i], "#,##0.000", f"{eqf} Mkt 頁 5B 節（Derived；CY2025＝SEMI 實績校準）；建置時直接讀成品原值")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"Siliconomics {SIL.VERSION} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：東京精密在探針台、切割研削設備的份額（層級只有市場值）；SPE 超額成長、中國占比與變化、探針台占比、計測機器（公司變數或 Assumed）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準（設備市場模型情境 B 的 AI 分項：先進封裝〔AI〕＋先進邏輯〔AI〕＋DRAM 含 HBM〔AI〕）與 Wattonomics 換算值兩層之間的差距，不是東京精密自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for y in (2027, 2028, 2029):
        i = y - 2025
        CW.cell(r, 1, "AI 設備（先進封裝＋先進邏輯＋DRAM 含 HBM；情境 B）：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, f"CY{y}")
        c = CW.cell(r, 4, EQB["apk"][i] + sum(v[i] for v in AIB.values())); c.number_format = "#,##0.000"; c.font = BLUE
        CW.cell(r, 5, f"=D{r}*{ratio(y)}").number_format = "#,##0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"{eqf} 情境 B（建置時讀成品）；Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；個股不裁決")
        r += 1
    ROW["ref2"] = r - 1
    r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝3/4×CY(t)＋1/4×CY(t+1)：FY2027–FY2029 及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；FY2026 為公司指引錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0"), T(" 日圓（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0"),
                    T(" 日圓（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的設備市場 AI 分項 × 另一側比值；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數縮放先進封裝〔AI〕與晶圓前段設備中的 AI 分項（只影響加工裝置需求指數）；測試與其他封裝設備（探針台需求指數）未縮放，FY2026 指數分母維持未縮放。設備市場由資本支出推得、與部署 GW 只間接連動，同年比值為近似。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    # ---- 4. C0 ----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.3.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績：FY2026 Q1–Q2（2026-04～2026-09；會計年度 4 月起）半導體製造裝置（SPE）部門營收（中間期決算短信部門資訊；Q1 為說明資料億日圓取整），不用指引或預測。層級側：設備市場層有「測試與其他封裝設備」市場值（探針台所屬；CY2025 以 SEMI 實績校準），沒有東京精密份額。"
                  "主讀數＝已公布期間累計 vs 層級 CY2026 × 6 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（Q2，2026-07～09）× 4 vs 層級 CY2026（季中點約 8 月中旬，晚於日曆年中點，營收逐季上升時偏高；不得寫成「夾住真值」）。"
                  "兩讀數差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」；層級沒有本公司份額時只標參考、不判定（tools/coord/README）。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查：設備市場值由晶圓廠資本支出與 SEMI 實績推得，不由本模型校準，與東京精密實績無同源關係；但層級沒有東京精密份額，要得到「層級隱含東京精密 SPE 營收」只能假設份額。"
                  "下表以「份額沿用 FY2025 實績」（FY2025 SPE 營收 ÷ CY2025 測試與其他封裝設備，匯率不變；FY 與 CY 錯開三個月，近似）換算：層級隱含值＝FY2025 SPE 營收 × 市場 CY2026 ÷ CY2025。份額不變是假設、不是層級值，故只列參考、不判定；結論「無可比實績（層級無東京精密份額）」，對帳中維持 0。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "缺的獨立揭露：(1) 層級的探針台市場、切割研削設備市場或東京精密份額；(2) 口徑差：SPE 營收含探針台與加工裝置（切割、研削、CMP），層級市場以美元計、含其他廠與測試機以外的封裝設備；(3) SPE 受注（訂單）層級無對應量。計測機器部門與合併營收列為其他實績。")
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
    put("q1", "SPE 營收：FY2026 Q1", "¥B", "2026-04～2026-06", "=Q!G8", "Interested-party（公司說明資料，億日圓取整）", "Q 頁 G8（連格）；公司 2026年度第1四半期 決算說明資料 285 億日圓（網址見 Sources 頁）", "#,##0.000")
    put("q2", "SPE 營收：FY2026 Q2（最近一季）", "¥B", "2026-07～2026-09", "=Q!H8", "Verified（上半 − Q1，Derived）", f"Q 頁 H8（連格）＝{TSQ2} 上半 72,645 百萬日圓 − Q1 28.5", "#,##0.000")
    put("h1", "SPE 營收：FY2026 Q1–Q2 累計", "¥B", "2026-04～2026-09", f"=D{r - 2}+D{r - 1}", "Verified", f"{TSQ2} 部門資訊：上半 SPE 72,645 百萬日圓（+22.3%）", "#,##0.000")
    put("mo", "已公布月數", "月", "2026-04～2026-09", 6, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("m26", "層級值：測試與其他封裝設備市場 CY2026（情境 B）", "US$B", "CY2026", "=tstB_26", "Derived（層級值）", f"In 頁 tstB_26（{eqf} Mkt 頁 5B；SEMI 測試＋組裝封裝 − 先進封裝〔AI〕）", "#,##0.000")
    put("m25", "層級值：測試與其他封裝設備市場 CY2025（SEMI 實績校準）", "US$B", "CY2025", "=tstB_25", "Verified（層級值）", f"In 頁 tstB_25（{eqf}）", "#,##0.000")
    put("a25", "FY2025 SPE 營收（實績）", "¥B", "2025-04～2026-03", "=spe_25", "Verified", "In 頁 spe_25（公司決算說明資料 127,878 百萬日圓）", "#,##0.000")
    put("ref", "參考：層級隱含東京精密 SPE 營收（份額沿用 FY2025 實績、匯率不變）", "¥B", "CY2026", f"=D{inp['a25']}*D{inp['m26']}/D{inp['m25']}", "Derived（份額為假設）", "＝FY2025 SPE 實績 × 市場 CY2026 ÷ CY2025；份額不變為假設（非層級值）", "#,##0.0")
    put("anc", "參考：模型 FY2026 SPE 營收（公司指引錨點）", "¥B", "FY2026", f"=PL!D{RB['spe']}", "Derived（公司指引）", "PL 頁情境 B FY2026（＝公司 11-04 指引；上半為實績），不算實績", "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "¥B"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("參考：主讀數（份額沿用 FY2025 實績）", "2026-04～09 vs CY2026 × 6/12", f"=D{inp['h1']}", f"=D{inp['ref']}*D{inp['mo']}/12",
             "實績＝FY2026 Q1＋Q2 SPE 營收；層級＝測試與其他封裝設備 CY2026 ÷ CY2025 × FY2025 SPE 營收 × 6 ÷ 12（線性時差近似）。份額不變與匯率不變為假設，故只列參考", "參考（份額假設，不判定）")
    rb = rd_("參考：第二讀數（份額沿用 FY2025 實績）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q2']}*4", f"=D{inp['ref']}",
             "最近一季年化；季中點（8 月中旬）晚於日曆年中點，成長期偏高；只作時差近似的另一端參考", "參考（份額假設，不判定）")
    CW.cell(r, 1, "C0（實績）判定：SPE 營收"); CW.cell(r, 7, "無可比實績（層級無東京精密份額）")
    CW.cell(r, 8, "結論：無可比實績（層級無東京精密份額）：設備市場層只有測試與其他封裝設備等市場值，缺東京精密份額或探針台市場。參考列（份額沿用 FY2025 實績）的兩個讀數只供方向對照，份額不變為假設，不判定；對帳中維持 0")
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占模型 FY2026 SPE 營收（公司指引錨點）"); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['h1']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（上半實績 ÷ 全年指引），不是差距；指引不算實績")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, "讀法：公司實績是決算短信部門數字、獨立；層級沒有東京精密份額，用 FY2025 實績份額換算時份額不變是假設，不能判定。結論：無可比實績，數字僅供參考。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    r += 1
    CW.cell(r, 1, "其他公司實績（無可比層級值；FY2026 Q1–Q2 與 Q2）").font = BOLD; ROW["oth_t"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "FY2026 Q1–Q2 累計", "FY2026 Q2", "", "判定", "出處／口徑／缺少的揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_oth"] = r; r += 1
    others = [("SPE 受注（訂單）", "¥B", "=Q!G10+Q!H10", "=Q!H10", "無可比實績（層級無訂單）", f"Q 頁第 10 列；{TSQ2} 上半 114,074 百萬日圓（+86.0%）"),
              ("SPE 營業利益", "¥B", "=Q!G9+Q!H9", "=Q!H9", "無可比實績（層級無對應量）", f"Q 頁第 9 列；上半 17,717 百萬日圓"),
              ("計測機器營收", "¥B", "=Q!G11+Q!H11", "=Q!H11", "無可比實績（層級無計測市場）", f"Q 頁第 11 列；上半 19,753 百萬日圓"),
              ("合併營收", "¥B", "=Q!G4+Q!H4", "=Q!H4", "無可比實績（層級無東京精密營收）", f"Q 頁第 4 列；{TSQ2} 上半 92,399 百萬日圓"),
              ("合併營業利益", "¥B", "=Q!G5+Q!H5", "=Q!H5", "無可比實績（層級無對應量）", "Q 頁第 5 列；上半 20,341 百萬日圓")]
    ROW["oth0"] = r
    for nm, u, a9, a3, judge, miss in others:
        CW.cell(r, 1, nm); CW.cell(r, 2, u); CW.cell(r, 3, "2026-04～09／2026-07～09")
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
    s_exc, s_chh, s_gm, s_wts = sens_r("SPE 相對設備市場"), sens_r("中國 SPE 營收"), sens_r("計測機器營收年增"), sens_r("情境權重")
    sd = SNS["dcfw"]
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rb = PLA["rows_B"]
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    pc = lambda c: TX(c, "+0.0%;-0.0%")
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 日圓，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("：DCF "), TX(dcf, "#,##0"), T(" 日圓，同業比較法（相對倍數調整後）"), TX(cadj, "#,##0"), T(" 日圓，各占一半。"))
    put("rec", T("投資建議「"), rec, T("」的由來：同業比較法採東京精密過去十年相對同業的本益比（"), TX(coef, "0.00"), T(" 倍，低於同業），把同業中位數法的 "), TX(cu, "#,##0"), T(" 日圓壓到 "), TX(cadj, "#,##0"),
        T(" 日圓。不採相對倍數調整時加權目標價 "), TX(noadj, "#,##0"), T(" 日圓（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{sd}", "#,##0"),
        T(" 日圓（"), TX(f"Sens!C{sd}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{sd}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 日圓，比現價 "), TX("px", "#,##0"), T(" 日圓"), f'IF({up}<0,"低 ","高 ")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "#,##0"), T(" 日圓）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 日圓。"))
    put("o2", T("最大變數（加權目標價相對基準）：SPE 相對設備市場的超額成長 低／高 "), pc(f"Sens!B{s_exc}"), T("／"), pc(f"Sens!C{s_exc}"),
        T("；中國 SPE 營收變化 低／高 "), pc(f"Sens!B{s_chh}"), T("／"), pc(f"Sens!C{s_chh}"), T("；計測機器年增 低／高 "), pc(f"Sens!B{s_gm}"), T("／"), pc(f"Sens!C{s_gm}"),
        T("；情境權重全給保守 "), pc(f"Sens!B{s_wts}"), T("。這些是測試案例，不是門檻。"))
    put("o3", T("DCF 低於現價的原因：WACC "), TX("wacc", "0.0%"), T("（股權成本 "), TX("ke", "0.0%"), T("）與基準、保守情境 FY2028 起設備市場持平或下滑。要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"), T(" 倍（現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 十億日圓，模型 "), TX(rv("rv0"), "#,##0"), T("）。"))
    put("o4", T("對帳中：CY2027–CY2029 的 AI 設備（先進封裝、先進邏輯、DRAM 含 HBM），Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 日圓（"), f"Coord!C{CR['rec']}", T("），現行基準讀數 "), TX(f"Coord!B{CR['tp']}", "#,##0"), T("（"), f"Coord!B{CR['rec']}", T("）"),
        f'IF(Coord!B{CR["rec"]}=Coord!C{CR["rec"]},"，評等相同","，評等不同")', T("。個股不裁決。公司實績（SPE 營收 FY2026 Q1–Q2）對層級：無可比實績（層級無東京精密份額；標參考）。"))
    put("o5", T("進度：FY2026 上半營收 "), TX("Q!G4+Q!H4", "#,##0.0"), T(" 十億日圓，占公司 11 月指引 "), TX("g26_rev", "#,##0"), T(" 的 "), TX("(Q!G4+Q!H4)/g26_rev", "0%"),
        T("；SPE 受注上半 "), TX("Q!G10+Q!H10", "#,##0.0"), T(" 十億日圓，約為同期 SPE 營收的 "), TX("(Q!G10+Q!H10)/(Q!G8+Q!H8)", "0.0"), T(" 倍。下一個檢驗是第三季決算（2027-02 上旬）的 SPE 受注。"))
    put("drv", T("營收推導（情境 B）：探針台＝FY2026 錨點 × 測試與其他封裝設備需求指數 × 超額成長 × 中國係數；加工裝置＝錨點 × 加工裝置需求指數（前段與先進封裝加權）× 同係數；計測機器逐年成長。FY2028 探針台 "),
        TX(f"PL!F{rb['pr']}", "#,##0.0"), T("、加工裝置 "), TX(f"PL!F{rb['mc']}", "#,##0.0"), T("、計測 "), TX(f"PL!F{rb['met']}", "#,##0.0"), T("，合計 "), TX(f"PL!F{rb['rev']}", "#,##0.0"),
        T(" 十億日圓；SPE 營業利益率 "), TX(f"PL!F{rb['spm']}", "0.0%"), T("。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達、Advantest）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 HistIS 頁（標準損益表歷史，部分列未取得留空）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數；讀取層級；C0：SPE 營收 FY2026 Q1–Q2 實績對層級，層級無東京精密份額，結論無可比實績，對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, PB, EQ_F):
    # check_layout_v13 的兩讀數檢查要求 Coord 頁該列標籤恰為「加權目標價」（只改文字標籤，數值與公式不變）
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    CO = coord_c0(wb, SIL, CR, HR, PB, EQ_F)
    PLA = pl_standard(wb, HR, CR, PB)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((("C0 實績（v0.3.5）：FY2026 上半 SPE 營收、受注、營業利益與合併數字取自 2027年3月期 第2四半期（中間期）決算短信（2026-11-04）；Q1 部門數字取自 2026年度第1四半期 決算說明資料（本頁既有列）", ""),
                                ("HistIS（v0.3.5 起；v0.3.6 補齊）：營收、營業利益、淨利、EPS 取既有 Hist 頁；FY2025 毛利、銷管費取 In 頁 a25_gp／a25_sga；FY2023–FY2024 毛利、銷管費與 FY2023–FY2025 稅前利益、所得稅取公司 2025年3月期、2026年3月期 決算短信", FS26),
                                ("層級：設備市場模型 Mkt 頁 5B 節（測試與其他封裝、先進封裝〔AI〕、晶圓前段設備；先進邏輯〔AI〕、DRAM〔AI〕）", ""))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v035 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v035_post：HistIS 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 加權目標價相對基準 {SNS['t0']} 起；Coord 讀取層級與 C0 {CO['rd0']}–{CO['last']}；README v1.3 白話列 {len(RMA)} 列"
