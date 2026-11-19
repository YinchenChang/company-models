# BESI v0.2.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5、Advantest v0.7.5、ASML v0.3.5）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.2.3 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2023–FY2025（公司新聞稿；IFRS；會計年度＝日曆年；FY2024、FY2025 連 Q 頁）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（快照 besi/sens.json）
#   Coord 頁末：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）；Sources、README 頁末
import json, os, re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.4", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
NA = "未取得（本輪無法讀取原文；待補）"


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
CCY, UB = "歐元", "€M"


# ---------- Hist：損益表歷史（FY2024、FY2025 連 Q 頁；FY2023 只有營收） ----------
SRC_Q = "Q 頁（連格）：公司新聞稿 Q4-25／FY2025（2026-02-19；stocktitan 轉載全文，網址見 Sources 頁）；stockanalysis 年度損益核對（Q 頁來源2 欄）"
SRC_23 = "公司年報 FY2023 營收 579.0（€M；本模型 DCF 參數 fc_da／fc_cx 的營收分母，In 頁說明）"


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.2.4 起；PL 頁「標準損益表」的歷史三年用）。會計年度＝日曆年。金額：百萬歐元（€M）；EPS：歐元。IFRS（公司新聞稿）。"
               "FY2024、FY2025 連 Q 頁（公司新聞稿實績）；營業成本、營業費用、稅前利益、所得稅、股數為由新聞稿各列相減或相除的 Derived 值（公司新聞稿另列銷管與研發，本頁未拆）。"
               "FY2023 只有營收（公司年報）；FY2023 其他各列與 FY2022 營收本輪未能讀取原文，標「未取得」、留空，不以第三方比率推估")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    R = {}
    rows = [("rev", "營收（Revenue）", 579.0, "=Q!B4", "=Q!G4", "Verified", SRC_Q + "；FY2023：" + SRC_23),
            ("gp", "毛利（Gross profit）", None, "=Q!B5", "=Q!G5", "Verified", SRC_Q),
            ("cogs", "營業成本（Cost of sales；＝營收 − 毛利）", None, "=D4-D5", "=E4-E5", "Derived", "營收 − 毛利"),
            ("oi", "營業利益（Operating income）", None, "=Q!B8", "=Q!G8", "Verified", SRC_Q),
            ("opx", "營業費用（銷管＋研發；＝毛利 − 營業利益）", None, "=D5-D7", "=E5-E7", "Derived", "毛利 − 營業利益（新聞稿另列銷管、研發，本頁合計）"),
            ("fin", "財務費用淨額（Financial expense, net）", None, "=Q!B9", "=Q!G9", "Verified", SRC_Q),
            ("pti", "稅前利益（＝營業利益 − 財務費用淨額）", None, "=D7-D9", "=E7-E9", "Derived", "營業利益 − 財務費用淨額"),
            ("ni", "稅後淨利（Net income）", None, "=Q!B10", "=Q!G10", "Verified", SRC_Q + "；FY2024 含 €18.2M 稅務利益（Q 頁說明）"),
            ("tax", "所得稅（＝稅前利益 − 稅後淨利）", None, "=D10-D11", "=E10-E11", "Derived", "稅前利益 − 稅後淨利（FY2024 所得稅低，因含一次性稅務利益）"),
            ("eps_b", "EPS 基本（公司公告值）", None, "=Q!B11", "=Q!G11", "Verified", SRC_Q),
            ("sh_b", "基本加權平均股數（百萬股；＝淨利 ÷ 基本 EPS）", None, "=D11/D13", "=E11/E13", "Derived", "淨利 ÷ 基本 EPS（EPS 兩位小數，股數為近似）")]
    r = 4
    for k, labl, v23, v24, v25, tag, src in rows:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in ((3, v23), (4, v24), (5, v25)):
            if v is None: continue
            c = H.cell(r, j, v); c.number_format = "0.00" if k == "eps_b" else "#,##0.0"
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        H.cell(r, 6, tag); H.cell(r, 7, src if k == "rev" else src + "；FY2023：" + NA)
        R[k] = r; r += 1
    assert (R["rev"], R["gp"], R["oi"], R["fin"], R["pti"], R["ni"], R["tax"], R["eps_b"], R["sh_b"]) == (4, 5, 7, 9, 10, 11, 12, 13, 14), R
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    H.cell(r, 2, "營收 − 營業成本 − 毛利"); [H.cell(r, j, f"={c}4-{c}6-{c}5") for j, c in ((4, "D"), (5, "E"))]; R["chk0"] = r; r += 1
    H.cell(r, 2, "毛利 − 營業費用 − 營業利益"); [H.cell(r, j, f"={c}5-{c}8-{c}7") for j, c in ((4, "D"), (5, "E"))]; r += 1
    H.cell(r, 2, "FY2025 稅後淨利 − Q 頁四季合計（Q!C10:F10；應約 0）"); H.cell(r, 5, "=E11-SUM(Q!C10:F10)"); r += 1
    H.cell(r, 2, "FY2025 營收 − Q 頁四季合計（Q!C4:F4；應約 0）"); H.cell(r, 5, "=E4-SUM(Q!C4:F4)"); r += 1
    for rr in range(R["chk0"], r):
        for c in "DE": H[f"{c}{rr}"].number_format = "0.000"
    H.column_dimensions["A"].width = 10; H.column_dimensions["B"].width = 56
    for c in "CDE": H.column_dimensions[c].width = 12
    H.column_dimensions["F"].width = 12; H.column_dimensions["G"].width = 110
    return R


# ---------- PL：標準損益表（€M） ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "C", "D": "D", "E": "E"}   # 標準損益表歷史欄 → Hist 欄
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}
PC4 = {"F": "D", "G": "E", "H": "F", "I": "G"}   # 標準損益表 FY2026–29 → PL 推導區塊欄（C＝2025 … G＝2029）
EC = {"J": "H", "K": "I", "L": "J"}   # FY2030–32 → PL 延伸區塊欄
DCOL = {"G": "C", "H": "D", "I": "E", "J": "F", "K": "G", "L": "H"}   # → DCF 頁欄（FY2027＝C … FY2032＝H）


def pl_standard(wb, HR, CR, PB):
    P = wb["PL"]; R = PB["B"]
    xb = lab(P, "情境（延伸）")
    xs = lab(P, "情境 B（基準）", xb, exact=True)
    X = dict(rev=lab(P, "營收", xs, exact=True), ni=lab(P, "淨利（收斂法延伸）", xs), eps=lab(P, "EPS（延伸）", xs))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, "標準損益表（v0.2.4 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加第四季指引的錨點）。會計年度＝日曆年；百萬歐元；EPS 歐元。"
                 "只引用本頁上方、Hist、Q、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：IFRS（國際財務報導準則；公司新聞稿；BESI 不公布 Non-GAAP EPS）。EPS 為基本 EPS（與模型、共識同口徑）。預測期模型推毛利、營業費用（固定＋變動）、營業利益與淨利（淨利＝（營業利益 − 財務費用淨額）×（1 − 稅率））；"
                 "其他收入（淨額）＝ − 財務費用淨額，稅前利益＝營業利益 − 財務費用淨額，所得稅＝稅前利益 − 淨利。FY2023 歷史只有營收（其他列本輪未取得原文，留空）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷（1 − 稅率）＋財務費用淨額），毛利率沿用 FY2029，營業費用＝毛利 − 營業利益（推算）。FY2026E 的財務費用＝前三季實績＋第四季假設。")
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
    hh = lambda k, c3=True: [(f"=Hist!C{HR[k]}" if c3 else None), f"=Hist!D{HR[k]}", f"=Hist!E{HR[k]}"]
    fc = lambda k: [f"=PL!{PC4[c]}{R[k]}" for c in "FGHI"]
    fin_f = ["=Q!H9+Q!I9+Q!J9+q4fin", "=fin", "=fin", "=fin", "=fin", "=fin", "=fin"]
    row("rev", "營收", UB, hh("rev") + fc("rev") + [f"=PL!{EC[c]}{X['rev']}" for c in "JKL"], M, "FY2026E＝錨點；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [None] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 年增率：FY2022 營收本輪未取得，留空")
    rc = r
    row("cogs", "營業成本（COGS）", UB, [None, f"=Hist!D{HR['cogs']}", f"=Hist!E{HR['cogs']}"] + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", UB, [None, f"=Hist!D{HR['gp']}", f"=Hist!E{HR['gp']}"] + fc("gp") + [f"={c}{L['rev']}*{c}{rc + 2}" for c in "JKL"], M, "FY2026E–FY2029E 取 PL 情境 B；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [None] + [f"={c}{rc + 1}/{c}{L['rev']}" for c in "DEFGHI"] + [f"={p}{rc + 2}" for p in "IJK"], PCT, "FY2030E–FY2032E 沿用前一年")
    row("opx", "營業費用（銷管＋研發）", UB, [None, f"=Hist!D{HR['opx']}", f"=Hist!E{HR['opx']}"] + fc("opx") + [f"={c}{rc + 1}-{c}{rc + 4}" for c in "JKL"], M,
        "FY2026E 為錨點；FY2027E–FY2029E 為模型固定＋變動；FY2030E 起＝毛利 − 營業利益（推算）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", UB, [None, f"=Hist!D{HR['oi']}", f"=Hist!E{HR['oi']}"] + fc("oi") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M, "FY2030E 起＝DCF 頁情境 B 的 EBIT（淨利 ÷（1 − tax）＋ fin）")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [None] + [f"={c}{L['oi']}/{c}{L['rev']}" for c in "DEFGHIJKL"], PCT)
    row("oth", "其他收入（淨額；＝ − 財務費用淨額）", UB, [None, f"=-Hist!D{HR['fin']}", f"=-Hist!E{HR['fin']}"] + ["=-(" + f[1:] + ")" for f in fin_f], M, "財務費用淨額取負；FY2026E＝前三季實績（Q 頁）＋第四季假設；FY2027E 起＝In 頁 fin")
    row("pti", "稅前利益", UB, [None] + [f"={c}{L['oi']}+{c}{L['oth']}" for c in "DEFGHIJKL"], M, "＝營業利益 ＋ 其他收入")
    row("tax", "所得稅", UB, [None] + [f"={c}{L['pti']}-{c}{r + 1}" for c in "DEFGHIJKL"], M, "＝稅前利益 − 稅後淨利（預測期＝稅前利益 × 稅率）")
    row("ni", "稅後淨利", UB, [None, f"=Hist!D{HR['ni']}", f"=Hist!E{HR['ni']}"] + fc("ni") + [f"=PL!{EC[c]}{X['ni']}" for c in "JKL"], M)
    row("nm", "淨利率", "%", [None] + [f"={c}{L['ni']}/{c}{L['rev']}" for c in "DEFGHIJKL"], PCT)
    row("sh", "基本加權平均股數（百萬股）", "百萬股", [None, f"=Hist!D{HR['sh_b']}", f"=Hist!E{HR['sh_b']}"] + ["=shares"] * 4 + [f"={c}{L['ni']}/{c}{r + 1}" for c in "JKL"], "#,##0.0",
        "歷史＝淨利 ÷ 基本 EPS（Derived）；FY2026E–FY2029E＝In 頁 shares；FY2030E 起＝淨利 ÷ EPS（延伸區塊）")
    row("eps", "EPS（IFRS 基本）", "€", [None, f"=Hist!D{HR['eps_b']}", f"=Hist!E{HR['eps_b']}"] + fc("eps") + [f"=PL!{EC[c]}{X['eps']}" for c in "JKL"], "#,##0.00",
        "基本 EPS；共識 EPS 同口徑")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" if c != "C" else None for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" if c != "C" else None for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" if c != "C" else None for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" if c != "C" else None for c in COLS], PCT)
    row("r_tax", "有效稅率（所得稅 ÷ 稅前利益）", "%", [None] + [f"={c}{L['tax']}/{c}{L['pti']}" for c in "DEFGHIJKL"], PCT, "預測期＝In 頁 tax；FY2024 含一次性稅務利益")
    row("r_niy", "淨利年增率（YoY）", "%", [None, None] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS[1:], COLS[2:])], PCT, "FY2023 淨利未取得，FY2023、FY2024 留空")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", UB, [None] + [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "DEFGHIJKL"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", UB, [None] + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "DEFGHIJKL"], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", UB, [None] + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in "DEFGHIJKL"], "0.000")
    row("c4", "核對：營業利益 − DCF 頁情境 B 的 EBIT（FY2027E–FY2032E）", UB, [None] * 4 + [f"={c}{L['oi']}-DCF!{DCOL[c]}{d_ebit}" for c in "GHIJKL"], "0.000")
    row("c5", "核對：稅後淨利 − PL 情境 B（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", UB, [None] * 3 + [f"={c}{L['ni']}-PL!{PC4[c]}{R['ni']}" for c in "FGHI"] + [f"={c}{L['ni']}-PL!{EC[c]}{X['ni']}" for c in "JKL"], "0.000")
    row("c6", "核對：預測期稅率 − In 頁 tax（FY2027E–FY2032E）", "%", [None] * 4 + [f"={c}{L['tax']}/{c}{L['pti']}-tax" for c in "GHIJKL"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度＝日曆年，引擎對 FY(t) 直接取 CY(t) 的 HBM 位元、CoWoS 晶圓與 Siliconomics 年別比值；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2032E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, rows_X=X, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF（DCF 頁 FY2027–FY2032＝C–H 欄） ----------


def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 6, 7, 8
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數")
    assert D.cell(5, 3).value == "FY2027" and D.cell(5, 8).value == "FY2032", (D.cell(5, 3).value, D.cell(5, 8).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.4 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
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
    kv("sh", "評價用稀釋股數（百萬股）", "=sh_v", "#,##0.00")
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
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 百萬歐元，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(f" {CCY}相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × BESI 過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給 BESI 的倍數是同業的 "),
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
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.2.4 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
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
    SE.cell(r, 4, "v0.2.4 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（besi/sens_besi.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
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
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 besi/sens_besi.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR, PB, EQ_F):
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]; N = SIL.NAMES
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    eqf = os.path.relpath(EQ_F, ROOT); eq_ver = re.search(r"v\d+(?:\.\d+)+", os.path.basename(EQ_F)).group(0)
    em = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
    rb_ = next(r for r in range(4, 40) if str(em.cell(r, 1).value or "").strip() == "情境 B")
    ry = next(r for r in range(rb_ + 1, rb_ + 6) if str(em.cell(r, 1).value or "").strip() == "年增")
    EQB = {y: em.cell(ry, 3 + y - 2025).value for y in (2027, 2028, 2029)}
    pkg = N["SIF_SEMI_Pkg"]
    assert abs(pkg[1] - 6.7) < 1e-9 and abs(N["SIF_HBMHybrid_Base"][1]) < 1e-12, (pkg, N["SIF_HBMHybrid_Base"])
    RB = PB["B"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.2.4；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（讀 M22 HBM 堆疊技術 hbm_den／hbm_stk／hbm_hbp、SIF_WTAnchor_Base、SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（後段鍵合設備），讀 Siliconomics"),
          ("HBM 出貨位元（三情境）", "唯讀上游", "markets/memory 經 SK 海力士模型成品 Mkt 頁「出貨＝MIN(需求 × 情境顆數係數, 供給)」（情境 B 與 Siliconomics SIF_HBM_ShipEB_Base 同值）", "skhynix dist", "In 頁 A 節 hs<情境>_<年>；HBM 混合鍵合段"),
          ("CoWoS 晶圓需求（基準路徑）與情境顆數係數", "唯讀上游", "台積電模型成品 CoWoS 頁「需求合計」、In 頁 um28／um29（Siliconomics SIF_UnitsCoef_Base 為基準）", "tsmc dist", "In 頁 C 節；邏輯混合鍵合段"),
          ("設備總額年增（三情境）與見頂年度", "唯讀上游", f"{eqf}（Mkt 頁）", eq_ver, "In 頁 F 節 eqg<情境>_<年>；主流業務與循環位置"),
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
        CW.cell(r, 6, f'=IF(E{r}=0,IF(D{r}=0,0,"—"),D{r}/E{r}-1)').number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
    nb = f"Siliconomics {SIL.VERSION}"
    for i, y in enumerate(range(2025, 2030)):
        line("HBM 出貨位元（情境 B；In 頁讀入）vs Siliconomics", "EB", f"CY{y}", f"=hsB_{y % 100}", N["SIF_HBM_ShipEB_Base"][i], "0.000",
             f"{nb} SIF_HBM_ShipEB_Base（Derived；markets/memory 出貨）；本模型經 SK 海力士成品讀入")
    for i, y in enumerate(range(2025, 2030)):
        line("混合鍵合占 HBM 工序比例（基準）vs Siliconomics", "%", f"CY{y}", f"=hbp_{y % 100}", N["SIF_HBMHybrid_Base"][i], "0.0%", f"{nb} SIF_HBMHybrid_Base（M22 hbm_hbp，Assumed）；情境 B 讀基準，A 讀高值、C 讀低值")
    for y in (2026, 2028):
        line("DRAM 核心晶粒容量（基準）vs Siliconomics", "Gb", f"CY{y}", f"=den_{y % 100}", SIL.vals(f"hbm_den_{y % 100}")[1], "0.0", f"{nb} M22 hbm_den_{y % 100}（Assumed）")
        line("HBM 平均堆疊層數（基準）vs Siliconomics", "層", f"CY{y}", f"=stk_{y % 100}", SIL.vals(f"hbm_stk_{y % 100}")[1], "0.0", f"{nb} M22 hbm_stk_{y % 100}（Assumed）")
    for y in (2027, 2028, 2029):
        line("設備總額年增（情境 B；In 頁讀入）vs 設備市場模型成品", "%", f"CY{y}", f"=eqgB_{y % 100}", EQB[y], "0.0%", f"{eqf} Mkt 頁情境 B「年增」（Derived）；建置時直接讀成品原值")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：BESI 的份額（層級只有 SEMI 組裝與封裝設備市場值，沒有 BESI 份額或混合鍵合設備市場）；混合鍵合營收、HBM／邏輯拆分、機台單價與產出（公司不揭露，模型 Assumed）；主流業務（隨設備總額年增）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準（HBM 出貨位元、CoWoS 晶圓需求，情境 B）與 Wattonomics 換算值兩層之間的差距，不是 BESI 自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for nm, u, fm, f_ in (("HBM 出貨位元（情境 B）", "EB", "0.000", lambda y: f"=hsB_{y % 100}"), ("CoWoS 晶圓需求（情境 B，含顆數係數）", "千片", "#,##0", lambda y: f"=PL!{chr(ord('C') + y - 2025)}{PB['B']['cow']}")):
        for y in (2027, 2028, 2029):
            CW.cell(r, 1, f"{nm}：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, u); CW.cell(r, 3, f"CY{y}")
            CW.cell(r, 4, f_(y)).number_format = fm; CW.cell(r, 5, f"=D{r}*{ratio(y)}").number_format = fm
            CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
            CW.cell(r, 8, f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；對帳中年份依 C4 並列兩讀數目標價（本頁上方），個股不裁決")
            r += 1
    ROW["ref2"] = r - 1
    r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t)：FY2027–FY2029 及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；FY2026 為前三季實績加第四季指引的錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 歐元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 歐元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的 HBM 出貨位元與 CoWoS 晶圓需求 × 另一側比值；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數直接縮放 HBM 出貨位元（不再與供給取小）與 CoWoS 晶圓需求；邏輯段校準係數仍用未縮放的 FY2026 新機需求指數（FY2026 錨點是實績，不重新校準）；主流業務（設備總額年增）未縮放。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    # ---- 4. C0 ----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.2.4；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績：2026 年 Q1–Q3（2026-01～2026-09；BESI 會計年度＝日曆年）合併營收（公司季報新聞稿，IFRS），不用指引或預測。層級側：Siliconomics 只有 SEMI 組裝與封裝設備市場值（CY2025 SEMI 實績、CY2026 起 SEMI 預測），沒有 BESI 份額，也沒有混合鍵合設備市場。"
                  "主讀數＝已公布期間累計 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（Q3，2026-07～09）× 4 vs 層級 CY2026（季中點約 8 月中旬，晚於日曆年中點，營收逐季上升時偏高；不得寫成「夾住真值」）。"
                  "兩讀數差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」；層級沒有本公司份額時只標參考、不判定（tools/coord/README）。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查：SEMI 組裝與封裝設備市場值不由本模型校準（BESI 模型不讀此值），與 BESI 實績無同源關係；但層級沒有 BESI 份額，要得到「層級隱含 BESI 營收」只能假設份額。"
                  "下表以「份額沿用 CY2025 實績」（2025 營收 ÷ SEMI 2025 組裝與封裝設備，匯率不變）換算：層級隱含值＝2025 營收 × 市場 CY2026 ÷ CY2025。這列確實用到層級市場值，但份額不變是假設、不是層級值，故只列參考、不判定；結論「無可比實績（層級無 BESI 份額）」，對帳中維持 0。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "缺的獨立揭露：(1) 層級的混合鍵合設備市場值或 BESI 份額（第三方份額屬外部估計，只能列參考）；(2) BESI 混合鍵合營收、HBM／邏輯拆分（公司不揭露）；(3) 口徑差：SEMI 組裝與封裝設備含打線、封裝、切割等，BESI 營收另含電鍍（Meco），兩者範圍不同。"
                  "HBM 混合鍵合：層級 CY2026 混合鍵合占 HBM 工序 0%（Siliconomics M22），公司說 HBM 混合鍵合導入延後（Q3-26 法說，Interested-party）；公司不揭露金額，無可比實績。")
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
    put("q1", "營收：Q1 2026", "€M", "2026-01～2026-03", "=Q!H4", "Verified", "Q 頁 H4（連格）；184.9 取自公司 Q2-26／H1-26 新聞稿的前季數（2026-07-23；stocktitan 轉載全文，網址見 Sources 頁）", "#,##0.0")
    put("q2", "營收：Q2 2026", "€M", "2026-04～2026-06", "=Q!I4", "Verified", "Q 頁 I4（連格）；公司 Q2-26 新聞稿 249.9（2026-07-23；網址見 Sources 頁）", "#,##0.0")
    put("q3", "營收：Q3 2026（最近一季）", "€M", "2026-07～2026-09", "=Q!J4", "Verified", "Q 頁 J4（連格）；公司 Q3-26 新聞稿 283.4（2026-10-22；網址見 Sources 頁）", "#,##0.0")
    put("ytd", "營收：2026 年 Q1–Q3 累計", "€M", "2026-01～2026-09", f"=D{r - 3}+D{r - 2}+D{r - 1}", "Verified（三季相加）", "三季相加", "#,##0.0")
    put("mo", "已公布月數", "月", "2026-01～2026-09", 9, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("pk26", "層級值：SEMI 組裝與封裝設備市場 CY2026", "US$B", "CY2026", pkg[1], "Interested-party（層級值；SEMI 預測）", f"Siliconomics {SIL.VERSION} SIF_SEMI_Pkg（快照 {SIL.FILE}；建置時讀入並核對 6.7）", "#,##0.000")
    put("pk25", "層級值：SEMI 組裝與封裝設備市場 CY2025（SEMI 實績）", "US$B", "CY2025", pkg[0], "Verified（層級值）", f"Siliconomics {SIL.VERSION} SIF_SEMI_Pkg 首年（由 SEMI 2026 值與年增回推）", "#,##0.000")
    put("a25", "2025 營收（實績）", "€M", "2025-01～2025-12", "=a25_rev", "Verified", "Q 頁 G4；公司 Q4-25／FY2025 新聞稿 591.3", "#,##0.0")
    put("ref", "參考：層級隱含 BESI 營收（份額沿用 CY2025 實績、匯率不變）", "€M", "CY2026", f"=D{inp['a25']}*D{inp['pk26']}/D{inp['pk25']}", "Derived（份額為假設）", "＝2025 實績 × 市場 CY2026 ÷ CY2025；份額不變為假設（非層級值）", "#,##0.0")
    put("anc", "參考：模型 FY2026 營收錨點（前三季實績＋第四季指引）", "€M", "FY2026", "=rev26", "Derived（含公司指引）", "In 頁 rev26；第四季為公司指引（持平 ±5%）中點，不算實績", "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "€M"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("參考：主讀數（份額沿用 CY2025 實績）", "2026-01～09 vs CY2026 × 9/12", f"=D{inp['ytd']}", f"=D{inp['ref']}*D{inp['mo']}/12",
             "實績＝Q1＋Q2＋Q3；層級＝SEMI 組裝與封裝設備 CY2026 ÷ CY2025 × 2025 營收 × 9 ÷ 12（線性時差近似）。份額不變與匯率不變為假設，故只列參考", "參考（份額假設，不判定）")
    rb = rd_("參考：第二讀數（份額沿用 CY2025 實績）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp['ref']}",
             "最近一季年化；季中點（8 月中旬）晚於日曆年中點，成長期偏高；只作時差近似的另一端參考", "參考（份額假設，不判定）")
    CW.cell(r, 1, "C0（實績）判定：營收"); CW.cell(r, 7, "無可比實績（層級無 BESI 份額）")
    CW.cell(r, 8, "結論：無可比實績（層級無 BESI 份額）：Siliconomics 只有 SEMI 組裝與封裝設備市場值，缺 BESI 份額或混合鍵合設備市場。參考列顯示 BESI 2026 年營收成長遠快於該市場（混合鍵合與 AI 相關訂單），即份額沿用 CY2025 的假設不成立，不是層級值的問題；對帳中維持 0")
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占模型 FY2026 營收錨點（個股自身）"); CW.cell(r, 2, "€M"); CW.cell(r, 3, "2026-01～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['ytd']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（前三季實績 ÷ 錨點），不是差距；錨點含第四季公司指引，非實績")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, "讀法：公司實績是季報新聞稿數字、獨立；層級沒有 BESI 份額，用 CY2025 實績份額換算時份額不變是假設，不能判定。結論：無可比實績，數字僅供參考。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    r += 1
    CW.cell(r, 1, "其他公司實績（無可比層級值；2026 年 Q1–Q3 與 Q3）").font = BOLD; ROW["oth_t"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "2026 年 Q1–Q3 累計", "Q3 2026", "", "判定", "出處／口徑／缺少的揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_oth"] = r; r += 1
    others = [("訂單（order intake）", "€M", "=Q!H12+Q!I12+Q!J12", "=Q!J12", "無可比實績（層級無訂單）", "Q 頁第 12 列（公司新聞稿 269.7／292.9／342.6）；近 12 月訂單約 1.2 十億歐元"),
              ("毛利", "€M", "=Q!H5+Q!I5+Q!J5", "=Q!J5", "無可比實績（層級無對應量）", "Q 頁第 5 列；Q3 為營收 × 毛利率 65.2%（Derived）"),
              ("營業利益", "€M", "=Q!H8+Q!I8+Q!J8", "=Q!J8", "無可比實績（層級無對應量）", "Q 頁第 8 列（公司新聞稿）"),
              ("淨利", "€M", "=Q!H10+Q!I10+Q!J10", "=Q!J10", "無可比實績（層級無對應量）", "Q 頁第 10 列（公司新聞稿）"),
              ("混合鍵合累計訂單台數（2025 年底 → 2026 年第三季）", "台", "=hb_cum26q3-hb_cum25", "=hb_cum26q3", "無可比實績（層級無台數）",
               "D 欄＝2026 年前三季新增（215＋ − 150＋），E 欄＝2026 年第三季累計；In 頁 hb_cum25（Q4-25 新聞稿，Verified）、hb_cum26q3（Q3-26 簡報摘要，Interested-party）")]
    ROW["oth0"] = r
    for nm, u, a9, a3, judge, miss in others:
        CW.cell(r, 1, nm); CW.cell(r, 2, u); CW.cell(r, 3, "2026-01～09／2026-07～09")
        for col, v in ((4, a9), (5, a3)):
            c = CW.cell(r, col, v); c.number_format = "#,##0" if u == "台" else "#,##0.0"
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
    s_att, s_hb, s_shm, s_wts = sens_r("邏輯混合鍵合導入比例"), sens_r("FY2026 混合鍵合占營收"), sens_r("BESI 占 HBM"), sens_r("情境權重")
    sd = SNS["dcfw"]
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rb = PLA["rows_B"]
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    pc = lambda c: TX(c, "+0.0%;-0.0%")
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.4 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0.0"), T(" 歐元，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("：DCF 只有 "), TX(dcf, "#,##0.0"), T(" 歐元，同業比較法 "), TX(cadj, "#,##0.0"), T(" 歐元，各占一半。"))
    put("rec", T("投資建議「"), rec, T("」的由來：DCF 遠低於現價（WACC 偏高、基準與保守情境混合鍵合放量慢），同業比較法採 BESI 過去十年相對同業的本益比（"), TX(coef, "0.00"), T(" 倍）把加權拉回現價附近。不採相對倍數調整時加權目標價 "),
        TX(noadj, "#,##0.0"), T(" 歐元（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{sd}", "#,##0.0"),
        T(" 歐元（"), TX(f"Sens!C{sd}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{sd}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0.0"), T(" 歐元，比現價 "), TX("px", "#,##0.0"), T(" 歐元"), f'IF({up}<0,"低 ","高 ")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.00"), T(" 歐元）得 "), TX(cadj, "#,##0.0"), T("，DCF 得 "), TX(dcf, "#,##0.0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0.0"), T(" 歐元高出許多。"))
    put("o2", T("最大變數是混合鍵合放量（加權目標價相對基準）：邏輯混合鍵合導入比例 ×0.7／×1.3 "), pc(f"Sens!B{s_att}"), T("／"), pc(f"Sens!C{s_att}"),
        T("；FY2026 混合鍵合占營收 18%／35% "), pc(f"Sens!B{s_hb}"), T("／"), pc(f"Sens!C{s_hb}"), T("；HBM 混合鍵合份額 ±10 個百分點 "), pc(f"Sens!B{s_shm}"), T("／"), pc(f"Sens!C{s_shm}"),
        T("；情境權重全給保守 "), pc(f"Sens!B{s_wts}"), T("。這些是測試案例，不是門檻。"))
    put("o3", T("DCF 偏低的原因：WACC "), TX("wacc", "0.0%"), T("（股權成本 "), TX("ke", "0.0%"), T("）偏高、基準與保守情境 FY2028 起營收成長慢。要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"), T(" 倍（現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 百萬歐元，模型 "), TX(rv("rv0"), "#,##0"), T("）。"))
    put("o4", T("對帳中：CY2027–CY2029 的 HBM 出貨位元與 CoWoS 晶圓需求，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0.0"), T(" 歐元（"), f"Coord!C{CR['rec']}", T("），現行基準讀數 "), TX(f"Coord!B{CR['tp']}", "#,##0.0"), T("（"), f"Coord!B{CR['rec']}", T("）"),
        f'IF(Coord!B{CR["rec"]}=Coord!C{CR["rec"]},"，評等相同","，評等不同")', T("。個股不裁決。公司實績（營收 2026 年 Q1–Q3）對層級：無可比實績（層級無 BESI 份額；標參考）。"))
    put("o5", T("進度：2026 年前三季營收 "), TX("Q!H4+Q!I4+Q!J4", "#,##0.0"), T(" 百萬歐元，占模型 FY2026 錨點 "), TX("rev26", "#,##0.0"), T(" 的 "), TX("(Q!H4+Q!I4+Q!J4)/rev26", "0%"),
        T("；前三季訂單 "), TX("Q!H12+Q!I12+Q!J12", "#,##0.0"), T(" 百萬歐元，混合鍵合累計訂單由 2025 年底 "), TX("hb_cum25", "0"), T(" 台增至第三季 "), TX("hb_cum26q3", "0"), T(" 台以上（公司說法）。下一個檢驗是第四季財報的混合鍵合訂單。"))
    put("drv", T("營收推導（情境 B）：混合鍵合＝HBM 段（HBM 位元 → 晶粒 → 混合鍵合機台 × 單價 × 份額）＋邏輯段（CoWoS 晶圓 × 導入比例的新機需求 × 校準係數 × 份額）；主流業務隨設備總額年增。FY2028 混合鍵合 "),
        TX(f"PL!F{rb['hbr']}", "#,##0"), T("、主流業務 "), TX(f"PL!F{rb['main']}", "#,##0"), T("，合計 "), TX(f"PL!F{rb['rev']}", "#,##0"), T(" 百萬歐元；混合鍵合占營收 "), TX(f"PL!F{rb['hsh']}", "0.0%"), T("。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達、Advantest、ASML）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史，IFRS；FY2023 只有營收）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數；讀取層級；C0：營收 2026 年 Q1–Q3 實績對層級，層級無 BESI 份額，結論無可比實績，對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, PB, SUMR, EQ_F):
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
    for i, (t, u) in enumerate((("C0 實績（v0.2.4）：Q1–Q3 2026 營收、訂單取自 Q 頁（公司季報新聞稿；Q1-26 數字見 Q2-26／Q3-26 新聞稿的前季欄）；Q3-26 新聞稿 2026-10-22（investing.com 簡報摘要，本頁第 2 列）", ""),
                                ("Hist（v0.2.4）：FY2024、FY2025 損益取自 Q4-25／FY2025 新聞稿（stocktitan 轉載全文，本頁第 6 列）；FY2023 營收取自公司年報（DCF 參數說明）；FY2023 其他列與 FY2022 未取得", ""),
                                ("層級：Siliconomics v0.22 SIF_SEMI_Pkg（SEMI 組裝與封裝設備，CY2025 實績、CY2026 起 SEMI 預測）、SIF_HBMHybrid_Base、SIF_HBM_ShipEB_Base", "siliconomics/data/" + SIL.FILE))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v024 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v024_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 加權目標價相對基準 {SNS['t0']} 起；Coord 讀取層級與 C0 {CO['rd0']}–{CO['last']}；README v1.3 白話列 {len(RMA)} 列"
