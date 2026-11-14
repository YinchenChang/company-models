# Astera Labs v0.3.8（2026-11-14）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調（讀取層級、C0 公司實績、兩讀數；規則 v1.4.2；
# 工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；樣板 NVIDIA v0.3.10、Broadcom v0.5.7）。只新增，不改既有數值格：
#   Hist 頁末：損益表歷史（FY2023–FY2025 non-GAAP 與 GAAP；Astera Labs 8-K 新聞稿、S-1）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；Non-GAAP＝模型口徑）、比率區、GAAP 對照、核對列
#   RevDCF 頁（新）：反向 DCF；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（引用既有目標價參數表與業務主因表，加估值框架封閉式案例）
#   Coord 頁 F–I 節：讀取層級、層級之間（參考）、損益表年度標示、C0（實績）
#   README 頁末：v1.3 展示層白話列
# build_alab.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。環境變數 COORD_RUN 設定時（alt_run.py 的建置）不執行（add_sheet 回傳 None，本檔不會被呼叫）。
import json, os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3.8", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
EDG = "https://www.sec.gov/Archives/edgar/data/1736297/"
SRC_Q424 = EDG + "000173629725000001/q424exhibit991.htm（FY2024 Q4 與全年 8-K 新聞稿；FY2023 為比較欄）"
SRC_Q425 = EDG + "000173629726000005/q425exhibit991.htm（FY2025 Q4 與全年 8-K 新聞稿）"
SRC_S1 = EDG + "000119312524040419/d285484ds1.htm（S-1，summary consolidated financial data：FY2022 營收）"
SRC_Q1 = EDG + "000173629726000033/q226exhibit991.htm（2Q26 8-K 新聞稿；1Q26 為比較欄）"
SRC_Q2 = EDG + "000173629726000033/q226exhibit991.htm（2Q26 8-K 新聞稿）"
SRC_Q3 = EDG + "000173629726000044/q326exhibit991.htm（3Q26 8-K 新聞稿，2026-11-03；九個月累計營收 1,285,660 千美元）"


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
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- Hist：損益表歷史（標準損益表用） ----------
# (鍵, 項目, FY2023, FY2024, FY2025, FY2022（基期）, 標記, 來源)；金額十億美元、EPS 美元、股數十億股
HIS = [
    ("rev", "營收", 0.115794, "=a24_rev", "=a25_rev", 0.079872, "Verified", f"FY2023：{SRC_Q424}；FY2024、FY2025：Q 頁（In 頁 a24／a25 列，{SRC_Q425}）；FY2022：{SRC_S1}"),
    ("gp", "毛利（non-GAAP）", 0.079851, "=a24_gp", "=a25_gp", None, "Verified", f"FY2023：{SRC_Q424}；FY2024、FY2025：Q 頁"),
    ("opx", "營業費用（non-GAAP；＝毛利 − 營業利益）", 0.098669, "=a24_opx", "=a25_opx", None, "Derived", f"FY2023：{SRC_Q424}（non-GAAP 營業費用 98,669 千美元，release 不單列、由毛利與營業損失推得）；FY2024、FY2025：Q 頁"),
    ("ninc", "利息與其他收益（non-GAAP；Astera 為利息收入）", 0.006549, "=a24_ninc", "=a25_ninc", None, "Verified", f"FY2023：{SRC_Q424}（利息收入 6,549 千美元）；FY2024、FY2025：Q 頁"),
    ("ni", "稅後淨利（non-GAAP）", -0.015578, "=a24_ni", "=a25_ni", None, "Verified", f"FY2023：{SRC_Q424}（non-GAAP 淨損 15,578 千美元）；FY2024、FY2025：Q 頁"),
    ("sh", "稀釋股數（non-GAAP；十億股；FY2023、FY2024 為上市前後擬制股數）", 0.128022, "=a24_sh", "=a25_sh", None, "Verified", f"FY2023：{SRC_Q424}（擬制稀釋股數 128,022 千股）；FY2024、FY2025：Q 頁"),
    ("eps", "EPS（non-GAAP 稀釋，公司公告；美元）", -0.12, "=a24_eps_in", "=a25_eps_in", None, "Verified", f"FY2023：{SRC_Q424}（−0.12）；FY2024、FY2025：In 頁 a24／a25 列"),
    ("g_oi", "營業利益（GAAP）", -0.029497, -0.116066, "=h_goi_25", None, "Verified", f"FY2023、FY2024：{SRC_Q424}；FY2025：Hist 頁 h_goi_25（{SRC_Q425}）"),
    ("g_ni", "稅後淨利（GAAP）", -0.026257, -0.083421, 0.219134, None, "Verified", f"FY2023、FY2024：{SRC_Q424}；FY2025：{SRC_Q425}（In 頁 a25 列備註：GAAP 淨利 219,134 千美元）"),
    ("g_eps", "EPS（GAAP 稀釋；美元）", -0.71, -0.64, 1.22, None, "Verified", f"FY2023、FY2024：{SRC_Q424}；FY2025：{SRC_Q425}（In 頁 a25 列備註：GAAP 稀釋 EPS 1.22）"),
]


def hist_block(wb):
    H = wb["Hist"]
    r = H.max_row + 2
    H.cell(r, 1, "損益表歷史（v0.3.8 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。金額十億美元；EPS 美元；股數十億股。non-GAAP 為公司口徑（排除股權報酬、IPO 相關費用等）；FY2024、FY2025 欄連到 Q 頁／In 頁的既有實績。").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2022（只供營收年增率基期）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {}
    for k, labl, v23, v24, v25, v22, tag, src in HIS:
        rows[k] = r
        H.cell(r, 1, f"hs_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25), ("F", "22", v22)):
            if v is None: continue
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000" if k not in ("eps", "g_eps") else "0.00"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"hs_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, tag); H.cell(r, 8, src)
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    for labl, f in (("毛利 − 營業費用 − 營業利益（FY2023；營業損失 −18.818 百萬美元，release 揭露值）", f"=C{rows['gp']}-C{rows['opx']}-(-0.018818)"),
                    ("營業利益（毛利 − 營業費用）＋利息與其他收益 − 稅後淨利 ＝ 所得稅（FY2023 應為 3.309 百萬美元）− release 值", f"=C{rows['gp']}-C{rows['opx']}+C{rows['ninc']}-C{rows['ni']}-0.003309")):
        H.cell(r, 1, labl); H[f"C{r}"] = f; H[f"C{r}"].number_format = "0.000000"; r += 1
    H.cell(r, 1, "核對：FY2024 營收、毛利連到 Q 頁；FY2025 營收 − Q 頁（應為 0）"); H[f"E{r}"] = f"=E{rows['rev']}-a25_rev"; r += 1
    H.cell(r, 1, "備註：FY2023、FY2024 non-GAAP 稀釋股數為公司公告的擬制股數（2024-03 上市前後），EPS 為淨利 ÷ 擬制股數；FY2023 為 non-GAAP 淨損，淨利年增率與以負值為基期者標「不適用」。").font = Font(italic=True)
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"      # 本表欄；Hist C–E＝FY2023–FY2025；預測欄 F–L＝PL 情境 B 區塊 D–J


def pl_standard(wb):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    rc = lab(P, "C. 合計", rb)
    g = lambda lbl, s=rb: lab(P, lbl, s, exact=True)
    R = dict(rev=g("營收", rc), gp=g("毛利", rc), opx=g("營業費用", rc), oi=g("營業利益", rc), oth=g("利息與其他收益", rc), tr=g("有效稅率", rc), ni=g("淨利", rc), sh=g("稀釋股數", rc), eps=g("EPS（non-GAAP）", rc))
    keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.3.8 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026E 錨點、FY2027E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "只引用本頁上方與 Hist、Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。Astera Labs 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股權報酬、無形資產攤銷與一次性項目；股權報酬約占營業費用三成），故本表用 Non-GAAP。"
                 "其他收入（淨額）＝利息與其他收益（利息收入為主）。GAAP 營業利益、稅後淨利與 EPS 列在本表下方對照（只列實績）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（1Q26–3Q26 實績＋4Q26 公司指引）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2032E＝延伸期（收斂法只推營收與淨利率，營業利益＝淨利 ÷（1 − 稅率）− 利息與其他收益，同 DCF 頁）。"
                 "延伸期毛利率與營業費用率沿用 FY2029E（v0.3.8 決定），據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r_hdr = r; r += 1
    L = {k: r + i for i, k in enumerate(keys)}      # 預先配置列號（公式需要往後引用）
    M = "#,##0.000"; PCT = "0.0%"

    def row(key, label, unit, cells, fmt, note=""):
        rr = L[key]
        P.cell(rr, 1, label); P.cell(rr, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{rr}"] = v; P[f"{col}{rr}"].number_format = fmt
        if note: P.cell(rr, 13, note)

    fc = lambda key, n=7: [f"=PL!{c}{R[key]}" for c in "DEFGHIJ"[:n]]       # FY2026E–FY2032E（PL 情境 B 區塊 D–J 欄）
    hh = lambda k: [f"=hs_{k}_23", f"=hs_{k}_24", f"=hs_{k}_25"]
    row("rev", "營收", "US$B", hh("rev") + fc("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/hs_rev_22-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 營收（Hist 頁，S-1）")
    row("cogs", "營業成本（COGS）", "US$B", [f"={c}{L['rev']}-{c}{L['gp']}" for c in "CDEFGHI"] + [f"={c}{L['rev']}*(1-$I{L['gm']})" for c in "JKL"], M,
        "FY2023A–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp", 4) + [f"={c}{L['rev']}-{c}{L['cogs']}" for c in "JKL"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx", 4) + [f"={c}{L['rev']}*$I{L['opx']}/$I{L['rev']}" for c in "JKL"], M,
        "FY2030E–FY2032E＝營收 × FY2029E 營業費用率（non-GAAP，不含股權報酬）")
    row("oi", "營業利益", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}" for c in "CDE"] + fc("oi", 4) + [f"={c}{L['ni']}/(1-tx_29)-oth_f" for c in "JKL"], M,
        "FY2023A–FY2029E＝毛利 − 營業費用；FY2030E–FY2032E 由淨利反推（同 DCF 頁；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝利息與其他收益", "US$B", hh("ninc") + fc("oth", 4) + ["=oth_f"] * 3, M, "non-GAAP；利息收入為主")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", [f"={c}{L['pti']}-{c}{L['ni']}" for c in COLS], M, "稅前 − 稅後淨利（歷史＝公司揭露推導；預測＝模型有效稅率）")
    row("ni", "稅後淨利", "US$B", hh("ni") + fc("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", hh("sh") + fc("sh"), "0.0000")
    row("eps", "EPS（完全稀釋，non-GAAP）", "US$", [f"={c}{L['ni']}/{c}{L['sh']}" for c in "CDE"] + fc("eps"), "0.00",
        "FY2023A–FY2025A＝淨利 ÷ 稀釋股數（公司公告 −0.12、0.84、1.84；四捨五入差）")
    r = L["eps"] + 1
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    X = {}
    for key, labl, cells, fm, note in (
            ("r_yoy", "營收年增率", [f"={c}{L['yoy']}" for c in COLS], PCT, ""),
            ("r_gm", "毛利率", [f"={c}{L['gm']}" for c in COLS], PCT, ""),
            ("r_opm", "營業利益率", [f"={c}{L['opm']}" for c in COLS], PCT, ""),
            ("r_nm", "淨利率", [f"={c}{L['nm']}" for c in COLS], PCT, ""),
            ("r_niy", "淨利年增率（YoY）", ['="不適用"'] + [f'=IF({p}{L["ni"]}<=0,"不適用",{c}{L["ni"]}/{p}{L["ni"]}-1)' for p, c in zip(COLS, COLS[1:])], PCT,
             "基期為負值（FY2023 non-GAAP 淨損）者不適用；FY2023 的基期 FY2022 淨利未取得，標不適用")):
        L[key] = r
        P.cell(r, 1, labl); P.cell(r, 2, "%")
        for col, v in zip(COLS, cells):
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fm
        if note: P.cell(r, 13, note)
        r += 1
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    for key, labl, unit, fm, note in (("g_oi", "GAAP：營業利益", "US$B", M, "含股權報酬；FY2024 含 IPO 相關一次性股權報酬"), ("g_ni", "GAAP：稅後淨利", "US$B", M, ""), ("g_eps", "GAAP：EPS（完全稀釋，公司公告）", "US$", "0.00", "")):
        L[key] = r
        P.cell(r, 1, labl); P.cell(r, 2, unit)
        for col, y in zip("CDE", ("23", "24", "25")):
            P[f"{col}{r}"] = f"=hs_{key}_{y}"; P[f"{col}{r}"].number_format = fm
        if note: P.cell(r, 13, note)
        r += 1
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1

    def chk(key, labl, unit, cells, fm):
        nonlocal r
        L[key] = r
        P.cell(r, 1, labl); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fm
        r += 1
    chk("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    chk("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None] * 3, "0.000")
    chk("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "US$B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JKL"], "0.000")
    chk("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JKL"], PCT)
    chk("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    chk("c4", "核對：稅後淨利 − 情境 B 區塊（FY2026E–FY2032E）；FY2025A − Q 頁", "US$B",
        [None, None, f"=E{L['ni']}-a25_ni"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJKL", "DEFGHIJ")], "0.000")
    chk("c5", "核對：EPS − 情境 B 區塊（FY2026E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHIJKL", "DEFGHIJ")], "0.000")
    P.column_dimensions["A"].width = max(P.column_dimensions["A"].width or 0, 62)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    P.column_dimensions["K"].width = max(P.column_dimensions["K"].width or 0, 12); P.column_dimensions["L"].width = max(P.column_dimensions["L"].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R, last=r - 1)
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
    tvr = lab(D, "終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"     # 情境 B 欄；含 stub 期 FCFF 與已宣告股利（stub FCFF 與 WACC 無關）
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.8 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（評價日後第一個完整年度）維持模型值不動，只反解 FY2028 起；反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("nc", "淨現金滾到評價日（情境 B；US$B）", f"={ncr}", "#,##0.000", "Val 頁：淨現金與投資＋stub 期 FCFF − 已宣告股利（與 WACC 無關）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.0000")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.00")
    kv("ev", "模型企業價值 EV（情境 B；US$B）", f"=DCF!$C${ev}", "#,##0.00")
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
    yr("rev", "營收（US$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.00")
    yr("ebit", "營業利益 EBIT（GAAP 口徑，扣股權報酬；US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.000")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元 EBIT 對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.000")
    yr("wc", "營運資金增加（US$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.000")
    yr("pv", "現值（US$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期 EBIT 率（FY2028–FY2032 統一一個 EBIT 率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.000")
    kv("E0", "與營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.000")
    kv("m0", "模型：基準的等效 EBIT 率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型 EBIT 率")
    kv("m1", "現價隱含的 EBIT 率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；EBIT 率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（US$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.000", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（US$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.000")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.00")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.00")
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
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%；若下一列核對不為零，表示到下限仍低於現價所需值（折現率單獨無法支持現價）")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（美元；接近 0＝已解出，為負＝搜尋到下限仍不足以支持現價）", f"=E{r - 4}-B{A['v0']}", "0.000000")
    kv("w1t", "現價隱含的 WACC（顯示）", f'=IF(B{A["w_res"]}<-0.01,"無解：到搜尋下限 "&TEXT(B{A["w1"]},"0.0%")&" 仍不足",TEXT(B{A["w1"]},"0.0%"))', "@")
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
                     T(" 倍。單靠一個因素達成時，需要："), f'IF(B{A["w_res"]}<-0.01,"折現率即使降到搜尋下限 "&TEXT(B{A["w1"]},"0.0%")&" 仍不足","折現率降到 "&TEXT(B{A["w1"]},"0.0%"))', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0.0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0.0")', T("）；或 EBIT 率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T("；Astera Labs 上市不足十年、資料不足時係數為 1，即不調整）。換言之，市場目前給 Astera Labs 的倍數是同業中位數的 "),
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
    dcf, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r_px = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.3.8 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "美元", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0%")', T("／B "), 'TEXT(tp_wB,"0%")', T("／C "), 'TEXT(tp_wC,"0%")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；進加權；相對倍數調整係數 1.00，未調整）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T("（資料不足，未調整；與市場主流同業法同值）× 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 美元"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable "), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（賣方平均）", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{r_px}"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px 列）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節） ----------
def sens_tor(wb):
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp = f"Val!$B${r_tp}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")               # B、C、D＝情境 A、B、C
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    dcfw, cadj = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}"
    epsr = {k: lab(P, "EPS（non-GAAP）", lab(P, f"情境 {k}：") , exact=True) for k in "ABC"}   # 情境區塊內 EPS 列（各情境 FY2027＝E、FY2028＝F）
    fe = lambda k: f"(2/12*PL!$E${epsr[k]}+10/12*PL!$F${epsr[k]})"
    r_par = lab(SE, "目標價參數敏感度"); r_biz = lab(SE, "目標價敏感度：業務主因")
    rowof = lambda start, p: lab(SE, p, start)
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.3.8 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例引用上方「目標價敏感度：業務主因」表（整本重算快照，三情境同時套用）；"
                 "折現參數案例引用上方「目標價參數敏感度」表；估值框架案例（情境權重、DCF 權重）為封閉式").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    # 1) 估值框架案例（封閉式）
    frame = [
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fe('C')}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fe('A')}",
         "Val 頁：單一情境的 DCF 價值與 12 個月後前瞻 EPS（2／12 × FY2027＋10／12 × FY2028）"),
        ("DCF 在加權中的權重 100%（只看現金流折現）／0%（只看同業比較法）（基準＝50%）", f"={dcfw}", f"={cadj}", "Val 頁：DCF 目標價與 Comparable 目標價（相對倍數調整後）"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    # 2) 營運輸入案例（引用業務主因表；整本重算快照）與 3) 折現參數案例（引用目標價參數表）：(上方列前綴, 白話標籤)
    def link(start, prefix, labl, note):
        nonlocal r
        sr = rowof(start, prefix)
        SE.cell(r, 1, labl)
        for col, scol, tcol in ((2, "D", 4), (3, "F", 5)):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—" or (isinstance(v, str) and v.startswith("—")): continue
            SE.cell(r, col, f"={scol}{sr}").number_format = "+0.0%;-0.0%"
            SE.cell(r, tcol, f"={chr(ord(scol) - 1)}{sr}").number_format = "#,##0.00"
        SE.cell(r, 6, note + f"（本頁第 {sr} 列）"); r += 1
    biz = [("Scorpio X 份額", "Scorpio X 有效份額（FY2027 起相對值）低／高路徑，三情境同用"),
           ("PCIe 等份額", "PCIe 等有效份額（FY2027 起相對值）低／高路徑，三情境同用"),
           ("GAAP ÷ non-GAAP", "GAAP 與 non-GAAP 營業利益比 低／高（股權報酬占比）"),
           ("最大客戶轉單 15%", "最大終端客戶轉單 15%／50%（其 PCIe 等內容轉給其他供應商或自製）"),
           ("評價股數原口徑", "評價股數用公司 4Q26 稀釋股數指引的原口徑（只有低案）")]
    for pre, labl in biz:
        link(r_biz, pre, labl, "整本重算快照（alab/sens.json）；上方「目標價敏感度：業務主因」表")
    par = [("WACC ±1pt", "折現率 WACC −1／＋1 個百分點（RONIC 隨 WACC）"), ("β 低／高", "β 低／高（各自重新加槓桿）"), ("ERP 低／高", "股權風險溢酬 低／高"),
           ("永續成長率", "永續成長率 −1／＋1 個百分點"), ("RONIC＝WACC＋5pt", "永續期新投資報酬率由等於 WACC 改為高 5 個百分點（只有高案）"),
           ("資本支出", "資本支出占營收比例 −1／＋1 個百分點"), ("營運資金", "營運資金占營收比例 低／高"), ("終值改用", "終值改用最後一年自由現金流成長（替代算法，只有高案）")]
    for pre, labl in par:
        link(r_par, pre, labl, "上方「目標價參數敏感度」表")
    last = r - 1
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f'=D{lab(SE, "模型狀態（CURRENT")}'); SE.cell(r, 3, "同上方「模型治理」節：CURRENT＝快照與本檔輸入一致；STALE 時須重跑 python3 alab/sens_alab.py"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調——讀取層級、層級之間（參考）、損益表年度標示、C0（實績） ----------
def coord_sections(wb, SIL):
    C, I, PLs = wb["Coord"], wb["In"], wb["PL"]
    N = SIL.NAMES
    ys = (2026, 2027, 2028, 2029)
    pct = '+0.0%;-0.0%'
    r_flag = lab(C, "對帳中旗標")
    r = C.max_row + 2
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    # ---- F. 讀取層級 ----
    C.cell(r, 1, "F. 讀取層級（三層基準協調規則 v1.4.2；工作單 20261114_第二批13家）：本模型讀哪一層、哪一版、哪些名稱。自層級讀入的同源值只列在此節，不進 C0 判定").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("情境顆數係數 CY2028、CY2029（um28_A／B／C、um29_A／B／C；基準＝情境 B）", "Siliconomics（SIF_UnitsCoef_Base 等；經 In 頁 um28_*／um29_* 讀入）", vtxt, "PL 頁情境 A／B／C：市場與 AI 加速器顆數 × 係數（CY2028 起）", '="情境 B："&TEXT(um28_B,"0.00")&"／"&TEXT(um29_B,"0.000")'),
        ("非 NVIDIA scale-up 交換晶片市場 msx_26–msx_29（Scorpio X 可服務市場）", "市場層 markets/network（AI 網通市場模型，經 In 頁 A 節讀入）", I.cell(lab(I, "msx_26", col=1), 8).value.split("｜")[0].split("/")[-1], "Scorpio X 營收＝市場 × 有效份額 × 相對值", '="CY2026 "&TEXT(msx_26,"0.00")&" 十億美元；CY2029 "&TEXT(msx_29,"0.00")'),
        ("AI 加速器顆數 u_26–u_29（PCIe 等可服務市場＝顆數 × c_pcie；c_pcie 為本模型 Assumed）", "市場層 markets/network（顆數取自 markets/memory）", "同上", "PCIe 與訊號調節營收＝可服務市場 × 有效份額 × 相對值", '="CY2026 "&TEXT(u_26,"0.0")&" 百萬顆；CY2029 "&TEXT(u_29,"0.0")'),
        ("AI IT 資本支出 itc_26–itc_29（Recon 對照用）", "市場層 markets/network（Mkt 頁）", "同上", "Recon：Astera 營收 ÷ AI IT 資本支出（只作對照，不進損益）", '="CY2026 "&TEXT(itc_26,"#,##0")&" 十億美元"'),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>）", "Siliconomics", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中（另一側讀數＝現行基準 × 0.62／0.67／0.67）"),
        ("Wattonomics 錨定係數 wta_27–wta_29（舊敏感度案例；SIF_WTAnchor_Base）", "Wattonomics（經 Siliconomics 轉入）", vtxt, "Sens 頁「Wattonomics 錨定」案例（情境 B × 係數）；兩讀數的另一側讀數取代其角色", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")'),
        ("Astera Labs 份額、Scorpio X 與 PCIe 等營收假設（x_26、rel_x_*、rel_p_*、c_pcie）", "本專案為權威（Units 頁；規則第 13 條）", "本模型", "不讀自層級；有效份額由 FY2026 錨點校準（見 I 節「同源校準」）", "Assumed"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。上列輸入的原值與來源見 In 頁；本節不改任何輸入。").font = Font(italic=True); r += 2
    # ---- G. 層級之間（參考） ----
    C.cell(r, 1, "G. 層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：產業部署 GW，Siliconomics 現行基準 vs 另一側讀數（Wattonomics 換算）。兩層之間的差距，不是 Astera Labs 自己的量對層級值；本公司不標、不計入對帳中項目數，依 C4 並列兩讀數目標價").font = BOLD; r += 1
    for j, h in enumerate(["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    for j, y in enumerate(ys):
        C.cell(r, 1, f"產業部署 GW CY{y}")
        C.cell(r, 2, N["SIF_DeployGW_Base"][j]).number_format = "#,##0.00"; C.cell(r, 3, N[f"SIF_DeployGW_Alt_{y}"]).number_format = "#,##0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
        C.cell(r, 5, "層間參考（Siliconomics 標對帳中；本公司不計）" if SIL.RECO_FLAG[y] else "層間參考（兩層一致）")
        C.cell(r, 6, "GW；SIF_DeployGW_Base vs SIF_DeployGW_Alt_<年>")
        C.cell(r, 7, f"Siliconomics {SIL.VERSION}（旗標 SIF_RecoFlag_{y}＝{int(SIL.RECO_FLAG[y])}）；Siliconomics 對帳單 r3")
        C.cell(r, 8, "Derived（基準）／Wattonomics 換算（另一側）"); C.cell(r, 9, "兩層之間的差距，不裁決；本模型 Coord 頁 A–E 節並列兩讀數目標價")
        r += 1
    r += 1
    # ---- H. 損益表年度標示 ----
    C.cell(r, 1, "H. 對帳中年份與損益表年度標示（C4；會計年度＝日曆年，FY(t)＝CY(t)；FY2030E–FY2032E 為 FY2029E 之後的延伸期，依賴 CY2029）").font = BOLD; r += 1
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in range(2026, 2033)], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年")
    dep = [2026, 2027, 2028, 2029, 2029, 2029, 2029]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2026E–FY2032E）")
    for j, fy in enumerate(range(2026, 2033)):
        cy = dep[j]
        C.cell(r, 2 + j, f'=IF({"$" + "BCDE"[cy - 2026] + "$" + str(r_flag)}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    # ---- I. C0（實績） ----
    C.cell(r, 1, "I. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家 第 2 節）").font = BOLD; j_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝Astera Labs 8-K 已公布的 2026 年季度營收（1Q26、2Q26、3Q26；公司揭露，Verified；公司指引不算實績，只列參考）。層級側＝同口徑的層級隱含值＝市場層值 × 份額："
                 "非 NVIDIA scale-up 交換晶片市場（markets/network，CY2026）× Scorpio X 有效份額＋AI 加速器顆數 × 每顆 PCIe 等內容（本模型 Assumed）× PCIe 等有效份額。"
                 "時差：Astera Labs 會計年度＝日曆年，1Q26–3Q26＝2026-01～2026-09（9 個月）。主讀數＝三季累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；"
                 "第二讀數＝最近一季（3Q26）× 4 vs 層級 CY2026。成長年中，主讀數偏低（後段季度較大）；3Q26 的季中點（2026-08）晚於年中，第二讀數偏高。"
                 "判定規則：兩讀數差距都超過 ±15% 且同號 → 對帳中；只有一個超過 → 時差邊界（參考）；都在內 → 一致。不用模型預測、同源值或賣方估計補位。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); r += 1
    C.cell(r, 1, "同源校準（本表實際檢驗什麼）：Astera Labs 沒有揭露產品線營收與市占，層級（markets/network）也沒有 Astera 份額的獨立來源。本模型的有效份額＝FY2026 錨點營收 ÷ 市場，"
                 "市場值（Scorpio X 可服務市場、顆數 × 每顆 PCIe 等內容）在「市場 × 份額」中消去，層級隱含值恆等於 FY2026 錨點（1Q26–3Q26 實績＋4Q26 公司指引）。"
                 "因此本表只檢驗「公司實績 vs 本模型錨點」的時差型態，不檢驗層級；判定標「參考（同源校準）」、不計入對帳中，C0 結論為「無可比實績（層級值同源）」。"
                 "缺的獨立揭露：公司未揭露 Scorpio X、Aries、Taurus、Leo 各產品線營收，或可交叉驗證份額的第三方出貨資料（公司只說 Scorpio 於 3Q26 成為最大產品線）。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000"):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); IN[key] = r; r += 1
    inp("q1", "營收：1Q26（公司 8-K）", "=q1_rev", "US$B", "2026-01-01～2026-03-31", "In 頁 q1_rev；" + SRC_Q1, "Verified", GREEN)
    inp("q2", "營收：2Q26（公司 8-K）", "=q2_rev", "US$B", "2026-04-01～2026-06-30", "In 頁 q2_rev；" + SRC_Q2, "Verified", GREEN)
    inp("q3", "營收：3Q26（公司 8-K）", "=q3_rev", "US$B", "2026-07-01～2026-09-30", "In 頁 q3_rev；" + SRC_Q3, "Verified", GREEN)
    inp("mo", "已公布月數（1Q26–3Q26）", 9, "月", "2026-01～2026-09", "會計年度＝日曆年；三季＝9 個月", "Derived", BLUE, fmt="0")
    inp("msx", "層級：非 NVIDIA scale-up 交換晶片市場 CY2026", "=msx_26", "US$B", "CY2026（日曆年）", "In 頁 msx_26（來源欄所列 markets/network 網通市場模型儲存格；顆數 × 每顆 scale-up 交換晶片內容）", "Derived（層）", GREEN)
    inp("shx", "本模型：Scorpio X 有效份額（FY2026 校準）", f"=PL!D{lab(PLs, 'Astera Scorpio X 有效份額（FY2026 校準）', lab(PLs, '情境 B：'))}", "%", "FY2026＝CY2026", "PL 頁情境 B 區塊（＝Scorpio X 營收假設 x_26 ÷ 市場；x_26 為 Assumed，公司未揭露）", "Derived（Assumed 輸入反推；同源校準）", GREEN, "0.0%")
    inp("mp", "層級：PCIe／CXL／AEC 可服務市場 CY2026（顆數 × 每顆內容 c_pcie）", "=u_26*c_pcie/1000", "US$B", "CY2026（日曆年）", "In 頁 u_26（markets/network 顆數）× c_pcie（本模型 Assumed）", "Derived（層顆數）／Assumed（c_pcie）", GREEN)
    inp("shp", "本模型：PCIe 等有效份額（FY2026 校準）", f"=PL!D{lab(PLs, 'Astera PCIe 等有效份額（FY2026 校準）', lab(PLs, '情境 B：'))}", "%", "FY2026＝CY2026", "PL 頁情境 B 區塊（＝PCIe 等營收殘差 ÷ 可服務市場）", "Derived（錨點殘差反推；同源校準）", GREEN, "0.0%")
    inp("Lx", "層級隱含：Scorpio X CY2026（市場 × 有效份額）", f"=B{IN['msx']}*B{IN['shx']}", "US$B", "CY2026（日曆年）", "上兩列相乘", "Derived（同源校準）")
    inp("Lp", "層級隱含：PCIe 與訊號調節 CY2026（可服務市場 × 有效份額）", f"=B{IN['mp']}*B{IN['shp']}", "US$B", "CY2026（日曆年）", "上兩列相乘", "Derived（同源校準）")
    inp("L", "層級隱含：營收 CY2026（Scorpio X＋PCIe 與訊號調節）", f"=B{IN['Lx']}+B{IN['Lp']}", "US$B", "CY2026（日曆年）", "上兩列加總", "Derived（同源校準）")
    inp("an", "參考：FY2026 錨點營收（1Q26–3Q26 實績＋4Q26 指引中點；本模型）", "=a26_rev", "US$B", "FY2026＝CY2026", "Q 頁 FY2026 錨點", "Verified（三季）／Interested-party（4Q26 指引）", GREEN)
    inp("g", "參考：公司 4Q26 營收指引中點（不判定）", "=q4_rev", "US$B", "4Q26（指引，2027-02 公布實績）", "In 頁 q4_rev（3Q26 8-K 展望；Interested-party）", "Interested-party", GREEN)
    inp("chk", "核對：層級隱含 − FY2026 錨點（同源，應為 0）", f"=B{IN['L']}-B{IN['an']}", "US$B", "—", "市場值消去的證明：層級隱含恆等於錨點", "—", fmt="0.000000")
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    J0 = r
    b = IN
    SAME = "參考（同源校準）"
    C.cell(r, 1, "主讀數：營收，1Q26–3Q26 累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12")
    C.cell(r, 2, f"=B{b['q1']}+B{b['q2']}+B{b['q3']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, SAME)
    C.cell(r, 6, "US$B；公司實績＝三季營收加總；層級＝CY2026 全年 × 9 ÷ 12")
    C.cell(r, 7, "公司：Astera Labs 8-K（In 頁 q1_rev、q2_rev、q3_rev）；層級：markets/network 市場 × 本模型校準份額（Siliconomics " + SIL.VERSION + " 顆數係數）")
    C.cell(r, 8, "Verified（公司）／Derived（層，同源校準）／Derived（線性時差）")
    C.cell(r, 9, f'="成長年中，線性分攤的層級值偏高（後段季度較大），主讀數偏低："&TEXT(B{r},"0.000")&" vs "&TEXT(C{r},"0.000")&"。同源校準，只說明錨點與實績的時差型態，不構成層級對帳"'); rm = r; r += 1
    C.cell(r, 1, "第二讀數：營收，最近一季（3Q26）× 4 vs 層級 CY2026")
    C.cell(r, 2, f"=B{b['q3']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, SAME)
    C.cell(r, 6, "US$B；3Q26 單季 × 4（年化）；層級＝CY2026 全年")
    C.cell(r, 7, "公司：Astera Labs 8-K（In 頁 q3_rev）；層級同上")
    C.cell(r, 8, "Verified（公司）／Derived（層，同源校準）／Derived（年化）")
    C.cell(r, 9, "3Q26 的季中點（2026-08）晚於年中，成長年中最近一季年化偏高；同源校準，不構成層級對帳"); rs = r; r += 1
    C.cell(r, 1, "機械結果（依判定規則；僅供參考，同源校準故不判定、不計入）")
    C.cell(r, 5, f'=IF(AND(ABS(D{rm})>0.15,ABS(D{rs})>0.15,SIGN(D{rm})=SIGN(D{rs})),"對帳中",IF(OR(ABS(D{rm})>0.15,ABS(D{rs})>0.15),"時差邊界（參考）","一致"))')
    C.cell(r, 9, f'="兩讀數反向（主讀數 "&TEXT(D{rm},"+0.0%;-0.0%")&"、第二讀數 "&TEXT(D{rs},"+0.0%;-0.0%")&"）：實績在層級隱含的兩側，不是同向偏離；即使不同源也不會判對帳中"'); rmech = r; r += 1
    C.cell(r, 1, "C0（實績）判定：營收 2026").font = BOLD
    C.cell(r, 5, "無可比實績（層級值同源）").font = BOLD
    C.cell(r, 6, "公司未揭露產品線營收與市占；層級無 Astera 份額的獨立來源；層級隱含值由本模型 FY2026 錨點校準，市場值消去")
    C.cell(r, 9, "不開對帳單、不改上游值；待公司揭露 Scorpio X 或各產品線營收，或有獨立的 scale-up 交換晶片市占資料時，才能對層級值做 C0 判定"); rv = r; r += 1
    C.cell(r, 1, "參考：FY2026 全年（1Q26–3Q26 實績＋4Q26 公司指引）vs 層級 CY2026（含指引，同源，不判定）")
    C.cell(r, 2, f"=B{b['an']}").number_format = "#,##0.000"; C.cell(r, 2).font = GREEN
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "參考（含指引，同源，不判定）")
    C.cell(r, 6, "US$B；Q 頁 FY2026 錨點（三季實績＋4Q26 指引中點）；會計年度＝日曆年")
    C.cell(r, 7, "Q 頁 FY2026 錨點；In 頁 q4_rev"); C.cell(r, 8, "Verified／Interested-party（4Q26 指引）")
    C.cell(r, 9, "差距恆為 0（市場值消去）；指引不算實績"); r += 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（I 節 C0（實績）判定；讀取層級、層級之間、參考列不計）"); C.cell(r, 2, f'=IF(E{rmech}="對帳中",IF(E{rv}="無可比實績（層級值同源）",0,1),0)'); cnt = r; r += 1
    C.cell(r, 1, "說明：F 節為讀取層級（同源，不判定）；G 節為層與層之間的差距（參考）；I 節為本公司唯一的 C0 實績節，結論「無可比實績（層級值同源）」，對帳中項目 0。層與層之間的差距只列參考，外部估計（共識、公司 2030 年 TAM 估計）只在 Recon 頁列參考，不判定。")
    r += 1
    for rr_ in (j_title + 1, j_title + 2):
        C.row_dimensions[rr_].height = 90
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, G0=G0, H_PL=H_PL, title=j_title, J0=J0, rm=rm, rs=rs, rmech=rmech, rv=rv, cnt=cnt, IN=IN, r_flag=r_flag)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, feps, coef = vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("相對倍數調整係數")
    rb = lab(P, "情境 B："); rc_ = lab(P, "C. 合計", rb)
    pr = lambda lbl, s=rc_: lab(P, lbl, s, exact=True)
    rev_b = pr("營收")
    x_b, p_b = lab(P, "Scorpio X 營收", rb, exact=True), lab(P, "PCIe 與訊號調節營收", rb, exact=True)
    msx_b, shx_b = lab(P, "非 NVIDIA scale-up 交換晶片市場 × 情境顆數係數", rb, exact=True), lab(P, "Scorpio X 有效份額", rb, exact=True)
    r27, r28e = lab(RC, "FY2027 EPS vs 共識"), lab(RC, "FY2028 EPS vs 共識")
    up_m, dn_m = lab(RC, "評等升一級所需的 Comparable 採用倍數"), lab(RC, "評等降一級所需的 Comparable 採用倍數")
    px_m = lab(RC, "支持現價所需的 Comparable 採用倍數")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    sb = lambda pre: lab(SE, pre, SEN["first"])
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.8 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元、同業比較法 "), f'TEXT({cadj},"0")', T(" 美元，加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0")',
        f'IF({up}<0," 美元低 "," 美元高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("。主因：現價給的倍數遠高於同業，且基準情境的 2028 年顆數下修。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：Comparable 採用同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍（相對倍數調整係數 "), f'TEXT({coef},"0.00")',
        T("，資料不足故未調整）；基準情境 CY2028 顆數 ×0.8；Scorpio X 與 PCIe 等的份額相對值走基準路徑。Comparable 倍數升到約 "), f'TEXT(Recon!B{up_m},"0.0")',
        T(" 倍評等升一級、降到約 "), f'TEXT(Recon!B{dn_m},"0.0")', T(" 倍降一級；要支持現價需約 "), f'TEXT(Recon!B{px_m},"0.0")', T(" 倍。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("）；同業比較法 "), f'TEXT({cadj},"0.00")', T(" 美元＝中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 前瞻 EPS "), f'TEXT({feps},"0.00")', T("。"))
    put("o2", T("現價隱含：要讓現金流折現法支持現價，"), f'IF({rv("w_res")}<-0.01,"折現率即使降到搜尋下限 "&TEXT({rv("w1")},"0.0%")&" 仍不足","折現率須降到 "&TEXT({rv("w1")},"0.0%"))', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0.0")', T("，模型 "), f'TEXT({rv("rv0")},"0.0")', T(" 十億美元）。"))
    put("o3", T("盈餘對共識：基準情境 FY2027 EPS "), f'TEXT(Recon!E{r27},"0.00")', T(" 美元，比共識 "), f'TEXT(Recon!B{r27},"0.00")', T(" "), f'IF(Recon!F{r27}>=0,"高 ","低 ")', f'TEXT(ABS(Recon!F{r27}),"0.0%")',
        T("；FY2028 "), f'TEXT(Recon!E{r28e},"0.00")', T(" 美元，比推導共識 "), f'TEXT(Recon!B{r28e},"0.00")', T(" "), f'IF(Recon!F{r28e}>=0,"高 ","低 ")', f'TEXT(ABS(Recon!F{r28e}),"0.0%")', T("（情境 B、C 的 2028 年顆數下修）。"))
    put("o4", T("最敏感的輸入：Scorpio X 份額相對值取低／高路徑，加權目標價 "), f'TEXT(Sens!B{sb("Scorpio X 有效份額")},"+0%;-0%")', T("／"), f'TEXT(Sens!C{sb("Scorpio X 有效份額")},"+0%;-0%")',
        T("；PCIe 等份額 "), f'TEXT(Sens!B{sb("PCIe 等有效份額")},"+0%;-0%")', T("／"), f'TEXT(Sens!C{sb("PCIe 等有效份額")},"+0%;-0%")', T("。Scorpio X FY2026 營收是假設（公司未揭露產品線）。"))
    put("o5", T("客戶集中與評等邊界：單一終端客戶逾七成（FY2025 10-K），其 PCIe 等內容轉單五成時加權目標價 "), f'TEXT(Sens!C{sb("最大終端客戶轉單")},"+0%;-0%")',
        T("；評等降一級只需 Comparable 倍數由 "), f'TEXT(Recon!C{dn_m},"0.0")', T(" 降到 "), f'TEXT(Recon!B{dn_m},"0.0")', T("。"))
    put("drv", T("分部營收推導（情境 B）：營收＝Scorpio X＋PCIe 與訊號調節。Scorpio X＝非 NVIDIA scale-up 交換晶片市場 × Astera 有效份額（FY2026 以假設營收校準）× 份額相對值；"
                 "PCIe 等＝AI 加速器顆數 × 每顆內容 × 有效份額 × 相對值（FY2026 為錨點殘差）。FY2028：市場 "), f'TEXT(PL!F{msx_b},"0.0")', T("、有效份額 "), f'TEXT(PL!F{shx_b},"0.0%")',
        T("，Scorpio X "), f'TEXT(PL!F{x_b},"0.00")', T("；PCIe 等 "), f'TEXT(PL!F{p_b},"0.00")', T("，合計 "), f'TEXT(PL!F{rev_b},"0.00")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型的情境顆數係數與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1']}", T("（見 Coord 頁 F 節）；市場層讀 markets/network。C0 以公司實績對層級隱含值（Coord 頁 I 節）：Astera Labs 8-K 的 2026 年前三季營收 "),
        f'TEXT(Coord!B{CO["rm"]},"0.00")', T(" 十億美元，對層級隱含依已公布月數折算 "), f'TEXT(Coord!C{CO["rm"]},"0.00")', T("，差 "), f'TEXT(Coord!D{CO["rm"]},"+0.0%;-0.0%")',
        T("；最近一季年化 "), f'TEXT(Coord!B{CO["rs"]},"0.00")', T("，差 "), f'TEXT(Coord!D{CO["rs"]},"+0.0%;-0.0%")', T("。層級隱含值由本模型 FY2026 錨點校準、市場值消去，判定「"), f"Coord!E{CO['rv']}",
        T("」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項。產業部署 GW 兩層差距另列參考、不計入；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（規則 v1.4.2；工作單 20261114_第二批13家）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁末損益表歷史、PL 頁末標準損益表、RevDCF 頁、"
                 "Val 0.7 四個目標價、Sens 末加權目標價相對基準；Coord 頁並列兩讀數目標價（對帳中年份 CY2027–CY2029，tools/coord）、F 節讀取層級、I 節 C0（實績）。"
                 "既有數值格、加權目標價與評等不變（逐格比對只有版號、build_id 與新增格）"); out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1":
        return "v038_post：略過（alt_run 建置）"
    RD = wb["README"]
    hist_block(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    val_g(wb)
    SEN = sens_tor(wb)
    CO = coord_sections(wb, SIL)
    RM = readme13(wb, RVA, PLA, SEN, CO)
    wb._v038 = dict(PLA=PLA, RVA=RVA, SEN=SEN, CO=CO, RM=RM)
    return f"v038_post：PL 標準損益表第 {PLA['start']} 列起；RevDCF；Sens 加權目標價相對基準第 {SEN['t0']} 列；Coord I 節第 {CO['title']} 列（判定第 {CO['rv']} 列）"
