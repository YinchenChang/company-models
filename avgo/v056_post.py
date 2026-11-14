# Broadcom v0.5.6（2026-11-14）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調新機制（共用工作單 D 節：讀哪一層哪一版、C0 由下往上對帳）。
# 樣板 NVIDIA v0.3.7／v0.3.8（nvidia/v037_post.py）。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史（FY2023–FY2025 non-GAAP 與 GAAP；Broadcom 8-K 新聞稿，SEC EDGAR）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；Non-GAAP＝模型口徑）、比率區、GAAP 對照、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（整本重算快照 avgo/sens.json＋Val 頁封閉式案例）
#   Coord 頁末 F 節：三層協調——讀哪一層哪一版、C0 由下往上對帳表、損益表年度標示（對帳中旗標列直接引用，不另設具名範圍）
#   README 頁末：v1.3 展示層白話列
# build_avgo.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。環境變數 COORD_RUN／AVGO_SENS 設定時（alt_run.py、sens_avgo.py 的建置）不執行。
import json, os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.5.6", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
EDG = "https://www.sec.gov/Archives/edgar/data/1730168/"
SRC = {"FY2025": EDG + "000173016825000116/avgo-11022025x8kxex99.htm（FY2025 Q4 與全年 8-K；FY2024 為比較欄）",
       "FY2024": EDG + "000173016824000125/avgo-11032024x8kxex99.htm（FY2024 Q4 與全年 8-K；FY2023 為比較欄）",
       "FY2023": EDG + "000173016823000093/avgo-10292023x8kxex99.htm（FY2023 Q4 與全年 8-K；FY2022 為比較欄，只用營收與 non-GAAP 淨利作年增率基期）"}


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


# ---------- Hist：損益表歷史 ----------
# (鍵, 項目, FY2023, FY2024, FY2025, FY2022（基期）, 標記)；金額十億美元、EPS 美元、股數十億股（2024-07 一拆十後）
HIS = [
    ("rev", "營收（GAAP＝non-GAAP）", 35.819, 51.574, "=a25_rev", 33.203, "Verified"),
    ("gp", "毛利（non-GAAP）", 26.757, 39.459, "=a25_gp", None, "Verified"),
    ("opx", "營業費用（non-GAAP）", 4.632, 8.723, "=a25_opx", None, "Verified"),
    ("oi", "營業利益（non-GAAP）", 22.125, 30.736, "=a25_oi", None, "Verified"),
    ("int", "利息費用（non-GAAP）", 1.622, 3.796, "=a25_int", None, "Verified"),
    ("oth", "其他收益淨額（non-GAAP）", 0.500, 0.418, "=a25_oth", None, "Verified"),
    ("tax", "所得稅（non-GAAP）", 2.625, 3.625, "=a25_tax", None, "Verified"),
    ("ni", "稅後淨利（non-GAAP）", 18.378, 23.733, "=a25_ni", 16.526, "Verified"),
    ("sh", "稀釋股數（non-GAAP；十億股）", 4.353, 4.877, "=a25_sh", None, "Verified"),
    ("g_gp", "毛利（GAAP）", 24.690, 32.509, 43.294, None, "Verified"),
    ("g_opx", "營業費用（GAAP）", 8.483, 19.046, 17.810, None, "Verified"),
    ("g_oi", "營業利益（GAAP）", 16.207, 13.463, 25.484, None, "Verified"),
    ("g_int", "利息費用（GAAP）", 1.622, 3.953, 3.210, None, "Verified"),
    ("g_oth", "其他收益淨額（GAAP）", 0.512, 0.406, 0.455, None, "Verified"),
    ("g_pti", "稅前利益（GAAP）", 15.097, 9.916, 22.729, None, "Verified"),
    ("g_tax", "所得稅（GAAP；FY2025 為淨利益）", 1.015, 3.748, -0.397, None, "Verified"),
    ("g_ni", "稅後淨利（GAAP；FY2024 含停業單位損失 0.273）", 14.082, 5.895, 23.126, None, "Verified"),
    ("g_eps", "EPS（GAAP 完全稀釋，公司公告；美元，一拆十後）", 3.30, 1.23, 4.77, None, "Verified"),
    ("g_sh", "稀釋股數（GAAP；十億股）", 4.272, 4.778, 4.853, None, "Verified"),
]


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 損益表歷史（v0.5.6 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。金額十億美元；EPS 美元；股數十億股。"
               "non-GAAP 為公司口徑（排除股票薪酬、無形資產攤銷、重組、一次性稅務項目）；FY2025 欄取 Q 頁，與 8-K 全年值一致。")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2022（只供年增率基期）", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; rows = {}
    for k, labl, v23, v24, v25, v22, tag in HIS:
        rows[k] = r
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25), ("F", "22", v22)):
            if v is None: continue
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"h_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, tag)
        H.cell(r, 8, f"FY2023：{SRC['FY2023']}；FY2024：{SRC['FY2024']}；FY2025：" + ("Q 頁（" + SRC["FY2025"] + "）" if isinstance(v25, str) else SRC["FY2025"]))
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    for labl, f in (("毛利 − 營業費用 − 營業利益（non-GAAP）", "={c}{gp}-{c}{opx}-{c}{oi}"),
                    ("營業利益 − 利息費用 ＋ 其他收益 − 所得稅 − 稅後淨利（non-GAAP）", "={c}{oi}-{c}{int}+{c}{oth}-{c}{tax}-{c}{ni}"),
                    ("GAAP：營業利益 − 利息費用 ＋ 其他收益 − 稅前利益", "={c}{g_oi}-{c}{g_int}+{c}{g_oth}-{c}{g_pti}"),
                    ("GAAP：毛利 − 營業費用 − 營業利益", "={c}{g_gp}-{c}{g_opx}-{c}{g_oi}")):
        H.cell(r, 1, labl)
        for c in "CDE":
            H[f"{c}{r}"] = f.format(c=c, **rows); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 1, "核對：FY2025 營收 − Q 頁（應為 0）"); H[f"E{r}"] = f"=E{rows['rev']}-a25_rev"; r += 1
    H.cell(r, 1, "備註：GAAP 稅後淨利＝稅前 − 所得稅；FY2024 另含停業單位損失 0.273（持續營業單位淨利 6.168），故 FY2024 稅前 − 所得稅 ≠ 稅後淨利。").font = Font(italic=True)
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 58
    for c in "CDEF": H.column_dimensions[c].width = 14
    H.column_dimensions["H"].width = 90
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"      # 本表欄；Hist C–E＝FY2023–FY2025；預測欄 F–I＝PL 情境 B 區塊 D–G；J–L＝PL 延伸表 H–J


def pl_standard(wb):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    rc = lab(P, "C. 報告值", rb)
    g = lambda lbl, s=rc: lab(P, lbl, s, exact=True)
    R = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), int=g("利息費用"), oth=g("其他收益"), ni=g("淨利（non-GAAP）"), sh=g("稀釋股數"), eps=g("EPS（non-GAAP）"))
    x0 = lab(P, "三情境延伸到 FY2032"); xb = lab(P, "情境 B", x0, exact=True)
    X = dict(rev=lab(P, "營收", xb, exact=True), ni=lab(P, "淨利（延伸表）", xb, exact=True), sh=lab(P, "稀釋股數", xb, exact=True), eps=lab(P, "EPS（延伸表）", xb, exact=True))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.5.6 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026E 錨點、FY2027E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "只引用本頁上方與 Hist、Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。Broadcom 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股票薪酬、無形資產攤銷、重組與一次性稅務項目），故本表用 Non-GAAP。"
                 "其他收入（淨額）＝其他收益 − 利息費用（Broadcom 的利息費用大，列入淨額）。GAAP 合計列在本表下方對照（只列實績）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（第一至第三季實績＋第四季指引）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2032E＝延伸期（收斂法只推營收與淨利，營業利益＝淨利 ÷（1 − 稅率）＋利息費用 − 其他收益，同 DCF 頁）。"
                 "延伸期毛利率與營業費用率沿用 FY2029（v0.5.6 決定），據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
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

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFG"]        # FY2026E–FY2029E（PL 情境 B 區塊 D–G 欄）
    ex = lambda key: [f"=PL!{c}{X[key]}" for c in "HIJ"]         # FY2030E–FY2032E（PL 延伸表情境 B 的 H–J 欄）
    hh = lambda k: [f"=h_{k}_23", f"=h_{k}_24", f"=h_{k}_25"]
    row("rev", "營收", "US$B", hh("rev") + fc("rev") + ex("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/h_rev_22-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "US$B", [f"=C{L['rev']}-C{r_c + 1}", f"=D{L['rev']}-D{r_c + 1}", f"=E{L['rev']}-E{r_c + 1}"]
        + [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "FGHI"] + [f"={c}{L['rev']}*(1-$I{r_c + 2})" for c in "JKL"], M,
        "FY2023A–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp") + [f"={c}{L['rev']}-{c}{L['cogs']}" for c in "JKL"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx") + [f"={c}{L['rev']}*$I{r_c + 3}/$I{L['rev']}" for c in "JKL"], M,
        "FY2030E–FY2032E＝營收 × FY2029E 營業費用率（non-GAAP，不含股票薪酬）")
    row("oi", "營業利益", "US$B", hh("oi") + fc("oi") + [f"={c}{r_c + 9}/(1-tx_f)+int_f-oth_f" for c in "JKL"], M,
        "FY2030E–FY2032E 由淨利反推（同 DCF 頁；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝其他收益 − 利息費用", "US$B", ["=h_oth_23-h_int_23", "=h_oth_24-h_int_24", "=h_oth_25-h_int_25"]
        + [f"=PL!{c}{R['oth']}-PL!{c}{R['int']}" for c in "DEFG"] + ["=oth_f-int_f"] * 3, M, "non-GAAP；利息費用為主")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", ["=h_tax_23", "=h_tax_24", "=h_tax_25"] + [f"={c}{L['pti']}-{c}{r_c + 9}" for c in "FGHIJKL"], M, "歷史＝non-GAAP 所得稅；預測＝稅前 − 稅後淨利（模型有效稅率）")
    assert L["tax"] == r_c + 8 and r == r_c + 9
    row("ni", "稅後淨利", "US$B", hh("ni") + fc("ni") + ex("ni"), M)
    assert L["ni"] == r_c + 9
    # 校正：上方公式以 r_c 偏移引用淨利列；r_c＋17 在此處必須等於淨利列（見下方斷言）
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", hh("sh") + fc("sh") + ex("sh"), M)
    row("eps", "EPS（完全稀釋）", "US$", [f"=C{L['ni']}/C{L['sh']}", f"=D{L['ni']}/D{L['sh']}", f"=E{L['ni']}/E{L['sh']}"] + fc("eps") + ex("eps"), "0.00",
        "FY2023A–FY2025A＝淨利 ÷ 稀釋股數（公司公告 FY2024 4.87、FY2025 6.82；四捨五入差）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/h_ni_22-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 non-GAAP 淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_oi", "GAAP：營業利益", "US$B", ["=h_g_oi_23", "=h_g_oi_24", "=h_g_oi_25"], M, "含股票薪酬、無形資產攤銷、重組")
    row("g_oth", "GAAP：其他收入（淨額）＝其他收益 − 利息費用", "US$B", ["=h_g_oth_23-h_g_int_23", "=h_g_oth_24-h_g_int_24", "=h_g_oth_25-h_g_int_25"], M)
    row("g_pti", "GAAP：稅前利益", "US$B", ["=h_g_pti_23", "=h_g_pti_24", "=h_g_pti_25"], M)
    row("g_tax", "GAAP：所得稅", "US$B", ["=h_g_tax_23", "=h_g_tax_24", "=h_g_tax_25"], M, "FY2024 含一次性非現金稅務 45 億美元；FY2025 為淨利益")
    row("g_ni", "GAAP：稅後淨利", "US$B", ["=h_g_ni_23", "=h_g_ni_24", "=h_g_ni_25"], M)
    row("g_eps", "GAAP：EPS（完全稀釋，公司公告）", "US$", ["=h_g_eps_23", "=h_g_eps_24", "=h_g_eps_25"], "0.00")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None] * 3, "0.000")
    row("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "US$B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JKL"], "0.000")
    row("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JKL"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 摘要／延伸表（FY2026E–FY2032E）；FY2025A − Q 頁", "US$B",
        [None, None, f"=E{L['ni']}-a25_ni"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"={c}{L['ni']}-PL!{s}{X['ni']}" for c, s in zip("JKL", "HIJ")], "0.000")
    row("c5", "核對：EPS − 情境 B 摘要／延伸表（FY2026E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHI", "DEFG")]
        + [f"={c}{L['eps']}-PL!{s}{X['eps']}" for c, s in zip("JKL", "HIJ")], "0.000")
    P.column_dimensions["A"].width = max(P.column_dimensions["A"].width or 0, 62)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R, X=X)
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
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"     # 情境 B 欄；含 stub 期 FCFF 與已宣告股利（stub FCFF＝FY2026 FCFF × 91／365＋FY2027 FCFF × 60／365，與 WACC 無關）
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.5.6 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027 視為錨點年不動（與 FY2026 錨點的校準相連、三情境相同）。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    yr("ebit", "營業利益 EBIT（扣股票薪酬；US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元 EBIT 對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
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
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或 EBIT 率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × Broadcom 過去十年相對同業的折價 "), f'TEXT({coef},"0.00")', T(" 倍）。換言之，市場目前給 Broadcom 的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("模型採用的歷史折價。")))
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
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.5.6 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
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


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節） ----------
def sens_tor(wb, here):
    SE, V = wb["Sens"], wb["Val"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")               # B、C、D＝情境 A、B、C
    r_eps = lab(V, "各情境 12 個月後的前瞻 EPS")                 # B、C、D
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.5.6 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（avgo/sens_avgo.py：三情境同時套用案例覆寫，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    base_f = f"{tp}"
    # 1) 估值框架案例（封閉式；與 Val 頁同式：加權＝DCF 權重 × DCF ＋（1 − DCF 權重）× Comparable）
    frame = [
        ("同業比較法的歷史折價係數 夾值下限 0.50／不調整 1.00（基準＝歷史折價）",
         f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*0.5", f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*1", "Val 頁：Comparable（未調整）× 係數；低＝夾值下限 0.5、高＝不調整"),
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$D${r_eps}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$B${r_eps}", "Val 頁：單一情境的 DCF 價值與前瞻 EPS"),
        ("DCF 在加權中的權重 0%（只看同業比較法）／100%（只看現金流折現）（基準＝50%）",
         f"={cadj}", f"={dcf}", "Val 頁：Comparable（調整後）與 DCF 目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    # 2) 營運輸入案例（整本重算快照）
    for k, d in J["biz"].items():
        SE.cell(r, 1, d["label"])
        SE.cell(r, 4, d["低"]["tp"]).number_format = "#,##0.00"; SE.cell(r, 5, d["高"]["tp"]).number_format = "#,##0.00"
        SE.cell(r, 4).font = BLUE; SE.cell(r, 5).font = BLUE
        SE.cell(r, 2, f"=D{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{base_f}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（avgo/sens.json）；覆寫同本頁上方「一次動一個輸入」案例，三情境同時套用"); r += 1
    # 3) 折現參數案例（引用既有表）
    skip = ("DCF 權重",)       # 已由 0%／100% 全區間案例涵蓋
    for sr in src:
        t = str(SE.cell(sr, 1).value)
        if t.startswith(skip): continue
        SE.cell(r, 1, t)
        for col, scol, tcol in ((2, "D", 4), (3, "F", 5)):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
            SE.cell(r, tcol, f"=Sens!{chr(ord(scol) - 1)}{sr}").number_format = "#,##0.00"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「目標價參數敏感度」表）"); r += 1
    last = r - 1
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f'=IF(ABS({tp}-{J["_base"]["tp"]!r})<=1E-6*ABS({tp}),"CURRENT","STALE")')
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 avgo/sens_avgo.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調——讀哪一層哪一版、C0 由下往上對帳（共用工作單 D 節） ----------
def coord_c0(wb, SIL, T_):
    if "Coord" not in wb.sheetnames:
        return None
    C, I = wb["Coord"], wb["In"]
    r = C.max_row + 2
    ys = (2026, 2027, 2028, 2029)
    r_flag = lab(C, "對帳中旗標")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    C.cell(r, 1, "F. 三層協調（共用工作單 D 節，2026-11-14）：本模型讀哪一層、哪一版").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    net_v = src_in("net_26").split("｜")[0]
    hbm_v = src_in("hbm_B_27").split("｜")[0]
    u_v = src_in("u_tpu_26")
    rows = [
        ("產業顆數：Google TPU、其他 ASIC 2026–2028（u_tpu_*、u_oth_*；2029 成長 g29_*）", "Siliconomics 轉自市場層 markets/memory", vtxt, "XPU 顆數 × Broadcom 份額 → XPU 營收", "見 In 頁 A 節；Siliconomics SIF_UnitsPlat_tpu／_oth"),
        ("基準顆數係數 CY2028 ×0.80、CY2029 ×0.710（um28_B、um29_B；情境 B）", "Siliconomics", vtxt, "情境 B 顆數＝原始顆數 × 係數（SIF_UnitsCoef_Base）", "0.80／0.710"),
        ("Broadcom 份額、核心售價、HBM 代購比例（sh_*、core_*、f_hbm）", "Siliconomics", vtxt, "XPU 營收的份額、單價輸入（本專案為權威，規則第 13 條）", "In 頁 B 節"),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中"),
        ("AI 網通營收（net_26–net_29；Broadcom 歸屬）", "市場層 markets/network", net_v, "AI 網通營收（再以 k_n 校準到 FY2026 推估）", "In 頁 D 節"),
        ("HBM 混合單價路徑 2026–2029（hbm_<情境>_<年>）", "市場層（SK 海力士成品）", hbm_v, "代購 HBM 金額與毛利", "In 頁 C 節"),
        ("加速器模組每顆價格（pm_*、pa_*；金額份額計算用）", "市場層（NVIDIA 成品）", src_in("pm_bw").split("｜")[0], "Recon 頁 ASIC 金額份額檢查", "In 頁 I 節"),
        ("Wattonomics 錨定係數 wta_27–wta_29（舊敏感度案例；SIF_WTAnchor_Base，v0.13 起）", "Wattonomics（經 Siliconomics 轉入）", "Wattonomics 現行 v0.3.2；係數取自 Siliconomics v0.13 起快照", "Sens 頁「Wattonomics 錨定」案例（情境 B × 係數）", "舊口徑，保留供對照；對帳中兩讀數（另一側讀數）取代其角色"),
        ("對帳中年份兩讀數 C4", "Wattonomics 現行 v0.3.2 換算（經 Siliconomics v0.19 另一側讀數）", vtxt, "本模型不直接讀 Wattonomics；C4 並列兩讀數（Coord 頁 A–E 節）", "CY2027／2028／2029 另一側讀數＝現行基準 ×0.62／0.67／0.67"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。上列各輸入的原值與來源列見 In 頁；本節不改任何輸入。").font = Font(italic=True); r += 2
    # ---- C0 ----
    C.cell(r, 1, "G. C0 由下往上對帳（三層基準協調規則 v1.4.1 C0／C5）：個股的產業相關量 vs 層級對應值；差距超過 ±15% 標「對帳中」，不改上游值（主控開對帳單）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "本模型值", "層級對照值", "差距", "狀態", "單位與口徑", "對照來源（層、版本）", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    N = SIL.NAMES
    pct = '+0.0%;-0.0%'
    P_ = wb["PL"]
    rbA = lab(P_, "情境 A：原始顆數"); rcB = lab(P_, "情境 B：")
    r_ku = lab(P_, "XPU 營收（未校準）", rbA, exact=True); r_xpu = lab(P_, "XPU 營收", rbA, exact=True)

    def c0(label, mine, layer, unit, source, ev, note, kind="ratio", fmt="0.000"):
        nonlocal r
        C.cell(r, 1, label)
        C.cell(r, 2, mine).number_format = fmt; C.cell(r, 3, layer).number_format = fmt
        if kind == "ratio":
            C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
            C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"對帳中","一致")')
        else:   # 恆等式：本模型值＋層內其他已建模者 ≤ 1；差距欄＝未建模殘差
            C.cell(r, 4, f"=1-C{r}").number_format = "0.0%"
            C.cell(r, 5, f'=IF(C{r}<=1,"一致（殘差 "&TEXT(D{r},"0%")&" 明列）","對帳中")')
        C.cell(r, 6, unit); C.cell(r, 7, source); C.cell(r, 8, ev); C.cell(r, 9, note)
        r += 1
    S = lambda k: SIL.base(k)
    up = lambda k, i: N[k][i]            # SIF_UnitsPlat_*：索引 1＝2026、2＝2027、3＝2028
    cb = N["SIF_UnitsCoef_Base"]
    c0("Google TPU 市場顆數 CY2028（情境 B）", "=u_tpu_28*um28_B", up("SIF_UnitsPlat_tpu", 3) * cb[3], "百萬顆；情境 B＝原始顆數 × 基準係數", f"Siliconomics {SIL.VERSION}（SIF_UnitsPlat_tpu × SIF_UnitsCoef_Base）", "Assumed（層）／同源",
       "本模型直接讀層級值；恆等，差距應為 0")
    c0("其他 ASIC 市場顆數 CY2028（情境 B）", "=u_oth_28*um28_B", up("SIF_UnitsPlat_oth", 3) * cb[3], "百萬顆；情境 B", f"Siliconomics {SIL.VERSION}（SIF_UnitsPlat_oth × SIF_UnitsCoef_Base）", "Assumed（層）／同源", "同上")
    c0("Google TPU 市場顆數 CY2029（情境 B）", "=u_tpu_28*(1+g29_tpu)*um29_B", up("SIF_UnitsPlat_tpu", 4) * cb[4], "百萬顆；情境 B（2029＝2028×(1＋成長)×係數）", f"Siliconomics {SIL.VERSION}", "Assumed（層）／同源", "同上；CY2029 在 Siliconomics 標對帳中（連續法）")
    c0("Broadcom 占 Google TPU 顆數份額 CY2027", "=sh_tpu_27", S("sh_tpu_27"), "%；顆數份額", f"Siliconomics {SIL.VERSION}（SIF_sh_tpu_27）", "Assumed（設計夥伴實績間接證據，C0 第 5 點）", "本模型直接讀層級值；恆等", fmt="0.0%")
    c0("Broadcom 占其他 ASIC 顆數份額 CY2027", "=sh_oth_27", S("sh_oth_27"), "%；顆數份額", f"Siliconomics {SIL.VERSION}（SIF_sh_oth_27）", "Assumed", "同上", fmt="0.0%")
    c0("XPU 每顆核心售價：Google TPU（2026 基準）", "=core_tpu", S("core_tpu"), "US$K；不含 HBM", f"Siliconomics {SIL.VERSION}（SIF_core_tpu）", "Assumed", "同上", fmt="0.0")
    # 份額恆等式（個股加總＋殘差＝100%）
    sumtpu, sumoth = N["SIF_ASICShareSum_tpu"], N["SIF_ASICShareSum_oth"]
    c0("份額恆等式：Google TPU CY2027 已建模同業份額合計（Broadcom＋聯發科＝Siliconomics 份額表）", f"=sh_tpu_27+{S('mtk_tpu_27')}", sumtpu[1],
       "%；Broadcom 份額 ＋ 聯發科（SIF_mtk_tpu_27）；層欄＝Siliconomics SIF_ASICShareSum_tpu", f"Siliconomics {SIL.VERSION}（SIF_ASICShareSum_tpu）", "Assumed",
       "個股加總＋未建模殘差＝100%（C0 第 1 點）；殘差由未建模同業（Alchip／GUC 等）吸收，不得默默歸零", kind="identity", fmt="0.0%")
    c0("份額恆等式：其他 ASIC CY2027 已建模同業份額合計（Broadcom＋聯發科＝Siliconomics 份額表）", f"=sh_oth_27+{S('mtk_oth_27')}", sumoth[1],
       "%；Broadcom 份額 ＋ 聯發科（SIF_mtk_oth_27）", f"Siliconomics {SIL.VERSION}（SIF_ASICShareSum_oth）", "Assumed",
       "殘差大（Marvell、Alchip／GUC 等未建模或由其他模型負責），明列不吸收", kind="identity", fmt="0.0%")
    # 營收口徑：FY2026 XPU
    c0("FY2026 XPU 營收：公司實績推估（錨點）vs 層級由下而上（顆數 × 份額 × 核心售價＋代購 HBM，未校準）", f"=PL!D{r_xpu}", f"=PL!D{r_ku}", "US$B；FY2026＝CY2026；差距＝XPU 校準係數 − 1",
       f"Siliconomics {SIL.VERSION} 顆數、份額、核心售價（層）；錨點＝Q 頁（AI 營收 57.6 − 網通 17.6）", "Verified（AI 營收，公司）／Derived（XPU＝AI − 網通）",
       "±15% 內：層級由下而上與公司實績相容；校準係數只用於 FY2026 之後的縮放", fmt="#,##0.0")
    c0("FY2026 AI 網通營收：公司實績推估（錨點）vs 市場層歸屬值（markets/network）", "=a26_net", "=net_26", "US$B；FY2026；差距＝網通校準係數 − 1",
       f"{net_v}（net_26：Broadcom 歸屬）", "Derived（錨點：法說 XPU 占 AI 73%，二手）／Assumed（層：廠商歸屬比例）",
       "超過 +15%：市場層的 Broadcom 歸屬比例偏低（Assumed），本模型以網通校準係數校準到推估；不改上游值，列對帳單候選（由網通層處理）", fmt="#,##0.0")
    G1 = r - 1
    r += 1
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：產業部署 GW，Siliconomics 現行基準 vs 另一側讀數（Wattonomics 換算）。這是兩層之間的差距，不是 Broadcom 自己的量對層級值；Siliconomics 已標對帳中，本公司不標、不計入對帳中項目數，依 C4 並列兩讀數目標價").font = BOLD; r += 1
    for j, h in enumerate(["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    L0 = r
    for j, y in enumerate(ys):
        C.cell(r, 1, f"產業部署 GW CY{y}")
        C.cell(r, 2, N["SIF_DeployGW_Base"][j]).number_format = "#,##0.00"; C.cell(r, 3, N[f"SIF_DeployGW_Alt_{y}"]).number_format = "#,##0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
        C.cell(r, 5, "層間參考（Siliconomics 標對帳中；本公司不計）" if SIL.RECO_FLAG[y] else "層間參考（兩層一致）")
        C.cell(r, 6, "GW；SIF_DeployGW_Base vs SIF_DeployGW_Alt_<年>")
        C.cell(r, 7, f"Siliconomics {SIL.VERSION}（旗標 SIF_RecoFlag_{y}＝{int(SIL.RECO_FLAG[y])}）；Siliconomics 對帳單 r3")
        C.cell(r, 8, "Derived（基準）／Wattonomics 換算（另一側）"); C.cell(r, 9, "兩層之間的差距，不裁決；本模型 Coord 頁 A–E 節並列兩讀數目標價")
        r += 1
    L1 = r - 1
    r += 1
    C.cell(r, 1, "I. 對帳中年份與損益表年度標示（C4；FY(t)＝CY(t)，FY2030E–FY2032E 為 FY2029E 之後的延伸期，依賴 CY2029）").font = BOLD; r += 1
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
        C.cell(r, 2 + j, f'=IF({"$"+"BCDE"[cy-2026]+"$"+str(r_flag)}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（G 表；不含層級之間參考列）"); C.cell(r, 2, f'=COUNTIF(E{G0}:E{G1},"對帳中")'); r_cnt = r; r += 1
    C.cell(r, 1, "說明：標「對帳中」的 C0 項目不改上游值；個股修正方向依證據（C0 第 3 點）：AI 網通 FY2026 為個股有實績推估、層（市場層歸屬比例）為 Assumed → 下次升版由層吸收，由主控開對帳單；"
                 "產業部署 GW 為層與層之間的差距（參考列，Siliconomics 對帳單 r3 處理），本模型並列兩讀數目標價（Coord 頁 A–E 節），不裁決，也不計入本公司對帳中項目數。"); r_note = r; r += 1
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, G0=G0, G1=G1, L0=L0, L1=L1, H_PL=H_PL, cnt=r_cnt, note=r_note)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cu, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, band = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), lab(V, "價格帶（DCF")
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    rb = lab(P, "情境 B："); rcB = lab(P, "C. 報告值", rb)
    pr = lambda lbl, s=rcB: lab(P, lbl, s, exact=True)
    rev_b, eps_b, gm_b = pr("營收"), pr("EPS（non-GAAP）"), pr("毛利率")
    xpu_b = lab(P, "XPU 營收", rb, exact=True); net_b = lab(P, "AI 網通營收", rb, exact=True); nai_b = lab(P, "非 AI 半導體", rb, exact=True); sw_b = lab(P, "基礎架構軟體", rb, exact=True)
    ai_b = lab(P, "AI 半導體營收", rb, exact=True)
    ra27, ra28, rm28, reps, rgm = (lab(RC, "FY2027 AI 營收 vs 公司展望"), lab(RC, "FY2028 AI 營收 vs 公司展望"), lab(RC, "FY2028：所需 XPU 顆數倍數"),
                                    lab(RC, "FY2027 EPS vs 共識"), lab(RC, "FY2027 毛利率 vs 共識隱含"))
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, SEN['first'])}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.6 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元與同業比較法 "), f'TEXT({cadj},"0")', T(" 美元都低於現價 "), 'TEXT(px,"0")',
        T(" 美元；後者已含 Broadcom 過去十年相對同業的折價（"), f'TEXT({coef},"0.00")', T(" 倍）。不採折價時加權目標價 "), f'TEXT({noadj},"0")', T(" 美元（"),
        f'TEXT({noadj}/px-1,"+0%;-0%")', T("），評等為"), REC5(f"({noadj}/px-1)"), T("。基準情境 FY2028 營收年增只有 "), f'TEXT(PL!F{rev_b}/PL!E{rev_b}-1,"0.0%")', T("。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：同業比較法採用 Broadcom 過去十年相對同業的本益比折價；基準情境的顆數在 2028 年先減（係數 0.8）；FY2028 的 AI 營收展望（"),
        'TEXT(co_ai28,"0")', T(" 十億美元）無法由市場層顆數支持。不採折價時評等為"), REC5(f"({noadj}/px-1)"), T("（"), f'TEXT({noadj}/px-1,"+0.0%;-0.0%")',
        T("）；現金流折現權重降到三成時為"), REC5(f"((0.3*{dcf}+0.7*{cadj})/px-1)"), T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("），進加權的同業比較法 "), f'TEXT({cadj},"0.00")',
        T(" 美元＝同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 歷史折價 "), f'TEXT({coef},"0.00")', T(" 倍 × 前瞻 EPS。"))
    put("o2", T("現價隱含的條件：要讓現金流折現法支持現價，折現率須降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0")', T("，模型 "), f'TEXT({rv("rv0")},"0")', T(" 十億美元）。"))
    put("o3", T("盈餘低於共識：FY2027 EPS "), f'TEXT(Recon!C{reps},"0.00")', T(" 美元，比共識 "), f'TEXT(Recon!B{reps},"0.00")', T(" 低 "), f'TEXT(-Recon!D{reps},"0.0%")',
        T("；毛利率 "), f'TEXT(Recon!C{rgm},"0.0%")', T("，比共識隱含的約 "), f'TEXT(Recon!B{rgm},"0.0%")', T(" 低，來自代購 HBM 占比上升（代購比例與毛利率為假設）。"))
    put("o4", T("核心問題：AI 營收達不到公司展望。FY2027 "), f'TEXT(Recon!C{ra27},"0.0")', T("（展望 "), f'TEXT(Recon!B{ra27},"0")', T("，差 "), f'TEXT(Recon!D{ra27},"+0.0%;-0.0%")',
        T("）；FY2028 情境 B "), f'TEXT(Recon!E{ra28},"0.0")', T("，只有展望 "), f'TEXT(Recon!B{ra28},"0")', T(" 的 "), f'TEXT(Recon!E{ra28}/Recon!B{ra28},"0%")',
        T("。要達到展望，FY2028 Broadcom XPU 顆數須為模型的 "), f'TEXT(Recon!E{rm28},"0.00")', T(" 倍（其他條件不變的反推，不是預測）。"))
    put("o5", T("框架假設的影響大於單一營運輸入：同業比較法不加歷史折價時加權目標價 "), f'TEXT({SB("同業比較法的歷史折價係數", "E")},"0")', T(" 美元（"), f'TEXT({SB("同業比較法的歷史折價係數", "C")},"+0%;-0%")',
        T("），情境全押樂觀 "), f'TEXT({SB("情境權重", "C")},"+0%;-0%")', T("、全押保守 "), f'TEXT({SB("情境權重", "B")},"+0%;-0%")', T("。營運輸入中 XPU 核心售價 ±20% 影響最大（"),
        f'TEXT({SB("XPU 核心售價", "B")},"+0.0%;-0.0%")', T("／"), f'TEXT({SB("XPU 核心售價", "C")},"+0.0%;-0.0%")', T("）。"))
    put("drv", T("分部營收推導（情境 B）：AI 半導體＝XPU 營收＋AI 網通；XPU 營收＝（Google TPU 與其他 ASIC 的市場顆數 × Broadcom 份額）×（每顆核心售價＋代購 HBM 比例 × 每顆 HBM 容量 × HBM 單價）× FY2026 校準係數；"
                 "AI 網通＝市場層歸屬金額 × 情境顆數係數 × 校準係數；非 AI 半導體與基礎架構軟體各以固定年增率推算。四者相加為合併營收。FY2028：XPU "), f'TEXT(PL!F{xpu_b},"0")', T("、AI 網通 "), f'TEXT(PL!F{net_b},"0")',
        T("、非 AI 半導體 "), f'TEXT(PL!F{nai_b},"0")', T("、軟體 "), f'TEXT(PL!F{sw_b},"0")', T("，合計 "), f'TEXT(PL!F{rev_b},"0")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型的產業顆數、份額與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1']}", T("（見 Coord 頁 F 節）。C0 由下往上對帳共 "), f'TEXT(COUNTA(Coord!A{CO["G0"]}:A{CO["G1"]}),"0")', T(" 項，其中「對帳中」"),
        f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項（只算 Broadcom 自己的量對它讀的層級值）；產業部署 GW 兩層差距另列參考、不計入；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調新機制（共用工作單 D 節）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史）、PL 頁末標準損益表"
                  "（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，Non-GAAP）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準（整本重算快照 avgo/sens.json）、"
                  "Coord 頁 F–I 節（讀哪一層哪一版、C0 由下往上對帳、損益表年度標示）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v0.5.5" in head, head[:120]
    RD["A1"].value = head.replace("v0.5.5", f"v{VER}", 1) + a1[len(head):] + (
        f"；v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調新機制（Coord 頁 F–I 節 C0 對帳；Hist、RevDCF 頁；PL 標準損益表；Val 0.7；Sens 加權目標價相對基準），既有數值不變")


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("AVGO_SENS") == "1":
        return "v056_post：略過（alt_run／sens_avgo 建置）"
    HR = hist_sheet(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb, here)
    CO = coord_c0(wb, SIL, None)
    assert CO, "Coord 頁不存在：先跑 alt_run.py"
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    version(wb)
    wb._v056 = dict(PL=PLA, RV=RVA, RM=RMA, SEN=SEN, CO=CO, VG=VG)
    return f"v056_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；Coord F–I 節；README v1.3 白話列 {len(RMA)} 列"
