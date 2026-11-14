# Amkor v0.4.4（2026-11-14）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調（規則 v1.4.2：C0 讀取層級與公司實績對帳、C4 兩讀數）。
# 樣板：日月光 v0.2.6／v0.2.7（ase/v026_post.py、v027_post.py）與 Broadcom（avgo/v056_post.py：整本重算敏感度快照）。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史 FY2022–FY2024（公司 10-K；FY2025 取 Q 頁）
#   PL 頁末：標準損益表（FY2023–FY2025 實績＋FY2026E–FY2031E 情境 B；GAAP）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列（四個目標價）
#   Sens 頁末：加權目標價相對基準（營運輸入案例＝整本重算快照 amkr/sens.json；折現參數案例＝引用上方表；估值框架案例＝封閉式）
#   Coord 頁末：三層協調（A 讀哪一層哪一版、B 讀取層級（同源，不判定）、C 口徑註記、D 摘要、E C0（實績）：公司實績 vs 層級隱含值）
#   README 頁末：v1.3 展示層白話列與版本說明；標題段版號更正
# 建檔程式在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。另一側讀數建置（COORD_RUN＝1）與 sens_amkr.py 建置（AMKR_SENS＝1）時不執行。
import json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.4", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
K25 = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000014/amkr-20251231.htm"
K24 = "https://www.sec.gov/Archives/edgar/data/1047127/000104712725000030/amkr-20241231.htm"
Q1 = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000020/amkr-20260331.htm"
Q2 = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000046/amkr-20260630.htm"
Q3 = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000052/amkr-20260930.htm"
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


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
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


# ---------- Hist：損益表歷史（US$M；公司 10-K 合併損益表，美國 GAAP） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025 公式)；None＝該年本模型不需要（不填）
HIS = [
    ("rev", "營收", 7091.585, 6503.065, 6317.692, "=Q!B4"),
    ("cogs", "營業成本", None, 5559.912, 5384.480, "=Q!B4-Q!B9"),
    ("gp", "毛利", 1329.987, 943.153, 933.212, "=Q!B9"),
    ("opx", "營業費用（銷管＋研發）", None, 472.866, 494.757, "=Q!B11"),
    ("oi", "營業利益", 897.186, 470.287, 438.455, "=Q!B12"),
    ("nonop", "營業外淨額（其他收益 − 利息費用）", None, -26.446, -7.439, "=Q!B13"),
    ("pbt", "稅前利益", 856.932, 443.841, 431.016, "=Q!B14"),
    ("tax", "所得稅", None, 81.710, 75.481, "=Q!B15"),
    ("nic", "稅後淨利（合併，含非控制權益）", None, 362.131, 355.535, "=Q!B14-Q!B15"),
    ("nci", "非控制權益損益", None, 2.318, 1.523, "=Q!B16"),
    ("ni", "歸屬母公司淨利", 765.823, 359.813, 354.012, "=Q!B17"),
    ("shr", "稀釋加權平均股數（百萬股）", None, 247.176, 247.818, "=Q!B18"),
    ("eps", "EPS（稀釋；FY2023–FY2024＝公司公告，FY2025＝Q 頁淨利 ÷ 股數；US$）", 3.11, 1.46, 1.43, "=Q!B19"),
    ("epsr", "EPS（稀釋，公司公告；US$）", 3.11, 1.46, 1.43, "=Q!B20"),
]


def hist_is(wb):
    Hs = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    Hs["A1"] = ("Hist — 損益表歷史（v0.4.4 起；百萬美元；EPS 美元；美國 GAAP 合併損益表、歸屬母公司）。PL 頁末「標準損益表」的歷史三年（FY2023–FY2025）與年增率基期用。"
                "FY2022 只取營收、毛利、營業利益、稅前利益、歸屬母公司淨利與 EPS（年增率基期）；FY2025 欄＝Q 頁（4Q25 新聞稿與 10-K）。只被 PL 頁標準損益表引用，不回饋任何計算")
    Hs["A1"].font = BOLD
    Hs["A2"] = (f"來源（公司原文＝SEC 10-K 合併損益表，Verified）：FY2023、FY2024（含 FY2022 比較欄）＝FY2024 年報（2025-02，accession 0001047127-25-000030；{K24}）；"
                f"FY2025、FY2024、FY2023＝FY2025 年報（2026-02，accession 0001047127-26-000014；{K25}）。FY2025 與 FY2024、FY2023 的重疊數字兩份年報相同")
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025（Q 頁）", "標記", "來源"], 1):
        c = Hs.cell(4, j, h); c.font = BOLD; c.fill = HDR
    rows = {}
    r = 5
    for k, labl, v22, v23, v24, v25 in HIS:
        rows[k] = r
        Hs.cell(r, 1, f"h_{k}"); Hs.cell(r, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if v is None: continue
            c = Hs[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k in ("eps", "epsr") else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"h_{k}_{yy}", f"Hist!${col}${r}")
        Hs.cell(r, 7, "Verified")
        Hs.cell(r, 8, "FY2024、FY2025 年報合併損益表（FY2023 各項取 FY2024 年報比較欄；FY2022 取 FY2024 年報比較欄）")
        r += 1
    r += 1
    for labl, f in (("核對：營收 − 營業成本 − 毛利（FY2023–FY2025，應為 0）", lambda c: f"={c}{rows['rev']}-{c}{rows['cogs']}-{c}{rows['gp']}"),
                    ("核對：毛利 − 營業費用 − 營業利益（應為 0）", lambda c: f"={c}{rows['gp']}-{c}{rows['opx']}-{c}{rows['oi']}"),
                    ("核對：營業利益 ＋ 營業外淨額 − 稅前利益（應為 0）", lambda c: f"={c}{rows['oi']}+{c}{rows['nonop']}-{c}{rows['pbt']}"),
                    ("核對：稅前利益 − 所得稅 − 稅後淨利（應為 0）", lambda c: f"={c}{rows['pbt']}-{c}{rows['tax']}-{c}{rows['nic']}"),
                    ("核對：稅後淨利 − 非控制權益 − 歸屬母公司淨利（應為 0）", lambda c: f"={c}{rows['nic']}-{c}{rows['nci']}-{c}{rows['ni']}")):
        Hs.cell(r, 2, labl)
        for col in "DEF":
            Hs[f"{col}{r}"] = f(col); Hs[f"{col}{r}"].number_format = "0.000"
        r += 1
    Hs.column_dimensions["A"].width = 12; Hs.column_dimensions["B"].width = 48
    for col in "CDEF": Hs.column_dimensions[col].width = 14
    S = wb["Sources"]; n = S.max_row + 1
    for i, (a, b) in enumerate((("公司：FY2025 年報 10-K（SEC，2026-02；FY2025、FY2024、FY2023 合併損益表，Hist 頁）", K25),
                                ("公司：FY2024 年報 10-K（SEC，2025-02；FY2024、FY2023、FY2022 合併損益表，Hist 頁）", K24),
                                ("公司：1Q26 10-Q（SEC；附註 2：產品組別營收、終端市場比重，Coord 頁 E 節）", Q1),
                                ("公司：2Q26 10-Q（SEC；附註 2，Coord 頁 E 節）", Q2),
                                ("公司：3Q26 10-Q（SEC；附註 2，Coord 頁 E 節）", Q3))):
        S.cell(n + i, 1, a); S.cell(n + i, 2, b)
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"       # 標準表欄；E 起（FY2025A…FY2031E）對應 PL 情境 B 區塊 C…I 欄
SRC = "CDEFGHI"


def pl_standard(wb):
    P = wb["PL"]; D = wb["DCF"]
    rbd = lab(P, "情境 B：")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    R = dict(rev=ex("營收"), gp=ex("毛利"), opx=ex("營業費用（銷管＋研發）"), oi=ex("營業利益"), oth=ex("營業外淨額"), pbt=ex("稅前利益"), tax=ex("所得稅"), nci=ex("非控制權益"),
             ni=ex("歸屬母公司淨利"), sh=ex("稀釋股數"), eps=ex("EPS（稀釋）"))
    dbr = lab(D, "情境 B"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.4.4 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026E–FY2031E，情境 B 基準；FY2026E 為錨點）。百萬美元；EPS 美元。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（美國一般公認會計原則合併財報、歸屬母公司）。艾克爾的新聞稿與 10-K 損益表即 GAAP，模型 EPS 與目標價皆用此口徑，沒有另設 Non-GAAP 口徑（Recon 頁的賣方共識 EPS 為 non-GAAP，只在該頁對照）。"
                 "EPS 用稀釋 EPS（完全稀釋，本模型口徑）。預測期的「其他收入（淨額）」＝營業外淨額（其他收益 − 利息費用）")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（1Q–3Q26 實績＋4Q26 依公司指引）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E、FY2031E＝延伸期（收斂法只推營收與歸屬母公司淨利）：毛利率與營業費用率沿用 FY2029E 推算營業成本、毛利與營業費用；"
                 "營業利益取 DCF 頁情境 B（由淨利反推）；其他收入＝FY2029 營業外淨額假設；所得稅＝稅前利益 −（歸屬母公司淨利＋非控制權益）。兩種推法的營業利益差額列在表末核對列，不改既有淨利與目標價")
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
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1

    M = "#,##0.0"; PCT = "0.0%"
    hh = lambda k: [f"=h_{k}_23", f"=h_{k}_24"]
    row("rev", "營收", "US$M", hh("rev") + [f"=PL!{c}{R['rev']}" for c in SRC], M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L), "=D{rev}/C{rev}-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS[1:], COLS[2:])], PCT)
    r_c = r   # 下一列＝cogs；J、K（FY2030E、FY2031E）用 FY2029E（I 欄）比率
    row("cogs", "營業成本（COGS）", "US$M", hh("cogs") + [f"=PL!{c}{R['rev']}-PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-I{r_c + 2})"], M, "FY2030E、FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$M", hh("gp") + [f"=PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（銷管＋研發）", "US$M", hh("opx") + [f"=PL!{c}{R['opx']}" for c in SRC[:5]] + [f"=J{L['rev']}*I{r_c + 3}/I{L['rev']}", f"=K{L['rev']}*I{r_c + 3}/I{L['rev']}"], M, "FY2030E、FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "US$M", hh("oi") + [f"=PL!{c}{R['oi']}" for c in SRC[:5]] + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外淨額）", "US$M", hh("nonop") + [f"=PL!{c}{R['oth']}" for c in SRC[:5]] + ["=nonop29", "=nonop29"], M, "利息、其他收益；FY2030E 起＝FY2029 營業外淨額假設")
    row("pbt", "稅前利益", "US$M", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$M", hh("tax") + [f"=PL!{c}{R['tax']}" for c in SRC[:5]] + [f"=J{L['pbt']}-(PL!H{R['ni']}+nci)", f"=K{L['pbt']}-(PL!I{R['ni']}+nci)"], M, "FY2030E、FY2031E＝稅前利益 −（歸屬母公司淨利＋非控制權益），等於稅率 × 稅前利益")
    row("nic", "稅後淨利（合併，含非控制權益）", "US$M", [f"={c}{L['pbt']}-{c}{L['tax']}" for c in COLS], M)
    row("nci", "非控制權益損益", "US$M", hh("nci") + [f"=PL!{c}{R['nci']}" for c in SRC[:5]] + ["=nci", "=nci"], M)
    row("ni", "歸屬母公司淨利（EPS 與目標價口徑）", "US$M", [f"={c}{L['nic']}-{c}{L['nci']}" for c in COLS], M)
    row("nm", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "百萬股", hh("shr") + [f"=PL!{c}{R['sh']}" for c in SRC], "#,##0.0", "FY2023–FY2024＝公司公告稀釋加權平均股數")
    row("eps", "EPS（稀釋；US$）", "US$", ["=h_eps_23", "=h_eps_24"] + [f"=PL!{c}{R['eps']}" for c in SRC], "0.00", "FY2023–FY2024＝公司公告稀釋 EPS；FY2025 起＝模型（歸屬母公司淨利 ÷ 稀釋股數）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY；歸屬母公司淨利）", "%", ["=C{ni}/h_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 歸屬母公司淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；公司公告稀釋 EPS，與上方 EPS 的差只在 FY2025 取模型淨利 ÷ 股數的捨入）").font = BOLD; r_gaap = r; r += 1
    row("g_epsd", "EPS（稀釋，公司公告；US$）", "US$", ["=h_epsr_23", "=h_epsr_24", "=h_epsr_25"], "0.00", "FY2025 公司公告 1.50（Q 頁第 20 列）")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$M", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E、FY2031E 見下兩列）", "US$M", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "US$M",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：歸屬母公司淨利 − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "US$M", [None, None] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    row("c4", "核對：EPS − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "US$", [None, None] + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    hr_ = lab(wb["Hist"], "h_nic", col=1)
    row("c5", "核對：稅後淨利（合併，Hist）− 本表（FY2023–FY2025）", "US$M", [f"=Hist!D{hr_}-C{L['nic']}", f"=Hist!E{hr_}-D{L['nic']}", f"=Hist!F{hr_}-E{L['nic']}"], "0.000")
    for col in "CDEFGHIJK": P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]; P = wb["PL"]; V = wb["Val"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    rbd = lab(P, "情境 B："); rev26 = lab(P, "營收", rbd, exact=True)
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.4 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In、PL 頁，無任何頁引用本頁，不影響目標價。金額百萬美元；價格美元")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。FY2027–FY2032 全部為預測年（艾克爾 FY2026 為錨點，不動）")
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
    kv("nc", "淨現金滾到評價日（百萬美元）", "=dcf_nc+stub_fcff4q-stub_div", "#,##0.0", "同 DCF 頁每股價值公式（淨現金與投資＋stub 期 FCFF − stub 期股利）")
    kv("sh", "評價用稀釋股數（百萬股）", "=sh_v", "0.0")
    kv("evr", "支持現價所需的企業價值 EV（百萬美元）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；百萬美元）", f"=DCF!$C${ev}", "#,##0.0")
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
    yr("rev", "營收（百萬美元）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（百萬美元）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027–FY2032 皆為預測年）", lambda dc, c: f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（百萬美元）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（百萬美元）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（百萬美元，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2027–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（百萬美元）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（百萬美元）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2027–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(B{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(B{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2027–FY2032 營收等比例放大；FY2026 錨點不動；營業利益率、稅率、折舊、資本支出占營收比率不變；亞利桑那第一期 FY2027 支出隨營收同比例放大，屬近似）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（百萬美元）", f"=fc_wc*PL!$D${rev26}*B{Y['fr']}*B{Y['df']}", "#,##0.0", "FY2027 營運資金增加中屬 FY2026 營收的部分（FY2026 為錨點，不放大）")
    kv("b", "隨營收倍數等比例變動的 EV（百萬美元）",
       f"=SUMPRODUCT(B{Y['cf']}:G{Y['cf']},B{Y['fr']}:G{Y['fr']},B{Y['df']}:G{Y['df']})-fc_wc*B{Y['rev']}*B{Y['fr']}*B{Y['df']}"
       f"-SUMPRODUCT(C{Y['wc']}:G{Y['wc']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2027 營收（百萬美元）", f"=B{A['s1']}*B{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2027 營收（情境 B；百萬美元）", f"=B{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2027 營收年增率（相對 FY2026 錨點營收）", f"=B{A['rv1']}/PL!$D${rev26}-1", "0.0%", "FY2026 為錨點（1Q–3Q26 實績＋4Q26 指引），不放大")
    kv("g0", "模型：FY2027 營收年增率（情境 B）", f"=B{A['rv0']}/PL!$D${rev26}-1", "0.0%")
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
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%；落在下界表示即使折現率壓到下界也不足以支持現價")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
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
                     T(" 美元；要支持現價 "), 'TEXT(px,"0.00")', T(" 美元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大為","縮小為")', T(" "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<=wacc,"降到","升到")', T(" "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2027–FY2032 營收全面"), f'IF(B{A["s1"]}>=1,"放大","縮小")',
                     T(" "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2027 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 百萬美元，相對 FY2026 錨點年增 "), f'TEXT(B{A["g1"]},"0%")',
                     T("）；或營業利益率"), f'IF(B{A["m1"]}>=B{A["m0"]},"提高到","降到")', T(" "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 艾克爾過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給艾克爾的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，比模型採用的相對倍數"), f'IF(D{A["rc1"]}<0,"低 ","高 ")', f'TEXT(ABS(D{A["rc1"]}),"0%")', T("。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
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
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.4.4 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "美元", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 美元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（賣方平均）", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列） ----------
TOR_PARAM = (("WACC", "WACC −1／＋1 個百分點（永續期新投資報酬率隨 WACC）"), ("β", "股權 β 低／高（產業無槓桿 β ∓0.3 後重新加槓桿）"), ("ERP", "股權風險溢酬 4%／6%（基準 5%）"),
             ("永續成長率", "永續成長率 2%／4%（基準 3%）"), ("RONIC", "永續期新投資報酬率高於 WACC 5 個百分點（只有高案）"), ("資本支出 ÷ 營收", "每年資本支出占營收 −1／＋1 個百分點"),
             ("營運資金", "營運資金占營收 5%／20%（基準 11.9%）"), ("淨現金不含", "淨現金不含權益投資（只有低案）"), ("終值改用", "DCF 終值改用 FCFF ×（1＋g）法（只有高案）"))


def sens_tor(wb, here):
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")                  # B、C、D＝情境 A、B、C
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    bl = {s: lab(P, f"情境 {s}：") for s in "ABC"}
    epr = {s: lab(P, "EPS（稀釋）", bl[s], exact=True) for s in "ABC"}
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.4.4 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（amkr/sens_amkr.py：三情境同時套用案例覆寫，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    # 1) 營運輸入案例（整本重算快照）
    for k, d in J["biz"].items():
        SE.cell(r, 1, d["label"])
        for col, key in ((4, "低"), (5, "高")):
            if d[key] is None: continue
            c = SE.cell(r, col, d[key]["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
            SE.cell(r, col - 2, f"={'D' if col == 4 else 'E'}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（amkr/sens.json）；覆寫同本頁上方「一次動一個輸入」案例，三情境同時套用"); r += 1
    # 2) 估值框架案例（封閉式；加權＝DCF 權重 × DCF ＋（1 − DCF 權重）× Comparable）
    frame = [
        ("同業比較法不做相對倍數調整（改用市場主流同業法；只有低案）", f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}", None, "Val 頁：Comparable（未調整）取代調整後值"),
        ("DCF 在加權中的權重 0%（只看同業比較法）／100%（只看現金流折現）（基準＝50%）", f"={cadj}", f"={dcf}", "Val 頁：Comparable（調整後）與 DCF 目標價"),
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*(2/12*PL!E{epr['C']}+10/12*PL!F{epr['C']})",
         f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*(2/12*PL!E{epr['A']}+10/12*PL!F{epr['A']})", "Val 頁：單一情境的 DCF 價值與 PL 頁該情境前瞻 EPS"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl)
        for col, f in ((4, lo), (5, hi)):
            if f is None: continue
            SE.cell(r, col, f).number_format = "#,##0.00"
            SE.cell(r, col - 2, f"={'D' if col == 4 else 'E'}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, note); r += 1
    # 3) 折現參數案例（引用既有表；DCF 權重案例已由 0%／100% 涵蓋）
    for sr in src:
        t = str(SE.cell(sr, 1).value)
        if t.startswith("DCF 權重"): continue
        labl = next((v for k, v in TOR_PARAM if t.startswith(k)), None)
        assert labl, t
        SE.cell(r, 1, labl)
        for col, scol, tcol in ((2, "D", 4), (3, "F", 5)):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
            SE.cell(r, tcol, f"=Sens!{chr(ord(scol) - 1)}{sr}").number_format = "#,##0.00"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「目標價參數敏感度」表）"); r += 1
    last = r - 1
    r += 1
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f'=IF(ABS({tp}-{J["_base"]["tp"]!r})<=1E-6*ABS({tp}),"CURRENT","STALE")')
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 amkr/sens_amkr.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調（A 讀哪一層、B 讀取層級、C 口徑、D 摘要、E C0 實績） ----------
def coord_c0(wb, SIL):
    C = wb["Coord"]; I = wb["In"]; P = wb["PL"]; V = wb["Val"]
    rbd = lab(P, "情境 B：")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    r_u = ex("AI 加速器顆數（含情境係數）")
    r_ai = lab(P, "AI 先進封裝營收（", rbd)
    ikey = {str(I.cell(i, 1).value): i for i in range(4, I.max_row + 1) if I.cell(i, 1).value}
    C.column_dimensions["F"].width = 60
    for col in "BCDE": C.column_dimensions[col].width = 24
    r_ratio = lab(C, "另一側讀數 ÷ 現行基準")
    r = C.max_row + 3
    C.cell(r, 1, "三層協調（v0.4.4 新增；三層基準協調規則 v1.4.2 C0 個股與層的一致性、C4 個股不裁決、C5 往上回饋只有證據）").font = BOLD; r += 1
    R = {}
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    # ---- A. 讀哪一層、哪一版 ----
    C.cell(r, 1, "A. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; R["A"] = r; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "讀自（層與版本）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["A_hdr"] = r; r += 1

    def arow(label, cells, src, fmt="0.000"):
        nonlocal r
        C.cell(r, 1, label)
        for j, v in enumerate(cells):
            if v is None: continue
            c = C.cell(r, 2 + j, v); c.number_format = fmt
        C.cell(r, 6, src); R[label] = r; r += 1
    tsm = "經台積電個股模型 v0.5.4 成品（tsmc/dist；該模型讀 Siliconomics v0.19）"
    arow("AI 加速器顆數（情境 A 原始顆數；百萬顆）", [f"=In!E{ikey[f'units_{y}']}" for y in (26, 27, 28, 29)], f"{ver}：SIF_UnitsPlat_Total，{tsm}取得（In 頁 units_26–units_29，唯讀）", "0.000")
    arow("情境 B 顆數係數（÷ 原始顆數）", [None, 1, f"=In!E{ikey['um28_B']}", f"=In!E{ikey['um29_B']}"], f"{ver}：SIF_UnitsCoef_Base，{tsm}取得（In 頁 um28_B、um29_B；CY2026–27 為 1）")
    arow("封裝面積指數（2026＝1）", [f"=In!E{ikey[f'ar_{y}']}" for y in (26, 27, 28, 29)], f"{ver}：CoWoS 每片封裝數 cw_pp_*，{tsm}取得（In 頁 ar_26–ar_29；＝加速器用 CoWoS 片數 ÷ 顆數）")
    arow("2026 年部署中 Rubin 世代占比（Recon 用）", [f"=In!E{ikey['vr26']}", None, None, None], f"{ver}：vr26（In 頁，唯讀）", "0%")
    arow("Wattonomics 錨定係數（只進 Sens 案例，不改基準）", [None] + [f"=In!E{ikey[f'wta_{y}']}" for y in (27, 28, 29)], f"{ver}：SIF_WTAnchor_Base（R4 比值＝Wattonomics ÷ 本庫推得上線 GW）；Wattonomics 現行版本 v0.3.2")
    arow("另一側讀數 ÷ 現行基準（對帳中年份；兩讀數目標價用）", [f"={'BCDE'[i]}{r_ratio}" for i in range(4)], f"{ver}：SIF_DeployGW_AltRatio_<年>（Wattonomics 換算的產業 GW ÷ Siliconomics 基準）；上方 Coord 第 6 列")
    C.cell(r, 1, "Wattonomics：本模型不直接讀 Wattonomics 數字；現行版本 v0.3.2，僅經 Siliconomics 的 R4 比值（Sens 錨定案例）與另一側讀數比值（對帳中年份兩讀數）間接使用。其餘產業量皆讀 Siliconomics 基準（經台積電成品）。"); R["A_note"] = r; r += 2
    # ---- B. 讀取層級（同源，不判定） ----
    C.cell(r, 1, "B. 讀取層級（同源，不判定）：艾克爾讀入的 Siliconomics 值 vs 層級值（個股自層級讀入，差距恆為 0，無檢驗力；不計入對帳中項目數；規則 v1.4.2 C0 第 2 點）。"
                 "公司實績對層級隱含值的對帳見本頁 E 節「C0（實績）」").font = BOLD; R["B"] = r; r += 1
    for j, h in enumerate(["項目（年）", "本模型原值", "層級對應值", "差距（本模型 ÷ 層級 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["B_hdr"] = r; r += 1
    units_L = SIL.NAMES["SIF_UnitsPlat_Total_Base"]; gw_b = SIL.NAMES["SIF_DeployGW_Base"]
    R["B_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        col = "DEFG"[i]
        C.cell(r, 1, f"AI 加速器顆數（情境 B）CY{y}（百萬顆）")
        C.cell(r, 2, f"=PL!{col}{r_u}").number_format = "0.000"
        c = C.cell(r, 3, f"=In!E{ikey[f'units_{y % 100}']}" + (f"*In!E{ikey[f'um{y % 100}_B']}" if y >= 2028 else "")); c.number_format = "0.000"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, f"本模型：PL 頁情境 B 顆數（原始顆數 × 基準係數；來自台積電個股成品）；層：{ver} SIF_UnitsPlat_Total_Base（供給側基準顆數）。同源讀取故差距 0；艾克爾不公布加速器顆數，兩邊都沒有公司實績（等級：Assumed／Interested-party）"
                     if i == 0 else "同上")
        r += 1
    R["B_last"] = r - 1
    C.cell(r, 1, "艾克爾 AI 先進封裝營收（FY2026–FY2029）vs 層級量"); C.cell(r, 2, "—"); C.cell(r, 3, "無可比的層級量"); C.cell(r, 4, "—"); C.cell(r, 5, "不判定（無可比實績，見 E 節）")
    C.cell(r, 6, "Siliconomics 沒有艾克爾單一公司的營收、份額或出貨量：只有「委外封測（日月光、艾克爾）補充 CoWoS 產能 ÷ 台積電 ＝ 10%」（SIF cw_osat，Assumed），以 CoWoS 月產能（千片）表達，"
                 "缺每片價格與艾克爾的分配，無法換算成艾克爾 AI 先進封裝美元營收；艾克爾不在 Siliconomics 份額表內，也沒有「個股加總＋殘差＝市場值」可對（C0 第 1 點）"); R["B_nocmp"] = r; r += 1
    R["B_end"] = r - 1
    C.cell(r, 1, "判定摘要"); R["B_sum"] = r; r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; R["B2"] = r; r += 1
    for j, h in enumerate(["項目（年）", "Siliconomics 基準", "外部估計（Wattonomics 換算）", "差距（基準 ÷ 外部 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    R["B2_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        C.cell(r, 1, f"產業部署 GW CY{y}（Siliconomics 基準 ÷ Wattonomics 換算；GW）")
        c = C.cell(r, 2, gw_b[i]); c.number_format = "0.00"; c.font = BOLD if False else BLUE
        c = C.cell(r, 3, f"=B{r}*{'BCDE'[i]}{r_ratio}"); c.number_format = "0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "參考")
        C.cell(r, 6, (f"基準＝{ver} SIF_DeployGW_Base（供給側部署 GW）；外部估計（參考）＝Wattonomics 換算 SIF_DeployGW_Alt_<年>（需求側）。"
                      "這是層與層之間的差距，不是艾克爾自己的量對層級的比較；由 Siliconomics 對帳單 r3 處理，不計入本公司 C0 項目數；本模型依 C4 並列兩讀數目標價，不裁決") if i == 0 else "同上")
        r += 1
    R["B2_last"] = r - 1
    r += 1
    # ---- C. 口徑註記 ----
    C.cell(r, 1, "C. 對帳中年份處理與口徑註記（C4）").font = BOLD; R["C"] = r; r += 1
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    C.cell(r, 1, F(T("對帳中年份：CY2027、CY2028、CY2029（Siliconomics 標示；本公司 C0 無可判定項目）。兩讀數目標價並列於上方（現行基準讀數 "), f'TEXT(B{rtp},"0.00")', T(" 美元、另一側讀數 "), f'TEXT(C{rtp},"0.00")',
                    T(" 美元），評等分別為"), f'B{rrec}', T("、"), f'C{rrec}', T("。個股不裁決（C4）；Siliconomics 與 Wattonomics 的差距由對帳單處理（期限 2027-02-28）；結案前兩個目標價並列。")))
    R["C_1"] = r; r += 1
    C.cell(r, 1, F(T("口徑註記（不改模型）：另一側讀數只改情境 B、C 的 AI 加速器顆數（CY2027 起乘比值 "), f'TEXT(C{r_ratio},"0.00")', T("／"), f'TEXT(D{r_ratio},"0.00")', T("／"), f'TEXT(E{r_ratio},"0.00")',
                    T("）；情境 A（樂觀，權重 "), 'TEXT(tp_wA,"0%")', T("）不縮放，AI 先進封裝以外的終端市場、亞利桑那營收與各項成本假設不變。AI 先進封裝每顆金額由 FY2026 校準、委外比例與份額為假設，故 AI 營收約隨顆數等比例下降。"
                    "這個口徑已含在另一側目標價內；對帳單處理時一併說明。")))
    R["C_2"] = r; r += 1
    C.cell(r, 1, "C0 結論：本模型的產業量（顆數、封裝面積）全部讀 Siliconomics 基準（經台積電成品），與層同源，只列「讀取層級」、不判定；公司實績（10-Q 的產品組別營收、終端市場比重）已列入 E 節，"
                 "但 Siliconomics 沒有封測服務或先進封裝的美元市場與艾克爾份額，AI 先進封裝金額公司也未揭露，故 C0（實績）全為無可比實績，不產生對帳中項目。往上回饋只有證據（C5）：待 4Q26 財報與 2027 法說的 AI 先進封裝揭露，"
                 "以及層級補上 OSAT 市場與份額後再評估是否送對帳單。"); R["C_3"] = r; r += 2
    # ---- D. 摘要 ----
    C.cell(r, 1, "D. 摘要（HTML 與版面檢查用；兩讀數與損益表年度標示）").font = BOLD; R["D"] = r; r += 1
    for j, h in enumerate(["項目", "現行基準讀數", "另一側讀數", "另一側相對現行", ""], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    r_tp0 = lab(C, "加權目標價（DCF 與 Comparable 加權）"); r_up0 = lab(C, "上檔空間（加權目標價"); r_rec0 = lab(C, "投資建議（五級")
    R["tp"] = r; C.cell(r, 1, "加權目標價")
    C.cell(r, 2, f"=B{r_tp0}").number_format = "0.00"; C.cell(r, 3, f"=C{r_tp0}").number_format = "0.00"; C.cell(r, 4, f"=D{r_tp0}").number_format = "+0.0%;-0.0%"; r += 1
    R["up"] = r; C.cell(r, 1, "上檔空間（加權目標價 ÷ 現價 − 1）")
    C.cell(r, 2, f"=B{r_up0}").number_format = "+0.0%;-0.0%"; C.cell(r, 3, f"=C{r_up0}").number_format = "+0.0%;-0.0%"; r += 1
    R["rec"] = r; C.cell(r, 1, "投資建議（五級；兩讀數）")
    C.cell(r, 2, f"=B{r_rec0}"); C.cell(r, 3, f"=C{r_rec0}"); r += 1
    R["pl"] = r; C.cell(r, 1, "損益表年度標示（對帳中年份欄標題；FY＝CY，FY2030E、FY2031E 為延伸期，承接 CY2029 水準）")
    for j, t in enumerate(["FY2027E（對帳中）", "FY2028E（對帳中）", "FY2029E（對帳中）", "FY2030E（對帳中；延伸期）", "FY2031E（對帳中；延伸期）"]):
        C.cell(r, 2 + j, t)
    r += 2
    # ---- E. C0（實績）：公司實績 vs 層級隱含值 ----
    C.cell(r, 1, "E. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格；差距超過 ±15% 且兩讀數同號標「對帳中」）").font = BOLD
    R["E"] = r; r += 1
    C.cell(r, 1, "比較方法：個股側只用公司已公布實績（SEC 10-Q，Verified：合併營收、產品組別營收為美元金額；終端市場只揭露整數百分比，美元值＝揭露比重 × 合併營收，Derived，捨入誤差約 ±0.5 個百分點，運算約 ±2%）；"
                 "層級側用 Siliconomics v0.19 同口徑隱含值。時差：層級值為日曆年 CY2026，艾克爾會計年度＝日曆年；主讀數＝2026 年 1–3Q 累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；"
                 "第二讀數＝3Q26 × 4 vs 層級 CY2026（約等於 3Q 中點的年化速率；3Q 落在年度後半，成長年中偏高、3Q 營收較 2Q 變動不大時兩讀數接近）。"
                 "判定：兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致；層級側沒有同口徑隱含值＝無可比實績（不以模型預測或同源值補位）。"
                 "公司指引、賣方估計、本模型自己的推算只列參考，不判定。本表「無可比實績」是指層級側沒有同口徑值（合併營收、產品組別、運算終端的公司實績都存在）；AI 先進封裝（2.5D、HDFO）則是公司側也沒有金額實績。")
    R["E_note"] = r; r += 1
    for j, h in enumerate(["項目", "公司實績：2026 年 1–3Q 累計（主讀數）", "公司實績：3Q26 × 4（第二讀數）", "層級隱含：CY2026 全年", "判定（±15%）", "口徑、出處與證據等級",
                           "主讀數差距（實績 ÷（層級 × 9÷12）− 1）", "第二讀數差距（實績 ÷ 層級 − 1）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["E_hdr"] = r; r += 1
    R["E_first"] = r
    r_in = r + 7                                   # 輸入區標題列（4 列表＋參考列＋判定摘要＋空列之後）
    r_rev, r_adv, r_mai, r_pct, r_cmp = r_in + 2, r_in + 3, r_in + 4, r_in + 5, r_in + 6

    def row(label, b, c, d, src, nb="#,##0.0"):
        nonlocal r
        C.cell(r, 1, label)
        for col, v in ((2, b), (3, c)):
            x = C.cell(r, col, v)
            if not isinstance(v, str) or v.startswith("="): x.number_format = nb
        C.cell(r, 4, d)
        C.cell(r, 7, f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r})),B{r}/(D{r}*9/12)-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 8, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}/D{r}-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 5, f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),IF(AND(ABS(G{r})>0.15,ABS(H{r})>0.15,SIGN(G{r})=SIGN(H{r})),"對帳中",IF(OR(ABS(G{r})>0.15,ABS(H{r})>0.15),"時差邊界（參考）","一致")),"無可比實績")')
        C.cell(r, 6, src)
        r += 1

    row("合併營收（US$M）", f"=E{r_rev}", f"=D{r_rev}*4", "無同口徑隱含值",
        "公司實績：合併營收 1Q26 1,684.701、2Q26 1,897.965、3Q26 1,990.557（SEC 10-Q 合併損益表；輸入區；Verified；與 Q 頁 C4:E4 相差 < 0.1，見檢查列）。"
        "層級側：Siliconomics v0.19 沒有艾克爾或封測服務的營收（層級內與封測相關者只有委外補充 CoWoS 月產能，單位與口徑皆不同）。無可比實績：缺層級的封測營收市場與艾克爾份額")
    row("先進產品組（Advanced products：覆晶、記憶體、晶圓級與相關測試；US$M）", f"=E{r_adv}", f"=D{r_adv}*4", "無同口徑隱含值",
        "公司實績：先進產品組營收 1Q26 1,372.001、2Q26 1,556.676、3Q26 1,631.580（SEC 10-Q 附註 2「net sales by product group」；輸入區；Verified）。"
        "此口徑含手機覆晶等非 AI 產品，不等於 AI 先進封裝；層級側無封裝市場值，無可比實績（缺層級的先進封裝美元市場與艾克爾份額）")
    row("運算終端（資料中心、基礎設施、PC、儲存；US$M）", f"=E{r_cmp}", f"=D{r_cmp}*4", "無同口徑隱含值",
        "公司實績：10-Q 附註 2「net sales by end market」揭露運算占合併營收 1Q26 21%、2Q26 22%、3Q26 26%（整數百分比，Verified），美元值＝比重 × 合併營收（輸入區；Derived；與 Q 頁第 6 列同值）。"
        "運算含 AI 先進封裝與非 AI（PC、儲存、基礎設施），公司未拆；層級側無對應的運算終端或 OSAT 營收市場，無可比實績")
    row("AI 先進封裝（2.5D、HDFO）營收（US$M）", "無實績", "無實績", "無同口徑隱含值",
        "公司只在法說口述「FY2026 約為 FY2025 的三倍」與個別專案量產（4Q25、3Q26 法說；Interested-party），未揭露金額；10-Q、新聞稿都沒有 AI 先進封裝美元營收。"
        "本模型 FY2026 AI 先進封裝（Derived，由運算營收與「約三倍」說法反推）是模型值，不當實績。層級側 Siliconomics 沒有 AI 先進封裝的美元市場值。兩側皆缺，無可比實績")
    R["E_last"] = r - 1
    C.cell(r, 1, "參考（不判定）：外部估計 OSAT 先進封裝市場 vs 本模型 FY2026 AI 先進封裝營收（US$B）")
    c = C.cell(r, 2, f"=tam_gw*B{R['B2_first']}"); c.number_format = "0.0"; c.font = BLUE
    c = C.cell(r, 3, f"=PL!D{r_ai}/1000"); c.number_format = "0.00"
    C.cell(r, 4, "—")
    C.cell(r, 5, "參考（不判定）")
    C.cell(r, 6, "B 欄＝外部估計 OSAT 先進封裝可服務市場約 US$0.5B／GW（艾克爾投資人日；In 頁 tam_gw，Interested-party；第二來源未找到）× Siliconomics v0.19 SIF_DeployGW_Base CY2026（上方「層級之間」第一列，約 23.4 GW）；"
                 "該說法的投影片為 2025–2030E 口徑，套用到 CY2026 有時間口徑落差；C 欄＝本模型 FY2026 AI 先進封裝營收（情境 B，PL 頁，Derived：模型值、不是公司實績）。C ÷ B 見 G 欄（同 Recon 頁 C 節）。這是模型值對外部估計，不是公司實績對層級隱含值，不是 C0 證據、不判定。"
                 "若要可判定，需 Siliconomics 補 OSAT 先進封裝美元市場（或委外 CoWoS 每片價格）與艾克爾份額，或公司揭露 AI 先進封裝金額。")
    C.cell(r, 7, f'=TEXT(C{r}/B{r},"0.0%")&"（模型 ÷ 外部估計）"')
    C.cell(r, 8, "—")
    R["E_ref"] = r; r += 1
    C.cell(r, 1, "判定摘要"); R["E_sum"] = r
    rg = f'E{R["E_first"]}:E{R["E_last"]}'
    C.cell(r, 2, F(T("C0（實績）共 "), f'COUNTA({rg})', T(" 項：可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})', T(" 項、對帳中 "), f'COUNTIF({rg},"對帳中")', T(" 項、時差邊界 "), f'COUNTIF({rg},"時差邊界（參考）")',
                   T(" 項、無可比實績 "), f'COUNTIF({rg},"無可比實績")', T(" 項。原因：公司公布合併營收、產品組別（先進／主流）與終端市場比重，但未公布 AI 先進封裝（2.5D、HDFO）金額；Siliconomics 沒有封測服務或先進封裝的美元市場與艾克爾份額。"
                   "本模型不以模型預測或同源值補位；待層級補上 OSAT 市場與份額，或公司揭露 AI 先進封裝金額後再判定。")))
    R["E_sum_rg"] = rg
    r += 2
    assert r == r_in, (r, r_in)
    C.cell(r, 1, "輸入：公司 10-Q 營收（百萬美元；SEC 10-Q 附註 2 與合併損益表；Verified；藍字為輸入；終端市場比重為公司揭露的整數百分比）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "1Q26", "2Q26", "3Q26", "1–3Q 累計", "出處"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    srcq = f"1Q26：{Q1}（10-Q，2026-03-31）；2Q26：{Q2}；3Q26：{Q3}（附註 2 與合併損益表）；Verified。百萬美元＝千美元 ÷ 1,000"
    for lab_, vals, fmt in (("合併營收", (1684.701, 1897.965, 1990.557), "#,##0.000"), ("先進產品組（Advanced products）", (1372.001, 1556.676, 1631.580), "#,##0.000"),
                            ("主流產品組（Mainstream products）", (312.700, 341.289, 358.977), "#,##0.000")):
        C.cell(r, 1, lab_)
        for j, v in enumerate(vals):
            c = C.cell(r, 2 + j, v); c.font = BLUE; c.number_format = fmt
        C.cell(r, 5, f"=B{r}+C{r}+D{r}").number_format = fmt
        C.cell(r, 6, srcq); r += 1
    assert r - 3 == r_rev and r - 2 == r_adv and r - 1 == r_mai, (r, r_rev)
    C.cell(r, 1, "運算終端占合併營收（10-Q 附註 2 揭露，整數百分比）")
    for j, v in enumerate((0.21, 0.22, 0.26)):
        c = C.cell(r, 2 + j, v); c.font = BLUE; c.number_format = "0%"
    C.cell(r, 5, f"=E{r + 1}/E{r_rev}").number_format = "0.0%"
    C.cell(r, 6, "公司 10-Q 附註 2「net sales by end market」：運算（data center、infrastructure、PC/laptop、storage）1Q26 21%（1Q26 10-Q）、2Q26 22%（2Q26 10-Q）、3Q26 26%（3Q26 10-Q 附註 2）；1–3Q 累計欄＝美元值 ÷ 累計營收的加權比重（Derived）；Verified（整數，±0.5 個百分點）")
    r += 1
    C.cell(r, 1, "運算終端營收（比重 × 合併營收；Derived）")
    for j, col in enumerate("BCD"):
        C.cell(r, 2 + j, f"={col}{r_pct}*{col}{r_rev}").number_format = "#,##0.0"
    C.cell(r, 5, f"=B{r}+C{r}+D{r}").number_format = "#,##0.0"
    C.cell(r, 6, "Derived：整數比重 × 10-Q 合併營收；捨入誤差 ±0.5 個百分點（約 ±2%）；與 Q 頁 C6:E6 同值"); r += 1
    assert r - 1 == r_cmp
    C.cell(r, 1, "檢查：產品組別合計 − 合併營收（1Q26、2Q26、3Q26；應 ≈0）"); R["E_chk1"] = r
    for j, col in enumerate("BCD"):
        C.cell(r, 2 + j, f"=ROUND({col}{r_adv}+{col}{r_mai}-{col}{r_rev},6)").number_format = "0.000"
    r += 1
    C.cell(r, 1, "檢查：10-Q 合併營收 − Q 頁（1Q26、2Q26、3Q26；Q 頁為四捨五入至 0.1，應 < 0.1）"); R["E_chk2"] = r
    for j, (col, qc) in enumerate(zip("BCD", "CDE")):
        C.cell(r, 2 + j, f"=ROUND({col}{r_rev}-Q!{qc}4,3)").number_format = "0.000"
    r += 1
    C.cell(r, 1, "檢查：運算終端 比重 × 營收 − Q 頁運算（1Q26、2Q26、3Q26；應 < 0.1）"); R["E_chk3"] = r
    for j, (col, qc) in enumerate(zip("BCD", "CDE")):
        C.cell(r, 2 + j, f"=ROUND({col}{r_cmp}-Q!{qc}6,3)").number_format = "0.000"
    r += 1
    R["E_end"] = r - 1
    # ---- B 節判定摘要改指向 E 節計數 ----
    C.cell(R["B_sum"], 2, F(T("本公司 C0 對帳中項目 "), f'COUNTIF({rg},"對帳中")', T(" 項：公司實績對層級隱含值的 C0（實績）表見 E 節（可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})',
                             T(" 項；其餘因層級側沒有同口徑隱含值，標無可比實績）。本節加速器顆數 CY2026–29 是個股讀入的層級值，與層級同源，標「讀取層級（同源，不判定）」，不計入。"
                               "下方「層級之間」的產業部署 GW 差距只列參考，不計入。對帳單由主控與 Siliconomics 層處理，本模型不改上游值。")))
    return R


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, TO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; C = wb["Coord"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, coef = vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    pmed10 = vb("同業 10 年中位數的中位數")
    geps, gmul = vb("與市場共識目標價的差距：EPS 因素"), vb("與市場共識目標價的差距：倍數因素")
    bl = {s: lab(P, f"情境 {s}：") for s in "ABC"}
    rb = bl["B"]
    aish = {s: lab(P, "AI 先進封裝 ÷ 營收", bl[s]) for s in "ABC"}
    epr = {s: lab(P, "EPS（稀釋）", bl[s], exact=True) for s in "ABC"}
    gmr = {s: lab(P, "毛利率", bl[s], exact=True) for s in "ABC"}
    azi = lab(P, "亞利桑那對毛利率的影響", rb)
    aird = lab(P, "AI 先進封裝營收（", rb)
    revr = lab(P, "營收", rb, exact=True)
    cpr = lab(P, "運算（非 AI", rb); cmr = lab(P, "通訊（", rb); aur = lab(P, "車用、工業與其他", rb); csr = lab(P, "消費", rb, exact=True); azr = lab(P, "亞利桑那廠營收", rb)
    unr = lab(P, "AI 加速器顆數（含情境係數）", rb)
    tor = lab(SE, "加權目標價相對基準（HTML")
    tr = lambda k, col: f"Sens!{col}{lab(SE, k, tor)}"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.4.4 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0.0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', T(" 美元"), f'IF({up}>=0," 高 "," 低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "),
        f'TEXT({dcf},"0.0")', T(" 美元"), f'IF({dcf}<px,"低於","高於")', T("現價（資本支出重、獲利率低），同業比較法（相對倍數調整後）"), f'TEXT({cadj},"0.0")', T(" 美元；上檔取決於同業倍數是否維持。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：同業封測股目前的前瞻本益比 "), f'TEXT({pmed},"0.0")', T(" 倍（約為其十年中位數的 "), f'TEXT({pmed}/{pmed10},"0.0")', T(" 倍）不回落，且 AI 先進封裝依計畫成長。"
                 "若只看現金流折現，目標價 "), f'TEXT({dcf},"0.0")', T(" 美元（"), f'TEXT({dcf}/px-1,"+0.0%;-0.0%")', T("），評等"), f'{REC5(f"({dcf}/px-1)")}', T("。"))
    put("o1", T("目標價："), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', T(" 美元"), f'IF({up}>=0,"高 ","低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "), f'TEXT({dcf},"0.00")',
        T(" 美元與同業比較法（相對倍數調整後）"), f'TEXT({cadj},"0.00")', T(" 美元各占一半；市場共識 "), 'TEXT(cons_tp,"0.0")', T(" 美元（賣方平均）高於本模型，差距來自前瞻 EPS（本模型"), f'IF({geps}>=0,"高 ","低 ")', f'TEXT(ABS({geps}),"0%")',
        T("）與倍數（共識隱含倍數"), f'IF({gmul}>=0,"高於","低於")', T("同業中位數 "), f'TEXT(ABS({gmul}),"0%")', T("）。"))
    put("o2", T("核心問題一：AI 先進封裝占營收 FY2026 "), f'TEXT(PL!D{aish["B"]},"0.0%")', T(" → FY2028 三情境 "), f'TEXT(PL!F{aish["C"]},"0%")', T("～"), f'TEXT(PL!F{aish["A"]},"0%")', T("，分歧來自 AI 加速器顆數。FY2027 EPS "),
        f'TEXT(PL!E{epr["C"]},"0.00")', T("～"), f'TEXT(PL!E{epr["A"]},"0.00")', T(" 美元，FY2028 "), f'TEXT(PL!F{epr["C"]},"0.00")', T("～"), f'TEXT(PL!F{epr["A"]},"0.00")', T(" 美元；情境 C（顆數 ×0.6、通訊下滑）的 FY2028 EPS 比情境 A 低 "),
        f'TEXT(1-PL!F{epr["C"]}/PL!F{epr["A"]},"0%")', T("。"))
    put("o3", T("核心問題二：亞利桑那在 FY2027–FY2028 是成本、FY2029 才改善。亞利桑那毛利 ÷ 營收（情境 B）FY2027 "), f'TEXT(PL!E{azi},"+0.0%;-0.0%")', T("、FY2028 "), f'TEXT(PL!F{azi},"+0.0%;-0.0%")', T("、FY2029 "),
        f'TEXT(PL!G{azi},"+0.0%;-0.0%")', T("；毛利率三情境 FY2028 約 "), f'TEXT(PL!F{gmr["A"]},"0.0%")', T("／"), f'TEXT(PL!F{gmr["B"]},"0.0%")', T("／"), f'TEXT(PL!F{gmr["C"]},"0.0%")', T("。折舊與資本支出占營收高，現金流折現因此低於現價。"))
    put("o4", T("反向推算（情境 B）：要支持現價，折現率須"), f'IF({rv("w1")}<=wacc,"降到","升到")', T(" "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2027 起營收全面"), f'IF({rv("s1")}>=1,"放大 ","縮小為 ")', f'TEXT({rv("s1")},"0.00")',
        T(" 倍。同業比較法：現價相當於前瞻 EPS 的 "), f'TEXT(RevDCF!B{RVA["pe1"]},"0.0")', T(" 倍，"), f'IF(RevDCF!B{RVA["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍。"))
    put("o5", T("對帳中年份（CY2027–29，Siliconomics 標示；層與層的差距由對帳單處理，非本公司 C0 項目）：另一側讀數目標價 "), f'TEXT(Coord!C{rtp},"0.00")', T(" 美元，評等"), f'Coord!C{rrec}', f'IF(Coord!B{rrec}<>Coord!C{rrec},"（與現行基準讀數"&Coord!B{rrec}&"評等不同）","")',
        T("。最敏感：現金流通率 30%／50% 使加權目標價 "), f'TEXT({tr("現金流通率", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("現金流通率", "C")},"+0%;-0%")', T("，AI 部署 ×0.8／×1.2 使 "),
        f'TEXT({tr("AI 部署顆數", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("AI 部署顆數", "C")},"+0%;-0%")', T("。"))
    put("drv", T("分部營收推導（情境 B）：合併營收＝AI 先進封裝＋非 AI 運算＋通訊＋車用與工業＋消費＋亞利桑那廠。AI 先進封裝＝AI 加速器顆數 × 後段委外比例 × 艾克爾份額 × 每顆封測金額（隨封裝面積上升，FY2026 校準）；通訊、車用與工業、消費、非 AI 運算依成長率；"
                 "亞利桑那廠第一期依情境營收。FY2028：AI 先進封裝 "), f'TEXT(PL!F{aird},"#,##0")', T("、非 AI 運算 "), f'TEXT(PL!F{cpr},"#,##0")', T("、通訊 "), f'TEXT(PL!F{cmr},"#,##0")', T("、車用與工業 "), f'TEXT(PL!F{aur},"#,##0")',
        T("、消費 "), f'TEXT(PL!F{csr},"#,##0")', T("、亞利桑那 "), f'TEXT(PL!F{azr},"#,##0")', T("，合計營收 "), f'TEXT(PL!F{revr},"#,##0")', T(" 百萬美元。"))
    out["v25"] = r
    RD.cell(r, 1, f"v0.4.4（{BUILT}）：只改展示層與新增列——個股 HTML 版面規格 v1.3（章節：命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態）；三層協調（規則 v1.4.2）：Coord 頁新增 A 讀哪一層哪一版、B 讀取層級（同源，不判定）、C 口徑註記、D 摘要、E C0（實績）：公司實績 vs 層級隱含值，"
                  "並列兩讀數目標價（C4；主引擎情境 B、C 於 CY2027 起產業量 × 另一側讀數比值；情境 A 不縮放）。新增 Hist 頁（損益表歷史 FY2022–FY2025）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B，GAAP）、Val 0.7（四個目標價）、"
                  "Sens 頁末加權目標價相對基準（營運輸入案例為整本重算快照 amkr/sens.json）、README 頁末 v1.3 白話列；README 標題段版號更正為 v0.4.4。既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異）。")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT or "Coord" not in wb.sheetnames:   # 另一側讀數建置（alt_run.py）只讀 Val 頁，不加新頁
        return "v044_post：略過（另一側讀數建置或無 Coord 頁）"
    HR = hist_is(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_tor(wb, here)
    CR = coord_c0(wb, SIL)
    RMA = readme13(wb, RVA, PLA, ST)
    RD = wb["README"]
    a1 = RD["A1"].value
    assert "收支與估值模型 v0.2.2（AI 半導體" in a1, a1[:80]
    RD["A1"].value = a1.replace("收支與估值模型 v0.2.2（AI 半導體", f"收支與估值模型 v{VER}（AI 半導體", 1)
    a1 = RD["A1"].value
    assert "\u30002026-10-25；" in a1
    RD["A1"].value = a1.replace("\u30002026-10-25；", f"\u3000{BUILT}；", 1) + "；版面 v1.3 與 C0（實績）"
    a2 = str(RD["A2"].value)
    assert a2.startswith("版本：v0.2.1（2026-10-25："), a2[:40]
    RD["A2"].value = a2.replace("版本：v0.2.1（", f"版本：v{VER}（{BUILT}：版面規格 v1.3、三層協調 C0 與兩讀數；前版 v0.4.3，2026-11-06）；v0.2.1（", 1)
    order = ["README", "In", "Q", "Hist", "Units", "PL", "Val", "DCF", "RevDCF", "Recon", "Sens", "Sources", "Coord"]
    wb._sheets = [wb[n] for n in order]
    return f"v044_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 相對基準 {ST['first']}–{ST['last']}；Coord E 節 {CR['E_first']}–{CR['E_last']}；README v1.3 {len(RMA)} 列"
