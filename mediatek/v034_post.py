# 聯發科 v0.3.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0 公司實績、兩讀數；規則 v1.4.2；工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）。
# 樣板 amkr/v044_post.py（Hist、標準損益表、RevDCF、Val 0.7、Sens 相對基準、Coord）。只新增，不改既有數值格：
#   Hist 頁（新）、PL 頁末標準損益表、RevDCF 頁（新）、Val 頁末 0.7、Sens 頁末加權目標價相對基準（整本重算快照 mediatek/sens.json）、Coord 頁 A–E 節、README 頁末 v1.3 白話列。
# build_mediatek.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。另一側讀數建置（COORD_RUN＝1）與 sens_mediatek.py 建置（MTK_SENS＝1）時不執行。
import json, os
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3.5", "2026-11-15"   # v0.3.5：資產負債表改公司 3Q26 合併財報、債務加計零息海外可轉債（版本列與 Sens 可轉債案例；其餘同 v0.3.4）
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
MK = "https://www.mediatek.com/hubfs/MediaTek%20Assets/Pdfs/"
S25 = MK + "Financial%20Reports/2025/Q4%20Consolidated%20Report.pdf"
S25R = MK + "Quarterly%20Earnings%20Release/2025/Quarterly%20Earnings%20Release-2025Q4/Press%20Release.pdf"
SDJ = "https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcq/zcqa_2454.djhtm"
SMS = "https://ca.marketscreener.com/quote/stock/MEDIATEK-INC-6495333/news/MediaTek-Inc-Reports-Earnings-Results-for-the-Full-Year-Ended-December-31-2023-45970394"
SQ3 = "https://cdnfinance.technews.tw/2026/10/30/mediateks-asic-revenue-will-reach-2-5-billion-in-2026/"
SQ3b = "https://uanalyze.com.tw/articles/6233958069"
SOCT = "https://technews.tw/2026/11/10/mediateks-october-revenue-increased-by-5-8-year-on-year/"
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


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
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) and p[0] == "t" else (p[0] if isinstance(p, tuple) else p) for p in parts)


T = lambda s: ("t", s)

# ---------- Hist：損益表歷史（NT$B；TIFRS 合併損益表） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025 公式)
HIS = [
    ("rev", "營收", 548.796, 433.446, 530.586, "=Q!B4"),
    ("cogs", "營業成本", 277.892, 226.079, 267.200, "=Q!B4-Q!B10"),
    ("gp", "毛利", 270.904, 207.367, 263.386, "=Q!B10"),
    ("opx", "營業費用（銷管＋研發）", 144.116, 135.568, 160.974, "=Q!B12"),
    ("oi", "營業利益", 126.788, 71.800, 102.412, "=Q!B13"),
    ("nonop", "營業外淨額", 8.773, 14.983, 17.107, "=Q!B14"),
    ("pbt", "稅前利益", 135.561, 86.782, 119.519, "=Q!B15"),
    ("tax", "所得稅", 16.936, 9.591, 12.378, "=Q!B16"),
    ("nic", "稅後淨利（合併，含非控制權益）", 118.625, 77.191, 107.141, "=Q!B15-Q!B16"),
    ("nci", "非控制權益損益", "=C{nic}-C{ni}", "=D{nic}-D{ni}", 0.754, "=Q!B17"),
    ("ni", "歸屬母公司淨利", 118.141, 76.979, 106.387, "=Q!B18"),
    ("eps", "EPS（基本；NT$）", 74.59, 48.51, 66.92, "=Q!B21"),
    ("epsd", "EPS（完全稀釋，公司公告；NT$）", 74.23, 48.34, 66.78, 66.03),
]


def hist_is(wb):
    Hs = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    Hs["A1"] = ("Hist — 損益表歷史（v0.3.4 起；十億新台幣；EPS 新台幣；TIFRS 合併損益表、歸屬母公司）。PL 頁末「標準損益表」的歷史三年（FY2023–FY2025）與年增率基期（FY2022）用。FY2025 欄＝Q 頁。只被 PL 頁標準損益表引用，不回饋任何計算")
    Hs["A1"].font = BOLD
    Hs["A2"] = (f"來源：FY2024、FY2025＝公司 2025 年 Q4 合併財報（{S25}；Verified，與 Q4 新聞稿 {S25R} 相同）。FY2022、FY2023＝富邦 e 點通（DJ）個股財報（{SDJ}；資料商，Interested-party），"
                f"營收、歸屬母公司淨利、基本 EPS 與公司原文或 MarketScreener 2023 年度業績（{SMS}）交叉核對一致；營業成本、費用、稅等細項只有資料商一個來源。非控制權益損益 FY2022、FY2023＝稅後淨利 − 歸屬母公司淨利（Derived）")
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025（Q 頁）", "標記", "來源"], 1):
        c = Hs.cell(4, j, h); c.font = BOLD; c.fill = HDR
    rows = {}; r = 5
    for k, *_ in HIS:
        rows[k] = r; r += 1
    r = 5
    for k, labl, v22, v23, v24, v25 in HIS:
        Hs.cell(r, 1, f"h_{k}"); Hs.cell(r, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if isinstance(v, str) and "{" in v:
                v = v.format(**rows)
            c = Hs[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k in ("eps", "epsd") else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"h_{k}_{yy}", f"Hist!${col}${r}")
        Hs.cell(r, 7, "Verified（FY2024–25）；Interested-party（FY2022–23 細項）")
        Hs.cell(r, 8, "FY2024–25：公司 Q4 合併財報；FY2022–23：富邦 DJ 財報頁（營收、淨利、EPS 另與公司原文交叉核對）")
        r += 1
    r += 1
    Hs.cell(r, 2, "加權股數（歸屬母公司淨利 ÷ 基本 EPS；十億股，Derived）")
    for col in "CDEF":
        Hs[f"{col}{r}"] = f"={col}{rows['ni']}/{col}{rows['eps']}"; Hs[f"{col}{r}"].number_format = "0.0000"
    rows["shr"] = r; dn(wb, "h_shr_23", f"Hist!$D${r}"); dn(wb, "h_shr_24", f"Hist!$E${r}"); dn(wb, "h_shr_22", f"Hist!$C${r}"); r += 2
    for labl, f in (("核對：營收 − 營業成本 − 毛利（應 ≈0）", lambda c: f"={c}{rows['rev']}-{c}{rows['cogs']}-{c}{rows['gp']}"),
                    ("核對：毛利 − 營業費用 − 營業利益（資料商四捨五入，|差| < 0.01）", lambda c: f"={c}{rows['gp']}-{c}{rows['opx']}-{c}{rows['oi']}"),
                    ("核對：營業利益 ＋ 營業外淨額 − 稅前利益（|差| < 0.01）", lambda c: f"={c}{rows['oi']}+{c}{rows['nonop']}-{c}{rows['pbt']}"),
                    ("核對：稅前利益 − 所得稅 − 稅後淨利（|差| < 0.01）", lambda c: f"={c}{rows['pbt']}-{c}{rows['tax']}-{c}{rows['nic']}")):
        Hs.cell(r, 2, labl)
        for col in "CDEF":
            Hs[f"{col}{r}"] = f(col); Hs[f"{col}{r}"].number_format = "0.000"
        r += 1
    Hs.column_dimensions["A"].width = 12; Hs.column_dimensions["B"].width = 52
    for col in "CDEF": Hs.column_dimensions[col].width = 14
    S = wb["Sources"]; n = S.max_row + 1
    for i, (a, b) in enumerate((("公司：2025 年 Q4 合併財報（FY2025、FY2024 合併損益表，Hist 頁）", S25), ("公司：2025 年 Q4 新聞稿（FY2025、FY2024）", S25R),
                                ("資料商：富邦 e 點通 個股財報（FY2022–FY2025 損益表，Hist 頁 FY2022–23 欄）", SDJ), ("媒體／資料商：MarketScreener 2023 年度業績（FY2023、FY2022 營收、淨利、EPS 交叉核對）", SMS),
                                ("媒體：科技新報 2026-10-30（3Q26 法說：首顆 AI ASIC 9 月出貨、4Q26 雲端 ASIC 預估 25 億美元；Coord 頁 E 節）", SQ3),
                                ("媒體：科技新報 2026-11-10（10 月營收 550.25 億元；前 10 月累計 5,244.16 億元；Coord 頁 E 節）", SOCT))):
        S.cell(n + i, 1, a); S.cell(n + i, 2, b)
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"
SRC = "CDEFGHI"


def pl_standard(wb):
    P = wb["PL"]; D = wb["DCF"]
    rbd = lab(P, "情境 B（基準）：")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    R = dict(rev=ex("營收"), gp=ex("毛利"), opx=ex("營業費用"), oi=ex("營業利益"), oth=ex("營業外淨額"), pbt=ex("稅前利益"), tax=ex("所得稅"), nci=ex("非控制權益"),
             ni=ex("歸屬母公司淨利"), sh=ex("加權股數"), eps=ex("EPS（基本）"))
    dbr = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, "標準損益表（v0.3.4 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026E–FY2031E，情境 B 基準；FY2026E 為錨點）。十億新台幣；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（台灣採用之國際財務報導準則 TIFRS 合併財報、歸屬母公司淨利）；模型 EPS 與目標價皆用此口徑，沒有 Non-GAAP 口徑。EPS 用基本 EPS（本模型口徑；公司另公告完全稀釋 EPS，見 Hist 頁，FY2025 為 66.03）。預測期的「其他收入（淨額）」＝營業外淨額")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（1H26 實績＋3Q26 營收實績與指引＋4Q26 假設）；FY2027E–FY2029E＝情境 B 由驅動因子推導；FY2030E、FY2031E＝延伸期（收斂法只推營收與歸屬母公司淨利）：毛利率與營業費用率沿用 FY2029E 推算營業成本、毛利與營業費用；"
                 "營業利益取 DCF 頁情境 B（由淨利反推）；其他收入＝營業外淨額假設；所得稅＝稅前利益 −（歸屬母公司淨利＋非控制權益）。兩種推法的營業利益差額列在表末核對列，不改既有淨利與目標價")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位", *YEARS, "說明"], 1):
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
    row("rev", "營收", "NT$B", hh("rev") + [f"=PL!{c}{R['rev']}" for c in SRC], M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L), "=D{rev}/C{rev}-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS[1:], COLS[2:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", hh("cogs") + [f"=PL!{c}{R['rev']}-PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-I{r_c + 2})"], M, "FY2030E、FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", hh("gp") + [f"=PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（銷管＋研發）", "NT$B", hh("opx") + [f"=PL!{c}{R['opx']}" for c in SRC[:5]] + [f"=J{L['rev']}*I{r_c + 3}/I{L['rev']}", f"=K{L['rev']}*I{r_c + 3}/I{L['rev']}"], M, "FY2030E、FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + [f"=PL!{c}{R['oi']}" for c in SRC[:5]] + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外淨額）", "NT$B", hh("nonop") + [f"=PL!{c}{R['oth']}" for c in SRC[:5]] + ["=nonop_f", "=nonop_f"], M, "利息、投資收益等；FY2030E 起＝營業外淨額假設")
    row("pbt", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "NT$B", hh("tax") + [f"=PL!{c}{R['tax']}" for c in SRC[:5]] + [f"=J{L['pbt']}-(PL!H{R['ni']}+nci)", f"=K{L['pbt']}-(PL!I{R['ni']}+nci)"], M, "FY2030E、FY2031E＝稅前利益 −（歸屬母公司淨利＋非控制權益）")
    row("nic", "稅後淨利（合併，含非控制權益）", "NT$B", [f"={c}{L['pbt']}-{c}{L['tax']}" for c in COLS], M)
    row("nci", "非控制權益損益", "NT$B", hh("nci") + [f"=PL!{c}{R['nci']}" for c in SRC[:5]] + ["=nci", "=nci"], M)
    row("ni", "歸屬母公司淨利（EPS 與目標價口徑）", "NT$B", [f"={c}{L['nic']}-{c}{L['nci']}" for c in COLS], M)
    row("nm", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "加權股數", "十億股", ["=h_shr_23", "=h_shr_24"] + [f"=PL!{c}{R['sh']}" for c in SRC], "0.000", "FY2023–FY2024＝歸屬母公司淨利 ÷ 基本 EPS（Hist 頁，Derived）")
    row("eps", "EPS（基本；NT$）", "NT$", ["=h_eps_23", "=h_eps_24"] + [f"=PL!{c}{R['eps']}" for c in SRC], "0.00", "FY2023–FY2024＝公司公告基本 EPS；FY2025 起＝模型（歸屬母公司淨利 ÷ 基本加權股數）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY；歸屬母公司淨利）", "%", ["=C{ni}/h_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 歸屬母公司淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；公司公告完全稀釋 EPS，與上方基本 EPS 口徑不同，只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_epsd", "EPS（完全稀釋，公司公告；NT$）", "NT$", ["=h_epsd_23", "=h_epsd_24", "=Hist!F{}".format(lab(wb["Hist"], "h_epsd", col=1))], "0.00", "FY2025 公司公告 66.03（Hist 頁）")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023–FY2024 為資料商四捨五入，|差| < 0.01；FY2030E、FY2031E 見下兩列）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：歸屬母公司淨利 − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$B", [None, None] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    row("c4", "核對：EPS − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$", [None, None] + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]; P = wb["PL"]; V = wb["Val"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"), dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    rbd = lab(P, "情境 B（基準）："); rev26 = lab(P, "營收", rbd, exact=True)
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncv = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.4 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In、PL 頁，無任何頁引用本頁，不影響目標價。金額十億新台幣；價格新台幣")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。FY2027–FY2032 全部為預測年（FY2026 為錨點，不動）")
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
    kv("px", "現價（新台幣）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（十億新台幣）", f"={ncv}", "#,##0.0", "同 Val 頁 0.1 節（淨現金與投資＋stub 期 FCFF − stub 期股利；三情境相同）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.0000")
    kv("evr", "支持現價所需的企業價值 EV（十億新台幣）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；十億新台幣）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（新台幣）", f"=DCF!$C${ps}", "0.00")
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
    yr("rev", "營收（十億新台幣）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（十億新台幣）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻", lambda dc, c: f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（十億新台幣）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（十億新台幣）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（十億新台幣，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2027–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（十億新台幣）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（十億新台幣）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2027–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(B{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(B{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2027–FY2032 營收等比例放大；FY2026 錨點不動；營業利益率、稅率、折舊、資本支出占營收比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（十億新台幣）", f"=fc_wc*PL!$D${rev26}*B{Y['fr']}*B{Y['df']}", "#,##0.0", "FY2027 營運資金增加中屬 FY2026 營收的部分（FY2026 為錨點，不放大）")
    kv("b", "隨營收倍數等比例變動的 EV（十億新台幣）",
       f"=SUMPRODUCT(B{Y['cf']}:G{Y['cf']},B{Y['fr']}:G{Y['fr']},B{Y['df']}:G{Y['df']})-fc_wc*B{Y['rev']}*B{Y['fr']}*B{Y['df']}"
       f"-SUMPRODUCT(C{Y['wc']}:G{Y['wc']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2027 營收（十億新台幣）", f"=B{A['s1']}*B{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2027 營收（情境 B；十億新台幣）", f"=B{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2027 營收年增率（相對 FY2026 錨點營收）", f"=B{A['rv1']}/PL!$D${rev26}-1", "0.0%", "FY2026 為錨點，不放大")
    kv("g0", "模型：FY2027 營收年增率（情境 B）", f"=B{A['rv0']}/PL!$D${rev26}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（新台幣）"], 1):
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
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（新台幣，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"#,##0")',
                     T(" 元；要支持現價 "), 'TEXT(px,"#,##0")', T(" 元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大為","縮小為")', T(" "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<=wacc,"降到","升到")', T(" "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2027–FY2032 營收全面"), f'IF(B{A["s1"]}>=1,"放大","縮小")',
                     T(" "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2027 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億元，相對 FY2026 錨點年增 "), f'TEXT(B{A["g1"]},"0%")',
                     T("）；或營業利益率"), f'IF(B{A["m1"]}>=B{A["m0"]},"提高到","降到")', T(" "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"#,##0")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 聯發科過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給聯發科的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，比模型採用的相對倍數"), f'IF(D{A["rc1"]}<0,"低 ","高 ")', f'TEXT(ABS(D{A["rc1"]}),"0%")', T("。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：估值並列四個目標價 ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.3.4 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（賣方平均）", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here):
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    bl = {s: lab(P, f"情境 {s}（") for s in "ABC"}   # 摘要區塊起點（第 4、18、32 列）
    bl = {s: lab(P, f"情境 {s}（", 60) for s in "ABC"}
    epr = {s: lab(P, "EPS（基本）", bl[s], exact=True) for s in "ABC"}
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.3.4 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（mediatek/sens_mediatek.py：三情境同時套用案例覆寫，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（新台幣）", "高案：加權目標價（新台幣）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1; first = r
    for k, d in J["biz"].items():
        SE.cell(r, 1, d["label"])
        for col, key in ((4, "低"), (5, "高")):
            if d[key] is None: continue
            c = SE.cell(r, col, d[key]["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
            SE.cell(r, col - 2, f"={'D' if col == 4 else 'E'}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（mediatek/sens.json）；覆寫同本頁上方「一次動一個輸入」案例，三情境同時套用"); r += 1
    frame = [
        ("同業比較法不做相對倍數調整（改用市場主流同業法；只有低案）", f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}", None, "Val 頁：Comparable（未調整）取代調整後值"),
        ("DCF 在加權中的權重 0%（只看同業比較法）／100%（只看現金流折現）（基準＝50%）", f"={cadj}", f"={dcf}", "Val 頁：Comparable（調整後）與 DCF 目標價"),
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*(2/12*PL!E{epr['C']}+10/12*PL!F{epr['C']})",
         f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*(2/12*PL!E{epr['A']}+10/12*PL!F{epr['A']})", "Val 頁：單一情境的 DCF 價值與 PL 頁該情境前瞻 EPS"),
    ]
    rcb = {k: lab(SE, k) for k in ("可轉債以面額計債務", "可轉債視為全數轉換")}   # v0.3.5：build_mediatek.py「可轉債處理敏感度」表（C＝加權目標價）
    frame.append(("可轉債處理：以面額計債務（低）／視為全數轉換（高）（基準＝帳面計債務、不計稀釋）", f"=Sens!C{rcb['可轉債以面額計債務']}", f"=Sens!C{rcb['可轉債視為全數轉換']}",
                  "本頁「可轉債處理敏感度」表（閉式；D/E、β、WACC 隨債務與股數重算）"))
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl)
        for col, f in ((4, lo), (5, hi)):
            if f is None: continue
            SE.cell(r, col, f).number_format = "#,##0.00"
            SE.cell(r, col - 2, f"={'D' if col == 4 else 'E'}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, note); r += 1
    TOR_PARAM = (("WACC", "WACC −1／＋1 個百分點（永續期新投資報酬率隨 WACC）"), ("β", "股權 β 低／高（產業無槓桿 β ∓0.3 後重新加槓桿）"), ("ERP", "股權風險溢酬 4%／6%（基準 5%）"),
                 ("永續成長率", "永續成長率 −1 個百分點／上限（基準 g＝新台幣 10 年期公債殖利率）"), ("RONIC", "永續期新投資報酬率高於 WACC 5 個百分點（只有高案）"), ("資本支出", "每年資本支出占營收 −1／＋1 個百分點"),
                 ("營運資金", "營運資金占營收 0%／15%（基準 8.1%）"), ("淨現金不含", "淨現金不含權益投資（只有低案）"), ("終值改用", "DCF 終值改用 FCFF ×（1＋g）法（只有高案）"), ("美元口徑", "美元口徑折現（rf 改美國 10 年期；只有低案）"))
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
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 mediatek/sens_mediatek.py）"); r += 1
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調（A 讀哪一層、B 讀取層級、C 口徑、D 摘要、E C0 實績） ----------
def coord_c0(wb, SIL):
    C = wb["Coord"]; P = wb["PL"]; I = wb["In"]
    rbA = lab(P, "情境 A（樂觀）：")
    ex = lambda t, s: lab(P, t, s, exact=True)
    r_xraw = lab(P, "AI ASIC 營收（未校準）", rbA, exact=True)
    r_utpu = lab(P, "Google TPU 市場顆數（含情境係數）", rbA)
    C.column_dimensions["F"].width = 60
    for col in "BCDE": C.column_dimensions[col].width = 24
    r_ratio = lab(C, "另一側讀數 ÷ 現行基準")
    r = C.max_row + 3
    C.cell(r, 1, "三層協調（v0.3.4 新增；三層基準協調規則 v1.4.2 C0 個股與層的一致性、C4 個股不裁決、C5 往上回饋只有證據）").font = BOLD; r += 1
    R = {}
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    C.cell(r, 1, "A. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; R["A"] = r; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "讀自（層與版本）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["A_hdr"] = r; r += 1

    def arow(label, cells, src, fmt="0.000"):
        nonlocal r
        C.cell(r, 1, label)
        for j, v in enumerate(cells):
            if v is None: continue
            C.cell(r, 2 + j, v).number_format = fmt
        C.cell(r, 6, src); r += 1
    arow("Google TPU 市場顆數（情境 A 原始顆數；百萬顆）", [f"=PL!{c}{r_utpu}" for c in "DEFG"], f"{ver}：SIF_UnitsPlat_tpu，經台積電個股模型 v0.5.4 成品（tsmc/dist；該模型讀 Siliconomics v0.20）取得（In 頁 u_tpu_26–28、g29_tpu；唯讀）")
    arow("聯發科占 Google TPU 顆數份額", ["=mtk_tpu_26", "=mtk_tpu_27", "=mtk_tpu_28", "=mtk_tpu_28"], f"{ver}：SIF_mtk_tpu_26／27／28（In 頁 M7；CY2029 沿用 CY2028；Assumed）")
    arow("TPU 每顆核心售價（US$K）", ["=core_tpu"] * 4, f"{ver} SIF_core_tpu，經 Broadcom 個股模型 v0.5.7（avgo/dist）取得（In 頁 core_tpu；唯讀）", "0.0")
    arow("Wattonomics 錨定係數（只進 Sens 案例，不改基準）", [None] + [f"=wta_{y}" for y in (27, 28, 29)], f"{ver}：SIF_WTAnchor_Base（R4 比值＝Wattonomics ÷ 本庫推得上線 GW）；Wattonomics 現行版本 v0.3.2")
    arow("另一側讀數 ÷ 現行基準（對帳中年份；兩讀數目標價用）", [f"={'BCDE'[i]}{r_ratio}" for i in range(4)], f"{ver}：SIF_DeployGW_AltRatio_<年>；上方 Coord 第 6 列")
    C.cell(r, 1, "Wattonomics：本模型不直接讀 Wattonomics 數字；現行版本 v0.3.2，僅經 Siliconomics 的 R4 比值（Sens 錨定案例）與另一側讀數比值（對帳中年份兩讀數）間接使用。其餘產業量皆讀 Siliconomics 基準（經台積電、Broadcom 成品）。"); R["A_note"] = r; r += 2
    # ---- B. 讀取層級 ----
    C.cell(r, 1, "B. 讀取層級（同源，不判定）：聯發科讀入的 Siliconomics 值 vs 層級值（個股自層級讀入，差距恆為 0，無檢驗力；不計入對帳中項目數；規則 v1.4.2 C0 第 2 點）。"
                 "公司實績對層級隱含值的對帳見本頁 E 節「C0（實績）」").font = BOLD; R["B"] = r; r += 1
    for j, h in enumerate(["項目（年）", "本模型原值", "層級對應值", "差距（本模型 ÷ 層級 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["B_hdr"] = r; r += 1
    units_L = SIL.NAMES["SIF_UnitsPlat_tpu"]; gw_b = SIL.NAMES["SIF_DeployGW_Base"]
    R["B_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        C.cell(r, 1, f"Google TPU 市場顆數 CY{y}（百萬顆）")
        C.cell(r, 2, f"=PL!{'DEFG'[i]}{r_utpu}").number_format = "0.000"
        c = C.cell(r, 3, units_L[i + 1]); c.number_format = "0.000"; c.font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, f"本模型：PL 頁情境 A 的 TPU 市場顆數（讀台積電個股成品）；層：{ver} SIF_UnitsPlat_tpu。同源讀取故差距 0；聯發科不揭露 TPU 顆數（等級：Interested-party）" if i == 0 else "同上")
        r += 1
    for i, y in enumerate((26, 27, 28)):
        C.cell(r, 1, f"聯發科占 Google TPU 顆數份額 CY20{y}")
        C.cell(r, 2, f"=mtk_tpu_{y}").number_format = "0.0%"
        c = C.cell(r, 3, SIL.NAMES[f"SIF_mtk_tpu_{y}"]); c.number_format = "0.0%"; c.font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, f"本模型：In 頁 mtk_tpu_{y}（自 Siliconomics 讀入）；層：{ver} SIF_mtk_tpu_{y}。Assumed；Siliconomics 該值於 v0.3 起取自本模型原寫死值（Siliconomics 盤點：mediatek In!42–44），故非獨立估計" if i == 0 else "同上")
        r += 1
    C.cell(r, 1, "TPU 每顆核心售價（US$K）"); C.cell(r, 2, "=core_tpu"); c = C.cell(r, 3, SIL.NAMES["SIF_core_tpu"]); c.font = BLUE
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"; C.cell(r, 5, "讀取層級（同源，不判定）")
    C.cell(r, 6, "本模型讀 Broadcom 個股模型 v0.5.7 的 In 頁 core_tpu；層：SIF_core_tpu（Assumed）；聯發科負責的世代是否與 Broadcom 同價：未找到"); r += 1
    R["B_last"] = r - 1
    C.cell(r, 1, "判定摘要"); R["B_sum"] = r; r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; R["B2"] = r; r += 1
    for j, h in enumerate(["項目（年）", "Siliconomics 基準", "外部估計（Wattonomics 換算）", "差距（基準 ÷ 外部 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    R["B2_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        C.cell(r, 1, f"產業部署 GW CY{y}（Siliconomics 基準 ÷ Wattonomics 換算；GW）")
        c = C.cell(r, 2, gw_b[i]); c.number_format = "0.00"; c.font = BLUE
        C.cell(r, 3, f"=B{r}*{'BCDE'[i]}{r_ratio}").number_format = "0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "參考")
        C.cell(r, 6, (f"基準＝Siliconomics {SIL.VERSION} 的供給側部署 GW；外部估計（參考）＝Wattonomics 換算的需求側部署 GW。這是層與層之間的差距，不是聯發科自己的量對層級的比較；"
                      "由 Siliconomics 對帳單 r3 處理，不計入本公司 C0 項目數；本模型依 C4 並列兩讀數目標價，不裁決") if i == 0 else "同上")
        r += 1
    R["B2_last"] = r - 1
    r += 1
    # ---- C ----
    C.cell(r, 1, "C. 對帳中年份處理與口徑註記（C4）").font = BOLD; R["C"] = r; r += 1
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    C.cell(r, 1, F(T("對帳中年份：CY2027、CY2028、CY2029（Siliconomics 標示；本公司 C0 無可判定項目）。兩讀數目標價並列於上方（現行基準讀數 "), f'TEXT(B{rtp},"#,##0.00")', T(" 元、另一側讀數 "), f'TEXT(C{rtp},"#,##0.00")',
                    T(" 元），評等分別為"), f'B{rrec}', T("、"), f'C{rrec}', T("。個股不裁決（C4）；Siliconomics 與 Wattonomics 的差距由對帳單處理（期限 2027-02-28）；結案前兩個目標價並列。")))
    R["C_1"] = r; r += 1
    C.cell(r, 1, F(T("口徑註記（不改模型）：另一側讀數只改情境 B、C 的聯發科 AI ASIC 顆數（CY2027 起乘比值 "), f'TEXT(C{r_ratio},"0.00")', T("／"), f'TEXT(D{r_ratio},"0.00")', T("／"), f'TEXT(E{r_ratio},"0.00")',
                    T("）；情境 A（樂觀，權重 "), 'TEXT(tp_wA,"0%")', T("）不縮放；手機、智慧邊緣平台、電源 IC 與各項成本假設不變；ASIC 校準係數由 FY2026 校準、不受影響，故 ASIC 營收約隨顆數等比例下降。這個口徑已含在另一側目標價內；對帳單處理時一併說明。")))
    R["C_2"] = r; r += 1
    C.cell(r, 1, "C0 結論：本模型的產業量（TPU 顆數）、聯發科份額與每顆核心售價都來自 Siliconomics（份額於 v0.3 起取自本模型原值，非獨立估計），列為「讀取層級」、不判定。公司實績（月營收、季報）已列入 E 節，"
                 "但聯發科至今沒有揭露 AI ASIC 營收金額（3Q26 僅口述 9 月開始出貨初期訂單；首顆 ASIC 4Q26 量產，4Q26 雲端 ASIC 營收 25 億美元為指引），層級也沒有聯發科合併營收，故 C0（實績）全為無可比實績，不產生對帳中項目。"
                 "往上回饋只有證據（C5）：待 4Q26 實績（2027-01 月營收與 2027-02 法說）揭露 ASIC 營收後，才能對層級隱含值判定。"); R["C_3"] = r; r += 2
    # ---- D ----
    C.cell(r, 1, "D. 摘要（HTML 與版面檢查用；兩讀數與損益表年度標示）").font = BOLD; R["D"] = r; r += 1
    for j, h in enumerate(["項目", "現行基準讀數", "另一側讀數", "另一側相對現行", ""], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    r_up0 = lab(C, "上檔空間（加權目標價")
    C.cell(r, 1, "加權目標價"); C.cell(r, 2, f"=B{rtp}").number_format = "0.00"; C.cell(r, 3, f"=C{rtp}").number_format = "0.00"; C.cell(r, 4, f"=D{rtp}").number_format = "+0.0%;-0.0%"; r += 1
    C.cell(r, 1, "上檔空間（加權目標價 ÷ 現價 − 1）"); C.cell(r, 2, f"=B{r_up0}").number_format = "+0.0%;-0.0%"; C.cell(r, 3, f"=C{r_up0}").number_format = "+0.0%;-0.0%"; r += 1
    C.cell(r, 1, "投資建議（五級；兩讀數）"); C.cell(r, 2, f"=B{rrec}"); C.cell(r, 3, f"=C{rrec}"); r += 1
    R["pl"] = r; C.cell(r, 1, "損益表年度標示（對帳中年份欄標題；FY＝CY，FY2030E、FY2031E 為延伸期，承接 CY2029 水準）")
    for j, t in enumerate(["FY2027E（對帳中）", "FY2028E（對帳中）", "FY2029E（對帳中）", "FY2030E（對帳中；延伸期）", "FY2031E（對帳中；延伸期）"]):
        C.cell(r, 2 + j, t)
    r += 2
    # ---- E ----
    C.cell(r, 1, "E. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格；差距超過 ±15% 且兩讀數同號標「對帳中」）").font = BOLD
    R["E"] = r; r += 1
    C.cell(r, 1, "比較方法：個股側只用公司已公布實績（月營收、季報；Verified）；層級側用 Siliconomics v0.20 同口徑隱含值。時差：層級值為日曆年 CY2026，聯發科會計年度＝日曆年；主讀數＝2026 年已公布期間累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12；"
                 "第二讀數＝最近一季 × 4 vs 層級 CY2026。判定：兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致；公司側或層級側沒有同口徑值＝無可比實績（不以模型預測、指引或同源值補位）。"
                 "AI ASIC 營收：聯發科至今未揭露金額（首顆 ASIC 4Q26 才量產，線性時差近似也不適用：全年金額幾乎全落在 4Q），故公司側無實績；層級側隱含值＝TPU 顆數 × 聯發科份額 × 每顆營收（核心售價＋代購 HBM），"
                 "取「校準前」值（本模型以係數 k 把 FY2026 校準到公司資料中心說法，校準後的 CY2026 值與公司說法同值，屬同源、無檢驗力，不採用）。公司指引（4Q26 雲端 ASIC 25 億美元、2027 年目標）、賣方估計只列參考，不判定。")
    R["E_note"] = r; r += 1
    for j, h in enumerate(["項目", "公司實績：2026 年已公布期間累計（主讀數）", "公司實績：最近一季 × 4（第二讀數）", "層級隱含：CY2026 全年", "判定（±15%）", "口徑、出處與證據等級",
                           "主讀數差距（實績 ÷（層級 × 月數÷12）− 1）", "第二讀數差距（實績 ÷ 層級 − 1）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["E_hdr"] = r; r += 1
    R["E_first"] = r
    r_in = r + 6
    r_oct, r_cum = r_in + 2, r_in + 3

    def row(label, b, c, d, src, nb="#,##0.0", months=None):
        nonlocal r
        C.cell(r, 1, label)
        for col, v in ((2, b), (3, c)):
            x = C.cell(r, col, v)
            if not isinstance(v, str) or v.startswith("="): x.number_format = nb
        x = C.cell(r, 4, d)
        if not isinstance(d, str) or d.startswith("="): x.number_format = nb
        C.cell(r, 7, f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r})),B{r}/(D{r}*{months}/12)-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 8, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}/D{r}-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 5, f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),IF(AND(ABS(G{r})>0.15,ABS(H{r})>0.15,SIGN(G{r})=SIGN(H{r})),"對帳中",IF(OR(ABS(G{r})>0.15,ABS(H{r})>0.15),"時差邊界（參考）","一致")),"無可比實績")')
        C.cell(r, 6, src); r += 1
    row("聯發科 AI ASIC（雲端／資料中心）營收（US$B）", "無實績（未揭露）", "無實績（未揭露）", f"=PL!D{r_xraw}",
        f"公司側：1Q26–3Q26 沒有 AI ASIC 營收金額實績。3Q26 法說（2026-10-30）只說首顆 AI 加速器 ASIC 9 月開始出貨初期訂單、未給金額（出處網址見本頁下方輸入表與 Sources 頁）；4Q26 雲端 ASIC 營收預估 25 億美元是指引、不是實績。"
        "層級側：PL 頁情境 A 區塊「AI ASIC 營收（未校準）」CY2026＝TPU 顆數 × 聯發科份額 × 每顆營收（Siliconomics v0.20 份額與顆數；每顆規格讀 Broadcom v0.5.7）。"
        "層級份額於 Siliconomics v0.3 取自本模型原值（同源，非獨立估計）。無可比實績：缺公司的 ASIC 營收揭露（最早 2027-01 月營收、2027-02 法說的 4Q26 實績）", "#,##0.00", 9)
    row("合併營收（NT$B；層級無聯發科合併營收）", f"=B{r_cum}", "=Q!E4*4", "無同口徑隱含值",
        "公司實績：1–10 月月營收累計 524.416 十億元（公司月營收公告，科技新報 2026-11-10；輸入區；Verified）；第二讀數＝3Q26 營收 168.057 × 4（Q 頁；7–9 月月營收合計，Verified）。"
        "層級側：Siliconomics v0.20 沒有聯發科（或手機 SoC、智慧邊緣平台）的營收或份額，只有 ASIC 份額與市場顆數。無可比實績：缺層級的聯發科合併營收", "#,##0.0", 10)
    R["E_last"] = r - 1
    C.cell(r, 1, "參考（不判定）：公司 FY2026 資料中心 ASIC 指引 vs 層級隱含 CY2026（US$B）")
    C.cell(r, 2, "=dc26").number_format = "0.00"; C.cell(r, 3, "—"); C.cell(r, 4, f"=PL!D{r_xraw}").number_format = "0.00"
    C.cell(r, 5, "參考（指引；層級份額同源，不判定）")
    C.cell(r, 6, "B 欄＝In 頁 dc26（公司 3Q26 法說指引：4Q26 雲端 ASIC 25 億美元＋1–3Q26 假設 3 億美元〔asic13，Assumed〕＝28 億美元；Interested-party，指引不是實績）。D 欄＝層級隱含（校準前）。"
                 "B ÷ D 見 G 欄：指引高於層級隱含，本模型以校準係數（Units 頁／PL 頁）吸收此差距。這是指引對層級，不是實績對層級，且層級份額取自本模型，不判定、不開對帳單；待 4Q26 實績揭露後可成為可判定的 C0 項目。")
    C.cell(r, 7, f'=B{r}/D{r}-1').number_format = "+0.0%;-0.0%"; C.cell(r, 8, "—"); R["E_ref"] = r; r += 1
    C.cell(r, 1, "參考（不判定）：公司 2027 年 ASIC 目標（SAM × 份額中點）vs 層級隱含 CY2027（US$B）")
    C.cell(r, 2, "=co_sam27*co_sh27").number_format = "0.0"; C.cell(r, 3, "—"); C.cell(r, 4, f"=PL!E{r_xraw}").number_format = "0.00"
    C.cell(r, 5, "參考（公司目標；不判定）")
    C.cell(r, 6, "B 欄＝公司 2Q26 法說：2027 年資料中心 ASIC 可服務市場 800 億美元 × 份額目標 15–20% 中點 17.5%（In 頁的 2027 年可服務市場與份額目標；Interested-party；3Q26 法說稱 160 億美元為原區間高標）。D 欄＝層級隱含 CY2027（校準前，PL 頁情境 A）。公司長期目標與外部估計只列參考。")
    C.cell(r, 7, f'=B{r}/D{r}-1').number_format = "+0.0%;-0.0%"; C.cell(r, 8, "—"); R["E_ref2"] = r; r += 1
    C.cell(r, 1, "判定摘要"); R["E_sum"] = r
    rg = f'E{R["E_first"]}:E{R["E_last"]}'
    C.cell(r, 2, F(T("C0（實績）共 "), f'COUNTA({rg})', T(" 項：可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})', T(" 項、對帳中 "), f'COUNTIF({rg},"對帳中")', T(" 項、時差邊界 "), f'COUNTIF({rg},"時差邊界（參考）")',
                   T(" 項、無可比實績 "), f'COUNTIF({rg},"無可比實績")', T(" 項。原因：聯發科尚未揭露 AI ASIC 營收金額（首顆 ASIC 4Q26 才量產）；層級沒有聯發科合併營收；層級份額取自本模型。不以模型預測、指引或同源值補位；待 4Q26 實績揭露後再判定。")))
    R["E_sum_rg"] = rg
    r += 2
    assert r == r_in, (r, r_in)
    C.cell(r, 1, "輸入：公司月營收與 3Q26 法說揭露（十億新台幣；藍字為輸入；Verified，除另註）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "出處"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "2026 年 10 月營收（NT$B）"); c = C.cell(r, 2, 55.025); c.font = BLUE; c.number_format = "#,##0.000"
    C.cell(r, 3, f"公司月營收公告（科技新報 2026-11-10 轉載：10 月營收 550.25 億元，月減 0.7%、年增 5.8%）：{SOCT}"); assert r == r_oct; r += 1
    C.cell(r, 1, "2026 年 1–10 月累計營收（NT$B）"); c = C.cell(r, 2, 524.416); c.font = BLUE; c.number_format = "#,##0.000"
    C.cell(r, 3, f"同上：前 10 個月營收 5,244.16 億元，年增 5.35%：{SOCT}"); assert r == r_cum; r += 1
    C.cell(r, 1, "檢查：1–10 月累計 − 10 月營收 − 1–3Q26 營收（Q 頁；應 ≈0，|差| < 0.01）"); R["E_chk1"] = r
    C.cell(r, 2, f"=ROUND(B{r_cum}-B{r_oct}-(Q!C4+Q!D4+Q!E4),3)").number_format = "0.000"; r += 1
    C.cell(r, 1, "4Q26 雲端 ASIC 營收指引（US$B；指引，不是實績）"); c = C.cell(r, 2, 2.5); c.font = BLUE
    C.cell(r, 3, f"公司 3Q26 法說（2026-10-30）：4Q26 雲端 ASIC 營收預估 25 億美元（{SQ3b}）；Interested-party（指引）"); R["E_guid"] = r; r += 1
    C.cell(r, 1, "檢查：In 頁 dc26（FY2026 資料中心）− 指引 − 1–3Q26 假設 asic13（應 ≈0）"); C.cell(r, 2, f"=ROUND(dc26-B{r - 1}-asic13,6)").number_format = "0.000"; R["E_chk2"] = r; r += 1
    R["E_end"] = r - 1
    C.cell(R["B_sum"], 2, F(T("本公司 C0 對帳中項目 "), f'COUNTIF({rg},"對帳中")', T(" 項：公司實績對層級隱含值的 C0（實績）表見 E 節（可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})',
                             T(" 項；其餘因公司或層級側沒有同口徑值，標無可比實績）。本節 TPU 顆數、聯發科份額、核心售價是個股讀入的層級值，與層級同源，標「讀取層級（同源，不判定）」，不計入。"
                               "下方「層級之間」的產業部署 GW 差距只列參考，不計入。對帳單由主控與 Siliconomics 層處理，本模型不改上游值。")))
    return R


# ---------- README：v1.3 白話列 ----------
def readme13(wb, RVA, PLA, TO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; C = wb["Coord"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed = vb("同業前瞻 P/E 中位數")
    geps, gmul = vb("與市場共識目標價的差距：EPS 因素"), vb("與市場共識目標價的差距：倍數因素")
    b0 = {s: lab(P, f"情境 {s}（", 60) for s in "ABC"}
    asu = {s: lab(P, "AI ASIC 營收（美元）", b0[s], exact=True) for s in "ABC"}
    epr = {s: lab(P, "EPS（基本）", b0[s], exact=True) for s in "ABC"}
    tor = lab(SE, "加權目標價相對基準（HTML")
    tr = lambda k, col: f"Sens!{col}{lab(SE, k, tor)}"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    r = RD.max_row + 2
    RD.cell(r, 1, "展示層 v1.3（HTML 用；v0.3.4 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，比現價 "), 'TEXT(px,"#,##0")', T(" 元"), f'IF({up}>=0," 高 "," 低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "),
        f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")', T(" 元，都"), f'IF(MAX({dcf},{cadj})<px,"低於","未全低於")', T("現價：市場已把 AI ASIC 的高成長價進股價，本模型 ASIC 受 TPU 顆數與份額上限限制。"))
    put("rec", T("投資建議「"), rec, T("」的前提：本模型 ASIC 營收只到 FY2028 "), f'TEXT(PL!F{asu["B"]},"0.0")', T(" 十億美元（情境 B），遠低於賣方的 50 十億美元。若份額、顆數或第二顆 ASIC 單價明顯上修，結論會改變（見敏感度）。"))
    put("o1", T("目標價 "), f'TEXT({tp},"#,##0")', T(" 元，比現價 "), 'TEXT(px,"#,##0")', T(" 元"), f'IF({up}>=0,"高 ","低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "), f'TEXT({dcf},"#,##0")',
        T(" 元與同業比較法 "), f'TEXT({cadj},"#,##0")', T(" 元各占一半；市場共識 "), 'TEXT(cons_tp,"#,##0")', T(" 元，差距來自前瞻 EPS（本模型"), f'IF({geps}>=0,"高 ","低 ")', f'TEXT(ABS({geps}),"0%")',
        T("）與倍數（共識隱含倍數"), f'IF({gmul}>=0,"高於","低於")', T("同業中位數 "), f'TEXT(ABS({gmul}),"0%")', T("）。"))
    put("o2", T("核心問題一（ASIC 能做多大）：FY2027 ASIC 營收 "), f'TEXT(PL!E{asu["B"]},"0.0")', T(" 十億美元，與公司目標約相當；FY2028 三情境 "), f'TEXT(PL!F{asu["A"]},"0.0")', T("／"), f'TEXT(PL!F{asu["B"]},"0.0")', T("／"), f'TEXT(PL!F{asu["C"]},"0.0")',
        T(" 十億美元，賣方 50。FY2028 EPS "), f'TEXT(PL!F{epr["C"]},"0")', T("～"), f'TEXT(PL!F{epr["A"]},"0")', T(" 元。"))
    put("o3", T("核心問題二（差距靠什麼補）：TPU 份額已近上限，主要看第二顆 ASIC 單價：單價 ×2 使加權目標價 "), f'TEXT({tr("第二顆 ASIC", "C")},"+0%;-0%")', T("，聯發科顆數 ×0.8／×1.2 使 "), f'TEXT({tr("聯發科 ASIC 顆數", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("聯發科 ASIC 顆數", "C")},"+0%;-0%")', T("。"))
    put("o4", T("反向推算（情境 B）：要支持現價，折現率須"), f'IF({rv("w1")}<=wacc,"降到","升到")', T(" "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2027 起營收全面"), f'IF({rv("s1")}>=1,"放大 ","縮小為 ")', f'TEXT({rv("s1")},"0.00")',
        T(" 倍；同業比較法：現價相當於前瞻 EPS 的 "), f'TEXT(RevDCF!B{RVA["pe1"]},"0.0")', T(" 倍，"), f'IF(RevDCF!B{RVA["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍。"))
    put("o5", T("對帳中（CY2027–29，層與層的差距）：另一側讀數目標價 "), f'TEXT(Coord!C{rtp},"#,##0")', T(" 元，評等"), f'Coord!C{rrec}', f'IF(Coord!B{rrec}<>Coord!C{rrec},"（與現行基準讀數"&Coord!B{rrec}&"評等不同）","")',
        T("。公司尚未揭露 ASIC 營收，公司實績對層級的 C0 對帳為無可比實績，待 4Q26 實績。"))
    put("drv", T("分部營收推導（情境 B）：合併營收＝AI ASIC＋手機＋智慧邊緣平台（不含 ASIC）＋電源 IC。AI ASIC＝TPU 市場顆數 × 聯發科份額 × 每顆營收（核心售價＋代購 HBM）× 校準係數（FY2026 對齊公司資料中心說法）× 匯率；其餘三塊由 FY2026 錨點依成長率推。"
                 "FY2028 合計營收 "), f'TEXT(PL!F{lab(P, "營收", b0["B"], exact=True)},"#,##0")', T(" 十億元，其中 AI ASIC "), f'TEXT(PL!F{lab(P, "AI ASIC 營收", b0["B"], exact=True)},"#,##0")', T(" 十億元。"))
    out["ver"] = r
    RD.cell(r, 1, F(T(f"v{VER}（{BUILT}）：資產負債表改公司 3Q26 合併財報（2026-09-30，會計師核閱），債務加計零息海外可轉債。更正：v0.3.4 以前債務只含短期借款 41.950，漏計 2026-09 發行的可轉債（帳面 111.338、面額 123.408，轉換價 4,513.75 元，2031-09 到期），淨現金高估約 111（十億新台幣）。"
                      "淨現金與投資 367.71 → "), 'TEXT(dcf_nc,"0.00")', T("（債務 41.95 → "), 'TEXT(debt_v,"0.000")', T("、租賃 9.785 → "), 'TEXT(lease_v,"0.000")', T("、現金與流動金融資產 277.687 → "), 'TEXT(cash_v,"0.000")',
                    T("、非流動投資 180.620 → "), 'TEXT(inv_v,"0.000")', T("）；D/E 0.0094 → "), 'TEXT(beta_de,"0.0000")', T("、β 1.44 → "), 'TEXT(beta,"0.00")', T("、WACC 10.23% → "), 'TEXT(wacc,"0.00%")',
                    T("。DCF 目標價 1,773.03 → "), f'TEXT({dcf},"#,##0.00")', T("，加權目標價 1,715.10 → "), f'TEXT({tp},"#,##0.00")', T("（上檔 −50.4% → "), f'TEXT({up},"+0.0%;-0.0%")', T("），評等 強力放空 → "), rec,
                    T("。可轉債基準以帳面計債、不計稀釋（現價為轉換價的 "), 'TEXT(cb_px,"0%")', T("）；Sens 新增以面額計債務、視為全數轉換兩案。")))
    out["ver34"] = r + 1; r += 1
    RD.cell(r, 1, f"v0.3.4（{BUILT}）：只改展示層與新增列——個股 HTML 版面規格 v1.3（章節：命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態）；三層協調（規則 v1.4.2）：Coord 頁新增 A 讀哪一層哪一版、B 讀取層級（同源，不判定）、C 口徑註記、D 摘要、E C0（實績）：公司實績 vs 層級隱含值，"
                  "並列兩讀數目標價（C4；主引擎情境 B、C 於 CY2027 起 ASIC 顆數 × 另一側讀數比值；情境 A 不縮放）。新增 Hist 頁（損益表歷史 FY2022–FY2025）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B，GAAP）、Val 0.7（四個目標價）、"
                  "Sens 頁末加權目標價相對基準（營運輸入案例為整本重算快照 mediatek/sens.json）、README 頁末 v1.3 白話列。既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異）。")
    r += 1
    return out


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT or "Coord" not in wb.sheetnames:
        return "v034_post：略過（另一側讀數建置或無 Coord 頁）"
    HR = hist_is(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_tor(wb, here)
    CR = coord_c0(wb, SIL)
    RMA = readme13(wb, RVA, PLA, ST)
    RD = wb["README"]
    RD["A2"].value = (f"版本：v{VER}（{BUILT}：資產負債表改公司 3Q26 合併財報、債務加計零息海外可轉債；前版 v0.3.4，同日）；"
                      f"v0.3.4（{BUILT}：版面規格 v1.3、三層協調 C0 與兩讀數；前版 v0.3.3，2026-11-06）；" + str(RD["A2"].value))
    a1 = str(RD["A1"].value); head = a1.split("；", 1)[0]
    assert "v0.3.4" in head, head
    RD["A1"].value = head.replace("v0.3.4", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：資產負債表改公司 3Q26 合併財報，債務加計零息海外可轉債（更正 v0.3.4 以前淨現金漏計），Sens 新增可轉債處理兩案"
    order = ["README", "In", "Q", "Hist", "Units", "PL", "Val", "DCF", "RevDCF", "Recon", "Sens", "Sources", "Coord"]
    wb._sheets = [wb[n] for n in order]
    return f"v034_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 相對基準 {ST['first']}–{ST['last']}；Coord E 節 {CR['E_first']}–{CR['E_last']}；README v1.3 {len(RMA)} 列"
