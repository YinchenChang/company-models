# 日月光投控 v0.2.6（2026-11-14）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調新機制（三層基準協調規則 v1.4.1 C0、C4）。
# 樣板 NVIDIA v0.3.7／v0.3.8（nvidia/v037_post.py、coord_post.py）。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史 FY2022–FY2024（公司 6-K 季報新聞稿；FY2025 取 Q 頁）
#   PL 頁末：標準損益表（FY2023–FY2025 實績＋FY2026E–FY2031E 情境 B；GAAP＝IFRS 合併口徑）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列（四個目標價）
#   Sens 頁末：加權目標價相對基準（低案、高案；HTML 龍捲風圖用）
#   Coord 頁末：三層協調（A 讀哪一層哪一版、B C0 由下往上對帳、C 口徑註記、D 摘要與損益表年度標示）；co_flag_<yyyy> 具名範圍
#   README 頁末：v1.3 展示層白話列與版本說明；標題段版號更正
# 建檔程式在 Coord 之後、存檔前呼叫 apply(wb, here, SIL)。另一側讀數建置（COORD_RUN＝1）時不執行（alt_run.py 只讀 Val 頁）。
import os
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.6", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
U4Q25 = "https://www.sec.gov/Archives/edgar/data/1122411/000095010326001662/dp241167_6k.htm"
U4Q24 = "https://www.sec.gov/Archives/edgar/data/1122411/000095010325001937/dp224882_6k.htm"
U4Q23 = "https://capedge.com/filing/1122411/0000950103-24-001548/ASX-6K"   # SEC 6-K accession 0000950103-24-001548（2024-02-01）；capedge 為 SEC 文件鏡像
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


# ---------- Hist：損益表歷史（NT$B；公司 6-K 季報新聞稿合併損益表，IFRS） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025 公式)；None＝該年本模型不需要（不填）
HIS = [
    ("rev", "營收", 670.873, 581.914, 595.410, "=Q!B4"),
    ("cogs", "營業成本", None, 490.157, 498.478, "=Q!B4-Q!B14"),
    ("gp", "毛利", 134.930, 91.757, 96.932, "=Q!B14"),
    ("opx", "營業費用（研發＋銷管）", None, 51.429, 57.765, "=Q!B16"),
    ("oi", "營業利益", 80.176, 40.328, 39.167, "=Q!B17"),
    ("nonop", "營業外淨額", None, 2.272, 2.517, "=Q!B18"),
    ("pbt", "稅前利益", 81.626, 42.600, 41.684, "=Q!B19"),
    ("tax", "所得稅", None, 9.043, 7.758, "=Q!B20"),
    ("nic", "稅後淨利（合併，含非控制權益）", None, 33.557, 33.926, "=Q!B19-Q!B20"),
    ("nci", "非控制權益損益", None, 1.832, 1.443, "=Q!B21"),
    ("ni", "歸屬母公司淨利", 62.090, 31.725, 32.483, "=Q!B22"),
    ("eps", "EPS（基本，公司公告；NT$）", 14.53, 7.39, 7.52, "=Q!B25"),
    ("epsd", "EPS（稀釋，公司公告；NT$）", 13.94, 7.18, 7.23, 8.89),
]


def hist_is(wb):
    Hs = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    Hs["A1"] = ("Hist — 損益表歷史（v0.2.6 起；十億新台幣；EPS 新台幣；IFRS 合併損益表、歸屬母公司）。PL 頁末「標準損益表」的歷史三年（FY2023–FY2025）與年增率基期用。"
                "FY2022 只取營收、毛利、營業利益、稅前利益、歸屬母公司淨利與 EPS（年增率基期）；FY2025 欄＝Q 頁（4Q25 新聞稿）。只被 PL 頁標準損益表引用，不回饋任何計算")
    Hs["A1"].font = BOLD
    Hs["A2"] = (f"來源（公司原文＝SEC 6-K 季報新聞稿，Verified）：FY2023、FY2022＝4Q23 新聞稿（2024-02-01，accession 0000950103-24-001548；{U4Q23}）；"
                f"FY2023、FY2024＝4Q24 新聞稿（2025-02-13；{U4Q24}）；FY2024、FY2025＝4Q25 新聞稿（2026-02-05；{U4Q25}）。新聞稿為未經查核數字，與年報有微小差異時以年報為準（未發現差異）")
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025（Q 頁）", "標記", "來源"], 1):
        c = Hs.cell(4, j, h); c.font = BOLD; c.fill = HDR
    rows = {}
    r = 5
    for k, labl, v22, v23, v24, v25 in HIS:
        rows[k] = r
        Hs.cell(r, 1, f"h_{k}"); Hs.cell(r, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if v is None: continue
            c = Hs[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k in ("eps", "epsd") else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"h_{k}_{yy}", f"Hist!${col}${r}")
        Hs.cell(r, 7, "Verified")
        Hs.cell(r, 8, "4Q23／4Q24／4Q25 新聞稿合併損益表（FY2023 營業成本、營業費用等取 4Q24 新聞稿的比較欄）" if k != "epsd" else "4Q23／4Q24／4Q25 新聞稿；FY2025 稀釋 EPS 8.89（Q 頁說明列，4Q25 新聞稿）")
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
    # Sources 頁補兩列（只新增）
    S = wb["Sources"]; n = S.max_row + 1
    S.cell(n, 1, "公司：4Q24 新聞稿（SEC 6-K，2025-02-13；FY2024、FY2023 損益表，Hist 頁）"); S.cell(n, 2, U4Q24)
    S.cell(n + 1, 1, "公司：4Q23 新聞稿（SEC 6-K，2024-02-01，accession 0000950103-24-001548；FY2023、FY2022 營收、毛利、營業利益、淨利、EPS，Hist 頁；capedge 為 SEC 鏡像）"); S.cell(n + 1, 2, U4Q23)
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"       # 標準表欄；E 起（FY2025A…FY2031E）對應 PL 情境 B 區塊 C…I 欄
SRC = "CDEFGHI"


def pl_standard(wb):
    P = wb["PL"]; D = wb["DCF"]
    rbd = lab(P, "情境 B：")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    R = dict(rev=ex("營收"), gp=ex("毛利"), opx=ex("營業費用"), oi=ex("營業利益"), oth=ex("營業外淨額"), pbt=ex("稅前利益"), tax=ex("所得稅"), nci=ex("非控制權益"),
             ni=ex("歸屬母公司淨利"), sh=ex("基本股數"), eps=ex("EPS（基本）"))
    dbr = lab(D, "情境 B"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.2.6 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026E–FY2031E，情境 B 基準；FY2026E 為錨點）。NT$B；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（IFRS 合併財報、歸屬母公司）。日月光在台灣以 IFRS 編製合併財報，公司財測、市場共識與本模型 EPS 皆以此為準，沒有另設 Non-GAAP 口徑，故本表即 GAAP 口徑。"
                 "EPS 用基本 EPS（本模型口徑；公司另揭露稀釋 EPS，因可轉債稀釋，見下方 GAAP 對照列；預測期未模擬稀釋）。預測期的「其他收入（淨額）」＝營業外淨額")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（1Q–3Q26 實績＋4Q26 依公司指引）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E、FY2031E＝延伸期（收斂法只推營收與歸屬母公司淨利）：毛利率與營業費用率沿用 FY2029E 推算營業成本、毛利與營業費用；"
                 "營業利益取 DCF 頁情境 B（由淨利反推）；其他收入＝營業外淨額假設；所得稅＝稅前利益 −（歸屬母公司淨利＋非控制權益）。兩種推法的營業利益差額列在表末核對列，不改既有淨利與目標價")
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
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in SRC[:7]]   # FY2025A–FY2031E 的 PL 情境 B 區塊（只在有值的欄使用）
    hh = lambda k: [f"=h_{k}_23", f"=h_{k}_24"]
    # 營收
    row("rev", "營收", "NT$B", hh("rev") + [f"=PL!{c}{R['rev']}" for c in SRC], M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L), "=D{rev}/C{rev}-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS[1:], COLS[2:])], PCT)
    r_c = r   # 下一列＝cogs；J、K（FY2030E、FY2031E）用 FY2029E（I 欄）比率
    row("cogs", "營業成本（COGS）", "NT$B", hh("cogs") + [f"=PL!{c}{R['rev']}-PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-I{r_c + 2})"], M, "FY2030E、FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", hh("gp") + [f"=PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "NT$B", hh("opx") + [f"=PL!{c}{R['opx']}" for c in SRC[:5]] + [f"=J{L['rev']}*I{r_c + 3}/I{L['rev']}", f"=K{L['rev']}*I{r_c + 3}/I{L['rev']}"], M, "FY2030E、FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + [f"=PL!{c}{R['oi']}" for c in SRC[:5]] + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外淨額）", "NT$B", hh("nonop") + [f"=PL!{c}{R['oth']}" for c in SRC[:5]] + ["=nonop_f", "=nonop_f"], M, "利息、投資評價、匯兌等；FY2030E 起＝營業外淨額假設")
    row("pbt", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "NT$B", hh("tax") + [f"=PL!{c}{R['tax']}" for c in SRC[:5]] + [f"=J{L['pbt']}-(PL!H{R['ni']}+nci)", f"=K{L['pbt']}-(PL!I{R['ni']}+nci)"], M, "FY2030E、FY2031E＝稅前利益 −（歸屬母公司淨利＋非控制權益），等於稅率 × 稅前利益")
    row("nic", "稅後淨利（合併，含非控制權益）", "NT$B", [f"={c}{L['pbt']}-{c}{L['tax']}" for c in COLS], M)
    row("nci", "非控制權益損益", "NT$B", hh("nci") + [f"=PL!{c}{R['nci']}" for c in SRC[:5]] + ["=nci", "=nci"], M)
    row("ni", "歸屬母公司淨利（EPS 與目標價口徑）", "NT$B", [f"={c}{L['nic']}-{c}{L['nci']}" for c in COLS], M)
    row("nm", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "基本股數", "百萬股", [f"=C{L['ni']}/C{r + 1}*1000", f"=D{L['ni']}/D{r + 1}*1000"] + [f"=PL!{c}{R['sh']}" for c in SRC], "#,##0.0", "FY2023、FY2024＝歸屬母公司淨利 ÷ 公司公告基本 EPS（推算）")
    row("eps", "EPS（基本；NT$）", "NT$", ["=h_eps_23", "=h_eps_24"] + [f"=PL!{c}{R['eps']}" for c in SRC], "0.00", "FY2023–FY2025＝公司公告基本 EPS；FY2026E 起＝模型（歸屬母公司淨利 ÷ 基本股數）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY；歸屬母公司淨利）", "%", ["=C{ni}/h_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 歸屬母公司淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；公司公告稀釋 EPS，預測期未模擬稀釋）").font = BOLD; r_gaap = r; r += 1
    row("g_epsd", "EPS（稀釋，公司公告；NT$）", "NT$", ["=h_epsd_23", "=h_epsd_24", "=h_epsd_25"], "0.00", "稀釋因可轉換公司債；FY2025 8.89")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E、FY2031E 見下兩列）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：歸屬母公司淨利 − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$B", [None, None] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    row("c4", "核對：EPS − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$", [None, None] + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    row("c5", "核對：稅後淨利（合併，Hist）− 本表（FY2023–FY2025）", "NT$B", [f"=Hist!D{lab(wb['Hist'], 'h_nic', col=1)}-C{L['nic']}", f"=Hist!E{lab(wb['Hist'], 'h_nic', col=1)}-D{L['nic']}", f"=Hist!F{lab(wb['Hist'], 'h_nic', col=1)}-E{L['nic']}"], "0.000")
    # 分部營收推導補充（情境 B；供 HTML 驅動因子章）：已在 PL 情境 B 區塊，不重複
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
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.6 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In、PL 頁，無任何頁引用本頁，不影響目標價。金額十億新台幣；價格新台幣")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。FY2027–FY2032 全部為預測年（日月光 FY2026 為錨點，不動）")
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
    kv("nc", "淨現金滾到評價日（十億新台幣）", "=dcf_nc+stub_fcff4q-stub_div", "#,##0.0", "同 DCF 頁每股價值公式（淨現金與投資＋stub 期 FCFF − stub 期股利）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
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
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027–FY2032 皆為預測年）", lambda dc, c: f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（十億新台幣）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（十億新台幣）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（十億新台幣，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2027–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（十億新台幣）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（十億新台幣）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2027–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(B{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2027–FY2032 營收等比例放大；FY2026 錨點不動；營業利益率、稅率、折舊、資本支出比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（十億新台幣）", f"=fc_wc*PL!$D${rev26}*B{Y['fr']}*B{Y['df']}", "#,##0.0", "FY2027 營運資金增加中屬 FY2026 營收的部分（FY2026 為錨點，不放大）")
    kv("b", "隨營收倍數等比例變動的 EV（十億新台幣）",
       f"=SUMPRODUCT(B{Y['cf']}:G{Y['cf']},B{Y['fr']}:G{Y['fr']},B{Y['df']}:G{Y['df']})-fc_wc*B{Y['rev']}*B{Y['fr']}*B{Y['df']}"
       f"-SUMPRODUCT(C{Y['wc']}:G{Y['wc']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2027 營收（十億新台幣）", f"=B{A['s1']}*B{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2027 營收（情境 B；十億新台幣）", f"=B{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2027 營收年增率（相對 FY2026 錨點營收）", f"=B{A['rv1']}/PL!$D${rev26}-1", "0.0%", "FY2026 為錨點（1Q–3Q26 實績＋4Q26 指引），不放大")
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
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − 未取下限的 DCF 情境 B 評價日每股價值（應為 0；情境 B 為正值，MAX(0,…) 不啟動）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
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
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 元；要支持現價 "), 'TEXT(px,"0.00")', T(" 元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.0")',
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2027–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.0")', T(" 倍（FY2027 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億元，相對 FY2026 錨點年增 "), f'TEXT(B{A["g1"]},"0%")',
                     T("）；或營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，高於同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 日月光過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給日月光的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，比模型採用的相對倍數高 "), f'TEXT(D{A["rc1"]},"0%")', T("。")))
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
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.2.6 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權；情境 C 價值為 0（股權價值取下限 0）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（BofA 單一券商，1 家）", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "富邦證券 e 點通／Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列） ----------
# 鍵 → (白話標籤, 低案列首字)；案例成對（低、高）排在同一鍵的連續兩列
TOR_LABEL = {
    "mf": "AI 部署顆數 ×0.8／×1.2（FY2027 起，三情境）", "el": "LEAP 每顆內容隨封裝面積的彈性 0.3／1.0", "fp_g": "全流程滲透成長 10%／50%",
    "cpo27": "CPO 2027 年金額 US$0.3B／0.8B", "g_ms27": "主流封測 FY2027 成長 +5%／+22%", "g_msL": "主流封測 FY2028–29 成長 +2%／+12%",
    "g_ems27": "EMS FY2027 成長 0%／+15%", "g_emsL": "EMS FY2028–29 成長 0%／+8%", "ems_gm": "EMS 毛利率 7.5%／9.5%", "fx": "匯率 30.5／33",
    "capex27": "FY2027 資本支出 US$11.5B／15B", "capi": "FY2028–29 資本支出占營收 22%／38%", "life": "折舊年限 6 年／12 年", "ft": "現金流通率 40%／60%",
    "phi": "營業費用彈性 0.3／1.0", "nonop_f": "營業外淨額每年 −8／−2（十億元）", "tax": "有效稅率 15%／20%", "wacc_adj": "WACC −1／＋1 個百分點", "beta": "β 低／高（β 1.19／1.87）",
    "erp": "股權風險溢酬 4%／6%", "g_tv": "永續成長率 1%／2%", "ronic_sp": "永續期新投資報酬率高於 WACC 5 個百分點（只有高案）",
    "fc_cx_d": "每年資本支出占營收 −1／＋1 個百分點", "fc_wc": "營運資金占營收 10%／25%", "inv_k": "淨現金不含長期投資（只有低案）", "rf_us": "DCF 改用美元口徑折現（只有低案）",
    "tp_w_dcf": "DCF 權重 0.3／0.7",
}
ALONE = {"ronic_sp": "高", "inv_k": "低", "rf_us": "低"}


def sens_tor(wb):
    SE = wb["Sens"]
    items = []   # (label, low_row, high_row)
    r = 5
    while r <= 56:
        k = str(SE.cell(r, 2).value)
        a = str(SE.cell(r, 1).value or "")
        if k == "—":
            if a.startswith("終值改用 FCFF"):
                items.append(("DCF 終值改用 FCFF ×（1＋g）法（只有低案）", r, None))
            r += 1; continue
        if k in ALONE:
            items.append((TOR_LABEL[k], r if ALONE[k] == "低" else None, r if ALONE[k] == "高" else None)); r += 1; continue
        assert str(SE.cell(r + 1, 2).value) == k, (r, k)
        items.append((TOR_LABEL[k], r, r + 1)); r += 2
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.2.6 起，版面規格 v1.3 H 節；只引用上方各案例列的「加權目標價相對基準」欄，白話標籤，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源列"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    import re
    for labl, lo, hi in items:
        assert not re.search(r"(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+", labl), labl
        SE.cell(r, 1, labl)
        for col, sr in ((2, lo), (3, hi)):
            if sr: SE.cell(r, col, f"=Sens!J{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{lo or hi}" + (f"、A{hi}" if lo and hi else "")); r += 1
    return r0


# ---------- Coord：三層協調（C0 由下往上對帳等） ----------
def coord_c0(wb, SIL):
    C = wb["Coord"]; I = wb["In"]; P = wb["PL"]
    rbd = lab(P, "情境 B：")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    r_u, r_core = ex("AI 加速器顆數（含情境係數）"), ex("LEAP 核心（2.5D oS、先進測試等）")
    r_fp, r_cpo = ex("全流程（CoWoS full process）"), ex("CPO（共同封裝光學）")
    ikey = {str(I.cell(i, 1).value): i for i in range(4, I.max_row + 1) if I.cell(i, 1).value}
    C.column_dimensions["F"].width = 60
    for col in "BCDE": C.column_dimensions[col].width = 24
    r_ratio = lab(C, "另一側讀數 ÷ 現行基準")
    r = C.max_row + 3
    C.cell(r, 1, "三層協調（v0.2.6 新增；三層基準協調規則 v1.4.1 C0 個股與層的一致性、C4 個股不裁決、C5 往上回饋只有證據）").font = BOLD; r += 1
    R = {}
    # ---- A. 讀哪一層、哪一版 ----
    C.cell(r, 1, "A. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; R["A"] = r; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "讀自（層與版本）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["A_hdr"] = r; r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"

    def arow(label, cells, src, fmt="0.000"):
        nonlocal r
        C.cell(r, 1, label)
        for j, v in enumerate(cells):
            if v is None: continue
            c = C.cell(r, 2 + j, v); c.number_format = fmt
        C.cell(r, 6, src); R[label] = r; r += 1
    arow("AI 加速器顆數（情境 A 原始顆數；百萬顆）", [f"=In!E{ikey[f'units_{y}']}" for y in (26, 27, 28, 29)], f"{ver}：SIF_UnitsPlat_Total（In 頁 units_26–units_29，唯讀）", "0.000")
    arow("情境 B 顆數係數（÷ 原始顆數）", [None, 1, f"=In!E{ikey['um28_B']}", f"=In!E{ikey['um29_B']}"], f"{ver}：SIF_UnitsCoef_Base（規則 v1.4；In 頁 um28_B、um29_B；CY2026–27 為 1）")
    arow("封裝面積指數（2026＝1）", [f"=In!E{ikey[f'ar_{y}']}" for y in (26, 27, 28, 29)], f"{ver}：CoWoS 每片封裝數 cw_pp_*（In 頁 ar_26–ar_29）")
    arow("CPO 占 scale-out 光學金額（情境 B）", [f"=In!E{ikey[f'cpo_B_{y}']}" for y in (26, 27, 28, 29)], f"{ver}：SIF_cpo_B_*（In 頁 cpo_B_26–cpo_B_29）", "0.0%")
    arow("匯率（NT$／US$；CY2027 起換算 LEAP 與資本支出）", [None, f"=In!E{ikey['fx']}", f"=In!E{ikey['fx']}", f"=In!E{ikey['fx']}"], f"{ver}：fx_twd（In 頁 fx）", "0.00")
    arow("Wattonomics 錨定係數（只進 Sens 案例，不改基準）", [None] + [f"=In!E{ikey[f'wta_{y}']}" for y in (27, 28, 29)], f"{ver}：SIF_WTAnchor_Base（R4 比值＝Wattonomics ÷ 本庫推得上線 GW）；Wattonomics 現行版本 v0.3.2")
    arow("另一側讀數 ÷ 現行基準（對帳中年份；兩讀數目標價用）", [f"={'BCDE'[i]}{r_ratio}" for i in range(4)], f"{ver}：SIF_DeployGW_AltRatio_<年>（Wattonomics 換算的產業 GW ÷ Siliconomics 基準）；上方 Coord 第 6 列")
    C.cell(r, 1, "Wattonomics：本模型不直接讀 Wattonomics 數字；現行版本 v0.3.2，僅經 Siliconomics 的 R4 比值（Sens 錨定案例）與另一側讀數比值（對帳中年份兩讀數）間接使用。其餘產業量皆讀 Siliconomics 基準。"); R["A_note"] = r; r += 2
    # ---- B. C0 由下往上對帳 ----
    C.cell(r, 1, "B. C0 由下往上對帳：日月光自己的量 vs 它讀的 Siliconomics 值（差距 > ±15% 標「對帳中」；不改上游值，由主控開對帳單；規則 v1.4.1 C0 第 2 點）").font = BOLD; R["B"] = r; r += 1
    for j, h in enumerate(["項目（年）", "本模型原值", "層級對應值", "差距（本模型 ÷ 層級 − 1）", "判定（±15%）", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["B_hdr"] = r; r += 1
    units_L = SIL.NAMES["SIF_UnitsPlat_Total_Base"]; gw_b = SIL.NAMES["SIF_DeployGW_Base"]
    R["B_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        col = "DEFG"[i]
        C.cell(r, 1, f"AI 加速器顆數（情境 B）CY{y}（百萬顆）")
        C.cell(r, 2, f"=PL!{col}{r_u}").number_format = "0.000"
        c = C.cell(r, 3, units_L[i + 1]); c.number_format = "0.000"; c.font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"對帳中","一致")')
        C.cell(r, 6, f"本模型：PL 頁情境 B 顆數（原始顆數 × 基準係數）；層：{ver} SIF_UnitsPlat_Total_Base（供給側基準顆數）。同源讀取故差距 0；日月光不公布加速器顆數，兩邊都沒有證據（等級：Assumed／Interested-party），依 C0 第 3 點層的值先成立"
                     if i == 0 else "同上")
        r += 1
    R["B_last"] = r - 1
    C.cell(r, 1, "日月光 LEAP 營收（FY2026–FY2029）vs 層級量"); C.cell(r, 2, "—"); C.cell(r, 3, "無可比的層級量"); C.cell(r, 4, "—"); C.cell(r, 5, "不判定")
    C.cell(r, 6, "Siliconomics 沒有日月光單一公司的營收、份額或出貨量：只有「委外封測（日月光、艾克爾）補充 CoWoS 產能 ÷ 台積電 ＝ 10%」（SIF cw_osat，Assumed，5%–20%），以 CoWoS 月產能（千片）表達，"
                 "缺每片價格與日月光、艾克爾的分配，無法換算成日月光 LEAP 美元營收；日月光不在 Siliconomics 份額表（加速器、ABF、液冷等）內，也沒有「個股加總＋殘差＝市場值」可對（C0 第 1 點）。下列橋接列只作參考，不判定"); R["B_nocmp"] = r; r += 1
    cap = [SIL.base(f"cw_capend{y}") for y in (26, 27, 28, 29)]; osat = SIL.base("cw_osat")
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        col = "DEFG"[i]
        C.cell(r, 1, f"參考：LEAP 核心營收 ÷ 層級委外補充 CoWoS 月產能 CY{y}（US$B／千片月；不判定）")
        C.cell(r, 2, f"=PL!{col}{r_core}/C{r}").number_format = "0.000"
        c = C.cell(r, 3, cap[i] * osat); c.number_format = "0.0"; c.font = BLUE
        C.cell(r, 4, "—"); C.cell(r, 5, "參考")
        C.cell(r, 6, f"分子：PL 情境 B 的 LEAP 核心（不含全流程、CPO；US$B）；分母：{ver} 該年年底 CoWoS 月產能 {cap[i]:g} 千片 × 委外補充比 {osat:g}（日月光＋艾克爾合計，Assumed）。比值下降＝日月光 LEAP 營收成長慢於層級委外產能擴張；口徑不同，不構成對帳"
                     if i == 0 else "同上")
        r += 1
    R["B_end"] = r - 1
    C.cell(r, 1, "判定摘要"); R["B_sum"] = r
    nB = f'COUNTIF(E{R["B_first"]}:E{R["B_last"]},"對帳中")'
    C.cell(r, 2, f'="本公司 C0 可判定的 4 項（日月光自己的量對它讀的 Siliconomics 值）中，"&{nB}&" 項差距超過 ±15%：加速器顆數同源讀取，一致；LEAP 營收無可比層級量，不判定。下方「層級之間」的產業部署 GW 差距只列參考，不計入。對帳單由主控與 Siliconomics 層處理，本模型不改上游值。"'); r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; R["B2"] = r; r += 1
    for j, h in enumerate(["項目（年）", "Siliconomics 基準", "外部估計（Wattonomics 換算）", "差距（基準 ÷ 外部 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    R["B2_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        C.cell(r, 1, f"產業部署 GW CY{y}（本模型讀 Siliconomics 基準 ÷ Wattonomics 換算；GW）")
        c = C.cell(r, 2, gw_b[i]); c.number_format = "0.00"; c.font = BLUE
        c = C.cell(r, 3, SIL.NAMES[f"SIF_DeployGW_Alt_{y}"]); c.number_format = "0.00"; c.font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "參考")
        C.cell(r, 6, (f"本模型讀值＝{ver} SIF_DeployGW_Base（供給側部署 GW；本模型顆數、每 GW 的 LEAP 營收皆依此）；外部估計（參考）＝Wattonomics 換算 SIF_DeployGW_Alt_<年>（需求側）。"
                      "這是層與層之間的差距，不是日月光自己的量對層級的比較；由 Siliconomics 對帳單 r3 處理，不計入本公司 C0 項目數；本模型依 C4 並列兩讀數目標價，不裁決") if i == 0 else "同上")
        r += 1
    R["B2_last"] = r - 1
    r += 1
    # ---- C. 口徑註記 ----
    C.cell(r, 1, "C. 對帳中年份處理與口徑註記（C4）").font = BOLD; R["C"] = r; r += 1
    C.cell(r, 1, F(T("對帳中年份：CY2027、CY2028、CY2029（Siliconomics 標示；本公司 C0 自身項目無超限）。兩讀數目標價並列於上方（現行基準讀數 "), f'TEXT(B{lab(C, "加權目標價")},"0.00")', T(" 元、另一側讀數 "), f'TEXT(C{lab(C, "加權目標價")},"0.00")',
                    T(" 元），評等同為"), f'B{lab(C, "投資建議（五級")}', T("。個股不裁決（C4）；Siliconomics 與 Wattonomics 的差距由對帳單處理（期限 2027-02-28）。")))
    R["C_1"] = r; r += 1
    C.cell(r, 1, F(T("口徑註記（不改模型）：全流程（CoWoS full process，FY2027 錨點）與 CPO 2027 是固定錨點（公司指引：全流程逾 US$1B、CPO 逾 US$0.5B），不隨另一側讀數縮放；另一側讀數下 CY2027 顆數 ×"),
                    f'TEXT(C{r_ratio},"0.00")', T("，但這兩項 FY2027 金額不變，FY2028–29 再依顆數比（另一側 CY2028 ÷ CY2027）成長，使兩項在 FY2028–29 反而較現行基準讀數高約 "),
                    f'TEXT(D{r_ratio}/C{r_ratio}-1,"0.0%")', T("（"), f'TEXT(D{r_ratio},"0.000")', T("／"), f'TEXT(C{r_ratio},"0.000")',
                    T("−1）。這是固定錨點與比例縮放並用的口徑效果，已含在另一側目標價內；對帳單處理時一併說明。")))
    R["C_2"] = r; r += 1
    C.cell(r, 1, "C0 結論：本模型的產業量（顆數、封裝面積、CPO 滲透率、匯率）全部讀 Siliconomics 基準，與層同源，沒有獨立證據可往上回饋；公司揭露（LEAP FY2025 US$1.6B、FY2026 逾 US$4B）是日月光自己的實績與指引，"
                 "Siliconomics 沒有對應的日月光量，故無可比層級量。往上回饋只有證據（C5）：待 4Q26 財報與 2027 法說的 LEAP、資本支出實績再評估是否送對帳單。"); R["C_3"] = r; r += 2
    # ---- D. 摘要（HTML、檢查用） ----
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
    for j, (t) in enumerate(["FY2027E（對帳中）", "FY2028E（對帳中）", "FY2029E（對帳中）", "FY2030E（對帳中；延伸期）", "FY2031E（對帳中；延伸期）"]):
        C.cell(r, 2 + j, t)
    r += 1
    # co_flag_<yyyy>：個股本地鏡像（版面檢查 I 項讀此；數值＝Siliconomics SIF_RecoFlag_<年>，上方 Coord 第 5 列）
    r_flag = lab(C, "對帳中旗標")
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        dn(wb, f"co_flag_{y}", f"Coord!${'BCDE'[i]}${r_flag}")
    return R


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, coef, cgap = vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數"), vb("與市場共識目標價的差距：EPS 因素")
    rb = lab(P, "情境 B：")
    ex = lambda t, s=rb: lab(P, t, s, exact=True)
    def block(sc):   # 情境 A、B、C 的詳細區塊起始列
        return lab(P, {"A": "情境 A：", "B": "情境 B：", "C": "情境 C："}[sc])
    shr = {s: lab(P, "LEAP ÷ 封測營收", block(s)) for s in "ABC"}
    epr = {s: lab(P, "EPS（基本）", block(s), exact=True) for s in "ABC"}
    roi = {s: lab(P, "營業利益增量（相對 FY2025）", block(s)) for s in "ABC"}
    fcf = {s: lab(P, "股東自由現金流", block(s)) for s in "ABC"}
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    zero = f"({cadj}/px-1)"
    tor = lab(SE, "加權目標價相對基準（HTML")
    tr = lambda k, col: f"Sens!{col}{lab(SE, k, tor)}"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.6 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 元，比現價 "), 'TEXT(px,"0")', T(" 元低 "), f'TEXT(-{up},"0.0%")', T("，評等"), rec, T("。現金流折現只有 "),
        f'TEXT({dcf},"0")', T(" 元（資本支出高峰壓低自由現金流），同業比較法 "), f'TEXT({cadj},"0")', T(" 元。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：資本支出約占營收三到四成，折舊快速上升，營業利益增量撐不起資本支出報酬；LEAP 升到封測營收三成以上後，同業比較法的倍數不再擴張。"
                 "若只看同業比較法（不採現金流折現），加權目標價 "), f'TEXT({cadj},"0")', T(" 元（"), f'TEXT({zero},"+0.0%;-0.0%")', T("），評等"), f'{REC5(zero)}', T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 元，比現價 "), 'TEXT(px,"0.00")', T(" 元低 "), f'TEXT(-{up},"0.0%")', T("，評等"), rec,
        T("。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 元與同業比較法 "), f'TEXT({cadj},"0.00")', T(" 元各占一半；市場共識（BofA 單一券商）"), 'TEXT(cons_tp,"0")', T(" 元，差距主要在盈餘（模型前瞻 EPS 低 "),
        f'TEXT(-{cgap},"0%")', T("）。"))
    put("o2", T("核心問題一：LEAP 占封測營收 FY2026 "), f'TEXT(PL!D{shr["B"]},"0.0%")', T(" → FY2027 "), f'TEXT(PL!E{shr["B"]},"0.0%")', T("（情境 B），FY2029 三情境 "),
        f'TEXT(PL!G{shr["C"]},"0%")', T("～"), f'TEXT(PL!G{shr["A"]},"0%")', T("，分歧來自 AI 加速器顆數。FY2028 EPS 三情境 "), f'TEXT(PL!F{epr["C"]},"0.0")', T("～"), f'TEXT(PL!F{epr["A"]},"0.0")', T(" 元。"))
    put("o3", T("核心問題二：資本支出報酬偏低。營業利益增量 ÷ 累計資本支出，FY2028 情境 B 只有 "), f'TEXT(PL!F{roi["B"]},"0.0%")', T("，低於 WACC "), 'TEXT(wacc,"0.0%")',
        T("；股東自由現金流 FY2029 仍為負，約 −"), f'TEXT(-PL!G{fcf["B"]},"0")', T(" 十億元。標準版 DCF 直接扣這段資本支出，DCF 目標價因此只有 "), f'TEXT({dcf},"0")', T(" 元。"))
    put("o4", T("反向推算（情境 B）：要支持現價，折現率須降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2027 起營收全面放大 "), f'TEXT({rv("s1")},"0.0")',
        T(" 倍。同業比較法：現價相當於前瞻 EPS 的 "), f'TEXT(RevDCF!B{RVA["pe1"]},"0.0")', T(" 倍，高於同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍。"))
    put("o5", T("對帳中年份（CY2027–29，Siliconomics 標示；層與層的差距由對帳單處理，非本公司 C0 項目）：另一側讀數目標價 "), f'TEXT(Coord!C{lab(wb["Coord"], "加權目標價（DCF 與 Comparable 加權）")},"0.00")',
        T(" 元，評等不變。最敏感：AI 部署 ×0.8／×1.2 使加權目標價 "), f'TEXT({tr("AI 部署顆數", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("AI 部署顆數", "C")},"+0%;-0%")',
        T("，現金流通率 40%／60% 使 "), f'TEXT({tr("現金流通率", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("現金流通率", "C")},"+0%;-0%")', T("。"))
    put("drv", T("分部營收推導（情境 B）：合併營收＝封測（ATM）＋EMS（環旭）＋沖銷。封測＝主流封測（依成長率）＋LEAP 先進封裝；LEAP＝AI 加速器顆數 × 封裝面積指數 × 委外與份額相對值 × 每顆核心內容（美元），"
                 "再加全流程與 CPO，乘匯率換新台幣；EMS 依成長率。FY2028 LEAP "), f'TEXT(PL!F{lab(P, "LEAP 先進封裝（命題子列）", rb)},"0")', T("、主流封測 "), f'TEXT(PL!F{lab(P, "主流封測", rb)},"0")',
        T("、EMS "), f'TEXT(PL!F{lab(P, "EMS（環旭）", rb)},"0")', T("、沖銷 "), f'TEXT(PL!F{lab(P, "沖銷及其他", rb)},"0")', T("，合計營收 "), f'TEXT(PL!F{ex("營收")},"0")', T(" 十億新台幣。"))
    RD.cell(r, 1, f"v0.2.5（2026-11-14）：三層協調——Coord 頁並列兩讀數目標價（C4）；主引擎情境 B、C 於 CY2027 起產業量 × 另一側讀數比值（0.62／0.67／0.67；情境 A 不縮放）；HTML 橫幅與估值章加另一側讀數；既有數值不變。")
    r += 1
    RD.cell(r, 1, f"v{VER}（{BUILT}）：只改展示層與新增列——個股 HTML 版面規格 v1.3（章節：命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態）；三層協調新機制（規則 v1.4.1）：Coord 頁新增 A 讀哪一層哪一版、B C0 由下往上對帳、C 口徑註記。"
                  "新增 Hist 頁（損益表歷史 FY2022–FY2024）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023–FY2025 實績＋FY2026E–FY2031E 情境 B，GAAP）、Val 0.7（四個目標價）、Sens 頁末（加權目標價相對基準）、README 頁末 v1.3 白話列；"
                  "README 標題段版號 v0.1 更正為 v0.2.6。既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價 264.40 與評等強力放空不變。")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT or "Coord" not in wb.sheetnames:   # 另一側讀數建置（alt_run.py）只讀 Val 頁，不加新頁
        return "v026_post：略過（另一側讀數建置或無 Coord 頁）"
    HR = hist_is(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_tor(wb)
    CR = coord_c0(wb, SIL)
    RMA = readme13(wb, RVA, PLA)
    RD = wb["README"]
    a1 = RD["A1"].value
    assert "個股模型 v0.1（AI 半導體" in a1, a1[:80]
    RD["A1"].value = a1.replace("個股模型 v0.1（AI 半導體", f"個股模型 v{VER}（AI 半導體", 1)
    a2 = str(RD["A2"].value)
    assert a2.startswith("版本：v0.1（2026-10-31：初版）。"), a2[:40]
    RD["A2"].value = a2.replace("版本：v0.1（2026-10-31：初版）。", f"版本：v{VER}（{BUILT}；初版 v0.1，2026-10-31）。", 1)
    order = ["README", "In", "Q", "Hist", "Units", "PL", "Val", "DCF", "RevDCF", "Recon", "Sens", "Sources", "Coord"]
    wb._sheets = [wb[n] for n in order]
    return f"v026_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 相對基準；Coord C0；README v1.3 {len(RMA)} 列"
