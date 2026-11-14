# 世芯-KY v0.2.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調（規則 v1.4.2：C0 讀取層級與公司實績對帳、C4 兩讀數）。
# 樣板：創意 v0.3.4（guc/v034_post.py）、艾克爾 v0.4.4、日月光、NVIDIA v0.3.10。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史 FY2022–FY2024（公司合併財報，資料商 stockanalysis 彙整；FY2025 取 Q 頁）
#   PL 頁末：標準損益表（FY2023–FY2025 實績＋FY2026E–FY2031E 情境 B；TIFRS 合併、歸屬母公司）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列（四個目標價）
#   Sens 頁末：加權目標價相對基準（營運輸入案例＝整本重算快照 alchip/sens.json；折現參數案例＝引用上方表；估值框架案例＝封閉式）
#   Coord 頁末：三層協調（A 讀哪一層哪一版、B 讀取層級（同源，不判定）、C 口徑註記、D 摘要、E C0（實績）：公司實績 vs 層級隱含值）
#   README 頁末：v1.3 展示層白話列與版本說明
# 建檔程式在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。另一側讀數建置（COORD_RUN＝1）與 sens_alchip.py 建置（ALCHIP_SENS＝1）時不執行。
import json, os
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.4", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SA = "https://stockanalysis.com/quote/tpe/3661/financials/income-statement/"
TN_OCT = "https://finance.technews.tw/2026/11/10/alchip-3661-202610-financial-report/"
TN_SEP = "https://finance.technews.tw/2026/10/12/alchip-3661-202609-financial-report/"
TN_AUG = "https://finance.technews.tw/2026/09/10/alchip-3661-202608-financial-report/"
TN_JUL = "https://finance.technews.tw/2026/08/10/alchip-3661-202607-financial-report/"
TN_MAR = "https://finance.technews.tw/2026/04/10/alchip-3661-202603-financial-report/"
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


# ---------- Hist：損益表歷史（NT$B；公司合併財報 TIFRS，歸屬母公司；資料商 stockanalysis 彙整，百萬元四捨五入） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025 公式)；None＝不填
HIS = [
    ("rev", "營收", 13.725, 30.482, 51.969, "=Q!B4"),
    ("cogs", "營業成本", 9.296, 23.687, 41.768, "=Q!B4-Q!B7"),
    ("gp", "毛利", 4.430, 6.794, 10.201, "=Q!B7"),
    ("opx", "營業費用（研發、銷管）", 2.118, 3.033, 3.705, "=Q!B10"),
    ("oi", "營業利益", 2.311, 3.761, 6.496, "=Q!B11"),
    ("nonop", "營業外淨額（稅前利益 − 營業利益）", "=2.383-2.311", "=4.198-3.761", "=7.808-6.496", "=Q!B12"),
    ("pbt", "稅前利益", 2.383, 4.198, 7.808, "=Q!B13"),
    ("tax", "所得稅", 0.550, 0.877, 1.362, "=Q!B14"),
    ("ni", "歸屬母公司淨利", 1.834, 3.325, 6.446, "=Q!B16"),
    ("shr", "加權平均股數（十億股；淨利 ÷ 基本 EPS）", "=1.834/25.69", "=3.325/45.47", "=6.446/81.34", "=Q!B17"),
    ("eps", "EPS（基本；NT$）", 25.69, 45.47, 81.34, "=Q!B18"),
    ("epsr", "EPS（基本，公司公告；NT$）", 25.69, 45.47, 81.34, "=Q!B19"),
]


def hist_is(wb):
    Hs = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    Hs["A1"] = ("Hist — 損益表歷史（v0.2.4 起；十億新台幣；EPS 新台幣；TIFRS 合併損益表、歸屬母公司）。PL 頁末「標準損益表」的歷史三年（FY2023–FY2025）與年增率基期用。"
                "FY2022 只取營收、毛利、營業利益、稅前利益、歸屬母公司淨利與 EPS（年增率基期）；FY2025 欄＝Q 頁（FY2025 財報）。只被 PL 頁標準損益表引用，不回饋任何計算")
    Hs["A1"].font = BOLD
    Hs["A2"] = (f"來源（Verified：公司合併財報，資料商 stockanalysis 年度損益表彙整，單位百萬元四捨五入，本表 ÷ 1,000；{SA}）。因四捨五入，核對列容許 ±0.005（FY2023 稅前利益 − 所得稅與歸屬母公司淨利差 0.004，資料商未拆非控制權益，Derived）。"
                "股數＝歸屬母公司淨利 ÷ 公司公告基本 EPS（Derived）")
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025（Q 頁）", "標記", "來源"], 1):
        c = Hs.cell(4, j, h); c.font = BOLD; c.fill = HDR
    rows = {}
    r = 5
    for k, labl, v22, v23, v24, v25 in HIS:
        rows[k] = r
        Hs.cell(r, 1, f"h_{k}"); Hs.cell(r, 2, labl)
        for col, yy, v in (("C", "22", v22), ("D", "23", v23), ("E", "24", v24), ("F", "25", v25)):
            if v is None: continue
            c = Hs[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k in ("eps", "epsr") else ("0.00000" if k == "shr" else "#,##0.000")
            if not isinstance(v, str) or (v.startswith("=") and "!" not in v): c.font = BLUE
            dn(wb, f"h_{k}_{yy}", f"Hist!${col}${r}")
        Hs.cell(r, 7, "Derived" if k in ("shr", "nonop") else "Verified")
        Hs.cell(r, 8, "stockanalysis 年度損益表（TWSE 公開財報彙整）；FY2025 欄取 Q 頁")
        r += 1
    r += 1
    for labl, f in (("核對：營收 − 營業成本 − 毛利（應約為 0；四捨五入容許 ±0.005）", lambda c: f"={c}{rows['rev']}-{c}{rows['cogs']}-{c}{rows['gp']}"),
                    ("核對：毛利 − 營業費用 − 營業利益（應約為 0）", lambda c: f"={c}{rows['gp']}-{c}{rows['opx']}-{c}{rows['oi']}"),
                    ("核對：營業利益 ＋ 營業外淨額 − 稅前利益（應為 0）", lambda c: f"={c}{rows['oi']}+{c}{rows['nonop']}-{c}{rows['pbt']}"),
                    ("核對：稅前利益 − 所得稅 − 歸屬母公司淨利（應約為 0；FY2023 差 0.004）", lambda c: f"={c}{rows['pbt']}-{c}{rows['tax']}-{c}{rows['ni']}")):
        Hs.cell(r, 2, labl)
        for col in "DEF":
            Hs[f"{col}{r}"] = f(col); Hs[f"{col}{r}"].number_format = "0.000"
        r += 1
    Hs.column_dimensions["A"].width = 12; Hs.column_dimensions["B"].width = 48
    for col in "CDEF": Hs.column_dimensions[col].width = 14
    S = wb["Sources"]; n = S.max_row + 1
    for i, (a, b) in enumerate((("資料商：stockanalysis 年度損益表（FY2022–FY2025 合併損益表，Hist 頁；v0.2.4 新增）", SA),
                                ("媒體：科技新報 2026 年 10 月營收（10 月 102.60 億、前 10 月累計 480.60 億；Coord 頁 E 節）", TN_OCT),
                                ("媒體：科技新報 2026 年 9 月營收（9 月 99.79 億、前 9 月累計 378.78 億；Coord 頁 E 節）", TN_SEP),
                                ("媒體：科技新報 2026 年 8 月營收（8 月 87.76 億；Coord 頁 E 節）", TN_AUG),
                                ("媒體：科技新報 2026 年 7 月營收（7 月 74.33 億；Coord 頁 E 節）", TN_JUL),
                                ("媒體：科技新報 2026 年 3 月營收（前 3 月累計 41.86 億；1Q26 營收財報 41.87 億；Coord 頁 E 節）", TN_MAR))):
        S.cell(n + i, 1, a); S.cell(n + i, 2, b)
    return rows


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"       # 標準表欄；E 起（FY2025A…FY2031E）對應 PL 情境 B 區塊 C…I 欄
SRC = "CDEFGHI"


def pl_standard(wb):
    P = wb["PL"]; D = wb["DCF"]
    rbd = lab(P, "情境 B（基準）")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    R = dict(rev=ex("營收"), gp=ex("毛利"), opx=ex("營業費用"), oi=ex("營業利益"), oth=ex("營業外淨額"), pbt=ex("稅前利益"), tax=ex("所得稅"),
             ni=ex("歸屬母公司淨利"), sh=ex("加權股數"), eps=ex("EPS（基本）"))
    dbr = lab(D, "情境 B"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.2.4 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026E–FY2031E，情境 B 基準；FY2026E 為錨點）。十億新台幣；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：TIFRS（台灣採國際財務報導準則；合併財報、歸屬母公司淨利）。世芯無重大非控制權益，稅後淨利即歸屬母公司淨利；EPS 用基本 EPS（本模型口徑，與公司公告同口徑）。預測期的「其他收入（淨額）」＝營業外淨額（利息、匯兌與其他）")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（1H26 實績、3Q26 營收實績與指引、4Q26 假設）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E、FY2031E＝延伸期（收斂法：營收年增與淨利率收斂，只推營收與歸屬母公司淨利）：毛利率與營業費用率沿用 FY2029E 推算營業成本、毛利與營業費用；"
                 "營業利益取 DCF 頁情境 B（由淨利反推）；其他收入＝每年營業外淨額假設；所得稅＝稅前利益 − 歸屬母公司淨利。兩種推法的營業利益差額列在表末核對列，不改既有淨利與目標價")
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

    M = "#,##0.000"; PCT = "0.0%"
    hh = lambda k: [f"=h_{k}_23", f"=h_{k}_24"]
    row("rev", "營收", "NT$B", hh("rev") + [f"=PL!{c}{R['rev']}" for c in SRC], M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L), "=D{rev}/C{rev}-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS[1:], COLS[2:])], PCT)
    r_c = r   # 下一列＝cogs；J、K（FY2030E、FY2031E）用 FY2029E（I 欄）比率
    row("cogs", "營業成本", "NT$B", hh("cogs") + [f"=PL!{c}{R['rev']}-PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-I{r_c + 2})"], M, "FY2030E、FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", hh("gp") + [f"=PL!{c}{R['gp']}" for c in SRC[:5]] + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發、銷管）", "NT$B", hh("opx") + [f"=PL!{c}{R['opx']}" for c in SRC[:5]] + [f"=J{L['rev']}*I{r_c + 3}/I{L['rev']}", f"=K{L['rev']}*I{r_c + 3}/I{L['rev']}"], M, "FY2030E、FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + [f"=PL!{c}{R['oi']}" for c in SRC[:5]] + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外淨額）", "NT$B", hh("nonop") + [f"=PL!{c}{R['oth']}" for c in SRC[:5]] + ["=nonop_f", "=nonop_f"], M, "利息、匯兌與其他；FY2030E 起＝每年營業外淨額假設")
    row("pbt", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "NT$B", hh("tax") + [f"=PL!{c}{R['tax']}" for c in SRC[:5]] + [f"=J{L['pbt']}-PL!H{R['ni']}", f"=K{L['pbt']}-PL!I{R['ni']}"], M, "FY2030E、FY2031E＝稅前利益 − 歸屬母公司淨利（世芯無重大非控制權益）")
    row("ni", "歸屬母公司淨利（EPS 與目標價口徑；＝稅後淨利）", "NT$B", [f"={c}{L['pbt']}-{c}{L['tax']}" for c in COLS], M)
    row("nm", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "加權股數", "十億股", hh("shr") + [f"=PL!{c}{R['sh']}" for c in SRC], "0.00000", "FY2023–FY2024＝歸屬母公司淨利 ÷ 公司公告基本 EPS（Derived）")
    row("eps", "EPS（基本；NT$）", "NT$", ["=h_eps_23", "=h_eps_24"] + [f"=PL!{c}{R['eps']}" for c in SRC], "0.00", "FY2023–FY2024＝公司公告基本 EPS；FY2025 起＝模型（歸屬母公司淨利 ÷ 加權股數）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY；歸屬母公司淨利）", "%", ["=C{ni}/h_ni_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 歸屬母公司淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；公司公告基本 EPS，與上方 EPS 的差只在 FY2025 取模型淨利 ÷ 股數的捨入）").font = BOLD; r_gaap = r; r += 1
    row("g_epsd", "EPS（基本，公司公告；NT$）", "NT$", ["=h_epsr_23", "=h_epsr_24", "=h_epsr_25"], "0.00", "FY2025 公司公告 69.18（Q 頁第 19 列）")
    P.cell(r, 1, "核對（應為 0；FY2023–FY2024 歷史列因百萬元四捨五入容許 ±0.005）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E、FY2031E 見下兩列）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：歸屬母公司淨利 − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$B", [None, None] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    row("c4", "核對：EPS − 情境 B（PL 上方區塊；FY2025A–FY2031E）", "NT$", [None, None] + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip(COLS[2:], SRC)], "0.000")
    hr_ = lab(wb["Hist"], "h_ni", col=1)
    row("c5", "核對：歸屬母公司淨利（Hist）− 本表（FY2023–FY2025）", "NT$B", [f"=Hist!D{hr_}-C{L['ni']}", f"=Hist!E{hr_}-D{L['ni']}", f"=Hist!F{hr_}-E{L['ni']}"], "0.000")
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
    rbd = lab(P, "情境 B（基準）"); rev26 = lab(P, "營收", rbd, exact=True)
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    nc_ref = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.4 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In、PL 頁，無任何頁引用本頁，不影響目標價。金額十億新台幣；價格新台幣")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。FY2027–FY2032 全部為預測年（世芯 FY2026 為錨點，不動）")
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
    kv("px", "現價（新台幣）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（十億新台幣）", f"={nc_ref}", "#,##0.000", "同 DCF 頁每股價值公式（Val 頁情境 B：淨現金與投資＋stub 期 FCFF − stub 期股利）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.00000")
    kv("evr", "支持現價所需的企業價值 EV（十億新台幣）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.000")
    kv("ev", "模型企業價值 EV（情境 B；十億新台幣）", f"=DCF!$C${ev}", "#,##0.000")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（新台幣）", f"=DCF!$C${ps}", "#,##0.00")
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
    yr("rev", "營收（十億新台幣）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.000")
    yr("ebit", "營業利益 EBIT（十億新台幣）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.000")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027–FY2032 皆為預測年）", lambda dc, c: f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（十億新台幣）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.000")
    yr("wc", "營運資金增加（十億新台幣）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.000")
    yr("pv", "現值（十億新台幣，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2027–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（十億新台幣）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.000")
    kv("E0", "與營業利益無關的 EV（十億新台幣）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.000")
    kv("m0", "模型：基準的等效營業利益率（FY2027–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(B{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(B{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2027–FY2032 營收等比例放大；FY2026 錨點不動；營業利益率、稅率、折舊、資本支出占營收比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（十億新台幣）", f"=fc_wc*PL!$D${rev26}*B{Y['fr']}*B{Y['df']}", "#,##0.000", "FY2027 營運資金增加中屬 FY2026 營收的部分（FY2026 為錨點，不放大）")
    kv("b", "隨營收倍數等比例變動的 EV（十億新台幣）",
       f"=SUMPRODUCT(B{Y['cf']}:G{Y['cf']},B{Y['fr']}:G{Y['fr']},B{Y['df']}:G{Y['df']})-fc_wc*B{Y['rev']}*B{Y['fr']}*B{Y['df']}"
       f"-SUMPRODUCT(C{Y['wc']}:G{Y['wc']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.000")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2027 營收（十億新台幣）", f"=B{A['s1']}*B{Y['rev']}", "#,##0.000")
    kv("rv0", "模型：FY2027 營收（情境 B；十億新台幣）", f"=B{Y['rev']}", "#,##0.000")
    kv("g1", "現價隱含的 FY2027 營收年增率（相對 FY2026 錨點營收）", f"=B{A['rv1']}/PL!$D${rev26}-1", "0.0%", "FY2026 為錨點（1H26 實績、3Q26 營收實績與指引、4Q26 假設），不放大")
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
        RV.cell(r, 5).number_format = "#,##0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%；落在下界表示即使折現率壓到下界也不足以支持現價")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（新台幣；二分法收斂時應接近 0，落在搜尋下界時不為 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "#,##0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"#,##0")',
                     T(" 元；要支持現價 "), 'TEXT(px,"#,##0")', T(" 元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大為","縮小為")', T(" "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要："), f'IF(B{A["w1"]}<=g_tv+0.0051,"折現率低於搜尋下界 "&TEXT(B{A["w1"]},"0.0%")&"（永續成長率＋0.5 個百分點，仍不足以支持現價）","折現率"&IF(B{A["w1"]}<=wacc,"降到 ","升到 ")&TEXT(B{A["w1"]},"0.0%"))', T("；或 FY2027–FY2032 營收全面"), f'IF(B{A["s1"]}>=1,"放大","縮小")',
                     T(" "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2027 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億元，相對 FY2026 錨點年增 "), f'TEXT(B{A["g1"]},"0%")',
                     T("）；或營業利益率"), f'IF(B{A["m1"]}>=B{A["m0"]},"提高到","降到")', T(" "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"#,##0")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 世芯過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給世芯的倍數是同業的 "),
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
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.2.4 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價（賣方平均）", cons, '="stockanalysis 賣方平均目標價（In 頁 cons_tp，Interested-party）"')]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "In 頁 px"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列） ----------
TOR_PARAM = (("WACC", "WACC −1／＋1 個百分點（永續期新投資報酬率隨 WACC）"), ("β", "股權 β 低／高（產業無槓桿 β ∓0.3 後重新加槓桿）"), ("ERP", "股權風險溢酬 4%／6%（基準 5%）"),
             ("永續成長率", "永續成長率 −1 個百分點／上限（基準已在上限，高案無差）"), ("RONIC", "永續期新投資報酬率高於 WACC 5 個百分點（只有高案）"), ("資本支出 ÷ 營收", "每年資本支出占營收 −1／＋1 個百分點"),
             ("營運資金", "營運資金占營收低／高"), ("淨現金不含", "淨現金不含權益投資（只有低案）"), ("終值改用", "DCF 終值改用 FCFF ×（1＋g）法（只有高案）"),
             ("美元口徑折現", "改用美元口徑折現（美國 10 年期 rf；只有低案）"))


def sens_tor(wb, here):
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")                  # B、C、D＝情境 A、B、C
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    bl = {s: lab(P, f"情境 {s}（") for s in "ABC"}
    epr = {s: lab(P, "EPS（基本）", bl[s], exact=True) for s in "ABC"}
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.2.4 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（alchip/sens_alchip.py：三情境同時套用案例覆寫，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（新台幣）", "高案：加權目標價（新台幣）", "來源"], 1):
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
        SE.cell(r, 6, "整本重算快照（alchip/sens.json）；三情境同時套用案例覆寫"); r += 1
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
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 alchip/sens_alchip.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調（A 讀哪一層、B 讀取層級、C 口徑、D 摘要、E C0 實績） ----------
MON = 10     # 2026 年已公布月營收月數（1–10 月）


def coord_c0(wb, SIL):
    C = wb["Coord"]; I = wb["In"]; P = wb["PL"]; V = wb["Val"]
    rbd = lab(P, "情境 B（基準）")
    ex = lambda t, s=rbd: lab(P, t, s, exact=True)
    r_u = ex("AWS Trainium 市場顆數（含情境係數）"); r_sh = ex("世芯份額：AWS Trainium"); r_pp = lab(P, "每顆世芯營收（核心＋代購 HBM；未乘 k_x）", rbd)
    r_kx = lab(P, "校準係數 k_x", rbd); r_trn = ex("Trainium 量產營收"); r_oth = lab(P, "其他（NRE＋非 Trainium 量產）", rbd); r_rev = ex("營收")
    ikey = {str(I.cell(i, 1).value): i for i in range(4, I.max_row + 1) if I.cell(i, 1).value}
    C.column_dimensions["F"].width = 60
    for col in "BCDE": C.column_dimensions[col].width = 24
    r_ratio = lab(C, "另一側讀數 ÷ 現行基準")
    r = C.max_row + 3
    C.cell(r, 1, "三層協調（v0.2.4 新增；三層基準協調規則 v1.4.2 C0 個股與層的一致性、C4 個股不裁決、C5 往上回饋只有證據）").font = BOLD; r += 1
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
    tsm = "經台積電個股模型成品（tsmc/dist；該模型讀 Siliconomics v0.19）"
    arow("AWS Trainium 顆數（原始顆數；百萬顆）", [f"=In!E{ikey[f'u_trn_{y}']}" for y in (26, 27, 28)] + [None], f"{ver}：SIF_UnitsPlat_trn（AWS Trainium），{tsm}取得（In 頁 u_trn_26–u_trn_28，唯讀；CY2029 由 g29_trn 外推）")
    arow("情境 B 顆數係數（÷ 原始顆數）", [None, 1, f"=In!E{ikey['um28_B']}", f"=In!E{ikey['um29_B']}"], f"{ver}：SIF_UnitsCoef_Base，{tsm}取得（In 頁 um28_B、um29_B；CY2026–27 為 1）")
    arow("世芯占 AWS Trainium 顆數份額（情境 B；基準）", [f"=In!E{ikey['al_trn_26']}", f"=In!E{ikey['al27_B']}", f"=In!E{ikey['al28_B']}", f"=In!E{ikey['al28_B']}"],
         f"{ver}：SIF_al_trn_26、SIF_al27_Base、SIF_al28_Base（設計夥伴份額；In 頁 al_trn_26、al27_B、al28_B 讀自 Siliconomics；CY2029 沿用 CY2028；Assumed）", "0.0%")
    arow("Wattonomics 錨定係數（只進 Sens 案例，不改基準）", [None] + [f"=In!E{ikey[f'wta_{y}']}" for y in (27, 28, 29)], f"{ver}：SIF_WTAnchor_Base（R4 比值＝Wattonomics ÷ 本庫推得上線 GW）；Wattonomics 現行版本 v0.3.2")
    arow("另一側讀數 ÷ 現行基準（對帳中年份；兩讀數目標價用）", [f"={'BCDE'[i]}{r_ratio}" for i in range(4)], f"{ver}：SIF_DeployGW_AltRatio_<年>（Wattonomics 換算的產業 GW ÷ Siliconomics 基準）；上方 Coord 第 6 列")
    C.cell(r, 1, "Wattonomics：本模型不直接讀 Wattonomics 數字；現行版本 v0.3.2，僅經 Siliconomics 的 R4 比值（Sens 錨定案例）與另一側讀數比值（對帳中年份兩讀數）間接使用。產業量（AWS Trainium 顆數）讀 Siliconomics 基準（經台積電成品）；世芯份額 CY2026–28 讀 Siliconomics 的設計夥伴份額；"
                 "每顆規格讀 Broadcom 模型（Siliconomics SIF_pa_trn）；HBM 代購比例、2 奈米每顆營收係數與毛利率為本模型假設，校準係數由 FY2026 公司 3 奈米營收說法反推。"); R["A_note"] = r; r += 2
    # ---- B. 讀取層級（同源，不判定） ----
    C.cell(r, 1, "B. 讀取層級（同源，不判定）：世芯讀入的 Siliconomics 值 vs 層級值（個股自層級讀入，差距恆為 0，無檢驗力；不計入對帳中項目數；規則 v1.4.2 C0 第 2 點）。"
                 "公司實績對層級隱含值的對帳見本頁 E 節「C0（實績）」").font = BOLD; R["B"] = r; r += 1
    for j, h in enumerate(["項目（年）", "本模型原值", "層級對應值", "差距（本模型 ÷ 層級 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["B_hdr"] = r; r += 1
    gw_b = SIL.NAMES["SIF_DeployGW_Base"]
    R["B_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        col = "DEFG"[i]
        C.cell(r, 1, f"AWS Trainium 顆數（情境 B）CY{y}（百萬顆）")
        C.cell(r, 2, f"=PL!{col}{r_u}").number_format = "0.000"
        if y <= 2027: C.cell(r, 3, f"=In!E{ikey[f'u_trn_{y % 100}']}")
        elif y == 2028: C.cell(r, 3, f"=In!E{ikey['u_trn_28']}*In!E{ikey['um28_B']}")
        else: C.cell(r, 3, f"=In!E{ikey['u_trn_28']}*(1+g29_trn)*In!E{ikey['um29_B']}")
        C.cell(r, 3).number_format = "0.000"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, f"本模型：PL 頁情境 B 顆數（原始顆數 × 情境係數；來自台積電個股成品）；層：{ver} 供給側基準顆數。同源讀取故差距 0；世芯不公布 Trainium 顆數，兩邊都沒有公司實績（等級：Assumed／Interested-party）" if i == 0 else "同上")
        r += 1
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        col = "DEFG"[i]
        C.cell(r, 1, f"世芯占 AWS Trainium 顆數份額（情境 B）CY{y}")
        C.cell(r, 2, f"=PL!{col}{r_sh}").number_format = "0.0%"
        C.cell(r, 3, f"=In!E{ikey['al_trn_26' if y == 2026 else 'al27_B' if y == 2027 else 'al28_B']}").number_format = "0.0%"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, f"本模型：PL 頁情境 B 份額；層：{ver} 設計夥伴份額（CY2029 沿用 CY2028）。同源讀取故差距 0；世芯未揭露 Trainium 份額（Assumed；公司僅說 Trainium 3 主供應商之一）" if i == 0 else "同上")
        r += 1
    R["B_last"] = r - 1
    C.cell(r, 1, "世芯 Trainium 營收（FY2026–FY2029）vs 層級量"); C.cell(r, 2, "—"); C.cell(r, 3, "層級量與本模型校準同源"); C.cell(r, 4, "—"); C.cell(r, 5, "不判定（參考（同源校準），見 E 節）")
    C.cell(r, 6, "Siliconomics 的 AWS Trainium 顆數、世芯份額與 Broadcom 模型每顆規格讀入本模型；每顆規格取已扣 HBM 的核心價值（世芯代購 HBM 比例 0），與世芯每顆營收（約 US$2.5K）口徑不同，需校準係數換算；"
                 "校準係數由 FY2026 3 奈米營收「超過 15 億美元」（媒體轉述外資報告，時報資訊 2026-08-14，Interested-party）反推（Derived），層級隱含值乘上校準係數後恆等於錨點（見 E 節同源檢查）"); R["B_nocmp"] = r; r += 1
    R["B_end"] = r - 1
    C.cell(r, 1, "判定摘要"); R["B_sum"] = r; r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; R["B2"] = r; r += 1
    for j, h in enumerate(["項目（年）", "Siliconomics 基準", "外部估計（Wattonomics 換算）", "差距（基準 ÷ 外部 − 1）", "判定", "口徑、來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    R["B2_first"] = r
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        C.cell(r, 1, f"產業部署 GW CY{y}（Siliconomics 基準 ÷ Wattonomics 換算；GW）")
        c = C.cell(r, 2, gw_b[i]); c.number_format = "0.00"; c.font = BLUE
        c = C.cell(r, 3, f"=B{r}*{'BCDE'[i]}{r_ratio}"); c.number_format = "0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "參考")
        C.cell(r, 6, (f"基準＝{ver} SIF_DeployGW_Base（供給側部署 GW）；外部估計（參考）＝Wattonomics 換算 SIF_DeployGW_Alt_<年>（需求側）。"
                      "這是層與層之間的差距，不是世芯自己的量對層級的比較；由 Siliconomics 對帳單 r3 處理，不計入本公司 C0 項目數；本模型依 C4 並列兩讀數目標價，不裁決") if i == 0 else "同上")
        r += 1
    R["B2_last"] = r - 1
    r += 1
    # ---- C. 口徑註記 ----
    C.cell(r, 1, "C. 對帳中年份處理與口徑註記（C4）").font = BOLD; R["C"] = r; r += 1
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    C.cell(r, 1, F(T("對帳中年份：CY2027、CY2028、CY2029（Siliconomics 標示；本公司 C0 無可判定項目）。兩讀數目標價並列於上方（現行基準讀數 "), f'TEXT(B{rtp},"#,##0.00")', T(" 元、另一側讀數 "), f'TEXT(C{rtp},"#,##0.00")',
                    T(" 元），評等分別為"), f'B{rrec}', T("、"), f'C{rrec}', T("。個股不裁決（C4）；Siliconomics 與 Wattonomics 的差距由對帳單處理（期限 2027-02-28）；結案前兩個目標價並列。")))
    R["C_1"] = r; r += 1
    C.cell(r, 1, F(T("口徑註記（不改模型）：另一側讀數只改情境 B、C 的 AWS Trainium 顆數（CY2027 起乘比值 "), f'TEXT(C{r_ratio},"0.00")', T("／"), f'TEXT(D{r_ratio},"0.00")', T("／"), f'TEXT(E{r_ratio},"0.00")',
                    T("）；情境 A（樂觀，權重 "), 'TEXT(tp_wA,"0%")', T("）不縮放，世芯份額、2 奈米每顆營收係數、毛利率與其他營收假設不變。世芯每顆營收由 FY2026 校準，故 Trainium 量產營收約隨顆數等比例下降。"
                    "這個口徑已含在另一側目標價內；對帳單處理時一併說明。")))
    R["C_2"] = r; r += 1
    C.cell(r, 1, "C0 結論：世芯的產業量（AWS Trainium 顆數）與設計夥伴份額全部讀 Siliconomics 基準（經台積電成品），與層同源，只列「讀取層級」、不判定；公司實績（月營收 1–10 月、季報 1Q26–2Q26 與 3Q26 營收）已列入 E 節。"
                 "Siliconomics 沒有世芯的營收：層級的 Trainium 每顆規格取已扣 HBM 的核心價值，與世芯每顆營收（晶片與封裝轉售，約 US$2.5K）口徑不同，層級隱含的 Trainium 營收由顆數、份額、每顆規格與校準係數算出，等於錨點（同源校準，E 節同源檢查），"
                 "公司又未揭露 Trainium 營收金額，故 C0（實績）無可判定項目，標「無可比實績」或「參考（同源校準）」，不產生對帳中項目。往上回饋只有證據（C5）：待世芯揭露 Trainium 營收、出貨顆數或客戶別營收，或層級補上獨立的 ASIC 設計服務市場口徑與份額後再評估是否送對帳單。"); R["C_3"] = r; r += 2
    # ---- D. 摘要 ----
    C.cell(r, 1, "D. 摘要（HTML 與版面檢查用；兩讀數與損益表年度標示）").font = BOLD; R["D"] = r; r += 1
    for j, h in enumerate(["項目", "現行基準讀數", "另一側讀數", "另一側相對現行", ""], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    r_tp0 = lab(C, "加權目標價（DCF 與 Comparable 加權）"); r_up0 = lab(C, "上檔空間（加權目標價"); r_rec0 = lab(C, "投資建議（五級")
    R["tp"] = r; C.cell(r, 1, "加權目標價")
    C.cell(r, 2, f"=B{r_tp0}").number_format = "#,##0.00"; C.cell(r, 3, f"=C{r_tp0}").number_format = "#,##0.00"; C.cell(r, 4, f"=D{r_tp0}").number_format = "+0.0%;-0.0%"; r += 1
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
    C.cell(r, 1, "比較方法：個股側只用公司已公布實績（月營收公告與季報，Verified：合併營收；世芯不揭露 Trainium 營收金額，Trainium 與其他的拆分為模型假設 Assumed，不算實績）；"
                 "層級側用 Siliconomics v0.19 同口徑隱含值（Trainium 顆數 × 世芯份額 × 每顆規格 × 校準係數 × 匯率）。時差：層級值為日曆年 CY2026，世芯會計年度＝日曆年；"
                 f"主讀數＝2026 年 1–{MON} 月累計月營收 vs 層級 CY2026 × {MON} ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026（約等於 3Q 中點的年化速率；世芯 3Q 為 Trainium 3 奈米量產首個完整季，成長年中偏高，年化偏高）。"
                 "判定：兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致；層級側沒有同口徑隱含值＝無可比實績；層級隱含值由本公司模型校準＝參考（同源校準），不計入。"
                 "公司指引、賣方估計、本模型自己的推算只列參考，不判定。")
    R["E_note"] = r; r += 1
    for j, h in enumerate(["項目", f"公司實績：2026 年 1–{MON} 月累計（主讀數）", "公司實績：3Q26 × 4（第二讀數）", "層級隱含：CY2026 全年", "判定（±15%）", "口徑、出處與證據等級",
                           f"主讀數差距（實績 ÷（層級 × {MON}÷12）− 1）", "第二讀數差距（實績 ÷ 層級 − 1）", "結論"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["E_hdr"] = r; r += 1
    R["E_first"] = r
    # 版面：C0 表 3 列＋參考列 2＋判定摘要 1 → 輸入區（月營收與季報）→ 同源檢查區
    r_in = r + 6
    r_rv = r_in + 2          # 輸入區：合併營收列
    r_ann = r_in + 3         # 公告累計列
    r_mo = r_in + 4          # 單月營收（參考）
    r_ch = r_in + 10         # 同源檢查區標題列（輸入區 5 列＋檢查 4 列＋空 1 列之後）
    r_chc = r_ch + 2

    def row(label, b, c, d, src, concl, verdict=None, nb="#,##0.000"):
        nonlocal r
        C.cell(r, 1, label)
        for col, v in ((2, b), (3, c)):
            x = C.cell(r, col, v)
            if not isinstance(v, str) or v.startswith("="): x.number_format = nb
        y = C.cell(r, 4, d)
        if not isinstance(d, str) or d.startswith("="): y.number_format = nb
        C.cell(r, 7, f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r})),B{r}/(D{r}*{MON}/12)-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 8, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}/D{r}-1,"—")').number_format = "+0.0%;-0.0%"
        if verdict:
            C.cell(r, 5, verdict)
        else:
            C.cell(r, 5, f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),IF(AND(ABS(G{r})>0.15,ABS(H{r})>0.15,SIGN(G{r})=SIGN(H{r})),"對帳中",IF(OR(ABS(G{r})>0.15,ABS(H{r})>0.15),"時差邊界（參考）","一致")),"無可比實績")')
        C.cell(r, 6, src)
        C.cell(r, 9, concl)
        r += 1

    row("合併營收（NT$B）", f"=G{r_ann}", f"=D{r_rv}*4", "無同口徑隱含值",
        f"公司實績：月營收公告 1–{MON} 月累計 48.060（科技新報 2026-11-10：10 月 10.260、前 10 月累計 480.60 億元）與 3Q26 26.051（前 9 月累計 378.78 億 − 1H26 118.27 億；Q 頁）；輸入區；Verified。"
        "層級側：Siliconomics v0.19 沒有世芯或 ASIC 設計服務的營收（層級內只有 AWS Trainium 顆數、設計夥伴份額與每顆規格；本模型取已扣 HBM 的核心價值，與世芯每顆營收口徑不同，需校準係數換算，校準係數為本模型校準）。",
        "無可比實績：缺層級的 ASIC 設計服務（晶片與封裝轉售）營收市場口徑，或世芯 Trainium 營收、出貨顆數的揭露")
    row("Trainium 量產營收（NT$B）", "無實績", "無實績", f"=G{r_chc}",
        "公司實績：世芯不揭露 Trainium 營收金額（公司只說 3 奈米 AI 加速器 5 月底量產、6 月出貨，FY2026 3 奈米營收「超過 15 億美元」為媒體轉述外資報告（時報資訊 2026-08-14），Interested-party，不是公司說法，也不是已實現營收；Q 頁第 5 列 2Q26、3Q26 的 Trainium 金額為媒體口述與本模型假設，Interested-party／Assumed，不算實績）。"
        "層級隱含：Siliconomics 的 AWS Trainium 顆數 CY2026 × 世芯份額 × 每顆規格 × 校準係數 × 匯率（E 節同源檢查區；校準係數由上述 FY2026 3 奈米營收數字反推，隱含值＝錨點）。",
        "無可比實績（層級值同源）：校準係數吸收顆數、份額、每顆規格與匯率，層級隱含＝本公司錨點；缺獨立揭露：世芯 Trainium 營收或出貨顆數", verdict="參考（同源校準）")
    row("其他（NRE＋非 Trainium 量產；NT$B）", "無實績", "無實績", "無同口徑隱含值",
        "公司實績：世芯只揭露單一營運部門與季度營收，委託設計（NRE）與量產（turnkey）的金額未公布（FY2025 應用別比重為公司揭露，FY2026 尚無；本模型其他營收以成長率簡化推算，Assumed）。"
        "層級側無 NRE 或非 Trainium 量產的市場值（層級的其他 ASIC 顆數與世芯其他營收無對應）。",
        "無可比實績：缺公司的 NRE 與非 Trainium 量產季度金額，且層級無對應市場值")
    R["E_last"] = r - 1
    C.cell(r, 1, "參考（不判定）：十月實績 × 3 vs 本模型 4Q26 營收假設（NT$B）")
    c = C.cell(r, 2, f"=F{r_rv}*3"); c.number_format = "#,##0.000"
    c = C.cell(r, 3, "=Q!F4"); c.number_format = "#,##0.000"
    C.cell(r, 4, "—")
    C.cell(r, 5, "參考（不判定）")
    C.cell(r, 6, "B 欄＝10 月單月營收 10.260（月營收公告，Verified）× 3（Derived：假設 11–12 月與 10 月持平）；C 欄＝本模型 4Q26 營收（Q 頁第 4 列：3Q26 ×（1＋4Q 季增假設），Assumed）。"
                 "這是公司月營收實績對本模型自己的假設，不是公司實績對層級隱含值，不是 C0 證據、不判定。")
    C.cell(r, 7, f'=TEXT(B{r}/C{r}-1,"+0.0%;-0.0%")&"（十月 × 3 ÷ 模型 4Q26 − 1）"')
    C.cell(r, 8, "—")
    C.cell(r, 9, "參考")
    R["E_ref"] = r; r += 1
    # PR #534 獨立核對建議：補 3Q26 法說製程占比（參考、不判定）與 3Q26 財報已公布註記（本版錨點未更新；佔用原判定摘要後的空列，輸入區列號不變）
    C.cell(r, 1, "參考（不判定）：世芯 3Q26 法說 3／2 奈米占營收比重")
    C.cell(r, 2, "80%（3Q26，公司法說）"); C.cell(r, 3, "—"); C.cell(r, 4, "—")
    C.cell(r, 5, "參考（不判定）")
    C.cell(r, 6, "世芯 3Q26 法說：3／2 奈米占營收 80%（公司法說說法，Interested-party）。製程別比重不是 Trainium 營收金額，層級側亦無同口徑值，不判定。"
                 "另：世芯 3Q26 財報已公布（優分析 2026-11-07：營收 8.193 億美元、EPS 37.68、毛利率低於 20%）；本版 FY2026 錨點仍為 3Q26 月營收實績＋指引，未更新，待下版。")
    C.cell(r, 7, "—"); C.cell(r, 8, "—"); C.cell(r, 9, "參考")
    R["E_ref2"] = r; r += 1
    C.cell(r, 1, "判定摘要"); R["E_sum"] = r
    rg = f'E{R["E_first"]}:E{R["E_last"]}'
    C.cell(r, 2, F(T("C0（實績）共 "), f'COUNTA({rg})', T(" 項：可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})', T(" 項、對帳中 "), f'COUNTIF({rg},"對帳中")', T(" 項、時差邊界 "), f'COUNTIF({rg},"時差邊界（參考）")',
                   T(" 項、參考（同源校準）"), f'COUNTIF({rg},"參考（同源校準）")', T(" 項、無可比實績 "), f'COUNTIF({rg},"無可比實績")',
                   T(" 項。原因：公司公布合併營收，但未公布 Trainium 營收金額；Siliconomics 沒有世芯營收，層級隱含 Trainium 營收由層級顆數、份額、每顆規格與校準係數算出，等於錨點（同源校準：校準係數吸收顆數、份額、每顆規格與匯率，"
                   "實際只檢驗「實績 vs 本公司錨點」，不是獨立對帳）。本模型不以模型預測或同源值補位；待世芯揭露 Trainium 營收或出貨顆數，或層級補上獨立的 ASIC 設計服務市場口徑與份額後再判定。")))
    R["E_sum_rg"] = rg
    r += 1
    assert r == r_in, (r, r_in)
    C.cell(r, 1, "輸入：公司合併營收（十億新台幣；月營收公告與季報；Verified；藍字為輸入）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "1Q26", "2Q26", "3Q26", "1–3Q 累計", "10 月單月", "1–10 月累計（單月加總）", "出處"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    assert r == r_rv, (r, r_rv)
    C.cell(r, 1, "合併營收（NT$B；分季）")
    for j, v in ((2, "=Q!C4"), (3, "=Q!D4"), (4, "=37.878-B{0}-C{0}".format(r)), (6, 10.260)):
        c = C.cell(r, j, v); c.number_format = "#,##0.000"
        if j == 6: c.font = BLUE
    C.cell(r, 5, f"=B{r}+C{r}+D{r}").number_format = "#,##0.000"
    C.cell(r, 7, f"=E{r}+F{r}").number_format = "#,##0.000"
    C.cell(r, 8, f"1Q26、2Q26：Q 頁（財報：1Q26 41.87 億、1H26 118.27 億 − 1Q26）；3Q26：前 9 月累計 378.78 億（{TN_SEP}）− 1H26（Q 頁）；10 月：102.60 億（{TN_OCT}）；Verified。G 欄＝1–3Q 累計＋10 月單月（單月加總，非公告值）")
    r += 1
    assert r == r_ann
    C.cell(r, 1, "公司公告累計（月營收公告；核對用；主讀數取此列 1–10 月累計）")
    c = C.cell(r, 5, 37.878); c.font = BLUE; c.number_format = "#,##0.000"
    c = C.cell(r, 7, 48.060); c.font = BLUE; c.number_format = "#,##0.000"
    C.cell(r, 8, f"前 9 月累計 378.78 億（{TN_SEP}）；前 10 月累計 480.60 億（{TN_OCT}）；Verified。公司公告累計與單月加總不連續（可能為外幣換算）：前 8 月差 0.054、前 10 月差 0.078、3Q 以累計減 1H 為 26.051 對單月加總 26.188 差 0.137（NT$B）；不影響 ±15% 判定"); r += 1
    assert r == r_mo
    C.cell(r, 1, "輸入（分月，參考）：7 月、8 月、9 月、10 月單月營收（NT$B）")
    for j, v in ((2, 7.433), (3, 8.776), (4, 9.979), (5, 10.260)):
        c = C.cell(r, j, v); c.font = BLUE; c.number_format = "#,##0.000"
    C.cell(r, 8, f"7 月 74.33 億（{TN_JUL}）；8 月 87.76 億（{TN_AUG}）；9 月 99.79 億（{TN_SEP}）；10 月 102.60 億（{TN_OCT}）；Verified"); r += 1
    C.cell(r, 1, "輸入（分月）：3Q26 營收 − In 頁 3Q26 營收輸入（應為 0）"); C.cell(r, 4, f"=ROUND(D{r_rv}-q3_rev,6)").number_format = "0.000"; R["E_chk0"] = r; r += 1
    C.cell(r, 1, "檢查：1–3Q 累計 − 公告前 9 月累計（應為 0）"); R["E_chk1"] = r
    C.cell(r, 5, f"=ROUND(E{r_rv}-E{r_ann},6)").number_format = "0.000"; r += 1
    C.cell(r, 1, "檢查：1Q26、2Q26 − Q 頁（應為 0）"); R["E_chk2"] = r
    C.cell(r, 2, f"=ROUND(B{r_rv}-Q!C4,6)").number_format = "0.000"; C.cell(r, 3, f"=ROUND(C{r_rv}-Q!D4,6)").number_format = "0.000"; r += 1
    C.cell(r, 1, "參考：公告 1–10 月累計 − 單月加總（非應為 0；公司公告累計與單月加總不連續，可能為外幣換算；NT$B）"); R["E_chk3"] = r
    C.cell(r, 7, f"=ROUND(G{r_ann}-G{r_rv},6)").number_format = "0.000"; r += 1
    r += 1
    assert r == r_ch, (r, r_ch)
    C.cell(r, 1, "層級隱含值的取法與同源檢查（CY2026；層級顆數 × 世芯份額 × 每顆規格 × 校準係數 × 匯率）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層級顆數 CY2026（百萬顆；Siliconomics 經台積電成品，In 頁）", "世芯份額（Siliconomics 基準；Assumed）", "每顆規格（US$K；核心＋代購 HBM，未乘校準係數；Broadcom 模型每顆規格扣 HBM）",
                           "校準係數（本模型；FY2026 3 奈米營收「超過 15 億美元」反推，媒體轉述外資報告，Derived）", "匯率（NT$/US$）", "層級隱含 CY2026（NT$B）", "同源檢查：隱含 − FY2026 錨點（應為 0）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    assert r == r_chc
    C.cell(r, 1, "Trainium 量產（Trainium 顆數 × 世芯份額 × 每顆規格 × 校準係數 × 匯率）")
    C.cell(r, 2, "=u_trn_26").number_format = "0.000"; C.cell(r, 3, "=al_trn_26").number_format = "0.0%"; C.cell(r, 4, f"=PL!D{r_pp}").number_format = "0.000"
    C.cell(r, 5, f"=PL!D{r_kx}").number_format = "0.0000"; C.cell(r, 6, "=fx26").number_format = "0.00"
    C.cell(r, 7, f"=B{r}*C{r}*D{r}*E{r}*F{r}").number_format = "#,##0.000"; C.cell(r, 8, f"=ROUND(G{r}-a26_trn,9)").number_format = "0.000"; r += 1
    C.cell(r, 1, F(T("相消的量：「層級顆數 × 份額 × 每顆規格」被校準係數完全吸收（校準係數＝FY2026 3 奈米營收「超過 15 億美元」（媒體轉述外資報告，時報資訊 2026-08-14，Interested-party）÷ 未校準營收），所以層級隱含值恆等於錨點（H 欄為 0）；不乘校準係數的層級未校準值 "),
                   f'TEXT(B{r - 1}*C{r - 1}*D{r - 1}*F{r - 1},"0.0")', T(" 十億元（已扣 HBM 的核心價值）是錨點 "), f'TEXT(a26_trn,"0.0")', T(" 十億元的 "), f'TEXT(B{r - 1}*C{r - 1}*D{r - 1}*F{r - 1}/a26_trn,"0.0")',
                   T(" 倍；差距原因：Broadcom 模型核心價值與世芯每顆營收（約 US$2.5K）口徑落差、份額與時程（CY2026 僅約 7 個月量產）等，皆被校準係數吸收。公司實績只能和這個錨點比，不是對層級顆數、份額或規格的獨立檢驗；"
                     "要獨立對帳需世芯揭露 Trainium 營收、出貨顆數或每顆價格，或 Siliconomics 補上世芯口徑的市場值。"))); R["E_cancel"] = r; r += 1
    C.cell(r, 1, "FY2026 錨點（NT$B）：Trainium／其他／合併營收（Q 頁 FY2026 錨點；1H26 實績、3Q26 營收實績與指引、4Q26 假設）")
    C.cell(r, 2, "=a26_trn").number_format = "#,##0.000"; C.cell(r, 3, "=a26_oth").number_format = "#,##0.000"; C.cell(r, 4, "=a26_rev").number_format = "#,##0.000"
    R["E_end"] = r; r += 1
    # ---- B 節判定摘要改指向 E 節計數 ----
    C.cell(R["B_sum"], 2, F(T("本公司 C0 對帳中項目 "), f'COUNTIF({rg},"對帳中")', T(" 項：公司實績對層級隱含值的 C0（實績）表見 E 節（可判定 "), f'COUNT(G{R["E_first"]}:G{R["E_last"]})',
                             T(" 項；其餘為「參考（同源校準）」或層級側沒有同口徑隱含值的「無可比實績」）。本節顆數與份額 CY2026–29 是個股讀入的層級值，與層級同源，標「讀取層級（同源，不判定）」，不計入。"
                               "下方「層級之間」的產業部署 GW 差距只列參考，不計入。對帳單由主控與 Siliconomics 層處理，本模型不改上游值。")))
    # 具名範圍（HTML 與核對程式用）
    wb.defined_names["COORD_C0"] = DefinedName("COORD_C0", attr_text=f"Coord!$A${R['E_first']}:$I${R['E_last']}")
    return R


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, TO, CR):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; C = wb["Coord"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, coef = vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    feps = vb("12 個月後的前瞻 EPS（情境加權）")
    geps, gmul = vb("與市場共識目標價的差距：EPS 因素"), vb("與市場共識目標價的差距：倍數因素")
    bl = {s: lab(P, f"情境 {s}（") for s in "ABC"}
    rb = bl["B"]
    epr = {s: lab(P, "EPS（基本）", bl[s], exact=True) for s in "ABC"}
    bur = {s: lab(P, "世芯顆數：AWS Trainium", bl[s], exact=True) for s in "ABC"}
    trr, otr = lab(P, "Trainium 量產營收", rb, exact=True), lab(P, "其他（NRE＋非 Trainium 量產）", rb)
    revr = lab(P, "營收", rb, exact=True)
    tor = lab(SE, "加權目標價相對基準（HTML")
    tr = lambda k, col: f"Sens!{col}{lab(SE, k, tor)}"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    rtp = lab(C, "加權目標價（DCF 與 Comparable 加權）"); rrec = lab(C, "投資建議（五級")
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.4 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，比現價 "), 'TEXT(px,"#,##0")', T(" 元"), f'IF({up}>=0,"高 ","低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "),
        f'TEXT({dcf},"#,##0")', T(" 元"), f'IF({dcf}<px,"低於","高於")', T("現價，同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")', T(" 元"), f'IF({cadj}<px,"也低於","也高於")', T("現價；獲利高度取決於 AWS Trainium 量產的顆數與世芯份額。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：現價相當於未來 12 個月前瞻 EPS（情境加權 "), f'TEXT({feps},"0.0")', T(" 元）的 "), f'TEXT(px/{feps},"0.0")', T(" 倍，同業中位數為 "), f'TEXT({pmed},"0.0")',
        T(" 倍；若 AWS Trainium 顆數、世芯份額或 2 奈米每顆營收明顯低於本模型（情境 C），EPS 下修、結論才會改變。只看現金流折現，目標價 "), f'TEXT({dcf},"#,##0")', T(" 元（"), f'TEXT({dcf}/px-1,"+0.0%;-0.0%")', T("），評等"), f'{REC5(f"({dcf}/px-1)")}', T("。"))
    put("o1", T("目標價："), f'TEXT({tp},"#,##0")', T(" 元，比現價 "), 'TEXT(px,"#,##0")', T(" 元"), f'IF({up}>=0,"高 ","低 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec, T("。現金流折現 "), f'TEXT({dcf},"#,##0")',
        T(" 元與同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")', T(" 元各占一半；市場共識 "), 'TEXT(cons_tp,"#,##0")', T(" 元（賣方平均）"), f'IF(cons_tp>{tp},"高於","低於")', T("本模型：本模型前瞻 EPS "), f'IF({geps}>=0,"高於","低於")', T("共識 "), f'TEXT(ABS({geps}),"0%")',
        T("，共識隱含倍數"), f'IF({gmul}>=0,"高於","低於")', T("同業中位數 "), f'TEXT(ABS({gmul}),"0%")', T("。"))
    put("o2", T("核心問題一：FY2027–28 能賺多少。CY2028 世芯 Trainium 顆數三情境 "), f'TEXT(PL!F{bur["C"]},"0.0")', T("／"), f'TEXT(PL!F{bur["B"]},"0.0")', T("／"), f'TEXT(PL!F{bur["A"]},"0.0")', T(" 百萬顆（保守／基準／樂觀）；FY2027 EPS "),
        f'TEXT(PL!E{epr["C"]},"0")', T("～"), f'TEXT(PL!E{epr["A"]},"0")', T(" 元，FY2028 "), f'TEXT(PL!F{epr["C"]},"0")', T("～"), f'TEXT(PL!F{epr["A"]},"0")', T(" 元；情境 C 的 FY2028 EPS 比情境 A 低 "),
        f'TEXT(1-PL!F{epr["C"]}/PL!F{epr["A"]},"0%")', T("。"))
    put("o3", T("核心問題二：差距靠什麼補。份額、顆數與 2 奈米每顆營收皆為假設；加權目標價整本重算：份額 −10／＋10 個百分點使目標價 "), f'TEXT({tr("Trainium 份額 −10", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("Trainium 份額 −10", "C")},"+0%;-0%")',
        T("，顆數 ×0.8／×1.2 使 "), f'TEXT({tr("世芯 Trainium 顆數", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("世芯 Trainium 顆數", "C")},"+0%;-0%")', T("，2 奈米每顆營收 ×0.8／×2.5 使 "), f'TEXT({tr("2 奈米每顆營收", "B")},"+0%;-0%")', T("／"), f'TEXT({tr("2 奈米每顆營收", "C")},"+0%;-0%")', T("。"))
    put("o4", T("反向推算（情境 B）：要支持現價，"), f'IF({rv("w1")}<=g_tv+0.0051,"折現率壓到搜尋下界 "&TEXT({rv("w1")},"0.0%")&" 仍不足","折現率須"&IF({rv("w1")}<=wacc,"降到 ","升到 ")&TEXT({rv("w1")},"0.0%"))', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2027 起營收全面"), f'IF({rv("s1")}>=1,"放大 ","縮小為 ")', f'TEXT({rv("s1")},"0.00")',
        T(" 倍。同業比較法：現價相當於前瞻 EPS 的 "), f'TEXT(RevDCF!B{RVA["pe1"]},"0.0")', T(" 倍，"), f'IF(RevDCF!B{RVA["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍。"))
    put("o5", T("對帳中年份（CY2027–29，Siliconomics 標示；層與層的差距由對帳單處理，非本公司 C0 項目）：另一側讀數目標價 "), f'TEXT(Coord!C{rtp},"#,##0")', T(" 元，評等"), f'Coord!C{rrec}', f'IF(Coord!B{rrec}<>Coord!C{rrec},"（與現行基準讀數"&Coord!B{rrec}&"評等不同）","")',
        T("。公司實績（月營收 1–10 月累計 "), f'TEXT(Coord!B{CR["E_first"]},"0.0")', T(" 十億元）只能對本公司錨點，Trainium 隱含值與錨點同源，C0 無可判定項目。"))
    put("drv", T("分部營收推導（情境 B）：合併營收＝Trainium 量產＋其他（NRE、非 Trainium 量產）。Trainium 量產營收＝AWS Trainium 顆數 × 世芯份額 × 每顆核心營收 × 校準係數（FY2026 由公司 3 奈米營收說法校準）× 匯率；其他營收依成長率。"
                 "FY2028："), T("Trainium "), f'TEXT(PL!F{trr},"0.0")', T("、其他 "), f'TEXT(PL!F{otr},"0.0")', T("，合計營收 "), f'TEXT(PL!F{revr},"0.0")', T(" 十億元。"))
    out["ver"] = r
    RD.cell(r, 1, f"v{VER}（{BUILT}）：只改展示層與新增列——個股 HTML 版面規格 v1.3（章節：命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態）；三層協調（規則 v1.4.2）：Coord 頁新增 A 讀哪一層哪一版、B 讀取層級（同源，不判定）、C 口徑註記、D 摘要、E C0（實績）：公司實績 vs 層級隱含值，"
                  "並列兩讀數目標價（C4；主引擎情境 B、C 於 CY2027 起產業量 × 另一側讀數比值；情境 A 不縮放）。新增 Hist 頁（損益表歷史 FY2022–FY2025）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B，TIFRS）、Val 0.7（四個目標價）、"
                  "Sens 頁末加權目標價相對基準（營運輸入案例為整本重算快照 alchip/sens.json）、README 頁末 v1.3 白話列。E 節 C0（實績）：月營收 1–10 月與 3Q26 對層級隱含值，Trainium 層級隱含值由本模型校準＝參考（同源校準），其餘無可比實績，對帳中項目 0。"
                  "既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價、評等不變。"
                  "註：世芯 3Q26 財報已公布（優分析 2026-11-07：營收 8.193 億美元、EPS 37.68、毛利率低於 20%），本版 FY2026 錨點未更新（仍為 3Q26 月營收實績＋指引），待下版。")
    r += 1
    return out


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT or "Coord" not in wb.sheetnames:   # 另一側讀數建置（alt_run.py）只讀 Val 頁，不加新頁
        return "v024_post：略過（另一側讀數建置或無 Coord 頁）"
    HR = hist_is(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_tor(wb, here)
    CR = coord_c0(wb, SIL)
    RMA = readme13(wb, RVA, PLA, ST, CR)
    RD = wb["README"]
    a1 = RD["A1"].value
    assert f"v{VER}（AI 半導體" in a1, a1[:80]
    RD["A1"].value = a1 + "；版面 v1.3 與 C0（實績）"
    a2 = str(RD["A2"].value)
    assert a2.startswith("版本：v0.1（初版）。"), a2[:40]
    RD["A2"].value = a2.replace("版本：v0.1（初版）。", f"版本：v{VER}（{BUILT}：版面規格 v1.3、三層協調 C0 與兩讀數；前版 v0.2.3，2026-11-06）；v0.1（初版）。", 1)
    order = ["README", "In", "Q", "Hist", "Units", "PL", "Val", "DCF", "RevDCF", "Recon", "Sens", "Sources", "Coord"]
    wb._sheets = [wb[n] for n in order]
    return f"v024_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 相對基準 {ST['first']}–{ST['last']}；Coord E 節 {CR['E_first']}–{CR['E_last']}；README v1.3 {len(RMA)} 列"
