# 記憶體樣板 v13 後處理（SK 海力士 v0.7.6；2026-11-14）：個股 HTML 版面規格 v1.3 與三層協調 C0 對帳所需的 xlsx 新增內容。
# 只在公司設定檔 CONFIG["v13"] 為真時由 tools/memory_model/build.py 呼叫（美光、三星不設，輸出不變）。
# 只新增，不改既有儲存格數值：
#   Hist 頁（新）：損益表歷史（2022 只供年增率；2023–2025 實績）
#   PL 頁（新）：標準損益表（2023A–2025A 實績、2026E–2030E 基準情境 B）、比率區、核對列；三情境並列（圖用）；分部營收推導（情境 B）
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC 與 Comparable 倍數（只讀其他頁）
#   Val 頁末：0.7 估值並列（四個目標價與現價）
#   Sens 頁末：加權目標價相對基準（HTML 用；既有目標價參數案例＋整本重建快照 skhynix/sens_snapshot.json）
#   Coord 頁末：C0 由下往上對帳（三層基準協調規則 v1.4.1 C0）
#   README 頁末：v1.3 展示層白話列（橫幅、投資建議、重點觀察、驅動說明、版本）
# 環境變數 V13_SKIP=1：不做後處理（sens_run.py 建置快照用）。
import json
import os

from openpyxl.styles import Font, PatternFill

BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    """公式字串：parts 為 ("t", 文字) 或 Excel 運算式（字串）。"""
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


def lab(ws, prefix, start=1, col=1, exact=False):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix.strip()) if exact else v.strip().startswith(prefix.strip())):
            return r
    raise KeyError((ws.title, prefix))


def hdr_row(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h)
        c.font = BOLD
        c.fill = HDR


# ---------------- Hist：損益表歷史 ----------------
SRC_H = ("SK hynix 424B4（2026-07；https://www.sec.gov/Archives/edgar/data/2120882/000119312526299963/d32785d424b4.htm）合併損益 2023–2025（營收、營業成本、稅前、所得稅、本期損益）；"
         "營業利益＝毛利 − 銷管 − 研發（424B4 未列營業利益，與公司新聞稿 2024 年 23.4673、2025 年 47.2063 兆相符）；EPS 與歸屬普通股淨利取自 stockanalysis.com/quote/krx/000660/financials/income-statement（2026-11-14 讀取；轉載公司財報，第三方整理）")
HIS = [  # (鍵, 項目, 2023, 2024, 2025, 2022, 標記)；兆韓元、EPS 韓元
    ("rev", "營收", 32.765719, 66.19296, 97.146675, 44.621568, "Verified"),
    ("cogs", "營業成本", 33.299167, 34.364814, 38.455885, None, "Verified"),
    ("gp", "毛利（＝營收 − 營業成本）", "=C{rev}-C{cogs}", "=D{rev}-D{cogs}", "=E{rev}-E{cogs}", None, "Derived"),
    ("oi", "營業利益（毛利 − 銷管費用 − 研發費用）", -7.730313, 23.467319, 47.206319, 6.809417, "Verified"),
    ("opx", "營業費用（銷管＋研發＝毛利 − 營業利益）", "=C{gp}-C{oi}", "=D{gp}-D{oi}", "=E{gp}-E{oi}", None, "Derived"),
    ("pti", "稅前利益", -11.657816, 23.88535, 50.465552, None, "Verified"),
    ("oth", "其他收入（淨額）＝稅前利益 − 營業利益（含財務損益、匯兌、權益法投資）", "=C{pti}-C{oi}", "=D{pti}-D{oi}", "=E{pti}-E{oi}", None, "Derived"),
    ("tax", "所得稅（負值＝稅額利益）", -2.520269, 4.088448, 7.51765, None, "Verified"),
    ("ni", "本期淨利（稅前 − 所得稅；合併，含非控制權益）", "=C{pti}-C{tax}", "=D{pti}-D{tax}", "=E{pti}-E{tax}", None, "Derived"),
    ("nic", "歸屬普通股淨利", -9.112428, 19.788681, 42.919287, 2.22956, "Verified"),
    ("eps", "EPS（完全稀釋；韓元）", -13244.0, 28418.66, 60377.79, None, "Verified"),
]


def hist_sheet(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.7.6 起；PL 頁「標準損益表」的歷史三年與 2023 年增率用）。兆韓元，K-IFRS 合併；EPS 韓元。2022 欄只供 2023 年增率。"
               "只引用本頁常數，不回饋任何計算")
    H["A1"].font = BOLD
    hdr_row(H, 3, ["鍵", "項目", "2023", "2024", "2025", "2022（只供年增率）", "標記", "來源"])
    # 欄：A 鍵、B 項目、C 2023、D 2024、E 2025、F 2022
    rows = {k: 4 + i for i, (k, *_x) in enumerate(HIS)}
    for k, labl, v23, v24, v25, v22, tag in HIS:
        r = rows[k]
        H.cell(r, 1, f"h_{k}")
        H.cell(r, 2, labl)
        for col, v in (("C", v23), ("D", v24), ("E", v25), ("F", v22)):
            if v is None:
                continue
            if isinstance(v, str):
                v = v.format(**rows)
            c = H[f"{col}{r}"]
            c.value = v
            c.number_format = "#,##0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str):
                c.font = BLUE
        H.cell(r, 7, tag)
        H.cell(r, 8, SRC_H if tag == "Verified" else "由上列計算")
    r = 4 + len(HIS) + 1
    H.cell(r, 1, "核對：2025 年營收（本頁）− Q 頁四季營收合計（應接近 0；Q 頁為 3 位小數）")
    H[f"E{r}"] = f"=E{rows['rev']}-SUM(Q!B4:E4)"
    H[f"E{r}"].number_format = "0.0000"
    r += 1
    H.cell(r, 1, "核對：2025 年營業利益（本頁）− Q 頁四季營業利益合計（應接近 0）")
    H[f"E{r}"] = f"=E{rows['oi']}-SUM(Q!B10:E10)"
    H[f"E{r}"].number_format = "0.0000"
    r += 1
    H.cell(r, 1, "核對：2025 年本期淨利（本頁）− Q 頁四季淨利合計（應接近 0）")
    H[f"E{r}"] = f"=E{rows['ni']}-SUM(Q!B12:E12)"
    H[f"E{r}"].number_format = "0.0000"
    H.column_dimensions["A"].width = 14
    H.column_dimensions["B"].width = 64
    for c in "CDEF":
        H.column_dimensions[c].width = 16
    H.column_dimensions["H"].width = 120
    return rows


# ---------------- PL：標準損益表 ----------------
COLS = "CDEFGHIJ"   # 2023A 2024A 2025A 2026E 2027E 2028E 2029E 2030E
YLABEL = ["2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
YB = {"2026": "C", "2027": "G", "2028": "H", "2029": "I", "2030": "P"}   # Y 頁情境 B 欄（2030 為收斂法）


def pl_sheet(wb, HR):
    P = wb.create_sheet("PL", index=wb.sheetnames.index("Y") + 1)
    Y, D, Co = wb["Y"], wb["DCF"], wb["Coord"]
    yr = lambda p: lab(Y, p, exact=True)
    R_ = dict(rev=yr("合併營收"), cogs=yr("營業成本"), opx=yr("營業費用"), oi=yr("合併營業利益"), ni=yr("淨利（合併）"), eps=yr("每股盈餘 EPS（期末股數）"), sh=yr("期末流通股數"))
    d_b = lab(D, "情境 B（")
    d_ebit = lab(D, "營業利益 EBIT", d_b)
    flag = lab(Co, "對帳中旗標")        # Coord 頁 B..E＝CY2026..CY2029
    P["A1"] = ("PL — 標準損益表（v0.7.6 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（2023–2025 實績）＋未來五年（2026–2030，情境 B 基準）；兆韓元，EPS 韓元。"
               "只引用本頁、Hist、Y、DCF、Coord 頁，不回饋任何計算")
    P["A1"].font = BOLD
    P["A2"] = ("口徑：GAAP（K-IFRS 合併財報；SK 海力士只公布 IFRS 損益，無另行定義的 Non-GAAP，市場共識與公司指引也用同一口徑）。"
               "預測：2026E 為基準年（前三季實績＋第四季推估）、2027E–2029E 為情境 B 由驅動因子推導、2030E 為收斂法（只推營收與淨利）。"
               "稅前利益＝淨利 ÷（1 − 有效稅率），其他收入（淨額）＝稅前利益 − 營業利益，所得稅＝稅前 − 淨利（預測不另估財務損益；2026E 的實績季度業外損益含在其他收入內）。"
               "EPS：歷史為公告稀釋 EPS；預測為淨利 ÷ 期末流通股數（Y 頁口徑，已扣除買回註銷），兩者股數基礎不同")
    P["A3"] = ("2030E：毛利率與營業費用率沿用 2029E 推算營業成本、毛利與營業費用；營業利益取 DCF 頁情境 B（由淨利反推）。兩種推法的差額見表末核對列（目標價用延伸期淨利，不受影響）。"
               "對帳中：Siliconomics 標對帳中的日曆年（Coord 頁）在欄位標題標示；2030E 由 2029E 收斂而來，一併標示")
    hr = 5
    hdr = ["項目", "單位"] + [f"=\"FY{y}\"&IF(Coord!{c}{flag}=1,\"（對帳中）\",\"\")" if c else f"FY{y}"
                           for y, c in zip(YLABEL, [None, None, None, None, "C", "D", "E", "E"])] + ["說明"]
    hdr_row(P, hr, hdr)
    # 2026E 欄的對帳中旗標＝Coord B 欄（CY2026＝0）：不標
    r = hr + 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label)
        P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None:
                continue
            P[f"{col}{r}"] = v
            P[f"{col}{r}"].number_format = fmt
        if note:
            P.cell(r, 11, note)
        L[key] = r
        r += 1

    M, PCT = "#,##0.000", "0.0%"
    h = lambda k, c: f"=Hist!{c}{HR[k]}"
    y = lambda k: [f"=Y!{YB[s]}{R_[k]}" for s in ("2026", "2027", "2028", "2029")]
    row("rev", "營收", "兆韓元", [h("rev", "C"), h("rev", "D"), h("rev", "E")] + y("rev") + [f"=Y!{YB['2030']}{R_['rev']}"], M, "預測取 Y 頁情境 B 合併營收（2030E 為收斂法）")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/Hist!F{HR['rev']}-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "2023 年的基期＝2022 年營收（Hist 頁）")
    r_c = r
    row("cogs", "營業成本（COGS）", "兆韓元", [h("cogs", "C"), h("cogs", "D"), h("cogs", "E")] + y("cogs") + [f"=J{L['rev']}*(1-I{r_c + 2})"], M, "2030E＝營收 ×（1 − 2029E 毛利率）")
    row("gp", "毛利", "兆韓元", [f"={c}{L['rev']}-{c}{L['cogs']}" for c in COLS], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "2030E＝2029E（沿用）")
    row("opx", "營業費用（銷管＋研發）", "兆韓元", [h("opx", "C"), h("opx", "D"), h("opx", "E")] + y("opx") + [f"=J{L['rev']}*I{r_c + 3}/I{L['rev']}"], M, "2030E＝營收 × 2029E 營業費用率")
    row("oi", "營業利益", "兆韓元", [h("oi", "C"), h("oi", "D"), h("oi", "E")] + y("oi") + [f"=DCF!F{d_ebit}"], M, "2030E 取 DCF 頁情境 B 的營業利益（由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    r_pti = r + 1
    row("oth", "其他收入（淨額）", "兆韓元", [f"={c}{r_pti}-{c}{L['oi']}" for c in COLS], M, "稅前利益 − 營業利益（歷史含財務損益、匯兌與權益法投資；預測依有效稅率由淨利還原；2026E 含 2Q 投資評價利益與 3Q 匯兌、投資評價損失的淨額）")
    row("pti", "稅前利益", "兆韓元", [h("pti", "C"), h("pti", "D"), h("pti", "E")] + [f"={c}{r_pti + 2}/(1-tax)" for c in COLS[3:]], M, "預測＝淨利 ÷（1 − 有效稅率 tax）")
    row("tax", "所得稅", "兆韓元", [h("tax", "C"), h("tax", "D"), h("tax", "E")] + [f"={c}{r_pti}-{c}{r_pti + 2}" for c in COLS[3:]], M, "歷史為損益表稅額（2023 年為稅額利益）；預測＝稅前 − 淨利（2026E 為依有效稅率推算，非公司公告）")
    row("ni", "稅後淨利", "兆韓元", [h("ni", "C"), h("ni", "D"), h("ni", "E")] + y("ni") + [f"=Y!{YB['2030']}{R_['ni']}"], M, "歷史＝本期淨利（含非控制權益）；預測取 Y 頁情境 B 淨利")
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數（百萬股）", "百萬股", [f"=Hist!{c}{HR['nic']}/Hist!{c}{HR['eps']}*1000000" for c in "CDE"] + y("sh") + [f"=Y!{YB['2030']}{R_['sh']}"], "#,##0.0",
        "歷史＝歸屬普通股淨利 ÷ 稀釋 EPS（稀釋加權平均股數的反推）；預測＝期末流通股數（Y 頁）")
    row("eps", "EPS（完全稀釋）", "韓元", [h("eps", "C"), h("eps", "D"), h("eps", "E")] + y("eps") + [f"=Y!{YB['2030']}{R_['eps']}"], "#,##0", "歷史＝公告稀釋 EPS；預測＝淨利 ÷ 期末流通股數")
    P.cell(r, 1, "比率").font = BOLD
    L["ratio"] = r
    r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["不適用（上年淨利為正、本年虧損，無意義）"] + [f"=IF({p}{L['ni']}<=0,\"不適用（上年虧損）\",{c}{L['ni']}/{p}{L['ni']}-1)" for p, c in zip(COLS, COLS[1:])], PCT,
        "2023 年虧損；2024 年的基期為虧損，兩年皆不適用")
    P.cell(r, 1, "核對（應為 0）").font = BOLD
    r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "兆韓元", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "兆韓元", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None], "0.000")
    row("c2b", "差額（2030E）：毛利 − 營業費用（沿用 2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "兆韓元",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}"], "0.000")
    row("c2c", "差額（2030E）占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}"], PCT)
    row("c3", "核對：稅前利益 − 稅後淨利 − 所得稅", "兆韓元", [f"={c}{L['pti']}-{c}{L['ni']}-{c}{L['tax']}" for c in COLS], "0.000")
    row("c4", "核對：2026E 營收 − Y 頁基準年（合併營收）", "兆韓元", [None, None, None, f"=F{L['rev']}-Y!C{R_['rev']}"] + [None] * 4, "0.000")
    r += 1

    # ---- 三情境並列（圖用；數字取自 Y 頁）----
    P.cell(r, 1, "三情境並列（圖用）：營收、淨利與 EPS（兆韓元；EPS 韓元）；2025 欄＝實績（情境相同）；2030 起為收斂法").font = BOLD
    r += 1
    scen_h = r
    hdr_row(P, r, ["項目", "單位", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E", "2031E", "2032E"])
    r += 1
    scol = {"A": ["D", "E", "F", "M", "N", "O"], "B": ["G", "H", "I", "P", "Q", "R"], "C": ["J", "K", "L", "S", "T", "U"]}
    nm = {"A": "樂觀", "B": "基準", "C": "保守"}
    S = {}
    for key, lbl, src, fmt in (("rev", "營收", R_["rev"], "#,##0.0"), ("ni", "淨利", R_["ni"], "#,##0.0"), ("eps", "EPS", R_["eps"], "#,##0")):
        for s in "ABC":
            P.cell(r, 1, f"{lbl}（情境 {s}　{nm[s]}）")
            P.cell(r, 2, "韓元" if key == "eps" else "兆韓元")
            if key == "rev":
                P[f"C{r}"] = "=Hist!E" + str(HR["rev"])
            elif key == "ni":
                P[f"C{r}"] = "=Hist!E" + str(HR["ni"])
            P[f"D{r}"] = f"=Y!C{src}"
            for j, yc in enumerate(scol[s]):
                P.cell(r, 5 + j, f"=Y!{yc}{src}")
            for j in range(3, 11):
                P.cell(r, j).number_format = fmt
            S[(key, s)] = r
            r += 1
    P.cell(r, 1, "核對：情境 B 2027E–2029E 營收 − 標準損益表（應為 0）")
    for j, c in enumerate("EFG"):
        P[f"{c}{r}"] = f"={c}{S[('rev', 'B')]}-{'GHI'[j]}{L['rev']}"
        P[f"{c}{r}"].number_format = "0.000"
    L["c5"] = r
    r += 2

    # ---- 分部營收推導（情境 B）----
    Mk, Pa = wb["Mkt"], wb["Path"]
    m25, m31, m34 = lab(Mk, "出貨＝MIN(需求, 供給)（B）"), lab(Mk, "HBM 單價（B）"), lab(Mk, "HBM 市場（B）")
    # Path 頁年均指數表：一般 DRAM／NAND 各三列（A、B、C），取 B 列
    pr_d = next(rr for rr in range(13, Pa.max_row + 1) if Pa.cell(rr, 1).value == "一般 DRAM" and Pa.cell(rr, 2).value == "B")
    pr_n = next(rr for rr in range(13, Pa.max_row + 1) if Pa.cell(rr, 1).value == "NAND" and Pa.cell(rr, 2).value == "B")
    P.cell(r, 1, "分部營收推導（情境 B；兆韓元）：合併營收＝一般 DRAM＋HBM＋NAND＋其他＋非記憶體事業；各分部由驅動因子算出（Y 頁同式，本表只引用）").font = BOLD
    r += 1
    seg_h = r
    hdr_row(P, r, ["項目", "單位", "", "2026E", "2027E", "2028E", "2029E", "說明"])
    r += 1
    G = {}
    ycols = ["C", "G", "H", "I"]          # Y 頁 2026、2027B、2028B、2029B
    mcols = ["D", "E", "F", "G"]          # Mkt／Path 頁 2026–2029
    D4 = "DEFG"

    def seg(key, label, unit, fn, fmt, note="", sub=False):
        nonlocal r
        P.cell(r, 1, label)
        if sub:
            P.cell(r, 1).font = BOLD
        P.cell(r, 2, unit)
        for j in range(4):
            v = fn(j)
            if v is None:
                continue
            P[f"{D4[j]}{r}"] = v
            P[f"{D4[j]}{r}"].number_format = fmt
        if note:
            P.cell(r, 8, note)
        G[key] = r
        r += 1

    P.cell(r, 1, "一、一般 DRAM＝前一年 ×（1＋位元成長）× 年均合約價指數比").font = BOLD
    r += 1
    seg("d_bit", "一般 DRAM 位元成長", "%", lambda j: None if j == 0 else f"=cb_{j}", "0.0%", "2027E 起；2026E 為前三季實績＋第四季推估")
    seg("d_idx", "一般 DRAM 年均合約價指數（Path 頁情境 B）", "指數", lambda j: f"=Path!{mcols[j]}{pr_d}", "0.000")
    seg("d_pr", "年均價指數比（÷ 前一年）", "x", lambda j: None if j == 0 else f"={D4[j]}{G['d_idx']}/{D4[j - 1]}{G['d_idx']}", "0.000")
    seg("d_rev", "一般 DRAM 營收", "兆韓元", lambda j: f"=Y!{ycols[j]}{lab(Y, '  一般 DRAM', exact=False)}", M)
    P.cell(r, 1, "二、HBM＝HBM 市場（出貨位元 × 單價）× 本公司營收市占 × 匯率").font = BOLD
    r += 1
    seg("h_bit", "HBM 出貨位元（市場，情境 B）", "EB", lambda j: f"=Mkt!{mcols[j]}{m25}", "0.00", "MIN(需求, 供給)；Siliconomics 基準")
    seg("h_px", "HBM 單價（情境 B）", "US$/GB", lambda j: f"=Mkt!{mcols[j]}{m31}", "0.00")
    seg("h_mkt", "HBM 市場（情境 B）", "US$B", lambda j: f"=Mkt!{mcols[j]}{m34}", "#,##0.0")
    seg("h_sh", "本公司 HBM 營收市占", "%", lambda j: f"=sh{26 + j}", "0.0%", "Siliconomics 市占表")
    seg("h_fx", "匯率", "韓元／美元", lambda j: "=fx", "#,##0")
    seg("h_rev", "HBM 營收", "兆韓元", lambda j: f"=Y!{ycols[j]}{lab(Y, '  HBM', exact=False)}", M, "＝市場 × 市占 × 匯率 ÷ 1000")
    P.cell(r, 1, "三、NAND＝前一年 ×（1＋位元成長）× 年均合約價指數比").font = BOLD
    r += 1
    seg("n_bit", "NAND 位元成長", "%", lambda j: None if j == 0 else f"=nb_{j}", "0.0%")
    seg("n_idx", "NAND 年均合約價指數（Path 頁情境 B）", "指數", lambda j: f"=Path!{mcols[j]}{pr_n}", "0.000")
    seg("n_pr", "年均價指數比（÷ 前一年）", "x", lambda j: None if j == 0 else f"={D4[j]}{G['n_idx']}/{D4[j - 1]}{G['n_idx']}", "0.000")
    seg("n_rev", "NAND 營收", "兆韓元", lambda j: f"=Y!{ycols[j]}{lab(Y, '  NAND', exact=False)}", M)
    P.cell(r, 1, "四、其他與合計").font = BOLD
    r += 1
    seg("o_rev", "其他營收", "兆韓元", lambda j: f"=Y!{ycols[j]}{lab(Y, '  其他', exact=False)}", M, "固定輸入")
    seg("nm_rev", "非記憶體事業營收", "兆韓元", lambda j: f"=Y!{ycols[j]}{yr('非記憶體事業營收')}", M, "本公司沒有，為零")
    seg("tot", "分部加總（營收）", "兆韓元", lambda j: f"={D4[j]}{G['d_rev']}+{D4[j]}{G['h_rev']}+{D4[j]}{G['n_rev']}+{D4[j]}{G['o_rev']}+{D4[j]}{G['nm_rev']}", M, sub=True)
    seg("rep", "合併營收（Y 頁）", "兆韓元", lambda j: f"=Y!{ycols[j]}{R_['rev']}", M)
    seg("gap", "核對：分部加總 − 合併營收（應為 0）", "兆韓元", lambda j: f"={D4[j]}{G['tot']}-{D4[j]}{G['rep']}", "0.000")
    seg("hsh", "HBM ÷ 營收", "%", lambda j: f"={D4[j]}{G['h_rev']}/{D4[j]}{G['rep']}", "0.0%")
    seg("dsh", "一般 DRAM ÷ 營收", "%", lambda j: f"={D4[j]}{G['d_rev']}/{D4[j]}{G['rep']}", "0.0%")
    seg("nsh", "NAND ÷ 營收", "%", lambda j: f"={D4[j]}{G['n_rev']}/{D4[j]}{G['rep']}", "0.0%")
    P.column_dimensions["A"].width = 60
    P.column_dimensions["B"].width = 10
    for c in "CDEFGHIJ":
        P.column_dimensions[c].width = 14
    P.column_dimensions["K"].width = 90
    L.update(hdr=hr, scen_h=scen_h, seg_h=seg_h)
    return L, S, G


# ---------------- RevDCF：反向 DCF ----------------
def revdcf(wb):
    D, V = wb["DCF"], wb["Val"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（")
    dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    vb = lambda p: f"Val!$B${lab(V, p)}"
    fbv, pmed, dcf_tp, coef = vb("12 個月後的前瞻每股淨值（情境加權）"), vb("同業前瞻 P/B 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.7.6 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻每股淨值（情境加權）× P/B 倍數＝現價」。2027E 為價格已議定的錨點年，不隨營業利益率或營收倍數變動；"
                "反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD
    r += 1
    hdr_row(RV, r, ["項目", "值", "說明"])
    r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label)
        RV.cell(r, 2, f).number_format = fmt
        if note:
            RV.cell(r, 3, note)
        A[key] = r
        r += 1

    kv("px", "現價（韓元）", "=px", "#,##0")
    kv("v0", "支持現價所需的評價日每股價值（韓元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（兆韓元；Val 頁）", f"={ncr}", "#,##0.0", "同 DCF 頁每股價值公式（淨現金與投資＋stub 期 FCFF − stub 期股利買回）")
    kv("sh", "評價用稀釋股數（兆股）", "=sh_v", "0.000000")
    kv("evr", "支持現價所需的企業價值 EV（兆韓元）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；兆韓元）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 須調整為模型的倍數（小於 1 為縮小）", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（韓元）", f"=DCF!$C${ps}", "#,##0")
    r += 1
    yh = r
    hdr_row(RV, r, ["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)])
    r += 1
    Y = {}

    def yr_(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            RV.cell(r, 2 + i, fn(dc, "BCDEFG"[i])).number_format = fmt
        Y[key] = r
        r += 1

    yr_("rev", "營收（兆韓元）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr_("ebit", "營業利益 EBIT（兆韓元）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr_("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr_("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr_("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr_("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr_("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）",
        lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr_("k", "每 1 兆韓元營業利益對 EV 的貢獻（2027E 錨點不動，記 0）",
        lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr_("cf", "NOPAT＋折舊 − 資本支出（兆韓元）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr_("wc", "營運資金增加（兆韓元）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr_("pv", "現值（兆韓元，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（2028E–2032E 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD
    r += 1
    kv("S", "營收 × 貢獻係數合計（兆韓元）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（兆韓元）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（2028E–2032E）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（2028E–2032E 營收等比例調整；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD
    r += 1
    kv("a", "不隨營收倍數變動的 EV（兆韓元）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "2027E 現值＋2028E 營運資金增加中屬 2027E 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（兆韓元）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 2028E 營收（兆韓元）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：2028E 營收（情境 B；兆韓元）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 2028E 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：2028E 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD
    r += 1
    hdr_row(RV, r, ["次", "下界", "上界", "中點", "中點的評價日每股價值（韓元）"])
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C$8+$B${A['nc']})/sh_v")

    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005")
            RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})")
            RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2")
        RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4):
            RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "#,##0"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（韓元，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
    r += 1
    RV.cell(r, 1, "B. Comparable：支持現價所需的倍數（Comparable＝同業前瞻 P/B 中位數 × 相對倍數調整係數 × 12 個月後前瞻每股淨值）").font = BOLD
    r += 1
    hdr_row(RV, r, ["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"])
    r += 1

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        RV.cell(r, 2, f_imp).number_format = fmt
        RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r
        r += 1

    RV.cell(r, 1, "12 個月後前瞻每股淨值（情境加權；韓元）")
    RV.cell(r, 3, f"={fbv}").number_format = "#,##0"
    A["fbv"] = r
    r += 1
    cmp("pb1", "只看 Comparable：支持現價所需的前瞻 P/B（倍）", f"=px/{fbv}", f"=tp_comp_adj/{fbv}", "0.00")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻 P/B 中位數）", f"=B{A['pb1']}/{pmed}", f"={coef}", "0.00")
    cmp("pb2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻 P/B（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{fbv})", f"=tp_comp_adj/{fbv}", "0.00")
    RV.cell(r, 1, "同業前瞻 P/B 中位數（倍）")
    RV.cell(r, 3, f"={pmed}").number_format = "0.00"
    A["pmed"] = r
    r += 2
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD
    r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"#,##0")',
                     T(" 韓元；要支持現價 "), 'TEXT(px,"#,##0")', T(" 韓元，企業價值須"), f'IF(B{A["evx"]}<1,"縮小為 ","放大為 ")&TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("；或 2028E–2032E 營收全面"),
                     f'IF(B{A["s1"]}<1,"縮小為 ","放大為 ")&TEXT(B{A["s1"]},"0.00")', T(" 倍（2028E 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 兆韓元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                     T("）；或營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r
    r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"#,##0")', T(" 韓元相當於 12 個月後前瞻每股淨值的 "), f'TEXT(B{A["pb1"]},"0.00")',
                     T(" 倍，同業前瞻 P/B 中位數為 "), f'TEXT({pmed},"0.00")', T(" 倍；模型採用 "), f'TEXT(C{A["pb1"]},"0.00")',
                     T(" 倍（同業中位數 × 本公司十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給的 P/B 只有同業中位數的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，模型採用 "), f'TEXT({coef},"0.00")', T(" 倍。")))
    A["sum_cmp"] = r
    r += 1
    RV.column_dimensions["A"].width = 70
    for c in "BCDEFGH":
        RV.column_dimensions[c].width = 16
    A["yh"] = yh
    A["bis0"] = b0
    A["ys"] = Y
    return A


# ---------------- Val 0.7：估值並列 ----------------
def val07(wb):
    V = wb["Val"]
    r = V.max_row + 2
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價（v0.7.6 起；個股 HTML 版面規格 v1.3 G 節；只引用本頁與 In 頁，不進任何計算）").font = BOLD
    r0 = r
    r += 1
    hdr_row(V, r, ["項目", "目標價（韓元）", "相對現價", "說明"])
    r += 1
    first = r
    rows = [
        ("DCF 目標價（情境加權）", f"=B{lab(V, 'DCF 目標價（情境加權）')}", "三情境 DCF 價值推到 12 個月後的加權（第 0.1 節）"),
        ("Comparable 目標價（相對倍數調整後；進加權）", "=tp_comp_adj", "同業前瞻 P/B 中位數 × 本公司相對倍數調整係數 × 12 個月後前瞻每股淨值（第 0.5 節）；加權目標價使用本列"),
        ("Comparable 目標價（市場主流同業法；未調整，不進加權）", f"=B{lab(V, 'Comparable 目標價（市場主流同業法）')}", "同業前瞻 P/B 中位數 × 前瞻每股淨值（第 0.2 節），未做相對倍數調整"),
        ("加權目標價（DCF 與 Comparable 加權；橫幅主值）", f"=B{lab(V, '目標價（DCF 與 Comparable 加權）')}", "DCF 權重 × DCF ＋（1 − DCF 權重）× Comparable（調整後）；投資建議依本列"),
        ("市場共識目標價（賣方平均）", f"=B{lab(V, '市場共識目標價')}", f"=E{lab(V, '市場共識目標價')}"),
    ]
    for labl, f, note in rows:
        V.cell(r, 1, labl)
        V.cell(r, 2, f).number_format = "#,##0"
        V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"
        V.cell(r, 4, note)
        r += 1
    V.cell(r, 1, "現價")
    V.cell(r, 1).value = "=In!B6"
    V.cell(r, 2, "=px").number_format = "#,##0"
    V.cell(r, 4, "In 頁 px（價格日見項目名稱）")
    r += 1
    return {"t": r0, "first": first}


# ---------------- Sens：加權目標價相對基準 ----------------
PLAIN_P = {   # Sens 目標價參數表的鍵 → 白話標籤
    "wacc_adj": "折現率（WACC）低／高一個百分點",
    "beta": "產業 β 低／高",
    "erp": "股權風險溢酬低／高",
    "g_tv": "永續成長率低／高一個百分點",
    "ronic_sp": "永續期新投資報酬率高於折現率五個百分點（只有高）",
    "fc_cx_d": "資本支出占營收低／高一個百分點",
    "fc_wc": "營運資金占營收低／高",
    "inv_k": "淨現金不含權益投資（只有低）",
    "tp_w_dcf": "DCF 權重調低／調高",
}
USD_LABEL = "美元口徑折現（rf 改用美國公債；只有低）"
ALT_LABEL = "終值改用最後一年自由現金流成長（只有高）"


def sens_rel(wb, here):
    S = wb["Sens"]
    snap_f = os.path.join(here, "sens_snapshot.json")
    snap = json.load(open(snap_f, encoding="utf-8")) if os.path.exists(snap_f) else None
    t0 = lab(S, "目標價參數敏感度")
    hd = t0 + 1
    r_end = lab(S, "檢查：閉式公式的基準") - 1
    r = S.max_row + 2
    S.cell(r, 1, "加權目標價相對基準（HTML 用；v0.7.6 起；個股 HTML 版面規格 v1.3 H 節）：一次只動一個輸入，輸入取低值、取高值各一欄；"
                 "前段＝上方目標價參數（閉式，引用上表），後段＝業務與市場輸入（整本重建快照，見下方快照狀態）").font = BOLD
    top = r
    r += 1
    hdr_row(S, r, ["輸入（白話）", "輸入取低值：加權目標價相對基準", "輸入取高值：加權目標價相對基準", "低：加權目標價", "高：加權目標價", "最大變動（絕對值）", "方法"])
    r += 1
    first = r
    for rr in range(hd + 1, r_end + 1):
        key = S.cell(rr, 8).value
        a = str(S.cell(rr, 1).value or "")
        if key in PLAIN_P:
            labl = PLAIN_P[key]
        elif a.startswith("美元口徑折現"):
            labl = USD_LABEL
        elif a.startswith("終值改用"):
            labl = ALT_LABEL
        else:
            raise SystemExit(f"Sens 目標價參數列未對應白話標籤：{a!r}")
        S.cell(r, 1, labl)
        S.cell(r, 2, f'=IF(ISNUMBER(D{rr}),D{rr},"")')
        S.cell(r, 3, f'=IF(ISNUMBER(F{rr}),F{rr},"")')
        S.cell(r, 4, f'=IF(ISNUMBER(C{rr}),C{rr},"")')
        S.cell(r, 5, f'=IF(ISNUMBER(E{rr}),E{rr},"")')
        S.cell(r, 7, "目標價參數（閉式，引用上方目標價參數敏感度表）")
        r += 1
    snap_first = r
    vb = lab(wb["Val"], "目標價（DCF 與 Comparable 加權）")
    if snap:
        for c in snap["cases"]:
            S.cell(r, 1, c["label"])
            lo, hi = c.get("lo"), c.get("hi")
            if lo is not None:
                S.cell(r, 4, lo).font = BLUE
                S.cell(r, 2, f"=D{r}/Val!$B${vb}-1")
            if hi is not None:
                S.cell(r, 5, hi).font = BLUE
                S.cell(r, 3, f"=E{r}/Val!$B${vb}-1")
            S.cell(r, 7, "整本重建快照（In 頁該輸入取敏感度低／高，其餘不變，LibreOffice 重算後讀加權目標價）：" + c["keys"])
            r += 1
    last = r - 1
    for rr in range(first, last + 1):
        for c in "BC":
            S[f"{c}{rr}"].number_format = "+0.0%;-0.0%"
        for c in "DE":
            S[f"{c}{rr}"].number_format = "#,##0"
        S[f"F{rr}"] = f'=IF(COUNT(B{rr}:C{rr})=0,"",MAX(ABS(N(B{rr})),ABS(N(C{rr}))))'
        S[f"F{rr}"].number_format = "0.0%"
    state = r
    S.cell(r, 1, "快照狀態")
    if snap:
        S.cell(r, 2, f'=IF(ABS(Val!$B${vb}-{snap["base_tp"]!r})<=1E-6*ABS({snap["base_tp"]!r}),"CURRENT","STALE")')
        S.cell(r, 3, f"快照 {snap['snapshot']}；建置時基準加權目標價 {snap['base_tp']:,.0f} 韓元；與 Val 頁不同即 STALE（改輸入後須重跑 skhynix/sens_run.py）")
    else:
        S.cell(r, 2, "無快照")
    r += 1
    S.cell(r, 1, "最大變動的輸入（白話）")
    S.cell(r, 2, f"=INDEX(A{first}:A{last},MATCH(MAX(F{first}:F{last}),F{first}:F{last},0))")
    S.cell(r, 3, f"=MAX(F{first}:F{last})").number_format = "0.0%"
    topc = r
    r += 1
    S.cell(r, 1, "核對：基準加權目標價（快照建置時）− Val 頁（應為 0）")
    if snap:
        S.cell(r, 2, f"=Val!$B${vb}-{snap['base_tp']!r}").number_format = "0.000000"
    S.column_dimensions["G"].width = 40
    return {"title": top, "first": first, "last": last, "snap_first": snap_first, "state": state, "top": topc}


# ---------------- Coord：C0 對帳 ----------------
def coord_c0(wb, SIL, here=None):
    Co, Mk, Y, V = wb["Coord"], wb["Mkt"], wb["Y"], wb["Val"]
    m = lambda p: lab(Mk, p)
    m25, m31, m34, m22, m17, m19 = (m("出貨＝MIN(需求, 供給)（B）"), m("HBM 單價（B）"), m("HBM 市場（B）"), m("需求 × 出貨係數（B）"),
                                    m("供給（三情境相同）"), m("顆數係數（B；"))
    m42 = lab(Mk, "已發布：SK 海力士自估")
    flag = lab(Co, "對帳中旗標")
    ratio = lab(Co, "另一側讀數 ÷ 現行基準")
    N = SIL.NAMES
    idx = {2026: 1, 2027: 2, 2028: 3, 2029: 4}   # SIF 數列索引（2025 起）
    cols = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
    mcol = {2026: "D", 2027: "E", 2028: "F", 2029: "G"}
    ccol = {2026: "B", 2027: "C", 2028: "D", 2029: "E"}   # Coord 上半部年度欄
    ycol = {2026: "C", 2027: "G", 2028: "H", 2029: "I"}
    r = Co.max_row + 2
    Co.cell(r, 1, ("C0 由下往上對帳：讀取層級（同源，不判定；三層基準協調規則 v1.4.2 C0 第 2 點，v0.7.7 改標）：本模型自己的產業相關量（HBM 市場規模與出貨位元、基準顆數係數 K1、HBM 營收市占、HBM 營收），"
                   f"與 Siliconomics {SIL.VERSION} 對應值並列。本模型直接讀層，差距恆為 0、無檢驗力，故不判定、不計入對帳中；有檢驗力的對帳見本頁末「C0（實績）」節（公司實績 vs 層級隱含值）。"
                   "層的值為建置時由 siliconomics/reader.py 讀入的常數（藍字），本模型的值為活公式")).font = BOLD
    top = r
    r += 1
    Co.cell(r, 1, "容差").font = BOLD
    Co.cell(r, 2, 0.15).number_format = "0%"
    tol = r
    r += 1
    hdr_row(Co, r, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "口徑、來源與證據等級"])
    hr = r
    r += 1
    L = {}

    def three(key, name, unit, own_f, layer_vals, fmt, own_note, layer_note, grade, state=True):
        """本模型一列、層一列、差距一列、狀態一列。"""
        nonlocal r
        Co.cell(r, 1, f"{name}：本模型")
        Co.cell(r, 2, unit)
        for y in cols:
            Co.cell(r, ord(cols[y]) - 64, own_f(y)).number_format = fmt
        Co.cell(r, 7, own_note)
        ro = r
        r += 1
        Co.cell(r, 1, f"{name}：層（Siliconomics）")
        Co.cell(r, 2, unit)
        for y in cols:
            v = layer_vals(y)
            c = Co.cell(r, ord(cols[y]) - 64, v)
            c.number_format = fmt
            if not isinstance(v, str):
                c.font = BLUE
        Co.cell(r, 7, layer_note + "；證據等級：" + grade)
        rl = r
        r += 1
        Co.cell(r, 1, f"{name}：差距（本模型 ÷ 層 − 1）")
        Co.cell(r, 2, "%")
        for y in cols:
            cc = cols[y]
            Co[f"{cc}{r}"] = f'=IF({cc}{rl}=0,"",{cc}{ro}/{cc}{rl}-1)'
            Co[f"{cc}{r}"].number_format = "+0.0%;-0.0%"
        rd = r
        r += 1
        if state:
            Co.cell(r, 1, f"{name}：狀態")
            for y in cols:
                cc = cols[y]
                Co[f"{cc}{r}"] = f'=IF({cc}{rd}="","","讀取層級（同源，不判定）")'
            L[key + "_s"] = r
            r += 1
        L[key] = (ro, rl, rd)

    sil = lambda nm: (lambda y: N[nm][idx[y]])
    three("mkt", "HBM 市場規模（情境 B）", "US$B", lambda y: f"=Mkt!{mcol[y]}{m34}", sil("SIF_HBM_Mkt_Base"), "#,##0.0",
          "Mkt 頁情境 B（出貨位元 × 單價）；本模型讀層，此列應為一致", f"SIF_HBM_Mkt_Base（{SIL.FILE}）", "層為 Assumed／Derived（Wattonomics 約束上的外推，丙狀態）")
    three("bit", "HBM 出貨位元（情境 B）", "EB", lambda y: f"=Mkt!{mcol[y]}{m25}", sil("SIF_HBM_ShipEB_Base"), "0.00",
          "Mkt 頁情境 B：MIN(需求 × K1, 供給)", "SIF_HBM_ShipEB_Base", "同上")
    three("k1", "基準顆數係數 K1（情境 B；÷ 原始顆數）", "x", lambda y: f"=Mkt!{mcol[y]}{m19}", sil("SIF_UnitsCoef_Base"), "0.000",
          "Mkt 頁情境 B 顆數係數（2026–2027 無情境差異）", "SIF_UnitsCoef_Base（三層基準協調規則 v1.4：基準係數）", "同上")
    three("sh", "SK 海力士 HBM 營收市占", "%", lambda y: f"=sh{y - 2000}", lambda y: SIL.V[f"hbmsh_skh_{y - 2000}"]["base"], "0.0%",
          "In 頁 sh26–sh29（讀層）", "Siliconomics M24 市占表 hbmsh_skh", "層為 Interested-party（Counterpoint、IDC、UBS 位元市占）")
    three("rev", "SK 海力士 HBM 營收", "US$B", lambda y: f"=Y!{ycol[y]}{lab(Y, '  HBM')}/(fx/1000)",
          lambda y: f"={cols[y]}{L['mkt'][1]}*{cols[y]}{L['sh'][1]}", "#,##0.0",
          "Y 頁情境 B HBM 營收 ÷ 匯率（韓元兆 → US$B）", "層：HBM 市場 × 層的市占（Σ 個股加總的一部分）", "由上兩列相乘（Derived）")
    # 恆等式：已建模個股加總＋殘差＝市場
    Co.cell(r, 1, "C0 恆等式：HBM 市占合計（SK 海力士＋三星＋美光＋未建模殘差）＝1")
    Co.cell(r, 2, "%")
    for y in cols:
        v = f"={cols[y]}{L['sh'][1]}+{SIL.V['hbmsh_ss_%d' % (y - 2000)]['base']}+{SIL.V['hbmsh_mu_%d' % (y - 2000)]['base']}"
        Co[f"{cols[y]}{r}"] = v
        Co[f"{cols[y]}{r}"].number_format = "0.0%"
        Co[f"{cols[y]}{r}"].font = BLUE
    Co.cell(r, 7, "SK 海力士取上方層列（連到格）加層的三星 hbmsh_ss、美光 hbmsh_mu（Siliconomics 常數）合計；殘差＝1 − 合計，層另列 SIF_HBMShareSum")
    L["shsum"] = r
    r += 1
    Co.cell(r, 1, "C0 恆等式：未建模殘差（1 − 三家合計）")
    Co.cell(r, 2, "%")
    for y in cols:
        Co[f"{cols[y]}{r}"] = f"=1-{cols[y]}{L['shsum']}"
        Co[f"{cols[y]}{r}"].number_format = "0.0%"
    Co.cell(r, 7, "殘差明列、不為空時不得默默吸收（C0 第 1 點）；2026–2029 為 0 表示三家已涵蓋層的 HBM 市場（SK 海力士、三星、美光）")
    L["resid"] = r
    r += 1
    # 外部證據
    Co.cell(r, 1, "對外證據：HBM 市場規模（US$B）")
    Co.cell(r, 7, "獨立來源的 HBM 市場估計與層的 HBM 市場（情境 B）比較；賣方與公司自估為 Interested-party，非逐顆證據")
    Co.cell(r, 1).font = BOLD
    r += 1
    ev = [("SK 海力士自估 2026（Mkt 頁）", 2026, "D"), ("BofA 2026（Mkt 頁）", 2026, "E"), ("BofA 2027（Mkt 頁）", 2027, "F")]
    L["ev"] = []
    for name, y, c in ev:
        Co.cell(r, 1, name)
        Co.cell(r, 2, "US$B")
        Co[f"{cols[y]}{r}"] = f"=Mkt!{c}{m42}"
        Co[f"{cols[y]}{r}"].number_format = "#,##0.0"
        r += 1
        Co.cell(r, 1, f"  差距：層 ÷ {name} − 1")
        Co.cell(r, 2, "%")
        Co[f"{cols[y]}{r}"] = f"={cols[y]}{L['mkt'][1]}/{cols[y]}{r - 1}-1"
        Co[f"{cols[y]}{r}"].number_format = "+0.0%;-0.0%"
        Co[f"{cols[y]}{r + 1}"] = f'=IF(ABS({cols[y]}{r})>$B${tol},"參考：超出容差（外部估計，不判定、不開單）","參考：在容差內（外部估計，不判定）")'
        Co.cell(r + 1, 1, f"  狀態：{name}")
        L["ev"].append((r - 1, r, r + 1))
        r += 2
    # 摘要
    Co.cell(r, 1, "C0 摘要（HTML 用；公式帶數字）")
    n_ev = "+".join(f'COUNTIF(C{x[2]}:F{x[2]},"*超出*")' for x in L["ev"])
    Co.cell(r, 2, F(T("讀取層級（同源，不判定、不計入對帳中）：HBM 市場規模、出貨位元、顆數係數、市占與 HBM 營收本模型直接讀層，差距恆為 0，無檢驗力；對帳以本頁末「C0（實績）」節的公司實績對層級隱含值為準。"),
                    T("對外證據（參考，不判定）："), f'IF(({n_ev})=0,"三筆外部估計都在容差內","有 "&({n_ev})&" 筆外部估計與層差距超過 ±15%（見下表，皆為賣方或公司自估，非逐顆證據）")'))
    L["sum"] = r
    r += 1
    # 另一側讀數
    r += 1
    Co.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：另一側讀數（Wattonomics 換算的產業 GW）對 HBM 出貨位元的影響（情境 B；情境 A 不縮放）").font = BOLD
    Co.cell(r, 7, "這是 Wattonomics 與 Siliconomics 兩層之間的差距，不是本公司對它讀的層值的差距；本公司只並列兩讀數，不裁決、不以 ±15% 判定本公司對帳中")
    L["lay"] = r
    r += 1
    Co.cell(r, 1, "另一側讀數 ÷ 現行基準（上方 Coord 頁；只在 Siliconomics 旗標年份套用，其他年份＝1）")
    Co.cell(r, 2, "x")
    for y in cols:
        Co[f"{cols[y]}{r}"] = f"=IF({ccol[y]}{flag}=1,{ccol[y]}{ratio},1)"
        Co[f"{cols[y]}{r}"].number_format = "0.000"
    L["ratio"] = r
    r += 1
    Co.cell(r, 1, "情境 B 需求位元（× K1，Mkt 頁）")
    Co.cell(r, 2, "EB")
    for y in cols:
        Co[f"{cols[y]}{r}"] = f"=Mkt!{mcol[y]}{m22}"
        Co[f"{cols[y]}{r}"].number_format = "0.00"
    L["dem"] = r
    r += 1
    Co.cell(r, 1, "情境 B 供給位元（三情境相同，Mkt 頁）")
    Co.cell(r, 2, "EB")
    for y in cols:
        Co[f"{cols[y]}{r}"] = f"=Mkt!{mcol[y]}{m17}"
        Co[f"{cols[y]}{r}"].number_format = "0.00"
    L["sup"] = r
    r += 1
    Co.cell(r, 1, "現行基準讀數：情境 B 出貨位元")
    Co.cell(r, 2, "EB")
    for y in cols:
        Co[f"{cols[y]}{r}"] = f"=Mkt!{mcol[y]}{m25}"
        Co[f"{cols[y]}{r}"].number_format = "0.00"
    L["cur"] = r
    r += 1
    Co.cell(r, 1, "另一側讀數：情境 B 出貨位元＝MIN(需求 × 比值, 供給)")
    Co.cell(r, 2, "EB")
    for y in cols:
        cc = cols[y]
        Co[f"{cc}{r}"] = f"=MIN({cc}{L['dem']}*{cc}{L['ratio']},{cc}{L['sup']})"
        Co[f"{cc}{r}"].number_format = "0.00"
    Co.cell(r, 7, "與整本重建的另一側讀數同式（主引擎掛鉤：情境 B 自 CY2027 起產業量 × 比值；供給不縮放）")
    L["alt"] = r
    r += 1
    Co.cell(r, 1, "另一側讀數 ÷ 現行基準 − 1（出貨位元）")
    Co.cell(r, 2, "%")
    for y in cols:
        cc = cols[y]
        Co[f"{cc}{r}"] = f"={cc}{L['alt']}/{cc}{L['cur']}-1"
        Co[f"{cc}{r}"].number_format = "+0.0%;-0.0%"
    L["altd"] = r
    r += 1
    Co.cell(r, 1, "層級之間：Siliconomics 對帳中旗標")
    for y in cols:
        cc = cols[y]
        Co[f"{cc}{r}"] = f'=IF({ccol[y]}{flag}=1,"對帳中（Siliconomics 旗標）","—")'
    Co.cell(r, 7, "取自上方 Coord 頁旗標列；對帳單由主控開立（Siliconomics 對帳單 r3），期限見上方目標價區說明")
    L["alt_s"] = r
    r += 1
    Co.cell(r, 1, "另一側讀數：情境 B HBM 市場（出貨位元 × 單價，Mkt 頁）")
    Co.cell(r, 2, "US$B")
    for y in cols:
        cc = cols[y]
        Co[f"{cc}{r}"] = f"={cc}{L['alt']}*Mkt!{mcol[y]}{m31}"
        Co[f"{cc}{r}"].number_format = "#,##0.0"
    L["altmkt"] = r
    r += 1
    Co.cell(r, 1, "層級之間摘要（HTML 用；公式帶數字）")
    Co.cell(r, 2, F(T("另一側讀數（Wattonomics 換算）使情境 B 的 HBM 出貨位元 CY2027 由 "), f'TEXT(D{L["cur"]},"0.00")', T(" 降到 "), f'TEXT(D{L["alt"]},"0.00")',
                    T(" EB（"), f'TEXT(D{L["altd"]},"+0%;-0%")', T("），CY2028 "), f'TEXT(E{L["altd"]},"+0%;-0%")', T("，CY2029 "), f'TEXT(F{L["altd"]},"+0%;-0%")',
                    T("。這是兩層之間的差距（Siliconomics 對帳單 r3 處理），不計入本公司對帳；個股兩讀數並列、不改上游值。")))
    L["lay_sum"] = r
    r += 1
    Co.column_dimensions["G"].width = 100
    Co.column_dimensions["F"].width = 22
    L.update(top=top, hdr=hr, tol=tol)
    # C0（實績）：公司實績 vs 層級隱含值（公司資料夾有 c0_actuals.py 才做；附在 Coord 頁尾，不位移既有列）
    import importlib.util
    _py = os.path.join(here, "c0_actuals.py") if here else None
    if _py and os.path.exists(_py):
        spec = importlib.util.spec_from_file_location("c0_actuals_co", _py)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        L["act"] = mod.add(wb, Co, L, F, T)
    return L


# ---------------- README：v1.3 展示層 ----------------
def readme13(wb, PLA, S_, G_, RVA, SNA, C0, ver_text):
    RD, V, P, Co, SE = wb["README"], wb["Val"], wb["PL"], wb["Coord"], wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cu = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）")
    ca = "tp_comp_adj"
    cons = vb("市場共識目標價")
    c_tp, c_rec, c_up = (f"Coord!C{lab(Co, '加權目標價')}", f"Coord!C{lab(Co, '投資建議（五級')}", f"Coord!C{lab(Co, '上檔空間（加權目標價')}")
    b_rec = f"Coord!B{lab(Co, '投資建議（五級')}"
    yrs = f"Coord!B{lab(Co, '對帳中年份')}"
    rr = lab(Co, "另一側讀數 ÷ 現行基準")
    REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
    # 情境淨利列（PL 頁三情境並列）
    ni = {s: S_[("ni", s)] for s in "ABC"}
    rev = {s: S_[("rev", s)] for s in "ABC"}
    # PL 頁欄：C 2025A D 2026E E 2027E F 2028E G 2029E H 2030E I 2031E J 2032E
    dcfw = "tp_w_dcf"
    dcsh = f"({dcfw}*({dcf}-px))/(({dcfw}*({dcf}-px))+((1-{dcfw})*({ca}-px)))"
    sn = SNA
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；個股 HTML 版面規格 v1.3：橫幅 ≤200 字、重點觀察 ≤5 點且 ≤750 字；公式帶數字，v0.7.6 起；上方各列保留給下游）").font = BOLD
    r0 = r
    rows = {}
    r += 1
    rows["hero"] = r
    RD.cell(r, 1, F(T("投資建議"), rec, T("：加權目標價 "), f'TEXT({tp},"#,##0")', T(" 韓元，較現價高 "), f'TEXT({up},"0.0%")',
                    T("；DCF 與同業 P/B 兩法都"), f'IF(AND({dcf}>px,{ca}>px),"高於","並非都高於")', T("現價。對帳中（"), yrs, T("）：另一側讀數目標價 "),
                    f'TEXT({c_tp},"#,##0")', T("（"), c_rec, T("）"), f'IF({b_rec}<>{c_rec},"，與現行基準讀數（"&{b_rec}&"）評等不同","")', T("。")))
    r += 1
    rows["rec"] = r
    RD.cell(r, 1, F(T("投資建議依加權目標價的上檔空間固定規則：大於 50% 強力買進、25%～50% 買進、−25%～25% 中立、−50%～−25% 放空、小於 −50% 強力放空。目前上檔空間 "),
                    f'TEXT({up},"0.0%")', T("，評等"), rec, T("；"),
                    f'IF({up}>0.5,"目標價再下修 "&TEXT(1-px*1.5/{tp},"0.0%")&" 就會降為買進（上檔空間降到 50%）",IF({up}>0.25,"目標價再上修 "&TEXT(px*1.5/{tp}-1,"0.0%")&" 就會升為強力買進","距相鄰評等邊界不遠"))',
                    T("。對帳中的另一側讀數上檔空間 "), f'TEXT({c_up},"0.0%")', T("，評等"), c_rec, T("。")))
    r += 1
    rows["o1"] = r
    RD.cell(r, 1, F(T("合理價值高於現價：加權目標價 "), f'TEXT({tp},"#,##0")', T(" 韓元（DCF "), f'TEXT({dcf},"#,##0")', T("、同業 P/B 法〔相對倍數調整後〕"), f'TEXT({ca},"#,##0")',
                    T("），較現價高 "), f'TEXT({up},"0.0%")', T("；上檔由 DCF 與同業比較約各半貢獻（DCF "), f'TEXT({dcsh},"0%")',
                    T("）。賣方共識目標價 "), f'TEXT({cons},"#,##0")', T("，本模型較共識"), f'IF({tp}<{cons},"低 ","高 ")&TEXT(ABS({tp}/{cons}-1),"0%")', T("。")))
    r += 1
    rows["o2"] = r
    RD.cell(r, 1, F(T("盈餘高點看情境、不是單一預測：基準情境淨利 2027 年 "), f'TEXT(PL!E{ni["B"]},"#,##0")', T(" 兆韓元，2028 年降到 "), f'TEXT(PL!F{ni["B"]},"#,##0")',
                    T("（"), f'TEXT(PL!F{ni["B"]}/PL!E{ni["B"]}-1,"+0%;-0%")', T("）；保守情境 2028 年 "), f'TEXT(PL!F{ni["C"]},"#,##0")', T("（"), f'TEXT(PL!F{ni["C"]}/PL!E{ni["C"]}-1,"+0%;-0%")',
                    T("）；樂觀情境高檔延到 2029 年（"), f'TEXT(PL!G{ni["A"]},"#,##0")', T("）。分歧集中在 2028 年以後。")))
    r += 1
    rows["o3"] = r
    g = lambda k, c="E": f"PL!{c}{G_[k]}"
    sg = lab(SE, "2027 A")
    RD.cell(r, 1, F(T("2027 年基準情境營收 "), f'TEXT({g("rep")},"#,##0")', T(" 兆韓元：一般 DRAM 占 "), f'TEXT({g("dsh")},"0%")', T("、NAND 占 "), f'TEXT({g("nsh")},"0%")', T("、HBM 占 "),
                    f'TEXT({g("hsh")},"0%")', T("。營業利益對一般 DRAM 價格最敏感：2027 年樂觀情境 DRAM 收入 +20% 使營業利益 "), f'TEXT(Sens!E{sg + 2},"+0.0%;-0.0%")',
                    T("，HBM 收入 +30% 只 "), f'TEXT(Sens!E{sg + 4},"+0.0%;-0.0%")', T("；HBM 年約鎖價，收入較穩、金額較小。")))
    r += 1
    rows["o4"] = r
    RD.cell(r, 1, F(T("敏感度：折現率、β、風險溢酬等 DCF 參數單獨變動，加權目標價最多變動 "), f'TEXT(MAX(Sens!F{sn["first"]}:F{sn["snap_first"] - 1}),"0.0%")', T("；最大的變動來自「"), f"Sens!B{sn['top']}", T("」，達 "), f'TEXT(Sens!C{sn["top"]},"0%")',
                    T("。現價隱含的折現率約 "), f'TEXT(RevDCF!B{RVA["w1"]},"0.0%")', T("（模型 "), f'TEXT(wacc,"0.0%")', T("），"),
                    f'IF(RevDCF!B{RVA["w1"]}>wacc,"即市場價格已反映高於模型的風險或較低的循環高點。","即市場比模型更樂觀。")'))
    r += 1
    rows["o5"] = r
    RD.cell(r, 1, F(T("對帳中："), yrs, T(" 在 Siliconomics 標對帳中。另一側讀數使情境 B 的 HBM 出貨位元 CY2027 由 "), f'TEXT(Coord!D{C0["cur"]},"0.00")', T(" 降到 "), f'TEXT(Coord!D{C0["alt"]},"0.00")',
                    T(" EB（"), f'TEXT(Coord!D{C0["altd"]},"+0%;-0%")', T("），目標價 "), f'TEXT({c_tp},"#,##0")', T("（"), c_rec, T("），與現行基準讀數（"), b_rec, T("）並列；個股不裁決。")))
    r += 1
    rows["drv"] = r
    RD.cell(r, 1, F(T("合併營收＝一般 DRAM＋HBM＋NAND＋其他。一般 DRAM、NAND＝前一年 ×（1＋位元成長）× 年均合約價指數比；HBM＝HBM 市場（出貨位元 × 單價）× 本公司營收市占 × 匯率；其他為固定輸入。"
                    "情境 B 2027 年："), f'TEXT({g("d_rev")},"#,##0")', T("（DRAM）＋"), f'TEXT({g("h_rev")},"#,##0")', T("（HBM）＋"), f'TEXT({g("n_rev")},"#,##0")', T("（NAND）＋"),
                    f'TEXT({g("o_rev")},"#,##0.0")', T("（其他）＝"), f'TEXT({g("tot")},"#,##0")', T(" 兆韓元。")))
    r += 1
    rows["ver"] = r
    RD.cell(r, 1, ver_text)
    return rows


def apply(wb, C, SIL, here):
    """build.py 呼叫：存檔前執行。回傳各區塊列號供檢查。"""
    if os.environ.get("V13_SKIP") == "1" or "Coord" not in wb.sheetnames:
        return None
    HR = hist_sheet(wb)
    PLA, S_, G_ = pl_sheet(wb, HR)
    RVA = revdcf(wb)
    V07 = val07(wb)
    SNA = sens_rel(wb, here)
    C0 = coord_c0(wb, SIL, here)
    RM = readme13(wb, PLA, S_, G_, RVA, SNA, C0, C["v13_ver"])
    return dict(hist=HR, pl=PLA, rd=RM, rv=RVA, val07=V07, sens=SNA, c0=C0)
