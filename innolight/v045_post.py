# 中際旭創 v0.4.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0（實績）公司實績對層級隱含值、兩讀數 C4）。
# 樣板：同次產業 Coherent v0.4.5（cohr/v045_post.py）、NVIDIA v0.3.10。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史 FY2023–FY2025（中國企業會計準則；年報數經券商研報財務摘要轉載）與 2026 年各季營收、歸母淨利（一季報、半年報、三季報，媒體轉載）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；FY2026E 含前三季實績與第四季假設）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的 EBIT 率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（整本重算快照 innolight/sens.json＋Val 頁封閉式案例）
#   Coord 頁末 F–H 節：讀取層級、C0（實績）、對帳中年份與損益表年度標示
#   README 頁末：v1.3 展示層白話列
# build_innolight.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL, CR)。環境變數 COORD_RUN（alt_run.py）或 CO_OVERRIDE（sens_innolight.py）設定時不執行。
import json, os
from openpyxl.styles import Font, PatternFill, Alignment

VER, BUILT = "0.4.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
KY23 = "https://pdf.dfcfw.com/pdf/H3_AP202510311772644632_1.pdf"   # 開源證券 2025-10-31 中際旭創三季報點評：財務摘要利潤表 2023A、2024A（百萬元）
KY25 = "https://pdf.dfcfw.com/pdf/H3_AP202604021820983821_1.pdf"   # 東吳證券 2026-04-02 中際旭創 2025 年報點評：財務摘要利潤表 2025A（百萬元）
Q1R = "https://www.ithome.com/0/940/030.htm"                        # IT之家 2026-04-16：2026 年一季報（營收 194.96 億元、歸母淨利 57.34 億元）
H1R = "https://finance.sina.com.cn/stock/hkstock/ggscyd/2026-08-21/doc-iniparan1339641.shtml"   # 新浪財經 2026-08-21：半年報（In 頁 h1_rev、h1_ni）
Q3R = "https://www.ithome.com/1/019/718.htm"                        # IT之家 2026-10-29：三季報（第三季營收 277.69 億元；前三季 695.47 億元、歸母淨利 222.38 億元）
SRC_H = (f"FY2023、FY2024：開源證券 2025-10-31 研報財務摘要利潤表（轉載公司年報數；{KY23}）；FY2025：東吳證券 2026-04-02 年報點評財務摘要利潤表（{KY25}），"
         "營收與歸母淨利連到 In 頁（公司年報，36氪、東方財富轉載）。百萬元換算十億元")


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
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- Hist：損益表歷史 ----------
# 欄：C＝FY2023、D＝FY2024、E＝FY2025；金額十億人民幣（百萬元 ÷ 1000）、EPS 人民幣
HIS = [
    ("rev", "營業收入", 10.718, 23.862, "=a25_rev_in"),
    ("cogs", "營業成本", 7.182, 15.796, 22.166),
    ("sur", "稅金及附加", 0.051, 0.047, 0.078),
    ("sell", "銷售費用", 0.125, 0.199, 0.226),
    ("adm", "管理費用", 0.434, 0.680, 0.761),
    ("rd", "研發費用", 0.739, 1.244, 1.615),
    ("fin", "財務費用（負值＝淨收益）", -0.084, -0.144, 0.183),
    ("imp", "減值損失（資產＋信用；負值＝損失）", -0.149, -0.079, -0.141),
    ("inv", "投資收益", 0.323, -0.033, 0.379),
    ("fv", "公允價值變動收益", -0.003, 0.067, 0.014),
    ("op", "營業利潤（公司報表口徑）", 2.494, 6.050, 13.597),
    ("nop", "營業外收支（淨額）", -0.002, 0.003, 0.003),
    ("pti", "利潤總額（稅前）", 2.492, 6.052, 13.600),
    ("tax", "所得稅費用", 0.285, 0.681, 2.020),
    ("nia", "淨利潤（含少數股東）", 2.208, 5.372, 11.580),
    ("mi", "少數股東損益", 0.034, 0.200, 0.782),
    ("ni", "歸屬母公司股東淨利潤", 2.174, 5.171, "=a25_ni_in"),
    ("eps", "EPS（按最新股本攤薄；券商財務摘要；人民幣）", 1.96, 4.65, 9.72),
]
QTR = [   # 2026 年各季（十億元）：C＝1Q26、D＝2Q26（＝上半年 − 1Q）、E＝3Q26、F＝前三季合計
    ("q_rev", "營業收入（2026 年各季）", 19.496, "=h1_rev-C{r}", "=q3_rev"),
    ("q_ni", "歸母淨利（2026 年各季）", 5.734, "=h1_ni-C{r}", "=q3_ni"),
]


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 損益表歷史（v0.4.5 起；PL 頁「標準損益表」的歷史三年、Coord 頁 C0（實績）的 2026 年各季實績）。會計年度＝日曆年。金額十億人民幣；EPS 人民幣。"
               "口徑：中國企業會計準則（公司只有一套報表，無 Non-GAAP）；EPS 為券商財務摘要「按最新股本攤薄」（2024 年起有資本公積轉增，與年報公告的基本 EPS 不同口徑）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        R[k] = r
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 6, "Verified（年報數，券商轉載）" if k != "eps" else "Interested-party（券商攤薄口徑）")
        H.cell(r, 7, SRC_H)
        r += 1
    r += 1
    H.cell(r, 1, "核對（十億元；公司以百萬元揭露、券商表四捨五入，加總可有 ±0.002 的捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("營業收入 − 營業成本 − 稅金及附加 − 銷售 − 管理 − 研發 − 財務費用 ＋ 減值 ＋ 投資收益 ＋ 公允價值變動 − 營業利潤：＝ −（其他收益、資產處置收益等未列項目）", "={c}{rev}-{c}{cogs}-{c}{sur}-{c}{sell}-{c}{adm}-{c}{rd}-{c}{fin}+{c}{imp}+{c}{inv}+{c}{fv}-{c}{op}"),
                     ("營業利潤 ＋ 營業外收支 − 利潤總額（應為 0）", "={c}{op}+{c}{nop}-{c}{pti}"),
                     ("利潤總額 − 所得稅 − 淨利潤（應為 0，捨入差）", "={c}{pti}-{c}{tax}-{c}{nia}"),
                     ("淨利潤 − 少數股東損益 − 歸母淨利（應為 0，捨入差）", "={c}{nia}-{c}{mi}-{c}{ni}")):
        H.cell(r, 2, labl)
        for c in "CDE":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    r += 1
    H.cell(r, 1, "2026 年各季實績（Coord 頁 C0（實績）用；十億元）").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "1Q26（1–3 月）", "2Q26（4–6 月）", "3Q26（7–9 月）", "前三季合計", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, labl, c1, c2, c3 in QTR:
        R[k] = r
        H.cell(r, 1, k); H.cell(r, 2, labl)
        H.cell(r, 3, c1).font = BLUE
        H.cell(r, 4, c2.format(r=r)); H.cell(r, 5, c3); H.cell(r, 6, f"=SUM(C{r}:E{r})")
        for c in "CDEF": H[f"{c}{r}"].number_format = "#,##0.000"
        H.cell(r, 7, "Verified（公司定期報告，媒體轉載）")
        H.cell(r, 8, f"1Q26：2026 年第一季度報告（IT之家 2026-04-16 轉載：營收 194.96 億元、歸母淨利 57.34 億元；{Q1R}）；2Q26＝半年報 − 1Q26（半年報：營收 417.78 億元、歸母淨利 136.51 億元，In 頁 h1_rev、h1_ni；{H1R}）；"
                        f"3Q26：第三季度報告（營收 277.69 億元、歸母淨利 85.87 億元，In 頁 q3_rev、q3_ni；{Q3R}）")
        r += 1
    H.cell(r, 2, "核對：前三季營收 − 三季報揭露 695.47 億元（應為 0）"); H[f"F{r}"] = f"=F{R['q_rev']}-69.547"; H[f"F{r}"].number_format = "0.000"; r += 1
    H.cell(r, 2, "核對：前三季歸母淨利 − 三季報揭露 222.38 億元（應為 0，捨入差）"); H[f"F{r}"] = f"=F{R['q_ni']}-22.238"; H[f"F{r}"].number_format = "0.000"; r += 1
    H.column_dimensions["A"].width = 10; H.column_dimensions["B"].width = 64
    for c in "CDEF": H.column_dimensions[c].width = 16
    H.column_dimensions["G"].width = 28; H.column_dimensions["H"].width = 90
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"      # 歷史 C–E＝Hist C–E；預測 F–I＝PL 情境 B 區塊 D–G（FY2026–FY2029）；J–L＝同區塊 H–J（延伸期）
DEPY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}   # 預測欄依賴的 Siliconomics 日曆年（FY(t)＝CY(t)；FY2030–32 延伸期沿用 CY2029 之後的路徑）
NOTE_COL = 13


def pl_standard(wb, HR, CR):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    g = lambda label: lab(P, label, rb, exact=True)
    R = dict(ai=g("AI 營收"), oth=g("其他營收（國內、電信等）"), rev=g("營收"), gp=g("毛利"), opx=g("營業費用、減值與其他"), pre=g("稅前利益（＝毛利 − 營業費用等）"),
             ni=g("歸母淨利"), sh=g("股數"), eps=g("EPS"))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.4.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 含前三季實績與第四季假設）。十億人民幣；EPS 人民幣。"
                 "只引用本頁上方、Hist 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（中國企業會計準則；公司只有一套報表，無 Non-GAAP）。營業利益＝稅前利益 − 非營業損益（投資收益＋公允價值變動 − 財務費用＋營業外收支），與 DCF 頁 EBIT 同一口徑；"
                 "營業費用＝毛利 − 營業利益（含稅金及附加、銷售、管理、研發、減值與其他收益）。其他收入（淨額）＝非營業損益。淨利＝歸屬母公司股東淨利。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E–FY2029E＝情境 B 由下而上（FY2026E＝錨點：前三季實績＋第四季季增假設；預測期非營業損益＝營收 × 0.55%，FY2025 年報比率）；FY2030E–FY2032E＝延伸期（收斂法只推營收與歸母淨利；"
                 "稅前利益＝歸母淨利 ÷（（1 − 稅率）× 歸母比例），營業利益＝稅前 − 非營業損益，與 DCF 頁 EBIT 相同）。延伸期毛利率沿用 FY2029E，營業費用為毛利 − 營業利益（殘差）。"
                 "歷史 EPS 為券商財務摘要的最新股本攤薄口徑；預測 EPS＝歸母淨利 ÷ 模型股數（FY2025 模型以 1H26 加權股數近似為 9.74，與攤薄口徑 9.72 不同）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":   # 預測欄標題：依賴 Siliconomics 對帳中年份者加「（對帳中）」（旗標在 Coord 頁）
        yr = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr}"&IF(Coord!${"BCDE"[DEPY[col] - 2026]}${CR["flag"]}=1,"（對帳中）","")'
    r_hdr = r; r += 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, NOTE_COL, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    hh = lambda k: [f"=Hist!{c}{HR[k]}" for c in "CDE"]
    fc = lambda k: [f"=PL!{c}{R[k]}" for c in "DEFG"]
    ex = lambda k: [f"=PL!{c}{R[k]}" for c in "HIJ"]
    FCST, EXT = "FGHI", "JKL"
    row("rev", "營收", "RMB B", hh("rev") + fc("rev") + ex("rev"), M, "FY2026E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [None] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 年增需 FY2022 營收（本表未列）")
    rc = r   # cogs＝rc、gp＝rc+1、gm＝rc+2、opx＝rc+3、oi＝rc+4、ni＝rc+10
    row("cogs", "營業成本（COGS）", "RMB B", [f"=Hist!{c}{HR['cogs']}" for c in "CDE"] + [f"={c}{L['rev']}-{c}{rc + 1}" for c in FCST] + [f"={c}{L['rev']}*(1-{p}{rc + 2})" for p, c in zip("IJK", EXT)], M,
        "歷史＝公司報表；FY2026E–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − 前一年毛利率）")
    row("gp", "毛利", "RMB B", [f"={c}{L['rev']}-{c}{rc}" for c in "CDE"] + fc("gp") + [f"={c}{L['rev']}-{c}{rc}" for c in EXT], M,
        "歷史＝營收 − 營業成本（全公司）；預測＝PL 情境 B（模型毛利率以光模組毛利率代用）")
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E 沿用前一年")
    row("opx", "營業費用（含稅金、銷售、管理、研發、減值與其他收益；＝毛利 − 營業利益）", "RMB B", [f"={c}{rc + 1}-{c}{rc + 4}" for c in COLS], M,
        "預測期＝PL「營業費用、減值與其他」＋非營業損益（該列在模型中以稅前利益倒推，含投資收益等）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "RMB B", [f"=Hist!{c}{HR['op']}-Hist!{c}{HR['inv']}-Hist!{c}{HR['fv']}+Hist!{c}{HR['fin']}" for c in "CDE"]
        + [f"=PL!{s}{R['pre']}-fc_nonop*{c}{L['rev']}" for c, s in zip(FCST, "DEFG")] + [f"={c}{rc + 10}/((1-tx)*attr)-fc_nonop*{c}{L['rev']}" for c in EXT], M,
        "歷史＝營業利潤 − 投資收益 − 公允價值變動 ＋ 財務費用；預測＝稅前利益 − 營收 × 非營業損益率（同 DCF 頁 EBIT）")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝非營業損益（投資收益＋公允價值變動 − 財務費用＋營業外收支）", "RMB B",
        [f"=Hist!{c}{HR['pti']}-{c}{L['oi']}" for c in "CDE"] + [f"=fc_nonop*{c}{L['rev']}" for c in FCST + EXT], M, "歷史＝利潤總額 − 營業利益；預測＝營收 × 0.55%（In 頁 fc_nonop）")
    row("pti", "稅前利益", "RMB B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "RMB B", [f"=Hist!{c}{HR['tax']}" for c in "CDE"] + [f"={c}{L['pti']}*tx" for c in FCST + EXT], M, "預測＝稅前利益 × 有效稅率（In 頁 tx）")
    row("mi", "少數股東損益", "RMB B", [f"=Hist!{c}{HR['mi']}" for c in "CDE"] + [f"={c}{L['pti']}*(1-tx)*(1-attr)" for c in FCST + EXT], M, "預測＝稅後利益 ×（1 − 歸母比例 attr）")
    row("ni", "稅後淨利（歸屬母公司股東）", "RMB B", hh("ni") + fc("ni") + ex("ni"), M)
    assert L["ni"] == rc + 10
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數（預測＝模型股數）", "十億股", [f"={c}{L['ni']}/Hist!{c}{HR['eps']}" for c in "CDE"] + [f"=PL!{c}{R['sh']}" for c in "DEFGHIJ"], "#,##0.0000",
        "歷史＝歸母淨利 ÷ 攤薄 EPS（隱含最新股本）；FY2026E＝1H 加權與發行後平均（模型）")
    row("eps", "EPS（歷史＝最新股本攤薄；預測＝歸母淨利 ÷ 模型股數）", "RMB", hh("eps") + [f"=PL!{c}{R['eps']}" for c in "DEFGHIJ"], "0.00",
        "歷史＝券商財務摘要（FY2023 1.96、FY2024 4.65、FY2025 9.72）；FY2026E 前三季公告基本 EPS 19.78 元")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" if c != "C" else None for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [None] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "歸母淨利")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "RMB B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "RMB B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 少數股東損益 − 歸母淨利（歷史有 ±0.002 的捨入差）", "RMB B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['mi']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅前利益 − PL 情境 B 稅前利益（FY2026E–FY2029E）", "RMB B", [None] * 3 + [f"={c}{L['pti']}-PL!{s}{R['pre']}" for c, s in zip(FCST, "DEFG")] + [None] * 3, "0.000")
    row("c5", "核對：稅後淨利 − PL 情境 B（FY2026E–FY2032E）；FY2025A − In 頁", "RMB B",
        [None, None, f"=E{L['ni']}-a25_ni_in"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJKL", "DEFGHIJ")], "0.000")
    row("c6", "核對：EPS − PL 情境 B（FY2026E–FY2032E）", "RMB", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHIJKL", "DEFGHIJ")], "0.000")
    row("c7", "參考：FY2025A 模型營業利益（PL 情境 B 稅前 − 非營業）− 公司報表口徑（模型毛利率以光模組代用，差額為代用誤差）", "RMB B",
        [None, None, f"=(PL!C{R['pre']}-fc_nonop*a25_rev_in)-E{L['oi']}"] + [None] * 7, "0.000")
    P.column_dimensions["A"].width = max(P.column_dimensions["A"].width or 0, 66)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, R=R)
    return L


# ---------- RevDCF：反向 DCF（同 Coherent v0.4.5；DCF 頁版面相同：C–H＝FY2027–FY2032） ----------
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
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年 EBIT 與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027 視為錨點年不動（三情境只差 CPO 情境、幾乎相同）。反解值超出合理範圍時照實列出，不外推")
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
    kv("px", "現價（人民幣）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（人民幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；RMB B）", f"={ncr}", "#,##0.0", "Val 頁：淨現金與投資＋stub 期 FCFF − 已宣告股利（與 WACC 無關）")
    kv("sh", "評價用股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（RMB B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；RMB B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（人民幣）", f"=DCF!$C${ps}", "0.00")
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
    yr("rev", "營收（RMB B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（RMB B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元 EBIT 對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（RMB B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（RMB B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（RMB B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期 EBIT 率（FY2028–FY2032 統一一個 EBIT 率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（RMB B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（RMB B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效 EBIT 率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型 EBIT 率")
    kv("m1", "現價隱含的 EBIT 率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(D{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；EBIT 率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（RMB B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（RMB B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（RMB B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；RMB B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（人民幣）"], 1):
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
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（人民幣，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；人民幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 元；要支持現價 "), 'TEXT(px,"0.00")', T(" 元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大為","縮小為")', f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<B{A["w0"]},"降到","升到")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面"),
                     f'IF(B{A["s1"]}>=1,"放大","縮小")', f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0.0")', T(" 十億元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0.0")', T("）；或 EBIT 率"),
                     f'IF(B{A["m1"]}>B{A["m0"]},"提高","降低")', T("到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 中際旭創過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給中際旭創的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("模型採用的歷史倍數比。")))
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
    V.cell(r, 1, "0.7 估值並列：四個目標價（v0.4.5 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["目標價", "人民幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T(" × 12 個月後前瞻 EPS（情境加權）"), f'TEXT({feps},"0.00")', T(" 元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{r_px}"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px 列；300308.SZ）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節） ----------
def sens_tor(wb, here):
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")               # B、C、D＝情境 A、B、C
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    e_ = {}
    for s in "ABC":   # 各情境 12 個月後前瞻 EPS（與 Val 頁同式：2／12 × FY2027 ＋ 10／12 × FY2028；PL 頁摘要區 E、F 欄）
        h = lab(P, f"情境 {s}", 1, exact=True); er = lab(P, "EPS", h, exact=True)
        e_[s] = f"(2/12*PL!$E${er}+10/12*PL!$F${er})"
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.4.5 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（innolight/sens_innolight.py：三情境同時套用 In 頁低／高欄，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（人民幣）", "高案：加權目標價（人民幣）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    base_f = f"{tp}"
    frame = [
        ("同業比較法的相對倍數調整係數 夾值下限 0.5／上限 2.0（基準＝歷史倍數比 約 0.88）",
         f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*0.5", f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*2", "Val 頁：Comparable（未調整）× 係數；低＝夾值下限 0.5、高＝夾值上限 2.0（tools/relative_multiple.py 的夾值區間）"),
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{e_['C']}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{e_['A']}", "Val 頁：單一情境的 DCF 價值與前瞻 EPS"),
        ("DCF 在加權中的權重 0%（只看同業比較法）／100%（只看現金流折現）（基準＝50%）",
         f"={cadj}", f"={dcf}", "Val 頁：Comparable（調整後）與 DCF 目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    for k, d in J["biz"].items():
        SE.cell(r, 1, d["label"])
        for side, col_v, col_p in (("低", 4, 2), ("高", 5, 3)):
            if d.get(side) is None: continue
            SE.cell(r, col_v, d[side]["tp"]).number_format = "#,##0.00"; SE.cell(r, col_v).font = BLUE
            SE.cell(r, col_p, f"={'DE'[col_v - 4]}{r}/{base_f}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（innolight/sens.json）；覆寫 In 頁低／高欄（或主引擎參數），三情境同時套用"); r += 1
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
    r += 1
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f'=IF(ABS({tp}-{J["_base"]["tp"]!r})<=1E-6*ABS({tp}),"CURRENT","STALE")')
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 innolight/sens_innolight.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調——讀取層級、C0（實績）、對帳中年份與損益表年度標示 ----------
def coord_c0(wb, SIL, HR):
    C, I, U, RC = wb["Coord"], wb["In"], wb["Units"], wb["Recon"]
    r = C.max_row + 2
    r_flag = lab(C, "對帳中旗標")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    net_file = src_in("mpl_26").split("｜")[0].split("/")[-1]
    tsmc_file = src_in("um28_B").split("｜")[0].split("/")[-1]
    # ---- F. 讀取層級 ----
    C.cell(r, 1, "F. 讀取層級（三層基準協調規則 v1.4.2；共用工作單 2026-11-14）：本模型讀哪一層、哪一版。自層級讀入的值只列在本節，不進 C0 判定").font = BOLD; f_title = r; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("AI 可插拔光模組市場、AI CPO 光學市場、AI IT 資本支出 CY2026–CY2029（情境 A，B／C 讀同字母 CPO 情境列）", "市場層 markets/network（其輸入由 Siliconomics 行業變數 M2、M8、M10 供給）", net_file,
         "AI 營收（美元）＝可插拔市場 × 有效份額 × 單價係數＋CPO 市場 × CPO 份額", '="CY2026 可插拔 "&TEXT(mpl_26,"0.00")&"、CPO "&TEXT(mcp_26,"0.00")&"；CY2027 可插拔 "&TEXT(mpl_27,"0.00")&"、CPO "&TEXT(mcp_27,"0.00")&"（US$B）"'),
        ("中際旭創占 CPO 光學金額 CY2027–CY2029", "Siliconomics（M12）", vtxt, "CPO 光學市場 × 中際旭創份額；AI 營收的 CPO 部分", '=TEXT(cpo_27,"0%")&"／"&TEXT(cpo_28,"0%")&"／"&TEXT(cpo_29,"0%")&"（Assumed）"'),
        ("情境顆數係數 CY2028／CY2029（情境 B＝基準）", "Siliconomics（M2b，經台積電個股模型 In 頁轉入；A、C 為個股自訂）", f"{vtxt}；{tsmc_file}", "市場值 × 係數", '="情境 B："&TEXT(um28_B,"0.00")&"／"&TEXT(um29_B,"0.000")&"；A："&TEXT(um28_A,"0.00")&"／"&TEXT(um29_A,"0.00")&"；C："&TEXT(um28_C,"0.00")&"／"&TEXT(um29_C,"0.00")'),
        ("人民幣兌美元（FY2026 校準與 FY2027 起換算）", "Siliconomics（SIF_fx_rmb）", vtxt, "AI 營收美元 → 人民幣", '=TEXT(fx,"0.00")'),
        ("Wattonomics 錨定係數 CY2027–CY2029（Sens 案例；兩讀數建置時改放另一側讀數比值）", "Wattonomics（經 Siliconomics SIF_WTAnchor_Base）", vtxt, "Sens 頁「Wattonomics 錨定」案例；C4 另一側讀數整本重建", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")'),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics（行業變數 M30）", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）",
         f'="CY2027–CY2029＝對帳中；比值 "&TEXT(C{lab(C, "另一側讀數 ÷ 現行基準")},"0.000")&"／"&TEXT(D{lab(C, "另一側讀數 ÷ 現行基準")},"0.000")&"／"&TEXT(E{lab(C, "另一側讀數 ÷ 現行基準")},"0.000")'),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（專案規則第 3 條）。中際旭創可插拔有效份額（Units 頁）是本模型自己的變數、為光模組份額的權威（Coherent、新易盛、Fabrinet 模型讀取），不在此節。").font = Font(italic=True); r += 2
    # ---- G. C0（實績） ----
    C.cell(r, 1, "G. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調）").font = BOLD; g_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝公司 2026 年已公布的營業收入（一季報、半年報、三季報；公司不揭露 AI 與其他營收拆分，用合併營收）。層級側＝同口徑的層級隱含值＝markets/network 的 AI 可插拔光模組市場 × "
                 "中際旭創有效份額 × 匯率（AI 營收）＋其他營收（層級沒有此項，取本模型 Q 頁推導值，屬 Assumed、非層級值）。時差：會計年度＝日曆年；已公布期間＝2026-01～2026-09（9 個月）。"
                 "主讀數＝前三季累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（3Q26，2026-07～09）× 4 vs 層級 CY2026（約等於該季中點 2026-08 的年化速率；"
                 "成長年中，最近一季落在年度後段時偏高，不得寫成夾住真值）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只有一個超過 → 時差邊界（參考）；都在內 → 一致。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); g_note1 = r; r += 1
    C.cell(r, 1, "同源校準（C0 第 2 點補充）：層級側的中際旭創有效份額由本模型校準——CY2026 有效份額＝FY2026 錨點 AI 營收（美元）÷ 可插拔市場，而 FY2026 錨點是用本模型 1H26、3Q26 實績與 4Q26 季增假設推得。"
                 "市場值乘回份額時消去，層級隱含營收恆等於本模型 FY2026 錨點營收（下列核對列＝0）。所以這張表實際檢驗的只有時間分攤（已公布 9 個月占全年的比例）與最近一季年化，"
                 "不檢驗層級市場值或份額、不能裁決層級對錯：判定標「參考（同源校準）」、不計入對帳中；C0 結論＝無可比實績（層級值同源）。要獨立檢驗需要公司揭露 AI 光模組（數通）營收或出貨量，"
                 "或獨立的中際旭創份額估計（研調排名只有 2025 年營收與 800G／1.6T 份額區間，非 CY2026 同口徑）；均未找到。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); g_note2 = r; r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000"):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); IN[key] = r; r += 1
    hq = HR["q_rev"]
    inp("q1", "營業收入：1Q26（一季報）", f"=Hist!C{hq}", "RMB B", "2026-01-01～2026-03-31", "Hist 頁；2026 年第一季度報告 194.96 億元（IT之家 2026-04-16 轉載；網址見 Hist 頁來源欄）", "Verified", GREEN)
    inp("q2", "營業收入：2Q26（半年報 − 一季報）", f"=Hist!D{hq}", "RMB B", "2026-04-01～2026-06-30", "Hist 頁；半年報 417.78 億元（新浪財經 2026-08-21 轉載；網址見 Hist 頁來源欄）− 1Q26", "Verified（推導）", GREEN)
    inp("q3", "營業收入：3Q26（三季報）", f"=Hist!E{hq}", "RMB B", "2026-07-01～2026-09-30", "Hist 頁；第三季度報告 277.69 億元（IT之家 2026-10-29 轉載；網址見 Hist 頁來源欄）；前三季 695.47 億元", "Verified", GREEN)
    inp("mo", "已公布月數（1Q26–3Q26）", 9, "月", "2026-01～2026-09", "三季＝9 個月", "Derived", BLUE, fmt="0")
    inp("m", "層級：AI 可插拔光模組市場 CY2026（markets/network）", "=mpl_26", "US$B", "CY2026（日曆年）", f"In 頁 mpl_26；{net_file}（Mkt 頁「其中可插拔」）", "Derived（層：顆數 × 每顆光學內容，行業變數 Assumed）", GREEN)
    inp("sh", "中際旭創可插拔有效份額 CY2026（本模型校準）", f"=Units!C{lab(U, '中際旭創可插拔有效份額')}", "%", "CY2026", "Units 頁；＝FY2026 錨點 AI 營收（美元）÷ 可插拔市場（PL 頁 A 節）：本模型以公司實績校準，不是層級值", "Derived（同源校準）", GREEN, fmt="0.0%")
    inp("fx", "人民幣兌美元（Siliconomics）", "=fx", "RMB／US$", "2026", "In 頁 fx（Siliconomics SIF_fx_rmb）", "Verified", GREEN, fmt="0.00")
    inp("cp", "層級：AI CPO 光學市場 CY2026 × 中際旭創 CPO 份額（CY2026 視為 0）", "=mcp_26*0", "US$B", "CY2026", "In 頁 mcp_26；FY2026 CPO 營收視為 0（PL 頁 A 節）", "Assumed", GREEN)
    inp("ai", "層級隱含：AI 營收 CY2026（（可插拔市場 × 份額＋CPO）× 匯率）", f"=(B{r - 4}*B{r - 3}+B{r - 1})*B{r - 2}", "RMB B", "CY2026（日曆年）", "上四列", "Derived")
    inp("oth", "其他營收 CY2026（本模型 Q 頁：營收 × 8%，非層級值）", "=a26_oth", "RMB B", "CY2026", "Q 頁「其他營收」列（In 頁 oth_s，Assumed）", "Assumed（非公司揭露）", GREEN)
    inp("L", "層級隱含：營業收入 CY2026（AI 營收層級隱含＋其他營收）", f"=B{IN['ai']}+B{IN['oth']}", "RMB B", "CY2026（日曆年）", "上兩列加總", "Derived")
    inp("chk", "核對：層級隱含營收 − 本模型 FY2026 錨點營收（應為 0；同源校準的恆等式）", f"=B{IN['L']}-a26_rev", "RMB B", "CY2026", "Q 頁 G4（具名 a26_rev）", "Derived", fmt="0.000000")
    inp("g", "參考：同花順一致預期 FY2026 營收（不判定）", "=cons_rev26", "RMB B", "FY2026（預測）", "In 頁 cons_rev26（同花順 15 家，2026-10-23）", "Interested-party", GREEN)
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    pct = '+0.0%;-0.0%'
    b = IN
    C.cell(r, 1, "主讀數：營業收入，2026-01～09 累計實績 vs 層級 CY2026 × 9 ÷ 12")
    C.cell(r, 2, f"=B{b['q1']}+B{b['q2']}+B{b['q3']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%（參考）","在 ±15% 內（參考）")')
    C.cell(r, 6, "RMB B；公司實績＝前三季營收加總；層級＝CY2026 全年 × 9 ÷ 12")
    C.cell(r, 7, "公司：一季報、半年報、三季報（Hist 頁）；層級：markets/network × 本模型校準份額（同源）× 匯率＋其他營收（Q 頁，Assumed）")
    C.cell(r, 8, "Verified（公司）／Derived（層級隱含、線性時差）")
    C.cell(r, 9, f'="層級全年值含 4Q26（2026-10～12，模型季增假設）；2026 年成長快、後段季度較大，線性分攤使層級值偏高、主讀數偏低。層級隱含全年 "&TEXT(C{r}/B{b["mo"]}*12,"0.00")&" 十億元"'); rm = r; r += 1
    C.cell(r, 1, "第二讀數：營業收入，最近一季（3Q26）× 4 vs 層級 CY2026")
    C.cell(r, 2, f"=B{b['q3']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%（參考）","在 ±15% 內（參考）")')
    C.cell(r, 6, "RMB B；3Q26（2026-07～09）單季 × 4（年化）；層級＝CY2026 全年")
    C.cell(r, 7, "公司：第三季度報告（Hist 頁）；層級同上")
    C.cell(r, 8, "Verified（公司）／Derived（層級隱含、年化）")
    C.cell(r, 9, "最近一季中點約 2026-08；年化速率與 CY2026 全年平均相比位於後段，成長年中偏高"); rs = r; r += 1
    C.cell(r, 1, "C0（實績）判定：營業收入 CY2026").font = BOLD
    C.cell(r, 5, f'=IF(AND(ABS(D{rm})>0.15,ABS(D{rs})>0.15,SIGN(D{rm})=SIGN(D{rs})),"參考（同源校準）：兩讀數同向超過 ±15%","參考（同源校準）：不判定")').font = BOLD
    C.cell(r, 6, "同源校準：判定標參考、不計入對帳中；機械結果見主讀數與第二讀數列")
    C.cell(r, 9, f'="結論：無可比實績（層級值同源）。機械結果：主讀數 "&TEXT(D{rm},"+0.0%;-0.0%")&"、第二讀數 "&TEXT(D{rs},"+0.0%;-0.0%")&"，"&IF(AND(ABS(D{rm})<=0.15,ABS(D{rs})<=0.15),"兩者都在 ±15% 內（若層級值獨立，會判一致）","有讀數超過 ±15%")&"；只表示已公布 9 個月與最近一季年化和本模型全年路徑相容，不是對層級市場值或份額的獨立檢驗。缺的獨立揭露：公司 AI 光模組（數通）營收或出貨量、獨立的中際旭創 CY2026 份額"')
    rv = r; r += 1
    G1 = r - 1
    r += 1
    C.cell(r, 1, "層級之間與外部估計（參考；不計入中際旭創的 C0）：市場層 markets/network 的 CY2026 AI 可插拔光模組市場 vs LightCounting（外部研調）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "市場層", "外部估計", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    L0 = r
    C.cell(r, 1, "AI 可插拔光模組市場 CY2026 vs LightCounting 2026 全球數通光模組市場")
    C.cell(r, 2, "=mpl_26").number_format = "#,##0.00"; C.cell(r, 3, "=lc_dc26").number_format = "#,##0.00"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "層間／外部參考（不判定、不計入）")
    C.cell(r, 6, "US$B；市場層＝AI 加速器 scale-out 可插拔光學；LightCounting＝全球數通光模組（含非 AI），口徑不同")
    C.cell(r, 7, f"{net_file}（Mkt 頁「其中可插拔」）；In 頁 LightCounting 2026 數通光模組市場列（愛建證券轉引；Recon 頁同值）")
    C.cell(r, 8, "Derived（市場層）／Interested-party（外部估計）"); C.cell(r, 9, "外部估計只列參考；差距屬市場層與外部估計之間，不是中際旭創自己的量，不開對帳單")
    L1 = r; r += 2
    # ---- H. 對帳中年份與損益表年度標示 ----
    C.cell(r, 1, "H. 對帳中年份與損益表年度標示（C4；FY(t)＝CY(t)，FY2030E–FY2032E 為 FY2029E 之後的延伸期，依賴 CY2029）").font = BOLD; r += 1
    fys = list(range(2026, 2033))
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in fys], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年")
    dep = [min(y, 2029) for y in fys]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2026E–FY2032E）")
    for j, fy in enumerate(fys):
        C.cell(r, 2 + j, f'=IF({"$" + "BCDE"[dep[j] - 2026] + "$" + str(r_flag)}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（G 節 C0（實績）；同源校準與層級之間參考列不計）"); C.cell(r, 2, f'=COUNTIF(E{G0}:E{G1},"對帳中")'); r_cnt = r; r += 1
    C.cell(r, 1, "說明：本公司 C0（實績）判定＝無可比實績（層級值同源），對帳中項目 0。產業部署 GW 為層與層之間的差距（Siliconomics 標對帳中，對帳單 r3 處理），本模型並列兩讀數目標價（Coord 頁 A–E 節），不裁決，也不計入對帳中項目數。"); r_note = r; r += 1
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, f_title=f_title, g_title=g_title, g_note1=g_note1, g_note2=g_note2, G0=G0, G1=G1, rm=rm, rs=rs, rv=rv, L0=L0, L1=L1, H_PL=H_PL, cnt=r_cnt, note=r_note, IN=IN)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cu, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef = vb("相對倍數調整係數")
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    rb = lab(P, "情境 B：")
    pr = lambda lbl: lab(P, lbl, rb, exact=True)
    rev_b, ai_b, oth_b = pr("營收"), pr("AI 營收"), pr("其他營收（國內、電信等）")
    reps27, reps28 = lab(RC, "FY2027 EPS vs 共識"), lab(RC, "FY2028 EPS vs 共識")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, SEN['first'])}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.4.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 元，比現價 "), 'TEXT(px,"0")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("。現金流折現只有 "), f'TEXT({dcf},"0")',
        T(" 元，同業比較法 "), f'TEXT({cadj},"0")', T(" 元（已含相對倍數調整 "), f'TEXT({coef},"0.00")', T(" 倍），上檔主要來自後者。"))
    put("rec", T("投資建議「"), rec, T("」依加權目標價判定，且取決於同業比較法：不做相對倍數調整時加權目標價 "), f'TEXT({noadj},"0")', T(" 元（"), f'TEXT({noadj}/px-1,"+0%;-0%")', T("），評等"), REC5(f"({noadj}/px-1)"),
        T("；現金流折現權重提高到七成時為 "), f'TEXT(0.7*{dcf}+0.3*{cadj},"0")', T(" 元（"), f'TEXT((0.7*{dcf}+0.3*{cadj})/px-1,"+0%;-0%")', T("），評等"), REC5(f"((0.7*{dcf}+0.3*{cadj})/px-1)"),
        T("。評等依上檔空間五級固定規則；現價隱含的條件見反向 DCF。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0.00")', T(" 元＝現金流折現 "), f'TEXT({dcf},"0.00")', T(" 與同業比較法 "), f'TEXT({cadj},"0.00")', T(" 各半；比現價 "), 'TEXT(px,"0.00")', T(" "), f'IF({up}<0,"低","高")', f'TEXT(ABS({up}),"0%")',
        T("，評等"), rec, T("；市場共識 "), 'TEXT(cons_tp,"0")', T(" 元。"))
    put("o2", T("現價隱含：現金流折現要支持現價，折現率須"), f'IF({rv("w1")}<wacc,"降到","升到")', f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收"), f'IF({rv("s1")}>=1,"放大","縮小")', f'TEXT({rv("s1")},"0.00")',
        T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0.0")', T("，模型 "), f'TEXT({rv("rv0")},"0.0")', T(" 十億元）。"))
    put("o3", T("盈餘對共識：情境 B 的 FY2027 EPS "), f'TEXT(Recon!E{reps27},"0.00")', T(" 元比共識 "), f'TEXT(Recon!B{reps27},"0.00")', f'IF(Recon!F{reps27}>=0," 高 "," 低 ")', f'TEXT(ABS(Recon!F{reps27}),"0%")',
        T("，FY2028 EPS "), f'TEXT(Recon!E{reps28},"0.00")', T(" 比共識 "), f'TEXT(Recon!B{reps28},"0.00")', f'IF(Recon!F{reps28}>=0," 高 "," 低 ")', f'TEXT(ABS(Recon!F{reps28}),"0%")',
        T("（顆數係數 CY2028 起下修）；共識只作對照，未用於校準。"))
    put("o4", T("最敏感：AI 光模組市場 ×0.8／×1.2（"), f'TEXT({SB("AI 光模組市場", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("AI 光模組市場", "C")},"+0%;-0%")',
        T("）、有效份額（"), f'TEXT({SB("中際旭創可插拔有效份額", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("中際旭創可插拔有效份額", "C")},"+0%;-0%")',
        T("）、毛利率路徑（低案 "), f'TEXT({SB("毛利率 FY2027", "B")},"+0%;-0%")', T("）；同業比較法的倍數調整夾值 0.5／2.0 為 "),
        f'TEXT({SB("同業比較法的相對倍數調整係數", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("同業比較法的相對倍數調整係數", "C")},"+0%;-0%")', T("（加權目標價相對基準）。"))
    put("o5", T("對帳中：產業部署 GW 換另一側讀數後，加權目標價 "), f'TEXT(Coord!C{CO["tp_row"]},"0.00")', T(" 元、評等"), f'Coord!C{CO["rec_row"]}', T("（現行 "), f'TEXT({tp},"0.00")', T("）。C0 實績：2026 年前三季營收 "), f'TEXT(Coord!B{CO["rm"]},"0.0")',
        T(" 對層級隱含 "), f'TEXT(Coord!C{CO["rm"]},"0.0")', T(" 十億元（"), f'TEXT(Coord!D{CO["rm"]},"+0%;-0%")', T("），但份額由本模型校準，無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：AI 營收＝（AI 可插拔光模組市場 × 情境顆數係數 × 中際旭創有效份額 × 單價係數＋AI CPO 光學市場 × CPO 份額）× 匯率；有效份額以 FY2026 錨點校準，之後乘份額相對值；"
                 "其他營收（國內、電信等）以固定年增推算；兩者相加為合併營收。FY2028：AI 營收 "), f'TEXT(PL!F{ai_b},"0.0")', T("、其他 "), f'TEXT(PL!F{oth_b},"0.0")', T("，合計 "), f'TEXT(PL!F{rev_b},"0.0")', T(" 十億元。"))
    put("coord", T("三層協調：本模型讀市場層 markets/network 的光學市場、Siliconomics 的 CPO 份額、匯率與情境顆數係數（Coord 頁 F 節）。C0 改以公司實績對層級隱含值（Coord 頁 G 節）：2026 年前三季營收累計 "), f'TEXT(Coord!B{CO["rm"]},"0.00")',
        T(" 十億元對層級隱含依月數折算 "), f'TEXT(Coord!C{CO["rm"]},"0.00")', T("，差 "), f'TEXT(Coord!D{CO["rm"]},"+0.0%;-0.0%")', T("；最近一季年化差 "), f'TEXT(Coord!D{CO["rs"]},"+0.0%;-0.0%")', T("；層級份額由本模型校準，判定「"), f'Coord!E{CO["rv"]}',
        T("」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項。產業部署 GW 兩層差距另列參考、不計入；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    put("ver", T("v0.4.6（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113；Yahoo Finance 日線收盤，closes_20261020_20261113.csv）。現價 493.33 → "), 'TEXT(px,"0.00")',
        T("；同業前瞻 P/E 中位數 23.34 → "), f'TEXT({vb("同業前瞻 P/E 中位數")},"0.00")', T("；加權目標價 644.77 → "), f'TEXT({tp},"0.00")',
        T("；上檔空間 +30.7%（買進）→ "), f'TEXT({up},"+0.0%;-0.0%")', T("（"), rec, T("）。共識 EPS、rf、β、預測用匯率、營運假設、情境權重與三層基準讀數不變；DCF 只經市值 D/E 對 WACC 微幅變動。未換價同業：無（五家同業與本公司皆取自收盤檔）。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3、三層協調（讀取層級、C0 實績）與兩讀數目標價（三層基準協調規則 v1.4.2 C4）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史與 2026 年各季實績）、PL 頁末標準損益表"
                  "（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，中國企業會計準則）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準（整本重算快照 innolight/sens.json）、"
                  "Coord 頁 A–E 節兩讀數目標價與 F–H 節（讀取層級、C0（實績）、損益表年度標示）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver_prev"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def apply(wb, here, SIL, CR):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("CO_OVERRIDE"):
        return "v045_post：略過（alt_run／sens_innolight 建置）"
    HR = hist_sheet(wb)
    PLA = pl_standard(wb, HR, CR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb, here)
    CO = coord_c0(wb, SIL, HR)
    CO["tp_row"], CO["rec_row"] = CR["tp"], CR["rec"]
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    wb._v045 = dict(PL=PLA, RV=RVA, RM=RMA, SEN=SEN, CO=CO, VG=VG, HR=HR)
    return f"v045_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；Coord F–H 節；README v1.3 白話列 {len(RMA)} 列"
