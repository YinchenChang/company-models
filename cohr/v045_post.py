# Coherent v0.4.5（2026-11-14）後處理：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調（讀取層級、C0（實績）公司實績對層級隱含值、兩讀數 C4）。
# 樣板 NVIDIA v0.3.10（nvidia/v037_post.py）、Broadcom v0.5.6／v0.5.7（avgo/v056_post.py、v057_post.py）、Ibiden v0.5.6。只新增，不改既有數值格：
#   Hist 頁（新）：損益表歷史（FY2023–FY2026 non-GAAP；FY2024–FY2026 另列 GAAP 對照；公司 8-K 新聞稿，SEC EDGAR）
#   PL 頁末：標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B；Non-GAAP＝模型口徑）、比率區、GAAP 對照、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的 EBIT 率、營收規模、WACC（情境 B 路徑）與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（整本重算快照 cohr/sens.json＋Val 頁封閉式案例）
#   Coord 頁末 F–H 節：讀取層級、C0（實績）、對帳中年份與損益表年度標示（對帳中旗標列直接引用，不另設具名範圍）
#   README 頁末：v1.3 展示層白話列
# build_cohr.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL, CR)。環境變數 COORD_RUN（alt_run.py）或 CO_OVERRIDE（sens_cohr.py）設定時不執行。
import json, os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.5", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
EDG = "https://www.sec.gov/Archives/edgar/data/820318/"
K26Q4 = EDG + "000119312526346860/d128030dex991.htm"     # FY2026 Q4 新聞稿（8-K Ex.99.1，2026-08）：Q4、Q3 單季分部營收、FY2026 全年
K27Q1 = EDG + "000119312526447200/d176836dex991.htm"     # FY2027 Q1 新聞稿（8-K Ex.99.1，2026-11-04）：Table 5 分部營收
K25Q4 = EDG + "000119312525179833/d22249dex991.htm"      # FY2025 Q4 新聞稿（8-K Ex.99.1，2025-08）：FY2025 與 FY2024（新非 GAAP 口徑）
K24Q4 = EDG + "000119312524201297/d873258dex991.htm"     # FY2024 Q4 新聞稿（8-K Ex.99.1，2024-08）：FY2024 與 FY2023（舊非 GAAP 口徑）
SRC_H = (f"FY2027 Q1：{K27Q1}（Table 5 分部營收，Coord 頁 C0 用）；FY2023 與 FY2024 舊口徑：{K24Q4}（FY2024 Q4 新聞稿）；FY2024 新口徑與 FY2025：{K25Q4}（FY2025 Q4 新聞稿；2025 年起非 GAAP 口徑調整，FY2024 比較欄已依新口徑重編）；"
         f"FY2026：{K26Q4}（FY2026 Q4 新聞稿；本模型 In 頁 G 節與 Q 頁）。百萬美元換算十億美元")


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
# 欄：C＝FY2023（舊口徑）、D＝FY2024（舊口徑，FY2024 原新聞稿）、E＝FY2024（新口徑，FY2025 新聞稿比較欄）、F＝FY2025、G＝FY2026；金額十億美元（百萬美元 ÷ 1000）、EPS 美元、股數十億股
HIS = [
    ("rev", "營收（GAAP＝non-GAAP）", 5.1601, 4.7077, 4.7077, "=a25_rev_in", "=a26_rev_in"),
    ("gp", "毛利（non-GAAP）", 1.9840, 1.6961, 1.6157, "=a25_gp_in", "=a26_gp_in"),
    ("opx", "營業費用（non-GAAP；研發＋銷管）", 1.0176, 0.9866, 0.9977, "=a25_opx_in", "=a26_opx_in"),
    ("oi", "營業利益（non-GAAP）", 0.9665, 0.7095, 0.6180, "=Q!B12", "=Q!C12"),
    ("int", "利息與其他淨支出（non-GAAP）", 0.2729, 0.2323, 0.2323, "=a25_int_in", "=a26_int_in"),
    ("tax", "所得稅（non-GAAP）", 0.1260, 0.0985, 0.0777, "=a25_tax_in", "=a26_tax_in"),
    ("ni", "歸屬 Coherent 淨利（non-GAAP）", 0.5674, 0.3812, 0.3107, "=a25_ni_in", "=a26_ni_in"),
    ("pref", "優先股股利（non-GAAP EPS 分子扣除；FY2023 由公告 EPS 反推）", 0.1156, 0.1234, 0.1234, 0.1299, 0),
    ("sh", "稀釋股數（non-GAAP；十億股）", 0.1506, 0.1543, 0.1543, "=a25_sh_in", "=a26_sh_in"),
    ("eps", "EPS（non-GAAP 稀釋，公司公告；美元）", 3.00, 1.67, 1.21, "=a25_eps_in", "=a26_eps_in"),
    ("g_gp", "毛利（GAAP）", 1.6183, 1.4560, 1.4560, 2.0433, 2.6690),
    ("g_oi", "營業利益（GAAP）", -0.0371, 0.0961, 0.0961, 0.2899, 0.8979),
    ("g_int", "利息與其他淨支出（GAAP）", 0.3184, 0.2438, 0.2438, 0.1957, 0.0501),
    ("g_tax", "所得稅（GAAP；FY2023 為淨利益）", -0.0961, 0.0111, 0.0111, 0.0641, 0.0608),
    ("g_ni", "歸屬 Coherent 淨利（GAAP）", -0.2595, -0.1562, -0.1562, 0.0494, 0.8050),
    ("g_eps", "EPS（GAAP 稀釋，公司公告；美元）", -2.93, -1.84, -1.84, -0.52, 4.12),
]


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 損益表歷史（v0.4.5 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。會計年度 FY(t)＝t−1 年 7 月～t 年 6 月。金額十億美元；EPS 美元；股數十億股。"
               "non-GAAP 口徑：公司於 FY2025 起調整非 GAAP 定義（FY2024 毛利 1,696.1→1,615.7、營業利益 709.5→618.0、淨利 381.2→310.7、EPS 1.67→1.21 百萬美元，新聞稿比較欄已重編）；"
               "FY2023 只有舊口徑，故 C、D 欄（舊口徑）只用來算 FY2024 的同口徑年增，E–G 欄（新口徑）進標準損益表。FY2025、FY2026 欄連到 In 頁 G 節與 Q 頁（已核實）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023（舊口徑）", "FY2024（舊口徑）", "FY2024（新口徑）", "FY2025", "FY2026", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        R[k] = r
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k in ("eps", "g_eps") else "#,##0.0000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 8, "Verified"); H.cell(r, 9, SRC_H + "；FY2025、FY2026 欄為連結" if isinstance(vals[-1], str) else SRC_H)
        r += 1
    r += 1
    H.cell(r, 1, "核對（單位：十億美元；公司以百萬美元揭露、四捨五入，加總可有 ±0.002 的捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("毛利 − 營業費用 − 營業利益（non-GAAP）", "={c}{gp}-{c}{opx}-{c}{oi}"),
                     ("營業利益 − 利息與其他 − 所得稅 + 少數股權調整 − 淨利：少數股權調整＝此列（應為小額正數；FY2025、FY2026 與 Q 頁一致）", "={c}{ni}-({c}{oi}-{c}{int}-{c}{tax})")):
        H.cell(r, 2, labl)
        for c in "CDEFG":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.0000"
        r += 1
    H.cell(r, 2, "核對：EPS（淨利 − 優先股股利）÷ 稀釋股數 − 公告 EPS（捨入差）")
    for c in "CDEFG":
        H[f"{c}{r}"] = f"=({c}{R['ni']}-{c}{R['pref']})/{c}{R['sh']}-{c}{R['eps']}"; H[f"{c}{r}"].number_format = "0.000"
    R["chk2"] = r; r += 1
    H.cell(r, 2, "核對：FY2026 營業利益（Q 頁）− 公司揭露 1,456.9 百萬美元（應為 0）"); H[f"G{r}"] = f"=G{R['oi']}-1.4569"; H[f"G{r}"].number_format = "0.0000"; r += 1
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 62
    for c in "CDEFG": H.column_dimensions[c].width = 16
    H.column_dimensions["I"].width = 90
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJK"      # 本表欄；歷史 C–E＝FY2024–FY2026（Hist E–G）；預測 F–H＝PL 情境 B 區塊 E–G；I–K＝同區塊 H–J（延伸期）
HC = {"C": "E", "D": "F", "E": "G"}
FLAGY = {"F": 2027, "G": 2028, "H": 2029, "I": 2029, "J": 2029, "K": 2029}   # 預測欄依賴的 Siliconomics 日曆年（FY(t) 含 hw × CY(t)；FY2030–32 延伸期沿用 FY2029 之後的路徑）


def pl_standard(wb, HR, CR):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    g = lambda label: lab(P, label, rb, exact=True)
    R = dict(dc=g("資料中心"), com=g("通訊"), ind=lab(P, "工業（FY2027 起排除", rb), rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), int=g("利息與其他淨支出"),
             ni=g("歸屬 Coherent 淨利"), sh=g("稀釋股數"), eps=g("EPS"), hdr=lab(P, "項目", rb, exact=True))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.4.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋未來六年（FY2027–FY2032，情境 B 基準；FY2027 含 Q1 實績與 Q2 指引）。US$B；EPS 美元。"
                 "只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。Coherent 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股票薪酬、收購無形資產攤銷、重組與整合費用等；公司於 FY2025 起調整定義，歷史欄已用新口徑）。"
                 "其他收入（淨額）＝−利息與其他淨支出（non-GAAP）。GAAP 合計（毛利、營業利益、淨利、EPS）列在本表下方對照，只列實績；預測不估 GAAP。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2027E–FY2029E＝情境 B 由下而上（FY2027 含 Q1 實績與 Q2 指引）；FY2030E–FY2032E＝延伸期（收斂法只推營收與淨利，營業利益＝（淨利 − 少數股權調整）÷（1 − 稅率）＋利息與其他，與 DCF 頁 EBIT 的起點相同、未扣股票薪酬）。"
                 "延伸期毛利率與營業費用率沿用 FY2029E，據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）。歷史淨利年增 FY2024 用舊口徑同口徑計算（見說明欄）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJK":   # 預測欄標題：依賴 Siliconomics 對帳中年份者加「（對帳中）」（旗標在 Coord 頁）
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
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    fc = lambda k: [f"=PL!{c}{R[k]}" for c in "EFG"]
    ex = lambda k: [f"=PL!{c}{R[k]}" for c in "HIJ"]
    row("rev", "營收", "US$B", hh("rev") + fc("rev") + ex("rev"), M, "FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT,
        "FY2024A 基期＝FY2023 營收 5.160（Hist 頁）；FY2026A 對 FY2025A 含已出售事業")
    rc = r
    row("cogs", "營業成本（COGS）", "US$B", [f"={c}{L['rev']}-{c}{rc + 1}" for c in "CDEFGH"] + [f"={c}{L['rev']}*(1-{p}{rc + 2})" for p, c in (("H", "I"), ("I", "J"), ("J", "K"))], M,
        "FY2024A–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − 前一年毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp") + [f"={c}{L['rev']}-{c}{rc}" for c in "IJK"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝前一年（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx") + [f"={c}{L['rev']}*{p}{rc + 3}/{p}{L['rev']}" for p, c in (("H", "I"), ("I", "J"), ("J", "K"))], M,
        "non-GAAP，不含股票薪酬；FY2030E–FY2032E＝營收 × 前一年營業費用率")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "US$B", hh("oi") + fc("oi") + [f"=({c}{rc + 10}-nci)/(1-tx)+int_f" for c in "IJK"], M, "FY2030E–FY2032E 由淨利反推（同 DCF 頁起點；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝−利息與其他淨支出", "US$B", [f"=-Hist!{HC[c]}{HR['int']}" for c in "CDE"] + [f"=-PL!{c}{R['int']}" for c in "EFG"] + ["=-int_f"] * 3, M, "non-GAAP；歷史為公司揭露的利息與其他淨支出（含匯兌以外之其他）")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", [f"=Hist!{HC[c]}{HR['tax']}" for c in "CDE"] + [f"={c}{L['pti']}-{c}{rc + 10}+{c}{rc + 9}" for c in "FGHIJK"], M, "歷史＝non-GAAP 所得稅；預測＝稅前 − 淨利 ＋ 少數股權調整（模型稅率 tx）")
    row("nci", "少數股權調整（加回；歷史＝淨利 −（稅前 − 所得稅））", "US$B", [f"={c}{rc + 10}-({c}{L['pti']}-{c}{L['tax']})" for c in "CDE"] + ["=nci"] * 6, M, "SiC 子公司少數股權分攤虧損加回；預測期＝In 頁 nci")
    assert L["nci"] == rc + 9, (L["nci"], rc)
    row("ni", "稅後淨利（歸屬 Coherent）", "US$B", hh("ni") + fc("ni") + ex("ni"), M)
    assert L["ni"] == rc + 10
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", hh("sh") + fc("sh_") if False else hh("sh") + [f"=PL!{c}{R['sh']}" for c in "EFGHIJ"], "#,##0.0000")
    row("eps", "EPS（完全稀釋；non-GAAP）", "US$", hh("eps") + [f"=PL!{c}{R['eps']}" for c in "EFGHIJ"], "0.00", "歷史＝公司公告值（FY2024 1.21、FY2025 3.53、FY2026 5.61）；預測＝（淨利 − 優先股股利）÷ 稀釋股數（PL 頁）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=Hist!D{HR['ni']}/Hist!C{HR['ni']}-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT,
        "FY2024A＝舊口徑同口徑（FY2024 原新聞稿 ÷ FY2023）；FY2025A 起為新口徑（FY2025A 對 FY2024A 新口徑重編值）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_gp", "GAAP：毛利", "US$B", hh("g_gp"), M, "含股票薪酬、收購無形資產攤銷（毛利內）、整合費用")
    row("g_oi", "GAAP：營業利益", "US$B", hh("g_oi"), M, "含股票薪酬、收購無形資產攤銷、重組、減損")
    row("g_ni", "GAAP：歸屬 Coherent 淨利", "US$B", hh("g_ni"), M, "FY2024 虧損，FY2025 近損益兩平，FY2026 轉盈")
    row("g_eps", "GAAP：EPS（完全稀釋，公司公告）", "US$", hh("g_eps"), "0.00")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2024A–FY2029E；歷史有 ±0.001 的捨入差）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:6]] + [None] * 3, "0.000")
    row("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "US$B",
        [None] * 6 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "IJK"], "0.000")
    row("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 6 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "IJK"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 ＋ 少數股權調整 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}+{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − PL 情境 B（FY2027E–FY2032E）；FY2025A、FY2026A − In 頁", "US$B",
        [None, f"=D{L['ni']}-a25_ni_in", f"=E{L['ni']}-a26_ni_in"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJK", "EFGHIJ")], "0.000")
    row("c5", "核對：EPS − PL 情境 B（FY2027E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHIJK", "EFGHIJ")], "0.000")
    P.column_dimensions["A"].width = max(P.column_dimensions["A"].width or 0, 66)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 14)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, gaap=r_gaap, chk=r_chk, start=s0, R=R)
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
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.4.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年 EBIT 與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027 視為錨點年不動（含 FY2027 Q1 實績與 Q2 指引、三情境幾乎相同）。反解值超出合理範圍時照實列出，不外推")
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
                     T(" 美元；要支持現價 "), 'TEXT(px,"0.00")', T(" 美元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大為","縮小為")', f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<B{A["w0"]},"降到","升到")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面"),
                     f'IF(B{A["s1"]}>=1,"放大","縮小")', f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0.0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0.0")', T("）；或 EBIT 率"),
                     f'IF(B{A["m1"]}>B{A["m0"]},"提高","降低")', T("到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × Coherent 過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T("）。換言之，市場目前給 Coherent 的倍數是同業的 "),
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
    SE, V, P = wb["Sens"], wb["Val"], wb["PL"]
    J = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")               # B、C、D＝情境 A、B、C
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    e_ = {}
    for s in "ABC":   # 各情境 12 個月後前瞻 EPS（與 Val 頁同式：8／12 × FY2028 ＋ 4／12 × FY2029；PL 頁摘要區 F、G 欄）
        h = lab(P, f"情境 {s}", 1, exact=True); er = lab(P, "EPS", h, exact=True)
        e_[s] = f"(8/12*PL!$F${er}+4/12*PL!$G${er})"
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        src.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v0.4.5 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝整本重算快照（cohr/sens_cohr.py：三情境同時套用 In 頁低／高欄，重算 DCF、Comparable 與加權）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用上方「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    base_f = f"{tp}"
    frame = [
        ("同業比較法的相對倍數調整係數 夾值下限 0.5／上限 2.0（基準＝歷史倍數比 約 1.1）",
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
        SE.cell(r, 4, d["低"]["tp"]).number_format = "#,##0.00"; SE.cell(r, 5, d["高"]["tp"]).number_format = "#,##0.00"
        SE.cell(r, 4).font = BLUE; SE.cell(r, 5).font = BLUE
        SE.cell(r, 2, f"=D{r}/{base_f}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{base_f}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, "整本重算快照（cohr/sens.json）；覆寫 In 頁低／高欄，三情境同時套用"); r += 1
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
    SE.cell(r, 3, f"快照基準加權目標價 {J['_base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 cohr/sens_cohr.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：三層協調——讀取層級、C0（實績）、對帳中年份與損益表年度標示 ----------
def coord_c0(wb, SIL):
    C, I, Q, U = wb["Coord"], wb["In"], wb["Q"], wb["Units"]
    r = C.max_row + 2
    r_flag = lab(C, "對帳中旗標")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    net_file = src_in("mpl_26").split("｜")[0].split("/")[-1]
    inl_file = src_in("inl_26").split("｜")[0].split("/")[-1]
    # ---- F. 讀取層級 ----
    C.cell(r, 1, "F. 讀取層級（三層基準協調規則 v1.4.2；共用工作單 2026-11-14）：本模型讀哪一層、哪一版。自層級讀入的值只列在本節，不進 C0 判定").font = BOLD; f_title = r; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("AI 可插拔光模組市場、AI CPO 光學市場、AI IT 資本支出 CY2026–CY2029（mpl_*、mcp_*、itc_*；情境 A，B／C 讀同字母 CPO 情境列）", "市場層 markets/network（其輸入由 Siliconomics 行業變數 M2、M8、M10 供給）", net_file,
         "資料中心營收＝可插拔市場 × 有效份額 × 單價係數＋CPO 市場 × CPO 份額", '="CY2026 可插拔 "&TEXT(mpl_26,"0.00")&"、CPO "&TEXT(mcp_26,"0.00")&"；CY2027 可插拔 "&TEXT(mpl_27,"0.00")&"、CPO "&TEXT(mcp_27,"0.00")&"（US$B）"'),
        ("CPO 光學金額份額 CY2027–CY2029（cpo_27–cpo_29）", "Siliconomics（M12）", vtxt, "CPO 光學市場 × Coherent 份額；資料中心營收的 CPO 部分", '=TEXT(cpo_27,"0%")&"／"&TEXT(cpo_28,"0%")&"／"&TEXT(cpo_29,"0%")&"（Assumed）"'),
        ("情境顆數係數 CY2028／CY2029（um28_*、um29_*；情境 B＝基準）", "Siliconomics（M2b，經台積電個股模型 In 頁轉入；A、C 為個股自訂）", vtxt, "市場值 × 係數", '="情境 B："&TEXT(um28_B,"0.00")&"／"&TEXT(um29_B,"0.000")&"；A："&TEXT(um28_A,"0.00")&"／"&TEXT(um29_A,"0.00")&"；C："&TEXT(um28_C,"0.00")&"／"&TEXT(um29_C,"0.00")'),
        ("中際旭創可插拔有效份額 CY2026–CY2029（inl_*；加總檢查用）", "個股模型 innolight（光模組份額權威）", inl_file, "Units 頁：兩家份額合計須低於 100%", '="CY2026 "&TEXT(inl_26,"0.0%")&"；CY2029 "&TEXT(inl_29,"0.0%")'),
        ("Wattonomics 錨定係數 wta_27–wta_29（Sens 案例；兩讀數建置時改放另一側讀數比值）", "Wattonomics（經 Siliconomics SIF_WTAnchor_Base）", vtxt, "Sens 頁「Wattonomics 錨定」案例；C4 另一側讀數整本重建", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")'),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics（行業變數 M30）", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中；比值 0.618／0.665／0.665"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（專案規則第 3 條）。本模型的有效份額（CY2026 校準）、份額相對值、單價係數是 Coherent 自己的變數，不在此節。").font = Font(italic=True); r += 2
    # ---- G. C0（實績） ----
    C.cell(r, 1, "G. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調）").font = BOLD; g_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝Coherent 8-K 已公布的「資料中心與通訊（DCC）」分部營收（公司只揭露合計，資料中心與通訊各自金額未揭露；公司指引不算實績）。層級側＝同口徑的層級隱含值＝markets/network 的 AI 可插拔光模組市場 × "
                 "Coherent 有效份額（資料中心）＋通訊（層級沒有通訊市場，取本模型 Q 頁推導值，屬 Derived、非層級值）。時差：層級值是日曆年 CY2026；已公布期間＝FY2026 Q3、Q4 與 FY2027 Q1＝2026-01～2026-09（9 個月）。"
                 "主讀數＝9 個月累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（FY2027 Q1，2026-07～09）× 4 vs 層級 CY2026（約等於該季中點 2026-08 的年化速率；"
                 "成長年中，最近一季落在年度後段時偏高，不得寫成夾住真值）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只有一個超過 → 時差邊界（參考）；都在內 → 一致。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); g_note1 = r; r += 1
    C.cell(r, 1, "同源校準（C0 第 2 點補充）：層級側的 Coherent 有效份額由本模型校準——CY2026 有效份額＝CY2026 資料中心 ÷ 可插拔市場，而 CY2026 資料中心是用本模型 FY2026 Q3、Q4、FY2027 Q1 實績與 Q2 假設推得。"
                 "市場值乘回份額時消去，層級隱含資料中心恆等於本模型 CY2026 資料中心（下列核對列＝0）。所以這張表實際檢驗的只有時間分攤（已公布 9 個月占 CY2026 全年的比例）與最近一季年化，"
                 "不檢驗層級市場值或份額、不能裁決層級對錯：判定標「參考（同源校準）」、不計入對帳中；C0 結論＝無可比實績（層級值同源）。要獨立檢驗需要公司揭露資料中心（不含通訊）營收，或獨立的 Coherent 份額估計；兩者均未找到。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); g_note2 = r; r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.0000"):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); IN[key] = r; r += 1
    inp("q3", "DCC 營收：FY2026 Q3（公司 8-K）", "=q3_dcc", "US$B", "2026-01-01～2026-03-31", f"In 頁 q3_dcc；FY2026 Q4 新聞稿 Segment Revenues 表 Q3 欄 1,361.6 百萬美元（SEC EDGAR 8-K Ex.99.1，2026-08；網址見 Hist 頁來源欄）", "Verified", GREEN)
    inp("q4", "DCC 營收：FY2026 Q4（公司 8-K）", "=q4_dcc", "US$B", "2026-04-01～2026-06-30", f"In 頁 q4_dcc；FY2026 Q4 新聞稿 Segment Revenues 表 1,615.0 百萬美元（SEC EDGAR 8-K Ex.99.1，2026-08；網址見 Hist 頁來源欄）", "Verified", GREEN)
    inp("q1", "DCC 營收：FY2027 Q1（公司 8-K）", "=q1_dcc", "US$B", "2026-07-01～2026-09-30", f"In 頁 q1_dcc；FY2027 Q1 新聞稿 Table 5 Segment Revenues 1,951.7 百萬美元（SEC EDGAR 8-K Ex.99.1，2026-11-04；網址見 Hist 頁來源欄）", "Verified", GREEN)
    inp("mo", "已公布月數（FY2026 Q3、Q4、FY2027 Q1）", 9, "月", "2026-01～2026-09", "三季＝9 個月（日曆年 2026 的前三季）", "Derived", BLUE, fmt="0")
    inp("m", "層級：AI 可插拔光模組市場 CY2026（markets/network）", "=mpl_26", "US$B", "CY2026（日曆年）", f"In 頁 mpl_26；{net_file}（Mkt 頁「其中可插拔」）", "Derived（層：顆數 × 每顆光學內容，行業變數 Assumed）", GREEN)
    inp("sh", "Coherent 可插拔有效份額 CY2026（本模型校準）", f"=Units!C{lab(U, 'Coherent 可插拔有效份額')}", "%", "CY2026", "Units 頁；＝CY2026 資料中心 ÷ 可插拔市場（In 頁 B 節說明）：本模型以公司實績校準，不是層級值", "Derived（同源校準）", GREEN, fmt="0.0%")
    inp("cp", "層級：AI CPO 光學市場 CY2026 × Coherent CPO 份額（CY2026 視為 0）", "=mcp_26*0", "US$B", "CY2026", "In 頁 mcp_26；CY2026 CPO 營收視為 0（PL 頁 A 節 Coherent 占 CPO 光學金額 CY2026＝0）", "Assumed", GREEN)
    inp("dc", "層級隱含：資料中心 CY2026（可插拔市場 × 份額＋CPO）", f"=B{r - 3}*B{r - 2}+B{r - 1}", "US$B", "CY2026（日曆年）", "上三列相乘加總", "Derived", fmt="#,##0.0000")
    inp("com", "通訊 CY2026（本模型 Q 頁推導，非層級值）", f"=Q!D{lab(Q, '通訊（電信', 1)}+Q!E{lab(Q, '通訊（電信', 1)}+Q!F{lab(Q, '通訊（電信', 1)}+Q!G{lab(Q, '通訊（電信', 1)}", "US$B", "CY2026＝FY2026 Q3＋Q4＋FY2027 Q1＋Q2（Q2 為指引口徑）",
        "Q 頁「通訊」列；由 DCC 合計與法說會季增率聯立推導（公司未揭露通訊金額）", "Derived（非公司揭露）", GREEN)
    inp("L", "層級隱含：DCC CY2026（資料中心層級隱含＋通訊）", f"=B{IN['dc']}+B{IN['com']}", "US$B", "CY2026（日曆年）", "上兩列加總", "Derived")
    inp("chk", "核對：層級隱含資料中心 − 本模型 CY2026 資料中心（應為 0；同源校準的恆等式）", f"=B{IN['dc']}-dc_cy26", "US$B", "CY2026", "Q 頁 B25（具名 dc_cy26）", "Derived", fmt="0.000000")
    inp("g", "參考：FY2027 Q2 公司指引營收中點（不判定）", f"=(Q!D{lab(Q, '營收', 36)}+Q!E{lab(Q, '營收', 36)})/2", "US$B", "FY2027 Q2（2026-10～12；指引）", "Q 頁 Q2 指引低、高（公司 8-K 2026-11-04；DCC 分部指引未揭露）", "Interested-party", GREEN)
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    pct = '+0.0%;-0.0%'
    b = IN
    C.cell(r, 1, "主讀數：DCC 營收，2026-01～09 累計實績 vs 層級 CY2026 × 9 ÷ 12")
    C.cell(r, 2, f"=B{b['q3']}+B{b['q4']}+B{b['q1']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%","在 ±15% 內")')
    C.cell(r, 6, "US$B；公司實績＝DCC 三季加總；層級＝CY2026 全年 × 9 ÷ 12")
    C.cell(r, 7, "公司：Coherent 8-K（Q3、Q4 見 FY2026 Q4 新聞稿；FY2027 Q1 見 Table 5）；層級：markets/network × 本模型校準份額（同源）＋通訊（Q 頁，Derived）")
    C.cell(r, 8, "Verified（公司）／Derived（層級隱含、線性時差）")
    C.cell(r, 9, f'="層級全年值含 CY2026 第四季（2026-10～12，模型 Q2 假設）；2026 年成長快、後段季度較大，線性分攤使層級值偏高、主讀數偏低。層級隱含全年 "&TEXT(C{r}/B{b["mo"]}*12,"0.00")&" 十億美元"'); rm = r; r += 1
    C.cell(r, 1, "第二讀數：DCC 營收，最近一季（FY2027 Q1）× 4 vs 層級 CY2026")
    C.cell(r, 2, f"=B{b['q1']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%","在 ±15% 內")')
    C.cell(r, 6, "US$B；FY2027 Q1（2026-07～09）單季 × 4（年化）；層級＝CY2026 全年")
    C.cell(r, 7, "公司：Coherent 8-K FY2027 Q1 新聞稿 Table 5；層級同上")
    C.cell(r, 8, "Verified（公司）／Derived（層級隱含、年化）")
    C.cell(r, 9, "最近一季中點約 2026-08；年化速率與 CY2026 全年平均相比位於後段，成長年中偏高"); rs = r; r += 1
    C.cell(r, 1, "C0（實績）判定：DCC 營收 CY2026").font = BOLD
    C.cell(r, 5, f'=IF(AND(ABS(D{rm})>0.15,ABS(D{rs})>0.15,SIGN(D{rm})=SIGN(D{rs})),"參考（同源校準）：兩讀數同向超過 ±15%","參考（同源校準）：不判定")').font = BOLD
    C.cell(r, 6, "同源校準：判定標參考、不計入對帳中；機械結果見主讀數與第二讀數列")
    C.cell(r, 9, f'="結論：無可比實績（層級值同源）。機械結果：主讀數 "&TEXT(D{rm},"+0.0%;-0.0%")&"、第二讀數 "&TEXT(D{rs},"+0.0%;-0.0%")&"，"&IF(AND(ABS(D{rm})<=0.15,ABS(D{rs})<=0.15),"兩者都在 ±15% 內（若層級值獨立，會判一致）","有讀數超過 ±15%")&"；只表示已公布 9 個月與最近一季年化和本模型全年路徑相容，不是對層級市場值或份額的獨立檢驗。缺的獨立揭露：公司資料中心（不含通訊）營收、獨立的 Coherent 份額"')
    rv = r; r += 1
    G1 = r - 1
    r += 1
    C.cell(r, 1, "層級之間與外部估計（參考；不計入 Coherent 的 C0）：市場層 markets/network 的 CY2026 AI 光學市場 vs LightCounting（外部研調）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "市場層", "外部估計", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    L0 = r
    C.cell(r, 1, "AI 光學市場 CY2026（可插拔＋CPO；LightCounting AI 叢集光學）")
    C.cell(r, 2, "=mpl_26+mcp_26").number_format = "#,##0.00"; C.cell(r, 3, 26).number_format = "#,##0.00"; C.cell(r, 3).font = BLUE
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "層間／外部參考（不判定、不計入）")
    C.cell(r, 6, "US$B；市場層 Mkt 頁光學合計 vs LightCounting（含 scale-up 光學的口徑差見市場層）")
    C.cell(r, 7, f"{net_file}（Mkt 頁第 39 列）；LightCounting 2026-05 光學廠商展望（見 Sources 頁）")
    C.cell(r, 8, "Derived（市場層）／Interested-party（外部估計）"); C.cell(r, 9, "外部估計只列參考；差距屬市場層與外部估計之間，不是 Coherent 自己的量，不開對帳單")
    L1 = r; r += 2
    # ---- H. 對帳中年份與損益表年度標示 ----
    C.cell(r, 1, "H. 對帳中年份與損益表年度標示（C4；FY(t)＝FY(t) 含 hw × CY(t)，FY2030E–FY2032E 為 FY2029E 之後的延伸期，依賴 CY2029）").font = BOLD; r += 1
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in range(2027, 2033)], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年")
    dep = [2027, 2028, 2029, 2029, 2029, 2029]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2027E–FY2032E）")
    for j, fy in enumerate(range(2027, 2033)):
        C.cell(r, 2 + j, f'=IF({"$"+"BCDE"[dep[j]-2026]+"$"+str(r_flag)}=1,"FY{fy}E（對帳中）","FY{fy}E")')
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
    coef, pmed = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數")
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    rb = lab(P, "情境 B：")
    pr = lambda lbl: lab(P, lbl, rb, exact=True)
    rev_b, dc_b, com_b = pr("營收"), pr("資料中心"), pr("通訊")
    ind_b = lab(P, "工業（FY2027 起排除", rb)
    reps27, reps28, rrev27 = lab(RC, "FY2027 EPS vs 共識"), lab(RC, "FY2028 EPS vs 共識"), lab(RC, "FY2027 營收 vs 共識")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, SEN['first'])}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.4.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("。現金流折現只有 "), f'TEXT({dcf},"0")',
        T(" 美元，同業比較法 "), f'TEXT({cadj},"0")', T(" 美元（已含相對倍數調整 "), f'TEXT({coef},"0.00")', T(" 倍），上檔主要來自後者。"))
    put("rec", T("投資建議「"), rec, T("」依加權目標價判定，且取決於同業比較法：不做相對倍數調整時加權目標價 "), f'TEXT({noadj},"0")', T(" 美元（"), f'TEXT({noadj}/px-1,"+0%;-0%")', T("），評等"), REC5(f"({noadj}/px-1)"),
        T("；現金流折現權重提高到七成時為 "), f'TEXT(0.7*{dcf}+0.3*{cadj},"0")', T(" 美元（"), f'TEXT((0.7*{dcf}+0.3*{cadj})/px-1,"+0%;-0%")', T("），評等"), REC5(f"((0.7*{dcf}+0.3*{cadj})/px-1)"),
        T("。評等依上檔空間五級固定規則；現價隱含的條件見反向 DCF。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0.00")', T(" 美元＝現金流折現 "), f'TEXT({dcf},"0.00")', T(" 與同業比較法 "), f'TEXT({cadj},"0.00")', T(" 各半；比現價 "), 'TEXT(px,"0.00")', T(" "), f'IF({up}<0,"低","高")', f'TEXT(ABS({up}),"0%")',
        T("，評等"), rec, T("；市場共識 "), 'TEXT(cons_tp,"0")', T("。"))
    put("o2", T("現價隱含：現金流折現要支持現價，折現率須"), f'IF({rv("w1")}<wacc,"降到","升到")', f'TEXT({rv("w1")},"0.0%")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收"), f'IF({rv("s1")}>=1,"放大","縮小")', f'TEXT({rv("s1")},"0.00")',
        T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0.0")', T("，模型 "), f'TEXT({rv("rv0")},"0.0")', T(" 十億美元）。"))
    put("o3", T("盈餘高於共識：情境 B 的 FY2027 EPS "), f'TEXT(Recon!E{reps27},"0.00")', T(" 美元比共識 "), f'TEXT(Recon!B{reps27},"0.00")', T(" 高 "), f'TEXT(Recon!F{reps27},"0%")', T("，FY2028 EPS "), f'TEXT(Recon!E{reps28},"0.00")', T(" 比共識 "),
        f'TEXT(Recon!B{reps28},"0.00")', f'IF(Recon!F{reps28}>=0," 高 "," 低 ")', f'TEXT(ABS(Recon!F{reps28}),"0%")', T("；共識只作對照，未用於校準。"))
    put("o4", T("最敏感：毛利率路徑（低案 "), f'TEXT({SB("non-GAAP 毛利率", "B")},"+0%;-0%")', T("）、有效份額（"), f'TEXT({SB("Coherent 可插拔有效份額", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("Coherent 可插拔有效份額", "C")},"+0%;-0%")',
        T("）、AI 光模組市場 ×0.8／×1.2（"), f'TEXT({SB("AI 光模組市場", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("AI 光模組市場", "C")},"+0%;-0%")', T("）；同業比較法的倍數調整夾值 0.5／2.0 為 "),
        f'TEXT({SB("同業比較法的相對倍數調整係數", "B")},"+0%;-0%")', T("／"), f'TEXT({SB("同業比較法的相對倍數調整係數", "C")},"+0%;-0%")', T("（加權目標價相對基準）。"))
    put("o5", T("對帳中：產業部署 GW 換另一側讀數後，加權目標價 "), f'TEXT(Coord!C{CO["tp_row"]},"0.00")', T(" 美元、評等"), f'Coord!C{CO["rec_row"]}', T("（現行 "), f'TEXT({tp},"0.00")', T("）。C0 實績：DCC 營收 2026 年前三季 "), f'TEXT(Coord!B{CO["rm"]},"0.0")',
        T(" 對層級隱含 "), f'TEXT(Coord!C{CO["rm"]},"0.0")', T("（"), f'TEXT(Coord!D{CO["rm"]},"+0%;-0%")', T("），但份額由本模型校準，無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：資料中心＝AI 可插拔光模組市場 × 情境顆數係數 × Coherent 有效份額（CY2026 校準；份額相對值與單價係數為假設）＋AI CPO 光學市場 × CPO 份額，日曆年再以上半年比重換成會計年度；"
                 "通訊與工業各以固定年增率推算；三者相加為合併營收。FY2028：資料中心 "), f'TEXT(PL!F{dc_b},"0.0")', T("、通訊 "), f'TEXT(PL!F{com_b},"0.0")', T("、工業 "), f'TEXT(PL!F{ind_b},"0.0")', T("，合計 "), f'TEXT(PL!F{rev_b},"0.0")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型讀市場層 markets/network 的光學市場、Siliconomics 的 CPO 份額與情境顆數係數（Coord 頁 F 節）。C0 改以公司實績對層級隱含值（Coord 頁 G 節）：DCC 營收 2026 年前三季累計 "), f'TEXT(Coord!B{CO["rm"]},"0.00")',
        T(" 十億美元對層級隱含依月數折算 "), f'TEXT(Coord!C{CO["rm"]},"0.00")', T("，差 "), f'TEXT(Coord!D{CO["rm"]},"+0.0%;-0.0%")', T("；最近一季年化差 "), f'TEXT(Coord!D{CO["rs"]},"+0.0%;-0.0%")', T("；層級份額由本模型校準，判定「"), f'Coord!E{CO["rv"]}',
        T("」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項。產業部署 GW 兩層差距另列參考、不計入；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3、三層協調（讀取層級、C0 實績）與兩讀數目標價（三層基準協調規則 v1.4.2 C4）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史）、PL 頁末標準損益表"
                  "（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B，Non-GAAP）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準（整本重算快照 cohr/sens.json）、"
                  "Coord 頁 A–E 節兩讀數目標價與 F–H 節（讀取層級、C0（實績）、損益表年度標示）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    RD["A1"].value = a1 + (f"；v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3、三層協調（Coord 頁 F–H 節：讀取層級、C0 實績對帳；Hist、RevDCF 頁；PL 標準損益表；Val 0.7；Sens 加權目標價相對基準），既有數值不變")


def apply(wb, here, SIL, CR):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("CO_OVERRIDE"):
        return "v045_post：略過（alt_run／sens_cohr 建置）"
    HR = hist_sheet(wb)
    PLA = pl_standard(wb, HR, CR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb, here)
    CO = coord_c0(wb, SIL)
    CO["tp_row"], CO["rec_row"] = CR["tp"], CR["rec"]
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    pass  # README A1 版號與說明由 COORD.stamp 處理；本版 v1.3 說明在 README 末列
    wb._v045 = dict(PL=PLA, RV=RVA, RM=RMA, SEN=SEN, CO=CO, VG=VG, HR=HR)
    return f"v045_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；Coord F–H 節；README v1.3 白話列 {len(RMA)} 列"
