# Kioxia v0.5.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Advantest v0.7.5、記憶體 SK 海力士 v0.7.7）、
# 三層協調讀取層級與 C0（含公司實績，規則 v1.4.2 C0 第 2 點）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.5.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2023–FY2025（repo 既有來源；找不到的列留空並標「未找到」）
#   PL 頁末：標準損益表（FY2023A–FY2025A＋FY2026E–FY2032E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（快照 kioxia/sens.json）
#   Coord 頁末：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）；Sources、README 頁末
import json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as GL

VER, BUILT = "0.5.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
ER1 = "https://links.sgx.com/1.0.0/corporate-announcements/DOATO9EIE4EEQ2Y8/898078_Earnings-Release-FY2026-1Q-EN.pdf"
SRC_Q2 = "Kioxia FY2026 第二季財報（2026-10-30；404K Research 法說整理、PANews／TechFlow 轉載；公司原文 PDF 本輪未重開）"


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
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
TX = lambda ref, f: f'TEXT({ref},"{f}")'
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h); c.font = BOLD; c.fill = HDR


# ---------- Hist：損益表歷史 ----------
# 值為 In 頁名稱公式（連格）或 repo 既有來源的常數；None＝未找到（本輪無法連網取公司年報原文，留空不補）
SA = "stockanalysis 季度／年度損益（TYO:285A；In 頁 p1–p4 列、M2 節 fc_da 來源）"
HIS = [  # (鍵, 項目, FY2023, FY2024, FY2025, 標記, 來源)
    ("rev", "營收", None, 1706.46, "=a25_rev", "Interested-party／Derived",
     "FY2024：" + SA + "（1,706,460 百萬日圓；In 頁 fc_da 來源）；FY2025：In 頁 a25_rev＝四季合計 2,337.6（公司 2,337.6，storagenewsletter 轉載 FY25 決算）；FY2023：未找到"),
    ("gp", "毛利", None, None, "=p1_gp+p2_gp+p3_gp+p4_gp", "Interested-party", "FY2025：In 頁 p1_gp–p4_gp 四季合計（" + SA + "）；FY2023、FY2024：未找到"),
    ("cogs", "營業成本", None, None, "=a25_rev-(p1_gp+p2_gp+p3_gp+p4_gp)", "Derived", "＝營收 − 毛利（FY2025）；FY2023、FY2024：未找到"),
    ("oi", "營業利益", None, None, "=p1_op+p2_op+p3_op+p4_op", "Interested-party／Verified", "FY2025：In 頁 p1_op–p4_op 四季合計 869.0（公司 870.4，取整差）；FY2023、FY2024：未找到"),
    ("pti", "稅前利益", None, None, None, "—", "未找到（FY2023–FY2025 稅前利益原文本輪未取得）"),
    ("tax", "所得稅", None, None, None, "—", "未找到"),
    ("ni", "稅後淨利（歸屬母公司）", None, 272.3, "=a25_ni", "Interested-party／Derived",
     "FY2024：MarketScreener 272.3（v0.4 In 頁 tp_roe 來源）；FY2025：In 頁 a25_ni＝四季合計 554.5（公司 554.0／MarketScreener 554.5）；FY2023：未找到"),
    ("eps", "EPS（分割後；口徑見來源）", None, None, "=a25_eps", "Interested-party", "FY2025：In 頁 a25_eps 341.4（MarketScreener，分割後）；FY2023、FY2024：未找到"),
]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.5.5 起；PL 頁「標準損益表」的歷史三年用）。會計年度 4–3 月：FY2023＝2023-04～2024-03、FY2024＝2024-04～2025-03、FY2025＝2025-04～2026-03。"
               "金額：十億日圓；EPS：日圓（2026-10-01 一股分三股後）。只用 repo 既有來源（In 頁與舊版來源）；空白＝未找到（本輪無法連網取公司年報原文，未補）")
    H["A1"].font = BOLD
    hdr(H, 3, ["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    r = 4; R = {}
    for k, labl, *vals, tag, src in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.0" if k != "eps" else "#,##0.0"
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        H.cell(r, 6, tag); H.cell(r, 7, src)
        R[k] = r; r += 1
    H.cell(r, 1, "h_opx"); H.cell(r, 2, "營業費用（＝毛利 − 營業利益）"); H["E" + str(r)] = f"=E{R['gp']}-E{R['oi']}"; H["E" + str(r)].number_format = "#,##0.0"
    H.cell(r, 6, "Derived"); H.cell(r, 7, "FY2025＝毛利 − 營業利益（Q 頁同口徑）"); R["opx"] = r; r += 1
    H.cell(r, 1, "h_sh"); H.cell(r, 2, "股數（＝淨利 ÷ EPS；百萬股，分割後）"); H["E" + str(r)] = f"=E{R['ni']}*1000/E{R['eps']}"; H["E" + str(r)].number_format = "#,##0.0"
    H.cell(r, 6, "Derived"); H.cell(r, 7, "FY2025 平均股數約 1,624 百萬股（In 頁 a25_eps 說明）"); R["sh"] = r; r += 2
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    H.cell(r, 2, "FY2025：營收 − 營業成本 − 毛利"); H[f"E{r}"] = f"=E{R['rev']}-E{R['cogs']}-E{R['gp']}"; H[f"E{r}"].number_format = "0.000"; R["chk1"] = r; r += 1
    H.cell(r, 2, "FY2025：營收 − Q 頁 FY2025 營收"); H[f"E{r}"] = f"=E{R['rev']}-Q!F4"; H[f"E{r}"].number_format = "0.000"; R["chk2"] = r; r += 1
    H.cell(r, 2, "FY2025：營業利益 − Q 頁 FY2025 營業利益"); H[f"E{r}"] = f"=E{R['oi']}-Q!F6"; H[f"E{r}"].number_format = "0.000"; R["chk3"] = r; r += 1
    H.column_dimensions["A"].width = 10; H.column_dimensions["B"].width = 44
    for c in "CDE": H.column_dimensions[c].width = 12
    H.column_dimensions["G"].width = 110
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "C", "D": "D", "E": "E"}   # 標準損益表歷史欄 → Hist 欄（FY2023＝C、FY2024＝D、FY2025＝E）
PC4 = {"F": "D", "G": "E", "H": "F", "I": "G"}   # FY2026–29 → PL 頁推導區塊欄（FYC）
DCOL = {"F": "C", "G": "D", "H": "E", "I": "F", "J": "G", "K": "H", "L": "I"}   # → DCF 頁欄
# 對帳中旗標依賴的日曆年：FY(t)＝CY(t) 4–12 月＋CY(t+1) 1–3 月；FY2026 第四季（2027Q1）在另一側讀數建置中受縮放，故 FY(t) 依 CY(t＋1)；FY2030–32 為延伸期，沿用 CY2029
FLAGY = {"F": 2027, "G": 2028, "H": 2029, "I": 2029, "J": 2029, "K": 2029, "L": 2029}


def pl_standard(wb, HR, CR, X):
    P = wb["PL"]; R = X["PB"]["B"]; FYC = X["FYC"]; EXTC = X["EXTC"]; XB = X["EXT"]["B"]
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr)
    hcell = lambda k, c: wb["Hist"][f"{HC[c]}{HR[k]}"].value
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, "標準損益表（v0.5.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025）＋未來七年（FY2026–FY2032，情境 B 基準）。會計年度 4–3 月；十億日圓；EPS 日圓（分割後）。只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：公司公布的 GAAP 數字（非 non-GAAP；公司另公布加回訴訟損失準備、股份報酬的 non-GAAP 營業利益，本表不用）；EPS 為稀釋 EPS（模型口徑：期末流通股數＋稀釋股數）。"
                 "預測期模型推毛利、營業費用（固定＋變動）、營業利益與淨利（淨利＝（營業利益＋業外＋淨現金收益）×（1 − 稅率））；稅前利益＝淨利 ÷（1 − 稅率），所得稅＝稅前利益 − 淨利，其他收入淨額＝稅前利益 − 營業利益，三者皆為由模型淨利反推的 Derived 列，不進 EPS、現金流與目標價。"
                 "FY2026E 的第一、二季為實績（淨利為公司實績，故 FY2026 的稅前利益反推只是近似）。歷史欄空白＝未找到（Hist 頁）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷（1 − 稅率）− 業外 − 淨現金收益），毛利率沿用 FY2029，營業費用＝毛利 − 營業利益（推算）。")
    r_note2 = r; r += 1
    hdr(P, r, ["項目", "單位"] + YEARS + ["說明"])
    for col in "FGHIJKL":
        yr = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr}"&IF(Coord!${"BCDE"[FLAGY[col] - 2026]}${CR["flag"]}=1,"（對帳中）","")'
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
    hh = lambda k: [(f"=Hist!{HC[c]}{HR[k]}" if hcell(k, c) is not None else None) for c in "CDE"]
    fc = lambda k: [f"=PL!{PC4[c]}{R[k]}" for c in "FGHI"]
    hv = lambda k, c: hcell(k, c) is not None
    row("rev", "營收", "¥B", hh("rev") + fc("yrev") + [f"=PL!{EXTC[2030 + i]}{XB['rev']}" for i in range(3)], M, "FY2026E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2032E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [None, None, f"=E{L['rev']}/D{L['rev']}-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip("EFGHIJK", "FGHIJKL")], PCT, "FY2023、FY2024 的基期未找到")
    rc = r
    row("cogs", "營業成本（COGS）", "¥B", hh("cogs") + fc("ycogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "JKL"], M, "FY2030E 起＝營收 − 毛利")
    row("gp", "毛利", "¥B", hh("gp") + [f"={c}{L['rev']}-{c}{rc}" for c in "FGHI"] + [f"={c}{L['rev']}*{c}{rc + 2}" for c in "JKL"], M, "FY2026E–FY2029E＝營收 − 營業成本（PL 情境 B）；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [(f"={c}{rc + 1}/{c}{L['rev']}" if hv("gp", c) else None) for c in "CDE"] + [f"={c}{rc + 1}/{c}{L['rev']}" for c in "FGHI"] + [f"={p}{rc + 2}" for p in "IJK"], PCT, "FY2030E–FY2032E 沿用前一年")
    row("opx", "營業費用", "¥B", [(f"=Hist!{HC[c]}{HR['opx']}" if c == "E" else None) for c in "CDE"] + fc("yox") + [f"={c}{rc + 1}-{c}{rc + 4}" for c in "JKL"], M, "歷史＝毛利 − 營業利益；FY2026E–FY2029E 為模型固定＋變動；FY2030E 起＝毛利 − 營業利益（推算）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", "¥B", hh("oi") + fc("yop") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M, "FY2030E 起＝DCF 頁情境 B 的 EBIT")
    assert L["oi"] == rc + 4
    row("opm", "營業利益率", "%", [(f"={c}{L['oi']}/{c}{L['rev']}" if hv("oi", c) else None) for c in "CDE"] + [f"={c}{L['oi']}/{c}{L['rev']}" for c in "FGHIJKL"], PCT)
    row("oth", "其他收入（淨額；業外）", "¥B", [None] * 3 + [f"={c}{r + 1}-{c}{L['oi']}" for c in "FGHIJKL"], M, "＝稅前利益 − 營業利益；預測期為由淨利反推的 Derived 值；歷史未找到")
    row("pti", "稅前利益", "¥B", hh("pti") + [f"={c}{r + 2}/(1-tax)" for c in "FGHIJKL"], M, "預測期＝淨利 ÷（1 − tax）")
    row("tax", "所得稅", "¥B", hh("tax") + [f"={c}{L['pti']}-{c}{r + 1}" for c in "FGHIJKL"], M, "預測期＝稅前利益 − 淨利")
    row("ni", "稅後淨利（歸屬母公司）", "¥B", hh("ni") + fc("yni") + [f"=PL!{EXTC[2030 + i]}{XB['ni']}" for i in range(3)], M)
    row("nm", "淨利率", "%", [(f"={c}{L['ni']}/{c}{L['rev']}" if hv("ni", c) and hv("rev", c) else None) for c in "CDE"] + [f"={c}{L['ni']}/{c}{L['rev']}" for c in "FGHIJKL"], PCT)
    row("sh", "稀釋股數（百萬股，分割後）", "百萬股", [None, None, f"=Hist!E{HR['sh']}"] + [f"={c}{L['ni']}*1000/{c}{r + 1}" for c in "FGHIJKL"], M, "預測期＝淨利 ÷ EPS（PL 頁 EPS＝淨利 ÷（期末流通股數＋稀釋股數））")
    row("eps", "EPS（稀釋，分割後）", "¥", hh("eps") + fc("eps") + [f"=PL!{EXTC[2030 + i]}{XB['eps']}" for i in range(3)], "#,##0.0", "FY2025 為 MarketScreener 值（分割後；基本或稀釋口徑未載明）")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [None, None] + [f"={c}{L['yoy']}" for c in "EFGHIJKL"], PCT)
    row("r_gm", "毛利率", "%", [(f"={c}{L['gm']}" if hv("gp", c) else None) for c in "CDE"] + [f"={c}{L['gm']}" for c in "FGHIJKL"], PCT)
    row("r_opm", "營業利益率", "%", [(f"={c}{L['opm']}" if hv("oi", c) else None) for c in "CDE"] + [f"={c}{L['opm']}" for c in "FGHIJKL"], PCT)
    row("r_nm", "淨利率", "%", [(f"={c}{L['nm']}" if hv("ni", c) and hv("rev", c) else None) for c in "CDE"] + [f"={c}{L['nm']}" for c in "FGHIJKL"], PCT)
    row("r_tax", "有效稅率（所得稅 ÷ 稅前利益）", "%", [None] * 3 + [f"={c}{L['tax']}/{c}{L['pti']}" for c in "FGHIJKL"], PCT, "預測期＝In 頁 tax")
    row("r_niy", "淨利年增率（YoY）", "%", [None, None, f"=E{L['ni']}/D{L['ni']}-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip("EFGHIJK", "FGHIJKL")], PCT)
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "¥B", [(f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" if hv("gp", c) else None) for c in "CDE"] + [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "FGHIJKL"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "¥B", [None, None, f"=E{L['gp']}-E{L['opx']}-E{L['oi']}"] + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "FGHIJKL"], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "¥B", [None] * 3 + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in "FGHIJKL"], "0.000")
    row("c4", "核對：營業利益 − DCF 頁情境 B 的 EBIT（FY2026E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['oi']}-DCF!{DCOL[c]}{d_ebit}" for c in "FGHIJKL"], "0.000")
    row("c5", "核對：稅後淨利 − PL 情境 B（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['ni']}-PL!{PC4[c]}{R['yni']}" for c in "FGHI"] + [f"={c}{L['ni']}-PL!{EXTC[2030 + i]}{XB['ni']}" for i, c in enumerate("JKL")], "0.000")
    row("c6", "核對：營收 − PL 情境 B（FY2026E–FY2029E）／延伸表（FY2030E–FY2032E）", "¥B", [None] * 3 + [f"={c}{L['rev']}-PL!{PC4[c]}{R['yrev']}" for c in "FGHI"] + [f"={c}{L['rev']}-PL!{EXTC[2030 + i]}{XB['rev']}" for i, c in enumerate("JKL")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度 4–3 月，FY(t)＝CY(t) 4–12 月＋CY(t+1) 1–3 月；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY(t) 依 CY(t+1)（FY2026 的第四季＝2027Q1 在另一側讀數建置中受縮放；CY2026 非對帳中）；FY2030E–FY2032E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = lab(D, "評價日後占該年度比例"), lab(D, "折現年數（期中"), lab(D, "終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    fbv, pmed, dcf_tp, coef = vb("12 個月後的前瞻每股淨值（情境加權）"), vb("同業前瞻 P/B 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.5.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與現金流倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑）；"
                "Comparable 部分反解「12 個月後前瞻每股淨值（情境加權）× 前瞻 P/B＝現價」（記憶體慣例用 P/B）。FY2026（只計評價日後一季）與 FY2027 不動，營業利益率與現金流倍數只動 FY2028–FY2032。反解值超出合理範圍時照實列出，不外推")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    hdr(RV, r, ["項目", "值", "說明"]); r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note: RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（日圓）", "=px", "#,##0")
    kv("v0", "支持現價所需的評價日每股價值（日圓）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數")
    kv("nc", "淨現金滾到評價日（情境 B；¥B）", "=" + ncr, "#,##0.0", "Val 頁 0.1 節（同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.0000")
    kv("evr", "支持現價所需的企業價值 EV（¥B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；¥B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（日圓）", f"=DCF!$C${ps}", "#,##0")
    r += 1
    yh = r
    hdr(RV, r, ["項目（情境 B）"] + [f"FY{y}" for y in range(2026, 2033)]); r += 1
    Y = {}
    NC = "BCDEFGH"

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGHI"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（¥B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（¥B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "I" else "=0"), "0.0000")
    yr("k", "每 1 日圓營業利益對 EV 的貢獻（FY2026、FY2027 不動，記 0）", lambda dc, c: ("=0" if dc in "CD" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（¥B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（¥B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（¥B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（¥B）", f"=SUMPRODUCT(D{Y['rev']}:H{Y['rev']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（¥B）", f"=B{A['ev']}-SUMPRODUCT(D{Y['ebit']}:H{Y['ebit']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2026–FY2027 現值（¥B）", f"=B{Y['pv']}+C{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（¥B）", f"=SUM(D{Y['pv']}:H{Y['pv']})+DCF!I{nop}*H{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（¥B）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；¥B）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近界限，表示區間內無解）").font = BOLD; r += 1
    hdr(RV, r, ["次", "下界", "上界", "中點", "中點的評價日每股價值（日圓）"]); r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$I${fcf},DCF!$C${fr}:$I${fr},1/(1+{w})^DCF!$C${ty}:$I${ty})"
                f"+DCF!$I${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tn}+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.001"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "#,##0"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.1 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（日圓，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
    r += 1
    RV.cell(r, 1, "B. Comparable：支持現價所需的倍數（記憶體慣例：前瞻 P/B）").font = BOLD; r += 1
    hdr(RV, r, ["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"]); r += 1

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    RV.cell(r, 1, "12 個月後前瞻每股淨值（情境加權；日圓）"); RV.cell(r, 3, f"={fbv}").number_format = "#,##0"; A["fbv"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻 P/B（倍）", f"=px/{fbv}", f"=tp_comp_adj/{fbv}", "0.00")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻 P/B 中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻 P/B（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{fbv})", f"=tp_comp_adj/{fbv}", "0.00")
    RV.cell(r, 1, "同業前瞻 P/B 中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.00"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(" 日圓；現價 "), TX("px", "#,##0"), T(" 日圓只需企業價值為模型的 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素反解到現價：折現率 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）乘 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億日圓，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("）。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(" 日圓相當於 12 個月後前瞻每股淨值的 "), TX(f"B{A['pe1']}", "0.00"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("記憶體同業前瞻 P/B 中位數 "), TX(pmed, "0.00"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.00"),
                     T(" 倍（同業中位數 × 相對倍數調整係數 "), TX(coef, "0.00"), T("）。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70
    for c in "BCDEFGH": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    return A


# ---------- Val：估值並列四個目標價 ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    dcf, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    fbv, pmed, coef = vb("12 個月後的前瞻每股淨值（情境加權）"), vb("同業前瞻 P/B 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價', exact=True)}"
    px_lab = f"Val!$A${lab(V, '現價（')}"
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.5.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    hdr(V, r, ["目標價", "日圓", "相對現價", "說明"]); r += 1
    rows = [("① DCF 目標價（情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（同業前瞻 P/B；進加權）", cadj, F(T("記憶體同業前瞻 P/B 中位數 "), TX(pmed, "0.00"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻每股淨值（情境加權）"), TX(fbv, "#,##0"), T(" 日圓；係數為 1 時與市場主流同業法同值"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable "), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0"; V.cell(r, 4, "=" + px_lab); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    snap = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    tp_cell = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）', exact=True)}"
    t0 = lab(SE, "目標價參數（v0.5"); hr = t0 + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.5.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（kioxia/sens_kioxia.py），目標價參數取上方「目標價參數」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    hdr(SE, r, ["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", "低案：加權目標價（日圓）", "高案：加權目標價（日圓）", "快照基準加權目標價（日圓）"])
    r += 1
    first = r
    b = snap["_base"]["tp"]
    for case in snap["cases"]:
        k = case["key"]; z = snap["biz"][k]
        SE.cell(r, 1, case["label"]); SE.cell(r, 4, case["setting"] + "（整本重算快照）")
        for col, side, tcol in ((2, "低", 5), (3, "高", 6)):
            if side in z:
                c = SE.cell(r, tcol, z[side]["tp"]); c.font = BLUE; c.number_format = "#,##0"
                SE.cell(r, col, f"={'EF'[tcol - 5]}{r}/$G{r}-1").number_format = "+0.0%;-0.0%"
        c = SE.cell(r, 7, b); c.font = BLUE; c.number_format = "#,##0"
        r += 1
    for sr in src:
        SE.cell(r, 1, f"=Sens!A{sr}"); SE.cell(r, 4, f"=\"Sens 頁目標價參數表：\"&Sens!B{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
        r += 1
    last = r - 1
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 kioxia/sens_kioxia.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, dcfw=lab(SE, "DCF 權重", t0))


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, X):
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CW = wb["Coord"]; N = SIL.NAMES
    import openpyxl
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    NF = X["NAND_F"]; wn = openpyxl.load_workbook(os.path.join(ROOT, NF), data_only=True); A_ = wn["Ann"]
    ann = {}; sc = None
    for rr in range(1, A_.max_row + 1):
        a = str(A_.cell(rr, 1).value or "")
        if a.startswith("情境 ") and len(a) > 3 and a[3] in "ABC": sc = a[3]
        if a == "Kioxia 營收（年）": ann[sc] = (rr, [A_.cell(rr, 5 + j).value for j in range(4)])   # CY2026–CY2029（Ann E–H）
    assert set(ann) == {"A", "B", "C"}, ann
    UR, UCOL = X["UR"], X["UCOL"]; urB = UR[("rev", "B")]
    U26 = [UCOL[f"2026Q{q}"] for q in (1, 2, 3, 4)]
    assert U26 == ["G", "H", "I", "J"], U26
    kx4, kx5 = N["SIF_nd_rv_kx_4"], N["SIF_nd_rv_kx_5"]
    assert abs(kx4 - 5.96) < 1e-9 and abs(kx5 - 10.72) < 1e-9, (kx4, kx5)
    SEW = wb["Sens"]; w0 = lab(SEW, "Wattonomics 錨定（三層校準原則）"); r_dai = lab(SEW, "需求：AI 企業級 SSD（情境 B）", w0)
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.5.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    hdr(CW, r, ["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"]); r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（經 markets/nand 讀 nd_*、nq_*；本模型 In 頁直接讀 SIF_WTAnchor_Base；不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股（記憶體 NAND），讀 Siliconomics"),
          ("NAND 市場層成品（唯讀抄錄）", NF.split("/")[-1], f"{NF}（Mkt 頁 Kioxia 營收、位元與合資群份額；Ann 頁 AI 企業級 SSD 占比）", "markets/nand v0.3", "Units 頁藍字；市場層讀 Siliconomics M23 NAND 輸入"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁敏感度案例；不改基準"),
          ("CY2026 Kioxia 營收（市場層，情境 B；TrendForce 美元口徑）", f"=SUM(Units!G{urB}:J{urB})", "markets/nand Mkt 頁 Kioxia 營收（2026Q1–Q2＝Siliconomics nd_rv_kx_4／5；2026Q3 起＝市場 × 份額）", "markets/nand v0.3", "PL 頁營收＝本值 × 公司口徑係數 × 匯率（第三季起）")]
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
    CW.cell(r, 1, "本模型會計年度 4–3 月；下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。日曆季與會計年度的對應在 PL 頁（第一季＝4–6 月）。")
    ROW["caveat"] = r; r += 1
    hdr(CW, r, ["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定", "層級名稱與版本／口徑／證據等級／備註"])
    ROW["hdr2"] = r; r += 1
    first = r

    def line(label, unit, yr_, f_model, f_layer, fmt, note):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_model).number_format = fmt
        c = CW.cell(r, 5, f_layer); c.number_format = fmt
        if not (isinstance(f_layer, str) and f_layer.startswith("=")): c.font = BLUE
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
    nb = f"Siliconomics {SIL.VERSION}"
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        cs = [UCOL[f"{y}Q{q}"] for q in (1, 2, 3, 4)]
        line("Kioxia 營收（情境 B；本模型 Units 頁四季合計）vs 市場層 Ann 頁", "US$B", f"CY{y}", f"=SUM(Units!{cs[0]}{urB}:{cs[3]}{urB})", float(ann["B"][1][i]), "#,##0.000",
             f"markets/nand v0.3 Ann!{GL(5 + i)}{ann['B'][0]}（Kioxia 營收（年），情境 B）；Derived；Units 頁為 Mkt 頁唯讀抄錄")
    line("Kioxia NAND 營收（本模型 Units 頁）vs Siliconomics", "US$B", "2026Q1", f"=Units!G{urB}", kx4, "#,##0.000", f"{nb} nd_rv_kx_4（Derived：TrendForce 2Q26 ÷（1＋季增 79.9%））；經 markets/nand 讀入")
    line("Kioxia NAND 營收（本模型 Units 頁）vs Siliconomics", "US$B", "2026Q2", f"=Units!H{urB}", kx5, "#,##0.000", f"{nb} nd_rv_kx_5（Interested-party：TrendForce 2026-08-18 前五大排名）；經 markets/nand 讀入")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    last = r - 1
    ROW["first"], ROW["last"] = first, last
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 1
    CW.cell(r, 1, "無可比層級量的項目：Kioxia 終端市場分部（SSD 與儲存、智慧裝置、其他；層級只有全產業應用別位元，沒有 Kioxia 分部份額）；營業利益、淨利（層級不持有）；加速器顆數（讀自 markets/memory，非本模型輸入）。")
    r += 1
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準（經 markets/nand）與 Wattonomics 換算值兩層之間的差距，不是 Kioxia 自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。")
    r += 1
    hdr(CW, r, ["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"])
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    ratio = lambda y: f"{'BCDE'[(y - 2026)]}${CR['ratio']}"
    for i, y in enumerate((2027, 2028, 2029)):
        CW.cell(r, 1, "AI 企業級 SSD 位元需求（情境 B）：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "EB"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, f"=Sens!{'DEF'[i]}{r_dai}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{r}*{ratio(y)}").number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"markets/nand Ann 頁（Sens 頁 Wattonomics 錨定節抄錄）；Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；個股不裁決")
        r += 1
    ROW["ref2"] = r - 1
    r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t) 4–12 月＋CY(t+1) 1–3 月：FY2026 第四季起至 FY2029，及其後由 FY2029 收斂推得的 FY2030–FY2032 延伸期；CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 日圓（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 日圓（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（情境 B、C 的 AI 企業級 SSD 位元需求 × 另一側比值，經 MIN(需求, 供給) 傳到 Kioxia 位元與營收；情境 A 不縮放），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準；CY2030 沿用 CY2029）。只縮放 AI 企業級 SSD 位元需求；非 AI 應用、位元供給、價格路徑與份額未縮放，需求低於供給的季度才影響出貨。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110
    # ---- 4. C0 ----
    r += 1
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.5.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績：FY2026 第一、二季（2026-04～2026-09；會計年度 4 月起）合併營收，以公司公布的各季平均匯率換成美元（第一季 160、第二季 161 日圓）；不用指引或預測。"
                  "層級側＝市場層 markets/nand（讀 Siliconomics）的 Kioxia CY2026 營收（TrendForce 美元口徑，情境 B；CY2026 三情境相同）。期間對應：實績 6 個月對層級 CY2026 × 6 ÷ 12（線性時差近似，Derived）＝主讀數；"
                  "第二讀數＝最近一季（第二季，2026-07～09）× 4 對層級 CY2026（季中點約 8 月中旬，晚於日曆年中點；營收逐季上升時偏高，不得寫成「夾住真值」）。"
                  "口徑差：公司營收含其他營收（零售、對 SanDisk 銷售等，第一季約 3.8%），TrendForce 口徑為 NAND 營收，公司側約高 3%（市場層公司口徑係數 1.030）。兩讀數都超過 ±15% 且同號為「對帳中」，只一個超過為「時差邊界（參考）」；層級隱含值任一部分與個股實績同源時只標「參考（同源校準）」、不判定。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（大部分消去）：市場層 2026Q3 的 Kioxia 份額由 Kioxia 第二季實績（¥2,480.0B ÷ 161 ÷ 公司口徑係數）反推（markets/nand In 頁的第二季實績、平均匯率與公司口徑係數三列），所以層級 2026Q3 營收＝Kioxia 自己的實績，市場值消去（下表核對列為 0）；"
                  "2026Q4 份額＝2026Q3 實績份額與第三季公司指引份額的平均；公司口徑係數由第一季實績 ÷ TrendForce 2Q26 校準；2026Q1–Q2 為 TrendForce 前五大排名（研調，依各廠財報整理，非獨立估計）。"
                  "所以本表實際檢驗的只是 Kioxia 自己的季別分布（與 TrendForce 對 Kioxia 財報的換算），不是對 Siliconomics 層級值的獨立檢驗；依規則標「參考（同源校準）」，結論「無可比實績（層級值同源）」，對帳中維持 0。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "缺的獨立揭露：(1) 不以 Kioxia 財報或指引校準的 CY2026 NAND 市場營收與 Kioxia 份額（TrendForce 3Q26 排名尚未進層級；1H26 排名本身依各廠財報）；(2) Kioxia 企業級 SSD（AI）營收或位元的公司揭露（公司只公布 SSD 與儲存分部，層級沒有 Kioxia 分部份額）。"
                  "分部營收、營業利益：層級不持有，無可比實績；實績列於下表供對照。")
    ROW["none"] = r; r += 1
    hdr(CW, r, ["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"])
    ROW["hdr_in"] = r; CW.cell(r, 9, "鍵名（追蹤用；HTML 不顯示）").font = BOLD; r += 1
    inp = {}

    def put(key, label, unit, per, val, grade, src, fmt):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        inp[key] = r; r += 1
    put("q1", "合併營收：FY2026 第一季", "¥B", "2026-04～2026-06", "=q1_rev", "Verified", "In 頁 q1_rev（連格）；Kioxia Earnings Release FY2026 1Q（2026-07-31；SGX 轉載，網址見 Sources 頁）：營收 ¥1,767.1B（SSD 與儲存 1,174.7、智慧裝置 525.7、其他 66.7）", "#,##0.0")
    put("fx1", "第一季平均匯率", "¥/US$", "2026-04～2026-06", 160, "Verified", "同上（公司揭露平均匯率 160 日圓；In 頁 q1_rev 說明）", "0")
    put("q2", "合併營收：FY2026 第二季（最近一季）", "¥B", "2026-07～2026-09", "=g2_rev", "Verified（二手轉載）", "In 頁 g2_rev（連格）；" + SRC_Q2 + "：¥2,480.0B（指引 2,390.0）", "#,##0.0")
    put("fx2", "第二季平均匯率", "¥/US$", "2026-07～2026-09", 161, "Verified（二手轉載）", "同上（平均匯率 161；markets/nand In 頁 kx_fxg 同值）", "0")
    put("h1", "合併營收（美元換算）：FY2026 第一、二季累計", "US$B", "2026-04～2026-09", f"=D{r - 4}/D{r - 3}+D{r - 2}/D{r - 1}", "Derived", "各季營收 ÷ 該季公司平均匯率後相加", "#,##0.000")
    put("q2u", "合併營收（美元換算）：FY2026 第二季", "US$B", "2026-07～2026-09", f"=D{inp['q2']}/D{inp['fx2']}", "Derived", "第二季營收 ÷ 161", "#,##0.000")
    put("mo", "已公布月數", "月", "2026-04～2026-09", 6, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("lay", "層級值：Kioxia CY2026 營收（市場層，情境 B；TrendForce 美元口徑）", "US$B", "CY2026", f"=SUM(Units!G{urB}:J{urB})", "Derived（層級；部分同源）",
        f"Units 頁 2026Q1–Q4 合計（{NF} Mkt 頁 Kioxia 營收（情境 B）唯讀抄錄；Ann 頁 CY2026 {ann['B'][1][0]:.3f}，三情境相同）", "#,##0.000")
    put("l1", "  其中 2026Q1（Siliconomics nd_rv_kx_4）", "US$B", "2026-01～03", f"=Units!G{urB}", "Derived（TrendForce 2Q26 ÷ 季增）", f"Siliconomics {SIL.VERSION} nd_rv_kx_4（TrendForce 2026-08-18：2Q26 營收 ÷（1＋79.9%））", "#,##0.000")
    put("l2", "  其中 2026Q2（Siliconomics nd_rv_kx_5）", "US$B", "2026-04～06", f"=Units!H{urB}", "Interested-party（研調）", f"Siliconomics {SIL.VERSION} nd_rv_kx_5（TrendForce 2026-08-18 前五大排名；依公司財報整理）", "#,##0.000")
    put("l3", "  其中 2026Q3（市場 × 份額；份額由 Kioxia 第二季實績反推）", "US$B", "2026-07～09", f"=Units!I{urB}", "Derived（同源）", "markets/nand Mkt 頁：份額＝¥2,480.0B ÷ 161 ÷ 公司口徑係數 ÷ 市場", "#,##0.000")
    put("l4", "  其中 2026Q4（市場 × 份額；份額＝第三季實績與第四季指引的平均）", "US$B", "2026-10～12", f"=Units!J{urB}", "Derived（部分同源）", "markets/nand Mkt 頁：份額＝（2026Q3 實績份額＋第三季指引 ¥2,500B ÷ 157 推得份額）÷ 2", "#,##0.000")
    put("rec", "核對：層級 2026Q3 × 公司口徑係數 × 161 − Kioxia 第二季實績（應約 0：市場值消去）", "¥B", "2026-07～09", f"=D{inp['l3']}*kx_k*D{inp['fx2']}-D{inp['q2']}", "Derived", "證明層級 2026Q3 營收＝Kioxia 自己的實績（同源）；In 頁公司口徑係數＝第一季營收 ÷ 160 ÷ TrendForce 2Q26（markets/nand 同式）", "0.000")
    put("kx", "參考：公司營收 ÷ TrendForce 口徑（公司口徑係數）", "x", "2026-04～06", "=kx_k", "Derived（同源校準）", "In 頁公司口徑係數列；公司第一季營收（美元換算）對層級 2026Q2 的比值，本身即校準值", "0.0000")
    put("anc", "參考：模型 FY2026 合併營收（情境 B；個股預測）", "¥B", "FY2026（2026-04～2027-03）", f"=PL!D{X['SUMR'][('B', 'yrev')]}", "Derived（模型）", "PL 頁情境 B FY2026；個股自身預測，非實績；只作參考", "#,##0.0")
    put("g3", "參考：公司指引 FY2026 第三季營收", "¥B", "2026-10～12", "=g3_rev", "Interested-party", "In 頁 g3_rev＝2,500.0（假設 157 日圓；公司指引，不算實績）", "#,##0.0")
    r += 1
    hdr(CW, r, ["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"])
    ROW["hdr_main"] = r; r += 1

    def rd_(label, per, f_a, f_l, note, judge="（讀數）"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "US$B"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.00"; CW.cell(r, 5, f_l).number_format = "#,##0.00"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = rd_("主讀數：已公布期間累計合併營收（美元換算）vs 層級 × 月數 ÷ 12", "2026-04～09 vs CY2026 × 6/12", f"=D{inp['h1']}", f"=D{inp['lay']}*D{inp['mo']}/12",
             "實績＝FY2026 第一＋第二季（各季平均匯率換算）；層級＝市場層 Kioxia CY2026 營收 × 6 ÷ 12（線性時差近似）。實績期間不含 2026-01～03 與 10～12 月，營收逐季大幅上升，時差近似的偏誤方向取決於季別分布，不預設；公司口徑含其他營收約高 3%")
    rb = rd_("第二讀數：最近一季（FY2026 第二季）× 4 vs 層級", "2026-07～09 × 4 vs CY2026", f"=D{inp['q2u']}*4", f"=D{inp['lay']}",
             "最近一季年化；季中點（8 月中旬）晚於日曆年中點，成長期偏高；只作時差近似的另一端參考")
    CW.cell(r, 1, "C0（實績）判定：Kioxia 合併營收"); CW.cell(r, 7, "參考（同源校準）")
    CW.cell(r, 8, "結論：無可比實績（層級值同源）：層級 CY2026 的 2026Q3 即 Kioxia 自己的第二季實績（市場值消去），2026Q4 份額由公司實績與指引推得，2026Q1–Q2 為研調依公司財報整理；缺不以 Kioxia 財報校準的市場與份額獨立揭露。"
                  "數字參考：若不論同源，主讀數在 ±15% 內、第二讀數超過 ±15%，會是「時差邊界（參考）」；對帳中維持 0")
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占模型 FY2026 合併營收（個股自身預測）"); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['q1']}+D{inp['q2']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anc']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（上半年實績 ÷ 模型全年），不是差距；模型值是個股預測，非公司指引也非實績")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, "讀法：公司實績是公司財報（第一季 SGX 轉載原文、第二季為二手轉載）；層級隱含值的 2026Q3 與 2026Q4 份額由 Kioxia 自己的實績與指引推得、2026Q1–Q2 為研調依財報整理，所以這一列檢驗的只是 Kioxia 自己的季別分布，結論：無可比實績（層級值同源），數字僅供參考。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["judge"]},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    r += 1
    CW.cell(r, 1, "其他公司實績（無可比層級值；FY2026 第一季與第二季）").font = BOLD; ROW["oth_t"] = r; r += 1
    hdr(CW, r, ["項目", "單位", "期間", "FY2026 第一季", "FY2026 第二季", "", "判定", "出處／口徑／缺少的揭露"])
    ROW["hdr_oth"] = r; r += 1
    ROW["oth0"] = r
    others = [("SSD 與儲存（分部營收）", "=q1_ss", None, "無可比實績（層級無 Kioxia 分部份額）", "第一季 In 頁 q1_ss（Earnings Release FY2026 1Q）；第二季分部數本輪未取得；層級只有全產業企業級 SSD 位元，沒有 Kioxia 份額"),
              ("智慧裝置（分部營收）", "=q1_sd", None, "無可比實績（層級無對應量）", "第一季 In 頁 q1_sd；第二季未取得；層級不持有"),
              ("營業利益", "=q1_op", "=g2_op", "無可比實績（層級無對應量）", "In 頁 q1_op（Earnings Release）、g2_op（" + SRC_Q2 + "：¥1,994.7B）；層級不持有 Kioxia 損益"),
              ("淨利（歸屬母公司）", "=q1_ni", "=g2_ni", "無可比實績（層級無對應量）", "In 頁 q1_ni、g2_ni（約 ¥1.36 兆，二手轉載）；層級不持有"),
              ("參考：FY2025 第四季合併營收（2026-01～03；不進主讀數）", "=p4_rev", None, "參考（不判定）", "In 頁 p4_rev 1,002.9（stockanalysis 季度表，Interested-party；四季合計＝公司 FY2025 2,337.6）；該季公司平均匯率未取得，故主讀數期間取 2026-04～09")]
    for nm, a1, a2, judge, miss in others:
        CW.cell(r, 1, nm); CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～06／2026-07～09" if "FY2025" not in nm else "2026-01～03")
        for col, v in ((4, a1), (5, a2)):
            if v: CW.cell(r, col, v).number_format = "#,##0.0"
        CW.cell(r, 7, judge); CW.cell(r, 8, miss)
        r += 1
    ROW["oth1"] = r - 1
    return ROW


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, PLA, RVA, SNS, CR, X):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, fbv, cons = vb("同業前瞻 P/B 中位數"), vb("12 個月後的前瞻每股淨值（情境加權）"), vb("市場共識目標價")
    dcfpx = lab(V, "DCF ÷ 現價", exact=True)
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_px, s_sh, s_fx, s_wts = sens_r("NAND 價格"), sens_r("Kioxia 份額"), sens_r("匯率"), sens_r("情境權重")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = X["SUMR"]; FYC = X["FYC"]
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 日圓，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("：DCF "), TX(dcf, "#,##0"), T(" 日圓與同業 P/B 法 "), TX(cadj, "#,##0"), T(" 日圓各占一半，兩者都高於現價。"))
    put("rec", T("投資建議「"), rec, T("」的由來：加權目標價相對現價 "), TX(up, "+0.0%;-0.0%"), T("，依上檔空間五級固定規則判定。三情境 DCF 價值為現價的 "), TX(f"Val!B{dcfpx}", "0.00"), T("／"), TX(f"Val!C{dcfpx}", "0.00"), T("／"), TX(f"Val!D{dcfpx}", "0.00"),
        T(" 倍（樂觀／基準／保守）；DCF 權重降到三成時加權目標價 "), TX(f"Sens!C{SNS['dcfw']}", "#,##0"), T(" 日圓（評等"), REC5(f"(Sens!C{SNS['dcfw']}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 日圓，比現價 "), TX("px", "#,##0"), T(" 日圓"), f'IF({up}<0,"低","高")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec, T("。同業比較法（記憶體同業前瞻 P/B 中位數 "), TX(pmed, "0.00"),
        T(" 倍 × 12 個月後每股淨值 "), TX(fbv, "#,##0"), T(" 日圓）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 日圓。"))
    put("o2", T("最大變數是 NAND 價格、份額與匯率（加權目標價相對基準）：第三季起市場層營收 ×0.8／×1.2 為 "), TX(f"Sens!B{s_px}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_px}", "+0.0%;-0.0%"),
        T("；份額 ×0.9／×1.1 為 "), TX(f"Sens!B{s_sh}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_sh}", "+0.0%;-0.0%"), T("；匯率 145／170 日圓為 "), TX(f"Sens!B{s_fx}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_fx}", "+0.0%;-0.0%"),
        T("；情境權重全給保守 "), TX(f"Sens!B{s_wts}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o3", T("現價隱含什麼：要讓 DCF 等於現價，折現率可升到 "), TX(rv("w1"), "0.0%"), T("（模型 "), TX("wacc", "0.0%"), T("），或 FY2028 起現金流只需模型的 "), TX(rv("s1"), "0.00"),
        T(" 倍；同業法下現價只相當於前瞻每股淨值的 "), TX(rv("pe1"), "0.00"), T(" 倍（同業中位數 "), TX(pmed, "0.00"), T(" 倍）。高峰期盈餘能否維持是關鍵。"))
    put("o4", T("對帳中：CY2027–CY2029 的 AI 加速器部署量，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 日圓（"), f"Coord!C{CR['rec']}", T("），現行基準讀數 "), TX(f"Coord!B{CR['tp']}", "#,##0"), T("（"), f"Coord!B{CR['rec']}", T("），"),
        f'IF(Coord!B{CR["rec"]}<>Coord!C{CR["rec"]},"評等不同","評等相同")', T("。個股不裁決。公司實績（FY2026 上半年營收）對層級隱含值：無可比實績（層級值同源；標參考）。"))
    put("o5", T("進度：FY2026 上半年營收 "), TX("q1_rev+g2_rev", "#,##0"), T(" 十億日圓，占模型 FY2026（情境 B）"), TX(f"PL!D{SB[('B', 'yrev')]}", "#,##0"), T(" 的 "), TX(f"(q1_rev+g2_rev)/PL!D{SB[('B', 'yrev')]}", "0%"),
        T("；第三季公司指引 "), TX("g3_rev", "#,##0"), T("，模型第三季比指引 "), TX("Recon!D7", "+0.0%;-0.0%"), T("（情境 A）。"))
    put("drv", T("分部營收推導（情境 B）：營收＝市場層 Kioxia 營收（TrendForce 美元口徑）× 公司口徑係數 × 匯率；SSD 與儲存、其他＝營收 × 占比，智慧裝置為差額；營業成本＝位元 × 每 EB 成本。FY2028 SSD 與儲存 "),
        TX(f"PL!{FYC[2028]}{SB[('B', 'yss')]}", "#,##0"), T("、智慧裝置 "), TX(f"PL!{FYC[2028]}{SB[('B', 'ysd')]}", "#,##0"), T("、其他 "), TX(f"PL!{FYC[2028]}{SB[('B', 'yot')]}", "#,##0"),
        T("，合計 "), TX(f"PL!{FYC[2028]}{SB[('B', 'yrev')]}", "#,##0"), T(" 十億日圓。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、Advantest、SK 海力士）、三層協調讀取層級與 C0（含公司實績）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史；FY2023 與部分列未找到，留空）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（業務主因整本重算快照）、Coord 頁（兩讀數；讀取層級；C0：FY2026 上半年合併營收實績對市場層 Kioxia CY2026 營收，標參考（同源校準），結論無可比實績（層級值同源），對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, X):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    CO = coord_c0(wb, SIL, CR, X)
    PLA = pl_standard(wb, HR, CR, X)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    RMA = readme13(wb, PLA, RVA, SNS, CR, X)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((("C0 實績：Kioxia Earnings Release FY2026 1Q（2026-07-31；SGX 轉載）：營收 ¥1,767.1B、平均匯率 160", ER1),
                                ("C0 實績：" + SRC_Q2 + "：營收 ¥2,480.0B、平均匯率 161", "In 頁 g2_rev 來源"),
                                ("C0 層級：markets/nand v0.3 Mkt 頁 Kioxia 營收（2026Q1–Q2＝Siliconomics nd_rv_kx_4／5，TrendForce 2026-08-18）", "https://www.trendforce.com/presscenter/news/20260818-13186.html"))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    order = [n for n in ["README", "In", "Q", "Units", "PL", "Val", "DCF", "RevDCF", "Recon", "Sens", "Sources", "Coord", "Hist"] if n in wb.sheetnames]
    assert len(order) == len(wb.sheetnames), (order, wb.sheetnames)
    wb._sheets = [wb[n] for n in order]
    return dict(hist=HR, coord=CO, pl=PLA, revdcf={k: v for k, v in RVA.items() if k in ("sum_dcf", "sum_cmp")}, val=VG, sens=SNS, readme=RMA)
