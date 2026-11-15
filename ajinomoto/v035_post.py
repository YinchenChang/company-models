# 味之素 v0.3.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10 版面；同次產業 Ibiden v0.5.6、日股決算短信 Advantest v0.7.5）＋三層協調
# （Coord 頁兩讀數目標價 tools/coord、讀取層級（同源，不判定）、C0（實績）；規則 docs/plan/三層基準協調規則.md v1.4.2、工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）。
# 只新增，不改既有數值格（tools/xlsx_diff.py 對 v0.3.4 逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司決算短信，IFRS；標準損益表的歷史三年與 FY2023 年增率）
#   In 頁末：N 節（標準損益表推算假設 opx_el；只用於推算毛利，不進 EPS 與目標價）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B）；RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準
#   Coord 頁末：讀取層級、讀取層級（同源，不判定）、層級之間（參考）、C0（實績）；Sources、README 頁末
# build_ajinomoto.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, SIL, CR)；alt_run.py 建置（無 Coord 頁）不呼叫。
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "v0.3.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
TOL = 0.15
JPX24 = "https://www2.jpx.co.jp/disc/28020/140120240508583570.pdf"
JPX25 = "https://www2.jpx.co.jp/disc/28020/140120250507532436.pdf"
JPX26S = "https://www2.jpx.co.jp/disc/28020/081220260501516433_tse-acedifsm-28020-20260507328020-ixbrl.htm"
SAI = "https://stockanalysis.com/quote/tyo/2802/financials/"
RES2Q = "https://www.ajinomoto.co.jp/company/jp/ir/event/presentation/main/0111111110/teaserItems1/00/linkList/01/link/FY26-2Q_Results_J.pdf"
TAN2Q = "https://www.ajinomoto.co.jp/company/jp/ir/event/presentation/main/0111111110/teaserItems1/00/linkList/0/link/FY26Q2_Tanshin_J.pdf"
REV1109 = "https://www.ajinomoto.co.jp/company/jp/ir/event/presentation/main/0111111110/teaserItems1/00/linkList/00/link/FY26Q2_Forecast_J.pdf"
TN1110 = "https://finance.technews.tw/2026/11/10/ajinomoto-raises-earnings-forecast-for-electronic-materials-business/"
FR2Q = "公司 2027 年 3 月期 第 2 四半期 決算概要（2026-11-09；p.2 セグメント別、ファンクショナルマテリアルズ）"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
TX = lambda ref, f: f'TEXT({ref},"{f}")'
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


def lab(ws, prefix, start=1, col=1, exact=False, end=None):
    for r in range(start, (end or ws.max_row) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- Hist：損益表歷史（IFRS；十億日圓） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025, 標記, 來源)；FY(t)＝t 年 4 月～t＋1 年 3 月（FY2025＝2026 年 3 月期）
SRC_H = (f"公司決算短信〔IFRS〕連結損益計算書（百萬日圓 ÷ 1000）：2024 年 3 月期（2024-05-09，p.10；FY2023 與比較欄 FY2022；{JPX24}）、"
         f"2025 年 3 月期（2025-05-07，p.10；FY2024；{JPX25}）")
HIS = [
    ("rev", "營收（売上高）", 1359.115, 1439.231, 1530.556, "=a_rev", "Verified", SRC_H + "；FY2025 引用 In 頁 a_rev（2026 年 3 月期決算短信）"),
    ("cogs", "營業成本（売上原価）", 888.727, 927.783, 979.792, "=F{rev}-F{gp}", "Verified／Derived", SRC_H + "；FY2025＝營收 − 毛利（Derived）"),
    ("gp", "毛利（売上総利益）", 470.387, 511.448, 550.764, 597.149, "Verified（FY2025 資料商）",
     SRC_H + f"；FY2025 597,149 百萬日圓取 stockanalysis 損益表（{SAI}；資料商轉載，未對決算短信原文；同站 FY2024 550,764 與決算短信相符）"),
    ("bp", "事業利益（IFRS；公司主要指標）", 135.341, 147.681, 159.302, "=a_bp", "Verified", SRC_H + "；FY2025 引用 In 頁 a_bp"),
    ("oi", "營業利益（IFRS 営業利益；含其他營業收支）", 148.928, 146.682, 113.968, "=a_oi", "Verified", SRC_H + "；FY2025 引用 In 頁 a_oi"),
    ("pti", "稅前利益（税引前当期利益）", 140.033, 142.043, 108.330, 196.115, "Verified", SRC_H + f"；FY2025 196,115 取 2026 年 3 月期決算短信摘要頁（2026-05-07；{JPX26S}）"),
    ("prof", "當期利益（含非支配持分）", 100.170, 102.032, 80.773, 145.060, "Verified", "同上（FY2025 當期利益 145,060）"),
    ("ni", "稅後淨利（歸屬母公司）", 94.065, 87.121, 70.272, "=a_ni", "Verified", SRC_H + "；FY2025 引用 In 頁 a_ni"),
    ("eps_r", "EPS 基本（公司公告值，公告當時股數口徑）", 175.97, 167.44, 69.77, "=a_eps", "Verified",
     SRC_H + "；FY2022、FY2023 為 2025-04-01 一拆二之前的公告值，FY2024 起為分割後口徑（2025 年 3 月期決算短信已追溯）"),
    ("split", "換算到 2025-04 分割後的倍數（公告值 ÷ 本列）", 2, 2, 1, 1, "Derived", "2025-04-01 一拆二（2026 年 3 月期決算短信註記）"),
]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = (f"Hist — 損益表歷史（{VER} 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。FY(t)＝t 年 4 月～t＋1 年 3 月：FY2022＝2023 年 3 月期 … FY2025＝2026 年 3 月期。"
               "金額：十億日圓；EPS：日圓。IFRS：事業利益＝營收 − 營業成本 − 販管費 − 研發費＋持分法損益（公司主要指標）；営業利益另含其他營業收支。只被 PL 頁標準損益表引用，不回饋任何計算")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    R = {k: 4 + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v22, v23, v24, v25, tag, src in HIS:
        r = R[k]
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for col, v in zip("CDEF", (v22, v23, v24, v25)):
            if isinstance(v, str):
                v = v.format(**R)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k.startswith("eps") else ("0" if k == "split" else "#,##0.000")
            if not isinstance(v, str):
                c.font = BLUE
        H.cell(r, 7, tag); H.cell(r, 8, src)
    r = max(R.values()) + 1
    for k, labl, f_, fmt, note in (("opx", "營業費用（＝毛利 − 事業利益：販管費＋研發費 − 持分法損益）", "={c}{gp}-{c}{bp}", "#,##0.000", "由財報毛利與事業利益推得"),
                                   ("ooi", "其他營業收支淨額（＝営業利益 − 事業利益）", "={c}{oi}-{c}{bp}", "#,##0.000", "由財報推得（FY2025 含本社大樓土地出售利益）"),
                                   ("fin", "金融收支淨額（＝稅前 − 営業利益）", "={c}{pti}-{c}{oi}", "#,##0.000", "由財報推得"),
                                   ("tax", "所得稅（＝稅前 − 當期利益）", "={c}{pti}-{c}{prof}", "#,##0.000", "由財報推得（FY2022–FY2024 與決算短信法人所得稅相同）"),
                                   ("nci", "非支配股東損益（＝當期利益 − 歸屬母公司）", "={c}{prof}-{c}{ni}", "#,##0.000", "由財報推得"),
                                   ("eps", "EPS 基本（分割後換算；模型口徑）", "={c}{eps_r}/{c}{split}", "0.00", "公告值 ÷ 分割倍數；FY2025 與 In 頁 a_eps 相同"),
                                   ("sh", "加權平均股數（十億股，分割後；＝淨利 ÷ EPS）", "={c}{ni}/{c}{eps}", "0.0000", "Derived")):
        R[k] = r; H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = fmt
        H.cell(r, 7, "Derived"); H.cell(r, 8, note); r += 1
    r += 1
    H.cell(r, 1, "核對（十億日圓；百萬日圓四捨五入，可有 ±0.002 的捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("事業利益 − 營收 ＋ 營業成本 ＋ 營業費用（恆等式）", "={c}{bp}-{c}{rev}+{c}{cogs}+{c}{opx}"),
                     ("稅前 − 所得稅 − 非支配 − 稅後淨利（恆等式）", "={c}{pti}-{c}{tax}-{c}{nci}-{c}{ni}")):
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "核對：FY2023、FY2024 所得稅（推得）− 決算短信法人所得稅（應為 0）")
    H[f"D{r}"] = f"=D{R['tax']}-40.011"; H[f"E{r}"] = f"=E{R['tax']}-27.556"
    for c in "DE": H[f"{c}{r}"].number_format = "0.000"
    R["chk1"] = r
    H.column_dimensions["A"].width = 10; H.column_dimensions["B"].width = 58
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["H"].width = 110
    return R


# ---------- In：N 節（標準損益表推算假設） ----------
def in_n(wb):
    IN = wb["In"]
    r = IN.max_row + 2
    IN.cell(r, 1, f"N. 標準損益表推算假設（{VER} 新增；只用於 PL 頁「標準損益表」推算毛利與營業費用，不進 EPS、現金流與目標價）").font = BOLD; r += 1
    vals = ("opx_el", "營業費用彈性（營業費用年增 ÷ 營收年增；FY2026–FY2029 推算毛利用）", "x", 0.7, 1.0, 1.3, "Assumed",
            "模型只推事業利益：毛利＝事業利益＋營業費用（推算）；營業費用＝前一年 ×（1＋彈性 × 營收年增）。FY2024、FY2025 實績營業費用年增 +7.6%、+6.3% 對營收年增 +6.3%、+3.5%（Hist 頁），彈性 1.2、1.8，無穩定比例；基準取 1（費用率不變）",
            "Assumed；低／高不進任何目標價（只供讀者觀察毛利率對費用假設的敏感度）")
    for j, v in enumerate(vals, 1):
        c = IN.cell(r, j, v)
        if j in (4, 5, 6): c.font = BLUE
    dn(wb, "opx_el", f"In!$E${r}")
    return r


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
COLS = "CDEFGHIJK"
HC = {"C": "D", "D": "E", "E": "F"}   # 標準損益表歷史欄 → Hist 欄（FY2023→D、FY2024→E、FY2025→F；FY2022 在 C）
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029}


def pl_standard(wb, HR, CR):
    P = wb["PL"]
    rb = lab(P, "情境 B（基準）：")
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 60)
    R = dict(rev=g("營收"), bp=g("事業利益"), ni=g("淨利（歸屬母公司）"), sh=g("股數（期中平均）"), eps=g("EPS（日圓）"))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來六年（FY2026–FY2031，情境 B 基準；FY2026E 為公司 2026-08-06 上修後預測的錨點）。¥B；EPS 日圓。"
                 "只引用本頁上方、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（IFRS，公司決算短信的連結損益）。公司以「事業利益」（營收 − 營業成本 − 販管費 − 研發費＋持分法損益）為主要獲利指標，本表的營業利益列即事業利益（模型與 DCF 的 EBIT 同一口徑）；"
                 "IFRS 営業利益另含其他營業收支（FY2025 含本社大樓土地出售利益），列在其他收入。公司與市場共識都以此為主，沒有另一套 Non-GAAP。預測期的其他營業收支與金融收支不估（其他收入＝稅前利益 − 營業利益，由淨利反推）。")
    r_note = r; r += 1
    P.cell(r, 1, "預測期只推事業利益與淨利：FY2026–FY2029 的營業費用＝前一年 ×（1＋彈性 × 營收年增）（彈性見 In 頁 N 節），毛利＝營業利益＋營業費用，營業成本＝營收 − 毛利；"
                 "稅前利益＝淨利 ÷（1 − 稅率），所得稅＝稅前利益 × 稅率，非支配股東損益併入其他收入。FY2030E–FY2031E（收斂法延伸期，只推營收與淨利）：毛利率與營業費用率沿用前一年，"
                 "營業利益取 DCF 頁 EBIT，與毛利 − 營業費用的差額列核對列，不改淨利與目標價。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJK":
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
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1
    M = "#,##0.0"; PCT = "0.0%"
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    fc4 = lambda k: [f"=PL!{c}{R[k]}" for c in "DEFG"]
    nxt = lambda k: [f"=PL!{c}{R[k]}" for c in "HI"]   # 本模型延伸期在情境 B 區塊 H、I 欄（FY2030、FY2031）
    row("rev", "營收", "¥B", hh("rev") + fc4("rev") + nxt("rev"), M, "FY2026E＝公司預測；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E–FY2031E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 取 Hist 頁")
    rc = r
    row("cogs", "營業成本（COGS）", "¥B", hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHI"] + [f"={c}{L['rev']}*(1-{p}{rc + 1}/{p}{L['rev']})" for p, c in (("I", "J"), ("J", "K"))], M,
        "FY2026E–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − 前一年毛利率）")
    ro = rc + 4
    row("gp", "毛利", "¥B", hh("gp") + [f"={c}{ro}+{c}{rc + 3}" for c in "FGHI"] + [f"={c}{L['rev']}-{c}{rc}" for c in "JK"], M, "FY2026E–FY2029E＝營業利益（事業利益）＋營業費用（推算）")
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "預測期為推算值（營業費用彈性假設）；FY2030E–FY2031E＝前一年")
    row("opx", "營業費用（販管費＋研發 − 持分法損益）", "¥B", hh("opx") + [f"={p}{rc + 3}*(1+opx_el*({c}{L['rev']}/{p}{L['rev']}-1))" for p, c in zip("EFGH", "FGHI")]
        + [f"={c}{L['rev']}*{p}{rc + 3}/{p}{L['rev']}" for p, c in (("I", "J"), ("J", "K"))], M, "歷史＝毛利 − 事業利益；FY2026E–FY2029E＝前一年 ×（1＋彈性 × 營收年增）；FY2030E–FY2031E＝營收 × 前一年營業費用率")
    assert L["opx"] == rc + 3
    row("oi", "營業利益（IFRS 事業利益）", "¥B", hh("bp") + fc4("bp") + [f"=DCF!{c}{d_ebit}" for c in "GH"], M, "公司主要指標；FY2030E–FY2031E 取 DCF 頁情境 B 的 EBIT（延伸期由淨利反推；不改）")
    assert L["oi"] == ro
    row("opm", "營業利益率（事業利益率）", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額：其他營業收支＋金融收支）", "¥B", [f"={c}{r + 1}-{c}{L['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益：實績含其他營業收支（FY2025 土地出售利益）與金融收支；預測期由淨利反推（非支配併入）")
    row("pti", "稅前利益", "¥B", hh("pti") + [f"={c}{r + 2}/(1-tax_w)" for c in "FGHIJK"], M, "預測期＝稅後淨利 ÷（1 − 稅率）")
    row("tax", "所得稅", "¥B", hh("tax") + [f"={c}{L['pti']}*tax_w" for c in "FGHIJK"], M, "預測期＝稅前利益 × 稅率（In 頁 WACC 用稅率 30.6%）")
    row("ni", "稅後淨利（歸屬母公司）", "¥B", hh("ni") + fc4("ni") + nxt("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("nci", "（備註）非支配股東損益", "¥B", hh("nci") + [None] * 6, M, "預測期併入其他收入，不另列")
    row("ifoi", "（備註）IFRS 営業利益（含其他營業收支；只列實績）", "¥B", hh("oi") + [None] * 6, M, "FY2026E 公司預測 184.2（In 頁 f_oi；Q1 上修後）")
    row("sh", "股數（十億股，分割後）", "十億股", [f"=Hist!{HC[c]}{HR['sh']}" for c in "CDE"] + fc4("sh") + nxt("sh"), "0.0000")
    row("eps", "EPS（基本；分割後）", "¥", [f"=Hist!{HC[c]}{HR['eps']}" for c in "CDE"] + fc4("eps") + nxt("eps"), "0.00", "模型口徑為基本 EPS（淨利 ÷ 期中平均股數）；公司無潛在股式，基本＝稀釋")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率（事業利益率）", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/Hist!C{HR['ni']}-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "¥B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2031E 為差額，見下）", "¥B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None, None], "0.000")
    row("c2b", "差額（FY2030E–FY2031E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利反推）；目標價用延伸期淨利，不受影響", "¥B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "差額（FY2030E–FY2031E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JK"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利（預測期；歷史欄差額＝非支配股東損益）", "¥B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − PL 情境 B 區塊（FY2026E–FY2031E）", "¥B", [None] * 3 + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJK", "DEFGHI")], "0.000")
    row("c5", "核對：營收 − PL 情境 B 區塊（FY2026E–FY2031E）", "¥B", [None] * 3 + [f"={c}{L['rev']}-PL!{s}{R['rev']}" for c, s in zip("FGHIJK", "DEFGHI")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型 FY(t)＝t 年 4 月至 (t+1) 年 3 月，引擎對 FY(t) 直接取 CY(t) 的加速器顆數與 Siliconomics 年別比值（重疊 9 個月，不平移；與 Ibiden 模型相同）；"
                 "標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2031E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。")
    L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF（同 Ibiden v0.5.5；DCF 頁 C＝FY2026（評價日後一季）… I＝FY2032） ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數"), (D.cell(fr, 1).value, D.cell(ty, 1).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"(dcf_nc+DCF!$C${fcf}*stub_frac-stub_div)"
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（{VER} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2026（錨點）與 FY2027（三情境相同）不動，營業利益率與現金流倍數只動 FY2028–FY2032。反解值超出合理範圍時照實列出，不外推")
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
    kv("px", "現價（日圓）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（日圓）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；¥B）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（同 DCF 頁每股價值公式）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（¥B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；¥B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（日圓）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2026, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGHI"):
            c = RV.cell(r, 2 + i, fn(dc, "BCDEFGH"[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（¥B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（事業利益；¥B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "I" else "=0"), "0.0000")
    yr("k", "每 1 日圓營業利益對 EV 的貢獻（FY2026–FY2027 不動，記 0）", lambda dc, c: ("=0" if dc in "CD" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（¥B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（¥B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（¥B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（¥B）", f"=SUMPRODUCT(D{Y['rev']}:H{Y['rev']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（¥B）", f"=B{A['ev']}-SUMPRODUCT(D{Y['ebit']}:H{Y['ebit']},D{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(D{Y["m"]}:H{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2026、FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2026＋FY2027 現值（¥B）", f"=B{Y['pv']}+C{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（¥B）", f"=SUM(D{Y['pv']}:H{Y['pv']})+DCF!I{nop}*H{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（¥B）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；¥B）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近下界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（日圓）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$I${fcf},DCF!$C${fr}:$I${fr},1/(1+{w})^DCF!$C${ty}:$I${ty})"
                f"+DCF!$I${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C$8+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.001"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "#,##0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.1 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（日圓，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；日圓）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 食品同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(" 日圓；要支持現價 "), TX("px", "#,##0"), T(" 日圓，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億日圓，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或事業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(" 日圓相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("食品同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 味之素過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給味之素的倍數是食品同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍（含 ABF 的溢價）。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFGH": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    return A


# ---------- Val：估值並列四個目標價（G 節） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    pxr = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, f"{VER} 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "日圓", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("食品同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(" 日圓；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("食品同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{pxr}"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（日期見左欄）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（H 節；引用 Sens 頁「目標價參數敏感度」表與 Val、Coord 頁） ----------
TOR_LABEL = {"WACC ±1pt": "折現率（WACC）低／高一個百分點",
             "β 低／高": "股權風險係數低／高（各自重新加槓桿）",
             "ERP 低／高": "股權風險溢酬低／高",
             "永續成長率 g ±1pt": "永續成長率低／高一個百分點",
             "RONIC＝WACC＋5pt": "永續期新投資報酬率提高到折現率加五個百分點（只有高）",
             "資本支出 ÷ 營收 ±1pt": "資本支出占營收低／高一個百分點",
             "營運資金 ÷ 營收 低／高": "營運資金占營收增量低／高",
             "淨現金不含權益投資": "淨現金不含權益證券等非營業投資（只有低）",
             "終值改用 FCFF ×（1＋g）": "終值改用最後一年自由現金流成長法（只有高）",
             "美元口徑折現": "美元口徑折現（無風險利率改美國十年期；只有低）",
             "DCF 權重 0.3／0.7": "加權目標價中 DCF 的權重 30%／70%（基準 50%）"}


def sens_tor(wb, CR):
    SE = wb["Sens"]; V = wb["Val"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    assert SE.cell(hr, 1).value == "參數"
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；{VER} 起，版面規格 v1.3 H 節；引用上方「目標價參數」表與 Val、Coord 頁，白話標籤，不回饋計算）").font = BOLD; t0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    for sr in src:
        t = str(SE.cell(sr, 1).value)
        key = next((k for k in TOR_LABEL if t.startswith(k)), None)
        SE.cell(r, 1, TOR_LABEL[key] if key else f"=A{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"={scol}{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{sr}"); r += 1
    tp = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    cu = f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    SE.cell(r, 1, "同業比較法不加歷史倍數調整（相對倍數係數＝1；Comparable 取市場主流同業法）")
    SE.cell(r, 2, f"=(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/{tp}-1").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Val 頁：DCF 與未調整 Comparable 依 DCF 權重加權"); r += 1
    SE.cell(r, 1, "產業部署 GW 改用另一側讀數（Wattonomics 換算；Siliconomics 標對帳中，整本重建）")
    SE.cell(r, 2, f"=Coord!D{CR['tp']}").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Coord 頁兩讀數目標價（另一側相對現行）"); r += 1
    return dict(t0=t0, r1=first, r2=r - 1)


# ---------- Coord：三層協調 C0（讀哪一層、讀取層級（同源）、層級之間、C0（實績）） ----------
def coord_c0(wb, SIL, CR):
    CW = wb["Coord"]; P = wb["PL"]
    N = SIL.NAMES; Vs = SIL.V
    rb = lab(P, "情境 B（基準）：")
    pb = lambda label: lab(P, label, rb, exact=True, end=rb + 60)
    r = CW.max_row + 3
    G = {}
    CW.cell(r, 1, f"三層協調 C0（{VER}；三層基準協調規則 v1.4.2 C0；工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）：讀哪一層、哪一版；讀取層級（同源，不判定）；C0（實績）").font = BOLD; G["c0"] = r; r += 1
    CW.cell(r, 1, "原則：同一數字只有一個權威層，下游註明版本、不自行重估（Project 規則第 3 條）；C0 只用公司實績對層級隱含值判定，兩讀數都超過 ±15% 且同號才標「對帳中」，不改上游值，由主控開對帳單；個股不裁決（C4），往上回饋只有證據（C5）。")
    r += 2
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    CW.cell(r, 1, "1. 本模型讀哪一層、哪一版（硬體供應鏈個股讀 Siliconomics；Wattonomics 只經 Siliconomics 的 R4 比值與另一側讀數進入，現行版本 v0.3.2）").font = BOLD; G["s1"] = r; r += 1
    for j, h in enumerate(["項目", "值", "層級名稱（版本）", "本模型用途"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("Siliconomics 快照版本", SIL.VERSION, ver, "行業變數單一來源（siliconomics/reader.py）；下游不讀 In、Path、Mkt 等內部頁", None),
            ("加速器顆數 CY2026／27／28／29（NVIDIA＋非 NVIDIA，情境 A 原始顆數；百萬顆）", '=TEXT(unv_26+uot_26,"0.0")&"／"&TEXT(unv_27+uot_27,"0.0")&"／"&TEXT(unv_28+uot_28,"0.0")&"／"&TEXT(unv_29+uot_29,"0.0")',
             "tsmc 模型成品 In 頁（與 Siliconomics SIF_UnitsPlat_* 同值，本頁第 2 節逐年核對）", "PL 引擎 A 節：有效顆數＝NVIDIA＋面積係數 × 非 NVIDIA", None),
            ("顆數情境係數 B：CY2028／CY2029", '=TEXT(um28_B,"0.00")&"／"&TEXT(um29_B,"0.00")', f"{ver} SIF_UnitsCoef_Base（經 tsmc 模型 In 頁讀入）", "情境 B（基準）；A＝1.0、C 個股自訂", None),
            ("ABF 可出貨量上限 CY2027／28／29（FY2026＝1）", '=TEXT(cap_27,"0.00")&"／"&TEXT(cap_28,"0.00")&"／"&TEXT(cap_29,"0.00")', f"{ver} abffilm_27–29（SIL.row9 讀入 In 頁 cap_；Siliconomics M15 註記「原在各公司模型」，即本模型校準值寫入）", "出貨量指數＝MIN(需求, 上限)", None),
            ("每顆載板面積 × 層數年增", "=g_asp", f"Ibiden 模型成品 In 頁（＝{ver} abf_gasp）", "AI 用 ABF 需求量指數", "0.0%"),
            ("非 NVIDIA 加速器每顆面積係數", "=v_ot", f"欣興模型成品 In 頁（{ver} abf_vot 同名量 0.80）", "有效顆數", "0.00"),
            ("Wattonomics 錨定係數 CY2027／28／29（SIF_WTAnchor_Base；R4 比值）", '=TEXT(wta_27,"0.000")&"／"&TEXT(wta_28,"0.000")&"／"&TEXT(wta_29,"0.000")',
             f"{ver}；Wattonomics 現行 v0.3.2（只經 Siliconomics 進入）", "只進敏感度（Sens「Wattonomics 錨定」案例）；不改基準", None),
            ("另一側讀數 ÷ 現行基準 CY2026／27／28／29（SIF_DeployGW_AltRatio_<年>）", f'=TEXT(B{CR["ratio"]},"0.000")&"／"&TEXT(C{CR["ratio"]},"0.000")&"／"&TEXT(D{CR["ratio"]},"0.000")&"／"&TEXT(E{CR["ratio"]},"0.000")',
             f"{ver}（比值列見本頁上方）", "兩讀數目標價（本頁上方）只換產業部署 GW 一個輸入（情境 B、C 的加速器顆數 × 比值）", None)]
    for label, val, who, use, fmt in rows:
        CW.cell(r, 1, label); c = CW.cell(r, 2, val)
        if fmt: c.number_format = fmt
        CW.cell(r, 3, who); CW.cell(r, 4, use); r += 1
    G["s1_end"] = r - 1
    r += 1
    CW.cell(r, 1, "2. 讀取層級（同源，不判定）：本模型自己的產業量與 Siliconomics 層級值逐年對照。個股的量直接讀自 Siliconomics（或由本公司模型校準後寫入 Siliconomics），差距恆為 0 或只反映本模型校準，無檢驗力，"
                  "不計入 C0 對帳中項目；C0 判定見下方第 4 節「C0（實績）」").font = BOLD; G["s2"] = r; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位", "口徑、來源、證據等級"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G["s2h"] = r - 1

    def block(key, title, unit, model, layer, note_m, note_l, fmt, verdict=SAME):
        nonlocal r
        CW.cell(r, 1, f"{title}：本模型"); CW.cell(r, 6, unit); CW.cell(r, 7, note_m)
        for j, f in enumerate(model):
            if f is not None: CW.cell(r, 2 + j, f).number_format = fmt
        rm = r; r += 1
        CW.cell(r, 1, f"{title}：層級值"); CW.cell(r, 6, unit); CW.cell(r, 7, note_l)
        for j, v in enumerate(layer):
            if v is None: continue
            c = CW.cell(r, 2 + j, v); c.number_format = fmt
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        rl = r; r += 1
        CW.cell(r, 1, f"{title}：差距（本模型 ÷ 層級值 − 1）")
        for j, col in enumerate("BCDE"):
            if model[j] is None or layer[j] is None: continue
            CW.cell(r, 2 + j, f'=IF({col}{rl}=0,"—",{col}{rm}/{col}{rl}-1)').number_format = "+0.0%;-0.0%"
        rg = r; r += 1
        CW.cell(r, 1, f"{title}：判定（同源，不判定）" if verdict == SAME else f"{title}：判定")
        for j in range(4):
            if model[j] is not None and layer[j] is not None: CW.cell(r, 2 + j, verdict)
        rj = r; r += 1
        G[key] = (rm, rl, rg, rj)
    ucols = "DEFG"
    unv, uot = pb("NVIDIA 加速器顆數 × 情境係數"), pb("非 NVIDIA 加速器顆數 × 情境係數")
    block("units", "加速器顆數合計（情境 B；NVIDIA＋非 NVIDIA）", "百萬顆", [f"=PL!{c}{unv}+PL!{c}{uot}" for c in ucols], list(N["SIF_UnitsPlat_Total_Base"][1:5]),
          "PL 情境 B 區塊（tsmc 成品顆數 × 情境 B 係數；FY(t) 對照 CY(t)）", f"Derived：Siliconomics {SIL.VERSION} SIF_UnitsPlat_Total_Base（原始顆數 × 基準係數）", "#,##0.000")
    block("cap", "ABF 可出貨量上限指數（FY2026＝1）", "x", [None, "=cap_27", "=cap_28", "=cap_29"], [None] + [Vs[f"abffilm_{y}"]["base"] for y in (27, 28, 29)],
          "In 頁 cap_27–29（SIL.row9 讀入）", f"Assumed：Siliconomics {SIL.VERSION} abffilm_27–29（M15 註記：原在味之素模型；同源校準）", "0.000")
    G["same_keys"] = ["units", "cap"]
    # 單值輸入的讀取核對（同源或外部估計，不判定）
    CW.cell(r, 1, "單值輸入：本模型 vs 層級值（同源或同一外部估計，不判定）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "本模型", "層級值", "差距", "", "單位", "口徑、來源、證據等級"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    G["sv_h"] = r; r += 1
    G["sv0"] = r
    for nm, ref, silk, unit, fmt, note in (("每顆載板面積 × 層數年增", "=g_asp", "abf_gasp", "%", "0.0%", "本模型讀 Ibiden 模型；Siliconomics 同名量同值（同源，不判定）"),
                                          ("非 NVIDIA 加速器每顆面積（金額）係數", "=v_ot", "abf_vot", "x", "0.00", "本模型讀欣興模型；Siliconomics 同名量（同一假設，不判定）"),
                                          ("非 AI ABF 需求年增", "=g_nai", "abf_gnai", "%", "0.0%", "Siliconomics 該列註記「欣興、南電、味之素模型同值」（同源，不判定）"),
                                          ("AI 用占比 FY2026（本模型：AI 用 ÷ FM 營收；層級：AI 用 ÷ ABF 需求金額）", "=ai_sh", "abf_ai26", "%", "0.0%",
                                           "兩者都引 Palliser Capital（股東行動派，Interested-party）：本模型取 40%（Assumed，FY2025E 29% 與 FY2026E 45% 之間），Siliconomics 取 FY2026E 45%；外部估計，不判定")):
        CW.cell(r, 1, nm); CW.cell(r, 2, ref).number_format = fmt
        c = CW.cell(r, 3, Vs[silk]["base"]); c.number_format = fmt; c.font = BLUE
        CW.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 6, unit); CW.cell(r, 7, f"{note}；Siliconomics {SIL.VERSION}"); CW.cell(r, 8, f"鍵名 {silk}（追蹤用；HTML 不顯示）"); r += 1
    G["sv1"] = r - 1
    r += 1
    CW.cell(r, 1, "3. 讀取層級與層級之間的結果（HTML 用；公式帶數字）").font = BOLD; G["s3"] = r; r += 1
    mx = lambda k: "MAX(" + ",".join(f"IFERROR(ABS({c}{G[k][2]}),0)" for c in "BCDE") + ")"
    CW.cell(r, 1, F(T(f"本模型讀 Siliconomics {SIL.VERSION}：加速器顆數（情境 B）、ABF 可出貨量上限與層級值的最大差距分別為 "), TX(mx("units"), "0.0%"), T("、"), TX(mx("cap"), "0.0%"),
                    T("（同源：本模型直接讀入，或由本模型校準後寫入 Siliconomics，差距只反映讀取一致，不判定）；AI 用占比本模型 "), TX(f"B{G['sv1']}", "0%"), T("、層級 "), TX(f"C{G['sv1']}", "0%"),
                    T("，兩者都是股東行動派的估計（外部估計，只列參考）。"))); G["sum1"] = r; r += 1
    CW.cell(r, 1, F(T("層級之間（參考）：本模型不讀部署 GW（以加速器顆數推導），兩讀數只換產業量；Siliconomics 部署 GW 的 Wattonomics 換算值 ÷ 現行基準 CY2027–CY2029 為 "),
                    TX(f"C{CR['ratio']}", "0.00"), T("／"), TX(f"D{CR['ratio']}", "0.00"), T("／"), TX(f"E{CR['ratio']}", "0.00"),
                    T("；這是 Siliconomics 與 Wattonomics 兩層之間的差距（Siliconomics 對帳單 r3 處理；已並列兩讀數目標價），不計入本公司 C0，本模型不改上游值、不裁決。另一側讀數只縮放 FY2027 起的顆數，"
                      "FY2025 實績與 FY2026 校準年（指數分母）不縮放（tools/coord/README.md 兩讀數規格）。"))); G["sum2"] = r; r += 1
    CW.cell(r, 1, "C4／C5：本模型不裁決兩層的對帳中年份；本版沒有可往上回饋的新證據（往上回饋只有實績、法定文件揭露的承諾、已簽約或已確認站點，個股預測與公司指引不往上流）。FM 營收實績只屬本公司，Siliconomics 沒有 ABF 膜市場可接，不往上流。")
    G["sum3"] = r; r += 1
    G["end"] = r - 1
    return G


def c0_actuals(wb, G, SIL):
    CW = wb["Coord"]; Vs = SIL.V
    r = G["end"] + 3
    R = {}

    def put(key, label, b=None, unit="", note="", fmt="#,##0.0", blue=False, c=None, url=""):
        nonlocal r
        CW.cell(r, 1, label)
        for col, v in ((2, b), (3, c)):
            if v is None: continue
            x = CW.cell(r, col, v); x.number_format = fmt
            if blue and not (isinstance(v, str) and v.startswith("=")): x.font = BLUE
        if unit: CW.cell(r, 6, unit)
        if note: CW.cell(r, 7, note)
        if url: CW.cell(r, 8, url)
        R[key] = r; r += 1
        return r - 1

    def bold(key, text):
        nonlocal r
        CW.cell(r, 1, text).font = BOLD; R[key] = r; r += 1

    def hdr(cols):
        nonlocal r
        for j, h in enumerate(cols, 1):
            c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
        r += 1
        return r - 1
    bold("title", "4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    put("method", "方法：個股側＝味之素已公布實績（2027 年 3 月期 第 2 四半期 決算概要 p.2：ファンクショナルマテリアルズ（FM，ABF 所屬子部門）營收；公司不單獨揭露 ABF 營收）；層級側＝Siliconomics 現行版同口徑的隱含值。"
               "會計年度 FY2026＝2026-04～2027-03，主讀數＝已公布期間累計（2026-04～09，6 個月）vs 層級 CY2026 × 6 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（2026-07～09）× 4 vs 層級 CY2026"
               "（約等於該季中點 2026-08 中的年化速率；FM 營收逐季加速時，季中點晚於年中使此讀數偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致；層級沒有同口徑值 → 無可比實績。")
    put("limit", "限制：Siliconomics 有加速器顆數、ABF 載板產值與年增、各載板廠產能指數與本模型校準寫入的 ABF 可出貨量上限（同源），沒有 ABF 增層膜市場值，也沒有味之素份額；"
               "公司不揭露 ABF 單獨營收與 AI 占比。結論：無可比實績；缺層級的 ABF 膜市場（或每片載板的膜用量 × 單價）與味之素份額、公司的 ABF 單獨營收揭露。顆數與產能上限的讀取一致見本頁第 2 節（同源，不判定）。")
    r += 1
    bold("a", "4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    R["a_h"] = hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    put("h1", "FM 營收：FY2026 上半年累計（2026-04～09）", 73.3, "¥B", f"Verified：{FR2Q}：733 億日圓（年增 63%；前年同期 449）", blue=True, url=RES2Q)
    put("q2", "FM 營收：FY2026 第二季（2026-07～09；最近一季，第二讀數的實績基礎）", 40.1, "¥B", f"Verified：{FR2Q}：7–9 月 401 億日圓（前年同期 234）", blue=True, url=RES2Q)
    put("q1", "FM 營收：FY2026 第一季（2026-04～06）", "=q1_fm", "¥B", "Verified：In 頁 q1_fm（2027 年 3 月期 Q1 決算概要 33.1，2026-08-06）")
    put("chk", "檢查：上半年 − 第一季 − 第二季（應約 0；億日圓四捨五入）", f"=B{R['h1']}-B{R['q1']}-B{R['q2']}", "¥B", "Derived；差 0.1 為兩份決算概要各自四捨五入", fmt="0.0")
    put("h1bp", "FM 事業利益：上半年累計（參考）", 42.9, "¥B", f"Verified：{FR2Q}：429 億日圓（年增 82%）；第二季 238", blue=True, url=RES2Q)
    put("cons", "全公司營收：上半年累計（參考）", 831.1, "¥B", f"Verified：{FR2Q} p.1：8,311 億日圓（事業利益 1,086、營業利益 1,024、歸屬母公司淨利 733）；另見同日決算短信", blue=True, url=RES2Q)
    put("abf", "公司揭露：ABF 單獨營收與 AI 占比", "未揭露", "", "公司只揭露 FM（電子材料等）子部門；AI 占比為 Palliser Capital（股東行動派）估計（Interested-party），不算實績", fmt="General")
    put("mo", "已公布月數（主讀數期間 2026-04～09）", 6, "月", "輸入：上半年實績", fmt="0", blue=True)
    r += 1
    bold("b", "4b. 層級隱含值（Siliconomics 現行版；同口徑值不存在者照實列出）")
    R["b_h"] = hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    put("lfilm", "層級：ABF 增層膜市場與味之素份額", "無", "", f"Siliconomics {SIL.VERSION} 無此量；不以本模型 FM 營收或可出貨量上限（同源）補位", fmt="General")
    put("lsub", "層級（參考，口徑不同）：全球 ABF 載板產值", f"={Vs['abfm_25']['base']}*(1+{Vs['abfm_g26']['base']})", "US$B",
        f"Derived：Siliconomics {SIL.VERSION} 全球 ABF 載板產值 2025 ×（1＋2026 年增）（TPCA／IEK，Interested-party；載板產值，不是增層膜）", fmt="#,##0.00")
    put("lcap", "層級：ABF 可出貨量上限 CY2027（FY2026＝1）", f"={Vs['abffilm_27']['base']}", "x", f"Siliconomics {SIL.VERSION} ABF 可出貨量上限 2027（本模型校準寫入，同源；只列讀取層級）", fmt="0.00")
    r += 1
    bold("tbl", "C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = hdr(["項目", "主讀數：2026-04～09 累計 vs 層級 CY2026 × 6 ÷ 12", "第二讀數：2026-07～09 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    put("fm_a", "FM（ABF 所屬子部門）營收：公司實績", f"=B{R['h1']}", "¥B", "Verified：決算概要（4a）", c=f"=B{R['q2']}*4")
    put("fm_l", "FM 營收：層級隱含", "無同口徑隱含值", "", "Siliconomics 沒有 ABF 膜市場與味之素份額（4b）", fmt="General", c="無同口徑隱含值")
    put("j_fm", "FM 營收：判定", "無可比實績", "", "缺層級的 ABF 膜市場與份額；不計入對帳中", fmt="General", c="無可比實績")
    put("abf_a", "ABF 增層膜營收：公司實績", "無（未揭露）", "", "公司不揭露 ABF 單獨營收；不以 AI 占比估計或模型值補位", fmt="General", c="無（未揭露）")
    put("abf_l", "ABF 增層膜營收：層級隱含", "無同口徑隱含值", "", "Siliconomics 只有載板產值與顆數，缺膜用量、膜單價與味之素份額", fmt="General", c="無同口徑隱含值")
    put("j_abf", "ABF 增層膜營收：判定", "無可比實績", "", "兩側皆缺；不計入對帳中", fmt="General", c="無可比實績")
    r += 1
    bold("ref", "參考：本模型 FY2026 錨點對已公布實績與公司 11-09 上修（模型對實績與公司指引，不是層級對帳；不計入 C0）")
    put("anc", "FY2026 FM 營收錨點（In 頁 f_fm；公司 2026-08-06 預測）", "=f_fm", "¥B", "Interested-party：Q1 決算概要（公司預測，非實績）")
    put("prog", "上半年實績 ÷ 錨點（進度）", f"=B{R['h1']}/B{R['anc']}", "%", "Derived；全年 12 個月中已過 6 個月", fmt="0.0%")
    put("g1109", "公司 FY2026 FM 營收預測（2026-11-09 上修）", 133.0, "¥B", f"Interested-party（公司指引，只列參考）：公司 FY2026 セグメント別修正業績予想（2026-11-09）ファンクショナルマテリアルズ營收 1,330 億日圓（8-06 為 1,206；事業利益 765，原 655）；MoneyDJ 經 TechNews 2026-11-10 轉載相同（網址見 Sources 頁）；模型錨點本版不改", blue=True, url=REV1109)
    put("g_vs", "11-09 上修 ÷ 模型錨點 − 1", f"=B{R['g1109']}/B{R['anc']}-1", "%", "Derived；錨點更新另案（Q2 決算後更新，In 頁 E 節標題）", fmt="+0.0%;-0.0%")
    r += 1
    bold("sum0", "C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_fm"]}:C{R["j_abf"]},"對帳中")'
    put("sum1", F(T("C0（實績）結論：無可比實績：味之素 FY2026 上半年 FM 營收 "), TX(f"B{R['h1']}", "#,##0.0"), T(" 十億日圓、第二季 "), TX(f"B{R['q2']}", "#,##0.0"),
                  T(" 十億日圓為公司決算概要實績，但 Siliconomics 沒有 ABF 增層膜市場與味之素份額，公司也未揭露 ABF 單獨營收，兩讀數無法計算。對帳中項目 "), nopen, T(" 項。")))
    put("sum2", F(T("參考（模型對實績）：上半年 FM 營收已達本模型 FY2026 錨點 "), TX(f"B{R['anc']}", "#,##0.0"), T(" 十億日圓的 "), TX(f"B{R['prog']}", "0%"),
                  T("；公司 11-09 把全年電子材料營收預測上修到 "), TX(f"B{R['g1109']}", "#,##0.0"), T(" 十億日圓（較錨點 "), TX(f"B{R['g_vs']}", "+0%;-0%"),
                  T("），屬公司指引、只列參考；本版錨點不改（既有數值不變），錨點更新另案。")))
    put("sum3", "限制與待查：決算概要數字由本版建置時以網路讀取（2026-11-15）；全公司上半年數字另見決算短信。Siliconomics 若補 ABF 增層膜市場與味之素份額（或公司揭露 ABF 營收），即可改為可判定的 C0 列。")
    put("cnt", "對帳中項目數（C0 實績表；讀取層級同源列不計）", None)
    CW.cell(R["cnt"], 7, f'={nopen}')
    R["end"] = r - 1
    CW.column_dimensions["A"].width = 70; CW.column_dimensions["G"].width = 110; CW.column_dimensions["H"].width = 60
    return R


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, PLA, RVA, TOR, CR):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    rb = lab(P, "情境 B（基準）：")
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 60)}"
    ra = lab(P, "情境 A（樂觀）：")
    pa = lambda label, c: f"PL!{c}{lab(P, label, ra, exact=True, end=ra + 60)}"
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    dw = lab(SE, "加權目標價中 DCF 的權重", TOR["t0"])
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；{VER} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    FM = "ファンクショナルマテリアルズ（FM）營收"
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 日圓，較現價"), f'IF({up}<0,"低 ","高 ")', TX(f"ABS({up})", "0.0%"), T("，評等"), rec,
        T("。現金流折現 "), TX(dcf, "#,##0"), T(" 日圓、同業比較法（食品同業，相對倍數調整後）"), TX(cadj, "#,##0"), T(" 日圓各半；AI 曝險在 ABF 增層膜，營收占比不到一成。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：本業以食品同業倍數計價、ABF 成長由加速器顆數與產能上限決定（份額、漲價為假設），DCF 與同業比較法各占一半。DCF 權重降到三成時為"),
        REC5(f"((0.3*{dcf}+0.7*{cadj})/px-1)"), T("，同業比較法不加歷史倍數調整時加權目標價 "), TX(noadj, "#,##0"), T(" 日圓（"), REC5(f"({noadj}/px-1)"),
        T("），產業部署量改用另一側讀數時為"), f"Coord!C{CR['rec']}", T("。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 日圓，"), f'IF({up}<0,"低於","高於")', T("現價 "), TX("px", "#,##0"), T(" 日圓 "), TX(f"ABS({up})", "0.0%"), T("，評等"), rec,
        T("。同業比較法（食品同業中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數 "), TX(coef, "0.00"), T(" × 12 個月後 EPS "), TX(feps, "0"), T(" 日圓）得 "), TX(cadj, "#,##0"),
        T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(" 日圓高於本模型。"))
    put("o2", T("ABF 是成長來源但規模小：情境 B 的 FM 營收由 FY2026 公司預測 "), TX(pb(FM, "D"), "#,##0.0"), T(" 十億日圓到 FY2028 "), TX(pb(FM, "F"), "#,##0.0"),
        T("，占營收由 "), TX(pb("FM ÷ 營收", "D"), "0.0%"), T(" 到 "), TX(pb("FM ÷ 營收", "F"), "0.0%"), T("；FM 事業利益占事業利益 "), TX(pb("FM 事業利益 ÷ 事業利益", "F"), "0%"),
        T("。FY2027 起需求高於可出貨量上限，ABF 成長由產能與漲價決定。"))
    put("o3", T("現價隱含：要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"), T("（模型 "), TX("wacc", "0.0%"), T("），或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"),
        T(" 倍；現價相當於食品同業倍數的 "), TX(rv("rc1"), "0.00"), T(" 倍。DCF 權重 30%／70% 時加權目標價各 "), TX(f"Sens!B{dw}", "+0%;-0%"), T("／"), TX(f"Sens!C{dw}", "+0%;-0%"), T("。"))
    put("o4", T("情境差異：情境 A 的 FY2028 EPS "), TX(pa("EPS（日圓）", "F"), "0"), T(" 日圓，情境 B "), TX(pb("EPS（日圓）", "F"), "0"),
        T(" 日圓；差在 CY2028 起加速器顆數與 ABF 價格乘數。上半年 FM 營收已公布，C0（實績）無可比實績：Siliconomics 沒有 ABF 增層膜市場與味之素份額。"))
    put("o5", T("對帳中：CY2027–CY2029 Siliconomics 標對帳中，另一側讀數（Wattonomics 換算）目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 日圓（"), f"Coord!C{CR['rec']}",
        T("），較現行基準 "), TX(f"Coord!D{CR['tp']}", "+0.0%;-0.0%"), T("；本模型的加速器顆數與 ABF 產能上限讀 Siliconomics，個股不裁決，對帳在 Siliconomics。"))
    put("drv", T("分部營收推導（情境 B）：FM 營收＝FY2026 錨點 ×〔AI 用：有效加速器顆數成長 × 每顆載板面積層數 × 份額；非 AI：成長率〕× 需求滿足率（可出貨量上限）× ABF 價格指數；"
                  "調味料・食品、冷凍食品、ヘルスケア等（扣除 FM）、その他依成長率；相加為合併營收。FY2028 FM "), TX(pb(FM, "F"), "#,##0.0"), T("、調味料・食品 "), TX(pb("調味料・食品營收", "F"), "#,##0"),
        T("、冷凍食品 "), TX(pb("冷凍食品營收", "F"), "#,##0"), T("，合計 "), TX(pb("營收", "F"), "#,##0"), T(" 十億日圓。"))
    RD.cell(r, 1, f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA；同次產業 Ibiden、日股決算短信 Advantest）與三層協調。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；"
                  "新增 Hist 頁（損益表歷史 FY2022–FY2025，公司決算短信 IFRS）、RevDCF 頁、PL 頁末標準損益表（營業利益列＝IFRS 事業利益）、Val 0.7 四個目標價、Sens 頁末加權目標價相對基準、In 頁 N 節（營業費用彈性，只用於推算毛利）；"
                  "Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、C0（實績）（2027 年 3 月期 第 2 四半期 決算概要 FM 營收；結論：無可比實績，Siliconomics 沒有 ABF 增層膜市場與味之素份額）。"
                  "README A1 標題段版號由 v0.1 更正為現行版號。既有數值格不變，加權目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def sources(wb):
    SO = wb["Sources"]
    r = SO.max_row + 1
    for t, u in ((f"公司：{FR2Q[2:]}（FM 營收 上半年 733、第二季 401 億日圓；全公司 8,311；C0 實績用）", RES2Q),
                 ("公司：2027 年 3 月期 第 2 四半期 決算短信（2026-11-09；全公司上半年）", TAN2Q),
                 ("公司：FY2026 セグメント別修正業績予想（2026-11-09；ファンクショナルマテリアルズ營收 1,330、事業利益 765 億日圓；只列參考）", REV1109),
                 ("媒體：MoneyDJ 經 TechNews 2026-11-10（上半年電子材料營收 733 億、事業利益 429 億、全年上修；交叉核對）", TN1110),
                 ("公司：2024 年 3 月期決算短信 p.10 連結損益計算書（FY2023、FY2022；Hist 頁）", JPX24),
                 ("公司：2025 年 3 月期決算短信 p.10 連結損益計算書（FY2024；Hist 頁）", JPX25),
                 ("公司：2026 年 3 月期決算短信 摘要頁 iXBRL（FY2025 稅前利益、當期利益；Hist 頁）", JPX26S),
                 ("資料商：stockanalysis 損益表（FY2025 毛利 597,149 百萬日圓；未對決算短信原文；Hist 頁）", SAI)):
        SO.cell(r, 1, t); SO.cell(r, 2, u); r += 1


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(r"v\d+(?:\.\d+)+（", f"{VER}（", head, count=1)
    assert n == 1, head
    head2 = re.sub("　20\\d\\d-\\d\\d-\\d\\d", "　" + BUILT, head2, count=1)
    RD["A1"].value = head2 + sep + rest + (f"；{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價相對基準敏感度）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
                                           "A1 標題版號由舊的 v0.1 更正；既有數值不變")


def apply(wb, SIL, CR):
    HR = hist_is(wb)
    in_n(wb)
    G = coord_c0(wb, SIL, CR)
    R4 = c0_actuals(wb, G, SIL)
    PLA = pl_standard(wb, HR, CR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    TOR = sens_tor(wb, CR)
    RMA = readme13(wb, PLA, RVA, TOR, CR)
    sources(wb)
    version(wb)
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, TOR=TOR, G=G, R4=R4, RM=RMA)
    return (f"v035_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {TOR['t0']} 列；"
            f"Coord C0 第 {G['c0']} 列起、C0（實績）第 {R4['title']}–{R4['end']} 列；README v1.3 白話列 {len(RMA)} 列")
