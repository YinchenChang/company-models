# Arm v0.3.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0 公司實績對帳、硬體股兩讀數）。
# 依據：docs/workorders/20261114_第二批13家_v1.3與三層協調.md、docs/workorders/20261114_C0實績對帳補列_七家.md、
#       docs/plan/三層基準協調規則.md v1.4.2（C0 第 2 點、C4）、docs/plan/個股HTML版面規格.md v1.3、tools/coord/README.md。
# 樣板：Broadcom v0.5.6／v0.5.7（avgo/v056_post.py）、AMD v0.4.7（amd/v047_post.py）、Credo v0.3.6（crdo/v036_post.py）。只新增，不改既有數值格：
#   Hist 頁末（新節）：損益表歷史（FY2024 non-GAAP，Arm 6-K 股東信 GAAP 對 non-GAAP 調節表；FY2025–FY2026 連 Q 頁；GAAP 對照三年）
#   PL 頁末：標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B；Non-GAAP）、比率、GAAP 對照、核對列
#   RevDCF 頁（新）；Val 頁末 0.7 四個目標價；Sens 頁末加權目標價相對基準（營運輸入＝本頁主表整本重算快照 arm/sens.json，低／高兩欄並列）
#   Coord 頁末 F–J 節；具名範圍 SIF_RecoFlag_<年>（指向 Coord 頁對帳中旗標列；FY(t)≈CY(t−1)，供版面檢查依會計年度對應）
#   README 頁末：v1.3 展示層白話列；A1、A2 版號；build_id 前綴（README A1 與 Sens 底部）
# build_arm.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL, CR)。COORD_RUN（alt_run.py）或 ARM_OVERRIDE（sens_arm.py）時不執行。
import os, re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, OLD, BUILT = "0.3.5", "0.3.4", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
ED = "https://www.sec.gov/Archives/edgar/data/1973239/"
SRC = {"FY2024": ED + "000197323925000010/exhibit992fye25q431-marx25.htm（Arm 6-K 2025-05-07 EX-99.2 股東信 Q4 FY25，第 15 頁 GAAP 對 non-GAAP 調節表：FY2024 為比較欄）",
       "FY2025": ED + "000197323925000010/exhibit992fye25q431-marx25.htm（Arm 6-K 2025-05-07 EX-99.2 股東信 Q4 FY25，第 15 頁調節表）",
       "FY2026": ED + "000197323926000062/exhibit992fye26q431-marx26.htm（Arm 6-K 2026-05-06 EX-99.2 股東信 Q4 FY26，調節表）",
       "FY2023": ED + "000197323924000007/exhibit992fye24q431-marx24.htm（Arm 6-K 2024-05-08 EX-99.2 股東信 Q4 FY24，FY2023 調節表；只用營收與 non-GAAP 淨利作年增率基期）"}
SRC_Q = {"Q1": ED + "000197323926000113/exhibit992fye27q130-junx26.htm（Arm 6-K 2026-07-29 EX-99.2 股東信 Q1 FY27）",
         "Q2": ED + "000197323926000141/exhibit992fye27q230-sepx26.htm（Arm 6-K 2026-11-04 EX-99.2 股東信 Q2 FY27）"}
PCODE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+(?![A-Za-z0-9])")


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
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


# ---------- Hist：損益表歷史（既有 Hist 頁末新節） ----------
# (鍵, 項目, FY2024, FY2025, FY2026, FY2023（基期）, 標記)；十億美元、EPS 美元、股數十億股。FY2025、FY2026 能連 Q 頁者連 Q 頁（Q 頁 B、C 欄＝公司全年值）
HIS = [
    ("rev", "營收（GAAP＝non-GAAP）", 3.233, "=Q!B6", "=Q!C6", 2.679, "Verified"),
    ("cogs", "營業成本（non-GAAP）", 0.102, "=Q!B7", "=Q!C7", None, "Verified"),
    ("gp", "毛利（non-GAAP）＝營收 − 營業成本", "=C{rev}-C{cogs}", "=D{rev}-D{cogs}", "=E{rev}-E{cogs}", None, "Derived"),
    ("opx", "營業費用（non-GAAP；研發＋銷管）", 1.723, "=Q!B8", "=Q!C8", None, "Verified"),
    ("oi", "營業利益（non-GAAP）", 1.408, "=Q!B9", "=Q!C9", None, "Verified"),
    ("oth", "其他收入（淨額；利息收入＋其他業外，non-GAAP）", 0.121, 0.126, 0.107, None, "Verified"),
    ("pti", "稅前利益（non-GAAP）", 1.529, 1.997, 2.222, None, "Verified"),
    ("tax", "所得稅（non-GAAP）", 0.205, 0.260, 0.333, None, "Verified"),
    ("ni", "稅後淨利（non-GAAP）", 1.324, "=Q!B13", "=Q!C13", 0.657, "Verified"),
    ("sh", "non-GAAP 稀釋加權股數（十億股）", 1.044, "=Q!B14", "=Q!C14", None, "Verified"),
    ("eps", "EPS（non-GAAP 稀釋，公司公告；美元）", 1.27, "=Q!B16", "=Q!C16", None, "Verified"),
    ("g_ni", "稅後淨利（GAAP）", 0.306, 0.792, 0.904, None, "Verified"),
    ("g_eps", "EPS（GAAP 稀釋，公司公告；美元）", 0.29, 0.75, 0.85, None, "Verified"),
]


def hist_pl(wb):
    H, Q = wb["Hist"], wb["Q"]
    for rr, t in ((6, "營收"), (7, "營業成本（non-GAAP）"), (8, "營業費用（non-GAAP；研發＋銷管）"), (9, "營業利益（non-GAAP）"), (13, "稅後淨利（non-GAAP）"), (14, "稀釋加權股數（十億股）")):
        assert str(Q.cell(rr, 1).value).strip() == t, (rr, Q.cell(rr, 1).value)
    assert str(Q.cell(16, 1).value).strip().startswith("公司揭露 non-GAAP EPS") and Q.cell(3, 2).value == "FY2025A" and Q.cell(3, 3).value == "FY2026A"
    r = H.max_row + 3
    r0 = r
    H.cell(r, 1, f"損益表歷史（v{VER} 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。十億美元；EPS 美元；股數十億股。會計年度 FY(t) 於 t 年 3 月 31 日結束。"
                 "non-GAAP 為公司口徑（排除股票薪酬、相關雇主稅、併購攤銷、權益投資損益等）；FY2025、FY2026 能連 Q 頁者連 Q 頁（與公司全年值一致），其他收入、稅前利益、所得稅與 GAAP 三年取自股東信調節表").font = BOLD
    r += 2
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "FY2023（只供年增率基期）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v24, v25, v26, v23, tag in HIS:
        H.cell(r, 1, f"hp_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "24", v24), ("D", "25", v25), ("E", "26", v26), ("F", "23", v23)):
            if v is None:
                continue
            if isinstance(v, str) and "{" in v:
                v = v.format(**rows)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"hp_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, tag)
        src = f"FY2024：{SRC['FY2024']}；FY2025：" + ("Q 頁 B 欄（" + SRC["FY2025"] + "）" if isinstance(v25, str) and v25.startswith("=Q") else SRC["FY2025"]) \
              + "；FY2026：" + ("Q 頁 C 欄（" + SRC["FY2026"] + "）" if isinstance(v26, str) and v26.startswith("=Q") else SRC["FY2026"])
        if v23 is not None: src += f"；FY2023：{SRC['FY2023']}"
        H.cell(r, 8, src)
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    for labl, f in (("營業利益＋其他收入 − 稅前利益（non-GAAP）", "={c}{oi}+{c}{oth}-{c}{pti}"),
                    ("稅前利益 − 所得稅 − 稅後淨利（non-GAAP）", "={c}{pti}-{c}{tax}-{c}{ni}"),
                    ("毛利 − 營業費用 − 營業利益（non-GAAP）", "={c}{gp}-{c}{opx}-{c}{oi}"),
                    ("稅後淨利 ÷ 稀釋股數 − 公司公告 EPS（四捨五入差）", "={c}{ni}/{c}{sh}-{c}{eps}")):
        H.cell(r, 2, labl)
        for c in "CDE":
            H[f"{c}{r}"] = f.format(c=c, **rows); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "備註：其他收入（淨額）＝利息收入淨額＋其他業外淨額（non-GAAP，排除權益投資損益與處分利益）：FY2024 110＋11、FY2025 116＋10、FY2026 111 −4（百萬美元）。").font = Font(italic=True)
    return dict(start=r0, rows=rows)


# ---------- PL：標準損益表 ----------
YEARS = ["FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJK"   # C–E＝Hist（FY2024A–FY2026A）；F–H＝PL 情境 B 推導區塊 E–G（FY2027 錨點、FY2028–FY2029）；I–K＝PL 延伸表情境 B 的 H–J（FY2030–FY2032，收斂法）


def pl_standard(wb):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    hb = rb + 1
    assert [P.cell(hb, c).value for c in (3, 4, 5, 6, 7)] == ["FY2025A", "FY2026A", "FY2027（錨點）", "FY2028", "FY2029"], [P.cell(hb, c).value for c in range(3, 8)]
    rd = lab(P, "D. 損益（non-GAAP）", rb)
    g = lambda lbl, s=rd: lab(P, lbl, s, exact=True)
    R = dict(rev=lab(P, "營收", lab(P, "C. 授權及其他", rb), exact=True), cogs=lab(P, "營業成本（", rd), gp=g("毛利"), opx=g("營業費用（non-GAAP）"), oi=g("營業利益（non-GAAP）"),
             ni=g("稅後淨利（non-GAAP）"), sh=g("稀釋股數"), eps=g("EPS（non-GAAP）"))
    xb = lab(P, "延伸｜情境 B")
    assert [P.cell(xb + 1, c).value for c in (8, 9, 10)] == ["FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）"]
    X = dict(rev=lab(P, "營收（延伸）", xb), ni=lab(P, "淨利（延伸）", xb), sh=lab(P, "股數（延伸）", xb), eps=lab(P, "EPS（延伸）", xb))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋未來六年（FY2027E 錨點、FY2028E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "會計年度 FY(t) 於 t 年 3 月 31 日結束。只引用本頁上方與 Hist 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。Arm 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股票薪酬與相關雇主稅、併購攤銷、權益投資損益），故本表用 Non-GAAP；"
                 "其他收入（淨額）＝利息收入＋其他業外淨額。GAAP 淨利與 EPS 列在本表下方對照（只列實績）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2027E＝錨點（1Q27、2Q27 實績＋3Q27 公司指引＋4Q27 假設；Arm AGI CPU 晶片成本另扣）；FY2028E–FY2029E＝情境 B 由下而上（權利金分市場＋授權及其他）；"
                 "FY2030E–FY2032E＝收斂法（PL 延伸表只推營收與淨利；營業利益＝淨利 ÷（1 − 稅率）− 業外收益，同 DCF 頁）。收斂期毛利率沿用 FY2029E，"
                 "據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列（目標價用收斂期淨利，不受影響）。EPS＝non-GAAP 稀釋")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r_hdr = r; r += 1
    L = {}
    NC = 3 + len(YEARS)

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, NC, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "EFG"]            # FY2027E–FY2029E（PL 情境 B 推導區塊 E–G）
    ex = lambda key: [f"=PL!{c}{X[key]}" for c in "HIJ"]            # FY2030E–FY2032E（延伸表情境 B 的 H–J）
    hh = lambda k: [f"=hp_{k}_24", f"=hp_{k}_25", f"=hp_{k}_26"]
    LAST = "H"   # 本表 FY2029E 欄
    EXT = "IJK"
    row("rev", "營收", "US$B", hh("rev") + fc("rev") + ex("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/hp_rev_23-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 的基期＝FY2023 營收（Hist 頁）")
    r_c = r
    row("cogs", "營業成本（COGS）", "US$B", hh("cogs") + fc("cogs") + [f"={c}{L['rev']}*(1-${LAST}{r_c + 2})" for c in EXT], M,
        "FY2027E 起含 Arm AGI CPU 晶片成本（晶片 ×（1 − 晶片毛利率））；FY2030E–FY2032E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp") + [f"={c}{L['rev']}-{c}{L['cogs']}" for c in EXT], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx") + [None] * 3, M,
        "FY2030E–FY2032E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）（non-GAAP，不含股票薪酬）")
    r_ni = r_c + 9
    row("oi", "營業利益", "US$B", hh("oi") + fc("oi") + [f"={c}{r_ni}/(1-tx_f)-oth_f" for c in EXT], M, "FY2030E–FY2032E 由淨利反推（同 DCF 頁；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝利息收入＋其他業外", "US$B", hh("oth") + ["=Q!H11", "=oth_f", "=oth_f", "=oth_f", "=oth_f", "=oth_f"], M,
        "FY2027E＝錨點業外收益（Q 頁 H11：1Q27、2Q27 實績＋3Q27、4Q27 每季假設）；FY2028E 起＝In 頁每年業外收益")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", hh("tax") + [f"={c}{L['pti']}-{c}{r_ni}" for c in "FGHIJK"], M,
        "歷史＝non-GAAP 所得稅；預測＝稅前 − 稅後淨利（non-GAAP 有效稅率約 15%，英國 OECD 最低稅）")
    assert L["tax"] == r_c + 8 and r == r_ni
    row("ni", "稅後淨利", "US$B", hh("ni") + fc("ni") + ex("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數", "十億股", hh("sh") + fc("sh") + ex("sh"), M)
    row("eps", "EPS（non-GAAP 稀釋）", "US$", hh("eps") + fc("eps") + ex("eps"), "0.00",
        "FY2024A–FY2026A＝公司公告 non-GAAP 稀釋 EPS；預測＝淨利 ÷ 稀釋股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/hp_ni_23-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 的基期＝FY2023 non-GAAP 淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_ni", "GAAP：稅後淨利", "US$B", ["=hp_g_ni_24", "=hp_g_ni_25", "=hp_g_ni_26"], M)
    row("g_eps", "GAAP：EPS（稀釋，公司公告）", "US$", ["=hp_g_eps_24", "=hp_g_eps_25", "=hp_g_eps_26"], "0.00", "含股票薪酬（FY2026 約 10.5 億美元）與權益投資損益")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2024A–FY2029E）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:6]] + [None] * 3, "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E–FY2032E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "IJK":
        P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"H{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E–FY2032E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "US$B", [None] * 6 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "IJK"], "0.000")
    row("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E–FY2032E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 5 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "HIJK"], "0.00%")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 區塊與延伸表（FY2025A–FY2032E）", "US$B",
        [None, f"=D{L['ni']}-PL!C{R['ni']}", f"=E{L['ni']}-PL!D{R['ni']}"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGH", "EFG")] + [f"={c}{L['ni']}-PL!{s}{X['ni']}" for c, s in zip("IJK", "HIJ")], "0.000")
    row("c5", "核對：EPS − 情境 B 區塊與延伸表（FY2027E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGH", "EFG")] + [f"={c}{L['eps']}-PL!{s}{X['eps']}" for c, s in zip("IJK", "HIJ")], "0.000")
    row("c6", "核對：稅前利益 − 營業利益 − 其他收入", "US$B", [f"={c}{L['pti']}-{c}{L['oi']}-{c}{L['oth']}" for c in COLS], "0.000")
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
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（v{VER} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。明確期第一年 FY2027（錨點；評價日後只含 2027-01～03）維持模型值不動，只動 FY2028–FY2032（含終值）。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    yr("ebit", "營業利益 EBIT（non-GAAP 營業利益 − 股票薪酬；US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元 EBIT 對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
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
    kv("w_ok", "判讀（隱含 WACC）", f'=IF(ABS(B{A["w_res"]})<0.01,"已收斂","搜尋下限仍不足：折現率降到下限（永續成長率＋0.5 個百分點）時每股價值仍低於支持現價所需值")', "@")
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
                     T(" 倍。單靠一個因素達成時，需要："), f'IF(ABS(B{A["w_res"]})<0.01,"折現率降到 "&TEXT(B{A["w1"]},"0.0%"),"折現率即使降到 "&TEXT(B{A["w1"]},"0.0%")&"（搜尋下限）仍不足")', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或 EBIT 率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 12 個月後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(" 倍（同業中位數 × 相對倍數係數 "), f'TEXT({coef},"0.00")', T("；Arm 上市未滿十年，資料不足，不調整）。換言之，市場目前給 Arm 的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("模型採用的係數。")))
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
    V.cell(r, 1, f"0.7 估值並列：四個目標價（v{VER} 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
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




# ---------- Sens：加權目標價相對基準（版面規格 v1.3 H 節） ----------
def _plain(t):
    """Sens 主表列名 → 白話（去掉「：低／高」與含參數代號的括號說明）。"""
    t = re.sub(r"：(低|高)$", "", str(t).strip())
    while PCODE.search(t):
        m = PCODE.search(t)
        i = max(t.rfind("（", 0, m.start()), t.rfind("；", 0, m.start()))
        if i <= 0:
            t = PCODE.sub("", t); break
        j = t.find("）", m.end())
        t = (t[:i] + ("）" if t[i] == "；" and j >= 0 else "") + (t[j + 1:] if j >= 0 else "")).strip()
    return t


def sens_tor(wb):
    SE, V = wb["Sens"], wb["Val"]
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")
    f35 = str(V.cell(lab(V, "12 個月後的前瞻 EPS（情境加權）"), 2).value)
    fw = dict(re.findall(r"tp_w([ABC])\*(\([^)]*\))", f35))
    assert set(fw) == set("ABC"), f35
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    # 主表（第 5 列起）：一列一個案例（C 欄＝低／高，F 欄＝鍵，D 欄＝重算加權目標價）；依鍵配對低／高
    assert SE.cell(3, 1).value == "驅動" and SE.cell(3, 6).value == "鍵" and SE.cell(4, 1).value == "基準"
    pairs = {}; order = []
    r = 5
    while SE.cell(r, 1).value and SE.cell(r, 6).value:
        k, side = SE.cell(r, 6).value, SE.cell(r, 3).value
        if k not in pairs: pairs[k] = {}; order.append(k)
        pairs[k][side] = r; r += 1
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    par = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        par.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；v{VER} 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝引用本頁主表（arm/sens_arm.py 整本重算快照，三情境同時套用），低案、高案並列一列；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例中主表沒有的，引用本頁「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    frame = [
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fw['C']}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*{fw['A']}",
         "Val 頁單一情境 DCF 價值；前瞻 EPS＝該情境 FY2028、FY2029 EPS 依月數加權（同 Val 頁公式）"),
        ("DCF 在加權中的權重 100%（只看現金流折現）／0%（只看同業比較法）（基準＝50%）", f"={dcf}", f"={cadj}", "Val 頁：DCF 與 Comparable（調整後）目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    KR = {}
    for k in order:
        if k == "tp_w_dcf":
            continue   # 估值框架：已由上方「DCF 在加權中的權重 100%／0%」涵蓋
        d = pairs[k]
        any_r = d.get("低") or d.get("高")
        SE.cell(r, 1, _plain(SE.cell(any_r, 1).value))
        for side, col, tcol in (("低", 2, 4), ("高", 3, 5)):
            if side not in d:
                continue
            SE.cell(r, tcol, f"=Sens!D{d[side]}").number_format = "#,##0.00"
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{d.get('低', any_r)}／A{d.get('高', any_r)}（本頁主表，整本重算快照）"); KR[k] = r; r += 1
    main_keys = set(order)
    for sr in par:
        key = SE.cell(sr, 8).value
        if key in main_keys or key == "beta":
            continue   # 主表已有（WACC、β、ERP、資本支出、營運資金、DCF 權重）
        t = str(SE.cell(sr, 1).value)
        SE.cell(r, 1, _plain(t))
        for col, tcol, scol in ((2, 4, "C"), (3, 5, "E")):
            v = SE[f"{scol}{sr}"].value
            if not isinstance(v, (int, float)) and not (isinstance(v, str) and v.startswith("=")):
                continue
            SE.cell(r, tcol, f"=Sens!{scol}{sr}").number_format = "#,##0.00"
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「目標價參數敏感度」表）"); r += 1
    last = r - 1
    r_st = lab(SE, "模型狀態（CURRENT")
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f"=Sens!D{r_st}")
    SE.cell(r, 3, "引用本頁底部模型狀態（CURRENT＝主表快照與本檔輸入一致；STALE＝In 頁已改、須重跑 python3 arm/sens_arm.py 與建置）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1, KR=KR)


# ---------- Coord：讀取層級、讀取層級核對、層級之間、損益表年度標示、C0（實績） ----------
def coord(wb, SIL):
    C, I, Q, P = wb["Coord"], wb["In"], wb["Q"], wb["PL"]
    N = SIL.NAMES
    r = C.max_row + 2
    ys = (2026, 2027, 2028, 2029)
    r_flag = lab(C, "對帳中旗標")
    for j, y in enumerate(ys):   # FY(t)＝9/12 × CY(t−1)＋3/12 × CY(t)：版面檢查依具名旗標以會計年度對應（同 NVIDIA、Marvell、Credo）；值連 Coord 頁旗標列
        dn(wb, f"SIF_RecoFlag_{y}", f"Coord!${'BCDE'[j]}${r_flag}")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    C.cell(r, 1, "F. 讀取層級（工作單 20261114 第二批第 2 節）：本模型讀哪一層、哪一版、哪些名稱。自層級讀入的同源值只列在此節與 G 節，不進 C0 判定").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("NVIDIA 資料中心 GPU 顆數 nvbw_／nvrb_／nvru_（CY2026–CY2029；SIF_UnitsPlat_bw／rb／ru）", "Siliconomics", src_in("nvbw_26").split("｜")[0],
         "NVIDIA Grace／Vera CPU 顆數＝GPU × 每顆搭配 CPU；資料中心 CPU 顆數指數的 NVIDIA 部分（FY2028 起影響資料中心權利金）", "In 頁資料中心節"),
        ("情境顆數係數 uc28／uc29（情境 B 0.80／0.710；um28／um29）", "市場層 markets/scen（與 Siliconomics 同值）", src_in("uc28_B").split("｜")[0], "情境 B、C 的 NVIDIA 顆數與雲端自研 CPU 顆數指數", "In 頁資料中心節"),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中"),
        ("Wattonomics 錨定係數 wta_27–wta_29（舊敏感度案例；SIF_WTAnchor_Base）", "Wattonomics（經 Siliconomics 轉入）", vtxt, "Sens 頁「Wattonomics 錨定」案例", "舊口徑，保留供對照；另一側讀數取代其角色"),
        ("NVIDIA 資料中心 GPU 顆數 CY2025 nv_cy25（FY2026 指數基期）", "未讀層級（In 頁 Assumed）", "—", "只用於 FY2026 的 NVIDIA CPU 顆數（指數基期）", "Siliconomics 加速器顆數自 CY2026 起；待併入"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。層級（Siliconomics、markets/scen）沒有 Arm 的權利金、份額或營收歸屬，也沒有讀本模型；"
                 "層級值只透過 NVIDIA 顆數成長進入 FY2028 起的資料中心權利金，FY2027 錨點不讀層級（見 J 節）。").font = Font(italic=True); r += 2
    C.cell(r, 1, "G. 讀取層級核對（同源，不判定）：個股讀入的層級值 vs 層級值，必然相同；只證明本模型讀到現行版，不計入對帳中").font = BOLD; r += 1
    for j, h in enumerate(["項目", "本模型值", "層級對照值", "差距", "狀態", "單位與口徑", "對照來源（層、版本）", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    up = lambda y: sum(N[f"SIF_UnitsPlat_{k}"][y - 2025] for k in ("bw", "rb", "ru"))
    for labl, f_m, v_l, unit, src, fmt in (
            ("NVIDIA 資料中心 GPU 顆數 CY2026（Blackwell＋Rubin＋Rubin Ultra）", "=nvbw_26+nvrb_26+nvru_26", up(2026), "百萬顆", f"Siliconomics {SIL.VERSION}（SIF_UnitsPlat_bw＋rb＋ru，2026）", "0.000"),
            ("NVIDIA 資料中心 GPU 顆數 CY2027", "=nvbw_27+nvrb_27+nvru_27", up(2027), "百萬顆", f"Siliconomics {SIL.VERSION}（SIF_UnitsPlat_*，2027）", "0.000"),
            ("情境 B：CY2028 顆數係數", "=uc28_B", SIL.base("um28_B"), "x", f"Siliconomics {SIL.VERSION}（um28_B）", "0.000"),
            ("情境 B：CY2029 顆數係數", "=uc29_B", SIL.base("um29_B"), "x", f"Siliconomics {SIL.VERSION}（um29_B）", "0.000")):
        C.cell(r, 1, labl); C.cell(r, 2, f_m).number_format = fmt; C.cell(r, 3, v_l).number_format = fmt; C.cell(r, 3).font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"; C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, unit); C.cell(r, 7, src); C.cell(r, 8, "Derived（層）")
        C.cell(r, 9, "本模型直接讀層級值；恆等，差距應為 0"); r += 1
    G1 = r - 1
    r += 1
    C.cell(r, 1, "H. 層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：產業部署 GW，Siliconomics 現行基準 vs 另一側讀數（Wattonomics 換算）。這是兩層之間的差距，不是 Arm 自己的量；本模型依 C4 並列兩讀數目標價").font = BOLD; r += 1
    for j, h in enumerate(["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態", "單位與口徑", "對照來源", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    L0 = r
    for j, y in enumerate(ys):
        C.cell(r, 1, f"產業部署 GW CY{y}")
        C.cell(r, 2, N["SIF_DeployGW_Base"][j]).number_format = "#,##0.00"; C.cell(r, 3, N[f"SIF_DeployGW_Alt_{y}"]).number_format = "#,##0.00"
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        C.cell(r, 5, "層間參考（Siliconomics 標對帳中；本公司不計）" if SIL.RECO_FLAG[y] else "層間參考（兩層一致）")
        C.cell(r, 6, "GW；SIF_DeployGW_Base vs SIF_DeployGW_Alt_<年>")
        C.cell(r, 7, f"Siliconomics {SIL.VERSION}（旗標 SIF_RecoFlag_{y}＝{int(SIL.RECO_FLAG[y])}）；Siliconomics 對帳單 r3")
        C.cell(r, 8, "Derived（基準）／Wattonomics 換算（另一側）"); C.cell(r, 9, "兩層之間的差距，不裁決；本模型 Coord 頁 A–E 節並列兩讀數目標價")
        r += 1
    L1 = r - 1
    r += 1
    C.cell(r, 1, "I. 對帳中年份與損益表年度標示（C4；Arm FY(t)＝9/12 × CY(t−1)＋3/12 × CY(t)，FY2027 為錨點（指數分母）、另一側讀數不縮放；FY2028 起含對帳中年份 CY2027；FY2030E–FY2032E 為收斂期，依賴 CY2029）").font = BOLD; r += 1
    FYS_ = list(range(2027, 2033))
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in FYS_], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年（主要）")
    dep = [2026, 2027, 2028, 2029, 2029, 2029]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2027E–FY2032E）")
    for j, fy in enumerate(FYS_):
        C.cell(r, 2 + j, f'=IF(${"BCDE"[dep[j] - 2026]}${r_flag}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    # ---- J：C0（實績） ----
    rbB = lab(P, "情境 B：")
    rdc, rnc, rro = lab(P, "權利金：資料中心", rbB), lab(P, "NVIDIA Grace／Vera CPU 顆數", rbB), lab(P, "權利金營收（Royalty）", rbB)
    C.cell(r, 1, "J. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）").font = BOLD; j_title = r; r += 1
    C.cell(r, 1, "方法：本模型自層級讀入的量是 NVIDIA 資料中心 GPU 顆數（→ Grace／Vera CPU 顆數），只用在 FY2028 起資料中心權利金的成長（NVIDIA 占 FY2027 資料中心權利金的比重為 Assumed 35%，其餘為雲端自研 CPU，不讀層級）。"
                 "同口徑的公司實績應是資料中心權利金金額或 NVIDIA Grace／Vera 權利金：Arm 不揭露（只揭露授權及其他、權利金兩類營收；資料中心只口述年增倍數：1Q27「逾一倍」、2Q27「一倍」，"
                 "FY2026 占權利金比重口述「十幾到接近 20%」）。口述倍數不算公司實績（Assumed／Interested-party），層級也沒有 Arm 的資料中心權利金，故無同口徑可比量。"
                 "下表兩列只列參考：(1) 公司口述的資料中心權利金年增 vs 層級隱含的 NVIDIA Grace／Vera CPU 顆數年增（口徑不同：後者不含雲端自研 CPU 與每顆權利金率上升）；"
                 "(2) 公司口述 vs 本模型 FY2027 錨點的資料中心權利金年增（Assumed 70%；這是本模型假設對公司說法，不是層級）。時差：FY2027＝2026-04～2027-03；已公布 1Q27、2Q27（2026-04～09，6 個月）。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); j_note = r; r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級", "網址／名稱"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000", url=""):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); C.cell(r, 7, url); IN[key] = r; r += 1
    assert str(Q.cell(3, 4).value).startswith("1Q27") and str(Q.cell(3, 5).value).startswith("2Q27") and str(Q.cell(5, 1).value).startswith("權利金營收")
    inp("ro", "權利金營收 1Q27＋2Q27（公司 6-K；Verified）", "=Q!D5+Q!E5", "US$B", "2026-04～2026-09（1H FY2027）", "Q 頁 D5、E5（股東信 1Q27 7.15、2Q27 6.70 億美元）", "Verified", GREEN, url=SRC_Q["Q2"])
    inp("q1", "資料中心權利金年增 1Q27（公司口述「逾一倍」；取 100%）", 1.0, "%", "1Q27（2026-04～06）對 1Q26", "Arm 6-K 2026-07-29 股東信：data center royalties more than doubling year over year", "Interested-party（口述，非金額）", BLUE, fmt="0%", url=SRC_Q["Q1"])
    inp("q2", "資料中心權利金年增 2Q27（公司口述「一倍」）", 1.0, "%", "2Q27（2026-07～09）對 2Q26", "Arm 6-K 2026-11-04 股東信：data center royalties doubling year over year", "Interested-party（口述，非金額）", BLUE, fmt="0%", url=SRC_Q["Q2"])
    inp("nv26", "層級：NVIDIA 資料中心 GPU 顆數 FY2026（會計年度加權；CY2025 為 In 頁 Assumed）", f"=PL!D{rnc - 1}", "百萬顆", "FY2026＝9/12 CY2025＋3/12 CY2026", f"PL 情境 B 區塊 D{rnc - 1}", "Derived（層＋Assumed 基期）", GREEN, url=f"PL!D{rnc - 1}")
    inp("nv27", "層級：NVIDIA 資料中心 GPU 顆數 FY2027（會計年度加權）", f"=PL!E{rnc - 1}", "百萬顆", "FY2027＝9/12 CY2026＋3/12 CY2027", f"PL 情境 B 區塊 E{rnc - 1}（Siliconomics SIF_UnitsPlat_*）", "Derived（層）", GREEN, url=f"PL!E{rnc - 1}")
    inp("lg", "層級隱含：NVIDIA Grace／Vera CPU 顆數年增 FY2027（＝GPU 顆數年增；每顆 GPU 搭配 CPU 不變）", f"=PL!E{rnc}/PL!D{rnc}-1", "%", "FY2027 對 FY2026", f"PL 情境 B 區塊 E{rnc}÷D{rnc}", "Derived（層）", fmt="0.0%")
    inp("mg", "本模型：FY2027 資料中心權利金年增（錨點，Assumed）", f"=PL!E{rdc}/PL!D{rdc}-1", "%", "FY2027 對 FY2026", f"PL 情境 B 區塊 E{rdc}÷D{rdc}（In 頁 FY2027 資料中心權利金年增）", "Assumed", GREEN, fmt="0.0%", url=f"PL!E{rdc}")
    inp("mdc", "本模型：FY2027 資料中心權利金（錨點）", f"=PL!E{rdc}", "US$B", "FY2027", f"PL 情境 B 區塊 E{rdc}", "Assumed（FY2026 比重口述 × 年增假設）", GREEN, url=f"PL!E{rdc}")
    r += 1
    for j, h in enumerate(["項目", "公司（口述）", "層級／模型隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    J0 = r
    b = IN; pct = "+0.0%;-0.0%"
    C.cell(r, 1, "參考（口徑不同）：資料中心權利金年增（公司口述，1H FY2027）vs 層級隱含 NVIDIA Grace／Vera CPU 顆數年增 FY2027")
    C.cell(r, 2, f"=AVERAGE(B{b['q1']},B{b['q2']})").number_format = "0%"
    C.cell(r, 3, f"=B{b['lg']}").number_format = "0.0%"
    C.cell(r, 4, f"=(1+C{r})/(1+B{r})-1").number_format = pct
    C.cell(r, 5, "參考（口徑不同）")
    C.cell(r, 6, "年增率；差距＝（1＋層級年增）÷（1＋公司口述年增）− 1（水準比）")
    C.cell(r, 7, "公司：股東信口述；層級：Siliconomics 顆數（PL 情境 B 區塊）"); C.cell(r, 8, "Interested-party（口述）／Derived（層）")
    C.cell(r, 9, "公司口徑含雲端自研 CPU（Graviton、Axion、Cobalt，約占六成五）與每顆權利金率上升（核心數、運算子系統）；層級只有 NVIDIA 顆數。差距不代表層級錯，不判定、不開對帳單")
    rm = r; r += 1
    C.cell(r, 1, "參考（本模型錨點）：資料中心權利金年增（公司口述，1H FY2027）vs 本模型 FY2027 資料中心權利金年增")
    C.cell(r, 2, f"=AVERAGE(B{b['q1']},B{b['q2']})").number_format = "0%"
    C.cell(r, 3, f"=B{b['mg']}").number_format = "0.0%"
    C.cell(r, 4, f"=(1+C{r})/(1+B{r})-1").number_format = pct
    C.cell(r, 5, "參考（本模型假設，非層級）")
    C.cell(r, 6, "年增率；差距＝FY2027 資料中心權利金水準（模型 ÷ 公司口述年增推得）− 1")
    C.cell(r, 7, "公司：股東信口述；模型：In 頁 FY2027 資料中心權利金年增（Assumed）"); C.cell(r, 8, "Interested-party（口述）／Assumed")
    C.cell(r, 9, '="本模型假設全年年增低於上半年口述（下半年基期較高）；若全年維持一倍，FY2027 資料中心權利金約高 "&TEXT(-D' + str(r) + ',"0%")&"（約 "&TEXT(B' + str(b["mdc"]) + '*(1+B' + str(r) + ')/(1+C' + str(r) + ')-B' + str(b["mdc"]) + ',"0.00")&" 十億美元，由物聯網與其他殘差吸收，權利金合計不變）。屬本模型假設複核，不計 C0"')
    rs = r; r += 1
    C.cell(r, 1, "C0（實績）判定：Arm 資料中心權利金（層級驅動的唯一項目）").font = BOLD
    C.cell(r, 5, "無可比實績").font = BOLD
    C.cell(r, 6, "缺：資料中心權利金金額、或 NVIDIA Grace／Vera 權利金、或 Arm 資料中心晶片出貨顆數（公司只口述年增倍數；FY2026 比重只有區間）")
    C.cell(r, 7, "Arm 6-K 股東信（1Q27、2Q27）、20-F FY2026")
    C.cell(r, 8, "—")
    C.cell(r, 9, '="上表兩列只供參考，不判定、不開對帳單；公司若在 20-F 或法說揭露資料中心權利金金額或比重，改以金額對層級隱含（NVIDIA 顆數 × 每顆權利金＋雲端自研）檢查"')
    rv = r; r += 1
    J1 = r - 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（J 節 C0（實績）判定；參考列與讀取層級列不計）"); C.cell(r, 2, f'=IF(E{rv}="對帳中",1,0)'); cnt = r; r += 1
    C.cell(r, 1, "說明：J 節為本公司唯一的 C0 判定，結論「無可比實績」：層級只透過 NVIDIA 顆數影響 FY2028 起的資料中心權利金，Arm 不揭露同口徑金額。"
                 "產業部署 GW 為層與層之間的差距（H 節），本模型並列兩讀數目標價，不裁決，也不計入對帳中項目數。")
    note = r; r += 1
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, G0=G0, G1=G1, L0=L0, L1=L1, H_PL=H_PL, J=j_title, Jn=j_note, J0=J0, J1=J1, rv=rv, rm=rm, rs=rs, cnt=cnt, note=note, IN=IN)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）")
    pmed, ttm = vb("同業前瞻 P/E 中位數"), vb("現價 ÷ 近 12 月 non-GAAP EPS")
    rb = lab(P, "情境 B：")
    pr = lambda lbl: lab(P, lbl, rb)
    dc_b, ro_b, lo_b, rev_b, dsh_b = pr("權利金：資料中心"), pr("權利金營收（Royalty）"), pr("授權及其他營收（License and other）"), lab(P, "營收", lab(P, "C. 授權及其他", rb), exact=True), pr("資料中心 ÷ 權利金")
    rv27, reps27, rflip = lab(RC, "FY2027 營收 vs 共識"), lab(RC, "FY2027 non-GAAP EPS vs 共識"), lab(RC, "評等翻轉（加權目標價＝現價 × 0.5）所需的 Compar")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    KR = SEN["KR"]; fr0 = SEN["first"] + 1   # 估值框架第二列＝DCF 權重 100%／0%
    SB = lambda k, col: f"Sens!{col}{KR[k]}"
    w30 = f"((0.3*{dcf}+0.7*{cadj})/px-1)"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；v{VER} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元、同業比較法 "), f'TEXT({cadj},"0")', T(" 美元，都遠低於現價 "), 'TEXT(px,"0")',
        T(" 美元；兩者各半加權 "), f'TEXT({tp},"0")', T(" 美元，評等「"), rec, T("」。現價約為近一年 non-GAAP EPS 的 "), f'TEXT({ttm},"0")', T(" 倍，同業前瞻本益比中位數約 "), f'TEXT({pmed},"0")', T(" 倍。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：現金流折現與同業比較法各占一半，同業倍數取六家同業中位數（上市未滿十年，不做相對倍數調整）。現金流折現權重降到三成時評等為"),
        REC5(w30), T("（"), f'TEXT({w30},"+0.0%;-0.0%")', T("），只看同業比較法時為"), REC5(f"({cadj}/px-1)"), T("（"), f'TEXT({cadj}/px-1,"+0.0%;-0.0%")',
        T("）；要翻為「放空」，同業比較法採用的本益比須升到 "), f'TEXT(Recon!B{rflip},"0")', T(" 倍。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("，扣股票薪酬），同業比較法 "), f'TEXT({cadj},"0.00")',
        T(" 美元＝同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 一年後前瞻 EPS。"))
    put("o2", T("現價隱含的條件：要讓現金流折現法支持現價，"), f'IF(ABS({rv("w_res")})<0.01,"折現率須降到 "&TEXT({rv("w1")},"0.0%"),"折現率即使降到 "&TEXT({rv("w1")},"0.0%")&" 仍不足")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.0")', T(" 倍；同業比較法則需 "), f'TEXT({rv("pe1")},"0")', T(" 倍本益比，是同業中位數的 "), f'TEXT({rv("rc1")},"0.0")', T(" 倍。"))
    put("o3", T("與共識的差距在倍數：FY2027 EPS "), f'TEXT(Recon!E{reps27},"0.00")', T(" 美元，比共識 "), f'TEXT(Recon!B{reps27},"0.00")', T(" 低 "), f'TEXT(-Recon!F{reps27},"0.0%")',
        T("，營收差 "), f'TEXT(Recon!F{rv27},"+0.0%;-0.0%")', T("；共識目標價 "), 'TEXT(cons_tp,"0")', T(" 美元隱含約 "), 'TEXT(cons_tp/cons_feps,"0")', T(" 倍前瞻本益比，本模型用同業中位數。"))
    put("o4", T("資料中心是成長主力但規模仍小：公司口述上半年資料中心權利金年增約一倍，本模型 FY2027 假設 "), f'TEXT(PL!E{dc_b}/PL!D{dc_b}-1,"0%")',
        T("；基準情境 FY2029 資料中心權利金 "), f'TEXT(PL!G{dc_b},"0.00")', T(" 十億美元，占權利金 "), f'TEXT(PL!G{dsh_b},"0%")', T("；層級只影響其中 NVIDIA 的部分。"))
    put("o5", T("影響最大的是估值框架與尾端事件：只看現金流折現時加權目標價 "), f'TEXT(Sens!B{fr0},"+0%;-0%")', T("，只看同業比較法 "), f'TEXT(Sens!C{fr0},"+0%;-0%")',
        T("；營運輸入中軟銀授權不續約 "), f'TEXT({SB("sb_lic", "B")},"+0%;-0%")', T("、高通停付權利金兩年 "), f'TEXT({SB("qc_stop", "B")},"+0%;-0%")',
        T("、授權成長放緩 "), f'TEXT({SB("g_lic", "B")},"+0%;-0%")', T("，都翻不過評等。"))
    put("drv", T("分部營收推導（情境 B）：權利金分四個市場，各為出貨 × 每顆權利金率：智慧型手機（出貨持平、費率年增）、資料中心（FY2027 錨點 × 處理器顆數指數 × 權利金率指數；NVIDIA 部分讀 Siliconomics 顆數）、車用、物聯網與其他；"
                 "授權及其他＝授權（年增）＋Arm 自研晶片。FY2028：權利金 "), f'TEXT(PL!F{ro_b},"0.00")', T("（其中資料中心 "), f'TEXT(PL!F{dc_b},"0.00")',
        T("）、授權及其他 "), f'TEXT(PL!F{lo_b},"0.00")', T("，合計 "), f'TEXT(PL!F{rev_b},"0.00")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型的 NVIDIA 資料中心 GPU 顆數、情境係數與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1'] + 2}", T("（Coord 頁 F 節）。C0（J 節）：層級只驅動 FY2028 起資料中心權利金中 NVIDIA 的部分，"
        "Arm 不揭露資料中心權利金金額（只口述上半年年增約一倍），結論「"), f"Coord!E{CO['rv']}", T("」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")',
        T(" 項；參考：層級隱含 NVIDIA 顆數年增 "), f'TEXT(Coord!C{CO["rm"]},"0%")', T("，口徑不同不判定。產業部署 GW 兩層差距另列參考；對帳中年份 CY2027–CY2029（FY2028 起）依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（工作單 20261114 第二批第 2 節）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；Hist 頁新增損益表歷史節、"
                  "PL 頁末標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B，Non-GAAP）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準（低案、高案並列，引用主表整本重算快照）、"
                  "Coord 頁（兩讀數目標價、讀取層級、C0 實績：資料中心權利金無同口徑揭露，結論無可比實績）；既有數值格不變（新增列、新增頁、版本字串與 build_id 以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    """README A1 標題段版號與 build_id 前綴改本版，末尾加本版說明；A2 版本列前加本版；Sens 底部 build_id 格同步。"""
    RD, SE = wb["README"], wb["Sens"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and f"個股模型 v{OLD}（" in a1.split("；", 1)[0] and f"build_id arm-v{OLD}-" in a1, a1[:120]
    RD["A1"].value = (a1.replace(f"個股模型 v{OLD}（", f"個股模型 v{VER}（", 1).replace(f"build_id arm-v{OLD}-", f"build_id arm-v{VER}-", 1)
                      + f"；v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（Coord 頁兩讀數目標價、讀取層級、C0 實績；Hist 損益表歷史、RevDCF、PL 標準損益表、Val 0.7、Sens 加權目標價相對基準），既有數值不變")
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{OLD}（"), a2[:40]
    RD["A2"].value = f"版本：v{VER}（{BUILT}）：v1.3 版面與三層協調（既有數值不變）；" + a2[len("版本："):]
    rb = lab(SE, "build_id（README、Sens、HTML")
    bid = str(SE.cell(rb, 4).value)
    assert bid.startswith(f"arm-v{OLD}-"), bid
    SE.cell(rb, 4).value = bid.replace(f"arm-v{OLD}-", f"arm-v{VER}-", 1)


def apply(wb, here, SIL, CR):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("ARM_OVERRIDE"):
        return "v035_post：略過（alt_run／sens_arm 建置）"
    assert CR, "Coord 頁不存在：先跑 python3 tools/coord/alt_run.py arm"
    HR = hist_pl(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb)
    CO = coord(wb, SIL)
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    version(wb)
    return (f"v035_post：Hist 損益表歷史第 {HR['start']} 列起；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；"
            f"Coord F–J 節（C0 判定第 {CO['rv']} 列）；README v1.3 白話列 {len(RMA)} 列")
