# AMD v0.4.7（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0 公司實績對帳、硬體股兩讀數）。
# 依據：docs/workorders/20261114_第二批13家_v1.3與三層協調.md、docs/workorders/20261114_C0實績對帳補列_七家.md、
#       docs/plan/三層基準協調規則.md v1.4.2（C0 第 2 點、C4）、docs/plan/個股HTML版面規格.md v1.3、tools/coord/README.md。
# 樣板：Broadcom v0.5.6／v0.5.7（avgo/v056_post.py、v057_post.py；同一建檔家族）、NVIDIA v0.3.10。只新增，不改既有數值格：
#   Hist 頁末（新增一節）：損益表歷史（FY2023–FY2024 non-GAAP，AMD 8-K 新聞稿；FY2025 連 Q 頁）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；Non-GAAP＝模型口徑）、比率、GAAP EPS 對照、核對列
#   RevDCF 頁（新）：反向 DCF；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（業務主因＝既有整本重算快照表，估值框架＝Val 頁封閉式，折現參數＝既有表）
#   Coord 頁末 F–J 節：讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、損益表年度標示、C0（實績）
#   README 頁末：v1.3 展示層白話列；A1 版號與本版說明；build_id 前綴
# build_amd.py 在 v045_post 與 COORD.add_sheet 之後、存檔前呼叫 apply(wb, here, SIL)。COORD_RUN（alt_run.py）或 AMD_OVERRIDE（sens_amd.py）時不執行。
import os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

VER, OLD, BUILT = "0.4.7", "0.4.6", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
E8K = "https://www.sec.gov/Archives/edgar/data/2488/"
SRC = {"FY2024": E8K + "000000248825000009/q42024991.htm（AMD 8-K 2025-02-04：Q4 與 FY2024 新聞稿；FY2023 為比較欄）",
       "FY2023": E8K + "000000248824000009/q42023991final.htm（AMD 8-K 2024-01-30：Q4 與 FY2023 新聞稿；FY2022 為比較欄，只用營收與 non-GAAP 淨利作年增率基期）"}
SRC_Q = {"Q1": E8K + "000000248826000072/q12026991.htm", "Q2": E8K + "000000248826000121/q22026991.htm", "Q3": E8K + "000000248826000188/q32026991.htm"}


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


# ---------- Hist：損益表歷史（新增一節於既有 Hist 頁末） ----------
# (鍵, 項目, FY2023, FY2024, FY2025（連 Q 頁）, FY2022（基期）, 標記)；十億美元、EPS 美元、股數十億股
HIS = [
    ("rev", "營收（GAAP＝non-GAAP）", 22.680, 25.785, "=Q!F4", 23.601, "Verified"),
    ("gp", "毛利（non-GAAP）", 11.436, 13.759, "=Q!F12", None, "Verified"),
    ("opx", "營業費用（non-GAAP，扣授權收益；FY2023 6.616 − 0.034、FY2024 7.669 − 0.048）", 6.582, 7.621, "=Q!F14", None, "Derived"),
    ("oi", "營業利益（non-GAAP；公司口徑含授權收益）", 4.854, 6.138, "=Q!F15", None, "Verified"),
    ("pti", "稅前利益（non-GAAP）＝淨利＋所得稅", "=C{ni}+C{tax}", "=D{ni}+D{tax}", "=Q!F21", None, "Derived"),
    ("tax", "所得稅（non-GAAP）＝GAAP 所得稅 ＋ 調節表所得稅調整（FY2023 −0.346＋0.988；FY2024 0.381＋0.429）", 0.642, 0.810, "=Q!F21-Q!F22", None, "Derived"),
    ("ni", "稅後淨利（non-GAAP）", 4.302, 5.420, "=Q!F22", 5.504, "Verified"),
    ("sh", "稀釋股數（十億股）", 1.625, 1.637, "=Q!F24", None, "Verified"),
    ("eps", "EPS（non-GAAP 稀釋，公司公告；美元）", 2.65, 3.31, "=Q!F26", None, "Verified"),
    ("g_eps", "EPS（GAAP 稀釋，公司公告；美元）", 0.53, 1.00, "=Q!F28", None, "Verified"),
    ("g_ni", "稅後淨利（GAAP）", 0.854, 1.641, None, None, "Verified"),
]


def hist_pl(wb):
    H = wb["Hist"]
    r = H.max_row + 3
    H.cell(r, 1, f"損益表歷史（v{VER} 起；PL 頁「標準損益表」的歷史三年與年增率基期用）。十億美元；EPS 美元；股數十億股。non-GAAP 為公司口徑"
                 "（排除股票薪酬、併購無形資產攤銷、收購相關費用、重組、投資損益與認股權證公允價值變動）；FY2025 欄連 Q 頁（與公司全年值一致）").font = BOLD
    r0 = r; r += 1
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2022（只供年增率基期）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v23, v24, v25, v22, tag in HIS:
        H.cell(r, 1, f"hp_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25), ("F", "22", v22)):
            if v is None:
                continue
            if isinstance(v, str) and "{" in v:
                v = v.format(**rows)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
            dn(wb, f"hp_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, tag)
        H.cell(r, 8, f"FY2023：{SRC['FY2023']}；FY2024：{SRC['FY2024']}" + ("；FY2025：Q 頁 F 欄（AMD 8-K 2026-02-03，Q4 與 FY2025 新聞稿）" if v25 is not None else ""))
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    for labl, f in (("毛利 − 營業費用 − 營業利益（non-GAAP）", "={c}{gp}-{c}{opx}-{c}{oi}"),
                    ("稅前利益 − 所得稅 − 稅後淨利（non-GAAP）", "={c}{pti}-{c}{tax}-{c}{ni}"),
                    ("稅後淨利 ÷ 稀釋股數 − 公司公告 EPS（四捨五入差）", "={c}{ni}/{c}{sh}-{c}{eps}")):
        H.cell(r, 1, labl)
        for c in "CDE":
            H[f"{c}{r}"] = f.format(c=c, **rows); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 1, "備註：AMD non-GAAP 營業利益含授權收益（FY2023 0.034、FY2024 0.048，公司損益表 Licensing gain），本表併入營業費用（扣除），使毛利 − 營業費用＝營業利益；"
                 "non-GAAP 所得稅＝GAAP 所得稅＋調節表「Income tax provision」調整（AMD 以 13% 正規化稅率計算）。").font = Font(italic=True)
    return dict(start=r0, rows=rows)


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"   # 本表欄；FY2023A–FY2025A＝Hist；F–I＝PL 情境 B 區塊 D–G（FY2026 錨點、FY2027–FY2029）；J–L＝PL 延伸表情境 B 的 H–J（FY2030–FY2032）


def pl_standard(wb):
    P = wb["PL"]
    rb = lab(P, "情境 B：")
    rc = lab(P, "B. 報告值", rb)
    g = lambda lbl, s=rc: lab(P, lbl, s, exact=True)
    R = dict(rev=g("營收"), gp=g("毛利（non-GAAP）"), opx=g("營業費用"), oi=g("營業利益"), oth=g("非營業淨額（non-GAAP）"), ni=g("淨利（non-GAAP）"),
             sh=g("稀釋股數"), eps=g("EPS（non-GAAP）"))
    x0 = lab(P, "三情境延伸到 FY2032"); xb = lab(P, "情境 B", x0, exact=True)
    X = dict(rev=lab(P, "營收", xb, exact=True), ni=lab(P, "淨利（延伸表）", xb, exact=True), sh=lab(P, "稀釋股數", xb, exact=True), eps=lab(P, "EPS（延伸表）", xb, exact=True))
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026E 錨點、FY2027E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "只引用本頁上方與 Hist、Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：Non-GAAP（本模型口徑）。AMD 的財測、指引與市場共識 EPS 皆以 non-GAAP 為主（排除股票薪酬、併購無形資產攤銷、收購相關費用、投資損益與認股權證公允價值變動），故本表用 Non-GAAP；"
                 "其他收入（淨額）＝利息費用與其他收益淨額（non-GAAP）。GAAP EPS 列在本表下方對照（只列實績）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點（第一至第三季實績＋第四季指引）；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2032E＝延伸期（收斂法只推營收與淨利，營業利益＝淨利 ÷（1 − 稅率）− 其他收入，同 DCF 頁）。"
                 "延伸期毛利率與營業費用率沿用 FY2029E，據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）。EPS＝稀釋（含客戶認股權證）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
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

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFG"]
    ex = lambda key: [f"=PL!{c}{X[key]}" for c in "HIJ"]
    hh = lambda k: [f"=hp_{k}_23", f"=hp_{k}_24", f"=hp_{k}_25"]
    row("rev", "營收", "US$B", hh("rev") + fc("rev") + ex("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/hp_rev_22-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "US$B", [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "CDEFGHI"] + [f"={c}{L['rev']}*(1-$I{r_c + 2})" for c in "JKL"], M,
        "FY2023A–FY2029E＝營收 − 毛利；FY2030E–FY2032E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "US$B", hh("gp") + fc("gp") + [f"={c}{L['rev']}-{c}{L['cogs']}" for c in "JKL"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E–FY2032E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "US$B", hh("opx") + fc("opx") + [f"={c}{L['rev']}*$I{r_c + 3}/$I{L['rev']}" for c in "JKL"], M,
        "歷史扣授權收益（Hist 頁）；FY2030E–FY2032E＝營收 × FY2029E 營業費用率（non-GAAP，不含股票薪酬）")
    r_ni = r_c + 9
    row("oi", "營業利益", "US$B", hh("oi") + fc("oi") + [f"={c}{r_ni}/(1-tx_f)-oth_f" for c in "JKL"], M, "FY2030E–FY2032E 由淨利反推（同 DCF 頁；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）＝利息費用與其他收益淨額", "US$B", ["=hp_pti_23-hp_oi_23", "=hp_pti_24-hp_oi_24", "=hp_pti_25-hp_oi_25"]
        + [f"=PL!{c}{R['oth']}" for c in "DEFG"] + ["=oth_f"] * 3, M, "non-GAAP；歷史＝稅前利益 − 營業利益")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    row("tax", "所得稅", "US$B", ["=hp_tax_23", "=hp_tax_24", "=hp_tax_25"] + [f"={c}{L['pti']}-{c}{r_ni}" for c in "FGHIJKL"], M,
        "歷史＝non-GAAP 所得稅（Hist 頁）；預測＝稅前 − 稅後淨利（non-GAAP 稅率 13%）")
    assert L["tax"] == r_c + 8 and r == r_ni
    row("ni", "稅後淨利", "US$B", hh("ni") + fc("ni") + ex("ni"), M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數（含客戶認股權證）", "十億股", hh("sh") + fc("sh") + ex("sh"), M)
    row("eps", "EPS（non-GAAP 稀釋）", "US$", ["=hp_eps_23", "=hp_eps_24", "=hp_eps_25"] + fc("eps") + ex("eps"), "0.00",
        "FY2023A–FY2025A＝公司公告 non-GAAP 稀釋 EPS；預測＝淨利 ÷ 稀釋股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/hp_ni_22-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 non-GAAP 淨利（Hist 頁）")
    P.cell(r, 1, "GAAP 對照（歷史；預測不估 GAAP，故只列實績）").font = BOLD; r_gaap = r; r += 1
    row("g_ni", "GAAP：稅後淨利", "US$B", ["=hp_g_ni_23", "=hp_g_ni_24", None], M, "FY2025 GAAP 淨利未列入 Hist（只列 EPS）")
    row("g_eps", "GAAP：EPS（稀釋，公司公告）", "US$", ["=hp_g_eps_23", "=hp_g_eps_24", "=hp_g_eps_25"], "0.00", "含股票薪酬、併購無形資產攤銷、收購相關費用與投資損益")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None] * 3, "0.000")
    row("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "US$B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JKL"], "0.000")
    row("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JKL"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 摘要／延伸表（FY2026E–FY2032E）；FY2025A − PL 情境 B 區塊", "US$B",
        [None, None, f"=E{L['ni']}-PL!C{R['ni']}"] + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"={c}{L['ni']}-PL!{s}{X['ni']}" for c, s in zip("JKL", "HIJ")], "0.000")
    row("c5", "核對：EPS − 情境 B 摘要／延伸表（FY2026E–FY2032E）", "US$", [None] * 3 + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip("FGHI", "DEFG")]
        + [f"={c}{L['eps']}-PL!{s}{X['eps']}" for c, s in zip("JKL", "HIJ")], "0.000")
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
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。明確期第一年 FY2027 維持模型值不動，只動 FY2028–FY2032（含終值）。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
                     T(" 倍（同業中位數 × AMD 過去十年相對同業的溢價 "), f'TEXT({coef},"0.00")', T(" 倍，已達夾值上限）。換言之，市場目前給 AMD 的倍數是同業的 "),
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
def sens_tor(wb):
    SE, V = wb["Sens"], wb["Val"]
    r_tp = lab(V, "目標價（DCF 與 Comparable 加權）")
    tp, dcf, cadj, cu = f"Val!$B${r_tp}", f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}", f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後")
    r_eps = lab(V, "各情境 12 個月後的前瞻 EPS")
    pmed, coef, raw = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}", f"Val!$B${lab(V, '原始比值（本公司 ÷ 同業）')}"
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    par = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        par.append(rr); rr += 1
    r_bz = lab(SE, "目標價敏感度：業務主因"); bh = r_bz + 2
    assert SE.cell(bh, 1).value == "案例", SE.cell(bh, 1).value
    biz = []; rr = bh + 1
    while SE.cell(rr, 1).value:
        biz.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；v{VER} 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝引用本頁「目標價敏感度：業務主因」整本重算快照表（amd/sens_amd.py，三情境同時套用）；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用本頁「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    hdr_r = r; r += 1
    first = r
    frame = [
        ("同業比較法的相對倍數係數 不調整 1.00／不夾值（基準＝夾值上限 2.00）",
         f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*1", f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}*{raw}", "Val 頁：Comparable（未調整）× 係數；低＝不調整、高＝原始比值不夾值（與業務主因表的不夾值案例相同）"),
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$D${r_eps}", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$B${r_eps}", "Val 頁：單一情境的 DCF 價值與前瞻 EPS"),
        ("DCF 在加權中的權重 100%（只看現金流折現）／0%（只看同業比較法）（基準＝50%）",
         f"={dcf}", f"={cadj}", "Val 頁：DCF 與 Comparable（調整後）目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    for sr in biz:
        if SE.cell(sr, 8).value == "rmhi":   # 已由框架案例「不夾值」涵蓋
            continue
        SE.cell(r, 1, str(SE.cell(sr, 1).value))
        for col, tcol, scol in ((2, 4, "C"), (3, 5, "E")):
            v = SE[f"{scol}{sr}"].value
            if v is None or (isinstance(v, str) and not v.startswith("=")):
                continue
            SE.cell(r, tcol, f"=Sens!{scol}{sr}").number_format = "#,##0.00"
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「業務主因」整本重算快照表；三情境同時套用）"); r += 1
    for sr in par:
        t = str(SE.cell(sr, 1).value)
        if t.startswith("DCF 權重"):
            continue
        SE.cell(r, 1, t)
        for col, tcol, scol in ((2, 4, "C"), (3, 5, "E")):
            v = SE[f"{scol}{sr}"].value
            if v is None or (isinstance(v, str) and not v.startswith("=")):
                continue
            SE.cell(r, tcol, f"=Sens!{scol}{sr}").number_format = "#,##0.00"
            SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{sr}（本頁「目標價參數敏感度」表）"); r += 1
    last = r - 1
    st = lab(SE, "模型狀態（CURRENT")
    SE.cell(r, 1, "快照狀態（營運輸入案例）"); SE.cell(r, 2, f"=D{st}")
    SE.cell(r, 3, "同本頁「模型狀態」（輸入指紋與加權目標價皆與建置快照一致＝CURRENT；改輸入後須重跑 python3 amd/sens_amd.py）"); r += 1
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last, state=r - 1)


# ---------- Coord：讀取層級、讀取層級核對、層級之間、損益表年度標示、C0（實績） ----------
def coord(wb, SIL):
    C, I, Q, P = wb["Coord"], wb["In"], wb["Q"], wb["PL"]
    N = SIL.NAMES
    r = C.max_row + 2
    ys = (2026, 2027, 2028, 2029)
    r_flag = lab(C, "對帳中旗標")
    src_in = lambda key: str(I.cell(lab(I, key, col=1, exact=True), 8).value or "")
    vtxt = f"{SIL.VERSION}（{SIL.FILE}）"
    C.cell(r, 1, "F. 讀取層級（工作單 20261114 第二批第 2 節）：本模型讀哪一層、哪一版、哪些名稱。自層級讀入的同源值只列在此節與 G 節，不進 C0 判定").font = BOLD; r += 1
    for j, h in enumerate(["項目", "層", "版本", "本模型用在哪裡", "值／範圍"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1 = r
    rows = [
        ("加速器顆數：各平台 2026–2028（u_*）與 2029 成長（g29_*）；AMD Instinct 為 u_amd_*", "Siliconomics（經市場層 markets/memory UNITS 與台積電成品轉入）", vtxt,
         "Units 頁：加速器市場金額與 AMD 金額份額 → 資料中心 GPU 營收", "In 頁 A 節；Siliconomics SIF_u_amd_26 等"),
        ("情境顆數係數 um28／um29（情境 B 0.80／0.710）", "Siliconomics（經台積電成品）", vtxt, "情境 B、C 顆數＝原始顆數 × 係數", "In 頁 A 節"),
        ("各平台每顆售價 pm_*、pa_*（AMD Instinct pa_amd）", "Siliconomics 行業變數（經 NVIDIA 成品 In 頁）", src_in("pa_amd").split("｜")[0], "Units 頁金額份額；k_rec 校準分母", "In 頁 B 節"),
        ("HBM 混合單價路徑 hbm_<情境>_<年>", "市場層（經 NVIDIA 成品）", src_in("hbm_B_27").split("｜")[0], "HBM 漲價轉嫁與 GPU 成本", "In 頁 C 節"),
        ("對帳中旗標與另一側讀數比值（SIF_RecoFlag、SIF_DeployGW_AltRatio_<年>）", "Siliconomics", vtxt, "Coord 頁 A–E 節：兩讀數目標價（tools/coord）", "CY2027–CY2029＝對帳中"),
        ("Wattonomics 錨定係數 wta_27–wta_29（舊敏感度案例；SIF_WTAnchor_Base）", "Wattonomics（經 Siliconomics 轉入）", vtxt, "Sens 頁「Wattonomics 錨定」案例", "舊口徑，保留供對照；另一側讀數取代其角色"),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。AMD 的 GPU 認列係數 k_rec 是本模型自己的校準（錨點 ÷ 層級值），不寫回層級；層級的 AMD 顆數與每顆售價不由本模型校準"
                 "（Siliconomics 快照註記：顆數 Morgan Stanley 等、售價 HSBC 等外部來源）。").font = Font(italic=True); r += 2
    # ---- G：讀取層級核對（同源，不判定） ----
    C.cell(r, 1, "G. 讀取層級核對（同源，不判定）：個股讀入的層級值 vs 層級值，必然相同；只證明本模型讀到現行版，不計入對帳中").font = BOLD; r += 1
    for j, h in enumerate(["項目", "本模型值", "層級對照值", "差距", "狀態", "單位與口徑", "對照來源（層、版本）", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    G0 = r
    for k, labl, unit, fmt in (("u_amd_26", "AMD Instinct 顆數 CY2026", "百萬顆", "0.000"), ("u_amd_27", "AMD Instinct 顆數 CY2027", "百萬顆", "0.000"),
                               ("u_amd_28", "AMD Instinct 顆數 CY2028（原始，未乘情境係數）", "百萬顆", "0.000"), ("pa_amd", "AMD Instinct 每顆售價", "US$K", "0.0")):
        C.cell(r, 1, labl); C.cell(r, 2, f"={k}").number_format = fmt; C.cell(r, 3, SIL.base(k)).number_format = fmt; C.cell(r, 3).font = BLUE
        C.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"; C.cell(r, 5, "讀取層級（同源，不判定）")
        C.cell(r, 6, unit); C.cell(r, 7, f"Siliconomics {SIL.VERSION}（SIF_{k}）"); C.cell(r, 8, SIL.V[k]["tag"] if k in SIL.V else "")
        C.cell(r, 9, "本模型直接讀層級值；恆等，差距應為 0"); r += 1
    G1 = r - 1
    r += 1
    C.cell(r, 1, "H. 層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）：產業部署 GW，Siliconomics 現行基準 vs 另一側讀數（Wattonomics 換算）。這是兩層之間的差距，不是 AMD 自己的量；本模型依 C4 並列兩讀數目標價").font = BOLD; r += 1
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
    C.cell(r, 1, "I. 對帳中年份與損益表年度標示（C4；AMD 會計年度＝日曆年 FY(t)＝CY(t)；FY2030E–FY2032E 為 FY2029E 之後的延伸期，依賴 CY2029）").font = BOLD; r += 1
    for j, h in enumerate(["項目"] + [f"FY{y}E" for y in range(2026, 2033)], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    C.cell(r, 1, "依賴的日曆年")
    dep = [2026, 2027, 2028, 2029, 2029, 2029, 2029]
    for j, y in enumerate(dep):
        C.cell(r, 2 + j, y)
    r += 1
    H_PL = r
    C.cell(r, 1, "損益表年度標示（FY2026E–FY2032E）")
    for j, fy in enumerate(range(2026, 2033)):
        C.cell(r, 2 + j, f'=IF(${"BCDE"[dep[j] - 2026]}${r_flag}=1,"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    # ---- J：C0（實績） ----
    C.cell(r, 1, "J. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）").font = BOLD; j_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝AMD 8-K 已公布的資料中心分部營收（公司揭露，Verified）。層級側＝Siliconomics 現行版隱含的 AMD 資料中心 GPU（Instinct）營收＝CY2026 顆數 × 每顆售價"
                 "（In 頁 AMD Instinct 顆數 × 每顆售價，皆讀自層級、非本模型校準）。口徑：資料中心分部＝Instinct GPU＋EPYC 伺服器 CPU＋DPU、FPGA 等；AMD 不揭露 Instinct GPU 營收，"
                 "Siliconomics 也沒有資料中心 CPU 層，故兩側沒有同口徑量：C0 結論為「無可比實績」。下列兩組只作參考、不判定：(1) 上限檢驗——層級 GPU 隱含應不超過資料中心分部實績；"
                 "(2) 本模型的 GPU 拆分（Assumed，公司不揭露）對層級隱含。時差讀法照工作單：主讀數＝2026 年已公布期間（Q1–Q3＝2026-01～09，9 個月）累計 vs 層級 CY2026 × 9 ÷ 12；"
                 "第二讀數＝最近一季（Q3）× 4 vs 層級 CY2026，約等於 Q3 中點（2026-08）的年化速率，晚於年中，在成長年偏高。")
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
    rdc, rgp = lab(Q, "Data Center", 3), lab(Q, "其中：資料中心 GPU", 3)
    assert [Q.cell(3, c).value for c in (7, 8, 9)] == ["Q1 2026", "Q2 2026", "Q3 2026"]
    inp("q1", "資料中心分部營收：Q1 2026（公司 8-K）", f"=Q!G{rdc}", "US$B", "2026-01～2026-03（Q1，季度結束 2026-03-28）", f"Q 頁 G{rdc}；AMD 8-K 2026-05-05 新聞稿（Q1 2026）", "Verified", GREEN, url=SRC_Q["Q1"])
    inp("q2", "資料中心分部營收：Q2 2026（公司 8-K）", f"=Q!H{rdc}", "US$B", "2026-04～2026-06（Q2，季度結束 2026-06-27）", f"Q 頁 H{rdc}；AMD 8-K 2026-08-04 新聞稿（Q2 2026）", "Verified", GREEN, url=SRC_Q["Q2"])
    inp("q3", "資料中心分部營收：Q3 2026（公司 8-K）", f"=Q!I{rdc}", "US$B", "2026-07～2026-09（Q3，季度結束 2026-09-26）", f"Q 頁 I{rdc}；AMD 8-K 2026-11-03 新聞稿（Q3 2026）", "Verified", GREEN, url=SRC_Q["Q3"])
    inp("mo", "已公布月數（2026 Q1–Q3）", 9, "月", "2026-01～2026-09", "AMD 會計年度＝日曆年（12 月最後一個週六結束）", "Derived", BLUE, fmt="0")
    inp("u", "層級：AMD Instinct 顆數 CY2026", "=u_amd_26", "百萬顆", "CY2026", f"In 頁 AMD Instinct 顆數 2026（Siliconomics {SIL.VERSION}；Morgan Stanley 等）", "Interested-party／Assumed（層）", GREEN, url="In!u_amd_26；SIF_u_amd_26")
    inp("p", "層級：AMD Instinct 每顆售價", "=pa_amd", "US$K", "2026", f"In 頁 每顆售價：AMD Instinct（Siliconomics {SIL.VERSION}；HSBC 等）", "Interested-party（層）", GREEN, fmt="0.0", url="In!pa_amd；SIF_pa_amd")
    inp("L", "層級隱含：AMD 資料中心 GPU 營收 CY2026（顆數 × 售價）", f"=B{IN['u']}*B{IN['p']}", "US$B", "CY2026", "上兩列相乘（百萬顆 × 千美元＝十億美元）", "Derived")
    inp("chk", "核對：層級隱含 − Units 頁 AMD 金額（PL 情境 B 區塊 FY2026 市場金額 × AMD 份額；應為 0）", f"=B{IN['L']}-PL!D{lab(P, '加速器市場金額', lab(P, '情境 B：'))}*PL!D{lab(P, 'AMD 金額份額', lab(P, '情境 B：'))}", "US$B", "CY2026", "PL 情境 B 區塊（Units 頁）", "Derived", fmt="0.000000")
    inp("g1", "參考：本模型 GPU 拆分 Q1 2026（Assumed）", f"=Q!G{rgp}", "US$B", "Q1 2026", f"Q 頁 G{rgp}（資料中心 × 假設的 GPU 占比；公司不揭露）", "Assumed", GREEN, url="In!gs_q1")
    inp("g2", "參考：本模型 GPU 拆分 Q2 2026（Assumed）", f"=Q!H{rgp}", "US$B", "Q2 2026", f"Q 頁 H{rgp}（資料中心 × 假設的 GPU 占比）", "Assumed", GREEN, url="In!gs_q2")
    inp("g3", "參考：本模型 GPU 拆分 Q3 2026（Assumed）", f"=Q!I{rgp}", "US$B", "Q3 2026", f"Q 頁 I{rgp}（資料中心 − 非 GPU；非 GPU 季增為假設）", "Assumed", GREEN, url="In!ng_q3")
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    J0 = r
    b = IN; pct = "+0.0%;-0.0%"
    C.cell(r, 1, "C0（實績）判定：AMD 資料中心 GPU（Instinct）營收 CY2026").font = BOLD
    C.cell(r, 5, "無可比實績").font = BOLD
    C.cell(r, 6, "缺：AMD Instinct GPU 營收揭露（公司只揭露資料中心分部）；Siliconomics 無資料中心 CPU 層，無法組成與分部同口徑的層級值")
    C.cell(r, 7, f"公司：AMD 8-K（Q 頁第 {rdc} 列）；層級：Siliconomics {SIL.VERSION}（AMD 顆數 × 每顆售價）")
    C.cell(r, 8, "—")
    C.cell(r, 9, "不用模型拆分（Assumed）或同源值補位；公司日後揭露 Instinct 營收（或年報分拆）時改以 GPU 營收實績判定")
    rv = r; r += 1
    C.cell(r, 1, "參考（上限檢驗，不判定）：主讀數——資料中心分部 Q1–Q3 累計實績 vs 層級 GPU 隱含 × 9 ÷ 12")
    C.cell(r, 2, f"=B{b['q1']}+B{b['q2']}+B{b['q3']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(C{r}<=B{r},"參考：層級 GPU 隱含未超過分部實績（上限檢驗通過）","參考：層級 GPU 隱含超過分部實績（上限檢驗不通過）")')
    C.cell(r, 6, "US$B；公司實績＝分部營收（含 CPU 等）；層級＝GPU 一項 CY2026 × 9 ÷ 12；差距＝實績 ÷ 層級 − 1（分部大於 GPU 為正常）")
    C.cell(r, 7, "公司：AMD 8-K；層級：Siliconomics " + SIL.VERSION); C.cell(r, 8, "Verified（公司）／Interested-party（層）／Derived（線性時差）")
    C.cell(r, 9, f'="分部實績減層級 GPU 隱含＝"&TEXT(B{r}-C{r},"0.00")&"（隱含的 9 個月非 GPU 營收，占分部 "&TEXT((B{r}-C{r})/B{r},"0%")&"）；本模型 Q1–Q3 的非 GPU（Assumed）為 "&TEXT(B{r}-(B{b["g1"]}+B{b["g2"]}+B{b["g3"]}),"0.00")&"。兩者差距大表示層級 GPU 售價或顆數偏高、或本模型 GPU 占比偏低；公司未揭露，無法判定"')
    rb1 = r; r += 1
    C.cell(r, 1, "參考（上限檢驗，不判定）：第二讀數——資料中心分部 Q3 × 4 vs 層級 GPU 隱含 CY2026")
    C.cell(r, 2, f"=B{b['q3']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(C{r}<=B{r},"參考：層級 GPU 隱含未超過分部實績（上限檢驗通過）","參考：層級 GPU 隱含超過分部實績（上限檢驗不通過）")')
    C.cell(r, 6, "US$B；Q3 單季 × 4（年化，約等於 2026-08 的速率）；層級＝GPU 一項 CY2026 全年")
    C.cell(r, 7, "公司：AMD 8-K（Q 頁 I 欄）；層級同上"); C.cell(r, 8, "Verified（公司）／Interested-party（層）／Derived（年化）")
    C.cell(r, 9, "Q3 中點晚於年中，成長年中年化值偏高；上限檢驗不受影響")
    rb2 = r; r += 1
    C.cell(r, 1, "參考（Assumed 拆分，不判定）：主讀數——本模型 GPU 拆分 Q1–Q3 累計 vs 層級 GPU 隱含 × 9 ÷ 12")
    C.cell(r, 2, f"=B{b['g1']}+B{b['g2']}+B{b['g3']}").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, '=IF(ABS(D' + str(r) + ')>0.15,"參考（Assumed 拆分）：超過 ±15%","參考（Assumed 拆分）：在 ±15% 內")')
    C.cell(r, 6, "US$B；本模型 GPU 拆分（GPU 占比為假設）；層級＝CY2026 × 9 ÷ 12")
    C.cell(r, 7, "本模型 Q 頁 GPU 列（Assumed）；層級同上"); C.cell(r, 8, "Assumed（拆分）／Interested-party（層）")
    C.cell(r, 9, "拆分是本模型假設、不是公司實績，不進判定（工作單：口述占比 × 營收屬 Assumed）。本模型的認列係數（FY2026 錨點 GPU ÷ 層級值）即此差距的全年版")
    ra1 = r; r += 1
    C.cell(r, 1, "參考（Assumed 拆分，不判定）：第二讀數——本模型 GPU 拆分 Q3 × 4 vs 層級 GPU 隱含 CY2026")
    C.cell(r, 2, f"=B{b['g3']}*4").number_format = "#,##0.000"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.000"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, '=IF(ABS(D' + str(r) + ')>0.15,"參考（Assumed 拆分）：超過 ±15%","參考（Assumed 拆分）：在 ±15% 內")')
    C.cell(r, 6, "US$B；Q3 拆分 × 4；層級＝CY2026 全年")
    C.cell(r, 7, "本模型 Q 頁 GPU 列（Assumed）；層級同上"); C.cell(r, 8, "Assumed（拆分）／Interested-party（層）")
    C.cell(r, 9, '=IF(AND(ABS(D' + str(ra1) + ')>0.15,ABS(D' + str(r) + ')>0.15,SIGN(D' + str(ra1) + ')=SIGN(D' + str(r) + ')),"若拆分屬實，兩讀數同向低於層級（層級 GPU 偏高）；拆分為假設，只列為對帳單候選，待公司揭露","兩讀數未同向超過 ±15%")')
    ra2 = r; r += 1
    J1 = r - 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（J 節 C0（實績）判定；G 節讀取層級與參考列不計）"); C.cell(r, 2, f'=IF(E{rv}="對帳中",1,0)'); cnt = r; r += 1
    C.cell(r, 1, "說明：J 節為本公司唯一的 C0 判定，結論「無可比實績」（缺 AMD Instinct GPU 營收揭露；Siliconomics 無資料中心 CPU 層）。上限檢驗與 Assumed 拆分兩組只列參考：上限檢驗兩個讀數都通過；"
                 "Assumed 拆分兩個讀數都低於層級，與本模型認列係數小於 1 一致，列為對帳單候選、不開改值對帳單。產業部署 GW 為層與層之間的差距（H 節），本模型並列兩讀數目標價，不裁決，也不計入對帳中項目數。")
    note = r; r += 1
    for col, w in zip("FGHI", (40, 60, 40, 70)):
        C.column_dimensions[col].width = max(C.column_dimensions[col].width or 0, w)
    return dict(F1=F1, G0=G0, G1=G1, L0=L0, L1=L1, H_PL=H_PL, J=j_title, Jn=j_note, J0=J0, J1=J1, rv=rv, rb1=rb1, rb2=rb2, ra1=ra1, ra2=ra2, cnt=cnt, note=note)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; P = wb["PL"]; RC = wb["Recon"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cu, cadj = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數")
    rb = lab(P, "情境 B："); rcB = lab(P, "B. 報告值", rb)
    pr = lambda lbl, s=rcB: lab(P, lbl, s, exact=True)
    rev_b, gpu_b, dcx_b, cg_b, emb_b = pr("營收"), pr("資料中心 GPU"), pr("資料中心非 GPU（CPU 等）"), pr("Client＋Gaming"), pr("Embedded")
    rdc27, reps27, rgm27 = lab(RC, "FY2027 資料中心營收 vs 公司"), lab(RC, "FY2027 EPS vs 共識（stockanalysis"), lab(RC, "FY2027 毛利率 vs FY2026 錨點")
    rkr = lab(RC, "k_rec（Derived）")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, SEN['first'])}"
    w30 = f"((0.3*{dcf}+0.7*{cadj})/px-1)"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；v{VER} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元遠低於現價 "), 'TEXT(px,"0")', T(" 美元；同業比較法 "), f'TEXT({cadj},"0")',
        T(" 美元已含 AMD 過去十年相對同業的本益比溢價（係數 "), f'TEXT({coef},"0.00")', T("，夾值上限）。兩者各半加權 "), f'TEXT({tp},"0")', T(" 美元，評等「"), rec, T("」。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：資料中心 GPU 營收由市場層顆數 × 售價推得，認列係數 "), f'TEXT(Recon!C{rkr},"0.00")',
        T("（FY2026 錨點 ÷ 層級值）沿用；同業比較法的溢價係數停在上限 2.0。現金流折現權重降到三成時評等為"), REC5(w30), T("（"), f'TEXT({w30},"+0.0%;-0.0%")', T("）；係數不夾值時為"),
        REC5(f"({SB('同業比較法的相對倍數係數', 'E')}/px-1)"), T("。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("），進加權的同業比較法 "), f'TEXT({cadj},"0.00")',
        T(" 美元＝同業前瞻本益比中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × 溢價係數 "), f'TEXT({coef},"0.00")', T(" 倍 × 前瞻 EPS。"))
    put("o2", T("現價隱含的條件：要讓現金流折現法支持現價，"), f'IF(ABS({rv("w_res")})<0.01,"折現率須降到 "&TEXT({rv("w1")},"0.0%"),"折現率即使降到 "&TEXT({rv("w1")},"0.0%")&" 仍不足")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0")', T("，模型 "), f'TEXT({rv("rv0")},"0")', T(" 十億美元）。"))
    put("o3", T("盈餘低於共識：FY2027 EPS 基準情境 "), f'TEXT(Recon!E{reps27},"0.00")', T(" 美元，比共識 "), f'TEXT(Recon!B{reps27},"0.00")', T(" 低 "), f'TEXT(-Recon!F{reps27},"0.0%")',
        T("；毛利率 "), f'TEXT(Recon!E{rgm27},"0.0%")', T("，比 FY2026 錨點低 "), f'TEXT(-Recon!F{rgm27}*100,"0.0")', T(" 個百分點，來自 GPU 占比上升與 HBM 漲價部分吸收（吸收比例為假設）。"))
    put("o4", T("核心問題：公司說資料中心營收 2027 年「逾一倍」（"), f'TEXT(Recon!B{rdc27},"0.0")', T(" 十億美元）；模型基準情境 "), f'TEXT(Recon!E{rdc27},"0.0")', T("（差 "),
        f'TEXT(Recon!F{rdc27},"+0.0%;-0.0%")', T("），樂觀情境 "), f'TEXT(Recon!C{rdc27},"0.0")', T("。差距來自市場層的 AMD 顆數與每顆售價，以及 CPU 成長假設。"))
    put("o5", T("框架假設的影響大於多數營運輸入：只看現金流折現時加權目標價 "), f'TEXT({SB("DCF 在加權中的權重", "D")},"0")', T(" 美元（"), f'TEXT({SB("DCF 在加權中的權重", "B")},"+0%;-0%")',
        T("），情境全押樂觀 "), f'TEXT({SB("情境權重", "C")},"+0%;-0%")', T("、全押保守 "), f'TEXT({SB("情境權重", "B")},"+0%;-0%")', T("。營運輸入中世代售價係數（高）"),
        f'TEXT({SB("世代售價係數", "C")},"+0.0%;-0.0%")', T("、資料中心 CPU 成長 "), f'TEXT({SB("資料中心 CPU 成長", "B")},"+0.0%;-0.0%")', T("／"), f'TEXT({SB("資料中心 CPU 成長", "C")},"+0.0%;-0.0%")', T(" 影響最大。"))
    put("drv", T("分部營收推導（情境 B）：資料中心 GPU＝加速器市場金額 × AMD 金額份額 × 認列係數 × 世代售價係數，再加 HBM 漲價中轉嫁給客戶的部分；資料中心非 GPU（EPYC CPU 等）、"
                 "Client、Gaming、Embedded 各以年增率推算。合計為營收。FY2028：GPU "), f'TEXT(PL!F{gpu_b},"0")', T("、非 GPU "), f'TEXT(PL!F{dcx_b},"0")',
        T("、Client＋Gaming "), f'TEXT(PL!F{cg_b},"0")', T("、Embedded "), f'TEXT(PL!F{emb_b},"0")', T("，合計 "), f'TEXT(PL!F{rev_b},"0")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型的加速器顆數、每顆售價與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1']}", T("（Coord 頁 F 節）。C0 以公司實績對層級隱含（J 節）：AMD 只揭露資料中心分部（2026 前三季 "),
        f'TEXT(Coord!B{CO["rb1"]},"0.0")', T(" 十億美元），不揭露 Instinct GPU 營收，Siliconomics 也沒有資料中心 CPU 層，結論「"), f"Coord!E{CO['rv']}", T("」，對帳中項目 "),
        f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項。參考：層級 GPU 隱含依月數折算 "), f'TEXT(Coord!C{CO["rb1"]},"0.0")', T("，未超過分部實績；對本模型的 GPU 拆分（假設）差 "),
        f'TEXT(Coord!D{CO["ra1"]},"+0%;-0%")', T("。產業部署 GW 兩層差距另列參考；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（工作單 20261114 第二批第 2 節）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；Hist 頁末新增損益表歷史、"
                  "PL 頁末標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，Non-GAAP）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準、"
                  "Coord 頁（兩讀數目標價、讀取層級、C0 實績：資料中心分部實績 vs 層級 GPU 隱含，結論無可比實績）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    """README A1（公式）版號與 build_id 前綴改 v0.4.7，末尾加本版說明；Sens build_id 格同步。"""
    RD, SE = wb["README"], wb["Sens"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and a1.startswith("=") and f"模型 v{OLD}（" in a1, a1[:120]
    a1 = a1.replace(f"模型 v{OLD}（", f"模型 v{VER}（", 1).replace(f"build_id amd-v{OLD}-", f"build_id amd-v{VER}-", 1)
    RD["A1"].value = a1 + "&" + lit(f"；v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（Coord 頁兩讀數目標價、讀取層級、C0 實績；Hist 損益表歷史、RevDCF、PL 標準損益表、Val 0.7、Sens 加權目標價相對基準），既有數值不變")
    rb = lab(SE, "build_id（README、Sens、HTML")
    bid = str(SE.cell(rb, 4).value)
    assert bid.startswith(f"amd-v{OLD}-"), bid
    SE.cell(rb, 4).value = bid.replace(f"amd-v{OLD}-", f"amd-v{VER}-", 1)


def apply(wb, here, SIL, CR):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("AMD_OVERRIDE"):
        return "v047_post：略過（alt_run／sens_amd 建置）"
    assert CR, "Coord 頁不存在：先跑 python3 tools/coord/alt_run.py amd"
    HR = hist_pl(wb)
    PLA = pl_standard(wb)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb)
    CO = coord(wb, SIL)
    RMA = readme13(wb, RVA, PLA, SEN, CO)
    version(wb)
    return (f"v047_post：Hist 損益表歷史第 {HR['start']} 列起；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；"
            f"Coord F–J 節（C0 判定第 {CO['rv']} 列）；README v1.3 白話列 {len(RMA)} 列")
