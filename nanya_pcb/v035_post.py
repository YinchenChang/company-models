# 南電 v0.3.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10；同次產業樣板 ibiden/v055_post.py、v056_post.py；欣興 unimicron/v044_post.py）、
# 三層協調 C0（規則 v1.4.2：讀取層級、C0（實績））。兩讀數由 tools/coord（建檔程式 COORD.add_sheet）產生，本檔只新增，不改既有數值格
# （以 tools/xlsx_diff.py 對前一版逐格比對驗證）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（公司 IFRS 合併財報，經 stockanalysis.com 轉載）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2030E 情境 B）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑）與 Comparable 倍數
#   Val 頁末：0.7 估值並列四個目標價
#   Sens 頁末：加權目標價相對基準（低案、高案；業務主因整本重算快照 sens.json＋目標價參數表）
#   Coord 頁末：1. 讀取層級、2. 讀取層級核對（同源，不判定）、層級之間（參考）、3. 對帳中年份、4. C0（實績）
#   README 頁末：v1.3 展示層白話列；A1 標題段版號更正為現行版號
# 本檔設計成載板三家（欣興、南電、景碩）共用結構，公司差異集中在 CO 設定（apply 的參數）。
import json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"


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


# ---------- Hist：損益表歷史 ----------
def hist_is(wb, CO):
    H = wb.create_sheet("Hist")
    H["A1"] = (f"Hist — 損益表歷史（v{CO['ver']} 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度＝日曆年。金額：十億新台幣（百萬元 ÷ 1000）；EPS：新台幣。"
               "FY2025 的營收、毛利、營業利益、歸屬母公司淨利、基本 EPS 連到 Q 頁（模型既有已核實值）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, vals in CO["hist"]:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"
            if not isinstance(v, str): c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, CO["hist_src"] + ("；FY2025 引用 Q 頁" if isinstance(vals[-1], str) else ""))
        R[k] = r; r += 1
    for k, labl, f_, fmt in (("noi", "營業外收支淨額（＝稅前利益 − 營業利益）", "={c}{pti}-{c}{oi}", "#,##0.000"),
                             ("sh", "加權平均股數（十億股；＝歸屬母公司淨利 ÷ 基本 EPS）", "={c}{ni}/{c}{eps}", "0.0000")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        sh_v = CO.get("hist_sh") if k == "sh" else None   # 小 EPS 年度（例：0.11、0.32 元）淨利 ÷ EPS 的捨入誤差大：改用來源的加權平均股數
        for j, c in enumerate("CDEF"):
            if sh_v and sh_v[j] is not None:
                H[f"{c}{r}"] = sh_v[j]; H[f"{c}{r}"].font = BLUE
            else:
                H[f"{c}{r}"] = f_.format(c=c, **R)
            H[f"{c}{r}"].number_format = fmt
        if sh_v:
            H.cell(r, 2, "加權平均股數（十億股；藍字＝來源的基本加權股數，FY2025＝歸屬母公司淨利 ÷ 基本 EPS）"); H.cell(r, 7, "Verified／Derived"); H.cell(r, 8, CO["hist_src"])
        else:
            H.cell(r, 7, "Derived"); H.cell(r, 8, "本頁上方各列推得")
        R[k] = r; r += 1
    r += 1
    H.cell(r, 1, "核對（單位：十億新台幣；來源以百萬元四捨五入，各列加總可有 ±0.002 的捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"), ("毛利 − 營業費用 − 營業利益", "={c}{gp}-{c}{opx}-{c}{oi}"),
                     ("稅前利益 − 所得稅 − 非控制權益 − 歸屬母公司淨利", "={c}{pti}-{c}{tax}-{c}{nci}-{c}{ni}")):
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 58
    for c in "CDEF": H.column_dimensions[c].width = 12
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表 ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
COLS = "CDEFGHIJ"
HC = {"C": "D", "D": "E", "E": "F"}   # 標準損益表歷史欄 → Hist 欄
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029}   # 預測欄依賴的 Siliconomics 日曆年（FY＝CY；FY2030E 延伸期沿用 CY2029）


def pl_standard(wb, HR, CR, CO):
    P = wb["PL"]
    rb = lab(P, CO["pl_b"])
    g = lambda label, ex=True: lab(P, label, rb, exact=ex)
    R = dict(rev=g("營收"), gp=g("毛利 ＝", False), opx=g("營業費用"), oi=g("營業利益"), nonop=g("營業外淨額"), pti=g("稅前利益"),
             ni=g(CO["pl_ni"]), sh=g("加權股數"), eps=g("EPS（基本）"))
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", dbr)
    cur = CO["cur"]
    r = P.max_row + 3; s0 = r
    P.cell(r, 1, f"標準損益表（v{CO['ver']} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來五年（FY2026–FY2030，情境 B 基準；FY2026E 為 1–3Q26 實績＋4Q26 假設的錨點）。"
                 f"十億{cur}；EPS {cur}。只引用本頁上方、Hist、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（台灣 IFRS 合併綜合損益表；公司財報、月營收與市場共識都以此為主，沒有另一套 Non-GAAP）。稅後淨利為歸屬母公司；"
                 + CO["pl_gaap_extra"]); r_note = r; r += 1
    P.cell(r, 1, "預測期 FY2026E–FY2029E 直接引用 PL 情境 B 區塊（毛利、營業費用、營業利益、營業外淨額、稅前利益、歸屬母公司淨利）；營業成本＝營收 − 毛利；所得稅列＝稅前利益 − 歸屬母公司淨利（模型把所得稅與非控制權益合併）。"
                 "FY2030E（收斂法延伸期，只推營收與淨利）：毛利率沿用 FY2029 據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2），營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷（1 − 稅率）− FY2029 業外），"
                 "稅前利益＝淨利 ÷（1 − 稅率），隱含營業費用率見表末揭露列，不改淨利與目標價（Andy 2026-11-14）。"); r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJ":
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
        if note: P.cell(r, 11, note)
        L[key] = r; r += 1

    M = "#,##0.0"; PCT = "0.0%"; U = f"十億{cur}"
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    fc4 = lambda k: [f"=PL!{c}{R[k]}" for c in "DEFG"]
    rc = r + 2   # 營業成本列（rev=r, yoy=r+1, cogs=r+2, gp=r+3, gm=r+4, opx=r+5, oi=r+6）
    row("rev", "營收", U, hh("rev") + fc4("rev") + [f"=PL!H{R['rev']}"], M, "FY2026E＝1–3Q26 實績＋4Q26 假設；FY2027E–FY2029E 為驅動因子推導（PL 情境 B）；FY2030E 為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{L['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    assert r == rc
    row("cogs", "營業成本（COGS）", U, hh("cogs") + [f"={c}{L['rev']}-{c}{rc + 1}" for c in "FGHIJ"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", U, hh("gp") + fc4("gp") + [f"=J{L['rev']}*I{rc + 1}/I{L['rev']}"], M, "FY2030E＝營收 × FY2029 毛利率")
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT)
    row("opx", "營業費用（含研發）", U, hh("opx") + fc4("opx") + [None] * 1, M, "FY2030E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）")
    assert L["opx"] == rc + 3
    row("oi", "營業利益", U, hh("oi") + fc4("oi") + [f"=DCF!F{d_ebit}"], M, "FY2030E 取 DCF 頁情境 B 的 EBIT（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外收支）", U, [f"=Hist!{HC[c]}{HR['noi']}" for c in "CDE"] + fc4("nonop") + [f"=J{r + 1}-J{L['oi']}"], M,
        "實績＝稅前利益 − 營業利益；FY2026E–FY2029E＝PL 情境 B 營業外淨額；FY2030E＝稅前利益 − 營業利益")
    row("pti", "稅前利益", U, hh("pti") + fc4("pti") + [f"=PL!H{R['ni']}/(1-tax)"], M, "FY2030E＝歸屬母公司淨利 ÷（1 − 稅率）")
    row("tax", CO["pl_tax_label"], U, hh("tax") + [f"={c}{L['pti']}-{c}{r + 1}" for c in "FGHIJ"], M, CO["pl_tax_note"])
    row("ni", "稅後淨利（歸屬母公司）", U, hh("ni") + fc4("ni") + [f"=PL!H{R['ni']}"], M)
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("nci", "（備註）非控制權益損益", U, hh("nci") + [None] * 5, M, "預測期併入所得稅列，不另列")
    row("sh", "加權股數（十億股）", "十億股", [f"=Hist!{HC[c]}{HR['sh']}" for c in "CDE"] + fc4("sh") + [f"=PL!H{R['sh']}"], "0.0000")
    row("eps", "EPS（基本）", cur, [f"=Hist!{HC[c]}{HR['eps']}" for c in "CDE"] + fc4("eps") + [f"=PL!H{R['eps']}"], "0.00", "模型口徑為基本 EPS（歸屬母公司淨利 ÷ 加權股數）；完全稀釋見下方對照列（只列實績）")
    row("epsd", "（備註）EPS 完全稀釋（公司財報；只列實績）", cur, [f"=Hist!{HC[c]}{HR['epsd']}" for c in "CDE"] + [None] * 5, "0.00")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [f"=C{L['ni']}/Hist!C{HR['ni']}-1"] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", U, [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E 見下一列）", U, [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "CDEFGHI"] + [None], "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "J":
        P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"I{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E；延伸期營業費用＝毛利 − 營業利益，應為 0）", U, [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "J"], "0.000")
    row("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 6 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "IJ"], "0.00%")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利（歷史欄差額＝非控制權益）", U, [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：營業利益 ＋ 其他收入 − 稅前利益", U, [f"={c}{L['oi']}+{c}{L['oth']}-{c}{L['pti']}" for c in COLS], "0.000")
    row("c5", "核對：稅後淨利 − PL 情境 B（FY2026E–FY2030E）", U, [None] * 3 + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHIJ", "DEFGH")], "0.000")
    row("c6", "核對：營收 − PL 情境 B（FY2026E–FY2030E）", U, [None] * 3 + [f"={c}{L['rev']}-PL!{s}{R['rev']}" for c, s in zip("FGHIJ", "DEFGH")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：會計年度＝日曆年，引擎對 FY(t) 直接取 CY(t) 的加速器顆數與 Siliconomics 年別比值；標題「（對帳中）」依 Coord 頁各日曆年旗標，"
                 "FY2030E 由 FY2029 之後的路徑收斂推得，沿用 CY2029 旗標。"); L["ycaveat"] = r; r += 1
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, rows_B=R, d_ebit=d_ebit)
    return L


# ---------- RevDCF：反向 DCF（DCF 頁 C–H＝FY2027–FY2032；評價日 2026-12-31） ----------
def revdcf(wb, CO):
    D = wb["DCF"]; cur = CO["cur"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = 6, 7
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數"), (D.cell(fr, 1).value, D.cell(ty, 1).value)
    assert D.cell(5, 3).value == "FY2027" and D.cell(5, 8).value == "FY2032", [D.cell(5, c).value for c in range(3, 9)]
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"   # 情境 B（Val 頁 C 欄）
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（v{CO['ver']} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與現金流倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（三情境營收相同）不動，營業利益率與現金流倍數只動 FY2028–FY2032。反解值超出合理範圍時照實列出，不外推")
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
    kv("px", f"現價（{cur}）", "=px", "#,##0.00")
    kv("v0", f"支持現價所需的評價日每股價值（{cur}）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", f"淨現金滾到評價日（情境 B；十億{cur}）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁情境 B 欄）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", f"支持現價所需的企業價值 EV（十億{cur}）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", f"模型企業價值 EV（情境 B；十億{cur}）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", f"模型：情境 B 評價日每股價值（{cur}）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}
    DC, RC = "CDEFGH", "BCDEFG"

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate(DC):
            c = RV.cell(r, 2 + i, fn(dc, RC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", f"營收（十億{cur}）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", f"營業利益 EBIT（十億{cur}）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C$8" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", f"NOPAT＋折舊 − 資本支出（十億{cur}）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", f"營運資金增加（十億{cur}）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", f"現值（十億{cur}，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", f"營收 × 貢獻係數合計（十億{cur}）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", f"與營業利益無關的 EV（十億{cur}）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<B{A["m0"]},"低於模型：現價不需要這麼高的利益率","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的現金流規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", f"不隨倍數變動的 EV：FY2027 現值（十億{cur}）", f"=B{Y['pv']}", "#,##0.0")
    kv("b", f"隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（十億{cur}）", f"=SUM(C{Y['pv']}:G{Y['pv']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", f"現價隱含的 FY2028 營收（十億{cur}）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", f"模型：FY2028 營收（情境 B；十億{cur}）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近區間端點，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", f"中點的評價日每股價值（{cur}）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C$8+$B${A['nc']})/sh_v")
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
    kv("w_res", f"核對：隱含 WACC 下的每股價值 − 支持現價所需值（{cur}，應接近 0；區間內無解時不為 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, f"12 個月後前瞻 EPS（情境加權；{cur}）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(f" {cur}；要支持現價 "), TX("px", "#,##0"), T(f" {cur}，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(f" 十億{cur}，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(f" {cur}相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(f" 倍（同業中位數 × {CO['name']}過去十年相對同業的倍數比 "), TX(coef, "0.00"), T(f"）。換言之，市場目前給{CO['name']}的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，歷史比值為 "), TX(coef, "0.00"), T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70
    for c in "BCDEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    return A


# ---------- Val：估值並列四個目標價（版面規格 v1.3 G 節；只新增列） ----------
def val_g(wb, CO):
    V = wb["Val"]; cur = CO["cur"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    pxr = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, f"v{CO['ver']} 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", cur, "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(f" {cur}；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + f"Val!$E${lab(V, '市場共識目標價')}")]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=Val!A{pxr}"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "股價基準日收盤（In 頁 px）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here, CO):
    SE = wb["Sens"]; V = wb["Val"]; cur = CO["cur"]
    snap = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    tp_cell = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    t0 = lab(SE, "目標價參數敏感度"); hr = t0 + 1
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, f"v{CO['ver']} 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（{CO['folder']}/{CO['sens_py']}），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", f"低案：加權目標價（{cur}）", f"高案：加權目標價（{cur}）", f"快照基準加權目標價（{cur}）"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    b = snap["_base"]["tp"]
    for case in snap["cases"]:
        k = case["key"]; z = snap["biz"][k]
        SE.cell(r, 1, case["label"]); SE.cell(r, 4, case["setting"] + "（整本重算快照）")
        for col, side, tcol in ((2, "低", 5), (3, "高", 6)):
            if side in z:
                c = SE.cell(r, tcol, z[side]["tp"]); c.font = BLUE; c.number_format = "#,##0.00"
                SE.cell(r, col, f"={'EF'[tcol - 5]}{r}/$G{r}-1").number_format = "+0.0%;-0.0%"
        c = SE.cell(r, 7, b); c.font = BLUE; c.number_format = "#,##0.00"
        r += 1
    for sr in src:
        SE.cell(r, 1, f"=Sens!A{sr}"); SE.cell(r, 4, f"=\"Sens 頁目標價參數敏感度：\"&Sens!B{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
        r += 1
    last = r - 1
    SE.cell(r, 1, f"快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 {CO['folder']}/{CO['sens_py']}）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r, src=src)


# ---------- Coord：讀取層級、讀取層級核對（同源）、層級之間、對帳中年份、C0（實績） ----------
def coord_c0(wb, SIL, CR, CO):
    CW = wb["Coord"]; P = wb["PL"]
    rb = lab(P, CO["pl_b"]); ra = lab(P, CO["pl_a"])
    R = dict(unv=lab(P, "NVIDIA 加速器顆數 × 情境係數", rb), uot=lab(P, "非 NVIDIA 加速器顆數 × 情境係數", rb))
    RA = dict(unv=lab(P, "NVIDIA 加速器顆數 × 情境係數", ra), uot=lab(P, "非 NVIDIA 加速器顆數 × 情境係數", ra))
    N = SIL.NAMES; Vs = SIL.V
    base_tot = N["SIF_UnitsPlat_Total_Base"]; raw_tot = N["SIF_UnitsPlat_Total"]
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    nb = f"Siliconomics {SIL.VERSION}"
    r = CW.max_row + 3
    CW.cell(r, 1, f"三層協調 C0（v{CO['ver']}；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; ROW["s1"] = r; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", SIL.VERSION, f"siliconomics/data/{SIL.FILE}（只讀 SIF_／abf_ 名稱，不讀 Wattonomics 覆蓋 Siliconomics；C4）", SIL.VERSION, "本模型屬硬體供應鏈個股，讀 Siliconomics"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 敏感度案例；不改基準"),
          ("加速器顆數（NVIDIA 四世代、非 NVIDIA 四平台）", "唯讀上游", "markets/network In 頁 u_*／g29_*（與 Siliconomics SIF_UnitsPlat_* 同值，本頁 2 節逐年核對）", "AI 網通市場模型", "PL 引擎 A 節；情境 A＝原始顆數"),
          ("顆數情境係數 B（CY2028）", "=um28_B", "Siliconomics SIF_UnitsCoef_Base（經 tsmc 模型 In 頁讀入；規則 C3）", SIL.VERSION, "情境 B（基準）；A＝1.0、C 個股自訂"),
          ("顆數情境係數 B（CY2029）", "=um29_B", "同上", SIL.VERSION, "同上")]
    for key, nm, silk in CO["sil_reads"]:
        rd.append((nm, f"={key}", f"Siliconomics {silk}（SIL 讀入 In 頁 {key}）", SIL.VERSION, CO["sil_use"]))
    for key, nm, src in CO["peer_reads"]:
        rd.append((nm, f"={key}", src, "唯讀同業模型", CO["sil_use"]))
    for y in (2027, 2028, 2029):
        rd.append((f"Wattonomics 錨定係數 CY{y}（敏感度案例用）", f"=wta_{y % 100}", "Siliconomics SIF_WTAnchor_Base（R4 比值：Wattonomics ÷ 本庫推得上線 GW）", SIL.VERSION, "Sens 頁「Wattonomics 錨定」案例；基準不乘"))
    rd.append(("對帳中旗標與另一側讀數比值（CY2026–CY2029）", "見本頁上方", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION, "兩讀數目標價（C4）"))
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
        r += 1
    r += 1
    CW.cell(r, 1, "2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0 或由同一層級值推得，無檢驗力，故不判定、不計入對帳中；C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）").font = BOLD
    ROW["s2"] = r; r += 1
    CW.cell(r, 1, "年度口徑：會計年度＝日曆年，引擎對 FY(t) 取 CY(t) 的顆數與年別比值，不平移。"); ROW["caveat"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定（同源，不判定）", "層級名稱與版本／口徑／證據等級／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr2"] = r; r += 1
    first = r

    def line(label, unit, yr_, f_model, f_layer, fmt, note, judge=SAME):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_model).number_format = fmt
        c = CW.cell(r, 5, f_layer); c.number_format = fmt
        if not (isinstance(f_layer, str) and f_layer.startswith("=")): c.font = BLUE
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ycol = {2026: "D", 2027: "E", 2028: "F", 2029: "G"}
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        c = ycol[y]
        line("加速器顆數合計（情境 B，NVIDIA＋非 NVIDIA）vs Siliconomics 基準", "百萬顆", f"CY{y}", f"=PL!{c}{R['unv']}+PL!{c}{R['uot']}", base_tot[i + 1], "#,##0.000",
             f"{nb} SIF_UnitsPlat_Total_Base（原始顆數 × 基準係數 SIF_UnitsCoef_Base）；Derived；模型情境 B 係數與基準相同，故應為 0")
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        c = ycol[y]
        line("加速器顆數合計（情境 A，原始顆數）vs Siliconomics 原始顆數", "百萬顆", f"CY{y}", f"=PL!{c}{RA['unv']}+PL!{c}{RA['uot']}", raw_tot[i + 1], "#,##0.000",
             f"{nb} SIF_UnitsPlat_Total；Derived；來源 markets/network")
    for key, nm, silk in CO["sil_reads"]:
        line(f"{nm}：模型 In 頁 vs Siliconomics", "x", "—", f"={key}", Vs[silk]["base"], "0.000",
             f"{nb} {silk}（直接讀取，鏈路完整性核對；證據等級見 Siliconomics In 頁）")
    last = r - 1
    ROW["first"], ROW["last"] = first, last
    CW.cell(r, 1, "讀取層級列數（同源，不判定）"); CW.cell(r, 7, f'=COUNTIF(G{first}:G{last},"{SAME}")'); ROW["cnt_same"] = r; r += 1
    CW.cell(r, 1, CO["c0_nocomp_note"]); ROW["nocomp"] = r; r += 2
    # ---- 層級之間（參考） ----
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, f"以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是{CO['name']}自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。"); r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    for i, y in enumerate((2026, 2027, 2028, 2029)):
        CW.cell(r, 1, "加速器顆數合計：Siliconomics 基準 vs Wattonomics 換算值"); CW.cell(r, 2, "百萬顆"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, f"=E{first + i}").number_format = "#,##0.000"; CW.cell(r, 5, f"=E{first + i}*{'BCDE'[i]}${CR['ratio']}").number_format = "#,##0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Wattonomics {wat_ver} 經 Siliconomics SIF_DeployGW_AltRatio_{y}（另一側 ÷ 現行基準）換算；Derived；對帳中年份依 C4 並列兩讀數目標價，個股不裁決")
        r += 1
    r += 1
    # ---- 3. 對帳中年份 ----
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY2027–FY2029 及其後由 FY2029 收斂推得的延伸期；FY2026 為實績＋第四季假設的錨點、CY2026 非對帳中）。"
                      "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(f" {CO['cur']}（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(f" {CO['cur']}（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入（顆數 × 另一側比值），個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3a"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110
    return ROW


def coord_actual(wb, SIL, CO, ROW):
    """4. C0（實績）：公司月營收／季報實績 vs Siliconomics 層級隱含值（同口徑合併營收）。"""
    CW = wb["Coord"]; A = CO["act"]; cur = CO["cur"]
    V = SIL.V[A["sil_key"]]
    r = CW.max_row + 3
    CW.cell(r, 1, f"4. C0（實績）：公司實績 vs 層級隱含值（v{CO['ver']}；規則 v1.4.2 C0 第 2 點、tools/coord/README.md；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["a4"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（2026 年 1–10 月合併營收月報累計、3Q26 合併營收；只用已公布期間，不用預測或指引，實績本身是公司原文、獨立）；"
                  f"層級側＝Siliconomics 現行基準的{CO['name']} 2026 ABF 載板營收（US$B）換算成同口徑合併營收（× 換算匯率 ÷ ABF 占營收比率）。層級值是日曆年 CY2026：主讀數＝已公布期間累計實績 vs 層級 × 已公布月數 ÷ 12（線性時差近似，Derived）；"
                  "第二讀數＝最近一季 × 4 vs 層級（約等於該季中點的年化速率：3Q26 中點 8 月中晚於年中，成長年中偏高；主讀數在成長年偏低）；兩讀數的差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, A["scope_note"]); ROW["scope"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑", "鍵名（追蹤用；HTML 不顯示）"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; r += 1
    inp = {}

    def put(key, label, unit, per, val, grade, src, fmt, keyname=""):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        if keyname: CW.cell(r, 9, keyname)
        inp[key] = r; r += 1
    U = f"十億{cur}"
    put("cum", "合併營收：2026 年 1–10 月累計（月營收）", U, "2026-01～2026-10", A["cum"], "Verified", A["cum_src"], "#,##0.000")
    put("mo", "已公布月數", "月", "2026-01～2026-10", 10, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("q3", "合併營收：3Q26（季報）", U, "2026-07～2026-09", A["q3"], "Verified", A["q3_src"], "#,##0.000")
    put("abf", f"Siliconomics 層級值：{CO['name']} 2026 ABF 載板營收", "US$B", "CY2026", V["base"], "Derived（層級值）",
        f"{SIL.VERSION}（快照檔 {SIL.FILE}；D 欄為寫死值，讀自快照；鍵名見 I 欄）；該列原註：{A['sil_note']}", "0.000", f"SIF_{A['sil_key']}")
    put("fx", "換算匯率（Siliconomics 該列原註的換算值）", f"{cur}／US$", "—", A["sil_fx"], "Derived", f"寫死值，取自 {SIL.VERSION} {A['sil_key']} 列原註（÷ {A['sil_fx']}）；本模型 In 頁 fx 為全庫共用匯率，與此不同，故用原註值以免換算失真", "0.00")
    put("sh", "ABF 占合併營收比率（FY2026；In 頁 mx_abf 連格）", "%", "FY2026", "=mx_abf", "Interested-party", A["sh_src"], "0.0%", "In 頁 mx_abf")
    put("lay", "層級隱含合併營收", U, "CY2026", f"=D{inp['abf']}*D{inp['fx']}/D{inp['sh']}", "Derived", "層級 ABF 營收 × 換算匯率 ÷ ABF 占比；連到 D 欄各輸入格", "#,##0.0")
    put("anchor", "參考：本模型 FY2026 營收錨點（Q 頁；1–3Q26 實績＋4Q26 假設）", U, "FY2026", "=a26_rev", "Derived", "Q 頁 FY2026 錨點；層級值由它推得（見讀法列）", "#,##0.0", "Q 頁 a26_rev")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def line(label, per, f_a, f_l, note, judge):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, U); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    REF = "參考（同源校準）"
    ra = line("主讀數：2026 年 1–10 月累計合併營收 vs 層級 × 月數 ÷ 12", "2026-01～10 vs CY2026 × 10/12", f"=D{inp['cum']}", f"=D{inp['lay']}*D{inp['mo']}/12",
              "實績＝月營收累計；層級＝Siliconomics ABF 載板營收換算的合併營收 × 10 ÷ 12（線性時差近似）。成長年中，前 10 個月占全年比重低於 10／12，此讀數偏低", REF)
    rb = line("第二讀數：3Q26 合併營收 × 4 vs 層級", "2026-07～09 × 4 vs CY2026", f"=D{inp['q3']}*4", f"=D{inp['lay']}",
              "最近一季年化，約等於該季中點（8 月中）的年化速率；中點晚於年中，成長年中此讀數偏高", REF)
    CW.cell(r, 1, "C0（實績）結論"); CW.cell(r, 7, "無可比實績（層級值同源）"); CW.cell(r, 8, A["concl"])
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：機械判定（兩讀數差距都超過 ±15% 且同號＝對帳中；只一個＝時差邊界；都在內＝一致）；因層級值同源，僅供參考、不計入")
    CW.cell(r, 7, f'="參考："&IF(AND(ABS(F{ra})>0.15,ABS(F{rb})>0.15,F{ra}*F{rb}>0),"對帳中",IF(OR(ABS(F{ra})>0.15,ABS(F{rb})>0.15),"時差邊界","一致"))')
    ROW["mech"] = r; r += 1
    CW.cell(r, 1, "參考：已公布期間實績占本模型 FY2026 營收錨點（進度，不是差距）"); CW.cell(r, 2, U); CW.cell(r, 3, "2026-01～10 vs FY2026")
    CW.cell(r, 4, f"=D{inp['cum']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['anchor']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"; CW.cell(r, 7, "參考（不判定）")
    CW.cell(r, 8, "錨點含 4Q26 假設（11、12 月尚未公布）；10 月營收 × 2 可對照 4Q26 剩餘月份")
    ROW["prog"] = r; r += 1
    CW.cell(r, 1, A["read"]); ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表；同源校準列不計）"); CW.cell(r, 7, f'=COUNTIF(G{ra}:G{rb},"對帳中")'); ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["inp"] = ra, rb, inp
    return ROW


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, PLA, RVA, SNS, CR, CO, ROW):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; P = wb["PL"]; DC = wb["DCF"]; cur = CO["cur"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    pa, pb = lab(P, CO["pl_a"]), lab(P, CO["pl_b"])
    gm_a = lab(P, "毛利率", pa, exact=True); eps_a = lab(P, "EPS（基本）", pa, exact=True); eps_b = lab(P, "EPS（基本）", pb, exact=True); da_a = lab(P, "折舊攤提 ＝", pa)
    tv_row = lab(DC, "終值現值占 EV", lab(DC, "情境 B（基準）"))
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_pm, s_sh, s_wts = sens_r(CO["sens_o3"][0]), sens_r(CO["sens_o3"][1]), sens_r("情境權重")
    s_w = next(sr for sr in SNS["src"] if "DCF 權重" in str(SE.cell(sr, 1).value))
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；v{CO['ver']} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    hi_lo = lambda x: f'IF({x}>=0,"高於","低於")&"現價 "&TEXT(ABS({x}),"0.0%")'
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(f" {cur}，"), hi_lo(up), T("，評等"), rec, T("：同業比較法 "), TX(cadj, "#,##0"), T(f" {cur}，現金流折現只有 "), TX(dcf, "#,##0"), T(f" {cur}（擴產期資本支出大）。"))
    put("rec", T("投資建議「"), rec, T(f"」成立的前提：Comparable 採用{CO['name']}過去十年相對同業的本益比比值（"), TX(coef, "0.00"), T(" 倍）、擴產期資本支出照本模型排程。不採相對倍數調整時加權目標價 "),
        TX(noadj, "#,##0"), T(f" {cur}（"), TX(f"{noadj}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"({noadj}/px-1)"), T("）；DCF 權重降到三成時 "), TX(f"Sens!C{s_w}", "#,##0"),
        T(f" {cur}（"), TX(f"Sens!C{s_w}/px-1", "+0.0%;-0.0%"), T("，評等"), REC5(f"(Sens!C{s_w}/px-1)"), T("）。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(f" {cur}，現價 "), TX("px", "#,##0"), T(f" {cur}，"), hi_lo(up), T("，評等"), rec, T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.0"), T(f" {cur}）得 "), TX(cadj, "#,##0"), T("，DCF 得 "), TX(dcf, "#,##0"), T("，各占一半；分析師共識 "), TX(cons, "#,##0"), T(f" {cur}遠高於本模型。"))
    put("o2", T("DCF 偏低：擴產期資本支出大，明確期自由現金流薄；終值現值占 EV "), TX(f"DCF!C{tv_row}", "0%"), T("。WACC "), TX("wacc", "0.0%"), T("；要支持現價，折現率須降到 "), TX(rv("w1"), "0.0%"),
        T("，或 FY2028 起現金流（營收與成本同步）放大 "), TX(rv("s1"), "0.00"), T(" 倍。"))
    put("o3", T(f"最大變數（加權目標價相對基準）：{CO['sens_o3_txt'][0]}"), TX(f"Sens!B{s_pm}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_pm}", "+0.0%;-0.0%"), T(f"；{CO['sens_o3_txt'][1]}"), TX(f"Sens!B{s_sh}", "+0.0%;-0.0%"),
        T("／"), TX(f"Sens!C{s_sh}", "+0.0%;-0.0%"), T("；情境權重全給保守 "), TX(f"Sens!B{s_wts}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o4", T("核心問題：漲價能否抵過擴產折舊。情境 A 毛利率由 FY2027 "), TX(f"PL!E{gm_a}", "0.0%"), T(" 到 FY2029 "), TX(f"PL!G{gm_a}", "0.0%"), T("，折舊由 FY2026 "), TX(f"PL!D{da_a}*10", "#,##0"),
        T(" 億元增至 FY2029 "), TX(f"PL!G{da_a}*10", "#,##0"), T(" 億元；情境 B 的 FY2028 EPS "), TX(f"PL!F{eps_b}", "0.0"), T(f" {cur}，比情境 A 的 "), TX(f"PL!F{eps_a}", "0.0"), T(" 低 "), TX(f"-(PL!F{eps_b}/PL!F{eps_a}-1)", "0.0%"), T("。"))
    put("o5", T("對帳中：CY2027–CY2029 的加速器顆數，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(f" {cur}（"), f"Coord!C{CR['rec']}", T("）。公司實績（1–10 月營收）對層級值：層級值由本模型校準，無可比實績。"))
    put("drv", *CO["drv_parts"](P, pb))
    RD.cell(r, 1, CO["ver_line"]); out["ver"] = r; r += 1
    return out


def version(wb, CO):
    RD = wb["README"]
    a1 = RD["A1"].value
    old = CO["a1_old"]
    assert old in a1, a1[:80]
    RD["A1"].value = a1.replace(old, old.replace(CO["a1_oldver"], f"v{CO['ver']}"), 1)


def apply(wb, here, SIL, CR, CO):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"   # check_layout_v13 的兩讀數檢查要求 Coord 頁該列標籤以「加權目標價」開頭（數值與公式不變）
    HR = hist_is(wb, CO)
    ROW = coord_c0(wb, SIL, CR, CO)
    ROW.update(coord_actual(wb, SIL, CO, ROW))
    PLA = pl_standard(wb, HR, CR, CO)
    RVA = revdcf(wb, CO)
    VG = val_g(wb, CO)
    SNS = sens_tor(wb, here, CO)
    RMA = readme13(wb, PLA, RVA, SNS, CR, CO, ROW)
    version(wb, CO)
    if CO.get("pl_b_title"):   # PL 情境 B 標題的舊顆數係數文字改讀 Siliconomics 基準係數（文字列；放最後，前面各步仍以文字找列）
        _pl = wb["PL"]; _r = lab(_pl, CO["pl_b"])
        _pl.cell(_r, 1).value = CO["pl_b_title"]
    if CO.get("sources"):
        SO = wb["Sources"]; rr = SO.max_row + 2
        for a, b in CO["sources"]:
            SO.cell(rr, 1, a); SO.cell(rr, 2, b); rr += 1
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=ROW)
    return (f"v{CO['ver']} post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 相對基準 {SNS['t0']} 起；"
            f"Coord 讀取層級 {ROW['first']}–{ROW['last']}、C0（實績）{ROW['a4']}–{ROW['cnt4']}；README v1.3 白話列 {len(RMA)} 列")
