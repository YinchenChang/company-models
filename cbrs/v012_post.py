# Cerebras v0.1.2（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調（讀取層級、C0 公司實績對帳、兩讀數判定）。
# 依據：docs/workorders/20261114_第二批13家_v1.3與三層協調.md、docs/workorders/20261114_C0實績對帳補列_七家.md、
#       docs/plan/三層基準協調規則.md v1.4.2（C0 第 2 點、C4）、docs/plan/個股HTML版面規格.md v1.3、tools/coord/README.md。
# 樣板：AMD v0.4.7（amd/v047_post.py）、NVIDIA v0.3.10。只新增，不改既有數值格：
#   PL 頁末：標準損益表（FY2024A–FY2025A GAAP 實績＋FY2026E–FY2032E 情境 B 模型口徑）、比率、核對列
#   RevDCF 頁（新）：反向 DCF；只讀其他頁，無任何頁引用本頁
#   Val 頁末 0.7：四個目標價；Sens 頁末：加權目標價相對基準（白話標籤）
#   Coord 頁（新）：讀取層級、兩讀數（C4）判定、層級之間（參考）、C0（實績）
#   README 頁末：v1.3 展示層白話列；A1 版號
# 本模型不讀 Siliconomics（唯一不讀的硬體個股）：估值不依賴 Siliconomics 對帳中年份，兩讀數依 C4 不適用（另一側讀數無傳導路徑，與現行相同），
# 不建 co_flag 具名範圍、不加橫幅「另一側讀數」格。
from openpyxl.styles import Font, PatternFill, Alignment

VER, OLD, BUILT = "0.1.2", "0.1.1", "2026-11-15"
CUR = "0.1.3"   # v0.1.3（2026-11-15）：滾入 3Q26 實績與 4Q26 指引；VER 保留 0.1.2（本檔新增頁與列自 v0.1.2 起）
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); HDR = PatternFill("solid", fgColor="DDEBF7")
S8K1 = "https://www.sec.gov/Archives/edgar/data/0002021728/000162828026044941/cbrsannouncesfinancialresu.htm"
S8K2 = "https://www.sec.gov/Archives/edgar/data/0002021728/000162828026056186/cbrsannouncesfinancialresu.htm"
SQ2 = "https://www.sec.gov/Archives/edgar/data/0002021728/000162828026056357/cbrs-20260630.htm"
S8K3 = "https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073366/cbrsannouncesfinancialresu.htm"
SQ3 = "https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073494/cbrs-20260930.htm"


def lit(t):
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
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
YEARS = ["FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJK"   # 本表欄：C、D＝Q 頁 FY2024、FY2025（GAAP 實績）；E–K＝PL 情境 B 區塊 D–J（FY2026E 錨點、FY2027E–FY2032E）


def hdr(ws, r, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR


# ---------- PL：標準損益表 ----------
def pl_standard(wb, PR):
    P, Q = wb["PL"], wb["Q"]
    R = PR["B"]
    q = dict(rev=lab(Q, "營收合計（GAAP）"), gp=lab(Q, "毛利（GAAP）"), oi=lab(Q, "營業損益（GAAP）"), ni=lab(Q, "淨利（GAAP）"))
    assert Q.cell(3, 2).value == "FY2024" and Q.cell(3, 3).value == "FY2025", (Q.cell(3, 2).value, Q.cell(3, 3).value)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史 FY2024A–FY2025A（GAAP 實績）＋未來七年（FY2026E 錨點、FY2027E–FY2032E，情境 B 基準）。US$B；EPS 美元。"
                 "只引用本頁上方與 Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP 近似。歷史欄＝公司 GAAP 財報（424B4；Q 頁）；預測欄＝本模型 EBIT 口徑——營收、毛利為公司核心口徑（排除客戶認股權證攤銷沖減與代墊資料中心收入），"
                 "營業費用含經常性股票薪酬（GAAP 費用），只排除 IPO 一次性股票薪酬與其薪資稅。公司另有 non-GAAP「核心」營業損益（不扣股票薪酬），但市場評價看 GAAP 虧損與股票薪酬負擔，故本表以 GAAP 近似呈現。")
    r_note = r; r += 1
    P.cell(r, 1, "歷史只列兩年：FY2023 GAAP 損益（424B4）本版未取得原文（待補，見 README 未解事項）；FY2024–FY2025 的其他收入與所得稅未分開取得，其他收入列＝淨利 − 營業損益（含所得稅與一次性項目，"
                 "FY2025 含遠期合約負債消滅利益約 3.6 億美元），稅前利益、所得稅與 EPS（上市前）列「—」。FY2026E＝1Q26–3Q26 實績＋4Q26 指引中點（Q 頁；v0.1.3 起）。")
    r_note2 = r; r += 1
    hdr(P, r, ["項目", "單位"] + YEARS + ["說明"])
    r_hdr = r; r += 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None:
                continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1

    M = "#,##0.000"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFGHIJ"]
    hq = lambda key: [f"=Q!B{q[key]}", f"=Q!C{q[key]}"]
    row("rev", "營收", "US$B", hq("rev") + fc("rev"), M, "歷史＝GAAP 營收合計；預測＝核心營收（雲端＋硬體）")
    row("yoy", "營收年增率（YoY）", "%", ["—"] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 年增率需 FY2023（未取得）")
    r_c = r
    row("cogs", "營業成本（COGS）", "US$B", [f"={c}{L['rev']}-{c}{r_c + 1}" for c in COLS], M, "＝營收 − 毛利（預測含雲端折舊）")
    row("gp", "毛利", "US$B", [f"=Q!B{q['gp']}", f"=Q!C{q['gp']}"] + fc("gp"), M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT)
    row("opx", "營業費用（研發＋銷管，含經常性股票薪酬）", "US$B", [f"={c}{L['gp']}-{c}{r_c + 4}" for c in COLS], M, "＝毛利 − 營業利益（預測＝核心營業費用＋經常性股票薪酬）")
    row("oi", "營業利益", "US$B", [f"=Q!B{q['oi']}", f"=Q!C{q['oi']}"] + fc("ebit"), M, "歷史＝GAAP 營業損益；預測＝EBIT（核心營業損益 − 經常性股票薪酬）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額）", "US$B", [f"=Q!B{q['ni']}-Q!B{q['oi']}", f"=Q!C{q['ni']}-Q!C{q['oi']}"] + fc("oth"), M, "歷史＝淨利 − 營業損益（含所得稅與一次性項目）；預測＝利息與其他收益淨額")
    row("pti", "稅前利益", "US$B", ["—", "—"] + [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS[2:]], M)
    row("tax", "所得稅", "US$B", ["—", "—"] + fc("tax"), M)
    row("ni", "稅後淨利", "US$B", [f"=Q!B{q['ni']}", f"=Q!C{q['ni']}"] + fc("ni"), M, "歷史＝GAAP 淨利（FY2025 含一次性利益）")
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "稀釋股數（含客戶認股權證）", "十億股", ["—", "—"] + fc("sh"), M)
    row("eps", "EPS（完全稀釋；模型口徑）", "US$", ["—", "—"] + fc("eps"), "0.00", "上市前（FY2024–FY2025）不適用；預測＝淨利 ÷ 稀釋股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", ["—"] + [f"={c}{L['yoy']}" for c in COLS[1:]], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["—"] + [f'=IF({p}{L["ni"]}<=0,"不適用（基期虧損）",{c}{L["ni"]}/{p}{L["ni"]}-1)' for p, c in zip(COLS, COLS[1:])], PCT, "基期為虧損時不計算")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS], "0.000")
    row("c3", "核對：稅前利益 − 所得稅 − 稅後淨利（FY2026E–FY2032E）", "US$B", [None, None] + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['ni']}" for c in COLS[2:]], "0.000")
    row("c4", "核對：EPS − 情境 B 區塊（FY2026E–FY2032E）", "US$", [None, None] + [f"={c}{L['eps']}-PL!{s}{R['eps']}" for c, s in zip(COLS[2:], "DEFGHIJ")], "0.000")
    P.column_dimensions["L"].width = max(P.column_dimensions["L"].width or 0, 50)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0)
    return L


# ---------- RevDCF：反向 DCF（樣板 AMD v0.4.7；Comparable 改為前瞻 P/S） ----------
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
    frev, pmed, dcf_tp, coef = vb("12 個月後的前瞻營收（情境加權）"), vb("同業前瞻 P/S 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（v{VER} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑）；Comparable 部分反解「12 個月後前瞻營收（情境加權）× 前瞻 P/S ÷ 股數＝現價」。"
                "明確期第一年 FY2027 維持模型值不動，只動 FY2028–FY2032（含終值）。反解值超出合理範圍時照實列出，不外推")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    hdr(RV, r, ["項目", "值", "說明"]); r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note: RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數")
    kv("nc", "淨現金滾到評價日（含 stub 期 FCFF；US$B）", f"={ncr}", "#,##0.0", "Val 頁：淨現金與投資＋stub 期 FCFF − 已宣告股利（與 WACC 無關）")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；US$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f'=IF(B{A["ev"]}>0,B{A["evr"]}/B{A["ev"]},"不適用（模型 EV 為負）")', "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${ps}", "0.00")
    r += 1
    yh = r
    hdr(RV, r, ["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)]); r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, "BCDEFG"[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.00")
    yr("ebit", "營業利益 EBIT（扣股票薪酬；US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.00")
    yr("m", "EBIT 率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvr}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元 EBIT 對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.00")
    yr("wc", "營運資金增加（US$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.00")
    yr("pv", "現值（US$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.00")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期 EBIT 率（FY2028–FY2032 統一一個 EBIT 率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.00")
    kv("E0", "與營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.00")
    kv("m0", "模型：基準的等效 EBIT 率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型 EBIT 率")
    kv("m1", "現價隱含的 EBIT 率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；EBIT 率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（US$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.00", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（US$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.00")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.00")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.00")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    hdr(RV, r, ["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"]); r += 1

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
    kv("w_ok", "判讀（隱含 WACC）", f'=IF(ABS(B{A["w_res"]})<0.01,"已收斂",IF(E{r - 5}<B{A["v0"]},"搜尋下限仍不足：折現率降到下限（永續成長率＋0.5 個百分點）時每股價值仍低於支持現價所需值","未收斂"))', "@")
    r += 1
    RV.cell(r, 1, "B. Comparable：支持現價所需的前瞻 P/S").font = BOLD; r += 1
    hdr(RV, r, ["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"]); r += 1

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    RV.cell(r, 1, "12 個月後前瞻營收（FY2028，情境加權；US$B）"); RV.cell(r, 3, f"={frev}").number_format = "0.00"; A["frev"] = r; r += 1
    cmp("ps1", "只看 Comparable：支持現價所需的前瞻 P/S（倍）", f"=px*sh_v/{frev}", f"=tp_comp_adj*sh_v/{frev}", "0.00")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻 P/S 中位數）", f"=B{A['ps1']}/{pmed}", f"={coef}", "0.00")
    cmp("ps2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻 P/S（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{frev}/sh_v)", f"=tp_comp_adj*sh_v/{frev}", "0.00")
    RV.cell(r, 1, "同業前瞻 P/S 中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.00"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 美元；要支持現價 "), 'TEXT(px,"0.00")', T(" 美元，企業價值須由 "), f'TEXT(B{A["ev"]},"0.0")',
                     T(" 提高到 "), f'TEXT(B{A["evr"]},"0.0")', T(" 十億美元。單靠一個因素達成時，需要："), f'IF(ABS(B{A["w_res"]})<0.01,"折現率降到 "&TEXT(B{A["w1"]},"0.0%"),"折現率即使降到 "&TEXT(B{A["w1"]},"0.0%")&"（搜尋下限）仍不足")', T("；或 FY2028–FY2032 營收全面放大 "),
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0.0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0.0")', T("）；或 EBIT 率提高到 "),
                     f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於 FY2028 前瞻營收（情境加權）的 "), f'TEXT(B{A["ps1"]},"0.0")',
                     T(" 倍營收，"), f'IF(B{A["ps1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["ps1"]},"0.0")',
                     T(" 倍（上市未滿一年、不做相對倍數調整）。加權法下 DCF 不變時，Comparable 只需 "), f'TEXT(B{A["ps2"]},"0.0")', T(" 倍即可支持現價。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70
    for c in "BCDEFGH": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# ---------- Val：估值並列四個目標價 ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    dcf, cu, cadj, tp = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）")
    cons = vb("市場共識目標價")
    frev, pmed, coef = vb("12 個月後的前瞻營收（情境加權）"), vb("同業前瞻 P/S 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價')}"
    r_px = lab(V, "現價（")
    r = V.max_row + 3
    V.cell(r, 1, f"0.7 估值並列：四個目標價（v{VER} 起；個股 HTML 版面規格 v1.3 G 節，Andy 2026-11-14 09:03；只引用上方各格，不回饋計算）").font = BOLD; r0 = r; r += 1
    hdr(V, r, ["目標價", "美元", "相對現價", "說明"]); r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後；負值取 0）以情境權重 A "), 'TEXT(tp_wA,"0.0")', T("／B "), 'TEXT(tp_wB,"0.0")', T("／C "), 'TEXT(tp_wC,"0.0")', T(" 加權"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻 P/S 中位數 "), f'TEXT({pmed},"0.00")', T(" 倍 × 相對倍數調整係數 "), f'TEXT({coef},"0.00")', T("（上市未滿一年不調整）× FY2028 前瞻營收（情境加權）"), f'TEXT({frev},"0.00")', T(" 十億美元 ÷ 稀釋股數"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法）", cu, F(T("同業前瞻 P/S 中位數 × FY2028 前瞻營收 ÷ 稀釋股數；係數為 1，與 ② 相同"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), 'TEXT(tp_w_dcf,"0%")', T("＋Comparable（相對倍數調整後）"), 'TEXT(1-tp_w_dcf,"0%")')),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, f"=A{r_px}"); V.cell(r, 2, "=px").number_format = "0.00"; V.cell(r, 4, "equibles 11-06 收盤 − 當日漲跌（In 頁 px 列；Derived）"); r += 1
    return r0


# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb):
    SE, V = wb["Sens"], wb["Val"]
    tp = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    dcf, cadj = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}", f"Val!$B${lab(V, 'Comparable 目標價（相對倍數調整後）')}"
    r_dcfs = lab(V, "DCF 情境價值推到 12 個月後"); r_rv = lab(V, "各情境 12 個月後的前瞻營收")
    pmed, coef = f"Val!$B${lab(V, '同業前瞻 P/S 中位數')}", f"Val!$B${lab(V, '相對倍數調整係數')}"
    assert SE["A3"].value == "案例" and SE["I3"].value == "情境 B 加權價值", (SE["A3"].value, SE["I3"].value)
    biz = []; rr = 5
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).strip().startswith("檢查"):
        biz.append(rr); rr += 1
    r_par = lab(SE, "目標價參數敏感度"); hr = r_par + 1
    assert SE.cell(hr, 1).value == "參數", SE.cell(hr, 1).value
    par = []; rr = hr + 1
    while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
        par.append(rr); rr += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；v{VER} 起，版面規格 v1.3 H 節；白話標籤，不回饋計算）。營運輸入案例＝本頁上方案例表（以情境 B 為底、只改一個輸入）："
                 "加權目標價＝基準加權目標價＋情境 B 權重 ×（案例的情境 B 加權價值 − 基準的情境 B 加權價值），即案例只套用在情境 B、情境 A／C 維持基準；"
                 "估值框架案例＝Val 頁封閉式；折現參數案例＝引用本頁「目標價參數敏感度」表").font = BOLD; r0 = r; r += 1
    hdr(SE, r, ["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "低案：加權目標價（美元）", "高案：加權目標價（美元）", "來源"])
    hdr_r = r; r += 1
    first = r
    frame = [
        ("情境權重 全押保守情境 C／全押樂觀情境 A（基準＝A 30%、B 50%、C 20%）",
         f"=tp_w_dcf*Val!$D${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$D${r_rv}/sh_v", f"=tp_w_dcf*Val!$B${r_dcfs}+(1-tp_w_dcf)*{pmed}*{coef}*Val!$B${r_rv}/sh_v", "Val 頁：單一情境的 DCF 價值與 FY2028 營收"),
        ("DCF 在加權中的權重 100%（只看現金流折現）／0%（只看同業比較法）（基準＝50%）", f"={dcf}", f"={cadj}", "Val 頁：DCF 與 Comparable（調整後）目標價"),
    ]
    for labl, lo, hi, note in frame:
        SE.cell(r, 1, labl); SE.cell(r, 4, lo).number_format = "#,##0.00"; SE.cell(r, 5, hi).number_format = "#,##0.00"
        SE.cell(r, 2, f"=D{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=E{r}/{tp}-1").number_format = "+0.0%;-0.0%"; SE.cell(r, 6, note); r += 1
    pairs = {}
    for sr in biz:
        t = str(SE.cell(sr, 1).value)
        base, side = t.rsplit("：", 1)
        pairs.setdefault(base, {})[side] = sr
    for base, d in pairs.items():
        SE.cell(r, 1, f"{base}（只套用情境 B）")
        for side, col, tcol in (("低", 2, 4), ("高", 3, 5)):
            if side in d:
                sr = d[side]
                SE.cell(r, tcol, f"={tp}+tp_wB*(Sens!I{sr}-Sens!$I$4)").number_format = "#,##0.00"
                SE.cell(r, col, f"={'DE'[col - 2]}{r}/{tp}-1").number_format = "+0.0%;-0.0%"
        SE.cell(r, 6, f"Sens!A{d.get('低', d.get('高'))}（本頁案例表；情境 B 加權價值）"); r += 1
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
    SE.column_dimensions["F"].width = max(SE.column_dimensions["F"].width or 0, 14)
    return dict(t0=r0, hdr=hdr_r, first=first, last=last)


# ---------- Coord：讀取層級、兩讀數判定、層級之間、C0（實績） ----------
def coord(wb, TKV, WATT):
    C = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    Q, V = wb["Q"], wb["Val"]
    C.column_dimensions["A"].width = 60
    for col, w in zip("BCDEFGHI", (16, 16, 14, 30, 40, 50, 30, 60)):
        C.column_dimensions[col].width = w
    C["A1"] = ("Coord — 三層協調（三層基準協調規則 v1.4.2；工作單 20261114 第二批第 2 節）：讀取層級、兩讀數（C4）判定、層級之間（參考）、C0（實績）。"
               "本模型是唯一不讀 Siliconomics 的硬體個股：量（平均計費 MW）與單價由公司合約與揭露推得，Tokenomics 只作 Recon 對照與一個敏感度高值。")
    C["A1"].font = BOLD
    r = 3
    C.cell(r, 1, "A. 讀取層級：本模型讀哪一層、哪一版、哪些名稱（自層級讀入的值只列在此節，不進 C0 判定）").font = BOLD; r += 1
    hdr(C, r, ["項目", "層", "版本", "本模型用在哪裡", "值／範圍"]); r += 1
    A1 = r
    rows = [
        ("Cerebras WSE-3 產出比、持有成本比（相對 VR200）；VR200 每 GW 持有成本與 IT 資本支出（In 頁 tk_* 列）", "Tokenomics", TKV,
         "Recon C 節（每 MW 年營收、每 MW 資本支出對照）；每 MW 資本支出敏感度高值", "In 頁 D 節；不進基準目標價"),
        ("Siliconomics（加速器顆數、份額、售價、對帳中旗標）", "Siliconomics", "v0.19（未讀）", "無：本模型不讀", "Siliconomics 無 Cerebras／晶圓級加速器的量或份額"),
        ("Wattonomics（玩家上線 GW、資本支出）", "Wattonomics", "v0.3.2（未讀）", "無：本模型不讀", WATT),
    ]
    for it in rows:
        for j, v in enumerate(it, 1):
            C.cell(r, j, v)
        r += 1
    C.cell(r, 1, "下游只註明版本、不自行重估（Instructions 第 3 條）。本模型的平均計費 MW、每 MW 年營收、每 MW 資本支出是本模型自己的假設，不寫回任何層級。").font = Font(italic=True); r += 2
    # ---- B：兩讀數（C4）判定 ----
    C.cell(r, 1, "B. 兩讀數（C4）判定：不適用——本模型不讀 Siliconomics 產業部署 GW，Siliconomics 對帳中年份（CY2027–CY2029）沒有傳導到本模型；另一側讀數（只換產業部署 GW）與現行讀數相同").font = BOLD; r += 1
    hdr(C, r, ["項目", "現行基準讀數", "另一側讀數（無傳導，同現行）", "另一側相對現行"]); r += 1
    B0 = r
    vb = lambda p: f"=Val!B{lab(V, p)}"
    for labl, ref, fmt in (("加權目標價（DCF 與 Comparable 加權）", vb("目標價（DCF 與 Comparable 加權）"), "#,##0.00"), ("DCF 目標價（情境加權）", vb("DCF 目標價（情境加權）"), "#,##0.00"),
                           ("Comparable 目標價（市場主流同業法；未調整）", vb("Comparable 目標價（市場主流同業法）"), "#,##0.00"), ("上檔空間（加權目標價 ÷ 現價 − 1）", vb("上檔空間：加權目標價"), "+0.0%;-0.0%")):
        C.cell(r, 1, labl); C.cell(r, 2, ref).number_format = fmt; C.cell(r, 3, f"=B{r}").number_format = fmt
        C.cell(r, 4, f"=C{r}-B{r}" if "上檔" in labl else f"=C{r}/B{r}-1").number_format = "+0.0%;-0.0%"; r += 1
    C.cell(r, 1, "投資建議（五級固定規則）"); C.cell(r, 2, "=" + REC5(f"B{r - 1}")); C.cell(r, 3, "=" + REC5(f"C{r - 1}")); B_rec = r; r += 1
    C.cell(r, 1, "是否有估值依賴的對帳中年份"); C.cell(r, 2, "否"); B_dep = r; r += 1
    C.cell(r, 1, "說明"); C.cell(r, 2, "C4 只適用估值依賴 Siliconomics 對帳中年份的硬體個股。本模型的營收量是平均計費 MW（公司合約與揭露推得），不是產業顆數；另一側讀數的唯一差異（產業部署 GW）"
                                 "在本模型沒有輸入格，整本重建結果與現行逐格相同，故另一側欄直接引用現行欄，不另建快照，也不建對帳中具名範圍；硬體個股清單（coord_meta）不需加入本公司。")
    B_note = r; r += 2
    # ---- C：層級之間（參考） ----
    C.cell(r, 1, "C. 層級之間（參考；不計入本公司 C0）").font = BOLD; r += 1
    hdr(C, r, ["項目", "值", "單位", "層、版本", "說明"]); r += 1
    C0r = r
    C.cell(r, 1, "Wattonomics 園區列：Cerebras（OpenAI），750 MW、2028 年上線、不納入加總"); C.cell(r, 2, 750); C.cell(r, 2).font = BLUE; C.cell(r, 3, "MW（IT）")
    C.cell(r, 4, "Wattonomics v0.3.2（Pipeline 列 cpx_neo_cer；營運者群 neo、使用者 oai、納入＝0）"); C.cell(r, 5, "不是 Cerebras 玩家值；Wattonomics 只把它列在敏感度，未計入任何玩家 GW。本模型 OpenAI 承諾 750 MW 讀公司揭露（In 頁 oai_mw），與此列同源（皆為公司公告），不比較")
    r += 1
    C.cell(r, 1, "Tokenomics 隱含 Cerebras 每 GW 年持有成本 vs 合約每 MW 年營收（Recon C 節第一列）"); C.cell(r, 2, f"=Recon!F{lab(wb['Recon'], '每 MW 年營收（FY2027）vs Tokenomics')}").number_format = "+0.0%;-0.0%"
    C.cell(r, 3, "相對差距"); C.cell(r, 4, TKV); C.cell(r, 5, "合約單價（Interested-party）對層級持有成本（Assumed），不是公司實績對層級，只列參考；既有 Recon 已說明待送算力收支")
    r += 2
    # ---- D：C0（實績） ----
    C.cell(r, 1, "D. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）").font = BOLD; D_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝Cerebras 8-K／10-Q 已公布的 2026 年季度實績（公司揭露，Verified）。層級側應為同口徑的層級隱含值，但三層都沒有 Cerebras 的量：Siliconomics 無晶圓級加速器顆數或份額；"
                 "Wattonomics 無 Cerebras 玩家（唯一相關的園區列不納入加總）；Tokenomics 只有每 GW 的產出與持有成本比，要換成營收或資本支出需要 Cerebras 在役（計費）MW，公司未揭露。"
                 "因此 C0 結論為「無可比實績」；下列實績只作記錄與未來對帳的起點。時差讀法照工作單：主讀數＝2026 年已公布期間累計 vs 層級 CY2026 × 月數 ÷ 12；第二讀數＝最近一季 × 4 vs 層級 CY2026"
                 "（最近一季 3Q26 中點約 2026-08，晚於年中，成長年第二讀數偏高）。v0.1.3 起滾入 3Q26（2026-11-12 公布），主讀數為 9M26（9 個月）。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); D_note = r; r += 2
    hdr(C, r, ["輸入", "值", "單位", "期間／口徑", "出處", "等級", "網址／名稱"]); r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000", url=""):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); C.cell(r, 7, url); IN[key] = r; r += 1
    assert Q.cell(3, 4).value == "1Q26" and Q.cell(3, 5).value == "2Q26" and Q.cell(3, 6).value == "3Q26"
    qr, qc, qh, qx = lab(Q, "營收合計（GAAP）"), lab(Q, "核心雲端營收"), lab(Q, "核心硬體營收"), lab(Q, "資本支出（季別未揭露")
    qx1 = lab(Q, "資本支出", lab(Q, "現金流量（FY"), exact=True)
    inp("r1", "營收合計（GAAP）：1Q26", f"=Q!D{qr}", "US$B", "2026-01～2026-03", f"Q 頁 D{qr}；Cerebras 8-K（1Q26 財報新聞稿）", "Verified", GREEN, url=S8K1)
    inp("r2", "營收合計（GAAP）：2Q26", f"=Q!E{qr}", "US$B", "2026-04～2026-06", f"Q 頁 E{qr}；Cerebras 8-K 2026-08-12（2Q26 財報新聞稿）", "Verified", GREEN, url=S8K2)
    inp("r3", "營收合計（GAAP）：3Q26", f"=Q!F{qr}", "US$B", "2026-07～2026-09", f"Q 頁 F{qr}；Cerebras 8-K 2026-11-12（3Q26 財報新聞稿）", "Verified", GREEN, url=S8K3)
    inp("c1", "核心雲端營收：1Q26", f"=Q!D{qc}", "US$B", "1Q26；公司核心口徑", f"Q 頁 D{qc}；Cerebras 8-K（1Q26 財報新聞稿）", "Verified", GREEN, url=S8K1)
    inp("c2", "核心雲端營收：2Q26", f"=Q!E{qc}", "US$B", "2Q26；公司核心口徑", f"Q 頁 E{qc}；Cerebras 8-K 2026-08-12", "Verified", GREEN, url=S8K2)
    inp("c3", "核心雲端營收：3Q26", f"=Q!F{qc}", "US$B", "3Q26；公司核心口徑", f"Q 頁 F{qc}；Cerebras 8-K 2026-11-12", "Verified", GREEN, url=S8K3)
    inp("h1", "核心硬體營收：1Q26–3Q26", f"=Q!D{qh}+Q!E{qh}+Q!F{qh}", "US$B", "9M26；公司核心口徑", f"Q 頁 D{qh}＋E{qh}＋F{qh}", "Verified", GREEN, url=S8K3)
    inp("x", "資本支出：9M26", f"=Q!F{qx1}", "US$B", "2026-01～2026-09（現金流量表）", f"Q 頁 F{qx1}；8-K 2026-11-12（9M26 現金流量表）", "Verified", GREEN, url=S8K3)
    inp("mo", "已公布月數（2026 1Q–3Q）", 9, "月", "2026-01～2026-09", "會計年度＝日曆年", "Derived", BLUE, fmt="0")
    r += 1
    hdr(C, r, ["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"]); r += 1
    D0 = r
    lines = [
        ("C0（實績）判定：雲端營收 CY2026（主讀數 9M26 累計；第二讀數 3Q26 × 4）", f"=B{IN['c1']}+B{IN['c2']}+B{IN['c3']}", f"B{IN['c3']}*4",
         "缺：Cerebras 在役／計費 MW 揭露（Tokenomics 每 GW 收支換成營收所需）；Wattonomics 無 Cerebras 玩家、Siliconomics 無 Cerebras 量"),
        ("C0（實績）判定：營收合計（GAAP）CY2026", f"=B{IN['r1']}+B{IN['r2']}+B{IN['r3']}", f"B{IN['r3']}*4",
         "缺：Siliconomics 晶圓級加速器（或「其他加速器」中 Cerebras）的顆數 × 售價；層級沒有 Cerebras 份額"),
        ("C0（實績）判定：資本支出 CY2026", f"=B{IN['x']}", None,
         "缺：Wattonomics 的 Cerebras 玩家資本支出（無此玩家；唯一相關園區列不納入）"),
    ]
    for labl, act, act2, miss in lines:
        C.cell(r, 1, labl).font = BOLD
        C.cell(r, 2, act).number_format = "#,##0.000"
        C.cell(r, 3, "—"); C.cell(r, 4, "—")
        C.cell(r, 5, "無可比實績").font = BOLD
        C.cell(r, 6, "US$B；公司實績欄＝2026 已公布期間累計（" + ("第二讀數＝最近一季 × 4，見說明" if act2 else "9M26，現金流量表") + "）")
        C.cell(r, 7, "公司：Cerebras 8-K／10-Q（Q 頁）；層級：Siliconomics v0.20、Wattonomics v0.3.6、" + TKV.split("（")[0] + "（皆無 Cerebras 同口徑量）")
        C.cell(r, 8, "Verified（公司）")
        C.cell(r, 9, F(T(miss + "。不以模型預測或合約推算值補位"), *( [T("；第二讀數（最近一季 × 4）＝"), f'TEXT({act2},"0.000")'] if act2 else []), T("。")))
        r += 1
    D1 = r - 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（D 節判定列；A 節讀取層級與 C 節參考不計）"); C.cell(r, 2, f'=COUNTIF(E{D0}:E{D1},"對帳中")'); cnt = r; r += 1
    C.cell(r, 1, "說明：D 節三列皆為「無可比實績」（層級無 Cerebras 的量）。Recon C 節「合約每 MW 年營收 vs Tokenomics 持有成本」是合約與層級假設的比較、不是公司實績，只列 C 節參考，不計入對帳中；"
                 "公司日後揭露在役或計費 MW（或 Wattonomics 新增 Cerebras 玩家）時，改以「雲端營收實績 vs Tokenomics 每 GW 營收 × 在役 GW」判定。")
    note = r; r += 1
    return dict(A1=A1, B0=B0, B_rec=B_rec, B_dep=B_dep, B_note=B_note, C0r=C0r, D=D_title, Dn=D_note, D0=D0, D1=D1, cnt=cnt, note=note, IN=IN)


# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, PR, RVA, PLA, SEN, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; RC = wb["Recon"]
    vb = lambda p: f"Val!B{lab(V, p)}"
    tp, up, rec = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    dcf, cadj, pmed = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("同業前瞻 P/S 中位數")
    B = PR["B"]
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    t0 = SEN["first"]
    SB = lambda lbl, col: f"Sens!{col}{lab(SE, lbl, t0)}"
    rx27 = lab(RC, "FY2027 營收 ÷ FY2026")
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（HTML 用；v{VER} 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("現金流折現 "), f'TEXT({dcf},"0")', T(" 美元遠低於現價 "), 'TEXT(px,"0")', T(" 美元（擴建期自由現金流為負）；同業營收倍數法 "), f'TEXT({cadj},"0")',
        T(" 美元。兩者各半加權 "), f'TEXT({tp},"0")', T(" 美元，評等「"), rec, T("」。關鍵在 OpenAI 合約每 MW 年營收與計費 MW 的擴張速度。"))
    put("rec", T("評等「"), rec, T("」的前提：兩種方法各半。只看現金流折現時評等為"), REC5(f"({dcf}/px-1)"), T("（"), f'TEXT({dcf}/px-1,"+0%;-0%")',
        T("），只看同業營收倍數法時為"), REC5(f"({cadj}/px-1)"), T("（"), f'TEXT({cadj}/px-1,"+0%;-0%")', T("）；兩法分歧大，加權結果落在中立區間。"))
    put("o1", T("目標價：加權目標價 "), f'TEXT({tp},"0.00")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等「"), rec,
        T("」。現金流折現 "), f'TEXT({dcf},"0.00")', T(" 美元（折現率 "), 'TEXT(wacc,"0.0%")', T("），同業比較法 "), f'TEXT({cadj},"0.00")',
        T(" 美元＝同業前瞻營收倍數中位數 "), f'TEXT({pmed},"0.0")', T(" 倍 × FY2028 營收（情境加權）÷ 股數。"))
    put("o2", T("現價隱含的條件：要讓現金流折現法支持現價，"), f'IF(ABS({rv("w_res")})<0.01,"折現率須降到 "&TEXT({rv("w1")},"0.0%"),"折現率即使降到 "&TEXT({rv("w1")},"0.0%")&" 仍不足")', T("（模型 "), 'TEXT(wacc,"0.0%")', T("），或 FY2028 起營收全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT({rv("rv1")},"0.0")', T("，模型 "), f'TEXT({rv("rv0")},"0.0")', T(" 十億美元）。"))
    put("o3", T("成長：情境 B 的 FY2027 營收 "), f'TEXT(PL!E{B["rev"]},"0.00")', T(" 十億美元，為 FY2026 的 "), f'TEXT(Recon!C{rx27},"0.0")', T(" 倍（公司說三倍）；FY2028 平均計費 "),
        f'TEXT(PL!F{B["mw"]},"#,##0")', T(" MW、營收 "), f'TEXT(PL!F{B["rev"]},"0.00")', T("，營業利益率 "), f'TEXT(PL!F{B["om"]},"0.0%")', T("（含股票薪酬）。"))
    put("o4", T("資本支出：情境 B FY2027–FY2028 資本支出合計 "), f'TEXT(PL!E{B["cx"]}+PL!F{B["cx"]},"0.0")', T(" 十億美元，對照 2026-09-30 現金與投資 "), 'TEXT(Q!F35+Q!F36+Q!F37,"0.0")',
        T("；每 MW 資本支出改用 Tokenomics 隱含值時，加權目標價 "), f'TEXT({SB("每 MW 資本支出", "C")},"+0.0%;-0.0%")', T("（只套用情境 B）。"))
    put("o5", T("敏感度：只看現金流折現時加權目標價 "), f'TEXT({SB("DCF 在加權中的權重", "B")},"+0%;-0%")', T("、只看同業比較法 "), f'TEXT({SB("DCF 在加權中的權重", "C")},"+0%;-0%")',
        T("；情境全押保守 "), f'TEXT({SB("情境權重", "B")},"+0%;-0%")', T("、全押樂觀 "), f'TEXT({SB("情境權重", "C")},"+0%;-0%")', T("。營運輸入中每 MW 年營收（合約期五年）"),
        f'TEXT({SB("每 MW 年營收", "B")},"+0.0%;-0.0%")', T("、平均計費 MW 減兩成 "), f'TEXT({SB("平均計費 MW", "B")},"+0.0%;-0.0%")', T("。"))
    put("drv", T("分部營收推導（情境 B）：推論雲端營收＝平均計費 MW × 每 MW 年營收（OpenAI 合約推得，FY2028 起每年調整）；系統（硬體）營收＝FY2026 錨點 × 情境年增。FY2028：雲端 "),
        f'TEXT(PL!F{B["cl"]},"0.00")', T("（"), f'TEXT(PL!F{B["mw"]},"#,##0")', T(" MW × "), f'TEXT(PL!F{B["pmw"]},"0.00")', T(" 百萬美元）、硬體 "), f'TEXT(PL!F{B["hw"]},"0.00")',
        T("，合計 "), f'TEXT(PL!F{B["rev"]},"0.00")', T(" 十億美元。"))
    put("coord", T("三層協調：本模型只讀 Tokenomics（Cerebras 產出比與持有成本比，只作對照與敏感度高值），不讀 Siliconomics 與 Wattonomics；Siliconomics 對帳中年份沒有傳導到本模型，"
                   "兩讀數不適用（另一側讀數與現行相同）。C0 以公司實績對層級：2026 前三季雲端營收 "), f'TEXT(Coord!B{CO["D0"]},"0.000")', T("、營收合計 "), f'TEXT(Coord!B{CO["D0"] + 1},"0.000")',
        T(" 十億美元，但三層都沒有 Cerebras 的量，結論「無可比實績」，對帳中項目 "), f'TEXT(Coord!B{CO["cnt"]},"0")', T(" 項。"))
    RD.cell(r, 1, f"v{CUR}（{BUILT}）：滾入 3Q26（2026-11-12 財報）——Q 頁 3Q26 改為實績、4Q26 改為 4Q26 指引中點（FY2026 錨點＝1Q26–3Q26 實績＋4Q26 指引）；"
                  "資產負債表日、淨現金、股數改為 2026-09-30（股數 2026-11-05）；stub 期改為 4Q26；RPO、容量、公司 2027 說法更新；C0 主讀數改為 9M26。情境假設與公式不變。｜"
                  f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3 與三層協調（工作單 20261114 第二批第 2 節）——章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；"
                  "PL 頁末新增標準損益表（FY2024A–FY2025A GAAP 實績＋FY2026E–FY2032E 情境 B）、RevDCF 頁（反向 DCF）、Val 頁 0.7（四個目標價）、Sens 頁末加權目標價相對基準、"
                  "Coord 頁（讀取層級、兩讀數判定：不適用、C0 實績：無可比實績）；既有數值格不變（新增列、新增頁與版本字串以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    for k, rr in out.items():
        RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return out


def version(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and f"模型 v{CUR}（" in a1, a1[:120]
    RD["A1"].value = a1.replace("2026-11-08", BUILT, 1) + f"；v{VER}：個股 HTML 版面規格 v1.3 與三層協調（Coord 頁）；v{CUR}：滾入 3Q26 實績與 4Q26 指引"


def apply(wb, PR, TKV):
    WATT = "Wattonomics 無 Cerebras 玩家；Pipeline 園區列 Cerebras（OpenAI）750 MW 不納入加總（營運者群 neo）"
    PLA = pl_standard(wb, PR)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SEN = sens_tor(wb)
    CO = coord(wb, TKV, WATT)
    RMA = readme13(wb, PR, RVA, PLA, SEN, CO)
    version(wb)
    return (f"v012_post：PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準 {SEN['t0']} 起；"
            f"Coord 頁（C0 第 {CO['D0']} 列起）；README v1.3 白話列 {len(RMA)} 列")
