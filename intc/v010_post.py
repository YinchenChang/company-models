# 英特爾 v0.1.0（2026-11-17）後處理：個股 HTML 版面規格 v1.3（樣板 台積電 v0.5.5、聯電 v0.2.4）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 本檔改寫自 umc/v024_post.py（RevDCF、Val 0.7、Sens 相對基準三段共用邏輯不變，只換幣別與公司字樣；Hist、標準損益表、分部推導、C0、README v1.3 為英特爾專屬）。
#   Hist 頁（新）：合併損益表 FY2022–FY2025（公司 4Q22／4Q23／4Q24／4Q25 財報新聞稿，美國 GAAP，十億美元）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，十億美元）、比率區、核對列、分部營收推導（情境 B）
#   RevDCF 頁（新）：反向 DCF
#   Val 頁末：0.7 估值並列四個目標價；Sens 頁末：加權目標價相對基準（連本頁既有逐案整本重算結果）
#   Coord 頁末：三層協調——讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）
#   Sources 頁：新增來源；README 頁末：v1.3 展示層白話列與本版說明
import os, re, json
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.1.1", "2026-11-21"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
IR = "https://www.intc.com/news-events/press-releases/detail/"
Q325 = IR + "1759/intel-reports-fourth-quarter-and-full-year-2025-financial-results"
Q424 = IR + "1726/intel-reports-fourth-quarter-and-full-year-2024-financial-results"
Q423 = "https://www.sec.gov/Archives/edgar/data/50863/000005086324000008/q423_earningsrelease.htm"
SRC_H25 = f"FY2025：公司 4Q25 與 FY2025 財報新聞稿（2026-01-22，{Q325}）合併損益表，百萬美元 ÷ 1,000；與 Q 頁 FY2025A 錨點同源"
SRC_H23 = (f"FY2024、FY2023：公司 4Q24 與 FY2024 財報新聞稿（2025-01-30，{Q424}）合併損益表十二個月欄；"
           f"FY2022：公司 4Q23 與 FY2023 財報新聞稿（2024-01-25，8-K 附件 {Q423}）十二個月比較欄；皆美國 GAAP，百萬美元 ÷ 1,000")

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))


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
TX = lambda ref, f: f'TEXT({ref},"{f}")'
SP = "+0.0%;-0.0%"


# ---------- RevDCF：反向 DCF（樣板 台積電 v0.5.5；英特爾 FY2027 不動、FY2028 起反解） ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 6, 7, 8
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.1.0 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（第一個明確期年度）不動，營業利益率與現金流倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("px", "現價（美元）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；十億美元）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節；同 DCF 頁每股價值公式）")
    kv("sh", "評價用股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（十億美元）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；十億美元）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${ps}", "#,##0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}
    NC = "BCDEFG"

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（十億美元）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（十億美元）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（十億美元）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（十億美元）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（十億美元，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（十億美元）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（十億美元）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(C{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2027 現值（十億美元）", f"=B{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（十億美元）", f"=SUM(C{Y['pv']}:G{Y['pv']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（十億美元）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；十億美元）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近邊界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tn}+$B${A['nc']})/sh_v")
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
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（美元，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
    r += 1
    RV.cell(r, 1, "B. Comparable：支持現價所需的倍數").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = SP
        A[key] = r; r += 1
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；美元）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0.0"),
                     T(" 美元；要支持現價 "), TX("px", "#,##0.0"), T(" 美元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大","縮小")', T("為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<wacc,"降到 ","升到 ")', TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面乘以 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億美元，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率改為 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0.0"), T(" 美元相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 英特爾過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給英特爾的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("歷史相對倍數 "), TX(coef, "0.00"), T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    return A


# ---------- Val：估值並列四個目標價（版面規格 v1.3 G 節；只新增列） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.1.0 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "美元", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.00"), T(" 美元；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "In 頁市場共識目標價（stockanalysis 分析師平均；Interested-party）")]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = SP; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=In!$B$" + str(lab(wb["In"], "px", exact=True))); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（連 Sens 頁既有逐案整本重算結果） ----------
SENS_LAB = {"ai_adj": "DCAI 中 AI 相關營收 ±30%", "ext_adj": "外部代工營收 ±50%", "fm_adj": "Foundry 營業利益率 ±5 個百分點", "dcm_adj": "DCAI 營業利益率 ±3 個百分點",
            "ccm_adj": "CCPG 營業利益率 ±3 個百分點", "ccg_adj": "CCPG 營收年增 ±3 個百分點", "cx_adj": "總資本支出 ±10%", "opx_adj": "營業費用 ±5%",
            "nonop_f": "營業外淨額", "tax": "有效稅率 10%／20%", "nci_r": "非控制權益報酬率 3%／7%",
            "cu_f": "Corporate Unallocated −4.8／−3.2 十億美元", "da_g": "折舊年增 2%／6%"}
FD_PAIRS = [("折現率 WACC −1／＋1 個百分點", "WACC −1pt", "WACC ＋1pt"), ("β（產業 β ∓ 0.3 重新加槓桿）", "β：低", "β：高"), ("股權風險溢酬 4%／6%", "ERP：低", "ERP：高"),
            ("永續成長率", "永續成長率 g：低", "永續成長率 g：高"), ("永續期新投資報酬率＝WACC＋5 個百分點（只有高案）", None, "RONIC＝WACC＋5pt"),
            ("資本支出占營收 −1／＋1 個百分點", "資本支出 ÷ 營收 −1pt", "資本支出 ÷ 營收 ＋1pt"), ("營運資金占營收", "營運資金 ÷ 營收：低", "營運資金 ÷ 營收：高"),
            ("淨現金不含權益投資（只有低案）", "淨現金不含權益投資", None), ("終值改用自由現金流成長法（只有低案）", "終值改用 FCFF ×（1＋g）", None),
            ("DCF 權重 0.3／0.7", "DCF 權重 0.3", "DCF 權重 0.7")]


def sens_tor(wb, G):
    SE = wb["Sens"]
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.1.0 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務案例連本頁上方逐案整本重算結果（相對基準欄），目標價參數取上方 DCF 參數案例；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    for key, *_ in G["C"]["sens_cases"]:
        lo, hi = G["SENS_ROWS"][(key, "lo")], G["SENS_ROWS"][(key, "hi")]
        SE.cell(r, 1, SENS_LAB[key]); SE.cell(r, 2, f"=J{lo}").number_format = SP; SE.cell(r, 3, f"=J{hi}").number_format = SP
        SE.cell(r, 4, f'="低案 "&B{lo}&"；高案 "&B{hi}&"（Sens 頁第 {lo}、{hi} 列）"'); r += 1
    TP = G["TP_ROWS"]
    for labl, lo, hi in FD_PAIRS:
        SE.cell(r, 1, labl)
        if lo: SE.cell(r, 2, f"=J{TP[lo]}").number_format = SP
        if hi: SE.cell(r, 3, f"=J{TP[hi]}").number_format = SP
        SE.cell(r, 4, "Sens 頁目標價參數案例：" + "／".join(x for x in (lo, hi) if x)); r += 1
    wr = lab(SE, "Wattonomics 錨定", exact=True)
    SE.cell(r, 1, "AI 相關營收依 Wattonomics 錨定（只有低案）"); SE.cell(r, 2, f"=J{wr}").number_format = SP
    SE.cell(r, 4, "Sens 頁 Wattonomics 錨定案例（既有列；DCAI 中 AI 相關營收 FY2027 起 × 錨定係數，三情境同乘）"); r += 1
    last = r - 1
    SE.cell(r, 1, "核對：基準列相對基準（應為 0）"); SE.cell(r, 2, "=J4").number_format = "0.000000"
    return dict(t0=r0, first=first, last=last, state=r)


# ---------- Hist：合併損益表歷史（十億美元） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)
HIS = [
    ("rev", "淨營收（Net revenue）", 63.054, 54.228, 53.101, 52.853),
    ("cogs", "銷貨成本（Cost of sales）", 36.188, 32.517, 35.756, 34.478),
    ("gp", "毛利（Gross margin）", 26.866, 21.711, 17.345, 18.375),
    ("oi", "營業利益（Operating income；已扣營業費用含重組）", 2.334, 0.093, -11.678, -2.214),
    ("pti", "稅前利益（Income before taxes）", 7.768, 0.762, -11.210, 1.557),
    ("tax", "所得稅費用（Provision for taxes；負數＝利益）", -0.249, -0.913, 8.023, 1.531),
    ("ni", "本期淨利（Net income；含非控制權益）", 8.017, 1.675, -19.233, 0.026),
    ("nip", "歸屬 Intel 淨利（模型淨利口徑）", 8.014, 1.689, -18.756, -0.267),
]
EPSR = [("eps", "EPS 稀釋（公司公告；美元）", (1.94, 0.40, -4.38, -0.06)),
        ("sh", "稀釋加權平均股數（十億股）", (4.123, 4.212, 4.280, 4.530))]
STD = ("美國 GAAP（4Q23 新聞稿）", "美國 GAAP（4Q24 新聞稿）", "美國 GAAP（4Q24 新聞稿）", "美國 GAAP（4Q25 新聞稿）")


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 合併損益表歷史（v0.1.0 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度近似日曆年（Intel 會計年度為 12 月最後一個週六）。金額：十億美元（美國 GAAP 合併，公司財報新聞稿）；EPS：美元（稀釋）。"
               "營業利益為公司 GAAP 營業利益（已含重組與其他費用），標準損益表的營業費用＝毛利 − 營業利益")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    H.cell(r, 1, "h_std"); H.cell(r, 2, "會計準則（出處）")
    for j, v in enumerate(STD): H.cell(r, 3 + j, v)
    H.cell(r, 7, "Verified"); H.cell(r, 8, "美國 GAAP；Intel 另公布 Non-GAAP（調整股權報酬、收購相關與重組等），模型與本表一律 GAAP，共識 EPS 多為 Non-GAAP，兩者不同")
    R["std"] = r; r += 1
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H25 + "；" + SRC_H23)
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k == "eps" else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, ("同上損益表加權平均股數（百萬股 ÷ 1,000；FY2025 取 Q 頁錨點 4.53）" if k == "sh" else "同上（每股盈餘；FY2025 公司公告 −0.06）"))
        R[k] = r; r += 1
    for k, labl, f_, note in (("nonop", "營業外淨額（＝稅前 − 營業利益；權益投資損益、利息及其他）", "={c}{pti}-{c}{oi}", "FY2022 含權益投資收益 4.268（Mobileye 等），FY2025 含利息及其他 3.3"),
                              ("opx", "營業費用（含重組；＝毛利 − 營業利益）", "={c}{gp}-{c}{oi}", "FY2024 含重組及其他費用 6.970"),
                              ("nci", "非控制權益損益（＝本期淨利 − 歸屬 Intel 淨利）", "={c}{ni}-{c}{nip}", "—")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "#,##0.000"
        H.cell(r, 7, "Derived"); H.cell(r, 8, note); R[k] = r; r += 1
    r += 1
    H.cell(r, 1, "核對（十億美元；千元換算可有 ±0.001 捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("營收 − 銷貨成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"), ("稅前利益 − 所得稅 − 本期淨利", "={c}{pti}-{c}{tax}-{c}{ni}"),
                     ("歸屬 Intel 淨利 ÷ 稀釋股數 − 稀釋 EPS（EPS 四捨五入到分）", "={c}{nip}/{c}{sh}-{c}{eps}")):
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "FY2025：本表 − Q 頁 FY2025A 錨點（營收、毛利、營業利益、歸屬 Intel 淨利）")
    for c, k, nm in (("C", "rev", "a25_rev"), ("D", "gp", "a25_gp"), ("E", "oi", "a25_oi"), ("F", "nip", "a25_ni")):
        H[f"{c}{r}"] = f"=F{R[k]}-{nm}"; H[f"{c}{r}"].number_format = "0.000"
    R["chk1"] = r; r += 2
    H.cell(r, 1, "口徑註記：四年皆取公司財報新聞稿的合併損益表十二個月欄（美國 GAAP）。Yahoo Finance 快照的 FY2022–FY2025 營業利益與稅前數字與公司新聞稿不同（例：FY2025 營業利益 Yahoo −23 百萬、公司 −2,214 百萬；差異來自 Yahoo 將重組與其他費用另列於營業利益之外），以公司原文為準。"
                  "FY2023、FY2024 的 EPS 與股數取公司公告，不重算。")
    R["note"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 60
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表（十億美元） ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "D", "D": "E", "E": "F"}                       # 標準損益表歷史欄 → Hist 欄
EC = {"F": "D", "G": "E", "H": "F", "I": "G", "J": "H", "K": "I", "L": "J"}   # 預測欄 → PL 頁推導區塊欄（FY2026–FY2032）
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}
DCOL = {"J": "F", "K": "G", "L": "H"}                     # FY2030–32 → DCF 頁欄


def pl_standard(wb, HR, CR, B):
    P = wb["PL"]
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr); d_rev = lab(D, "營收", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.1.0 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加 4Q26 指引的錨點）。十億美元；EPS 美元（稀釋）。"
                 "只引用本頁上方情境 B 推導區塊、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（美國 GAAP 合併財報；歷史三年與 FY2022 基期皆取公司財報新聞稿）。Intel 另公布 Non-GAAP（排除股權報酬、收購相關、重組等），市場共識 EPS 多為 Non-GAAP 口徑，本表與模型一律 GAAP，不做 Non-GAAP 調整。"
                 "稅後淨利＝歸屬 Intel 淨利（模型淨利口徑）；EPS 為稀釋 EPS（歷史為公司公告值）。營業費用含重組（＝毛利 − 營業利益）；其他收入（淨額）＝稅前利益 − 營業利益，FY2026E 含 Escrowed Shares 評價損失 126 億美元（非現金，FY2027 起不預測）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝（淨利 ＋ 非控制權益）÷（1 − 稅率）− 營業外淨額，由淨利還原）；毛利率沿用前一年 據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列，不改淨利與目標價。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":
        yr_ = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr_}"&IF(Coord!${"BCDE"[FLAGY[col] - 2026]}${CR["flag"]}=1,"（對帳中）","")'
    r_hdr = r; r += 1
    order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "nci", "ni", "nm", "sh", "eps"]
    RW = {k: r + i for i, k in enumerate(order)}
    M = "#,##0.0"; PCT = "0.0%"
    g = RW

    def row(key, label, unit, cells, fmt, note=""):
        P.cell(RW[key], 1, label); P.cell(RW[key], 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{RW[key]}"] = v; P[f"{col}{RW[key]}"].number_format = fmt
        if note: P.cell(RW[key], 13, note)
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    eb = lambda k, cols="FGHI": [f"=PL!{EC[c]}{B[k]}" for c in cols]
    row("rev", "營收", "US$B", hh("rev") + eb("rev", "FGHIJKL"), M, "FY2026E–FY2029E＝PL 頁情境 B 推導區塊（CCPG＋DCAI＋All Other＋外部代工）；FY2030E 起為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{g['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{g['rev']}/{p}{g['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 取 Hist 頁")
    row("cogs", "營業成本（COGS）", "US$B", hh("cogs") + [f"={c}{g['rev']}-{c}{g['gp']}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "US$B", hh("gp") + eb("gp") + [f"={c}{g['rev']}*{c}{g['gm']}" for c in "JKL"], M, "FY2026E–FY2029E＝PL 頁毛利（＝營業利益＋營業費用）；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{g['gp']}/{c}{g['rev']}" for c in "CDEFGHI"] + [f"={p}{g['gm']}" for p in "IJK"], PCT, "FY2030E–FY2032E＝前一年")
    row("opx", "營業費用（含重組）", "US$B", [f"=Hist!{HC[c]}{HR['opx']}" for c in "CDE"] + eb("opx")
        + [None] * 3, M, "歷史＝毛利 − 營業利益；FY2026E–FY2029E＝PL 頁營業費用；FY2030E 起＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）")
    row("oi", "營業利益", "US$B", hh("oi") + eb("oi") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M, "FY2030E–FY2032E＝DCF 頁情境 B EBIT（淨利還原）")
    row("opm", "營業利益率", "%", [f"={c}{g['oi']}/{c}{g['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外）", "US$B", [f"={c}{g['pti']}-{c}{g['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（權益投資損益、利息及其他）；FY2026E 含 Escrowed Shares 評價損失（非現金）")
    row("pti", "稅前利益", "US$B", hh("pti") + eb("pti") + [f"={c}{g['oi']}+nonop_f" for c in "JKL"], M, "FY2030E 起＝延伸期營業利益＋營業外淨額（In 頁，與 DCF 還原同式）")
    row("tax", "所得稅", "US$B", hh("tax") + eb("tax") + [f"={c}{g['pti']}*tax" for c in "JKL"], M, "FY2030E 起＝稅前 × 有效稅率（In 頁）")
    row("nci", "非控制權益損益", "US$B", [f"=Hist!{HC[c]}{HR['nci']}" for c in "CDE"] + eb("mi") + [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['pti'] + 3}" for c in "JKL"], M,
        "FY2027E 起＝期初非控制權益餘額 × 報酬率（In 頁）；FY2030E 起＝稅前 − 所得稅 − 稅後淨利")
    row("ni", "稅後淨利（歸屬 Intel）", "US$B", hh("nip") + eb("ni", "FGHIJKL"), M, "預測期＝PL 頁情境 B 歸屬 Intel 淨利")
    row("nm", "淨利率", "%", [f"={c}{g['ni']}/{c}{g['rev']}" for c in COLS], PCT)
    row("sh", "股數（十億股；稀釋加權平均）", "十億股", hh("sh") + eb("shr", "FGHIJKL"), "0.000")
    row("eps", "EPS（美元；稀釋，歷史為公司公告值）", "US$", hh("eps") + eb("eps", "FGHIJKL"), "0.00", "預測期＝PL 頁 EPS（＝歸屬 Intel 淨利 ÷ 稀釋股數）")
    r = RW["eps"] + 1
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    L = dict(RW)

    def row2(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, 13, note)
        L[key] = r; r += 1
    row2("r_yoy", "營收年增率", "%", [f"={c}{g['yoy']}" for c in COLS], PCT)
    row2("r_gm", "毛利率", "%", [f"={c}{g['gm']}" for c in COLS], PCT)
    row2("r_opm", "營業利益率", "%", [f"={c}{g['opm']}" for c in COLS], PCT)
    row2("r_nm", "淨利率", "%", [f"={c}{g['nm']}" for c in COLS], PCT)
    row2("r_niy", "淨利年增率（YoY）", "%", [f'=IF(AND(Hist!C{HR["nip"]}>0,C{g["ni"]}>0),C{g["ni"]}/Hist!C{HR["nip"]}-1,"n/a")'] + [f'=IF(AND({p}{g["ni"]}>0,{c}{g["ni"]}>0),{c}{g["ni"]}/{p}{g["ni"]}-1,"n/a")' for p, c in zip(COLS, COLS[1:])], PCT,
         "基期或當期淨利為負或零時標 n/a（FY2022 基期＝Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row2("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{g['rev']}-{c}{g['cogs']}-{c}{g['gp']}" for c in COLS], "0.000")
    row2("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2032E 見下一列）", "US$B", [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "CDEFGHI"] + [None] * 3, "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E–FY2032E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "JKL":
        P[f"{c}{g['opx']}"] = f"={c}{g['gp']}-{c}{g['oi']}"; P[f"{c}{g['opx']}"].number_format = P[f"I{g['opx']}"].number_format
    row2("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E–FY2032E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "US$B", [None] * 7 + [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "JKL"], "0.000")
    row2("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E–FY2032E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 6 + [f"={c}{g['opx']}/{c}{g['rev']}" for c in "IJKL"], "0.00%")
    row2("c3", "核對：稅前利益 − 所得稅 − 非控制權益 − 稅後淨利", "US$B", [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['nci']}-{c}{g['ni']}" for c in COLS], "0.000")
    row2("c4", "核對：稅後淨利 ÷ 股數 − EPS（預測期）", "US$", [None] * 3 + [f"={c}{g['ni']}/{c}{g['sh']}-{c}{g['eps']}" for c in "FGHIJKL"], "0.0000")
    row2("c5", "核對：營收 − DCF 頁情境 B 營收（FY2027E–FY2032E）／ − Q 頁 FY2026 錨點（FY2026E）", "US$B", [None] * 3 + [f"=F{g['rev']}-a26_rev"] + [f"={c}{g['rev']}-DCF!{d}{d_rev}" for c, d in zip("GHIJKL", "CDEFGH")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度近似日曆年，FY(t) 直接對應 CY(t)；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2032E 由 FY2029 之後的收斂路徑推得，沿用 CY2029 旗標。FY2026 不是對帳中年份。")
    L["ycaveat"] = r; r += 2
    # ---- 分部營收推導（情境 B；HTML 命題、驅動因子章用） ----
    P.cell(r, 1, "分部營收推導（情境 B；v0.1.0）：營收＝CCPG（客戶端運算與實體 AI）＋DCAI（資料中心與 AI；拆 AI 相關與其他，AI 相關 FY2028 起乘 Siliconomics 顆數係數）＋All Other＋外部代工營收（Foundry 營收扣除對 Intel Products 的內部銷售）；Foundry 營業利益率為假設路徑。只引用本頁上方情境 B 推導區塊").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "", "", ""] + ["FY2026E", "FY2027E", "FY2028E", "FY2029E"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    L["seg_hdr"] = r; r += 1

    def seg(key, label, unit, fn, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for c in "FGHI":
            P[f"{c}{r}"] = fn(c); P[f"{c}{r}"].number_format = fmt
        if note: P.cell(r, 13, note)
        L[key] = r; r += 1
    pb = lambda k: (lambda c: f"=PL!{EC[c]}{B[k]}")
    seg("s_ccg", "CCPG 營收", "US$B", pb("ccg"), M, "FY2026 為錨點；FY2027 起依年增假設")
    seg("s_ccm", "CCPG 營業利益率", "%", pb("ccm"), PCT)
    seg("s_dai", "DCAI 中 AI 相關營收", "US$B", pb("dai"), M, "FY2026 占 DCAI 35%（假設拆分）；FY2028 起乘 Siliconomics 顆數係數")
    seg("s_dnai", "DCAI 其他（非 AI 相關）營收", "US$B", pb("dnai"), M)
    seg("s_dcr", "DCAI 營收", "US$B", pb("dcr"), M)
    seg("s_dcm", "DCAI 營業利益率", "%", pb("dcm"), PCT)
    seg("s_oth", "All Other 營收", "US$B", pb("oth"), M)
    seg("s_ext", "外部代工營收", "US$B", pb("ext"), M, "公司揭露 Foundry 營收含對內銷售，扣除分部間沖銷後為外部代工營收（18A／14A、先進封裝）")
    seg("s_fnd", "Intel Foundry 營收（含對內銷售）", "US$B", pb("fnd"), M)
    seg("s_fm", "Intel Foundry 營業利益率", "%", pb("fm"), PCT, "占 Foundry 營收（含對內銷售）；假設路徑")
    seg("s_fndoi", "Intel Foundry 營業利益（虧損）", "US$B", pb("fndoi"), M)
    seg("s_cx", "總資本支出", "US$B", pb("cx"), M, "投資活動與融資活動的 PP&E 增加合計（公司口徑，未扣政府補助與夥伴挹注）")
    seg("s_cxs", "總資本支出 ÷ 營收", "%", pb("cxs"), PCT)
    seg("s_rev", "營收", "US$B", pb("rev"), M)
    seg("s_chk", "核對：CCPG＋DCAI＋All Other＋外部代工 − 營收（應為 0）", "US$B", lambda c: f"={c}{L['s_ccg']}+{c}{L['s_dcr']}+{c}{L['s_oth']}+{c}{L['s_ext']}-{c}{L['s_rev']}", "0.000")
    seg("s_sh", "AI 相關 ÷ 營收", "%", lambda c: f"={c}{L['s_dai']}/{c}{L['s_rev']}", PCT)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0)
    return L


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, B):
    CW = wb["Coord"]; N = SIL.NAMES
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.1.0；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（顆數情境係數 SIF_UnitsCoef_Base；Wattonomics 錨定 SIF_WTAnchor_Base；代工與邏輯設備成長 eq_g_fd；對帳中旗標與另一側比值）", SIL.VERSION,
           "英特爾屬硬體供應鏈個股，讀 Siliconomics；不讀 Wattonomics 覆蓋 Siliconomics（C4）"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁 Wattonomics 錨定案例；不改基準"),
          ("顆數情境係數 情境 B（CY2028）", "=um28_B", "In 頁 um28_B（Siliconomics SIF_UnitsCoef_Base；情境 A、C 讀 markets/scen）", SIL.VERSION, "DCAI 中 AI 相關營收 FY2028 起乘同字母係數"),
          ("顆數情境係數 情境 B（CY2029）", "=um29_B", "In 頁 um29_B（同上）", SIL.VERSION, "同上"),
          ("公司自有變數（不由層級提供）", "—", "CCPG 與 DCAI 營收與營業利益率、Foundry 營業利益率路徑、外部代工營收、資本支出、營業費用、業外、稅率、非控制權益", "本模型",
           "Siliconomics 尚無伺服器 CPU、客戶端 PC、代工利益率或外部代工營收變數（README 列「建議補進 Siliconomics」）")]
    for y in (2027, 2028, 2029):
        rd.append((f"Wattonomics 錨定係數 CY{y}（敏感度案例用）", f"=wta_{y % 100}", "Siliconomics SIF_WTAnchor_Base（R4 比值）", SIL.VERSION, "Sens 頁「Wattonomics 錨定」案例；基準不乘"))
    for y in (26, 27, 28):
        rd.append((f"代工與邏輯設備年增 CY20{y}（對照，不驅動）", f"=eqfd{y}", f"Siliconomics eq_g_fd_{y}（SEMI）", SIL.VERSION, "只用於 Recon 與本頁 C0 參考列，不進任何計算"))
    rd.append(("對帳中旗標與另一側讀數比值（CY2026–CY2029）", "見本頁上方", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION, "兩讀數目標價（C4）"))
    ROW["rd0"] = r
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
        r += 1
    r += 1
    CW.cell(r, 1, "2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0，無檢驗力，故不判定、不計入對帳中；C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）").font = BOLD; r += 1
    CW.cell(r, 1, "本模型會計年度近似日曆年，FY(t) 對 CY(t) 直接對應。下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。")
    ROW["caveat"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定", "層級名稱與版本／口徑／證據等級／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr2"] = r; r += 1
    first = r

    def line(label, unit, yr_, f_model, f_layer, fmt, note):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, yr_)
        CW.cell(r, 4, f_model).number_format = fmt
        c = CW.cell(r, 5, f_layer); c.number_format = fmt; c.font = BLUE
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = SP
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
    nb = f"Siliconomics {SIL.VERSION}"
    line("顆數情境係數 情境 B", "x", "CY2028", "=um28_B", N["SIF_UnitsCoef_Base"][3], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um28_B（本版不改）")
    line("顆數情境係數 情境 B", "x", "CY2029", "=um29_B", N["SIF_UnitsCoef_Base"][4], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um29_B（本版不改）")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    for y in (26, 27, 28):
        line("代工與邏輯設備年增", "%/年", f"CY20{y}", f"=eqfd{y}", SIL.base(f"eq_g_fd_{y}"), "0.0%", f"{nb} eq_g_fd_{y}（SEMI）；In 頁 eqfd{y}（對照用，不進計算）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 2
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是英特爾自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。英特爾只有 DCAI 中的 AI 相關營收（約占營收 15%）隨顆數變動。"); r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    for i, y in enumerate((2027, 2028, 2029)):
        CW.cell(r, 1, "AI 加速器顆數合計（Siliconomics 基準係數）：基準 vs Wattonomics 換算值"); CW.cell(r, 2, "百萬顆"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, N["SIF_UnitsPlat_Total_Base"][i + 2]).number_format = "0.000"; CW.cell(r, 4).font = BLUE
        CW.cell(r, 5, f"=D{r}*{'CDE'[i]}${CR['ratio']}").number_format = "0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = SP; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Siliconomics SIF_UnitsPlat_Total_Base；Wattonomics {wat_ver} 經 SIF_DeployGW_AltRatio_{y} 換算；Derived；對帳中年份依 C4 並列兩讀數目標價，個股不裁決")
        r += 1
    r += 1
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t)；FY2030–FY2032 延伸期由 FY2029 收斂推得；FY2026 為前三季實績加 4Q26 指引的錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 美元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 美元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入，個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數建置：情境 B、C 於 CY2027 起 DCAI 中 AI 相關營收同乘比值，情境 A 不縮放；CCPG、Foundry、All Other 與其他假設不變，整本重建（tools/coord/alt_run.py）。"
                      "AI 相關營收只占營收約 15%，且 Intel 的獲利主要由 Foundry 虧損與資本支出決定，兩讀數差距有限。本模型會計年度與日曆年相近，無月數近似；敏感度頁另有 Wattonomics 錨定案例（用 SIF_WTAnchor_Base，與本比值不同）。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110
    r += 1
    # ---- 4. C0（實績）----
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.1.0；規則 v1.4.2 C0 第 2 點；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（1Q26–3Q26 季報的合併與分部數字、前三季資本支出；不用指引或預測）。層級側＝Siliconomics 現行基準的同口徑隱含值。"
                  "英特爾讀入的層級值只有 CY2028 起的 AI 加速器顆數係數（比例）與 Wattonomics 錨定係數；層級沒有 Intel 的營收、分部營收、Foundry 利益率或資本支出，"
                  "只有代工與邏輯設備的產業成長率（SEMI 群組，非 Intel 口徑）。主讀數與第二讀數都需要層級的同口徑年度值，本輪各項皆無，故不計算。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（結論）：本表沒有任何一列同時具備「公司已公布實績」與「獨立於本公司模型的層級隱含值」。本模型 FY2026 的 DCAI 中 AI 相關營收是依公司 DCAI 分部營收的假設拆分（公司未揭露 AI 相關金額），不是層級值；"
                  "CY2028 起的顆數係數只是比例，乘的是本公司的 AI 相關營收路徑，因此層級對英特爾沒有可對照的絕對量。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "公司揭露缺口與層級缺口：公司未揭露 DCAI 中 AI 相關營收、外部代工客戶營收與先進封裝營收的細項；Siliconomics 尚無伺服器 CPU、客戶端 PC、代工利益率與外部代工營收的行業層，"
                  "設備層只有「代工與邏輯」設備成長率（SEMI 群組），不是 Intel 資本支出口徑。以上實績取自公司財報新聞稿（IR 原文；網址見 Sources 頁）。")
    ROW["gap"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; r += 1
    inp = {}

    def put(key, label, unit, per, val, grade, src, fmt):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        inp[key] = r; r += 1
    put("q1", "合併營收：1Q26", "US$B", "2026-01～03", "=Q!C4", "Verified", "Q 頁 1Q26（公司 1Q26 財報新聞稿 2026-04-23；連格）", "#,##0.000")
    put("q2", "合併營收：2Q26", "US$B", "2026-04～06", "=Q!D4", "Verified", "Q 頁 2Q26（公司 2Q26 財報新聞稿 2026-07-23；連格）", "#,##0.000")
    put("q3", "合併營收：3Q26", "US$B", "2026-07～09", "=Q!E4", "Verified", "Q 頁 3Q26（公司 3Q26 財報新聞稿 2026-10-29；連格）", "#,##0.000")
    put("s9", "合併營收：2026 年前三季累計", "US$B", "2026-01～09", f"=D{r - 3}+D{r - 2}+D{r - 1}", "Derived", "＝上三列；與公司 3Q26 新聞稿累計營收一致（Q 頁核對列）", "#,##0.000")
    put("d3", "DCAI 營收：3Q26", "US$B", "2026-07～09", "=Q!E6", "Verified", "Q 頁 3Q26 DCAI 分部營收（公司新聞稿）", "#,##0.000")
    put("f3", "Intel Foundry 營業利益（虧損）：3Q26", "US$B", "2026-07～09", "=Q!E14", "Verified", "Q 頁 3Q26 Foundry 分部營業虧損（公司新聞稿）", "#,##0.000")
    put("e3", "外部代工營收：3Q26（Foundry 營收扣分部間沖銷）", "US$B", "2026-07～09", "=Q!E10", "Derived", "Q 頁：Foundry 含對內銷售營收 ＋ 分部間沖銷（公司新聞稿分部表）", "#,##0.000")
    put("cx9", "總資本支出：2026 年前三季", "US$B", "2026-01～09", "=Q!C37+Q!D37+Q!E37", "Verified／Derived", "Q 頁各季總資本支出（投資活動＋融資活動的 PP&E 增加；公司新聞稿現金流量表）", "#,##0.000")
    put("cxg", "參考：FY2026 總資本支出年增（前三季實績＋4Q26 假設）", "%", "FY2026 vs FY2025", "=Q!G37/Q!B37-1", "Derived", "Q 頁 FY2026 錨點 ÷ FY2025A − 1；第四季為假設（Assumed），故只列參考", "+0.0%;-0.0%")
    put("eqg", "參考：層級「代工與邏輯」設備成長率 CY2026（群組，非 Intel 口徑）", "%/年", "Siliconomics 基準", "=eqfd26", "Assumed（層級標記）",
        f"Siliconomics {SIL.VERSION} 設備層 eq_g_fd_26（SEMI 代工與邏輯設備成長；全產業群組，不是 Intel 資本支出），只列參考", "0.0%")
    put("ai26", "參考：DCAI 中 AI 相關營收 FY2026（本模型假設拆分，非實績、非層級值）", "US$B", "FY2026", "=PL!D" + str(B["dai"]), "Assumed", "PL 頁 FY2026 錨點：DCAI 營收 × 假設 AI 占比（In 頁 ai_sh26）；公司未揭露 AI 相關金額", "0.000")
    r += 1
    CW.cell(r, 1, "C0（實績）判定").font = BOLD; ROW["hdr_j"] = r; r += 1
    for j, h in enumerate(["項目", "", "", "", "", "", "判定（±15%）", "結論與缺漏揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1
    J = [("C0（實績）判定：合併營收與分部營收（US$）", "無可比實績",
          "結論：無可比實績。公司營收與分部實績完整（見上），但層級（Siliconomics、Wattonomics）沒有 Intel 營收、CCPG、DCAI 或外部代工市場規模；缺：Siliconomics 伺服器 CPU、客戶端 PC、外部代工營收層。不用本模型預測補位"),
         ("C0（實績）判定：Foundry 營業利益率", "無可比實績",
          "結論：無可比實績。公司 Foundry 分部營業虧損實績可得（見上），層級沒有代工利益率變數；待 Siliconomics 建代工利益率層後重做"),
         ("C0（實績）判定：DCAI 中 AI 相關營收", "無可比實績",
          "結論：無可比實績。公司未揭露 AI 相關營收金額；層級對本項只提供 CY2028 起的顆數係數（比例），沒有 Intel 口徑的絕對量"),
         ("C0（實績）判定：資本支出", "無可比實績",
          "結論：無可比實績。公司前三季資本支出實績可得，但層級只有「代工與邏輯」設備成長率（群組，非 Intel 口徑），且 Intel 資本支出含夥伴挹注與政府補助的毛額口徑差異，只列參考，不判定")]
    ROW["j1"] = r
    for a, gd, h in J:
        CW.cell(r, 1, a); CW.cell(r, 7, gd); CW.cell(r, 8, h); r += 1
    ROW["j4"] = r - 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["j1"]}:G{ROW["j4"]},"對帳中")')
    ROW["cnt4"] = r; ROW["last"] = r; r += 1
    CW.cell(r, 1, "讀法：英特爾 2026 年用前三季實績加 4Q 指引，營收、分部、資本支出都是公司已公布實績，不構成對層級的檢驗；層級對英特爾只約束 CY2028 起 DCAI 中 AI 相關營收的顆數比例。因此本輪 C0（實績）沒有可判定項目，對帳中 0；"
                  "外部估計（賣方共識、券商目標）只在 Recon 頁列參考。")
    ROW["read"] = r; r += 1
    ROW["inp"] = inp
    return ROW


def readme13(wb, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    r12 = lab(V, "DCF 情境價值推到 12 個月後")
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_fm, s_ai, s_cx, s_dcm = sens_r("Foundry 營業利益率"), sens_r("DCAI 中 AI 相關營收"), sens_r("總資本支出"), sens_r("DCAI 營業利益率")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.1.0 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0.0"), T(" 美元，相對現價 "), TX(up, SP), T("，評等"), rec, T("：DCF "), TX(dcf, "#,##0.0"), T(" 美元，同業比較法（相對倍數調整後）"), TX(cadj, "#,##0.0"), T(" 美元。"))
    put("rec", T("評等「"), rec, T("」：同業比較法採英特爾過去十年相對同業的本益比倍數（"), TX(coef, "0.00"), T(" 倍）；未調整的市場主流同業法 "), TX(cu, "#,##0.0"), T(" 美元。"))
    put("o1", T("主因：加權目標價 "), TX(tp, "#,##0.0"), T(" 美元、現價 "), TX("px", "#,##0.0"), T(" 美元（"), TX(up, SP), T("）。現價已預先計入 Foundry 轉盈與 AI 成長：以 12 個月後 EPS "), TX(feps, "0.00"),
        T(" 美元計，相當 "), TX(rv("pe1"), "0"), T(" 倍，同業中位數 "), TX(pmed, "0.0"), T(" 倍；情境 A 的 DCF 也只有 "), TX(f"Val!B{r12}", "#,##0.0"), T(" 美元（WACC "), TX("wacc", "0.0%"), T("），共識 "), TX(cons, "#,##0.0"), T(" 美元。"))
    put("o2", T("最大變數是 Foundry：FY2026 營業虧損 "), TX(f"-PL!F{PLA['s_fndoi']}", "0.0"), T(" 十億美元（利益率 "), TX(f"PL!F{PLA['s_fm']}", "0%"), T("），情境 B 到 FY2029 收斂到 "), TX(f"PL!I{PLA['s_fm']}", "0%"),
        T("；利益率 ±5 個百分點使目標價 "), TX(f"Sens!B{s_fm}", SP), T("／"), TX(f"Sens!C{s_fm}", SP), T("。總資本支出占營收 "), TX(f"PL!F{PLA['s_cxs']}", "0%"), T(" 降到 "), TX(f"PL!I{PLA['s_cxs']}", "0%"), T("。"))
    put("o3", T("AI 與 CPU：DCAI 營收 FY2029 "), TX(f"PL!I{PLA['s_dcr']}", "#,##0"), T(" 十億美元（AI 相關 "), TX(f"PL!I{PLA['s_dai']}", "#,##0"), T("）；AI 營收 ±30% 使目標價 "), TX(f"Sens!B{s_ai}", SP), T("／"), TX(f"Sens!C{s_ai}", SP),
        T("，DCAI 利益率 ±3 個百分點 "), TX(f"Sens!B{s_dcm}", SP), T("／"), TX(f"Sens!C{s_dcm}", SP), T("。外部代工營收 FY2029 "), TX(f"PL!I{PLA['s_ext']}", "0.0"), T(" 十億美元。"))
    put("o4", T("反向 DCF：要使 DCF 等於現價，折現率須降到 "), TX(rv("w1"), "0.0%"), T("（模型 "), TX("wacc", "0.0%"), T("），或 FY2028 起現金流放大 "), TX(rv("s1"), "0.0"), T(" 倍。"))
    put("o5", T("對帳中：CY2027–CY2029 AI 顆數的另一側讀數只影響 DCAI 中 AI 相關營收，另一側加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0.0"), T(" 美元（"), f"Coord!C{CR['rec']}", T("），個股不裁決；公司實績對層級：無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：營收＝CCPG＋DCAI＋All Other＋外部代工。FY2029 CCPG "), TX(f"PL!I{PLA['s_ccg']}", "#,##0"), T("、DCAI "), TX(f"PL!I{PLA['s_dcr']}", "#,##0"), T("（AI 相關 "), TX(f"PL!I{PLA['s_dai']}", "#,##0"),
        T("）、外部代工 "), TX(f"PL!I{PLA['s_ext']}", "0.0"), T("，營收 "), TX(f"PL!I{PLA['s_rev']}", "#,##0"), T(" 十億美元；Foundry 營業利益率 "), TX(f"PL!I{PLA['s_fm']}", "0%"), T("。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）；其餘沿用 v0.1.0（2026-11-17 初版）。個股 HTML 版面規格 v1.3（樣板 台積電 v0.5.5、聯電 v0.2.4）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數。章節為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；"
                  "含 Hist 頁（損益表歷史 FY2022–FY2025，公司財報新聞稿，美國 GAAP）、RevDCF 頁（反向 DCF）、PL 頁末標準損益表與分部營收推導、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（連既有逐案整本重算）、"
                  "Coord 頁（兩讀數目標價；讀取層級；C0：季營收、分部、資本支出實績對層級，皆無可比實績，對帳中 0）、README 頁末 v1.3 白話列。")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, G):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    PLA = pl_standard(wb, HR, CR, G["PB"]["B"])
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, G)
    CO = coord_c0(wb, SIL, CR, G["PB"]["B"])
    RMA = readme13(wb, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((
        ("公司：4Q24 與 FY2024 財報新聞稿（2025-01-30；FY2024、FY2023 損益表；Hist 頁用）", Q424),
        ("公司：4Q23 與 FY2023 財報新聞稿（2024-01-25，8-K 附件；FY2022 損益表比較欄；Hist 頁用）", Q423))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v010 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    BID = governance(wb, here)
    return f"v010_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 相對基準 {SNS['t0']} 起；Coord C0 {CO['rd0']}–{CO['last']}；README v1.3 {len(RMA)} 列；build_id {BID}"


# ---------- 治理（check_governance）：單一主要標記、來源誘因、build_id、模型狀態 ----------
import hashlib, os as _os
_INC = "公司指引、自揭數字或轉載的公司說法：預期管理誘因，口徑由提供者定義"


def governance(wb, here):
    from openpyxl.styles import Font
    for sh, tcol, scol, icol in (("In", 7, 8, 10), ("Q", 8, 9, 11)):
        ws = wb[sh]
        hr = next(r for r in range(1, 15) if ws.cell(r, tcol).value == "標記")
        ws.cell(hr, icol, "來源誘因").font = Font(bold=True); ws.column_dimensions[ws.cell(hr, icol).column_letter].width = 50
        for r in range(hr + 1, ws.max_row + 1):
            t = ws.cell(r, tcol).value
            if not isinstance(t, str) or t == "標記": continue
            if t not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"):
                p = [x.strip() for x in re.split(r"／|（", t)]
                prim = "Interested-party" if any(x.startswith("Interested-party") for x in p) else p[0]
                if prim not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"): prim = "Assumed"
                ws.cell(r, tcol).value = prim
                ws.cell(r, icol, ("混標「" + t + "」：主要標記取較弱者；" + _INC) if prim == "Interested-party" else "混標「" + t + "」：主要標記取第一項")
                t = prim
            if t == "Interested-party" and not ws.cell(r, icol).value: ws.cell(r, icol, _INC)
            src = str(ws.cell(r, scol).value or "")
            if t == "Verified" and re.search(r"；(Derived|Interested-party)）$", src): ws.cell(r, tcol).value = "Derived"
    SE = wb["Sens"]; V = wb["Val"]
    tpr = next(r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").startswith("目標價（DCF 與 Comparable 加權）"))
    snap = json.load(open(_os.path.join(here, "gov_snapshot.json"))) if _os.path.exists(_os.path.join(here, "gov_snapshot.json")) else {}
    IN = wb["In"]; hr = next(r for r in range(1, 15) if IN.cell(r, 7).value == "標記")
    last = next(r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")) - 1
    rng = f"In!$D${hr + 1}:$F${last}"
    r = SE.max_row + 3
    SE.cell(r, 1, "模型治理：Sens 主表為活頁簿公式；HTML 與敏感度快照是建置快照。只在 Excel 改 In 頁而未重跑建置時，下方模型狀態顯示 STALE").font = Font(bold=True); r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄，情境驅動對照以上）"); SE.cell(r, 4, f"=SUMPRODUCT(N(+{rng}),ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; fp = r; r += 1
    SE.cell(r, 1, "輸入指紋（建置快照）"); SE.cell(r, 4, snap.get("fp")); fb = r; r += 1
    SE.cell(r, 1, "加權目標價（建置快照）"); SE.cell(r, 4, snap.get("tp")); tb = r; r += 1
    SE.cell(r, 1, "模型狀態（CURRENT＝快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑建置）").font = Font(bold=True)
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{fb}),ISNUMBER(D{tb}),ABS(D{fp}-D{fb})<=0.000000001*MAX(1,ABS(D{fb})),ABS(Val!$B${tpr}-D{tb})<0.005),"CURRENT","STALE")'); r += 1
    h = hashlib.sha256(b"0.1.1")
    for fn in sorted(["build_intc.py", "v010_post.py", "show_intc.py", "segments.json", "coord_alt.json"]):
        p = _os.path.join(here, fn)
        if _os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    bid = "intc-v0.1.1-" + h.hexdigest()[:10]
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid)
    a1 = wb["README"]["A1"].value
    wb["README"]["A1"].value = (a1 + ("" if str(a1).endswith('"') or not str(a1).startswith("=") else "") ) if False else (str(a1) + '&"；build_id ' + bid + '。"' if str(a1).startswith("=") else str(a1) + "；build_id " + bid + "。")
    return bid
