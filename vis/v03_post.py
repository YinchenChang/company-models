# 世界先進 v0.3（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 聯電 umc/v024_post.py、台積電、NVIDIA v0.3.10）＋三層協調讀取層級與 C0（實績）。
# 世界先進不讀產業量（只讀 Siliconomics 匯率），兩讀數不適用、不加入 coord_meta 的 depends_on_years；Coord 頁只列讀取層級與 C0（實績）。
# 新增（不回饋計算）：Hist 頁（stockanalysis 年度損益表 FY2022–FY2025）、PL 頁末標準損益表與分部營收推導、RevDCF 頁、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準、Coord 頁、README 頁末 v1.3 白話列。
# build_vis.py 存檔前呼叫 apply(wb, here, SIL, G)。
import os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
SAIS = "https://stockanalysis.com/quote/tpex/5347/financials/income-statement/"
VQ3 = "https://news.nextapple.com/finance/20261106/8A39A01AEB8B43D60002F799C6B267EB"
VQ3B = "https://udn.com/news/story/7240/9862010"
VSEP = "https://finance.technews.tw/2026/10/08/vis-5347-202609-financial-report/"
VDA_KEYS = ["設備 50%＋廠房 35%", "設備 70%＋廠房 20%", "設備年限 5 年", "設備年限 8 年"]
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




# ---------- RevDCF：反向 DCF（樣板 世界先進 umc/v024_post.py；世界先進 v0.3 為 SOTP：DCF 頁情境 B 路徑＝既有 8 吋，淨現金含 VSMC SOTP 調整） ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 6, 7, 8
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：世界先進 v0.3 起 DCF 為 SOTP（既有 8 吋業務 100%＋VSMC 60%），本頁只動 8 吋路徑，VSMC 價值固定在基準（WACC 反解時亦不重估 VSMC）。DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
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
    kv("px", "現價（新台幣）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日＋VSMC SOTP 調整（情境 B；新台幣十億元）", "=" + ncr + "+sotp_x_B", "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節）＋ VSMC 60% 企業價值 − VSMC 淨現金 × 40%（VSMC 頁；同 DCF 頁每股價值公式）")
    kv("sh", "評價用股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（新台幣十億元）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；新台幣十億元）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（新台幣）", f"=DCF!$C${ps}", "#,##0.00")
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
    yr("rev", "營收（新台幣十億元）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（新台幣十億元）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率（EBIT ÷ 營收）", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tn}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（新台幣十億元）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（新台幣十億元）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（新台幣十億元，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（新台幣十億元）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（新台幣十億元）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年",IF(B{A["m1"]}<MIN(C{Y["m"]}:G{Y["m"]}),"低於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 各年自由現金流與終值等比例放大，近似「營收與成本同步放大」；FY2027 不動）").font = BOLD; r += 1
    kv("a", "不隨倍數變動的 EV：FY2027 現值（新台幣十億元）", f"=B{Y['pv']}", "#,##0.0")
    kv("b", "隨倍數變動的 EV：FY2028–FY2032 現值＋終值現值（新台幣十億元）", f"=SUM(C{Y['pv']}:G{Y['pv']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的放大倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（新台幣十億元）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；新台幣十億元）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近邊界，表示區間內無解）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（新台幣）"], 1):
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
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（新台幣元，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-w_dcf_eff*{dcf_tp})/((1-w_dcf_eff)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0.0"),
                     T(" 元；要支持現價 "), TX("px", "#,##0.0"), T(" 元，企業價值須"), f'IF(B{A["evx"]}>=1,"放大","縮小")', T("為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}<wacc,"降到 ","升到 ")', TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面乘以 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億新台幣，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率改為 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0.0"), T(" 元相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 相對倍數調整係數 "), TX(coef, "0.00"), T("）。換言之，市場目前給世界先進的倍數是同業的 "),
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
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.3 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.00"), T(" 新台幣；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("w_dcf_eff", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-w_dcf_eff", "0%"), T("（DCF 有效權重＝設定權重 × DCF 適用，VSMC 頁）"))),
            ("④ 市場共識目標價", cons, "In 頁市場共識目標價（stockanalysis 分析師平均；Interested-party）")]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = SP; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=In!$B$" + str(lab(wb["In"], "px", exact=True))); r += 1
    return r0




# ---------- Sens：加權目標價相對基準（連 Sens 頁既有逐案整本重算結果） ----------
SENS_LAB = {"fx": "匯率", "cap_adj": "8 吋產能係數", "ut_adj": "8 吋稼動率 ±3 個百分點", "asp_adj": "8 吋售價水準 ±3%", "dag_adj": "折舊年增 ±5 個百分點",
            "cc_g": "每片現金成本年增 0%／4%", "phi": "營業費用彈性", "nonop_f": "營業外淨額", "tax": "有效稅率", "vcap_adj": "VSMC 產能 ±20%", "vasp": "VSMC 售價（美元／片）"}
FD_PAIRS = [("折現率 WACC −1／＋1 個百分點", "WACC −1pt", "WACC ＋1pt"), ("β（產業 β ∓ 0.3 重新加槓桿）", "β：低", "β：高"), ("股權風險溢酬 4%／6%", "ERP：低", "ERP：高"),
            ("永續成長率", "永續成長率 g：低", "永續成長率 g：高"), ("永續期新投資報酬率＝WACC＋5 個百分點（只有高案）", None, "RONIC＝WACC＋5pt"),
            ("資本支出占營收 −1／＋1 個百分點", "資本支出 ÷ 營收 −1pt", "資本支出 ÷ 營收 ＋1pt"), ("營運資金占營收", "營運資金 ÷ 營收：低", "營運資金 ÷ 營收：高"),
            ("淨現金不含長期投資（只有低案）", "淨現金不含長期投資", None), ("終值改用自由現金流成長法（只有低案）", "終值改用 FCFF ×（1＋g）", None),
            ("美元口徑折現（只有低案）", "美元口徑折現", None), ("DCF 權重 0.3／0.7", "DCF 權重 0.3", "DCF 權重 0.7")]


def sens_tor(wb, G):
    SE = wb["Sens"]; V = wb["Val"]
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.3 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務案例連本頁上方逐案整本重算結果（相對基準欄），目標價參數取上方 DCF 參數案例；白話標籤，不回饋計算")
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
    for nm in VDA_KEYS:   # v0.3：VSMC 建廠折舊拆分（驗收 C-04 建議區間；Sens 頁 VSMC 折舊表）
        rr_ = lab(SE, nm, exact=True)
        SE.cell(r, 1, "VSMC 折舊：" + nm); SE.cell(r, 2, f"=J{rr_}").number_format = SP
        SE.cell(r, 4, f"Sens 頁 VSMC 建廠折舊假設表第 {rr_} 列（每案三情境重算 PL、8 吋 DCF 與 VSMC SOTP）"); r += 1
    last = r - 1
    SE.cell(r, 1, "核對：基準列相對基準（應為 0）"); SE.cell(r, 2, "=J4").number_format = "0.000000"
    return dict(t0=r0, first=first, last=last, state=r)



# ---------- PL：標準損益表（新台幣十億元） ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "D", "D": "E", "E": "F"}                       # 標準損益表歷史欄 → Hist 欄
EC = {"F": "D", "G": "E", "H": "F", "I": "G", "J": "H", "K": "I", "L": "J"}   # 預測欄 → PL 頁推導區塊欄（FY2026–FY2032）
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}
DCOL = {"J": "F", "K": "G", "L": "H"}                     # FY2030–32 → DCF 頁欄


def pl_standard(wb, HR, B):
    P = wb["PL"]
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.3 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加 4Q26 指引的錨點）。新台幣十億元；EPS 新台幣。"
                 "只引用本頁上方情境 B 推導區塊、Hist 與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（TIFRS 合併財報，含 VSMC 100%；淨利為歸屬母公司，VSMC 損益的 40% 歸 NXP 列非控制權益）。世界先進不另發布 Non-GAAP，法說與市場共識 EPS 皆以此口徑。"
                 "營業費用＝毛利 − 營業利益；其他收入（淨額）＝稅前利益 − 營業利益。歷史三年取 stockanalysis 年度損益表（Hist 頁）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益＝（歸屬母公司淨利 ×（1＋FY2029 非控制權益 ÷ 淨利））÷（1 − 稅率）− 營業外淨額（由淨利還原，與 v0.2.4 DCF 同式）；"
                 "毛利率與營業費用率沿用前一年據以算出營業成本、毛利與營業費用，與延伸期營業利益的差額列核對列，不改淨利與目標價。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":
        P[f"{col}{r}"] = YEARS[COLS.index(col)]
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
    row("rev", "營收", "NT$B", hh("rev") + eb("rev", "FGHIJKL"), M, "FY2026E–FY2029E＝PL 頁情境 B 推導區塊（8 吋晶圓＋VSMC 12 吋）；FY2030E 起為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{g['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{g['rev']}/{p}{g['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 取 Hist 頁")
    row("cogs", "營業成本（COGS）", "NT$B", hh("cogs") + [f"={c}{g['rev']}-{c}{g['gp']}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "NT$B", hh("gp") + eb("gp") + [f"={c}{g['rev']}*{c}{g['gm']}" for c in "JKL"], M, "FY2026E–FY2029E＝PL 頁毛利；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{g['gp']}/{c}{g['rev']}" for c in "CDEFGHI"] + [f"={p}{g['gm']}" for p in "IJK"], PCT, "FY2030E–FY2032E＝前一年")
    row("opx", "營業費用（含其他營業損益）", "NT$B", [f"=Hist!{HC[c]}{HR['opx']}" for c in "CDE"] + eb("opx")
        + [f"={c}{g['rev']}*{p}{g['opx']}/{p}{g['rev']}" for p, c in (("I", "J"), ("J", "K"), ("K", "L"))], M, "歷史＝毛利 − 營業利益；FY2026E–FY2029E＝PL 頁營業費用；FY2030E 起＝營收 × 前一年營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + eb("oi") + [f"=(PL!{EC[c]}{B['ni']}*(1+PL!$G${B['mi']}/PL!$G${B['ni']}))/(1-tax)-nonop_f" for c in "JKL"], M, "FY2030E–FY2032E＝由淨利還原（同 v0.2.4 合併 DCF 的 EBIT）")
    row("opm", "營業利益率", "%", [f"={c}{g['oi']}/{c}{g['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外）", "NT$B", [f"={c}{g['pti']}-{c}{g['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（利息、投資損益、匯兌）")
    row("pti", "稅前利益", "NT$B", hh("pti") + eb("pti") + [f"={c}{g['oi']}+nonop_f" for c in "JKL"], M, "FY2030E 起＝延伸期營業利益＋營業外淨額（In 頁，與 DCF 還原同式）")
    row("tax", "所得稅", "NT$B", hh("tax") + eb("tax") + [f"={c}{g['pti']}*tax" for c in "JKL"], M, "FY2030E 起＝稅前 × 有效稅率（In 頁）")
    row("nci", "非控制權益損益", "NT$B", [f"=Hist!{HC[c]}{HR['nci']}" for c in "CDE"] + eb("mi") + [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['pti'] + 3}" for c in "JKL"], M,
        "FY2027E–FY2029E＝VSMC 毛利 ×（1 − 稅率）× 40%（虧損為負）；FY2030E 起＝稅前 − 所得稅 − 稅後淨利")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", hh("nip") + eb("ni", "FGHIJKL"), M, "預測期＝PL 頁情境 B 歸屬母公司淨利")
    row("nm", "淨利率", "%", [f"={c}{g['ni']}/{c}{g['rev']}" for c in COLS], PCT)
    row("sh", "股數（十億股；基本加權平均）", "十億股", hh("sh") + eb("shr", "FGHIJKL"), "0.000")
    row("eps", "EPS（新台幣；基本，歷史為公司公告值）", "NT$", hh("eps") + eb("eps", "FGHIJKL"), "0.00", "預測期＝PL 頁 EPS（＝歸屬母公司淨利 ÷ 股數）")
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
    row2("r_niy", "淨利年增率（YoY）", "%", [f"=C{g['ni']}/Hist!C{HR['nip']}-1"] + [f"={c}{g['ni']}/{p}{g['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 歸屬母公司淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row2("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{g['rev']}-{c}{g['cogs']}-{c}{g['gp']}" for c in COLS], "0.000")
    row2("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2032E 為差額，見下）", "NT$B", [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "CDEFGHI"] + [None] * 3, "0.000")
    row2("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利還原）；目標價用延伸期淨利，不受影響", "NT$B", [None] * 7 + [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "JKL"], "0.000")
    row2("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{g['rev']}" for c in "JKL"], PCT)
    row2("c3", "核對：稅前利益 − 所得稅 − 非控制權益 − 稅後淨利", "NT$B", [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['nci']}-{c}{g['ni']}" for c in COLS], "0.000")
    row2("c4", "核對：稅後淨利 ÷ 股數 − EPS（預測期）", "NT$", [None] * 3 + [f"={c}{g['ni']}/{c}{g['sh']}-{c}{g['eps']}" for c in "FGHIJKL"], "0.0000")
    row2("c5", "核對：營收 − Q 頁 FY2026 錨點（FY2026E）／ − PL 頁情境 B 營收（FY2027E–FY2029E）", "NT$B", [None] * 3 + [f"=F{g['rev']}-a26_rev"] + [f"={c}{g['rev']}-PL!{EC[c]}{B['rev']}" for c in "GHI"], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度＝日曆年，FY(t) 直接對應 CY(t)。世界先進不讀產業量（只讀 Siliconomics 匯率），沒有估值依賴的對帳中年份，標題不加「對帳中」。")
    L["ycaveat"] = r; r += 2
    # ---- 分部營收推導（情境 B；HTML 命題、驅動因子章用） ----
    P.cell(r, 1, "分部營收推導（情境 B；v0.3）：8 吋晶圓＝產能 × 稼動率 × 美元售價 × 匯率；VSMC 12 吋＝產能 × 稼動率 × 售價 × 匯率（FY2027 第一季量產）；兩者相加＝營收。只引用本頁情境 B 推導區塊").font = BOLD; r += 1
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
    seg("s_cap", "產能（8 吋）", "千片", pb("cap"), "#,##0", "FY2026 為公司揭露；FY2027 起持平（粗線寬轉細線寬）")
    seg("s_ut", "平均稼動率（8 吋）", "%", pb("ut"), PCT)
    seg("s_ship", "晶圓出貨（8 吋）", "千片", pb("ship"), "#,##0")
    seg("s_asp", "平均售價（8 吋晶圓）", "US$/片", pb("asp"), "#,##0")
    seg("s_fx", "匯率", "NT$/US$", pb("fxr"), "0.00")
    seg("s_w", "8 吋晶圓營收", "NT$B", pb("wrev"), M)
    seg("s_vcap", "VSMC 產能（12 吋）", "千片", pb("vcap"), "#,##0")
    seg("s_vut", "VSMC 稼動率", "%", pb("vut"), PCT)
    seg("s_v", "VSMC 營收（12 吋）", "NT$B", pb("vrev"), M)
    seg("s_rev", "營收", "NT$B", pb("rev"), M)
    seg("s_chk", "核對：8 吋＋VSMC − 營收（應為 0）", "NT$B", lambda c: f"={c}{L['s_w']}+{c}{L['s_v']}-{c}{L['s_rev']}", "0.000")
    seg("s_sh", "12 吋（VSMC）÷ 營收", "%", lambda c: f"={c}{L['s_v']}/{c}{L['s_rev']}", PCT)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0)
    return L


# ---------- Hist：損益表歷史（新台幣十億元；stockanalysis） ----------
HIS = [
    ("rev", "營業收入", 51.694, 38.273, 44.055, 48.591),
    ("cogs", "營業成本", 27.771, 27.827, 32.122, 34.937),
    ("gp", "營業毛利", 23.924, 10.446, 11.933, 13.654),
    ("oi", "營業利益", 17.807, 5.774, 7.111, 7.773),
    ("pti", "稅前淨利", 18.851, 8.834, 8.752, 9.251),
    ("tax", "所得稅費用", 3.571, 1.464, 1.705, 1.480),
    ("ni", "本期淨利（合併，含非控制權益）", 15.280, 7.370, 7.047, 7.770),
    ("nip", "歸屬母公司業主淨利（模型淨利口徑）", 15.280, 7.370, 7.046, 7.908),
]
EPSR = [("eps", "EPS 基本（stockanalysis；新台幣元）", (9.32, 4.50, 4.21, 4.30)),
        ("eps_d", "EPS 稀釋（stockanalysis；新台幣元）", (9.07, 4.43, 4.16, 4.22)),
        ("sh", "基本加權平均股數（十億股）", (1.639, 1.639, 1.675, 1.839))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 損益表歷史（v0.3 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度＝日曆年。金額：新台幣十億元（stockanalysis 年度損益表，百萬元 ÷ 1,000，S&P Global 轉載公司財報；"
               f"2026-11-15 讀取，{SAIS}）。標準損益表的營業費用＝毛利 − 營業利益")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, f"stockanalysis 年度損益表（{SAIS}）")
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, "同上（EPS (Basic)、EPS (Diluted)、Shares Outstanding (Basic)）；Q 頁 FY2025 EPS 4.22 與 stockanalysis 稀釋 EPS 相同，基本 EPS 4.30 以 18.39 億股計（公司公告基本 EPS：未找到，待核）")
        R[k] = r; r += 1
    for k, labl, f_, note in (("nonop", "營業外收入及支出淨額（＝稅前 − 營業利益）", "={c}{pti}-{c}{oi}", "Derived"),
                              ("opx", "營業費用（＝毛利 − 營業利益）", "={c}{gp}-{c}{oi}", "Derived"),
                              ("nci", "非控制權益損益（＝本期淨利 − 歸屬母公司淨利）", "={c}{ni}-{c}{nip}", "FY2025 為 VSMC 量產前費用的 NXP 分攤（負值）")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "#,##0.000"
        H.cell(r, 7, "Derived"); H.cell(r, 8, note); R[k] = r; r += 1
    r += 1
    H.cell(r, 1, "核對（新台幣十億元；百萬元換算可有 ±0.002 捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"), ("稅前淨利 − 所得稅 − 本期淨利", "={c}{pti}-{c}{tax}-{c}{ni}"),
                     ("歸屬母公司淨利 ÷ 基本股數 − 基本 EPS（EPS 四捨五入到分）", "={c}{nip}/{c}{sh}-{c}{eps}")):
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "FY2025：本表 − Q 頁 FY2025A 錨點（營收、毛利、營業利益、歸屬母公司淨利）")
    for c, k, nm in (("C", "rev", "a25_rev"), ("D", "gp", "a25_gp"), ("E", "oi", "a25_oi"), ("F", "nip", "a25_ni")):
        H[f"{c}{r}"] = f"=F{R[k]}-{nm}"; H[f"{c}{r}"].number_format = "0.000"
    R["chk1"] = r; r += 2
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 60
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["H"].width = 100
    return R


# ---------- Coord：讀取層級與 C0（實績）；兩讀數不適用 ----------
def coord_c0(wb, SIL):
    CW = wb.create_sheet("Coord")
    r = 1
    CW.cell(r, 1, "Coord — 三層協調（v0.3；三層基準協調規則 v1.4.2 C0、C4、C5；工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）。只新增，不改上游值").font = BOLD; r += 2
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估）").font = BOLD; ROW["s1"] = r; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["h1"] = r; r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", SIL.VERSION, f"siliconomics/data/{SIL.FILE}（只讀匯率 fx_twd）", SIL.VERSION, "世界先進為硬體供應鏈個股，讀 Siliconomics；不讀 Wattonomics"),
          ("匯率（FY2027 起）", "=fx", "In 頁 fx（Siliconomics fx_twd）", SIL.VERSION, "美元售價換算新台幣（8 吋與 VSMC）"),
          ("產業量（部署 GW、加速器顆數、AI 市場）", "不讀", "—", "—", "世界先進營收＝8 吋產能 × 稼動率 × 售價＋VSMC；AI 相關占比只作記憶列。顆數情境係數、Wattonomics 錨定與另一側讀數皆不適用"),
          ("兩讀數目標價（C4）", "不適用", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION,
           "估值不依賴對帳中年份的產業量；不加入 coord_meta 的 hardware_companies／depends_on_years（PR 說明）"),
          ("公司自有變數（不由層級提供）", "—", "8 吋產能、稼動率、售價、每片現金成本、折舊；VSMC 產能、稼動率、售價、毛利率、建廠折舊；營業費用、業外、稅、股數", "—",
           "Siliconomics 尚無成熟製程代工的稼動率、價格或市場規模（待併入 Siliconomics）")]
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.00"
        r += 1
    r += 1
    CW.cell(r, 1, "2. 讀取層級（同源，不判定）：個股讀入的層級值 vs 層級值").font = BOLD; ROW["s2"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "年度", "模型值", "層值", "差距（模型 ÷ 層 − 1）", "判定", "層級名稱與版本／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["h2"] = r; r += 1
    CW.cell(r, 1, "新台幣兌美元（FY2027 起）"); CW.cell(r, 2, "NT$/US$"); CW.cell(r, 3, "CY2027")
    CW.cell(r, 4, "=fx").number_format = "0.00"; c = CW.cell(r, 5, SIL.base("fx_twd")); c.number_format = "0.00"; c.font = BLUE
    CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = SP; CW.cell(r, 7, SAME); CW.cell(r, 8, f"Siliconomics {SIL.VERSION} fx_twd"); ROW["fx"] = r; r += 2
    CW.cell(r, 1, "3. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點）").font = BOLD; ROW["s3"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績；層級側＝同口徑的層級隱含值。世界先進的營收來自 8 吋成熟與特殊製程代工與 VSMC 12 吋，Siliconomics 沒有成熟製程代工的市場、份額、稼動率或價格，"
                  "Wattonomics 與 Tokenomics 也沒有對應量，故沒有可比的層級隱含值；實績只列出，不判定。"); ROW["m"] = r; r += 1
    CW.cell(r, 1, "結論：無可比實績（層級無此量）；對帳中項目 0 項。缺的不是公司揭露，而是層級：成熟製程代工市場與份額尚未併入 Siliconomics。"); ROW["c"] = r; r += 1
    for j, h in enumerate(["項目", "期間", "值", "單位", "層級隱含值", "判定", "口徑、來源、證據等級"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["h3"] = r; r += 1
    acts = [("營收（月營收累計）", "2026-01～09", 42.479, "NT$B", "Verified：月營收公告前 9 月 424.79 億元（科技新報 2026-10-08）"),
            ("營收（季）", "3Q26", "=Q!E4", "NT$B", "Verified：157.02 億元（季增 10.2%、年增 27.2%；壹蘋新聞網 2026-11-06 法說）"),
            ("毛利率（季）", "3Q26", "=q3_gm", "%", "Verified：33.6%（同上）"),
            ("歸屬母公司淨利（季）", "3Q26", "=q3_ni", "NT$B", "Verified：30.99 億元、EPS 1.62 元（同上；經濟日報同值）"),
            ("稼動率（季）", "3Q26", "=Q!E6", "%", "Interested-party：公司法說約 90%（同上）"),
            ("資本支出（年度指引）", "2026", "=cxg_26", "NT$B", "Interested-party：公司 600–700 億元（約九成 VSMC；同上）；指引不算實績")]
    ROW["a0"] = r
    for labl, per, v, u, note in acts:
        CW.cell(r, 1, labl); CW.cell(r, 2, per)
        c = CW.cell(r, 3, v); c.number_format = "0.0%" if u == "%" else "#,##0.000"
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        CW.cell(r, 4, u); CW.cell(r, 5, "無（層級無此量）"); CW.cell(r, 6, "無可比實績"); CW.cell(r, 7, note); r += 1
    ROW["a1"] = r - 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 6, f'=COUNTIF(F{ROW["a0"]}:F{ROW["a1"]},"對帳中")'); ROW["cnt"] = r; r += 2
    CW.cell(r, 1, "C4／C5：本模型沒有估值依賴的對帳中年份（不讀產業量），不並列兩讀數；本版沒有可往上回饋的新證據（3Q26 實績為公司層級，層級無對應量）。"); ROW["c45"] = r; r += 1
    CW.column_dimensions["A"].width = 46
    for c in "BCDEF": CW.column_dimensions[c].width = 16
    CW.column_dimensions["G"].width = 100; CW.column_dimensions["H"].width = 40
    ROW["last"] = r - 1
    return ROW


def readme13(wb, PLA, RVA, SNS, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; VS = wb["VSMC"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_fx, s_asp, s_ut, s_vasp = sens_r("匯率"), sens_r("8 吋售價水準"), sens_r("8 吋稼動率"), sens_r("VSMC 售價")
    vs = lambda p: next(r for r in range(VS.max_row, 0, -1) if str(VS.cell(r, 1).value or "").startswith(p))
    r_evv, r_ok = vs("VSMC EV（100%）"), vs("DCF 適用（8 吋")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0.0"), T(" 元，相對現價 "), TX(up, SP), T("，評等"), rec, T("：DCF（8 吋＋VSMC 六成，SOTP）"), TX(dcf, "#,##0.0"), T(" 元，同業比較法 "), TX(cadj, "#,##0.0"),
        T(" 元；VSMC 量產初期折舊重，壓低近兩年 EPS。"))
    put("rec", T("評等「"), rec, T("」：DCF 有效權重 "), TX("w_dcf_eff", "0%"), T("（三情境 DCF 皆適用才計入）；同業比較法用 12 個月後 EPS "), TX(feps, "0.00"),
        T(" 元，受 VSMC 折舊壓低；未調整的市場主流同業法 "), TX(cu, "#,##0.0"), T(" 元。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0.0"), T(" 元，現價 "), TX("px", "#,##0.0"), T(" 元，差距 "), TX(up, SP), T("；DCF "), TX(dcf, "#,##0.0"), T(" 元、同業比較法 "), TX(cadj, "#,##0.0"),
        T(" 元，分析師共識 "), TX(cons, "#,##0.0"), T(" 元。"))
    put("o2", T("8 吋稼動率與售價決定本業：情境 B 的 FY2028 稼動率 "), TX(f"PL!H{PLA['s_ut']}", "0%"), T("、售價每片約 "), TX(f"PL!H{PLA['s_asp']}", "#,##0"),
        T(" 美元；VSMC 占營收 FY2028 "), TX(f"PL!H{PLA['s_sh']}", "0.0%"), T("，100% 企業價值情境 B 為 "), TX(f"VSMC!D{r_evv}", "#,##0.0"), T(" 十億元（建廠支出與折舊期虧損大於穩態現金流）。"))
    put("o3", T("最大變數（加權目標價相對基準）：匯率低／高案 "), TX(f"Sens!B{s_fx}", SP), T("／"), TX(f"Sens!C{s_fx}", SP),
        T("；8 吋售價 ±3% "), TX(f"Sens!B{s_asp}", SP), T("／"), TX(f"Sens!C{s_asp}", SP), T("；8 吋稼動率 ±3 個百分點 "), TX(f"Sens!B{s_ut}", SP), T("／"), TX(f"Sens!C{s_ut}", SP),
        T("；VSMC 售價低／高 "), TX(f"Sens!B{s_vasp}", SP), T("／"), TX(f"Sens!C{s_vasp}", SP), T("。"))
    put("o4", T("反向 DCF：WACC "), TX("wacc", "0.0%"), T("，要使 DCF 等於現價，折現率須為 "), TX(rv("w1"), "0.0%"), T("，或 FY2028 起 8 吋現金流放大 "), TX(rv("s1"), "0.00"),
        T(" 倍；現價相當於 12 個月後 EPS 的 "), TX(rv("pe1"), "0.0"), T(" 倍。"))
    put("o5", T("三層協調：世界先進只讀 Siliconomics 匯率，不讀產業量，沒有對帳中年份，兩讀數不適用；公司實績（月營收、3Q26 損益）對層級無可比量。C0 對帳中 "), TX(f"Coord!F{CO['cnt']}", "0"), T(" 項。"))
    put("drv", T("分部營收推導（情境 B）：8 吋晶圓＝產能 × 稼動率 × 美元售價 × 匯率，加 VSMC 12 吋。FY2028 8 吋出貨 "),
        TX(f"PL!H{PLA['s_ship']}", "#,##0"), T(" 千片、8 吋營收 "), TX(f"PL!H{PLA['s_w']}", "#,##0.0"), T("、VSMC "), TX(f"PL!H{PLA['s_v']}", "#,##0.0"), T("、營收 "), TX(f"PL!H{PLA['s_rev']}", "#,##0.0"), T(" 十億新台幣。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：依 v0.2.4 正式驗收報告（2026-11-09，不通過）修正四項 Critical 並滾入 3Q26 實績：(C-01) 不再於負 NOPAT 套終值、每股價值不截零，VSMC 明確期延到 FY2040、未正常化者終值不適用，DCF 不適用時權重取 0；"
                  "(C-02) DCF 改 SOTP：既有 8 吋 100%＋VSMC 企業價值 60%，NXP 40% 以公允價值扣除；(C-03) 不再以帳面倒推 NXP 剩餘注資；(C-04) VSMC 總投資拆為設備 60%（6 年）、廠房 30%（20 年）、其他 10% 不折舊；"
                  "FY2026 錨點改用 3Q26 實績與 4Q26 指引。另：個股 HTML 版面規格 v1.3（Hist、標準損益表、RevDCF、Val 0.7、Sens 末表、Coord 讀取層級與 C0）；兩讀數不適用")
    out["ver"] = r; r += 1
    # v0.3 獨立核對（文字）：新舊對照另起一列，數值寫死（歷史列不隨重算變動；建置後以 xlsx 核對）
    RD.cell(r, 1, "v0.3 新舊對照（v0.2.4 → v0.3，當時數值）：現價 126.0 → 126.0，同業前瞻 P/E 中位數 21.1 → 21.1，DCF 目標價 0.00 → 33.85（情境價值 A／B／C 71.05／33.94／−22.17，不截零），"
                  "Comparable（相對倍數調整後）19.36 → 60.96，加權目標價 9.68 → 47.41，上檔空間 −92.3% → −62.4%，投資建議 強力放空 → 強力放空。")
    out["cmp"] = r; r += 1
    # README A4 核心問題：倍數改由公式帶出（原文為 v0.1 的 25 倍、27.4 倍）
    RD["A4"] = F(T("核心問題：8 吋稼動率與售價回升能走多遠（情境 A 稼動率 95%、售價 +6%），VSMC 量產初期的毛利率與非控制權益讓 EPS 增加多少；現價約 "), TX(vb("現價 ÷ FY2026 EPS（錨點）"), "0.0"),
                 T(" 倍 FY2026 EPS、"), TX(vb("現價 ÷ 未來 12 月 EPS（模型 B）"), "0.0"), T(" 倍未來 12 月 EPS（模型 B），與同業前瞻 P/E 中位數 "), TX(pmed, "0.0"), T(" 倍差多少。"))
    return out


def apply(wb, here, SIL, G):
    HR = hist_is(wb)
    PLA = pl_standard(wb, HR, G["PB"]["B"])
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, G)
    CO = coord_c0(wb, SIL)
    RMA = readme13(wb, PLA, RVA, SNS, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((("媒體：壹蘋新聞網 2026-11-06（3Q26 法說：營收、毛利率、營業利益率、淨利、EPS、4Q26 指引、2026 折舊與資本支出）", VQ3),
                                ("媒體：經濟日報 2026-11-06（3Q26 法說：營收、淨利、EPS、4Q26 指引）", VQ3B),
                                ("資料商：stockanalysis 年度損益表 FY2022–FY2025（Hist 頁）", SAIS),
                                ("驗收報告：世界先進個股模型_v0.2.4_正式驗收與建議報告（2026-11-09；Drive 02-AI供應鏈）", "Google Drive（Andy 專案資料夾）"))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v03 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v03_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 相對基準 {SNS['t0']} 起；Coord 頁；README v1.3 {len(RMA)} 列"
