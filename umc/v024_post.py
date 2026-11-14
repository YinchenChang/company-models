# 聯電 v0.2.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 台積電 v0.5.5、NVIDIA v0.3.10）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.2.3 逐格比對驗證；In 頁既有列與 um28／um29、wta 等名稱與值不動）：
#   Hist 頁（新）：合併損益表 FY2022–FY2025（四年皆取公司 6-K 年度合併財務報告，台灣版 IFRS〔TIFRS〕，新台幣十億元；不混用 20-F 的 IASB 版 IFRS）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，新台幣十億元）、比率區、核對列、分部營收推導（情境 B）
#   RevDCF 頁（新）：反向 DCF
#   Val 頁末：0.7 估值並列四個目標價；Sens 頁末：加權目標價相對基準（本頁既有逐案整本重算結果，連格）
#   Coord 頁末：三層協調——讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）
#   Sources 頁：新增來源；README 頁末：v1.3 展示層白話列與本版說明
import os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.4", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
EDGAR = "https://www.sec.gov/Archives/edgar/data/1033767/"
AR25 = EDGAR + "000119312526068366/umc-ex99_1.htm"       # 6-K 2026-02-25：2025 年合併財務報告（英譯）
Q326 = EDGAR + "000119312526435003/umc-ex99_1.htm"       # 6-K 2026-10-28：2026 年第三季合併財務報告（英譯）
OCT26 = EDGAR + "000119312526448404/umc-ex99_1.htm"      # 6-K 2026-11-05：2026 年 10 月營收公告
F24 = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001033767&type=20-F"   # 20-F 清單（FY2024 20-F：0001193125-25-092142）
AR24 = EDGAR + "000129281425000981/ex99-1.htm"           # 6-K 2025-03-20：2024 年合併財務報告（英譯；FY2023 比較欄，TIFRS）
AR23 = EDGAR + "000129281424000982/ex99-1.htm"           # 6-K 2024-03-25：2023 年合併財務報告（英譯；FY2022 比較欄，TIFRS）
XBRL = "https://data.sec.gov/api/xbrl/companyfacts/CIK0001033767.json"
SRC_H25 = f"FY2025、FY2024：公司 2025 年合併財務報告（6-K 2026-02-25，{AR25}）綜合損益表，千元 ÷ 1,000,000；FY2024 取該報告的比較欄（TIFRS）"
SRC_H23 = (f"FY2023：公司 2024 年合併財務報告比較欄（6-K 2025-03-20，accession 0001292814-25-000981，{AR24}）；FY2022：公司 2023 年合併財務報告比較欄（6-K 2024-03-25，accession 0001292814-24-000982，{AR23}）；"
           "皆為 TIFRS 綜合損益表，千元 ÷ 1,000,000（2024 年報與 2023 年報的 FY2023 欄相同）")

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


# ---------- RevDCF：反向 DCF（樣板 台積電 v0.5.5；聯電 FY2027 不動、FY2028 起反解） ----------
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
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.4 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
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
    kv("px", "現價（新台幣）", "=px", "#,##0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "#,##0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（情境 B；新台幣十億元）", "=" + ncr, "#,##0.0", "＝淨現金與投資 ＋ stub 期 FCFF − stub 期已宣告股利（Val 頁 0.1 節；同 DCF 頁每股價值公式）")
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
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
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
                     T(" 倍（同業中位數 × 聯電過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給聯電的倍數是同業的 "),
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
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.2.4 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.00"), T(" 新台幣；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "In 頁市場共識目標價（stockanalysis 分析師平均；Interested-party）")]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = SP; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=In!$B$" + str(lab(wb["In"], "px", exact=True))); r += 1
    return r0


# ---------- Sens：加權目標價相對基準（連 Sens 頁既有逐案整本重算結果） ----------
SENS_LAB = {"fx": "匯率", "cap_adj": "產能係數（擴產提前／延後 ±5%）", "ut_adj": "稼動率 ±3 個百分點", "asp_adj": "售價水準 ±3%", "dag_adj": "折舊年增 ±5 個百分點",
            "cc_g": "每片現金成本年增 0%／4%", "phi": "營業費用彈性", "nonop_f": "營業外淨額", "tax": "有效稅率", "ai_adj": "AI 相關營收 ±30%", "m_new": "新業務邊際貢獻率"}
FD_PAIRS = [("折現率 WACC −1／＋1 個百分點", "WACC −1pt", "WACC ＋1pt"), ("β（產業 β ∓ 0.3 重新加槓桿）", "β：低", "β：高"), ("股權風險溢酬 4%／6%", "ERP：低", "ERP：高"),
            ("永續成長率", "永續成長率 g：低", "永續成長率 g：高"), ("永續期新投資報酬率＝WACC＋5 個百分點（只有高案）", None, "RONIC＝WACC＋5pt"),
            ("資本支出占營收 −1／＋1 個百分點", "資本支出 ÷ 營收 −1pt", "資本支出 ÷ 營收 ＋1pt"), ("營運資金占營收", "營運資金 ÷ 營收：低", "營運資金 ÷ 營收：高"),
            ("淨現金不含長期投資（只有低案）", "淨現金不含長期投資", None), ("終值改用自由現金流成長法（只有低案）", "終值改用 FCFF ×（1＋g）", None),
            ("美元口徑折現（只有低案）", "美元口徑折現", None), ("DCF 權重 0.3／0.7", "DCF 權重 0.3", "DCF 權重 0.7")]


def sens_tor(wb, G):
    SE = wb["Sens"]; V = wb["Val"]
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.2.4 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務案例連本頁上方逐案整本重算結果（相對基準欄），目標價參數取上方 DCF 參數案例；白話標籤，不回饋計算")
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
    SE.cell(r, 4, "Sens 頁 Wattonomics 錨定案例（既有列；AI 相關營收 FY2027 起 × 錨定係數，三情境同乘）"); r += 1
    last = r - 1
    SE.cell(r, 1, "核對：基準列相對基準（應為 0）"); SE.cell(r, 2, "=J4").number_format = "0.000000"
    return dict(t0=r0, first=first, last=last, state=r)


# ---------- Hist：合併損益表歷史（新台幣十億元） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)
HIS = [
    ("rev", "營業收入（Operating revenues）", 278.705264, 222.533000, 232.302584, 237.553199),
    ("cogs", "營業成本（Operating costs）", 152.940887, 144.789162, 156.648504, 168.646700),
    ("gp", "營業毛利（Gross profit）", 125.764377, 77.743838, 75.654080, 68.906499),
    ("oi", "營業利益（Operating income；含其他營業收益及費損淨額）", 104.292225, 57.890661, 51.612570, 43.948688),
    ("pti", "稅前淨利（Income before income tax）", 106.097237, 70.912228, 56.219713, 49.647706),
    ("tax", "所得稅費用（Income tax expense）", 18.078986, 9.472411, 9.113457, 8.112958),
    ("ni", "本期淨利（合併，含非控制權益）", 88.018251, 61.439817, 47.106256, 41.534748),
    ("nip", "歸屬母公司業主淨利（模型淨利口徑）", 87.198291, 60.989633, 47.210930, 41.716249),
]
EPSR = [("eps", "EPS 基本（公司公告；新台幣元）", (7.09, 4.93, 3.80, 3.34)),
        ("eps_d", "EPS 稀釋（公司公告；新台幣元）", (6.87, 4.83, 3.74, 3.31)),
        ("sh", "基本加權平均股數（十億股）", (12.305519, 12.370951, 12.436599, 12.485461))]
# 各年準則與出處（v0.2.4 修正：原 FY2022–23 取 20-F〔IASB 版 IFRS〕，與 FY2024–25 的 TIFRS 混用；改為四年皆 TIFRS）
STD = ("TIFRS（6-K 2023 年報）", "TIFRS（6-K 2024 年報）", "TIFRS（6-K 2025 年報）", "TIFRS（6-K 2025 年報）")


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 合併損益表歷史（v0.2.4 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度＝日曆年。金額：新台幣十億元（TIFRS 合併，四年皆取公司 6-K 年度合併財務報告，同一準則）；EPS：新台幣元。"
               "營業利益為公司「營業利益」（已含其他營業收益及費損淨額），標準損益表的營業費用＝毛利 − 營業利益")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    H.cell(r, 1, "h_std"); H.cell(r, 2, "會計準則（出處）")
    for j, v in enumerate(STD): H.cell(r, 3 + j, v)
    H.cell(r, 7, "Verified"); H.cell(r, 8, "TIFRS＝台灣金管會認可之 IFRS（公司 6-K 年度合併財務報告英譯本）；20-F 為 IASB 版 IFRS，數字不同，不混用（見頁末口徑註記）")
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
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, ("各年年度合併財務報告附註 6(27)（千股；FY2025、FY2024 取 2025 年報，FY2023 取 2024 年報，FY2022 取 2023 年報）" if k == "sh" else "同上（綜合損益表每股盈餘）"))
        R[k] = r; r += 1
    for k, labl, f_, note in (("nonop", "營業外收入及支出淨額（＝稅前 − 營業利益）", "={c}{pti}-{c}{oi}", "與財報「營業外收入及支出合計」相同（FY2022–FY2025：1.805、13.022、4.607、5.699）"),
                              ("opx", "營業費用（含其他營業收益及費損淨額；＝毛利 − 營業利益）", "={c}{gp}-{c}{oi}", "FY2025：營業費用小計 26.948 − 其他營業收益淨額 1.990；FY2024：25.365 − 1.324"),
                              ("nci", "非控制權益損益（＝本期淨利 − 歸屬母公司淨利）", "={c}{ni}-{c}{nip}", "—")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "#,##0.000"
        H.cell(r, 7, "Derived"); H.cell(r, 8, note); R[k] = r; r += 1
    r += 1
    H.cell(r, 1, "核對（新台幣十億元；千元換算可有 ±0.001 捨入差）").font = BOLD; r += 1
    R["chk0"] = r
    for labl, f_ in (("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"), ("稅前淨利 − 所得稅 − 本期淨利", "={c}{pti}-{c}{tax}-{c}{ni}"),
                     ("歸屬母公司淨利 ÷ 基本股數 − 基本 EPS（EPS 四捨五入到分）", "={c}{nip}/{c}{sh}-{c}{eps}")):
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    H.cell(r, 2, "FY2025：本表 − Q 頁 FY2025A 錨點（營收、毛利、營業利益、歸屬母公司淨利；Q 頁營業利益 43.850 取自資料商，財報為 43.949，模型既有值不改）")
    for c, k, nm in (("C", "rev", "a25_rev"), ("D", "gp", "a25_gp"), ("E", "oi", "a25_oi"), ("F", "nip", "a25_ni")):
        H[f"{c}{r}"] = f"=F{R[k]}-{nm}"; H[f"{c}{r}"].number_format = "0.000"
    R["chk1"] = r; r += 2
    H.cell(r, 1, "口徑註記：四年皆取公司 6-K 年度合併財務報告（TIFRS），同一準則，年增率不跨準則。20-F 原報數不同，原因為準則不同（20-F 為 IASB 版 IFRS；6-K 年報為 TIFRS），非重編："
                  "FY2024 20-F 稅前 57.045、歸屬母公司 48.779、基本 EPS 4.01、股數 121.73 億股（本表 56.220、47.211、3.80、124.37 億股）；FY2023 20-F 稅前 68.450、歸屬母公司 59.689、EPS 4.92（本表 70.912、60.990、4.93）；"
                  "FY2022 20-F 稅前 109.626、歸屬母公司 89.479、EPS 7.40（本表 106.097、87.198、7.09）。營收、毛利、營業利益兩準則相同。")
    R["note"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 60
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表（新台幣十億元） ----------
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
    P.cell(r, 1, "標準損益表（v0.2.4 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加 4Q26 指引的錨點）。新台幣十億元；EPS 新台幣元。"
                 "只引用本頁上方情境 B 推導區塊、Hist、DCF 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（TIFRS 合併財報；歷史三年 FY2023A–FY2025A 與 FY2022 基期皆取公司 6-K 年度合併財報，同一準則，不混用 20-F 的 IASB 版 IFRS）。聯電不另發布 Non-GAAP，法說與市場共識 EPS 皆以此口徑，本表不做 Non-GAAP 調整。稅後淨利＝歸屬母公司淨利（模型淨利口徑）；EPS 為基本 EPS（歷史為公司公告值）。"
                 "營業費用含其他營業收益及費損淨額（＝毛利 − 營業利益）；其他收入（淨額）＝稅前利益 − 營業利益。FY2027E 起非控制權益假設為 0（模型既有假設）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（＝淨利 ÷（1 − 稅率）− 營業外淨額，由淨利還原）；毛利率與營業費用率沿用前一年據以算出營業成本、毛利與營業費用，"
                 "與延伸期營業利益的差額列核對列，不改淨利與目標價。")
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
    row("rev", "營收", "NT$B", hh("rev") + eb("rev", "FGHIJKL"), M, "FY2026E–FY2029E＝PL 頁情境 B 推導區塊（晶圓＋AI 相關＋Intel 12 奈米）；FY2030E 起為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{g['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{g['rev']}/{p}{g['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 取 Hist 頁")
    row("cogs", "營業成本（COGS）", "NT$B", hh("cogs") + [f"={c}{g['rev']}-{c}{g['gp']}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "NT$B", hh("gp") + eb("gp") + [f"={c}{g['rev']}*{c}{g['gm']}" for c in "JKL"], M, "FY2026E–FY2029E＝PL 頁毛利；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{g['gp']}/{c}{g['rev']}" for c in "CDEFGHI"] + [f"={p}{g['gm']}" for p in "IJK"], PCT, "FY2030E–FY2032E＝前一年")
    row("opx", "營業費用（含其他營業損益）", "NT$B", [f"=Hist!{HC[c]}{HR['opx']}" for c in "CDE"] + eb("opx")
        + [f"={c}{g['rev']}*{p}{g['opx']}/{p}{g['rev']}" for p, c in (("I", "J"), ("J", "K"), ("K", "L"))], M, "歷史＝毛利 − 營業利益；FY2026E–FY2029E＝PL 頁營業費用；FY2030E 起＝營收 × 前一年營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + eb("oi") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M, "FY2030E–FY2032E＝DCF 頁情境 B EBIT（淨利還原）")
    row("opm", "營業利益率", "%", [f"={c}{g['oi']}/{c}{g['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外）", "NT$B", [f"={c}{g['pti']}-{c}{g['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（利息、投資損益、權益法損益、匯兌）；FY2026E 含第二季一次性投資收益")
    row("pti", "稅前利益", "NT$B", hh("pti") + eb("pti") + [f"={c}{g['oi']}+nonop_f" for c in "JKL"], M, "FY2030E 起＝延伸期營業利益＋營業外淨額（In 頁，與 DCF 還原同式）")
    row("tax", "所得稅", "NT$B", hh("tax") + eb("tax") + [f"={c}{g['pti']}*tax" for c in "JKL"], M, "FY2030E 起＝稅前 × 有效稅率（In 頁）")
    row("nci", "非控制權益損益", "NT$B", [f"=Hist!{HC[c]}{HR['nci']}" for c in "CDE"] + eb("mi") + [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['pti'] + 3}" for c in "JKL"], M,
        "FY2027E 起模型假設 0；FY2030E 起＝稅前 − 所得稅 − 稅後淨利（應為 0）")
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
    row2("c5", "核對：營收 − DCF 頁情境 B 營收（FY2027E–FY2032E）／ − Q 頁 FY2026 錨點（FY2026E）", "NT$B", [None] * 3 + [f"=F{g['rev']}-a26_rev"] + [f"={c}{g['rev']}-DCF!{d}{d_rev}" for c, d in zip("GHIJKL", "CDEFGH")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度＝日曆年，FY(t) 直接對應 CY(t)；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2032E 由 FY2029 之後的收斂路徑推得，沿用 CY2029 旗標。FY2026 不是對帳中年份。")
    L["ycaveat"] = r; r += 2
    # ---- 分部營收推導（情境 B；HTML 命題、驅動因子章用） ----
    P.cell(r, 1, "分部營收推導（情境 B；v0.2.4）：成熟製程晶圓＝產能 × 稼動率 × 美元售價 × 匯率；AI 相關（先進封裝矽中介層、矽光子）＝FY2026 公司說法 × 年增，FY2028 起乘 Siliconomics 顆數係數；Intel 12 奈米為假設路徑；三者相加＝營收。只引用本頁上方情境 B 推導區塊").font = BOLD; r += 1
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
    seg("s_cap", "產能（12 吋約當）", "千片", pb("cap"), "#,##0", "FY2026 由各季出貨 ÷ 稼動率推得；FY2027 起依已核准擴產年增")
    seg("s_ut", "平均稼動率", "%", pb("ut"), PCT)
    seg("s_ship", "晶圓出貨（12 吋約當）", "千片", pb("ship"), "#,##0")
    seg("s_asp", "平均售價（成熟製程晶圓）", "US$/片", pb("asp"), "#,##0")
    seg("s_fx", "匯率", "NT$/US$", pb("fxr"), "0.00")
    seg("s_w", "成熟製程晶圓營收", "NT$B", pb("wrev"), M)
    seg("s_aiu", "AI 相關營收（美元）", "US$B", pb("ai_usd"), "0.000")
    seg("s_ai", "AI 相關營收", "NT$B", pb("ai_rev"), M)
    seg("s_int", "Intel 12 奈米營收", "NT$B", pb("int_rev"), M)
    seg("s_rev", "營收", "NT$B", pb("rev"), M)
    seg("s_chk", "核對：晶圓＋AI 相關＋Intel 12 奈米 − 營收（應為 0）", "NT$B", lambda c: f"={c}{L['s_w']}+{c}{L['s_ai']}+{c}{L['s_int']}-{c}{L['s_rev']}", "0.000")
    seg("s_sh", "AI 相關與 Intel 12 奈米 ÷ 營收", "%", lambda c: f"=({c}{L['s_ai']}+{c}{L['s_int']})/{c}{L['s_rev']}", PCT)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0)
    return L


# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR):
    CW = wb["Coord"]; N = SIL.NAMES
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.2.4；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（匯率 fx_twd；顆數情境係數 SIF_UnitsCoef_Base；Wattonomics 錨定 SIF_WTAnchor_Base；對帳中旗標與另一側比值）", SIL.VERSION,
           "聯電屬硬體供應鏈個股，讀 Siliconomics；不讀 Wattonomics 覆蓋 Siliconomics（C4）"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁 Wattonomics 錨定案例；不改基準"),
          ("匯率（FY2027 起）", "=fx", "In 頁 fx（Siliconomics fx_twd）", SIL.VERSION, "美元售價換算新台幣"),
          ("顆數情境係數 情境 B（CY2028）", "=um28_B", "In 頁 um28_B（Siliconomics SIF_UnitsCoef_Base；情境 A、C 讀 markets/scen）", SIL.VERSION, "AI 相關營收 FY2028 起乘同字母係數"),
          ("顆數情境係數 情境 B（CY2029）", "=um29_B", "In 頁 um29_B（同上）", SIL.VERSION, "同上"),
          ("公司自有變數（不由層級提供）", "—", "產能、稼動率、美元售價、每片現金成本、折舊、AI 相關營收起點與年增、Intel 12 奈米、營業費用、業外、稅率", "本模型",
           "Siliconomics 尚無成熟製程代工的稼動率、價格或市場規模（模型標「待併入 Siliconomics」）")]
    for y in (2027, 2028, 2029):
        rd.append((f"Wattonomics 錨定係數 CY{y}（敏感度案例用）", f"=wta_{y % 100}", "Siliconomics SIF_WTAnchor_Base（R4 比值）", SIL.VERSION, "Sens 頁「Wattonomics 錨定」案例；基準不乘"))
    rd.append(("對帳中旗標與另一側讀數比值（CY2026–CY2029）", "見本頁上方", "Siliconomics SIF_RecoFlag_<年>、SIF_DeployGW_AltRatio_<年>", SIL.VERSION, "兩讀數目標價（C4）"))
    ROW["rd0"] = r
    for it in rd:
        for j, v in enumerate(it, 1):
            c = CW.cell(r, j, v)
            if j == 2 and isinstance(v, str) and v.startswith("="): c.number_format = "0.000"
        r += 1
    r += 1
    CW.cell(r, 1, "2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0，無檢驗力，故不判定、不計入對帳中；C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）").font = BOLD; r += 1
    CW.cell(r, 1, "本模型會計年度＝日曆年，FY(t) 對 CY(t) 直接對應。下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。")
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
    line("新台幣兌美元（2027 起）", "NT$/US$", "CY2027", "=fx", SIL.base("fx_twd"), "0.00", f"{nb} fx_twd")
    line("顆數情境係數 情境 B", "x", "CY2028", "=um28_B", N["SIF_UnitsCoef_Base"][3], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um28_B（本版不改）")
    line("顆數情境係數 情境 B", "x", "CY2029", "=um29_B", N["SIF_UnitsCoef_Base"][4], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um29_B（本版不改）")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 2
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是聯電自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。聯電只有 AI 相關營收（約占營收 3–7%）隨顆數變動。"); r += 1
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
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入，個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數建置：情境 B、C 於 CY2027 起 AI 相關營收（先進封裝矽中介層、矽光子）同乘比值，情境 A 不縮放；成熟製程晶圓、Intel 12 奈米與其他假設不變，整本重建（tools/coord/alt_run.py）。"
                      "AI 相關營收只占營收數個百分點，兩讀數差距小。本模型會計年度與日曆年相同，無月數近似；敏感度頁另有 Wattonomics 錨定案例（用 SIF_WTAnchor_Base，與本比值不同）。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110
    r += 1
    # ---- 4. C0（實績）----
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.2.4；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（2026 年 1–10 月月營收、1Q26–3Q26 季報、3Q26 法說揭露的出貨與稼動率、前三季資本支出；不用指引或預測）。層級側＝Siliconomics 現行基準的同口徑隱含值。"
                  "聯電讀入的層級值只有匯率、CY2028 起的 AI 加速器顆數係數與 Wattonomics 錨定係數，層級沒有成熟製程代工的營收、出貨、稼動率或價格，也沒有聯電的 AI 相關營收或資本支出口徑。"
                  "主讀數（已公布期間累計 vs 層級 CY2026 × 月數 ÷ 12）與第二讀數（最近一季 × 4 vs 層級 CY2026）都需要層級的同口徑年度值，本輪各項皆無，故不計算。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（結論）：本表沒有任何一列同時具備「公司已公布實績」與「獨立於本公司模型的層級隱含值」。本模型 CY2026 的 AI 相關營收（約 3 億美元）是公司說法（媒體轉述），不是層級值；"
                  "CY2028 起的顆數係數只是比例，乘的是本公司的 AI 相關營收路徑，因此層級對聯電沒有可對照的絕對量。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "公司揭露缺口與層級缺口：公司未揭露 AI 相關（先進封裝矽中介層、矽光子）營收金額、產能與平均售價絕對值；Siliconomics 尚無成熟製程代工行業層（稼動率、價格、市場規模），"
                  "設備層只有「其他代工與邏輯」群組設備成長率（含中芯、Intel、三星晶圓代工、華虹、聯電、GlobalFoundries 等），不是聯電口徑。以上實績取自公司 6-K（SEC EDGAR 原文）。")
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
    put("q1", "合併營收：1Q26", "NT$B", "2026-01～03", "=Q!C4", "Verified", "Q 頁 1Q26（公司 6-K 1Q26 財報；連格）", "#,##0.000")
    put("q2", "合併營收：2Q26", "NT$B", "2026-04～06", "=Q!D4", "Verified", "Q 頁 2Q26（＝1H26 129.771 − 1Q26；公司 6-K 2Q26 財報；連格）", "#,##0.000")
    put("q3", "合併營收：3Q26", "NT$B", "2026-07～09", "=Q!E4", "Verified", "Q 頁 3Q26（連格）；公司 2026 年第三季合併財務報告（6-K 2026-10-28；網址見 Sources 頁）綜合損益表：3Q26 74,239,117 千元、前三季 204,009,681 千元", "#,##0.000")
    put("m10", "月營收：2026 年 10 月", "NT$B", "2026-10", 25.197532, "Verified", "公司 10 月營收公告（6-K 2026-11-05；網址見 Sources 頁）：Net sales 25,197,532 千元，年增 18.33%（2025 年 10 月 21,294,974 千元）", "#,##0.000")
    put("ytd", "月營收：2026 年 1–10 月累計", "NT$B", "2026-01～10", 229.207213, "Verified", "同上公告：Year-to-Date Net sales 229,207,213 千元，年增 16.33%（2025 年同期 197,038,628 千元）", "#,##0.000")
    put("xchk", "核對：1–10 月累計 − 1Q26 − 2Q26 − 3Q26 − 10 月（應約 0；月營收為合併自結數，季營收為財報）", "NT$B", "—", f"=D{r - 1}-D{r - 5}-D{r - 4}-D{r - 3}-D{r - 2}", "Derived", "差額為季營收輸入（百萬元四捨五入）與千元數的捨入差", "0.000")
    put("ship3", "晶圓出貨：3Q26（12 吋約當）", "千片", "2026-07～09", "=Q!E5", "公司揭露（下限）", "Q 頁 3Q26 出貨（公司 3Q26 法說用語「超過 120 萬片」，取下限 1,200 千片）", "#,##0")
    put("ut3", "產能利用率：3Q26", "%", "2026-07～09", "=Q!E6", "Verified", "Q 頁 3Q26 稼動率（公司 6-K 3Q26：92%）", "0%")
    put("cx9", "資本支出（購置不動產、廠房及設備）：2026 年前三季", "NT$B", "2026-01～09", 32.608565, "Verified", "公司 2026 年第三季合併財務報告現金流量表（6-K 2026-10-28；網址見 Sources 頁）：Acquisition of property, plant and equipment 32,608,565 千元", "#,##0.000")
    put("cx9p", "資本支出：2025 年前三季", "NT$B", "2025-01～09", 32.918719, "Verified", "同上比較欄：32,918,719 千元", "#,##0.000")
    put("cxg", "資本支出前三季年增", "%", "2026 vs 2025", f"=D{r - 2}/D{r - 1}-1", "Derived", "＝上兩列", "+0.0%;-0.0%")
    put("eqg", "參考：層級「其他代工與邏輯」設備成長率（群組，非聯電口徑）", "%/年", "Siliconomics 基準", SIL.base("eq_g_oth_fd"), "Assumed（層級標記）",
        f"Siliconomics {SIL.VERSION} 設備層「其他代工與邏輯設備成長率」（中芯、Intel、三星晶圓代工、華虹、聯電、GlobalFoundries 等群組）；群組設備成長率，不是聯電資本支出，只列參考", "0.0%")
    put("ai26", "參考：AI 相關營收 CY2026（本模型輸入＝公司說法，非實績、非層級值）", "US$B", "CY2026", "=ai_26", "Interested-party", "In 頁 FY2026 AI 相關營收（科技新報 2026-10-06 轉述摩根士丹利：公司說 2026 年約 3 億美元）；公司財報未揭露金額", "0.000")
    r += 1
    CW.cell(r, 1, "C0（實績）判定").font = BOLD; ROW["hdr_j"] = r; r += 1
    for j, h in enumerate(["項目", "", "", "", "", "", "判定（±15%）", "結論與缺漏揭露"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1
    J = [("C0（實績）判定：合併營收與月營收（NT$）", "無可比實績",
          "結論：無可比實績。公司營收實績完整（見上），但層級（Siliconomics、Wattonomics）沒有聯電營收或成熟製程代工市場規模；缺：Siliconomics 成熟製程代工行業層（市場規模與聯電份額）。不用本模型預測補位"),
         ("C0（實績）判定：晶圓出貨與產能利用率", "無可比實績",
          "結論：無可比實績。層級沒有成熟製程代工的出貨或稼動率；公司側只揭露「超過 120 萬片」與整數稼動率，缺產能絕對值。待 Siliconomics 建成熟製程稼動率與價格層後重做"),
         ("C0（實績）判定：AI 相關營收（先進封裝矽中介層、矽光子）", "無可比實績",
          "結論：無可比實績。公司未揭露 AI 相關營收金額（只有媒體轉述的 2026 年約 3 億美元、2029 年約 10 億美元目標）；層級對本項只提供 CY2028 起的顆數係數（比例），沒有聯電口徑的絕對量"),
         ("C0（實績）判定：資本支出", "無可比實績",
          "結論：無可比實績。公司前三季資本支出實績可得（年增見上），但層級只有「其他代工與邏輯」群組設備成長率，不是聯電口徑，只列參考，不判定")]
    ROW["j1"] = r
    for a, gd, h in J:
        CW.cell(r, 1, a); CW.cell(r, 7, gd); CW.cell(r, 8, h); r += 1
    ROW["j4"] = r - 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["j1"]}:G{ROW["j4"]},"對帳中")')
    ROW["cnt4"] = r; ROW["last"] = r; r += 1
    CW.cell(r, 1, "讀法：聯電 2026 年用前三季實績加 4Q 指引，營收、出貨、稼動率都是公司已公布實績，不構成對層級的檢驗；層級對聯電只約束匯率與 CY2028 起 AI 相關營收的顆數比例。因此本輪 C0（實績）沒有可判定項目，對帳中 0；"
                  "外部估計（賣方共識、券商目標）只在 Recon 頁列參考。")
    ROW["read"] = r; r += 1
    ROW["inp"] = inp
    return ROW


def readme13(wb, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_fx, s_asp, s_ut, s_ai = sens_r("匯率"), sens_r("售價水準"), sens_r("稼動率"), sens_r("AI 相關營收 ±30%")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.4 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0.0"), T(" 元，相對現價 "), TX(up, SP), T("，評等"), rec, T("：DCF "), TX(dcf, "#,##0.0"), T(" 元，同業比較法（相對倍數調整後）"), TX(cadj, "#,##0.0"), T(" 元。"))
    put("rec", T("評等「"), rec, T("」：同業比較法採聯電過去十年相對同業的本益比倍數（"), TX(coef, "0.00"), T(" 倍，觸及下限）；未調整的市場主流同業法 "), TX(cu, "#,##0.0"), T(" 元。"))
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0.0"), T(" 元，現價 "), TX("px", "#,##0.0"), T(" 元，差距 "), TX(up, SP), T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.00"), T(" 元）得 "), TX(cadj, "#,##0.0"), T(" 元，DCF 得 "), TX(dcf, "#,##0.0"), T(" 元，分析師共識 "), TX(cons, "#,##0.0"), T(" 元。"))
    put("o2", T("獲利由成熟製程的稼動率與美元售價決定：情境 B 的 FY2028 稼動率 "), TX(f"PL!H{PLA['s_ut']}", "0%"), T("、售價每片約 "), TX(f"PL!H{PLA['s_asp']}", "#,##0"),
        T(" 美元；AI 相關與 Intel 12 奈米合計只占營收 "), TX(f"PL!H{PLA['s_sh']}", "0.0%"), T("，兩者金額公司都未揭露。"))
    put("o3", T("最大變數（加權目標價相對基準）：匯率低／高案 "), TX(f"Sens!B{s_fx}", SP), T("／"), TX(f"Sens!C{s_fx}", SP),
        T("；售價 ±3% "), TX(f"Sens!B{s_asp}", SP), T("／"), TX(f"Sens!C{s_asp}", SP), T("；稼動率 ±3 個百分點 "), TX(f"Sens!B{s_ut}", SP), T("／"), TX(f"Sens!C{s_ut}", SP),
        T("；AI 相關營收 ±30% 只有 "), TX(f"Sens!B{s_ai}", SP), T("／"), TX(f"Sens!C{s_ai}", SP), T("。這些是測試案例，不是門檻。"))
    put("o4", T("反向 DCF：WACC "), TX("wacc", "0.0%"), T("，要使 DCF 等於現價，折現率須降到 "), TX(rv("w1"), "0.0%"), T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"),
        T(" 倍；現價相當於 12 個月後 EPS 的 "), TX(rv("pe1"), "0.0"), T(" 倍。"))
    put("o5", T("對帳中：CY2027–CY2029 的 AI 加速器顆數，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；聯電只有 AI 相關營收受影響，另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0.0"), T(" 元（"), f"Coord!C{CR['rec']}", T("），個股不裁決。公司實績（月營收、出貨、稼動率、資本支出）對層級：無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：成熟製程晶圓＝產能 × 稼動率 × 美元售價 × 匯率，加上 AI 相關與 Intel 12 奈米。FY2028 出貨 "),
        TX(f"PL!H{PLA['s_ship']}", "#,##0"), T(" 千片、晶圓營收 "), TX(f"PL!H{PLA['s_w']}", "#,##0"), T("、營收 "), TX(f"PL!H{PLA['s_rev']}", "#,##0"), T(" 十億新台幣。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 台積電 v0.5.5、NVIDIA）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2022–FY2025，四年皆取公司 6-K 年度合併財報，TIFRS 同一準則）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表與分部營收推導、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（連既有逐案整本重算）、Coord 頁（兩讀數目標價；讀取層級；C0：月營收、季營收、出貨、稼動率與資本支出實績對層級，皆無可比實績，對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異；um28_B／um29_B、wta 等名稱與值不動），加權目標價與評等不變。前版 README 數字照舊，本節為補充說明")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, G):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    PLA = pl_standard(wb, HR, CR, G["PB"]["B"])
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, G)
    CO = coord_c0(wb, SIL, CR)
    RMA = readme13(wb, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((
        ("公司：2026 年 10 月營收公告 6-K（2026-11-05；10 月 25,197,532 千元、1–10 月累計 229,207,213 千元；C0 實績用）", OCT26),
        ("公司：2026 年第三季合併財務報告 6-K（2026-10-28；3Q26 與前三季損益、前三季資本支出；C0 實績用）", Q326),
        ("公司：2025 年合併財務報告 6-K（2026-02-25；FY2025、FY2024 綜合損益表與每股盈餘附註；Hist 頁用）", AR25),
        ("公司：2024 年合併財務報告 6-K（2025-03-20；FY2023 比較欄，TIFRS；Hist 頁用）", AR24),
        ("公司：2023 年合併財務報告 6-K（2024-03-25；FY2022 比較欄，TIFRS；Hist 頁用）", AR23),
        ("公司：20-F XBRL（IASB 版 IFRS；FY2022–FY2024 原報數，只供 Hist 頁口徑註記對照）", XBRL),
        ("公司：20-F 清單（FY2024 20-F accession 0001193125-25-092142）", F24))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v024 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v024_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 相對基準 {SNS['t0']} 起；Coord C0 {CO['rd0']}–{CO['last']}；README v1.3 {len(RMA)} 列"
