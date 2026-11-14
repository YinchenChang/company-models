# 台積電 v0.5.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數（Coord 頁，tools/coord）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.5.4 逐格比對驗證；In 頁既有列與 um28／um29、wta 等名稱與值不動）：
#   Hist 頁（新）：合併損益表 FY2022–FY2025（公司年報，IFRS，新台幣十億元）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，新台幣十億元）、比率區、核對列、分部營收推導（情境 B，十億美元）
#   RevDCF 頁（新）：反向 DCF
#   Val 頁末：0.7 估值並列四個目標價；Sens 頁末：加權目標價相對基準（整本重算快照 tsmc/sens.json）
#   Coord 頁末：三層協調——讀取層級、讀取層級核對（同源，不判定）、層級之間（參考）、對帳中年份、C0（實績）
#   Sources 頁：新增來源；README 頁末：v1.3 展示層白話列與本版說明
import json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.5.5", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
AR25 = "https://investor.tsmc.com/static/annualReports/2025/chinese/pdf/2025_tsmc_ar_c_ch6.pdf"
FS25 = "https://investor.tsmc.com/japanese/encrypt/files/encrypt_file/reports/2026-02/cb73d5f4e019a8f6d7a494a0f8f6c6da2dfc4ee2/FS.pdf"
AR23 = "https://investor.tsmc.com/static/annualReports/2023/english/pdf/2023_tsmc_ar_e_ch6.pdf"
REV10 = "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000796/tsm-revenue20261110.htm"
PR10 = "https://pr.tsmc.com/english/news/3349"
SRC_H = (f"公司合併損益表（IFRS，新台幣）：FY2025、FY2024 取 2025 年報第六章（{AR25}）與 2025 財報 FS（{FS25}）；FY2023、FY2022 取 2023 年報第六章（{AR23}）；"
         "千元 ÷ 1,000,000 換算十億元。摘錄經檢索工具讀取原文 PDF，營業利益、營業外、稅前、所得稅、淨利五項另與 NI_OI 頁（前版獨立來源）核對")


ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

#HELP
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
TX = lambda ref, f: f'TEXT({ref},"{f}")'
REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'


#REV
# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）", exact=True); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty, tn = 6, 7, 8   # DCF 頁：評價日後占比、折現年數（期中）、終值折現年數
    assert D.cell(fr, 1).value.startswith("評價日後占") and D.cell(ty, 1).value.startswith("折現年數") and D.cell(tn, 1).value.startswith("終值折現年數"), (D.cell(fr, 1).value, D.cell(ty, 1).value, D.cell(tn, 1).value)
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    ncr = f"Val!$C${lab(V, '淨現金滾到評價日')}"   # 淨現金滾到評價日（含 stub 期 FCFF 與已宣告股利；三情境相同，取情境 B 欄）
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.5.5 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「12 個月後前瞻 EPS（情境加權）× 倍數＝現價」。FY2027（第一個明確期年度，三情境接近）不動，營業利益率與營收倍數只動 FY2028–FY2032。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推")
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
    kv("sh", "評價用稀釋股數（百萬股）", "=sh_v", "0.000")
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
    NC = "BCDEFG"   # 本頁 FY2027–FY2032 欄；對應 DCF 頁 C–H 欄

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, NC[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（新台幣十億元）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（新台幣十億元；事業別稅前利益合計）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
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
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
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
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次；若隱含值貼近下界，表示區間內無解）").font = BOLD; r += 1
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
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    RV.cell(r, 1, "12 個月後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), TX("wacc", "0.0%"), T(" 的折現率，每股價值 "), TX(f"B{A['v_b']}", "#,##0"),
                     T(" 元；要支持現價 "), TX("px", "#,##0"), T(" 元，企業價值須放大為 "), TX(f"B{A['evx']}", "0.00"),
                     T(" 倍。單靠一個因素達成時，需要：折現率降到 "), TX(f"B{A['w1']}", "0.0%"), T("；或 FY2028 起現金流（營收與成本同步）全面放大 "),
                     TX(f"B{A['s1']}", "0.00"), T(" 倍（FY2028 營收 "), TX(f"B{A['rv1']}", "#,##0"), T(" 十億新台幣，模型為 "), TX(f"B{A['rv0']}", "#,##0"), T("）；或營業利益率提高到 "),
                     TX(f"B{A['m1']}", "0%"), T("（模型等效約 "), TX(f"B{A['m0']}", "0%"), T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), TX("px", "#,##0"), T(" 元相當於 12 個月後前瞻 EPS 的 "), TX(f"B{A['pe1']}", "0.0"),
                     T(" 倍，"), f'IF(B{A["pe1"]}>{pmed},"高於","低於")', T("同業中位數 "), TX(pmed, "0.0"), T(" 倍；模型採用 "), TX(f"C{A['pe1']}", "0.0"),
                     T(" 倍（同業中位數 × 台積電過去十年相對同業的倍數比 "), TX(coef, "0.00"), T("）。換言之，市場目前給台積電的倍數是同業的 "),
                     TX(f"B{A['rc1']}", "0.00"), T(" 倍，"), f'IF(B{A["rc1"]}>{coef},"高於","低於")', T("歷史相對倍數 "), TX(coef, "0.00"), T(" 倍。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


#VAL
# ---------- Val：估值並列四個目標價（版面規格 v1.3 G 節；只新增列） ----------
def val_g(wb):
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p, exact=True)}"
    dcf, cu, cadj, tp, cons = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）"), vb("目標價（DCF 與 Comparable 加權）"), vb("市場共識目標價")
    feps, pmed, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("相對倍數調整係數")
    cons_src = f"Val!$E${lab(V, '市場共識目標價', exact=True)}"
    px_lab = "In!$B$" + str(lab(wb["In"], "px", exact=True, col=1))
    r = V.max_row + 3
    V.cell(r, 1, "0.7 估值並列：四個目標價與現價").font = BOLD; V.cell(r, 4, "v0.5.5 起；個股 HTML 版面規格 v1.3 G 節；只引用上方各格，不回饋計算"); r0 = r; r += 1
    for j, h in enumerate(["目標價", "新台幣", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = [("① DCF 目標價（單純 DCF，情境加權）", dcf, F(T("三情境 DCF 價值（推到 12 個月後）以情境權重 A "), TX("tp_wA", "0.0"), T("／B "), TX("tp_wB", "0.0"), T("／C "), TX("tp_wC", "0.0"), T(" 加權；標準版 DCF（無槓桿自由現金流 FCFF＋WACC）"))),
            ("② Comparable 目標價（情境加權；相對倍數調整後，進加權）", cadj, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 相對倍數調整係數 "), TX(coef, "0.00"), T(" × 12 個月後前瞻 EPS（情境加權）"), TX(feps, "0.0"), T(" 新台幣；進加權目標價"))),
            ("②' Comparable 目標價（情境加權；未調整，市場主流同業法，不進加權）", cu, F(T("同業前瞻本益比中位數 "), TX(pmed, "0.0"), T(" 倍 × 12 個月後前瞻 EPS（情境加權）；只供對照，不進加權目標價"))),
            ("③ 加權目標價（DCF 與 Comparable 加權；橫幅主值，評等依此判定）", tp, F(T("DCF "), TX("tp_w_dcf", "0%"), T("＋Comparable（相對倍數調整後）"), TX("1-tp_w_dcf", "0%"))),
            ("④ 市場共識目標價", cons, "=" + cons_src)]
    for labl, ref, note in rows:
        V.cell(r, 1, labl); V.cell(r, 2, "=" + ref).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, note); r += 1
    V.cell(r, 1, "現價（股價基準日收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "=" + px_lab); r += 1
    return r0


#SENS
# ---------- Sens：加權目標價相對基準 ----------
def sens_tor(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    snap = json.load(open(os.path.join(here, "sens.json"), encoding="utf-8"))
    tp_cell = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）', exact=True)}"
    t0 = lab(SE, "目標價參數敏感度"); hr = t0 + 1
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用）").font = BOLD
    SE.cell(r, 4, "v0.5.5 起，版面規格 v1.3 H 節：一次動一個輸入，三情境同時受影響。業務主因為整本重算快照（tsmc/sens_tsmc.py），目標價參數取上方「目標價參數敏感度」表；白話標籤，不回饋計算")
    r0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "設定／來源", "低案：加權目標價（新台幣）", "高案：加權目標價（新台幣）", "快照基準加權目標價（新台幣）"], 1):
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
    wr = lab(SE, "Wattonomics 錨定（情境 B × wta）", exact=True)
    SE.cell(r, 1, "產業量依 Wattonomics 錨定（情境 B × 錨定係數；只有低案）"); SE.cell(r, 4, "Sens 頁 Wattonomics 錨定案例（既有列，目標價相對基準）")
    SE.cell(r, 2, f"=Sens!I{wr}").number_format = "+0.0%;-0.0%"; r += 1
    last = r - 1
    SE.cell(r, 1, "快照狀態（快照基準加權目標價 vs Val 頁；不同＝STALE，改輸入後須重跑 tsmc/sens_tsmc.py）")
    SE.cell(r, 2, f'=IF(ABS({tp_cell}-G{first})<=0.000001*ABS(G{first}),"CURRENT","STALE")')
    return dict(t0=r0, first=first, last=last, state=r)




# ---------- Hist：合併損益表歷史（新台幣十億元） ----------
# (鍵, 項目, FY2022, FY2023, FY2024, FY2025)
HIS = [
    ("rev", "營業收入淨額（Net revenue）", 2263.891292, 2161.735841, 2894.307699, 3809.054272),
    ("cogs", "營業成本（Cost of revenue）", 915.536486, 986.625213, 1269.954135, 1527.760293),
    ("gp", "營業毛利（Gross profit）", 1348.354806, 1175.110628, 1624.353564, 2281.293979),
    ("opx", "營業費用合計（研發、管理與行銷；Operating expenses）", 226.707552, 253.833716, 301.070315, 345.649630),
    ("oox", "其他營業收益及費損淨額（正值＝收益）", -0.368403, 0.188694, -1.230199, 0.447328),
    ("oi", "營業淨利（Income from operations）", 1121.278851, 921.465606, 1322.053050, 1936.091677),
    ("nonop", "營業外收入及支出合計（Non-operating, net）", 22.911867, 57.705718, 83.785585, 105.571163),
    ("pti", "稅前淨利（Income before income tax）", 1144.190718, 979.171324, 1405.838635, 2041.662840),
    ("tax", "所得稅費用（Income tax expense）", 127.290203, 141.403807, 233.406876, 326.266060),
    ("ni", "本年度淨利（合併，含非控制權益；模型淨利口徑）", 1016.900515, 837.767517, 1172.431759, 1715.396780),
    ("nip", "歸屬於母公司業主淨利", 1016.530249, 838.497664, 1173.267703, 1717.882627),
]
EPSR = [("eps_d", "EPS 稀釋（公司公告值；新台幣元）", (39.20, 32.34, 45.25, 66.25)),
        ("sh_d", "稀釋加權平均股數（十億股）", (None, None, 25.930, 25.931))]


def hist_is(wb):
    H = wb.create_sheet("Hist")
    H["A1"] = ("Hist — 合併損益表歷史（v0.5.5 起；PL 頁「標準損益表」的歷史三年與 FY2023 年增率用）。會計年度＝日曆年。金額：新台幣十億元（IFRS，合併）；EPS：新台幣元。"
               "營業費用列為公司「營業費用」合計，另列「其他營業收益及費損淨額」，標準損益表的營業費用＝兩者相抵（毛利 − 營業費用 ＝ 營業利益）")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4; R = {}
    for k, labl, *vals in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            c = H.cell(r, 3 + j, v); c.number_format = "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified"); H.cell(r, 8, SRC_H)
        R[k] = r; r += 1
    for k, labl, vals in EPSR:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, v in enumerate(vals):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.number_format = "0.00" if k == "eps_d" else "#,##0.000"; c.font = BLUE
        H.cell(r, 7, "Verified" if k == "eps_d" else "Verified（FY2022、FY2023 年報未列，不填）"); H.cell(r, 8, SRC_H + ("；FY2022、FY2023 年報為基本／稀釋同值" if k == "eps_d" else ""))
        R[k] = r; r += 1
    H.cell(r, 1, "h_nci"); H.cell(r, 2, "非控制權益損益（＝本年度淨利 − 歸屬母公司淨利）")
    for c in "CDEF":
        H[f"{c}{r}"] = f"={c}{R['ni']}-{c}{R['nip']}"; H[f"{c}{r}"].number_format = "#,##0.000"
    H.cell(r, 7, "Derived"); R["nci"] = r; r += 2
    H.cell(r, 1, "核對（新台幣十億元；千元換算可有 ±0.001 捨入差）").font = BOLD; r += 1
    chk = [("營收 − 營業成本 − 毛利", "={c}{rev}-{c}{cogs}-{c}{gp}"),
           ("毛利 − 營業費用合計 ＋ 其他營業收益 − 營業淨利", "={c}{gp}-{c}{opx}+{c}{oox}-{c}{oi}"),
           ("營業淨利 ＋ 營業外 − 稅前淨利", "={c}{oi}+{c}{nonop}-{c}{pti}"),
           ("稅前淨利 − 所得稅 − 本年度淨利", "={c}{pti}-{c}{tax}-{c}{ni}")]
    R["chk0"] = r
    for labl, f_ in chk:
        H.cell(r, 2, labl)
        for c in "CDEF":
            H[f"{c}{r}"] = f_.format(c=c, **R); H[f"{c}{r}"].number_format = "0.000"
        r += 1
    NI = wb["NI_OI"]
    for labl, col in (("營業淨利", "B"), ("營業外", "C"), ("稅前淨利", "D"), ("所得稅", "E"), ("本年度淨利", "F")):
        key = {"營業淨利": "oi", "營業外": "nonop", "稅前淨利": "pti", "所得稅": "tax", "本年度淨利": "ni"}[labl]
        H.cell(r, 2, f"核對：{labl}（本頁）− NI_OI 頁（前版獨立來源；兩位小數）")
        for c, nr in zip("CDEF", (4, 5, 6, 7)):
            H[f"{c}{r}"] = f"={c}{R[key]}-NI_OI!{col}{nr}"; H[f"{c}{r}"].number_format = "0.000"
        r += 1
    R["chk1"] = r - 1
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 60
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["H"].width = 100
    return R


# ---------- PL：標準損益表（新台幣十億元） ----------
YEARS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
COLS = "CDEFGHIJKL"
HC = {"C": "D", "D": "E", "E": "F"}
FLAGY = {"F": 2026, "G": 2027, "H": 2028, "I": 2029, "J": 2029, "K": 2029, "L": 2029}
PCOL = {"F": "B", "G": "F", "H": "G", "I": "H", "J": "O", "K": "P", "L": "Q"}   # 標準損益表欄 → PL 頁情境 B 欄（2026；2027–29 B；2030–32 B 收斂）
FXN = {"F": "fx_26", "G": "fx", "H": "fx", "I": "fx", "J": "fx", "K": "fx", "L": "fx"}
DCOL = {"J": "F", "K": "G", "L": "H"}   # FY2030–32 → DCF 頁欄
ETR = {"F": "NI_OI!$J$11", "G": "etr_new", "H": "etr_new", "I": "etr_new", "J": "etr_new", "K": "etr_new", "L": "etr_new"}


def pl_standard(wb, HR, CR, PR, ukeys):
    P = wb["PL"]
    D = wb["DCF"]; dbr = lab(D, "情境 B（基準）", exact=True); d_ebit = lab(D, "營業利益 EBIT", dbr); d_rev = lab(D, "營收", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, "標準損益表（v0.5.5 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋未來七年（FY2026–FY2032，情境 B 基準；FY2026E 為前三季實績加 4Q26 指引中點的錨點）。新台幣十億元；EPS 新台幣元。"
                 "只引用本頁上方、Hist、DCF、NI_OI 頁與 In 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：IFRS（公司合併財報；GAAP 口徑）。稅後淨利＝合併本年度淨利（含非控制權益，模型淨利口徑，EPS ＝ 淨利 × 匯率 ÷ 股數）；歷史 EPS 為公司公告稀釋 EPS（歸屬母公司）。"
                 "預測期模型只推營業利益與淨利：毛利、營業利益取 PL 頁情境 B（美元 × 匯率；FY2026 用 2026 平均匯率），營業費用＝毛利 − 營業利益；稅前利益＝淨利 ÷（1 − 有效稅率），"
                 "其他收入（淨額）＝稅前利益 − 營業利益，所得稅＝稅前利益 − 淨利（FY2026 稅率取前三季實績 NI_OI 頁，FY2027 起取 etr_new）。")
    r_note = r; r += 1
    P.cell(r, 1, "FY2030E–FY2032E（收斂法延伸期，只推營收與淨利）：營業利益取 DCF 頁情境 B 的 EBIT（淨利 ÷ ni_op 還原）；毛利率與營業費用率沿用前一年據以算出營業成本、毛利與營業費用，"
                 "與延伸期營業利益的差額列核對列，不改淨利與目標價。")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    for col in "FGHIJKL":
        yr = YEARS[COLS.index(col)]
        P[f"{col}{r}"] = f'="{yr}"&IF(Coord!${"BCDE"[FLAGY[col]-2026]}${CR["flag"]}=1,"（對帳中）","")'
    r_hdr = r; r += 1
    order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    RW = {k: r + i for i, k in enumerate(order)}
    M = "#,##0.0"; PCT = "0.0%"

    def row(key, label, unit, cells, fmt, note=""):
        P.cell(RW[key], 1, label); P.cell(RW[key], 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{RW[key]}"] = v; P[f"{col}{RW[key]}"].number_format = fmt
        if note: P.cell(RW[key], 13, note)
    g = RW
    hh = lambda k: [f"=Hist!{HC[c]}{HR[k]}" for c in "CDE"]
    pc = lambda k, cols: [f"=PL!{PCOL[c]}{PR[k]}*{FXN[c]}" for c in cols]
    row("rev", "營收", "NT$B", hh("rev") + pc("rev", "FGHIJKL"), M, "FY2026E＝PL 頁情境 B（美元）× 2026 平均匯率（＝前三季新台幣實績加 4Q26 指引中點）；FY2027E 起 × 匯率；FY2030E 起為收斂法")
    row("yoy", "營收年增率（YoY）", "%", [f"=C{g['rev']}/Hist!C{HR['rev']}-1"] + [f"={c}{g['rev']}/{p}{g['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    row("cogs", "營業成本（COGS）", "NT$B", hh("cogs") + [f"={c}{g['rev']}-{c}{g['gp']}" for c in "FGHIJKL"], M, "預測期＝營收 − 毛利")
    row("gp", "毛利", "NT$B", hh("gp") + pc("gp", "FGHI") + [f"={c}{g['rev']}*{c}{g['gm']}" for c in "JKL"], M, "FY2026E–FY2029E＝PL 頁毛利（美元）× 匯率；FY2030E 起＝營收 × 毛利率")
    row("gm", "毛利率", "%", [f"={c}{g['gp']}/{c}{g['rev']}" for c in "CDE"] + [f"=PL!{PCOL[c]}{PR['gm']}" for c in "FGHI"] + [f"={p}{g['gm']}" for p in "IJK"], PCT, "FY2026E–FY2029E 取 PL 頁毛利率；FY2030E–FY2032E＝前一年")
    row("opx", "營業費用（含其他營業損益）", "NT$B", [f"=Hist!{HC[c]}{HR['opx']}-Hist!{HC[c]}{HR['oox']}" for c in "CDE"] + [f"={c}{g['gp']}-{c}{g['oi']}" for c in "FGHI"]
        + [f"={c}{g['rev']}*{p}{g['opx']}/{p}{g['rev']}" for p, c in (("I", "J"), ("J", "K"), ("K", "L"))], M, "歷史＝營業費用合計 − 其他營業收益淨額；FY2026E–FY2029E＝毛利 − 營業利益；FY2030E 起＝營收 × 前一年營業費用率")
    row("oi", "營業利益", "NT$B", hh("oi") + pc("op", "FGHI") + [f"=DCF!{DCOL[c]}{d_ebit}" for c in "JKL"], M, "FY2026E–FY2029E＝PL 頁營業利益 × 匯率；FY2030E–FY2032E＝DCF 頁情境 B EBIT（淨利還原）")
    row("opm", "營業利益率", "%", [f"={c}{g['oi']}/{c}{g['rev']}" for c in COLS], PCT)
    row("oth", "其他收入（淨額；營業外）", "NT$B", [f"={c}{g['pti']}-{c}{g['oi']}" for c in COLS], M, "＝稅前利益 − 營業利益（利息、投資損益、匯兌、處分利益）；預測期為推算值")
    row("pti", "稅前利益", "NT$B", hh("pti") + [f"={c}{g['ni']}/(1-{ETR[c]})" for c in "FGHIJKL"], M, "預測期＝淨利 ÷（1 − 有效稅率）；FY2026E 稅率＝前三季實績（NI_OI 頁），FY2027E 起 etr_new")
    row("tax", "所得稅", "NT$B", hh("tax") + [f"={c}{g['pti']}-{c}{g['ni']}" for c in "FGHIJKL"], M)
    row("ni", "稅後淨利（合併，含非控制權益）", "NT$B", hh("ni") + pc("ni", "FGHIJKL"), M, "預測期＝PL 頁淨利（美元）× 匯率；非控制權益損益小（FY2025 −2.5），模型不另列")
    row("nm", "淨利率", "%", [f"={c}{g['ni']}/{c}{g['rev']}" for c in COLS], PCT)
    row("sh", "股數（十億股；歷史為稀釋加權平均）", "十億股", [f"=Hist!{HC[c]}{HR['sh_d']}" if c != "C" else None for c in "CDE"] + ["=shares"] * 7, "0.000", "FY2023 年報未列股數，不填")
    row("eps", "EPS（新台幣；歷史為公司公告稀釋 EPS）", "NT$", [f"=Hist!{HC[c]}{HR['eps_d']}" for c in "CDE"] + [f"=PL!{PCOL[c]}{PR['eps']}" for c in "FGHIJKL"], "0.00", "預測期＝PL 頁 EPS（＝淨利 × 匯率 ÷ 股數）")
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
    row2("r_niy", "淨利年增率（YoY）", "%", [f"=C{g['ni']}/Hist!C{HR['ni']}-1"] + [f"={c}{g['ni']}/{p}{g['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    P.cell(r, 1, "核對（應為 0；實績欄可有 ±0.002 的捨入差）").font = BOLD; r_chk = r; r += 1
    row2("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{g['rev']}-{c}{g['cogs']}-{c}{g['gp']}" for c in COLS], "0.000")
    row2("c2", "核對：毛利 − 營業費用 − 營業利益（FY2030E–FY2032E 為差額，見下）", "NT$B", [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "CDEFGHI"] + [None] * 3, "0.000")
    row2("c2b", "差額（FY2030E–FY2032E）：毛利 − 營業費用（沿用前一年比率推算）− 延伸期營業利益（由淨利還原）；目標價用延伸期淨利，不受影響", "NT$B", [None] * 7 + [f"={c}{g['gp']}-{c}{g['opx']}-{c}{g['oi']}" for c in "JKL"], "0.000")
    row2("c2c", "差額（FY2030E–FY2032E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{g['rev']}" for c in "JKL"], PCT)
    row2("c3", "核對：稅前利益 − 所得稅 − 稅後淨利", "NT$B", [f"={c}{g['pti']}-{c}{g['tax']}-{c}{g['ni']}" for c in COLS], "0.000")
    row2("c4", "核對：稅後淨利 ÷ 股數 − EPS（預測期；歷史 EPS 為歸屬母公司，另有非控制權益差）", "NT$", [None] * 3 + [f"={c}{g['ni']}/{c}{g['sh']}-{c}{g['eps']}" for c in "FGHIJKL"], "0.0000")
    row2("c5", "核對：營收 − DCF 頁情境 B 營收（FY2027E–FY2032E）／ − In 頁 ntd26（FY2026E）", "NT$B", [None] * 3 + [f"=F{g['rev']}-ntd26"] + [f"={c}{g['rev']}-DCF!{d}{d_rev}" for c, d in zip("GHIJKL", "CDEFGH")], "0.000")
    r += 1
    P.cell(r, 1, "年度與對帳中：本模型會計年度＝日曆年，FY(t) 直接對應 CY(t)；標題「（對帳中）」依 Coord 頁各日曆年旗標，FY2030E–FY2032E 由 FY2029 之後的收斂路徑推得，沿用 CY2029 旗標。FY2026 不是對帳中年份。")
    L["ycaveat"] = r; r += 2
    # ---- 分部營收推導（情境 B；十億美元；HTML 驅動因子章用） ----
    P.cell(r, 1, "分部營收推導（情境 B，十億美元；v0.5.5）：加速器顆數（Siliconomics／markets/memory）× 每顆台積電內容 × 價格調升 ×（1＋其他 AI）＝AI 相關營收（HPC 子列）；非 AI 平台＝前一年 ×（1＋非 AI 基準成長＋平台差距）；各平台相加＝營收。只引用本頁上方各列").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "", "", ""] + ["FY2026E", "FY2027E", "FY2028E", "FY2029E"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    L["seg_hdr"] = r; r += 1
    SC4 = {"F": 26, "G": 27, "H": 28, "I": 29}
    um = {26: "1", 27: "1", 28: "um28_B", 29: "um29_B"}

    def ucell(y):
        if y < 29: return "+".join(f"u_{k}_{y}" for k in ukeys) if y < 28 else "(" + "+".join(f"u_{k}_28" for k in ukeys) + ")*um28_B"
        return "(" + "+".join(f"u_{k}_28*(1+g29_{k})" for k in ukeys) + ")*um29_B"

    def seg(key, label, unit, fn, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for c in "FGHI":
            P[f"{c}{r}"] = fn(c); P[f"{c}{r}"].number_format = fmt
        if note: P.cell(r, 13, note)
        L[key] = r; r += 1
    seg("s_u", "AI 加速器顆數（情境 B，含情境係數）", "百萬顆", lambda c: "=" + ucell(SC4[c]), "0.00", "各平台顆數加總（In 頁 u_*；2028、2029 乘情境 B 係數）")
    seg("s_acc", "AI 加速器內容（Σ 顆數 × 每顆內容 × 價格）", "US$B", lambda c: f"=PL!{PCOL[c]}{PR['ai_acc']}", M)
    seg("s_oth", "其他 AI 相關加成（網通 ASIC、HBM 基底晶粒、CPU）", "US$B", lambda c: f"=PL!{PCOL[c]}{PR['ai']}-PL!{PCOL[c]}{PR['ai_acc']}", M)
    seg("s_ai", "AI 相關營收（HPC 子列）", "US$B", lambda c: f"=PL!{PCOL[c]}{PR['ai']}", M)
    for key, lbl in (("hpc_n", "HPC：非 AI"), ("sp", "智慧型手機"), ("iot", "物聯網"), ("au", "車用"), ("dce", "消費性電子（DCE）"), ("oth", "其他")):
        pr_ = next(rr for rr in range(26, 36) if str(P.cell(rr, 1).value or "").strip() == lbl)
        seg("s_" + key, lbl, "US$B", lambda c, pr_=pr_: f"=PL!{PCOL[c]}{pr_}", M)
    seg("s_rev", "營收（美元）", "US$B", lambda c: f"=PL!{PCOL[c]}{PR['rev']}", M)
    seg("s_chk", "核對：AI 相關營收 ＋ 五個非 AI 平台 − 營收（應為 0）", "US$B", lambda c: f"={c}{L['s_ai']}+SUM({c}{L['s_hpc_n']}:{c}{L['s_oth']})-{c}{L['s_rev']}", "0.000")
    seg("s_sh", "AI 相關營收 ÷ 營收", "%", lambda c: f"={c}{L['s_ai']}/{c}{L['s_rev']}", PCT)
    for col in COLS: P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0)
    return L



# ---------- Coord：讀取層級、C0 ----------
def coord_c0(wb, SIL, CR, HR, PR):
    CW = wb["Coord"]; N = SIL.NAMES
    wat_ver = re.search(r"v\d+(?:\.\d+)*", open(os.path.join(ROOT, "wattonomics", "model", "CURRENT"), encoding="utf-8").read()).group(0)
    P = wb["PL"]; CWS = wb["CoWoS"]
    r = CW.max_row + 3
    CW.cell(r, 1, "三層協調：讀取層級與 C0 對帳（v0.5.5；三層基準協調規則 v1.4.2 C0、C4、C5；本節只新增，不改上游值，也不改本模型既有數值）").font = BOLD; r += 1
    ROW = {}
    CW.cell(r, 1, "1. 讀哪一層、哪一版（下游只註明版本、不自行重估；Instructions 第 3 條）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "層級與讀取名稱", "版本", "模型使用處／備註"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rd = [("硬體股基準：Siliconomics 現行基準", f"{SIL.VERSION}", f"siliconomics/data/{SIL.FILE}（加速器顆數 SIF 單位列、每片 CoWoS 封裝數 cw_pp_*、其他 CoWoS 用途 cw_oth*、CoWoS 月產能 cw_capend*、委外 cw_osat、HBM 缺口、2029 顆數成長 g29_*、匯率 fx_twd、SIF_WTAnchor_Base）", SIL.VERSION, "台積電屬硬體供應鏈個股，讀 Siliconomics；不讀 Wattonomics 覆蓋 Siliconomics（C4）"),
          ("Wattonomics 現行版本（另一側讀數來源）", wat_ver, "wattonomics/model/CURRENT；經 Siliconomics SIF_DeployGW_AltRatio_<年> 換算成另一側讀數", wat_ver, "只用於兩讀數（本頁上方）與 Sens 頁 Wattonomics 錨定案例；不改基準"),
          ("顆數情境係數 情境 B（CY2028、CY2029；說明用）", "=um28_B", "In 頁 um28_B／um29_B（Siliconomics SIF_UnitsCoef_Base 為基準，siliconomics/coef_rule.py）；本版不改這些名稱與值（下游經本模型讀入）", SIL.VERSION, "PL 情境 B 顆數＝原始顆數 × 本係數；情境 A、C 為個股自訂"),
          ("公司自有變數（不由層級提供）", "—", "每顆台積電內容 ct_*、其他 AI 加成 ai_oth、價格調升 px27–29、非 AI 成長、毛利率、營業費用率、淨利轉換、資本支出", "本模型", "Siliconomics 沒有台積電營收口徑的量；層級只提供顆數與 CoWoS 片數／產能")]
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
    CW.cell(r, 1, "本模型會計年度＝日曆年，FY(t) 對 CY(t) 直接對應。下列各列都是模型從上游讀入的值對上游原值（鏈路完整性核對），不是對層級的獨立檢驗。CoWoS 使用率只比 2026、2027（情境 A 與層級基準係數相同）；2028 起兩者情境係數不同，不比。")
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
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, SAME); CW.cell(r, 8, note)
        r += 1
    nb = f"Siliconomics {SIL.VERSION}"
    ukeys = ["bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"]
    for i, y in enumerate((2026, 2027)):
        line("AI 加速器顆數合計（各平台加總；情境 A）", "百萬顆", f"CY{y}", "=" + "+".join(f"u_{k}_{y % 100}" for k in ukeys), N["SIF_UnitsPlat_Total"][i + 1], "0.000", f"{nb} SIF_UnitsPlat_Total（markets/memory 共用市場層；Interested-party／Assumed）")
    line("AI 加速器顆數合計（情境 A，＝原始顆數 × um28_A）", "百萬顆", "CY2028", "=(" + "+".join(f"u_{k}_28" for k in ukeys) + ")*um28_A", N["SIF_UnitsPlat_Total"][3], "0.000", f"{nb} SIF_UnitsPlat_Total")
    line("AI 加速器顆數合計（情境 A）", "百萬顆", "CY2029", "=(" + "+".join(f"u_{k}_28*(1+g29_{k})" for k in ukeys) + ")*um29_A", N["SIF_UnitsPlat_Total"][4], "0.000", f"{nb} SIF_UnitsPlat_Total；2029 成長取 Siliconomics g29_*")
    line("顆數情境係數 情境 B", "x", "CY2028", "=um28_B", N["SIF_UnitsCoef_Base"][3], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um28_B（本版不改）")
    line("顆數情境係數 情境 B", "x", "CY2029", "=um29_B", N["SIF_UnitsCoef_Base"][4], "0.0000", f"{nb} SIF_UnitsCoef_Base；In 頁 um29_B（本版不改）")
    for i, y in enumerate((2026, 2027)):
        line("CoWoS 需求 ÷ 可用產能（CoWoS 頁，情境 A）vs Siliconomics 基準使用率", "%", f"CY{y}", f"=CoWoS!{'BC'[i]}{lab(CWS, '需求 ÷ 可用產能')}", N["SIF_CoWoSUtil_Base"][i], "0.0%", f"{nb} SIF_CoWoSUtil_Base；同一組每片封裝數、其他用途、月產能與委外參數（原在本模型，v0.5 前後搬入 Siliconomics M14）")
    for y in (2026, 2027):
        line("CoWoS 月產能年底值", "千片／月", f"CY{y}", f"=cap_end{y % 100}", SIL.base(f"cw_capend{y % 100}"), "#,##0", f"{nb} cw_capend{y % 100}（TrendForce／Morgan Stanley；Interested-party）")
    line("新台幣兌美元（2027 起）", "NT$/US$", "CY2027", "=fx", SIL.base("fx_twd"), "0.00", f"{nb} fx_twd")
    for y in (2027, 2028, 2029):
        line("Wattonomics 錨定係數", "x", f"CY{y}", f"=wta_{y % 100}", SIL.WTA[y], "0.000", f"{nb} SIF_WTAnchor_Base（Derived）；SIL.wta_rows 直接讀入 In 頁（基準不乘）")
    ROW["first"], ROW["last"] = first, r - 1
    CW.cell(r, 1, "對帳中項目數（讀取層級核對：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"); CW.cell(r, 7, 0)
    ROW["cnt"] = r; r += 2
    CW.cell(r, 1, "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目數）").font = BOLD; ROW["ref"] = r; r += 1
    CW.cell(r, 1, "以下是 Siliconomics 現行基準與 Wattonomics 換算值兩層之間的差距，不是台積電自己的量對層級的比較；本公司不裁決（C4），只在兩讀數目標價並列。"); r += 1
    for j_, h in enumerate(["項目", "單位", "年度", "Siliconomics 基準", "Wattonomics 換算值", "差距（基準 ÷ 換算值 − 1）", "判定", "層級名稱與版本"], 1):
        c = CW.cell(r, j_, h); c.font = BOLD; c.fill = HDR
    ROW["refh"] = r; r += 1
    ROW["ref1"] = r
    for i, y in enumerate((2027, 2028, 2029)):
        CW.cell(r, 1, "AI 加速器顆數合計（Siliconomics 基準係數）：基準 vs Wattonomics 換算值"); CW.cell(r, 2, "百萬顆"); CW.cell(r, 3, f"CY{y}")
        CW.cell(r, 4, N["SIF_UnitsPlat_Total_Base"][i + 2]).number_format = "0.000"; CW.cell(r, 4).font = BLUE
        CW.cell(r, 5, f"=D{r}*{'CDE'[i]}${CR['ratio']}").number_format = "0.000"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; CW.cell(r, 7, f'=IF(ABS(F{r})>0.15,"層間差距（參考）","一致")')
        CW.cell(r, 8, f"Siliconomics SIF_UnitsPlat_Total_Base；Wattonomics {wat_ver} 經 SIF_DeployGW_AltRatio_{y} 換算；Derived；對帳中年份依 C4 並列兩讀數目標價，個股不裁決")
        r += 1
    ROW["ref2"] = r - 1; r += 2
    CW.cell(r, 1, "3. 對帳中年份與兩讀數（C4）").font = BOLD; r += 1
    CW.cell(r, 1, F(T("硬體股：Siliconomics 標對帳中的年份為 CY2027–CY2029，本模型估值依賴（FY(t)＝CY(t)；FY2030–FY2032 延伸期由 FY2029 收斂推得；FY2026 為前三季實績加 4Q26 指引的錨點、CY2026 非對帳中）。"
                    "依 C4 並列兩個讀數的目標價：現行基準讀數 "), TX(f"B{CR['tp']}", "#,##0.00"), T(" 元（"), f"B{CR['rec']}", T("）、另一側讀數 "), TX(f"C{CR['tp']}", "#,##0.00"),
                    T(" 元（"), f"C{CR['rec']}", T("）；兩讀數只差產業部署 GW 一個輸入，個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    ROW["s3"] = r; r += 1
    CW.cell(r, 1, F(T("兩讀數的縮放比值：CY2027 "), TX(f"C{CR['ratio']}", "0.000"), T("、CY2028 "), TX(f"D{CR['ratio']}", "0.000"), T("、CY2029 "), TX(f"E{CR['ratio']}", "0.000"),
                    T("（另一側 ÷ 現行基準）。另一側讀數建置：情境 B、C 於 CY2027 起 AI 加速器內容（含其他 AI 加成）同乘比值，情境 A 不縮放；非 AI 營收、毛利率、資本支出與其他假設不變，整本重建（tools/coord/alt_run.py）。"
                      "本模型會計年度與日曆年相同，無月數近似；敏感度頁另有 Wattonomics 錨定案例（用 SIF_WTAnchor_Base，與本比值不同）。")))
    ROW["s3b"] = r; r += 1
    CW.column_dimensions["F"].width = 18; CW.column_dimensions["G"].width = 26; CW.column_dimensions["H"].width = 110; CW.column_dimensions["I"].width = 40
    r += 1
    # ---- 4. C0（實績）----
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.5.5；規則 v1.4.2 C0 第 2 點、工作單 20261114_第二批13家；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW["title"] = r; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（2026 年 1–10 月月營收、1Q26–3Q26 季營收與 HPC 平台占比；不用指引或預測）。層級側＝Siliconomics 現行基準的同口徑隱含值。台積電的層級值只有「加速器顆數」與「CoWoS 片數／產能」；換成營收須乘每顆台積電內容、其他 AI 加成等本公司變數，"
                  "這些變數以公司 2025 年揭露占比（high-teens %）與 2024–2029 年 AI 加速器營收 CAGR 指引校準，屬個股預測而非層級值。時差：主讀數＝已公布期間累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12；第二讀數＝最近一季 × 4 vs 層級 CY2026；兩者差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "同源檢查（結論）：本表沒有任何一列同時具備「公司已公布實績」與「獨立於本公司模型的層級隱含值」。層級值消去哪些量：AI 營收隱含值＝顆數（層級）× 每顆內容（本公司，由公司 CAGR／占比校準）× 其他 AI 加成（本公司假設，使 2026 AI 營收貼近公司口徑推算）；"
                  "校準目標與可得實績同為公司說法，故 2026 AI 營收的比較只檢驗本公司的校準，不是層級。")
    ROW["same"] = r; r += 1
    CW.cell(r, 1, "公司揭露缺口：公司 2026 年 1–3 季（含 3Q26 法說與 Management Report）未揭露 AI 加速器營收金額或占比（只重申 2024–2029 年 AI 加速器營收 CAGR 為 50% 中後段至高 50%，3Q26 法說稱可能更高）、未揭露 CoWoS 片數、月產能或先進封裝營收金額（僅敘述性說明）。"
                  "因此 AI 與 CoWoS 兩項無可比實績；實績可比的只有合併營收與 HPC 平台，而層級沒有對應量。以上依公司新聞稿、法說逐字稿的二手摘錄與本模型 Sources 頁；公司原文 PDF（3Q26 逐字稿、Management Report）本輪未能開啟，待人工複核。")
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
    put("q1", "合併營收：1Q26", "NT$B", "2026-01～03", "=ntd26_1q", "Verified", "In 頁 ntd26_1q（公司 1Q26 新聞稿 pr.tsmc.com/english/news/3297；連格）", "#,##0.00")
    put("q2", "合併營收：2Q26", "NT$B", "2026-04～06", "=ntd26_2q", "Verified", "In 頁 ntd26_2q（公司 2Q26 新聞稿；連格）", "#,##0.00")
    put("q3", "合併營收：3Q26", "NT$B", "2026-07～09", "=ntd26_3q", "Verified", "In 頁 ntd26_3q（公司 3Q26 新聞稿 pr.tsmc.com/english/news/3346；連格）", "#,##0.00")
    put("m10", "月營收：2026 年 10 月", "NT$B", "2026-10", 586.890319, "Verified", "公司 2026-11-10 月營收 6-K 與新聞稿（網址見 Sources 頁）：NT$586,890,319 千元，月增 14.7%、年增 59.7%", "#,##0.000")
    put("ytd", "月營收：2026 年 1–10 月累計", "NT$B", "2026-01～10", 4485.617293, "Verified", "同上 6-K：NT$4,485,617,293 千元，年增 43.3%（累計 2025 同期 3,130,436,902 千元）", "#,##0.000")
    put("xchk", "核對：1–10 月累計 − 1Q26 − 2Q26 − 3Q26 − 10 月（應約 0；月營收為未經查核自結數，季營收為財報）", "NT$B", "—", f"=D{r - 1}-D{r - 5}-D{r - 4}-D{r - 3}-D{r - 2}", "Derived", "差額來自月營收自結與季財報的四捨五入", "0.000")
    for q, k, nm in ((1, "q1", "1Q26"), (2, "q2", "2Q26"), (3, "q3", "3Q26")):
        put(f"hs{q}", f"HPC 平台占營收：{nm}", "%", f"2026 Q{q}", f"=pf26_{q}q_hpc", "Verified（占比，整數四捨五入）", "In 頁 pf26_*_hpc（公司 Management Report「Net Revenue by Platform」；連格）。占比為公司揭露的整數百分比，換算金額有 ±0.5 個百分點（約 ±1%）的四捨五入誤差", "0%")
    for q, k, nm in ((1, "q1", "1Q26"), (2, "q2", "2Q26"), (3, "q3", "3Q26")):
        put(f"h{q}", f"HPC 平台營收：{nm}（占比 × 合併營收）", "NT$B", f"2026 Q{q}", f"=D{inp[k]}*D{inp['hs' + str(q)]}", "Derived（占比 Verified；金額為換算）", "公司只揭露占比，未揭露平台金額；口述占比 × 營收屬換算值，不當作公司揭露金額", "#,##0.0")
    put("h9", "HPC 平台營收：2026 年 Q1–Q3 累計", "NT$B", "2026-01～09", f"=D{inp['h1']}+D{inp['h2']}+D{inp['h3']}", "Derived", "三季換算值加總", "#,##0.0")
    put("hqoq", "參考：3Q26 HPC 營收季增（換算值）", "%", "3Q26 vs 2Q26", f"=D{inp['h3']}/D{inp['h2']}-1", "Derived", "公司法說稱 HPC 季增 22%；換算值與之差距反映占比四捨五入的誤差量級", "0.0%")
    put("h9u", "HPC 平台營收：2026 年 Q1–Q3 累計（美元）", "US$B", "2026-01～09", "=rev26_1q*pf26_1q_hpc+rev26_2q*pf26_2q_hpc+rev26_3q*pf26_3q_hpc", "Derived", "In 頁美元季營收 × 占比（連格）", "#,##0.00")
    put("h3u", "HPC 平台營收：3Q26（美元）", "US$B", "2026-07～09", "=rev26_3q*pf26_3q_hpc", "Derived", "同上", "#,##0.00")
    put("ai", "層級隱含：AI 相關營收 CY2026（顆數 × 每顆內容 ×（1＋其他 AI））", "US$B", "CY2026", f"=PL!B{PR['ai']}", "Derived（顆數＝層級；內容與加成＝本公司）", "PL 頁 2026 欄 AI 相關營收（連格）；顆數取 Siliconomics／markets/memory，每顆內容與其他 AI 加成為本公司變數（同源校準，見上）", "#,##0.00")
    put("cw", "層級隱含：CoWoS 需求 CY2026（千片；顆數 ÷ 每片封裝數＋其他用途）", "千片", "CY2026", f"=CoWoS!B{lab(CWS, '需求合計')}", "Derived（層級參數）", "CoWoS 頁需求合計（連格）；每片封裝數 cw_pp_* 由外部券商與媒體的片數估計校準（Derived／Assumed）", "#,##0")
    put("cu", "層級隱含：CoWoS 需求 ÷ 可用產能 CY2026", "%", "CY2026", f"=CoWoS!B{lab(CWS, '需求 ÷ 可用產能')}", "Derived", "Siliconomics SIF_CoWoSUtil_Base 同值（讀取層級核對）；產能為 TrendForce／Morgan Stanley 估計（Interested-party），不是公司揭露", "0.0%")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "AI 隱含 ÷ HPC 實績", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1
    CW.cell(r, 1, "參考：主讀數：AI 隱含 CY2026 × 9/12 占 HPC 平台 Q1–Q3 累計實績"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, "2026-01～09 vs CY2026 × 9/12")
    CW.cell(r, 4, f"=D{inp['h9u']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['ai']}*9/12").number_format = "#,##0.0"
    CW.cell(r, 6, f"=E{r}/D{r}").number_format = "0.0%"; CW.cell(r, 7, "參考（同源校準）")
    CW.cell(r, 8, "AI 只是 HPC 的一部分（HPC 另含 CPU、PC、一般伺服器），公司沒有揭露 AI 占 HPC 的實際比例，故這是占比而不是差距，無檢驗力；只顯示層級隱含 AI 營收未超過 HPC 實績（必要條件）")
    ROW["ra"] = r; r += 1
    CW.cell(r, 1, "參考：第二讀數：AI 隱含 CY2026 占 HPC 平台 3Q26 × 4"); CW.cell(r, 2, "US$B"); CW.cell(r, 3, "2026-07～09 × 4 vs CY2026")
    CW.cell(r, 4, f"=D{inp['h3u']}*4").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['ai']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=E{r}/D{r}").number_format = "0.0%"; CW.cell(r, 7, "參考（同源校準）")
    CW.cell(r, 8, "最近一季年化（約等於 3Q26 中點的年化速率；HPC 季成長快，年化值高於全年）；同上不判定")
    ROW["rb"] = r; r += 1
    CW.cell(r, 1, "C0（實績）判定：合併營收與月營收（NT$）")
    CW.cell(r, 7, "無可比實績"); CW.cell(r, 8, "結論：無可比實績。公司營收實績完整（見上），但層級（Siliconomics、Wattonomics）沒有台積電營收口徑的量；缺：層級側台積電營收或晶圓／先進封裝營收的獨立估計。不用本模型預測補位")
    ROW["j1"] = r; r += 1
    CW.cell(r, 1, "C0（實績）判定：HPC 平台營收")
    CW.cell(r, 7, "無可比實績"); CW.cell(r, 8, "結論：無可比實績。HPC 含 AI 加速器與非 AI（CPU、PC 等），層級只有 AI 加速器顆數；非 AI 部分為本公司假設。缺：層級側 HPC 或 AI 加速器營收的獨立估計；公司側缺 HPC 金額（只有整數占比）")
    ROW["j2"] = r; r += 1
    CW.cell(r, 1, "C0（實績）判定：AI 加速器營收（CY2026）")
    CW.cell(r, 7, "參考（同源校準）"); CW.cell(r, 8, "結論：無可比實績（層級值同源）。公司 2026 年尚無 AI 加速器營收實績（只有 CAGR 指引與 2025 年占比）；層級隱含值的每顆內容與其他 AI 加成由公司 CAGR／占比校準，與可得公司說法同源。待公司揭露 FY2026 AI 加速器營收（約 2027 年初法說）後重做")
    ROW["j3"] = r; r += 1
    CW.cell(r, 1, "C0（實績）判定：CoWoS 片數／產能／先進封裝營收")
    CW.cell(r, 7, "無可比實績"); CW.cell(r, 8, "結論：無可比實績。缺：公司對 CoWoS 片數、月產能或先進封裝營收金額的揭露（2026 年 1–3 季僅有敘述性說明）。層級的月產能與使用率為 TrendForce、Morgan Stanley 等外部估計，只列參考，不判定")
    ROW["j4"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ROW["j1"]}:G{ROW["j4"]},"對帳中")')
    ROW["cnt4"] = r; ROW["last"] = r; r += 1
    CW.cell(r, 1, "讀法：本模型 2026 年用前三季實績加 4Q 指引，營收與 HPC 占比都是公司已公布實績，不構成對層級的檢驗；層級對台積電只約束顆數與 CoWoS 供需。因此本輪 C0（實績）沒有可判定項目，對帳中 0；外部估計（賣方共識、券商 CoWoS 片數）只在 Recon 頁列參考。")
    ROW["read"] = r; r += 1
    return ROW


def readme13(wb, HR, PLA, RVA, SNS, CR, CO):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]
    vb = lambda p: f"Val!B{lab(V, p, exact=True)}"
    tp, up, rec, dcf, cu, cadj = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議"), vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（市場主流同業法）"), vb("Comparable 目標價（相對倍數調整後）")
    coef, pmed, feps, cons = vb("相對倍數調整係數"), vb("同業前瞻 P/E 中位數"), vb("12 個月後的前瞻 EPS（情境加權）"), vb("市場共識目標價")
    sens_r = lambda key: next(r for r in range(SNS["first"], SNS["last"] + 1) if str(SE.cell(r, 1).value or "").startswith(key))
    s_u, s_c, s_p, s_g, s_w = sens_r("AI 加速器顆數"), sens_r("每顆台積電內容"), sens_r("AI 內容價格調升"), sens_r("毛利率"), sens_r("情境權重")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.5 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1
    out = {}

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), TX(tp, "#,##0"), T(" 元，相對現價 "), TX(up, "+0.0%;-0.0%"), T("，評等"), rec, T("：DCF "), TX(dcf, "#,##0"), T(" 元，同業比較法（相對倍數調整後）"), TX(cadj, "#,##0"), T(" 元。"))
    put("rec", T("評等「"), rec, T("」：現價已接近模型價值。同業比較法採台積電過去十年相對同業的本益比倍數（"), TX(coef, "0.00"), T(" 倍）；未調整的市場主流同業法 "), TX(cu, "#,##0"), T(" 元。")) 
    put("o1", T("目標價：加權目標價 "), TX(tp, "#,##0"), T(" 元，現價 "), TX("px", "#,##0"), T(" 元，差距 "), TX(up, "+0.0%;-0.0%"), T("。同業比較法（中位數 "), TX(pmed, "0.0"), T(" 倍 × 係數 "), TX(coef, "0.00"),
        T(" × 12 個月後 EPS "), TX(feps, "0.0"), T(" 元）得 "), TX(cadj, "#,##0"), T(" 元，DCF 得 "), TX(dcf, "#,##0"), T(" 元，分析師共識 "), TX(cons, "#,##0"), T(" 元。"))
    put("o2", T("模型的核心是 AI 加速器：2026 年 AI 相關營收約 "), TX(f"PL!B{CO['_pr']['ai']}", "0.0"), T(" 十億美元，占營收 "), TX(f"PL!B{CO['_pr']['ai']}/PL!B{CO['_pr']['rev']}", "0%"),
        T("；顆數取 Siliconomics 現行基準，每顆台積電內容與其他 AI 加成是本公司變數，由公司 2024–2029 年 AI 加速器營收年複合成長率指引校準。"))
    put("o3", T("最大變數（加權目標價相對基準）：顆數 ×0.8／×1.2 為 "), TX(f"Sens!B{s_u}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_u}", "+0.0%;-0.0%"),
        T("；毛利率 ±3 個百分點 "), TX(f"Sens!B{s_g}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_g}", "+0.0%;-0.0%"),
        T("；價格調升不漲／加倍半 "), TX(f"Sens!B{s_p}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_p}", "+0.0%;-0.0%"),
        T("；情境權重全給保守／樂觀 "), TX(f"Sens!B{s_w}", "+0.0%;-0.0%"), T("／"), TX(f"Sens!C{s_w}", "+0.0%;-0.0%"), T("。這些是測試案例，不是門檻。"))
    put("o4", T("反向 DCF：WACC "), TX("wacc", "0.0%"), T("，要使 DCF 等於現價，折現率須降到 "), TX(rv("w1"), "0.0%"), T("，或 FY2028 起現金流放大 "), TX(rv("s1"), "0.00"),
        T(" 倍；現價隱含 FY2028 營收約 "), TX(rv("rv1"), "#,##0"), T(" 十億新台幣，模型 "), TX(rv("rv0"), "#,##0"), T("。"))
    put("o5", T("對帳中：CY2027–CY2029 的 AI 加速器顆數，Wattonomics 換算值只有 Siliconomics 基準的 "), TX(f"Coord!C{CR['ratio']}", "0.00"), T("～"), TX(f"Coord!D{CR['ratio']}", "0.00"),
        T(" 倍；另一側讀數加權目標價 "), TX(f"Coord!C{CR['tp']}", "#,##0"), T(" 元（"), f"Coord!C{CR['rec']}", T("），現行基準讀數（"), f"Coord!B{CR['rec']}", T("），個股不裁決。公司實績（月營收、HPC 平台）對層級隱含值：無可比實績（層級值同源；標參考）。"))
    P = CO["_pr"]
    put("drv", T("分部營收推導（情境 B）：加速器顆數（層級）× 每顆台積電內容 × 價格調升 ×（1＋其他 AI）＝AI 相關營收；非 AI 平台逐年成長；各平台相加為營收。FY2028 顆數 "),
        TX(f"PL!H{PLA['s_u']}", "0.0"), T(" 百萬顆、AI 相關營收 "), TX(f"PL!H{PLA['s_ai']}", "0.0"), T("、營收 "), TX(f"PL!H{PLA['s_rev']}", "0.0"), T(" 十億美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達）、三層協調讀取層級與 C0（含公司實績，規則 v1.4.2）、硬體股兩讀數。章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 Hist 頁（損益表歷史 FY2022–FY2025，公司年報）、RevDCF 頁（反向 DCF）、"
                  "PL 頁末標準損益表與分部營收推導、Val 頁末 0.7 四個目標價、Sens 頁末加權目標價相對基準（整本重算快照）、Coord 頁（兩讀數目標價；讀取層級；C0：月營收與 HPC 平台實績對層級隱含值，AI 營收標參考（同源校準），其餘無可比實績，對帳中 0）、README 頁末 v1.3 白話列。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異；um28_B／um29_B 等名稱與值不動），加權目標價與評等不變。前版 README 數字照舊，本節為補充說明")
    out["ver"] = r; r += 1
    return out


def apply(wb, here, SIL, CR, ukeys, PR, COLOF=None, SC=None):
    wb["Coord"].cell(CR["tp"], 1).value = "加權目標價"
    HR = hist_is(wb)
    PLA = pl_standard(wb, HR, CR, PR, ukeys)
    RVA = revdcf(wb)
    VG = val_g(wb)
    SNS = sens_tor(wb, here)
    CO = coord_c0(wb, SIL, CR, HR, PR)
    CO["_pr"] = PR
    RMA = readme13(wb, HR, PLA, RVA, SNS, CR, CO)
    SO = wb["Sources"]
    rr = SO.max_row + 1
    for i, (t, u) in enumerate((
        ("公司：2026 年 10 月月營收 6-K（2026-11-10；NT$586,890,319 千元；1–10 月累計 NT$4,485,617,293 千元；C0 實績用）", REV10),
        ("公司：2026 年 10 月月營收新聞稿", PR10),
        ("公司：2025 年報第六章 合併損益表（FY2024–FY2025；Hist 頁用）", AR25),
        ("公司：2025 財報 FS（FY2025、FY2024）", FS25),
        ("公司：2023 年報第六章 合併損益表（FY2022–FY2023；Hist 頁用）", AR23))):
        SO.cell(rr + i, 1, t); SO.cell(rr + i, 2, u)
    wb._v055 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, SNS=SNS, RM=RMA, CO=CO)
    return f"v055_post：Hist 頁；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 {VG} 起；Sens 相對基準 {SNS['t0']} 起；Coord C0 {CO['rd0']}–{CO['last']}；README v1.3 {len(RMA)} 列"
