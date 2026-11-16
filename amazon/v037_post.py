# Amazon v0.3.7（2026-11-15）後處理：個股 HTML 版面規格 v1.3 全套＋Coord 頁「三層協調」（MAG 類；樣板 Alphabet v0.3.9 alphabet/v039_post.py、Microsoft v0.3.8 microsoft/v038_post.py）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點；tools/coord/README.md（追公式、同源校準、實績期間）。
# 只新增，不改任何既有儲存格（既有數值、目標價、評等不變；逐格比對只允許版號文字與新增格）：
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E（9M 實績＋Q4 模型）–FY2032E 情境 B；GAAP；另列 9M26A 比較欄）、比率區、核對列、分部營收（圖表用）
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑，封閉解與二分法）與分部加總隱含倍數；只讀其他頁，無任何頁引用本頁
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用，白話標籤；引用上方各案例 E 欄）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v21.10 快照；Wattonomics v0.3.2 只作對帳列）、F1b 追公式（tools/coord/trace_wattonomics.py --player amzn）、F2 C0（實績）
#   README 頁末：v1.3 展示層白話列
# build_amazon.py 在 resolve_all() 之後、存檔前呼叫 apply(wb, ctx)；--set／--q3v03（敏感度、拆解重算）時不呼叫。
import json
import os
import re

from openpyxl.styles import Font, PatternFill

VER, BUILT = "0.3.8", "2026-11-16"
U_PR = "https://s2.q4cdn.com/299287126/files/doc_earnings/2026/q3/earnings-result/AMZN-Q3-2026-Earnings-Release.pdf"   # 出處網址只放 Coord 頁 L 欄（HTML 不顯示外部網址）
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
PARAM_RE = re.compile(r'(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+(?![A-Za-z0-9])')


def lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F_(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def T(s):
    return ("t", s)


def last_text_row(ws):
    r = ws.max_row
    while r > 0 and ws.cell(r, 1).value in (None, ""):
        r -= 1
    return r


def hdr(ws, r, labels):
    for j, h in enumerate(labels, 1):
        if h is None:
            continue
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR


# =====================================================================================================
# PL：標準損益表（版面規格 v1.3 D 節）
# 欄：C FY2023A（Q 頁）、D FY2024A、E FY2025A、F FY2026E（9M 實績＋Q4 模型）、G–L FY2027E–FY2032E（PL 情境 B 區塊 C、D、G、H–M 欄）、M 9M26A（比較欄）、N 說明
PCOL = dict(zip("DEFGHIJKL", "CDGHIJKLM"))   # 本表欄 → PL 情境 B 區塊欄
YC = "CDEFGHIJKL"


def pl_standard(wb, C):
    P = wb["PL"]
    B, QR = C["B"], C["QR"]
    r = last_text_row(P) + 3
    L = {"start": r}
    P.cell(r, 1, "標準損益表（v0.3.7 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023A–FY2025A）＋FY2026E（9M 實績＋Q4 模型）–FY2032E（情境 B 基準）；最右欄 9M26A 為今年已公布九個月實績（比較用）。US$B；EPS 美元。"
                 "只引用本頁上方與 Q 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（財報口徑）。Amazon 以 GAAP 揭露損益，市場共識 EPS 也以 GAAP 稀釋為主；本表營業利益至 EPS 取 PL 頁情境 B 財報口徑，不含自用 AI 影子效益（評價口徑列見上方情境區塊）。"
                 "其他收入（淨額）＝稅前淨利 − 營業利益（利息收支與業外其他；FY2026 含 Anthropic 等持股的評價利益，預測期不估），所以 FY2027E 淨利年增率為負。")
    L["note"] = r; r += 1
    P.cell(r, 1, "營業成本、毛利：Amazon 損益表把成本分列於銷貨成本、履約、技術與基礎設施、銷售與行銷、管理等多列，本模型以分部營業利益建模、沒有分列，事實總帳也未收銷貨成本，所以營業成本、毛利、毛利率三列記「—」（不推估，避免假精確）；"
                 "「營業費用」列為營收 − 營業利益（全部營業成本與費用合計）。")
    L["note2"] = r; r += 1
    hdr(P, r, ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "9M26A（比較）", "說明"])
    L["hdr"] = r; r += 1
    M = "#,##0.000"; PCT = "0.0%"
    idx = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in cells.items():
            if v is None:
                continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note:
            P.cell(r, 14, note)
        idx[key] = r; r += 1

    pl = lambda k: {c: f"=PL!{PCOL[c]}{B[k]}" for c in "DEFGHIJKL"}
    q = lambda k, col: f"=Q!{col}{QR[k]}"
    dash = {c: "—" for c in "CDEFGHIJKLM"}
    row("rev", "營收", "US$B", {"C": q("tot", "C"), **pl("tot"), "M": q("tot", "O")}, M, "合併營收（分部加總；FY2023A 取 Q 頁；9M26A 取 Q 頁 9M26 欄）")
    yo = {c: f"={c}{idx['rev']}/{p}{idx['rev']}-1" for p, c in zip(YC, YC[1:])}
    yo["C"] = "—"; yo["M"] = f"=M{idx['rev']}/Q!N{QR['tot']}-1"
    row("yoy", "營收年增率（YoY）", "%", yo, PCT, "FY2023A 的基期（FY2022）不在 Q 頁，記「—」；9M26A 對 9M25")
    row("cogs", "營業成本（COGS）", "US$B", dash, "@", "模型未分列（見上方說明）")
    row("gp", "毛利", "US$B", dash, "@", "模型未分列")
    row("gm", "毛利率", "%", dash, "@", "模型未分列")
    row("opx", "營業費用（全部營業成本與費用合計＝營收 − 營業利益）", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['rev'] + 6}" for c in "CDEFGHIJKLM"}, M, "含銷貨成本、履約、技術與基礎設施、銷售與行銷、管理與其他營業費用")
    row("oi", "營業利益", "US$B", {"C": q("oi", "C"), **pl("oi"), "M": q("oi", "O")}, M, "財報口徑合併營業利益（不含自用 AI 影子效益）")
    assert idx["oi"] == idx["rev"] + 6
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("oth", "其他收入（淨額）＝稅前淨利 − 營業利益", "US$B", {c: f"={c}{idx['oi'] + 3}-{c}{idx['oi']}" for c in "CDEFGHIJKLM"}, M, "利息收支、權益法投資與業外其他（含持股評價損益）")
    row("pti", "稅前利益", "US$B", {"C": q("pretax", "C"), **pl("pretax"), "M": q("pretax", "O")}, M)
    assert idx["pti"] == idx["oi"] + 3
    row("tax", "所得稅", "US$B", {"C": q("taxx", "C"), **pl("taxx"), "M": q("taxx", "O")}, M)
    row("ni", "稅後淨利", "US$B", {"C": q("ni", "C"), **pl("ni"), "M": q("ni", "O")}, M, "財報淨利（PL 頁、Q 頁）；與稅前 − 所得稅的差額為權益法投資損益（稅後，損益表另列；預測期設 0），見核對列")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("sh", "稀釋加權股數", "十億股", {"C": q("shd", "C"), **pl("sh"), "M": q("shd", "O")}, M)
    row("eps", "EPS（GAAP 完全稀釋）", "US$", {"C": q("eps", "C"), **pl("eps_fin"), "M": q("eps", "O")}, "0.00", "淨利 ÷ 稀釋加權股數（PL 頁財報口徑 EPS；FY2023A、9M26A 取 Q 頁）")
    P.cell(r, 1, "比率").font = BOLD; L["ratio"] = r; r += 1
    row("r_yoy", "營收年增率", "%", {c: f"={c}{idx['yoy']}" for c in "DEFGHIJKLM"}, PCT)
    row("r_opm", "營業利益率", "%", {c: f"={c}{idx['opm']}" for c in "CDEFGHIJKLM"}, PCT)
    row("r_nm", "淨利率", "%", {c: f"={c}{idx['nm']}" for c in "CDEFGHIJKLM"}, PCT)
    ny = {c: f"={c}{idx['ni']}/{p}{idx['ni']}-1" for p, c in zip(YC, YC[1:])}
    ny["C"] = "—"; ny["M"] = f"=M{idx['ni']}/Q!N{QR['ni']}-1"
    row("r_niy", "淨利年增率（YoY）", "%", ny, PCT, "FY2026 含持股評價利益（預測期不估），所以 FY2027E 年增率為負；9M26A 對 9M25")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; L["chk"] = r; r += 1
    row("c1", "差額：稅後淨利 −（稅前利益 − 所得稅）＝權益法投資損益（稅後；實績年非 0，預測期為 0）", "US$B", {c: f"={c}{idx['ni']}-({c}{idx['pti']}-{c}{idx['tax']})" for c in "CDEFGHIJKLM"}, "0.000")
    row("c2", "核對：營收 − 分部營收合計（下方）", "US$B", {c: f"={c}{idx['rev']}-{c}{r + 8}" for c in "DEFGHIJKL"}, "0.000")
    row("c3", "核對：營業利益＋其他收入 − 稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}-{c}{idx['pti']}" for c in "CDEFGHIJKLM"}, "0.000")
    r += 1
    P.cell(r, 1, "分部營收（US$B；圖表用；情境 B）").font = BOLD; L["seg"] = r; r += 1
    sidx = {}
    for k, lab in (("na", "North America"), ("intl", "International"), ("aws_o", "AWS 其他（非 AI 雲端）"), ("aws_ai", "AWS AI（對外 AI 算力；估計）")):
        P.cell(r, 1, lab); P.cell(r, 2, "US$B")
        for c in "DEFGHIJKL":
            P[f"{c}{r}"] = f"=PL!{PCOL[c]}{B[k]}"; P[f"{c}{r}"].number_format = M
        sidx[k] = r; r += 1
    P.cell(r, 1, "合計"); P.cell(r, 2, "US$B")
    for c in "DEFGHIJKL":
        P[f"{c}{r}"] = f"=SUM({c}{sidx['na']}:{c}{sidx['aws_ai']})"; P[f"{c}{r}"].number_format = M
    sidx["tot"] = r; r += 1
    assert sidx["tot"] == idx["c2"] + 8, (sidx, idx)
    for col in "CDEFGHIJKLM":
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(idx=idx, sidx=sidx)
    return L


# =====================================================================================================
# RevDCF：反向 DCF（情境 B 路徑；DCF 明確期 Val 頁 H–M 欄＝FY2027–FY2032）
def revdcf(wb, C):
    B, V, VR = C["B"], C["BV"], C["VR"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    vc = lambda k: f"Val!$C${V[k]}"
    cols = "HIJKLM"
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.7 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益為線性，用封閉解；WACC 用二分法）。本頁只讀 Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 每股價值推到目標時點＝現價」（與目標價同一時點口徑：2026 年底每股價值 ×（1＋股權成本）^年數 − 期間股利；營業利益為評價口徑，含自用 AI 影子效益）；"
                "分部加總部分反解「分部加總每股價值＝現價」。營業利益率與營收規模兩項只動 FY2028 起的營業利益（FY2027 不動；折舊、資本支出、融資租賃、其他營運項目不變）。反解值超出合理範圍時照實列出，不外推")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    hdr(RV, r, ["項目", "值", "說明"]); r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note:
            RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的 2026 年底每股價值（美元）", "=(px+dps*tp_yf)/(1+ke)^tp_yf", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（Val 頁 DCF 推到目標時點的反算）")
    kv("sh", "流通股數（十億股；Val 頁 DCF 用 sh_out）", "=sh_out", "0.000")
    kv("oth", "持股價值 − 2026 年底淨負債（US$B；含融資租賃與融資義務）", f"={vc('hold')}-{vc('nd26')}", "#,##0.0", "同 Val 頁 DCF 股東權益橋的非 EV 項")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{A['v0']}*B{A['sh']}-B{A['oth']}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B）", f"={vc('ev')}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 2026 年底每股價值（美元）", f"={vc('ps')}", "0.00")
    r += 1
    yh = r
    hdr(RV, r, ["項目（情境 B）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"]); r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, vcol in enumerate(cols):
            RV.cell(r, 2 + i, fn(vcol, "BCDEFG"[i], i)).number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$B）", lambda v, c, i: f"=PL!{v}{B['tot']}", "#,##0.0")
    yr("ebit", "營業利益（評價口徑，含自用 AI 影子效益；US$B）", lambda v, c, i: f"=PL!{v}{B['oi_v']}", "#,##0.0")
    yr("m", "營業利益率（評價口徑）", lambda v, c, i: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "折現年數（期中，距 2026 年底）", lambda v, c, i: f"={i}+0.5", "0.0")
    yr("df", "折現因子（Val 頁）", lambda v, c, i: f"=Val!{v}{V['df']}", "0.0000")
    yr("fcf", "FCFF（Val 頁；US$B）", lambda v, c, i: f"=Val!{v}{V['ufcf']}", "#,##0.0")
    yr("tvf", "終值係數（只在 FY2032：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) ÷（1＋WACC）^6）",
       lambda v, c, i: ("=(1+tg)*(1-tg/ronic)/(wacc-tg)/(1+wacc)^6" if v == "M" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（FY2027 不動，記 0；之後＝(1 − 稅率) × (折現因子＋終值係數)）", lambda v, c, i: ("=0" if v == "H" else f"=(1-tax)*({c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與 FY2028 起營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032，以同一係數加權）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "終值年（FY2032）權重大")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收與營業利益等比例放大；營業利益率不變，折舊、資本支出不變）").font = BOLD; r += 1
    kv("EB", "FY2028–FY2032 營業利益 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=1+(B{A['evr']}-B{A['ev']})/B{A['EB']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年 FCFF 不變；永續期 RONIC＝WACC＋原差；股權成本不變；二分法 40 次）").font = BOLD; r += 1
    hdr(RV, r, ["次", "下界", "上界", "中點", "中點的 DCF 每股價值推到目標時點（美元）"]); r += 1

    def px_at(w):
        ev = (f"SUMPRODUCT($B${Y['fcf']}:$G${Y['fcf']},1/(1+{w})^$B${Y['t']}:$G${Y['t']})"
              f"+Val!$M${V['nopat']}*(1+tg)*(1-tg/({w}+ronic-wacc))/({w}-tg)/(1+{w})^6")
        return f"(({ev})+$B${A['oth']})/$B${A['sh']}*(1+ke)^tp_yf-dps*tp_yf"
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=tg+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>px,D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>px,C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + px_at(f"D{r}"))
        for c in (2, 3, 4):
            RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式的目標時點每股價值 − Val 頁情境 B（應為 0）", "=" + px_at("wacc") + f"-{vc('dcf12')}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的目標時點每股價值 − 現價（美元，應接近 0）", f"=E{b0 + 39}-px", "0.000000")
    kv("w_ok", "判讀（隱含 WACC）", f'=IF(ABS(B{A["w_res"]})<0.01,"已收斂","未收斂：搜尋區間內無解")', "@")
    r += 1
    RV.cell(r, 1, "B. 分部加總（Comparable）：支持現價所需的倍數").font = BOLD; r += 1
    hdr(RV, r, ["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"]); r += 1
    ntm = f"({vc('e_ads')}+{vc('e_rt')}+{vc('e_awso')}+{vc('e_ai')})"

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    cmp("sps", "情境 B 分部加總每股價值（美元；目標時點）", "=px", f"={vc('sotp')}", "0.00")
    cmp("sev", "情境 B 分部 EV 合計（US$B；只看分部加總、淨負債與持股不變）", f"=px*sh_val+{vc('nd27')}-{vc('hold')}", f"={vc('sev')}", "#,##0.0")
    cmp("smul", "隱含分部 EV ÷ 未來 12 個月 EBITDA（倍；四個分部 EBITDA 合計，AWS AI 以自身 DCF 計價，含於分子）", f"=B{A['sev']}/{ntm}", f"=C{A['sev']}/{ntm}", "0.0")
    cmp("scmp", "加權（DCF 不變）：支持現價所需的分部加總目標價（情境加權；美元）", f"=(px-w_dcf*Val!$B${VR['dcftp']})/(1-w_dcf)", f"=Val!$B${VR['comptp']}", "0.00")
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F_(T("現金流折現法：情境 B 的企業價值 "), f'TEXT(B{A["ev"]},"#,##0")', T(" 十億美元，目標時點每股價值 "), f'TEXT({vc("dcf12")},"0.00")', T(" 美元；要支持現價 "), 'TEXT(px,"0.00")',
                     T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")', T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("（模型 "), f'TEXT(B{A["w0"]},"0.0%")',
                     T("）；或 FY2028–FY2032 營收與營業利益全面放大 "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                     T("）；或預測期營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F_(T("分部加總法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於情境 B 分部 EV ÷ 未來 12 個月 EBITDA 約 "), f'TEXT(B{A["smul"]},"0.0")', T(" 倍，模型為 "), f'TEXT(C{A["smul"]},"0.0")',
                     T(" 倍；加權法下 DCF 不變時，分部加總目標價只需 "), f'TEXT(B{A["scmp"]},"0.00")', T(" 美元（模型 "), f'TEXT(C{A["scmp"]},"0.00")', T("）即可支持現價。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 76
    for c in "BCDEFG":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["b0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# =====================================================================================================
# Sens：加權目標價相對基準（HTML 龍捲風圖用；只新增列）
LABEL_FIX = {
    "IF_HoldEcon 四世代同時取低／高（每 MW 年持有成本；k_長約的分母同步變動）": "四個 GPU 世代的每 MW 年持有成本同時取低／高（長約價格比的分母同步變動）",
    "Tokenomics IF_NonNV_Hold_Trn（cf_chip、cf_cx 同時跟隨）": "Trainium 每 GW 持有成本比（相對 VR200；收入與資本支出係數同時跟隨）",
    "Trainium 收入係數改用產出比 IF_NonNV_Out_Trn（資本支出仍用持有比）": "Trainium 收入係數改用產出比（資本支出仍用持有比）",
    "k_長約改用長約下緣價格（新產能與期初機隊）": "長約價格比改用長約下緣價格（新產能與期初機隊）",
    "k_現貨改用 AWS 牌價": "現貨價格比改用 AWS 牌價",
}


def sens_tor(wb, C):
    SE = wb["Sens"]
    last = C["SENS_LAST"]
    r0 = last_text_row(SE) + 3
    SE.cell(r0, 1, "加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；本節只引用上方各案例的相對基準（E 欄），只新增、不改既有格；只有一側的案例只列在該側，沒有低／高之分的單一案例依變動方向列在低案或高案側）").font = BOLD
    hdr(SE, r0 + 1, ["輸入（白話）", "輸入取低值", "輸入取高值"])
    r = r0 + 2
    sfx = lambda s: ("低" if s.endswith("：低") else "高" if s.endswith("：高") else "")
    base = lambda s: s[:-2] if sfx(s) else s
    groups = {}
    order = []
    for i in range(4, last + 1):
        a = str(SE.cell(i, 1).value)
        key = base(a)
        if key not in groups:
            groups[key] = {}; order.append(key)
        side = sfx(a) or "單"
        groups[key][side] = i
    for key in order:
        g = groups[key]
        label = LABEL_FIX.get(key, key)
        if "單" in g and ("低" in g or "高" in g):
            label_s = label + "（基準端點）"
            SE.cell(r, 1, label_s)
            SE[f"B{r}"] = f"=IF(Sens!E{g['單']}<0,Sens!E{g['單']},\"\")"; SE[f"C{r}"] = f"=IF(Sens!E{g['單']}>=0,Sens!E{g['單']},\"\")"
            SE[f"B{r}"].number_format = SE[f"C{r}"].number_format = "+0.0%;-0.0%;0.0%"
            assert not PARAM_RE.search(label_s), label_s
            r += 1
        SE.cell(r, 1, label if ("低" in g or "高" in g) else label + "（單一案例）")
        for col, side in (("B", "低"), ("C", "高")):
            if side in g:
                SE[f"{col}{r}"] = f"=Sens!E{g[side]}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        if "低" not in g and "高" not in g:
            SE[f"B{r}"] = f"=IF(Sens!E{g['單']}<0,Sens!E{g['單']},\"\")"; SE[f"C{r}"] = f"=IF(Sens!E{g['單']}>=0,Sens!E{g['單']},\"\")"
            SE[f"B{r}"].number_format = SE[f"C{r}"].number_format = "+0.0%;-0.0%;0.0%"
        assert not PARAM_RE.search(str(SE.cell(r, 1).value)), SE.cell(r, 1).value
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調（讀取層級＋追公式＋C0 實績對帳）
def coord(wb, C):
    QR, HERE, tk, SNAP, F = C["QR"], C["HERE"], C["tk"], C["SNAP"], C["FACTS"]
    W = json.load(open(os.path.join(HERE, "data", "wattonomics_c0_v0.3.2.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    TR = json.load(open(os.path.join(HERE, "data", "wattonomics_trace_v0.3.2.json"), encoding="utf-8"))
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    TKN = f"Tokenomics {C['TKV']}（amazon/data/tokenomics_snapshot_{C['TKV']}.json）"
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJKL", (30, 12, 16, 16, 12, 30, 60, 24, 60, 22, 40)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.3.7 起；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板 Alphabet v0.3.9、Microsoft v0.3.8）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。Amazon 是算力買方兼賣方（AWS 對外出租與自用 AI），不做兩讀數目標價；"
                "C0 的個股側只用公司已公布實績（新聞稿、10-Q），層級側用 Wattonomics v0.3.2 隱含值（玩家資本支出、上線與新增 GW），同源者標「參考」。Amazon 會計年度＝日曆年，CY2026 已公布期間＝2026Q1–Q3（9 個月）。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    hdr(ws, r, ["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"]); r += 1
    F1a = r
    ws.cell(r, 1, "讀取層級：Tokenomics（算力單位經濟）——本模型讀 v21.10 快照（工作單所列現行 v22.0 不在本輪換快照；換快照另走 Tokenomics 同步通知）；Wattonomics——本模型不讀（無 WIF_ 名稱進模型），v0.3.2 快照只用於下方 C0 對帳列").font = BOLD; r += 1

    def f1(name, yr, lv, adopt, unit, layer, note, fmt="0.000"):
        nonlocal r
        ws.cell(r, 1, name); ws.cell(r, 2, yr)
        c = ws.cell(r, 3, lv); c.number_format = fmt
        if isinstance(lv, (int, float)):
            c.font = BLUE
        if adopt is not None:
            ws.cell(r, 4, adopt).number_format = fmt
        else:
            ws.cell(r, 4, "不讀入模型")
        ws.cell(r, 5, unit); ws.cell(r, 6, layer); ws.cell(r, 7, note)
        r += 1
        return r - 1
    F1R = {}
    G4 = (("h100", "H100"), ("gb200", "GB200"), ("gb300", "GB300"), ("vr200", "VR200"))
    rows = []
    for code, nm in G4:
        rows.append((f"IF_HoldEcon｜{nm}", nm, tk("IF_HoldEcon", nm), f"=tk_he_{code}", "US$m/MW-年", "每 MW 年持有成本（經濟口徑；收入錨）"))
    for code, nm in G4:
        rows.append((f"IF_CapexIT｜{nm}", nm, tk("IF_CapexIT", nm), f"=tk_cx_{code}", "US$m/MW", "每 MW IT 資本支出"))
    for code, nm in G4:
        rows.append((f"IF_OpexGW｜{nm}", nm, tk("IF_OpexGW", nm), f"=tk_ox_{code}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    rows.append(("IF_GPUhrEcon｜H100", "H100", tk("IF_GPUhrEcon", "H100"), "=tk_gpuhr", "US$/GPU-hr", "現貨價格比的分母"))
    rows.append(("L1_FacCapexMW", "—", SNAP["items"]["L1_FacCapexMW"]["values"], "=tk_fac", "US$m/MW", "每 MW 機房資本支出"))
    rows.append(("IF_DeprLifeIT", "H100", tk("IF_DeprLifeIT", "H100"), "=tk_life", "年", "AI IT 設備折舊年限"))
    rows.append(("L1_RevGW_PQ_VR200", "VR200", SNAP["items"]["L1_RevGW_PQ_VR200"]["values"], "=tk_pq", "US$m/MW-年", "模型商每 MW 年實收（只作上限檢查）"))
    trn = SNAP["items"]["IF_NonNV_Hold_Trn"]["values"]; tro = SNAP["items"]["IF_NonNV_Out_Trn"]["values"]
    pick = lambda v: v[1] if isinstance(v, list) else (v.get("基準") if isinstance(v, dict) else v)
    rows.append(("IF_NonNV_Hold_Trn（基準）", "Trainium3", pick(trn), "=trn_hold", "x（÷ VR200）", "Trainium 持有成本比（Tokenomics 標 Assumed）"))
    rows.append(("IF_NonNV_Out_Trn（基準）", "Trainium3", pick(tro), "=trn_out", "x（÷ VR200）", "Trainium 產出比（只用於敏感度）"))
    for nm_, yr_, lv, ad, un, nt in rows:
        F1R[nm_] = f1(nm_, yr_, lv, ad, un, TKN, nt)
    F1b = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{F1a + 1}:D{F1b}-C{F1a + 1}:C{F1b}))").number_format = "0.000000"; F1chk = r; r += 1
    ws.cell(r, 1, "Wattonomics v0.3.2 快照名稱與 In 列（只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (
            ("WIF_Gap_Player_amzn_2026_CapexAI", "capai", "US$B", "Amazon 2026 AI 資料中心資本支出（Interface!F622 ← Gap!G154 ← G137＝G134＋G135＋G136：新增 GW × Tokenomics 每 GW 成本＋次年、後年廠房提前；新增 GW 由 CapexGW 從公司資本支出〔In!K181 FY2026 220＝本模型寫回的 2026 指引〕反推，見下方追公式）"),
            ("WIF_Gap_Player_amzn_2026_CapexNonAI", "capnai", "US$B", "非 AI 資本支出（Gap!G144 ← G142 ← G141 ← In!K180 gnai_amzn，本模型 v0.3.1 寫回的 FY2026 非 AI 資本支出；同源）"),
            ("WIF_Gap_Player_amzn_2026_CapexAIModel", "capaim", "US$B", "並列讀法：本模型寫回的 2026 AI 資本支出（In!K179；不進 CapexAI，因 Gap!G153 空白）"),
            ("WIF_Gap_Player_amzn_2026_AddGW", "addgw", "GW", "2026 年新增 IT GW（Gap!G130 ← CapexGW!G96 ← G94〔上線時差加權〕← G93 ← G92＝G91 ÷ 每 GW 成本 ← G83＝G80 × ai_amzn 0.68 ← G80 ← G79 ← G75 ← G73 ← In!K181＝220；同源）"),
            ("WIF_Gap_Player_amzn_2025_AddGW", "add25", "GW", "2025 年新增 IT GW（CapexGW!F94；以 FY2025 資本支出 128.32〔公司現金淨額實績〕換算；只供新增容量 CY2025 參考列）"),
            ("WIF_PlayerGW_YE_amzn_2025", "ye25", "GW", "Amazon 2025 年底上線 IT GW（Interface!E26 ← CapexGW!F97 ← Players!F68 ← F62 ← F61＝In!J100 晶片功耗 1,672.5 MW × D16〔＝ovr25_amzn 2,200 ÷ 晶片功耗〕；Wattonomics C 段證據更正值，與公司資本支出無關）"),
            ("WIF_PlayerGW_YE_amzn_2026", "ye26", "GW", "Amazon 2026 年底上線 IT GW（CapexGW!G97＝2025 年底＋2026 新增 GW〔同源於公司資本支出〕）"),
            ("WIF_PlayerGW_YE_amzn_SiteOnly_2026", "ye26s", "GW", "2026 年底上線 IT GW（只有園區證據路徑）")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt)
    gc = WR["gcap_amzn"]["values"]; gn = WR["gnai_amzn"]["values"]
    wrows["g26"] = f1("In 列 gcap_amzn（K181）｜FY2026 資本支出合計（財年＝日曆年）", "FY2026", gc[3], None, "US$B", WV,
                      "Wattonomics In!K181（Amazon 個股模型 v0.3.1 基準寫回 CF!G54：FY2026 現金淨額資本支出＝C10 校準到公司 2026 指引約 220；Interested-party）；是 CapexGW 2026 年新增 GW 的輸入", "0.0")
    wrows["n26"] = f1("In 列 gnai_amzn（K180）｜FY2026 非 AI 資本支出", "FY2026", gn[3], None, "US$B", WV, "Amazon 個股模型基準寫回（同源）", "0.0")
    wrows["aish"] = f1("In 列 ai_amzn｜資本支出 AI 占比覆寫", "2026", WR["ai_amzn"]["values"][1], None, "比", WV, "AWS 占全公司 PP&E 淨增加比（FY2025：96.5 ÷ 142.4；本模型事實總帳；以 AWS 占比代 AI 占比，偏高）", "0.00")
    wrows["act26"] = f1("In 列 act26_amzn｜2026 已公布三季資本支出實績（現金淨額）", "CY2026Q1–Q3", WR["act26_amzn"]["values"][1], None, "US$B", WV, "Wattonomics 只在 Compare 頁讀（追公式反向查詢），不進 GW 或資本支出路徑", "0.000")
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 Amazon 路徑是否讀公司實績或目標（v0.3.2 快照對應的 20261114_Wattonomics_v0.3.2.xlsx；tools/coord/trace_wattonomics.py --player amzn）").font = BOLD; r += 1
    hdr(ws, r, ["輸出名稱 ← 公式鏈", None, None, None, None, None, "同源判定", "讀到什麼（輸入、公式、儲存格）"]); r += 1
    rdr = sorted({x.split("!")[0] for x in TR["act26_readers"]})
    TRC = [
        ("WIF_Gap_Player_amzn_2026_AddGW ← Interface!F630 ← Gap!G130 ← CapexGW!G96 ← G94（上線時差：lag_on 4.5 個月、inv_on）← G93 ← G92 ← G91 ← G83（＝G80 × ai_amzn）← G80（日曆年；財年結束月 12）← G79 ← G75 ← G73 ← In!K181；F93 ← F83 ← F80 ← In!J181（FY2025 128.32）",
         "同源（本模型寫回的 2026 指引＋公司實績）",
         "In!K181 gcap_amzn＝220（Amazon 個股模型 v0.3.1 CF!G54，以 C10 校準到公司 2026 現金資本支出指引約 220）；In!J181＝128.32（FY2025 現金淨額，公司實績）；× AI 占比 0.68（AWS 占 PP&E 淨增加，取自本模型事實總帳）÷ Tokenomics 每 GW 成本（IT＋廠房，Trainium 比 0.67）。2026 年新增 GW 因此是公司資本支出換算，不是獨立的物理證據"),
        ("WIF_Gap_Player_amzn_2026_CapexAI ← Interface!F622 ← Gap!G154（＝IF(G153<>\"\",G148,G137)；G153 空白，取 G137）← G137＝G134＋G135＋G136 ← G130、H130、I130（2026–2028 新增 GW）", "同源（同上）",
         "新增 GW 同上；廠房提前入帳讀 2027、2028 年新增 GW，其輸入為 In!L181、M181（本模型寫回 FY2027 253.08、FY2028 266.97）"),
        ("WIF_Gap_Player_amzn_2026_CapexNonAI ← Interface!F623 ← Gap!G144 ← G142 ← G141 ← In!K180", "同源（本模型寫回）", "In!K180 gnai_amzn＝68.11（Amazon 個股模型基準 FY2026 非 AI 資本支出）"),
        ("WIF_PlayerGW_YE_amzn_2025 ← Interface!E26 ← CapexGW!F97 ← Players!F68 ← F62 ← F61（＝In!J100 晶片功耗 × D16）← In!I36 ovr25_amzn", "不同源（Wattonomics C 段證據）",
         "2025 年底 2.2 GW＝ovr25_amzn（C 段 5.2 更正：NVIDIA 約 1.0＋Trainium 約 1.2；Derived）；與公司資本支出無關。In!I26 obs_amzn（2026Q2 AI 在役 2.6 GW）引自本模型 v0.1 的事實總帳估計（第三方估計換算），本路徑未生效"),
        ("WIF_PlayerGW_YE_amzn_2026 ← Interface!F26 ← CapexGW!G97＝F97＋G96", "部分同源", "2025 年底（C 段證據，不同源）＋2026 新增 GW（同源於公司資本支出）；公司不揭露年底在役容量存量，存量不做 C0"),
        ("act26_amzn（Wattonomics In!I450；2026 已公布三季資本支出實績 172.188）", "未進 GW／資本支出路徑",
         f"反向查詢：公式讀 act26_amzn 的只有 {'、'.join(rdr)} 頁（{len(TR['act26_readers'])} 格），不回饋 CapexGW、Gap、Interface 的 Amazon 值；本表直接以公司新聞稿的同一實績為個股側"),
    ]
    rtr0 = r
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 2026 年 Amazon 的資本支出（AI＋非 AI）與新增 GW 由公司資本支出（FY2026＝本模型寫回的 2026 指引 220、FY2025 實績 128.32、本模型寫回的 FY2027–FY2028 與非 AI 資本支出）換算，整列同源；"
                 "2025 年底 GW 來自 Wattonomics C 段證據（不同源）；公司實績 act26_amzn 未進路徑（只被 Compare 頁讀）。因此資本支出列標「參考（同源校準）」，結論「無可比實績（層級值同源）」；"
                 "新增 GW 部分同源，依規則改比 GW 物理量：公司 2026 年各季未揭露新增 GW（見 F2），只有 2025 全年約 3.9 GW（全部負載）可作參考。")
    rtr_conc = r; r += 1
    ws.cell(r, 1, f"追公式傾印（生效格與有旗標格；園區旗標格只計數；來源 {TR['source'].split('/')[-1]}；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 4, "值"); ws.cell(r, 7, "標記"); ws.cell(r, 8, "公式（省略開頭等號）"); r += 1
    rdump0 = r
    for x in TR["rows"]:
        ws.cell(r, 1, ("　" * x["depth"]) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 7, (x["flag"] + "；" if x["flag"] else "") + x["status"]); ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    ws.cell(r, 1, f"傾印範圍：生效格與有旗標格共 {len(TR['rows'])} 列；另有園區旗標格 {TR['n_campus_flag']} 格（只計數不列）；全部追到 {TR['n_total']} 格，其中生效 {TR['n_effective']} 格；無法解析的公式 0 格。")
    r += 2
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝新聞稿／10-Q 實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    hdr(ws, r, ["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級", "出處網址"])
    F2hh = r; r += 1
    F2a = r
    SRC_N = f"現金流量表「Purchases of property and equipment, net of proceeds from sales and incentives」（現金淨額；Wattonomics gcap、act26 同口徑）；Q 頁第 {QR['cxn']} 列（2026Q1–Q3、9M26 欄；新聞稿與 10-Q）"
    SRC_G = f"現金流量表「Purchases of property and equipment」（總額）；Q 頁第 {QR['cxg']} 列"
    GR = "Verified（財報新聞稿與 10-Q 現金流量表；事實總帳標 Interested-party 指公司自報）"

    def inp(label, per, unit, f_act, src, grade, layer_f=None, lsrc="", lgrade="", url=U_PR):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, per); ws.cell(r, 3, unit)
        if f_act is not None:
            ws.cell(r, 4, f_act).number_format = "#,##0.000"
        if layer_f is not None:
            ws.cell(r, 5, layer_f).number_format = "#,##0.000"
        ws.cell(r, 8, src); ws.cell(r, 9, grade); ws.cell(r, 10, lsrc); ws.cell(r, 11, lgrade); ws.cell(r, 12, url)
        r += 1
        return r - 1
    rq1 = inp("資本支出 2026Q1（現金淨額）", "CY2026Q1（2026-01～03）", "US$B", f"=Q!K{QR['cxn']}", SRC_N, GR)
    rq2 = inp("資本支出 2026Q2（現金淨額）", "CY2026Q2（04～06）", "US$B", f"=Q!L{QR['cxn']}", SRC_N, GR)
    rq3 = inp("資本支出 2026Q3（現金淨額）", "CY2026Q3（07～09）", "US$B", f"=Q!M{QR['cxn']}", SRC_N, GR)
    rq9 = inp("資本支出 2026 已公布三季累計（現金淨額）", "CY2026Q1–Q3（9 個月）", "US$B", f"=Q!O{QR['cxn']}", SRC_N, GR)
    ws.cell(r, 1, "核對：三季合計 − 9M26 欄；9M26 − Wattonomics act26_amzn（應皆為 0）"); ws.cell(r, 4, f"=ABS(D{rq1}+D{rq2}+D{rq3}-D{rq9})+ABS(D{rq9}-C{wrows['act26']})").number_format = "0.000000"; r += 1
    rg9 = inp("資本支出 2026 已公布三季累計（總額，口徑參考）", "CY2026Q1–Q3（9 個月）", "US$B", f"=Q!O{QR['cxg']}", SRC_G, GR)
    rl = inp("層級 CY2026 玩家資本支出：AI＋非 AI（Wattonomics）", "CY2026（全年）", "US$B", None, "—", "—",
             f"=C{wrows['capai']}+C{wrows['capnai']}", f"{WV}：Gap 頁 Amazon 2026 AI 資本支出（新增 {WN['WIF_Gap_Player_amzn_2026_AddGW']:.2f} GW × Tokenomics 每 GW 成本，含廠房提前）＋非 AI 資本支出（連 F1 兩列）；輸入都是本模型寫回的公司資本支出，屬同源",
             "Derived；輸入同源於公司資本支出", url=f"wattonomics/model/{WS['file']}；快照 {WS['snapshot']}")
    ws.cell(r, 1, "主讀數：已公布期間（CY2026Q1–Q3）累計 vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計實績（現金淨額）；層級 × 9/12（線性時差近似）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq9}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "三季累計（上列）"); ws.cell(r, 9, GR); ws.cell(r, 10, "層級 CY2026 × 9 ÷ 12（Derived）"); ws.cell(r, 11, "Derived")
    rd1 = r; r += 1
    ws.cell(r, 1, "第二讀數：最近一季（CY2026Q3）× 4 vs 層級 CY2026"); ws.cell(r, 2, "CY2026Q3 × 4（約等於該季中點 8 月中的年化速率）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq3}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "CY2026Q3 單季 × 4（Derived）"); ws.cell(r, 9, GR); ws.cell(r, 10, "層級 CY2026 全年"); ws.cell(r, 11, "Derived")
    rd2 = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：資本支出").font = BOLD
    ws.cell(r, 2, "層級值同源於公司資本支出（見追公式）：參考（同源校準），不判定")
    ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由本模型寫回的 2026 指引與公司實績換算＝同源，不是獨立證據；兩個讀數仍列出供參考。消去的是公司資本支出本身，實際只檢驗「層級把 220 換成 GW 再換回資本支出」的時差與每 GW 成本"); ws.cell(r, 9, GR)
    ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL},SIGN(F{rd1})=SIGN(F{rd2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rjr = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：公司 2026 現金資本支出指引 vs 層級 CY2026"); ws.cell(r, 2, "指引（不是實績）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, F["guide.2026.capex"]["value"] / 1000.0).number_format = "#,##0.0"; ws.cell(r, 4).font = BLUE; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（指引，不判定）")
    ws.cell(r, 8, "2026Q3 新聞稿 Financial Guidance 與法說「still expect 2026 cash CapEx of approximately $220 billion」（事實總帳 guide.2026.capex；Interested-party）；指引與層級輸入（In!K181）相同，差距全部來自層級的 GW 換算（上線時差、廠房提前入帳）"); ws.cell(r, 9, "Interested-party（指引）")
    ws.cell(r, 10, "同上"); ws.cell(r, 11, "Derived")
    rg = r; r += 1
    ws.cell(r, 1, "口徑參考（不判定）：資本支出總額三季累計 vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計（總額，含出售與獎勵抵減前）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rg9}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（口徑差，不判定）"); ws.cell(r, 8, "總額比淨額多出售與獎勵抵減（9M 約 3.5）；層級與指引用淨額"); ws.cell(r, 9, GR); ws.cell(r, 10, "同主讀數"); ws.cell(r, 11, "Derived")
    rc9 = r; r += 1
    # ---- AI 營收、在役存量、新增 GW ----
    ws.cell(r, 1, "AI 營收或 run-rate（公司實績）vs 層級隱含 AWS AI 營收"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績（缺 AWS AI 營收揭露）")
    ws.cell(r, 8, F_(T("Amazon 只揭露 AWS 全部營收與年化 run-rate（2026Q3 約 "), f'TEXT({F["aws.runRate.2026q3"]["value"] / 1000.0},"0")', T(" 十億美元，含非 AI 雲端），不揭露 AI 營收或 AI run-rate 的最新值；本模型 AWS AI／AWS 其他的拆分為估計（Assumed），不算公司實績。層級隱含 AI 營收（上線 GW × 對外比例 × Tokenomics 每 GW 營收）的 GW 又同源於公司資本支出，所以不判定")))
    ws.cell(r, 9, "Interested-party（公司 run-rate）"); ws.cell(r, 10, "Tokenomics 每 GW 營收 × Wattonomics GW（未算：無可比實績）"); ws.cell(r, 11, "—"); ws.cell(r, 12, U_PR)
    rai = r; r += 1
    ws.cell(r, 1, "在役／簽約電力 MW 存量（公司實績）"); ws.cell(r, 2, "無可比實績（存量）"); ws.cell(r, 7, "無可比實績（存量）")
    ws.cell(r, 8, "Amazon 不揭露總電力容量或 AI 在役 MW（公司只說「目前容量為 2022 年的兩倍」、目標「2027 年底為 2025 年的兩倍」）；Epoch 站點估計與本模型 AI 在役 MW 為第三方估計或假設，不算公司實績")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "Wattonomics 上線 GW：2025、2026 年底見 F1；存量只供參考"); ws.cell(r, 11, "Derived")
    rmw = r; r += 1
    ws.cell(r, 1, "公司口頭揭露新增電力容量 vs 層級新增 IT GW（新增 GW 部分同源：公司資本支出 ÷ Tokenomics 每 GW 成本，所以改比 GW 物理量）").font = BOLD; r += 1
    GRG = "Interested-party（公司新聞稿與法說；二手轉錄）"
    mw25 = F["mw.co.added2025"]; mwq = F["mw.ai.addSpeed"]
    rgw25 = r
    ws.cell(r, 1, "參考：CY2025 全年新增電力容量（公司）vs 層級 2025 新增 IT GW（不在 CY2026 C0 期間）"); ws.cell(r, 2, "CY2025（全年）"); ws.cell(r, 3, "GW")
    c = ws.cell(r, 4, mw25["value"] / 1000.0); c.number_format = "0.0"; c.font = BLUE
    ws.cell(r, 5, f"=C{wrows['add25']}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"; ws.cell(r, 7, "參考（2025 年，口徑無法對齊，不判定）")
    ws.cell(r, 8, f"2025Q4 法說 Jassy「added 3.9 gigawatts of power」（2025Q3 新聞稿：過去 12 個月 3.8 GW）；全部負載（未分 AI／非 AI）、口徑未說明（推定設施電力）；事實總帳 mw.co.added2025"); ws.cell(r, 9, GRG)
    ws.cell(r, 10, f"{WV}：WIF_Gap_Player_amzn_2025_AddGW（AI IT GW）"); ws.cell(r, 11, "Derived"); ws.cell(r, 12, mw25.get("url", ""))
    r += 1
    rgq = r
    ws.cell(r, 1, "新增容量 CY2026Q1–Q3（公司口頭揭露）"); ws.cell(r, 2, "CY2026：各季未揭露 GW"); ws.cell(r, 3, "GW")
    ws.cell(r, 8, "2026Q1、Q2 法說未揭露新增 GW（事實總帳 mw.ai.addSpeed：2026H1 找不到）；2026Q3 新聞稿與事實總帳無新增 GW 數字（法說逐字稿不在 repo，本輪亦無法取得原文）；公司只重申「2027 年底容量為 2025 年的兩倍」（目標，不是實績）"); ws.cell(r, 9, "—")
    ws.cell(r, 12, mwq.get("url", ""))
    r += 1
    rgl = r
    ws.cell(r, 1, "層級 CY2026 新增 IT GW（Wattonomics）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "GW"); ws.cell(r, 5, f"=C{wrows['addgw']}").number_format = "0.000"
    ws.cell(r, 10, f"{WV}：WIF_Gap_Player_amzn_2026_AddGW；＝公司資本支出 × AI 占比 ÷ 每 GW 成本再依上線時差加權（部分同源）"); ws.cell(r, 11, "Derived；部分同源")
    r += 1
    rgj = r
    ws.cell(r, 1, "C0（實績）判定：新增容量 GW").font = BOLD; ws.cell(r, 2, "公司 2026 年未揭露新增 GW：無可比實績"); ws.cell(r, 7, "無可比實績（缺 2026 年新增 GW 揭露）")
    ws.cell(r, 8, F_(T("CY2026 沒有公司新增 GW 實績可比；CY2025 參考列公司 "), f'TEXT(D{rgw25},"0.0")', T(" GW 對層級 "), f'TEXT(E{rgw25},"0.00")', T(" GW（"), f'TEXT(F{rgw25},"+0%;-0%")',
                    T("），但公司為全部負載、推定設施電力，層級為 AI 的 IT GW，口徑無法對齊，只列參考。公司若揭露 2026 年新增 GW 及其口徑（AI／IT、自有與租用），再依兩讀數規則判定。")))
    ws.cell(r, 9, GRG); ws.cell(r, 10, "Wattonomics 新增 IT GW（部分同源）"); ws.cell(r, 11, "Derived")
    r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "無可比實績列數（資本支出〔層級值同源〕、AI 營收、在役存量、新增 GW）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F_(T("C0（實績）對帳：公司已公布資本支出（現金淨額，2026 年 Q1–Q3 累計 "), f'TEXT(D{rd1},"0.0")', T(" 十億美元）與層級 CY2026 隱含值（"), f'TEXT(E{rl},"0.0")', T("；同期間換算 "), f'TEXT(E{rd1},"0.0")', T("）相差 "),
                     f'TEXT(F{rd1},"+0%;-0%")', T("（主讀數）；最近一季年化 "), f'TEXT(D{rd2},"0.0")', T(" 十億美元，相差 "), f'TEXT(F{rd2},"+0%;-0%")',
                     T("（第二讀數）。但 Wattonomics 的 Amazon 2026 年資本支出與新增 GW 由本模型寫回的 2026 指引（約 "), f'TEXT(C{wrows["g26"]},"0")', T("）換算（追公式：CapexGW!G73 ← In!K181），屬同源，判定："), f"G{rj}",
                     T("（參考，不計入對帳中）。公司 2026 年各季未揭露新增 GW，AWS 也不揭露 AI 營收，在役容量存量不揭露，三項皆無可比實績；對帳中 "), f'TEXT(D{rn},"0")', T(" 項。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向（依季中點）")
    ws.cell(r, 2, F_(T("資本支出是流量，逐季上升（Q1 "), f'TEXT(D{rq1},"0.0")', T("、Q2 "), f'TEXT(D{rq2},"0.0")', T("、Q3 "), f'TEXT(D{rq3},"0.0")',
                      T("）。主讀數用 9 個月平均速率（中點約 5 月中）代表全年，上升趨勢下偏低；第二讀數用 Q3 單季年化（中點約 8 月中），晚於年中，偏高。兩讀數相差 "), f'TEXT((F{rd2}-F{rd1})*100,"0")',
                      T(" 個百分點即時點差的幅度。公司全年指引 "), f'TEXT(D{rg},"0")', T(" 隱含 Q4 約 "), f'TEXT(D{rg}-D{rq9},"0")', T("，與第二讀數的速率相近；層級值低於指引是 GW 換算的上線時差與廠房提前入帳。全部用流量比較，不拿 9 月底存量對年底存量。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出列：追公式（F1b）顯示 Wattonomics 2026 年 Amazon AI 資本支出（Gap!G137）與非 AI 資本支出（Gap!G144）的輸入是本模型寫回的資本支出（In!K181＝220〔C10 校準到公司指引〕、In!K180 非 AI 68.11）；新增 GW 也由此換算。"
                   "層級值因此同源，整列標「參考（同源校準）」，結論「無可比實績（層級值同源）」；比式消去的是公司資本支出，實際只檢驗層級的 GW 換算時差。公司已公布實績 act26_amzn 沒有進 GW 或資本支出路徑（只在 Compare 頁）。"
                   "（二）AI 營收列：AWS 不揭露 AI 營收，層級隱含值又用到同源 GW，不判定。（三）新增 GW 列：層級新增 GW 部分同源，應改比公司揭露的新增 GW 物理量，但 2026 年各季公司未揭露；CY2025 只作參考。"
                   "（四）所以本頁 C0 在 Amazon 對帳中為 0，是因為沒有獨立可比實績，不代表兩層一致。若公司揭露 2026 年新增 GW（AI／IT 口徑），或 Wattonomics 改以 Epoch／園區證據獨立推 2026 年 GW，再依兩讀數規則判定。"))
    rsame = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 輸入的 CY2026 資本支出（In!K181）vs 其 Gap 頁 GW 推算的玩家資本支出（AI＋非 AI；輸出相對輸入）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['g26']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics In!K181（輸入）"); ws.cell(r, 10, "Wattonomics Gap 頁玩家資本支出（輸出）"); rb_ = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：2026 資本支出指引（上列）、Morgan Stanley 新增 GW 估計（事實總帳 mw.ms.*）、Epoch 站點估計不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：本模型讀 Tokenomics v21.10（工作單所列現行 v22.0 未換）；Wattonomics v0.3.2 也讀 Tokenomics v21.10。Amazon 模型值曾被 Wattonomics 讀入（資本支出寫回 In!K179–K181 等，來源為 v0.3.1），本表只列層級值與追公式，不用該回饋判定。")
    r += 1
    return dict(F1=(F1a, F1b), F1chk=F1chk, F2h=F2h, F2hh=F2hh, F2=(F2a, F2b), rq=(rq1, rq2, rq3, rq9), rg9=rg9, rl=rl, rd1=rd1, rd2=rd2, rj=rj, rjr=rjr, rg=rg, rc9=rc9,
                rai=rai, rmw=rmw, rgw=dict(y25=rgw25, q=rgq, l=rgl, j=rgj), rn=rn, rnc=rnc, rs=rs, rr=rr_, rsame=rsame, rb=rb_, rtr=(rtr0, rtr_conc), rdump=(rdump0, rdump1), wrows=wrows, tol=TOL)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, RVA, PLA, CRD):
    RD = wb["README"]
    VR, B, CB = C["VR"], C["B"], C["CB"]
    vb = lambda k, c="B": f"Val!{c}{VR[k]}"
    tp, up, rec, dcf, cmpv, cons = vb("tpw"), vb("upw"), vb("rec"), vb("dcftp"), vb("comptp"), vb("cons")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    PLB = lambda k, c: f"PL!{c}{B[k]}"
    out = {}
    r = last_text_row(RD) + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.7 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F_(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("：現金流折現 "), f'TEXT({dcf},"0")',
        T(" 與分部加總 "), f'TEXT({cmpv},"0")', T(" 各半。AI 資本支出大、基準情境 AWS AI 報酬低於資金成本，是現金流折現偏低的主因。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：自用 AI 採價值中性（讀法 2）；改為只計成本（讀法 1）時加權目標價 "), f'TEXT(Val!B{C["R1"]["tp"]},"0")', T(" 美元（"), f'TEXT(Val!B{C["R1"]["up"]},"+0%;-0%")', T("），評等"),
        f'Val!B{C["R1"]["rec"]}', T("。2026 資本支出依公司指引約 220 十億美元校準 AI 容量速度；現價 "), 'TEXT(px,"0.00")', T(" 美元為 2026-11-13 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；市場共識 "), 'TEXT(cons_tp,"0")', T("。三個情境 "), f'TEXT({vb("tps")},"0")', T("／"), f'TEXT({vb("tps", "C")},"0")', T("／"), f'TEXT({vb("tps", "D")},"0")', T("。"))
    put("o2", T("現金流折現偏低：反向推算（情境 B），要支持現價需把折現率降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), f'TEXT({rv("w0")},"0.0%")', T("），或 FY2028 起營收與營業利益全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍；分部加總隱含 "), f'TEXT(RevDCF!B{RVA["smul"]},"0.0")', T(" 倍 EBITDA，模型 "), f'TEXT(RevDCF!C{RVA["smul"]},"0.0")', T(" 倍。"))
    put("o3", T("AI 報酬：FY2028 對外 AI 投資報酬率 "), f'TEXT({vb("roic28", "C")},"0.0%")', T("，低於 WACC "), 'TEXT(wacc,"0.0%")', T("；AWS 營業利益率 FY2028 "), f'TEXT({vb("awsm28", "C")},"0.0%")', T("（基準）。"))
    put("o4", T("現金：五年自由現金流扣股利合計 "), 'TEXT(CF!C6,"0.0")', T(" 十億美元（基準），FY2028 為 "), 'TEXT(CF!C10,"0.0")', T("；是否需要舉債："), "CF!C8", T("。營業利益率由 FY2025 "), f'TEXT(PL!E{PLA["idx"]["opm"]},"0.0%")', T(" 到 FY2028 "), f'TEXT(PL!H{PLA["idx"]["opm"]},"0.0%")', T("。"))
    put("o5", T("實績對帳（本版新增）：2026 前三季現金資本支出 "), f'TEXT(Coord!D{CRD["rd1"]},"0.0")', T(" 十億美元，與 Wattonomics 同期間隱含值相差 "), f'TEXT(Coord!F{CRD["rd1"]},"+0%;-0%")',
        T("（最近一季年化 "), f'TEXT(Coord!F{CRD["rd2"]},"+0%;-0%")', T("），但層級值由本模型寫回的指引換算（同源），列參考；新增 GW、AI 營收公司未揭露，無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：North America、International＝零售六條營收線（線上商店、實體店、第三方賣家、廣告、訂閱、其他）依年增率推算再按北美比例分配；AWS＝AWS AI（對外 AI 在役 MW × 每 MW 年持有成本 × 價格比）＋AWS 其他（非 AI 雲端，近四季年增向長期值收斂）；四者加總為合併營收。FY2028：North America "),
        f'TEXT({PLB("na", "I")},"0")', T("、International "), f'TEXT({PLB("intl", "I")},"0")', T("、AWS 其他 "), f'TEXT({PLB("aws_o", "I")},"0")', T("、AWS AI "), f'TEXT({PLB("aws_ai", "I")},"0")',
        T("，合計 "), f'TEXT({PLB("tot", "I")},"0")', T(" 十億美元。"))
    RD.cell(r, 1, F_(T("v0.3.8（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 260.21（11-05）→ "), 'TEXT(px,"0.00")',
                     T("；加權目標價 217.69 → "), f'TEXT({tp},"0.00")', T("；上檔空間 −16.3% → "), f'TEXT({up},"+0.0%;-0.0%")',
                     T("；投資建議 中立 → "), rec, T("（只換現價時為 217.69、−18.4%、中立）。本模型現價只影響上檔空間與評等（WACC 權益權重用 In 頁 px_val〔2026-10-30 收盤 240.35，依設計不隨現價換〕）；加權目標價的變動來自同業倍數換價。同業倍數（報價日 2026-10-07／10-08）依價格檔同日收盤換到 2026-11-13（倍數 × 收盤比，近似，無市值）：" + _PXTXT + "；零售同業 WMT、COST、TGT 同法換價（N 批補做，報價日 2026-10-07）；共識與其他假設不變。")))
    out["ver"] = r; r += 1
    RD.cell(r, 1, "v0.3.7（2026-11-15）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁（反向 DCF）、PL 頁末標準損益表"
                  "（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，GAAP；另列 9M26A）、Sens 頁末加權目標價相對基準（龍捲風圖用）、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」（F1 讀取層級；F1b 追公式；F2 C0（實績）：現金資本支出 vs 層級隱含值，層級值同源）。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    r += 1
    return out


_HD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
_PXC = json.load(open(os.path.join(_HD, "peers_cloud_20261113.json"), encoding="utf-8"))
_PXA = json.load(open(os.path.join(_HD, "peers_ads_20261113.json"), encoding="utf-8"))
_PXR = json.load(open(os.path.join(_HD, "peers_retail_20261113.json"), encoding="utf-8"))
_PXTXT = f"非 AI 雲端中位數 {_PXC['median1007']['value']} → {_PXC['median']['value']}、數位廣告 {_PXA['median1007']['value']} → {_PXA['median']['value']}、零售 {_PXR['median1007']['value']} → {_PXR['median']['value']}"


def apply(wb, C):
    PLA = pl_standard(wb, C)
    RVA = revdcf(wb, C)
    tor = sens_tor(wb, C)
    CRD = coord(wb, C)
    RMA = readme13(wb, C, RVA, PLA, CRD)
    return dict(PL=PLA, RV=RVA, tor=tor, CRD=CRD, RM=RMA)
