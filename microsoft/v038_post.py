# Microsoft v0.3.8（2026-11-15）後處理：個股 HTML 版面規格 v1.3 全套＋Coord 頁「三層協調」節（MAG 類樣板；樣板 Alphabet v0.3.9 alphabet/v039_post.py）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點。
# 只新增，不改任何既有儲存格（既有數值、目標價、評等不變；逐格比對只允許版號文字與新增格）：
#   PL 頁末：標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B；GAAP；另列 FY27Q1A 單季比較欄）、比率區、核對列、實績輸入、分部營收（圖表用）
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑，封閉解與二分法）與分部加總隱含倍數；只讀其他頁，無任何頁引用本頁
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用，白話標籤；低案、高案）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v21.10 快照；Wattonomics v0.3.2 只作對帳列、不進模型）、F1b 追公式、F2 C0（實績）：10-Q／IR 公司口徑資本支出 vs 層級隱含值
#   README 頁末：v1.3 展示層白話列（橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字、分部營收推導、版本沿革）
# 會計年度：Microsoft 財年結束於 6 月；日曆年 2026（CY2026）＝FY26Q3（2026-01～03）、FY26Q4（04～06）、FY27Q1（07～09）已公布三季。
# build_microsoft.py 在 README 之後、存檔前呼叫 apply(wb, ctx)。
import json
import os
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3.11", "2026-11-24"
U_IR = "https://www.microsoft.com/en-us/investor/earnings/fy-2027-q1/press-release-webcast（公司 IR 財務報表檔；本 repo 內為 mag/extract/docs/2026Q3/microsoft_FY27Q1_財務報表.xlsx）"   # 出處網址只放 Coord 頁 L 欄（HTML 不顯示外部網址）
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
PARAM_RE = re.compile(r'(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+(?![A-Za-z0-9])')


def lit(t):
    """Excel 公式字串常數上限 255 字元：切段以 & 相接。"""
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F_(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def T(s):
    return ("t", s)


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def last_text_row(ws):
    """最後一個 A 欄有字的列（Sens 頁 max_row 因整本重算區塊而異常大；用此函式）。"""
    r = ws.max_row
    while r > 0 and ws.cell(r, 1).value in (None, ""):
        r -= 1
    return r


# =====================================================================================================
# PL：標準損益表（版面規格 v1.3 D 節）
# 欄位：C FY2024A、D FY2025A、E FY2026A、F–K FY2027E–FY2032E（與 PL 頁情境 B 區塊同欄）、L FY27Q1A（最新單季比較欄）、M 說明
YC = "CDEFGHIJK"
YEAR_OF = {c: 2024 + i for i, c in enumerate(YC)}


def pl_standard(wb, C):
    P, Q = wb["PL"], wb["Q"]
    B, QR, F, fb = C["BR"], C["QR"], C["F"], C["fb"]
    r = last_text_row(P) + 3
    L = {}
    P.cell(r, 1, "標準損益表（v0.3.8 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績）＋FY2027E–FY2032E（情境 B 基準）；最右欄 FY27Q1A 為最新單季實績（比較用）。US$B；EPS 美元。"
                 "只引用本頁上方、Q 頁與下方實績輸入，不回饋任何計算").font = BOLD
    L["start"] = r; r += 1
    P.cell(r, 1, ('="口徑：GAAP（財報口徑）。Microsoft 以 GAAP 揭露損益；本表營業利益、稅前、稅後與 EPS 取 PL 頁情境 B 財報口徑，不含自用 AI 影子效益（評價口徑列見上方情境區塊）。'
                  'FY2026A 其他收入（淨額）含 OpenAI 等持股損益 "&TEXT(Q!M' + str(QR["oai_g"]) + ',"0.0")&" US$B（預測期不估），所以 FY2027E 淨利年增率為負。"'))
    L["note"] = r; r += 1
    P.cell(r, 1, "營業成本與營業費用：模型以分部營業利益與折舊推導營業利益，沒有分列營業成本與營業費用。預測期營業成本＝營收 × 非折舊營業成本率（FY27Q1A 營業成本扣折舊後占營收，下方假設列）＋折舊攤銷合計（AI 折舊加非 AI 折舊；假設全數記入營業成本）；"
                 "FY2027E 為 FY27Q1A 實績加第二至第四季推估；營業費用＝毛利 − 營業利益（殘差，含研發、行銷、管理）。營業利益、稅前、稅後與 EPS 全取自 PL 頁情境 B，不因此改變；歷史年度與 FY27Q1A 為財報實績。")
    L["note2"] = r; r += 1
    hdr = ["項目", "單位", "FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "FY27Q1A（單季）", "說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    L["hdr"] = r; r += 1
    M = "#,##0.000"; PCT = "0.0%"

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in cells.items():
            if v is None:
                continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note:
            P.cell(r, 13, note)
        L[key] = r; r += 1

    PLR = lambda k: B[k][1]
    rows_order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    base = r
    idx = {k: base + i for i, k in enumerate(rows_order)}
    r_ratio = base + len(rows_order)
    ridx = {k: r_ratio + 1 + i for i, k in enumerate(["r_yoy", "r_gm", "r_opm", "r_nm", "r_niy"])}
    r_chk = r_ratio + 1 + 5
    cidx = {k: r_chk + 1 + i for i, k in enumerate(["c1", "c2", "c3", "c4", "c5", "c6"])}
    r_seg = r_chk + 1 + 6 + 1
    sidx = {k: r_seg + 1 + i for i, k in enumerate(["s_ai", "s_dc", "s_tot", "s_chk"])}
    r_h = r_seg + 1 + 4 + 1
    hkeys = ["h_rev", "h_cogs", "h_opx", "h_oi", "h_oie", "h_pti", "h_tax", "h_ni", "h_eps", "h_sh", "h_gm", "h_cex", "h_9", "h_9n"]
    hidx = {k: r_h + 2 + i for i, k in enumerate(hkeys)}

    pl = lambda key, c: f"=PL!{c}{PLR(key)}"
    qq = lambda key: f"=Q!I{QR[key]}"          # FY27Q1A（Q 頁 I 欄）
    # 營收
    row("rev", "營收", "US$B", {"C": f"=C{hidx['h_rev']}", **{c: pl("p_tot", c) for c in "DEFGHIJK"}, "L": qq("rev")}, M, "分部營收加總（PL 頁情境 B 區塊）；FY2024A 為 10-K 實績（下方輸入列）；FY27Q1A 為新聞稿單季")
    assert L["rev"] == idx["rev"]
    yo = {c: f"={c}{idx['rev']}/{p}{idx['rev']}-1" for p, c in zip("CDEFGHIJ", "DEFGHIJK")}
    yo["L"] = f"=L{idx['rev']}/Q!E{QR['rev']}-1"
    row("yoy", "營收年增率（YoY）", "%", yo, PCT, "FY2024A 的基期（FY2023）不在事實總帳，故空白；FY27Q1A 對 FY26Q1（Q 頁）")
    cg = {"C": f"=C{hidx['h_cogs']}", "D": f"=D{hidx['h_cogs']}", "E": f"=E{hidx['h_cogs']}", "L": f"=L{hidx['h_cogs']}",
          "F": f"=L{r}+(F{idx['rev']}-L{idx['rev']})*cogs_ex_r+(PL!F{PLR('p_aida')}+PL!F{PLR('p_nada')}-Q!I{QR['da']})"}
    for c in "GHIJK":
        cg[c] = f"={c}{idx['rev']}*cogs_ex_r+PL!{c}{PLR('p_aida')}+PL!{c}{PLR('p_nada')}"
    row("cogs", "營業成本（COGS）", "US$B", cg, M,
        "歷史與 FY27Q1A＝財報（下方輸入列）；FY2027E＝FY27Q1A＋第二至第四季（營收 × 非折舊成本率＋折舊攤銷）；FY2028E 起＝營收 × 非折舊成本率＋折舊攤銷合計（假設列）")
    row("gp", "毛利", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['cogs']}" for c in "CDEFGHIJKL"}, M)
    row("gm", "毛利率", "%", {c: f"={c}{idx['gp']}/{c}{idx['rev']}" for c in "CDEFGHIJKL"}, PCT, "預測期毛利率下降＝折舊（AI 資本支出）上升")
    ox = {"C": f"=C{hidx['h_opx']}", "D": f"=D{hidx['h_opx']}", "E": f"=E{hidx['h_opx']}", "L": f"=L{hidx['h_opx']}"}
    for c in "FGHIJK":
        ox[c] = f"={c}{idx['gp']}-{c}{idx['oi']}"
    row("opx", "營業費用（研發＋行銷＋管理）", "US$B", ox, M, "歷史與 FY27Q1A＝財報；預測期＝毛利 − 營業利益（殘差）")
    row("oi", "營業利益", "US$B", {"C": f"=C{hidx['h_oi']}", **{c: pl("oi", c) for c in "DEFGHIJK"}, "L": qq("oi")}, M, "財報口徑合併營業利益（PL 頁；不含自用 AI 影子效益）")
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in "CDEFGHIJKL"}, PCT)
    row("oth", "其他收入（淨額）", "US$B", {"C": f"=C{hidx['h_oie']}", **{c: pl("oie", c) for c in "DEFGHIJK"}, "L": qq("oie")}, M, "利息、持股損益與其他；FY2026A 含 OpenAI 等持股損益；預測期持股損益設 0")
    row("pti", "稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}" for c in "CDEFGHIJKL"}, M)
    row("tax", "所得稅", "US$B", {"C": f"=C{hidx['h_tax']}", **{c: pl("tax", c) for c in "DEFGHIJK"}, "L": qq("tax")}, M)
    row("ni", "稅後淨利", "US$B", {c: f"={c}{idx['pti']}-{c}{idx['tax']}" for c in "CDEFGHIJKL"}, M, "稅前 − 所得稅（核對列：與 PL 頁淨利相同）")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in "CDEFGHIJKL"}, PCT)
    row("sh", "稀釋股數", "十億股", {"C": f"=C{hidx['h_sh']}", **{c: pl("sh", c) for c in "DEFGHIJK"}, "L": qq("sh")}, M, "FY2024A＝10-K 稀釋加權平均；預測期取 PL 頁（回購與員工股票發行後的稀釋加權股數）")
    row("eps", "EPS（GAAP 完全稀釋）", "US$", {"C": f"=C{hidx['h_eps']}", **{c: pl("eps_r", c) for c in "DEFGHIJK"}, "L": qq("eps")}, "0.00", "FY2024A＝10-K 公告值；其餘取 PL 頁財報口徑 EPS（淨利 ÷ 稀釋股數）；FY27Q1A＝新聞稿公告值")
    assert L["eps"] == idx["eps"] and r == r_ratio
    P.cell(r, 1, "比率").font = BOLD; r += 1
    row("r_yoy", "營收年增率", "%", {c: f"={c}{idx['yoy']}" for c in "DEFGHIJKL"}, PCT)
    row("r_gm", "毛利率", "%", {c: f"={c}{idx['gm']}" for c in "CDEFGHIJKL"}, PCT)
    row("r_opm", "營業利益率", "%", {c: f"={c}{idx['opm']}" for c in "CDEFGHIJKL"}, PCT)
    row("r_nm", "淨利率", "%", {c: f"={c}{idx['nm']}" for c in "CDEFGHIJKL"}, PCT)
    ny = {c: f"={c}{idx['ni']}/{p}{idx['ni']}-1" for p, c in zip("CDEFGHIJ", "DEFGHIJK")}
    ny["L"] = f"=L{idx['ni']}/Q!E{QR['ni']}-1"
    row("r_niy", "淨利年增率（YoY）", "%", ny, PCT, "FY2026A 含 OpenAI 等持股損益，預測期不估，所以 FY2027E 年增率為負；FY27Q1A 對 FY26Q1")
    assert r == r_chk
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    row("c1", "核對：毛利 − 營業費用 − 營業利益", "US$B", {c: f"={c}{idx['gp']}-{c}{idx['opx']}-{c}{idx['oi']}" for c in "CDEFGHIJKL"}, "0.000", "歷史與 FY27Q1A：財報三項（營業成本、費用、營業利益）互相核對；預測期以殘差定義，恆為 0")
    row("c2", "核對：稅前利益 − PL 頁稅前淨利", "US$B", {c: f"={c}{idx['pti']}-PL!{c}{PLR('pbt')}" for c in "DEFGHIJK"}, "0.000")
    row("c3", "核對：稅後淨利 − PL 頁淨利", "US$B", {c: f"={c}{idx['ni']}-PL!{c}{PLR('ni')}" for c in "DEFGHIJK"}, "0.000")
    row("c4", "核對：營收 − PL 頁合併營收", "US$B", {c: f"={c}{idx['rev']}-PL!{c}{PLR('p_tot')}" for c in "DEFGHIJK"}, "0.000")
    row("c5", "核對：FY2024A 稅後淨利 − 10-K 淨利（輸入列）；FY27Q1A 稅後淨利 − 新聞稿淨利", "US$B", {"C": f"=C{idx['ni']}-C{hidx['h_ni']}", "L": f"=L{idx['ni']}-L{hidx['h_ni']}"}, "0.000")
    row("c6", "核對：實績輸入 − PL 頁／Q 頁（營收、營業利益、淨利三項絕對值合計；FY2025A、FY2026A、FY27Q1A）", "US$B",
        {c: f"=ABS({c}{hidx['h_rev']}-{c}{idx['rev']})+ABS({c}{hidx['h_oi']}-{c}{idx['oi']})+ABS({c}{hidx['h_ni']}-{c}{idx['ni']})" for c in "DEL"}, "0.000")
    assert r == r_seg - 1
    r += 1
    P.cell(r, 1, "分部營收（US$B；圖表用；FY2024A 不在重編後的分部口徑，空白）").font = BOLD; r += 1
    row("s_ai", "Agents and Infra", "US$B", {c: pl("p_ai_seg", c) for c in "DEFGHIJK"}, M)
    row("s_dc", "Devices and Consumer", "US$B", {c: pl("p_dc", c) for c in "DEFGHIJK"}, M)
    row("s_tot", "合計", "US$B", {c: f"={c}{sidx['s_ai']}+{c}{sidx['s_dc']}" for c in "DEFGHIJK"}, M)
    row("s_chk", "核對：合計 − 營收（應為 0）", "US$B", {c: f"={c}{sidx['s_tot']}-{c}{idx['rev']}" for c in "DEFGHIJK"}, "0.000")
    assert r == r_h - 1
    r += 1
    P.cell(r, 1, "實績輸入與假設（藍字＝輸入；FY2024–FY2026 取事實總帳 microsoft/data/microsoft_facts_20261008.json 的 is.*，FY27Q1 取 microsoft/data/microsoft_facts_FY27Q1_20261029.json 的 q.*，皆標 Verified）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "FY2024A", "FY2025A", "FY2026A", "", "", "", "", "", "", "FY27Q1A", "來源"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    src = "Microsoft Form 10-K FY2026（FY2024–FY2026，經查核）；FY27Q1：新聞稿 mag/extract/docs/2026Q3/microsoft_FY27Q1_新聞稿.docx（事實總帳，Verified）"

    def fq(i, k):
        return F[i]["value"] / 1000.0 if i in F and k else None
    items = [
        ("h_rev", "營收", "US$B", "is.rev.fy{}", "q.rev.fy27q1"),
        ("h_cogs", "營業成本", "US$B", "is.cogs.fy{}", "q.cogs.fy27q1"),
        ("h_opx", "營業費用（研發＋行銷＋管理）", "US$B", None, None),
        ("h_oi", "營業利益", "US$B", "is.oi.fy{}", "q.oi.fy27q1"),
        ("h_oie", "其他收入（淨額）", "US$B", "is.oie.fy{}", "q.oie.fy27q1"),
        ("h_pti", "稅前利益", "US$B", "is.pretax.fy{}", "q.pretax.fy27q1"),
        ("h_tax", "所得稅", "US$B", "is.tax.fy{}", "q.tax.fy27q1"),
        ("h_ni", "淨利", "US$B", "is.ni.fy{}", "q.ni.fy27q1"),
    ]
    for k, labl, un, fid, qid in items:
        assert hidx[k] == r, (k, hidx[k], r)
        P.cell(r, 1, labl); P.cell(r, 2, un)
        for col, y in (("C", 2024), ("D", 2025), ("E", 2026)):
            v = (sum(fb(f"is.{x}.fy{y}") for x in ("rd", "sm", "ga")) if k == "h_opx" else fb(fid.format(y)))
            c = P[f"{col}{r}"]; c.value = v; c.number_format = M; c.font = BLUE
        vq = (sum(fb(f"q.{x}.fy27q1") for x in ("rd", "sm", "ga")) if k == "h_opx" else fb(qid))
        c = P[f"L{r}"]; c.value = vq; c.number_format = M; c.font = BLUE
        P.cell(r, 13, src + ("；研發＋行銷＋管理三項相加" if k == "h_opx" else ""))
        r += 1
    assert hidx["h_eps"] == r
    P.cell(r, 1, "EPS（GAAP 稀釋，美元）"); P.cell(r, 2, "US$")
    for col, y in (("C", 2024), ("D", 2025), ("E", 2026)):
        c = P[f"{col}{r}"]; c.value = F[f"is.epsDil.fy{y}"]["value"]; c.number_format = "0.00"; c.font = BLUE
    c = P[f"L{r}"]; c.value = F["q.epsDil.fy27q1"]["value"]; c.number_format = "0.00"; c.font = BLUE
    P.cell(r, 13, src); r += 1
    assert hidx["h_sh"] == r
    P.cell(r, 1, "稀釋加權股數（十億股）"); P.cell(r, 2, "十億股")
    for col, y in (("C", 2024), ("D", 2025), ("E", 2026)):
        c = P[f"{col}{r}"]; c.value = fb(f"is.sharesDilW.fy{y}"); c.number_format = "0.000"; c.font = BLUE
    c = P[f"L{r}"]; c.value = fb("q.sharesDilW.fy27q1"); c.number_format = "0.000"; c.font = BLUE
    P.cell(r, 13, src); r += 1
    assert hidx["h_gm"] == r
    P.cell(r, 1, "毛利（公告值；核對營收 − 營業成本）"); P.cell(r, 2, "US$B")
    for col, y in (("C", 2024), ("D", 2025), ("E", 2026)):
        c = P[f"{col}{r}"]; c.value = fb(f"is.gm.fy{y}"); c.number_format = M; c.font = BLUE
    c = P[f"L{r}"]; c.value = fb("q.gm.fy27q1"); c.number_format = M; c.font = BLUE
    P.cell(r, 13, "10-K／新聞稿公告毛利；與營收 − 營業成本相同（核對：下列）"); r += 1
    assert hidx["h_cex"] == r
    P.cell(r, 1, "非折舊營業成本率（FY27Q1A；＝（營業成本 − 折舊攤銷及其他）÷ 營收；預測期營業成本假設）"); P.cell(r, 2, "%")
    P.cell(r, 3, f"=(L{hidx['h_cogs']}-Q!I{QR['da']})/L{hidx['h_rev']}").number_format = "0.0%"
    P.cell(r, 13, "Derived：假設現金流量表折舊攤銷及其他（Q 頁）全數記入營業成本，其餘營業成本占營收比例不變；Assumed（模型未分列營業成本）")
    dn(wb, "cogs_ex_r", f"PL!$C${r}"); r += 1
    assert hidx["h_9"] == r
    P.cell(r, 1, "核對：營收 − 營業成本 − 公告毛利（FY2024A–FY2026A、FY27Q1A，應為 0）"); P.cell(r, 2, "US$B")
    for col in "CDEL":
        P[f"{col}{r}"] = f"={col}{hidx['h_rev']}-{col}{hidx['h_cogs']}-{col}{hidx['h_gm']}"; P[f"{col}{r}"].number_format = "0.000"
    r += 1
    assert hidx["h_9n"] == r
    P.cell(r, 1, "核對：稅前 − 所得稅 − 淨利（輸入列；應為 0）"); P.cell(r, 2, "US$B")
    for col in "CDEL":
        P[f"{col}{r}"] = f"={col}{hidx['h_pti']}-{col}{hidx['h_tax']}-{col}{hidx['h_ni']}"; P[f"{col}{r}"].number_format = "0.000"
    r += 1
    for col in "CDEFGHIJKL":
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(idx=idx, ridx=ridx, cidx=cidx, sidx=sidx, hidx=hidx, ratio=r_ratio, chk=r_chk, seg=r_seg, h=r_h)
    return L


# =====================================================================================================
# RevDCF：反向 DCF（版面規格 v1.3；列＝A–G，一頁）
def revdcf(wb, C):
    BR = C["BR"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    vc = lambda k: f"Val!$C${BR[k][1]}"                # 引擎 Val 區塊的純量（C 欄）
    vy = lambda k, c: f"Val!{c}{BR[k][1]}"
    pl_ = lambda k, c: f"PL!{c}{BR[k][1]}"
    cols = ["F", "G", "H", "I", "J", "K"]               # FY2027–FY2032
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.8 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益為線性，用封閉解；WACC 用二分法）。本頁只讀 Val、PL、CF、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 每股價值推到目標時點＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋WACC）^年數 − 期間股利；自用 AI 影子效益依 Val 頁讀法）；"
                "分部加總部分反解「情境 B 分部加總每股價值＝現價」。反解值超出合理範圍時照實列出，不外推。營業利益率與營收規模兩項只動 FY2028 起的財報口徑營業利益（FY2027 為 Q1 實績加指引錨點不動；折舊、資本支出、營運資金、自用 AI 影子效益不變）")
    r = 4
    RV.cell(r, 1, "A. DCF（情境 B 路徑）：支持現價所需的企業價值").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "說明"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    A = {}

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f).number_format = fmt
        if note:
            RV.cell(r, 3, note)
        A[key] = r; r += 1
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", f"=(px+{vc('divroll')})/(1+wacc)^t_roll", "0.00", "＝（現價＋評價日到目標時點的每股股利）÷（1＋WACC）^目標時點年數（Val 頁 DCF 每股價值推到目標時點的反算）")
    kv("sh", "股數（十億股；評價日稀釋）", "=shd_val", "0.000")
    kv("oth", "淨現金、持股減在帳融資租賃負債（US$B）", "=cash_val-debt_val-fl_val+oai_v", "#,##0.0", "同 Val 頁 DCF 股東權益橋的非 EV 項（淨現金＋OpenAI 持股 − 融資租賃負債）")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{A['v0']}*B{A['sh']}-B{A['oth']}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；主值）", f"={vc('ev')}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=(B{A['ev']}+B{A['oth']})/B{A['sh']}", "0.00")
    kv("v_chk", "核對：模型情境 B DCF 每股價值推到目標時點 − Val 頁情境 B（應為 0）", f"=B{A['v_b']}*(1+wacc)^t_roll-{vc('divroll')}-{vc('dcf12')}", "0.000000")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, vcol in enumerate(cols):
            RV.cell(r, 2 + i, fn(vcol, "BCDEFG"[i])).number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$B）", lambda v, c: f"={pl_('p_tot', v)}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$B；財報口徑）", lambda v, c: f"={pl_('oi', v)}", "#,##0.0")
    yr("m", "營業利益率", lambda v, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("df", "折現因子", lambda v, c: f"={vy('df', v)}", "0.0000")
    yr("tvf", "終值係數（只在 FY2032：(1＋g)(1 − g ÷ (WACC＋新投資報酬差)) ÷ (WACC − g) × 終值折現因子）",
       lambda v, c: ("=(1+g_tv)*(1-g_tv/(wacc+ronic_prem))/(wacc-g_tv)/(1+wacc)^" + vy('tend', 'K') if v == "K" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（FY2027 為錨點不動，記 0；FY2028 起 ＝(1 − 稅率) × (折現因子＋終值係數)）", lambda v, c: ("=0" if v == "F" else f"=(1-tax_f)*({c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(B{Y['rev']}:G{Y['rev']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與 FY2028 起營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032，以同一係數加權）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "終值年（FY2032）權重大")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(C{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收與營業利益等比例放大；營業利益率不變，折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("EB", "FY2028–FY2032 營業利益 × 貢獻係數合計（US$B；放大一倍所增加的 EV）", f"=SUMPRODUCT(B{Y['ebit']}:G{Y['ebit']},B{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=1+(B{A['evr']}-B{A['ev']})/B{A['EB']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年現金流不變；永續期新投資報酬＝WACC＋差，同 Val 頁；目標時點每股價值含 (1＋WACC)^t；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的 DCF 每股價值推到目標時點（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def px_at(w):
        fcf, tm = f"Val!$F${BR['fcf_val'][1]}:$K${BR['fcf_val'][1]}", f"Val!$F${BR['tm'][1]}:$K${BR['tm'][1]}"
        ev = (f"SUMPRODUCT({fcf},1/(1+{w})^{tm})+{vc('nopat')}*(1-g_tv/({w}+ronic_prem))/({w}-g_tv)/(1+{w})^Val!$K${BR['tend'][1]}")
        return f"(({ev})+$B${A['oth']})/$B${A['sh']}*(1+{w})^t_roll-{vc('divroll')}"
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
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
    r += 1
    RV.cell(r, 1, "B. 分部加總（Comparable）：支持現價所需的倍數").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ssh = f"(CF!$F${BR['sh_end'][1]}+dil_gap)"
    divq = f"CF!$G${BR['dps'][1]}*roll_div_q/q_per_yr"
    ntm = f"(w_ntm28*({vy('e_aiseg', 'G')}+{vy('e_dc', 'G')})+(1-w_ntm28)*({vy('e_aiseg', 'H')}+{vy('e_dc', 'H')}))"

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    cmp("sps", "情境 B 分部加總每股價值（美元；目標時點）", "=px", f"={vc('sotp')}", "0.00")
    cmp("sev", "情境 B 分部加總 EV（US$B；只看分部加總、其他項不變）", f"=(px+{divq})*{ssh}-{vc('sbr')}", f"={vc('sev')}", "#,##0.0")
    cmp("smul", "隱含 EV ÷ NTM EBITDA（倍；NTM＝FY2028、FY2029 依目標時點加權；AI 基礎設施模組以自身 DCF 計價，含於分子）", f"=B{A['sev']}/{ntm}", f"=C{A['sev']}/{ntm}", "0.0")
    cmp("scmp", "加權（DCF 不變）：支持現價所需的分部加總每股價值（情境加權；美元）", f"=(px-w_dcf*Val!$B${C['VR']['dcftp']})/(1-w_dcf)", f"=Val!$B${C['VR']['comp']}", "0.00")
    RV.cell(r, 1, "同業倍數見 In 頁 B 節與 Val 頁 0.2 節；此處只列整體隱含倍數，不拆到分部"); r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F_(T("現金流折現法：情境 B 的企業價值 "), f'TEXT(B{A["ev"]},"#,##0")', T(" 十億美元，目標時點每股價值 "), f'TEXT({vc("dcf12")},"0.00")', T(" 美元；要支持現價 "), 'TEXT(px,"0.00")',
                     T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")', T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("（模型 "), f'TEXT(B{A["w0"]},"0.0%")',
                     T("）；或 FY2028–FY2032 營收與營業利益全面放大 "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                     T("）；或預測期營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                     f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F_(T("分部加總法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於情境 B 分部加總 EV ÷ NTM EBITDA 約 "), f'TEXT(B{A["smul"]},"0.0")', T(" 倍（含以自身現金流折現的 AI 基礎設施模組），模型整體隱含 "), f'TEXT(C{A["smul"]},"0.0")',
                     T(" 倍；換言之，只看分部加總，市場價格對應的整體倍數比模型 "), f'IF(B{A["smul"]}<C{A["smul"]},"低","高")', f'TEXT(ABS(D{A["smul"]}),"0%")', T("（EV 差距）。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 76
    for c in "BCDEFG":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["b0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# =====================================================================================================
# Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列）
LABEL_FIX = {
    "IF_HoldEcon 四世代同時（每 MW 年持有成本）": "四個 GPU 世代的每 MW 年持有成本同時取低／高",
    "IF_NonNV_Hold_Maia": "Maia 每 MW 持有成本比（÷ VR200）",
    "現貨證據價格（高＝Azure 隨需牌價）": "現貨證據價格（Azure 隨需牌價）",
}


def sens_tor(wb, C):
    SE = wb["Sens"]
    last = C["SENS_LAST"]
    r0 = last_text_row(SE) + 3
    SE.cell(r0, 1, "加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；本節只引用上方各案例的相對基準（E 欄），只新增、不改既有格；只有一側的案例只列在該側）").font = BOLD
    for j, h in enumerate(["輸入（白話）", "輸入取低值", "輸入取高值"], 1):
        c = SE.cell(r0 + 1, j, h); c.font = BOLD; c.fill = HDR
    r = r0 + 2
    sfx = lambda s: ("低" if s.endswith("：低") else "高" if s.endswith("：高") else "")
    i = 4
    while i <= last:
        a = str(SE.cell(i, 1).value); b = str(SE.cell(i, 2).value)
        nxt = i + 1 <= last
        a2 = str(SE.cell(i + 1, 1).value) if nxt else ""
        b2 = str(SE.cell(i + 1, 2).value) if nxt else ""
        low = high = None; label = None
        if nxt and b == b2 and sfx(a) == "低" and sfx(a2) == "高" and a[:-2] == a2[:-2]:
            low, high, label = i, i + 1, a[:-2]
        elif nxt and b == b2 and not sfx(a) and not sfx(a2):
            low, high, label = i, i + 1, f"{a}／{a2}"
        elif sfx(a) == "高":
            high, label = i, a[:-2]
        elif sfx(a) == "低":
            low, label = i, a[:-2]
        else:
            low, label = i, a
        i += 2 if (low and high) else 1
        label = LABEL_FIX.get(label, label)
        label = re.sub(r"（IF_[A-Za-z_]+ [低高]）", "", label)
        SE.cell(r, 1, label)
        for col, rr in (("B", low), ("C", high)):
            if rr:
                SE[f"{col}{r}"] = f"=Sens!E{rr}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        assert not PARAM_RE.search(label), label
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調（讀取層級＋C0 實績對帳；MAG 類樣板——Alphabet 樣板複製）
def coord(wb, C):
    GEN, QR, SNAP, TKV, TKN, tk, tkv, HERE, F, fb = C["GEN"], C["QR"], C["SNAP"], C["TKV"], C["TKN"], C["tk"], C["tkv"], C["HERE"], C["F"], C["fb"]
    W = json.load(open(os.path.join(HERE, "data", "wattonomics_c0_v0.3.2.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    TR = json.load(open(os.path.join(HERE, "data", "wattonomics_trace_v0.3.2.json"), encoding="utf-8"))
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    tcell = lambda cell: next(x for x in TR["rows"] if x["cell"] == cell)
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJKL", (30, 12, 16, 16, 12, 28, 60, 22, 60, 22, 40)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.3.8 起；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。Microsoft 是算力買方兼賣方（Azure 對外出租與自用 AI），不做兩讀數目標價；"
                "C0 的個股側只用公司已公布實績（新聞稿／10-Q／IR 財務報表），層級側用 Wattonomics v0.3.2 隱含值（玩家資本支出、上線 GW）與 Tokenomics 每 GW 營收錨，同源者標「參考」。"
                "會計年度：Microsoft 財年結束於 6 月；層級是日曆年，所以 CY2026 已公布期間＝FY26Q3（2026-01～03）、FY26Q4（04～06）、FY27Q1（07～09）三季。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    for j, h in enumerate(["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1a = r
    ws.cell(r, 1, "讀取層級：Tokenomics（算力單位經濟）——本模型讀 " + TKV + " 快照（與 Tokenomics 現行版相同；外部稽核優化 W3-5，2026-11-23 由 v21.10 換入）；Wattonomics——本模型不讀（無 WIF_ 名稱進模型），v0.3.2 快照只用於下方 C0 對帳列").font = BOLD; r += 1

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
    tkrows = []
    for code, g in (("H100", "h100"), ("GB200", "gb200"), ("GB300", "gb300"), ("VR200", "vr200")):
        tkrows.append((f"IF_HoldEcon｜{code}", code, tk("IF_HoldEcon", g), f"=tk_hold_{g}", "US$m/MW-年（＝$B／GW／年）", "每 MW 年持有成本（經濟口徑；收入錨 k＝1）"))
    tkrows.append(("IF_GPUhrEcon｜H100", "H100", tk("IF_GPUhrEcon", "h100"), "=tk_gpuhr_h100", "US$/GPU-hr", "k 的現貨分母"))
    for code, g in (("H100", "h100"), ("GB200", "gb200"), ("GB300", "gb300"), ("VR200", "vr200")):
        tkrows.append((f"IF_OpexGW｜{code}", code, tk("IF_OpexGW", g), f"=tk_opex_{g}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    for code, g in (("H100", "h100"), ("GB200", "gb200"), ("GB300", "gb300"), ("VR200", "vr200")):
        tkrows.append((f"IF_CapexIT｜{code}", code, tk("IF_CapexIT", g), f"=tk_cxit_{g}", "US$m/MW", "每 MW IT 資本支出"))
    tkrows.append(("IF_CapexFacility｜VR200", "VR200", tk("IF_CapexFacility", "vr200"), "=tk_cxfac", "US$m/MW", "每 MW 機房資本支出"))
    tkrows.append(("IF_NonNV_Hold_Maia（基準）", "Maia 200", tkv("IF_NonNV_Hold_Maia", 1), "=nv_hold_mai", "x（÷ VR200）", "Maia 持有成本比（Tokenomics 標 Assumed）"))
    tkrows.append(("IF_NonNV_Out_Maia（基準）", "Maia 200", tkv("IF_NonNV_Out_Maia", 1), "=nv_out_mai", "x（÷ VR200）", "Maia 產出比（Tokenomics 標 Assumed）"))
    for nm_, yr_, lv, ad, un, nt in tkrows:
        F1R[nm_] = f1(nm_, yr_, lv, ad, un, TKN, nt, "0.000")
    F1b = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{F1a + 2}:D{F1b}-C{F1a + 2}:C{F1b}))").number_format = "0.000000"; F1chk = r; r += 1
    ws.cell(r, 1, "Wattonomics v0.3.2 快照名稱與 In 列（只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (
            ("WIF_Gap_Player_msft_2026_CapexAI", "capai", "US$B", "Microsoft 2026 AI 資料中心資本支出（Interface!F596 ← Gap!G62 ← G45＝G42＋G43＋G44：新增 GW × Tokenomics 每 GW 成本＋次年、後年廠房提前；新增 GW 由 CapexGW 從公司資本支出（FY2026 實績 145.3、FY2027 個股模型寫回 204.2）反推，見下方追公式）"),
            ("WIF_Gap_Player_msft_2026_CapexNonAI", "capnai", "US$B", "非 AI 資本支出（Gap!G52＝G50＝（In!K168＋In!L168）× 6/12，讀個股模型寫回的 FY2026、FY2027 非 AI 資本支出；同源）"),
            ("WIF_Gap_Player_msft_2026_AddGW", "addgw", "GW", "2026 年新增 IT GW（Gap!G38←CapexGW!G34←G32←G31←G30←G29←G21＝G18 × AI 占比；輸入為公司資本支出 FY2026 實績 145.3 與個股模型寫回 FY2027 204.2，見追公式；不是 Epoch 或園區證據推得）"),
            ("WIF_Gap_Player_msft_2025_AddGW", "add25", "GW", "2025 年新增 IT GW（Interface!E604←Gap!F38←CapexGW!F34←F33←Players!F36；2025 年走 Players 路徑〔Epoch 晶片功耗〕，與公司資本支出無關；只供 Coord 新增容量 CY2025 參考列）"),
            ("WIF_PlayerGW_YE_msft_2025", "ye25", "GW", "Microsoft 2025 年底上線 IT GW（Interface!E24←CapexGW!F35←Players!F30←F24←F23＝Epoch 晶片功耗 In!J98 2,015.7 MW × k_chip2it 1.85；與公司資本支出無關，不同源）"),
            ("WIF_PlayerGW_YE_msft_2026", "ye26", "GW", "Microsoft 2026 年底上線 IT GW（CapexGW!G35＝2025 年底（Epoch，不同源）＋2026 新增 GW（同源於公司資本支出，見追公式））"),
            ("WIF_PlayerGW_YE_msft_SiteOnly_2026", "ye26s", "GW", "2026 年底上線 IT GW（只有園區證據路徑）"),
            ("WIF_Gap_Player_msft_2027_CapexAIMC", "cap27", "US$B", "並列讀法：2027 AI 資本支出（個股模型值；2026 無此值）")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt, "0.000")
    wrows["g26"] = f1("In 列 gcap_msft（K169）｜FY2026（資本支出，財年）", "FY2026", WR["gcap_msft"]["values"][3], None, "US$B", WV,
                      "Wattonomics In!K169（gcap_msft，Microsoft 個股模型基準寫回：FY2026＝10-K／法說的公司口徑資本支出 145.3，含融資租賃；Interested-party）；是 CapexGW 2026 年新增 GW 的輸入", "0.0")
    wrows["g27"] = f1("In 列 gcap_msft（L169）｜FY2027（資本支出，財年）", "FY2027", WR["gcap_msft"]["values"][4], None, "US$B", WV,
                      "Wattonomics In!L169（Microsoft 個股模型 v0.3.3 基準寫回 FY2027 全年 204.2；公司指引推估，不是實績）；也是 CapexGW 2026 日曆年換算（財年結束月 6）的輸入", "0.0")
    wrows["gcy"] = f1("CapexGW!G18｜CY2026 日曆年資本支出（輸入）", "CY2026", tcell("CapexGW!G18")["value"], None, "US$B", WV,
                      "CapexGW!G18＝G17 × 財年結束月 ÷ 12＋H17 ×（12 − 財年結束月）÷ 12＝FY2026 的 6/12＋FY2027 的 6/12（公司資本支出換成日曆年）；之後 × AI 占比 0.85 進 AI 路徑、× 0.15 進非 AI 路徑（本值由追公式傾印取得）", "0.00")
    wrows["self"] = f1("In 列 al_msft_self｜自用比例", "2026", WR["al_msft_self"]["values"][3], None, "比", WV, "自用＝1 − 對外 60%（Wattonomics 沿用 Microsoft 舊模型；Assumed）；對外比例與本模型 ext_sh 同源", "0.00")
    wrows["chip"] = f1("In 列 chip_msft（J98）｜2025 年底 Epoch 晶片功耗", "2025", WR["chip_msft"]["values"][2], None, "MW", WV, "Epoch AI Chip Owners（NVIDIA＋AMD；累計未扣退役；Analogy）；× k_chip2it 1.85 ＝ 2025 年底上線 IT GW", "#,##0.0")
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 Microsoft 路徑是否讀公司實績或目標（v0.3.2 快照對應的 20261114_Wattonomics_v0.3.2.xlsx；規則：有讀者的部分屬同源，不當獨立證據）").font = BOLD; r += 1
    for j, h in enumerate(["輸出名稱 ← 公式鏈", "", "", "", "", "", "同源判定", "讀到什麼（輸入、公式、儲存格）"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    nrd = len(TR["act26_readers"])
    TRC = [
        ("WIF_Gap_Player_msft_2026_AddGW ← Interface!F604 ← Gap!G38 ← CapexGW!G34 ← G32（上線年時差加權；lag_on＝4.5 個月、inv_on）← G31 ← G30 ← G29 ← G21（＝G18 × ai_sh）← G18（日曆年換算）← G17 ← G13 ← G11 ← In!K169；H17 ← H13 ← H11 ← In!L169；F13 ← F12 ← In!J121；E13 ← E12 ← In!I121",
         "同源（公司資本支出實績＋個股模型寫回）",
         "公司實績：FY2026 資本支出 145.3（In!K169 gcap_msft，個股模型基準寫回，＝10-K／法說的公司口徑含融資租賃）、FY2025 64.551（In!J121）、FY2024 44.477（In!I121，現金 PP&E，10-K）；"
         "個股模型寫回：FY2027 204.2（In!L169，Microsoft 個股模型基準：Q1 實績＋指引推估）。G18＝145.3 × 6/12＋204.2 × 6/12＝174.75；× AI 占比 ai_sh（In!I13，0.85，Assumed）進 AI 路徑。"
         "GW＝資金 ÷ 每 GW 成本（Tokenomics），所以 2026 年新增 GW 是公司資本支出換算，不是獨立的物理證據"),
        ("WIF_Gap_Player_msft_2026_CapexAI ← Interface!F596 ← Gap!G62（＝IF(G61<>\"\",G56,G45)；G61 空白，取 G45）← G45＝G42＋G43＋G44（G42＝2026 新增 GW Gap!G38 ×（每 GW IT 成本 G39＋廠房 G40 × D6）；G43＝E6 × H38 × H40、G44＝F6 × I38 × I40＝2027、2028 年廠房提前）← CapexGW!H34、I34 ← In!L169、M169、N169", "同源（同上）",
         "新增 GW 同上；廠房提前入帳讀 CapexGW 2027、2028 年新增 GW（Gap!H38、I38），其輸入來自 In!L169、M169、N169（個股模型寫回 FY2027–FY2029）；Gap!G56（個股模型 AI 資本支出曆年＝G55 × 6/12＋H55 × 6/12；G55 ← In!K167 空白、H55 ← In!L167 168.45）因 G55 空白而 2026 空白，G62 不改取它"),
        ("WIF_Gap_Player_msft_2026_CapexNonAI ← Interface!F597 ← Gap!G52 ← G50（＝G49 × 6/12＋H49 × 6/12）← G49 ← In!K168；H49 ← In!L168", "同源（個股模型寫回）",
         "In!K168（FY2026 非 AI 資本支出 30.18）與 In!L168（FY2027 35.75）都是 Microsoft 個股模型基準寫回值（gnai_msft）；日曆年換算＝兩者各半＝32.97；Gap!G50 有值，所以不走 CapexGW!G18 × 0.15 的備援路徑"),
        ("WIF_PlayerGW_YE_msft_2025 ← Interface!E24 ← CapexGW!F35 ← Players!F30 ← F24 ← F23（＝In!J98 chip_msft × Players!D14）", "不同源（Epoch 晶片功耗 × 倍數）",
         "2025 年底 3.729 GW＝Epoch 晶片功耗 2,015.7 MW（In!J98）× k_chip2it 1.85（In!I7，Analogy）；與公司資本支出無關，也不讀公司實績（obs_msft 備援分支未生效）"),
        ("WIF_PlayerGW_YE_msft_2026 ← Interface!F24 ← CapexGW!G35＝F35＋G34", "部分同源",
         "2025 年底（Epoch 推得，不同源）＋2026 年增量 G34（同源於公司資本支出，見第一列）。層級另有 SiteOnly 路徑 WIF_PlayerGW_YE_msft_SiteOnly_2026（Epoch 加園區證據，4.5 GW）；公司不揭露年底在役容量存量，存量不做 C0。2026 年增量 G34 部分同源（公司資本支出 ÷ Tokenomics 每 GW 成本），依規則改比每 GW／新增 GW 物理量：公司法說口頭揭露各季新增容量（見 F2「公司口頭揭露新增容量」列）"),
        ("act26_msft（Wattonomics In 頁 act26_msft；2026 年已公布三季資本支出實績 118.1）", "未進 GW／資本支出路徑",
         "反向查詢：公式讀 act26_msft 的只有 " + "、".join(sorted({x.split("!")[0] for x in TR["act26_readers"]})) + " 頁（" + str(nrd) + " 格；Compare 頁第 227 列），不回饋 CapexGW、Gap、Interface 的 Microsoft 值；本表直接以公司新聞稿與 IR 財務報表的同一實績為個股側，不經 Wattonomics"),
        ("WIF_Gap_Player_msft_2026_CapexAIModel ← Interface!F601 ← Gap!G56 ← G55 ← In!K167", "空白（無回寫值）",
         "2026 年個股模型 AI 資本支出 In!K167 為空（FY2026 未拆 AI／非 AI 寫回），所以這一欄不替代 GW 推算；2027 年起的並列讀法才讀個股模型（本表不用）"),
    ]
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 2026 年 Microsoft 的資本支出（AI＋非 AI）與新增 GW 由公司資本支出（FY2026 實績 145.3、FY2025 與 FY2024 現金 PP&E、個股模型寫回的 FY2027 204.2 與非 AI 資本支出）換算，整列同源；2025 年底 GW 來自 Epoch，不同源；公司實績 act26_msft 未進路徑（只被 Compare 頁讀）。"
                 "因此資本支出列標「參考（同源校準）」，結論「無可比實績（層級值同源）」。新增 GW 是部分同源（公司資本支出 ÷ Tokenomics 每 GW 成本，加上線時差），依規則改比每 GW 物理量："
                 "公司法說逐季口頭揭露新增容量（FY26Q2 約 1 GW、FY26Q3 再 1 GW、FY26Q4 31 座資料中心約 1 GW；FY27Q1 只說新增 24 座資料中心、未給 GW），見 F2「公司口頭揭露新增容量」列；"
                 "這個比較消去的是公司資本支出，實際檢驗的是 Tokenomics 每 GW 成本、AI 占比與上線時差。證據回饋走對帳單（C5）。")
    rtr_conc = r; r += 1
    ws.cell(r, 1, "追公式傾印（由 tools/coord/trace_wattonomics.py --player msft 產生，貼生效格與有旗標格；園區旗標格只計數；來源 " + TR["source"].split("/")[-1] + "；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 7, "標記"); ws.cell(r, 8, "公式（省略開頭等號）"); ws.cell(r, 4, "值"); r += 1
    rdump0 = r
    for x in TR["rows"]:
        ws.cell(r, 1, ("　" * x["depth"]) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 7, (x["flag"] + "；" if x["flag"] else "") + x["status"]); ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    ws.cell(r, 1, f"傾印範圍：生效格與有旗標格共 {len(TR['rows'])} 列；另有園區旗標格 {TR['n_campus_flag']} 格（Pipeline 園區列與園區輸入列，條件讀取或分支未生效，只計數不列）；全部追到 {TR['n_total']} 格，其中生效 {TR['n_effective']} 格；無法解析的公式 0 格。")
    rdump_note = r; r += 2
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝新聞稿／10-Q／IR 財務報表實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    for j, h in enumerate(["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級", "出處網址"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    F2hh = r; r += 1
    SRC_CO = ("公司口徑資本支出（含融資租賃取得的資產；公司 IR 財務報表檔 CapEx 工作表「Capital expenditures including assets acquired under capital leases」逐季：FY26Q3 31.9、FY26Q4 41.0、FY27Q1 45.2；"
              "mag/extract/docs/2026Q3/text/microsoft_FY27Q1_財務報表.txt 的 sheet CapEx（第 312 行起）；Q 頁第 " + str(QR["capex_co"]) + " 列）")
    GR_CO = "Verified（公司 IR 財務報表原檔；口徑為公司自訂含融資租賃，Q 頁標 Interested-party 指此口徑）"
    SRC_CASH = "現金流量表「Additions to property and equipment」（現金 PP&E）：FY27Q1 10-Q 第 537 行 39,367；FY26Q3、FY26Q4 各季新聞稿現金流量表（30.876、35.802）；Q 頁第 " + str(QR["capex"]) + " 列"
    GR_CASH = "Verified（10-Q／新聞稿財報表）"
    F2a = r

    def inp(label, per, unit, f_act, src, grade, layer_f=None, lsrc="", lgrade="", url=U_IR):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, per); ws.cell(r, 3, unit)
        ws.cell(r, 4, f_act).number_format = "#,##0.000"
        if layer_f is not None:
            ws.cell(r, 5, layer_f).number_format = "#,##0.000"
        ws.cell(r, 8, src); ws.cell(r, 9, grade); ws.cell(r, 10, lsrc); ws.cell(r, 11, lgrade); ws.cell(r, 12, url)
        r += 1
        return r - 1
    rq1 = inp("資本支出 CY2026Q1＝FY26Q3（公司口徑）", "CY2026Q1（2026-01～03）", "US$B", f"=Q!G{QR['capex_co']}", SRC_CO, GR_CO)
    rq2 = inp("資本支出 CY2026Q2＝FY26Q4（公司口徑）", "CY2026Q2（04～06）", "US$B", f"=Q!H{QR['capex_co']}", SRC_CO, GR_CO)
    rq3 = inp("資本支出 CY2026Q3＝FY27Q1（公司口徑）", "CY2026Q3（07～09）", "US$B", f"=Q!I{QR['capex_co']}", SRC_CO, GR_CO)
    rq9 = inp("資本支出 CY2026 已公布三季累計（公司口徑）", "CY2026Q1–Q3（9 個月）", "US$B", f"=D{rq1}+D{rq2}+D{rq3}", SRC_CO, GR_CO)
    ws.cell(r, 1, "核對：三季累計 − Wattonomics act26_msft（公司已公布實績輸入 118.1；應為 0）"); ws.cell(r, 4, f"=D{rq9}-{WR['act26_msft']['values'][1]}").number_format = "0.000000"
    ws.cell(r, 8, "Wattonomics In!I448 act26_msft＝FY26Q3 31.9＋FY26Q4 41.0＋FY27Q1 45.2；與本表同一實績（只被 Compare 頁讀）"); r += 1
    rp1 = inp("資本支出 CY2026Q1＝FY26Q3（現金 PP&E）", "CY2026Q1", "US$B", f"=Q!G{QR['capex']}", SRC_CASH, GR_CASH, url="https://www.microsoft.com/en-us/investor/earnings/fy-2027-q1/press-release-webcast（FY27Q1 10-Q；Form 10-Q 現金流量表）")
    rp2 = inp("資本支出 CY2026Q2＝FY26Q4（現金 PP&E）", "CY2026Q2", "US$B", f"=Q!H{QR['capex']}", SRC_CASH, GR_CASH, url="")
    rp3 = inp("資本支出 CY2026Q3＝FY27Q1（現金 PP&E）", "CY2026Q3", "US$B", f"=Q!I{QR['capex']}", SRC_CASH, GR_CASH, url="")
    rp9 = inp("資本支出 CY2026 已公布三季累計（現金 PP&E）", "CY2026Q1–Q3（9 個月）", "US$B", f"=D{rp1}+D{rp2}+D{rp3}", SRC_CASH, GR_CASH, url="")
    rl = inp("層級 CY2026 玩家資本支出：AI＋非 AI（Wattonomics）", "CY2026（全年）", "US$B", None, "—", "—",
             f"=C{wrows['capai']}+C{wrows['capnai']}", f"{WV}：Gap 頁 Microsoft 2026 AI 資本支出（新增 {WN['WIF_Gap_Player_msft_2026_AddGW']:.2f} GW × Tokenomics 每 GW 成本，含廠房提前）＋非 AI 資本支出（連 F1 兩列）；"
             "兩者的輸入都是公司資本支出（FY2026 實績與個股模型寫回 FY2027），屬同源（見下方追公式）", "Derived；輸入同源於公司資本支出",
             url=f"wattonomics/model/{WS['file']}；快照 {WS['snapshot']}（repo 內路徑；層級值出處）")
    ws.cell(rl, 4).value = None
    ws.cell(r, 1, "主讀數：已公布期間（CY2026Q1–Q3）累計 vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計實績（公司口徑）；層級 × 9/12（線性時差近似）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq9}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "三季累計＝CY2026Q1＋Q2＋Q3（上列，公司口徑）"); ws.cell(r, 9, GR_CO); ws.cell(r, 10, "層級 CY2026 × 9 ÷ 12（Derived；線性時差近似）"); ws.cell(r, 11, "Derived")
    rd1 = r; r += 1
    ws.cell(r, 1, "第二讀數：最近一季（CY2026Q3＝FY27Q1）× 4 vs 層級 CY2026"); ws.cell(r, 2, "CY2026Q3 × 4（約等於該季中點的年化速率）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq3}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "CY2026Q3 單季 × 4（Derived；公司口徑）"); ws.cell(r, 9, GR_CO); ws.cell(r, 10, "層級 CY2026 全年（同上）"); ws.cell(r, 11, "Derived")
    rd2 = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：資本支出").font = BOLD
    ws.cell(r, 2, "層級值同源於公司資本支出（見追公式）：參考（同源校準），不判定")
    ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由公司資本支出實績與個股模型寫回換算＝同源，不是獨立證據；上方兩個讀數仍列出供參考"); ws.cell(r, 9, GR_CO); ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL},SIGN(F{rd1})=SIGN(F{rd2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rjr = r; r += 1
    ws.cell(r, 1, "口徑參考（不判定）：現金 PP&E 三季累計 vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計（現金 PP&E；不含融資租賃）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rp9}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（口徑差，不判定）"); ws.cell(r, 8, "現金 PP&E 不含融資租賃；層級資本支出為「完工口徑」近似（含融資租賃的設施），所以以公司口徑為主讀數；此列只顯示口徑差（9M 約差融資租賃增添）")
    ws.cell(r, 9, GR_CASH); ws.cell(r, 10, "層級 CY2026 × 9 ÷ 12"); ws.cell(r, 11, "Derived")
    rc9 = r; r += 1
    ws.cell(r, 1, "口徑參考（不判定）：現金 PP&E CY2026Q3 × 4 vs 層級 CY2026"); ws.cell(r, 2, "CY2026Q3 × 4（現金 PP&E）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rp3}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（口徑差，不判定）"); ws.cell(r, 8, "同上（現金 PP&E 單季年化）"); ws.cell(r, 9, GR_CASH); ws.cell(r, 10, "層級 CY2026 全年"); ws.cell(r, 11, "Derived")
    rc3 = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：公司 CY2026 資本支出指引 vs 層級 CY2026（指引為公司口徑，差距來自 GW 換算與廠房提前入帳的時點、以及層級以 FY2026 實績與個股模型 FY2027 推估各半換算）"); ws.cell(r, 2, "指引（公司指引不是實績）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, fb("guid.cy26.capex")).number_format = "#,##0.0"; ws.cell(r, 4).font = BLUE; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（指引，不判定）")
    ws.cell(r, 8, "FY27Q1 法說（2026-10-28）：曆年 2026 資本支出約 170（事實總帳 guid.cy26.capex；Interested-party；逐字稿不在 repo，chat 端提供）"); ws.cell(r, 9, "Interested-party（指引）"); ws.cell(r, 10, "同上"); ws.cell(r, 11, "Derived")
    rg = r; r += 1
    # ---- AI 營收與在役 MW（參考；無可比實績）----
    ws.cell(r, 1, "層級隱含 Azure AI 營收（CY2026；參考）：年均上線 GW × 對外比例 × 每 GW 營收錨（k＝1）").font = BOLD; r += 1
    ra1 = r; ws.cell(r, 1, "年均上線 IT GW（2025、2026 年底平均；Wattonomics）"); ws.cell(r, 3, "GW"); ws.cell(r, 5, f"=(C{wrows['ye25']}+C{wrows['ye26']})/2").number_format = "0.000"; ws.cell(r, 10, "Wattonomics PlayerGW_YE msft 2025、2026（連 F1）；2026 年增量由公司資本支出換算，屬同源（追公式）"); ws.cell(r, 11, "Derived（同源）"); r += 1
    ra2 = r; ws.cell(r, 1, "對外比例（1 − Wattonomics 自用比例；同源：Wattonomics 沿用 Microsoft 舊模型）"); ws.cell(r, 3, "比"); ws.cell(r, 5, f"=1-C{wrows['self']}").number_format = "0.00"; ws.cell(r, 10, "Wattonomics al_msft_self；Assumed"); ws.cell(r, 11, "Assumed"); r += 1
    ra3 = r; ws.cell(r, 1, "每 GW 年營收錨（k＝1；Tokenomics IF_HoldEcon × 本模型期初世代組合）"); ws.cell(r, 3, "US$B／GW")
    ws.cell(r, 5, f"=mix_h*C{F1R['IF_HoldEcon｜H100']}+mix_g2*C{F1R['IF_HoldEcon｜GB200']}+mix_g3*C{F1R['IF_HoldEcon｜GB300']}").number_format = "0.000"
    ws.cell(r, 10, "Tokenomics IF_HoldEcon（連 F1）× 本模型期初世代組合 mix_h／mix_g2／mix_g3（Assumed）"); ws.cell(r, 11, "Derived／Assumed"); r += 1
    ra5 = r; ws.cell(r, 1, "層級隱含 Azure AI 營收（CY2026；US$B）"); ws.cell(r, 3, "US$B"); ws.cell(r, 5, f"=E{ra1}*E{ra2}*E{ra3}").number_format = "#,##0.0"; ws.cell(r, 10, "上列三項相乘（層級的對外比例同源，k 取 1；Maia 調整不計，Wattonomics 無 Microsoft 的 nnv 列）"); ws.cell(r, 11, "Derived"); r += 1
    ws.cell(r, 1, "公司 AI 業務年化營收（FY26Q3 法說 >37；含 Copilot 與自有 AI 產品）vs 層級隱含 Azure AI 營收（CY2026）"); ws.cell(r, 2, "FY26Q3 年化（2026-01～03）；層級 CY2026 全年"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, "=ai_runrate").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{ra5}").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（無可比實績，不判定）")
    ws.cell(r, 8, "Microsoft FY26Q3 法說：AI 業務年化營收超過 37（In 頁 ai_runrate；事實總帳 call.aiRunRate.fy26q3；含 Copilot 等自有 AI 產品，公司不揭露 Azure AI 單獨營收）；F 欄＝層級隱含 Azure AI 營收占公司 AI 年化的比例")
    ws.cell(r, 9, "Interested-party（公司法說）"); ws.cell(r, 10, "層級隱含 Azure AI 營收 CY2026（上列）；口徑不同（層級只含對外 Azure AI，公司值含 Copilot），只作參考"); ws.cell(r, 11, "Derived／Assumed"); ws.cell(r, 12, "https://www.microsoft.com/en-us/investor/earnings/fy-2026-q3/press-release-webcast")
    rai = r; r += 1
    ws.cell(r, 1, "上限檢查（層級隱含 Azure AI 營收不得高於 Azure 分部實績年化）"); ws.cell(r, 7, f'=IF(E{ra5}<=Q!I{QR["azure"]}*4,"通過","不通過：層級隱含 Azure AI 營收高於 Azure 分部實績年化")')
    ws.cell(r, 8, "Azure 分部營收 FY27Q1 × 4（Q 頁第 " + str(QR["azure"]) + " 列；Verified）"); rcu = r; r += 1
    ws.cell(r, 1, "在役／簽約電力 MW 存量（公司實績）"); ws.cell(r, 2, "無可比實績（存量）"); ws.cell(r, 7, "無可比實績（存量）")
    ws.cell(r, 8, "Microsoft 不揭露年底在役或簽約容量存量（MW／GW）；法說只口頭揭露各季新增容量（見下方「公司口頭揭露新增容量」列）與「兩年內整體容量約翻倍」；Epoch AI 晶片功耗與園區估計是第三方估計，不算公司實績；本模型 AI 在役 MW 為假設（Assumed）再以資本支出指引重解速度（C10）")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "Wattonomics 上線 GW：2025、2026 年底見 F1；存量只供參考"); ws.cell(r, 11, "Derived"); rmw = r; r += 1
    # ---- 新增容量 GW：公司法說口頭揭露 vs 層級新增 IT GW（新增 GW 部分同源→改比每 GW 物理量；規則 tools/coord/README.md）----
    ws.cell(r, 1, "公司口頭揭露新增容量 GW（Interested-party，法說）vs 層級新增 IT GW（新增 GW 部分同源：公司資本支出 ÷ Tokenomics 每 GW 成本，所以改比 GW 物理量）").font = BOLD; r += 1
    U_MB = "https://www.marketbeat.com/instant-alerts/microsoft-q3-earnings-call-highlights-2026-04-29/"
    U_NP = "https://www.nextplatform.com/cloud/2026/05/04/microsoft-committed-to-doubling-ai-infrastructure-in-two-years/5219208"
    U_DCD = "https://www.datacenterdynamics.com/en/news/microsoft-brought-88-data-centers-online-in-fy2026/"
    U_Q1 = "https://www.webull.com/news/15828153888932864"
    GR_GW = "Interested-party（公司法說口頭約數；二手轉述）"
    LGW = f"{WV}：WIF_Gap_Player_msft_2026_AddGW（Gap!G38←CapexGW!G34；2026 年上線 IT GW，含上線時差）"

    def gwrow(label, per, val, src, url, unit="GW"):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, per); ws.cell(r, 3, unit)
        if val is not None:
            c = ws.cell(r, 4, val); c.number_format = "0.0"; c.font = BLUE
        ws.cell(r, 8, src); ws.cell(r, 9, GR_GW if val is not None else "—"); ws.cell(r, 12, url)
        r += 1
        return r - 1
    rgw25 = gwrow("參考：CY2025 全年新增容量（公司）vs 層級 2025 新增 IT GW（不在 CY2026 C0 期間）", "CY2025（全年）", 2.0,
                  "FY26Q4 法說（2026-07）轉述：2025 年全年上線約 2 GW；FY26Q2（2025-10～12）約 1 GW（DCD 轉述「前兩季各約 1 GW」）；口徑同下方各季（整體容量，未說明 AI／IT）", U_DCD)
    ws.cell(rgw25, 5, f"=C{wrows['add25']}").number_format = "0.000"
    ws.cell(rgw25, 6, f"=D{rgw25}/E{rgw25}-1").number_format = "+0.0%;-0.0%"; ws.cell(rgw25, 7, "參考（2025 年，不判定）")
    ws.cell(rgw25, 10, f"{WV}：WIF_Gap_Player_msft_2025_AddGW（2025 年新增 IT GW；連 F1 讀取列）"); ws.cell(rgw25, 11, "Derived")
    rgq1 = gwrow("新增容量 CY2026Q1＝FY26Q3（公司口頭揭露）", "CY2026Q1（2026-01～03）", 1.0,
                 "FY26Q3 法說（2026-04-29）Nadella：本季「another gigawatt of capacity」／「added 1 gigawatt of additional capacity in the March quarter」（MarketBeat、NextPlatform 轉述）；未說明 AI 或整體、IT 或電力口徑", U_MB + "；" + U_NP)
    rgq2 = gwrow("新增容量 CY2026Q2＝FY26Q4（公司口頭揭露）", "CY2026Q2（04～06）", 1.0,
                 "FY26Q4 法說（2026-07）Nadella：上一季新增 31 座資料中心（五大洲），「contributing 1GW of capacity」；FY2026 全年上線 88 座；「roughly double our overall capacity in just two years」（DCD 轉述）", U_DCD)
    rgq3 = gwrow("新增容量 CY2026Q3＝FY27Q1（公司口頭揭露）", "CY2026Q3（07～09）：未揭露 GW", None,
                 "FY27Q1 法說（2026-10-28）Nadella：本季新增 24 座資料中心（五大洲），交機到上線時間季減約 20%；未給 GW／MW（本輪查逐字稿轉載與報導：Webull 逐字稿、Windows Forum「did not give locations or megawatt figures」）", U_Q1)
    rgl = r; ws.cell(r, 1, "層級 CY2026 新增 IT GW（Wattonomics）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "GW")
    ws.cell(r, 5, f"=C{wrows['addgw']}").number_format = "0.000"; ws.cell(r, 8, "—"); ws.cell(r, 9, "—")
    ws.cell(r, 10, LGW + "；＝公司資本支出日曆年 × AI 占比 ÷ 每 GW 成本（IT＋廠房）再依上線時差加權；部分同源（公司資本支出），GW 物理量本身不是公司揭露")
    ws.cell(r, 11, "Derived；部分同源"); ws.cell(r, 12, f"wattonomics/model/{WS['file']}；快照 {WS['snapshot']}"); r += 1
    rgd1 = r; ws.cell(r, 1, "主讀數（GW）：CY2026 已揭露季合計（Q1＋Q2，6 個月）vs 層級 CY2026 新增 GW × 6 ÷ 12"); ws.cell(r, 2, "6 個月合計（公司口頭）；層級 × 6/12"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=D{rgq1}+D{rgq2}").number_format = "0.00"; ws.cell(r, 5, f"=E{rgl}*6/12").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "CY2026Q1＋Q2（上列；Q3 公司未揭露 GW，所以月數取 6）"); ws.cell(r, 9, GR_GW); ws.cell(r, 10, "層級 CY2026 新增 IT GW × 6 ÷ 12（線性時差近似）"); ws.cell(r, 11, "Derived"); r += 1
    rgd2 = r; ws.cell(r, 1, "第二讀數（GW）：最近揭露一季（CY2026Q2＝FY26Q4）× 4 vs 層級 CY2026 新增 GW"); ws.cell(r, 2, "CY2026Q2 × 4（該季中點年化）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=D{rgq2}*4").number_format = "0.00"; ws.cell(r, 5, f"=E{rgl}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "CY2026Q2 單季 × 4（Derived；季中點 5 月中，接近年中，時點偏誤小）"); ws.cell(r, 9, GR_GW); ws.cell(r, 10, "層級 CY2026 全年"); ws.cell(r, 11, "Derived"); r += 1
    rfac = r; ws.cell(r, 1, "橋接參數：每 IT GW 對應設施電力（Tokenomics 基準；區間 1.1–1.3）"); ws.cell(r, 3, "GW／IT GW")
    c = ws.cell(r, 4, tk("IF_FacilityGW", "gb300")); c.number_format = "0.00"; c.font = BLUE
    ws.cell(r, 8, "若公司口頭的 GW 是設施（電網）電力而非 IT 關鍵電力，÷ 本值換成 IT GW；公司未說明口徑，此換算為假設"); ws.cell(r, 9, "Assumed（Tokenomics IF_FacilityGW）"); r += 1
    rais = r; ws.cell(r, 1, "橋接參數：AI 占比（Wattonomics 資本支出 AI 占比，借用為容量 AI 占比）"); ws.cell(r, 3, "比")
    c = ws.cell(r, 4, WR["ai_sh"]["values"][1]); c.number_format = "0.00"; c.font = BLUE
    ws.cell(r, 8, "公司容量為整體容量（含 Azure 一般雲端、Microsoft 365 等非 AI 負載）；以 Wattonomics ai_sh 0.85（區間 0.75–0.95，資本支出口徑）近似容量中的 AI 占比；非 AI 機房每 GW 成本較低，容量口徑的 AI 占比可能更低"); ws.cell(r, 9, "Assumed（Wattonomics ai_sh）"); r += 1
    rgb1 = r; ws.cell(r, 1, "口徑橋接後主讀數（參考）：已揭露季合計 ÷ 設施電力係數 × AI 占比 vs 層級 × 6 ÷ 12"); ws.cell(r, 2, "橋接（Assumed）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=D{rgd1}/D{rfac}*D{rais}").number_format = "0.000"; ws.cell(r, 5, f"=E{rgd1}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（橋接參數 Assumed）"); ws.cell(r, 8, "未扣租用的第三方託管與 neocloud 容量（層級只算自有＋融資租賃；租用容量歸出租方），扣除後會更低"); ws.cell(r, 9, "Assumed（橋接）"); ws.cell(r, 10, "同主讀數"); ws.cell(r, 11, "Derived"); r += 1
    rgb2 = r; ws.cell(r, 1, "口徑橋接後第二讀數（參考）：最近揭露一季 × 4 ÷ 設施電力係數 × AI 占比 vs 層級 CY2026"); ws.cell(r, 2, "橋接（Assumed）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=D{rgd2}/D{rfac}*D{rais}").number_format = "0.000"; ws.cell(r, 5, f"=E{rgd2}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（橋接參數 Assumed）"); ws.cell(r, 8, "同上"); ws.cell(r, 9, "Assumed（橋接）"); ws.cell(r, 10, "同第二讀數"); ws.cell(r, 11, "Derived"); r += 1
    rgb25 = r; ws.cell(r, 1, "同參數橋接套到 CY2025 參考列（檢查用）：CY2025 全年 ÷ 設施電力係數 × AI 占比 vs 層級 2025 新增 IT GW"); ws.cell(r, 2, "CY2025；橋接（Assumed）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=D{rgw25}/D{rfac}*D{rais}").number_format = "0.000"; ws.cell(r, 5, f"=E{rgw25}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（橋接參數 Assumed）"); ws.cell(r, 8, "同一組橋接參數套到 CY2025 得到遠低於層級的結果：橋接後 CY2026 落在門檻內是參數使然，不代表差距已解；口徑因素大致固定，解釋不了公司對層級比值由 CY2025 到 CY2026 的跳升"); ws.cell(r, 9, "Assumed（橋接）"); ws.cell(r, 10, "同 CY2025 參考列"); ws.cell(r, 11, "Derived"); r += 1
    rgj = r; ws.cell(r, 1, "C0（實績）判定：新增容量 GW").font = BOLD
    ws.cell(r, 2, "公司口徑無法與層級對齊：參考（無可比實績），不判定"); ws.cell(r, 7, "參考（無可比實績：口徑無法對齊，缺 AI／IT 口徑新增容量揭露）")
    ws.cell(r, 8, F_(T("口徑無法對齊的原因：（1）公司說的是整體容量（兩年內整體容量約翻倍），含非 AI 負載，層級是 AI 的 IT GW；（2）公司未說明是設施電力或 IT 電力，層級是 IT 關鍵電力；"
                       "（3）新增的資料中心可能含向第三方租用的託管機房與 neocloud 容量，層級只算自有＋融資租賃；（4）公司數字是口頭約數（1 GW、another gigawatt），FY27Q1 未揭露，CY2026 只有兩季。"
                       "未換算時兩讀數為 "), f'TEXT(F{rgd1},"+0%;-0%")', T(" 與 "), f'TEXT(F{rgd2},"+0%;-0%")', T("（同號、都超過 ±15%）；以設施電力係數 "), f'TEXT(D{rfac},"0.0")', T(" 與 AI 占比 "), f'TEXT(D{rais},"0.00")',
                    T(" 橋接後為 "), f'TEXT(F{rgb1},"+0%;-0%")', T(" 與 "), f'TEXT(F{rgb2},"+0%;-0%")', T("，但這是 Assumed 參數下的讀數，不得用來說差距已解："),
                    T("（a）同一組參數套到 CY2025 參考列（公司 "), f'TEXT(D{rgw25},"0.0")', T(" GW 對層級 "), f'TEXT(E{rgw25},"0.000")', T(" GW）得 "), f'TEXT(F{rgb25},"+0%;-0%")',
                    T("；未換算時 CY2025 與 CY2026 分別為 "), f'TEXT(F{rgw25},"+0%;-0%")', T(" 與 "), f'TEXT(F{rgd1},"+0%;-0%")', T("，公司對層級比值由 "), f'TEXT(D{rgw25}/E{rgw25},"0.00")', T(" 升到 "), f'TEXT(D{rgd1}/E{rgd1},"0.00")',
                    T("。口徑因素（整體對 AI、設施對 IT、自有對租用）大致固定，解釋不了比值跳升；真正差異在 2026 年增速：層級 2025→2026 新增 GW 只增 "), f'TEXT(E{rgl}/E{rgw25}-1,"0%")',
                    T("，公司口頭約翻倍（最近揭露一季年化對 2025 全年 "), f'TEXT(D{rgd2}/D{rgw25}-1,"+0%;-0%")', T("）；可能原因為租用 neocloud 容量增加、層級每 GW 成本或上線時差。"),
                    T("（b）設施電力係數 1.1–1.3 與 AI 占比 "), f'TEXT({WR["ai_sh"]["values"][0]},"0.00")', T("–"), f'TEXT({WR["ai_sh"]["values"][2]},"0.00")', T(" 一併取區間時主讀數為 "),
                    f'TEXT(D{rgd1}/1.3*{WR["ai_sh"]["values"][0]}/E{rgd1}-1,"+0%;-0%")', T(" 至 "), f'TEXT(D{rgd1}/1.1*{WR["ai_sh"]["values"][2]}/E{rgd1}-1,"+0%;-0%")',
                    T("，且租用容量未扣。結論隨 Assumed 的橋接參數跨過門檻，所以列參考（無可比實績）、不計入對帳中；公司若揭露 AI 或 IT 口徑、自有與租用拆分，再依兩讀數規則判定。")))
    ws.cell(r, 9, GR_GW); ws.cell(r, 10, "Wattonomics 新增 IT GW（部分同源；比的是每 GW 成本、AI 占比與上線時差）"); ws.cell(r, 11, "Derived（部分同源）"); r += 1
    rgjr = r; ws.cell(r, 1, "參考：未換算時兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rgd1})>{TOL},ABS(F{rgd2})>{TOL},SIGN(F{rgd1})=SIGN(F{rgd2})),"參考（口徑無法對齊）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rgd1})>{TOL},ABS(F{rgd2})>{TOL}),"參考（口徑無法對齊）：只一個讀數超過門檻（時差邊界）","參考（口徑無法對齊）：兩讀數都在門檻內"))'); r += 1
    rgpg = r; ws.cell(r, 1, "每 GW 資本支出（參考）：公司口徑資本支出 CY2026Q1＋Q2 ÷ 公司口頭揭露新增 GW vs Wattonomics 每 IT GW 建置成本（CapexGW!G22＋G25；IT＋廠房）"); ws.cell(r, 2, "CY2026 上半年"); ws.cell(r, 3, "US$B／GW")
    ws.cell(r, 4, f"=(D{rq1}+D{rq2})/(D{rgq1}+D{rgq2})").number_format = "0.0"
    ws.cell(r, 5, tcell("CapexGW!G22")["value"] + tcell("CapexGW!G25")["value"]).number_format = "0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（口徑無法對齊）")
    ws.cell(r, 8, "公司側＝公司口徑資本支出（含融資租賃、含非 AI、付款早於上線）÷ 口頭揭露新增容量（整體容量、口徑不明）；兩者期間與口徑都不一致，只看量級"); ws.cell(r, 9, GR_GW)
    ws.cell(r, 10, f"{WV}：CapexGW!G22（每 IT GW 的 IT 資本支出，含 Maia 調整）＋G25（每 IT GW 廠房），追公式傾印值"); ws.cell(r, 11, "Derived"); r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "時差邊界（參考）項目數"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"時差邊界（參考）")'); rnb = r; r += 1
    ws.cell(r, 1, "無可比實績列數（資本支出 1、AI 營收 1、在役 MW 存量 1、新增容量 GW 1〔「參考（無可比實績：口徑無法對齊…）」〕；每 GW 資本支出列「參考（口徑無法對齊）」不在此計）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")+COUNTIF(G{F2a}:G{F2b},"參考（無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F_(T("C0（實績）對帳：公司已公布資本支出（公司口徑，CY2026 年 Q1–Q3 累計 "), f'TEXT(D{rd1},"0.0")', T(" 十億美元）與層級 CY2026 隱含值（"), f'TEXT(E{rl},"0.0")', T("；同期間換算 "), f'TEXT(E{rd1},"0.0")', T("）相差 "),
                     f'TEXT(F{rd1},"+0%;-0%")', T("（主讀數）；最近一季年化 "), f'TEXT(D{rd2},"0.0")', T(" 十億美元，相差 "), f'TEXT(F{rd2},"+0%;-0%")',
                     T("（第二讀數）。但 Wattonomics 的 Microsoft 2026 年資本支出與新增 GW 由公司資本支出（FY2026 實績與個股模型寫回的 FY2027）換算（追公式：CapexGW!G17、H17←In!K169、L169），屬同源，判定："), f'G{rj}',
                     T("（參考，不計入對帳中）。新增容量：公司法說口頭揭露 CY2026 上半年約 "), f'TEXT(D{rgd1},"0.0")', T(" GW，比層級同期 "), f'TEXT(E{rgd1},"0.00")', T(" GW "), f'IF(F{rgd1}<0,"低 ","高 ")', f'TEXT(ABS(F{rgd1}),"0%")', T("（最近揭露一季年化 "), f'TEXT(F{rgd2},"+0%;-0%")', T("），但公司為整體容量、未說明 AI 或 IT 口徑、可能含租用，以假設參數橋接後 "), f'TEXT(F{rgb1},"+0%;-0%")',
                     T("（Assumed，不代表差距已解：同參數套到 CY2025 得 "), f'TEXT(F{rgb25},"+0%;-0%")', T("；未換算時公司對層級比值由 CY2025 的 "), f'TEXT(D{rgw25}/E{rgw25},"0.00")', T(" 升到 CY2026 的 "), f'TEXT(D{rgd1}/E{rgd1},"0.00")',
                     T("，口徑因素大致固定解釋不了，真正差異在 2026 增速——層級新增 GW 只增 "), f'TEXT(E{rgl}/E{rgw25}-1,"0%")', T("、公司口頭約翻倍；設施電力係數與 AI 占比一併取區間時主讀數為 "),
                     f'TEXT(D{rgd1}/1.3*{WR["ai_sh"]["values"][0]}/E{rgd1}-1,"+0%;-0%")', T(" 至 "), f'TEXT(D{rgd1}/1.1*{WR["ai_sh"]["values"][2]}/E{rgd1}-1,"+0%;-0%")', T("），結論隨橋接參數跨過門檻，判定："), f'G{rgj}', T("。AI 營收（公司只揭露含 Copilot 的 AI 年化）與在役容量存量沒有可比的公司實績，不判定。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向（依季中點）")
    ws.cell(r, 2, F_(T("資本支出是流量，逐季上升（Q1 "), f'TEXT(D{rq1},"0.0")', T("、Q2 "), f'TEXT(D{rq2},"0.0")', T("、Q3 "), f'TEXT(D{rq3},"0.0")',
                      T("）。主讀數用 9 個月平均速率（中點約 5 月中）代表全年，全年中點在 7 月初，所以在上升趨勢下主讀數偏低；第二讀數用 Q3 單季年化（中點約 8 月中），比全年平均速率的中點晚一個半月，所以偏高。"
                        "兩讀數相差 "), f'TEXT((F{rd2}-F{rd1})*100,"0")', T(" 個百分點，就是時點差的幅度。公司曆年 2026 全年實績要等 FY27Q2 財報（2027-01）才完整；層級 CY2026 與公司曆年指引相差 "), f'TEXT(F{rg},"+0%;-0%")',
                      T("，來自 GW 換算與廠房提前入帳的時點，以及層級以財年換算日曆年。這裡全部用流量（資本支出）比較，不拿 9 月底存量對年底存量。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出列：追公式（F1b）顯示 Wattonomics 2026 年 Microsoft AI 資本支出（Interface!F596 ← Gap!G62 ← G45＝G42＋G43＋G44）與非 AI 資本支出（Gap!G52）的輸入是公司資本支出（CapexGW!G17←In!K169 FY2026 實績 145.3；H17←In!L169 個股模型寫回 FY2027 204.2；非 AI 讀 In!K168、L168 個股模型寫回）；"
                   "新增 GW 也由此換算（CapexGW!G34）。層級值因此同源於公司資本支出與個股模型，不是獨立證據：整列標「參考（同源校準）」，結論「無可比實績（層級值同源）」。公司已公布實績 act26_msft 沒有進 Microsoft 的 GW 或資本支出路徑（只在 Compare 頁作權重）。"
                   "（二）AI 營收列：用到上述 GW（同源）、對外比例 60%（沿用 Microsoft 舊模型，同源）、期初世代組合（本模型假設）；公司只揭露含 Copilot 的 AI 年化營收（>37），不揭露 Azure AI 單獨營收，所以只列層級隱含值占公司 AI 年化的比例與上限檢查，不判定。"
                   "（三）新增容量 GW 列：層級新增 GW＝公司資本支出 × AI 占比 ÷ Tokenomics 每 GW 成本再加上線時差，屬部分同源；拿公司法說口頭揭露的新增 GW 比，公司資本支出不再是比較對象，實際檢驗的是每 GW 成本、AI 占比與上線時差（每 GW 物理檢驗）。但公司 GW 為整體容量、口徑（電力或 IT、自有或租用）未說明，橋接參數為假設且結論跨過門檻（同參數套到 CY2025 得到相反方向的大差距），所以列「參考（無可比實績：口徑無法對齊，缺 AI／IT 口徑新增容量揭露）」。在役容量存量公司不揭露，列「無可比實績（存量）」。""（四）所以本頁 C0 在 Microsoft 對帳中為 0：資本支出列同源、新增 GW 列口徑無法對齊，不代表兩層一致。若公司揭露 AI 或 IT 口徑的新增容量，或 Wattonomics 改以 Epoch／園區證據獨立推 2026 年 GW，再依兩讀數規則判定。"))
    rsame = r; r += 1
    ws.cell(r, 1, "口徑換算（層級 vs 公司資本支出）")
    ws.cell(r, 2, ("公司口徑資本支出＝現金購置不動產與設備＋融資租賃增添（公司 IR 財務報表 CapEx 工作表）；層級資本支出＝新增 GW × 每 GW 建置成本（IT＋廠房）＋非 AI 資本支出，屬完工口徑近似。"
                   "Microsoft 的融資租賃占比大（FY27Q1 單季公司口徑 45.2 對現金 PP&E 39.4，差額為融資租賃等非現金取得），所以主讀數用含融資租賃的公司口徑，現金 PP&E 只列口徑參考；"
                   "Wattonomics 對 Microsoft 為擁有者加融資租賃口徑（In 頁 cap_msft 標「自有＋融資租賃 AI 在役」），租用 neocloud 的容量屬出租方，不重複計。"))
    rk = r; r += 1
    ws.cell(r, 1, "證據回饋註記（待對帳單；不改上游值）"); ws.cell(r, 2, F_(T("2025→2026 公司口頭容量增速遠高於層級（公司最近揭露一季年化對 2025 全年 "), f'TEXT(D{rgd2}/D{rgw25}-1,"+0%;-0%")', T("，層級新增 IT GW "), f'TEXT(E{rgl}/E{rgw25}-1,"+0%;-0%")', T("），待公司揭露口徑（AI／IT、自有與租用）；揭露後若差距仍在，走對帳單（C5）檢查層級每 GW 成本、AI 占比與上線時差。")))
    ws.cell(r, 7, "參考（證據回饋，待公司揭露口徑）"); rfb = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 輸入的 CY2026 資本支出（CapexGW!G18，公司資本支出換成日曆年）vs 其 Gap 頁 GW 推算的玩家資本支出（AI＋非 AI；輸出相對輸入）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['gcy']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics CapexGW!G18（公司資本支出日曆年換算；輸入）"); ws.cell(r, 10, "Wattonomics Gap 頁玩家資本支出（GW 推算；輸出）"); rb_ = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：CY2026 資本支出指引（上列）、RPO（In 頁事實）等不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：Tokenomics 本模型讀 " + TKV + "（現行版；本模型所讀名稱數值與 v21.10 相同）；Wattonomics v0.3.2 內含的 Tokenomics 為 v21.10，與本模型的 " + TKV + " 不是同一版（層級值之間的版本差異只列參考）。Microsoft 模型值曾被 Wattonomics 讀入（資本支出寫回，In!K169、L169、K168、L168），本表只列層級值與追公式，不用該回饋判定。")
    r += 1
    return dict(F1=(F1a, F1b), F1chk=F1chk, F2h=F2h, F2hh=F2hh, F2=(F2a, F2b), rq=(rq1, rq2, rq3, rq9), rp=(rp1, rp2, rp3, rp9), rl=rl, rd1=rd1, rd2=rd2, rj=rj, rjr=rjr, rtr=rtr_conc, rdump=(rdump0, rdump1), rdump_note=rdump_note,
                rg=rg, rgw=dict(y25=rgw25, q=(rgq1, rgq2, rgq3), l=rgl, d1=rgd1, d2=rgd2, fac=rfac, ais=rais, b1=rgb1, b2=rgb2, b25=rgb25, j=rgj, jr=rgjr, pg=rgpg), rc9=rc9, rc3=rc3, ra=(ra1, ra2, ra3, ra5), rai=rai, rcu=rcu, rmw=rmw, rn=rn, rnb=rnb, rnc=rnc, rs=rs, rr=rr_, rsame=rsame, rk=rk, rfb=rfb, rb=rb_, wrows=wrows, tol=TOL)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, RVA, PLA, CRD):
    RD = wb["README"]
    VR, BR, BC, QR = C["VR"], C["BR"], C["BC"], C["QR"]
    vb = lambda k, c="B": f"Val!{c}{VR[k]}"
    tp, up, rec, dcf, cmpv, cons = vb("tp"), vb("up"), vb("rec"), vb("dcftp"), vb("comp"), vb("cons")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    ix = PLA["idx"]
    PLB = lambda k, c: f"PL!{c}{BR[k][1]}"
    TH = lambda k, c: f"Thesis!{c}{BR[k][1]}"
    CFB = lambda k, c: f"CF!{c}{BR[k][1]}"
    out = {}
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.8 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F_(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("：現金流折現 "), f'TEXT({dcf},"0")',
        T(" 與分部加總 "), f'TEXT({cmpv},"0")', T(" 各占一半。AI 折舊把營業利益率從 FY2026 的 "), f'TEXT({PLB("oim", "E")},"0%")', T(" 壓到 FY2028 的 "), f'TEXT({PLB("oim", "G")},"0%")', T("。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：自用 AI 採價值中性（讀法 2）；改為只計成本（讀法 1）時加權目標價 "), f'TEXT({vb("tp_r1")},"0")', T(" 美元（"), f'TEXT({vb("up_r1")},"+0%;-0%")', T("），評等"),
        vb("rec_r1"), T("；FY2027 資本支出依公司指引約 "), f'TEXT({CFB("capex", "F")},"0")', T(" 十億美元，Azure 的 AI 營收照該速度長出來；現價 "), 'TEXT(px,"0.00")', T(" 美元為 2026-11-05 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價 "), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；市場共識 "), 'TEXT(cons_tp,"0")', T("（財報前）。三個情境（"), f'TEXT({vb("tps")},"0")', T("／"), f'TEXT({vb("tps", "C")},"0")', T("／"), f'TEXT({vb("tps", "D")},"0")', T("）都低於現價。"))
    put("o2", T("現金流折現偏貴：反向推算（情境 B），要支持現價需把折現率降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), f'TEXT({rv("w0")},"0.0%")', T("），或 FY2028 起營收與營業利益全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍；分部加總隱含 "), f'TEXT(RevDCF!B{RVA["smul"]},"0.0")', T(" 倍 EBITDA，模型 "), f'TEXT(RevDCF!C{RVA["smul"]},"0.0")', T(" 倍。"))
    put("o3", T("AI 投資賺不賺得回資金成本：FY2029 對外 Azure AI 報酬率 "), f'TEXT({TH("roic_x", "H")},"0.0%")', T("，高於 WACC "), 'TEXT(wacc,"0.0%")', T("；加上自用算力（讀法 2）全部新增 AI 資產 "),
        f'TEXT({TH("roic", "H")},"0.0%")', T("，自用只計成本（讀法 1）只有 "), f'TEXT({TH("roic_r1", "H")},"0.0%")', T("。"))
    put("o4", T("真正壓力在 AI 折舊與現金：基準情境新增 AI 折舊 FY2028 "), f'TEXT({PLB("p_aida", "G")},"0")', T(" 十億美元，毛利率由 FY2026 "), f'TEXT(PL!E{ix["gm"]},"0.0%")', T(" 降到 FY2028 "),
        f'TEXT(PL!G{ix["gm"]},"0.0%")', T("；五年自由現金流扣股利合計 "), 'TEXT(CF!C5,"0.0")', T(" 十億美元（基準），"), 'CF!C6', T("。"))
    put("o5", T("實績對帳（本版新增）：CY2026 前三季（FY26Q3 至 FY27Q1）公司口徑資本支出 "), f'TEXT(Coord!D{CRD["rd1"]},"0.0")', T(" 十億美元，與 Wattonomics 隱含的同期間值相差 "), f'TEXT(Coord!F{CRD["rd1"]},"+0%;-0%")',
        T("（最近一季年化 "), f'TEXT(Coord!F{CRD["rd2"]},"+0%;-0%")', T("），但層級值由公司資本支出換算（同源），列參考。公司法說口頭揭露上半年新增約 "), f'TEXT(Coord!D{CRD["rgw"]["d1"]},"0")', T(" GW，比層級同期"), f'IF(Coord!F{CRD["rgw"]["d1"]}<0,"低 ","高 ")', f'TEXT(ABS(Coord!F{CRD["rgw"]["d1"]}),"0%")', T("，但公司說的是整體容量，口徑無法對齊，只列參考；以假設參數橋接後 "), f'TEXT(Coord!F{CRD["rgw"]["b1"]},"+0%;-0%")',
        T("（設施電力係數與 AI 占比一併取區間為 "), f'TEXT(Coord!D{CRD["rgw"]["d1"]}/1.3*0.75/Coord!E{CRD["rgw"]["d1"]}-1,"+0%;-0%")', T(" 至 "), f'TEXT(Coord!D{CRD["rgw"]["d1"]}/1.1*0.95/Coord!E{CRD["rgw"]["d1"]}-1,"+0%;-0%")',
        T("）屬 Assumed，不代表差距已解：同參數套到 CY2025 得 "), f'TEXT(Coord!F{CRD["rgw"]["b25"]},"+0%;-0%")', T("，公司對層級比值由 CY2025 的 "), f'TEXT(Coord!D{CRD["rgw"]["y25"]}/Coord!E{CRD["rgw"]["y25"]},"0.00")',
        T(" 升到 CY2026 的 "), f'TEXT(Coord!D{CRD["rgw"]["d1"]}/Coord!E{CRD["rgw"]["d1"]},"0.00")', T("，真正差異在 2026 年增速（層級新增 GW 只增 "), f'TEXT(Coord!E{CRD["rgw"]["l"]}/Coord!E{CRD["rgw"]["y25"]}-1,"0%")', T("，公司口頭約翻倍）。"))
    put("drv", T("分部營收推導（情境 B）：Agents and Infra＝Azure（Azure AI〔在役 AI MW × 每 MW 持有成本 × k〕＋Azure 其他〔近四季年增向長期值收斂〕）＋Microsoft 365 cloud（基礎＋Copilot〔座席 × 月費〕）＋授權、Industry solutions、Frontier；"
                 "Devices and Consumer＝Search and advertising、XBOX、Windows OEM；兩分部加總為合併營收。FY2028 Agents and Infra "),
        f'TEXT({PLB("p_ai_seg", "G")},"0")', T("（其中 Azure AI "), f'TEXT({PLB("p_s_azai", "G")},"0")', T("、Copilot "), f'TEXT({PLB("p_cop", "G")},"0")', T("）、Devices and Consumer "), f'TEXT({PLB("p_dc", "G")},"0")',
        T("，合計 "), f'TEXT({PLB("p_tot", "G")},"0")', T(" 十億美元。"))
    # v0.3.11（2026-11-24；外部稽核優化第三波，會改值）：W3-7 同業換價分離淨負債、W3-5 Tokenomics 快照 v22.0；v0.3.10 列凍結為固定文字
    RD.cell(r, 1, F_(T("v0.3.11（2026-11-24）：外部稽核優化第三波（會改值；現價不變，2026-11-20 收盤 "), 'TEXT(px,"0.00")',
                     T("）。(1) 同業 EV／EBITDA 的換價改為分離淨負債（W3-7）：Microsoft 自身與企業軟體同業的 ORCL、CRM，企業價值只有市值部分隨股價變、淨負債不變；其餘同業 repo 內沒有市值，仍為倍數 × 收盤比的整體縮放（近似）。各組中位數（11-13 → 11-20）：" + _PXTXT
                       + "。(2) Tokenomics 快照 v21.10 → v22.0（W3-5）：本模型讀的名稱數值相同，只有版本與出處文字改變。加權目標價 354.14 → "), f'TEXT({vb("tp")},"0.00")',
                     T("；上檔空間 −26.9% → "), f'TEXT({vb("up")},"+0.0%;-0.0%")', T("；投資建議 放空 → "), f'{vb("rec")}',
                     T("。上檔空間距放空與中立的分界（−25%）約 1 個百分點。共識、rf、β、ERP 與其他假設不變。")))
    out["ver"] = r; r += 1
    RD.cell(r, 1, "v0.3.10（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）。現價 523.88（11-13）→ 484.52；加權目標價 382.78 → 354.14；上檔空間 −26.9% → -26.9%；投資建議 放空 → 放空。同業倍數（11-13 值）依價格檔 11-13 與 11-20 兩日收盤比換到 2026-11-20（倍數 × 收盤比，近似，無市值；Microsoft 自身用市值法）：非 AI 雲端中位數 14.66 → 12.79、數位廣告 8.54 → 7.35、企業軟體 10.48 → 8.69、遊戲 13.86 → 13.73、PC 7.89 → 7.32、Microsoft 自身 15.5 → 14.35；本模型 WACC 權益權重與回購價隨現價自動重算，共識、rf、β、ERP 與其他假設不變；未換價同業：無。")
    r += 1
    RD.cell(r, 1, "v0.3.9（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 535.93（11-05）→ 523.88；加權目標價 384.07 → 382.78；上檔空間 −28.3% → −26.9%；投資建議 放空 → 放空。同業倍數由 2026-10-07／10-08 報價日換到 11-13（倍數 × 收盤比，近似）")
    r += 1
    RD.cell(r, 1, "v0.3.8（2026-11-15）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁（反向 DCF）、PL 頁末標準損益表"
                  "（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B，GAAP；另列 FY27Q1A）、Sens 頁末加權目標價相對基準（龍捲風圖用）、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」（F1 讀取層級；F1b 追公式；F2 C0（實績）：公司口徑資本支出 vs 層級隱含值）。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    r += 1
    return out


_PX = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "peers_px1120_20261120.json"), encoding="utf-8"))
_PXC = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "peers_cloud_20261120.json"), encoding="utf-8"))
_PXA = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "peers_ads_20261120.json"), encoding="utf-8"))
_PXTXT = (f"非 AI 雲端中位數 {_PXC['median1113']['value']} → {_PXC['median']['value']}、數位廣告 {_PXA['median1113']['value']} → {_PXA['median']['value']}、企業軟體 {_PX['mediansOld']['sw']} → {_PX['medians']['sw']}、"
          f"遊戲 {_PX['mediansOld']['game']} → {_PX['medians']['game']}、PC {_PX['mediansOld']['pc']} → {_PX['medians']['pc']}、Microsoft 自身 {_PX['items']['val.msft.ownNtm']['old']} → {_PX['items']['val.msft.ownNtm']['new']}")


def apply(wb, C):
    PLA = pl_standard(wb, C)
    RVA = revdcf(wb, C)
    tor = sens_tor(wb, C)
    CRD = coord(wb, C)
    RMA = readme13(wb, C, RVA, PLA, CRD)
    return dict(PL=PLA, RV=RVA, tor=tor, CRD=CRD, RM=RMA)
