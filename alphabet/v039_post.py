# Alphabet v0.3.9（2026-11-14）後處理：個股 HTML 版面規格 v1.3 全套＋Coord 頁「三層協調」節（MAG 類樣板）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點；樣板 NVIDIA v0.3.10（nvidia/v037_post.py、coord_post.py）。
# 只新增，不改任何既有儲存格（既有數值、目標價、評等不變；逐格比對只允許版號文字與新增格）：
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；GAAP；另列 9M26A 比較欄）、比率區、核對列、FY2023 實績輸入、分部營收（圖表用）
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑，封閉解與二分法）與分部加總隱含倍數；只讀其他頁，無任何頁引用本頁
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用，白話標籤）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v21.10 快照；Wattonomics v0.3.2 只作對帳列、不進模型）、F2 C0（實績）：10-Q／財報資本支出與 Cloud 營收 vs 層級隱含值
#   README 頁末：v1.3 展示層白話列（橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字、分部營收推導、版本沿革）
# build_alphabet.py 在 README 之後、存檔前呼叫 apply(wb, ctx)。
import json
import os
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.3.9", "2026-11-14"
U_10Q = "https://s206.q4cdn.com/479360582/files/doc_financials/2026/q3/GOOG-10-Q-Q3-2026.pdf"   # 出處網址只放 Coord 頁 L 欄（HTML 不顯示外部網址）
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
SUBF = PatternFill("solid", fgColor="E2EFDA")


def lit(t):
    """Excel 公式字串常數上限 255 字元：切段以 & 相接。"""
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def T(s):
    return ("t", s)


def lab(ws, prefix, start=1, col=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix):
            return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# =====================================================================================================
# PL：標準損益表（版面規格 v1.3 D 節）
# 欄位：C FY2023A、D FY2024A、E FY2025A、F–L FY2026E–FY2032E（連續年度，供圖表）、M 9M26A（最新累計比較欄）、N 說明
# 對應 PL 頁（情境 B 區塊）欄：D→C、E→D、F→G、G→H、…、L→M；M(9M26A)→E
SCOL = "CDEFGHIJKL"          # 年度欄（FY2023A–FY2032E）
S2P = {"D": "C", "E": "D", "F": "G", "G": "H", "H": "I", "I": "J", "J": "K", "K": "L", "L": "M", "M": "E"}
S2Q = {"D": "J", "E": "K", "M": "L"}     # Q 頁：FY2024、FY2025、9M26


def pl_standard(wb, C):
    P, Q, B = wb["PL"], wb["Q"], C["PR"]["B"]
    QR, Fct = C["QR"], C["F"]
    fv = lambda i: Fct[i]["value"] / 1000.0
    r = P.max_row + 3
    L = {}
    P.cell(r, 1, "標準損益表（v0.3.9 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2032E（情境 B 基準）；最右欄 9M26A 為最新累計實績。US$B；EPS 美元。只引用本頁上方、Q 頁與下方 FY2023 實績輸入，不回饋任何計算").font = BOLD
    L["start"] = r; r += 1
    P.cell(r, 1, F(T("口徑：GAAP。Alphabet 以 GAAP 揭露損益，市場共識 EPS 也以 GAAP 稀釋為主；本表其他收入（淨額）含權益證券未實現損益（FY2026E 含 9M26 實績 "),
                   f'TEXT(PL!E{B["eqg"]},"0.0")', T(" US$B，預測期不估），所以 FY2026E EPS 遠高於不含權益證券損益的 EPS（見 PL 頁 EPS 不含權益證券損益列）。")))
    L["note"] = r; r += 1
    P.cell(r, 1, "營業成本與營業費用：模型以分部 EBITDA 與折舊推導營業利益，沒有分列營業成本與營業費用。預測期營業成本＝營收 × 非折舊營業成本率（9M26A 營業成本扣折舊後占營收，下方假設列）＋折舊合計（技術基礎設施折舊記入營業成本，假設新增折舊全數記入）；"
                 "營業費用＝毛利 − 營業利益（殘差，含研發、行銷、管理）。營業利益、稅前、稅後與 EPS 全取自 PL 頁情境 B，不因此改變；歷史年度與 9M26A 為財報實績。")
    L["note2"] = r; r += 1
    hdr = ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "9M26A（累計）", "說明"]
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
            P.cell(r, 14, note)
        L[key] = r; r += 1

    rv, oi, da, oie, pt, tx, ni, she, eps = (B[k] for k in ("rev", "oi", "da", "oie", "pt", "tx", "ni", "she", "eps"))
    pr = lambda row_, c: f"=PL!{S2P[c]}{row_}"
    cell_pl = lambda row_, hist23: {**({"C": hist23} if hist23 else {}), **{c: pr(row_, c) for c in "DEFGHIJKLM"}}
    # 假設列（放在表前的位置不能動；改放表後，名稱供公式引用）
    r_ass = None
    # --- 先預留表列，假設與 FY2023 輸入放表後 ---
    rows_order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    # 列號預先配置（公式互相引用）
    base = r
    idx = {k: base + i for i, k in enumerate(rows_order)}
    # 比率區與核對區在其後
    r_ratio = base + len(rows_order)
    ratio_keys = ["r_yoy", "r_gm", "r_opm", "r_nm", "r_niy"]
    ridx = {k: r_ratio + 1 + i for i, k in enumerate(ratio_keys)}
    r_chk = r_ratio + 1 + len(ratio_keys)
    chk_keys = ["c1", "c2", "c3", "c4", "c5"]
    cidx = {k: r_chk + 1 + i for i, k in enumerate(chk_keys)}
    r_seg = r_chk + 1 + len(chk_keys) + 1                 # 分部營收區
    seg_keys = ["s_gs", "s_cl", "s_ob", "s_tot", "s_chk"]
    sidx = {k: r_seg + 1 + i for i, k in enumerate(seg_keys)}
    r_h23 = r_seg + 1 + len(seg_keys) + 1                 # FY2023 輸入與假設
    h23_keys = ["h_rev", "h_cogs", "h_opx", "h_oi", "h_oie", "h_pt", "h_tax", "h_ni", "h_eps", "h_sh", "h_gs", "h_cl", "h_ob", "h_gp", "h_a_cex", "h_9m25rev", "h_9m25ni"]
    hidx = {k: r_h23 + 2 + i for i, k in enumerate(h23_keys)}
    rcx = hidx["h_a_cex"]

    # 營收
    row("rev", "營收", "US$B", {"C": f"=C{hidx['h_rev']}", **{c: pr(rv, c) for c in "DEFGHIJKLM"}}, M, "分部營收加總（PL 頁情境 B 區塊）；FY2023A 為 10-K 實績")
    assert L["rev"] == idx["rev"]
    yo = {c: f"={c}{idx['rev']}/{p}{idx['rev']}-1" for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    yo["M"] = f"=M{idx['rev']}/C{hidx["h_9m25rev"]}-1"
    row("yoy", "營收年增率（YoY）", "%", yo, PCT, "FY2023A 的基期（FY2022）不在事實總帳，故空白；9M26A 對 2025 年前三季合計（Q 頁）")
    cg = {"C": f"=C{hidx['h_cogs']}", "D": f"=Q!J{QR['cogs']}", "E": f"=Q!K{QR['cogs']}", "M": f"=Q!L{QR['cogs']}",
          "F": f"=M{r}+PL!F{rv}*cogs_ex_r+PL!F{da}"}
    for c in "GHIJKL":
        cg[c] = f"={c}{idx['rev']}*cogs_ex_r+PL!{S2P[c]}{da}"
    row("cogs", "營業成本（COGS）", "US$B", cg, M, "歷史與 9M26A＝財報；FY2026E＝9M26A＋Q4 26E（營收 × 非折舊成本率＋折舊）；FY2027E 起＝營收 × 非折舊成本率＋折舊合計（假設列）")
    row("gp", "毛利", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['cogs']}" for c in "CDEFGHIJKLM"}, M)
    row("gm", "毛利率", "%", {c: f"={c}{idx['gp']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT, "預測期毛利率下降＝折舊（AI 資本支出）上升")
    ox = {"C": f"=C{hidx['h_opx']}", "D": f"=Q!J{QR['rd']}+Q!J{QR['sm']}+Q!J{QR['ga']}", "E": f"=Q!K{QR['rd']}+Q!K{QR['sm']}+Q!K{QR['ga']}",
          "M": f"=Q!L{QR['rd']}+Q!L{QR['sm']}+Q!L{QR['ga']}"}
    for c in "FGHIJKL":
        ox[c] = f"={c}{idx['gp']}-{c}{idx['oi']}"
    row("opx", "營業費用（研發＋行銷＋管理）", "US$B", ox, M, "歷史與 9M26A＝財報；預測期＝毛利 − 營業利益（殘差）")
    oc = {"C": f"=C{hidx['h_oi']}", **{c: pr(oi, c) for c in "DEFGHIJKLM"}}
    row("oi", "營業利益", "US$B", oc, M, "財報口徑合併營業利益（PL 頁；不含自用 AI 影子效益）")
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("oth", "其他收入（淨額）", "US$B", {"C": f"=C{hidx['h_oie']}", **{c: pr(oie, c) for c in "DEFGHIJKLM"}}, M, "利息收入淨額＋權益證券損益＋其他；預測期權益證券損益設 0")
    row("pti", "稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}" for c in "CDEFGHIJKLM"}, M)
    row("tax", "所得稅", "US$B", {"C": f"=C{hidx['h_tax']}", **{c: pr(tx, c) for c in "DEFGHIJKLM"}}, M)
    row("ni", "稅後淨利", "US$B", {c: f"={c}{idx['pti']}-{c}{idx['tax']}" for c in "CDEFGHIJKLM"}, M, "稅前 − 所得稅（核對列：與 PL 頁淨利相同）")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("sh", "稀釋股數", "十億股", {"C": f"=C{hidx['h_sh']}", **{c: (f"={c}{idx['ni']}/{c}{idx['eps']}" if c in "DEF" else pr(she, c)) for c in "DEFGHIJKLM"}}, M, "FY2023A＝10-K 稀釋加權平均；FY2024A、FY2025A、FY2026E＝淨利 ÷ EPS（PL 頁 EPS 列）；FY2027E 起含強制轉換特別股轉換後股數（PL 頁）")
    row("eps", "EPS（GAAP 完全稀釋）", "US$", {"C": f"=C{hidx['h_eps']}", **{c: pr(eps, c) for c in "DEFGHIJKLM"}}, "0.00", "FY2023A＝10-K 公告值；其餘取 PL 頁（淨利 ÷ 稀釋股數；FY2024A、FY2025A 與公告值差 0.01 內）")
    assert L["eps"] == idx["eps"] and r == r_ratio
    P.cell(r, 1, "比率").font = BOLD; r += 1
    row("r_yoy", "營收年增率", "%", {c: f"={c}{idx['yoy']}" for c in "DEFGHIJKLM"}, PCT)
    row("r_gm", "毛利率", "%", {c: f"={c}{idx['gm']}" for c in "CDEFGHIJKLM"}, PCT)
    row("r_opm", "營業利益率", "%", {c: f"={c}{idx['opm']}" for c in "CDEFGHIJKLM"}, PCT)
    row("r_nm", "淨利率", "%", {c: f"={c}{idx['nm']}" for c in "CDEFGHIJKLM"}, PCT)
    ny = {c: f"={c}{idx['ni']}/{p}{idx['ni']}-1" for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    ny["M"] = f"=M{idx['ni']}/C{hidx["h_9m25ni"]}-1"
    row("r_niy", "淨利年增率（YoY）", "%", ny, PCT, "FY2026E 含權益證券損益 127.3（9M26A），預測期不估，所以 FY2027E 年增率為負")
    assert r == r_chk
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    row("c1", "核對：毛利 − 營業費用 − 營業利益", "US$B", {c: f"={c}{idx['gp']}-{c}{idx['opx']}-{c}{idx['oi']}" for c in "CDEFGHIJKLM"}, "0.000", "歷史與 9M26A：財報三項（營業成本、費用、營業利益）互相核對；預測期以殘差定義，恆為 0")
    row("c2", "核對：稅前利益 − PL 頁稅前利益", "US$B", {c: f"={c}{idx['pti']}-PL!{S2P[c]}{pt}" for c in "DEFGHIJKLM"}, "0.000")
    row("c3", "核對：稅後淨利 − PL 頁淨利", "US$B", {c: f"={c}{idx['ni']}-PL!{S2P[c]}{ni}" for c in "DEFGHIJKLM"}, "0.000")
    row("c4", "核對：營收 − PL 頁合併營收", "US$B", {c: f"={c}{idx['rev']}-PL!{S2P[c]}{rv}" for c in "DEFGHIJKLM"}, "0.000")
    row("c5", "核對：FY2023A 稅後淨利 − 10-K 淨利（輸入列）", "US$B", {"C": f"=C{idx['ni']}-C{hidx['h_ni']}"}, "0.000")
    assert r == r_seg - 1
    r += 1
    P.cell(r, 1, "分部營收（US$B；圖表與驅動因子章用）").font = BOLD; r += 1
    segcells = lambda k23, pk: {"C": f"=C{hidx[k23]}", **{c: pr(B[pk], c) for c in "DEFGHIJKLM"}}
    row("s_gs", "Google Services", "US$B", segcells("h_gs", "gs"), M)
    row("s_cl", "Google Cloud", "US$B", segcells("h_cl", "cloud"), M)
    obc = {"C": f"=C{hidx['h_ob']}", **{c: f"=PL!{S2P[c]}{B['ob']}+PL!{S2P[c]}{B['hedge']}" for c in "DEFGHIJKLM"}}
    row("s_ob", "Other Bets＋避險損益", "US$B", obc, M)
    row("s_tot", "合計", "US$B", {c: f"=SUM({c}{sidx['s_gs']}:{c}{sidx['s_ob']})" for c in "CDEFGHIJKLM"}, M)
    row("s_chk", "核對：合計 − 營收（應為 0）", "US$B", {c: f"={c}{sidx['s_tot']}-{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, "0.000")
    assert r == r_h23 - 1
    r += 1
    P.cell(r, 1, "FY2023 實績輸入與假設（藍字＝輸入；FY2023 為 10-K FY2025 的比較欄，取自事實總帳 data/alphabet_facts_20261008.json 的 is.*.fy2023，標 Verified）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "數值", "", "", "", "", "", "", "", "", "", "", "來源"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    src23 = "Alphabet Form 10-K FY2025（FY2023 比較欄；事實總帳 data/alphabet_facts_20261008.json，Verified）"
    hv = [("h_rev", "FY2023 合併營收", fv("is.rev.total.fy2023")), ("h_cogs", "FY2023 營業成本", fv("is.cogs.fy2023")),
          ("h_opx", "FY2023 營業費用（研發＋行銷＋管理）", fv("is.rd.fy2023") + fv("is.sm.fy2023") + fv("is.ga.fy2023")),
          ("h_oi", "FY2023 營業利益", fv("is.ebit.total.fy2023")), ("h_oie", "FY2023 其他收入（淨額）", fv("is.oie.fy2023")),
          ("h_pt", "FY2023 稅前利益", fv("is.pretax.fy2023")), ("h_tax", "FY2023 所得稅", fv("is.tax.fy2023")),
          ("h_ni", "FY2023 淨利", fv("is.ni.fy2023")), ("h_eps", "FY2023 EPS（GAAP 稀釋，美元）", Fct["is.epsDil.fy2023"]["value"]),
          ("h_sh", "FY2023 稀釋加權股數（十億股）", Fct["sh.dilW.fy2023"]["value"] / 1000.0),
          ("h_gs", "FY2023 Google Services 營收", fv("is.rev.gs.fy2023")), ("h_cl", "FY2023 Google Cloud 營收", fv("is.rev.cloud.fy2023")),
          ("h_ob", "FY2023 Other Bets＋避險損益", fv("is.rev.ob.fy2023") + fv("is.rev.hedge.fy2023"))]
    for k, labl, v in hv:
        assert hidx[k] == r, (k, hidx[k], r)
        P.cell(r, 1, labl); P.cell(r, 2, "US$B" if "EPS" not in labl and "股數" not in labl else ("US$" if "EPS" in labl else "十億股"))
        c = P.cell(r, 3, v); c.number_format = "#,##0.000" if "EPS" not in labl else "0.00"; c.font = BLUE
        P.cell(r, 14, src23)
        r += 1
    P.cell(r, 1, "FY2023 毛利（營收 − 營業成本；核對 10-K 毛利 174.062）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=C{hidx['h_rev']}-C{hidx['h_cogs']}").number_format = "#,##0.000"; assert hidx["h_gp"] == r; r += 1
    P.cell(r, 1, "非折舊營業成本率（9M26A；＝（營業成本 − 折舊）÷ 營收，Q 頁；預測期營業成本假設）"); P.cell(r, 2, "%")
    P.cell(r, 3, f"=(Q!L{QR['cogs']}-Q!L{QR['da']})/Q!L{QR['rev']}").number_format = "0.0%"
    P.cell(r, 14, "Derived：假設折舊（現金流量表 PP&E 折舊）全數記入營業成本，其餘營業成本占營收比例不變；Assumed（模型未分列營業成本）")
    assert hidx["h_a_cex"] == r; dn(wb, "cogs_ex_r", f"PL!$C${r}"); r += 1
    P.cell(r, 1, "2025 年前三季合計營收（Q 頁 2025Q1–Q3；9M26A 年增基期）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=SUM(Q!C{QR['rev']}:E{QR['rev']})").number_format = "#,##0.000"; assert hidx["h_9m25rev"] == r; r += 1
    P.cell(r, 1, "2025 年前三季合計淨利（Q 頁；9M26A 年增基期）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=SUM(Q!C{QR['ni']}:E{QR['ni']})").number_format = "#,##0.000"; assert hidx["h_9m25ni"] == r; r += 1
    for col in "CDEFGHIJKLM":
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(idx={**idx}, ridx=ridx, cidx=cidx, sidx=sidx, hidx=hidx, ratio=r_ratio, chk=r_chk, seg=r_seg, h23=r_h23)
    return L


# =====================================================================================================
# RevDCF：反向 DCF（版面規格 v1.3；列＝A–H，一頁）
def revdcf(wb, C):
    V, P = wb["Val"], wb["PL"]
    VR, DC, PRB = C["VR"], C["DC"]["B"], C["PR"]["B"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    ev, nc, hold, dl, ob, pf, shv = (f"Val!$C${VR[k]}" for k in ("ev", "nc0", "hold", "dl", "obv", "ppref", "shv"))
    cols = ["F", "H", "I", "J", "K", "L", "M"]      # Q4 26E、FY2027E–FY2032E
    tr, dfr, ebr, tvn, shd, ufcf = DC["t"], DC["df"], DC["ebit"], DC["tvn"], DC["shd"], DC["ufcf"]
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.3.9 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益為線性，用封閉解；WACC 用二分法）。本頁只讀 Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 每股價值推到目標時點＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利；自用 AI 依 In 頁讀法）；"
                "分部加總部分反解「情境 B 分部加總每股價值＝現價」。反解值超出合理範圍時照實列出，不外推。營業利益率與營收規模兩項只動 FY2028 起的財報口徑營業利益（FY2026 Q4 與 FY2027 為錨點不動；折舊、資本支出、營運資金、自用 AI 影子效益不變）")
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
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（Val 頁 DCF 每股價值推到目標時點的反算）")
    kv("sh", "股數（十億股；評價日期末稀釋＋強制轉換特別股）", f"={shv}", "0.000")
    kv("oth", "淨現金、持股、Other Bets 減類債與特別股股利現值（US$B）", f"={nc}+{hold}-{dl}-{pf}+{ob}", "#,##0.0", "同 Val 頁 DCF 每股價值公式的非 EV 項")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{A['v0']}*B{A['sh']}-B{A['oth']}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；主值）", f"={ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=(B{A['ev']}+B{A['oth']})/B{A['sh']}", "0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）", "FY2026 Q4", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, vc in enumerate(cols):
            RV.cell(r, 2 + i, fn(vc, "BCDEFGH"[i])).number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$B）", lambda vc, c: f"=PL!{vc}{PRB['rev']}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$B；財報口徑）", lambda vc, c: f"=Val!{vc}{ebr}", "#,##0.0")
    yr("m", "營業利益率", lambda vc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("df", "折現因子", lambda vc, c: f"=Val!{vc}{dfr}", "0.0000")
    yr("tvf", "終值係數（只在 FY2032：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）",
       lambda vc, c: ("=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^(Val!$M$" + str(tr) + "+mid)" if vc == "M" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（Q4 26 與 FY2027 為錨點不動，記 0）", lambda vc, c: ("=0" if vc in ("F", "H") else f"=(1-tax)*({c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(B{Y['rev']}:H{Y['rev']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與 FY2028 起營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032，以同一係數加權）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "終值年（FY2032）權重大")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內"))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收與營業利益等比例放大；營業利益率不變，折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("EB", "FY2028–FY2032 營業利益 × 貢獻係數合計（US$B；放大一倍所增加的 EV）", f"=SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=1+(B{A['evr']}-B{A['ev']})/B{A['EB']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 Val 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def ev_at(w):
        cf = f"Val!$F${ufcf}:$M${ufcf}+(internal_ai_reading=2)*Val!$F${shd}:$M${shd}"
        tvc = f"(Val!$M${DC['tvn']}+(internal_ai_reading=2)*Val!$M${shd})"
        return (f"SUMPRODUCT({cf},1/(1+{w})^Val!$F${tr}:$M${tr})"
                f"+{tvc}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^(Val!$M${tr}+mid)")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, f"=({ev_at(f'D{r}')}+$B${A['oth']})/$B${A['sh']}")
        for c in (2, 3, 4):
            RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式的 EV − Val 頁情境 B EV（應為 0）", f"={ev_at('wacc')}-B{A['ev']}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的評價日每股價值 − 支持現價所需值（美元，應接近 0）", f"=E{b0 + 39}-B{A['v0']}", "0.000000")
    r += 1
    RV.cell(r, 1, "B. 分部加總（Comparable）：支持現價所需的倍數").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    sev, sebt, snc, shold, sdl, ssh, simp = (f"Val!$C${VR[k]}" for k in ("sev", "sebt", "snc", "shold", "sdl", "ssh", "simp"))
    stp = f"Val!$C${VR['sotp']}"

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    cmp("sps", "情境 B 分部加總每股價值（美元；目標時點）", "=px", f"={stp}", "0.00")
    cmp("sev", "情境 B 分部加總 EV（US$B；只看分部加總、其他項不變）", f"=px*{ssh}-({snc}+{shold}-{sdl})", f"={sev}", "#,##0.0")
    cmp("smul", "隱含 EV ÷ NTM EBITDA（倍）", f"=B{A['sev']}/{sebt}", f"={simp}", "0.0")
    cmp("scmp", "加權（DCF 不變）：支持現價所需的分部加總每股價值（情境加權；美元）", f"=(px-w_dcf*Val!$B${VR['tpdcf']})/(1-w_dcf)", f"=Val!$B${VR['tpcomp']}", "0.00")
    RV.cell(r, 1, "同業倍數：數位廣告（六家中位數）與採用倍數見 Val 頁 0.2 節；此處只列整體隱含倍數，不拆到分部"); r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：情境 B 的企業價值 "), f'TEXT(B{A["ev"]},"#,##0")', T(" 十億美元，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")', T(" 美元（評價日）；要支持現價 "), 'TEXT(px,"0.00")',
                    T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.00")', T(" 倍。單靠一個因素達成時，需要：折現率降到 "), f'TEXT(B{A["w1"]},"0.0%")', T("（模型 "), f'TEXT(B{A["w0"]},"0.0%")',
                    T("）；或 FY2028–FY2032 營收與營業利益全面放大 "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                    T("）；或預測期營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                    f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("分部加總法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於情境 B 分部加總 EV ÷ NTM EBITDA 約 "), f'TEXT(B{A["smul"]},"0.0")', T(" 倍，模型採用各分部同業倍數，整體隱含 "), f'TEXT(C{A["smul"]},"0.0")',
                    T(" 倍；換言之，只看分部加總，市場價格對應的整體倍數比模型 "), f'IF(B{A["smul"]}<C{A["smul"]},"低","高")', f'TEXT(ABS(D{A["smul"]}),"0%")', T("（EV 差距）。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 76
    for c in "BCDEFGH":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["b0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# =====================================================================================================
# Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列）
TOR_PLAIN = {
    "IF_HoldEcon 四世代同時取低／高（每 MW 年持有成本；k 的分母同步變動）": "四個 GPU 世代的每 MW 年持有成本同時取低／高",
    "Tokenomics IF_NonNV_Hold_TPU（tpu_f、tpu_cf 同時跟隨）": "TPU 持有成本比（收入與資本支出係數同時跟隨）",
    "TPU 收入係數改用產出比低／高端（資本支出仍用持有比）": "TPU 收入係數改用產出比低／高端（資本支出仍用持有比）",
}


def sens_tor(wb, C):
    SE = wb["Sens"]
    r0 = SE.max_row + 3
    SE.cell(r0, 1, "加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；本節只引用上方各案例的相對基準，只新增、不改既有格）").font = BOLD
    for j, h in enumerate(["輸入（白話）", "輸入取低值", "輸入取高值"], 1):
        c = SE.cell(r0 + 1, j, h); c.font = BOLD; c.fill = HDR
    r = r0 + 2
    pairs = {}
    order = []
    for key, side, val, rr in C["SE_ROWS"]:
        lab_ = str(SE.cell(rr, 1).value)
        base = re.sub(r"：(低|高)$", "", lab_)
        if base not in pairs:
            pairs[base] = {}
            order.append(base)
        pairs[base][side] = rr
    for base in order:
        plain = TOR_PLAIN.get(base, base)
        SE.cell(r, 1, plain)
        for col, side in (("B", "低"), ("C", "高")):
            if side in pairs[base]:
                SE[f"{col}{r}"] = f"=Sens!E{pairs[base][side]}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調（讀取層級＋C0 實績對帳；MAG 類樣板——Microsoft、Meta、SpaceX、Oracle 複製）
def coord(wb, C):
    TK, TKV, tk, GEN, QR = C["TK"], C["TKV"], C["tk"], C["GEN"], C["QR"]
    W = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_c0_v0.3.2.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    TKS = f"Tokenomics {TKV}（{TK['source']['file']}，master {TK['source']['commit'][:7]}；快照 alphabet/data/tokenomics_snapshot_{TKV}.json）"
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJK", (30, 12, 16, 16, 12, 28, 60, 22, 60, 22)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.3.9 起；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。Alphabet 是算力買方（雲端與自用 AI 的資本支出方），不做兩讀數目標價；"
                "C0 的個股側只用公司已公布實績（新聞稿／10-Q 財報表），層級側用 Wattonomics v0.3.2 隱含值（玩家資本支出、上線 GW）與 Tokenomics 每 GW 營收錨，同源者標「參考」。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    for j, h in enumerate(["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1a = r
    ws.cell(r, 1, "讀取層級：Tokenomics（算力單位經濟）——本模型讀 v21.10 快照（工作單所列現行 v22.0 不在本輪換快照；換快照另走 Tokenomics 同步通知）；Wattonomics——本模型不讀（無 WIF_ 名稱進模型），v0.3.2 快照只用於下方 C0 對帳列"); ws.cell(r, 1).font = BOLD; r += 1

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
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_HoldEcon｜{code}", code, tk("IF_HoldEcon", code), f"=he_{nm}", "US$m/MW-年（＝$B／GW／年）", "每 MW 年持有成本（經濟口徑；收入錨 k＝1）"))
    tkrows.append(("IF_GPUhrEcon｜H100", "H100", tk("IF_GPUhrEcon", "H100"), "=gpuhr_h", "US$/GPU-hr", "k 的現貨分母"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_OpexGW｜{code}", code, tk("IF_OpexGW", code), f"=opx_{nm}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_CapexIT｜{code}", code, tk("IF_CapexIT", code), f"=cit_{nm}", "US$m/MW", "每 MW IT 資本支出"))
    tkrows.append(("IF_CapexFacility｜VR200", "VR200", tk("IF_CapexFacility", "VR200"), "=cfac", "US$m/MW", "每 MW 機房資本支出"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3")):
        tkrows.append((f"IF_DeprIT｜{code}", code, tk("IF_DeprIT", code), f"=dpi_{nm}", "US$m/MW-年", "每 MW IT 年折舊（期初機隊汰換）"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IFW_RevGWFleet_Util｜{code}", code, tk("IFW_RevGWFleet_Util", code), f"=cap_{nm}", "US$m/MW-年", "上限檢查（不得作預測）"))
    tkrows.append(("IF_NonNV_Hold_TPU（基準）", "TPU v7", TK["items"]["IF_NonNV_Hold_TPU"]["values"][1], "=tpu_hold", "x（÷ VR200）", "TPU 持有成本比（Tokenomics 標 Assumed）"))
    tkrows.append(("IF_NonNV_Out_TPU（基準）", "TPU v7", TK["items"]["IF_NonNV_Out_TPU"]["values"][1], "=tpu_out", "x（÷ VR200）", "TPU 產出比（只用於 Sens）"))
    for nm_, yr_, lv, ad, un, nt in tkrows:
        F1R[nm_] = f1(nm_, yr_, lv, ad, un, TKS, nt, "0.000")
    F1b = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{F1a + 2}:D{F1b}-C{F1a + 2}:C{F1b}))").number_format = "0.000000"; F1chk = r; r += 1
    ws.cell(r, 1, "Wattonomics v0.3.2 快照名稱（只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (("WIF_Gap_Player_googl_2026_CapexAI", "capai", "US$B", "Alphabet 2026 AI 資料中心資本支出（Gap!G108←G91＝新增 GW × Tokenomics 每 GW 成本＋次年、後年廠房提前；GW 由 CapexGW 從公司資本支出（指引中點）反推，見下方追公式）"),
                             ("WIF_Gap_Player_googl_2026_CapexNonAI", "capnai", "US$B", "非 AI 資本支出（Gap!G98＝CapexGW!G49 × (1 − In!I13 AI 占比)＝公司資本支出 200 × 0.15，直接由指引中點推得；Assumed）"),
                             ("WIF_Gap_Player_googl_2026_AddGW", "addgw", "GW", "2026 年新增 IT GW（Gap!G84←CapexGW!G65←G63←G62←G61←G60←G52、G57、G58；輸入為公司指引＋公司 2024–25 實績＋個股模型 2027–28 寫回，見追公式；不是 Epoch 或園區證據推得）"),
                             ("WIF_PlayerGW_YE_googl_2025", "ye25", "GW", "Alphabet 2025 年底上線 IT GW（CapexGW!F66←Players!F49←F42＝Epoch 晶片功耗 2,953 MW × Google 倍數 1.456（In!I35 ovr25_googl 4,300 MW，Derived）；與公司資本支出無關，不同源）"),
                             ("WIF_PlayerGW_YE_googl_2026", "ye26", "GW", "Alphabet 2026 年底上線 IT GW（CapexGW!G66＝2025 年底（Epoch，不同源）＋2026 新增 GW（同源於公司指引等，見追公式））"),
                             ("WIF_PlayerGW_YE_googl_SiteOnly_2026", "ye26s", "GW", "2026 年底上線 IT GW（只有園區證據路徑）"),
                             ("WIF_Gap_Player_googl_2027_CapexAIMC", "cap27", "US$B", "並列讀法：2027 AI 資本支出（個股模型值；2026 無此值）")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt, "0.000")
    wr_ = WR["cap_googl"]
    wrows["guid"] = f1("In 列 gcap_googl（K175）｜2026（資本支出，財年）", "2026", wr_["values"][5 - 2 + 0] if False else wr_["values"][3], None, "US$B", WV, "Wattonomics In!K175（gcap_googl，Alphabet 個股模型 v0.3.3 基準寫回：FY2026E＝9M 實績＋Q4 公式＝指引中點 200；Interested-party，公司指引不是實績）；是 CapexGW 2026 年新增 GW 與非 AI 資本支出的輸入", "0.0")
    wrows["self"] = f1("In 列 al_googl_self｜自用比例", "2026", WR["al_googl_self"]["values"][3], None, "比", WV, "自用＝1 − 對外 30%（Wattonomics 沿用 Alphabet 舊模型；Assumed）；對外比例與本模型 ext_sh 同源", "0.00")
    wrows["nnv"] = f1("In 列 nnv_googl｜新增 GW 非 NVIDIA（TPU）占比", "2025", WR["nnv_googl"]["values"][1], None, "比", WV, "Epoch 2025 年底 TPU 2,223 MW ÷（2,223＋730）；Derived", "0.000")
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 Alphabet 路徑是否讀公司實績或目標（v0.3.2 快照對應的 20261114_Wattonomics_v0.3.2.xlsx；規則：有讀者的部分屬同源，不當獨立證據）").font = BOLD; r += 1
    for j, h in enumerate(["輸出名稱 ← 公式鏈", "", "", "", "", "", "同源判定", "讀到什麼（輸入、公式、儲存格）"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    TR = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_trace_v0.3.2.json"), encoding="utf-8"))
    TRC = [
        ("WIF_Gap_Player_googl_2026_AddGW ← Interface!F617 ← Gap!G84 ← CapexGW!G65 ← G63（上線年時差加權；lag_on、inv_on）← G62 ← G61 ← G60 ← G52（曆年 AI 資本支出）、G57／G58（前一、前二年 AI 資本支出，廠房提前入帳）← G49 ← G48 ← G44 ← G42 ← In!K175",
         "同源（公司指引＋公司 2024–25 實績＋個股模型 2027–28 寫回）",
         "公司指引：2026 財年資本支出 200（In!K175 gcap_googl＝Alphabet 個股模型 v0.3.3 基準寫回，FY2026E＝9M 實績＋Q4 公式＝指引中點；備援 In!K122 cap_googl 同為 200）× AI 占比 ai_sh（In!I13，0.85，Assumed）。"
         "公司 2024–25 實績：G57、G58 讀 2025（91.447）與 2024（52.535）財年資本支出（In!J175、I175），G63 的上線時差再加權前一年 GW。"
         "個股模型 2027–28 寫回：廠房提前入帳（Gap!G89、G90）讀 CapexGW!H65、I65，其輸入為 In!L175、M175（Alphabet 個股模型基準 2027–28 資本支出）。GW＝資金 ÷ 每 GW 成本（Tokenomics），所以 2026 年新增 GW 是公司資本支出換算，不是獨立的物理證據"),
        ("WIF_Gap_Player_googl_2026_CapexAI ← Interface!F609 ← Gap!G108 ← G91（＝G88＋G89＋G90）", "同源（同上）",
         "G88＝2026 新增 GW（上列）×（每 GW IT 成本＋廠房）；G89、G90＝2027、2028 新增 GW 的廠房提前入帳，讀 In!L175、M175（個股模型寫回）；G102（個股模型 AI 資本支出 In!K173）2026 空白，不替代"),
        ("WIF_Gap_Player_googl_2026_CapexNonAI ← Interface!F610 ← Gap!G98", "同源（公司指引）",
         "G96（In!K174 個股模型非 AI 資本支出）空白，改用 CapexGW!G49 × (1 − ai_sh)＝200 × 0.15＝30；200 即公司指引中點。Wattonomics 對此列已標 Assumed（待證據）"),
        ("WIF_PlayerGW_YE_googl_2025 ← Interface!E25 ← CapexGW!F66 ← Players!F49 ← F43 ← F42 ← In!J99 × Players!D15（＝In!I35 ovr25_googl ÷ In!J99）", "不同源（Epoch 晶片功耗 × 倍數）",
         "2025 年底 4.3 GW＝Epoch 晶片功耗 2,953 MW × Google 倍數 1.456（ovr25_googl 4,300 MW；Derived）；與公司資本支出無關，也不讀公司實績"),
        ("WIF_PlayerGW_YE_googl_2026 ← Interface!F25 ← CapexGW!G66＝F66＋G65", "部分同源",
         "2025 年底（Epoch 推得，不同源）＋2026 年增量 G65（同源於公司指引等，見第一列）。層級另有 SiteOnly 路徑 WIF_PlayerGW_YE_googl_SiteOnly_2026（Epoch 加園區證據，約 5.59 GW），但公司不揭露 GW，仍不能做 C0"),
        ("act26_googl（Wattonomics In 頁 act26_googl；9M 資本支出實績 131.953）", "未進 GW／資本支出路徑",
         "反向查詢：公式讀 act26_googl 的只有 " + "、".join(sorted({x.split("!")[0] for x in TR["act26_readers"]})) + " 頁（" + str(len(TR["act26_readers"])) + " 格；Compare 頁第 231 列作 ev_actyr 年的權重），不回饋 CapexGW、Gap、Interface 的 Alphabet 值；本表直接以公司新聞稿與 10-Q 的同一實績為個股側，不經 Wattonomics"),
        ("WIF_Gap_Player_googl_2026_CapexAIModel ← Gap!G102 ← G101 ← In!K173", "空白（無回寫值）",
         "2026 年個股模型 AI 資本支出 In!K173 為空（9M 未拆 AI／非 AI），所以這一欄不替代 GW 推算；2027 年起的並列讀法才讀個股模型（本表不用）"),
        ("TPU 相關推得值：In!nnv_googl（2025 年新增 GW 非 NVIDIA 占比）、Players!G57、Tokenomics IF_NonNV_Hold_TPU", "不同源",
         "nnv 來自 Epoch 晶片擁有者資料（TPU 2,223 MW ÷ 合計 2,953 MW），不讀公司實績或目標；TPU 持有成本比來自 Tokenomics（本模型同讀，屬層級共用）。公司不揭露 TPU 金額，因此沒有可比實績"),
    ]
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 2026 年 Alphabet 的資本支出（AI＋非 AI）與新增 GW 由公司指引（2026 中點 200）、公司 2024–25 實績（廠房提前入帳、上線時差）與個股模型 2027–28 寫回換算，整列同源；2025 年底 GW 來自 Epoch，不同源；公司實績 act26_googl 未進路徑。層級有 SiteOnly 路徑（約 5.59 GW，Epoch 加園區），但公司不揭露 GW，每 GW 資本支出也不可比；因此資本支出列標「參考（同源校準）」，結論「無可比實績（層級值同源）」。若 Wattonomics 改為只用 Epoch／園區證據推 2026 GW，才恢復比較（證據回饋走對帳單 C5）。")
    rtr_conc = r; r += 1
    ws.cell(r, 1, "追公式傾印（由 alphabet/trace_wattonomics.py 產生，貼程式輸出；每格公式與值；來源 " + TR["source"].split("/")[-1] + "；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 7, "標記"); ws.cell(r, 8, "公式（省略開頭等號）"); ws.cell(r, 4, "值"); r += 1
    rdump0 = r
    for x in TR["rows"]:
        ws.cell(r, 1, ("　" * x["depth"]) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 7, x["flag"]); ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    rtr = rtr_conc
    r += 1
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝新聞稿／10-Q 財報表實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    for j, h in enumerate(["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級", "出處網址"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    F2hh = r; r += 1
    SRC_Q = "Alphabet 2026Q3 Form 10-Q 現金流量表「購置不動產與設備」（網址見本列 L 欄）；2026Q3 財報新聞稿同值（mag/extract/docs/2026Q3/text/alphabet_2026Q3_新聞稿.txt 第 283 行；補充表第 388 行逐季）；Q 頁第 " + str(QR["capex"]) + " 列。融資租賃本金 9M 為 1.95（10-Q Note 4），不另調整"
    GR_Q = "Verified（10-Q 財報表；Wattonomics act26_googl 同值標 Verified）"
    F2a = r

    def inp(label, per, unit, f_act, src, grade, layer_f=None, lsrc="", lgrade=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, per); ws.cell(r, 3, unit)
        ws.cell(r, 4, f_act).number_format = "#,##0.000"
        if layer_f is not None:
            ws.cell(r, 5, layer_f).number_format = "#,##0.000"
        ws.cell(r, 8, src); ws.cell(r, 9, grade); ws.cell(r, 10, lsrc); ws.cell(r, 11, lgrade); ws.cell(r, 12, U_10Q)
        r += 1
        return r - 1
    rq1 = inp("資本支出 2026Q1（公司實績）", "2026Q1", "US$B", f"=Q!G{QR['capex']}", SRC_Q, GR_Q)
    rq2 = inp("資本支出 2026Q2（公司實績）", "2026Q2", "US$B", f"=Q!H{QR['capex']}", SRC_Q, GR_Q)
    rq3 = inp("資本支出 2026Q3（公司實績）", "2026Q3", "US$B", f"=Q!I{QR['capex']}", SRC_Q, GR_Q)
    rq9 = inp("資本支出 2026 年 9M 累計（公司實績）", "2026Q1–Q3（9 個月）", "US$B", f"=Q!L{QR['capex']}", SRC_Q, GR_Q)
    ws.cell(r, 1, "核對：Q1＋Q2＋Q3 − 9M 累計（應為 0）"); ws.cell(r, 4, f"=D{rq1}+D{rq2}+D{rq3}-D{rq9}").number_format = "0.000000"; r += 1
    rl = inp("層級 CY2026 玩家資本支出：AI＋非 AI（Wattonomics）", "CY2026（全年）", "US$B", None, "—", "—",
             f"=C{wrows['capai']}+C{wrows['capnai']}", f"{WV}：Gap 頁 Alphabet 2026 AI 資本支出（新增 {WN['WIF_Gap_Player_googl_2026_AddGW']:.2f} GW × Tokenomics 每 GW 成本，含廠房提前）＋非 AI 資本支出（連 F1 兩列）；兩者的輸入都是公司 2026 財年資本支出 200（指引中點，Alphabet 模型寫回 In!K175），屬同源（見下方追公式）", "Derived；輸入同源於公司指引")
    ws.cell(rl, 4).value = None
    ws.cell(r, 1, "主讀數：已公布期間（2026Q1–Q3）累計 vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計實績；層級 × 9/12（線性時差近似）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq9}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "9M 累計＝Q1＋Q2＋Q3（上列）"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "層級 CY2026 × 9 ÷ 12（Derived；線性時差近似）"); ws.cell(r, 11, "Derived")
    rd1 = r; r += 1
    ws.cell(r, 1, "第二讀數：最近一季（2026Q3）× 4 vs 層級 CY2026"); ws.cell(r, 2, "2026Q3 × 4（約等於該季中點的年化速率）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq3}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "2026Q3 單季 × 4（Derived）"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "層級 CY2026 全年（同上）"); ws.cell(r, 11, "Derived")
    rd2 = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：資本支出").font = BOLD
    ws.cell(r, 2, "層級值同源於公司資本支出指引（見追公式）：參考（同源校準），不判定")
    ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由公司指引中點換算＝同源，不是獨立證據；上方兩個讀數仍列出供參考"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL},SIGN(F{rd1})=SIGN(F{rd2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rjr = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：公司 2026 資本支出指引中點 vs 層級 CY2026（層級的輸入就是指引中點；差距來自 GW 換算與廠房提前入帳的時點）"); ws.cell(r, 2, "指引中點（公司指引不是實績）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, "=capex_g").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（指引，不判定）")
    ws.cell(r, 8, "2026Q3 法說：資本支出指引維持 195–205（中點；In 頁 capex_g；Interested-party）"); ws.cell(r, 9, "Interested-party（指引）"); ws.cell(r, 10, "同上"); ws.cell(r, 11, "Derived")
    rg = r; r += 1
    # ---- Cloud 營收（參考；無可比實績）----
    ws.cell(r, 1, "層級隱含 Cloud AI 營收（CY2026；參考）：年均上線 GW × 對外比例 × 每 GW 營收錨（k＝1）× TPU 調整").font = BOLD; r += 1
    ra1 = r; ws.cell(r, 1, "年均上線 IT GW（2025、2026 年底平均；Wattonomics）"); ws.cell(r, 3, "GW"); ws.cell(r, 5, f"=(C{wrows['ye25']}+C{wrows['ye26']})/2").number_format = "0.000"; ws.cell(r, 10, "Wattonomics PlayerGW_YE googl 2025、2026（連 F1）；2026 年增量由公司資本支出指引換算，屬同源（追公式）"); ws.cell(r, 11, "Derived（同源）"); r += 1
    ra2 = r; ws.cell(r, 1, "對外比例（1 − Wattonomics 自用比例；同源：Wattonomics 沿用 Alphabet 舊模型）"); ws.cell(r, 3, "比"); ws.cell(r, 5, f"=1-C{wrows['self']}").number_format = "0.00"; ws.cell(r, 10, "Wattonomics al_googl_self；Assumed"); ws.cell(r, 11, "Assumed"); r += 1
    ra3 = r; ws.cell(r, 1, "每 GW 年營收錨（k＝1；Tokenomics IF_HoldEcon × 本模型期初世代組合）"); ws.cell(r, 3, "US$B／GW"); ws.cell(r, 5, f"=mix0_h*C{F1R['IF_HoldEcon｜H100']}+mix0_b2*C{F1R['IF_HoldEcon｜GB200']}+mix0_b3*C{F1R['IF_HoldEcon｜GB300']}").number_format = "0.000"
    ws.cell(r, 10, "Tokenomics IF_HoldEcon（連 F1）× 本模型期初世代組合 mix0（Assumed）"); ws.cell(r, 11, "Derived／Assumed"); r += 1
    ra4 = r; ws.cell(r, 1, "TPU 調整＝(1 − TPU 占比) ＋ TPU 占比 × 持有成本比（Wattonomics 非 NVIDIA 占比；Tokenomics IF_NonNV_Hold_TPU）"); ws.cell(r, 3, "x"); ws.cell(r, 5, f"=(1-C{wrows['nnv']})+C{wrows['nnv']}*C{F1R['IF_NonNV_Hold_TPU（基準）']}").number_format = "0.000"
    ws.cell(r, 10, "「TPU 相關推得值」：Wattonomics nnv_googl（2025 年新增 GW 非 NVIDIA 占比）與 Tokenomics TPU 持有成本比"); ws.cell(r, 11, "Derived／Assumed"); r += 1
    ra5 = r; ws.cell(r, 1, "層級隱含 Cloud AI 營收（CY2026；US$B）"); ws.cell(r, 3, "US$B"); ws.cell(r, 5, f"=E{ra1}*E{ra2}*E{ra3}*E{ra4}").number_format = "#,##0.0"; ws.cell(r, 10, "上列四項相乘（層級的對外比例同源，k 取 1）"); ws.cell(r, 11, "Derived"); r += 1
    ws.cell(r, 1, "Google Cloud 分部營收：9M 累計（公司實績；含 Workspace 與非 AI 雲端）"); ws.cell(r, 2, "2026Q1–Q3"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!L{QR['cloud']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{ra5}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（無可比實績，不判定）")
    ws.cell(r, 8, "Alphabet 2026Q3 Form 10-Q Note 2 分部資訊 Google Cloud 營收（網址見本列 L 欄；新聞稿 Segment Operating Results 同值；Q 頁第 " + str(QR["cloud"]) + " 列）；公司不揭露 AI 營收，分部含 Workspace 與非 AI 雲端"); ws.cell(r, 9, "Verified（10-Q Note 2）")
    ws.cell(r, 10, "層級隱含 Cloud AI 營收 × 9/12；F 欄＝層級隱含 AI 營收占 Cloud 分部實績比例"); ws.cell(r, 11, "Derived／Assumed")
    ws.cell(r, 12, U_10Q); rc1 = r; r += 1
    ws.cell(r, 1, "Google Cloud 分部營收：2026Q3 × 4（公司實績年化）"); ws.cell(r, 2, "2026Q3 × 4"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!I{QR['cloud']}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{ra5}").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（無可比實績，不判定）"); ws.cell(r, 8, "同上（2026Q3 單季 × 4；10-Q Note 2）"); ws.cell(r, 9, "Verified（10-Q Note 2）"); ws.cell(r, 10, "層級隱含 Cloud AI 營收 CY2026"); ws.cell(r, 11, "Derived／Assumed")
    ws.cell(r, 12, U_10Q); rc2 = r; r += 1
    ws.cell(r, 1, "上限檢查（層級隱含 AI 營收不得高於 Cloud 分部實績）"); ws.cell(r, 7, f'=IF(AND(E{rc1}<=D{rc1},E{rc2}<=D{rc2}),"通過","不通過：層級隱含 AI 營收高於 Cloud 分部實績")'); rcu = r; r += 1
    ws.cell(r, 1, "TPU 系統銷售（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "Alphabet 不揭露 TPU 系統銷售金額（2026Q3 法說：Q3 已認列一部分、Q4 持續擴大，不揭露金額；事實總帳 guid.tpuRev.2026q3）；10-Q 只揭露 TPU 回購保證最大潛在付款 6.6（保證不是營收）；本模型 TPU 銷售為假設（Assumed）")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "層級隱含值已列於 Cloud AI 營收參考列的 TPU 調整"); ws.cell(r, 11, "Derived"); rtp = r; r += 1
    ws.cell(r, 1, "在役／簽約電力 MW（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "Alphabet 不揭露自家在役或簽約 AI 電力（MW／GW）；Epoch AI 晶片功耗與園區估計是第三方估計，不算公司實績；本模型 AI 在役 MW 為 Epoch 推得（Derived）再以資本支出指引重解速度（C10）")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "Wattonomics 上線 GW：2025 年底、2026 年底見 F1；層級值只供參考"); ws.cell(r, 11, "Derived"); rmw = r; r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "時差邊界（參考）項目數"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"時差邊界（參考）")'); rnb = r; r += 1
    ws.cell(r, 1, "無可比實績列數（資本支出 1、Cloud 營收 2、TPU 系統銷售 1、在役 MW 1）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")+COUNTIF(G{F2a}:G{F2b},"參考（無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F(T("C0（實績）對帳：公司已公布資本支出（2026 年 Q1–Q3 累計 "), f'TEXT(D{rd1},"0.0")', T(" 十億美元）與層級 CY2026 隱含值（"), f'TEXT(E{rl},"0.0")', T("；同期間換算 "), f'TEXT(E{rd1},"0.0")', T("）相差 "),
                    f'TEXT(F{rd1},"+0%;-0%")', T("（主讀數）；最近一季年化 "), f'TEXT(D{rd2},"0.0")', T(" 十億美元，相差 "), f'TEXT(F{rd2},"+0%;-0%")',
                    T("（第二讀數）。但 Wattonomics 的 Alphabet 2026 年資本支出與新增 GW 由公司資本支出指引中點換算（追公式：CapexGW!G49←In!K175），屬同源，判定："), f'G{rj}',
                    T("（參考，不計入對帳中）。Cloud AI 營收、TPU 系統銷售與在役電力沒有可比的公司實績（公司不揭露），不判定。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向（依季中點）")
    ws.cell(r, 2, F(T("資本支出是流量，逐季上升（Q1 "), f'TEXT(D{rq1},"0.0")', T("、Q2 "), f'TEXT(D{rq2},"0.0")', T("、Q3 "), f'TEXT(D{rq3},"0.0")',
                     T("）。主讀數用 9 個月平均速率（中點約 5 月中）代表全年，全年中點在 7 月初，所以在上升趨勢下主讀數偏低；第二讀數用 Q3 單季年化（中點約 8 月中），比全年平均速率的中點晚一個半月，所以偏高。"
                       "兩讀數相差 "), f'TEXT((F{rd2}-F{rd1})*100,"0")', T(" 個百分點，就是時點差的幅度。公司全年實績要等 2026Q4 財報（2027-02）；層級 CY2026 與公司全年指引（195–205）相差 "), f'TEXT(F{rg},"+0%;-0%")',
                     T("，來自 GW 換算與廠房提前入帳的時點。這裡全部用流量（資本支出）比較，不拿 9 月底存量對年底存量。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出列：追公式（F1b）顯示 Wattonomics 2026 年 Alphabet AI 資本支出（Gap!G108）與非 AI 資本支出（Gap!G98）的共同輸入是公司 2026 財年資本支出 200（CapexGW!G49←In!K175，Alphabet 個股模型基準寫回，＝指引中點）；"
                   "新增 GW 也由此換算（CapexGW!G65）。層級值因此同源於公司目標，不是獨立證據：整列標「參考（同源校準）」，結論「無可比實績（層級值同源）」；每 GW 資本支出同樣無可比（公司不揭露 GW）。公司實績 act26_googl 沒有進 Alphabet 的 GW 或資本支出路徑（只在 Compare 頁作權重）。"
                   "（二）Cloud AI 營收列：用到上述 GW（同源）、對外比例 30%（沿用 Alphabet 舊模型，同源）、期初世代組合（本模型假設）；公司不揭露 AI 營收，Cloud 分部又含非 AI，所以只列層級隱含 AI 營收占分部比例與上限檢查，不判定。"
                   "（三）TPU 銷售與在役 MW：公司不揭露，列「無可比實績」。（四）所以本頁 C0 在 Alphabet 無檢驗力：對帳中為 0 不代表兩層一致，只代表層級值是由公司指引換算而來。要恢復檢驗，需 Wattonomics 改以 Epoch 或園區證據獨立推 2026 年 GW。"))
    rsame = r; r += 1
    ws.cell(r, 1, "口徑換算（層級 vs 公司資本支出）")
    ws.cell(r, 2, ("公司資本支出＝購置不動產與設備（現金流量表）；層級資本支出＝新增 GW × 每 GW 建置成本（IT＋廠房）＋非 AI 資本支出。兩者都是現金／完工口徑近似；"
                   "Alphabet 融資租賃本金很小（9M 為 1.95，10-Q Note 4），本頁不另調整（其他公司須逐家判斷，租賃占比大者要調整）；Wattonomics 對 Alphabet 為擁有者口徑（租用 SpaceX／xAI Colossus 的 0.19 GW 屬出租方，不重複計）。"))
    rk = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 輸入的 2026 資本支出（指引中點）vs 其 Gap 頁 GW 推算的玩家資本支出（AI＋非 AI）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['guid']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics 資本支出輸入列 cap_googl（公司指引中點）"); ws.cell(r, 10, "Wattonomics Gap 頁玩家資本支出（GW 推算）"); rb_ = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：2026 資本支出指引（上列）、Cloud 積壓 RPO（In 頁事實）等不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：Tokenomics 本模型讀 v21.10（工作單所列現行 v22.0 未換；v21.10 換 v22.0 只影響 IFW_RevGWFleet_Util，約 −2%，只作上限檢查，不影響基準）；Wattonomics v0.3.2 也讀 Tokenomics v21.10，兩側版本一致。Alphabet 模型值曾被 Wattonomics 讀入（營業現金流、資本支出寫回），本表只列層級值與追公式，不用該回饋判定。")
    r += 1
    return dict(F1=(F1a, F1b), F1chk=F1chk, F2h=F2h, F2hh=F2hh, F2=(F2a, F2b), rq=(rq1, rq2, rq3, rq9), rl=rl, rd1=rd1, rd2=rd2, rj=rj, rjr=rjr, rtr=rtr, rdump=(rdump0, rdump1), rg=rg, ra=(ra1, ra2, ra3, ra4, ra5),
                rc1=rc1, rc2=rc2, rcu=rcu, rtp=rtp, rmw=rmw, rn=rn, rnb=rnb, rnc=rnc, rs=rs, rr=rr_, rsame=rsame, rk=rk, rb=rb_, wrows=wrows, tol=TOL)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, RVA, PLA, CRD):
    RD, V, P = wb["README"], wb["Val"], wb["PL"]
    VR, PRB, PRC = C["VR"], C["PR"]["B"], C["PR"]["C"]
    vb = lambda k, c="B": f"Val!{c}{VR[k]}"
    tp, up, rec, dcf, cmpv, cons = vb("tp"), vb("up"), vb("rec"), vb("tpdcf"), vb("tpcomp"), vb("cons")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    ox = PLA["idx"]
    G_ = f"CF!C{C['GATE']}"
    out = {}
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.3.9 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("：現金流折現 "), f'TEXT({dcf},"0")',
        T(" 與分部加總 "), f'TEXT({cmpv},"0")', T(" 各占一半。Cloud 成長足以抵銷搜尋被 AI 分流的損失，但 AI 折舊與資本支出把營業利益率從 "), f'TEXT(PL!D{PRB["om"]},"0%")', T(" 壓到 "), f'TEXT(PL!I{PRB["om"]},"0%")', T("（FY2028）。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：自用 AI 採價值中性（讀法 2）；改為只計成本（讀法 1）時加權目標價 "), f'TEXT(Val!B{VR["tp_r1"]},"0")', T(" 美元（"), f'TEXT(Val!C{VR["tp_r1"]},"+0%;-0%")', T("），評等"),
        f'Val!D{VR["tp_r1"]}', T("；Cloud 的資本支出轉成營收的速度符合 2026 年指引（195–205）的假設；現價 "), 'TEXT(px,"0.00")', T(" 美元為 2026-11-05 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價 "), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；市場共識 "), f'TEXT({cons},"0")', T("。三個情境（"), f'TEXT({vb("scen")},"0")', T("／"), f'TEXT({vb("scen", "C")},"0")', T("／"), f'TEXT({vb("scen", "D")},"0")', T("）都低於共識。"))
    put("o2", T("現金流折現偏貴：反向推算（情境 B），要支持現價需把折現率降到 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), f'TEXT({rv("w0")},"0.0%")', T("），或 FY2028 起營收與營業利益全面放大 "),
        f'TEXT({rv("s1")},"0.00")', T(" 倍；分部加總隱含 "), f'TEXT(RevDCF!B{RVA["smul"]},"0.0")', T(" 倍 EBITDA，模型 "), f'TEXT(RevDCF!C{RVA["smul"]},"0.0")', T(" 倍。"))
    put("o3", T("搜尋被侵蝕，Cloud 抵銷："), T("悲觀情境 FY2028 Search & other 比基準少 "), f'TEXT(1-PL!I{PRC["search"]}/PL!I{PRB["search"]},"0.0%")', T("，Google Services 營業利益少 "),
        f'TEXT(PL!I{PRB["oi_gs"]}-PL!I{PRC["oi_gs"]},"0.0")', T(" 十億美元；同年 Cloud 營業利益比 FY2025 增加 "), f'TEXT(PL!I{PRB["oi_cl"]}-PL!D{PRB["oi_cl"]},"0.0")', T("，足以抵銷。"))
    put("o4", T("真正壓力在 AI 折舊與現金：基準情境新增 AI 折舊 FY2028 "), f'TEXT(PL!I{PRB["da_ai"]},"0")', T(" 十億美元，毛利率由 FY2025 "), f'TEXT(PL!E{ox["gm"]},"0.0%")', T(" 降到 FY2028 "),
        f'TEXT(PL!H{ox["gm"]},"0.0%")', T("；五年自由現金流扣股利合計 "), f'TEXT({G_},"0.0")', T(" 十億美元（基準），"), f'CF!C{C["GATE"] + 3}', T("。"))
    put("o5", T("實績對帳（本版新增）：2026 年 Q1–Q3 資本支出 "), f'TEXT(Coord!D{CRD["rd1"]},"0.0")', T(" 十億美元，與 Wattonomics 隱含的 CY2026 同期間值相差 "), f'TEXT(Coord!F{CRD["rd1"]},"+0%;-0%")',
        T("（最近一季年化 "), f'TEXT(Coord!F{CRD["rd2"]},"+0%;-0%")', T("），但該層級值由公司資本支出指引換算（同源），結論："), f'Coord!G{CRD["rj"]}', T("；AI 營收、TPU 與在役電力公司不揭露，無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：Google Services＝Search & other（B 路徑年增，情境 A／C 再乘分流與變現變動）＋YouTube ads＋Google Network＋訂閱與裝置（各自近四季年增向長期值收斂）；"
                 "Google Cloud＝Cloud AI（AI 在役 MW × 對外比例 × 每 MW 年收入）＋TPU 系統銷售＋Cloud 其他（近四季年增收斂）；另加 Other Bets 與避險損益，加總為合併營收。FY2028 Services "),
        f'TEXT(PL!I{PRB["gs"]},"0")', T("、Cloud "), f'TEXT(PL!I{PRB["cloud"]},"0")', T("（其中 Cloud AI "), f'TEXT(PL!I{PRB["ai"]},"0")', T("）、Other Bets "), f'TEXT(PL!I{PRB["ob"]},"0.0")', T("，合計 "), f'TEXT(PL!I{PRB["rev"]},"0")', T(" 十億美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁（反向 DCF）、PL 頁末標準損益表"
                  "（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，GAAP）、Sens 頁末加權目標價相對基準（龍捲風圖用）、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」（F1 讀取層級；F2 C0（實績）：公司資本支出與 Cloud 營收 vs 層級隱含值）。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, C):
    PLA = pl_standard(wb, C)
    RVA = revdcf(wb, C)
    tor = sens_tor(wb, C)
    CRD = coord(wb, C)
    RMA = readme13(wb, C, RVA, PLA, CRD)
    return dict(PL=PLA, RV=RVA, tor=tor, CRD=CRD, RM=RMA)
