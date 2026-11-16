# Meta v0.2.1（2026-11-15）後處理：個股 HTML 版面規格 v1.3 全套＋Coord 頁「三層協調」節（MAG 類樣板；照 alphabet/v039_post.py 改寫，Meta 的損益、DCF 與資料來源不同處逐項調整）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點；樣板 Alphabet v0.3.9。
# 只新增，不改任何既有儲存格的數值（逐格比對只允許版號／build_id 文字、Recon 第 7 節判定欄文字與新增格）：
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；GAAP；另列 9M26A 比較欄）、比率區、核對列、FY2023 實績輸入、分部營收（圖表用）
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（情境 B 路徑，封閉解與二分法）與分部加總隱含倍數；只讀其他頁，無任何頁引用本頁
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用，白話標籤）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v22.0；Wattonomics v0.3 快照）、F1b 追公式（tools/coord/trace_wattonomics.py 輸出）、F2 C0（實績）：10-Q／財報資本支出 vs 層級隱含值
#   README 頁末：v1.3 展示層白話列（橫幅、重點觀察 5 點、分部營收推導、版本沿革）
# build_meta.py 在 README 之後、存檔前呼叫 apply(wb, ctx)。
import json
import os
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.2", "2026-11-16"
U_10Q = "https://www.sec.gov/Archives/edgar/data/0001326801/000162828026066773/meta-20260930.htm"   # 出處網址只放 Coord 頁 L 欄（HTML 不顯示外部網址）
U_PR = "https://www.sec.gov/Archives/edgar/data/0001326801/000162828026066700/meta-09302026xexhibit991.htm"
U_Q4_23 = "https://investor.atmeta.com/investor-news/press-release-details/2024/Meta-Reports-Fourth-Quarter-and-Full-Year-2023-Results-Initiates-Quarterly-Dividend/"
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")


def lit(t):
    """Excel 公式字串常數上限 255 字元：切段以 & 相接。"""
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def T(s):
    return ("t", s)


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# FY2023 實績（US$ 百萬；Meta 2023Q4 財報新聞稿 8-K 附件 99.1 之 FY2023 欄，2024-02-01；IR 頁與 SEC 檔各讀一次，營收、總成本、營業利益、所得稅、淨利、EPS 一致；
# 其餘明細另以恆等式核對：成本四項合計＝總成本、分部營業利益合計＝合併營業利益、稅前－稅＝淨利、淨利÷股數＝EPS）
FY23 = dict(rev=134902, cogs=25959, rd=38483, ms=12301, ga=11408, costs=88151, oi=46751, oie=677, pt=47428, tax=8330, ni=39098, eps=14.87, sh=2629,
            ads=131948, other=1058, foa=133006, rl=1896, oi_foa=62871, oi_rl=-16120)

# =====================================================================================================
# PL：標準損益表（版面規格 v1.3 D 節）
# 欄位：C FY2023A、D FY2024A、E FY2025A、F–L FY2026E–FY2032E（連續年度，供圖表）、M 9M26A、N 說明
# 對應 PL 頁（情境 B 區塊）欄：D→C（FY2024A）、E→D（FY2025A）、F→G（FY2026E）、…、L→M（FY2032E）；M(9M26A)→E
SCOL = "CDEFGHIJKL"
S2P = {"D": "C", "E": "D", "F": "G", "G": "H", "H": "I", "I": "J", "J": "K", "K": "L", "L": "M", "M": "E"}


def pl_standard(wb, C):
    P, Q, B = wb["PL"], wb["Q"], C["PR"]["B"]
    QR = C["QR"]
    r = P.max_row + 3
    L = {}
    P.cell(r, 1, "標準損益表（v0.2.1 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2032E（情境 B 基準）；最右欄 9M26A 為最新累計實績。US$B；EPS 美元。只引用本頁上方、Q 頁與下方 FY2023 實績輸入，不回饋任何計算").font = BOLD
    L["start"] = r; r += 1
    P.cell(r, 1, "口徑：GAAP（財報口徑，含一次性項目）。Meta 以 GAAP 揭露損益；FY2025A 稅額含一次性稅務費用、9M26A 營業利益含法律與遣散一次性費用 13.47、所得稅含 Q1 CAMT 稅務利益 8.03，預測期不估一次性（排除一次性口徑見 PL 頁情境 B 區塊）。"
                 "EPS＝歸屬普通股淨利（淨利減特別股股利）÷ 稀釋加權股數；FY2027E–FY2029E 有特別股股利，所以 EPS × 股數略小於淨利。")
    L["note"] = r; r += 1
    P.cell(r, 1, "營業成本與營業費用：模型以分部 EBITDA 與折舊推導營業利益，沒有分列營業成本與營業費用。預測期營業成本＝營收 × 非折舊營業成本率（9M26A 營業成本扣折舊後占營收，下方假設列）＋折舊合計（假設新增折舊全數記入營業成本；Meta 實際折舊分攤於營業成本、研發與管理費用，Assumed）；"
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

    rv, oi, da, oie, pt, tx, ni, sh, eps = (B[k] for k in ("rev", "oi", "da", "oie", "pt", "tx", "ni", "sh", "eps"))
    pr = lambda row_, c: f"=PL!{S2P[c]}{row_}"
    rows_order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    base = r
    idx = {k: base + i for i, k in enumerate(rows_order)}
    r_ratio = base + len(rows_order)
    ratio_keys = ["r_yoy", "r_gm", "r_opm", "r_nm", "r_niy"]
    ridx = {k: r_ratio + 1 + i for i, k in enumerate(ratio_keys)}
    r_chk = r_ratio + 1 + len(ratio_keys)
    chk_keys = ["c1", "c2", "c3", "c4", "c5"]
    cidx = {k: r_chk + 1 + i for i, k in enumerate(chk_keys)}
    r_seg = r_chk + 1 + len(chk_keys) + 1
    seg_keys = ["s_ads", "s_oth", "s_rl", "s_tot", "s_chk"]
    sidx = {k: r_seg + 1 + i for i, k in enumerate(seg_keys)}
    r_h23 = r_seg + 1 + len(seg_keys) + 1
    h23_keys = ["h_rev", "h_cogs", "h_opx", "h_oi", "h_oie", "h_pt", "h_tax", "h_ni", "h_eps", "h_sh", "h_ads", "h_oth", "h_rl", "h_gp", "h_a_cex", "h_9m25rev", "h_9m25ni"]
    hidx = {k: r_h23 + 2 + i for i, k in enumerate(h23_keys)}

    row("rev", "營收", "US$B", {"C": f"=C{hidx['h_rev']}", **{c: pr(rv, c) for c in "DEFGHIJKLM"}}, M, "分部營收加總（PL 頁情境 B 區塊）；FY2023A 為 2023Q4 新聞稿 FY2023 欄")
    assert L["rev"] == idx["rev"]
    yo = {c: f"={c}{idx['rev']}/{p}{idx['rev']}-1" for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    yo["M"] = f"=M{idx['rev']}/C{hidx['h_9m25rev']}-1"
    row("yoy", "營收年增率（YoY）", "%", yo, PCT, "FY2023A 的基期（FY2022）不在事實總帳，故空白；9M26A 對 2025 年前三季合計（Q 頁）")
    cg = {"C": f"=C{hidx['h_cogs']}", "D": f"=Q!J{QR['cogs']}", "E": f"=Q!K{QR['cogs']}", "M": f"=Q!L{QR['cogs']}",
          "F": f"=M{r}+PL!F{rv}*cogs_ex_r+PL!F{da}"}
    for c in "GHIJKL":
        cg[c] = f"={c}{idx['rev']}*cogs_ex_r+PL!{S2P[c]}{da}"
    row("cogs", "營業成本（COGS）", "US$B", cg, M, "歷史與 9M26A＝財報；FY2026E＝9M26A＋Q4 26E（營收 × 非折舊成本率＋折舊）；FY2027E 起＝營收 × 非折舊成本率＋折舊合計（假設列）")
    row("gp", "毛利", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['cogs']}" for c in "CDEFGHIJKLM"}, M)
    row("gm", "毛利率", "%", {c: f"={c}{idx['gp']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT, "預測期毛利率下降＝折舊（AI 資本支出）上升")
    ox = {"C": f"=C{hidx['h_opx']}", "D": f"=Q!J{QR['rd']}+Q!J{QR['ms']}+Q!J{QR['ga']}", "E": f"=Q!K{QR['rd']}+Q!K{QR['ms']}+Q!K{QR['ga']}",
          "M": f"=Q!L{QR['rd']}+Q!L{QR['ms']}+Q!L{QR['ga']}"}
    for c in "FGHIJKL":
        ox[c] = f"={c}{idx['gp']}-{c}{idx['oi']}"
    row("opx", "營業費用（研發＋行銷＋管理）", "US$B", ox, M, "歷史與 9M26A＝財報；預測期＝毛利 − 營業利益（殘差）")
    row("oi", "營業利益", "US$B", {"C": f"=C{hidx['h_oi']}", **{c: pr(oi, c) for c in "DEFGHIJKLM"}}, M, "財報口徑合併營業利益（PL 頁；不含自用 AI 影子效益）")
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("oth", "其他收入（淨額）", "US$B", {"C": f"=C{hidx['h_oie']}", **{c: pr(oie, c) for c in "DEFGHIJKLM"}}, M, "利息與其他收益淨額；預測期＝利息收入 − 利息費用（PL 頁）")
    row("pti", "稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}" for c in "CDEFGHIJKLM"}, M)
    row("tax", "所得稅", "US$B", {"C": f"=C{hidx['h_tax']}", **{c: pr(tx, c) for c in "DEFGHIJKLM"}}, M)
    row("ni", "稅後淨利", "US$B", {c: f"={c}{idx['pti']}-{c}{idx['tax']}" for c in "CDEFGHIJKLM"}, M, "稅前 − 所得稅（核對列：與 PL 頁淨利相同）")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, PCT)
    row("sh", "稀釋加權股數", "十億股", {"C": f"=C{hidx['h_sh']}", **{c: pr(sh, c) for c in "DEFGHIJKLM"}}, M, "FY2023A＝新聞稿稀釋加權股數；其餘取 PL 頁平均稀釋股數（每股為 Class C 股票股利 1:1 之前的口徑）")
    row("eps", "EPS（GAAP 稀釋）", "US$", {"C": f"=C{hidx['h_eps']}", **{c: pr(eps, c) for c in "DEFGHIJKLM"}}, "0.00", "FY2023A＝新聞稿公告值；其餘取 PL 頁（歸屬普通股淨利 ÷ 稀釋股數；FY2024A 與公告值差 0.01 內）")
    assert L["eps"] == idx["eps"] and r == r_ratio
    P.cell(r, 1, "比率").font = BOLD; r += 1
    row("r_yoy", "營收年增率", "%", {c: f"={c}{idx['yoy']}" for c in "DEFGHIJKLM"}, PCT)
    row("r_gm", "毛利率", "%", {c: f"={c}{idx['gm']}" for c in "CDEFGHIJKLM"}, PCT)
    row("r_opm", "營業利益率", "%", {c: f"={c}{idx['opm']}" for c in "CDEFGHIJKLM"}, PCT)
    row("r_nm", "淨利率", "%", {c: f"={c}{idx['nm']}" for c in "CDEFGHIJKLM"}, PCT)
    ny = {c: f"={c}{idx['ni']}/{p}{idx['ni']}-1" for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    ny["M"] = f"=M{idx['ni']}/C{hidx['h_9m25ni']}-1"
    row("r_niy", "淨利年增率（YoY）", "%", ny, PCT, "FY2025A 含一次性稅務費用、9M26A 含 Q1 CAMT 稅務利益，年增率受一次性項目影響")
    assert r == r_chk
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    row("c1", "核對：毛利 − 營業費用 − 營業利益", "US$B", {c: f"={c}{idx['gp']}-{c}{idx['opx']}-{c}{idx['oi']}" for c in "CDEFGHIJKLM"}, "0.000", "歷史與 9M26A：財報三項（營業成本、費用、營業利益）互相核對；預測期以殘差定義，恆為 0")
    row("c2", "核對：稅前利益 − PL 頁稅前利益", "US$B", {c: f"={c}{idx['pti']}-PL!{S2P[c]}{pt}" for c in "DEFGHIJKLM"}, "0.000")
    row("c3", "核對：稅後淨利 − PL 頁淨利", "US$B", {c: f"={c}{idx['ni']}-PL!{S2P[c]}{ni}" for c in "DEFGHIJKLM"}, "0.000")
    row("c4", "核對：營收 − PL 頁合併營收", "US$B", {c: f"={c}{idx['rev']}-PL!{S2P[c]}{rv}" for c in "DEFGHIJKLM"}, "0.000")
    row("c5", "核對：FY2023A 稅後淨利 − 新聞稿淨利（輸入列）", "US$B", {"C": f"=C{idx['ni']}-C{hidx['h_ni']}"}, "0.000")
    assert r == r_seg - 1
    r += 1
    P.cell(r, 1, "分部營收（US$B；圖表與驅動因子章用）").font = BOLD; r += 1
    segcells = lambda k23, pk: {"C": f"=C{hidx[k23]}", **{c: pr(B[pk], c) for c in "DEFGHIJKLM"}}
    row("s_ads", "Advertising（Family of Apps）", "US$B", segcells("h_ads", "ads"), M)
    row("s_oth", "Other revenue（Family of Apps）", "US$B", segcells("h_oth", "other"), M)
    row("s_rl", "Reality Labs", "US$B", segcells("h_rl", "rl"), M)
    row("s_tot", "合計", "US$B", {c: f"=SUM({c}{sidx['s_ads']}:{c}{sidx['s_rl']})" for c in "CDEFGHIJKLM"}, M)
    row("s_chk", "核對：合計 − 營收（應為 0）", "US$B", {c: f"={c}{sidx['s_tot']}-{c}{idx['rev']}" for c in "CDEFGHIJKLM"}, "0.000")
    assert r == r_h23 - 1
    r += 1
    P.cell(r, 1, "FY2023 實績輸入與假設（藍字＝輸入；FY2023 取 Meta 2023Q4 財報新聞稿〔8-K 附件 99.1，2024-02-01〕FY2023 欄，標 Verified；IR 頁與 SEC 檔各讀一次，恆等式另核對，見下方核對列）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "數值", "", "", "", "", "", "", "", "", "", "", "來源"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    src23 = "Meta 2023Q4 財報新聞稿 FY2023 欄（8-K 附件 99.1，2024-02-01；investor.atmeta.com 與 SEC EDGAR 0001326801-24-000010；Verified）"
    f = {k: v / 1000.0 for k, v in FY23.items() if k not in ("eps", "sh")}
    hv = [("h_rev", "FY2023 合併營收", f["rev"]), ("h_cogs", "FY2023 營業成本", f["cogs"]),
          ("h_opx", "FY2023 營業費用（研發＋行銷＋管理）", (FY23["rd"] + FY23["ms"] + FY23["ga"]) / 1000.0),
          ("h_oi", "FY2023 營業利益", f["oi"]), ("h_oie", "FY2023 其他收入（淨額）", f["oie"]),
          ("h_pt", "FY2023 稅前利益", f["pt"]), ("h_tax", "FY2023 所得稅", f["tax"]),
          ("h_ni", "FY2023 淨利", f["ni"]), ("h_eps", "FY2023 EPS（GAAP 稀釋，美元）", FY23["eps"]),
          ("h_sh", "FY2023 稀釋加權股數（十億股）", FY23["sh"] / 1000.0),
          ("h_ads", "FY2023 Advertising", f["ads"]), ("h_oth", "FY2023 Other revenue", f["other"]), ("h_rl", "FY2023 Reality Labs 營收", f["rl"])]
    for k, labl, v in hv:
        assert hidx[k] == r, (k, hidx[k], r)
        P.cell(r, 1, labl); P.cell(r, 2, "US$B" if "EPS" not in labl and "股數" not in labl else ("US$" if "EPS" in labl else "十億股"))
        c = P.cell(r, 3, v); c.number_format = "#,##0.000" if "EPS" not in labl else "0.00"; c.font = BLUE
        P.cell(r, 14, src23)
        r += 1
    P.cell(r, 1, "FY2023 毛利（營收 − 營業成本；新聞稿營業成本 25.959）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=C{hidx['h_rev']}-C{hidx['h_cogs']}").number_format = "#,##0.000"; assert hidx["h_gp"] == r; r += 1
    P.cell(r, 1, "非折舊營業成本率（9M26A；＝（營業成本 − 折舊與攤銷）÷ 營收，Q 頁；預測期營業成本假設）"); P.cell(r, 2, "%")
    P.cell(r, 3, f"=(Q!L{QR['cogs']}-Q!L{QR['da']})/Q!L{QR['rev']}").number_format = "0.0%"
    P.cell(r, 14, "Derived：假設折舊與攤銷（現金流量表）全數記入營業成本，其餘營業成本占營收比例不變；Assumed（模型未分列營業成本）")
    assert hidx["h_a_cex"] == r; dn(wb, "cogs_ex_r", f"PL!$C${r}"); r += 1
    P.cell(r, 1, "2025 年前三季合計營收（Q 頁 2025Q1–Q3；9M26A 年增基期）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=SUM(Q!C{QR['rev']}:E{QR['rev']})").number_format = "#,##0.000"; assert hidx["h_9m25rev"] == r; r += 1
    P.cell(r, 1, "2025 年前三季合計淨利（Q 頁；9M26A 年增基期）"); P.cell(r, 2, "US$B"); P.cell(r, 3, f"=SUM(Q!C{QR['ni']}:E{QR['ni']})").number_format = "#,##0.000"; assert hidx["h_9m25ni"] == r; r += 1
    # FY2023 恆等式核對（第二讀）：成本四項合計＝總成本、分部營業利益合計＝營業利益、稅前−稅＝淨利、淨利÷股數＝EPS
    P.cell(r, 1, "核對（FY2023 輸入恆等式；應為 0）：成本四項合計 − 新聞稿總成本 88.151 ＋ 營收 − 總成本 − 營業利益")
    P.cell(r, 3, f"=ABS(C{hidx['h_cogs']}+C{hidx['h_opx']}-{FY23['costs'] / 1000.0})+ABS(C{hidx['h_rev']}-{FY23['costs'] / 1000.0}-C{hidx['h_oi']})").number_format = "0.000000"; L["h_id1"] = r; r += 1
    P.cell(r, 1, "核對（FY2023 輸入恆等式；應為 0）：營業利益 ＋ 其他收入 − 稅前；稅前 − 所得稅 − 淨利；分部營業利益合計（62.871 − 16.120）− 營業利益；Advertising＋Other＋RL − 營收")
    P.cell(r, 3, f"=ABS(C{hidx['h_oi']}+C{hidx['h_oie']}-C{hidx['h_pt']})+ABS(C{hidx['h_pt']}-C{hidx['h_tax']}-C{hidx['h_ni']})+ABS({FY23['oi_foa'] / 1000.0}+({FY23['oi_rl'] / 1000.0})-C{hidx['h_oi']})+ABS(C{hidx['h_ads']}+C{hidx['h_oth']}+C{hidx['h_rl']}-C{hidx['h_rev']})").number_format = "0.000000"; L["h_id2"] = r; r += 1
    P.cell(r, 1, "核對（FY2023 EPS）：淨利 ÷ 稀釋股數 − 公告 EPS（四捨五入容差 0.01 內）")
    P.cell(r, 3, f"=C{hidx['h_ni']}/C{hidx['h_sh']}-C{hidx['h_eps']}").number_format = "0.0000"; L["h_id3"] = r; r += 1
    for col in "CDEFGHIJKLM":
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(idx={**idx}, ridx=ridx, cidx=cidx, sidx=sidx, hidx=hidx, ratio=r_ratio, chk=r_chk, seg=r_seg, h23=r_h23)
    return L


# =====================================================================================================
# RevDCF：反向 DCF（版面規格 v1.3；列＝A–H，一頁）。Meta 的 DCF：NOPAT＝(EBIT＋sbc_add×股權報酬)×(1−稅率)；終值 tv_alt＝0（價值驅動）或 1（FY2032 UFCF）；自用 AI 影子效益讀法 2 另計
def revdcf(wb, C):
    VR, DC, PRB = C["VR"], C["DC"]["B"], C["PR"]["B"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    ev, nc, post, agtv, dl, pf, shv = (f"Val!$C${VR[k]}" for k in ("ev", "nc0", "post", "agtv", "dl", "ppref", "shv"))
    cols = ["F", "H", "I", "J", "K", "L", "M"]      # Q4 26E、FY2027E–FY2032E
    ebr, nor, dar, cxr, wcr, agr, dadf, ufcf, shd, tr, dfr = (DC[k] for k in ("ebit", "nopat", "da", "cx", "wc", "ag", "dadf", "ufcf", "shd", "t", "df"))
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.1 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益為線性，用封閉解；WACC 用二分法）。本頁只讀 Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 每股價值（主值，依自用 AI 讀法）推到目標時點＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "分部加總部分反解「情境 B 分部加總每股價值＝現價」。反解值超出合理範圍時照實列出，不外推。營業利益率與營收規模兩項只動 FY2028 起的財報口徑營業利益（FY2026 Q4 與 FY2027 為錨點不動；折舊、資本支出、營運資金、自用 AI 影子效益不變）；"
                "現價低於模型值時，「現價隱含」是把模型往下調到現價所需的幅度（倍數小於 1、折現率高於模型）")
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
    kv("sh", "股數（十億股；評價日流通股＋RSU 淨股數＋遠期出售＋強制轉換特別股）", f"={shv}", "0.000")
    kv("oth", "淨現金、評價日後股權融資減協議付款、類債與特別股股利現值（US$B）", f"={nc}+{post}-{agtv}-{dl}-{pf}", "#,##0.0", "同 Val 頁 DCF 每股價值公式的非 EV 項（協議付款現值隨 WACC 變動，二分法內另算）")
    kv("evr", "支持現價所需的企業價值 EV（US$B；主值口徑）", f"=B{A['v0']}*B{A['sh']}-B{A['oth']}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；主值）", f"={ev}", "#,##0.0")
    kv("evx", "現價所需 EV 為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
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
    yr("tvf", "終值係數（只在 FY2032：價值驅動 (1＋g)(1 − g ÷ RONIC) ÷ (WACC − g)，或 tv_alt＝1 時 (1＋g) ÷ (WACC − g)，× 終值折現因子）",
       lambda vc, c: ("=IF(tv_alt=1,(1+g_tv)/(wacc-g_tv),(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))/(1+wacc)^(Val!$M$" + str(tr) + "+mid)" if vc == "M" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（Q4 26 與 FY2027 為錨點不動，記 0）", lambda vc, c: ("=0" if vc in ("F", "H") else f"=(1-tax)*({c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(B{Y['rev']}:H{Y['rev']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與 FY2028 起營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032，以同一係數加權）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "終值年（FY2032）權重大")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}<0,"低於 0：現價低於任何正營業利益率的價值",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收與營業利益等比例調整；營業利益率不變，折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("EB", "FY2028–FY2032 營業利益 × 貢獻係數合計（US$B；調整一倍所改變的 EV）", f"=SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=1+(B{A['evr']}-B{A['ev']})/B{A['EB']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 Val 頁；協議付款 FY2033–FY2036 現值隨 WACC 重算；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def ev_at(w):
        cf = f"Val!$F${ufcf}:$M${ufcf}+(internal_ai_reading=2)*Val!$F${shd}:$M${shd}"
        tvc = (f"IF(tv_alt=1,(Val!$M${ufcf}-Val!$M${agr}+Val!$M${dadf}*tax+(internal_ai_reading=2)*Val!$M${shd})*(1+g_tv)/({w}-g_tv),"
               f"((Val!$M${nor}-Val!$M${dadf}*(1-tax))*(1-g_tv/({w}+ronic_sp))+(internal_ai_reading=2)*Val!$M${shd}*(1-g_tv/({w}+ronic_sp)))*(1+g_tv)/({w}-g_tv))")
        return (f"SUMPRODUCT({cf},1/(1+{w})^Val!$F${tr}:$M${tr})"
                f"+{tvc}/(1+{w})^(Val!$M${tr}+mid)")

    def ag_at(w):
        return f"ag_rest*((1+{w})^-6.75+(1+{w})^-7.75+(1+{w})^-8.75+(1+{w})^-9.75)"
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, f"=({ev_at(f'D{r}')}+{nc}+{post}-{ag_at(f'D{r}')}-{dl}-{pf})/$B${A['sh']}")
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
    sev, sebt, snc, sag, sdl, spf, ssh, simp = (f"Val!$C${VR[k]}" for k in ("sev", "sebt", "snc", "sag", "sdl", "spref", "ssh", "simp"))
    stp = f"Val!$C${VR['sotp']}"

    def cmp(key, label, f_imp, f_mod, fmt):
        nonlocal r
        RV.cell(r, 1, label); RV.cell(r, 2, f_imp).number_format = fmt; RV.cell(r, 3, f_mod).number_format = fmt
        RV.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
        A[key] = r; r += 1
    cmp("sps", "情境 B 分部加總每股價值（美元；目標時點）", "=px", f"={stp}", "0.00")
    cmp("sev", "情境 B 分部加總 EV（US$B；只看分部加總、其他項不變）", f"=px*{ssh}-({snc}-{sag}-{sdl}-{spf})", f"={sev}", "#,##0.0")
    cmp("smul", "隱含 EV ÷ NTM EBITDA（倍）", f"=B{A['sev']}/{sebt}", f"={simp}", "0.0")
    cmp("scmp", "加權（DCF 不變）：支持現價所需的分部加總每股價值（情境加權；美元）", f"=(px-w_dcf*Val!$B${VR['tpdcf']})/(1-w_dcf)", f"=Val!$B${VR['tpcomp']}", "0.00")
    RV.cell(r, 1, "同業倍數：數位廣告（六家中位數）與採用倍數見 Val 頁 0.2 節；此處只列整體隱含倍數，不拆到分部"); r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：情境 B 的企業價值 "), f'TEXT(B{A["ev"]},"#,##0")', T(" 十億美元，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")', T(" 美元（評價日）；現價 "), 'TEXT(px,"0.00")',
                    T(" 美元相當於企業價值為模型的 "), f'TEXT(B{A["evx"]},"0.00")', T(" 倍。單靠一個因素達成時，需要：折現率 "), f'TEXT(B{A["w1"]},"0.0%")', T("（模型 "), f'TEXT(B{A["w0"]},"0.0%")',
                    T("）；或 FY2028–FY2032 營收與營業利益為模型的 "), f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                    T("）；或預測期營業利益率 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
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
    "IF_HoldEcon 四世代同時取低／高（讀法 2 影子效益的每 MW 值）": "四個 GPU 世代的每 MW 年持有成本同時取低／高（影子效益的每 MW 值）",
    "在建工程不另增 MW（cip_ai＝0，搭配 mw_src＝0 資本支出推得口徑；對照 mw_src 案）": "在建工程不另增 MW（改用資本支出推得的在役容量）",
    "協議應計以名目上限 12.775 計（含隱私案；legal_oth 少扣 0.885）": "協議應計以名目上限 12.775 計（含隱私案；其他法律應計少扣 0.885）",
}


def sens_tor(wb, C):
    SE = wb["Sens"]
    r0 = SE.max_row + 3
    SE.cell(r0, 1, "加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；本節只引用上方各案例的相對基準，只新增、不改既有格）").font = BOLD
    for j, h in enumerate(["輸入（白話）", "輸入取低值", "輸入取高值"], 1):
        c = SE.cell(r0 + 1, j, h); c.font = BOLD; c.fill = HDR
    r = r0 + 2
    pairs, order = {}, []
    for key, side, val, rr in C["SE_ROWS"]:
        lab_ = str(SE.cell(rr, 1).value)
        base = re.sub(r"：(低|高)$", "", lab_)
        if base not in pairs:
            pairs[base] = {}
            order.append(base)
        pairs[base][side] = rr
    for base in order:
        SE.cell(r, 1, TOR_PLAIN.get(base, base))
        for col, side in (("B", "低"), ("C", "高")):
            if side in pairs[base]:
                SE[f"{col}{r}"] = f"=Sens!E{pairs[base][side]}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調（讀取層級＋C0 實績對帳；MAG 類樣板）
def coord(wb, C):
    TK, TKV, tk, QR = C["TK"], C["TKV"], C["tk"], C["QR"]
    WN_ = C["WAT"]["names"]
    W = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_c0_v0.3.2.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    WV_RD = f"Wattonomics {C['WAT']['source']['version']}（{C['WAT']['source']['file']}；快照 wattonomics/data/archive/{C['WAT_FN']}）"
    TKS = f"Tokenomics {TKV}（{TK['source']['file']}，master {TK['source']['commit'][:7]}；快照 meta/data/tokenomics_snapshot_{TKV}.json）"
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJKL", (30, 12, 16, 16, 12, 28, 60, 22, 60, 22, 60)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.2.1 起；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。Meta 是算力買方（自用 AI 的資本支出方；不對外售算力），不做兩讀數目標價；"
                "C0 的個股側只用公司已公布實績（新聞稿／10-Q 財報表），層級側用 Wattonomics v0.3.2 隱含值（玩家資本支出、上線 GW），同源者標「參考」。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    for j, h in enumerate(["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1a = r
    ws.cell(r, 1, f"讀取層級：Tokenomics——本模型讀 {TKV}（與工作單所列現行版相同）；Wattonomics——本模型讀 {C['WAT']['source']['version']} 快照的 Meta 上線 GW（In 頁 wif_*，供自用在役容量），現行為 v0.3.2（20261114；本輪不換快照：Meta 上線 GW 兩版 2026 年底前相同、2027–2032 相差 ≤ 6e-8 GW）；"
                      "Wattonomics v0.3.2 內含的 Tokenomics 為 v21.10（wattonomics/data/tokenomics_snapshot_v21.10.json），與本模型的 v22.0 不是同一版（層級值之間的版本差異只列參考）。下方 C0 與追公式用 v0.3.2。")
    ws.cell(r, 1).font = BOLD; r += 1

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
    tkrows = []
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_HoldEcon｜{code}", code, tk("IF_HoldEcon", code), f"=he_{nm}", "US$m/MW-年（＝$B／GW／年）", "每 MW 年持有成本（經濟口徑；讀法 2 影子效益的每 MW 值，k＝1）"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_OpexGW｜{code}", code, tk("IF_OpexGW", code), f"=opx_{nm}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_CapexIT｜{code}", code, tk("IF_CapexIT", code), f"=cit_{nm}", "US$m/MW", "每 MW IT 資本支出"))
    tkrows.append(("IF_CapexFacility｜VR200", "VR200", tk("IF_CapexFacility", "VR200"), "=cfac", "US$m/MW", "每 MW 機房資本支出"))
    for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3")):
        tkrows.append((f"IF_DeprIT｜{code}", code, tk("IF_DeprIT", code), f"=dpi_{nm}", "US$m/MW-年", "每 MW IT 年折舊（期初機隊汰換）"))
    for nm_, yr_, lv, ad, un, nt in tkrows:
        f1(nm_, yr_, lv, ad, un, TKS, nt, "0.000")
    F1b_tk = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{F1a + 1}:D{F1b_tk}-C{F1a + 1}:C{F1b_tk}))").number_format = "0.000000"; F1chk = r; r += 1
    wif_first = r
    for y in range(2026, 2033):
        f1(f"WIF_PlayerGW_YE_meta_{y}", str(y), WN_[f"WIF_PlayerGW_YE_meta_{y}"], f"=wif_{y}", "GW（IT）", WV_RD,
           "Meta 年底上線 IT GW（In 頁 wif_*；自用在役容量，Thesis 讀入；上游權威，個股不改值）；v0.3.2 同名值：" + f"{WN[f'WIF_PlayerGW_YE_meta_{y}']:.9f}", "0.000000")
    wif_last = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Wattonomics 讀入列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{wif_first}:D{wif_last}-C{wif_first}:C{wif_last}))").number_format = "0.000000"; F1chk_w = r; r += 1
    ws.cell(r, 1, "Wattonomics v0.3.2 快照名稱（只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (("WIF_Gap_Player_meta_2026_CapexAI", "capai", "US$B", "Meta 2026 AI 資料中心資本支出（Interface!F635←Gap!G200←G183＝2026 新增 GW × Tokenomics 每 GW 成本＋次年、後年廠房提前；GW 由 CapexGW 從公司資本支出〔In!K187 gcap_meta〕反推，見下方追公式）"),
                             ("WIF_Gap_Player_meta_2026_CapexNonAI", "capnai", "US$B", "非 AI 資本支出（Interface!F636←Gap!G190＝CapexGW!G111 ×（1 − ai_sh）＝140 × 0.15；In!K186 gnai_meta 2026 空白，Assumed）"),
                             ("WIF_Gap_Player_meta_2026_AddGW", "addgw", "GW", "2026 年新增 IT GW（Interface!F643←Gap!G176←CapexGW!G127←G125←G124←…←G104←In!K187；輸入為公司資本支出〔含個股模型寫回與 2024–25 實績〕，見追公式；不是園區證據推得）"),
                             ("WIF_PlayerGW_YE_meta_2025", "ye25", "GW", "Meta 2025 年底上線 IT GW（Interface!E27←CapexGW!F128←Players!F87←F81←F80＝chip_meta 1,338.7 MW × k_chip2it 1.85；Epoch 晶片功耗 × 換算係數，Analogy；與公司資本支出無關，不同源）"),
                             ("WIF_PlayerGW_YE_meta_2026", "ye26", "GW", "Meta 2026 年底上線 IT GW（Interface!F27←CapexGW!G128＝2025 年底（Epoch，不同源）＋2026 新增 GW（同源於公司資本支出，見追公式））"),
                             ("WIF_PlayerGW_YE_meta_SiteOnly_2026", "ye26s", "GW", "2026 年底上線 IT GW（只有園區證據路徑；Prometheus 等園區列多為公司公告；不在主路徑）"),
                             ("WIF_Gap_Player_meta_2027_CapexAIModel", "cap27", "US$B", "並列讀法：2027 AI 資本支出（個股模型值回寫；2026 無此值）")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt, "0.000")
    gc = WR["gcap_meta"]
    wrows["guid"] = f1("In 列 gcap_meta（K187）｜2026（資本支出，財年）", "2026", gc["values"][3], None, "US$B", WV,
                       "Wattonomics In!K187（gcap_meta：Meta 個股模型基準寫回，FY2026E＝9M 實績＋Q4 公式＝指引中點 140，含融資租賃本金；Interested-party，指引不是實績）；是 CapexGW 2026 年新增 GW 與非 AI 資本支出的輸入；J187＝2025 實績 72.215、I187＝2024 實績 39.225、L187／M187＝個股模型 FY2027／FY2028（205.706／219.283）", "0.0")
    wrows["act"] = f1("In 列 act26_meta｜2026 已公布資本支出（Q1–Q3）", "2026", WR["act26_meta"]["values"][1], None, "US$B", WV,
                      "Wattonomics In!act26_meta（9M 實績 85.824，含融資租賃本金；Verified）；只被 Compare 頁讀（權重），不進 GW 或資本支出路徑（見追公式）", "0.000")
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 Meta 路徑是否讀公司實績或目標（v0.3.2 xlsx；規則：有讀者的部分屬同源，不當獨立證據）").font = BOLD; r += 1
    for j, h in enumerate(["輸出名稱 ← 公式鏈", "", "", "", "", "", "同源判定", "讀到什麼（輸入、公式、儲存格）"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    TR = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_trace_v0.3.2.json"), encoding="utf-8"))
    campus = TR["campus_rows"]
    TRC = [
        ("WIF_Gap_Player_meta_2026_AddGW ← Interface!F643 ← Gap!G176 ← CapexGW!G127 ← G125（上線時差加權；lag_on、inv_on）← G124 ← G123 ← G122 ← G114（＝G111 × ai_sh）← G111 ← G110 ← G106 ← G104 ← In!K187（gcap_meta）；前兩年 G125 另讀 F124 ← … ← In!J187、I187",
         "同源（公司資本支出指引＋公司 2024–25 實績＋個股模型寫回）",
         "In!K187 gcap_meta＝Meta 個股模型基準寫回（FY2026E＝9M 實績＋Q4 公式＝2026 指引中點 140）；In!J187＝2025 實績 72.215、In!I187＝2024 實績 39.225；ai_sh（In!I13，0.85）為 Assumed、不同源（人工確認）。"
         "GW＝資金 ÷ 每 GW 成本（Tokenomics），所以 2026 年新增 1.79 GW 是 Meta 資本支出換算，不是獨立的物理證據。備援輸入 In!K124 cap_meta（同為 140）為「分支未生效」，園區列（Pipeline，" + str(len(campus)) + " 列命中 Meta，其輸入列多為公司公告）在 2026 路徑上全部為「分支未生效」，不傳值"),
        ("WIF_Gap_Player_meta_2026_CapexAI ← Interface!F635 ← Gap!G200 ← G183（＝G180＋G181＋G182）", "同源（同上）",
         "G180＝2026 新增 GW（上列）×（每 GW IT 成本＋廠房）；G181、G182＝2027、2028 新增 GW 的廠房提前入帳，讀 CapexGW!H127、I127，其輸入為 In!L187、M187（Meta 個股模型 FY2027／FY2028 資本支出）；"
         "G194（個股模型 AI 資本支出 CapexAIModel）2026 空白，不替代"),
        ("WIF_Gap_Player_meta_2026_CapexNonAI ← Interface!F636 ← Gap!G190", "同源（公司指引）",
         "G188（In!K186 gnai_meta 個股模型非 AI 資本支出）2026 空白，改用 CapexGW!G111 ×（1 − ai_sh）＝140 × 0.15＝21；140 即公司指引中點（個股模型寫回）"),
        ("WIF_PlayerGW_YE_meta_2025 ← Interface!E27 ← CapexGW!F128 ← Players!F87 ← F81 ← F80 ← In!J101 chip_meta × Players!D17（＝k_chip2it）", "不同源（Epoch 晶片功耗 × 換算係數）",
         "2025 年底 2.477 GW＝chip_meta 1,338.7 MW（Epoch，Analogy）× k_chip2it 1.85；Players!F79（mw_meta，In!J108）空白為條件讀取；與公司資本支出無關，也不讀公司實績"),
        ("WIF_PlayerGW_YE_meta_2026 ← Interface!F27 ← CapexGW!G128＝F128＋G127", "部分同源",
         "2025 年底（Epoch 推得，不同源）＋2026 年增量 G127（同源於公司資本支出，見第一列）。層級另有 SiteOnly 路徑 WIF_PlayerGW_YE_meta_SiteOnly_2026（園區證據，約 3.15 GW，不在主路徑），但 Meta 不揭露 GW，仍不能做 C0"),
        ("act26_meta（Wattonomics In 頁 act26_meta；9M 資本支出實績 85.824）", "未進 GW／資本支出路徑",
         "反向查詢：公式讀 act26_meta 的只有 " + "、".join(sorted({x.split("!")[0] for x in TR["act26_readers"]})) + " 頁（" + str(len(TR["act26_readers"])) + " 格；Compare 頁第 239 列），不回饋 CapexGW、Gap、Interface 的 Meta 值；本表直接以公司新聞稿與 10-Q 的同一實績為個股側，不經 Wattonomics"),
        ("WIF_Gap_Player_meta_2026_CapexAIModel ← Gap!G194", "空白（無回寫值）",
         "2026 年個股模型 AI 資本支出 In!K185 gcapai_meta 為空，所以這一欄不替代 GW 推算；2027 年起的並列讀法才讀個股模型（本表不用）"),
    ]
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 2026 年 Meta 的資本支出（AI＋非 AI）與新增 GW 由公司 2026 資本支出指引中點（In!K187＝個股模型基準寫回）、公司 2024–25 實績（上線時差與廠房提前入帳）與個股模型 2027–28 資本支出換算，整列同源；2025 年底 GW 來自 Epoch 晶片功耗（不同源）；"
                 "公司實績 act26_meta 未進路徑。層級有 SiteOnly 路徑（約 3.15 GW，園區證據），但 Meta 不揭露 GW，每 GW 資本支出也不可比；因此資本支出列標「參考（同源校準）」，結論「無可比實績（層級值同源）」。"
                 "已知限制須人工確認：追公式工具對 In 頁鍵名的同源旗標只標「路徑上讀到公司輸入」，園區列（Prometheus、Hyperion 等）多為公司公告，標「園區輸入列（人工確認）」；本次 2026 路徑上園區列全部為分支未生效。若 Wattonomics 改為只用 Epoch／園區證據推 2026 GW，才恢復比較（證據回饋走對帳單 C5）。")
    rtr_conc = r; r += 1
    ws.cell(r, 1, "追公式傾印（由 tools/coord/trace_wattonomics.py --player meta 產生，貼程式輸出；只貼狀態為「生效」的格與旗標格；來源 " + TR["source"].split("/")[-1] + "；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 7, "標記"); ws.cell(r, 8, "公式（省略開頭等號）"); ws.cell(r, 4, "值"); r += 1
    rdump0 = r
    for x in TR["rows"]:
        if not (x["status"] == "生效" or x["flag"]):
            continue
        ws.cell(r, 1, ("　" * x["depth"]) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 7, (x["status"] + ("；" + x["flag"] if x["flag"] else "")))
        ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    r += 1
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝新聞稿／10-Q 財報表實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    for j, h in enumerate(["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級", "出處網址"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    F2hh = r; r += 1
    SRC_Q = ("Meta 2026Q1／Q2／Q3 財報新聞稿現金流量表「購置不動產與設備」加「融資租賃本金支付」＝公司資本支出口徑（含融資租賃本金；Meta 2026Q3 新聞稿 CFO outlook 同口徑）；2026Q3 Form 10-Q 現金流量表同值（網址見本列 L 欄；"
             "mag/extract/docs/2026Q3/text/meta_2026Q3_10-Q.txt）；事實總帳 meta/data/meta_facts_2026q3.json cf.capexInclFL；Q 頁第 " + str(QR["capex"]) + " 列")
    GR_Q = "Verified（財報新聞稿與 10-Q 財報表，兩讀核對）"
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
    ws.cell(r, 1, "核對：9M 累計 − Wattonomics In 頁 act26_meta（同口徑實績；應為 0）"); ws.cell(r, 4, f"=D{rq9}-C{wrows['act']}").number_format = "0.000000"; ract = r; r += 1
    rl = inp("層級 CY2026 玩家資本支出：AI＋非 AI（Wattonomics）", "CY2026（全年）", "US$B", None, "—", "—",
             f"=C{wrows['capai']}+C{wrows['capnai']}", f"{WV}：Gap 頁 Meta 2026 AI 資本支出（新增 {WN['WIF_Gap_Player_meta_2026_AddGW']:.2f} GW × Tokenomics 每 GW 成本，含廠房提前）＋非 AI 資本支出（連 F1 兩列）；兩者的輸入都是公司 2026 財年資本支出（指引中點 140，Meta 模型寫回 In!K187），屬同源（見上方追公式）", "Derived；輸入同源於公司資本支出")
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
    ws.cell(r, 2, "層級值同源於公司資本支出（見追公式）：參考（同源校準），不判定")
    ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由公司資本支出（指引中點、2024–25 實績、個股模型 2027–28）換算＝同源，不是獨立證據；上方兩個讀數仍列出供參考"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL},SIGN(F{rd1})=SIGN(F{rd2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rjr = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：公司 2026 資本支出指引中點 vs 層級 CY2026（層級的輸入就是指引中點；差距來自 GW 換算與廠房提前入帳的時點）"); ws.cell(r, 2, "指引中點（公司指引不是實績）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, "=capex_g").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（指引，不判定）")
    ws.cell(r, 8, "2026Q3 財報新聞稿 CFO outlook：2026 資本支出（含融資租賃本金）1,350–1,450 億美元（中點；In 頁 capex_g；Interested-party）"); ws.cell(r, 9, "Interested-party（指引）"); ws.cell(r, 10, "同上"); ws.cell(r, 11, "Derived")
    rg = r; r += 1
    ws.cell(r, 1, "AI 營收或 run-rate（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "Meta 不對外出售算力，財報不分列 AI 營收；AI 效益嵌在廣告營收內（法說與 follow-up call 未量化 AI 對曝光、價格的歸因；Recon 6d「未找到」）。本模型自用 AI 影子效益是讀法 2 的假設（Assumed），不是公司實績；缺的揭露：AI 歸因的廣告營收或利潤")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "Tokenomics 每 GW 營收錨 × Wattonomics 上線 GW 只能得層級隱含的自用經濟價值，無對應實績可比；不列數字"); ws.cell(r, 11, "—"); rai = r; r += 1
    ws.cell(r, 1, "在役／簽約電力 MW（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "Meta 不揭露自家在役或簽約 AI 電力（MW／GW；2026Q3 10-Q 全文無 MW／GW 字樣，已查）；Prometheus、Hyperion 等為公司公告的園區規模與時程（前瞻，不是在役實績）；Epoch AI 晶片功耗與園區估計是第三方估計，不算公司實績")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "Wattonomics 上線 GW：2025 年底 2.477、2026 年底 4.264（見 F1）；2025 年底來自 Epoch，2026 年增量同源於公司資本支出；層級值只供參考"); ws.cell(r, 11, "Derived"); rmw = r; r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "時差邊界（參考）項目數"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"時差邊界（參考）")'); rnb = r; r += 1
    ws.cell(r, 1, "無可比實績列數（資本支出 1、AI 營收 1、在役 MW 1）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "讀取層級（同源，不判定）：Recon 第 7 節「自用在役容量 vs Wattonomics」的差距判定"); ws.cell(r, 4, "參考，不計入對帳中")
    ws.cell(r, 8, "本模型自用在役容量讀 Wattonomics 上線時程（mw_src＝1，差距恆為 0）；Recon 第 7 節對照的「支出推得」與 Wattonomics 都由同一條 Meta 資本支出路徑換算（同源），v0.2.1 起該節判定欄改標「參考（讀取層級，不判定）」，年數不計入對帳中項目數"); rsrc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F(T("C0（實績）對帳：公司已公布資本支出（2026 年 Q1–Q3 累計 "), f'TEXT(D{rd1},"0.0")', T(" 十億美元）與層級 CY2026 隱含值（"), f'TEXT(E{rl},"0.0")', T("；同期間換算 "), f'TEXT(E{rd1},"0.0")', T("）相差 "),
                    f'TEXT(F{rd1},"+0%;-0%")', T("（主讀數）；最近一季年化 "), f'TEXT(D{rd2},"0.0")', T(" 十億美元，相差 "), f'TEXT(F{rd2},"+0%;-0%")',
                    T("（第二讀數）。但 Wattonomics 的 Meta 2026 年資本支出與新增 GW 由公司資本支出（指引中點、2024–25 實績與個股模型 2027–28；追公式：CapexGW!G104←In!K187）換算，屬同源，判定："), f'G{rj}',
                    T("（參考，不計入對帳中）。AI 營收與在役電力沒有可比的公司實績（公司不揭露），不判定。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向（依季中點）")
    ws.cell(r, 2, F(T("資本支出是流量，逐季上升（Q1 "), f'TEXT(D{rq1},"0.0")', T("、Q2 "), f'TEXT(D{rq2},"0.0")', T("、Q3 "), f'TEXT(D{rq3},"0.0")',
                     T("）。主讀數用 9 個月平均速率（中點約 5 月中）代表全年，全年中點在 7 月初，所以在上升趨勢下主讀數偏低；第二讀數用 Q3 單季年化（中點約 8 月中），比全年平均速率的中點晚一個半月，所以偏高。"
                       "兩讀數相差 "), f'TEXT((F{rd2}-F{rd1})*100,"0")', T(" 個百分點，就是時點差的幅度。公司全年實績要等 2026Q4 財報（2027-01 下旬）；層級 CY2026 與公司全年指引中點（140）相差 "), f'TEXT(F{rg},"+0%;-0%")',
                     T("，來自 GW 換算與廠房提前入帳的時點。這裡全部用流量（資本支出）比較，不拿 9 月底存量對年底存量。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出列：追公式（F1b）顯示 Wattonomics 2026 年 Meta AI 資本支出（Gap!G200）與非 AI 資本支出（Gap!G190）的共同輸入是公司 2026 財年資本支出 140（CapexGW!G104←In!K187，Meta 個股模型基準寫回，＝指引中點）；"
                   "新增 GW 也由此換算（CapexGW!G127），2027–28 的廠房提前入帳讀個股模型 FY2027／FY2028 資本支出（In!L187、M187）。層級值因此同源於公司資本支出，不是獨立證據：整列標「參考（同源校準）」，結論「無可比實績（層級值同源）」；每 GW 資本支出同樣無可比（公司不揭露 GW）。"
                   "公司實績 act26_meta（85.824，與本表 9M 累計相同）沒有進 Meta 的 GW 或資本支出路徑（只在 Compare 頁作權重）。"
                   "（二）AI 營收列與在役 MW 列：公司不揭露，列「無可比實績」。（三）所以本頁 C0 在 Meta 無檢驗力：對帳中為 0 不代表兩層一致，只代表層級值是由公司資本支出換算而來。要恢復檢驗，需 Wattonomics 改以 Epoch 或園區證據獨立推 2026 年 GW，或公司揭露 AI 營收／在役電力。"))
    rsame = r; r += 1
    ws.cell(r, 1, "口徑換算（層級 vs 公司資本支出）")
    ws.cell(r, 2, ("公司資本支出＝購置不動產與設備＋融資租賃本金支付（公司口徑；Meta 2026Q3 新聞稿 CFO outlook 同口徑，Wattonomics In!act26_meta 同口徑）；層級資本支出＝新增 GW × 每 GW 建置成本（IT＋廠房）＋非 AI 資本支出。兩者都是現金／完工口徑近似；"
                   "Meta 融資租賃本金 9M 為 " + "2.769" + "（占資本支出約 3%，已含在公司口徑內，不另調整）；Hyperion、El Paso 為合資（資本支出不在 Meta 帳上），Wattonomics 對該類園區以 Meta 使用者口徑列入上線 GW，但園區列在 2026 主路徑為分支未生效；未起租租賃 355.7（10-Q）屬表外承諾，不是實績資本支出。"))
    rk = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 輸入的 2026 資本支出（指引中點，In!K187）vs 其 Gap 頁 GW 推算的玩家資本支出（AI＋非 AI）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['guid']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics 資本支出輸入列 gcap_meta（公司指引中點）"); ws.cell(r, 10, "Wattonomics Gap 頁玩家資本支出（GW 推算）"); rb_ = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：2026 資本支出指引（上列）、市場共識目標價與年度共識（Recon 第 1 節）不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：Tokenomics 本模型讀 " + TKV + "；Wattonomics 本模型讀 " + C["WAT"]["source"]["version"] + " 快照（Meta 上線 GW），現行 v0.3.2 內含 Tokenomics v21.10，與本模型的 v22.0 不是同一版。Meta 模型值被 Wattonomics 讀入（In!gcap_meta 等寫回，v0.1.1 基準），本表只列層級值與追公式，不用該回饋判定。")
    r += 1
    return dict(F1=(F1a, F1b_tk), F1w=(wif_first, wif_last), F1chk=F1chk, F1chk_w=F1chk_w, F2h=F2h, F2hh=F2hh, F2=(F2a, F2b), rq=(rq1, rq2, rq3, rq9), rl=rl, rd1=rd1, rd2=rd2, rj=rj, rjr=rjr, rtr=rtr_conc, rdump=(rdump0, rdump1), rg=rg,
                rai=rai, rmw=rmw, rn=rn, rnb=rnb, rnc=rnc, rsrc=rsrc, rs=rs, rr=rr_, rsame=rsame, rk=rk, rb=rb_, wrows=wrows, tol=TOL, ract=ract)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, RVA, PLA, CRD):
    RD = wb["README"]
    VR, PRB, PRC = C["VR"], C["PR"]["B"], C["PR"]["C"]
    TB = C["T"]["B"]
    vb = lambda k, c="B": f"Val!{c}{VR[k]}"
    tp, up, rec, dcf, cmpv, cons = vb("tp"), vb("up"), vb("rec"), vb("tpdcf"), vb("tpcomp"), vb("cons")
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    ox = PLA["idx"]
    RC6 = C["RC6"]
    kk = lambda key: f'IF(ISNUMBER(Recon!B{RC6[key]}),TEXT(Recon!B{RC6[key]},"0.00"),"未重算")'
    out = {}
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.1 起，個股 HTML 版面規格 v1.3：橫幅說明、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), f'TEXT(ABS({up}),"0%")', f'IF({up}<0," 低，"," 高，")', T("評等"), rec, T("；營業利益率由 "), f'TEXT(PL!D{PRB["om"]},"0%")', T("（FY2025）降到 "), f'TEXT(PL!I{PRB["om"]},"0%")', T("（FY2028）。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：自用 AI 採價值中性（讀法 2，k＝"), 'TEXT(k_ai,"0.00")', T("）；改為只計成本（讀法 1）時加權目標價 "), f'TEXT({vb("tp_r1")},"0")', T(" 美元（"), f'TEXT({vb("tp_r1", "C")},"+0%;-0%")', T("），評等"),
        f'{vb("tp_r1", "D")}', T("；現價 "), 'TEXT(px,"0.00")', T(" 美元為 2026-11-05 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價 "), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；情境 A／B／C 為 "), f'TEXT({vb("scen")},"0")', T("／"), f'TEXT({vb("scen", "C")},"0")', T("／"), f'TEXT({vb("scen", "D")},"0")', T("，市場共識 "), f'TEXT({cons},"0")', T("。"))
    put("o2", T("讀法 2 依賴度高：自用 AI 價值中性（k＝1）占目標價 "), f'TEXT({vb("dep_tot", "C")},"0%")', T("；只計成本（讀法 1）降到 "), f'TEXT({vb("tp_r1")},"0")', T("，支持現價所需 k＝"), kk("kreq"), T("、評等翻轉 k＝"), kk("kflip"),
        T("。反向推算（情境 B）：支持現價的折現率 "), f'TEXT({rv("w1")},"0.0%")', T("（模型 "), f'TEXT({rv("w0")},"0.0%")', T("）。"))
    put("o3", T("AI 折舊壓低利益率：基準情境新增 AI 折舊 FY2028 "), f'TEXT(PL!I{PRB["da_ai"]}*10,"#,##0")', T(" 億美元，毛利率由 FY2025 "), f'TEXT(PL!E{ox["gm"]},"0.0%")', T(" 降到 FY2028 "), f'TEXT(PL!H{ox["gm"]},"0.0%")',
        T("；FY2028 自用 AI 影子效益 "), f'TEXT(Thesis!I{TB["shadow"]},"0.0")', T("，對自用 AI 增量成本 "), f'TEXT(Thesis!I{TB["int_cost"]},"0.0")', T("（十億美元）。"))
    put("o4", T("收支關卡（FY2026–2030 自由現金流扣股利）A／B／C＝"), f'TEXT(CF!B{C["GATE"]},"0.0")', T("／"), f'TEXT(CF!C{C["GATE"]},"0.0")', T("／"), f'TEXT(CF!D{C["GATE"]},"0.0")',
        T("；基準情境單年最多新增舉債 "), f'TEXT(CF!C{C["GATE"] + 5},"0.0")', T("（"), f'CF!C{C["GATE"] + 6}', T("），融資可行性 WARN 年數 "), f'TEXT(CF!C{C["GATE"] + 14},"0")', T("。"))
    put("o5", T("實績對帳（本版新增）：2026 年 Q1–Q3 資本支出 "), f'TEXT(Coord!D{CRD["rd1"]},"0.0")', T(" 十億美元，與 Wattonomics 隱含的 CY2026 同期間值相差 "), f'TEXT(Coord!F{CRD["rd1"]},"+0%;-0%")',
        T("（最近一季年化 "), f'TEXT(Coord!F{CRD["rd2"]},"+0%;-0%")', T("），但該層級值由 Meta 資本支出換算（同源），結論："), f'Coord!G{CRD["rj"]}', T("；AI 營收、在役電力公司不揭露，無可比實績。"))
    put("drv", T("分部營收推導（情境 B）：Advertising＝廣告曝光數 × 每則廣告平均價格（各自近四季年增向長期值收斂；情境 A／C 只調價格），Other revenue 與 Reality Labs 各依近四季年增收斂；"
                 "三者加總為合併營收。FY2028 Advertising "), f'TEXT(PL!I{PRB["ads"]},"0")', T("、Other revenue "), f'TEXT(PL!I{PRB["other"]},"0.0")', T("、Reality Labs "), f'TEXT(PL!I{PRB["rl"]},"0.0")', T("，合計 "), f'TEXT(PL!I{PRB["rev"]},"0")', T(" 十億美元。"))
    RD.cell(r, 1, F(T("v0.2.2（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 605.08（11-05）→ "), 'TEXT(px,"0.00")',
                    T("；股價基準日 2026-11-05 → 2026-11-13，目標時點隨公式 2027-11-05 → "), 'TEXT(tp_date,"yyyy-mm-dd")',
                    T("；加權目標價 778.44 → "), f'TEXT({vb("tp")},"0.00")', T("；上檔空間 +28.7% → "), f'TEXT({vb("up")},"+0.0%;-0.0%")',
                    T("；投資建議 買進 → "), f'{vb("rec")}', T("（只換現價時為 780.65、+21.7%、中立）；加權目標價的變動來自同業倍數換價。評價日股價與 WACC 權益權重（px_val）不隨現價換。同業倍數（報價日 2026-10-07／10-08；Meta 自身 10-28）依價格檔同日收盤換到 2026-11-13（GOOGL 與 Meta 自身用市值法，PINS、TTD、APP、RDDT、SNAP 倍數 × 收盤比，近似）：" + _PXTXT + "；共識與其他假設不變。")))
    out["ver"] = r; r += 1
    RD.cell(r, 1, "v0.2.1（2026-11-15）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁、PL 頁末標準損益表、Sens 頁末加權目標價相對基準、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」。既有數值格不變，目標價與評等不變；Wattonomics 固定讀 v0.3 快照，本輪不換快照")
    r += 1
    return out


_PXJ = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "peers_ads_20261113.json"), encoding="utf-8"))
_PXG = [p for p in _PXJ["peers"] if p["ticker"] == "GOOGL"][0]
_PXTXT = f"六家中位數 {_PXJ['median1030']['value']} → {_PXJ['median']['value']}、GOOGL {_PXG['ntmEvEbitda1007']} → {_PXG['ntmEvEbitda']}、Meta 自身 {_PXJ['selfReference']['ntmEvEbitda1007']} → {_PXJ['selfReference']['ntmEvEbitda']}"


def apply(wb, C):
    PLA = pl_standard(wb, C)
    RVA = revdcf(wb, C)
    tor = sens_tor(wb, C)
    CRD = coord(wb, C)
    RMA = readme13(wb, C, RVA, PLA, CRD)
    return dict(PL=PLA, RV=RVA, tor=tor, CRD=CRD, RM=RMA)
