# SpaceX v0.2.8（2026-11-15）後處理（v0.2.9 起 Coord 改讀 Wattonomics v0.3.3）：個股 HTML 版面規格 v1.3 全套＋Coord 頁「三層協調」節（MAG 類樣板；Alphabet v0.3.9 為底）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點。
# 只新增，不改任何既有儲存格（既有數值、目標價、評等不變；逐格比對只允許版號文字、build_id 與新增格）：
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B；GAAP；9M26A 比較欄）、比率區、核對列、424B4 歷史輸入、分部營收（圖表用）
#   RevDCF 頁（新）：反向 DCF（情境 B；封閉解與二分法）＋分部加總隱含倍數；只讀其他頁，無任何頁引用本頁
#   Sens 頁末：加權目標價相對基準（龍捲風圖用，白話標籤）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v21.10；Wattonomics v0.3.3 xai 玩家只作對帳；v0.2.9 由 v0.3.2 換讀）、F1b 追公式、F2 C0（實績）
#   README 頁末：v1.3 展示層白話列
# build_spacex.py 在 README 之後、存檔前呼叫 apply(wb, ctx)。
import json
import os
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.9", "2026-11-15"   # v0.2.9：只換 Coord 讀的 Wattonomics 快照（v0.3.2 → v0.3.3）
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")


def lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


def T(s):
    return ("t", s)


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


SRC_424 = "SpaceX 424B4 最終招股書（mag/extract/docs/2026Q3/spacex_424B4.pdf）MD&A 合併損益表（FY2025 vs FY2024 與 FY2024 vs FY2023 兩表；Verified）"
# =====================================================================================================
# PL：標準損益表
# 欄位：C FY2023A、D FY2024A、E FY2025A、F–L FY2026E–FY2032E、M 9M26A、N 說明
S2P = {"E": "C", "F": "G", "G": "H", "H": "I", "I": "J", "J": "K", "K": "L", "L": "M", "M": "E"}   # 標準表欄 → PL 頁欄（FY2025A→C、9M26A→E）
ALLC = "CDEFGHIJKLM"
PCOLS = "EFGHIJKLM"       # 有 PL 頁對應欄者


def pl_standard(wb, C):
    P, Q, B = wb["PL"], wb["Q"], C["PR"]["B"]
    QR = C["QR"]
    r = P.max_row + 3
    L = {}
    P.cell(r, 1, "標準損益表（v0.2.8 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績，424B4）＋FY2026E–FY2032E（情境 B 基準）；最右欄 9M26A 為最新累計實績。US$B；EPS 美元。只引用本頁上方、Q 頁與下方歷史輸入，不回饋任何計算").font = BOLD
    L["start"] = r; r += 1
    P.cell(r, 1, "口徑：GAAP（財報口徑營業利益，不含自用 AI 影子效益）。歷史年度的稀釋股數與 EPS 為 IPO 前資本結構（普通股加權股數，FY2024 含特別股轉換因子），與 FY2026E 起（含強制轉換特別股、IPO 後股數）不可直接比較；FY2024 淨利為正是因為稅務利益 0.549 與其他收益 0.985。")
    L["note"] = r; r += 1
    P.cell(r, 1, "營業成本與營業費用：模型以分部調整後 EBITDA、折舊與股權報酬推導營業利益，沒有分列營業成本與研發、行銷管理。預測期「成本費用合計＝營收 − 營業利益」，營業成本＝合計 × 9M26A 營業成本占合計之比（下方假設列；Assumed），營業費用（研發＋行銷與管理＋重組＋減損）＝合計 − 營業成本。"
                 "營業利益、稅前、稅後與 EPS 全取自 PL 頁情境 B，不因此改變；歷史年度與 9M26A 為財報實績（424B4、2026Q3 10-Q）。FY2026E 其他收入（淨額）＝PL 頁稅前利益 − 營業利益（含 9M26A 實績利息與其他收益）。")
    L["note2"] = r; r += 1
    hdr = ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "9M26A（累計）", "說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    L["hdr"] = r; r += 1
    M3 = "#,##0.000"; PCT = "0.0%"
    order = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    base = r
    idx = {k: base + i for i, k in enumerate(order)}
    r_ratio = base + len(order)
    ridx = {"r_niy": r_ratio + 1}
    r_chk = r_ratio + 2
    cidx = {k: r_chk + 1 + i for i, k in enumerate(["c1", "c2", "c3", "c4", "c5", "c6", "c7"])}
    r_seg = r_chk + 1 + 7 + 1
    sidx = {k: r_seg + 1 + i for i, k in enumerate(["s_sp", "s_cn", "s_si", "s_tot", "s_chk"])}
    r_h = r_seg + 1 + 5 + 1
    hkeys = ["h_rev", "h_cogs", "h_opx", "h_oi", "h_oth", "h_pt", "h_tax", "h_ni", "h_sh", "h_eps", "h_sp", "h_cn", "h_si", "h_c9m", "cogs_sh"]
    hidx = {k: r_h + 2 + i for i, k in enumerate(hkeys)}

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
    pr = lambda row_, cols=PCOLS: {c: f"=PL!{S2P[c]}{row_}" for c in cols}
    H = lambda k, c: f"={c}{hidx[k]}"      # 歷史輸入表：C=FY2023、D=FY2024、E=FY2025
    rv, oi, pt, tx, ni, she, eps = (B[k] for k in ("rev", "oi", "pt", "tx", "ni", "she", "eps"))
    row("rev", "營收", "US$B", {"C": H("h_rev", "C"), "D": H("h_rev", "D"), **pr(rv)}, M3, "分部營收加總（PL 頁情境 B 區塊）；FY2023A、FY2024A 為 424B4 實績")
    assert L["rev"] == idx["rev"]
    yo = {c: f"={c}{idx['rev']}/{p}{idx['rev']}-1" for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    yo["M"] = f"=M{idx['rev']}/PL!D{rv}-1"
    row("yoy", "營收年增率（YoY）", "%", yo, PCT, "FY2023A 的基期（FY2022）不在 424B4 損益表，空白；9M26A 對 9M25A（PL 頁）")
    cg = {"C": H("h_cogs", "C"), "D": H("h_cogs", "D"), "E": H("h_cogs", "E"), "M": f"=C{hidx['h_c9m']}",
          "F": f"=M{idx['cogs']}+((F{idx['rev']}-F{idx['oi']})-(M{idx['rev']}-M{idx['oi']}))*cogs_sh_r"}
    for c in "GHIJKL":
        cg[c] = f"=({c}{idx['rev']}-{c}{idx['oi']})*cogs_sh_r"
    row("cogs", "營業成本（COGS）", "US$B", cg, M3, "歷史與 9M26A＝財報；FY2026E＝9M26A＋Q4 26E（成本費用合計 × 占比）；FY2027E 起＝成本費用合計 × 占比（假設列）")
    row("gp", "毛利", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['cogs']}" for c in ALLC}, M3)
    row("gm", "毛利率", "%", {c: f"={c}{idx['gp']}/{c}{idx['rev']}" for c in ALLC}, PCT, "預測期毛利率隨營收組合與營業利益率變動（成本占比為常數的結果）")
    ox = {"C": H("h_opx", "C"), "D": H("h_opx", "D"), "E": H("h_opx", "E"), "M": f"=Q!H{QR['costs']}-M{idx['cogs']}"}
    for c in "FGHIJKL":
        ox[c] = f"=({c}{idx['rev']}-{c}{idx['oi']})-{c}{idx['cogs']}"
    row("opx", "營業費用（研發＋行銷與管理＋重組＋減損）", "US$B", ox, M3, "歷史與 9M26A＝財報（9M26A＝成本費用合計 26.378 − 營業成本）；預測期＝成本費用合計 − 營業成本")
    oc = {"C": H("h_oi", "C"), "D": H("h_oi", "D"), **pr(B["oi"])}
    row("oi", "營業利益", "US$B", oc, M3, "財報口徑合併營業利益（PL 頁；不含自用 AI 影子效益）")
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in ALLC}, PCT)
    ot = {"C": H("h_oth", "C"), "D": H("h_oth", "D"), "E": H("h_oth", "E"), **{c: f"=PL!{S2P[c]}{pt}-PL!{S2P[c]}{B['oi']}" for c in "FGHIJKLM"}}
    row("oth", "其他收入（淨額；利息收支與其他收益）", "US$B", ot, M3, "利息收入 − 利息費用＋其他收益（費用）；預測期＝PL 頁稅前利益 − 營業利益")
    row("pti", "稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}" for c in ALLC}, M3)
    row("tax", "所得稅（負值＝利益）", "US$B", {"C": H("h_tax", "C"), "D": H("h_tax", "D"), "E": H("h_tax", "E"), **pr(tx, "FGHIJKLM")}, M3)
    row("ni", "稅後淨利", "US$B", {c: f"={c}{idx['pti']}-{c}{idx['tax']}" for c in ALLC}, M3, "稅前 − 所得稅（核對列：與 424B4 與 PL 頁淨利相同）")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in ALLC}, PCT)
    row("sh", "稀釋加權股數", "十億股", {"C": H("h_sh", "C"), "D": H("h_sh", "D"), "E": H("h_sh", "E"), **{c: f"=PL!{S2P[c]}{B['sh']}" for c in "FGHIJKLM"}}, M3,
        "歷史＝424B4（IPO 前普通股；FY2024 含特別股轉換因子 9.956）；9M26A 為 IPO 前後加權（10-Q）；FY2026E 起為模型平均稀釋股數（含轉換後特別股）")
    row("eps", "EPS（GAAP 稀釋）", "US$", {"C": H("h_eps", "C"), "D": H("h_eps", "D"), "E": H("h_eps", "E"), **{c: f"=PL!{S2P[c]}{eps}" for c in "FGHIJKLM"}}, "0.00",
        "歷史＝424B4 公告值；FY2026E 起取 PL 頁（歸屬普通股淨利 ÷ 平均稀釋股數）；歷史與預測股數基礎不同，不宜直接比成長")
    assert L["eps"] == idx["eps"] and r == r_ratio
    P.cell(r, 1, "比率").font = BOLD; r += 1
    ny = {c: f'=IF(AND({p}{idx["ni"]}>0,{c}{idx["ni"]}>0),{c}{idx["ni"]}/{p}{idx["ni"]}-1,"NM")' for p, c in zip("CDEFGHIJK", "DEFGHIJKL")}
    ny["M"] = f'=IF(AND(PL!D{ni}>0,M{idx["ni"]}>0),M{idx["ni"]}/PL!D{ni}-1,"NM")'
    row("r_niy", "淨利年增率（YoY；基期或本期為負值＝NM）", "%", ny, PCT, "NM＝不具比較意義（基期或本期淨利為負）")
    assert r == r_chk
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    row("c1", "核對：毛利 − 營業費用 − 營業利益", "US$B", {c: f"={c}{idx['gp']}-{c}{idx['opx']}-{c}{idx['oi']}" for c in ALLC}, "0.000", "歷史：財報分項互相核對；預測期以殘差定義，恆為 0")
    row("c2", "核對：稅前利益 − PL 頁稅前利益", "US$B", {c: f"={c}{idx['pti']}-PL!{S2P[c]}{pt}" for c in "FGHIJKLM"}, "0.000")
    row("c3", "核對：稅後淨利 − PL 頁淨利", "US$B", {c: f"={c}{idx['ni']}-PL!{S2P[c]}{ni}" for c in "FGHIJKLM"}, "0.000")
    row("c4", "核對：營收 − PL 頁合併營收", "US$B", {c: f"={c}{idx['rev']}-PL!{S2P[c]}{rv}" for c in PCOLS}, "0.000")
    row("c5", "核對：歷史稅後淨利 − 424B4 淨利（輸入列）", "US$B", {c: f"={c}{idx['ni']}-{c}{hidx['h_ni']}" for c in "CDE"}, "0.000")
    row("c6", "核對：FY2025A 營收（PL 頁）− 424B4 營收（輸入列）", "US$B", {"E": f"=E{idx['rev']}-E{hidx['h_rev']}"}, "0.000")
    row("c7", "核對：FY2025A 營業利益（PL 頁）− 424B4 營業利益（輸入列）", "US$B", {"E": f"=E{idx['oi']}-E{hidx['h_oi']}"}, "0.000")
    assert r == r_seg - 1
    r += 1
    P.cell(r, 1, "分部營收（US$B；圖表與驅動因子章用）").font = BOLD; r += 1
    sg = lambda k, pk: {"C": H(k, "C"), "D": H(k, "D"), **pr(B[pk])}
    row("s_sp", "Space", "US$B", sg("h_sp", "space"), M3)
    row("s_cn", "Connectivity", "US$B", sg("h_cn", "conn"), M3)
    row("s_si", "SI（Super Intelligence；424B4 稱 AI）", "US$B", sg("h_si", "si"), M3)
    row("s_tot", "合計", "US$B", {c: f"=SUM({c}{sidx['s_sp']}:{c}{sidx['s_si']})" for c in ALLC}, M3)
    row("s_chk", "核對：合計 − 營收（應為 0）", "US$B", {c: f"={c}{sidx['s_tot']}-{c}{idx['rev']}" for c in ALLC}, "0.000")
    assert r == r_h - 1
    r += 1
    P.cell(r, 1, "歷史輸入與假設（藍字＝輸入；C＝FY2023、D＝FY2024、E＝FY2025；來源 424B4，標 Verified；US$B，EPS 美元）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "單位", "FY2023", "FY2024", "FY2025"] + [""] * 8 + ["來源"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    hv = [("h_rev", "營收", (10.387, 14.015, 18.674)), ("h_cogs", "營業成本", (6.110, 7.996, 9.451)),
          ("h_opx", "營業費用（研發＋行銷與管理＋重組＋減損）", (2.105 + 1.665 + 0.237 + 3.775, 3.464 + 1.813 + 0.213 + 0.063, 8.643 + 2.644 + 0.487 + 0.038)),
          ("h_oi", "營業利益（損失）", (-3.505, 0.466, -2.589)), ("h_oth", "其他收入（淨額）＝利息費用＋利息收入＋其他收益", (-1.693 + 0.249 - 0.042, -1.580 + 0.371 + 0.985, -1.945 + 0.492 - 0.177)),
          ("h_pt", "稅前利益", (-4.991, 0.242, -4.219)), ("h_tax", "所得稅（負值＝利益）", (-0.363, -0.549, 0.718)), ("h_ni", "淨利（損）", (-4.628, 0.791, -4.937)),
          ("h_sh", "稀釋加權股數（十億股）", (2.759, 9.956, 2.926)), ("h_eps", "EPS（稀釋，美元）", (-1.68, 0.00, -1.69)),
          ("h_sp", "Space 營收", (3.557, 3.796, 4.086)), ("h_cn", "Connectivity 營收", (3.869, 7.599, 11.387)),
          ("h_si", "SI 營收（FY2023 為合計 − Space − Connectivity，Derived；FY2024、FY2025 為 424B4 分部表）", (10.387 - 3.557 - 3.869, 2.620, 3.201))]
    for k, labl, vals in hv:
        assert hidx[k] == r, (k, hidx[k], r)
        P.cell(r, 1, labl); P.cell(r, 2, "US$" if "EPS" in labl else ("十億股" if "股數" in labl else "US$B"))
        for j, v in enumerate(vals):
            c = P.cell(r, 3 + j, round(v, 3)); c.number_format = "0.00" if "EPS" in labl else "#,##0.000"; c.font = BLUE
        P.cell(r, 14, SRC_424)
        r += 1
    P.cell(r, 1, "9M26 營業成本（2026Q3 10-Q 損益表；Verified）"); P.cell(r, 2, "US$B"); c = P.cell(r, 3, 11.361); c.font = BLUE; c.number_format = "#,##0.000"
    P.cell(r, 14, "SpaceX 2026Q3 10-Q 合併損益表 Nine Months Ended September 30, 2026（mag/extract/docs/2026Q3/spacex_2026Q3_10-Q.pdf）：Cost of revenue 11,361；R&D 11,658；SG&A 3,365；Restructuring (6)；合計 26,378")
    assert hidx["h_c9m"] == r; r += 1
    P.cell(r, 1, "營業成本占成本費用合計之比（9M26A；預測期營業成本假設）"); P.cell(r, 2, "%")
    P.cell(r, 3, f"=C{hidx['h_c9m']}/Q!H{QR['costs']}").number_format = "0.0%"
    P.cell(r, 14, "Derived／Assumed：模型未分列營業成本，預測期成本費用合計（營收 − 營業利益）按此比例分攤；只影響毛利與營業費用的分配，不影響營業利益、稅前、稅後與 EPS")
    assert hidx["cogs_sh"] == r; dn(wb, "cogs_sh_r", f"PL!$C${r}"); r += 1
    for col in "CDEFGHIJKLM":
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 12)
    L.update(idx=idx, ridx=ridx, cidx=cidx, sidx=sidx, hidx=hidx, ratio=r_ratio, chk=r_chk, seg=r_seg, h=r_h)
    return L


# =====================================================================================================
# RevDCF：反向 DCF（情境 B；讀 Val、PL、In 頁；無任何頁引用本頁）
def revdcf(wb, C):
    V = wb["Val"]
    VR, DC, PRB = C["VR"], C["DC"]["B"], C["PR"]["B"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    ev, nc, hold, dl, shv = (f"Val!$C${VR[k]}" for k in ("ev", "nc0", "hold", "dl", "shv"))
    cols = ["F", "H", "I", "J", "K", "L", "M"]      # Q4 26E、FY2027E–FY2032E
    tr, dfr, ebr, shd, ufcf, tvn, tvf, shdt = DC["t"], DC["df"], DC["ebit"], DC["shd"], DC["ufcf"], DC["tvn"], DC["tvf"], DC["shdt"]
    RV["A1"] = ("RevDCF — 反向 DCF：現價隱含什麼（v0.2.8 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式"
                "（DCF 對各年營業利益為線性，用封閉解；WACC 用二分法）。本頁只讀 Val、PL、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：以情境 B（基準）路徑反解「情境 B 的 DCF 每股價值推到目標時點＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數；本模型無股利；自用 AI 依 In 頁讀法）；"
                "分部加總部分反解「情境 B 分部加總每股價值＝現價」。反解值超出合理範圍時照實列出，不外推。營業利益率與營收規模兩項只動 FY2028 起的財報口徑營業利益（FY2026 Q4 與 FY2027 為錨點不動；折舊、資本支出、營運資金、自用 AI 影子效益不變）。"
                "WACC 反解的搜尋下界為永續成長率＋0.01 個百分點（WACC 趨近永續成長率時終值發散，每股價值趨於無限大，故區間內必有解）；隱含 WACC 若幾乎等於永續成長率，照實列出並註明不具經濟意義。")
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
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=px/(1+ke)^t_tgt", "0.00", "＝現價 ÷（1＋股權成本）^目標時點年數（Val 頁 DCF 每股價值推到目標時點的反算；本模型無股利）")
    kv("sh", "股數（十億股；評價日 A＋B＋RSU 與選擇權淨股數）", f"={shv}", "0.000")
    kv("oth", "淨現金、數位資產減類債（US$B）", f"={nc}+{hold}-{dl}", "#,##0.0", "同 Val 頁 DCF 每股價值公式的非 EV 項")
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
    yr("tvf", "終值係數（只在 FY2032：(1＋g)×(1 − g ÷ RONIC；終值用自由現金流時省略)÷(WACC − g)× 終值折現因子）",
       lambda vc, c: ("=(1+g_tv)*IF(tv_alt=1,1,1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^(Val!$M$" + str(tr) + "+mid)" if vc == "M" else "=0"), "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（Q4 26 與 FY2027 為錨點不動，記 0）",
       lambda vc, c: ("=0" if vc in ("F", "H") else f"=(1-tax)*{c}{Y['df']}" + (f"+(1-tax_tv)*{c}{Y['tvf']}" if vc == "M" else "")), "0.0000")
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
        tvm = f"IF(tv_alt=1,Val!$M${tvf}*(1+g_tv)/({w}-g_tv),Val!$M${tvn}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv))"
        tvs = f"(internal_ai_reading=2)*Val!$M${shdt}*(1+g_tv)*IF(tv_alt=1,1,1-g_tv/({w}+ronic_sp))/({w}-g_tv)"
        return f"SUMPRODUCT({cf},1/(1+{w})^Val!$F${tr}:$M${tr})+({tvm}+{tvs})/(1+{w})^(Val!$M${tr}+mid)"
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.0001"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, f"=({ev_at(f'D{r}')}+$B${A['oth']})/$B${A['sh']}")
        for c in (2, 3, 4):
            RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC（無解時為搜尋區間端點）", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.01 個百分點～40%（v0.2.8 獨立核對：原下界＋0.5 個百分點使解落在區間外，誤判為無解）")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_ok", "判讀：二分法收斂（隱含 WACC 下的每股價值 ≈ 支持現價所需值）", f'=IF(ABS(E{b0 + 39}-B{A["v0"]})<0.05,"有解："&TEXT(B{A["w1"]},"0.00%")&IF(B{A["w1"]}-g_tv<0.01,"，幾乎等於永續成長率 "&TEXT(g_tv,"0.0%")&"，不具經濟意義",""),"無解：隱含 WACC 低於搜尋下界 "&TEXT(g_tv+0.0001,"0.00%")&"（幾乎等於永續成長率 "&TEXT(g_tv,"0.0%")&"，不具經濟意義）")&"；模型 "&TEXT(wacc,"0.0%")', "@")
    kv("w_chk", "核對：以模型 WACC 代入同一公式的 EV − Val 頁情境 B EV（應為 0）", f"={ev_at('wacc')}-B{A['ev']}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的評價日每股價值 − 支持現價所需值（美元；無解時不為 0）", f"=E{b0 + 39}-B{A['v0']}", "0.000000")
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
    cmp("smul", "隱含 EV ÷ NTM 調整後 EBITDA（倍）", f"=B{A['sev']}/{sebt}", f"={simp}", "0.0")
    cmp("scmp", "加權（DCF 不變）：支持現價所需的分部加總每股價值（情境加權；美元）", f"=(px-w_dcf*Val!$B${VR['tpdcf']})/(1-w_dcf)", f"=Val!$B${VR['tpcomp']}", "0.00")
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：情境 B 的企業價值 "), f'TEXT(B{A["ev"]},"#,##0")', T(" 十億美元，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")', T(" 美元（評價日）；要支持現價 "), 'TEXT(px,"0.00")',
                    T(" 美元，企業價值須放大為 "), f'TEXT(B{A["evx"]},"0.0")', T(" 倍。單靠一個因素達成時，需要："),
                    f'IF(LEFT(B{A["w_ok"]},2)="有解","折現率降到 "&TEXT(B{A["w1"]},"0.00%"),"折現率低於 "&TEXT(g_tv+0.0001,"0.00%"))&IF(B{A["w1"]}-g_tv<0.01,"（幾乎等於永續成長率 "&TEXT(g_tv,"0%")&"，不具經濟意義；模型 "&TEXT(B{A["w0"]},"0.0%")&"）","（模型 "&TEXT(B{A["w0"]},"0.0%")&"）")',
                    T("；或 FY2028–FY2032 營收與營業利益全面放大 "), f'TEXT(B{A["s1"]},"0.0")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")',
                    T("）；或預測期營業利益率提高到 "), f'TEXT(B{A["m1"]},"0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0%")', T("，終值年權重大）"),
                    f'IF(B{A["m1"]}>1,"，超過 100%，不可能單靠利益率達成","")', T("。")))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("分部加總法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於情境 B 分部加總 EV ÷ NTM 調整後 EBITDA 約 "), f'TEXT(B{A["smul"]},"0.0")', T(" 倍，模型採用各分部同業倍數，整體隱含 "), f'TEXT(C{A["smul"]},"0.0")',
                    T(" 倍；換言之，只看分部加總，市場價格對應的整體倍數比模型 "), f'IF(B{A["smul"]}<C{A["smul"]},"低","高")', f'TEXT(ABS(D{A["smul"]}),"0%")', T("（EV 差距）。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 76
    for c in "BCDEFGH":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["b0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


# =====================================================================================================
# Sens：加權目標價相對基準（龍捲風圖用）
TOR_PLAIN = {
    "G2 校準＋k 不收斂（k0_ovr、conv_k）": "G2 校準後 k 不收斂（每 GW 年營收相對持有成本的倍數不下降）",
    "Tokenomics IF_HoldEcon 四世代同時": "Tokenomics 四個 GPU 世代的每 MW 年持有成本同時取低／高",
    "融資受限：FY2027 起建置速度減半（pace × 0.5；驗收 B4）": "融資受限：FY2027 起算力建置速度減半（驗收 B4）",
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
            pairs[base] = {}; order.append(base)
        pairs[base][side] = rr
    for base in order:
        SE.cell(r, 1, TOR_PLAIN.get(base, base))
        for col, side in (("B", "低"), ("C", "高")):
            if side in pairs[base]:
                SE[f"{col}{r}"] = f"=Sens!E{pairs[base][side]}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調
def coord(wb, C):
    TK, TKV, tk, GEN, QR = C["TK"], C["TKV"], C["tk"], C["GEN"], C["QR"]
    W = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_c0_v0.3.3.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    TKS = f"Tokenomics {TKV}（{TK['source']['file']}，master {TK['source']['commit'][:7]}；快照 spacex/data/tokenomics_snapshot_{TKV}.json）"
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Recon") + 1)
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJKL", (30, 12, 16, 16, 12, 28, 60, 22, 60, 22, 40)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.2.8 起；v0.2.9 改讀 Wattonomics v0.3.3；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。SpaceX（含 xAI）是算力賣方兼買方（Colossus 自建自用並對外出租），不做兩讀數目標價；"
                "本模型只讀 Tokenomics，不讀 Wattonomics（無 WIF_ 名稱進模型）。C0 的個股側只用公司已公布實績（424B4、2026Q3 10-Q），層級側用 Wattonomics v0.3.3 的 xai 玩家隱含值；"
                "追公式（F1b）顯示 xai 的資本支出與 2026 年底 GW 皆由個股模型 v0.2.1 的寫回值換算（同源），2025 年底上線 GW 在 v0.3.2 來自 Epoch 晶片功耗（不同源），v0.3.3 起改讀公司 424B4 揭露 0.8 GW（層級值吸收公司揭露，同源）：對帳單 20261115 SpaceX／xAI 2025 年底算力（Drive 02-AI供應鏈/對帳/已處理/） 已由 Wattonomics v0.3.3（PR #581）吸收，本頁對帳中 0 項。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    for j, h in enumerate(["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1a = r
    ws.cell(r, 1, "讀取層級：Tokenomics（算力單位經濟）——本模型讀 v21.10 快照（工作單所列現行 v22.0 本輪不換；v22.0 與 v21.10 對本模型所讀名稱數值相同，換快照另走 Tokenomics 同步通知）；Wattonomics——本模型不讀，v0.3.3 快照只用於下方 C0 對帳列"); ws.cell(r, 1).font = BOLD; r += 1

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
    gens = (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr"))
    for code, nm in gens:
        tkrows.append((f"IF_HoldEcon｜{code}", code, tk("IF_HoldEcon", code), f"=he_{nm}", "US$m/MW-年（＝$B／GW／年）", "每 MW 年持有成本（經濟口徑；收入錨 k＝1）"))
    for code, nm in gens:
        tkrows.append((f"IF_OpexGW｜{code}", code, tk("IF_OpexGW", code), f"=opx_{nm}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    for code, nm in (("GB300", "b3"), ("VR200", "vr")):
        tkrows.append((f"IF_CapexIT｜{code}", code, tk("IF_CapexIT", code), f"=cit_{nm}", "US$m/MW", "每 MW IT 資本支出"))
    tkrows.append(("IF_CapexFacility｜VR200", "VR200", tk("IF_CapexFacility", "VR200"), "=cfac", "US$m/MW", "每 MW 機房資本支出"))
    for code, nm in gens[:3]:
        tkrows.append((f"IF_DeprIT｜{code}", code, tk("IF_DeprIT", code), f"=dpi_{nm}", "US$m/MW-年", "每 MW IT 年折舊（期初機隊汰換）"))
    F1R = {}
    for nm_, yr_, lv, ad, un, nt in tkrows:
        F1R[nm_] = f1(nm_, yr_, lv, ad, un, TKS, nt, "0.000")
    F1b = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；應為 0）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{F1a + 2}:D{F1b}-C{F1a + 2}:C{F1b}))").number_format = "0.000000"; F1chk = r; r += 1
    ws.cell(r, 1, "Wattonomics v0.3.3 快照名稱（xai＝SpaceX／xAI 玩家碼；只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (
            ("WIF_PlayerGW_YE_xai_2025", "ye25", "GW", "xai 2025 年底上線 IT GW（CapexGW!F240←Players!F182←F176←F174←In!J113 mw_xai 2025＝800 MW：SpaceX 424B4 名目算力 0.8 GW〔Wattonomics 來源 GL-346，Interested-party〕；v0.3.3 吸收對帳單，層級值吸收公司揭露：同源。v0.3.2 為 Epoch 晶片功耗 308 MW × 1.85＝0.570，該分支 v0.3.3 起未生效）"),
            ("WIF_PlayerGW_YE_xai_2026", "ye26", "GW", "xai 2026 年底上線 IT GW（＝2025 年底〔公司揭露〕＋2026 新增 GW〔個股模型 AI 資本支出寫回 In!K215 換算〕：兩段皆同源）"),
            ("WIF_PlayerGW_YE_xai_SiteOnly_2026", "ye26s", "GW", "只有園區證據路徑（Wattonomics 對 2026 年底採 mw_xai 目標 2 GW：公司目標，Interested-party）"),
            ("WIF_Gap_Player_xai_2026_AddGW", "addgw", "GW", "2026 新增 IT GW（Gap!G409←CapexGW!G239←G228←In!K215 gcapai_xai＝SpaceX 個股模型 v0.2.1 基準寫回 74.058：同源）"),
            ("WIF_Gap_Player_xai_2026_CapexAI", "capai", "US$B", "2026 AI 資本支出（Gap!G415＝新增 GW × 每 GW 成本，含 2027–28 廠房提前；2027–28 輸入也是個股模型寫回 In!L215、M215：同源）"),
            ("WIF_Gap_Player_xai_2026_CapexNonAI", "capnai", "US$B", "2026 非 AI 資本支出（Gap!G419←In!K216 gnai_xai＝個股模型 Space＋Connectivity 資本支出寫回：同源）"),
            ("WIF_Gap_Player_xai_2026_CapexAIModel", "capmod", "US$B", "個股模型 AI 資本支出回寫（Interface!F704；＝In!K215）"),
            ("WIF_Gap_Player_xai_2025_CapexAI", "capai25", "US$B", "2025 AI 資本支出：新增 GW 2025＝0，全額為 2026、2027 新增 GW 的廠房提前入帳（Gap!F413、F414 讀 CapexGW!G239、H239←個股模型寫回）：同源，不是 2025 年的獨立估計")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt, "0.000")
    wrows["chip"] = f1("In 列 chip_xai｜2025 年底擁有 AI 晶片功耗", "2025", WR["chip_xai"]["values"][2], None, "MW", WV, "Epoch Colossus 晶片功耗（Analogy；C 段 6.2）；v0.3.3 起只作對照（追公式：分支未生效）", "0")
    wrows["klo"] = f1("In 列 k_chip2it｜低", "—", WR["k_chip2it"]["values"][0], None, "x", WV, "晶片功耗 → IT 電力係數低端（repo mw.conv.itPerTdp；Assumed）", "0.00")
    wrows["kmid"] = f1("In 列 k_chip2it｜基準", "—", WR["k_chip2it"]["values"][1], None, "x", WV, "基準", "0.00")
    wrows["khi"] = f1("In 列 k_chip2it｜高", "—", WR["k_chip2it"]["values"][2], None, "x", WV, "高端：Epoch 設施係數 2.5 ÷ 1.2", "0.00")
    wrows["obs"] = f1("In 列 obs_xai｜2026Q2 名目算力", "2026-06-30", WR["obs_xai"]["values"][1] / 1000.0, None, "GW", WV, "公司 2026Q2 10-Q／新聞稿揭露 1.4 GW（Interested-party；B#46 屬實）；Wattonomics 註明「2026 年底已有目標值，本觀測不進計算」：與公司揭露同源", "0.00")
    wrows["tgt"] = f1("In 列 mw_xai｜2026 年底目標", "2026", WR["mw_xai"]["values"][3] / 1000.0, None, "GW", WV, "公司目標「2026 年底 >2 GW」（Interested-party；目標不是實績）", "0.00")
    wrows["gog"] = f1("In 列 al_xai_googl｜出租給 Google", "2026", WR["al_xai_googl"]["values"][4], None, "GW", WV, "Google 租用 Colossus 2（約 11 萬顆 GPU，153–226 MW-IT 取中 190；新聞與 Wattonomics 證據；公司未揭露客戶別 MW）", "0.00")
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 xai（SpaceX／xAI）玩家路徑是否讀公司實績或目標（20261115_Wattonomics_v0.3.3.xlsx；追公式工具為 main 版 tools/coord/trace_wattonomics.py；規則：有讀者的部分屬同源，不當獨立證據）").font = BOLD; r += 1
    for j, h in enumerate(["輸出名稱 ← 公式鏈", "", "", "", "", "", "同源判定", "讀到什麼（輸入、公式、儲存格）"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    TR = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_trace_v0.3.3.json"), encoding="utf-8"))
    TRC = [
        ("WIF_Gap_Player_xai_2026_AddGW ← Interface!F707 ← Gap!G409 ← CapexGW!G239 ← G237（上線時差加權）← G236 ← G235 ← G234 ← G228 ← In!K215（AI 資本支出寫回）",
         "同源（個股模型 v0.2.1 基準寫回）",
         "In!K215（AI 資本支出寫回）＝SpaceX 個股模型基準情境 B 的 FY2026E SI 資本支出 74.058（CF!G55；含 9M26 實績，Wattonomics 標 Derived，來源編號 GL-318，commit 246d492）；追公式標記「公司實績／指引／個股模型寫回」。GW＝資金 ÷ 每 GW 成本（Tokenomics；xai 世代組合全為 NVIDIA、GB300 78%／VR200 22%），所以 2026 年新增 1.607 GW 是個股模型資本支出的換算，不是獨立的物理證據。本模型 v0.2.7 的 CF!G55 仍為 74.058（與寫回值一致，無過期差）"),
        ("WIF_Gap_Player_xai_2026_CapexAI ← Interface!F699 ← Gap!G431 ← G415（＝G412＋G413＋G414）", "同源（同上）",
         "G412＝2026 新增 GW × 每 GW 成本；G413、G414＝2027、2028 新增 GW 的廠房提前入帳，讀 CapexGW!H239、I239←In!L215、M215（個股模型 2027–28 寫回 136.252、144.121）；G426（個股模型回寫）74.058 並列，不替代"),
        ("WIF_Gap_Player_xai_2026_CapexNonAI ← Interface!F700 ← Gap!G422 ← G419 ← In!K216（非 AI 資本支出寫回）", "同源（個股模型寫回）",
         "In!K216＝個股模型 Space＋Connectivity 資本支出 11.511（CF!G53＋G54）；Wattonomics 對 xai 不另估非 AI"),
        ("WIF_PlayerGW_YE_xai_2025 ← Interface!E32 ← CapexGW!F240 ← Players!F182 ← F176 ← F174 ← In!J113（mw_xai 2025＝公司揭露 800 MW）", "同源（層級值吸收公司揭露）",
         "v0.3.3（PR #581；Decisions GL01-108）把 In!J113 由空白改為 800 MW（SpaceX 424B4 Summary 分部營運資料表與 MD&A Key Business Metrics 的 2025-12-31 Nameplate Compute Draw 0.8 GW；Wattonomics 來源 GL-346；IT 口徑），Players 已知值優先序改走 F174（生效）；F175＝Epoch 晶片功耗 308 MW（In!J103）× 1.85（In!I7）＝0.570 變為分支未生效，只作對照。追公式工具旗標以關鍵字判斷未自動標 J113，本版以 --manual 人工補標「公司揭露」。因此 2025 年底層級值＝公司揭露，差距恆為 0，已無獨立檢驗力"),
        ("WIF_PlayerGW_YE_xai_2026 ← Interface!F32 ← CapexGW!G240＝F240＋G239", "同源（兩段皆是）",
         "2025 年底（公司揭露 0.8，同源）＋2026 新增 GW 1.607（同源於個股模型寫回）＝2.407 GW（v0.3.2 為 2.177）；層級另有只計園區的路徑 2 GW（採公司目標 2026 年底 >2 GW：目標，不是實績）"),
        ("WIF_Gap_Player_xai_2025_CapexAI ← Interface!E699 ← Gap!F431 ← F415", "同源（個股模型寫回）",
         "2025 新增 GW＝0（Gap!F412＝0）；22.893 全部是 Gap!F413、F414（2026、2027 新增 GW 的廠房提前入帳，讀 G409、H409←個股模型寫回）。所以層級對 2025 年資本支出的值不是 2025 年的獨立估計，不能拿公司 FY2025 資本支出當獨立 C0 對照"),
        ("公司已公布實績列（Wattonomics In 頁；本玩家無此列）", "未進路徑（無此列）",
         "反向查詢：讀公司已公布實績列的公式 0 格；Wattonomics 沒有 xai 的 act26 列，公司 9M26 實績不進 xai 的 GW 或資本支出路徑"),
        ("公司目標列（In 頁：Musk「2027 年底接近 10 GW」、2026 年底 >2 GW）", "公司目標（不進主路徑或只作 SiteOnly）",
         "2027 年目標不入主路徑，只作比較敏感度；2026 年底 2,000 MW 目標只進只計園區的路徑。目標不是實績，不能當 C0 的個股側"),
        ("追公式限制（人工確認）", "已人工確認",
         "(1) 追公式旗標以關鍵字（個股模型、指引、法說、act26、公告）判斷，漏掉「目標」字樣：本表以人工補標 2026 年底 MW 目標與 2027 年目標兩列為公司目標（Interested-party）。(2) 工具的條件讀取判斷以快取值遞迴：2026 根的 IF 條件皆可算（無「不確定」）。(3) 玩家碼確認為 xai（不是 spacex、spcx）；非 AI 資本支出 In!K216 由個股模型 Space＋Connectivity 寫回，Wattonomics 另無獨立估計"),
    ]
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 對 SpaceX／xAI 的 2026 年資本支出（AI＋非 AI）與新增 GW 全由個股模型 v0.2.1 寫回值換算，2025 年的資本支出也是 2026–27 寫回值的廠房提前，整列同源；2025 年底上線 GW 在 v0.3.2 為 Epoch 晶片功耗 × 係數（0.570，與公司揭露獨立），v0.3.3 吸收對帳單後改讀公司 424B4 揭露 0.8（同源）。因此資本支出列與名目算力列都標「參考（同源校準）」，結論「無可比實績（層級值同源／層級值吸收公司揭露）」；本版 xai 路徑上已沒有可做獨立 C0 的量。")
    rtr_conc = r; r += 1
    ws.cell(r, 1, "追公式傾印（由 main 版 tools/coord/trace_wattonomics.py --player xai --year 2026 產生；每格公式、值與生效狀態；來源 " + TR["source"].split("/")[-1] + "；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 7, "標記"); ws.cell(r, 8, "公式（省略開頭等號）"); ws.cell(r, 4, "值"); ws.cell(r, 2, "狀態（生效／條件讀取／分支未生效）"); r += 1
    rdump0 = r
    for x in TR["rows"]:
        ws.cell(r, 1, ("　" * x["depth"]) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 2, x.get("status", "")); ws.cell(r, 7, x["flag"]); ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    r += 1
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝424B4／10-Q 實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    for j, h in enumerate(["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F2a = r
    SRC_Q = "SpaceX 2026Q3 Form 10-Q 附註 18 分部資訊「資本支出」與現金流量表（mag/extract/docs/2026Q3/spacex_2026Q3_10-Q.pdf）；424B4 MD&A Capital Expenditures；Q 頁第 " + str(QR["capex_si"]) + " 列（SI）、第 " + str(QR["capex_space"]) + "／" + str(QR["capex_conn"]) + " 列"
    GR_Q = "Verified（10-Q 附註 18）"
    SRC_G = "SpaceX 424B4 MD&A Key Business Metrics「Nameplate Compute Draw」：安裝 GPU 數 × 各自全載功耗，不含冷卻、配電損失與機房負載（＝IT 口徑；2025-12-31 為 0.8 GW）；2026Q3 10-Q 同表 2026-09-30 為 1.7 GW、2026-06-30 為 1.4 GW；Q 頁第 " + str(QR["gw"]) + " 列"
    GR_G = "Interested-party（公司自訂 KPI；見公司口徑說明）"
    # ---- 區塊 1：名目算力 2025 年底（唯一獨立 C0）----
    ws.cell(r, 1, "A. 名目算力（IT GW；期末存量）：公司揭露 vs 層級 xai 2025 年底上線 GW（v0.3.3 起＝In!J113 公司揭露 800 MW：層級值吸收公司揭露，同源，參考不判定）").font = BOLD; r += 1
    ws.cell(r, 1, "主讀數（參考，同源）：2025-12-31 名目算力（公司）vs 層級 2025 年底上線 IT GW（Wattonomics v0.3.3；讀公司同一揭露）"); ws.cell(r, 2, "2025-12-31 期末（同日，無時差）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=Q!I{QR['gw']}").number_format = "0.000"; ws.cell(r, 5, f"=C{wrows['ye25']}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, SRC_G); ws.cell(r, 9, GR_G); ws.cell(r, 10, "Wattonomics WIF_PlayerGW_YE_xai_2025（連 F1；In!J113＝424B4 0.8 GW；追公式：同源）"); ws.cell(r, 11, "Interested-party（同源）"); ws.cell(r, 7, "參考（同源校準）")
    rg1 = r; r += 1
    ws.cell(r, 1, "對照（已吸收，不判定）：同一公司值 vs v0.3.2 舊層級路徑 Epoch 晶片功耗 308 MW × 基準換算係數 1.85（v0.3.3 起分支未生效）"); ws.cell(r, 2, "2025-12-31；舊層級＝308 MW × 1.85"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=Q!I{QR['gw']}").number_format = "0.000"; ws.cell(r, 5, f"=C{wrows['chip']}/1000*C{wrows['kmid']}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "同上（同一日期的公司值）"); ws.cell(r, 9, GR_G); ws.cell(r, 10, "Epoch 晶片功耗 308 MW（In!J103）× In!k_chip2it 基準 1.85（連 F1）＝v0.3.2 的 WIF_PlayerGW_YE_xai_2025；即 " + "對帳單 20261115 SpaceX／xAI 2025 年底算力（Drive 02-AI供應鏈/對帳/已處理/）" + " 的層級側原值（v0.2.8 主讀數 +40%、區間上緣 2.08 時 +25%，判定對帳中）；v0.3.3 已吸收，本列只留作差距來源的紀錄"); ws.cell(r, 11, "Analogy／Assumed"); ws.cell(r, 7, "參考（已吸收，不判定）")
    rg2 = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：名目算力（2025 年底）").font = BOLD
    ws.cell(r, 2, "層級值吸收公司揭露（Wattonomics v0.3.3，PR #581 已合併）：差距恆為 0，無檢驗力；規則 v1.4.2 C0 第 2 點同源校準，不判定")
    ws.cell(r, 7, "無可比實績（層級值吸收公司揭露）")
    ws.cell(r, 8, "判定列；v0.2.8 判定對帳中（公司 424B4 0.8 GW vs 層級 Epoch 0.570 GW），已開對帳單並由 Wattonomics v0.3.3 吸收（In!J113 改讀公司揭露）；層級值與公司揭露同源，依規則 v1.4.2 C0 第 2 點標「參考（同源校準）／無可比實績（層級值吸收公司揭露）」，不計入對帳中。缺的獨立揭露：同日的第三方安裝量（Epoch 等）更新到涵蓋 Colossus II。SpaceX 估值不讀層級 GW，不觸發兩讀數目標價")
    ws.cell(r, 9, GR_G); ws.cell(r, 10, "C0 第 3 點：個股（公司揭露）有證據、層（Epoch 估計）無同日獨立證據 → 層在下一次升版吸收；已由 Wattonomics v0.3.3 吸收（" + "對帳單 20261115 SpaceX／xAI 2025 年底算力（Drive 02-AI供應鏈/對帳/已處理/）" + "），本模型不改層級值"); ws.cell(r, 11, "Interested-party（同源）")
    rgj = r; r += 1
    ws.cell(r, 1, "捨入範圍（參考）：公司值取 0.75 GW（揭露 0.8 為一位小數，捨入下限）vs 層級（＝公司揭露 0.8）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, 0.75).number_format = "0.000"; ws.cell(r, 4).font = BLUE; ws.cell(r, 5, f"=E{rg1}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 2, F(T("對層級（同源）"), f'TEXT(D{r}/E{rg1}-1,"+0%;-0%")', T("；對 v0.3.2 舊層級 "), f'TEXT(D{r}/E{rg2}-1,"+0%;-0%")'))
    ws.cell(r, 7, "參考（同源；捨入範圍內，不判定）")
    ws.cell(r, 8, "424B4 揭露 0.8 GW 為一位小數；層級 v0.3.3 取揭露值 0.8，捨入下限 0.75 時差 −6%，在 ±門檻內；v0.3.2 舊層級對照仍差 +32%（吸收前的差距不因捨入而消失）"); ws.cell(r, 9, GR_G); ws.cell(r, 10, "同主讀數與對照列的層級值"); ws.cell(r, 11, "Interested-party（同源）")
    rgr = r; r += 1
    ws.cell(r, 1, "公司值與層級的口徑差（參考）")
    ws.cell(r, 2, F(T("兩側皆 IT 口徑：v0.3.3 層級直接讀公司揭露（GPU 數 × 全載功耗，不含冷卻、配電等設施負載；Wattonomics In!C113 標 IT），口徑一致。v0.3.2 舊路徑（Epoch 晶片功耗 × 換算係數 1.85）要補足公司值，換算係數需 "),
                    f'TEXT(D{rg1}/(C{wrows["chip"]}/1000),"0.00")', T("（> Epoch 設施係數 2.5），口徑無法解釋，差距來自範圍（含 Colossus II）或 Epoch 時點；這是對帳單判層吸收的理由。")))
    rkq = r; r += 1
    # ---- 區塊 2：資本支出（同源，參考）----
    ws.cell(r, 1, "B. 資本支出（流量）：公司 9M26 實績 vs 層級 CY2026（同源於個股模型寫回，參考不判定）").font = BOLD; r += 1
    ws.cell(r, 1, "層級 CY2026 xai AI 資本支出（Wattonomics；同源）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "US$B"); ws.cell(r, 5, f"=C{wrows['capai']}").number_format = "#,##0.0"
    ws.cell(r, 10, WV + "：Gap!G415（新增 GW × 每 GW 成本＋2027–28 廠房提前；輸入為個股模型 v0.2.1 寫回，見追公式）"); ws.cell(r, 11, "Derived；輸入同源於個股模型")
    rl = r; r += 1
    ws.cell(r, 1, "主讀數：SI 分部資本支出 2026Q1–Q3 累計（公司實績）vs 層級 CY2026 × 9 ÷ 12"); ws.cell(r, 2, "9M 累計；層級 × 9/12（線性時差近似）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!H{QR['capex_si']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, SRC_Q); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "層級 CY2026 × 9 ÷ 12（Derived）"); ws.cell(r, 11, "Derived")
    rc1 = r; r += 1
    ws.cell(r, 1, "第二讀數：SI 分部資本支出最近一季（2026Q3）× 4 vs 層級 CY2026"); ws.cell(r, 2, "2026Q3 × 4"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!F{QR['capex_si']}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 8, "2026Q3 單季 × 4（Derived；Q 頁）"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "層級 CY2026 全年"); ws.cell(r, 11, "Derived")
    rc2 = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：SI 資本支出").font = BOLD
    ws.cell(r, 2, "層級值同源於個股模型寫回：參考（同源校準），不判定"); ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由本公司模型 v0.2.1 寫回（In!K215）換算＝同源，不是獨立證據；上方兩個讀數仍列出供參考"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rcj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rc1})>{TOL},ABS(F{rc2})>{TOL},SIGN(F{rc1})=SIGN(F{rc2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rc1})>{TOL},ABS(F{rc2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rcjr = r; r += 1
    ws.cell(r, 1, "非 AI 資本支出（Space＋Connectivity）2026Q1–Q3 累計 vs 層級 CY2026 × 9 ÷ 12（同源，參考）"); ws.cell(r, 2, "9M 累計"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!H{QR['capex_space']}+Q!H{QR['capex_conn']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=C{wrows['capnai']}*9/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "無可比實績（層級值同源；參考同源校準）"); ws.cell(r, 8, SRC_Q); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "Wattonomics In!K216 gnai_xai（個股模型寫回）；層級 CY2026 × 9/12"); ws.cell(r, 11, "Derived（同源校準）")
    rcn = r; r += 1
    # ---- 區塊 3：無可比實績與並列 ----
    ws.cell(r, 1, "C. 沒有可比公司實績的項目與並列讀法（參考；不判定）").font = BOLD; r += 1
    ws.cell(r, 1, "名目算力 2026 年底（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "2026 年底尚未發生；公司最近揭露 2026-09-30 為 1.7 GW（10-Q）。層級 2026 年底 2.177 GW（同源於個股模型寫回）與公司目標 >2 GW（SiteOnly）都不是實績"); ws.cell(r, 9, "—"); ws.cell(r, 10, "WIF_PlayerGW_YE_xai_2026／SiteOnly_2026（連 F1）"); ws.cell(r, 11, "Derived（同源）"); rn1 = r; r += 1
    ws.cell(r, 1, "2026-06-30 名目算力：公司 1.4 GW vs Wattonomics obs_xai 1.4 GW"); ws.cell(r, 2, "2026-06-30（同源）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=Q!E{QR['gw']}").number_format = "0.00"; ws.cell(r, 5, f"=C{wrows['obs']}").number_format = "0.00"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（同源），不判定"); ws.cell(r, 8, SRC_G); ws.cell(r, 9, GR_G); ws.cell(r, 10, "Wattonomics In obs_xai＝公司同一揭露（B#46）：不是獨立證據，差距恆為 0"); ws.cell(r, 11, "Interested-party（同源）"); rn2 = r; r += 1
    ws.cell(r, 1, "對外出租 GW／營收的客戶別實績（例：出租給 Google 0.19 GW）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "公司只揭露 SI 對外雲端服務營收（合計）與 Backlog，不揭露客戶別 GW 或 MW；Wattonomics 的 0.19 GW 取自新聞與調查證據（Derived），無公司揭露可比"); ws.cell(r, 9, "—"); ws.cell(r, 10, "Wattonomics al_xai_googl（連 F1）"); ws.cell(r, 11, "Derived"); rn3 = r; r += 1
    ws.cell(r, 1, "並列（不判定）：SpaceXAI 對 Anthropic 的付款讀法——Anthropic 模型基準每月 1.25 × 36 個月"); ws.cell(r, 2, "既定決定（Andy 11-04「依建議」）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, "=1.25*36").number_format = "#,##0.0"; ws.cell(r, 7, "並列（既定基準；不判定）")
    ws.cell(r, 8, "基準＝每月 1.25 十億美元 × 36 個月＝45（Anthropic 收支模型 Cost!D143；Andy 2026-11-04「依建議」，維持）；本表只並列、不改。此為客戶側（Anthropic）付款假設，SpaceX 公司未揭露客戶別合約金額"); ws.cell(r, 9, "Assumed（既定決定）"); ws.cell(r, 10, "—"); ws.cell(r, 11, "—"); rn4 = r; r += 1
    ws.cell(r, 1, "並列（不判定）：S-1 月費路徑——滿載每月約 3.1、上限 120.5（2027Q1 後可 90 天通知取消）"); ws.cell(r, 2, "S-1／A 付款表（並列）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, 120.5).number_format = "#,##0.0"; ws.cell(r, 4).font = BLUE; ws.cell(r, 7, "並列（S-1 路徑；不判定）")
    ws.cell(r, 8, "S-1／A 寫滿載每月約 3.1、上限 120.5、2027Q1 後可 90 天通知取消（Anthropic 收支模型 SRC_ANT!E519、E457；本專案 Anthropic 交接）。若照此路徑付滿，Anthropic 側結論會翻轉；本表只並列、基準不改。公司揭露的合併 Backlog 為 62.959（10-Q，67% 一年內；含全部客戶與分部）："); ws.cell(r, 9, "Interested-party（S-1 條款；非本公司實績）"); ws.cell(r, 10, "—"); ws.cell(r, 11, "—"); rn5 = r; r += 1
    ws.cell(r, 1, "　　對照：公司 Backlog（10-Q，2026-09-30）"); ws.cell(r, 2, "2026-09-30"); ws.cell(r, 3, "US$B"); ws.cell(r, 4, 62.959).number_format = "#,##0.0"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 7, "參考（總額含全部客戶，無客戶別，不判定）"); ws.cell(r, 8, "SpaceX 2026Q3 10-Q Backlog 62,959 百萬美元（其中遞延收入 17,520；約 67% 一年內認列）；Recon 頁 G7 另有與模型的對照"); ws.cell(r, 9, GR_Q); ws.cell(r, 10, "—"); ws.cell(r, 11, "—"); rn6 = r; r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定；v0.2.8 為 1，v0.2.9 起 0：名目算力已由 Wattonomics v0.3.3 吸收）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "時差邊界（參考）項目數"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"時差邊界（參考）")'); rnb = r; r += 1
    ws.cell(r, 1, "無可比實績列數（不含參考列）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F(T("C0（實績）對帳：公司揭露 2025 年底名目算力 "), f'TEXT(D{rg1},"0.00")', T(" GW；Wattonomics v0.3.3 已吸收對帳單，層級 xai 2025 年底上線 GW 改讀公司同一揭露（"), f'TEXT(E{rg1},"0.00")', T(" GW，相差 "), f'TEXT(F{rg1},"+0%;-0%")',
                    T("），判定："), f'G{rgj}', T("（v0.3.2 舊層級 Epoch 晶片功耗 × 1.85＝"), f'TEXT(E{rg2},"0.00")', T(" GW，差 "), f'TEXT(F{rg2},"+0%;-0%")', T("，即已處理對帳單的原差距）；對帳中 "), f'TEXT(D{rn},"0")', T(" 項"),
                    T("。資本支出的層級值（xai 2026 年 "), f'TEXT(E{rl},"0.0")', T(" 十億美元）由本公司模型 v0.2.1 寫回值換算（追公式：CapexGW!G228←In!K215），屬同源，判定：無可比實績（層級值同源），不計入對帳中。"
                    "2026 年底算力、客戶別出租與 Anthropic 付款讀法沒有可比的公司實績；每月 1.25 × 36＝45 為既定基準，S-1 月費路徑只並列。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向")
    ws.cell(r, 2, F(T("名目算力是同日期末存量，沒有時差：公司揭露 2025-12-31 為 "), f'TEXT(D{rg1},"0.0")', T(" GW，之後 2026Q1 "), f'TEXT(Q!D{QR["gw"]},"0.0")', T("、Q2 "), f'TEXT(Q!E{QR["gw"]},"0.0")', T("、Q3 "), f'TEXT(Q!F{QR["gw"]},"0.0")',
                    T(" GW。層級 2025 年底 v0.3.3 起讀公司揭露（差距恆為 0）；v0.3.2 舊層級 0.57 GW 低於公司揭露，方向是層級偏低（Epoch 晶片功耗可能未涵蓋 Colossus II 已安裝或在途的 GPU，或 Epoch 資料時點較早；換算係數需高於設施係數才補足，口徑差無法解釋），已由對帳單送層吸收；"
                      "公司口徑（GPU 數 × 全載功耗）是安裝量，不代表實際用電或稼動。資本支出的兩讀數（SI 9M 累計與 Q3 年化）都低於層級 CY2026，但層級值同源於個股模型 FY2026E，差距反映的是個股模型 Q4 26E 與 FY2027–28 廠房提前入帳，不是公司實績與獨立證據的差距。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出：追公式（F1b）顯示 Wattonomics xai 的 2026 年 AI 資本支出與新增 GW 的輸入是個股模型 v0.2.1 寫回的 FY2026E SI 資本支出 74.058（In!K215），2027–28 廠房提前也讀個股模型寫回；2025 年的資本支出值全額來自這些寫回值的廠房提前。整列標「參考（同源校準）」，不判定。"
                   "（二）名目算力：2025 年底層級值在 v0.3.2 為 Epoch 晶片功耗 × 換算係數（0.570 GW，不讀公司揭露，是 v0.2.8 唯一獨立的比較，判定對帳中）；v0.3.3 吸收對帳單後改讀公司 424B4 揭露 0.8 GW（In!J113），公司值與層級值消去，比較只剩「公司揭露 vs 公司揭露」。2026 年底層級值同源（寫回）或為公司目標；2026-06-30 的觀測 1.4 GW 即公司同一揭露。"
                   "（三）所以本頁 C0 已無獨立檢驗力；對帳中 0 項不代表兩層由獨立證據確認一致。要恢復檢驗力，需 Wattonomics 對 xai 改用獨立證據（園區、Epoch 更新到涵蓋 Colossus II）而非公司揭露或個股模型寫回（對帳單第四節 (3)，列 Wattonomics 待辦）。"))
    rsame = r; r += 1
    ws.cell(r, 1, "口徑換算（層級 vs 公司資本支出）")
    ws.cell(r, 2, ("公司資本支出＝購置不動產與設備（現金流量表，不含以其他融資取得的 PP&E 3.921；SI 分部口徑見 10-Q 附註 18）；層級資本支出＝新增 GW × 每 GW 建置成本（IT＋廠房）。"
                   "SpaceX 有融資租賃與關係人融資（Valor 售後租回列融資、不扣），Wattonomics 對 xai 的 Funding 列另有融資結構；本頁不另調整。Wattonomics 對 xai 為擁有者與出租方口徑（出租給 Google 的 0.19 GW 屬 xai 的出租，不重複計）。"))
    rk = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 的個股模型寫回（In!K215）vs 本模型現值 CF 頁 FY2026E SI 資本支出（差距＝寫回是否過期）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['capmod']}").number_format = "#,##0.000"; ws.cell(r, 5, f"=CF!G{C['CR']['B']['cx_si']}").number_format = "#,##0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics 寫回值（commit 246d492，個股模型 v0.2.1）"); ws.cell(r, 10, "本模型 v0.2.9 CF 頁情境 B FY2026E（與 v0.2.8 相同）"); rb_ = r; r += 1
    ws.cell(r, 1, "Wattonomics 層級 CY2026 xai AI 資本支出（GW 推算）vs 其讀入的個股模型 AI 資本支出（廠房提前造成的差距）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=C{wrows['capai']}").number_format = "#,##0.0"; ws.cell(r, 5, f"=C{wrows['capmod']}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Gap!G415 對 G426"); rb2 = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：公司 2026Q4 指引、2026 年底 >2 GW 目標、Musk「2027 年底接近 10 GW」不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：Tokenomics 本模型讀 v21.10（工作單所列現行 v22.0 本輪未換；兩版對本模型所讀名稱數值相同）；Wattonomics v0.3.3 也讀 Tokenomics v21.10，兩側版本一致（v0.2.9 由 Wattonomics v0.3.2 換讀 v0.3.3：xai 只有 2025 年底 GW 起變動，2026 年底 GW 隨之 +0.230；資本支出與寫回值不變）。SpaceX 個股模型值曾被 Wattonomics 讀入（營業現金流、資本支出寫回），本表只列層級值與追公式，不用該回饋判定。")
    r += 1
    return dict(F1=(F1a, F1b), F1chk=F1chk, F2h=F2h, F2=(F2a, F2b), rg=(rg1, rg2, rgj), rgr=rgr, rkq=rkq, rl=rl, rc=(rc1, rc2, rcj, rcjr, rcn), rtr=rtr_conc, rdump=(rdump0, rdump1),
                rn=(rn1, rn2, rn3, rn4, rn5, rn6), rncount=rn, rnb=rnb, rnc=rnc, rs=rs, rr=rr_, rsame=rsame, rk=rk, rb=(rb_, rb2), wrows=wrows, tol=TOL)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, RVA, PLA, CRD):
    RD = wb["README"]
    VR, PRB = C["VR"], C["PR"]["B"]
    vb = lambda k, c="B": f"Val!{c}{VR[k]}"
    RCw = wb["Recon"]
    r_cons = next(x for x in range(1, RCw.max_row + 1) if str(RCw.cell(x, 1).value or "").startswith("市場共識目標價"))
    tp, up, rec, dcf, cmpv, cons = vb("tp"), vb("up"), vb("rec"), vb("tpdcf"), vb("tpcomp"), f"Recon!C{r_cons}"   # 共識經 Recon 頁引用（Val 頁共識格須維持顯示用）
    rv = lambda k: f"RevDCF!B{RVA[k]}"
    G_ = f"CF!C{C['GATE']}"
    out = {}
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.2.8 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec, T("：現金流折現 "), f'TEXT({dcf},"0")',
        T(" 與分部加總 "), f'TEXT({cmpv},"0")', T(" 各占一半。xAI 算力出租撐起營收，但 AI 資本支出使五年自由現金流合計 "), f'TEXT({G_},"0")', T(" 十億美元。"))
    put("rec", T("投資建議「"), rec, T("」的前提：自用 AI 採價值中性（讀法 2）；改為只計成本（讀法 1）時加權目標價 "), f'TEXT(Val!B{VR["tp_r1"]},"0")', T(" 美元；每 GW 年營收相對持有成本的倍數依情境路徑收斂；現價 "), 'TEXT(px,"0.00")', T(" 美元為 2026-11-05 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價 "), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；市場共識 "), f'TEXT({cons},"0")', T("。三個情境（"), f'TEXT({vb("scen")},"0")', T("／"), f'TEXT({vb("scen", "C")},"0")', T("／"), f'TEXT({vb("scen", "D")},"0")', T("）都低於現價。"))
    put("o2", T("反向推算（情境 B）：要支持現價，企業價值須放大 "), f'TEXT({rv("evx")},"0.0")', T(" 倍，或 FY2028 起營收與營業利益全面放大 "), f'TEXT({rv("s1")},"0.0")', T(" 倍；分部加總隱含 "),
        f'TEXT(RevDCF!B{RVA["smul"]},"0.0")', T(" 倍 EBITDA，模型 "), f'TEXT(RevDCF!C{RVA["smul"]},"0.0")', T(" 倍。"))
    SEw = wb["Sens"]
    srow = lambda pre: next(x for x in range(1, SEw.max_row + 1) if str(SEw.cell(x, 1).value or "").startswith(pre))
    r_knc, r_he = srow("k 不收斂（三情境"), srow("Tokenomics IF_HoldEcon 四世代同時")
    put("o3", T("核心問題：FY2028 每 GW 年營收 ÷ 持有成本，情境 A／B／C 為 "), f'TEXT({vb("s2")},"0.00")', T("／"), f'TEXT({vb("s2", "C")},"0.00")', T("／"), f'TEXT({vb("s2", "D")},"0.00")',
        T("；若這個倍數不收斂（維持 2026Q3 校準值），加權目標價相對基準 "), f'TEXT(Sens!E{r_knc},"+0%;-0%")', T("，是影響最大的單一假設之一。"))
    put("o4", T("資金：FY2026–FY2030 自由現金流合計（收支關卡）情境 A／B／C 為 "), f'TEXT(CF!B{C["GATE"]},"0")', T("／"), f'TEXT({G_},"0")', T("／"), f'TEXT(CF!D{C["GATE"]},"0")',
        T(" 十億美元，三情境皆為負，需外部融資；Tokenomics 四個世代的持有成本同時取低值時，加權目標價相對基準 "), f'TEXT(Sens!E{r_he},"+0%;-0%")', T("。"))
    put("o5", T("實績對帳：公司揭露 2025 年底名目算力 "), f'TEXT(Coord!D{CRD["rg"][0]},"0.00")', T(" GW，原層級 Epoch 推得 "), f'TEXT(Coord!E{CRD["rg"][1]},"0.00")', T(" GW（相差 "), f'TEXT(Coord!F{CRD["rg"][1]},"+0%;-0%")',
        T("）；對帳單已由 Wattonomics v0.3.3 吸收，層級改讀公司揭露，判定："), f'Coord!G{CRD["rg"][2]}', T("，對帳中 "), f'TEXT(Coord!D{CRD["rncount"]},"0")', T(" 項；資本支出與 2026 年底 GW 也同源，無可比實績。"))
    put("drv", T("營收推導（情境 B）：Space＝客戶發射 × 每次收入＋發射與開發；Connectivity＝平均訂戶 × ARPU ×（含 Kits 的調整）＋企業與政府；SI＝對外雲端服務（在役 GW × 對外比例 × 每 GW 年營收）＋Grok／Cursor／X 訂閱、資料授權與 API＋廣告，加總為合併營收。FY2028 Space "),
        f'TEXT(PL!I{PRB["space"]},"0.0")', T("、Connectivity "), f'TEXT(PL!I{PRB["conn"]},"0")', T("、SI "), f'TEXT(PL!I{PRB["si"]},"0")', T("，合計 "), f'TEXT(PL!I{PRB["rev"]},"0")', T(" 十億美元。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：Coord 頁改讀 Wattonomics v0.3.3（PR #581，xai 2025 年底 GW 吸收公司揭露 0.8 GW），名目算力 C0 改「參考（同源校準）／無可比實績（層級值吸收公司揭露）」，對帳中 1→0；估值不讀層級 GW，目標價與評等不變。"
                  "v0.2.8（2026-11-15）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；新增 RevDCF 頁（反向 DCF）、PL 頁末標準損益表"
                  "（FY2023A–FY2025A 實績＋FY2026E–FY2032E 情境 B，GAAP）、Sens 頁末加權目標價相對基準（龍捲風圖用）、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」（F1 讀取層級；F1b 追公式〔main 版 tools/coord/trace_wattonomics.py〕；F2 C0（實績））。"
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
