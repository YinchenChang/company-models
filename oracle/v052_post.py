# Oracle v0.5.2（2026-11-15）後處理：個股 HTML 版面規格 v1.3 補缺項＋Coord 頁「三層協調」節（MAG 類樣板；樣板 Alphabet v0.3.9 alphabet/v039_post.py）。
# 工作單：docs/workorders/20261114_第二批13家_v1.3與三層協調.md；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點。
# 只新增，不改任何既有儲存格（既有數值、目標價、評等不變；逐格比對只允許版號文字與新增格）：
#   IS 頁（新）：標準損益表（FY2024A–FY2026A 10-K 實績＋FY2027E–FY2032E 情境 B；GAAP）、比率區、核對列、分部營收（圖表用）、10-K 實績輸入
#   Sens 頁末：加權目標價相對基準（HTML 龍捲風圖用，白話標籤）
#   Coord 頁（新）：F1 讀取層級（Tokenomics v21.10 快照；Wattonomics v0.3.2 只作對帳列、不進模型）、F1b 追公式、F2 C0（實績）：10-Q 資本支出（扣 2025-12）與 OCI 營收、MW vs 層級隱含值
#   README 頁末：v1.3 展示層白話列（橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字、版本沿革）
# build_oracle.py 在 README 與收支模組之後、存檔前呼叫 apply(wb, ctx)。
import json
import os
import re

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.5.2", "2026-11-15"
U_10Q_Q1 = "https://www.sec.gov/Archives/edgar/data/1341439/000119312526389274/orcl-20260831.htm"   # Form 10-Q（季末 2026-08-31）
U_10K = "https://www.sec.gov/Archives/edgar/data/1341439/000119312526277521/orcl-20260531.htm"       # Form 10-K FY2026（期末 2026-05-31）
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
M3 = "#,##0.000"
PCT = "0.0%"
SPCT = "+0.0%;-0.0%"
PARAM_RE = re.compile(r'(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+(?![A-Za-z0-9])')


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


# =====================================================================================================
# IS：標準損益表（版面規格 v1.3 D 節）；欄：C FY2024A、D FY2025A、E FY2026A、F–K FY2027E–FY2032E、L 說明、M 出處網址
HC = "CDE"                      # 歷史三年
FC = "FGHIJK"                   # 預測六年
ALLC = HC + FC
S2P = dict(zip(FC, "DEFGHI"))   # IS 欄 → PL 頁欄（FY2027→D）
S2T = dict(zip(FC, "EFGHIJ"))   # IS 欄 → Thesis 頁欄（FY2027→E）
S2Q = dict(zip(HC, "JKL"))      # IS 欄 → Q 頁欄（FY2024→J）

# 10-K FY2026 合併損益表（USD 十億；FY2024／FY2025／FY2026；Verified，經查核）
K10 = {
    "i_cs": ("Cloud and software expenses（雲端與軟體費用）", (9.427, 11.569, 17.597)),
    "i_hw": ("Hardware expenses（硬體費用）", (0.891, 0.782, 0.868)),
    "i_svc": ("Services expenses（服務費用）", (4.825, 4.576, 4.556)),
    "i_sm": ("Sales and marketing（行銷）", (8.274, 8.651, 8.331)),
    "i_rd": ("Research and development（研發）", (8.915, 9.860, 10.272)),
    "i_ga": ("General and administrative（管理）", (1.548, 1.602, 1.618)),
    "i_am": ("Amortization of intangible assets（無形資產攤銷）", (3.010, 2.307, 1.671)),
    "i_rs": ("Restructuring and other（重組與其他）", (0.718, 0.374, 1.838)),
    "i_no": ("Non-operating income (expenses), net（非營業收入淨額，不含利息費用）", (-0.098, 0.060, 3.547)),
    "i_pt": ("Income before income taxes（稅前利益；核對用）", (11.741, 14.160, 19.554)),
    "i_tx": ("Provision for income taxes（所得稅）", (1.274, 1.717, 2.467)),
}


def income_statement(wb, C):
    P = wb["PL"]
    IS = wb.create_sheet("IS")
    PB, TB = C["PR"]["B"], C["TR"]["B"]
    rv, ai, oi, it, ii, pt, tx, ni, pf, sh, eps, da = (PB[k] for k in ("rev", "ai", "oi", "int", "ii", "pti", "tax", "ni", "pref", "sh", "eps", "da"))
    L = {}
    r = 1
    IS.cell(r, 1, "IS — 標準損益表（v0.5.2 起；個股 HTML 版面規格 v1.3 D 節）：FY2024A–FY2026A（10-K 實績）＋FY2027E–FY2032E（情境 B 基準）。US$B；EPS 美元。"
                  "只引用 PL、Thesis、Q 頁與本頁下方 10-K 輸入，不回饋任何計算（既有各頁數值與目標價不變）").font = BOLD
    r = 4
    keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "pref", "sh", "eps"]
    hdr_r = r
    idx = {k: hdr_r + 1 + i for i, k in enumerate(keys)}
    r_ratio = hdr_r + 1 + len(keys) + 1
    r_niy = r_ratio + 1
    r_chk = r_niy + 2
    chk_keys = ["c1", "c2", "c3", "c4", "c5"]
    cidx = {k: r_chk + 1 + i for i, k in enumerate(chk_keys)}
    r_seg = r_chk + 1 + len(chk_keys) + 1
    seg_keys = ["s_oci", "s_ai", "s_saas", "s_sw", "s_hw", "s_svc", "s_tot", "s_chk"]
    sidx = {k: r_seg + 1 + i for i, k in enumerate(seg_keys)}
    r_in = r_seg + 1 + len(seg_keys) + 1
    in_keys = list(K10.keys()) + ["h_cogsr"]
    hidx = {k: r_in + 2 + i for i, k in enumerate(in_keys)}

    IS.cell(2, 1, F(T("口徑：GAAP（Oracle 以 GAAP 揭露損益，市場共識 EPS 也以 GAAP 稀釋為主，共識口徑見 Recon 頁）。其他收入（淨額）＝非營業收入淨額減利息費用；FY2026A 的非營業收入淨額為 "),
                    f'TEXT(E{hidx["i_no"]},"0.0")', T(" 十億美元（10-K 損益表；預測期不估），所以 FY2027E 的稅前與淨利相對 FY2026A 較低。")))
    L["note"] = 2
    IS.cell(3, 1, "營業成本與營業費用：模型以分部利潤與未分攤費用推導營業利益，沒有分列成本與費用。預測期營業成本＝非 AI 營收 × 非 AI 營業成本率（FY2024 營業成本扣非 AI 折舊後占營收，下方假設列）"
                  "＋AI 營運費用與機房租金（Thesis 頁）＋折舊（D&A 扣無形資產攤銷；攤銷與其餘費用記在營業費用）；營業費用＝毛利 − 營業利益（殘差，含研發、行銷、管理、攤銷、重組）。"
                  "營業利益、稅前、稅後與 EPS 全取自 PL 頁情境 B，不因此改變；FY2024A–FY2026A 為 10-K 實績。")
    L["note2"] = 3
    hdr = ["項目", "單位", "FY2024A", "FY2025A", "FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "說明", "出處網址"]
    for j, h in enumerate(hdr, 1):
        c = IS.cell(hdr_r, j, h); c.font = BOLD; c.fill = HDR
    L["hdr"] = hdr_r

    def row(key, label, unit, cells, fmt, note="", url=""):
        rr = {**idx, **{"r_niy": r_niy}, **cidx, **sidx}[key]
        IS.cell(rr, 1, label); IS.cell(rr, 2, unit)
        for col, v in cells.items():
            if v is None:
                continue
            IS[f"{col}{rr}"] = v; IS[f"{col}{rr}"].number_format = fmt
        if note:
            IS.cell(rr, 12, note)
        if url:
            IS.cell(rr, 13, url)
        return rr

    X = hidx
    prevc = dict(zip("DEFGHIJK", "CDEFGHIJ"))
    row("rev", "營收", "US$B", {**{c: f"=Q!{S2Q[c]}10" for c in HC}, **{c: f"=PL!{S2P[c]}{rv}" for c in FC}}, M3, "歷史＝10-K 合併營收（Q 頁）；預測＝PL 頁情境 B 合併營收", U_10K)
    row("yoy", "營收年增率（YoY）", "%", {c: f"={c}{idx['rev']}/{prevc[c]}{idx['rev']}-1" for c in "DEFGHIJK"}, PCT, "FY2024A 的基期（FY2023）不在事實總帳，故空白")
    cg = {c: f"={c}{X['i_cs']}+{c}{X['i_hw']}+{c}{X['i_svc']}" for c in HC}
    for c in FC:
        cg[c] = f"=(PL!{S2P[c]}{rv}-PL!{S2P[c]}{ai})*cogs_nonai_r+Thesis!{S2T[c]}{TB['opex']}+Thesis!{S2T[c]}{TB['rent']}+(PL!{S2P[c]}{da}-amort)"
    row("cogs", "營業成本（COGS）", "US$B", cg, M3, "歷史＝雲端與軟體、硬體、服務三項費用（10-K）；預測＝（營收 − OCI AI 營收）× 非 AI 營業成本率＋AI 營運費用＋AI 機房租金＋折舊（假設列）")
    row("gp", "毛利", "US$B", {c: f"={c}{idx['rev']}-{c}{idx['cogs']}" for c in ALLC}, M3)
    row("gm", "毛利率", "%", {c: f"={c}{idx['gp']}/{c}{idx['rev']}" for c in ALLC}, PCT, "預測期毛利率下降＝AI 折舊與機房租金（損益進營業成本）上升")
    ox = {c: f"={c}{X['i_sm']}+{c}{X['i_rd']}+{c}{X['i_ga']}+{c}{X['i_am']}+{c}{X['i_rs']}" for c in HC}
    for c in FC:
        ox[c] = f"={c}{idx['gp']}-{c}{idx['oi']}"
    row("opx", "營業費用（行銷、研發、管理、攤銷、重組）", "US$B", ox, M3, "歷史＝10-K 五項費用；預測＝毛利 − 營業利益（殘差）")
    row("oi", "營業利益", "US$B", {**{c: f"=Q!{S2Q[c]}21" for c in HC}, **{c: f"=PL!{S2P[c]}{oi}" for c in FC}}, M3, "GAAP 營業利益（PL 頁；歷史＝Q 頁）")
    row("opm", "營業利益率", "%", {c: f"={c}{idx['oi']}/{c}{idx['rev']}" for c in ALLC}, PCT)
    ot = {c: f"={c}{X['i_no']}-Q!{S2Q[c]}26" for c in HC}
    for c in FC:
        ot[c] = f"=-PL!{S2P[c]}{it}+PL!{S2P[c]}{ii}"
    row("oth", "其他收入（淨額，含利息費用）", "US$B", ot, M3, "歷史＝非營業收入淨額 − 利息費用；預測＝利息收入 − 利息費用（PL 頁；非營業損益設 0）")
    row("pti", "稅前利益", "US$B", {c: f"={c}{idx['oi']}+{c}{idx['oth']}" for c in ALLC}, M3)
    row("tax", "所得稅", "US$B", {**{c: f"={c}{X['i_tx']}" for c in HC}, **{c: f"=PL!{S2P[c]}{tx}" for c in FC}}, M3)
    row("ni", "稅後淨利", "US$B", {c: f"={c}{idx['pti']}-{c}{idx['tax']}" for c in ALLC}, M3, "稅前 − 所得稅（核對列：與 Q 頁、PL 頁淨利相同）")
    row("nm", "淨利率", "%", {c: f"={c}{idx['ni']}/{c}{idx['rev']}" for c in ALLC}, PCT)
    row("pref", "特別股股利", "US$B", {**{"E": "=Q!L48"}, **{c: f"=PL!{S2P[c]}{pf}" for c in FC}}, M3, "FY2026 起有強制轉換特別股；FY2024A、FY2025A 空白＝無")
    row("sh", "稀釋股數", "十億股", {**{c: f"=Q!{S2Q[c]}49" for c in HC}, **{c: f"=PL!{S2P[c]}{sh}" for c in FC}}, M3, "歷史＝10-K 稀釋加權平均；預測＝PL 頁（含新股與強制轉換特別股）")
    row("eps", "EPS（GAAP 稀釋）", "US$", {**{c: f"=Q!{S2Q[c]}50" for c in HC}, **{c: f"=PL!{S2P[c]}{eps}" for c in FC}}, "0.00", "歷史＝10-K 公告值；預測＝PL 頁（（淨利 − 特別股股利）÷ 稀釋股數）")
    IS.cell(r_ratio, 1, "比率").font = BOLD
    ny = {c: f"={c}{idx['ni']}/{prevc[c]}{idx['ni']}-1" for c in "DEFGHIJK"}
    row("r_niy", "淨利年增率（YoY）", "%", ny, PCT, "FY2026A 含非營業收入淨額 3.5（10-K），預測期不估，所以 FY2027E 年增率為負")
    IS.cell(r_chk, 1, "核對（應為 0；EPS 列因公告值四捨五入，歷史差在 0.01 內）").font = BOLD
    row("c1", "核對：毛利 − 營業費用 − 營業利益", "US$B", {c: f"={c}{idx['gp']}-{c}{idx['opx']}-{c}{idx['oi']}" for c in ALLC}, "0.000", "歷史：10-K 三項成本、五項費用與營業利益互相核對；預測以殘差定義，恆為 0")
    row("c2", "核對：稅前利益 − 10-K（歷史）／PL 頁（預測）", "US$B", {**{c: f"={c}{idx['pti']}-{c}{X['i_pt']}" for c in HC}, **{c: f"={c}{idx['pti']}-PL!{S2P[c]}{pt}" for c in FC}}, "0.000")
    row("c3", "核對：稅後淨利 − Q 頁（歷史）／PL 頁（預測）", "US$B", {**{c: f"={c}{idx['ni']}-Q!{S2Q[c]}27" for c in HC}, **{c: f"={c}{idx['ni']}-PL!{S2P[c]}{ni}" for c in FC}}, "0.000")
    row("c4", "核對：營收 − Q 頁（歷史）／PL 頁（預測）", "US$B", {**{c: f"={c}{idx['rev']}-Q!{S2Q[c]}10" for c in HC}, **{c: f"={c}{idx['rev']}-PL!{S2P[c]}{rv}" for c in FC}}, "0.000")
    row("c5", "核對：（淨利 − 特別股股利）÷ 稀釋股數 − EPS", "US$", {c: f"=({c}{idx['ni']}-N({c}{idx['pref']}))/{c}{idx['sh']}-{c}{idx['eps']}" for c in ALLC}, "0.000")
    IS.cell(r_seg, 1, "營收分部（US$B；圖表與驅動因子章用）").font = BOLD
    qs = lambda c, rows_: "=" + "+".join(f"Q!{S2Q[c]}{x}" for x in rows_)
    ps = lambda c, rows_: "=" + "+".join(f"PL!{S2P[c]}{x}" for x in rows_)
    row("s_oci", "Cloud Infrastructure（OCI）", "US$B", {**{c: qs(c, [4]) for c in HC}, **{c: ps(c, [PB["oci"]]) for c in FC}}, M3, "歷史＝新聞稿營收線（Q 頁）；預測＝PL 頁")
    row("s_ai", "　其中 OCI AI 算力（命題子列；歷史不拆）", "US$B", {c: ps(c, [PB["ai"]]) for c in FC}, M3, "公開資料沒有 OCI 的 AI／非 AI 拆分，所以歷史空白")
    row("s_saas", "Cloud Applications（SaaS）", "US$B", {**{c: qs(c, [5]) for c in HC}, **{c: ps(c, [PB["saas"]]) for c in FC}}, M3)
    row("s_sw", "Software license＋support", "US$B", {**{c: qs(c, [6, 7]) for c in HC}, **{c: ps(c, [PB["lic"], PB["sup"]]) for c in FC}}, M3)
    row("s_hw", "Hardware", "US$B", {**{c: qs(c, [8]) for c in HC}, **{c: ps(c, [PB["hw"]]) for c in FC}}, M3)
    row("s_svc", "Services", "US$B", {**{c: qs(c, [9]) for c in HC}, **{c: ps(c, [PB["svc"]]) for c in FC}}, M3)
    row("s_tot", "合計（OCI＋SaaS＋軟體＋硬體＋服務）", "US$B", {c: f"={c}{sidx['s_oci']}+{c}{sidx['s_saas']}+{c}{sidx['s_sw']}+{c}{sidx['s_hw']}+{c}{sidx['s_svc']}" for c in ALLC}, M3)
    row("s_chk", "核對：合計 − 營收（應為 0）", "US$B", {c: f"={c}{sidx['s_tot']}-{c}{idx['rev']}" for c in ALLC}, "0.000")
    IS.cell(r_in, 1, "FY2024–FY2026 實績輸入與假設（藍字＝輸入；取自 Oracle Form 10-K FY2026 合併損益表〔期末 2026-05-31〕，Verified，經查核）").font = BOLD
    for j, h in enumerate(["項目", "單位", "FY2024", "FY2025", "FY2026", "", "", "", "", "", "", "來源", "出處網址"], 1):
        c = IS.cell(r_in + 1, j, h); c.font = BOLD; c.fill = HDR
    for k, (labl, vals) in K10.items():
        rr = hidx[k]
        IS.cell(rr, 1, labl); IS.cell(rr, 2, "US$B")
        for col, v in zip(HC, vals):
            c = IS[f"{col}{rr}"]; c.value = v; c.number_format = M3; c.font = BLUE
        IS.cell(rr, 12, "Oracle Form 10-K FY2026 合併損益表（USD 百萬換算十億；Verified）"); IS.cell(rr, 13, U_10K)
    rr = hidx["h_cogsr"]
    IS.cell(rr, 1, "非 AI 營業成本率（＝（FY2024 營業成本 − 4 × 每季非 AI 折舊）÷ FY2024 營收；預測期營業成本假設）"); IS.cell(rr, 2, "%")
    IS[f"C{rr}"] = f"=(C{idx['cogs']}-4*nonai_dep_q)/C{idx['rev']}"; IS[f"C{rr}"].number_format = PCT
    IS.cell(rr, 12, "Derived：FY2024 為 AI 規模化前的年度（同 m_sw 的校準年）；每季非 AI 折舊（In 頁，FY2024 水準）自營業成本扣除，其餘占營收比例在預測期不變；Assumed（模型未分列營業成本）")
    dn(wb, "cogs_nonai_r", f"IS!$C${rr}")
    IS.column_dimensions["A"].width = 60
    IS.column_dimensions["B"].width = 9
    for col in ALLC:
        IS.column_dimensions[col].width = 12
    IS.column_dimensions["L"].width = 80
    IS.column_dimensions["M"].width = 40
    IS.freeze_panes = "C5"
    L.update(idx=idx, r_niy=r_niy, cidx=cidx, sidx=sidx, hidx=hidx, ratio=r_ratio, chk=r_chk, seg=r_seg, inp=r_in)
    return L


# =====================================================================================================
# Sens：加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；只新增列，白話標籤）
TOR_PLAIN = {
    "「OCI AI 以外」利潤率 m_sw": "「OCI AI 以外」的分部利潤率",
    "IF_HoldEcon 四世代同時（每 MW 年持有成本）": "四個 GPU 世代的每 MW 年持有成本同時取低／高",
    "長約價格證據 p_long（k_長約隨之）": "長約價格證據（長約定價倍數隨之）",
    "新增投資級債利率 r_new（瀑布補缺口的新債）": "新增投資級債利率（瀑布補缺口的新債）",
    "營運資金增加 ÷ 營收增量 nwc_r": "營運資金增加占營收增量的比率",
}


def sens_tor(wb, C):
    SE = wb["Sens"]
    r0 = SE.max_row + 3
    SE.cell(r0, 1, "加權目標價相對基準（HTML 龍捲風圖用；版面規格 v1.3 H 節；本節只引用上方各案例的相對基準，只新增、不改既有格；標籤為白話）").font = BOLD
    for j, h in enumerate(["輸入（白話）", "輸入取低值", "輸入取高值"], 1):
        c = SE.cell(r0 + 1, j, h); c.font = BOLD; c.fill = HDR
    r = r0 + 2
    a, b = C["SENS_ROWS"]
    for rr in range(a, b + 1):
        lab_ = str(SE.cell(rr, 1).value)
        plain = TOR_PLAIN.get(lab_, lab_)
        assert not PARAM_RE.search(plain), ("Sens 白話標籤含參數代號", plain)
        SE.cell(r, 1, plain)
        for col, src in (("B", "D"), ("C", "F")):
            SE[f"{col}{r}"] = f"=Sens!{src}{rr}"; SE[f"{col}{r}"].number_format = "+0.0%;-0.0%;0.0%"
        r += 1
    return r0


# =====================================================================================================
# Coord：三層協調（讀取層級＋C0 實績對帳；MAG 類樣板）
def coord(wb, C):
    TKS, TKV, GEN, GNAME, QR = C["TKS"], C["TKV"], C["GEN"], C["GNAME"], C["QR"]
    TK = TKS["items"]
    TB = C["TR"]["B"]
    W = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_c0_v0.3.2.json"), encoding="utf-8"))
    WN, WR, WS = W["names"], W["rows"], W["source"]
    TR = json.load(open(os.path.join(C["HERE"], "data", "wattonomics_trace_v0.3.2.json"), encoding="utf-8"))
    TRV = {x["cell"]: x["value"] for x in TR["rows"]}
    WV = f"Wattonomics {WS['version']}（{WS['file']}，sha256 {WS['sha256'][:8]}；快照 {WS['snapshot']}）"
    TKSRC = f"Tokenomics {TKV}（{TKS['source']['file']}，master {TKS['source']['commit'][:7]}；快照 oracle/data/tokenomics_snapshot_{TKV}.json）"
    ws = wb.create_sheet("Coord")
    ws.column_dimensions["A"].width = 64
    for c, w in zip("BCDEFGHIJKL", (30, 12, 16, 16, 12, 28, 60, 22, 60, 22, 50)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Coord — 三層協調：讀取層級與 C0（實績）對帳（v0.5.2 起；三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_第二批13家_v1.3與三層協調；MAG 類樣板）"
    ws["A1"].font = BOLD
    ws["A2"] = ("本頁只新增、不改任何既有格，也不改上游值；個股不裁決，差距只標示，證據回饋走對帳單（規則 C5）。Oracle 是算力賣方（OCI 出租 AI 算力、資本支出由自己出，機房向房東承租），不做兩讀數目標價；"
                "C0 的個股側只用公司已公布實績（10-Q／10-K 財報表），層級側用 Wattonomics v0.3.2 隱含值（玩家資本支出、上線 GW）與 Tokenomics 每 GW 營收錨，同源者標「參考」。"
                "會計年度 5 月底：2026 日曆年已公布期間＝FY26 Q3（2025-12～2026-02，含 2025-12 一個月，已扣除）、FY26 Q4、FY27 Q1。")
    r = 4
    ws.cell(r, 1, "F. 三層協調：讀取層級與 C0（實績）對帳").font = BOLD; r += 1
    ws.cell(r, 1, "F1. 本模型讀取的層級、版本與名稱（只列層級值與本模型採用值；同源，不判定，不計入對帳中）").font = BOLD; r += 1
    for j, h in enumerate(["名稱", "年度／世代", "層級值", "本模型採用值", "單位", "層級與版本", "說明"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    F1a = r
    ws.cell(r, 1, f"讀取層級：Tokenomics（算力單位經濟）——本模型讀 {TKV} 快照（工作單所列現行 v22.0 不在本輪換快照；換快照另走 Tokenomics 同步通知，不影響本輪 C0）；"
                  "Wattonomics——本模型不讀（無 WIF_ 名稱進模型），v0.3.2 快照只用於下方 C0 對帳列；反向：Wattonomics 讀本模型寫回的資本支出（見下列「反向讀取」）"); ws.cell(r, 1).font = BOLD; r += 1

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
    for g in GEN:
        tkrows.append((f"IF_HoldEcon｜{GNAME[g]}", GNAME[g], TK["IF_HoldEcon"]["values"][GNAME[g]]["基準"], f"=he_{g}", "US$m/MW-年（＝$B／GW／年）", "每 MW 年持有成本（經濟口徑；收入錨 k＝1）"))
    for g in GEN:
        tkrows.append((f"IF_CapexIT｜{GNAME[g]}", GNAME[g], TK["IF_CapexIT"]["values"][GNAME[g]]["基準"], f"=cx_{g}", "US$m/MW", "每 MW IT 資本支出（機房以租賃取得，不含）"))
    for g in GEN:
        tkrows.append((f"IF_OpexGW｜{GNAME[g]}", GNAME[g], TK["IF_OpexGW"]["values"][GNAME[g]]["基準"], f"=ox_{g}", "US$m/MW-年", "每 MW 年營運費用（不含折舊）"))
    tkrows.append(("IF_DeprLifeIT｜GB300 NVL72", "GB300 NVL72", TK["IF_DeprLifeIT"]["values"]["GB300 NVL72"]["基準"], "=life_ai", "年", "AI 設備折舊年限"))
    tkrows.append(("IF_GPUhrEcon｜Hopper H100", "Hopper H100", TK["IF_GPUhrEcon"]["values"]["Hopper H100"]["基準"], "=gh_h100", "US$/GPU-hr", "k 的現貨分母"))
    t0 = r
    for nm_, yr_, lv, ad, un, nt in tkrows:
        F1R[nm_] = f1(nm_, yr_, lv, ad, un, TKSRC, nt, "0.0000")
    F1b_ = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（Tokenomics 列；In 頁採用值四捨五入到四位小數，每列誤差小於 0.00005，合計應小於 0.001）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{t0}:D{F1b_}-C{t0}:C{F1b_}))").number_format = "0.000000"; F1chk = r
    ws.cell(r, 7, f'=IF(D{r}<0.001,"通過","不通過")'); r += 1
    ws.cell(r, 1, "Wattonomics v0.3.2 快照名稱與輸入列（只供 C0 對帳列；不進模型）").font = BOLD; r += 1
    wrows = {}
    for nm_, key, un, nt in (
            ("WIF_Gap_Player_orcl_2026_CapexAI", "capai", "US$B", "Oracle 2026 年 AI 資本支出（Interface!F648←Gap!G251←G229＝新增 GW × (每 GW IT 成本＋融資租賃機房成本)＋2027、2028 年廠房提前；GW 由 CapexGW 從公司資本支出（FY26 實績、FY27 指引中點）反推，見下方追公式）"),
            ("WIF_Gap_Player_orcl_2026_CapexNonAI", "capnai", "US$B", "非 AI 資本支出（Gap!G236＝CapexGW!G142 日曆年資本支出流量 − G145 AI 流量；兩者都是公司資本支出 FY26 實績與 FY27 指引中點的加權，直接同源）"),
            ("WIF_Gap_Player_orcl_2026_AddGW", "addgw", "GW", "2026 年新增 IT GW（Interface!F656←Gap!G222←CapexGW!G158←G156←G155←G154←G153←G145←G144←G141←G137←G135←In!K193；輸入為公司 FY26 資本支出實績、FY27 指引中點與本模型 2027–28 寫回，見追公式；不是 Epoch 或園區證據推得）"),
            ("WIF_PlayerGW_YE_orcl_2025", "ye25", "GW", "Oracle 2025 年底上線 IT GW（Interface!E28←CapexGW!F159←Players!F106←F100←F99＝In!J102 晶片功耗 650.2 MW × k_chip2it 1.85；Epoch 晶片功耗 × 換算係數，Analogy；與公司資本支出無關，不同源）"),
            ("WIF_PlayerGW_YE_orcl_2026", "ye26", "GW", "Oracle 2026 年底上線 IT GW（Interface!F28←CapexGW!G159＝2025 年底（Epoch，不同源）＋2026 新增 GW（同源於公司資本支出，見追公式））"),
            ("WIF_PlayerGW_YE_orcl_SiteOnly_2026", "ye26s", "GW", "2026 年底上線 IT GW（只有園區證據路徑；Pipeline 頁園區加總，不經資本支出）"),
            ("WIF_Gap_Player_orcl_2027_CapexAIMC", "cap27", "US$B", "並列讀法：2027 年 AI 資本支出（個股模型寫回值；2026 年無此值）")):
        yr_ = re.search(r"(20\d\d)", nm_).group(1)
        wrows[key] = f1(nm_, yr_, WN[nm_], None, un, WV, nt, "0.000")
    A26 = WR["act26_orcl"]["values"][1]
    wrows["act26"] = f1("In 列 act26_orcl｜2026 已公布資本支出實績（Q1–Q3；Dec 2025–Aug 2026）", "2025-12～2026-08", A26, f"=Q!G{QR['capex']}+Q!H{QR['capex']}+Q!I{QR['capex']}", "US$B", WV,
                        "＝FY26 Q3＋FY26 Q4＋FY27 Q1 現金資本支出（含 2025-12 一個月）；Wattonomics 只在 Compare 頁讀（不進 GW 或資本支出路徑，見追公式）；採用值＝本模型 Q 頁三季合計", "0.000")
    wrows["cap25"] = f1("In 列 cap_orcl｜FY2025 資本支出（J125）", "FY2025", WR["cap_orcl"]["values"][2], f"=Q!K{QR['capex']}", "US$B", WV,
                        "10-K 現金資本支出（Verified）；Wattonomics 寫入的公司實績；採用值＝本模型 Q 頁 FY2025", "0.000")
    wrows["gcap26"] = f1("In 列 gcap_orcl｜FY2026 資本支出合計（K193）", "FY2026", WR["gcap_orcl"]["values"][3], f"=Q!L{QR['capex']}", "US$B", WV,
                         "＝本模型 FY2026 實績（10-K 現金資本支出）寫回；是 CapexGW 2026 年資本支出流量（5 ÷ 12 權重）的輸入", "0.000")
    wrows["guid"] = f1("In 列 gcap_orcl｜FY2027 資本支出合計（L193）", "FY2027", WR["gcap_orcl"]["values"][4], "=capex_g27", "US$B", WV,
                       "＝公司 FY27 資本支出指引中點（90–95；本模型 In 頁 capex_g27）寫回；是 CapexGW 2026 年資本支出流量（7 ÷ 12 權重）的輸入；Interested-party（指引不是實績）", "0.0")
    wrows["lease"] = f1("In 列 lease_orcl｜承租比例", "2026", WR["lease_orcl"]["values"][1], None, "比", WV, "廠房、土地、電力由房東出資的占比（Oracle 管理層法說：資本支出不含土地、建物、電力；單一來源）", "0.00")
    wrows["finl"] = f1("In 列 finl_orcl｜承租廠房中融資租賃占比", "2026", WR["finl_orcl"]["values"][1], None, "比", WV, "FY27Q1 帳上融資租賃負債 9.185 ÷（營業租賃 34.621＋9.185）＝0.21；層級資本支出因此含 21% 的機房成本（融資租賃使用權資產視同資本支出）", "0.00")
    wrows["obs"] = f1("In 列 obs_orcl｜Oracle 已連網 MW-IT（FY27Q1 末，2026-08-31）", "2026", WR["obs_orcl"]["values"][1], f"=Thesis!C{TB['conn']}", "MW-IT", WV,
                      "＝FY26 末 1,835＋Q1 交付 850 ÷ 1.2（Derived；換算係數為假設）；本模型 Thesis 頁 Q1 末已連網 MW 同值（同源）；追公式顯示 obs_orcl 不在 2026 年資本支出與 GW 的有效路徑上", "#,##0")
    r += 1
    ws.cell(r, 1, "反向讀取：Wattonomics 讀入本模型寫回的資本支出（In 頁 gcap_orcl、gcapai_orcl、gnai_orcl；快照值 vs 本模型情境 B 現值；同源，不判定）").font = BOLD; r += 1
    CB = C["CR"]["B"]
    rb0 = r
    for key, nm_, crow in (("gcap_orcl", "資本支出合計", CB["cx"]), ("gcapai_orcl", "AI 資本支出", CB["cxai"]), ("gnai_orcl", "非 AI 資本支出", CB["cxna"])):
        vals = WR[key]["values"]
        for i, v in enumerate(vals):
            if v is None:
                continue
            yr = 2023 + i
            col = "CDEFGHI"[yr - 2026]
            f1(f"In 列 {key}｜{nm_}", f"FY{yr}", v, f"=CF!{col}{crow}", "US$B", WV, "Wattonomics 讀入的本模型寫回值（v0.4.2 版寫回）vs 本模型 CF 頁情境 B 現值", "0.000")
    rb1 = r - 1
    ws.cell(r, 1, "反向讀取核對：Σ|本模型現值 − Wattonomics 讀入值|（應為 0；不為 0 代表寫回值已過期，需走對帳單）")
    ws.cell(r, 4, f"=SUMPRODUCT(ABS(D{rb0}:D{rb1}-C{rb0}:C{rb1}))").number_format = "0.000000"; rbc = r
    ws.cell(r, 7, f'=IF(D{r}<0.001,"一致（寫回值未過期）","寫回值已過期")'); r += 1
    ws.cell(r, 1, "F1 末（Wattonomics 值與 In 頁採用值的核對）：In 列公司實績與本模型 Q 頁相同"); ws.cell(r, 4, f"=ABS(C{wrows['act26']}-D{wrows['act26']})+ABS(C{wrows['cap25']}-D{wrows['cap25']})+ABS(C{wrows['gcap26']}-D{wrows['gcap26']})+ABS(C{wrows['guid']}-D{wrows['guid']})").number_format = "0.000000"; rF1c = r; r += 1
    r += 1
    ws.cell(r, 1, "F1b. 追公式：Wattonomics 的 Oracle 路徑是否讀公司實績或目標（v0.3.2 快照對應的 20261114_Wattonomics_v0.3.2.xlsx；規則：有讀者的部分屬同源，不當獨立證據）").font = BOLD; r += 1
    ws.cell(r, 1, "追公式工具：tools/coord/trace_wattonomics.py --player orcl（分支 claude/trace-wattonomics，PR 520；輸出存 oracle/data/wattonomics_trace_v0.3.2.json）；下方傾印只列「生效」格與旗標格，條件讀取與未生效的分支（園區列等）不貼。")
    r += 1
    for j, h in enumerate(["輸出名稱 ← 公式鏈", "", "", "", "", "", "同源判定", "讀到什麼（輸入、公式、儲存格）"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    f126 = TRV.get("CapexGW!G142")
    TRC = [
        ("WIF_Gap_Player_orcl_2026_AddGW ← Interface!F656 ← Gap!G222 ← CapexGW!G158 ← G156（上線時差 lag_on 加權前一年）← G155 ← G154（＝資金 ÷（每 GW IT 成本＋融資租賃機房成本））← G153 ← G145（日曆年 AI 資本支出＝FY26 × 5/12＋FY27 × 7/12）← G144（＝總資本支出 × AI 占比）← G141 ← G137 ← G135 ← In!K193",
         "同源（公司 FY26 實績＋FY27 指引中點＋本模型 2027–28 寫回）",
         "公司 FY26 實績：In!K193（gcap_orcl＝55.663，10-K 現金資本支出）；公司 FY27 指引中點：In!L193（92.5，本模型寫回）；AI 占比：FY26 用 ai_sh（In!I13，0.85，人工確認不同源、Assumed），FY27 用本模型寫回的 AI ÷ 合計（In!L191 ÷ L193）。"
         "再扣除 2025 年已入帳的資本支出（CapexGW!F 欄，讀 In!J125 cap_orcl＝FY25 實績 21.215）。GW＝資金 ÷ 每 GW 成本（Tokenomics），所以 2026 年新增 GW 是公司資本支出換算，不是獨立的物理證據"),
        ("WIF_Gap_Player_orcl_2026_CapexAI ← Interface!F648 ← Gap!G251 ← G229（＝G226＋G227＋G228）", "同源（同上）",
         "G226＝2026 新增 GW（上列）×（每 GW IT 成本 33.065＋機房成本在資本支出的部分 2.566＝機房 12.217 × (1 − lease_orcl × (1 − finl_orcl))）；G227、G228＝2027、2028 新增 GW 的廠房提前入帳（lag_bldg 18 個月），讀 CapexGW!H158、I158，其輸入是本模型寫回的 FY27–FY28 資本支出（In!L193、M193、L191、M191）。lease_orcl 與 finl_orcl 為人工確認（Oracle 管理層法說、10-Q 租賃負債）"),
        ("WIF_Gap_Player_orcl_2026_CapexNonAI ← Interface!F649 ← Gap!G236 ← CapexGW!G142 − G145", "同源（公司實績與指引）",
         "G236＝CapexGW!G142（日曆年資本支出流量＝FY26 實績 5/12＋FY27 指引中點 7/12" + (f"＝{f126:.2f}" if isinstance(f126, (int, float)) else "") + "）減 G145（日曆年 AI 流量）；層級的非 AI 資本支出是公司總資本支出與 AI 占比假設（0.85）的差，不是獨立證據"),
        ("WIF_PlayerGW_YE_orcl_2025 ← Interface!E28 ← CapexGW!F159 ← Players!F106 ← F100 ← F99 ← In!J102 × k_chip2it（In!I7，1.85）", "不同源（Epoch 晶片功耗 × 換算係數）",
         "2025 年底上線 GW＝In!J102 chip_orcl（650.2 MW，Epoch 晶片功耗，Analogy）× 1.85；Players!F100 的觀測分支（obs_orcl，In!I28）因觀測年為 2026 而不啟用；與公司資本支出無關"),
        ("WIF_PlayerGW_YE_orcl_2026 ← Interface!F28 ← CapexGW!G159＝F159＋G158", "部分同源",
         "2025 年底（Epoch，不同源）＋2026 年增量 G158（同源於公司資本支出，見第一列）。另有 SiteOnly 路徑 WIF_PlayerGW_YE_orcl_SiteOnly_2026（園區證據加總），與本路徑相差很小，但公司不揭露 GW，仍不能做 C0（見 F2 MW 列）"),
        ("act26_orcl（Wattonomics In 頁 act26_orcl；2025-12 至 2026-08 資本支出實績 63.627）", "未進 GW／資本支出路徑",
         "反向查詢：公式讀 act26_orcl 的只有 " + "、".join(sorted({x.split("!")[0] for x in TR["act26_readers"]})) + " 頁（" + str(len(TR["act26_readers"])) + " 格；Compare 頁第 243 列），不回饋 CapexGW、Gap、Interface 的 Oracle 值；本表直接以 10-Q 與 10-K 的同一實績為個股側"),
        ("WIF_Gap_Player_orcl_2026_CapexAIModel ← Interface!F653 ← Gap!G240", "空白（無回寫值）",
         "2026 年個股模型 AI 資本支出 In!K191 為空（本模型 FY2026A 不拆 AI／非 AI），所以這一欄不替代 GW 推算；2027 年起才讀本模型寫回（並列讀法，本表不用）"),
    ]
    for a_, j_, t_ in TRC:
        ws.cell(r, 1, a_); ws.cell(r, 7, j_); ws.cell(r, 8, t_); r += 1
    ws.cell(r, 1, "追公式結論：Wattonomics 2026 年 Oracle 的資本支出（AI＋非 AI）與新增 GW 由公司 FY26 資本支出實績、FY27 指引中點（日曆年權重 5/12、7/12）與本模型 2027–28 寫回換算，整列同源；"
                  "2025 年底 GW 來自 Epoch 晶片功耗，不同源；公司實績 act26_orcl 只在 Compare 頁讀，未進路徑。層級另有 SiteOnly 路徑，但公司不揭露 GW，每 GW 資本支出也不可比；"
                  "因此資本支出列標「參考（同源校準）」，結論「無可比實績（層級值同源）」。若 Wattonomics 改為只用 Epoch／園區證據推 2026 GW，才恢復比較（證據回饋走對帳單 C5）。")
    rtr_conc = r; r += 1
    eff = [x for x in TR["rows"] if "生效" in x["roots"].values()]
    keepc = {x["cell"] for x in eff}
    flg = [x for x in TR["rows"] if x["flag"] and x["flag"] not in ("園區列", "園區輸入列（人工確認：多為公司公告）", "公司公告園區（人工確認）") and x["cell"] not in keepc]
    seen, dump = set(), []
    for x in eff + flg:
        if x["cell"] in seen:
            continue
        seen.add(x["cell"]); dump.append(x)
    SHORT = {"WIF_Gap_Player_orcl_2026_AddGW": "新增GW", "WIF_Gap_Player_orcl_2026_CapexAI": "AI資本支出", "WIF_Gap_Player_orcl_2026_CapexNonAI": "非AI資本支出",
             "WIF_PlayerGW_YE_orcl_2025": "2025底GW", "WIF_PlayerGW_YE_orcl_2026": "2026底GW", "WIF_Gap_Player_orcl_2026_CapexAIModel": "個股回寫"}
    def stat(x):
        ok = [SHORT.get(k, k) for k, v in x["roots"].items() if v == "生效"]
        return ("生效（" + "、".join(ok) + "）") if ok else x["status"]
    ws.cell(r, 1, f"追公式傾印（trace_wattonomics.py 的輸出，只列生效格 {len(eff)} 與其餘旗標格 {len(dump) - len(eff)}；每格公式與值；來源 {TR['source'].split('/')[-1]}；縮排＝依賴深度）").font = BOLD
    ws.cell(r, 7, "標記｜狀態"); ws.cell(r, 8, "公式（省略開頭等號）"); ws.cell(r, 4, "值"); r += 1
    rdump0 = r
    for x in dump:
        ws.cell(r, 1, ("　" * min(x["depth"], 12)) + x["cell"] + (("｜" + x["label"]) if x["label"] else ""))
        ws.cell(r, 4, x["value"] if isinstance(x["value"], (int, float)) else (None if x["value"] is None else str(x["value"])))
        ws.cell(r, 7, (x["flag"] + "｜" if x["flag"] else "") + stat(x)); ws.cell(r, 8, x["formula"][1:] if x["formula"] else "（輸入或常數）")
        r += 1
    rdump1 = r - 1
    r += 1
    # ------------------------------------------------------------------ F2
    ws.cell(r, 1, "F2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；個股側＝10-Q／10-K 財報表實績，不含指引與模型預測；層級側＝Wattonomics 隱含值）").font = BOLD; F2h = r; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE
    ws.cell(r, 8, "規則：兩個讀數的差距都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致。門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    for j, h in enumerate(["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級", "出處網址"], 1):
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
    F2hh = r; r += 1
    F2a = r
    cx = QR["capex"]
    SRC3 = "Oracle Form 10-Q（季末 2026-02-28）現金流量表 9 個月「Capital expenditures」39.170 減上半年（FY26 Q1＋Q2＝20.535）；Q 頁第 " + str(cx) + " 列；2026-03 新聞稿同值。10-Q 網址未取得（SEC EDGAR 在建置環境被擋；數字以 10-Q 轉載與新聞稿交叉）"
    SRC4 = "Oracle Form 10-K FY2026 現金流量表全年「Capital expenditures」55.663 減 10-Q 前三季 39.170；Q 頁第 " + str(cx) + " 列；新聞稿同值"
    SRC1 = "Oracle Form 10-Q（季末 2026-08-31）現金流量表「Capital expenditures」28.499（本輪查證；去年同季 8.502）；Q 頁第 " + str(cx) + " 列"
    GR = "Verified（10-Q／10-K 財報表；10-Q 為核閱未查核，標 Verified 為財報表原值）"

    def inp(label, per, unit, f_act, src, grade, url="", fmt=M3):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, per); ws.cell(r, 3, unit)
        ws.cell(r, 4, f_act).number_format = fmt
        ws.cell(r, 8, src); ws.cell(r, 9, grade); ws.cell(r, 12, url)
        r += 1
        return r - 1
    rq3 = inp("資本支出 FY26 Q3（公司實績；含 2025-12 一個月）", "2025-12～2026-02（FY26 Q3）", "US$B", f"=Q!G{cx}", SRC3, GR)
    rq4 = inp("資本支出 FY26 Q4（公司實績）", "2026-03～2026-05（FY26 Q4）", "US$B", f"=Q!H{cx}", SRC4, GR, U_10K)
    rq1 = inp("資本支出 FY27 Q1（公司實績）", "2026-06～2026-08（FY27 Q1）", "US$B", f"=Q!I{cx}", SRC1, GR, U_10Q_Q1)
    rq9 = inp("已公布三季合計（2025-12～2026-08，9 個月）", "FY26 Q3＋Q4＋FY27 Q1", "US$B", f"=D{rq3}+D{rq4}+D{rq1}", "三季加總（含 2025-12，尚未扣除）", GR)
    ws.cell(r, 1, "核對：三季合計 − Wattonomics act26_orcl（應為 0；層級讀的是同一個三季合計）"); ws.cell(r, 4, f"=D{rq9}-C{wrows['act26']}").number_format = "0.000000"; rchk = r; r += 1
    rdec = inp("扣除 2025-12（FY26 Q3 × 1 ÷ 3；公司不揭露月度資本支出，線性分攤）", "2025-12（1 個月）", "US$B", f"=D{rq3}/3", "Assumed：FY26 Q3 三個月（2025-12、2026-01、2026-02）等分；資本支出逐季上升，12 月占比可能略低於三分之一；占比 25%–40% 時 8 個月累計約 56.2–59.0（基準 57.4），不改結論", "Assumed（線性分攤）")
    r8 = inp("2026-01～2026-08 累計（8 個月；日曆年已公布期間）", "2026-01～2026-08", "US$B", f"=D{rq9}-D{rdec}", "三季合計扣除 2025-12", "Derived（財報表加線性時差）")
    rl = r
    ws.cell(r, 1, "層級 CY2026 玩家資本支出：AI＋非 AI（Wattonomics）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 5, f"=C{wrows['capai']}+C{wrows['capnai']}").number_format = M3
    ws.cell(r, 10, f"{WV}：Interface 頁 Oracle 2026 AI 資本支出（新增 {WN['WIF_Gap_Player_orcl_2026_AddGW']:.2f} GW × 每 GW 成本，含融資租賃機房與廠房提前）＋非 AI 資本支出（連 F1 兩列）；"
                  "輸入皆來自公司 FY26 資本支出實績與 FY27 指引中點（日曆年權重 5/12、7/12），屬同源（見追公式）")
    ws.cell(r, 11, "Derived；輸入同源於公司實績與指引"); rl = r; r += 1
    rli = r
    ws.cell(r, 1, "層級輸入側：CY2026 資本支出流量（＝FY26 實績 × 5 ÷ 12＋FY27 指引中點 × 7 ÷ 12；CapexGW!G142）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 5, f"=5/12*Q!L{cx}+7/12*capex_g27").number_format = M3
    ws.cell(r, 10, "Wattonomics CapexGW!G142；本列以公式重現（FY26 實績＝Q 頁、FY27 指引中點＝In 頁 capex_g27），與 Wattonomics 值核對見下列；整列由公司數字組成"); ws.cell(r, 11, "Derived（同源）"); r += 1
    ws.cell(r, 1, "核對：本列重現值 − Wattonomics CapexGW!G142（追公式傾印值；應為 0）"); ws.cell(r, 5, f"=E{rli}-{f126 if isinstance(f126, (int, float)) else 'NA()'}").number_format = "0.000000"; r += 1
    ws.cell(r, 1, "主讀數：已公布期間累計（2026-01～08，8 個月）vs 層級 CY2026 × 8 ÷ 12"); ws.cell(r, 2, "8 個月累計實績；層級 × 8/12（線性時差近似）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{r8}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}*8/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 8, "8 個月累計＝FY26 Q3 × 2/3＋FY26 Q4＋FY27 Q1（上列）"); ws.cell(r, 9, GR); ws.cell(r, 10, "層級 CY2026 × 8 ÷ 12（Derived；線性時差近似）"); ws.cell(r, 11, "Derived")
    rd1 = r; r += 1
    ws.cell(r, 1, "第二讀數：最近一季（FY27 Q1，2026-06～08）× 4 vs 層級 CY2026"); ws.cell(r, 2, "FY27 Q1 × 4（約等於該季中點的年化速率）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{rq1}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 8, "FY27 Q1 單季 × 4（Derived）"); ws.cell(r, 9, GR); ws.cell(r, 10, "層級 CY2026 全年（同上）"); ws.cell(r, 11, "Derived")
    rd2 = r; r += 1
    ws.cell(r, 1, "口徑換算：融資租賃新增使用權資產（FY27 Q1；非現金，10-Q 附註）"); ws.cell(r, 2, "2026-06～2026-08"); ws.cell(r, 3, "US$B")
    c = ws.cell(r, 4, 1.539); c.number_format = M3; c.font = BLUE
    ws.cell(r, 8, "Oracle Form 10-Q（季末 2026-08-31）補充現金流資訊「ROU assets obtained in exchange for finance lease liabilities」1.539；營業租賃 4.903 為房東出資，不計入（層級也以房東出資處理）"); ws.cell(r, 9, "Verified（10-Q 附註；核閱未查核）"); ws.cell(r, 12, U_10Q_Q1)
    rfl = r; r += 1
    ws.cell(r, 1, "敏感度：第二讀數加回融資租賃新增（（FY27 Q1 資本支出＋融資租賃新增）× 4 vs 層級 CY2026）"); ws.cell(r, 2, "（FY27 Q1＋融資租賃新增）× 4"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=(D{rq1}+D{rfl})*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 7, "參考（口徑敏感度，不判定）"); ws.cell(r, 8, "層級的資本支出含 21% 機房成本（融資租賃視同資本支出，finl_orcl），公司現金資本支出不含；加回後差距擴大約 10 個百分點，不改結論"); ws.cell(r, 9, "Derived"); ws.cell(r, 10, "同上"); ws.cell(r, 11, "Derived")
    rfs = r; r += 1
    ws.cell(r, 1, "C0（實績）判定：資本支出").font = BOLD
    ws.cell(r, 2, "層級值同源於公司資本支出實績與指引（見追公式）：參考（同源校準），不判定")
    ws.cell(r, 7, "無可比實績（層級值同源）")
    ws.cell(r, 8, "規則：層級值由公司 FY26 實績與 FY27 指引中點換算＝同源，不是獨立證據；上方讀數仍列出供參考"); ws.cell(r, 9, GR); ws.cell(r, 10, "Wattonomics CapexGW／Gap（追公式列）"); ws.cell(r, 11, "Derived（同源校準）")
    rj = r; r += 1
    ws.cell(r, 1, "參考：若層級值是獨立證據，兩讀數規則會得到什麼（不計入判定）")
    ws.cell(r, 7, f'=IF(AND(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL},SIGN(F{rd1})=SIGN(F{rd2})),"參考（同源校準）：兩讀數都超過 ±門檻且同號",IF(OR(ABS(F{rd1})>{TOL},ABS(F{rd2})>{TOL}),"參考（同源校準）：只一個讀數超過門檻（時差邊界）","參考（同源校準）：兩讀數都在門檻內"))')
    rjr = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：層級輸入側流量 × 8 ÷ 12 vs 已公布 8 個月累計（兩邊同源：層級輸入就是公司實績與指引）"); ws.cell(r, 2, "8 個月累計 vs 輸入側 × 8/12"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=D{r8}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rli}*8/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 7, "參考（同源，不判定）"); ws.cell(r, 8, "層級輸入側＝FY26 實績 5/12＋FY27 指引中點 7/12（流量）"); ws.cell(r, 9, GR); ws.cell(r, 10, "Wattonomics CapexGW!G142"); ws.cell(r, 11, "Derived（同源）")
    rg = r; r += 1
    ws.cell(r, 1, "對照（參考，不判定）：層級輸入側流量 vs 層級輸出（Interface 頁 AI＋非 AI；兩者差來自 GW 換算與每 GW 成本）"); ws.cell(r, 2, "CY2026（全年）"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=E{rli}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{rl}").number_format = "#,##0.0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics CapexGW!G142"); ws.cell(r, 10, "Wattonomics Interface F648＋F649"); ws.cell(r, 11, "Derived")
    rb_ = r; r += 1
    # ---- OCI 營收（參考；口徑不同）----
    ws.cell(r, 1, "層級隱含 OCI AI 營收（CY2026；參考）：年均上線 GW × 每 GW 年營收錨（k＝1）").font = BOLD; r += 1
    ra1 = r; ws.cell(r, 1, "年均上線 IT GW（2025、2026 年底平均；Wattonomics）"); ws.cell(r, 3, "GW"); ws.cell(r, 5, f"=(C{wrows['ye25']}+C{wrows['ye26']})/2").number_format = "0.000"; ws.cell(r, 10, "Wattonomics PlayerGW_YE orcl 2025、2026（連 F1）；2026 年增量由公司資本支出換算，屬同源（追公式）"); ws.cell(r, 11, "Derived（部分同源）"); r += 1
    ra2 = r; ws.cell(r, 1, "每 GW 年營收錨（k＝1；Tokenomics IF_HoldEcon × 本模型 Q1 末世代組合；Thesis 頁每 MW 錨）"); ws.cell(r, 3, "US$B／GW"); ws.cell(r, 5, f"=Thesis!C{TB['anc']}").number_format = "0.000"
    ws.cell(r, 10, "Tokenomics IF_HoldEcon（連 F1）× 本模型 Q1 末世代組合（In 頁，Assumed）"); ws.cell(r, 11, "Derived／Assumed"); r += 1
    ra3 = r; ws.cell(r, 1, "層級隱含 OCI AI 營收（CY2026；US$B）"); ws.cell(r, 3, "US$B"); ws.cell(r, 5, f"=E{ra1}*E{ra2}").number_format = "#,##0.0"; ws.cell(r, 10, "上列兩項相乘（Oracle 出租 AI 算力，沒有自用比例；k 取 1）"); ws.cell(r, 11, "Derived"); r += 1
    oc = QR.get("oci", 4)
    ws.cell(r, 1, "OCI（Cloud Infrastructure）營收：2026-01～08 累計（公司實績；含非 AI 與 AI）"); ws.cell(r, 2, "FY26 Q3 × 2/3＋Q4＋FY27 Q1"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!G{oc}*2/3+Q!H{oc}+Q!I{oc}").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{ra3}*8/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（口徑不同，不判定）")
    ws.cell(r, 8, "Oracle Form 10-Q／新聞稿 Cloud infrastructure（IaaS）營收（Q 頁第 " + str(oc) + " 列；FY27 Q1 為 10-Q 值 7.388）；公司不揭露 AI 營收，OCI 含非 AI"); ws.cell(r, 9, "Verified（FY27 Q1，10-Q）／Interested-party（其餘季，新聞稿）")
    ws.cell(r, 10, "層級隱含 OCI AI 營收 × 8/12；F 欄＝層級隱含 AI 營收占 OCI 實績比例"); ws.cell(r, 11, "Derived／Assumed"); ws.cell(r, 12, U_10Q_Q1); rc1 = r; r += 1
    ws.cell(r, 1, "OCI 營收：FY27 Q1 × 4（公司實績年化）"); ws.cell(r, 2, "FY27 Q1 × 4"); ws.cell(r, 3, "US$B")
    ws.cell(r, 4, f"=Q!I{oc}*4").number_format = "#,##0.0"; ws.cell(r, 5, f"=E{ra3}").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（口徑不同，不判定）"); ws.cell(r, 8, "同上（FY27 Q1 單季 × 4；10-Q）"); ws.cell(r, 9, "Verified（10-Q）"); ws.cell(r, 10, "層級隱含 OCI AI 營收 CY2026"); ws.cell(r, 11, "Derived／Assumed"); ws.cell(r, 12, U_10Q_Q1); rc2 = r; r += 1
    ws.cell(r, 1, "CPU 與 GPU 營收（公司簡報揭露；含非 AI 的 CPU）：2026-03～08（FY26 Q4＋FY27 Q1）"); ws.cell(r, 2, "2026-03～2026-08（FY26 Q3 未取得）"); ws.cell(r, 3, "US$B")
    c = ws.cell(r, 4, "=4.8+6.5"); c.number_format = "#,##0.0"; c.font = BLUE
    ws.cell(r, 5, f"=E{ra3}*6/12").number_format = "#,##0.0"; ws.cell(r, 6, f"=E{r}/D{r}").number_format = "0%"
    ws.cell(r, 7, "參考（口徑不同，不判定）"); ws.cell(r, 8, "Oracle FY26 Q4 與 FY27 Q1 財報簡報（IR 網站）「CPU & GPU」營收 4.8、6.5（oracle/data/oracle_facts_20261007.json rev.cpuGpu；簡報未定義，含非 AI 的 CPU 雲端）；FY26 Q3 簡報未取得，所以只有 6 個月")
    ws.cell(r, 9, "Interested-party（簡報）"); ws.cell(r, 10, "層級隱含 OCI AI 營收 × 6/12"); ws.cell(r, 11, "Derived／Assumed"); rc3 = r; r += 1
    ws.cell(r, 1, "上限檢查（層級隱含 AI 營收不得高於 OCI 實績與 CPU＋GPU 實績）"); ws.cell(r, 7, f'=IF(AND(E{rc1}<=D{rc1},E{rc2}<=D{rc2},E{rc3}<=D{rc3}),"通過","不通過：層級隱含 AI 營收高於公司揭露的上限")'); rcu = r; r += 1
    # ---- 在役 MW（參考；口徑不明）----
    ws.cell(r, 1, "已連網 MW-IT（公司揭露換算）：2026-08-31"); ws.cell(r, 2, "2026-08-31（FY27 Q1 末）"); ws.cell(r, 3, "MW-IT")
    ws.cell(r, 4, f"=Thesis!C{TB['conn']}").number_format = "#,##0"
    ws.cell(r, 5, f"=(C{wrows['ye25']}+(C{wrows['ye26']}-C{wrows['ye25']})*8/12)*1000").number_format = "#,##0"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 7, "參考（口徑不明，不判定）")
    ws.cell(r, 8, "Oracle 簡報與法說揭露交付量（FY26 +1.2 GW；FY27 Q1 +850 MW，未說明 IT 或設施電力）；本模型換算 FY26 末 1,835＋850 ÷ 1.2＝2,543 MW-IT（Derived；÷ 1.2 為假設）；公司不揭露在役 MW 總量")
    ws.cell(r, 9, "Derived（換算係數 Assumed）"); ws.cell(r, 10, "層級 2025、2026 年底上線 GW 在 2026-08-31 的線性內插；2025 年底來自 Epoch 晶片功耗（不同源），2026 年增量來自資本支出（同源）"); ws.cell(r, 11, "Derived（部分同源）"); rmw = r; r += 1
    ws.cell(r, 1, "AI 營收與簽約電力（公司實績）"); ws.cell(r, 2, "無可比實績"); ws.cell(r, 7, "無可比實績")
    ws.cell(r, 8, "Oracle 不揭露 AI 營收，也不揭露在役或簽約 AI 電力總量（只揭露新增交付量與 RPO 664；RPO 不是營收）；Epoch 與園區估計是第三方估計，不算公司實績")
    ws.cell(r, 9, "—"); ws.cell(r, 10, "層級隱含值已列於上方 OCI AI 營收參考列"); ws.cell(r, 11, "Derived"); rmw2 = r; r += 1
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（F2 的 C0（實績）判定）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); rn = r; r += 1
    ws.cell(r, 1, "時差邊界（參考）項目數"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"時差邊界（參考）")'); rnb = r; r += 1
    ws.cell(r, 1, "無可比實績列數（資本支出 1、AI 營收與簽約電力 1）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"無可比實績*")'); rnc = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用；公式帶數字）")
    ws.cell(r, 4, F(T("C0（實績）對帳：公司已公布資本支出（2026 年 1–8 月累計 "), f'TEXT(D{rd1},"0.0")', T(" 十億美元，已扣除 2025-12）與層級 CY2026 隱含值（"), f'TEXT(E{rl},"0.0")', T("；同期間換算 "), f'TEXT(E{rd1},"0.0")', T("）相差 "),
                    f'TEXT(F{rd1},"+0%;-0%")', T("（主讀數）；最近一季年化 "), f'TEXT(D{rd2},"0.0")', T(" 十億美元，相差 "), f'TEXT(F{rd2},"+0%;-0%")',
                    T("（第二讀數）。但 Wattonomics 的 Oracle 2026 年資本支出與新增 GW 由公司 FY26 實績與 FY27 指引中點換算（追公式：CapexGW!G142←In!K193、L193），屬同源，判定："), f'G{rj}',
                    T("（參考，不計入對帳中）。OCI AI 營收與在役電力沒有可比的公司實績（公司不揭露），不判定。個股不裁決，不改上游值；證據回饋走對帳單（C5）。")))
    rs = r; r += 1
    ws.cell(r, 1, "讀法：兩個讀數的偏誤方向（依季中點）")
    ws.cell(r, 2, F(T("資本支出是流量，逐季大幅上升（FY26 Q3 "), f'TEXT(D{rq3},"0.0")', T("、Q4 "), f'TEXT(D{rq4},"0.0")', T("、FY27 Q1 "), f'TEXT(D{rq1},"0.0")',
                     T("）。主讀數用 8 個月的平均速率（中點約 5 月初）代表全年，全年中點在 7 月初，所以在上升趨勢下主讀數偏低；第二讀數用 FY27 Q1 單季年化（中點約 7 月中），比全年中點晚約兩週，且季內仍在上升，所以偏高。"
                       "兩讀數相差 "), f'TEXT((F{rd2}-F{rd1})*100,"0")', T(" 個百分點，就是時點差的幅度。公司 CY2026 全年實績要等 FY27 Q2（2026-09～11）與 FY27 Q3（2026-12～2027-02）財報合併後才完整；"
                       "層級輸入側流量與層級輸出（Interface 頁）相差 "), f'TEXT(F{rb_},"+0%;-0%")', T("，來自 GW 換算、每 GW 成本與廠房提前入帳的時點。這裡全部用流量（資本支出）比較，不拿 9 月底存量對年底存量。")))
    rr_ = r; r += 1
    ws.cell(r, 1, "F2 同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
    ws.cell(r, 2, ("（一）資本支出列：追公式（F1b）顯示 Wattonomics 2026 年 Oracle AI 資本支出（Gap!G251）與非 AI 資本支出（Gap!G236）的共同輸入是公司 FY26 資本支出實績（In!K193＝55.663）與 FY27 指引中點（In!L193＝92.5，本模型寫回），"
                   "以日曆年權重 5/12、7/12 合成（CapexGW!G142）；新增 GW 也由此換算（CapexGW!G158）。層級值因此同源於公司實績與目標，不是獨立證據：整列標「參考（同源校準）」，結論「無可比實績（層級值同源）」；"
                   "每 GW 資本支出同樣無可比（公司不揭露 GW）。公司實績 act26_orcl 沒有進 Oracle 的 GW 或資本支出路徑（只在 Compare 頁讀）。"
                   "（二）OCI AI 營收列：用到上述 GW（部分同源）、Tokenomics 每 GW 營收錨（層級共用）與本模型 Q1 末世代組合；公司不揭露 AI 營收，OCI 又含非 AI，CPU＋GPU 簡報營收含非 AI 的 CPU，所以只列層級隱含 AI 營收占公司實績比例與上限檢查，不判定。"
                   "（三）在役 MW：2025 年底來自 Epoch（不同源），2026 年增量同源於資本支出，公司換算值又是層級 obs_orcl 的輸入但不進有效路徑，口徑不明，列參考。"
                   "（四）所以本頁 C0 在 Oracle 對資本支出無檢驗力：對帳中為 0 不代表兩層一致，只代表層級值是由公司實績與指引換算而來。要恢復檢驗，需 Wattonomics 改以 Epoch 或園區證據獨立推 2026 年 GW（SiteOnly 路徑可作）。"))
    rsame = r; r += 1
    ws.cell(r, 1, "口徑換算（層級 vs 公司資本支出）")
    ws.cell(r, 2, F(T("公司資本支出＝現金流量表「Capital expenditures」（購置不動產與設備，現金）；層級資本支出＝新增 GW × (每 GW IT 成本＋機房成本中的融資租賃部分)＋廠房提前。Oracle 的機房由房東出資（營業租賃，不進資本支出，兩邊都不計）；"
                     "融資租賃占承租機房 21%（層級）：層級資本支出含這部分，公司現金資本支出不含（融資租賃新增使用權資產為非現金，FY27 Q1 為 "), f'TEXT(D{rfl},"0.0")', T(" 十億美元，占該季現金資本支出 "), f'TEXT(D{rfl}/D{rq1},"0%")',
                     T("）。判斷：以現金資本支出為個股側（與 Wattonomics act26_orcl 同口徑，也是公司唯一經查核的季度資本支出），融資租賃不加回；加回只會讓第二讀數差距再擴大（見敏感度列），不改變結論。"
                       "另需注意：Oracle 另有第三方製造商安排的短期融資與客戶預付（淨資本支出非 GAAP 口徑），本頁不採用淨額。")))
    rk = r; r += 1
    ws.cell(r, 1, "層級之間（參考；不計入個股 C0）").font = BOLD; r += 1
    ws.cell(r, 1, "Wattonomics 2026 年底上線 GW：資本支出路徑 vs 園區證據路徑（SiteOnly）"); ws.cell(r, 3, "GW")
    ws.cell(r, 4, f"=C{wrows['ye26']}").number_format = "0.000"; ws.cell(r, 5, f"=C{wrows['ye26s']}").number_format = "0.000"; ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = SPCT
    ws.cell(r, 7, "參考（層級之間，不判定）"); ws.cell(r, 8, "Wattonomics PlayerGW_YE orcl 2026（資本支出路徑，同源於公司實績與指引）"); ws.cell(r, 10, "Wattonomics PlayerGW_YE orcl SiteOnly 2026（園區證據）"); rly = r; r += 1
    ws.cell(r, 1, "外部估計（賣方、研調、指引）只列參考，不判定：公司 FY27 資本支出指引（90–95）、RPO 664 等不進 C0。"); r += 1
    ws.cell(r, 1, "版本與快照狀態：Tokenomics 本模型讀 v21.10（工作單所列現行 v22.0 未換）；Wattonomics v0.3.2 也讀 Tokenomics，版本由該層自行維護，本頁只標示。Oracle 模型值曾被 Wattonomics 讀入（資本支出寫回），本表只列層級值與追公式，不用該回饋判定；反向讀取核對見 F1。")
    r += 1
    return dict(F1=(F1a, F1b_), F1chk=F1chk, F1c=rF1c, rb=(rb0, rb1), rbc=rbc, wrows=wrows, F1R=F1R, F2h=F2h, F2hh=F2hh, F2=(F2a, F2b), rq=(rq3, rq4, rq1, rq9), rchk=rchk, rdec=rdec, r8=r8, rl=rl, rli=rli,
                rd1=rd1, rd2=rd2, rfl=rfl, rfs=rfs, rj=rj, rjr=rjr, rg=rg, rb_=rb_, ra=(ra1, ra2, ra3), rc1=rc1, rc2=rc2, rc3=rc3, rcu=rcu, rmw=rmw, rmw2=rmw2,
                rn=rn, rnb=rnb, rnc=rnc, rs=rs, rr=rr_, rsame=rsame, rk=rk, rly=rly, rtr=rtr_conc, rdump=(rdump0, rdump1), tol=TOL, f1h=None)


# =====================================================================================================
# README：v1.3 展示層白話列
def readme13(wb, C, PLA, CRD):
    RD = wb["README"]
    CON = C["CON"]
    v = lambda k, c="C": f"Val!{c}{CON[k]}"
    tp, up, rec, dcf, cmpv, cons = v("tp"), v("up"), v("rec"), v("dcf"), v("comp"), v("cons")
    TB = C["TR"]["B"]
    out = {}
    r = RD.max_row + 2
    RD.cell(r, 1, "v1.3 展示層（HTML 用；v0.5.2 起，個股 HTML 版面規格 v1.3：橫幅說明 ≤200 字、重點觀察 ≤5 點且 ≤750 字、不出現參數代號；上方原列保留給下游）").font = BOLD; r += 1

    def put(key, *parts):
        nonlocal r
        RD.cell(r, 1, F(*parts)); out[key] = r; r += 1
    put("hero", T("加權目標價 "), f'TEXT({tp},"0")', T(" 美元，比現價 "), 'TEXT(px,"0.00")', f'IF({up}<0," 低 "," 高 ")', f'TEXT(ABS({up}),"0%")', T("，評等"), rec,
        T("：雲端營收長得快，但 AI 算力的利潤大多被設備折舊與機房租金吃掉；長約價格達 "), f'TEXT({C["REV_NEU"]},"0.00")', T(" 時評等翻為中立。"))
    put("rec", T("投資建議「"), rec, T("」成立的前提：長約價格 "), 'TEXT(p_long,"0.00")', T("（百萬美元／MW-IT·年）、計費比例 "), 'TEXT(u_bill,"0%")', T("、AI 設備折舊 "), 'TEXT(life_ai,"0")',
        T(" 年（Tokenomics）的基準假設；現價 "), 'TEXT(px,"0.00")', T(" 美元為 "), 'TEXT(px_date,"yyyy-mm-dd")', T(" 收盤。"))
    put("o1", T("目標價 "), f'TEXT({tp},"0")', T("＝現金流折現 "), f'TEXT({dcf},"0")', T("、分部加總 "), f'TEXT({cmpv},"0")', T(" 各半，比現價 "), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0%")',
        T("；市場共識 "), f'TEXT({cons},"0")', T("。三情境依序 "), f'TEXT({C["TPS_A"]},"0")', T("／"), f'TEXT({C["TPS_B"]},"0")', T("／"), f'TEXT({C["TPS_C"]},"0")', T("。"))
    put("o2", T("AI 算力賺的錢被折舊與租金吃掉：情境 B 的 FY2029 OCI 營收 "), f'TEXT(PL!F{C["PR"]["B"]["oci"]},"0")', T(" 十億美元，AI 營業利益率 "), f'TEXT(Thesis!G{TB["om"]},"0.0%")', T("、AI 增量 ROIC "),
        f'TEXT(Thesis!G{TB["roic"]},"0.0%")', T("，低於 WACC "), 'TEXT(wacc,"0.0%")', T("；毛利率由 FY2026 "), f'TEXT(IS!E{PLA["idx"]["gm"]},"0%")', T(" 降到 FY2029 "), f'TEXT(IS!H{PLA["idx"]["gm"]},"0%")', T("。"))
    put("o3", T("錢不夠用：FY2027–FY2031 自由現金流扣股利合計，情境 A／B／C 為 "), f'TEXT(CF!C{C["GATE"]},"0")', T("／"), f'TEXT(CF!D{C["GATE"]},"0")', T("／"), f'TEXT(CF!E{C["GATE"]},"0")',
        T(" 十億美元；情境 B 的調整後槓桿峰值 "), f'TEXT(MAX(Fund!D{C["FB_LEV"]}:H{C["FB_LEV"]}),"0.0")', T(" 倍，"), f'IF(MAX(Fund!D{C["FB_LEV"]}:H{C["FB_LEV"]})>lev_max,"超過","低於")', T("投資級上限 "), 'TEXT(lev_max,"0.0")', T(" 倍。"))
    put("o4", T("反向 DCF：現價隱含的長約價格約 "), f'TEXT({C["REV_EQ"]},"0.0")', T("（基準 "), 'TEXT(p_long,"0.0")', T("，敏感度上限 "), f'TEXT({C["REV_HI"]},"0.0")', T("）；分部加總的 OCI AI 以 AI 模組現金流折現計，只占情境 B 分部企業價值 "),
        f'TEXT({C["OCI_SH"]},"0%")', T("。"))
    put("o5", T("實績對帳（本版新增）：2026 年 1–8 月已公布資本支出 "), f'TEXT(Coord!D{CRD["rd1"]},"0")', T(" 十億美元，與 Wattonomics 隱含值相差 "), f'TEXT(Coord!F{CRD["rd1"]},"+0%;-0%")',
        T("（最近一季年化 "), f'TEXT(Coord!F{CRD["rd2"]},"+0%;-0%")', T("），但層級值由公司實績與指引換算（同源），結論："), f'Coord!G{CRD["rj"]}', T("。"))
    PB = C["PR"]["B"]
    put("drv", T("營收推導（情境 B）：Cloud Infrastructure（OCI）＝OCI AI 算力（已連網 MW × 每 MW 年收入，Thesis 頁）＋OCI 其他（非 AI，近四季年增向長期值收斂）；Cloud Applications（SaaS）、Software license、Software support、Hardware、Services 各自由近四季年增線性收斂到長期年增；"
                 "加總為官方三分部（Cloud and software、Hardware、Services）與合併營收。FY2029 合併營收 "), f'TEXT(PL!F{PB["rev"]},"0")', T(" 十億美元，其中 OCI "), f'TEXT(PL!F{PB["oci"]},"0")', T("（AI 算力 "), f'TEXT(PL!F{PB["ai"]},"0")', T("）、SaaS "),
        f'TEXT(PL!F{PB["saas"]},"0")', T("、軟體授權與支援 "), f'TEXT(PL!F{PB["lic"]}+PL!F{PB["sup"]},"0")', T("、硬體與服務 "), f'TEXT(PL!F{PB["hw"]}+PL!F{PB["svc"]},"0")', T("。"))
    RD.cell(r, 1, f"v{VER}（{BUILT}）：只改展示層與新增頁——個股 HTML 版面規格 v1.3：章節改為命題→估值→敏感度→反向 DCF→損益→驅動因子→收支模組→對帳→資料狀態；新增 IS 頁（標準損益表：FY2024A–FY2026A 10-K 實績＋FY2027E–FY2032E 情境 B，GAAP）、"
                  "Sens 頁末加權目標價相對基準（龍捲風圖用，白話標籤）、README 頁末 v1.3 白話列；新增 Coord 頁「三層協調」（F1 讀取層級；F1b 追公式；F2 C0（實績）：10-Q 資本支出（扣 2025-12）、OCI 營收與在役 MW vs 層級隱含值）。"
                  "既有數值格不變（新增列與新增頁以外逐格比對 0 差異），目標價與評等不變")
    out["ver"] = r; r += 1
    return out


def apply(wb, C):
    PLA = income_statement(wb, C)
    tor = sens_tor(wb, C)
    CRD = coord(wb, C)
    RMA = readme13(wb, C, PLA, CRD)
    return dict(IS=PLA, tor=tor, CRD=CRD, RM=RMA)
