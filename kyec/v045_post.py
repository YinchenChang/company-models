# 京元電 v0.4.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10 版面、同次產業 Advantest v0.7.5／泰瑞達 v0.4.5（測試）、日月光 v0.2.7（台股封測 C0））。
# 共通步驟取自 kyec/c13.py（與電源散熱四家同一份，未改）；本模型損益沒有業外、稅與非控制權益列（淨利＝營業利益 × 淨利 ÷ 營業利益），
# 故 Hist 與標準損益表改由本檔產生（hist_kyec、pl_kyec），其餘（RevDCF、Val 0.7、Sens 末表、Coord 讀取層級、README 白話列）呼叫 c13。
# build_kyec.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet, dn, BOLD, BLUE, HDR

VER, BUILT = "v0.4.5", "2026-11-15"   # 各 v1.3 區塊實際起始版（「v0.4.5 起」標籤用）
NEWVER = "v0.4.7"   # 現行版號（README A1 標題段；apply 末段改）
TN1008 = "https://finance.technews.tw/2026/10/08/kyec-2449-202609-financial-report/"
MDJ1105 = "https://www.moneydj.com/kmdj/news/newsviewer.aspx?a=2bcf599f-fe54-431a-8b6a-467b995f9ff7"
HSM = "https://histock.tw/stock/financial.aspx?no=2449"
SA_CF = "https://stockanalysis.com/quote/tpe/2449/financials/cash-flow-statement/"
CNA0306 = "https://www.cna.com.tw/news/afe/202603060261.aspx"

V46 = ("v0.4.6（2026-11-15）：3Q26 損益出處改公開資訊觀測站 t164sb04 單季（數值同 v0.3 已滾入值，只差捨入）；DCF 資產負債表日 2026-06-30 → 09-30（t164sb03：現金、長期投資、長期借款、租賃、非控制權益、股數），"
       "2026-08 已發放現金股利不再扣，stub 期改 4Q26 單季（4Q 資本支出＝下半年 ÷ 2，Assumed）；4Q26 假設不改。")

V47 = ("v0.4.7（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）：現價 228.5（11-05）→ 213.5（11-13）；同業前瞻 P/E 中位數 14.71 → 16.42；"
       "加權目標價 103.72 → 111.82（上檔空間 −54.6% → −47.6%，評等強力放空 → 放空）；DCF 目標價 69.50 → 69.68（只因現價改變 D/E、β 與 WACC 權重）、Comparable（相對倍數調整後）137.94 → 153.96；"
       "另一側讀數（Coord）99.40 → 106.73（強力放空）。共識 EPS、rf、預測用匯率與模型假設不變；日月光 ADS 換算的 USD/TWD 33.034 → 32.939（TWD=X 2026-11-13 收盤；K 批）；同業全數有 11-13 收盤，無未換價同業。")

C = dict(
    name="京元電", ver=VER, built=BUILT,
    fy26_short="FY2026E 含 1H26、3Q26 實績",
    fy26_long="1H26 實績＋3Q26 實績（公開資訊觀測站 t164sb04 單季）＋4Q26 假設",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 股數（本模型不計稀釋）",
    gw_label=None,   # 本模型不讀部署 GW（以加速器顆數推導）
    same_names="AI 加速器顆數（八平台合計，情境 B）",
    a1_tail=("v0.4.5（2026-11-15）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.4 更正；既有數值不變；" + V46),
    ver_line=(V47 + V46 + "v0.4.5（2026-11-15）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、Advantest、泰瑞達）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～09 與 3Q26 財報；結論：無可比實績，Siliconomics 沒有測試代工市場與京元電份額）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


# ---------- Hist：損益表歷史（只放可追到出處者；未取得者留空並註明） ----------
def hist_kyec(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = (f"Hist — 損益表歷史（{VER} 起；PL 頁「標準損益表」的歷史三年用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
               "FY2025＝Q 頁（公司季簡報與中央社 2026-03-06 全年數字）；FY2023、FY2024 只有營收（資料商轉載，繼續營業單位口徑：2025 年出售蘇州京隆後重編），"
               "毛利以下未取得（本版建置環境無法連網讀取年報；待補），標準損益表該格留空，不以推估補位。只被 PL 頁標準損益表引用，不回饋任何計算")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    rows = [
        ("rev", "營收", 23.992, "=rev24", "=a25_rev", "Verified",
         f"FY2023 23,992、FY2024 26,856（NT$M；stockanalysis 現金流量表頁的營收列，In 頁 fc_da 說明；{SA_CF}）；FY2024 本模型 In 頁 rev24＝349.34 億 ÷ 1.3008（雅虎新聞 2026-01-27 年增 30.08%）＝26.856，兩者一致；FY2025＝Q 頁四季合計"),
        ("gp", "毛利", None, None, "=a25_gp", "Verified", f"FY2025＝Q 頁四季（營收 × 毛利率，公司簡報）；中央社 2026-03-06 全年毛利率 35.84%（{CNA0306}）"),
        ("opx", "營業費用（推銷管理＋研發）", None, None, "=a25_opx", "Verified", "FY2025＝Q 頁四季（公司簡報）"),
        ("oi", "營業利益", None, None, "=a25_oi", "Verified", "FY2025＝Q 頁（毛利 − 營業費用；中央社全年 90.02 億，差額見 Q 頁核對列）"),
        ("ni", "稅後淨利（歸屬母公司）", None, None, "=a25_ni", "Verified", "FY2025 110.15 億（中央社 2026-03-06；含 1Q25 出售蘇州京隆的處分利益，金額未拆）"),
        ("sh", "EPS 換算股數（十億股）", None, None, "=a25_sh", "Derived", "FY2025＝淨利 ÷ 公司公告 EPS 9.01（Q 頁）"),
        ("eps", "EPS（公司公告，基本，NT$）", None, None, "=Q!B16", "Verified", "FY2025 9.01（Q 頁「公司揭露 EPS（基本）」）"),
    ]
    R = {}
    r = 4
    for k, labl, v23, v24, v25, tag, src in rows:
        R[k] = r
        H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl)
        for col, yy, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25)):
            if v is None:
                H[f"{col}{r}"] = "未取得"
                continue
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str):
                c.font = BLUE
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        H.cell(r, 6, tag); H.cell(r, 7, src)
        r += 1
    H.cell(r, 1, "核對"); H.cell(r, 2, "FY2025：毛利 − 營業費用 − 營業利益（應為 0）")
    H[f"E{r}"] = f"=E{R['gp']}-E{R['opx']}-E{R['oi']}"; H[f"E{r}"].number_format = "0.000"
    R["chk"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 40
    for c in "CDE":
        H.column_dimensions[c].width = 14
    H.column_dimensions["G"].width = 110
    return R


# ---------- PL：標準損益表（本模型沒有業外、稅、非控制權益列：預測期以法定稅率還原，見說明列） ----------
YEARS = c13.YEARS
COLS = c13.COLS


def pl_kyec(wb, CR):
    P = wb["PL"]
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    B = dict(rev=g("營收"), gp=g("毛利 ＝ 營收 − 現金營業成本 − 折舊攤提"), opx=g("營業費用"), oi=g("營業利益"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"))
    x0 = next(r for r in range(rb, P.max_row + 1) if str(P.cell(r, 1).value or "").startswith("收斂法") and str(P.cell(r, 1).value).endswith("情境 B"))
    gx = lambda label: lab(P, label, x0, exact=True, end=x0 + 8)
    X = dict(rev=gx("營收"), ni=gx("歸屬母公司淨利"), sh=gx("股數（沿用 sh_g）"), eps=gx("EPS"))
    D = wb["DCF"]
    dbr = lab(D, "情境 B")
    d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2031E（情境 B 基準；{C['fy26_short']}，FY2030E–FY2031E 為延伸期）。"
                 "NT$B；EPS 新台幣。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（台灣採用的 IFRS 合併財報，歸屬母公司淨利）。京元電不另發布 Non-GAAP，市場共識 EPS 亦為此口徑，故本表不做 Non-GAAP 調整。"
                 "本模型以「歸屬母公司淨利 ÷ 營業利益」合併處理業外、所得稅與非控制權益，沒有單獨的稅率列；本表預測期的稅前利益＝稅後淨利 ÷（1 − 法定稅率 20%，In 頁 WACC 用稅率），"
                 "所得稅＝稅前 × 20%，其他收入（業外淨額）＝稅前 − 營業利益，非控制權益列 0（評價用非控制權益僅 0.008 十億元）；這是展示用的還原，不是模型輸入。"
                 "歷史年度：FY2023、FY2024 只有營收（資料商轉載），FY2025 業外、稅前與所得稅未取得，皆留空")
    r_note = r; r += 1
    P.cell(r, 1, f"FY2026E＝{C['fy26_long']}；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益取 DCF 頁由淨利還原）："
                 "毛利率沿用 FY2029E 據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    fl = lambda y: f"Coord!${'BCDE'[y - 2026]}${CR['flag']}"
    for col, fy, cy in (("G", "FY2027E", 2027), ("H", "FY2028E", 2028), ("I", "FY2029E", 2029), ("J", "FY2030E", 2029), ("K", "FY2031E", 2029)):
        P[f"{col}{r}"] = f'=IF({fl(cy)}=1,"{fy}（對帳中）","{fy}")'
    r_hdr = r; r += 1
    L = {}
    M = "#,##0.0"; PCT = "0.0%"

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None:
                continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note:
            P.cell(r, 12, note)
        L[key] = r; r += 1
    eng = lambda key: [f"=PL!{c}{B[key]}" for c in "DEFG"]
    ext = lambda key: [f"=PL!H{X[key]}", f"=PL!I{X[key]}"]
    h3 = lambda k: [f"=h_is_{k}_23", f"=h_is_{k}_24", f"=h_is_{k}_25"]
    h1 = lambda k: [None, None, f"=h_is_{k}_25"]
    FC = "FGHIJK"   # FY2026E–FY2031E
    row("rev", "營收", "NT$B", h3("rev") + eng("rev") + ext("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [None] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 未取得")
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", [None, None] + [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "EFGHI"] + [f"=J{L['rev']}*(1-$I{r_c + 2})", f"=K{L['rev']}*(1-$I{r_c + 2})"], M,
        "FY2025A–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", h1("gp") + eng("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [None, None] + [f"={c}{L['gp']}/{c}{L['rev']}" for c in "EFGHIJK"], PCT, "FY2030E–FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "NT$B", h1("opx") + eng("opx") + [None] * 2, M,
        "FY2030E–FY2031E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）")
    row("oi", "營業利益", "NT$B", h1("oi") + eng("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E–FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利 ÷ 淨利對營業利益比率還原；不改）")
    row("opm", "營業利益率", "%", [None, None] + [f"={c}{L['oi']}/{c}{L['rev']}" for c in "EFGHIJK"], PCT)
    r_o = r
    row("oth", "其他收入（淨額，營業外收益）", "NT$B", [None, None, None] + [f"={c}{r_o + 1}-{c}{L['oi']}" for c in FC], M,
        "FY2026E 起＝稅前利益 − 營業利益（展示用還原）；FY2025 含處分利益，業外金額未取得")
    row("pti", "稅前利益", "NT$B", [None, None, None] + [f"={c}{r_o + 4}/(1-tax_w)" for c in FC], M, "FY2026E 起＝稅後淨利 ÷（1 − 法定稅率 20%）（展示用還原；非控制權益視為 0）")
    row("tax", "所得稅", "NT$B", [None, None, None] + [f"={c}{L['pti']}*tax_w" for c in FC], M, "FY2026E 起＝稅前 × 20%（In 頁 WACC 用稅率）")
    row("nci", "非控制權益損益", "NT$B", [None, None, None] + ["=0"] * 6, M, "模型不分非控制權益（評價用非控制權益 0.008 十億元）")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", h1("ni") + eng("ni") + ext("ni"), M)
    assert L["ni"] == r_o + 4
    row("nm", "淨利率", "%", [None, None] + [f"={c}{L['ni']}/{c}{L['rev']}" for c in "EFGHIJK"], PCT)
    row("sh", "股數", "十億股", h1("sh") + eng("sh") + ext("sh"), "0.000", "FY2026 起為配股後股數（FY2025 現金股利外另配股票股利）")
    row("eps", "EPS（基本，公司公告）", "NT$", h1("eps") + eng("eps") + ext("eps"), "0.00", C["eps_note"])
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [None] + [f"={c}{L['yoy']}" for c in COLS[1:]], PCT)
    row("r_gm", "毛利率", "%", [None, None] + [f"={c}{L['gm']}" for c in "EFGHIJK"], PCT)
    row("r_opm", "營業利益率", "%", [None, None] + [f"={c}{L['opm']}" for c in "EFGHIJK"], PCT)
    row("r_nm", "淨利率", "%", [None, None] + [f"={c}{L['nm']}" for c in "EFGHIJK"], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [None, None, None] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip("EFGHIJ", "FGHIJK")], PCT, "FY2026E 對 FY2025A：FY2025 含處分利益")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [None, None] + [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "EFGHIJK"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2025A–FY2029E）", "NT$B", [None, None] + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "EFGHI"] + [None, None], "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2030E–FY2031E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    for c in "JK":
        P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"I{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2030E–FY2031E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "NT$B", [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2030E–FY2031E 的營業費用是殘差，不是假設；FY2029E 欄供對照）", "%", [None] * 6 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "IJK"], "0.00%")
    row("c3", "核對：稅後淨利 − 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "NT$B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{B['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    row("c4", "核對：稅前 − 所得稅 − 非控制權益 − 稅後淨利（FY2026E 起）", "NT$B", [None] * 3 + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in FC], "0.000")
    row("c5", "核對：EPS − 稅後淨利 ÷ 股數", "NT$", [None, None, f"=E{L['eps']}-E{L['ni']}/E{L['sh']}"] + [f"={c}{L['eps']}-{c}{L['ni']}/{c}{L['sh']}" for c in FC], "0.000",
        "FY2025A：公司公告 EPS 對 Q 頁股數（Q 頁股數由淨利 ÷ 公告 EPS 推得，應為 0）")
    for col in COLS:
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    P.column_dimensions["L"].width = 70
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, B=B, X=X)
    return L


def layer_rows(ver, CR):
    tsmc = f"{ver}，經 tsmc/dist 成品 In 頁讀取"
    usum = lambda y: "+".join(f"u_{k}_{y}" for k in ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"))
    return [
        ("AI 加速器顆數 CY2026／27／28（八平台合計，原始顆數；u_<平台>_26–28；百萬顆）", f'=TEXT({usum(26)},"0.00")&"／"&TEXT({usum(27)},"0.00")&"／"&TEXT({usum(28)},"0.00")', tsmc,
         "加權顆數的分子（× 每顆測試時間 × 曝險）；顆數本身多為 Interested-party／Assumed，見上游", None),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取",
         "基準情境 B 的加速器顆數係數；A（原始顆數）、C 為個股自訂情境", None),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "只用於每顆美元換算的顯示，不影響營收與 EPS", "0.00"),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("u", "AI 加速器顆數（八平台合計，未加權）", "百萬顆", "加速器顆數合計（未加權）", list(N["SIF_UnitsPlat_Total_Base"][1:5]),
         "Derived：tsmc 成品顆數 × 情境 B 顆數係數（CY2029＝CY2028 ×（1＋平台成長）× 係數）", f"Derived：Siliconomics {SIL.VERSION} SIF_UnitsPlat_Total_Base（原始顆數 × SIF_UnitsCoef_Base）", "0.00"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sl = lambda k: f"Sens!L{lab(SE, k, 4)}"
    AI = "AI 加速器測試營收 ＝ 每加權顆營收 × 加權顆數 × 份額 × 單價"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；2026 年擴產的折舊能否被 AI 加速器測試量吸收決定盈餘，目標價對顆數情境與折舊率最敏感。")),
        premise="AI 加速器顆數（Siliconomics 基準，經台積電模型讀入）、每顆測試時間與京元電曝險、份額與單價路徑（假設）成立",
        o2=(T("AI 加速器測試是成長主力：情境 B 的 AI 測試營收由 FY2026 約 "), f'TEXT({pb(AI, "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb(AI, "F")}*10,"#,##0")',
            T(" 億元，占營收由 "), f'TEXT({pb("AI 加速器測試 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("AI 加速器測試 ÷ 營收", "F")},"0%")',
            T("；同期折舊占營收由 "), f'TEXT({pb("折舊攤提 ÷ 營收", "D")},"0%")', T(" 變為 "), f'TEXT({pb("折舊攤提 ÷ 營收", "F")},"0%")', T("。")),
        o3=(T("最大的變數是顆數情境與折舊：以情境 A 為基準一次動一個輸入，顆數情境 C 使 FY2028 EPS "), f'TEXT({sl("顆數情境 C")},"+0.0%;-0.0%")', T("，折舊率 14% 為 "),
            f'TEXT({sl("折舊率 14%")},"+0.0%;-0.0%")', T("，非 AI 成長取低值為 "), f'TEXT({sl("非 AI 成長 低")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。")),
        o5_tail="本模型的加速器顆數讀 Siliconomics（經台積電模型）；C0（實績）無可比實績（Siliconomics 沒有測試代工市場與京元電份額，公司也未揭露 AI 營收）。",
        drv=(T("分部營收推導（情境 B）：AI 加速器測試營收＝每加權顆營收 × 加權顆數（顆數 × 每顆測試時間 × 曝險）× 份額相對值 × 測試單價指數；非 AI 營收依年增率；相加為營收。"
               "FY2028 AI 測試 "), f'TEXT({pb(AI, "F")},"#,##0.0")', T("、非 AI "), f'TEXT({pb("非 AI 營收", "F")},"#,##0.0")', T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0.0")',
             T(" 十億元；資本支出 "), f'TEXT({pb("資本支出", "F")},"#,##0.0")', T("、折舊 "), f'TEXT({pb("折舊攤提", "F")},"#,##0.0")', T(" 十億元。")),
    )


def c0_actuals(wb, G, SIL):
    CW = wb["Coord"]; Q = wb["Q"]
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝京元電已公布實績（月營收與 3Q26 合併財報；公司未發布 AI 營收）；層級側＝Siliconomics 現行版同口徑的隱含值。層級值是日曆年（京元電會計年度＝日曆年），"
                         "主讀數＝2026 年已公布期間累計實績 vs 層級 CY2026 × 月數 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（3Q26）× 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致；層級沒有同口徑值 → 無可比實績。"
                         "主讀數期間用 2026-01～09：10 月營收（11 月上旬公告）本版建置環境無法連網讀取，未納入；判定不受影響（層級側無同口徑值）。")
    R["limit"] = S.text("限制：Siliconomics 只有 AI 加速器顆數（本模型讀入，同源）與測試機設備市場（Advantest、泰瑞達的設備銷售，CY2026 SoC 測試機市場；設備商營收，不是測試代工服務營收），"
                        "沒有 OSAT 測試代工（晶圓測試、最終測試、燒機）的市場值，也沒有京元電份額；公司未揭露 AI 加速器測試營收（In 頁 AI 占比 35% 為 Morgan Stanley 估計，Interested-party）。"
                        "結論：無可比實績；缺層級的測試代工市場與京元電份額、公司的 AI 測試營收揭露。顆數讀取一致見本頁第 2 節（同源，不判定）。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["jul"] = S.row("月營收：2026-07", 3.991, None, "NT$B", "Verified：月營收公告 39.91 億元（科技新報 2026-10-08 轉載；本模型 In 頁 q3_rev 說明）", blue=True, url=TN1008)
    R["aug"] = S.row("月營收：2026-08", 4.080, None, "NT$B", "Verified：月營收公告 40.80 億元（同上）", blue=True, url=TN1008)
    R["sep"] = S.row("月營收：2026-09", 4.071, None, "NT$B", "Verified：月營收公告 40.71 億元（科技新報 2026-10-08，標題 9 月營收）", blue=True, url=TN1008)
    R["h1"] = S.row("累計：2026-01～06（1H26）", "=Q!C4", None, "NT$B", "Verified：Q 頁 1H26 營收 213.34 億元（1Q26＋2Q26 公司簡報；HiStock 月營收合計相同）", url=HSM)
    R["cum9"] = S.row("累計：2026-01～09（主讀數的實績基礎）", 33.475818, None, "NT$B", "Verified：前三季合併營收 334.76 億元（MoneyDJ 轉載公開資訊觀測站重大訊息 2026-11-05 董事會通過前三季財報；科技新報 2026-10-08 前 9 月月營收累計同值）", blue=True, url=MDJ1105)
    R["chk1"] = S.row("檢查：累計 9 月 − 1H26 − 7～9 月月營收（應接近 0；月營收四捨五入到百萬元）", f"=B{R['cum9']}-B{R['h1']}-B{R['jul']}-B{R['aug']}-B{R['sep']}", None, "NT$B", "Derived：累計相減驗算")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎）", "=q3_rev", None, "NT$B", "Verified：In 頁 q3_rev＝公開資訊觀測站 t164sb04 3Q26 營業收入 12,142,259 千元（v0.4.6；＝前三季 334.76 − 上半年 213.34 億元）")
    R["chk2"] = S.row("檢查：3Q26 − 7～9 月月營收合計（應接近 0）", f"=B{R['q3']}-B{R['jul']}-B{R['aug']}-B{R['sep']}", None, "NT$B", "Derived")
    R["oct"] = S.row("月營收：2026-10", "未取得", None, "", "公司 11 月上旬公告；本版建置環境無法連網讀取，待下一版補入主讀數（改為 2026-01～10）", fmt="General")
    R["ai"] = S.row("公司揭露：AI 加速器測試營收或占比", "未揭露", None, "", "公司只揭露測試類別（晶圓測試、最終測試、燒機）與應用別占比；In 頁 ai_26 35% 為 Morgan Stanley 估計（Interested-party）", fmt="General")
    R["months"] = S.row("已公布月數（主讀數期間 2026-01～09）", 9, None, "月", "輸入：本版納入的月營收至 2026-09", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；同口徑值不存在者照實列出）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lsvc"] = S.row("層級：測試代工（OSAT 測試服務）市場與京元電份額", "無", None, "", f"Siliconomics {SIL.VERSION} 無此量；不以本模型 AI 占比或營收錨點補位", fmt="General")
    R["lsoc"] = S.row("層級（參考，口徑不同）：SoC 測試機設備市場", SIL.NAMES["SIF_TestSoCMkt_Base"][1] if isinstance(SIL.NAMES["SIF_TestSoCMkt_Base"], list) else SIL.NAMES["SIF_TestSoCMkt_Base"],
                      None, "US$B", f"Siliconomics {SIL.VERSION} SIF_TestSoCMkt_Base（CY2026；測試機設備商營收，與測試代工服務營收口徑不同，不可比）", fmt="#,##0.0", blue=True)
    R["lu"] = S.row("層級：AI 加速器顆數（八平台合計）", f"=B{G['u'][1]}", None, "百萬顆", "本頁第 2 節層級值（本模型讀入，同源；只列讀取層級）", fmt="#,##0.00")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～09 累計 vs 層級 CY2026 × 9 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    S.row("全公司營收：公司實績", f"=B{R['cum9']}", f"=B{R['q3']}*4", "NT$B", "Verified：月營收與 3Q26 財報（4a）", fmt="#,##0.0")
    S.row("全公司營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 沒有測試代工市場與京元電份額（4b）", fmt="General")
    R["j_rev"] = S.row("全公司營收：判定", "無可比實績", "無可比實績", "", "缺層級的測試代工市場與份額；不計入對帳中", fmt="General")
    S.row("AI 加速器測試營收：公司實績", "無（未揭露）", "無（未揭露）", "", "公司未揭露 AI 營收；不以券商占比或模型值補位", fmt="General")
    S.row("AI 加速器測試營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 只有加速器顆數，缺每顆測試營收與京元電份額", fmt="General")
    R["j_ai"] = S.row("AI 加速器測試營收：判定", "無可比實績", "無可比實績", "", "兩側皆缺；不計入對帳中", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Assumed：1H26、3Q26 實績＋4Q26 假設（季增 In 頁 q4g）", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 4Q26 月均營收（錨點 − 1～9 月累計）÷ 3", f"=(B{R['anc']}-B{R['cum9']})/3", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 9 月實績 − 1", f"=B{R['imp']}/B{R['sep']}-1", None, "%", "Derived：正值＝錨點假設 4Q 月營收高於 9 月水準", fmt="+0.0%;-0.0%")
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_rev"]}:C{R["j_ai"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績：京元電 2026-01～09 營收 "), f'TEXT(B{R["cum9"]}*10,"#,##0.0")', T(" 億元、3Q26 "), f'TEXT(B{R["q3"]}*10,"#,##0.0")',
                         T(" 億元為公司實績，但 Siliconomics 沒有測試代工市場與京元電份額，公司也未揭露 AI 測試營收，兩讀數無法計算。對帳中項目 "), nopen, T(" 項。")))
    R["sum2"] = S.text(F(T("參考（模型對實績）：FY2026 營收錨點 "), f'TEXT(B{R["anc"]}*10,"#,##0")', T(" 億元隱含 4Q26 月均 "), f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元，較 9 月實績 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("；累計相減驗算差 "), f'TEXT(B{R["chk1"]}*10,"0.00")', T(" 億元。")))
    R["sum3"] = S.text("限制與待查：10 月營收未納入（本版建置環境無法連網）；Siliconomics 若補測試代工市場與京元電份額（或公司揭露 AI 測試營收），即可改為可判定的 C0 列。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    HR = hist_kyec(wb)
    PLA = pl_kyec(wb, CR)
    RVA = c13.revdcf(wb, C)
    VG = c13.val_g(wb, C)
    ST = c13.sens_tor(wb, CR, C)
    G = c13.coord_c0(wb, SIL, CR, C)
    R4 = c0_actuals(wb, G, SIL)
    RMA = c13.readme13(wb, RVA, PLA, ST, CR, C, lines(wb))
    c13.version_a1(wb, C)
    _a1 = wb["README"]["A1"].value; _h, _s, _r = _a1.partition("；")   # v0.4.6：標題段版號改現行版（共用 c13 不改）
    assert VER + "（" in _h, _h
    wb["README"]["A1"].value = _h.replace(VER + "（", NEWVER + "（", 1) + _s + _r
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, G=G, R4=R4, RM=RMA)
    return (f"v045_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {ST['t0']} 列；"
            f"Coord C0 第 {G['c0']} 列起、C0（實績）第 {R4['title']}–{R4['end']} 列；README v1.3 白話列 {len(RMA)} 列")
