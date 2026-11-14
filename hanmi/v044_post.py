# 韓美半導體 v0.4.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10 版面、Advantest v0.7.5／泰瑞達 v0.4.5（設備）、京元電 v0.4.5、旺矽 v0.2.5（同一寫法））。
# 共通步驟取自 hanmi/c13.py（與電源散熱四家同一份，未改）；本模型以 ₩億、韓元建模，損益沒有業外與稅列（淨利＝營業利益 × 淨利 ÷ 營業利益），
# 故 Hist 與標準損益表由本檔產生，c13 寫入的新台幣字樣在本檔改為韓元（只改本版新增的格）。build_hanmi.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。
# alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet, dn, BOLD, BLUE, HDR

VER, BUILT = "v0.4.4", "2026-11-15"
SA_IS = "https://stockanalysis.com/quote/krx/042700/financials/"

C = dict(
    name="韓美半導體", ver=VER, built=BUILT,
    fy26_short="FY2026E 含 1–3Q26 實績（第三季為初步實績）",
    fy26_long="1–3Q26 實績（第三季 2026-10-30 初步實績）＋第四季假設",
    eps_note="歷史為公司公告 EPS（資料商轉載，基本／稀釋口徑未解）；預測為淨利 ÷ 股數",
    gw_label=None,   # 本模型不讀部署 GW（以 HBM 位元推導）
    same_names="HBM 出貨位元、混合鍵合占比（情境 B）",
    a1_tail=(f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.3 更正；既有數值不變"),
    ver_line=(f"{VER}（{BUILT}）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、Advantest、泰瑞達）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（DART 初步實績 2026 年 1–3Q；結論：無可比實績，Siliconomics 沒有 TC 鍵合機市場與韓美份額）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)
YEARS = c13.YEARS
COLS = c13.COLS


def hist_hanmi(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = (f"Hist — 損益表歷史（{VER} 起；PL 頁「標準損益表」的歷史三年用）。金額：億韓元（₩億）；EPS：韓元；股數：百萬股。"
               f"FY2024、FY2025＝Q 頁（DART 轉載：valueline、butler.works、stockanalysis）；稅前與所得稅取 stockanalysis 損益表（本模型 In 頁 WACC 用稅率說明：FY2024 稅 45,746／稅前 198,360、FY2025 64,358／278,366 百萬韓元）；"
               f"FY2023 只有營收 1,590.09（stockanalysis，In 頁營運資金說明；{SA_IS}），其餘未取得（本版建置環境無法連網；待補）。只被 PL 頁標準損益表引用，不回饋任何計算")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    keys = ["rev", "gp", "oi", "pti", "tax", "ni", "nci", "sh", "eps"]
    R = {k: 4 + i for i, k in enumerate(keys)}
    rows = [
        ("rev", "營收", 1590.09, "=Q!B4", "=a25_rev", "Verified", "FY2023 159,009 百萬韓元（stockanalysis）；FY2024 5,589、FY2025 5,767（Q 頁；公司初步實績與 DART 轉載）"),
        ("gp", "毛利", None, "=Q!B5", "=a25_gp", "Verified", "Q 頁（valueline、stockanalysis，DART 轉載）"),
        ("oi", "營業利益", None, "=Q!B6", "=a25_oi", "Verified", "Q 頁（公司初步實績；valueline FY2025 2,513.9）"),
        ("pti", "稅前利益", None, 1983.60, 2783.66, "Verified", "stockanalysis 損益表：198,360、278,366（百萬韓元）"),
        ("tax", "所得稅", None, 457.46, 643.58, "Verified", "stockanalysis 損益表：45,746、64,358（百萬韓元）"),
        ("ni", "稅後淨利（歸屬母公司）", None, "=Q!B9", "=a25_ni", "Verified", "Q 頁（butler.works、valueline）"),
        ("nci", "非控制權益損益（＝稅前 − 所得稅 − 淨利）", None, "=D{pti}-D{tax}-D{ni}", "=E{pti}-E{tax}-E{ni}", "Derived", "資料商四捨五入差；接近 0"),
        ("sh", "EPS 換算股數（百萬股）", None, "=D{ni}/D{eps}*100", "=shares", "Derived", "淨利 ÷ 公司 EPS（FY2025 即 In 頁 shares）"),
        ("eps", "EPS（公司公告，韓元）", None, "=Q!B10", "=a25_eps", "Verified（來源不一致）", "Q 頁：FY2024 1,584、FY2025 2,231.7（stockanalysis 2,248、etfshopping 2,256；口徑未解）"),
    ]
    for k, labl, v23, v24, v25, tag, src in rows:
        r = R[k]
        H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl)
        for col, yy, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25)):
            if v is None:
                H[f"{col}{r}"] = "未取得"
                continue
            if isinstance(v, str):
                v = v.format(**R)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "#,##0.0"
            if not isinstance(v, str):
                c.font = BLUE
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        H.cell(r, 6, tag); H.cell(r, 7, src)
    r = max(R.values()) + 1
    H.cell(r, 1, "核對"); H.cell(r, 2, "營業利益 ＋ 業外（＝稅前 − 營業利益）− 稅前（恆等式；業外 FY2024 為負，含金融資產評價損失）")
    for col in "DE":
        H[f"{col}{r}"] = f"={col}{R['oi']}+({col}{R['pti']}-{col}{R['oi']})-{col}{R['pti']}"; H[f"{col}{r}"].number_format = "0.0"
    R["chk"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 44
    for c in "CDE":
        H.column_dimensions[c].width = 14
    H.column_dimensions["G"].width = 110
    return R


def pl_hanmi(wb, CR):
    P = wb["PL"]
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 35)
    B = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), ni=g("淨利"), eps=g("EPS"))
    x0 = lab(P, "情境 B（基準）", lab(P, "延伸到 FY2032"))
    gx = lambda label: lab(P, label, x0, exact=True, end=x0 + 6)
    X = dict(rev=gx("營收"), ni=gx("淨利（收斂法延伸）"), eps=gx("EPS（延伸）"))
    D = wb["DCF"]
    dbr = lab(D, "情境 B")
    d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2031E（情境 B 基準；{C['fy26_short']}，FY2030E–FY2031E 為延伸期）。"
                 "₩億；EPS 韓元；股數百萬股。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（韓國採用的 K-IFRS 合併財報）。韓美半導體不另發布 Non-GAAP，故本表不做 Non-GAAP 調整。"
                 "本模型以「淨利 ÷ 營業利益」合併處理業外與所得稅，沒有單獨的稅率列；本表預測期的稅前利益＝淨利 ÷（1 − 有效稅率，In 頁 WACC 用稅率：FY2024–25 平均），"
                 "所得稅＝稅前 × 有效稅率，其他收入（業外淨額）＝稅前 − 營業利益，非控制權益 0；這是展示用的還原，不是模型輸入。歷史年度：FY2023 只有營收（資料商轉載），其餘未取得，留空")
    r_note = r; r += 1
    P.cell(r, 1, f"FY2026E＝{C['fy26_long']}；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益取 DCF 頁由淨利還原）："
                 "毛利率與營業費用率沿用 FY2029E 據以算出營業成本、毛利與營業費用，與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    for j, h in enumerate(["項目", "單位"] + YEARS + ["說明"], 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    fl = lambda y: f"Coord!${'BCDE'[y - 2026]}${CR['flag']}"
    for col, fy, cy in (("G", "FY2027E", 2027), ("H", "FY2028E", 2028), ("I", "FY2029E", 2029), ("J", "FY2030E", 2029), ("K", "FY2031E", 2029)):
        P[f"{col}{r}"] = f'=IF({fl(cy)}=1,"{fy}（對帳中）","{fy}")'
    r_hdr = r; r += 1
    L = {}
    M = "#,##0"; PCT = "0.0%"

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
    h2 = lambda k: [None, f"=h_is_{k}_24", f"=h_is_{k}_25"]
    FC = "FGHIJK"
    row("rev", "營收", "₩億", ["=h_is_rev_23", "=h_is_rev_24", "=h_is_rev_25"] + eng("rev") + ext("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [None] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 未取得")
    r_c = r
    row("cogs", "營業成本（COGS）", "₩億", [None] + [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "DEFGHI"] + [f"=J{L['rev']}*(1-$I{r_c + 2})", f"=K{L['rev']}*(1-$I{r_c + 2})"], M,
        "FY2024A–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "₩億", h2("gp") + eng("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [None] + [f"={c}{L['gp']}/{c}{L['rev']}" for c in "DEFGHIJK"], PCT, "FY2030E–FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "₩億", [None, f"=D{L['gp']}-h_is_oi_24", f"=E{L['gp']}-h_is_oi_25"] + eng("opx") + [f"=J{L['rev']}*$I{r_c + 3}/$I{L['rev']}", f"=K{L['rev']}*$I{r_c + 3}/$I{L['rev']}"], M,
        "歷史＝毛利 − 營業利益；FY2030E–FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "₩億", h2("oi") + eng("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E–FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利還原；不改）")
    row("opm", "營業利益率", "%", [None] + [f"={c}{L['oi']}/{c}{L['rev']}" for c in "DEFGHIJK"], PCT)
    r_o = r
    row("oth", "其他收入（淨額，營業外收益）", "₩億", [None] + [f"={c}{r_o + 1}-{c}{L['oi']}" for c in "DEFGHIJK"], M, "歷史＝稅前 − 營業利益（FY2024 含金融資產評價損失）；FY2026E 起為展示用還原")
    row("pti", "稅前利益", "₩億", h2("pti") + [f"={c}{r_o + 4}/(1-tax_w)" for c in FC], M, "FY2026E 起＝淨利 ÷（1 − 有效稅率）（展示用還原；非控制權益視為 0）")
    row("tax", "所得稅", "₩億", h2("tax") + [f"={c}{L['pti']}*tax_w" for c in FC], M, "FY2026E 起＝稅前 × 有效稅率（In 頁 WACC 用稅率）")
    row("nci", "非控制權益損益", "₩億", h2("nci") + ["=0"] * 6, M, "歷史為資料商四捨五入差；預測 0")
    row("ni", "稅後淨利（歸屬母公司）", "₩億", h2("ni") + eng("ni") + ext("ni"), M)
    assert L["ni"] == r_o + 4
    row("nm", "淨利率", "%", [None] + [f"={c}{L['ni']}/{c}{L['rev']}" for c in "DEFGHIJK"], PCT)
    row("sh", "股數", "百萬股", h2("sh") + ["=shares"] * 6, "#,##0.00", "FY2026E 起＝In 頁 shares（不建模庫藏股註銷）")
    row("eps", "EPS（基本，公司公告）", "₩", h2("eps") + eng("eps") + ext("eps"), "#,##0", C["eps_note"])
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [None] + [f"={c}{L['yoy']}" for c in COLS[1:]], PCT)
    row("r_gm", "毛利率", "%", [None] + [f"={c}{L['gm']}" for c in COLS[1:]], PCT)
    row("r_opm", "營業利益率", "%", [None] + [f"={c}{L['opm']}" for c in COLS[1:]], PCT)
    row("r_nm", "淨利率", "%", [None] + [f"={c}{L['nm']}" for c in COLS[1:]], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [None, None] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip("DEFGHIJ", "EFGHIJK")], PCT)
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "₩億", [None] + [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "DEFGHIJK"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2024A–FY2029E）", "₩億", [None] + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "DEFGHI"] + [None, None], "0.000")
    row("c2b", "差額（FY2030E–FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利還原）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "₩億",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "差額（FY2030E–FY2031E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JK"], PCT)
    row("c3", "核對：淨利 − 情境 B 推導區塊（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "₩億",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{B['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    row("c4", "核對：稅前 − 所得稅 − 非控制權益 − 淨利", "₩億", [None] + [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in "DEFGHIJK"], "0.000")
    row("c5", "核對：EPS − 淨利 × 100 ÷ 股數（FY2026E 起）", "₩", [None] * 3 + [f"={c}{L['eps']}-{c}{L['ni']}*100/{c}{L['sh']}" for c in FC], "0.000")
    for col in COLS:
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    P.column_dimensions["L"].width = 70
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, B=B, X=X)
    return L


def layer_rows(ver, CR):
    return [
        ("HBM 核心晶粒容量 CY2026／27／28（Gb；hbm_den）", '=TEXT(den_26,"0")&"／"&TEXT(den_27,"0")&"／"&TEXT(den_28,"0")', f"{ver}（M22，與 Disco 模型同一值）", "核心晶粒數＝位元 ÷ 容量", None),
        ("HBM 平均堆疊層數 CY2026／27／28（hbm_stk）", '=TEXT(stk_26,"0.0")&"／"&TEXT(stk_27,"0.0")&"／"&TEXT(stk_28,"0.0")', f"{ver}（M22）", "層數加權（每顆晶粒鍵合時間）", None),
        ("混合鍵合占比 CY2028／CY2029（情境 B；hbm_hbp）", '=TEXT(hbp_28,"0%")&"／"&TEXT(hbp_29,"0%")', f"{ver}（M22；A 取高、C 取低）", "TC 鍵合產能需求 ×（1 − 混合鍵合）", None),
        ("HBM 出貨位元 CY2026／27／28／29（情境 B；EB）", '=TEXT(hsB_26,"0.00")&"／"&TEXT(hsB_27,"0.00")&"／"&TEXT(hsB_28,"0.00")&"／"&TEXT(hsB_29,"0.00")',
         "記憶體市場層（markets/memory，經 SK 海力士模型成品 Mkt 頁）；與 Siliconomics SIF_HBM_ShipEB_Base 同值", "鍵合產能需求的起點", None),
        ("匯率（₩/US$；fx_krw）", "=fx", ver, "只用於對帳的美元換算，不進損益", "#,##0"),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("ship", "HBM 出貨位元", "EB", "HBM 出貨位元", list(N["SIF_HBM_ShipEB_Base"][1:5]),
         "Derived：記憶體市場層 HBM 出貨（MIN(需求 × 情境 B 係數, 供給)），經 SK 海力士模型成品", f"Derived：Siliconomics {SIL.VERSION} SIF_HBM_ShipEB_Base", "0.00"),
        ("hb", "混合鍵合占比", "%", "混合鍵合占比", list(N["SIF_HBMHybrid_Base"][1:5]),
         "讀自 Siliconomics（In 頁 hbp_26–29；情境 B 取基準）", f"Assumed：Siliconomics {SIL.VERSION} SIF_HBMHybrid_Base", "0.0%"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 35)}"
    sh = lambda k: f"Sens!H{lab(SE, k, 4)}"
    HB = "HBM TC 鍵合機營收（三大記憶體廠）"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 韓元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 韓元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 韓元各半；HBM 鍵合機需求取決於 HBM 位元成長與混合鍵合導入，盈餘對混合鍵合導入速度與顆數情境最敏感。")),
        premise="HBM 出貨位元（記憶體市場層）、晶粒容量、堆疊層數與混合鍵合占比（Siliconomics）、韓美在三家記憶體廠的份額（假設）成立",
        o2=(T("HBM TC 鍵合機是營收主力：情境 B 由 FY2026 約 "), f'TEXT({pb(HB, "D")},"#,##0")', T(" 億韓元變為 FY2028 約 "), f'TEXT({pb(HB, "F")},"#,##0")',
            T(" 億韓元，占營收由 "), f'TEXT({pb("HBM TC 鍵合機占營收", "D")},"0%")', T(" 變為 "), f'TEXT({pb("HBM TC 鍵合機占營收", "F")},"0%")',
            T("；鍵合機比產能早一年購入，需求看下一年 HBM 位元增量。")),
        o3=(T("最大的變數是混合鍵合導入與顆數情境：以情境 A 為基準一次動一個輸入，混合鍵合不納入使 FY2028 EPS "), f'TEXT({sh("混合鍵合 不納入")},"+0.0%;-0.0%")',
            T("、改取基準路徑 "), f'TEXT({sh("混合鍵合 Disco 基準")},"+0.0%;-0.0%")', T("，顆數情境 C 為 "), f'TEXT({sh("顆數情境 C")},"+0.0%;-0.0%")',
            T("，淨利對營業利益比率取低值為 "), f'TEXT({sh("淨利÷營業利益 0.72")},"+0.0%;-0.0%")', T("；韓美在 SK 海力士的份額減十個百分點只有 "),
            f'TEXT({sh("SK 海力士份額 −10pt")},"+0.0%;-0.0%")', T("。基準情境 B 的 FY2028 EPS 為 "), f'TEXT({pb("EPS", "F")},"#,##0")', T(" 韓元。")),
        o5_tail="本模型的晶粒、堆疊層數與混合鍵合讀 Siliconomics，HBM 位元讀記憶體市場層；C0（實績）無可比實績（Siliconomics 沒有 TC 鍵合機市場與韓美份額，公司也未揭露 HBM 鍵合機營收）。",
        drv=(T("分部營收推導（情境 B）：HBM TC 鍵合機營收＝校準係數 × 韓美可得新機需求（HBM 位元 → 核心晶粒 × 層數加權 ×（1 − 混合鍵合）→ 產能增量＋汰換 → × 三家記憶體廠占比與韓美份額）；"
               "其他業務依年成長；相加為營收。FY2028 HBM TC 鍵合機 "), f'TEXT({pb(HB, "F")},"#,##0")', T("、其他業務 "),
             f'TEXT({pb("其他業務營收（中國 TC 鍵合機、MSVP、EMI、2.5D 等）", "F")},"#,##0")', T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0")', T(" 億韓元。")),
    )


def c0_actuals(wb, G, SIL):
    CW = wb["Coord"]; Q = wb["Q"]
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝韓美半導體已公布實績（DART 公平揭露的季度初步營收與營業利益；公司不揭露產品別營收）；層級側＝Siliconomics 現行版同口徑的隱含值。會計年度＝日曆年，"
                         "主讀數＝2026 年已公布期間累計實績（1–3Q）vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（3Q26）× 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率；"
                         "設備營收按交機認列、季度波動大，偏誤方向不定）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致；層級沒有同口徑值 → 無可比實績。")
    R["limit"] = S.text("限制：Siliconomics 只有 HBM 位元、晶粒容量、堆疊層數與混合鍵合占比（本模型讀入，同源），沒有 TC 鍵合機（或 HBM 後段鍵合設備）市場值，也沒有韓美份額；"
                        "公司不揭露 HBM TC 鍵合機營收（In 頁 TC 鍵合機占營收、中國占比為 CGSI 估計，Interested-party）；本模型的「三大廠 TC 鍵合機市場」由韓美營收校準（同源），不是層級值。"
                        "結論：無可比實績；缺層級的 TC 鍵合機市場與韓美份額、公司的 HBM 鍵合機營收揭露。HBM 位元與混合鍵合的讀取一致見本頁第 2 節（同源，不判定）。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["q1"] = S.row("1Q26 營收", "=q1_rev", None, "₩億", "Verified：Q 頁 26Q1 509（DART 初步實績；ddaily 2026-07-14 轉載）", fmt="#,##0")
    R["q2"] = S.row("2Q26 營收", "=q2_rev", None, "₩億", "Verified：Q 頁 26Q2 2,511（DART 初步實績；etfshopping 2,512）", fmt="#,##0")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎）", "=q3_rev", None, "₩億", "Verified：Q 頁 26Q3 1,810（2026-10-30 DART 公平揭露初步實績；SBS Biz、Bizwatch、dealsite 轉載）", fmt="#,##0")
    R["cum"] = S.row("累計：2026 年 1–3Q（主讀數的實績基礎）", f"=B{R['q1']}+B{R['q2']}+B{R['q3']}", None, "₩億", "Derived：三季初步實績相加（公司未另公布累計數；季報待複核）", fmt="#,##0")
    R["q3oi"] = S.row("3Q26 營業利益（參考）", "=q3_oi", None, "₩億", "Verified：908（營益率 50.2%；同上）", fmt="#,##0")
    R["chk1"] = S.row("檢查：累計 − 1Q − 2Q − 3Q（應為 0）", f"=B{R['cum']}-B{R['q1']}-B{R['q2']}-B{R['q3']}", None, "₩億", "Derived", fmt="0.0")
    R["hbm"] = S.row("公司揭露：HBM TC 鍵合機營收", "未揭露", None, "", "公司不揭露產品別營收；TC 鍵合機占 FY2025 營收 60%、中國占 17% 為 CGSI 估計（Interested-party）", fmt="General")
    R["months"] = S.row("已公布月數（主讀數期間 2026-01～09）", 9, None, "月", "輸入：1–3Q 初步實績", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；同口徑值不存在者照實列出）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lsvc"] = S.row("層級：TC 鍵合機（HBM 後段鍵合設備）市場與韓美份額", "無", None, "", f"Siliconomics {SIL.VERSION} 無此量；不以本模型校準的三大廠鍵合機市場補位", fmt="General")
    R["lhbm"] = S.row("層級（參考，口徑不同）：HBM 市場", SIL.NAMES["SIF_HBM_Mkt_Base"][1], None, "US$B", f"Siliconomics {SIL.VERSION} SIF_HBM_Mkt_Base（CY2026；記憶體營收，不是鍵合設備）", fmt="#,##0.0", blue=True)
    R["lship"] = S.row("層級：HBM 出貨位元", f"=B{G['ship'][1]}", None, "EB", "本頁第 2 節層級值（本模型讀入，同源；只列讀取層級）", fmt="#,##0.00")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026 年 1–3Q 累計 vs 層級 CY2026 × 9 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    S.row("全公司營收：公司實績", f"=B{R['cum']}", f"=B{R['q3']}*4", "₩億", "Verified：DART 初步實績（4a）", fmt="#,##0")
    S.row("全公司營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 沒有 TC 鍵合機市場與韓美份額（4b）", fmt="General")
    R["j_rev"] = S.row("全公司營收：判定", "無可比實績", "無可比實績", "", "缺層級的鍵合機市場與份額；不計入對帳中", fmt="General")
    S.row("HBM TC 鍵合機營收：公司實績", "無（未揭露）", "無（未揭露）", "", "公司不揭露產品別營收；不以券商占比或模型值補位", fmt="General")
    S.row("HBM TC 鍵合機營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 只有 HBM 位元與堆疊技術變數，缺鍵合機單價、裝機量與韓美份額", fmt="General")
    R["j_ai"] = S.row("HBM TC 鍵合機營收：判定", "無可比實績", "無可比實績", "", "兩側皆缺；不計入對帳中", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對已公布實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（In 頁）", "=rev26", None, "₩億", "Derived／Assumed：1–3Q 實績＋第四季（h2f × 2 × 第二季 − 第三季）", fmt="#,##0")
    R["imp"] = S.row("錨點隱含 4Q26 營收（錨點 − 1–3Q 累計）", f"=B{R['anc']}-B{R['cum']}", None, "₩億", "Derived", fmt="#,##0")
    R["impg"] = S.row("隱含月均 ÷ 3Q26 月均 − 1", f"=B{R['imp']}/B{R['q3']}-1", None, "%", "Derived：第四季對第三季的季增假設", fmt="+0.0%;-0.0%")
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_rev"]}:C{R["j_ai"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績：韓美 2026 年 1–3Q 營收 "), f'TEXT(B{R["cum"]},"#,##0")', T(" 億韓元、3Q26 "), f'TEXT(B{R["q3"]},"#,##0")',
                         T(" 億韓元為公司初步實績，但 Siliconomics 沒有 TC 鍵合機市場與韓美份額，公司也未揭露 HBM 鍵合機營收，兩讀數無法計算。對帳中項目 "), nopen, T(" 項。")))
    R["sum2"] = S.text(F(T("參考（模型對實績）：FY2026 營收錨點 "), f'TEXT(B{R["anc"]},"#,##0")', T(" 億韓元隱含第四季 "), f'TEXT(B{R["imp"]},"#,##0")', T(" 億韓元，較第三季 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("；累計相減驗算差 "), f'TEXT(B{R["chk1"]},"0.0")', T(" 億韓元。")))
    R["sum3"] = S.text("限制與待查：3Q26 為 DART 初步實績（季報與淨利待公布，本版建置環境無法連網讀取 DART 原文）；Siliconomics 若補 TC 鍵合機市場與韓美份額（或公司揭露產品別營收），即可改為可判定的 C0 列。")
    R["end"] = S.r - 1
    return R


CCY = (("十億元新台幣", "億韓元"), ("（NT$B；", "（₩億；"), ("（NT$B）", "（₩億）"), ("NT$B", "₩億"), ("（NT$）", "（₩）"), ("新台幣元", "韓元"), ("新台幣", "韓元"),
       ("十億元", "億韓元"), ("十億股", "億股"), ("NT$", "₩"))


def ccy_fix(ws, rows=None, cols=None):
    """把 c13 寫入的新台幣字樣改為韓元（只動指定範圍的字串格；公式內的字串常數一併改）。"""
    n = 0
    for row in ws.iter_rows(min_row=rows[0] if rows else 1, max_row=rows[1] if rows else ws.max_row):
        for c in row:
            if cols and c.column not in cols:
                continue
            v = c.value
            if not isinstance(v, str):
                continue
            w = v
            for a, b in CCY:
                w = w.replace(a, b)
            w = re.sub(r"(?<![韓])( )元", r"\1韓元", w)   # 「 元」→「 韓元」（已是「韓元」者不動）
            if w != v:
                c.value = w; n += 1
    return n


_lab0 = c13.lab


def _lab_val(ws, prefix, start=1, col=1, exact=False, end=None):
    """Val 頁有同名的標題列（B 欄空白，例：第 3 列「目標價（DCF 與 Comparable 加權）」）：c13 取 Val 頁列時跳過 B 欄空白的列，其餘同 c13.lab。"""
    if ws.title == "Val" and col == 1:
        for r in range(start, (end or ws.max_row) + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)) and ws.cell(r, 2).value is not None:
                return r
    return _lab0(ws, prefix, start, col, exact, end)


def apply(wb, SIL, CR):
    c13.lab = _lab_val   # c13 各函式在執行時取模組內的 lab
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    HR = hist_hanmi(wb)
    PLA = pl_hanmi(wb, CR)
    RVA = c13.revdcf(wb, C)
    ccy_fix(wb["RevDCF"])
    V = wb["Val"]; v0 = V.max_row
    VG = c13.val_g(wb, C)
    ccy_fix(V, rows=(VG, V.max_row))
    ST = c13.sens_tor(wb, CR, C)
    G = c13.coord_c0(wb, SIL, CR, C)
    R4 = c0_actuals(wb, G, SIL)
    RD = wb["README"]; rd0 = RD.max_row
    RMA = c13.readme13(wb, RVA, PLA, ST, CR, C, lines(wb))
    ccy_fix(RD, rows=(rd0 + 1, RD.max_row))
    c13.version_a1(wb, C)
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, G=G, R4=R4, RM=RMA)
    return (f"v044_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {ST['t0']} 列；"
            f"Coord C0 第 {G['c0']} 列起、C0（實績）第 {R4['title']}–{R4['end']} 列；README v1.3 白話列 {len(RMA)} 列")
