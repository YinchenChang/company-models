# 旺矽 v0.2.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10 版面、同次產業 Advantest v0.7.5／泰瑞達 v0.4.5（測試）、京元電 v0.4.5（台股測試，同一寫法））。
# 共通步驟取自 mpi/c13.py（與電源散熱四家同一份，未改）；本模型損益沒有業外、稅列（淨利＝營業利益 × 淨利 ÷ 營業利益）、股數為百萬股，
# 故 Hist、標準損益表與敏感度末表（加上本模型逐案整本重算的加權目標價）由本檔產生；其餘（RevDCF、Val 0.7、Coord 讀取層級、README 白話列）呼叫 c13。
# build_mpi.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet, dn, BOLD, BLUE, HDR

VER, BUILT = "v0.2.5", "2026-11-15"
TN9 = "https://finance.technews.tw/2026/10/08/mpi-6223-202609-financial-report/"
SA_IS = "https://stockanalysis.com/quote/tpex/6223/financials/"
FG4 = "https://blog.fugle.tw/post/earnings-call-6223-2026-03-03"
PK2 = "https://www.pocket.tw/school/articles/6433/"
FG1 = "https://blog.fugle.tw/post/earnings-call-6223-2026-05-15"

C = dict(
    name="旺矽", ver=VER, built=BUILT,
    fy26_short="FY2026E 含 1H26 實績與 3Q26 營收實績",
    fy26_long="1Q、2Q26 實績＋3Q26 營收實績（月營收）＋3Q26 損益依公司說法＋4Q26 依指引（3Q26 財報本版未取得）",
    eps_note="歷史為公司公告 EPS；預測為歸屬母公司淨利 ÷ 股數（可轉換公司債轉換稀釋未計）",
    gw_label=None,   # 本模型不讀部署 GW（以加速器顆數推導）
    same_names="AI 加速器顆數（八平台合計，情境 B）",
    a1_tail=(f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.2.2 更正；既有數值不變"),
    ver_line=(f"{VER}（{BUILT}）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、Advantest、泰瑞達）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～09；結論：無可比實績，Siliconomics 沒有探針卡市場與旺矽份額）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


# ---------- Hist：損益表歷史（只放可追到出處者；未取得者留空並註明） ----------
def hist_mpi(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = (f"Hist — 損益表歷史（{VER} 起；PL 頁「標準損益表」的歷史三年用）。金額：十億新台幣；EPS：新台幣；股數：百萬股；歸屬母公司。"
               f"FY2023、FY2024 營收、稅前、淨利取資料商轉載（stockanalysis 財務摘要，本模型 In 頁 DCF 用稅率說明；{SA_IS}）；FY2025＝Q 頁（公司法說摘要）。"
               "毛利、營業利益與 FY2023–24 EPS 未取得（本版建置環境無法連網讀取年報；待補），標準損益表該格留空，不以推估補位。只被 PL 頁標準損益表引用，不回饋任何計算")
    H["A1"].font = BOLD
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    rows = [
        ("rev", "營收", 8.147, 10.172, "=a25_rev", "Verified", "FY2023 8,147、FY2024 10,172（NT$M；stockanalysis）；FY2025＝Q 頁 13.377（公司，+31.5%；stockanalysis 13,371）"),
        ("gp", "毛利", None, None, "=a25_gp", "Interested-party", "FY2025＝Q 頁營收 × 毛利率 57.4%（高盛，財訊 2026-10-12；公司全年毛利未取得）"),
        ("pti", "稅前利益", "=8.147*0.1942", "=10.172*0.2747", "=13.371*0.2875", "Derived", "營收 × 稅前淨利率（stockanalysis：19.42%、27.47%、28.75%）"),
        ("ni", "稅後淨利（歸屬母公司）", 1.312, 2.301, "=a25_ni", "Verified", "FY2023 1,312、FY2024 2,301（NT$M；stockanalysis）；FY2025＝Q 頁 3.18（富果法說摘要；stockanalysis 3,177）"),
        ("tax", "所得稅（＝稅前 − 淨利；無非控制權益）", "=C{pti}-C{ni}", "=D{pti}-D{ni}", "=E{pti}-E{ni}", "Derived", "公司無非控制權益（stockanalysis 資產負債表）；所得稅直接數字未取得"),
        ("sh", "EPS 換算股數（百萬股）", None, None, "=a25_sh", "Derived", "FY2025＝淨利 ÷ 公司揭露 EPS（Q 頁）"),
        ("eps", "EPS（公司公告，NT$）", None, None, "=Q!B18", "Verified", "FY2025 33.49（Q 頁「公司揭露 EPS」，富果法說摘要）"),
    ]
    R = {}
    r = 4
    for i, row in enumerate(rows):
        R[row[0]] = r + i
    for k, labl, v23, v24, v25, tag, src in rows:
        r = R[k]
        H.cell(r, 1, f"h_is_{k}"); H.cell(r, 2, labl)
        for col, yy, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25)):
            if v is None:
                H[f"{col}{r}"] = "未取得"
                continue
            if isinstance(v, str):
                v = v.format(**R)
            c = H[f"{col}{r}"]; c.value = v; c.number_format = "0.00" if k == "eps" else "#,##0.000"
            if not isinstance(v, str):
                c.font = BLUE
            dn(wb, f"h_is_{k}_{yy}", f"Hist!${col}${r}")
        H.cell(r, 6, tag); H.cell(r, 7, src)
    r = max(R.values()) + 1
    H.cell(r, 1, "核對"); H.cell(r, 2, "FY2025：EPS × 股數 ÷ 1000 − 淨利（應為 0）")
    H[f"E{r}"] = f"=E{R['eps']}*E{R['sh']}/1000-E{R['ni']}"; H[f"E{r}"].number_format = "0.000"
    R["chk"] = r
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 40
    for c in "CDE":
        H.column_dimensions[c].width = 14
    H.column_dimensions["G"].width = 110
    return R


# ---------- PL：標準損益表 ----------
YEARS = c13.YEARS
COLS = c13.COLS


def pl_mpi(wb, CR):
    P = wb["PL"]
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 40)
    B = dict(rev=g("營收"), gp=g("毛利"), opx=g("營業費用"), oi=g("營業利益"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"))
    D = wb["DCF"]
    dbr = lab(D, "情境 B")
    d_ebit = lab(D, "營業利益 EBIT", dbr)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績）＋FY2026E–FY2031E（情境 B 基準；{C['fy26_short']}，FY2030E–FY2031E 為延伸期）。"
                 "NT$B；EPS 新台幣；股數百萬股。只引用本頁上方、Hist、Q、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, "口徑：GAAP（台灣採用的 IFRS 合併財報，歸屬母公司淨利）。旺矽不另發布 Non-GAAP，市場共識 EPS 亦為此口徑，故本表不做 Non-GAAP 調整。"
                 "本模型以「歸屬母公司淨利 ÷ 營業利益」合併處理業外與所得稅，沒有單獨的稅率列；本表預測期的稅前利益＝稅後淨利 ÷（1 − 有效稅率，In 頁 DCF 用稅率：近三年 1 − 淨利 ÷ 稅前平均），"
                 "所得稅＝稅前 × 有效稅率，其他收入（業外淨額）＝稅前 − 營業利益，非控制權益 0（公司無非控制權益）；這是展示用的還原，不是模型輸入。"
                 "歷史年度：FY2023、FY2024 只有營收、稅前與淨利（資料商轉載），FY2025 營業利益未取得，皆留空")
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
    ext = lambda key: [f"=PL!H{B[key]}", f"=PL!I{B[key]}"]
    h3 = lambda k: [f"=h_is_{k}_23", f"=h_is_{k}_24", f"=h_is_{k}_25"]
    h1 = lambda k: [None, None, f"=h_is_{k}_25"]
    FC = "FGHIJK"
    row("rev", "營收", "NT$B", h3("rev") + eng("rev") + ext("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", [None] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 未取得")
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", [None, None] + [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "EFGHI"] + [f"=J{L['rev']}*(1-$I{r_c + 2})", f"=K{L['rev']}*(1-$I{r_c + 2})"], M,
        "FY2025A–FY2029E＝營收 − 毛利；FY2030E–FY2031E＝營收 ×（1 − FY2029E 毛利率）")
    row("gp", "毛利", "NT$B", h1("gp") + eng("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M, "FY2025A 毛利率為券商數字（Interested-party），公司全年毛利未取得")
    row("gm", "毛利率", "%", [None, None] + [f"={c}{L['gp']}/{c}{L['rev']}" for c in "EFGHIJK"], PCT, "FY2030E–FY2031E＝FY2029E（沿用）")
    row("opx", "營業費用（研發＋銷管）", "NT$B", [None, None, None] + eng("opx") + [f"=J{L['rev']}*$I{r_c + 3}/$I{L['rev']}", f"=K{L['rev']}*$I{r_c + 3}/$I{L['rev']}"], M,
        "FY2025A 未取得；FY2030E–FY2031E＝營收 × FY2029E 營業費用率")
    row("oi", "營業利益", "NT$B", [None, None, None] + eng("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2025A 未取得；FY2030E–FY2031E 取 DCF 頁情境 B 的營業利益（由淨利 ÷ 淨利對營業利益比率還原；不改）")
    row("opm", "營業利益率", "%", [None, None, None] + [f"={c}{L['oi']}/{c}{L['rev']}" for c in FC], PCT)
    r_o = r
    row("oth", "其他收入（淨額，營業外收益）", "NT$B", [None, None, None] + [f"={c}{r_o + 1}-{c}{L['oi']}" for c in FC], M, "FY2026E 起＝稅前利益 − 營業利益（展示用還原）")
    row("pti", "稅前利益", "NT$B", h3("pti") + [f"={c}{r_o + 4}/(1-tax_w)" for c in FC], M, "歷史＝營收 × 稅前淨利率（資料商）；FY2026E 起＝稅後淨利 ÷（1 − 有效稅率）（展示用還原）")
    row("tax", "所得稅", "NT$B", h3("tax") + [f"={c}{L['pti']}*tax_w" for c in FC], M, "歷史＝稅前 − 淨利；FY2026E 起＝稅前 × 有效稅率（In 頁 DCF 用稅率）")
    row("nci", "非控制權益損益", "NT$B", ["=0"] * 9, M, "公司無非控制權益")
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", h3("ni") + eng("ni") + ext("ni"), M)
    assert L["ni"] == r_o + 4
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    row("sh", "股數", "百萬股", h1("sh") + eng("sh") + ext("sh"), "#,##0.00")
    row("eps", "EPS（基本，公司公告）", "NT$", h1("eps") + eng("eps") + ext("eps"), "0.00", C["eps_note"])
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [None] + [f"={c}{L['yoy']}" for c in COLS[1:]], PCT)
    row("r_gm", "毛利率", "%", [None, None] + [f"={c}{L['gm']}" for c in "EFGHIJK"], PCT)
    row("r_opm", "營業利益率", "%", [None, None, None] + [f"={c}{L['opm']}" for c in FC], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", [None] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023A 的基期 FY2022 未取得")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [None, None] + [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in "EFGHIJK"], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2026E–FY2029E）", "NT$B", [None, None, None] + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "FGHI"] + [None, None], "0.000")
    row("c2b", "差額（FY2030E–FY2031E）：毛利 − 營業費用（沿用 FY2029E 比率推算）− 延伸期營業利益（由淨利還原）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", "差額（FY2030E–FY2031E）占營收", "%", [None] * 7 + [f"={c}{L['c2b']}/{c}{L['rev']}" for c in "JK"], PCT)
    row("c3", "核對：稅後淨利 − 情境 B 推導區塊（FY2026E–FY2031E）", "NT$B", [None] * 3 + [f"={c}{L['ni']}-PL!{s}{B['ni']}" for c, s in zip(FC, "DEFGHI")], "0.000")
    row("c4", "核對：稅前 − 所得稅 − 非控制權益 − 稅後淨利", "NT$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c5", "核對：EPS − 稅後淨利 ÷ 股數 × 1000", "NT$", [None, None] + [f"={c}{L['eps']}-{c}{L['ni']}/{c}{L['sh']}*1000" for c in "EFGHIJK"], "0.000")
    for col in COLS:
        P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    P.column_dimensions["L"].width = 70
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, chk=r_chk, start=s0, B=B)
    return L


# ---------- Sens：加權目標價相對基準（本模型逐案整本重算的驅動案例＋目標價參數；H 節） ----------
DRV = {"mf": "AI 加速器部署量（FY2027 起顆數乘數低／高）", "ex_nv": "旺矽對 NVIDIA 加速器的曝險低／高", "ai_sh26": "AI／HPC 占探針卡（FY2026 校準）低／高",
       "g_ct": "每顆加速器的探針卡內容年增低／高", "g_nonai": "非 AI SoC 市場年增低／高", "d_npc": "其他探針卡超額成長低／高", "k_cpo": "CPO 測試設備營收低／高",
       "gm_adj": "毛利率三情境同時下調／上調", "phi": "營業費用對營收成長的彈性低／高", "ni_oi": "淨利對營業利益的比率低／高"}


def sens_mpi(wb, CR):
    SE = wb["Sens"]; V = wb["Val"]
    assert str(SE.cell(3, 10).value).startswith("加權目標價相對基準"), SE.cell(3, 10).value
    pairs = []
    r = 5
    while SE.cell(r, 2).value in DRV:
        k = SE.cell(r, 2).value
        assert SE.cell(r + 1, 2).value == k, (r, k)
        pairs.append((k, r, r + 1)); r += 2
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    src = []
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        src.append(r); r += 1
    r = SE.max_row + 3
    SE.cell(r, 1, f"加權目標價相對基準（HTML 敏感度主圖用；{VER} 起，版面規格 v1.3 H 節；驅動案例取本頁上方逐案整本重算的加權目標價（J 欄），參數案例取「目標價參數」表；白話標籤，不回饋計算）").font = BOLD
    t0 = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    first = r
    for k, lo, hi in pairs:
        SE.cell(r, 1, DRV[k])
        SE.cell(r, 2, f"=J{lo}").number_format = "+0.0%;-0.0%"; SE.cell(r, 3, f"=J{hi}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{lo}、A{hi}"); r += 1
    for sr in src:
        t = str(SE.cell(sr, 1).value)
        key = next((x for x in c13.TOR_LABEL if t.startswith(x)), None)
        SE.cell(r, 1, c13.TOR_LABEL[key] if key else f"=A{sr}")
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{sr}"].value
            if v is None or v == "—":
                continue
            SE.cell(r, col, f"={scol}{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{sr}"); r += 1
    tp = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    cu = f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    SE.cell(r, 1, "同業比較法不加歷史倍數調整（相對倍數係數＝1；Comparable 取市場主流同業法）")
    SE.cell(r, 2, f"=(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})/{tp}-1").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Val 頁：DCF 與未調整 Comparable 依 DCF 權重加權"); r += 1
    SE.cell(r, 1, "產業部署 GW 改用另一側讀數（Wattonomics 換算；Siliconomics 標對帳中，整本重建）")
    SE.cell(r, 2, f"=Coord!D{CR['tp']}").number_format = "+0.0%;-0.0%"
    SE.cell(r, 4, "Coord 頁兩讀數目標價（另一側相對現行）"); r += 1
    return dict(t0=t0, r1=first, r2=r - 1)


def layer_rows(ver, CR):
    return [
        ("AI 加速器顆數 CY2026／27／28／29（八平台合計，原始顆數；units_26–29；百萬顆）", '=TEXT(units_26,"0.00")&"／"&TEXT(units_27,"0.00")&"／"&TEXT(units_28,"0.00")&"／"&TEXT(units_29,"0.00")',
         f"{ver}（SIF_UnitsPlat_Total）", "晶圓測試量指數的分子（× 測試時間 × 旺矽曝險）", None),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}（um28_B、um29_B）",
         "基準情境 B 的加速器顆數係數；A（原始顆數）、C 為市場層同字母情境", None),
        ("測試時間加權係數 CY2026（NVIDIA ÷ 總顆數；ASIC 與 AMD × 旺矽曝險 ÷ 總顆數）", '=TEXT(nvtt_26,"0.000")&"／"&TEXT(astt_26,"0.000")', f"{ver}（平台顆數 × ts_tt_*）",
         "每顆加權係數（測試時間係數讀 Siliconomics；旺矽曝險為本模型假設）", None),
        ("SEMI 測試設備 CY2027 ÷ CY2026（SIF_SEMI_Test）", "=semi_27", ver, "設備分部 FY2027 成長", "0.000"),
        ("非 AI SoC 市場年增（ts_g_nonai）", "=g_nonai", ver, "其他探針卡的行業成長", "0.0%"),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("u", "AI 加速器顆數（八平台合計，含情境係數）", "百萬顆", "AI 加速器顆數（含情境係數）", list(N["SIF_UnitsPlat_Total_Base"][1:5]),
         "Derived：Siliconomics 原始顆數 × 情境 B 顆數係數", f"Derived：Siliconomics {SIL.VERSION} SIF_UnitsPlat_Total_Base（原始顆數 × SIF_UnitsCoef_Base）", "0.00"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 40)}"
    sj = lambda k: f"Sens!J{lab(SE, k, 4, exact=True)}"
    AI = "AI／HPC 探針卡（VPC、MEMS；命題子列）"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；AI／HPC 探針卡隨加速器晶圓測試量成長，目標價對 AI 部署量與每顆探針卡內容最敏感。")),
        premise="AI 加速器顆數與測試時間係數（Siliconomics 基準）、旺矽份額與價格相對值、每顆探針卡內容與 CPO 測試設備（假設）成立",
        o2=(T("AI／HPC 探針卡是成長主力：情境 B 由 FY2026 約 "), f'TEXT({pb(AI, "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb(AI, "F")}*10,"#,##0")',
            T(" 億元，占營收由 "), f'TEXT({pb("AI／HPC 探針卡 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("AI／HPC 探針卡 ÷ 營收", "F")},"0%")',
            T("；晶圓測試量指數 FY2028 為 FY2026 的 "), f'TEXT({pb("晶圓測試量指數（2026＝1）", "F")},"0.00")', T(" 倍。")),
        o3=(T("最大的變數是 AI 部署量與每顆內容：一次動一個輸入整本重算，AI 部署 ×0.8 使加權目標價 "), f'TEXT({sj("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("，每顆內容年增 0% 為 "),
            f'TEXT({sj("每顆內容年增 0%")},"+0.0%;-0.0%")', T("，毛利率下調兩個百分點為 "), f'TEXT({sj("毛利率 −2pt")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"#,##0")', T(" 元。")),
        o5_tail="本模型的加速器顆數與測試時間係數讀 Siliconomics；C0（實績）無可比實績（Siliconomics 沒有探針卡市場與旺矽份額，公司也未揭露 AI 營收）。",
        drv=(T("分部營收推導（情境 B）：AI／HPC 探針卡＝FY2026 × 晶圓測試量指數（加速器顆數 × 測試時間 × 曝險）× 份額與價格相對值 × 每顆內容指數；其他探針卡依非 AI SoC 成長；"
               "設備＝CPO 測試設備＋其他設備（SEMI 測試設備路徑）。FY2028 AI／HPC 探針卡 "), f'TEXT({pb(AI, "F")},"#,##0.0")', T("、其他探針卡 "),
             f'TEXT({pb("其他探針卡（CPC 等）", "F")},"#,##0.0")', T("、設備 "), f'TEXT({pb("設備", "F")},"#,##0.0")', T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0.0")', T(" 十億元。")),
    )


def c0_actuals(wb, G, SIL):
    CW = wb["Coord"]; Q = wb["Q"]
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝旺矽已公布實績（月營收、1Q26 與 2Q26 季報；公司未發布 AI 營收金額）；層級側＝Siliconomics 現行版同口徑的隱含值。層級值是日曆年（旺矽會計年度＝日曆年），"
                         "主讀數＝2026 年已公布期間累計實績 vs 層級 CY2026 × 月數 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（3Q26，月營收合計）× 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致；層級沒有同口徑值 → 無可比實績。"
                         "主讀數期間用 2026-01～09：10 月營收（11 月上旬公告）與 3Q26 財報（11 月中）本版建置環境無法連網讀取，未納入；判定不受影響（層級側無同口徑值）。")
    R["limit"] = S.text("限制：Siliconomics 只有 AI 加速器顆數與每顆測試時間係數（本模型讀入，同源）、測試機設備市場與 SEMI 測試設備（設備商營收，與探針卡耗材營收口徑不同），"
                        "沒有探針卡（垂直、MEMS、懸臂）市場值，也沒有旺矽份額；公司只揭露探針卡、設備、其他的營收占比（法說），未揭露 AI／HPC 探針卡營收（In 頁 AI 占探針卡六成為假設）。"
                        "結論：無可比實績；缺層級的探針卡市場與旺矽份額、公司的 AI 探針卡營收揭露。顆數讀取一致見本頁第 2 節（同源，不判定）。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["q1"] = S.row("1Q26 營收", "=Q!C4", None, "NT$B", "Verified：Q 頁 1Q26 3.933（公司法說，富果 2026-05）", url=FG1)
    R["q2"] = S.row("2Q26 營收", "=Q!D4", None, "NT$B", "Verified：Q 頁 2Q26 5.233（公司財報，口袋學堂 2026-09-08）", url=PK2)
    R["cum9"] = S.row("累計：2026-01～09（主讀數的實績基礎）", 15.068, None, "NT$B", "Verified：前 9 月月營收累計 150.68 億元，年增 58.0%（科技新報 2026-10-08；本模型 Q 頁 3Q26 欄）", blue=True, url=TN9)
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎；累計 9 月 − 1H26）", "=Q!E4", None, "NT$B", "Derived：前 9 月 150.68 − 1H26 91.66 億元（7–9 月月營收合計；3Q26 財報本版未取得）")
    R["chk1"] = S.row("檢查：累計 9 月 − 1Q26 − 2Q26 − 3Q26（應為 0）", f"=B{R['cum9']}-B{R['q1']}-B{R['q2']}-B{R['q3']}", None, "NT$B", "Derived：累計相減驗算（3Q26 由相減得出，本列驗證 Q 頁一致；單月 7–9 月營收本版未取得）")
    R["oct"] = S.row("月營收：2026-10", "未取得", None, "", "公司 11 月上旬公告；本版建置環境無法連網讀取，待下一版補入主讀數（改為 2026-01～10）", fmt="General")
    R["ai"] = S.row("公司揭露：AI／HPC 探針卡營收", "未揭露", None, "", "公司只揭露探針卡、設備、其他占比（1Q26 71.1%／26.9%／2.0%，2Q26 探針卡 74.7%）；AI 占探針卡 60% 為假設（In 頁 ai_sh26）", fmt="General")
    R["pcq2"] = S.row("公司揭露：2Q26 探針卡營收（占比 × 營收）", "=Q!D5", None, "NT$B", "Derived：2Q26 營收 × 探針卡占比 74.7%（公司法說）；口述占比 × 營收屬 Assumed，不作 C0 實績")
    R["months"] = S.row("已公布月數（主讀數期間 2026-01～09）", 9, None, "月", "輸入：本版納入的月營收至 2026-09", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；同口徑值不存在者照實列出）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lsvc"] = S.row("層級：探針卡市場與旺矽份額", "無", None, "", f"Siliconomics {SIL.VERSION} 無此量；不以本模型 AI 占比或營收錨點補位", fmt="General")
    soc = SIL.NAMES["SIF_TestSoCMkt_Base"][1]
    R["lsoc"] = S.row("層級（參考，口徑不同）：SoC 測試機設備市場", soc, None, "US$B", f"Siliconomics {SIL.VERSION} SIF_TestSoCMkt_Base（CY2026；測試機設備商營收，不是探針卡）", fmt="#,##0.0", blue=True)
    R["lu"] = S.row("層級：AI 加速器顆數（八平台合計）", f"=B{G['u'][1]}", None, "百萬顆", "本頁第 2 節層級值（本模型讀入，同源；只列讀取層級）", fmt="#,##0.00")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～09 累計 vs 層級 CY2026 × 9 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    S.row("全公司營收：公司實績", f"=B{R['cum9']}", f"=B{R['q3']}*4", "NT$B", "Verified：月營收（4a）", fmt="#,##0.0")
    S.row("全公司營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 沒有探針卡市場與旺矽份額（4b）", fmt="General")
    R["j_rev"] = S.row("全公司營收：判定", "無可比實績", "無可比實績", "", "缺層級的探針卡市場與份額；不計入對帳中", fmt="General")
    S.row("AI／HPC 探針卡營收：公司實績", "無（未揭露）", "無（未揭露）", "", "公司未揭露 AI 探針卡營收；不以占比假設或模型值補位", fmt="General")
    S.row("AI／HPC 探針卡營收：層級隱含", "無同口徑隱含值", "無同口徑隱含值", "", "Siliconomics 只有加速器顆數與測試時間係數，缺每顆探針卡內容與旺矽份額", fmt="General")
    R["j_ai"] = S.row("AI／HPC 探針卡營收：判定", "無可比實績", "無可比實績", "", "兩側皆缺；不計入對帳中", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!G{lab(Q, '營收')}", None, "NT$B", "Derived／Interested-party：1H26 實績＋3Q26 營收實績＋4Q26 依指引季增（In 頁 q4_g）", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 4Q26 月均營收（錨點 − 1～9 月累計）÷ 3", f"=(B{R['anc']}-B{R['cum9']})/3", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 3Q26 月均 − 1", f"=B{R['imp']}/(B{R['q3']}/3)-1", None, "%", "Derived：等於錨點的 4Q26 季增假設（公司指引 5–10%）", fmt="+0.0%;-0.0%")
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_rev"]}:C{R["j_ai"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績：旺矽 2026-01～09 營收 "), f'TEXT(B{R["cum9"]}*10,"#,##0.0")', T(" 億元、3Q26 "), f'TEXT(B{R["q3"]}*10,"#,##0.0")',
                         T(" 億元為公司實績，但 Siliconomics 沒有探針卡市場與旺矽份額，公司也未揭露 AI 探針卡營收，兩讀數無法計算。對帳中項目 "), nopen, T(" 項。")))
    R["sum2"] = S.text(F(T("參考（模型對實績）：FY2026 營收錨點 "), f'TEXT(B{R["anc"]}*10,"#,##0")', T(" 億元隱含 4Q26 月均 "), f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元，較 3Q26 月均 "),
                         f'TEXT(B{R["impg"]},"+0.0%;-0.0%")', T("；累計相減驗算差 "), f'TEXT(B{R["chk1"]}*10,"0.00")', T(" 億元。")))
    R["sum3"] = S.text("限制與待查：10 月營收與 3Q26 財報未納入（本版建置環境無法連網）；Siliconomics 若補探針卡市場與旺矽份額（或公司揭露 AI 探針卡營收），即可改為可判定的 C0 列。")
    R["end"] = S.r - 1
    return R


def version_a1(wb):
    """README A1：標題段（「個股模型 vX（」）版號更正為新版、日期改建置日，末尾加本版說明（標題含全形分號，c13.version_a1 的切段不適用）。"""
    import re
    RD = wb["README"]
    a1 = RD["A1"].value
    a1, n = re.subn(r"個股模型 v\d+(?:\.\d+)+（", f"個股模型 {VER}（", a1, count=1)
    assert n == 1, a1[:80]
    a1 = re.sub("\u300020\\d\\d-\\d\\d-\\d\\d", "\u3000" + BUILT, a1, count=1)
    RD["A1"].value = a1 + "；" + C["a1_tail"]


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    HR = hist_mpi(wb)
    PLA = pl_mpi(wb, CR)
    RVA = c13.revdcf(wb, C)
    VG = c13.val_g(wb, C)
    ST = sens_mpi(wb, CR)
    G = c13.coord_c0(wb, SIL, CR, C)
    R4 = c0_actuals(wb, G, SIL)
    RMA = c13.readme13(wb, RVA, PLA, ST, CR, C, lines(wb))
    version_a1(wb)
    wb._v13 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, G=G, R4=R4, RM=RMA)
    return (f"v025_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF 頁；Val 0.7 第 {VG} 列；Sens 加權目標價相對基準第 {ST['t0']} 列；"
            f"Coord C0 第 {G['c0']} 列起、C0（實績）第 {R4['title']}–{R4['end']} 列；README v1.3 白話列 {len(RMA)} 列")
