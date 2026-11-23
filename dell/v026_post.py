# Dell v0.2.6（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Broadcom、Astera Labs；同 lineage 緯穎 v0.2.6 wiwynn/v026_post.py）＋三層協調 C0（實績）公司實績對帳（規則 v1.4.2）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.2.5 逐格比對）：
#   Hist 頁（新）：non-GAAP 損益表歷史 FY2024–FY2026（公司 8-K 新聞稿調節表；FY2023 營收與淨利只供年增率）
#   PL 頁末：標準損益表（FY2024A–FY2026A 實績＋FY2027E–FY2032E 情境 B；公司 non-GAAP 口徑（與模型一致），歷史另列 GAAP 稀釋 EPS 對照）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC 與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末：0.7 四個目標價並列（HTML 用）
#   Sens 頁末：目標價敏感度：業務主因（整本重算快照 dell/sens.json）與「加權目標價相對基準」白話表（HTML 敏感度主圖用）
#   README 頁末：v1.3 展示層白話列（橫幅說明、重點觀察 ≤5 點、分部營收推導說明、本版要點）
#   Coord 頁末：C0 由下往上對帳（讀哪一層、哪一版；讀取層級同源列不判定）＋第 4 節 C0（實績）：1Q–2Q FY27 AI 伺服器營收（10-Q） vs 層級隱含值（同源校準，無可比實績）
# build_dell.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, SIL, here, coord_rows)。
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.6", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
PR426 = "https://www.sec.gov/Archives/edgar/data/1571996/000157199626000003/exhibit991earnings8kq4fy26.htm"
PR425 = "https://www.sec.gov/Archives/edgar/data/1571996/000157199625000014/exhibit991earnings8kq4fy25.htm"
PR424 = "https://www.sec.gov/Archives/edgar/data/1571996/000157199624000008/exhibit991earnings8kq4fy24.htm"
Q10_227 = "https://www.sec.gov/Archives/edgar/data/0001571996/000157199626000046/dell-20260731.htm"
SP = ("公司 8-K 新聞稿 non-GAAP 調節表（SEC EDGAR）：FY2026、FY2025 取 4Q FY26 新聞稿 " + PR426 + "；FY2024 取 4Q FY25 新聞稿（修正後數字，Exhibit A）" + PR425 + "；FY2023 營收與 non-GAAP 淨利取 4Q FY24 新聞稿 " + PR424 + "。"
      "百萬美元換算為十億美元。稅前利益＝non-GAAP 營業利益＋GAAP 利息及其他淨額（公司未列 non-GAAP 值）；所得稅＝稅前利益 − non-GAAP 淨利（隱含）；稀釋股數＝non-GAAP 淨利 ÷ non-GAAP 稀釋 EPS（推得）")
K = dict(
    cn="Dell", folder="dell", ENV="DELL",
    # (鍵, 項目, FY2024, FY2025, FY2026, FY2023（只供年增率）)；金額十億美元（non-GAAP）、EPS 美元、股數十億股
    HIS=[
        ("rev", "營收", 88.425, 95.567, 113.538, 102.301),
        ("cogs", "營業成本（non-GAAP；＝營收 − non-GAAP 毛利）", 66.781, 73.757, 90.379, None),
        ("gp", "non-GAAP 毛利", 21.644, 21.810, 23.159, None),
        ("opx", "non-GAAP 營業費用", 13.766, 13.281, 13.168, None),
        ("oi", "non-GAAP 營業利益", 7.878, 8.529, 9.991, None),
        ("pti", "稅前利益（non-GAAP 營業利益＋GAAP 利息及其他淨額；推得）", 6.554, 7.340, 9.105, None),
        ("tax", "所得稅（隱含：稅前利益 − non-GAAP 淨利）", 1.132, 1.475, 2.059, None),
        ("nci", "少數股權損益（non-GAAP 淨利已為歸屬 Dell；記 0）", 0.0, 0.0, 0.0, None),
        ("ni_rep", "non-GAAP 淨利（公司列示值）", 5.422, 5.865, 7.046, 5.727),
        ("eps_b", "non-GAAP 稀釋 EPS（美元）", 7.37, 8.14, 10.30, None),
        ("eps_d", "GAAP 稀釋 EPS（美元；對照）", 4.60, 6.38, 8.68, None),
        ("sh_b", "稀釋股數（十億股；＝non-GAAP 淨利 ÷ non-GAAP EPS，推得）", 0.73569, 0.72052, 0.68408, None),
        ("sh_d", "稀釋股數（十億股；同上）", 0.73569, 0.72052, 0.68408, None),
    ],
    SP=SP, SP22="；FY2023 營收 102,301、non-GAAP 淨利 5,727 百萬美元（4Q FY24 新聞稿）",
    PLB="情境 B（基準）：",
    R=dict(rev="營收", gp=None, opx=None, oi="分部營業利益（≈ non-GAAP 營業利益）", ni="non-GAAP 淨利", sh="稀釋股數", eps="non-GAAP EPS"),
    nci_mode="none", ext_mode="inblock", DCFX="GH", OIDIV="/gaap_k", LAST_A=2026, DEP=(2028, 2029, 2030, 2031, 2032), FY_LAG=1,   # FY_LAG：會計年度對應（FY 減 CY），寫入 Coord 頁具名格 co_fy_lag 供 check_layout_v13 判讀
    HIST_A1="會計年度：FY(t)＝t−1 年 2 月～t 年 1 月（≈ CY(t−1)）。來源為公司 8-K 新聞稿（SEC EDGAR）的 non-GAAP 調節表；FY2024 為公司修正後數字；稅前利益、所得稅與股數為推得值（見各列）。FY2026 營業利益為公司揭露 non-GAAP 值，模型 Q 頁用分部營業利益合計（差 0.047）",
    HIST_TAG="Verified（公司 8-K）／Derived（推得列）",
    PXL="現價（2026-11-20 收盤）", ANCH="1H FY27 實績＋公司全年指引＋2H 假設",
    EPSL="EPS（non-GAAP 稀釋；美元）", EPSN="歷史＝公司列示 non-GAAP 稀釋 EPS；預測＝non-GAAP 淨利 ÷ 模型稀釋股數",
    EPSDL="EPS（GAAP 稀釋；公司列示值，對照）", DILH="GAAP 稀釋 EPS 對照（歷史；預測不估 GAAP）", SHN="歷史與預測皆為稀釋股數",
    GAAP_NOTE=("口徑：Non-GAAP（公司 8-K 新聞稿口徑，排除攤銷、股權報酬等；與本模型 PL、Q 頁及市場共識一致）。會計年度 FY(t)≈CY(t−1)（2 月～隔年 1 月）。"
               "預測期模型只推分部營業利益（≈ non-GAAP 營業利益）與「淨利 ÷ 營業利益」一個係數，本表把它拆成：營業費用＝營收 × FY2026 實績營業費用率、毛利＝營業利益＋營業費用、"
               "其他收入＝營業利益 × FY2026 實績比例、所得稅＝稅前利益 − 淨利（由模型淨利倒推）。FY2031E–FY2032E 的營業利益取 DCF 頁 EBIT ÷ GAAP 換算係數（DCF 用 GAAP 口徑）。這幾列是展示用拆分，不回饋計算、不改淨利與目標價。"
               "標「對帳中」的年度：FY2028E–FY2030E 依賴 CY2027–CY2029（FY(t) 對 CY(t−1)）；FY2031E–FY2032E 為延伸期（收斂法由 FY2030 路徑推得，另一側讀數下數值隨之變動），亦標示。FY2027E 為錨點（公司全年指引），另一側讀數下數值不變，不標。"),
    RV_NOTE=("Dell FY2027（2026-02～2027-01）為錨點（公司全年指引），評價日後只含 2027-01 一個月，不隨反解變動；DCF 的 EBIT 為 GAAP 口徑（non-GAAP × 換算係數）"),
    COMP_NOTE="同業前瞻 P/E 中位數 × Dell 十年相對同業倍數比 × 約一年後前瞻 EPS（情境加權）",
    BIZ=[  # (白話標籤, 低案設定, 高案設定, 鍵)；設定＝DELL_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄；mf＝引擎 AI 部署係數）
        ("AI 部署（FY2028 起商用 GPU 顆數 ×0.8 ～ ×1.2）", "mf=0.8", "mf=1.2", "AI 部署"),
        ("Dell 份額（相對值 0.75／0.65／0.55 ～ 1.05／1.0／1.0）", "rel_28=lo;rel_29=lo;rel_30=lo", "rel_28=hi;rel_29=hi;rel_30=hi", "份額"),
        ("每顆系統價值（橋接表低值 ～ 高值）", "v_gb=lo;v_vr=lo", "v_gb=hi;v_vr=hi", "每顆價值"),
        ("AI 伺服器營益率變動（−3 ～ +1 個百分點）", "ai_dm=lo", "ai_dm=hi", "AI 營益率"),
        ("非 AI ISG 營益率（13.5% ～ 17.5%）", "m_naif=lo", "m_naif=hi", "非 AI 營益率"),
        ("CSG（PC）營益率（6% ～ 8%）", "m_csgf=lo", "m_csgf=hi", "CSG 營益率"),
        ("傳統伺服器成長（低 ～ 高）", "g_tr28=lo;g_tr29=lo;g_tr30=lo", "g_tr28=hi;g_tr29=hi;g_tr30=hi", "傳統伺服器"),
        ("儲存成長（低 ～ 高）", "g_st28=lo;g_st29=lo;g_st30=lo", "g_st28=hi;g_st29=hi;g_st30=hi", "儲存"),
        ("PC 成長（低 ～ 高）", "g_csg28=lo;g_csg29=lo;g_csg30=lo", "g_csg28=hi;g_csg29=hi;g_csg30=hi", "PC"),
        ("淨利 ÷ 營業利益（0.70 ～ 0.78）", "nio_f=lo", "nio_f=hi", "淨利係數"),
        ("股數年變動（−5% ～ −1%）", "sh_g=lo", "sh_g=hi", "回購"),
    ],
)


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def F(*parts):
    return "=" + "&".join(lit(p[1]) if isinstance(p, tuple) else p for p in parts)


T = lambda s: ("t", s)


def lab(ws, prefix, start=1, col=1, exact=False):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((ws.title, prefix))


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- Hist：損益表歷史 ----------
HIS = K['HIS']; SP = K['SP']


def hist_is(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = (f"Hist — {K['cn']}損益表歷史（v{VER} 新增；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億美元；EPS：美元。"
               + K.get("HIST_A1", "FY(t)＝CY(t)。來源為第三方彙整（S&P Global Market Intelligence），未逐項對公司年報原文；與模型 Q 頁 FY2026 可能有口徑或四捨五入差異，見表末對照列"))
    H["A1"].font = BOLD
    r = 3
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "FY2023（只供年增率）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v23, v24, v25, v22 in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "24", v23), ("D", "25", v24), ("E", "26", v25), ("F", "23", v22)):
            if v is None: continue
            c = H[f"{col}{r}"]; c.value = v; c.font = BLUE
            c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"
            dn(wb, f"h_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, K.get("HIST_TAG", "Verified／第三方彙整")); H.cell(r, 8, SP + (K["SP22"] if k in ("rev", "ni_rep") else ""))
        r += 1
    r += 1
    H.cell(r, 1, "核對（應為 0）").font = BOLD; r += 1
    chk = {}

    def crow(key, text, f, fmt="0.000"):
        nonlocal r
        H.cell(r, 1, text)
        for col in "CDE":
            H[f"{col}{r}"] = f.format(c=col, **{k: f"{col}{v}" for k, v in rows.items()}); H[f"{col}{r}"].number_format = fmt
        chk[key] = r; r += 1
    crow("c1", "營收 − 營業成本 − 毛利", "={rev}-{cogs}-{gp}")
    crow("c2", "毛利 − 營業費用 − 營業利益", "={gp}-{opx}-{oi}")
    crow("c3", "稅前利益 − 所得稅 − 少數股權損益 − 歸屬母公司淨利（列示值；四捨五入差 ≤ 0.001）", "={pti}-{tax}-{nci}-{ni_rep}")
    crow("c4", "歸屬母公司淨利 ÷ 流通股數 − 基本 EPS（資料來源股數為百萬股四捨五入，差 ≤ 0.05）", "={ni_rep}/{sh_b}-{eps_b}", "0.000")
    r += 1
    H.cell(r, 1, "對照：歷史（本頁，S&P Global）− 模型 FY2026A（Q 頁）。非零差異＝兩個來源口徑或四捨五入不同；模型既有 FY2026 欄不改").font = BOLD; r += 1
    for key, text, h_k, a_nm in (("d_rev", "營收", "rev", "a26_rev"), ("d_gp", "毛利", "gp", "a26_gp"), ("d_oi", "營業利益", "oi", "a26_oi"), ("d_ni", "歸屬母公司淨利", "ni_rep", "a26_ni"), ("d_eps", "EPS（基本）", "eps_b", "a26_eps")):
        if a_nm not in wb.defined_names:
            continue
        H.cell(r, 1, f"差異：{text}（本頁 − Q 頁 FY2026）")
        H[f"E{r}"] = f"=E{rows[h_k]}-{a_nm}"; H[f"E{r}"].number_format = "0.000"; chk[key] = r; r += 1
    H.column_dimensions["A"].width = 52; H.column_dimensions["B"].width = 44
    for c in "CDEF": H.column_dimensions[c].width = 14
    H.column_dimensions["H"].width = 100
    return rows


# ---------- PL：標準損益表 ----------
YR = list(range(2024, 2033))
COLS = "CDEFGHIJK"


def dep_years(SIL):
    if K.get("DEP"):   # 會計年度與日曆年錯開者：依賴對帳中日曆年的會計年度由公司設定檔指定
        return set(K["DEP"])
    ys = [y for y in (2026, 2027, 2028, 2029) if SIL.RECO_FLAG.get(y) == 1]
    out = set(ys)
    if 2029 in ys:
        out |= {2030, 2031}   # 延伸期由 FY2030 路徑收斂而來
    return out


def pl_standard(wb, SIL):
    P = wb["PL"]; D = wb["DCF"]
    rb = lab(P, K["PLB"])
    g = lambda p: lab(P, p, rb, exact=True)
    KR = K["R"]
    R = dict(rev=g(KR["rev"]), gp=(g(KR["gp"]) if KR.get("gp") else None), oi=g(KR["oi"]), ni=g(KR["ni"]), sh=g(KR["sh"]), eps=g(KR["eps"]))
    R["opx"] = g(KR["opx"]) if KR.get("opx") else None
    if K["nci_mode"] == "model":
        R["nci"] = lab(P, KR["nci"], rb); R["cons"] = g(KR["cons"])
    if K["ext_mode"] == "inblock":   # 緯穎：FY2031–FY2032 收斂法與情境 B 區塊同欄（H、I、J 欄），不另有延伸表
        X = {k: R[k] for k in ("rev", "ni", "sh", "eps")}
    elif K["ext_mode"] == "named":   # 延伸表（標題列 K["EXT"]["title"] 之後的「情境 B」列起；FY2031、FY2032 在 H、I 欄）
        E_ = K["EXT"]; xt = lab(P, E_["title"]); xb = lab(P, "情境 B", xt, exact=True)
        X = {k: lab(P, E_[k], xb, exact=True) for k in ("rev", "ni", "sh", "eps")}
    elif K["ext_mode"] == "block":
        xb = lab(P, "延伸｜情境 B")
        X = {k: lab(P, p, xb, exact=True) for k, p in (("rev", "營收（延伸）"), ("ni", "淨利（延伸）"), ("sh", "股數（延伸）"), ("eps", "EPS（延伸）"))}
    else:   # 摘要區塊（情境 B）H、I 欄＝FY2031、FY2032；股數＝淨利 ÷ EPS
        sb = lab(P, "情境 B", exact=True); SR = K["SUM_R"]
        X = {"rev": lab(P, SR["rev"], sb, exact=True), "ni": lab(P, SR["ni"], sb, exact=True), "eps": lab(P, SR["eps"], sb, exact=True)}
    db = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", db)
    dep = dep_years(SIL)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2024–FY2026 實績，Hist 頁）＋FY2027E–FY2032E（情境 B 基準；FY2027 為錨點＝" + K.get("ANCH", "1H26 實績＋3Q 營收實績＋假設") + "）。US$B；EPS 美元。"
                 "只引用本頁上方、Hist、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, K["GAAP_NOTE"])
    r_note = r; r += 1
    P.cell(r, 1, "FY2027E＝錨點；FY2028E–FY2030E＝情境 B 由下而上；FY2031E–FY2032E＝延伸期（收斂法只推營收與淨利，營業利益取 DCF 頁情境 B＝淨利 ÷ 淨利係數）。"
                 "延伸期毛利率沿用前一年，據以算出營業成本與毛利；營業費用＝毛利 − 營業利益（殘差，不是假設；外部稽核優化 W3-2）；隱含營業費用率見表末揭露列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + [f"FY{y}{'A' if y <= K.get('LAST_A', 2026) else 'E'}" + ("（對帳中）" if y in dep else "") for y in YR] + ["說明"]
    for j, h in enumerate(hdr, 1):
        c = P.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r_hdr = r; r += 1
    L = {}

    def row(key, label, unit, cells, fmt, note=""):
        nonlocal r
        P.cell(r, 1, label); P.cell(r, 2, unit)
        for col, v in zip(COLS, cells):
            if v is None: continue
            P[f"{col}{r}"] = v; P[f"{col}{r}"].number_format = fmt
        if note: P.cell(r, 12, note)
        L[key] = r; r += 1
    M = "#,##0.0"; PCT = "0.0%"
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFG"]            # FY2027E–FY2030E（PL 情境 B 區塊 D–G 欄）
    xe = lambda key: [f"=PL!H{X[key]}", f"=PL!I{X[key]}"]            # FY2031E–FY2032E（延伸表或摘要區塊 H、I 欄）
    h3 = lambda k: [f"=h_{k}_24", f"=h_{k}_25", f"=h_{k}_26"]
    row("rev", "營收", "US$B", h3("rev") + fc("rev") + xe("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_23-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    if R["gp"] is None:   # 模型不推毛利與營業費用（只推營業利益）：預測期營業費用＝營收 × 最近一年實績營業費用率，毛利＝營業利益＋營業費用（展示用拆分，不回饋）
        r_c = r
        ox = "h_opx_26/h_rev_26"
        row("cogs", "營業成本（COGS）", "US$B", h3("cogs") + [f"={c}{L['rev']}-{c}{r_c + 1}" for c in "FGHIJK"], M, "預測＝營收 − 毛利")
        row("gp", "毛利", "US$B", h3("gp") + [f"={c}{r_c + 4}+{c}{r_c + 3}" for c in "FGHIJK"], M, "預測＝營業利益＋營業費用（倒推；展示用拆分）")
        row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT)
        r_o = r
        row("opx", "營業費用（銷管＋研發）", "US$B", h3("opx") + [f"={c}{L['rev']}*{ox}" for c in "FGHIJK"], M,
            "預測＝營收 × FY2026 實績營業費用率（模型只推營業利益；展示用拆分，不回饋）")
    else:
      r_c = r
      row("cogs", "營業成本（COGS）", "US$B", h3("cogs") + [f"=F{L['rev']}-F{r_c + 1}", f"=G{L['rev']}-G{r_c + 1}", f"=H{L['rev']}-H{r_c + 1}", f"=I{L['rev']}-I{r_c + 1}",
                                                            f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-J{r_c + 2})"], M,
          "FY2031E、FY2032E＝營收 ×（1 − 前一年毛利率）")
      row("gp", "毛利", "US$B", h3("gp") + fc("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
      row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2031E、FY2032E＝前一年（沿用）")
      r_o = r
      opx_fc = fc("opx") if R["opx"] else [f"=PL!{c}{R['gp']}-PL!{c}{R['oi']}" for c in "DEFG"]
      row("opx", "營業費用（銷管＋研發）", "US$B", h3("opx") + opx_fc + [None] * 2, M,
          "FY2031E、FY2032E＝毛利 − 營業利益（殘差，不是假設；隱含費用率見表末揭露列）" + ("" if R["opx"] else "；預測期營業費用＝毛利 − 營業利益（模型以營業費用成長推導，與此式相同）"))
    row("oi", "營業利益", "US$B", h3("oi") + fc("oi") + ([f"=DCF!F{d_ebit}+{K['SBC_ADD']}*J{L['rev']}", f"=DCF!G{d_ebit}+{K['SBC_ADD']}*K{L['rev']}"] if K.get("SBC_ADD") else [f"=DCF!{K.get('DCFX', 'FG')[0]}{d_ebit}{K.get('OIDIV', '')}", f"=DCF!{K.get('DCFX', 'FG')[1]}{d_ebit}{K.get('OIDIV', '')}"]), M,
        "FY2031E、FY2032E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）" + ("＋股票薪酬（DCF 的 EBIT 已扣股票薪酬，本表為 non-GAAP 口徑加回）" if K.get("SBC_ADD") else ""))
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    ro = "(h_pti_26-h_oi_26)/h_oi_26"
    if K.get("OTH"):   # 模型有「其他收益」列：FY2027E–FY2030E 取模型列，FY2031E–FY2032E 沿用 FY2030E
        _ro = lab(P, K["OTH"], rb, exact=True)
        row("oth", "其他收入（淨額：利息、匯兌、投資收益等）", "US$B", [f"=h_pti_{y}-h_oi_{y}" for y in ("24", "25", "26")] + [f"=PL!{s}{_ro}" for s in "DEFG"] + ["=I{0}".format(r), "=I{0}".format(r)], M,
            "歷史＝稅前利益 − 營業利益；FY2027E–FY2030E＝模型「其他收益」列；FY2031E–FY2032E＝FY2030E（沿用）")
    else:
      row("oth", "其他收入（淨額：利息、匯兌、投資收益等）", "US$B", [f"=h_pti_{y}-h_oi_{y}" for y in ("24", "25", "26")] + [f"={c}{L['oi']}*{ro}" for c in "FGHIJK"], M,
          "歷史＝稅前利益 − 營業利益；預測＝營業利益 × FY2026 實績比例（展示用拆分）")
    row("pti", "稅前利益", "US$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    r_t = r
    nm = K["nci_mode"]
    if nm == "none":
        tax_f = [f"={c}{L['pti']}-{c}{r_t + 2}" for c in "FGHIJK"]
        nci_f = ["=0"] * 6
        tax_note, nci_note = "預測＝稅前利益 − 淨利（無少數股權；由模型淨利倒推）", "預測＝0（合併財報無少數股權損益）"
    elif nm == "model":
        # 緯創：FY2027E–FY2030E 取模型「合併淨利（含非控制權益）」與「非控制權益」列；FY2031E–FY2032E 無模型列，沿用 FY2030E 非控制權益 ÷ 稅前利益
        tax_f = [f"={c}{L['pti']}-PL!{s}{R['cons']}" for c, s in zip("FGHI", "DEFG")] + [f"={c}{L['pti']}-{c}{r_t + 1}-{c}{r_t + 2}" for c in "JK"]
        nci_f = [f"=PL!{s}{R['nci']}" for s in "DEFG"] + [f"={c}{L['pti']}*$I{r_t + 1}/$I{L['pti']}" for c in "JK"]
        tax_note = "FY2027E–FY2030E＝稅前利益 − 模型合併淨利（含非控制權益）；FY2031E–FY2032E＝稅前利益 − 少數股權損益 − 淨利"
        nci_note = "FY2027E–FY2030E＝模型非控制權益（（1 − 持股）× 緯穎淨利）；FY2031E–FY2032E＝稅前利益 × FY2030E 比例（模型無延伸期此列）"
    else:
        tax_f = [f"={c}{L['pti']}-{c}{r_t + 1}-{c}{r_t + 2}" for c in "FGHIJK"]
        nci_f = [f"={c}{L['pti']}*h_nci_26/h_pti_26" for c in "FGHIJK"]
        tax_note, nci_note = "預測＝稅前利益 − 少數股權損益 − 淨利（由模型淨利倒推）", "預測＝稅前利益 × FY2026 實績比例（展示用拆分）"
    row("tax", "所得稅", "US$B", h3("tax") + tax_f, M, tax_note)
    row("nci", "少數股權損益", "US$B", h3("nci") + nci_f, M, nci_note)
    row("ni", "稅後淨利（歸屬母公司）", "US$B", h3("ni_rep") + fc("ni") + xe("ni"), M, "歷史＝資料來源列示值（Hist 頁）")
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    sh_x = xe("sh") if K["ext_mode"] in ("block", "inblock", "named") else [f"=H{L['ni']}/PL!H{X['eps']}", f"=I{L['ni']}/PL!I{X['eps']}"]
    row("sh", "股數（十億股）", "十億股", h3("sh_b") + fc("sh") + sh_x, "0.000", K.get("SHN", "歷史＝流通股數；預測＝模型股數（不含稀釋）") + ("" if K["ext_mode"] in ("block", "inblock", "named") else "；FY2031E–FY2032E＝淨利 ÷ EPS"))
    row("eps", K.get("EPSL", "EPS（基本；美元）"), "US$", h3("eps_b") + fc("eps") + xe("eps"), "0.00", K.get("EPSN", "歷史＝資料來源列示基本 EPS；預測＝淨利 ÷ 模型股數"))
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/h_ni_rep_23-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2024 的基期＝FY2023 淨利（Hist 頁）")
    row("etr", "有效稅率（所得稅 ÷ 稅前利益；預測為隱含值）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT)
    P.cell(r, 1, K.get("DILH", "稀釋 EPS 對照（歷史；預測不估稀釋）")).font = BOLD; r_dil = r; r += 1
    row("g_epsd", K.get("EPSDL", "EPS（稀釋；資料來源列示值）"), "US$", h3("eps_d"), "0.00", "稀釋加權股數見 Hist 頁")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "US$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2024A–FY2030E；歷史有四捨五入差，≤ 0.001）", "US$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    # 外部稽核優化 W3-2（2026-11-23；Andy 決定 J3）：延伸期（FY2031E、FY2032E）營業費用改為殘差＝毛利 − 營業利益，使「毛利 − 營業費用 − 營業利益＝0」每一年都成立。
    # 原因：原本營業費用＝營收 × 沿用的費用率，營業利益卻由淨利反推，兩者對不上（差額占營收約 1–4%）。只改營業費用這一列的呈現；毛利、營業利益、淨利、EPS、目標價都不動。
    # 隱含的營業費用率列在下方揭露列（不是新假設，是殘差算出來的結果）。
    if R["gp"] is not None:   # 模型只推營業利益的公司：營業費用＝營收 × 實績費用率、毛利＝營業利益＋營業費用，恆等式本來就成立，不覆寫（覆寫會循環參照）
        for c in "JK":
            P[f"{c}{L['opx']}"] = f"={c}{L['gp']}-{c}{L['oi']}"; P[f"{c}{L['opx']}"].number_format = P[f"I{L['opx']}"].number_format
    row("c2b", "核對：延伸期毛利 − 營業費用 − 營業利益（FY2031E、FY2032E；延伸期營業費用＝毛利 − 營業利益，應為 0）", "US$B", [None] * 7 + [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in "JK"], "0.000")
    row("c2c", ("揭露（不是核對列）：營業費用率＝營業費用 ÷ 營收（預測期營業費用＝營收 × 實績費用率、毛利＝營業利益＋營業費用，展示用拆分，各年相同）" if R["gp"] is None else
                "揭露（不是核對列）：隱含營業費用率＝營業費用 ÷ 營收（FY2031E、FY2032E 的營業費用是殘差，不是假設；FY2030E 欄供對照）"), "%", [None] * 6 + [f"={c}{L['opx']}/{c}{L['rev']}" for c in "IJK"], "0.00%")
    row("c3", "核對：稅前利益 − 所得稅 − 少數股權損益 − 稅後淨利（歷史有四捨五入差，≤ 0.001）", "US$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 摘要（FY2027E–FY2030E）／延伸表（FY2031E–FY2032E）", "US$B",
        [None] * 3 + [f"={c}{L['ni']}-PL!{s}{R['ni']}" for c, s in zip("FGHI", "DEFG")] + [f"=J{L['ni']}-PL!H{X['ni']}", f"=K{L['ni']}-PL!I{X['ni']}"], "0.000")
    for col in "CDEFGHIJK": P.column_dimensions[col].width = max(P.column_dimensions[col].width or 0, 13)
    L.update(hdr=r_hdr, note=r_note, note2=r_note2, ratio=r_ratio, dil=r_dil, chk=r_chk, start=s0, R=R, X=X)
    return L


# ---------- RevDCF：反向 DCF ----------
def revdcf(wb):
    D = wb["DCF"]
    RV = wb.create_sheet("RevDCF", index=wb.sheetnames.index("DCF") + 1)
    b = lab(D, "情境 B（基準）"); dr = lambda p: lab(D, p, b)
    rv, eb, tx, nop, da, cx, wc, fcf, df, pv, ev, ps = (dr("營收"), dr("營業利益 EBIT"), dr("稅率"), dr("NOPAT"), dr("＋折舊攤銷"), dr("− 資本支出"),
                                                          dr("− 營運資金增加"), dr("無槓桿自由現金流 FCFF"), dr("折現因子"), dr("現值（FCFF"),
                                                          dr("企業價值 EV"), dr("每股價值（評價日）"))
    fr, ty = lab(D, "評價日後占該年度比例"), lab(D, "折現年數（期中")
    tvy = lab(D, "終值折現年數")
    V = wb["Val"]
    vb = lambda p: f"Val!$B${lab(V, p)}"
    feps, pmed, dcf_tp, coef = vb("12 個月後的前瞻 EPS（情境加權）"), vb("同業前瞻 P/E 中位數"), vb("DCF 目標價（情境加權）"), vb("相對倍數調整係數")
    nc_cell = f"Val!$B${lab(V, '淨現金滾到評價日')}"
    RV["A1"] = (f"RevDCF — 反向 DCF：現價隱含什麼（v{VER} 起；個股 HTML 版面規格 v1.3）。一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式（DCF 對各年營業利益與營收倍數為線性，用封閉解；WACC 用二分法）。"
                "本頁只讀 DCF、Val、In 頁，無任何頁引用本頁，不影響目標價")
    RV["A1"].font = BOLD
    RV["A2"] = ("讀法：DCF 部分以情境 B（基準）路徑反解「情境 B 的 DCF 價值推到 12 個月後＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數 − 期間股利）；"
                "Comparable 部分反解「約一年後的前瞻 EPS（情境加權）× 倍數＝現價」。反解值超出合理範圍（例如營業利益率高於 100%）時照實列出，不外推。"
                + K["RV_NOTE"])
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
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（US$B）", f"={nc_cell}", "#,##0.0", "Val 頁 0.1 節：淨現金與投資＋stub 期 FCFF − 已宣告股利")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（US$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；US$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=DCF!$C${ps}", "0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）"] + [f"FY{y}" for y in range(2027, 2033)], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    Y = {}

    def yr(key, label, fn, fmt):
        nonlocal r
        RV.cell(r, 1, label)
        for i, dc in enumerate("CDEFGH"):
            c = RV.cell(r, 2 + i, fn(dc, "BCDEFG"[i])); c.number_format = fmt
        Y[key] = r; r += 1
    yr("rev", "營收（US$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvy}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（US$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（US$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（US$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（US$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（US$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}<=0,"現價低於營業利益率為零時的價值：現價甚至沒有給預測期獲利定價",IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（US$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（US$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        c = RV.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1

    def val_at(w):
        return (f"(SUMPRODUCT(DCF!$C${fcf}:$H${fcf},DCF!$C${fr}:$H${fr},1/(1+{w})^DCF!$C${ty}:$H${ty})"
                f"+DCF!$H${nop}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv)/(1+{w})^DCF!$C${tvy}+$B${A['nc']})/sh_v")
    b0 = r
    for i in range(40):
        RV.cell(r, 1, i + 1)
        if i == 0:
            RV.cell(r, 2, "=g_tv+0.005"); RV.cell(r, 3, "=0.4")
        else:
            RV.cell(r, 2, f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})"); RV.cell(r, 3, f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        RV.cell(r, 4, f"=(B{r}+C{r})/2"); RV.cell(r, 5, "=" + val_at(f"D{r}"))
        for c in (2, 3, 4): RV.cell(r, c).number_format = "0.0000%"
        RV.cell(r, 5).number_format = "0.00"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式 − DCF 頁情境 B 評價日每股價值（應為 0）", "=" + val_at("wacc") + f"-DCF!$C${ps}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（美元，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "約一年後前瞻 EPS（情境加權；美元）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 美元；要支持現價 "), 'TEXT(px,"0.00")', T(" 美元，企業價值須"), f'IF(B{A["evx"]}<1,"縮小為","放大為")', T("模型的 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}>B{A["w0"]},"升到 ","降到 ")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面"), f'IF(B{A["s1"]}<1,"縮小為 ","放大為 ")',
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億美元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或營業利益率"),
                     f'IF(B{A["m1"]}<B{A["m0"]},"降到 ","提高到 ")', f'TEXT(B{A["m1"]},"0.0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0.0%")', T("，終值年權重大）。"),
                     f'IF(B{A["m1"]}<=0,"現價低於預測期獲利為零時的價值。","")'))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 美元相當於約一年後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
                     T(" 倍，"), f'IF(B{A["pe1"]}<{pmed},"低於","高於")', T("同業中位數 "), f'TEXT({pmed},"0.0")', T(" 倍；模型採用 "), f'TEXT(C{A["pe1"]},"0.0")',
                     T(f" 倍（同業中位數 × {K['cn']}過去十年相對同業的倍數比 "), f'TEXT({coef},"0.00")', T(f"）。換言之，市場目前給{K['cn']}的倍數是同業的 "),
                     f'TEXT(B{A["rc1"]},"0.00")', T(" 倍，歷史相對倍數比為 "), f'TEXT({coef},"0.00")', T("。")))
    A["sum_cmp"] = r; r += 1
    RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 14; RV.column_dimensions["C"].width = 14
    for c in "DEFG": RV.column_dimensions[c].width = 14
    A["yh"] = yh; A["bis0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


def val_g(wb):
    V = wb["Val"]
    r = V.max_row + 2
    V.cell(r, 1, f"0.7 四個目標價並列（v{VER} 新增；個股 HTML 版面規格 v1.3 G 節；只引用本頁既有列，不回饋計算）").font = BOLD; g0 = r; r += 1
    for j, h in enumerate(["項目", "目標價（US$）", "相對現價", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    cons = lab(V, "市場共識目標價", exact=True)
    r_dcf_ = lab(V, "DCF 目標價（情境加權）"); r_cmp0_ = lab(V, "Comparable 目標價（市場主流同業法）"); r_tpw_ = lab(V, "目標價（DCF 與 Comparable 加權）")
    items = [("DCF 目標價（情境加權；進加權）", f"=$B${r_dcf_}", "單純 DCF：三情境無槓桿自由現金流折現，價值推到 12 個月後，依情境權重加權"),
             ("Comparable 目標價（相對倍數調整後；進加權）", "=tp_comp_adj", K["COMP_NOTE"]),
             ("Comparable 目標價（市場主流同業法，未調整；不進加權）", f"=$B${r_cmp0_}", "同業前瞻 P/E 中位數 × 約一年後前瞻 EPS，未做相對倍數調整"),
             ("加權目標價（橫幅主值）", f"=$B${r_tpw_}", "DCF 與 Comparable（調整後）依 DCF 權重加權；投資建議依它的上檔空間判定"),
             ("市場共識目標價", "=cons_tp", f"=Val!E{cons}")]
    for k, (a, f, d) in enumerate(items):
        V.cell(r, 1, a); V.cell(r, 2, f).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, d); r += 1
    V.cell(r, 1, K["PXL"]); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return g0


# ---------- Sens：業務主因快照與加權目標價相對基準 ----------
BIZ = K['BIZ']   # (白話標籤, 低案設定, 高案設定, 鍵)；設定＝{ENV}_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄；x＝基準 × 倍數；mf＝引擎 AI 部署係數）
TOR_LABEL = {"WACC ±1pt": "折現率 WACC 低一個百分點／高一個百分點", "β 低／高": "β（風險係數）取低／高", "ERP 低／高": "股權風險溢酬 4%／6%",
             "永續成長率": "永續成長率低一個百分點／上限", "RONIC": "永續期新投資報酬率提高（只有高）", "資本支出 ÷ 營收": "資本支出占營收 −1／＋1 個百分點",
             "營運資金 ÷ 營收": "營運資金占營收 8%／18%", "淨現金不含權益投資": "淨現金不含權益投資（只有低）", "終值改用 FCFF": "終值改用 FCFF ×（1＋g）替代法（只有高）",
             "美元口徑折現": "改用美元口徑折現（只有低）", "DCF 權重": "DCF 權重 0.3／0.7"}


def sens_post(wb, here):
    SE = wb["Sens"]; V = wb["Val"]
    f = os.path.join(here, "sens.json")
    S = json.load(open(f, encoding="utf-8")) if os.path.exists(f) else None
    if S is None:
        assert os.environ.get(K["ENV"] + "_OVR") or os.environ.get(K["ENV"] + "_NOSENS"), f"{K['folder']}/sens.json 不存在：先跑 python3 {K['folder']}/sens_{K['folder']}.py"
    r = SE.max_row + 3
    SE.cell(r, 1, f"目標價敏感度：業務主因（v{VER} 新增；每個案例以 {K['ENV']}_OVR 把 In 頁輸入改為敏感度低／高值後整本重建、LibreOffice 重算，讀加權目標價；{K['folder']}/sens_{K['folder']}.py → {K['folder']}/sens.json）").font = BOLD
    t0 = r; r += 1
    for j, h in enumerate(["案例", "設定（In 輸入）", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "快照狀態"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    tpc = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"
    rows = {}
    for lab_, lo, hi, key in BIZ:
        SE.cell(r, 1, lab_); SE.cell(r, 2, f"低：{lo}；高：{hi}（lo／hi＝In 頁敏感度低／高欄；x＝基準 × 倍數）")
        e = (S or {}).get("biz", {}).get(key)
        if e:
            for col, side in (("C", "低"), ("E", "高")):
                if side in e:
                    SE[f"{col}{r}"] = e[side]["tp"]; SE[f"{col}{r}"].number_format = "#,##0.00"; SE[f"{col}{r}"].font = BLUE
                    SE[f"{chr(ord(col) + 1)}{r}"] = f"={col}{r}/G{r}-1"; SE[f"{chr(ord(col) + 1)}{r}"].number_format = "+0.0%;-0.0%"
        SE[f"G{r}"] = f"={tpc}"; SE[f"G{r}"].number_format = "#,##0.00"
        SE[f"H{r}"] = f'=IF(ABS(G{r}-{(S or {}).get("_base", {}).get("tp", 0)!r})<=1E-6*ABS(G{r}),"CURRENT","STALE")'
        rows[key] = r; r += 1
    SE.cell(r, 1, (f"快照狀態：快照基準加權目標價與本頁即時值相同為 CURRENT；改 In 頁輸入或模型後須重跑 python3 {K['folder']}/sens_{K['folder']}.py，再正式建置。"
                   "每個案例只改該列輸入，其他輸入維持基準；三情境與 DCF、Comparable 整本重算。")); r += 2
    # 加權目標價相對基準（HTML 敏感度主圖用）
    SE.cell(r, 1, "加權目標價相對基準（HTML 敏感度主圖用；v" + VER + " 起，版面規格 v1.3 H 節；只引用上方「業務主因」與「目標價參數」兩表，白話標籤，不回饋計算）").font = BOLD; tt = r; r += 1
    for j, h in enumerate(["案例（白話）", "低案：加權目標價相對基準", "高案：加權目標價相對基準", "來源列"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for lab_, lo, hi, key in BIZ:
        sr = rows[key]
        SE.cell(r, 1, lab_)
        for col, scol in ((2, "D"), (3, "F")):
            if SE[f"{scol}{sr}"].value is not None:
                SE.cell(r, col, f"=Sens!{scol}{sr}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{sr}"); r += 1
    tp0 = lab(SE, "目標價參數敏感度"); hr = lab(SE, "參數", tp0, exact=True); q = hr + 1
    while SE.cell(q, 1).value and not str(SE.cell(q, 1).value).startswith("檢查"):
        t = str(SE.cell(q, 1).value)
        labl = next((v for k, v in TOR_LABEL.items() if t.startswith(k)), t)
        assert not re.search(r"(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+", labl), labl
        SE.cell(r, 1, labl)
        for col, scol in ((2, "D"), (3, "F")):
            v = SE[f"{scol}{q}"].value
            if v is None or v == "—": continue
            SE.cell(r, col, f"=Sens!{scol}{q}").number_format = "+0.0%;-0.0%"
        SE.cell(r, 4, f"Sens!A{q}"); r += 1; q += 1
    return tt



# ---------- README：v1.3 展示層白話列 ----------
def readme13(wb, SIL, PLA, RVA, coord_rows):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; RC = wb["Recon"]; D = wb["DCF"]; P = wb["PL"]
    vb = lambda p, **k: f"Val!B{lab(V, p, **k)}"
    TPW, UP, REC = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    DCF, CMP, CMP0, PX = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("Comparable 目標價（市場主流同業法）"), vb("現價（")
    NTM12 = f"Val!B{lab(V, '現價 ÷ 未來 12 月 EPS')}"
    sv = lambda p, c: f"Sens!{c}{lab(SE, p)}"
    ra, rb_, rc_ = lab(P, "情境 A（樂觀）："), lab(P, "情境 B（基準）："), lab(P, "情境 C（保守）：")
    eps = lambda r0, c: f"PL!{c}{lab(P, 'non-GAAP EPS', r0, exact=True)}"
    rcn = lambda p: lab(RC, p)
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（v{VER} 起；個股 HTML 版面規格 v1.3：橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字；公式帶數字；原展示層各列保留）").font = BOLD
    r += 1
    M = {}

    def put(key, f):
        nonlocal r
        RD.cell(r, 1, f); M[key] = r; r += 1
    e28 = f"Recon!F{rcn('FY2028 EPS vs 共識')}"
    put("hero", F(f'{REC}', T("：加權目標價 "), f'TEXT({TPW},"0")', T(" 美元，比現價 "), f'TEXT({PX},"0")', T(" 美元"), f'IF({UP}>=0,"高 ","低 ")', f'TEXT(ABS({UP}),"0%")',
                  T("；DCF "), f'TEXT({DCF},"0")', T(" 美元與 Comparable（相對倍數調整後）"), f'TEXT({CMP},"0")', T(" 美元各占一半。AI 伺服器營收由 GPU 系統市場 × Dell 份額推得，份額下修與營益率走低壓抑估值。")))
    put("rec", F(T("投資建議依加權目標價相對現價的上檔空間判定：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、低於 −50% 強力放空。目前上檔空間 "),
                 f'TEXT({UP},"+0.0%;-0.0%")', T("，投資建議 "), f'{REC}', T("。")))
    put("o1", F(T("1｜目標價 "), f'TEXT({TPW},"0")', T(" 美元、上檔 "), f'TEXT({UP},"+0%;-0%")', T("：DCF "), f'TEXT({DCF},"0")', T(" 美元、Comparable（調整後）"), f'TEXT({CMP},"0")',
                T(" 美元各占一半；現價為本模型未來 12 個月 EPS 的 "), f'TEXT({NTM12},"0.0")', T(" 倍。只看市場主流同業法的 Comparable 為 "), f'TEXT({CMP0},"0")', T(" 美元。")))
    put("o2", F(T("2｜AI 伺服器份額與部署量是最大的兩個假設：份額相對值降到 0.75／0.65／0.55 時 FY2028 EPS "), f'TEXT({sv("份額流失（低）", "F")},"+0%;-0%")',
                T("；商用 GPU 顆數 ×0.8 時 "), f'TEXT({sv("AI 部署 ×0.8", "F")},"+0%;-0%")', T("。份額以公司 FY2027 AI 伺服器營收指引（740 億美元）反推。")))
    put("o3", F(T("3｜三情境從 FY2029 才分開：FY2029 EPS（non-GAAP）樂觀、基準、保守約 "), f'TEXT({eps(ra, "F")},"0.00")', T("、"), f'TEXT({eps(rb_, "F")},"0.00")', T("、"), f'TEXT({eps(rc_, "F")},"0.00")',
                T(" 美元；FY2028 EPS 三情境相同，比法人共識 "), f'TEXT({e28},"+0%;-0%")', T("。")))
    put("o4", F(T("4｜DCF 為標準版（無槓桿自由現金流折現，EBIT 換成 GAAP 口徑）：折現率 "), 'TEXT(wacc,"0.0%")', T("、永續成長率 "), 'TEXT(g_tv,"0.00%")', T("，終值現值占情境 B 企業價值 "),
                f'TEXT(DCF!C{lab(D, "終值現值占 EV", lab(D, "情境 B（基準）"))},"0%")', T("；折現率低／高一個百分點時加權目標價 "), f'TEXT({sv("WACC ±1pt", "D")},"+0%;-0%")', T("／"),
                f'TEXT({sv("WACC ±1pt", "F")},"+0%;-0%")', T("。")))
    if coord_rows:
        CR = coord_rows
        put("o5", F(T("5｜對帳中：Siliconomics 標示 CY"), 'Coord!B' + str(CR["yrs"]), T(" 對帳中；另一側讀數（Wattonomics 換算的部署量）目標價 "), f'TEXT(Coord!C{CR["tp"]},"0")', T(" 美元（"), f'Coord!C{CR["rec"]}', T("），現行 "),
                    f'TEXT(Coord!B{CR["tp"]},"0")', T(" 美元（"), f'Coord!B{CR["rec"]}', T("）；個股不裁決。公司實績（AI 伺服器營收）對層級：份額同源校準，無可比實績。")))
    else:
        put("o5", "5｜本次建置不含 Coord 頁（三層協調兩讀數快照重建用）。")
    pr = lambda p: lab(P, p, rb_, exact=True)
    put("drv", F(T("分部營收怎麼推導（情境 B）：AI 伺服器營收＝商用 GPU 系統市場（GPU 顆數 × 每顆系統價值）× Dell 有效份額（FY2027 以公司指引校準）；FY2028 為 "),
                 f'TEXT(PL!E{pr("AI 伺服器營收")},"#,##0.0")', T(" 十億美元。傳統伺服器與網通、儲存、PC 由 FY2027 錨點乘成長率；FY2028 分別為 "), f'TEXT(PL!E{pr("傳統伺服器與網通")},"#,##0.0")', T("、"),
                 f'TEXT(PL!E{pr("儲存")},"#,##0.0")', T("、"), f'TEXT(PL!E{pr("CSG 營收（PC）")},"#,##0.0")', T(" 十億美元。合計＝合併營收 "), f'TEXT(PL!E{PLA["R"]["rev"]},"#,##0.0")', T(" 十億美元。")))
    put("ver", F(T("v0.2.6（2026-11-15）：版面規格 v1.3 全面改版（章節順序、橫幅 ≤200 字、重點觀察 ≤5 點、標準損益表 FY2024A–FY2032E、反向 DCF 頁、四個目標價與現價圖、敏感度改加權目標價相對基準低案／高案主圖）；"
                 "三層協調：Coord 頁並列兩讀數目標價，並新增 C0（實績）：1Q–2Q FY27 AI 伺服器營收（10-Q）對層級隱含值，列兩讀數但不判定（Dell 份額由公司指引校準，無可比實績）；同源列標「讀取層級（同源，不判定）」。"
                 "xlsx 只新增 Hist、RevDCF、Coord 頁與各頁末新增區塊，既有數值、加權目標價與投資建議不變。"
                 "舊 README 內文（下方各節）與 v0.2.5 沿革節的數字早於 Siliconomics v0.19 全下游重建，與現行不符；現行值以 Val 頁與本展示層為準。")))
    return M


# ---------- Coord：C0（實績）——公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.2.6 新增） ----------
QTRS = [   # (季, AI 伺服器營收 US$B, 來源, 說明)
    ("1Q FY27（2026-02～04）", "=Q!C4-{q2}", Q10_227, "＝上半年 32.533 − 2Q 16.401（10-Q 2Q FY27：六個月與三個月 AI-optimized servers）"),
    ("2Q FY27（2026-05～07）", 16.401, Q10_227, "AI-optimized servers 16,401 百萬美元（10-Q 2Q FY27 ISG 營收表；去年同期 8,208）"),
]
VERIF = "Verified（公司 10-Q，SEC EDGAR）"


def c0_actuals(wb, C, r, lay):
    out = {"sec": r}
    C.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.2.6 新增）。結論：無可比實績（層級值同源）——層級沒有 Dell 份額的獨立來源，Dell 有效份額由本模型以公司 FY2027 AI 伺服器營收指引校準；下表只作參考，不判定、不計對帳中").font = BOLD; r += 1
    C.cell(r, 1, ("方法：公司實績側＝公司揭露的 AI-optimized servers 營收（10-Q ISG 營收表，每季揭露）。"
                  "層級側＝商用 GPU 系統市場（Siliconomics 基準層，CY2026）× Dell 有效份額＝CY2026 全年隱含 AI 伺服器營收（Dell FY2027＝2026-02～2027-01 ≈ CY2026）。"
                  "主讀數＝FY2027 已公布期間（1Q–2Q，2026-02～07，6 個月）累計 vs 層級 × 6 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季（2Q FY27，2026-05～07）× 4 vs 層級全年。"
                  "2Q FY27 季中點（6 月中）接近年中，第二讀數大致無偏；成長年主讀數偏低。"))
    r += 1
    C.cell(r, 1, ("為什麼不判定：Dell 有效份額＝公司 FY2027 AI 伺服器營收指引（740 億美元，Interested-party）÷ 商用 GPU 系統市場，所以層級全年隱含值恆等於公司指引（見下方核對列），"
                  "GPU 顆數與每顆系統價值在比式中消去；本表不檢驗 Siliconomics 的 GPU 系統市場，也不得讀成層級獲得獨立驗證。"
                  "本表實際檢驗的是錨點（公司指引）：(a) 上半年 AI 伺服器營收占全年指引的比例；(b) 指引隱含下半年 AI 伺服器營收對上半年的成長。"
                  "3Q FY27 財報（約 11 月下旬）尚未公布。"))
    r += 1
    for j, h in enumerate(["項目", "主讀數（1Q–2Q FY27 累計）", "第二讀數（2Q FY27 × 4）", "單位與口徑", "來源與證據等級", "出處網址（只在 xlsx）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    out["head"] = r; r += 1
    C.cell(r, 1, "4.1 公司實績：AI-optimized servers 營收（US$B）").font = BOLD; r += 1
    m0 = r
    for q, v, src, note in QTRS:
        C.cell(r, 1, f"　{q} AI 伺服器營收")
        C[f"B{r}"] = v.format(q2=f"B{m0 + 1}") if isinstance(v, str) else v; C[f"B{r}"].number_format = "#,##0.000"
        if not isinstance(v, str): C[f"B{r}"].font = BLUE
        C.cell(r, 4, "US$B；ISG AI-optimized servers（公司揭露）；" + note); C.cell(r, 5, VERIF if not isinstance(v, str) else "Derived（公司揭露的六個月減三個月）"); C.cell(r, 6, src)
        r += 1
    m1 = r - 1
    out["m0"], out["m1"] = m0, m1
    C.cell(r, 1, "　1Q–2Q FY27 累計（主讀數用）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m1})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "US$B；10-Q 六個月 AI-optimized servers 32,533 百萬美元"); C.cell(r, 5, "Verified（加總）"); cum = r; r += 1
    C.cell(r, 1, "　核對：上半年合計 − Q 頁 1H FY27 AI 伺服器（32.533）"); C[f"B{r}"] = f"=B{cum}-Q!C4"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "US$B；應為 0"); C.cell(r, 5, "核對列"); ck1 = r; r += 1
    q3 = m1
    out["cum"], out["q3"] = cum, q3
    C.cell(r, 1, "4.2 口徑換算參數（公司揭露與指引；參考）").font = BOLD; r += 1
    C.cell(r, 1, "　參考：公司 FY2027 AI 伺服器營收指引（In 頁 g_ai，本模型錨點）"); C[f"B{r}"] = "=g_ai"; C[f"B{r}"].number_format = "#,##0.0"
    C.cell(r, 4, "US$B；公司指引（非實績）；為 Dell 份額的校準值"); C.cell(r, 5, "Interested-party"); g_ai = r; r += 1
    C.cell(r, 1, "　參考：2Q FY27 ISG 營收（公司揭露）"); C[f"B{r}"] = 31.782; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "US$B；10-Q：AI 16,401＋傳統伺服器與網通 10,531＋儲存 4,850 百萬美元"); C.cell(r, 5, VERIF); C.cell(r, 6, Q10_227); r += 1
    C.cell(r, 1, "4.3 對帳（實績側 vs 層級隱含）").font = BOLD; r += 1

    def row2(label, fb, fc, fmt, unit, src, bold=False):
        nonlocal r
        C.cell(r, 1, label)
        C[f"B{r}"] = fb; C[f"C{r}"] = fc
        for c in "BC": C[f"{c}{r}"].number_format = fmt
        C.cell(r, 4, unit); C.cell(r, 5, src)
        if bold:
            for c in "ABC": C[f"{c}{r}"].font = BOLD
        r += 1; return r - 1
    a_rev = row2("　公司實績：AI 伺服器營收（主：1Q–2Q FY27 累計；第二：2Q FY27 × 4）", f"=B{cum}", f"=B{q3}*4", "#,##0.00", "US$B；主讀數為 6 個月累計、第二讀數為年化", VERIF, True)
    l_mk = row2("　層級：商用 GPU 系統市場 CY2026（Siliconomics 基準層）", f"=B{lay['l_mk']}", f"=B{lay['l_mk']}", "#,##0.0", "US$B；連到下方「讀取層級」第 (1) 項的 Siliconomics 層列", "Siliconomics v0.22（商用 GPU 系統市場基準值）；Derived")
    l_sh = row2("　× Dell 有效份額（FY2027 校準）", f"=PL!D{lay['sh0']}", f"=PL!D{lay['sh0']}", "0.00%", "＝公司 FY2027 AI 伺服器營收指引 ÷ 市場（同源）", "Derived（由本模型 FY2027 錨點校準）")
    l_fy = row2("　＝ 層級隱含 CY2026 全年 AI 伺服器營收", f"=B{l_mk}*B{l_sh}", f"=C{l_mk}*C{l_sh}", "#,##0.00", "US$B；全年（≈ Dell FY2027）", "Derived", True)
    l_tf = row2("　× 期間係數（主：已公布月數 ÷ 12；第二：1）", "=6/12", "=1", "0.000", "線性時差近似（層級值為日曆年）", "Derived")
    l_pd = row2("　＝ 層級側同期隱含 AI 伺服器營收", f"=B{l_fy}*B{l_tf}", f"=C{l_fy}*C{l_tf}", "#,##0.00", "US$B", "Derived", True)
    out["fy"] = l_fy
    C.cell(r, 1, "　差距（實績側 ÷ 層級側 − 1）"); C[f"B{r}"] = f"=B{a_rev}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_rev}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"; C[f"{c}{r}"].font = BOLD
    C.cell(r, 4, "正值＝公司實績高於層級隱含"); out["d"] = r; d = r; r += 1
    C.cell(r, 1, "　判定：無可比實績（層級值同源）——份額由公司指引同源校準，本表不檢驗 GPU 系統市場；參考（同源校準），不判定、不計對帳中")
    C[f"B{r}"] = "無可比實績（不判定）"; C[f"B{r}"].font = BOLD
    C.cell(r, 4, "缺：Dell 在商用 GPU 系統市場份額的獨立揭露（規則 v1.4.2：沒有可比實績時不以同源值補位）"); C.cell(r, 5, "規則 v1.4.2 C0 第 2 點、工作單 20261114"); out["judge"] = r; r += 1
    C.cell(r, 1, "　參考：兩讀數是否都在 ±15% 內（僅描述，不構成判定）"); C[f"B{r}"] = f'=IF(AND(ABS(B{d})<=0.15,ABS(C{d})<=0.15),"都在 ±15% 內","有讀數超過 ±15%")'
    C.cell(r, 4, "描述用；層級側恆等於公司指引（見下方核對列）"); out["ref_in"] = r; r += 1
    C.cell(r, 1, "　核對：層級隱含 CY2026 全年 AI 伺服器營收 − 公司 FY2027 指引（應為 0）"); C[f"B{r}"] = f"=B{l_fy}-B{g_ai}"; C[f"B{r}"].number_format = "0.000000"
    C.cell(r, 4, "US$B；為 0＝層級隱含值與本模型錨點同源（GPU 系統市場消去）"); C.cell(r, 5, "核對列"); out["ck_same"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (a)：上半年 AI 伺服器營收 ÷ 公司全年指引"); C[f"B{r}"] = f"=B{cum}/B{g_ai}"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "比；6 個月線性＝50%"); C.cell(r, 5, "Derived"); out["ref_a"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (b)：指引隱含下半年 AI 伺服器營收 ÷ 上半年實績 − 1"); C[f"B{r}"] = f"=(B{g_ai}-B{cum})/B{cum}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "比；公司稱季末在手訂單大增（10-Q MD&A），下半年需較上半年成長才能達成指引"); C.cell(r, 5, "Derived"); out["ref_b"] = r; r += 1
    C.cell(r, 1, "　參考（口徑不同，不判定）：Siliconomics 組裝市場 CY2026 未建模殘差（1 − 已建模 ODM 份額；含 Dell、Supermicro 等）"); C[f"B{r}"] = f"={lay['res']}"; C[f"B{r}"].number_format = "#,##0.0"
    C.cell(r, 4, "US$B；ODM 整櫃價口徑（Dell 為品牌出貨價，口徑不同）；只作量級參考"); C.cell(r, 5, "Siliconomics v0.22（整櫃組裝市場基準值 ×（1 − 已建模 ODM 份額合計））；Derived"); out["ref_res"] = r; r += 1
    out["ck"] = (ck1,); out["end"] = r
    C.column_dimensions["A"].width = max(C.column_dimensions["A"].width or 0, 60)
    return out


# ---------- Coord：C0 由下往上對帳 ----------
def coord_c0(wb, SIL, R):
    C = wb["Coord"]; P = wb["PL"]
    C.column_dimensions["F"].width = 46; C.column_dimensions["G"].width = 70
    rb = lab(P, K["PLB"]); g = lambda p: lab(P, p, rb, exact=True)
    N = SIL.NAMES
    r = C.max_row + 3
    C.cell(r, 1, "C0 由下往上對帳（三層基準協調規則 v1.4.2 C0 第 2 點；v0.2.6 新增）：比較基準是公司實績（第 4 節 C0〔實績〕：AI 伺服器營收 vs 層級隱含值）；本模型自層級讀入的同源值只列在第 2 節「讀取層級」，不判定、不計對帳中。差距超過 ±15% 標「對帳中」，不改上游值，由主控開對帳單。").font = BOLD
    c0 = r; r += 1
    C.cell(r, 1, "1. 讀哪一層、哪一版").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "版本／檔案", "本模型用途", "備註"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    flg = "／".join(str(SIL.RECO_FLAG[y]) for y in (2026, 2027, 2028, 2029)); alt = "／".join(f"{SIL.ALT_RATIO[y]:.2f}" for y in (2026, 2027, 2028, 2029))
    wt = "／".join(f"{N['SIF_WTAnchor_Base'][i]:.2f}" for i in (1, 2, 3))
    for a, b, c, d, e in (
        ("硬體股讀哪一層", "Siliconomics（零組件層；基準）", ver, "商用 GPU 顆數、情境顆數係數、Wattonomics 錨定、對帳中旗標；每顆系統價值讀 markets/chip_bridge", "下游只註明版本、不自行重估（專案規則第 3 條）"),
        ("商用 GPU 顆數（GB300 級、VR200 級）", "In 頁 ug_26–ug_29、uv_26–uv_29", "tsmc v0.5.4（In 頁；Siliconomics v0.22 平台顆數）", "PL 商用 GPU 系統市場", "情境 B 的系統市場＝Siliconomics SIF_GPUSysMkt_Base（見第 2 節）"),
        ("每顆系統價值（GB300、VR200 NVL72 機架價 ÷ 72）", "59.7／116.7 US$K／顆", "markets/chip_bridge v0.2.3（橋接表）", "In 頁 v_gb、v_vr", "唯讀"),
        ("情境顆數係數（2028、2029）", f"A 1.0／1.0；B {SIL.COEF_B_TXT}；C 0.6／0.6", "tsmc v0.5.4（In!E37:E42；Siliconomics v0.22 基準係數）", "In 頁 um28_A–um29_C（FY(t) 用 CY(t−1)）", "情境 B＝Siliconomics 基準"),
        ("Wattonomics 錨定 SIF_WTAnchor_Base（wta_27–29）", wt, ver, "只用於 Sens「Wattonomics 錨定」案例，不進目標價", "Wattonomics 現行版本 v0.3.2（本模型不直接讀）"),
        ("對帳中旗標 SIF_RecoFlag_<yyyy>、SIF_DeployGW_AltRatio_<yyyy>", f"旗標 {flg}；比值 {alt}", ver, "C4 兩讀數（本頁上方）", "另一側讀數＝Wattonomics 換算的產業 GW；本模型 FY2028–FY2030 的 GPU 顆數同乘 CY2027–CY2029 比值")):
        for j, v in enumerate((a, b, c, d, e), 1):
            C.cell(r, j, v)
        r += 1
    r += 1
    C.cell(r, 1, "2. 讀取層級（同源，不判定；Dell FY(t)≈CY(t−1)：FY2027–FY2030 對 CY2026–CY2029）：本模型自層級讀入的值，差距恆為 0，無檢驗力，故不進 C0 判定（規則 v1.4.2）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "CY2026（FY2027）", "CY2027（FY2028）", "CY2028（FY2029）", "CY2029（FY2030）", "單位與口徑", "來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    flags = []
    YC = "BCDE"

    def block(title, model_fs, layer_vals, unit, src, fmt="#,##0.0", model_label="本模型", layer_label="Siliconomics 層", tol=0.15, ref=False):
        nonlocal r
        C.cell(r, 1, title).font = BOLD; r += 1
        C.cell(r, 1, f"　{model_label}"); C.cell(r, 6, unit); C.cell(r, 7, src)
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = model_fs[k]; C[f"{c}{r}"].number_format = fmt
        m = r; r += 1
        C.cell(r, 1, f"　{layer_label}")
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = layer_vals[k]; C[f"{c}{r}"].number_format = fmt; C[f"{c}{r}"].font = BLUE
        l = r; r += 1
        C.cell(r, 1, f"　差距（{model_label} ÷ {layer_label} − 1）" if ref else "　差距（本模型 ÷ 層 − 1）")
        for c in YC:
            C[f"{c}{r}"] = f"={c}{m}/{c}{l}-1"; C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
        d = r; r += 1
        if ref:
            C.cell(r, 1, f"　判定（參考：|差距| > {tol:.0%}＝層級差距；不標本公司對帳中）")
            for c in YC:
                C[f"{c}{r}"] = f'=IF(ABS({c}{d})>{tol},"層級差距","一致")'
            r += 1
        else:
            C.cell(r, 1, "　判定（讀取層級，同源，不判定；不計入對帳中項目數）")
            for c in YC:
                C[f"{c}{r}"] = "同源（不判定）"
            flags.append(r); r += 1
        return m, l, d
    m1_, l1_, d1_ = block("(1) 商用 GPU 系統市場（情境 B；US$B）", [f"=PL!{c}{g('商用 GPU 系統市場（顆數 × 每顆系統價值）')}" for c in "DEFG"], list(N["SIF_GPUSysMkt_Base"]),
                          "US$B；商用 GPU 顆數 × 每顆系統價值 × 顆數係數", "Siliconomics 基準層（商用 GPU 系統市場）；Derived")
    r += 1
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; r += 1
    C.cell(r, 1, "　下列差距是 Wattonomics 與 Siliconomics 兩層之間的差距，不是 Dell 自己的量對其所讀層級值的差距；本公司依 C4 並列兩讀數，不裁決、不改上游。"); r += 1
    m5, l5, d5 = block("(2) 產業部署 GW：Siliconomics 基準對 Wattonomics 換算值（兩層之間的差距）", list(N["SIF_DeployGW_Base"]), [N[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)],
                       "GW；Siliconomics 基準部署 GW 對 Wattonomics 資本支出與電力路徑換算的需求側讀數", "Wattonomics v0.3.8 經 Siliconomics 換算；Derived（需求側）", "0.00",
                       model_label="Siliconomics 基準", layer_label="Wattonomics 換算值（需求側）", ref=True)
    res = N["SIF_AsmMkt_Base"][0] * (1 - N["SIF_AsmShareSum"][0])
    lay = dict(sh0=lab(P, "Dell 份額（FY2027 校準）", rb), l_mk=l1_, res=f"{N['SIF_AsmMkt_Base'][0]!r}*(1-{N['SIF_AsmShareSum'][0]!r})")
    r += 1
    C.cell(r, 1, "3. 結論（C0；判定見第 4 節 C0〔實績〕表）").font = BOLD; r += 1
    C.cell(r, 1, "對帳中項目數（C0〔實績〕判定為「對帳中」者；讀取層級的同源列不判定、不計入）")
    n_other = r; r += 1
    C.cell(r, 1, "C0 結論（白話）")
    c0_concl = r; r += 2
    A4 = c0_actuals(wb, C, r, lay)
    C[f"B{n_other}"] = f'=IF(B{A4["judge"]}="對帳中",1,0)'
    C[f"B{c0_concl}"] = F(T("C0（實績）：無可比實績（層級值同源）——層級沒有 Dell 份額的獨立來源，不判定、不計對帳中（對帳中項目數 0）。"
                          "Dell 有效份額是本模型以公司 FY2027 AI 伺服器營收指引（740 億美元）反推，層級隱含值恆等於指引，GPU 系統市場在比式中消去，不得讀成層級獲得獨立驗證。"
                          "第 4 節表實際檢驗的是錨點：上半年 AI 伺服器營收已達全年指引的 "), f'TEXT(B{A4["ref_a"]},"0.0%")',
                          T("；指引隱含下半年比上半年 "), f'TEXT(B{A4["ref_b"]},"+0.0%;-0.0%")',
                          T("。（描述：上半年累計對層級同期隱含值 "), f'TEXT(B{A4["d"]},"+0.0%;-0.0%")', T("、最近一季年化對全年 "), f'TEXT(C{A4["d"]},"+0.0%;-0.0%")', T("。）"
                          "本模型自層級讀入的同源值（第 2 節）不判定。Siliconomics 基準部署 GW 與 Wattonomics 換算的需求側讀數在 CY2027–CY2029 差距 "),
                          f'TEXT(C{d5},"+0%")', T("、"), f'TEXT(D{d5},"+0%")', T("、"), f'TEXT(E{d5},"+0%")',
                          T("：這是兩層之間的差距（見「層級之間」小節），已由 Siliconomics 對帳單 r3 處理，不計入本公司 C0；本模型依 C4 並列兩讀數目標價（上方），不改上游值，不裁決哪一側正確。"))
    return dict(c0=c0, concl=c0_concl, flags=flags, d5=d5, **A4)


# ---------- 版本字串 ----------
def version(wb, here, coord_rows):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    if f"v{VER}" not in head:   # 標題段版號未隨版更新（舊檔）：改為本版（版本沿革文字不動）
        a1 = re.sub(r"v\d+\.\d+(\.\d+)?", f"v{VER}", head, count=1) + a1[len(head):]
    a1 += (f"；v{VER}（{BUILT}）同版另：版面規格 v1.3 全面改版（新增 Hist、RevDCF 頁，PL、Val、Sens、README 頁末新增區塊）與三層協調 C0（實績）補公司實績對帳"
           "（Coord 頁末「C0〔實績〕」：1Q–2Q FY27 AI 伺服器營收 vs 層級隱含值，Dell 份額由公司指引校準→無可比實績，不判定；同源列標「讀取層級（同源，不判定）」）；既有數值、加權目標價與評等不變")
    RD["A1"].value = a1
    return a1


def apply(wb, SIL, here, coord_rows):
    HR = hist_is(wb)
    PLA = pl_standard(wb, SIL)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_post(wb, here)
    C0 = coord_c0(wb, SIL, coord_rows) if coord_rows else None
    if coord_rows and K.get("FY_LAG") is not None:   # 會計年度對應（版面規格 v1.3 I 節；check_layout_v13 讀 co_fy_lag）：寫在 C4 區塊與 C0 之間的空白列（C0 建好後才寫，不移動 C0 列號）
        C = wb["Coord"]; rl = lab(C, "橫幅揭露") + 1
        assert C.cell(rl, 1).value is None and C.cell(rl + 1, 1).value is None
        C.cell(rl, 1, "會計年度對應（FY 減 CY）：Dell FY(t)＝t−1 年 2 月～t 年 1 月，依賴 CY(t−1)；版面規格 v1.3 I 節，check_layout_v13 據此判讀損益表應標「對帳中」的會計年度")
        C.cell(rl, 2, K["FY_LAG"]); dn(wb, "co_fy_lag", f"Coord!$B${rl}")
    RMA = readme13(wb, SIL, PLA, RVA, coord_rows)
    if coord_rows:
        version(wb, here, coord_rows)
    SO = wb["Sources"]; rs = SO.max_row + 1
    SO.cell(rs, 1, "[Verified] Dell 10-Q 2Q FY27（2026-07-31 季）：ISG AI-optimized servers 三個月 16,401、六個月 32,533 百萬美元（C0〔實績〕用）")
    SO.cell(rs, 2, Q10_227)
    SO.cell(rs + 1, 1, "[Verified] Dell non-GAAP 損益表歷史 FY2023–FY2026（Hist 頁；4Q FY24、4Q FY25、4Q FY26 新聞稿調節表；推得列標 Derived）")
    SO.cell(rs + 1, 2, PR426)
    wb._v026 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, C0=C0, RM=RMA)
    return f"v026_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 主圖表 {ST}；README v1.3 {len(RMA)} 列；C0 {'有' if C0 else '無（無 Coord 頁）'}"
