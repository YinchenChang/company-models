# 緯穎 v0.2.6（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7／v0.3.8、廣達 v0.4.8 quanta/v047_post.py）＋三層協調 C0（實績）公司實績對帳（規則 v1.4.2）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.2.5 逐格比對）：
#   Hist 頁（新）：損益表歷史 FY2023–FY2025（S&P Global Market Intelligence，經 stockanalysis.com；FY2022 營收與淨利只供年增率；除權還原後）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B；緯穎只有 IFRS 合併財報、無少數股權，標 GAAP 口徑）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC 與 Comparable 倍數；只讀其他頁，無任何頁引用本頁
#   Val 頁末：0.7 四個目標價並列（HTML 用）
#   Sens 頁末：目標價敏感度：業務主因（整本重算快照 wiwynn/sens.json）與「加權目標價相對基準」白話表（HTML 敏感度主圖用）
#   README 頁末：v1.3 展示層白話列（橫幅說明、重點觀察 ≤5 點、分部營收推導說明、本版要點）；A1 標題段版號更正為 v0.2.6
#   Coord 頁末：C0 由下往上對帳（讀哪一層、哪一版；讀取層級同源列不判定）＋第 4 節 C0（實績）：2026 年 1–10 月月營收 × 公司揭露 AI 占比 vs 層級隱含值（同源校準，無可比實績）
# build_wiwynn.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, SIL, here, coord_rows)。
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.2.6", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SP = ("S&P Global Market Intelligence 經 stockanalysis.com 損益表頁（https://stockanalysis.com/quote/tpe/6669/financials/income-statement/；2026-11-15 讀取，頁面更新 2026-08-09；"
      "百萬元新台幣，此處換算為十億元；EPS 與股數已依 2026-09-02 除權還原，與本模型口徑一致）")
K = dict(
    cn="緯穎", folder="wiwynn", ENV="WIWYNN",
    # (鍵, 項目, FY2023, FY2024, FY2025, FY2022（只供年增率）)；金額十億新台幣、EPS 新台幣（除權還原後）、股數十億股（除權還原後）
    HIS=[
        ("rev", "營收", 241.901, 360.541, 950.663, 292.876),
        ("cogs", "營業成本", 219.244, 323.141, 872.209, None),
        ("gp", "毛利", 22.657, 37.400, 78.454, None),
        ("opx", "營業費用（銷管＋研發）", 6.787, 9.301, 14.537, None),
        ("oi", "營業利益", 15.871, 28.099, 63.918, None),
        ("pti", "稅前利益", 15.444, 28.830, 66.052, None),
        ("tax", "所得稅", 3.400, 6.054, 14.934, None),
        ("nci", "少數股權損益（緯穎無少數股權，記 0）", 0.0, 0.0, 0.0, None),
        ("ni_rep", "歸屬母公司淨利（資料來源列示值）", 12.044, 22.776, 51.118, 14.175),
        ("eps_b", "EPS（基本，新台幣；除權還原後）", 23.09, 42.43, 92.22, None),
        ("eps_d", "EPS（稀釋，新台幣；除權還原後；資料來源列示值）", 23.02, 41.05, 89.02, None),
        ("sh_b", "流通股數（十億股；除權還原後）", 0.522, 0.537, 0.554, None),
        ("sh_d", "稀釋加權股數（十億股；除權還原後）", 0.523, 0.557, 0.576, None),
    ],
    SP=SP, SP22="；FY2022 取同頁（營收 292,876、淨利 14,175 百萬元）",
    PLB="情境 B（基準）：",
    R=dict(rev="營收", gp="毛利", opx="營業費用", oi="營業利益", ni="歸屬母公司淨利", sh="股數（除權還原後）", eps="EPS（除權還原後）"),
    nci_mode="none", ext_mode="inblock",
    GAAP_NOTE=("口徑：GAAP（IFRS 合併財報，歸屬母公司；緯穎無少數股權，歸屬母公司淨利＝稅後淨利）。緯穎不另編 Non-GAAP，市場共識與公司法說皆用合併財報口徑，故本表用 GAAP。"
               "歷史 EPS 為基本 EPS，與模型一致已依 2026-09-02 除權還原（還原係數 2.9828；歷史股數同樣還原）；稀釋 EPS 在表末對照列。"
               "預測期模型只推營業利益與「淨利 ÷ 營業利益」一個係數（涵蓋業外、稅），本表把它拆成：其他收入＝營業利益 × FY2025 實績比例、少數股權損益＝0、"
               "所得稅＝稅前利益 − 淨利（由模型淨利倒推，等於隱含有效稅率）。這幾列是展示用拆分，不回饋計算、不改淨利與目標價。"),
    RV_NOTE=("緯穎 FY2027 三情境相同（顆數係數從 CY2028 起），營收年增約 66%、營運資金大增（無槓桿自由現金流受拖累），所以 FY2027 視為錨點、不隨反解變動"),
    COMP_NOTE="同業前瞻 P/E 中位數 × 緯穎十年相對同業倍數比 × 約一年後前瞻 EPS（情境加權）",
    BIZ=[  # (白話標籤, 低案設定, 高案設定, 鍵)；設定＝WIWYNN_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄；x＝基準 × 倍數）
        ("緯穎份額（相對值 0.85／0.75／0.70 ～ 1.1）", "rel_27=lo;rel_28=lo;rel_29=lo", "rel_27=hi;rel_28=hi;rel_29=hi", "份額"),
        ("ASIC 比重（85% ～ 95%；基準 90%）", "asic_mix=lo", "asic_mix=hi", "ASIC 比重"),
        ("客供比例（全為零 ～ 15%／25%／35%）", "cs_27=lo;cs_28=lo;cs_29=lo", "cs_27=hi;cs_28=hi;cs_29=hi", "客供比例"),
        ("客供部分保留毛利比例（70% ～ 100%）", "keep=lo", "keep=hi", "客供保留毛利"),
        ("每 GW 整櫃價值 2028–29（×0.85 ～ ×1.2）", "vpg_28=lo;vpg_29=lo", "vpg_28=hi;vpg_29=hi", "每 GW 價值"),
        ("AI 部署量（FY2027 起 AI IT 資本支出 ×0.8 ～ ×1.2）", "itc_27=x0.8;itc_28=x0.8;itc_29=x0.8", "itc_27=x1.2;itc_28=x1.2;itc_29=x1.2", "AI 部署量"),
        ("通用伺服器成長（2027 −10%、2028–29 −5% ～ +15%／+10%／+8%）", "g_nai27=lo;g_nai28=lo;g_nai29=lo", "g_nai27=hi;g_nai28=hi;g_nai29=hi", "通用成長"),
        ("AI 毛利率差距（低 2 個百分點 ～ 0）", "d_ai=lo", "d_ai=hi", "AI 毛利率差距"),
        ("GPU 機櫃毛利率稀釋（−0.5 個百分點 ～ 不稀釋）", "vrd=lo", "vrd=hi", "GPU 稀釋"),
        ("營業費用彈性（營業費用年增 ÷ 營收年增：0.5 ～ 1.0；基準 0.8）", "phi=lo", "phi=hi", "營業費用彈性"),
        ("淨利 ÷ 營業利益（0.72 ～ 0.80）", "nio_f=lo", "nio_f=hi", "淨利係數"),
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
    H["A1"] = (f"Hist — {K['cn']}損益表歷史（v{VER} 新增；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣。"
               "FY(t)＝CY(t)。來源為第三方彙整（S&P Global Market Intelligence），未逐項對公司年報原文；與模型 Q 頁 FY2025 可能有口徑或四捨五入差異，見表末對照列")
    H["A1"].font = BOLD
    r = 3
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "FY2022（只供年增率）", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {k: r + i for i, (k, *_) in enumerate(HIS)}
    for k, labl, v23, v24, v25, v22 in HIS:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for col, y, v in (("C", "23", v23), ("D", "24", v24), ("E", "25", v25), ("F", "22", v22)):
            if v is None: continue
            c = H[f"{col}{r}"]; c.value = v; c.font = BLUE
            c.number_format = "0.00" if k.startswith("eps") else "#,##0.000"
            dn(wb, f"h_{k}_{y}", f"Hist!${col}${r}")
        H.cell(r, 7, "Verified／第三方彙整"); H.cell(r, 8, SP + (K["SP22"] if k in ("rev", "ni_rep") else ""))
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
    H.cell(r, 1, "對照：歷史（本頁，S&P Global）− 模型 FY2025A（Q 頁）。非零差異＝兩個來源口徑或四捨五入不同；模型既有 FY2025 欄不改").font = BOLD; r += 1
    for key, text, h_k, a_nm in (("d_rev", "營收", "rev", "a25_rev"), ("d_gp", "毛利", "gp", "a25_gp"), ("d_oi", "營業利益", "oi", "a25_oi"), ("d_ni", "歸屬母公司淨利", "ni_rep", "a25_ni"), ("d_eps", "EPS（基本）", "eps_b", "a25_eps")):
        if a_nm not in wb.defined_names:
            continue
        H.cell(r, 1, f"差異：{text}（本頁 − Q 頁 FY2025）")
        H[f"E{r}"] = f"=E{rows[h_k]}-{a_nm}"; H[f"E{r}"].number_format = "0.000"; chk[key] = r; r += 1
    H.column_dimensions["A"].width = 52; H.column_dimensions["B"].width = 44
    for c in "CDEF": H.column_dimensions[c].width = 14
    H.column_dimensions["H"].width = 100
    return rows


# ---------- PL：標準損益表 ----------
YR = list(range(2023, 2032))
COLS = "CDEFGHIJK"


def dep_years(SIL):
    ys = [y for y in (2026, 2027, 2028, 2029) if SIL.RECO_FLAG.get(y) == 1]
    out = set(ys)
    if 2029 in ys:
        out |= {2030, 2031}   # 延伸期由 FY2029 路徑收斂而來
    return out


def pl_standard(wb, SIL):
    P = wb["PL"]; D = wb["DCF"]
    rb = lab(P, K["PLB"])
    g = lambda p: lab(P, p, rb, exact=True)
    KR = K["R"]
    R = dict(rev=g(KR["rev"]), gp=g(KR["gp"]), oi=g(KR["oi"]), ni=g(KR["ni"]), sh=g(KR["sh"]), eps=g(KR["eps"]))
    R["opx"] = g(KR["opx"]) if KR.get("opx") else None
    if K["nci_mode"] == "model":
        R["nci"] = lab(P, KR["nci"], rb); R["cons"] = g(KR["cons"])
    if K["ext_mode"] == "inblock":   # 緯穎：FY2030–FY2032 收斂法與情境 B 區塊同欄（H、I、J 欄），不另有延伸表
        X = {k: R[k] for k in ("rev", "ni", "sh", "eps")}
    elif K["ext_mode"] == "block":
        xb = lab(P, "延伸｜情境 B")
        X = {k: lab(P, p, xb, exact=True) for k, p in (("rev", "營收（延伸）"), ("ni", "淨利（延伸）"), ("sh", "股數（延伸）"), ("eps", "EPS（延伸）"))}
    else:   # 摘要區塊（情境 B）H、I 欄＝FY2030、FY2031；股數＝淨利 ÷ EPS
        sb = lab(P, "情境 B", exact=True); SR = K["SUM_R"]
        X = {"rev": lab(P, SR["rev"], sb, exact=True), "ni": lab(P, SR["ni"], sb, exact=True), "eps": lab(P, SR["eps"], sb, exact=True)}
    db = lab(D, "情境 B（基準）"); d_ebit = lab(D, "營業利益 EBIT", db)
    dep = dep_years(SIL)
    r = P.max_row + 3
    s0 = r
    P.cell(r, 1, f"標準損益表（v{VER} 起；個股 HTML 版面規格 v1.3 D 節）：歷史三年（FY2023–FY2025 實績，Hist 頁）＋FY2026E–FY2031E（情境 B 基準；FY2026 為錨點＝1–3Q26 財報實績＋4Q26 假設）。NT$B；EPS 新台幣。"
                 "只引用本頁上方、Hist、DCF 頁，不回饋任何計算").font = BOLD; r += 1
    P.cell(r, 1, K["GAAP_NOTE"])
    r_note = r; r += 1
    P.cell(r, 1, "FY2026E＝錨點；FY2027E–FY2029E＝情境 B 由下而上；FY2030E–FY2031E＝延伸期（收斂法只推營收與淨利，營業利益取 DCF 頁情境 B＝淨利 ÷ 淨利係數）。"
                 "延伸期毛利率與營業費用率沿用前一年，據以算出營業成本、毛利與營業費用；與延伸期營業利益的差額見表末核對列（目標價用延伸期淨利，不受影響）")
    r_note2 = r; r += 1
    hdr = ["項目", "單位"] + [f"FY{y}{'A' if y <= 2025 else 'E'}" + ("（對帳中）" if y in dep else "") for y in YR] + ["說明"]
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
    fc = lambda key: [f"=PL!{c}{R[key]}" for c in "DEFG"]            # FY2026E–FY2029E（PL 情境 B 區塊 D–G 欄）
    xe = lambda key: [f"=PL!H{X[key]}", f"=PL!I{X[key]}"]            # FY2030E–FY2031E（延伸表或摘要區塊 H、I 欄）
    h3 = lambda k: [f"=h_{k}_23", f"=h_{k}_24", f"=h_{k}_25"]
    row("rev", "營收", "NT$B", h3("rev") + fc("rev") + xe("rev"), M)
    row("yoy", "營收年增率（YoY）", "%", ["=C{rev}/h_rev_22-1".format(**L)] + [f"={c}{L['rev']}/{p}{L['rev']}-1" for p, c in zip(COLS, COLS[1:])], PCT)
    r_c = r
    row("cogs", "營業成本（COGS）", "NT$B", h3("cogs") + [f"=F{L['rev']}-F{r_c + 1}", f"=G{L['rev']}-G{r_c + 1}", f"=H{L['rev']}-H{r_c + 1}", f"=I{L['rev']}-I{r_c + 1}",
                                                          f"=J{L['rev']}*(1-I{r_c + 2})", f"=K{L['rev']}*(1-J{r_c + 2})"], M,
        "FY2030E、FY2031E＝營收 ×（1 − 前一年毛利率）")
    row("gp", "毛利", "NT$B", h3("gp") + fc("gp") + [f"=J{L['rev']}-J{L['cogs']}", f"=K{L['rev']}-K{L['cogs']}"], M)
    row("gm", "毛利率", "%", [f"={c}{L['gp']}/{c}{L['rev']}" for c in COLS], PCT, "FY2030E、FY2031E＝前一年（沿用）")
    r_o = r
    opx_fc = fc("opx") if R["opx"] else [f"=PL!{c}{R['gp']}-PL!{c}{R['oi']}" for c in "DEFG"]
    row("opx", "營業費用（銷管＋研發）", "NT$B", h3("opx") + opx_fc + [f"=J{L['rev']}*I{r_o}/I{L['rev']}", f"=K{L['rev']}*J{r_o}/J{L['rev']}"], M,
        "FY2030E、FY2031E＝營收 × 前一年營業費用率" + ("" if R["opx"] else "；預測期營業費用＝毛利 − 營業利益（模型以營業費用成長推導，與此式相同）"))
    row("oi", "營業利益", "NT$B", h3("oi") + fc("oi") + [f"=DCF!F{d_ebit}", f"=DCF!G{d_ebit}"], M, "FY2030E、FY2031E 取 DCF 頁情境 B 的營業利益（延伸期由淨利反推；不改）")
    row("opm", "營業利益率", "%", [f"={c}{L['oi']}/{c}{L['rev']}" for c in COLS], PCT)
    ro = "(h_pti_25-h_oi_25)/h_oi_25"
    row("oth", "其他收入（淨額：利息、匯兌、投資收益等）", "NT$B", [f"=h_pti_{y}-h_oi_{y}" for y in ("23", "24", "25")] + [f"=-2.801894+Q!E8*{ro}"] + [f"={c}{L['oi']}*{ro}" for c in "GHIJK"], M,
        "歷史＝稅前利益 − 營業利益；FY2026E＝1–9 月實績 −2.802（公開資訊觀測站）＋4Q 營業利益 × FY2025 實績比例；FY2027E 起＝營業利益 × FY2025 實績比例（展示用拆分）")
    row("pti", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    r_t = r
    nm = K["nci_mode"]
    if nm == "none":
        tax_f = [f"={c}{L['pti']}-{c}{r_t + 2}" for c in "FGHIJK"]
        nci_f = ["=0"] * 6
        tax_note, nci_note = "預測＝稅前利益 − 淨利（無少數股權；由模型淨利倒推；FY2026E 的 1–9 月部分等於實績 14.877）", "預測＝0（合併財報無少數股權損益）"
    elif nm == "model":
        # 緯創：FY2026E–FY2029E 取模型「合併淨利（含非控制權益）」與「非控制權益」列；FY2030E–FY2031E 無模型列，沿用 FY2029E 非控制權益 ÷ 稅前利益
        tax_f = [f"={c}{L['pti']}-PL!{s}{R['cons']}" for c, s in zip("FGHI", "DEFG")] + [f"={c}{L['pti']}-{c}{r_t + 1}-{c}{r_t + 2}" for c in "JK"]
        nci_f = [f"=PL!{s}{R['nci']}" for s in "DEFG"] + [f"={c}{L['pti']}*$I{r_t + 1}/$I{L['pti']}" for c in "JK"]
        tax_note = "FY2026E–FY2029E＝稅前利益 − 模型合併淨利（含非控制權益）；FY2030E–FY2031E＝稅前利益 − 少數股權損益 − 淨利"
        nci_note = "FY2026E–FY2029E＝模型非控制權益（（1 − 持股）× 緯穎淨利）；FY2030E–FY2031E＝稅前利益 × FY2029E 比例（模型無延伸期此列）"
    else:
        tax_f = [f"={c}{L['pti']}-{c}{r_t + 1}-{c}{r_t + 2}" for c in "FGHIJK"]
        nci_f = [f"={c}{L['pti']}*h_nci_25/h_pti_25" for c in "FGHIJK"]
        tax_note, nci_note = "預測＝稅前利益 − 少數股權損益 − 淨利（由模型淨利倒推）", "預測＝稅前利益 × FY2025 實績比例（展示用拆分）"
    row("tax", "所得稅", "NT$B", h3("tax") + tax_f, M, tax_note)
    row("nci", "少數股權損益", "NT$B", h3("nci") + nci_f, M, nci_note)
    row("ni", "稅後淨利（歸屬母公司）", "NT$B", h3("ni_rep") + fc("ni") + xe("ni"), M, "歷史＝資料來源列示值（Hist 頁）")
    row("nm", "淨利率", "%", [f"={c}{L['ni']}/{c}{L['rev']}" for c in COLS], PCT)
    sh_x = xe("sh") if K["ext_mode"] in ("block", "inblock") else [f"=H{L['ni']}/PL!H{X['eps']}", f"=I{L['ni']}/PL!I{X['eps']}"]
    row("sh", "股數（十億股）", "十億股", h3("sh_b") + fc("sh") + sh_x, "0.000", "歷史＝流通股數；預測＝模型股數（不含稀釋）" + ("" if K["ext_mode"] in ("block", "inblock") else "；FY2030E–FY2031E＝淨利 ÷ EPS"))
    row("eps", "EPS（基本；新台幣）", "NT$", h3("eps_b") + fc("eps") + xe("eps"), "0.00", "歷史＝資料來源列示基本 EPS；預測＝淨利 ÷ 模型股數")
    P.cell(r, 1, "比率").font = BOLD; r_ratio = r; r += 1
    row("r_yoy", "營收年增率", "%", [f"={c}{L['yoy']}" for c in COLS], PCT)
    row("r_gm", "毛利率", "%", [f"={c}{L['gm']}" for c in COLS], PCT)
    row("r_opm", "營業利益率", "%", [f"={c}{L['opm']}" for c in COLS], PCT)
    row("r_nm", "淨利率", "%", [f"={c}{L['nm']}" for c in COLS], PCT)
    row("r_niy", "淨利年增率（YoY）", "%", ["=C{ni}/h_ni_rep_22-1".format(**L)] + [f"={c}{L['ni']}/{p}{L['ni']}-1" for p, c in zip(COLS, COLS[1:])], PCT, "FY2023 的基期＝FY2022 淨利（Hist 頁）")
    row("etr", "有效稅率（所得稅 ÷ 稅前利益；預測為隱含值）", "%", [f"={c}{L['tax']}/{c}{L['pti']}" for c in COLS], PCT)
    P.cell(r, 1, "稀釋 EPS 對照（歷史；預測不估稀釋）").font = BOLD; r_dil = r; r += 1
    row("g_epsd", "EPS（稀釋；資料來源列示值）", "NT$", h3("eps_d"), "0.00", "稀釋加權股數見 Hist 頁")
    P.cell(r, 1, "核對（應為 0）").font = BOLD; r_chk = r; r += 1
    row("c1", "核對：營收 − 營業成本 − 毛利", "NT$B", [f"={c}{L['rev']}-{c}{L['cogs']}-{c}{L['gp']}" for c in COLS], "0.000")
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E；歷史有四捨五入差，≤ 0.001）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
    row("c2b", "差額（FY2030E、FY2031E）：毛利 − 營業費用（沿用前一年比率）− 延伸期營業利益（由淨利反推）；不為 0 表示兩種推法不一致，目標價用延伸期淨利", "NT$B",
        [None] * 7 + [f"=J{L['gp']}-J{L['opx']}-J{L['oi']}", f"=K{L['gp']}-K{L['opx']}-K{L['oi']}"], "0.000")
    row("c2c", "差額（FY2030E、FY2031E）占營收", "%", [None] * 7 + [f"=J{L['c2b']}/J{L['rev']}", f"=K{L['c2b']}/K{L['rev']}"], PCT)
    row("c3", "核對：稅前利益 − 所得稅 − 少數股權損益 − 稅後淨利（歷史有四捨五入差，≤ 0.001）", "NT$B", [f"={c}{L['pti']}-{c}{L['tax']}-{c}{L['nci']}-{c}{L['ni']}" for c in COLS], "0.000")
    row("c4", "核對：稅後淨利 − 情境 B 摘要（FY2026E–FY2029E）／延伸表（FY2030E–FY2031E）", "NT$B",
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
    # 外部稽核優化 W4-4（2026-11-23）：讀情境 B 欄（C 欄），與本頁「以情境 B 路徑反解」一致（同美超微 W2-2）。原讀 B 欄＝情境 A；
    # 目前 stub 期落在三情境相同的年度、兩欄值相同，本次只改參照、數值不變。
    nc_cell = f"Val!$C${lab(V, '淨現金滾到評價日')}"
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
    kv("px", "現價（新台幣）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（新台幣）", "=(px+dps_tgt)/(1+ke)^t_tgt", "0.00", "＝（現價＋期間股利）÷（1＋股權成本）^目標時點年數（DCF 頁每股價值推到 12 個月後的反算）")
    kv("nc", "淨現金滾到評價日（NT$B）", f"={nc_cell}", "#,##0.0", "Val 頁 0.1 節：淨現金與投資＋stub 期 FCFF − 已宣告股利")
    kv("sh", "評價用稀釋股數（十億股）", "=sh_v", "0.000")
    kv("evr", "支持現價所需的企業價值 EV（NT$B）", f"=B{r - 3}*B{r - 1}-B{r - 2}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B；NT$B）", f"=DCF!$C${ev}", "#,##0.0")
    kv("evx", "EV 需放大為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（新台幣）", f"=DCF!$C${ps}", "0.00")
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
    yr("rev", "營收（NT$B）", lambda dc, c: f"=DCF!{dc}{rv}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（NT$B）", lambda dc, c: f"=DCF!{dc}{eb}", "#,##0.0")
    yr("m", "營業利益率", lambda dc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("t", "稅率", lambda dc, c: f"=DCF!{dc}{tx}", "0.0%")
    yr("fr", "評價日後占該年度比例", lambda dc, c: f"=DCF!{dc}{fr}", "0.000")
    yr("df", "折現因子", lambda dc, c: f"=DCF!{dc}{df}", "0.0000")
    yr("tvf", "終值係數（只在最後一年：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) × 終值折現因子）", lambda dc, c: (f"=(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)/(1+wacc)^DCF!$C${tvy}" if dc == "H" else "=0"), "0.0000")
    yr("k", "每 1 元營業利益對 EV 的貢獻（FY2027 錨點不動，記 0）", lambda dc, c: ("=0" if dc == "C" else f"=(1-{c}{Y['t']})*({c}{Y['fr']}*{c}{Y['df']}+{c}{Y['tvf']})"), "0.0000")
    yr("cf", "NOPAT＋折舊 − 資本支出（NT$B）", lambda dc, c: f"=DCF!{dc}{nop}+DCF!{dc}{da}-DCF!{dc}{cx}", "#,##0.0")
    yr("wc", "營運資金增加（NT$B）", lambda dc, c: f"=DCF!{dc}{wc}", "#,##0.0")
    yr("pv", "現值（NT$B，DCF 頁）", lambda dc, c: f"=DCF!{dc}{pv}", "#,##0.0")
    r += 1
    RV.cell(r, 1, "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）").font = BOLD; r += 1
    kv("S", "營收 × 貢獻係數合計（NT$B）", f"=SUMPRODUCT(C{Y['rev']}:G{Y['rev']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("E0", "與營業利益無關的 EV（NT$B）", f"=B{A['ev']}-SUMPRODUCT(C{Y['ebit']}:G{Y['ebit']},C{Y['k']}:G{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "＝以同一係數加權的模型營業利益率")
    kv("m1", "現價隱含的營業利益率", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}<=0,"現價低於營業利益率為零時的價值：現價甚至沒有給預測期獲利定價",IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}>MAX(D{Y["m"]}:G{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    RV.cell(r, 1, "(b) 現價隱含的營收規模（FY2028–FY2032 營收等比例放大；營業利益率、稅率、折舊、資本支出、營運資金比率不變）").font = BOLD; r += 1
    kv("a", "不隨營收倍數變動的 EV（NT$B）", f"=B{Y['pv']}+fc_wc*B{Y['rev']}*C{Y['fr']}*C{Y['df']}", "#,##0.0", "FY2027 現值＋FY2028 營運資金增加中屬 FY2027 營收的部分")
    kv("b", "隨營收倍數等比例變動的 EV（NT$B）",
       f"=SUMPRODUCT(C{Y['cf']}:G{Y['cf']},C{Y['fr']}:G{Y['fr']},C{Y['df']}:G{Y['df']})-fc_wc*C{Y['rev']}*C{Y['fr']}*C{Y['df']}"
       f"-SUMPRODUCT(D{Y['wc']}:G{Y['wc']},D{Y['fr']}:G{Y['fr']},D{Y['df']}:G{Y['df']})+DCF!H{nop}*G{Y['tvf']}", "#,##0.0")
    kv("ab_chk", "核對：前兩列合計 − 模型 EV（應為 0）", f"=B{A['a']}+B{A['b']}-B{A['ev']}", "0.000000")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=(B{A['evr']}-B{A['a']})/B{A['b']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（NT$B）", f"=B{A['s1']}*C{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；NT$B）", f"=C{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/B{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=C{Y['rev']}/B{Y['rev']}-1", "0.0%")
    r += 1
    RV.cell(r, 1, "(c) 現價隱含的 WACC（各年自由現金流不變；永續期 RONIC＝WACC＋超額報酬差，同 DCF 頁；二分法 40 次）").font = BOLD; r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（新台幣）"], 1):
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
    kv("w_res", "核對：隱含 WACC 下的每股價值 − 支持現價所需值（新台幣，應接近 0）", f"=E{r - 4}-B{A['v0']}", "0.000000")
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
    RV.cell(r, 1, "約一年後前瞻 EPS（情境加權；新台幣）"); RV.cell(r, 3, f"={feps}").number_format = "0.00"; A["feps"] = r; r += 1
    cmp("pe1", "只看 Comparable：支持現價所需的前瞻本益比（倍）", f"=px/{feps}", f"=tp_comp_adj/{feps}", "0.0")
    cmp("rc1", "只看 Comparable：所需的相對倍數係數（÷ 同業前瞻本益比中位數）", f"=B{A['pe1']}/{pmed}", f"={coef}", "0.00")
    cmp("pe2", "加權（DCF 不變）：支持現價所需的 Comparable 前瞻本益比（倍）", f"=(px-tp_w_dcf*{dcf_tp})/((1-tp_w_dcf)*{feps})", f"=tp_comp_adj/{feps}", "0.0")
    RV.cell(r, 1, "同業前瞻本益比中位數（倍）"); RV.cell(r, 3, f"={pmed}").number_format = "0.0"; A["pmed"] = r; r += 1
    r += 1
    RV.cell(r, 1, "C. 摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    RV.cell(r, 1, F(T("現金流折現法：以情境 B 的現金流與 "), 'TEXT(wacc,"0.0%")', T(" 的折現率，每股價值 "), f'TEXT(B{A["v_b"]},"0.00")',
                     T(" 元；要支持現價 "), 'TEXT(px,"0.00")', T(" 元，企業價值須"), f'IF(B{A["evx"]}<1,"縮小為","放大為")', T("模型的 "), f'TEXT(B{A["evx"]},"0.00")',
                     T(" 倍。單靠一個因素達成時，需要：折現率"), f'IF(B{A["w1"]}>B{A["w0"]},"升到 ","降到 ")', f'TEXT(B{A["w1"]},"0.0%")', T("；或 FY2028–FY2032 營收全面"), f'IF(B{A["s1"]}<1,"縮小為 ","放大為 ")',
                     f'TEXT(B{A["s1"]},"0.00")', T(" 倍（FY2028 營收 "), f'TEXT(B{A["rv1"]},"#,##0")', T(" 十億元，模型為 "), f'TEXT(B{A["rv0"]},"#,##0")', T("）；或營業利益率"),
                     f'IF(B{A["m1"]}<B{A["m0"]},"降到 ","提高到 ")', f'TEXT(B{A["m1"]},"0.0%")', T("（模型等效約 "), f'TEXT(B{A["m0"]},"0.0%")', T("，終值年權重大）。"),
                     f'IF(B{A["m1"]}<=0,"現價低於預測期獲利為零時的價值。","")'))
    A["sum_dcf"] = r; r += 1
    RV.cell(r, 1, F(T("同業比較法：現價 "), 'TEXT(px,"0.00")', T(" 元相當於約一年後前瞻 EPS 的 "), f'TEXT(B{A["pe1"]},"0.0")',
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
    for j, h in enumerate(["項目", "目標價（NT$）", "相對現價", "說明"], 1):
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
    V.cell(r, 1, "現價（2026-11-20 收盤，除權後）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
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
    ra, rb_, rc_ = lab(P, "情境 A（樂觀）"), lab(P, "情境 B（基準）："), lab(P, "情境 C（保守）")
    eps = lambda r0, c: f"PL!{c}{lab(P, 'EPS（除權還原後）', r0, exact=True)}"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（v{VER} 起；個股 HTML 版面規格 v1.3：橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字；公式帶數字；原展示層各列保留）").font = BOLD
    r0 = r; r += 1
    M = {}

    def put(key, f):
        nonlocal r
        RD.cell(r, 1, f); M[key] = r; r += 1
    e27 = f"Recon!D{lab(RC, 'FY2027 EPS vs 共識')}"
    put("hero", F(f'{REC}', T("：加權目標價 "), f'TEXT({TPW},"0")', T(" 元，比現價 "), f'TEXT({PX},"0")', T(" 元高 "), f'TEXT({UP},"0%")',
                  T("；DCF "), f'TEXT({DCF},"0")', T(" 元與 Comparable（相對倍數調整後）"), f'TEXT({CMP},"0")', T(" 元各占一半。FY2027 盈餘比法人高 "), f'TEXT({e27},"0%")', T("，客戶 ASIC 放量是上檔的主因。")))
    put("rec", F(T("投資建議依加權目標價相對現價的上檔空間判定：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、低於 −50% 強力放空。目前上檔空間 "),
                 f'TEXT({UP},"+0.0%;-0.0%")', T("，投資建議 "), f'{REC}', T("。")))
    put("o1", F(T("1｜目標價 "), f'TEXT({TPW},"0")', T(" 元、上檔 "), f'TEXT({UP},"0%")', T("：DCF "), f'TEXT({DCF},"0")', T(" 元、Comparable（調整後）"), f'TEXT({CMP},"0")',
                T(" 元各占一半；現價只有本模型未來 12 個月 EPS 的 "), f'TEXT({NTM12},"0.0")', T(" 倍。只看市場主流同業法的 Comparable 為 "), f'TEXT({CMP0},"0")', T(" 元。"),
                T("剩餘可轉債（價內）若全數轉換（約 3,473 萬股、債務減 436 億元，Derived 近似）加權目標價約 "), f'TEXT(tp_w_dcf*(({DCF})*sh_v+43.579411*(1+ke))/(sh_v+0.034727)+(1-tp_w_dcf)*({CMP})*sh_v/(sh_v+0.034727),"0")',
                T(" 元（"), f'TEXT((tp_w_dcf*(({DCF})*sh_v+43.579411*(1+ke))/(sh_v+0.034727)+(1-tp_w_dcf)*({CMP})*sh_v/(sh_v+0.034727))/{TPW}-1,"+0.0%;-0.0%")', T("），評等不變。")))
    put("o2", F(T("2｜ASIC 傾斜是最大的單一假設：FY2027 EPS "), f'TEXT(Recon!C{lab(RC, "FY2027 EPS vs 共識")},"0")', T(" 元，比法人高 "), f'TEXT({e27},"0%")',
                T("；無 ASIC 傾斜時 FY2028 EPS "), f'TEXT({sv("無 ASIC 傾斜", "J")},"+0%;-0%")', T("。傾斜係數取上游「其他 ASIC」顆數（含非緯穎客戶），第三季財報後仍待驗證。")))
    put("o3", F(T("3｜三情境從 FY2028 才分開：FY2028 EPS 樂觀、基準、保守約 "), f'TEXT({eps(ra, "F")},"0")', T("、"), f'TEXT({eps(rb_, "F")},"0")', T("、"), f'TEXT({eps(rc_, "F")},"0")',
                T(" 元。份額相對值降到 0.75 時 FY2028 EPS "), f'TEXT({sv("份額流失（低）", "J")},"+0%;-0%")', T("；AI 部署 ×0.8 時 "), f'TEXT({sv("AI 部署 ×0.8", "J")},"+0%;-0%")', T("。")))
    put("o4", F(T("4｜DCF 為標準版（無槓桿自由現金流折現）：折現率 "), 'TEXT(wacc,"0.0%")', T("、永續成長率 "), 'TEXT(g_tv,"0.00%")', T("，終值現值占情境 B 企業價值 "),
                f'TEXT(DCF!C{lab(D, "終值現值占 EV", lab(D, "情境 B（基準）"))},"0%")', T("；折現率低／高一個百分點時加權目標價 "), f'TEXT({sv("WACC ±1pt", "D")},"+0%;-0%")', T("／"),
                f'TEXT({sv("WACC ±1pt", "F")},"+0%;-0%")', T("。")))
    if coord_rows:
        CR = coord_rows
        put("o5", F(T("5｜對帳中：Siliconomics 標示 CY"), 'Coord!B' + str(CR["yrs"]), T(" 對帳中；另一側讀數（Wattonomics 換算的部署量）目標價 "), f'TEXT(Coord!C{CR["tp"]},"0")', T(" 元（"), f'Coord!C{CR["rec"]}', T("），現行 "),
                    f'TEXT(Coord!B{CR["tp"]},"0")', T(" 元（"), f'Coord!B{CR["rec"]}', T("）；個股不裁決。公司實績（月營收）對層級：份額同源校準，無可比實績。")))
    else:
        put("o5", "5｜本次建置不含 Coord 頁（三層協調兩讀數快照重建用）。")
    pr = lambda p: lab(P, p, rb_)
    put("drv", F(T("分部營收怎麼推導（情境 B）：AI 伺服器營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 每 GW ODM 整櫃價值 × 緯穎有效份額（含 ASIC 傾斜係數）×（1 − 客供比例）；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("AI 伺服器營收")},"#,##0")', T(" 十億元。通用伺服器與其他（非 AI）由 FY2026 錨點乘成長率；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("通用伺服器與其他（非 AI）")},"#,##0")', T(" 十億元。兩者相加＝合併營收 "), f'TEXT(PL!E{PLA["R"]["rev"]},"#,##0")', T(" 十億元。")))
    PMED = vb("同業前瞻 P/E 中位數")
    put("ver", F(T("v0.2.9（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價，共識 EPS、rf、β、ERP、模型假設與情境權重不變）。現價（除權後）1,635 → "), 'TEXT(px,"#,##0")',
                 T("；同業前瞻 P/E 中位數 9.70 → "), f'TEXT({PMED},"0.00")', T("；DCF 目標價 3,247.53 → "), f'TEXT({DCF},"#,##0.00")', T("，加權目標價 3,145.45 → "), f'TEXT({TPW},"#,##0.00")',
                 T("（上檔 +92.4% → "), f'TEXT({UP},"+0.0%;-0.0%")', T("），評等 強力買進 → "), REC, T("。")))
    put("ver_v027", F(T("v0.2.8（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）：現價 1,650 → 1,635，同業前瞻 P/E 中位數 10.52 → 9.70，加權目標價 3,274.94 → 3,145.45（上檔 +98.5% → +92.4%），評等強力買進（不變）。v0.2.7（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（海外第一次可轉債 9-30 非流動餘額為 0，推定已轉換〔Derived〕；應付公司債流動部分改 9-30 推估值 45.80〔Derived：6-30 59.65 − 9-08 後第二次可轉債轉換約 1,103 萬股對應帳面約 13.85〕；已發放股利不再扣）；評價用股數改 9-30 股本 ÷ 面額 584.34 百萬股（9-02 除權每股配 1.98 股、其後可轉債轉換；與除權後現價同口徑）。"
                 "v0.2.6（2026-11-15）：版面規格 v1.3 全面改版（章節順序、橫幅 ≤200 字、重點觀察 ≤5 點、標準損益表 FY2023A–FY2031E、反向 DCF 頁、四個目標價與現價圖、敏感度改加權目標價相對基準低案／高案主圖）；"
                 "三層協調：Coord 頁並列兩讀數目標價，並新增 C0（實績）：2026 年 1–10 月月營收 × 公司揭露 AI 占比對層級隱含值，列兩讀數但不判定（層級份額由本模型校準，無可比實績）；同源列改標「讀取層級（同源，不判定）」。"
                 "xlsx 只新增 Hist、RevDCF、Coord 頁與各頁末新增區塊，既有數值、加權目標價與投資建議不變。"
                 "舊 README 內文（下方各節）與 v0.2.5 沿革節的數字早於 Siliconomics v0.19 全下游重建（例：FY2027 EPS 240.7、加權目標價 3,314 或 3,494.88），與現行不符；現行值以 Val 頁與本展示層為準（FY2027 EPS 231.02、加權目標價見上）。")))
    return M


# ---------- Coord：C0（實績）——公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.2.6 新增） ----------
TN = "https://finance.technews.tw/2026/{}/wiwynn-6669-2026{}-financial-report/"
MONTHS = [   # (月份, 營收 NT$B＝公告億元 ÷ 10, 來源網址, 說明)
    ("2026-01", 83.226, "https://finance.technews.tw/2026/07/08/wiwynn-6669-202606-financial-report/；https://technews.tw/2026/02/06/wiwynn-6669-202601-financial-report/",
     "Derived：公司公布前 6 月累計 5,546.60 億元 − 2–6 月各月 ＝ 832.26 億元；交叉核對：2025 年 12 月 1,042.91 億元 ×（1 − 20.2%，TechNews 標題月減 20.2%）≈ 832.2 億元"),
    ("2026-02", 94.633, "https://technews.tw/2026/03/06/wiwynn-6669-202602-financial-report/", "2 月營收 946.33 億元（TechNews 標題）"),
    ("2026-03", 98.649, TN.format("04/08", "03"), "3 月營收 986.49 億元（TechNews 標題）"),
    ("2026-04", 82.731, TN.format("05/07", "04"), "4 月營收 827.31 億元（TechNews 標題）"),
    ("2026-05", 84.050, TN.format("06/08", "05"), "5 月營收 840.50 億元（TechNews 標題；新聞稿：6 月較上月的 840.50 億元成長 32.5%）"),
    ("2026-06", 111.371, TN.format("07/08", "06"), "6 月營收 1,113.71 億元，較 5 月成長 32.5%，累計前 6 月 5,546.60 億元（TechNews 內文）"),
    ("2026-07", 117.686, TN.format("08/07", "07"), "7 月營收 1,176.86 億元（TechNews 標題）"),
    ("2026-08", 144.312, TN.format("09/08", "08"), "8 月營收 1,443.12 億元（TechNews 標題）"),
    ("2026-09", 172.620, TN.format("10/08", "09"), "9 月合併營收 1,726.20 億元，累計前 9 月 9,892.78 億元（TechNews 摘要）"),
    ("2026-10", 165.132, TN.format("11/06", "10"), "10 月營收 1,651.32 億元，較 9 月 −4.3%、年增 80.9%，累計前 10 月約 1.15 兆元（TechNews 內文，2026-11-06 公布）"),
]
def srcname(u):
    return "TechNews 月營收報導" + ("（內文已讀）" if any(k in u for k in ("202606", "202610")) else "（標題或摘要）")


VERIF = "Verified（公司月營收公告；經媒體轉載，MOPS 原文未直接讀取）"
URL_Q3 = "https://uanalyze.com.tw/articles/2502258599"
URL_AI = "https://uanalyze.com.tw/articles/4033848176"
URL_BRK = "https://finance.technews.tw/2026/11/09/6669-q326"


def c0_actuals(wb, C, r, l3, rs_l, d5):
    out = {"sec": r}
    C.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.2.6 新增）。結論：無可比實績（層級值同源）——層級份額由本模型校準，缺獨立的緯穎份額或 AI 伺服器營收金額揭露；下表只作參考，不判定、不計對帳中").font = BOLD; r += 1
    C.cell(r, 1, ("方法：公司實績側＝公司月營收（法定揭露）× 公司揭露的 AI 占營收比例（約五成，4Q25 法說：2025 年 AI 約占 5 成、2026 年微幅提高）＝AI 伺服器營收（緯穎未揭露金額，只能換算；比例等級 Interested-party、換算 Derived）。"
                  "層級側＝Siliconomics 整櫃組裝市場（基準層，CY2026）× 緯穎份額（份額表）× 匯率＝CY2026 全年隱含 AI 伺服器營收。"
                  "主讀數＝2026 年已公布月份（1–10 月）累計 vs 層級 CY2026 × 已公布月數 ÷ 12（線性時差近似，Derived）；第二讀數＝最近已公布一季（3Q26，7–9 月）× 4 vs 層級 CY2026。"
                  "成長年主讀數偏低；最近一季落在年度後半，第二讀數偏高（3Q 營收已遠高於全年平均速率）。"))
    r += 1
    C.cell(r, 1, ("為什麼不判定：層級份額（Siliconomics 份額表）是緯穎模型以自己的 FY2026 營收錨點 × AI 占營收（In 頁 48%）反推後存入的快照（Siliconomics 快照註記「公司模型校準值」），所以層級全年隱含值恆等於錨點營收 × 48%（見下方核對列），"
                  "整櫃組裝市場值與匯率在比式中消去，本表不檢驗 Siliconomics 的組裝市場，也不得讀成層級獲得獨立驗證。"
                  "本表實際檢驗的是三件事：(a) 錨點採用的 AI 占營收（48%）與公司揭露（約五成、2026 年微幅提高）的差；(b) 錨點的 4Q 營收假設（以 11–12 月隱含月均對 10 月實績衡量）；(c) 3Q26 公司公布的毛利率、營業利益率與淨利對錨點 3Q 假設。"
                  "3Q26 財報（2026-11-06 公告）已於 v0.2.7 滾入錨點（公開資訊觀測站），(c) 各列因此約為 0；記憶體代採購使認列營收低於毛額價值，會使實績側偏低（未調整）。"))
    r += 1
    for j, h in enumerate(["項目", "主讀數（1–10 月累計）", "第二讀數（3Q26 × 4）", "單位與口徑", "來源與證據等級", "出處網址（只在 xlsx）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    out["head"] = r; r += 1
    C.cell(r, 1, "4.1 公司實績：月營收（NT$B；Excel 儲存格單位十億新台幣）").font = BOLD; r += 1
    m0 = r
    for mo, v, src, note in MONTHS:
        C.cell(r, 1, f"　{mo} 月營收"); C[f"B{r}"] = v; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.000"
        C.cell(r, 4, "NT$B；合併營收（公司月營收公告）；" + note); C.cell(r, 5, ("Derived（由公司公布的累計與各月相減）" if mo == "2026-01" else VERIF) + "；出處：" + srcname(src)); C.cell(r, 6, src)
        r += 1
    m1 = r - 1
    out["m0"], out["m1"] = m0, m1
    C.cell(r, 1, "　2026-01～2026-10 累計（主讀數用）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m1})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布前 10 月累計約 1.15 兆元（TechNews 內文）"); C.cell(r, 5, "Verified（加總）"); cum = r; r += 1
    C.cell(r, 1, "　3Q26（7–9 月）合計（第二讀數用）"); C[f"B{r}"] = f"=SUM(B{m0+6}:B{m0+8})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布 3Q26 營收 4,346.18 億元（TechNews 摘要、優分析 3Q26 財報）"); C.cell(r, 5, "Verified（加總）"); q3 = r; r += 1
    C.cell(r, 1, "　核對：1–6 月合計 − 公司公布 1H26 營收（Q 頁）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+5})-Q!C4"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（公司以億元位數公布，容許 ±0.05）"); C.cell(r, 5, "核對列"); ck1 = r; r += 1
    C.cell(r, 1, "　核對：1–9 月合計 − 公司公布前 9 月累計 989.278（9,892.78 億元）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+8})-989.278"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（容許 ±0.05）"); C.cell(r, 5, "核對列（出處：TechNews 9 月營收報導、優分析 3Q26 財報）"); C.cell(r, 6, "https://finance.technews.tw/2026/10/08/wiwynn-6669-202609-financial-report/；" + URL_Q3); ck2 = r; r += 1
    C.cell(r, 1, "　核對：3Q26 合計 − 3Q26 營收輸入（In 頁 q3_rev）"); C[f"B{r}"] = f"=B{q3}-q3_rev"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（In 頁 q3_rev 取 434.618）"); C.cell(r, 5, "核對列"); ck3 = r; r += 1
    out["cum"], out["q3"] = cum, q3
    C.cell(r, 1, "4.2 口徑換算參數與 3Q26 財報實績（公司揭露；原簡報與 MOPS 未直接讀取）").font = BOLD; r += 1
    SRC_AI = "Interested-party（法說摘要：優分析 2026-02-26 4Q25 法說報導；公司：2025 年 AI 約占 5 成，2026 年微幅提高）"
    C.cell(r, 1, "　AI 伺服器占營收比例（主：公司揭露『約五成』）"); C[f"B{r}"] = 0.50; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司：2025 年 AI 產品約占營收 5 成、ASIC 約占 AI 九成；2026 年占比微幅提高（故 50% 為下限性質，實績側偏低）；Assumed 點值"); C.cell(r, 5, SRC_AI); C.cell(r, 6, URL_AI); ratio = r; r += 1
    C.cell(r, 1, "　參考：本模型 In 頁 AI 占營收（FY2026 錨點採用）"); C[f"B{r}"] = "=ais_26"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "本模型錨點採用的比例（In 頁 ais_26）；非公司揭露點值"); C.cell(r, 5, "Interested-party"); ratio_md = r; r += 1
    C.cell(r, 1, "　參考：In 頁敏感度高值（AI 占營收）"); C[f"B{r}"] = "=ais_26_hi"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司揭露『微幅提高』的情境上緣；Assumed"); C.cell(r, 5, "Assumed"); ratio_hi = r; r += 1
    C.cell(r, 1, "　3Q26 毛利率（公司公布，單季）"); C[f"B{r}"] = 0.0887; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.00%"
    C.cell(r, 4, "公司 3Q26 財報（2026-11-06 公告）：毛利率 8.87%"); C.cell(r, 5, "Verified（公司財報；經媒體轉載，MOPS 原文未直接讀取）"); C.cell(r, 6, URL_Q3); g3 = r; r += 1
    C.cell(r, 1, "　3Q26 營業利益率（公司公布，單季）"); C[f"B{r}"] = 0.0747; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.00%"
    C.cell(r, 4, "公司 3Q26 財報：營益率 7.47%"); C.cell(r, 5, "Verified（同上）"); C.cell(r, 6, URL_Q3); o3 = r; r += 1
    C.cell(r, 1, "　3Q26 稅後淨利（公司公布，單季）"); C[f"B{r}"] = 23.364; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司 3Q26 財報：稅後淨利 233.64 億元（利息費用上升壓抑業外）"); C.cell(r, 5, "Verified（同上）"); C.cell(r, 6, URL_Q3); n3 = r; r += 1
    C.cell(r, 1, "　3Q26 EPS（公司公布，單季；基本、除權還原後）"); C[f"B{r}"] = 41.46; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.00"
    C.cell(r, 4, "NT$；公司 3Q26 財報：EPS 41.46 元（去年同期 27.80 元，已依除權還原）"); C.cell(r, 5, "Verified（同上）"); C.cell(r, 6, URL_Q3); e3 = r; r += 1
    out["ratio"] = ratio
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
    a_rev = row2("　公司實績：營收（主：1–10 月累計；第二：3Q26 × 4）", f"=B{cum}", f"=B{q3}*4", "#,##0.0", "NT$B；合併營收", VERIF)
    a_rat = row2("　× AI 伺服器占營收比例", f"=B{ratio}", f"=B{ratio}", "0.0%", "見 4.2", "Derived／Interested-party")
    a_ai = row2("　＝ 公司實績側 AI 伺服器營收", f"=B{a_rev}*B{a_rat}", f"=C{a_rev}*C{a_rat}", "#,##0.0", "NT$B；主讀數為 10 個月累計；第二讀數為年化", "Derived（口徑換算）", True)
    l_mk = row2("　層級：整櫃組裝市場 CY2026（Siliconomics 基準層）", f"=B{l3}", f"=B{l3}", "#,##0.0", "US$B；連到下方「讀取層級」第 (3) 項的 Siliconomics 層列", "Siliconomics v0.22；Derived")
    l_sh = row2("　× 緯穎份額（Siliconomics 份額表）", f"=B{rs_l}", f"=B{rs_l}", "0.0%", "連到下方 (5) 的 Siliconomics 份額列（FY2026 尚未乘 ASIC 傾斜係數，傾斜＝1）", "Siliconomics v0.22；Derived（由本模型 FY2026 錨點校準，見提醒）")
    l_fx = row2("　× 匯率（新台幣 ÷ 美元；In 頁 fx）", "=fx", "=fx", "0.0", "NT$/US$", "Derived（本模型 In 頁 fx）")
    l_fy = row2("　＝ 層級隱含 CY2026 全年 AI 伺服器營收", f"=B{l_mk}*B{l_sh}*B{l_fx}", f"=C{l_mk}*C{l_sh}*C{l_fx}", "#,##0.0", "NT$B；全年", "Derived", True)
    l_tf = row2("　× 期間係數（主：已公布月數 ÷ 12；第二：1）", f"=COUNT(B{m0}:B{m1})/12", "=1", "0.000", "線性時差近似（層級值為日曆年）", "Derived")
    l_pd = row2("　＝ 層級側同期隱含 AI 伺服器營收", f"=B{l_fy}*B{l_tf}", f"=C{l_fy}*C{l_tf}", "#,##0.0", "NT$B", "Derived", True)
    out["fy"] = l_fy
    C.cell(r, 1, "　差距（實績側 ÷ 層級側 − 1）"); C[f"B{r}"] = f"=B{a_ai}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_ai}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"; C[f"{c}{r}"].font = BOLD
    C.cell(r, 4, "正值＝公司實績高於層級隱含"); out["d"] = r; d = r; r += 1
    C.cell(r, 1, "　判定：無可比實績（層級值同源）——份額同源校準，本表不檢驗組裝市場；參考（同源校準），不判定、不計對帳中")
    C[f"B{r}"] = "無可比實績（不判定）"; C[f"B{r}"].font = BOLD
    C.cell(r, 4, "缺獨立的緯穎份額或 AI 伺服器營收金額揭露（規則 v1.4.2：沒有可比實績時不以同源值補位）"); C.cell(r, 5, "規則 v1.4.2 C0 第 2 點、工作單 20261114"); out["judge"] = r; r += 1
    C.cell(r, 1, "　參考：兩讀數是否都在 ±15% 內（僅描述，不構成判定）"); C[f"B{r}"] = f'=IF(AND(ABS(B{d})<=0.15,ABS(C{d})<=0.15),"都在 ±15% 內","有讀數超過 ±15%")'
    C.cell(r, 4, "描述用；層級側恆等於錨點營收 × 48%（見下方核對列）；第二讀數落在年度後半，成長年偏高"); out["ref_in"] = r; r += 1
    C.cell(r, 1, "　核對：層級隱含 CY2026 全年 AI 伺服器營收 − 本模型 FY2026 錨點 AI 伺服器營收（Q 頁；應為 0）"); C[f"B{r}"] = f"=B{l_fy}-a26_ai"; C[f"B{r}"].number_format = "0.000000"
    C.cell(r, 4, "NT$B；為 0＝層級隱含值與本模型錨點同源（市場與匯率消去）"); C.cell(r, 5, "核對列"); out["ck_same"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：AI 占營收取本模型 In 頁比例（48%）時的差距"); C[f"B{r}"] = f"=B{a_rev}*B{ratio_md}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_rev}*B{ratio_md}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "本模型錨點採用的比例；非公司揭露。層級側＝錨點營收 × 48%，主讀數因此只剩月營收期間的差"); out["ref_md"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：AI 占營收取敏感度高值時的差距"); C[f"B{r}"] = f"=B{a_rev}*B{ratio_hi}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_rev}*B{ratio_hi}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "公司揭露『微幅提高』的上緣情境"); out["ref_hi"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (a)：錨點 AI 占營收（In 頁）÷ 公司揭露『約五成』 − 1"); C[f"B{r}"] = f"=B{ratio_md}/B{ratio}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "負值＝錨點採用的 AI 占營收低於公司揭露（公司：2026 年微幅提高，高於 2025 年的約 5 成）"); C.cell(r, 5, "Derived"); out["ref_a"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (b)：錨點隱含 11–12 月月均營收 ÷ 10 月實績 − 1"); C[f"B{r}"] = f"=(a26_rev-B{cum})/2/B{m1}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "NT$B 比；錨點＝1H26 實績＋3Q26 實績＋4Q 假設（4Q 營收 ＝ 3Q ×（1＋季增 8%））；負值＝錨點假設 11–12 月低於 10 月實績（參考，非實績：花旗預期 4Q 營收季增約 20%、大摩約 5,200 億元，見 TechNews 2026-11-09）")
    C.cell(r, 5, "Derived"); C.cell(r, 6, URL_BRK); out["ref_b"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (c1)：3Q26 毛利率實績 − 錨點 3Q 毛利率（百分點）"); C[f"B{r}"] = f"=B{g3}-Q!D6"; C[f"B{r}"].number_format = "+0.00%;-0.00%"
    C.cell(r, 4, "百分點；正值＝公司實績高於錨點假設（v0.2.7 起錨點 3Q＝財報實績，應約為 0；4Q 毛利率假設 8.3%）"); C.cell(r, 5, "Derived"); out["ref_c1"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (c2)：3Q26 營業利益率實績 − 錨點 3Q 營業利益率（百分點）"); C[f"B{r}"] = f"=B{o3}-Q!D9"; C[f"B{r}"].number_format = "+0.00%;-0.00%"
    C.cell(r, 4, "百分點；錨點 3Q 營益率＝毛利率 − 營業費用率假設"); C.cell(r, 5, "Derived"); out["ref_c2"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (c3)：3Q26 稅後淨利實績 ÷ 錨點 3Q 淨利 − 1"); C[f"B{r}"] = f"=B{n3}/Q!D10-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "NT$B 比；錨點 3Q 淨利＝營業利益 × 淨利係數（0.78）；實績含利息費用上升"); C.cell(r, 5, "Derived"); out["ref_c3"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (c4)：3Q26 EPS 實績 ÷ 錨點 3Q EPS（除權還原）− 1"); C[f"B{r}"] = f"=B{e3}/Q!D13-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "比；錨點 3Q EPS＝3Q 淨利 ÷ 除權後股數"); C.cell(r, 5, "Derived"); out["ref_c4"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：公司實績側營收 1–10 月累計 ÷ 本模型 FY2026 營收錨點（Q 頁）"); C[f"B{r}"] = f"=B{cum}/a26_rev"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "NT$B 比；錨點全年中已實現的比例"); C.cell(r, 5, "Derived"); out["ref_rev"] = r; r += 1
    out["ck"] = (ck1, ck2, ck3); out["end"] = r
    C.column_dimensions["A"].width = max(C.column_dimensions["A"].width or 0, 60)
    return out


# ---------- Coord：C0 由下往上對帳 ----------
def coord_c0(wb, SIL, R):
    C = wb["Coord"]; P = wb["PL"]
    C.column_dimensions["F"].width = 46; C.column_dimensions["G"].width = 70
    rb = lab(P, "情境 B（基準）：", exact=False); ra = lab(P, "情境 A（樂觀）：")
    gb = lambda p, ex=False: lab(P, p, rb, exact=ex); ga = lambda p, ex=False: lab(P, p, ra, exact=ex)
    N = SIL.NAMES
    r = C.max_row + 3
    C.cell(r, 1, "C0 由下往上對帳（三層基準協調規則 v1.4.2 C0 第 2 點；v0.2.6 新增）：比較基準是公司實績（第 4 節 C0〔實績〕：月營收與公司揭露 AI 占比 vs 層級隱含值）；本模型自層級讀入的同源值只列在第 2 節「讀取層級」，不判定、不計對帳中。差距超過 ±15% 標「對帳中」，不改上游值，由主控開對帳單。").font = BOLD
    c0 = r; r += 1
    C.cell(r, 1, "1. 讀哪一層、哪一版").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "版本／檔案", "本模型用途", "備註"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    flg = "／".join(str(SIL.RECO_FLAG[y]) for y in (2026, 2027, 2028, 2029)); alt = "／".join(f"{SIL.ALT_RATIO[y]:.2f}" for y in (2026, 2027, 2028, 2029))
    wt = "／".join(f"{N['SIF_WTAnchor_Base'][i]:.2f}" for i in (1, 2, 3))
    for a, b, c, d, e in (
        ("硬體股讀哪一層", "Siliconomics（零組件層；基準）", ver, "產業部署量、AI IT 資本支出、整櫃組裝市場、顆數係數、份額表、對帳中旗標（reader.py）", "下游只註明版本、不自行重估（專案規則第 3 條）"),
        ("AI IT 資本支出 SIF_AIITCapex_Base（情境 A 供給側；情境 B 以顆數係數縮放）", "US$B；見下表", "markets/network v0.2.3（Mkt!C27:F27）", "In 頁 itc_26–itc_29", "數值與 Siliconomics 基準相同（核對見下表）"),
        ("每 GW IT 資本支出 SIF_ITCapexPerGW、每 GW 機架數", "33.4／37.6／37.6／37.6 US$B", "markets/chip_bridge v0.2.3（Tokenomics v21.x 經橋接表）", "In 頁 itpg_gb、itpg_vr、racks_gb、racks_vr", "唯讀"),
        ("情境顆數係數（2028、2029）", f"A 1.0／1.0；B {SIL.COEF_B_TXT}；C 0.6／0.6", "tsmc v0.5.4（In!E37:E42；Siliconomics v0.22 基準係數）", "In 頁 um28_A–um29_C", "情境 B＝Siliconomics 基準"),
        ("客戶 ASIC 顆數與全部加速器顆數 → ASIC 傾斜係數 SIF_AsmWyTilt", "1.00／" + "／".join(f"{N['SIF_AsmWyTilt'][i]:.2f}" for i in (1, 2, 3)), "tsmc v0.5.4（In!E5:E36）；" + ver, "In 頁 ua_26–ua_29、ut_26–ut_29、asic_mix", "傾斜係數放大緯穎有效份額（Trainium 等 ASIC 成長快於全部加速器）；「其他 ASIC」含非緯穎客戶（已知限制）"),
        ("ODM 整櫃價格、每 GW 價值", "4.0／6.75 US$M／架；每 GW 價值 2028、2029 ÷ 2027＝1.0", "quanta v0.4.8（In!D24:F27；Siliconomics M18）", "In 頁 odm_gb、odm_vr、vpg_28、vpg_29", "唯讀（緯穎假設與廣達相同的每 GW 整櫃價值）"),
        ("份額快照 asm_wysh_26、asm_wyrel_27–29、asm_wyam、asm_wyai_26", f"{SIL.V['asm_wysh_26']['base']:.4f}／1.0／1.0／0.95；ASIC 比重 90%；AI 占營收 48%", ver, "In 頁 rel_27–rel_29、asic_mix、ais_26；FY2026 有效份額由公司揭露比例校準", "公司模型以自身營收校準，Siliconomics 只存快照（快照註記「公司模型校準值」）"),
        ("Wattonomics 錨定 SIF_WTAnchor_Base（wta_27–29）", wt, ver, "只用於 Sens「Wattonomics 錨定」案例，不進目標價", "Wattonomics 現行版本 v0.3.2（本模型不直接讀）"),
        ("對帳中旗標 SIF_RecoFlag_<yyyy>、SIF_DeployGW_AltRatio_<yyyy>", f"旗標 {flg}；比值 {alt}", ver, "C4 兩讀數（本頁上方）", "另一側讀數＝Wattonomics 換算的產業 GW")):
        for j, v in enumerate((a, b, c, d, e), 1):
            C.cell(r, j, v)
        r += 1
    r += 1
    C.cell(r, 1, "2. 讀取層級（同源，不判定；CY＝FY；情境 B 基準對 Siliconomics 基準；金額 US$B）：本模型自層級讀入的值，差距恆為 0，無檢驗力，故不進 C0 判定（規則 v1.4.2）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位與口徑", "來源與證據等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    flags = []
    YC = "BCDE"; PC = "DEFG"   # Coord 欄 B–E ↔ PL 欄 D–G（FY2026–FY2029）

    def block(title, model_f, layer_vals, unit, src, fmt="#,##0.0", layer_label="Siliconomics 層", tol=0.15, ref=False):
        nonlocal r
        C.cell(r, 1, title).font = BOLD; r += 1
        C.cell(r, 1, "　本模型"); C.cell(r, 6, unit); C.cell(r, 7, src)
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = model_f(PC[k]); C[f"{c}{r}"].number_format = fmt
        m = r; r += 1
        C.cell(r, 1, f"　{layer_label}")
        for k, c in enumerate(YC):
            C[f"{c}{r}"] = layer_vals[k]; C[f"{c}{r}"].number_format = fmt; C[f"{c}{r}"].font = BLUE
        l = r; r += 1
        C.cell(r, 1, "　差距（本模型 ÷ 層 − 1）")
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
    dg = lambda k: [N[k][i] for i in range(4)]
    block("(1) AI IT 資本支出（情境 B）", lambda c: f"=PL!{c}{gb('AI IT 資本支出 × 情境顆數係數')}", dg("SIF_AIITCapex_Base"), "US$B；加速器顆數 × 每顆 IT capex（含 ASIC）× 顆數係數", "Siliconomics 基準層（AI IT 資本支出）；Derived（markets/network v0.2.3 讀入）")
    block("(2) 部署 GW（IT 電力；情境 B）", lambda c: f"=PL!{c}{gb('部署 GW（IT 電力）')}", dg("SIF_DeployGW_Base"), "GW；AI IT 資本支出 ÷ 每 GW IT 資本支出", "Siliconomics 基準層（部署 GW）；Derived", "0.00")
    m3, l3, d3 = block("(3) 整櫃組裝市場（ODM 價；情境 B）", lambda c: f"=PL!{c}{gb('AI 伺服器組裝市場（ODM 價，美元）')}", dg("SIF_AsmMkt_Base"), "US$B；部署 GW × 每 GW ODM 整櫃價值（含 ASIC 機架，Assumed：每 GW 價值相同）",
                       "Siliconomics 基準層（整櫃組裝市場）；Derived／Assumed（每 GW 價值）")
    block("(4) 部署 GW（情境 A；供給側讀數）", lambda c: f"=PL!{c}{ga('部署 GW（IT 電力）')}", [N[f"SIF_DeployGW_Supply_{y}"] for y in (2026, 2027, 2028, 2029)], "GW；情境 A＝原始顆數（供給側）", "Siliconomics 供給側部署 GW 逐年值；Derived", "0.00")
    # 份額與未建模殘差
    C.cell(r, 1, "(5) 份額與恆等式（已建模 ODM 加總＋未建模殘差＝市場）").font = BOLD; r += 1
    C.cell(r, 1, "　緯穎有效份額（本模型，情境 B；含 ASIC 傾斜係數）"); C.cell(r, 6, "%；緯穎 AI 伺服器毛額價值 ÷ 組裝市場"); C.cell(r, 7, "PL 有效份額；FY2026 由公司揭露比例校準（Derived），2027–29＝份額相對值（Assumed）× ASIC 傾斜係數（Derived／Assumed）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = f"=PL!{PC[k]}{gb('緯穎有效份額', True)}"; C[f"{c}{r}"].number_format = "0.0%"
    rs_m = r; r += 1
    C.cell(r, 1, "　緯穎有效份額（Siliconomics 份額表：asm_wysh_26 × 份額相對值 × ASIC 傾斜係數；SIF_AsmShare_wy）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = N["SIF_AsmShare_wy"][k]; C[f"{c}{r}"].number_format = "0.0%"; C[f"{c}{r}"].font = BLUE
    rs_l = r; r += 1
    C.cell(r, 1, "　差距（百分點）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}-{c}{rs_l}"; C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    r += 1
    C.cell(r, 1, "　已建模 ODM 份額合計（廣達、鴻海、緯創、緯穎；SIF_AsmShareSum）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = N["SIF_AsmShareSum"][k]; C[f"{c}{r}"].number_format = "0.0%"; C[f"{c}{r}"].font = BLUE
    rs_s = r; C.cell(r, 7, "Siliconomics 已建模份額合計；Derived／Assumed"); r += 1
    C.cell(r, 1, "　未建模殘差（1 − 合計；Dell、Supermicro、Celestica 等；明列，不為 0）")
    for c in YC:
        C[f"{c}{r}"] = f"=1-{c}{rs_s}"; C[f"{c}{r}"].number_format = "0.0%"
    rs_r = r; r += 1
    C.cell(r, 1, "　恆等式核對：已建模合計 ＋ 殘差 − 100%（應為 0）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_s}+{c}{rs_r}-1"; C[f"{c}{r}"].number_format = "0.0000"
    r += 1
    C.cell(r, 1, "　緯穎占已建模合計")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}/{c}{rs_s}"; C[f"{c}{r}"].number_format = "0.0%"
    r += 1
    C.cell(r, 1, "　緯穎 AI 伺服器毛額價值（US$B；＝組裝市場 × 有效份額）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{m3}*{c}{rs_m}"; C[f"{c}{r}"].number_format = "#,##0.0"
    C.cell(r, 7, "本模型：PL 情境 B 組裝市場 × 有效份額；上限為市場（恆等式：加總不得超過上層）"); r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; r += 1
    C.cell(r, 1, "　下列差距是 Wattonomics 與 Siliconomics 兩層之間的差距，不是緯穎自己的量對其所讀 Siliconomics 值的差距；本公司依 C4 並列兩讀數，不裁決、不改上游。"); r += 1
    m5, l5, d5 = block("(6) 部署 GW（情境 B）對 Wattonomics 換算值（兩層之間的差距）", lambda c: f"=PL!{c}{gb('部署 GW（IT 電力）')}", [N[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)],
                       "GW；Siliconomics 的需求側部署 GW 逐年值＝Wattonomics 資本支出與電力路徑換算的需求側讀數", "Wattonomics v0.3.8 經 Siliconomics 換算；Derived（需求側）", "0.00", layer_label="Wattonomics 換算值（需求側）", ref=True)
    r += 1
    C.cell(r, 1, "3. 結論（C0；判定見第 4 節 C0〔實績〕表）").font = BOLD; r += 1
    C.cell(r, 1, "對帳中項目數（C0〔實績〕判定為「對帳中」者；讀取層級的同源列不判定、不計入）")
    n_other = r; r += 1
    C.cell(r, 1, "C0 結論（白話）")
    c0_concl = r; r += 2
    A4 = c0_actuals(wb, C, r, l3, rs_l, d5)
    C[f"B{n_other}"] = f'=IF(B{A4["judge"]}="對帳中",1,0)'
    C[f"B{c0_concl}"] = F(T("C0（實績）：無可比實績（層級值同源）——缺獨立的緯穎份額或 AI 伺服器營收金額揭露，不判定、不計對帳中（對帳中項目數 0）。"
                          "層級份額是本模型以自己的 FY2026 錨點反推後存入的快照，層級隱含值恆等於錨點營收 × 錨點 AI 占比，組裝市場值與匯率在比式中消去，不得讀成層級獲得獨立驗證。"
                          "第 4 節表實際檢驗的是錨點：採用的 AI 占營收比公司揭露（約五成）低 "), f'TEXT(-B{A4["ref_a"]},"0.0%")',
                          T("；錨點隱含 11–12 月月均營收比 10 月實績 "), f'TEXT(B{A4["ref_b"]},"+0.0%;-0.0%")',
                          T("；3Q26 毛利率實績比錨點高 "), f'TEXT(B{A4["ref_c1"]}*100,"0.00")', T(" 個百分點、淨利高 "), f'TEXT(B{A4["ref_c3"]},"0.0%")',
                          T("。（描述：以公司揭露比例換算的 1–10 月累計對層級同期隱含值 "), f'TEXT(B{A4["d"]},"+0.0%;-0.0%")', T("、最近一季年化對全年 "), f'TEXT(C{A4["d"]},"+0.0%;-0.0%")', T("，第二讀數落在年度後半而偏高。）"
                          "本模型自層級讀入的同源值（第 2 節）不判定。與 Wattonomics 換算的需求側讀數相比，部署 GW 在 CY2027–CY2029 差距 "),
                          f'TEXT(C{d5},"+0%")', T("、"), f'TEXT(D{d5},"+0%")', T("、"), f'TEXT(E{d5},"+0%")',
                          T("，超過 ±15%：這是兩層之間的差距（見「層級之間」小節），已由 Siliconomics 對帳單 r3 處理，不計入本公司 C0；本模型依 C4 並列兩讀數目標價（上方），不改上游值，不裁決哪一側正確。"))
    return dict(c0=c0, concl=c0_concl, flags=flags, d3=d3, d5=d5, m3=m3, l3=l3, **A4)


# ---------- 版本字串 ----------
def version(wb, here, coord_rows):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v0.1.2" in head, head
    a1 = head.replace("v0.1.2", f"v{VER}", 1) + a1[len(head):]
    a1 += (f"；v{VER}（{BUILT}）同版另：版面規格 v1.3 全面改版（新增 Hist、RevDCF 頁，PL、Val、Sens、README 頁末新增區塊）與三層協調 C0（實績）補公司實績對帳"
           "（Coord 頁末「C0〔實績〕」：1–10 月月營收 × 公司揭露 AI 占比 vs Siliconomics 整櫃組裝市場 × 緯穎份額，層級份額由本模型校準→無可比實績，不判定；同源列改標「讀取層級（同源，不判定）」）；既有數值、加權目標價與評等不變"
           "；v0.2.7（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 t164sb03／t164sb04）：FY2026 錨點改 1–3Q26 財報實績＋4Q26 假設；DCF 資產負債表日 2026-09-30；評價用股數 9-30 股本 ÷ 面額 584.34 百萬股（除權後）")
    head = a1.split("；", 1)[0]
    a1 = head.replace(f"v{VER}", "v0.2.9", 1) + a1[len(head):] + "；v0.2.8（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）；v0.2.9（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
    RD["A1"].value = a1
    return a1


def apply(wb, SIL, here, coord_rows):
    HR = hist_is(wb)
    PLA = pl_standard(wb, SIL)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_post(wb, here)
    C0 = coord_c0(wb, SIL, coord_rows) if coord_rows else None
    RMA = readme13(wb, SIL, PLA, RVA, coord_rows)
    version(wb, here, coord_rows)
    SO = wb["Sources"]; rs = SO.max_row + 1
    SO.cell(rs, 1, "[Verified] 緯穎 2026 年 1–10 月月營收（C0〔實績〕；TechNews 各月營收報導；1 月為 1–6 月累計減 2–6 月的推得值；MOPS 原文未直接讀取，網址見 Coord 頁第 4.1 節各列）")
    SO.cell(rs, 2, "https://finance.technews.tw/2026/11/06/wiwynn-6669-202610-financial-report/")
    SO.cell(rs + 1, 1, "[Verified] 緯穎 3Q26 財報（2026-11-06 公告：營收 4,346.18 億、毛利率 8.87%、營益率 7.47%、稅後淨利 233.64 億、EPS 41.46；優分析／鉅亨網轉載；v0.2.7 起以公開資訊觀測站 t164sb03／t164sb04 115Q3 為主要出處並滾入錨點，資產負債表日 2026-09-30）")
    SO.cell(rs + 1, 2, URL_Q3)
    SO.cell(rs + 2, 1, "[Interested-party] 法說摘要：4Q25 法說（2026-02-26）AI 產品約占營收 5 成、ASIC 約占 AI 九成、2026 年 AI 占比微幅提高（優分析）：C0〔實績〕AI 占比用，原簡報未直接讀取；3Q26 法說揭露的 AI 占比查無公開來源")
    SO.cell(rs + 2, 2, URL_AI)
    SO.cell(rs + 3, 1, "[Verified／第三方彙整] 緯穎損益表歷史 FY2022–FY2025（Hist 頁；S&P Global Market Intelligence 經 stockanalysis.com，2026-08-09 更新，除權還原後）")
    SO.cell(rs + 3, 2, "https://stockanalysis.com/quote/tpe/6669/financials/income-statement/")
    for _r in range(1, rs):   # 待補清單：3Q26 財報已取得（本版只比對、未滾入錨點）；法說 AI 占比仍待補
        _v = SO.cell(_r, 1).value
        if isinstance(_v, str) and _v.startswith("待補（人工）：3Q26 財報與法說"):
            SO.cell(_r, 1, _v.replace("3Q26 財報與法說（約 11 月上旬）", "3Q26 法說揭露的 AI 占比（財報已於 v0.2.7 滾入錨點）"))
    wb._v026 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, C0=C0, RM=RMA)
    return f"v026_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 主圖表 {ST}；README v1.3 {len(RMA)} 列；C0 {'有' if C0 else '無（無 Coord 頁）'}"
