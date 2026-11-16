# 鴻海 v0.4.6（2026-11-15）後處理：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10；伺服器組裝樣板＝廣達 v0.4.8）＋三層協調 C0（實績）與兩讀數（規則 v1.4.2）。
# 只新增，不改既有數值格（以 tools/xlsx_diff.py 對 v0.4.5 逐格比對）：
#   Hist 頁（新）：損益表歷史 FY2022–FY2025（S&P Global Market Intelligence，經 stockanalysis.com；含少數股權損益；FY2022 營收與淨利只供年增率）
#   PL 頁末：標準損益表（FY2023A–FY2025A 實績＋FY2026E–FY2031E 情境 B；GAAP 口徑）、比率區、核對列
#   RevDCF 頁（新）：反向 DCF——現價隱含的營業利益率、營收規模、WACC（DCF 情境 B 路徑）與 Comparable 倍數；只讀其他頁
#   Val 頁末：0.7 四個目標價並列（HTML 用）
#   Sens 頁末：目標價敏感度：業務主因（整本重算快照 foxconn/sens.json）與「加權目標價相對基準」白話表（HTML 敏感度主圖用）
#   README 頁末：v1.3 展示層白話列（橫幅說明、重點觀察 ≤5 點、分部營收推導說明、本版要點）
#   Coord 頁末（coord.add_sheet 之後）：C0 由下往上對帳（讀取層級、同源列不判定）＋第 4 節 C0（實績）：2026 年 1–10 月月營收與公司揭露的雲端網路占比 vs 層級隱含值
# build_foxconn.py 在 COORD.add_sheet 之後、存檔前呼叫 apply(wb, SIL, here, coord_rows)。
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

VER, BUILT = "0.4.6", "2026-11-15"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")

SP = ("S&P Global Market Intelligence 經 stockanalysis.com 損益表頁（https://stockanalysis.com/quote/tpe/2317/financials/income-statement/；2026-11-15 讀取，頁面更新 2026-06-30；"
      "單位百萬元新台幣，此處換算為十億元）；FY2025 營收、淨利、EPS 與公司揭露相符（去年稅後淨利 1,893 億元、EPS 13.61 元；優分析 https://uanalyze.com.tw/articles/4369349056）")
# (鍵, 項目, FY2023, FY2024, FY2025, FY2022（只供年增率）)；金額十億新台幣、EPS 新台幣、股數十億股
HIS = [
    ("rev", "營收", 6162.221, 6859.615, 8103.105, 6626.997),
    ("cogs", "營業成本", 5774.274, 6430.670, 7604.944, None),
    ("gp", "毛利", 387.947, 428.946, 498.161, None),
    ("opx", "營業費用（銷管＋研發）", 221.419, 228.339, 238.938, None),
    ("oi", "營業利益", 166.528, 200.607, 259.223, None),
    ("pti", "稅前利益", 192.224, 211.875, 293.445, None),
    ("tax", "所得稅", 37.435, 40.196, 78.410, None),
    ("nci", "少數股權損益（非控制權益；本期淨利中不屬於母公司的部分）", 12.691, 18.974, 25.681, None),
    ("ni_rep", "歸屬母公司淨利（資料來源列示值）", 142.098, 152.705, 189.354, 141.483),
    ("eps_b", "EPS（基本，新台幣；資料來源列示值）", 10.25, 11.01, 13.61, None),
    ("eps_d", "EPS（稀釋，新台幣；資料來源列示值）", 10.07, 10.79, 13.40, None),
    ("sh_b", "加權平均流通股數（基本；十億股）", 13.862, 13.874, 13.918, None),
    ("sh_d", "加權平均股數（稀釋；十億股）", 14.106, 14.146, 14.131, None),
]
BIZ = [  # (白話標籤, 低案設定, 高案設定, 鍵)；鍵＝FOXCONN_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄；x0.8＝基準 × 0.8；mf＝引擎 AI 部署係數）
    ("鴻海份額（相對值 0.85／0.75／0.70 ～ 1.15）", "rel_27=lo;rel_28=lo;rel_29=lo", "rel_27=hi;rel_28=hi;rel_29=hi", "份額"),
    ("FY2026 AI 占營收的錨點校準（雲端網路占比 51% × AI 占比 52% ～ 56% × 68%）", "cl_26=lo;ais_26=lo", "cl_26=hi;ais_26=hi", "AI 占營收"),
    ("客供比例（全為零 ～ 25%／40%／50%）", "cs_27=lo;cs_28=lo;cs_29=lo", "cs_27=hi;cs_28=hi;cs_29=hi", "客供比例"),
    ("客供部分保留毛利比例（70% ～ 100%）", "keep=lo", "keep=hi", "客供保留毛利"),
    ("每 GW 整櫃價值 2028–29（×0.85 ～ ×1.2）", "vpg_28=lo;vpg_29=lo", "vpg_28=hi;vpg_29=hi", "每 GW 價值"),
    ("AI 部署量（FY2027 起 AI IT 資本支出 ×0.8 ～ ×1.2）", "mf=x0.8", "mf=x1.2", "AI 部署量"),
    ("雲端網路非 AI 成長（0%／−5%／−5% ～ 25%／20%／15%）", "g_cn27=lo;g_cn28=lo;g_cn29=lo", "g_cn27=hi;g_cn28=hi;g_cn29=hi", "雲端非 AI 成長"),
    ("消費智能成長（iPhone 等；−8% ～ +8%／+6%）", "g_ce27=lo;g_ce28=lo;g_ce29=lo", "g_ce27=hi;g_ce28=hi;g_ce29=hi", "消費智能成長"),
    ("AI 毛利率差距（低 1.5 個百分點 ～ 高 0.5 個百分點）", "d_ai=lo", "d_ai=hi", "AI 毛利率差距"),
    ("VR 世代毛利率稀釋（−0.5 個百分點 ～ 不稀釋）", "vrd=lo", "vrd=hi", "VR 稀釋"),
    ("營業費用彈性（營業費用年增 ÷ 營收年增：0.2 ～ 0.7；基準 0.4）", "phi=lo", "phi=hi", "營業費用彈性"),
    ("淨利 ÷ 營業利益（0.60 ～ 0.73）", "nio_f=lo", "nio_f=hi", "淨利係數"),
]
K = dict(
    folder="foxconn", ENV="FOXCONN", cn="鴻海", HIS=HIS, SP=SP, SP22="；FY2022 取同表 FY2022 欄（營收 6,626,997、歸屬母公司淨利 141,483 百萬元）",
    PLB="情境 B：", R=dict(rev="營收", gp="毛利", opx="營業費用", oi="營業利益", ni="歸屬母公司淨利", sh="股數", eps="EPS"),
    nci_mode="ratio", ext_mode="block", BIZ=BIZ,
    GAAP_NOTE=("口徑：GAAP（IFRS 合併財報，歸屬母公司）。鴻海不另編 Non-GAAP，市場共識與公司法說皆用合併財報口徑，故本表用 GAAP；模型淨利為歸屬母公司淨利。"
               "歷史 EPS 為基本 EPS（資料來源列示值，與公司揭露一致；FY2025 13.61 元）；稀釋 EPS 在表末對照列。"
               "少數股權（非控制權益）金額大：FY2025 少數股權損益占稅前利益約 9%（工業富聯、鴻騰、富智康、樺漢等上市子公司的外部持股），表中「少數股權損益」列單獨列出，淨利為扣除後的歸屬母公司淨利。"
               "預測期模型只推營業利益與「淨利 ÷ 營業利益」一個係數（涵蓋業外、稅與少數股權），本表把它拆成：其他收入＝營業利益 × FY2025 實績比例、少數股權損益＝稅前利益 × FY2025 實績比例、"
               "所得稅＝稅前利益 − 少數股權損益 − 淨利（由模型淨利倒推，等於隱含有效稅率；因模型係數 0.66 低於 FY2025 的 0.73，隱含稅率高於 FY2025）。這三列是展示用拆分，不回饋計算、不改淨利與目標價。"),
    COMP_NOTE="同業前瞻 P/E 中位數 × 約一年後前瞻 EPS × 相對倍數調整係數（本公司十年相對同業的倍數比）",
    RV_NOTE="鴻海 FY2027 營運資金大增（營收年增約兩成，無槓桿自由現金流遠低於其他年度），所以 FY2027 視為錨點、不隨反解變動",
)


import glob
HERE = ''

def inval(wb, name):
    t = wb.defined_names[name].attr_text; sh, c = t.split("!"); return wb[sh.strip("'")][c.replace("$", "")].value

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
    crow("c1", "營收 − 營業成本 − 毛利（列示值；四捨五入差 ≤ 0.001）", "={rev}-{cogs}-{gp}")
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
    if K["ext_mode"] == "block":
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
    row("oth", "其他收入（淨額：利息、匯兌、投資收益等）", "NT$B", [f"=h_pti_{y}-h_oi_{y}" for y in ("23", "24", "25")] + [f"={c}{L['oi']}*{ro}" for c in "FGHIJK"], M,
        "歷史＝稅前利益 − 營業利益；預測＝營業利益 × FY2025 實績比例（展示用拆分）")
    row("pti", "稅前利益", "NT$B", [f"={c}{L['oi']}+{c}{L['oth']}" for c in COLS], M)
    r_t = r
    nm = K["nci_mode"]
    if nm == "none":
        tax_f = [f"={c}{L['pti']}-{c}{r_t + 2}" for c in "FGHIJK"]
        nci_f = ["=0"] * 6
        tax_note, nci_note = "預測＝稅前利益 − 淨利（無少數股權；由模型淨利倒推）", "預測＝0（合併財報無少數股權損益）"
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
    sh_x = xe("sh") if K["ext_mode"] == "block" else [f"=H{L['ni']}/PL!H{X['eps']}", f"=I{L['ni']}/PL!I{X['eps']}"]
    row("sh", "股數（十億股）", "十億股", h3("sh_b") + fc("sh") + sh_x, "0.000", "歷史＝流通股數；預測＝模型股數（不含稀釋）" + ("" if K["ext_mode"] == "block" else "；FY2030E–FY2031E＝淨利 ÷ EPS"))
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
    row("c2", "核對：毛利 − 營業費用 − 營業利益（FY2023A–FY2029E）", "NT$B", [f"={c}{L['gp']}-{c}{L['opx']}-{c}{L['oi']}" for c in COLS[:7]] + [None, None], "0.000")
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
    items = [("DCF 目標價（情境加權；進加權）", "=tp_dcf", "單純 DCF：三情境無槓桿自由現金流折現，價值推到 12 個月後，依情境權重加權"),
             ("Comparable 目標價（相對倍數調整後；進加權）", "=tp_comp_adj", K["COMP_NOTE"]),
             ("Comparable 目標價（市場主流同業法，未調整；不進加權）", "=tp_comp", "同業前瞻 P/E 中位數 × 約一年後前瞻 EPS，未做相對倍數調整"),
             ("加權目標價（橫幅主值）", "=tp_final", "DCF 與 Comparable（調整後）依 DCF 權重加權；投資建議依它的上檔空間判定"),
             ("市場共識目標價", "=cons_tp", f"=Val!E{cons}")]
    for k, (a, f, d) in enumerate(items):
        V.cell(r, 1, a); V.cell(r, 2, f).number_format = "#,##0.00"; V.cell(r, 3, f"=B{r}/px-1").number_format = "+0.0%;-0.0%"; V.cell(r, 4, d); r += 1
    V.cell(r, 1, "現價（2026-11-05 收盤）"); V.cell(r, 2, "=px").number_format = "#,##0.00"; V.cell(r, 4, "Yahoo Finance 日線收盤（In 頁 px）"); r += 1
    return g0


# ---------- Sens：業務主因快照與加權目標價相對基準 ----------
BIZ = K['BIZ']   # (白話標籤, 低案設定, 高案設定, 鍵)；設定＝{ENV}_OVR 覆寫的 In 輸入（lo／hi 取 In 頁敏感度低／高欄；x＝基準 × 倍數；mf＝引擎 AI 部署係數）
TOR_LABEL = {"WACC ±1pt": "折現率 WACC 低一個百分點／高一個百分點", "β 低／高": "β（風險係數）取低／高", "ERP 低／高": "股權風險溢酬 4%／6%",
             "永續成長率": "永續成長率低一個百分點／上限", "RONIC": "永續期新投資報酬率提高（只有高）", "資本支出 ÷ 營收": "資本支出占營收 −1／＋1 個百分點",
             "營運資金 ÷ 營收": "營運資金占營收 10%／20%", "淨現金不含權益投資": "淨現金不含權益投資（只有低）", "終值改用 FCFF": "終值改用 FCFF ×（1＋g）替代法（只有高）",
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
    tpc = "tp_final"
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
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; RC = wb["Recon"]; D = wb["DCF"]
    vb = lambda p, **k: f"Val!B{lab(V, p, **k)}"
    TPW, UP, REC = vb("目標價（DCF 與 Comparable 加權）"), vb("上檔空間：加權目標價"), vb("投資建議")
    DCF, CMP, CMP0, PX = vb("DCF 目標價（情境加權）"), vb("Comparable 目標價（相對倍數調整後）"), vb("Comparable 目標價（市場主流同業法）"), vb("現價（")
    NTM12 = f"Val!B{lab(V, '現價 ÷ 未來 12 月 EPS')}"
    sv = lambda p, c: f"Sens!{c}{lab(SE, p)}"
    r = RD.max_row + 2
    RD.cell(r, 1, f"v1.3 展示層（v{VER} 起；個股 HTML 版面規格 v1.3：橫幅 ≤200 字、重點觀察 ≤5 點 ≤750 字；公式帶數字；原展示層各列保留）").font = BOLD
    r0 = r; r += 1
    M = {}

    def put(key, f):
        nonlocal r
        RD.cell(r, 1, f); M[key] = r; r += 1
    put("hero", F(f'{REC}', T("：加權目標價 "), f'TEXT({TPW},"0")', T(" 元，比現價 "), f'TEXT({PX},"0")', T(" 元高 "), f'TEXT({UP},"0%")',
                  T("；DCF "), f'TEXT({DCF},"0")', T(" 元與 Comparable（相對倍數調整後）"), f'TEXT({CMP},"0")', T(" 元各占一半。上檔主要來自 DCF，並且假設 AI 伺服器份額與淨利係數守得住。")))
    put("rec", F(T("投資建議依加權目標價相對現價的上檔空間判定：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、低於 −50% 強力放空。目前上檔空間 "),
                 f'TEXT({UP},"+0.0%;-0.0%")', T("，投資建議 "), f'{REC}', T("。")))
    put("o1", F(T("1｜目標價 "), f'TEXT({TPW},"0")', T(" 元、上檔 "), f'TEXT({UP},"0%")', T("：DCF "), f'TEXT({DCF},"0")', T(" 元、Comparable（調整後）"), f'TEXT({CMP},"0")',
                T(" 元各占一半；現價只有本模型未來 12 個月 EPS 的 "), f'TEXT({NTM12},"0.0")', T(" 倍。只看市場主流同業法的 Comparable 為 "), f'TEXT({CMP0},"0")', T(" 元。")))
    put("o2", F(T("2｜盈餘比法人高：FY2027 EPS "), f'TEXT(Recon!C{lab(RC, "FY2027 EPS vs 法人")},"0.00")', T(" 元，比法人（08-18）高 "),
                f'TEXT(Recon!D{lab(RC, "FY2027 EPS vs 法人")},"0%")', T("；淨利 ÷ 營業利益用 "), 'TEXT(nio_f,"0.00")', T("（FY2025 為 "), 'TEXT(a25_ni/a25_oi,"0.00")', T("）；第三季財報（11-12）已公布，FY2026 錨點尚待更新。")))
    put("o3", F(T("3｜三情境從 FY2028 才分開：FY2028 EPS 樂觀、基準、保守約 "), f'TEXT(PL!F{lab(wb["PL"], "EPS", lab(wb["PL"], "情境 A："), exact=True)},"0.0")', T("、"),
                f'TEXT(PL!F{PLA["R"]["eps"]},"0.0")', T("、"), f'TEXT(PL!F{lab(wb["PL"], "EPS", lab(wb["PL"], "情境 C："), exact=True)},"0.0")', T(" 元。要追蹤份額與部署量：AI 部署量 ×0.8 時 FY2028 EPS 變動 "),
                f'TEXT({sv("AI 部署 ×0.8", "L")},"+0.0%;-0.0%")', T("，份額相對值降到 0.75 時 "), f'TEXT({sv("份額流失（低）", "L")},"+0.0%;-0.0%")', T("。")))
    put("o4", F(T("4｜DCF 為標準版（無槓桿自由現金流折現）：WACC "), 'TEXT(wacc,"0.0%")', T("、永續成長率 "), 'TEXT(g_tv,"0.00%")', T("，終值現值占情境 B 企業價值 "),
                f'TEXT(DCF!C{lab(D, "終值現值占 EV", lab(D, "情境 B（基準）"))},"0%")', T("；WACC 低／高一個百分點時加權目標價 "), f'TEXT({sv("WACC ±1pt", "D")},"+0%;-0%")', T("／"),
                f'TEXT({sv("WACC ±1pt", "F")},"+0%;-0%")', T("。")))
    if coord_rows:
        CR = coord_rows
        put("o5", F(T("5｜對帳中：Siliconomics 標示 CY"), 'Coord!B' + str(CR["yrs"]), T(" 對帳中；另一側讀數（Wattonomics 換算的部署量）目標價 "), f'TEXT(Coord!C{CR["tp"]},"0.00")', T(" 元（"), f'Coord!C{CR["rec"]}', T("），現行基準讀數 "),
                    f'TEXT(Coord!B{CR["tp"]},"0.00")', T(" 元（"), f'Coord!B{CR["rec"]}', T("）；個股不裁決，對帳在 Siliconomics 對帳單處理。")))
    else:
        put("o5", "5｜本次建置不含 Coord 頁（三層協調兩讀數快照重建用）。")
    rb = lab(wb["PL"], "情境 B：")
    pr = lambda p: lab(wb["PL"], p, rb)
    put("drv", F(T("分部營收怎麼推導（情境 B）：AI 伺服器營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 每 GW ODM 整櫃價值 × 鴻海有效份額 ×（1 − 客供比例）；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("AI 伺服器營收")},"#,##0")', T(" 十億元。非 AI 營收＝雲端網路非 AI＋消費智能＋電腦終端＋元件及其他，由 FY2026 錨點乘各自成長率；"
                   "FY2027 為 "), f'TEXT(PL!E{pr("非 AI 營收合計")},"#,##0")', T(" 十億元。兩者相加＝合併營收 "), f'TEXT(PL!E{PLA["R"]["rev"]},"#,##0")', T(" 十億元；四大產品類別加總與合併營收的檢查列為 0。")))
    put("ver", F(T("v0.4.7（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 t164sb03／t164sb04，Verified）。FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 季增、毛利率、營業費用率、淨利 ÷ 營業利益不改）：營收 11,150.7 → 11,129.0 十億元、歸屬母公司淨利 275.26 → 275.54、EPS 19.64 → 19.63（股數改 1–3Q 加權 14.039 十億股）。"
                 "資產負債表日 2026-06-30 → 09-30：借款 1,315.9 → 1,505.9（一年內到期長期負債 64.3 為 6-30 代理，9-30 未取得）、租賃 41.8 → 44.7、現金及流動金融資產 1,524.7 → 1,449.9、長期投資 491.2 → 477.3、非控制權益帳面 211.7 → 223.1、評價用股數 14.017 → 14.074；"
                 "FY2025 股利已於 07-31 發放，stub 期不再扣；淨現金與投資（扣非控制權益）−165.5 → −458.4，滾到評價日 −329.3 → −489.7。DCF 486.28 → 472.48、加權目標價 411.81 → 404.09（強力買進不變）；另一側讀數 371.32 → 363.69（強力買進）。"
                 "C0：Siliconomics 鴻海份額快照仍以 v0.4.6 錨點反推，同源核對列差 6.9 十億元（−0.19% 營收錨點變動），待 Siliconomics 下次升版重存。")))
    return M


# ---------- Coord：C0（實績）——公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.4.6 新增） ----------
TN = "https://finance.technews.tw/2026/{}/hon-hai-2317-2026{}-financial-report/"
MONTHS = [   # (月份, 營收 NT$B, 來源網址, 說明)；公司月營收公告（億元，TechNews 標題列示至小數兩位；9 月、10 月為 TechNews 月報導）
    ("2026-01", 730.039, TN.format("02/05", "01"), "1 月合併營收 7,300.39 億元，年增 35.5%"),
    ("2026-02", 595.813, TN.format("03/05", "02"), "2 月營收 5,958.13 億元，累計前 2 月 1.33 兆元"),
    ("2026-03", 803.738, TN.format("04/05", "03"), "3 月合併營收 8,037.38 億元，累計前 3 月 2.13 兆元"),
    ("2026-04", 832.098, TN.format("05/05", "04"), "4 月營收 8,320.98 億元"),
    ("2026-05", 859.409, TN.format("06/05", "05"), "5 月合併營收 8,594.09 億元"),
    ("2026-06", 821.763, TN.format("07/05", "06"), "6 月營收 8,217.63 億元，累計前 6 月 4.64 兆元"),
    ("2026-07", 946.513, TN.format("08/05", "07"), "7 月營收 9,465.13 億元"),
    ("2026-08", 921.766, TN.format("09/05", "08"), "8 月營收 9,217.66 億元"),
    ("2026-09", 1158.633, "https://finance.technews.tw/2026/10/05/2026-sep-foxconn/", "9 月合併營收 11,586.33 億元（1 兆 1,586.33 億元）；3Q26 3 兆 269 億元，累計前 9 月 7 兆 6,698 億元"),
    ("2026-10", 1305.3, "https://finance.technews.tw/2026/11/05/oct-2026-foxconn/", "10 月營收 1 兆 3,053 億元（報導只列到億元位數），月增 12.66%、年增 45.73%；累計前 10 月 8 兆 9,751 億元"),
]
VERIF = "Verified（公司月營收公告；經 TechNews 報導轉載，MOPS 原文未直接讀取）"
URL_MIX = {"1Q26": "https://blog.fugle.tw/post/earnings-call-2317-2026-05-14", "2Q26": "https://blog.fugle.tw/post/earnings-call-2317-2026-08-12",
           "3Q26": "https://news.nextapple.com/finance/20261112/D566D07E5C054836DE510FB38F6781C9；https://www.bnext.com.tw/article/93040/2026-foxconn--earnings-call-；https://uanalyze.com.tw/articles/3008659196"}
SRC_MIX = "Interested-party（法說摘要：富果／壹蘋／數位時代／優分析；原簡報未直接讀取，公司揭露的整數百分比）"


def c0_actuals(wb, C, r, l3, rs_l, d5):
    out = {"sec": r}
    C.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；v0.4.6 新增）。結論：無可比實績（層級值同源）——層級的鴻海份額由本模型 FY2026 錨點校準，且公司未揭露 AI 伺服器營收金額或其占比（只揭露雲端網路產品占營收）；下表只作參考，不判定、不計對帳中").font = BOLD; r += 1
    C.cell(r, 1, ("方法：公司實績側＝公司月營收（法定揭露）× 公司法說揭露的雲端網路產品占營收比例（1Q26 48%、2Q26 51%、3Q26 54%；整數百分比）＝雲端網路產品營收（Derived）；"
                  "再 × AI 占雲端網路比例（本模型 In 頁，Assumed；公司只說「超過五成」下限，未揭露實際比例）＝AI 伺服器營收（含假設，不是公司實績）。"
                  "層級側＝Siliconomics 整櫃組裝市場（基準層，CY2026）× 鴻海份額（份額表）× 匯率＝CY2026 全年隱含 AI 伺服器營收。"
                  "主讀數＝2026 年已公布產品占比期間（1Q–3Q，9 個月）累計 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived；成長年偏低）；"
                  "第二讀數＝最近已公布一季（3Q26）× 4 vs 層級 CY2026（約等於 3Q 季中點即 8 月中的年化速率；季中點晚於年中，成長下偏高）。10 月營收已公布但產品占比尚未公布，只用於錨點的 4Q 假設檢驗。"))
    r += 1
    C.cell(r, 1, ("為什麼不判定：層級份額（Siliconomics 份額表）是本模型以自己的 FY2026 錨點（營收 × 雲端網路占比 × AI 占雲端網路 ÷ 組裝市場）反推後存入的快照，所以層級全年隱含值恆等於錨點營收 × 雲端網路占比 × AI 占雲端網路，"
                  "整櫃組裝市場值與匯率在比式中消去，本表不檢驗 Siliconomics 的組裝市場，也不得讀成層級獲得獨立驗證；AI 占雲端網路是假設，公司實績側也含同一個假設，兩邊相同的部分消去。"
                  "本表實際檢驗的是兩件事：(a) 錨點採用的雲端網路占營收（In 頁 53%）對公司 1Q–3Q 揭露的加權占比，以及推得的 4Q 占比是否合理；(b) 錨點的 4Q 營收假設（以 11–12 月隱含月均對 10 月實績衡量）。"
                  "2H26 起部分高價料件改客供（公司：1Q26 起已有半成品模式）會使認列營收低於毛額價值（未調整）。"))
    r += 1
    for j, h in enumerate(["項目", "主讀數（1Q–3Q 累計）", "第二讀數（3Q26 × 4）", "單位與口徑", "來源與證據等級", "出處網址（只在 xlsx）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    out["head"] = r; r += 1
    C.cell(r, 1, "4.1 公司實績：月營收（NT$B；Excel 儲存格單位十億新台幣）").font = BOLD; r += 1
    m0 = r
    for mo, v, src, note in MONTHS:
        C.cell(r, 1, f"　{mo} 月營收"); C[f"B{r}"] = v; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.000"
        C.cell(r, 4, "NT$B；合併營收（公司月營收公告）；" + note); C.cell(r, 5, VERIF + "；出處：TechNews 月營收報導"); C.cell(r, 6, src)
        r += 1
    m1 = r - 1
    out["m0"], out["m1"] = m0, m1
    C.cell(r, 1, "　2026-01～2026-10 累計"); C[f"B{r}"] = f"=SUM(B{m0}:B{m1})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布前 10 月累計 8 兆 9,751 億元"); C.cell(r, 5, "Verified（加總）"); cum10 = r; r += 1
    C.cell(r, 1, "　1Q26（1–3 月）合計"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+2})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布累計前 3 月 2.13 兆元；財報營收（法說：2 兆 1,195 億元）與月營收合計差約 0.5%（合併沖銷口徑）"); C.cell(r, 5, "Verified（加總）"); q1 = r; r += 1
    C.cell(r, 1, "　2Q26（4–6 月）合計"); C[f"B{r}"] = f"=SUM(B{m0+3}:B{m0+5})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B"); C.cell(r, 5, "Verified（加總）"); q2 = r; r += 1
    C.cell(r, 1, "　3Q26（7–9 月）合計"); C[f"B{r}"] = f"=SUM(B{m0+6}:B{m0+8})"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布 3Q26 營收 3 兆 269 億元；法說財報營收 3.0156 兆元，與月營收合計差約 0.4%"); C.cell(r, 5, "Verified（加總）"); q3 = r; r += 1
    C.cell(r, 1, "　1Q–3Q26（1–9 月）合計（主讀數用）"); C[f"B{r}"] = f"=B{q1}+B{q2}+B{q3}"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；公司公布前 9 月累計 7 兆 6,698 億元"); C.cell(r, 5, "Verified（加總）"); cum9 = r; r += 1
    C.cell(r, 1, "　核對：1–6 月合計 − 公司公布 1H26 營收（Q 頁）"); C[f"B{r}"] = f"=SUM(B{m0}:B{m0+5})-Q!C4"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；v0.4.7 起 Q 頁 1H26 為財報（MOPS 1–9 月累計 − 3Q26），本列＝月營收自結數與財報營收之差（合併沖銷口徑），不再應為 0"); C.cell(r, 5, "核對列（參考）"); ck1 = r; r += 1
    C.cell(r, 1, "　核對：1–9 月合計 − 公司公布前 9 月累計 7,669.8（7 兆 6,698 億元）"); C[f"B{r}"] = f"=B{cum9}-7669.8"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（累計以億元位數列示，容許 ±0.05）"); C.cell(r, 5, "核對列（出處：TechNews 2026-10-05 月營收報導）"); C.cell(r, 6, "https://finance.technews.tw/2026/10/05/2026-sep-foxconn/"); ck2 = r; r += 1
    C.cell(r, 1, "　核對：3Q26 合計 − 3Q26 營收輸入（In 頁 q3_rev）"); C[f"B{r}"] = f"=B{q3}-q3_rev"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；v0.4.7 起 In 頁 q3_rev 為財報單季 3,015.634（MOPS），本列＝月營收自結數與財報營收之差（約 +0.4%），不再應為 0"); C.cell(r, 5, "核對列（參考）"); ck3 = r; r += 1
    C.cell(r, 1, "　核對：1–10 月合計 − 公司公布前 10 月累計 8,975.1（8 兆 9,751 億元）"); C[f"B{r}"] = f"=B{cum10}-8975.1"; C[f"B{r}"].number_format = "0.000"
    C.cell(r, 4, "NT$B；應約為 0（10 月只到億元位數，容許 ±0.5）"); C.cell(r, 5, "核對列（出處：TechNews 2026-11-05 月營收報導）"); C.cell(r, 6, "https://finance.technews.tw/2026/11/05/oct-2026-foxconn/"); ck4 = r; r += 1
    out["cum10"], out["cum9"], out["q3"] = cum10, cum9, q3
    C.cell(r, 1, "4.2 口徑換算參數（公司法說揭露的產品占營收比例；原簡報未直接讀取）").font = BOLD; r += 1
    cl = {}
    for k, v, qq in (("1Q26", 0.48, q1), ("2Q26", 0.51, q2), ("3Q26", 0.54, q3)):
        C.cell(r, 1, f"　雲端網路產品占營收：{k} 法說揭露"); C[f"B{r}"] = v; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
        C.cell(r, 4, {"1Q26": "公司：雲端網路產品接近五成（48%）", "2Q26": "公司：雲端網路首度超過五成（51%）；消費智能 29%、電腦終端 15%、元件及其他 5%",
                      "3Q26": "公司：雲端網路產品受惠 AI 伺服器需求，占比進一步提升至 54%；消費智能 29%、電腦終端 12%、元件及其他 5%"}[k])
        C.cell(r, 5, SRC_MIX); C.cell(r, 6, URL_MIX[k]); cl[k] = r; r += 1
    for k, qq in (("1Q26", q1), ("2Q26", q2), ("3Q26", q3)):
        C.cell(r, 1, f"　＝ {k} 雲端網路產品營收（月營收合計 × 占比）"); C[f"B{r}"] = f"=B{qq}*B{cl[k]}"; C[f"B{r}"].number_format = "#,##0.000"
        C.cell(r, 4, "NT$B；Derived（占比為整數百分比，±0.5 個百分點的四捨五入誤差未計）"); C.cell(r, 5, "Derived／Interested-party"); cl[k + "r"] = r; r += 1
    C.cell(r, 1, "　1Q–3Q26 雲端網路產品營收合計（主讀數用）"); C[f"B{r}"] = f"=B{cl['1Q26r']}+B{cl['2Q26r']}+B{cl['3Q26r']}"; C[f"B{r}"].number_format = "#,##0.000"
    C.cell(r, 4, "NT$B；Derived"); C.cell(r, 5, "Derived／Interested-party"); cloud9 = r; r += 1
    C.cell(r, 1, "　1Q–3Q26 雲端網路產品占營收（加權）"); C[f"B{r}"] = f"=B{cloud9}/B{cum9}"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "Derived；各季占比以月營收合計加權"); C.cell(r, 5, "Derived／Interested-party"); w9 = r; r += 1
    C.cell(r, 1, "　AI 伺服器占雲端網路產品比例（本模型 In 頁 ais_26）"); C[f"B{r}"] = "=ais_26"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "公司未揭露實際比例；本模型 FY2026 取 60%（Assumed／Interested-party）。同一個值也在層級份額的校準中，故兩邊相同"); C.cell(r, 5, "Assumed／Interested-party（見 In 頁 ais_26）"); ais = r; r += 1
    C.cell(r, 1, "　參考：公司揭露的 AI 占雲端網路下限"); C[f"B{r}"] = 0.5; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "2Q25 法說：AI 伺服器占雲端網路超過五成（公司；下限，不是實際值）；1Q26 法說：AI 伺服器占伺服器營收已突破五成（分母為伺服器，口徑不同）"); C.cell(r, 5, "Interested-party（法說摘要：優分析 2Q25、富果 1Q26）")
    C.cell(r, 6, "https://blog.fugle.tw/post/earnings-call-2317-2026-05-14"); ais_lo = r; r += 1
    C.cell(r, 1, "　參考：本模型 In 頁 FY2026 雲端網路產品占營收（錨點採用）"); C[f"B{r}"] = "=cl_26"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "本模型錨點採用的全年占比（1Q26 48%、2Q26 51% 的外推；Interested-party）"); C.cell(r, 5, "Interested-party"); cl26 = r; r += 1
    C.cell(r, 1, "　參考：3Q26 財報營收（MOPS；In 頁 q3_rev）"); C[f"B{r}"] = "=q3_rev"; C[f"B{r}"].number_format = "#,##0.0"
    C.cell(r, 4, "NT$B；MOPS t164sb04 115Q3 營業收入 3,015,633,915 千元（法說 3.0156 兆元一致）；與月營收合計的差距見下列"); C.cell(r, 5, "Verified（MOPS）"); C.cell(r, 6, "https://mops.twse.com.tw/mops/#/web/t164sb04"); rev3 = r; r += 1
    C.cell(r, 1, "　參考：3Q26 財報營收 ÷ 月營收合計 − 1"); C[f"B{r}"] = f"=B{rev3}/B{q3}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "月營收（自結）與財報營收的口徑差；C0 兩讀數的實績側仍用月營收（與 1Q、2Q 一致、可延伸到 10 月），錨點 v0.4.7 起用財報"); C.cell(r, 5, "Derived"); r += 1
    C.cell(r, 1, "　參考：3Q26 營業利益（公司揭露）"); C[f"B{r}"] = 115.350449; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "#,##0.00"
    C.cell(r, 4, "NT$B；法說：3Q26 營業利益 1,153.5 億元（毛利率 6.02%、營業利益率 3.83%；首度超過千億元）"); C.cell(r, 5, "Interested-party（法說摘要）"); C.cell(r, 6, URL_MIX["3Q26"]); oi3 = r; r += 1
    C.cell(r, 1, "　參考：3Q26 EPS（公司揭露）"); C[f"B{r}"] = 5.51; C[f"B{r}"].font = BLUE; C[f"B{r}"].number_format = "0.00"
    C.cell(r, 4, "新台幣（基本）；法說：3Q26 歸屬母公司淨利 775.24 億元、EPS 5.51 元（單季新高）；前三季 EPS 13.35 元"); C.cell(r, 5, "Interested-party（法說摘要）"); C.cell(r, 6, URL_MIX["3Q26"]); eps3 = r; r += 1
    out["cl"], out["w9"], out["ais"] = cl, w9, ais
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
    a_cl = row2("　公司實績：雲端網路產品營收（主：1Q–3Q 累計；第二：3Q26 × 4）", f"=B{cloud9}", f"=B{cl['3Q26r']}*4", "#,##0.0", "NT$B；合併營收 × 公司揭露占比", "Derived／Interested-party")
    a_ai = row2("　× AI 占雲端網路產品比例", f"=B{ais}", f"=B{ais}", "0.0%", "見 4.2；假設值，兩邊相同", "Assumed／Interested-party")
    a_ns = row2("　＝ 公司側換算的 AI 伺服器營收", f"=B{a_cl}*B{a_ai}", f"=C{a_cl}*C{a_ai}", "#,##0.0", "NT$B；主讀數為 9 個月累計；第二讀數為年化（含假設，非公司實績）", "Assumed 級（換算含假設）", True)
    l_mk = row2("　層級：整櫃組裝市場 CY2026（Siliconomics 基準層）", f"=B{l3}", f"=B{l3}", "#,##0.0", "US$B；連到上方「讀取層級」第 (3) 項的 Siliconomics 層列", "Siliconomics v0.20；Derived")
    l_sh = row2("　× 鴻海份額（Siliconomics 份額表）", f"=B{rs_l}", f"=B{rs_l}", "0.0%", "連到上方 (5) 的 Siliconomics 份額列", "Siliconomics v0.20；Derived（由本模型 FY2026 錨點校準，見說明）")
    l_fx = row2("　× 匯率（新台幣 ÷ 美元；In 頁 fx）", "=fx", "=fx", "0.0", "NT$/US$", "Derived（與廣達、台達電模型相同）")
    l_fy = row2("　＝ 層級隱含 CY2026 全年 AI 伺服器營收", f"=B{l_mk}*B{l_sh}*B{l_fx}", f"=C{l_mk}*C{l_sh}*C{l_fx}", "#,##0.0", "NT$B；全年", "Derived", True)
    row2("　核對（同源證據）：層級隱含 CY2026 全年 − 本模型 FY2026 AI 伺服器營收錨點（Q 頁）", f"=B{l_fy}-a26_ai", f"=C{l_fy}-a26_ai", "0.000", "NT$B；應為 0：層級值與本模型錨點同源，所以本表不構成獨立驗證", "Derived（核對列）")
    l_tf = row2("　× 期間係數（主：9 ÷ 12；第二：1）", f"=COUNT(B{m0}:B{m0+8})/12", "=1", "0.000", "線性時差近似（層級值為日曆年）", "Derived")
    l_pd = row2("　＝ 層級側同期隱含 AI 伺服器營收", f"=B{l_fy}*B{l_tf}", f"=C{l_fy}*C{l_tf}", "#,##0.0", "NT$B", "Derived", True)
    out["fy"] = l_fy
    C.cell(r, 1, "　差距（公司側 ÷ 層級側 − 1）"); C[f"B{r}"] = f"=B{a_ns}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_ns}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"; C[f"{c}{r}"].font = BOLD
    C.cell(r, 4, "正值＝公司側換算高於層級隱含；兩讀數方向相反，來自時差（主讀數偏低、第二讀數偏高），不是獨立證據"); out["d"] = r; d = r; r += 1
    C.cell(r, 1, "　判定：無可比實績（層級值同源）——份額由本模型錨點校準，公司未揭露 AI 伺服器金額；不判定、不計對帳中")
    C[f"B{r}"] = "無可比實績（不判定）"; C[f"B{r}"].font = BOLD
    C.cell(r, 4, "缺獨立揭露：AI 伺服器營收金額、AI 占雲端網路的實際比例（規則 v1.4.2：沒有可比實績時不以同源值補位）"); C.cell(r, 5, "規則 v1.4.2 C0 第 2 點、工作單 20261114"); out["judge"] = r; r += 1
    C.cell(r, 1, "　參考：兩讀數是否都在 ±15% 內（僅描述，不構成判定）"); C[f"B{r}"] = f'=IF(AND(ABS(B{d})<=0.15,ABS(C{d})<=0.15),"都在 ±15% 內","有讀數超過 ±15%")'
    C.cell(r, 4, "描述用；層級側恆等於錨點營收 × 雲端網路占比 × AI 占雲端網路（見上方說明），兩讀數的差距只反映 1Q–3Q 實際占比與全年錨點占比的差、以及時差"); out["ref_in"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：AI 占雲端網路取公司揭露下限 50% 時的差距"); C[f"B{r}"] = f"=B{a_cl}*B{ais_lo}/B{l_pd}-1"; C[f"C{r}"] = f"=C{a_cl}*B{ais_lo}/C{l_pd}-1"
    for c in "BC": C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "下限不是實際值；只顯示若實際比例在 50% 附近，錨點的 60% 會高估多少（與 In 頁敏感度低值 52% 相近）"); out["ref_lo"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (a1)：錨點 FY2026 雲端網路占營收 − 公司 1Q–3Q 揭露加權占比（百分點）"); C[f"B{r}"] = f"=B{cl26}-B{w9}"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "正值＝全年錨點高於前三季實際加權；要追上需 4Q 占比更高（下列）"); C.cell(r, 5, "Derived"); out["ref_a1"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (a2)：錨點隱含的 4Q 雲端網路占營收"); C[f"B{r}"] = f"=(B{cl26}*a26_rev-B{cloud9})/(a26_rev-B{cum9})"; C[f"B{r}"].number_format = "0.0%"
    C.cell(r, 4, "＝（錨點全年雲端網路營收 − 前三季實際）÷ 錨點 4Q 營收；對照 3Q26 實際占比與公司「4Q AI 機櫃出貨季增高雙位數」、10 月雲端網路略低於 9 月（客供模式、供料、新機種放量）"); C.cell(r, 5, "Derived"); out["ref_a2"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (a3)：錨點隱含 4Q 占比 − 3Q26 實際占比（百分點）"); C[f"B{r}"] = f"=B{r-1}-B{cl['3Q26']}"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "正值＝錨點要求 4Q 占比再上升"); C.cell(r, 5, "Derived"); out["ref_a3"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (b)：錨點隱含 11–12 月月均營收 ÷ 10 月實績 − 1"); C[f"B{r}"] = f"=(a26_rev-B{cum10})/2/B{m1}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "NT$B 比；錨點＝1H26 實績＋3Q26 實績＋4Q 假設（季增）；負值＝錨點假設 11–12 月低於 10 月實績（10 月為歷史新高單月）"); C.cell(r, 5, "Derived"); out["ref_b"] = r; r += 1
    C.cell(r, 1, "　實際檢驗 (b2)：若 11、12 月維持 10 月水準，全年營收 ÷ 錨點 − 1"); C[f"B{r}"] = f"=(B{cum10}+2*B{m1})/a26_rev-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "描述用，不是預測；正值＝錨點的 4Q 營收低於 10 月年化水準"); C.cell(r, 5, "Derived"); out["ref_b2"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：3Q26 營業利益（公司揭露）÷ Q 頁 3Q26 營業利益（模型）− 1"); C[f"B{r}"] = f"=B{oi3}/Q!D8-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "NT$B；v0.4.7 起 Q 頁 3Q26 為財報實績（MOPS），本列應為 0"); C.cell(r, 5, "Verified（MOPS）；Derived"); out["ref_oi"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：3Q26 EPS（公司揭露）÷ Q 頁 3Q26 EPS（模型）− 1"); C[f"B{r}"] = f"=B{eps3}/Q!D13-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "新台幣；v0.4.7 起 Q 頁 3Q26 EPS＝財報淨利 ÷ 1–3Q 加權股數（q_sh），與公司揭露 5.51 的差只來自股數口徑"); C.cell(r, 5, "Verified（MOPS）；Derived"); out["ref_eps"] = r; r += 1
    C.cell(r, 1, "　參考（不判定）：錨點 4Q26 EPS（模型）÷ 3Q26 實際 EPS − 1"); C[f"B{r}"] = f"=Q!E13/B{eps3}-1"; C[f"B{r}"].number_format = "+0.0%;-0.0%"
    C.cell(r, 4, "＝錨點要求 4Q 盈餘比 3Q 實績再成長多少（v0.4.7 起錨點＝1–3Q26 財報＋4Q 假設）"); C.cell(r, 5, "Derived"); out["ref_q4"] = r; r += 1
    out["ck"] = (ck1, ck2, ck3, ck4); out["end"] = r
    C.column_dimensions["A"].width = max(C.column_dimensions["A"].width or 0, 60)
    return out

# ---------- Coord：C0 由下往上對帳 ----------
def coord_c0(wb, SIL, R):
    C = wb["Coord"]; P = wb["PL"]
    C.column_dimensions["F"].width = 46; C.column_dimensions["G"].width = 70
    rb = lab(P, "情境 B：", exact=False); ra = lab(P, "情境 A：")
    gb = lambda p, ex=False: lab(P, p, rb, exact=ex); ga = lambda p, ex=False: lab(P, p, ra, exact=ex)
    N = SIL.NAMES
    r = C.max_row + 3
    C.cell(r, 1, "C0 由下往上對帳（三層基準協調規則 v1.4.2 C0 第 2 點；v0.4.6 新增）：比較基準是公司實績（第 4 節 C0〔實績〕：月營收與公司揭露的雲端網路占比 vs 層級隱含值）；本模型自層級讀入的同源值只列在第 2 節「讀取層級」，不判定、不計對帳中。差距超過 ±15% 標「對帳中」，不改上游值，由主控開對帳單。").font = BOLD
    c0 = r; r += 1
    C.cell(r, 1, "1. 讀哪一層、哪一版").font = BOLD; r += 1
    for j, h in enumerate(["項目", "值", "版本／檔案", "本模型用途", "備註"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ver = f"Siliconomics {SIL.VERSION}（{SIL.FILE}）"
    ROOT = os.path.dirname(HERE)
    fn = lambda d: os.path.basename(sorted(glob.glob(os.path.join(ROOT, d, "dist", "*.xlsx")))[-1])
    inv = lambda nm: inval(wb, nm)
    f2 = lambda xs, p=2: "／".join(f"{x:.{p}f}" for x in xs)
    um = lambda k: f"{inv('um28_' + k):.2f}／{inv('um29_' + k):.2f}"
    fxsh = SIL.V["asm_fxsh_26"]["base"]; fxrel = [SIL.V[f"asm_fxrel_{y}"]["base"] for y in (27, 28, 29)]
    for a, b, c, d, e in (
        ("硬體股讀哪一層", "Siliconomics（零組件層；基準）", ver, "產業部署量、AI IT 資本支出、整櫃組裝市場、顆數係數、份額表、對帳中旗標（reader.py）", "下游只註明版本、不自行重估（專案規則第 3 條）"),
        ("AI IT 資本支出 SIF_AIITCapex_Base（情境 A 供給側；情境 B 以顆數係數縮放）", "US$B；見下表", f"markets/network（{fn('markets/network')}）", "In 頁 itc_26–itc_29", "數值與 Siliconomics 基準相同（核對見下表）"),
        ("每 GW IT 資本支出 SIF_ITCapexPerGW、每 GW 機架數、整櫃買方價", f2(SIL.NAMES["SIF_ITCapexPerGW"], 1) + " US$B", f"markets/chip_bridge（{fn('markets/chip_bridge')}）", "In 頁 itpg_gb、itpg_vr、racks_gb、racks_vr、rack_gb、rack_vr", "唯讀（In 頁 itpg_gb 為 GB300 值，其餘年度用 VR200 值）"),
        ("情境顆數係數（2028、2029）", f"A {um('A')}；B {um('B')}；C {um('C')}", f"tsmc（{fn('tsmc')}；In!E37:E42）", "In 頁 um28_A–um29_C", "個股自訂情境，讀 tsmc 成品；產業量與 Siliconomics 基準的關係見下表 (1)–(3)"),
        ("份額快照 asm_fxsh_26、asm_fxrel_27–29", f"{fxsh:.4f}／" + f2(fxrel), ver, "In 頁 rel_27–rel_29；FY2026 有效份額由公司揭露比例校準", "公司模型以自身營收校準，Siliconomics 只存快照（asm_fxsh_26 註記為公司模型校準值）"),
        ("Wattonomics 錨定 SIF_WTAnchor_Base（wta_27–29）", f2(SIL.NAMES["SIF_WTAnchor_Base"][1:4]), ver, "只用於 Sens「Wattonomics 錨定」案例，不進目標價", "Wattonomics 現行版本 v0.3.2（本模型不直接讀）"),
        ("對帳中旗標 SIF_RecoFlag_<yyyy>、SIF_DeployGW_AltRatio_<yyyy>", "旗標 " + "／".join(str(SIL.RECO_FLAG[y]) for y in (2026, 2027, 2028, 2029)) + "；比值 " + f2([SIL.NAMES[f"SIF_DeployGW_AltRatio_{y}"] for y in (2026, 2027, 2028, 2029)]), ver, "C4 兩讀數（本頁上方）", "另一側讀數＝Wattonomics 換算的產業 GW")):
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
    qs = SIL.V["asm_fxsh_26"]["base"]; qrel = [1.0] + [SIL.V[f"asm_fxrel_{y}"]["base"] for y in (27, 28, 29)]
    C.cell(r, 1, "　鴻海有效份額（本模型，情境 B）"); C.cell(r, 6, "%；鴻海 AI 伺服器毛額價值 ÷ 組裝市場"); C.cell(r, 7, "PL 有效份額；FY2026 由公司揭露比例校準（Derived），2027–29 為份額相對值（Assumed）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = f"=PL!{PC[k]}{gb('鴻海有效份額', True)}"; C[f"{c}{r}"].number_format = "0.0%"
    rs_m = r; r += 1
    C.cell(r, 1, "　鴻海有效份額（Siliconomics 份額表：FY2026 基準份額 × 份額相對值）")
    for k, c in enumerate(YC):
        C[f"{c}{r}"] = qs * qrel[k]; C[f"{c}{r}"].number_format = "0.0%"; C[f"{c}{r}"].font = BLUE
    rs_l = r; r += 1
    C.cell(r, 1, "　差距（百分點）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}-{c}{rs_l}"; C[f"{c}{r}"].number_format = "+0.0%;-0.0%"
    r += 1
    C.cell(r, 1, "　已建模 ODM 份額合計（廣達、鴻海、緯創、緯穎）")
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
    C.cell(r, 1, "　鴻海占已建模合計")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{rs_m}/{c}{rs_s}"; C[f"{c}{r}"].number_format = "0.0%"
    r += 1
    C.cell(r, 1, "　鴻海 AI 伺服器毛額價值（US$B；＝組裝市場 × 有效份額）")
    for c in YC:
        C[f"{c}{r}"] = f"={c}{m3}*{c}{rs_m}"; C[f"{c}{r}"].number_format = "#,##0.0"
    C.cell(r, 7, "本模型：PL 情境 B 組裝市場 × 有效份額；上限為市場（恆等式：加總不得超過上層）"); r += 2
    C.cell(r, 1, "層級之間（參考；Siliconomics 對帳單 r3 處理，不計入本公司 C0）").font = BOLD; r += 1
    C.cell(r, 1, "　下列差距是 Wattonomics 與 Siliconomics 兩層之間的差距，不是鴻海自己的量對其所讀 Siliconomics 值的差距；本公司依 C4 並列兩讀數，不裁決、不改上游。"); r += 1
    m5, l5, d5 = block("(6) 部署 GW（情境 B）對 Wattonomics 換算值（兩層之間的差距）", lambda c: f"=PL!{c}{gb('部署 GW（IT 電力）')}", [N[f"SIF_DeployGW_Alt_{y}"] for y in (2026, 2027, 2028, 2029)],
                       "GW；Siliconomics 的需求側部署 GW 逐年值＝Wattonomics 資本支出與電力路徑換算的需求側讀數", "Wattonomics v0.3.6 經 Siliconomics 換算；Derived（需求側）", "0.00", layer_label="Wattonomics 換算值（需求側）", ref=True)
    r += 1
    C.cell(r, 1, "3. 結論（C0；判定見第 4 節 C0〔實績〕表）").font = BOLD; r += 1
    C.cell(r, 1, "對帳中項目數（C0〔實績〕判定為「對帳中」者；讀取層級的同源列不判定、不計入）")
    n_other = r; r += 1
    C.cell(r, 1, "C0 結論（白話）")
    c0_concl = r; r += 2
    A4 = c0_actuals(wb, C, r, l3, rs_l, d5)
    C[f"B{n_other}"] = f'=IF(B{A4["judge"]}="對帳中",1,0)'
    C[f"B{c0_concl}"] = F(T("C0（實績）：無可比實績（層級值同源）——Siliconomics 的鴻海份額是本模型以自己的 FY2026 錨點（營收 × 雲端網路占比 × AI 占雲端網路 ÷ 組裝市場）反推後存入的快照，層級隱含值恆等於錨點營收 × 雲端網路占比 × AI 占雲端網路，組裝市場值與匯率在比式中消去；"
                          "公司未揭露 AI 伺服器營收金額，也未揭露 AI 占雲端網路的實際比例（只說超過五成），所以沒有獨立實績可判定，不判定、不計對帳中（對帳中項目數 0），不得讀成層級獲得獨立驗證。"
                          "第 4 節表實際檢驗的是錨點：(a) 錨點全年雲端網路占營收比公司 1Q–3Q 揭露的加權占比 "), f'TEXT(B{A4["ref_a1"]},"+0.0%;-0.0%")',
                          T("（百分點），4Q 需達 "), f'TEXT(B{A4["ref_a2"]},"0.0%")', T("，比 3Q26 實際 "), f'TEXT(B{A4["cl"]["3Q26"]},"0%")', T(" 高 "), f'TEXT(B{A4["ref_a3"]},"0.0%")',
                          T("（百分點）；(b) 錨點隱含 11–12 月月均營收比 10 月實績 "), f'TEXT(B{A4["ref_b"]},"+0.0%;-0.0%")',
                          T("，11、12 月若維持 10 月水準，全年營收比錨點高 "), f'TEXT(B{A4["ref_b2"]},"0.0%")',
                          T("（描述：以公司揭露占比與假設的 AI 占雲端網路換算的 1Q–3Q 累計對層級同期隱含值 "), f'TEXT(B{A4["d"]},"+0.0%;-0.0%")', T("、3Q 年化對全年 "), f'TEXT(C{A4["d"]},"+0.0%;-0.0%")', T("，兩讀數反向來自時差與假設，不是獨立證據）。"
                          "本模型自層級讀入的同源值（第 2 節）不判定。與 Wattonomics 換算的需求側讀數相比，部署 GW 在 CY2027–CY2029 差距 "),
                          f'TEXT(C{d5},"+0%")', T("、"), f'TEXT(D{d5},"+0%")', T("、"), f'TEXT(E{d5},"+0%")',
                          T("，超過 ±15%：這是兩層之間的差距（見「層級之間」小節），已由 Siliconomics 對帳單 r3 處理，不計入本公司 C0；本模型依 C4 並列兩讀數目標價（上方），不改上游值，不裁決哪一側正確。"))
    return dict(c0=c0, concl=c0_concl, flags=flags, d3=d3, d5=d5, m3=m3, l3=l3, **A4)


# ---------- 版本字串 ----------
def version(wb, here, coord_rows):
    RD = wb["README"]
    a1 = RD["A1"].value
    a1 += (f"；v{VER}（{BUILT}）：版面規格 v1.3 全面改版（新增 Hist、RevDCF 頁與各頁末新增區塊）；三層協調 C0：Coord 頁新增由下往上對帳，讀取層級的同源列改標不判定，第 4 節 C0〔實績〕以 2026 年 1–10 月月營收與公司揭露的雲端網路占比對層級隱含值，"
           "結論無可比實績（層級值同源）並實際檢驗錨點的雲端網路占比與 4Q 營收假設；既有數值、加權目標價與評等不變"
           + "；v0.4.7（2026-11-15）：滾入 3Q26 財報（MOPS）——FY2026 錨點改 1–3Q26 財報實績＋4Q26 假設、資產負債表日改 2026-09-30（stub 不再扣已發放股利）")
    RD["A1"].value = a1
    return a1


def apply(wb, SIL, here, coord_rows):
    global HERE
    HERE = here
    HR = hist_is(wb)
    PLA = pl_standard(wb, SIL)
    RVA = revdcf(wb)
    VG = val_g(wb)
    ST = sens_post(wb, here)
    C0 = coord_c0(wb, SIL, coord_rows) if coord_rows else None
    RMA = readme13(wb, SIL, PLA, RVA, coord_rows)
    version(wb, here, coord_rows)
    SO = wb["Sources"]; rs = SO.max_row + 1
    SO.cell(rs, 1, "[Verified] 鴻海 2026 年 1–10 月月營收（C0〔實績〕；TechNews 各月月營收報導，MOPS 原文未直接讀取，網址見 Coord 頁第 4.1 節各列）")
    SO.cell(rs, 2, "https://finance.technews.tw/2026/11/05/oct-2026-foxconn/")
    SO.cell(rs + 1, 1, "[Interested-party] 法說摘要：1Q26、2Q26、3Q26 雲端網路產品占營收 48%、51%、54%（富果、壹蘋、數位時代、優分析）；3Q26 營收、營業利益、EPS：C0〔實績〕與錨點對照用，原簡報未直接讀取；公司未揭露 AI 伺服器營收金額")
    SO.cell(rs + 1, 2, "https://blog.fugle.tw/post/earnings-call-2317-2026-05-14；https://blog.fugle.tw/post/earnings-call-2317-2026-08-12；https://news.nextapple.com/finance/20261112/D566D07E5C054836DE510FB38F6781C9；https://www.bnext.com.tw/article/93040/2026-foxconn--earnings-call-")
    SO.cell(rs + 2, 1, "[Verified／第三方彙整] 損益表歷史 FY2022–FY2025（Hist 頁；S&P Global Market Intelligence 經 stockanalysis.com；含少數股權損益）；FY2025 與公司揭露相符")
    SO.cell(rs + 2, 2, "https://stockanalysis.com/quote/tpe/2317/financials/income-statement/；https://uanalyze.com.tw/articles/4369349056")
    wb._v046 = dict(HR=HR, PL=PLA, RV=RVA, VG=VG, ST=ST, C0=C0, RM=RMA)
    return f"v046_post：Hist {len(HR)} 列；PL 標準損益表 {PLA['start']} 起；RevDCF；Val 0.7；Sens 主圖表 {ST}；README v1.3 {len(RMA)} 列；C0 {'有' if C0 else '無（無 Coord 頁）'}"
