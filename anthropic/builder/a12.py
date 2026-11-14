"""Anthropic v0.2.1 標準損益表頁 PL（個股 HTML 版面規格 v1.3 D 節；docs/plan/個股HTML版面規格.md）。

- 口徑：GAAP（Anthropic 招股書只提供 GAAP 財報，另以「調整後營業損益」為主要 non-GAAP 指標；市場與媒體引用的營收、淨損皆為 GAAP）。
- 欄：C–E＝2023–2025（S-1/A 經查核年度）、F＝2026 前九月（S-1/A 期中，比較欄）、G–K＝2026–2030（本模型基準情境 B）。
- 歷史列取 SRC_ANT（S-1/A 損益表，SRC_ANT_594–621 與既有 410–412、422、433、547、548）；預測列只引用既有計算頁的列（Revenue、Cost、Funding）
  與 Val 頁的評價股數，另加兩個 Inputs（pl_cash_yield、pl_q4_oth_other）。
- 預測口徑：營收＝營收總額（截頂後；招股書以總額認列雲端市集銷售，雲端平台抽成列在銷售、行銷與合作夥伴費用）；
  營業成本＝推論算力成本（COST_InfCompute；2025 年 S-1 營收成本 2.773 為模型推論算力 3.226 的 0.86 倍，見說明欄）；
  營業利益＝Funding F06（營收淨額 − 算力 − 非算力 − 股權報酬＝營收總額 − 平台抽成 − 算力 − 非算力 − 股權報酬；2025 年 −8.061 對 S-1 −8.060）；
  營業費用＝毛利 − 營業利益（研發算力與閒置、人事與管銷、平台抽成、股權報酬）。
  其他收入＝現金利息（收益率 × 期初期末平均現金）；2026＝前九月實際 ＋ Q4 利息 ＋ Q4 其他（pl_q4_oth_other，可轉債轉換前公允價值變動未揭露，預設 0）。
  所得稅：2026＝前九月實際；2027 起＝稅率 × MAX(0, 稅前 − 可抵用虧損期初)，可抵用虧損自 2025 年底聯邦 NOL 8.2 起滾動（2026 加營業虧損；不計 80% 上限與州稅）。
  EPS：歷史＝招股書每股淨損（虧損年度基本＝稀釋，分母為特別股轉換前加權股數）；預測＝稅後淨利 ÷ Val 評價股數（IPO 後完全稀釋）。
- 本頁只列示，不被任何計算頁引用（Checks 除外）；新增本頁不改變任何既有工作表的數值。
"""
from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, header, nm, put, title

HC = "CDE"                 # 2023–2025
QC = "F"                   # 2026 前九月
FC = "GHIJK"               # 2026–2030
SRCC = "EFGHI"             # 來源頁 2026–2030 欄（D–I＝2025–2030）
ALL = HC + QC + FC
YEAR_COLS = HC + FC        # 年度欄（年增率用）


def build(ctx):
    D, wb = ctx.D, ctx.wb
    S, I = D.S, D.I
    from a2 import Sheet
    R, K, F = Sheet.all["R"], Sheet.all["K"], Sheet.all["F"]
    VP = ctx.P["a11"]["P"]
    ws = wb.create_sheet("PL")
    title(ws, "PL — 標準損益表（v0.2.1；GAAP；個股 HTML 版面規格 v1.3 D 節）：歷史 2023–2025＋2026 前九月（S-1/A）、2026–2030 本模型基準情境",
          "口徑：GAAP（招股書只有 GAAP 財報；non-GAAP 只揭露調整後營業損益）。預測列由既有計算頁推得：營收＝營收總額（截頂後）；營業成本＝推論算力；營業利益＝Funding F06；"
          "營業費用＝毛利 − 營業利益；其他收入＝現金利息；所得稅＝稅率 × 抵用 NOL 後的正稅前利益；EPS＝稅後淨利 ÷ 評價股數（IPO 後完全稀釋）。",
          "本頁只列示，不被任何計算頁引用。單位：$B（EPS 為 $／股、股數為十億股）。")
    header(ws, 4, ["項目", "單位", "2023", "2024", "2025", "2026 前九月", "2026", "2027", "2028", "2029", "2030", "說明"])
    put(ws, "A5", "期間（歷史＝S-1/A 實際；2026–2030＝本模型基準情境 B 預測）", F_NOTE)
    for c in HC + QC:
        put(ws, f"{c}5", "實際", F_NOTE)
    for c in FC:
        put(ws, f"{c}5", "預測", F_NOTE)
    P = {}
    r = [6]

    def row(key, label, unit, vals, note="", fmt="#,##0.000", name=None, bold=False):
        n = r[0]
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for col, v in zip(ALL, vals):
            if v is not None:
                put(ws, f"{col}{n}", v, fmt=fmt)
        if note:
            put(ws, f"L{n}", note, F_NOTE)
        P[key] = n
        if name:
            nm(wb, name, f"PL!$C${n}:$K${n}")
        r[0] += 1
        return n

    def sec(t):
        put(ws, f"A{r[0]}", t, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    def ref(key, col):
        return f"{col}{P[key]}"

    def yoy(key):
        """年增率：2024、2025 對前一年；前九月欄空白；2026 對 2025（E 欄）；之後對前一欄。"""
        out = [None]
        for j, col in enumerate(ALL[1:], start=1):
            if col == QC:
                out.append(None)
                continue
            prev = "E" if col == "G" else ALL[j - 1]
            out.append(f"={ref(key, col)}/{ref(key, prev)}-1")
        return out

    def ratio(num, den):
        return [f"={ref(num, c)}/{ref(den, c)}" for c in ALL]

    stub = f"(1-{I('m_9m')}/{I('months_per_year')})"
    sec("一、損益表（GAAP）")
    row("rev", "營收", "$B",
        [f"={S('rev_s1f_fy2023')}", f"={S('rev_s1f_fy2024')}", f"={S('rev_s1f_fy2025')}", f"={S('pros_rev_9m26')}"] +
        [f"=Revenue!{c}{R.rows['gross_c']}" for c in SRCC],
        note="歷史 SRC_ANT_410–412、547；預測＝營收總額（截頂後；REV_GrossCapped）：招股書以總額認列雲端市集銷售", name="PL_Rev", bold=True)
    row("yoy", "營收年增率（YoY）", "比例", yoy("rev"), fmt="0.0%", name="PL_RevYoY", note="2026 對 2025 全年；前九月欄不計")
    row("cogs", "營業成本（COGS）", "$B",
        [f"={S('is_s1a_cor_2023')}", f"={S('is_s1a_cor_2024')}", f"={S('cor_s1f_fy2025')}", f"={S('is_s1a_cor_9m26')}"] +
        [f"=Cost!{c}{K.rows['c_inf']}" for c in SRCC],
        note="歷史＝S-1/A 功能別費用表的營收成本（SRC_ANT_594、595、422、596）；預測＝推論算力成本（COST_InfCompute；2025 年 S-1 營收成本為模型推論算力的 0.86 倍，模型口徑偏保守）",
        name="PL_COGS")
    row("gp", "毛利", "$B", [f"={ref('rev', c)}-{ref('cogs', c)}" for c in ALL], name="PL_GP", bold=True)
    row("gpm", "毛利率", "比例", ratio("gp", "rev"), fmt="0.0%", name="PL_GPM")
    row("opex", "營業費用", "$B", [None] * len(ALL))   # 先佔位，營業利益列建好後回填
    row("oi", "營業利益", "$B",
        [f"={S('is_s1a_oi_2023')}", f"={S('is_s1a_oi_2024')}", f"={S('oi_s1f_fy2025')}", f"={S('pros_oi_9m26')}"] +
        [f"=Funding!{c}{F.rows['gap']}" for c in SRCC],
        note="歷史 SRC_ANT_597、598、433、548；預測＝Funding F06（營收淨額 − 算力 − 非算力 − 股權報酬；2025 年模型 −8.061 對 S-1 −8.060）", name="PL_OI", bold=True)
    for c in ALL:
        put(ws, f"{c}{P['opex']}", f"={ref('gp', c)}-{ref('oi', c)}", fmt="#,##0.000")
    put(ws, f"L{P['opex']}", "＝毛利 − 營業利益（研發算力與閒置、人事與管理、銷售行銷與雲端平台抽成、股權報酬；歷史含 2025 年 Bartz 和解 1.5）", F_NOTE)
    nm(wb, "PL_Opex", f"PL!$C${P['opex']}:$K${P['opex']}")
    row("oim", "營業利益率", "比例", ratio("oi", "rev"), fmt="0.0%", name="PL_OIM")
    # 其他收入：歷史 SRC；2026＝前九月實際＋Q4 利息＋Q4 其他；2027 起＝收益率 × 平均現金
    oth_f = [f"={S('is_s1a_oth_2023')}", f"={S('is_s1a_oth_2024')}", f"={S('is_s1a_oth_2025')}", f"={S('is_s1a_oth_9m26')}"]
    oth_f.append(f"={S('is_s1a_oth_9m26')}+{I('pl_cash_yield')}*Funding!E{F.rows['end']}*{stub}+{I('pl_q4_oth_other')}")
    oth_f += [f"={I('pl_cash_yield')}*(Funding!{c}{F.rows['open']}+Funding!{c}{F.rows['end']})/{I('arith_k2')}" for c in SRCC[1:]]
    row("oth", "其他收入（淨額）", "$B", oth_f,
        note="歷史 SRC_ANT_599–602（利息收入＋可轉債等公允價值變動＋其他；2025 年公允價值變動 −34.3、2026 前九月 −94.3，非現金）；"
             "2026＝前九月實際＋Q4 利息（收益率 × 年底現金 × 0.25）＋Q4 其他（pl_q4_oth_other，預設 0）；2027 起＝收益率 × 期初期末平均現金（FND_CashOpen、FND_CashEnd）",
        name="PL_Oth")
    row("pbt", "稅前利益", "$B", [f"={ref('oi', c)}+{ref('oth', c)}" for c in ALL], name="PL_PBT", bold=True)
    # 所得稅與可抵用虧損
    tax_r = r[0]
    nol_b, nol_e = tax_r + 5, tax_r + 6        # 下方「二、」節的列（先算位置）
    tax_f = [f"={S('is_s1a_tax_2023')}", f"={S('is_s1a_tax_2024')}", f"={S('is_s1a_tax_2025')}", f"={S('is_s1a_tax_9m26')}", f"={S('is_s1a_tax_9m26')}"]
    tax_f += [f"={I('val_tax')}*MAX(0,{c}{P['pbt']}-{c}{nol_b})" for c in FC[1:]]
    row("tax", "所得稅", "$B", tax_f,
        note="歷史 SRC_ANT_603–606；2026＝前九月實際（Q4 虧損、評價備抵，不再提列）；2027 起＝稅率 21% × MAX(0, 稅前利益 − 可抵用虧損期初)", name="PL_Tax")
    row("ni", "稅後淨利", "$B", [f"={ref('pbt', c)}-{ref('tax', c)}" for c in ALL], name="PL_NI", bold=True)
    row("nim", "淨利率", "比例", ratio("ni", "rev"), fmt="0.0%", name="PL_NIM")
    row("eps", "EPS（完全稀釋）", "$/股",
        [f"={S('is_s1a_eps_2023')}", f"={S('is_s1a_eps_2024')}", f"={S('is_s1a_eps_2025')}", f"={S('is_s1a_eps_9m26')}"] +
        [f"={ref('ni', c)}/VAL_Shares" for c in FC],
        fmt="#,##0.00", name="PL_EPS",
        note="歷史＝招股書每股淨損（虧損年度基本＝稀釋；分母為特別股轉換前加權股數，見下方股數列）；預測＝稅後淨利 ÷ 評價股數（VAL_Shares，IPO 後完全稀釋，含選擇權庫藏股法與 RSU）。兩段分母口徑不同，不宜直接比較")
    assert r[0] == nol_b - 1, (r[0], nol_b)
    sec("二、EPS 分母與可抵用虧損（附屬列）")
    assert r[0] == nol_b
    nb = [None] * 5 + [f"={S('tax_s1a_nol_fed_2025')}+MAX(0,-G{P['oi']})"] + [f"={FC[j - 1]}{nol_e}" for j in range(2, 5)]
    row("nolb", "可抵用虧損期初（2027＝2025 年底聯邦 NOL 8.2＋2026 營業虧損）", "$B", [None] * 5 + nb[5:], note="SRC_ANT_621；不計 80% 抵用上限、州稅與研發扣抵（簡化）")
    row("nole", "可抵用虧損期末＝期初 ＋ 虧損 − 抵用", "$B",
        [None] * 5 + [f"={c}{nol_b}+MAX(0,-{c}{P['pbt']})-MIN({c}{nol_b},MAX(0,{c}{P['pbt']}))" for c in FC[1:]])
    assert P["nolb"] == nol_b and P["nole"] == nol_e
    row("sh", "EPS 分母股數", "十億股",
        [f"={S('is_s1a_wsh_2023')}/{I('m_to_b')}", f"={S('is_s1a_wsh_2024')}/{I('m_to_b')}", f"={S('is_s1a_wsh_2025')}/{I('m_to_b')}", f"={S('is_s1a_wsh_9m26')}/{I('m_to_b')}"] +
        ["=VAL_Shares"] * 5, fmt="0.000", name="PL_Shares",
        note="歷史＝招股書加權平均股數（SRC_ANT_615–618）；預測＝評價股數（Val 頁）")

    sec("三、比率（表尾）")
    row("r_yoy", "營收年增率", "比例", [f"={ref('yoy', c)}" if ws[f"{c}{P['yoy']}"].value is not None else None for c in ALL], fmt="0.0%")
    row("r_gpm", "毛利率", "比例", [f"={ref('gpm', c)}" for c in ALL], fmt="0.0%")
    row("r_oim", "營業利益率", "比例", [f"={ref('oim', c)}" for c in ALL], fmt="0.0%")
    row("r_nim", "淨利率", "比例", [f"={ref('nim', c)}" for c in ALL], fmt="0.0%")
    niy = [None]
    for j, col in enumerate(ALL[1:], start=1):
        if col == QC:
            niy.append(None)
            continue
        prev = "E" if col == "G" else ALL[j - 1]
        niy.append(f'=IF({ref("ni", prev)}<=0,"不適用",{ref("ni", col)}/{ref("ni", prev)}-1)')
    row("r_niy", "淨利年增率（前期虧損時不適用）", "比例", niy, fmt="0.0%", name="PL_NIYoY")

    sec("四、情境 A／C 對照（取自 Val 頁情境來源；A、C 為快照值）")
    for s in ("A", "B", "C"):
        row(f"sc_rev_{s}", f"營收（情境 {s}）", "$B", [None] * 4 + [f"=Val!{c}{VP[s + '_gross']}" for c in "CDEFG"], name=f"PL_Rev_{s}")
    for s in ("A", "B", "C"):
        row(f"sc_oi_{s}", f"營業利益（情境 {s}）", "$B", [None] * 4 + [f"=Val!{c}{VP[s + '_ebit']}" for c in "CDEFG"], name=f"PL_OI_{s}")

    sec("五、檢查（應為 0）")
    row("chk_ni", "歷史：計算稅後淨利 − 招股書淨損（2023、2024、2025、2026 前九月）", "$B",
        [f"={ref('ni', c)}-{S(k)}" for c, k in zip(HC + QC, ("is_s1a_ni_2023", "is_s1a_ni_2024", "is_s1a_ni_2025", "is_s1a_ni_9m26"))] + [None] * 5,
        note="SRC_ANT_607–610")
    row("chk_b", "情境 B 對照列 − 本表（營收、營業利益；2026–2030）", "$B",
        [None] * 4 + [f"=ABS({ref('sc_rev_B', c)}-{ref('rev', c)})+ABS({ref('sc_oi_B', c)}-{ref('oi', c)})" for c in FC])
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 9
    for col in ALL:
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["L"].width = 90
    ws.freeze_panes = "C6"
    ctx.P["a12"] = dict(P=P)
    build_revdcf(ctx)


def build_revdcf(ctx):
    """RevDCF 頁（v0.2.1；個股 HTML 版面規格 v1.3 A 節「反向 DCF」）：以現價（IPO 區間中點）反解現價隱含的
    (a) WACC（二分法）、(b) 2027–2033 營收等比例倍數（二分法；累積虧損扣抵逐年重算）、(c) 終值期營業利益率（封閉解）、
    (d) Comparable 倍數與相對同業係數（封閉解）。DCF 以情境 B 路徑反解「情境 B 每股價值推到目標時點＝現價」。
    只讀 Val 頁與 Inputs，無任何頁引用本頁，不影響目標價；各段附核對列（應為 0）。"""
    D, wb = ctx.D, ctx.wb
    I = D.I
    VP = ctx.P["a11"]["P"]
    ws = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    title(ws, "RevDCF — 反向 DCF：現價隱含什麼（v0.2.1；個股 HTML 版面規格 v1.3）：一次只動一個驅動、其餘維持基準，反解到現價（IPO 區間中點）",
          "DCF：以情境 B 路徑反解「情境 B 評價日每股價值 ×（1＋股權成本）^目標年數＝現價」；(a) WACC 與 (b) 營收倍數用二分法 40 次，(c) 終值期營業利益率與 (d) Comparable 倍數為封閉解。"
          "（DCF 目標價為三情境加權，情境 B 為基準路徑。）反解值超出合理範圍時照實列出，不外推。",
          "本頁只讀 Val 頁與 Inputs，無任何頁引用本頁。單位：$B、十億股、$／股。")
    V = lambda key, col="C": f"Val!${col}${VP[key]}"  # noqa: E731
    R = {}
    r = [5]

    def kv(key, label, f, fmt="#,##0.000", note="", name=None, bold=False):
        n = r[0]
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", f, fmt=fmt)
        if note:
            put(ws, f"C{n}", note, F_NOTE)
        R[key] = n
        if name:
            nm(wb, name, f"RevDCF!$B${n}")
        r[0] += 1
        return n

    def sec(t_):
        put(ws, f"A{r[0]}", t_, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    b = lambda key: f"$B${R[key]}"  # noqa: E731
    sec("一、支持現價所需的 DCF 企業價值（情境 B 路徑）")
    kv("px", "現價（IPO 區間中點；美元）", "=VAL_Px", "0.00", name="RVD_Px")
    kv("ps0r", "支持現價所需的評價日每股價值＝現價 ÷（1＋股權成本）^目標年數", f"=VAL_Px/(1+VAL_Ke)^{I('val_t_tgt')}", "0.00", name="RVD_PS0Req")
    kv("ps0b", "模型：情境 B 評價日每股價值", f"={V('ps0', 'D')}", "0.00", name="RVD_PS0B")
    kv("evr", "支持現價所需的企業價值 EV＝每股 × 評價股數 − 淨現金", f"={b('ps0r')}*VAL_Shares-VAL_NetCash", name="RVD_EVReq", bold=True)
    kv("evb", "模型：企業價值 EV（情境 B）", "=VAL_EV_B", name="RVD_EVB")
    kv("evx", "EV 需為模型的倍數", f"={b('evr')}/{b('evb')}", "0.00", name="RVD_EVMult")

    # (a) WACC
    sec("二、(a) 現價隱含的 WACC（各年自由現金流不變；永續期新投資報酬＝WACC＋超額報酬差，同 Val 頁；二分法 40 次）")
    cols = "CDEFGHI"                           # Val 明確期 2026Q4–2032

    def val_w(w):
        pv = "+".join(f"Val!${c}${VP['dB_fcff']}/(1+{w})^Val!${c}${VP['t']}" for c in cols)
        tv = (f"Val!$D${VP['tvn']}*(1-VAL_g/({w}+{I('val_ronic_add')}))/({w}-VAL_g)/(1+{w})^(Val!$I${VP['t']}+Val!$C${VP['mid']})")
        return f"(({pv})+{tv}+VAL_NetCash)/VAL_Shares"
    kv("wchk", "核對：WACC＝模型值時重算的評價日每股價值 − 模型情境 B（應為 0）", f"={val_w('VAL_WACC')}-{b('ps0b')}", "0.000000")
    header(ws, r[0], ["次", "下界", "上界", "中點", "中點的評價日每股價值"])
    r[0] += 1
    w0 = r[0]
    for i in range(40):
        n = r[0]
        put(ws, f"A{n}", i + 1)
        if i == 0:
            put(ws, f"B{n}", f"=VAL_g+{I('val_bisect_lo_add')}", fmt="0.0000%")
            put(ws, f"C{n}", f"={I('val_bisect_w_hi')}", fmt="0.0000%")
        else:
            put(ws, f"B{n}", f"=IF(E{n - 1}>{b('ps0r')},D{n - 1},B{n - 1})", fmt="0.0000%")
            put(ws, f"C{n}", f"=IF(E{n - 1}>{b('ps0r')},C{n - 1},D{n - 1})", fmt="0.0000%")
        put(ws, f"D{n}", f"=(B{n}+C{n})/{I('arith_k2')}", fmt="0.0000%")
        put(ws, f"E{n}", "=" + val_w(f"D{n}"), fmt="0.00")
        r[0] += 1
    wl = r[0] - 1
    kv("w1", "現價隱含的 WACC", f"=D{wl}", "0.00%", name="RVD_WACC", bold=True)
    kv("w0", "模型：WACC", "=VAL_WACC", "0.00%", name="RVD_WACC0")
    kv("w1chk", "核對：隱含 WACC 下的每股價值 − 支持現價所需（應為 0）", f"=E{wl}-{b('ps0r')}", "0.000000", name="RVD_WACCChk")

    # (b) 營收倍數
    sec("三、(b) 現價隱含的營收規模（2027–2033 營收淨額等比例放大 s 倍；各年營業利益率、資本支出比率不變；2026Q4 不動；累積虧損扣抵與稅逐年重算；二分法 40 次）")
    ycol = "DEFGHI"                            # 2027–2032（明確期現值；2033 只進終值）
    tax, inf, wcr = I("val_tax"), I("val_nol_inf_switch"), I("val_wc_rate")
    hdr = ["次", "下界", "上界", "中點 s"] + [f"扣抵期初 {y}" for y in range(2027, 2033)] + [f"稅 {y}" for y in range(2027, 2033)] + ["EV", "評價日每股價值"]
    header(ws, r[0], hdr)
    r[0] += 1
    NC = ["E", "F", "G", "H", "I", "J"]        # 扣抵期初欄
    TC = ["K", "L", "M", "N", "O", "P"]        # 稅欄

    def srow(n, s_ref):
        for j, vc in enumerate(ycol):
            eb = f"{s_ref}*Val!${vc}${VP['dB_ebit']}"
            if j == 0:
                put(ws, f"{NC[j]}{n}", f"=Val!$C${VP['dB_nolc']}", fmt="#,##0.000")
            else:
                pe = f"{s_ref}*Val!${ycol[j - 1]}${VP['dB_ebit']}"
                put(ws, f"{NC[j]}{n}", f"={NC[j - 1]}{n}+MAX(0,-{pe})-MIN({NC[j - 1]}{n},MAX(0,{pe}))", fmt="#,##0.000")
            put(ws, f"{TC[j]}{n}", f"={tax}*MAX(0,{eb}-{NC[j]}{n})*(1-{inf})", fmt="#,##0.000")
        terms = []
        for j, vc in enumerate(ycol):
            wc = (f"(Val!$D${VP['dB_wc']}+{wcr}*({s_ref}-1)*Val!$D${VP['dB_rev']})" if j == 0 else f"{s_ref}*Val!${vc}${VP['dB_wc']}")
            terms.append(f"({s_ref}*Val!${vc}${VP['dB_ebit']}-{TC[j]}{n}-{s_ref}*Val!${vc}${VP['dB_cx']}-{wc})*Val!${vc}${VP['df']}")
        put(ws, f"Q{n}", f"=Val!$C${VP['dB_pv']}+" + "+".join(terms) + f"+{s_ref}*VAL_PVTV_B", fmt="#,##0.000")
        put(ws, f"R{n}", f"=(Q{n}+VAL_NetCash)/VAL_Shares", fmt="0.00")
    n = r[0]
    put(ws, f"A{n}", "核對 s＝1")
    put(ws, f"D{n}", f"={I('months_per_year')}/{I('months_per_year')}", fmt="0.0000")
    srow(n, f"D{n}")
    R["s_base"] = n
    r[0] += 1
    s0 = r[0]
    for i in range(40):
        n = r[0]
        put(ws, f"A{n}", i + 1)
        if i == 0:
            put(ws, f"B{n}", f"={I('val_bisect_s_lo')}", fmt="0.0000")
            put(ws, f"C{n}", f"={I('val_bisect_s_hi')}", fmt="0.0000")
        else:
            put(ws, f"B{n}", f"=IF(R{n - 1}>{b('ps0r')},B{n - 1},D{n - 1})", fmt="0.0000")
            put(ws, f"C{n}", f"=IF(R{n - 1}>{b('ps0r')},D{n - 1},C{n - 1})", fmt="0.0000")
        put(ws, f"D{n}", f"=(B{n}+C{n})/{I('arith_k2')}", fmt="0.0000")
        srow(n, f"D{n}")
        r[0] += 1
    sl = r[0] - 1
    kv("schk0", "核對：s＝1 重算的 EV − 模型 EV（情境 B；應為 0）", f"=Q{R['s_base']}-VAL_EV_B", "0.000000", name="RVD_ScaleChk0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=D{sl}", "0.000", name="RVD_RevMult", bold=True)
    kv("r30n", "現價隱含的 2030 營收淨額", f"={b('s1')}*Val!$G${VP['dB_rev']}", name="RVD_Rev2030Net")
    kv("r30n0", "模型：2030 營收淨額（情境 B）", f"=Val!$G${VP['dB_rev']}", name="RVD_Rev2030Net0")
    kv("r30g", "現價隱含的 2030 營收總額（同倍數）", f"={b('s1')}*Val!$G${VP['B_gross']}", name="RVD_Rev2030Gross")
    kv("r30g0", "模型：2030 營收總額（情境 B）", f"=Val!$G${VP['B_gross']}", name="RVD_Rev2030Gross0")
    kv("schk", "核對：隱含倍數下的每股價值 − 支持現價所需（應為 0）", f"=R{sl}-{b('ps0r')}", "0.000000", name="RVD_ScaleChk")

    # (c) 終值期營業利益率
    sec("四、(c) 現價隱含的終值期營業利益率（只改 2033 年營業利益率；終值 NOPAT 與其成正比；封閉解）")
    kv("mJ", "模型：2033 年營業利益率（情境 B；終值基礎）", f"=Val!$J${VP['dB_m']}", "0.0%", name="RVD_TermMargin0")
    kv("m30", "對照：模型 2030 年營業利益率（情境 B）", f"=Val!$G${VP['dB_m']}", "0.0%", name="RVD_Margin2030")
    kv("m1", "現價隱含的終值期營業利益率＝模型值 ×（終值現值 ＋ 所需 EV − 模型 EV）÷ 終值現值", f"={b('mJ')}*(VAL_PVTV_B+{b('evr')}-{b('evb')})/VAL_PVTV_B", "0.0%",
       name="RVD_TermMargin", bold=True)
    kv("mchk", "核對：明確期現值 ＋ 終值現值 × 隱含利益率 ÷ 模型利益率 − 所需 EV（應為 0）", f"=VAL_PVExp_B+VAL_PVTV_B*{b('m1')}/{b('mJ')}-{b('evr')}", "0.000000", name="RVD_TermChk")

    # (d) Comparable
    sec("五、(d) Comparable：現價隱含的營收倍數與相對同業係數（情境 B 的未來 12 個月營收總額；封閉解）")
    kv("cev", "現價隱含的企業價值＝現價 × 評價股數 − 淨現金", "=VAL_Px*VAL_Shares-VAL_NetCash", name="RVD_CompEV")
    kv("cm", "現價隱含的 EV／未來 12 個月營收倍數", f"={b('cev')}/VAL_NTMRev", "0.00", name="RVD_CompMult", bold=True)
    kv("pm", "模型：同業 EV／遠期營收中位數（未調整）", "=VAL_PeerMult", "0.00", name="RVD_PeerMult")
    kv("am", "模型：調整後倍數＝同業中位數 × 調整係數（情境 B）", "=VAL_PeerMult*VAL_AdjFactor", "0.00", name="RVD_AdjMult")
    kv("cc", "現價隱含的相對同業係數＝隱含倍數 ÷ 同業中位數", f"={b('cm')}/VAL_PeerMult", "0.000", name="RVD_CompCoef", bold=True)
    kv("ac", "模型：調整係數（情境 B）", "=VAL_AdjFactor", "0.000", name="RVD_AdjCoef")
    kv("cchk", "核對：（隱含倍數 × 營收 ＋ 淨現金）÷ 股數 − 現價（應為 0）", f"=({b('cm')}*VAL_NTMRev+VAL_NetCash)/VAL_Shares-VAL_Px", "0.000000", name="RVD_CompChk")
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.freeze_panes = "A5"
    ctx.P["a12"]["R"] = R


def checks(ctx):
    return []
