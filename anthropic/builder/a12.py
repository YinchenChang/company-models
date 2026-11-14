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


def checks(ctx):
    return []
