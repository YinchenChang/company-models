"""Anthropic v0.2.0：初步招股說明書（2026-11-10）滾入（工作單 20261111 v0.2.0 A–E；原文 docs/reports/20261111_Prospectus原文摘錄.md）。
Cost 第二十七節（含 D1／D2 對照：Revenue 頁尾有價格事件表，不附加列）、Funding 第十六節，皆附加在頁尾，既有列編號不變。估值頁 Val 另在 a11.py。

- B1（p_ipo_base_switch）：IPO＋同步私募淨募資 − RSU 代扣稅款進 2026 基準股權流入（Funding 第三節 c_ipo、第四節 eq 引用本節 p_ipo_net、p_ipo_base）；
  淨募資與代扣稅款依每 $1 敏感度隨 ipo_px 連動；超額配售由 p_overallot_switch 控制（基準 0）。
- B2（p_cond_pp_switch）：Google NPA 剩餘 20 → 10、NVIDIA 10 → 併入私募（0）、AMD 5 → 2.5（a4 第三節公式直接選用）。
- B3：9-30 流動性與 pro forma 對照（只對照）。
- C1–C5：9-30 後新簽 105.7（a9 s1a_post）、Microsoft 47.2（a3 合約 T）、附條件 186.2（a9 s1a_cond）、Lambda 34.9 扣除讀法（本節 p_lamd → a7 s1u_raw）、
  2030 年前付款 451.6 對照（只對照）。
- D1（p_q4_rr_switch）：2026Q4 營收＝10-31 Run-Rate Revenue ÷ 季數（a2 target 公式）；D2：2027 營收總額 ÷ 2026Q4 年化（Checks WARN）。
- E1（p_sbc_ipo_switch）：2026 股權報酬加 14.2（本節 p_sbcipo → a3 sbc；不進每人基準）；E2（p_sbc_floor_switch）：2027–2028 股權報酬下限＝24.1 ÷ 2.0 年直線攤銷（p_sbcfl）。
Python 只產生結構；數值一律引用 SRC_ANT／Inputs／既有列（E6）。
"""
from a2 import Sheet  # noqa: F401


def build(ctx):
    D = ctx.D
    S, I = D.S, D.I
    from a2 import Sheet as _S
    R, K, F, C = _S.all["R"], _S.all["K"], _S.all["F"], _S.all["C"]
    months, qpy, m9 = I("months_per_year"), I("quarters_per_year"), I("m_9m")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731
    zero = "(«C:vr_unit@$»-«C:vr_unit@$»)"
    one = "«C:vr_unit@$»"
    y26, y27, y28 = I("year_2026"), I("year_2027"), I("year_2028")

    # ═════════════ Cost 第二十七節：C1–C5、E1–E2 ═════════════
    K.section("二十七、v0.2.0 初步招股說明書（2026-11-10）：C1 9-30 後新簽 105.7、C2 Microsoft 47.2、C3 附條件 186.2、C4 Lambda 34.9 讀法、C5 2030 年前付款 451.6 對照、E1 RSU 修改 14.2、E2 未認列股權報酬攤銷下限")
    K.add("p_tot", "C1 初步招股說明書：承諾總額（S-1/A 895.4）", "$B", only25(f"={S('cmt_pros_total')}"), "SRC_ANT_567")
    K.add("p_totc", "C1 檢查：三段（表內＋未起租＋9-30 後）＋短期 − 總額（原文捨入）", "$B", only25("=«s1a_seg@$»+«s1a_short@$»-«p_tot@$»"), "開關 1 時約 −0.007")
    K.add("p_seg_old", "C1 對照：S-1/A 三段合計（表內＋未起租＋100.2）", "$B", only25(f"=«s1a_tbl@$»+«s1a_lnc@$»+{S('cmt_s1a_post0930')}"), "871.4")
    K.add("p_segd", "C1 差：本版三段 − S-1/A 三段", "$B", only25("=«s1a_seg@$»-«p_seg_old@$»"), "開關 1 時 +5.5")
    K.add("p_msft", "C2 Microsoft 合約：下限（D）／上限（E；只列）／S-1/A（F）", "$B",
          [f"={S('cmt_pros_msft')}", f"={S('cmt_pros_msft')}_Hi", f"={S('cmt_s1a_msft')}", None, None, None], "SRC_ANT_573、514；欄位不是年度")
    K.add("p_cond", "C3 附條件或可終止承諾：本版（D）／S-1/A（E）", "$B", [f"={S('cmt_pros_cond')}", f"={S('cmt_s1a_conditional')}", None, None, None, None],
          "SRC_ANT_569、508；R5d 扣除讀法用 D 欄（開關 p_cond_cmt_switch）")
    K.add("p_lam", "C4 Lambda 8 月合約（目前可終止；公司無意終止）", "$B", only25(f"={S('cmt_pros_lambda_term')}"), "SRC_ANT_575；Lambda 合計 41.3（SRC_ANT_574）")
    K.add("p_lamd", "C4 自 W1(b) 三段合計扣除的 Lambda 額（開關 × 34.9；基準 0＝照計）", "$B", only25(f"={I('p_lambda_deduct_switch')}*«p_lam@$»"),
          "第二十五節 s1u_raw 減本列（J4 預設照計）", name="COST_ProsLambdaDeduct", name_cols="D")
    K.add("p_n12", "C 對照：未來 12 個月預計支付 70.7（D；其中附條件 4.2，E）vs 模型 2026Q4＋2027 前三季合約實付（F；年內平均）", "$B",
          [f"={S('cmt_pros_next12m')}", f"={S('cmt_pros_next12m_cond')}", f"=«pay@E»/{qpy}+«pay@F»*({qpy}-{one})/{qpy}", None, None, None], "SRC_ANT_570、571；欄位不是年度")
    # C5：2026Q4–2030 合約付款 vs 451.6
    K.add("p_c5q4", "C5 模型 2026 Q4 合約實付（年內平均＝全年 ÷ 季數）", "$B", only26(f"=«pay»/{qpy}"), "近似（年內平均）")
    K.add("p_c5m", "C5 模型 2026Q4–2030 合約實付合計（全部）", "$B", only25("=«p_c5q4@E»+SUM(«pay@F»:«pay@I»)"), "")
    K.add("p_c5nf", "C5 其中：無固定金額安排（AMD、Google 額外 TPU）的模型計入額", "$B",
          only25(f"=«pay_amd@E»/{qpy}+SUM(«pay_amd@F»:«pay_amd@I»)+«pay_gadd@E»/{qpy}+SUM(«pay_gadd@F»:«pay_gadd@I»)"), "不在承諾總額內（S-1/A 腳註 5）")
    K.add("p_c5x", "C5 模型（不含無固定金額安排）", "$B", only25("=«p_c5m@$»-«p_c5nf@$»"), "")
    K.add("p_c5p", "C5 初步招股說明書：承諾中 2030-12-31 前發生", "$B", only25(f"={S('cmt_pros_by2030')}"), "SRC_ANT_568")
    K.add("p_c5d", "C5 差：模型（不含無固定金額）− 451.6", "$B", only25("=«p_c5x@$»-«p_c5p@$»"), "只對照", name="COST_ProsBy2030Gap", name_cols="D")
    K.add("p_c5r", "C5 比例：模型（不含無固定金額）÷ 451.6 − 1", "比例", only25(f"=«p_c5x@$»/«p_c5p@$»-{one}"),
          "取代既有限制「2027–2030 合約付款比推算高約 22%」的口徑（推算 → 原文）", name="COST_ProsBy2030Ratio", name_cols="D")
    # E1／E2
    K.add("p_sbcipo", "E1 pro forma RSU 修改追加股權報酬（2026Q4 一次性；開關 × 14.2；不進每人基準、不外推）", "$B",
          lambda i, c: f"={I('p_sbc_ipo_switch')}*{S('pros_sbc_pf_mod')}*({c}$6={y26})", "SRC_ANT_551；非現金", name="COST_SBC_IPO")
    K.add("p_sbcr", "E2 未認列股權報酬（RSU）年攤銷＝24.1 ÷ 2.0 年", "$B/年", only25(f"={S('pros_sbc_unrec')}/{S('pros_sbc_unrec_yrs')}"), "SRC_ANT_553、554；直線（Derived）")
    stub = f"({one}-{m9}/{months})"
    K.add("p_sbca", "E2 攤銷時程（2026＝Q4；2027＝全年；2028＝餘額）", "$B",
          lambda i, c: (f"=«p_sbcr@$»*{stub}" if i == 1 else f"=«p_sbcr@$»*({c}$6={y27})+({S('pros_sbc_unrec')}-«p_sbca@E»-«p_sbcr@$»)*({c}$6={y28})" if i >= 2 else f"={zero}"),
          "2026Q4 3.01、2027 12.05、2028 9.04", name="COST_SBC_Amort")
    K.add("p_sbcfl", "E2 股權報酬下限（開關 × 攤銷；只 2027、2028）", "$B",
          lambda i, c: f"={I('p_sbc_floor_switch')}*«p_sbca»*(({c}$6={y27})+({c}$6={y28}))", "Cost 第五節 sbc＝MAX(模型值, 本列)", name="COST_SBC_Floor")
    K.add("p_sbcb", "E2 對照：未套下限的模型股權報酬（人數 × 每人＋捐贈＋E1）", "$B", lambda i, c: f"=«hc»*«sbc_pp»/{I('m_to_b')}+«s1s_chc»+«p_sbcipo»", "")
    K.add("p_sbcd", "E2 下限追加額＝股權報酬 − 未套下限值", "$B", lambda i, c: "=«sbc»-«p_sbcb»", "非現金")
    K.add("p_sbcpf9", "對照：pro forma 績效條件 RSU 9M26 追加 7.487（只列示，不計）", "$B", only26(f"={S('pros_sbc_pf_9m')}"), "SRC_ANT_552")

    # ═════════════ Cost 第二十七之一節：D1／D2（Revenue 頁尾有價格事件表，不能附加列） ═════════════
    K.section("二十七之一、v0.2.0 Revenue 對照（Revenue 頁尾有價格事件表，列放 Cost 頁）：D1 2026Q4 營收改用 10-31 Run-Rate Revenue ÷ 季數（p_q4_rr_switch）、D2 2027 營收總額 ÷ 2026Q4 年化（只對照、不校準）")
    sq3 = I("s1a_q3_switch")
    q3 = f"({sq3}*{S('rev_s1a_2026q3')}+(1-{sq3})*{S('rev_2026_q3')})"
    K.add("p_rr", "初步招股說明書：Run-Rate Revenue 2026-10-31（D）／2026-09-30 推算（E）", "$B/年", [f"={S('pros_rr_1031')}", f"={S('pros_rr_0930')}", None, None, None, None],
          "SRC_ANT_557、559；欄位不是年度；公司註明不等於營收預測")
    K.add("p_q3", "2026 Q3 營收（S-1/A 實績）", "$B", only26(f"={q3}"), "SRC_ANT_488")
    K.add("p_q4", "2026 Q4 營收（模型）＝2026 營收總額基準 − 上半年實際 − Q3", "$B", only26("=«R:target@E»-«R:h1act@E»-«p_q3»"),
          "開關 1：81 ÷ 4＝20.25；開關 0：Q3 ×（1＋q4_growth）＝16.956", name="REV_Q4_2026", name_cols="E")
    K.add("p_q4o", "對照：v0.1.9 讀法 Q4＝Q3 ×（1＋q4_growth）", "$B", only26(f"={q3}*(1+{I('q4_growth')})"), "")
    K.add("p_q4a", "2026 Q4 年化＝Q4 × 季數", "$B/年", only26(f"=«p_q4»*{qpy}"), "", name="REV_Q4Ann_2026", name_cols="E")
    K.add("p_d2", "D2：2027 營收總額（截頂後）÷ 2026 Q4 年化", "倍", [None, None, "=«R:gross_c»/«p_q4a@E»", None, None, None],
          "< 1.0 立 WARN（Checks）：模型 2027 年平均低於 2026 年底年化；只對照、不校準", name="REV_D2Ratio", name_cols="F")
    K.add("p_q4g", "對照：2026 Q4 ÷ Q3 − 1（隱含季增率）", "比例", only26(f"=«p_q4»/«p_q3»-{one}"), "")

    # ═════════════ Funding 第十六節：B1–B3 ═════════════
    F.section("十六、v0.2.0 初步招股說明書（2026-11-10）：B1 IPO＋同步私募進基準（p_ipo_base_switch）、B2 條件式改讀、B3 流動性與 pro forma 對照")
    px, mid = I("ipo_px"), S("pros_px_mid")
    F.add("p_px", "現價／發行價（ipo_px；區間中點 74，低 70、高 78）", "$/股", only25(f"={px}"), "INP；SRC_ANT_522")
    F.add("p_net", "B1 公司淨募資（IPO＋私募；不含超額配售）＝84.6 ＋ 0.793 ×（價格 − 74）", "$B", only25(f"={S('pros_net_proceeds')}+{S('pros_net_per_dollar')}*(«p_px@$»-{mid})"),
          "SRC_ANT_534、536")
    F.add("p_wh", "B1 RSU 淨結算代扣稅款＝3.8 ＋ 0.052 ×（價格 − 74）", "$B", only25(f"={S('pros_rsu_withhold')}+{S('pros_rsu_withhold_per_dollar')}*(«p_px@$»-{mid})"),
          "SRC_ANT_537、538；以募資支付")
    F.add("p_ipo_net", "B1 進基準的 IPO＋私募淨額＝淨募資 − 代扣稅款", "$B", only25("=«p_net@$»-«p_wh@$»"), "價格 74 時 80.8", name="FND_IPONet", name_cols="D")
    F.add("p_oa", "B1 超額配售淨額（不計；Sens）＝（93.7 − 84.6）× 價格 ÷ 74", "$B",
          only25(f"=({S('pros_net_proceeds_oa')}-{S('pros_net_proceeds')})*«p_px@$»/{mid}"), "SRC_ANT_535、534（Derived）")
    F.add("p_ipo_base", "B1 計入基準的 IPO 流入（開關 × IPO 列＋超額配售開關 × 超額配售淨額）", "$B",
          lambda i, c: f"={I('p_ipo_base_switch')}*«c_ipo»+{I('p_overallot_switch')}*«p_oa@$»*({c}$6={I('ipo_year')})", "第四節 eq 引用本列", name="FND_IPOInBase")
    F.add("p_pp", "B1 其中：同步私募（Google 10、NVIDIA 10、AMD 2.5、Broadcom 1、Micron 1、Microsoft 1、TSMC 0.5）", "$B",
          only25("=" + "+".join(S(f"pros_pp_{k}") for k in ("google", "nvda", "amd", "avgo", "mu", "msft", "tsm"))), "SRC_ANT_526–532")
    F.add("p_ppc", "B1 檢查：私募逐家合計 − 26.0", "$B", only25(f"=«p_pp@$»-{S('pros_pp_total')}"), "應為 0")
    F.add("p_pp_sh", "B1 私募股數（百萬股）＝26.0 ÷ 價格 × 1000", "百萬股", only25(f"={S('pros_pp_total')}/«p_px@$»*{I('m_to_b')}"),
          "SRC_ANT_525；價格 74 時 351.35（SRC_ANT_533）")
    F.add("p_cond", "B2 條件式（2027）：Amazon 15＋Google（本版 10）＋AMD（本版 2.5）；NVIDIA（本版 0）", "$B",
          only25("=«c_amzn@F»+«c_googl@F»+«c_amd@F»+«c_nvda@E»"), "第三節")
    # B3
    F.add("p_liq", "B3 9-30 流動性（發行前）＝現金 95.6＋短投 36.3＋循環額度可用 14.8", "$B", only26(f"={S('pros_liq_0930')}"), "SRC_ANT_561；只對照")
    F.add("p_liqc", "B3 檢查：流動性 −（S-1/A 9-30 現金與短投＋循環額度可用）", "$B", only26(f"=«p_liq»-({S('cash_s1a_2026_09_30')}+{S('pros_revolver_avail')})"), "原文捨入")
    F.add("p_pf", "B3 pro forma as adjusted 現金、約當現金與短投（9-30；以 74；已扣代扣稅款）", "$B", only26(f"={S('pros_cash_pf')}"), "SRC_ANT_560")
    F.add("p_pfm", "B3 pro forma 依價格連動＝212.684 ＋（淨募資 − 代扣稅款）對價格的變動", "$B",
          only26(f"=«p_pf»+(«p_ipo_net@$»-({S('pros_net_proceeds')}-{S('pros_rsu_withhold')}))"), "估值頁淨現金使用", name="FND_ProFormaCash", name_cols="E")
    F.add("p_q4f", "B3 模型 2026 Q4 自由現金流＝2026 全年 − 前九月（S-1/A 營業現金流 − 資本支出 − 租賃資產預付）", "$B",
          only26(f"=«fcf»-({S('ocf_s1a_2026_9m')}-{S('capex_s1a_2026_9m')}-{S('leasepre_s1a_2026_9m')})"), "SRC_ANT_518–520")
    F.add("p_pfx", "B3 模型 2026 年底現金 − Q4 自由現金流（模型隱含 9-30 pro forma）", "$B", only26("=«end»-«p_q4f»"), "")
    F.add("p_pfd", "B3 差：模型隱含 − pro forma", "$B", only26("=«p_pfx»-«p_pfm»"),
          "只對照（Checks）；組成：模型年初現金＋已到位股權＋前九月自由現金流（模型）＋IPO 淨額 vs 原文 9-30 現金＋IPO 淨額", name="FND_ProFormaGap", name_cols="E")

    for sh in (R, C, K, F):
        sh.resolve()
    ctx.P["a10"] = dict(R=R, K=K, F=F)


def checks(ctx):
    R, K, F = (ctx.P["a10"][k] for k in "RKF")
    I = ctx.D.I
    d = lambda sh, key, col="D": f"={sh.name}!${col}${sh.rows[key]}"  # noqa: E731
    return [
        ("p_d2", "v0.2.0 D2：2027 營收總額 ÷ 2026Q4 年化 低於門檻的差（倍；> 0 立 WARN）", f"=MAX({I('d2_warn_ratio')}-REV_D2Ratio,0)", "0", "warn",
         "工作單 D2：< 1.0 立 WARN（模型 2027 年平均低於 2026 年底年化）；比值見 Revenue 第十四節"),
        ("p_q4", "v0.2.0 D1：2026 Q4 營收（$B）", "=REV_Q4_2026", None, "info", "開關 1：20.25；0：16.956"),
        ("p_totc", "v0.2.0 C1：三段＋短期 − 承諾總額 900.9（$B）", d(K, "p_totc"), None, "info", "原文捨入"),
        ("p_c5", "v0.2.0 C5：模型 2026Q4–2030 合約付款（不含無固定金額）÷ 2030 年前發生 451.6 − 1", "=COST_ProsBy2030Ratio", None, "info", "只對照"),
        ("p_sbcc", "v0.2.0 E1：2026 股權報酬中的 RSU 修改 − 開關 × 14.2（$B）", f"={K.name}!$E${K.rows['p_sbcipo']}-{I('p_sbc_ipo_switch')}*{ctx.D.S('pros_sbc_pf_mod')}",
         0, "tol", "只在 2026 欄；不進每人基準"),
        ("p_sbca", "v0.2.0 E2：攤銷時程 2026Q4–2028 合計 − 24.1（$B）", f"=SUM({K.name}!$E${K.rows['p_sbca']}:$G${K.rows['p_sbca']})-{ctx.D.S('pros_sbc_unrec')}", 0, "tol", ""),
        ("p_ppc", "v0.2.0 B1：私募逐家合計 − 26.0（$B）", d(F, "p_ppc"), 0, "tol", ""),
        ("p_ipo", "v0.2.0 B1：進基準的 IPO＋私募淨額（$B）", "=FND_IPONet", None, "info", "74 時 80.8"),
        ("p_liqc", "v0.2.0 B3：流動性 146.6 −（現金與短投＋循環額度可用）（$B）", d(F, "p_liqc", "E"), None, "info", "原文捨入"),
        ("p_pf", "v0.2.0 B3：模型 2026 年底現金 − Q4 自由現金流 − pro forma 212.684（價格連動）（$B）", "=FND_ProFormaGap", None, "info", "只對照"),
    ]
