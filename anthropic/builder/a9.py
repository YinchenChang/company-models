"""Anthropic v0.1.8：S-1/A No.1 滾動（工作單 20261102 v0.1.8 R4、R5a–R5f；原文 docs/reports/20261103_S1A原文摘錄.md）。Cost 第二十六節、Funding 第十五節，皆附加在頁尾，既有列編號不變。

- R4 SpaceXAI S-1 路徑（開關 s1f_spacex_fee_switch；基準 0＝既有月費 1.25 × 36）：2026-05～11 交付期間線性爬升（第 k 月＝滿載月費 × k ÷（爬升月數＋1）〔Assumed〕）、
  2026-12 起每月 3.1、累計以上限 120.5 截頂、至 2029；Cost 第二節 SpaceX 付款與第十一節合約總額在 a3 以開關選用本節列。
- R5a 2026 Q3 實績（開關 s1a_q3_switch）：營收 16.956（Revenue 營收基準，a2）；W4 下半年非算力＝Q3（營收 − 算力 − 營業損益 − 股權報酬 − 非現金捐贈）× 2（a7）；股權報酬代理下半年＝（1.070＋1.960）× 2（a7）。
- R5b S-1/A 承諾（開關 s1a_commit_switch）：W1(b) 三段＝表內 320.0＋未起租 451.2＋9-30 後 100.2＝871.4（a7 s1u_sum 選用本節 s1a_seg）；Google 合約 121.3（a3）。
- R5c 未歸屬餘額付款期（開關 s1a_unalloc_switch）：表內依 S-1/A 逐年表（之後 148.4 依 2031–2036 平均）、未起租＋9-30 後依起租年起加權租期 10 年平均，三段金額加權合成年度比例（a7 pay_s1u 乘本節 s1a_sh）。
- R5d 附條件 213.8（開關 s1a_cond_deduct_switch；基準 0＝計入）：扣除讀法時自三段合計扣除（a7 s1u_raw 減本節 s1a_cdd）。
- R5e／R5f Funding 第十五節：9-30 現金對照與差額拆解（只對照）；Q3 無新股權 → 已到位融資維持（對帳 S-1/A 前九月特別股收入）。
Python 只產生結構；數值一律引用 SRC_ANT／Inputs／既有列（E6）。
"""
from __future__ import annotations

from a2 import Sheet


def build(ctx):
    D = ctx.D
    S, I = D.S, D.I
    C, K, R = Sheet.all["C"], Sheet.all["K"], Sheet.all["R"]
    F = ctx.P["a4"]["F"]
    months, k2, m9 = I("months_per_year"), I("arith_k2"), I("m_9m")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731
    zero = "(«C:vr_unit@$»-«C:vr_unit@$»)"
    one = "«C:vr_unit@$»"

    K.section("二十六、v0.1.8 S-1/A No.1 滾動（工作單 20261102 v0.1.8；2026-11-02 申報）：R5a Q3 實績、R4 SpaceXAI S-1 路徑、R5b 承諾總額與付款表、R5c 未歸屬餘額付款期、R5d 附條件 213.8；列放在本節，使既有列編號不變")
    # ── R5a
    K.add("s1a_rev", "R5a S-1/A：2026 Q3 營收", "$B", only26(f"={S('rev_s1a_2026q3')}"), "SRC_ANT_488（Verified）；舊：SRC_ANT_384 報導 16.8")
    K.add("s1a_revd", "R5a Q3 營收：S-1/A ÷ 報導 − 1", "比例", only26(f"=«s1a_rev»/{S('rev_2026_q3')}-{one}"), "SRC_ANT_488 ÷ 384")
    K.add("s1a_ci", "R5a S-1/A：2026 Q3 算力與基礎設施費用", "$B", only26(f"={S('ci_s1a_2026q3')}"), "SRC_ANT_489")
    K.add("s1a_oi", "R5a S-1/A：2026 Q3 營業損益", "$B", only26(f"={S('oi_s1a_2026q3')}"), "SRC_ANT_490")
    K.add("s1a_sbc", "R5a S-1/A：2026 Q3 股權報酬", "$B", only26(f"={S('sbc_s1a_2026q3')}"), "SRC_ANT_491")
    K.add("s1a_chr", "R5a S-1/A：2026 Q3 非現金捐贈配對", "$B", only26(f"={S('charity_s1a_2026q3')}"), "SRC_ANT_492（10 月以 Class N 股份結清；一次性性質）")
    K.add("s1a_adj", "R5a S-1/A：2026 Q3 調整後營業損益", "$B", only26(f"={S('adjoi_s1a_2026q3')}"), "SRC_ANT_493")
    K.add("s1a_adjc", "R5a 檢查：調整後營業損益 −（營業損益＋股權報酬＋非現金捐贈）", "$B", only26("=«s1a_adj»-(«s1a_oi»+«s1a_sbc»+«s1a_chr»)"), "應為 0（Checks）")
    K.add("s1a_m", "R5a 調整後營業利益率：推算＝調整後營業損益 ÷ 營收（S-1/A 原文 15%，SRC_ANT_494；取代報導 11%，SRC_ANT_385）", "比例", only26("=«s1a_adj»/«s1a_rev»"), "")
    K.add("s1a_q3nc", "R5a 2026 Q3 非算力費用（含通路平台費；不含股權報酬與非現金捐贈）＝營收 − 算力 − 營業損益 − 股權報酬 − 非現金捐贈", "$B",
          only26("=«s1a_rev»-«s1a_ci»-«s1a_oi»-«s1a_sbc»-«s1a_chr»"), "s1a_q3_switch＝1 時進 W4 下半年（第二十五節 s1n_h2，Q4＝Q3）", name="COST_S1aQ3NonComp", name_cols="E")
    K.add("s1a_q3nco", "R5a 對照：v0.1.7.1 的 Q3 隱含非算力（第十九節；報導營收 16.8 × 報導利率 11% 推）", "$B", only26("=«u5_nc»"), "")
    K.add("s1a_q3ncd", "R5a 差：Q3 實際 − Q3 隱含", "$B", only26("=«s1a_q3nc»-«s1a_q3nco»"), "")
    K.add("s1a_cimp", "R5a 對照：模型 Q3 隱含算力＝Q3 報導營收 ÷ 每 $1 算力的營收 1.70（SRC Derived）", "$B", only26(f"={S('cmp_s1_spend_2026_q3_derived')}"), "SRC_ANT_384 ÷ 387")
    K.add("s1a_cid", "R5a 差：Q3 算力 S-1/A 實際 − 模型 Q3 隱含", "$B", only26("=«s1a_ci»-«s1a_cimp»"), "只對照，不校準", name="COST_S1aQ3ComputeGap", name_cols="E")
    K.add("s1a_sbcq", "R5a 股權報酬代理每季：Q3 實際（股權報酬＋非現金捐贈）", "$B", only26("=«s1a_sbc»+«s1a_chr»"), "s1a_q3_switch＝1 時 × 下半年季數進第二十五節 s1s_fy")
    K.add("s1a_sbcqo", "R5a 對照：v0.1.7.1 每季代理（Q2 調整後營業利益率 × Q2 營收 − Q2 營業損益）", "$B", only26("=«s1s_q2»"), "")

    # ── 獨立核對後（2026-11-03）：下半年非現金捐贈不進每人股權報酬基準；2026 全成本另計（判斷題）
    sws = I("s1f_sbc_proxy_switch")
    hq = f"{I('quarters_per_year')}*{I('h2_months')}/{months}"
    sq3a, chq = I("s1a_q3_switch"), S("charity_s1a_2026q3")
    K.add("s1s_chx", "v0.1.8 核對後：2026 下半年非現金捐贈（S-1/A Q3 1.960 × 下半年季數；一次性）——自每人股權報酬基準扣除，不沿用到 2027 起", "$B",
          lambda i, c: f"={sq3a}*{chq}*{hq}" if i == 1 else f"={zero}", "SRC_ANT_492；s1a_q3_switch＝0 時為 0")
    K.add("s1s_chc", "v0.1.8 核對後：2026 全成本另計的非現金捐贈（判斷題：Q4＝Q3 也計一次〔s1a_chr_q4_switch＝1，保守〕；0＝只計 Q3 一次）", "$B",
          lambda i, c: (f"={sws}*{sq3a}*{chq}*({one}+{I('s1a_chr_q4_switch')}*({hq}-{one}))" if i == 1 else f"={zero}"),
          "加在股權報酬（單列）2026；2027 起 0", name="COST_S1aCharityInCost")

    # ── R4 SpaceXAI S-1 路徑
    fF, cap = S("cmt_s1a_spacex_monthly"), S("cmt_s1f_spacex")
    ss, se = I("spacex_start"), I("spacex_end")
    K.add("s1a_spx_f", "R4 SpaceXAI 滿載月費（S-1/A；預期 2026-12）", "$B/月", only25(f"={fF}"), "SRC_ANT_515")
    K.add("s1a_spx_cap", "R4 SpaceXAI 合約上限（至 2029）", "$B", only25(f"={cap}"), "SRC_ANT_453（S-1；S-1/A 同）")
    K.add("s1a_spx_unc", "R4 S-1 路徑未截頂年付款＝滿載月費 ×（起始年：爬升月數 ÷ 2 ＋ 滿載月數；之後至結束年：12）", "$B",
          lambda i, c: f"=«s1a_spx_f@$»*(({c}$6={ss})*({I('spacex_ramp_months')}/{k2}+{I('spacex_full_m26')})+({c}$6>{ss})*({c}$6<={se})*{months})",
          "爬升第 k 月＝滿載 × k ÷（爬升月數＋1）〔Assumed，線性〕→ 爬升期合計＝滿載 × 爬升月數 ÷ 2")
    K.add("s1a_spx_cum", "R4 S-1 路徑未截頂累計", "$B", lambda i, c: "=«s1a_spx_unc»" if i == 0 else "=«s1a_spx_cum@p»+«s1a_spx_unc»", "")
    K.add("s1a_spx_pay", "R4 S-1 路徑年付款（累計以上限截頂）", "$B",
          lambda i, c: "=MIN(«s1a_spx_cum»,«s1a_spx_cap@$»)" if i == 0 else "=MIN(«s1a_spx_cum»,«s1a_spx_cap@$»)-MIN(«s1a_spx_cum@p»,«s1a_spx_cap@$»)",
          "s1f_spacex_fee_switch＝1 時為第二節 SpaceX 付款", name="COST_S1aSpaceXPath")
    K.add("s1a_spx_tot", "R4 S-1 路徑合約總額（2025–2030 合計；＝上限 120.5）", "$B", only25("=SUM(«s1a_spx_pay@R»)"), "第十一節 SpaceX 總額（開關 1 時）",
          name="COST_S1aSpaceXTot", name_cols="D")
    K.add("s1a_spx_d", "R4 對照：S-1 路徑 − 現行第二節 SpaceX 付款", "$B", lambda i, c: "=«s1a_spx_pay»-«pay_spacex»", "開關 0 時＝S-1 路徑 − 既有月費路徑")

    # ── R5b 承諾總額
    K.add("s1a_tbl", "R5b S-1/A：表內合約義務合計（2026-09-30；營運租賃＋融資租賃＋採購承諾）", "$B", only25(f"={S('cmt_s1a_table')}"), "SRC_ANT_496")
    K.add("s1a_lnc", "R5b S-1/A：已簽未起租租賃（表外）", "$B", only25(f"={S('cmt_s1a_leases_nc')}"), "SRC_ANT_504")
    K.add("s1a_post", "R5b S-1/A：9-30 後新簽或修訂（v0.2.0 C1：開關 p_cmt_post_switch＝1 改讀初步招股說明書 105.7）", "$B", only25(f"={I('p_cmt_post_switch')}*{S('cmt_pros_post0930')}+(1-{I('p_cmt_post_switch')})*{S('cmt_s1a_post0930')}"), "SRC_ANT_572；舊：SRC_ANT_506（100.2）")
    K.add("s1a_seg", "R5b S-1/A 三段合計（進基準）＝表內＋未起租＋9-30 後", "$B", only25("=«s1a_tbl@$»+«s1a_lnc@$»+«s1a_post@$»"),
          "s1a_commit_switch＝1 時為第二十五節 W1(b) 三段合計；短期 17.4、期後短期 6.6 照 N2 不進基準", name="COST_S1aCommit3", name_cols="D")
    K.add("s1a_short", "R5b 對照：短期承諾（一年內可無罰取消）＋9-30 後短期（不進基準）", "$B", only25(f"={S('cmt_s1a_short')}+{S('cmt_s1a_post0930_short')}"), "SRC_ANT_503＋507")
    K.add("s1a_tot", "R5b S-1/A：承諾總額", "$B", only25(f"={S('cmt_s1a_total')}"), "SRC_ANT_495")
    K.add("s1a_totc", "R5b 檢查：三段＋短期 − 總額（原文捨入）", "$B", only25("=«s1a_seg@$»+«s1a_short@$»-«s1a_tot@$»"), "約 −0.007")
    K.add("s1a_old", "R5b 對照：v0.1.7.1 三段（S-1 6-30：未起租 321.4＋採購 265.6＋6-30 後 245）", "$B", only25("=«s1u_l@$»+«s1u_p@$»+«s1u_n@$»"), "SRC_ANT_445–447")
    K.add("s1a_oldd", "R5b 差：S-1/A 三段 − S-1 三段", "$B", only25("=«s1a_seg@$»-«s1a_old@$»"), "")
    K.add("s1a_goog", "R5b Google 合約：S-1/A（D）／S-1（E）／差（F）", "$B",
          [f"={S('cmt_s1a_google')}", f"={S('cmt_s1f_google')}", "=«s1a_goog@D»-«s1a_goog@E»", None, None, None], "SRC_ANT_513／450；欄位不是年度")
    K.add("s1a_n12", "R5b 對照：S-1/A 未來 12 個月預計支付（D；其中附條件 E）vs 模型 2027 合約實付（F；期間差一季）", "$B",
          [f"={S('cmt_s1a_next12m')}", f"={S('cmt_s1a_next12m_cond')}", "=«pay@F»", None, None, None], "SRC_ANT_509、510；欄位不是年度")
    K.add("s1a_conv", "R5b 對照：可再發行可轉債上限 77.0（Broadcom 42＋Google 20＋Amazon 15；照 N3 不計入已到位融資）", "$B", only25(f"={S('debt_s1a_conv_reissue')}"), "SRC_ANT_512")
    K.add("s1a_nofx", "R5f 對照：無固定金額的其他算力安排估計付款（低 D／高 E；不在總額內，含 Google 額外 TPU 與 AMD；不進基準）", "$B",
          [f"={S('cmt_s1a_nofixed')}_Lo", f"={S('cmt_s1a_nofixed')}_Hi", None, None, None, None], "SRC_ANT_511；欄位不是年度")
    # ── R5d 附條件 213.8
    K.add("s1a_cond", "R5d S-1/A：附條件或可終止承諾（v0.2.0 C3：開關 p_cond_cmt_switch＝1 改讀初步招股說明書 186.2）", "$B", only25(f"={I('p_cond_cmt_switch')}*{S('cmt_pros_cond')}+(1-{I('p_cond_cmt_switch')})*{S('cmt_s1a_conditional')}"), "SRC_ANT_569；舊：SRC_ANT_508（213.8）")
    K.add("s1a_cdd", "R5d 自三段合計扣除的附條件額（開關 × 213.8；基準 0＝計入）", "$B", only25(f"={I('s1a_cond_deduct_switch')}*«s1a_cond@$»"),
          "第二十五節 s1u_raw 減本列", name="COST_S1aCondDeduct", name_cols="D")

    # ── R5c 未歸屬餘額付款期（S-1/A 結構合成）
    tb = {2026: "cmt_s1a_tbl_2026q4", 2027: "cmt_s1a_tbl_2027", 2028: "cmt_s1a_tbl_2028", 2029: "cmt_s1a_tbl_2029", 2030: "cmt_s1a_tbl_2030"}
    K.add("s1a_tby", "R5c S-1/A 表內義務逐年（2026＝Q4）", "$B", lambda i, c: (f"={zero}" if i == 0 else f"={S(tb[2025 + i])}"), "SRC_ANT_497–501")
    K.add("s1a_aft", "R5c S-1/A 表內義務 2030 年之後", "$B", only25(f"={S('cmt_s1a_tbl_after')}"), "SRC_ANT_502")
    K.add("s1a_aftn", "R5c 之後部分攤提年數＝結束 − 起始 ＋ 1（Inputs 2031–2036）", "年", only25(f"=({I('s1a_after_end')}-{I('s1a_after_start')}+1)"), "")
    K.add("s1a_tbyc", "R5c 檢查：表內逐年合計＋之後 − 表內合計（原文捨入）", "$B", only25("=SUM(«s1a_tby@R»)+«s1a_aft@$»-«s1a_tbl@$»"), "")
    ls, wt = I("s1a_lease_start"), S("cmt_s1a_lease_wterm")
    K.add("s1a_lpy", "R5c 未起租＋9-30 後新簽的年付款＝（451.2＋100.2）÷ 加權租期 10 年", "$B", only25(f"=(«s1a_lnc@$»+«s1a_post@$»)/{wt}"), "SRC_ANT_505；起租年 Inputs s1a_lease_start（2027）")
    K.add("s1a_ly", "R5c 未起租＋9-30 後新簽逐年（起租年起 10 年）", "$B", lambda i, c: f"=«s1a_lpy@$»*({c}$6>={ls})*({c}$6<{ls}+{wt})", "")
    K.add("s1a_segy", "R5c 三段逐年付款（S-1/A 結構）＝表內逐年＋未起租與 9-30 後逐年", "$B", lambda i, c: "=«s1a_tby»+«s1a_ly»", "")
    K.add("s1a_sh", "R5c 未歸屬餘額年度比例＝三段逐年 ÷ 三段合計", "比例", lambda i, c: "=«s1a_segy»/«s1a_seg@$»",
          "s1a_unalloc_switch＝1 時：第二十五節未歸屬餘額付款＝餘額 × 本列", name="COST_S1aUnallocShare")
    K.add("s1a_shs", "R5c 2025–2030 比例合計（其餘在 2031 年後）", "比例", only25("=SUM(«s1a_sh@R»)"), "")
    K.add("s1a_shc", "R5c 檢查：全期比例合計 − 1（表內逐年＋之後＋租約 10 年）", "比例",
          only25(f"=(SUM(«s1a_tby@R»)+«s1a_aft@$»+«s1a_lpy@$»*{wt})/«s1a_seg@$»-{one}"), "應約為 0（表內原文捨入）")
    # 並列讀法（獨立核對建議；基準 0）：殘差法——未歸屬餘額逐年＝MAX(0, S-1/A 三段逐年 − 模型逐家〔承諾表內〕合約逐年付款)，累計以餘額為上限
    from a7 import IN_TABLE
    K.add("s1a_rm", "R5c 並列：模型逐家合約（S-1/A 承諾表內，不含 AMD、既有雲端、未歸屬餘額）逐年付款", "$B", lambda i, c: "=" + "+".join(f"«pay_{k}»" for k in IN_TABLE), "")
    K.add("s1a_rr", "R5c 並列：殘差＝MAX(0, 三段逐年 − 逐家逐年)", "$B", lambda i, c: f"=MAX({zero},«s1a_segy»-«s1a_rm»)", "")
    K.add("s1a_rc", "R5c 並列：殘差累計", "$B", lambda i, c: "=«s1a_rr»" if i == 0 else "=«s1a_rc@p»+«s1a_rr»", "")
    K.add("s1a_rs", "R5c 並列：殘差法未歸屬餘額付款（累計以未歸屬餘額為上限；s1a_resid_switch＝1 時進第二十五節）", "$B",
          lambda i, c: "=MIN(«s1a_rc»,«s1u_tot@$»)" if i == 0 else "=MIN(«s1a_rc»,«s1u_tot@$»)-MIN(«s1a_rc@p»,«s1u_tot@$»)",
          "基準不用：上限型合約（SpaceXAI、Nscale）模型付款可能超過最低承諾，殘差會低估未歸屬付款", name="COST_S1aUnallocResid")
    s_, e_ = I("s1u_start"), I("s1u_end")
    K.add("s1a_sho", "R5c 對照：v0.1.7 平均攤提的年度比例（s1u_start–s1u_end）", "比例", lambda i, c: f"=({c}$6>={s_})*({c}$6<={e_})/«s1u_ny@$»", "")
    K.add("s1a_pm", "R5c 新 Checks：模型 2027–2030 合約實付合計（不含 AMD，不在 S-1/A 承諾表內）", "$B", only25("=SUM(«pay@F»:«pay@I»)-SUM(«pay_amd@F»:«pay_amd@I»)"), "")
    K.add("s1a_ps", "R5c 新 Checks：S-1/A 推算 2027–2030＝表內逐年＋未起租與 9-30 後（起租 2027、10 年）", "$B",
          only25("=SUM(«s1a_tby@F»:«s1a_tby@I»)+SUM(«s1a_ly@F»:«s1a_ly@I»)"), "推算口徑（只對照）")
    K.add("s1a_pr", "R5c 模型 ÷ S-1/A 推算 − 1（2027–2030）", "比例", only25("=«s1a_pm@$»/«s1a_ps@$»-«C:vr_unit@$»"), "只對照", name="COST_S1aPay2730Gap", name_cols="D")

    # ═════════════ Funding 第十五節 ═════════════
    F.section("十五、v0.1.8 R5e／R5f：S-1/A 2026-09-30 現金對照（第十三節 6-30 對照的 9-30 版；只對照）與 Q3 股權（無新股權；已到位融資維持）")
    cash25 = S("cash_2025_12_31")
    F.add("c930_s1", "S-1/A：2026-09-30 現金及約當現金＋短期投資（95.6＋36.3）", "$B", only26(f"={S('cash_s1a_2026_09_30')}"), "SRC_ANT_516（Verified）")
    F.add("c930_rc", "S-1/A：2026-09-30 受限現金（另計；不在上列內）", "$B", only26(f"={S('cash_s1a_restricted_0930')}"), "SRC_ANT_517")
    F.add("c930_mod", "模型 2026-09-30 近似＝2025 年底現金 ＋ 2026 已到位融資 ＋ 2026 自由現金流 × 9 ÷ 12", "$B",
          only26(f"={cash25}+«committed»+«fcf»*{m9}/{months}"), "年內平均近似；模型自由現金流為現金口徑")
    F.add("c930_s1b", "對照：以 S-1/A 前九月現金流推算＝2025 年底現金 ＋ 特別股收入 ＋ 營業現金流 − 資本支出 − 租賃資產預付", "$B",
          only26(f"={cash25}+{S('fund_s1a_pref_2026_9m')}+{S('ocf_s1a_2026_9m')}-{S('capex_s1a_2026_9m')}-{S('leasepre_s1a_2026_9m')}"), "SRC_ANT_518–521")
    F.add("c930_gap", "差：S-1/A 9-30 現金 − 模型近似", "$B", only26("=«c930_s1»-«c930_mod»"), "只對照，不校準", name="FND_Cash0930Gap", name_cols="E")
    F.add("c930_d_eq", "差額拆解 ①股權收入：S-1/A 前九月特別股收入 − 模型已到位融資", "$B", only26(f"={S('fund_s1a_pref_2026_9m')}-«committed»"), "SRC_ANT_521 − FND_Committed")
    F.add("c930_d_tm", "差額拆解 ②前九月自由現金流時點：S-1/A（營業現金流 − 資本支出 − 租賃資產預付）− 模型 2026 自由現金流 × 9 ÷ 12", "$B",
          only26(f"={S('ocf_s1a_2026_9m')}-{S('capex_s1a_2026_9m')}-{S('leasepre_s1a_2026_9m')}-«fcf»*{m9}/{months}"), "租賃資產預付 1.955 為 S-1/A 新科目（投資活動）")
    F.add("c930_d_rc", "差額拆解 ③受限現金：S-1/A 9-30 現金不含受限現金（負號）", "$B", only26("=-«c930_rc»"), "")
    F.add("c930_d_ot", "差額拆解 ④其餘＝差 − ① − ② − ③（短期投資與現金口徑、普通股發行、融資租賃本金、其他投資等）", "$B",
          only26("=«c930_gap»-«c930_d_eq»-«c930_d_tm»-«c930_d_rc»"), "只對照，不校準", name="FND_Cash0930Resid", name_cols="E")
    F.add("c930_q3o", "對照：S-1/A Q3 營業現金流（推算）＝前九月 − 上半年", "$B", only26(f"={S('ocf_s1a_2026_9m')}-{S('ocf_s1f_2026h1')}"), "SRC_ANT_518 − 442")
    F.add("c930_q3m", "對照：模型 2026 自由現金流 ÷ 季數（年內平均一季）", "$B", only26(f"=«fcf»/{I('quarters_per_year')}"), "")
    F.add("c930_eq", "R5f：模型 2026 已到位融資 − S-1/A 前九月特別股收入（Q3 無新股權）", "$B", only26(f"=«committed»-{S('fund_s1a_pref_2026_9m')}"),
          "|差| ≤ Inputs s1f_cmt_tol（0.5）為 OK（Checks）", name="FND_CommittedS1aGap", name_cols="E")

    for sh in (R, C, K, F):
        sh.resolve()
    ctx.P["a9"] = dict(K=K, F=F)


def checks(ctx):
    I = ctx.D.I
    K, F = ctx.P["a9"]["K"], ctx.P["a9"]["F"]
    d = lambda sh, key, col="D": f"={sh.name}!${col}${sh.rows[key]}"  # noqa: E731
    return [
        ("i1_q3rev", "v0.1.8 R5a：2026 Q3 營收 S-1/A ÷ 報導 − 1", d(K, "s1a_revd", "E"), None, "info", "16.956 vs 16.8"),
        ("i1_adj", "v0.1.8 R5a：Q3 調整後營業損益 −（營業損益＋股權報酬＋非現金捐贈）（$B）", d(K, "s1a_adjc", "E"), 0, "tol", "原文調節"),
        ("i1_q3nc", "v0.1.8 R5a：2026 Q3 非算力費用（不含股權報酬與非現金捐贈；$B）", "=COST_S1aQ3NonComp", None, "info", "取代 Q3 隱含 5.07"),
        ("i1_q3ci", "v0.1.8 R5a：Q3 算力 S-1/A − 模型 Q3 隱含（$B）", "=COST_S1aQ3ComputeGap", None, "info", "只對照"),
        ("i1_spx", "v0.1.8 R4：SpaceXAI S-1 路徑合計（$B；＝上限）", "=COST_S1aSpaceXTot", None, "info", "基準開關 0：不進基準"),
        ("i1_seg", "v0.1.8 R5b：S-1/A 三段合計（$B）", "=COST_S1aCommit3", None, "info", "320.0＋451.2＋100.2"),
        ("i1_tot", "v0.1.8 R5b：三段＋短期 − 承諾總額 895.4（$B）", d(K, "s1a_totc"), None, "info", "原文捨入"),
        ("i1_unal", "v0.1.8 R5b：未歸屬餘額（$B）", "=COST_S1Unalloc", None, "info", "871.4 − 模型逐家（不含 AMD）"),
        ("i1_shc", "v0.1.8 R5c：未歸屬餘額年度比例全期合計 − 1", d(K, "s1a_shc"), 0, "tol", "表內逐年合計＝表內合計 319.993（原文）"),
        ("i1_pay", "v0.1.8 R5c：模型 2027–2030 合約實付（不含 AMD）÷ S-1/A 表內＋租約推算 − 1", "=COST_S1aPay2730Gap", None, "info", "只對照"),
        ("i1_cond", "v0.1.8 R5d：自三段扣除的附條件額（$B；基準 0）", "=COST_S1aCondDeduct", None, "info", "扣除讀法見報告"),
        ("i1_c930", "v0.1.8 R5e：S-1/A 9-30 現金 − 模型近似（$B）", "=FND_Cash0930Gap", None, "info", "只對照"),
        ("i1_c930r", "v0.1.8 R5e：9-30 差額拆解的其餘項（$B）", "=FND_Cash0930Resid", None, "info", "只對照"),
        ("i1_sbcx", "v0.1.8 核對後：2026 每人股權報酬基準 × 人數 −（股權報酬代理 − 下半年非現金捐贈）（$B；應為 0＝捐贈不沿用到 2027 起）",
         f"={I('s1f_sbc_proxy_switch')}*({K.name}!$E${K.rows['sbc']}-{K.name}!$E${K.rows['s1s_chc']}-{K.name}!$E${K.rows['p_sbcipo']}-({K.name}!$E${K.rows['s1s_fy']}-{K.name}!$E${K.rows['s1s_chx']}))", 0, "tol", "只在股權報酬代理開關 1 時成立；Checks"),
        ("i1_eq", "v0.1.8 R5f：2026 已到位融資 − S-1/A 前九月特別股收入 110.432（$B）", "=FND_CommittedS1aGap", I("s1f_cmt_tol"), "warn", "Q3 無新股權；|差| ≤ 0.5 為 OK"),
    ]
