"""Anthropic v0.1.7：正式 S-1 全面套用＋自建資本支出改 GW 法＋快照同步（工作單 20261031 v0.1.7；Cost 第二十五節、Funding 第十二節，皆附加在頁尾，既有列編號不變）。

- W1(b) S-1 承諾未歸屬餘額：S-1 三段合計（6-30 尚未起租租賃＋6-30 固定或最低承諾＋6-30 後新簽與修訂）− 模型逐家合約總額（第十一節；AMD 不在 S-1 承諾表內，不扣）
  ＝未歸屬餘額；以通用合約價（VR200 族，INP contract_price × 持有比 1）換 GW，付款期間 Inputs s1u_start–s1u_end 平均攤提（不爬坡）。開關 s1f_commit_switch＝0 時為 0。
- W1(a) SpaceXAI 上限讀法係數（基準 1＝沿用月費；s1f_spacex_cap_switch＝1 時按正式 S-1 上限 120.5 ÷ 月費總額等比例放大）。
- W1(c) Google 額外 TPU 125–150（不在承諾表；情境，基準 0）：中點 ÷ 期間平均攤提，GW＝付款 ÷ TPU v7 合約價。
- W2 自建資本支出 GW 法的對照列（S-1 上半年資本支出年化 vs 模型；舊法 Fluidstack 均攤值列示）。
- W3 2025 校準：正式 S-1 vs 草擬版（列示）。
- W4 2026 非算力成本：正式 S-1 上半年與 Q2 的「營收 − 算力與基礎設施費用 − 營業損益」＝非算力營業費用（含雲端平台抽成與股權報酬）；
  全年＝上半年 ＋ Q2 × 2（下半年 Q2 持平，同 U1 的 Q4＝Q3 慣例）；扣模型 2026 平台抽成（REV_PartnerShare）與股權報酬（COST_SBC）＝不含股權報酬的非算力成本，
  Cost 第五節「每人年成本」2026 欄改以此值 ÷ 人數（開關 s1f_d22_switch）。
- W6 Amazon 2026 實際投入（10-Q：Series G 5＋Series H 5＝10.0）vs 模型已到位融資中 Amazon 部分（Funding 第十二節；只對照）。
最後一次解析 Compute、Cost、Funding 的列鍵記號（a3／a4 的公式引用本節新列）。Python 只產生結構；數值一律引用 SRC_ANT／Inputs／既有列（E6）。
"""
from __future__ import annotations

from a2 import FMT, YEARS, Sheet
from a3 import CONTRACTS

FMT.setdefault("年", "0")

IN_TABLE = [c[0] for c in CONTRACTS if c[0] != "amd"]          # S-1 承諾表內的模型合約（AMD「不在承諾表內」，SRC_ANT_457）


def build(ctx):
    D = ctx.D
    S, I = D.S, D.I
    C, K, R = Sheet.all["C"], Sheet.all["K"], Sheet.all["R"]
    F = ctx.P["a4"]["F"]
    months, h2, qpy, k2, m2b = I("months_per_year"), I("h2_months"), I("quarters_per_year"), I("arith_k2"), I("m_to_b")
    sw = I("s1f_commit_switch")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731
    zero = "(«C:vr_unit@$»-«C:vr_unit@$»)"
    one = "«C:vr_unit@$»"

    # ═════════════ Cost 第二十五節 ═════════════
    K.section("二十五、v0.1.7 正式 S-1 全面套用（工作單 20261031 v0.1.7）：W1 承諾進基準（未歸屬餘額、SpaceXAI 上限讀法、Google 額外 TPU 情境）、W2 自建資本支出對照、W3 2025 校準、W4 2026 非算力成本校準；列放在本節，使既有列編號不變")
    # W1(b)
    K.add("s1u_l", "W1(b) S-1：截至 6-30 尚未起租的租賃（未折現）", "$B", only25(f"={S('cmt_s1f_leases_nc_0630')}"), "SRC_ANT_445（二手轉載）")
    K.add("s1u_p", "W1(b) S-1：截至 6-30 固定或最低採購承諾", "$B", only25(f"={S('cmt_s1f_purchase_0630')}"), "SRC_ANT_446")
    K.add("s1u_n", "W1(b) S-1：6-30 後新簽與修訂合約的未來最低付款", "$B", only25(f"={S('cmt_s1f_post0630')}"), "SRC_ANT_447；與 Note 16 的 417（SRC_ANT_448）重疊，不相加")
    K.add("s1u_sum", "W1(b) S-1 三段合計", "$B", only25("=«s1u_l@$»+«s1u_p@$»+«s1u_n@$»"), "", name="COST_S1fCommit3", name_cols="D")
    K.add("s1u_mod", "W1(b) 模型逐家合約總額（第十一節；W1(a) 後；不含 AMD〔不在 S-1 承諾表內〕）", "$B",
          only25("=" + "+".join(f"«tot_{k}@$»" for k in IN_TABLE)), "Google、Broadcom、Amazon、Microsoft、SpaceX、Nscale、Lambda、Volta、Akamai", name="COST_S1fCommitModel", name_cols="D")
    K.add("s1u_raw", "W1(b) 差額＝S-1 三段合計 − 模型逐家合約總額", "$B", only25("=«s1u_sum@$»-«s1u_mod@$»"), "負＝模型逐家已超過 S-1 合計（未歸屬餘額取 0）")
    K.add("s1u_tot", "W1(b) S-1 承諾未歸屬餘額（進基準；開關 × MAX(0, 差額)）", "$B", only25(f"={sw}*MAX({zero},«s1u_raw@$»)"),
          "工作單 W1(b)；s1f_commit_switch＝0 時為 0", name="COST_S1Unalloc", name_cols="D")
    s_, e_ = I("s1u_start"), I("s1u_end")
    K.add("s1u_ny", "W1(b) 付款年數＝結束 − 起始 ＋ 1", "年", only25(f"=({e_}-{s_}+1)"), "Inputs s1u_start／s1u_end（預設 2027–2033；快 2027–2031、慢 2028–2036）")
    K.add("pay_s1u", "合約實付：S-1 承諾未歸屬餘額（W1(b)；期間內平均）", "$B", lambda i, c: f"=«s1u_tot@$»/«s1u_ny@$»*({c}$6>={s_})*({c}$6<={e_})",
          "GW＝付款 ÷ 通用合約價（Compute 第十六節，VR200 族）", name="COST_Pay_S1Unalloc")
    K.add("s1u_after", "W1(b) 未歸屬餘額中 2030 年後應付", "$B", only25("=«s1u_tot@$»-SUM(«pay_s1u@R»)"), "")
    K.add("s1u_all", "W1 模型合約總額合計（全部新合約〔含 AMD〕＋未歸屬餘額）", "$B", only25("=«tot_all@$»+«s1u_tot@$»"), "對照 S-1 三段合計＋AMD", name="COST_CommitAllS1", name_cols="D")
    # W1(a) SpaceXAI 上限讀法
    sp = dict(CONTRACTS[[c[0] for c in CONTRACTS].index("spacex")][4])
    sp = {k: D.fill(v) for k, v in sp.items()}
    K.add("s1_spx_mt", "W1(a) SpaceX 月費總額（月費 × 總月數；既有規則）", "$B", only25(f"={sp['F']}*({sp['m1']}+{sp['m2']}+({sp['e']}-({sp['s']}+1))*{months})"),
          "SRC_ANT 月費 × Inputs 月數（同第十一節舊式）")
    K.add("s1_spx_fac", "W1(a) SpaceXAI 上限讀法係數（基準 1；s1f_spacex_cap_switch＝1 且承諾開關 1 時＝正式 S-1 上限 ÷ 月費總額）", "倍",
          only25(f"={one}+{I('s1f_spacex_cap_switch')}*{sw}*({S('cmt_s1f_spacex')}/«s1_spx_mt@$»-{one})"),
          "工作單預設：上限型沿用既有規則（SpaceX＝月費為基準，上限只作區間）；第二節付款與第十一節總額 × 本係數", name="COST_S1fSpaceXFactor", name_cols="D")
    # W1(a) AMD 讀法（基準 1＝維持 GW 型；s1f_amd_switch＝1 時合約總額＝正式 S-1 下限 20，時程與 GW 等比例縮放）
    am = dict(CONTRACTS[[c[0] for c in CONTRACTS].index("amd")][4])
    am = {k: D.fill(v) for k, v in am.items()}
    r_ = I("ramp_years")
    K.add("s1_amd_raw", "W1(a) AMD 模型合約總額（既有 GW 型；未乘讀法係數）", "$B",
          only25(f"=«C:gw_unit_amd@$»*«C:price_mi455@D»*(({am['term']})+({am['term']})-{r_}+1)/{k2}"), "同第十一節舊式")
    K.add("s1_amd_fac", "W1(a) AMD 讀法係數（基準 1；s1f_amd_switch＝1 且承諾開關 1 時＝正式 S-1 下限 20 ÷ 模型總額）", "倍",
          only25(f"={one}+{I('s1f_amd_switch')}*{sw}*({S('cmt_s1f_amd')}/«s1_amd_raw@$»-{one})"),
          "S-1：「若全數購買初始容量，超過 $20B」、不在承諾表內 → 下限；基準維持模型 GW 型（chat 端核對後確認）；Compute AMD 供給與第十一節總額 × 本係數", name="COST_S1fAMDFactor", name_cols="D")
    # W1(c) Google 額外 TPU（情境）
    gs, ge = I("gadd_start"), I("gadd_end")
    K.add("gadd_tot", "W1(c) Google 額外 TPU 容量（情境；開關 × SRC 低高中點）", "$B",
          only25(f"={I('s1f_gaddl_switch')}*{sw}*({S('cmt_s1f_google_addl')}_Lo+{S('cmt_s1f_google_addl')}_Hi)/{k2}"),
          "SRC_ANT_458（125–150，不在承諾表內）；基準開關 0", name="COST_GoogleAddl", name_cols="D")
    K.add("pay_gadd", "合約實付：Google 額外 TPU（W1(c) 情境；期間內平均）", "$B",
          lambda i, c: f"=«gadd_tot@$»/({ge}-{gs}+1)*({c}$6>={gs})*({c}$6<={ge})", "GW＝付款 ÷ TPU v7 合約價（Compute 第十六節）", name="COST_Pay_GoogleAddl")
    # W1 對照：逐家（正式 S-1 vs 模型第十一節）
    K.add("s1g_gap", "W1 對照：模型合約總額合計（含未歸屬餘額，不含 AMD）÷ S-1 三段合計 − 1", "比例", only25("=(«s1u_mod@$»+«s1u_tot@$»-«s1u_sum@$»)/«s1u_sum@$»"),
          "承諾開關 1 時應為 0（模型逐家不超過 S-1 合計時）", name="COST_S1fCommitGapAfter", name_cols="D")

    # W2 對照
    K.add("s2_s1capex", "W2 對照：S-1 2026 上半年資本支出年化（× 12 ÷ 上半年月數）", "$B", only26(f"={S('capex_s1f_2026h1')}*{months}/{h2}"), "SRC_ANT_443（二手轉載）")
    K.add("s2_model", "W2 對照：模型 2026 自建資本支出（現行算法）", "$B", lambda i, c: "=«own»", "COST_OwnedCapex")
    K.add("s2_diff", "W2 對照：模型 − S-1 年化（2026）", "$B", only26("=«s2_model»-«s2_s1capex»"), "S-1 資本支出含辦公室等非算力項目；模型只計自建算力")
    fs_, fy = I("fluid_start"), I("fluid_years")
    K.add("s2_old", "W2 對照：舊法（Fluidstack $50B ÷ 均攤年數；own_capex_method＝0）", "$B",
          lambda i, c: f"={S('cmp_fluidstack_50b')}/{fy}*({c}$6>={fs_})*({c}$6<{fs_}+{fy})", "SRC_ANT_212；Fluidstack 站點視為 Fluidstack 擁有、經承諾租回（已在 S-1 承諾內）→ 新法不另計")

    # W3 對照
    K.add("s3_rev", "W3 2025 校準營收：正式 S-1（D）／草擬版（E）／差（F）", "$B",
          [f"={S('rev_s1f_fy2025')}", f"={S('rev_2025_recognized')}", "=«s3_rev@D»-«s3_rev@E»", None, None, None], "SRC_ANT_412／SRC_ANT_001；欄位不是年度")
    K.add("s3_ci", "W3 2025 算力與基礎設施費用：正式 S-1（D）／草擬版（E）／差（F）", "$B",
          [f"={S('ci_s1f_fy2025')}", f"={S('cmp_spend_2025_actual_prospectus')}", "=«s3_ci@D»-«s3_ci@E»", None, None, None], "SRC_ANT_424／SRC_ANT_252；欄位不是年度")
    K.add("s3_eta", "W3 η（2025；現行開關下）", "倍", only25("=«C:eta25@$»"), "D8；開關 0 的值見報告（情境重算）")

    # W4 2026 非算力成本（v0.1.7 獨立核對後：下半年與營收同期〔Q3 隱含、Q4＝Q3〕；股權報酬以 S-1 調整數代理；兩者各有開關）
    swq, sws = I("s1f_d22_q3_switch"), I("s1f_sbc_proxy_switch")
    hq = f"{qpy}*{h2}/{months}"            # 下半年季數
    K.add("s1n_h1rev", "W4 S-1：2026 上半年營收", "$B", only26(f"={S('rev_s1f_2026h1')}"), "SRC_ANT_414（Verified）")
    K.add("s1n_h1ci", "W4 S-1：2026 上半年算力與基礎設施費用", "$B", only26(f"={S('ci_s1f_2026h1')}"), "SRC_ANT_426（Verified）")
    K.add("s1n_h1oi", "W4 S-1：2026 上半年營業損益", "$B", only26(f"={S('oi_s1f_2026h1')}"), "SRC_ANT_435（Verified）")
    K.add("s1n_h1nc", "W4 2026 上半年非算力營業費用＝營收 − 算力與基礎設施 − 營業損益（含雲端平台抽成與股權報酬）", "$B", only26("=«s1n_h1rev»-«s1n_h1ci»-«s1n_h1oi»"),
          "口徑：S-1 損益表以性質別列示；營收為總額（雲端市集抽成列在費用）")
    K.add("s1n_q2nc", "W4 2026 Q2 非算力營業費用＝Q2 營收 − Q2 算力與基礎設施 − Q2 營業損益（含抽成與股權報酬）", "$B",
          only26(f"={S('rev_s1f_2026q2')}-{S('ci_s1f_2026q2')}-{S('oi_s1f_2026q2')}"), "SRC_ANT_420、432、441（二手轉載；Q2 營業利益 +0.974）")
    K.add("s1s_h1", "股權報酬代理：2026 上半年＝S-1 調整後營業損益 − GAAP 營業損益", "$B", only26(f"={S('adjoi_s1f_2026h1')}-{S('oi_s1f_2026h1')}"),
          "SRC_ANT_465 − 435＝2.311＝股權報酬 2.185＋非現金捐贈配對 0.126（2026Q1；S-1 第 161 頁，v0.1.7.1 更正 v0.1.7「約 0.11 其他調整項」）；兩者皆非現金，代理口徑成立；SRC_ANT_467（2.2）只對照")
    K.add("s1s_q2", "股權報酬代理：2026 Q2＝Q2 調整後營業利益率 × Q2 營收 − Q2 GAAP 營業損益", "$B",
          only26(f"={S('adjoim_s1f_2026q2')}*{S('rev_s1f_2026q2')}-{S('oi_s1f_2026q2')}"), "SRC_ANT_466（約 19%）、420、441")
    K.add("s1s_fy", "股權報酬代理：2026 全年＝上半年 ＋ Q2 × 下半年季數（照營收同法：Q3＝Q2、Q4＝Q3）", "$B", only26(f"=«s1s_h1»+«s1s_q2»*{hq}"),
          "s1f_sbc_proxy_switch＝1 時進基準（第五節每人股權報酬 2026 欄）", name="COST_S1fSBCProxy2026", name_cols="E")
    K.add("s1s_old", "對照：v0.1.6 算法的 2026 股權報酬（人數 × 2025 每人 ×（1＋年變動率））", "$B", only26(f"=«hc»*«sbc_pp@D»*(1+{I('sbc_pp_g')})/{m2b}"), "舊值")
    K.add("s1s_h1u", "W4 扣：上半年股權報酬（代理開關 1＝上半年代理；0＝舊值 × 上半年月數 ÷ 12）", "$B",
          only26(f"={sws}*«s1s_h1»+(1-{sws})*«s1s_old»*{h2}/{months}"), "")
    K.add("s1s_qu", "W4 扣：每季股權報酬（代理開關 1＝Q2 代理；0＝舊值 ÷ 4）", "$B", only26(f"={sws}*«s1s_q2»+(1-{sws})*«s1s_old»/{qpy}"), "")
    K.add("s1n_h2", "W4 2026 下半年非算力費用（含抽成、不含股權報酬）：Q3 口徑開關 1＝Q3 報導隱含（第十九節）× 下半年季數；0＝（Q2 − 每季股權報酬）× 下半年季數", "$B",
          only26(f"={swq}*«u5_nc»*{hq}+(1-{swq})*(«s1n_q2nc»-«s1s_qu»)*{hq}"), "Q3 隱含＝Q3 營收 16.8 ×（1 − 算力占比 − 調整後營業利益率 11%）；與營收下半年（Q3 實際、Q4＝Q3）同期")
    K.add("s1n_fy", "W4 2026 全年非算力費用（含抽成、不含股權報酬）＝上半年 − 上半年股權報酬 ＋ 下半年", "$B", only26("=«s1n_h1nc»-«s1s_h1u»+«s1n_h2»"), "")
    K.add("s1n_fee", "W4 扣：模型 2026 雲端平台抽成（REV_PartnerShare；模型自營收扣除）", "$B", only26("=«R:pshare»"), "避免與營收淨額重複扣除（D5 r1）")
    K.add("s1n_ncx", "W4 2026 非算力成本（不含股權報酬；正式 S-1 校準）", "$B", only26("=«s1n_fy»-«s1n_fee»"),
          "s1f_d22_switch＝1 時進基準（第五節每人年成本 2026 欄）", name="COST_S1fNonComp2026", name_cols="E")
    K.add("s1n_old", "W4 對照：v0.1.6 算法的 2026 非算力成本（2025 每人成本 ×（1＋年變動率）× 人數）", "$B",
          only26(f"=«pp@D»*(1+{I('pp_cost_g')})*«hc»/{m2b}"), "舊值（並列）")
    K.add("s1n_diff", "W4 推得值 ÷ 舊值 − 1", "比例", only26("=«s1n_ncx»/«s1n_old»-«C:vr_unit@$»"), "", name="COST_S1fNonCompDiff", name_cols="E")
    K.add("s1n_flag", "W4 旗標：|差距| > 門檻（工作單：先報告、照推得值入基準）", "旗標", only26(f"=(ABS(«s1n_diff»)>{I('s1f_d22_warn')})*«C:vr_unit@$»"), "1＝超過 30%")
    K.add("s1n_25", "W4 交叉核對：S-1 2025 非算力營業費用（營收 − 算力 − 營業損益）− 模型 2025 非算力（SRC_ANT_277）", "$B",
          only25(f"={S('rev_s1f_fy2025')}-{S('ci_s1f_fy2025')}-{S('oi_s1f_fy2025')}-«nc25»"), "應接近 0（模型 2025 非算力＝草擬版 12.65 − 7.33）")
    K.add("s1n_q25", "W4 交叉核對：S-1 2025 四季非算力營業費用合計（季度營收 − 算力 − 營業損益）", "$B",
          only25("=" + "+".join(f"({S(f'rev_s1f_2025q{q}')}-{S(f'ci_s1f_2025q{q}')}-{S(f'oi_s1f_2025q{q}')})" for q in (1, 2, 3, 4))), "季度（二手轉載）")
    K.add("s1n_adjoi", "W4 對照：模型 2026 營業結果（營收總額 − 合約實付 − 自建 − 非算力〔含抽成、股權報酬〕）", "$B", only26("=«R:gross»-«cc»-«ncx»-«sbc»-«R:pshare»"),
          "對照 S-1 上半年 −0.693、Q2 +0.974（現金口徑 vs 費用口徑）")
    K.add("s1s_25", "對照：S-1 2025 股權報酬（人事附註）vs 模型 2025（人數 × 每人 0.8）", "$B", [f"={S('sbc_s1f_fy2025')}", None, None, None, None, None], "模型 2025 見第五節（不改 2025）")
    # 持續轉正年（之後各年皆為正；只取值與比較，不改任何計算）
    K.add("ps_gap_y", "命題 1 差額（每 VR 等值 GW）持續為正的年度標示（該年起至 2030 皆 > 0）", "年",
          lambda i, c: None if i == 0 else f'=IF(COUNTIF(«p_gap_vr»:«p_gap_vr@I»,"<=0"),"",{c}$6)', "")
    K.add("ps_gap", "命題 1 持續轉正年（之後各年皆為正；無則「無」）", "年", only25('=IF(COUNT(«ps_gap_y@E»:«ps_gap_y@I»),MIN(«ps_gap_y@E»:«ps_gap_y@I»),"無")'),
          "與單年轉正不同：v0.1.7 的 2026 單年為正、2027–2029 為負", name="COST_GapPosSustainYear", name_cols="D")

    # ═════════════ Funding 第十二節（W6） ═════════════
    Y26 = I("year_2026")
    F.section("十二、v0.1.7 W6：Amazon 2026 實際投入（Amazon 2026Q3 10-Q：Series G 5.0＋Series H 5.0＝10.0；Wattonomics 證據 F#81 附帶觀察）vs 模型已到位融資中的 Amazon 部分（只對照）")
    F.add("w6_10q", "Amazon 10-Q：2026 年對 Anthropic 實際投入（Series G 特別股＋Series H 特別股）", "$B",
          lambda i, c: f"=({S('strat_amzn_2026_series_g')}+{S('strat_amzn_2026_series_h')})*({c}$6={Y26})", "SRC_ANT_319、320（Amazon 10-Q，Verified）")
    F.add("w6_model", "模型：2026 已到位融資中的 Amazon 部分（Amazon Series G-2 5〔開關〕＋Series H 內含的 Amazon 5）", "$B",
          lambda i, c: f"=«in_amzg»+{S('strat_amzn_2026_series_h')}*({c}$6={Y26})", "Series H $65B 含 $15B 先前承諾的雲端業者投資（其中 Amazon $5B，SRC_ANT_300、320）")
    F.add("w6_diff", "差：模型 − 10-Q（應為 0）", "$B", lambda i, c: "=«w6_model»-«w6_10q»", "", name="FND_AmznCheck2026")
    F.add("w6_cum", "對照：Amazon 累計投入 18.0（可轉債 8＋G 5＋H 5；SRC_ANT_323）只用於回流對照（第七節 D 欄），不進 2026 流入", "$B",
          only25(f"={S('strat_amzn_funded_cum_2026h1')}"), "可轉債 8.0 為 2023–2025 年投入，已反映在 2025 年底現金")

    F.section("十三、v0.1.7 補充：營運自由現金流持續轉正年（FND_FCFFirstPos 為首次轉正，v0.1.7 的 2026 單年為正後 2027–2029 又轉負）、正式 S-1 2026-06-30 現金對照（只對照）")
    F.add("ps_fcf_y", "營運自由現金流持續為正的年度標示（該年起至 2030 皆 > 0）", "年",
          lambda i, c: None if i == 0 else f'=IF(COUNTIF(«fcf»:«fcf@I»,"<=0"),"",{c}$6)', "")
    F.add("ps_fcf", "營運自由現金流持續轉正年（無則「無」）", "年", only25('=IF(COUNT(«ps_fcf_y@E»:«ps_fcf_y@I»),MIN(«ps_fcf_y@E»:«ps_fcf_y@I»),"無")'),
          "HTML 與報告用本列；FND_FCFFirstPos 保留（首次轉正）", name="FND_FCFPosSustainYear", name_cols="D")
    F.add("c630_s1", "S-1：2026-06-30 現金、約當現金與短期投資", "$B", only26(f"={S('cash_s1f_2026_06_30')}"), "SRC_ANT_444（Verified）")
    F.add("c630_mod", "模型 2026-06-30 近似＝2025 年底現金 ＋ 2026 已到位融資（皆在上半年交割）＋ 2026 自由現金流 × 上半年月數 ÷ 12", "$B",
          only26(f"={S('cash_2025_12_31')}+«committed»+«fcf»*{h2}/{months}"), "年內平均近似；模型自由現金流為現金口徑")
    F.add("c630_s1b", "對照：以 S-1 上半年營業現金流與資本支出推算＝2025 年底現金 ＋ 已到位融資 ＋ 營業現金流 − 資本支出", "$B",
          only26(f"={S('cash_2025_12_31')}+«committed»+{S('ocf_s1f_2026h1')}-{S('capex_s1f_2026h1')}"), "SRC_ANT_442、443（二手轉載）")
    F.add("c630_gap", "差：S-1 6-30 現金 − 模型近似", "$B", only26("=«c630_s1»-«c630_mod»"),
          "正＝模型少計上半年流入或多計流出；可能原因：利息收入、融資活動的其他流入（員工認股、可轉債轉換）、預付款與營運資金時點、2025 年底現金口徑（只對照，不校準）", name="FND_Cash0630Gap", name_cols="E")
    for sh in (C, K, F):
        sh.resolve()
    ctx.P["a7"] = dict(K=K, F=F)


def checks(ctx):
    D = ctx.D
    I = D.I
    K, F = ctx.P["a7"]["K"], ctx.P["a7"]["F"]
    C = Sheet.all["C"]
    d = lambda sh, key, col="D": f"{sh.name}!${col}${sh.rows[key]}"  # noqa: E731
    mx = lambda sh, key: f"=MAX(MAX({sh.name}!$D${sh.rows[key]}:$I${sh.rows[key]}),-MIN({sh.name}!$D${sh.rows[key]}:$I${sh.rows[key]}))"  # noqa: E731
    return [
        ("g1_s1u", "v0.1.7 W1(b)：S-1 承諾未歸屬餘額（$B；進基準）", "=COST_S1Unalloc", None, "info", "S-1 三段 832 − 模型逐家合約總額（不含 AMD）"),
        ("g1_raw", "v0.1.7 W1(b)：S-1 三段 − 模型逐家（$B；負＝模型已超過）", f"={d(K, 's1u_raw')}", None, "info", ""),
        ("g1_gap", "v0.1.7 W1：模型（逐家＋未歸屬餘額）÷ S-1 三段 − 1", "=COST_S1fCommitGapAfter", None, "info", "承諾開關 1 時為 0"),
        ("g1_spx", "v0.1.7 W1(a)：SpaceXAI 上限讀法係數（基準 1）", "=COST_S1fSpaceXFactor", None, "info", ""),
        ("g2_cpg", "v0.1.7 W2：TK IF_CapexIT＋IF_CapexFacility − IF_CapexTotal（$B/GW）", f"={d(C, 'own_cpgchk')}", 0, "tol", "與舊法每 GW 資本支出同一口徑"),
        ("g2_own26", "v0.1.7 W2：模型 2026 自建資本支出（$B）vs S-1 上半年年化", f"={d(K, 's2_model', 'E')}", None, "info", "S-1 年化見 Cost 第二十五節"),
        ("g3_eta", "v0.1.7 W3：η（2025）", "=CMP_Eta2025", None, "info", "2025 校準改正式 S-1 後重校"),
        ("g1_amd", "v0.1.7 W1(a)：AMD 讀法係數（基準 1；1＝維持模型 GW 型）", "=COST_S1fAMDFactor", None, "info", "S-1 >20 為下限、不在承諾表內"),
        ("g4_sbc", "v0.1.7：2026 股權報酬代理（S-1 調整後 − GAAP；$B）", "=COST_S1fSBCProxy2026", None, "info", "上半年 2.311＋Q2 差額 × 2"),
        ("g5_sus", "v0.1.7：命題 1 持續轉正年", "=COST_GapPosSustainYear", None, "info", ""),
        ("g5_fcf", "v0.1.7：營運自由現金流持續轉正年", "=FND_FCFPosSustainYear", None, "info", "FND_FCFFirstPos 為首次轉正"),
        ("g7_cash", "v0.1.7：S-1 2026-06-30 現金 − 模型近似（$B）", "=FND_Cash0630Gap", None, "info", "只對照"),
        ("g4_nc", "v0.1.7 W4：2026 非算力成本（不含股權報酬；正式 S-1 校準，$B）", "=COST_S1fNonComp2026", None, "info", ""),
        ("g4_diff", "v0.1.7 W4：推得值 ÷ v0.1.6 算法 − 1（> 30% 已在報告首節說明）", "=COST_S1fNonCompDiff", None, "info", "工作單：照推得值入基準"),
        ("g4_25", "v0.1.7 W4：S-1 2025 非算力 − 模型 2025 非算力（$B）", f"={d(K, 's1n_25')}", None, "info", "交叉核對"),
        ("g6_amzn", "v0.1.7 W6：模型 2026 Amazon 流入 − 10-Q 實際 10.0（$B）", mx(F, "w6_diff"), None, "info", "模型未以 18.0 計入 2026 流入"),
    ]
