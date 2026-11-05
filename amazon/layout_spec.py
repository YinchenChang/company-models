# Amazon 個股 HTML 版面（docs/plan/個股HTML版面規格.md v1.0；tools/company_html/relayout.py 讀本檔 SPEC）
SPEC = {
 "in_skip": ["H", "I", "J"],
 "in_groups": [
  ["AWS AI（對外 AI 算力）", [
    ["對外 AI 容量與擴產速度", ["mw_ye23", "mw_ye24", "mw_ye25", "mw_open", "pace_A", "pace_B", "pace_C", "pace_x", "mw_q3", "ext_sh", "cip_lag", "capex_g26"]],
    ["價格與 k（每 MW 年收入 ÷ 持有成本）", ["long_sh", "p_spot", "p_list", "p_long", "p_long_lo", "k_long", "k_spot", "k_long_lo2", "k_spot_lst", "k_A", "k_B", "k_C", "tk_pq"]],
    ["每 MW 成本（Tokenomics 快照）", ["tk_he_h100", "tk_he_gb200", "tk_he_gb300", "tk_he_vr200", "tk_cx_h100", "tk_cx_gb200", "tk_cx_gb300", "tk_cx_vr200",
                                       "tk_ox_h100", "tk_ox_gb200", "tk_ox_gb300", "tk_ox_vr200", "tk_gpuhr", "tk_fac", "tk_life", "fac_life"]],
    ["世代組合與 Trainium", ["mo_trn2", "mo_h100", "mo_gb200", "mo_gb300", "he_open", "ox_open", "rpm_open", "cust_sh", "na_vr26", "na_vr27", "na_vr28", "trn_hold", "trn_out", "cf_chip", "cf_cx"]],
    ["機房租用", ["colo_sh", "rent_rec", "rent_mw"]],
    ["自用 AI 讀法與公司說法", ["internal_ai_reading", "ai_rr", "aws_rr", "aws_bl", "aws_g40"]],
    ["分部加總中的 AWS AI", ["ai_leg", "m_ai"]]]],
  ["AWS 其他（非 AI 雲端）", [
    ["成長與資本支出", ["g_awso_lt", "awso_g0", "cx_awso"]]]],
  ["零售與廣告（North America、International）", [
    ["各營收線長期成長", ["g_os", "g_ps", "g_sp", "g_ad", "g_sub", "g_ot"]],
    ["利益率與資本支出", ["m_ad", "step_na", "step_in", "cx_rt"]]]],
  ["全公司：價格、折現率、現金與持股", [
    ["價格、股數與共識目標價", ["px", "px_val", "sh_out", "sh_awd", "sh_val", "cons_tp"]],
    ["折現率（CAPM）與目標價權重", ["rf", "beta_sa", "beta_gf", "beta", "erp", "kd_new", "kd_spr", "kd", "tax", "debt_q3", "fl_pv", "fin_obl", "fin_debt", "ke", "wacc_adj", "wacc", "tg", "ronic", "w_dcf", "w_A", "w_B", "w_C", "tp_yf"]],
    ["公司層與現金", ["sbc_r", "oth_r", "oth_q4", "lease_pr", "sh_g", "kd_book", "y_cash", "cash_q3", "min_cash", "inv_q4", "dps"]],
    ["持股與分部加總設定", ["ant_val", "ant_sh", "oai_val", "oai_sh", "liq_disc", "sbc_ab", "m_x"]]]]],
 "blocks_to": {"驅動 → 推導量": "AWS AI（對外 AI 算力）"},
 "main": {"0. 目標價": ["DCF 情境價值推到 12 個月後", "Comparable（分部加總）情境價值", "情境目標價", "情境權重", "FY2028 AWS 營業利益率"],
          "sheet:Recon": ["FY2027 營收", "FY2028 營收", "FY2028 營業利益", "FY2027 資本支出", "FY2027 AWS", "FY2026 現金資本支出 vs 指引", "AWS 年增 FY2027"]},
 "n_recon": 6,
}
