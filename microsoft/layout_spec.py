# Microsoft 個股模型 HTML 版面（個股 HTML 版面規格 v1.0，docs/plan/個股HTML版面規格.md）：relayout 用的 SPEC。
# 用法：python3 tools/company_html/relayout.py microsoft microsoft/dist/<成品>.xlsx（make_v03.sh 在 LibreOffice 重算後執行）
# In 頁每個鍵須剛好歸入一組（relayout 建置時檢查）；H（共識與指引，只作 Recon 對照）、J（情境驅動對照，已在情境驅動表）不歸組。
# 摘要、小組標題不得含阿拉伯數字（FY／CY 加四位年度除外）。
AI = "Azure 與 AI 基礎設施（AI 算力、自研晶片、OpenAI 持股）"
M365 = "生產力與企業軟體（Office 雲端訂閱、Copilot、授權）"
DC = "Windows、搜尋與 XBOX（Devices and Consumer）"
CO = "全公司：資產負債、資本配置、折現率、目標價規則、稅與現金流"

SPEC = {
    "in_skip": ["H", "J"],
    "in_groups": [
        [AI, [
            ["AI 容量與擴產速度（在役對外 MW、租用、對外比例、資本支出指引）",
             ["mw_ext_24", "mw_ext_25", "mw_ext_26", "rent_ext_24", "rent_ext_25", "rent_ext_26", "rent_ext_f",
              "pace", "pace_v03", "pace_nv0", "pace_y0", "pace_y1", "pace_epoch", "cap_epoch", "ext_sh", "capex_floor27"]],
            ["每 MW 價格 k（長約與現貨證據、情境長約占比）",
             ["p_long", "p_spot", "k_long", "p_long_lo", "k_long_lo", "k_spot", "ls_A", "ls_B", "ls_C", "k_A", "k_B", "k_C", "k_ramp", "rent_mw"]],
            ["世代組合與自研晶片（Maia）",
             ["mix_h", "mix_g2", "mix_g3", "a_g3_27", "a_g3_28", "a_g3_29", "nonnv", "a_mai", "nv_out_mai", "nv_hold_mai", "chip_f"]],
            ["Tokenomics 每 MW 成本、機房與折舊",
             ["tk_hold_h100", "tk_hold_gb200", "tk_hold_gb300", "tk_hold_vr200", "tk_cxit_h100", "tk_cxit_gb200", "tk_cxit_gb300", "tk_cxit_vr200",
              "tk_opex_h100", "tk_opex_gb200", "tk_opex_gb300", "tk_opex_vr200", "tk_cxfac", "tk_gpuhr_h100", "life_it", "life_fac", "fac_wacc", "fac_rent",
              "self_build", "cxit_open", "ai_net_open"]],
            ["自用 AI 讀法", ["internal_ai_reading", "shadow_base"]],
            ["Azure 其他（非 AI 雲端）", ["glt_azo", "cx_azo", "m_cloud"]],
            ["估值：AI 倍數對照與 OpenAI 持股", ["m_ai", "oai_stake", "oai_stake26", "oai_val", "liq_disc", "oai_v"]]]],
        [M365, [
            ["Copilot 座席與月費", ["cop_s24", "cop_s25", "cop_s26", "cop_px", "cop_add_A", "cop_add_B", "cop_add_C", "cop_tail"]],
            ["產品線長期成長", ["glt_m365b", "glt_lic", "glt_ind", "glt_fro"]],
            ["非 AI 資本支出", ["cx_m365", "cx_lic", "cx_ind"]],
            ["估值：企業軟體同業倍數與利益率分配", ["m_sw", "al_m365", "al_lic", "al_ind", "al_fro"]]]],
        [DC, [
            ["產品線長期成長", ["glt_srch", "glt_xbox", "glt_win"]],
            ["非 AI 資本支出", ["cx_srch", "cx_dev"]],
            ["估值：數位廣告、遊戲、PC 同業倍數與利益率分配", ["m_ads", "m_game", "m_pc", "al_srch"]]]],
        [CO, [
            ["股價、股數與資產負債（評價日）",
             ["px", "sh_dil26", "sh_basic26", "dil_gap", "sh_out26", "cash_open", "debt_face", "sh_dil_q1", "sh_out_q1", "cash_q1", "debt_q1", "fl_liab_q1",
              "sh_val", "shd_val", "cash_val", "debt_val", "fl_val", "debt_mat_27", "debt_mat_28", "debt_mat_29", "debt_mat_30", "debt_mat_31", "debt_mat_32",
              "fl_liab", "fl_rate", "fl_pay_27", "fl_pay_28", "fl_pay_29", "fl_pay_30", "fl_pay_31", "fl_pay_32"]],
            ["資本配置：最低現金、回購、員工股票與股利",
             ["min_cash", "bb_target", "bb_g", "sh_iss", "dps_q0", "dps_q1", "dps_27", "div_rem27", "dps_g"]],
            ["折現率", ["rf", "beta_sa", "beta_gf", "blume_w", "beta", "erp", "kd_aaa", "kd_spread", "kd", "tax_w", "ke", "wacc", "g_tv", "ronic_prem"]],
            ["目標價規則與時點",
             ["base_fy", "mid_yr", "val_d0", "roll_d1", "roll_d2", "days_yr", "v_off", "t_roll", "roll_div_q", "q_per_yr", "stub_d0", "t_stub", "fy28_end",
              "w_ntm28", "w_dcf", "w_A", "w_B", "w_C", "m_own"]],
            ["損益、現金流與成長收斂", ["n_conv", "dm_seg", "oie_f", "tax_f", "sbc_r", "wc27", "wc26q1", "wc_f", "life_nonai"]],
            ["前版對照（只用於對帳）", ["old_tp", "old_dcf", "old_ev", "old_w"]]]],
    ],
    # 既有驅動區塊：Thesis 說明與命題模組表放 Azure 與 AI 基礎設施塊
    "blocks_to": {"@Thesis!A1": AI, "@Thesis!A2": AI, "情境 B": AI},
    "main": {
        "0.1 DCF": ["DCF 企業價值", "DCF 每股價值（評價日）", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）", "WACC／永續成長率"],
        "0.2 分部加總": ["EV：AI 基礎設施模組", "EV：Microsoft 365 cloud", "分部加總 EV", "分部加總每股價值", "Comparable 目標價"],
        "0.3 結論": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "市場共識目標價", "現價", "上檔空間：加權目標價", "上檔空間：Comparable",
                   "投資建議", "情境目標價", "目標價（讀法 1"],
        "收支關卡": ["收支關卡：", "是否需要舉債", "自由現金流 FY2027", "回購合計", "新增舉債合計"],
        "1. 合併層級": ["目標價：模型加權（讀法 2）", "營收 FY2027", "營業利益 FY2029", "EPS FY2029"],
        "2. 公司指引": ["FY27Q1 實績 vs 指引中點：合併營收", "FY27Q1 實績 vs 指引：Azure 年增", "Azure：第二至第四季", "資本支出 FY2027（模型）vs 指引"],
        "3. 命題驗證": ["Azure AI＋Copilot 營收 FY2026", "Azure 實績 FY2026 中 Azure AI 占比", "Copilot 營收 FY2026", "Azure AI FY2027"],
        "4. 產品線層級": ["Azure：", "Microsoft 365 cloud：", "Search and advertising：", "XBOX："],
        "5. v0.3 → v0.3.1": ["v0.3 加權目標價", "(c)", "(d)", "v0.3.1 加權目標價"],
        "6. 新季實績": ["Azure", "合併營收", "合併營業利益", "資本支出（公司口徑）"],
        "7. v0.3.1 vs v0.2.2": ["v0.2.2 基準目標價", "② DCF 腿", "③ 倍數腿", "v0.3.1 加權目標價"],
    },
}
