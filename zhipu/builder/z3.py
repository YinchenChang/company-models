"""智譜 v0.1-Z3：算力頁 Compute 與成本頁 Cost（命題表），以及 Z3 的 Checks 列。

版面同 Z2：A 編號、B 項目、C 單位、D–I 2025–2030、J 空白、K 1H2026、L 2H2026、M 說明。
- GW 列＝期間平均 GW（IT 關鍵電力）；半年欄的 GW＝半年 token 或金額 ÷（每 GW 年值 × 半年天數 ÷ 年天數）；2026（E 欄）＝兩個半年依天數加權。
- 金額列（RMB 億）：2026（E 欄）＝1H26＋2H26。
機制（規格 D6–D13、D23、r1 的 D8r、D9r、D10r、D11r；工作單 Z3）：
  晶片族（Hopper＝TK Hopper H100 欄；H20、國產＝Hopper × Inputs 比例，Tokenomics 缺口）→ 每 GW 年產能（IF_Util ÷ Σ 層級占比 ÷ 每 GW 產出）
  → token 換算 GW；算力服務費（2025、1H26 財報推得）÷ 每 GW 價格（卡時租價 × 每 GW 卡數 × 小時）→ 供給 GW、支出換算推論 GW → η；
  2H26 起：有效推論 GW＝token GW ÷ η；供給＝有效推論 ÷ [(1−閒置)(1−研發占比)]；研發 GW＝殘差；容量上限（V1a 同式）回乘雲端營收。
所有數值引用 SRC_ZP／Inputs／TK_Link 或同簿儲存格（E6）；Python 只組公式文字。同頁後建列以「«C.鍵»」「«K.鍵»」佔位，由 z2.fill() 代換。
"""
from __future__ import annotations

from common import F_NOTE, put
from z2 import ALL, HC, LATE, PREV, YC, Sheet

FAM = ("hop", "h20", "dom")
FAM_ZH = {"hop": "Hopper（H800／A800 等；TK Hopper H100 欄）", "h20": "H20", "dom": "國產（昇騰 910B／910C 等）"}
TIERS = ("sol", "luna")
TIER_ZH = {"sol": "Sol", "luna": "Luna"}
YEARCOLS = ("D",) + tuple(LATE)          # 年度驅動欄（不含 2026 E；E 由半年合成）


def build(wb, D, Z, snap):
    S, I = D.S, D.I
    Dm, V = Z["D"], Z["V"]
    dr, vr = Dm.rows, V.rows
    dy, d1, d2 = I("days_year"), I("days_1h26"), I("days_2h26")
    halves, pct, fx, bn = I("halves_year"), I("pct"), I("fx_usdcny"), I("mul_bn_yi")
    y26 = I("year_2026")
    hop_key, vr_key, base_cost = snap["hdr_gen"][0], snap["hdr_gen"][9], snap["hdr_cost"][1]
    assert "Hopper" in hop_key and "VR200" in vr_key, (hop_key, vr_key)

    def frac(c):
        """期間占一年的比例（年度欄為 None）。"""
        return {"K": f"{d1}/{dy}", "L": f"{d2}/{dy}"}.get(c)

    def yr(c):
        return y26 if c in HC else f"{c}$6"

    def per(c, x):
        """每 GW 年值 × 期間比例。"""
        f = frac(c)
        return f"({x}*{f})" if f else x

    def wavg(key):
        """2026 E 欄：兩個半年依天數加權（GW 等存量）。"""
        return f"=(K«C.{key}»*{d1}+L«C.{key}»*{d2})/{dy}"

    def share_in(kind):
        return {c: I(f"share_{kind}_{2025 if c == 'D' else (2026 if c in ('E', 'K', 'L') else 2027 + LATE.index(c))}") for c in ALL}

    # ══════════════════════════════════ Compute ══════════════════════════════════
    C = Sheet(wb, "Compute", "G",
              "Compute — 算力：晶片族（Hopper／H20／國產）每 GW 產能、組合、token 換算 GW、算力服務費與每 GW 價格 → 供給 GW、η、有效推論 GW、研發 GW（殘差）、容量上限、VR 等值、10 萬國產晶片對照、敏感度",
              "規格 D6–D12、r1 D8r／D9r／D10r／D11r。產能取 TK_Link（Hopper H100 欄的 IF_TokGW_*、IF_Util）；H20、國產＝Hopper × Inputs 比例（Tokenomics 缺口）。"
              "2025、1H26：供給 GW＝算力服務費 ÷ 每 GW 價格（智譜不揭露 GW，只能由金額推）；η＝token 換算推論 GW ÷ 支出換算推論 GW。2H26 起：供給依需求配置（見第八節）。",
              "GW＝IT 關鍵電力、期間平均；半年欄（K、L）為該半年平均 GW；2026（E）＝兩個半年依天數加權。每 GW 價格 RMB 億／GW／年。token：T（兆）；產能：M tok／GW／年。")
    C.years(D)
    cr = C.rows

    # 一、晶片族每 GW 產能
    C.section("一、晶片族每 GW 年產能（100% 利用率；TK IF_TokGW_* Hopper H100／VR200 NVL72 基準欄；H20、國產＝Hopper × 比例）")
    C.add("Tokenomics 世代鍵：Hopper", "文字", {"D": hop_key}, "TK_HdrGen 的世代名稱（文字鍵，SUMIFS 用）", key="k_hop")
    C.add("Tokenomics 世代鍵：VR200（VR 等值基準）", "文字", {"D": vr_key}, "", key="k_vr")
    C.add("Tokenomics 成本情境鍵：基準", "文字", {"D": base_cost}, "TK_HdrCost（每 GW 總產出三個成本情境同值）", key="k_cost")
    kh, kv, kc = "$D$«C.k_hop»", "$D$«C.k_vr»", "$D$«C.k_cost»"
    for t in TIERS:
        C.add(f"TK 每 GW 總產出：Hopper × {TIER_ZH[t]}", "M tok/GW/年", {c: f"=SUMIFS(TK_IF_TokGW_{t.capitalize()},TK_HdrGen,{kh},TK_HdrCost,{kc})" for c in ALL},
              f"TK IF_TokGW_{t.capitalize()}（Hopper H100、基準欄）", key=f"tk_hop_{t}")
    C.add("TK 每 GW 總產出：VR200 × Sol", "M tok/GW/年", {c: f"=SUMIFS(TK_IF_TokGW_Sol,TK_HdrGen,{kv},TK_HdrCost,{kc})" for c in ALL},
          "VR 等值換算基準（同 OpenAI v0.6：Sol 層級產能比）", key="tk_vr_sol")
    C.add("H20 產出比（÷ Hopper）", "倍", {c: f"={I('h20_ratio')}" for c in ALL}, "Inputs（SRC_ZP_490 區間；D6）", key="r_h20")
    C.add("國產產出比（÷ Hopper）", "倍", {c: f"={I('dom_ratio')}" for c in ALL}, "Inputs（D6，0.4–0.9）", key="r_dom")
    for f in FAM:
        for t in TIERS:
            fx_ = {"hop": "", "h20": f"*{{c}}«C.r_h20»", "dom": f"*{{c}}«C.r_dom»"}[f]
            C.add(f"每 GW 總產出：{FAM_ZH[f]} × {TIER_ZH[t]}", "M tok/GW/年",
                  lambda c, t=t, fx_=fx_: f"={c}«C.tk_hop_{t}»" + fx_.format(c=c), "Hopper × 產出比（Tokenomics 缺口：無 H20／昇騰欄）", key=f"cap_{f}_{t}")

    # 二、晶片族組合
    C.section("二、晶片族組合（實體 GW 占比；D7：時變，無新 NVIDIA 採購；Hopper＝1 − 國產 − H20）")
    sd, sh = share_in("dom"), share_in("h20")
    C.add("國產", "比例", {c: f"={sd[c]}" for c in ALL}, "Inputs（Assumed；2026 起依『10 萬級國產晶片用於推論』）", key="s_dom", name="CMP_Share_Domestic")
    C.add("H20", "比例", {c: f"={sh[c]}" for c in ALL}, "Inputs（Assumed）", key="s_h20", name="CMP_Share_H20")
    C.add("Hopper（H800／A800 等）", "比例", {c: f"=(1-{c}«C.s_dom»-{c}«C.s_h20»)" for c in ALL}, "＝1 − 國產 − H20", key="s_hop", name="CMP_Share_Hopper")
    C.add("組合產出係數（÷ Hopper）＝Σ 占比 × 產出比", "倍", {c: f"={c}«C.s_hop»+{c}«C.s_h20»*{c}«C.r_h20»+{c}«C.s_dom»*{c}«C.r_dom»" for c in ALL},
          "各族每 GW 產出＝Hopper × 比例，故組合產能＝Hopper 產能 × 本係數", key="mixf")

    # 三、token 與層級組合
    C.section("三、token（Demand；T）與層級組合：付費、免費分開（半年欄：API／Coding Plan 取 Demand 半年欄；清言為年度值 × 天數比例）")

    def tok(kind, t):
        def f(c):
            if c == "E":
                return f"=K«C.t_{kind}_{t}»+L«C.t_{kind}_{t}»"
            if c in YC:
                return f"=Demand!{c}{dr[f'{kind}_{t}']}"
            if kind == "paid":
                return f"=Demand!{c}{dr[f'api_tok_{t}']}+Demand!{c}{dr[f'cptok_{t}']}"
            q = f"Demand!$E${dr[f'qy_{t}']}*{frac(c)}"
            return f"={q}+Demand!{c}{dr['api_free']}" if t == "luna" else f"={q}"
        return f
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        for t in TIERS:
            C.add(f"{zh} token：{TIER_ZH[t]}", "T", tok(kind, t), "Demand DEM_Tok_*（半年欄見本節標題）", key=f"t_{kind}_{t}")
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        C.add(f"{zh} token 合計", "T", {c: f"={c}«C.t_{kind}_sol»+{c}«C.t_{kind}_luna»" for c in ALL}, "", key=f"t_{kind}")
        C.add(f"{zh}：Sol 占比", "比例", {c: f"={c}«C.t_{kind}_sol»/{c}«C.t_{kind}»" for c in ALL}, "", key=f"m_{kind}_sol")
        C.add(f"{zh}：Luna 占比", "比例", {c: f"={c}«C.t_{kind}_luna»/{c}«C.t_{kind}»" for c in ALL}, "", key=f"m_{kind}_luna")

    # 四、每 GW 年產能與 token 換算 GW
    C.section("四、每 GW 年產能（基準利用率 TK IF_Util；IF_Util ÷ Σ 層級占比 ÷ 每 GW 產出，同 Tokenomics Alloc 與 OpenAI v0.6）與 token 換算推論 GW")
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        for f in FAM:
            C.add(f"{zh}每 GW 年產能：{FAM_ZH[f]}", "M tok/GW/年",
                  {c: f"=TK_IF_Util/({c}«C.m_{kind}_sol»/{c}«C.cap_{f}_sol»+{c}«C.m_{kind}_luna»/{c}«C.cap_{f}_luna»)" for c in ALL}, "", key=f"pc_{kind}_{f}")
        C.add(f"{zh}每 GW 年產能：組合後", "M tok/GW/年",
              {c: "=" + "+".join(f"{c}«C.s_{f}»*{c}«C.pc_{kind}_{f}»" for f in FAM) for c in ALL}, "Σ 晶片族占比 × 各族產能", key=f"pc_{kind}",
              name=f"CMP_Cap{kind.capitalize()}")
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        C.add(f"推論 GW（token 換算）：{zh}", "GW",
              lambda c, kind=kind: wavg(f"g_{kind}") if c == "E" else f"={c}«C.t_{kind}»*{I('mul_t_m')}/{per(c, f'{c}«C.pc_{kind}»')}",
              "token（T）× 10^6 ÷（每 GW 年產能 × 期間比例）", key=f"g_{kind}", name=f"CMP_TokGW_{kind.capitalize()}")
    C.add("推論 GW（token 換算）：合計", "GW", {c: f"={c}«C.g_paid»+{c}«C.g_free»" for c in ALL}, "", key="g_tok", name="CMP_TokGW")

    # 五、算力服務費（D8r、D10r）
    C.section("五、算力服務費（財報推得；RMB 億；2025 年報、1H26 中期公告）：推論＝開放平台及 API 銷售成本 × 占比（D8r）；研發＝研發開支（扣股權報酬）× 算力占比（D10r）")
    cal = ("D", "K")
    sv = lambda k25, k1h: {"D": f"={S(k25)}", "K": f"={S(k1h)}"}  # noqa: E731
    C.add("開放平台及 API 銷售成本（推算）", "RMB 億", sv("cogs_api_fy25", "cogs_api_1h26"), "SRC_ZP_197、195（Derived：毛利 − 收入）", key="f_cogs_api")
    C.add("推論算力支出＝API 銷售成本 × 算力費占比", "RMB 億", {c: f"={c}«C.f_cogs_api»*{I('inf_fee_share')}" for c in cal},
          "Inputs（D8r：基準 1＝上限；下限 0.22＝FY2024 銷售成本中算力費占比）", key="f_inf", name="CMP_InfFee")
    C.add("研發開支", "RMB 億", sv("rd_fy25", "rd_1h26"), "SRC_ZP_139、199", key="f_rd_tot")
    C.add("股權報酬（總額）", "RMB 億", sv("sbc_fy25", "sbc_1h26"), "SRC_ZP_142、202", key="f_sbc")
    C.add("研發開支（扣股權報酬中研發部分）", "RMB 億", {c: f"={c}«C.f_rd_tot»-{c}«C.f_sbc»*{I('sbc_rd_share')}" for c in cal},
          "股權報酬 × 研發比例（Inputs，Analogy：研發人員占比 74.4%）", key="f_rd_x")
    C.add("研發算力支出（基準：占比法）＝研發（扣股權報酬）× 算力占比", "RMB 億", {c: f"={c}«C.f_rd_x»*{I('rd_compute_share')}" for c in cal},
          "D10r 兩法並列、取本法為基準；占比 Inputs（FY2024 70.7%，60–80%）", key="f_rd", name="CMP_RDFee")
    C.add("研發算力支出（對照：扣除法）＝研發（扣股權報酬）−（薪酬總額 − 股權報酬）× 研發人員占比", "RMB 億",
          {c: f"={c}«C.f_rd_x»-({S('remun_fy25' if c == 'D' else 'remun_1h26')}-{c}«C.f_sbc»)*{S('rd_staff_pct_1h25')}/{pct}" for c in cal},
          "D10r 第一法；薪酬 SRC_ZP_163、224；研發人員占比 SRC_ZP_240（Interested-party）；扣除法把研發中的非人事非算力費用也算成算力（上限）", key="f_rd_a")
    C.add("對照：FY2024 研發算力占研發開支（SRC）", "比例", {"D": f"={S('rd_compute_fy24')}/{S('rd_fy24')}"}, "SRC_ZP_232 ÷ SRC_ZP_048（媒體轉述招股章程）", key="f_fy24")
    C.add("算力服務費合計（推論＋研發）", "RMB 億", {c: f"={c}«C.f_inf»+{c}«C.f_rd»" for c in cal}, "", key="f_tot", name="CMP_ComputeFee")
    C.add("研發占算力服務費", "比例", {c: f"={c}«C.f_rd»/{c}«C.f_tot»" for c in cal}, "1H26 值為 2H26 起研發占比路徑的起點（第八節）", key="f_rdsh")

    # 六、每 GW 價格（D9、D9r）
    C.section("六、每 GW 年算力價格（租用；D9r）＝卡時租價 × 每 GW 卡數 × 每年小時 ÷ 10^8（元→億）；2H26 起 ×（1＋年變動率），2H26 以半年計")
    for f, k in (("hop", "kw_hopper"), ("h20", "kw_h20"), ("dom", "kw_dom")):
        C.add(f"每卡 IT 功耗：{FAM_ZH[f]}", "kW／卡", {c: f"={I(k)}" for c in ALL}, "Inputs（Analogy；Hopper 依 TK DC_Cost 每 GW GPU 數換算）", key=f"kw_{f}")
        C.add(f"每 GW 卡數：{FAM_ZH[f]}", "卡／GW", {c: f"={I('kw_per_gw')}/{c}«C.kw_{f}»" for c in ALL}, "10^6 kW ÷ 每卡 kW", key=f"n_{f}")
    rent = {"hop": f"={S('rent_h800_cardhr_derived')}", "h20": f"={S('rent_h20_cardhr_sep26')}",
            "dom": f"=AVERAGE({S('rent_910b4_cardhr_derived')},{S('rent_910b_zxy_cardhr_derived')})*{I('dom_card_mult')}"}
    rnote = {"hop": "SRC_ZP_474 H800（SMM 兩年約折算 13.36；阿里雲 18.75／圖靈 11.83 為區間 SRC_ZP_472–473）",
             "h20": "SRC_ZP_462 H20 京津冀均價 7.47（7.03–8.16）",
             "dom": "昇騰 910B 折算卡時價 2.74／4.28 的平均（SRC_ZP_478、479）× 國產卡別倍數（Inputs；910C 租價找不到）"}
    chg = I("px_gw_chg")
    for f in FAM:
        def rf(c, f=f):
            if c in ("D", "K"):
                return rent[f]
            if c == "L":
                return f"=K«C.rent_{f}»*(1+{chg}/{halves})"
            if c == "E":
                return wavg(f"rent_{f}")
            return f"={PREV[c]}«C.rent_{f}»*(1+{chg})"
        C.add(f"卡時租價：{FAM_ZH[f]}", "元／卡／時", rf, rnote[f] + "；2025 以 2026 觀察值計（2025 租價找不到）", key=f"rent_{f}")
    for f in FAM:
        C.add(f"每 GW 年價格：{FAM_ZH[f]}", "RMB 億／GW／年",
              {c: f"={c}«C.rent_{f}»*{c}«C.n_{f}»*{I('hours_year')}/{I('div_yuan_yi')}" for c in ALL}, "", key=f"p_{f}")
    C.add("每 GW 年價格：組合後（Σ 占比 × 各族價格）", "RMB 億／GW／年", {c: "=" + "+".join(f"{c}«C.s_{f}»*{c}«C.p_{f}»" for f in FAM) for c in ALL},
          "供給 GW 換算與 2H26 起算力成本用", key="p", name="CMP_PricePerGW")
    C.add("對照：組合後每 GW 年價格（美元）", "$B／GW／年", {c: f"={c}«C.p»/{fx}/{bn}" for c in ALL}, "÷ USD/CNY ÷ 10（億→十億）；OpenAI v0.6 合約價 12（INP_229，8–20）", key="p_usd")

    # 七、η（D8、D8r）
    C.section("七、η（有效產出係數）＝token 換算推論 GW ÷ 支出換算推論 GW（2025、1H26 各自校準；2H26 起依 Inputs 路徑，基準沿用 1H26）")
    C.add("支出換算推論 GW＝推論算力支出 ÷（每 GW 價格 × 期間比例）", "GW", {c: f"={c}«C.f_inf»/{per(c, f'{c}«C.p»')}" for c in cal},
          "D8r", key="e_spend", name="CMP_SpendGW")
    C.add("η（校準值）", "倍", {c: f"={c}«C.g_tok»/{c}«C.e_spend»" for c in cal}, "D 欄＝2025、K 欄＝1H26（Derived）", key="e_cal")
    C.add("η 2025", "倍", {"D": "=D«C.e_cal»"}, "", key="e25", name="CMP_Eta2025", name_col="D")
    C.add("η 1H26", "倍", {"K": "=K«C.e_cal»"}, "", key="e1h", name="CMP_Eta1H26", name_col="K")
    sw, tgt, ty = I("eta_sw"), I("eta_target"), I("eta_target_year")

    def eta(c):
        if c in cal:
            return f"={c}«C.e_cal»"
        if c == "E":
            return "=E«C.g_tok»/E«C.eff»"
        return f"=$K$«C.e_cal»+{sw}*({tgt}-$K$«C.e_cal»)*MIN({ty}-{y26},{yr(c)}-{y26})/({ty}-{y26})"
    C.add("η（逐年）", "倍", eta, "2H26 起＝η₁H₂₆＋開關 ×（目標 − η₁H₂₆）× MIN(目標年 − 2026, 年 − 2026) ÷（目標年 − 2026）；開關 0（基準）＝沿用 1H26；E 欄＝token GW ÷ 有效 GW",
          key="eta", name="CMP_Eta")
    C.add("有效推論 GW（未截頂）＝token 換算 GW ÷ η", "GW", lambda c: wavg("eff") if c == "E" else f"={c}«C.g_tok»/{c}«C.eta»",
          "2025、1H26＝支出換算 GW（校準恆等式）", key="eff", name="CMP_InfGW_Eff")

    # 八、供給 GW
    C.section("八、供給 GW：2025、1H26＝算力服務費 ÷ 每 GW 價格（租用）；2H26 起＝有效推論 GW ÷［(1−閒置)×(1−研發占比)］（依需求配置）；自有 GW 依資本支出（D11r 基準 0）")
    tgt_rd = I("rd_share_2030")
    y30 = I("year_2030")

    def rdsh(c):
        if c in cal:
            return f"={c}«C.f_rdsh»"
        if c == "E":
            return None
        return f"=$K$«C.f_rdsh»+{I('rd_share_sw')}*({tgt_rd}-$K$«C.f_rdsh»)*MIN({y30}-{y26},{yr(c)}-{y26})/({y30}-{y26})"
    C.add("研發占非閒置供給比例（路徑）", "比例", rdsh, "2025、1H26＝實際（研發算力費 ÷ 合計）；2H26 起＝1H26 實際＋開關 ×（2030 目標 − 1H26）× 進度；開關 0（基準）＝沿用 1H26",
          key="rdsh", name="CMP_RDShare")
    C.add("每 GW 資本支出（國產；IT＋廠房）", "RMB 億／GW",
          {c: f"=SUMIFS(TK_IF_CapexTotal,TK_HdrGen,{kh},TK_HdrCost,{kc})*{I('hold_ratio_dom')}*{fx}*{bn}" for c in ALL},
          "TK IF_CapexTotal（Hopper 基準欄，$B/GW）× 國產持有比 × USD/CNY × 10（D11）", key="cpx_gw")
    C.add("自有資本支出：1 GW 資料中心情境", "RMB 億",
          {c: f"={I('dc1gw_sw')}*{S('cmp_1gw_dc')}/{I('pue')}*{c}«C.cpx_gw»*(({c}$6+1)={I('dc1gw_year')})" for c in LATE},
          "開關 × 1 GW（SRC_ZP_458，口徑未明→視為設施）÷ PUE × 每 GW 資本支出；計入投產前一年（D11r：基準關）", key="cpx_dc")
    C.add("自有資本支出：其他", "RMB 億", {c: f"={I('owned_capex')}" for c in LATE}, "Inputs（D11r：基準 0）", key="cpx_oth")
    C.add("自有資本支出合計", "RMB 億", {c: f"={c}«C.cpx_dc»+{c}«C.cpx_oth»" for c in LATE}, "2025、2026 無（1H26 資本開支為辦公樓，D11r）", key="cpx", name="CMP_OwnedCapex")
    C.add("自有 GW＝至前一年累計自有資本支出 ÷ 每 GW 資本支出", "GW",
          {c: f"=SUMPRODUCT(($F$6:$I$6<{c}$6)*($F$«C.cpx»:$I$«C.cpx»))/{c}«C.cpx_gw»" for c in LATE}, "落後一年投產（同 OpenAI v0.6 第十三節）", key="own", name="CMP_Supply_Owned")
    idle = I("idle")

    def sup(c):
        if c in cal:
            return f"={c}«C.f_tot»/{per(c, f'{c}«C.p»')}"
        if c == "E":
            return wavg("sup")
        req = f"{c}«C.eff»/((1-{idle})*(1-{c}«C.rdsh»))"
        return f"={req}" if c == "L" else f"=MAX({req},{c}«C.own»)"
    C.add("供給 GW（租用＋自有）", "GW", sup, "2025、1H26：算力服務費 ÷（每 GW 價格 × 期間比例）；之後：有效推論 ÷［(1−閒置)(1−研發占比)］，且不低於自有 GW", key="sup", name="CMP_SupplyGW")
    C.add("租用 GW＝供給 − 自有", "GW", lambda c: wavg("rent_gw") if c == "E" else (f"={c}«C.sup»-{c}«C.own»" if c in LATE else f"={c}«C.sup»"),
          "", key="rent_gw", name="CMP_SupplyRentGW")
    C.add("閒置 GW（2H26 起）", "GW", lambda c: wavg("idle") if c == "E" else (None if c in cal else f"={c}«C.sup»*{idle}"),
          "OpenAI v0.6：閒置保留自校準期後起算；校準期的閒置含在研發殘差內", key="idle")

    # 九、容量上限、研發 GW、總需求
    C.section("九、容量上限（V1a 同式）、研發 GW（殘差，D12）與總需求 GW")
    C.add("推論可用 GW＝供給 ×（1−閒置）×（1−訓練最低占比）", "GW",
          lambda c: wavg("avail") if c == "E" else f"={c}«C.sup»*(1-{idle})*(1-{I('min_train')})", "", key="avail", name="CMP_InfAvailGW")
    C.add("推論 GW（截頂後）", "GW", lambda c: wavg("infcap") if c == "E" else (f"={c}«C.eff»" if c in cal else f"=MIN({c}«C.avail»,{c}«C.eff»)"),
          "校準期＝有效推論（營收＝實際）；之後＝MIN（可用, 有效）", key="infcap", name="CMP_InfGW")
    C.add("容量上限係數（1＝未受限）", "倍", {c: f"={c}«C.infcap»/{c}«C.eff»" for c in ALL}, "Revenue 容量上限列引用（只回乘雲端營收）", key="cap", name="CMP_CapFactor")
    C.add("缺口旗標（1＝需求超過可用 GW）", "旗標", {c: f"=({c}«C.eff»>{c}«C.infcap»)" for c in ALL}, "不自動假設新增算力", key="flag", name="CMP_CapFlag")
    C.add("研發 GW（殘差）", "GW",
          lambda c: wavg("rd") if c == "E" else (f"={c}«C.sup»-{c}«C.infcap»" if c in cal else f"={c}«C.sup»*(1-{idle})-{c}«C.infcap»"),
          "校準期＝供給 − 推論（＝研發算力費 ÷ 每 GW 價格）；之後＝供給 ×（1−閒置）− 推論（截頂後）", key="rd", name="CMP_RDGW")
    C.add("總需求 GW＝有效推論（未截頂）＋研發", "GW", {c: f"={c}«C.eff»+{c}«C.rd»" for c in ALL}, "", key="dem", name="CMP_DemandGW")
    C.add("檢查：推論（截頂後）＋研發＋閒置 − 供給", "GW", {c: f"={c}«C.infcap»+{c}«C.rd»+{c}«C.idle»-{c}«C.sup»" for c in ALL}, "各期應為 0", key="ident")
    C.add("對照：研發 GW（扣除法研發算力費 ÷ 每 GW 價格）", "GW", {c: f"={c}«C.f_rd_a»/{per(c, f'{c}«C.p»')}" for c in cal},
          "D12：智譜揭露研發算力費，可作獨立對照（D10r 第一法）", key="rd_a")
    C.add("對照：TK IF_AllocRDGW（OpenAI 規模研發 GW·年）", "GW·年", {"D": "=TK_IF_AllocRDGW"}, "Tokenomics Alloc（k ×（N_major × 家族＋N_refresh × 改版））；只並列", key="rd_tk")

    # 十、VR 等值
    C.section("十、VR 等值換算（各族 Sol 層級每 GW 產能 ÷ VR200 Sol；同 OpenAI v0.6）")
    for f in FAM:
        C.add(f"VR200 等值係數：{FAM_ZH[f]}", "倍", {c: f"={c}«C.cap_{f}_sol»/{c}«C.tk_vr_sol»" for c in ALL}, "", key=f"vr_{f}")
    C.add("機隊 VR 等值係數＝Σ 占比 × 係數", "倍", {c: "=" + "+".join(f"{c}«C.s_{f}»*{c}«C.vr_{f}»" for f in FAM) for c in ALL}, "",
          key="vrf", name="CMP_VReqFactor")
    C.add("供給 VR 等值 GW（命題分母）", "GW", lambda c: wavg("vr_sup") if c == "E" else f"={c}«C.sup»*{c}«C.vrf»", "", key="vr_sup", name="CMP_Supply_VReq")
    C.add("推論 GW（截頂後）VR 等值", "GW", lambda c: wavg("vr_inf") if c == "E" else f"={c}«C.infcap»*{c}«C.vrf»", "", key="vr_inf", name="CMP_InfGW_VReq")
    C.add("研發 GW VR 等值", "GW", lambda c: wavg("vr_rd") if c == "E" else f"={c}«C.rd»*{c}«C.vrf»", "", key="vr_rd", name="CMP_RDGW_VReq")

    # 十一、對照：10 萬級國產晶片、1 GW 資料中心
    C.section("十一、對照：『10 萬級國產晶片』換算 GW（卡數 × 每卡 IT 功耗）與『1 GW 國產算力資料中心』（只列差距，不校準）")
    C.add("公司揭露：推論用國產晶片（張）", "張", {"L": f"={S('cmp_chips_domestic_100k')}"}, "SRC_ZP_452（截至 2026-08-31；Interested-party；下限）", key="chips")
    C.add("換算 IT GW＝卡數 × 每卡 IT 功耗 ÷ 10^6", "GW", {"L": f"=L«C.chips»*L«C.kw_dom»/{I('kw_per_gw')}"}, "每卡 IT 功耗 Inputs（0.65–1.4 kW）", key="chip_gw",
          name="CMP_ChipGW", name_col="L")
    C.add("換算設施 GW（× PUE）", "GW", {"L": f"=L«C.chip_gw»*{I('pue')}"}, "", key="chip_gw_f")
    C.add("模型：2H26 國產供給 GW＝供給 × 國產占比", "GW", {c: f"={c}«C.sup»*{c}«C.s_dom»" for c in ALL}, "", key="dom_gw", name="CMP_DomesticGW")
    C.add("模型：2H26 國產推論 GW＝推論（截頂後）× 國產占比", "GW", {c: f"={c}«C.infcap»*{c}«C.s_dom»" for c in ALL}, "國產晶片用於推論（公司說法）", key="dom_inf")
    C.add("差距：模型 2H26 國產供給 GW ÷ 換算 GW − 1", "比例", {"L": "=(L«C.dom_gw»-L«C.chip_gw»)/L«C.chip_gw»"}, "Checks WARN 門檻 50%", key="chip_gap",
          name="CMP_ChipGWGap", name_col="L")
    C.add("差距：模型 2H26 國產推論 GW ÷ 換算 GW − 1", "比例", {"L": "=(L«C.dom_inf»-L«C.chip_gw»)/L«C.chip_gw»"}, "只列", key="chip_gap_inf")
    C.add("公司說法：1 GW 國產算力資料中心 ÷ PUE（IT GW）", "GW", {"L": f"={S('cmp_1gw_dc')}/{I('pue')}"}, "SRC_ZP_458（口徑未明，視為設施）；D11r 只列情境", key="dc_gw")

    # 十二、R4：Coding Plan token 對帳
    C.section("十二、R4 對帳：Coding Plan token 換算 GW 與 API 銷售成本推得的推論 GW（2025、1H26；不自動調整，見報告）")
    C.add("Coding Plan token（T）", "T", {c: f"=Demand!{c}{dr['cptok']}" for c in cal}, "Demand DEM_Tok_CP", key="r4_tok")
    C.add("Coding Plan token 換算 GW（以付費組合後產能）", "GW", {c: f"=Demand!{c}{dr['cptok']}*{I('mul_t_m')}/{per(c, f'{c}«C.pc_paid»')}" for c in cal}, "", key="r4_gw")
    C.add("Coding Plan 占 token 換算推論 GW", "比例", {c: f"={c}«C.r4_gw»/{c}«C.g_tok»" for c in cal}, "", key="r4_sh")
    C.add("倍數：Coding Plan token GW ÷ 支出換算推論 GW", "倍", {c: f"={c}«C.r4_gw»/{c}«C.e_spend»" for c in cal}, ">1＝Coding Plan 單項已超過推論算力支出買得到的 GW", key="r4_mult",
          name="CMP_R4Mult")
    C.add("對照：使 token GW＝支出 GW 的 Coding Plan 額度使用率", "比例",
          {c: f"={I('cp_util')}*({c}«C.e_spend»-({c}«C.g_tok»-{c}«C.r4_gw»))/{c}«C.r4_gw»" for c in cal},
          "＝Inputs 使用率 ×（支出 GW − 非 Coding Plan token GW）÷ Coding Plan token GW；負值＝其他 token 已超過；只對照，不回饋（R4）", key="r4_util")

    # 十三、敏感度
    C.section("十三、敏感度（Excel 公式；引用 Inputs 低／高欄與 SRC 區間；K＝1H26、I＝2030）：A 每 GW 價格、B η 路徑、C 國產產出比")
    pr = {"lo": {"hop": S("rent_h800_cardhr_turing_b"), "h20": S("rent_h20_cardhr_sep26", "lo"), "dom": S("rent_910b4_cardhr_derived")},
          "hi": {"hop": S("rent_h800_cardhr_aliyun"), "h20": S("rent_h20_cardhr_sep26", "hi"), "dom": S("rent_910b_zxy_cardhr_derived")}}
    base_r = {"hop": S("rent_h800_cardhr_derived"), "h20": S("rent_h20_cardhr_sep26"),
              "dom": f"AVERAGE({S('rent_910b4_cardhr_derived')},{S('rent_910b_zxy_cardhr_derived')})"}
    two = ("K", "I")
    for tag, zh in (("lo", "低端"), ("hi", "高端")):
        k = f"A_{tag}"
        C.add(f"A 每 GW 價格（{zh}租價）", "RMB 億／GW／年",
              {c: "=" + "+".join(f"{c}«C.s_{f}»*{c}«C.p_{f}»*{pr[tag][f]}/{base_r[f]}" for f in FAM) for c in two},
              "各族租價取區間端點（Hopper 11.83／18.75；H20 7.03／8.16；910B 2.74／4.28）", key=f"{k}_p")
        C.add(f"A η 1H26（{zh}）", "倍", {"K": f"=K«C.g_tok»/(K«C.f_inf»/(K«C.{k}_p»*{frac('K')}))"}, "", key=f"{k}_eta")
        C.add(f"A 供給 GW（{zh}）", "GW",
              {"K": f"=K«C.f_tot»/(K«C.{k}_p»*{frac('K')})",
               "I": f"=I«C.g_tok»/$K$«C.{k}_eta»/((1-{idle})*(1-I«C.rdsh»))"}, "1H26＝算力費 ÷ 價格；2030＝有效推論 ÷［(1−閒置)(1−研發占比)］（η 沿用）", key=f"{k}_sup")
        C.add(f"A 供給 VR 等值 GW（{zh}）", "GW", {c: f"={c}«C.{k}_sup»*{c}«C.vrf»" for c in two}, "", key=f"{k}_vr")
    C.add("B η 回升至 2030＝目標：2030 有效推論 GW", "GW", {"I": f"=I«C.g_tok»/{tgt}"}, "η 線性回升情境（開關＝1）", key="B_eff")
    C.add("B η 回升：2030 供給 GW", "GW", {"I": f"=I«C.B_eff»/((1-{idle})*(1-I«C.rdsh»))"}, "", key="B_sup")
    C.add("B η 回升：2030 供給 VR 等值 GW", "GW", {"I": "=I«C.B_sup»*I«C.vrf»"}, "", key="B_vr")
    for tag, col, zh in (("lo", "G", "低端"), ("hi", "H", "高端")):
        k = f"C_{tag}"
        ralt = D.I_cell("dom_ratio", col)
        C.add(f"C 國產產出比（{zh}）", "倍", {c: f"={ralt}" for c in two}, f"Inputs 國產產出比的{zh}欄", key=f"{k}_r")
        C.add(f"C 組合產出係數（{zh}）", "倍", {c: f"={c}«C.s_hop»+{c}«C.s_h20»*{c}«C.r_h20»+{c}«C.s_dom»*{c}«C.{k}_r»" for c in two}, "", key=f"{k}_mixf")
        C.add(f"C token 換算 GW（{zh}）", "GW", {c: f"={c}«C.g_tok»*{c}«C.mixf»/{c}«C.{k}_mixf»" for c in two}, "各族產能同比例變動：token GW ∝ 1 ÷ 組合產出係數", key=f"{k}_g")
        C.add(f"C η 1H26（{zh}）", "倍", {"K": f"=K«C.{k}_g»/K«C.e_spend»"}, "", key=f"{k}_eta")
        C.add(f"C 2030 供給 GW（{zh}）", "GW", {"I": f"=I«C.{k}_g»/$K$«C.{k}_eta»/((1-{idle})*(1-I«C.rdsh»))"}, "", key=f"{k}_sup")
        C.add(f"C 2030 供給 VR 等值 GW（{zh}）", "GW",
              {"I": f"=I«C.{k}_sup»*(I«C.s_hop»*I«C.vr_hop»+I«C.s_h20»*I«C.vr_h20»+I«C.s_dom»*I«C.vr_hop»*I«C.{k}_r»)"}, "", key=f"{k}_vr")

    # ══════════════════════════════════ Cost ══════════════════════════════════
    K = Sheet(wb, "Cost", "K",
              "Cost — 成本與命題表：算力成本（2025、1H26＝算力服務費實付；之後＝租用 GW × 每 GW 價格＋自有資本支出）、供應商持有成本與雲端毛利、本地化部署交付成本、非算力成本、股權報酬、每 VR 等值 GW 命題表（人民幣與美元）",
              "規格 D4、D10、D11r、D13、D23。全成本＝算力成本＋本地化部署交付成本＋非算力成本（不含股權報酬）＋股權報酬（命題含股權報酬，較保守；另列不含者）。"
              "2025 全成本對財報（銷售成本＋研發＋銷售＋管理）差距＝0（非算力成本為財報減項的殘差）。",
              "金額 RMB 億；命題表 RMB 億／VR 等值 GW／年，美元口徑 $B／VR 等值 GW／年（÷ USD/CNY ÷ 10）。2026（E）＝1H26＋2H26。")
    K.years(D)
    kr = K.rows
    cmp_ = lambda key, c: f"Compute!{c}«C.{key}»"  # noqa: E731

    K.section("一、算力成本（D10）：2025、1H26＝算力服務費實付（Compute 第五節）；2H26 起＝租用 GW × 每 GW 價格 × 期間比例＋自有資本支出")
    K.add("推論算力支出（實付）", "RMB 億", {c: f"={cmp_('f_inf', c)}" for c in cal}, "", key="inf_fee")
    K.add("研發算力支出（實付）", "RMB 億", {c: f"={cmp_('f_rd', c)}" for c in cal}, "", key="rd_fee")

    def rentpay(c):
        if c in cal:
            return f"={c}«K.inf_fee»+{c}«K.rd_fee»"
        if c == "E":
            return "=K«K.rentpay»+L«K.rentpay»"
        return f"={cmp_('rent_gw', c)}*{per(c, cmp_('p', c))}"
    K.add("算力服務費（租用）", "RMB 億", rentpay, "2H26 起＝租用 GW × 每 GW 年價格 × 期間比例", key="rentpay", name="COST_ComputeRent")
    K.add("自有資本支出（現金）", "RMB 億", {c: f"={cmp_('cpx', c)}" for c in LATE}, "Compute 第八節（D11r：基準 0）", key="own")
    K.add("算力成本（現金口徑）＝租用＋自有資本支出", "RMB 億",
          lambda c: f"={c}«K.rentpay»+{c}«K.own»" if c in LATE else ("=K«K.cc»+L«K.cc»" if c == "E" else f"={c}«K.rentpay»"),
          "命題輸出基準", key="cc", name="COST_Compute")
    for key_, zh, g in (("cc_inf", "推論", "infcap"), ("cc_rd", "研發", "rd"), ("cc_idle", "閒置", "idle")):
        K.add(f"算力成本拆分：{zh}", "RMB 億",
              lambda c, g=g, key_=key_: f"=K«K.{key_}»+L«K.{key_}»" if c == "E" else f"={c}«K.cc»*{cmp_(g, c)}/{cmp_('sup', c)}",
              "依 GW 占比拆分（校準期：推論＝推論算力支出、研發＝研發算力支出）", key=key_, name={"cc_inf": "COST_InfCompute", "cc_rd": "COST_RDCompute"}.get(key_))
    K.add("檢查：算力成本 −（實付推論＋研發）", "RMB 億", {c: f"={c}«K.cc»-{c}«K.inf_fee»-{c}«K.rd_fee»" for c in cal}, "2025、1H26 應為 0（Checks）", key="ck_fee")

    K.section("二、供應商持有成本與雲端毛利（D10：持有成本＝租用 GW × TK IF_HoldEcon（Hopper）× 持有比 × 匯率；差額＝雲端毛利，只作參考）")
    hr = {"hop": None, "h20": I("hold_ratio_h20"), "dom": I("hold_ratio_dom")}
    for f in FAM:
        base_h = f"SUMIFS(TK_IF_HoldEcon,TK_HdrGen,Compute!{kh},TK_HdrCost,Compute!{kc})*{fx}*{bn}"
        K.add(f"每 GW 年持有成本：{FAM_ZH[f]}", "RMB 億／GW／年", {c: f"={base_h}" + (f"*{hr[f]}" if hr[f] else "") for c in ALL},
              "TK IF_HoldEcon（Hopper 基準欄，經濟口徑，$B/GW/年）× 持有比（Inputs；Hopper＝1）× USD/CNY × 10", key=f"h_{f}")
    K.add("每 GW 年持有成本：組合後", "RMB 億／GW／年", {c: "=" + "+".join(f"{cmp_('s_' + f, c)}*{c}«K.h_{f}»" for f in FAM) for c in ALL}, "", key="hw", name="COST_HoldW")
    K.add("供應商持有成本＝租用 GW × 持有成本 × 期間比例", "RMB 億",
          lambda c: "=K«K.sh»+L«K.sh»" if c == "E" else f"={cmp_('rent_gw', c)}*{per(c, f'{c}«K.hw»')}", "", key="sh", name="COST_SupplierHold")
    K.add("雲端毛利＝算力服務費 − 供應商持有成本", "RMB 億", {c: f"={c}«K.rentpay»-{c}«K.sh»" for c in ALL}, "", key="gm", name="COST_CloudGM")
    K.add("雲端毛利率", "比例", {c: f"=IF({c}«K.rentpay»,{c}«K.gm»/{c}«K.rentpay»,{c}«K.rentpay»)" for c in ALL}, "租用為 0（全自有）時顯示 0", key="gmp", name="COST_CloudGMPct")

    K.section("三、本地化部署交付成本（D4：不耗智譜算力；2025、1H26＝財報推算；2H26 起＝本地化營收 × 1H26 成本率）")
    K.add("本地化部署銷售成本（推算）", "RMB 億", sv("cogs_onprem_fy25", "cogs_onprem_1h26"), "SRC_ZP_198、196（Derived）", key="op_src")
    K.add("本地化部署成本率", "比例", {c: f"={c}«K.op_src»/Revenue!{c}{vr['onprem']}" for c in cal}, "2025 0.51、1H26 0.62；2H26 起沿用 1H26（最新實際值，較保守）", key="op_ratio")

    def opd(c):
        if c in cal:
            return f"={c}«K.op_src»"
        if c == "E":
            return "=K«K.opd»+L«K.opd»"
        return f"=Revenue!{c}{vr['onprem']}*$K$«K.op_ratio»"
    K.add("本地化部署交付成本", "RMB 億", opd, "", key="opd", name="COST_OnPremDelivery")

    K.section("四、非算力成本（D13）：2025、1H26＝財報費用 − 算力服務費 − 本地化交付成本 − 股權報酬（殘差）；2H26 起＝人數 × 每人年成本；股權報酬單列")
    K.add("銷售成本", "RMB 億", sv("cogs_fy25", "cogs_1h26"), "SRC_ZP_123、182", key="x_cogs")
    K.add("研發開支", "RMB 億", sv("rd_fy25", "rd_1h26"), "SRC_ZP_139、199", key="x_rd")
    K.add("銷售及營銷開支", "RMB 億", sv("sm_fy25", "sm_1h26"), "SRC_ZP_140、200", key="x_sm")
    K.add("一般及行政開支", "RMB 億", sv("ga_fy25", "ga_1h26"), "SRC_ZP_141、201", key="x_ga")
    K.add("財報費用合計（銷售成本＋研發＋銷售＋管理）", "RMB 億", {c: f"={c}«K.x_cogs»+{c}«K.x_rd»+{c}«K.x_sm»+{c}«K.x_ga»" for c in cal}, "", key="x_tot", name="COST_Reported")
    K.add("股權報酬（財報）", "RMB 億", sv("sbc_fy25", "sbc_1h26"), "SRC_ZP_142、202", key="x_sbc")
    K.add("非算力成本（不含股權報酬；財報殘差）", "RMB 億", {c: f"={c}«K.x_tot»-{c}«K.cc»-{c}«K.opd»-{c}«K.x_sbc»" for c in cal},
          "＝研發、銷售、管理扣算力服務費與股權報酬，加 API 銷售成本中非算力部分（推論占比 <1 時）", key="nc_cal")
    hcg, ppg, sbg = I("hc_g_2h26"), I("pp_g"), I("sbc_pp_g")

    def hc(c):
        if c == "D":
            return f"=AVERAGE({S('employees_1h25')},{S('employees_fy25')})"
        if c == "K":
            return f"=AVERAGE({S('employees_fy25')},{S('employees_1h26')})"
        if c == "L":
            return f"=K«K.hc»*(1+{hcg})"
        if c == "E":
            return f"=(K«K.hc»+L«K.hc»)/{halves}"
        return f"={PREV[c]}«K.hc»*(1+{I('hc_g')})"
    K.add("員工人數（期間平均）", "人", hc, "2025＝AVERAGE（2025-06-30 883, 2025-12-31 1,094）；1H26＝AVERAGE（1,094, 2026-06-30 981）（SRC_ZP_223、162、222）；之後 Inputs 成長率",
          key="hc", name="COST_Headcount")

    def perhead(src_key, g):
        def f(c):
            if c == "D":
                return f"=D«K.{src_key}»/D«K.hc»"
            if c == "K":
                return f"=K«K.{src_key}»/(K«K.hc»*{frac('K')})"
            if c == "L":
                return f"=K«K.{src_key}_pp»*(1+{g}/{halves})"
            if c == "E":
                return f"=(K«K.{src_key}_pp»*{d1}+L«K.{src_key}_pp»*{d2})/{dy}"
            return f"={PREV[c]}«K.{src_key}_pp»*(1+{g})"
        return f
    K.add("每人年成本（不含股權報酬）", "RMB 億／人／年", perhead("nc_cal", ppg), "校準期＝財報殘差 ÷ 人數（1H26 年化）；2H26 起 ×（1＋年變動率；2H26 為半年）", key="nc_cal_pp")
    K.add("每人年股權報酬", "RMB 億／人／年", perhead("x_sbc", sbg), "1H26 年化為基期（上市後）；2025 含上市前授予，只列", key="x_sbc_pp")

    def bypp(src_key):
        def f(c):
            if c in cal:
                return f"={c}«K.{src_key}»"
            if c == "E":
                return f"=K«K.{src_key}_m»+L«K.{src_key}_m»"
            return f"={c}«K.hc»*{per(c, f'{c}«K.{src_key}_pp»')}"
        return f
    K.add("非算力成本（不含股權報酬）", "RMB 億", bypp("nc_cal"), "校準期＝財報殘差；2H26 起＝人數 × 每人年成本 × 期間比例", key="nc_cal_m", name="COST_NonCompExSBC")
    K.add("股權報酬", "RMB 億", bypp("x_sbc"), "校準期＝財報；2H26 起＝人數 × 每人股權報酬 × 期間比例", key="x_sbc_m", name="COST_SBC")

    K.section("五、全成本與對帳（命題全成本含股權報酬；2025、1H26 對財報差距應為 0）")
    K.add("全成本（含股權報酬）＝算力＋本地化交付＋非算力＋股權報酬", "RMB 億",
          {c: f"={c}«K.cc»+{c}«K.opd»+{c}«K.nc_cal_m»+{c}«K.x_sbc_m»" for c in ALL}, "", key="full", name="COST_FullCash")
    K.add("全成本（不含股權報酬）", "RMB 億", {c: f"={c}«K.full»-{c}«K.x_sbc_m»" for c in ALL}, "", key="fullx", name="COST_FullExSBC")
    K.add("對帳：全成本 − 財報費用合計（應為 0）", "RMB 億", {c: f"={c}«K.full»-{c}«K.x_tot»" for c in cal}, "工作單 Z3 第 2 步：2025 差距＝0（1H26 同）", key="recon", name="COST_Recon")
    K.add("營收淨額（截頂後）", "RMB 億", {c: f"=Revenue!{c}{vr['net_c']}" for c in YC}, "REV_NetCapped", key="rev")
    K.add("雲端營收（截頂後）＝雲端 × 容量上限係數", "RMB 億", {c: f"=Revenue!{c}{vr['cloud']}*Revenue!{c}{vr['cap']}" for c in YC}, "", key="rev_cl", name="COST_CloudRevCapped")
    K.add("差額＝營收淨額 − 全成本（含股權報酬）", "RMB 億", {c: f"={c}«K.rev»-{c}«K.full»" for c in YC}, "負＝缺口（Z4 融資）", key="gap", name="COST_GapCash")
    K.add("對照：財報年內虧損（2025）", "RMB 億", {"D": f"=-{S('net_loss_fy25')}"}, "SRC_ZP_145；差額另含其他收入、財務成本、投資人金融工具公允值變動等", key="nl")

    K.section("六、命題表（人民幣；每 VR 等值 GW；分母＝供給 VR 等值 GW）")
    K.add("分母：供給 VR 等值 GW", "GW", {c: f"={cmp_('vr_sup', c)}" for c in YC}, "CMP_Supply_VReq", key="den")
    prop = [("p_rev", "營收淨額", "rev", "COST_PropRev_VR"), ("p_cl", "雲端營收", "rev_cl", "COST_PropCloudRev_VR"),
            ("p_cc", "算力成本", "cc", "COST_PropCompute_VR"), ("p_op", "本地化部署交付成本", "opd", "COST_PropOnPrem_VR"),
            ("p_nc", "非算力成本（不含股權報酬）", "nc_cal_m", "COST_PropNonComp_VR"), ("p_sbc", "股權報酬", "x_sbc_m", "COST_PropSBC_VR"),
            ("p_full", "全成本（含股權報酬）", "full", "COST_PropFull_VR"), ("p_fullx", "全成本（不含股權報酬）", "fullx", "COST_PropFullExSBC_VR")]
    for key_, zh, src, nm_ in prop:
        K.add(f"每 VR 等值 GW：{zh}", "RMB 億／GW／年", {c: f"={c}«K.{src}»/{c}«K.den»" for c in YC}, "", key=key_, name=nm_)
    K.add("每 VR 等值 GW：差額（含股權報酬）", "RMB 億／GW／年", {c: f"={c}«K.p_rev»-{c}«K.p_full»" for c in YC}, "命題：負＝營收不足以覆蓋全成本", key="p_gap", name="COST_PropGap_VR")
    K.add("每 VR 等值 GW：差額（不含股權報酬）", "RMB 億／GW／年", {c: f"={c}«K.p_rev»-{c}«K.p_fullx»" for c in YC}, "", key="p_gapx", name="COST_PropGapExSBC_VR")
    K.add("每 VR 等值 GW：雲端口徑差額＝雲端營收 −（全成本 − 本地化交付成本）", "RMB 億／GW／年", {c: f"={c}«K.p_cl»-{c}«K.p_full»+{c}«K.p_op»" for c in YC},
          "D4：分子只計雲端營收、分母不變；本地化交付成本隨本地化營收一併排除（其餘成本全留，較保守）", key="p_gap_cl", name="COST_PropGap_VR_Cloud")
    K.add("覆蓋率＝營收淨額 ÷ 全成本（含股權報酬）", "倍", {c: f"={c}«K.rev»/{c}«K.full»" for c in YC}, "", key="p_cov", name="COST_Coverage")
    K.add("雲端覆蓋率＝雲端營收 ÷（全成本 − 本地化交付成本）", "倍", {c: f"={c}«K.rev_cl»/({c}«K.full»-{c}«K.opd»)" for c in YC}, "", key="p_cov_cl", name="COST_CoverageCloud")

    K.section("七、命題表（美元；$B／VR 等值 GW／年＝人民幣億 ÷ USD/CNY ÷ 10；同一匯率，D23）")
    for key_, zh, nm_ in (("p_rev", "營收淨額", "COST_PropRev_VR_USD"), ("p_cl", "雲端營收", None), ("p_cc", "算力成本", "COST_PropCompute_VR_USD"),
                          ("p_full", "全成本（含股權報酬）", "COST_PropFull_VR_USD"), ("p_gap", "差額（含股權報酬）", "COST_PropGap_VR_USD"),
                          ("p_gapx", "差額（不含股權報酬）", None), ("p_gap_cl", "雲端口徑差額", "COST_PropGap_VR_Cloud_USD")):
        K.add(f"每 VR 等值 GW：{zh}（美元）", "$B／GW／年", {c: f"={c}«K.{key_}»/{fx}/{bn}" for c in YC}, "", key=f"u_{key_}", name=nm_)
    K.add("營收淨額（美元）", "$B", {c: f"={c}«K.rev»/{fx}/{bn}" for c in YC}, "", key="u_rev")
    K.add("全成本（含股權報酬；美元）", "$B", {c: f"={c}«K.full»/{fx}/{bn}" for c in YC}, "", key="u_full")
    K.add("差額（美元）", "$B", {c: f"={c}«K.gap»/{fx}/{bn}" for c in YC}, "", key="u_gap", name="COST_GapCash_USD")

    K.section("八、Tokenomics 單位成本參考（每 VR200 GW 年全成本＝TK $/M × 每 GW 年產出（Sol）× 基準利用率；只作參考，算力成本仍＝實付）")
    for nm_, zh, key_ in (("IF_FullCostDefault_Sol", "下游預設 IF_FullCostDefault_Sol；基準", "tk_def"), ("IF_FullCost_Sol", "IF_FullCost_Sol；並列", "tk_full")):
        K.add(f"TK 參考：每 VR200 GW 年全成本（{zh}）", "RMB 億／GW／年",
              {c: f"=SUMIFS(TK_{nm_},TK_HdrGen,Compute!{kv},TK_HdrCost,Compute!{kc})*SUMIFS(TK_IF_TokGW_Sol,TK_HdrGen,Compute!{kv},TK_HdrCost,Compute!{kc})"
                  f"*TK_IF_Util/{I('div_usd_b')}*{fx}*{bn}" for c in YC}, "TK $/M（VR200 Sol 基準欄）× 每 GW 年產出 × IF_Util ÷ 10^9 × USD/CNY × 10", key=key_)
    K.add("差距：每 VR 等值 GW 算力成本 − TK 參考（下游預設）", "RMB 億／GW／年", {c: f"={c}«K.p_cc»-{c}«K.tk_def»" for c in YC}, "", key="tk_gap")

    for S_, w in ((C, 62), (K, 62)):
        ws = S_.ws
        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = w
        ws.column_dimensions["C"].width = 14
        for c in YC + HC:
            ws.column_dimensions[c].width = 12
        ws.column_dimensions["J"].width = 2
        ws.column_dimensions["M"].width = 80
        ws.freeze_panes = "D7"
    return dict(C=C, K=K)


def checks(D, Z):
    """Z3 的 Checks 列。kind：eq／tol／warn（ARR）／warnrng（(lo 名, hi 名)）／warnabs（門檻名）／info。«C.»／«K.» 由 fill() 代換。"""
    I = D.I
    c8 = ("D", "E", "F", "G", "H", "I", "K", "L")
    shares = "+".join(f"(Compute!{c}«C.s_{f}»<0)+(Compute!{c}«C.s_{f}»>1)" for c in c8 for f in FAM)
    supinf = "+".join(f"(Compute!{c}«C.sup»<Compute!{c}«C.infcap»)" for c in c8)
    rdneg = "+".join(f"(Compute!{c}«C.rd»<0)" for c in c8)
    ident = "MAX(" + ",".join(f"ABS(Compute!{c}«C.ident»)" for c in c8) + ")"
    capv = "+".join(f"(Compute!{c}«C.cap»<=0)+(Compute!{c}«C.cap»>1)" for c in c8)
    revcap = "+".join(f"(Revenue!{c}«V.gross_c»>Revenue!{c}«V.gross»)" for c in YC)
    ncneg = "(Cost!D«K.nc_cal»<0)+(Cost!K«K.nc_cal»<0)"
    rows = [
        ("Compute：晶片族 GW 占比 ∉[0,1] 的格數（Hopper＝1 − 國產 − H20）", "=" + shares, 0, "eq", "Inputs 國產＋H20 ≤ 1"),
        ("Compute：供給 GW < 推論 GW（截頂後）的期數", "=" + supinf, 0, "eq", "工作單 Z3 第 3 步：供給 ≥ 推論"),
        ("Compute：研發 GW（殘差）< 0 的期數", "=" + rdneg, 0, "eq", "研發 GW 非負旗標（校準期：研發算力費 < 0 或推論超過供給時轉 ERR）"),
        ("Compute：推論（截頂後）＋研發＋閒置 − 供給（各期最大絕對值，GW）", "=" + ident, 0, "tol", "恆等式"),
        ("Compute：容量上限係數 ∉(0,1] 的期數", "=" + capv, 0, "eq", ""),
        ("Revenue：截頂後營收總額 > 未截頂的年數", "=" + revcap, 0, "eq", "容量上限只能使營收下降"),
        ("Cost：2025 算力成本 −（推論＋研發算力服務費）（RMB 億）", "=Cost!D«K.ck_fee»", 0, "tol", "工作單：2025 算力費對帳"),
        ("Cost：1H26 算力成本 −（推論＋研發算力服務費）（RMB 億）", "=Cost!K«K.ck_fee»", 0, "tol", "工作單：1H26 算力費對帳"),
        ("Cost：2025 全成本 − 財報費用合計（銷售成本＋研發＋銷售＋管理）（RMB 億）", "=Cost!D«K.recon»", 0, "tol", "工作單 Z3 第 2 步：差距＝0"),
        ("Cost：1H26 全成本 − 財報費用合計（RMB 億）", "=Cost!K«K.recon»", 0, "tol", ""),
        ("Cost：校準期非算力成本（財報殘差）< 0 的期數", "=" + ncneg, 0, "eq", "算力服務費估計（推論占比、研發算力占比）超過財報費用時轉 ERR"),
        ("η 2025 ∉［下限, 上限］", "=CMP_Eta2025", ("eta_warn_lo", "eta_warn_hi"), "warnrng", "η 合理範圍 WARN（Inputs 0.1–1）；>1＝模型 token 量超過推論算力支出以 TK 產能買得到的量（見報告 R4）"),
        ("η 1H26 ∉［下限, 上限］", "=CMP_Eta1H26", ("eta_warn_lo", "eta_warn_hi"), "warnrng", "同上"),
        ("國產晶片 GW 對照：模型 2H26 國產供給 GW ÷『10 萬級國產晶片』換算 GW − 1", "=CMP_ChipGWGap", "chip_gap_warn", "warnabs", "絕對值 >50% 為 WARN；只對照"),
        ("對照：Coding Plan token GW ÷ 支出換算推論 GW（2025）", "=Compute!D«C.r4_mult»", None, "info", "R4"),
        ("對照：Coding Plan token GW ÷ 支出換算推論 GW（1H26）", "=Compute!K«C.r4_mult»", None, "info", "R4"),
        ("對照：組合後每 GW 年價格 1H26（RMB 億）", "=Compute!K«C.p»", None, "info", "D9r"),
        ("對照：研發 GW 1H26（殘差）÷ 扣除法研發算力費換算 GW", "=Compute!K«C.rd»/Compute!K«C.rd_a»", None, "info", "D12 獨立對照"),
        ("對照：雲端毛利率 2025（供應商）", "=Cost!D«K.gmp»", None, "info", ""),
        ("預覽：2026 供給 GW", "=INDEX(CMP_SupplyGW,1,2)", None, "info", ""),
        ("預覽：2030 供給 GW", "=INDEX(CMP_SupplyGW,1,6)", None, "info", ""),
        ("預覽：2030 每 VR 等值 GW 差額（RMB 億）", "=INDEX(COST_PropGap_VR,1,6)", None, "info", "命題"),
        ("預覽：2030 每 VR 等值 GW 雲端口徑差額（RMB 億）", "=INDEX(COST_PropGap_VR_Cloud,1,6)", None, "info", ""),
        ("預覽：2030 覆蓋率", "=INDEX(COST_Coverage,1,6)", None, "info", ""),
        ("對照：2030 每 VR 等值 GW 差額：智譜（$B）− OpenAI v0.6（$B）", "=INDEX(COST_PropGap_VR_USD,1,6)-INDEX(OAI_COST_PropGap_VR,1,6)", None, "info", "OAI_Link；只並排"),
    ]
    return rows
