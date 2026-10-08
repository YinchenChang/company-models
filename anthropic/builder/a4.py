"""Anthropic v0.1（A4）：融資頁 Funding 與反向模式頁 Reverse，及 Cost 頁的 A4 補充（工作單 A4；規格 D16 r1、D17、D18、D19、D21、D22）。

結構沿用 OpenAI v0.6 builder/p5.py（同名具名範圍：FND_*、RVS_*），差異：
- 自由現金流＝營收淨額（截頂後）− 算力成本（現金）− 非算力成本（不含股權報酬；股權報酬為非現金，加回）；2025 淨損中約 $34B 非現金費用不計（D21）。
- 來源順序（D16 r1；工作單 A4 預設）：①2025 年底現金 → ②2026 已宣布且已交割的股權（Series G $30B、Amazon Series G 特別股 $5B、Series H $65B〔含 $15B 先前承諾的雲端業者投資〕；
  Microsoft $5B 已在 2025 年底現金內、Google 2026-04 $10B 視為 Series H 的 $15B 之一，皆以開關避免重複）→ ③外部資金（補足至最低現金）。
  條件式（Amazon $15B 里程碑、Google $30B 績效、NVIDIA 最高 $10B、AMD 最高 $5B 認股）與 IPO 約 $100B 只列情境；循環信貸 $15B、Broadcom 租賃融資、表外 TPU 融資只列或有。
- 最低現金（D17）＝次年非算力成本（不含股權報酬）× 月數 ÷ 12。
- 回流對照（D18）：策略投資人的投資額 vs 對同一方的算力合約，只列對照、不沖銷。
- Cost 頁補充（chat 端 A3 審查）：逐家合約總額對招股書（WARN）、「逐家錨定招股書金額」敏感度、機房租約租金敏感度、D22 非算力低估量化。
- Reverse（D19）：管理層營收目標 → 所需倍數與反向資金（同一支出）；只被 Checks 引用（test_reverse_not_fed_back）。
Python 只產生結構；數值一律引用 SRC_ANT／Inputs／REV_／COST_ 具名範圍或同簿儲存格（E6：公式不含常數）。
"""
from __future__ import annotations

from a2 import FMT, YC, YEARS, Sheet
from common import put

FMT.update({"開關": "0", "混合": "#,##0.00"})

# 招股書六家合作方 → 本模型合約鍵（Cost 第十一節 tot_*／pay_*）；SRC 鍵與欄（與 a3.PARTNERS 同）
PARTNER_MAP = [("Google", "google", "cmp_prospectus_google", "_Lo"), ("Amazon", "aws", "cmp_prospectus_amazon", ""),
               ("Microsoft", "azure", "cmp_msft_azure_prospectus", ""), ("Broadcom", "broadcom", "cmp_prospectus_broadcom", ""),
               ("xAI／SpaceX", "spacex", "cmp_spacex_total_prospectus", "_Hi"), ("AMD", "amd", "cmp_amd_usd_prospectus", "_Lo")]
NVDA_FAMILY = ("azure", "spacex", "nscale", "lambda", "volta", "akamai")     # 以 NVIDIA 晶片為主的合約（D18 間接對照）
LEASE_RENT = [("terawulf", "TeraWulf Hawesville（KY；20 年 $19B；Anthropic 直接承租）", "cmp_terawulf_ky_usd", "lease_term_terawulf"),
              ("riot", "Riot Rockdale（TX；20 年 $9.1B）", "cmp_riot_rockdale_usd", "lease_term_riot"),
              ("hut8", "Hut 8 River Bend（LA；15 年 $7B；經 Fluidstack，可能已含於自建 $50B）", "cmp_hut8_riverbend_usd", "lease_term_hut8")]


def waterfall(S, tag, zh, fcf, inflow, cash25, minref, names=None, note=""):
    """現金瀑布（同 S7 規則）：補足前＝期初＋自由現金流＋股權流入；外部資金＝補足至最低現金；年底＝補足前＋外部資金；累計。
    fcf：記號字串（例如 «fcf»）；inflow：callable(i, c) → 運算式（不含 =）；names：{'ext':…, 'end':…, 'cum':…, 'pre':…}。"""
    names = names or {}
    S.add(f"{tag}_in", f"股權流入（{zh}）", "$B", lambda i, c: None if i == 0 else "=" + inflow(i, c), note, name=names.get("in"))
    S.add(f"{tag}_pre", f"補足前年底現金（{zh}）", "$B",
          lambda i, c: None if i == 0 else (f"={cash25}+{fcf}+«{tag}_in»" if i == 1 else f"=«{tag}_end@p»+{fcf}+«{tag}_in»"), "", name=names.get("pre"))
    S.add(f"{tag}_ext", f"當年外部資金需求（{zh}）", "$B", lambda i, c: None if i == 0 else f"=({minref}-«{tag}_pre»)*({minref}>«{tag}_pre»)", "",
          name=names.get("ext"))
    S.add(f"{tag}_end", f"年底現金（{zh}）", "$B", lambda i, c: f"={cash25}" if i == 0 else f"=«{tag}_pre»+«{tag}_ext»", "", name=names.get("end"))
    S.add(f"{tag}_cum", f"累計外部資金需求（{zh}）", "$B",
          lambda i, c: None if i == 0 else (f"=«{tag}_ext»" if i == 1 else f"=«{tag}_cum@p»+«{tag}_ext»"), "", name=names.get("cum"))


def build(ctx):
    D, wb = ctx.D, ctx.wb
    S, I = D.S, D.I
    years = [I(f"year_{y}") for y in YEARS]
    Y26, months, qpy, capfix = I("year_2026"), I("months_per_year"), I("quarters_per_year"), I("cap_fixed_year")
    idle, mintr = I("idle_share"), I("train_min_share")
    cash25 = S("cash_2025_12_31")
    C, K = ctx.P["a3"]["C"], ctx.P["a3"]["K"]
    kr = K.rows
    from a3 import CT, FAM  # noqa: E402  （合約族與型別）
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731

    # ═════════════════════ Cost 頁補充（A4；chat 端 A3 審查 1、2、4） ═════════════════════
    K.section("十二、逐家合約總額對帳（chat 端 A3 審查：模型 ÷ 招股書逐家 − 1；任一家 |差距| > 門檻 → Checks WARN）")
    for zh, ck, sk, suf in PARTNER_MAP:
        K.add(f"pg_{ck}", f"模型 ÷ 招股書 − 1：{zh}（模型＝第十一節合約總額）", "比例", only25(f"=(«tot_{ck}@$»-{S(sk)}{suf})/({S(sk)}{suf})"), "")
    K.add("pg_max", "逐家 |差距| 最大值", "比例", only25("=MAX(" + ",".join(f"ABS(«pg_{ck}@$»)" for _z, ck, _s, _f in PARTNER_MAP) + ")"),
          "> 門檻（Inputs 0.25）→ WARN；合計對得上（−3%）是因為各家誤差互相抵銷", name="COST_PartnerGapMax", name_cols="D")

    K.section("十三、敏感度：逐家合約錨定招股書金額（基準不變：GW × 合約價，D10）。付款按同一時程等比例縮放；金額型、GW 型、月費型合約的 GW 隨付款同比例（D10：GW＝金額 ÷ 合約價），Broadcom（雙揭露）GW 不變")
    scaled = [ck for _z, ck, _s, _f in PARTNER_MAP if CT[ck][3] in ("amt", "gw", "monthly")]
    for zh, ck, sk, suf in PARTNER_MAP:
        K.add(f"an_k_{ck}", f"錨定係數：{zh}＝招股書 ÷ 模型合約總額", "倍", only25(f"={S(sk)}{suf}/«tot_{ck}@$»"), "")
    K.add("an_cc", "算力成本（現金；錨定招股書）", "$B",
          lambda i, c: "=«cc»+" + "+".join(f"«pay_{ck}»*(«an_k_{ck}@$»-«C:vr_unit@$»)" for _z, ck, _s, _f in PARTNER_MAP), "", name="COST_ComputeAnchor")
    K.add("an_dcc", "差：錨定 − 基準（算力成本）", "$B", lambda i, c: "=«an_cc»-«cc»", "")
    K.add("an_sup", "供給 GW（錨定）", "GW", lambda i, c: "=«C:sup»+" + "+".join(f"«C:sup_{ck}»*(«an_k_{ck}@$»-«C:vr_unit@$»)" for ck in scaled), "")
    K.add("an_vr", "供給 VR 等值 GW（錨定）", "GW",
          lambda i, c: "=«C:vr_sup»+" + "+".join(f"«C:sup_{ck}»*(«an_k_{ck}@$»-«C:vr_unit@$»)*«C:vrf_{CT[ck][2]}»" for ck in scaled), "")
    K.add("an_cap", "容量上限係數（錨定）", "倍",
          lambda i, c: f"=IF({c}$6<={capfix},«C:vr_unit@$»,MIN(«an_sup»*(1-{idle})*(1-{mintr}),«C:eff»)/«C:eff»)", "有效推論 GW 不變（η 以 2025 校準，2025 不受影響）")
    K.add("an_rev", "營收淨額（錨定；截頂後）", "$B", lambda i, c: "=«R:net»*«an_cap»", "")
    K.add("an_gap_b", "差額（含股權報酬；錨定）", "$B", lambda i, c: "=«an_rev»-«an_cc»-«ncx»-«sbc»", "")
    K.add("an_gap_vr", "每 VR 等值 GW 差額（含股權報酬；錨定）", "$B/GW/年", lambda i, c: "=«an_gap_b»/«an_vr»", "", name="COST_PropGap_VR_Anchor")
    K.add("an_dgap", "差：錨定 − 基準（每 VR 等值 GW 差額）", "$B/GW/年", lambda i, c: "=«an_gap_vr»-«p_gap_vr»", "")
    K.add("an_fcf", "自由現金流（錨定）＝營收淨額 − 算力 − 非算力（不含股權報酬）", "$B", lambda i, c: "=«an_rev»-«an_cc»-«ncx»", "Funding 情境 S4")

    K.section("十四、敏感度：機房租約租金（基準不計，D10 r1 只規定不另計 GW；租金＝總值 ÷ 期間，自 Inputs 起算年）")
    ls = I("lease_start_year")
    for key, zh, sk, tk in LEASE_RENT:
        K.add(f"rent_{key}", f"年租金：{zh}", "$B", lambda i, c, sk=sk, tk=tk: f"={S(sk)}/{I(tk)}*({c}$6>={ls})", "")
    K.add("rent_nexus", "（未計）Nexus Hubbard（TX；首期 500 MW；Google 擔保）：租金未揭露", "$B", [None] * 6, "只揭露園區融資 $15B（SRC_ANT_239），租金找不到 → 不計（不杜撰）")
    K.add("rent", "機房租約租金合計（敏感度）", "$B", lambda i, c: "=" + "+".join(f"«rent_{k}»" for k, *_ in LEASE_RENT), "", name="COST_LeaseRent")
    K.add("rent_gap_vr", "每 VR 等值 GW 差額（含股權報酬；加計租金）", "$B/GW/年", lambda i, c: "=«p_gap_vr»-«rent»/«p_den»", "", name="COST_PropGap_VR_Rent")
    K.add("rent_dgap", "差：加計租金 − 基準", "$B/GW/年", lambda i, c: "=«rent_gap_vr»-«p_gap_vr»", "")

    K.section("十五、D22 量化：2026 非算力成本低估（報導 2026 Q2 隱含調整後非算力費用年化 − 模型；不校準，只作敏感度）")
    K.add("d22_qnc", "報導：2026 Q2 隱含非算力費用（含平台抽成，不含股權報酬）＝營收預測 ×（1 − 算力占營收）− 營業利益預測", "$B",
          only26(f"={S('proj_rev_2026q2_wsj')}*(1-{S('cost_compute_per_rev_2026_q2')})-{S('proj_opinc_2026q2_wsj')}"),
          "SRC_ANT_347（10.9）、052（0.56）、348（0.559）；Interested-party（WSJ 5 月預測；實際營收 >11.5、調整後營業利益為正）")
    K.add("d22_ann", "年化（× 季數）", "$B", only26(f"=«d22_qnc»*{qpy}"), "")
    K.add("d22_model", "模型 2026：非算力（不含股權報酬）＋平台抽成", "$B", only26("=«ncx»+«R:pshare»"), "說明書把抽成列行銷費用；模型自營收扣除 → 比較時加回")
    K.add("d22_under", "隱含低估＝年化 − 模型", "$B", only26("=«d22_ann»-«d22_model»"), "正＝模型 2026 非算力成本偏低", name="COST_D22Under", name_cols="E")
    K.add("d22_alt", "交叉對照：（模型調整後營業利益率 − 報導）× 2026 營收淨額", "$B", only26("=(«adj_oim»-«adj_rep»)*«R:net»"), "利潤率法（同一差距的另一算法）")
    K.add("d22_k", "低估比例＝低估 ÷ 模型 2026 非算力（不含股權報酬）", "倍", only26("=«d22_under»/«ncx»"), "敏感度：2026 起各年非算力成本 ×（1＋本比例）")
    K.add("d22_dncx", "非算力成本增加（敏感度）", "$B", lambda i, c: None if i == 0 else "=«ncx»*«d22_k@E»", "", name="COST_D22DeltaNC")
    K.add("d22_gap_vr", "每 VR 等值 GW 差額（含股權報酬；D22 低估補足）", "$B/GW/年", lambda i, c: "=«p_gap_vr»" if i == 0 else "=«p_gap_vr»-«d22_dncx»/«p_den»", "",
          name="COST_PropGap_VR_D22")
    K.add("d22_dgap", "差：D22 補足 − 基準", "$B/GW/年", lambda i, c: "=«d22_gap_vr»-«p_gap_vr»", "")

    # ═════════════════════ Funding ═════════════════════
    F = Sheet(wb, "Funding", "F", "F", years,
              "Funding — 融資：自由現金流、已到位融資、最低現金與外部資金需求、條件式與 IPO 情境、敏感度情境、或有負債、回流對照（D18）",
              "規格 D16 r1、D17、D18、D21：自由現金流＝營收淨額（截頂後）− 算力成本（現金）− 非算力成本（不含股權報酬；股權報酬非現金，加回）。來源順序：①2025 年底現金 → ②2026 已宣布且已交割的股權 → "
              "③外部資金（新股權、債務或延後承諾；補足至最低現金）。條件式與 IPO 不計入基準（情境）；循環信貸與表外融資只列或有。",
              "D–I＝2025–2030；2025 為實際年：只列年底現金（SRC）與自由現金流對照，現金流自 2026 起。$B＝十億美元。")
    F.section("一、自由現金流（A3 交接：COST_GapCash ＋ 股權報酬加回；D21：2025 非現金費用不計）")
    F.add("rev", "營收淨額（截頂後；扣雲端平台抽成）", "$B", lambda i, c: "=«R:net_c»", "REV_NetCapped")
    F.add("cc", "算力成本（現金）＝合約實付＋自建資本支出", "$B", lambda i, c: "=«K:cc»", "COST_Compute")
    F.add("own", "其中：自建資本支出（Fluidstack）", "$B", lambda i, c: "=«K:own»", "COST_OwnedCapex（2026–2028 每年 16.7）")
    F.add("ncx", "非算力成本（不含股權報酬）", "$B", lambda i, c: "=«K:ncx»", "COST_NonCompExSBC")
    F.add("sbc", "股權報酬（非現金；只列示）", "$B", lambda i, c: "=«K:sbc»", "COST_SBC：不耗用現金")
    F.add("gap", "差額（含股權報酬；命題口徑）", "$B", lambda i, c: "=«K:gap_b»", "COST_GapCash（A3 交接的融資缺口起點）")
    F.add("fcf", "自由現金流＝差額＋股權報酬", "$B", lambda i, c: "=«gap»+«sbc»", "＝營收淨額 − 算力成本 − 非算力成本（Checks 恆等式）；2025 欄只對照", name="FND_FCF")
    F.add("ck_fcf", "檢查：自由現金流 −（營收淨額 − 算力 − 非算力）", "$B", lambda i, c: "=«fcf»-(«rev»-«cc»-«ncx»)", "應為 0")
    F.add("fcf_m", "自由現金流率＝自由現金流 ÷ 營收總額（截頂後）", "比例", lambda i, c: "=«fcf»/«R:gross_c»", "對照：管理層燒錢占營收預測 2026 33%、2027 9%（Reverse 第五節）")
    F.add("d21", "D21：2025 淨損中的非現金會計費用（可轉換融資重估；不進現金流）", "$B", only25(f"={S('cost_netloss_2025_noncash_charge')}"), "SRC_ANT_281（說明書）")
    F.add("burn25", "對照：2025 現金燃燒預測（2025-02 版；2025 實際燃燒未揭露）", "$B", only25(f"={S('burn_2025_proj')}"),
          "SRC_ANT_352（3）；2024 實際燃燒 5.6（SRC_ANT_290）；模型 2025 自由現金流見上列 D 欄")

    F.section("二、已到位融資（來源順序 ②；工作單 A4 預設：2026 已宣布且已完成交割者；避免與期初現金、Series H 的 $15B 重複）")
    yr26 = lambda c: f"({c}$6={Y26})"  # noqa: E731
    F.add("in_g", "Series G（2026-02，$30B；GIC、Coatue 領投）", "$B", lambda i, c: f"={S('fund_series_g_amount')}*{yr26(c)}", "SRC_ANT_302；含 Microsoft、NVIDIA 先前承諾的一部分（金額未揭露）")
    F.add("in_amzg", "Amazon Series G 無投票權特別股（2026 Q2，$5B；開關）", "$B", lambda i, c: f"={S('strat_amzn_2026_series_g')}*{I('amzn_seriesg_switch')}*{yr26(c)}",
          "SRC_ANT_319（Amazon 10-Q，Verified）；晚於 Series G 交割 → 另計（Inputs 開關 1）")
    F.add("in_h", "Series H（2026-05，$65B；含 $15B 先前承諾的雲端業者投資）", "$B", lambda i, c: f"={S('fund_series_h_amount')}*{yr26(c)}",
          "SRC_ANT_298、300；$15B 中 Amazon $5B 經融資額度（SRC_ANT_320），其餘 $10B 推定為 Google 2026-04 即時投資")
    F.add("in_h15", "　其中：先前承諾的雲端業者投資（只列示）", "$B", lambda i, c: f"={S('fund_series_h_hyperscaler_component')}*{yr26(c)}", "SRC_ANT_300；不另計")
    F.add("in_msft", "Microsoft $5B（2025-11；開關，基準 0＝已在 2025 年底現金內）", "$B", lambda i, c: f"={S('strat_msft_funded')}*{I('msft_2026_switch')}*{yr26(c)}",
          "SRC_ANT_331；規格 D16 r1 列為已到位 → 以期初現金涵蓋，避免重複")
    F.add("in_googl", "Google 2026-04 即時投資 $10B（開關，基準 0＝視為 Series H 的 $15B 之一）", "$B",
          lambda i, c: f"={S('strat_googl_2026_04_upfront')}*{I('googl_upfront_switch')}*{yr26(c)}", "SRC_ANT_328；與 Series H 雲端業者部分是否同一筆未證實")
    F.add("committed", "已到位融資合計", "$B", lambda i, c: "=«in_g»+«in_amzg»+«in_h»+«in_msft»+«in_googl»", "基準 2026＝30＋5＋65＝100", name="FND_Committed")
    F.add("tender", "員工股份出售（2026-02，$5–6B）：不計入（現金不進公司；工作單 A4 預設）", "$B", [None] * 6, "SRC_ANT_335、336")
    F.add("raised", "對照：成立以來私募累計 >$125B（招股書草稿）vs 已知各輪加總（E、F、G、H、Amazon 可轉債、Google、Microsoft、Amazon G）", "$B",
          [f"={S('fund_private_raised_cum_2026')}_Lo",
           f"={S('fund_series_e_amount')}+{S('fund_series_f_amount')}+{S('fund_series_g_amount')}+{S('fund_series_h_amount')}+{S('strat_amzn_notes_cum_2025')}"
           f"+{S('strat_googl_cum_2025_03')}+{S('strat_msft_funded')}+{S('strat_amzn_2026_series_g')}*{I('amzn_seriesg_switch')}"] + [None] * 4,
          "D＝招股書下限；E＝已知各輪（不含 2023 年以前各輪，故應 ≥ 招股書下限附近）；只對照")

    F.section("三、條件式融資與 IPO（D16 r1：不計入基準，列情境；開關＝1 時 IPO＋Amazon＋Google＋AMD 計入基準）")
    F.add("c_ipo", "IPO 募資（報導約 $100B，2026-11）", "$B", lambda i, c: f"={S('ipo_raise_size')}*({c}$6={I('ipo_year')})", "SRC_ANT_341（Interested-party）；含 NVIDIA 可能的基石認購")
    F.add("c_amzn", "Amazon 剩餘融資額度（依算力交付里程碑；$15B）", "$B", lambda i, c: f"={S('strat_amzn_facility_remaining')}*({c}$6={I('amzn_cond_year')})",
          "SRC_ANT_322（Verified）；以新可轉債提領（屬債務性質）")
    F.add("c_googl", "Google 條件式（績效目標；最高 $30B）", "$B", lambda i, c: f"={S('strat_googl_2026_04_conditional')}*({c}$6={I('googl_cond_year')})", "SRC_ANT_329")
    F.add("c_nvda", "NVIDIA 最高 $10B（已到位金額未揭露）", "$B",
          lambda i, c: f"=({S('strat_msft_nvda_announce')}-{S('strat_msft_funded')})*({c}$6={I('nvda_cond_year')})", "SRC_ANT_330 − 331；與 IPO 基石（SRC_ANT_333）可能同一筆")
    F.add("c_amd", "AMD 最高 $5B 認股", "$B", lambda i, c: f"={S('strat_amd_commit')}*({c}$6={I('amd_equity_year')})", "SRC_ANT_332")
    F.add("c_inbase", "計入基準的條件式（開關 × IPO＋Amazon＋Google＋AMD）", "$B", lambda i, c: f"={I('cond_in_base_switch')}*(«c_ipo»+«c_amzn»+«c_googl»+«c_amd»)",
          "基準開關 0")
    F.add("eq", "計入基準的股權流入＝已到位＋計入基準的條件式", "$B", lambda i, c: "=«committed»+«c_inbase»", "", name="FND_EquityIn")

    F.section("四、現金與外部資金需求（基準；D17 最低現金；外部資金需求＝年底現金低於最低現金的差額）")
    mcm = I("min_cash_months")
    F.add("min", "最低現金＝次年非算力成本（不含股權報酬）× 月數 ÷ 12", "$B",
          lambda i, c: (f"=Cost!{YC[i + 1]}{kr['ncx']}*{mcm}/{months}" if i < 5 else f"=«K:ncx»*{mcm}/{months}"),
          "D17：Inputs 6 個月（3–12）；2030 無次年 → 用 2030 本年", name="FND_MinCash")
    F.add("open", "期初現金（來源順序 ①）", "$B", lambda i, c: None if i == 0 else (f"={cash25}" if i == 1 else "=«end@p»"),
          "2026＝2025 年底現金 20.28（SRC_ANT_305）；之後＝前一年年底現金", name="FND_CashOpen")
    F.add("pre", "補足前年底現金＝期初＋自由現金流＋股權流入", "$B", lambda i, c: None if i == 0 else "=«open»+«fcf»+«eq»", "")
    F.add("ext", "當年外部資金需求（來源順序 ③）", "$B", lambda i, c: None if i == 0 else "=(«min»-«pre»)*(«min»>«pre»)",
          "新股權、債務或延後承諾（不分種類）", name="FND_ExtNeed")
    F.add("end", "年底現金", "$B", lambda i, c: f"={cash25}" if i == 0 else "=«pre»+«ext»", "2025＝SRC（實際）", name="FND_CashEnd")
    F.add("cum", "累計外部資金需求（命題 2）", "$B", lambda i, c: None if i == 0 else ("=«ext»" if i == 1 else "=«cum@p»+«ext»"), "", name="FND_ExtNeedCum")
    F.add("ck_cash", "檢查：年底現金 −（期初＋自由現金流＋股權流入＋外部資金）", "$B", lambda i, c: None if i == 0 else "=«end»-(«open»+«fcf»+«eq»+«ext»)", "應為 0")
    F.add("room", "現金餘裕＝年底現金 − 最低現金", "$B", lambda i, c: None if i == 0 else "=«end»-«min»",
          "2030 值＝若額外支出（例如 2029 起新簽算力）不早於 2029，2026–2030 可再承擔而不需外部資金的上限（仍受各年餘裕約束）", name="FND_Headroom")
    F.add("flag", "缺口旗標（1＝當年需要外部資金）", "旗標", lambda i, c: None if i == 0 else "=SIGN(«ext»)", "", name="FND_GapFlag")
    F.add("flag_y", "缺口年度標示", "年", lambda i, c: None if i == 0 else f'=IF(«flag»,{c}$6,"")', "")
    F.add("noext", "只靠已到位融資（不另募）的年底現金", "$B", lambda i, c: None if i == 0 else (f"={cash25}+«fcf»+«eq»" if i == 1 else "=«noext@p»+«fcf»+«eq»"),
          "已到位融資能撐到哪一年（可為負）", name="FND_CashNoExt")
    F.add("noext_y", "只靠已到位融資：低於最低現金的年度標示", "年", lambda i, c: None if i == 0 else f'=IF(«noext»<«min»,{c}$6,"")', "")
    F.add("peak_y", "峰值年度標示（當年外部資金需求＝最大且 > 0）", "年",
          lambda i, c: None if i == 0 else f'=IF(AND(«flag»,«ext»=MAX(«ext@R»)),{c}$6,"")', "")
    F.add("s_first", "首次缺口年（無則「無」）", "年", only25('=IF(COUNT(«flag_y@R»),MIN(«flag_y@R»),"無")'), "", name="FND_FirstGapYear", name_cols="D")
    F.add("s_peak", "單年外部資金需求峰值", "$B", only25("=MAX(«ext@R»)"), "", name="FND_ExtNeedPeak", name_cols="D")
    F.add("s_peak_y", "峰值年", "年", only25('=IF(COUNT(«peak_y@R»),MIN(«peak_y@R»),"無")'), "", name="FND_PeakYear", name_cols="D")
    F.add("s_trough", "只靠已到位融資的現金谷底（2026–2030 最小值）", "$B", only25("=MIN(«noext@R»)"), "負＝需外部資金補到 0 以上的金額（不含最低現金）",
          name="FND_CashNoExtTrough", name_cols="D")
    F.add("s_noext_first", "只靠已到位融資：首次低於最低現金年", "年", only25('=IF(COUNT(«noext_y@R»),MIN(«noext_y@R»),"無")'), "")

    F.section("五、情境（同一規則；基準開關不影響情境列）：S1 IPO、S2 條件式、S3 全部、S4 逐家錨定招股書、S5 機房租金、S6 D22 非算力補足、S7 S4＋S5＋S6 合併")
    mn = "«min»"
    waterfall(F, "s1", "S1：＋IPO 約 $100B", "«fcf»", lambda i, c: "«committed»+«c_ipo»", cash25, mn, dict(cum="FND_ExtNeedCumIPO", end="FND_CashEndIPO"))
    waterfall(F, "s2", "S2：＋條件式（Amazon、Google、NVIDIA、AMD）", "«fcf»", lambda i, c: "«committed»+«c_amzn»+«c_googl»+«c_nvda»+«c_amd»", cash25, mn,
              dict(cum="FND_ExtNeedCumCond"))
    waterfall(F, "s3", "S3：＋IPO＋條件式（NVIDIA 視為 IPO 基石，不另計）", "«fcf»", lambda i, c: "«committed»+«c_ipo»+«c_amzn»+«c_googl»+«c_amd»", cash25, mn,
              dict(cum="FND_ExtNeedCumAll", end="FND_CashEndAll"))
    waterfall(F, "s4", "S4：逐家錨定招股書金額", "«K:an_fcf»", lambda i, c: "«eq»", cash25, mn, dict(cum="FND_ExtNeedCumAnchor"),
              "自由現金流改用 Cost 第十三節（算力成本、容量上限、營收連動）")
    waterfall(F, "s5", "S5：加計機房租約租金", "(«fcf»-«K:rent»)", lambda i, c: "«eq»", cash25, mn, dict(cum="FND_ExtNeedCumRent"), "Cost 第十四節")
    waterfall(F, "s6", "S6：D22 非算力成本補足", "(«fcf»-«K:d22_dncx»)", lambda i, c: "«eq»", cash25, mn, dict(cum="FND_ExtNeedCumD22"), "Cost 第十五節")
    waterfall(F, "s7", "S7：S4＋S5＋S6 合併（不利）", "(«K:an_fcf»-«K:rent»-«K:d22_dncx»)", lambda i, c: "«eq»", cash25, mn, dict(cum="FND_ExtNeedCumAdverse"), "")

    F.section("六、或有負債與承諾（D16 r1：只列示，不作資金來源）")
    F.add("cl_rev", "循環信貸額度（$15B；招股書稱已有／定案中）", "$B", only25(f"={S('debt_credit_facility_15b')}"), "SRC_ANT_306、344")
    F.add("cl_bcm", "Broadcom 租賃融資（可轉換票據，最高 $42B；IPO 後）", "$B", only25(f"={S('cmp_broadcom_financing_42b')}_Hi"), "SRC_ANT_251（招股書草稿經 Reuters）")
    F.add("cl_spv", "表外 Google TPU 融資 SPV（>$71B；單一來源，未證實）", "$B", only25(f"={S('debt_spv_tpu_financing')}"), "SRC_ANT_307（只有 ZeroHedge 轉述）")
    F.add("cl", "或有負債合計（只列示）", "$B", only25("=«cl_rev»+«cl_bcm»+«cl_spv»"), "不進現金流與外部資金需求", name="FND_Contingent", name_cols="D")
    from a3 import CONTRACTS  # noqa: E402
    ckeys = [k[0] for k in CONTRACTS]
    F.add("after", "合約剩餘承諾（2030 年後應付）＝各新合約總額 − 2025–2030 實付", "$B",
          only25("=(" + "+".join(f"«K:tot_{k}@$»" for k in ckeys) + ")-(" + "+".join(f"SUM(«K:pay_{k}@R»)" for k in ckeys) + ")"),
          "Cost 第十一節總額（模型）；只列示", name="FND_CommitAfter2030", name_cols="D")
    F.add("noncancel", "對照：招股書承諾中不可取消比例（約）", "比例", only25(f"={S('cmp_prospectus_noncancel_share')}"), "SRC_ANT_246；xAI 部分可 90 天取消")

    F.section("七、回流對照（D18；只列對照、不沖銷）：欄位不是年度——D 已到位投資、E 條件式／最高可達（未到位）、F 算力合約總額（招股書）、G 模型 2026–2030 付款、H 已到位 ÷ 模型付款、I（已到位＋條件式）÷ 合約總額")
    F.add("rt_hdr", "表頭", "混合", ["已到位投資", "條件式（未到位）", "合約總額（招股書）", "模型 2026–30 付款", "已到位÷付款", "（到位＋條件）÷合約"], "")

    def paysum(keys):
        return "+".join(f"SUM(Cost!$E${kr['pay_' + k]}:$I${kr['pay_' + k]})" for k in keys)

    rows = [
        ("amzn", "Amazon（可轉債 8＋Series G 5＋Series H 5）", f"={S('strat_amzn_funded_cum_2026h1')}", f"={S('strat_amzn_facility_remaining')}",
         f"={S('cmp_prospectus_amazon')}", "=" + paysum(["aws"]), "既有雲端（2025 前簽）中付 AWS 的部分未拆分，未計入"),
        ("googl", "Google（2025-03 前 ≥3＋2026-04 即時 10）", f"={S('strat_googl_cum_2025_03')}+{S('strat_googl_2026_04_upfront')}", f"={S('strat_googl_2026_04_conditional')}",
         f"={S('cmp_prospectus_google')}_Lo", "=" + paysum(["google"]), "Google 即時 $10B 在基準中視為 Series H 的一部分（已到位）"),
        ("bcm", "Broadcom（股權 0；租賃融資最高 42 屬債務）", None, f"={S('cmp_broadcom_financing_42b')}_Hi", f"={S('cmp_prospectus_broadcom')}", "=" + paysum(["broadcom"]),
         "TPU 設備租賃；融資為可轉換票據"),
        ("msft", "Microsoft（2025-11，$5B）", f"={S('strat_msft_funded')}", None, f"={S('cmp_msft_azure_prospectus')}", "=" + paysum(["azure"]), ""),
        ("nvda", "NVIDIA（最高 $10B，已到位未揭露）", None, f"={S('strat_msft_nvda_announce')}-{S('strat_msft_funded')}", None, "=" + paysum(NVDA_FAMILY),
         "無直接合約：G＝以 NVIDIA 晶片為主的合約付款（Azure、SpaceX、Nscale、Lambda、Volta、Akamai；與 Microsoft 列重複 Azure）"),
        ("amd", "AMD（最高 $5B 認股）", None, f"={S('strat_amd_commit')}", f"={S('cmp_amd_usd_prospectus')}_Lo", "=" + paysum(["amd"]), ""),
    ]
    for key, zh, d_, e_, f_, g_, note in rows:
        h_ = f"=«rt_{key}@D»/«rt_{key}@G»" if d_ else None
        i_ = (f"=(«rt_{key}@D»+«rt_{key}@E»)/«rt_{key}@F»" if d_ and e_ else f"=«rt_{key}@D»/«rt_{key}@F»" if d_ else f"=«rt_{key}@E»/«rt_{key}@F»") if f_ else None
        F.add(f"rt_{key}", zh, "混合", [d_, e_, f_, g_, h_, i_], note)
    tk = [r[0] for r in rows]
    F.add("rt_tot", "合計（G 不含 NVIDIA 間接列）", "混合",
          ["=" + "+".join(f"«rt_{k}@D»" for k in tk), "=" + "+".join(f"«rt_{k}@E»" for k in tk), "=" + "+".join(f"«rt_{k}@F»" for k in tk),
           "=" + "+".join(f"«rt_{k}@G»" for k in tk if k != "nvda"), "=«rt_tot@D»/«rt_tot@G»", "=(«rt_tot@D»+«rt_tot@E»)/«rt_tot@F»"],
          "已到位策略投資 ÷ 對同一方 2026–2030 付款", name="FND_RoundTrip", name_cols="DI")

    # ═════════════════════ Reverse ═════════════════════
    X = Sheet(wb, "Reverse", "X", "X", years,
              "Reverse — 反向模式（D19）：假設管理層營收目標達成，反解所需倍數與反向資金（同一支出）；只作對照，不回饋基準",
              "目標＝各年最新一版公司預測（Interested-party；The Information 2025-11、2026-01 轉述）；2030 無公開目標 → 取正向值（倍數＝1）。本頁只被 Checks 引用（共同規則第 4 節；test_reverse_not_fed_back）。",
              "注意：這些目標都早於 2026 年的實際成長（2026 上半年實際已超過 2026-01 版全年目標），所以多數年度目標低於本模型正向值——倍數 < 1 代表「管理層舊目標比模型保守」。")
    X.section("一、管理層營收目標（只驅動本頁）")
    tg = [f"={S('rev_2025_recognized')}", f"={S('proj_rev_2026_2026_01')}", f"={S('proj_rev_2027_2026_01')}", f"={S('proj_rev_2028_2025_11')}",
          f"={S('proj_rev_2029_2026_01')}", "=«fwd»"]
    X.add("tgt", "管理層營收目標（總額口徑）", "$B", tg,
          "2025＝實際（SRC_ANT_001）；2026＝18（2026-01，SRC_ANT_368）；2027＝55（2026-01 樂觀，369）；2028＝最高 70（2025-11，362）；2029＝最高 148（2026-01，370）；2030＝無目標，取正向值",
          name="RVS_Target")
    X.add("tgt_g", "目標年增率", "比例", lambda i, c: None if i == 0 else "=(«tgt»-«tgt@p»)/«tgt@p»", "")
    X.add("tgt_old", "對照：2025-02 版 2027 營收（基準 12／樂觀 34.5）", "$B", [None, None, f"={S('proj_rev_2027_base_2025_02')}", None, None, None], "SRC_ANT_359、360")
    X.add("tgt_arr", "對照：2026 年底年化營收（2025-11 目標 20–26；投資人預期 100–120 為下限）", "$B/年",
          [None, f"={S('proj_arr_2026_target_2025_11')}_Hi", f"={S('proj_arr_2026_end_investors')}_Lo", None, None, None], "E＝2025-11 目標上緣；F＝投資人預期下限（年化，非年度營收）")

    X.section("二、正向對照（本模型基準；截頂後）")
    X.add("fwd", "正向營收總額（截頂後）", "$B", lambda i, c: "=«R:gross_c»", "REV_GrossCapped")
    for k, zh in (("sub", "訂閱"), ("api", "API"), ("other", "其他")):
        X.add(f"f_{k}", f"正向{zh}（截頂後）", "$B", lambda i, c, k=k: f"=«R:{k}»*«R:cap»", "各營收線同比例截頂")
    X.add("gap", "差距（目標 − 正向）", "$B", lambda i, c: "=«tgt»-«fwd»", "負＝目標低於正向", name="RVS_Gap")

    X.section("三、所需倍數（相對正向；1＝正向剛好；< 1＝目標低於正向；廣告＝0，D3）")
    X.add("m_api", "只靠 API：API 營收所需倍數", "倍", lambda i, c: "=(«tgt»-«f_sub»-«f_other»)/«f_api»", "", name="RVS_MultAPI")
    X.add("m_sub", "只靠訂閱：訂閱營收所需倍數", "倍", lambda i, c: "=(«tgt»-«f_api»-«f_other»)/«f_sub»", "", name="RVS_MultSub")
    X.add("m_prop", "兩線等比例：共同倍數", "倍", lambda i, c: "=(«tgt»-«f_other»)/(«f_sub»+«f_api»)", "", name="RVS_MultProp")

    X.section("四、反向資金（目標營收＋同一支出；同 S7 規則；只作對照）")
    X.add("net", "營收淨額（目標；平台抽成占比同正向）", "$B", lambda i, c: "=«tgt»*«R:net_c»/«R:gross_c»", "")
    X.add("cc", "算力成本（現金；同正向）", "$B", lambda i, c: "=«K:cc»", "COST_Compute")
    X.add("ncx", "非算力成本（不含股權報酬；同正向，不隨目標放大）", "$B", lambda i, c: "=«K:ncx»", "同 OpenAI v0.6 S5 預設「同一支出」")
    X.add("fcf", "自由現金流（反向）", "$B", lambda i, c: "=«net»-«cc»-«ncx»", "", name="RVS_FCF")
    waterfall(X, "r", "反向", "«fcf»", lambda i, c: "«F:eq»", cash25, "«F:min»", dict(cum="RVS_ExtNeedCum", end="RVS_CashEnd", ext="RVS_ExtNeed"),
              "FND_EquityIn（同基準）；最低現金＝FND_MinCash")

    X.section("五、與管理層說法對照（現金流轉正年、2028 現金流、燒錢占營收）")
    X.add("pos_f", "正向自由現金流 > 0 的年度標示", "年", lambda i, c: None if i == 0 else f'=IF(«F:fcf»>«F:fcf»-«F:fcf»,{c}$6,"")', "")
    X.add("pos_r", "反向自由現金流 > 0 的年度標示", "年", lambda i, c: None if i == 0 else f'=IF(«fcf»>«fcf»-«fcf»,{c}$6,"")', "")
    X.add("pos_sum", "現金流轉正年：管理層（2026-01 版）／正向首年／反向首年", "年",
          [f"={S('proj_cfpos_year_2026_01')}", '=IF(COUNT(«pos_f@R»),MIN(«pos_f@R»),"無")', '=IF(COUNT(«pos_r@R»),MIN(«pos_r@R»),"無")', None, None, None],
          "D＝SRC_ANT_371（2028）；E＝本模型正向；F＝反向", name="RVS_CFPosYear", name_cols="DF")
    X.add("cf28", "2028 自由現金流：管理層（2025-11 版 17）／正向／反向", "$B",
          [f"={S('proj_cf_2028_2025_11')}", "=«F:fcf@G»", "=«fcf@G»", None, None, None], "D＝SRC_ANT_365")
    X.add("burn", "燒錢占營收：管理層預測（2026 33%、2027 9%）vs 正向（−自由現金流 ÷ 營收總額）", "比例",
          [None, f"={S('burn_2026_ratio_proj')}", f"={S('burn_2027_ratio_proj')}", "=-«F:fcf_m@E»", "=-«F:fcf_m@F»", None],
          "E、F＝管理層（SRC_ANT_353、354）；G、H＝本模型 2026、2027")

    for sh in (F, X):
        sh.resolve()
    K.resolve()
    for sh in (F.ws, X.ws):
        sh.column_dimensions["A"].width = 7
        sh.column_dimensions["B"].width = 66
        sh.column_dimensions["C"].width = 12
        for c in YC:
            sh.column_dimensions[c].width = 13
        sh.column_dimensions["K"].width = 80
        sh.freeze_panes = "D7"
    ctx.P["a4"] = dict(F=F, X=X)


def checks(ctx):
    D = ctx.D
    I = D.I
    F, X, K = ctx.P["a4"]["F"], ctx.P["a4"]["X"], ctx.P["a3"]["K"]
    yr = lambda sh, key, a="D": f"{sh.name}!${a}${sh.rows[key]}:$I${sh.rows[key]}"  # noqa: E731
    mx = lambda sh, key, a="D": f"=MAX(MAX({yr(sh, key, a)}),-MIN({yr(sh, key, a)}))"  # noqa: E731
    d = lambda sh, key, col="D": f"{sh.name}!${col}${sh.rows[key]}"  # noqa: E731
    return [
        ("a4_fcf", "自由現金流 −（營收淨額 − 算力 − 非算力）（各年最大絕對差，$B）", mx(F, "ck_fcf"), 0, "tol", "A3 交接：股權報酬加回"),
        ("a4_cash", "年底現金 −（期初＋自由現金流＋股權流入＋外部資金）（2026–2030 最大絕對差）", mx(F, "ck_cash", "E"), 0, "tol", "來源順序恆等式"),
        ("a4_extneg", "外部資金需求 < 0 的年數", f'=COUNTIF({yr(F, "ext", "E")},"<0")', 0, "eq", ""),
        ("a4_mincash", "最低現金 ≤ 0 的年數", f'=COUNTIF({yr(F, "min")},"<=0")', 0, "eq", "D17：月數或非算力成本為負時轉 ERR"),
        ("a4_cum", "累計外部資金需求 2030 − Σ 當年需求", f"={d(F, 'cum', 'I')}-SUM({yr(F, 'ext', 'E')})", 0, "tol", ""),
        ("a4_rvs_iso", "Reverse 被基準頁引用的公式格數（builder 掃描）", ctx.scan["rvs_refs"], 0, "eq", "D19：只作對照、不回饋；tests 每次重驗（test_reverse_not_fed_back）"),
        ("a4_partner", "逐家合約總額：模型 ÷ 招股書 − 1 的最大 |差距|（WARN）", "=COST_PartnerGapMax", I("partner_warn_threshold"), "warn",
         "chat 端 A3 審查：AMD 模型 77.8 vs 20、SpaceX 45 vs 84.5、Broadcom 125.2 vs 161.2；基準不改（D10），見 Cost 第十三節錨定敏感度"),
        ("a4_committed", "結果：2026 已到位融資（$B）", f"={d(F, 'committed', 'E')}", None, "info", "Series G 30＋Amazon G 5＋Series H 65"),
        ("a4_first", "結果：首次缺口年", "=FND_FirstGapYear", None, "info", "命題 2"),
        ("a4_peak", "結果：單年外部資金需求峰值（$B）", "=FND_ExtNeedPeak", None, "info", ""),
        ("a4_peak_y", "結果：峰值年", "=FND_PeakYear", None, "info", ""),
        ("a4_cum30", "結果：2030 累計外部資金需求（$B；命題 2）", f"={d(F, 'cum', 'I')}", None, "info", ""),
        ("a4_cash30", "結果：2030 年底現金（$B）", f"={d(F, 'end', 'I')}", None, "info", ""),
        ("a4_trough", "結果：只靠已到位融資的現金谷底（$B）", "=FND_CashNoExtTrough", None, "info", ""),
        ("a4_anchor", "敏感度：逐家錨定招股書的 2030 累計外部資金需求 − 基準（$B）", f"=INDEX(FND_ExtNeedCumAnchor,1,6)-{d(F, 'cum', 'I')}", None, "info", ""),
        ("a4_d22", "D22：2026 非算力成本隱含低估（$B；報導 Q2 年化 − 模型）", "=COST_D22Under", None, "info", "不校準"),
        ("a4_rent", "敏感度：機房租約年租金 2030（$B）", "=INDEX(COST_LeaseRent,1,6)", None, "info", "TeraWulf、Riot、Hut 8；Nexus 未揭露"),
        ("a4_rt", "D18：已到位策略投資 ÷ 對同一方 2026–2030 付款", f"={d(F, 'rt_tot', 'H')}", None, "info", "只對照、不沖銷"),
        ("a4_raised", "對照：已知各輪加總 − 招股書私募累計下限 $125B", f"={d(F, 'raised', 'E')}-{d(F, 'raised')}", None, "info", "不含 2023 年以前各輪"),
        ("a4_rvs_cum30", "Reverse：2030 反向累計外部資金需求（$B）", "=INDEX(RVS_ExtNeedCum,1,6)", None, "info", "管理層目標營收＋同一支出"),
        ("a4_rvs_m29", "Reverse：2029 兩線等比例所需倍數", "=INDEX(RVS_MultProp,1,5)", None, "info", "< 1＝管理層目標低於正向"),
        ("a4_oai_cum", "OpenAI 並排：2030 累計外部資金需求 Anthropic − OpenAI（$B）", f"={d(F, 'cum', 'I')}-INDEX(OAI_FND_ExtNeedCum,1,6)", None, "info", "OAI_Link（D20）"),
        ("a4_oai_cash", "OpenAI 並排：2030 年底現金 Anthropic − OpenAI（$B）", f"={d(F, 'end', 'I')}-INDEX(OAI_FND_CashEnd,1,6)", None, "info", ""),
    ]
