"""Anthropic v0.1-A2：需求頁 Demand 與營收頁 Revenue（工作單 A2 第 4–6 步；規格 D2–D5、D14、D15 r1）。

Python 只產生結構：每列的公式文字由此組裝，數值一律引用 SRC_ANT／Inputs／TK_ 具名範圍（E6：公式不含常數；恆等式 1−x、1＋成長率、年數 +1 除外）。
兩頁互相引用（API 2025／2026 token 由 Revenue 的營收倒推；訂閱營收用 Demand 的人數），公式先以列鍵記號撰寫，兩頁建好後一次換成儲存格位址：
  «k»＝本頁同欄、«k@p»＝本頁前一欄、«k@D»＝本頁 D 欄（2025）、«k@$»＝本頁 $D$（只在 D 欄有值的參數列）、«k@R»＝本頁整列 $D:$I、
  «D:k»／«R:k»／«C:k»／«K:k»＝Demand／Revenue／Compute／Cost 同欄（亦可加 @p、@D、@$；A3 起 Compute、Cost 沿用本類別）。
層級對應：Haiku→低層（Tokenomics Luna）、Sonnet→中層（Sol）、Opus→頂層（Astra）；具名範圍沿用 OpenAI v0.6 的 Top／Mid／Low。
容量上限：Revenue「容量上限係數」列（REV_CapFactor）由 A3（builder/a3.py）填入 Compute 頁 CMP_CapFactor 同欄（2025、2026 固定 1）；其餘截頂列自動連動。
"""
from __future__ import annotations

import re

from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, header, nm, put, title

YEARS = (2025, 2026, 2027, 2028, 2029, 2030)
YC = "DEFGHI"
CONS = (("pro", "Pro", "Pro"), ("max5", "Max 5x", "Max5"), ("max20", "Max 20x", "Max20"))
SEATS = (("team_std", "Team 標準", "TeamStd"), ("team_prem", "Team Premium", "TeamPrem"), ("ent", "Enterprise", "Ent"))
PLANS = (("free", "Free", "Free"),) + CONS + SEATS
PAID = CONS + SEATS
CATS = (("chat", "對話"), ("code", "程式代理"), ("other", "其他代理"))
TIERS = (("top", "頂層（Opus→Astra）", "Top"), ("mid", "中層（Sonnet→Sol）", "Mid"), ("low", "低層（Haiku→Luna）", "Low"))
TIER_INP = {"low": "haiku", "mid": "sonnet"}            # Inputs 只設低、中層占比，頂層＝1 − 其餘
# 各層級的牌價序列（tierRole：每層級以該層最新模型的牌價為準，新模型上市視為該層變價；同 OpenAI v0.6 INP_050）。
# 只列結構（哪些 SRC 列構成同一層級的序列）；價格與日期全部取自 SRC_ANT。Fable／Mythos（限量或更高價旗艦）不納入頂層序列（見報告「已套用的預設」）。
CHAINS = {
    "top": ["price_claude3_opus", "price_opus4", "price_opus41", "price_opus45", "price_opus46", "price_opus47", "price_opus48", "price_opus5", "price_opus55"],
    "mid": ["price_claude35_sonnet", "price_claude37_sonnet", "price_sonnet4", "price_sonnet45", "price_sonnet46", "price_sonnet5", "price_sonnet55"],
    "low": ["price_claude35_haiku_cut", "price_haiku45", "price_haiku55"],
}
FMT = {"$B": "#,##0.000", "T": "#,##0", "M": "#,##0.000", "比例": "0.0%", "比例/年": "0.0%", "$/M": "0.000", "$/月": "0.00", "倍": "0.000",
       "日期": "0", "任務/日": "0.00", "K tok/任務": "#,##0.0", "十億任務": "#,##0.00", "$B/年": "#,##0.000"}
TOK = re.compile(r"«(?:(D|R|C|K):)?([A-Za-z0-9_]+)(?:@(p|[D-I]|\$|R))?»")


class Sheet:
    """年度表：A 編號、B 項目、C 單位、D–I 2025–2030、K 說明。"""
    all = {}

    def __init__(self, wb, name, code, prefix, years, t1, t2, t3=None):
        self.wb, self.name, self.prefix = wb, name, prefix
        self.ws = wb.create_sheet(name)
        title(self.ws, t1, t2, t3)
        header(self.ws, 5, ["編號", "項目", "單位"] + [str(y) for y in YEARS] + ["", "說明"])
        put(self.ws, "B6", "年度", F_BOLD)
        for i in range(6):
            put(self.ws, f"{YC[i]}6", f"={years[i]}")
        self.r, self.n, self.rows, self.meta = 7, 0, {}, []
        Sheet.all[code] = self

    def section(self, label):
        put(self.ws, f"A{self.r}", label, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def add(self, key, label, unit, fs, note="", name=None, name_cols=None):
        """fs：callable(i, c) → 公式（None＝空白），或長度 6 的清單。name_cols：具名範圍只涵蓋的欄（例如 "D" 或 "DE"）。"""
        self.n += 1
        code = f"{self.prefix}{self.n:02d}"
        r = self.r
        self.rows[key] = r
        put(self.ws, f"A{r}", code, F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        for i in range(6):
            f = fs(i, YC[i]) if callable(fs) else fs[i]
            if f is not None:
                put(self.ws, f"{YC[i]}{r}", f, fmt=FMT.get(unit))
        put(self.ws, f"K{r}", note, F_NOTE)
        if name:
            cols = name_cols or "DI"
            ref = f"{self.name}!${cols[0]}${r}" + (f":${cols[-1]}${r}" if cols[-1] != cols[0] else "")
            nm(self.wb, name, ref)
        self.meta.append((code, key, label, unit, name))
        self.r += 1
        return r

    def resolve(self, extra=None):
        for r in self.rows.values():
            for i, col in enumerate(YC):
                c = self.ws[f"{col}{r}"]
                if isinstance(c.value, str) and "«" in c.value:
                    c.value = _resolve(c.value, self, col, i, extra or {})


def _resolve(f, me, col, i, extra):
    def rep(m):
        sh, key, at = m.group(1), m.group(2), m.group(3)
        if key in extra:
            return extra[key](col)
        tgt = Sheet.all[sh] if sh else me
        if at == "p" and i == 0:
            raise ValueError(f"«{key}@p» 用在 2025 欄：{f}")
        if at == "$":                                  # 只在 D 欄有值的參數列：絕對位址
            ref = f"$D${tgt.rows[key]}"
        elif at == "R":                                # 整列 2025–2030（SUMPRODUCT 用）
            ref = f"$D${tgt.rows[key]}:$I${tgt.rows[key]}"
        else:
            c = YC[i - 1] if at == "p" else (at or col)
            ref = f"{c}{tgt.rows[key]}"
        return ref if tgt is me else f"{tgt.name}!{ref}"
    return TOK.sub(rep, f)


def build(ctx):
    D, wb = ctx.D, ctx.wb
    S, I = D.S, D.I
    Sheet.all = {}
    years = [I(f"year_{y}") for y in YEARS]
    days, months, h2, m2b, b2t = I("days_per_year"), I("months_per_year"), I("h2_months"), I("m_to_b"), I("b_to_t")
    start = I("elasticity_start_year")

    def by_year(prefix, i, first=2026):
        """逐年 Inputs（key＝prefix_YYYY）；first 之前回傳 None。"""
        y = YEARS[i]
        return I(f"{prefix}_{y}") if y >= first else None

    def two(k25, k26):
        """2025 一欄、2026 起一欄的 Inputs。"""
        return lambda i: I(k25) if i == 0 else I(k26)

    def grow(key, base, rate):
        """D 欄＝base；之後＝前一年 ×（1＋rate(i)）。"""
        return lambda i, c: (base if i == 0 else f"=«{key}@p»*(1+{rate(i)})")

    # ═════════════════════ Revenue（先建：價格與校準基準） ═════════════════════
    R = Sheet(wb, "Revenue", "R", "R", years,
              "Revenue — 營收：API 牌價（價格事件天數加權）與有效單價、訂閱方案別、API 層級別、其他、廣告（＝0）、總額、雲端平台抽成、淨額、容量上限（A2 佔位）",
              "2025 為校準年（D14）：訂閱＝SRC 訂閱營收、API＝SRC 用量計價營收、其他＝殘差，總額＝SRC 認列營收。2026 為半校準年（D15 r1）：總額＝Q1＋Q2 實際＋2026-07 年化 × 6／12，"
              "API 為倒推項。2027 起正向推導（Demand 的任務數與每任務 token × 價格反應）。報導口徑為總額（雲端市集以總額認列），淨額＝總額 − 雲端平台抽成（D5 r1）；命題與現金流用淨額。",
              "單位：營收 $B；牌價與有效單價 $/M token（百萬 token）；月費 $/月。run-rate 里程碑只作對照（D14、D23），不反推參數。")
    R.section("一、年度日期（價格事件天數加權用）")
    R.add("yst", "年初日（日期序列值）", "日期", lambda i, c: f"=DATE({c}$6,{I('fy_start_month')},{I('fy_start_day')})", "DATE(年度, 起始月, 起始日)")
    R.add("yen", "年底次日（日期序列值）", "日期", lambda i, c: f"=DATE(({c}$6+1),{I('fy_start_month')},{I('fy_start_day')})", "下一年度的年初日")

    R.section("二、API 牌價（$/M token；層級以該層最新模型為準；2025、2026 依價格事件公告日天數加權，2027 起＝2026 年底牌價 ×（1＋年變動率））")
    chg = I("price_chg")
    for t, tz, T in TIERS:
        for io, ioz in (("in", "輸入"), ("out", "輸出")):
            R.add(f"w_{t}_{io}", f"{tz} {ioz}牌價（價格事件天數加權）", "$/M", lambda i, c, t=t, io=io: f"=SUMPRODUCT(«EVM_{t}_{io}»,«EVC_{t}_{io}»)",
                  "Σ（本事件價 − 前一事件價）× 該事件之後占當年的比例（第十二節事件表）；2027 起各事件比例皆為 1，即 2026 年底牌價")
            R.add(f"lp_{t}_{io}", f"{tz} {ioz}牌價（FY 平均）", "$/M",
                  lambda i, c, t=t, io=io: f"=«w_{t}_{io}»" if i < 2 else (f"=«w_{t}_{io}»*(1+{chg})" if i == 2 else f"=«lp_{t}_{io}@p»*(1+{chg})"),
                  "2025、2026＝事件加權；2027＝2026 年底牌價 ×（1＋年變動率）；之後逐年", name=f"REV_ListPrice_{T}_{io.capitalize()}")

    R.section("三、API 有效單價（$/M token；含輸出占比、快取讀寫、Batch 與議價折扣）")
    h, w, b, neg, phi = I("api_cache_hit"), I("api_cache_write"), I("api_batch_share"), I("api_neg_disc"), I("api_out_share")
    R.add("cache_mult", "輸入 token 價格倍數（快取讀寫加權）", "倍", lambda i, c: f"=(1-{h}-{w})+{h}*{S('price_cache_read_mult')}+{w}*{S('price_cache_write_5m_mult')}",
          "（1 − 讀 − 寫）× 1 ＋ 讀 × 讀取倍數（SRC_ANT 0.1）＋ 寫 × 寫入倍數（SRC_ANT 1.25）")
    R.add("disc", "折扣係數（Batch × 議價）", "倍", lambda i, c: f"=(1-{b}*{S('price_batch_discount')})*(1-{neg})", "（1 − Batch 占比 × Batch 折扣）×（1 − 議價折扣）")
    for t, tz, T in TIERS:
        R.add(f"ep_{t}", f"{tz} 有效單價", "$/M", lambda i, c, t=t: f"=((1-{phi})*«lp_{t}_in»*«cache_mult»+{phi}*«lp_{t}_out»)*«disc»",
              "（（1 − 輸出占比）× 輸入價 × 快取倍數 ＋ 輸出占比 × 輸出價）× 折扣係數", name=f"REV_ApiPrice_{T}")
    for t, tz, T in TIERS[1:]:
        k = two(f"api_mix_{TIER_INP[t]}_2025", f"api_mix_{TIER_INP[t]}_2026")
        R.add(f"mix_{t}", f"API token 層級組合：{tz}", "比例", lambda i, c, k=k: f"={k(i)}", "Inputs（2025 一欄；2026 起一欄）")
    R.add("mix_top", f"API token 層級組合：{TIERS[0][1]}", "比例", lambda i, c: "=(1-«mix_mid»-«mix_low»)", "＝1 − 中層 − 低層")
    R.add("mixprice", "API 組合有效單價", "$/M", lambda i, c: "=" + "+".join(f"«mix_{t}»*«ep_{t}»" for t, _, _ in TIERS), "Σ 層級占比 × 有效單價",
          name="REV_ApiPrice")

    R.section("四、訂閱實收月費（$/月；個人方案與企業席位）")
    pa, ta = I("pro_annual_share"), I("team_annual_share")
    R.add("p_pro", "Pro 實收月費", "$/月", lambda i, c: f"={S('plan_pro_monthly')}*(1-{pa})+{S('plan_pro_annual')}*{pa}", "月付 ×（1 − 年付占比）＋ 年付月均 × 年付占比")
    R.add("p_max5", "Max 5x 月費", "$/月", lambda i, c: f"={S('plan_max5x')}", "只有月付")
    R.add("p_max20", "Max 20x 月費", "$/月", lambda i, c: f"={S('plan_max20x')}", "只有月付")
    R.add("p_team_std", "Team 標準席位實收月費", "$/月", lambda i, c: f"={S('plan_team_standard_monthly')}*(1-{ta})+{S('plan_team_standard_annual')}*{ta}", "同上（年付占比）")
    R.add("p_team_prem", "Team Premium 席位實收月費", "$/月", lambda i, c: f"={S('plan_team_premium_monthly')}*(1-{ta})+{S('plan_team_premium_annual')}*{ta}", "同上")
    R.add("p_ent", "Enterprise 席位月費", "$/月", lambda i, c: f"={S('plan_enterprise_seat_legacy')}" if i == 0 else f"={S('plan_enterprise_seat_2026')}",
          "2025＝舊制（第三方指稱 $60，SRC Analogy）；2026 起＝$20 席位費，用量另按 API 牌價計（落在 API 營收）")
    R.add("p_cons_w", "個人付費加權月費", "$/月", lambda i, c: "=" + "+".join(f"«D:cmix_{p}»*«p_{p}»" for p, _, _ in CONS), "Σ 方案占比（Demand）× 月費")
    R.add("p_seat_w", "企業席位加權月費", "$/月", lambda i, c: "=" + "+".join(f"«D:smix_{p}»*«p_{p}»" for p, _, _ in SEATS), "Σ 席位占比（Demand）× 月費")

    R.section("五、校準基準（2025 實際：D14；2026 半校準：D15 r1）")
    R.add("cal_cons", "2025 個人訂閱營收＝消費者訂閱占營收（SRC，Analogy）× 2025 營收", "$B", [f"={S('rev_mix_consumer_share_2025_jaipuria')}*{S('rev_2025_recognized')}"] + [None] * 5,
          "SRC_ANT 分析師估計 15%（10–20%）；個人付費人數由此倒推（工作單預設）")
    R.add("cal_seat", "2025 企業席位訂閱營收＝SRC 訂閱營收 − 個人訂閱", "$B", [f"={S('rev_mix_2025_subscription')}-«cal_cons»"] + [None] * 5, "企業席位數由此倒推")
    R.add("target", "營收總額基準（2025＝SRC 認列營收；2026＝Q1＋Q2 實際＋2026-07 年化 × 下半年月數 ÷ 12）", "$B",
          [f"={S('rev_2025_recognized')}", f"={S('rev_2026_q1')}+{S('rev_2026_q2')}+{S('rev_runrate_2026_07')}*{h2}/{months}"] + [None] * 4,
          "D14；D15 r1：下半年持平於最新觀測年化營收（保守；Decision）", name="REV_Target", name_cols="DE")

    R.section("六、訂閱營收（$B＝人數（M）× 實收月費 × 12 ÷ 1000）")
    for p, pz, P in PAID:
        src = "u" if (p, pz, P) in CONS else "s"
        R.add(f"sub_{p}", f"{pz} 訂閱營收", "$B", lambda i, c, p=p, src=src: f"=«D:{src}_{p}»*«p_{p}»*{months}/{m2b}", "", name=f"REV_Sub_{P}")
    R.add("sub_cons", "個人訂閱合計（Pro＋Max）", "$B", lambda i, c: "=" + "+".join(f"«sub_{p}»" for p, _, _ in CONS), "", name="REV_SubConsumer")
    R.add("sub_seat", "企業席位訂閱合計（Team＋Enterprise）", "$B", lambda i, c: "=" + "+".join(f"«sub_{p}»" for p, _, _ in SEATS), "", name="REV_SubSeats")
    R.add("sub", "訂閱合計", "$B", lambda i, c: "=«sub_cons»+«sub_seat»", "", name="REV_Sub")

    R.section("七、API 營收（$B＝計費 token（T）× 組合有效單價 ÷ 1000）")
    R.add("calib", "API 營收校準值（2025＝SRC 用量計價營收；2026＝總額基準 − 訂閱 − 其他 − 廣告）", "$B",
          [f"={S('rev_mix_2025_consumption')}", "=«target»-«sub»-«other»-«ads»"] + [None] * 4, "Demand 的 2025、2026 API token 由此倒推（唯一倒推項）")
    R.add("api", "API 營收", "$B", lambda i, c: f"=«D:api_tok»*«mixprice»/{m2b}", "計費 token × 組合有效單價", name="REV_API")
    for t, tz, T in TIERS:
        R.add(f"api_{t}", f"API 營收：{tz}", "$B", lambda i, c, t=t: f"=«D:api_tok_{t}»*«ep_{t}»/{m2b}", "", name=f"REV_API_{T}")
    R.add("apichk", "檢查：層級合計 − API 營收（應為 0）", "$B", lambda i, c: "=" + "+".join(f"«api_{t}»" for t, _, _ in TIERS) + "-«api»", "")

    R.section("八、其他與廣告")
    R.add("other", "其他營收（政府、授權等；2025＝認列營收 − 訂閱 − 用量計價）", "$B",
          grow("other", f"={S('rev_2025_recognized')}-{S('rev_mix_2025_subscription')}-{S('rev_mix_2025_consumption')}", lambda i: I("g_other")),
          "2025 殘差（含四捨五入）；2026 起 ×（1＋年增率）", name="REV_Other")
    R.add("ads", "廣告營收（D3：不放廣告）", "$B", lambda i, c: f"={I('ads_revenue')}", "Decision＝0", name="REV_Ads")

    R.section("九、總額、雲端平台抽成（D5 r1）、淨額；個人 vs 企業")
    R.add("gross", "營收總額（報導口徑：雲端市集以總額認列）", "$B", lambda i, c: "=«sub»+«api»+«other»+«ads»", "", name="REV_Gross")
    R.add("ch_share", "雲端通路（AWS Bedrock、Google Vertex 等）占 API 營收比例", "比例",
          lambda i, c: f"={S('rev_2025_recognized')}*{S('rev_mix_2025_cloud_partner_share')}/{S('rev_mix_2025_consumption')}" if i == 0 else f"=«ch_share@D»+{I('channel_delta')}",
          "2025＝認列營收 × 雲端通路占比 47%（SRC）÷ 用量計價營收（通路營收全歸 API）；2026 起＝2025 ＋ Inputs 變動")
    R.add("ch_rev", "雲端通路營收", "$B", lambda i, c: "=«api»*«ch_share»", "", name="REV_Channel")
    R.add("fee_rate", "平台抽成率（每 $1 通路營收）", "比例", lambda i, c: f"={S('rev_partner_fee_2025')}/«ch_rev»" if i == 0 else f"={I('partner_fee_rate')}",
          "2025＝報導通路費 ÷ 通路營收（隱含值）；2026 起 Inputs（基準 16%，區間至 BofA 推得 30%）")
    R.add("pshare", "雲端平台抽成", "$B", lambda i, c: f"={S('rev_partner_fee_2025')}" if i == 0 else "=«ch_rev»*«fee_rate»",
          "2025＝SRC 報導值 $0.351B；2026 起＝通路營收 × 抽成率。說明書將抽成記為行銷費用，本模型改自營收扣除（不再列入非算力費用，避免重複）", name="REV_PartnerShare")
    R.add("net", "營收淨額（總額 − 平台抽成；命題與現金流用此）", "$B", lambda i, c: "=«gross»-«pshare»", "", name="REV_Net")
    R.add("cons", "個人（Pro、Max 訂閱；含廣告）", "$B", lambda i, c: "=«sub_cons»+«ads»", "D2：個人 vs 企業", name="REV_Consumer")
    R.add("ent", "企業（Team、Enterprise 席位＋API）", "$B", lambda i, c: "=«sub_seat»+«api»", "", name="REV_Enterprise")
    R.add("entshare", "企業占（個人＋企業）", "比例", lambda i, c: "=«ent»/(«cons»+«ent»)", "對照：企業與商業客戶約占 80%（2025-10，SRC_ANT_036）")
    R.add("sumchk", "檢查：個人＋企業＋其他 − 總額（應為 0）", "$B", lambda i, c: "=«cons»+«ent»+«other»-«gross»", "")
    R.add("growth", "總額年增率", "比例", lambda i, c: None if i == 0 else "=(«gross»-«gross@p»)/«gross@p»", "")
    R.add("subshare", "訂閱占總額", "比例", lambda i, c: "=«sub»/«gross»", "對照：2025 說明書約 17%；Sacra 2026 估 10–15%（SRC_ANT_062）")

    R.section("十、容量上限（Compute 的 CMP_CapFactor；2025、2026 固定 1）")
    R.add("cap", "容量上限係數（＝Compute 容量上限係數 CMP_CapFactor）", "倍", lambda i, c: None, "A3 填入：＝Compute 同欄（推論 GW 截頂後 ÷ 有效推論 GW）；2025、2026 為校準年固定 1",
          name="REV_CapFactor")
    R.add("gross_c", "營收總額（截頂後）", "$B", lambda i, c: "=«gross»*«cap»", "", name="REV_GrossCapped")
    R.add("pshare_c", "雲端平台抽成（截頂後）", "$B", lambda i, c: "=«pshare»*«cap»", "", name="REV_PartnerShareCapped")
    R.add("net_c", "營收淨額（截頂後）", "$B", lambda i, c: "=«gross_c»-«pshare_c»", "", name="REV_NetCapped")

    R.section("十一、對照與差距（只列差距，不校準；D14、D15 r1、D23）")
    R.add("gap25", "2025：模型總額 − SRC 認列營收（校準，應為 0）", "$B", ["=«gross»-" + S("rev_2025_recognized")] + [None] * 5, "", name="REV_Gap2025", name_cols="D")
    R.add("gap26", "2026：模型總額 − 半校準基準（應為 0）", "$B", [None, "=«gross»-«target»"] + [None] * 4, "", name="REV_Gap2026", name_cols="E")
    R.add("unc_cons", "2025 未校準個人訂閱＝月活（SRC，Analogy）× 付費轉換率（Inputs）× 加權月費 × 12", "$B",
          [f"={S('users_mau_web_2025')}*{I('paid_conv_2025')}*«p_cons_w»*{months}/{m2b}"] + [None] * 5, "自下而上口徑")
    R.add("unc_gap", "2025 未校準個人訂閱 − 校準值", "$B", ["=«unc_cons»-«cal_cons»"] + [None] * 5, "", name="REV_GapUncal2025", name_cols="D")
    R.add("rr25", "2025 由 run-rate 推估＝（2025-12 − 2024-12 年化）÷ ln（兩者比）", "$B",
          [f"=({S('rev_runrate_2025_12')}-{S('rev_runrate_2024_12')})/LN({S('rev_runrate_2025_12')}/{S('rev_runrate_2024_12')})"] + [None] * 5,
          "年度營收≈當年各月年化營收的平均；年內等比成長時為對數平均（D14：只作對照）")
    R.add("rr25gap", "2025 模型總額 ÷ run-rate 推估 − 1", "比例", ["=(«gross»-«rr25»)/«rr25»"] + [None] * 5, "run-rate 為未明口徑，且月初／月底不一")
    mid = f"AVERAGE({S('proj_arr_2026_end_investors')}_Lo,{S('proj_arr_2026_end_investors')}_Hi)"
    R.add("rr26", "2026 里程碑推估＝上半年實際＋下半年（2026-07 年化 → 年底投資人預期中點的對數平均）× 6 ÷ 12", "$B",
          [None, f"={S('rev_2026_q1')}+{S('rev_2026_q2')}+({mid}-{S('rev_runrate_2026_07')})/LN({mid}/{S('rev_runrate_2026_07')})*{h2}/{months}"] + [None] * 4,
          "年底投資人預期 $100–120B（SRC_ANT_373，非公司指引）只作對照", name="REV_RREst2026", name_cols="E")
    R.add("rr26gap", "2026 模型總額 ÷ 里程碑推估 − 1（|差距| > 門檻 → Checks WARN）", "比例", [None, "=(«gross»-«rr26»)/«rr26»"] + [None] * 4, "D15：差距 >25% 立 WARN，不自動校準",
          name="REV_RRGap2026", name_cols="E")
    R.add("exit26", "2026 下半年隱含年化（模型）＝（總額 − 上半年實際）× 12 ÷ 6", "$B/年", [None, f"=(«gross»-{S('rev_2026_q1')}-{S('rev_2026_q2')})*{months}/{h2}"] + [None] * 4,
          "＝2026-07 年化（D15 r1 持平假設）")
    R.add("mgmt", "管理層營收預測（2026-01 版；只對照，不驅動）", "$B",
          [None, f"={S('proj_rev_2026_2026_01')}", f"={S('proj_rev_2027_2026_01')}", None, f"={S('proj_rev_2029_2026_01')}", None], "SRC_ANT_368–370（Interested-party）")
    R.add("mgmtgap", "模型總額 − 管理層預測", "$B", [None, "=«gross»-«mgmt»", "=«gross»-«mgmt»", None, "=«gross»-«mgmt»", None], "")
    R.add("bofa", "BofA 2026 付給雲端金額 ÷ 模型 2026 通路營收（隱含抽成率上緣）", "比例", [None, f"={S('rev_partner_fee_2026_est')}/«ch_rev»"] + [None] * 4,
          "BofA 口徑可能含算力租金（SRC_ANT_033 註），只作 Inputs 抽成率上緣的依據")

    # ── 十二、價格事件表（事件 × 年度：事件之後占當年的比例）
    ev0 = R.r + 1
    put(R.ws, f"A{ev0}", "十二、API 價格事件（層級牌價序列；價格與公告日取自 SRC_ANT；D–I＝該事件之後占當年的比例）", F_BOLD, fill=FILL_SEC)
    header(R.ws, ev0 + 1, ["編號", "模型（層級／輸入輸出）", "單位"] + [str(y) for y in YEARS] + ["", "說明", "", "SRC 牌價", "公告日", "Δ（本價 − 前價）"])
    rr = ev0 + 2
    ev_rng = {}
    for t, tz, T in TIERS:
        for io, ioz in (("in", "輸入"), ("out", "輸出")):
            first = rr
            for j, pk in enumerate(CHAINS[t]):
                key = f"{pk}_{io}"
                put(R.ws, f"A{rr}", f"E{T[0]}{io[0].upper()}{j + 1}", F_CALC)
                put(R.ws, f"B{rr}", f"{tz} {ioz}：{D.src(key)['metric']}", F_CALC)
                put(R.ws, f"C{rr}", "比例", F_CALC)
                put(R.ws, f"M{rr}", f"={S(key)}", fmt="0.000")
                put(R.ws, f"N{rr}", f"={S(key)}_Date", fmt="0")
                put(R.ws, f"O{rr}", f"=M{rr}" if j == 0 else f"=M{rr}-M{rr - 1}", fmt="0.000")
                for i, c in enumerate(YC):
                    yst, yen = R.rows["yst"], R.rows["yen"]
                    put(R.ws, f"{c}{rr}", f"=({c}${yen}-MAX({c}${yst},MIN({c}${yen},$N{rr})))/({c}${yen}-{c}${yst})", fmt="0.0%")
                put(R.ws, f"K{rr}", "序列起點（早於 2025，比例＝1）" if j == 0 else "", F_NOTE)
                rr += 1
            ev_rng[f"{t}_{io}"] = (first, rr - 1)
    R.ev = ev_rng

    # ═════════════════════ Demand ═════════════════════
    Dm = Sheet(wb, "Demand", "D", "D", years,
               "Demand — 需求：個人方案人數、企業席位、任務類別 × 每任務 token × 層級組合；API 任務數 × 每任務 token；token 量（DEM_Tok_*：層級 × 付費／免費）",
               "2025：個人付費人數＝個人訂閱營收 ÷ 加權月費（D14 校準）；企業席位同理；免費＝2025 月活 − 個人付費；API token＝API 營收 ÷ 組合有效單價（唯一倒推）。"
               "2026：API token 由半校準總額倒推（D15 r1）。2027 起 API＝任務數 ×（1＋成長）× 價格反應 (p/p前)^ε × 每任務 token（同 OpenAI v0.6）。",
               "人數 M＝百萬；token T＝10^12；每任務 K token＝千 token；API 任務數＝十億。訂閱與免費 token 不產生營收，只供 A3 算力（推論 GW）。")
    Dm.section("一、方案組合與人數（M，FY 平均）")
    for k25, k26, p in (("cons_mix_pro_2025", "cons_mix_pro_2026", "pro"), ("cons_mix_max5_2025", "cons_mix_max5_2026", "max5")):
        f = two(k25, k26)
        Dm.add(f"cmix_{p}", f"個人付費方案組合：{dict((a, b) for a, b, _ in CONS)[p]}", "比例", lambda i, c, f=f: f"={f(i)}", "Inputs（2025 一欄；2026 起一欄）")
    Dm.add("cmix_max20", "個人付費方案組合：Max 20x", "比例", lambda i, c: "=(1-«cmix_pro»-«cmix_max5»)", "＝1 − Pro − Max 5x")
    gc = lambda i: by_year("g_cons", i)  # noqa: E731
    Dm.add("u_cons", "個人付費人數合計", "M", grow("u_cons", f"=«R:cal_cons»*{m2b}/(«R:p_cons_w»*{months})", gc),
           "2025＝個人訂閱營收 ÷（加權月費 × 12）（工作單預設：付費人數無官方值時由營收倒推）；2026 起 ×（1＋年增率）", name="DEM_Users_PaidConsumer")
    for p, pz, P in CONS:
        Dm.add(f"u_{p}", f"{pz} 人數", "M", lambda i, c, p=p: f"=«u_cons»*«cmix_{p}»", "", name=f"DEM_Users_{P}")
    Dm.add("u_free", "Free 人數", "M", grow("u_free", f"={S('users_mau_web_2025')}-«u_cons»", lambda i: by_year("g_free", i)),
           "2025＝月活（SRC，第三方估計 1,890 萬）− 個人付費；2026 起 ×（1＋年增率）", name="DEM_Users_Free")
    for k25, k26, p in (("seat_mix_std_2025", "seat_mix_std_2026", "team_std"), ("seat_mix_prem_2025", "seat_mix_prem_2026", "team_prem")):
        f = two(k25, k26)
        Dm.add(f"smix_{p}", f"企業席位組合：{dict((a, b) for a, b, _ in SEATS)[p]}", "比例", lambda i, c, f=f: f"={f(i)}", "Inputs（2025 一欄；2026 起一欄）")
    Dm.add("smix_ent", "企業席位組合：Enterprise", "比例", lambda i, c: "=(1-«smix_team_std»-«smix_team_prem»)", "＝1 − 標準 − Premium")
    Dm.add("s_total", "企業席位數合計", "M", grow("s_total", f"=«R:cal_seat»*{m2b}/(«R:p_seat_w»*{months})", lambda i: by_year("g_seats", i)),
           "2025＝企業席位訂閱營收 ÷（加權月費 × 12）；2026 起 ×（1＋年增率）", name="DEM_Seats_Total")
    for p, pz, P in SEATS:
        Dm.add(f"s_{p}", f"{pz} 席位數", "M", lambda i, c, p=p: f"=«s_total»*«smix_{p}»", "", name=f"DEM_Seats_{P}")

    Dm.section("二、訂閱與席位使用量：每日任務數 × 任務類別 × 每任務 token（2026 起年增）")
    for p, pz, P in PLANS:
        Dm.add(f"tasks_{p}", f"{pz} 每人（席）每日任務數", "任務/日", grow(f"tasks_{p}", f"={I('tasks_' + p)}", lambda i: I("g_tasks")), "Inputs 2025；之後 ×（1＋年增率）")
    for c_, cz in CATS:
        task = I(f"tk_task_{c_}")
        tk = "+".join(f"SUMIFS(TK_IF_TaskTok{x},TK_IF_HdrTask,{task})" for x in ("Fresh", "Cached", "Dec"))
        Dm.add(f"k_{c_}", f"每任務 token：{cz}", "K tok/任務", grow(f"k_{c_}", f"=({tk})/{I('tok_per_k')}*{I('k_ratio_' + c_)}", lambda i: I("g_k")),
               "2025＝Tokenomics 對應任務每次嘗試 token（新鮮＋快取＋decode）× Inputs 倍數；之後 ×（1＋年增率）")

    def share(p, c_):
        if c_ == "other":
            return f"(1-{I('cat_chat_' + p)}-{I('cat_code_' + p)})"
        return I(f"cat_{c_}_{p}")

    def users(p):
        return "u_free" if p == "free" else (f"u_{p}" if p in dict((a, 1) for a, _, _ in CONS) else f"s_{p}")

    for p, pz, P in PLANS:
        mixk = "+".join(f"{share(p, c_)}*«k_{c_}»" for c_, _ in CATS)
        Dm.add(f"tok_{p}", f"{pz} token", "T", lambda i, c, p=p, mixk=mixk: f"=«{users(p)}»*«tasks_{p}»*({mixk})*{days}/{b2t}",
               "人數（M）× 每日任務 × Σ 類別占比 × 每任務 K token × 天數 ÷ 1000", name=f"DEM_Tok_Plan_{P}")
    Dm.add("tok_sub", "訂閱、席位與 Free token 合計", "T", lambda i, c: "=" + "+".join(f"«tok_{p}»" for p, _, _ in PLANS), "")
    for c_, cz in CATS:
        Dm.add(f"ct_{c_}_paid", f"付費方案 token：{cz}", "T",
               lambda i, c, c_=c_: "=(" + "+".join(f"«{users(p)}»*«tasks_{p}»*{share(p, c_)}" for p, _, _ in PAID) + f")*«k_{c_}»*{days}/{b2t}",
               "Σ 付費方案（Pro、Max、席位）人數 × 每日任務 × 類別占比 × 該類每任務 token")
    for c_, cz in CATS:
        Dm.add(f"ct_{c_}_free", f"Free token：{cz}", "T", lambda i, c, c_=c_: f"=«u_free»*«tasks_free»*{share('free', c_)}*«k_{c_}»*{days}/{b2t}", "")
    Dm.add("catchk", "檢查：（類別合計 − 方案合計）÷ 方案合計（應為 0）", "比例",
           lambda i, c: "=((" + "+".join(f"«ct_{c_}_paid»+«ct_{c_}_free»" for c_, _ in CATS) + ")-«tok_sub»)/«tok_sub»", "以相對差表示，避免大數相減的浮點雜訊")
    Dm.add("code_share", "程式代理占訂閱、席位與 Free token（D4：Claude Code 不另立營收線）", "比例",
           lambda i, c: "=(«ct_code_paid»+«ct_code_free»)/«tok_sub»", "Claude Code 營收落在 Max、Team Premium 訂閱與 API 用量")

    Dm.section("三、API 需求：任務數 × 每任務 token（2025、2026 由營收倒推；2027 起任務成長 × 每任務 token 成長 × 價格反應）")
    Dm.add("api_b", "價格反應底數 p/p前（API 組合有效單價）", "倍", lambda i, c: None if i == 0 else "=«R:mixprice»/«R:mixprice@p»",
           "價格反應只在年度 ≥ 彈性起算年（Inputs 2027）生效：底數^(ε ×（年度 ≥ 起算年））")
    ek, en = I("api_elast_k"), I("api_elast_n")
    Dm.add("api_k", "API 每任務 token", "K tok/任務",
           lambda i, c: f"={I('api_k_base')}" if i == 0 else f"=«api_k@p»*(1+{by_year('api_gk', i)})*«api_b»^({ek}*({c}$6>={start}))",
           "層級正規化（只影響任務數與每任務 token 的拆分，不影響營收）；之後 ×（1＋成長）× 價格反應")
    Dm.add("api_n", "API 任務數", "十億任務",
           lambda i, c: "=«api_tok»/«api_k»" if i < 2 else f"=«api_n@p»*(1+{by_year('api_gn', i, 2027)})*«api_b»^({en}*({c}$6>={start}))",
           "2025、2026＝計費 token ÷ 每任務 token；2027 起 ×（1＋成長）× 價格反應")
    Dm.add("api_tok", "API 計費 token", "T", lambda i, c: f"=«R:calib»*{m2b}/«R:mixprice»" if i < 2 else "=«api_n»*«api_k»",
           "2025、2026：API 營收校準值 ÷ 組合有效單價（D14、D15 r1）；之後＝任務數 × 每任務 token", name="DEM_Tok_API")
    Dm.add("api_gn", "API 任務數年增率（2026＝半校準隱含值）", "比例/年", lambda i, c: None if i == 0 else "=(«api_n»-«api_n@p»)/«api_n@p»", "2027 起含價格反應", name="DEM_ApiTaskGrowth")
    for t, tz, T in TIERS:
        Dm.add(f"api_tok_{t}", f"API token：{tz}", "T", lambda i, c, t=t: f"=«api_tok»*«R:mix_{t}»", "計費 token × 層級組合（Revenue）", name=f"DEM_Tok_API_{T}")

    Dm.section("四、輸出：token 量（層級 × 付費／免費；A3 推論 GW 用）")

    def tmix(c_, t):
        if t == "top":
            return f"(1-{I('tier_haiku_' + c_)}-{I('tier_sonnet_' + c_)})"
        return I(f"tier_{TIER_INP[t]}_{c_}")

    for t, tz, T in TIERS:
        Dm.add(f"free_{t}", f"免費 token：{tz}", "T", lambda i, c, t=t: "=" + "+".join(f"«ct_{c_}_free»*{tmix(c_, t)}" for c_, _ in CATS),
               "Σ 類別 token × 層級組合（Inputs）", name=f"DEM_Tok_Free_{T}")
    for t, tz, T in TIERS:
        Dm.add(f"paid_{t}", f"付費 token：{tz}", "T",
               lambda i, c, t=t: "=" + "+".join(f"«ct_{c_}_paid»*{tmix(c_, t)}" for c_, _ in CATS) + f"+«api_tok_{t}»",
               "Σ 付費類別 token × 層級組合 ＋ API 該層級 token", name=f"DEM_Tok_Paid_{T}")
    Dm.add("free_all", "免費 token 合計", "T", lambda i, c: "=" + "+".join(f"«free_{t}»" for t, _, _ in TIERS), "", name="DEM_Tok_Free")
    Dm.add("paid_all", "付費 token 合計", "T", lambda i, c: "=" + "+".join(f"«paid_{t}»" for t, _, _ in TIERS), "", name="DEM_Tok_Paid")
    Dm.add("tok_all", "token 合計（付費＋免費）", "T", lambda i, c: "=«free_all»+«paid_all»", "", name="DEM_Tok_Total")
    Dm.add("tok_chk", "對照：訂閱、席位、Free 與 API 計費 token（層級拆分前；應＝合計）", "T", lambda i, c: "=«tok_sub»+«api_tok»", "Checks 比對")
    Dm.add("api_tok_share", "API 占 token 合計", "比例", lambda i, c: "=«api_tok»/«tok_all»", "")

    Dm.section("五、對照（只列示）")
    Dm.add("cc_rr", "Claude Code 年化營收（SRC：2025-12、2026-02；D4 不另立營收線）", "$B/年", [f"={S('cc_runrate_2025_12')}", f"={S('cc_runrate_2026_02')}"] + [None] * 4,
           "只對照；其營收已含在 Max、Team Premium 訂閱與 API 用量")

    # ── 解析列鍵記號（兩頁都建好後）
    def evm(key):
        a, z = ev_rng[key]
        return lambda col: f"$O${a}:$O${z}"

    def evc(key):
        a, z = ev_rng[key]
        return lambda col: f"{col}${a}:{col}${z}"

    extra = {}
    for key in ev_rng:
        extra[f"EVM_{key}"] = evm(key)
        extra[f"EVC_{key}"] = evc(key)
    R.resolve(extra)
    Dm.resolve()
    # 頁序：Demand 在 Revenue 之前（因果鏈）
    wb.move_sheet("Demand", offset=-1)
    for sh, wB in ((Dm.ws, 52), (R.ws, 60)):
        sh.column_dimensions["A"].width = 7
        sh.column_dimensions["B"].width = wB
        sh.column_dimensions["C"].width = 10
        for c in YC:
            sh.column_dimensions[c].width = 12
        sh.column_dimensions["K"].width = 70
        sh.freeze_panes = "D7"
    R.ws.column_dimensions["M"].width = 10
    R.ws.column_dimensions["N"].width = 11
    R.ws.column_dimensions["O"].width = 10
    ctx.P["a2"] = dict(D=Dm, R=R)


def checks(ctx):
    D = ctx.D
    S, I = D.S, D.I
    Dm, R = ctx.P["a2"]["D"], ctx.P["a2"]["R"]
    dr, rr = Dm.rows, R.rows
    yr = lambda sh, key: f"{sh.name}!$D${sh.rows[key]}:$I${sh.rows[key]}"  # noqa: E731
    mx = lambda sh, key: f"=MAX(MAX({yr(sh, key)}),-MIN({yr(sh, key)}))"  # noqa: E731

    def viol(a, b):
        return f"({a}<0)+({a}>1)+({b}<0)+({b}>1)+((1-{a}-{b})<0)"

    pairs = [(I("cons_mix_pro_2025"), I("cons_mix_max5_2025")), (I("cons_mix_pro_2026"), I("cons_mix_max5_2026")),
             (I("seat_mix_std_2025"), I("seat_mix_prem_2025")), (I("seat_mix_std_2026"), I("seat_mix_prem_2026")),
             (I("api_mix_haiku_2025"), I("api_mix_sonnet_2025")), (I("api_mix_haiku_2026"), I("api_mix_sonnet_2026"))]
    pairs += [(I(f"cat_chat_{p}"), I(f"cat_code_{p}")) for p, _, _ in PLANS]
    pairs += [(I(f"tier_haiku_{c_}"), I(f"tier_sonnet_{c_}")) for c_, _ in CATS]
    order = []
    for key, (a, z) in R.ev.items():
        order += [f"(Revenue!$N${r}<Revenue!$N${r - 1})" for r in range(a + 1, z + 1)]
    return [
        ("a2_gap25", "2025 營收校準：模型總額 − SRC 認列營收（$B）", "=REV_Gap2025", 0, "tol", "D14；訂閱、API、其他三項分別等於 SRC，總額恆等於 SRC_ANT_001"),
        ("a2_sub25", "2025 訂閱：模型 − SRC 訂閱營收（$B）", f"=Revenue!$D${rr['sub']}-{S('rev_mix_2025_subscription')}", 0, "tol", "個人＋席位由營收倒推人數"),
        ("a2_api25", "2025 API：模型 − SRC 用量計價營收（$B）", f"=Revenue!$D${rr['api']}-{S('rev_mix_2025_consumption')}", 0, "tol", "API token 由營收 ÷ 有效單價倒推"),
        ("a2_ps25", "2025 雲端平台抽成：模型 − SRC 報導通路費（$B）", f"=Revenue!$D${rr['pshare']}-{S('rev_partner_fee_2025')}", 0, "tol", "D5 r1"),
        ("a2_gap26", "2026 半校準：模型總額 − 基準（Q1＋Q2＋2026-07 年化 × 6／12）（$B）", "=REV_Gap2026", 0, "tol", "D15 r1；API 為倒推項"),
        ("a2_rr26", "2026 對照：模型總額 ÷ 里程碑推估 − 1（|差距| > 門檻 → WARN）", "=REV_RRGap2026", I("runrate_warn_threshold"), "warn",
         "D15：只提示、不校準；推估＝上半年實際＋下半年由 2026-07 年化至年底投資人預期中點的對數平均"),
        ("a2_sumchk", "個人＋企業＋其他 − 總額（各年最大絕對差，$B）", mx(R, "sumchk"), 0, "tol", ""),
        ("a2_apichk", "API 層級合計 − API 營收（各年最大絕對差，$B）", mx(R, "apichk"), 0, "tol", ""),
        ("a2_toksplit", "token 層級拆分合計 vs 拆分前總量（違反年數）",
         f"=SUMPRODUCT(--(ABS({yr(Dm, 'tok_all')}-{yr(Dm, 'tok_chk')})>0.000000001*{yr(Dm, 'tok_chk')}))", 0, "eq", "付費＋免費（層級）＝訂閱、席位、Free＋API 計費 token"),
        ("a2_catchk", "任務類別合計 vs 方案合計（違反年數）",
         f"=SUMPRODUCT(--(ABS({yr(Dm, 'catchk')})>0.000000001))", 0, "eq", "相對差"),
        ("a2_mix", "組合占比 ∉ [0,1] 或餘項 < 0 的違反數（方案、席位、API 層級、任務類別、層級組合）", "=" + "+".join(viol(a, b) for a, b in pairs), 0, "eq",
         "Inputs 只設兩項、第三項＝1 − 其餘"),
        ("a2_backsolve", "2025 倒推：個人付費 ≤0、席位 ≤0、Free <0、API token ≤0 的違反數",
         f"=(Demand!$D${dr['u_cons']}<=0)+(Demand!$D${dr['s_total']}<=0)+(Demand!$D${dr['u_free']}<0)+(Demand!$D${dr['api_tok']}<=0)", 0, "eq",
         "例：消費者訂閱占營收（SRC Analogy）取高值 20% 時個人訂閱 0.92 > 訂閱合計 0.789，席位轉負"),
        ("a2_api26", "2026 API 校準營收 ≤ 0（半校準後訂閱＋其他已超過總額）", f"=IF(Revenue!$E${rr['calib']}<=0,1,0)", 0, "eq", ""),
        ("a2_cap", "容量上限係數 ∉ (0,1] 的年數", "=SUMPRODUCT(--(REV_CapFactor<=0))+SUMPRODUCT(--(REV_CapFactor>1.000000001))", 0, "eq", "A3 起＝Compute CMP_CapFactor"),
        ("a2_ads", "廣告營收 ≠ 0 的年數（D3）", "=SUMPRODUCT(--(REV_Ads<>0))", 0, "eq", "Anthropic 公開承諾不放廣告"),
        ("a2_evorder", "價格事件公告日順序違反數（各層級序列須依日期遞增）", "=" + "+".join(order), 0, "eq", "SRC 日期被改錯時轉 ERR"),
        ("a2_tktask", "任務類別對應的 Tokenomics 任務名稱在 TK_IF_HdrTask 找不到的數",
         "=" + "+".join(f"(COUNTIF(TK_IF_HdrTask,{I('tk_task_' + c_)})<>1)" for c_, _ in CATS), 0, "eq", "Tokenomics 改任務名稱時轉 ERR"),
        ("a2_tkcache", "對照：Tokenomics 程式代理任務的快取輸入占輸入 token（vs Inputs API 快取命中）",
         f"=SUMIFS(TK_IF_TaskTokCached,TK_IF_HdrTask,{I('tk_task_code')})/(SUMIFS(TK_IF_TaskTokCached,TK_IF_HdrTask,{I('tk_task_code')})+SUMIFS(TK_IF_TaskTokFresh,TK_IF_HdrTask,{I('tk_task_code')}))",
         None, "info", f"Inputs {I('api_cache_hit')} 基準 0.6（API 全體）"),
        ("a2_src12", "對照：2025 營收兩筆 SRC 之差（4.6 − 4.59，$B）", f"={S('rev_2025_recognized')}-{S('rev_2025_recognized_b')}", None, "info", "基準用 SRC_ANT_001（Reuters 經 Fortune）"),
        ("a2_unc25", "對照：2025 未校準個人訂閱 − 校準值（$B）", "=REV_GapUncal2025", None, "info", "月活 × 付費轉換率（Analogy）× 加權月費"),
        ("a2_rr25", "對照：2025 模型總額 ÷ run-rate 對數平均推估 − 1", f"=Revenue!$D${rr['rr25gap']}", None, "info", "D14：run-rate 只作對照"),
        ("a2_fee25", "對照：2025 隱含平台抽成率（報導通路費 ÷ 通路營收）", f"=Revenue!$D${rr['fee_rate']}", None, "info", "Inputs 2026 起基準 0.16"),
        ("a2_bofa", "對照：BofA 2026 付給雲端金額 ÷ 模型 2026 通路營收", f"=Revenue!$E${rr['bofa']}", None, "info", "Inputs 抽成率上緣 0.30 的依據"),
        ("a2_subshare26", "對照：2026 訂閱占總額（Sacra 估 10–15%）", f"=Revenue!$E${rr['subshare']}", None, "info", "SRC_ANT_062（Analogy）"),
        ("a2_apig26", "對照：2026 API 任務數隱含年增率（半校準）", f"=Demand!$E${dr['api_gn']}", None, "info", "D15 r1：2026 由實際值倒推"),
        ("a2_mgmt27", "對照：2027 模型總額 − 管理層預測 $55B（2026-01 版）", f"=Revenue!$F${rr['mgmtgap']}", None, "info", "D19：只對照、不回饋"),
        ("a2_net30", "結果：2030 營收淨額（$B）", "=INDEX(REV_Net,1,6)", None, "info", "總額 − 雲端平台抽成"),
        ("a2_oai25", "OpenAI 並排：2025 營收總額 Anthropic ÷ OpenAI", "=INDEX(REV_Gross,1,1)/INDEX(OAI_REV_Gross,1,1)", None, "info", "OAI_Link（D20）"),
        ("a2_oai30", "OpenAI 並排：2030 營收淨額 Anthropic ÷ OpenAI（未截頂；OpenAI 為截頂前淨額）", "=INDEX(REV_Net,1,6)/INDEX(OAI_REV_Net,1,6)", None, "info", ""),
        ("a2_oaitok30", "OpenAI 並排：2030 token 合計 Anthropic ÷ OpenAI", "=INDEX(DEM_Tok_Total,1,6)/INDEX(OAI_DEM_Tok_Total,1,6)", None, "info", "需求口徑不同，只並列"),
    ]
