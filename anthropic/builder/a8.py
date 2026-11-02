"""Anthropic v0.1.7.1：正式 S-1 原文落地＋已到位融資更正（工作單 20261102 v0.1.7.1；Funding 第十三節尾 1 組列、第十四節，皆附加在頁尾，既有列編號不變）。

- V2 已到位融資對帳：正式 S-1 Note 4、9、10 的各輪金額（Q1 Series G 30.2、Q2 Series G-2 15.0〔Google 38.6 百萬股、Amazon 19.3 百萬股，每股 259.1364〕、
  Q2 Series H 65.2〔含 Google 10.0、Amazon 5.0〕）與合併現金流量表特別股發行收入 110.432 vs 模型 2026 已到位融資（FND_Committed）。
  已到位融資的公式本身在 a4（Funding 第二節；開關 googl_upfront_switch、s1f_round_switch）。
- 6-30 現金對照（第十三節）補 Broadcom 租約保證金（受限現金）約 2.5 與差額拆解（股權收入、上半年自由現金流時點、受限現金、其餘）；只對照。
- V4 最低現金依據：A&R 循環額度 15.0（或有，D16）、最低流動性條款 5.0（自 2026-12-31 季末起）vs 模型最低現金；只對照。
Python 只產生結構；數值一律引用 SRC_ANT／Inputs／既有列（E6）。
"""
from __future__ import annotations

from a2 import Sheet


def build(ctx):
    D = ctx.D
    S, I = D.S, D.I
    F = ctx.P["a4"]["F"]
    months, h2, m2b = I("months_per_year"), I("h2_months"), I("m_to_b")
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731

    # ═════════════ Funding 第十三節（續）：6-30 現金對照補充（列接在第十三節最後一列之後） ═════════════
    F.add("c630_bcm", "其中：Broadcom 租約保證金（受限現金）約 2.5（2026-04 存入；S-1 Note 16；不在 6-30 現金、約當現金與短期投資內）", "$B",
          only26(f"={S('cash_s1f_broadcom_deposit')}"), "SRC_ANT_487；預設不另計（之後抵付租金），只對照（v0.1.7.1）")
    F.add("c630_d_eq", "差額拆解 ①股權收入：S-1 特別股發行收入（淨額）− 模型已到位融資", "$B", only26(f"={S('fund_s1f_pref_2026h1')}-«committed»"),
          "SRC_ANT_472 − FND_Committed；V2 後約 0")
    F.add("c630_d_tm", "差額拆解 ②上半年自由現金流時點：S-1 上半年（營業現金流 − 資本支出）− 模型 2026 自由現金流 × 上半年月數 ÷ 12", "$B",
          only26(f"={S('ocf_s1f_2026h1')}-{S('capex_s1f_2026h1')}-«fcf»*{h2}/{months}"), "SRC_ANT_442、443；模型年內平均近似")
    F.add("c630_d_rc", "差額拆解 ③受限現金：Broadcom 租約保證金（S-1 現金不含；負號）", "$B", only26("=-«c630_bcm»"), "")
    F.add("c630_d_ot", "差額拆解 ④其餘＝差 − ① − ② − ③（利息與投資收益、其他融資與投資活動、2025 年底現金口徑等）", "$B",
          only26("=«c630_gap»-«c630_d_eq»-«c630_d_tm»-«c630_d_rc»"), "只對照，不校準", name="FND_Cash0630Resid", name_cols="E")

    # ═════════════ Funding 第十四節 ═════════════
    F.section("十四、v0.1.7.1：已到位融資對帳（正式 S-1 Note 4、9、10 與合併現金流量表；V2）、最低現金依據（S-1 信用額度條款；V4，只對照）")
    F.add("s1r_pref", "S-1：2026 上半年特別股發行收入（淨額；合併現金流量表 F-13）", "$B", only26(f"={S('fund_s1f_pref_2026h1')}"), "SRC_ANT_472（Verified）")
    F.add("s1r_g", "S-1：Q1 Series G-1／G-2（每股 259.1364）", "$B", only26(f"={S('fund_s1f_series_g_q1')}"), "SRC_ANT_473（Note 10）")
    F.add("s1r_g2", "S-1：Q2 Series G-2（Google 2026 NPA＋Amazon 2026 NPA）", "$B", only26(f"={S('fund_s1f_series_g2_q2')}"), "SRC_ANT_474（Note 10、Note 4）")
    F.add("s1r_g2g", "　其中 Google：38.6 百萬股 × 259.1364 ÷ 1000", "$B", only26(f"={S('fund_s1f_g2_googl_shares')}*{S('fund_s1f_g_price')}/{m2b}"),
          "SRC_ANT_475 × 477；＝已到位融資 Google G-2 列（開關 1 時）")
    F.add("s1r_g2a", "　其中 Amazon：19.3 百萬股 × 259.1364 ÷ 1000", "$B", only26(f"={S('fund_s1f_g2_amzn_shares')}*{S('fund_s1f_g_price')}/{m2b}"),
          "SRC_ANT_476 × 477；模型 Amazon G 列引用 Amazon 10-Q 5.0（SRC_ANT_319）")
    F.add("s1r_g2c", "　檢查：G-2 合計 − Google − Amazon（股數捨入）", "$B", only26("=«s1r_g2»-«s1r_g2g»-«s1r_g2a»"), "股數為百萬股一位小數，差在 ±0.02 內")
    F.add("s1r_h", "S-1：Q2 Series H-1／H-2（每股 589.0095）", "$B", only26(f"={S('fund_s1f_series_h_q2')}"), "SRC_ANT_478（Note 10）")
    F.add("s1r_hg", "　其中 Google（2026-05；Note 9）", "$B", only26(f"={S('fund_s1f_googl_series_h')}"), "SRC_ANT_480")
    F.add("s1r_ha", "　其中 Amazon（2026-05；Note 9）", "$B", only26(f"={S('fund_s1f_amzn_series_h')}"), "SRC_ANT_481（＝Amazon 10-Q，SRC_ANT_320）")
    F.add("s1r_sum", "S-1 各輪合計＝Series G＋Series G-2＋Series H", "$B", only26("=«s1r_g»+«s1r_g2»+«s1r_h»"), "")
    F.add("s1r_sumgap", "S-1 各輪合計 − 現金流量表特別股收入（發行成本與捨入）", "$B", only26("=«s1r_sum»-«s1r_pref»"), "")
    F.add("s1r_mod", "模型：2026 已到位融資（FND_Committed）", "$B", only26("=«committed»"), "第二節")
    F.add("s1r_gap", "差：模型 2026 已到位融資 − S-1 特別股發行收入", "$B", only26("=«s1r_mod»-«s1r_pref»"),
          "|差| ≤ Inputs s1f_cmt_tol（0.5）為 OK（Checks）；v0.1.7 讀法（googl_upfront_switch、s1f_round_switch 皆 0）為 −10.4", name="FND_CommittedS1Gap", name_cols="E")
    F.add("s1r_googl", "模型：2026 已到位融資中的 Google 部分（Series G-2〔開關〕＋Series H 內含 10）", "$B",
          only26(f"=«in_googl»+{I('s1f_round_switch')}*{S('fund_s1f_googl_series_h')}"), "對照 S-1 Note 9、10：Google 2026 共約 20.0")
    # V4
    yr = lambda c: f"({c}$6>={I('year_2026')})"  # noqa: E731
    F.add("s1m_liq", "S-1：信用額度最低流動性條款（自 2026-12-31 季末起）", "$B",
          lambda i, c: None if i == 0 else f"={S('debt_s1f_min_liquidity')}*{yr(c)}", "SRC_ANT_486（MD&A 第 163–164 頁；Note 17）")
    F.add("s1m_over", "模型最低現金 − 契約最低流動性（正＝模型較保守）", "$B", lambda i, c: None if i == 0 else "=«min»-«s1m_liq»",
          "模型最低現金＝次年非算力成本 × Inputs min_cash_months（6；Assumed 3–12）÷ 12（D17）", name="FND_MinCashOverCovenant")
    F.add("s1m_rev", "S-1：A&R 循環信用額度（2031-09 到期；或有，照 D16 不作資金來源）", "$B", only26(f"={S('debt_s1f_revolver')}"), "SRC_ANT_485；第六節同值")
    F.resolve()
    ctx.P["a8"] = dict(F=F)


def checks(ctx):
    I = ctx.D.I
    F = ctx.P["a8"]["F"]
    d = lambda key, col="E": f"={F.name}!${col}${F.rows[key]}"  # noqa: E731
    r = F.rows["s1m_over"]
    return [
        ("h2_cmt", "v0.1.7.1 V2：2026 已到位融資 − 正式 S-1 特別股發行收入 110.432（$B）", "=FND_CommittedS1Gap", I("s1f_cmt_tol"), "warn",
         "|差| ≤ 0.5 為 OK；V2 開關皆 0（v0.1.7 讀法）時約 −10.4 → WARN"),
        ("h2_round", "v0.1.7.1 V2：S-1 各輪合計 − 現金流量表特別股收入（$B）", d("s1r_sumgap"), None, "info", "發行成本與捨入"),
        ("h2_g2", "v0.1.7.1 V2：Series G-2 合計 − Google − Amazon（股數 × 每股價格；$B）", d("s1r_g2c"), None, "info", "股數捨入"),
        ("h2_c630", "v0.1.7.1：6-30 現金差額拆解的其餘項（$B）", "=FND_Cash0630Resid", None, "info", "利息、其他融資與投資活動等；只對照"),
        ("h4_liq", "v0.1.7.1 V4：模型最低現金 − 契約最低流動性 5.0 的 2026–2030 最小值（$B）", f"=MIN({F.name}!$E${r}:$I${r})", None, "info", "正＝模型最低現金高於契約下限"),
    ]
