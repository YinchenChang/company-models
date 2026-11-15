"""智譜 v0.1.6：估值頁 Val 與反向 DCF 頁 RevDCF（Andy 2026-11-15 決定智譜與 MiniMax 都給目標價與評等）。

排版與公式結構在 builder/valcore.py（與 minimax/valcore.py 同一份程式）；本檔只提供參照：
- 現金流沿用模型既有路徑（Funding 頁 F01 營收、F06 差額〔含股權報酬〕＝EBIT、F10 非算力資本支出、F12 併購）：2H26 取 Funding L 欄（驅動），
  2027–2030 取 F–I 欄。股權報酬不加回（個股DCF標準規格第 1 節）。算力為服務採購／租賃付款，已在 EBIT 扣除（折舊 − 資本支出＝0；自有算力資本支出 D11r 基準 0）。
- 評價日 2026-06-30（中期報告資產負債表日）；之後完成的 2026-07、2026-09 配售與可換股債券淨額計入淨現金，股數取 2026-09 配售後總股本；
  現價 2026-11-05 收盤 HK$476（SRC_ZP_1055）。可換股債券換股價 HK$892.5 高於現價 → 視為債務（本金 201.4 億）。
- 情境 A／C：2H26 API 任務數成長率取 Inputs 高／低端（INP_038；樣板 Anthropic 以 API 任務數高／低端定情境），快照 data/val_scenarios.yaml（tools/val_snapshots.py）。
- 同業：MiniMax（00100.HK），市值與智譜同日（2026-11-05 收盤）；其餘中國與海外上市 AI 模型／應用公司找不到同日可查證市值，未納入（README 列待確認）。
Val 與 RevDCF 不被任何既有計算頁或 Checks 引用；新增後既有工作表數值不變（快照狀態列在 Val 頁內）。
"""
from __future__ import annotations

from pathlib import Path

import yaml

import valcore

ROOT = Path(__file__).resolve().parent.parent
SNAP_YAML = ROOT / "data" / "val_scenarios.yaml"
SENS_YAML = ROOT / "data" / "val_sens.yaml"
ALIAS = {"half": "val_half", "val_pct": "pct"}


def build(wb, D, Z):
    S, I0 = D.S, D.I
    I = lambda k: I0(ALIAS.get(k, k))  # noqa: E731
    F = Z["F"].rows
    snap = yaml.safe_load(SNAP_YAML.read_text(encoding="utf-8")) if SNAP_YAML.exists() else {}
    sens = yaml.safe_load(SENS_YAML.read_text(encoding="utf-8")) if SENS_YAML.exists() else {}
    f = lambda key, col: f"Funding!${col}${F[key]}"  # noqa: E731
    src_B = {k: [f"={f(key, col)}" for col in "LFGHI"] for k, key in (("rev", "rev"), ("ebit", "gap"), ("cx", "capex"), ("ma", "ma"))}
    src_B["fy"] = [f"={f('rev', col)}" for col in "DEF"]
    # 2H26 的併購欄（L）在 Funding 為空白時視為 0（SUM 對空白格回 0）
    src_B["ma"][0] = f"=SUM({f('ma', 'L')})"
    yi, hk = I("div_yuan_yi"), I("fx_hkdcny")
    musd = I("div_mn_yi")
    mmx_fx = S("peer_mmx_hkdusd")
    cfg = dict(
        I=I, unit="RMB 億", ps_unit="RMB／股", cur="人民幣",
        title="Val — 估值頁（v0.1.6；Andy 2026-11-15 決定給目標價與評等）：DCF（FCFF）50%＋同業 EV／NTM 營收 50%；情境 A／B／C 25／50／25；評價日 2026-06-30、現價 2026-11-05 收盤",
        sub="FCFF＝EBIT（Funding F06：營收 − 算力 − 本地化交付 − 非算力 − 股權報酬；股權報酬不加回）− 稅（25% × 正 EBIT，扣評價日起自身虧損）− 非算力資本支出 − 已承諾併購 − 營運資金增加。"
            "終值＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)。情境 A、C 為快照值（藍字；tools/val_snapshots.py），B 為活公式。判斷類參數（WACC 組成、同業、權重）依樣板慣例預設，待 Andy 確認。",
        sub2="本頁不被任何既有計算頁引用。單位：人民幣億元、億股、人民幣元／股；目標價以港元表示（÷ HKD/CNY 中間價）。",
        val_date="2026-06-30（pro forma：計入其後完成的配售與可換股債券）", px_date="2026-11-05",
        px=S("price_20261105"), px_note="SRC_ZP_1055（Yahoo Finance、StockAnalysis 兩來源一致）",
        pxfx=hk, pxfx_note="HKD/CNY 中間價 2026-09-28（Inputs fx_hkdcny）",
        shares=f"{S('shares_total_post_placing2')}/{yi}", shares_unit="億股", shares_label="評價股數＝2026-09 配售後已發行股份總數（億股）",
        shares_note="SRC_ZP 2026-09 配售完成後總股本；受限制股份有揭露（SRC_ZP_876，招股章程附註 29(b)：1H25 尚未行使 6,620 千份〔拆股前〕，1 拆 10 後約 6,620 萬股，行使價 0.10–2.00 元），未計入評價股數；若全部為增量約占 13.6%，加權目標價約降至 138–140 港元，評等不變（是否已含在已發行股數內未查證）；可換股債券價外，不計換股股數",
        cash_items=[
            ("2026-06-30 現金及現金等價物（中期報告）", S("cash_end_1h26"), "SRC_ZP_211"),
            ("＋2026-07 H 股配售淨額（07-13 完成）", f("pl1", "L"), "Funding F15"),
            ("＋2026-09 H 股配售淨額（09-16 完成）", f("pl2", "L"), "Funding F16"),
            ("＋2026-09 可換股債券淨額（09-18 發行）", f("cb", "L"), "Funding F18"),
            ("− 銀行借款（2026-06-30）", f"-{S('bank_loans_1h26')}", "SRC_ZP_213"),
            ("− 可換股債券本金（價外，視為債務）", f"-{S('cb_principal')}", "SRC_ZP_545"),
            ("− 租賃負債（2025-12-31；1H26 未揭露，沿用最近一期）", f"-{S('lease_liab_fy25')}", "SRC_ZP 年度業績公告"),
        ],
        debt=f"{S('bank_loans_1h26')}+{S('cb_principal')}+{S('lease_liab_fy25')}", debt_note="銀行借款＋可換股債券本金＋租賃負債（個股DCF標準規格第 2 節：債務本金＋融資租賃）",
        rf_note="美國 10 年期公債 CMT 4.86%（2026-11-06；與 Anthropic v0.2.0 同一來源 SRC_ANT_576）；港股以港元計價、聯繫匯率，沿用美元無風險利率（判斷類）",
        beta_note="Analogy：Anthropic v0.2.0 同業（高成長 AI／資料軟體）5 年 β 中位數 1.37 經 Blume 調整＝1.248；智譜 2026-01 上市，自身 β 期間不足（判斷類）",
        ttgt_note="評價日 2026-06-30 至現價日約 0.35 年 ＋ 12 個月",
        src_zh={"rev": "營收淨額（截頂後；Funding F01）", "ebit": "EBIT＝差額（含股權報酬；Funding F06）", "cx": "非算力資本支出（Funding F10）", "ma": "併購現金支出（Funding F12）"},
        scn_desc={"A": "2H26 API 任務數成長率取 Inputs 高端（INP_038 高欄）", "B": "基準（活公式）", "C": "2H26 API 任務數成長率取 Inputs 低端（INP_038 低欄）"},
        sens=sens or {},
        snap=(snap or {}).get("scenarios", {}), snap_note=f"{(snap or {}).get('model', '')}；tools/val_snapshots.py",
        snap_check=((snap or {}).get("scenarios", {}).get("B", {}) or {}).get("check"),
        src_B=src_B,
        ebit_note="Funding F06（營收 − 算力現金成本 − 本地化交付 − 非算力 − 股權報酬）；利息與其他收益不計（FCFF 口徑）",
        cx_note="Funding F10 非算力資本支出；自有算力資本支出已在算力成本內（D11r 基準 0）",
        ma_note="Funding F12：中科加禾約 60%（2H26，3 億元，Assumed）",
        peer_date_note="市值一律 2026-11-05 收盤；淨現金為 2026-06-30 pro forma；共識營收為最近一次彙整",
        peers=[dict(name="MiniMax（00100.HK）", unit="USD M", px=S("peer_mmx_price_20261105"), pxfx=f"1/{mmx_fx}", shares=f"{S('peer_mmx_shares')}/{I('mul_t_m')}",
                    netcash=(f"{S('peer_mmx_cash_0630')}+{S('peer_mmx_placing_net_hkd')}/{mmx_fx}+{S('peer_mmx_cb_net_hkd')}/{mmx_fx}"
                             f"-{S('peer_mmx_bank_0630')}-{S('peer_mmx_cb_principal_hkd')}/{mmx_fx}"),
                    c26=S("peer_mmx_cons_fy26"), c27=S("peer_mmx_cons_fy27"),
                    note="收盤價 2026-11-05 HK$170.2（MiniMax 模型 S113）；股數 2026-09-30 FF301；淨現金＝6-30 現金 1,322.8＋07 配售與可轉債淨額 − 銀行借款 − 可轉債本金（港元 ÷ 7.78）；"
                         "共識 FY2026 452、FY2027 1,030（stockanalysis）；與 MiniMax 模型 v0.2.3 同一組數字")],
        mult_note="同業只有 MiniMax 一家（中國與海外上市 AI 模型／應用公司中，取得 2026-11-05 同日可查證市值者）；其餘候選未納入，見 README（判斷類，待 Andy 確認）",
        self=dict(c26=f"{S('cons_rev_fy26')}/{musd}", c27=f"{S('cons_rev_fy27')}/{musd}", note="Simply Wall St 共識（23 位）FY2026 60.48、FY2027 180.13 億元（2026-10-06）"),
        gate="SUM(Funding!$E$%d:$I$%d)" % (F["fcf"], F["fcf"]),
        gate_note="Funding F09 自由現金流 FY2026–FY2030 合計（與 AI 投資地圖收支關卡同一口徑）",
        cons_tp=S("cons_tp_avg"), cons_label="Investing.com 24 位平均，2026-10-03",
        cons_note="SRC_ZP_587（19 買入、4 持有、1 賣出）；日期早於現價日一個月，股價其間大跌，僅作對照",
        method=("DCF（FCFF；EBIT＝Funding F06、股權報酬不加回；稅 25% 扣自身虧損；WACC＝CAPM〔rf 4.86%、β 1.248、ERP 5%〕與債務加權；終值 RONIC＝WACC、g 3%；"
                "延伸期 2031–2033 營收年增收斂到 10%、利益率維持 FY2030）50%＋同業 EV／NTM 營收（MiniMax，同日市值；不做相對倍數調整）50%；情境 A／B／C 25／50／25"),
    )
    P = valcore.build_val(wb, cfg)
    R = valcore.build_revdcf(wb, cfg, P)
    return dict(P=P, R=R, snap=snap, sens=sens)

