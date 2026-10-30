"""Anthropic v0.1.6：正式公開 S-1 對照（工作單 20261031 W3；Cost 第二十四節，附加在 Cost 頁尾，既有列編號不變）。

- (a) S-1 算力與基礎設施費用 vs 模型 2026 算力成本（現金口徑）；口徑差異（費用 vs 現金、預付、自建資本支出）寫在說明欄。
- (b) S-1 營收成本（Note 8）vs 模型 2025 推論算力成本。
- (c) η 診斷：以 2026 上半年營收成本 ÷ 既有雲端加權合約價（與 2025 η 同法）推得推論 GW，得 η_2026H1；
      B 情境「η＝2026H1 值」（2025 沿用校準值，2026 起取 η_2026H1）：容量上限係數、每 VR 等值 GW 營收、差額、覆蓋率。基準 η 不改。
- (d) S-1 承諾三段（6-30 尚未起租租賃、6-30 固定或最低承諾、6-30 後新簽與修訂）vs 模型逐家合約總額（Cost 第十一、十二節口徑）。
本節只對照：不被任何其他計算列引用（Checks f1_iso 由 builder 掃描；tests 重驗）。數值一律引用 SRC_ANT／Inputs／既有列（E6：公式不含常數）。
"""
from __future__ import annotations

import re

from a2 import FMT, Sheet

FMT.setdefault("GW", "#,##0.000")
FMT.setdefault("$B/GW/年", "#,##0.000")

# 正式 S-1 逐家（合約鍵、名稱、SRC 鍵、欄位尾碼〔下限 _Lo／上限 _Hi／點值 ""〕）
S1F_PARTNERS = [("google", "Google Cloud（至少）", "cmt_s1f_google", ""), ("aws", "Amazon（至少）", "cmt_s1f_amazon", ""),
                ("azure", "Microsoft Azure（至少）", "cmt_s1f_msft", ""), ("spacex", "SpaceXAI（最高）", "cmt_s1f_spacex", ""),
                ("broadcom", "Broadcom 租賃（不含利息）", "cmt_s1f_broadcom", ""), ("nscale", "Nscale（最高）", "cmt_s1f_nscale", ""),
                ("lambda", "Lambda（約）", "cmt_s1f_lambda", ""), ("amd", "AMD（超過；不在承諾表內）", "cmt_s1f_amd", "")]
NAMED7 = [p for p in S1F_PARTNERS if p[0] != "amd"]


def build(ctx):
    D = ctx.D
    S, I = D.S, D.I
    K = Sheet.all["K"]
    months, h2, qpy, capfix = I("months_per_year"), I("h2_months"), I("quarters_per_year"), I("cap_fixed_year")
    idle = I("idle_share")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731
    first = K.r

    K.section("二十四、v0.1.6 正式 S-1 對照（工作單 20261031 W3；只對照、不回饋基準；B 情境「η＝2026H1」只作情境，基準 η 不改；列放在本節，使既有列編號不變）")
    # (a) 算力與基礎設施費用 vs 模型算力成本
    K.add("s1a_h1", "(a) S-1：2026 上半年算力與基礎設施費用（損益費用）", "$B", only26(f"={S('ci_s1f_2026h1')}"), "SRC_ANT_426（Verified；期中未經查核）")
    K.add("s1a_q2rr", "(a) S-1：2026 Q2 算力與基礎設施費用 × 每年季數（年化）", "$B", only26(f"={S('ci_s1f_2026q2')}*{qpy}"), "SRC_ANT_432（二手轉載）")
    K.add("s1a_pay", "(a) 模型：2026 合約實付 × 上半年月數 ÷ 12", "$B", only26(f"=«pay»*{h2}/{months}"),
          "口徑差異：模型為現金（年內爬坡，上半年應低於全年一半）；S-1 為費用（預付雲端款先列資產再攤銷、融資租賃以折舊與利息入帳）")
    K.add("s1a_cc", "(a) 模型：2026 算力成本（含自建資本支出）× 上半年月數 ÷ 12", "$B", only26(f"=«cc»*{h2}/{months}"), "自建資本支出是現金資本支出，費用口徑只認列折舊")
    K.add("s1a_gpay", "(a) 模型合約實付（上半年等值）÷ S-1 上半年費用 − 1", "比例", only26("=(«s1a_pay»-«s1a_h1»)/«s1a_h1»"), "", name="COST_S1fGapPayH1", name_cols="E")
    K.add("s1a_gcc", "(a) 模型算力成本（含自建；上半年等值）÷ S-1 上半年費用 − 1", "比例", only26("=(«s1a_cc»-«s1a_h1»)/«s1a_h1»"), "", name="COST_S1fGapCcH1", name_cols="E")
    K.add("s1a_gfy", "(a) 模型 2026 全年合約實付 ÷ S-1 Q2 年化 − 1", "比例", only26("=(«pay»-«s1a_q2rr»)/«s1a_q2rr»"), "Q2 年化為下半年的粗略下限（費用逐季上升）")
    K.add("s1a_capex", "(a) S-1：2026 上半年資本支出 × 12 ÷ 上半年月數（年化）", "$B", only26(f"={S('capex_s1f_2026h1')}*{months}/{h2}"),
          "SRC_ANT_443（二手轉載）；對照模型 Fluidstack 自建資本支出（D11：$50B 三年均攤）")
    K.add("s1a_own", "(a) 模型：2026 自建資本支出 ÷ S-1 年化資本支出", "倍", only26("=«own»/«s1a_capex»"), "遠大於 1＝模型均攤的自建支出在 S-1 上半年現金流中看不到（可能由 Fluidstack 出資、以租賃承擔，或時點較晚）")
    K.add("s1a_draft", "(a) 草擬版轉述『近 9』− S-1 上半年費用", "$B", only26(f"={S('cmp_s1_spend_2026_h1')}-«s1a_h1»"), "SRC_ANT_394 − SRC_ANT_426；U4（Cost 第十八節）仍用草擬版，不改")
    # (b) 營收成本 vs 模型 2025 推論算力
    K.add("s1b_cor25", "(b) S-1：2025 年營收成本（Note 8，功能別）", "$B", only25(f"={S('cor_s1f_fy2025')}"), "SRC_ANT_422（二手轉載）")
    K.add("s1b_inf25", "(b) 模型：2025 推論算力成本（＝說明書算力 7.33 − 訓練 4.1）", "$B", only25("=«c_inf»"), "COST_InfCompute（D8 η 校準用）")
    K.add("s1b_g25", "(b) S-1 營收成本 ÷ 模型 2025 推論算力 − 1", "比例", only25("=(«s1b_cor25»-«s1b_inf25»)/«s1b_inf25»"),
          "口徑差異：營收成本＝服務客戶的推論算力＋其他直接成本（可能含部分人事、支付手續費）；模型推論＝非訓練算力（含免費用戶、內部用量）", name="COST_S1fGapCOR2025", name_cols="D")
    K.add("s1b_sh25", "(b) 營收成本占算力與基礎設施費用（S-1；2025、2026 上半年）", "比例",
          [f"={S('cor_s1f_fy2025')}/{S('ci_s1f_fy2025')}", f"={S('cor_s1f_2026h1')}/{S('ci_s1f_2026h1')}"] + [None] * 4, "S-1 營收成本可能含非算力項目，比例為上緣")
    K.add("s1b_shm", "(b) 模型：推論算力占算力成本（2025、2026）", "比例", [f"=«c_inf»/«cc»", f"=«c_inf»/«cc»"] + [None] * 4, "COST_InfCompute ÷ COST_Compute")
    # (c) η 診斷
    K.add("s1c_cor", "(c) S-1：2026 上半年營收成本 × 12 ÷ 上半年月數（年化）", "$B/年", only26(f"={S('cor_s1f_2026h1')}*{months}/{h2}"), "SRC_ANT_423（二手轉載）")
    K.add("s1c_lp", "(c) 既有雲端加權合約價（與 2025 η 同一值）", "$B/GW/年", only26("=«C:lprice@D»"), "Compute 第三節（D9）")
    K.add("s1c_spgw", "(c) 2026 上半年推論 GW（支出換算）＝年化營收成本 ÷ 合約價", "GW", only26("=«s1c_cor»/«s1c_lp»"), "與 Compute CMP_SpendGW2025 同法")
    K.add("s1c_sh", "(c) 上半年占 2026 營收（S-1 上半年 ÷ 模型 2026 總額）", "比例", only26(f"={S('rev_s1f_2026h1')}/«R:gross_c»"),
          "預設：年內 token 與營收同比例（同一年內牌價與組合不變）；用來把模型全年 token 換算 GW 折成上半年平均")
    K.add("s1c_tok", "(c) 2026 上半年推論 GW（token 換算）＝全年 token 換算 GW × 上半年占比 × 12 ÷ 上半年月數", "GW",
          only26(f"=«C:gw_tok»*«s1c_sh»*{months}/{h2}"), "Compute CMP_TokGW（2026）")
    K.add("s1c_eta", "(c) η_2026H1＝token 換算 GW ÷ 支出換算 GW（上半年）", "倍", only26("=«s1c_tok»/«s1c_spgw»"), "診斷；基準 η 仍＝2025 校準值（CMP_Eta）",
          name="COST_Eta2026H1", name_cols="E")
    K.add("s1c_eta_alt", "(c) 替代：全年 token 換算 GW ÷ 上半年支出換算 GW（上緣）", "倍", only26("=«C:gw_tok»/«s1c_spgw»"), "不折上半年；年內 token 成長使本值偏高")
    K.add("s1c_eta25", "(c) 對照：η（2025）改以 S-1 2025 營收成本換算＝2025 token 換算 GW ÷（營收成本 ÷ 合約價）", "倍",
          only25(f"=«C:gw_tok»/({S('cor_s1f_fy2025')}/«C:lprice»)"), "基準 η（2025）以 說明書算力 − 訓練（3.23）校準（D8）；本列只對照", name="COST_Eta2025COR", name_cols="D")
    K.add("s1c_ratio", "(c) η_2026H1 ÷ 基準 η（2025）", "倍", only26("=«s1c_eta»/«C:eta25@D»"), ">1＝2026 上半年每 GW 實際兌現的 token 多於 2025（或營收成本未含全部推論算力）")
    # (c) B 情境：η＝2026H1 值（2025 沿用校準值）
    K.add("s1h_eta", "(c) B 情境 η 路徑：2025＝校準值；2026 起＝η_2026H1", "倍", lambda i, c: "=«C:eta25@$»" if i == 0 else "=«s1c_eta@E»", "只作情境（工作單預設）")
    K.add("s1h_eff", "(c) B 情境 有效推論 GW＝token 換算 GW ÷ η", "GW", lambda i, c: "=«C:gw_tok»/«s1h_eta»", "")
    K.add("s1h_cap", "(c) B 情境 容量上限係數", "倍", lambda i, c: f"=IF({c}$6<={capfix},«C:vr_unit@$»,MIN(«C:avail»,«s1h_eff»)/«s1h_eff»)",
          "同 Compute 第十四節 B（供給不變；2025、2026 固定 1）", name="CMP_CapFactor_EtaH1")
    K.add("s1h_rd", "(c) B 情境 研發 GW＝供給 ×（1 − 閒置）− 推論（截頂後）", "GW",
          lambda i, c: "=«C:rd»" if i == 0 else f"=«C:sup»*(1-{idle})-IF({c}$6<={capfix},«s1h_eff»,MIN(«C:avail»,«s1h_eff»))", "2025＝基準（訓練支出換算）")
    K.add("s1h_rev", "(c) B 情境 每 VR 等值 GW：營收淨額", "$B/GW/年", lambda i, c: "=«R:net»*«s1h_cap»/«p_den»", "", name="COST_PropRev_VR_EtaH1")
    K.add("s1h_gap", "(c) B 情境 每 VR 等值 GW：差額（含股權報酬）", "$B/GW/年", lambda i, c: "=«s1h_rev»-«p_full_vr»", "全成本不變（合約決定）", name="COST_PropGap_VR_EtaH1")
    K.add("s1h_cov", "(c) B 情境 覆蓋率＝營收淨額 ÷ 全成本（含股權報酬）", "倍", lambda i, c: "=«R:net»*«s1h_cap»/«full»", "", name="COST_Coverage_EtaH1")
    K.add("s1h_dgap", "(c) 差：B 情境 − 基準（每 VR 等值 GW 差額）", "$B/GW/年", lambda i, c: "=«s1h_gap»-«p_gap_vr»", "")
    K.add("s1h_dcov", "(c) 差：B 情境 − 基準（覆蓋率）", "倍", lambda i, c: "=«s1h_cov»-«p_cov»", "")
    # (d) 承諾三段 vs 模型逐家合約總額
    K.add("s1d_l", "(d) S-1：截至 6-30 尚未起租的租賃（未折現）", "$B", only25(f"={S('cmt_s1f_leases_nc_0630')}"), "SRC_ANT_445（二手轉載）")
    K.add("s1d_p", "(d) S-1：截至 6-30 固定或最低採購承諾", "$B", only25(f"={S('cmt_s1f_purchase_0630')}"), "SRC_ANT_446")
    K.add("s1d_n", "(d) S-1：6-30 後新簽與修訂合約的未來最低付款（不含短期承諾）", "$B", only25(f"={S('cmt_s1f_post0630')}"), "SRC_ANT_447；短期承諾 S-1 未單獨揭露")
    K.add("s1d_sum", "(d) S-1 三段合計（不含已起租租賃與短期承諾；S-1 總額留白）", "$B", only25("=«s1d_l@$»+«s1d_p@$»+«s1d_n@$»"), "")
    K.add("s1d_all", "(d) 模型：全部新合約加總（Cost 第十一節）", "$B", only25("=«tot_all@$»"), "COST_CommitAll；不含既有雲端、Fluidstack 自建、機房租約")
    K.add("s1d_gall", "(d) 模型全部新合約 ÷ S-1 三段合計 − 1", "比例", only25("=(«s1d_all@$»-«s1d_sum@$»)/«s1d_sum@$»"), "", name="COST_S1fCommitGap", name_cols="D")
    K.add("s1d_named", "(d) S-1：七家具名合約合計（Google、Amazon、Microsoft、SpaceXAI、Broadcom、Nscale、Lambda）", "$B",
          only25("=" + "+".join(f"{S(k)}{suf}" for _c, _z, k, suf in NAMED7)), "下限、上限與約值混加（S-1 未給合計）")
    K.add("s1d_mnamed", "(d) 模型：同七家合約總額", "$B", only25("=" + "+".join(f"«tot_{c}@$»" for c, *_r in NAMED7)), "")
    K.add("s1d_gnamed", "(d) 模型七家 ÷ S-1 七家 − 1", "比例", only25("=(«s1d_mnamed@$»-«s1d_named@$»)/«s1d_named@$»"), "", name="COST_S1fNamedGap", name_cols="D")
    for c, zh, k, suf in S1F_PARTNERS:
        K.add(f"s1d_pg_{c}", f"(d) 模型 ÷ 正式 S-1 − 1：{zh}（模型＝第十一節合約總額；同第十二節口徑）", "比例", only25(f"=(«tot_{c}@$»-{S(k)}{suf})/({S(k)}{suf})"), "")
    K.add("s1d_pgmax", "(d) 逐家 |差距| 最大值（正式 S-1）", "比例", only25("=MAX(" + ",".join(f"ABS(«s1d_pg_{c}@$»)" for c, *_r in S1F_PARTNERS) + ")"),
          "只對照（INFO）；第十二節草擬版 WARN（C89）不變", name="COST_PartnerGapMaxS1f", name_cols="D")
    K.resolve()
    ctx.P["a6"] = dict(rows=set(range(first, K.r)), keys=[k for k, r in K.rows.items() if r >= first])


def _iso_refs(ctx):
    """第二十四節的列是否被本節以外的公式引用（Checks 除外）：儲存格位址或本節具名範圍。"""
    wb, rows = ctx.wb, ctx.P["a6"]["rows"]
    names = {n for n, v in wb.defined_names.items() if v.attr_text.startswith("Cost!") and
             int(re.search(r"\$(\d+)", v.attr_text).group(1)) in rows}
    pat_local = re.compile(r"(?<![A-Za-z!_$])\$?[D-I]\$?(\d+)\b")
    pat_cost = re.compile(r"Cost!\$?[A-Z]{1,3}\$?(\d+)")
    hits = []
    for ws in wb.worksheets:
        if ws.title == "Checks":
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                if ws.title == "Cost" and c.row in rows:
                    continue
                refs = [int(x) for x in pat_cost.findall(v)]
                if ws.title == "Cost":
                    refs += [int(x) for x in pat_local.findall(re.sub(r"[A-Za-z_]+![$A-Z0-9:]+", "", v))]
                if any(r in rows for r in refs) or any(re.search(rf"\b{n}\b", v) for n in names):
                    hits.append(f"{ws.title}!{c.coordinate}")
    return hits


def checks(ctx):
    D = ctx.D
    S, I = D.S, D.I
    K = Sheet.all["K"]
    kr = K.rows
    d = lambda key, col="D": f"Cost!${col}${kr[key]}"  # noqa: E731
    mx = lambda key: f"=MAX(MAX(Cost!$D${kr[key]}:$I${kr[key]}),-MIN(Cost!$D${kr[key]}:$I${kr[key]}))"  # noqa: E731
    hits = _iso_refs(ctx)
    ctx.scan["s1f_refs"], ctx.scan["s1f_cells"] = len(hits), hits
    q = lambda *ks: "+".join(S(k) for k in ks)  # noqa: E731
    return [
        ("f1_iso", "v0.1.6：Cost 第二十四節（正式 S-1 對照與 η_2026H1 情境）被其他公式引用的格數（builder 掃描）", len(hits), 0, "eq", "W3 只對照、不回饋基準"),
        ("f1_h1", "v0.1.6 W2：2026 上半年營收 正式 S-1 − 草擬版（$B）", f"={S('rev_s1f_2026h1')}-{S('rev_2026_h1_s1')}", None, "info", "16.286 − 16.3"),
        ("f1_h1pct", "v0.1.6 W2：正式 S-1 ÷ 草擬版 − 1（2026 上半年營收；|差距| > 5% → WARN，工作單停止條件）", f"=({S('rev_s1f_2026h1')}-{S('rev_2026_h1_s1')})/{S('rev_2026_h1_s1')}",
         I("s1f_warn_threshold"), "warn", "超過門檻時先報告、不改基準"),
        ("f1_target", "v0.1.6 W2：2026 營收基準（上半年 正式 S-1＋Q3 實際＋Q4）（$B）", "=INDEX(REV_Target,1,2)", None, "info", "16.286＋16.8＋16.8"),
        ("f1_rq25", "v0.1.6：S-1 2025 四季營收合計 − 2025 年營收（$B）", f"={q('rev_s1f_2025q1', 'rev_s1f_2025q2', 'rev_s1f_2025q3', 'rev_s1f_2025q4')}-{S('rev_s1f_fy2025')}", 0, "tol",
         "季度（二手轉載）與年度（Verified）互相核對"),
        ("f1_rq26", "v0.1.6：S-1 2026 Q1＋Q2（Kepler）− 上半年營收（$B）", f"={q('rev_s1f_2026q1', 'rev_s1f_2026q2')}-{S('rev_s1f_2026h1')}", 0, "tol", ""),
        ("f1_rq2b", "v0.1.6：2026 Q2 營收 thevccorner − Kepler（$B；進位差）", f"={S('rev_s1f_2026q2_b')}-{S('rev_s1f_2026q2')}", None, "info", "兩值並列登錄"),
        ("f1_oq25", "v0.1.6：S-1 2025 四季營業損益合計 − 2025 年（$B）", f"={q('oi_s1f_2025q1', 'oi_s1f_2025q2', 'oi_s1f_2025q3', 'oi_s1f_2025q4')}-{S('oi_s1f_fy2025')}", 0, "tol", ""),
        ("f1_oq26", "v0.1.6：S-1 2026 Q1＋Q2 營業損益 − 上半年（$B）", f"={q('oi_s1f_2026q1', 'oi_s1f_2026q2')}-{S('oi_s1f_2026h1')}", 0, "tol", ""),
        ("f1_cq", "v0.1.6：S-1 算力與基礎設施費用 2026 Q1＋Q2 − 上半年（$B；季度為兩位小數）", f"={q('ci_s1f_2026q1', 'ci_s1f_2026q2')}-{S('ci_s1f_2026h1')}", None, "info", "進位差"),
        ("f1_rev25", "v0.1.6：基準 2025 校準營收（SRC_ANT_001）− 正式 S-1 2025 營收（$B）", f"={S('rev_2025_recognized')}-{S('rev_s1f_fy2025')}", None, "info",
         "工作單：基準只改 W2，2025 校準不改（差 0.3%）"),
        ("f1_a_pay", "v0.1.6 W3(a)：模型 2026 合約實付（上半年等值）÷ S-1 上半年算力與基礎設施費用 − 1", "=COST_S1fGapPayH1", None, "info", "現金 vs 費用"),
        ("f1_a_cc", "v0.1.6 W3(a)：模型 2026 算力成本（含自建；上半年等值）÷ S-1 − 1", "=COST_S1fGapCcH1", None, "info", ""),
        ("f1_a_own", "v0.1.6 W3(a)：模型 2026 自建資本支出 ÷ S-1 年化資本支出（倍）", f"={d('s1a_own', 'E')}", None, "info", "D11 Fluidstack 均攤"),
        ("f1_b", "v0.1.6 W3(b)：S-1 2025 營收成本 ÷ 模型 2025 推論算力 − 1", "=COST_S1fGapCOR2025", None, "info", ""),
        ("f1_c_eta", "v0.1.6 W3(c)：η_2026H1", "=COST_Eta2026H1", None, "info", "基準 η 不改"),
        ("f1_c_eta25", "v0.1.6 W3(c)：η（2025）改用 S-1 營收成本換算（對照）", "=COST_Eta2025COR", None, "info", "基準 η（2025）見 a3_eta25"),
        ("f1_h_gap30", "v0.1.6 B 情境（η＝2026H1）：2030 每 VR 等值 GW 差額（$B/GW/年）", "=INDEX(COST_PropGap_VR_EtaH1,1,6)", None, "info", "基準 a3_gap30"),
        ("f1_h_cov30", "v0.1.6 B 情境（η＝2026H1）：2030 覆蓋率", "=INDEX(COST_Coverage_EtaH1,1,6)", None, "info", ""),
        ("f1_h_dmax", "v0.1.6 B 情境 − 基準：每 VR 等值 GW 差額的最大絕對差（2025–2030）", mx("s1h_dgap"), None, "info", ""),
        ("f1_d_gap", "v0.1.6 W3(d)：模型全部新合約 ÷ S-1 承諾三段合計 − 1", "=COST_S1fCommitGap", None, "info", ""),
        ("f1_d_named", "v0.1.6 W3(d)：模型七家 ÷ S-1 七家具名合約 − 1", "=COST_S1fNamedGap", None, "info", ""),
        ("f1_d_pgmax", "v0.1.6 W3(d)：逐家 |模型 ÷ 正式 S-1 − 1| 最大值", "=COST_PartnerGapMaxS1f", None, "info", "只對照；第十二節草擬版 WARN（C89）不變"),
        ("f1_cash", "v0.1.6：S-1 2026-06-30 現金與短期投資（$B；U8 缺口補上，只列示）", f"={S('cash_s1f_2026_06_30')}", None, "info", "Funding 不改（工作單）"),
        ("f1_oai_rev26", "v0.1.6 OpenAI 並排：2026 每 VR 等值 GW 營收淨額 Anthropic ÷ OpenAI", "=INDEX(COST_PropRev_VR,1,2)/INDEX(OAI_COST_PropRev_VR,1,2)", None, "info", "OAI_Link（D20）"),
    ]
