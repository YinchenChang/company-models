"""Anthropic v0.1.3 B 段：Bridge 頁（需求 → token → GW → 成本逐層對照）與 EtaDiag 頁（η 診斷拆解）。

兩頁都是「唯讀」：只引用既有頁與 TK_Link，不被任何計算頁引用（Checks 除外；builder 掃描＋tests 每次重驗）。
- Bridge：由上而下逐層列出每一步的公式來源（編號）、單位、2025–2030 值、2030 敏感度包絡（低／高）、標記；最後一節與 Tokenomics 同口徑量對照。
  不以「與 Tokenomics 逐項閉合」為驗收條件（TK token 帳 6.26 倍未閉合是 Tokenomics 端的待辦）。
- EtaDiag：η（2025）＝計價口徑因子 × 利用率因子 × 工作負載組合因子 × 未解釋殘差。只作診斷，不改 η 的值；
  並列 TK 的 L1_Gap*（OpenAI 口徑，Analogy；方向相反：TK 為支出路線 ÷ token 路線）。
新列／新頁皆附加在尾端，既有列編號不變。數值一律引用既有具名範圍或儲存格（公式不含常數；常數經 Inputs 定義常數）。
"""
from __future__ import annotations

from openpyxl.styles import Alignment

from a2 import FMT, TOK, YC, YEARS, Sheet
from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, header, nm, put, title

for _u, _f in (("十億任務/年", "#,##0.00"), ("M tok/年", "#,##0"), ("M tok/GW/年", "#,##0"), ("GW", "#,##0.000"), ("$B/GW/年", "#,##0.000"),
               ("倍", "0.000"), ("$B", "#,##0.000"), ("K tok/任務", "#,##0.0"), ("任務/日", "0.00"), ("T", "#,##0"), ("M", "#,##0.000")):
    FMT.setdefault(_u, _f)


def build(ctx):
    D, wb = ctx.D, ctx.wb
    I = D.I
    years = [I(f"year_{y}") for y in YEARS]
    Cs, Ks, Rs, Ds = Sheet.all["C"], Sheet.all["K"], Sheet.all["R"], Sheet.all["D"]
    cr, kr = Cs.rows, Ks.rows
    days, m2b, tpk = I("days_per_year"), I("m_to_b"), I("tok_per_k")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    only26 = lambda f: [None, f] + [None] * 4  # noqa: E731

    B = Sheet(wb, "Bridge", "B", "B", years,
              "Bridge — 需求 → token → GW → 成本：逐層對照（唯讀；不被任何計算頁引用）",
              "由上而下：用戶與席位 → 任務 → 每任務 token（TK IF_TaskTok*）→ 付費／免費 × 層級 token → 有效單價 → 營收；token → 每 GW 理論產能（TK IF_TokGW_* × 世代與加速器組合）→ ÷ η → 有效推論 GW → ＋研發 GW 殘差 → 供給 GW → 合約付款與自建 → 成本 → 每 VR 等值 GW 差額。"
              "D–I 欄＝2025–2030 的同一公式（綠字＝引用他頁）；J 欄標記；K 欄＝公式來源（頁＋編號）；L、M 欄＝低／高（依據見 N 欄）。",
              "本頁只列示與核對，不產生任何數值給其他頁。最後一節與 Tokenomics 對照：差額若大，部分來自 TK 端 token 帳未閉合（L1_TokLedger_2026Q2），屬 Tokenomics 範圍，不作為本模型的通過條件。")
    ws = B.ws
    header(ws, 5, ["低", "高", "低／高依據"], start_col=12)
    put(ws, "J5", "標記", F_BOLD)

    def src_note(fs):
        f = next((x for x in ([fs(i, YC[i]) for i in range(6)] if callable(fs) else fs) if x), "")
        parts = []
        for m in TOK.finditer(f):
            tgt = Sheet.all[m.group(1)] if m.group(1) else B
            r = tgt.rows[m.group(2)]
            p = f"{tgt.name} {tgt.ws[f'A{r}'].value}"
            if p not in parts:
                parts.append(p)
        if "TK_" in f:
            parts.append("TK_Link")
        if "Inputs!" in f or "INP_" in f:
            parts.append("Inputs")
        return "、".join(parts)

    def add(key, label, unit, fs, tag, extra="", lohi=None):
        """lohi：(低公式, 高公式, 依據文字)；公式用明確儲存格位址（不含記號）。"""
        note = src_note(fs) + (f"｜{extra}" if extra else "")
        r = B.add(key, label, unit, fs, note)
        put(ws, f"J{r}", tag, F_NOTE)
        if lohi:
            lo, hi, why = lohi
            put(ws, f"L{r}", lo, fmt=FMT.get(unit))
            put(ws, f"M{r}", hi, fmt=FMT.get(unit))
            put(ws, f"N{r}", why, F_NOTE)
        return r

    def env(sheet, keys, col="I"):
        """2030 敏感度包絡：Compute／Cost 第十四節等既有情境列（欄 I＝2030）的最小與最大。"""
        rows = Cs.rows if sheet == "Compute" else Ks.rows
        cells = ",".join(f"{sheet}!${col}${rows[k]}" for k in keys)
        return f"=MIN({cells})", f"=MAX({cells})"

    def lohi(sheet, keys, why, col="I"):
        lo, hi = env(sheet, keys, col)
        return lo, hi, why

    TAG_D = "Derived"
    # ═════════════════════ 一、用戶與席位 ═════════════════════
    B.section("一、用戶與席位（M；FY 平均）")
    add("u_cons", "個人付費人數（Pro＋Max）", "M", lambda i, c: "=«D:u_cons»", "Derived（2025 由訂閱營收校準；2026 起 Assumed 成長）")
    add("s_total", "企業席位數（Team＋Enterprise）", "M", lambda i, c: "=«D:s_total»", "Derived（同上）")
    add("u_free", "Free 人數", "M", lambda i, c: "=«D:u_free»", "Derived（2025 月活 − 個人付費；2026 起 Assumed 成長）")
    # ═════════════════════ 二、任務 ═════════════════════
    B.section("二、任務（每日任務數 → 年任務數）")
    for k, zh in (("free", "Free"), ("pro", "Pro"), ("max5", "Max 5x"), ("max20", "Max 20x"), ("team_std", "Team 標準"), ("team_prem", "Team Premium"), ("ent", "Enterprise")):
        add(f"tasks_{k}", f"每人（席）每日任務數：{zh}", "任務/日", lambda i, c, k=k: f"=«D:tasks_{k}»", "Analogy／Assumed（OpenAI v0.6 類比；區間 ×0.5–×2）")
    users = ["u_free", "u_pro", "u_max5", "u_max20", "s_team_std", "s_team_prem", "s_ent"]
    plans = ["free", "pro", "max5", "max20", "team_std", "team_prem", "ent"]
    sub_expr = "+".join(f"«D:{u}»*«D:tasks_{p}»" for u, p in zip(users, plans))
    add("tasks_sub", "訂閱、席位與 Free 年任務數", "十億任務/年", lambda i, c: f"=({sub_expr})*{days}/{m2b}", "Derived", "Σ 人數 × 每日任務數 × 天數 ÷ 單位換算")
    add("tasks_api", "API 年任務數", "十億任務/年", lambda i, c: "=«D:api_n»", "Derived（2025 由 API 營收倒推；2027 起任務成長 × 價格反應）")
    add("tasks_all", "年任務數合計", "十億任務/年", lambda i, c: "=«tasks_sub»+«tasks_api»", "Derived")
    # ═════════════════════ 三、每任務 token ═════════════════════
    B.section("三、每任務 token（K token；2025 取 Tokenomics 任務模板 × 倍數；2026 起 × 年增）")
    tkfmt = ("=(SUMIFS(TK_IF_TaskTokFresh,TK_IF_HdrTask,{k})+SUMIFS(TK_IF_TaskTokCached,TK_IF_HdrTask,{k})"
             "+SUMIFS(TK_IF_TaskTokDec,TK_IF_HdrTask,{k}))/" + tpk)
    cats = (("chat", "對話", "tk_task_chat", "k_ratio_chat"), ("code", "程式代理", "tk_task_code", "k_ratio_code"), ("other", "其他代理", "tk_task_other", "k_ratio_other"))
    for ck, zh, tkk, rk in cats:
        add(f"tk_{ck}", f"TK 任務模板每次嘗試 token：{zh}（新鮮＋快取＋decode；IF_TaskTok*）", "K tok/任務", only25(tkfmt.format(k=I(tkk))),
            "Analogy（Tokenomics 任務模板；非 Anthropic 實測）", "Inputs 對應鍵：" + I(tkk))
    for ck, zh, tkk, rk in cats:
        add(f"ratio_{ck}", f"倍數（Claude 每任務 ÷ TK 任務模板）：{zh}", "倍", only25(f"={I(rk)}"), "Assumed（區間見 L、M）", "",
            (f"=Inputs!$F${D.inp_row(rk)}", f"=Inputs!$G${D.inp_row(rk)}", "2025；Inputs 低／高欄"))
    for ck, zh, tkk, rk in cats:
        r_tk, r_ra = B.rows[f"tk_{ck}"], B.rows[f"ratio_{ck}"]
        add(f"k_{ck}", f"模型每任務 token：{zh}", "K tok/任務", lambda i, c, ck=ck: f"=«D:k_{ck}»", "Derived（＝TK 模板 × 倍數；2026 起 × 年增）", "",
            (f"=$D${r_tk}*$L${r_ra}", f"=$D${r_tk}*$M${r_ra}", "2025：TK 模板 × 倍數低／高"))
    add("k_api", "API 每任務 token", "K tok/任務", lambda i, c: "=«D:api_k»", "Assumed（2025 起點；2027 起 × 年增 × 價格反應）")
    # ═════════════════════ 四、token ═════════════════════
    B.section("四、付費／免費 × 層級 token（T）")
    for pf, zh in (("paid", "付費"), ("free", "免費")):
        for tk_, zt in (("top", "頂層（Opus→Astra）"), ("mid", "中層（Sonnet→Sol）"), ("low", "低層（Haiku→Luna）")):
            add(f"{pf}_{tk_}", f"{zh} token：{zt}", "T", lambda i, c, pf=pf, tk_=tk_: f"=«D:{pf}_{tk_}»", "Derived（層級占比 Assumed）")
    add("paid_all", "付費 token 合計", "T", lambda i, c: "=«D:paid_all»", "Derived")
    add("free_all", "免費 token 合計", "T", lambda i, c: "=«D:free_all»", "Derived")
    add("tok_all", "token 合計（付費＋免費）", "T", lambda i, c: "=«D:tok_all»", "Derived")
    add("api_share", "API 占 token 合計", "比例", lambda i, c: "=«D:api_tok_share»", "Derived")
    # ═════════════════════ 五、單價與營收 ═════════════════════
    B.section("五、有效單價與營收（單價 $/M token；營收 $B）")
    for tk_, zt in (("top", "頂層（Opus→Astra）"), ("mid", "中層（Sonnet→Sol）"), ("low", "低層（Haiku→Luna）")):
        add(f"ep_{tk_}", f"API 有效單價：{zt}", "$/M", lambda i, c, tk_=tk_: f"=«R:ep_{tk_}»", "Derived（牌價 Verified × 輸出占比、快取、批次、折扣 Assumed）")
    add("mixprice", "API 組合有效單價", "$/M", lambda i, c: "=«R:mixprice»", "Derived")
    add("api_recalc", "API 營收（重算＝API token × 組合有效單價 ÷ 單位換算）", "$B", lambda i, c: f"=«D:api_tok»*«mixprice»/{m2b}", "Derived", "與 Revenue 的 API 營收核對（下方差額列）")
    add("sub_rev", "訂閱營收（個人＋席位）", "$B", lambda i, c: "=«R:sub»", "Derived（2025 校準至 SRC 訂閱營收；Interested-party）")
    add("api_rev", "API 營收（Revenue）", "$B", lambda i, c: "=«R:api»", "Derived（2025 校準至 SRC 用量計價營收）")
    add("other_rev", "其他營收（含廣告＝0）", "$B", lambda i, c: "=«R:other»+«R:ads»", "Derived")
    add("gross", "營收總額（報導口徑）", "$B", lambda i, c: "=«R:gross»", "Derived（2025 校準至 SRC 認列營收）")
    add("ck_rev", "核對：（訂閱＋重算 API＋其他）− 營收總額（應為 0）", "$B", lambda i, c: "=«sub_rev»+«api_recalc»+«other_rev»-«gross»", "核對")
    add("pshare", "雲端平台抽成", "$B", lambda i, c: "=«R:pshare»", "Interested-party（抽成率，Reuters／報導）")
    add("net", "營收淨額（截頂後；命題與現金流用）", "$B", lambda i, c: "=«R:net_c»", "Derived")
    # ═════════════════════ 六、token → GW ═════════════════════
    B.section("六、token → 每 GW 理論產能 → 推論 GW → 供給 GW（GW；產能 M token／GW／年）")
    for tk_, zt in (("top", "頂層（Astra）"), ("mid", "中層（Sol）"), ("low", "低層（Luna）")):
        add(f"tokgw_{tk_}", f"TK 每 GW 總產出（VR200，100% 利用率）：{zt}（IF_TokGW_*）", "M tok/GW/年", only25(f"=«C:tok_vr200_{tk_}@$»"),
            "Analogy（Tokenomics 自下而上；非 Anthropic 實測）")
    add("util", "TK 基準利用率（IF_Util）", "比例", only25("=TK_IF_Util"), "Analogy（TK；實際利用率 TK_IF_UtilActual 見 EtaDiag）")
    for tk_, zt in (("top", "頂層"), ("mid", "中層"), ("low", "低層")):
        add(f"mpaid_{tk_}", f"付費層級占比：{zt}", "比例", lambda i, c, tk_=tk_: f"=«C:m_paid_{tk_}»", "Derived（Assumed 層級組合）")
    fam = (("hopper", "NVIDIA Hopper"), ("gb200", "NVIDIA GB200"), ("gb300", "NVIDIA GB300"), ("vr200", "NVIDIA VR200"), ("tpu6", "Google TPU v6e"),
           ("tpu7", "Google TPU v7／次世代"), ("trn2", "AWS Trainium2"), ("trn3", "AWS Trainium3"), ("mi455", "AMD MI455X"))
    for fk, zf in fam:
        add(f"share_{fk}", f"加速器組合占比：{zf}", "比例", lambda i, c, fk=fk: f"=«C:s_{fk}»", "Derived（由逐合約供給 GW 加權；產出比 Analogy）")
    add("cap_paid", "組合後每 GW 年產能（付費）", "M tok/GW/年", lambda i, c: "=«C:cap_paid»", "Analogy（TK IF_TokGW × 組合；區間見 L、M＝產出比低／高）",
        "", lohi("Compute", ["C_lo_cap_paid", "C_hi_cap_paid"], "2030；Compute 第十四節 C：TPU／Trainium／AMD 產出比低／高"))
    add("cap_free", "組合後每 GW 年產能（免費）", "M tok/GW/年", lambda i, c: "=«C:cap_free»", "Analogy（同上）",
        "", lohi("Compute", ["C_lo_cap_free", "C_hi_cap_free"], "2030；同上"))
    add("tok_paid", "付費 token", "M tok/年", lambda i, c: "=«C:tok_paid»", "Derived")
    add("tok_free", "免費 token", "M tok/年", lambda i, c: "=«C:tok_free»", "Derived")
    add("gw_tpaid", "推論 GW（token 換算）：付費", "GW", lambda i, c: "=«C:gw_tpaid»", "Derived")
    add("gw_tfree", "推論 GW（token 換算）：免費", "GW", lambda i, c: "=«C:gw_tfree»", "Derived")
    add("gw_tok", "推論 GW（token 換算）：合計", "GW", lambda i, c: "=«C:gw_tok»", "Derived", "",
        lohi("Compute", ["C_lo_tok", "C_hi_tok"], "2030；產出比低／高"))
    add("eta", "η（有效產出係數；2025 由推論支出校準，2026 起沿用）", "倍", lambda i, c: "=«C:eta»", "Derived（2025）／Assumed（2026 起沿用）", "拆解見 EtaDiag 頁",
        lohi("Compute", ["B_keep_eta", "B_rise_eta"], "2030；η 沿用（基準）／線性回升至目標（Compute 第十四節 B）"))
    add("eff", "有效推論 GW＝token 換算 GW ÷ η（未截頂）", "GW", lambda i, c: "=«C:eff»", "Derived", "",
        lohi("Compute", ["A_lo_eff", "A_hi_eff", "B_keep_eff", "B_rise_eff", "C_lo_eff", "C_hi_eff"], "2030；三組敏感度（合約價 8／20、η 沿用／回升、產出比低／高）的最小／最大"))
    add("gw_chk", "核對：token 換算 GW ÷ η − 有效推論 GW（應為 0）", "GW", lambda i, c: "=«gw_tok»/«eta»-«eff»", "核對")
    add("inf", "推論 GW（截頂後）", "GW", lambda i, c: "=«C:infcap»", "Derived")
    add("idle", "閒置 GW", "GW", lambda i, c: "=«C:idle»", "Assumed（閒置保留比例 Inputs 低／高）")
    add("rd", "研發 GW（殘差）＝供給 ×（1 − 閒置）− 推論", "GW", lambda i, c: "=«C:rd»", "Derived（殘差；非獨立估計）")
    add("sup_con", "供給 GW：合約（既有雲端＋新合約）", "GW", lambda i, c: "=«C:sup_con»", "Interested-party（公開合約揭露）＋Derived（金額型＝付款 ÷ 合約價）", "",
        lohi("Compute", ["A_lo_sup", "A_hi_sup"], "2030；合約價 8／20 $B/GW/年（含自建）"))
    add("sup_own", "供給 GW：自建（Fluidstack）", "GW", lambda i, c: "=«C:sup_own»", "Derived（$50B 自有資本支出 ÷ TK 每 GW 資本支出）")
    add("sup", "供給 GW 合計", "GW", lambda i, c: "=«C:sup»", "Derived", "",
        lohi("Compute", ["A_lo_sup", "A_hi_sup"], "2030；合約價 8／20"))
    add("vrf", "機隊 VR 等值係數", "倍", lambda i, c: "=«C:vrf»", "Analogy（各族 Sol 產能比；NonNV 產出比）", "",
        lohi("Compute", ["C_lo_vrf", "C_hi_vrf"], "2030；產出比低／高"))
    add("vr_sup", "供給 VR 等值 GW（命題分母）", "GW", lambda i, c: "=«C:vr_sup»", "Derived", "",
        lohi("Compute", ["A_lo_vr", "A_hi_vr"], "2030；合約價 8／20"))
    # ═════════════════════ 七、合約付款與自建 → 成本 ═════════════════════
    B.section("七、合約付款與自建 → 成本（$B）")
    pays = (("legacy", "既有雲端（2025 前簽）"), ("google", "Google Cloud TPU"), ("broadcom", "Google／Broadcom 次世代 TPU"), ("aws", "Amazon 擴大協議"),
            ("azure", "Microsoft Azure＋NVIDIA"), ("amd", "AMD MI455X"), ("spacex", "SpaceX／xAI Colossus"), ("nscale", "Nscale Monarch"), ("lambda", "Lambda"),
            ("volta", "Volta Infra"), ("akamai", "Akamai"))
    for pk, zp in pays:
        add(f"pay_{pk}", f"合約實付：{zp}", "$B", lambda i, c, pk=pk: f"=«K:pay_{pk}»", "Interested-party（金額、期間）＋Derived（爬坡）")
    add("pay", "合約實付合計", "$B", lambda i, c: "=«K:pay»", "Derived")
    add("own", "自建資本支出（Fluidstack）", "$B", lambda i, c: "=«K:own»", "Interested-party（$50B 公告）")
    add("cc", "算力成本（現金口徑）", "$B", lambda i, c: "=«K:cc»", "Derived", "",
        lohi("Cost", ["x_cc"], "（僅供參照）") if False else lohi("Compute", ["A_lo_cc", "A_hi_cc"], "2030；合約價 8／20"))
    add("ncx", "非算力成本（不含股權報酬）", "$B", lambda i, c: "=«K:ncx»", "Derived（2025 說明書；2026 起人數 × 每人成本 Assumed）")
    add("sbc", "股權報酬（非現金，單列）", "$B", lambda i, c: "=«K:sbc»", "Assumed（每人 $0.8M）")
    add("full", "全成本（含股權報酬）", "$B", lambda i, c: "=«K:full»", "Derived")
    add("p_rev_vr", "每 VR 等值 GW：營收淨額", "$B/GW/年", lambda i, c: "=«K:p_rev_vr»", "Derived")
    add("p_full_vr", "每 VR 等值 GW：全成本（含股權報酬）", "$B/GW/年", lambda i, c: "=«K:p_full_vr»", "Derived")
    add("p_gap_vr", "每 VR 等值 GW：差額（含股權報酬；命題 1）", "$B/GW/年", lambda i, c: "=«K:p_gap_vr»", "Derived", "",
        lohi("Compute", ["A_lo_gap", "A_hi_gap", "C_lo_gap", "C_hi_gap"], "2030；合約價 8／20 與產出比低／高的最小／最大"))
    add("p_cov", "覆蓋率（營收淨額 ÷ 全成本）", "倍", lambda i, c: "=«K:p_cov»", "Derived")
    # ═════════════════════ 八、與 Tokenomics 對照 ═════════════════════
    B.section("八、與 Tokenomics 對照（2026；TK 的 P × Q、Q 與 token 帳診斷；只對照，不回饋）。TK token 帳未閉合（6.26 倍）是 Tokenomics 端待辦，差額大不代表本模型錯")
    add("tk_pq", "TK：VR200 每 GW 營收 P × Q（L1_RevGW_PQ_VR200；2026 Q2 年化 m）", "$B/GW/年", only26("=«K:pq_base@E»"),
        "Analogy（TK；OpenAI 口徑推導的機隊產出）", "", ("=Cost!$E$" + str(kr["pq_lo"]), "=Cost!$E$" + str(kr["pq_hi"]), "2026；TK 實際利用率 0.30／0.65"))
    add("md_pq", "本模型：營收總額 ÷ 付費推論 VR 等值 GW（與 TK P × Q 最接近的口徑）", "$B/GW/年", only26("=«K:pq_rev_paid@E»"), "Derived")
    add("df_pq", "差額（本模型 − TK）", "$B/GW/年", only26("=«md_pq»-«tk_pq»"), "Derived")
    add("rt_pq", "比值（TK ÷ 本模型）", "倍", only26("=«tk_pq»/«md_pq»"), "Derived", "1＝一致")
    add("tk_q", "TK：VR200 每 GW 付費 token 實際產出 Q（L1_TokGW_PaidPQ_VR200）", "M tok/年", only26("=«K:pq_q_tk@E»"), "Analogy（TK；實際利用率 0.41）", "",
        ("=TK_L1_TokGW_PaidPQ_VR200_Lo", "=TK_L1_TokGW_PaidPQ_VR200_Hi", "2026；TK 實際利用率 0.30／0.65"))
    add("md_q", "本模型：每付費推論 VR 等值 GW 的付費 token", "M tok/年", only26("=«K:pq_q_model@E»"), "Derived（η 校準後的有效 GW）")
    add("df_q", "差額（本模型 − TK）", "M tok/年", only26("=«md_q»-«tk_q»"), "Derived")
    add("rt_q", "比值（本模型 ÷ TK）", "倍", only26("=«md_q»/«tk_q»"), "Derived", "對照 TK 帳未閉合倍數")
    add("tk_led", "TK：token 帳未閉合倍數（L1_TokLedger_2026Q2；TK 模型付費 token ÷ 需求面付費 token）", "倍", only26("=«K:pq_led@E»"), "Analogy（TK 診斷；OpenAI 口徑）",
        "Tokenomics 待辦；本模型不據以調整")
    B.resolve()
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 12
    for c in YC:
        ws.column_dimensions[c].width = 13
    ws.column_dimensions["J"].width = 34
    ws.column_dimensions["K"].width = 42
    ws.column_dimensions["L"].width = 13
    ws.column_dimensions["M"].width = 13
    ws.column_dimensions["N"].width = 56
    ws.freeze_panes = "D7"

    # ═════════════════════ EtaDiag ═════════════════════
    E = wb.create_sheet("EtaDiag")
    title(E, "EtaDiag — η 診斷拆解（唯讀；只作診斷，不改 η 的值）",
          "η（2025）＝計價口徑因子 × 利用率因子 × 工作負載組合因子 × 未解釋殘差。因子皆以 η 的方向定義（η ＝ token 換算 GW ÷ 支出換算 GW）；TK 的 L1_Gap* 為相反方向（支出路線 ÷ token 路線），並列時取倒數對照。",
          "D＝基準、E＝低、F＝高（區間端點；依據見 H 欄）。TK 的 L1_Gap* 為 OpenAI 口徑（Analogy）：本模型的 Anthropic 數字與其只作並排，不得因拆分改變 η。"
          "「輸入輸出組合」（TK L1_GapPrompt、L1_GapISL）在本模型無法分離：每任務 token 已直接取自 TK 任務模板（含輸入、快取、輸出），並入殘差。")
    header(E, 5, ["編號", "項目", "單位", "基準", "低", "高", "標記", "說明／公式來源"])
    st = {"r": 6, "n": 0, "rows": {}}

    def sec(label):
        put(E, f"A{st['r']}", label, F_BOLD, fill=FILL_SEC)
        st["r"] += 1

    def row(key, label, unit, d, e=None, f=None, tag="Derived", note="", name=None):
        st["n"] += 1
        r = st["r"]
        st["rows"][key] = r
        put(E, f"A{r}", f"E{st['n']:02d}", F_CALC)
        put(E, f"B{r}", label, F_CALC)
        put(E, f"C{r}", unit, F_CALC)
        for col, v in zip("DEF", (d, e, f)):
            if v is not None:
                put(E, f"{col}{r}", v, fmt=FMT.get(unit, "#,##0.000"))
        put(E, f"G{r}", tag, F_NOTE)
        put(E, f"H{r}", note, F_NOTE)
        if name:
            nm(wb, name, f"EtaDiag!$D${r}")
        st["r"] += 1
        return r

    R_ = st["rows"]
    d = lambda k, col="D": f"${col}${R_[k]}"  # noqa: E731
    cD = lambda k: f"Compute!$D${cr[k]}"  # noqa: E731
    kD = lambda k: f"Cost!$D${kr[k]}"  # noqa: E731

    sec("一、被拆解的量（2025）")
    row("eta", "η（2025）＝token 換算 GW ÷ 支出換算 GW", "倍", "=CMP_Eta2025", tag="Derived", note="Compute C133（CMP_Eta2025）；本頁不改此值", name="ETA_Eta")
    row("tokgw", "token 換算推論 GW（2025；付費＋免費；以 TK IF_Util 為基準利用率）", "GW", f"={cD('gw_tok')}", tag="Derived", note="Compute C130")
    row("spendgw", "支出換算推論 GW（2025）＝推論支出 ÷ 既有雲端加權合約價", "GW", f"={cD('e_spendgw')}", tag="Derived", note="Compute C132（支出為說明書；Interested-party）")
    row("ck_eta", "核對：token GW ÷ 支出 GW − η（應為 0）", "倍", f"={d('tokgw')}/{d('spendgw')}-{d('eta')}", tag="核對")

    sec("二、計價口徑因子 f_P ＝ 每 GW 合約實付 ÷ TK 持有成本（加權）。η 用合約價把支出換成 GW；若改用持有成本換算，支出 GW 會較大（價 > 持有成本時）")
    row("price", "既有雲端加權每 GW 合約實付（2025）", "$B/GW/年", f"={cD('lprice')}", tag="Interested-party／Derived", note="Compute C58（合約價加權；2025 前簽之 AWS、Google、NVIDIA）")
    row("hold", "TK 持有成本（經濟口徑）加權每 GW", "$B/GW/年", f"=INDEX(COST_SupplierHold,1,1)/INDEX(CMP_SupplyContractGW,1,1)",
        f"=Cost!$D${kr['sh_lo']}/INDEX(CMP_SupplyContractGW,1,1)", f"=Cost!$D${kr['sh_hi']}/INDEX(CMP_SupplyContractGW,1,1)",
        tag="Analogy（TK IF_HoldEcon × 持有比；低／高＝TK 低／高成本情境）", note="Cost K61–K63 ÷ 合約供給 GW")
    row("fp", "f_P：計價口徑因子", "倍", f"={d('price')}/{d('hold')}", f"={d('price')}/{d('hold','F')}", f"={d('price')}/{d('hold','E')}",
        tag="Derived（以 Analogy 持有成本；區間＝TK 成本情境）", note="低＝除以高成本情境；高＝除以低成本情境", name="ETA_fP")

    sec("三、利用率因子 f_U ＝ TK 實際利用率 ÷ TK 基準利用率（token 換算 GW 以基準利用率計；若實際較低，token 換算 GW 偏低，使 η 偏低）")
    row("u0", "TK 基準利用率（IF_Util；本模型每 GW 產能的基準）", "比例", "=TK_IF_Util", tag="Analogy（TK）", note="Compute 第八節")
    row("ua", "TK 實際利用率（IF_UtilActual；TK 對 OpenAI 機隊的估計）", "比例", "=TK_IF_UtilActual", "=TK_IF_UtilActual_Lo", "=TK_IF_UtilActual_Hi",
        tag="Analogy（TK；非 Anthropic 實測）", note="本模型無 Anthropic 自己的利用率資料；以 TK 區間作 Analogy")
    row("fu", "f_U：利用率因子", "倍", f"={d('ua')}/{d('u0')}", f"={d('ua','E')}/{d('u0')}", f"={d('ua','F')}/{d('u0')}", tag="Derived（以 Analogy）",
        note="TK 方向（L1_GapUtil＝IF_Util ÷（實際 × 生產折減））約為本因子倒數再除以生產折減；本列（第二節）不計生產折減，計入者見第七節 f_U×生產折減（v0.1.4）", name="ETA_fU")

    sec("四、工作負載（層級）組合因子 f_W ＝ TK 參考層級組合下 2025 機隊每 GW 付費產能 ÷ 本模型付費層級組合下的同一機隊每 GW 付費產能（兩者同一基準利用率、同一機隊組成）")
    for key, lab, ix in (("mx_low", "TK 參考付費層級組合（B4_MixPaid）：低層（Luna）", 1), ("mx_mid", "TK 參考付費層級組合（B4_MixPaid）：中層（Sol）", 2),
                         ("mx_top", "TK 參考付費層級組合（B4_MixPaid）：頂層（Astra）", 3)):
        row(key, lab, "比例", f"=INDEX(TK_B4_MixPaid,1,{ix})", tag="Analogy（TK；OpenAI 口徑的付費層級組合）", note="Cap_In C31:E31（Luna／Sol／Astra）")
    gens = (("hopper", "Hopper"), ("gb200", "GB200"), ("gb300", "GB300"), ("vr200", "VR200"))
    for gk, gz in gens:
        t_ = {t: cD(f"tok_{gk}_{t}") for t in ("top", "mid", "low")}
        mm = f"({cD('m_paid_top')}/{t_['top']}+{cD('m_paid_mid')}/{t_['mid']}+{cD('m_paid_low')}/{t_['low']})"
        rr = f"({d('mx_top')}/{t_['top']}+{d('mx_mid')}/{t_['mid']}+{d('mx_low')}/{t_['low']})"
        row(f"r_{gk}", f"層級組合換算比（{gz}）＝（本模型組合的 Σ占比÷每 GW 總產出）÷（TK 參考組合的同式）", "倍", f"={mm}/{rr}", tag="Derived",
            note="每 GW 產能 ∝ 1 ÷ Σ（層級占比 ÷ 該層每 GW 總產出）；比值＝TK 參考組合的產能 ÷ 本模型組合的產能")
    fam_gen = (("hopper", "hopper"), ("gb200", "gb200"), ("gb300", "gb300"), ("vr200", "vr200"), ("tpu6", "hopper"), ("tpu7", "vr200"), ("trn2", "hopper"), ("trn3", "vr200"), ("mi455", "vr200"))
    num = "+".join(f"{cD('s_' + f)}*{cD('cap_paid_' + f)}*{d('r_' + g)}" for f, g in fam_gen)
    row("capmod", "本模型付費層級組合下 2025 機隊每 GW 付費產能（組合後）", "M tok/GW/年", f"={cD('cap_paid')}", tag="Derived（Assumed 層級組合；Analogy 產能）", note="Compute C115")
    row("capref", "TK 參考層級組合下 2025 機隊每 GW 付費產能（逐族：本模型該族產能 × 該族所屬世代的換算比）", "M tok/GW/年", f"={num}",
        tag="Derived（假設：TPU v6e、Trainium2 隨 Hopper；TPU v7、Trainium3、MI455X 隨 VR200，同本模型 NonNV 設定）", note="Σ 組合占比 × 該族產能 × 換算比")
    row("fw", "f_W：工作負載（層級）組合因子", "倍", f"={d('capref')}/{d('capmod')}", tag="Derived",
        note="只含層級組合；本模型層級組合的 Inputs 區間未接入此因子（無區間）。輸入輸出組合無法分離（並入殘差）", name="ETA_fW")
    row("fw_vr", "對照：僅以 VR200 近似的 f_W（舊口徑，TK L1_TokGW_PaidCap_VR200 × IF_Util ÷ 本模型 VR200 付費產能）", "倍",
        f"=TK_L1_TokGW_PaidCap_VR200*TK_IF_Util/{cD('cap_paid_vr200')}", tag="Derived", note="2025 機隊是 Hopper／TPU v6e／Trainium2，故正式口徑用機隊加權")
    row("fio", "輸入輸出組合因子（請求形狀、每則 token 數）", "倍", '="無法分離"', tag="無法分離", note="每任務 token 已取自 TK 任務模板（新鮮＋快取＋decode）；無獨立可比的 TK 參考請求組合 → 並入殘差")

    sec("五、乘積與未解釋殘差")
    row("prod", "已解釋部分 f_P × f_U × f_W", "倍", f"={d('fp')}*{d('fu')}*{d('fw')}", f"={d('fp','E')}*{d('fu','E')}*{d('fw')}", f"={d('fp','F')}*{d('fu','F')}*{d('fw')}",
        tag="Derived", name="ETA_Prod")
    row("res", "未解釋殘差 ＝ η ÷ 已解釋部分", "倍", f"={d('eta')}/{d('prod')}", f"={d('eta')}/{d('prod','F')}", f"={d('eta')}/{d('prod','E')}",
        tag="Derived（殘差；非獨立估計）", note="含：輸入輸出組合、每任務 token 倍數（Assumed）、任務數與支出的口徑差、研發占比等；低＝除以高端乘積", name="ETA_Resid")
    row("res_inv", "殘差倒數（TK 方向；支出路線 ÷ token 路線）", "倍", f"=1/{d('res')}", f"=1/{d('res','F')}", f"=1/{d('res','E')}", tag="Derived")

    sec("六、並列：TK L1_Gap*（OpenAI 口徑，Analogy；TK 方向）與本模型各因子倒數（TK 方向）")
    for key, nm_, zh in (("gp", "L1_GapPrompt", "TK 落差 1：每則提示 token 數（對應本模型：無法分離）"),
                         ("gs", "L1_GapSpendBasis", "TK 落差 2：推論支出計價 ÷ 持有成本"),
                         ("gi", "L1_GapISL", "TK 落差 3：請求形狀（1 ÷ κ）"),
                         ("gu", "L1_GapUtil", "TK 落差 4：利用率"),
                         ("gpr", "L1_GapProduct", "TK 四項乘積（已解釋部分）"),
                         ("gx", "L1_GapUnexplained", "TK 落差 5：未解釋")):
        row(key, zh, "倍", f"=TK_{nm_}", f"=TK_{nm_}_Lo", f"=TK_{nm_}_Hi", tag="Analogy（TK；OpenAI 口徑）", note=nm_)
    row("m_fp", "本模型：1 ÷ f_P（計價口徑，TK 方向）", "倍", f"=1/{d('fp')}", f"=1/{d('fp','F')}", f"=1/{d('fp','E')}", tag="Derived")
    row("m_fu", "本模型：1 ÷ f_U（利用率，TK 方向；未計生產折減）", "倍", f"=1/{d('fu')}", f"=1/{d('fu','F')}", f"=1/{d('fu','E')}", tag="Derived")
    row("m_fw", "本模型：1 ÷ f_W（工作負載層級組合，TK 方向）", "倍", f"=1/{d('fw')}", tag="Derived")
    row("m_res", "本模型：1 ÷ 殘差（TK 方向）", "倍", f"={d('res_inv')}", f"={d('res_inv','E')}", f"={d('res_inv','F')}", tag="Derived")
    row("m_prod", "本模型：1 ÷ 已解釋乘積（TK 方向）", "倍", f"=1/{d('prod')}", f"=1/{d('prod','F')}", f"=1/{d('prod','E')}", tag="Derived")
    row("r_all", "本模型 1 ÷ η（TK 方向的總落差；支出路線 ÷ token 路線）", "倍", f"=1/{d('eta')}", tag="Derived", note="＝1 ÷ η；TK 對應 L1_ScaleServe2000（約 10.4；OpenAI 口徑）")
    # ── v0.1.4 補充（新列加在尾端；η 值不變）：利用率因子計入 TK 生產折減、殘差對照 TK L1_GapUnexplained ──
    sec("七、v0.1.4 補充：f_U 計入 TK 生產折減（Andy 2026-10-22 預設）；殘差隨之重算。η 值不變，第五節（不計生產折減）原列保留以供前後對照")
    row("pd", "TK 生產折減（IF_ProdDerate；具名範圍）", "比例", "=TK_IF_ProdDerate", tag="Analogy（TK；OpenAI 口徑）",
        note="Interface C355（＝Serving CTL_ProdDerate）；token 換算 GW 的每 GW 產能是 100% 理論值，實際生產環境再打此折減")
    row("fu2", "f_U'：利用率因子（計入生產折減）＝f_U × 生產折減", "倍", f"={d('fu')}*{d('pd')}", f"={d('fu','E')}*{d('pd')}", f"={d('fu','F')}*{d('pd')}",
        tag="Derived（以 Analogy）", note="TK 方向 L1_GapUtil＝IF_Util ÷（實際 × 生產折減）＝本因子倒數", name="ETA_fU_Prod")
    row("prod2", "已解釋部分 f_P × f_U' × f_W", "倍", f"={d('fp')}*{d('fu2')}*{d('fw')}", f"={d('fp','E')}*{d('fu2','E')}*{d('fw')}", f"={d('fp','F')}*{d('fu2','F')}*{d('fw')}",
        tag="Derived", name="ETA_Prod2")
    row("res2", "未解釋殘差（計入生產折減後）＝η ÷ 已解釋部分", "倍", f"={d('eta')}/{d('prod2')}", f"={d('eta')}/{d('prod2','F')}", f"={d('eta')}/{d('prod2','E')}",
        tag="Derived（殘差；非獨立估計）", note="低＝除以高端乘積；高＝除以低端乘積", name="ETA_Resid2")
    row("res_chg", "殘差：計入生產折減後 ÷ 前（第五節）", "倍", f"={d('res2')}/{d('res')}", tag="Derived", note="＝1 ÷ 生產折減")
    row("res2_inv", "殘差倒數（TK 方向；計入生產折減後）", "倍", f"=1/{d('res2')}", f"=1/{d('res2','F')}", f"=1/{d('res2','E')}", tag="Derived")
    row("res_tk", "對照：TK L1_GapUnexplained（OpenAI 口徑，Analogy；TK 方向）", "倍", "=TK_L1_GapUnexplained", "=TK_L1_GapUnexplained_Lo", "=TK_L1_GapUnexplained_Hi",
        tag="Analogy（TK；OpenAI 口徑）", note="L1 D111；區間為 TK 低／高")
    row("res_ratio", "比值：本模型殘差倒數（計入生產折減）÷ TK L1_GapUnexplained", "倍", f"={d('res2_inv')}/{d('res_tk')}", tag="Derived",
        note="＜1：本模型未解釋部分較 TK 小；兩者口徑不同（Anthropic vs OpenAI），只並排；基準對基準，區間端點不相除")
    row("res_ratio_old", "比值（前）：第五節殘差倒數（不計生產折減）÷ TK L1_GapUnexplained", "倍", f"={d('res_inv')}/{d('res_tk')}", tag="Derived", note="v0.1.3 口徑，供前後對照")
    row("gu_chk", "核對：f_U' 的倒數 − TK L1_GapUtil（同一組輸入 0.41、0.85、0.60，應為 0；確認生產折減讀值與 TK 一致）", "倍", f"=1/{d('fu2')}-{d('gu')}", tag="Derived", note="本模型借用 TK 的 OpenAI 利用率輸入（Analogy），故與 L1_GapUtil 相同；若不為 0 表示 TK 改了定義")
    E.column_dimensions["A"].width = 7
    E.column_dimensions["B"].width = 78
    E.column_dimensions["C"].width = 12
    for c in "DEF":
        E.column_dimensions[c].width = 14
    E.column_dimensions["G"].width = 40
    E.column_dimensions["H"].width = 90
    E.freeze_panes = "D6"
    for r in range(6, st["r"]):
        E[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ctx.P["a5"] = dict(B=B, E=E, er=st["rows"])
    # Bridge 排在 Cost 之後（其餘頁順序不變；EtaDiag 緊接 Bridge）
    names = wb.sheetnames
    for sh in ("Bridge", "EtaDiag"):
        wb.move_sheet(sh, offset=names.index("Cost") + 1 + (0 if sh == "Bridge" else 1) - wb.sheetnames.index(sh))


def checks(ctx):
    B, er = ctx.P["a5"]["B"], ctx.P["a5"]["er"]
    br = B.rows
    mx = lambda key: f"=MAX(MAX(Bridge!$D${br[key]}:$I${br[key]}),-MIN(Bridge!$D${br[key]}:$I${br[key]}))"  # noqa: E731
    e = lambda k, col="D": f"EtaDiag!${col}${er[k]}"  # noqa: E731
    return [
        ("e1_iso", "Bridge、EtaDiag 被計算頁引用的公式格數（builder 掃描）", ctx.scan["bridge_refs"], 0, "eq", "v0.1.3 B：兩頁唯讀；tests 每次重驗（test_bridge_not_fed_back）"),
        ("e1_rev", "Bridge：（訂閱＋重算 API＋其他）− 營收總額（各年最大絕對差，$B）", mx("ck_rev"), 0, "tol", "token × 單價 ＝ API 營收的重算核對"),
        ("e1_gw", "Bridge：token 換算 GW ÷ η − 有效推論 GW（各年最大絕對差，GW）", mx("gw_chk"), 0, "tol", "Compute 第十節恆等式"),
        ("e1_eta", "EtaDiag：被拆解的 η − Compute 的 η（2025）", f"={e('eta')}-INDEX(CMP_Eta,1,1)", 0, "tol", "拆解不改 η"),
        ("e1_idn", "EtaDiag：f_P × f_U × f_W × 殘差 − η", f"={e('prod')}*{e('res')}-{e('eta')}", 0, "tol", "拆解恆等式"),
        ("e1_fp", "結果：η 拆解 f_P（計價口徑）", f"={e('fp')}", None, "info", "合約實付 ÷ TK 持有成本（加權）"),
        ("e1_fu", "結果：η 拆解 f_U（利用率；TK 實際 ÷ 基準）", f"={e('fu')}", None, "info", "Analogy"),
        ("e1_fw", "結果：η 拆解 f_W（層級組合）", f"={e('fw')}", None, "info", ""),
        ("e1_res", "結果：η 拆解未解釋殘差", f"={e('res')}", None, "info", "含輸入輸出組合、每任務 token 倍數等"),
        ("e1_tkq", "結果：Bridge 本模型 Q ÷ TK Q（2026）", f"=Bridge!$E${br['rt_q']}", None, "info", "對照 TK 帳未閉合倍數（Tokenomics 待辦）"),
        ("e1_fu2", "結果：η 拆解 f_U'（計入 TK 生產折減）", f"={e('fu2')}", None, "info", "v0.1.4：f_U × IF_ProdDerate"),
        ("e1_res2", "結果：未解釋殘差（計入生產折減後）", f"={e('res2')}", None, "info", "v0.1.4；前值見 e1_res"),
        ("e1_idn2", "EtaDiag：f_P × f_U' × f_W × 殘差' − η", f"={e('prod2')}*{e('res2')}-{e('eta')}", 0, "tol", "v0.1.4 拆解恆等式"),
        ("e1_gapr", "結果：本模型殘差倒數 ÷ TK L1_GapUnexplained", f"={e('res_ratio')}", None, "info", "v0.1.4；OpenAI 口徑 Analogy，只並排"),
    ]
