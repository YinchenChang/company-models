"""v0.6-P3（工作單 S3）：算力頁 Compute 與 Revenue 第八節（容量上限 V1a）；v0.6-P4（S4）加第十三節自建 GW（計入供給）。

機制：r6 第 P3 節 1–7 點、V1、V2、V8、V11；v0.5「算力MW」頁的供給（S2、N4(b)）與容量上限（V1a）結構。
產能一律取 TK_Link（IF_TokGW_* 世代 × 層級、IF_Util）；本模型不重算算力物理量。
數值一律引用 SRC_OAI／Inputs／TK_／DEM_ 具名範圍或同簿儲存格（E6：公式不內含常數；恆等式 1−x、年數 +1 除外）。
Python 只產生結構：每列的公式文字由此組裝，計算由 Excel（LibreOffice）與 engine 執行。
"""
from __future__ import annotations

from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, put
from p2 import YC, YEARS, Sheet, nm

GEN_KEYS = ("hopper", "gb200", "gb300", "vr200", "ru", "custom")
GEN_ZH = {"hopper": "Hopper", "gb200": "GB200", "gb300": "GB300", "vr200": "VR200", "ru": "Rubin Ultra", "custom": "自研／其他"}
GEN_NAME = {"hopper": "Hopper", "gb200": "GB200", "gb300": "GB300", "vr200": "VR200", "ru": "RubinUltra", "custom": "Custom"}
SUP_NAME = {"azure": "Azure", "oracle": "Oracle", "aws": "AWSNvidia", "trn": "AWSTrainium", "cw": "CoreWeave", "cer": "Cerebras"}
TIER_TK = {"top": "Astra", "mid": "Sol", "low": "Luna"}
TIER_ZH = {"top": "頂層（Astra）", "mid": "中層（Sol）", "low": "低層（Luna）"}


def build_p3(wb, R, S, I, inp_row, P2, snap):
    """建立 Compute 頁，並在 Revenue 加第八節（容量上限）。回傳供 Checks 使用的物件。"""
    D, V = P2["D"], P2["V"]
    yidx = lambda name, y: I(name, str(y))  # noqa: E731
    cal = I("校準年（第 1 個）")
    tm, idle, mintr, price = I("單位換算：T token → M token"), I("閒置保留比例"), I("訓練最低占比"), I("每 GW 年合約價")
    custom = I("世代產能係數：自研／其他")
    # Tokenomics 世代名稱與成本情境標籤（取自快照表頭；只作 SUMIFS 的對應鍵，不是數值）
    tk_gens = list(dict.fromkeys(snap["hdr_gen"]))
    base_cost = snap["hdr_cost"][1]
    assert len(tk_gens) == 5, tk_gens
    GC = dict(zip(GEN_KEYS, YC))                       # 世代 → 第一節的欄（D–I）

    C = Sheet(wb, "Compute", "G", "Compute — 算力需求與供給：世代組合、每 GW 年產能、推論 GW（token 換算）、η、研發 GW、供給 GW（逐合約）、容量上限、VR 等值",
              "r6 第 P3 節：推論 GW＝token ÷（Σ 世代占比 × 各世代每 GW 年產能），付費／免費分開（V11）；V1 世代組合時變（Blackwell 拆 GB200／GB300；自研＝VR200 × 係數）；"
              "V2 η＝2025 token 換算 GW ÷ 支出換算 GW；供給依合約（v0.5 S2、N4(b)）；營收以可用 GW 為上限（V1a）。",
              "GW＝IT 關鍵電力（實體 GW；外部揭露 GW 口徑未明者視同）。產能取 TK_Link（IF_TokGW_*、IF_Util），不重算物理量。第一節 D–I 欄＝世代；其餘各節 D–I＝2025–2030。")
    for c in "KLMNOP":
        C.ws[f"{c}5"].value = None                     # 本頁無 v0.5 原值欄

    # ═════ 一、世代參數 ═════
    C.section("一、世代參數（本節 D–I 欄＝世代：Hopper／GB200／GB300／VR200／Rubin Ultra／自研／其他；非年度）")
    C.add("Tokenomics 世代名稱（對應 TK_HdrGen）", "文字", tk_gens + ["—（VR200 × 自研係數）"], None,
          "Tokenomics 快照表頭的世代名稱（文字鍵）；v0.5 Vera Rubin＝VR200，v0.5 下一代（Rubin Ultra 以後）＝Rubin Ultra", key="g_lab")
    C.add("Tokenomics 成本情境（對應 TK_HdrCost）", "文字", [base_cost] * 5 + ["—"], None,
          "取基準成本欄（每 GW 總產出三個成本情境同值）", key="g_cost")
    rl, rc = C.rows["g_lab"], C.rows["g_cost"]
    for t in ("top", "mid", "low"):
        tk = TIER_TK[t]
        C.add(f"每 GW 總產出：{TIER_ZH[t]}（100% 利用率）", "M tok/GW/年",
              lambda i, c, tk=tk, t=t: f"=SUMIFS(TK_IF_TokGW_{tk},TK_HdrGen,{c}${rl},TK_HdrCost,{c}${rc})" if i < 5
              else f"=$G${C.r}*{custom}",
              None, f"TK IF_TokGW_{tk}（世代 × 基準成本欄）；自研／其他＝VR200 × Inputs 自研係數（r6 P3-2）", key=f"g_tok_{t}")
    rsol = C.rows["g_tok_mid"]
    C.add("VR200 等值係數（Sol 層級每 GW 產能比）", "倍", lambda i, c: f"={c}{rsol}/$G${rsol}", None,
          "工作單預設：VR 等值以 Sol 層級產能比換算（供 S4 命題輸出）；自研＝係數本身", name="CMP_GenVReq", key="g_vr")

    # ═════ 二、V1 世代組合 ═════
    C.section("二、V1 世代組合（實體 GW 占比；時變；Blackwell 拆 GB200／GB300）")
    C.add("Hopper", "比例", lambda i, c: f"={yidx('實體 GW 世代占比：Hopper', YEARS[i])}", None, "Inputs（v0.5 fleetMix）",
          name="CMP_Share_Hopper", key="s_hopper")
    C.add("Blackwell 合計（v0.5）", "比例", lambda i, c: f"={yidx('實體 GW 世代占比：Blackwell', YEARS[i])}", None, "Inputs（v0.5 fleetMix）", key="s_bw")
    g3 = [I("GB300 占 Blackwell 比例", "2025")] + [I("GB300 占 Blackwell 比例", "2026 起")] * 5
    C.add("GB300 占 Blackwell 比例", "比例", lambda i, c: f"={g3[i]}", None, "Inputs（r6 P3-2：2025 為 0；2026 起 0.5；區間 0–0.8）", key="g3")
    C.add("GB200", "比例", lambda i, c: f"={c}{C.rows['s_bw']}*(1-{c}{C.rows['g3']})", None, "Blackwell ×（1 − GB300 占比）",
          name="CMP_Share_GB200", key="s_gb200")
    C.add("GB300", "比例", lambda i, c: f"={c}{C.rows['s_bw']}*{c}{C.rows['g3']}", None, "Blackwell × GB300 占比", name="CMP_Share_GB300", key="s_gb300")
    C.add("VR200（v0.5 Vera Rubin）", "比例", lambda i, c: f"={yidx('實體 GW 世代占比：Vera Rubin', YEARS[i])}", None, "Inputs（v0.5 fleetMix）",
          name="CMP_Share_VR200", key="s_vr200")
    C.add("Rubin Ultra（v0.5 下一代）", "比例", lambda i, c: f"={yidx('實體 GW 世代占比：下一代', YEARS[i])}", None,
          "Inputs（v0.5 fleetMix『下一代』＝Rubin Ultra 以後；TK 第 5 世代）", name="CMP_Share_RubinUltra", key="s_ru")
    C.add("自研／其他", "比例",
          lambda i, c: f"=(1-{c}{C.rows['s_hopper']}-{c}{C.rows['s_bw']}-{c}{C.rows['s_vr200']}-{c}{C.rows['s_ru']})", None,
          "＝1 − 其餘（v0.5 fleetMix 註記）", name="CMP_Share_Custom", key="s_custom")
    C.add("合計（應＝1）", "比例", lambda i, c: "=" + "+".join(f"{c}{C.rows['s_' + g]}" for g in GEN_KEYS), None, "Checks 檢查", key="s_sum")

    # ═════ 三、層級組合 ═════
    C.section("三、層級組合（Demand 的 token 量 ÷ 合計；付費、免費分開）")
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        for t in ("top", "mid", "low"):
            C.add(f"{zh}：{TIER_ZH[t]}占比", "比例",
                  lambda i, c, kind=kind, t=t: f"=Demand!{c}{D.rows[f'{kind}_{t}']}/Demand!{c}{D.rows[f'{kind}_all']}", None,
                  f"Demand {zh} token（{TIER_ZH[t]}）÷ {zh} token 合計", key=f"m_{kind}_{t}")

    # ═════ 四、每 GW 年產能 ═════
    C.section("四、每 GW 年產能（M tok/GW/年；基準利用率 TK IF_Util；V11：本模型組合 × TK 世代產能；式同 Tokenomics Alloc）")
    for kind, zh in (("paid", "付費"), ("free", "免費")):
        for g in GEN_KEYS:
            X = GC[g]
            C.add(f"{zh}：{GEN_ZH[g]}", "M tok/GW/年",
                  lambda i, c, kind=kind, X=X: "=TK_IF_Util/(" + "+".join(
                      f"{c}{C.rows[f'm_{kind}_{t}']}/${X}${C.rows[f'g_tok_{t}']}" for t in ("top", "mid", "low")) + ")",
                  None, "IF_Util ÷ Σ（層級占比 ÷ 該世代該層級每 GW 總產出）", key=f"cap_{kind}_{g}")
        C.add(f"{zh}：世代組合後每 GW 年產能", "M tok/GW/年",
              lambda i, c, kind=kind: "=" + "+".join(f"{c}{C.rows['s_' + g]}*{c}{C.rows[f'cap_{kind}_{g}']}" for g in GEN_KEYS), None,
              "Σ 世代占比 × 各世代產能（r6 P3-1）", name=f"CMP_Cap{kind.capitalize()}", key=f"cap_{kind}")

    # ═════ 五、推論 GW（token 換算） ═════
    C.section("五、推論 GW（token 換算）＝token ÷ 每 GW 年產能（r6 P3-1；付費、免費分開）")
    C.add("付費 token", "M tok/年", lambda i, c: f"=Demand!{c}{D.rows['paid_all']}*{tm}", None, "Demand 付費 token（T）× 10^6（Inputs 定義常數）；API 為計費 token（S2 預設 7）", key="tok_paid")
    C.add("免費 token", "M tok/年", lambda i, c: f"=Demand!{c}{D.rows['free_all']}*{tm}", None, "Demand 免費 token（T）× 10^6", key="tok_free")
    C.add("推論 GW（token 換算）：付費", "GW", lambda i, c: f"={c}{C.rows['tok_paid']}/{c}{C.rows['cap_paid']}", None, "", name="CMP_TokGW_Paid", key="gw_paid")
    C.add("推論 GW（token 換算）：免費", "GW", lambda i, c: f"={c}{C.rows['tok_free']}/{c}{C.rows['cap_free']}", None, "", name="CMP_TokGW_Free", key="gw_free")
    C.add("推論 GW（token 換算）：合計", "GW", lambda i, c: f"={c}{C.rows['gw_paid']}+{c}{C.rows['gw_free']}", None, "", name="CMP_TokGW", key="gw_tok")

    # ═════ 六、V2 η ═════
    C.section("六、V2 有效產出係數 η（2025 計算鏈；之後各年依 Inputs 路徑開關）與有效推論 GW")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    C.add("2025 推論支出（SRC）", "$B", only25(f"={S('推論成本合計：2025')}"), None, "SRC_OAI（The Information 2026-02；付費 4.5＋非付費 3.9）", key="e_spend")
    C.add("每 GW 年合約價", "$B/GW/年", only25(f"={price}"), None, "Inputs（V8：唯一合約價參數；Analogy 12，8–20）", key="e_price")
    C.add("2025 推論 GW（支出換算）＝推論支出 ÷ 合約價", "GW", only25(f"=D{C.rows['e_spend']}/D{C.rows['e_price']}"), None, "r6 P3-3",
          name="CMP_SpendGW2025", name_col="D", key="e_spendgw")
    C.add("2025 推論 GW（token 換算）", "GW", only25(f"=D{C.rows['gw_tok']}"), None, "第五節合計列 2025 欄", key="e_tokgw")
    C.add("η（2025）＝token 換算 GW ÷ 支出換算 GW", "倍", only25(f"=D{C.rows['e_tokgw']}/D{C.rows['e_spendgw']}"), None,
          "V2：模型產能實際兌現的比例（Derived）", name="CMP_Eta2025", name_col="D", key="eta25")
    e25 = f"$D${C.rows['eta25']}"
    sw, tgt, ty = I("η 逐年路徑開關"), I("η 回升情境：目標值"), I("η 回升情境：目標年")
    C.add("η（逐年）", "倍", lambda i, c: f"={e25}+{sw}*({tgt}-{e25})*MIN({ty}-{cal},{c}$6-{cal})/({ty}-{cal})", None,
          "＝η₂₀₂₅＋開關 ×（目標 − η₂₀₂₅）× MIN(目標年 − 校準年, 年 − 校準年) ÷（目標年 − 校準年）；開關 0（基準）＝沿用 2025", name="CMP_Eta", key="eta")
    C.add("有效推論 GW：付費", "GW", lambda i, c: f"={c}{C.rows['gw_paid']}/{c}{C.rows['eta']}", None, "token 換算 GW ÷ η", key="eff_paid")
    C.add("有效推論 GW：免費", "GW", lambda i, c: f"={c}{C.rows['gw_free']}/{c}{C.rows['eta']}", None, "token 換算 GW ÷ η", key="eff_free")
    C.add("有效推論 GW（未截頂）", "GW", lambda i, c: f"={c}{C.rows['gw_tok']}/{c}{C.rows['eta']}", None,
          "2025＝支出換算 GW（校準恆等式）", name="CMP_InfGW_Eff", key="eff")

    # ═════ 七、三角對照 ═════
    C.section("七、三角對照（2025；r6 P3-3：只列差距，不校準）")
    C.add("2025 算力支出＝推論 ＋ 訓練（10.59＋其他雲端）", "$B",
          only25(f"={S('推論成本合計：2025')}+{S('研發費用中付 Microsoft 部分：2025')}+{I('2025 其他雲端')}"), None, "SRC_OAI＋Inputs（V6）", key="t_spend")
    C.add("2025 年均算力 GW＝AVERAGE（2024 年底, 2025 年底）", "GW", only25(f"=AVERAGE({S('算力規模：2024 年底')},{S('算力規模：2025 年底')})"), None,
          "SRC_OAI 0.6、1.9（CFO 部落格；口徑未明）", key="t_avg")
    C.add("隱含每 GW 年算力支出", "$B/GW/年", only25(f"=D{C.rows['t_spend']}/D{C.rows['t_avg']}"), None, "r6 參考值約 16.3",
          name="CMP_ImpliedPrice2025", name_col="D", key="t_impl")
    C.add("差距：隱含 − 合約價", "$B/GW/年", only25(f"=D{C.rows['t_impl']}-{price}"), None, "", key="t_gap")
    C.add("差距占合約價", "比例", only25(f"=D{C.rows['t_gap']}/{price}"), None, "r6：約 +36%", key="t_gapp")
    C.add("以隱含值計算的 η", "倍", only25(f"=D{C.rows['gw_tok']}/(D{C.rows['e_spend']}/D{C.rows['t_impl']})"), None, "只並列，不取代 η", key="t_eta")
    C.add("η 差距：隱含值口徑 − 基準", "倍", only25(f"=D{C.rows['t_eta']}-D{C.rows['eta25']}"), None, "", key="t_etagap")

    # ═════ 八、研發 GW 與總需求 GW ═════
    C.section("八、研發 GW 與總需求 GW（r6 P3-4、P3-5；沿用 v0.5：推論優先、其餘歸研發）")
    C.add("2025 訓練支出（研發算力）＝10.59＋其他雲端", "$B", only25(f"={S('研發費用中付 Microsoft 部分：2025')}+{I('2025 其他雲端')}"), None,
          "V6：SRC_OAI＋Inputs", key="r_train25")
    C.add("研發 GW", "GW", lambda i, c: f"=D{C.rows['r_train25']}/{price}" if i == 0 else f"={c}{{SUP}}*(1-{idle})-{c}{{INFCAP}}", None,
          "2025＝訓練支出 ÷ 合約價（r6 P3-4）；2026 起＝供給 ×（1 − 閒置）− 推論 GW（截頂後）（v0.5 算力MW 第 34 列）", name="CMP_RDGW", key="rd")
    C.add("總需求 GW＝有效推論（未截頂）＋研發", "GW", lambda i, c: f"={c}{C.rows['eff']}+{c}{C.rows['rd']}", None, "r6 P3-5", name="CMP_DemandGW", key="dem")
    C.add("對照：SRC 2024 年底算力", "GW", only25(f"={S('算力規模：2024 年底')}"), None, "CFO 部落格（口徑未明）", key="r_e24")
    C.add("對照：SRC 2025 年底算力", "GW", only25(f"={S('算力規模：2025 年底')}"), None, "同上（約 1.9）", key="r_e25")
    C.add("差距：2025 總需求 − 年均（0.6／1.9 平均）", "GW", only25(f"=D{C.rows['dem']}-D{C.rows['t_avg']}"), None, "只列差距", name="CMP_DemGapAvg2025", name_col="D", key="r_gapavg")
    C.add("差距：2025 總需求 − 2025 年底", "GW", only25(f"=D{C.rows['dem']}-D{C.rows['r_e25']}"), None, "只列差距", key="r_gapend")

    # ═════ 九、供給 GW ═════
    C.section("九、供給 GW（逐合約；v0.5 S2、N4(b)：揭露 GW 者期間內爬升至揭露值；未揭露者＝付款 ÷ 合約價）")
    prof, k2 = I("未揭露時程合約的攤提方式"), I("線性爬升付款：等差級數和")
    contracts = [  # key, 名稱, 總額, 起, 迄, 揭露 GW, 型態（und＝未揭露；ramp＝揭露 GW 爬升；flat＝揭露年額）, 說明
        ("azure", "Microsoft Azure", S("合約總額：Microsoft Azure（OpenAI 增購）"), I("合約期間起：Microsoft Azure"), I("合約期間迄：Microsoft Azure"), None, "und",
         "未揭露 GW：付款 ÷ 合約價；期間 Inputs（V13、V16）"),
        ("oracle", "Oracle", S("合約總額：Oracle"), S("合約期間起：Oracle"), S("合約期間迄：Oracle"), S("合約容量：Oracle"), "flat",
         "揭露年額（300÷5＝60，E8f）與 4.5 GW：期間內固定（v0.5 揭露年額欄）"),
        ("aws", "AWS（Nvidia 晶片）", S("合約總額：AWS（Nvidia 晶片）"), S("合約期間起：AWS（Nvidia 晶片）"), S("合約期間迄：AWS（Nvidia 晶片）"), None, "und",
         "未揭露 GW：付款 ÷ 合約價"),
        ("trn", "AWS（Trainium）", S("合約總額：AWS（Trainium）"), S("合約期間起：AWS（Trainium）"), S("合約期間迄：AWS（Trainium）"),
         I("合約容量（採用值）：AWS（Trainium）"), "ramp", "揭露 2 GW（Inputs 採用值，2–5）：期間內線性爬升至揭露值"),
        ("cw", "CoreWeave", S("合約總額：CoreWeave"), S("合約期間起：CoreWeave"), S("合約期間迄：CoreWeave"), None, "und", "未揭露 GW：付款 ÷ 合約價"),
        ("cer", "Cerebras", I("合約總額（估計）：Cerebras"), S("合約期間起：Cerebras"), I("合約期間迄：Cerebras"), S("合約容量：Cerebras"), "ramp",
         "揭露 0.75 GW：期間內線性爬升至揭露值；總額與終點 Inputs（V13）"),
    ]

    def supply_f(c, T, s, e, gw, kind):
        inn = f"({c}$6>={s})*({c}$6<={e})"
        n = f"({e}-{s}+1)"
        if kind == "flat":
            return f"={inn}*{gw}"
        if kind == "ramp":
            return f'={inn}*IF({prof}="even",{gw},{gw}*({c}$6-{s}+1)/{n})'
        return f'={inn}*IF({prof}="even",{T}/{n},{T}*({c}$6-{s}+1)/({n}*({n}+1)/{k2}))/{price}'

    for key, zh, T, s, e, gw, kind, note in contracts:
        C.add(f"供給 GW：{zh}", "GW", lambda i, c, T=T, s=s, e=e, gw=gw, kind=kind: supply_f(c, T, s, e, gw, kind), None,
              note + "；攤提方式 Inputs（N4(b) ramp）", name=f"CMP_Supply_{SUP_NAME[key]}", key=f"sup_{key}")
    disc = [k for k, *_r in contracts if _r[5] != "und"]
    und = [k for k, *_r in contracts if _r[5] == "und"]
    C.add("小計：揭露 GW 的合約", "GW", lambda i, c: "=" + "+".join(f"{c}{C.rows['sup_' + k]}" for k in disc), None, "Oracle、AWS Trainium、Cerebras", key="sup_disc")
    C.add("小計：未揭露 GW（付款 ÷ 合約價）", "GW", lambda i, c: "=" + "+".join(f"{c}{C.rows['sup_' + k]}" for k in und), None,
          "Microsoft Azure、AWS（Nvidia）、CoreWeave；與合約價成反比（敏感度 A）", key="sup_und")
    C.add("供給 GW 合計（合約＋自建）", "GW", lambda i, c: f"={c}{C.rows['sup_disc']}+{c}{C.rows['sup_und']}+{c}{{OWN}}", None,
          "合約（不自動假設新增合約）＋自建 GW（S4 預設：自有資本支出自投產年起計入供給與命題分母；見第十三節）", name="CMP_SupplyGW", key="sup")
    C.add("對照：2025 實際年均算力 − 合約供給", "GW", only25(f"=D{C.rows['t_avg']}-D{C.rows['sup']}"), None, "只列差距", key="sup_gap25")

    # ═════ 十、容量上限 ═════
    C.section("十、容量上限（V1a）：推論可用 GW＝供給 ×（1 − 閒置）×（1 − 訓練最低占比）；超出時營收按比例截頂並立旗標")
    C.add("閒置 GW（校準年後）", "GW", lambda i, c: f"={c}{C.rows['sup']}*{idle}*({c}$6>{cal})", None, "v0.5：閒置保留自 2026 起", key="idle")
    C.add("推論可用 GW", "GW", lambda i, c: f"={c}{C.rows['sup']}*(1-{idle})*(1-{mintr})", None, "v0.5 V1a（Inputs 閒置 0.1、訓練最低占比 0.3）",
          name="CMP_InfAvailGW", key="avail")
    C.add("推論 GW（截頂後）", "GW", lambda i, c: f"=IF({c}$6<={cal},{c}{C.rows['eff']},MIN({c}{C.rows['avail']},{c}{C.rows['eff']}))", None,
          "校準年＝有效推論 GW（2025 營收＝實際）；之後＝MIN（可用, 有效）", name="CMP_InfGW", key="infcap")
    C.add("容量上限係數（1＝未受限）", "倍", lambda i, c: f"={c}{C.rows['infcap']}/{c}{C.rows['eff']}", None, "＝截頂後 ÷ 有效（v0.5 MIN(1, 可用÷需求)）",
          name="CMP_CapFactor", key="cap")
    C.add("缺口旗標（1＝需求超過可用 GW）", "旗標", lambda i, c: f"=({c}{C.rows['eff']}>{c}{C.rows['avail']})*({c}$6>{cal})", None,
          "工作單預設：不自動假設新增合約", name="CMP_CapFlag", key="flag")
    C.add("缺口 GW（有效 − 截頂後）", "GW", lambda i, c: f"={c}{C.rows['eff']}-{c}{C.rows['infcap']}", None, "", key="excess")
    C.add("檢查：推論（截頂後）＋研發＋閒置 − 供給", "GW",
          lambda i, c: f"={c}{C.rows['infcap']}+{c}{C.rows['rd']}+{c}{C.rows['idle']}-{c}{C.rows['sup']}", None,
          "2026 起應為 0；2025 研發 GW 以支出換算，差額＝2025 支出換算 GW 對合約供給的差距", key="ident")
    # 回填研發 GW 的列號
    r = C.rows["rd"]
    for i in range(1, 6):
        C.ws[f"{YC[i]}{r}"].value = C.ws[f"{YC[i]}{r}"].value.replace("{SUP}", str(C.rows["sup"])).replace("{INFCAP}", str(C.rows["infcap"]))

    # ═════ 十一、VR 等值換算 ═════
    C.section("十一、VR 等值換算（供 S4 命題輸出；工作單預設：Sol 層級產能比）")
    C.add("機隊 VR 等值係數＝Σ 世代占比 × VR200 等值係數", "倍",
          lambda i, c: "=" + "+".join(f"{c}{C.rows['s_' + g]}*${GC[g]}${C.rows['g_vr']}" for g in GEN_KEYS), None, "", name="CMP_VReqFactor", key="vrf")
    C.add("供給 VR 等值 GW", "GW", lambda i, c: f"={c}{C.rows['sup']}*{c}{C.rows['vrf']}", None, "", name="CMP_Supply_VReq", key="vr_sup")
    C.add("推論 GW（截頂後）VR 等值", "GW", lambda i, c: f"={c}{C.rows['infcap']}*{c}{C.rows['vrf']}", None, "", name="CMP_InfGW_VReq", key="vr_inf")
    C.add("研發 GW VR 等值", "GW", lambda i, c: f"={c}{C.rows['rd']}*{c}{C.rows['vrf']}", None, "", name="CMP_RDGW_VReq", key="vr_rd")
    C.add("對照：依付費層級組合加權的 VR 等值係數", "倍", lambda i, c: f"={c}{C.rows['cap_paid']}/{c}{C.rows['cap_paid_vr200']}", None,
          "替代口徑（工作單預設表）：付費組合後產能 ÷ VR200 付費產能", key="vrf_alt")

    # ═════ 十二、敏感度 ═════
    C.section("十二、敏感度（Excel 公式；引用 Inputs 值／低／高欄；A、C、D 以 η 沿用路徑計算）")
    rp = inp_row(price)

    def coef(c, eff, sup):
        return f"=IF({c}$6<={cal},{c}{eff},MIN({c}{sup}*(1-{idle})*(1-{mintr}),{c}{eff}))/{c}{eff}"

    sens = {}
    for tag, col, zh in (("lo", "F", "低端"), ("base", "E", "基準"), ("hi", "G", "高端")):
        P = f"Inputs!${col}${rp}"
        k = f"A_{tag}"
        C.add(f"A 合約價（{zh}）", "$B/GW/年", only25(f"={P}"), None, f"Inputs {price} 的{zh}欄", key=f"{k}_p")
        C.add(f"A η（2025；合約價{zh}）", "倍", only25(f"=D{C.rows['gw_tok']}/(D{C.rows['e_spend']}/D{C.rows[k + '_p']})"), None, "", key=f"{k}_eta")
        C.add(f"A 有效推論 GW（合約價{zh}）", "GW", lambda i, c, k=k: f"={c}{C.rows['gw_tok']}/$D${C.rows[k + '_eta']}", None, "η 沿用", key=f"{k}_eff")
        C.add(f"A 合約供給 GW（合約價{zh}）", "GW",
              lambda i, c, k=k: f"={c}{C.rows['sup_disc']}+{c}{C.rows['sup_und']}*{price}/$D${C.rows[k + '_p']}+{c}{{OWN}}", None,
              "揭露 GW 不變；未揭露者與合約價成反比；自建 GW 不隨合約價變動（S4）", key=f"{k}_sup")
        C.add(f"A 容量上限係數（合約價{zh}）", "倍", lambda i, c, k=k: coef(c, C.rows[k + "_eff"], C.rows[k + "_sup"]), None, "", key=f"{k}_cap")
        sens[k] = k
    for tag, zh, f in (("keep", "沿用（基準）", lambda i, c: f"={e25}"),
                       ("rise", "線性回升至目標", lambda i, c: f"={e25}+({tgt}-{e25})*MIN({ty}-{cal},{c}$6-{cal})/({ty}-{cal})")):
        k = f"B_{tag}"
        C.add(f"B η 路徑：{zh}", "倍", f, None, "", key=f"{k}_eta")
        C.add(f"B 有效推論 GW（η {zh}）", "GW", lambda i, c, k=k: f"={c}{C.rows['gw_tok']}/{c}{C.rows[k + '_eta']}", None, "", key=f"{k}_eff")
        C.add(f"B 容量上限係數（η {zh}）", "倍", lambda i, c, k=k: coef(c, C.rows[k + "_eff"], C.rows["sup"]), None, "供給為基準", key=f"{k}_cap")
    r3 = inp_row(I("GB300 占 Blackwell 比例", "2026 起"))
    rcu = inp_row(custom)
    for grp, rowx, zh_grp, delta in (
        ("C", r3, "GB300 占 Blackwell（2026 起）",
         lambda c, G, kind: f"{c}{C.rows['s_bw']}*({G}-{c}{C.rows['g3']})*({c}{C.rows[f'cap_{kind}_gb300']}-{c}{C.rows[f'cap_{kind}_gb200']})*({c}$6>{cal})"),
        ("D", rcu, "自研產能係數",
         lambda c, G, kind: f"{c}{C.rows['s_custom']}*({G}-{custom})*{c}{C.rows[f'cap_{kind}_vr200']}"),
    ):
        for tag, col, zh in (("lo", "F", "低端"), ("base", "E", "基準"), ("hi", "G", "高端")):
            G = f"Inputs!${col}${rowx}"
            k = f"{grp}_{tag}"
            C.add(f"{grp} {zh_grp}（{zh}）", "比例" if grp == "C" else "倍", only25(f"={G}"), None, f"Inputs 的{zh}欄", key=f"{k}_x")
            C.add(f"{grp} 推論 GW（token 換算；{zh}）", "GW",
                  lambda i, c, G=G, delta=delta: f"={c}{C.rows['tok_paid']}/({c}{C.rows['cap_paid']}+{delta(c, G, 'paid')})"
                  f"+{c}{C.rows['tok_free']}/({c}{C.rows['cap_free']}+{delta(c, G, 'free')})", None,
                  "組合後產能隨參數線性變動（Σ 世代占比 × 各世代產能）", key=f"{k}_tok")
            C.add(f"{grp} 有效推論 GW（{zh}）", "GW",
                  lambda i, c, k=k: f"={c}{C.rows[k + '_tok']}/($D${C.rows[k + '_tok']}/$D${C.rows['e_spendgw']})", None,
                  "η 以該情境 2025 token GW 重新校準（沿用路徑）", key=f"{k}_eff")
            C.add(f"{grp} 容量上限係數（{zh}）", "倍", lambda i, c, k=k: coef(c, C.rows[k + "_eff"], C.rows["sup"]), None, "", key=f"{k}_cap")
            if grp == "D":
                C.add(f"D 機隊 VR 等值係數（{zh}）", "倍", lambda i, c, G=G: f"={c}{C.rows['vrf']}+{c}{C.rows['s_custom']}*({G}-{custom})", None, "", key=f"{k}_vrf")

    # ═════ 十三、自建 GW（P4；S4） ═════
    C.section("十三、自建 GW（P4；S4 預設：自有資本支出自投產年起計入供給與命題分母；v0.5 算力MW 第 11 列）")
    rl, rc = C.rows["g_lab"], C.rows["g_cost"]
    lag = I("自建 GW 投產落後年數")
    C.add("自有資本支出（毛額）", "$B", lambda i, c: f"={yidx('自有資本支出', YEARS[i])}", None, "Inputs（v0.5 ownedCapex；Assumed，低／高情境倍數 0.6／1.5）", key="own_capex")
    rcx = C.rows["own_capex"]
    C.add("累計自有資本支出（至 年 − 投產落後年數）", "$B",
          lambda i, c: f"=SUMPRODUCT(($D$6:$I$6<={c}$6-{lag})*($D${rcx}:$I${rcx}))", None,
          "v0.5：累計至前一年（落後 1 年；Inputs）；毛額（合作方出資部分仍形成 GW）", key="own_cum")
    C.add("每 GW 資本支出（VR200 基準；TK IF_CapexTotal）", "$B/GW",
          lambda i, c: f"=SUMIFS(TK_IF_CapexTotal,TK_HdrGen,$G${rl},TK_HdrCost,$G${rc})", None,
          "TK IF_CapexTotal（IT 設備＋廠房）VR200 基準欄（v0.5 用 Tokenomics DC_Cost 43.2）", key="own_cpg")
    C.add("自建 GW", "GW", lambda i, c: f"={c}{C.rows['own_cum']}/{c}{C.rows['own_cpg']}", None,
          "累計自有資本支出 ÷ 每 GW 資本支出（v0.5 同式）", name="CMP_Supply_Owned", key="own")
    C.add("合約供給 GW（不含自建）", "GW", lambda i, c: f"={c}{C.rows['sup_disc']}+{c}{C.rows['sup_und']}", None,
          "揭露＋未揭露（S3 的合約供給；供應商持有成本的 GW）", name="CMP_SupplyContractGW", key="sup_con")
    for r_ in (C.rows["sup"], *[C.rows[f"A_{t}_sup"] for t in ("lo", "base", "hi")]):
        for i in range(6):
            cell = C.ws[f"{YC[i]}{r_}"]
            cell.value = cell.value.replace("{OWN}", str(C.rows["own"]))

    for c, w in (("A", 7), ("B", 58), ("C", 12)):
        C.ws.column_dimensions[c].width = w
    for c in YC:
        C.ws.column_dimensions[c].width = 13
    C.ws.column_dimensions["Q"].width = 80
    C.ws.freeze_panes = "D7"

    # ═════ Revenue 第八節：容量上限（V1a） ═════
    V.r = V.ev_last + 2
    V.section("八、容量上限（V1a；P3）：截頂後營收＝未截頂 × 容量上限係數（v0.5 營收彙總第 4 列）；未截頂值（R27–R30）保留")
    V.add("容量上限係數（Compute）", "倍", lambda i, c: f"=Compute!{c}{C.rows['cap']}", None, "1＝未受限", name="REV_CapFactor", key="capf")
    V.add("截頂後總額", "$B", lambda i, c: f"={c}{V.rows['gross']}*{c}{V.rows['capf']}", None, "各營收線同比例截頂（v0.5）", name="REV_GrossCapped", key="gross_c")
    rate, mcap, thru, cfrom = S("Microsoft 營收分成率"), S("Microsoft 分成總額上限"), S("Microsoft 分成付款持續至"), I("分成上限起算年")
    gr, msr, cumr = V.rows["gross_c"], V.r, V.r + 1

    def ms(i, c):
        pre = f"{rate}*{c}{gr}*({c}$6<{cfrom})"
        if i == 0:
            return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{c}{gr},{mcap})"
        return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{c}{gr},{mcap}-{YC[i - 1]}{cumr})"

    V.add("截頂後 Microsoft 分成", "$B", ms, None, "同 R28 公式，以截頂後總額與截頂後累計計算", name="REV_MSShareCapped", key="ms_c")
    V.add("截頂後 Microsoft 分成累計（自起算年）", "$B",
          lambda i, c: f"={c}{msr}*({c}$6>={cfrom})" if i == 0 else f"={YC[i - 1]}{cumr}+{c}{msr}*({c}$6>={cfrom})", None, "",
          name="REV_MSCumCapped", key="mscum_c")
    V.add("截頂後淨額", "$B", lambda i, c: f"={c}{gr}-{c}{msr}", None, "估值與命題用（容量受限時）", name="REV_NetCapped", key="net_c")
    V.add("截頂減少的總額（未截頂 − 截頂後）", "$B", lambda i, c: f"={c}{V.rows['gross']}-{c}{gr}", None, "", key="cut")
    V.add("容量截頂旗標（Compute）", "旗標", lambda i, c: f"=Compute!{c}{C.rows['flag']}", None, "1＝需求超過可用 GW", key="flag")
    return dict(C=C, V=V, contracts=[k for k, *_ in contracts], contracts_full=contracts, prof=prof, k2=k2, tk_gens=tk_gens, base_cost=base_cost)
