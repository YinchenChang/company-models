"""v0.6-P5（工作單 S5）：融資頁 Funding 與反向模式頁 Reverse。

機制：r6 第 P5 節 1–3 點；v0.5「資金與檢查」「反向」兩頁的公式結構（S7 最低現金、來源順序、缺口旗標；S4b 供應商融資只列或有負債；
Amazon 條件式 $35B 不計入基準、另列情境；外部資金需求＝年底現金低於最低現金之差額累計；R1–R3 反向模式只作對照，不回饋基準）。
S4「給 S5」：融資缺口起點 COST_GapCash；股權報酬 COST_SBC 為非現金，現金流加回；2025 現金對帳用 COST_Actual2025。
數值一律引用 SRC_OAI／Inputs／REV_／DEM_／COST_ 具名範圍或同簿儲存格（E6：公式不內含常數；恆等式 1−x、1＋成長率、年數 +1 除外）。
Python 只產生結構；計算由 Excel（LibreOffice）與 engine 執行。Reverse 頁只被 Checks 引用（不回饋基準；test_builder 檢查）。
"""
from __future__ import annotations

from p2 import V05, YC, YEARS, Sheet


def build_p5(wb, R, S, I, inp_row, P2, P3, P4):
    D, V = P2["D"], P2["V"]
    dr, vr = D.rows, V.rows
    kr = P4["K"].rows
    v05 = V05()
    mincash, nv_ratio = I("最低現金"), I("Nvidia $30B 現金比例")
    y_round, amz_on, y_amz = I("2026-03 融資輪無條件部分的現金到位年"), I("計入 Amazon 條件式 $35B"), I("Amazon 條件式 $35B 到位年")
    cash25 = S("現金：2025 年底")
    amz = S("股權融資：條件式部分（Amazon）")
    mb, months = I("單位換算：$M → $B 的除數"), I("每年月數")
    flow = lambda i: i > 0  # noqa: E731  現金流自 2026 起（期初＝2025 年底現金，SRC）
    cost = lambda key, c: f"Cost!{c}{kr[key]}"  # noqa: E731
    rev = lambda key, c: f"Revenue!{c}{vr[key]}"  # noqa: E731

    # ═══════════════════════════ Funding ═══════════════════════════
    F = Sheet(wb, "Funding", "F", "Funding — 融資（S7）：自由現金流、已到位融資、最低現金與外部資金需求、Amazon 條件式情境、或有負債（S4b）",
              "r6 第 P5 節第 1 點（沿用 v0.5「資金與檢查」頁）：自由現金流＝營收淨額 − 算力成本（現金）− 非算力成本（不含股權報酬；股權報酬為非現金，加回）；"
              "來源順序：①期初現金（2025 年底，SRC）→ ②已到位股權融資（2026-03 輪無條件部分）→ ③外部資金（新股權、債務或延後承諾；不分種類）。"
              "外部資金需求＝各年年底現金低於最低現金的差額（補足至最低現金）之累計。Amazon 條件式 $35B 不計入基準（第四節另列情境）；供應商融資只列或有負債。",
              "D–I＝2025–2030；2025 為校準年：只列年底現金（SRC）與自由現金流對照，現金流自 2026 起。K–P＝v0.5 原值（對照，不參與計算；v0.5 支出列為負號者不列）。$B＝十億美元。")

    F.section("一、自由現金流（S4 給 S5：COST_GapCash ＋ 股權報酬加回）")
    F.add("營收淨額（截頂後；扣 Microsoft 分成）", "$B", lambda i, c: f"={rev('net_c', c)}", v05.row("資金與檢查", 5),
          "REV_NetCapped", key="rev")
    F.add("算力成本（現金口徑）＝合約實付＋自有資本支出", "$B", lambda i, c: f"={cost('cc', c)}", None, "COST_Compute（v0.5 資金第 6、7 列，負號）", key="cc")
    F.add("非算力成本（不含股權報酬）", "$B", lambda i, c: f"={cost('ncx', c)}", None, "COST_NonCompExSBC（v0.5 資金第 8 列：營運費用，不含股票薪酬）", key="ncx")
    F.add("股權報酬（非現金；只列示）", "$B", lambda i, c: f"={cost('sbc', c)}", None, "COST_SBC：不耗用現金（S4 給 S5：計算現金缺口時加回）", key="sbc")
    F.add("差額（含股權報酬；S4 命題口徑）", "$B", lambda i, c: f"={cost('p_gap_b', c)}", None, "COST_GapCash（S4 給 S5 的融資缺口起點）", key="gap")
    F.add("自由現金流＝差額＋股權報酬", "$B", lambda i, c: f"={c}{F.rows['gap']}+{c}{F.rows['sbc']}", v05.row("資金與檢查", 9),
          "＝營收淨額 − 算力成本 − 非算力成本（Checks 恆等式）；2025 欄為模型值（合約實付口徑），不進現金流", name="FND_FCF", key="fcf")
    F.add("對照：2025 以實際算力支出（20.4）計的自由現金流", "$B", [f"=D{F.rows['rev']}-COST_Actual2025-D{F.rows['ncx']}"] + [None] * 5, None,
          "S4 給 S5：2025 現金對帳用實際算力支出（推論 8.4＋訓練 10.59＋其他雲端 1.41）", name="FND_FCF2025Act", name_col="D", key="fcf25")

    F.section("二、已到位融資（來源順序 ②；2026-03 輪無條件部分＝SRC 四個組成；Nvidia 以現金比例計，S4b）")
    F.add("2026-03 輪到位年旗標", "旗標", lambda i, c: f"=N({c}$6={y_round})", None, "Inputs：到位年（v0.5：全額計入 2026）", key="fl_r")
    for k, lab, f_, note in (
        ("amzn1", "Amazon 首筆", lambda c: f"={S('股權融資 2026-03 輪：Amazon 首筆')}*{c}{F.rows['fl_r']}", "SRC_OAI"),
        ("sb", "SoftBank（三期）", lambda c: f"={S('股權融資 2026-03 輪：SoftBank（三期）')}*{c}{F.rows['fl_r']}", "SRC_OAI；三期分期時點未揭露，全額計入到位年"),
        ("nv", "Nvidia（現金部分）", lambda c: f"={S('股權融資 2026-03 輪：Nvidia')}*{nv_ratio}*{c}{F.rows['fl_r']}", "SRC_OAI × Nvidia 現金比例（Inputs，基準 1；S4b）"),
        ("oth", "其他", lambda c: f"={S('股權融資 2026-03 輪：其他')}*{c}{F.rows['fl_r']}", "SRC_OAI（約）"),
    ):
        F.add(f"已到位：{lab}", "$B", lambda i, c, f_=f_: f_(c), None, note, key=f"in_{k}")
    F.add("已到位融資合計", "$B", lambda i, c: "=" + "+".join(f"{c}{F.rows['in_' + k]}" for k in ("amzn1", "sb", "nv", "oth")), None,
          "v0.5 無條件部分 87（Nvidia 以現金計）", name="FND_Committed", key="committed")
    F.add("對照：Nvidia 投資的非現金（算力）部分（只列示）", "$B", lambda i, c: f"={S('股權融資 2026-03 輪：Nvidia')}*(1-{nv_ratio})*{c}{F.rows['fl_r']}", None,
          "S4b：另有報導稱 Nvidia 投資多為算力；基準現金比例 1 時為 0", key="nv_nc")
    F.add("Amazon 條件式 $35B（基準不計入；Inputs 開關）", "$B", lambda i, c: f"={amz}*{amz_on}*N({c}$6={y_amz})", None,
          "S5 預設：不計入基準（開關 0）；計入情境見第四節", key="amz")
    F.add("計入基準的股權流入", "$B", lambda i, c: f"={c}{F.rows['committed']}+{c}{F.rows['amz']}", v05.row("資金與檢查", 10), "", name="FND_EquityIn", key="eq")

    F.section("三、現金與外部資金需求（S7：最低現金、來源順序、缺口旗標；外部資金需求＝年底現金低於最低現金的差額累計）")
    F.add("最低現金", "$B", lambda i, c: f"={mincash}", None, "Inputs（v0.5 S7 沿用 CRWV：10，區間 5–20）", name="FND_MinCash", key="min")
    r_end = F.r + 3
    F.add("期初現金（來源順序 ①）", "$B", lambda i, c: None if i == 0 else (f"={cash25}" if i == 1 else f"={YC[i - 1]}{r_end}"), v05.row("資金與檢查", 11),
          "2026＝2025 年底現金（SRC_OAI，40；區間 35–45）；之後＝前一年年底現金", name="FND_CashOpen", key="open")
    F.add("補足前年底現金＝期初＋自由現金流＋股權流入", "$B",
          lambda i, c: f"={c}{F.rows['open']}+{c}{F.rows['fcf']}+{c}{F.rows['eq']}" if flow(i) else None, v05.row("資金與檢查", 12), "", key="pre")
    F.add("當年外部資金需求（來源順序 ③）＝最低現金 − 補足前（補足前低於最低現金時；否則 0）", "$B",
          lambda i, c: f"=({c}{F.rows['min']}-{c}{F.rows['pre']})*({c}{F.rows['min']}>{c}{F.rows['pre']})" if flow(i) else None, v05.row("資金與檢查", 13),
          "新股權、債務或延後承諾（v0.5 externalNeed；不分種類）", name="FND_ExtNeed", key="ext")
    F.add("年底現金", "$B", lambda i, c: f"={cash25}" if i == 0 else f"={c}{F.rows['pre']}+{c}{F.rows['ext']}", v05.row("資金與檢查", 14),
          "2025＝SRC_OAI（實際）；2026 起＝補足前＋外部資金", name="FND_CashEnd", key="end")
    assert F.rows["end"] == r_end
    F.add("累計外部資金需求", "$B", lambda i, c: None if i == 0 else (f"={c}{F.rows['ext']}" if i == 1 else f"={YC[i - 1]}{F.r}+{c}{F.rows['ext']}"),
          v05.row("資金與檢查", 15), "r6 P5-1：外部資金需求＝年底現金低於最低現金的累計差額", name="FND_ExtNeedCum", key="cum")
    F.add("缺口旗標（1＝當年需要外部資金）", "旗標", lambda i, c: f"=SIGN({c}{F.rows['ext']})" if flow(i) else None, None,
          "S7 缺口旗標（v0.5 沿用 CRWV）", name="FND_GapFlag", key="flag")
    F.add("缺口年度標示（旗標＝1 的年度）", "年", lambda i, c: f'=IF({c}{F.rows["flag"]},{c}$6,"")' if flow(i) else None, None, "Checks 取最小值＝首次需要外部資金的年度", key="flag_y")
    r_ne = F.r
    F.add("只靠已到位融資（不另募）的年底現金", "$B",
          lambda i, c: None if i == 0 else (f"={cash25}+{c}{F.rows['fcf']}+{c}{F.rows['eq']}" if i == 1 else f"={YC[i - 1]}{r_ne}+{c}{F.rows['fcf']}+{c}{F.rows['eq']}"),
          None, "顯示已到位融資能撐到哪一年（無外部資金時的現金路徑；可為負）", name="FND_CashNoExt", key="noext")
    F.add("只靠已到位融資：低於最低現金的年度標示", "年", lambda i, c: f'=IF({c}{F.rows["noext"]}<{c}{F.rows["min"]},{c}$6,"")' if flow(i) else None, None,
          "Checks 取最小值＝已到位融資撐不過的第一年", key="noext_y")
    F.add("峰值年度標示（當年外部資金需求＝2026–2030 最大值）", "年",
          lambda i, c: f'=IF(AND({c}{F.rows["flag"]},{c}{F.rows["ext"]}=MAX($E${F.rows["ext"]}:$I${F.rows["ext"]})),{c}$6,"")' if flow(i) else None, None, "", key="peak_y")

    F.section("四、情境：計入 Amazon 條件式 $35B（S5 預設：不計入基準，列情境；不受 Inputs 開關影響）")
    F.add("Amazon 條件式流入（情境）", "$B", lambda i, c: f"={amz}*N({c}$6={y_amz})", None, "SRC_OAI_082 × 到位年（Inputs，2026；2026–2027）", key="s_amz")
    F.add("股權流入（情境）＝已到位融資＋Amazon 條件式", "$B", lambda i, c: f"={c}{F.rows['committed']}+{c}{F.rows['s_amz']}", None, "", key="s_eq")
    r_send = F.r + 2
    F.add("補足前年底現金（情境）", "$B",
          lambda i, c: None if i == 0 else (f"={cash25}+{c}{F.rows['fcf']}+{c}{F.rows['s_eq']}" if i == 1 else f"={YC[i - 1]}{r_send}+{c}{F.rows['fcf']}+{c}{F.rows['s_eq']}"),
          None, "", key="s_pre")
    F.add("當年外部資金需求（情境）", "$B", lambda i, c: f"=({c}{F.rows['min']}-{c}{F.rows['s_pre']})*({c}{F.rows['min']}>{c}{F.rows['s_pre']})" if flow(i) else None, None, "", key="s_ext")
    F.add("年底現金（情境）", "$B", lambda i, c: f"={cash25}" if i == 0 else f"={c}{F.rows['s_pre']}+{c}{F.rows['s_ext']}", None, "", key="s_end")
    assert F.rows["s_end"] == r_send
    F.add("累計外部資金需求（情境）", "$B", lambda i, c: None if i == 0 else (f"={c}{F.rows['s_ext']}" if i == 1 else f"={YC[i - 1]}{F.r}+{c}{F.rows['s_ext']}"),
          None, "", name="FND_ExtNeedCumAmzn", key="s_cum")

    F.section("五、或有負債與承諾（S4b：只列示，不作資金來源）")
    F.add("或有負債：Nvidia 擔保", "$B", [f"={S('或有負債：Nvidia 擔保')}"] + [None] * 5, v05.row("資金與檢查", 66), "SRC_OAI（WSJ）；v0.5 資金第 66 列", key="cl_g")
    F.add("或有負債：Nvidia 晶片融資（洽談中）", "$B", [f"={S('或有負債：Nvidia 晶片融資洽談')}"] + [None] * 5, v05.row("資金與檢查", 67),
          "SRC_OAI（WSJ）；v0.5 資金第 67 列", key="cl_c")
    F.add("或有負債合計（只列示）", "$B", [f"=D{F.rows['cl_g']}+D{F.rows['cl_c']}"] + [None] * 5, None, "不進現金流與外部資金需求（S4b）", name="FND_Contingent", name_col="D", key="cl")
    F.add("對照：表外承諾存量（S-1 草案轉述）", "$B", [f"={S('表外承諾存量')}"] + [None] * 5, None, "SRC_OAI；只對照", key="obs")
    keys = [k for k, *_ in P3["contracts_full"]]
    tot = "+".join(t[2] for t in P3["contracts_full"])
    F.add("合約剩餘承諾（2030 年後）＝各合約總額 − 2025–2030 實付合計", "$B",
          [f"=({tot})-SUM({cost('pay', 'D')}:{cost('pay', 'I')})"] + [None] * 5, [v05.wb["支出"]["J11"].value] + [None] * 5,
          "v0.5 支出頁 J 欄（2030 年後剩餘）；只列示", name="FND_CommitAfter2030", name_col="D", key="after")

    for c, w in (("A", 7), ("B", 62), ("C", 10)):
        F.ws.column_dimensions[c].width = w
    for c in YC:
        F.ws.column_dimensions[c].width = 12
    F.ws.column_dimensions["Q"].width = 80
    F.ws.freeze_panes = "D7"

    # ═══════════════════════════ Reverse ═══════════════════════════
    X = Sheet(wb, "Reverse", "X", "Reverse — 反向模式（R1–R3）：假設管理層營收目標達成，反解充分條件；只作對照區塊，不回饋基準",
              "r6 第 P5 節第 2 點（沿用 v0.5「反向」頁與「資金與檢查」反向段）：R1 目標只驅動本頁；R2 三組充分條件（只靠 API、只靠訂閱、兩線等比例）以正向值的倍數表示；"
              "R3 廣告取正向值（內部廣告目標只作對照）。本頁不被任何基準頁引用（共同規則第 4 節：不以管理層目標反推參數）。",
              "目標路徑：2025＝實際；2026、2030、Σ840＝SRC_OAI（FT 2026-09-18，7 月簡報）；2027–28＝擬合值（Inputs）；2029＝Σ840 殘差。K–P＝v0.5 原值（對照）。")
    X.section("一、管理層目標營收路徑（R1：只驅動本頁）")
    fit = {2027: I("管理層營收目標擬合值（反向模式）", "2027"), 2028: I("管理層營收目標擬合值（反向模式）", "2028")}
    t26, t30, tsum = S("管理層營收目標：2026"), S("管理層營收目標：2030"), S("管理層營收目標：2026–2030 累計")

    def target(i, c):
        if i == 0:
            return f"={S('營收：FY2025')}"
        if i == 1:
            return f"={t26}"
        if YEARS[i] in fit:
            return f"={fit[YEARS[i]]}"
        if i == 4:
            return f"={tsum}-E{X.r}-F{X.r}-G{X.r}-I{X.r}"
        return f"={t30}"
    X.add("管理層營收目標", "$B", target, v05.row("反向", 4), "2029＝Σ840 − 其餘四年（v0.5／Tokenomics v4 FT_Sep2026 列 24）", name="RVS_Target", key="tgt")
    X.add("檢查：2026–2030 合計 − SRC Σ840", "$B", [f"=SUM(E{X.rows['tgt']}:I{X.rows['tgt']})-{tsum}"] + [None] * 5, None, "應為 0", key="tsum_ck")
    X.add("目標年增率", "比例", lambda i, c: None if i == 0 else f"=({c}{X.rows['tgt']}-{YC[i - 1]}{X.rows['tgt']})/{YC[i - 1]}{X.rows['tgt']}", None,
          "v0.5：成長率約每年遞減 0.2", key="tgt_g")
    X.add("對照：舊版 2030 目標（2026-02 版）", "$B", [None] * 5 + [f"={S('管理層營收目標（舊版）：2030（2026-02 版）')}"], None, "SRC_OAI；只對照", key="tgt_old")

    X.section("二、正向對照（本模型基準；截頂後）")
    capf = lambda c: rev("capf", c)  # noqa: E731
    X.add("正向總額（截頂後）", "$B", lambda i, c: f"={rev('gross_c', c)}", v05.row("反向", 5), "REV_GrossCapped", key="fwd")
    for k, lab, key_ in (("sub", "訂閱", "sub"), ("api", "API", "api"), ("ads", "廣告", "ads"), ("oth", "其他", "other")):
        X.add(f"正向{lab}（截頂後）", "$B", lambda i, c, key_=key_: f"={rev(key_, c)}*{capf(c)}", None, "各營收線同比例截頂（v0.5）", key=f"f_{k}")
    X.add("差距（目標 − 正向）", "$B", lambda i, c: f"={c}{X.rows['tgt']}-{c}{X.rows['fwd']}", v05.row("反向", 6), "", name="RVS_Gap", key="gap")
    X.add("核心需求＝目標 − 正向廣告 − 正向其他（R3：廣告取正向值）", "$B", lambda i, c: f"={c}{X.rows['tgt']}-{c}{X.rows['f_ads']}-{c}{X.rows['f_oth']}",
          v05.row("反向", 7), "v0.5 反向第 7 列（其他營收 v0.5 為 0）", key="core")

    X.section("三、R2 充分條件組：所需倍數（相對正向；1＝正向已足夠）")
    X.add("只靠 API：API 營收所需倍數", "倍", lambda i, c: f"=({c}{X.rows['core']}-{c}{X.rows['f_sub']})/{c}{X.rows['f_api']}", v05.row("反向", 9), "",
          name="RVS_MultAPI", key="m_api")
    X.add("只靠訂閱：訂閱營收所需倍數", "倍", lambda i, c: f"=({c}{X.rows['core']}-{c}{X.rows['f_api']})/{c}{X.rows['f_sub']}", v05.row("反向", 10), "",
          name="RVS_MultSub", key="m_sub")
    X.add("兩線等比例：訂閱與 API 共同倍數", "倍", lambda i, c: f"={c}{X.rows['core']}/({c}{X.rows['f_sub']}+{c}{X.rows['f_api']})", v05.row("反向", 11), "",
          name="RVS_MultProp", key="m_prop")
    X.add("對照：廣告取內部路徑時的等比例倍數", "倍",
          lambda i, c: None if i == 0 else f"=({c}{X.rows['tgt']}-{S(f'內部廣告營收目標：{YEARS[i]}')}-{c}{X.rows['f_oth']})/({c}{X.rows['f_sub']}+{c}{X.rows['f_api']})",
          v05.row("反向", 12), "SRC_OAI 內部廣告營收目標（Axios 2026-04-09）；只對照", key="m_ads")

    X.section("四、2030 換算為可觀測量（I 欄）")
    only30 = lambda f: [None] * 5 + [f]  # noqa: E731
    X.add("正向付費用戶（Go＋Plus＋Pro＋席位）", "M", only30(f"=Demand!I{dr['u_paid']}"), v05.row("反向", 14), "DEM_Users_Paid", key="o_users")
    X.add("正向混合 ARPU＝訂閱營收 ÷ 付費用戶", "$/月", only30(f"=I{X.rows['f_sub']}*{mb}/{months}/I{X.rows['o_users']}"), v05.row("反向", 15), "", key="o_arpu")
    X.add("只靠訂閱：所需付費用戶（同 ARPU）", "M", only30(f"=(I{X.rows['core']}-I{X.rows['f_api']})*{mb}/{months}/I{X.rows['o_arpu']}"), v05.row("反向", 16), "", key="o_need_u")
    X.add("對照：2030 週用戶數計畫", "M", only30(f"={S('2030 週用戶數計畫')}"), None, "SRC_OAI（Axios 2026-04-09；廣告預測的用戶假設）", key="o_wau")
    X.add("正向 API 計費 token", "T", only30(f"=Demand!I{dr['api_tok']}"), v05.row("反向", 17), "DEM_Tok_API", key="o_tok")
    X.add("只靠 API：所需計費 token（同單價路徑）", "T", only30(f"=(I{X.rows['core']}-I{X.rows['f_sub']})/{rev('mixprice', 'I')}*{mb}"), v05.row("反向", 18),
          "REV_ApiPrice（$/M）", key="o_need_t")

    X.section("五、橋接 2030（示意；簡報營收組成未揭露，順序不同則各項不同）")
    X.add("正向總額", "$B", only30(f"=I{X.rows['fwd']}"), v05.row("反向", 21), "", key="b_fwd")
    X.add("＋廣告由正向提高至內部路徑", "$B", only30(f"={S('內部廣告營收目標：2030')}-I{X.rows['f_ads']}"), v05.row("反向", 22), "", key="b_ads")
    X.add("＋用戶提高至計畫（訂閱等比放大）", "$B",
          only30(f"=I{X.rows['f_sub']}*({S('2030 週用戶數計畫')}-(Demand!I{dr['u_free']}+I{X.rows['o_users']}))/(Demand!I{dr['u_free']}+I{X.rows['o_users']})"),
          v05.row("反向", 23), "總用戶＝Free＋付費（v0.5 同）", key="b_users")
    X.add("＋核心殘差", "$B", only30(f"=I{X.rows['tgt']}-I{X.rows['b_fwd']}-I{X.rows['b_ads']}-I{X.rows['b_users']}"), v05.row("反向", 24), "", key="b_res")
    X.add("＝管理層目標", "$B", only30(f"=I{X.rows['b_fwd']}+I{X.rows['b_ads']}+I{X.rows['b_users']}+I{X.rows['b_res']}"), v05.row("反向", 25), "", key="b_tgt")
    X.add("核心殘差折算：API 所需倍數", "倍", only30(f"=(I{X.rows['f_api']}+I{X.rows['b_res']})/I{X.rows['f_api']}"), v05.row("反向", 26), "", key="b_mapi")
    X.add("廣告步驟占差距比例", "比例", only30(f"=I{X.rows['b_ads']}/(I{X.rows['tgt']}-I{X.rows['b_fwd']})"), v05.row("反向", 27), "", key="b_adsh")
    X.add("廣告內部路徑占管理層 2030 目標", "比例", only30(f"={S('內部廣告營收目標：2030')}/I{X.rows['tgt']}"), v05.row("反向", 28), "", key="b_adtgt")

    X.section("六、反向資金（管理層目標營收＋同一支出；S7 同一規則；只作對照）")
    rate, cap, thru, cfrom = S("Microsoft 營收分成率"), S("Microsoft 分成總額上限"), S("Microsoft 分成付款持續至"), I("分成上限起算年")
    msr, cumr = X.r, X.r + 1

    def ms(i, c):
        tg = f"{c}{X.rows['tgt']}"
        pre = f"{rate}*{tg}*({c}$6<{cfrom})"
        if i == 0:
            return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{tg},{cap})"
        return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{tg},{cap}-{YC[i - 1]}{cumr})"
    X.add("Microsoft 分成（目標營收；同上限規則）", "$B", ms, v05.row("資金與檢查", 18), "同 Revenue R28 公式，以目標營收計", key="ms")
    X.add("Microsoft 分成累計（自起算年）", "$B", lambda i, c: f"={c}{msr}*({c}$6>={cfrom})" if i == 0 else f"={YC[i - 1]}{cumr}+{c}{msr}*({c}$6>={cfrom})",
          v05.row("資金與檢查", 19), "", name="RVS_MSCum", key="mscum")
    X.add("淨營收（目標）", "$B", lambda i, c: f"={c}{X.rows['tgt']}-{c}{X.rows['ms']}", v05.row("資金與檢查", 22), "", key="net")
    X.add("算力成本（現金；同正向）", "$B", lambda i, c: f"={cost('cc', c)}", None, "COST_Compute（v0.5 反向第 23、24 列，負號）", key="cc")
    X.add("非算力成本（不含股權報酬；同正向，不隨目標營收放大）", "$B", lambda i, c: f"={cost('ncx', c)}", None,
          "S5 預設：『同一支出』（v0.6 非算力成本為人數法，不以營收占比推）；v0.5 反向第 25 列隨目標營收放大，見下方對照列", key="ncx")
    X.add("自由現金流（反向）", "$B", lambda i, c: f"={c}{X.rows['net']}-{c}{X.rows['cc']}-{c}{X.rows['ncx']}", v05.row("資金與檢查", 26), "", name="RVS_FCF", key="fcf")
    X.add("股權流入（同基準）", "$B", lambda i, c: f"=Funding!{c}{F.rows['eq']}", v05.row("資金與檢查", 27), "FND_EquityIn", key="eq")
    r_xend = X.r + 3
    X.add("期初現金", "$B", lambda i, c: None if i == 0 else (f"={cash25}" if i == 1 else f"={YC[i - 1]}{r_xend}"), v05.row("資金與檢查", 28), "", key="open")
    X.add("補足前年底現金", "$B", lambda i, c: f"={c}{X.rows['open']}+{c}{X.rows['fcf']}+{c}{X.rows['eq']}" if flow(i) else None, v05.row("資金與檢查", 29), "", key="pre")
    X.add("當年外部資金需求", "$B", lambda i, c: f"=(Funding!{c}{F.rows['min']}-{c}{X.rows['pre']})*(Funding!{c}{F.rows['min']}>{c}{X.rows['pre']})" if flow(i) else None, v05.row("資金與檢查", 30), "", key="ext")
    X.add("年底現金", "$B", lambda i, c: f"={cash25}" if i == 0 else f"={c}{X.rows['pre']}+{c}{X.rows['ext']}", v05.row("資金與檢查", 31), "", key="end")
    assert X.rows["end"] == r_xend
    X.add("累計外部資金需求（反向）", "$B", lambda i, c: None if i == 0 else (f"={c}{X.rows['ext']}" if i == 1 else f"={YC[i - 1]}{X.r}+{c}{X.rows['ext']}"),
          v05.row("資金與檢查", 32), "", name="RVS_ExtNeedCum", key="cum")
    X.add("對照：非算力成本隨目標營收等比放大（v0.5 口徑）", "$B", lambda i, c: f"={c}{X.rows['ncx']}*{c}{X.rows['tgt']}/{c}{X.rows['fwd']}",
          v05.row("支出", 24), "v0.5：營運費用＝目標營收 × 營收占比；此列以 v0.6 非算力成本 × 目標 ÷ 正向總額近似", key="ncx_v05")
    X.add("對照：放大口徑下自由現金流的變動（放大 − 同一支出）", "$B", lambda i, c: f"={c}{X.rows['ncx']}-{c}{X.rows['ncx_v05']}", None, "負＝放大口徑的現金流較差", key="ncx_d")

    X.section("七、與管理層簡報對帳（2026–2030 累計；D 欄）")
    one = lambda f: [f] + [None] * 5  # noqa: E731
    X.add("本模型反向 Σ 自由現金流 2026–30", "$B", one(f"=SUM(E{X.rows['fcf']}:I{X.rows['fcf']})"), v05.row("資金與檢查", 35), "v0.5 資金第 35 列", key="r_fcf")
    X.add("簡報 Σ 自由現金流 2026–30", "$B", one(f"={S('管理層目標：2026–2030 累計自由現金流')}"), v05.row("資金與檢查", 36), "SRC_OAI（FT 2026-09-18）", key="r_deck")
    X.add("差異（本模型 − 簡報）", "$B", one(f"=D{X.rows['r_fcf']}-D{X.rows['r_deck']}"), v05.row("資金與檢查", 37), "", key="r_diff")
    X.add("其中：算力（簡報 856 − 本模型 Σ 算力成本 2026–30）", "$B", one(f"={S('計畫算力支出：2026–2030 累計')}-SUM({cost('cc', 'E')}:{cost('cc', 'I')})"),
          v05.row("資金與檢查", 38), "SRC_OAI（FT 2026-09-18：$856B 算力與基礎設施）", key="r_cc")
    X.add("其中：其餘（營運費用、分成上限、口徑）", "$B", one(f"=D{X.rows['r_diff']}-D{X.rows['r_cc']}"), v05.row("資金與檢查", 39), "", key="r_oth")
    X.add("對照：本模型正向 Σ 自由現金流 2026–30", "$B", one(f"=SUM(Funding!E{F.rows['fcf']}:I{F.rows['fcf']})"), [v05.wb["資金與檢查"]["I9"].value] + [None] * 5, "v0.5 資金第 9 列 I 欄", key="r_fwd")

    for c, w in (("A", 7), ("B", 62), ("C", 10)):
        X.ws.column_dimensions[c].width = w
    for c in YC:
        X.ws.column_dimensions[c].width = 12
    X.ws.column_dimensions["Q"].width = 80
    X.ws.freeze_panes = "D7"
    return dict(F=F, X=X)
