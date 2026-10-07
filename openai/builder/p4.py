"""v0.6-P4（工作單 S4）：成本頁 Cost — 算力成本（V3）、供應商持有成本與雲端毛利、非算力成本（V4）、Q3、命題輸出（每 VR 等值 GW）、V16 Azure 攤入檢查。

機制：r6 第 P4 節 1–5 點、V3、V4、V6、V16；S4 工作單預設表（自建 GW 計入供給、現金口徑為基準、經濟口徑並列、
TK 單位成本參考以 IF_FullCostDefault_* 為基準、自研持有成本＝VR200 × genFactor.custom）。
合約實付沿用 v0.5「支出」頁（S1c：揭露年額者期間內固定；其餘依 N4(b) 攤提方式），與 Compute 第九節供給同一時程。
數值一律引用 SRC_OAI／Inputs／TK_／CMP_／REV_ 具名範圍或同簿儲存格（E6：公式不內含常數；恆等式 1−x、1＋成長率、年數 +1 除外）。
Python 只產生結構；計算由 Excel（LibreOffice）與 engine 執行。
"""
from __future__ import annotations

from p2 import V05, YC, YEARS, Sheet
from p3 import GEN_KEYS, GEN_ZH, SUP_NAME

GC = dict(zip(GEN_KEYS, YC))                       # 世代 → 第一節的欄（D–I）


def build_p4(wb, R, S, I, inp_row, P2, P3, snap):
    C, V = P3["C"], P3["V"]
    cr, vr = C.rows, V.rows
    v05 = V05()
    cal = I("校準年（第 1 個）")
    price, custom, k2, prof = I("每 GW 年合約價"), I("世代產能係數：自研／其他"), I("線性爬升付款：等差級數和"), I("未揭露時程合約的攤提方式")
    m2b, d2b = I("單位換算：$M → $B 的除數"), I("單位換算：$ → $B 的除數")
    lo_cost, hi_cost = snap["hdr_cost"][0], snap["hdr_cost"][2]

    K = Sheet(wb, "Cost", "K", "Cost — 成本、Q3 與單位經濟：算力成本（合約實付＋自有資本支出）、供應商持有成本與雲端毛利、非算力成本、Q3、命題輸出（每 VR 等值 GW）、V16 Azure 檢查",
              "r6 第 P4 節：V3 算力成本＝合約實付（v0.5 S1c、S2、N4(b)）＋自有資本支出；TK IF_HoldEcon（世代加權）只作供應商成本參考，差額＝雲端毛利；"
              "V4 非算力成本：2025＝財報；2026 起＝人數 × 每人年成本，股權報酬單列；命題輸出＝每 VR 等值 GW 的營收、成本、差額（現金口徑為基準，經濟口徑並列）。",
              "第一節 D–I 欄＝世代（Hopper／GB200／GB300／VR200／Rubin Ultra／自研／其他）；其餘各節 D–I＝2025–2030。K–P＝v0.5 原值（對照，不參與計算）。$B＝十億美元；GW＝IT 實體 GW。")
    only25 = lambda f: [f] + [None] * 5  # noqa: E731
    cmp_ = lambda key, c: f"Compute!{c}{cr[key]}"  # noqa: E731

    # ═════ 一、世代參數 ═════
    K.section("一、世代參數（本節 D–I 欄＝世代；TK IF_HoldEcon 經濟口徑持有成本、IF_CapexTotal；自研＝VR200 × 自研係數，S4 預設）")
    K.add("Tokenomics 世代名稱（Compute 第一節）", "文字", lambda i, c: f"=Compute!{c}{cr['g_lab']}" if i < 5 else None, None, "SUMIFS 的鍵", key="g_lab")
    K.add("成本情境鍵：低成本", "文字", [lo_cost] * 5 + [None], None, "TK_HdrCost 的低成本欄標籤（文字鍵）", key="g_lo")
    K.add("成本情境鍵：基準", "文字", lambda i, c: f"=Compute!{c}{cr['g_cost']}" if i < 5 else None, None, "", key="g_base")
    K.add("成本情境鍵：高成本", "文字", [hi_cost] * 5 + [None], None, "", key="g_hi")
    for tag, zh in (("base", "基準"), ("lo", "低成本"), ("hi", "高成本")):
        K.add(f"每 GW 年持有成本（經濟口徑；{zh}）", "$B/GW/年",
              lambda i, c, tag=tag: f"=SUMIFS(TK_IF_HoldEcon,TK_HdrGen,{c}${K.rows['g_lab']},TK_HdrCost,{c}${K.rows['g_' + tag]})" if i < 5
              else f"=$G${K.r}*{custom}", None,
              "TK IF_HoldEcon（資本回收年金 WACC 10%＋營運費用）；自研／其他＝VR200 × 自研係數（S4 預設，Analogy，區間隨係數 0.4–1.0）", key=f"hold_{tag}")
    K.add("每 GW 資本支出（基準；IT＋廠房）", "$B/GW",
          lambda i, c: f"=SUMIFS(TK_IF_CapexTotal,TK_HdrGen,{c}${K.rows['g_lab']},TK_HdrCost,{c}${K.rows['g_base']})" if i < 5 else None, None,
          "TK IF_CapexTotal；自建 GW 用 VR200 欄（Compute 第十三節）", key="capex_gw")

    # ═════ 二、算力成本（V3，現金口徑） ═════
    K.section("二、算力成本（V3；現金口徑）＝合約實付（v0.5 S1c：揭露年額者期間內固定；其餘依攤提方式 N4(b)）＋自有資本支出")
    v05_pay = {"azure": 5, "oracle": 6, "aws": 7, "trn": 8, "cw": 9, "cer": 10}
    for key, zh, T, s, e, gw, kind, note in P3["contracts_full"]:
        inn = lambda c, s=s, e=e: f"({c}$6>={s})*({c}$6<={e})"  # noqa: E731
        n = f"({e}-{s}+1)"
        if kind == "flat":
            f = lambda i, c, T=T, n=n, inn=inn: f"={inn(c)}*{T}/{n}"  # noqa: E731
            nt = "揭露年額＝總額 ÷ 年數（E8f；v0.5 揭露年額欄 60），期間內固定"
        else:
            f = lambda i, c, T=T, s=s, n=n, inn=inn: f'={inn(c)}*IF({prof}="even",{T}/{n},{T}*({c}$6-{s}+1)/({n}*({n}+1)/{k2}))'  # noqa: E731
            nt = "總額依攤提方式（Inputs：ramp＝期間內線性爬升，權重＝年序 ÷ n(n+1)/2；even＝平均）"
        K.add(f"合約實付：{zh}", "$B", f, v05.row("支出", v05_pay[key]), nt + "；與 Compute 供給同一時程", name=f"COST_Pay_{SUP_NAME[key]}", key=f"pay_{key}")
    keys = [k for k, *_ in P3["contracts_full"]]
    K.add("合約實付合計", "$B", lambda i, c: "=" + "+".join(f"{c}{K.rows['pay_' + k]}" for k in keys), v05.row("支出", 11), "", name="COST_ContractPay", key="pay")
    K.add("自有資本支出（毛額）", "$B", lambda i, c: f"={cmp_('own_capex', c)}", v05.row("支出", 12), "Inputs（v0.5 ownedCapex，Assumed；Compute 第十三節）", key="own_g")
    K.add("合作方出資占比", "比例", lambda i, c: f"={I('合作方出資占比')}", None, "Inputs（v0.5 S3；0，區間 0–0.3）", key="own_ps")
    K.add("自有資本支出（扣合作方出資）", "$B", lambda i, c: f"={c}{K.rows['own_g']}*(1-{c}{K.rows['own_ps']})", v05.row("支出", 12),
          "v0.5 支出第 12 列", name="COST_OwnedCapex", key="own")
    K.add("算力成本（現金口徑）＝合約實付＋自有資本支出", "$B", lambda i, c: f"={c}{K.rows['pay']}+{c}{K.rows['own']}", v05.row("支出", 14),
          "r6 P4-1（V3）；命題輸出基準", name="COST_Compute", key="cc")
    K.add("對照：2025 實際算力支出＝推論 8.4＋訓練（10.59＋其他雲端）", "$B",
          only25(f"={S('推論成本合計：2025')}+{S('研發費用中付 Microsoft 部分：2025')}+{I('2025 其他雲端')}"), None,
          "SRC_OAI＋Inputs（V6）；v0.5 算力損益成本 2025＝20.4", name="COST_Actual2025", name_col="D", key="act25")
    K.add("差距：2025 算力成本（合約實付）− 實際算力支出", "$B", only25(f"=D{K.rows['cc']}-D{K.rows['act25']}"), None,
          "只列差距（2025 合約推得供給 1.23 GW 對支出換算 1.70 GW；S3 已註）", key="gap25")
    K.add("每 GW 合約實付＝合約實付 ÷ 合約供給 GW", "$B/GW/年", lambda i, c: f"={c}{K.rows['pay']}/{cmp_('sup_con', c)}", None,
          "未揭露 GW 的三家恆等於合約價（Checks）；Oracle 13.3、Trainium／Cerebras 依爬升", key="pay_gw")

    # ═════ 三、供應商持有成本與雲端毛利 ═════
    K.section("三、供應商持有成本與雲端毛利（V3：持有成本＝合約 GW × TK IF_HoldEcon 世代加權；只作參考；差額＝雲端毛利）")
    for tag, zh in (("base", "基準"), ("lo", "低成本"), ("hi", "高成本")):
        K.add(f"世代加權持有成本（{zh}）", "$B/GW/年",
              lambda i, c, tag=tag: "=" + "+".join(f"{cmp_('s_' + g, c)}*${GC[g]}${K.rows['hold_' + tag]}" for g in GEN_KEYS), None,
              "Σ 世代占比（Compute 第二節）× 各世代每 GW 年持有成本", name="COST_HoldW" if tag == "base" else None, key=f"hw_{tag}")
    K.add("合約供給 GW（不含自建）", "GW", lambda i, c: f"={cmp_('sup_con', c)}", None, "Compute 第十三節（S3 的合約供給）", key="gw_con")
    K.add("供應商持有成本（基準）", "$B", lambda i, c: f"={c}{K.rows['gw_con']}*{c}{K.rows['hw_base']}", None, "", name="COST_SupplierHold", key="sh")
    K.add("雲端毛利＝合約實付 − 供應商持有成本", "$B", lambda i, c: f"={c}{K.rows['pay']}-{c}{K.rows['sh']}", None, "日後接 CRWV 模型（r6 V3）", name="COST_CloudGM", key="gm")
    K.add("雲端毛利率", "比例", lambda i, c: f"={c}{K.rows['gm']}/{c}{K.rows['pay']}", None, "", name="COST_CloudGMPct", key="gmp")
    for tag, zh in (("lo", "低成本"), ("hi", "高成本")):
        K.add(f"雲端毛利率（持有成本{zh}）", "比例", lambda i, c, tag=tag: f"=({c}{K.rows['pay']}-{c}{K.rows['gw_con']}*{c}{K.rows['hw_' + tag]})/{c}{K.rows['pay']}",
              None, "TK 成本角落情境", key=f"gmp_{tag}")

    # ═════ 四、經濟口徑 ═════
    K.section("四、經濟口徑（S4 預設：自有資本支出改以 TK 持有成本世代加權 × 自建 GW 年化）")
    K.add("自建 GW", "GW", lambda i, c: f"={cmp_('own', c)}", None, "Compute 第十三節", key="own_gw")
    K.add("自建持有成本（經濟）", "$B", lambda i, c: f"={c}{K.rows['own_gw']}*{c}{K.rows['hw_base']}", None, "", name="COST_OwnedHold", key="own_hold")
    K.add("算力成本（經濟口徑）＝合約實付＋自建持有成本", "$B", lambda i, c: f"={c}{K.rows['pay']}+{c}{K.rows['own_hold']}", v05.row("算力MW", 41),
          "v0.5 對照欄＝算力損益成本（2025＝實際）", name="COST_ComputeEcon", key="ce")
    K.add("現金口徑 − 經濟口徑（資本支出與持有成本的時間差）", "$B", lambda i, c: f"={c}{K.rows['cc']}-{c}{K.rows['ce']}", None, "", key="cash_econ")

    # ═════ 五、非算力成本（V4） ═════
    K.section("五、非算力成本（V4）：2025＝財報（研發總額 − 訓練支出；銷售、管理）；2026 起＝人數 × 每人年成本，股權報酬單列")
    K.add("研發費用總額（SRC）", "$B", only25(f"={S('研發費用總額：2025')}"), None, "外流財報", key="rd_tot")
    K.add("研發算力（訓練支出）＝10.59＋其他雲端", "$B", only25(f"={S('研發費用中付 Microsoft 部分：2025')}+{I('2025 其他雲端')}"), None, "V6（區間 10.59–13.50）", key="rd_cmp25")
    K.add("非算力研發（財報口徑）＝研發總額 − 訓練支出", "$B", only25(f"=D{K.rows['rd_tot']}-D{K.rows['rd_cmp25']}"), None,
          "r6 P4-2：≈7.19（區間 5.68–8.59）", name="COST_NonCompRD2025", name_col="D", key="ncrd25")
    K.add("銷售費用（SRC）", "$B", only25(f"={S('銷售費用（sales and marketing）：FY2025')}"), None, "外流財報（v0.5 opexExComputePctRevenue 2025 的拆分來源）", key="sm25")
    K.add("管理費用（SRC）", "$B", only25(f"={S('管理費用（general and administrative）：FY2025')}"), None, "同上", key="ga25")
    K.add("非算力營運費用（財報口徑，含股權報酬）", "$B", only25(f"=D{K.rows['ncrd25']}+D{K.rows['sm25']}+D{K.rows['ga25']}"), None,
          "S4 預設：財報費用列含股權報酬（GAAP 常見口徑）", key="nc25")
    for k_, zh in (("ncrd25", "研發"), ("sm25", "銷售"), ("ga25", "管理")):
        K.add(f"功能別占比（2025）：{zh}", "比例", only25(f"=D{K.rows[k_]}/D{K.rows['nc25']}"), None, "2026 起沿用（S4 預設）", key=f"sh_{k_}")
    hc_g = [I("員工人數年增率", str(y)) for y in YEARS[1:]]
    K.add("員工人數（年均）", "人", lambda i, c: f"={S('員工人數：2025（約）')}" if i == 0 else f"={YC[i - 1]}{K.r}*(1+{hc_g[i - 1]})", None,
          "2025＝SRC（WSJ 約 4,000）；2026 起＝前一年 ×（1＋年增率，Inputs）", name="COST_Headcount", key="hc")
    K.add("每人股權報酬", "$M/人/年", lambda i, c: f"={S('平均每人股權報酬：2025')}" if i == 0 else f"={YC[i - 1]}{K.r}*(1+{I('每人股權報酬年變動率')})", None,
          "2025＝SRC（WSJ 約 $1.5M）；2026 起 Inputs 年變動率", key="sbc_pp")
    K.add("股權報酬（單列）", "$B", lambda i, c: f"={c}{K.rows['hc']}*{c}{K.rows['sbc_pp']}/{m2b}", None, "人數 × 每人股權報酬（r6 P4-2：單列一列）", name="COST_SBC", key="sbc")
    r_pp = K.r + 1
    K.add("非算力營運費用（不含股權報酬）", "$B",
          lambda i, c: f"=D{K.rows['nc25']}-D{K.rows['sbc']}" if i == 0 else f"={c}{K.rows['hc']}*{c}{r_pp}/{m2b}", v05.row("支出", 23),
          "2025＝財報 − 股權報酬；2026 起＝人數 × 每人年成本（V4）。v0.5 對照欄＝營收占比路徑（v0.5 註：不含股票薪酬）", name="COST_NonCompExSBC", key="ncx")
    K.add("每人年成本（不含股權報酬；含現金薪酬與其他）", "$M/人/年",
          lambda i, c: f"=D{K.rows['ncx']}*{m2b}/D{K.rows['hc']}" if i == 0 else f"={YC[i - 1]}{K.r}*(1+{I('每人年成本（不含股權報酬）年變動率')})", None,
          "2025＝Derived（財報 ÷ 人數）；2026 起 Inputs 年變動率", key="pp")
    assert K.rows["pp"] == r_pp
    for k_, zh, nmx in (("ncrd25", "非算力研發", "COST_NonCompRD"), ("sm25", "銷售費用", None), ("ga25", "管理費用", None)):
        K.add(f"{zh}（不含股權報酬）", "$B", lambda i, c, k_=k_: f"={c}{K.rows['ncx']}*$D${K.rows['sh_' + k_]}", None, "非算力營運費用（不含股權報酬）× 2025 功能別占比",
              name=nmx, key=f"x_{k_}")
    K.add("銷售＋管理（不含股權報酬）", "$B", lambda i, c: f"={c}{K.rows['x_sm25']}+{c}{K.rows['x_ga25']}", None, "", name="COST_SGA", key="sga")
    K.add("非算力成本（含股權報酬）", "$B", lambda i, c: f"={c}{K.rows['ncx']}+{c}{K.rows['sbc']}", None, "", name="COST_NonCompInclSBC", key="nci")
    v05pct = [None] + [I("非算力營運費用占營收比（v0.5）", str(y)) for y in YEARS[1:]]
    K.add("暫代對照：v0.5 營收占比路徑（研發不含付 Microsoft＋銷售＋管理）÷ 總額營收", "比例",
          lambda i, c: (f"=({S('研發費用總額：2025')}-{S('研發費用中付 Microsoft 部分：2025')}+{S('銷售費用（sales and marketing）：FY2025')}"
                        f"+{S('管理費用（general and administrative）：FY2025')})/{S('營收：FY2025')}") if i == 0 else f"={v05pct[i]}",
          [1.216, 0.33, 0.265, 0.22, 0.18, 0.14], "2025 由財報推得（v0.5 1.216）；2026 起 Inputs INP_196–200（v0.5 S5a；P4 取代為人數法，只作對照）", key="v05pct")
    K.add("暫代對照：v0.5 路徑金額＝占比 × 總額營收（截頂後）", "$B", lambda i, c: f"={c}{K.rows['v05pct']}*Revenue!{c}{vr['gross_c']}", v05.row("支出", 23), "", key="v05amt")
    K.add("差距：人數法（含股權報酬）− v0.5 路徑", "$B", lambda i, c: f"={c}{K.rows['nci']}-{c}{K.rows['v05amt']}", None,
          "v0.5 路徑 2025 含股權報酬（財報口徑）、2026 起註明不含；兩者口徑差異列於報告", key="v05gap")
    K.add("對照：WSJ 股權報酬路徑（2025＝占營收 46.2%；之後每年 +$3B）", "$B",
          lambda i, c: f"={S('股權報酬占營收比：2025（投資人預測）')}*{S('營收：FY2025')}" if i == 0 else f"={YC[i - 1]}{K.r}+{S('股權報酬年增額：至 2030（投資人預測）')}", None,
          "公司預測，只作對照（不回饋基準）", key="wsj")
    K.add("差距：本模型股權報酬 − WSJ 路徑", "$B", lambda i, c: f"={c}{K.rows['sbc']}-{c}{K.rows['wsj']}", None, "", key="wsj_gap")
    K.add("對照：FT 人數隱含 2026 年均成長＝AVERAGE（2026-03, 年底目標）÷ 2025 − 1", "比例",
          [None, f"=(AVERAGE({S('員工人數：2026-03（約）')},{S('員工人數目標：2026 年底（約）')})-{S('員工人數：2025（約）')})/{S('員工人數：2025（約）')}"] + [None] * 4,
          None, "對照 Inputs 2026 年增率（0.55）", key="ft_g")

    # ═════ 六、Q3 ═════
    K.section("六、Q3（研發占整體成本）＝（研發算力＋非算力研發）÷ 全部營運成本；研發算力對 TK L1_Ans3 的差距（$B/年）")
    sup = lambda c: cmp_("sup", c)  # noqa: E731
    K.add("研發算力", "$B", lambda i, c: f"=D{K.rows['rd_cmp25']}" if i == 0 else f"={c}{K.rows['cc']}*{cmp_('rd', c)}/{sup(c)}", v05.row("算力MW", 45),
          "2025＝訓練支出公式（r6 P4-2）；2026 起＝算力成本 × 研發 GW ÷ 供給 GW（v0.5：訓練成本＝訓練 GW × 每 GW 成本）", name="COST_RDCompute", key="q_rd")
    K.add("推論算力", "$B", lambda i, c: f"={S('推論成本合計：2025')}" if i == 0 else f"={c}{K.rows['cc']}*{cmp_('infcap', c)}/{sup(c)}", v05.row("算力MW", 44),
          "2025＝SRC；2026 起＝算力成本 × 推論 GW（截頂後）÷ 供給 GW", key="q_inf")
    r_q = K.r + 1
    K.add("閒置算力", "$B", lambda i, c: f"=D{r_q}-D{K.rows['q_rd']}-D{K.rows['q_inf']}" if i == 0 else f"={c}{K.rows['cc']}*{cmp_('idle', c)}/{sup(c)}",
          v05.row("算力MW", 46), "2026 起＝算力成本 × 閒置 GW ÷ 供給 GW", key="q_idle")
    K.add("算力成本（Q3 口徑）", "$B", lambda i, c: f"=D{K.rows['act25']}" if i == 0 else f"={c}{K.rows['cc']}", None,
          "2025＝實際算力支出（校準年）；2026 起＝算力成本（現金口徑）", key="q_cc")
    assert K.rows["q_cc"] == r_q
    K.add("研發合計＝研發算力＋非算力研發（不含股權報酬）", "$B", lambda i, c: f"={c}{K.rows['q_rd']}+{c}{K.rows['x_ncrd25']}", None, "", key="q_rdt")
    K.add("全部營運成本＝算力（Q3 口徑）＋非算力（不含股權報酬）＋股權報酬", "$B",
          lambda i, c: f"={c}{K.rows['q_cc']}+{c}{K.rows['ncx']}+{c}{K.rows['sbc']}", None, "", key="q_tot")
    K.add("Q3＝研發合計 ÷ 全部營運成本", "比例", lambda i, c: f"={c}{K.rows['q_rdt']}/{c}{K.rows['q_tot']}", None, "r6 P4-2", name="COST_Q3", key="q3")
    K.add("研發算力占研發合計", "比例", lambda i, c: f"={c}{K.rows['q_rd']}/{c}{K.rows['q_rdt']}", None, "", key="q_rdshare")
    K.add("TK L1_Ans3（研發算力成本；2025 穩態）", "$B/年", lambda i, c: "=TK_L1_Ans3", None, "Tokenomics 問 3（算力部分）；區間 TK_L1_Ans3_Lo–Hi", key="q_tk")
    K.add("差距：研發算力 − TK L1_Ans3", "$B/年", lambda i, c: f"={c}{K.rows['q_rd']}-{c}{K.rows['q_tk']}", None, "r6 P4-2：只列差距", name="COST_Q3GapTK", key="q_gap")
    K.add("倍數：研發算力 ÷ TK L1_Ans3", "倍", lambda i, c: f"={c}{K.rows['q_rd']}/{c}{K.rows['q_tk']}", None, "交接第 8y 節 W3：Tokenomics 物理下限約小一個數量級", key="q_mult")

    # ═════ 七、命題輸出 ═════
    K.section("七、命題輸出（每 VR 等值 GW；分母＝供給 VR 等值 GW，含自建）：營收、算力成本、非算力成本、全成本、差額；現金口徑為基準，經濟口徑並列")
    K.add("供給 VR 等值 GW（分母）", "GW", lambda i, c: f"={cmp_('vr_sup', c)}", None, "CMP_Supply_VReq（S4 起含自建 GW）", key="p_den")
    K.add("營收（截頂後淨額）", "$B", lambda i, c: f"=Revenue!{c}{vr['net_c']}", None, "REV_NetCapped（扣 Microsoft 分成；估值與命題用）", key="p_rev")
    K.add("全成本（含股權報酬；現金口徑）", "$B", lambda i, c: f"={c}{K.rows['cc']}+{c}{K.rows['ncx']}+{c}{K.rows['sbc']}", None,
          "S4 預設：命題的全成本含股權報酬（較保守）", name="COST_FullCash", key="p_full")
    K.add("全成本（不含股權報酬；現金口徑）", "$B", lambda i, c: f"={c}{K.rows['cc']}+{c}{K.rows['ncx']}", None, "", key="p_fullx")
    K.add("差額＝營收 − 全成本（含股權報酬；現金口徑）", "$B", lambda i, c: f"={c}{K.rows['p_rev']}-{c}{K.rows['p_full']}", None, "負＝缺口（P5 融資）",
          name="COST_GapCash", key="p_gap_b")
    K.add("每 VR 等值 GW：營收", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_rev']}/{c}{K.rows['p_den']}", v05.row("算力MW", 51),
          "v0.5 對照欄＝總額營收 ÷ VR-eq GW（v0.5 口徑）", name="COST_PropRev_VR", key="p_rev_vr")
    K.add("每 VR 等值 GW：算力成本（現金）", "$B/GW/年", lambda i, c: f"={c}{K.rows['cc']}/{c}{K.rows['p_den']}", v05.row("算力MW", 42),
          "v0.5 對照欄＝算力損益成本 ÷ VR-eq GW", name="COST_PropCompute_VR", key="p_cc_vr")
    K.add("每 VR 等值 GW：非算力成本（不含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['ncx']}/{c}{K.rows['p_den']}", None, "", key="p_nc_vr")
    K.add("每 VR 等值 GW：股權報酬", "$B/GW/年", lambda i, c: f"={c}{K.rows['sbc']}/{c}{K.rows['p_den']}", None, "", key="p_sbc_vr")
    K.add("每 VR 等值 GW：全成本（含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_cc_vr']}+{c}{K.rows['p_nc_vr']}+{c}{K.rows['p_sbc_vr']}", None, "",
          name="COST_PropFull_VR", key="p_full_vr")
    K.add("每 VR 等值 GW：全成本（不含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_cc_vr']}+{c}{K.rows['p_nc_vr']}", None, "", key="p_fullx_vr")
    K.add("每 VR 等值 GW：差額（含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_rev_vr']}-{c}{K.rows['p_full_vr']}", None, "命題：負＝營收不足以覆蓋全成本",
          name="COST_PropGap_VR", key="p_gap_vr")
    K.add("每 VR 等值 GW：差額（不含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_rev_vr']}-{c}{K.rows['p_fullx_vr']}", None, "", key="p_gapx_vr")
    K.add("覆蓋率＝營收 ÷ 全成本（含股權報酬）", "倍", lambda i, c: f"={c}{K.rows['p_rev']}/{c}{K.rows['p_full']}", None, "", name="COST_Coverage", key="p_cov")
    K.add("經濟口徑：每 VR 等值 GW 算力成本", "$B/GW/年", lambda i, c: f"={c}{K.rows['ce']}/{c}{K.rows['p_den']}", None, "合約實付＋自建持有成本", key="e_cc_vr")
    K.add("經濟口徑：每 VR 等值 GW 全成本（含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['e_cc_vr']}+{c}{K.rows['p_nc_vr']}+{c}{K.rows['p_sbc_vr']}", None, "",
          key="e_full_vr")
    K.add("經濟口徑：每 VR 等值 GW 差額（含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_rev_vr']}-{c}{K.rows['e_full_vr']}", None, "",
          name="COST_PropGapEcon_VR", key="e_gap_vr")
    K.add("經濟口徑：每 VR 等值 GW 差額（不含股權報酬）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_rev_vr']}-{c}{K.rows['e_cc_vr']}-{c}{K.rows['p_nc_vr']}", None, "", key="e_gapx_vr")
    K.add("經濟口徑：差額（含股權報酬）", "$B", lambda i, c: f"={c}{K.rows['p_rev']}-{c}{K.rows['ce']}-{c}{K.rows['ncx']}-{c}{K.rows['sbc']}", None, "",
          name="COST_GapEcon", key="e_gap_b")
    K.add("對照：2025 以實際算力支出（20.4）計的每 VR 等值 GW 差額（含股權報酬）", "$B/GW/年",
          only25(f"=(D{K.rows['p_rev']}-D{K.rows['act25']}-D{K.rows['ncx']}-D{K.rows['sbc']})/D{K.rows['p_den']}"), None, "只並列（基準依 r6 用合約實付）", key="p_gap25act")
    vr_lab, vr_cost = f"Compute!$G${cr['g_lab']}", f"Compute!$G${cr['g_cost']}"
    for nm_, zh, key in (("IF_FullCostDefault_Sol", "下游預設 IF_FullCostDefault_Sol；基準", "tk_def"), ("IF_FullCost_Sol", "IF_FullCost_Sol；並列", "tk_full")):
        K.add(f"TK 參考：每 VR200 GW 年全成本（{zh}）", "$B/GW/年",
              lambda i, c, nm_=nm_: f"=SUMIFS(TK_{nm_},TK_HdrGen,{vr_lab},TK_HdrCost,{vr_cost})*SUMIFS(TK_IF_TokGW_Sol,TK_HdrGen,{vr_lab},TK_HdrCost,{vr_cost})*TK_IF_Util/{d2b}",
              None, "TK $/M（VR200 Sol 基準欄）× 每 GW 年產出（Sol）× 基準利用率 ÷ 10⁹；只作參考（S4 預設），算力成本仍＝合約實付", key=key)
    K.add("差距：每 VR 等值 GW 算力成本（現金）− TK 參考（下游預設）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_cc_vr']}-{c}{K.rows['tk_def']}", None, "", key="tk_gap")
    K.add("差距：每 VR 等值 GW 算力成本（現金）− TK 參考（IF_FullCost）", "$B/GW/年", lambda i, c: f"={c}{K.rows['p_cc_vr']}-{c}{K.rows['tk_full']}", None, "", key="tk_gap2")

    # ═════ 八、V16 Azure 攤入檢查 ═════
    K.section("八、V16 Azure 攤入檢查（r6 P4-5；只列差距、立旗標，不校準；基準起點 2025）")
    az = next(t for t in P3["contracts_full"] if t[0] == "azure")
    _, _, T, s, e, *_ = az
    K.add("2025 Azure $250B 攤入額（基準起點 2025）", "$B", only25(f"=D{K.rows['pay_azure']}"), None, "合約實付：Microsoft Azure 2025 欄", key="v_az")
    K.add("2025 實際算力支出（推論 8.4＋訓練支出公式）", "$B", only25(f"=D{K.rows['act25']}"), None, "", key="v_act")
    K.add("差距：攤入額 − 實際算力支出", "$B", only25(f"=D{K.rows['v_az']}-D{K.rows['v_act']}"), None, "", key="v_gap")
    K.add("旗標（1＝攤入額超過實際算力支出）", "旗標", only25(f"=(D{K.rows['v_az']}>D{K.rows['v_act']})*(D$6<={cal})"), None, "S4 預設：只立旗標", name="COST_V16Flag", name_col="D", key="v_flag")
    s26 = f"Inputs!$G${inp_row(s)}"
    n26 = f"({e}-{s26}+1)"
    K.add("起點 2026 情境：Azure 攤入額", "$B",
          lambda i, c: f'=({c}$6>={s26})*({c}$6<={e})*IF({prof}="even",{T}/{n26},{T}*({c}$6-{s26}+1)/({n26}*({n26}+1)/{k2}))', None,
          "Azure 起點取 Inputs 區間高端（V16：2025–2026）；終點不變", key="v_az26")
    K.add("差異：起點 2026 − 基準", "$B", lambda i, c: f"={c}{K.rows['v_az26']}-{c}{K.rows['pay_azure']}", None, "", key="v_diff")

    # ═════ 九、檢查列（供 Checks 取各年最大絕對值） ═════
    K.section("九、檢查列（各年應為 0；Checks 取最大絕對值）")
    K.add("逐合約實付加總 − 合約實付合計", "$B", lambda i, c: "=" + "+".join(f"{c}{K.rows['pay_' + k]}" for k in keys) + f"-{c}{K.rows['pay']}", None, "", key="ck_pay")
    K.add("未揭露 GW 合約：|實付 − 供給 GW × 合約價| 合計", "$B",
          lambda i, c: "=" + "+".join(f"ABS({c}{K.rows['pay_' + k]}-{cmp_('sup_' + k, c)}*{price})" for k in ("azure", "aws", "cw")), None,
          "S3：未揭露 GW＝付款 ÷ 合約價（Azure、AWS Nvidia、CoreWeave）", key="ck_und")
    K.add("算力成本（現金）−（合約實付＋自有資本支出）", "$B", lambda i, c: f"={c}{K.rows['cc']}-{c}{K.rows['pay']}-{c}{K.rows['own']}", None, "", key="ck_cc")
    K.add("推論＋研發＋閒置算力 − 算力成本（Q3 口徑）", "$B",
          lambda i, c: f"={c}{K.rows['q_inf']}+{c}{K.rows['q_rd']}+{c}{K.rows['q_idle']}-{c}{K.rows['q_cc']}", None, "v0.5 算力MW 第 47 列", key="ck_split")
    K.add("非算力研發＋銷售＋管理（不含股權報酬）− 合計", "$B",
          lambda i, c: f"={c}{K.rows['x_ncrd25']}+{c}{K.rows['x_sm25']}+{c}{K.rows['x_ga25']}-{c}{K.rows['ncx']}", None, "", key="ck_nc")

    for c, w in (("A", 7), ("B", 62), ("C", 12)):
        K.ws.column_dimensions[c].width = w
    for c in YC:
        K.ws.column_dimensions[c].width = 13
    K.ws.column_dimensions["Q"].width = 80
    K.ws.freeze_panes = "D7"
    return dict(K=K)
