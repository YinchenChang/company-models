"""智譜 v0.1-Z4：融資頁 Funding 與反向模式頁 Reverse，以及 Z4 的 Checks 列。

版面同 Z2／Z3：A 編號、B 項目、C 單位、D–I 2025–2030、J 空白、K 1H2026、L 2H2026、M 說明。
- 流量列（RMB 億）：2026（E）＝1H26（實際）＋2H26（驅動）。現金存量列：K＝2026-06-30 實際（SRC），L 與 E＝2026 年底。
- 外部資金需求只在年底（E–I）判斷：年底現金低於最低現金的差額（補足至最低現金），累計＝命題 2（同 OpenAI v0.6 p5.py）。
機制（規格 D16–D22、r1 的 D16r、D21r、D22r；工作單 Z4 與 chat 端補充）：
  自由現金流＝營收淨額（截頂後）− 算力成本 − 本地化交付成本 − 非算力成本（不含股權報酬；股權報酬非現金，加回）＋其他收益 − 利息
  → 減非算力資本支出、併購 → 來源順序：①期初現金 ②已到位股權（IPO、2026-07、2026-09 配售淨額，港幣 × Inputs 匯率）
  ③已提用債務（銀行借款續借；可換股債券淨額流入、2027-09 到期償還）④外部資金需求（補足至最低現金）。
  Reverse：①管理層量化目標不存在（ARR 只對照）②分析師共識營收 → 所需倍數、反向外部資金（同一支出）③市值隱含營收（EV ÷ 同業倍數）。
  Reverse 只被 Checks 引用（不回饋基準；test_reverse_not_fed_back）。
所有數值引用 SRC_ZP／Inputs 或同簿儲存格（E6）；Python 只組公式文字。同頁後建列以「«F.鍵»」「«X.鍵»」佔位，由 z2.fill() 代換。
"""
from __future__ import annotations

from z2 import ALL, HC, LATE, PREV, YC, Sheet

FLOW = ("D", "K", "L", "E") + tuple(LATE)
YE = ("E",) + tuple(LATE)                 # 年底判斷外部資金需求的欄（2026–2030）
NEXT = {"E": "F", "F": "G", "G": "H", "H": "I", "I": "I"}


def build(wb, D, Z):
    S, I = D.S, D.I
    halves, pct = I("halves_year"), I("pct")
    dy, d2 = I("days_year"), I("days_2h26")
    fxu, fxh, bn, mn = I("fx_usdcny"), I("fx_hkdcny"), I("mul_bn_yi"), I("div_mn_yi")
    ycol = lambda c: f"{c}$6"  # noqa: E731

    # ═══════════════════════════ Funding ═══════════════════════════
    F = Sheet(wb, "Funding", "F",
              "Funding — 融資：自由現金流、投資活動（非算力資本支出、併購）、已到位融資（IPO、2026-07／09 配售、可換股債券、銀行借款）、最低現金與外部資金需求、可換股債券轉股情境、或有、配售用途對照、回流對照、2025／1H26 現金對帳",
              "規格 D16–D20、D16r；工作單 Z4。自由現金流＝營收淨額（截頂後）− 算力成本 − 本地化部署交付成本 − 非算力成本（不含股權報酬；股權報酬非現金，加回）＋其他收益（D20：2025、1H26 實際，之後 Inputs，基準 0）− 利息（2025、1H26＝財報財務成本；之後銀行借款餘額 × 利率）。"
              "來源順序：①期初現金 ②已到位股權（港幣淨額 × HKD/CNY；入帳時點＝完成日）③已提用債務（銀行借款到期續借；可換股債券淨額 2H26 流入、2027-09 到期現金償還，D16r）④外部資金需求（年底現金補足至最低現金）。",
              "金額 RMB 億。K＝1H26（實際；期末現金＝中期報告 39.94 億，IPO 淨額已全數動用並已反映）；L＝2H26（驅動）；E＝2026 全年（流量＝K＋L；現金＝年底）。外部資金需求只在年底（E–I）判斷。")
    F.years(D)
    cost = lambda key, c: f"Cost!{c}«K.{key}»"  # noqa: E731

    F.section("一、自由現金流（Cost COST_GapCash 為起點：營收 − 全成本；股權報酬非現金，加回；加其他收益、減利息）")

    def rev(c):
        if c in YC and c != "E":
            return f"=Revenue!{c}«V.net_c»"
        if c == "K":
            return "=Revenue!K«V.gross»"
        if c == "L":
            return "=Revenue!L«V.cloud»*Compute!L«C.cap»+Revenue!L«V.onprem»+Revenue!L«V.ads»"
        return "=K«F.rev»+L«F.rev»"
    F.add("營收淨額（截頂後）", "RMB 億", rev, "REV_NetCapped；1H26＝實際（不截頂）；2H26＝雲端 × 容量上限係數＋本地化＋廣告；2026＝1H＋2H", key="rev")
    F.add("算力成本（現金口徑）", "RMB 億", {c: f"={cost('cc', c)}" for c in ALL}, "COST_Compute", key="cc")
    F.add("本地化部署交付成本", "RMB 億", {c: f"={cost('opd', c)}" for c in ALL}, "COST_OnPremDelivery", key="opd")
    F.add("非算力成本（不含股權報酬）", "RMB 億", {c: f"={cost('nc_cal_m', c)}" for c in ALL}, "COST_NonCompExSBC", key="ncx")
    F.add("股權報酬（非現金；只列示，現金流加回）", "RMB 億", {c: f"={cost('x_sbc_m', c)}" for c in ALL}, "COST_SBC", key="sbc")
    F.add("差額（含股權報酬；命題口徑）＝營收 − 全成本", "RMB 億", {c: f"={c}«F.rev»-{c}«F.cc»-{c}«F.opd»-{c}«F.ncx»-{c}«F.sbc»" for c in ALL},
          "年度欄＝COST_GapCash（Checks 對帳）", key="gap", name="FND_Gap")

    def oth(c):
        if c == "D":
            return f"={S('other_inc_fy25')}"
        if c == "K":
            return f"={S('other_inc_1h26')}"
        if c == "L":
            return f"={I('other_inc_future')}*{d2}/{dy}"
        if c == "E":
            return "=K«F.oth»+L«F.oth»"
        return f"={I('other_inc_future')}"
    F.add("其他收益（其他收入：政府補助、利息收入等）", "RMB 億", oth,
          "D20：2025（SRC_ZP_143）、1H26（SRC_ZP_203，主要為上市後存款利息）實際；2H26 起 Inputs（基準 0，Decision）", key="oth")

    def intr(c):
        if c == "D":
            return f"={S('fin_cost_fy25')}"
        if c == "K":
            return f"={S('fin_cost_1h26')}"
        if c == "L":
            return f"=K«F.loan_end»*{I('bank_rate')}*{d2}/{dy}"
        if c == "E":
            return "=K«F.int»+L«F.int»"
        return f"={PREV[c]}«F.loan_end»*{I('bank_rate')}"
    F.add("利息支出（非算力現金成本）", "RMB 億", intr,
          "工作單 Z4：2025、1H26＝財報財務成本（SRC_ZP_144、204；含匯兌損失與租賃利息）；之後＝期初銀行借款餘額 × 利率（Inputs，利率未揭露 → LPR Analogy）；可換股債券零息", key="int", name="FND_Interest")
    F.add("自由現金流＝差額＋股權報酬＋其他收益 − 利息", "RMB 億", {c: f"={c}«F.gap»+{c}«F.sbc»+{c}«F.oth»-{c}«F.int»" for c in ALL},
          "＝營收 − 算力成本 − 本地化交付成本 − 非算力成本（不含股權報酬）＋其他收益 − 利息（Checks 恆等式）", key="fcf", name="FND_FCF", name_h="FND_FCF_H")

    F.section("二、投資活動（D18）：非算力資本支出、併購現金支出（自有算力資本支出已在算力成本內，D11r 基準 0）")

    def capex(c):
        if c == "D":
            return f"={S('capex_fy25')}"
        if c == "K":
            return f"={S('capex_1h26')}"
        if c == "L":
            return f"={I('capex_nc')}*{d2}/{dy}"
        if c == "E":
            return "=K«F.capex»+L«F.capex»"
        return f"={I('capex_nc')}"
    F.add("非算力資本支出", "RMB 億", capex, "2025、1H26＝財報公告資本開支（SRC_ZP_149、209；1H26 主要為北京紅鑽物業；r2 P2：公告口徑可能含使用權資產添置，FY2025、1H26 現金口徑未揭露，見第十一節）；2H26 起 Inputs（FY2025 水準延續）", key="capex")
    F.add("對照：北京紅鑽 100% 總代價上限（已含於 1H26 資本開支，只列）", "RMB 億", {"K": f"={S('ma_hongzuan_consideration')}/{mn}"},
          "SRC_ZP_555（股權現金 0.82 億＋承接債務約 2.79 億；海淀辦公樓，自用）", key="ma_hz")
    F.add("併購現金支出：中科加禾約 60%", "RMB 億", {"L": f"={I('zkjh_cash')}", "E": "=L«F.ma»"},
          "D18；對價未揭露（『數億元』，SRC_ZP_559）→ Inputs；被併公司營收與成本不併入驅動", key="ma", name="FND_MA")

    def netop(c):
        if c in ("L", "E"):
            return f"={c}«F.fcf»-{c}«F.capex»-{c}«F.ma»"
        return f"={c}«F.fcf»-{c}«F.capex»"
    F.add("融資前淨現金流＝自由現金流 − 資本支出 − 併購", "RMB 億", netop, "", key="netop", name="FND_NetOp")

    F.section("三、已到位融資（來源順序 ②③；港幣、美元淨額 × Inputs 匯率；入帳時點＝完成日）")
    F.add("2026-01 IPO 淨額（含超額配股；1H26 已全數動用）", "RMB 億", {"K": f"={S('ipo_net_incl_overallot')}*{fxh}/{mn}", "E": "=K«F.ipo»"},
          "SRC_ZP_522 HK$4,896.2 百萬 × HKD/CNY ÷ 100；已反映於 2026-06-30 現金", key="ipo")
    F.add("2026-07 H 股配售淨額（2026-07-13 完成）", "RMB 億", {"L": f"={S('placing1_net')}*{fxh}/{mn}", "E": "=L«F.pl1»"},
          "SRC_ZP_536 HK$31,374.95 百萬 × HKD/CNY ÷ 100", key="pl1")
    F.add("2026-09 H 股配售淨額（2026-09-16 完成）", "RMB 億", {"L": f"={S('placing2_net')}*{fxh}/{mn}", "E": "=L«F.pl2»"},
          "SRC_ZP_544 HK$15,664.13 百萬 × HKD/CNY ÷ 100", key="pl2")
    F.add("已到位股權合計", "RMB 億", {"K": "=K«F.ipo»", "L": "=L«F.pl1»+L«F.pl2»", "E": "=K«F.eq»+L«F.eq»"}, "", key="eq", name="FND_EquityIn")
    F.add("2026-09 可換股債券淨額（2026-09-18 發行；零息）", "RMB 億", {"L": f"={S('cb_net_usd')}*{fxu}/{mn}", "E": "=L«F.cb»"},
          "SRC_ZP_547 US$3,010.77 百萬 × USD/CNY ÷ 100（本金 RMB 201.4 億、發行價 100.5%）", key="cb")

    def rep(c):
        return {"F": f"={I('bank_repay_sw')}*{S('bank_loans_cur_1h26')}", "G": f"={I('bank_repay_sw')}*{S('bank_loans_ncur_1h26')}"}.get(c)
    F.add("銀行借款償還（開關＝1 時：流動部分 2027、非流動部分 2028）", "RMB 億", rep, "工作單 Z4 預設：到期續借（開關 0）；SRC_ZP_214、215", key="loan_rep")

    def loan_end(c):
        if c == "D":
            return f"={S('bank_loans_fy25')}"
        if c == "K":
            return f"={S('bank_loans_1h26')}"
        if c == "L":
            return "=K«F.loan_end»"
        if c == "E":
            return "=L«F.loan_end»"
        if c in ("F", "G"):
            return f"={PREV[c]}«F.loan_end»-{c}«F.loan_rep»"
        return f"={PREV[c]}«F.loan_end»"
    F.add("銀行借款餘額（期末）", "RMB 億", loan_end, "2025-12-31、2026-06-30＝財報（SRC_ZP_155、213；無擔保、主要人民幣、利率與到期未揭露）；之後續借", key="loan_end", name="FND_BankLoans")
    prev_loan = {"K": "D", "L": "K", "E": "D", "F": "E", "G": "F", "H": "G", "I": "H"}

    def loan_chg(c):
        if c == "D":
            return f"=D«F.loan_end»-{S('bank_loans_p_fy24')}"
        return f"={c}«F.loan_end»-{prev_loan[c]}«F.loan_end»"
    F.add("銀行借款淨增（減）", "RMB 億", loan_chg, "1H26 +15.35 已反映於 2026-06-30 現金", key="loan_chg")

    def committed(c):
        if c == "D":
            return None
        if c == "K":
            return "=K«F.eq»+K«F.loan_chg»"
        if c == "L":
            return "=L«F.eq»+L«F.cb»+L«F.loan_chg»"
        if c == "E":
            return "=K«F.committed»+L«F.committed»"
        return f"={c}«F.loan_chg»"
    F.add("已到位融資合計（股權＋可換股債券＋銀行借款淨增）", "RMB 億", committed, "來源順序 ②③", key="committed", name="FND_Committed")
    F.add("可換股債券到期償還（基準：現金償還；開關＝1 時轉股不償還）", "RMB 億",
          {c: f"={S('cb_principal')}*{I('cb_redeem')}*(1-{I('cb_conv_sw')})*({ycol(c)}={I('cb_year')})" for c in LATE},
          "D16r：本金 RMB 201.4 億（SRC_ZP_545）× 贖回價（Inputs 1）；2027-09-16 到期", key="cb_rep", name="FND_DebtRepay")

    F.section("四、現金與外部資金需求（命題 2；D17 最低現金；年底現金低於最低現金的差額＝當年外部資金需求）")
    mm, my = I("min_cash_months"), I("months_year")
    F.add("最低現金＝次年（算力成本＋非算力成本不含股權報酬）× 月數 ÷ 12", "RMB 億",
          {c: f"=({cost('cc', NEXT.get(c, 'E'))}+{cost('nc_cal_m', NEXT.get(c, 'E'))})*{mm}/{my}" for c in ("D",) + YE},
          "D17（Inputs 6 個月，3–12）；2030 以當年代替次年；2025 欄只列示", key="min", name="FND_MinCash")
    open_ = {"K": "=D«F.end»", "L": "=K«F.end»", "E": "=D«F.end»", **{c: f"={PREV[c]}«F.end»" for c in LATE}}
    F.add("期初現金（來源順序 ①）", "RMB 億", open_, "2026 期初＝2025-12-31 現金（SRC_ZP_151）", key="open", name="FND_CashOpen")
    F.add("1H26 其他現金流（對帳差＝實際期末 − 模型推算）", "RMB 億",
          {"K": f"={S('cash_end_1h26')}-(K«F.open»+K«F.netop»+K«F.committed»)", "E": "=K«F.recon»"},
          "營運資金、租賃付款、匯兌、定期存款到期等未建模項目；使 1H26 期末＝中期報告 39.94 億（SRC_ZP_211）", key="recon")

    def pre(c):
        if c == "D":
            return None
        if c == "K":
            return "=K«F.open»+K«F.netop»+K«F.committed»+K«F.recon»"
        if c == "L":
            return "=L«F.open»+L«F.netop»+L«F.committed»"
        if c == "E":
            return "=L«F.pre»"
        return f"={c}«F.open»+{c}«F.netop»+{c}«F.committed»-{c}«F.cb_rep»"
    F.add("補足前期末現金＝期初＋融資前淨現金流＋已到位融資 − 債務償還", "RMB 億", pre, "K＝2026-06-30 實際；E＝2026 年底（＝L）", key="pre")
    F.add("當年外部資金需求（來源順序 ④）＝最低現金 − 補足前（低於時；否則 0）", "RMB 億",
          {c: f"=({c}«F.min»-{c}«F.pre»)*({c}«F.min»>{c}«F.pre»)" for c in YE}, "新股權、債務或其他（不分種類；同 OpenAI v0.6）", key="ext", name="FND_ExtNeed")

    def end(c):
        if c == "D":
            return f"={S('cash_end_fy25')}"
        if c == "K":
            return "=K«F.pre»"
        if c == "L":
            return "=L«F.pre»+E«F.ext»"
        if c == "E":
            return "=L«F.end»"
        return f"={c}«F.pre»+{c}«F.ext»"
    F.add("期末現金", "RMB 億", end, "2025＝SRC_ZP_151；1H26＝SRC_ZP_211（實際）；2026 起＝補足前＋外部資金", key="end", name="FND_CashEnd", name_h="FND_CashEnd_H")
    F.add("累計外部資金需求", "RMB 億", {"E": "=E«F.ext»", **{c: f"={PREV[c]}«F.cum»+{c}«F.ext»" for c in LATE}}, "命題 2", key="cum", name="FND_ExtNeedCum")
    F.add("缺口旗標（1＝當年需要外部資金）", "旗標", {c: f"=SIGN({c}«F.ext»)" for c in YE}, "", key="flag", name="FND_GapFlag")
    F.add("缺口年度標示", "年", {c: f'=IF({c}«F.flag»,{ycol(c)},"")' for c in YE}, "", key="flag_y")
    F.add("峰值年度標示（當年外部資金需求＝2026–2030 最大且 >0）", "年",
          {c: f'=IF(AND({c}«F.flag»,{c}«F.ext»=MAX($E$«F.ext»:$I$«F.ext»)),{ycol(c)},"")' for c in YE}, "", key="peak_y")

    def noext(c):
        if c == "E":
            return "=L«F.pre»"
        if c in LATE:
            return f"={PREV[c]}«F.noext»+{c}«F.netop»+{c}«F.committed»-{c}«F.cb_rep»"
        return None
    F.add("只靠已到位融資（不另募）的年底現金", "RMB 億", noext, "已到位融資能撐到哪一年（可為負）", key="noext", name="FND_CashNoExt")
    F.add("只靠已到位融資：低於最低現金的年度標示", "年", {c: f'=IF({c}«F.noext»<{c}«F.min»,{ycol(c)},"")' for c in YE}, "", key="noext_y")
    F.add("摘要：首次缺口年", "年", {"D": f'=IF(COUNT($E$«F.flag_y»:$I$«F.flag_y»),MIN($E$«F.flag_y»:$I$«F.flag_y»),"期間內無")'}, "", key="s_first",
          name="FND_FirstGapYear", name_col="D")
    F.add("摘要：峰值年（當年外部資金需求最大）", "年", {"D": f'=IF(COUNT($E$«F.peak_y»:$I$«F.peak_y»),MIN($E$«F.peak_y»:$I$«F.peak_y»),"期間內無")'}, "",
          key="s_peak_y", name="FND_PeakYear", name_col="D")
    F.add("摘要：峰值（當年外部資金需求最大值）", "RMB 億", {"D": "=MAX($E$«F.ext»:$I$«F.ext»)"}, "", key="s_peak", name="FND_PeakExtNeed", name_col="D")
    F.add("摘要：只靠已到位融資撐不過的第一年", "年", {"D": f'=IF(COUNT($E$«F.noext_y»:$I$«F.noext_y»),MIN($E$«F.noext_y»:$I$«F.noext_y»),"期間內無")'}, "",
          key="s_noext", name="FND_NoExtFailYear", name_col="D")
    F.add("摘要：2030 年後跑道＝2030 年底現金 ÷ 2030 融資前淨現金流出", "年",
          {"D": '=IF(I«F.netop»<ABS(I«F.netop»),I«F.end»/(-I«F.netop»),"融資前現金流為正")'}, "粗估：以 2030 年流出速度計", key="s_run", name="FND_RunwayAfter2030",
          name_col="D")

    F.section("五、情境：可換股債券轉股（D16r；2027-07-01 起可換股，不需現金償還；不受 Inputs 開關影響）")

    def s_pre(c):
        if c == "E":
            return "=E«F.pre»"
        return f"={PREV[c]}«F.s_end»+{c}«F.netop»+{c}«F.committed»"
    F.add("補足前期末現金（轉股）", "RMB 億", {c: s_pre(c) for c in YE}, "", key="s_pre")
    F.add("當年外部資金需求（轉股）", "RMB 億", {c: f"=({c}«F.min»-{c}«F.s_pre»)*({c}«F.min»>{c}«F.s_pre»)" for c in YE}, "", key="s_ext")
    F.add("期末現金（轉股）", "RMB 億", {c: f"={c}«F.s_pre»+{c}«F.s_ext»" for c in YE}, "", key="s_end", name="FND_CashEndConv")
    F.add("累計外部資金需求（轉股）", "RMB 億", {"E": "=E«F.s_ext»", **{c: f"={PREV[c]}«F.s_cum»+{c}«F.s_ext»" for c in LATE}}, "", key="s_cum", name="FND_ExtNeedCumConv")
    F.add("轉股稀釋：換股股數 ÷ 全面攤薄總股數", "比例", {"D": f"={S('cb_conv_shares')}/({S('shares_total_post_placing2')}+{S('cb_conv_shares')})"},
          "SRC_ZP_549、563", key="s_dil")

    F.section("六、或有與承諾（只列示；不作資金來源、不進外部資金需求）")
    F.add("或有：可換股債券本金（轉股則免償還）", "RMB 億", {"D": f"={S('cb_principal')}*{I('cb_redeem')}"}, "基準視為債務；換股價 HK$892.50", key="cl_cb",
          name="FND_Contingent", name_col="D")
    F.add("現價 ÷ 換股價", "倍", {"D": f"={S('price_20261007')}/{S('cb_conv_price')}"}, "SRC_ZP_569 ÷ SRC_ZP_548；<1＝價外（偏向償還）", key="cl_moneyness")
    F.add("對照：2025-10-31 現金＋短期投資＋可用承諾銀行授信合計（未拆）", "RMB 億", {"D": f"={S('liquidity_2510')}"},
          "SRC_ZP_110（Interested-party）；未提用授信額度單獨金額找不到 → 只列", key="cl_liq")
    F.add("未提用授信額度、合約承諾（算力採購等）", "—", {}, "找不到（中期公告未列；Z1 not_found）", key="cl_nf")
    F.add("對照：2027-01-08 解禁占總股本（媒體估算）", "%", {"D": f"={S('lockup_2027_01', 'lo')}"}, "SRC_ZP_568（Interested-party）；可能影響 2027 年再融資條件", key="cl_lock")

    F.section("七、配售用途對照（D16r；公告用途 vs 模型支出；只列）")
    F.add("2026-07 配售：研發（模型、人才、算力）用途金額", "RMB 億", {"L": f"=L«F.pl1»*{S('placing1_use_rd')}/{pct}"}, "SRC_ZP_539：55%；預計 2027 年底前用畢", key="u_pl1")
    F.add("2026-07 配售：截至 2026-08-31 已動用", "RMB 億", {"L": f"={S('placing1_utilised_0831')}*{fxh}/{mn}"}, "SRC_ZP_540：HK$10,954.8 百萬（34.92%）", key="u_pl1_used")
    F.add("模型：2026-07-01～08-31 融資前淨現金流出（2H26 按天數）", "RMB 億", {"L": f"=-L«F.netop»*{I('days_jul_aug')}/{d2}"}, "含中科加禾併購", key="u_model_ja")
    F.add("倍數：已動用 ÷ 模型同期流出", "倍", {"L": "=L«F.u_pl1_used»/L«F.u_model_ja»"}, ">1＝公司實際動用速度高於模型支出（可能含預付、押金、理財或模型低估的支出）", key="u_ratio",
          name="FND_UseVsModel", name_col="L")
    F.add("2026-09 配售＋可換股債券：研發、訓練推論、算力基建、人才用途金額", "RMB 億", {"L": f"=(L«F.pl2»+L«F.cb»)*{S('placing2_cb_use_rd')}/{pct}"},
          "SRC_ZP_550：60%；2028-06-30 前用畢", key="u_pl2")
    F.add("模型：2H26–2028 算力成本＋非算力成本（不含股權報酬）合計", "RMB 億",
          {"L": "=L«F.cc»+L«F.ncx»+F«F.cc»+F«F.ncx»+G«F.cc»+G«F.ncx»"}, "對照兩次用途合計（期間至 2028-06 vs 2028 全年，模型較寬）", key="u_model")
    F.add("倍數：兩次用途（研發與算力）合計 ÷ 模型同期支出", "倍", {"L": "=(L«F.u_pl1»+L«F.u_pl2»)/L«F.u_model»"}, "", key="u_ratio2")

    F.section("八、回流對照（D19：誰出錢；只列對照、不沖銷）")
    F.add("最大客戶 A：智譜對其銷售（FY2024＋1H25）", "RMB 億", {"D": f"={S('cust_a_overlap_p')}"}, "SRC_ZP_785（招股章程會計師報告，Verified；r2 P2：媒體 SRC_ZP_433 的 2.42 億差 10 倍）", key="rf_sales")
    F.add("最大客戶 A：智譜向其採購（同期）", "RMB 億", {"D": f"={S('cust_a_purch_p')}"}, "SRC_ZP_786（Verified；採購內容為數據庫及知識產權授權）；客戶兼供應商", key="rf_purch")
    F.add("採購 ÷ 銷售", "倍", {"D": "=D«F.rf_purch»/D«F.rf_sales»"}, "≈1：營收與採購互相抵銷", key="rf_ratio")
    F.add("2025 地方國資投資（杭州≥10、珠海華發 5、成都高新 3、上海浦東張江 10）", "RMB 億",
          {"D": f"={S('round_2025_03_hangzhou', 'lo')}+{S('round_2025_03_huafa')}+{S('round_2025_03_chengdu')}+{S('round_2025_07_shanghai')}"},
          "SRC_ZP_510–513（Interested-party）；對這些股東（或其算力平台）的採購額找不到", key="rf_soe")

    F.section("九、2025、1H26 現金對帳（經營現金流量表找不到 → 以期末現金變動 − 已知融資推得隱含經營＋投資＋其他）")
    F.add("期末現金變動", "RMB 億", {"D": f"={S('cash_end_fy25')}-{S('cash_end_fy24_b')}", "K": f"={S('cash_end_1h26')}-{S('cash_end_fy25')}"},
          "SRC_ZP_151、152、211", key="rc_dc")
    F.add("已知融資流入（2025：上市前各輪（國資）＋銀行借款淨增；1H26：IPO 淨額＋銀行借款淨增）", "RMB 億",
          {"D": "=D«F.rf_soe»+D«F.loan_chg»", "K": "=K«F.committed»"}, "2025 各輪到帳時點與金額部分為媒體轉述（下限）", key="rc_fin")
    F.add("隱含：經營＋投資＋其他現金流＝現金變動 − 已知融資", "RMB 億", {c: f"={c}«F.rc_dc»-{c}«F.rc_fin»" for c in ("D", "K")}, "", key="rc_impl")
    F.add("模型：融資前淨現金流", "RMB 億", {c: f"={c}«F.netop»" for c in ("D", "K")}, "", key="rc_model")
    F.add("差異（隱含 − 模型）", "RMB 億", {c: f"={c}«F.rc_impl»-{c}«F.rc_model»" for c in ("D", "K")}, "1H26＝第四節對帳差；2025 含理財、定期存款、營運資金與融資到帳時點差", key="rc_diff",
          name="FND_CashRecon")
    F.add("對照：經調整淨虧損（非 IFRS，負值）", "RMB 億", {"D": f"=-{S('adj_loss_fy25')}", "K": f"=-{S('adj_loss_1h26')}"}, "SRC_ZP_146、206", key="rc_adj")

    F.section("十、美元口徑（與 OpenAI v0.6 並排；÷ USD/CNY ÷ 10，同一匯率 D23）")
    for key_, zh, nm_ in (("fcf", "自由現金流", "FND_FCF_USD"), ("committed", "已到位融資合計", "FND_Committed_USD"), ("ext", "當年外部資金需求", "FND_ExtNeed_USD"),
                          ("cum", "累計外部資金需求", "FND_ExtNeedCum_USD"), ("end", "年底現金", "FND_CashEnd_USD"), ("s_cum", "累計外部資金需求（轉股）", "FND_ExtNeedCumConv_USD"),
                          ("s_end", "年底現金（轉股）", "FND_CashEndConv_USD")):
        cols = YC if key_ in ("fcf", "end") else YE
        F.add(f"{zh}（美元）", "$B", {c: f"={c}«F.{key_}»/{fxu}/{bn}" for c in cols}, "", key=f"u_{key_}", name=nm_)

    F.section("十一、r2 P2（招股章程，只列對照，不進現金流）：資本開支口徑")
    pairs = (("fy24", "FY2024"), ("1h25", "1H2025"))
    F.add("P2 公告資本開支（權責口徑；含使用權資產添置）：FY2024", "RMB 億", {"D": f"={S('capex_fy24')}"}, "SRC_ZP_150（年度公告比較數）", key="p2_acc_fy24")
    F.add("P2 公告資本開支（權責口徑）：1H2025", "RMB 億", {"D": f"={S('capex_1h25')}"}, "SRC_ZP_210（中期公告比較數）", key="p2_acc_1h25")
    for k, zh in pairs:
        F.add(f"P2 現金資本開支（購買物業及設備＋無形資產）：{zh}", "RMB 億", {"D": f"={S('capex_cash_total_' + k)}"}, "SRC_ZP_889／891（招股章程現金流量表）", key=f"p2_cash_{k}")
        F.add(f"P2 其中使用權資產添置（不是現金資本開支）：{zh}", "RMB 億", {"D": f"={S('rou_add_' + k)}"}, "SRC_ZP_898／899", key=f"p2_rou_{k}")
        F.add(f"P2 權責 − 現金：{zh}", "RMB 億", {"D": f"=D«F.p2_acc_{k}»-D«F.p2_cash_{k}»"}, "差額主要是使用權資產（租賃算力硬件與辦公室）", key=f"p2_diff_{k}")

    for c, w in (("A", 7), ("B", 66), ("C", 10)):
        F.ws.column_dimensions[c].width = w
    for c in YC + HC:
        F.ws.column_dimensions[c].width = 12
    F.ws.column_dimensions["J"].width = 2
    F.ws.column_dimensions["M"].width = 80
    F.ws.freeze_panes = "D7"

    # ═══════════════════════════ Reverse ═══════════════════════════
    X = Sheet(wb, "Reverse", "X",
              "Reverse — 反向模式（D22、D22r）：①管理層量化目標 ②分析師共識營收 → 所需倍數與反向外部資金 ③市值隱含營收；只作對照區塊，不回饋基準",
              "共同規則第 4 節：不以公司或分析師數字反推參數。本頁不被任何計算頁引用（只被 Checks 引用；test_reverse_not_fed_back）。"
              "共識：全部登錄（Simply Wall St 聚合共識＋國信、國證國際兩家券商），取平均為基準、最高最低為區間（工作單 Z4）；2029–2030 依 Inputs（基準持平於 FY2028）。",
              "反向資金＝同一支出（算力、本地化交付、非算力成本、利息、資本支出、融資與債務償還皆同 Funding）、營收換成目標營收；同一最低現金規則。金額 RMB 億。")
    X.years(D)
    fnd = lambda key, c: f"Funding!{c}«F.{key}»"  # noqa: E731

    X.section("一、管理層目標（D22 ①）：營收、獲利時點、毛利率的量化目標不存在（SRC_ZP_581）；只有 ARR 指引（D14r，只對照）")
    X.add("管理層量化營收／獲利目標", "—", {}, "不存在：中期業績公告與 2026-09-16 電話會只有 ARR 指引；『毛利率有望企穩回升』為質化說法（SRC_ZP_581、Z1 not_found）", key="t_none")
    X.add("ARR 指引：2026 年末（美元十億 × USD/CNY × 10）", "RMB 億", {"L": f"={S('mgmt_arr_guidance_ye26')}*{fxu}*{bn}"}, "SRC_ZP_576（Interested-party；ARR 定義未揭露）", key="t_arr")
    X.add("模型：2H26 年化營收（× 半年數）", "RMB 億", {"L": f"={fnd('rev', 'L')}*{halves}"}, "", key="t_model")
    X.add("差距：模型 ÷ ARR 指引 − 1", "比例", {"L": "=(L«X.t_model»-L«X.t_arr»)/L«X.t_arr»"}, "只對照（年末 ARR 為時點年化，模型為下半年平均）", key="t_gap")

    X.section("二、分析師共識營收（D22 ②；Interested-party：券商有承銷與交易誘因；2026-09 配售承銷商為中金、國泰君安國際）")
    E3 = ("E", "F", "G")
    sws = {"E": f"={S('cons_rev_fy26')}/{mn}", "F": f"={S('cons_rev_fy27')}/{mn}", "G": f"={S('cons_rev_fy28')}/{mn}"}
    gs = {"E": f"={S('broker_guosen_rev')}", "F": f"={S('broker_guosen_rev_fy27')}", "G": f"={S('broker_guosen_rev_fy28')}"}
    gz = {"E": f"={S('broker_guozheng_rev')}", "F": f"={S('broker_guozheng_rev_fy27')}", "G": f"={S('broker_guozheng_rev_fy28')}"}
    X.add("Simply Wall St 聚合共識（21–23 位分析師）", "RMB 億", sws, "SRC_ZP_582–584（2026-10-06）；百萬元 ÷ 100", key="c_sws")
    X.add("國信證券", "RMB 億", gs, "SRC_ZP_591、606、607（2026-09-04）", key="c_gs")
    X.add("國證國際", "RMB 億", gz, "SRC_ZP_592、608、609", key="c_gz")
    X.add("對照：大和 2029 年營收預測（單一券商）", "RMB 億", {"H": f"={S('broker_daiwa_rev_fy29')}"}, "SRC_ZP_610；只對照", key="c_dw")
    X.add("登錄筆數（家數：聚合共識 1 筆＋券商 2 家）", "筆", {c: f"=COUNT({c}«X.c_sws»,{c}«X.c_gs»,{c}«X.c_gz»)" for c in E3}, "工作單 Z4：全部登錄", key="c_n")
    X.add("平均（基準）", "RMB 億", {c: f"=AVERAGE({c}«X.c_sws»,{c}«X.c_gs»,{c}«X.c_gz»)" for c in E3}, "工作單 Z4：取平均為基準", key="c_avg")
    X.add("最低", "RMB 億", {c: f"=MIN({c}«X.c_sws»,{c}«X.c_gs»,{c}«X.c_gz»)" for c in E3}, "", key="c_min")
    X.add("最高", "RMB 億", {c: f"=MAX({c}«X.c_sws»,{c}«X.c_gs»,{c}«X.c_gz»)" for c in E3}, "", key="c_max")
    g_late = I("cons_g_late")

    def tgt(src):
        def f(c):
            if c == "D":
                return "=Revenue!D«V.gross»"
            if c in E3:
                return f"={c}«X.{src}»"
            if c in ("H", "I"):
                return f"={PREV[c]}«X.{'tgt' if src == 'c_avg' else 'tgt_' + src}»*(1+{g_late})"
            return None
        return f
    X.add("目標營收（反向；2025＝實際、2026–2028＝共識平均、2029–2030＝FY2028 ×（1＋Inputs 年增率））", "RMB 億", tgt("c_avg"), "", key="tgt", name="RVS_Target")
    X.add("目標營收：共識最低路徑", "RMB 億", tgt("c_min"), "", key="tgt_c_min", name="RVS_TargetLo")
    X.add("目標營收：共識最高路徑", "RMB 億", tgt("c_max"), "", key="tgt_c_max", name="RVS_TargetHi")

    X.section("三、正向對照（本模型基準；截頂後，容量上限只回乘雲端）")
    X.add("正向營收淨額（截頂後）", "RMB 億", {c: f"={fnd('rev', c)}" for c in YC}, "Funding（REV_NetCapped）", key="fwd")
    X.add("正向：按量計費（截頂後）", "RMB 億", {c: f"=Revenue!{c}«V.paygo»*Revenue!{c}«V.cap»" for c in YC}, "", key="f_api")
    X.add("正向：Coding Plan 訂閱（截頂後）", "RMB 億", {c: f"=Revenue!{c}«V.cp»*Revenue!{c}«V.cap»" for c in YC}, "", key="f_cp")
    X.add("正向：本地化部署＋廣告（不截頂）", "RMB 億", {c: f"=Revenue!{c}«V.onprem»+Revenue!{c}«V.ads»" for c in YC}, "", key="f_op")
    X.add("差距（目標 − 正向）", "RMB 億", {c: f"={c}«X.tgt»-{c}«X.fwd»" for c in YC}, "", key="gap", name="RVS_Gap")
    X.add("目標 ÷ 正向", "倍", {c: f"={c}«X.tgt»/{c}«X.fwd»" for c in YC}, "", key="ratio", name="RVS_Ratio")
    X.add("對照：共識最低 ÷ 正向", "倍", {c: f"={c}«X.tgt_c_min»/{c}«X.fwd»" for c in YC}, "", key="ratio_lo")
    X.add("對照：共識最高 ÷ 正向", "倍", {c: f"={c}«X.tgt_c_max»/{c}«X.fwd»" for c in YC}, "", key="ratio_hi")

    X.section("四、所需倍數（相對正向；1＝正向已足夠；本地化部署不耗算力，固定為正向值）")
    X.add("只靠 API 按量計費：所需倍數", "倍", {c: f"=({c}«X.tgt»-{c}«X.f_op»-{c}«X.f_cp»)/{c}«X.f_api»" for c in YC}, "", key="m_api", name="RVS_MultAPI")
    X.add("只靠 Coding Plan 訂閱：所需倍數", "倍", {c: f"=({c}«X.tgt»-{c}«X.f_op»-{c}«X.f_api»)/{c}«X.f_cp»" for c in YC}, "", key="m_sub", name="RVS_MultSub")
    X.add("兩線等比例：API 與訂閱共同倍數", "倍", {c: f"=({c}«X.tgt»-{c}«X.f_op»)/({c}«X.f_api»+{c}«X.f_cp»)" for c in YC}, "", key="m_prop", name="RVS_MultProp")
    X.add("2028：只靠 API 所需按量計費 token（同單價路徑）", "T", {"G": f"=(G«X.tgt»-G«X.f_op»-G«X.f_cp»)*{I('div_rev')}/Revenue!G«V.p»"},
          "REV_ApiPrice（元／百萬 token）", key="o_tok")
    X.add("2028：正向按量計費 token", "T", {"G": "=Demand!G«D.api_tok»"}, "DEM_Tok_API", key="o_tok_f")
    X.add("2028：只靠訂閱所需 Coding Plan 訂閱者（同 ARPU）", "萬人", {"G": "=Demand!G«D.subs»*G«X.m_sub»"}, "DEM_Subs_CP × 倍數", key="o_subs")
    X.add("2028：兩線等比例所需有效推論 GW（同 η、同 token 結構）", "GW", {"G": "=Compute!G«C.eff»*G«X.m_prop»"}, "正向有效推論 GW × 等比例倍數；只對照（反向資金不增加算力，見第五節說明）",
          key="o_gw")
    X.add("反向覆蓋率（目標營收 ÷ 正向全成本）", "倍", {c: f"={c}«X.tgt»/Cost!{c}«K.full»" for c in YC}, "同一支出下的覆蓋率；目標營收若需更多算力，實際覆蓋率較低", key="cov",
          name="RVS_Coverage")

    X.section("五、反向資金（目標營收＋同一支出；同一最低現金規則；只作對照）")
    X.add("自由現金流（反向）＝正向自由現金流＋（目標 − 正向）", "RMB 億", {c: f"={fnd('fcf', c)}+{c}«X.gap»" for c in YE}, "", key="fcf", name="RVS_FCF")
    X.add("期初現金", "RMB 億", {"E": f"={fnd('end', 'D')}", **{c: f"={PREV[c]}«X.end»" for c in LATE}}, "", key="open")

    def xpre(c):
        if c == "E":
            return f"={fnd('pre', 'E')}+E«X.gap»"
        return f"={c}«X.open»+{fnd('netop', c)}+{c}«X.gap»+{fnd('committed', c)}-{fnd('cb_rep', c)}"
    X.add("補足前期末現金", "RMB 億", {c: xpre(c) for c in YE}, "2026＝Funding 補足前＋（目標 − 正向）", key="pre")
    X.add("當年外部資金需求", "RMB 億", {c: f"=({fnd('min', c)}-{c}«X.pre»)*({fnd('min', c)}>{c}«X.pre»)" for c in YE}, "", key="ext")
    X.add("期末現金", "RMB 億", {c: f"={c}«X.pre»+{c}«X.ext»" for c in YE}, "", key="end", name="RVS_CashEnd")
    X.add("累計外部資金需求（反向）", "RMB 億", {"E": "=E«X.ext»", **{c: f"={PREV[c]}«X.cum»+{c}«X.ext»" for c in LATE}}, "", key="cum", name="RVS_ExtNeedCum")

    X.section("六、市值隱含營收（D22 ③、D21r）：EV ÷ 同業 EV／年化營收倍數（MiniMax，Analogy）→ 與正向營收比較")
    X.add("收盤價（2026-10-07）", "HKD／股", {"D": f"={S('price_20261007')}"}, "SRC_ZP_569", key="v_px")
    X.add("總股本（2026-09-16 配售後）", "股", {"D": f"={S('shares_total_post_placing2')}"}, "SRC_ZP_563（含非上市股份，同價計）", key="v_sh")
    X.add("市值（港幣）＝股價 × 股數 ÷ 10^8", "HKD 億", {"D": f"=D«X.v_px»*D«X.v_sh»/{I('div_yuan_yi')}"}, "對照：行情頁 3,396.48（SRC_ZP_571）", key="v_mc_hkd")
    X.add("市值（人民幣）", "RMB 億", {"D": f"=D«X.v_mc_hkd»*{fxh}"}, "× HKD/CNY", key="v_mc", name="RVS_MktCap", name_col="D")
    X.add("淨現金（估計，Derived）＝2026-06-30 現金＋2H26 配售與可換股債券淨額 − 2026-07 配售款已動用 − 銀行借款 − 可換股債券本金", "RMB 億",
          {"D": f"={S('cash_end_1h26')}+{fnd('pl1', 'L')}+{fnd('pl2', 'L')}+{fnd('cb', 'L')}-{fnd('u_pl1_used', 'L')}-{S('bank_loans_1h26')}-{S('cb_principal')}*{I('cb_redeem')}"},
          "chat 端補充：含 2026-07、09 新資金；Z5b V3：扣 2026-07 配售款截至 2026-08-31 已動用（Funding 第七節，SRC_ZP_540 × HKD/CNY）", key="v_nc", name="RVS_NetCash", name_col="D")
    X.add("對照：模型 2026 年底淨現金（期末現金 − 銀行借款 − 可換股債券本金）", "RMB 億",
          {"D": f"={fnd('end', 'E')}-{fnd('loan_end', 'E')}-{S('cb_principal')}*{I('cb_redeem')}"}, "含 2H26 模型營運消耗與中科加禾", key="v_nc_m")
    X.add("企業價值 EV＝市值 − 淨現金", "RMB 億", {"D": "=D«X.v_mc»-D«X.v_nc»"}, "", key="v_ev", name="RVS_EV", name_col="D")
    X.add("MiniMax 市值（美元）＝港幣市值 × HKD/CNY ÷ USD/CNY", "USD 億", {"D": f"={S('peer_mmx_mktcap')}*{fxh}/{fxu}"}, "SRC_ZP_596（2026-10-07 盤中）", key="p_mc")
    X.add("MiniMax 現金、等價物及短期投資", "USD 億", {"D": f"={S('peer_mmx_cash_1h26')}/{mn}"}, "SRC_ZP_603（2026-06-30；未扣債務，資料缺口）", key="p_cash")
    X.add("MiniMax EV", "USD 億", {"D": "=D«X.p_mc»-D«X.p_cash»"}, "", key="p_ev")
    X.add("MiniMax 年化營收＝1H26 × 半年數", "USD 億", {"D": f"={S('peer_mmx_rev_1h26')}*{halves}/{mn}"}, "SRC_ZP_601", key="p_rev")
    X.add("MiniMax EV ÷ 年化營收（推導值）", "倍", {"D": "=D«X.p_ev»/D«X.p_rev»"}, "SRC_ZP_604 所述公式，改用年化營收（chat 端補充）", key="p_mult")
    pm = "peer_mult_range"
    X.add("同業倍數：基準（推導值 × Inputs 係數）", "倍", {"D": f"=D«X.p_mult»*{I(pm)}"}, "Inputs（Analogy，1；0.5–2.0）", key="m_base", name="RVS_PeerMult", name_col="D")
    X.add("同業倍數：低端", "倍", {"D": f"=D«X.p_mult»*{D.I_cell(pm, 'G')}"}, "Inputs 低欄", key="m_lo")
    X.add("同業倍數：高端", "倍", {"D": f"=D«X.p_mult»*{D.I_cell(pm, 'H')}"}, "Inputs 高欄", key="m_hi")
    X.add("市值隱含營收（基準）＝EV ÷ 同業倍數", "RMB 億", {"D": "=D«X.v_ev»/D«X.m_base»"}, "市值『要求』的年營收（以同業當前倍數計）", key="mi", name="RVS_MktImpliedRev",
          name_col="D")
    X.add("市值隱含營收（倍數低端 → 營收高端）", "RMB 億", {"D": "=D«X.v_ev»/D«X.m_lo»"}, "", key="mi_hi", name="RVS_MktImpliedRevHi", name_col="D")
    X.add("市值隱含營收（倍數高端 → 營收低端）", "RMB 億", {"D": "=D«X.v_ev»/D«X.m_hi»"}, "", key="mi_lo", name="RVS_MktImpliedRevLo", name_col="D")
    X.add("市值隱含營收（美元）", "$B", {"D": f"=D«X.mi»/{fxu}/{bn}"}, "÷ USD/CNY ÷ 10", key="mi_usd")
    X.add("對照：智譜 EV ÷ 1H26 年化營收（實際）", "倍", {"D": f"=D«X.v_ev»/(Revenue!K«V.gross»*{halves})"}, "", key="z_mult1h")
    X.add("對照：智譜 EV ÷ 模型 2H26 年化營收", "倍", {"D": "=D«X.v_ev»/L«X.t_model»"}, "", key="z_mult2h")
    X.add("正向營收 ÷ 市值隱含營收（基準）", "倍", {c: f"={c}«X.fwd»/$D$«X.mi»" for c in YC}, ">1＝正向該年營收已達市值要求（以同業當前倍數計）", key="mi_ratio",
          name="RVS_FwdOverImplied")
    X.add("正向營收 ≥ 市值隱含營收的年度標示", "年", {c: f'=IF({c}«X.fwd»>=$D$«X.mi»,{ycol(c)},"")' for c in YC}, "", key="mi_y")
    X.add("摘要：正向營收首次達到市值隱含營收的年度", "年", {"D": f'=IF(COUNT($D$«X.mi_y»:$I$«X.mi_y»),MIN($D$«X.mi_y»:$I$«X.mi_y»),"期間內未達")'}, "", key="mi_first",
          name="RVS_MktImpliedYear", name_col="D")
    X.add("共識 FY2028 ÷ 市值隱含營收", "倍", {"G": "=G«X.tgt»/$D$«X.mi»"}, "", key="mi_cons")

    for c, w in (("A", 7), ("B", 66), ("C", 10)):
        X.ws.column_dimensions[c].width = w
    for c in YC + HC:
        X.ws.column_dimensions[c].width = 12
    X.ws.column_dimensions["J"].width = 2
    X.ws.column_dimensions["M"].width = 80
    X.ws.freeze_panes = "D7"
    return dict(F=F, X=X)


def checks(D, Z):
    """Z4 的 Checks 列。«F.»／«X.» 由 fill() 代換。"""
    fcf = "MAX(" + ",".join(f"ABS(Funding!{c}«F.fcf»-(Funding!{c}«F.rev»-Funding!{c}«F.cc»-Funding!{c}«F.opd»-Funding!{c}«F.ncx»+Funding!{c}«F.oth»-Funding!{c}«F.int»))"
                            for c in ALL) + ")"
    gapc = "MAX(" + ",".join(f"ABS(Funding!{c}«F.gap»-Cost!{c}«K.gap»)" for c in YC) + ")"
    rev26 = "ABS(Funding!E«F.rev»-Revenue!E«V.net_c»)"
    flows = ("rev", "cc", "opd", "ncx", "sbc", "oth", "int", "fcf", "capex", "netop", "eq", "committed")
    eadd = "MAX(" + ",".join(f"ABS(Funding!E«F.{k}»-Funding!K«F.{k}»-Funding!L«F.{k}»)" for k in flows) + ")"
    negext = "+".join(f"(Funding!{c}«F.ext»<0)" for c in YE)
    belowmin = "+".join(f"(Funding!{c}«F.end»<Funding!{c}«F.min»-0.00000001)" for c in YE)
    conv = "+".join(f"(Funding!{c}«F.s_cum»>Funding!{c}«F.cum»)" for c in YE)
    rows = [
        ("Funding：自由現金流恆等式（營收 − 算力 − 本地化交付 − 非算力 ＋ 其他收益 − 利息；各期最大絕對差）", "=" + fcf, 0, "tol", "工作單 Z4 第 1 步"),
        ("Funding：差額 − Cost COST_GapCash（年度欄最大絕對差）", "=" + gapc, 0, "tol", "融資缺口起點＝命題全成本口徑"),
        ("Funding：2026 營收（1H＋2H）− Revenue REV_NetCapped 2026", "=" + rev26, 0, "tol", ""),
        ("Funding：2026 流量列 E −（K＋L）最大絕對差", "=" + eadd, 0, "tol", ""),
        ("Funding：1H26 期末現金 − 中期報告（SRC_ZP_211）", f"=Funding!K«F.end»-{D.S('cash_end_1h26')}", 0, "tol", "1H26 對帳差使期末＝實際"),
        ("Funding：外部資金需求 < 0 的年數", "=" + negext, 0, "eq", ""),
        ("Funding：年底現金 < 最低現金的年數（補足後）", "=" + belowmin, 0, "eq", "外部資金需求補足至最低現金（容差 1e-8 億）"),
        ("Funding：轉股情境累計外部資金需求 > 基準的年數", "=" + conv, 0, "eq", "轉股不需償還，需求不得增加"),
        ("Funding：2026 年底現金 − 2H26 期末現金", "=Funding!E«F.end»-Funding!L«F.end»", 0, "tol", ""),
        ("Reverse：2025 目標營收 − 正向（應為 0：2025 為實際）", "=Reverse!D«X.tgt»-Reverse!D«X.fwd»", 0, "tol", ""),
        ("Reverse：共識登錄筆數 < 1 的年數", "=(Reverse!E«X.c_n»<1)+(Reverse!F«X.c_n»<1)+(Reverse!G«X.c_n»<1)", 0, "eq", "工作單 Z4：全部登錄"),
        ("Reverse：共識 最低 ≤ 平均 ≤ 最高 不成立的年數", "=" + "+".join(f"(Reverse!{c}«X.c_min»>Reverse!{c}«X.c_avg»)+(Reverse!{c}«X.c_avg»>Reverse!{c}«X.c_max»)" for c in ("E", "F", "G")),
         0, "eq", ""),
        ("對照：1H26 其他現金流（對帳差，RMB 億）", "=Funding!K«F.recon»", None, "info", "營運資金、租賃、匯兌等未建模項目"),
        ("對照：2025 現金對帳差異（隱含 − 模型，RMB 億）", "=INDEX(FND_CashRecon,1,1)", None, "info", "含理財、定期存款、上市前融資到帳時點"),
        ("對照：2026-07 配售已動用 ÷ 模型同期流出", "=FND_UseVsModel", None, "info", "公司實際動用速度 vs 模型"),
        ("預覽：累計外部資金需求 2030（RMB 億）", "=INDEX(FND_ExtNeedCum,1,6)", None, "info", "命題 2"),
        ("預覽：累計外部資金需求 2030（轉股情境）", "=INDEX(FND_ExtNeedCumConv,1,6)", None, "info", ""),
        ("預覽：首次缺口年", "=FND_FirstGapYear", None, "info", ""),
        ("預覽：2030 年底現金（RMB 億）", "=INDEX(FND_CashEnd,1,6)", None, "info", ""),
        ("預覽：只靠已到位融資撐不過的第一年", "=FND_NoExtFailYear", None, "info", ""),
        ("預覽：反向累計外部資金需求 2030（RMB 億）", "=INDEX(RVS_ExtNeedCum,1,6)", None, "info", "只對照"),
        ("預覽：市值隱含營收（RMB 億）", "=RVS_MktImpliedRev", None, "info", "D22r"),
        ("預覽：正向營收首次達到市值隱含營收的年度", "=RVS_MktImpliedYear", None, "info", ""),
        ("對照：MiniMax EV ÷ 年化營收", "=RVS_PeerMult", None, "info", ""),
        ("對照：2030 累計外部資金需求：智譜（$B）− OpenAI v0.6（$B）", f"=INDEX(FND_ExtNeedCum,1,6)/{D.I('fx_usdcny')}/{D.I('mul_bn_yi')}-INDEX(OAI_FND_ExtNeedCum,1,6)", None, "info",
         "OAI_Link；只並排"),
    ]
    return rows
