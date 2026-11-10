# Oracle 個股模型 v0.5.0：收支模組（收支型個股樣板）。
# 由 build_oracle.py 在存檔前呼叫 add_funding_sheets(ctx)，新增 Fund、RPO、Sites、Debt、Lease、RevDCF 六張工作表。
# 原則：
#   1. 只新增工作表，全部以公式引用既有的 In／Q／Thesis／PL／CF／Val；不改任何既有儲存格，v0.4.7 既有頁的數值逐格不變。
#   2. 現金數字只有一組：資金來源與用途一律引用 CF 頁（淨利起算的營運現金流），不另用 EBITDAR 率推算（v0.3.1 的做法不搬）。
#   3. 事實（債務逐檔、租賃到期表、RPO 分桶、站點）讀 data/funding_facts_v05.json（自 v0.3.1 company.json 抽出，資料截至 FY27Q1 10-Q）。
#   4. 排程還本在基準仍假設以原利率展延（v0.4.7 口徑不變）；本模組把到期表與「以 r_new 再融資」的增額利息列為備忘，是否進基準為判斷題。
import os, json, datetime
from openpyxl.styles import Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
PER = ["FY27 Q2–Q4", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"]     # 欄 D..I（與 CF 頁瀑布同欄）
PC = "DEFGHI"
# 各期起訖日（會計年度 6 月～5 月；首期自評價日次日起）
PDATE = [((2026, 9, 1), (2027, 5, 31))] + [((y - 1, 6, 1), (y, 5, 31)) for y in range(2028, 2033)]


def _hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


def _find(ws, text):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith(text):
            return r
    raise KeyError(text)


def add_funding_sheets(ctx):
    wb = ctx["wb"]; TR = ctx["TR"]; PR = ctx["PR"]; CR = ctx["CR"]; SC = ctx["SC"]; SCN = ctx["SCN"]; tc = ctx["tc"]
    F = json.load(open(os.path.join(HERE, "data", "funding_facts_v05.json"), encoding="utf-8"))
    rev = ctx.get("REV5")
    out = {}
    thc = lambda i: "D" if i == 0 else tc(2027 + i)          # Thesis 欄：D＝FY27 Q2–Q4；F..J＝FY2028..FY2032
    QREV = _find(wb["Q"], "合併營收（公司揭露）")

    NCOL = {"Debt": 11, "Lease": 11, "RPO": 11, "Sites": 8, "Fund": 10}
    def put(ws, r, lab, unit, vals, fmt="#,##0.000", note=None, bold=False, blue=False, ncol=None):
        ncol = ncol or NCOL.get(ws.title, 10)
        ws.cell(r, 1, lab).font = Font(bold=bold); ws.cell(r, 2, unit)
        for c, v in vals.items():
            if v is None: continue
            x = ws[f"{c}{r}"]; x.value = v; x.number_format = fmt
            if blue: x.font = BLUE
        if note: ws.cell(r, ncol, note)

    # ------------------------------------------------------------------ Debt
    D = wb.create_sheet("Debt")
    D["A1"] = "Debt — 債務逐檔、到期表與再融資備忘（收支模組）。藍字＝申報數字；黑字＝公式。金額十億美元。資料截至 FY27Q1 10-Q（2026-08-31）＋2026-11-04 交割的五檔新票據（In 頁 I2 節）。"
    D["A2"] = "基準口徑（v0.4.7 起不變）：到期借款以原利率展延，所以排程還本不進 CF 頁的融資需求。本頁列出逐檔到期與「若以新債利率 r_new 再融資」的增額利息，作為再融資壓力的備忘；是否改進基準為判斷題。"
    _hdr(D, 4, ["工具", "類別", "到期（月底）", "票面／有效利率", "本金（2026-08-31）", "年化利息", "來源與標記"])
    r0 = 5; ins = F["debt"]["instruments"]
    for i, (nm, cls, mat, rate, prin, src) in enumerate(ins):
        r = r0 + i
        y, m = int(mat[:4]), int(mat[5:7])
        last = (datetime.date(y + (m == 12), (m % 12) + 1, 1) - datetime.timedelta(days=1))
        D.cell(r, 1, nm); D.cell(r, 2, cls)
        c = D.cell(r, 3, last); c.number_format = "yyyy-mm-dd"; c.font = BLUE
        c = D.cell(r, 4, rate); c.number_format = "0.000%"; c.font = BLUE
        c = D.cell(r, 5, prin); c.number_format = "#,##0.000"; c.font = BLUE
        c = D.cell(r, 6, f"=D{r}*E{r}"); c.number_format = "#,##0.000"
        D.cell(r, 7, src)
    r1 = r0 + len(ins) - 1
    rT = r1 + 1
    put(D, rT, "合計（2026-08-31 面額）", None, {"E": f"=SUM(E{r0}:E{r1})", "F": f"=SUM(F{r0}:F{r1})"}, bold=True)
    D.cell(rT, 7, F["debt"]["note"])
    put(D, rT + 1, "加權有效利率", None, {"D": f"=F{rT}/E{rT}"}, "0.000%")
    put(D, rT + 2, "對照：In 頁借款本金 debt_q1（CF 頁利息與槓桿用）", None, {"E": "=debt_q1"})
    put(D, rT + 3, "差異：逐檔合計 − debt_q1（面額與 In 頁取值的差；不影響基準）", None, {"E": f"=E{rT}-E{rT+2}"})
    put(D, rT + 4, "對照：In 頁既有借款平均利率 r_old（Q1 利息費用年化 ÷ 本金）", None, {"D": "=r_old"}, "0.000%")
    # 新票據
    rn = rT + 7
    D.cell(rn - 1, 1, "2026-11-04 交割的五檔優先無擔保票據（In 頁 I2 節；到期月日找不到，暫列各到期年的 11 月 [Assumed]）").font = BOLD
    _hdr(D, rn, ["工具", "類別", "到期（月底）", "票息", "面額", "年化利息", "來源與標記"])
    NT = [(2029, "ex_n29", "ex_c29"), (2031, "ex_n31", "ex_c31"), (2033, "ex_n33", "ex_c33"), (2037, "ex_n37", "ex_c37"), (2046, "ex_n46", "ex_c46")]
    for i, (y, kn, kc) in enumerate(NT):
        r = rn + 1 + i
        D.cell(r, 1, f"優先無擔保票據 {y} 年到期（2026-11 發行）"); D.cell(r, 2, "無擔保")
        c = D.cell(r, 3, datetime.date(y, 11, 30)); c.number_format = "yyyy-mm-dd"; c.font = BLUE
        c = D.cell(r, 4, f"={kc}"); c.number_format = "0.000%"
        c = D.cell(r, 5, f"={kn}"); c.number_format = "#,##0.000"
        c = D.cell(r, 6, f"=D{r}*E{r}"); c.number_format = "#,##0.000"
        D.cell(r, 7, "PRNewswire 發行人原文 [Verified]（In 頁 I2 節）；到期月 [Assumed]")
    n0, n1 = rn + 1, rn + len(NT)
    put(D, n1 + 1, "新票據合計", None, {"E": f"=SUM(E{n0}:E{n1})", "F": f"=SUM(F{n0}:F{n1})"}, bold=True)
    put(D, n1 + 2, "檢查：新票據合計 − In 頁 ex_nt_par（應為 0）", None, {"E": f"=E{n1+1}-ex_nt_par"})
    # 強制轉換特別股
    rm = n1 + 5
    D.cell(rm - 1, 1, "強制轉換特別股（權益，不列入借款；股數已在 In 頁 mcps_sh、股利在 pref27–pref29）").font = BOLD
    _hdr(D, rm, ["工具", "清算價", "股利率", "強制轉換日", "轉換價下限", "", "來源與標記"])
    for i, cv in enumerate(F["debt"]["mcps"]):
        r = rm + 1 + i
        D.cell(r, 1, cv[0]); D.cell(r, 2, cv[1]).font = BLUE; c = D.cell(r, 3, cv[2]); c.font = BLUE; c.number_format = "0.00%"
        D.cell(r, 4, cv[3]).font = BLUE; D.cell(r, 5, cv[5]).font = BLUE; D.cell(r, 7, cv[6])
    # 到期表
    rl = rm + 5
    D.cell(rl - 1, 1, "到期表（會計年度）：逐檔推得 vs 10-Q 申報").font = BOLD
    _hdr(D, rl, ["項目", "單位", ""] + PER + ["FY2033 以後", "說明"])
    def sumifs(col, lo, hi, a, b):
        return f'SUMIFS({col}{a}:{col}{b},C{a}:C{b},">="&DATE({lo[0]},{lo[1]},{lo[2]}),C{a}:C{b},"<="&DATE({hi[0]},{hi[1]},{hi[2]}))'
    L = {}
    L["raw"] = rl + 1; L["tl"] = rl + 2; L["old"] = rl + 3; L["rep"] = rl + 4; L["chk"] = rl + 5; L["new"] = rl + 6; L["tot"] = rl + 7
    L["oint"] = rl + 8; L["nint"] = rl + 9; L["cpn"] = rl + 10; L["dint"] = rl + 11; L["cum"] = rl + 12
    put(D, L["raw"], "8/31 存量：依逐檔到期日加總", "US$B", {**{c: "=" + sumifs("E", *PDATE[i], r0, r1) for i, c in enumerate(PC)}, "J": f'=SUMIFS(E{r0}:E{r1},C{r0}:C{r1},">"&DATE(2032,5,31))'},
        note="到期日落在該會計年度內的本金")
    put(D, L["tl"], "定期貸款每季攤還（FY27 Q2–Q4 三次 × 0.14，自 2027-08 到期餘額扣回）", "US$B", {"D": 0.42, "E": -0.42}, blue=True,
        note="定期貸款 2026-09-30 起每季攤還約 0.14、2027-08-16 到期 [Interested-party／Derived]（逐檔表備註）")
    put(D, L["old"], "8/31 存量：排程還本（逐檔＋攤還調整）", "US$B", {c: f"={c}{L['raw']}+{c}{L['tl']}" for c in PC + "J"}, bold=True)
    rep = F["debt"]["reported_ladder"]
    put(D, L["rep"], "10-Q 申報到期表（v0.3.1 事實總帳；FY27 剩餘、FY2028–FY2031、其後）", "US$B",
        {**{c: rep[i] for i, c in enumerate("DEFGH")}}, blue=True, note="v0.3.1 把這五欄標成曆年，經逐檔對帳確認是會計年度（FY27 剩餘 3.007＝商業本票 0.337＋2027-04 票據 2.25＋定期貸款攤還 0.42）")
    put(D, L["chk"], "檢查：逐檔推得 − 申報（FY27–FY2031，應為 0）", "US$B", {c: f"={c}{L['old']}-{c}{L['rep']}" for c in "DEFGH"})
    put(D, L["new"], "2026-11 新票據到期", "US$B", {**{c: "=" + sumifs("E", *PDATE[i], n0, n1) for i, c in enumerate(PC)}, "J": f'=SUMIFS(E{n0}:E{n1},C{n0}:C{n1},">"&DATE(2032,5,31))'})
    put(D, L["tot"], "排程還本合計（存量＋新票據）", "US$B", {c: f"={c}{L['old']}+{c}{L['new']}" for c in PC + "J"}, bold=True,
        note="基準假設以原利率展延，不計入 CF 頁融資需求；Fund 頁列為「再融資需求」備忘")
    put(D, L["oint"], "到期債的年化利息（存量逐檔）", "US$B", {c: "=" + sumifs("F", *PDATE[i], r0, r1) for i, c in enumerate(PC)})
    put(D, L["nint"], "到期債的年化利息（新票據）", "US$B", {c: "=" + sumifs("F", *PDATE[i], n0, n1) for i, c in enumerate(PC)})
    put(D, L["cpn"], "到期債加權票息", "%", {c: f"=IF({c}{L['raw']}+{c}{L['new']}>0,({c}{L['oint']}+{c}{L['nint']})/({c}{L['raw']}+{c}{L['new']}),0)" for c in PC}, "0.00%",
        note="定期貸款攤還部分以同檔利率計，併入 FY2028 欄（簡化）")
    put(D, L["dint"], "備忘：若到期債以 r_new 再融資，新增的年化利息", "US$B", {c: f"=({c}{L['raw']}+{c}{L['new']})*r_new-({c}{L['oint']}+{c}{L['nint']})" for c in PC},
        note="＝到期本金 ×（r_new − 原票息）；基準不計（展延假設）。r_new＝In 頁 2026-11 五檔新票據加權票息")
    put(D, L["cum"], "備忘：累計新增年化利息（期末水準）", "US$B", {c: "=" + "+".join(f"{x}{L['dint']}" for x in PC[: i + 1]) for i, c in enumerate(PC)}, bold=True,
        note="到 FY2032 年底，利息費用每年比基準多這個金額（稅前）")
    rs = L["cum"] + 2
    S = {"mat5": rs, "cum32": rs + 1, "eps": rs + 2}
    put(D, S["mat5"], "FY27 Q2–Q4 至 FY2031 排程還本合計", "US$B", {"D": f"=SUM(D{L['tot']}:H{L['tot']})"}, bold=True)
    put(D, S["cum32"], "FY2032 年底累計新增年化利息（稅前）", "US$B", {"D": f"=I{L['cum']}"})
    put(D, S["eps"], "　折合情境 B 的 FY2032 每股稅後影響", "US$／股", {"D": f"=-I{L['cum']}*(1-etr)/CF!I{CR['B']['shd']}"}, "0.00",
        note="＝−累計新增年化利息 ×（1 − 有效稅率）÷ 情境 B FY2032 稀釋股數；只是量級備忘，未進 EPS 與目標價")
    for col, w in zip("ABCDEFGHIJK", [58, 10, 13, 13, 16, 12, 12, 12, 12, 14, 90]): D.column_dimensions[col].width = w
    D.column_dimensions["G"].width = 60
    D.freeze_panes = "B5"
    out["Debt"] = {"r0": r0, "r1": r1, "tot": rT, "n0": n0, "n1": n1, "L": L, "S": S, "hdr": rl}

    # ------------------------------------------------------------------ Lease
    Ls = wb.create_sheet("Lease")
    Ls["A1"] = "Lease — 租賃到期表與未起租承諾（收支模組）。藍字＝申報數字（10-Q FY27Q1 附註 6；v0.3.1 事實總帳）。金額十億美元、未折現。"
    Ls["A2"] = "在帳租賃付款表依 10-Q 的會計年度分欄（v0.3.1 標成曆年；債務到期表已逐檔確認為會計年度，租賃表無逐筆資料可驗證 [Derived]）。未起租 288 的現金租金路徑引用 Thesis 頁（三情境相同：合約性，不隨 MW 縮放）。"
    _hdr(Ls, 4, ["項目", "單位", ""] + PER[:5] + ["其後", "", "說明"])
    op, fn = F["leases"]["operating"], F["leases"]["finance"]
    put(Ls, 5, "在帳營業租賃付款", "US$B", {c: op[i] for i, c in enumerate("DEFGHI")}, blue=True, note="I 欄＝FY2032 以後合計")
    put(Ls, 6, "在帳融資租賃付款", "US$B", {c: fn[i] for i, c in enumerate("DEFGHI")}, blue=True)
    put(Ls, 7, "在帳租賃付款合計（未折現）", "US$B", {c: f"={c}5+{c}6" for c in "DEFGHI"}, bold=True)
    put(Ls, 8, "在帳租賃未折現付款總額", "US$B", {"D": "=SUM(D7:I7)"}, note="對照 10-Q 揭露 63.245")
    put(Ls, 9, "檢查：付款總額 − 申報 63.245（應為 0）", "US$B", {"D": f"=D8-{F['leases']['facts']['onBal']}"})
    put(Ls, 10, "在帳租賃負債（現值，FY27Q1 末；In 頁 lease_on）", "US$B", {"D": "=lease_on"})
    put(Ls, 12, "未起租租賃承諾（未折現；In 頁 lease_unc）", "US$B", {"D": "=lease_unc"}, bold=True, note=F["leases"]["uncommenced_note"])
    _hdr(Ls, 14, ["未起租租約的現金租金路徑（Thesis 頁「AI 機房租金」）", "單位", ""] + PER + ["", "說明"])
    LR = {}
    r = 15
    for s in SC:
        T = TR[s]
        LR[s] = {"rent": r, "sum": r + 1, "tail": r + 2, "ratio": r + 3, "liab": r + 4, "permw": r + 5}
        put(Ls, r, f"情境 {s}：AI 機房租金（未起租 288 依序起租＋超過已簽範圍的新租約）", "US$B", {c: f"=Thesis!{thc(i)}{T['rent']}" for i, c in enumerate(PC)}, bold=True)
        put(Ls, r + 1, "　FY27 Q2–Q4 至 FY2032 合計", "US$B", {"D": f"=SUM(D{r}:I{r})"})
        put(Ls, r + 2, "　模型期後尚未支付（288 − 合計；負值＝已含新租約）", "US$B", {"D": f"=lease_unc-D{r+1}"})
        put(Ls, r + 3, "　模型期內支付 ÷ 承諾總額", "%", {"D": f"=D{r+1}/lease_unc"}, "0.0%")
        put(Ls, r + 4, "　租賃負債（在帳＋已起租的未起租租約現值，期末；CF 頁）", "US$B", {c: f"=CF!{c}{CR[s]['lease']}" for c in PC}, "#,##0.0")
        put(Ls, r + 5, "　模型每 MW 年租金（（在帳年租金＋AI 機房租金）÷ 平均已連網 MW）", "US$M／MW·年",
            {c: f"=(rent_on*CF!{c}{CR[s]['L']}+{c}{r})/Thesis!{thc(i)}{T['avg']}/CF!{c}{CR[s]['L']}*1000" for i, c in enumerate(PC)}, "0.00",
            note="在帳年租金含非 AI 據點，所以是上限口徑")
        r += 7
    for col, w in zip("ABCDEFGHIJK", [66, 12, 4, 12, 12, 12, 12, 12, 12, 4, 100]): Ls.column_dimensions[col].width = w
    Ls.freeze_panes = "B5"
    out["Lease"] = {"R": LR}

    # ------------------------------------------------------------------ Sites
    St = wb.create_sheet("Sites")
    St["A1"] = "Sites — 具名站點（收支模組；自 v0.3.1 照搬）。站點表不驅動 MW 路徑（MW 由 In 頁併網速度與合約容量上限決定），用來對照容量上限與租賃承諾。MW 為 IT 口徑。"
    St["A2"] = "契約 MW 多為第三方（Epoch）估計的設施功率 ÷ 1.2 換算 [Derived]；逐站租約金額公司未揭露，所以無法逐站對到租賃承諾。"
    _hdr(St, 4, ["站點", "開發與承租", "契約 MW-IT", "已驗收 MW-IT", "時程", "信心", "現況與來源"])
    sites = F["sites"]; s0 = 5
    for i, s_ in enumerate(sites):
        r = s0 + i
        St.cell(r, 1, s_["name"]); St.cell(r, 2, s_["operator"])
        St.cell(r, 3, s_["planned"]).font = BLUE; St.cell(r, 4, s_["accepted"]).font = BLUE
        St.cell(r, 5, s_["date"]); St.cell(r, 6, s_["confidence"]); St.cell(r, 7, s_["status"])
    s1 = s0 + len(sites) - 1; sr = s1 + 1
    SR_ = {"tot": sr, "cap": sr + 1, "res": sr + 2, "ressh": sr + 3, "acc": sr + 4, "lease": sr + 6, "unc": sr + 7, "sum": sr + 8, "named": sr + 9}
    put(St, sr, "具名站點合計", "MW-IT", {"C": f"=SUM(C{s0}:C{s1})", "D": f"=SUM(D{s0}:D{s1})"}, "#,##0", bold=True)
    put(St, sr + 1, "對照：合約容量上限（情境 A／B；In 頁 cap_B）", "MW-IT", {"C": "=cap_B"}, "#,##0")
    put(St, sr + 2, "未具名容量（上限 − 具名站點）", "MW-IT", {"C": f"=C{sr+1}-C{sr}"}, "#,##0")
    put(St, sr + 3, "未具名容量占上限", "%", {"C": f"=C{sr+2}/C{sr+1}"}, "0.0%")
    put(St, sr + 4, "對照：FY27Q1 末已連網 AI MW（In 頁 mw_q1）÷ 具名站點已驗收", "x", {"C": f"=mw_q1/D{sr}"}, "0.00", note="已連網 MW 遠多於具名站點已驗收：多數在役產能在未具名的既有機房")
    put(St, sr + 6, "租賃承諾：在帳（未折現）", "US$B", {"C": F["leases"]["facts"]["onBal"]}, "#,##0.0", blue=True)
    put(St, sr + 7, "租賃承諾：未起租（未折現）", "US$B", {"C": "=lease_unc"}, "#,##0.0")
    put(St, sr + 8, "租賃承諾合計", "US$B", {"C": f"=C{sr+6}+C{sr+7}"}, "#,##0.0", bold=True)
    put(St, sr + 9, "其中可對到具名站點的合約金額", "US$B", {"C": 0}, "#,##0.0", blue=True, note="公司未揭露逐站租約金額 → 0；全部為未對到的殘差（本表不能驗證租金路徑）")
    for col, w in zip("ABCDEFG", [46, 44, 13, 14, 12, 8, 110]): St.column_dimensions[col].width = w
    out["Sites"] = {"s0": s0, "s1": s1, "R": SR_}

    # ------------------------------------------------------------------ RPO
    Rp = wb.create_sheet("RPO")
    Rp["A1"] = "RPO — 合約排程 vs 模型交付（收支模組）。RPO 664（FY27Q1 末）含全公司合約（OCI、SaaS、支援等），所以先與模型合併營收比，再扣掉非 OCI 營收推出「須由 OCI 認列」的下限，與模型 OCI 營收比。"
    Rp["A2"] = "與 v0.3.1 的差別：v0.3.1 直接拿全部排程 RPO 對算力容量上限，瓶頸金額偏高；本頁改為同口徑比較。排程權重沿用 v0.3.1（10-Q 分桶：≤12 個月 13%、13–36 個月 37%、37–60 個月 34%、其後 16%，桶內平均分攤 [Derived]）。RPO 排程不驅動營收（營收由 MW 驅動），本頁是命題的檢查。"
    _hdr(Rp, 4, ["項目", "單位", ""] + PER[:5] + ["", "", "說明"])
    w = F["rpo"]["bucketWeights"]
    put(Rp, 5, "排程權重（占 FY27Q1 末 RPO）", "%", {c: w[i] for i, c in enumerate("DEFGH")}, "0.000%", blue=True, note="FY27 Q2–Q4＝≤12 個月桶 × 9/12；其餘依桶內月數平均（v0.3.1 rpo.bucketWeights）")
    put(Rp, 6, "排程 RPO 認列（全公司合約）", "US$B", {c: f"=rpo*{c}5" for c in "DEFGH"}, bold=True)
    put(Rp, 7, "五期合計 ÷ RPO", "%", {"D": "=SUM(D6:H6)/rpo"}, "0.0%", note="其餘在 FY2032 以後認列")
    RR = {}
    r = 9
    for s in SC:
        T = TR[s]; P = PR[s]; C_ = CR[s]
        Rp.cell(r, 1, SCN[s]).font = BOLD
        k = {n: r + 1 + i for i, n in enumerate(["rev", "oci", "ai", "non", "cov", "impl", "gap", "gaphi", "cum", "b0", "b1", "dbl"])}
        RR[s] = k
        pl = lambda i: "DEFGH"[i]
        put(Rp, k["rev"], "模型合併營收（期間）", "US$B", {"D": f"=PL!D{P['rev']}-Q!I{QREV}", **{c: f"=PL!{c}{P['rev']}" for c in "EFGH"}}, note="FY27 Q2–Q4＝全年 − Q1 實績")
        put(Rp, k["oci"], "　其中 OCI 營收（含非 AI）", "US$B", {c: f"=Thesis!{thc(i)}{T['oci']}" for i, c in enumerate("DEFGH")})
        put(Rp, k["ai"], "　　其中 OCI AI 算力營收（MW 驅動）", "US$B", {c: f"=Thesis!{thc(i)}{T['rev']}" for i, c in enumerate("DEFGH")})
        put(Rp, k["non"], "　非 OCI 營收（SaaS、授權、支援、硬體、服務）", "US$B", {c: f"={c}{k['rev']}-{c}{k['oci']}" for c in "DEFGH"})
        put(Rp, k["cov"], "排程 RPO ÷ 模型合併營收", "x", {c: f"={c}$6/{c}{k['rev']}" for c in "DEFGH"}, "0.00", note="＞1＝已簽約排程超過模型全部營收")
        put(Rp, k["impl"], "須由 OCI 認列的排程 RPO（下限＝排程 − 非 OCI 營收）", "US$B", {c: f"=MAX(0,{c}$6-{c}{k['non']})" for c in "DEFGH"},
            note="假設非 OCI 營收全部來自 RPO（實際有一部分不是），所以這是 OCI 須認列金額的下限 [Derived]")
        put(Rp, k["gap"], "OCI 交付缺口（須認列下限 − 模型 OCI 營收）", "US$B", {c: f"={c}{k['impl']}-{c}{k['oci']}" for c in "DEFGH"}, bold=True,
            note="正值＝合約排程多於模型交付能力（Backlog 來不及變成營收）；負值＝模型交付超過排程")
        put(Rp, k["gaphi"], "OCI 交付缺口上限（排程全部由 OCI 認列 − 模型 OCI 營收）", "US$B", {c: f"={c}$6-{c}{k['oci']}" for c in "DEFGH"},
            note="假設非 OCI 營收完全不來自 RPO（另一個極端）。公司未揭露 RPO 的業務別，真實缺口落在上一列與本列之間：正負號不一致時，「Backlog 來不及交付」無法由公開資料證實")
        put(Rp, k["cum"], "累計未如期交付（下限口徑，只累加正值）", "US$B", {c: "=" + "+".join(f"MAX(0,{x}{k['gap']})" for x in "DEFGH"[: i + 1]) for i, c in enumerate("DEFGH")})
        put(Rp, k["b0"], "期初 RPO", "US$B", {"D": "=rpo", **{c: f"={chr(ord(c)-1)}{k['b1']}" for c in "EFGH"}}, "#,##0.0")
        put(Rp, k["b1"], "期末 RPO（依排程認列、不含新簽約）", "US$B", {c: f"={c}{k['b0']}-{c}$6" for c in "DEFGH"}, "#,##0.0")
        put(Rp, k["dbl"], "借款合計 ÷ 期末 RPO", "x", {c: f"=CF!{c}{C_['debt']}/{c}{k['b1']}" for c in "DEFGH"}, "0.00", note="v0.3.1 的債務／backlog 檢查；RPO 遞減而借款不降時上升")
        r = k["dbl"] + 2
    for col, w_ in zip("ABCDEFGHIJK", [58, 8, 4, 12, 12, 12, 12, 12, 4, 4, 100]): Rp.column_dimensions[col].width = w_
    Rp.freeze_panes = "B5"
    out["RPO"] = {"R": RR, "sched": 6, "w": 5}

    # ------------------------------------------------------------------ Fund
    Fu = wb.create_sheet("Fund")
    Fu["A1"] = "Fund — 資金來源、用途與融資瀑布（收支模組；收支型個股的主頁）。全部引用 CF 頁，同一組現金數字換成「來源／用途／缺口／融資」的排法；首期 FY27 Q2–Q4 自 Q1 末（2026-08-31）現金起算。金額十億美元。"
    Fu["A2"] = "讀法：營運來源 − 用途（資本支出、股利）＝融資前現金流；累加到期初現金上就是「不做任何新融資的累積現金」，低於最低現金的部分要靠外部融資。排程還本列為備忘（基準假設展延）。"
    _hdr(Fu, 3, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
    SUMR = {n: 4 + i for i, n in enumerate(["trough", "tyear", "short", "exf", "wf", "ext", "mat", "tot", "lev", "chk"])}
    labs = {"trough": ("融資前累積現金谷底（不做任何新融資）", "US$B"), "tyear": ("　谷底出現的期間", "—"), "short": ("　低於最低現金的缺口（最低現金 − 谷底；負值＝不需融資）", "US$B"),
            "exf": ("已發生增資淨額（2026-11-04 交割）", "US$B"), "wf": ("瀑布另需外部融資（投資級債＋新股＋高息債，五期合計）", "US$B"), "ext": ("外部融資合計（已發生＋瀑布）", "US$B"),
            "mat": ("備忘：排程還本（FY27 Q2–Q4 至 FY2031；基準假設展延）", "US$B"), "tot": ("備忘：含再融資的資本市場總需求", "US$B"),
            "lev": ("調整後槓桿峰值（投資級上限 lev_max）", "x"), "chk": ("檢查：現金橋差異絕對值合計（應為 0）", "US$B")}
    for n, (lab, u) in labs.items():
        Fu.cell(SUMR[n], 1, lab).font = Font(bold=n in ("trough", "ext")); Fu.cell(SUMR[n], 2, u)
    Fu.cell(SUMR["trough"], 6, "＝MIN（各期末「融資前累積現金」）；五期＝FY27 Q2–Q4 至 FY2031")
    Fu.cell(SUMR["mat"], 6, "Debt 頁到期表；公司的融資上限 20 也不含既有債再融資")
    FR = {}
    r = SUMR["chk"] + 3
    FUH = r
    _hdr(Fu, r, ["項目", "單位", ""] + PER + ["說明"]); r += 1
    DL = out["Debt"]["L"]
    for j, s in enumerate(SC):
        T = TR[s]; C_ = CR[s]; P = PR[s]
        Fu.cell(r, 1, SCN[s]).font = BOLD
        names = ["L", "h_src", "core", "pp", "nwc", "ocf", "h_use", "cxai", "cxna", "cx", "div", "use", "h_gap", "opgap", "flow", "chk1",
                 "h_cum", "cash0", "cumx", "exf", "need", "ig", "eq", "hy", "cash1", "chk2", "h_memo", "rent", "int", "mat", "lev", "debt"]
        k = {n: r + 1 + i for i, n in enumerate(names)}
        FR[s] = k
        cf = lambda key, c: f"CF!{c}{C_[key]}"
        for n, t in (("h_src", "一、營運來源"), ("h_use", "二、用途"), ("h_gap", "三、缺口"), ("h_cum", "四、現金橋與融資瀑布"), ("h_memo", "五、備忘（已含在營運現金或基準不計）")):
            Fu.cell(k[n], 1, t).font = BOLD
        put(Fu, k["L"], "期間長度", "年", {c: "=" + cf("L", c) for c in PC}, "0.00")
        put(Fu, k["ocf"], "營運現金流（期間）", "US$B", {"D": f"={cf('cfo','D')}-cfo_q1", **{c: "=" + cf("cfo", c) for c in "EFGHI"}}, bold=True, note="FY27 Q2–Q4＝全年 − Q1 實績 23.103")
        put(Fu, k["pp"], "　客戶預付淨額（流入 − 認列）", "US$B", {c: f"={cf('ppin',c)}-{cf('pprec',c)}" for c in PC}, note="客戶出資：AI 資本支出 × 覆蓋比，扣掉已收現、本期只認列營收的部分")
        put(Fu, k["nwc"], "　營運資金變動（負值＝占用現金）", "US$B", {c: f"=-{cf('nwc',c)}" for c in PC})
        put(Fu, k["core"], "　本業營運現金（淨利＋折舊攤銷＋SBC；已扣利息、租金、稅）", "US$B", {c: f"={c}{k['ocf']}-{c}{k['pp']}-{c}{k['nwc']}" for c in PC})
        put(Fu, k["cxai"], "　AI 資本支出（Thesis）", "US$B", {c: f"=Thesis!{thc(i)}{T['cx']}" for i, c in enumerate(PC)})
        put(Fu, k["cx"], "資本支出合計", "US$B", {"D": f"={cf('cx','D')}-capex_q1", **{c: "=" + cf("cx", c) for c in "EFGHI"}}, note="FY27 Q2–Q4＝指引中點 92.5 − Q1 實績 28.499")
        put(Fu, k["cxna"], "　非 AI 資本支出", "US$B", {c: f"={c}{k['cx']}-{c}{k['cxai']}" for c in PC})
        put(Fu, k["div"], "股利（普通股＋特別股）", "US$B", {"D": f"={cf('div','D')}-div_q1", **{c: "=" + cf("div", c) for c in "EFGHI"}})
        put(Fu, k["use"], "用途合計", "US$B", {c: f"={c}{k['cx']}+{c}{k['div']}" for c in PC}, bold=True)
        put(Fu, k["opgap"], "自由現金流（營運現金流 − 資本支出）", "US$B", {c: f"={c}{k['ocf']}-{c}{k['cx']}" for c in PC})
        put(Fu, k["flow"], "融資前現金流（營運來源 − 用途）", "US$B", {c: f"={c}{k['ocf']}-{c}{k['use']}" for c in PC}, bold=True, note="負值＝當期要動用現金或外部融資")
        put(Fu, k["chk1"], "　檢查：與 CF 頁「融資前現金流」差異（應為 0）", "US$B", {c: f"={c}{k['flow']}-{cf('flow',c)}" for c in PC})
        put(Fu, k["cash0"], "期初現金（首期＝FY27Q1 末）", "US$B", {c: "=" + cf("cash0", c) for c in PC})
        put(Fu, k["cumx"], "融資前累積現金（不做任何新融資，期末）", "US$B", {c: "=cash_q1+" + "+".join(f"{x}{k['flow']}" for x in PC[: i + 1]) for i, c in enumerate(PC)}, bold=True,
            note="＝Q1 末現金＋累計融資前現金流；不含已發生增資與瀑布融資、也不含其利息影響以外的調整（利息已在營運現金流內）")
        put(Fu, k["exf"], "＋ 已發生增資淨額（新股 9.9＋票據 10.0）", "US$B", {c: "=" + cf("exf", c) for c in PC})
        put(Fu, k["need"], "外部融資需求（補足到最低現金）", "US$B", {c: "=" + cf("need", c) for c in PC})
        put(Fu, k["ig"], "＋ 瀑布：新增投資級債", "US$B", {c: "=" + cf("ig", c) for c in PC})
        put(Fu, k["eq"], "＋ 瀑布：新股", "US$B", {c: "=" + cf("eq", c) for c in PC})
        put(Fu, k["hy"], "＋ 瀑布：高息債", "US$B", {c: "=" + cf("hy", c) for c in PC})
        put(Fu, k["cash1"], "期末現金", "US$B", {c: "=" + cf("cash1", c) for c in PC}, bold=True)
        put(Fu, k["chk2"], "　檢查：期初＋融資前現金流＋增資＋瀑布 − 期末（應為 0）", "US$B",
            {c: f"={c}{k['cash0']}+{c}{k['flow']}+{c}{k['exf']}+{c}{k['ig']}+{c}{k['eq']}+{c}{k['hy']}-{c}{k['cash1']}" for c in PC})
        put(Fu, k["rent"], "現金租金（在帳年租金＋AI 機房租金；已含在營運現金流）", "US$B", {c: f"=rent_on*{c}{k['L']}+Thesis!{thc(i)}{T['rent']}" for i, c in enumerate(PC)})
        put(Fu, k["int"], "利息費用（已含在營運現金流）", "US$B", {"D": f"={cf('int','D')}-int_q1", **{c: "=" + cf("int", c) for c in "EFGHI"}})
        put(Fu, k["mat"], "排程還本（Debt 頁；基準假設展延，不計入需求）", "US$B", {c: f"=Debt!{c}{DL['tot']}" for c in PC})
        put(Fu, k["lev"], "調整後槓桿（借款＋租賃負債）÷（EBITDA＋租金）", "x", {c: "=" + cf("lev", c) for c in PC}, "0.00")
        put(Fu, k["debt"], "借款合計（期末）", "US$B", {c: "=" + cf("debt", c) for c in PC}, "#,##0.0")
        c = "CDE"[j]
        Fu[f"{c}{SUMR['trough']}"] = f"=MIN(D{k['cumx']}:H{k['cumx']})"
        Fu[f"{c}{SUMR['tyear']}"] = f"=INDEX($D${FUH}:$H${FUH},MATCH({c}{SUMR['trough']},D{k['cumx']}:H{k['cumx']},0))"
        Fu[f"{c}{SUMR['short']}"] = f"=min_cash-{c}{SUMR['trough']}"
        Fu[f"{c}{SUMR['exf']}"] = f"=SUM(D{k['exf']}:H{k['exf']})"
        Fu[f"{c}{SUMR['wf']}"] = f"=SUM(D{k['ig']}:H{k['hy']})"
        Fu[f"{c}{SUMR['ext']}"] = f"={c}{SUMR['exf']}+{c}{SUMR['wf']}"
        Fu[f"{c}{SUMR['mat']}"] = f"=SUM(D{k['mat']}:H{k['mat']})"
        Fu[f"{c}{SUMR['tot']}"] = f"={c}{SUMR['ext']}+{c}{SUMR['mat']}"
        Fu[f"{c}{SUMR['lev']}"] = f"=MAX(D{k['lev']}:H{k['lev']})"
        Fu[f"{c}{SUMR['chk']}"] = "=" + "+".join(f"ABS({x}{k[n]})" for n in ("chk1", "chk2") for x in PC)
        for n in SUMR:
            if n != "tyear": Fu[f"{c}{SUMR[n]}"].number_format = "0.00" if n == "lev" else "#,##0.0"
        r = k["debt"] + 2
    for col, w_ in zip("ABCDEFGHIJ", [64, 8, 13, 13, 13, 13, 13, 13, 13, 90]): Fu.column_dimensions[col].width = w_
    Fu.column_dimensions["F"].width = 14
    Fu.freeze_panes = "C4"
    out["Fund"] = {"SUM": SUMR, "R": FR, "hdr": FUH}

    # ------------------------------------------------------------------ RevDCF
    Rv = wb.create_sheet("RevDCF")
    Rv["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（一次只動一個驅動，其餘維持基準；整本重算反解到「加權目標價＝現價」）。由 scripts/reverse_v05.py 產生 data/reverse_v05.json，建置時寫入；本頁是建置快照。"
    Rv["A2"] = "讀法：「現價隱含值」落在 In 頁敏感度低／高區間之外，表示單靠這個驅動無法在合理範圍內支持現價。區間外或無解都照實列出，不外推。"
    _hdr(Rv, 4, ["驅動", "In 鍵", "單位", "基準", "現價隱含值", "隱含 ÷ 基準", "敏感度低", "敏感度高", "是否在區間內", "中立門檻隱含值（目標價＝現價 × 0.75）", "說明"])
    RV = {}
    if rev:
        Rv["A3"] = f"快照：現價 {rev['px']}、基準加權目標價 {rev['base_tp']:.2f}、求解日 {rev['date']}；容差 ±{rev['tol']} 美元。"
        for i, d in enumerate(rev["drivers"]):
            r = 5 + i; RV[d["id"]] = r
            Rv.cell(r, 1, d["label"]); Rv.cell(r, 2, d["key"]); Rv.cell(r, 3, d["unit"])
            for col, key in ((4, "base"), (5, "eq_px"), (7, "lo"), (8, "hi"), (10, "neutral")):
                v = d.get(key)
                c = Rv.cell(r, col, v if v is not None else ("—" if key in ("lo", "hi") else "搜尋範圍內無解")); c.number_format = d.get("fmt", "0.00")
            if isinstance(d.get("eq_px"), (int, float)) and d.get("base"):
                c = Rv.cell(r, 6, f"=E{r}/D{r}"); c.number_format = "0.00"
            Rv.cell(r, 9, d["inrange"]); Rv.cell(r, 11, d["note"])
        n = 5 + len(rev["drivers"])
        Rv.cell(n + 1, 1, "快照檢查：求解時的基準加權目標價 − 本檔加權目標價（應為 0；不為 0 表示快照過期，重跑 scripts/reverse_v05.py）")
        c = Rv.cell(n + 1, 4, f"={rev['base_tp']!r}-Val!C{ctx['CON']['tp']}"); c.number_format = "0.0000"
        RV["_chk"] = n + 1
    else:
        Rv["A3"] = "尚未求解（data/reverse_v05.json 不存在，或為敏感度建置）。"
    for col, w_ in zip("ABCDEFGHIJK", [44, 16, 14, 11, 13, 12, 11, 11, 14, 22, 110]): Rv.column_dimensions[col].width = w_
    out["RevDCF"] = {"R": RV}
    return out
