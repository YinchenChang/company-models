# v0.3.13（2026-11-21）：供應承諾對採購恆等式（Recon I 節）＋ Q3 FY27 承諾、擔保、存貨三項（Hist 頁）＋ 風險登錄更新。
# 依據：Andy 2026-11-21 10:27（於 AI供應鏈）「精读原文以及最新季度…我们把这块给做掉」；對帳單 Drive 02-AI供應鏈/對帳/
#       20261121_對帳單_NVIDIA供應承諾對模型銷貨成本.md。
# 性質：只新增（Hist 頁末、Recon 頁末、README 兩列風險登錄改文字、README 頁末沿革列）；既有數值格、目標價、評等不變。
#       本節只作對照與敏感度，不進目標價、不用於校準（X3）。Q3 FY27 損益錨點滾入不在本版（等 10-Q，另版）。
# 資料：NVIDIA 8-K 2026-11-17（accession 0001045810-26-000088）EX-99.2 CFO Commentary、EX-99.3 Supplemental Financial
#       Information（未經查核、furnished）；10-Q 期間結束 2026-07-26（accession 0001045810-26-000075）Note 7（存貨、應計負債）、Note 10。
# build_nvidia.py 在 coord_post.stamp 之後、存檔前呼叫 apply(wb)。
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName

SRC_Q3 = "NVIDIA 8-K 2026-11-17（accession 0001045810-26-000088）EX-99.2 CFO Commentary／EX-99.3 Supplemental Financial Information（未經查核）"
SRC_Q2 = "NVIDIA 10-Q（期間結束 2026-07-26；accession 0001045810-26-000075）"
V, IP, AS, DV = "Verified", "Interested-party", "Assumed", "Derived"
BLUE = Font(color="000000FF")
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")

# 鍵, 項目, 值, 標記, 來源
HIST = [
    ("h_sc_q3", "供應與產能承諾合計（2026-10-25）", 390, V, SRC_Q3 + "：Future commitments 表"),
    ("h_sc_q3_27r", "　其中 FY2027 剩餘到期（2026-10-25）", 69, V, "同上"),
    ("h_sc_q3_28", "　其中 FY2028 到期", 203, V, "同上"),
    ("h_sc_q3_29", "　其中 FY2029 到期", 92, V, "同上"),
    ("h_sc_q3_30p", "　其中 FY2030 以後到期（10＋8＋8）", 26, V, "同上"),
    ("h_sc_q2_27r", "供應與產能承諾：FY2027 剩餘到期（2026-07-26）", 92, V, SRC_Q2 + " Note 10"),
    ("h_sc_q2_28", "　FY2028 到期（2026-07-26）", 87, V, "同上"),
    ("h_sc_q2_29", "　FY2029 到期（2026-07-26）", 88, V, "同上"),
    ("h_cm_cloud_q3", "雲端服務協議承諾（2026-10-25；前一季 29）", 28, V, SRC_Q3),
    ("h_cm_lease_q3", "未起租資料中心租約承諾（2026-10-25；前一季 自用 25＋代第三方 20）", 74, V, SRC_Q3 + "：公司稱部分可能轉讓第三方"),
    ("h_cm_eq_q3", "股權投資承諾（2026-10-25；前一季 25）", 20, V, SRC_Q3),
    ("h_cm_cx_q3", "資本支出承諾（2026-10-25；前一季 8）", 6, V, SRC_Q3),
    ("h_cm_neo_q3", "加碼承諾：新雲協議（2026-10-25；前一季 AI cloud agreements 36）", 32, V, SRC_Q3),
    ("h_gu_lps_q3", "擔保上限：新雲的土地、電力、機房（2026-10-25）", 3.5, V, SRC_Q3),
    ("h_gu_sbe_q3", "擔保上限：SB Energy（PORTS-Pike，約 4.25 GW，OpenAI 20 年租約；首期預計 FY2029）", 105, V, SRC_Q3),
    ("h_gu_pwr_q3", "擔保上限：電力與系統信用支持（2026-10-25；本季新增）", 9.5, V, SRC_Q3),
    ("h_inv_rm_q3", "存貨：原物料（2026-10-25）", 15.802, V, SRC_Q3 + "：Inventories 表"),
    ("h_inv_wip_q3", "存貨：在製品（2026-10-25）", 17.767, V, "同上"),
    ("h_inv_fg_q3", "存貨：製成品（2026-10-25）", 6.155, V, "同上"),
    ("h_inv_rm_q2", "存貨：原物料（2026-07-26）", 11.341, V, SRC_Q2 + " Note 7"),
    ("h_inv_wip_q2", "存貨：在製品（2026-07-26）", 13.377, V, "同上"),
    ("h_inv_fg_q2", "存貨：製成品（2026-07-26）", 6.857, V, "同上"),
    ("h_cogs_h1", "營業成本 H1 FY27", 44.538, V, SRC_Q2 + " 損益表"),
    ("h_cogs_q3", "營業成本 Q3 FY27", 29.168, V, SRC_Q3 + "：損益表"),
    ("h_rev_q3a", "營收 Q3 FY27（實績；外部稽核優化 W3-4 起已滾入 Q 頁 I 欄，本列供 Recon I 節用）", 115.844, V, SRC_Q3 + "：損益表"),
    ("h_q4g_rev", "Q4 FY27 營收指引中點（14 週；±3%）", 130, IP, SRC_Q3 + "：Outlook。來源誘因：公司對自身展望"),
    ("h_q4g_gm", "Q4 FY27 毛利率指引中點（71.0%–72.0%）", 0.715, IP, "同上"),
    ("h_g28_co", "公司 FY2028 營收成長展望（約 75%；法說口述，二手）", 0.75, IP, "2026-11-17 法說（investing.com 逐字稿摘要，二手）；見對帳單 20261116 合併單 r3。來源誘因：公司對自身展望"),
    ("h_ar_q3", "應收帳款（2026-10-25；DSO 64 天，前一季 60 天）", 81.033, V, SRC_Q3 + "：公司稱因 Rubin 出貨時點與對投資等級客戶延長付款條件"),
    ("h_eq_imp_q3", "Q3 FY27 非上市權益證券減損與未實現損失", 1.846, V, SRC_Q3 + "：非上市權益證券變動表"),
    ("h_xs_acc_q2", "超額存貨採購義務應計（2026-07-26；年初 2.739）", 2.138, V, SRC_Q2 + " Note 7"),
    ("sc_cover", "供應承諾模組假設：採購等值中由承諾支應的比例（基準 1；Q3 點讀數見 Recon I1）", 1.0, AS, "無揭露；取 1 時超額估計最小（保守）"),
    ("sc_q4inv_tr", "供應承諾模組假設：Q4 FY27 存貨增加的趨勢案（＝Q3 增幅）", "=h_inv_rm_q3+h_inv_wip_q3+h_inv_fg_q3-h_inv_rm_q2-h_inv_wip_q2-h_inv_fg_q2", DV, "Q3 存貨季增（8-K EX-99.3）"),
]


def _name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _find(ws, col, text, start=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(text)


def apply(wb):
    HS, RC, PL, RD = wb["Hist"], wb["Recon"], wb["PL"], wb["README"]

    # ---------- Hist：頁末追加（不位移既有列） ----------
    r = HS.max_row + 2
    HS.cell(r, 1, "供應承諾、擔保、存貨三項與 Q3 FY27 已公布數字（v0.3.13 起；Recon I 節用。Q3 損益已於外部稽核優化 W3-4（2026-11-23）滾入 Q 頁；本區不回饋 PL、DCF、Val）").font = BOLD
    r += 1
    for j, h in enumerate(("鍵", "項目", "值"), 1):
        HS.cell(r, j, h).font = BOLD
    HS.cell(r, 7, "標記").font = BOLD; HS.cell(r, 8, "來源").font = BOLD
    for key, item, val, tag, src in HIST:
        r += 1
        HS.cell(r, 1, key); HS.cell(r, 2, item)
        c = HS.cell(r, 3, val); c.number_format = "#,##0.000"
        if not (isinstance(val, str) and val.startswith("=")):
            c.font = BLUE
        HS.cell(r, 7, tag); HS.cell(r, 8, src)
        _name(wb, key, f"Hist!$C${r}")

    # ---------- PL 位置（依標籤找，不寫死列號） ----------
    def blk(n):   # 第 n 個「B. 報告值」區塊（依序為情境 A、B、C；第 4 個為 v0.1 對照）
        b = 1
        for _ in range(n):
            b = _find(PL, 1, "B. 報告值", start=b + 1, exact=False)
        rev = _find(PL, 1, "營收", start=b); gp = _find(PL, 1, "毛利", start=rev)
        eps = _find(PL, 1, "EPS（不含權益證券損益）", start=gp)
        return rev, gp, eps
    S = {"A": blk(1), "B": blk(2), "C": blk(3)}
    assert S["A"][0] < S["B"][0] < S["C"][0], S
    CO = wb["Coord"]
    co28 = _find(CO, 1, "營收（情境 B）FY2028"); co29 = _find(CO, 1, "營收（情境 B）FY2029")

    # ---------- Recon：I 節（頁末追加） ----------
    r0 = RC.max_row + 2
    R = {}

    def put(key, label, cells=(), note=None, bold=False, fmt=None):
        nonlocal r0
        R[key] = r0
        c = RC.cell(r0, 1, label)
        if bold:
            c.font = BOLD
        for j, v in enumerate(cells, 2):
            if v is None:
                continue
            cc = RC.cell(r0, j, v)
            if fmt:
                cc.number_format = fmt
            if bold:
                cc.font = BOLD
        if note:
            RC.cell(r0, 10, note).alignment = WRAP
        r0 += 1

    put("t", "I. 供應承諾對採購恆等式（v0.3.13 起；只作對照與敏感度，不進目標價、不校準；對帳單 20261121_對帳單_NVIDIA供應承諾對模型銷貨成本）", bold=True)
    put("n", "恆等式：採購等值＝銷貨成本＋存貨增加。年底存貨＝年初存貨＋到期承諾 ÷ 承諾支應比例 − 銷貨成本。限制：銷貨成本含保固、提列等非採購項；承諾含製造產能（部分進預付與其他資產，未單獨揭露）；公司稱部分協議在下正式訂單前可取消、改期或調整（比例未揭露）；Q3 數字未經查核。",
        note="出處：" + SRC_Q3 + "；" + SRC_Q2)
    r0 += 1
    put("h1", "I1. 已實現期間的恆等式核對", ("銷貨成本", "存貨增加", "採購等值", "FY2027 剩餘到期承諾的減少", "承諾支應比例（點讀數）"), bold=True)
    put("a_h1", "H1 FY27（2026-02～07）", ("=h_cogs_h1", f"=h_inv_rm_q2+h_inv_wip_q2+h_inv_fg_q2-Hist!E{_find(HS, 1, 'h_inv')}", "=B{r}+C{r}", "—", "—"),
        note="年初存貨＝Hist 頁 FY2026 年底存貨", fmt="#,##0.00")
    put("a_q3", "Q3 FY27（2026-08～10）", ("=h_cogs_q3", "=sc_q4inv_tr", "=B{r}+C{r}", "=h_sc_q2_27r-h_sc_q3_27r", "=E{r}/D{r}"),
        note="FY2027 剩餘到期承諾由 92 降到 69。點讀數，不是上下限：期間若有新增歸入 FY2027 的承諾，實際動用高於 23；若有承諾改期到以後年度（同季 FY2028 到期額由 87 升到 203），實際動用低於 23", fmt="#,##0.00")
    r0 += 1
    put("h2", "I2. 假設與 Q4 FY27 推算", ("值", "標記"), bold=True)
    put("dn", "正常存貨水位（年；＝Q3 年底存貨 ÷ Q3 銷貨成本年化）", ("=(h_inv_rm_q3+h_inv_wip_q3+h_inv_fg_q3)/(h_cogs_q3*4)", DV), fmt="0.0000",
        note="約 124 天；FY2026 年底為 115 天。存貨已含 Rubin 備料，取現值偏寬（超額估計偏小）。口徑：當季存貨 ÷ 當季銷貨成本年化（落後口徑），I3 套用在次年銷貨成本（前瞻口徑），兩者不完全可比：次年銷貨成本成長時低估超額（情境 A），下降時高估（情境 C）")
    put("q4c", "Q4 FY27 銷貨成本（指引中點）", ("=h_q4g_rev*(1-h_q4g_gm)", IP), fmt="#,##0.00")
    put("i27f", "FY2027 年底存貨：FY2027 剩餘承諾 69 全額執行", ("=h_inv_rm_q3+h_inv_wip_q3+h_inv_fg_q3+h_sc_q3_27r/sc_cover-B{q4c}", DV), fmt="#,##0.00",
        note="預登錄檢驗點：FY27 Q4 財報（約 2027-02）的年底存貨。全額執行約 71.7；趨勢案約 47.9；落在兩者之間表示部分承諾改期或未執行")
    put("i27t", "FY2027 年底存貨：趨勢案（Q4 增幅＝Q3 增幅）", ("=h_inv_rm_q3+h_inv_wip_q3+h_inv_fg_q3+sc_q4inv_tr", DV), fmt="#,##0.00")
    put("tx", "提列情境：超額存貨中認列損失的比例（低／中／高）", (0.1, 0.3, 1.0, AS), fmt="0%",
        note="無實證依據，只給區間。參考：FY2026 Q1 H20 超額存貨與採購義務費用 4.5；FY2027 前九個月存貨提列 2.3（銷貨成本的 3.1%）")
    r0 += 1
    put("h3", "I3. 各路徑的 FY2028（US$B）", ("情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "另一側讀數（情境 B）", "公司說法（成長 75%、毛利率指引中點）"), bold=True)
    cols = "BCDEF"

    def row5(key, label, fA, fO, fCo, note=None, fmt="#,##0.00"):
        vals = [fA(s) for s in "ABC"] + [fO, fCo]
        put(key, label, [v.replace("{c}", cols[i]) for i, v in enumerate(vals)], note=note, fmt=fmt)
    row5("rev", "FY2028 營收", lambda s: f"=PL!E{S[s][0]}", f"=Coord!C{co28}", "=(h_rev_h1fy27+h_rev_q3a+h_q4g_rev)*(1+h_g28_co)",
         note="公司說法：FY2027＝Q1、Q2、Q3 實績＋Q4 指引中點；成長率為法說口述（Interested-party）")
    row5("cg", "FY2028 銷貨成本", lambda s: f"=PL!E{S[s][0]}-PL!E{S[s][1]}", f"={{c}}{{rev}}*(1-PL!E{S['B'][1]}/PL!E{S['B'][0]})", "={c}{rev}*(1-g28_gm)",
         note="另一側讀數用情境 B 的毛利率；公司說法用 FY2028 毛利率指引中點（In 頁 g28_gm）")
    row5("ra", "FY2028 到期供應承諾 ÷ FY2028 銷貨成本（倍）", lambda s: "=h_sc_q3_28/{c}{cg}", "=h_sc_q3_28/{c}{cg}", "=h_sc_q3_28/{c}{cg}", fmt="0.00")
    row5("cg9", "FY2029 銷貨成本", lambda s: f"=PL!F{S[s][0]}-PL!F{S[s][1]}", f"=Coord!C{co29}*(1-PL!F{S['B'][1]}/PL!F{S['B'][0]})", "={c}{cg}",
         note="公司說法欄：無 FY2029 展望，假設與 FY2028 持平（Assumed；持平使正常水位偏低、超額偏大）")
    row5("i28", "FY2028 年底存貨（FY2027 與 FY2028 承諾全額執行）", lambda s: "=$B${i27f}+h_sc_q3_28/sc_cover-{c}{cg}", "=$B${i27f}+h_sc_q3_28/sc_cover-{c}{cg}", "=$B${i27f}+h_sc_q3_28/sc_cover-{c}{cg}")
    row5("dio", "　存貨天數（對 FY2029 銷貨成本）", lambda s: "={c}{i28}/{c}{cg9}*365", "={c}{i28}/{c}{cg9}*365", "={c}{i28}/{c}{cg9}*365", fmt="#,##0")
    row5("xf", "超出正常水位的存貨：全額執行", lambda s: "={c}{i28}-$B${dn}*{c}{cg9}", "={c}{i28}-$B${dn}*{c}{cg9}", "={c}{i28}-$B${dn}*{c}{cg9}",
         note="正值＝該路徑下承諾多於需求；負值＝承諾不足以支應該路徑")
    row5("xt", "超出正常水位的存貨：Q4 FY27 只照趨勢（未執行部分視為取消）", lambda s: "={c}{xf}-($B${i27f}-$B${i27t})", "={c}{xf}-($B${i27f}-$B${i27t})", "={c}{xf}-($B${i27f}-$B${i27t})")
    row5("xs", "　全額執行的超額 ÷ FY2028 以前到期承諾（69＋203）", lambda s: "={c}{xf}/(h_sc_q3_27r+h_sc_q3_28)", "={c}{xf}/(h_sc_q3_27r+h_sc_q3_28)", "={c}{xf}/(h_sc_q3_27r+h_sc_q3_28)", fmt="0.0%",
         note="正值時＝要讓存貨回到正常水位，需取消或改期的承諾比例")
    row5("eps", "FY2028 EPS（模型；美元）", lambda s: f"=PL!E{S[s][2]}", '="—"', '="—"')
    for k, col, lab in (("l1", "B", "低"), ("l2", "C", "中"), ("l3", "D", "高")):
        row5(k, f"提列損失（{lab}）：稅後每股（美元；全額執行的超額為底，負值不計）", lambda s, col=col: f"=MAX(0,{{c}}{{xf}})*${col}${{tx}}*(1-tx_f)/sh_v",
             f"=MAX(0,{{c}}{{xf}})*${col}${{tx}}*(1-tx_f)/sh_v", f"=MAX(0,{{c}}{{xf}})*${col}${{tx}}*(1-tx_f)/sh_v")
    row5("l3p", "　高案稅後每股 ÷ FY2028 EPS", lambda s: "={c}{l3}/{c}{eps}", '="—"', '="—"', fmt="0.0%")
    r0 += 1
    put("h4", "I4. 承諾可支撐的 FY2028 營收（年底存貨回到正常水位；FY2029 銷貨成本與 FY2028 持平）", ("情境 A 毛利率", "情境 B 毛利率", "情境 C 毛利率", "—", "毛利率指引中點"), bold=True)
    gm = [f"PL!E{S[s][1]}/PL!E{S[s][0]}" for s in "ABC"] + [None, "g28_gm"]

    def row4(key, label, inv, fmt="#,##0.0", ratio=False):
        vals = []
        for i, g in enumerate(gm):
            if g is None:
                vals.append('="—"'); continue
            f = f"=({inv}+h_sc_q3_28/sc_cover)/(1+$B${R['dn']})/(1-{g})"
            vals.append(f)
        put(key, label, vals, fmt=fmt)
    row4("s_f", "可支撐營收：FY2027 剩餘承諾全額執行", f"$B${R['i27f']}")
    row4("s_t", "可支撐營收：Q4 FY27 只照趨勢", f"$B${R['i27t']}")
    put("s_fr", "　全額執行 ÷ 該情境模型營收 − 1", [f"={c}{R['s_f']}/{c}{R['rev']}-1" for c in "BCD"] + ['="—"', f"=F{R['s_f']}/F{R['rev']}-1"], fmt="0.0%",
        note="正值＝承諾對應的營收高於該路徑；兩種解釋不能同時成立：模型營收偏低，或公司備料偏多（I3 的超額）")
    put("s_tr", "　趨勢案 ÷ 該情境模型營收 − 1", [f"={c}{R['s_t']}/{c}{R['rev']}-1" for c in "BCD"] + ['="—"', f"=F{R['s_t']}/F{R['rev']}-1"], fmt="0.0%")
    r0 += 1
    put("ck", "檢查：Q3 存貨三項合計 − 39.724；FY2027＋FY2028＋FY2029＋以後 − 390（應為 0）",
        ("=h_inv_rm_q3+h_inv_wip_q3+h_inv_fg_q3-39.724", "=h_sc_q3_27r+h_sc_q3_28+h_sc_q3_29+h_sc_q3_30p-h_sc_q3"), fmt="0.000")
    put("wf", "附：在製品＋製成品 Q3 季增（合併單 r3 的 ΔWF；門檻 +6）", ("=h_inv_wip_q3+h_inv_fg_q3-h_inv_wip_q2-h_inv_fg_q2",), fmt="0.000",
        note="8-K EX-99.3 已揭露存貨三項，不必等 10-Q。低於 +6 → 依 20261116 合併單預登錄規則，R2 判定由條件式成立轉為成立（由 R 線處理，不在本模型裁決）")

    # 佔位符解析：{r}＝本列；{key}＝該鍵的列號
    for rr in range(R["t"], r0):
        for cc in range(2, 8):
            v = RC.cell(rr, cc).value
            if isinstance(v, str) and v.startswith("=") and "{" in v:
                v = v.replace("{r}", str(rr))
                for k, n in R.items():
                    v = v.replace("{" + k + "}", str(n))
                assert "{" not in v, (rr, cc, v)
                RC.cell(rr, cc).value = v
        RC.cell(rr, 1).alignment = WRAP

    # ---------- README：風險登錄兩列改文字（列號不變） ----------
    rs = _find(RD, 1, '="供應與產能承諾 ', exact=False)
    RD.cell(rs, 1).value = ('="供應與產能承諾 "&TEXT(h_sc_q3,"#,##0")&" 十億美元（Q3 FY27；前兩季 "&TEXT(h_supply_c_q1,"#,##0")&"、"&TEXT(h_supply_c,"#,##0")'
                            '&"；其中 FY2028 到期 "&TEXT(h_sc_q3_28,"#,##0")&"）：未進 FCFF（隨採購認列進成本）；v0.3.13 起 Recon I 節以採購恆等式對照各情境銷貨成本，情境 B 下 FY2028 到期承諾為銷貨成本的 "'
                            f'&TEXT(Recon!C{R["ra"]},"0.00")&" 倍，只作對照與敏感度、不進目標價。觀察：FY27 Q4 年底存貨（全額執行約 "&TEXT(Recon!B{R["i27f"]},"0.0")&"、趨勢約 "&TEXT(Recon!B{R["i27t"]},"0.0")&"）、承諾到期表、存貨提列與超額採購義務應計"')
    rf = _find(RD, 1, "客戶融資與延長付款條件", exact=False)
    RD.cell(rf, 1).value = ('="客戶融資、擔保與延長付款條件：未建模（不進 FCFF 與淨現金）。已揭露金額（Q3 FY27）：擔保上限 "&TEXT(h_gu_lps_q3+h_gu_sbe_q3+h_gu_pwr_q3,"0.0")'
                            '&" 十億美元（SB Energy／OpenAI "&TEXT(h_gu_sbe_q3,"0")&"，逐棟完工才生效、首期預計 FY2029）、新雲協議承諾 "&TEXT(h_cm_neo_q3,"0")&"、應收帳款 "&TEXT(h_ar_q3,"0.0")'
                            '&"（DSO 64 天）；本季非上市權益證券減損與未實現損失 "&TEXT(h_eq_imp_q3,"0.0")&"。觀察：擔保是否動用、DSO、權益證券減損、10-Q Note 10"')
    for rr in (rs, rf):
        RD.cell(rr, 1).alignment = WRAP

    # ---------- README：沿革列 ----------
    r = RD.max_row + 1
    RD.cell(r, 1, "v0.3.13（2026-11-21）：只新增對照與揭露，目標價、評等、EPS 不變。(1) Recon I 節「供應承諾對採購恆等式」：以 8-K（2026-11-17）揭露的供應承諾到期表（合計 390；FY2028 到期 203）與存貨三項，"
               "對照各情境、另一側讀數與公司說法的 FY2028 銷貨成本，算年底存貨、超出正常水位的存貨、提列敏感度與承諾可支撐的營收；含 FY27 Q4 年底存貨的預登錄檢驗點。"
               "(2) Hist 頁末新增 Q3 FY27 承諾、擔保、存貨與相關已公布數字。(3) 風險登錄兩列改寫（供應承諾、客戶融資與擔保；後者原寫金額未揭露，已過時）。"
               "Q3 FY27 損益錨點滾入不在本版。依據：Andy 2026-11-21；對帳單 20261121_對帳單_NVIDIA供應承諾對模型銷貨成本。").alignment = WRAP
    return f"v0313_post：Hist 追加 {len(HIST)} 列；Recon I 節 {R['t']}–{r0 - 1}；README 風險登錄 {rs}、{rf} 改寫；沿革 README!A{r}"
