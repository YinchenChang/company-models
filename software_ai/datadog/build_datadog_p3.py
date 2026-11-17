# build_datadog.py 的第三部分（DCF、Val、RevDCF）；由 build_datadog_p2.py exec，共用命名空間。
SC3 = {"A": "B", "B": "C", "C": "D"}
S3 = lambda f: {SC3[s]: f(s) for s in "ABC"}

# 補 PL 其他收入淨額（預測年：現金收益 − 票據發行成本攤銷；現金取前一年底，避免循環）
for s in "ABC":
    for c in FCOLS:
        PL[f"{c}{P_[s]['onet']}"] = f"=y_cash*CF!{prevc(c)}{C_[s]['cash']}-int_exp"
        PL[f"{c}{P_[s]['onet']}"].number_format = "#,##0.0"

# ====================================================================================================== DCF（FCFF；評價日 2026-09-30；期中折現）
DC["A1"] = ("DCF — 無槓桿自由現金流（FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；NOPAT＝GAAP 營業利益 ×(1 − 稅率) ＋ sbc_add × 股票薪酬；預設股票薪酬視為費用、不加回）。"
            "Q4 26E 為 stub（評價日至年底），FY2027–FY2032 期中折現；終值＝NOPAT(FY2033)×(1 − g÷RONIC)÷(WACC − g)，RONIC＝WACC，終值折現年數＝stub＋6。")
DC["A2"] = ("每情境另算一組「加回股票薪酬（稅後）」的 FCFF 與現值，供 Val 頁落差分解使用（不影響主值）；FY2026 起利息收入不在 FCFF 內（現金另加在股權價值）。")
DC["A3"] = "Q4 26E stub 年數（評價日至年底 ÷ 365）"; DC["B3"] = "=(DATE(2026,12,31)-val_date)/365"; DC["B3"].number_format = "0.0000"; dn("t_stub", "DCF!$B$3")
DC["A4"] = "終值折現年數（stub ＋ 6）"; DC["B4"] = "=t_stub+6"; DC["B4"].number_format = "0.0000"; dn("t_tv", "DCF!$B$4")
DR_ = {}
r = 6
DCOLS = ["G"] + FCOLS
for s in "ABC":
    r = blk_header(DC, r, f"情境 {s}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s] }）")
    R = {k: r + i for i, k in enumerate(["yr", "ebit", "taxr", "nopat", "da", "cx", "dwc", "fcff", "t", "df", "pv", "nops", "fcffs", "pvs"])}
    PR = P_[s]; CR = C_[s]
    put(DC, R["yr"], "年度（數值）", "年", {c: (2026 if c == "G" else YR_OF[c]) for c in DCOLS}, "0")
    put(DC, R["ebit"], "營業利益（GAAP）", "US$M", {c: f"=PL!{c}{PR['oi']}" for c in DCOLS})
    put(DC, R["taxr"], "稅率", "%", {c: f"=PL!{c}{PR['taxr']}" for c in DCOLS}, "0.0%")
    put(DC, R["nopat"], "NOPAT＝營業利益 ×(1 − 稅率) ＋ sbc_add × 股票薪酬（加回為全額，稅盾已在營業利益內）", "US$M", {c: f"={c}{R['ebit']}*(1-{c}{R['taxr']})+sbc_add*PL!{c}{PR['sbc']}" for c in DCOLS})
    put(DC, R["da"], "＋折舊攤銷", "US$M", {c: f"=CF!{c}{CR['da']}" for c in DCOLS})
    put(DC, R["cx"], "− 資本支出（含資本化軟體）", "US$M", {c: f"=CF!{c}{CR['cx']}" for c in DCOLS})
    put(DC, R["dwc"], "− 營運資金增加（負＝流入）", "US$M", {c: f"=CF!{c}{CR['dwc']}" for c in DCOLS})
    put(DC, R["fcff"], "FCFF", "US$M", {c: f"={c}{R['nopat']}+{c}{R['da']}-{c}{R['cx']}-{c}{R['dwc']}" for c in DCOLS}, bold=True)
    put(DC, R["t"], "折現年數（期中）", "年", {"G": "=t_stub/2", "I": "=t_stub+0.5", **{c: f"={prevc(c)}{R['t']}+1" for c in FCOLS[1:]}}, "0.000")
    put(DC, R["df"], "折現因子 1÷(1＋WACC)^年數", "x", {c: f"=1/(1+wacc)^{c}{R['t']}" for c in DCOLS}, "0.0000")
    put(DC, R["pv"], "現值", "US$M", {c: f"={c}{R['fcff']}*{c}{R['df']}" for c in DCOLS})
    put(DC, R["nops"], "NOPAT（加回全額股票薪酬；落差分解用）", "US$M", {c: f"={c}{R['ebit']}*(1-{c}{R['taxr']})+PL!{c}{PR['sbc']}" for c in DCOLS})
    put(DC, R["fcffs"], "FCFF（加回股票薪酬）", "US$M", {c: f"={c}{R['nops']}+{c}{R['da']}-{c}{R['cx']}-{c}{R['dwc']}" for c in DCOLS})
    put(DC, R["pvs"], "現值（加回股票薪酬）", "US$M", {c: f"={c}{R['fcffs']}*{c}{R['df']}" for c in DCOLS})
    DR_[s] = R
    r += len(R) + 1
widths(DC)

# ====================================================================================================== Val
V["A1"] = '="Val — 目標價（股價基準日＋12 個月＝"&TEXT(tp_date,"yyyy-mm-dd")&"）＝DCF "&TEXT(w_dcf,"0%")&"＋Comparable "&TEXT(1-w_dcf,"0%")&"；情境權重 A "&TEXT(tp_wA,"0%")&"／B "&TEXT(tp_wB,"0%")&"／C "&TEXT(tp_wC,"0%")&"。金額 US$M，每股美元"'
V["A2"] = ("DCF：FCFF（股票薪酬視為費用、不加回）以 WACC 折現到評價日（期中），每股價值以 ke 推到目標時點；評價股數含未歸屬 RSU（扣預期沒收）與選擇權庫藏股法，2029 票據預設視為債務。"
           "Comparable（G6）：主腿為 P/S——同業中位 P/S（peer_ps_20261113，2026-11-13 現價，市值÷最新 TTM 營收）× rel_m × 本模型目標時點 TTM 營收 ÷ 稀釋股數，與 PS／PSG 相對值同口徑；盈餘型倍數（Forward P/E）只列對照，不進加權目標價；相對倍數 rel_m 取近 20 財季本公司 ÷ 同業中位的 PS 與 PSG 相對值幾何平均（Derived）。"
           "五級評等：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空；先看收支關卡。")
VR = {}
r = 4
V.cell(r, 1, "0.3 結論").font = BOLD; r += 1
hdr(V, r, ["項目", "數值", "對照", "對照", "說明"]); r += 1
for k in ("tpdcf", "psv_s", "tpdcf_s", "tpcomp", "tp", "px", "cons", "up", "upc", "tier", "valrate", "gate_chk", "gate_in", "rec", "scen_t", "scen_up", "gap_hint"):
    VR[k] = r; r += 1
r += 1
V.cell(r, 1, "0.2 折現率（CAPM；Blume β；稅後債務成本）").font = BOLD; r += 1
for k, lab in (("emkt", "股權市值 E＝現價 × 流通股（US$M）"), ("dmkt", "債務 D＝票據本金＋融資租賃（US$M）"), ("we", "股權權重"), ("rf", "無風險利率（10 年期美國公債）"), ("beta", "β（Blume 調整、兩來源中位數）"),
               ("erp", "股權風險溢酬"), ("ke", "股權成本 ke＝rf＋β × ERP"), ("kd", "稅前債務成本＝rf＋利差（票面 0%，採經濟成本）"), ("kdat", "稅後債務成本"), ("wacc", "WACC"),
               ("ronic", "RONIC＝WACC＋ronic_sp"), ("g", "永續成長率 g"), ("t_tgt", "評價日至目標時點年數")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "0.1 DCF 彙總（三情境）").font = BOLD; r += 1
VR["dh"] = r; hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
for k, lab in (("pvx", "明確期 FCFF 現值（含 Q4 26E stub）"), ("nop33", "NOPAT(FY2033)＝FY2032 NOPAT ×(1＋g)"), ("tv", "終值＝NOPAT(FY2033)×(1 − g÷RONIC)÷(WACC − g)"), ("pvtv", "終值現值"),
               ("ev", "企業價值 EV"), ("tvsh", "終值占 EV"), ("nc", "＋淨現金（現金＋有價證券 − 債務本金 − 融資租賃；票據視為股份時不扣票據）"), ("eqv", "股權價值（評價日）"),
               ("shv", "評價股數（流通股＋RSU×(1 − 沒收率)＋選擇權庫藏股法＋轉換股）"), ("psv", "每股價值（評價日，美元）"), ("dcf12", "DCF 情境價值推到目標時點（× (1＋ke)^年數）"),
               ("ev_s", "（落差分解用）EV：加回股票薪酬"), ("dcf12_s", "（落差分解用）DCF 情境價值：加回股票薪酬（推到目標時點）")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "0.0 Comparable（同業 Forward 倍數 × 目標時點 NTM；每股，目標時點）").font = BOLD; r += 1
VR["ch"] = r; hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
for k, lab in (("fmix", "權重：目標年度（FY2027）剩餘比例＝(該年 12-31 − 目標時點) ÷ 365"), ("ttm_rev", "目標時點 TTM 營收（US$M）＝FY2026 與 FY2027 加權（主腿分母）"), ("ttm_sh", "目標時點 TTM 稀釋加權股數（百萬股）"), ("ntm_rev", "目標時點 NTM 營收（US$M）＝FY2027 與 FY2028 加權"), ("ntm_eps", "目標時點 NTM non-GAAP EPS（美元）"),
               ("ntm_sh", "目標時點 NTM 稀釋加權股數（百萬股）"), ("m_ps", "同業 P/S 中位數（五家；2026-11-13 現價，市值÷最新 TTM 營收；股權口徑）"), ("m_pe", "對照：同業 Forward P/E 中位數（五家；分母期間未載明）"),
               ("px_ps", "P/S 腿每股＝同業 P/S 中位 × TTM 營收 ÷ TTM 股數（未乘相對倍數）"), ("comp", "Comparable 每股（主腿）＝rel_m × P/S 腿"), ("px_pe", "對照列：P/E 腿每股（rel_m＝1.00；同業 Forward P/E 中位 × NTM non-GAAP EPS；不進加權）"), ("px_pe_c", "對照列：P/E 腿每股 × 本公司現行 Forward P/E ÷ 同業中位（本公司現行相對值；不進加權）")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "情境目標價與收支關卡").font = BOLD; r += 1
for k, lab in (("t_A", "情境 A 目標價"), ("t_B", "情境 B 目標價"), ("t_C", "情境 C 目標價"), ("wsum", "情境權重合計（應為 1）"), ("fcf_gate", "收支關卡（基準情境 B：FY2026–FY2030 自由現金流合計，US$M）"),
               ("gate_A", "情境 A 收支關卡合計"), ("gate_C", "情境 C 收支關卡合計"), ("ai_net", "AI 淨效果（FY2032E，Thesis 頁；US$M）A／B／C"), ("tvc", "檢查：加權目標價 − (DCF 權重 × DCF ＋ Comparable 權重 × Comparable)（應為 0）")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "0.4 DCF 與 Comparable 落差分解（情境 B，每股美元；兩法皆在目標時點）").font = BOLD; r += 1
for k, lab in (("g_tot", "落差＝Comparable − DCF"), ("g_sbc", "其中：股票薪酬（DCF 加回股票薪酬後的增加量；Comparable 倍數隱含 non-GAAP 口徑）"), ("g_mult", "其中：倍數與成長殘差（Comparable − 加回股票薪酬的 DCF）"),
               ("g_lev", "備註：槓桿（兩法皆為股權每股價值，淨現金已含在兩法內；每股淨現金如右）"), ("g_time", "備註：時點（DCF 評價日至目標時點以 ke 滾動的增加量；兩法皆在目標時點，不另計入落差）"),
               ("g_chk", "檢查：股票薪酬 ＋ 殘差 − 落差（應為 0）"), ("g_evx", "DCF 隱含 EV 占 NTM 營收（倍）"), ("g_cmx", "Comparable 隱含股權價值占 TTM 營收（倍；＝rel_m × P/S 腿，主腿）")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "0.5 本公司現行倍數 ÷ 同業中位數與隱含重估（原則第 47 條）").font = BOLD; r += 1
for k, lab in (("c_ps", "本公司現行 Forward P/S（資料商同日同口徑）／同業中位數／比值／隱含重估（同業中位數 ÷ 現行 − 1）"), ("c_pe", "本公司現行 Forward P/E／同業中位數／比值／隱含重估"),
               ("c_den", "資料商 Forward P/S 隱含營收分母（US$M）＝(現價 × 流通股) ÷ 本公司 Forward P/S；對照模型 FY2026E／FY2027E／FY2028E 營收")):
    VR[k] = r; V.cell(r, 1, lab); r += 1
r += 1
V.cell(r, 1, "0.6 同業倍數逐家與逾期核對（原則第 15 條；與現價基準日相差超過 2 個交易日標逾期；皆為美國會計準則 US GAAP，租賃處理資料商未載明＝未找到）").font = BOLD; r += 1
VR["ph"] = r; hdr(V, r, ["項目", "P/S（主腿；2026-11-13 現價）", "Forward P/E（對照）", "P/S 價格日", "狀態／說明"]); r += 1
for k, nm, dt in PEER:
    VR["p_" + k] = r; V.cell(r, 1, nm + "（US GAAP；租賃處理未找到）"); r += 1
VR["p_ddog"] = r; V.cell(r, 1, "Datadog 現行（資料商，只作對照；US GAAP）"); r += 1
VR["o_rf"] = r; V.cell(r, 1, "無風險利率基準日"); r += 1
VR["o_cons"] = r; V.cell(r, 1, "共識目標價基準日"); r += 1

b = lambda k: f"B{VR[k]}"
vf = lambda ws_cell, f, fmt="#,##0.00": None
def vset(cell, f, fmt="#,##0.00"):
    V[cell] = f; V[cell].number_format = fmt
vset(b("emkt"), "=px*sh_out", "#,##0"); vset(b("dmkt"), "=debt_0+fin_lease", "#,##0"); vset(b("we"), f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})", "0.0%")
vset(b("rf"), "=rf", "0.00%"); vset(b("beta"), "=beta", "0.000"); vset(b("erp"), "=erp", "0.0%")
vset(b("ke"), f"=B{VR['rf']}+B{VR['beta']}*B{VR['erp']}", "0.00%"); dn("ke", f"Val!$B${VR['ke']}")
vset(b("kd"), f"=B{VR['rf']}+kd_sp", "0.00%"); vset(b("kdat"), f"=B{VR['kd']}*(1-tax_lr)", "0.00%")
vset(b("wacc"), f"=B{VR['we']}*B{VR['ke']}+(1-B{VR['we']})*B{VR['kdat']}", "0.00%"); dn("wacc", f"Val!$B${VR['wacc']}")
vset(b("ronic"), f"=B{VR['wacc']}+ronic_sp", "0.00%"); vset(b("g"), "=g_tv", "0.0%"); vset(b("t_tgt"), "=(tp_date-val_date)/365", "0.000"); dn("t_tgt", f"Val!$B${VR['t_tgt']}")

for s in "ABC":
    c = SC3[s]; D = DR_[s]; PR = P_[s]
    vset(f"{c}{VR['pvx']}", f"=SUM(DCF!G{D['pv']}:N{D['pv']})", "#,##0.0")
    vset(f"{c}{VR['nop33']}", f"=DCF!N{D['nopat']}*(1+$B${VR['g']})", "#,##0.0")
    vset(f"{c}{VR['tv']}", f"={c}{VR['nop33']}*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})", "#,##0.0")
    vset(f"{c}{VR['pvtv']}", f"={c}{VR['tv']}/(1+$B${VR['wacc']})^t_tv", "#,##0.0")
    vset(f"{c}{VR['ev']}", f"={c}{VR['pvx']}+{c}{VR['pvtv']}", "#,##0.0")
    vset(f"{c}{VR['tvsh']}", f"={c}{VR['pvtv']}/{c}{VR['ev']}", "0.0%")
    vset(f"{c}{VR['nc']}", "=cash_0-(1-conv_if)*debt_0-fin_lease", "#,##0.0")
    vset(f"{c}{VR['eqv']}", f"={c}{VR['ev']}+{c}{VR['nc']}", "#,##0.0")
    vset(f"{c}{VR['shv']}", "=sh_out+dil_inc*(rsu_unv*(1-rsu_fo)+MAX(0,opt_n*(1-opt_k/px)))+conv_if*conv_sh", "#,##0.000")
    vset(f"{c}{VR['psv']}", f"={c}{VR['eqv']}/{c}{VR['shv']}", "#,##0.00")
    vset(f"{c}{VR['dcf12']}", f"=MAX(0,{c}{VR['psv']})*(1+$B${VR['ke']})^t_tgt", "#,##0.00")
    vset(f"{c}{VR['ev_s']}", f"=SUM(DCF!G{D['pvs']}:N{D['pvs']})+DCF!N{D['nops']}*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^t_tv", "#,##0.0")
    vset(f"{c}{VR['dcf12_s']}", f"=MAX(0,({c}{VR['ev_s']}+{c}{VR['nc']})/{c}{VR['shv']})*(1+$B${VR['ke']})^t_tgt", "#,##0.00")
    # Comparable
    vset(f"{c}{VR['fmix']}", "=(DATE(YEAR(tp_date),12,31)-tp_date)/365", "0.000")
    vset(f"{c}{VR['ttm_rev']}", f"={c}{VR['fmix']}*PL!H{PR['rev']}+(1-{c}{VR['fmix']})*PL!I{PR['rev']}", "#,##0.0")
    vset(f"{c}{VR['ttm_sh']}", f"={c}{VR['fmix']}*PL!H{PR['dsh']}+(1-{c}{VR['fmix']})*PL!I{PR['dsh']}", "#,##0.0")
    vset(f"{c}{VR['ntm_rev']}", f"={c}{VR['fmix']}*PL!I{PR['rev']}+(1-{c}{VR['fmix']})*PL!J{PR['rev']}", "#,##0.0")
    vset(f"{c}{VR['ntm_eps']}", f"={c}{VR['fmix']}*PL!I{PR['ngeps']}+(1-{c}{VR['fmix']})*PL!J{PR['ngeps']}", "#,##0.00")
    vset(f"{c}{VR['ntm_sh']}", f"={c}{VR['fmix']}*PL!I{PR['dsh']}+(1-{c}{VR['fmix']})*PL!J{PR['dsh']}", "#,##0.0")
    vset(f"{c}{VR['m_ps']}", f"=B{VR['m_ps']}", "0.00") if False else None
PSQ_L = ",".join(f"pm_psq_{k}" for k, _, _ in PEER); PS_L = ",".join(f"pm_ps_{k}" for k, _, _ in PEER); PE_L = ",".join(f"pm_pe_{k}" for k, _, _ in PEER)
vset(f"B{VR['m_ps']}", f"=IF(ps_db=1,MEDIAN({PSQ_L}),MEDIAN({PS_L}))", "0.00"); vset(f"B{VR['m_pe']}", f"=MEDIAN({PE_L})", "0.00")
for s in "ABC":
    c = SC3[s]
    vset(f"{c}{VR['px_ps']}", f"=$B${VR['m_ps']}*{c}{VR['ttm_rev']}/{c}{VR['ttm_sh']}", "#,##0.00")
    vset(f"{c}{VR['px_pe']}", f"=$B${VR['m_pe']}*{c}{VR['ntm_eps']}", "#,##0.00")
    vset(f"{c}{VR['comp']}", f"=rel_m*{c}{VR['px_ps']}", "#,##0.00")
    vset(f"{c}{VR['px_pe_c']}", f"={c}{VR['px_pe']}*$D${VR['c_pe']}", "#,##0.00")
    vset(f"B{VR['t_' + s]}", f"=w_dcf*{c}{VR['dcf12']}+(1-w_dcf)*{c}{VR['comp']}", "#,##0.00")
    vset(f"B{VR['gate_' + s] if s != 'B' else VR['fcf_gate']}".replace("BB", "B") if False else (f"B{VR['gate_' + s]}" if s != "B" else f"B{VR['fcf_gate']}"), f"=SUM(CF!H{C_[s]['fcf']}:L{C_[s]['fcf']})", "#,##0.0")
    vset(f"{c}{VR['ai_net']}", f"=Thesis!N{T[s]['net']}", "#,##0.0")
W3 = lambda row: f"tp_wA*B{row}+tp_wB*C{row}+tp_wC*D{row}"
vset(f"B{VR['tpdcf']}", "=" + W3(VR["dcf12"])); vset(f"B{VR['tpdcf_s']}", "=" + W3(VR["dcf12_s"]))
vset(f"B{VR['psv_s']}", "=" + "+".join(f"tp_w{s}*MAX(0,({c}{VR['ev_s']}+{c}{VR['nc']})/{c}{VR['shv']})" for s, c in (("A", "B"), ("B", "C"), ("C", "D")))); vset(f"B{VR['tpcomp']}", "=" + W3(VR["comp"]))
vset(f"B{VR['tp']}", f"=w_dcf*B{VR['tpdcf']}+(1-w_dcf)*B{VR['tpcomp']}"); vset(f"B{VR['px']}", "=px")
vset(f"B{VR['cons']}", "=cons_tp", "#,##0.00")
vset(f"B{VR['up']}", f"=B{VR['tp']}/B{VR['px']}-1", "0.0%"); vset(f"B{VR['upc']}", f"=B{VR['tpcomp']}/B{VR['px']}-1", "0.0%")
U_ = f"B{VR['up']}"
V[f"B{VR['tier']}"] = f'=IF({U_}>0.5,"> +50%",IF({U_}>0.25,"+25% 至 +50%",IF({U_}>=-0.25,"−25% 至 +25%",IF({U_}>=-0.5,"−50% 至 −25%","< −50%"))))'
V[f"B{VR['valrate']}"] = f'=IF({U_}>0.5,"強力買進",IF({U_}>0.25,"買進",IF({U_}>=-0.25,"中立",IF({U_}>=-0.5,"放空","強力放空"))))'
vset(f"B{VR['gate_in']}", f"=CF!C{GATE}", "#,##0.0")
V[f"B{VR['gate_chk']}"] = f'=IF(B{VR["fcf_gate"]}>=0,"通過（FY2026 至 FY2030 自由現金流合計為正）","不通過（資金缺口，不得買進）")'
V[f"B{VR['rec']}"] = f'=IF(AND(B{VR["gate_in"]}<0,OR(B{VR["valrate"]}="強力買進",B{VR["valrate"]}="買進")),"收支不足",B{VR["valrate"]})'
for cc, k in zip("BCD", "ABC"):
    vset(f"{cc}{VR['scen_t']}", f"=B{VR['t_' + k]}", "#,##0.00"); vset(f"{cc}{VR['scen_up']}", f"={cc}{VR['scen_t']}/px-1", "0.0%")
vset(f"B{VR['wsum']}", "=ROUND(tp_wA+tp_wB+tp_wC,6)", "0.000")
vset(f"B{VR['tvc']}", f"=B{VR['tp']}-(w_dcf*B{VR['tpdcf']}+(1-w_dcf)*B{VR['tpcomp']})", "0.000")
# 落差分解
vset(f"B{VR['g_tot']}", f"=C{VR['comp']}-C{VR['dcf12']}"); vset(f"B{VR['g_sbc']}", f"=C{VR['dcf12_s']}-C{VR['dcf12']}")
vset(f"B{VR['g_mult']}", f"=C{VR['comp']}-C{VR['dcf12_s']}"); vset(f"B{VR['g_lev']}", f"=(cash_0-(1-conv_if)*debt_0-fin_lease)/C{VR['shv']}")
vset(f"B{VR['g_time']}", f"=C{VR['dcf12']}-MAX(0,C{VR['psv']})"); vset(f"B{VR['g_chk']}", f"=B{VR['g_sbc']}+B{VR['g_mult']}-B{VR['g_tot']}", "0.000")
vset(f"B{VR['g_evx']}", f"=C{VR['ev']}/C{VR['ntm_rev']}", "0.00")
vset(f"B{VR['g_cmx']}", f"=C{VR['comp']}*C{VR['ttm_sh']}/C{VR['ttm_rev']}", "0.00")
# 現行倍數 ÷ 同業中位數
vset(f"B{VR['c_ps']}", "=pm_ps_ddog", "0.00"); vset(f"C{VR['c_ps']}", f"=B{VR['m_ps']}", "0.00"); vset(f"D{VR['c_ps']}", f"=B{VR['c_ps']}/C{VR['c_ps']}", "0.00"); vset(f"E{VR['c_ps']}", f"=C{VR['c_ps']}/B{VR['c_ps']}-1", "+0.0%;-0.0%")
vset(f"B{VR['c_pe']}", "=pm_pe_ddog", "0.00"); vset(f"C{VR['c_pe']}", f"=B{VR['m_pe']}", "0.00"); vset(f"D{VR['c_pe']}", f"=B{VR['c_pe']}/C{VR['c_pe']}", "0.00"); vset(f"E{VR['c_pe']}", f"=C{VR['c_pe']}/B{VR['c_pe']}-1", "+0.0%;-0.0%")
vset(f"B{VR['c_den']}", f"=px*sh_out/pm_ps_ddog", "#,##0"); vset(f"C{VR['c_den']}", f"=PL!H{P_['B']['rev']}", "#,##0"); vset(f"D{VR['c_den']}", f"=PL!I{P_['B']['rev']}", "#,##0"); vset(f"E{VR['c_den']}", f"=PL!J{P_['B']['rev']}", "#,##0")
# 同業逐家
for k, nm, dt in PEER:
    rr = VR["p_" + k]
    vset(f"B{rr}", f"=pm_ps_{k}", "0.00"); vset(f"C{rr}", f"=pm_pe_{k}", "0.00"); vset(f"D{rr}", f"=pm_dp_{k}", "yyyy-mm-dd")
    V[f"E{rr}"] = f'=IF(NETWORKDAYS(D{rr},px_date)-1>2,"逾期","正常")'
rr = VR["p_ddog"]; vset(f"B{rr}", "=pm_ps_ddog", "0.00"); vset(f"C{rr}", "=pm_pe_ddog", "0.00"); vset(f"D{rr}", "=px_date", "yyyy-mm-dd"); V[f"E{rr}"] = "同日（2026-11-13）"
vset(f"B{VR['o_rf']}", "=rf_date", "yyyy-mm-dd"); V[f"C{VR['o_rf']}"] = f'=IF(NETWORKDAYS(B{VR["o_rf"]},px_date)-1>2,"逾期","正常")'
vset(f"B{VR['o_cons']}", "=cons_date", "yyyy-mm-dd"); V[f"C{VR['o_cons']}"] = f'=IF(NETWORKDAYS(B{VR["o_cons"]},px_date)-1>2,"逾期","正常")'
# 標籤與說明
NOTES = {"psv_s": ("DCF（股票薪酬加回）每股價值（情境加權，評價日；只作對照）", "FCFF 與終值 NOPAT 加回全額股票薪酬；不進加權目標價（工作單 G5 第 3 條）"),
         "tpdcf_s": ("DCF（股票薪酬加回）目標價（情境加權；只作對照）", "每股價值以 ke 推到目標時點；不進加權目標價、不影響評等"),
         "tpdcf": ("DCF 目標價（情境加權）", "三情境 DCF 每股價值推到目標時點，依情境權重加權；股票薪酬視為費用"),
         "tpcomp": ("Comparable 目標價（P/S 主腿，情境加權）", "同業 P/S 中位（2026-11-13 現價）× rel_m × 本模型目標時點 TTM 營收 ÷ 稀釋股數；與 PS／PSG 相對值同口徑，盈餘型倍數只作對照"),
         "tp": ("加權目標價（DCF 與 Comparable 加權）", "DCF 與 Comparable 依 w_dcf 加權；投資建議依此判定"),
         "px": ("現價（2026-11-13 收盤）", "兩來源前收價相符；財報後（2026-11-04）單日大跌"),
         "cons": ("市場共識目標價（賣方平均）", "TipRanks 2026-11-12，高 335、低 171；只作對照，不進模型"),
         "up": ("上檔空間：加權目標價", ""), "upc": ("上檔空間：Comparable 目標價", ""),
         "tier": ("上檔空間級距", "五級固定規則：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空"),
         "valrate": ("估值評等（依加權目標價上檔空間，不含收支關卡）", "五級之一"),
         "gate_chk": ("收支關卡判定", "先看收支再看估值（個股模型原則第 10 節第 63 條）"),
         "gate_in": ("入不敷出判定用：基準情境 B 收支關卡合計（US$M）", "供 check_rating_gate 核對；負值時買進以上改判收支不足"),
         "rec": ("投資建議", "估值評等；收支關卡不通過且評等為買進以上時改為收支不足"),
         "scen_t": ("情境目標價 A／B／C（DCF 與 Comparable 加權）", ""), "scen_up": ("情境上檔空間 A／B／C", ""),
         "gap_hint": ("DCF 與 Comparable 落差", "")}
for k, (lab, note) in NOTES.items():
    V.cell(VR[k], 1, lab); V.cell(VR[k], 5, note)
V.cell(VR["gap_hint"], 1, "DCF 與 Comparable 落差（Comparable − DCF，情境 B；每股）"); vset(f"B{VR['gap_hint']}", f"=B{VR['g_tot']}", "+#,##0.00;-#,##0.00")
V.cell(VR["gap_hint"], 5, "分解見 0.4：股票薪酬與倍數成長殘差")
for k in ("tpdcf", "tpcomp", "tp", "up", "rec"):
    for cc in range(1, 6): V.cell(VR[k], cc).fill = KEY
V.cell(VR["g_lev"], 3, f"=B{VR['g_lev']}"); V.cell(VR["g_lev"], 3).number_format = "#,##0.00"
V.column_dimensions["A"].width = 70
for c_ in "BCD": V.column_dimensions[c_].width = 14
V.column_dimensions["E"].width = 70

# ====================================================================================================== RevDCF（封閉解：FY2028–FY2032 統一 GAAP 營業利益率 m*；情境 B 營收路徑）
RV["A1"] = "RevDCF — 現價隱含什麼（情境 B 營收路徑；以封閉解反解 FY2028–FY2032 統一 GAAP 營業利益率，其餘輸入不變）"
RV["A2"] = ("FCFF 對 FY2028–FY2032 營業利益率為線性，故以 m＝0 與 m＝1 兩個端點求 EV，線性內插到現價所需 EV；終值亦隨 FY2032 營業利益率線性變動。"
            "現價視為評價日價值（兩者相差約六週，不另折現）；同時列本公司現行 Forward 倍數相對同業中位數的隱含重估。")
RVR = {}
r = 4
for k, lab in (("px", "現價（美元）"), ("eqt", "現價隱含股權價值（US$M；現價 × 評價股數）"), ("evr", "現價隱含企業價值 EV（US$M；扣淨現金）"), ("ev0", "EV(m＝0)"), ("ev1", "EV(m＝1)"),
               ("m1", "現價隱含 FY2028–FY2032 GAAP 營業利益率"), ("m0", "模型情境 B FY2028–FY2032 平均 GAAP 營業利益率"), ("m0n", "參考：模型情境 B FY2032 non-GAAP 營業利益率"),
               ("rv0", "模型情境 B FY2028 營收（US$M）"), ("w0", "模型 WACC"), ("sps", "Comparable 目標價 ÷ 現價（情境 B；現價高於 Comparable 的程度）"),
               ("mps", "現行 Forward P/S ÷ 同業中位數"), ("mpe", "現行 Forward P/E ÷ 同業中位數"), ("rrps", "隱含重估（同業中位數 ÷ 現行 − 1）：P/S"), ("rrpe", "隱含重估：P/E")):
    RVR[k] = r; RV.cell(r, 1, lab); r += 1
RB = P_["B"]; DB = DR_["B"]
yrs_c = ["J", "K", "L", "M", "N"]
def ev_expr(m):
    parts = [f"SUM(DCF!G{DB['pv']}:I{DB['pv']})"]
    for c in yrs_c:
        parts.append(f"(({m})*PL!{c}{RB['rev']}*(1-PL!{c}{RB['taxr']})+DCF!{c}{DB['da']}-DCF!{c}{DB['cx']}-DCF!{c}{DB['dwc']})*DCF!{c}{DB['df']}")
    parts.append(f"({m})*PL!N{RB['rev']}*(1-PL!N{RB['taxr']})*(1+g_tv)*(1-g_tv/Val!$B${VR['ronic']})/(Val!$B${VR['wacc']}-g_tv)/(1+Val!$B${VR['wacc']})^t_tv")
    return "+".join(parts)
RV[f"B{RVR['px']}"] = "=px"; RV[f"B{RVR['px']}"].number_format = "#,##0.00"
RV[f"B{RVR['eqt']}"] = f"=px*Val!C{VR['shv']}"; RV[f"B{RVR['eqt']}"].number_format = "#,##0"
RV[f"B{RVR['evr']}"] = f"=B{RVR['eqt']}-Val!C{VR['nc']}"; RV[f"B{RVR['evr']}"].number_format = "#,##0"
RV[f"B{RVR['ev0']}"] = "=" + ev_expr("0"); RV[f"B{RVR['ev1']}"] = "=" + ev_expr("1")
for k in ("ev0", "ev1"): RV[f"B{RVR[k]}"].number_format = "#,##0"
RV[f"B{RVR['m1']}"] = f"=(B{RVR['evr']}-B{RVR['ev0']})/(B{RVR['ev1']}-B{RVR['ev0']})"; RV[f"B{RVR['m1']}"].number_format = "0.0%"
RV[f"B{RVR['m0']}"] = "=AVERAGE(" + ",".join(f"PL!{c}{RB['om']}" for c in yrs_c) + ")"; RV[f"B{RVR['m0']}"].number_format = "0.0%"
RV[f"B{RVR['m0n']}"] = f"=PL!N{RB['omn']}"; RV[f"B{RVR['m0n']}"].number_format = "0.0%"
RV[f"B{RVR['rv0']}"] = f"=PL!J{RB['rev']}"; RV[f"B{RVR['rv0']}"].number_format = "#,##0"
RV[f"B{RVR['w0']}"] = f"=Val!B{VR['wacc']}"; RV[f"B{RVR['w0']}"].number_format = "0.00%"
RV[f"B{RVR['sps']}"] = f"=Val!C{VR['comp']}/px"; RV[f"B{RVR['sps']}"].number_format = "0.00"
RV[f"B{RVR['mps']}"] = f"=Val!D{VR['c_ps']}"; RV[f"B{RVR['mpe']}"] = f"=Val!D{VR['c_pe']}"; RV[f"B{RVR['rrps']}"] = f"=Val!E{VR['c_ps']}"; RV[f"B{RVR['rrpe']}"] = f"=Val!E{VR['c_pe']}"
for k, f_ in (("mps", "0.00"), ("mpe", "0.00"), ("rrps", "+0.0%;-0.0%"), ("rrpe", "+0.0%;-0.0%")): RV[f"B{RVR[k]}"].number_format = f_
r += 1
RVR["sum"] = r
RV[f"A{r}"] = (f'=IF(B{RVR["m1"]}>1,"現價隱含的 FY2028–FY2032 GAAP 營業利益率超過 100%（"&TEXT(B{RVR["m1"]},"0%")&"），代表在模型營收路徑與 WACC 下，無論利潤率多高都無法由現金流折現支持現價；模型情境 B 平均 "&TEXT(B{RVR["m0"]},"0.0%")&"（FY2032 non-GAAP 營業利益率 "&TEXT(B{RVR["m0n"]},"0.0%")&"）；現價隱含 EV "&TEXT(B{RVR["evr"]},"#,##0")&" 百萬美元。","現價隱含 FY2028–FY2032 GAAP 營業利益率約 "&TEXT(B{RVR["m1"]},"0.0%")&"，模型情境 B 平均 "&TEXT(B{RVR["m0"]},"0.0%")&"；現價隱含 EV "&TEXT(B{RVR["evr"]},"#,##0")&" 百萬美元。")')
r += 1
RVR["sum2"] = r
RV[f"A{r}"] = (f'="現行 Forward P/S 為同業中位數的 "&TEXT(B{RVR["mps"]},"0.00")&" 倍、Forward P/E 為 "&TEXT(B{RVR["mpe"]},"0.00")&" 倍；若回到同業中位數，隱含重估 "&TEXT(B{RVR["rrps"]},"+0.0%;-0.0%")&"（P/S）與 "&TEXT(B{RVR["rrpe"]},"+0.0%;-0.0%")&"（P/E）。"')
RV.column_dimensions["A"].width = 90
exec(open(os.path.join(HERE, "build_datadog_p4.py"), encoding="utf-8").read())
