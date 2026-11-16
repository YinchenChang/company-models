# build_datadog.py 的第二部分（Thesis、PL、CF、DCF）；由 build_datadog.py exec，共用命名空間。
HN = set(wb.defined_names.keys())
FCOLS = ["I", "J", "K", "L", "M", "N"]
HISTC = {"C": "23", "D": "24", "E": "25"}
YR_OF = {c: y for c, y in zip("HIJKLMN", [2026] + [2000 + a for a in ANN])}


def blk_header(ws, r, title, fill=SUB):
    ws.cell(r, 1, title).font = BOLD
    for cc in range(1, NOTE_C + 1): ws.cell(r, cc).fill = fill
    for j, h in enumerate(MODEL_HDR, 1):
        c = ws.cell(r + 1, j, h); c.font = BOLD; c.fill = HDR
    return r + 2


def ex(n): return n in HN


# ====================================================================================================== Thesis（命題模組：用量單位 × 單價）
TH["A1"] = ("Thesis — 命題模組（原型 U 按用量計價）：營收＝Σ 客戶群（用量單位 × 單價）。非 AI 客戶、AI 原生客戶、AI 產品 credits 三條子列；用量成長由 FY2027 值線性收斂，單價依讓利路徑；"
            "AI 淨效果＝AI 營收增量 − AI 原生客戶價格讓利 − 推論成本（Tokenomics）。變數名稱與 software_ai/unit_economics 一致，該模型完成前先在 In 頁自建（v0.2 改讀 unit_economics）。")
TH["A2"] = ("單位毛利機制：銷貨成本隨「用量」（以 FY2026 價格計的營收）而非隨營收，單價讓利直接壓低毛利率；推論成本只由 AI credits 任務數 × 每任務 token × token 價格決定，token 價格只讀 Tokenomics 快照。"
            "每任務 token 由公司揭露校準（credit 價格 × 耗用 credits × 成本占比 ÷ 校準價格），不用聊天基數 × IF_TaskTokMult；後者只列 Recon 第 6 節對照。")
T = {}
r = 4
for s in "ABC":
    r = blk_header(TH, r, f"情境 {s}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s] }）：用量 × 單價與 AI 淨效果")
    R = {}
    def row(k, label, unit, vals, fmt="#,##0.0", note=None, bold=False):
        global r
        R[k] = r; put(TH, r, label, unit, vals, fmt, note, bold); r += 1
    def cols(fn, cs=FCOLS): return {c: fn(c) for c in cs}
    nxt = {}
    # 列號先預留（公式互相引用）
    keys = ["yr", "rev", "revN", "revA", "aic", "chk", "vN", "volN", "priN", "vA", "volA", "pA", "priA", "gai", "units", "core", "pxt", "crp", "tasks", "ctask", "inf", "rm", "inc", "eros", "inf2", "net", "nerosN", "pbe", "pact", "pbe_task", "pidx"]
    base = r
    for i, k in enumerate(keys): R[k] = base + i
    def rr(k): return R[k]
    P = lambda c: prevc(c)
    r = base
    put(TH, R["yr"], "年度（數值）", "年", {c: YR_OF[c] for c in "HIJKLMN"}, "0")
    put(TH, R["rev"], "營收合計（Datadog 平台，單一應報導部門）", "US$M",
        {"F": "=h_rev_9m26", "G": "=q_rev_3q26*(1+q4_g)", "H": f"=F{R['rev']}+G{R['rev']}", **cols(lambda c: f"={c}{R['revN']}+{c}{R['revA']}+{c}{R['aic']}")}, bold=True,
        note="FY2026E＝9M26 實績＋Q4 26E（3Q26 × (1＋季增率)）；FY2027 起＝三條子列加總")
    put(TH, R["revN"], "非 AI 原生客戶平台用量營收＝基期營收 × 用量指數 × 單價指數", "US$M",
        {"H": f"=(H{R['rev']}-ai_26)*(1-sa_26)", **cols(lambda c: f"=$H{R['revN']}*{c}{R['volN']}*{c}{R['priN']}")}, note="FY2026E 基期＝(營收 − AI credits)×(1 − AI 原生占比)")
    put(TH, R["revA"], "AI 原生客戶平台用量營收＝基期營收 × 用量指數 × 單價指數", "US$M",
        {"H": f"=(H{R['rev']}-ai_26)*sa_26", **cols(lambda c: f"=$H{R['revA']}*{c}{R['volA']}*{c}{R['priA']}")}, note="含最大客戶（公司未揭露占比）")
    put(TH, R["aic"], "AI 產品 credits 營收（Bits AI 等）", "US$M",
        {"H": "=ai_26", "I": f"=ai_27_{s}", **{c: f"={P(c)}{R['aic']}*(1+{c}{R['gai']})" for c in FCOLS[1:]}}, note="公司不揭露 AI 產品營收（未找到）；Assumed")
    put(TH, R["chk"], "檢查：營收合計 − 三條子列合計（應為 0）", "US$M", {c: f"={c}{R['rev']}-{c}{R['revN']}-{c}{R['revA']}-{c}{R['aic']}" for c in "HIJKLMN"}, "0.000")
    put(TH, R["vN"], "非 AI 客戶用量年增", "%", cols(lambda c: f"=vn_27_{s}+(vn_32-vn_27_{s})*({c}{R['yr']}-2027)/5"), "0.0%")
    put(TH, R["volN"], "非 AI 客戶用量指數（FY2026E＝1）", "x", {"H": 1, **cols(lambda c: f"={P(c)}{R['volN']}*(1+{c}{R['vN']})")}, "0.000")
    put(TH, R["priN"], "非 AI 客戶單價指數（價格保留率；FY2026E＝1）", "x", {"H": 1, **cols(lambda c: f"={P(c)}{R['priN']}*(1+pn_g)")}, "0.000")
    put(TH, R["vA"], "AI 原生客戶用量年增", "%", cols(lambda c: f"=va_27_{s}+(va_32-va_27_{s})*({c}{R['yr']}-2027)/5"), "0.0%")
    put(TH, R["volA"], "AI 原生客戶用量指數（FY2026E＝1；FY2027 乘 (1 − 最大客戶尾端下降比例)）", "x",
        {"H": 1, "I": f"=H{R['volA']}*(1+I{R['vA']})*(1-lc_cut)", **{c: f"={P(c)}{R['volA']}*(1+{c}{R['vA']})" for c in FCOLS[1:]}}, "0.000")
    put(TH, R["pA"], "AI 原生客戶單價年變動（讓利自 FY2027 起三年線性消退）", "%", cols(lambda c: f"=pa_{s}*MAX(0,1-({c}{R['yr']}-2027)/3)"), "0.0%")
    put(TH, R["priA"], "AI 原生客戶單價指數（價格保留率；FY2026E＝1）", "x", {"H": 1, **cols(lambda c: f"={P(c)}{R['priA']}*(1+{c}{R['pA']})")}, "0.000")
    put(TH, R["gai"], "AI credits 營收年增（FY2028 起）", "%", {c: f"=MAX(gai_floor,gai28-gai_step*({c}{R['yr']}-2028))" for c in FCOLS[1:]}, "0.0%")
    put(TH, R["units"], "平台用量（以 FY2026 價格計的營收；不含 AI credits，其成本由推論成本列計）", "US$M", {"H": f"=H{R['revN']}+H{R['revA']}", **cols(lambda c: f"=$H{R['revN']}*{c}{R['volN']}+$H{R['revA']}*{c}{R['volA']}")})
    put(TH, R["core"], "銷貨成本（non-GAAP，不含推論）＝用量 ×（1 − 毛利率）×（1＋單位成本年變動）^年數", "US$M", {c: f"={c}{R['units']}*(1-gm_core)*(1+uc_g)^({c}{R['yr']}-2026)" for c in "HIJKLMN"})
    put(TH, R["pxt"], "採用的 token 價格（路由組合 × 價格年變動；Tokenomics IF_PriceRef）", "US$/M", {c: f"=((1-route_sh)*tk_px+route_sh*tk_cheap)*(1+tp_dec)^({c}{R['yr']}-2026)" for c in "HIJKLMN"}, "0.000", note="價格只讀 Tokenomics 快照；年變動為 Assumed（待 AI供應鏈 登錄）")
    put(TH, R["crp"], "每 credit 價格（隨 token 價格變動轉嫁 pt_k）", "US$", {c: f"=cr_px*(1+pt_k*({c}{R['pxt']}/tk_px0-1))" for c in "HIJKLMN"}, "0.000")
    put(TH, R["tasks"], "AI 任務數（Bits Investigation 當量）", "百萬次", {c: f"={c}{R['aic']}/(cr_task*{c}{R['crp']})" for c in "HIJKLMN"}, "#,##0.0")
    put(TH, R["ctask"], "每任務推論成本＝每任務 token × token 價格", "US$", {c: f"=tk_tokens*{c}{R['pxt']}" for c in "HIJKLMN"}, "0.000")
    put(TH, R["inf"], "推論成本（US$M）", "US$M", {c: f"={c}{R['tasks']}*{c}{R['ctask']}" for c in "HIJKLMN"}, "#,##0.0", note="進 FY2027 起銷貨成本；FY2026E 已含在實績與 Q4 毛利率中，本列只作顯示")
    put(TH, R["rm"], "AI credits 轉售毛利率（1 − 每任務推論成本 ÷ 每任務 credit 收入）", "%", {c: f"=1-{c}{R['ctask']}/(cr_task*{c}{R['crp']})" for c in "HIJKLMN"}, "0.0%")
    put(TH, R["inc"], "AI 營收增量＝AI 原生客戶用量增量（以 FY2026 價格）＋AI credits 營收", "US$M", {c: f"=$H{R['revA']}*({c}{R['volA']}-1)+{c}{R['aic']}" for c in "HIJKLMN"}, bold=True)
    put(TH, R["eros"], "侵蝕：AI 原生客戶價格讓利＝基期營收 × 用量指數 ×（1 − 單價指數）", "US$M", {c: f"=$H{R['revA']}*{c}{R['volA']}*(1-{c}{R['priA']})" for c in "HIJKLMN"}, bold=True)
    put(TH, R["inf2"], "推論成本", "US$M", {c: f"={c}{R['inf']}" for c in "HIJKLMN"}, bold=True)
    put(TH, R["net"], "AI 淨效果（對營業利益前毛利）＝增量 − 侵蝕 − 推論成本", "US$M", {c: f"={c}{R['inc']}-{c}{R['eros']}-{c}{R['inf2']}" for c in "HIJKLMN"}, bold=True, fill=KEY)
    put(TH, R["nerosN"], "參考：非 AI 客戶單價讓利（不計入 AI 淨效果）", "US$M", {c: f"=$H{R['revN']}*{c}{R['volN']}*(1-{c}{R['priN']})" for c in "HIJKLMN"})
    put(TH, R["pbe"], "反解：AI 淨效果＝0 所需的 AI 原生客戶價格保留率", "x", {c: f"=($H{R['revA']}-{c}{R['aic']}+{c}{R['inf2']})/($H{R['revA']}*{c}{R['volA']})" for c in "HIJKLMN"}, "0.000", note="公式反解：(基期 − AI credits 營收 ＋ 推論成本) ÷ (基期 × 用量指數)；低於實際單價指數＝淨效果為正")
    put(TH, R["pact"], "對照：模型的 AI 原生客戶單價指數", "x", {c: f"={c}{R['priA']}" for c in "HIJKLMN"}, "0.000")
    put(TH, R["pbe_task"], "反解：AI credits 轉售毛利率＝0 的每任務 credit 收入（美元）＝每任務推論成本", "US$", {c: f"={c}{R['ctask']}" for c in "HIJKLMN"}, "0.000", note="對照每任務 credit 收入＝耗用 credits × 每 credit 價格")
    put(TH, R["pidx"], "平台加權單價指數＝平台營收 ÷ 平台用量", "x", {c: f"=({c}{R['revN']}+{c}{R['revA']})/{c}{R['units']}" for c in "HIJKLMN"}, "0.000")
    r = base + len(keys) + 1
    T[s] = R
widths(TH)

# ====================================================================================================== PL
PL["A1"] = "PL — 損益（US$M）：官方單一分部 Datadog 平台（子列＝命題拆分，合計等於分部）→ non-GAAP 營業利益 → GAAP（扣股票薪酬等）→ 稅 → 淨利與 EPS。三情境，末段為標準損益表（GAAP）。"
PL["A2"] = ("營業費用率與毛利率比率以近四季實績全期持平（G2 第 3 條）；毛利率只因單價讓利、AI 推論成本與單位成本機制而變。GAAP 營業利益＝non-GAAP − 股票薪酬 − 員工股票交易薪資稅 − 攤銷與併購費用。"
            "FY2026E＝9M26 實績＋Q4 26E。non-GAAP EPS 依公司口徑（稅前加回發行成本攤銷，21% 稅率）。")
P_ = {}
r = 4
SEGL = "Datadog 平台（單一應報導部門）"
LAB = {"N": "非 AI 原生客戶平台用量", "A": "AI 原生客戶平台用量", "I": "AI 產品 credits 營收"}
for s in "ABC":
    r = blk_header(PL, r, f"情境 {s}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s] }）")
    R = {}
    ks = ["yr", "seg", "sN", "sA", "sI", "rev", "yoy", "cng", "gpn", "gmn", "rdn", "smn", "gan", "oin", "omn", "sbc", "ptx", "oth", "oi", "om", "cogs", "gp", "gm", "opex", "onet", "pti", "taxr", "tax", "ni", "nm", "dsh", "eps", "ngpti", "ngni", "ngeps", "chk1", "chk2"]
    for i, k in enumerate(ks): R[k] = r + i
    TR = T[s]
    def pl(k, label, unit, vals, fmt="#,##0.0", note=None, bold=False, fill=None): put(PL, R[k], label, unit, vals, fmt, note, bold, fill)
    def hist(k, p_cols):   # 歷史欄：{欄: 名稱}
        return {c: f"=h_{k}_{p}" for c, p in p_cols.items() if ex(f"h_{k}_{p}")}
    H3 = {"C": "23", "D": "24", "E": "25", "F": "9m26"}
    pl("yr", "年度（數值）", "年", {c: YR_OF[c] for c in "HIJKLMN"}, "0")
    pl("seg", SEGL, "US$M", {**{c: f"={c}{R['rev']}" for c in "CDEFG"}, **{c: f"={c}{R['sN']}+{c}{R['sA']}+{c}{R['sI']}" for c in "HIJKLMN"}}, bold=True, note="單一應報導部門（10-K Note 2）；FY2026E 起以三條命題子列加總")
    pl("sN", LAB["N"], "US$M", {c: f"=Thesis!{c}{TR['revN']}" for c in "HIJKLMN"}, note="命題子列（Thesis 頁）")
    pl("sA", LAB["A"], "US$M", {c: f"=Thesis!{c}{TR['revA']}" for c in "HIJKLMN"})
    pl("sI", LAB["I"], "US$M", {c: f"=Thesis!{c}{TR['aic']}" for c in "HIJKLMN"})
    pl("rev", "合併營收", "US$M", {**hist("rev", H3), "G": f"=Thesis!G{TR['rev']}", "H": f"=F{R['rev']}+G{R['rev']}", **{c: f"=Thesis!{c}{TR['rev']}" for c in FCOLS}}, bold=True)
    pl("yoy", "營收年增率", "%", {"D": f"=D{R['rev']}/C{R['rev']}-1", "E": f"=E{R['rev']}/D{R['rev']}-1", "F": f"=F{R['rev']}/h_rev_9m25-1", "G": f"=G{R['rev']}/q_rev_4q25-1", "H": f"=H{R['rev']}/E{R['rev']}-1",
                                    **{c: f"={c}{R['rev']}/{prevc(c)}{R['rev']}-1" for c in FCOLS}}, "0.0%")
    pl("cng", "銷貨成本（non-GAAP；含 AI 推論成本）", "US$M", {**{c: f"={c}{R['rev']}-h_ngp_{p}" for c, p in H3.items() if ex(f"h_ngp_{p}")}, "G": f"=G{R['rev']}*(1-gm_core)", "H": f"=F{R['cng']}+G{R['cng']}",
                                                               **{c: f"=Thesis!{c}{TR['core']}+Thesis!{c}{TR['inf']}" for c in FCOLS}})
    pl("gpn", "毛利（non-GAAP）", "US$M", {c: f"={c}{R['rev']}-{c}{R['cng']}" for c in "CDEFGHIJKLMN" if True})
    pl("gmn", "毛利率（non-GAAP）", "%", {c: f"={c}{R['gpn']}/{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    share = "(rd_ng+sm_ng+ga_ng)"
    for k, nm, lab_ in (("rdn", "rd_ng", "研發費用（non-GAAP）"), ("smn", "sm_ng", "銷售與行銷費用（non-GAAP）"), ("gan", "ga_ng", "一般及行政費用（non-GAAP）")):
        hk = {"rdn": "ngrd", "smn": "ngsm", "gan": "ngga"}[k]
        pl(k, lab_, "US$M", {**{c: f"=h_{hk}_{p}" for c, p in H3.items() if ex(f"h_{hk}_{p}")}, "G": f"=G{R['rev']}*{nm}", "H": f"=F{R[k]}+G{R[k]}",
                              **{c: f"={c}{R['rev']}*{nm}/{share}*MAX(opx_floor,{share}+opx_st*({c}{R['yr']}-2026))" for c in FCOLS}})
    pl("oin", "營業利益（non-GAAP）", "US$M", {**{c: f"=h_ngoi_{p}" for c, p in H3.items() if ex(f"h_ngoi_{p}")}, "G": f"=G{R['gpn']}-G{R['rdn']}-G{R['smn']}-G{R['gan']}", "H": f"=F{R['oin']}+G{R['oin']}",
                                  **{c: f"={c}{R['gpn']}-{c}{R['rdn']}-{c}{R['smn']}-{c}{R['gan']}" for c in FCOLS}}, bold=True)
    pl("omn", "營業利益率（non-GAAP）", "%", {c: f"={c}{R['oin']}/{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    pl("sbc", "股票薪酬（總額）", "US$M", {**hist("sbc", H3), "G": f"=G{R['rev']}*sbc_0", "H": f"=F{R['sbc']}+G{R['sbc']}", **{c: f"={c}{R['rev']}*MAX(sbc_floor,sbc_0+sbc_st*({c}{R['yr']}-2026))" for c in FCOLS}})
    pl("ptx", "員工股票交易薪資稅", "US$M", {"D": "=h_ptx_24", "E": "=h_ptx_25", "F": "=h_ptx_9m26", "G": f"=G{R['sbc']}*ptx_sh", "H": f"=F{R['ptx']}+G{R['ptx']}", **{c: f"={c}{R['sbc']}*ptx_sh" for c in FCOLS}})
    pl("oth", "取得無形資產攤銷＋併購費用", "US$M", {"D": "=h_amort_24+h_mna_24", "E": "=h_amort_25+h_mna_25", "F": "=h_amort_9m26+h_mna_9m26", "G": f"=G{R['rev']}*oth_adj", "H": f"=F{R['oth']}+G{R['oth']}", **{c: f"={c}{R['rev']}*oth_adj" for c in FCOLS}})
    pl("oi", "營業利益（GAAP）", "US$M", {**hist("oi", H3), **{c: f"={c}{R['oin']}-{c}{R['sbc']}-{c}{R['ptx']}-{c}{R['oth']}" for c in ["G"] + FCOLS}, "H": f"=F{R['oi']}+G{R['oi']}"}, bold=True, fill=KEY)
    pl("om", "營業利益率（GAAP）", "%", {c: f"={c}{R['oi']}/{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    pl("cogs", "營業成本（GAAP）", "US$M", {**hist("cogs", H3), "G": f"=G{R['cng']}+G{R['rev']}*gp_adj", "H": f"=F{R['cogs']}+G{R['cogs']}", **{c: f"={c}{R['cng']}+{c}{R['rev']}*gp_adj" for c in FCOLS}})
    pl("gp", "毛利（GAAP）", "US$M", {c: f"={c}{R['rev']}-{c}{R['cogs']}" for c in "CDEFGHIJKLMN"})
    pl("gm", "毛利率（GAAP）", "%", {c: f"={c}{R['gp']}/{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    pl("opex", "營業費用（GAAP；研發＋銷售行銷＋一般行政）", "US$M", {c: f"={c}{R['gp']}-{c}{R['oi']}" for c in "CDEFGHIJKLMN"})
    CFR = None   # CF 列號稍後回填（期初現金）
    pl("onet", "其他收入淨額（利息收入 − 利息費用）", "US$M", {**hist("onet", H3), "G": "=cash_0*y_cash/4-int_exp/4", "H": f"=F{R['onet']}+G{R['onet']}"})
    pl("pti", "稅前淨利", "US$M", {c: f"={c}{R['oi']}+{c}{R['onet']}" for c in "CDEFGHIJKLMN"})
    pl("taxr", "所得稅率（GAAP）", "%", {**{c: f"={c}{R['tax']}/{c}{R['pti']}" for c in "CDEFH"}, "G": "=tax_q4", **{c: f"=tax_27+(tax_lr-tax_27)*MIN(1,({c}{R['yr']}-2027)/3)" for c in FCOLS}}, "0.0%")
    pl("tax", "所得稅費用", "US$M", {**hist("tax", H3), "G": f"=G{R['pti']}*G{R['taxr']}", "H": f"=F{R['tax']}+G{R['tax']}", **{c: f"={c}{R['pti']}*{c}{R['taxr']}" for c in FCOLS}})
    pl("ni", "淨利（GAAP）", "US$M", {c: f"={c}{R['pti']}-{c}{R['tax']}" for c in "CDEFGHIJKLMN"}, bold=True)
    pl("nm", "淨利率（GAAP）", "%", {c: f"={c}{R['ni']}/{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    pl("dsh", "稀釋加權股數", "百萬股", {**{c: f"=h_dsh_{p}" for c, p in H3.items() if ex(f"h_dsh_{p}")}, "G": "=q_dsh_3q26*(1+dil/4)", "H": f"=(F{R['dsh']}*3+G{R['dsh']})/4", "I": f"=H{R['dsh']}*(1+dil)", **{c: f"={prevc(c)}{R['dsh']}*(1+dil)" for c in FCOLS[1:]}}, "#,##0.0")
    pl("eps", "EPS（GAAP，稀釋，US$）", "US$", {c: f"={c}{R['ni']}/{c}{R['dsh']}" for c in "CDEFGHIJKLMN"}, "0.00", bold=True)
    pl("ngpti", "non-GAAP 稅前淨利（營業利益＋其他收入淨額＋發行成本攤銷）", "US$M", {**{c: f"={c}{R['oin']}+{c}{R['onet']}+int_exp*{0.75 if c == 'F' else (0.25 if c == 'G' else 1)}" for c in "FGIJKLMN"}, "H": f"=F{R['ngpti']}+G{R['ngpti']}"})
    pl("ngni", "non-GAAP 淨利（公司口徑，21% 稅率）", "US$M", {c: f"={c}{R['ngpti']}*(1-tax_ng)" for c in "FGHIJKLMN"})
    pl("ngeps", "EPS（non-GAAP，稀釋，US$）", "US$", {c: f"={c}{R['ngni']}/{c}{R['dsh']}" for c in "FGHIJKLMN"}, "0.00", bold=True)
    pl("chk1", "檢查：子列合計 − 合併營收", "US$M", {c: f"={c}{R['sN']}+{c}{R['sA']}+{c}{R['sI']}-{c}{R['rev']}" for c in "HIJKLMN"}, "0.000")
    pl("chk2", "檢查：分部合計 − 合併營收", "US$M", {c: f"={c}{R['seg']}-{c}{R['rev']}" for c in "CDEFGHIJKLMN"}, "0.000")
    P_[s] = R
    r += len(ks) + 1
PB = P_["B"]
# 標準損益表（GAAP；情境 B）
r += 1
PL.cell(r, 1, "標準損益表（GAAP；US$M，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）、9M26A 為 2026 年前三季累計實績。用 GAAP：公司指引與市場多用 non-GAAP，但本模型 DCF 與損益以 GAAP 為底（股票薪酬為真實成本），non-GAAP 列見上方情境區塊與 Recon").font = BOLD
PL_NOTE = r; r += 1
for j, h in enumerate(["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "9M26A", "Q4 26E", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"], 1):
    c = PL.cell(r, j, h); c.font = BOLD; c.fill = HDR
PL_HDR = r; r += 1
IS = {}
def isrow(k, label, unit, src, fmt="#,##0.0"):
    global r
    IS[k] = r; put(PL, r, label, unit, {c: f"={c}{PB[src]}" for c in "CDEFGHIJKLMN"}, fmt); r += 1
isrow("rev", "營收", "US$M", "rev"); isrow("yoy", "營收年增率（YoY）", "%", "yoy", "0.0%"); isrow("cogs", "營業成本（COGS）", "US$M", "cogs"); isrow("gp", "毛利", "US$M", "gp"); isrow("gm", "毛利率", "%", "gm", "0.0%")
isrow("opex", "營業費用", "US$M", "opex"); isrow("oi", "營業利益", "US$M", "oi"); isrow("om", "營業利益率", "%", "om", "0.0%"); isrow("oth", "其他收入（淨額）", "US$M", "onet"); isrow("pti", "稅前利益", "US$M", "pti")
isrow("tax", "所得稅", "US$M", "tax"); isrow("ni", "稅後淨利", "US$M", "ni"); isrow("nm", "淨利率", "%", "nm", "0.0%"); isrow("eps", "EPS（完全稀釋）", "US$", "eps", "0.00")
PL.cell(r, 1, "比率區").font = BOLD; r += 1
IS_RATIO = r - 1
IS["r_yoy"] = r; put(PL, r, "營收年增率", "%", {c: f"={c}{PB['yoy']}" for c in "DEFGHIJKLMN"}, "0.0%"); r += 1
IS["r_gm"] = r; put(PL, r, "毛利率", "%", {c: f"={c}{PB['gm']}" for c in "CDEFGHIJKLMN"}, "0.0%"); r += 1
IS["r_om"] = r; put(PL, r, "營業利益率", "%", {c: f"={c}{PB['om']}" for c in "CDEFGHIJKLMN"}, "0.0%"); r += 1
IS["r_nm"] = r; put(PL, r, "淨利率", "%", {c: f"={c}{PB['nm']}" for c in "CDEFGHIJKLMN"}, "0.0%"); r += 1
IS["r_niy"] = r; put(PL, r, "淨利年增率", "%", {"D": f"=IF(C{PB['ni']}>0,D{PB['ni']}/C{PB['ni']}-1,\"N/M\")", "E": f"=IF(D{PB['ni']}>0,E{PB['ni']}/D{PB['ni']}-1,\"N/M\")",
                                                  "H": f"=IF(E{PB['ni']}>0,H{PB['ni']}/E{PB['ni']}-1,\"N/M\")", **{c: f"=IF({prevc(c)}{PB['ni']}>0,{c}{PB['ni']}/{prevc(c)}{PB['ni']}-1,\"N/M\")" for c in FCOLS}}, "0.0%",
                                    note="前期淨利 ≤ 0 時顯示 N/M"); r += 1
IS_END = r
r += 1
PL.cell(r, 1, "核對：標準損益表 營業利益 − 情境 B 區塊營業利益（應為 0；FY2023A–FY2032E）"); PL_SUM = r
for c in "CDEFGHIJKLMN": PL[f"{c}{r}"] = f"=ABS({c}{IS['oi']}-{c}{PB['oi']})"
r += 1
PL.cell(r, 1, "核對：三情境 × 全部年度 分部合計 − 合併營收 絕對值最大者（應為 0）")
for c in "CDEFGHIJKLMN": PL[f"{c}{r}"] = "=MAX(" + ",".join(f"ABS({c}{P_[s]['chk2']})" for s in "ABC") + ")"
PL_SUM2 = r; r += 1
PL.cell(r, 1, "核對：三情境 × 全部年度 子列合計 − 合併營收 絕對值最大者（應為 0）")
for c in "HIJKLMN": PL[f"{c}{r}"] = "=MAX(" + ",".join(f"ABS({c}{P_[s]['chk1']})" for s in "ABC") + ")"
PL_SUM3 = r; r += 1
widths(PL)

# ====================================================================================================== CF
CF["A1"] = "CF — 現金流（US$M）：自由現金流（公司定義：營運現金流 − 資本支出 − 資本化軟體）；不含股利與回購（Datadog 無）；2029 年到期可轉換票據以現金償還；收支關卡（FY2026–FY2030 自由現金流合計）見下方。"
CF["A2"] = "營運現金流＝淨利＋折舊攤銷＋股票薪酬 − 營運資金增加（營運資金增加＝wc_k × 營收增量；歷史由 Hist 頁勾稽）。併購現金支出不建模（9M26 為 115.6，列風險登錄）。"
C_ = {}
r = 4
for s in "ABC":
    r = blk_header(CF, r, f"情境 {s}（{ {'A': '樂觀', 'B': '基準', 'C': '保守'}[s] }）")
    R = {}
    ks = ["yr", "da", "sbc", "dwc", "ocf", "cx", "fcf", "fm", "repay", "cash", "debt", "nc", "chk"]
    for i, k in enumerate(ks): R[k] = r + i
    PR = P_[s]
    put(CF, R["yr"], "年度（數值）", "年", {c: YR_OF[c] for c in "HIJKLMN"}, "0")
    put(CF, R["da"], "折舊與攤銷", "US$M", {"C": "=h_da_23", "D": "=h_da_24", "E": "=h_da_25", "F": "=h_da_9m26", "G": f"=PL!G{PR['rev']}*da_r", "H": f"=F{R['da']}+G{R['da']}", **{c: f"=PL!{c}{PR['rev']}*da_r" for c in FCOLS}})
    put(CF, R["sbc"], "股票薪酬（加回）", "US$M", {c: f"=PL!{c}{PR['sbc']}" for c in "CDEFGHIJKLMN"})
    put(CF, R["dwc"], "營運資金增加（負＝流入）", "US$M", {"D": "=-h_wcx_24", "E": "=-h_wcx_25", "F": "=-h_wcx_9m26", "G": f"=wc_k*(PL!G{PR['rev']}-q_rev_4q25)", "H": f"=F{R['dwc']}+G{R['dwc']}", **{c: f"=wc_k*(PL!{c}{PR['rev']}-PL!{prevc(c)}{PR['rev']})" for c in FCOLS}},
        note="實績＝−(營運現金流 − 淨利 − 折舊攤銷 − 股票薪酬)；預測＝wc_k × 營收增量")
    put(CF, R["ocf"], "營運現金流", "US$M", {"C": "=h_ocf_23", "D": "=h_ocf_24", "E": "=h_ocf_25", "F": "=h_ocf_9m26", **{c: f"=PL!{c}{PR['ni']}+{c}{R['da']}+{c}{R['sbc']}-{c}{R['dwc']}" for c in ["G"] + FCOLS}, "H": f"=F{R['ocf']}+G{R['ocf']}"}, bold=True)
    put(CF, R["cx"], "資本支出（含資本化軟體）", "US$M", {"C": "=h_capex_23+h_capsw_23", "D": "=h_capex_24+h_capsw_24", "E": "=h_capex_25+h_capsw_25", "F": "=h_capex_9m26+h_capsw_9m26", "G": f"=PL!G{PR['rev']}*cx_r", "H": f"=F{R['cx']}+G{R['cx']}", **{c: f"=PL!{c}{PR['rev']}*cx_r" for c in FCOLS}})
    put(CF, R["fcf"], "自由現金流（公司定義）", "US$M", {c: f"={c}{R['ocf']}-{c}{R['cx']}" for c in "CDEFGHIJKLMN"}, bold=True, fill=KEY)
    put(CF, R["fm"], "自由現金流率", "%", {c: f"={c}{R['fcf']}/PL!{c}{PR['rev']}" for c in "CDEFGHIJKLMN"}, "0.0%")
    put(CF, R["repay"], "償還可轉換票據（2029-12-01 到期；現金償還；轉換價 217.60 高於現價）", "US$M", {"K": "=-debt_0"}, "#,##0.0")
    put(CF, R["cash"], "期末現金＋有價證券", "US$M", {"G": f"=cash_0+G{R['fcf']}", "H": f"=G{R['cash']}", **{c: f"={prevc(c)}{R['cash']}+{c}{R['fcf']}+{c}{R['repay']}" for c in FCOLS}}, note="Q4 26E 起算自 2026-09-30 餘額；不含併購、ESPP 收入與回購")
    put(CF, R["debt"], "期末債務本金", "US$M", {"G": "=debt_0", "H": f"=G{R['debt']}", **{c: f"={prevc(c)}{R['debt']}+{c}{R['repay']}" for c in FCOLS}})
    put(CF, R["nc"], "淨現金（不含營業租賃）", "US$M", {c: f"={c}{R['cash']}-{c}{R['debt']}-fin_lease" for c in "GHIJKLMN"})
    put(CF, R["chk"], "檢查：期末現金 − (期初 ＋ 自由現金流 ＋ 償債)（應為 0）", "US$M", {c: f"={c}{R['cash']}-({prevc(c) if c != 'I' else 'H'}{R['cash']}+{c}{R['fcf']}+{c}{R['repay']})" for c in FCOLS}, "0.000")
    C_[s] = R
    r += len(ks) + 1
GATE = r
CF.cell(r, 1, "收支關卡（FY2026–FY2030 自由現金流扣股利合計；US$M）").font = BOLD
for j, s in enumerate("ABC", 2):
    CF.cell(r, j, f"=SUM(H{C_[s]['fcf']}:L{C_[s]['fcf']})").number_format = "#,##0.0"
CF.cell(r, 5, "情境 A／B／C；Datadog 無股利與回購，故等於自由現金流合計；判準＝基準情境（B）合計 ≥ 0 為通過（原則第 10 節第 63 條）"); r += 1
CF.cell(r, 1, "期間最大單年新增舉債（FY2026–FY2030）");
for j, s in enumerate("ABC", 2): CF.cell(r, j, 0)
CF.cell(r, 5, "無新增舉債（現金充足；票據到期以現金償還）"); r += 1
CF.cell(r, 1, "期間最低年末現金（FY2026–FY2030；US$M）")
for j, s in enumerate("ABC", 2): CF.cell(r, j, f"=MIN(H{C_[s]['cash']}:L{C_[s]['cash']})").number_format = "#,##0.0"
GATE_MINC = r; r += 1
CF.cell(r, 1, "融資可行性檢查：淨債務 ÷ 調整後 EBITDA（負＝淨現金）；利息保障倍數")
CF.cell(r, 2, "不適用：淨現金，無新債缺口（原則第 10 節第 9 條僅適用以新債補缺口者）"); r += 1
widths(CF)
exec(open(os.path.join(HERE, "build_datadog_p3.py"), encoding="utf-8").read())
