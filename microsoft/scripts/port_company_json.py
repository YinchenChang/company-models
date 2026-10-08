# MAG v0.1b′（Microsoft）：以 Amazon r2 的 company.json 為骨架，填入 Microsoft v0.1a 資料（mag 區段）與移植預設（已套用預設見進度檔與 v0.1b′ 報告）。
# 用法：python3 scripts/port_company_json.py <amazon company.json（git show origin/claude/amazon-v0.1:amazon/company.json）> <含 mag 區段的 company.json> <data/tokenomics_snapshot_v5.27.json> <輸出 company.json>；之後執行 node scripts/calib_pace.js --write 與 python3 scripts/fields_doc.py --write
import json, sys, copy

A = json.load(open(sys.argv[1], encoding='utf-8'))
MS = json.load(open(sys.argv[2], encoding='utf-8'))
G = MS['mag']
C = copy.deepcopy(A)
C.pop('mag', None)
C['mag'] = G
# mag_sens.js 讀 mag.mw.externalIT（期初對外計費 MW 區間）；v0.1a split.extMW 的同一組數字
G['mw']['externalIT'] = {"value": {"low": 2150, "base": 3250, "high": 4875}, "unit": "MW-IT", "ref": "split.extMW",
                         "mapTo": "defaults.billableOpen（自有 × 對外占比＋neocloud 租用 × 50%；敏感度區間）"}

r3 = lambda x: round(x, 3)
r6 = lambda x: round(x, 6)
Q = 'FY26Q4'
RENT_FULL, RENT_MW = 3.438 + 3.48 + 1.94 + 4.6 + 0.5, 700 + 205 + 200 + 411 + 55  # neocloud 全額年租金、MW
OPEX_MW = 2.34  # 期初世代加權 IF_OpexGW（Hopper 20%／GB200 45%／GB300 35%；US$m/MW）
RENT_GROSS = RENT_FULL / RENT_MW * 1000  # US$m/MW·年
RCF = 0.5 + 0.5 * (RENT_GROSS - OPEX_MW) / RENT_GROSS

# ---------- meta / calendar ----------
C['meta'] = {"company": "Microsoft", "ticker": "MSFT", "updateDate": "2026-10-09", "priceDate": "2026-10-07",
  "sourceOrderNote": "時序先 FY26 10-K（2026-07-29，經查核）與 FY26Q4 財報新聞稿／法說（2026-07-29），再 2026-09-02 8-K（FY27 分部重編與 Azure 絕對金額，[Interested-party]）。入帳數字聽 10-K；指引聽法說與 8-K。AI 在役 MW 公司未揭露（只揭露新增 GW），期初對外 AI MW 取第三方與公司增量三法交叉估計，不以 Azure 營收、RPO 或 AI run-rate 反推。FY27Q1 財報預計 2026-10-28（本版未納入，對照表 C7）。",
  "consensusFile": "data/consensus_msft_20261008.json"}
C['calendar'] = {"fiscalYearEndMonth": 6, "latestQuarterFiled": Q, "latestQuarterReported": Q, "firstModelFY": None,
                 "modelYears": 4, "targetHorizon": "valuationPlus12m", "periodLabel": "FY"}
C['asOf'] = {k: (Q if not k.startswith('_') else v) for k, v in A['asOf'].items()}

# ---------- ytdActual（FY27 年初至今＝0：FY26 10-K 為最新已申報）----------
ya = {k: 0.0 for k in ['revenue', 'capex', 'cashCapex', 'interest', 'leasePaid', 'debtRepaid', 'borrow', 'equity', 'cappedCall', 'jv', 'cfo', 'prepay', 'da', 'sbc', 'opInc', 'ni', 'eps', 'ngEps', 'adjEbitda', 'dividends', 'buyback']}
C['ytdActual'] = {"throughQuarter": Q, "months": 0, "label": "FY27 年初至今（無：最新已申報為 FY26 10-K）", "cash1231": 76.843, **ya,
  "adjEbitdaMeta": {"tag": "Derived", "sources": [], "crossCheck": "不適用（首期＝FY27 全年模型，年初至今無實際數）", "note": "FY27 年初至今 0", "usage": "只用於「與市場的差異」"},
  "jvSplit": {"jv": 0.0, "strategic": 0.0},
  "notes": {"cash1231": "上一財年底（2026-06-30）現金及約當現金 20,935＋短期投資 55,908＝76,843（10-K 經查核；模型現金口徑含短期投資）[Verified]",
            "cfo": "不適用（FY27 年初至今無實際數）", "cashCapex": "不適用", "capex": "不適用", "jv": "不適用", "borrow": "不適用", "debtRepaid": "不適用",
            "cappedCall": "不適用", "equity": "不適用", "interest": "不適用", "leasePaid": "不適用", "revenue": "不適用", "opInc": "不適用", "ni": "不適用",
            "prepay": "不適用（無客戶預付融資機制）", "da": "不適用", "sbc": "不適用", "eps": "不適用", "ngEps": "不適用", "dividends": "不適用", "buyback": "不適用"}}

# ---------- historicalPL ----------
pl = G['pl']['value'] if 'value' in G['pl'] else None
def hp(y, rev, oi, ni, eps, sh, sbc):
    return {"year": y, "revenue": rev, "opInc": oi, "ni": ni, "eps": eps, "ngEps": round((ni + sbc) / sh, 2), "shares": sh, "tag": "Verified"}
C['historicalPL'] = [hp("FY24", 245.122, 109.433, 88.136, 11.80, 7.469, 10.734), hp("FY25", 281.724, 128.528, 101.832, 13.64, 7.465, 11.974),
                     hp("FY26", 331.839, 155.237, 133.749, 17.95, 7.453, 12.405),
                     {"year": "FY27 年初至今", "revenue": 0.0, "opInc": 0.0, "ni": 0.0, "eps": 0.0, "ngEps": 0.0, "shares": 7.427, "tag": "Derived"}]  # 引擎以第 4 格為年初至今（FY26Q4 已申報 → FY27 年初至今 0）

# ---------- RPO（只作對照；rpoOpen＝0）----------
C['rpo'] = {"scheduledShare": 0.9, "bucketWeights": [0.3, 0.25, 0.2, 0.1, 0.05], "bucketLabels": ["≤12 個月約 30%（其餘未分桶；加權平均 2.3 年）"], "split": [1.0], "within36m": 0.75}

# ---------- 租賃 ----------
opP = [6.082, 4.334, 3.146, 2.612, 2.316, 6.216]
fnP = [7.121, 7.294, 6.668, 6.570, 6.543, 55.490]
C['leases'] = {
  "onBalanceCash": [r3(a + b) for a, b in zip(opP[:5], fnP[:5])], "afterFY30": r3(opP[5] + fnP[5]),
  "facts": {"onBal": r3(sum(opP) + sum(fnP)), "notCommenced": 329.1, "singleCap": 0.0, "share": 1.0},
  "operatingPayments": opP, "financePayments": fnP,
  "liability": {"discRate": 0.043, "tailYears": 10, "note": "租賃負債（每期末）＝剩餘租金現值。折現率＝10-K 附註 13 加權平均折現率（融資 4.5%、營業 3.7%，按負債 66.6／21.9 加權 4.3%）[Derived]；在帳租約＝附註 13 到期表各期現金（營業＋融資）＋模型期後尾端（61.7）平均分 10 年（融資租賃加權剩餘 13 年、營業 6 年，[Assumed]）；未起租租約只計已起租部分。"},
  "uncommenced": {"startQ": 0, "quarters": 16, "termYears": 15, "delayLink": 0.0, "delayLinkNote": "不適用（不建模建設延誤；三情境延誤月數皆 0）", "termSens": [12, 18],
    "note": "10-K 附註 13：尚未起租租賃 329.1B（主要為資料中心；FY27–FY33 起租、租期 1–20 年；逐年時程與營業／融資拆分未揭露）。模型：平均分 16 季（FY27–FY30）起租、每筆 15 年直線付租 [Assumed]；敏感度 12／18 年。一年內序列 106→155→197→329（每季新增 40–130）。",
    "opShare": 0.75, "cashShareLabel": None,
    "opShareNote": "三家統一定義（MAG 對照表 r1 C21）：未起租租約中預期列為營業租賃的比例（租金入 EBITDA 的部分；不另自現金扣除；1 − opShare＝融資部分，自現金扣除）。Microsoft：75%（公司稱 FY27 起更多資料中心租約改列營業租賃；在帳負債營業占 25% 是改分類前的結構；[Assumed]，區間 25%–100%）。注意：本家資本支出採含融資租賃口徑，融資部分起租時也進資本支出（與自現金扣除的融資租金重複，偏保守）；營業部分的機房租金依 C22 計入對外 AI 營運成本（只進 AI 增量報酬）"},
  "operatingInEbitda": True,
  "operatingInEbitdaNote": "分部營業利益已扣在帳營業租賃成本（FY26 6.968）→ 在帳營業租賃現金不再自現金扣除；在帳融資租賃本息仍為現金支出（過去以融資租賃取得、尚未付清的資產）；未起租部分見 uncommenced.opShareNote（對照表 r1 C8 f）",
  "rentedCompute": [
    {"name": "CoreWeave", "start": "2025-07", "years": 10, "annualRent": round(3.438 * RCF, 3), "mw": 700, "use": "自用 50%／對外 50%"},
    {"name": "Nebius", "start": "2025-11", "years": 10, "annualRent": round(3.48 * RCF, 3), "mw": 205, "use": "自用 50%／對外 50%"},
    {"name": "IREN", "start": "2026-06", "years": 10, "annualRent": round(1.94 * RCF, 3), "mw": 200, "use": "自用 50%／對外 50%"},
    {"name": "Nscale", "start": "2026-10", "years": 10, "annualRent": round(4.6 * RCF, 3), "mw": 411, "use": "自用 50%／對外 50%"},
    {"name": "Lambda", "start": "2026-01", "years": 10, "annualRent": round(0.5 * RCF, 3), "mw": 55, "use": "自用 50%／對外 50%"}],
  "rentedComputeNote": "租用算力（向 neocloud 租 GPU；對照表 r1 第 5 節第 7 條、C8 d、C20）：租金計入營運成本（扣 EBITDA 與營運現金，不是資本支出）；用途無揭露 → 對外／自用各半。年租金（全額）：CoreWeave 3.438（2025 實際推估；起訖未揭露，假設 FY26 起 5 年）、Nebius 3.48（17.4 ÷ 5 年）、IREN 1.94（9.7 ÷ 5 年）、Nscale 4.6（金額未揭露，二手 23 ÷ 5 年，[Interested-party 二手／Assumed]）、Lambda 0.5（multibillion，[Assumed]）。年期：合約 5 年＋假設續約 5 年（years 10，與 rentedExt 的續約假設一致；查核 a）。本表每筆＝自用 50% 全額＋對外 50% 扣 Tokenomics 營運成本後的淨額（對外 MW 的營運成本已在 AI 雲端 EBITDA 內）＝全額 × " + f"{RCF:.3f}" + "（[Derived]）；對外 50% 的 MW 列 capexModel.rentedExt（收入照算、不計資本支出與折舊）。分部 EBITDA 率的起點已加回 FY26 neocloud 租金約 6.3（[Assumed]），避免與本表重複扣。"}

# ---------- 債務 ----------
notes = [  # (名稱, 面額, 到期年區間, 票面區間)
  ("2009 票據", 0.52, (2039, 2039), (5.20, 5.20)), ("2010 票據", 0.486, (2040, 2040), (4.50, 4.50)), ("2011 票據", 0.718, (2041, 2041), (5.30, 5.30)),
  ("2012 票據", 0.454, (2042, 2042), (3.50, 3.50)), ("2013 美元票據", 0.314, (2043, 2043), (3.75, 4.88)), ("2013 歐元票據", 2.63, (2028, 2033), (2.63, 3.13)),
  ("2015 票據", 4.555, (2035, 2055), (3.50, 4.75)), ("2016 票據", 7.93, (2026, 2056), (2.40, 3.95)), ("2017 票據", 6.833, (2026, 2057), (3.30, 4.50)),
  ("2020 票據（交換）", 10.111, (2030, 2060), (1.35, 2.68)), ("2021 票據（交換）", 8.185, (2052, 2062), (2.92, 3.04)),
  ("2023 票據（Activision 交換）", 0.056, (2026, 2050), (1.35, 4.50)), ("2024 票據（Activision 交換）", 3.344, (2026, 2050), (1.35, 4.50))]
ins = []
for n, f, (y0, y1), (c0, c1) in notes:
    ym = (y0 + y1) // 2
    ins.append([n, "無擔保", f"{ym}-06", round((c0 + c1) / 200, 5), f,
                f"10-K FY26 附註 10：面額 {f * 1000:,.0f}、票面 {c0:.2f}%–{c1:.2f}%（取中點）、到期 {y0}–{y1}；到期日取區間中點 [Derived]"])
assert abs(sum(x[4] for x in ins) - 46.136) < 1e-6
C['debt'] = {"amortization": [9.25, 0.0, 2.001, 0.0, 0.5], "amortAfterFY30": 34.385, "instruments": ins,
  "convertible": {"principal": 0.0, "coupon": 0.0}, "convertibles": [],
  "convertibleBridge": {"exchangedAccreted": 0.0, "newIssuesAccreted": 0.0, "note": "無可轉債；債務明細＝13 次票據發行＝面額 46.136（10-K 附註 10；帳面 40.294 含交換折價）；無商業本票。評價日後未見新發債。"},
  "amortizationNote": "10-K 附註 10 到期表（本金）：FY27 9.250、FY28 0、FY29 2.001、FY30 0、FY31 0.500、其後 34.385 [Verified]"}

# ---------- 最新季 ----------
LQ = copy.deepcopy(A['latestQuarter'])
for k in LQ:
    if isinstance(LQ[k], (int, float)) and not isinstance(LQ[k], bool): LQ[k] = None
LQ.update({"filed": "2026-07-29", "periodEnd": "2026-06-30", "revenue": 90.007, "yoy": r6(90.007 / 76.441 - 1), "h1Revenue": 331.839,
  "opInc": 40.603, "interest": None, "h1Interest": 0.0, "ni": None, "h1Ni": 0.0, "epsDiluted": None, "sbc": 3.122, "da": 11.022,
  "adjEbitda": r3(40.603 + 11.022), "adjOpInc": 40.603, "rpo": 684.0, "cash": 20.935, "restricted": 0.0, "marketable": 55.908, "availability": 0.0,
  "debtPrincipal": 46.136, "ddtlOut": 0.0, "notesOut": 40.294, "sharesOut": 7.427, "basicWaso": None,
  "capexQ2": 0.0, "capexH1": 0.0, "cashCapexH1": 0.0, "cfoH1": 0.0, "cashInterestH1": 0.0,
  "deferredTotal": r3(72.965 + 2.747), "deferredIn": 0.0, "ppe": 313.076, "opLeaseLiab": 21.925, "finLeaseLiab": 66.594,
  "onBalanceUndiscounted": C['leases']['facts']['onBal'], "offBalanceLease": 329.1, "singleSiteCap": 0.0, "jvCommit": 1.1, "jvPaidH1": 0.0,
  "leaseCashH1": 0.0, "custA": 0.0, "custB": 0.0, "custC": 0.0, "debtIssuedH1": 0.0, "debtRepaidH1": 0.0, "equityH1": 0.0})
C['latestQuarter'] = LQ

# ---------- 法說事實 ----------
CAPEX27 = round(2 * (175.0 - (31.9 + 41.0)), 1)  # 204.2：CY2026 約 175 − FY26 下半年實際（Q3 31.9＋Q4 41.0）＝FY27 上半年 102.1；下半年＝上半年
CF = {k: None for k in A['callFacts']}
CF.update({"capexLo": CAPEX27, "capexHi": CAPEX27, "nextQCapexLo": 50.0, "nextQCapexHi": None, "nextQRevLo": 89.85, "nextQRevHi": 90.95,
  "nextQAdjOpLo": round(89.85 - 29.8 - 16.9, 2), "nextQAdjOpHi": round(90.95 - 29.6 - 16.8, 2), "convert": 0.0, "convertNet": 0.0,
  "availability": 0.0, "priceLast": 529.76, "priceDate": "2026-10-07",
  "postQShortDated": "評價日後：2026-09-02 8-K 分部重編（Agents and Infra／Devices and Consumer）與 Azure 絕對金額；2026-09-15 宣告季度股利 0.98（+8%）；FY27Q1 財報 2026-10-28（未納入）",
  "ttmRev": 331.839, "ttmOpInc": 155.237, "ttmDa": 38.534, "ttmOpLease": 21.925, "mktCapLast": round(529.76 * 7.427, 1),
  "nextEarn": "2026-10-28（FY27Q1 財報）", "aiRunRate": 37.0,
  "aiRunRateNote": "FY26Q3 法說：AI 業務年化營收「surpassed $37 billion」（含 M365／GitHub Copilot 等自有 AI 產品，非僅 Azure 對外；事實總帳 call.aiRunRate.fy26q3）[Interested-party]；只作驗證列，不回頭改 MW 或 k（對照表 r1 C3）。扣 Copilot（約 30M 席 × $30／月 ≈ 10.8，[Assumed]）後隱含對外 AI 約 26"})
C['callFacts'] = CF

# ---------- 情境（容量軸）----------
OPEN = 3250
LOW_CAP = round(OPEN + (2263 - 369) * 0.6)
C['scenarios'] = {
  "labels": {"low": "保守 第三方可辨識站點", "base": "基準 依資本支出指引", "high": "積極 依公司容量目標"},
  "descriptions": {
    "low": f"對外 AI MW-IT 只計 Epoch AI 可辨識站點的已交付＋在建：Fairwater Wisconsin 369 → 2028Q2 預測 2,263（+1,894，約 1.75 年）× 對外 60% → 每年 +650、上限 {LOW_CAP:,}（Atlanta 636 已全數運轉）[Derived]",
    "base": "對外 AI MW 新增速度＝近端資本支出隱含的建置量：FY27 資本支出（公司口徑含融資租賃）204.2 −（非 AI 公式值＋GPU 汰換）÷ 每 MW 全成本（Tokenomics IT＋自建比例 × 機房）× 對外 60%，依 λ 遞延上線（scripts/calib_pace.js 解，對帳落差＝0）；之後各年資本支出＝MW 路徑 × 每 MW 成本（同一套數字）[Derived]。FY27 資本支出＝2 ×（曆年 2026 約 175 − FY26 下半年實際 72.9）＝204.2（下半年＝上半年；FY27Q1 指引 >50 吻合）",
    "high": "公司目標：總產能約兩年翻倍（FY26 起多次重申）。以 Bloomberg 總容量約 12 GW（2026-09，口徑不明）為基數 → 每年 +6 GW（設施）× AI 占新增 80%（70%–90% 中點）÷ PUE 1.2 × 對外 60%＝每年 +2,400 MW-IT [Derived；總容量為第三方報導，非公司揭露]"},
  "mwPath": {"connectedStart": None,
    "contracted": {"low": [LOW_CAP] * 5, "base": [99999] * 5, "high": [99999] * 5},
    "calibrate": {"scenario": "base", "decimals": 1},
    "pace": {"low": 650, "base": 2200, "high": 2400},
    "note": "對外 AI MW-IT（在役）：評價日 3,250（v0.1a：自有在役 AI 4,500 × 對外 60%＋neocloud 租用 1,100 × 50%；三法交叉估計，事實總帳 split.extMW）[Derived／Assumed]；各期期末＝MIN(上限, 前期＋速度 × 期間長度)，首期自評價日起算（connectedStart＝null）。MAG r1 C10：基準速度由 calibrate 設定以資本支出指引解出（scripts/calib_pace.js；verify 檢查）；保守＝第三方可辨識站點速度、積極＝公司容量目標換算。"},
  "billableRatio": {"mode": "ratio", "ratio": [1, 1, 1, 1, 1], "note": "對外 AI MW 依 MW-year 計價：已連網即計費（在役比例 100%）；期初計費 MW＝三法交叉估計 3,250，不以 Azure 營收校準 [Assumed]", "openAnnualRevenue": None},
  "mw31": {"low": 0, "base": 2200, "high": 2400},
  "delayMonths": {"low": 0, "base": 0, "high": 0, "note": "不適用（Microsoft 不建模建設延誤；機制保留，輸入 0）"},
  "convCap": {"low": 0, "base": 0, "high": 0},
  "capexTemplate": {"costMW": [44.1] * 5, "div": [0.0] * 5}}

# ---------- 非 AI 事業各線 ----------
# FY26／FY25 產品別營收（8-K 2026-09-02 重述；mag.lines 季度加總）
qs = lambda k, fy: round(sum(v for q, v in G['lines'][k]['revQ']['value'].items() if q.startswith(fy)), 3)
REV = {k: (qs(k, 'FY26'), qs(k, 'FY25')) for k in G['lines']}
DA, REV_T = 38.534, 331.839
DAR = DA / REV_T  # D&A 按營收比例分攤（對照表第 5 節第 5 條；分部 D&A 不揭露）
RENT26 = 6.3      # FY26 neocloud 租金估計（加回 Azure EBITDA 起點，避免與 rentedCompute 重複扣）
OI_A, OI_D = 136.365, 18.872
om = {"m365cloud": 0.60, "licensing": 0.75, "indfront": 0.30, "searchads": 0.30}
rev = {"m365cloud": REV['m365cloud'], "licensing": REV['licensing'],
       "indfront": (r3(REV['industry'][0] + REV['frontier'][0]), r3(REV['industry'][1] + REV['frontier'][1])),
       "searchads": REV['searchads'], "devgame": (r3(REV['windows'][0] + REV['xbox'][0]), r3(REV['windows'][1] + REV['xbox'][1])), "azure": REV['azure']}
oi_az = OI_A - sum(rev[k][0] * om[k] for k in ("m365cloud", "licensing", "indfront"))
oi_dg = OI_D - rev['searchads'][0] * om['searchads']
om['devgame'] = oi_dg / rev['devgame'][0]
m0 = {k: r6(om[k] + DAR) for k in om}
az_eb = r6(oi_az + rev['azure'][0] * DAR + RENT26)
g0 = {k: r6(rev[k][0] / rev[k][1] - 1) for k in rev}
# 期初 AI 計費 MW（對外）：FY24 末 640、FY25 末 1,630、FY26 末 3,250 → FY26／FY25 平均
aiTTM, aiPrev = (1630 + 3250) / 2, (640 + 1630) / 2
CX = {"m365cloud": 0.10, "licensing": 0.02, "indfront": 0.05, "searchads": 0.08, "devgame": 0.03, "azure": 0.20}
lines = [
  {"key": "m365", "label": "M365 雲端（商用＋消費）", "kind": "growth", "peer": "software", "fyBase": rev['m365cloud'][0], "ytd": 0.0, "g0": g0['m365cloud'], "gLT": 0.10, "m0": m0['m365cloud'], "mLT": None, "oa": 0.0, "cx": CX['m365cloud'], "m0Note": "營業利益率分配 60%（舊 PBP 分部 59.9% 類比）＋D&A 按營收 11.6% [Assumed]"},
  {"key": "licensing", "label": "生產力與伺服器授權", "kind": "growth", "peer": "software", "fyBase": rev['licensing'][0], "ytd": 0.0, "g0": g0['licensing'], "gLT": 0.02, "m0": m0['licensing'], "mLT": None, "oa": 0.0, "cx": CX['licensing'], "m0Note": "營業利益率分配 75%＋D&A 11.6% [Assumed]（產品線利潤不揭露）"},
  {"key": "indFrontier", "label": "產業解決方案＋Frontier 與支援服務", "kind": "growth", "peer": "software", "fyBase": rev['indfront'][0], "ytd": 0.0, "g0": g0['indfront'], "gLT": 0.07, "m0": m0['indfront'], "mLT": None, "oa": 0.0, "cx": CX['indfront'], "m0Note": "營業利益率分配 30%＋D&A 11.6% [Assumed]"},
  {"key": "searchAds", "label": "搜尋與廣告（含 LinkedIn 行銷）", "kind": "growth", "peer": "ads", "fyBase": rev['searchads'][0], "ytd": 0.0, "g0": g0['searchads'], "gLT": 0.07, "m0": m0['searchads'], "mLT": None, "oa": 0.0, "cx": CX['searchads'], "m0Note": "營業利益率分配 30%＋D&A 11.6% [Assumed]"},
  {"key": "devGaming", "label": "Windows 與裝置＋XBOX", "kind": "growth", "peer": "devGaming", "fyBase": rev['devgame'][0], "ytd": 0.0, "g0": g0['devgame'], "gLT": 0.02, "m0": m0['devgame'], "mLT": None, "oa": 0.0, "cx": CX['devgame'], "m0Note": "Devices and Consumer 分部扣搜尋與廣告後的殘差 29.4%＋D&A 11.6% [Derived]"},
  {"key": "azureNonAi", "label": "Azure 非 AI 雲端", "kind": "cloudResidual", "peer": "cloud", "ytd": 0.0, "priorStub": rev['azure'][0], "g4q": g0['azure'], "ttm": rev['azure'][0], "prevTTM": rev['azure'][1],
   "ebitdaTTM": az_eb, "aiMwTTM": aiTTM, "aiMwPrevTTM": aiPrev, "gLT": 0.08, "mLT": None, "oa": 0.0, "cx": CX['azure']}]
DFL = C['defaults']
DFL['legacyBiz'] = {"lines": lines,
  "note": "各線：growth＝全年營收＝上一財年（FY26，8-K 2026-09-02 新口徑重述）×(1＋年增率)，年增率自 FY26 對 FY25 年增率（g0，[Derived]）線性收斂到長期值（gLT，[Assumed]，區間 ±3 個百分點）；首期＝FY27 全年模型（年初至今 0）。cloudResidual（Azure 非 AI）＝Azure 首期估計〔FY26 101.938 ×(1＋近四季年增率 40.4%)〕− 對外 AI 雲端模型值；之後年增率由推估的近四季非 AI 年增率〔(Azure FY26 − AI FY26 估計) ÷ (FY25 − AI FY25 估計) − 1〕收斂到 gLT（8%，5%–12%，[Assumed]）；EBITDA 率＝(Azure EBITDA 估計 − AI 雲端 EBITDA 估計) ÷ 非 AI 營收（[Derived]，固定）。EBITDA 率（m0，五期固定，mLT 空白＝m0）＝分配的營業利益率＋D&A 按營收分攤 11.61%（38.534 ÷ 331.839；分部 D&A 公司不揭露，對照表第 5 節第 5 條，[Derived]）。營業利益分配：Agents and Infra 分部 136.365 中 M365 雲端 60%、授權 75%、產業＋Frontier 30%（[Assumed]），Azure＝殘差；Devices and Consumer 分部 18.872 中搜尋與廣告 30%（[Assumed]），Windows＋XBOX＝殘差。Azure EBITDA 起點另加回 FY26 neocloud 租金 6.3（[Assumed]；租金改由 leases.rentedCompute 列支）。oa＝0（無影音內容類其他攤銷；無形資產攤銷併入 D&A 池校準）。peer＝評價時套用的同業倍數組（valuation.segmentMultiples）。",
  "split": {"opMargins": om, "note": "營業利益率分配（[Assumed]；Agents and Infra 與 Devices and Consumer 分部合計＝10-K／8-K 實際）：M365 雲端 60%（舊 PBP 分部 FY26 59.9%）、授權 75%（地端授權毛利高、銷售費用低）、產業＋Frontier 30%（含支援服務）、搜尋與廣告 30%；Azure 與 Windows＋XBOX 為殘差（" + f"Azure {oi_az / rev['azure'][0]:.1%}、Windows＋XBOX {om['devgame']:.1%}" + "）；區間 ±10pt", "daShare": r6(DAR), "rentAddBack": RENT26},
  "cxNote": "cx＝非 AI 資本支出占全年營收比（[Assumed]；公司不揭露分部資本支出）：Azure 非 AI 20%（同 Amazon AWS 非 AI 口徑）、M365 雲端 10%（雲端服務自有基礎設施）、搜尋與廣告 8%（Bing 基礎設施）、產業＋Frontier 5%、Windows＋XBOX 3%、授權 2%"}
DFL.update({"mwYearEnd": {"2024": 640, "2025": 1630, "2026": OPEN}, "mw31": 2200, "capexFloorFY0": CAPEX27, "billableOpen": OPEN,
  "ebStart": 0.79, "ebSteady": 0.79, "cashTaxRate": 0.20, "ppeOpen": 313.076, "jvCommit": [1.1, 0.0, 0.0, 0.0, 0.0], "cash": 76.843,
  "eqPx": 529.76, "eqCapShares": 7.427, "minCash": 50, "useFacility": False, "facility": 0.0,
  "otherEbitdaNote": "未分攤公司層費用：Microsoft 分部營業利益合計＝合併營業利益（無未分攤項）→ 0；租用算力租金另由 leases.rentedCompute 自本列扣除",
  "cdsDate": "不適用（MSFT CDS 無可引用的一手報價）",
  "prepay": {**A['defaults']['prepay'], "note": "不適用（Microsoft 無客戶預付融資；機制保留，輸入 0）"},
  "gpuLifeNote": "Tokenomics IF_DeprLifeIT 6 年（tk.IF_DeprLifeIT）；10-K 伺服器與網路設備 2–6 年（事實總帳 pl.usefulLife.servers）[Verified]",
  "dividend": {"perShareQ": 0.98, "growth": 0.08, "growthNote": "每股股利年增 8%（2026-09-15 宣告 0.98，較前季 +8%；FY25→FY26 +9.6%）[Interested-party]", "sharesBase": 7.427, "preferred": [0.0] * 5, "note": "每季 0.98（2026-09-15 宣告，+8%；事實總帳 sh.dpsQ.fy27q1）[Interested-party]；第 n 期 ×(1＋8%)^n（defaults.dividend.growth）"},
  "buyback": {"annual": 22.271, "floorShare": 0.0, "note": "回購基準＝近四季實際年額 22.271（FY26 現金流量表普通股買回，含員工扣稅買回 5.6；計畫內 16.719；授權餘額 40.6；事實總帳 cf.buyback.fy2026）[Verified]；floorShare 0＝可全數取消（對照表 r1 第 5 節第 8 條）"},
  "sbcRate": r6(12.405 / 331.839), "sbcNote": "股權報酬（SBC）為非現金費用：分部營業利益已扣除，營運現金加回＝SBC 占營收 × 模型期總營收；FY26 SBC 12.405 ÷ 營收 331.839＝3.74%（現金流量表；[Derived]）。股數稀釋沿用模板每年 1%",
  "wcStub": -4.895, "wcStubNote": "首期（FY27 全年）營運資金變動＝上一年度同期實際：FY26 營運資產與負債變動合計 −4.895（應收 −12.737、存貨 −0.461、其他流動資產 −2.627、其他長期資產 −3.964、應付 +5.268、遞延收入 +9.361、所得稅 −1.875、其他流動負債 +6.847、其他長期負債 −4.707；10-K 現金流量表，SEC XBRL）[Derived]；之後年度沿用營收增量比例（MAG 對照表 r1 C14；首期為全年，季節性不適用，照字面套用上一年度全年）",
  "sites": [
    {"id": "fairwaterWI", "name": "Fairwater Wisconsin（Mount Pleasant）", "operator": "Microsoft 自建（OpenAI 與 Microsoft，B200）", "planned": 2263, "energized": 369, "accepted": 369, "billable": 369, "status": "Epoch AI：目前 IT 電力 369 MW（2026-10-05 更新；2028Q2 預測 2,263；事實總帳 mw.epoch.fairwaterWI）[Analogy／第三方]", "next": "2028Q2 預測 2,263 MW-IT", "date": "2026", "confidence": "中"},
    {"id": "fairwaterATL", "name": "Fairwater Atlanta", "operator": "Microsoft 自建（OpenAI 與 Microsoft，GB200）", "planned": 636, "energized": 636, "accepted": 636, "billable": 636, "status": "Epoch AI：Buildings 1–4 已運轉，636 MW-IT（2026-06-14）", "next": "—", "date": "2026", "confidence": "中"}]})

# ---------- 評價 ----------
V = C['valuation']
V.update({"price": 529.76, "shares": 7.453, "netDebt": round(46.136 - 76.843, 3), "tax": 0.20, "sbc": 12.405, "evYear": 2,
  "holdings": [["OpenAI（2026-03 投後估值）", 852.0, 0.25, "10-K 附註 1：持股約 25%（as-converted；重組時約 27%）[Verified]；投後估值 852B（2026-03 結案、募資 122B；新創自報估值 [Interested-party]）；Anthropic 持股比例未揭露 → 不估值（列資料缺口；帳列成本法約 12.4 內）", False]],
  "debtLike": []})

V['capm'] = {"beta": round((1.10 + 1.1741) / 2, 4),
  "betaSources": [{"vendor": "StockAnalysis（Yahoo Finance 同值 1.10，推測同源，合算 1 個）", "value": 1.10, "period": "5 年", "frequency": "月", "asOf": "2026-10-08", "ref": "val.beta.sa"},
                  {"vendor": "GuruFocus", "value": 1.1741, "period": "3 年", "frequency": "頁面未寫明", "asOf": "2026-10-08", "ref": "val.beta.gurufocus"}],
  "erp": 0.05, "kdPretax": 0.0626, "betaSens": [1.10, 1.1741], "kdNote": "Moody's Aaa 公司債殖利率 6.26%（2026-10-06）類比；MSFT 個別長債殖利率找不到 [Analogy]",
  "note": "WACC＝E/(D+E)×(rf＋β×ERP)＋D/(D+E)×kd×(1−稅率)；E＝現價 × 流通股數（latestQuarter.sharesOut）、D＝評價日債務面額（latestQuarter.debtPrincipal）；β 1.137（獨立來源平均，見 betaNote）、rf 5.28%（10 年美債 2026-10-07）、ERP 5%、kd 6.26%（Moody's Aaa 公司債殖利率類比；MSFT 個別長債殖利率找不到）；稅率 20%（FY27 指引）。valuation.wacc 為 null 時採 CAPM。",
  "betaNote": "β＝獨立來源平均（MAG 對照表 r1 C13）：StockAnalysis 5 年月 1.10（Yahoo 同值，推測同源，只算 1 個）、GuruFocus 3 年 1.1741 → 1.137；敏感度＝獨立來源區間 1.10–1.174 [Derived]"}
cloud = copy.deepcopy(A['valuation']['segmentMultiples']['cloud'])
V['segmentMultiples'] = {
  "software": {"label": "企業軟體", "peers": [{"ticker": "ORCL", "name": "Oracle", "ntmEvEbitda": 9.54, "ref": "val.peer.sw.orcl"}, {"ticker": "SAP", "name": "SAP", "ntmEvEbitda": 15.02, "ref": "val.peer.sw.sap"},
     {"ticker": "CRM", "name": "Salesforce", "ntmEvEbitda": 10.8, "ref": "val.peer.sw.crm"}, {"ticker": "ADBE", "name": "Adobe", "ntmEvEbitda": 6.6, "ref": "val.peer.sw.adbe"}, {"ticker": "NOW", "name": "ServiceNow", "ntmEvEbitda": 20.65, "ref": "val.peer.sw.now"}]},
  "ads": copy.deepcopy(A['valuation']['segmentMultiples']['ads']),
  "devGaming": {"label": "遊戲與 PC／裝置", "peers": [{"ticker": "TTWO", "name": "Take-Two", "ntmEvEbitda": 50.9, "ref": "val.peer.game.ttwo（TTM，非 NTM）"}, {"ticker": "NTDOY", "name": "Nintendo", "ntmEvEbitda": 12.63, "ref": "val.peer.game.ntdoy（TTM）"},
     {"ticker": "NTES", "name": "NetEase", "ntmEvEbitda": 8.23, "ref": "val.peer.game.ntes（TTM）"}, {"ticker": "HPQ", "name": "HP", "ntmEvEbitda": 7.6, "ref": "val.peer.pc.hpq（TTM）"},
     {"ticker": "DELL", "name": "Dell", "ntmEvEbitda": 22.1, "ref": "val.peer.pc.dell（TTM，含 AI 伺服器）"}, {"ticker": "0992.HK", "name": "Lenovo", "ntmEvEbitda": 14.47, "ref": "val.peer.pc.0992.hk（TTM）"}]},
  "cloud": cloud}
V['segmentMultiplesNote'] = "分部 EV/EBITDA：各非 AI 分部 × 所屬同業組 EV/EBITDA 中位數（legacyBiz.lines.peer；每組至少 3 家）：企業軟體 ORCL、SAP、CRM、ADBE、NOW（NTM，不含 MSFT 本身）；數位廣告共用同業組 META、PINS、TTD、APP、RDDT、SNAP（C18，data/peers_ads_20261009.json 自 amazon/ 複製）；遊戲與 PC TTWO、NTDOY、NTES、HPQ、DELL、Lenovo（TTM，NTM 找不到）；非 AI 雲端＝共用同業組 8 家中位數（C12，data/peers_cloud_20261008.json 自 amazon/ 複製）；AI 雲端 × valuation.evEbitda（6×）；非 AI 事業倍數＝錨定年度各線 EBITDA 加權（valuation.legacyEvEbitda 非 null 時為手動覆蓋）"
V['ownMultiple'] = {"value": 15.67, "ev": 3990, "ntmEbitda": 254.683, "ref": "val.msft.ownNtm",
  "note": "公司自身 NTM EV/EBITDA＝EV 3.99T（StockAnalysis 2026-10-07，含租賃負債；事實總帳 val.ev）÷ NTM EBITDA 254.7（MarketScreener 共識 FY27 239.452、FY28 295.606，FY1 權重 0.7288＝(2027-06-30 − 2026-10-07) ÷ 365）[Derived]；只用於評價口徑敏感度（非 AI 分部改用自身倍數）"}
V['holdingsNote'] = "持股清單：[名稱, 估值（100%，US$bn）, 持股比例, 說明, 上市（true＝不折價）]；價值＝Σ 估值 × 比例 ×（1 − 折價；上市持股不折價），自淨負債扣除（分部加總項）。流動性折價 20%（0%–40%，[Assumed]；對照表 r1 第 5 節第 9 條、C8 c）。OpenAI 25% × 852 × 0.8＝170.4；Anthropic 比例未揭露，不估值（少計約 5–14）"
V['legacyEvEbitdaNote'] = "非 AI 分部 EV/EBITDA：null＝各分部同業倍數加權。"

# ---------- Tokenomics ----------
snap = json.load(open(sys.argv[3], encoding='utf-8'))
C['tokenomics'] = copy.deepcopy(A['tokenomics'])
C['tokenomics']['names'] = list(snap['items'])
C['tokenomics']['optional'] = [n for n, v in snap['items'].items() if v.get('missing')]

# ---------- 定價 ----------
P = copy.deepcopy(A['pricing'])
P.update({"_note": P['_note'].replace('kevidence_amazon_20261008', 'kevidence_microsoft_20261008'),
  "axisLabels": {"low": "低 長約下緣", "base": "基準 C1 混合", "high": "高 現貨改用 Azure 牌價"},
  "kLong": {"low": 0.45, "base": 0.76, "high": 0.76}, "kSpot": {"low": 1.76, "base": 1.76, "high": 7.66}, "longShare": {"low": 0.7, "base": 0.7, "high": 0.7},
  "kNote": {"kLong": "長約 k：IREN–Microsoft 0.76（一手 MW）、Nscale 約 0.88（金額二手）、Anthropic–Azure 0.39–0.71（年期假設）、Nebius ≤1.33 → 基準 0.76；低＝長約區間下緣 0.45（Anthropic 隱含）[Assumed]（kevidence longContracts）",
            "kSpot": "現貨 k：基準 1.76＝與 Nebius／CRWV 同一市場現貨指數（對照表 r1 C1）；高＝Azure 牌價換算（H100 隨需 12.29 ÷ IF_GPUhrEcon H100 1.605＝7.66，kevidence az.h100.od；與 Amazon 同取隨需層；牌價不是大客戶成交價，只作上限；替代：3 年預留 3.36 → 混合 1.54）",
            "longShare": "長約占比 70%（OpenAI 約占商用 RPO 45%＋Anthropic 等長約；45%–100% [Assumed]）"},
  "customNote": "自研晶片每 MW 錨＝同期 NVIDIA 世代 IF_HoldEcon × 1.0（對照表 r1 C2）：Maia 100 → Hopper（期初未列，部署量不存在）、Maia 200（2026 起量產）→ GB300（v0.1 交付前修訂，查核 b；原 v0.1 誤對 Hopper）；敏感度 × 0.7（收入端單邊與收入＋資本支出雙邊兩版）。Tokenomics 無 NonNV 具名範圍（IF_NonNVRatio、IF_NonNVCostRatio、IF_NonNV_Maia 為 missing），已列缺口回報。",
  "customCapexFactor": 1.0, "customCapexNote": "自研晶片每 MW IT 資本支出係數：預設 1（Maia 200 的資本支出以同期 NVIDIA GB300 的 TK_CapexIT 計）；雙邊敏感度時與 customFactor 同乘 0.7（收入錨與資本支出同降）；機房成本不變 [Assumed]",
  "chips": [{"key": "h100", "label": "Hopper（H100／H200）", "tk": "H100", "custom": False, "mixOpen": 0.20, "mixAdds": [0] * 5},
            {"key": "gb200", "label": "GB200", "tk": "GB200", "custom": False, "mixOpen": 0.45, "mixAdds": [0] * 5},
            {"key": "gb300", "label": "GB300", "tk": "GB300", "custom": False, "mixOpen": 0.35, "mixAdds": [0.6, 0.2, 0, 0, 0]},
            {"key": "maia", "label": "Maia 200", "tk": "GB300", "custom": True, "mixOpen": 0, "mixAdds": [0.1, 0.1, 0.1, 0.1, 0.1]},
            {"key": "vr200", "label": "Vera Rubin（VR200）", "tk": "VR200", "custom": False, "mixOpen": 0, "mixAdds": [0.3, 0.7, 0.9, 0.9, 0.9]}],
  "mixNote": "期初在役（評價日 2026-06-30）：Hopper 20%、GB200 45%、GB300 35%（v0.1a mw.genMix.open，[Assumed]；Maia 部署量未揭露，期初不另列）。新增 MW：Maia 200 10%（對應 GB300；「continues to scale」，部署量不存在，[Assumed] 0%–20%）；NVIDIA 部分 FY27 GB300 60%／VR200 30%、FY28 GB300 20%／VR200 70%、FY29 起 VR200 90% [Assumed]（VR200 2026 下半年起出貨）。在役占比＝各世代累計 MW ÷ 對外 AI MW。"})
C['pricing'] = P

# ---------- 資本支出模型 ----------
CM = copy.deepcopy(A['capexModel'])
CM.update({"_note": "資本支出（MAG v0.1b，對照表 r1 第 5 節第 6 條、D3）：首期＝MAX(全年 AI 成長型＋非 AI, 公司指引) − 年初至今實際（0）；之後＝新增對外 AI MW ÷ 對外比例 × 每 MW 成本（Σ 新增世代占比 × (TK_CapexIT＋自建比例 × TK_CapexFacility)）＋GPU 汰換（只換 IT；壽命 6 年）＋非 AI 資本支出（各線全年營收 × legacyBiz.lines.cx）。口徑＝公司口徑（含以融資租賃取得的資產；FY26 145.3，其中融資租賃 28.1）：融資租賃取得的資產在取得時視同現金支出（偏保守；之後的融資租賃本息只計在帳部分）。",
  "extShare": 0.6, "extShareNote": "對外 AI MW 占 AI 總 MW 比例 60%（50%–75%，[Assumed]；v0.1a mw.ai.externalShare）：自用 AI（M365／GitHub Copilot、MAI）同樣需要資本支出，影子收入只作對照。期初對外計費 MW 3,250 含 neocloud 租用對外 550：依 C20 列 capexModel.rentedExt，自有 AI MW＝3,250 ÷ 60% − 550＝4,867（v0.1a 自有 4,500 加上自有部分的自用；差額來自對外比例套在含租用的總數上）",
  "selfBuild": 0.6, "selfBuildNote": "機房自建比例 60%（自建＋融資租賃；[Assumed] 40%–80%）：公司口徑資本支出含融資租賃取得的機房；營業租賃部分（約 40%）不計機房資本支出，租金由未起租租賃的營業部分承擔（leases.uncommenced）。未起租 329.1B 遠大於在帳租賃，顯示租用機房占比上升",
  "segDaRunRate": 44.088, "segDaRunRateNote": "最新季（FY26Q4）折舊、攤銷及其他 11.022 × 4（現金流量表；分部 D&A 不揭露，以合併數代替；含無形資產攤銷約 1.2／季）[Derived]：校準非 AI 折舊年限",
  "guideNote": "FY27 公司未給金額（只說「年增」）；以曆年 2026 約 175 − FY26 下半年實際 72.9＝FY27 上半年 102.1，下半年＝上半年 → 204.2 [Derived]（FY27Q1 指引 >50 吻合；共識 FY27 現金口徑 192.8）",
  "aiNetShare": 0.8, "aiNetShareNote": "期初 AI PP&E 淨額 ÷ 毛額 80%（在役機隊約 FY24 末 0.9、FY25 末 2.3、FY26 末 4.5 GW-IT → 平均約 1.2 年、IT 壽命 6 年；[Assumed]，區間 65%–90%）：AI 增量 ROIC 的期初投入資本",
  "rentedExt": {"open": 550, "path": [756, 756, 756, 756, 756], "rentMW": round(RENT_GROSS - OPEX_MW, 2),
    "note": "租用的對外 AI MW（MAG 對照表 r1 C20）：期初 550＝neocloud 在役 1,100 × 對外 50%（v0.1a nc.mwLeased.open）；FY27 起加 Nscale 對外 205.5 → 756，之後到期假設以相同條件續約（[Assumed]）。收入照算（含在對外計費 MW 內）、不計資本支出、投入資本與折舊；每 MW 年租金＝已揭露合約全額年租金 " + f"{RENT_FULL:.2f}" + " ÷ " + f"{RENT_MW:,}" + " MW＝" + f"{RENT_GROSS:.2f}" + " 扣 Tokenomics 營運成本 2.34（已在 AI EBITDA 內）＝淨額 [Derived]"},
  "nonAiLife": 10, "nonAiLifeRange": [8, 15], "nonAiLifeNote": "非 AI 折舊年限預設 10 年（區間 8–15 年）[Assumed]；FY27 起資料中心與辦公建物耐用年限 15 → 25 年、伺服器 2–6 年（10-K、法說）；差額列 D&A 對帳殘差（MAG 對照表 r1 C17）",
  "roicYear": 2, "roicYearNote": "打平 k 的錨定期（模型第 3 期＝FY29，截至 2029-06；對照表 r1 第 5 節第 10 條）"})
C['capexModel'] = CM

# ---------- 關聯方 ----------
fy = lambda d: {f"FY{(y + 1) % 100:02d}": round((d[y] + d[y + 1]) / 2, 3) for y in range(2025, 2030)}
oaiAz = {2025: 11.905, 2026: 23.81, 2027: 35.714, 2028: 47.619, 2029: 59.524, 2030: 71.429}
oaiRs = {2025: 2.614, 2026: 6.841, 2027: 10.959, 2028: 15.685, 2029: 4.515, 2030: 0.0}
antAz = {2025: 0.0, 2026: 0.0, 2027: 2.415, 2028: 4.831, 2029: 4.831, 2030: 4.831}
C['related'] = {"_note": "關聯方：OpenAI／Anthropic 模型快照與合約並排，只讀、不回饋計算（對照表 r1 D5、第 5 節第 11 條）；對手方集中度＝合約年化 ÷ 模型對外 AI 雲端收入。日曆年快照換財年：FY＝(前一曆年＋當年) ÷ 2（6 月財年各含半年，[Derived]）",
  "contracts": [{"name": "OpenAI 增量 Azure 承諾（$250B；年期未揭露，依 openai/ 模型 2025–2030）", "annual": round(250 / 6, 6), "startFY": 2026, "note": "2025-10-28 公告（10-Q FY26Q1；事實總帳 oai.azureCommit）[Interested-party]；年化＝總額 ÷ 6 年（openai/ v0.6 INP_182／183，[Analogy]）"},
                {"name": "Anthropic Azure 承諾（$30B，至多 1 GW）", "annual": round(30 / 6.5, 6), "startFY": 2027, "note": "2025-11-18 公告（事實總帳 ant.azureCommit）[Interested-party]；年期依 anthropic/ v0.1 INP_135／136（2026-11 至 2033-05，約 6.5 年）[Analogy]"}],
  "snapshots": [{"name": "openai/ 模型：OpenAI 付 Azure（合約實付）", "series": fy(oaiAz), "note": "openai/ v0.6（main）成品 Cost K09 日曆年 → 財年（只讀，不連動）；FY26 實際：來自 OpenAI 營收 24.1（含營收分成，10-K）"},
                {"name": "openai/ 模型：OpenAI 付 Microsoft 營收分成（20%，上限 38B 至 2030；對照）", "series": fy(oaiRs), "note": "openai/ v0.6 Revenue R28（只讀）；營收分成只列對照，不進評價（已入帳營收除外）"},
                {"name": "anthropic/ 模型：Anthropic 付 Azure", "series": fy(antAz), "note": "anthropic/ v0.1（main）成品 Cost K35（只讀，不連動）"}],
  "rpoShareMax": 0.46, "rpoShareNote": "OpenAI 約占商用 RPO 45%（FY26Q2，約 281B）＋Anthropic 30B → 約 311 ÷ 商用 RPO 678 ≈ 46%（FY26Q4 最新比例找不到；事實總帳 oai.rpoShare.fy26q2）"}

# ---------- 季度層 ----------
QT = copy.deepcopy(A['quarterly'])
keys = ['FY27Q1', 'FY27Q2', 'FY27Q3', 'FY27Q4', 'FY28Q1', 'FY28Q2']
QT['quarters'] = [{"key": k, "label": f"Q{k[-1]} FY{k[2:4]}", "period": 0 if k.startswith('FY27') else 1, **({"reportNote": "預計 2026-10-28"} if k == 'FY27Q1' else {})} for k in keys]
QT['periodNames'] = ["FY27", "FY28"]
QT['focus'] = 'FY27Q1'
QT['driver'] = {"type": "mw", "endStart": OPEN, "endStartNote": "評價日（2026-06-30）對外 AI MW-IT 未揭露：三法交叉估計 3,250（自有在役 AI 4,500 × 對外 60%＋neocloud 租用 1,100 × 50%，事實總帳 split.extMW）[Derived／Assumed]"}
QT['revenueAnchor'] = {"value": 90.007, "label": "FY26Q4 實際營收", "tag": "Interested-party"}
QT['guidance'] = {"FY27Q1": {"revenue": [89.85, 90.95], "adjOpInc": [CF['nextQAdjOpLo'], CF['nextQAdjOpHi']]}}
QT['guidanceMeta'] = {"source": "2026-09-02 8-K（FY27Q1 調整後指引）：營收 89.85–90.95B；營業成本 29.6–29.8、營業費用 16.8–16.9 → 營業利益 43.15–44.55 [Derived]；資本支出 >50（含融資租賃）[Interested-party]", "tag": "Interested-party"}
QT['periodGuidance'] = {"0": {"capex": {"range": [CAPEX27, CAPEX27], "less": 0.0}}}
QT['periodGuidanceNote'] = "FY27 資本支出公司未給金額：2 ×（曆年 2026 約 175 − FY26 下半年實際 72.9）＝204.2 [Derived]；首期為全年模型，年初至今 0；只用於判斷季度差距是否為「拆法」"
QT['actuals'] = {k: {"revenue": None, "adjEbitda": None, "adjOpInc": None, "capex": None, "mw": None, "source": None, "date": None, "tag": None} for k in keys}
C['quarterly'] = QT

VR = copy.deepcopy(A['varianceReasons'])
for x in VR['list']:
    t = x['text']
    t = t.replace('零售／廣告／訂閱／AWS 非 AI', 'M365 雲端／授權／產業＋Frontier／搜尋與廣告／Windows＋XBOX／Azure 非 AI').replace('首期資本支出＝2026 指引約 220 − 上半年實際；之後由 MW 推導（Tokenomics 每 MW 成本）', '首期資本支出＝FY27 推估 204.2（曆年 2026 約 175 換算，含融資租賃）；之後由 MW 推導（Tokenomics 每 MW 成本）')
    t = t.replace('首期取公司指引約 220；公司未給 2027 以後指引', '首期取 FY27 推估 204.2（公司口徑含融資租賃；共識推定現金口徑）；公司未給 FY28 以後指引').replace('對帳列顯示 2026 指引隱含的 AI 建置量約為 MW 路徑的 2 倍（共識口徑未揭露）', '基準情境 MW 速度由首期資本支出解出（對帳落差 0）')
    t = t.replace('分部 EBITDA 率依 C5 放大分攤 D&A，近四季收斂到近三年平均', '各線 EBITDA 率＝分配營業利益率＋D&A 按營收分攤，五期固定').replace('；年度總額線性分到各季，未建模第四季季節性', '；年度總額線性分到各季，未建模季節性')
    x['text'] = t
VR['list'].append({"scope": "annual", "period": "FY29", "metric": "nd", "vs": "consensus", "type": "已知限制", "text": "共識 FY29 淨負債 +99.6 與 FY28 −38.8、FY29 FCF 89.0 方向矛盾（MarketScreener 原文，2026-10-09 複查同值，原因無法確認）→ 資料可疑，本期差距不具意義"})
C['varianceReasons'] = VR

# ---------- 同業（市場比較用，不進目標價）----------
PE = copy.deepcopy(A['peers'])
PE['textHtml']['headerTip'] = PE['textHtml']['headerTip'].replace('零售同業的 NTM EV/EBITDA 用於非 AI 分部倍數（評價頁）', '企業軟體同業的 NTM EV/EBITDA 用於非 AI 分部倍數（評價頁）')
PE['textHtml']['readingTip'] = "MSFT 列為整體公司（M365、授權、遊戲、搜尋＋Azure），倍數不可直接與純 AI 基礎設施同業比較；前瞻倍數見本模型列。"
PE['textXlsx']['notes'] = ["讀法：MSFT 為整體公司（M365、授權、遊戲、搜尋＋Azure），倍數不可直接與純 AI 基礎設施同業比較；",
                           "CRWV 以租約取得機房，加入營業租賃負債後倍數上升；MSFT 自建與租賃並行（另有未起租承諾 329.1B 尚未入表），並向 CoreWeave、Nebius、IREN、Nscale、Lambda 租用 GPU。",
                           "前瞻：見下方 MSFT 模型列。此表為市場比較用，不進入目標價。"]
PE['software'] = copy.deepcopy(V['segmentMultiples']['software']['peers'])
PE['softwareNote'] = "企業軟體同業 NTM EV/EBITDA（ORCL、SAP、CRM、ADBE、NOW；data/microsoft_facts val.peer.sw.*，2026-10-07，不含 MSFT 本身）；中位數 10.8× 用於 M365、授權、產業＋Frontier 的分部倍數 [Derived]"
C['peers'] = PE

# ---------- 文字 ----------
T = copy.deepcopy(A['texts'])
T.update({
  "fy0EquityNote": "«P0»＝全年模型（年初至今 0）：不發股；股利每季 0.98、回購基準 22.3／年",
  "cashTaxNote": "FY26 有效稅率 19%（FY24 18%、FY25 18%）；FY27 指引約 20% → 現金稅率 20%",
  "sourceLine": "FY26 10-K（2026-07-29）＋FY26Q4 財報新聞稿與法說（2026-07-29）＋2026-09-02 8-K 分部重編 · 口徑分層",
  "mwYearEndNotes": {"2024": "FY24 末對外 AI MW-IT 估 640（自有 AI 約 900 × 對外 60%＋neocloud 租用約 200 × 50%；Omdia 2023–24 Hopper 約 635k 顆，[Derived／Assumed]；只用於 GPU 汰換批次）",
                     "2025": "FY25 末對外 AI MW-IT 估 1,630（自有約 2,300 × 60%＋租用約 500 × 50%；FY25 新增 >2 GW（設施）× AI 70–90% ÷ PUE 1.2，[Derived／Assumed]）",
                     "2026": "FY26 末（評價日）對外 AI MW-IT 3,250（自有 4,500 × 60%＋租用 1,100 × 50%；三法交叉，事實總帳 split.extMW）[Derived／Assumed]"},
  "guideLine": "FY27 公司指引：資本支出年增（未給金額；曆年 2026 約 $175B、FY27Q1 >$50B，皆含融資租賃）；FY27Q1 營收 $89.85–90.95B、營業利益率全年下降 <1pt；有效稅率約 20% [Interested-party]",
  "taxNote": "FY24–FY26 有效稅率 18%／18%／19%；FY27 指引約 20%（採用）；IRS 移轉訂價補稅主張 28.9B 為或有事項，模型不計",
  "revMwCompare": "Microsoft 對照值：IREN 9.7B／5 年 ↔ 200 MW-IT 隱含 9.70 US$m/MW-年（k 0.76），只作對照",
  "rpoNote": "商用 RPO 678B（約 30% 於 12 個月內、加權平均 2.3 年；約 45% 來自 OpenAI，FY26Q2）含 M365 與非 AI 雲端，只作對照（模型排程歸 0：AI 雲端收入＝MW × 每 MW）",
  "leaseNote": "Microsoft 資料中心自建與租賃並行（融資租賃負債 66.6、營業 21.9、未起租 329.1）；另向 neocloud 租用 GPU 算力（租金屬營運成本）",
  "offBalanceLeaseTerm": "FY27–FY33 起租、租期 1–20 年；逐年時程與營業／融資拆分未揭露",
  "leaseLiabNote": "10-K 附註 13；融資租賃加權平均折現率 4.5%、剩餘 13 年；營業 3.7%、6 年",
  "capexGuideSource": "FY27 公司未給金額：曆年 2026 約 175（2026-07-29 法說，含融資租賃、租賃改分類後）換算 [Derived]",
  "prepayCheck": "不適用（Microsoft 無客戶預付融資機制）。",
  "mwFacts": "公司：FY25 新增 >2 GW、FY26 Q2–Q4 各約 1 GW（口徑不明，推定設施）；總產能約兩年翻倍；FY26 AI 總產能 +80% 以上。第三方：Bloomberg 總容量約 12 GW（2026-09）、2032 目標約 38 GW；Epoch AI Fairwater Wisconsin 369、Atlanta 636 MW-IT 在役 [Interested-party／第三方]。",
  "costMwNote": "每 MW 成本＝Σ 新增世代占比 × (TK_CapexIT＋自建比例 60% × TK_CapexFacility)（Tokenomics v5.27；US$m/MW-IT）",
  "ppeOpenNote": "評價日 PP&E 淨額 313.1（10-K；含融資租賃資產 67.3）[Verified]；D&A 分池：AI 期初毛額估計（評價日 AI 總 MW × Tokenomics 每 MW 成本）＋非 AI（其餘，壽命以最新季 D&A 年化校準）",
  "facilityName": "無已揭露的循環信用額度（不列入瀑布）",
  "newDebtRateNote": "瀑布新債利率以 Moody's Aaa 公司債殖利率 6.26% 類比（MSFT 個別長債殖利率找不到）[Analogy]",
  "concentration": {"title": "客戶集中：OpenAI 約占商用 RPO 45%，加 Anthropic 約 46%", "detail": "10-K：商用 RPO 678B（加權 2.3 年）；FY26Q2 法說 OpenAI 約占 45%（約 281B）；OpenAI 增量 Azure 承諾 250B（年期未揭露）、Anthropic 30B（至多 1 GW）。FY26 來自 OpenAI 營收 24.1（占 7.3%，含營收分成）。Microsoft 同時持有 OpenAI 約 25%（持股與關聯方頁）。"},
  "sharesNote": "稀釋加權股數 7.453bn（FY26）；期末流通 7.427bn",
  "netDebtNote": "不含租賃（債務面額 46.1 − 現金及短期投資 76.8）",
  "otherRevNote": "非算力服務預設 0（各分部另列）",
  "rvRevNote": "Microsoft 對照：IREN 合約隱含 9.70（IT）US$m/MW-年",
  "rvCostNote": "Tokenomics 每 MW 全成本（IT＋自建機房 60%）約 $44m/MW-IT",
  "priceNote": "每 MW 年收入＝Tokenomics 持有成本（世代加權）× k；價格軸低／基準／高見 pricing；不建模續約價格",
  "prepayCoverNote": "不適用（無客戶預付）", "prepayOpenNote": "不適用（無客戶預付）", "ytdEquityNote": "不適用（年初至今 0）", "prepayRiskNote": "不適用（Microsoft 無客戶預付）。",
  "callVsFiling": [["RPO", "OpenAI 約占 45%（FY26Q2）", "商用 RPO 678B（加權 2.3 年）", "RPO 只作對照；營收由 MW × 每 MW 與各分部驅動"],
                   ["期初現金", "未提", "現金 20.9、短期投資 55.9", "期初用現金＋短期投資 76.8"],
                   ["FY27 CapEx", "年增（未給金額）；曆年 2026 約 175；Q1 >50", "FY26 145.3（含融資租賃 28.1）", "首期＝204.2（換算）；之後由 MW 推導"],
                   ["AI 容量", "FY26 Q2–Q4 各約 1 GW；兩年翻倍", "未揭露 MW", "三法交叉估計；不以營收反推"],
                   ["在帳租賃", "未量化", "融資 66.6、營業 21.9；未折現 114.4", "聽 10-K 到期表，固定輸入"],
                   ["表外租賃", "租賃改分類（更多營業租賃）", "尚未起租 329.1（FY27–FY33）", "平均分 16 季起租、15 年"],
                   ["債務", "無新發", "面額 46.1（13 次發行）", "逐批；到期由瀑布再融資"],
                   ["AI run-rate", "AI 業務年化 >37B（FY26Q3，含 Copilot）", "未揭露", "只作驗證列，不作輸入（對照表 C3）"]],
  "commitments": [["長期債（13 次票據發行，評價日）", "46.14", "FY27 9.25、FY29 2.00、FY31 0.50；其後 34.4", "是（帳面 40.3）", "排程還本（由瀑布再融資）"],
                  ["在帳營業＋融資租賃（未折現）", "114.39", "FY27 13.2／FY28 11.6／之後遞減／其後 61.7", "是（PV 88.5）", "在帳現金租金（固定）"],
                  ["尚未起租租賃（未折現）", "329.10", "FY27–FY33 起租、1–20 年", "否", "表外現金租金（營業部分）"],
                  ["採購承諾（主要為資料中心；含 take-or-pay）", "194.06", "FY27 169.0", "否", "已含在 CapEx 與 EBITDA（可能含 neocloud 算力；不另列）"],
                  ["建設承諾", "34.57", "FY27 29.8", "否", "已含在 CapEx"],
                  ["OpenAI 出資承諾（未出資）", "1.10", "承諾 13.0、已出資 11.9", "否", "列入 jvCommit（首期）"]],
  "consistency": [["FY27 CapEx 年增", "法說", "FY26 145.3", "公司未給金額；首期依曆年 2026 約 175 換算 204.2"],
                  ["AI run-rate >37B", "法說（含 Copilot）", "無", "模型 AI 雲端為 MW × 每 MW；差距列驗證（對照表 C3）"],
                  ["兩年翻倍", "公司目標", "無", "總容量未揭露，積極情境以 Bloomberg 12 GW 為基數"],
                  ["FCF", "未給", "FY26 FCF 67.0（現金口徑）", "模型五期 FCF 另列"]],
  "sources": [["Verified", "FY26 10-K（2026-07-29，經查核）：營收 331.8（FY27 新口徑 Agents and Infra 268.1、Devices and Consumer 63.7）；營業利益 155.2；D&A 38.5；SBC 12.4；利息費用 3.1；淨利 133.7（含 OpenAI 利得 6.5）。"],
              ["Verified", "資產負債（2026-06-30）：現金 20.9、短期投資 55.9；PP&E 淨額 313.1；債務面額 46.1；融資租賃負債 66.6、營業 21.9；尚未起租 329.1；股東權益 442.4。"],
              ["Verified", "現金流（FY26）：CFO 182.9；現金資本支出 115.9（含融資租賃 145.3）；股利 26.4、回購 22.3；無新發債、償還 3.0。"],
              ["Interested-party", "2026-09-02 8-K：Azure（新定義）FY26 101.9、Q4 29.4（+42%）；FY27Q1 指引營收 89.85–90.95、資本支出 >50；曆年 2026 資本支出約 175（法說）。"],
              ["Interested-party", "關聯方：OpenAI 持股約 25%、來自 OpenAI 營收 24.1、增量 Azure 承諾 250；Anthropic Azure 承諾 30（至多 1 GW）。"]],
  "headerCards": [["產能", "FY26 Q2–Q4 各約 1 GW · 兩年翻倍目標 · AI 在役 MW 未揭露（三法交叉估計）"],
                  ["合約", "OpenAI 增量 Azure $250B · Anthropic $30B（≤1 GW）· 商用 RPO 678B（OpenAI 約 45%）"],
                  ["年度指引", "FY27 資本支出年增（曆年 2026 約 175）· Q1 營收 89.85–90.95 · 營業利益率降 <1pt"]],
  "legacyMarginNote": "各線 EBITDA 率＝分配營業利益率（分部合計＝實際）＋D&A 按營收分攤 11.6%（分部 D&A 不揭露）；Azure 非 AI 雲端＝殘差（Azure EBITDA 估計 − AI 雲端 EBITDA 估計），[Derived]",
  "scenarioTip": "三個情境只改容量軸（對外 AI MW 路徑，對照表 r1 第 5 節第 1 條）：保守＝只計第三方（Epoch AI）可辨識站點；基準＝依 FY27 資本支出隱含的建置量（C10）；積極＝依公司總產能兩年翻倍的目標。價格軸（k 低／基準／高）另列 3 × 3 矩陣，三容量情境都用基準 k。毛 CapEx 由 MW 公式自動增減。切換情境只改寫 MW 路徑與預建 MW；你手動調整的其他數字會保留，按頁首「重設」才回到預設值。"})
C['texts'] = T
C['methodology'] = copy.deepcopy(A['methodology'])

json.dump(C, open(sys.argv[4], "w", encoding='utf-8'), ensure_ascii=False, indent=1)
print('lines:', [(x['key'], x.get('fyBase') or x.get('priorStub'), round(x.get('g0') or x.get('g4q'), 4), x.get('m0'), x.get('ebitdaTTM')) for x in lines])
print('Azure OI', round(oi_az, 3), 'devgame om', round(om['devgame'], 4), 'capex27', CAPEX27)
