/*APP_START*/
// ===== 公司資料：全部來自 company.json（建置時注入為 COMPANY_DATA；Excel 建置讀同一檔）=====
// 本段只做「由原始輸入推導」：情境 Billable 比率、Q3 新增 RPO 權重、債務合計與平均利率等。
var PERIODS = COMPANY_DATA.periods,
  PERIOD_YEARS = COMPANY_DATA.periodYears,
  UPDATE_DATE = COMPANY_DATA.meta.updateDate,
  SOURCE_ORDER_NOTE = COMPANY_DATA.meta.sourceOrderNote,
  ACTUAL_1H = COMPANY_DATA.ytdActual, // v4.5：年初至今實際（原 actual1H）
  CALQ = COMPANY_DATA.cal, // v4.5：期間與日期（calendar_q.py 推算）
  PERIOD_FY = PERIODS.map(p => 2000 + +p.slice(2)), // 5a：各期財年年份（FYyy → 20yy）
  RPO_SCHEDULED_SHARE = COMPANY_DATA.rpo.scheduledShare,
  RPO_BUCKET_W = COMPANY_DATA.rpo.bucketWeights,
  RPO_Q3ADD_W = RPO_BUCKET_W.map(e => e / RPO_SCHEDULED_SHARE),
  LEASE_CASH_ON_BAL = COMPANY_DATA.leases.onBalanceCash,
  DEBT_AMORT = COMPANY_DATA.debt.amortization,
  LEASE_AFTER_FY30 = COMPANY_DATA.leases.afterFY30,
  LATEST_Q = COMPANY_DATA.latestQuarter,
  CALL_FACTS = COMPANY_DATA.callFacts,
  SC_MWP = COMPANY_DATA.scenarios.mwPath, // v0.1b：已連網 MW＝MIN(合約上限, 前期＋併網速度×期間長度)；首期期末三情境共用
  SC_ACC = Object.fromEntries([`low`, `base`, `high`].map(k => [k, PERIOD_YEARS.reduce((a, L, i) => (a.push(Math.min(SC_MWP.contracted[k][i], i === 0 ? (SC_MWP.connectedStart ?? COMPANY_DATA.defaults.billableOpen + SC_MWP.pace[k] * L) : a[i - 1] + SC_MWP.pace[k] * L)), a), [])])), // MAG v0.1b：connectedStart＝null 時首期自評價日（期初計費 MW）起算
  PRICING = COMPANY_DATA.pricing || null, // MAG v0.1b：對外 AI 雲端定價（company.json → pricing；Tokenomics × k）
  TKV = COMPANY_DATA.tk || {}, // MAG v0.1b：Tokenomics 快照基準值 {名稱: {世代代碼: 值} 或單值}（建置時由 calendar_q.tk_base 注入）
  K_AX = [`low`, `base`, `high`],
  kAxQ = ax => { const a = K_AX[ax] ?? ax, p = PRICING; return p.longShare[a] * p.kLong[a] + (1 - p.longShare[a]) * p.kSpot[a] }, // 混合 k＝長約占比 × k_長約＋(1 − 長約占比) × k_現貨
  // MAG v0.1b：每 MW 年收入路徑。各世代在役 MW＝期初 × 期初占比＋Σ 本期新增（MAX(0, 期末 − 前期末)）× 新增占比；
  // 每 MW 年收入＝Σ 世代占比 × TK_HoldEcon × 晶片係數（自研 × customFactor）÷ 1000 × k；營運成本（TK_OpexGW）與參考付費營收（TK_RevGWFleet）同權重、不乘係數。Excel 同式。
  revPathQ = (acc, open, ax, cf) => {
    const k = kAxQ(ax), G = PRICING.chips.map(c => ({ ...c, he: TKV.IF_HoldEcon[c.tk] * (c.custom ? cf ?? PRICING.customFactor : 1), op: TKV.IF_OpexGW[c.tk], rf: TKV.IF_RevGWFleet[c.tk] }));
    let mw = G.map(c => c.mixOpen * open), prev = open;
    return acc.map((x, r) => {
      const add = Math.max(0, x - prev); prev = x;
      mw = mw.map((m, j) => m + add * G[j].mixAdds[r]);
      const T = mw.reduce((a, b) => a + b, 0), W = f => mw.reduce((a, m, j) => a + m * G[j][f], 0) / Math.max(T, 1e-9);
      const hold = W(`he`), opex = W(`op`), ref = W(`rf`);
      return { k, gens: [...mw], tot: T, hold, opex, ref, rev: hold * k / 1e3, opexMW: opex / 1e3, capRatio: hold * k / Math.max(ref, 1e-9) }
    })
  },
  SC_REV = PRICING ? Object.fromEntries(K_AX.map(k => [k, revPathQ(SC_ACC[k], COMPANY_DATA.defaults.billableOpen, K_AX.indexOf(PRICING.axis)).map(x => x.rev)])) : COMPANY_DATA.scenarios.revMW, // v0.1b：每 MW 年收入隨情境；MAG v0.1b：由 pricing 推得（容量軸改變世代組合）
  SC_BR = COMPANY_DATA.scenarios.billableRatio.ratio,
  SC_BRM = COMPANY_DATA.scenarios.billableRatio.mode || `ratio`, // v0.1b（Oracle）：converge＝期初可計費 MW 以最新季實際營收年化 ÷ 每 MW 年收入校準，之後向已連網 MW 收斂
  SC_BOPEN = k => SC_BRM === `converge` ? Math.round(COMPANY_DATA.scenarios.billableRatio.openAnnualRevenue / SC_REV[k][0]) : COMPANY_DATA.defaults.billableOpen,
  UL = COMPANY_DATA.leases.uncommenced, UL_TOT = COMPANY_DATA.leases.facts.notCommenced, // v0.1b（Oracle）：未起租租賃起租排程
  ulPath = (T = UL.termYears, N = UL.quarters, S = UL.startQ, tot = UL_TOT) => { // 每季起租 tot/N（未折現），每筆期限 T 年直線付租；各期租金＝每筆季租 ×(期末累計已起租筆季數 − 期初累計)
    const q = tot / N / T / 4, F = x => { x = Math.max(0, x - S); const m = Math.min(x, N); return m * (m + 1) / 2 + N * Math.max(0, x - N) };
    let c = 0;
    return PERIOD_YEARS.map(L => { const a = c; c += Math.round(L * 4); return q * (F(c) - F(a)) })
  },
  ulTail = p => Math.max(0, UL_TOT - p.reduce((e, t) => e + t, 0)), // 模型期後尚未支付的未起租租金（未折現）＝總額 − 五期路徑
  SC_MW31 = COMPANY_DATA.scenarios.mw31,
  scA = e => {
    let T = COMPANY_DATA.scenarios.capexTemplate;
    return {
      costMW: COMPANY_DATA.capexModel?.mode === `tk` ? capexCostQ(COMPANY_DATA.capexModel.selfBuild).map(x => x.all) : [...T.costMW], // MAG v0.1b：tk＝Tokenomics 每 MW 成本（顯示用；引擎於 runFunding 依自建比例重算）
      customerFund: PERIOD_YEARS.map(() => COMPANY_DATA.defaults.prepay.shareOfDeals * COMPANY_DATA.defaults.prepay.capexCover), // v0.1b：預付比率＝有預付的合約比例 × 預付占相關資本支出比
      newLease: ulPath(), // v0.1b（Oracle）：合約性，三情境相同
      div: [...T.div]
    }
  },
  SCENARIOS = Object.fromEntries([`low`, `base`, `high`].map(k => [k, {
    label: COMPANY_DATA.scenarios.labels[k],
    acc: SC_ACC[k],
    bil: SC_ACC[k].map((e, t) => SC_BRM === `converge` ? Math.round(SC_BOPEN(k) + (e - SC_BOPEN(k)) * SC_BR[t]) : Math.round(e * SC_BR[t])), // v0.1b：converge＝校準起點＋(已連網 − 起點)× 收斂比例
    bOpen: SC_BOPEN(k), // v0.1b：期初可計費 MW（隨情境）
    rev: SC_REV[k],
    mw31: SC_MW31[k],
    cvCap: COMPANY_DATA.scenarios.convCap[k], // v0.1b：瀑布可轉債每年新發行上限（保守 0＝不新發）
    delay: (COMPANY_DATA.scenarios.delayMonths || {})[k] ?? 0, // v0.2：建設延誤月數（計費 MW 平移；GPU 資本支出照原時程）
    a: scA(k)
  }])),
  DEBT_TOOLS = COMPANY_DATA.debt.instruments,
  DBT_P = DEBT_TOOLS.reduce((e, t) => e + t[4], 0),
  DBT_I = DEBT_TOOLS.reduce((e, t) => e + t[3] * t[4], 0),
  DBT_R = DBT_I / DBT_P,
  CONV_P = COMPANY_DATA.debt.convertible.principal,
  CONV_I = CONV_P * COMPANY_DATA.debt.convertible.coupon,
  // v0.1b：可轉債逐檔（company.json → debt.convertibles：[名稱, 原始本金, 票息, 到期 YYYY-MM, 到期累積倍數, 轉換價, 備註]）
  // M＝到期本金（原始 × 累積）、S＝若轉換股數（原始 ÷ 轉換價）、x＝有效轉換價（轉換價 × 累積）、t＝到期所屬模型期（0–4；5＝模型期後）
  CVN = COMPANY_DATA.debt.convertibles.map(c => ({ name: c[0], P: c[1], c: c[2], mat: c[3], acc: c[4], k: c[5], M: c[1] * c[4], S: c[1] / c[5], x: c[5] * c[4], mand: c[7] === true, // v0.1b（Oracle）：第 8 格＝強制轉換（一律轉股）
    t: (i => i < 0 ? 5 : i)(CALQ.periodEnd.findIndex(d => d.slice(0, 7) >= c[3])) })),
  cvConvQ = px => CVN.map(n => n.mand || n.x < px), // 價內（有效轉換價 < 判斷價）或強制轉換＝若轉換法
  cvFlowQ = (cv, inc, r, L) => CVN.reduce((a, n, i) => cv[i] ? a : { // 債務處理者的還本、期末餘額與票息
    amort: a.amort + (inc && n.t === r ? n.M : 0),
    end: a.end + (inc ? (n.t > r ? n.M : 0) : n.M),
    int: a.int + n.P * n.c * L * (inc ? (n.t > r ? 1 : n.t === r ? .5 : 0) : 1)
  }, { amort: 0, end: 0, int: 0 }),
  CHECK_TH = COMPANY_DATA.methodology.checks, // 5a：連動檢查門檻（Excel「檢查_連動」同一來源）
  W36 = PERIOD_YEARS.map((L, i) => Math.max(0, Math.min(L, 3 - PERIOD_YEARS.slice(0, i).reduce((a, b) => a + b, 0))) / L), // v0.1b：評價日起 36 個月落在各期的比例（RPO 對照）
  TXQ = COMPANY_DATA.texts, // v0.1b（Oracle）：公司特有說明文字
  REV_GUIDE_TXT = CALL_FACTS.revLo == null ? `未給全年營收指引` : CALL_FACTS.revHi == null ? `≥${CALL_FACTS.revLo}` : `${CALL_FACTS.revLo}–${CALL_FACTS.revHi}`, // v0.1b：營收指引只有下限時寫「≥」；MAG v0.1b：未給時寫「未給」
  inRevGuideQ = (x, t = 0) => CALL_FACTS.revLo == null || x >= CALL_FACTS.revLo - t && (CALL_FACTS.revHi == null || x <= CALL_FACTS.revHi + t), // MAG v0.1b：未給指引時不判斷
  CX_OLD = COMPANY_DATA.legacy.capexV14,
  INT_OLD = COMPANY_DATA.legacy.interestV14,
  LEASE_FACTS = COMPANY_DATA.leases.facts,
  DEFAULTS = Object.fromEntries(Object.entries(structuredClone(COMPANY_DATA.defaults)).flatMap(([k, v]) => k === `intCal` ? [[k, v], [`a`, structuredClone(SCENARIOS.base.a)]] : [[k, v]])),
  Qk = DEFAULTS; // 模板函式庫片段（mid1–mid3）仍以 Qk 引用預設值，保留別名
DEFAULTS.cvCap = SCENARIOS[COMPANY_DATA.defaults.scenario].cvCap; // v0.1b：預設情境的可轉債年上限
if (SC_MWP.connectedStart == null) DEFAULTS.m.accepted = [...SCENARIOS[COMPANY_DATA.defaults.scenario].acc], DEFAULTS.m.billable = [...SCENARIOS[COMPANY_DATA.defaults.scenario].bil]; // MAG v0.1b：首期自評價日起算——預設 MW 路徑隨日曆（與 calendar_q.apply 同一規則）
if (PRICING) DEFAULTS.priceLink = !0, DEFAULTS.kAxis = K_AX.indexOf(PRICING.axis), DEFAULTS.customFactor = PRICING.customFactor, DEFAULTS.m.revMW = [...SC_REV[COMPANY_DATA.defaults.scenario]]; // MAG v0.1b：價格軸（0 低、1 基準、2 高）與自研晶片係數；priceLink 時每 MW 年收入由 MW 路徑與 k 推得
DEFAULTS.delayMonths = SCENARIOS[COMPANY_DATA.defaults.scenario].delay; // v0.2：預設情境的建設延誤月數
DEFAULTS.delayLink = UL.delayLink ?? 0; // v0.2：未起租租約起租隨延誤後移的比例（company.json → leases.uncommenced.delayLink）
DEFAULTS.ulTerm = UL.termYears; // v0.2：未起租租約租期（延誤平移與租賃負債用；租期敏感度同時改 a.newLease 與此值）
// v0.2：租賃負債（期末剩餘租金現值；company.json → leases.liability）。LLQ＝租賃折現率、LLN＝在帳到期表模型期後尾端年數
var LLQ = (COMPANY_DATA.leases.liability || {}).discRate ?? 0, LLN = (COMPANY_DATA.leases.liability || {}).tailYears ?? 12;
// 在帳租約：期末之後各期到期表現金（期中付款）＋模型期後尾端（LEASE_AFTER_FY30 平均分 LLN 年、年中付款）的現值
function llOnQ(r) {
  const D = x => (1 + LLQ) ** -x, T = PERIOD_T, tail = LLQ > 0 ? LEASE_AFTER_FY30 / LLN * (1 + LLQ) ** .5 * (1 - (1 + LLQ) ** -LLN) / LLQ : LEASE_AFTER_FY30;
  return LEASE_CASH_ON_BAL.reduce((a, c, k) => a + (k > r ? c * D(T[k] - PERIOD_YEARS[k] / 2 - T[r]) : 0), 0) + tail * D(T[4] - T[r])
}
// 未起租租約已起租部分：每季起租 tot/N、每筆 4T 季、季末付 q＝tot/N/T/4；期末累計季數 x、起算季 S、權重 w：
// ＝w × q ÷ i_q × [m − v^(4T − x') ×(1 − v^m) ÷ (1 − v)]，x'＝MAX(0, x − S)、m＝MIN(x', N)、v＝(1＋折現率)^(−1/4)、i_q＝(1＋折現率)^(1/4) − 1（Excel 同式）
function llUlQ(x, S, w, T) {
  const xx = Math.max(0, x - S), m = Math.min(xx, UL.quarters), q = UL_TOT / UL.quarters / T / 4, R = 1 + LLQ;
  if (!(LLQ > 0)) return w * q * (m * 4 * T - (m * (2 * xx - m + 1)) / 2);
  return w * q / (R ** .25 - 1) * (m - R ** (-(4 * T - xx) / 4) * (1 - R ** (-m / 4)) / (1 - R ** -.25))
}
// v0.2：期末存量路徑往後平移 dm 個月（以期間長度線性內插；評價日之前取 v0）。V＝各期末值、v0＝評價日值。
// ＝v0＋Σ_k (V_k − V_{k−1}) × MIN(1, MAX(0, (期末時點_i − dm/12 − 期初時點_k) ÷ 期間長度_k))；Excel 同一公式。dm＝0 時原樣回傳。
var PERIOD_T = PERIOD_YEARS.reduce((a, L, i) => (a.push((i ? a[i - 1] : 0) + L), a), []);
function shiftQ(V, v0, dm) {
  if (!(dm > 0)) return [...V];
  return V.map((x, i) => { const tau = PERIOD_T[i] - dm / 12; return V.reduce((a, v, k) => a + (v - (k ? V[k - 1] : v0)) * Math.min(1, Math.max(0, (tau - (k ? PERIOD_T[k - 1] : 0)) / PERIOD_YEARS[k])), v0) })
}
// v0.1b（Oracle）：非 AI 事業（company.json → defaults.legacyBiz）；MAG v0.1b：改為 N 線（kind），各線自有 EBITDA 率與其他攤銷。
// growth：全年營收＝上一財年實際 ×(1＋年增率)，年增率自 g0 線性收斂到 gLT；首期模型部分＝首期全年 − 年初至今實際。
// cloudResidual：雲端分部的非 AI 部分＝雲端分部首期估計〔年初至今＋去年同期 ×(1＋近四季年增率)〕− 對外 AI 雲端模型值（年初至今以期初 MW 平均 × 首期每 MW 年收入估計）；
//   g0＝(雲端近四季 − AI 近四季估計) ÷ (前四季 − AI 前四季估計) − 1；m0＝(雲端近四季 EBITDA − AI 近四季估計 × AI 首期 EBITDA 率) ÷ (雲端近四季 − AI 近四季估計)。X＝{ aiRev（各期 AI 模型期營收）, rev0（首期每 MW 年收入）, m0（AI 首期 EBITDA 率）}。
// explicit：各期模型期營收直接輸入（rev；首期為模型部分），EBITDA 率同 m0→mLT 或 ebitda 陣列直接輸入（未分攤公司層費用：rev 全 0、ebitda 為負）。
// EBITDA 率 m＝m0＋(mLT − m0)× 期數 ÷ 4（mLT 空白＝m0）；oa＝其他攤銷占營收比（進損益 D&A、自營運現金扣除）。沒有非 AI 事業時 lines 為空清單，全部為 0。Excel「輸入與假設」B2 同式。
var B_LINE = k => (COMPANY_DATA.defaults.legacyBiz?.lines || []).find(x => x.key === k) || {}; // MAG v0.1b：各線原始輸入（檢查文字用）
function legacyQ(e, X) {
  let B = e.legacyBiz || { lines: [] }, Y0 = CALQ.ytdMonths / 12,
    lines = B.lines.map(x => {
      const k = x.kind || `growth`;
      let g0 = x.g0, m0 = x.m0, A = [], ytd = x.ytd || 0, ai = null;
      if (k === `cloudResidual`) {
        const aiY = X.rev0 * ((e.mwYearEnd[PERIOD_FY[0] - 1] ?? 0) + e.billableOpen) / 2 * Y0, aiT = X.rev0 * x.aiMwTTM, aiP = X.rev0 * x.aiMwPrevTTM;
        g0 = (x.ttm - aiT) / (x.prevTTM - aiP) - 1, m0 = (x.ebitdaTTM - aiT * X.m0) / (x.ttm - aiT), ytd = x.ytd - aiY;
        ai = { ytd: aiY, ttm: aiT, prevTTM: aiP, segFy0: x.ytd + x.priorStub * (1 + x.g4q) }
      }
      const mLT = x.mLT ?? m0, g = PERIOD_YEARS.map((L, r) => k === `explicit` ? 0 : g0 + (x.gLT - g0) * r / 4), m = PERIOD_YEARS.map((L, r) => m0 + (mLT - m0) * r / 4);
      if (k === `explicit`) A = x.rev.map((v, r) => r === 0 ? v + ytd : v);
      else g.forEach((gr, r) => A.push(r === 0 ? (k === `cloudResidual` ? ai.segFy0 - ai.ytd - X.aiRev[0] : x.fyBase * (1 + gr)) : A[r - 1] * (1 + gr)));
      const rev = A.map((a, r) => r === 0 ? a - ytd : a), eb = rev.map((v, r) => x.ebitda ? x.ebitda[r] : v * m[r]);
      return { key: x.key, label: x.label, kind: k, peer: x.peer, g, m, g0, m0, annual: A, rev, ytd, ebitda: eb, oa: rev.map(v => v * (x.oa || 0)), cx: x.cx || 0, ai }
    }),
    sumQ = f => PERIOD_YEARS.map((L, r) => lines.reduce((a, x) => a + x[f][r], 0)),
    rev = sumQ(`rev`), ebitda = sumQ(`ebitda`);
  return { lines, rev, ebitda, oa: sumQ(`oa`), margin: rev.map((v, r) => ebitda[r] / Math.max(v, 1e-9)), annual: sumQ(`annual`) }
}

// MAG v0.1b：關聯方並排（related.contracts：年化 × 是否已起始；snapshots：依期間標籤取值）；ai＝對外 AI 雲端收入年化（模型期 ÷ 期間長度）
function relatedQ(o) {
  const RL = COMPANY_DATA.related, ai = o.map((y, r) => y.isRev / PERIOD_YEARS[r]),
    rows = RL.contracts.map(c => PERIOD_FY.map(fy => fy >= c.startFY ? c.annual : 0)),
    tot = PERIOD_YEARS.map((L, r) => rows.reduce((a, x) => a + x[r], 0));
  return { ai, rows, tot, ratio: tot.map((x, r) => x / Math.max(ai[r], 1e-9)), snap: RL.snapshots.map(c => PERIODS.map(p => c.series[p] ?? null)) }
}
// MAG v0.1b：租用算力排程（leases.rentedCompute：[{name, start＝YYYY-MM 起租, years, annualRent（US$bn／年）, mw, use}]）：各期租金＝年租金 × 期間內在租月數 ÷ 12（Excel 同式，月數由建置時推算）
var LEASE_OIE = !!COMPANY_DATA.leases.operatingInEbitda,
  monQ = s => { const [y, m] = s.split(`-`).map(Number); return y * 12 + m - 1 },
  RENT_OV = (c => PERIOD_YEARS.map((L, r) => { const v = monQ(CALQ.valuationDate.slice(0, 7)) + 1, a = v + Math.round(CALQ.tStart[r] * 12), b = v + Math.round(CALQ.tEnd[r] * 12), s0 = monQ(c.start), s1 = s0 + Math.round(c.years * 12); return Math.max(0, Math.min(b, s1) - Math.max(a, s0)) / 12 })),
  rentedQ = L => PERIOD_YEARS.map((x, r) => L.reduce((a, c) => a + c.annualRent * RENT_OV(c)[r], 0));
// MAG v0.1b：每 MW 資本支出（US$m/MW）＝Σ 新增占比 × (TK_CapexIT＋自建比例 × TK_CapexFacility)；IT 部分＝Σ 新增占比 × TK_CapexIT（Excel「輸入與假設」D 區同式）
function capexCostQ(sb) {
  return PERIOD_YEARS.map((L, r) => {
    const it = PRICING.chips.reduce((a, c) => a + c.mixAdds[r] * TKV.IF_CapexIT[c.tk], 0), fc = PRICING.chips.reduce((a, c) => a + c.mixAdds[r] * TKV.IF_CapexFacility[c.tk], 0) * sb;
    return { it, fac: fc, all: it + fc, itShare: it / Math.max(it + fc, 1e-9) }
  })
}
// MAG v0.1b：期初 AI 毛 PP&E 估計＝評價日 AI 總 MW（對外 ÷ 對外比例）× Σ 期初占比 × Tokenomics 每 MW 成本；機房壽命＝TK_CapexFacility ÷ TK_DeprFac；
// 非 AI 期初基礎＝期初 PP&E − AI 估計，壽命＝非 AI 基礎 ÷ (最新季分部 D&A 年化 − AI 期初 D&A)（校準）
function aiOpenQ(e, XS) {
  const CM = COMPANY_DATA.capexModel, sb = e.selfBuild ?? CM.selfBuild, mw = Math.max(0, e.billableOpen / XS - ((CM.rentedExt || {}).open || 0)), // MAG v0.1b r3（C20）：自有 AI 總 MW＝對外 ÷ 對外比例 − 租用對外 MW
    W = f => PRICING.chips.reduce((a, c) => a + c.mixOpen * TKV[f][c.tk], 0),
    it = mw * W(`IF_CapexIT`) / 1e3, fac = mw * sb * W(`IF_CapexFacility`) / 1e3, facLife = W(`IF_CapexFacility`) / W(`IF_DeprFac`),
    da0 = it / e.gpuLife + fac / facLife, n0 = e.ppeOpen - it - fac;
  const nLife = CM.nonAiLife ?? 10; // MAG v0.1b r3（對照表 r1 C17）：非 AI 折舊年限＝預設年限（不再以分部 D&A 反解，避免 AI 期初 D&A > 合併 D&A 時退化）
  return { mw, it, fac, facLife, da0, n0, nLife, daRecon: CM.segDaRunRate - da0 - n0 / nLife } // daRecon＝最新季分部 D&A 年化 −（AI 期初 D&A＋非 AI 期初 ÷ 年限）：對帳殘差（示警，不回填）
}

function siteBenchQ(e) {
  let t = e.filter(e => !e.residual && e.contract && e.years && e.planned);
  return t.reduce((e, t) => e + t.contract / t.years, 0) / Math.max(t.reduce((e, t) => e + t.planned, 0), 1) * 1e3
}

function $k(e) {
  return e.filter(e => !e.residual)
}

function eA(e) {
  let t = $k(e),
    n = e => t.reduce((t, n) => t + (Number(n[e]) || 0), 0);
  return {
    planned: n(`planned`),
    energized: n(`energized`),
    accepted: n(`accepted`),
    billable: n(`billable`)
  }
}

function tA(e) {
  let t = $k(e.sites).map(e => ({
      ...e,
      residual: !1
    })),
    n = eA(t),
    r = e.m.accepted[0],
    i = Math.min(e.m.billable[0], r),
    a = e.m.accepted[4];
  return [...t, {
    id: `other`,
    name: `其他／未列名站點`,
    operator: `模型殘差（連動）`,
    planned: Math.max(0, a - n.planned),
    energized: Math.max(0, r - n.energized),
    accepted: Math.max(0, r - n.accepted),
    billable: Math.max(0, i - n.billable),
    status: `用來讓站點加總 = 年度 MW。公司不揭露逐站 MW，殘差必然偏大`,
    next: `有具名專案時改列名列`,
    date: `連動`,
    confidence: `模型`,
    residual: !0
  }]
}

function nA(e) {
  let t = {
      accepted: [...e.m.accepted],
      billable: [...e.m.billable],
      util: [...e.m.util],
      revMW: [...e.m.revMW],
      aiShare: [...e.m.aiShare],
      fill: [...(e.m.fill || [100, 100, 100, 100, 100])],
      power: [...e.m.power],
      pue: [...e.m.pue],
      maint: [...e.m.maint],
      defaultP: [...e.m.defaultP],
      recovery: [...e.m.recovery],
      rate: [...e.m.rate]
    },
    n = eA(e.sites);
  e.revenueDriver === `mw` && (t.fill = t.fill.map(() => 100)); // v0.1b：MW 驅動——營收＝容量上限（平均在役 MW × 每 MW 年收入 × 利用率），RPO 只作對照
  e.linkSites && (t.accepted[0] = Math.max(t.accepted[0], n.accepted), t.billable[0] = Math.max(t.billable[0], n.billable));
  for (let e = 0; e < 5; e++) t.accepted[e] = Math.max(0, t.accepted[e]), e > 0 && (t.accepted[e] = Math.max(t.accepted[e], t.accepted[e - 1])), t.billable[e] = Math.min(Math.max(0, t.billable[e]), t.accepted[e]), t.util[e] = Math.min(100, Math.max(0, t.util[e])), t.aiShare[e] = Math.min(100, Math.max(0, t.aiShare[e])), t.fill[e] = Math.min(100, Math.max(0, t.fill[e])), t.defaultP[e] = Math.min(100, Math.max(0, t.defaultP[e])), t.recovery[e] = Math.min(100, Math.max(0, t.recovery[e]));
  if (PRICING && e.priceLink) { // MAG v0.1b：每 MW 年收入與營運成本由（單調化後的）在役 MW 路徑 × 世代組合 × k 推得（Excel「輸入與假設」B0 同式）
    const R = revPathQ(t.accepted, e.billableOpen, e.kAxis, e.customFactor);
    t.revMW = R.map(x => x.rev), t.opexMW = R.map(x => x.opexMW), t.price = R
  }
  return t
}

function rA(e) {
  return LEASE_CASH_ON_BAL[4] * e.terminal.residualLeaseYears + ulTail(e.a.newLease) // v0.1b：表外尾端＝未起租總額 − 五期路徑（取代末期 × 年數）
}

function iA(e) {
  return e.a.newLease.reduce((e, t) => e + t, 0) + ulTail(e.a.newLease)
}

function aA(e, t) {
  let n = t.totals.operatingGap,
    r = t.totals.atm,
    f = t.totals.facility,
    i = t.years.reduce((e, t) => e + t.debtPay, 0),
    a = e.cash + n + r + f - i;
  return {
    open: e.cash,
    operatingGap: n,
    atm: r,
    facility: f,
    debtPay: i,
    implied: a,
    end: t.totals.end,
    ok: Math.abs(a - t.totals.end) < .05
  }
}

function rpoBridge(e) {
  let t = e.years.reduce((e, t) => e + t.scheduled, 0),
    n = e.years.reduce((e, t) => e + t.bottleneck, 0);
  return {
    scheduled: t,
    bottleneck: n,
    afterCap: t - n,
    credit: e.years.reduce((e, t) => e + t.loss, 0),
    collected: e.years.reduce((e, t) => e + t.collected, 0)
  }
}

function sA(e, t) {
  let n = e.rpoOpen,
    r = e.rpoPendingAdd,
    i = n + r;
  return {
    begin: n,
    q1Net: r,
    q1End: i,
    convert: t.years.reduce((e, t) => e + t.scheduled, 0),
    remaining: i * (1 - e.rp / 100)
  }
}

function cA(e) {
  let t = e.a.newLease.reduce((e, t) => e + t, 0),
    n = ulTail(e.a.newLease),
    r = t + n,
    i = LEASE_CASH_ON_BAL.reduce((e, t) => e + t, 0),
    a = LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap;
  return {
    uncommenced: LATEST_Q.offBalanceLease,
    singleSite: LATEST_Q.singleSiteCap,
    committed: a,
    cashFive: t,
    tail: n,
    mapped: r,
    onFive: i,
    onAfter: LEASE_AFTER_FY30,
    gap: r - a
  }
}

function lA(e) {
  return e >= 800 ? {
    label: `危險`,
    tone: `bad`
  } : e >= 600 ? {
    label: `壓力`,
    tone: `stress`
  } : e > 450 ? {
    label: `偏高／觀察`,
    tone: `watch`
  } : {
    label: `改善`,
    tone: `ok`
  }
}

function uA(e, t) {
  let n = e.m.rate[t];
  if (!e.cdsLink) return n;
  let r = (e.cds - e.cdsBaseBp) / 100 * e.cdsPassThrough; // 5a：CDS 門檻與傳導比例讀 company.json → defaults
  return Math.max(0, n + r)
}

function runFunding(e) {
  let t = nA(e),
    n = tA({
      ...e,
      m: t
    }),
    r = eA(e.sites),
    i = e.rp / 100,
    a = e.cash,
    MB = [e.mwYearEnd[PERIOD_FY[0] - 1], ...t.accepted.slice(0, 4)], // 5a：期初＝首期前一財年末主動電力（defaults.mwYearEnd 以年份為鍵）
    MN = t.accepted.map((e, n) => e - MB[n]),
    t_acc = n => t.accepted[n],
    MX = t.accepted.map((n, r) => r < 4 ? t.accepted[r + 1] - n : e.mw31),
    DM = e.delayMonths ?? 0, // v0.2：建設延誤月數
    BD = shiftQ(t.billable, e.billableOpen, DM), // v0.2：計費用可計費 MW（原路徑平移延誤月數）
    AIR = PERIOD_YEARS.map((L, r) => (e.useAvgMw ? ((r === 0 ? e.billableOpen : BD[r - 1]) + BD[r]) / 2 : BD[r]) * t.revMW[r] * (e.revScale ?? 1) * (t.util[r] / 100) * L), // MAG v0.1b：對外 AI 雲端模型期營收（＝下方容量 u；非 AI 雲端殘差用）
    AIM = PERIOD_YEARS.map((L, r) => e.ebitdaBasis === `tk` ? 1 - t.opexMW[r] / Math.max(t.revMW[r] * (e.revScale ?? 1), 1e-9) : null), // MAG v0.1b：tk＝AI 雲端 EBITDA 率＝1 − Tokenomics 加權每 MW 營運成本 ÷ 每 MW 年收入（價格倍數改變收入、不改成本）
    LG = legacyQ(e, { aiRev: AIR, rev0: t.revMW[0] * (e.revScale ?? 1) * (t.util[0] / 100), m0: AIM[0] ?? e.ebStart }), // v0.1b（Oracle）：非 AI 事業營收與 EBITDA；MAG v0.1b：N 線
    RENTC = rentedQ(e.rentedCompute ?? COMPANY_DATA.leases.rentedCompute ?? []), // MAG v0.1b：租用算力租金（各期；C8 d）
    CXM = COMPANY_DATA.capexModel && COMPANY_DATA.capexModel.mode === `tk` ? COMPANY_DATA.capexModel : null, // MAG v0.1b：資本支出由 Tokenomics 每 MW 成本推導（AI 與非 AI 分池）
    XS = CXM ? e.extShare ?? CXM.extShare : 1, // 對外比例：新增 AI 總 MW＝新增對外 MW ÷ 對外比例（自用 AI 同樣需要資本支出）
    CC = CXM ? capexCostQ(e.selfBuild ?? CXM.selfBuild) : null, // 各期新增 MW 的每 MW 成本（US$m/MW：IT＋自建比例 × 機房）與 IT 占比
    CMW = CXM ? CC.map(x => x.all) : e.a.costMW, CIT = CXM ? CC.map(x => x.it) : e.a.costMW, // 汰換只換 IT（機房不換）；非 tk 模式沿用 e.a.costMW
    RXP = CXM ? rentedExtQ() : null, RXD = n => RXP ? (n < 5 ? RXP.path[n] - (n ? RXP.path[n - 1] : RXP.open) : 0) : 0, // MAG v0.1b r3（C20）：租用對外 MW 的新增（不需資本支出）
    CXF = MN.map((t, n) => (CXM ? Math.max(0, (t * (1 - e.lambda) + MX[n] * e.lambda) / XS - (RXD(n) * (1 - e.lambda) + RXD(n + 1) * e.lambda)) : t * (1 - e.lambda) + MX[n] * e.lambda) * CMW[n] * (e.capexScale ?? 1) / 1e3),
    CXNF = PERIOD_YEARS.map((L, r) => CXM ? LG.lines.reduce((a, x) => a + x.annual[r] * (x.cx || 0), 0) : 0), // MAG v0.1b：非 AI 資本支出（全年）＝Σ 各線全年營收 × 資本支出強度
    VIN = Object.fromEntries(Object.entries(e.mwYearEnd).map(([y, m]) => [y, m - (e.mwYearEnd[y - 1] ?? 0)])), // 各年新增 MW（汰換批次）
    RFV = PERIOD_FY.map( // 5a：期間的財年年份由日曆推算（滾動後不寫死）
      (t, n) => (VIN[t - e.gpuLife] || 0) / XS * CIT[n] * (e.capexScale ?? 1) / 1e3),
    // v0.1c（Oracle）：穩態汰換——已連網 MW 不再增加的期間（觸頂後）及終值年，汰換 CapEx＝期間平均已連網 MW × 每 MW GPU 資本支出 ÷ GPU 經濟壽命 × 期間長度（建築與電力屬租賃不計）；取代批次汰換
    RFF = PERIOD_FY.map((t, n) => !!e.refreshSteady && (n === 4 || MN[n] <= 0)),
    RFS = PERIOD_FY.map((t, n) => (MB[n] + t_acc(n)) / 2 / XS * CIT[n] * (e.capexScale ?? 1) / 1e3 / e.gpuLife * PERIOD_YEARS[n]),
    REF = RFV.map((x, n) => RFF[n] ? RFS[n] : x),
    // 成長型（模型期）：首期＝MAX(全年 AI 成長型＋非 AI, 指引下限) − 年初至今實際；非 AI 首期模型部分＝全年 × 模型期月數 ÷ 12，其餘歸 AI（指引高於公式的差額視為 AI，保守）
    CXN = CXNF.map((x, n) => n === 0 ? x * CALQ.stubMonths / 12 : x),
    CXG = CXF.map((t, n) => n === 0 ? Math.max(t + CXNF[0], e.capexFloorFY0 ?? 0) - ACTUAL_1H.capex : t + CXN[n]),
    CXA = CXG.map((x, n) => x - CXN[n]), // AI 成長型（模型期）
    CX = CXG.map((e, t) => e + REF[t]),
    CXC = CXG.reduce((a, x, i) => (a.push((i ? a[i - 1] : 0) + x), a), []), // v0.2：成長型 CapEx 累計（原時程）
    CXSC = shiftQ(CXC, 0, DM), // v0.2：已投入使用的成長型 CapEx 累計（延誤後）
    CXS = DM > 0 ? CXSC.map((x, i) => x - (i ? CXSC[i - 1] : 0)) : CXG, // v0.2：本期投入使用的成長型 CapEx（折舊基礎）
    IDLE = CXC.map((x, i) => DM > 0 ? x - CXSC[i] : 0), // v0.2：閒置資本＝已支出而尚未產生收入的累計成長型 CapEx（期末）
    LK = e.delayLink ?? 0, // v0.2：未起租租約起租連動比例
    ULT = e.ulTerm ?? UL.termYears,
    ULS = DM > 0 && LK > 0 ? ulPath(ULT, UL.quarters, UL.startQ + DM / 3) : null, // v0.2：未起租租金全部隨延誤後移（起算季＋延誤月數 ÷ 3）
    ULX = PERIOD_YEARS.reduce((a, L, i) => (a.push((i ? a[i - 1] : 0) + Math.round(L * 4)), a), []), // 期末累計季數
    LLON = PERIOD_YEARS.map((L, r) => llOnQ(r)), // v0.2：在帳租賃負債（期末）
    LLUL = ULX.map(x => ULS ? llUlQ(x, UL.startQ, 1 - LK, ULT) + llUlQ(x, UL.startQ + DM / 3, LK, ULT) : llUlQ(x, UL.startQ, 1, ULT)), // v0.2：未起租租約已起租部分的租賃負債（期末）
    PPE = [],
    // MAG v0.1b：D&A 分池——AI（IT 依 GPU 壽命、機房依機房壽命）＋非 AI（期初＝PP&E − AI 期初毛額估計，壽命以最新季分部 D&A 年化校準）；汰換不增加基礎（模板慣例）
    AP0 = CXM ? aiOpenQ(e, XS) : null,
    AI_IT = [], AI_FC = [], NA_P = [],
    DAI = CXS.map((x, n) => { if (!CXM) return 0;
      const a = DM > 0 ? x - CXN[n] : CXA[n], it = a * CC[n].itShare, fc = a - it,
        b1 = n === 0 ? AP0.it : AI_IT[n - 1][0] + AI_IT[n - 1][1], b2 = n === 0 ? AP0.fac : AI_FC[n - 1][0] + AI_FC[n - 1][1];
      AI_IT.push([b1, it]), AI_FC.push([b2, fc]);
      return ((b1 + .5 * it) / e.gpuLife + (b2 + .5 * fc) / AP0.facLife) * PERIOD_YEARS[n] }),
    DAN = CXN.map((x, n) => { if (!CXM) return 0;
      const b = n === 0 ? AP0.n0 : NA_P[n - 1][0] + NA_P[n - 1][1]; NA_P.push([b, x]);
      return (b + .5 * x) / AP0.nLife * PERIOD_YEARS[n] }),
    DAF = CXS.map((t, n) => {
      let r = n === 0 ? e.ppeOpen : PPE[n - 1] + CXS[n - 1];
      return PPE.push(r), CXM ? DAI[n] + DAN[n] : (r + .5 * t) / e.gpuLife * PERIOD_YEARS[n]
    }),
    CVF = cvConvQ(e.eqPx), // v0.1b：融資現金流的可轉債分類（判斷價＝股權發行參考價，預設＝現價）
    CVP = PERIOD_YEARS.map((L, n) => cvFlowQ(CVF, e.includeDebt, n, L)),
    PB = [],
    IX = DEBT_AMORT.map((t, n) => {
      let r = n === 0 ? DBT_P : PB[n - 1][1],
        i = r - t;
      return PB.push([r, i]), (r + i) / 2 * DBT_R * PERIOD_YEARS[n] + CONV_I * PERIOD_YEARS[n] + CVP[n].int + (n === 0 ? e.intCal : 0)
    }),
    WF = {
      Jn: 0,
      pc: e.cash,
      Dn: 0,
      fr: e.useFacility ? e.facility : 0,
      B: e.rpoOpen + e.rpoPendingAdd,
      pnr: 0,
      sh: 0,
      cl: e.prepay.openBalance, // v0.1b：合約負債（客戶預付餘額）期初
      Cn: 0 // v0.1b：瀑布新發可轉債餘額
    },
    o = PERIODS.map((n, r) => {
      let L = PERIOD_YEARS[r],
        o = e.rpoOpen * (RPO_BUCKET_W[r] / RPO_SCHEDULED_SHARE) * i + e.rpoPendingAdd * RPO_Q3ADD_W[r] * i,
        s = o,
        c = r === 0 ? e.billableOpen : BD[r - 1],
        l = e.useAvgMw ? (c + BD[r]) / 2 : BD[r], // v0.2：計費用 MW＝延誤後路徑
        u = l * t.revMW[r] * (e.revScale ?? 1) * (t.util[r] / 100) * L,
        d = Math.max(0, s - u),
        f = o - d,
        p = 1 - t.recovery[r] / 100,
        m = f * (t.defaultP[r] / 100) * p,
        h = f - m,
        nR = Math.max(0, u - s) * (t.fill[r] / 100),
        b0 = LEASE_CASH_ON_BAL[r],
        x0 = ULS ? (1 - LK) * e.a.newLease[r] + LK * ULS[r] : e.a.newLease[r], // v0.2：延誤連動部分的起租往後平移
        // MAG v0.1b：operatingInEbitda＝分部 EBITDA 已扣營業租賃成本——在帳只扣融資租賃（含融資義務）現金、未起租只扣融資部分（1 − opShare）；營業部分視為已含在 EBITDA 率
        b = LEASE_OIE ? COMPANY_DATA.leases.financePayments[r] : b0,
        x = LEASE_OIE ? x0 * (1 - (UL.opShare ?? 0)) : x0,
        c0 = r === 0 ? e.billableOpen : t.billable[r - 1], u0 = (e.useAvgMw ? (c0 + t.billable[r]) / 2 : t.billable[r]) * t.revMW[r] * (e.revScale ?? 1) * (t.util[r] / 100) * L, // v0.2：未延誤的容量上限（對照）
        lost = DM > 0 ? Math.max(0, u0 - u) : 0, // v0.2：應計費而未計費營收（延誤造成）
        pen = (e.delayPenalty ?? 0) * lost, // v0.2：延誤罰則／服務抵減（營業費用：扣 EBITDA、營運來源、稅基、債務上限）
        S = b + x,
        svc = e.services[r],
        totRev = f + nR + svc,
        // v0.1c（Oracle）：ebitdaBasis＝ebitdar 時，EBITDA＝EBITDAR 率 × 營收 − 租金（租金為固定成本）；EBITDAR 率＝EBITDA 率＋基準情境租金÷AI 雲端 營收（defaults.ebitdarAdj，校準於基準情境起點與穩態）
        eR = e.ebitdaBasis === `ebitdar` ? e.ebStart + e.ebitdarAdj[0] + (e.ebSteady + e.ebitdarAdj[1] - e.ebStart - e.ebitdarAdj[0]) * r / 4 : null,
        ebM = eR !== null ? eR - S / Math.max(totRev, .01) : AIM[r] ?? e.ebStart + (e.ebSteady - e.ebStart) * r / 4,
        cm = eR !== null ? eR : e.ebitdaBasis === `tk` ? ebM : ebM + S / Math.max(totRev, .01), // MAG v0.1b：tk＝AI 雲端 EBITDA 率由 Tokenomics 營運成本推得（未扣租金），租金另列現金支出，不加回
        g = h * cm,
        nC = nR * (1 - (t.defaultP[r] / 100) * p) * cm,
        svcCash = svc * cm,
        ob = ((e.otherEbitda || [])[r] || 0) - RENTC[r], // v0.1b：其他事業 EBITDA（負值＝燒錢），同時進入 EBITDA 與營運來源；MAG v0.1b：扣租用算力租金（營運成本，C8 d）
        _ = CX[r],
        v = (CXG[r] + (e.prepay.coverRefresh ? REF[r] : 0)) * e.a.customerFund[r], // v0.1b：客戶預付流入＝成長型 CapEx × 預付比率；v0.1c（Oracle）：prepay.coverRefresh 時汰換 CapEx 同樣適用覆蓋比
        y = _ - v,
        clB = WF.cl,
        ppI = (e.prepay.financingRate ?? 0) * (clB + .5 * v) * L, // v0.1b（Oracle）：重大財務組成——合約負債以隱含利率累積的非現金利息（期初餘額＋本期流入一半）
        pr = Math.min(clB + ppI, (clB + ppI) / e.prepay.recogYears * L), // v0.1b：預付認列（非現金營收）＝(期初合約負債＋累積利息) ÷ 認列年數 × 期間長度
        clE = clB + v + ppI - pr,
        C = t.accepted[r] * 8760 * t.pue[r] * t.power[r] / 1e9 * L,
        w = t.accepted[r] * t.maint[r] / 1e3 * L,
        T = e.overlay ? C + w : 0,
        O = (e.includeDebt ? DEBT_AMORT[r] : 0) + CVP[r].amort, // v0.1b：含債務處理可轉債的到期還本（到期累積本金）
        k = r === 0 && e.includeAtm ? e.atm : 0,
        lgR = LG.rev[r], lgE = LG.ebitda[r], lgO = LG.oa[r], // v0.1b（Oracle）：非 AI 事業營收與 EBITDA（EBITDA 視為現金，稅另列）；MAG v0.1b：lgO＝其他攤銷（損益 D&A；視為等額現金支出，自營運來源扣除）
        tx = (e.cashTaxRate ?? 0) * Math.max(0, totRev * ebM + ob + lgE - lgO - pen - DAF[r] - IX[r]), // v0.1b：現金稅＝稅率 × MAX(0, 損益 EBITDA − 車隊 D&A − 存量利息)（不含瀑布新債利息，避免循環；偏保守）
        dvSh = e.dividend ? e.dividend.sharesBase + WF.sh + CVN.reduce((a, n) => a + (n.mand && n.t < r ? n.S : 0), 0) : 0, // v0.1b（Oracle）：股利股數＝期初股數（基礎＋前期累計瀑布新股＋已強制轉換的特別股）
        dvC = e.dividend ? 4 * e.dividend.perShareQ * Math.pow(1 + (e.dividend.growth ?? 0), r) * L * dvSh : 0, // MAG v0.1b r3（C19）：每股股利年成長 defaults.dividend.growth // 普通股股利＝每股（每季 × 4）× 期間長度 × 期初股數
        dvP = e.dividend ? e.dividend.preferred[r] : 0, // 特別股股利（強制轉換前；company.json → defaults.dividend.preferred）
        sbcC = (e.sbcRate ?? 0) * (totRev + lgR), // MAG v0.1b：股權報酬（非現金）加回營運現金＝SBC 占營收 × 模型期總營收
        wcS = r === 0 ? e.wcStub ?? 0 : 0, // MAG v0.1b r2（C14）：首期剩餘季度營運資金變動＝上一年度同期實際（流入為正；company.json → defaults.wcStub）
        A = g + nC + svcCash + ob + lgE - lgO - pen + v - pr + sbcC + wcS, // v0.1b：預付認列的營收已在預付時收現，自營運來源扣除（不重複計入）
        j0 = _ + S + IX[r] + e.jvCommit[r] + e.a.div[r] + T + O + tx + dvC + dvP,
        wRL = uA(e, r) / 100 * L,
        wRJ = (e.junkRate + (e.cdsLink ? Math.max(0, e.cds - e.cdsBaseBp) / 1e4 * e.cdsPassThrough : 0)) * L,
        wRC = e.convIssue.coupon * L, // v0.1b：瀑布可轉債票息 × 期間長度
        wI0 = wRL * WF.Dn + wRJ * WF.Jn + wRC * WF.Cn,
        wPre0 = a + A + k - j0 - wI0,
        // MAG v0.1b：瀑布第二步「減少回購」——計畫回購＝基準年額 × 期間長度；實際＝MIN(計畫, MAX(下限, 融資前現金 − 最低現金))；不足才進入新債
        bbP = (e.buyback?.annual ?? 0) * L, bbA = Math.min(bbP, Math.max(bbP * (e.buyback?.floorShare ?? 0), wPre0 - e.minCash)), bbSh = bbA / Math.max(e.eqPx, .01),
        wPre = wPre0 - bbA,
        wX = Math.max(0, e.minCash - wPre),
        wB = WF.B - o - nR + e.ctrTerm * Math.max(0, nR / L - WF.pnr),
        wEx = (e.includeDebt ? PB[r][1] : DBT_P) + CONV_P + CVP[r].end, // 5a：評價日後新發可轉債本金讀 company.json → debt.convertible；v0.1b：加債務處理可轉債餘額
        wCapB = e.debtCapBasis === `leaseAdj` ? e.debtEbitdaMax * (totRev * ebM + ob + lgE - pen + S) / L - LLON[r] - LLUL[r] : e.debtCapBasis === `ebitda` ? e.debtEbitdaMax * (totRev * ebM + ob + lgE - pen) / L : e.debtBacklog * wB, // v0.2：leaseAdj＝(總債務＋租賃負債) ≤ 倍數 ×(EBITDA＋租金)（年化） // v0.1b（Oracle）：債務上限＝倍數 × 當期 EBITDA（年化）；模板＝債務／backlog
        wCap = wCapB - (wEx + WF.Dn + WF.Cn),
        wCapD = Math.max(0, wCap, WF.fr),
        wD = Math.min(wCapD, wX / (1 - wRL)),
        wRem = Math.max(0, wX + wRL * wD - wD),
        wCapC = Math.max(0, e.cvCap ?? 0) * L, // v0.1b：可轉債步驟（資產擔保融資之後、ATM 股權之前）
        wC = Math.min(wCapC, wRem / (1 - wRC)),
        wRem2 = Math.max(0, wRem + wRC * wC - wC),
        wCapEq = e.eqCapPct >= 9 ? 1 / 0 : e.eqCapPct * e.eqPx * e.eqCapShares * L,
        wEq = Math.min(wRem2, wCapEq),
        wJ = (wRem2 - wEq) / (1 - wRJ),
        E = wI0 + wRL * wD + wRC * wC + wRJ * wJ,
        wSh = wEq / (e.eqPx * (1 - e.eqDisc)),
        D = IX[r] + E,
        j = j0 + bbA + E,
        F = wD + wC + wEq + wJ,
        M = A + k + F,
        N = M - j,
        ee = A - (j - O),
        wFr0 = WF.fr,
        wDn0 = WF.Dn,
        wCn0 = WF.Cn,
        wPc = WF.pc + A + k - j0;
      return a += N, WF.fr -= Math.min(WF.fr, wD), WF.Dn += wD, WF.Jn += wJ, WF.pc = wPc, WF.B = wB, WF.pnr = nR / L, WF.sh += wSh - bbSh, WF.cl = clE, WF.Cn += wC, {
        convNew: wC,
        convCap: wCapC,
        convBeg: wCn0,
        convEnd: WF.Cn,
        convNeedAfter: wRem2,
        cvAmort: CVP[r].amort,
        cvEnd: CVP[r].end,
        cvInt: CVP[r].int,
        prepayIn: v,
        prepayRecog: pr,
        prepayAccr: ppI,
        clBeg: clB,
        clEnd: clE,
        junk: wJ,
        junkEnd: WF.Jn,
        eqCap: wCapEq,
        preFinCum: wPc,
        preFinGap: A + k - j0,
        newDebt: wD,
        newDebtInt: E,
        equity: wEq,
        newShares: wSh,
        buybackPlan: bbP, buyback: bbA, buybackCut: bbP - bbA, buybackShares: bbSh, // MAG v0.1b：回購（計畫、實際、被迫減少、股數）
        fcf: (r === 0 ? ACTUAL_1H.cfo - ACTUAL_1H.cashCapex : 0) + A - tx - D - _, // MAG v0.1b：自由現金流（公司定義近似：營運現金 − 現金資本支出；首期含年初至今實際）
        cumNewShares: WF.sh,
        newDebtBeg: wDn0,
        newDebtEnd: WF.Dn,
        facBeg: wFr0,
        facEnd: WF.fr,
        backlogEnd: wB,
        debtCap: wCapB,
        dividend: dvC + dvP,
        divCommon: dvC,
        divPref: dvP,
        divShares: dvSh,
        existDebtEnd: wEx,
        totalDebtEnd: wEx + WF.Dn + WF.Cn + WF.Jn,
        preCash: wPre0, // MAG v0.1b：融資前現金（回購前；Excel 同列）
        need: wX,
        capD: wCapD,
        year: n,
        scheduled: o,
        capacity: u,
        ai: s,
        bottleneck: d,
        revenue: f,
        loss: m,
        collected: h,
        rpoCash: g,
        newRev: nR,
        newCash: nC,
        isRev: f + nR,
        unsold: Math.max(0, u - f - nR),
        oci36: (f + nR) * W36[r], // v0.1b：評價日起 36 個月的 MW 驅動營收（RPO 36 個月內轉換比例的對照）
        gross: _,
        external: v,
        cashCapex: y,
        lease: S,
        interest: D,
        overdraft: E,
        div: e.jvCommit[r] + e.a.div[r],
        jvC: e.jvCommit[r],
        jvF: e.a.div[r],
        intStock: IX[r],
        pBeg: PB[r][0],
        pEnd: PB[r][1],
        mwNew: MN[r],
        mwNext: MX[r],
        capexFull: CXF[r],
        capexFormulaFY: CXF[r] + CXNF[r], // MAG v0.1b：MW 公式全年（AI 成長型＋非 AI；不含汰換）
        capexNonAiFull: CXNF[r], capexNonAi: CXN[r], capexAi: CXA[r], costMW: CMW[r], costIT: CIT[r], // MAG v0.1b：AI／非 AI 資本支出
        daAi: DAI[r], daNonAi: DAN[r], aiItBeg: CXM ? AI_IT[r][0] : 0, aiFacBeg: CXM ? AI_FC[r][0] : 0, nonAiBeg: CXM ? NA_P[r][0] : 0,
        capexGrowth: CXG[r],
        capexGrowthCash: CXG[r] * (1 - e.a.customerFund[r]), // v0.1c：成長型 CapEx 扣客戶預付後（終值基準加回用）
        refresh: REF[r],
        refreshFlag: RFF[r] ? 1 : 0,
        refreshSteadyV: RFS[r],
        refreshVintage: RFV[r],
        billDelayed: BD[r], // v0.2
        capexCum: CXC[r], capexInSvcCum: DM > 0 ? CXSC[r] : CXC[r], capexInSvc: CXS[r], idleCap: IDLE[r], // v0.2：閒置資本
        avgAccepted: (MB[r] + t.accepted[r]) / 2,
        ebitdarM: eR,
        ppeBeg: PPE[r],
        daFleet: DAF[r],
        capexOld: CX_OLD[r],
        intOld: INT_OLD[r],
        rentPerMW: (b + x) / L / ((MB[r] + t.accepted[r]) / 2) * 1e3,
        rentBench: (MB[r] + t.accepted[r]) / 2 * siteBenchQ(e.sites) / 1e3 * LEASE_FACTS.share * L,
        legacy: svcCash,
        servicesRev: svc,
        legacyRev: lgR,
        legacyEbitda: lgE,
        legacyOa: lgO, // MAG v0.1b：非 AI 事業其他攤銷
        aiRevQ: AIR[r],
        cashTax: tx,
        ebM: ebM,
        cashMargin: cm,
        totRev: totRev,
        ebitdaPL: totRev * ebM + ob + lgE - pen,
        capUndelayed: u0, lostRev: lost, delayPen: pen, // v0.2
        leaseLiabOn: LLON[r], leaseLiabUl: LLUL[r], leaseLiab: LLON[r] + LLUL[r], ebitdarAnn: (totRev * ebM + ob + lgE - pen + S) / L, // v0.2：租賃負債與 EBITDAR（年化）
        adjLev: (wEx + WF.Dn + WF.Cn + WF.Jn + LLON[r] + LLUL[r]) / Math.max((totRev * ebM + ob + lgE - pen + S) / L, .01), // v0.2：調整後槓桿（期末）
        otherEbitda: ob,
        rentedCompute: RENTC[r], sbcCash: sbcC, wcStub: r === 0 ? e.wcStub ?? null : null, onBalLeaseAll: b0, offLeaseAll: x0, // MAG v0.1b
        cashEbitda: g + nC + svcCash + ob + lgE - lgO - pen - S,
        creditAdj: (m + nR * (t.defaultP[r] / 100) * p) * cm,
        atm: k,
        facility: F,
        power: C,
        maint: w,
        op: T,
        debtPay: O,
        sources: M,
        uses: j,
        gap: N,
        operatingGap: ee,
        cum: a,
        avgBillable: l,
        onBalLease: b0,
        offLease: x0,
        leaseCashOn: b, leaseCashOff: x, // MAG v0.1b：自現金扣除的租金（operatingInEbitda 時只含融資部分）
        sourcesOp: A
      }
    }),
    hOp = ACTUAL_1H.cfo - ACTUAL_1H.cashCapex - ACTUAL_1H.jv - (ACTUAL_1H.dividends || 0) - (ACTUAL_1H.buyback || 0), // MAG v0.1b：年初至今回購 // v0.1b：年初至今股利與模型期一樣列在營運缺口
    hFin = ACTUAL_1H.borrow - ACTUAL_1H.cappedCall + ACTUAL_1H.equity,
    hPlug = e.cash - (ACTUAL_1H.cash1231 + hOp + hFin - ACTUAL_1H.debtRepaid),
    _fy = (o[0].fyRevenue = ACTUAL_1H.revenue + o[0].isRev + o[0].legacyRev, o[0].fyGross = ACTUAL_1H.capex + o[0].gross, o[0].fyCashCapex = ACTUAL_1H.cashCapex + o[0].cashCapex, o[0].fyLease = ACTUAL_1H.leasePaid + o[0].lease, o[0].fyInterest = ACTUAL_1H.interest + o[0].interest, o[0].fyDebtPay = ACTUAL_1H.debtRepaid + o[0].debtPay, o[0].fyDiv = ACTUAL_1H.jv + o[0].div, o[0].fyAtm = ACTUAL_1H.equity - ACTUAL_1H.cappedCall + o[0].atm, o[0].fyBorrow = ACTUAL_1H.borrow, o[0].fySourcesOp = ACTUAL_1H.cfo + o[0].sourcesOp, o[0].fyOperatingGap = hOp + o[0].operatingGap, o[0].fyExternal = ACTUAL_1H.prepay + o[0].external, o[0].h1 = ACTUAL_1H, o[0].hOp = hOp, o[0].hFin = hFin, o[0].hPlug = hPlug, o[0].fyCfo = ACTUAL_1H.cfo, o[0].fyCapexUse = ACTUAL_1H.cashCapex + o[0].gross, o[0].fyEquity = ACTUAL_1H.equity + o[0].atm, o[0].fyCapped = ACTUAL_1H.cappedCall, o[0].fyUsesCash = ACTUAL_1H.cashCapex + o[0].gross + o[0].lease + o[0].interest + ACTUAL_1H.jv + o[0].div + ACTUAL_1H.debtRepaid + o[0].debtPay + ACTUAL_1H.cappedCall + o[0].op + o[0].cashTax + (ACTUAL_1H.dividends || 0) + o[0].dividend + (ACTUAL_1H.buyback || 0) + o[0].buyback, o[0].fyBuyback = (ACTUAL_1H.buyback || 0) + o[0].buyback, o[0].fyDividend = (ACTUAL_1H.dividends || 0) + o[0].dividend, o[0].fySrcTotal = ACTUAL_1H.cfo + o[0].sourcesOp + ACTUAL_1H.equity + o[0].atm + ACTUAL_1H.borrow + o[0].newDebt + o[0].equity, 0),
    s = e => o.reduce((t, n) => t + n[e], 0),
    c = rA(e),
    l = iA(e),
    u = s(`gross`),
    d = s(`collected`),
    f = s(`cashCapex`) + s(`op`) + s(`lease`) + s(`interest`) + s(`div`) + s(`debtPay`),
    p = Math.max(0, f - s(`legacy`)) / Math.max(d + s(`newRev`), 1),
    m = (e.rpoOpen + e.rpoPendingAdd) * (1 - i),
    h = u * (e.terminal.residual / 100) * (e.terminal.rerent / 100) + m * (e.terminal.margin / 100) - c,
    g = [],
    _ = e => g.push(e);
  _({
    id: `rev-mw`,
    ok: o.every(y => Math.abs(y.isRev - y.capacity) < 1e-9) || e.revenueDriver !== `mw`,
    severity: `ok`,
    title: `營收＝平均在役 MW × 每 MW 年收入 × 利用率 × 期間長度`,
    detail: `${e.revenueDriver === `mw` ? `MW 驅動` : `RPO 驅動`}：五期算力營收 ${o.map(y => Y(y.isRev, 2)).join(`／`)}；容量上限 ${o.map(y => Y(y.capacity, 2)).join(`／`)}。每 MW 年收入 ${t.revMW.map(x => Y(x * 1e3, 2)).join(`／`)} US$m/MW-IT（${PRICING && e.priceLink ? `Tokenomics 持有成本 × k ${Y(t.price[0].k, 3)}` : `Tokenomics 正向推導`}；${TXQ.revMwCompare}）。`
  }), PRICING && e.priceLink && _({ // MAG v0.1b：上限檢查——每 MW 年收入 ÷ 參考付費營收（IF_RevGWFleet）> 門檻時示警
    id: `cap-ratio`,
    ok: t.price.every(x => x.capRatio <= PRICING.capWarn),
    severity: t.price.every(x => x.capRatio <= PRICING.capWarn) ? `ok` : `watch`,
    title: `每 MW 年收入 ÷ Tokenomics 參考付費營收：最高 ${hA(Math.max(...t.price.map(x => x.capRatio)) * 100, 1)}（門檻 ${hA(PRICING.capWarn * 100, 0)}）`,
    detail: `五期 ${t.price.map(x => hA(x.capRatio * 100, 1)).join(`／`)}；參考付費營收＝Σ 世代占比 × IF_RevGWFleet（1 GW 參考機隊付費營收，理論上限）。超過門檻表示 hyperscaler 每 MW 收入占客戶 token 營收的比例偏高（${PRICING.capWarnNote}）。`
  }), LG.lines.filter(x => x.kind === `cloudResidual`).forEach(x => _({ // MAG v0.1b：非 AI 雲端殘差的合理性（殘差為正、推估年增率介於 0 與雲端分部年增率之間）
    id: `residual-${x.key}`,
    ok: x.rev[0] > 0 && x.g0 >= 0 && x.g0 <= (B_LINE(x.key).g4q ?? 1),
    severity: x.rev[0] > 0 && x.g0 >= 0 && x.g0 <= (B_LINE(x.key).g4q ?? 1) ? `ok` : `watch`,
    title: `${x.label}（殘差）：${PERIODS[0]} 模型期 ${Y(x.rev[0], 2)}bn、推估近四季年增率 ${hA(x.g0 * 100, 1)}（雲端分部 ${hA((B_LINE(x.key).g4q ?? 0) * 100, 1)}）、EBITDA 率 ${hA(x.m0 * 100, 1)}`,
    detail: `＝雲端分部首期估計 ${Y(x.ai.segFy0, 2)} − AI 年初至今估計 ${Y(x.ai.ytd, 2)} − AI 模型期 ${Y(o[0].isRev, 2)}；AI 近四季估計 ${Y(x.ai.ttm, 2)}、前四季 ${Y(x.ai.prevTTM, 2)}。殘差為負或年增率超出區間時示警（不回頭改 k 或 MW；敏感度見報告）。`
  })), CALL_FACTS.aiRunRate != null && (() => { // MAG v0.1b：公司 AI run-rate 驗證列（對照表 r1 C3）——不作輸入；寫出要對上須多少 MW 或 k
    const r0 = t.revMW[0] * (e.revScale ?? 1), mdl = e.billableOpen * r0, rr = CALL_FACTS.aiRunRate;
    return _({ id: `ai-runrate`, ok: mdl >= rr, severity: `watch`,
      title: `公司 AI run-rate >${Y(rr, 0)}B vs 模型評價日年化 ${Y(mdl, 1)}B（${hA((mdl / rr - 1) * 100, 1)}）`,
      detail: `模型＝評價日在役 ${Y(e.billableOpen, 0)} MW × 首期每 MW 年收入 ${Y(r0 * 1e3, 2)}m。要對上公司說法：對外 MW 需 ${Y(rr / r0, 0)}（k 不變），或 k 需 ${Y((t.price ? t.price[0].k : 1) * rr / mdl, 3)}（MW 不變）。${CALL_FACTS.aiRunRateNote || ``}` })
  })(), COMPANY_DATA.related && (() => { // MAG v0.1b：關聯方——合約年化與兩家模型快照並排（只讀、不連動）；對手方集中度＝合約年化 ÷ 對外 AI 雲端收入（年化）
    const R = relatedQ(o);
    return _({ id: `related`, ok: !0, severity: `watch`,
      title: `對手方集中度：關聯方合約年化 ${PERIODS[2]} $${Y(R.tot[2], 1)}bn ÷ 對外 AI 雲端收入 $${Y(R.ai[2], 1)}bn＝${hA(R.ratio[2] * 100, 0)}（RPO 上限約 ${hA(COMPANY_DATA.related.rpoShareMax * 100, 0)}）`,
      detail: `${COMPANY_DATA.related.contracts.map((c, j) => `${c.name} ${R.rows[j].map(x => Y(x, 1)).join(`／`)}`).join(`；`)}；模型快照：${COMPANY_DATA.related.snapshots.map((c, j) => `${c.name} ${R.snap[j].map(x => x == null ? `—` : Y(x, 1)).join(`／`)}`).join(`；`)}（US$bn／年，${PERIODS.join(`／`)}；只讀，不回饋計算）。` })
  })(), _({
    id: `rpo-weights`,
    ok: Math.abs(RPO_BUCKET_W.reduce((e, t) => e + t, 0) - RPO_SCHEDULED_SHARE) < 1e-6,
    severity: `watch`,
    title: `RPO 桶 → 五期權重 [Derived]（只作對照，不驅動營收）`,
    detail: `季報：RPO $${Y(LATEST_Q.rpo, 1)}bn，${COMPANY_DATA.rpo.bucketLabels.map((b, j) => `${b} ${hA(COMPANY_DATA.rpo.split[j] * 100, 0)}`).join(`、`)}（${TXQ.rpoNote}）。桶內線性分攤得五期 ${RPO_BUCKET_W.map(x => hA(x * 100, 1)).join(`／`)}，合計 ${hA(RPO_SCHEDULED_SHARE * 100, 0)}。營收由 MW × 每 MW 年收入驅動，RPO 排程只用於產能瓶頸旗標；桶內前載或後載是 Derived，不是 Verified。`
  }), _({
    id: `legacy-fy`,
    ok: LG.lines.every(x => Math.abs(x.ytd + x.rev[0] - x.annual[0]) < 1e-9),
    severity: `ok`,
    title: `非 AI 事業：${PERIODS[0]} 全年 ${Y(LG.annual[0], 2)}bn（年初至今實際 ${Y(LG.lines.reduce((a, x) => a + x.ytd, 0), 3)}＋模型 ${Y(LG.rev[0], 2)}）`,
    detail: `${LG.lines.map(x => `${x.label} ${Y(x.annual[0], 2)}（年增 ${hA(x.g[0] * 100, 1)} → ${PERIODS[4]} ${hA(x.g[4] * 100, 1)}）`).join(`；`)}。EBITDA 率 ${LG.margin.map(m => hA(m * 100, 1)).join(`／`)}（${TXQ.legacyMarginNote}）。${PERIODS[0]} 合計營收 ${Y(ACTUAL_1H.revenue + o[0].isRev + o[0].legacyRev, 2)}bn 對照公司指引 ${REV_GUIDE_TXT}。`
  }), _({
    id: `rpo-36m`,
    ok: !0,
    severity: `watch`,
    title: `RPO 36 個月內轉換 vs 模型（只作對照）`,
    detail: `季報：RPO $${Y(LATEST_Q.rpo, 1)}bn 中約 ${hA(COMPANY_DATA.rpo.within36m * 100, 0)} 預計 36 個月內認列＝$${Y(LATEST_Q.rpo * COMPANY_DATA.rpo.within36m, 1)}bn（含非 AI 事業 RPO）。模型評價日起 36 個月的 MW 驅動營收 $${Y(o.reduce((a, y) => a + y.oci36, 0), 1)}bn（各期權重 ${W36.map(x => hA(x * 100, 0)).join(`／`)}）。差額不回推每 MW 單價。`
  }), _({
    id: `lease-10q`,
    ok: !0,
    severity: `ok`,
    title: `在帳租賃現金（季報到期表）`,
    detail: `營業租賃未折現付款：五期 ${LEASE_CASH_ON_BAL.map(x => Y(x, 3)).join(`／`)}，其後 ${Y(LEASE_AFTER_FY30, 2)}；租賃負債現值 ${Y(LATEST_Q.opLeaseLiab, 3)}（另融資租賃 ${Y(LATEST_Q.finLeaseLiab, 3)}）。${TXQ.leaseNote}；不含 ${Y(LATEST_Q.offBalanceLease, 1)}bn 未起租。`
  }), _({
    id: `off-balance-lease`,
    ok: l >= (LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap) * CHECK_TH.leaseVsCommitMin,
    severity: `watch`,
    title: `已簽約未起租租賃 $${Y(LATEST_Q.offBalanceLease, 1)}bn`,
    detail: `季報附註：已簽約未起租租賃未折現 $${Y(LATEST_Q.offBalanceLease, 2)}bn（${TXQ.offBalanceLeaseTerm}）。模型排程：平均分 ${UL.quarters} 季起租（評價日後第 ${UL.startQ + 1} 季起）、每筆期限 ${UL.termYears} 年直線付租，三情境相同；五期 ${e.a.newLease.map(x => Y(x, 2)).join(`／`)}（合計 ${e.a.newLease.reduce((e,t)=>e+t,0).toFixed(1)}）＋模型期後 ${ulTail(e.a.newLease).toFixed(1)} = ${l.toFixed(1)}bn，為已承諾 ${Y(LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap, 1)}bn 的 ${(l/Math.max(LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap, .01)).toFixed(1)} 倍；低於 ${hA(CHECK_TH.leaseVsCommitMin * 100, 0)} 才標示不一致。`
  }), _({
    id: `rou-h1`,
    ok: !0,
    severity: `ok`,
    title: `營業租賃負債 $${Y(LATEST_Q.opLeaseLiab, 2)}bn`,
    detail: `使用權資產對應的租賃負債 ${Y(LATEST_Q.opLeaseLiab, 3)}bn（${TXQ.leaseLiabNote}）。$${Y(LATEST_Q.offBalanceLease, 1)}bn 尚未起租，還沒有使用權資產。`
  });
  let v = o[0].gross,
    y = v >= CALL_FACTS.capexLo - LATEST_Q.capexH1 - .5 && v <= CALL_FACTS.capexHi - LATEST_Q.capexH1 + .5;
  _({
    id: `h2-capex`,
    ok: y,
    severity: y ? `ok` : `watch`,
    title: `${PERIOD_FY[0]} CapEx 指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi} → ${PERIODS[0]} 模型期 ${(CALL_FACTS.capexLo-LATEST_Q.capexH1).toFixed(1)}–${(CALL_FACTS.capexHi-LATEST_Q.capexH1).toFixed(1)}`,
    detail: `全年資本支出指引 $${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}bn（${TXQ.capexGuideSource}）；年初至今現金資本支出 ${Y(LATEST_Q.capexH1, 2)}（應計口徑未揭露，以現金口徑代替）。模型 ${PERIOD_FY[0]}（年初至今實際＋首期模型）毛額 ${(v+ACTUAL_1H.capex).toFixed(1)}。`
  }), _({
    id: `prepay`,
    ok: !0,
    severity: `watch`,
    title: `客戶預付：合約負債 $${Y(LATEST_Q.deferredTotal, 2)}bn，年初至今淨增 $${Y(LATEST_Q.deferredIn, 2)}bn`,
    detail: `季報：遞延營收（合約負債）${Y(LATEST_Q.deferredTotal, 3)}bn，年初至今增加 ${Y(LATEST_Q.deferredIn, 3)}；年初至今客戶預付現金 ${Y(ACTUAL_1H.prepay, 3)}。${TXQ.prepayCheck}模型首期預付流入 ${Y(o[0].external, 2)}bn（${PERIOD_FY[0]} 全年 ${Y(ACTUAL_1H.prepay + o[0].external, 2)}）、認列 ${Y(o[0].prepayRecog, 2)}；期初合約負債 ${Y(e.prepay.openBalance, 3)}。`
  }), _({
    id: `call-mw`,
    ok: !0,
    severity: `watch`,
    title: `已連網 MW-IT：${PERIOD_FY[0] - 1} 年底 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW → ${PERIODS[0]} 年底 ${t.accepted[0]} MW`,
    detail: `${TXQ.mwFacts}口徑不明者以 ÷ PUE 1.2 換成 MW-IT。目前情境：${PERIODS[0]} 已連網 ${t.accepted[0]} MW、${PERIODS[1]} ${t.accepted[1]} MW、${PERIODS[4]} ${t.accepted[4]} MW。累計在役 MW 未揭露，不以營收反推（期初可計費 MW 以實際營收校準）。`
  }), _({
    id: `call-util`,
    ok: t.util[0] <= 100,
    severity: `watch`,
    title: `利用率 ${t.util[0]}%：每 MW 年收入已含可計費利用率`,
    detail: `每 MW 年收入＝Tokenomics 持有成本 × k，按 MW-year 計價，所以利用率欄預設 100%（收入端不乘 IF_Util）；欄位保留供壓力測試。目前每 MW 年收入 $${Y(t.revMW[0] * 1e3 * (e.revScale ?? 1), 2)}m/MW-IT。`
  });
  for (let e = 0; e < 5; e++) t.billable[e] - t.accepted[e] > .5 && _({
    id: `billable-${e}`,
    ok: !1,
    severity: `block`,
    title: `${PERIODS[e]} 可計費 > 已驗收`,
    detail: `${t.billable[e]} MW > ${t.accepted[e]} MW。引擎已向下截斷。`
  });
  e.linkSites && r.accepted > e.m.accepted[0] + .5 ? _({
    id: `site-floor`,
    ok: !0,
    severity: `watch`,
    title: `站點已驗收拉高 ${PERIODS[0]}`,
    detail: `具名站點已驗收 ${r.accepted} MW，${PERIODS[0]} 已連網取 max(輸入, 站點)。`
  }) : !e.linkSites && r.accepted > t.accepted[0] + 1 && _({
    id: `site-mismatch`,
    ok: !1,
    severity: `block`,
    title: `站點與年度 MW 不一致`,
    detail: `具名站點已驗收 ${r.accepted} MW，高於 ${PERIODS[0]} 已連網 ${t.accepted[0]} MW。請打開站點連動。`
  }), t.accepted[4] > r.planned + 50 && _({
    id: `unlisted-mw`,
    ok: !0,
    severity: `watch`,
    title: `未列名容量`,
    detail: `${PERIODS[4]} 已連網 ${t.accepted[4]} MW 高於具名規劃 ${r.planned} MW，差額進「其他／未列名站點」。公司多數站點未逐站揭露 MW。`
  }), e.overlay && _({
    id: `overlay`,
    ok: !1,
    severity: `watch`,
    title: `Overlay 與貢獻率`,
    detail: `電力＋維護已另扣；EBITDAR 率由 EBITDA 率推導、已含電費，開啟 overlay 會重複扣除，僅供壓力測試。`
  }), e.cdsLink && _({
    id: `cds-link`,
    ok: !0,
    severity: `watch`,
    title: `CDS 傳入透支利率`,
    detail: `透支利率 = 新債利率 + max(0, CDS−${e.cdsBaseBp})×${e.cdsPassThrough}，只在累積現金為負時生效。沒有 CDS 報價時不適用。`
  }), e.includeDebt ? _({
    id: `debt-sched`,
    ok: !0,
    severity: `ok`,
    title: `債務排程攤還（季報到期表，基本情境）`,
    detail: `五期還本 ${DEBT_AMORT.map(x => Y(x, 3)).join(`／`)}，其後 ${Y(COMPANY_DATA.debt.amortAfterFY30, 3)}（合計 ${Y(DBT_P, 3)}）。可轉債以到期累積本金計，屬契約排程，非壓力測試。關閉＝假設全額再融資。`
  }) : _({
    id: `debt-sched-off`,
    ok: !0,
    severity: `watch`,
    title: `債務攤還已關閉（假設全額再融資）`,
    detail: `關閉等於假設 ${Y(DEBT_AMORT.reduce((a, b) => a + b, 0), 1)}bn 五期本金全部借新還舊。`
  }), _({
    id: `identity`,
    ok: Math.abs(o[4].cum - (e.cash + o.reduce((e, t) => e + t.gap, 0))) < .01 && Math.abs(o[4].cum - (e.cash + o.reduce((e, t) => e + t.operatingGap + t.atm + t.facility - t.debtPay, 0))) < .01,
    severity: `ok`,
    title: `現金恆等式（期前融資瀑布）`,
    detail: `期末 ${o[4].cum.toFixed(1)} = 評價日現金 ${e.cash} + 營運缺口 ${o.reduce((e,t)=>e+t.operatingGap,0).toFixed(1)} + 期後股權／可轉債 ${o.reduce((e,t)=>e+t.atm,0).toFixed(1)} + 瀑布新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)} + 瀑布可轉債 ${o.reduce((e,t)=>e+t.convNew,0).toFixed(1)} + 瀑布股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)} + 高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)} − 排程還本 ${o.reduce((e,t)=>e+t.debtPay,0).toFixed(1)}。每期期末現金不低於最低現金 ${e.minCash}bn——缺口在發生前一期就先融好。`
  }), _({
    id: `waterfall`,
    ok: !0,
    severity: `watch`,
    title: `融資瀑布：新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)}bn、可轉債 ${o.reduce((e,t)=>e+t.convNew,0).toFixed(1)}bn、股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)}bn（新股 ${o.reduce((e,t)=>e+t.newShares,0).toFixed(2)}bn 股）、高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)}bn`,
    detail: `順序：客戶預付（營運來源）→ 現金（高於最低現金 ${e.minCash}bn 的部分）→ ${e.useFacility ? `未動用額度（${TXQ.facilityName}）→ ` : ``}新債（${e.debtCapBasis === `leaseAdj` ? `(總債務＋租賃負債) ≤ ${multTxt(e.debtEbitdaMax)}×(EBITDA＋租金)（投資級上限，租賃調整後槓桿）` : e.debtCapBasis === `ebitda` ? `總債務 ≤ ${multTxt(e.debtEbitdaMax)}× 當期 EBITDA（投資級上限）` : `總債務 ≤ ${e.debtBacklog}× backlog`}）→ ${(e.cvCap ?? 0) > 0 ? `可轉債（每年上限 ${Y(e.cvCap ?? 0, 1)}bn、票息 ${hA(e.convIssue.coupon * 100, 1)}）→ ` : ``}股權（發行價＝$${e.eqPx}×(1−${(e.eqDisc*100).toFixed(0)}%)，每年上限＝現市值 ${e.eqCapPct>=9?`無上限`:(e.eqCapPct*100).toFixed(0)+`%`}）→ 超出部分以高息債 ${(e.junkRate*100).toFixed(0)}% 補足${e.debtCapBasis === `ebitda` || e.debtCapBasis === `leaseAdj` ? `（＝需失去投資級才能融資的金額）` : ``}。股利 ${Y(o.reduce((a, t) => a + t.dividend, 0), 1)}bn 列為用途。${(q => q.length ? `本情境股權需求落在 ${q.join(`、`)}。` : `本情境不需股權。`)(o.filter(t => t.equity > .05).map(t => t.year))}`
  }), _({
    id: `ebitda-link`,
    ok: o.every(y => Math.abs(y.cashEbitda + y.creditAdj + (y.legacyOa || 0) + (e.ebitdaBasis === `tk` ? y.lease : 0) - y.ebitdaPL) < .01), // MAG v0.1b：tk 時 AI 雲端 EBITDA 未扣租金（租金另列現金支出），其他攤銷自現金扣除
    severity: `ok`,
    title: `資金與損益同一組 EBITDA：${PERIODS[4]} EBITDA 率 ${(o[4].ebM*100).toFixed(1)}%、EBITDAR 率 ${(o[4].cashMargin*100).toFixed(1)}%`,
    detail: `${e.ebitdaBasis === `ebitdar` ? `EBITDAR 率由 ${((e.ebStart+e.ebitdarAdj[0])*100).toFixed(1)}% 線性變動至 ${PERIODS[4]} ${((e.ebSteady+e.ebitdarAdj[1])*100).toFixed(1)}%（三情境共用；校準使基準情境起點 ${(e.ebStart*100).toFixed(1)}%、穩態 ${(e.ebSteady*100).toFixed(1)}% EBITDA 率不變），EBITDA 率＝EBITDAR 率 − 租金÷營收（租金為固定成本）。` : `EBITDA 率由 ${(e.ebStart*100).toFixed(0)}%（${TXQ.ebStartSource}）線性變動至 ${PERIODS[4]} ${(e.ebSteady*100).toFixed(0)}%（穩態）。資金模型用 EBITDAR 率＝EBITDA 率＋租金÷營收（因租金在支出端另列），`}非算力服務現金＝服務營收×同一 EBITDAR 率。恆等式：現金 EBITDA（營運來源不含預付 − 租金）＋信用損失調整＝損益 EBITDA，五期皆成立。`
  }), _({
    id: `capex-mw`,
    ok: o[0].capexFormulaFY >= CALL_FACTS.capexLo && o[0].capexFormulaFY <= CALL_FACTS.capexHi,
    severity: o[0].capexFormulaFY >= CALL_FACTS.capexLo && o[0].capexFormulaFY <= CALL_FACTS.capexHi ? `ok` : `watch`,
    title: `毛 CapEx 由 MW 推導：${PERIODS[0]} 全年 ${o[0].capexFormulaFY.toFixed(1)}${CXM ? `（AI ${o[0].capexFull.toFixed(1)}＋非 AI ${o[0].capexNonAiFull.toFixed(1)}）` : ``}（指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}）`,
    detail: `公式＝(本期新增 MW×(1−λ)＋次期新增 MW×λ)×每 MW 成本。${PERIOD_FY[0] - 1} 年底 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW → ${PERIODS[0]} 年底 ${t.accepted[0]} MW、λ ${(e.lambda*100).toFixed(0)}%、每 MW $${e.a.costMW[0]}m（${TXQ.costMwNote}）。五期（首期為模型部分）合計 ${CX.reduce((e,t)=>e+t,0).toFixed(1)}。指引口徑：${TXQ.capexGuideSource}。`
  }), CXM && CALL_FACTS.capexLo != null && (() => { // MAG v0.1b：對帳列——指引 ÷ 每 MW 全成本＝隱含 AI 建置 MW，對照 MW 路徑（對照表 r1 第 5 節第 6 條）
    const g = (CALL_FACTS.capexLo + CALL_FACTS.capexHi) / 2, imp = (g - CXNF[0] - REF[0]) / (CMW[0] * (e.capexScale ?? 1) / 1e3), pth = (MN[0] * (1 - e.lambda) + MX[0] * e.lambda) / XS - (RXD(0) * (1 - e.lambda) + RXD(1) * e.lambda); // MAG v0.1b r3（C20）：扣租用對外 MW 新增
    return _({ id: `capex-recon`, ok: Math.abs(imp / pth - 1) <= .2, severity: `watch`, imp, pth, // MAG v0.1b r2：imp／pth 供 scripts/calib_pace.js 讀取（C10）
      title: `對帳：${PERIODS[0]} 資本支出指引隱含 AI 建置 ${Y(imp, 0)} MW vs MW 路徑 ${Y(pth, 0)} MW（${hA((imp / pth - 1) * 100, 0)}）`,
      detail: `隱含＝(指引 ${Y(g, 1)} − 非 AI ${Y(CXNF[0], 1)} − 汰換 ${Y(REF[0], 1)}) ÷ 每 MW 成本 $${Y(CMW[0] * (e.capexScale ?? 1), 1)}m；路徑＝(本期新增 ${Y(MN[0], 0)}×(1−λ)＋次期新增 ${Y(MX[0], 0)}×λ) ÷ 對外比例 ${hA(XS * 100, 0)}（對外當量 ${Y(imp * XS, 0)} vs ${Y(pth * XS, 0)} MW）。落差表示指引含 MW 路徑以外的支出（未上線容量的預付與在建、記憶體漲價、網路與土地），或 MW 路徑偏低；首期以指引為準（差額歸 AI 成長型）。` })
  })(), CXM && (() => { // MAG v0.1b r2（對照表 r1 C15）：一致性檢查——k＝1、Tokenomics 成本、無爬坡延遲、無稅的對外 AI 稅前 ROIC vs IF_HoldEcon 隱含報酬
    const Q = aiRoicQ({ years: o, m: t }, e, { tax: COMPANY_DATA.valuation.tax, wacc: 0 }), c = Q && Q.c15; if (!c) return;
    const tol = CHECK_TH.c15Tol ?? .05;
    return _({ id: `c15`, ok: Math.abs(c.clean - c.hurdle) <= tol, severity: `watch`, clean: c.clean,
      title: `一致性：${PERIODS[c.ry]} 對外 AI 稅前 ROIC（k＝1、Tokenomics 成本、穩態、無稅）${hA(c.clean * 100, 1)} vs IF_HoldEcon 隱含 ${hA(c.hurdle * 100, 0)}（模型稅後 ${hA(c.model * 100, 1)}）`,
      detail: `差距 ${Y(c.gap * 100, 1)}pt 依序拆解：` + c.items.map(x => `${x[0]} ${x[1] >= 0 ? `+` : `−`}${Y(Math.abs(x[1]) * 100, 2)}pt`).join(`；`) + `（加總＝總差距；序列拆解，順序固定）。` })
  })(), _({
    id: `fleet-da`,
    ok: !0,
    severity: `watch`,
    title: CXM ? `D&A 分池：${PERIODS[4]} AI ${o[4].daAi.toFixed(1)}＋非 AI ${o[4].daNonAi.toFixed(1)}＋其他攤銷 ${o[4].legacyOa.toFixed(1)}bn` : `D&A 改由車隊推算：${PERIODS[4]} ${o[4].daFleet.toFixed(1)}bn（壽命 ${e.gpuLife} 年）`,
    detail: CXM ? `AI：IT 依 GPU 壽命 ${e.gpuLife} 年、機房依 ${Y(AP0.facLife, 1)} 年（期初毛額估計 IT ${Y(AP0.it, 1)}、機房 ${Y(AP0.fac, 1)}）；非 AI：期初基礎 ${Y(AP0.n0, 1)}、折舊年限 ${Y(AP0.nLife, 1)} 年（預設年限；最新季分部 D&A 年化 ${Y(COMPANY_DATA.capexModel.segDaRunRate, 1)} 只作對帳，C17）；其他攤銷依各線占營收比（C5）。季報 D&A ${Y(LATEST_Q.da, 3)}／季。` : `D&A＝(期初毛 PP&E＋本期成長型 CapEx×½)÷壽命。期初 PP&E 基礎 ${e.ppeOpen}（${TXQ.ppeOpenNote}）；季報 D&A ${Y(LATEST_Q.da, 3)}／季。CapEx 隨 MW 增加，折舊跟著增加。`
  }), CXM && (() => { const R = AP0.daRecon, B = COMPANY_DATA.capexModel.segDaRunRate, tol = CHECK_TH.daReconTol ?? .1; // MAG v0.1b r3（C17）：D&A 對帳殘差
    return _({ id: `da-recon`, ok: Math.abs(R) <= tol * B, severity: `watch`, resid: R,
      title: `D&A 對帳：最新季分部 D&A 年化 ${Y(B, 1)} vs AI 期初 ${Y(AP0.da0, 1)}＋非 AI ${Y(AP0.n0 / AP0.nLife, 1)}（殘差 ${Y(R, 1)}，${hA(R / B * 100, 0)}）`,
      detail: `非 AI 折舊＝非 AI 期初基礎 ${Y(AP0.n0, 1)} ÷ 預設年限 ${Y(AP0.nLife, 0)} 年（capexModel.nonAiLife，區間 ${(COMPANY_DATA.capexModel.nonAiLifeRange || []).join('–')} 年 [Assumed]）＋新增非 AI 資本支出同年限；AI 期初折舊依 Tokenomics。殘差不回填（容許 ±${hA(tol * 100, 0)}）：正值＝模型 D&A 低於公司實際（年限偏長或 AI 期初估計偏低）。` })
  })(), _({
    id: `gpu-refresh`,
    ok: !0,
    severity: `watch`,
    title: `GPU 汰換 CapEx：${PERIODS[3]} ${o[3].refresh.toFixed(1)}、${PERIODS[4]} ${o[4].refresh.toFixed(1)}`,
    detail: `${e.refreshSteady ? `已連網 MW 不再增加的期間（觸頂後）及終值年採穩態汰換＝平均已連網 MW × 每 MW GPU 成本 ÷ 壽命 ${e.gpuLife} 年（建築與電力屬租賃不計；客戶出資覆蓋比同樣適用），終值以含汰換的末期 UFCF 為基準。其餘期間：` : ``}${PERIOD_FY[0] - 1} 年底批次（${e.mwYearEnd[PERIOD_FY[0] - 1]} MW）在第 ${e.gpuLife} 年汰換（落在模型期之後則不出現）；更早批次的 MW 未揭露。汰換取代已折舊完的設備，不增加折舊基礎。`
  }), _({
    id: `capex-floor`,
    ok: o[0].capexFormulaFY < (e.capexFloorFY0 ?? 0) ? !1 : !0,
    severity: `watch`,
    title: o[0].capexFormulaFY < (e.capexFloorFY0 ?? 0) ? `${PERIODS[0]} CapEx 公式值 ${o[0].capexFormulaFY.toFixed(1)} 低於下限 ${e.capexFloorFY0}：差額 ${(e.capexFloorFY0 - o[0].capexFormulaFY).toFixed(1)}${CXM ? `（歸 AI 成長型）` : ` 為轉向太晚的成本`}` : `${PERIODS[0]} CapEx 公式值 ${o[0].capexFormulaFY.toFixed(1)} 高於下限 ${e.capexFloorFY0}`,
    detail: `${PERIOD_FY[0]} 的支出多已下單（全年指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}，年初至今 ${Y(ACTUAL_1H.capex, 1)}）。若次年新增 MW 少到公式值低於下限，代表已採購的設備超過實際上線需求——這部分在模型中不帶來額外收入。`
  }), _({
    id: `debt-sched-int`,
    ok: Math.abs(DBT_P + CVN.reduce((a, n) => a + (n.mand ? 0 : n.M), 0) - (LATEST_Q.debtPrincipal - COMPANY_DATA.debt.convertibleBridge.exchangedAccreted + COMPANY_DATA.debt.convertibleBridge.newIssuesAccreted)) < .01,
    severity: `ok`,
    title: `存量利息由既有債務與可轉債逐檔推算`,
    detail: `其他借款 ${DBT_P.toFixed(3)}＋可轉債到期本金 ${Y(CVN.reduce((a, n) => a + (n.mand ? 0 : n.M), 0), 3)}（強制轉換特別股為權益，不列入），對照季報本金 ${LATEST_Q.debtPrincipal} − 以股換債 ${COMPANY_DATA.debt.convertibleBridge.exchangedAccreted} ＋ 期後新發 ${COMPANY_DATA.debt.convertibleBridge.newIssuesAccreted}。存量利息＝其他借款利息＋債務處理可轉債票息（原始本金 × 票息；到期當期計半年）＋首期校準 ${e.intCal}；價內可轉債（有效轉換價 < 現價）以若轉換法計，不計利息與還本。可轉債以現金票息計，不含折價攤銷（非現金）。五期合計 ${o.reduce((e,t)=>e+t.intStock,0).toFixed(2)}。`
  }), _({
    id: `jv-commit`,
    ok: Math.abs(e.jvCommit.reduce((e,t)=>e+t,0) - COMPANY_DATA.defaults.jvCommit.reduce((a,b)=>a+b,0)) < .01,
    severity: `ok`,
    title: `JV 已承諾出資：${Y(e.jvCommit.reduce((a, b) => a + b, 0), 2)}bn`,
    detail: `季報未揭露 JV 出資承諾（不適用）；年初至今收購子公司淨額 ${Y(ACTUAL_1H.jv, 3)} 列為策略投資。後續增資 ${e.a.div.join('／')}，屬 [Assumed]。`
  }), _({
    id: `rent-bench`,
    ok: o[4].rentPerMW >= siteBenchQ(e.sites) * LEASE_FACTS.share * CHECK_TH.rentVsBenchMin,
    severity: `watch`,
    title: `租金檢驗：模型每 MW 年租金 ${PERIODS[4]} $${o[4].rentPerMW.toFixed(2)}m vs 市場基準 $${siteBenchQ(e.sites).toFixed(2)}m`,
    detail: `具名站點${siteBenchQ(e.sites) > 0 ? `加權每 MW 年租金 $${siteBenchQ(e.sites).toFixed(2)}m` : `沒有租約金額（自有為主），市場租金基準不適用`}。模型每 MW 年租金 ${PERIODS[0]} $${o[0].rentPerMW.toFixed(2)}m、${PERIODS[4]} $${o[4].rentPerMW.toFixed(2)}m。`
  }), _({
    id: `facility`,
    ok: !0,
    severity: `ok`,
    title: e.useFacility ? `未動用額度 ${e.facility}bn（${TXQ.facilityName}）為瀑布第一順位` : `未動用額度不動用`,
    detail: e.useFacility ? `${CALL_FACTS.postQShortDated || ``}。已承諾額度，不受債務上限限制；用罄後才進入新債與股權。` : `${TXQ.facilityName}：不列入瀑布（流動性備援；新債一律受債務上限約束）。`
  }), _({
    id: `cds-level`,
    ok: e.cds == null || e.cds < 800,
    severity: e.cds != null && e.cds >= 800 ? `watch` : `ok`,
    title: e.cds == null ? `CDS：${e.cdsDate}` : `5Y CDS ${Y(e.cds,0)} bps（${e.cdsDate}）`,
    detail: e.cds == null ? `沒有 CDS 報價資料；${TXQ.newDebtRateNote}` : `買賣 ${Y(e.cdsBid,0)}／${Y(e.cdsAsk,0)}，近期區間 ${Y(e.cdsLo,0)}–${Y(e.cdsHi,0)}。以 40% 回收率的簡化式估算，隱含年化違約機率約 ${(e.cds/100/(1-.4)).toFixed(1)}%。`
  }), _({
    id: `concentration`,
    ok: !0,
    severity: `watch`,
    title: TXQ.concentration.title,
    detail: TXQ.concentration.detail
  }), _({
    id: `tier34`,
    ok: !0,
    severity: `watch`,
    title: `MW 口徑：合約／已連網電力未說明 IT 或設施口徑`,
    detail: `模型以 MW-IT 為主口徑（Tokenomics 與 active power 皆為 IT）；公司合約電力與已連網電力 ÷ PUE 1.2 換算。若公司數字其實是 IT 口徑，可部署 MW 約多 20%，營收與 CapEx 同步放大。`
  });
  let C = r.accepted,
    w = Math.max(0, t.accepted[0] - C),
    T = w / Math.max(t.accepted[0], 1);
  return _({
    id: `residual-mw`,
    ok: T <= .5,
    severity: T > .5 || T > .3 ? `watch` : `ok`,
    title: `未列名站點占 ${PERIODS[0]} 已連網 ${Math.round(T*100)}%`,
    detail: `具名已驗收 ${C} MW，${PERIODS[0]} 已連網 ${t.accepted[0]} MW，殘差 ${w.toFixed(0)} MW。公司不逐站揭露 MW，殘差必然偏高；超過 50% 表示容量預測大多不是具名站點支撐。`
  }), _({
    id: `fill-timing`,
    ok: !0,
    severity: `watch`,
    title: `營收＝在役 MW × 每 MW 年收入（新產能全數可出租）`,
    detail: `營收由 MW 驅動：已連網 MW 依在役比例轉為可計費，乘每 MW 年收入（Tokenomics 正向推導）；新產能簽約率 ${t.fill[0]}%，RPO 只作對照（${PERIODS[2]}–${PERIODS[4]} 超出期初 RPO 的部分合計 ${(o[2].newRev+o[3].newRev+o[4].newRev).toFixed(0)}bn）。風險：已連網但尚未簽約的產能；續約價格衰退未建模（預付款優勢可能高估）。`
  }), e.useAvgMw || _({
    id: `year-end-mw`,
    ok: !0,
    severity: `watch`,
    title: `收入用年末 Billable 全年化`,
    detail: `預設關閉時用年末存量×全年單價，年中才交付的 MW 會被當成全年在役；建議打開「平均在役 MW」。`
  }), {
    years: o,
    totals: {
      revenue: s(`revenue`),
      gross: u,
      cashCapex: s(`cashCapex`),
      op: s(`op`),
      operatingGap: s(`operatingGap`),
      gap: s(`gap`),
      end: o[4].cum,
      legacy: s(`legacy`),
      legacyRev: s(`legacyRev`),
      legacyEbitda: s(`legacyEbitda`),
      legacyOa: s(`legacyOa`),
      cashTax: s(`cashTax`),
      dividend: s(`dividend`),
      prepayAccr: s(`prepayAccr`),
      newRev: s(`newRev`),
      oci36: s(`oci36`),
      newCash: s(`newCash`),
      lease: s(`lease`),
      interest: s(`interest`),
      overdraft: s(`overdraft`),
      div: s(`div`),
      atm: s(`atm`),
      facility: s(`facility`),
      newDebt: s(`newDebt`),
      convNew: s(`convNew`),
      junk: s(`junk`),
      preFinEnd: o[4].preFinCum,
      equity: s(`equity`),
      newShares: s(`newShares`),
      be: p,
      terminal: h,
      leaseTail: c,
      scheduled: s(`scheduled`),
      bottleneck: s(`bottleneck`),
      collected: s(`collected`),
      debtPay: s(`debtPay`),
      ntmHole: Math.max(0, -o[1].cum),
      capexFull26: o[0].capexFull,
      capexOld: CX_OLD.reduce((e, t) => e + t, 0),
      intStock: s(`intStock`),
      intOld: INT_OLD.reduce((e, t) => e + t, 0),
      jvCommit: s(`jvC`),
      rentGap: o.reduce((e, t) => e + (t.rentBench - t.lease), 0),
      bench: siteBenchQ(e.sites),
      h1Op: hOp,
      h1Fin: hFin,
      h1Plug: hPlug,
      fyRevenue: o[0].fyRevenue,
      fyGross: o[0].fyGross,
      fyInterest: o[0].fyInterest,
      fyLease: o[0].fyLease,
      fyDebtPay: o[0].fyDebtPay
    },
    checks: g,
    lg: LG,
    cvFund: CVF,
    sites: n,
    m: t,
    leaseTail: c
  }
}

function reverseDcf(e, v) {
  let P = v.price,
    run = (m, q) => {
      let n = structuredClone(e),
        r = structuredClone(v);
      m && m(n), q && q(r);
      let i = runFunding(n),
        a = runValuation(i, n, r);
      return {
        dcf: a.d.invalid ? 0 : a.d.perShare,
        tgt: a.call.blended,
        d: i,
        p: a
      }
    },
    solve = (f, lo, hi, inc, key = `dcf`) => {
      let a = f(lo)[key] - P,
        b = f(hi)[key] - P;
      if (!(Number.isFinite(a) && Number.isFinite(b)) || a * b > 0) return NaN;
      for (let k = 0; k < 22; k++) {
        let m = (lo + hi) / 2,
          c = f(m)[key] - P;
        (c > 0) === inc ? hi = m : lo = m
      }
      return (lo + hi) / 2
    },
    base = run(),
    fR = x => run(t => {
      t.revScale = x
    }),
    fC = x => run(t => {
      t.capexScale = x
    }),
    fE = x => run(t => {
      t.ebSteady = x
    }),
    R = solve(fR, .5, 4, !0),
    C = solve(fC, .1, 1.5, !1),
    Eb = solve(fE, .3, .99, !0),
    Rt = solve(fR, .5, 4, !0, `tgt`),
    caps = [.7, .8, .9, 1, 1.1],
    ebs = [.35, .47, .59, .7], // v0.1b：可觀察 neocloud 區間（IREN 約 35%、中點 47%、CRWV 約 59%）＋ 70%
    grid = caps.map(c => ebs.map(s => solve(x => run(t => {
      t.capexScale = c, t.ebSteady = s, t.revScale = x
    }), .5, 4, !0)));
  return {
    price: P,
    baseDcf: base.dcf,
    baseTgt: base.tgt,
    R,
    C,
    Eb,
    Rt,
    caps,
    ebs,
    grid,
    rev30: e.m.revMW[4] * (e.revScale ?? 1) * 1e3,
    util30: e.m.util[4] / 100,
    cost30: e.a.costMW[4] * (e.capexScale ?? 1),
    eb30: e.ebSteady
  }
}

function fA(e, t, v, q) {
  let n = structuredClone(e),
    r = structuredClone(v);
  t && t(n), q && q(r);
  let i = runFunding(n),
    a = runValuation(i, n, r);
  return {
    tgt: a.call.blended,
    eq: i.totals.equity,
    gap: Math.max(0, -i.totals.preFinEnd)
  }
}

function sensitivities(e, v) {
  let t = fA(e, null, v),
    n = [],
    r = (r, la, lb, i, a, qi, qa) => {
      let x = fA(e, i, v, qi),
        y = fA(e, a, v, qa);
      n.push({
        name: r,
        la: la,
        lb: lb,
        a: x,
        b: y,
        low: Math.min(x.tgt, y.tgt),
        high: Math.max(x.tgt, y.tgt),
        base: t.tgt,
        baseEq: t.eq,
        baseGap: t.gap
      })
    };
  return r(`穩態 EBITDA 率`, `35%`, `59%`, t => { // v0.1b：可觀察 neocloud 區間（IREN 約 35%、CRWV 約 59%）
    t.ebSteady = .35
  }, t => {
    t.ebSteady = .59
  }), (e.debtCapBasis === `leaseAdj` ? r(`投資級上限（調整後槓桿）`, `4.0x`, `5.0x`, t => { // v0.2：(債務＋租賃負債) ÷ (EBITDA＋租金)
    t.debtEbitdaMax = 4
  }, t => {
    t.debtEbitdaMax = 5
  }) : e.debtCapBasis === `ebitda` ? r(`投資級上限（總債務 ÷ EBITDA）`, `3.5x`, `4.5x`, t => {
    t.debtEbitdaMax = 3.5
  }, t => {
    t.debtEbitdaMax = 4.5
  }) : r(`債務／backlog 上限`, `0.4x`, `1.2x`, t => {
    t.debtBacklog = .4
  }, t => {
    t.debtBacklog = 1.2
  })), r(`每 MW 建置成本`, `+20%`, `−20%`, e => {
    e.capexScale = (e.capexScale ?? 1) * 1.2 // MAG v0.1b：tk 模式每 MW 成本由 Tokenomics 推得，敏感度改乘倍數（等價）
  }, e => {
    e.capexScale = (e.capexScale ?? 1) * .8
  }), r(`GPU 經濟壽命`, `4 年`, `6 年`, e => { // v0.1b：Nebius 會計 5 年、Tokenomics 6 年
    e.gpuLife = 4
  }, e => {
    e.gpuLife = 6
  }), r(`${CALQ.nextFYLabel} 新增 MW`, `1,000`, `0`, e => { // v0.1b：公司 2027 起每年部署 >1 GW
    e.mw31 = 1000
  }, e => {
    e.mw31 = 0
  }), r(`未起租租期`, `${UL.termSens[0]} 年`, `${UL.termSens[1]} 年`, e => { // v0.1b（Oracle）：租期越短年租越高（租金含在 EBITDA 率內，現金中性）
    e.a.newLease = ulPath(UL.termSens[0]), e.ulTerm = UL.termSens[0]
  }, e => {
    e.a.newLease = ulPath(UL.termSens[1]), e.ulTerm = UL.termSens[1]
  }), r(`建設延誤月數`, `12 個月`, `0 個月`, e => { // v0.2：計費 MW 平移、GPU 資本支出照原時程
    e.delayMonths = 12
  }, e => {
    e.delayMonths = 0
  }), r(`未起租租約連動延誤`, `0%`, `100%`, e => { // v0.2：delayLink
    e.delayLink = 0
  }, e => {
    e.delayLink = 1
  }), r(`預付重大財務組成`, hA((e.prepay.financingRate ?? 0) * 100, 2), `0%`, null, e => {
    e.prepay = { ...e.prepay, financingRate: 0 }
  }), r(`新債利率`, `+300bps`, `−300bps`, e => {
    e.m.rate = e.m.rate.map(e => e + 3)
  }, e => {
    e.m.rate = e.m.rate.map(e => Math.max(0, e - 3))
  }), r(`客戶預付比率`, `−5pt`, `+5pt`, e => {
    e.a.customerFund = e.a.customerFund.map(e => Math.max(0, e - .05))
  }, e => {
    e.a.customerFund = e.a.customerFund.map(e => Math.min(.8, e + .05))
  }), r(`Billable MW`, `−15%`, `+15%`, e => {
    e.m.billable = e.m.billable.map(e => e * .85)
  }, e => {
    e.m.billable = e.m.billable.map(e => e * 1.15)
  }), r(`每 MW 年收入`, `−15%`, `+15%`, e => {
    e.revScale = (e.revScale ?? 1) * .85 // MAG v0.1b：priceLink 時每 MW 年收入由 k 推得，敏感度改乘倍數（等價）
  }, e => {
    e.revScale = (e.revScale ?? 1) * 1.15
  }), e.revenueDriver !== `mw` && r(`新產能簽約率`, `−20pt`, `+20pt`, e => {
    e.m.fill = e.m.fill.map(e => Math.max(0, e - 20))
  }, e => {
    e.m.fill = e.m.fill.map(e => Math.min(100, e + 20))
  }), r(`五期認列比例`, `74%`, `94%`, e => {
    e.rp = 74
  }, e => {
    e.rp = 94
  }), r(`債務攤還`, `開`, `關（展期）`, e => {
    e.includeDebt = !0
  }, e => {
    e.includeDebt = !1
  }), r(`股權發行折價`, `30%`, `0%`, e => {
    e.eqDisc = .3
  }, e => {
    e.eqDisc = 0
  }), r(`每年股權吸收上限`, `10%`, `無上限`, e => {
    e.eqCapPct = .1
  }, e => {
    e.eqCapPct = 99
  }), r(`WACC（CAPM β ${COMPANY_DATA.valuation.capm.betaSens[1]}／${COMPANY_DATA.valuation.capm.betaSens[0]}）`, hA(CAPM_Q(v, COMPANY_DATA.valuation.capm.betaSens[1]).wacc * 100, 1), hA(CAPM_Q(v, COMPANY_DATA.valuation.capm.betaSens[0]).wacc * 100, 1), null, null, e => { // v0.1b（Oracle）：β 敏感度
    e.wacc = CAPM_Q(e, COMPANY_DATA.valuation.capm.betaSens[1]).wacc
  }, e => {
    e.wacc = CAPM_Q(e, COMPANY_DATA.valuation.capm.betaSens[0]).wacc
  }), r(`EV/EBITDA 倍數`, `5x`, `7x`, null, null, e => {
    e.evEbitda = 5
  }, e => {
    e.evEbitda = 7
  }), n.sort((e, t) => Math.abs(t.high - t.low) - Math.abs(e.high - e.low)), n
}

function Y(e, t = 1) {
  return Number.isFinite(e) ? e.toLocaleString(`zh-TW`, {
    minimumFractionDigits: t,
    maximumFractionDigits: t
  }) : `—`
}

function mA(e, t = 1) {
  return `${e<0?`−`:``}$${Y(Math.abs(e),t)}`
}

function hA(e, t = 1) {
  return `${Y(e,t)}%`
}
