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
  SC_ACC = COMPANY_DATA.scenarios.accepted,
  SC_BR = COMPANY_DATA.scenarios.billableRatio.billable.map((b, i) => b / COMPANY_DATA.scenarios.billableRatio.accepted[i]),
  SC_LEASE_HI = COMPANY_DATA.scenarios.leaseHighPath,
  SC_MW31 = COMPANY_DATA.scenarios.mw31,
  scA = e => {
    let T = COMPANY_DATA.scenarios.capexTemplate, F = COMPANY_DATA.scenarios.leaseRampFloorMw;
    return {
      costMW: [...T.costMW],
      customerFund: [...T.customerFund],
      newLease: SC_LEASE_HI.map((t, n) => t * Math.max(0, SC_ACC[e][n] - F) / (SC_ACC.high[n] - F)),
      div: [...T.div]
    }
  },
  SCENARIOS = Object.fromEntries([`low`, `base`, `high`].map(k => [k, {
    label: COMPANY_DATA.scenarios.labels[k],
    acc: SC_ACC[k],
    bil: SC_ACC[k].map((e, t) => Math.round(e * SC_BR[t])),
    mw31: SC_MW31[k],
    a: scA(k)
  }])),
  DEBT_TOOLS = COMPANY_DATA.debt.instruments,
  DBT_P = DEBT_TOOLS.reduce((e, t) => e + t[4], 0),
  DBT_I = DEBT_TOOLS.reduce((e, t) => e + t[3] * t[4], 0),
  DBT_R = DBT_I / DBT_P,
  CONV_P = COMPANY_DATA.debt.convertible.principal,
  CONV_I = CONV_P * COMPANY_DATA.debt.convertible.coupon,
  CHECK_TH = COMPANY_DATA.methodology.checks, // 5a：連動檢查門檻（Excel「檢查_連動」同一來源）
  CX_OLD = COMPANY_DATA.legacy.capexV14,
  INT_OLD = COMPANY_DATA.legacy.interestV14,
  LEASE_FACTS = COMPANY_DATA.leases.facts,
  DEFAULTS = Object.fromEntries(Object.entries(structuredClone(COMPANY_DATA.defaults)).flatMap(([k, v]) => k === `intCal` ? [[k, v], [`a`, structuredClone(SCENARIOS.base.a)]] : [[k, v]])),
  Qk = DEFAULTS; // 模板函式庫片段（mid1–mid3）仍以 Qk 引用預設值，保留別名
// ===== W2：每 MW 經濟性（世代組合 × Tokenomics 快照 × 公司專屬輸入；Excel「輸入與假設」B 區後「世代組合」、「每MW經濟性」頁同一邏輯）=====
// 方法開關 company.json → methodology.perMw：capex（tokenomics｜legacy）、cost（bottomUp｜ebitdaPct）、revenue（gpuHr｜legacy）。
// 全為 legacy（ebitdaPct）時與 v4.5 完全相同，畫面不顯示 W2 內容。Tokenomics 值只讀建置時內嵌的快照（COMPANY_DATA.tkSnap），不寫死數字。
var PMWQ = Object.assign({ capex: `legacy`, cost: `ebitdaPct`, revenue: `legacy` }, (COMPANY_DATA.methodology || {}).perMw || {}),
  PMW_ONQ = PMWQ.capex !== `legacy` || PMWQ.cost !== `ebitdaPct` || PMWQ.revenue !== `legacy`,
  FLEETQ = COMPANY_DATA.fleet || null,
  GENQ = FLEETQ ? FLEETQ.generations : [],
  TKSQ = (COMPANY_DATA.tkSnap || {}).items || {},
  MWBASISQ = COMPANY_DATA.meta.mwBasis || `IT`,
  PRICINGQ = COMPANY_DATA.pricing || {},
  AMQ = (COMPANY_DATA.pricing || {}).anchorMultiple || null, // W4：Tokenomics 錨的公司因素（定價倍數 k、隨需占比、證據表）
  CAQ = COMPANY_DATA.companyAdjust || null, // W5：公司實況驗證與公司調整（既有合約 k、新約價格調整、營運成本倍數；證據與文字）
  COSTSQ = COMPANY_DATA.costs || {},
  TK_PHQ = [`IF_MaintIT`, `IF_StaffSW`, `IF_TaxIns`, `IF_DeprLifeIT`], // 可用暫代值的名稱（待 Tokenomics v5.26）；其餘名稱缺少時建置失敗
  COSTMW_LEGQ = [...COMPANY_DATA.scenarios.capexTemplate.costMW]; // v4.5 舊值（對照列）
function tkHasQ(n) { let it = TKSQ[n]; return !!it && !it.missing }
function tkQ(n, g, cs) { // Tokenomics 快照值：成本情境 cs＝低成本／基準／高成本（單值名稱的低／高取 range）
  let it = TKSQ[n]; cs = cs || `基準`;
  if (!it || it.missing) return null;
  if (it.kind === `gen_cost`) return it.values[g] ? it.values[g][cs] : null;
  return cs === `基準` ? it.values : ((it.range || {})[cs.slice(0, 1)] ?? it.values)
}
function mwConvQ(g) { return MWBASISQ === `facility` ? 1 / tkQ(`IF_FacilityGW`, g) : 1 } // 每 IT MW → 每公司 MW（設施口徑：IT MW＝設施 MW ÷ IF_FacilityGW；基準值）
function tkMwQ(n, g, cs) { let x = tkQ(n, g, cs); return x == null ? null : x * mwConvQ(g) } // 每 GW 值 ＝ 每 MW 的百萬美元值（換成公司 MW 口徑）
function wMixQ(mx, f) { return GENQ.reduce((a, g, j) => a + (mx[j] || 0) * ((mx[j] || 0) ? f(g, j) : 0), 0) } // 依世代占比加權（占比 0 的世代不取值）
function tkNeedQ() { // 目前方法需要、快照卻沒有的 Tokenomics 名稱（只列可暫代者；檢查頁警告「成本為暫代值」）
  let need = [];
  PMWQ.capex === `tokenomics` && need.push(`IF_DeprLifeIT`);
  PMWQ.cost === `bottomUp` && need.push(`IF_MaintIT`, `IF_StaffSW`, `IF_TaxIns`);
  return need.filter(n => !tkHasQ(n))
}
var TK_MISSQ = tkNeedQ();
function newMixQ(e) { let A = FLEETQ.newMixAlt; return (e && e.mixAlt && A ? A.mix : FLEETQ.newMix).map(m => GENQ.map(g => m[g] || 0)) }
function capexTkQ(e) { // 每 MW 建置成本（US$m/MW）＝Σ 新增世代占比 × IF_CapexIT（不含廠房：CRWV 機房以租約取得）
  let cs = (e && e.tkCase) || `基準`;
  return newMixQ(e).map(mx => wMixQ(mx, g => tkMwQ(`IF_CapexIT`, g, cs)))
}
function lifeTkQ() { // GPU 經濟壽命：IF_DeprLifeIT 依期初在役世代加權、取整數（v5.26 前暫代＝company.json defaults.gpuLife）
  if (!tkHasQ(`IF_DeprLifeIT`)) return COMPANY_DATA.defaults.gpuLife;
  return Math.round(GENQ.reduce((a, g) => a + (FLEETQ.openMix.mix[g] || 0) * tkQ(`IF_DeprLifeIT`, g), 0))
}
function sgaPctQ(e) { // 公司管銷率（扣 D&A、SBC）＝（銷售行銷 − SBC ＋ 一般管理 − SBC）÷ 營收，最近一季實際，各期持平 [Derived]
  let q = LATEST_Q, b = (e && e.sgaBasis) || COSTSQ.sgaBasis || `exSbc`;
  return b === `gaap` ? (q.sm + q.ga) / q.revenue : (q.sm - q.smSbc + q.ga - q.gaSbc) / q.revenue
}
if (PMWQ.capex === `tokenomics`) { // 建置成本與壽命改為 Tokenomics 推導（輸入格仍可手動覆寫）
  let c = capexTkQ(null);
  for (let k of [`low`, `base`, `high`]) SCENARIOS[k].a.costMW = [...c];
  DEFAULTS.a.costMW = [...c];
  DEFAULTS.gpuLife = lifeTkQ()
}
AMQ && Object.assign(DEFAULTS, { kLong: AMQ.long.base, kSpot: AMQ.spot.base, odShare: AMQ.onDemandShare.base }); // W4 r2：定價倍數 k 與隨需占比（敏感度以 e 覆寫）
CAQ && Object.assign(DEFAULTS, { kExOn: CAQ.existingK.on, kNewAdj: CAQ.newK.adjBase, opexScale: CAQ.opexScale.base }); // W5：既有合約 k 開關、新約價格調整、營運成本倍數（敏感度以 e 覆寫）
function exFleetQ(F) { // W5：既有合約 MW＝最新季末在役（fleet.openMix.activeMW）逐期扣汰換（汰換由最舊世代先出；Excel「每MW經濟性」W4 區同列）
  let s = FLEETQ.openMix.activeMW, st = [], en = [];
  for (let n = 0; n < 5; n++) { let x = Math.max(0, s - F.ret[n]); st.push(s), en.push(x), s = x }
  return { st, en, avg: st.map((x, n) => (x + en[n]) / 2) }
}
function kExQ(e, t, cs) { // W5：k_既有＝Q2 實現單價 ÷ Q2 錨＝（Q2 每在役 MW 年收入 − 服務）÷（計費比例 × 利用率 × 季末在役世代 × IF_HoldEcon）
  let lq = LATEST_Q, oMW = FLEETQ.openMix.activeMW, qMW = oMW - CALL_FACTS.activeAddQ2 / 2,
    rq = lq.revenue * 4 / qMW * 1e3, oq = e.services[0] / PERIOD_YEARS[0] / qMW * 1e3, bq = e.billableOpen / oMW, uq = t.util[0] / 100,
    aq = GENQ.reduce((a, g) => a + (FLEETQ.openMix.mix[g] || 0) * tkMwQ(`IF_HoldEcon`, g, cs), 0);
  return { rq, oq, bq, uq, aq, k: (rq - oq) / (bq * uq * aq) }
}
PMWQ.cost === `bottomUp` && (DEFAULTS.ebSteady = null); // 由下而上：穩態 EBITDA 率預設＝由下而上的 FY30 值（null）；輸入數值時視為 FY30 目標，差額線性分攤

function ebPathQ(e, d) { return PMWQ.cost === `bottomUp` ? [d.years[0].ebM, d.years[4].ebM] : [e.ebStart, e.ebSteady] } // W2：畫面上的 EBITDA 率起點／FY30（由下而上時取模型路徑）
function fleetQ(e, acc) { // 世代組合：期初（最新季末）在役機隊 → 各期期末／平均在役 MW（依世代）；汰換由最舊世代先出、以當期新增世代補回
  if (!FLEETQ) return null;
  let O = FLEETQ.openMix, NM = newMixQ(e),
    VIN = Object.fromEntries(Object.entries(e.mwYearEnd).map(([y, m]) => [y, m - (e.mwYearEnd[y - 1] ?? 0)])),
    prev = GENQ.map(g => O.activeMW * (O.mix[g] || 0)), start = [], end = [], avg = [], add = [], ret = [];
  for (let n = 0; n < 5; n++) {
    let pt = prev.reduce((a, b) => a + b, 0), rt = Math.min(VIN[PERIOD_FY[n] - e.gpuLife] || 0, pt), ad = Math.max(0, acc[n] - pt), left = rt,
      cur = prev.map(x => { let k = Math.min(x, left); return left -= k, x - k });
    cur = cur.map((x, j) => x + NM[n][j] * (ad + rt));
    start.push(prev), end.push(cur), avg.push(cur.map((x, j) => (prev[j] + x) / 2)), add.push(ad), ret.push(rt), prev = cur
  }
  let tot = avg.map(a => a.reduce((x, y) => x + y, 0));
  return { start, end, avg, tot, add, ret, nm: NM, endTot: end.map(a => a.reduce((x, y) => x + y, 0)), mix: avg.map((a, n) => a.map(x => x / Math.max(tot[n], 1e-9))) }
}
function gpuPerMwQ(mx) { return wMixQ(mx, g => tkMwQ(`IF_GPUsPerGW`, g) / 1e3) } // 每公司 MW GPU 數（依世代加權）
function revGpuQ(e, F) { // 每 MW 年收入（US$bn/MW，100% 計費時數；利用率在收入端另乘）＝Σ 平均在役占比 × 每 MW GPU 數 × GPU 小時合約價 × 8,760 ÷ 10⁹
  let P = PRICINGQ.gpuHr || {}, k = (e && e.pxCase) || `base`;
  return F.mix.map(mx => wMixQ(mx, g => P[g] && P[g][k] != null ? tkMwQ(`IF_GPUsPerGW`, g) / 1e3 * P[g][k] * 8760 / 1e9 : NaN))
}
function anchorRevQ(e, F, t) { // W4 r2：每 MW 年收入＝Tokenomics 錨（Σ 平均在役占比 × IF_HoldEcon）× 定價倍數 k（Excel「每MW經濟性」W4 區同列同算式）
  // k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約；k_長約、k_現貨 為基準成本情境的值，其他情境以基準證據世代的成本比例重算（價格是事實）；RPO 覆蓋率只作對照
  let cs = e.tkCase || `基準`, i = e.rp / 100, ref = side => AMQ.evidence.find(x => x.label === AMQ[side].refEvidence),
    rr = x => tkQ(x.tkName, x.gen, `基準`) / tkQ(x.tkName, x.gen, cs), rL = ref(`long`), rS = ref(`spot`),
    kL0 = e.kLong ?? AMQ.long.base, kS0 = e.kSpot ?? AMQ.spot.base, kL = kL0 * rr(rL), kS = kS0 * rr(rS), od = e.odShare ?? AMQ.onDemandShare.base,
    R = { anchor: [], sched: [], avgB: [], util: [], longCap: [], cover: [], kL: [], kS: [], od: [], k: [], rev: [], kN: [], exS: [] },
    W5 = !!CAQ, X = W5 ? exFleetQ(F) : null, KE = W5 ? kExQ(e, t, cs) : null, on = W5 ? (e.kExOn ?? CAQ.existingK.on) : 0, adj = W5 ? (e.kNewAdj ?? CAQ.newK.adjBase) : 0;
  R.ex = X, R.ke = KE, R.adj = adj, R.on = on;
  for (let r = 0; r < 5; r++) {
    let L = PERIOD_YEARS[r], A = wMixQ(F.mix[r], g => tkMwQ(`IF_HoldEcon`, g, cs)),
      o = e.rpoOpen * (RPO_BUCKET_W[r] / RPO_SCHEDULED_SHARE) * i + e.rpoPendingAdd * RPO_Q3ADD_W[r] * i,
      c = r === 0 ? e.billableOpen : t.billable[r - 1], l = e.useAvgMw ? (c + t.billable[r]) / 2 : t.billable[r],
      lc = l * A / 1e3 * kL0 * (t.util[r] / 100) * L, kN = od * kS + (1 - od) * kL * (1 + adj), // W5：新約 k（新約價格調整只作用於長約部分）
      xs = W5 ? on * X.avg[r] / Math.max(1e-9, F.tot[r]) : 0, k = W5 ? xs * KE.k + (1 - xs) * kN : od * kS + (1 - od) * kL; // W5：k＝既有占比 × k_既有＋（1 − 既有占比）× k_新約
    R.anchor.push(A), R.sched.push(o), R.avgB.push(l), R.util.push(t.util[r] / 100), R.longCap.push(lc), R.cover.push(o / Math.max(1e-9, lc)), R.kL.push(kL), R.kS.push(kS), R.od.push(od), R.k.push(k), R.rev.push(A * k), R.kN.push(kN), R.exS.push(xs)
  }
  return R
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
    name: `其他／未列名站點（51 站中未具名者）`,
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
  e.linkSites && (t.accepted[0] = Math.max(t.accepted[0], n.accepted), t.billable[0] = Math.max(t.billable[0], n.billable));
  for (let e = 0; e < 5; e++) t.accepted[e] = Math.max(0, t.accepted[e]), e > 0 && (t.accepted[e] = Math.max(t.accepted[e], t.accepted[e - 1])), t.billable[e] = Math.min(Math.max(0, t.billable[e]), t.accepted[e]), t.util[e] = Math.min(100, Math.max(0, t.util[e])), t.aiShare[e] = Math.min(100, Math.max(0, t.aiShare[e])), t.fill[e] = Math.min(100, Math.max(0, t.fill[e])), t.defaultP[e] = Math.min(100, Math.max(0, t.defaultP[e])), t.recovery[e] = Math.min(100, Math.max(0, t.recovery[e]));
  return t
}

function rA(e) {
  return (LEASE_CASH_ON_BAL[4] + e.a.newLease[4]) * e.terminal.residualLeaseYears
}

function iA(e) {
  return e.a.newLease.reduce((e, t) => e + t, 0) + e.a.newLease[4] * e.terminal.residualLeaseYears
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
    n = e.a.newLease[4] * e.terminal.residualLeaseYears,
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

function revPassQ(e, t, r) { // W2：與 runFunding 同一組收入算式（由下而上的成本率需要各期營收；算式順序相同，數值逐位相同）
  let i = e.rp / 100, L = PERIOD_YEARS[r],
    o = e.rpoOpen * (RPO_BUCKET_W[r] / RPO_SCHEDULED_SHARE) * i + e.rpoPendingAdd * RPO_Q3ADD_W[r] * i,
    c = r === 0 ? e.billableOpen : t.billable[r - 1],
    l = e.useAvgMw ? (c + t.billable[r]) / 2 : t.billable[r],
    u = l * t.revMW[r] * (e.revScale ?? 1) * (t.util[r] / 100) * L,
    d = Math.max(0, o - u), f = o - d, nR = Math.max(0, u - o) * (t.fill[r] / 100);
  return f + nR + e.services[r]
}

function buCostQ(e, t, F, r, rev, S) { // W2：由下而上營運成本（租金前；US$bn）——電費、IT 維護、人員軟體、稅險依平均在役世代加權 × 平均在役 MW；管銷＝營收 × 管銷率
  let L = PERIOD_YEARS[r], M = F.tot[r], mx = F.mix[r], cs = e.tkCase || `基準`,
    pw = wMixQ(mx, g => tkMwQ(`IF_PowerCost`, g, cs)),
    mt = tkHasQ(`IF_MaintIT`) ? wMixQ(mx, g => tkMwQ(`IF_MaintIT`, g, cs)) : t.maint[r],
    st = tkHasQ(`IF_StaffSW`) ? wMixQ(mx, g => tkMwQ(`IF_StaffSW`, g, cs)) : 0,
    tx = tkHasQ(`IF_TaxIns`) ? wMixQ(mx, g => tkMwQ(`IF_TaxIns`, g, cs) * tkQ(`IF_CapexIT`, g, cs) / tkQ(`IF_CapexTotal`, g, cs)) : 0,
    k = M / 1e3 * L, R = Math.max(rev, .01), sp = sgaPctQ(e), sga = rev * sp;
  if (CAQ) { let z = e.opexScale ?? CAQ.opexScale.base; pw *= z, mt *= z, st *= z, tx *= z } // W5：營運成本倍數（基準 1；公司實況敏感度）
  let cash = (pw + mt + st + tx) * k + sga;
  return { mw: M, pwMW: pw, mtMW: mt, stMW: st, txMW: tx, sgaPct: sp, power: pw * k, maint: mt * k, staff: st * k, tax: tx * k, sga, cash, rev, rent: S, ebr: 1 - cash / R, eb: 1 - cash / R - S / R }
}

// W3（v4.6）：一頁摘要「每 MW 一句」——首個完整財年的每 MW 年收入、現金成本（含租金）、EBITDA，與 Tokenomics 經濟持有成本的關係。
// 數字取自 perMwQ（與 Excel「每MW經濟性」頁同列；Excel「摘要」→「結論｜每 MW 經濟性句」同一句，cmp31 逐字比對）；只在每 MW 使用新方法時顯示。
function pmLineQ(d, e) {
  if (!PMW_ONQ || !FLEETQ) return null;
  let P = perMwQ(d, e);
  if (!P) return null;
  let g = (sec, k) => P[sec].find(r => r[0] === k)[2][1], R = g(`rev`, `每 MW 年收入 ÷ 經濟持有成本`);
  return `每 MW（${PERIODS[1]}）：年收入 $${Y(g(`sum`, `每 MW 年收入（算力＋服務）`), 1)}m、現金成本 $${Y(g(`sum`, `現金成本合計（含租金）`), 1)}m（含租金 $${Y(g(`sum`, `租金`), 1)}m）、EBITDA $${Y(g(`sum`, `EBITDA`), 1)}m；`
    + `計費單價為 Tokenomics 經濟持有成本（含廠房資本回收的打平線）$${Y(g(`rev`, `每 MW 經濟持有成本（不賠錢下限）`), 1)}m 的 ${Y(R, 2)} 倍${R < 1 ? `，未回收全部持有成本` : ``}。`
}

function perMwQ(d, e) { // W2：每 MW 經濟性（與 Excel「每MW經濟性」頁同列名、同算式；cmp31 逐列比對）。回傳 { 區段: [[列名, 單位, 5 期值, 說明]] }
  let F = d.fleet;
  if (!F) return null;
  let y = d.years, L = PERIOD_YEARS, cs = e.tkCase || `基準`, I5 = [0, 1, 2, 3, 4], G = GENQ, NG = G.length,
    M = F.tot, mx = F.mix, nm = F.nm,
    pm = (x, i) => x / Math.max(1e-9, M[i]) / L[i] * 1e3,
    fac = mx.map(m => wMixQ(m, g => tkQ(`IF_FacilityGW`, g))),
    itMw = I5.map(i => MWBASISQ === `facility` ? M[i] / fac[i] : M[i]),
    gpu = mx.map(m => gpuPerMwQ(m)),
    gEcon = mx.map((m, i) => wMixQ(m, g => tkMwQ(`IF_GPUsPerGW`, g) / 1e3 * tkQ(`IF_GPUhrEcon`, g, cs)) / Math.max(1e-9, gpu[i])),
    hold = mx.map(m => wMixQ(m, g => tkMwQ(`IF_HoldEcon`, g, cs))),
    revIn = I5.map(i => d.m.revMW[i] * (e.revScale ?? 1) * 1e3),
    px = PRICINGQ.gpuHr || {}, hasPx = G.every(g => px[g] && px[g].base != null),
    gpuRev = hasPx ? revGpuQ(e, F).map(x => x * 1e3) : I5.map(() => `不適用`),
    imp = I5.map(i => revIn[i] * 1e6 / Math.max(1e-9, gpu[i] * 8760)),
    capIT = nm.map(m => wMixQ(m, g => tkMwQ(`IF_CapexIT`, g, cs))),
    capTot = nm.map(m => wMixQ(m, g => tkMwQ(`IF_CapexTotal`, g, cs))),
    buM = (k) => y.map(t => t.bu[k]),
    lq = LATEST_Q, oMW = FLEETQ.openMix.activeMW, qMW = oMW - CALL_FACTS.activeAddQ2 / 2,
    qCost = lq.costRev + lq.techInfra - lq.da - lq.sbcCostTi - lq.opLeaseCost - lq.varLeaseCost,
    qPM = qCost * 4 / qMW * 1e3, buPM = I5.map(i => y[i].bu.pwMW + y[i].bu.mtMW + y[i].bu.stMW + y[i].bu.txMW),
    NA = [null, null, null, null],
    sum = [
      [`平均在役 MW（IT 關鍵電力）`, `MW`, itMw],
      [`平均在役 MW（設施＝IT × IF_FacilityGW）`, `MW`, I5.map(i => itMw[i] * fac[i])],
      [`每 MW 年收入（算力＋服務）`, `US$m/MW`, I5.map(i => pm(y[i].totRev, i))],
      [`　其中：算力收入`, `US$m/MW`, I5.map(i => pm(y[i].isRev, i))],
      [`電費`, `US$m/MW`, I5.map(i => pm(y[i].bu.power, i))],
      [`IT 維護`, `US$m/MW`, I5.map(i => pm(y[i].bu.maint, i))],
      [`人員、軟體、水與耗材`, `US$m/MW`, I5.map(i => pm(y[i].bu.staff, i))],
      [`財產稅與保險`, `US$m/MW`, I5.map(i => pm(y[i].bu.tax, i))],
      [`公司管銷與其他`, `US$m/MW`, I5.map(i => pm(y[i].bu.sga, i))],
      [`租金`, `US$m/MW`, I5.map(i => pm(y[i].lease, i))],
      [`現金成本合計（含租金）`, `US$m/MW`, I5.map(i => pm(y[i].totRev - y[i].ebitdaPL, i))],
      [`EBITDA`, `US$m/MW`, I5.map(i => pm(y[i].ebitdaPL, i))],
      [`D&A（模型車隊折舊）`, `US$m/MW`, I5.map(i => pm(y[i].daFleet, i))],
      [`利息`, `US$m/MW`, I5.map(i => pm(y[i].interest, i))],
      [`稅前`, `US$m/MW`, I5.map(i => pm(y[i].ebitdaPL, i) - pm(y[i].daFleet, i) - pm(y[i].interest, i))],
      [`每 MW 資本支出（新增 MW 的建置成本）`, `US$m/MW`, I5.map(i => e.a.costMW[i] * (e.capexScale ?? 1))],
      [`每設施 MW 年收入`, `US$m/MW`, I5.map(i => pm(y[i].totRev, i) * itMw[i] / (itMw[i] * fac[i]))],
      [`每設施 MW EBITDA`, `US$m/MW`, I5.map(i => pm(y[i].ebitdaPL, i) * itMw[i] / (itMw[i] * fac[i]))]
    ],
    fleet = [
      [`期初在役 MW 合計`, `MW`, F.start.map(a => a.reduce((x, z) => x + z, 0))],
      ...G.map((g, j) => [`期初在役 MW｜${g}`, `MW`, F.start.map(a => a[j])]),
      [`汰換 MW（壽命到期批次，最舊世代先出）`, `MW`, F.ret],
      [`新增 MW（期末 Accepted − 期初，不為負）`, `MW`, F.add],
      ...G.map((g, j) => [`期末在役 MW｜${g}`, `MW`, F.end.map(a => a[j])]),
      [`期末在役 MW 合計`, `MW`, F.endTot],
      ...G.map((g, j) => [`平均在役 MW｜${g}`, `MW`, F.avg.map(a => a[j])]),
      [`平均在役 MW 合計`, `MW`, M],
      ...G.map((g, j) => [`平均在役占比｜${g}`, `%`, mx.map(a => a[j])])
    ],
    bu = [
      [`每 MW 電費（世代加權）`, `US$m/MW`, buM(`pwMW`)], [`每 MW IT 維護（世代加權）`, `US$m/MW`, buM(`mtMW`)],
      [`每 MW 人員、軟體、水與耗材`, `US$m/MW`, buM(`stMW`)], [`每 MW 財產稅與保險（只算 IT 部分）`, `US$m/MW`, buM(`txMW`)],
      [`電費（金額）`, `US$bn`, buM(`power`)], [`IT 維護（金額）`, `US$bn`, buM(`maint`)], [`人員、軟體、水與耗材（金額）`, `US$bn`, buM(`staff`)], [`財產稅與保險（金額）`, `US$bn`, buM(`tax`)],
      [`模型期總營收（算力＋服務）`, `US$bn`, buM(`rev`)], [`管銷率`, `%`, buM(`sgaPct`)], [`公司管銷與其他（金額）`, `US$bn`, buM(`sga`)],
      [`現金營運成本合計（租金前）`, `US$bn`, buM(`cash`)], [`由下而上 EBITDAR 率`, `%`, buM(`ebr`)], [`租金合計`, `US$bn`, buM(`rent`)],
      [`由下而上 EBITDA 率`, `%`, buM(`eb`)], [`EBITDA 率（模型採用）`, `%`, y.map(t => t.ebM)]
    ],
    rev = [
      [`每 MW 年收入（模型採用，100% 計費時數）`, `US$m/MW`, revIn], [`每 MW 年收入（計費後＝× 利用率）`, `US$m/MW`, I5.map(i => revIn[i] * d.m.util[i] / 100)],
      [PMWQ.revenue === `tkAnchor` ? `每 MW 年收入（v4.6 舊輸入 m.revMW，對照）` : `每 MW 年收入（v4.5 舊值，對照）`, `US$m/MW`, COMPANY_DATA.defaults.m.revMW.map(x => x * 1e3)],
      [`期末 ARR 指引 ÷ 年底主動電力（公司數字，只作對照）`, `US$m/MW`, [(CALL_FACTS.arrLo + CALL_FACTS.arrHi) / 2 / CALL_FACTS.yeActiveGw, ...NA]],
      [`每 MW GPU 數（世代加權）`, `顆/MW`, gpu], [`GPU 小時價格路線：每 MW 年收入`, `US$m/MW`, gpuRev],
      [`隱含每 GPU 小時價格（反算對照，不是輸入）`, `US$/GPU-hr`, imp], [`每 GPU 小時經濟持有成本（GPU 數加權）`, `US$/GPU-hr`, gEcon],
      [`隱含價格 ÷ GPU 小時持有成本`, `倍`, I5.map(i => imp[i] / Math.max(1e-9, gEcon[i]))],
      [`每 MW 經濟持有成本（不賠錢下限）`, `US$m/MW`, hold], [`每 MW 年收入 ÷ 經濟持有成本`, `倍`, I5.map(i => revIn[i] / Math.max(1e-9, hold[i]))],
      ...(PRICINGQ.peerRevPerMw || []).map(p => [`同業｜${p.label}`, p.unit, I5.map(() => p.value)]),
      ...(PRICINGQ.marketRefs || []).map(p => [`市場價格｜${p.label}`, `US$/GPU-hr`, [p.value, ...NA]])
    ],
    cap = [
      [`每 MW 建置成本（模型採用）`, `US$m/MW`, [...e.a.costMW]], [`Tokenomics IT 設備（IF_CapexIT）`, `US$m/MW`, capIT],
      [`Tokenomics 含廠房合計（IF_CapexTotal；對照）`, `US$m/MW`, capTot], [`v4.5 舊值（以 FY26 CapEx 指引校準）`, `US$m/MW`, [...COSTMW_LEGQ]],
      [`每 MW 建置成本 ÷ 壽命（在役世代加權；D&A 參考）`, `US$m/MW`, mx.map(m => wMixQ(m, g => tkMwQ(`IF_CapexIT`, g, cs)) / e.gpuLife)]
    ],
    q2 = [
      [`最近一季平均在役 MW`, `MW`, [qMW, ...NA]], [`最近一季現金營運成本（租金前，不含管銷）`, `US$bn`, [qCost, ...NA]],
      [`最近一季每 MW 現金營運成本（年化，不含管銷）`, `US$m/MW`, [qPM, ...NA]], [`模型由下而上每 MW 現金營運成本（不含管銷）`, `US$m/MW`, buPM],
      [`差距（模型 ${PERIODS[0]} − 實際）`, `US$m/MW`, [buPM[0] - qPM, ...NA]],
      [`最近一季每 MW 年租金（營業＋變動，年化）`, `US$m/MW`, [(lq.opLeaseCost + lq.varLeaseCost) * 4 / qMW * 1e3, ...NA]],
      [`模型每 MW 年租金`, `US$m/MW`, I5.map(i => pm(y[i].lease, i))],
      ...(AMQ && PMWQ.revenue === `tkAnchor` ? (() => { // W4：Q2 收入對照（只作驗證，不校準 k）
        let qr = lq.revenue * 4 / qMW * 1e3, br = e.billableOpen / oMW, qa = GENQ.reduce((a, g) => a + (FLEETQ.openMix.mix[g] || 0) * tkMwQ(`IF_HoldEcon`, g, cs), 0);
        return [[`最近一季每 MW 年收入（營收 × 4 ÷ 平均在役 MW）`, `US$m/MW`, [qr, ...NA]], [`最近一季計費比例（季末 Billable ÷ 季末在役 MW，假設）`, `%`, [br, ...NA]],
          [`最近一季每 MW 年收入（÷ 平均計費 MW）`, `US$m/MW`, [qr / br, ...NA]], [`最近一季錨（季末在役世代 × IF_HoldEcon）`, `US$m/MW`, [qa, ...NA]],
          [`最近一季隱含 k（年化營收 ÷ 平均在役 MW ÷ 錨，未調整）`, `倍`, [qr / qa, ...NA]]]
      })() : [])
    ];
  let tk = null, kev = null;
  if (AMQ) { // W4：Tokenomics 錨 × k（Excel「每MW經濟性」W4 區）與 k 證據表
    let A = anchorRevQ(e, F, d.m), old = COMPANY_DATA.defaults.m.revMW.map(x => x * 1e3), hasRF = tkHasQ(`IF_RevGWFleet`),
      rf = hasRF ? mx.map(m => wMixQ(m, g => tkMwQ(`IF_RevGWFleet`, g, cs))) : I5.map(() => `不適用`),
      ratio = hasRF ? I5.map(i => revIn[i] * d.m.util[i] / 100 / Math.max(1e-9, rf[i])) : I5.map(() => `不適用`),
      th = CHECK_TH.revCapShareMax;
    tk = [
      [`錨｜每 MW 經濟持有成本（IF_HoldEcon，在役世代加權）`, `US$m/MW`, A.anchor], [`錨｜排程 RPO（模型期）`, `US$bn`, A.sched],
      [`錨｜平均計費 MW`, `MW`, A.avgB], [`錨｜利用率`, `%`, A.util],
      [`長約價產能收入（平均計費 MW × 錨 × k_長約 × 利用率 × 期間長度）`, `US$bn`, A.longCap],
      [`RPO 涵蓋的產能 ÷ 在役計費產能（對照，不驅動）`, `%`, A.cover],
      [`k_長約（依目前成本情境重算）`, `倍`, A.kL], [`k_現貨（依目前成本情境重算）`, `倍`, A.kS], [`隨需占比（輸入）`, `%`, A.od],
      ...(CAQ ? [ // W5：既有合約（最新季末在役）以 Q2 實現單價為 k_既有；新增與汰換補回的 MW 按 k_新約
        [`新約價格調整（輸入；只作用於長約部分）`, `%`, I5.map(() => A.adj)],
        [`k_新約（隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約 ×（1＋新約價格調整））`, `倍`, A.kN],
        [`既有合約 MW｜期初（最新季末在役，逐期扣汰換）`, `MW`, A.ex.st], [`既有合約 MW｜期末`, `MW`, A.ex.en],
        [`既有合約占比（平均既有 MW ÷ 平均在役 MW × 開關）`, `%`, A.exS],
        [`k_既有｜Q2 每 MW 年收入（營收 × 4 ÷ Q2 平均在役 MW）`, `US$m/MW`, [A.ke.rq, ...NA]], [`k_既有｜Q2 服務收入（首期服務年化 ÷ Q2 平均在役 MW）`, `US$m/MW`, [A.ke.oq, ...NA]],
        [`k_既有｜Q2 計費比例（季末 Billable ÷ 季末在役，假設）`, `%`, [A.ke.bq, ...NA]], [`k_既有｜Q2 利用率（＝首期利用率，假設）`, `%`, [A.ke.uq, ...NA]],
        [`k_既有｜Q2 錨（季末在役世代 × IF_HoldEcon，目前成本情境）`, `US$m/MW`, [A.ke.aq, ...NA]],
        [`k_既有（Q2 實現單價 ÷ Q2 錨；合約期內固定）`, `倍`, I5.map(() => A.ke.k)], [`既有合約 k 開關（1＝套用、0＝不套用）`, ``, I5.map(() => A.on)],
        [`定價倍數 k（既有占比 × k_既有 ＋（1 − 既有占比）× k_新約）`, `倍`, A.k]] :
        [[`定價倍數 k（隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約）`, `倍`, A.k]]),
      [`Tokenomics 錨 × k：每 MW 年收入（100% 計費時數）`, `US$m/MW`, A.rev],
      [`錨 × k 相對 v4.6 舊輸入 m.revMW`, `%`, I5.map(i => A.rev[i] / Math.max(1e-9, old[i]) - 1)],
      [`上限檢查｜客戶每 MW 付費 token 營收（IF_RevGWFleet，在役世代加權）`, `US$m/MW`, rf],
      [`上限檢查｜CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收`, `%`, ratio],
      [`上限檢查｜各期結果`, ``, hasRF ? ratio.map(x => x > th ? `警示` : `通過`) : I5.map(() => `不適用`)]
    ];
    kev = [
      ...AMQ.evidence.map(x => { let v = tkQ(x.tkName, x.gen); return [`k 證據｜${x.label}`, x.unit, v == null ? [x.price, `不適用`, `不適用`, x.use, null] : [x.price, v, x.price / v, x.use, null]] }),
      ...(AMQ.contractMix || []).map(x => [`合約組合｜${x.label}`, x.unit, [x.value ?? x.tag, null, null, null, null]])
    ]
  }
  let cv = null, q2r = null;
  if (CAQ && AMQ && tk && PMWQ.revenue === `tkAnchor`) { // W5：公司實況驗證（Excel「公司實況驗證」頁同列名、同算式；欄＝Tokenomics 值、CRWV 實際、差距、採用值、處理規則）
    let A = anchorRevQ(e, F, d.m), m0 = mx[0], tw = n => wMixQ(m0, g => tkMwQ(n, g, cs)),
      tPw = tw(`IF_PowerCost`), tMt = tw(`IF_MaintIT`), tSt = tw(`IF_StaffSW`), tTx = wMixQ(m0, g => tkMwQ(`IF_TaxIns`, g, cs) * tkQ(`IF_CapexIT`, g, cs) / tkQ(`IF_CapexTotal`, g, cs)),
      qAct = (lq.costRev + lq.techInfra - lq.da - lq.sbcCostTi - lq.opLeaseCost) * 4 / qMW * 1e3, qRent = lq.opLeaseCost * 4 / qMW * 1e3,
      CX = CAQ.capexActual, dMW = CX.mwEnd - CX.mwStart, cxAct = lq.capexH1 / dMW * 1e3, cxTech = (CX.techEquip[1] - CX.techEquip[0]) / dMW * 1e3,
      lifeTk = GENQ.reduce((a, g) => a + (FLEETQ.openMix.mix[g] || 0) * tkQ(`IF_DeprLifeIT`, g), 0), P = CAQ.params, tol = CAQ.gapTol,
      gp = (x, y) => typeof x === `number` && typeof y === `number` ? y / Math.max(1e-9, x) - 1 : `不適用`,
      rl = (key, g) => { let q = P.find(z => z.key === key).rule; return q === `4a` || typeof g !== `number` ? q : Math.abs(g) <= tol ? `1` : q },
      row = (key, lab, u, c, dv, f) => { let g = gp(c, dv); return [lab, u, [c, dv, g, f, rl(key, g)]] },
      L0 = P.reduce((o, z) => (o[z.key] = z.label, o), {}), NAx = `不適用`, rp = x => pm(x, 0);
    cv = [
      row(`kExist`, L0.kExist, `倍`, A.kL[0], A.ke.k, A.on ? A.ke.k : A.kN[0]),
      row(`kNew`, L0.kNew, `倍`, A.kL[0], NAx, A.kN[0]),
      row(`odShare`, L0.odShare, `%`, AMQ.onDemandShare.base, NAx, A.od[0]),
      row(`power`, L0.power, `US$m/MW`, tPw, NAx, y[0].bu.pwMW), row(`maint`, L0.maint, `US$m/MW`, tMt, NAx, y[0].bu.mtMW),
      row(`staff`, L0.staff, `US$m/MW`, tSt, NAx, y[0].bu.stMW), row(`taxIns`, L0.taxIns, `US$m/MW`, tTx, NAx, y[0].bu.txMW),
      row(`opexBundle`, L0.opexBundle, `US$m/MW`, tPw + tMt + tSt + tTx, qAct, buPM[0]),
      row(`rent`, L0.rent, `US$m/MW`, rp(y[0].lease), qRent, rp(y[0].lease)),
      row(`capex`, L0.capex, `US$m/MW`, capIT[0], cxAct, e.a.costMW[0] * (e.capexScale ?? 1)),
      [`每 MW 資本支出｜技術設備毛額增加 ÷ 新增 MW（對照）`, `US$m/MW`, [capIT[0], cxTech, gp(capIT[0], cxTech), null, null]],
      row(`life`, L0.life, `年`, lifeTk, P.find(z => z.key === `life`).actualValue, e.gpuLife),
      ...(() => { let Z = cvSensQ(); return [
        [`敏感度輸入｜Q2 季末世代 Tokenomics 營運成本合計（基準成本情境）`, `US$m/MW`, [Z.tb, null, null, null, null]],
        [`敏感度輸入｜營運成本倍數＝Q2 實際 ÷ Q2 季末世代 Tokenomics 合計`, `倍`, [Z.opex, null, null, null, null]],
        [`敏感度輸入｜首期新增世代 Tokenomics IT 資本（基準成本情境）`, `US$m/MW`, [Z.cxTk, null, null, null, null]],
        [`敏感度輸入｜每 MW 建置成本倍數＝${CALQ.ytdLabel} 實際 ÷ 首期新增世代 Tokenomics`, `倍`, [Z.capex, null, null, null, null]],
        [`敏感度輸入｜CRWV 短天期 k（短天期合約價 ÷ 同世代 Tokenomics，基準）`, `倍`, [Z.kcw, null, null, null, null]]] })()
    ];
    let qSga = (lq.sm - lq.smSbc + lq.ga - lq.gaSbc) * 4 / qMW * 1e3, qEb = lq.adjEbitda * 4 / qMW * 1e3, r2 = (lab, u, c, dv) => [lab, u, [c, dv, gp(c, dv), null, null]];
    q2r = [
      r2(`Q2 對帳｜每 MW 年收入（算力＋服務）`, `US$m/MW`, rp(y[0].totRev), A.ke.rq),
      r2(`Q2 對帳｜每 MW 營運成本（租金前、不含管銷；Q2 含變動租賃）`, `US$m/MW`, buPM[0], qAct),
      r2(`Q2 對帳｜每 MW 固定租金`, `US$m/MW`, rp(y[0].lease), qRent),
      r2(`Q2 對帳｜每 MW 管銷（扣 SBC）`, `US$m/MW`, rp(y[0].bu.sga), qSga),
      r2(`Q2 對帳｜每 MW EBITDA（Q2＝調整後 EBITDA）`, `US$m/MW`, rp(y[0].ebitdaPL), qEb),
      [`Q2 對帳｜EBITDA 率（差距＝百分點）`, `%`, [y[0].ebM, lq.adjEbitda / lq.revenue, y[0].ebM - lq.adjEbitda / lq.revenue, null, null]]
    ]
  }
  return { sum, fleet, bu, rev, cap, q2, tk, kev, cv, q2r }
}

function cvSensQ() { // W5：公司實況敏感度的輸入值（與情境無關：Q2 季末世代與首期新增世代、基準成本情境；Excel「公司實況驗證」頁敏感度輸入列同算式）
  let lq = LATEST_Q, O = FLEETQ.openMix, qMW = O.activeMW - CALL_FACTS.activeAddQ2 / 2, CX = CAQ.capexActual,
    om = g => O.mix[g] || 0, tb = GENQ.reduce((a, g) => a + om(g) * (tkMwQ(`IF_PowerCost`, g) + tkMwQ(`IF_MaintIT`, g) + tkMwQ(`IF_StaffSW`, g) + tkMwQ(`IF_TaxIns`, g) * tkQ(`IF_CapexIT`, g) / tkQ(`IF_CapexTotal`, g)), 0),
    qAct = (lq.costRev + lq.techInfra - lq.da - lq.sbcCostTi - lq.opLeaseCost) * 4 / qMW * 1e3, cxAct = lq.capexH1 / (CX.mwEnd - CX.mwStart) * 1e3,
    cxTk = wMixQ(newMixQ(null)[0], g => tkMwQ(`IF_CapexIT`, g)), ev = AMQ.evidence.find(x => x.label === CAQ.spotCw.evidenceLabel);
  return { opex: qAct / tb, capex: cxAct / cxTk, kcw: ev.price / tkQ(ev.tkName, ev.gen), tb, cxTk }
}
function pmwSensQ(e, v) { // W2：每 MW 敏感度（與 scripts/permw_sens.py 同一組設定；Excel 為建置時快照，cmp31 逐格比對）
  let C = [[`base`, `基準（目前輸入）`, {}], [`tkLow`, `Tokenomics 低成本`, { tkCase: `低成本` }], [`tkHigh`, `Tokenomics 高成本`, { tkCase: `高成本` }],
    [`pxLow`, `GPU 小時價格 低`, PMWQ.revenue === `gpuHr` ? { pxCase: `low` } : null], [`pxHigh`, `GPU 小時價格 高`, PMWQ.revenue === `gpuHr` ? { pxCase: `high` } : null],
    [`mixRU`, `世代組合 Rubin Ultra 版`, { mixAlt: !0 }], [`sgaGaap`, `管銷率 GAAP（含 SBC）`, PMWQ.cost === `bottomUp` ? { sgaBasis: `gaap` } : null],
    ...(PMWQ.revenue === `tkAnchor` ? [ // W4 r2：定價倍數 k 與隨需占比（與 scripts/permw_sens.py 同一組設定）
      [`kLongLo`, `k_長約 ${Y(AMQ.long.low, 2)}`, { kLong: AMQ.long.low }], [`kLongMed`, `k_長約 ${Y(AMQ.long.sensMedian, 2)}（三筆長約中位數）`, { kLong: AMQ.long.sensMedian }],
      [`kLongHi`, `k_長約 ${Y(AMQ.long.high, 2)}`, { kLong: AMQ.long.high }],
      ...AMQ.onDemandShare.sens.map((x, j) => [`od${j + 1}`, `隨需占比 ${Math.round(x * 100)}%（k_現貨 ${Y(AMQ.spot.base, 2)}）`, { odShare: x }])]: []),
    ...(PMWQ.revenue === `tkAnchor` && CAQ ? (() => { let Z = cvSensQ(); return [ // W5：公司實況驗證的敏感度（與 scripts/permw_sens.py 同一組設定）
      [`kExOff`, `既有合約 k 不套用（全部按 k_新約）`, { kExOn: 0 }], [`kNewUp`, `新約價格 +${Math.round(CAQ.newK.adjSens * 100)}%（公司說法，只作用於長約）`, { kNewAdj: CAQ.newK.adjSens }],
      [`spotCw`, `隨需 ${Math.round(CAQ.spotCw.od * 100)}% × CRWV 短天期 k ${Y(Z.kcw, 2)}`, { odShare: CAQ.spotCw.od, kSpot: Z.kcw }],
      [`opexQ2`, `營運成本＝Q2 實際比率（× ${Y(Z.opex, 2)}）`, { opexScale: Z.opex }], [`capexQ2`, `每 MW 建置成本＝${CALQ.ytdLabel} 實際比率（× ${Y(Z.capex, 2)}）`, { capexScale: Z.capex }],
      [`actBoth`, `營運成本與建置成本皆用公司實際比率`, { opexScale: Z.opex, capexScale: Z.capex }]] })() : [])],
    rows = Object.fromEntries([`low`, `base`, `high`].map(sk => [sk, Object.fromEntries(C.map(([k, , ch]) => {
      if (!ch) return [k, null];
      let s2 = { ...scnQ(e, sk), ...ch }, d2 = runFunding(s2), p2 = runValuation(d2, s2, v);
      return [k, [p2.call.blended, Math.max(0, -d2.totals.preFinEnd)]]
    }))]));
  return { cases: C.map(([k, n]) => [k, n]), rows }
}

function runFunding(e) {
  PMWQ.capex === `tokenomics` && (e.tkCase || e.mixAlt) && (e = { ...e, a: { ...e.a, costMW: capexTkQ(e) } }); // W2 敏感度：成本情境／世代組合改變時重算建置成本
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
    MX = t.accepted.map((n, r) => r < 4 ? t.accepted[r + 1] - n : e.mw31),
    CXF = MN.map((t, n) => (t * (1 - e.lambda) + MX[n] * e.lambda) * e.a.costMW[n] * (e.capexScale ?? 1) / 1e3),
    VIN = Object.fromEntries(Object.entries(e.mwYearEnd).map(([y, m]) => [y, m - (e.mwYearEnd[y - 1] ?? 0)])), // 各年新增 MW（汰換批次）
    REF = PERIOD_FY.map( // 5a：期間的財年年份由日曆推算（滾動後不寫死）
      (t, n) => (VIN[t - e.gpuLife] || 0) * e.a.costMW[n] * (e.capexScale ?? 1) / 1e3),
    CXG = CXF.map((t, n) => n === 0 ? Math.max(t, e.capexFloorFY0 ?? 0) - ACTUAL_1H.capex : t),
    CX = CXG.map((e, t) => e + REF[t]),
    FQ = fleetQ(e, t.accepted), // W2：世代組合（legacy 組合只作對照，不影響任何數字）
    _gh = PMWQ.revenue === `gpuHr` && (t.revMW = revGpuQ(e, FQ)),
    _ta = PMWQ.revenue === `tkAnchor` && (t.revMW = anchorRevQ(e, FQ, t).rev.map(x => x / 1e3)), // W4：Tokenomics 錨 × k（US$bn/MW，100% 計費時數）
    BUQ = FQ ? PERIODS.map((n, r) => buCostQ(e, t, FQ, r, revPassQ(e, t, r), LEASE_CASH_ON_BAL[r] + e.a.newLease[r])) : null,
    PPE = [],
    DAF = CXG.map((t, n) => {
      let r = n === 0 ? e.ppeOpen : PPE[n - 1] + CXG[n - 1];
      return PPE.push(r), (r + .5 * t) / e.gpuLife * PERIOD_YEARS[n]
    }),
    PB = [],
    IX = DEBT_AMORT.map((t, n) => {
      let r = n === 0 ? DBT_P : PB[n - 1][1],
        i = r - t;
      return PB.push([r, i]), (r + i) / 2 * DBT_R * PERIOD_YEARS[n] + CONV_I * PERIOD_YEARS[n] + (n === 0 ? e.intCal : 0)
    }),
    WF = {
      Jn: 0,
      pc: e.cash,
      Dn: 0,
      fr: e.useFacility ? e.facility : 0,
      B: e.rpoOpen + e.rpoPendingAdd,
      pnr: 0,
      sh: 0
    },
    o = PERIODS.map((n, r) => {
      let L = PERIOD_YEARS[r],
        o = e.rpoOpen * (RPO_BUCKET_W[r] / RPO_SCHEDULED_SHARE) * i + e.rpoPendingAdd * RPO_Q3ADD_W[r] * i,
        s = o,
        c = r === 0 ? e.billableOpen : t.billable[r - 1],
        l = e.useAvgMw ? (c + t.billable[r]) / 2 : t.billable[r],
        u = l * t.revMW[r] * (e.revScale ?? 1) * (t.util[r] / 100) * L,
        d = Math.max(0, s - u),
        f = o - d,
        p = 1 - t.recovery[r] / 100,
        m = f * (t.defaultP[r] / 100) * p,
        h = f - m,
        nR = Math.max(0, u - s) * (t.fill[r] / 100),
        b = LEASE_CASH_ON_BAL[r],
        x = e.a.newLease[r],
        S = b + x,
        svc = e.services[r],
        ebM = PMWQ.cost === `bottomUp` ? BUQ[r].eb + (e.ebSteady != null ? (e.ebSteady - BUQ[4].eb) * r / 4 : 0) : e.ebStart + (e.ebSteady - e.ebStart) * r / 4,
        totRev = f + nR + svc,
        cm = ebM + S / Math.max(totRev, .01),
        g = h * cm,
        nC = nR * (1 - (t.defaultP[r] / 100) * p) * cm,
        svcCash = svc * cm,
        _ = CX[r],
        v = _ * e.a.customerFund[r],
        y = _ - v,
        C = t.accepted[r] * 8760 * t.pue[r] * t.power[r] / 1e9 * L,
        w = t.accepted[r] * t.maint[r] / 1e3 * L,
        T = e.overlay && PMWQ.cost !== `bottomUp` ? C + w : 0, // W2：由下而上已含電費與維護，overlay 自動停用
        O = e.includeDebt ? DEBT_AMORT[r] : 0,
        k = r === 0 && e.includeAtm ? e.atm : 0,
        A = g + nC + svcCash + v,
        j0 = _ + S + IX[r] + e.jvCommit[r] + e.a.div[r] + T + O,
        wRL = uA(e, r) / 100 * L,
        wRJ = (e.junkRate + (e.cdsLink ? Math.max(0, e.cds - e.cdsBaseBp) / 1e4 * e.cdsPassThrough : 0)) * L,
        wI0 = wRL * WF.Dn + wRJ * WF.Jn,
        wPre = a + A + k - j0 - wI0,
        wX = Math.max(0, e.minCash - wPre),
        wB = WF.B - o - nR + e.ctrTerm * Math.max(0, nR / L - WF.pnr),
        wEx = (e.includeDebt ? PB[r][1] : DBT_P) + CONV_P, // 5a：評價日後新發可轉債本金讀 company.json → debt.convertible
        wCap = e.debtBacklog * wB - (wEx + WF.Dn),
        wCapD = Math.max(0, wCap, WF.fr),
        wD = Math.min(wCapD, wX / (1 - wRL)),
        wRem = Math.max(0, wX + wRL * wD - wD),
        wCapEq = e.eqCapPct >= 9 ? 1 / 0 : e.eqCapPct * e.eqPx * e.eqCapShares * L,
        wEq = Math.min(wRem, wCapEq),
        wJ = (wRem - wEq) / (1 - wRJ),
        E = wI0 + wRL * wD + wRJ * wJ,
        wSh = wEq / (e.eqPx * (1 - e.eqDisc)),
        D = IX[r] + E,
        j = j0 + E,
        F = wD + wEq + wJ,
        M = A + k + F,
        N = M - j,
        ee = A - (j - O),
        wFr0 = WF.fr,
        wDn0 = WF.Dn,
        wPc = WF.pc + A + k - j0;
      return a += N, WF.fr -= Math.min(WF.fr, wD), WF.Dn += wD, WF.Jn += wJ, WF.pc = wPc, WF.B = wB, WF.pnr = nR / L, WF.sh += wSh, {
        junk: wJ,
        junkEnd: WF.Jn,
        eqCap: wCapEq,
        preFinCum: wPc,
        preFinGap: A + k - j0,
        newDebt: wD,
        newDebtInt: E,
        equity: wEq,
        newShares: wSh,
        cumNewShares: WF.sh,
        newDebtBeg: wDn0,
        newDebtEnd: WF.Dn,
        facBeg: wFr0,
        facEnd: WF.fr,
        backlogEnd: wB,
        debtCap: e.debtBacklog * wB,
        existDebtEnd: wEx,
        totalDebtEnd: wEx + WF.Dn + WF.Jn,
        preCash: wPre,
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
        capexGrowth: CXG[r],
        refresh: REF[r],
        ppeBeg: PPE[r],
        daFleet: DAF[r],
        capexOld: CX_OLD[r],
        intOld: INT_OLD[r],
        rentPerMW: (b + x) / L / ((MB[r] + t.accepted[r]) / 2) * 1e3,
        rentBench: (MB[r] + t.accepted[r]) / 2 * siteBenchQ(e.sites) / 1e3 * LEASE_FACTS.share * L,
        legacy: svcCash,
        servicesRev: svc,
        ebM: ebM,
        cashMargin: cm,
        totRev: totRev,
        ebitdaPL: totRev * ebM,
        cashEbitda: g + nC + svcCash - S,
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
        onBalLease: b,
        offLease: x,
        sourcesOp: A,
        bu: BUQ ? BUQ[r] : null
      }
    }),
    hOp = ACTUAL_1H.cfo - ACTUAL_1H.cashCapex - ACTUAL_1H.jv,
    hFin = ACTUAL_1H.borrow - ACTUAL_1H.cappedCall + ACTUAL_1H.equity,
    hPlug = e.cash - (ACTUAL_1H.cash1231 + hOp + hFin - ACTUAL_1H.debtRepaid),
    _fy = (o[0].fyRevenue = ACTUAL_1H.revenue + o[0].isRev, o[0].fyGross = ACTUAL_1H.capex + o[0].gross, o[0].fyCashCapex = ACTUAL_1H.cashCapex + o[0].cashCapex, o[0].fyLease = ACTUAL_1H.leasePaid + o[0].lease, o[0].fyInterest = ACTUAL_1H.interest + o[0].interest, o[0].fyDebtPay = ACTUAL_1H.debtRepaid + o[0].debtPay, o[0].fyDiv = ACTUAL_1H.jv + o[0].div, o[0].fyAtm = ACTUAL_1H.equity - ACTUAL_1H.cappedCall + o[0].atm, o[0].fyBorrow = ACTUAL_1H.borrow, o[0].fySourcesOp = ACTUAL_1H.cfo + o[0].sourcesOp, o[0].fyOperatingGap = hOp + o[0].operatingGap, o[0].fyExternal = ACTUAL_1H.prepay + o[0].external, o[0].h1 = ACTUAL_1H, o[0].hOp = hOp, o[0].hFin = hFin, o[0].hPlug = hPlug, o[0].fyCfo = ACTUAL_1H.cfo, o[0].fyCapexUse = ACTUAL_1H.cashCapex + o[0].gross, o[0].fyEquity = ACTUAL_1H.equity + o[0].atm, o[0].fyCapped = ACTUAL_1H.cappedCall, o[0].fyUsesCash = ACTUAL_1H.cashCapex + o[0].gross + o[0].lease + o[0].interest + ACTUAL_1H.jv + o[0].div + ACTUAL_1H.debtRepaid + o[0].debtPay + ACTUAL_1H.cappedCall + o[0].op, o[0].fySrcTotal = ACTUAL_1H.cfo + o[0].sourcesOp + ACTUAL_1H.equity + o[0].atm + ACTUAL_1H.borrow + o[0].newDebt + o[0].equity, 0),
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
    id: `rpo-weights`,
    ok: Math.abs(RPO_BUCKET_W.reduce((e, t) => e + t, 0) - RPO_SCHEDULED_SHARE) < 1e-6,
    severity: `watch`,
    title: `RPO 桶（Q2 10-Q）→ 五期權重 [Derived]`,
    detail: `10-Q：$103.7bn，41% 於 24 個月內（至 2028-06）、39% 於 25–48 月、20% 於 49–78 月。模型假設桶內線性：每月 1.708%／1.625%／0.667%，得 2H26 10.25%、FY27 20.5%、FY28 20.0%、FY29 19.5%、FY30 13.75%，合計 84%；其後 16%。桶內前載或後載會改變 2H26／FY27 排程，是 Derived 不是 Verified。`
  }), _({
    id: `lease-10q`,
    ok: !0,
    severity: `ok`,
    title: `在帳租賃現金（10-Q 到期表）`,
    detail: `營業＋融資租賃未折現付款：2026 餘 1.028+0.012、2027 2.116+0.223、2028 2.306、2029 2.373、2030 2.289、其後 19.023；現值負債 16.319+0.221。對應已入帳 ROU 16.595bn，不含 35.5bn 未起租。`
  }), _({
    id: `off-balance-lease`,
    ok: l >= (LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap) * CHECK_TH.leaseVsCommitMin,
    severity: `watch`,
    title: `表外租賃 $35.5bn＋單站上限 $14.7bn（未起租）`,
    detail: `10-Q Note 8：已簽約未起租租賃未折現 $35.5bn（2026–2029 起租、7–16 年），另有單站 393 MW 按造價計租、上限 $14.7bn／16 年，以及 355 MW 按造價計租（金額未定）未含在 35.5 內。模型「表外現金租金」五期 ${e.a.newLease.reduce((e,t)=>e+t,0).toFixed(1)} + 尾端 = ${l.toFixed(0)}bn，為已承諾 50.2bn 的 ${(l/50.2).toFixed(1)} 倍：8 GW 路徑需要的新租約多數尚未簽署，路徑高於已承諾是假設而非錯誤；低於 80% 才標示不一致。`
  }), _({
    id: `rou-h1`,
    ok: !0,
    severity: `ok`,
    title: `H1 新取得營業 ROU $8.36bn`,
    detail: `ROU＝使用權資產。上半年因起租入帳 8.359bn（ROU 由 8.231 增至 16.595）；同期現金租賃支付僅 0.748bn。$35.5bn 尚未 commence，還沒有 ROU。`
  });
  PMW_ONQ && TK_MISSQ.length && _({ // W2：缺漏時顯示、Tokenomics 補齊名稱後自動消失（Excel「檢查_連動」同一列）
    id: `tk-missing`,
    ok: !1,
    severity: `watch`,
    title: `Tokenomics 名稱缺漏 ${TK_MISSQ.length} 項，成本為暫代值`,
    detail: `${TK_MISSQ.join(`、`)} 尚未在 Tokenomics ${(COMPANY_DATA.tkSnap || { source: {} }).source.version || ``} 提供（待 v5.26）。暫代：IT 維護＝輸入頁「維護成本」（CRWV 現值）、人員軟體與稅險＝0、GPU 經濟壽命＝${COMPANY_DATA.defaults.gpuLife} 年。Tokenomics 補齊後重抓快照即自動改用正式值。`
  });
  if (PMWQ.revenue === `tkAnchor` && FQ && tkHasQ(`IF_RevGWFleet`)) { // W4：收入上限檢查（Excel「檢查_連動」同一列；門檻 company.json → methodology.checks.revCapShareMax）
    let rq = PERIODS.map((n, r) => t.revMW[r] * (e.revScale ?? 1) * 1e3 * t.util[r] / 100 / Math.max(1e-9, wMixQ(FQ.mix[r], g => tkMwQ(`IF_RevGWFleet`, g, e.tkCase || `基準`)))),
      mx = Math.max(...rq), th = CHECK_TH.revCapShareMax;
    _({
      id: `rev-cap`,
      ok: mx <= th,
      severity: mx <= th ? `ok` : `watch`,
      title: `收入上限：CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收 最高 ${hA(mx * 100, 0)}（門檻 ${hA(th * 100, 0)}）`,
      detail: `neocloud 拿走客戶 token 營收的比例＝每 MW 計費收入 ÷ Tokenomics IF_RevGWFleet（OpenAI 有效單價、層級組合的理想上限，依在役世代加權）。各期：${PERIODS.map((n, r) => `${n} ${hA(rq[r] * 100, 0)}`).join(`、`)}。超過門檻代表 CRWV 單價相對客戶端 token 經濟性偏高。`
    })
  }
  let v = o[0].gross,
    y = v >= CALL_FACTS.capexLo - LATEST_Q.capexH1 - .5 && v <= CALL_FACTS.capexHi - LATEST_Q.capexH1 + .5;
  _({
    id: `h2-capex`,
    ok: y,
    severity: y ? `ok` : `watch`,
    title: `2026 CapEx 35–39 → 2H26 ${(CALL_FACTS.capexLo-LATEST_Q.capexH1).toFixed(1)}–${(CALL_FACTS.capexHi-LATEST_Q.capexH1).toFixed(1)}`,
    detail: `法說上修全年 CapEx $35–39bn（原 31–35）；Q3 指引 11.5–13.5。H1 已認列 16.139（Q2 9.352），故 2H26 隱含 18.9–22.9。模型 FY26（1H 實際 16.1＋2H 模型）毛 ${(v+ACTUAL_1H.capex).toFixed(1)}。CapEx 口徑＝PP&E 增加＋融資租賃資產－CIP 變動，含 OEM 融資的非現金部分（H1 現金購置 14.117 vs 認列 16.139）。`
  }), _({
    id: `prepay`,
    ok: !0,
    severity: `watch`,
    title: `客戶預付：遞延收入 $9.7bn，H1 淨流入 $1.37bn`,
    detail: `10-Q：遞延收入 9.692bn（年初 8.185）。H1 遞延收入淨增 1.365，對 H1 CapEx 16.139 約 8.5%。此欄是「客戶預付」占毛 CapEx 比率，只降當期現金 CapEx、形成遞延收入，不降專案總成本。FY26 下半年假設 ${((e.a.customerFund[0]||0)*CX[0]).toFixed(1)}bn。`
  }), _({
    id: `call-mw`,
    ok: !0,
    severity: `watch`,
    title: `Q2 主動電力 1.5 GW（+500 MW；6 月 +300 MW）不年化`,
    detail: `法說：Q2 新增近 500 MW、6 月單月 300 MW（後載）；YE26 指引 >1.85 GW；簽約電力 4.2 GW（8/11）；另 1.5 GW powered land／選擇權／LOI。基準 YE26 Accepted ${t.accepted[0]} MW 對上指引下限；FY27 ${t.accepted[1]} MW 對應「3.5 GW 多數於 2027 年底上線」；FY30 ${t.accepted[4]} MW 對上「≥8 GW by 2030」。這是公司路徑，不是模型獨立驗證。`
  }), _({
    id: `call-util`,
    ok: t.util[0] <= 97,
    severity: `watch`,
    title: `「近期產能實質售罄」→ 利用率上限 95%`,
    detail: `法說：near-term capacity effectively sold out、A100 續約至 2029、7 月 SKU 漲價約 25%。基準利用率 ${t.util[0]}%，rev/MW 2H26 ${t.revMW[0]} bn/MW（≈ 期末 ARR 18.5–19.5 ÷ 1.85 GW ≈ 10.3 M/MW，加漲價）。不擬合「售罄」為 100%。`
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
    title: `站點已驗收拉高 2H26`,
    detail: `具名站點已驗收 ${r.accepted} MW，2H26 Accepted 取 max(輸入, 站點)。`
  }) : !e.linkSites && r.accepted > t.accepted[0] + 1 && _({
    id: `site-mismatch`,
    ok: !1,
    severity: `block`,
    title: `站點與年度 MW 不一致`,
    detail: `具名站點已驗收 ${r.accepted} MW，高於 2H26 Accepted ${t.accepted[0]} MW。請打開站點連動。`
  }), t.accepted[4] > r.planned + 50 && _({
    id: `unlisted-mw`,
    ok: !0,
    severity: `watch`,
    title: `未列名容量`,
    detail: `FY30 Accepted ${t.accepted[4]} MW 高於具名規劃 ${r.planned} MW，差額進「其他／未列名站點」。4.2 GW 簽約電力多數未具名。`
  }), e.overlay && _({
    id: `overlay`,
    ok: !1,
    severity: `watch`,
    title: `Overlay 與貢獻率`,
    detail: `電力＋維護已另扣，v2.3 起EBITDAR 率由 EBITDA 率推導、已含電費；開啟 overlay 會重複扣除，僅供壓力測試。CoreWeave 電費多數由房東轉嫁、已在營收成本內。`
  }), e.cdsLink && _({
    id: `cds-link`,
    ok: !0,
    severity: `watch`,
    title: `CDS 傳入透支利率`,
    detail: `透支利率 = 基準邊際利率 + max(0, CDS−450)×40%，只在累積現金為負時生效。基準 9% 已對應 DDTL 5.0（SOFR+450）／9.75% 高收益債水準；再開傳導會部分重複。`
  }), e.includeDebt ? _({
    id: `debt-sched`,
    ok: !0,
    severity: `ok`,
    title: `債務排程攤還（10-Q 本金表，基本情境）`,
    detail: `10-Q 未來本金：2026 餘 4.413、2027 6.184、2028 4.416、2029 2.421、2030 3.221、其後 14.896（合計 35.551）。DDTL 隨客戶付款攤還、OEM 融資 2026–2028 到期，屬契約排程，非壓力測試。關閉＝假設全額再融資。`
  }) : _({
    id: `debt-sched-off`,
    ok: !0,
    severity: `watch`,
    title: `債務攤還已關閉（假設全額再融資）`,
    detail: `DDTL 為契約性攤還、OEM 融資短期到期，關閉等於假設 20.7bn 五年本金全部借新還舊。`
  }), _({
    id: `identity`,
    ok: Math.abs(o[4].cum - (e.cash + o.reduce((e, t) => e + t.gap, 0))) < .01 && Math.abs(o[4].cum - (e.cash + o.reduce((e, t) => e + t.operatingGap + t.atm + t.facility - t.debtPay, 0))) < .01,
    severity: `ok`,
    title: `現金恆等式（期前融資瀑布）`,
    detail: `期末 ${o[4].cum.toFixed(1)} = 6/30 現金 ${e.cash} + 營運缺口 ${o.reduce((e,t)=>e+t.operatingGap,0).toFixed(1)} + 9/17 可轉債／ATM ${o.reduce((e,t)=>e+t.atm,0).toFixed(1)} + 瀑布新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)} + 瀑布股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)} + 高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)} − 排程還本 ${o.reduce((e,t)=>e+t.debtPay,0).toFixed(1)}。每期期末現金不低於最低現金 ${e.minCash}bn——缺口在發生前一期就先融好。`
  }), _({
    id: `waterfall`,
    ok: !0,
    severity: `watch`,
    title: `融資瀑布：新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)}bn、股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)}bn（新股 ${o.reduce((e,t)=>e+t.newShares,0).toFixed(2)}bn 股）、高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)}bn`,
    detail: `順序：未動用額度 → 新債（總債務 ≤ ${e.debtBacklog}× backlog；預設 1.0x 依 2026 年實際融資組合——債務約 77%、股權約 23%——校準，公司 9/17 簡報揭露的當前比率約 0.4x）→ 股權（發行價＝$${e.eqPx}×(1−${(e.eqDisc*100).toFixed(0)}%)，每年上限＝現市值 ${e.eqCapPct>=9?`無上限`:(e.eqCapPct*100).toFixed(0)+`%`}）→ 超出部分以高息債 ${(e.junkRate*100).toFixed(0)}% 補足。backlog 依合約遞減、新簽約以 ${e.ctrTerm} 年合約補入。backlog 因認列而下降、債務上限跟著收縮後，才需要股權；${(q => q.length ? `本情境股權需求落在 ${q.join(`、`)}。` : `本情境不需股權。`)(o.filter(t => t.equity > .05).map(t => t.year))}`
  }), _({
    id: `ebitda-link`,
    ok: o.every(e => Math.abs(e.cashEbitda + e.creditAdj - e.ebitdaPL) < .01),
    severity: `ok`,
    title: `資金與損益同一組 EBITDA：FY30 EBITDA 率 ${(o[4].ebM*100).toFixed(1)}%、EBITDAR 率 ${(o[4].cashMargin*100).toFixed(1)}%`,
    detail: `${PMWQ.cost === `bottomUp` ? `EBITDA 率由下而上推導（Tokenomics 電費、IT 維護等 × 平均在役 MW＋管銷率 ${(sgaPctQ(e)*100).toFixed(1)}%；EBITDAR 率 − 租金 ÷ 營收）：${PERIODS[0]} ${(o[0].ebM*100).toFixed(1)}% → ${PERIODS[4]} ${(o[4].ebM*100).toFixed(1)}%。` : `EBITDA 率由 ${(e.ebStart*100).toFixed(0)}%（Q2 實際）線性爬升至 FY30 ${(e.ebSteady*100).toFixed(0)}%（穩態）。`}資金模型用 EBITDAR 率＝EBITDA 率＋租金÷營收（因租金在支出端另列），非算力服務現金＝服務營收×同一 EBITDAR 率。恆等式：現金 EBITDA（營運來源不含預付 − 租金）＋信用損失調整＝損益 EBITDA，五期皆成立。`
  }), _({
    id: `capex-mw`,
    ok: o[0].capexFull >= 35 && o[0].capexFull <= 39,
    severity: o[0].capexFull >= 35 && o[0].capexFull <= 39 ? `ok` : `watch`,
    title: `毛 CapEx 由 MW 推導：FY26 全年 ${o[0].capexFull.toFixed(1)}（指引 35–39）`,
    detail: `公式＝(本期新增 MW×(1−λ)＋次期新增 MW×λ)×每 MW 成本。YE25 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW（Q4 法說 >850）→ YE26 ${t.accepted[0]} MW、λ ${(e.lambda*100).toFixed(0)}%、每 MW $${e.a.costMW[0]}m。五期（FY26 為下半年）合計 ${CX.reduce((e,t)=>e+t,0).toFixed(1)}，v1.4 手動值 ${CX_OLD.reduce((e,t)=>e+t,0)}，差 ${(CX.reduce((e,t)=>e+t,0)-CX_OLD.reduce((e,t)=>e+t,0)).toFixed(1)}——舊值對 FY28–30 每年 +1.5–1.7 GW 明顯不足。FY26 指引把每 MW 成本釘在約 $32–37m（λ 25–45%）；GPU 汰換的維持性 CapEx 未納入。`
  }), _({
    id: `fleet-da`,
    ok: !0,
    severity: `watch`,
    title: `D&A 改由車隊推算：FY30 ${o[4].daFleet.toFixed(1)}bn（壽命 ${e.gpuLife} 年）`,
    detail: `D&A＝(期初毛 PP&E＋本期成長型 CapEx×½)÷壽命。v1.6 以營收比例（FY30 36%）估 D&A，與 $${e.a.costMW[4]}m/MW 的建置成本不一致——CapEx 隨 MW 增加，折舊也必須跟著增加。Q2 實際 D&A／營收 54%，等於每 MW 成本 $34m÷6 年÷每 MW 年收入約 $10.5m，本版在穩態下重現此比例。後果：GAAP 營業利益＝EBITDA 率約 59% − D&A 約 54–70%，擴張愈快愈接近零或為負。`
  }), _({
    id: `gpu-refresh`,
    ok: !0,
    severity: `watch`,
    title: `GPU 汰換 CapEx：FY29 ${o[3].refresh.toFixed(1)}、FY30 ${o[4].refresh.toFixed(1)}`,
    detail: `2023／2024 年批次在第 ${e.gpuLife} 年汰換（YE24 360 MW 為 S-1 揭露，YE23 ${Object.values(e.mwYearEnd)[0]} MW 為推估）。汰換取代已折舊完的設備，不增加折舊基礎。法說稱 A100 續約至 2029——舊世代可延長使用，汰換可能遞延，此處取壽命年限為基準。`
  }), _({
    id: `capex-floor`,
    ok: o[0].capexFull < (e.capexFloorFY0 ?? 0) ? !1 : !0,
    severity: `watch`,
    title: o[0].capexFull < (e.capexFloorFY0 ?? 0) ? `FY26 CapEx 公式值 ${o[0].capexFull.toFixed(1)} 低於下限 ${e.capexFloorFY0}：差額 ${(e.capexFloorFY0 - o[0].capexFull).toFixed(1)} 為轉向太晚的成本` : `FY26 CapEx 公式值 ${o[0].capexFull.toFixed(1)} 高於下限 ${e.capexFloorFY0}`,
    detail: `2026 年的支出多已下單（全年指引 35–39，1H 已認列 16.1）。若 FY27 的新增 MW 少到公式值低於下限，代表已採購的設備超過實際上線需求——這部分在模型中不帶來額外收入。`
  }), _({
    id: `debt-sched-int`,
    ok: Math.abs(DBT_P - 35.551) < .01,
    severity: `ok`,
    title: `存量利息由既有債務明細推算：加權有效利率 ${(DBT_R*100).toFixed(1)}%`,
    detail: `16 筆工具合計 ${DBT_P.toFixed(3)}（10-Q 35.551）。存量利息＝本金依到期表遞減的平均值×${(DBT_R*100).toFixed(1)}%×期間長度＋9/17 可轉債 3.7×2.875%＋FY26 校準 ${e.intCal}（對上 Q3 指引 0.86–0.94／季）。五期合計 ${o.reduce((e,t)=>e+t.intStock,0).toFixed(2)}，v1.4 手動值 ${INT_OLD.reduce((e,t)=>e+t,0).toFixed(1)}——舊值未隨還本遞減。限制：利率組合假設不變。`
  }), _({
    id: `jv-commit`,
    ok: Math.abs(e.jvCommit.reduce((e,t)=>e+t,0) - 1.15) < .01,
    severity: `ok`,
    title: `JV 已承諾餘額 1.15 於 2026 年內履行`,
    detail: `10-Q Note 3：兩個 JV 承諾最高 1.7bn，預計 2026 年內履行，入夥後須依協議追加出資。H1 已付 0.55 → 餘 1.15 落在 FY26 下半年（v1.4 只放 0.6）。後續增資 ${e.a.div.join('／')} 未揭露，屬 [Assumed]。`
  }), _({
    id: `rent-bench`,
    ok: o[4].rentPerMW >= siteBenchQ(e.sites) * LEASE_FACTS.share * CHECK_TH.rentVsBenchMin,
    severity: `watch`,
    title: `租金檢驗：模型每 MW 年租金 FY30 $${o[4].rentPerMW.toFixed(2)}m vs 市場基準 $${siteBenchQ(e.sites).toFixed(2)}m`,
    detail: `具名站點（Helios／Polaris Forge 1／Core Scientific，1,516 MW、合約 36.2bn）加權每 MW 年租金 $${siteBenchQ(e.sites).toFixed(2)}m。模型從 FY26 $${o[0].rentPerMW.toFixed(2)}m 降到 FY30 $${o[4].rentPerMW.toFixed(2)}m；以第三方租賃占比 ${LEASE_FACTS.share*100}% 計，五期租金可能低估約 ${o.reduce((e,t)=>e+(t.rentBench-t.lease),0).toFixed(1)}bn，集中在 FY29–30。本版未改為 MW 驅動，避免與 CapEx 同時放大同一條 8 GW 假設。`
  }), _({
    id: `facility`,
    ok: !0,
    severity: `ok`,
    title: e.useFacility ? `未動用額度 ${e.facility}bn 為瀑布第一順位` : `未動用額度不動用`,
    detail: `10-Q：可用額度 ${LATEST_Q.availability}bn（RCF＋DDTL 未動用，DDTL 4.0 可動用至 2027-06-30）。已承諾額度，不受 債務／backlog 上限限制；用罄後才進入資產層新債與股權。`
  }), _({
    id: `cds-level`,
    ok: e.cds < 800,
    severity: e.cds >= 800 ? `watch` : `ok`,
    title: `5Y CDS ${Y(e.cds,0)} bps（${e.cdsDate}）`,
    detail: `買賣 ${Y(e.cdsBid,0)}／${Y(e.cdsAsk,0)}，近期區間 ${Y(e.cdsLo,0)}–${Y(e.cdsHi,0)}。以 40% 回收率的簡化式估算，720bps 隱含年化違約機率約 ${(720/100/(1-.4)).toFixed(1)}%、五年累積約 ${((1-Math.pow(1-720/1e4/(1-.4),5))*100).toFixed(0)}%。較 7 月峰值 855 收窄、仍高於 6 月低點 452。這是 CoreWeave 自身信用，不是客戶違約率。`
  }), _({
    id: `concentration`,
    ok: !0,
    severity: `watch`,
    title: `客戶集中：A 36%／B 26%／C 10%（Q2 營收）`,
    detail: `10-Q：前三大客戶合計 72% 營收；應收 A 32%、B 32%。Microsoft、OpenAI（$6.5bn 至 2031-05）、Meta（$21bn 至 2032-12）為揭露之重大客戶；Jane Street 承諾約 $6.0bn（私人公司）。信用損失欄的違約率是對此集中度的定價，不是預測。`
  }), _({
    id: `tier34`,
    ok: !0,
    severity: `watch`,
    title: `Bernstein：主動電力 25%、簽約電力 74% 位於 Tier 3/4 市場`,
    detail: `9/14 報告（Underperform，目標 $74）：訓練需求放緩將先衝擊鄉村站點；既有 backlog 為 take-or-pay 受保護，風險在「已簽約電力尚未售出」部分。本模型未列名容量占比即此風險的量化入口。`
  });
  let C = r.accepted,
    w = Math.max(0, t.accepted[0] - C),
    T = w / Math.max(t.accepted[0], 1);
  return _({
    id: `residual-mw`,
    ok: T <= .5,
    severity: T > .5 || T > .3 ? `watch` : `ok`,
    title: `未列名站點占 2H26 Accepted ${Math.round(T*100)}%`,
    detail: `具名已驗收 ${C} MW，2H26 Accepted ${t.accepted[0]} MW，殘差 ${w.toFixed(0)} MW。公司揭露 51 站但不給逐站 MW，殘差必然偏高；超過 50% 表示容量預測大多不是具名站點支撐。`
  }), _({
    id: `fill-timing`,
    ok: !0,
    severity: `watch`,
    title: `新簽約收入採當期入帳（無爬坡遞延）`,
    detail: `模型把「未被期初 RPO 占用的產能 × 簽約率」在當期全額認列為收入（FY28–FY30 合計 ${(o[2].newRev+o[3].newRev+o[4].newRev).toFixed(0)}bn）。真實合約自簽署到滿載有數季爬坡（法說：新部署會壓低貢獻率、穩定後回到 mid-20%s）。此設定使收入與現金偏樂觀——若改為半期遞延，FY28–30 收入約少 15–20%，缺口會再擴大。方向上對本模型的「賣出」結論不利於翻案，但使用者若要做多方情境，須先把這條收緊。`
  }), e.useAvgMw || _({
    id: `year-end-mw`,
    ok: !0,
    severity: `watch`,
    title: `收入用年末 Billable 全年化`,
    detail: `預設關閉時用年末存量×全年單價，年中才交付的 MW 會被當成全年在役。CoreWeave 交付後載（6 月 300 MW），建議打開「平均在役 MW」。`
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
      newRev: s(`newRev`),
      newCash: s(`newCash`),
      lease: s(`lease`),
      interest: s(`interest`),
      overdraft: s(`overdraft`),
      div: s(`div`),
      atm: s(`atm`),
      facility: s(`facility`),
      newDebt: s(`newDebt`),
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
    sites: n,
    m: t,
    fleet: FQ,
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
    ebs = [.59, .65, .7, .75],
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
    rev30: base.d.m.revMW[4] * (e.revScale ?? 1) * 1e3, // W2：GPU 小時路線時為推導值（舊方法＝輸入值，數值相同）
    util30: e.m.util[4] / 100,
    cost30: e.a.costMW[4] * (e.capexScale ?? 1),
    eb30: e.ebSteady ?? base.d.years[4].ebM // W2 由下而上：穩態預設＝由下而上 FY30
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
  return r(`穩態 EBITDA 率`, `59%`, `70%`, t => {
    t.ebSteady = .59
  }, t => {
    t.ebSteady = .7
  }), r(`債務／backlog 上限`, `0.4x`, `1.2x`, t => {
    t.debtBacklog = .4
  }, t => {
    t.debtBacklog = 1.2
  }), r(`每 MW 建置成本`, `+20%`, `−20%`, e => {
    e.a.costMW = e.a.costMW.map(e => e * 1.2)
  }, e => {
    e.a.costMW = e.a.costMW.map(e => e * .8)
  }), r(`GPU 經濟壽命`, `5 年`, `8 年`, e => {
    e.gpuLife = 5
  }, e => {
    e.gpuLife = 8
  }), r(`FY31 新增 MW`, `1,700`, `0`, e => {
    e.mw31 = 1700
  }, e => {
    e.mw31 = 0
  }), r(`表外現金租金`, `+30%`, `−30%`, e => {
    e.a.newLease = e.a.newLease.map(e => e * 1.3)
  }, e => {
    e.a.newLease = e.a.newLease.map(e => e * .7)
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
    PMWQ.revenue !== `legacy` ? e.revScale = (e.revScale ?? 1) * .85 : e.m.revMW = e.m.revMW.map(e => e * .85) // W2：GPU 小時路線的 revMW 由價格推導，改用整體倍數
  }, e => {
    PMWQ.revenue !== `legacy` ? e.revScale = (e.revScale ?? 1) * 1.15 : e.m.revMW = e.m.revMW.map(e => e * 1.15)
  }), r(`新產能簽約率`, `−20pt`, `+20pt`, e => {
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
  }), r(`WACC`, `12.5%`, `9.5%`, null, null, e => {
    e.wacc = .125
  }, e => {
    e.wacc = .095
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
