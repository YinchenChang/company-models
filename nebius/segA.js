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
  SC_ACC = Object.fromEntries([`low`, `base`, `high`].map(k => [k, PERIOD_YEARS.reduce((a, L, i) => (a.push(Math.min(SC_MWP.contracted[k][i], i === 0 ? SC_MWP.connectedStart : a[i - 1] + SC_MWP.pace[k] * L)), a), [])])),
  SC_REV = COMPANY_DATA.scenarios.revMW, // v0.1b：每 MW 年收入隨情境（Tokenomics 正向推導三情境）
  SC_BR = COMPANY_DATA.scenarios.billableRatio.ratio,
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
    rev: SC_REV[k],
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
    MX = t.accepted.map((n, r) => r < 4 ? t.accepted[r + 1] - n : e.mw31),
    CXF = MN.map((t, n) => (t * (1 - e.lambda) + MX[n] * e.lambda) * e.a.costMW[n] * (e.capexScale ?? 1) / 1e3),
    VIN = Object.fromEntries(Object.entries(e.mwYearEnd).map(([y, m]) => [y, m - (e.mwYearEnd[y - 1] ?? 0)])), // 各年新增 MW（汰換批次）
    REF = PERIOD_FY.map( // 5a：期間的財年年份由日曆推算（滾動後不寫死）
      (t, n) => (VIN[t - e.gpuLife] || 0) * e.a.costMW[n] * (e.capexScale ?? 1) / 1e3),
    CXG = CXF.map((t, n) => n === 0 ? Math.max(t, e.capexFloorFY0 ?? 0) - ACTUAL_1H.capex : t),
    CX = CXG.map((e, t) => e + REF[t]),
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
        ebM = e.ebStart + (e.ebSteady - e.ebStart) * r / 4,
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
        T = e.overlay ? C + w : 0,
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
        sourcesOp: A
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
    id: `rev-mw`,
    ok: o.every(y => Math.abs(y.isRev - y.capacity) < 1e-9) || e.revenueDriver !== `mw`,
    severity: `ok`,
    title: `營收＝平均在役 MW × 每 MW 年收入 × 利用率 × 期間長度`,
    detail: `${e.revenueDriver === `mw` ? `MW 驅動` : `RPO 驅動`}：五期算力營收 ${o.map(y => Y(y.isRev, 2)).join(`／`)}；容量上限 ${o.map(y => Y(y.capacity, 2)).join(`／`)}。每 MW 年收入 ${t.revMW.map(x => Y(x * 1e3, 2)).join(`／`)} US$m/MW-IT（Tokenomics 正向推導；公司 ACV $20–25M 只作對照）。`
  }), _({
    id: `rpo-weights`,
    ok: Math.abs(RPO_BUCKET_W.reduce((e, t) => e + t, 0) - RPO_SCHEDULED_SHARE) < 1e-6,
    severity: `watch`,
    title: `RPO 桶 → 五期權重 [Derived]（只作對照，不驅動營收）`,
    detail: `季報：RPO $${Y(LATEST_Q.rpo, 1)}bn，${hA(LATEST_Q.rpo24m * 100, 0)} 於 24 個月內、${hA(LATEST_Q.rpo25to48 * 100, 0)} 於 25–48 個月、其餘 ${hA(LATEST_Q.rpo49to78 * 100, 0)} 之後（假設 49–72 個月）。桶內線性分攤得五期 ${RPO_BUCKET_W.map(x => hA(x * 100, 1)).join(`／`)}，合計 ${hA(RPO_SCHEDULED_SHARE * 100, 0)}。營收由 MW × 每 MW 年收入驅動，RPO 排程只用於產能瓶頸旗標與債務上限的 backlog；桶內前載或後載是 Derived，不是 Verified。`
  }), _({
    id: `lease-10q`,
    ok: !0,
    severity: `ok`,
    title: `在帳租賃現金（季報到期表）`,
    detail: `營業租賃未折現付款：五期 ${LEASE_CASH_ON_BAL.map(x => Y(x, 3)).join(`／`)}，其後 ${Y(LEASE_AFTER_FY30, 2)}；租賃負債現值 ${Y(LATEST_Q.opLeaseLiab, 3)}。自有站點占合約電力 >75%，租金占比低；不含 ${Y(LATEST_Q.offBalanceLease, 1)}bn 未起租。`
  }), _({
    id: `off-balance-lease`,
    ok: l >= (LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap) * CHECK_TH.leaseVsCommitMin,
    severity: `watch`,
    title: `已簽約未起租租賃 $${Y(LATEST_Q.offBalanceLease, 1)}bn`,
    detail: `季報附註：已簽約未起租租賃未折現 $${Y(LATEST_Q.offBalanceLease, 2)}bn（預計 2026–2027 起租、租期最長 12 年）。模型「表外現金租金」五期 ${e.a.newLease.reduce((e,t)=>e+t,0).toFixed(1)} + 尾端 = ${l.toFixed(1)}bn，為已承諾 ${Y(LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap, 1)}bn 的 ${(l/Math.max(LATEST_Q.offBalanceLease + LATEST_Q.singleSiteCap, .01)).toFixed(1)} 倍；低於 ${hA(CHECK_TH.leaseVsCommitMin * 100, 0)} 才標示不一致。`
  }), _({
    id: `rou-h1`,
    ok: !0,
    severity: `ok`,
    title: `營業租賃負債 $${Y(LATEST_Q.opLeaseLiab, 2)}bn`,
    detail: `使用權資產對應的租賃負債 ${Y(LATEST_Q.opLeaseLiab, 3)}bn（季報：加權平均剩餘約 8.7 年、折現率 6.3%）。$${Y(LATEST_Q.offBalanceLease, 1)}bn 尚未起租，還沒有使用權資產。`
  });
  let v = o[0].gross,
    y = v >= CALL_FACTS.capexLo - LATEST_Q.capexH1 - .5 && v <= CALL_FACTS.capexHi - LATEST_Q.capexH1 + .5;
  _({
    id: `h2-capex`,
    ok: y,
    severity: y ? `ok` : `watch`,
    title: `${PERIOD_FY[0]} CapEx 指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi} → ${PERIODS[0]} 模型期 ${(CALL_FACTS.capexLo-LATEST_Q.capexH1).toFixed(1)}–${(CALL_FACTS.capexHi-LATEST_Q.capexH1).toFixed(1)}`,
    detail: `全年資本支出指引 $${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}bn（法說會轉述，股東信未列金額 [Interested-party]）；年初至今現金資本支出 ${Y(LATEST_Q.capexH1, 2)}（應計口徑未揭露，以現金口徑代替）。模型 ${PERIOD_FY[0]}（年初至今實際＋首期模型）毛額 ${(v+ACTUAL_1H.capex).toFixed(1)}。`
  }), _({
    id: `prepay`,
    ok: !0,
    severity: `watch`,
    title: `客戶預付：合約負債 $${Y(LATEST_Q.deferredTotal, 2)}bn，年初至今淨增 $${Y(LATEST_Q.deferredIn, 2)}bn`,
    detail: `季報：遞延營收（合約負債）${Y(LATEST_Q.deferredTotal, 3)}bn，年初至今增加 ${Y(LATEST_Q.deferredIn, 3)}，推估預付現金 ${Y(ACTUAL_1H.prepay, 3)}（[Derived]）。股東信：約 70% 合約含預付、覆蓋相關資本支出 50–60%、2026 年預期預付 >$9B（[Interested-party]）。模型首期預付流入 ${Y(o[0].external, 2)}bn。`
  }), _({
    id: `call-mw`,
    ok: !0,
    severity: `watch`,
    title: `已連網 MW-IT：${PERIOD_FY[0] - 1} 年底 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW → ${PERIODS[0]} 年底 ${t.accepted[0]} MW`,
    detail: `股東信：2025 年底 active power 約 170 MW；2026 年底 connected 目標 800–1,000 MW、合約電力目標 5 GW（目前 >3.5 GW）；2027 起每年部署 >1 GW。口徑不明者以 ÷ PUE 1.2 換成 MW-IT。目前情境：${PERIODS[0]} 已連網 ${t.accepted[0]} MW、${PERIODS[1]} ${t.accepted[1]} MW、${PERIODS[4]} ${t.accepted[4]} MW。季末 active MW 未揭露，不以營收反推。`
  }), _({
    id: `call-util`,
    ok: t.util[0] <= 100,
    severity: `watch`,
    title: `利用率 ${t.util[0]}%：每 MW 年收入已含可計費利用率`,
    detail: `每 MW 年收入取 Tokenomics 正向推導（路徑 B 已乘可計費利用率 80／85／90%），所以利用率欄預設 100%、不重複扣除；欄位保留供壓力測試。目前每 MW 年收入 $${Y(t.revMW[0] * 1e3 * (e.revScale ?? 1), 2)}m/MW-IT。`
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
    detail: `期末 ${o[4].cum.toFixed(1)} = 評價日現金 ${e.cash} + 營運缺口 ${o.reduce((e,t)=>e+t.operatingGap,0).toFixed(1)} + 期後股權／可轉債 ${o.reduce((e,t)=>e+t.atm,0).toFixed(1)} + 瀑布新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)} + 瀑布股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)} + 高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)} − 排程還本 ${o.reduce((e,t)=>e+t.debtPay,0).toFixed(1)}。每期期末現金不低於最低現金 ${e.minCash}bn——缺口在發生前一期就先融好。`
  }), _({
    id: `waterfall`,
    ok: !0,
    severity: `watch`,
    title: `融資瀑布：新債 ${o.reduce((e,t)=>e+t.newDebt,0).toFixed(1)}bn、股權 ${o.reduce((e,t)=>e+t.equity,0).toFixed(1)}bn（新股 ${o.reduce((e,t)=>e+t.newShares,0).toFixed(2)}bn 股）、高息債 ${o.reduce((e,t)=>e+t.junk,0).toFixed(1)}bn`,
    detail: `順序：未動用額度（資產擔保融資）→ 新債（總債務 ≤ ${e.debtBacklog}× backlog）→ 股權（發行價＝$${e.eqPx}×(1−${(e.eqDisc*100).toFixed(0)}%)，每年上限＝現市值 ${e.eqCapPct>=9?`無上限`:(e.eqCapPct*100).toFixed(0)+`%`}）→ 超出部分以高息債 ${(e.junkRate*100).toFixed(0)}% 補足。backlog 依合約遞減、新簽約以 ${e.ctrTerm} 年合約補入。${(q => q.length ? `本情境股權需求落在 ${q.join(`、`)}。` : `本情境不需股權。`)(o.filter(t => t.equity > .05).map(t => t.year))}`
  }), _({
    id: `ebitda-link`,
    ok: o.every(e => Math.abs(e.cashEbitda + e.creditAdj - e.ebitdaPL) < .01),
    severity: `ok`,
    title: `資金與損益同一組 EBITDA：FY30 EBITDA 率 ${(o[4].ebM*100).toFixed(1)}%、EBITDAR 率 ${(o[4].cashMargin*100).toFixed(1)}%`,
    detail: `EBITDA 率由 ${(e.ebStart*100).toFixed(0)}%（Q2 實際）線性爬升至 FY30 ${(e.ebSteady*100).toFixed(0)}%（穩態）。資金模型用 EBITDAR 率＝EBITDA 率＋租金÷營收（因租金在支出端另列），非算力服務現金＝服務營收×同一 EBITDAR 率。恆等式：現金 EBITDA（營運來源不含預付 − 租金）＋信用損失調整＝損益 EBITDA，五期皆成立。`
  }), _({
    id: `capex-mw`,
    ok: o[0].capexFull >= CALL_FACTS.capexLo && o[0].capexFull <= CALL_FACTS.capexHi,
    severity: o[0].capexFull >= CALL_FACTS.capexLo && o[0].capexFull <= CALL_FACTS.capexHi ? `ok` : `watch`,
    title: `毛 CapEx 由 MW 推導：${PERIODS[0]} 全年 ${o[0].capexFull.toFixed(1)}（指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}）`,
    detail: `公式＝(本期新增 MW×(1−λ)＋次期新增 MW×λ)×每 MW 成本。${PERIOD_FY[0] - 1} 年底 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW → ${PERIODS[0]} 年底 ${t.accepted[0]} MW、λ ${(e.lambda*100).toFixed(0)}%、每 MW $${e.a.costMW[0]}m（Tokenomics IF_CapexTotal：IT＋機房，自有站點）。五期（首期為模型部分）合計 ${CX.reduce((e,t)=>e+t,0).toFixed(1)}。指引只有法說會二手轉述，且口徑（現金或應計）未定義。`
  }), _({
    id: `fleet-da`,
    ok: !0,
    severity: `watch`,
    title: `D&A 改由車隊推算：FY30 ${o[4].daFleet.toFixed(1)}bn（壽命 ${e.gpuLife} 年）`,
    detail: `D&A＝(期初毛 PP&E＋本期成長型 CapEx×½)÷壽命。期初毛 PP&E ${e.ppeOpen} 含尚未啟用資產；季報 D&A ${Y(LATEST_Q.da, 3)}／季，模型首期明顯偏高（模板已知限制，折舊修正不在本次範圍）。CapEx 隨 MW 增加，折舊跟著增加。`
  }), _({
    id: `gpu-refresh`,
    ok: !0,
    severity: `watch`,
    title: `GPU 汰換 CapEx：FY29 ${o[3].refresh.toFixed(1)}、FY30 ${o[4].refresh.toFixed(1)}`,
    detail: `${PERIOD_FY[0] - 1} 年底批次（${e.mwYearEnd[PERIOD_FY[0] - 1]} MW，active power）在第 ${e.gpuLife} 年汰換；更早批次的 MW 未揭露。汰換取代已折舊完的設備，不增加折舊基礎。`
  }), _({
    id: `capex-floor`,
    ok: o[0].capexFull < (e.capexFloorFY0 ?? 0) ? !1 : !0,
    severity: `watch`,
    title: o[0].capexFull < (e.capexFloorFY0 ?? 0) ? `FY26 CapEx 公式值 ${o[0].capexFull.toFixed(1)} 低於下限 ${e.capexFloorFY0}：差額 ${(e.capexFloorFY0 - o[0].capexFull).toFixed(1)} 為轉向太晚的成本` : `FY26 CapEx 公式值 ${o[0].capexFull.toFixed(1)} 高於下限 ${e.capexFloorFY0}`,
    detail: `${PERIOD_FY[0]} 的支出多已下單（全年指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}，年初至今 ${Y(ACTUAL_1H.capex, 1)}）。若次年新增 MW 少到公式值低於下限，代表已採購的設備超過實際上線需求——這部分在模型中不帶來額外收入。`
  }), _({
    id: `debt-sched-int`,
    ok: Math.abs(DBT_P - LATEST_Q.debtPrincipal) < .01,
    severity: `ok`,
    title: `存量利息由既有債務明細推算：加權有效利率 ${(DBT_R*100).toFixed(1)}%`,
    detail: `${DEBT_TOOLS.length} 筆工具合計 ${DBT_P.toFixed(3)}（季報 ${LATEST_Q.debtPrincipal}）。存量利息＝本金依到期表遞減的平均值×${(DBT_R*100).toFixed(1)}%×期間長度＋期後新發可轉債 ${CONV_P}×${hA(COMPANY_DATA.debt.convertible.coupon * 100, 2)}＋首期校準 ${e.intCal}。可轉債以現金票息計，不含折價攤銷（非現金）。五期合計 ${o.reduce((e,t)=>e+t.intStock,0).toFixed(2)}。`
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
    title: `租金檢驗：模型每 MW 年租金 FY30 $${o[4].rentPerMW.toFixed(2)}m vs 市場基準 $${siteBenchQ(e.sites).toFixed(2)}m`,
    detail: `具名站點${siteBenchQ(e.sites) > 0 ? `加權每 MW 年租金 $${siteBenchQ(e.sites).toFixed(2)}m` : `沒有租約金額（自有為主），市場租金基準不適用`}。模型每 MW 年租金 ${PERIODS[0]} $${o[0].rentPerMW.toFixed(2)}m、${PERIODS[4]} $${o[4].rentPerMW.toFixed(2)}m。`
  }), _({
    id: `facility`,
    ok: !0,
    severity: `ok`,
    title: e.useFacility ? `未動用額度 ${e.facility}bn（資產擔保融資）為瀑布第一順位` : `未動用額度不動用`,
    detail: `${CALL_FACTS.postQShortDated || ``}。已承諾額度，不受 債務／backlog 上限限制；用罄後才進入新債與股權。`
  }), _({
    id: `cds-level`,
    ok: e.cds == null || e.cds < 800,
    severity: e.cds != null && e.cds >= 800 ? `watch` : `ok`,
    title: e.cds == null ? `CDS：${e.cdsDate}` : `5Y CDS ${Y(e.cds,0)} bps（${e.cdsDate}）`,
    detail: e.cds == null ? `沒有 CDS 報價資料；新債利率改用資產擔保融資 SOFR＋2.50% 與可轉債票息推估。` : `買賣 ${Y(e.cdsBid,0)}／${Y(e.cdsAsk,0)}，近期區間 ${Y(e.cdsLo,0)}–${Y(e.cdsHi,0)}。以 40% 回收率的簡化式估算，隱含年化違約機率約 ${(e.cds/100/(1-.4)).toFixed(1)}%。`
  }), _({
    id: `concentration`,
    ok: !0,
    severity: `watch`,
    title: `客戶集中：${hA(LATEST_Q.custA * 100, 0)}／${hA(LATEST_Q.custB * 100, 0)}／${hA(LATEST_Q.custC * 100, 0)}（最新季營收）`,
    detail: `季報：占營收 ≥10% 的客戶三家合計 ${hA((LATEST_Q.custA + LATEST_Q.custB + LATEST_Q.custC) * 100, 0)}；名稱未揭露（可能含 Microsoft、Meta，未經確認）。已揭露大單：Microsoft 5 年 TCV 約 $17.4B、Meta 兩份協議（$2.9B；$12B 專用＋最高 $15B 未售容量承購）。信用損失欄的違約率是對此集中度的定價，不是預測。`
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
    e.m.revMW = e.m.revMW.map(e => e * .85)
  }, e => {
    e.m.revMW = e.m.revMW.map(e => e * 1.15)
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
