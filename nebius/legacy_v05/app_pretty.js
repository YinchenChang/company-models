zk = [`FY27`, `FY28`, `FY29`, `FY30`, `FY31`], Bk = `2026-09-14`, Vk = `時序先法說、後 10-Q。入帳數字聽 10-Q；全年指引聽法說；事實打架聽 10-Q。Q1 的 MW／利用率當觀察，不年化進基準。`, Hk = {
  version: `4.0`,
  date: `2026-09-11`,
  rpo: 664,
  rp: 84,
  endCash: -79.9,
  target: 118,
  call: `賣出`,
  fy27Accepted: 2200,
  util0: 92
}, Uk = .84, Wk = [.13, .1665, .2035, .153, .187], Gk = [.13 / .84, .1665 / .84, .2035 / .84, .153 / .84, .187 / .84], Kk = [4.975, 4.987, 5.02, 5.056, 5.06], qk = [7.199, 10.145, 5.5, 7.25, 9.75], Jk = 39.283, Yk = {
  filed: `2026-09-11`,
  periodEnd: `2026-08-31`,
  revenue: 19.345,
  yoy: .3,
  cloud: 11.607,
  oci: 7.388,
  ociYoy: 1.21,
  apps: 4.219,
  appsYoy: .1,
  software: 5.55,
  license: .655,
  support: 4.895,
  hardware: .774,
  services: 1.414,
  opInc: 6.728,
  opM: .35,
  interest: 1.428,
  taxRate: .151,
  ni: 4.76,
  niCommon: 4.679,
  epsDiluted: 1.56,
  epsBasic: 1.58,
  sbc: 1.127,
  da: 3.156,
  rpo: 664,
  rpoFy26: 638,
  rpoNext12: .13,
  rpoM13to36: .37,
  rpoM37to60: .34,
  cashMs: 37.077,
  cashMsYe: 31.894,
  notesCurrent: 7.625,
  notesLong: 117.712,
  netDebtQ1: 88.26,
  netDebtYe: 97.647,
  sharesOut: 3.023736,
  dilutedWaso: 3,
  basicWaso: 2.966,
  atmNet: 19.909,
  atmShares: .141,
  capexCash: 28.499,
  capexUnpaid: 6.247,
  cfo: 23.103,
  fcf: -5.396,
  prepaySfc: 11.363,
  deferredTotal: 30.789,
  deferredCurrent: 14.686,
  ppe: 127.845,
  offBalanceLease: 288,
  onBalanceLeaseLiab: 43.806,
  onBalanceUndiscounted: 63.245,
  purchaseObligations: 34.15,
  leaseCashQ1: 1.136,
  rouOp: 33.967,
  rouFin: 8.856,
  rouOpObtained: 4.903,
  rouFinObtained: 1.539,
  buybackAuth: 6.3,
  commonDiv: .5,
  ellisonPlanShares: 50
}, Xk = {
  mwDelivered: 850,
  gpus: 3e5,
  gpuUtil: 97.9,
  fy26Delivered: 1164,
  abileneMw: 618,
  abileneShare: .75,
  abileneCampusMw: 824,
  abileneGpusQ1: 131e3,
  aiContractsNoCash: 30,
  rpo36m: .5,
  capexLo: 90,
  capexHi: 95,
  netCashCapexMax: 70,
  q1NetCashCapex: 18,
  fy27RevFloor: 90,
  ngEpsFy: 8.1,
  q1NgEps: 1.92,
  q1NgOpInc: 8.2,
  q1NgOpM: .42,
  q2RevGLo: .3,
  q2RevGHi: .34,
  q2CloudGLo: .65,
  q2CloudGHi: .71,
  q2NgEpsLo: 1.85,
  q2NgEpsHi: 1.93,
  renewalPremium: .2,
  investorDay: `2026-10`,
  nextEarn: `2026-12-14`
}, Zk = {
  low: {
    label: `低資本`,
    a: {
      capex: [88, 80, 50, 35, 25],
      customerFund: [.3, .3, .3, .3, .3],
      newLease: [3, 8, 12, 14, 15],
      interest: [5.2, 5.5, 5.4, 5.2, 5],
      div: [6.37, 6.45, 6.5, 6.5, 6.6],
      legacy: [30, 30, 30, 30, 30],
      margin: .4
    }
  },
  base: {
    label: `基準`,
    a: {
      capex: [92.5, 95, 70, 50, 35],
      customerFund: [.243, .25, .25, .25, .25],
      newLease: [4, 12, 16, 18, 18],
      interest: [5.7, 6, 6.2, 6, 5.8],
      div: [6.37, 6.5, 6.6, 6.7, 6.8],
      legacy: [25, 25, 25, 25, 25],
      margin: .35
    }
  },
  high: {
    label: `高資本`,
    a: {
      capex: [95, 110, 90, 70, 55],
      customerFund: [.2, .2, .2, .2, .2],
      newLease: [6, 16, 20, 22, 22],
      interest: [6, 7, 8, 8.5, 9],
      div: [6.37, 6.55, 6.7, 6.8, 7.1],
      legacy: [20, 20, 20, 20, 20],
      margin: .25
    }
  }
}, Qk = {
  scenario: `base`,
  a: structuredClone(Zk.base.a),
  rpoFy26: 638,
  rpoQ1Add: 26,
  rp: 84,
  cash: 31.9,
  includeDebt: !1,
  includeAtm: !0,
  atm: 19.9,
  overlay: !1,
  cdsLink: !1,
  linkSites: !0,
  linkLeaseTail: !0,
  useAvgMw: !1,
  billableOpen: 1200,
  cds: 186.275,
  cdsBid: 183.6,
  cdsAsk: 188.95,
  cdsDate: `2026-09-10`,
  m: {
    accepted: [2200, 3800, 5500, 6800, 7800],
    billable: [2e3, 3500, 5100, 6300, 7200],
    util: [95, 94, 93, 92, 90],
    revMW: [.015, .0145, .014, .0135, .013],
    aiShare: [35, 65, 75, 80, 80],
    power: [55, 58, 60, 62, 64],
    pue: [1.18, 1.17, 1.16, 1.15, 1.15],
    maint: [.12, .12, .13, .13, .14],
    defaultP: [.5, .75, 1, 1.5, 2],
    recovery: [60, 55, 50, 45, 40],
    rate: [5.7, 6, 6.2, 6, 5.8]
  },
  terminal: {
    residual: 25,
    rerent: 75,
    margin: 25,
    residualLeaseYears: 12
  },
  sites: structuredClone([{
    id: `abilene`,
    name: `Abilene, Texas`,
    operator: `Oracle / Crusoe`,
    planned: 1200,
    energized: 618,
    accepted: 618,
    billable: 600,
    status: `六棟已交付 618 MW（占八棟校園 75%，校園約 824 MW）；GPT-6 Astra 在此訓練。1.2 GW 視為後續棟`,
    next: `餘下兩棟`,
    date: `FY27 Q2 起`,
    confidence: `高`
  }, {
    id: `shackelford`,
    name: `Shackelford, Texas`,
    operator: `Oracle / Vantage`,
    planned: 1400,
    energized: 115,
    accepted: 0,
    billable: 0,
    status: `下一座 GW 級校園；法說：進度正常。Q1 未交付`,
    next: `首批 Vera Rubin 客戶交付`,
    date: `Q2 FY27`,
    confidence: `中`
  }, {
    id: `dona-ana`,
    name: `Doña Ana, New Mexico`,
    operator: `Oracle / developers`,
    planned: 2200,
    energized: 0,
    accepted: 0,
    billable: 0,
    status: `施工中；Bloom 燃料電池現場發電。空氣許可進行中。Clay：不影響 FY27 營收`,
    next: `空氣許可與首期通電`,
    date: `FY28 起貢獻營收`,
    confidence: `中`
  }, {
    id: `wisconsin`,
    name: `Wisconsin`,
    operator: `Oracle / Vantage`,
    planned: 1e3,
    energized: 0,
    accepted: 0,
    billable: 0,
    status: `機房施工正常；電網與 PSC／ATC／We Energies 設計中。Clay：不影響 FY27 營收`,
    next: `電網核准`,
    date: `FY28 起貢獻營收`,
    confidence: `中`
  }, {
    id: `michigan`,
    name: `Saline Township, Michigan`,
    operator: `Oracle / Related Digital`,
    planned: 1200,
    energized: 0,
    accepted: 0,
    billable: 0,
    status: `施工中（The Barn；三棟）。公告 1+ GW，DTE 1.4 GW。Clay：Q1 未交付`,
    next: `分階段交付 Oracle`,
    date: `2027 起`,
    confidence: `高`
  }])
};

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
    name: `其他／未列名園區`,
    operator: `模型殘差（連動）`,
    planned: Math.max(0, a - n.planned),
    energized: Math.max(0, r - n.energized),
    accepted: Math.max(0, r - n.accepted),
    billable: Math.max(0, i - n.billable),
    status: `用來讓園區加總 = 年度 MW`,
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
      power: [...e.m.power],
      pue: [...e.m.pue],
      maint: [...e.m.maint],
      defaultP: [...e.m.defaultP],
      recovery: [...e.m.recovery],
      rate: [...e.m.rate]
    },
    n = eA(e.sites);
  e.linkSites && (t.accepted[0] = Math.max(t.accepted[0], n.accepted), t.billable[0] = Math.max(t.billable[0], n.billable));
  for (let e = 0; e < 5; e++) t.accepted[e] = Math.max(0, t.accepted[e]), e > 0 && (t.accepted[e] = Math.max(t.accepted[e], t.accepted[e - 1])), t.billable[e] = Math.min(Math.max(0, t.billable[e]), t.accepted[e]), t.util[e] = Math.min(100, Math.max(0, t.util[e])), t.aiShare[e] = Math.min(100, Math.max(0, t.aiShare[e])), t.defaultP[e] = Math.min(100, Math.max(0, t.defaultP[e])), t.recovery[e] = Math.min(100, Math.max(0, t.recovery[e]));
  return t
}

function rA(e) {
  return (Kk[4] + e.a.newLease[4]) * e.terminal.residualLeaseYears
}

function iA(e) {
  return e.a.newLease.reduce((e, t) => e + t, 0) + e.a.newLease[4] * e.terminal.residualLeaseYears
}

function aA(e, t) {
  let n = t.totals.operatingGap,
    r = t.totals.atm,
    i = t.years.reduce((e, t) => e + t.debtPay, 0),
    a = e.cash + n + r - i;
  return {
    open: e.cash,
    operatingGap: n,
    atm: r,
    debtPay: i,
    implied: a,
    end: t.totals.end,
    ok: Math.abs(a - t.totals.end) < .05
  }
}

function oA(e) {
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
  let n = e.rpoFy26,
    r = e.rpoQ1Add,
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
    i = Kk.reduce((e, t) => e + t, 0);
  return {
    uncommenced: Yk.offBalanceLease,
    cashFive: t,
    tail: n,
    mapped: r,
    onFive: i,
    onAfter: Jk,
    gap: r - Yk.offBalanceLease
  }
}

function lA(e) {
  return e >= 500 ? {
    label: `危險`,
    tone: `bad`
  } : e >= 250 ? {
    label: `壓力`,
    tone: `stress`
  } : e > 150 ? {
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
  let r = (e.cds - 150) / 100 * .4;
  return Math.max(0, n + r)
}

function dA(e) {
  let t = nA(e),
    n = tA({
      ...e,
      m: t
    }),
    r = eA(e.sites),
    i = e.rp / 100,
    a = e.cash,
    o = zk.map((n, r) => {
      let o = e.rpoFy26 * (Wk[r] / Uk) * i + e.rpoQ1Add * Gk[r] * i,
        s = o * (t.aiShare[r] / 100),
        c = r === 0 ? e.billableOpen : t.billable[r - 1],
        l = e.useAvgMw ? (c + t.billable[r]) / 2 : t.billable[r],
        u = l * t.revMW[r] * (t.util[r] / 100),
        d = Math.max(0, s - u),
        f = o - d,
        p = 1 - t.recovery[r] / 100,
        m = f * (t.defaultP[r] / 100) * p,
        h = f - m,
        g = h * e.a.margin,
        _ = e.a.capex[r],
        v = _ * e.a.customerFund[r],
        y = _ - v,
        b = Kk[r],
        x = e.a.newLease[r],
        S = b + x,
        C = t.accepted[r] * 8760 * t.pue[r] * t.power[r] / 1e9,
        w = t.accepted[r] * t.maint[r] / 1e3,
        T = e.overlay ? C + w : 0,
        E = Math.max(0, -a) * (uA(e, r) / 100),
        D = e.a.interest[r] + E,
        O = e.includeDebt ? qk[r] : 0,
        k = r === 0 && e.includeAtm ? e.atm : 0,
        A = g + e.a.legacy[r] + v,
        j = _ + S + D + e.a.div[r] + T + O,
        M = A + k,
        N = M - j,
        ee = A - (j - O);
      return a += N, {
        year: n,
        scheduled: o,
        capacity: u,
        ai: s,
        bottleneck: d,
        revenue: f,
        loss: m,
        collected: h,
        rpoCash: g,
        gross: _,
        external: v,
        cashCapex: y,
        lease: S,
        interest: D,
        div: e.a.div[r],
        legacy: e.a.legacy[r],
        atm: k,
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
    s = e => o.reduce((t, n) => t + n[e], 0),
    c = rA(e),
    l = iA(e),
    u = s(`gross`),
    d = s(`collected`),
    f = s(`cashCapex`) + s(`op`) + s(`lease`) + s(`interest`) + s(`div`) + s(`debtPay`),
    p = Math.max(0, f - s(`legacy`)) / Math.max(d, 1),
    m = (e.rpoFy26 + e.rpoQ1Add) * (1 - i),
    h = u * (e.terminal.residual / 100) * (e.terminal.rerent / 100) + m * (e.terminal.margin / 100) - c,
    g = [],
    _ = e => g.push(e);
  _({
    id: `rpo-weights`,
    ok: Math.abs(Wk.reduce((e, t) => e + t, 0) - Uk) < 1e-6,
    severity: `ok`,
    title: `RPO 權重（Q1 10-Q）`,
    detail: `13%+16.65%+20.35%+15.3%+18.7% = 84%，對上 10-Q 五年桶。法說「約一半 36 個月內轉營收」= 13%+37%。`
  }), _({
    id: `lease-10q`,
    ok: !0,
    severity: `ok`,
    title: `在帳租賃`,
    detail: `FY27 用 Q1 已付 1.136 + 剩餘 3.839；FY28–FY31 對上 8/31 到期表（營業+融資）。對應已入帳 ROU，不是 288bn 表外。`
  }), _({
    id: `off-balance-lease`,
    ok: Math.abs(l - Yk.offBalanceLease) < 40,
    severity: `watch`,
    title: `表外機房租賃 $288bn`,
    detail: `10-Q Note 6：已簽約、尚未 commence 的機房租賃承諾 $288bn（未折現）。模型「表外現金租金」五年 + 尾端 = ${l.toFixed(0)}bn，這是現金支付路徑，不是新簽承諾總額、也不是新 ROU。法說沒提這筆。`
  }), _({
    id: `rou-q1`,
    ok: !0,
    severity: `ok`,
    title: `Q1 新取得 ROU $6.4bn`,
    detail: `ROU＝使用權資產。Q1 已起租入帳：營業 4.903 + 融資 1.539。現金租賃只有 1.136。$288bn 尚未 commence，所以還沒有 ROU。`
  });
  let v = o[0].cashCapex,
    y = e.a.capex[0] >= Xk.capexLo - .5 && e.a.capex[0] <= Xk.capexHi + .5;
  _({
    id: `q1-capex`,
    ok: y && v <= Xk.netCashCapexMax + .6,
    severity: y ? `ok` : `watch`,
    title: `FY27 CapEx 90–95／淨現金 ≤70`,
    detail: `法說確認全年毛 capex $90–95bn、淨現金 capex 不超過 $70bn；Q1 毛 28.5、淨約 18，並稱「非線性」。模型 FY27 毛 ${e.a.capex[0]}、淨 ${v.toFixed(1)}。Q1 前載不是把指引作廢。`
  }), _({
    id: `q1-prepay`,
    ok: !0,
    severity: `watch`,
    title: `Q1 客戶預付 $11.4bn`,
    detail: `具重大融資成分的預付 $11.4bn 已進 CFO。此欄是「客戶預付／BYOH」，只降當期 Cash CapEx、形成遞延收入，不降專案總成本，也不是 landlord 租賃（那筆在表外現金租金）。FY27 假設 ${((e.a.customerFund[0]||0)*e.a.capex[0]).toFixed(1)}bn。`
  });
  let b = Xk.fy26Delivered + Xk.mwDelivered;
  _({
    id: `call-mw`,
    ok: !0,
    severity: `watch`,
    title: `Q1 已交付 850 MW（不年化）`,
    detail: `法說 Q1 增量 850 MW（>30 萬 GPU），約為 FY26 全年 73%（≈${Xk.fy26Delivered} MW）。若把剩餘三季用 Q4 節奏往後推，YE 可到 ~${b.toFixed(0)}+ MW——那是觀察，不是預設。基準 Accepted 維持 ${t.accepted[0]} MW（v3.5 路徑）：Clay 點名 Shackelford／NM／WI／Michigan Q1 皆未交付，Q1 爆發主要是 Abilene 補交，不拿來改寫五年存量。`
  }), _({
    id: `call-util`,
    ok: t.util[0] <= Xk.gpuUtil + .05,
    severity: `watch`,
    title: `GPU 利用率 97.9% 是 Q1 時點`,
    detail: `法說 Q1 GPU 艦隊 97.9%，到期 GPU +20% 溢價續約。基準 FY27 ${t.util[0]}%：新產能要爬坡，不當成年均峰值。這是 GPU，不是全部 OCI。`
  });
  for (let e = 0; e < 5; e++) t.billable[e] - t.accepted[e] > .5 && _({
    id: `billable-${e}`,
    ok: !1,
    severity: `block`,
    title: `${zk[e]} 可計費 > 已驗收`,
    detail: `${t.billable[e]} MW > ${t.accepted[e]} MW。引擎已向下截斷。`
  });
  e.linkSites && r.accepted > e.m.accepted[0] + .5 ? _({
    id: `site-floor`,
    ok: !0,
    severity: `watch`,
    title: `園區已驗收拉高 FY27`,
    detail: `具名園區已驗收 ${r.accepted} MW，FY27 Accepted 取 max(輸入, 園區)。`
  }) : !e.linkSites && r.accepted > t.accepted[0] + 1 && _({
    id: `site-mismatch`,
    ok: !1,
    severity: `block`,
    title: `園區與年度 MW 不一致`,
    detail: `具名園區已驗收 ${r.accepted} MW，高於 FY27 Accepted ${t.accepted[0]} MW。請打開園區連動。`
  }), t.accepted[4] > r.planned + 50 && _({
    id: `unlisted-mw`,
    ok: !0,
    severity: `watch`,
    title: `未列名容量`,
    detail: `FY31 Accepted ${t.accepted[4]} MW 高於具名規劃 ${r.planned} MW，差額進「其他／未列名園區」。`
  }), e.overlay && _({
    id: `overlay`,
    ok: !1,
    severity: `watch`,
    title: `Overlay 與貢獻率`,
    detail: `電力+維護已另扣，貢獻率 ${e.a.margin*100}% 應視為尚未扣電費的毛現金率，否則雙重扣除。`
  }), e.cdsLink && _({
    id: `cds-link`,
    ok: !0,
    severity: `watch`,
    title: `CDS 傳入透支利率`,
    detail: `透支利率 = 基準全包利率 + max(0, CDS−150bps)×40%，只在累積現金為負時生效。基準利率若已依當期 CDS 校準，再開傳導會部分重複——兩種模式不要同時當真。`
  }), e.includeDebt && _({
    id: `debt-stress`,
    ok: !0,
    severity: `watch`,
    title: `到期全數還本`,
    detail: `只加本金流出、不含再融資。這是壓力測試，不是基本預測。`
  }), e.a.capex[4] < e.a.capex[0] * .45 && _({
    id: `capex-down`,
    ok: !0,
    severity: `watch`,
    title: `CapEx 速降路徑`,
    detail: `FY31 Gross ${e.a.capex[4]} 相對 FY27 ${e.a.capex[0]}，屬研究假設。法說只確認 FY27 為高峰年之一（與 FY28），未承諾 FY29 起下降。FCF 轉正時點公司拒給。`
  });
  let x = Math.abs(o[4].cum - (e.cash + o.reduce((e, t) => e + t.gap, 0))) < 1e-6,
    S = Math.abs(o[4].cum - (e.cash + s(`operatingGap`) + s(`atm`) - s(`debtPay`))) < .05;
  _({
    id: `identity`,
    ok: x && S,
    severity: x && S ? `ok` : `block`,
    title: `現金恆等式（四層）`,
    detail: x && S ? `期末 ${o[4].cum.toFixed(1)} = 期初 ${e.cash} + 營運缺口 ${s(`operatingGap`).toFixed(1)} + ATM ${s(`atm`).toFixed(1)} − 還本 ${s(`debtPay`).toFixed(1)}。ATM 不算營運來源。` : `累積現金與期初+營運缺口+ATM−還本不一致。`
  });
  let C = r.accepted,
    w = Math.max(0, t.accepted[0] - C),
    T = w / Math.max(t.accepted[0], 1);
  return _({
    id: `residual-mw`,
    ok: T <= .5,
    severity: T > .5 || T > .3 ? `watch` : `ok`,
    title: `未列名園區占 FY27 Accepted ${Math.round(T*100)}%`,
    detail: `具名已驗收 ${C} MW，FY27 Accepted ${t.accepted[0]} MW，殘差 ${w.toFixed(0)} MW。超過 50% 代表容量預測大多不是具名園區支撐。`
  }), e.useAvgMw || _({
    id: `year-end-mw`,
    ok: !0,
    severity: `watch`,
    title: `收入用年末 Billable 全年化`,
    detail: `預設用年末存量×全年單價，年中才交付的 MW 會被當成全年在役。打開「平均在役 MW」改用（期初 ${e.billableOpen} + 期末）/2。`
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
      lease: s(`lease`),
      interest: s(`interest`),
      div: s(`div`),
      atm: s(`atm`),
      be: p,
      terminal: h,
      leaseTail: c,
      scheduled: s(`scheduled`),
      bottleneck: s(`bottleneck`),
      collected: s(`collected`),
      debtPay: s(`debtPay`)
    },
    checks: g,
    sites: n,
    m: t,
    leaseTail: c
  }
}

function fA(e, t) {
  let n = structuredClone(e);
  return t(n), dA(n).totals.end
}

function pA(e) {
  let t = dA(e).totals.end,
    n = [],
    r = (r, i, a) => {
      n.push({
        name: r,
        low: fA(e, i),
        high: fA(e, a),
        base: t
      })
    };
  return r(`RPO 現金貢獻率 ±10pt`, t => {
    t.a.margin = Math.max(.05, e.a.margin - .1)
  }, t => {
    t.a.margin = Math.min(.7, e.a.margin + .1)
  }), r(`非 RPO 現金 ±10bn/年`, e => {
    e.a.legacy = e.a.legacy.map(e => e - 10)
  }, e => {
    e.a.legacy = e.a.legacy.map(e => e + 10)
  }), r(`Gross CapEx ±20%`, e => {
    e.a.capex = e.a.capex.map(e => e * .8)
  }, e => {
    e.a.capex = e.a.capex.map(e => e * 1.2)
  }), r(`客戶資金比率 ±10pt`, e => {
    e.a.customerFund = e.a.customerFund.map(e => Math.max(0, e - .1))
  }, e => {
    e.a.customerFund = e.a.customerFund.map(e => Math.min(.8, e + .1))
  }), r(`Billable MW ±15%`, e => {
    e.m.billable = e.m.billable.map(e => e * .85)
  }, e => {
    e.m.billable = e.m.billable.map(e => e * 1.15)
  }), r(`五年認列比例 70% / 90%`, e => {
    e.rp = 70
  }, e => {
    e.rp = 90
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
var gA = (e, t) => {
    let n = Array(e.length + t.length);
    for (let t = 0; t < e.length; t++) n[t] = e[t];
    for (let r = 0; r < t.length; r++) n[e.length + r] = t[r];
    return n
  },
  _A = (e, t) => ({
    classGroupId: e,
    validator: t
  }),
  vA = (e = new Map, t = null, n) => ({
    nextPart: e,
    validators: t,
    classGroupId: n
  }),
  yA = `-`,
  bA = [],
  xA = `arbitrary..`,
  SA = e => {
    let t = TA(e),
      {
        conflictingClassGroups: n,
        conflictingClassGroupModifiers: r
      } = e;
    return {
      getClassGroupId: e => {
        if (e.startsWith(`[`) && e.endsWith(`]`)) return wA(e);
        let n = e.split(yA);
        return CA(n, +(n[0] === `` && n.length > 1), t)
      },
      getConflictingClassGroupIds: (e, t) => {
        if (t) {
          let t = r[e],
            i = n[e];
          return t ? i ? gA(i, t) : t : i || bA
        }
        return n[e] || bA
      }
    }
  },
  CA = (e, t, n) => {
    if (e.length - t === 0) return n.classGroupId;
    let r = e[t],
      i = n.nextPart.get(r);
    if (i) {
      let n = CA(e, t + 1, i);
      if (n) return n
    }
    let a = n.validators;
    if (a === null) return;
    let o = t === 0 ? e.join(yA) : e.slice(t).join(yA),
      s = a.length;
    for (let e = 0; e < s; e++) {
      let t = a[e];
      if (t.validator(o)) return t.classGroupId
    }
  },
  wA = e => e.slice(1, -1).indexOf(`:`) === -1 ? void 0 : (() => {
    let t = e.slice(1, -1),
      n = t.indexOf(`:`),
      r = t.slice(0, n);
    return r ? xA + r : void 0
  })(),
  TA = e => {
    let {
      theme: t,
      classGroups: n
    } = e;
    return EA(n, t)
  },
  EA = (e, t) => {
    let n = vA();
    for (let r in e) {
      let i = e[r];
      DA(i, n, r, t)
    }
    return n
  },
  DA = (e, t, n, r) => {
    let i = e.length;
    for (let a = 0; a < i; a++) {
      let i = e[a];
      OA(i, t, n, r)
    }
  },
  OA = (e, t, n, r) => {
    if (typeof e == `string`) {
      kA(e, t, n);
      return
    }
    if (typeof e == `function`) {
      AA(e, t, n, r);
      return
    }
    jA(e, t, n, r)
  },
  kA = (e, t, n) => {
    let r = e === `` ? t : MA(t, e);
    r.classGroupId = n
  },
  AA = (e, t, n, r) => {
    if (NA(e)) {
      DA(e(r), t, n, r);
      return
    }
    t.validators === null && (t.validators = []), t.validators.push(_A(n, e))
  },
  jA = (e, t, n, r) => {
    let i = Object.entries(e),
      a = i.length;
    for (let e = 0; e < a; e++) {
      let [a, o] = i[e];
      DA(o, MA(t, a), n, r)
    }
  },
  MA = (e, t) => {
    let n = e,
      r = t.split(yA),
      i = r.length;
    for (let e = 0; e < i; e++) {
      let t = r[e],
        i = n.nextPart.get(t);
      i || (i = vA(), n.nextPart.set(t, i)), n = i
    }
    return n
  },
  NA = e => `isThemeGetter` in e && e.isThemeGetter === !0,
  PA = e => {
    if (e < 1) return {
      get: () => void 0,
      set: () => {}
    };
    let t = 0,
      n = Object.create(null),
      r = Object.create(null),
      i = (i, a) => {
        n[i] = a, t++, t > e && (t = 0, r = n, n = Object.create(null))
      };
    return {
      get(e) {
        let t = n[e];
        if (t !== void 0) return t;
        if ((t = r[e]) !== void 0) return i(e, t), t
      },
      set(e, t) {
        e in n ? n[e] = t : i(e, t)
      }
    }
  },
  FA = `!`,
  IA = `:`,
  LA = [],
  RA = (e, t, n, r, i) => ({
    modifiers: e,
    hasImportantModifier: t,
    baseClassName: n,
    maybePostfixModifierPosition: r,
    isExternal: i
  }),
  zA = e => {
    let {
      prefix: t,
      experimentalParseClassName: n
    } = e, r = e => {
      let t = [],
        n = 0,
        r = 0,
        i = 0,
        a, o = e.length;
      for (let s = 0; s < o; s++) {
        let o = e[s];
        if (n === 0 && r === 0) {
          if (o === IA) {
            t.push(e.slice(i, s)), i = s + 1;
            continue
          }
          if (o === `/`) {
            a = s;
            continue
          }
        }
        o === `[` ? n++ : o === `]` ? n-- : o === `(` ? r++ : o === `)` && r--
      }
      let s = t.length === 0 ? e : e.slice(i),
        c = s,
        l = !1;
      s.endsWith(FA) ? (c = s.slice(0, -1), l = !0) : s.startsWith(FA) && (c = s.slice(1), l = !0);
      let u = a && a > i ? a - i : void 0;
      return RA(t, l, c, u)
    };
    if (t) {
      let e = t + IA,
        n = r;
      r = t => t.startsWith(e) ? n(t.slice(e.length)) : RA(LA, !1, t, void 0, !0)
    }
    if (n) {
      let e = r;
      r = t => n({
        className: t,
        parseClassName: e
      })
    }
    return r
  },
  BA = e => {
    let t = new Map;
    return e.orderSensitiveModifiers.forEach((e, n) => {
      t.set(e, 1e6 + n)
    }), e => {
      let n = [],
        r = [];
      for (let i = 0; i < e.length; i++) {
        let a = e[i],
          o = a[0] === `[`,
          s = t.has(a);
        o || s ? (r.length > 0 && (r.sort(), n.push(...r), r = []), n.push(a)) : r.push(a)
      }
      return r.length > 0 && (r.sort(), n.push(...r)), n
    }
  },
  VA = e => ({
    cache: PA(e.cacheSize),
    parseClassName: zA(e),
    sortModifiers: BA(e),
    postfixLookupClassGroupIds: HA(e),
    ...SA(e)
  }),
  HA = e => {
    let t = Object.create(null),
      n = e.postfixLookupClassGroups;
    if (n)
      for (let e = 0; e < n.length; e++) t[n[e]] = !0;
    return t
  },
  UA = /\s+/,
  WA = (e, t) => {
    let {
      parseClassName: n,
      getClassGroupId: r,
      getConflictingClassGroupIds: i,
      sortModifiers: a,
      postfixLookupClassGroupIds: o
    } = t, s = [], c = e.trim().split(UA), l = ``;
    for (let e = c.length - 1; e >= 0; --e) {
      let t = c[e],
        {
          isExternal: u,
          modifiers: d,
          hasImportantModifier: f,
          baseClassName: p,
          maybePostfixModifierPosition: m
        } = n(t);
      if (u) {
        l = t + (l.length > 0 ? ` ` + l : l);
        continue
      }
      let h = !!m,
        g;
      if (h) {
        g = r(p.substring(0, m));
        let e = g && o[g] ? r(p) : void 0;
        e && e !== g && (g = e, h = !1)
      } else g = r(p);
      if (!g) {
        if (!h) {
          l = t + (l.length > 0 ? ` ` + l : l);
          continue
        }
        if (g = r(p), !g) {
          l = t + (l.length > 0 ? ` ` + l : l);
          continue
        }
        h = !1
      }
      let _ = d.length === 0 ? `` : d.length === 1 ? d[0] : a(d).join(`:`),
        v = f ? _ + FA : _,
        y = v + g;
      if (s.indexOf(y) > -1) continue;
      s.push(y);
      let b = i(g, h);
      for (let e = 0; e < b.length; ++e) {
        let t = b[e];
        s.push(v + t)
      }
      l = t + (l.length > 0 ? ` ` + l : l)
    }
    return l
  },
  GA = (...e) => {
    let t = 0,
      n, r, i = ``;
    for (; t < e.length;)(n = e[t++]) && (r = KA(n)) && (i && (i += ` `), i += r);
    return i
  },
  KA = e => {
    if (typeof e == `string`) return e;
    let t, n = ``;
    for (let r = 0; r < e.length; r++) e[r] && (t = KA(e[r])) && (n && (n += ` `), n += t);
    return n
  },
  qA = (e, ...t) => {
    let n, r, i, a, o = o => (n = VA(t.reduce((e, t) => t(e), e())), r = n.cache.get, i = n.cache.set, a = s, s(o)),
      s = e => {
        let t = r(e);
        if (t) return t;
        let a = WA(e, n);
        return i(e, a), a
      };
    return a = o, (...e) => a(GA(...e))
  },
  JA = [],
  YA = e => {
    let t = t => t[e] || JA;
    return t.isThemeGetter = !0, t
  },
  XA = /^\[(?:(\w[\w-]*):)?(.+)\]$/i,
  ZA = /^\((?:(\w[\w-]*):)?(.+)\)$/i,
  QA = /^\d+(?:\.\d+)?\/\d+(?:\.\d+)?$/,
  $A = /^(\d+(\.\d+)?)?(xs|sm|md|lg|xl)$/,
  ej = /\d+(%|px|r?em|[sdl]?v([hwib]|min|max)|pt|pc|in|cm|mm|cap|ch|ex|r?lh|cq(w|h|i|b|min|max))|\b(calc|min|max|clamp)\(.+\)|^0$/,
  tj = /^(rgba?|hsla?|hwb|(ok)?(lab|lch)|color-mix)\(.+\)$/,
  nj = /^(inset_)?-?((\d+)?\.?(\d+)[a-z]+|0)_-?((\d+)?\.?(\d+)[a-z]+|0)/,
  rj = /^(url|image|image-set|cross-fade|element|(repeating-)?(linear|radial|conic)-gradient)\(.+\)$/,
  ij = e => QA.test(e),
  X = e => !!e && !Number.isNaN(Number(e)),
  aj = e => !!e && Number.isInteger(Number(e)),
  oj = e => e.endsWith(`%`) && X(e.slice(0, -1)),
  sj = e => $A.test(e),
  cj = () => !0,
  lj = e => ej.test(e) && !tj.test(e),
  uj = () => !1,
  dj = e => nj.test(e),
  fj = e => rj.test(e),
  pj = e => !Z(e) && !Q(e),
  mj = e => e.startsWith(`@container`) && (e[10] === `/` && e[11] !== void 0 || e[11] === `s` && e[16] !== void 0 && e.startsWith(`-size/`, 10) || e[11] === `n` && e[18] !== void 0 && e.startsWith(`-normal/`, 10)),
  hj = e => Aj(e, Pj, uj),
  Z = e => XA.test(e),
  gj = e => Aj(e, Fj, lj),
  _j = e => Aj(e, Ij, X),
  vj = e => Aj(e, Rj, cj),
  yj = e => Aj(e, Lj, uj),
  bj = e => Aj(e, Mj, uj),
  xj = e => Aj(e, Nj, fj),
  Sj = e => Aj(e, zj, dj),
  Q = e => ZA.test(e),
  Cj = e => jj(e, Fj),
  wj = e => jj(e, Lj),
  Tj = e => jj(e, Mj),
  Ej = e => jj(e, Pj),
  Dj = e => jj(e, Nj),
  Oj = e => jj(e, zj, !0),
  kj = e => jj(e, Rj, !0),
  Aj = (e, t, n) => {
    let r = XA.exec(e);
    return r ? r[1] ? t(r[1]) : n(r[2]) : !1
  },
  jj = (e, t, n = !1) => {
    let r = ZA.exec(e);
    return r ? r[1] ? t(r[1]) : n : !1
  },
  Mj = e => e === `position` || e === `percentage`,
  Nj = e => e === `image` || e === `url`,
  Pj = e => e === `length` || e === `size` || e === `bg-size`,
  Fj = e => e === `length`,
  Ij = e => e === `number`,
  Lj = e => e === `family-name`,
  Rj = e => e === `number` || e === `weight`,
  zj = e => e === `shadow`,
  Bj = qA(() => {
    let e = YA(`color`),
      t = YA(`font`),
      n = YA(`text`),
      r = YA(`font-weight`),
      i = YA(`tracking`),
      a = YA(`leading`),
      o = YA(`breakpoint`),
      s = YA(`container`),
      c = YA(`spacing`),
      l = YA(`radius`),
      u = YA(`shadow`),
      d = YA(`inset-shadow`),
      f = YA(`text-shadow`),
      p = YA(`drop-shadow`),
      m = YA(`blur`),
      h = YA(`perspective`),
      g = YA(`aspect`),
      _ = YA(`ease`),
      v = YA(`animate`),
      y = () => [`auto`, `avoid`, `all`, `avoid-page`, `page`, `left`, `right`, `column`],
      b = () => [`center`, `top`, `bottom`, `left`, `right`, `top-left`, `left-top`, `top-right`, `right-top`, `bottom-right`, `right-bottom`, `bottom-left`, `left-bottom`],
      x = () => [...b(), Q, Z],
      S = () => [`auto`, `hidden`, `clip`, `visible`, `scroll`],
      C = () => [`auto`, `contain`, `none`],
      w = () => [Q, Z, c],
      T = () => [ij, `full`, `auto`, ...w()],
      E = () => [aj, `none`, `subgrid`, Q, Z],
      D = () => [`auto`, {
        span: [`full`, aj, Q, Z]
      }, aj, Q, Z],
      O = () => [aj, `auto`, Q, Z],
      k = () => [`auto`, `min`, `max`, `fr`, Q, Z],
      A = () => [`start`, `end`, `center`, `between`, `around`, `evenly`, `stretch`, `baseline`, `center-safe`, `end-safe`],
      j = () => [`start`, `end`, `center`, `stretch`, `center-safe`, `end-safe`],
      M = () => [`auto`, ...w()],
      N = () => [ij, `auto`, `full`, `dvw`, `dvh`, `lvw`, `lvh`, `svw`, `svh`, `min`, `max`, `fit`, ...w()],
      ee = () => [ij, `screen`, `full`, `dvw`, `lvw`, `svw`, `min`, `max`, `fit`, ...w()],
      te = () => [ij, `screen`, `full`, `lh`, `dvh`, `lvh`, `svh`, `min`, `max`, `fit`, ...w()],
      P = () => [e, Q, Z],
      F = () => [...b(), Tj, bj, {
        position: [Q, Z]
      }],
      ne = () => [`no-repeat`, {
        repeat: [``, `x`, `y`, `space`, `round`]
      }],
      re = () => [`auto`, `cover`, `contain`, Ej, hj, {
        size: [Q, Z]
      }],
      ie = () => [oj, Cj, gj],
      ae = () => [``, `none`, `full`, l, Q, Z],
      I = () => [``, X, Cj, gj],
      oe = () => [`solid`, `dashed`, `dotted`, `double`],
      se = () => [`normal`, `multiply`, `screen`, `overlay`, `darken`, `lighten`, `color-dodge`, `color-burn`, `hard-light`, `soft-light`, `difference`, `exclusion`, `hue`, `saturation`, `color`, `luminosity`],
      ce = () => [X, oj, Tj, bj],
      le = () => [``, `none`, m, Q, Z],
      ue = () => [`none`, X, Q, Z],
      de = () => [`none`, X, Q, Z],
      fe = () => [X, Q, Z],
      pe = () => [ij, `full`, ...w()];
    return {
      cacheSize: 500,
      theme: {
        animate: [`spin`, `ping`, `pulse`, `bounce`],
        aspect: [`video`],
        blur: [sj],
        breakpoint: [sj],
        color: [cj],
        container: [sj],
        "drop-shadow": [sj],
        ease: [`in`, `out`, `in-out`],
        font: [pj],
        "font-weight": [`thin`, `extralight`, `light`, `normal`, `medium`, `semibold`, `bold`, `extrabold`, `black`],
        "inset-shadow": [sj],
        leading: [`none`, `tight`, `snug`, `normal`, `relaxed`, `loose`],
        perspective: [`dramatic`, `near`, `normal`, `midrange`, `distant`, `none`],
        radius: [sj],
        shadow: [sj],
        spacing: [`px`, X],
        text: [sj],
        "text-shadow": [sj],
        tracking: [`tighter`, `tight`, `normal`, `wide`, `wider`, `widest`]
      },
      classGroups: {
        aspect: [{
          aspect: [`auto`, `square`, ij, Z, Q, g]
        }],
        container: [`container`],
        "container-type": [{
          "@container": [``, `normal`, `size`, Q, Z]
        }],
        "container-named": [mj],
        columns: [{
          columns: [X, Z, Q, s]
        }],
        "break-after": [{
          "break-after": y()
        }],
        "break-before": [{
          "break-before": y()
        }],
        "break-inside": [{
          "break-inside": [`auto`, `avoid`, `avoid-page`, `avoid-column`]
        }],
        "box-decoration": [{
          "box-decoration": [`slice`, `clone`]
        }],
        box: [{
          box: [`border`, `content`]
        }],
        display: [`block`, `inline-block`, `inline`, `flex`, `inline-flex`, `table`, `inline-table`, `table-caption`, `table-cell`, `table-column`, `table-column-group`, `table-footer-group`, `table-header-group`, `table-row-group`, `table-row`, `flow-root`, `grid`, `inline-grid`, `contents`, `list-item`, `hidden`],
        sr: [`sr-only`, `not-sr-only`],
        float: [{
          float: [`right`, `left`, `none`, `start`, `end`]
        }],
        clear: [{
          clear: [`left`, `right`, `both`, `none`, `start`, `end`]
        }],
        isolation: [`isolate`, `isolation-auto`],
        "object-fit": [{
          object: [`contain`, `cover`, `fill`, `none`, `scale-down`]
        }],
        "object-position": [{
          object: x()
        }],
        overflow: [{
          overflow: S()
        }],
        "overflow-x": [{
          "overflow-x": S()
        }],
        "overflow-y": [{
          "overflow-y": S()
        }],
        overscroll: [{
          overscroll: C()
        }],
        "overscroll-x": [{
          "overscroll-x": C()
        }],
        "overscroll-y": [{
          "overscroll-y": C()
        }],
        position: [`static`, `fixed`, `absolute`, `relative`, `sticky`],
        inset: [{
          inset: T()
        }],
        "inset-x": [{
          "inset-x": T()
        }],
        "inset-y": [{
          "inset-y": T()
        }],
        start: [{
          "inset-s": T(),
          start: T()
        }],
        end: [{
          "inset-e": T(),
          end: T()
        }],
        "inset-bs": [{
          "inset-bs": T()
        }],
        "inset-be": [{
          "inset-be": T()
        }],
        top: [{
          top: T()
        }],
        right: [{
          right: T()
        }],
        bottom: [{
          bottom: T()
        }],
        left: [{
          left: T()
        }],
        visibility: [`visible`, `invisible`, `collapse`],
        z: [{
          z: [aj, `auto`, Q, Z]
        }],
        basis: [{
          basis: [ij, `full`, `auto`, s, ...w()]
        }],
        "flex-direction": [{
          flex: [`row`, `row-reverse`, `col`, `col-reverse`]
        }],
        "flex-wrap": [{
          flex: [`nowrap`, `wrap`, `wrap-reverse`]
        }],
        flex: [{
          flex: [X, ij, `auto`, `initial`, `none`, Z]
        }],
        grow: [{
          grow: [``, X, Q, Z]
        }],
        shrink: [{
          shrink: [``, X, Q, Z]
        }],
        order: [{
          order: [aj, `first`, `last`, `none`, Q, Z]
        }],
        "grid-cols": [{
          "grid-cols": E()
        }],
        "col-start-end": [{
          col: D()
        }],
        "col-start": [{
          "col-start": O()
        }],
        "col-end": [{
          "col-end": O()
        }],
        "grid-rows": [{
          "grid-rows": E()
        }],
        "row-start-end": [{
          row: D()
        }],
        "row-start": [{
          "row-start": O()
        }],
        "row-end": [{
          "row-end": O()
        }],
        "grid-flow": [{
          "grid-flow": [`row`, `col`, `dense`, `row-dense`, `col-dense`]
        }],
        "auto-cols": [{
          "auto-cols": k()
        }],
        "auto-rows": [{
          "auto-rows": k()
        }],
        gap: [{
          gap: w()
        }],
        "gap-x": [{
          "gap-x": w()
        }],
        "gap-y": [{
          "gap-y": w()
        }],
        "justify-content": [{
          justify: [...A(), `normal`]
        }],
        "justify-items": [{
          "justify-items": [...j(), `normal`]
        }],
        "justify-self": [{
          "justify-self": [`auto`, ...j()]
        }],
        "align-content": [{
          content: [`normal`, ...A()]
        }],
        "align-items": [{
          items: [...j(), {
            baseline: [``, `last`]
          }]
        }],
        "align-self": [{
          self: [`auto`, ...j(), {
            baseline: [``, `last`]
          }]
        }],
        "place-content": [{
          "place-content": A()
        }],
        "place-items": [{
          "place-items": [...j(), `baseline`]
        }],
        "place-self": [{
          "place-self": [`auto`, ...j()]
        }],
        p: [{
          p: w()
        }],
        px: [{
          px: w()
        }],
        py: [{
          py: w()
        }],
        ps: [{
          ps: w()
        }],
        pe: [{
          pe: w()
        }],
        pbs: [{
          pbs: w()
        }],
        pbe: [{
          pbe: w()
        }],
        pt: [{
          pt: w()
        }],
        pr: [{
          pr: w()
        }],
        pb: [{
          pb: w()
        }],
        pl: [{
          pl: w()
        }],
        m: [{
          m: M()
        }],
        mx: [{
          mx: M()
        }],
        my: [{
          my: M()
        }],
        ms: [{
          ms: M()
        }],
        me: [{
          me: M()
        }],
        mbs: [{
          mbs: M()
        }],
        mbe: [{
          mbe: M()
        }],
        mt: [{
          mt: M()
        }],
        mr: [{
          mr: M()
        }],
        mb: [{
          mb: M()
        }],
        ml: [{
          ml: M()
        }],
        "space-x": [{
          "space-x": w()
        }],
        "space-x-reverse": [`space-x-reverse`],
        "space-y": [{
          "space-y": w()
        }],
        "space-y-reverse": [`space-y-reverse`],
        size: [{
          size: N()
        }],
        "inline-size": [{
          inline: [`auto`, ...ee()]
        }],
        "min-inline-size": [{
          "min-inline": [`auto`, ...ee()]
        }],
        "max-inline-size": [{
          "max-inline": [`none`, ...ee()]
        }],
        "block-size": [{
          block: [`auto`, ...te()]
        }],
        "min-block-size": [{
          "min-block": [`auto`, ...te()]
        }],
        "max-block-size": [{
          "max-block": [`none`, ...te()]
        }],
        w: [{
          w: [s, `screen`, ...N()]
        }],
        "min-w": [{
          "min-w": [s, `screen`, `none`, ...N()]
        }],
        "max-w": [{
          "max-w": [s, `screen`, `none`, `prose`, {
            screen: [o]
          }, ...N()]
        }],
        h: [{
          h: [`screen`, `lh`, ...N()]
        }],
        "min-h": [{
          "min-h": [`screen`, `lh`, `none`, ...N()]
        }],
        "max-h": [{
          "max-h": [`screen`, `lh`, ...N()]
        }],
        "font-size": [{
          text: [`base`, n, Cj, gj]
        }],
        "font-smoothing": [`antialiased`, `subpixel-antialiased`],
        "font-style": [`italic`, `not-italic`],
        "font-weight": [{
          font: [r, kj, vj]
        }],
        "font-stretch": [{
          "font-stretch": [`ultra-condensed`, `extra-condensed`, `condensed`, `semi-condensed`, `normal`, `semi-expanded`, `expanded`, `extra-expanded`, `ultra-expanded`, oj, Z]
        }],
        "font-family": [{
          font: [wj, yj, t]
        }],
        "font-features": [{
          "font-features": [Z]
        }],
        "fvn-normal": [`normal-nums`],
        "fvn-ordinal": [`ordinal`],
        "fvn-slashed-zero": [`slashed-zero`],
        "fvn-figure": [`lining-nums`, `oldstyle-nums`],
        "fvn-spacing": [`proportional-nums`, `tabular-nums`],
        "fvn-fraction": [`diagonal-fractions`, `stacked-fractions`],
        tracking: [{
          tracking: [i, Q, Z]
        }],
        "line-clamp": [{
          "line-clamp": [X, `none`, Q, _j]
        }],
        leading: [{
          leading: [a, ...w()]
        }],
        "list-image": [{
          "list-image": [`none`, Q, Z]
        }],
        "list-style-position": [{
          list: [`inside`, `outside`]
        }],
        "list-style-type": [{
          list: [`disc`, `decimal`, `none`, Q, Z]
        }],
        "text-alignment": [{
          text: [`left`, `center`, `right`, `justify`, `start`, `end`]
        }],
        "placeholder-color": [{
          placeholder: P()
        }],
        "text-color": [{
          text: P()
        }],
        "text-decoration": [`underline`, `overline`, `line-through`, `no-underline`],
        "text-decoration-style": [{
          decoration: [...oe(), `wavy`]
        }],
        "text-decoration-thickness": [{
          decoration: [X, `from-font`, `auto`, Q, gj]
        }],
        "text-decoration-color": [{
          decoration: P()
        }],
        "underline-offset": [{
          "underline-offset": [X, `auto`, Q, Z]
        }],
        "text-transform": [`uppercase`, `lowercase`, `capitalize`, `normal-case`],
        "text-overflow": [`truncate`, `text-ellipsis`, `text-clip`],
        "text-wrap": [{
          text: [`wrap`, `nowrap`, `balance`, `pretty`]
        }],
        indent: [{
          indent: w()
        }],
        "tab-size": [{
          tab: [aj, Q, Z]
        }],
        "vertical-align": [{
          align: [`baseline`, `top`, `middle`, `bottom`, `text-top`, `text-bottom`, `sub`, `super`, Q, Z]
        }],
        whitespace: [{
          whitespace: [`normal`, `nowrap`, `pre`, `pre-line`, `pre-wrap`, `break-spaces`]
        }],
        break: [{
          break: [`normal`, `words`, `all`, `keep`]
        }],
        wrap: [{
          wrap: [`break-word`, `anywhere`, `normal`]
        }],
        hyphens: [{
          hyphens: [`none`, `manual`, `auto`]
        }],
        content: [{
          content: [`none`, Q, Z]
        }],
        "bg-attachment": [{
          bg: [`fixed`, `local`, `scroll`]
        }],
        "bg-clip": [{
          "bg-clip": [`border`, `padding`, `content`, `text`]
        }],
        "bg-origin": [{
          "bg-origin": [`border`, `padding`, `content`]
        }],
        "bg-position": [{
          bg: F()
        }],
        "bg-repeat": [{
          bg: ne()
        }],
        "bg-size": [{
          bg: re()
        }],
        "bg-image": [{
          bg: [`none`, {
            linear: [{
              to: [`t`, `tr`, `r`, `br`, `b`, `bl`, `l`, `tl`]
            }, aj, Q, Z],
            radial: [``, Q, Z],
            conic: [aj, Q, Z]
          }, Dj, xj]
        }],
        "bg-color": [{
          bg: P()
        }],
        "gradient-from-pos": [{
          from: ie()
        }],
        "gradient-via-pos": [{
          via: ie()
        }],
        "gradient-to-pos": [{
          to: ie()
        }],
        "gradient-from": [{
          from: P()
        }],
        "gradient-via": [{
          via: P()
        }],
        "gradient-to": [{
          to: P()
        }],
        rounded: [{
          rounded: ae()
        }],
        "rounded-s": [{
          "rounded-s": ae()
        }],
        "rounded-e": [{
          "rounded-e": ae()
        }],
        "rounded-t": [{
          "rounded-t": ae()
        }],
        "rounded-r": [{
          "rounded-r": ae()
        }],
        "rounded-b": [{
          "rounded-b": ae()
        }],
        "rounded-l": [{
          "rounded-l": ae()
        }],
        "rounded-ss": [{
          "rounded-ss": ae()
        }],
        "rounded-se": [{
          "rounded-se": ae()
        }],
        "rounded-ee": [{
          "rounded-ee": ae()
        }],
        "rounded-es": [{
          "rounded-es": ae()
        }],
        "rounded-tl": [{
          "rounded-tl": ae()
        }],
        "rounded-tr": [{
          "rounded-tr": ae()
        }],
        "rounded-br": [{
          "rounded-br": ae()
        }],
        "rounded-bl": [{
          "rounded-bl": ae()
        }],
        "border-w": [{
          border: I()
        }],
        "border-w-x": [{
          "border-x": I()
        }],
        "border-w-y": [{
          "border-y": I()
        }],
        "border-w-s": [{
          "border-s": I()
        }],
        "border-w-e": [{
          "border-e": I()
        }],
        "border-w-bs": [{
          "border-bs": I()
        }],
        "border-w-be": [{
          "border-be": I()
        }],
        "border-w-t": [{
          "border-t": I()
        }],
        "border-w-r": [{
          "border-r": I()
        }],
        "border-w-b": [{
          "border-b": I()
        }],
        "border-w-l": [{
          "border-l": I()
        }],
        "divide-x": [{
          "divide-x": I()
        }],
        "divide-x-reverse": [`divide-x-reverse`],
        "divide-y": [{
          "divide-y": I()
        }],
        "divide-y-reverse": [`divide-y-reverse`],
        "border-style": [{
          border: [...oe(), `hidden`, `none`]
        }],
        "divide-style": [{
          divide: [...oe(), `hidden`, `none`]
        }],
        "border-color": [{
          border: P()
        }],
        "border-color-x": [{
          "border-x": P()
        }],
        "border-color-y": [{
          "border-y": P()
        }],
        "border-color-s": [{
          "border-s": P()
        }],
        "border-color-e": [{
          "border-e": P()
        }],
        "border-color-bs": [{
          "border-bs": P()
        }],
        "border-color-be": [{
          "border-be": P()
        }],
        "border-color-t": [{
          "border-t": P()
        }],
        "border-color-r": [{
          "border-r": P()
        }],
        "border-color-b": [{
          "border-b": P()
        }],
        "border-color-l": [{
          "border-l": P()
        }],
        "divide-color": [{
          divide: P()
        }],
        "outline-style": [{
          outline: [...oe(), `none`, `hidden`]
        }],
        "outline-offset": [{
          "outline-offset": [X, Q, Z]
        }],
        "outline-w": [{
          outline: [``, X, Cj, gj]
        }],
        "outline-color": [{
          outline: P()
        }],
        shadow: [{
          shadow: [``, `none`, u, Oj, Sj]
        }],
        "shadow-color": [{
          shadow: P()
        }],
        "inset-shadow": [{
          "inset-shadow": [`none`, d, Oj, Sj]
        }],
        "inset-shadow-color": [{
          "inset-shadow": P()
        }],
        "ring-w": [{
          ring: I()
        }],
        "ring-w-inset": [`ring-inset`],
        "ring-color": [{
          ring: P()
        }],
        "ring-offset-w": [{
          "ring-offset": [X, gj]
        }],
        "ring-offset-color": [{
          "ring-offset": P()
        }],
        "inset-ring-w": [{
          "inset-ring": I()
        }],
        "inset-ring-color": [{
          "inset-ring": P()
        }],
        "text-shadow": [{
          "text-shadow": [`none`, f, Oj, Sj]
        }],
        "text-shadow-color": [{
          "text-shadow": P()
        }],
        opacity: [{
          opacity: [X, Q, Z]
        }],
        "mix-blend": [{
          "mix-blend": [...se(), `plus-darker`, `plus-lighter`]
        }],
        "bg-blend": [{
          "bg-blend": se()
        }],
        "mask-clip": [{
          "mask-clip": [`border`, `padding`, `content`, `fill`, `stroke`, `view`]
        }, `mask-no-clip`],
        "mask-composite": [{
          mask: [`add`, `subtract`, `intersect`, `exclude`]
        }],
        "mask-image-linear-pos": [{
          "mask-linear": [X]
        }],
        "mask-image-linear-from-pos": [{
          "mask-linear-from": ce()
        }],
        "mask-image-linear-to-pos": [{
          "mask-linear-to": ce()
        }],
        "mask-image-linear-from-color": [{
          "mask-linear-from": P()
        }],
        "mask-image-linear-to-color": [{
          "mask-linear-to": P()
        }],
        "mask-image-t-from-pos": [{
          "mask-t-from": ce()
        }],
        "mask-image-t-to-pos": [{
          "mask-t-to": ce()
        }],
        "mask-image-t-from-color": [{
          "mask-t-from": P()
        }],
        "mask-image-t-to-color": [{
          "mask-t-to": P()
        }],
        "mask-image-r-from-pos": [{
          "mask-r-from": ce()
        }],
        "mask-image-r-to-pos": [{
          "mask-r-to": ce()
        }],
        "mask-image-r-from-color": [{
          "mask-r-from": P()
        }],
        "mask-image-r-to-color": [{
          "mask-r-to": P()
        }],
        "mask-image-b-from-pos": [{
          "mask-b-from": ce()
        }],
        "mask-image-b-to-pos": [{
          "mask-b-to": ce()
        }],
        "mask-image-b-from-color": [{
          "mask-b-from": P()
        }],
        "mask-image-b-to-color": [{
          "mask-b-to": P()
        }],
        "mask-image-l-from-pos": [{
          "mask-l-from": ce()
        }],
        "mask-image-l-to-pos": [{
          "mask-l-to": ce()
        }],
        "mask-image-l-from-color": [{
          "mask-l-from": P()
        }],
        "mask-image-l-to-color": [{
          "mask-l-to": P()
        }],
        "mask-image-x-from-pos": [{
          "mask-x-from": ce()
        }],
        "mask-image-x-to-pos": [{
          "mask-x-to": ce()
        }],
        "mask-image-x-from-color": [{
          "mask-x-from": P()
        }],
        "mask-image-x-to-color": [{
          "mask-x-to": P()
        }],
        "mask-image-y-from-pos": [{
          "mask-y-from": ce()
        }],
        "mask-image-y-to-pos": [{
          "mask-y-to": ce()
        }],
        "mask-image-y-from-color": [{
          "mask-y-from": P()
        }],
        "mask-image-y-to-color": [{
          "mask-y-to": P()
        }],
        "mask-image-radial": [{
          "mask-radial": [Q, Z]
        }],
        "mask-image-radial-from-pos": [{
          "mask-radial-from": ce()
        }],
        "mask-image-radial-to-pos": [{
          "mask-radial-to": ce()
        }],
        "mask-image-radial-from-color": [{
          "mask-radial-from": P()
        }],
        "mask-image-radial-to-color": [{
          "mask-radial-to": P()
        }],
        "mask-image-radial-shape": [{
          "mask-radial": [`circle`, `ellipse`]
        }],
        "mask-image-radial-size": [{
          "mask-radial": [{
            closest: [`side`, `corner`],
            farthest: [`side`, `corner`]
          }]
        }],
        "mask-image-radial-pos": [{
          "mask-radial-at": b()
        }],
        "mask-image-conic-pos": [{
          "mask-conic": [X]
        }],
        "mask-image-conic-from-pos": [{
          "mask-conic-from": ce()
        }],
        "mask-image-conic-to-pos": [{
          "mask-conic-to": ce()
        }],
        "mask-image-conic-from-color": [{
          "mask-conic-from": P()
        }],
        "mask-image-conic-to-color": [{
          "mask-conic-to": P()
        }],
        "mask-mode": [{
          mask: [`alpha`, `luminance`, `match`]
        }],
        "mask-origin": [{
          "mask-origin": [`border`, `padding`, `content`, `fill`, `stroke`, `view`]
        }],
        "mask-position": [{
          mask: F()
        }],
        "mask-repeat": [{
          mask: ne()
        }],
        "mask-size": [{
          mask: re()
        }],
        "mask-type": [{
          "mask-type": [`alpha`, `luminance`]
        }],
        "mask-image": [{
          mask: [`none`, Q, Z]
        }],
        filter: [{
          filter: [``, `none`, Q, Z]
        }],
        blur: [{
          blur: le()
        }],
        brightness: [{
          brightness: [X, Q, Z]
        }],
        contrast: [{
          contrast: [X, Q, Z]
        }],
        "drop-shadow": [{
          "drop-shadow": [``, `none`, p, Oj, Sj]
        }],
        "drop-shadow-color": [{
          "drop-shadow": P()
        }],
        grayscale: [{
          grayscale: [``, X, Q, Z]
        }],
        "hue-rotate": [{
          "hue-rotate": [X, Q, Z]
        }],
        invert: [{
          invert: [``, X, Q, Z]
        }],
        saturate: [{
          saturate: [X, Q, Z]
        }],
        sepia: [{
          sepia: [``, X, Q, Z]
        }],
        "backdrop-filter": [{
          "backdrop-filter": [``, `none`, Q, Z]
        }],
        "backdrop-blur": [{
          "backdrop-blur": le()
        }],
        "backdrop-brightness": [{
          "backdrop-brightness": [X, Q, Z]
        }],
        "backdrop-contrast": [{
          "backdrop-contrast": [X, Q, Z]
        }],
        "backdrop-grayscale": [{
          "backdrop-grayscale": [``, X, Q, Z]
        }],
        "backdrop-hue-rotate": [{
          "backdrop-hue-rotate": [X, Q, Z]
        }],
        "backdrop-invert": [{
          "backdrop-invert": [``, X, Q, Z]
        }],
        "backdrop-opacity": [{
          "backdrop-opacity": [X, Q, Z]
        }],
        "backdrop-saturate": [{
          "backdrop-saturate": [X, Q, Z]
        }],
        "backdrop-sepia": [{
          "backdrop-sepia": [``, X, Q, Z]
        }],
        "border-collapse": [{
          border: [`collapse`, `separate`]
        }],
        "border-spacing": [{
          "border-spacing": w()
        }],
        "border-spacing-x": [{
          "border-spacing-x": w()
        }],
        "border-spacing-y": [{
          "border-spacing-y": w()
        }],
        "table-layout": [{
          table: [`auto`, `fixed`]
        }],
        caption: [{
          caption: [`top`, `bottom`]
        }],
        transition: [{
          transition: [``, `all`, `colors`, `opacity`, `shadow`, `transform`, `none`, Q, Z]
        }],
        "transition-behavior": [{
          transition: [`normal`, `discrete`]
        }],
        duration: [{
          duration: [X, `initial`, Q, Z]
        }],
        ease: [{
          ease: [`linear`, `initial`, _, Q, Z]
        }],
        delay: [{
          delay: [X, Q, Z]
        }],
        animate: [{
          animate: [`none`, v, Q, Z]
        }],
        backface: [{
          backface: [`hidden`, `visible`]
        }],
        perspective: [{
          perspective: [h, Q, Z]
        }],
        "perspective-origin": [{
          "perspective-origin": x()
        }],
        rotate: [{
          rotate: ue()
        }],
        "rotate-x": [{
          "rotate-x": ue()
        }],
        "rotate-y": [{
          "rotate-y": ue()
        }],
        "rotate-z": [{
          "rotate-z": ue()
        }],
        scale: [{
          scale: de()
        }],
        "scale-x": [{
          "scale-x": de()
        }],
        "scale-y": [{
          "scale-y": de()
        }],
        "scale-z": [{
          "scale-z": de()
        }],
        "scale-3d": [`scale-3d`],
        skew: [{
          skew: fe()
        }],
        "skew-x": [{
          "skew-x": fe()
        }],
        "skew-y": [{
          "skew-y": fe()
        }],
        transform: [{
          transform: [Q, Z, ``, `none`, `gpu`, `cpu`]
        }],
        "transform-origin": [{
          origin: x()
        }],
        "transform-style": [{
          transform: [`3d`, `flat`]
        }],
        translate: [{
          translate: pe()
        }],
        "translate-x": [{
          "translate-x": pe()
        }],
        "translate-y": [{
          "translate-y": pe()
        }],
        "translate-z": [{
          "translate-z": pe()
        }],
        "translate-none": [`translate-none`],
        zoom: [{
          zoom: [aj, Q, Z]
        }],
        accent: [{
          accent: P()
        }],
        appearance: [{
          appearance: [`none`, `auto`]
        }],
        "caret-color": [{
          caret: P()
        }],
        "color-scheme": [{
          scheme: [`normal`, `dark`, `light`, `light-dark`, `only-dark`, `only-light`]
        }],
        cursor: [{
          cursor: [`auto`, `default`, `pointer`, `wait`, `text`, `move`, `help`, `not-allowed`, `none`, `context-menu`, `progress`, `cell`, `crosshair`, `vertical-text`, `alias`, `copy`, `no-drop`, `grab`, `grabbing`, `all-scroll`, `col-resize`, `row-resize`, `n-resize`, `e-resize`, `s-resize`, `w-resize`, `ne-resize`, `nw-resize`, `se-resize`, `sw-resize`, `ew-resize`, `ns-resize`, `nesw-resize`, `nwse-resize`, `zoom-in`, `zoom-out`, Q, Z]
        }],
        "field-sizing": [{
          "field-sizing": [`fixed`, `content`]
        }],
        "pointer-events": [{
          "pointer-events": [`auto`, `none`]
        }],
        resize: [{
          resize: [`none`, ``, `y`, `x`]
        }],
        "scroll-behavior": [{
          scroll: [`auto`, `smooth`]
        }],
        "scrollbar-thumb-color": [{
          "scrollbar-thumb": P()
        }],
        "scrollbar-track-color": [{
          "scrollbar-track": P()
        }],
        "scrollbar-gutter": [{
          "scrollbar-gutter": [`auto`, `stable`, `both`]
        }],
        "scrollbar-w": [{
          scrollbar: [`auto`, `thin`, `none`]
        }],
        "scroll-m": [{
          "scroll-m": w()
        }],
        "scroll-mx": [{
          "scroll-mx": w()
        }],
        "scroll-my": [{
          "scroll-my": w()
        }],
        "scroll-ms": [{
          "scroll-ms": w()
        }],
        "scroll-me": [{
          "scroll-me": w()
        }],
        "scroll-mbs": [{
          "scroll-mbs": w()
        }],
        "scroll-mbe": [{
          "scroll-mbe": w()
        }],
        "scroll-mt": [{
          "scroll-mt": w()
        }],
        "scroll-mr": [{
          "scroll-mr": w()
        }],
        "scroll-mb": [{
          "scroll-mb": w()
        }],
        "scroll-ml": [{
          "scroll-ml": w()
        }],
        "scroll-p": [{
          "scroll-p": w()
        }],
        "scroll-px": [{
          "scroll-px": w()
        }],
        "scroll-py": [{
          "scroll-py": w()
        }],
        "scroll-ps": [{
          "scroll-ps": w()
        }],
        "scroll-pe": [{
          "scroll-pe": w()
        }],
        "scroll-pbs": [{
          "scroll-pbs": w()
        }],
        "scroll-pbe": [{
          "scroll-pbe": w()
        }],
        "scroll-pt": [{
          "scroll-pt": w()
        }],
        "scroll-pr": [{
          "scroll-pr": w()
        }],
        "scroll-pb": [{
          "scroll-pb": w()
        }],
        "scroll-pl": [{
          "scroll-pl": w()
        }],
        "snap-align": [{
          snap: [`start`, `end`, `center`, `align-none`]
        }],
        "snap-stop": [{
          snap: [`normal`, `always`]
        }],
        "snap-type": [{
          snap: [`none`, `x`, `y`, `both`]
        }],
        "snap-strictness": [{
          snap: [`mandatory`, `proximity`]
        }],
        touch: [{
          touch: [`auto`, `none`, `manipulation`]
        }],
        "touch-x": [{
          "touch-pan": [`x`, `left`, `right`]
        }],
        "touch-y": [{
          "touch-pan": [`y`, `up`, `down`]
        }],
        "touch-pz": [`touch-pinch-zoom`],
        select: [{
          select: [`none`, `text`, `all`, `auto`]
        }],
        "will-change": [{
          "will-change": [`auto`, `scroll`, `contents`, `transform`, Q, Z]
        }],
        fill: [{
          fill: [`none`, ...P()]
        }],
        "stroke-w": [{
          stroke: [X, Cj, gj, _j]
        }],
        stroke: [{
          stroke: [`none`, ...P()]
        }],
        "forced-color-adjust": [{
          "forced-color-adjust": [`auto`, `none`]
        }]
      },
      conflictingClassGroups: {
        "container-named": [`container-type`],
        overflow: [`overflow-x`, `overflow-y`],
        overscroll: [`overscroll-x`, `overscroll-y`],
        inset: [`inset-x`, `inset-y`, `inset-bs`, `inset-be`, `start`, `end`, `top`, `right`, `bottom`, `left`],
        "inset-x": [`right`, `left`],
        "inset-y": [`top`, `bottom`],
        flex: [`basis`, `grow`, `shrink`],
        gap: [`gap-x`, `gap-y`],
        p: [`px`, `py`, `ps`, `pe`, `pbs`, `pbe`, `pt`, `pr`, `pb`, `pl`],
        px: [`pr`, `pl`],
        py: [`pt`, `pb`],
        m: [`mx`, `my`, `ms`, `me`, `mbs`, `mbe`, `mt`, `mr`, `mb`, `ml`],
        mx: [`mr`, `ml`],
        my: [`mt`, `mb`],
        size: [`w`, `h`],
        "font-size": [`leading`],
        "fvn-normal": [`fvn-ordinal`, `fvn-slashed-zero`, `fvn-figure`, `fvn-spacing`, `fvn-fraction`],
        "fvn-ordinal": [`fvn-normal`],
        "fvn-slashed-zero": [`fvn-normal`],
        "fvn-figure": [`fvn-normal`],
        "fvn-spacing": [`fvn-normal`],
        "fvn-fraction": [`fvn-normal`],
        "line-clamp": [`display`, `overflow`],
        rounded: [`rounded-s`, `rounded-e`, `rounded-t`, `rounded-r`, `rounded-b`, `rounded-l`, `rounded-ss`, `rounded-se`, `rounded-ee`, `rounded-es`, `rounded-tl`, `rounded-tr`, `rounded-br`, `rounded-bl`],
        "rounded-s": [`rounded-ss`, `rounded-es`],
        "rounded-e": [`rounded-se`, `rounded-ee`],
        "rounded-t": [`rounded-tl`, `rounded-tr`],
        "rounded-r": [`rounded-tr`, `rounded-br`],
        "rounded-b": [`rounded-br`, `rounded-bl`],
        "rounded-l": [`rounded-tl`, `rounded-bl`],
        "border-spacing": [`border-spacing-x`, `border-spacing-y`],
        "border-w": [`border-w-x`, `border-w-y`, `border-w-s`, `border-w-e`, `border-w-bs`, `border-w-be`, `border-w-t`, `border-w-r`, `border-w-b`, `border-w-l`],
        "border-w-x": [`border-w-r`, `border-w-l`],
        "border-w-y": [`border-w-t`, `border-w-b`],
        "border-color": [`border-color-x`, `border-color-y`, `border-color-s`, `border-color-e`, `border-color-bs`, `border-color-be`, `border-color-t`, `border-color-r`, `border-color-b`, `border-color-l`],
        "border-color-x": [`border-color-r`, `border-color-l`],
        "border-color-y": [`border-color-t`, `border-color-b`],
        translate: [`translate-x`, `translate-y`, `translate-none`],
        "translate-none": [`translate`, `translate-x`, `translate-y`, `translate-z`],
        "scroll-m": [`scroll-mx`, `scroll-my`, `scroll-ms`, `scroll-me`, `scroll-mbs`, `scroll-mbe`, `scroll-mt`, `scroll-mr`, `scroll-mb`, `scroll-ml`],
        "scroll-mx": [`scroll-mr`, `scroll-ml`],
        "scroll-my": [`scroll-mt`, `scroll-mb`],
        "scroll-p": [`scroll-px`, `scroll-py`, `scroll-ps`, `scroll-pe`, `scroll-pbs`, `scroll-pbe`, `scroll-pt`, `scroll-pr`, `scroll-pb`, `scroll-pl`],
        "scroll-px": [`scroll-pr`, `scroll-pl`],
        "scroll-py": [`scroll-pt`, `scroll-pb`],
        touch: [`touch-x`, `touch-y`, `touch-pz`],
        "touch-x": [`touch`],
        "touch-y": [`touch`],
        "touch-pz": [`touch`]
      },
      conflictingClassGroupModifiers: {
        "font-size": [`leading`]
      },
      postfixLookupClassGroups: [`container-type`],
      orderSensitiveModifiers: [`*`, `**`, `after`, `backdrop`, `before`, `details-content`, `file`, `first-letter`, `first-line`, `marker`, `placeholder`, `selection`]
    }
  });

function Vj(...e) {
  return Bj(x(e))
}
var Hj = o((e => {
    var t = Symbol.for(`react.transitional.element`);

    function n(e, n, r) {
      var i = null;
      if (r !== void 0 && (i = `` + r), n.key !== void 0 && (i = `` + n.key), `key` in n)
        for (var a in r = {}, n) a !== `key` && (r[a] = n[a]);
      else r = n;
      return n = r.ref, {
        $$typeof: t,
        type: e,
        key: i,
        ref: n === void 0 ? null : n,
        props: r
      }
    }
    e.jsx = n, e.jsxs = n
  })),
  $ = o(((e, t) => {
    t.exports = Hj()
  }))();

function Uj({
  className: e,
  tone: t = `neutral`,
  children: n
}) {
  return (0, $.jsx)(`span`, {
    className: Vj(`inline-flex items-center rounded-full border px-2.5 py-0.5 text-xs font-medium`, {
      neutral: `bg-surface text-muted border-border`,
      ok: `bg-ok-bg text-ok border-ok/20`,
      watch: `bg-watch-bg text-watch border-watch/20`,
      stress: `bg-watch-bg text-stress border-stress/20`,
      bad: `bg-bad-bg text-bad border-bad/20`
    } [t], e),
    children: n
  })
}
var Wj = e => typeof e == `boolean` ? `${e}` : e === 0 ? `0` : e,
  Gj = x,
  Kj = ((e, t) => n => {
    if (t?.variants == null) return Gj(e, n?.class, n?.className);
    let {
      variants: r,
      defaultVariants: i
    } = t, a = Object.keys(r).map(e => {
      let t = n?.[e],
        a = i?.[e];
      if (t === null) return null;
      let o = Wj(t) || Wj(a);
      return r[e][o]
    }), o = n && Object.entries(n).reduce((e, t) => {
      let [n, r] = t;
      return r === void 0 || (e[n] = r), e
    }, {});
    return Gj(e, a, t?.compoundVariants?.reduce((e, t) => {
      let {
        class: n,
        className: r,
        ...a
      } = t;
      return Object.entries(a).every(e => {
        let [t, n] = e;
        return Array.isArray(n) ? n.includes({
          ...i,
          ...o
        } [t]) : {
          ...i,
          ...o
        } [t] === n
      }) ? [...e, n, r] : e
    }, []), n?.class, n?.className)
  })(`inline-flex items-center justify-center gap-2 rounded-md text-sm font-medium transition-colors duration-150 disabled:pointer-events-none disabled:opacity-40 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/40`, {
    variants: {
      variant: {
        default: `bg-accent text-accent-fg hover:bg-ink`,
        outline: `border border-border bg-card text-fg hover:bg-accent-soft`,
        ghost: `text-fg hover:bg-accent-soft`,
        solid: `bg-ink text-accent-fg hover:bg-fg`
      },
      size: {
        default: `h-10 px-3.5`,
        sm: `h-8 px-2.5 text-xs`,
        lg: `h-11 px-4`
      }
    },
    defaultVariants: {
      variant: `default`,
      size: `default`
    }
  });

function qj({
  className: e,
  variant: t,
  size: n,
  ...r
}) {
  return (0, $.jsx)(`button`, {
    className: Vj(Kj({
      variant: t,
      size: n
    }), e),
    ...r
  })
}

function Jj({
  className: e,
  ...t
}) {
  return (0, $.jsx)(`div`, {
    className: Vj(`rounded-xl border border-border bg-card p-4 shadow-card md:p-5`, e),
    ...t
  })
}
var Yj = e => e.replace(/([a-z0-9])([A-Z])/g, `$1-$2`).toLowerCase(),
  Xj = e => e.replace(/^([A-Z])|[\s-_]+(\w)/g, (e, t, n) => n ? n.toUpperCase() : t.toLowerCase()),
  Zj = e => {
    let t = Xj(e);
    return t.charAt(0).toUpperCase() + t.slice(1)
  },
  Qj = (...e) => e.filter((e, t, n) => !!e && e.trim() !== `` && n.indexOf(e) === t).join(` `).trim(),
  $j = e => {
    for (let t in e)
      if (t.startsWith(`aria-`) || t === `role` || t === `title`) return !0
  },
  eM = {
    xmlns: `http://www.w3.org/2000/svg`,
    width: 24,
    height: 24,
    viewBox: `0 0 24 24`,
    fill: `none`,
    stroke: `currentColor`,
    strokeWidth: 2,
    strokeLinecap: `round`,
    strokeLinejoin: `round`
  },
  tM = (0, v.forwardRef)(({
    color: e = `currentColor`,
    size: t = 24,
    strokeWidth: n = 2,
    absoluteStrokeWidth: r,
    className: i = ``,
    children: a,
    iconNode: o,
    ...s
  }, c) => (0, v.createElement)(`svg`, {
    ref: c,
    ...eM,
    width: t,
    height: t,
    stroke: e,
    strokeWidth: r ? Number(n) * 24 / Number(t) : n,
    className: Qj(`lucide`, i),
    ...!a && !$j(s) && {
      "aria-hidden": `true`
    },
    ...s
  }, [...o.map(([e, t]) => (0, v.createElement)(e, t)), ...Array.isArray(a) ? a : [a]])),
  nM = (e, t) => {
    let n = (0, v.forwardRef)(({
      className: n,
      ...r
    }, i) => (0, v.createElement)(tM, {
      ref: i,
      iconNode: t,
      className: Qj(`lucide-${Yj(Zj(e))}`, `lucide-${e}`, n),
      ...r
    }));
    return n.displayName = Zj(e), n
  },
  rM = nM(`download`, [
    [`path`, {
      d: `M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4`,
      key: `ih7n3h`
    }],
    [`polyline`, {
      points: `7 10 12 15 17 10`,
      key: `2ggqvy`
    }],
    [`line`, {
      x1: `12`,
      x2: `12`,
      y1: `15`,
      y2: `3`,
      key: `1vk2je`
    }]
  ]),
  iM = nM(`rotate-ccw`, [
    [`path`, {
      d: `M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8`,
      key: `1357e3`
    }],
    [`path`, {
      d: `M3 3v5h5`,
      key: `1xhq8a`
    }]
  ]),
  aM = nM(`upload`, [
    [`path`, {
      d: `M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4`,
      key: `ih7n3h`
    }],
    [`polyline`, {
      points: `17 8 12 3 7 8`,
      key: `t8dd8p`
    }],
    [`line`, {
      x1: `12`,
      x2: `12`,
      y1: `3`,
      y2: `15`,
      key: `widbto`
    }]
  ]),
  oM = [`FY27`, `FY28`, `FY29`, `FY30`, `FY31`],
  sM = [{
    year: `FY22`,
    oci: 2.92,
    apps: 7.89,
    software: 25.24,
    other: 6.39,
    revenue: 42.44,
    opInc: 10.96,
    ni: 6.72,
    eps: 2.41,
    ngEps: 4.21,
    shares: 2.786,
    ociTag: `Reconstructed`
  }, {
    year: `FY23`,
    oci: 4.56,
    apps: 11.32,
    software: 25.21,
    other: 8.87,
    revenue: 49.954,
    opInc: 13.09,
    ni: 8.5,
    eps: 3.07,
    ngEps: 4.89,
    shares: 2.766,
    ociTag: `Reconstructed`
  }, {
    year: `FY24`,
    oci: 6.84,
    apps: 12.934,
    software: 24.69,
    other: 8.5,
    revenue: 52.961,
    opInc: 15.353,
    ni: 10.467,
    eps: 3.71,
    ngEps: 5.56,
    shares: 2.823,
    ociTag: `Verified`
  }, {
    year: `FY25`,
    oci: 10.234,
    apps: 14.272,
    software: 24.724,
    other: 8.169,
    revenue: 57.399,
    opInc: 17.678,
    ni: 12.443,
    eps: 4.34,
    ngEps: 6.03,
    shares: 2.866,
    ociTag: `Verified`
  }, {
    year: `FY26`,
    oci: 18.101,
    apps: 15.888,
    software: 24.541,
    other: 8.827,
    revenue: 67.357,
    opInc: 20.606,
    ni: 16.984,
    eps: 5.83,
    ngEps: 7.63,
    shares: 2.914,
    ociTag: `Verified`
  }],
  cM = {
    price: 152.94,
    shares: 3.024,
    netDebt: 97.6,
    tax: .151,
    wacc: .095,
    g: .03,
    appsG: .1,
    softwareG: -.01,
    otherG: .03,
    otherOci: [12, 14, 16, 18, 20],
    opMargin: [.32, .285, .28, .295, .31],
    daOci: .18,
    daOther: .04,
    sbc: 4.5,
    tradEvSales: 5.5,
    ociEvSales: 8.5,
    ntmPe: 22
  },
  lM = [{
    ticker: `MSFT`,
    name: `Microsoft`,
    role: `雲＋應用上沿`,
    pe: 27.5,
    fwdPe: 24.9,
    evSales: 12.1,
    mkt: 3650
  }, {
    ticker: `AMZN`,
    name: `Amazon`,
    role: `AWS／IaaS`,
    pe: 20.3,
    fwdPe: 24,
    evSales: 3.4,
    mkt: 2720
  }, {
    ticker: `GOOGL`,
    name: `Alphabet`,
    role: `GCP／廣告現金`,
    pe: 16.6,
    fwdPe: 22.5,
    evSales: 6.2,
    mkt: 4040
  }, {
    ticker: `CRM`,
    name: `Salesforce`,
    role: `企業 SaaS`,
    pe: 22.3,
    fwdPe: 20,
    evSales: 5,
    mkt: 200
  }, {
    ticker: `SAP`,
    name: `SAP`,
    role: `企業應用`,
    pe: 26,
    fwdPe: 24,
    evSales: 6.6,
    mkt: 254
  }, {
    ticker: `NOW`,
    name: `ServiceNow`,
    role: `高成長 SaaS`,
    pe: 48,
    fwdPe: 38,
    evSales: 12.5,
    mkt: 168
  }];

function uM(e, t) {
  return t ? e / t - 1 : 0
}

function dM(e) {
  return Math.max(0, -e)
}
var fM = .141;

function pM(e, t) {
  let n = e.includeAtm ? 0 : fM;
  return Math.max(.5, t.shares - n)
}

function mM(e, t, n) {
  return e - t / Math.max(n, .01)
}

function hM(e, t) {
  let n = sM[4],
    r = n.revenue,
    i = n.apps,
    a = n.software,
    o = n.other,
    s = Math.max(t.shares, n.shares * .9);
  return oM.map((n, c) => {
    let l = e.years[c].capacity,
      u = l + t.otherOci[c];
    i *= 1 + t.appsG, a *= 1 + t.softwareG, o *= 1 + t.otherG;
    let d = u + i + a + o,
      f = uM(d, r),
      p = d * t.opMargin[c],
      m = e.years[c].interest,
      h = (p - m) * (1 - t.tax);
    s = c === 0 ? t.shares : s * 1.01;
    let g = h / s,
      _ = (h + t.sbc * (1 - t.tax)) / s,
      v = u * t.daOci + (d - u) * t.daOther,
      y = e.years[c].cashCapex,
      b = .02 * (d - r),
      x = p * (1 - t.tax) + v - y - b,
      S = e.years[c].revenue,
      C = o,
      w = d - S,
      T = d > 0 ? Math.abs(w) / d : 0;
    return r = d, {
      year: n,
      gpu: l,
      oci: u,
      apps: i,
      software: a,
      other: o,
      revenue: d,
      fundingRev: S,
      nonRpo: C,
      inYear: w,
      plugRatio: T,
      yoy: f,
      opInc: p,
      opM: t.opMargin[c],
      ni: h,
      nm: h / d,
      eps: g,
      ngEps: _,
      shares: s,
      da: v,
      cashCapex: y,
      interest: m,
      ufcf: x
    }
  })
}

function gM(e, t) {
  let n = e.map(e => e.ufcf),
    r = n.map((e, n) => e / (1 + t.wacc) ** (n + 1)),
    i = r.reduce((e, t) => e + t, 0),
    a = n[n.length - 1] * (1 + t.g) / (t.wacc - t.g),
    o = a / (1 + t.wacc) ** n.length,
    s = i + o,
    c = s - t.netDebt;
  return {
    pvFcf: i,
    tv: a,
    pvTv: o,
    ev: s,
    equity: c,
    perShare: c / t.shares,
    tvShare: o / Math.max(Math.abs(s), 1) * Math.sign(s || 1),
    fcf: n,
    pv: r
  }
}

function _M(e, t) {
  let n = e[0],
    r = n.oci * t.ociEvSales,
    i = (n.revenue - n.oci) * t.tradEvSales,
    a = r + i,
    o = a - t.netDebt;
  return {
    ociEv: r,
    tradEv: i,
    ev: a,
    equity: o,
    perShare: o / t.shares
  }
}

function vM(e, t) {
  return e[0].ngEps * t.ntmPe
}
var yM = {
    dcf: .45,
    sotp: .35,
    pe: .2
  },
  bM = [{
    method: `DCF`,
    capex: `已扣 Cash CapEx`,
    hole: `不另扣未籌缺口`,
    shares: `含 ATM 後股數`,
    netDebt: `FY26 末 97.6，不含表外未起租`
  }, {
    method: `SOTP`,
    capex: `倍數不含建置現金`,
    hole: `另扣未籌缺口／股數`,
    shares: `含 ATM 後股數`,
    netDebt: `FY26 末 97.6，不含表外未起租`
  }, {
    method: `NTM PE`,
    capex: `EPS 未扣未來 Cash CapEx`,
    hole: `另扣未籌缺口／股數`,
    shares: `含 ATM 後股數`,
    netDebt: `不進 PE`
  }];

function xM(e) {
  let {
    spot: t,
    dcfPx: n,
    sotpPx: r,
    pePx: i,
    endCash: a,
    tvShare: o
  } = e, s = yM.dcf * n + yM.sotp * r + yM.pe * i, c = s / Math.max(t, .01) - 1, l = dM(a), u = a >= 0, d = [];
  u ? d.push(`資金狀態 Funded：五年期末現金 ${a.toFixed(0)}bn，評價與資金同向。`) : d.push(`資金狀態 Unfunded：五年期末現金 ${a.toFixed(0)}bn，尚待籌約 ${l.toFixed(0)}bn。這是必須定價的風險，不是拿來代替估值的開關；但未指定融資前不准買進（避免評價說買、資金說破產）。SOTP／PE 已扣缺口；DCF 已內含 Cash CapEx。`), o > .85 && d.push(`終值占企業價值過高，DCF 對 WACC／永續成長極敏感。`), !u && n < t * .75 && d.push(`DCF（內含建置支出）明顯低於現價，與資金缺口方向一致。`);
  let f;
  f = u ? c >= .25 && o < .9 ? `買進` : c <= -.15 && n < t ? `賣出` : `中立` : a < -40 || c <= -.08 || n < t * .9 ? `賣出` : `中立`;
  let p = Math.min(n, r, i) * .9,
    m = Math.max(n, r, i) * 1.05;
  return {
    blended: s,
    upside: c,
    call: f,
    lo: p,
    hi: m,
    notes: d,
    hole: l,
    funded: u,
    weights: {
      ...yM
    }
  }
}

function SM(e, t) {
  let n = [.08, .09, .095, .1, .11];
  return [.02, .025, .03, .035, .04].map(r => n.map(n => n <= r ? NaN : gM(e, {
    ...t,
    wacc: n,
    g: r
  }).perShare))
}

function CM(e, t, n = 0) {
  let r = [6, 7.5, 8.5, 10, 12],
    i = [4, 5, 5.5, 6.5, 8],
    a = n / Math.max(t.shares, .01);
  return i.map(n => r.map(r => _M(e, {
    ...t,
    tradEvSales: n,
    ociEvSales: r
  }).perShare - a))
}

function wM(e) {
  let t = [...e].sort((e, t) => e - t),
    n = Math.floor(t.length / 2);
  return t.length % 2 ? t[n] : (t[n - 1] + t[n]) / 2
}

function TM(e, t, n) {
  let r = pM(t, n),
    i = {
      ...n,
      shares: r
    },
    a = hM(e, i),
    o = gM(a, i),
    s = _M(a, i),
    c = vM(a, i),
    l = dM(e.totals.end),
    u = mM(s.perShare, l, r),
    d = mM(c, l, r),
    f = xM({
      spot: i.price,
      dcfPx: o.perShare,
      sotpPx: u,
      pePx: d,
      endCash: e.totals.end,
      tvShare: o.tvShare
    });
  return {
    v: i,
    fwd: a,
    d: o,
    s,
    sRaw: s.perShare,
    sAdj: u,
    peRaw: c,
    peAdj: d,
    hole: l,
    call: f,
    shares: r
  }
}

function EM(e, t, n) {
  let r = e.fwd[0],
    i = e.v.otherOci[0],
    a = Math.abs(r.oci - (r.gpu + i)) < .05,
    o = e.fwd.every((e, n) => Math.abs(e.cashCapex - t.years[n].cashCapex) < 1e-6),
    s = e.fwd.every((e, n) => Math.abs(e.interest - t.years[n].interest) < 1e-6),
    c = t.totals.end >= 0 || e.call.call !== `買進`,
    l = n.includeAtm || e.shares < cM.shares - .05;
  return [{
    id: `val-oci`,
    ok: a,
    severity: a ? `ok` : `block`,
    title: `OCI 連動產能`,
    detail: `FY27 GPU 產能 ${r.gpu.toFixed(1)} + 非GPU ${i.toFixed(1)} = OCI ${r.oci.toFixed(1)}bn。改 Billable／利用率／rev/MW 會改這列。`
  }, {
    id: `val-capex`,
    ok: o,
    severity: o ? `ok` : `block`,
    title: `DCF Cash CapEx 連動`,
    detail: `五年 Cash CapEx ${t.totals.cashCapex.toFixed(0)}bn 與損益／DCF 為同一組數字。`
  }, {
    id: `val-int`,
    ok: s,
    severity: s ? `ok` : `block`,
    title: `利息連動`,
    detail: `損益表利息（含現金為負時的循環利息、CDS 連動加碼）取自資金模型。`
  }, {
    id: `val-veto`,
    ok: c,
    severity: c ? e.call.funded ? `ok` : `watch` : `block`,
    title: `資金缺口禁止買進`,
    detail: e.call.funded ? `期末現金 ${t.totals.end.toFixed(0)}bn，評價結論「${e.call.call}」與資金同向。` : `期末現金 ${t.totals.end.toFixed(0)}bn，缺口 ${e.hole.toFixed(0)}bn 已扣進 SOTP／PE。結論「${e.call.call}」，不會是買進。`
  }, {
    id: `val-atm`,
    ok: l,
    severity: `ok`,
    title: `ATM 與股數`,
    detail: n.includeAtm ? `含 ATM，評價股數 ${e.shares.toFixed(2)}bn（10-Q 流通 3.024bn）。` : `ATM 關閉，股數由 ${cM.shares.toFixed(2)} 扣回 141m 股為 ${e.shares.toFixed(2)}bn，缺口同步擴大。`
  }, {
    id: `val-rev-guide`,
    ok: r.revenue + .05 >= 90,
    severity: r.revenue >= 90 ? `ok` : `watch`,
    title: `FY27 營收指引 ≥$90bn`,
    detail: `法說上修全年「至少 $90bn」。當下限、不倒推產能。本模型 FY27 IS ${r.revenue.toFixed(1)}bn。n/g EPS 指引 $8.10；模型 ${r.ngEps.toFixed(2)}（GAAP 加回 SBC，不是公司 non-GAAP）。`
  }, {
    id: `val-plug`,
    ok: r.plugRatio <= .1,
    severity: r.plugRatio > .1 ? `watch` : `ok`,
    title: `IS 對資金 RPO 的殘差`,
    detail: `FY27 總營收 ${r.revenue.toFixed(1)} − RPO 轉換 ${r.fundingRev.toFixed(1)} = 對帳差額 ${r.inYear.toFixed(1)}（${(r.plugRatio*100).toFixed(0)}%）。硬體／服務 ${r.nonRpo.toFixed(1)} 已在 IS 單列。差額／營收超過 10% 表示分部成長假設大於期初合約轉換，不是獨立 bookings。`
  }, {
    id: `val-double`,
    ok: !0,
    severity: `ok`,
    title: `評價法防雙重扣除`,
    detail: `DCF 已含 Cash CapEx，不另扣缺口。SOTP／PE 未含建置現金，另扣未籌缺口。加權前各算各的，加權後不再扣一次。`
  }]
}

function DM(e) {
  return Number.isFinite(e) ? String(parseFloat(e.toPrecision(10))) : `0`
}

function OM({
  value: e,
  onChange: t,
  step: n = .1
}) {
  return (0, $.jsx)(`input`, {
    type: `number`,
    step: n,
    value: DM(e),
    onChange: e => t(Number(e.target.value)),
    className: `h-9 w-full min-w-14 rounded-sm border border-border bg-card px-2 text-right font-mono text-xs tabular-nums text-fg focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/30`
  })
}

function kM({
  label: e,
  hint: t,
  children: n
}) {
  return (0, $.jsxs)(`label`, {
    className: `block space-y-1.5`,
    children: [(0, $.jsxs)(`span`, {
      className: `flex items-baseline justify-between gap-2 text-sm font-medium`,
      children: [e, t ? (0, $.jsx)(`span`, {
        className: `font-mono text-xs text-muted tabular-nums`,
        children: t
      }) : null]
    }), n]
  })
}

function AM({
  result: e,
  state: t,
  v: n,
  pack: r,
  onPatch: i,
  tab: a,
  onTab: o
}) {
  let {
    fwd: s,
    d: c,
    sRaw: l,
    sAdj: u,
    peRaw: d,
    peAdj: f,
    hole: p,
    call: m,
    shares: h
  } = r, g = (0, v.useMemo)(() => SM(s, r.v), [s, r.v]), _ = (0, v.useMemo)(() => CM(s, r.v, p), [s, r.v, p]), y = n.price * h, b = y + n.netDebt, x = s[0].revenue, S = s[0].ngEps, C = S > 0 ? n.price / S : 0, w = x > 0 ? b / x : 0, T = wM(lM.map(e => e.fwdPe)), E = [`損益簡表`, `Comps`, `DCF`, `目標價`], D = m.call === `買進` ? `ok` : m.call === `賣出` ? `bad` : `watch`, O = e.totals.end;
  return (0, $.jsxs)(`div`, {
    className: `space-y-4`,
    children: [(0, $.jsxs)(Jj, {
      className: Vj(`border px-4 py-3`, m.funded ? `border-ok/30 bg-ok/5` : `border-bad/30 bg-bad/5`),
      children: [(0, $.jsxs)(`div`, {
        className: `flex flex-wrap items-start justify-between gap-3`,
        children: [(0, $.jsxs)(`div`, {
          className: `min-w-0`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: `資金模型 → 評價 連動`
          }), (0, $.jsxs)(`p`, {
            className: `mt-1 text-sm leading-relaxed`,
            children: [`GPU 產能、Cash CapEx、利息、期末現金皆來自資金模型。SOTP／本益比另扣未籌資金`, p > .5 ? ` ${Y(p,0)}bn` : `（本路徑無缺口）`, `。DCF 不另扣。資金狀態與投資評等分開：Unfunded 不准買進，但評等仍由融資後目標價相對現價決定。`]
          })]
        }), (0, $.jsxs)(`div`, {
          className: `flex shrink-0 flex-row flex-wrap items-center gap-2`,
          children: [(0, $.jsx)(Uj, {
            tone: m.funded ? `ok` : `bad`,
            children: m.funded ? `Funded` : `Unfunded`
          }), (0, $.jsx)(Uj, {
            tone: D,
            children: m.call
          })]
        })]
      }), (0, $.jsx)(`div`, {
        className: `mt-3 grid grid-cols-2 gap-2 md:grid-cols-5`,
        children: [
          [`FY27 GPU 產能`, mA(s[0].gpu) + `bn`],
          [`五年 Cash CapEx`, mA(e.totals.cashCapex) + `bn`],
          [`期末現金`, mA(O) + `bn`],
          [`ATM／股數`, `${t.includeAtm?`含`:`不含`} · ${Y(h,2)}bn`],
          [`評價結論`, m.call]
        ].map(([e, t]) => (0, $.jsxs)(`div`, {
          className: `rounded-md bg-card/80 px-3 py-2`,
          children: [(0, $.jsx)(`div`, {
            className: `text-[11px] text-muted`,
            children: e
          }), (0, $.jsx)(`div`, {
            className: Vj(`mt-0.5 font-mono text-sm tabular-nums`, e === `期末現金` && (O < 0 ? `text-bad` : `text-ok`), e === `評價結論` && (D === `ok` ? `text-ok` : D === `bad` ? `text-bad` : `text-watch`)),
            children: t
          })]
        }, e))
      })]
    }), (0, $.jsx)(`div`, {
      className: `grid grid-cols-2 gap-3 md:grid-cols-3 lg:grid-cols-6`,
      children: [
        [`現價（9/10 收）`, `$${Y(n.price,2)}`],
        [`市值`, `${Y(y,0)}bn`],
        [`企業價值`, `${Y(b,0)}bn`],
        [`本模型目標價`, `$${Y(m.blended,0)}`],
        [`潛在空間`, hA(m.upside * 100, 0)],
        [`投資結論`, m.call]
      ].map(([e, t], n) => (0, $.jsxs)(Jj, {
        className: `p-3 md:p-4`,
        children: [(0, $.jsx)(`div`, {
          className: `text-xs text-muted`,
          children: e
        }), (0, $.jsx)(`div`, {
          className: Vj(`mt-1 font-display text-xl font-semibold tabular-nums`, n === 4 && (m.upside >= 0 ? `text-ok` : `text-bad`), n === 5 && (D === `ok` ? `text-ok` : D === `bad` ? `text-bad` : `text-watch`)),
          children: t
        })]
      }, e))
    }), (0, $.jsxs)(`div`, {
      className: `grid gap-4 lg:grid-cols-[minmax(0,20rem)_minmax(0,1fr)]`,
      children: [(0, $.jsxs)(Jj, {
        className: `min-w-0`,
        children: [(0, $.jsx)(`h2`, {
          className: `font-display text-lg font-semibold`,
          children: `評價假設`
        }), (0, $.jsx)(`p`, {
          className: `mt-1 text-xs leading-relaxed text-muted`,
          children: `現價為 2026-09-10 收盤。股數隨資金模型 ATM 開關連動（關閉則扣回 Q1 已發行 141m 股）。OCI 量＝資金模型 Billable 產能＋非 GPU OCI。稅率已改 15.1%（Q1 ETR）。淨負債用 FY26 末 97.6，不用 Q1 現貨 88.3，以免 ATM 現金在評價與資金兩邊雙計。`
        }), (0, $.jsxs)(`div`, {
          className: `mt-4 space-y-3`,
          children: [(0, $.jsx)(kM, {
            label: `現價 US$`,
            hint: Y(n.price, 2),
            children: (0, $.jsx)(OM, {
              value: n.price,
              onChange: e => i({
                price: e
              }),
              step: .01
            })
          }), (0, $.jsx)(kM, {
            label: `ATM 後股數（十億）`,
            hint: `10-Q 3.024 · 連動後 ${Y(h,2)}`,
            children: (0, $.jsx)(OM, {
              value: n.shares,
              onChange: e => i({
                shares: e
              }),
              step: .01
            })
          }), (0, $.jsx)(kM, {
            label: `淨負債 US$bn`,
            hint: `FY26 末 97.6（Q1 現貨 88.3）`,
            children: (0, $.jsx)(OM, {
              value: n.netDebt,
              onChange: e => i({
                netDebt: e
              })
            })
          }), (0, $.jsx)(kM, {
            label: `WACC`,
            hint: hA(n.wacc * 100, 1),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 70,
              max: 130,
              value: Math.round(n.wacc * 1e3),
              onChange: e => i({
                wacc: Number(e.target.value) / 1e3
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `永續成長 g`,
            hint: hA(n.g * 100, 1),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 15,
              max: 45,
              value: Math.round(n.g * 1e3),
              onChange: e => i({
                g: Number(e.target.value) / 1e3
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `Apps SaaS 年增`,
            hint: hA(n.appsG * 100, 0),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 0,
              max: 20,
              value: Math.round(n.appsG * 100),
              onChange: e => i({
                appsG: Number(e.target.value) / 100
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `SOTP：OCI EV/Sales`,
            hint: `${Y(n.ociEvSales,1)}x`,
            children: (0, $.jsx)(OM, {
              value: n.ociEvSales,
              onChange: e => i({
                ociEvSales: e
              }),
              step: .5
            })
          }), (0, $.jsx)(kM, {
            label: `SOTP：傳統 EV/Sales`,
            hint: `${Y(n.tradEvSales,1)}x`,
            children: (0, $.jsx)(OM, {
              value: n.tradEvSales,
              onChange: e => i({
                tradEvSales: e
              }),
              step: .5
            })
          }), (0, $.jsx)(kM, {
            label: `NTM 非 GAAP PE`,
            hint: `${Y(n.ntmPe,0)}x`,
            children: (0, $.jsx)(OM, {
              value: n.ntmPe,
              onChange: e => i({
                ntmPe: e
              }),
              step: 1
            })
          }), (0, $.jsx)(`p`, {
            className: `text-xs leading-relaxed text-muted`,
            children: `FY27 營收錨在公司指引 ≥$90bn：OCI ≈ 產能＋other OCI，Apps +10%，授權／支援略降。改資金模型的 MW／CapEx／ATM，此頁 OCI、UFCF、目標價與結論會跟著動。`
          })]
        })]
      }), (0, $.jsxs)(`div`, {
        className: `min-w-0 space-y-4`,
        children: [(0, $.jsx)(`div`, {
          className: `w-full max-w-full overflow-x-auto`,
          children: (0, $.jsx)(`div`, {
            className: `flex w-max gap-1 rounded-lg bg-surface p-1`,
            children: E.map((e, t) => (0, $.jsx)(`button`, {
              onClick: () => o(t),
              className: Vj(`h-10 shrink-0 rounded-md px-3 text-sm font-medium`, a === t ? `bg-card text-fg shadow-card` : `text-muted hover:text-fg`),
              children: e
            }, e))
          })
        }), a === 0 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(`h2`, {
            className: `font-display text-lg font-semibold`,
            children: `過去五年＋未來五年（損益為主）`
          }), (0, $.jsx)(`p`, {
            className: `mt-1 text-xs leading-relaxed text-muted`,
            children: `傳統業務＝Apps SaaS＋授權／支援＋硬體／服務。前瞻 OCI 連動資金模型可計費產能。RPO 轉換是期初合約產能約束後的數字；其餘拆成年內新單／其他，硬體服務單列，不再混成一筆殘差。`
          }), (0, $.jsxs)(`table`, {
            className: `mt-3 w-full min-w-[860px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `US$bn`
                }), sM.map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: e.year
                }, e.year)), s.map(e => (0, $.jsxs)(`th`, {
                  className: `py-2 text-right font-medium text-accent`,
                  children: [e.year, `E`]
                }, e.year))]
              })
            }), (0, $.jsxs)(`tbody`, {
              children: [
                [
                  [`GPU／AI 產能（資金）`, [...sM.map(() => NaN), ...s.map(e => e.gpu)]],
                  [`OCI（IaaS）`, [...sM.map(e => e.oci), ...s.map(e => e.oci)]],
                  [`Apps SaaS`, [...sM.map(e => e.apps), ...s.map(e => e.apps)]],
                  [`授權＋支援`, [...sM.map(e => e.software), ...s.map(e => e.software)]],
                  [`硬體＋服務`, [...sM.map(e => e.other), ...s.map(e => e.other)]],
                  [`總營收（IS）`, [...sM.map(e => e.revenue), ...s.map(e => e.revenue)]],
                  [`RPO 轉換（資金）`, [...sM.map(() => NaN), ...s.map(e => e.fundingRev)]],
                  [`對帳差額（年內新單／分類）`, [...sM.map(() => NaN), ...s.map(e => e.inYear)]]
                ].map(([e, t]) => (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: e
                  }), t.map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: Number.isFinite(e) ? Y(e, 1) : `—`
                  }, t))]
                }, e)), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `營收 YoY`
                  }), [...sM, ...s].map((e, t, n) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: t === 0 ? `—` : hA((`yoy` in e ? e.yoy : uM(e.revenue, n[t - 1].revenue)) * 100, 0)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `GAAP 營業利益`
                  }), [...sM.map(e => e.opInc), ...s.map(e => e.opInc)].map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: Y(e, 1)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `營利率`
                  }), [...sM.map(e => e.opInc / e.revenue), ...s.map(e => e.opM)].map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: hA(e * 100, 0)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `淨利率`
                  }), [...sM.map(e => e.ni / e.revenue), ...s.map(e => e.nm)].map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: hA(e * 100, 0)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `GAAP EPS`
                  }), [...sM.map(e => e.eps), ...s.map(e => e.eps)].map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: Y(e, 2)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: `非 GAAP EPS`
                  }), [...sM.map(e => e.ngEps), ...s.map(e => e.ngEps)].map((e, t) => (0, $.jsx)(`td`, {
                    className: `py-1.5 text-right font-mono tabular-nums`,
                    children: Y(e, 2)
                  }, t))]
                })
              ]
            })]
          }), (0, $.jsxs)(`p`, {
            className: `mt-3 text-xs text-muted`,
            children: [`公司 FY27 指引總營收至少 $90bn、非 GAAP EPS $8.10。本表 FY27 總營收 `, Y(s[0].revenue, 1), `bn、非 GAAP EPS $`, Y(s[0].ngEps, 2), `。RPO 轉換 `, Y(s[0].fundingRev, 1), `，對帳差額 `, Y(s[0].inYear, 1), `（占營收`, ` `, hA(s[0].plugRatio * 100, 0), s[0].plugRatio > .1 ? `，超過 10%` : ``, `）。差額含年內新單、非 RPO 與兩套口徑差，不是獨立驅動項。`]
          }), (0, $.jsx)(`div`, {
            className: `mt-3 overflow-x-auto`,
            children: (0, $.jsxs)(`table`, {
              className: `w-full min-w-[640px] text-xs`,
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsxs)(`tr`, {
                  className: `text-muted`,
                  children: [(0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `前瞻營利率（可改）`
                  }), oM.map(e => (0, $.jsx)(`th`, {
                    className: `py-2 text-right font-medium`,
                    children: e
                  }, e))]
                })
              }), (0, $.jsxs)(`tbody`, {
                children: [(0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    className: `font-medium`,
                    children: `GAAP 營利率`
                  }), n.opMargin.map((e, t) => (0, $.jsx)(`td`, {
                    className: `p-1`,
                    children: (0, $.jsx)(OM, {
                      value: e * 100,
                      onChange: e => {
                        let r = [...n.opMargin];
                        r[t] = e / 100, i({
                          opMargin: r
                        })
                      },
                      step: .5
                    })
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    className: `font-medium`,
                    children: `非 GPU OCI US$bn`
                  }), n.otherOci.map((e, t) => (0, $.jsx)(`td`, {
                    className: `p-1`,
                    children: (0, $.jsx)(OM, {
                      value: e,
                      onChange: e => {
                        let r = [...n.otherOci];
                        r[t] = e, i({
                          otherOci: r
                        })
                      }
                    })
                  }, t))]
                })]
              })]
            })
          })]
        }), a === 1 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(`h2`, {
            className: `font-display text-lg font-semibold`,
            children: `可比公司（約 2026-09-10）`
          }), (0, $.jsxs)(`p`, {
            className: `mt-1 text-xs leading-relaxed text-muted`,
            children: [`OCI 對標雲基礎設施。傳統 DB＋SaaS 對標 CRM／SAP。下表敏感度已扣除資金缺口每股`, ` `, Y(p / Math.max(h, .01), 0), `。`]
          }), (0, $.jsxs)(`table`, {
            className: `mt-3 w-full min-w-[640px] text-sm`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsx)(`tr`, {
                className: `text-xs text-muted`,
                children: [`公司`, `角色`, `TTM PE`, `Forward PE`, `EV/Sales`, `市值 US$bn`].map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium first:text-left`,
                  children: e
                }, e))
              })
            }), (0, $.jsxs)(`tbody`, {
              children: [lM.map(e => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsxs)(`td`, {
                  className: `py-2`,
                  children: [(0, $.jsx)(`span`, {
                    className: `font-medium`,
                    children: e.ticker
                  }), (0, $.jsx)(`span`, {
                    className: `ml-2 text-xs text-muted`,
                    children: e.name
                  })]
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right text-xs text-muted`,
                  children: e.role
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(e.pe, 1), `x`]
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(e.fwdPe, 1), `x`]
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(e.evSales, 1), `x`]
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: Y(e.mkt, 0)
                })]
              }, e.ticker)), (0, $.jsxs)(`tr`, {
                className: `border-t-2 border-ink`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-2 font-medium`,
                  children: `ORCL NTM`
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right text-xs text-muted`,
                  children: `本模型`
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(C, 1), `x`]
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(C, 1), `x`]
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [Y(w, 1), `x`]
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: Y(y, 0)
                })]
              })]
            })]
          }), (0, $.jsxs)(`p`, {
            className: `mt-3 text-xs leading-relaxed text-muted`,
            children: [`Comp forward PE 中位數 `, Y(T, 1), `x。ORCL NTM 非 GAAP PE `, Y(C, 1), `x、EV/Sales`, ` `, Y(w, 1), `x。SOTP 未扣缺口 $`, Y(l, 0), `，扣資金缺口後 $`, Y(u, 0), `。`]
          }), (0, $.jsx)(`h3`, {
            className: `mt-5 text-sm font-medium`,
            children: `SOTP 敏感度（已扣資金缺口，每股 US$）`
          }), (0, $.jsxs)(`table`, {
            className: `mt-2 w-full min-w-[520px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `傳統＼OCI`
                }), [6, 7.5, 8.5, 10, 12].map(e => (0, $.jsxs)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: [e, `x`]
                }, e))]
              })
            }), (0, $.jsx)(`tbody`, {
              children: [4, 5, 5.5, 6.5, 8].map((e, t) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsxs)(`td`, {
                  className: `py-1.5 font-medium`,
                  children: [e, `x`]
                }), _[t].map((t, r) => (0, $.jsx)(`td`, {
                  className: Vj(`py-1.5 text-right font-mono tabular-nums`, Math.abs(e - n.tradEvSales) < .01 && Math.abs([6, 7.5, 8.5, 10, 12][r] - n.ociEvSales) < .01 ? `bg-accent-soft font-medium` : ``),
                  children: Y(t, 0)
                }, r))]
              }, e))
            })]
          })]
        }), a === 2 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(`h2`, {
            className: `font-display text-lg font-semibold`,
            children: `DCF（與資金模型 Cash CapEx 連動）`
          }), (0, $.jsx)(`p`, {
            className: `mt-1 text-xs leading-relaxed text-muted`,
            children: `UFCF = EBIT×(1−t)＋D&A−Cash CapEx−營運資金。Cash CapEx 與利息取自資金模型。DCF 已內含建置支出，不再重複扣期末現金缺口。`
          }), (0, $.jsxs)(`table`, {
            className: `mt-3 w-full min-w-[640px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `US$bn`
                }), s.map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: e.year
                }, e.year))]
              })
            }), (0, $.jsx)(`tbody`, {
              children: [
                [`EBIT`, s.map(e => e.opInc)],
                [`利息（資金）`, s.map(e => e.interest)],
                [`D&A`, s.map(e => e.da)],
                [`Cash CapEx`, s.map(e => e.cashCapex)],
                [`UFCF`, s.map(e => e.ufcf)],
                [`PV`, c.pv]
              ].map(([e, t]) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-1.5 font-medium`,
                  children: e
                }), t.map((t, n) => (0, $.jsx)(`td`, {
                  className: Vj(`py-1.5 text-right font-mono tabular-nums`, e === `UFCF` && t < 0 && `text-bad`),
                  children: Y(t, 1)
                }, n))]
              }, e))
            })]
          }), (0, $.jsxs)(`div`, {
            className: `mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-4`,
            children: [(0, $.jsx)(jM, {
              k: `FCF 現值`,
              v: mA(c.pvFcf)
            }), (0, $.jsx)(jM, {
              k: `終值現值`,
              v: mA(c.pvTv)
            }), (0, $.jsx)(jM, {
              k: `終值占比`,
              v: hA(c.tvShare * 100, 0)
            }), (0, $.jsx)(jM, {
              k: `DCF 每股`,
              v: `$${Y(c.perShare,0)}`
            })]
          }), (0, $.jsx)(`h3`, {
            className: `mt-5 text-sm font-medium`,
            children: `WACC × g 敏感度（每股 US$）`
          }), (0, $.jsxs)(`table`, {
            className: `mt-2 w-full min-w-[520px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `g＼WACC`
                }), [`8.0%`, `9.0%`, `9.5%`, `10.0%`, `11.0%`].map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: e
                }, e))]
              })
            }), (0, $.jsx)(`tbody`, {
              children: [`2.0%`, `2.5%`, `3.0%`, `3.5%`, `4.0%`].map((e, t) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-1.5 font-medium`,
                  children: e
                }), g[t].map((e, n) => (0, $.jsx)(`td`, {
                  className: Vj(`py-1.5 text-right font-mono tabular-nums`, t === 2 && n === 2 && `bg-accent-soft font-medium`),
                  children: Number.isFinite(e) ? Y(e, 0) : `—`
                }, n))]
              }, e))
            })]
          })]
        }), a === 3 && (0, $.jsxs)(Jj, {
          children: [(0, $.jsxs)(`div`, {
            className: `flex flex-wrap items-center gap-2`,
            children: [(0, $.jsx)(`h2`, {
              className: `font-display text-lg font-semibold`,
              children: `目標價與操作區間`
            }), (0, $.jsx)(Uj, {
              tone: D,
              children: m.call
            }), (0, $.jsx)(Uj, {
              tone: m.funded ? `ok` : `bad`,
              children: m.funded ? `資金可自籌` : `資金缺口未解`
            })]
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 text-sm leading-relaxed`,
            children: [`加權目標價 `, (0, $.jsxs)(`span`, {
              className: `font-semibold tabular-nums`,
              children: [`$`, Y(m.blended, 0)]
            }), `＝ DCF $`, Y(c.perShare, 0), ` × `, hA(yM.dcf * 100, 0), ` ＋ SOTP（扣缺口）$`, Y(u, 0), ` ×`, ` `, hA(yM.sotp * 100, 0), ` ＋ PE（扣缺口）$`, Y(f, 0), ` × `, hA(yM.pe * 100, 0), `。相對現價 $`, Y(n.price, 2), ` 為`, ` `, (0, $.jsx)(`span`, {
              className: m.upside >= 0 ? `text-ok` : `text-bad`,
              children: hA(m.upside * 100, 0)
            }), `。下檔 $`, Y(m.lo, 0), `、上檔 $`, Y(m.hi, 0), `。資金狀態 `, m.funded ? `Funded` : `Unfunded`, `。`]
          }), (0, $.jsx)(`div`, {
            className: `mt-4 overflow-x-auto`,
            children: (0, $.jsxs)(`table`, {
              className: `w-full min-w-[640px] text-xs`,
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsxs)(`tr`, {
                  className: `text-muted`,
                  children: [(0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `方法`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `CapEx`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `未籌缺口`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `股數`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `淨負債`
                  })]
                })
              }), (0, $.jsx)(`tbody`, {
                children: bM.map(e => (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: e.method
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.capex
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.hole
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.shares
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.netDebt
                  })]
                }, e.method))
              })]
            })
          }), (0, $.jsx)(`table`, {
            className: `mt-4 w-full text-sm`,
            children: (0, $.jsx)(`tbody`, {
              children: [
                [`DCF（已含 CapEx）`, c.perShare, `不重複扣缺口`],
                [`SOTP 未扣缺口`, l, `OCI ${Y(n.ociEvSales,1)}x + 傳統 ${Y(n.tradEvSales,1)}x`],
                [`SOTP 扣資金缺口`, u, `缺口 ${Y(p,0)}bn ／ ${Y(h,2)}bn 股`],
                [`NTM PE 未扣缺口`, d, `${Y(n.ntmPe,0)}x × FY27E $${Y(s[0].ngEps,2)}`],
                [`NTM PE 扣資金缺口`, f, `與 SOTP 同一缺口`],
                [`Street 平均（參考）`, 241, `約 44 家、並非本模型；未反映本資金缺口`]
              ].map(([e, t, n]) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-2 font-medium`,
                  children: e
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [`$`, Y(Number(t), 0)]
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right text-xs text-muted`,
                  children: n
                })]
              }, String(e)))
            })
          }), (0, $.jsxs)(`div`, {
            className: `mt-4 grid gap-3 md:grid-cols-3`,
            children: [(0, $.jsx)(MM, {
              title: `賣出／下檔`,
              body: `低於 $${Y(m.lo,0)}`,
              hint: `資金缺口擴大、WACC 上移或 OCI 倍數壓縮。`
            }), (0, $.jsx)(MM, {
              title: `中立區間`,
              body: `$${Y(m.lo,0)}–$${Y(m.hi,0)}`,
              hint: `現價已部分反映成長，但 FCF 與 CDS 限制加碼。`
            }), (0, $.jsx)(MM, {
              title: `買進／上檔`,
              body: m.funded ? `高於 $${Y(m.hi,0)} 的確認` : `本路徑不給買進`,
              hint: m.funded ? `產能如期、Cash CapEx 下降、期末現金翻正。` : `先把資金模型的期末現金做正（降低 CapEx、提高客戶出資或 ATM），結論才會開放買進。`
            })]
          }), m.notes.length ? (0, $.jsx)(`ul`, {
            className: `mt-4 space-y-1 text-xs leading-relaxed text-muted`,
            children: m.notes.map(e => (0, $.jsxs)(`li`, {
              children: [`— `, e]
            }, e))
          }) : null, (0, $.jsx)(`p`, {
            className: `mt-4 text-xs leading-relaxed text-muted`,
            children: `結論規則：資金模型期末現金為負 → 禁止買進；缺口大於約 40bn 或 DCF 明顯低於現價 → 賣出。現金翻正後，上檔 ≥25% 才買進。這是研究框架，不是投資建議。Street $241 隱含更樂觀的 FCF 與倍數，與本資金模型的現金缺口並不一致。`
          })]
        })]
      })]
    })]
  })
}

function jM({
  k: e,
  v: t
}) {
  return (0, $.jsxs)(`div`, {
    className: `rounded-md border border-border bg-surface px-3 py-2`,
    children: [(0, $.jsx)(`div`, {
      className: `text-xs text-muted`,
      children: e
    }), (0, $.jsx)(`div`, {
      className: `mt-1 font-mono text-sm tabular-nums`,
      children: t
    })]
  })
}

function MM({
  title: e,
  body: t,
  hint: n
}) {
  return (0, $.jsxs)(`div`, {
    className: `rounded-lg border border-border bg-surface p-3`,
    children: [(0, $.jsx)(`div`, {
      className: `text-xs text-muted`,
      children: e
    }), (0, $.jsx)(`div`, {
      className: `mt-1 font-display text-lg font-semibold tabular-nums`,
      children: t
    }), (0, $.jsx)(`p`, {
      className: `mt-1 text-xs leading-relaxed text-muted`,
      children: n
    })]
  })
}
var NM = [`年度假設`, `年度結果`, `收入／容量`, `成本`, `信用`, `終值`, `園區`, `敏感性`, `連動檢查`, `來源`];

function PM() {
  return structuredClone(Qk)
}

function FM(e) {
  return Number.isFinite(e) ? String(parseFloat(e.toPrecision(10))) : `0`
}

function IM({
  value: e,
  onChange: t,
  step: n = .1
}) {
  return (0, $.jsx)(`input`, {
    type: `number`,
    step: n,
    value: FM(e),
    onChange: e => t(Number(e.target.value)),
    className: `h-9 w-full min-w-16 rounded-sm border border-border bg-card px-2 text-right font-mono text-xs tabular-nums text-fg focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/30`
  })
}

function LM({
  label: e,
  hint: t,
  children: n
}) {
  return (0, $.jsxs)(`label`, {
    className: `block space-y-1.5`,
    children: [(0, $.jsxs)(`span`, {
      className: `flex items-baseline justify-between gap-2 text-sm font-medium`,
      children: [e, t ? (0, $.jsx)(`span`, {
        className: `font-mono text-xs text-muted tabular-nums`,
        children: t
      }) : null]
    }), n]
  })
}

function RM({
  checked: e,
  onChange: t,
  label: n,
  hint: r
}) {
  return (0, $.jsxs)(`label`, {
    className: `flex cursor-pointer items-start gap-2.5 py-1`,
    children: [(0, $.jsx)(`input`, {
      type: `checkbox`,
      checked: e,
      onChange: e => t(e.target.checked),
      className: `mt-1 size-4 accent-accent`
    }), (0, $.jsxs)(`span`, {
      children: [(0, $.jsx)(`span`, {
        className: `block text-sm font-medium`,
        children: n
      }), (0, $.jsx)(`span`, {
        className: `block text-xs leading-relaxed text-muted`,
        children: r
      })]
    })]
  })
}

function zM() {
  let [e, t] = (0, v.useState)(PM), [n, r] = (0, v.useState)(0), [i, a] = (0, v.useState)(`funding`), [o, s] = (0, v.useState)(() => structuredClone(cM)), [c, l] = (0, v.useState)(0), u = (0, v.useRef)(null), d = (0, v.useMemo)(() => dA(e), [e]), f = (0, v.useMemo)(() => TM(d, e, o), [d, e, o]), p = (0, v.useMemo)(() => EM(f, d, e), [f, d, e]), m = (0, v.useMemo)(() => pA(e), [e]), h = lA(e.cds), g = d.totals.end >= 0, _ = d.totals.operatingGap + e.cash >= 0, y = (0, v.useMemo)(() => aA(e, d), [e, d]), b = (0, v.useMemo)(() => oA(d), [d]), x = (0, v.useMemo)(() => sA(e, d), [e, d]), S = (0, v.useMemo)(() => cA(e), [e]), C = e => s(t => ({
    ...t,
    ...e
  })), w = e => t(t => ({
    ...t,
    ...e,
    scenario: e.scenario ?? `custom`
  })), T = e => t(t => ({
    ...t,
    scenario: e,
    a: structuredClone(Zk[e].a)
  })), E = (e, n, r) => t(t => {
    let i = structuredClone(t.a);
    return e === `margin` ? i.margin = r : i[e][n] = r, {
      ...t,
      a: i,
      scenario: `custom`
    }
  }), D = (e, n, r) => t(t => {
    let i = structuredClone(t.m);
    return i[e][n] = r, {
      ...t,
      m: i,
      scenario: `custom`
    }
  }), O = (e, n, r) => t(t => {
    let i = structuredClone(t.sites),
      a = i[e];
    return !a || a.residual ? t : (a[n] = r, {
      ...t,
      sites: i,
      scenario: `custom`
    })
  }), k = (e, t) => {
    let n = new Blob([`﻿` + t.map(e => e.join(`,`)).join(`
`)], {
        type: `text/csv;charset=utf-8`
      }),
      r = document.createElement(`a`);
    r.href = URL.createObjectURL(n), r.download = e, r.click(), URL.revokeObjectURL(r.href)
  };
  return Math.max(...d.years.flatMap(e => [e.sources, e.uses]), 1), (0, $.jsxs)(`div`, {
    className: `min-h-screen overflow-x-hidden bg-bg text-fg`,
    children: [(0, $.jsx)(`header`, {
      className: `border-b border-border bg-ink text-accent-fg`,
      children: (0, $.jsxs)(`div`, {
        className: `mx-auto flex max-w-7xl flex-col gap-5 px-4 py-8 md:px-6`,
        children: [(0, $.jsxs)(`div`, {
          className: `flex flex-wrap items-center gap-2 text-xs tracking-wide text-accent-soft`,
          children: [(0, $.jsx)(`span`, {
            children: `ORCL FY27–FY31`
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsx)(`span`, {
            children: `產能約束 · 資金優先`
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsxs)(`span`, {
            children: [`v`, `4.3`, ` · `, Bk]
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsx)(`span`, {
            children: `Q1 FY27 10-Q + 法說 · 口徑分層`
          })]
        }), (0, $.jsx)(`h1`, {
          className: `max-w-3xl font-display text-3xl font-semibold leading-tight tracking-tight md:text-4xl`,
          children: `RPO 不是現金`
        }), (0, $.jsx)(`p`, {
          className: `max-w-3xl font-display text-lg font-medium leading-snug text-accent-fg md:text-xl`,
          children: `先問建置誰出錢，再談本益比`
        }), (0, $.jsxs)(`p`, {
          className: `max-w-3xl text-sm leading-relaxed text-pretty text-accent-soft`,
          children: [`RPO `, (0, $.jsx)(`span`, {
            className: `font-medium text-accent-fg`,
            children: `638→664`
          }), `、五年認列`, ` `, (0, $.jsx)(`span`, {
            className: `font-medium text-accent-fg`,
            children: `80%→84%`
          }), `（10-Q；法說「一半在 36 個月內」只是把 13+37 講成人話）。期末現金`, ` `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [mA(d.totals.end), `bn`]
          }), `，目標價`, ` `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(f.call.blended, 0)]
          }), `（`, f.call.upside >= 0 ? `+` : ``, hA(f.call.upside * 100, 0), `），結論「`, f.call.call, `」`, f.call.call === Hk.call ? `未變` : ``, `——產能故事法說講得滿，建置帳單在 10-Q。Q1 的 850 MW、GPU 利用率 97.9% 寫在觀察項，不年化。v4.3 把 ATM 從營運來源拆開；Unfunded 是資金狀態，不是拿來代替估值的開關，但現金翻正前仍不准買進。$288bn 表外租賃法說沒提，聽 10-Q。`]
        }), (0, $.jsxs)(`div`, {
          className: `flex flex-wrap gap-2`,
          children: [(0, $.jsxs)(qj, {
            variant: `solid`,
            onClick: () => k(`oracle_annual_v43.csv`, [Object.keys(d.years[0]), ...d.years.map(e => Object.values(e).map(e => typeof e == `number` ? e : String(e)))]),
            children: [(0, $.jsx)(rM, {
              className: `size-4`
            }), `年度 CSV`]
          }), (0, $.jsx)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => k(`oracle_sites_v43.csv`, [Object.keys(d.sites[0]), ...d.sites.map(e => Object.values(e).map(e => String(e ?? ``)))]),
            children: `園區 CSV`
          }), (0, $.jsx)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => {
              let t = new Blob([JSON.stringify(e, null, 2)], {
                  type: `application/json`
                }),
                n = document.createElement(`a`);
              n.href = URL.createObjectURL(t), n.download = `oracle_model_v43.json`, n.click()
            },
            children: `儲存 JSON`
          }), (0, $.jsxs)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => u.current?.click(),
            children: [(0, $.jsx)(aM, {
              className: `size-4`
            }), `載入`]
          }), (0, $.jsx)(`input`, {
            ref: u,
            hidden: !0,
            type: `file`,
            accept: `application/json`,
            onChange: e => {
              let n = e.target.files?.[0];
              if (!n) return;
              let r = new FileReader;
              r.onload = () => {
                try {
                  t({
                    ...PM(),
                    ...JSON.parse(String(r.result))
                  })
                } catch {}
              }, r.readAsText(n)
            }
          }), (0, $.jsxs)(qj, {
            variant: `ghost`,
            className: `text-accent-fg hover:bg-white/10`,
            onClick: () => {
              t(PM()), s(structuredClone(cM))
            },
            children: [(0, $.jsx)(iM, {
              className: `size-4`
            }), `重設`]
          })]
        })]
      })
    }), (0, $.jsxs)(`div`, {
      className: `mx-auto max-w-7xl space-y-4 px-4 py-5 md:px-6 md:py-6`,
      children: [(0, $.jsx)(`div`, {
        className: `flex w-full max-w-full overflow-x-auto`,
        children: (0, $.jsx)(`div`, {
          className: `flex gap-1 rounded-lg bg-surface p-1`,
          children: [
            [`funding`, `資金模型`],
            [`value`, `損益與評價`]
          ].map(([e, t]) => (0, $.jsx)(`button`, {
            onClick: () => a(e),
            className: Vj(`h-11 min-w-28 rounded-md px-4 text-sm font-medium`, i === e ? `bg-ink text-accent-fg` : `text-muted hover:text-fg`),
            children: t
          }, e))
        })
      }), (0, $.jsx)(`div`, {
        className: Vj(i !== `value` && `hidden`),
        children: (0, $.jsx)(AM, {
          result: d,
          state: e,
          v: o,
          pack: f,
          onPatch: C,
          tab: c,
          onTab: l
        })
      }), (0, $.jsxs)(`div`, {
        className: Vj(i !== `funding` && `hidden`),
        children: [(0, $.jsx)(`div`, {
          className: `grid grid-cols-2 gap-3 md:grid-cols-3 lg:grid-cols-6`,
          children: [
            [`排程 RPO（五年）`, mA(b.scheduled) + `bn`],
            [`可實現（扣產能／信用）`, mA(b.collected) + `bn`],
            [`營運缺口（不含 ATM）`, mA(d.totals.operatingGap) + `bn`],
            [`期初 + ATM`, mA(e.cash + d.totals.atm) + `bn`],
            [`期末現金`, mA(d.totals.end) + `bn`],
            [`連動評價／結論`, `$${Y(f.call.blended,0)} · ${f.call.call}`]
          ].map(([e, t], n) => (0, $.jsxs)(Jj, {
            className: `p-3 md:p-4`,
            children: [(0, $.jsx)(`div`, {
              className: `text-xs text-muted`,
              children: e
            }), (0, $.jsx)(`div`, {
              className: Vj(`mt-1 font-display text-xl font-semibold tabular-nums`, n === 2 && (d.totals.operatingGap >= 0 ? `text-ok` : `text-bad`), n === 4 && (g ? `text-ok` : `text-bad`), n === 5 && (f.call.call === `買進` ? `text-ok` : f.call.call === `賣出` ? `text-bad` : `text-watch`)),
              children: t
            })]
          }, e))
        }), (0, $.jsxs)(Jj, {
          className: `p-4`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: `資金四層橋（同一條公式產生所有缺口）`
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 font-mono text-sm tabular-nums leading-relaxed`,
            children: [`期初 `, Y(y.open), ` + 營運缺口 `, Y(y.operatingGap), ` + ATM `, Y(y.atm), y.debtPay ? ` − 還本 ${Y(y.debtPay)}` : ``, ` = 期末 `, Y(y.end), y.ok ? `` : `（對不上）`]
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 text-xs leading-relaxed text-muted`,
            children: [`產能橋：排程 `, Y(b.scheduled), ` − 瓶頸 `, Y(b.bottleneck), ` − 信用損失 `, Y(b.credit), ` = 收現 `, Y(b.collected), `。這就是「RPO 不是現金」。 RPO 滾動：FY26 末 `, Y(x.begin), ` + Q1 淨增加 `, Y(x.q1Net), ` = `, Y(x.q1End), `；五年排程 `, Y(x.convert), `，五年後仍未認列 `, Y(x.remaining), `。 表外現金租金五年 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), `，對 288 差 `, Y(S.gap), `。`]
          })]
        }), (0, $.jsxs)(Jj, {
          className: `border-accent/20 bg-accent/5 p-4`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: Vk
          }), (0, $.jsx)(`div`, {
            className: `mt-3 grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-6`,
            children: [
              [`營收 / OCI`, `${Y(Yk.revenue,1)}bn +${hA(Yk.yoy*100,0)} · OCI ${Y(Yk.oci,1)} +${hA(Yk.ociYoy*100,0)}`],
              [`RPO 桶`, `${Y(Yk.rpo,0)}bn · 13/37/34/16 · 36個月 50%`],
              [`現金 capex`, `Q1 ${Y(Yk.capexCash,1)} · 指引毛 90–95／淨≤70`],
              [`ATM / 股數`, `${Y(Yk.atmNet,1)}bn · ${Y(Yk.sharesOut,3)}bn 股`],
              [`客戶預付`, `${Y(Yk.prepaySfc,1)}bn · 遞延 ${Y(Yk.deferredTotal,1)}`],
              [`表外租賃`, `${Y(Yk.offBalanceLease,0)}bn · Q1 新 ROU ${Y(Yk.rouOpObtained+Yk.rouFinObtained,1)}`]
            ].map(([e, t]) => (0, $.jsxs)(`div`, {
              className: `rounded-md bg-card/80 px-3 py-2`,
              children: [(0, $.jsx)(`div`, {
                className: `text-xs text-muted`,
                children: e
              }), (0, $.jsx)(`div`, {
                className: `mt-0.5 font-mono text-xs tabular-nums leading-snug`,
                children: t
              })]
            }, e))
          }), (0, $.jsx)(`div`, {
            className: `mt-2 grid grid-cols-2 gap-2 sm:grid-cols-4`,
            children: [
              [`Q1 交付 MW`, `${Xk.mwDelivered} MW · 觀察，不年化`],
              [`GPU 利用率`, `Q1 ${Xk.gpuUtil}% · 基準全年 ${Y(e.m.util[0],1)}%`],
              [`Abilene`, `${Xk.abileneMw} MW · 六／八棟已入園區`],
              [`全年指引`, `營收 ≥$${Xk.fy27RevFloor}bn · 當下限，非擬合`]
            ].map(([e, t]) => (0, $.jsxs)(`div`, {
              className: `rounded-md bg-card/80 px-3 py-2`,
              children: [(0, $.jsx)(`div`, {
                className: `text-xs text-muted`,
                children: e
              }), (0, $.jsx)(`div`, {
                className: `mt-0.5 font-mono text-xs tabular-nums leading-snug`,
                children: t
              })]
            }, e))
          })]
        }), (0, $.jsxs)(`div`, {
          className: `grid gap-4 lg:grid-cols-[minmax(0,20rem)_minmax(0,1fr)]`,
          children: [(0, $.jsx)(`aside`, {
            className: `min-w-0 space-y-4`,
            children: (0, $.jsxs)(Jj, {
              children: [(0, $.jsx)(`h2`, {
                className: `font-display text-lg font-semibold`,
                children: `場景與全域假設`
              }), (0, $.jsx)(`div`, {
                className: `mt-3 grid grid-cols-3 gap-1 rounded-lg bg-surface p-1`,
                children: [`low`, `base`, `high`].map(t => (0, $.jsx)(`button`, {
                  onClick: () => T(t),
                  className: Vj(`h-10 rounded-md text-sm font-medium`, e.scenario === t ? `bg-ink text-accent-fg` : `text-muted hover:bg-card`),
                  children: Zk[t].label
                }, t))
              }), (0, $.jsxs)(`div`, {
                className: `mt-4 space-y-4`,
                children: [(0, $.jsx)(LM, {
                  label: `FY26 末 RPO（US$bn）`,
                  hint: `638 10-K`,
                  children: (0, $.jsx)(IM, {
                    value: e.rpoFy26,
                    onChange: e => w({
                      rpoFy26: e
                    })
                  })
                }), (0, $.jsx)(LM, {
                  label: `Q1 RPO 淨增加（US$bn）`,
                  hint: `664−638，不是 gross bookings`,
                  children: (0, $.jsx)(IM, {
                    value: e.rpoQ1Add,
                    onChange: e => w({
                      rpoQ1Add: e
                    })
                  })
                }), (0, $.jsx)(LM, {
                  label: `五年認列比例`,
                  hint: hA(e.rp, 0),
                  children: (0, $.jsx)(`input`, {
                    type: `range`,
                    min: 50,
                    max: 100,
                    value: e.rp,
                    onChange: e => w({
                      rp: Number(e.target.value)
                    }),
                    className: `w-full accent-accent`
                  })
                }), (0, $.jsxs)(LM, {
                  label: `RPO 現金貢獻率`,
                  hint: hA(e.a.margin * 100, 0),
                  children: [(0, $.jsx)(`input`, {
                    type: `range`,
                    min: 10,
                    max: 70,
                    value: Math.round(e.a.margin * 100),
                    onChange: e => E(`margin`, 0, Number(e.target.value) / 100),
                    className: `w-full accent-accent`
                  }), (0, $.jsx)(`p`, {
                    className: `text-xs leading-relaxed text-muted`,
                    children: `已收現 RPO 收入，扣除履約現金成本與現金稅後，可支應資本用途的比例。預設已含電費；打開 overlay 會再扣一次。`
                  })]
                }), (0, $.jsx)(LM, {
                  label: `期初現金 FY26 末（US$bn）`,
                  hint: `31.894`,
                  children: (0, $.jsx)(IM, {
                    value: e.cash,
                    onChange: e => w({
                      cash: e
                    })
                  })
                }), (0, $.jsx)(RM, {
                  checked: e.useAvgMw,
                  onChange: e => w({
                    useAvgMw: e
                  }),
                  label: `收入改用平均在役 MW`,
                  hint: `關掉則用年末存量全年化。打開用（FY26 末 ${e.billableOpen} + 當年期末）/2，年中交付不會被當成全年在役。`
                }), (0, $.jsx)(LM, {
                  label: `FY26 末 Billable MW`,
                  hint: `Assumed · FY26 交付約 1,164`,
                  children: (0, $.jsx)(IM, {
                    value: e.billableOpen,
                    onChange: e => w({
                      billableOpen: e
                    }),
                    step: 50
                  })
                }), (0, $.jsx)(RM, {
                  checked: e.includeAtm,
                  onChange: t => w({
                    includeAtm: t,
                    scenario: e.scenario
                  }),
                  label: `計入 Q1 已完成 ATM 股權 US$19.9bn`,
                  hint: `141m 股、淨額 $19.9bn。已發生事實，預設打開。不算進「營運缺口」。`
                }), (0, $.jsx)(RM, {
                  checked: e.includeDebt,
                  onChange: e => w({
                    includeDebt: e
                  }),
                  label: `到期債務全數還本（壓力）`,
                  hint: `只加本金流出，不含再融資。不是基本預測。`
                }), (0, $.jsx)(RM, {
                  checked: e.overlay,
                  onChange: e => w({
                    overlay: e
                  }),
                  label: `另扣電力＋維護 overlay`,
                  hint: `貢獻率若已淨額化，請保持關閉。`
                }), (0, $.jsx)(RM, {
                  checked: e.linkSites,
                  onChange: t => w({
                    linkSites: t,
                    scenario: e.scenario
                  }),
                  label: `園區 MW 連動年度容量`,
                  hint: `FY27 Accepted/Billable 不低於具名園區加總；不足部分進殘差列。`
                }), (0, $.jsx)(RM, {
                  checked: e.cdsLink,
                  onChange: t => w({
                    cdsLink: t,
                    scenario: e.scenario
                  }),
                  label: `CDS 溢價傳入透支利率`,
                  hint: `透支利率 = 基準全包利率 + max(0, CDS−150)×40%。基準利率若已含 CDS，不要再開，避免雙計。不進入客戶違約。`
                })]
              }), (0, $.jsxs)(`div`, {
                className: `mt-5 rounded-lg border border-border bg-surface p-3`,
                children: [(0, $.jsx)(`div`, {
                  className: `text-sm font-medium`,
                  children: `Oracle 自身 5Y Senior CDS`
                }), (0, $.jsxs)(`p`, {
                  className: `mt-1 text-xs leading-relaxed text-muted`,
                  children: [`截至 `, e.cdsDate, `；Bid/Ask `, Y(e.cdsBid, 3), ` / `, Y(e.cdsAsk, 3), `。不是客戶違約率。`]
                }), (0, $.jsx)(LM, {
                  label: `CDS（bps）`,
                  hint: `${Y(e.cds,3)}`,
                  children: (0, $.jsx)(`input`, {
                    type: `range`,
                    min: 50,
                    max: 600,
                    value: e.cds,
                    onChange: e => w({
                      cds: Number(e.target.value)
                    }),
                    className: `w-full accent-accent`
                  })
                }), (0, $.jsxs)(`div`, {
                  className: `mt-2 flex items-center justify-between`,
                  children: [(0, $.jsx)(Uj, {
                    tone: h.tone,
                    children: h.label
                  }), (0, $.jsxs)(`span`, {
                    className: `text-xs text-muted`,
                    children: [150, ` / `, 250, ` / `, 500]
                  })]
                })]
              })]
            })
          }), (0, $.jsxs)(`main`, {
            className: `min-w-0 space-y-4`,
            children: [(0, $.jsxs)(Jj, {
              children: [(0, $.jsxs)(`div`, {
                className: `flex flex-wrap items-center justify-between gap-2`,
                children: [(0, $.jsx)(`h2`, {
                  className: `font-display text-lg font-semibold`,
                  children: `年度資金狀況`
                }), (0, $.jsxs)(`div`, {
                  className: `flex flex-wrap gap-2`,
                  children: [(0, $.jsxs)(Uj, {
                    tone: _ ? `ok` : `watch`,
                    children: [`營運`, _ ? `可覆蓋（含期初現金）` : `仍缺資金`]
                  }), (0, $.jsx)(Uj, {
                    tone: g ? `ok` : `bad`,
                    children: g ? `含 ATM 後期末現金為正` : `含 ATM 後仍需額外資金`
                  })]
                })]
              }), (0, $.jsxs)(`div`, {
                className: `mt-4 grid gap-4 lg:grid-cols-2`,
                children: [(0, $.jsxs)(`div`, {
                  children: [(0, $.jsx)(`p`, {
                    className: `mb-2 text-xs text-muted`,
                    children: `來源分層：RPO 現金／非 RPO／客戶預付／ATM。ATM 不是營運來源。`
                  }), (0, $.jsx)(`div`, {
                    className: `h-56`,
                    children: (0, $.jsx)(Gs, {
                      width: `100%`,
                      height: `100%`,
                      children: (0, $.jsxs)(Lk, {
                        data: d.years.map(e => ({
                          name: e.year,
                          RPO現金: Number(e.rpoCash.toFixed(1)),
                          非RPO: Number(e.legacy.toFixed(1)),
                          客戶預付: Number(e.external.toFixed(1)),
                          ATM: Number(e.atm.toFixed(1)),
                          用途: Number(e.uses.toFixed(1))
                        })),
                        children: [(0, $.jsx)(uD, {
                          stroke: `#ddd6c8`,
                          strokeDasharray: `3 3`
                        }), (0, $.jsx)(XD, {
                          dataKey: `name`,
                          tick: {
                            fontSize: 11
                          }
                        }), (0, $.jsx)(pO, {
                          tick: {
                            fontSize: 11
                          }
                        }), (0, $.jsx)(Es, {}), (0, $.jsx)(ow, {
                          dataKey: `RPO現金`,
                          stackId: `s`,
                          fill: `#0f5c61`
                        }), (0, $.jsx)(ow, {
                          dataKey: `非RPO`,
                          stackId: `s`,
                          fill: `#3d8b8f`
                        }), (0, $.jsx)(ow, {
                          dataKey: `客戶預付`,
                          stackId: `s`,
                          fill: `#8fb9bb`
                        }), (0, $.jsx)(ow, {
                          dataKey: `ATM`,
                          stackId: `s`,
                          fill: `#c4a35a`
                        }), (0, $.jsx)(ow, {
                          dataKey: `用途`,
                          fill: `#5b5778`
                        })]
                      })
                    })
                  })]
                }), (0, $.jsx)(`div`, {
                  className: `h-56`,
                  children: (0, $.jsx)(Gs, {
                    width: `100%`,
                    height: `100%`,
                    children: (0, $.jsxs)(Rk, {
                      data: d.years.map(e => ({
                        name: e.year,
                        累積現金: Number(e.cum.toFixed(1))
                      })),
                      children: [(0, $.jsx)(uD, {
                        stroke: `#ddd6c8`,
                        strokeDasharray: `3 3`
                      }), (0, $.jsx)(XD, {
                        dataKey: `name`,
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(pO, {
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(Es, {}), (0, $.jsx)(ND, {
                        type: `monotone`,
                        dataKey: `累積現金`,
                        stroke: `#0f5c61`,
                        fill: `#d7ecec`
                      })]
                    })
                  })
                })]
              })]
            }), (0, $.jsx)(`div`, {
              className: `w-full max-w-full overflow-x-auto`,
              children: (0, $.jsx)(`div`, {
                className: `flex w-max gap-1 rounded-lg bg-surface p-1`,
                children: NM.map((e, t) => (0, $.jsx)(`button`, {
                  onClick: () => r(t),
                  className: Vj(`h-10 shrink-0 rounded-md px-3 text-sm font-medium`, n === t ? `bg-card text-fg shadow-card` : `text-muted hover:text-fg`),
                  children: e
                }, e))
              })
            }), (0, $.jsxs)(Jj, {
              className: `max-w-full overflow-x-auto`,
              children: [n === 0 && (0, $.jsx)(BM, {
                rows: [
                  [`Gross CapEx`, e.a.capex, (e, t) => E(`capex`, e, t)],
                  [`客戶預付／BYOH %`, e.a.customerFund.map(e => e * 100), (e, t) => E(`customerFund`, e, t / 100)],
                  [`表外現金租金`, e.a.newLease, (e, t) => E(`newLease`, e, t)],
                  [`在帳現金租金（固定）`, [...Kk], void 0],
                  [`基本利息`, e.a.interest, (e, t) => E(`interest`, e, t)],
                  [`股息`, e.a.div, (e, t) => E(`div`, e, t)],
                  [`非 RPO 現金`, e.a.legacy, (e, t) => E(`legacy`, e, t)]
                ]
              }), n === 1 && (0, $.jsx)(BM, {
                rows: [
                  [`排程 RPO`, d.years.map(e => e.scheduled)],
                  [`AI 排程`, d.years.map(e => e.ai)],
                  [`容量上限`, d.years.map(e => e.capacity)],
                  [`平均在役 MW`, d.years.map(e => e.avgBillable)],
                  [`產能瓶頸`, d.years.map(e => e.bottleneck)],
                  [`收入（產能後）`, d.years.map(e => e.revenue)],
                  [`信用損失`, d.years.map(e => e.loss)],
                  [`RPO 現金`, d.years.map(e => e.rpoCash)],
                  [`客戶預付／BYOH`, d.years.map(e => e.external)],
                  [`ATM（融資，非營運）`, d.years.map(e => e.atm)],
                  [`營運來源`, d.years.map(e => e.sourcesOp)],
                  [`總來源（含 ATM）`, d.years.map(e => e.sources)],
                  [`總用途`, d.years.map(e => e.uses)],
                  [`營運缺口`, d.years.map(e => e.operatingGap)],
                  [`融資後年度缺口`, d.years.map(e => e.gap)],
                  [`累積現金`, d.years.map(e => e.cum)]
                ]
              }), n === 2 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: `收入 = 排程 RPO − max(0, AI 排程 − 在役 MW×收入/MW×利用率)。預設在役＝年末 Billable；打開平均在役則用期初／期末平均。非 AI RPO 不受 MW 限制。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`Accepted MW`, d.m.accepted, (e, t) => D(`accepted`, e, t)],
                    [`Billable MW`, d.m.billable, (e, t) => D(`billable`, e, t)],
                    [`利用率 %`, e.m.util, (e, t) => D(`util`, e, t)],
                    [`收入 US$bn/MW`, e.m.revMW, (e, t) => D(`revMW`, e, t), 5e-4],
                    [`AI 占 RPO %`, e.m.aiShare, (e, t) => D(`aiShare`, e, t)]
                  ]
                })]
              }), n === 3 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: `電力 = Accepted×8760×PUE×電價／1e9（十億美元）。維護 = Accepted×單價／1000。Overlay 關閉時不進用途。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`Accepted MW`, d.m.accepted, (e, t) => D(`accepted`, e, t)],
                    [`電價 $/MWh`, e.m.power, (e, t) => D(`power`, e, t)],
                    [`PUE`, e.m.pue, (e, t) => D(`pue`, e, t), .01],
                    [`維護 US$m/MW`, e.m.maint, (e, t) => D(`maint`, e, t), .01],
                    [`電力（計算）`, d.years.map(e => e.power)],
                    [`維護（計算）`, d.years.map(e => e.maint)]
                  ]
                })]
              }), n === 4 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: `信用損失 = 收入 × 違約率 × (1−回收率)。不再另扣收款率，避免雙重計算。全包利率是透支用；存量利息在「年度假設」。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`客戶違約率 %`, e.m.defaultP, (e, t) => D(`defaultP`, e, t)],
                    [`回收率 %`, e.m.recovery, (e, t) => D(`recovery`, e, t)],
                    [`透支／邊際利率 %`, e.m.rate, (e, t) => D(`rate`, e, t)],
                    [`客戶預付／BYOH %`, e.a.customerFund.map(e => e * 100), (e, t) => E(`customerFund`, e, t / 100)]
                  ]
                })]
              }), n === 5 && (0, $.jsx)(`table`, {
                className: `w-full min-w-[480px] text-sm`,
                children: (0, $.jsxs)(`tbody`, {
                  children: [
                    [
                      [`資產殘值 %`, `residual`],
                      [`再出租率 %`, `rerent`],
                      [`剩餘 RPO 現金率 %`, `margin`],
                      [`租賃剩餘年數`, `residualLeaseYears`]
                    ].map(([n, r]) => (0, $.jsxs)(`tr`, {
                      className: `border-b border-border`,
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: n
                      }), (0, $.jsx)(`td`, {
                        className: `py-2`,
                        children: (0, $.jsx)(IM, {
                          value: e.terminal[r],
                          onChange: e => t(t => ({
                            ...t,
                            terminal: {
                              ...t.terminal,
                              [r]: e
                            },
                            scenario: `custom`
                          }))
                        })
                      })]
                    }, r)), (0, $.jsxs)(`tr`, {
                      className: `border-b border-border`,
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: `連動租賃尾端`
                      }), (0, $.jsxs)(`td`, {
                        className: `py-2 font-mono tabular-nums`,
                        children: [Y(d.leaseTail), `bn =（在帳 FY31 `, Kk[4], ` + 表外現金租金 `, e.a.newLease[4], `）×`, ` `, e.terminal.residualLeaseYears, ` 年。五年表外現金 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), ` vs 承諾 288。`]
                      })]
                    }), (0, $.jsxs)(`tr`, {
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: `終值淨額`
                      }), (0, $.jsxs)(`td`, {
                        className: `py-2 font-mono tabular-nums`,
                        children: [Y(d.totals.terminal), `bn`]
                      })]
                    })
                  ]
                })
              }), n === 6 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: `具名園區可編輯。最後一列是殘差＝年度 Accepted − 具名加總。殘差占年度五成以上會在連動檢查標紅——容量預測有多少是「其他／未列名」。`
                }), (0, $.jsxs)(`table`, {
                  className: `w-full min-w-[960px] text-xs`,
                  children: [(0, $.jsx)(`thead`, {
                    children: (0, $.jsx)(`tr`, {
                      className: `text-muted`,
                      children: [`園區`, `營運商`, `Planned`, `Energized`, `Accepted`, `Billable`, `狀態`, `里程碑`, `日期`, `信心`].map(e => (0, $.jsx)(`th`, {
                        className: `px-1 py-2 text-left font-medium`,
                        children: e
                      }, e))
                    })
                  }), (0, $.jsx)(`tbody`, {
                    children: d.sites.map((e, t) => (0, $.jsx)(`tr`, {
                      className: `border-t border-border`,
                      children: [`name`, `operator`, `planned`, `energized`, `accepted`, `billable`, `status`, `next`, `date`, `confidence`].map(n => (0, $.jsx)(`td`, {
                        className: `px-1 py-1`,
                        children: e.residual || ![`planned`, `energized`, `accepted`, `billable`, `name`, `operator`, `status`, `next`, `date`, `confidence`].includes(n) ? (0, $.jsx)(`span`, {
                          className: Vj(`block px-1 py-2`, e.residual && `text-muted`),
                          children: String(e[n])
                        }) : typeof e[n] == `number` ? (0, $.jsx)(IM, {
                          value: e[n],
                          onChange: e => O(t, n, e),
                          step: 1
                        }) : (0, $.jsx)(`input`, {
                          value: String(e[n]),
                          onChange: e => O(t, n, e.target.value),
                          className: `h-9 w-full rounded-sm border border-border bg-card px-2 text-xs`
                        })
                      }, n))
                    }, e.id))
                  })]
                })]
              }), n === 7 && (0, $.jsxs)(`div`, {
                className: `h-80`,
                children: [(0, $.jsx)(Gs, {
                  width: `100%`,
                  height: `100%`,
                  children: (0, $.jsxs)(Lk, {
                    layout: `vertical`,
                    data: m.map(e => ({
                      name: e.name,
                      不利: Number(Math.min(e.low, e.high).toFixed(1)),
                      有利: Number(Math.max(e.low, e.high).toFixed(1))
                    })),
                    margin: {
                      left: 24,
                      right: 12
                    },
                    children: [(0, $.jsx)(uD, {
                      stroke: `#ddd6c8`,
                      strokeDasharray: `3 3`
                    }), (0, $.jsx)(XD, {
                      type: `number`,
                      tick: {
                        fontSize: 11
                      }
                    }), (0, $.jsx)(pO, {
                      type: `category`,
                      dataKey: `name`,
                      width: 150,
                      tick: {
                        fontSize: 11
                      }
                    }), (0, $.jsx)(Es, {}), (0, $.jsx)(ow, {
                      dataKey: `不利`,
                      fill: `#9f1239`
                    }), (0, $.jsx)(ow, {
                      dataKey: `有利`,
                      fill: `#0f6b4c`
                    })]
                  })
                }), (0, $.jsxs)(`p`, {
                  className: `mt-2 text-xs text-muted`,
                  children: [`指標為五年期末現金（含期初與 ATM）。基準 `, Y(d.totals.end), `bn。`]
                })]
              }), n === 8 && (0, $.jsxs)(`ul`, {
                className: `space-y-2`,
                children: [d.checks.map(e => (0, $.jsxs)(`li`, {
                  className: `rounded-md border border-border bg-surface px-3 py-2`,
                  children: [(0, $.jsxs)(`div`, {
                    className: `flex items-center gap-2`,
                    children: [(0, $.jsx)(Uj, {
                      tone: e.ok ? e.severity === `watch` ? `watch` : `ok` : e.severity === `block` ? `bad` : `watch`,
                      children: e.ok ? e.severity === `watch` ? `觀察` : `通過` : `不一致`
                    }), (0, $.jsx)(`span`, {
                      className: `text-sm font-medium`,
                      children: e.title
                    })]
                  }), (0, $.jsx)(`p`, {
                    className: `mt-1 text-xs leading-relaxed text-muted`,
                    children: e.detail
                  })]
                }, e.id)), p.map(e => (0, $.jsxs)(`li`, {
                  className: `rounded-md border border-border bg-surface px-3 py-2`,
                  children: [(0, $.jsxs)(`div`, {
                    className: `flex items-center gap-2`,
                    children: [(0, $.jsx)(Uj, {
                      tone: e.ok ? e.severity === `watch` ? `watch` : `ok` : `bad`,
                      children: e.ok ? e.severity === `watch` ? `觀察` : `通過` : `不一致`
                    }), (0, $.jsx)(`span`, {
                      className: `text-sm font-medium`,
                      children: e.title
                    })]
                  }), (0, $.jsx)(`p`, {
                    className: `mt-1 text-xs leading-relaxed text-muted`,
                    children: e.detail
                  })]
                }, e.id))]
              }), n === 9 && (0, $.jsxs)(`div`, {
                className: `space-y-4 text-sm leading-relaxed`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: `時序先法說、後 10-Q。入帳數字聽 10-Q；全年指引聽法說；事實打架聽 10-Q。Q1 的 MW／利用率當觀察，不年化進基準。`
                }), (0, $.jsx)(`div`, {
                  className: `overflow-x-auto`,
                  children: (0, $.jsxs)(`table`, {
                    className: `w-full min-w-[720px] text-xs`,
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsxs)(`tr`, {
                        className: `border-b border-border text-left text-muted`,
                        children: [(0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `項目`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `法說（前瞻）`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `10-Q（入帳）`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 font-medium`,
                          children: `本模型取捨`
                        })]
                      })
                    }), (0, $.jsx)(`tbody`, {
                      className: `font-mono tabular-nums`,
                      children: [
                        [`RPO`, `一半 36 個月內轉營收`, `664 · 13/37/34/16＝84%`, `聽 10-Q 精確桶；法說＝13+37 的口語`],
                        [`期初現金`, `未給 YE 拆解`, `期末 37.077；YE 31.894`, `期初用 YE 31.9，避免 ATM／Q1 capex 雙計`],
                        [`ATM`, `約 $20bn 已完成`, `19.909／141m 股`, `聽 10-Q：19.9；股數 3.024bn`],
                        [`FY27 毛 CapEx`, `維持 90–95，非線性`, `只說高於 FY26；Q1 已 28.499`, `聽法說指引：毛 92.5。Q1 前載不年化成 114`],
                        [`淨現金 CapEx`, `全年 ≤70；Q1 約 18`, `Q1：28.499−11.363＝17.1`, `全年聽法說天花板 70；Q1 聽 10-Q，不把 18 當輸入`],
                        [`客戶資金`, `新 AI 約 30「不需 Oracle 現金」`, `Q1 預付 11.363`, `維持 24%（淨 70）。$30bn 不併入 RPO；Clay 改口後對得上 10-Q`],
                        [`在帳租賃`, `未提`, `已付 1.136＋剩餘 3.839；負債 43.8`, `聽 10-Q 到期表`],
                        [`表外租賃 $288bn`, `完全沒提`, `Note 6：已簽約未起租、15–19 年`, `聽 10-Q。這是法說最大的沉默`],
                        [`Q1 新 ROU`, `未提`, `營業 4.903＋融資 1.539`, `聽 10-Q；已含在帳，288 還沒 ROU`],
                        [`股數／淨負債`, `ATM 約 20bn`, `流通 3.024；Q1 現貨淨負債 88.3`, `股數聽 10-Q。淨負債用 YE 97.6，與期初現金同日`],
                        [`稅率／利息`, `未重述`, `稅 15.1%；Q1 利息 1.428`, `聽 10-Q：稅 15.1、FY27 利息 5.7`],
                        [`Q1 交付 MW`, `850 MW／>30 萬 GPU`, `無園區 MW`, `寫進園區與觀察。不年化進 Accepted`],
                        [`GPU 利用率`, `Q1 艦隊 97.9%；續約 +20%`, `無`, `基準全年 95：時點不是年均。不擬合 97.9`],
                        [`Abilene`, `六／八棟 618 MW＝75%`, `無`, `採用：具名實績，10-Q 不衝突。1.2 GW 留後續棟`],
                        [`FY27 Accepted`, `未給全年存量`, `無`, `維持 2,200 MW（v3.5）。Q1 爆發主要是 Abilene 補交`],
                        [`營收／EPS 指引`, `至少 $90bn；n/g EPS $8.10`, `無年度指引`, `營收當下限、不倒推產能。n/g EPS 不強制擬合 GAAP 模型`],
                        [`FCF 轉正`, `拒給時點`, `Q1 FCF −5.4`, `不編時點。模型五年現金仍為負`]
                      ].map(([e, t, n, r]) => (0, $.jsxs)(`tr`, {
                        className: `border-b border-border/70`,
                        children: [(0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3 font-sans font-medium`,
                          children: e
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3 text-muted`,
                          children: t
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3`,
                          children: n
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 font-sans text-muted`,
                          children: r
                        })]
                      }, e))
                    })]
                  })
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `$288bn 租賃承諾：法律效力與毀約`
                  }), (0, $.jsxs)(`div`, {
                    className: `mt-2 space-y-2 text-xs leading-relaxed text-muted`,
                    children: [(0, $.jsx)(`p`, {
                      children: `10-Q Note 6 原文：截至 2026-08-31「we had $288 billion of additional lease commitments, substantially all related to data center arrangements, that are generally expected to commence between the second quarter of fiscal 2027 and fiscal 2029 and for terms of fifteen to nineteen years that were not reflected on our condensed consolidated balance sheets … or in the maturities table.」MD&A 補一句：Q1「entered into certain significant leases for data centers and other contractual commitments。」`
                    }), (0, $.jsxs)(`p`, {
                      children: [(0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `這是契約，不是或有事項。`
                      }), `ASC 842 規定租賃在 commencement（資產可供使用）才入表。已簽約、機房還沒交屋 → 表外揭露。同一註腳把無條件採購 $34.2bn 寫成「enforceable and legally binding」；租賃承諾放在同一節「LEASES AND OTHER COMMITMENTS」，實務上視為已簽署、待起租。`]
                    }), (0, $.jsxs)(`p`, {
                      children: [`「generally expected to commence」是時間對沖：許可、電力、施工可以推遲起租，`, (0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `不自動解除合約`
                      }), `。`]
                    }), (0, $.jsxs)(`p`, {
                      children: [(0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `毀約代價 10-Q 沒有揭露`
                      }), `——沒有 termination for convenience、沒有違約金、沒有 walk-away 數字。典型大型機房／build-to-suit：`]
                    }), (0, $.jsxs)(`ul`, {
                      className: `list-disc space-y-1 pl-4`,
                      children: [(0, $.jsx)(`li`, {
                        children: `起租前（施工期）：開發協議。走掉 ≈ 地主已投入建造成本、開發費、可能的 take-or-pay 電力。不是明天付 $288bn。`
                      }), (0, $.jsx)(`li`, {
                        children: `起租後：幾乎 hell-or-high-water。違約 → 剩餘租金（或現值）加速到期。這才接近未折現剩餘租金。`
                      })]
                    }), (0, $.jsxs)(`p`, {
                      children: [`$288bn 是`, (0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `未折現未來租金`
                      }), `（對齊到期表口徑）。在帳租賃未折現 $63.2bn、負債現值 $43.8bn。$288bn 若平均 17 年、6% 折現，PV 約 $160–180bn；起租時以 PV 列 ROU／負債，不是 288 一次入帳。模型把這筆當「誰出錢」的現金租賃路徑（新增租賃五年+尾），既不假設明天付 288 現金，也不假設可以零成本取消。`]
                    })]
                  })]
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `ROU 是什麼？Q1「新取得 ROU」又是什麼？`
                  }), (0, $.jsxs)(`div`, {
                    className: `mt-2 space-y-2 text-xs leading-relaxed text-muted`,
                    children: [(0, $.jsxs)(`p`, {
                      children: [`ROU = Right-of-Use，`, (0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `使用權資產`
                      }), `（ASC 842）。承租人在租期內使用機房／辦公室的權利，對應一筆租賃負債。不是現金、不是 capex。`]
                    }), (0, $.jsxs)(`ul`, {
                      className: `list-disc space-y-1 pl-4`,
                      children: [(0, $.jsx)(`li`, {
                        children: `營業租賃 ROU：獨立列 $33.967bn（5/31 為 29.690）。`
                      }), (0, $.jsx)(`li`, {
                        children: `融資租賃 ROU：進 PPE $8.856bn（5/31 為 7.464）。`
                      }), (0, $.jsxs)(`li`, {
                        children: [`Q1 新取得 ROU＝「ROU assets obtained in exchange for lease obligations」：營業 $4.903bn + 融資 $1.539bn = `, (0, $.jsx)(`span`, {
                          className: `font-medium text-fg`,
                          children: `$6.442bn`
                        }), `。這是本季已起租（或修改增加）的入帳現值。`]
                      }), (0, $.jsx)(`li`, {
                        children: `同一季現金租賃支出只有 $1.136bn（營業 0.950 + 融資 0.186）。`
                      })]
                    }), (0, $.jsx)(`p`, {
                      children: `三層不要混：`
                    }), (0, $.jsxs)(`ol`, {
                      className: `list-decimal space-y-1 pl-4`,
                      children: [(0, $.jsx)(`li`, {
                        children: `已入帳租賃（ROU ≈ $42.8bn／負債 $43.8bn）→ 模型「在帳租賃」現金。`
                      }), (0, $.jsx)(`li`, {
                        children: `Q1 新取得 ROU $6.4bn → 已含在 (1) 的增量。`
                      }), (0, $.jsx)(`li`, {
                        children: `已簽約未起租 $288bn → 還沒有 ROU，模型「新增租賃」。`
                      })]
                    })]
                  })]
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `法說 vs 10-Q：一致、改口、含糊`
                  }), (0, $.jsx)(`div`, {
                    className: `mt-2 overflow-x-auto`,
                    children: (0, $.jsxs)(`table`, {
                      className: `w-full min-w-[640px] text-xs`,
                      children: [(0, $.jsx)(`thead`, {
                        children: (0, $.jsxs)(`tr`, {
                          className: `border-b border-border text-left text-muted`,
                          children: [(0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `項目`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `法說`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `10-Q`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 font-medium`,
                            children: `判定`
                          })]
                        })
                      }), (0, $.jsx)(`tbody`, {
                        children: [
                          [`RPO 36 個月一半`, `明確`, `13%+37%=50%`, `一致。精確桶聽 10-Q`],
                          [`CapEx 90–95`, `確認全年、非線性`, `只說高於 FY26`, `指引聽法說；Q1 28.5 不年化`],
                          [`淨現金 CapEx ≤70`, `確認；Q1 約 18`, `無全年數字；Q1 數學 17.1`, `全年天花板聽法說；Q1 聽 10-Q`],
                          [`不需 Oracle 資本`, `Hilary 用 capital；Clay 改口「是 CapEx、不是 Oracle 現金」`, `毛 capex 28.5、預付 11.4`, `改口才對得上 10-Q。不因此下修 capex`],
                          [`新 AI 約 $30bn vs RPO +26`, `未對帳`, `638→664`, `RPO 聽 10-Q。$30bn 不當新 RPO`],
                          [`表外租賃 $288bn`, `完全沒提`, `Note 6 全文`, `聽 10-Q。法說最大沉默`],
                          [`FCF 轉正`, `拒給時點；ramp 後「約 100% 稅後 EBITDA 轉 FCF」`, `Q1 FCF −5.4`, `不編時點`],
                          [`Abilene 618 MW／75%`, `六／八棟`, `無園區 MW`, `具名實績，採用`],
                          [`NM／WI`, `施工正常、不影響 FY27 營收`, `未點名`, `Accepted 維持 0；許可仍是風險`],
                          [`GPU 97.9%`, `Q1 艦隊`, `無`, `時點。基準全年 95，不擬合峰值`],
                          [`Q1 交付 850 MW`, `3× Q4、73% of FY26`, `無`, `觀察。不年化進 2.8 GW`],
                          [`營收 ≥90／EPS 8.10`, `上修`, `無年度指引`, `營收當下限；n/g 不強制擬合`],
                          [`毛利率`, `今年下台階、北極星改營業利益率`, `GAAP 營業利益率 35%`, `未給 GM%。不改貢獻率`]
                        ].map(([e, t, n, r]) => (0, $.jsxs)(`tr`, {
                          className: `border-b border-border/70`,
                          children: [(0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 font-medium`,
                            children: e
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: t
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: n
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5`,
                            children: r
                          })]
                        }, e))
                      })]
                    })
                  })]
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `FY2026 末（10-K／10-Q 比較欄）：RPO 638bn、現金及有價證券 31.894bn、票據 7.199+122.342、淨負債 97.6bn、流通 2.880bn 股、FY26 capex 約 55.7bn。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `Q1 FY27 10-Q（2026-08-31，2026-09-11 申報）：營收 19.345bn +30%；OCI 7.388 +121%；Apps 4.219 +10%；授權 0.655 −15%；支援 4.895 −1%；GAAP 營業利益 6.728（利潤率 35%）；淨利 4.760；稀釋 EPS 1.56；稅率 15.1%；SBC 1.127；利息 1.428；折舊 3.156。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `Q1 現金：CFO 23.103（含重大融資成分預付 11.363 + 其他遞延 3.997）；現金 capex 28.499、未付 capex 6.247；FCF −5.396。期末現金+有價證券 37.077。ATM 用盡：141m 股、淨額 19.909bn。無庫藏股；授權剩餘 6.3bn。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `RPO 664bn。認列桶：次 12 個月 13%、月 13–36 為 37%、月 37–60 為 34%、其後 16%（五年 84%）。法說：約一半在未來 36 個月轉營收。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `租賃：在帳營業負債 34.621 + 融資 9.185。營業 ROU 33.967、融資 ROU 8.856。Q1 現金租賃 1.136；Q1 新取得 ROU 4.903+1.539。表外追加機房租賃承諾 288bn（未折現、尚未 commence），多半 Q2 FY27–FY29 起租、15–19 年。無條件採購／電力 34.150bn。`
                }), (0, $.jsx)(VM, {
                  tag: `Interested-party`,
                  text: `法說（Hilary）：全年毛 capex 維持 90–95、淨現金 capex 不超過 70、Q1 淨約 18；FY27 營收至少 90bn（+34%）、non-GAAP EPS 8.10；Q2 營收 +30–34%、雲 +65–71%、n/g EPS 1.85–1.93。FCF 轉正時點拒給。`
                }), (0, $.jsx)(VM, {
                  tag: `Interested-party`,
                  text: `法說（Clay）：Q1 交付 850 MW／>30 萬 GPU（約為 FY26 全年 73%）；GPU 利用率 97.9%；到期 GPU +20% 溢價續約。Abilene 六棟 618 MW（75%）。Shackelford／NM／WI／Michigan Q1 皆未交付；NM、WI 稱不影響 FY27 營收。新 AI 約 30bn 不需 Oracle 現金——隨即改口「仍是 CapEx，只是現金不從 Oracle 出」。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `Ellison 10b5-1（2026-06-22 至 10-24）最多出售 50m 股。Ken Bond 本季為最後一次法說；Investor Day 2026-10；下次財報 2026-12-14。`
                }), (0, $.jsx)(VM, {
                  tag: `User-provided`,
                  text: `5Y Senior CDS ${Y(e.cds,3)}bps（${e.cdsDate}），Bid/Ask ${Y(e.cdsBid,3)}/${Y(e.cdsAsk,3)}。中價由買賣報價平均。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `Michigan／Saline Township（The Barn）：OpenAI、Oracle、Related Digital 於 2025-10-30 公告 1+ GW；DTE 1.4 GW。法說點名 Q1 未交付。模型 Planned 1,200 MW。`
                }), (0, $.jsx)(VM, {
                  tag: `Assumed`,
                  text: `FY28–FY31 capex 下降、legacy 現金 25bn、貢獻率 35%、rev/MW、客戶違約路徑。新增租賃路徑用來對上 288bn 表外承諾的現金型態，不是逐筆租約。FY27 Accepted 維持 2,200 MW：Q1 850 MW 是增量實績，不年化。FY27 利用率 95%：Q1 GPU 97.9% 是時點，不是全年峰值。`
                }), (0, $.jsxs)(`p`, {
                  className: `text-xs text-muted`,
                  children: [`損益兩平貢獻率 `, hA(d.totals.be * 100), `。終值 `, mA(d.totals.terminal), `bn。本工具是現金資金模型，不是評等預測或投資建議。`]
                })]
              })]
            }), (0, $.jsxs)(Jj, {
              className: `bg-ink text-accent-fg`,
              children: [(0, $.jsx)(`h2`, {
                className: `font-display text-2xl font-semibold`,
                children: g ? `目前假設下五年期末現金為正。連動評價 $${Y(f.call.blended,0)}，結論「${f.call.call}」。` : `目前假設下五年仍缺 ${mA(Math.abs(d.totals.end))}bn。連動評價 $${Y(f.call.blended,0)}，結論鎖定「${f.call.call}」，不會是買進。`
              }), (0, $.jsxs)(`p`, {
                className: `mt-2 text-sm leading-relaxed text-accent-soft`,
                children: [`營運缺口（不含 ATM、含 capex／租賃／利息／股息）為 `, mA(d.totals.operatingGap), `bn。 自給與否請同時看「營運」與「含已完成股權」兩欄。CDS `, Y(e.cds, 3), ` bps。SOTP／本益比已扣除未籌資金，DCF 已內含 Cash CapEx。`]
              })]
            })]
          })]
        })]
      })]
    })]
  })
}

function BM({
  rows: e,
  step: t = .1
}) {
  return (0, $.jsx)(`div`, {
    className: `max-w-full overflow-x-auto`,
    children: (0, $.jsxs)(`table`, {
      className: `w-full min-w-[720px] text-sm`,
      children: [(0, $.jsx)(`thead`, {
        children: (0, $.jsxs)(`tr`, {
          children: [(0, $.jsx)(`th`, {
            className: `py-2 text-left font-medium text-muted`,
            children: `US$bn / 單位`
          }), zk.map(e => (0, $.jsx)(`th`, {
            className: `py-2 text-right font-medium text-muted`,
            children: e
          }, e))]
        })
      }), (0, $.jsx)(`tbody`, {
        children: e.map(([e, n, r, i]) => (0, $.jsxs)(`tr`, {
          className: `border-t border-border`,
          children: [(0, $.jsx)(`td`, {
            className: `py-1.5 font-medium`,
            children: e
          }), n.map((e, n) => (0, $.jsx)(`td`, {
            className: `py-1.5 text-right`,
            children: r ? (0, $.jsx)(IM, {
              value: e,
              onChange: e => r(n, e),
              step: i ?? t
            }) : (0, $.jsx)(`span`, {
              className: `font-mono text-xs tabular-nums`,
              children: Y(e)
            })
          }, n))]
        }, e))
      })]
    })
  })
}

function VM({
  tag: e,
  text: t
}) {
  return (0, $.jsxs)(`div`, {
    className: `rounded-md border-l-4 border-accent bg-surface px-3 py-2`,
    children: [(0, $.jsx)(`span`, {
      className: `font-mono text-xs text-accent`,
      children: e
    }), (0, $.jsx)(`p`, {
      className: `mt-1 text-xs leading-relaxed text-fg`,
      children: t
    })]
  })
}(0, y.createRoot)(document.getElementById(`root`)).render((0, $.jsx)(v.StrictMode, {
  children: (0, $.jsx)(zM, {})
}));