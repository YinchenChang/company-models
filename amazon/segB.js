PERIOD_LABELS = PERIODS, // 接續模板片段 mid1 結尾未結束的宣告鏈，所以不可加 var
  HIST_PL = COMPANY_DATA.historicalPL,
  CAPM_Q = ((v, b) => { // v0.1b（Oracle）：WACC 以 CAPM 計算（company.json → valuation.capm）；b＝β（敏感度用）
    const c = v.capm, E = v.price * COMPANY_DATA.latestQuarter.sharesOut, D = COMPANY_DATA.latestQuarter.debtPrincipal, be = b ?? c.beta,
      ke = v.rf + be * c.erp, kd = c.kdPretax * (1 - v.tax), wE = E / (D + E);
    return { beta: be, ke, kd, kdPre: c.kdPretax, E, D, wE, wD: 1 - wE, wacc: wE * ke + (1 - wE) * kd }
  }),
  MEDIAN_Q = a => { const s = [...a].sort((x, y) => x - y), n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2 },
  SW_MEDIAN = MEDIAN_Q((COMPANY_DATA.peers.software || []).map(x => x.ntmEvEbitda)), // 同業 NTM EV/EBITDA 中位數
  SEGM_Q = COMPANY_DATA.valuation.segmentMultiples || null, // MAG v0.1b：各分部同業倍數組（legacyBiz.lines.peer）
  segMultQ = (k, o) => { const g = SEGM_Q && SEGM_Q[k]; return !g || g.useAi ? o.evEbitda : MEDIAN_Q(g.peers.map(x => x.ntmEvEbitda)) },
  VAL_DEFAULTS = (v => (v.wacc = v.wacc ?? CAPM_Q(v).wacc, v.legacyEvEbitda = v.legacyEvEbitda ?? (SEGM_Q ? null : SW_MEDIAN), v))(structuredClone(COMPANY_DATA.valuation)), // MAG v0.1b：分部倍數時 null＝錨定年度各線 EBITDA 加權 // null＝採 CAPM／同業中位數；數字＝手動覆蓋
  lM = COMPANY_DATA.peers.list.map(e => ({ ...e, ev: e.mkt + e.netDebt, ebitda: e.opInc + e.da })); // 同業 Comps：資料只在 company.json → peers；EV 與 EBITDA 現算（與 Excel 公式一致）

function uM(e, t) {
  return t ? e / t - 1 : 0
}

function dM(e) {
  return Math.max(0, -e)
}
var fM = COMPANY_DATA.valuation.atmSharesInValuation ?? 0; // v0.1b：評價股數中「期後股權發行上限」的股數（ATM 關閉時扣回；CRWV 0.035、Nebius 0）

function shareCount(e, t) {
  let n = e.includeAtm ? 0 : fM;
  return Math.max(.1, t.shares - n)
}

function mM(e, t, n) {
  return Math.max(0, e - t / Math.max(n, .01))
}

function mR(e, t, n) {
  return e - t / Math.max(n, .01)
}

function forwardPL(e, t) {
  let NOL_USE = t.nolUsePct, // 5a：NOL 每年可抵用比例讀 company.json → valuation
    n = HIST_PL[2],
    r = n.revenue,
    s = t.shares,
    sb = t.shares,
    N = t.nol ?? 0,
    N2 = t.nol ?? 0;
  return PERIOD_LABELS.map((n, c) => {
    let L = PERIOD_YEARS[c],
      l = e.years[c].isRev,
      u = e.years[c].servicesRev,
      lg = e.years[c].legacyRev || 0, // v0.1b（Oracle）：非 AI 事業營收
      d = l + u + lg,
      f = c === 0 ? uM(HIST_PL[3].revenue + d, r) : uM(d, r),
      EB = (l + u) * e.years[c].ebM + (e.years[c].otherEbitda || 0) + (e.years[c].legacyEbitda || 0) - (e.years[c].delayPen || 0) + (t.read2 ? e.years[c].shadowEb1 || 0 : 0), // MAG v0.1b r3（C16）：讀法 2 加自用 AI 影子 EBITDA（k＝1）； // v0.2：減延誤罰則 // v0.1b：加其他事業 EBITDA；Oracle：EBITDA 率只套算力＋服務，非 AI 事業另計
      lgO = e.years[c].legacyOa || 0, // MAG v0.1b：非 AI 事業其他攤銷（影音內容、營業租賃資產等；進 D&A，現金上視為等額支出，UFCF 不加回）
      p = EB - e.years[c].daFleet - lgO,
      m = e.years[c].interest,
      ai = e.years[c].prepayAccr || 0, // v0.1b（Oracle）：預付重大財務組成的非現金利息（進稅前損益，不進現金）
      pt = p - m - ai,
      tb = Math.max(0, pt - Math.min(N, Math.max(0, pt) * NOL_USE)),
      tx = tb * t.tax,
      h = pt - tx;
    pt < 0 ? N += -pt : N -= Math.min(N, pt * NOL_USE);
    sb = c === 0 ? t.shares : sb * 1.01, s = sb + (e.years[c].cumNewShares || 0);
    let g = h / s,
      _ = (h + t.sbc * L) / s,
      v = e.years[c].daFleet + lgO,
      vF = e.years[c].daFleet,
      y = e.years[c].cashCapex,
      b = c === 0 && e.years[0].wcStub != null ? -e.years[0].wcStub : t.wcPctOfRevGrowth * Math.max(0, d - r), // MAG v0.1b r2（C14）：首期＝上一年度同期實際（defaults.wcStub，流入為正）；之後年度沿用營收增量比例
      tb2 = Math.max(0, p - Math.min(N2, p * NOL_USE)),
      x = p - tb2 * t.tax + vF - y - b - (e.years[c].prepayRecog || 0), // v0.1b：預付認列為非現金營收，自 UFCF 扣除（預付流入已在現金 CapEx 抵減）
      S = e.years[c].revenue,
      w = e.years[c].newRev,
      T = e.years[c].capacity > 0 ? e.years[c].unsold / e.years[c].capacity : 0,
      E = p + v;
    return N2 -= Math.min(N2, p * NOL_USE), r = c === 0 ? HIST_PL[3].revenue + d : d, {
      fyRevenue: c === 0 ? HIST_PL[3].revenue + d : d,
      fyOpInc: c === 0 ? HIST_PL[3].opInc + p : p,
      fyNi: c === 0 ? HIST_PL[3].ni + h : h,
      fyDa: c === 0 ? ACTUAL_1H.da + v : v,
      fyEbitda: c === 0 ? HIST_PL[3].opInc + ACTUAL_1H.da + EB : EB,
      fyEps: c === 0 ? HIST_PL[3].eps + h / s : h / s,
      fyNgEps: c === 0 ? HIST_PL[3].ngEps + (h + t.sbc * L) / s : (h + t.sbc * L) / s,
      year: n,
      gpu: l,
      services: u,
      legacy: lg,
      legacyEbitda: e.years[c].legacyEbitda || 0,
      segEb: (e.lg ? e.lg.lines : []).map(x => [x.peer, x.ebitda[c]]), // MAG v0.1b：各線 EBITDA（分部加總倍數）
      revenue: d,
      fundingRev: S,
      inYear: w,
      plugRatio: T,
      yoy: f,
      opInc: p,
      opM: p / d,
      ebitda: E,
      ni: h,
      tax: tx,
      nm: h / d,
      eps: g,
      ngEps: _,
      shares: s,
      da: v,
      cashCapex: y,
      capexGrowthCash: e.years[c].capexGrowthCash || 0,
      interest: m,
      prepayAccr: ai,
      ufcf: x
    }
  })
}

function ncdfQ(x) {
  let t = 1 / (1 + .2316419 * Math.abs(x)),
    d = .3989423 * Math.exp(-x * x / 2),
    p = d * t * (.3193815 + t * (-.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
  return x > 0 ? 1 - p : p
}

function bsCallQ(S, K, T, r, s) {
  if (S <= 0) return 0;
  if (K <= 0) return S;
  let d1 = (Math.log(S / K) + (r + s * s / 2) * T) / (s * Math.sqrt(T)),
    d2 = d1 - s * Math.sqrt(T);
  return S * ncdfQ(d1) - K * Math.exp(-r * T) * ncdfQ(d2)
}

function dcfValue(e, t) {
  let n = e.map(e => e.ufcf),
    r = n.map((e, n) => e / (1 + t.wacc) ** CALQ.tEnd[n]), // v4.5：評價日至各期期末（月數 ÷ 12）
    i = r.reduce((e, t) => e + t, 0),
    c = e[e.length - 1],
    l = t.tvBasis === `ufcf` ? c.ufcf + c.capexGrowthCash : c.opInc * (1 - t.tax) + c.da - c.da * t.maintRatio, // v0.1c（Oracle）：tvBasis＝ufcf 時終值以末期 UFCF（含汰換 CapEx）為基準，加回末期成長型 CapEx（扣客戶預付後；成長由 g 表達，預設情境末期為 0）；模板＝常態化 FCF（D&A × 維持比率）
    bad = t.wacc <= t.g ? `WACC ≤ 永續成長率，Gordon 終值無定義` : l <= 0 ? `常態化 FCF ≤ 0，終值無經濟意義` : ``,
    a = bad ? 0 : l * (1 + t.g) / (t.wacc - t.g),
    o = a / (1 + t.wacc) ** CALQ.tEnd[CALQ.tEnd.length - 1],
    s = i + o,
    u = s - t.netDebt,
    q = t.fund ? t.fund.eq.reduce((e, n, r) => e + n / (1 + t.wacc) ** CALQ.tStart[r], 0) : 0,
    z = t.shares + (t.fund ? t.fund.sh : 0),
    zr = Math.max(0, (u + q) / z),
    op = bsCallQ(s + q, t.netDebt, t.optT ?? 4.5, t.rf ?? .04, t.sigma ?? .5) / z,
    ps = bad ? NaN : (t.dcfMode === `option` ? op : zr);
  return {
    invalid: !!bad,
    invalidReason: bad,
    zeroPerShare: zr,
    optPerShare: op,
    pvFcf: i,
    tv: a,
    pvTv: o,
    ev: s,
    equity: u,
    pvEquityRaised: q,
    postShares: z,
    perShare: ps,
    // v4.5：DCF 腿以 WACC 推到目標價時點（評價日＋12 個月；0 截斷與選擇權兩種模式皆同），與 EV/EBITDA 腿同一時點；perShare 仍為評價日現值（反向 DCF 用）
    perShareT: ps * (1 + t.wacc) ** CALQ.targetT,
    perShareRaw: (u + q) / z,
    tvShare: o / Math.max(Math.abs(s), 1) * Math.sign(s || 1),
    normFcf: l,
    fcf: n,
    pv: r
  }
}

function evEbitdaLeg(e, t) {
  // v3.5：(錨定年 EBITDA × 倍數 − 錨定年末淨負債) ÷ 錨定年末股數，再以 WACC 折回 FY27 末（目標價時點）
  let k = Math.min(4, Math.max(1, Math.round(t.evYear ?? 1))),
    n = t.fund ? t.fund.ndY[k] : t.netDebt,
    r = t.fund ? e[k].shares : t.shares; // 損益股數：SBC 逐年複利稀釋＋瀑布新股（FY27 與 v3.4 相同）
  return (evSotpQ(e[k], t.evEbitda, t) - n) / r / Math.pow(1 + t.wacc, k - CALQ.evOffset) // v4.5：錨定期末折回目標價時點（calendar.targetHorizon）
}
// v0.1b（Oracle）：分部加總企業價值＝AI 雲端 EBITDA × AI 雲端 倍數 m＋非 AI 事業 EBITDA × 非 AI 事業倍數（同業 NTM 中位數）
function evSotpQ(f, m, t) {
  const lg = f.legacyEbitda || 0;
  return (f.ebitda - lg) * m + lg * (t.legacyEvEbitda ?? m)
}
// v3.5：錨定年度 × 倍數矩陣（加權目標價與 EV/EBITDA 腿，均為每股）
var EVG_MULTS = [3.4, 4.5, 5, 6, 7];
function evAnchorGrid(d, st, o) {
  let base = runValuation(d, st, o);
  let leg = EVG_MULTS.map(m => [1, 2, 3, 4].map(k => {
    return Math.max(0, (evSotpQ(base.fwd[k], m, base.v) - base.v.fund.ndY[k]) / base.fwd[k].shares / Math.pow(1 + o.wacc, k - CALQ.evOffset)) // MAG v0.1b：非 AI 倍數取該次評價的加權值
  }));
  let DI = base.d.invalid, wd = DI ? 0 : BLEND_W.dcf, dcf = DI ? 0 : base.d.perShareT;
  return {
    mults: EVG_MULTS,
    years: [1, 2, 3, 4].map(k => d.years[k].year),
    leg: leg,
    tgt: leg.map(r => r.map(x => wd * dcf + (1 - wd) * x)),
    dcf: dcf,
    wd: wd,
    legM: base.v.legacyEvEbitda // MAG v0.1b：非 AI 事業加權倍數（顯示用）
  }
}
var BLEND_W = COMPANY_DATA.methodology.blendWeights,
  bM = [{
    method: `DCF（融資後每股）`,
    capex: VAL_DEFAULTS.tvBasis === `ufcf` ? `已扣五期 Cash CapEx；終值用末期 UFCF（含穩態 GPU 汰換）` : `已扣五期 Cash CapEx；終值用常態化 FCF（D&A×維持比率）`,
    hole: `不扣缺口：新債視為公允價值、價值中性；新股募得現金折現後加回，股數同步增加`,
    shares: `含 ATM ＋ 瀑布新股`,
    netDebt: `評價日本金 ${Y(LATEST_Q.debtPrincipal, 2)} − 現金 ${Y(LATEST_Q.cash + LATEST_Q.marketable, 2)}（另含期後調整）`
  }, {
    method: `EV/EBITDA 分部加總（錨定年度，可選 ${PERIODS[1]}–${PERIODS[4]}）`,
    capex: `EBITDA 未扣 CapEx；AI 雲端 EBITDA × AI 雲端 倍數＋非 AI 事業 EBITDA × 同業 NTM 中位數`,
    hole: `不扣缺口：改用錨定年末淨負債（含瀑布新債）與股數（含瀑布新股）；${CALQ.evDiscText}`,
    shares: `錨定年末（含 SBC 稀釋與新股）`,
    netDebt: `錨定年末總債務 − 現金`
  }]

// 評等規則（v4.1 自 blendCall 抽出；v4.2 門檻改讀 company.json → methodology.rating）：c＝相對現價的空間、o＝終值占 EV、B＝股權募資超過現市值上限倍數
var RATE_TH = COMPANY_DATA.methodology.rating,
  SELL_TH = RATE_TH.sellUpsideMax;
function rateCall(c, o, B) {
  return c >= RATE_TH.buyUpsideMin && o < RATE_TH.buyTvShareMax && !B ? `買進` : c <= SELL_TH || B && c <= RATE_TH.sellUpsideMaxIfEquityOver ? `賣出` : `中立`
}
function pctQ(x) { // 門檻的百分比文字（取絕對值，正負號由呈現處決定）：0.25 → 25%、-0.15 → 15%
  return `${multTxt(Math.round(Math.abs(x) * 1e6) / 1e4)}%`
}

// MAG v0.1b r3（對照表 r1 C16）：讀法 2（自用 AI 價值中性）——同一組評價輸入（股數、淨負債、可轉債分類、倍數沿用讀法 1），EBITDA 加自用 AI 影子 EBITDA（k＝1），
// 影子收入進評價（DCF 的 EBIT、稅、UFCF、終值與 EV/EBITDA 錨定年 AI 雲端 EBITDA）；只有對外 AI 的超額報酬影響目標價。Excel「評價_DCF與目標價」讀法 2 區同式。
function read2Q(e, p) {
  const i = { ...p.v, read2: !0 }, a = forwardPL(e, i), o = dcfValue(a, i), c = evEbitdaLeg(a, i), d = Math.max(0, c),
    WD = o.invalid ? 0 : BLEND_W.dcf, tp = (o.invalid ? 0 : WD * o.perShareT) + (1 - WD) * d;
  return { tp, dcfT: o.perShareT, ev: d, diff: tp - p.call.blended, shadow: e.years.map(y => y.shadowEb1 || 0) }
}
// v0.1 交付前修訂：評價口徑敏感度（WACC 以 ERP 調整達成、ERP、非 AI 改公司自身 NTM 倍數、AI 雲端 15×、組合）與評等翻轉點（WACC 降到多少時由賣出轉中立）。
// Excel 為建置時快照（scripts/rv_solve.py「vs」「vflip」，同一演算法：ERP 對 WACC 線性；翻轉點 30 次對半）。
function erpForWaccQ(o, w) { const c = o.capm || COMPANY_DATA.valuation.capm, f = e => CAPM_Q({ ...o, capm: { ...c, erp: e } }).wacc, a = f(c.erp), b = f(c.erp + .01); return c.erp + (w - a) / (b - a) * .01 }
function valAtQ(d, st, o, chg) { let i = { ...o }; if (chg.erp != null) i = { ...i, capm: { ...(o.capm || COMPANY_DATA.valuation.capm), erp: chg.erp } }, i.wacc = CAPM_Q(i).wacc;
  if (chg.own) i.legacyEvEbitda = chg.own; if (chg.ai) i.evEbitda = chg.ai; const p = runValuation(d, st, i); return { tp: p.call.blended, call: p.call.call, wacc: i.wacc } }
function valSensQ(d, st, o) {
  const OM = COMPANY_DATA.valuation.ownMultiple, own = OM && OM.value, th = o.price * (1 + SELL_TH), ew = w => erpForWaccQ(o, w),
    R = [[`WACC 9%（ERP 調整）`, { erp: ew(.09) }], [`WACC 10%（ERP 調整）`, { erp: ew(.10) }], [`WACC 11%（ERP 調整）`, { erp: ew(.11) }], [`ERP 4%`, { erp: .04 }], [`ERP 6%`, { erp: .06 }],
      ...(own ? [[`非 AI 分部改用公司自身 NTM EV/EBITDA ${Y(own, 2)}×`, { own }]] : []), [`AI 雲端 15×`, { ai: 15 }],
      [`組合：WACC 9%＋${own ? `非 AI 自身倍數＋` : ``}AI 15×`, { erp: ew(.09), own, ai: 15 }]].map(([l, c]) => ({ label: l, erp: c.erp, ...valAtQ(d, st, o, c) })),
    base = valAtQ(d, st, o, {});
  let flip = null;
  if (base.tp < th) { let lo = .05, hi = o.wacc; const f = w => valAtQ(d, st, o, { erp: ew(w) }).tp - th;
    if (f(lo) >= 0) { for (let i = 0; i < 30; i++) { const m = (lo + hi) / 2; f(m) >= 0 ? lo = m : hi = m } flip = { wacc: (lo + hi) / 2, erp: ew((lo + hi) / 2) } } }
  return { rows: R, base, th, flip }
}
function blendCall(e) {
  let {
    spot: t,
    dcfPx: n,
    pePx: i,
    eqTotal: a,
    newSh: q,
    tvShare: o,
    mktCap: M,
    shares0: S0,
    dcfInvalid: DI
  } = e, WD = DI ? 0 : BLEND_W.dcf, WE = 1 - WD, s = (DI ? 0 : WD * n) + WE * i, c = s / Math.max(t, .01) - 1, u = a <= .01, B = a > RATE_TH.equityRaiseMaxMult * M, d = [];
  u ? d.push(`融資狀態：五期缺口可全由債務（受債務上限約束）支應，不需新股。`) : d.push(`融資狀態：五期需股權募資 ${a.toFixed(0)}bn（約現市值 ${(a/Math.max(M,1)).toFixed(1)} 倍），新發行 ${q.toFixed(2)}bn 股，原股東最終持股約 ${(S0/(S0+q)*100).toFixed(0)}%。股權在需要前一期募足（期前融資），不再以缺口本金扣減估值。`), B && d.push(`股權募資超過現市值 ${multTxt(RATE_TH.equityRaiseMaxMult)} 倍：市場吸收能力存疑，禁止買進。`), o > RATE_TH.tvShareWarn && d.push(`終值占企業價值過高，DCF 對 WACC／永續成長／維持性 CapEx 比率極敏感。`);
  let f = rateCall(c, o, B);
  DI && d.unshift(`DCF 失效，已排除於目標價（權重改為 EV/EBITDA 100%）。`);
  return {
    blended: s,
    upside: c,
    call: f,
    notes: d,
    hole: a,
    funded: u,
    blocked: B,
    weights: {
      dcf: WD,
      pe: WE
    },
    dcfInvalid: DI
  }
}

function dcfGridWaccG(e, t) {
  let n = [.09, .1, .11, .12, .13];
  return [.02, .025, .03, .035, .04].map(r => n.map(n => n <= r ? NaN : dcfValue(e, {
    ...t,
    wacc: n,
    g: r
  }).perShare))
}
function wM(e) {
  let t = [...e].filter(e => Number.isFinite(e)).sort((e, t) => e - t),
    n = Math.floor(t.length / 2);
  return t.length ? t.length % 2 ? t[n] : (t[n - 1] + t[n]) / 2 : NaN
}

// MAG v0.1b：AI 增量報酬（主命題）。各期（年化＝模型期 ÷ 期間長度）：AI 雲端 EBITDA＝對外 AI 雲端營收 × AI EBITDA 率；
// 投入資本（全 AI）：期初＝AI 期初毛額（IT＋機房）× 淨額比；期末＝期初＋AI 成長型＋汰換 − AI 折舊。
// MAG v0.1b r2（對照表 r1 C11）：主值＝對外 AI ROIC＝(對外 EBITDA − AI 折舊 × 對外比例)×(1 − 稅率) ÷ (平均投入資本 × 對外比例)（投入資本與折舊按對外 MW 比例分攤）；
// 對照＝全 AI ROIC（含影子收入）＝(對外 EBITDA＋影子收入 EBITDA − AI 折舊)×(1 − 稅率) ÷ 全 AI 平均投入資本；影子收入＝自用 AI MW（＝對外 ×(1 ÷ 對外比例 − 1)）× 每 MW 年收入 × 同一 EBITDA 率。
// 打平 k（錨定期，對外口徑）＝k ×(WACC × 平均投入資本 × 對外比例 ÷ (1 − 稅率)＋營運成本＋折舊 × 對外比例) ÷ 營收（AI EBITDA 對 k 線性，其他不變）。Excel「AI增量報酬」同式。
function aiRoicQ(d, st, o) {
  const CM = COMPANY_DATA.capexModel; if (!CM || CM.mode !== `tk`) return null;
  const Y = d.years, XS = st.extShare ?? CM.extShare, A0 = aiOpenQ(st, XS), tax = o.tax, P = d.m.price, k = P ? P[0].k : 1, ry = CM.roicYear ?? 3,
    sb = st.selfBuild ?? CM.selfBuild, RX = rentedExtQ(), FR = r => P ? facRentMWQ(P[r], A0.facLife) : 0;
  let ic = (A0.it + A0.fac) * (CM.aiNetShare ?? 1), out = { xs: XS, icBeg: [], icEnd: [], ebitda: [], da: [], daExt: [], icExt: [], opex: [], rev: [], nopat: [], roic: [], spread: [], shRev: [], shEbitda: [], roicSh: [],
    mwEff: [], mwRent: [], ownShare: [], facRentMW: [], rentFac: [], rentExt: [], ebitdaNet: [] };
  Y.forEach((y, r) => {
    const L = PERIOD_YEARS[r], rev = y.isRev / L, eb = y.isRev * y.ebM / L, da = y.daAi / L, opx = rev - eb, b = ic, e = b + y.capexAi + y.refresh - y.daAi, avg = (b + e) / 2,
      mwE = P ? rev / Math.max(P[r].hold * P[r].k / 1e3, 1e-12) : 0, Ra = ((r ? RX.path[r - 1] : RX.open) + RX.path[r]) / 2, own = Math.max(0, mwE - Ra), tot = Math.max(1e-9, mwE / XS - Ra),
      so = P ? own / tot : XS, fr = FR(r), rf = (1 - sb) * own * fr, rx = Ra * RX.rentMW / 1e3, ebx = eb - rf - rx,
      np = (ebx - da * so) * (1 - tax), sr = rev * (1 / XS - 1), se = sr * y.ebM, rfAll = (1 - sb) * tot * fr;
    ic = e;
    out.icBeg.push(b), out.icEnd.push(e), out.ebitda.push(eb), out.da.push(da), out.daExt.push(da * so), out.icExt.push(avg * so), out.opex.push(opx), out.rev.push(rev), out.nopat.push(np),
      out.roic.push(np / Math.max(1e-9, avg * so)), out.spread.push(np / Math.max(1e-9, avg * so) - o.wacc),
      out.mwEff.push(mwE), out.mwRent.push(Ra), out.ownShare.push(so), out.facRentMW.push(fr), out.rentFac.push(rf), out.rentExt.push(rx), out.ebitdaNet.push(ebx),
      out.shRev.push(sr), out.shEbitda.push(se), out.roicSh.push((eb + se - rfAll - rx - da) * (1 - tax) / Math.max(1e-9, avg))
  });
  out.k = k, out.ry = ry, out.breakevenK = k * (o.wacc * out.icExt[ry] / (1 - tax) + out.opex[ry] + out.rentFac[ry] + out.rentExt[ry] + out.daExt[ry]) / Math.max(out.rev[ry], 1e-9), out.wacc = o.wacc;
  out.c15 = c15Q(d, st, out, A0, tax);
  return out
}
// MAG v0.1b r3（對照表 r1 C20、C22）：租用對外 MW（capexModel.rentedExt：期初 open、各期末 path、每 MW 年租金 rentMW）——收入照算、投入資本與折舊不計、租金計入對外 AI 營運成本；
// 機房租用部分（1 − 自建比例）的租金＝在役世代加權 TK_CapexFacility × 資本回收係數（tokenomics.holdEconWacc、機房年限）——出租方打平租金 [Derived]。
// 兩項租金只進 AI 增量報酬（ROIC、打平 k、C15）：資金模型與評價的租金現金已在租賃（leases）與租用算力（rentedCompute）內，不重複扣。
function rentedExtQ() { const R = (COMPANY_DATA.capexModel || {}).rentedExt || {}, o = R.open || 0; return { open: o, path: R.path || PERIOD_YEARS.map(() => o), rentMW: R.rentMW || 0 } }
function facRentMWQ(pr, facLife) { const w = (COMPANY_DATA.tokenomics || {}).holdEconWacc ?? .1, crf = w / (1 - Math.pow(1 + w, -facLife));
  return PRICING.chips.reduce((a, c, j) => a + pr.gens[j] * TKV.IF_CapexFacility[c.tk], 0) / Math.max(pr.tot, 1e-9) / 1e3 * crf }
// MAG v0.1b r2（對照表 r1 C15）：一致性檢查——錨定期對外 AI ROIC 依序改成「k＝1、Tokenomics 成本、無爬坡延遲、無稅」，與 IF_HoldEcon 隱含報酬（Tokenomics WACC，tokenomics.holdEconWacc）比較；
// 各步差額依序相加＝總差距（序列拆解，順序固定）：稅 → k＝1（含自研晶片係數）→ 營運成本來源 → 汰換 → 爬坡／閒置（自有對外 MW 穩態，淨額比 ½）→ 機房與租用（全部自有、自建 100%、無租金；r3 C20／C22）→ 會計口徑殘差（平均淨帳面 vs 年金）。Excel「AI增量報酬」同式。
function c15Q(d, st, A, A0, tax) {
  const TKH = COMPANY_DATA.tokenomics && COMPANY_DATA.tokenomics.holdEconWacc, P = d.m.price; if (TKH == null || !P) return null;
  const ry = A.ry, pr = P[ry], CM = COMPANY_DATA.capexModel, sb = st.selfBuild ?? CM.selfBuild, life = st.gpuLife,
    W = f => PRICING.chips.reduce((a, c, j) => a + pr.gens[j] * TKV[f][c.tk], 0) / Math.max(pr.tot, 1e-9) / 1e3,
    so = A.ownShare[ry], eb = A.ebitdaNet[ry], opx = A.opex[ry], rent = A.rentFac[ry] + A.rentExt[ry], da = A.daExt[ry], ic = A.icExt[ry],
    mw = A.mwEff[ry], own = Math.max(0, mw - A.mwRent[ry]), rev1 = mw * W(`IF_HoldEcon`), opTK = mw * W(`IF_OpexGW`),
    cR = so * (d.years.slice(0, ry).reduce((a, y) => a + y.refresh, 0) + d.years[ry].refresh / 2), dR = cR / life,
    cIT = W(`IF_CapexIT`), cF = W(`IF_CapexFacility`), ss = (m, b) => ({ ic: m * (cIT + b * cF) * .5, da: m * (cIT / life + b * cF / A0.facLife) }),
    s5 = ss(own, sb), s6 = ss(mw, 1),
    R = [A.roic[ry], (eb - da) / ic, (rev1 - opx - rent - da) / ic, (rev1 - opTK - rent - da) / ic, (rev1 - opTK - rent - (da - dR)) / Math.max(1e-9, ic - cR), (rev1 - opTK - rent - s5.da) / Math.max(1e-9, s5.ic), (rev1 - opTK - s6.da) / s6.ic],
    lb = [`稅（稅後 → 稅前）`, `k＝1（含自研晶片係數）`, `營運成本來源（模型 → Tokenomics IF_OpexGW）`, `汰換（自投入資本與折舊移除累計汰換）`, `爬坡／閒置（投入資本與折舊改為自有在役 MW 穩態，淨額比 ½）`, `機房與租用（全部自有、自建 100%、無租金）`];
  const items = lb.map((l, i) => [l, R[i + 1] - R[i]]); items.push([`會計口徑殘差（平均淨帳面 vs ${Math.round(TKH * 100)}% 年金）`, TKH - R[6]]);
  return { ry, hurdle: TKH, model: R[0], clean: R[6], gap: TKH - R[0], items, stages: R, mw, rev1, opTK, cR, dR, ss5: s5, ss6: s6 }
}
function legBlendQ(f, o) { const S = f.segEb.reduce((a, x) => a + x[1], 0); return S > 1e-9 ? f.segEb.reduce((a, x) => a + x[1] * segMultQ(x[0], o), 0) / S : o.evEbitda }
function holdValQ(n) { return (n.holdings || []).reduce((a, h) => a + h[1] * h[2] * (1 - (h[4] ? 0 : n.holdingsDiscount ?? 0)), 0) } // MAG v0.1b：持股價值（分部加總項）
function ndAdjQ(n) { // v0.1b：淨負債調整項＝類債項目合計 − Σ 持股估值 × 持股比例 ×（1 − 持股折價）
  return (n.debtLike || []).reduce((a, x) => a + x[1], 0) - holdValQ(n) // MAG v0.1b：持股清單第 5 格＝上市（不折價）
}

function runValuation(e, t, n) {
  // v0.1b：可轉債稀釋（八檔，若轉換法）。第一輪判斷價＝現價；第二輪判斷價＝MIN(現價, 第一輪加權目標價)（「以較保守者」），
  // 評價（股數、淨負債、錨定年末淨負債）依第二輪分類。融資現金流（票息、到期還本）依 runFunding 的分類（判斷價＝發行參考價）。
  let r0 = shareCount(t, n),
    Y2 = e.years,
    CF = e.cvFund || CVN.map(() => !1),
    adj = ndAdjQ(n), // 淨負債調整項（v0.1b：類債 − 持股 ×（1 − 折價））
    mk = cv => {
      let sh = r0 + CVN.reduce((a, c, j) => a + (cv[j] ? c.S : 0), 0),
        mDebt = CVN.reduce((a, c, j) => a + (cv[j] ? 0 : c.M), 0),
        mAdj = CVN.reduce((a, c, j) => a + c.M * ((CF[j] ? 1 : 0) - (cv[j] ? 1 : 0)), 0) + adj, // 錨定年末淨負債：融資分類與評價分類不同的部分＋調整項
        i = {
          ...n,
          shares: sh,
          netDebt: n.netDebt + mDebt + adj,
          fund: {
            eq: Y2.map(e => (e.equity || 0) - (e.buyback || 0)), // MAG v0.1b：淨股權（扣瀑布回購）
            sh: Y2[Y2.length - 1].cumNewShares || 0, // MAG v0.1b：淨新股（扣回購股數；Excel「累計新股」末期）
            sh27: Y2[1].cumNewShares || 0,
            nd27: Y2[1].totalDebtEnd - Y2[1].cum + mAdj,
            shY: Y2.map(e => e.cumNewShares || 0),
            ndY: Y2.map(e => e.totalDebtEnd - e.cum + mAdj)
          }
        },
        a = forwardPL(e, i),
        _lm = i.legacyEvEbitda == null && (i.legacyEvEbitda = legBlendQ(a[Math.min(4, Math.max(1, Math.round(i.evYear ?? 1)))], i)), // MAG v0.1b：非 AI 事業加權倍數（錨定年度；Excel「非 AI 事業加權倍數」同式）
        o = dcfValue(a, i),
        c = evEbitdaLeg(a, i),
        d = Math.max(0, c),
        f = blendCall({
          spot: i.price,
          dcfPx: o.perShareT,
          pePx: d,
          eqTotal: e.totals.equity || 0,
          newSh: e.totals.newShares || 0,
          tvShare: o.tvShare,
          mktCap: i.price * r0,
          shares0: sh,
          dcfInvalid: o.invalid
        });
      return { i, a, o, c, d, f, sh, mDebt, mAdj }
    },
    cv1 = n.cvFix || cvConvQ(n.price), // cvFix：固定分類（方法區間只換倍數、不重新分類，與 Excel 相同）
    R1 = mk(cv1),
    px2 = n.cvFix ? NaN : Math.min(n.price, R1.f.blended),
    cv2 = n.cvFix || cvConvQ(px2),
    R = mk(cv2),
    { i, a, o, c, d, f } = R;
  return {
    v: i,
    fwd: a,
    d: o,
    peRaw: c,
    peAdj: d,
    hole: e.totals.equity || 0,
    fullHole: e.totals.equity || 0,
    nd27: i.fund.nd27,
    sh27: R.sh * 1.01 + i.fund.sh27,
    evK: Math.min(4, Math.max(1, Math.round(i.evYear ?? 1))),
    ndA: i.fund.ndY[Math.min(4, Math.max(1, Math.round(i.evYear ?? 1)))],
    shA: a[Math.min(4, Math.max(1, Math.round(i.evYear ?? 1)))].shares,
    call: f,
    shares: R.sh,
    shares0: r0,
    cv: { cv1, cv2, px2, tp1: R1.f.blended, sh1: R1.sh, sh2: R.sh, m1: R1.mDebt, m2: R.mDebt, cvFund: CF } // v0.1b：可轉債分類（第一輪／第二輪）
  }
}

// 依情境改寫 Accepted／Billable／FY31 預建／表外租金，其餘手動調整保留（與頁首情境按鈕相同規則；v4.1 自 segE 移入，供目標價區間共用）
function scnQ(e, sc) {
  return {
    ...e,
    scenario: sc,
    a: { ...structuredClone(e.a), newLease: [...SCENARIOS[sc].a.newLease] },
    mw31: SCENARIOS[sc].mw31,
    cvCap: SCENARIOS[sc].cvCap,
    delayMonths: SCENARIOS[sc].delay, // v0.2
    billableOpen: SCENARIOS[sc].bOpen, // v0.1b：期初可計費 MW 隨情境（以實際營收校準時）
    m: { ...e.m, accepted: [...SCENARIOS[sc].acc], billable: [...SCENARIOS[sc].bil], revMW: [...SCENARIOS[sc].rev] }
  };
}

// MAG v0.1b：容量軸（三情境 MW 路徑）× 價格軸（k 低／基準／高）3 × 3 加權目標價；其餘手動調整保留（Excel 為 rv_solve 快照，cmp31 比對）
function grid33Q(st, o) {
  if (!PRICING) return null;
  return K_AX.map(sc => [0, 1, 2].map(ax => { const s2 = { ...scnQ(st, sc), kAxis: ax }; return runValuation(runFunding(s2), s2, o).call.blended }))
}

// v4.1：目標價區間。點位＝目前輸入的加權目標價（評等仍依點位）；
// 情境區間＝保守與積極情境的加權目標價（保留手動調整）；方法區間＝目前輸入、EV/EBITDA 倍數換成 methodology.rangeMultiples 兩端。
// 判斷句與 DCF 權重說明在此產生（HTML 各頁共用；Excel 以文字公式產生同一字串，cmp31 逐字比對）。
var RANGE_MULTS = COMPANY_DATA.methodology.rangeMultiples,
  MARGIN_ALERT = RATE_TH.sellMarginAlert; // 任一情境與賣出門檻的距離小於現價的此比例（預設 5%）時，判斷句揭露該距離
function multTxt(m) {
  return Number.isInteger(m) ? m.toFixed(0) : m.toFixed(1)
}
function targetRange(d, st, o, base) {
  base = base || runValuation(d, st, o);
  let P = o.price, th = P * (1 + SELL_TH), pt = base.call.blended,
    sc = [`low`, `base`, `high`].map(k => {
      let s2 = scnQ(st, k), d2 = runFunding(s2), p2 = runValuation(d2, s2, o);
      return { sc: k, name: SCENARIOS[k].label.split(` `)[0], tgt: p2.call.blended, call: p2.call.call, gap: p2.call.blended - th, tgt2: COMPANY_DATA.capexModel && COMPANY_DATA.capexModel.mode === `tk` ? read2Q(d2, p2).tp : p2.call.blended } // MAG v0.1b r3（C16）：讀法 2
    }),
    A = [Math.min(sc[0].tgt, sc[2].tgt), Math.max(sc[0].tgt, sc[2].tgt)],
    mLo = Math.min(...RANGE_MULTS), mHi = Math.max(...RANGE_MULTS),
    bm = [mLo, mHi].map(m => runValuation(d, st, { ...o, evEbitda: m, cvFix: base.cv.cv2, legacyEvEbitda: base.v.legacyEvEbitda }).call.blended), // MAG v0.1b：方法區間只換 AI 雲端倍數（非 AI 加權倍數固定，與 Excel 相同）
    B = [Math.min(...bm), Math.max(...bm)],
    m = o.evEbitda, eq = (a, b) => Math.abs(a - b) < 1e-9,
    wd = base.call.weights.dcf, we = base.call.weights.pe,
    ev100 = base.peAdj, ev100Call = rateCall(ev100 / Math.max(P, .01) - 1, base.d.tvShare, base.call.blocked),
    raw = base.d.perShareRaw,
    notSell = sc.filter(s => s.call !== `賣出`).map(s => s.name),
    pos = eq(m, mHi) ? `目前點位位於方法區間上緣（${multTxt(mHi)}x 為穩態倍數上緣）`
      : eq(m, mLo) ? `目前點位位於方法區間下緣（${multTxt(mLo)}x）`
      : `目前點位${m > mLo && m < mHi ? `位於方法區間內` : m > mHi ? `高於方法區間上緣` : `低於方法區間下緣`}（目前倍數 ${multTxt(m)}x）`,
    judge = (A[1] < P
      ? `情境區間上緣 $${Y(A[1], 1)} 低於現價 $${Y(P, 2)}` + (notSell.length ? `，但${notSell.join(`、`)}情境未達賣出門檻 $${Y(th, 1)}` : `：賣出在三個擴張情境下都成立`)
      : `情境區間上緣 $${Y(A[1], 1)} 不低於現價 $${Y(P, 2)}：賣出並非在三個擴張情境下都成立`) + `。` +
      sc.filter(s => Math.abs(s.gap) < MARGIN_ALERT * P).map(s => `${s.name}情境 $${Y(s.tgt, 1)}，${s.gap < 0 ? `距賣出門檻` : `高於賣出門檻`} $${Y(th, 1)} 僅 $${Y(Math.abs(s.gap), 1)}。`).join(``),
    note = base.d.invalid ? `DCF 失效，已以 EV/EBITDA 100% 計。`
      : o.dcfMode !== `option` && raw < 0 ? `DCF 為 $0 仍計 ${hA(wd * 100, 0)} 權重的原因：DCF 沒有失效，而是算出的股權價值為負（每股約 ${mA(raw, 1)}），代表自由現金流折現後的企業價值不足以償還淨負債。EV/EBITDA 沒有扣除 GPU 汰換與擴張支出，保留 DCF 權重，是為了讓目標價反映「EBITDA 付不起資本成本」。若不計 DCF，目標價為 $${Y(ev100, 1)}（${ev100Call}）；${hA(wd * 100, 0)}／${hA(we * 100, 0)} 為判斷值。`
      : ``;
  return {
    pt, P, th, scen: sc, A, B, mults: [mLo, mHi], ev100, ev100Call, pos, judge, note,
    bLabel: `方法區間（目前情境，${multTxt(mLo)}–${multTxt(mHi)}x）`,
    head: `目標價 $${Y(pt, 1)}（情境區間 $${Y(A[0], 1)}–$${Y(A[1], 1)}）`
  }
}

function EM(e, t, n) {
  let r = e.fwd[1],
    i = r.services,
    a = Math.abs(r.revenue - (r.gpu + i + r.legacy)) < .05,
    o = e.fwd.every((e, n) => Math.abs(e.cashCapex - t.years[n].cashCapex) < 1e-6),
    s = e.fwd.every((e, n) => Math.abs(e.interest - t.years[n].interest) < 1e-6),
    c = !e.call.blocked || e.call.call !== `買進`,
    l = n.includeAtm || e.shares < VAL_DEFAULTS.shares - .01,
    h = e.fwd[0].revenue * 2,
    g = h + 0;
  return [{
    id: `val-cap`,
    ok: a,
    severity: a ? `ok` : `block`,
    title: `營收連動產能`,
    detail: `${PERIODS[1]} 算力產能 ${r.gpu.toFixed(1)} + 非算力服務 ${i.toFixed(1)} + 非 AI 事業 ${r.legacy.toFixed(1)} = 營收 ${r.revenue.toFixed(1)}bn。改 Billable／利用率／rev/MW 會改這列。`
  }, {
    id: `val-capex`,
    ok: o,
    severity: o ? `ok` : `block`,
    title: `DCF Cash CapEx 連動`,
    detail: `五期 Cash CapEx ${t.totals.cashCapex.toFixed(0)}bn 與損益／DCF 為同一組數字。`
  }, {
    id: `val-int`,
    ok: s,
    severity: s ? `ok` : `block`,
    title: `利息連動`,
    detail: `損益表利息（存量債務利息＋現金為負時的透支利息 ${t.totals.overdraft.toFixed(1)}bn、CDS 連動加碼）取自資金模型。`
  }, {
    id: `val-veto`,
    ok: c,
    severity: c ? e.call.funded ? `ok` : `watch` : `block`,
    title: `股權募資過大禁止買進`,
    detail: e.call.funded ? `五期缺口可全由債務支應，無需新股。` : `五期需股權 ${e.hole.toFixed(0)}bn；${e.call.blocked ? `超過現市值 ${multTxt(RATE_TH.equityRaiseMaxMult)} 倍，結論不會是買進` : `未超過現市值 ${multTxt(RATE_TH.equityRaiseMaxMult)} 倍`}。結論「${e.call.call}」。`
  }, {
    id: `val-atm`,
    ok: l,
    severity: `ok`,
    title: `ATM／可轉債與股數`,
    detail: `評價股數 ${e.shares.toFixed(4)}bn（季末流通 ${Y(LATEST_Q.sharesOut, 4)}bn＋${TXQ.sharesNote}）。${n.includeAtm ? `期後股權／可轉債淨現金計入首期來源。` : `期後股權／可轉債關閉，缺口同步擴大。`}`
  }, {
    id: `val-rev-guide`,
    ok: inRevGuideQ(HIST_PL[3].revenue + e.fwd[0].revenue, .05),
    severity: inRevGuideQ(HIST_PL[3].revenue + e.fwd[0].revenue) ? `ok` : `watch`,
    title: `${PERIODS[0]} 營收指引 ${REV_GUIDE_TXT}bn`,
    detail: `${TXQ.guideLine}。本模型年初至今實際 ${Y(HIST_PL[3].revenue, 3)} + 模型期 ${e.fwd[0].revenue.toFixed(2)} = ${(HIST_PL[3].revenue+e.fwd[0].revenue).toFixed(2)}bn；營收由 MW × 每 MW 年收入驅動，不回推指引，差距見「市場共識」分頁的差異原因。`
  }, {
    id: `val-plug`,
    ok: r.plugRatio <= .2,
    severity: r.plugRatio > .2 ? `watch` : `ok`,
    title: `IS 與資金模型同一組營收`,
    detail: `${PERIODS[1]} 算力營收 ${r.gpu.toFixed(1)}（＝平均在役 MW × 每 MW 年收入）；其中期初 RPO 可涵蓋 ${r.fundingRev.toFixed(1)}、其餘 ${r.inYear.toFixed(1)}；未售產能占容量 ${(r.plugRatio*100).toFixed(0)}%。非算力服務 ${r.services.toFixed(1)} 單列。`
  }, {
    id: `dcf-validity`,
    ok: !e.d.invalid,
    severity: e.d.invalid ? `watch` : `ok`,
    title: e.d.invalid ? `DCF 失效：${e.d.invalidReason}` : `DCF 有效（WACC − g ＝ ${((e.v.wacc - e.v.g)*100).toFixed(1)}%，常態化 FCF ${e.d.normFcf.toFixed(1)}）`,
    detail: `失效只定義為數學上的失效：WACC ≤ g，或常態化 FCF ≤ 0。股權價值為負不算失效——那是有效的經濟結論（企業價值低於淨負債）。目前下限方式：${e.v.dcfMode === `option` ? `選擇權（Merton，σ ${(e.v.sigma*100).toFixed(0)}%）` : `0 截斷`}；0 截斷 $${e.d.zeroPerShare.toFixed(1)}、選擇權 $${e.d.optPerShare.toFixed(1)}。失效時 DCF 權重歸零、EV/EBITDA 權重 100%。`
  }, {
    id: `val-double`,
    ok: !0,
    severity: `ok`,
    title: `舉債與股權的分工（期前融資瀑布）`,
    detail: `新債受債務上限約束，超出部分改以股權按現價折價募集：股權沒有利息、也不增加淨負債，但增加股數；新債的利息與本金都進入現金流與淨負債，兩者不重複計費。`
  }]
}

// v4.3：市場共識對照（一頁摘要與「市場共識」分頁共用）。共識數字只讀 company.json → meta.consensusFile（建置時併入 COMPANY_DATA.consensus）。
// 模型端用目前輸入（預設＝基準情境）；FY26 為全年口徑：營收＝1H 實際＋2H 模型，調整後 EBITDA＝1H 實際調整後 EBITDA＋2H 模型，CapEx＝1H 實際毛額＋2H 模型毛額。
// 判斷句、隱含倍數句、結論句由此產生（v4.4：Q3 句改由季度層 quarterlyView 產生）；Excel「摘要」頁以文字公式產生同一字串，cmp31 逐字比對。
var PRICE_DATE = COMPANY_DATA.meta.priceDate, // 現價收盤日（HTML 畫面文字用；Excel 同讀 meta.priceDate）
  mdQ = x => x.slice(5).split(`-`).map(Number).join(`/`), // 2026-09-21 → 9/21
  CONSENSUS = COMPANY_DATA.consensus,
  CONS_TOL = COMPANY_DATA.methodology.consensusGapTol, // 判斷句門檻：營收、EBITDA、CapEx 與共識差距（絕對值）超過此比例即視為分歧
  CONS_YEARS = PERIODS.slice(0, 3), // v0.1b（Oracle）：共識年度＝模型前三期（依公司財年；共識檔年度標籤須相同）
  CONS_JUDGE_KEYS = [[`營收`, `rev`], [`EBITDA`, `ebitda`], [`CapEx`, `capex`]];
function gapTxtQ(g, pt) { // 差距文字：比例 → +1.4%；pt＝true 時為百分點 → −2.8pt
  return `${g < 0 ? `−` : `+`}${Y(Math.abs(g) * 100, 1)}${pt ? `pt` : `%`}`
}
function annGapTxtQ(r, k) { // v4.4：年度對照的差距文字——營收、CapEx 為比例；EBITDA 為金額差／EBITDA 率百分點；淨負債為金額差；EBITDA 率為百分點
  return k === `ebitda` ? `${dTxtQ(r.dif.ebitda, `US$bn`)}／${gapTxtQ(r.gap.ebM, true)}` : k === `nd` ? dTxtQ(r.dif.nd, `US$bn`) : gapTxtQ(r.gap[k], k === `ebM`)
}
function consensusView(d, p, o, TR, st) { // d＝runFunding、p＝runValuation、o＝評價參數、TR＝targetRange、st＝目前輸入（v4.4：差異原因的數字）
  let E = CONSENSUS.annualEstimates, y = d.years, f = p.fwd,
    rows = CONS_YEARS.map((yr, i) => {
      let rev = f[i].fyRevenue,
        ebitda = i === 0 ? ACTUAL_1H.adjEbitda + f[0].ebitda : f[i].fyEbitda,
        m = { rev, ebitda, ebM: ebitda / rev, capex: i === 0 ? ACTUAL_1H.capex + y[0].gross : y[i].gross, nd: y[i].totalDebtEnd - y[i].cum },
        c = { rev: E[yr].revenue, ebitda: E[yr].ebitda, ebM: E[yr].ebitdaMargin, capex: E[yr].capex, nd: E[yr].netDebt };
      return { yr, m, c, gap: { rev: m.rev / c.rev - 1, ebitda: m.ebitda / c.ebitda - 1, ebM: m.ebM - c.ebM, capex: m.capex / c.capex - 1, nd: m.nd / c.nd - 1 },
        dif: { ebitda: m.ebitda - c.ebitda, nd: m.nd - c.nd } } // v4.4：利潤類與淨負債以金額差呈現（門檻判斷仍用比例）
    }),
    tol = pctQ(CONS_TOL),
    ex = rows.map(r => CONS_JUDGE_KEYS.filter(([, k]) => Math.abs(r.gap[k]) > CONS_TOL).map(([n, k]) => `${n} ${annGapTxtQ(r, k)}`)),
    first = ex.findIndex(x => x.length),
    Y0 = CONS_YEARS[0], YN = CONS_YEARS[CONS_YEARS.length - 1],
    judge = first < 0 ? `${Y0}–${YN} 營收、EBITDA、CapEx 與共識差距皆在 ${tol} 以內。`
      : (first > 0 ? `${first === 1 ? Y0 : `${Y0}–${CONS_YEARS[first - 1]}`} 營收、EBITDA、CapEx 與共識差距在 ${tol} 以內；` : ``) +
        `分歧始於 ${CONS_YEARS[first]}（${ex[first].join(`、`)}）` +
        ex.slice(first + 1).map((x, j) => `；${CONS_YEARS[first + 1 + j]} ${x.length ? x.join(`、`) : `回到 ${tol} 以內`}`).join(``) + `。`,
    PT = CONSENSUS.priceTarget, e28 = E[YN],
    impTgt = (PT.mean * o.shares + e28.netDebt) / e28.ebitda,
    impPx = (o.price * o.shares + e28.netDebt) / e28.ebitda,
    lgE = p.fwd[CONS_YEARS.length - 1].legacyEbitda || 0, lgM = p.v.legacyEvEbitda ?? o.legacyEvEbitda ?? 0, // v0.1b（Oracle）：分部加總——扣除非 AI 事業（模型 EBITDA × 非 AI 事業倍數）後的 AI 雲端 隱含倍數
    impTgtOci = (x => Math.abs(x) < .05 ? 0 : x)((PT.mean * o.shares + e28.netDebt - lgE * lgM) / Math.max(e28.ebitda - lgE, .01)), // MAG v0.1b：|x|<0.05 視為 0（與 Excel TEXT 一致，避免 −0.0）
    impPxOci = (x => Math.abs(x) < .05 ? 0 : x)((o.price * o.shares + e28.netDebt - lgE * lgM) / Math.max(e28.ebitda - lgE, .01)),
    mHi = Math.max(...RANGE_MULTS),
    rel = x => { let r = Math.round(x * 10) / 10; return r > mHi ? `高於` : r < mHi ? `低於` : `等於` },
    implied = `共識平均目標價 $${Y(PT.mean, 2)} 隱含 ${YN} EV/EBITDA ${Y(impTgt, 1)}x（扣除非 AI 事業 ${Y(lgM, 1)}x × 模型 EBITDA $${Y(lgE, 1)}bn 後，AI 雲端 ${Y(impTgtOci, 1)}x），AI 雲端 倍數${rel(impTgtOci)}模型方法區間上緣 ${multTxt(mHi)}x；現價 $${Y(o.price, 2)} 隱含 ${Y(impPx, 1)}x（AI 雲端 ${Y(impPxOci, 1)}x）。`,
    up = TR.pt / o.price - 1, gapTh = TR.pt - TR.th,
    head = `${p.call.call}：點位 $${Y(TR.pt, 1)}，較現價 $${Y(o.price, 2)} ${up >= 0 ? `高` : `低`} ${hA(Math.abs(up) * 100, 0)}；情境區間 $${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}，方法區間 $${Y(TR.B[0], 1)}–$${Y(TR.B[1], 1)}；點位${gapTh < 0 ? `低於` : `高於`}賣出門檻 $${Y(TR.th, 1)} 達 $${Y(Math.abs(gapTh), 1)}。`;
  // v4.4：差距超過門檻的項目附差異原因（已決定事項 2）；原因與類型讀 company.json → varianceReasons，找不到即為「未歸類」
  let rsn = rows.flatMap((r, i) => ANN_RSN_KEYS.filter(([, k]) => k === `ebitda` && PM_TOL != null ? Math.abs(r.gap.ebM) > PM_TOL : Math.abs(r.gap[k]) > CONS_TOL).map(([n, k]) => {
      let c = findRsnQ(`annual`, r.yr, k, `consensus`),
        ctx = { m: r.m, c: r.c, gap: r.gap, y: { mwNew: y[i].mwNew, accepted: d.m.accepted[i], costMW: st ? st.a.costMW[i] : NaN }, in: st || {} };
      return { yr: r.yr, key: k, name: n, gap: r.gap[k], gtxt: annGapTxtQ(r, k), type: c ? c.type : `未歸類`, text: c ? fillTokQ(c.text, ctx) : `差異原因待補` }
    })),
    rsnSum = reasonSumQ(rsn, `原因見「損益與評價 → 市場共識」。`),
    jk = d.years.reduce((a, t) => a + t.junk, 0), S0 = st || DEFAULTS, // v0.1b（Oracle）：高息債溢出＝需失去投資級才能融資的金額（一頁摘要一句）
    igLine = S0.debtCapBasis === `ebitda` || S0.debtCapBasis === `leaseAdj` ? `需失去投資級才能融資的金額：${jk > .05 ? `$${Y(jk, 1)}bn（五期高息債溢出）` : `$0`}；投資級上限＝${S0.debtCapBasis === `leaseAdj` ? `(總債務＋租賃負債) ≤ ${multTxt(S0.debtEbitdaMax)}×(EBITDA＋租金)` : `總債務 ≤ ${multTxt(S0.debtEbitdaMax)}× 當期 EBITDA`}，股權每年 ≤ 現市值 ${pctQ(S0.eqCapPct)}。` : ``;
  // v0.2：調整後槓桿一句（路徑、距上限空間；任一期超過上限時標示）
  let al = d.years.map(t => t.adjLev), amx = Math.max(...al), ami = al.indexOf(amx), cap = S0.debtEbitdaMax,
    adjLine = `調整後槓桿（(債務＋租賃負債) ÷ (EBITDA＋租金)）路徑 ${al.map(x => Y(x, 1)).join(`／`)}×；上限 ${multTxt(cap)}×，` + (amx > cap + 1e-9 ? `${PERIODS[ami]} 超過 ${Y(amx - cap, 1)}×：需股權或失去投資級。` : `最小空間 ${Y(cap - amx, 1)}×（${PERIODS[ami]}）。`);
  // v0.2：建設延誤一句（閒置資本峰值）
  let dm = S0.delayMonths ?? 0, idl = d.years.map(t => t.idleCap || 0), ipk = Math.max(...idl), ipi = idl.indexOf(ipk),
    delayLine = dm > 0 ? `建設延誤 ${multTxt(dm)} 個月（GPU 資本支出照原時程）：閒置資本（已支出、尚未產生收入）峰值 $${Y(ipk, 1)}bn（${PERIODS[ipi]} 末）。` : `建設延誤：本情境 0 個月（無閒置資本）。`;
  // MAG v0.1b：股東回饋與 FCF 句（一頁摘要；Excel「摘要」同句，cmp31 逐字比對）
  let Yd = d.years, fn = Yd.findIndex(t => t.fcf < 0), bp = Yd.reduce((a, t) => a + t.buybackPlan, 0), bcut = Yd.map((t, i) => [PERIODS[i], t.buybackCut]).filter(x => x[1] > .05),
    ndS = Yd.reduce((a, t) => a + t.newDebt + t.junk, 0), eqS = Yd.reduce((a, t) => a + t.equity, 0),
    fcfLine = `股東回饋與 FCF：FCF ${fn < 0 ? `五期皆為正` : `首次為負 ${PERIODS[fn]}（−$${Y(-Yd[fn].fcf, 1)}bn）`}；` +
      (bp <= .05 ? `無回購計畫（減少回購步驟不適用）` : bcut.length ? `回購被迫減少：${bcut.map(x => `${x[0]} $${Y(x[1], 1)}`).join(`、`)}bn` : `回購未被迫減少`) +
      `；五期新債 $${Y(ndS, 1)}bn、股權 $${Y(eqS, 1)}bn。`;
  let AQ = aiRoicQ(d, st || DEFAULTS, o), thesisLine = AQ ? `主命題（AI 資本支出有沒有賺到資金成本）：${PERIODS[AQ.ry]} 對外 AI ROIC ${Y(AQ.roic[AQ.ry] * 100, 1)}% vs WACC ${Y(AQ.wacc * 100, 1)}%（${AQ.spread[AQ.ry] < 0 ? `−` : `+`}${Y(Math.abs(AQ.spread[AQ.ry]) * 100, 1)}pt）；打平 k ${Y(AQ.breakevenK, 2)}（目前 ${Y(AQ.k, 2)}）；全 AI（含影子收入）${Y(AQ.roicSh[AQ.ry] * 100, 1)}%。` : ``; // MAG v0.1b
  // v0.1 交付前修訂：評價口徑敏感度句、終值占比規則句
  let VS = valSensQ(d, st || DEFAULTS, o), CODEq = { 買進: `買進`, 中立: `中立`, 賣出: `賣出` },
    vsLine = `評價口徑敏感度：` + VS.rows.filter((x, i) => i < 3).map(x => `${x.label.replace(`（ERP 調整）`, ``)} $${Y(x.tp, 2)}（${x.call}）`).join(`、`) + `；` + VS.rows.slice(5).map(x => `${x.label} $${Y(x.tp, 2)}（${x.call}）`).join(`、`) +
      `；評等翻轉點：${VS.flip ? `WACC 約 ${Y(VS.flip.wacc * 100, 1)}%（ERP ${Y(VS.flip.erp * 100, 2)}%）以下由「賣出」轉為「中立」` : VS.base.tp >= VS.th ? `目前已高於賣出門檻` : `WACC 5% 仍為「賣出」`}。`,
    tvq = p.d.tvShare, tvLine = `評等規則：買進須空間 ≥ +${pctQ(RATE_TH.buyUpsideMin)} 且終值占 EV < ${pctQ(RATE_TH.buyTvShareMax)}；目前終值占 EV ${hA(tvq * 100, 0)}${tvq >= RATE_TH.buyTvShareMax ? `，此規則下只可能「賣出」或「中立」` : ``}（規則是否修改待 Andy 決定）。`;
  // MAG v0.1b r3（對照表 r1 C16、C23）：兩種讀法並列＋差額；對外比例（無揭露）±20pt 目標價
  let CMq = COMPANY_DATA.capexModel, TK = CMq && CMq.mode === `tk`, R2 = TK ? read2Q(d, p) : null, S0q = st || DEFAULTS,
    read2Line = R2 ? `兩種讀法（自用 AI）：讀法 1（影子收入不進評價）$${Y(p.call.blended, 2)}；讀法 2（自用 AI 價值中性：自用 MW 以 k＝1 計影子收入並進評價）$${Y(R2.tp, 2)}；差額 ${R2.diff < 0 ? `−` : `+`}$${Y(Math.abs(R2.diff), 2)}。主值取哪一個是 Andy 的判斷（待決）。` : ``,
    XSq = TK ? S0q.extShare ?? CMq.extShare : null,
    xsTp = TK ? [XSq - .2, Math.min(1, XSq + .2)].map(x => { const s2 = { ...S0q, extShare: x }, d2 = runFunding(s2); return { x, tp: runValuation(d2, s2, o).call.blended } }) : null,
    extLine = TK ? `對外比例無揭露（目前 ${Math.round(XSq * 100)}%，[Assumed]），是最大不確定：${xsTp.map(z => `${Math.round(z.x * 100)}% → $${Y(z.tp, 2)}`).join(`、`)}（基準 $${Y(p.call.blended, 2)}）。` : ``;
  return { rows, ex, first, judge, impTgt, impPx, impTgtOci, impPxOci, lgE, mHi, implied, head, up, gapTh, rsn, rsnSum, igLine, adjLine, delayLine, junk: jk, fcfLine, thesisLine, read2Line, read2: R2, extLine, xsTp, vsLine, vsens: VS, tvLine, aiRoic: AQ }
}

// v4.4：差異原因的共用工具（年度共識對照與季度層共用）。類型固定為四種（已決定事項 2）；原因文字中的 {路徑:格式} 由模型數字帶入。
var RSN_TYPES = [`觀點`, `拆法`, `已知限制`, `未歸類`],
  PM_TOL = COMPANY_DATA.methodology.profitMarginTolPt ?? null, // v4.4 第 5 輪：利潤類附原因門檻（利潤率百分點，比例表示；0.02＝2pt）
  ANN_RSN_KEYS = [[`營收`, `rev`], [`EBITDA`, `ebitda`], [`CapEx`, `capex`], [`淨負債`, `nd`]],
  VAR_RSN = (COMPANY_DATA.varianceReasons || {}).list || [];
function cfgPathQ(s, root) { // 以「a.b.0」路徑讀值（company.json 或指定物件）
  return String(s).split(`.`).reduce((x, k) => x == null ? x : x[k], root || COMPANY_DATA)
}
function tokFmtQ(x, f) { // 原因文字的數字格式：gap＝差距、%1＝百分比 1 位、2＝數字 2 位、s2＝帶正負號 2 位；負號用 −
  if (x == null || !Number.isFinite(x)) return `不適用`;
  if (f === `gap`) return gapTxtQ(x);
  if (f[0] === `s`) return `${x < 0 ? `−` : `+`}${Y(Math.abs(x), +f.slice(1))}`;
  return f[0] === `%` ? `${x < 0 ? `−` : ``}${Y(Math.abs(x) * 100, +f.slice(1))}%` : `${x < 0 ? `−` : ``}${Y(Math.abs(x), +f)}`
}
function tokValQ(expr, ctx) { // {a-b+c}：路徑的加減（例如 c.ebitda-g.adjOpInc-q.da）；任一項缺值即 NaN
  let parts = expr.split(/(?=[+-])/), v = 0;
  for (let t of parts) { let sg = t[0] === `-` ? -1 : 1, x = cfgPathQ(t.replace(/^[+-]/, ``), ctx); if (typeof x !== `number` || !Number.isFinite(x)) return NaN; v += sg * x }
  return v
}
function fillTokQ(t, ctx) {
  return t.replace(/\{([\w.+-]+):([%s]?\w+)\}/g, (_, a, f) => tokFmtQ(tokValQ(a, ctx), f))
}
function findRsnQ(scope, period, metric, vs) {
  return VAR_RSN.find(x => x.scope === scope && (x.period === `*` || x.period === period) && x.metric === metric && (x.vs === `*` || x.vs === vs))
}
function reasonSumQ(rsn, tail) { // 例：需附原因的 6 項差距：觀點 6 項；原因見…
  if (!rsn.length) return ``;
  return `需附原因的 ${rsn.length} 項差距：${RSN_TYPES.map(t => [t, rsn.filter(x => x.type === t).length]).filter(([, n]) => n).map(([t, n]) => `${t} ${n} 項`).join(`、`)}；${tail}`
}

// v4.4：季度層（company.json → quarterly）。由年度模型拆分，只用於追蹤：不回寫任何年度數字、不影響目標價。
// 拆分：營收依 revenueSplit（anchor＝從最新一季實際線性爬升、driverAvg＝依平均在役 MW、equal＝平分）；調整後 EBITDA＝營收 × EBITDA 率，
// EBITDA 率為跨季一條直線（兩個期間時同時符合兩期年度 EBITDA；一個期間時為常數）；車隊折舊依期內平均 PP&E（與年度公式相同）；
// 調整後營業利益＝EBITDA − 折舊；CapEx 依季內新增 Accepted MW（線性內插時等於平分）。沒有 MW 驅動（driver.type≠mw）時一律平分，MW 列為「不適用」。
var QTR_CFG = COMPANY_DATA.quarterly || null;
function valTxtQ(x, u, dp) { // 數值文字：US$bn → $1.23bn（dp 位，預設 2）；MW → 1,500 MW；% → 58.8%；負號用 −
  if (x == null || !Number.isFinite(x)) return `不適用`;
  let s = x < 0 ? `−` : ``, a = Math.abs(x);
  return u === `%` ? `${s}${Y(Math.round(+(a * 1000).toPrecision(13)) / 10, 1)}%` : // v0.1b：先以 13 位有效數字四捨五入，與 Excel TEXT 的進位一致（例如 60.65%）
     u === `MW` ? `${s}${Y(a, 0)} MW` : `${s}$${Y(a, dp ?? 2)}bn`
}
function rngTxtQ(g, u) { // 指引區間文字：$3.45–3.60bn
  let n = x => `${x < 0 ? `−` : ``}${Y(Math.abs(x), u === `MW` ? 0 : 2)}`;
  return u === `MW` ? `${n(g[0])}–${n(g[1])} MW` : u === `%` ? `${n(g[0] * 100)}–${n(g[1] * 100)}%` : `$${n(g[0])}–${n(g[1])}bn`
}
function qPosQ(x, g) { // 相對指引區間的位置
  return x > g[1] ? `高於上緣` : x < g[0] ? `低於下緣` : x >= (g[0] + g[1]) / 2 ? `區間上半部` : `區間下半部`
}
function dTxtQ(g, u, dp) { // 金額差文字：+$0.10bn／−25 MW
  return g == null || !Number.isFinite(g) ? `不適用` : `${g < 0 ? `−` : `+`}${valTxtQ(Math.abs(g), u, dp)}`
}
function gTxtQ(g, kind, u, pt) { // 差距文字（v4.4 第 4 輪）：ratio＝比例（營收、CapEx）；diff＝金額差，利潤類另附利潤率百分點；pt＝百分點
  if (g == null || !Number.isFinite(g)) return `不適用`;
  if (kind === `diff`) return dTxtQ(g, u) + (pt != null && Number.isFinite(pt) ? `／${gapTxtQ(pt, true)}` : ``);
  return gapTxtQ(g, kind === `pt`)
}
function quarterlyView(d, p, st) { // d＝runFunding、p＝runValuation、st＝目前輸入
  let C = QTR_CFG;
  if (!C) return null;
  let Qs = C.quarters, n = Qs.length, f = p.fwd, y = d.years, drv = !!(C.driver && C.driver.type === `mw`),
    pers = [...new Set(Qs.map(q => q.period))], inP = P => Qs.map((q, i) => q.period === P ? i : -1).filter(i => i >= 0),
    gq = (qk, mk) => ((C.guidance || {})[qk] || {})[mk] || null,
    endAcc = Array(n).fill(null), avgBil = Array(n).fill(null), rev = Array(n), da = Array(n), capex = Array(n), idx = Qs.map((q, i) => i + 1);
  pers.forEach(P => {
    let ix = inP(P), k = ix.length, L = PERIOD_YEARS[P],
      a0 = P === pers[0] ? C.driver?.endStart : d.m.accepted[P - 1], a1 = d.m.accepted[P],
      b0 = P === pers[0] ? st.billableOpen : y[P - 1].billDelayed, b1 = y[P].billDelayed, // v0.2：計費用（延誤後）可計費 MW
      R = f[P].revenue, how = (C.revenueSplit || [])[pers.indexOf(P)] || `equal`;
    ix.forEach((i, j) => {
      if (drv) endAcc[i] = a0 + (a1 - a0) * (j + 1) / k, avgBil[i] = b0 + (b1 - b0) * (2 * j + 1) / (2 * k);
      da[i] = COMPANY_DATA.capexModel?.mode === `tk` ? y[P].daFleet / k : (y[P].ppeBeg + y[P].capexInSvc * (2 * j + 1) / (2 * k)) / st.gpuLife * L / k; // MAG v0.1b：D&A 分池時季度＝期間 D&A 平均分配
    });
    if (how === `anchor` && C.revenueAnchor) {
      let A = C.revenueAnchor.value, g = (R - k * A) / (k * (k + 1) / 2);
      ix.forEach((i, j) => rev[i] = A + (j + 1) * g)
    } else {
      let w = ix.map(i => how === `driverAvg` && drv ? avgBil[i] : 1), s = w.reduce((a, b) => a + b, 0);
      ix.forEach((i, j) => rev[i] = R * w[j] / s)
    }
    // CapEx：依季內新增 Accepted MW；capexSplit＝guidanceAnchor 時，有季度指引的季取指引中點，其餘季依新增 MW 分配期間餘數
    let adds = ix.map((i, j) => drv ? endAcc[i] - (j === 0 ? a0 : endAcc[ix[j - 1]]) : 1),
      anc = ix.map(i => C.capexSplit === `guidanceAnchor` ? gq(Qs[i].key, `capex`) : null),
      fixed = anc.reduce((s, g) => s + (g ? (g[0] + g[1]) / 2 : 0), 0),
      free = adds.map((a, j) => anc[j] ? 0 : a), sf = free.reduce((a, b) => a + b, 0), nf = anc.filter(g => !g).length;
    ix.forEach((i, j) => capex[i] = anc[j] ? (anc[j][0] + anc[j][1]) / 2 : (y[P].gross - fixed) * (sf > 0 ? free[j] / sf : 1 / nf))
  });
  let S = pers.map(P => inP(P).reduce((s, i) => s + rev[i], 0)), T = pers.map(P => inP(P).reduce((s, i) => s + rev[i] * idx[i], 0)),
    E = pers.map(P => f[P].ebitda), mg = Array(n);
  if (pers.length === 2) {
    let det = S[0] * T[1] - T[0] * S[1], a = (E[0] * T[1] - T[0] * E[1]) / det, b = (S[0] * E[1] - S[1] * E[0]) / det;
    idx.forEach((x, i) => mg[i] = a + b * x)
  } else Qs.forEach((q, i) => { let j = pers.indexOf(q.period); mg[i] = E[j] / S[j] });
  let oaQ = Qs.map((q, i) => (y[q.period].legacyOa || 0) * rev[i] / Math.max(S[pers.indexOf(q.period)], 1e-9)), // MAG v0.1b：非 AI 事業其他攤銷依季營收占比分配
    model = { revenue: rev, adjEbitda: rev.map((r, i) => r * mg[i]), ebitdaMargin: mg, adjOpInc: rev.map((r, i) => r * mg[i] - da[i] - oaQ[i]), capex, mw: endAcc },
    annual = { revenue: P => f[P].revenue, adjEbitda: P => f[P].ebitda, adjOpInc: P => f[P].opInc, capex: P => y[P].gross },
    QE = (typeof CONSENSUS !== `undefined` && CONSENSUS && CONSENSUS.quarterlyEstimates) || {},
    num = x => typeof x === `number` && Number.isFinite(x) ? x : null,
    mref = k => C.metrics.find(x => x.key === k),
    consOf = (m, key) => { let q = QE[key]; if (!m || !q || !m.consensus) return null;
      if (m.consensus === `derived`) { let a = num(q.ebitda), b = num(q.revenue); return a != null && b ? a / b : null }
      return num(q[m.consensus]) },
    actOf = (m, key) => { let a = (C.actuals || {})[key] || {}; if (!m) return null;
      if (m.key === `ebitdaMargin`) { let x = num(a.adjEbitda), r = num(a.revenue); return x != null && r ? x / r : null }
      return num(a[m.key]) },
    gmidOf = (qk, mk) => { let g = gq(qk, mk); return g ? (g[0] + g[1]) / 2 : null },
    kindOf = m => m.gap || (m.unit === `%` ? `pt` : `ratio`),
    gapOf = (x, b, kind) => x == null || b == null ? null : kind === `ratio` ? (b !== 0 ? x / b - 1 : null) : x - b,
    tol = CONS_TOL,
    exc = (g, b, kind) => g != null && (kind === `diff` ? (b ? Math.abs(g) > tol * Math.abs(b) : g !== 0) : Math.abs(g) > tol), // 金額差：|差| > 門檻 × |比較基準|
    mar = (a, b) => a != null && b != null && b !== 0 ? a / b : null, ptd = (a, b) => a != null && b != null ? a - b : null,
    VS = { consensus: `vs 共識`, guidance: `vs 指引中點`, actual: `實際 vs 模型` },
    rows = C.metrics.map(m => { let kind = kindOf(m), u = m.unit, RM = m.marginOf ? mref(m.marginOf) : null;
      return { ...m, kind, q: Qs.map((q, i) => {
      let mv = model[m.key] ? model[m.key][i] : null, cv = consOf(m, q.key), gd = gq(q.key, m.key), gmid = gd ? (gd[0] + gd[1]) / 2 : null, av = actOf(m, q.key),
        cs = m.consensusSecondary && QE[q.key] ? num(QE[q.key][m.consensusSecondary]) : null,
        gt = ((C.guidanceText || {})[q.key] || {})[m.key] || null,
        mM = RM ? mar(mv, model[RM.key][i]) : null, cM = RM ? mar(cv, consOf(RM, q.key)) : null,
        gM = RM ? mar(gmid, gmidOf(q.key, RM.key)) : null, aM = RM ? mar(av, actOf(RM, q.key)) : null,
        r = { model: mv, cons: cv, consSec: cs, guide: gd, gmid, gtext: gt, actual: av,
          mc: gapOf(mv, cv, kind), mg: gapOf(mv, gmid, kind), am: gapOf(av, mv, kind), ac: gapOf(av, cv, kind), ag: gapOf(av, gmid, kind),
          mcP: ptd(mM, cM), mgP: ptd(mM, gM), amP: ptd(aM, mM), acP: ptd(aM, cM), agP: ptd(aM, gM),
          posM: gd && mv != null ? qPosQ(mv, gd) : null, posA: gd && av != null ? qPosQ(av, gd) : null, rsn: [] },
        P = q.period, same = inP(P), pn = C.periodNames[pers.indexOf(P)],
        ctx = { q: { ...Object.fromEntries(Object.keys(model).map(k => [k, model[k][i]])), da: da[i], mwAvg: avgBil[i] }, lq: LATEST_Q,
          c: QE[q.key] || {}, g: Object.fromEntries(C.metrics.map(x => [x.key, gmidOf(q.key, x.key)])) },
        // 期間合計的比較基準：共識＝同期各季共識合計（全部季都有才算）；指引＝全年指引中點 − 已實現（periodGuidance）
        aggC = () => { let v = same.map(j => consOf(m, Qs[j].key)); return v.every(x => x != null) ? v.reduce((a, c) => a + c, 0) : null },
        aggG = () => { let pg = ((C.periodGuidance || {})[P] || {})[m.key]; return pg ? (pg.range[0] + pg.range[1]) / 2 - (pg.less || 0) : null },
        pc = x => `${Y(x * 100, 1)}%`, CN = { consensus: `共識合計`, guidance: `指引隱含` },
        // 差異原因（v4.4 第 5 輪）：有期間合計時拆解差距——水準＝(模型期間 − 比較期間) × 比較的季占比；分配＝(模型季占比 − 比較季占比) × 模型期間；
        // 兩項相加＝季度差距。分配較大→「拆法」；水準較大→「觀點」（或 varianceReasons 設定的類型與依據）。沒有期間合計時讀 varianceReasons
        push = (vs, g, gp, cq, agg) => { let c = findRsnQ(`quarter`, q.key, m.key, vs), gt2 = gTxtQ(g, kind, u, gp), B = agg ? agg() : null, AP = annual[m.key] ? annual[m.key](P) : null;
          if (B != null && AP != null && B !== 0 && AP !== 0 && cq != null && mv != null) {
            let cS = cq / B, mS = mv / AP, lv = (AP - B) * cS, al = (mS - cS) * AP;
            return r.rsn.push(Math.abs(al) > Math.abs(lv)
              ? { vs, gap: g, gtxt: gt2, type: `拆法`, text: `分配 ${dTxtQ(al, u, 3)}（${q.label} 占 ${pn} ${pc(mS)}，${CN[vs]} ${pc(cS)}）；水準 ${dTxtQ(lv, u, 3)}` }
              : { vs, gap: g, gtxt: gt2, type: c ? c.type : `觀點`, text: `水準 ${dTxtQ(lv, u, 3)}（${pn} 模型 ${valTxtQ(AP, u, 3)}，${CN[vs]} ${valTxtQ(B, u, 3)}）；分配 ${dTxtQ(al, u, 3)}${c ? `；${fillTokQ(c.text, ctx)}` : ``}` })
          }
          r.rsn.push(c ? { vs, gap: g, gtxt: gt2, type: c.type, text: fillTokQ(c.text, ctx) } : { vs, gap: g, gtxt: gt2, type: `未歸類`, text: `差異原因待補` }) },
        // 需附原因：營收、CapEx、淨負債、MW＝差距超過比較基準的門檻；利潤類（marginOf）＝利潤率差超過 methodology.profitMarginTolPt（未設定時沿用比較基準門檻）
        need = (g, b, p_) => RM && PM_TOL != null && p_ != null ? Math.abs(p_) > PM_TOL : exc(g, b, kind);
      if (kind !== `pt`) {
        if (need(r.mc, cv, r.mcP)) push(`consensus`, r.mc, r.mcP, cv, aggC);
        if (need(r.mg, gmid, r.mgP) || (gd && mv != null && (mv > gd[1] || mv < gd[0]))) push(`guidance`, r.mg, r.mgP, gmid, aggG); // 落在指引區間外也須附原因
        if (need(r.am, mv, r.amP)) { let c = findRsnQ(`quarter`, q.key, m.key, `actual`);
          if (c) r.rsn.push({ vs: `actual`, gap: r.am, gtxt: gTxtQ(r.am, kind, u, r.amP), type: c.type, text: fillTokQ(c.text, ctx) }) }
      }
      return r
    }) } }),
    fi = Math.max(0, Qs.findIndex(q => q.key === C.focus)),
    rsnLine = i => rows.flatMap(m => m.q[i].rsn.map(x => `${m.label} ${VS[x.vs]} ${x.gtxt}：${x.type}（${x.text}）`)).join(`；`),
    line = (m, i) => { let r = m.q[i], u = m.unit, k = m.kind;
      return `${m.label}：模型 ${valTxtQ(r.model, u)}` + (r.cons != null ? `｜共識 ${valTxtQ(r.cons, u)}（${gTxtQ(r.mc, k, u, r.mcP)}）` : ``) +
        (r.guide ? `｜指引 ${rngTxtQ(r.guide, u)}（中點 ${gTxtQ(r.mg, k, u, r.mgP)}，${r.posM}）` : r.gtext ? `｜指引（文字）${r.gtext.text}` : ``) +
        (r.actual != null ? `｜實際 ${valTxtQ(r.actual, u)}（較模型 ${gTxtQ(r.am, k, u, r.amP)}${r.cons != null ? `、較共識 ${gTxtQ(r.ac, k, u, r.acP)}` : ``}${r.guide ? `、${r.posA}` : ``}）` : ``) +
        (r.rsn.length ? `〔${[...new Set(r.rsn.map(x => x.type))].join(`、`)}〕` : ``) },
    key = (C.keyMetrics || []).map(k => rows.find(m => m.key === k)).filter(Boolean),
    sums = pers.map((P, j) => ({ P, name: C.periodNames[j], ...Object.fromEntries(Object.keys(annual).map(k => [k, { q: inP(P).reduce((s, i) => s + model[k][i], 0), a: annual[k](P) }])),
      da: { q: inP(P).reduce((s, i) => s + da[i], 0), a: y[P].daFleet } }));
  return { Qs, rows, model, da, avgBil, endAcc, fi, focus: Qs[fi], rsnLine, line, key, keyLines: key.map(m => line(m, fi)), sums, drv, tol }
}

// v4.3：共識資料逐筆清單（HTML「市場共識」分頁與 Excel「輸入與假設」I 區同一組列；標籤即 Excel A 欄，cmp31 以標籤逐列比對數值、擷取日期、標記）。
// vals 為數值陣列（年度列＝FY26–FY28、季度列＝Q3／Q4、區間列＝低／高）；文字值放 text。不增補、不改任何數字；資料檔沒有的欄位顯示「未列」。
function consensusItems() {
  // v0.1b：依共識檔實際有的欄位列出（換公司時欄位不同：對照目標價可為多筆、EPS 可為 GAAP 或一般口徑、公司指引項目不同）；Excel cons_items() 同一規則
  let C = CONSENSUS, it = [], num = x => typeof x === `number`,
    add = (sec, label, vals, unit, dp, m, extra) => it.push({ sec, label: `共識｜${label}`, vals, unit, dp, src: m.source, url: m.url || ``, date: m.retrieved || m.date || `未列`, tag: m.tag || ``, note: ``, text: ``, ...(extra || {}) }),
    PR = C.priceReference, PT = C.priceTarget, XCS = PT.crossCheck == null ? [] : Array.isArray(PT.crossCheck) ? PT.crossCheck : [PT.crossCheck], RA = C.ratings, AE = C.annualEstimates, AL = C.annualEstimatesAlt,
    QE = C.quarterlyEstimates, CG = C.companyGuidance || {}, IC = C.independentCrossCheck, RM = C.recentActionsMeta || {}, S1 = `價格與目標價`;
  add(S1, `現價參考（收盤）`, [PR.close], `US$`, 2, { ...PR, retrieved: PR.date }, { note: `收盤日 ${PR.date}` });
  [[`平均`, PT.mean], [`中位數`, PT.median], [`最低`, PT.low], [`最高`, PT.high]].filter(([, x]) => num(x)).forEach(([a, x]) => add(S1, `目標價｜${a}`, [x], `US$`, 2, PT));
  add(S1, `目標價｜分析師家數`, [PT.analysts], `家`, 0, PT);
  add(S1, `目標價｜共識評等`, [], ``, 0, PT, { text: PT.consensusRating });
  XCS.forEach(XC => {
    let XM = { ...XC, tag: PT.tag }, nm = XC.source.split(`（`)[0].trim();
    add(S1, `目標價｜${nm} 對照平均`, [XC.mean], `US$`, 2, XM, { note: XC.note || `` });
    num(XC.analysts) && add(S1, `目標價｜${nm} 對照家數`, [XC.analysts], `家`, 0, XM);
    XC.consensusRating && add(S1, `目標價｜${nm} 對照評等`, [], ``, 0, XM, { text: XC.consensusRating });
  });
  [[`強力買進`, RA.strongBuy], [`買進`, RA.buy], [`持有`, RA.hold], [`賣出`, RA.sell], [`強力賣出`, RA.strongSell], [`合計`, RA.total]]
    .forEach(([a, x]) => add(`評等分布（${RA.month}）`, `評等分布｜${a}`, [x], `家`, 0, RA));
  let YR = CONS_YEARS, SA = `年度共識（${YR.join(`／`)}）`; // v0.1b（Oracle）：年度＝模型前三期
  [[`營收`, `revenue`, `US$bn`, 3], [`調整後 EBITDA`, `ebitda`, `US$bn`, 3], [`EBITDA 率`, `ebitdaMargin`, `%`, 1], [`EBIT`, `ebit`, `US$bn`, 3], [`非 GAAP 營業利益`, `ebitNonGaap`, `US$bn`, 3], [`利息費用`, `interest`, `US$bn`, 3],
   [`利息支付`, `interestPaid`, `US$bn`, 3], [`淨利`, `netIncome`, `US$bn`, 3], [`GAAP EPS`, `epsGaap`, `US$`, 3], [`EPS`, `eps`, `US$`, 3], [`CapEx`, `capex`, `US$bn`, 3], [`自由現金流`, `fcf`, `US$bn`, 3], [`淨負債`, `netDebt`, `US$bn`, 3], [`每股股利`, `dps`, `US$`, 3]]
    .filter(([, k]) => YR.every(y => AE[y] && num(AE[y][k])))
    .forEach(([a, k, u, dp]) => add(SA, `年度｜${a}`, YR.map(y => AE[y][k]), u, dp, AE));
  AE.definition && add(SA, `年度｜口徑說明`, [], ``, 0, AE, { text: AE.definition });
  if (AL) {
    let A0 = AL.sp ? { ...AL.sp, note: AL.note } : AL, // v0.1b：annualEstimatesAlt 可為單一來源或 { sp, lseg }（只列 S&P；LSEG 見獨立對照）
      SB = `S&P 對照（經 StockAnalysis.com）`, Y01 = YR.slice(0, 2), Y2 = Y01.filter(y => A0[y] && num(A0[y].revenue)), F6 = A0[YR[0]] || {},
      ek = Y01.some(y => A0[y] && num(A0[y].epsAdjusted)) ? `epsAdjusted` : `eps`, en = ek === `epsAdjusted` ? `調整後 EPS` : `EPS`, Ye = Y01.filter(y => A0[y] && num(A0[y][ek]));
    Y2.length && add(SB, `S&P 對照｜營收（${Y2.join(`／`)}）`, Y2.map(y => A0[y].revenue), `US$bn`, 2, A0);
    num(F6.revenueLow) && add(SB, `S&P 對照｜${YR[0]} 營收區間（低／高）`, [F6.revenueLow, F6.revenueHigh], `US$bn`, 2, A0);
    Ye.length && add(SB, `S&P 對照｜${en}（${Ye.join(`／`)}）`, Ye.map(y => A0[y][ek]), `US$`, 2, A0);
    num(F6[ek + `Low`]) && add(SB, `S&P 對照｜${YR[0]} ${en} 區間（低／高）`, [F6[ek + `Low`], F6[ek + `High`]], `US$`, 2, A0);
    num(F6.analysts) && add(SB, `S&P 對照｜${YR[0]} 分析師家數`, [F6.analysts], `家`, 0, A0);
    A0.note && it.length && it[it.length - 1].sec === SB && (it[it.length - 1].note = A0.note);
  }
  let QK = Object.keys(QE || {}).filter(k => /^(\d{4}|FY\d{2})Q\d$/.test(k)), SQ = `季度共識（${QK.join(`／`)}）`; // v4.4：季度依共識檔實際列出的季別；v0.1b：季別鍵可為 2026Q3 或 FY27Q2
  if (QK.length) {
    [[`營收`, `revenue`], [`EBITDA`, `ebitda`], [`EBIT`, `ebit`], [`非 GAAP 營業利益`, `ebitNonGaap`], [`淨利`, `netIncome`]]
      .filter(([, k]) => QK.every(q => num(QE[q][k])))
      .forEach(([a, k]) => add(SQ, `季度｜${a}`, QK.map(q => QE[q][k]), `US$bn`, 3, QE));
    add(SQ, `季度｜說明`, [], ``, 0, QE, { text: QE.note || `` });
  }
  let SG = `公司指引（管理層預估）`, GM = { ...CG, tag: CG.tag }, G6 = CG[YR[0]] || {};
  [[`營收（低／高）`, [`revenueLow`, `revenueHigh`], `US$bn`, 2], [`營收（下限）`, [`revenueMin`], `US$bn`, 2], [`調整後營業利益（低／高）`, [`adjOpIncomeLow`, `adjOpIncomeHigh`], `US$bn`, 2],
   [`CapEx（低／高）`, [`capexLow`, `capexHigh`], `US$bn`, 2], [`淨現金 CapEx（上限）`, [`netCashCapexMax`], `US$bn`, 2], [`非 GAAP EPS`, [`epsNonGaap`], `US$`, 2],
   [`資本市場融資`, [`capitalMarketsFunding`], `US$bn`, 2], [`年底 ARR（低／高）`, [`arrYearEndLow`, `arrYearEndHigh`], `US$bn`, 2],
   [`調整後 EBITDA 率`, [`adjEbitdaMargin`], `%`, 1], [`年底合約電力`, [`contractedPowerGW`], `GW`, 1],
   [`年底已連網電力（低／高）`, [`connectedPowerGWLow`, `connectedPowerGWHigh`], `GW`, 1], [`客戶預付（下限）`, [`prepaymentsMin`], `US$bn`, 2]]
    .filter(([, ks]) => ks.every(k => num(G6[k])))
    .forEach(([a, ks, u, dp]) => add(SG, `公司指引｜${YR[0]} ${a}`, ks.map(k => G6[k]), u, dp, GM));
  CALL_FACTS.nextQRevLo != null && add(SG, `公司指引｜下一季營收（低／高）`, [CALL_FACTS.nextQRevLo, CALL_FACTS.nextQRevHi], `US$bn`, 2, GM, { note: CG.tagNote }); // 以 company.json 為準（建置時已檢查與共識檔一致）
  if (it.length && it[it.length - 1].sec === SG && !it[it.length - 1].note) it[it.length - 1].note = CG.tagNote || ``;
  if (IC) {
    let SL = `獨立對照（${IC.provider}）`, LM = { ...IC, source: `${IC.provider}（經 ${IC.source}）` },
      ICL = { FY26RevenuePreQ2: [`FY26 營收（Q2 前）`, `US$bn`, 2], [`2026Q3RevenuePreQ2`]: [`Q3 營收（Q2 前）`, `US$bn`, 2], [`2026Q2RevenuePre`]: [`Q2 營收（財報前）`, `US$bn`, 2], [`2026Q2EpsPre`]: [`Q2 EPS（財報前）`, `US$`, 2], analysts: [`家數`, `家`, 0],
        FY27Revenue: [`FY27 營收`, `US$bn`, 2], FY28Revenue: [`FY28 營收`, `US$bn`, 2], FY27EpsAdjusted: [`FY27 調整後 EPS`, `US$`, 2], FY28EpsAdjusted: [`FY28 調整後 EPS`, `US$`, 2] },
      ks = Object.keys(IC).filter(k => num(IC[k]));
    ks.forEach((k, i) => { let [a, u, dp] = ICL[k] || [k, ``, 2]; add(SL, `${IC.provider} 對照｜${a}`, [IC[k]], u, dp, LM, i === ks.length - 1 ? { note: IC.note || `` } : {}) });
  }
  (C.recentActions || []).forEach(x => add(`最新分析師動作`, `分析師動作｜${x.date} ${x.firm}`, x.target == null ? [] : [x.target], `US$`, 0, RM,
    { text: x.target == null ? `目標價未列` : ``, note: `${x.rating}${x.note ? `；${x.note}` : ``}` }));
  add(`來源與限制`, `來源獨立性`, [], ``, 0, { source: `資料檔說明`, retrieved: C.asOf }, { text: C.sourceIndependence });
  (C.notFound || []).forEach((t, i) => add(`來源與限制`, `未取得｜${i + 1}`, [], ``, 0, { source: `資料檔說明`, retrieved: C.asOf }, { text: t }));
  return it
}

// v4.3：「來源」頁引用句（同業市值讀 company.json → peers／callFacts；目標價讀共識資料檔）。Excel「來源」頁以同一規則組字。
function consSourceTxtQ() {
  let PE = COMPANY_DATA.peers, PT = CONSENSUS.priceTarget, XC = [].concat(PT.crossCheck || [])[0];
  return {
    peers: `市值（${PE.priceDate} 收盤，${PE.priceSource}）：${COMPANY_DATA.meta.ticker} 現價 $${Y(CALL_FACTS.priceLast, 2)}（${CALL_FACTS.priceDate} 收盤）、市值 ${Y(CALL_FACTS.mktCapLast, 2)}bn（${PE.priceDate}）、流通 ${Y(LATEST_Q.sharesOut * 1e3, 2)}m；` +
      `${PE.list.map(p => `${p.ticker} ${Y(p.mkt, 2)}`).join(`；`)}。淨負債取各公司最新申報：${PE.list.map(p => `${p.ticker} ${Y(p.netDebt, 2)}`).join(`、`)}。`,
    targets: `賣方目標價（${PT.source}，擷取 ${PT.retrieved}）：平均 $${Y(PT.mean, 2)}、中位數 $${Y(PT.median, 2)}、區間 $${Y(PT.low, 2)}–$${Y(PT.high, 2)}（${PT.analysts} 家）；共識評等 ${PT.consensusRating}。` +
      `對照 ${XC.source}：平均 $${Y(XC.mean, 2)}（${XC.analysts} 家，${XC.consensusRating}）。數字為賣方意見；逐筆來源與日期見「損益與評價 → 市場共識」。`
  }
}
