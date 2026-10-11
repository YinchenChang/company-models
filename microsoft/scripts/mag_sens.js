// MAG v0.1b：報告用敏感度（HTML 引擎重算；Excel 為計算事實來源，基準值已由 cmp31 核對一致）。不進成品、不寫檔。
// 用法：node scripts/mag_sens.js [--json]；每列只改一個輸入，其餘為預設（基準容量情境、基準價格軸）。
// 輸出：加權目標價、DCF 腿、EV/EBITDA 腿、錨定期對外 AI ROIC、打平 k（對外口徑；MAG v0.1b r2 C11）。
const path = require('path'); require(path.join(__dirname, '..', 'load_engine.js'))(path.join(__dirname, '..'));
const CM = COMPANY_DATA.capexModel, RY = (CM && CM.roicYear) ?? 3;
function evalQ(fs = s => s, fo = o => o) {
  const st = fs(structuredClone(DEFAULTS)), o = fo(structuredClone(VAL_DEFAULTS));
  const d = runFunding(st), p = runValuation(d, st, o), A = aiRoicQ(d, st, o);
  return { tgt: p.call.blended, dcf: p.d.perShareT, ev: p.peAdj, roic: A ? A.roic[RY] : NaN, be: A ? A.breakevenK : NaN, legM: p.v.legacyEvEbitda, wacc: o.wacc };
}
const pathQ = (st, open) => { // 期初在役 MW 改變時，MW 路徑（預設情境速度）同步平移
  const mp = COMPANY_DATA.scenarios.mwPath, k = COMPANY_DATA.defaults.scenario; let a = [];
  PERIOD_YEARS.forEach((L, i) => a.push(Math.min(mp.contracted[k][i], (i ? a[i - 1] : open) + mp.pace[k] * L)));
  st.billableOpen = open; st.m.accepted = a; st.m.billable = a.map((x, i) => Math.round(x * COMPANY_DATA.scenarios.billableRatio.ratio[i])); return st };
const line = k => st => (st.legacyBiz.lines.find(x => x.key === k) || {});
const rows = [
  ['基準', s => s],
  ['價格軸 k 低（' + kAxQ(0).toFixed(3) + '）', s => (s.kAxis = 0, s)], ['價格軸 k 高（' + kAxQ(2).toFixed(3) + '）', s => (s.kAxis = 2, s)],
  [`自研晶片係數 ${PRICING.customFactorSens ?? .7}（收入端單邊壓力測試）`, s => (s.customFactor = PRICING.customFactorSens ?? .7, s)],
  [`自研晶片係數 ${PRICING.customFactorSens ?? .7}（雙邊：收入錨與每 MW IT 資本支出同乘）`, s => (s.customFactor = PRICING.customFactorSens ?? .7, s.customCapexFactor = PRICING.customFactorSens ?? .7, s)], // v0.1 交付前修訂：機房成本不變；MW 速度不重解
  ...(([lo, hi]) => [[`期初對外 AI MW ${lo.toLocaleString()}（區間下緣）`, s => pathQ(s, lo)], [`期初對外 AI MW ${hi.toLocaleString()}（區間上緣）`, s => pathQ(s, hi)]])(
    COMPANY_DATA.mag?.mw?.externalIT?.value ? [COMPANY_DATA.mag.mw.externalIT.value.low, COMPANY_DATA.mag.mw.externalIT.value.high] : [Math.round(DEFAULTS.billableOpen * .7), Math.round(DEFAULTS.billableOpen * 1.3)]), // v0.1a 區間（mag 區段）；沒有時 ±30%
  [`對外比例 ${Math.round(((CM.extShare ?? .8) - .2) * 100)}%（−20pt，C9）`, s => (s.extShare = (CM.extShare ?? .8) - .2, s)], [`對外比例 ${Math.round(Math.min(1, (CM.extShare ?? .8) + .2) * 100)}%（＋20pt）`, s => (s.extShare = Math.min(1, (CM.extShare ?? .8) + .2), s)], // MAG v0.1b′：列名依 company.json 對外比例
  ...((COMPANY_DATA.leases.rentedCompute || []).length ? [['租用算力租金 ×2', s => (s.rentedCompute = (COMPANY_DATA.leases.rentedCompute || []).map(c => ({ ...c, annualRent: c.annualRent * 2 })), s)], ['租用算力租金 0', s => (s.rentedCompute = [], s)]] : []), // MAG v0.1b′（Microsoft C8 d）
  ['非 AI 雲端長期年增率 5%', s => (s.legacyBiz.lines.filter(x => x.kind === `cloudResidual`).forEach(x => x.gLT = .05), s)],
  ['非 AI 雲端長期年增率 12%', s => (s.legacyBiz.lines.filter(x => x.kind === `cloudResidual`).forEach(x => x.gLT = .12), s)],
  ['GPU 壽命 5 年', s => (s.gpuLife = 5, s)],
  ['廣告 EBITDA 率 40%', s => (s.legacyBiz.lines.filter(x => x.peer === `ads`).forEach(x => x.m0 = .4), s)], ['廣告 EBITDA 率 60%', s => (s.legacyBiz.lines.filter(x => x.peer === `ads`).forEach(x => x.m0 = .6), s)],
];
const orows = [
  ...[10, 15].map(m => [`AI 雲端倍數 ${m}×` + (SEGM_Q && SEGM_Q.cloud && SEGM_Q.cloud.useAi ? `（非 AI 雲端沿用）` : `（非 AI 雲端仍用同業 ${segMultQ('cloud', VAL_DEFAULTS).toFixed(1)}×）`), o => (o.evEbitda = m, o)]), // MAG v0.1b r2：C12 後非 AI 雲端不再沿用 AI 倍數
  ['流動性折價 0%', o => (o.holdingsDiscount = 0, o)], ['流動性折價 40%', o => (o.holdingsDiscount = .4, o)],
  ...COMPANY_DATA.valuation.capm.betaSens.map(b => [`β ${b}（獨立來源區間）`, o => (o.wacc = CAPM_Q(o, b).wacc, o)]), // MAG v0.1b r2（C13）
];
const out = rows.map(([n, f]) => [n, evalQ(f)]).concat(orows.map(([n, f]) => [n, evalQ(s => s, f)]));
// 非 AI 雲端單獨改倍數（AI 雲端維持 6×）：以手動覆蓋非 AI 加權倍數模擬
const base = out[0][1];
for (const m of [10, 15]) {
  const st = structuredClone(DEFAULTS), o = structuredClone(VAL_DEFAULTS), d = runFunding(st), f = forwardPL(d, { ...o, shares: o.shares });
  const k = Math.min(4, Math.max(1, Math.round(o.evYear ?? 1))), segs = f[k].segEb, S = segs.reduce((a, x) => a + x[1], 0);
  o.legacyEvEbitda = segs.reduce((a, x) => a + x[1] * (x[0] === `cloud` ? m : segMultQ(x[0], o)), 0) / S;
  const p = runValuation(d, st, o), A = aiRoicQ(d, st, o);
  out.push([`只改非 AI 雲端倍數 ${m}×（AI 雲端 6×）`, { tgt: p.call.blended, dcf: p.d.perShareT, ev: p.peAdj, roic: A.roic[RY], be: A.breakevenK, legM: o.legacyEvEbitda, wacc: o.wacc }]);
}
if (process.argv.includes('--json')) { console.log(JSON.stringify(out)); process.exit(0) }
const f1 = x => Number.isFinite(x) ? x.toFixed(2) : '—', pc = x => Number.isFinite(x) ? (x * 100).toFixed(1) + '%' : '—';
console.log(`| 敏感度（一次只改一項） | 加權目標價 | 較基準 | DCF 腿 | EV/EBITDA 腿 | ${PERIODS[RY]} 對外 AI ROIC | 打平 k |`);
console.log('|---|---|---|---|---|---|---|');
for (const [n, r] of out) console.log(`| ${n} | ${f1(r.tgt)} | ${n === '基準' ? '—' : (r.tgt - base.tgt >= 0 ? '+' : '−') + Math.abs(r.tgt - base.tgt).toFixed(2)} | ${f1(r.dcf)} | ${f1(r.ev)} | ${pc(r.roic)} | ${f1(r.be)} |`);
