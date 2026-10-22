// WhiteFiber v0.2：三情境（價格基準）關鍵數字摘要（HTML 引擎；報告與交接檔取數用）。用法：node scripts/summary_v02.js [輸出.json]
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..'); require(path.join(dir, 'load_engine.js'))(dir);
const out = {};
for (const sc of ['low', 'base', 'high']) {
  const s = scnQ(DEFAULTS, sc), f = runFunding(structuredClone(s)), a = runValuation(f, structuredClone(s), structuredClone(VAL_DEFAULTS)), Y = f.years, T = f.totals;
  out[sc] = { tgt: a.call.blended, call: a.call.call, upside: a.call.upside, pf: Y.reduce((x, y) => x + y.pfDraw, 0), newDebt: T.newDebt, conv: T.convNew, equity: T.equity, junk: T.junk, newShares: T.newShares, preFinGap: Math.max(0, -T.preFinEnd),
    rev: Y.map(y => y.fyRevenue ?? y.totRev), ebitda: Y.map(y => y.ebitdaPL), totalDebt: Y.map(y => y.totalDebtEnd), lev: Y.map(y => y.adjLev), levMax: Math.max(...Y.map(y => y.adjLev)), mwEnd: Y.map(y => y.capacity !== undefined ? y.avgBillable : null), avgBill: Y.map(y => y.avgBillable), revMW: s.m.revMW, billableOpen: s.billableOpen };
}
const p = process.argv[2] || path.join(dir, 'out', 'summary_v02.json'); fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, JSON.stringify(out, null, 1));
for (const [k, o] of Object.entries(out)) console.log(k, o.tgt.toFixed(2), o.call, 'pf', o.pf.toFixed(1), 'nd', o.newDebt.toFixed(1), 'cv', o.conv.toFixed(1), 'eq', o.equity.toFixed(1), 'junk', o.junk.toFixed(1), 'sh', o.newShares.toFixed(2), 'rev', o.rev.map(x => x.toFixed(0)).join('/'), 'debtFY30', o.totalDebt[4].toFixed(0), 'levMax', o.levMax.toFixed(1));
