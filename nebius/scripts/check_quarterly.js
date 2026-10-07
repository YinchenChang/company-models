// v4.4：季度層與差異原因的建置檢查（build_html_portable.py 建置時呼叫；verify.sh 另列一步）。任何一項不符即以代碼 1 結束。
// 1. 季度加總＝年度模型（三情境；營收、調整後 EBITDA、調整後營業利益、CapEx、車隊折舊；誤差 ≤ 1e-9）
// 2. company.json → quarterly.consistency：兩個路徑的值必須相同（第三格為倍數，例如 GW → MW 為 1000）
// 3. 差距超過門檻卻沒有原因（「未歸類」）：季度（模型 vs 共識／指引）與年度共識對照，三情境
// 用法：node scripts/check_quarterly.js [repo 目錄]（預設為本檔上一層）
const path = require('path');
const dir = process.argv[2] || path.join(__dirname, '..');
require(path.join(dir, 'load_engine.js'))(dir);
const bad = [], info = [];
for (const [a, b, k] of (COMPANY_DATA.quarterly?.consistency || [])) {
  const x = cfgPathQ(a), y = cfgPathQ(b) * (k ? +k : 1);
  if (typeof x !== 'number' || !Number.isFinite(y) || Math.abs(x - y) > 1e-9) bad.push(`一致性：${a}＝${x} 與 ${b}${k ? ' × ' + k : ''}＝${y} 不同`);
}
info.push(`一致性檢查 ${(COMPANY_DATA.quarterly?.consistency || []).length} 組`);
for (const SC of Object.keys(SCENARIOS)) {
  const q = structuredClone(DEFAULTS); q.scenario = SC; q.a = structuredClone(SCENARIOS[SC].a); q.mw31 = SCENARIOS[SC].mw31; q.cvCap = SCENARIOS[SC].cvCap;
  q.m.accepted = [...SCENARIOS[SC].acc]; q.m.billable = [...SCENARIOS[SC].bil]; q.m.revMW = [...SCENARIOS[SC].rev];
  const d = runFunding(q), p = runValuation(d, q, VAL_DEFAULTS), qv = quarterlyView(d, p, q);
  const cv = consensusView(d, p, VAL_DEFAULTS, targetRange(d, q, VAL_DEFAULTS, p), q);
  if (qv) {
    let n = 0;
    for (const s of qv.sums) for (const k of ['revenue', 'adjEbitda', 'adjOpInc', 'capex', 'da']) {
      n++; if (!(Math.abs(s[k].q - s[k].a) <= 1e-9)) bad.push(`${SC} ${s.name} ${k}：季度加總 ${s[k].q} ≠ 年度 ${s[k].a}`);
    }
    info.push(`${SC}：季度加總＝年度 ${n} 項`);
    qv.rows.forEach(m => m.q.forEach((r, i) => r.rsn.filter(x => x.type === '未歸類').forEach(x =>
      bad.push(`${SC} ${qv.Qs[i].key} ${m.label} ${x.vs} ${gapTxtQ(x.gap)}：差距超過門檻卻沒有原因（company.json → varianceReasons）`))));
  }
  cv.rsn.filter(x => x.type === '未歸類').forEach(x => bad.push(`${SC} ${x.yr} ${x.name} ${gapTxtQ(x.gap)}：年度共識差距超過門檻卻沒有原因`));
  info.push(`${SC}：年度差異原因 ${cv.rsn.length} 項、季度差異原因 ${qv ? qv.rows.reduce((s, m) => s + m.q.reduce((t, r) => t + r.rsn.length, 0), 0) : 0} 項`);
}
console.log(info.join('\n'));
if (bad.length) { console.log('check_quarterly：失敗 ' + bad.length + ' 項'); bad.forEach(x => console.log('  ' + x)); process.exit(1); }
console.log('check_quarterly：全部通過');
