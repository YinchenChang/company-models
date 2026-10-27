// MAG v0.1b r2（對照表 r1 第 8 節 C10）：基準情境 MW 路徑對齊資本支出——解 scenarios.mwPath.pace[情境]，
// 使首期全年資本支出公式值（AI 成長型＋汰換＋非 AI）＝全年指引中點（callFacts.capexLo／capexHi），即「對帳：指引隱含 AI 建置 MW vs MW 路徑」落差＝0。
// 非 AI 雲端資本支出隨 AI 營收而變（cloudResidual），所以用二分法在整個引擎上解（不是封閉公式）。這是把支出換算成容量（實體換算），不是從營收倒推價格。
// 同時寫入 scenarios.mw31[情境]（模型期後次期新增＝同一速度）與 defaults.mw31（預設情境相同時）。
// 設定：company.json → scenarios.mwPath.calibrate＝{ "scenario": "base", "decimals": 1 }；無此設定＝不適用。
// 用法：node scripts/calib_pace.js [--write]；未加 --write 時與 company.json 不一致即以代碼 1 結束（verify.sh 用）。
const fs = require('fs'), path = require('path'), vm = require('vm'), cp = require('child_process');
const dir = path.join(__dirname, '..'), P = path.join(dir, 'company.json');
const raw = fs.readFileSync(P, 'utf8'), co0 = JSON.parse(raw);
const cal = co0.scenarios.mwPath.calibrate;
if (!cal) { console.log('scenarios.mwPath.calibrate 未設定：不適用'); process.exit(0); }
const cf = co0.callFacts || {};
if (cf.capexLo == null || cf.capexHi == null) { console.log('callFacts.capexLo／capexHi 空白：不適用'); process.exit(0); }
const SC = cal.scenario || 'base', DEC = cal.decimals ?? 1, G = (cf.capexLo + cf.capexHi) / 2;
const py = a => JSON.parse(cp.execFileSync('python3', [path.join(dir, 'calendar_q.py'), dir, ...a], { encoding: 'utf8' }));
const calq = py([]), tk = py(['--tk']), cons = py(['--consensus']);
const src = fs.readFileSync(path.join(dir, 'segA.js'), 'utf8') + '\nvar ' + fs.readFileSync(path.join(dir, 'segB.js'), 'utf8');
// 以指定速度跑一次引擎（新 context，常數全部重算），回傳首期全年資本支出公式值與對帳列
function run(p) {
  const co = JSON.parse(raw);
  Object.assign(co, { cal: calq, periods: calq.periods, periodYears: calq.periodYears, tk, consensus: cons }); co.valuation.optT = calq.tEnd[calq.tEnd.length - 1];
  co.scenarios.mwPath.pace[SC] = p; co.scenarios.mw31[SC] = p; if (co.defaults.scenario === SC) co.defaults.mw31 = p;
  const ctx = vm.createContext({ console, Math, JSON, Object, Array, Number, String, isFinite, structuredClone });
  vm.runInContext('var COMPANY_DATA = ' + JSON.stringify(co) + ';\n' + src +
    `\n;var __q = structuredClone(DEFAULTS); __q.scenario = ${JSON.stringify(SC)}; Object.assign(__q, { a: structuredClone(SCENARIOS.${SC}.a), mw31: SCENARIOS.${SC}.mw31, cvCap: SCENARIOS.${SC}.cvCap, billableOpen: SCENARIOS.${SC}.bOpen }); __q.m.accepted = [...SCENARIOS.${SC}.acc]; __q.m.billable = [...SCENARIOS.${SC}.bil]; __q.m.revMW = [...SCENARIOS.${SC}.rev];` +
    `\nvar __r = runFunding(__q), __c = __r.checks.find(c => c.id === 'capex-recon');`, ctx);
  return { fy: ctx.__r.years[0].capexFormulaFY, imp: ctx.__c && ctx.__c.imp, pth: ctx.__c && ctx.__c.pth };
}
let lo = 0, hi = Math.max(1, co0.scenarios.mwPath.pace[SC]) * 8, flo = run(lo).fy - G, fhi = run(hi).fy - G;
if (flo > 0 || fhi < 0) { console.error(`無法夾住解：速度 ${lo} → ${(flo + G).toFixed(1)}、${hi} → ${(fhi + G).toFixed(1)}（指引 ${G}）`); process.exit(2); }
for (let i = 0; i < 60 && hi - lo > 1e-6; i++) { const m = (lo + hi) / 2, f = run(m).fy - G; if (f < 0) lo = m; else hi = m; }
const sol = +((lo + hi) / 2).toFixed(DEC), r = run(sol), cur = co0.scenarios.mwPath.pace[SC];
console.log(`${SC} 併網速度解＝${sol} MW／年（company.json：${cur}）；首期全年資本支出公式值 ${r.fy.toFixed(3)} vs 指引中點 ${G}；對帳 隱含 ${r.imp?.toFixed(1)} vs 路徑 ${r.pth?.toFixed(1)} MW（落差 ${r.imp && r.pth ? ((r.imp / r.pth - 1) * 100).toFixed(2) : '—'}%）`);
const same = Math.abs(sol - cur) < Math.pow(10, -DEC) / 2 + 1e-9 && co0.scenarios.mw31[SC] === cur && (co0.defaults.scenario !== SC || co0.defaults.mw31 === cur);
if (process.argv.includes('--write')) {
  if (same) { console.log('無變動'); process.exit(0); }
  let s = raw;
  const sub = (re, f) => { if (!re.test(s)) { console.error('company.json 找不到 ' + re); process.exit(2); } s = s.replace(re, f); };
  sub(new RegExp(`("pace": \\{[^}]*"${SC}": )[-0-9.]+`), `$1${sol}`);
  sub(new RegExp(`("mw31": \\{[^}]*"${SC}": )[-0-9.]+`), `$1${sol}`);
  if (co0.defaults.scenario === SC) sub(/("mw31": )[-0-9.]+(,\s*"capexFloorFY0")/, `$1${sol}$2`);
  fs.writeFileSync(P, s); console.log(`已寫入 company.json → scenarios.mwPath.pace.${SC}、scenarios.mw31.${SC}${co0.defaults.scenario === SC ? '、defaults.mw31' : ''}`); process.exit(0);
}
if (!same) { console.error('併網速度與資本支出指引校準不一致：執行 node scripts/calib_pace.js --write'); process.exit(1); }
console.log('一致');
