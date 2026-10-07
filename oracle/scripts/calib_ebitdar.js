// v0.1c（Oracle）：EBITDAR 率校準（defaults.ebitdarAdj）——基準情境起點（首期）與穩態（末期）的租金 ÷ OCI 營收。
// EBITDAR 率＝EBITDA 率（ebStart／ebSteady）＋此比例，使基準情境起點與穩態 EBITDA 率維持 ebStart／ebSteady；保守與積極用同一 EBITDAR 率，承擔固定租金。
// 用法：node scripts/calib_ebitdar.js [--write]；未加 --write 時與 company.json 不一致（差 > 1e-6）即以代碼 1 結束（verify.sh 用）。
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..');
require(path.join(dir, 'load_engine.js'))(dir);
if (DEFAULTS.ebitdaBasis !== 'ebitdar') { console.log('ebitdaBasis 不是 ebitdar：不適用'); process.exit(0); }
const SC = 'base', q = structuredClone(DEFAULTS);
q.scenario = SC; q.a = structuredClone(SCENARIOS[SC].a); q.mw31 = SCENARIOS[SC].mw31; q.cvCap = SCENARIOS[SC].cvCap; q.billableOpen = SCENARIOS[SC].bOpen;
q.m.accepted = [...SCENARIOS[SC].acc]; q.m.billable = [...SCENARIOS[SC].bil]; q.m.revMW = [...SCENARIOS[SC].rev];
const y = runFunding(q).years, n = y.length - 1;
const adj = [y[0].lease / y[0].totRev, y[n].lease / y[n].totRev].map(x => +x.toFixed(6));
const cur = DEFAULTS.ebitdarAdj || [];
const ebm = [y[0].ebM, y[n].ebM];
console.log(`基準情境 租金÷OCI 營收：起點 ${(adj[0]*100).toFixed(4)}%、穩態 ${(adj[1]*100).toFixed(4)}%（company.json：${cur.map(x => (x*100).toFixed(4) + '%').join('／')}）；基準 EBITDA 率 起點 ${(ebm[0]*100).toFixed(4)}%、穩態 ${(ebm[1]*100).toFixed(4)}%`);
const same = cur.length === 2 && adj.every((x, i) => Math.abs(x - cur[i]) <= 1e-6);
if (process.argv.includes('--write')) {
  if (same) { console.log('無變動'); process.exit(0); }
  const p = path.join(dir, 'company.json'), s = fs.readFileSync(p, 'utf8');
  const re = /("ebitdarAdj": \[)[^\]]*(\])/;
  if (!re.test(s)) { console.error('company.json 找不到 defaults.ebitdarAdj'); process.exit(2); }
  fs.writeFileSync(p, s.replace(re, `$1\n   ${adj[0]},\n   ${adj[1]}\n  $2`));
  console.log('已寫入 company.json → defaults.ebitdarAdj'); process.exit(0);
}
if (!same) { console.error('defaults.ebitdarAdj 與基準情境校準不一致：執行 node scripts/calib_ebitdar.js --write'); process.exit(1); }
if (Math.abs(ebm[0] - DEFAULTS.ebStart) > 1e-5 || Math.abs(ebm[1] - DEFAULTS.ebSteady) > 1e-5) { console.error('基準情境 EBITDA 率未回到 ebStart／ebSteady'); process.exit(1); }
console.log('一致');
