// v0.2a：每 MW 收入（Tokenomics 錨 × k）的 3 × 3 容量 × 價格矩陣與敏感度快照（HTML 引擎；與 Excel 經 cmp31 逐項一致的同一引擎）。
// 輸出 tk_sens.json（repo 根目錄）；build_xlsx.py 寫入「每MW收入_錨定」G 區（建置時快照，附過期檢查）。verify.sh 在建 Excel 前執行。
// 用法：node scripts/tk_sens.js [輸出檔]
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..'); require(path.join(dir, 'load_engine.js'))(dir);
if (PMW_REVQ !== 'tkAnchor') { fs.writeFileSync(process.argv[2] || path.join(dir, 'tk_sens.json'), JSON.stringify({ method: PMW_REVQ }) + '\n'); console.log('legacy：不產生矩陣'); process.exit(0); }
const R = tkSensQ(DEFAULTS, VAL_DEFAULTS), rd = x => Math.round(x * 1e6) / 1e6, rr = o => ({ tgt: rd(o.tgt), eq: rd(o.eq), gap: rd(o.gap) });
const out = { method: 'tkAnchor', matrix: Object.fromEntries(Object.entries(R.matrix).map(([s, v]) => [s, Object.fromEntries(Object.entries(v).map(([p, o]) => [p, rr(o)]))])),
  sens: R.sens.map(x => ({ key: x.key, label: x.label, res: Object.fromEntries(Object.entries(x.res).map(([s, o]) => [s, rr(o)])) })) };
const p = process.argv[2] || path.join(dir, 'tk_sens.json'), old = fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : '', txt = JSON.stringify(out, null, 1) + '\n';
fs.writeFileSync(p, txt);
const M = out.matrix; console.log('3×3（容量 × 價格，加權目標價）：' + ['low', 'base', 'high'].map(s => s + ' ' + ['low', 'base', 'high'].map(x => M[s][x].tgt.toFixed(2)).join('／')).join('；') + (old === txt ? '（無變動）' : '（已更新）'));
