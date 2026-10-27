// MAG v0.2.1（Alphabet 專用）：Alphabet 層級費用線（legacyBiz 的 corp，kind＝explicit）EBITDA ＝ −corpShareOfRevenue ×（非 AI 各線營收＋AI 雲端營收）
// 的逐期值存在 company.json（explicit 線不隨營收即時重算）。每當 AI 雲端營收改變（k、Tokenomics、MW 速度），須重算該線；v0.2 的 4.83% 即此關係。
// 用法：node scripts/calib_corp.js [--write]；與 company.json 不一致時以代碼 1 結束（--write 則寫入）。以預設情境的 HTML 引擎計算（與 Excel 三情境共用同一條 explicit 線）。
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..'), P = path.join(dir, 'company.json');
require(path.join(dir, 'load_engine.js'))(dir);
const d = runFunding(structuredClone(DEFAULTS)), sh = COMPANY_DATA.defaults.legacyBiz.split.corpShareOfRevenue;
const want = d.lg.rev.map((v, i) => +(-sh * (v + d.years[i].isRev)).toFixed(3));
const corp = COMPANY_DATA.defaults.legacyBiz.lines.find(x => x.key === 'corp'), cur = corp.ebitda;
const ok = want.every((v, i) => Math.abs(v - cur[i]) < 5e-4);
console.log(`層級費用線 EBITDA：company.json ${cur.join(' / ')}；重算 ${want.join(' / ')}（占比 ${sh}）`);
if (ok) { console.log('一致'); process.exit(0); }
if (!process.argv.includes('--write')) { console.error('層級費用線與營收不一致：執行 node scripts/calib_corp.js --write'); process.exit(1); }
let s = fs.readFileSync(P, 'utf8');
const re = /("key": "corp",[\s\S]*?"ebitda": \[)([\s\S]*?)(\])/;
if (!re.test(s)) { console.error('company.json 找不到 corp 線'); process.exit(2); }
s = s.replace(re, (m, a, b, c) => a + '\n      ' + want.join(',\n      ') + '\n     ' + c);
fs.writeFileSync(P, s); console.log('已寫入 company.json → legacyBiz corp ebitda');
