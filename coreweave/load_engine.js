// 在 node 中載入引擎（核對腳本用）：CO 來自 company.json（併入 meta.consensusFile 的共識資料），再依序執行 segA、segB；以間接 eval 放入全域。
const fs = require('fs'), path = require('path');
module.exports = function loadEngine(dir = __dirname) {
  const co = JSON.parse(fs.readFileSync(path.join(dir, 'company.json'), 'utf8'));
  // v4.5：期間與日期由 calendar_q.py 推算（與建置相同的單一來源），併入 cal、periods、periodYears、valuation.optT
  const cal = JSON.parse(require('child_process').execFileSync('python3', [path.join(dir, 'calendar_q.py'), dir], { encoding: 'utf8' }));
  Object.assign(co, { cal, periods: cal.periods, periodYears: cal.periodYears }); co.valuation.optT = cal.tEnd[cal.tEnd.length - 1];
  co.consensus = JSON.parse(fs.readFileSync(path.join(dir, co.meta.consensusFile), 'utf8'));  // v4.3：與 build_html_portable.py 相同，共識資料併入 COMPANY_DATA
  (0, eval)('var COMPANY_DATA = ' + JSON.stringify(co) + ';\n' +
    fs.readFileSync(path.join(dir, 'segA.js'), 'utf8') + '\nvar ' + fs.readFileSync(path.join(dir, 'segB.js'), 'utf8'));  // segB 在建置中接續模板宣告鏈，node 中補 var
};
