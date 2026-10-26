/**
 * AI 投資地圖 造訪計數器（Google Apps Script；綁定在 Drive「AI投資地圖_造訪紀錄」試算表）
 * 網站每次開啟、切換分頁、開公司檔案，各送一筆到這裡，寫進「紀錄」頁；「彙總」頁用公式統計。
 * 不收 IP、不用 cookie；訪客代號是瀏覽器端隨機產生、只存在讀者自己的瀏覽器。
 * 安裝：試算表 → 擴充功能 → Apps Script → 整段貼上 → 部署 → 新增部署作業 → 類型「網頁應用程式」
 *       → 執行身分「我」、存取權「任何人」→ 部署 → 授權 → 複製網址交給 Claude。
 * 權威位置：company-models/markets/invest_map/counter/Code.gs（改版改這裡，再貼回 Apps Script 並「管理部署作業 → 編輯 → 新版本」）。
 */
var TZ = 'Asia/Taipei';
var LOG = '紀錄';
var SUM = '彙總';
var HEAD = ['時間（台北）', '日期', '事件', '對象', '訪客代號', '來源網站', '裝置', '螢幕寬', '網站版本'];
var TABS = ['map', 'position', 'invest', 'profiles', 'method'];
var TAB_NAMES = ['01 供應鏈地圖', '02 產業地位', '03 可投資性', '04 公司檔案', '附 方法與來源'];

function doGet(e) { return record_(e); }
function doPost(e) { return record_(e); }

function record_(e) {
  var p = (e && e.parameter) || {};
  var ev = clip_(p.ev, 12);
  if (['view', 'tab', 'profile'].indexOf(ev) < 0) return out_('skip');   // 只收三種事件
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensure_(ss);
    var now = new Date();
    // 時間與日期存成文字（前置 '），避免試算表自動改格式；彙總頁以文字比對日期
    sh.appendRow(["'" + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
                  "'" + Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
                  ev, clip_(p.t, 40), clip_(p.vid, 40), clip_(p.ref, 120), clip_(p.dev, 10), clip_(p.w, 6), clip_(p.ver, 20)]);
  } finally {
    lock.releaseLock();
  }
  return out_('ok');
}

function clip_(s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]/g, ' ').replace(/^[=+\-@]/, "'$&").slice(0, n); }
function out_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }

/** 重建「彙總」頁公式（程式改版後在編輯器選此函式按「執行」；不動「紀錄」頁資料）。 */
function rebuildSummary() { buildSummary_(SpreadsheetApp.getActiveSpreadsheet()); }

/** 第一次收到資料時自動建立「紀錄」與「彙總」兩頁（也可在編輯器手動執行 setup）。 */
function setup() { ensure_(SpreadsheetApp.getActiveSpreadsheet()); }

function ensure_(ss) {
  var sh = ss.getSheetByName(LOG);
  if (sh) return sh;
  ss.setSpreadsheetTimeZone(TZ);
  sh = ss.getSheets()[0];
  sh.setName(LOG);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  buildSummary_(ss);
  return sh;
}

function buildSummary_(ss) {
  var s = ss.getSheetByName(SUM) || ss.insertSheet(SUM);
  s.clear();
  var L = "'" + LOG + "'!";
  // 總計
  s.getRange('A1').setValue('總計').setFontWeight('bold');
  s.getRange('A2:B5').setValues([['開啟次數', ''], ['不重複訪客', ''], ['分頁切換', ''], ['公司檔案開啟', '']]);
  s.getRange('B2').setFormula('=COUNTIF(' + L + 'C2:C,"view")');
  s.getRange('B3').setFormula('=IFERROR(ROWS(UNIQUE(FILTER(' + L + 'E2:E,' + L + 'C2:C="view"))),0)');   // 無資料時 FILTER 報錯 → 0（COUNTUNIQUE 會把錯誤當 1）
  s.getRange('B4').setFormula('=COUNTIF(' + L + 'C2:C,"tab")');
  s.getRange('B5').setFormula('=COUNTIF(' + L + 'C2:C,"profile")');
  // 最近 60 天
  s.getRange('A7:E7').setValues([['日期', '開啟次數', '不重複訪客', '分頁切換', '公司檔案開啟']]).setFontWeight('bold');
  var f = [];
  for (var i = 0; i < 60; i++) {
    var r = 8 + i, d = 'TEXT($A' + r + ',"yyyy-mm-dd")';
    f.push(['=TODAY()-' + i,
            '=COUNTIFS(' + L + 'B2:B,' + d + ',' + L + 'C2:C,"view")',
            '=IFERROR(ROWS(UNIQUE(FILTER(' + L + 'E2:E,' + L + 'B2:B=' + d + ',' + L + 'C2:C="view"))),0)',
            '=COUNTIFS(' + L + 'B2:B,' + d + ',' + L + 'C2:C,"tab")',
            '=COUNTIFS(' + L + 'B2:B,' + d + ',' + L + 'C2:C,"profile")']);
  }
  s.getRange(8, 1, 60, 5).setFormulas(f);
  s.getRange(8, 1, 60, 1).setNumberFormat('yyyy-mm-dd');
  // 分頁熱門度（開啟時的落地分頁＋切換）
  s.getRange('G1:H1').setValues([['分頁', '瀏覽次數（落地＋切換）']]).setFontWeight('bold');
  for (var j = 0; j < TABS.length; j++) {
    s.getRange(2 + j, 7).setValue(TAB_NAMES[j]);
    s.getRange(2 + j, 8).setFormula('=COUNTIF(' + L + 'D2:D,"' + TABS[j] + '")');
  }
  // 公司檔案熱門前 20
  s.getRange('G9').setValue('公司檔案開啟（前 20）').setFontWeight('bold');
  s.getRange('G10').setFormula('=IFERROR(QUERY(' + L + 'C2:D,"select D, count(C) where C = \'profile\' group by D order by count(C) desc limit 20 label D \'公司\', count(C) \'次數\'",0),"尚無資料")');
  s.setColumnWidth(1, 100); s.setColumnWidth(7, 170); s.setColumnWidth(8, 150);
}
