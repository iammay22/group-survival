/**
 * 分組求生記：遊玩統計接收程式（遊戲設計 may22）
 * 使用方式：
 *   1. 在試算表選「擴充功能 → Apps Script」，把這整段貼上取代原本內容，存檔。
 *   2. 上方函式選單選「setup」，按「執行」，依畫面授權。會自動建立「原始資料」和「統計」兩個分頁。
 *   3. 右上角「部署 → 新增部署作業」，類型選「網頁應用程式」，
 *      執行身分選「我」，存取權選「所有人」，按部署，複製「網頁應用程式網址」。
 * 只接收匿名資料：事件、隨機玩家代號、成績、分數、人數、題目、裝置類型。不收任何個人資料。
 */

const RAW = '原始資料';
const STATS = '統計';
const HEADERS = ['時間', '事件', '玩家代號', '成績', '總分', '人數', '題目', '報告進度', '組內氣氛', '專長完整度', '退組人數', '裝置', '版本'];
const EVENTS = ['visit', 'start', 'finish', 'share'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (EVENTS.indexOf(d.event) < 0) return out_('ignored');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(RAW) || setup();
    sh.appendRow([
      new Date(), d.event, clip_(d.sid, 40), clip_(d.grade, 3), num_(d.score), num_(d.size),
      clip_(d.topic, 30), num_(d.progress), num_(d.mood), num_(d.coverage), num_(d.quits),
      clip_(d.device, 10), clip_(d.version, 10)
    ]);
    return out_('ok');
  } catch (err) {
    return out_('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_('分組求生記統計接收中');
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let raw = ss.getSheetByName(RAW);
  if (!raw) {
    raw = ss.getSheets()[0];
    if (raw.getName() === STATS || raw.getLastRow() > 0) raw = ss.insertSheet(RAW);
    else raw.setName(RAW);
  }
  raw.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#efe7ff');
  raw.setFrozenRows(1);
  raw.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');

  let st = ss.getSheetByName(STATS);
  if (st) ss.deleteSheet(st);
  st = ss.insertSheet(STATS, 0);
  const R = "'" + RAW + "'!";
  const ev = R + 'B2:B', grade = R + 'D2:D', score = R + 'E2:E', size = R + 'F2:F', dev = R + 'L2:L', sid = R + 'C2:C';

  // 總覽
  st.getRange('A1').setValue('分組求生記 遊玩統計').setFontSize(16).setFontWeight('bold');
  st.getRange('A2').setValue('資料會自動更新；遊戲設計 may22').setFontColor('#655e80');
  const overview = [
    ['總覽', '數值'],
    ['瀏覽次數', '=COUNTIF(' + ev + ',"visit")'],
    ['不重複玩家', '=IFERROR(COUNTUNIQUE(FILTER(' + sid + ',' + ev + '="visit")),0)'],
    ['開始遊戲局數', '=COUNTIF(' + ev + ',"start")'],
    ['完成遊戲局數', '=COUNTIF(' + ev + ',"finish")'],
    ['完成率', '=IFERROR(B8/B7,0)'],
    ['分享圖卡次數', '=COUNTIF(' + ev + ',"share")'],
    ['平均總分', '=IFERROR(AVERAGEIF(' + ev + ',"finish",' + score + '),0)'],
    ['A+ 比例', '=IFERROR(COUNTIFS(' + ev + ',"finish",' + grade + ',"A+")/B8,0)'],
    ['每位玩家平均局數', '=IFERROR(B7/B6,0)']
  ];
  st.getRange(4, 1, overview.length, 2).setValues(overview);
  st.getRange('B9').setNumberFormat('0.0%');
  st.getRange('B11').setNumberFormat('0.0');
  st.getRange('B12').setNumberFormat('0.0%');
  st.getRange('B13').setNumberFormat('0.0');

  // 成績分布
  const grades = ['A+', 'A', 'B', 'C', 'F'];
  st.getRange('D4:F4').setValues([['成績分布', '局數', '比例']]);
  grades.forEach(function (g, i) {
    const r = 5 + i;
    st.getRange(r, 4, 1, 3).setValues([[g,
      '=COUNTIFS(' + ev + ',"finish",' + grade + ',"' + g + '")',
      '=IFERROR(E' + r + '/SUM($E$5:$E$9),0)']]);
  });
  st.getRange('F5:F9').setNumberFormat('0.0%');

  // 依人數
  st.getRange('H4:K4').setValues([['人數', '完成局數', '平均總分', 'A+ 局數']]);
  [3, 4, 5, 6].forEach(function (n, i) {
    const r = 5 + i;
    st.getRange(r, 8, 1, 4).setValues([[n + ' 人一組',
      '=COUNTIFS(' + ev + ',"finish",' + size + ',' + n + ')',
      '=IFERROR(AVERAGEIFS(' + score + ',' + ev + ',"finish",' + size + ',' + n + '),0)',
      '=COUNTIFS(' + ev + ',"finish",' + size + ',' + n + ',' + grade + ',"A+")']]);
  });
  st.getRange('J5:J8').setNumberFormat('0.0');

  // 依裝置
  st.getRange('H11:I11').setValues([['裝置', '瀏覽次數']]);
  st.getRange('H12:I13').setValues([
    ['手機平板', '=COUNTIFS(' + ev + ',"visit",' + dev + ',"手機平板")'],
    ['電腦', '=COUNTIFS(' + ev + ',"visit",' + dev + ',"電腦")']
  ]);

  // 每日瀏覽與完成
  st.getRange('A16').setValue('每日瀏覽').setFontWeight('bold');
  st.getRange('A17').setFormula('=IFERROR(QUERY(' + "'" + RAW + "'!A2:B" + ',"select toDate(A), count(B) where B = \'visit\' group by toDate(A) label toDate(A) \'日期\', count(B) \'瀏覽次數\'",0),"還沒有資料")');
  st.getRange('A18:A400').setNumberFormat('yyyy-mm-dd');

  // 題目平均分數
  st.getRange('D16').setValue('各題目平均總分').setFontWeight('bold');
  st.getRange('D17').setFormula('=IFERROR(QUERY(' + "'" + RAW + "'!B2:G" + ',"select G, count(B), avg(E) where B = \'finish\' group by G order by avg(E) desc label G \'題目\', count(B) \'完成局數\', avg(E) \'平均總分\'",0),"還沒有資料")');

  ['A4:B4', 'D4:F4', 'H4:K4', 'H11:I11'].forEach(function (a) {
    st.getRange(a).setFontWeight('bold').setBackground('#efe7ff');
  });
  st.setColumnWidths(1, 1, 150);
  st.setColumnWidth(4, 160);
  st.setColumnWidth(8, 120);

  // 圖表
  st.insertChart(st.newChart().asColumnChart()
    .addRange(st.getRange('D4:E9'))
    .setPosition(20, 8, 0, 0)
    .setOption('title', '成績分布').setOption('legend', { position: 'none' })
    .setOption('colors', ['#b9a6f2']).build());
  st.insertChart(st.newChart().asLineChart()
    .addRange(st.getRange('A17:B400'))
    .setPosition(38, 8, 0, 0)
    .setOption('title', '每日瀏覽次數').setOption('legend', { position: 'none' })
    .setOption('colors', ['#ff9f7a']).build());

  return raw;
}

function out_(s) {
  return ContentService.createTextOutput(s);
}
function clip_(v, n) {
  return v === undefined || v === null ? '' : String(v).slice(0, n);
}
function num_(v) {
  const x = Number(v);
  return isFinite(x) ? x : '';
}
