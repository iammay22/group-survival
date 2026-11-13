/**
 * 分組求生記：即時排名功能（遊戲設計 may22）
 * 讓遊戲結算時可以讀取「所有完成局的分數分布」，算出「擊敗了 XX% 的分組受難者」。
 * 只回傳匯總後的分數分布（每個分數有幾局），不回傳任何單筆紀錄或玩家代號。
 *
 * 安裝方式：
 *   1. 在 Apps Script 左側「檔案」旁按 ＋ → 指令碼，命名為「排名功能」，貼上這整段。
 *   2. 回到原本的程式碼檔，找到 function doGet() { ... } 那幾行，整段刪掉（這裡已經有新的 doGet）。
 *   3. 存檔 → 部署 → 管理部署作業 → 鉛筆圖示 → 版本選「新版本」→ 部署。網址不會變。
 *   不需要重新執行 setup，你的統計版面不會被動到。
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.scores) {
    const cache = CacheService.getScriptCache();
    let json = cache.get('scoreDist');
    if (!json) {
      json = JSON.stringify(scoreDist_());
      cache.put('scoreDist', json, 300); // 每 5 分鐘重新計算一次
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('分組求生記統計接收中');
}

function scoreDist_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('原始資料');
  if (!sh || sh.getLastRow() < 2) return { n: 0, hist: [] };
  const vals = sh.getDataRange().getValues();
  const head = vals[0];
  const iEvent = head.indexOf('事件'), iScore = head.indexOf('總分'), iVer = head.indexOf('版本');
  const counts = {};
  let n = 0;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[iEvent] !== 'finish' || r[iScore] === '' || (iVer >= 0 && r[iVer] === 'test')) continue;
    const s = Math.round(Number(r[iScore]));
    if (!isFinite(s) || s < 0 || s > 100) continue;
    counts[s] = (counts[s] || 0) + 1;
    n++;
  }
  const hist = Object.keys(counts).map(function (k) { return [Number(k), counts[k]]; })
    .sort(function (a, b) { return a[0] - b[0]; });
  return { n: n, hist: hist, updated: new Date().toISOString() };
}

/** 在編輯器裡選這個函式按「執行」，可以在下方記錄看到目前的分布，用來測試。 */
function testScoreDist() {
  Logger.log(JSON.stringify(scoreDist_()));
}
