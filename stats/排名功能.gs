/**
 * 分組求生記：接收資料＋即時排名＋最新戰報（遊戲設計 may22）
 *
 * 這個檔案現在包含 doPost（接收遊戲資料）和 doGet（排名與最新戰報）。
 * 安裝方式：
 *   1. 把「排名功能」檔案的內容整段換成這份。
 *   2. 回到原本的程式碼檔（Code.gs / 程式碼.gs），找到 function doPost(e) { ... } 整段刪掉，
 *      因為這裡已經有新的 doPost。（doGet 上次已經刪過了）
 *   3. 存檔 → 部署 → 管理部署作業 → 鉛筆圖示 → 版本選「新版本」→ 部署。網址不會變。
 *   不需要執行 setup，統計版面不會被動到。
 *
 * 對外只回傳匯總分數分布和最近 5 局的：時間、等級、分數、稱號、人數、題目。
 * 不回傳玩家代號、裝置、來源等任何其他資料。
 */

const GS_RAW = '原始資料';
const GS_EVENTS = ['visit', 'start', 'finish', 'share'];
const GS_RECENT = 5;

/* ---------- 接收遊戲資料：依欄位名稱寫入，欄位順序可以自由調整 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (GS_EVENTS.indexOf(d.event) < 0) return gsText_('ignored');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GS_RAW);
    if (!sh) return gsText_('no sheet');
    const lastCol = Math.max(1, sh.getLastColumn());
    const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    if (head.indexOf('稱號') < 0) { // 第一次收到稱號時自動補上欄位標題
      sh.getRange(1, lastCol + 1).setValue('稱號').setFontWeight('bold');
      head.push('稱號');
    }
    const v = {
      '時間': new Date(), '事件': d.event, '玩家代號': gsClip_(d.sid, 40), '成績': gsClip_(d.grade, 3),
      '總分': gsNum_(d.score), '人數': gsNum_(d.size), '題目': gsClip_(d.topic, 30),
      '報告進度': gsNum_(d.progress), '組內氣氛': gsNum_(d.mood), '專長完整度': gsNum_(d.coverage),
      '退組人數': gsNum_(d.quits), '裝置': gsClip_(d.device, 10), '版本': gsClip_(d.version, 10),
      '來源': gsClip_(d.source, 24), '來源網域': gsClip_(d.ref, 60), '稱號': gsClip_(d.title, 16)
    };
    sh.appendRow(head.map(function (h) { return h in v ? v[h] : ''; }));
    if (d.event === 'finish') CacheService.getScriptCache().remove('recent');
    return gsText_('ok');
  } catch (err) {
    return gsText_('error');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 排名分布＋最新戰報 ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.scores) {
    const cache = CacheService.getScriptCache();
    let dist = cache.get('scoreDist'), recent = cache.get('recent');
    if (!dist || !recent) {
      const data = gsScan_();
      if (!dist) { dist = JSON.stringify(data.dist); cache.put('scoreDist', dist, 300); }
      if (!recent) { recent = JSON.stringify(data.recent); cache.put('recent', recent, 60); }
    }
    const out = JSON.parse(dist);
    out.recent = JSON.parse(recent);
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  return gsText_('分組求生記統計接收中');
}

function gsScan_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GS_RAW);
  if (!sh || sh.getLastRow() < 2) return { dist: { n: 0, hist: [] }, recent: [] };
  const vals = sh.getDataRange().getValues();
  const head = vals[0], col = function (name) { return head.indexOf(name); };
  const iT = col('時間'), iE = col('事件'), iG = col('成績'), iS = col('總分'), iN = col('人數'),
    iP = col('題目'), iV = col('版本'), iTi = col('稱號');
  const counts = {}, recent = [];
  let n = 0;
  for (let i = vals.length - 1; i >= 1; i--) {
    const r = vals[i];
    if (r[iE] !== 'finish' || r[iS] === '' || (iV >= 0 && r[iV] === 'test')) continue;
    const s = Math.round(Number(r[iS]));
    if (!isFinite(s) || s < 0 || s > 100) continue;
    counts[s] = (counts[s] || 0) + 1;
    n++;
    if (recent.length < GS_RECENT) {
      const t = r[iT] instanceof Date ? r[iT].getTime() : Date.parse(r[iT]);
      recent.push({
        t: isFinite(t) ? t : null, grade: String(r[iG] || ''), score: s,
        size: Number(r[iN]) || null, topic: String(r[iP] || ''), title: iTi >= 0 ? String(r[iTi] || '') : ''
      });
    }
  }
  const hist = Object.keys(counts).map(function (k) { return [Number(k), counts[k]]; })
    .sort(function (a, b) { return a[0] - b[0]; });
  return { dist: { n: n, hist: hist, updated: new Date().toISOString() }, recent: recent };
}

/** 測試用：選這個函式按「執行」，在下方記錄看結果。 */
function testBoard() {
  Logger.log(JSON.stringify(gsScan_()));
}

function gsText_(s) { return ContentService.createTextOutput(s); }
function gsClip_(v, n) { return v === undefined || v === null ? '' : String(v).slice(0, n); }
function gsNum_(v) { const x = Number(v); return v === '' || v === null || v === undefined || !isFinite(x) ? '' : x; }
