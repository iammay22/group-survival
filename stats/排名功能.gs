/**
 * 分組求生記：接收資料＋即時排名＋最新戰報（遊戲設計 may22）・v9
 *
 * 以 may22 目前版本為基礎合併：保留 ?action=scores、?nocache=1、avg、錯誤處理、標題去空白；
 * 新增：doPost（依欄位名稱寫入，自動補「稱號」欄）、最近 5 局戰報 recent。
 *
 * 安裝方式：
 *   1. 「排名功能」檔整段換成這份。
 *   2. 原本的程式碼檔（Code.gs / 程式碼.gs）裡，把 function doPost(e) { ... } 整段刪掉。
 *      （doGet 已經在這個檔案，原檔若還有 doGet 也一起刪掉。）
 *   3. 存檔 → 部署 → 管理部署作業 → 鉛筆圖示 → 版本選「新版本」→ 部署（網址不會變）。
 *   不需要執行 setup，統計版面不會被動到。
 *
 * 對外只回傳：分數分布、平均分、最近 5 局的時間／等級／分數／稱號／人數／題目。
 * 不回傳玩家代號、裝置、來源等任何其他資料。
 */

const GS_RAW = '原始資料';
const GS_EVENTS = ['visit', 'start', 'finish', 'share'];
const GS_RECENT = 5;
const GS_KEY_DIST = 'scoreDist_v1';   // 快取 5 分鐘
const GS_KEY_RECENT = 'recent_v1';    // 快取 1 分鐘，有新的完成局時立刻清掉

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
    const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
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
    if (d.event === 'finish') CacheService.getScriptCache().remove(GS_KEY_RECENT);
    return gsText_('ok');
  } catch (err) {
    return gsText_('error');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 排名分布＋最新戰報：?scores=1 或 ?action=scores，?nocache=1 強制重算 ---------- */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.scores || params.action === 'scores') {
    try {
      const cache = CacheService.getScriptCache();
      let dist = params.nocache ? null : cache.get(GS_KEY_DIST);
      let recent = params.nocache ? null : cache.get(GS_KEY_RECENT);
      if (!dist || !recent) {
        const data = scoreDist_();
        if (!dist) { dist = JSON.stringify(data.dist); cache.put(GS_KEY_DIST, dist, 300); }
        if (!recent) { recent = JSON.stringify(data.recent); cache.put(GS_KEY_RECENT, recent, 60); }
      }
      const out = JSON.parse(dist);
      out.recent = JSON.parse(recent);
      return gsJson_(out);
    } catch (err) {
      return gsJson_({ n: 0, avg: 0, hist: [], recent: [], error: String(err) });
    }
  }
  return gsText_('分組求生記統計接收中');
}

function scoreDist_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GS_RAW);
  const nowIso = new Date().toISOString();
  const empty = { dist: { n: 0, avg: 0, hist: [], updated: nowIso }, recent: [] };
  if (!sh || sh.getLastRow() < 2) return empty;

  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(function (h) { return String(h).trim(); });
  const iT = head.indexOf('時間'), iE = head.indexOf('事件'), iG = head.indexOf('成績'),
    iS = head.indexOf('總分'), iN = head.indexOf('人數'), iP = head.indexOf('題目'),
    iV = head.indexOf('版本'), iTi = head.indexOf('稱號');
  if (iE < 0 || iS < 0) return empty;

  const counts = {}, recent = [];
  let n = 0, sum = 0;
  for (let i = vals.length - 1; i >= 1; i--) { // 由新到舊掃，順便抓最近 5 局
    const r = vals[i];
    if (String(r[iE]).trim() !== 'finish') continue;
    if (r[iS] === '' || r[iS] === null) continue;
    if (iV >= 0 && String(r[iV]).trim() === 'test') continue;
    const s = Math.round(Number(r[iS]));
    if (!isFinite(s) || s < 0 || s > 100) continue;
    counts[s] = (counts[s] || 0) + 1;
    sum += s;
    n++;
    if (recent.length < GS_RECENT) {
      const t = r[iT] instanceof Date ? r[iT].getTime() : Date.parse(r[iT]);
      recent.push({
        t: isFinite(t) ? t : null,
        grade: iG >= 0 ? String(r[iG] || '').trim() : '',
        score: s,
        size: iN >= 0 ? (Number(r[iN]) || null) : null,
        topic: iP >= 0 ? String(r[iP] || '').trim() : '',
        title: iTi >= 0 ? String(r[iTi] || '').trim() : ''
      });
    }
  }
  const hist = Object.keys(counts)
    .map(function (k) { return [Number(k), counts[k]]; })
    .sort(function (a, b) { return a[0] - b[0]; });
  const avg = n > 0 ? Math.round((sum / n) * 10) / 10 : 0;
  return { dist: { n: n, avg: avg, hist: hist, updated: nowIso }, recent: recent };
}

/** 測試：選 testScoreDist 按「執行」，在執行記錄看樣本數與最近 5 局。 */
function testScoreDist() {
  Logger.log(JSON.stringify(scoreDist_()));
}

function gsText_(s) { return ContentService.createTextOutput(s); }
function gsJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function gsClip_(v, n) { return v === undefined || v === null ? '' : String(v).slice(0, n); }
function gsNum_(v) { const x = Number(v); return v === '' || v === null || v === undefined || !isFinite(x) ? '' : x; }
