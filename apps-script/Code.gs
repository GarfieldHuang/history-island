// 宋元明清闖關島 — 排行榜後端（Google Apps Script）
// 用法：在一份新的 Google 試算表 → 擴充功能 → Apps Script，把這整份貼上，
// 部署 → 新增部署作業 → 類型「網頁應用程式」→ 執行身分「我」→ 存取權「所有人」→ 部署，
// 複製結尾是 /exec 的網址，填進 index.html 的 DATA.meta.leaderboard。

const SHEET_NAME = '排行榜';
const MAX_ISLANDS = 10;   // 島數上限（防亂填）
const MAX_STARS = 200;    // 星星上限（防亂填；目前題目 72 題）
const TOP_N = 50;

// 綁定在試算表上（擴充功能 → Apps Script）就用那份；
// 單獨建立的腳本則第一次自動建一份試算表，id 存在指令碼屬性裡。
function spreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('宋元明清闖關島排行榜');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// 在編輯器裡執行這個函式，就會在「執行記錄」印出試算表網址
function showSheetUrl() {
  Logger.log(spreadsheet_().getUrl());
}

function sheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const blank = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0] : null;
    sh = blank ? blank.setName(SHEET_NAME) : ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', '名字', '過關島數', '星星', '答對率', '全部完成', '更新時間']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clampInt_(v, lo, hi) {
  const n = Math.round(Number(v) || 0);
  return Math.max(lo, Math.min(hi, n));
}

// 名字：去掉試算表公式開頭、控制字元，最多 10 個字
function cleanName_(s) {
  let name = String(s || '').replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, 10);
  if (/^[=+\-@]/.test(name)) name = "'" + name;
  return name || '玩家';
}

function doGet() {
  const values = sheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
  values.sort((a, b) =>
    (b[2] - a[2]) || (b[3] - a[3]) || (b[4] - a[4]) ||
    (new Date(a[6]).getTime() - new Date(b[6]).getTime()));
  const out = values.slice(0, TOP_N).map(r => ({
    id: String(r[0]).slice(0, 6),
    name: String(r[1]).replace(/^'/, ''),
    islands: Number(r[2]) || 0,
    stars: Number(r[3]) || 0,
    accuracy: Number(r[4]) || 0,
    finished: r[5] === true || r[5] === 'TRUE'
  }));
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const id = String(d.id || '').replace(/[^\w-]/g, '').slice(0, 40);
    if (!id) return json_({ ok: false, error: 'missing id' });
    const name = cleanName_(d.name);
    const islands = clampInt_(d.islands, 0, MAX_ISLANDS);
    const stars = clampInt_(d.stars, 0, MAX_STARS);
    const accuracy = clampInt_(d.accuracy, 0, 100);
    const finished = d.finished === true;
    const now = new Date();

    const sh = sheet_();
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
    const idx = ids.indexOf(id);
    if (idx < 0) {
      sh.appendRow([id, name, islands, stars, accuracy, finished, now]);
    } else {
      const row = idx + 2;
      const cur = sh.getRange(row, 1, 1, 7).getValues()[0];
      const better = islands > cur[2] || (islands === cur[2] && stars >= cur[3]);
      sh.getRange(row, 2).setValue(name);
      if (better) sh.getRange(row, 3, 1, 5).setValues([[islands, stars, accuracy, finished || cur[5] === true, now]]);
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
