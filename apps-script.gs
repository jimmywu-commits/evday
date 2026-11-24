// 蛋白質追蹤 ↔ Google 試算表同步
// 貼到試算表的「擴充功能 → Apps Script」，執行一次 setup 後部署為網頁應用程式。

const REC_SHEET = '紀錄';
const DAY_SHEET = '每日統計';
const CFG_SHEET = '設定';
const REC_HEAD = ['id', '日期', '時間', '時段', '食物', '份量', '蛋白質g', 'foodId', '寫入時間'];

// 知識庫初始資料的位置（GitHub Pages 上的 kb-seed.json）
const SEED_URL = 'https://jimmywu-commits.github.io/evday/kb-seed.json';

const KB = {
  '食物庫': { head: ['id', '名稱', '分類', '份量', '蛋白質g', '熱量kcal', '備註', '啟用'], key: ['名稱', '份量'], need: ['名稱', '份量', '蛋白質g'], prefix: 'f' },
  '知識庫': { head: ['id', '分類', '標題', '內容', '啟用'], key: ['標題'], need: ['分類', '標題', '內容'], prefix: 't' },
  '測驗題': { head: ['id', '題目', '答案', '解說', '啟用'], key: ['題目'], need: ['題目', '答案', '解說'], prefix: 'q' },
  '建議文案': { head: ['id', '代碼', '文案', '啟用'], key: ['代碼', '文案'], need: ['代碼', '文案'], prefix: 'm' },
};

// 第一次貼上後，在上方函式選單選「setup」按執行：完成授權並建立分頁。
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  recSheet_(ss);
  ensureDaySheet_(ss);
  cfgSheet_(ss);
  Object.keys(KB).forEach(name => kbSheet_(ss, name));
  Logger.log('完成：已建立紀錄、每日統計、設定與 4 個知識庫分頁。下一步：執行 seedKnowledgeBase 匯入初始資料。');
}

// 從 GitHub Pages 匯入知識庫初始資料；只會新增試算表裡還沒有的資料，可以重複執行。
function seedKnowledgeBase() {
  const res = UrlFetchApp.fetch(SEED_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('讀不到 ' + SEED_URL + '（HTTP ' + res.getResponseCode() + '），請確認 kb-seed.json 已上傳到 GitHub。');
  const seed = JSON.parse(res.getContentText('UTF-8'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = Object.keys(KB).map(name => {
    const r = appendKb_(ss, name, seed[name] || []);
    return name + '：新增 ' + r.added + '、略過 ' + r.skipped;
  });
  Logger.log('知識庫匯入完成（版本 ' + seed.version + '）\n' + report.join('\n'));
}

function doGet() {
  return out_({ ok: true, message: '蛋白質追蹤同步服務運作中' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: '資料格式錯誤' });
  }
  if (!body) return out_({ ok: false, error: '資料格式錯誤' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (body.action === 'kb') return out_({ ok: true, kb: readKb_(ss) });
    if (body.action === 'kbAppend') {
      if (!KB[body.sheet]) return out_({ ok: false, error: '沒有這個知識庫分頁：' + body.sheet });
      const rows = Array.isArray(body.rows) ? body.rows.slice(0, 500) : [];
      return out_(Object.assign({ ok: true }, appendKb_(ss, body.sheet, rows)));
    }
    const rec = recSheet_(ss);
    ensureDaySheet_(ss);
    const ops = Array.isArray(body.ops) ? body.ops.slice(0, 2000) : [];
    const applied = [];
    for (const op of ops) {
      if (!op || !op.opId) continue;
      if (op.type === 'add' && op.rec) addRecord_(rec, op.rec);
      else if (op.type === 'del' && op.id) deleteRecord_(rec, String(op.id));
      else if (op.type === 'config' && op.config) writeConfig_(ss, op.config);
      applied.push(op.opId);
    }
    return out_({ ok: true, applied: applied, records: readRecords_(rec), config: readConfig_(ss) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function recSheet_(ss) {
  let sh = ss.getSheetByName(REC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REC_SHEET);
    sh.getRange(1, 1, 1, REC_HEAD.length).setValues([REC_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:F').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
    sh.setColumnWidth(1, 120);
    sh.hideColumns(1);
    sh.hideColumns(8);
  }
  return sh;
}

function ensureDaySheet_(ss) {
  if (ss.getSheetByName(DAY_SHEET)) return;
  const sh = ss.insertSheet(DAY_SHEET);
  sh.getRange('A1').setFormula(
    "=QUERY('" + REC_SHEET + "'!B:G,\"select B, sum(G), count(G) where B is not null group by B order by B desc label B '日期', sum(G) '蛋白質g', count(G) '筆數'\",1)"
  );
  sh.setFrozenRows(1);
}

function idIndex_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return {};
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
  const map = {};
  ids.forEach((r, i) => { if (r[0]) map[r[0]] = i + 2; });
  return map;
}

function addRecord_(sh, r) {
  const id = String(r.id || '');
  if (!id || idIndex_(sh)[id]) return;
  const date = String(r.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
  sh.appendRow([
    id, date, String(r.time || '').slice(0, 10), String(r.slot || '').slice(0, 10),
    String(r.name || '').slice(0, 40), String(r.unit || '').slice(0, 40),
    Number(r.g) || 0, String(r.foodId || '').slice(0, 40), new Date(),
  ]);
}

function deleteRecord_(sh, id) {
  const row = idIndex_(sh)[id];
  if (row) sh.deleteRow(row);
}

function readRecords_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, 8).getDisplayValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ id: r[0], date: r[1], time: r[2], slot: r[3], name: r[4], unit: r[5], g: Number(r[6]) || 0, foodId: r[7] }));
}

function cfgSheet_(ss) {
  let sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG_SHEET);
    sh.getRange('A1:B1').setValues([['設定（App 自動寫入，請勿手動修改）', '']]);
    sh.getRange('A2:B2').setNumberFormat('@');
  }
  return sh;
}

function writeConfig_(ss, config) {
  const cur = readConfig_(ss);
  if (cur && Number(cur.updatedAt) > Number(config.updatedAt)) return;
  cfgSheet_(ss).getRange('A2:B2').setValues([[JSON.stringify(config), String(config.updatedAt || '')]]);
}

function readConfig_(ss) {
  const sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) return null;
  try {
    return JSON.parse(sh.getRange('A2').getDisplayValue());
  } catch (err) {
    return null;
  }
}

/* ---------- 知識庫 ---------- */
function kbSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const head = KB[name].head;
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    head.forEach((h, i) => {
      if (h !== '蛋白質g' && h !== '熱量kcal') sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sh;
}

function kbRows_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const head = values[0];
  return values.slice(1).filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  });
}

function readKb_(ss) {
  const out = {};
  Object.keys(KB).forEach(name => {
    const sh = ss.getSheetByName(name);
    out[name] = sh ? kbRows_(sh) : [];
  });
  return out;
}

function kbKey_(name, row) {
  return KB[name].key.map(k => String(row[k] || '').trim()).join('|');
}

function cleanKbRow_(name, row) {
  const o = {};
  KB[name].head.forEach(h => {
    let v = row[h] === undefined || row[h] === null ? '' : row[h];
    if (h === '蛋白質g' || h === '熱量kcal') {
      v = v === '' ? '' : Number(v);
      if (v !== '' && !(v >= 0 && v <= 1000)) v = '';
    } else {
      v = String(v).trim().slice(0, h === '內容' || h === '解說' || h === '文案' ? 500 : 80);
    }
    o[h] = v;
  });
  if (KB[name].need.some(h => o[h] === '' || o[h] === undefined)) return null;
  if (name === '測驗題' && o['答案'] !== '對' && o['答案'] !== '錯') return null;
  if (name === '建議文案' && !/^[A-Z0-9_]+$/.test(o['代碼'])) return null;
  if (!o['啟用']) o['啟用'] = 'Y';
  return o;
}

function appendKb_(ss, name, rows) {
  const sh = kbSheet_(ss, name);
  const existing = kbRows_(sh);
  const ids = new Set(existing.map(r => r.id));
  const keys = new Set(existing.map(r => kbKey_(name, r)));
  const head = KB[name].head;
  const stamp = Date.now().toString(36);
  const toWrite = [];
  let skipped = 0;
  rows.forEach((raw, i) => {
    const o = raw && cleanKbRow_(name, raw);
    if (!o) { skipped++; return; }
    if (!o.id) o.id = KB[name].prefix + stamp + i;
    const k = kbKey_(name, o);
    if (ids.has(o.id) || keys.has(k)) { skipped++; return; }
    ids.add(o.id);
    keys.add(k);
    toWrite.push(head.map(h => o[h]));
  });
  if (toWrite.length) sh.getRange(sh.getLastRow() + 1, 1, toWrite.length, head.length).setValues(toWrite);
  return { added: toWrite.length, skipped: skipped };
}
