// 蛋白質追蹤 ↔ Google 試算表同步
// 貼到試算表的「擴充功能 → Apps Script」，改好 TOKEN 後部署為網頁應用程式。

const TOKEN = '請改成你自己的同步密碼';

const REC_SHEET = '紀錄';
const DAY_SHEET = '每日統計';
const CFG_SHEET = '設定';
const REC_HEAD = ['id', '日期', '時間', '時段', '食物', '份量', '蛋白質g', 'foodId', '寫入時間'];

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
  if (!body || body.token !== TOKEN || TOKEN === '請改成你自己的同步密碼') {
    return out_({ ok: false, error: 'token' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
