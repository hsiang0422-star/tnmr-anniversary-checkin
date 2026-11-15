/**
 * 臺南特教｜校慶來賓簽到 — 資料 API（Google Apps Script）
 *
 * 網頁放在 GitHub Pages，這支程式只負責讀寫兩份試算表：
 *   「臺南特教｜校慶來賓簽到（資料）」→ 分頁「來賓資料」：原始資料（列的順序＝名單順序）
 *   「臺南特教｜校慶來賓簽到名單」    → 分頁「來賓名單」：每次變動後自動重寫，給人看、列印用
 *
 * API：
 *   GET  ?action=list                             → { ok, guests:[{id,name,title,time}] }
 *   POST {action:'check', pass}                   → { ok }
 *   POST {action:'add',    pass, name, title}     → { ok, guests }
 *   POST {action:'edit',   pass, id, name, title} → { ok, guests }
 *   POST {action:'delete', pass, id}              → { ok, guests }
 *   POST {action:'move',   pass, id, dir}         → { ok, guests }   dir：-1 上移、1 下移
 *   失敗時 → { ok:false, error:'訊息' }
 *
 * 第一次使用：在編輯器選 setup 執行一次，到「執行紀錄」看管理密碼。
 * 管理密碼存在「指令碼屬性」，不會出現在 GitHub。
 */

const CFG = {
  TZ: 'Asia/Taipei',
  DATA_SS_ID: '1GDYeThip-NZPYwDaDiGbpTjzP5kErhQMoYkJBJVN2ec', // 臺南特教｜校慶來賓簽到（資料）
  LIST_SS_ID: '1oT9g3zbINKStvBeoyfCLpx4Miqj07MGLpmVcoQoIO2Y', // 臺南特教｜校慶來賓簽到名單
  DATA_SHEET: '來賓資料',
  LIST_SHEET: '來賓名單',
  LIST_FIRST_ROW: 4, // 來賓名單：第 1–2 列標題說明、第 3 列欄名、第 4 列起為名單
  MAX_FAILS: 20,     // 10 分鐘內密碼錯誤超過這個次數就暫停管理功能
};

const DATA_HEADERS = ['ID', '姓名', '職稱', '簽到時間'];

/* ======================== 初始化（執行一次） ======================== */

function setup() {
  const ss = SpreadsheetApp.openById(CFG.DATA_SS_ID);
  let sh = ss.getSheetByName(CFG.DATA_SHEET);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(CFG.DATA_SHEET);
  }
  if (sh.getRange(1, 1).getValue() !== 'ID') {
    sh.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0f8');
    sh.setFrozenRows(1);
  }
  sh.getRange('D:D').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  writeListSheet_(readGuests_());

  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADMIN_PASS')) {
    p.setProperty('ADMIN_PASS', String(Math.floor(100000 + Math.random() * 900000)));
  }
  Logger.log('完成。簽到處管理密碼：' + p.getProperty('ADMIN_PASS'));
}

// 想自訂密碼：把下面的 '' 改成新密碼，執行一次 setPasscode，再改回 ''
function setPasscode() {
  const NEW_PASS = '';
  if (!NEW_PASS) throw new Error('請先在程式裡填入 NEW_PASS。');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASS', NEW_PASS);
  Logger.log('管理密碼已更新。');
}

/* ======================== API 進入點 ======================== */

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') return json_({ ok: true, guests: readGuests_().map(toClient_) });
  return json_({ ok: true, service: '臺南特教｜校慶來賓簽到 API' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkPass_(req.pass);
    switch (req.action) {
      case 'check': return json_({ ok: true });
      case 'add': return json_({ ok: true, guests: addGuest_(req.name, req.title) });
      case 'edit': return json_({ ok: true, guests: editGuest_(req.id, req.name, req.title) });
      case 'delete': return json_({ ok: true, guests: deleteGuest_(req.id) });
      case 'move': return json_({ ok: true, guests: moveGuest_(req.id, req.dir) });
      default: throw new Error('不明的操作。');
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ======================== 操作 ======================== */

function addGuest_(name, title) {
  name = clean_(name, 40);
  title = clean_(title, 60);
  if (!name) throw new Error('請輸入來賓姓名。');
  return withLock_(guests => {
    guests.push({ id: Utilities.getUuid().slice(0, 8), name, title, signedAt: new Date() });
    return guests;
  });
}

// 更正姓名與職稱，保留原簽到時間與名單位置
function editGuest_(id, name, title) {
  name = clean_(name, 40);
  title = clean_(title, 60);
  if (!name) throw new Error('請輸入來賓姓名。');
  return withLock_(guests => {
    const g = guests.find(x => x.id === id);
    if (!g) throw new Error('找不到這位來賓，可能已被刪除。');
    g.name = name;
    g.title = title;
    return guests;
  });
}

function deleteGuest_(id) {
  return withLock_(guests => {
    const next = guests.filter(x => x.id !== id);
    if (next.length === guests.length) throw new Error('找不到這位來賓，可能已被刪除。');
    return next;
  });
}

function moveGuest_(id, dir) {
  return withLock_(guests => {
    const i = guests.findIndex(x => x.id === id);
    const j = i + (Number(dir) < 0 ? -1 : 1);
    if (i < 0) throw new Error('找不到這位來賓，可能已被刪除。');
    if (j < 0 || j >= guests.length) return guests;
    [guests[i], guests[j]] = [guests[j], guests[i]];
    return guests;
  });
}

/* ======================== 內部 ======================== */

function checkPass_(pass) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= CFG.MAX_FAILS) throw new Error('密碼錯誤次數過多，請 10 分鐘後再試。');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
  if (!real || String(pass || '') !== real) {
    cache.put('fails', String(fails + 1), 600);
    throw new Error('管理密碼錯誤。');
  }
}

function dataSheet_() {
  return SpreadsheetApp.openById(CFG.DATA_SS_ID).getSheetByName(CFG.DATA_SHEET);
}

function readGuests_() {
  const sh = dataSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), name: String(r[1]), title: String(r[2]), signedAt: r[3] }));
}

function writeGuests_(guests) {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  if (guests.length) {
    sh.getRange(2, 1, guests.length, 4)
      .setValues(guests.map(g => [g.id, g.name, g.title, g.signedAt]));
  }
}

// 「來賓名單」：依目前順序從 1 連續編號，簽到時間只顯示時:分
function writeListSheet_(guests) {
  const sh = SpreadsheetApp.openById(CFG.LIST_SS_ID).getSheetByName(CFG.LIST_SHEET);
  if (!sh) throw new Error('找不到「' + CFG.LIST_SHEET + '」分頁。');
  const first = CFG.LIST_FIRST_ROW;
  const last = sh.getLastRow();
  if (last >= first) sh.getRange(first, 1, last - first + 1, 4).clearContent();
  if (guests.length) {
    const rng = sh.getRange(first, 1, guests.length, 4);
    rng.offset(0, 3, guests.length, 1).setNumberFormat('@');
    rng.setValues(guests.map((g, i) => [i + 1, g.name, g.title, hhmm_(g.signedAt)]));
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const next = fn(readGuests_());
    writeGuests_(next);
    writeListSheet_(next);
    SpreadsheetApp.flush();
    return next.map(toClient_);
  } finally {
    lock.releaseLock();
  }
}

function toClient_(g) {
  return { id: g.id, name: g.name, title: g.title, time: hhmm_(g.signedAt) };
}

function hhmm_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '--:--';
  return Utilities.formatDate(d, CFG.TZ, 'HH:mm');
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
