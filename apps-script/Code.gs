/**
 * 臺南特教｜校慶來賓簽到（Google Apps Script 網頁應用程式）
 *
 * 網頁與資料都在這支程式裡，只允許學校網域帳號登入使用。
 * 兩個入口（同一個網址）：
 *   簽到處管理：<網址>
 *   校長手機查閱：<網址>?page=view
 *
 * 資料存放：
 *   「臺南特教｜校慶來賓簽到（資料）」→ 分頁「來賓資料」：原始資料（列的順序＝名單順序）
 *   「臺南特教｜校慶來賓簽到名單」    → 分頁「來賓名單」：每次變動後自動重寫，給人看、列印用
 *
 * 管理密碼存在「指令碼屬性」ADMIN_PASS，不寫在程式碼裡。
 * 第一次使用：在編輯器選 setup 執行一次。
 * 之後要改密碼：專案設定 → 指令碼屬性 → 編輯 ADMIN_PASS。
 */

const CFG = {
  TZ: 'Asia/Taipei',
  TITLE: '臺南特教｜校慶來賓簽到',
  DATA_SS_ID: '1GDYeThip-NZPYwDaDiGbpTjzP5kErhQMoYkJBJVN2ec', // 臺南特教｜校慶來賓簽到（資料）
  LIST_SS_ID: '1oT9g3zbINKStvBeoyfCLpx4Miqj07MGLpmVcoQoIO2Y', // 臺南特教｜校慶來賓簽到名單
  DATA_SHEET: '來賓資料',
  LIST_SHEET: '來賓名單',
  LIST_FIRST_ROW: 4, // 來賓名單：第 1–3 列為標題與欄名，第 4 列起為名單
  MAX_FAILS: 20,     // 10 分鐘內密碼錯誤超過這個次數就暫停管理功能
  XLSX_FOLDER: '校慶來賓簽到_Excel', // 雲端硬碟中存放 Excel 資料庫與存檔的資料夾
};

const DATA_HEADERS = ['ID', '姓名', '職稱', '簽到時間'];

/* ======================== 網頁進入點 ======================== */

function doGet(e) {
  const page = e && e.parameter && e.parameter.page === 'view' ? 'view' : 'checkin';
  const t = HtmlService.createTemplateFromFile('Index');
  t.mode = page;
  t.baseUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle(page === 'view' ? CFG.TITLE + '（查閱）' : CFG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

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

/* ======================== 給網頁呼叫的函式（google.script.run） ======================== */

// 名單：學校網域內的使用者都能看
function getList() {
  return readGuests_().map(toClient_);
}

// 登入時確認密碼
function checkPass(pass) {
  checkPass_(pass);
  return true;
}

function addGuest(pass, name, title) {
  checkPass_(pass);
  name = clean_(name, 40);
  title = clean_(title, 60);
  if (!name) throw new Error('請輸入來賓姓名。');
  return withLock_(guests => {
    guests.push({ id: Utilities.getUuid().slice(0, 8), name, title, signedAt: new Date() });
    return guests;
  }, name);
}

// 更正姓名與職稱，保留原簽到時間與名單位置
function editGuest(pass, id, name, title) {
  checkPass_(pass);
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

function deleteGuest(pass, id) {
  checkPass_(pass);
  return withLock_(guests => {
    const next = guests.filter(x => x.id !== id);
    if (next.length === guests.length) throw new Error('找不到這位來賓，可能已被刪除。');
    return next;
  });
}

// dir：-1 上移、1 下移
function moveGuest(pass, id, dir) {
  checkPass_(pass);
  return withLock_(guests => {
    const i = guests.findIndex(x => x.id === id);
    if (i < 0) throw new Error('找不到這位來賓，可能已被刪除。');
    const j = i + (Number(dir) < 0 ? -1 : 1);
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

// archiveName：新增來賓時傳入姓名，會另外存一份 Excel 存檔
function withLock_(fn, archiveName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const next = fn(readGuests_());
    writeGuests_(next);
    writeListSheet_(next);
    SpreadsheetApp.flush();
    syncExcel_(next, archiveName);
    return next.map(toClient_);
  } finally {
    lock.releaseLock();
  }
}

/* ======================== Excel 資料庫與存檔 ======================== */

// 雲端硬碟：「校慶來賓簽到_Excel」→ 資料庫檔案 ＋ 「來賓存檔」資料夾
function getXlsxFolders_() {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CFG.XLSX_FOLDER);
  const main = it.hasNext() ? it.next() : root.createFolder(CFG.XLSX_FOLDER);
  const it2 = main.getFoldersByName('來賓存檔');
  const archive = it2.hasNext() ? it2.next() : main.createFolder('來賓存檔');
  return { main, archive };
}

// 資料庫檔：每次變動都覆寫為最新版本（舊檔丟到垃圾桶）
function syncExcel_(guests, archiveName) {
  try {
    const f = getXlsxFolders_();
    const dbName = '校慶來賓簽到資料庫.xlsx';
    const blob = DriveApp.getFileById(CFG.DATA_SS_ID).getAs(MimeType.MICROSOFT_EXCEL).setName(dbName);
    const olds = f.main.getFilesByName(dbName);
    while (olds.hasNext()) olds.next().setTrashed(true);
    f.main.createFile(blob);

    // 新增來賓時，另存一份名單 Excel 存檔
    if (archiveName) {
      const stamp = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd_HHmmss');
      const no = guests.length;
      const arch = DriveApp.getFileById(CFG.LIST_SS_ID).getAs(MimeType.MICROSOFT_EXCEL)
        .setName(no + '_' + archiveName + '_' + stamp + '.xlsx');
      f.archive.createFile(arch);
    }
  } catch (err) {
    // Excel 存檔失敗不影響簽到本身
    Logger.log('Excel 存檔失敗：' + err);
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
