/**
 * 呼召聚會報名網頁的後端（掛在「2026下 全時間訓練呼召訓練 台中場 (回覆)」試算表上）
 *
 * 三個動作：
 *   GET  ?action=list           → 回傳報名者名單（只有名字和是否已繳費，不含電話、身分證）
 *   POST {action:'register',…}  → 新增一列報名資料，回傳報名編號
 *   POST {action:'pay',…}       → 上傳匯款截圖到雲端硬碟，把連結寫回該列
 *   POST {action:'ask',…}       → 匿名提問，寫進「提問單」分頁（不記任何個人資料）
 *   每個動作都可帶 site：'taichung'（預設，這份試算表）或 'fengchia'（另一份「逢甲報名」試算表）
 *   提問單寫在各場自己的試算表裡
 *
 * 部署方式見同資料夾的 README.md。
 */

// 網頁送來的 site 決定寫到哪一份試算表、哪一頁（沒給就是台中場）
// book 留空 = 這支程式所在的試算表（台中場回覆）；有填 = 另一份試算表的 ID
const SITES = {
  taichung: { book: '', sheet: '表單回覆 1', label: '台中場' },
  fengchia: { book: '1b4CTYNX_g53ofGdCGk6yYba7ShvIoKdgB0u9KaYhDgg', sheet: '逢甲報名', label: '逢甲場' }
};
const FOLDER_NAME = '呼召聚會 匯款截圖';
const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10 MB

const COL = {
  time: '時間戳記',
  name: '我叫什麼名字？',
  campus: '校園',
  gender: '性別',
  role: '身份',
  phone: '手機號碼',
  nid: '身分證字號',
  bday: '出生年月日',
  stop: '哪一站上車？',
  note: '備註（若有不可抗因素無法搭乘遊覽車，請務必填寫清楚，並與服事者交通）',
  receipt: '匯款截圖或拍照證明',
  id: '報名編號'
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'list') return json_({ ok: true, people: listPeople_(e.parameter.site) });
  return json_({ ok: false, error: '未知的動作' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'register') return json_(register_(body));
    if (body.action === 'pay') return json_(pay_(body));
    if (body.action === 'ask') return json_(ask_(body));
    return json_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 名單 ---------- */
function listPeople_(site) {
  const { sheet, cols } = sheet_(site);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const rows = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  return rows
    .filter(r => String(r[cols.name - 1]).trim())
    .map(r => ({
      name: String(r[cols.name - 1]).trim(),
      gender: String(r[cols.gender - 1]).trim(),
      paid: !!String(r[cols.receipt - 1]).trim()
    }));
}

/* ---------- 報名 ---------- */
function register_(b) {
  const need = ['name', 'campus', 'gender', 'role', 'phone', 'nid', 'bday', 'stop'];
  for (const k of need) if (!String(b[k] || '').trim()) return { ok: false, error: '有必填欄位沒填' };
  const phone = String(b.phone).replace(/\D/g, '');
  if (!/^09\d{8}$/.test(phone)) return { ok: false, error: '手機號碼格式不對' };
  const nid = String(b.nid).toUpperCase().trim();
  if (!/^[A-Z][0-9A-D]\d{8}$/.test(nid)) return { ok: false, error: '身分證字號格式不對' };

  const { sheet, cols } = sheet_(b.site);
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const row = new Array(sheet.getLastColumn()).fill('');
  const put = (k, v) => { row[cols[k] - 1] = v; };
  put('time', new Date());
  put('name', clean_(b.name, 40));
  put('campus', clean_(b.campus, 40));
  put('gender', clean_(b.gender, 10));
  put('role', clean_(b.role, 40));
  put('phone', "'" + phone);          // 前面加 ' 保留開頭的 0
  put('nid', nid);
  put('bday', clean_(b.bday, 10).replace(/-/g, '/'));
  put('stop', clean_(b.stop, 40));
  put('note', clean_(b.note, 500));
  put('id', id);
  sheet.appendRow(row);
  return { ok: true, id };
}

/* ---------- 上傳匯款截圖 ---------- */
function pay_(b) {
  const { sheet, cols } = sheet_(b.site);
  const rowNo = findRow_(sheet, cols, b);
  if (!rowNo) return { ok: false, error: '找不到這筆報名，請確認名字和手機末四碼' };

  const m = String(b.file || '').match(/^data:(image\/(?:png|jpe?g|gif|webp|heic|heif));base64,(.+)$/);
  if (!m) return { ok: false, error: '請上傳圖片檔' };
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_FILE_BYTES) return { ok: false, error: '圖片超過 10 MB' };

  const name = sheet.getRange(rowNo, cols.name).getValue();
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(bytes, m[1], `${site_(b.site).label}_${name}_${Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd-HHmm')}.${ext}`);
  const file = folder_().createFile(blob);   // 只有試算表擁有者看得到
  sheet.getRange(rowNo, cols.receipt).setValue(file.getUrl());
  return { ok: true, name };
}

function findRow_(sheet, cols, b) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const id = String(b.id || '').trim();
  const name = String(b.name || '').trim();
  const p4 = String(b.phone4 || '').replace(/\D/g, '');
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    if (id && String(r[cols.id - 1]) === id) return i + 2;
    const phone = String(r[cols.phone - 1]).replace(/\D/g, '');
    if (!id && name && p4.length === 4 && String(r[cols.name - 1]).trim() === name && phone.slice(-4) === p4) return i + 2;
  }
  return 0;
}

/* ---------- 匿名提問單 ---------- */
const ASK_SHEET = '提問單';
const ASK_TOPICS = ['真理追求', '如何愛主追求主'];

function ask_(b) {
  const items = (Array.isArray(b.questions) ? b.questions : [])
    .map(q => ({ topic: String(q.topic || ''), text: clean_(q.text, 1000) }))
    .filter(q => ASK_TOPICS.indexOf(q.topic) >= 0 && q.text);
  if (!items.length) return { ok: false, error: '請至少寫一個問題' };

  const ss = book_(b.site);
  let sh = ss.getSheetByName(ASK_SHEET);
  if (!sh) sh = ss.insertSheet(ASK_SHEET);
  if (sh.getLastColumn() === 0) {
    sh.appendRow(['日期', '主題', '問題']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 520);
  }
  // 匿名：只記日期（不記時間）、主題、問題，不記名字或任何識別資料
  const day = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  items.forEach(q => sh.appendRow([day, q.topic, q.text]));
  return { ok: true, count: items.length };
}

/* ---------- 工具 ---------- */
function site_(site) {
  return SITES[site] || SITES.taichung;
}

function book_(site) {
  const id = site_(site).book;
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function sheet_(site) {
  const name = site_(site).sheet;
  const ss = book_(site);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastColumn() === 0) {              // 新場次第一次用：寫入標題列
    sheet.appendRow(Object.keys(COL).map(k => COL[k]));
    sheet.setFrozenRows(1);
  }
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  if (headers.indexOf(COL.id) < 0) {              // 第一次使用時加上「報名編號」欄
    sheet.getRange(1, headers.length + 1).setValue(COL.id);
    headers.push(COL.id);
  }
  const cols = {};
  for (const k in COL) {
    const i = headers.indexOf(COL[k]);
    if (i < 0) throw new Error('找不到欄位：' + COL[k]);
    cols[k] = i + 1;
  }
  return { sheet, cols };
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function clean_(v, max) {
  // 去掉開頭的 = + - @，避免被試算表當成公式
  return String(v || '').trim().replace(/^[=+\-@]+/, '').slice(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 在編輯器裡手動執行一次，用來授權並檢查欄位都找得到 */
function setupCheck() {
  for (const k in SITES) Logger.log(SITES[k].sheet + ' 欄位：' + JSON.stringify(sheet_(k).cols));
  folder_();
}
