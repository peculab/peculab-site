/**
 * PecuLab 網站報名表 → Google Sheet
 *
 * 用法：在 Sheet「peculab-site」裡開啟「擴充功能 → Apps Script」，
 * 把這整份貼進 Code.gs，存檔後先執行一次 setup()，再部署為網頁應用程式。
 *
 * 網站用 fetch POST JSON 到 /exec，這裡的 doPost 驗證後寫入「registrations」分頁，
 * 寫入成功才回傳 ok:true 與報名編號。網站會讀到這個回應，不使用 no-cors。
 *
 * 這是公開可呼叫的網址：欄位驗證、honeypot、寫入鎖、同編號防重複
 * 都不能取代身分驗證與完整防濫用。請勿用來收集敏感資料。
 */

// ===== 設定 =====
const SHEET_NAME = 'registrations';
const HEADERS = [
  'received_at', 'request_id', 'name', 'email', 'phone',
  'program', 'organization', 'goal', 'referral', 'page', 'status'
];
// 必須與 index.html 報名表「報名項目」的選項完全相同
const PROGRAMS = ['Claude AI Work System 90 天陪跑計畫', '其他（請在備註說明）'];
// 必須與 index.html「從哪裡得知」的選項相同；空字串代表未選
const REFERRALS = ['', '朋友或同事介紹', '社群媒體', '搜尋引擎', '過去的課程或活動', '其他'];

// ===== 第一次設定：手動執行一次 =====
function setup() {
  const sh = getSheet_();
  Logger.log('已就緒：分頁「' + sh.getName() + '」，欄位 ' + HEADERS.join(', '));
}

// ===== 網址健康檢查：在瀏覽器打開 /exec 會看到 ok:true =====
function doGet() {
  return json_({ ok: true, service: 'peculab-signup' });
}

// ===== 接收報名 =====
function doPost(e) {
  let data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return json_({ ok: false, error: '資料格式錯誤，請重新整理頁面後再試。' });
  }
  if (!data || typeof data !== 'object') {
    return json_({ ok: false, error: '資料格式錯誤，請重新整理頁面後再試。' });
  }

  let v;
  try {
    if (data.website) throw new Error('送出失敗，請重新整理頁面後再試。'); // honeypot
    if (data.consent !== true) throw new Error('請勾選同意資料用途。');
    v = {
      id: text_(data.request_id, 10, 80, '報名編號'),
      name: text_(data.name, 1, 60, '姓名'),
      email: text_(data.email, 3, 254, 'Email'),
      phone: text_(data.phone || '', 0, 30, '電話'),
      program: text_(data.program, 1, 60, '報名項目'),
      organization: text_(data.organization || '', 0, 100, '服務單位或領域'),
      goal: text_(data.goal || '', 0, 1000, '想解決的問題'),
      referral: text_(data.referral || '', 0, 30, '得知管道'),
      page: text_(data.page || '', 0, 300, '來源頁面')
    };
    if (!/^[A-Za-z0-9_-]+$/.test(v.id)) throw new Error('報名編號格式錯誤，請重新整理頁面。');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.email)) throw new Error('請檢查 Email 格式。');
    if (v.phone && !/^[0-9+\-()#\s]{6,30}$/.test(v.phone)) throw new Error('電話只能包含數字、空格與 + - ( )。');
    if (PROGRAMS.indexOf(v.program) === -1) throw new Error('請選擇報名項目。');
    if (REFERRALS.indexOf(v.referral) === -1) throw new Error('請重新選擇得知管道。');
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }

  const lock = LockService.getScriptLock();
  let held = false;
  try {
    held = lock.tryLock(10000);
    if (!held) return json_({ ok: false, retryable: true, error: '目前報名人數較多，請稍後用相同編號重試。' });

    const sh = getSheet_();
    const last = sh.getLastRow();
    // 同一個報名編號只寫入一次（網路逾時後重試不會重複報名）
    if (last > 1) {
      const found = sh.getRange(2, 2, last - 1, 1)
        .createTextFinder(v.id).matchEntireCell(true).findNext();
      if (found) return json_({ ok: true, request_id: v.id, duplicate: true });
    }
    sh.appendRow([
      new Date(),
      v.id,
      safe_(v.name),
      safe_(v.email),
      v.phone ? "'" + v.phone : '',   // 前置 ' 讓 0912… 保留開頭的 0
      safe_(v.program),
      safe_(v.organization),
      safe_(v.goal),
      safe_(v.referral),
      safe_(v.page),
      'new'
    ]);
    SpreadsheetApp.flush();
    return json_({ ok: true, request_id: v.id, duplicate: false });
  } catch (err) {
    console.error('doPost failed: ' + String(err && err.message)); // 不記錄個資
    return json_({ ok: false, retryable: true, error: '目前無法確認是否已儲存，請保留此頁，稍後用相同編號重試。' });
  } finally {
    if (held) lock.releaseLock();
  }
}

// ===== 工具函式 =====
function getSheet_() {
  // 這份程式綁定在 Sheet 裡，直接用目前的試算表。
  // 若改成獨立專案，可在「專案設定 → 指令碼屬性」加入 SPREADSHEET_ID。
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet not found');

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const actual = sh.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  if (actual.join('|') !== HEADERS.join('|')) throw new Error('Header mismatch');
  return sh;
}

function text_(v, min, max, label) {
  if (typeof v !== 'string') throw new Error(label + '格式不正確。');
  const t = v.trim();
  if (t.length < min) throw new Error('請填寫' + label + '。');
  if (t.length > max) throw new Error(label + '請在 ' + max + ' 字以內。');
  return t;
}

function safe_(s) {
  // 避免使用者輸入被當成試算表公式
  return /^[\s]*[=+@\-\t\r\n]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
