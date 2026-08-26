/**
 * 情緒素養量表｜Google Sheets + Apps Script 後端
 *
 * 使用方式：
 * 1. 將這份檔案與 index.html 放進「由試算表開啟」的 Apps Script 專案。
 * 2. 先執行 setupScaleBackend() 建立「回覆」分頁。
 * 3. 執行 setAdminToken()，在試算表授權視窗輸入一組後台代碼。
 * 4. 部署為網頁應用程式：執行身分選「我」、誰可以存取選「所有人」。
 */

const SHEET_NAME = '回覆';
const TIME_ZONE = 'Asia/Taipei';
const HEADERS = [
  '時間', '匿名編號', '總分', '總量表平均',
  '自我覺察總分', '自我覺察平均',
  '自我管理總分', '自我管理平均',
  '社會覺察總分', '社會覺察平均',
  '人際技能總分', '人際技能平均', '答案JSON'
];

const FACTORS = [
  { id: 'self-awareness', items: [1, 2, 3, 4, 5, 6, 7, 8, 9] },
  { id: 'self-management', items: [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20] },
  { id: 'social-awareness', items: [21, 22, 23, 24, 25, 26, 27] },
  { id: 'relationship-skills', items: [28, 29, 30, 31, 32, 33, 34, 35] }
];

/** 公開頁面入口 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('情緒素養量表｜35 題自我檢視')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 第一次設定：建立資料分頁並檢查表頭。 */
function setupScaleBackend() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  Logger.log('工作表就緒：' + sheet.getName() + '，目前有 ' + Math.max(0, sheet.getLastRow() - 1) + ' 筆資料');
}

/** 只在第一次設定或要更換代碼時手動執行。代碼只會存於 Script Properties。 */
function setAdminToken() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('設定情緒素養量表後台代碼', '請輸入至少 8 碼的代碼（不要使用姓名或容易猜到的內容）', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const token = String(result.getResponseText() || '').trim();
  if (token.length < 8) throw new Error('後台代碼至少需要 8 碼。');
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', token);
  ui.alert('後台代碼已儲存。');
}

/** 填答者送出後呼叫：在鎖定狀態下驗證並新增一列。 */
function submitResponse(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const answers = normalizeAnswers_(payload && payload.answers);
    if (!answers) return { ok: false, error: '答案資料不完整，請重新填答。' };

    const scores = scoreAnswers_(answers);
    const id = 'EL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(), id,
      scores.overall.total, scores.overall.mean,
      scores.factors['self-awareness'].total, scores.factors['self-awareness'].mean,
      scores.factors['self-management'].total, scores.factors['self-management'].mean,
      scores.factors['social-awareness'].total, scores.factors['social-awareness'].mean,
      scores.factors['relationship-skills'].total, scores.factors['relationship-skills'].mean,
      JSON.stringify(answers)
    ]);
    return { ok: true, id: id };
  } catch (error) {
    return { ok: false, error: '資料暫時無法寫入，請稍後再試。' };
  } finally {
    lock.releaseLock();
  }
}

/** 後台呼叫：只有輸入 Script Properties 中的代碼才會回傳個別回覆。 */
function getAdminData(token) {
  const savedToken = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!savedToken) return { ok: false, error: '尚未設定後台代碼，請先在 Apps Script 執行 setAdminToken()。'};
  if (String(token || '') !== savedToken) return { ok: false, error: '後台代碼不正確。' };

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const rows = values.map(function (row) {
    const answers = parseAnswers_(row[12]);
    if (!answers) return null;
    const scores = scoreAnswers_(answers);
    const time = row[0] instanceof Date ? row[0] : new Date(row[0]);
    return {
      id: String(row[1] || ''),
      submittedAt: Utilities.formatDate(time, TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss"),
      answers: answers,
      scores: scores
    };
  }).filter(Boolean);
  return { ok: true, rows: rows };
}

function getSheet_() {
  // 本檔案預設為「由試算表開啟」的容器繫結腳本，因此不需要硬寫試算表 ID。
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function normalizeAnswers_(answers) {
  if (!answers || typeof answers !== 'object') return null;
  const normalized = {};
  for (let id = 1; id <= 35; id += 1) {
    const value = Number(answers[id]);
    if (!Number.isInteger(value) || value < 1 || value > 6) return null;
    normalized[id] = value;
  }
  return normalized;
}

function parseAnswers_(json) {
  try {
    return normalizeAnswers_(JSON.parse(String(json || '{}')));
  } catch (error) {
    return null;
  }
}

function scoreAnswers_(answers) {
  const factors = {};
  let overallTotal = 0;
  FACTORS.forEach(function (factor) {
    const total = factor.items.reduce(function (sum, id) { return sum + answers[id]; }, 0);
    const mean = Number((total / factor.items.length).toFixed(4));
    overallTotal += total;
    factors[factor.id] = {
      total: total,
      max: factor.items.length * 6,
      count: factor.items.length,
      mean: mean
    };
  });
  return {
    overall: {
      total: overallTotal,
      max: 210,
      count: 35,
      mean: Number((overallTotal / 35).toFixed(4))
    },
    factors: factors
  };
}

