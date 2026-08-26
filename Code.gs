/**
 * 情緒素養量表｜Google Sheets + Apps Script 後端
 *
 * 使用方式：
 * 1. 將這份檔案與 index.html 放進 Apps Script 專案；可使用試算表綁定專案，也可使用獨立專案。
 * 2. 先執行 setupScaleBackend() 建立「回覆」分頁。
 * 3. 在「專案設定 → 指令碼屬性」新增 ADMIN_TOKEN，設定至少 8 碼的管理者代碼。
 * 4. 部署為網頁應用程式：執行身分選「我」、誰可以存取選「所有人」。
 * 5. 公開 GitHub Pages 會以 POST 呼叫 doPost()；一般填答者不需要直接開啟 Apps Script 頁面。
 */

const SHEET_NAME = '回覆';
const TIME_ZONE = 'Asia/Taipei';
// 若是獨立 Apps Script 專案，請在這裡填入試算表網址中的 ID；綁定專案可維持原樣。
const SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
const HEADERS = [
  '時間', '匿名編號', '填答身分', '性別', '年齡', '任教階段', '教學年資', '目前主要工作／身分',
  '總分', '總量表平均',
  '自我覺察總分', '自我覺察平均',
  '自我管理總分', '自我管理平均',
  '社會覺察總分', '社會覺察平均',
  '人際技能總分', '人際技能平均', '答案JSON'
];

const PROFILE_OPTIONS = {
  respondentTypes: ['teacher', 'non-teacher'],
  genders: ['女性', '男性', '其他／不便回答'],
  ages: ['25 歲以下', '26-30 歲', '31-40 歲', '41-50 歲', '51 歲以上', '不便回答'],
  teachingStages: ['幼兒園', '國小', '國中', '高中', '其他教育階段'],
  teachingExperiences: ['０師培生', '初任教師（未滿 2 年）', '2 年以上未達 5 年', '5 年以上未達 10 年', '10 年以上未達 15 年', '15 年以上未達 20 年', '20 年以上'],
  occupations: ['學生／在學者', '公務／軍警', '教育／研究（非教師）', '醫療／健康／社會照顧', '科技／工程／製造', '商業／金融／行政', '服務／銷售／餐旅', '文化／媒體／自由業', '農林漁牧／營建／運輸', '家務／退休／待業', '其他']
};

const FACTORS = [
  { id: 'self-awareness', items: [1, 2, 3, 4, 5, 6, 7, 8, 9] },
  { id: 'self-management', items: [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20] },
  { id: 'social-awareness', items: [21, 22, 23, 24, 25, 26, 27] },
  { id: 'relationship-skills', items: [28, 29, 30, 31, 32, 33, 34, 35] }
];

/** 公開頁面入口；mode=bridge 是給 GitHub Pages 呼叫 Apps Script 的隱藏橋接頁。 */
function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'bridge') return createBridge_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('情緒素養量表｜35 題自我檢視')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 公開 GitHub Pages 的無登入送出入口。
 * 前端使用 text/plain POST，避免跨網域 JSON 預檢；回覆內容不含任何試算表資料。
 */
function doPost(e) {
  try {
    const request = parsePostRequest_(e);
    if (!request) return jsonOutput_({ ok: false, error: '無法解析送出資料。' });
    if (request.action && request.action !== 'submitResponse') {
      return jsonOutput_({ ok: false, error: '不支援的雲端操作。' });
    }
    return jsonOutput_(submitResponse(request.payload || request));
  } catch (error) {
    return jsonOutput_({ ok: false, error: '資料暫時無法寫入，請稍後再試。' });
  }
}

/**
 * 讓 GitHub Pages 在不使用 CORS、不暴露試算表資料的情況下呼叫後端。
 * 這個頁面只接受兩個白名單操作：送出回覆、以管理者代碼讀取後台。
 */
function createBridge_() {
  const html = `<!doctype html>
<html lang="zh-Hant">
<head><meta charset="utf-8"><title>情緒素養量表資料服務</title></head>
<body>
<script>
(function () {
  function reply(requestId, response) {
    window.parent.postMessage({
      source: 'el-scale-gas-bridge',
      type: 'response',
      requestId: requestId,
      response: response
    }, '*');
  }

  window.addEventListener('message', function (event) {
    var data = event.data || {};
    if (data.source !== 'el-scale-gas-client' || data.type !== 'call') return;

    var runner = google.script.run
      .withSuccessHandler(function (response) { reply(data.requestId, response); })
      .withFailureHandler(function (error) {
        reply(data.requestId, { ok: false, error: error && error.message ? error.message : '雲端服務發生錯誤。' });
      });

    if (data.method === 'submitResponse') runner.submitResponse(data.payload);
    else if (data.method === 'getAdminData') runner.getAdminData(data.payload);
    else reply(data.requestId, { ok: false, error: '不支援的雲端操作。' });
  });

  window.parent.postMessage({ source: 'el-scale-gas-bridge', type: 'ready' }, '*');
}());
</script>
</body>
</html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('情緒素養量表資料服務')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 第一次設定：建立資料分頁並檢查表頭。 */
function setupScaleBackend() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  Logger.log('工作表就緒：' + sheet.getName() + '，目前有 ' + Math.max(0, sheet.getLastRow() - 1) + ' 筆資料');
}

/** 填答者送出後呼叫：在鎖定狀態下驗證並新增一列。 */
function submitResponse(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const answers = normalizeAnswers_(payload && payload.answers);
    if (!answers) return { ok: false, error: '答案資料不完整，請重新填答。' };
    const profile = normalizeProfile_(payload && payload.profile);
    if (!profile) return { ok: false, error: '基本資料不完整，請返回前一步重新填寫。' };

    const scores = scoreAnswers_(answers);
    const sheet = getSheet_();
    // 公開頁面可能在網路逾時後重試；以同一個前端 ID 去重，避免重複寫入。
    const requestedId = String(payload && payload.id || '').trim().slice(0, 80);
    if (requestedId && responseIdExists_(sheet, requestedId)) {
      return { ok: true, id: requestedId, duplicate: true };
    }
    const id = requestedId || 'EL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([
      new Date(), id, profile.respondentType, profile.gender, profile.age,
      profile.teachingStage, profile.teachingExperience, profile.occupation,
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

function parsePostRequest_(e) {
  const raw = e && e.postData && e.postData.contents;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (error) {
      // 也接受表單 fallback 的 payload 欄位。
    }
  }
  const formPayload = e && e.parameter && e.parameter.payload;
  if (!formPayload) return null;
  try {
    const parsed = JSON.parse(String(formPayload));
    return parsed && typeof parsed === 'object' ? { action: 'submitResponse', payload: parsed } : null;
  } catch (error) {
    return null;
  }
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function responseIdExists_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  return sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().some(function (row) {
    return String(row[0] || '') === id;
  });
}

/** 後台呼叫：只有輸入 Script Properties 中的代碼才會回傳個別回覆。 */
function getAdminData(token) {
  const savedToken = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!savedToken) return { ok: false, error: '尚未設定後台代碼，請到「專案設定 → 指令碼屬性」新增 ADMIN_TOKEN。'};
  if (String(token || '') !== savedToken) return { ok: false, error: '後台代碼不正確。' };

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const rows = values.map(function (row) {
    const answers = parseAnswers_(row[HEADERS.indexOf('答案JSON')]);
    if (!answers) return null;
    const scores = scoreAnswers_(answers);
    const time = row[0] instanceof Date ? row[0] : new Date(row[0]);
    const profile = normalizeProfile_({
      respondentType: row[2],
      gender: row[3],
      age: row[4],
      teachingStage: row[5],
      teachingExperience: row[6],
      occupation: row[7]
    });
    return {
      id: String(row[1] || ''),
      submittedAt: Utilities.formatDate(time, TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss"),
      profile: profile,
      answers: answers,
      scores: scores
    };
  }).filter(Boolean);
  return { ok: true, rows: rows };
}

function getSheet_() {
  // 綁定專案可直接取得目前試算表；獨立專案則使用上方設定的試算表 ID。
  let spreadsheet = null;
  try {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    spreadsheet = null;
  }
  if (!spreadsheet) {
    if (!SPREADSHEET_ID || SPREADSHEET_ID === 'PASTE_YOUR_SPREADSHEET_ID_HERE') {
      throw new Error('請先在 Code.gs 設定 SPREADSHEET_ID。');
    }
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
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

function normalizeProfile_(profile) {
  if (!profile || typeof profile !== 'object') return null;
  const respondentType = String(profile.respondentType || '').trim();
  const gender = String(profile.gender || '').trim();
  const age = String(profile.age || '').trim();
  if (PROFILE_OPTIONS.respondentTypes.indexOf(respondentType) === -1) return null;
  if (PROFILE_OPTIONS.genders.indexOf(gender) === -1) return null;
  if (PROFILE_OPTIONS.ages.indexOf(age) === -1) return null;
  const normalized = { respondentType: respondentType, gender: gender, age: age, teachingStage: '', teachingExperience: '', occupation: '' };
  if (respondentType === 'teacher') {
    const teachingStage = String(profile.teachingStage || '').trim();
    const teachingExperience = String(profile.teachingExperience || '').trim();
    if (PROFILE_OPTIONS.teachingStages.indexOf(teachingStage) === -1) return null;
    if (PROFILE_OPTIONS.teachingExperiences.indexOf(teachingExperience) === -1) return null;
    normalized.teachingStage = teachingStage;
    normalized.teachingExperience = teachingExperience;
  } else {
    const occupation = String(profile.occupation || '').trim();
    if (PROFILE_OPTIONS.occupations.indexOf(occupation) === -1) return null;
    normalized.occupation = occupation;
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
