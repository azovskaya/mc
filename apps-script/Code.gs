/**
 * Meal Kiosk — бэкенд на Google Apps Script
 *
 * Установка:
 * 1. Создайте Google Таблицу
 * 2. Расширения → Apps Script → вставьте этот код
 * 3. Развернуть → Новое развертывание → Веб-приложение
 *    - Выполнять от имени: Я
 *    - Доступ: Все пользователи
 * 4. Скопируйте URL (/exec) в настройки киоска
 */

var SHEET_EMPLOYEES = 'Employees';
var SHEET_MEALS = 'Meals';
var SHEET_LOGS = 'Logs';

var EMPLOYEE_HEADERS = [
  'employeeId', 'fullName', 'staffId', 'department', 'position',
  'photo', 'faceDescriptor', 'active', 'createdAt', 'updatedAt'
];

var MEAL_HEADERS = [
  'mealId', 'timestamp', 'date', 'time', 'employeeId', 'employeeName',
  'staffId', 'department', 'mealType', 'siteName', 'operator',
  'matchScore', 'photo', 'verified', 'note'
];

var LOG_HEADERS = [
  'eventId', 'timestamp', 'type', 'employeeId', 'employeeName',
  'status', 'message', 'photo', 'matchScore'
];

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    switch (action) {
      case 'ping':
        return respond({ ok: true, message: 'pong', version: '1.0' });
      case 'listEmployees':
        return respond({ ok: true, items: listEmployees() });
      case 'listMeals':
        return respond({ ok: true, items: listMeals(Number(e.parameter.limit) || 500) });
      default:
        return respond({ ok: false, status: 'error', message: 'Неизвестное действие: ' + action });
    }
  } catch (err) {
    return respond({ ok: false, status: 'error', message: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    var body = JSON.parse(e.postData.contents);
    switch (action) {
      case 'saveEmployee':
        return respond(saveEmployee(body));
      case 'saveMeal':
        return respond(saveMeal(body));
      case 'logEvent':
        return respond(logEvent(body));
      default:
        return respond({ ok: false, status: 'error', message: 'Неизвестное действие: ' + action });
    }
  } catch (err) {
    return respond({ ok: false, status: 'error', message: String(err.message || err) });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(String);
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row.some(function(cell) { return cell !== '' && cell !== null; })) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    rows.push(obj);
  }
  return rows;
}

function findRowById_(sheet, idCol, id) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var headers = data[0].map(String);
  var col = headers.indexOf(idCol);
  if (col < 0) return 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col]) === String(id)) return i + 1;
  }
  return 0;
}

function objectToRow_(headers, obj) {
  return headers.map(function(h) { return obj[h] !== undefined && obj[h] !== null ? obj[h] : ''; });
}

function nowIso_() {
  return new Date().toISOString();
}

function listEmployees() {
  var sheet = getOrCreateSheet_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  return sheetToObjects_(sheet).filter(function(e) {
    return e.active !== false && e.active !== 'false' && String(e.active).toLowerCase() !== 'нет';
  });
}

function listMeals(limit) {
  var sheet = getOrCreateSheet_(SHEET_MEALS, MEAL_HEADERS);
  var items = sheetToObjects_(sheet);
  items.sort(function(a, b) {
    return String(b.timestamp || '').localeCompare(String(a.timestamp || ''));
  });
  return items.slice(0, limit || 500);
}

function saveEmployee(body) {
  if (!body || !body.fullName) {
    return { ok: false, status: 'error', message: 'Укажите ФИО сотрудника' };
  }

  var sheet = getOrCreateSheet_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  var id = body.employeeId || Utilities.getUuid();
  var now = nowIso_();
  var rowNum = findRowById_(sheet, 'employeeId', id);
  var existing = rowNum ? sheetToObjects_(sheet).find(function(e) { return String(e.employeeId) === String(id); }) : null;

  var record = {
    employeeId: id,
    fullName: String(body.fullName).trim(),
    staffId: body.staffId || '',
    department: body.department || '',
    position: body.position || '',
    photo: truncatePhoto_(body.photo || (existing && existing.photo) || ''),
    faceDescriptor: body.faceDescriptor || (existing && existing.faceDescriptor) || '',
    active: body.active !== false,
    createdAt: (existing && existing.createdAt) || now,
    updatedAt: now
  };

  var row = objectToRow_(EMPLOYEE_HEADERS, record);
  if (rowNum) {
    sheet.getRange(rowNum, 1, 1, EMPLOYEE_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  logEvent_({
    type: rowNum ? 'employee_update' : 'employee_create',
    employeeId: id,
    employeeName: record.fullName,
    status: 'ok',
    message: rowNum ? 'Сотрудник обновлён' : 'Сотрудник создан'
  });

  return { ok: true, employeeId: id, employee: record };
}

function saveMeal(body) {
  if (!body || !body.employeeId || !body.mealType) {
    return { ok: false, status: 'error', message: 'Неполные данные для регистрации питания' };
  }

  var sheet = getOrCreateSheet_(SHEET_MEALS, MEAL_HEADERS);
  var date = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var meals = sheetToObjects_(sheet);

  var duplicate = meals.some(function(m) {
    return String(m.employeeId) === String(body.employeeId) &&
      String(m.mealType) === String(body.mealType) &&
      mealDate_(m) === date;
  });

  if (duplicate) {
    return {
      ok: false,
      status: 'error',
      message: body.mealType + ' уже зарегистрирован сегодня для этого сотрудника'
    };
  }

  var record = {
    mealId: body.mealId || Utilities.getUuid(),
    timestamp: body.timestamp || nowIso_(),
    date: date,
    time: body.time || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'),
    employeeId: body.employeeId,
    employeeName: body.employeeName || '',
    staffId: body.staffId || '',
    department: body.department || '',
    mealType: body.mealType,
    siteName: body.siteName || '',
    operator: body.operator || '',
    matchScore: body.matchScore !== undefined ? body.matchScore : '',
    photo: truncatePhoto_(body.photo || ''),
    verified: body.verified !== false,
    note: body.note || ''
  };

  sheet.appendRow(objectToRow_(MEAL_HEADERS, record));

  logEvent_({
    type: 'meal',
    employeeId: record.employeeId,
    employeeName: record.employeeName,
    status: 'ok',
    message: record.mealType + ' зарегистрирован',
    photo: record.photo,
    matchScore: record.matchScore
  });

  return { ok: true, mealId: record.mealId, meal: record };
}

function logEvent(body) {
  logEvent_(body || {});
  return { ok: true };
}

function logEvent_(body) {
  try {
    var sheet = getOrCreateSheet_(SHEET_LOGS, LOG_HEADERS);
    sheet.appendRow(objectToRow_(LOG_HEADERS, {
      eventId: Utilities.getUuid(),
      timestamp: nowIso_(),
      type: body.type || 'info',
      employeeId: body.employeeId || '',
      employeeName: body.employeeName || '',
      status: body.status || '',
      message: body.message || '',
      photo: truncatePhoto_(body.photo || ''),
      matchScore: body.matchScore !== undefined ? body.matchScore : ''
    }));
  } catch (e) {
    // журнал не должен ломать основной поток
  }
}

function mealDate_(meal) {
  if (meal.date) return String(meal.date).slice(0, 10);
  if (meal.timestamp) return String(meal.timestamp).slice(0, 10);
  return '';
}

function truncatePhoto_(photo) {
  if (!photo) return '';
  var s = String(photo);
  if (s.length <= 45000) return s;
  return s.slice(0, 45000);
}

/** Запустите один раз из редактора Apps Script для создания листов */
function setupSheets() {
  getOrCreateSheet_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  getOrCreateSheet_(SHEET_MEALS, MEAL_HEADERS);
  getOrCreateSheet_(SHEET_LOGS, LOG_HEADERS);
  Logger.log('Листы Employees, Meals, Logs готовы.');
}
