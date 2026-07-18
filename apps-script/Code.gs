/**
 * Питание Персонала — бэкенд на Google Apps Script
 *
 * Установка (один раз):
 *  1. Создайте новую Google Таблицу.
 *  2. Расширения → Apps Script → удалите Code.gs по умолчанию,
 *     вставьте этот файл и сохраните (значок дискеты).
 *  3. Запустите функцию setupSheets (выберите её в выпадающем
 *     списке наверху, нажмите ▶ «Выполнить»). Разрешите доступ.
 *  4. Развернуть → Новое развертывание → тип «Веб-приложение»:
 *        Выполнять от имени: «Я»
 *        Доступ:           «Все»
 *     Нажмите «Развернуть» и скопируйте URL, оканчивающийся на /exec.
 *  5. Вставьте этот URL в файле public/kiosk.html в константу
 *     API_URL (вверху <script>). Опубликуйте сайт заново.
 *
 * После этого приложение всегда работает с этим бэкендом без
 * каких-либо ручных настроек URL в интерфейсе.
 *
 * ─────────────────────────────────────────────────────────────────
 * ИСПРАВЛЕНИЯ v2.1:
 *  1. Колонки date/time листа Meals принудительно форматируются как
 *     ОБЫЧНЫЙ ТЕКСТ ('@') перед каждой записью. Раньше Google Sheets
 *     автоматически распознавал строку "2026-07-18" как дату и
 *     конвертировал ячейку в тип Date. При чтении через getValues()
 *     это превращалось в JS Date-объект и сериализовалось в JSON как
 *     полноценный ISO-timestamp (например "2026-07-18T00:00:00.000Z"),
 *     из-за чего фронтенд при СТРОКОВОМ сравнении дат в отчётах и
 *     журнале терял часть (а иногда и все) строки. Теперь date/time
 *     остаются простыми текстовыми строками "YYYY-MM-DD"/"HH:mm:ss".
 *  2. Добавлена колонка extraMeal в лист Meals (раньше фронт её
 *     отправлял, а бэкенд тихо отбрасывал, т.к. её не было в
 *     MEAL_HEADERS).
 *  3. Доп. питание (extraMeal=true) больше не блокируется общим
 *     cooldown между обычными приёмами пищи.
 *  4. Добавлены action=saveSettings/getSettings (лист Settings) —
 *     фронтенд их уже вызывает (сохранение названия точки/оператора
 *     на сервере), но в исходном Code.gs обработчиков не было, и
 *     запросы просто падали с "Неизвестное действие".
 *  5. Если у вас уже есть данные, накопленные СТАРОЙ версией скрипта
 *     (колонка date уже "испорчена" как тип Date) — запустите один
 *     раз вручную функцию fixMealsDateColumn() из редактора Apps
 *     Script, она пересчитает date/time из timestamp и переформатирует
 *     колонки в текст.
 * ─────────────────────────────────────────────────────────────────
 */

var SHEET_EMPLOYEES = 'Employees';
var SHEET_MEALS     = 'Meals';
var SHEET_LOGS      = 'Logs';
var SHEET_SETTINGS  = 'Settings';

// Минимальный интервал между двумя выдачами ОБЫЧНОГО питания одному
// сотруднику (минуты). Защищает от повторного получения
// независимо от того, Завтрак/Обед/Ужин. НЕ применяется к
// доп. питанию (extraMeal=true) — см. saveMeal().
var COOLDOWN_MINUTES = 60;

var EMPLOYEE_HEADERS = [
  'employeeId', 'fullName', 'staffId', 'department', 'position',
  'photo', 'faceDescriptor', 'active', 'createdAt', 'updatedAt'
];

// Добавлена колонка 'extraMeal' (была потеряна в исходной версии).
var MEAL_HEADERS = [
  'mealId', 'timestamp', 'date', 'time', 'employeeId', 'employeeName',
  'staffId', 'department', 'mealType', 'extraMeal', 'siteName', 'operator',
  'matchScore', 'photo', 'verified', 'note'
];

var LOG_HEADERS = [
  'eventId', 'timestamp', 'type', 'employeeId', 'employeeName',
  'status', 'message', 'photo', 'matchScore'
];

var SETTINGS_HEADERS = ['key', 'value', 'updatedAt'];

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    switch (action) {
      case 'ping':
        return respond({ ok: true, message: 'pong', version: '2.1', cooldownMinutes: COOLDOWN_MINUTES });
      case 'listEmployees':
        return respond({ ok: true, items: listEmployees() });
      case 'listMeals':
        return respond({ ok: true, items: listMeals(Number(e.parameter.limit) || 500) });
      case 'getSettings':
        return respond({ ok: true, settings: getSettings() });
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
      case 'saveEmployee': return respond(saveEmployee(body));
      case 'saveMeal':     return respond(saveMeal(body));
      case 'logEvent':     return respond(logEvent(body));
      case 'saveSettings': return respond(saveSettings(body));
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

function getSpreadsheet_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getOrCreateSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // Для листа Meals принудительно держим колонки date/time текстовыми,
  // чтобы Google Sheets не авто-конвертировал их в тип Date при записи.
  // Это выполняется при каждом обращении к листу (дёшево и идемпотентно),
  // поэтому защищает и вновь создаваемые, и уже существующие листы.
  if (name === SHEET_MEALS) {
    ensureTextColumns_(sheet, headers, ['date', 'time']);
  }
  return sheet;
}

function ensureTextColumns_(sheet, headers, colNames) {
  var maxRows = Math.max(sheet.getMaxRows(), 2);
  colNames.forEach(function (colName) {
    var idx = headers.indexOf(colName);
    if (idx < 0) return;
    var range = sheet.getRange(2, idx + 1, maxRows - 1, 1);
    if (range.getNumberFormat() !== '@') {
      range.setNumberFormat('@');
    }
  });
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
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = normalizeCell_(row[j]);
    rows.push(obj);
  }
  return rows;
}

// Если ячейка всё же оказалась типом Date (например, старые строки,
// записанные до применения текстового формата колонки), приводим её
// к строке "YYYY-MM-DD" вместо полного ISO-timestamp с "T", чтобы не
// ломать сравнения на фронтенде.
function normalizeCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
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

function nowIso_() { return new Date().toISOString(); }

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
    fullName: String(body.fullName).trim().slice(0, 120),
    staffId: String(body.staffId || '').slice(0, 40),
    department: String(body.department || '').slice(0, 80),
    position: String(body.position || '').slice(0, 80),
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
    employeeId: id, employeeName: record.fullName,
    status: 'ok',
    message: rowNum ? 'Сотрудник обновлён' : 'Сотрудник создан'
  });

  return { ok: true, employeeId: id, employee: record };
}

/**
 * Регистрация выдачи питания.
 * Сервер — последняя линия защиты: даже если клиент обойдёт UI,
 * выдача чаще чем раз в COOLDOWN_MINUTES будет отклонена — но ТОЛЬКО
 * для обычного питания (завтрак/обед/ужин). Доп. питание
 * (body.extraMeal === true) не подчиняется этому кулдауну: это
 * отдельная позиция (перекус, кефир и т.п.), которую можно выдавать
 * независимо от основного приёма пищи.
 */
function saveMeal(body) {
  if (!body || !body.employeeId || !body.mealType) {
    return { ok: false, status: 'error', message: 'Неполные данные для регистрации питания' };
  }

  var isExtra = body.extraMeal === true || String(body.extraMeal).toLowerCase() === 'true';

  // Глобальная блокировка — нельзя одновременно записывать
  // две выдачи (защищает от двойного нажатия).
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) {
    return { ok: false, status: 'error', message: 'Сервер занят, повторите через секунду' };
  }

  try {
    var sheet = getOrCreateSheet_(SHEET_MEALS, MEAL_HEADERS);
    var meals = sheetToObjects_(sheet);
    var now = new Date();
    var nowMs = now.getTime();
    var cooldownMs = COOLDOWN_MINUTES * 60 * 1000;

    if (!isExtra) {
      // Ищем последнюю ОБЫЧНУЮ выдачу этому сотруднику (доп. питание
      // в расчёт cooldown не берём и им не блокируем).
      var lastTs = 0, lastMeal = null;
      for (var i = 0; i < meals.length; i++) {
        if (String(meals[i].employeeId) !== String(body.employeeId)) continue;
        var rowIsExtra = meals[i].extraMeal === true || String(meals[i].extraMeal).toLowerCase() === 'true' || meals[i].mealType === 'Доп. питание';
        if (rowIsExtra) continue;
        var t = Date.parse(meals[i].timestamp);
        if (!isNaN(t) && t > lastTs) { lastTs = t; lastMeal = meals[i]; }
      }

      if (lastTs && (nowMs - lastTs) < cooldownMs) {
        var leftMin = Math.ceil((cooldownMs - (nowMs - lastTs)) / 60000);
        logEvent_({
          type: 'meal_blocked',
          employeeId: body.employeeId,
          employeeName: body.employeeName || '',
          status: 'blocked',
          message: 'Повторная попытка через ' + Math.round((nowMs - lastTs) / 60000) + ' мин (последний: ' + (lastMeal && lastMeal.mealType) + ')'
        });
        return {
          ok: false,
          status: 'cooldown',
          message: 'Питание уже выдано (' + (lastMeal && lastMeal.mealType) + '). Следующая выдача через ' + leftMin + ' мин.',
          minutesLeft: leftMin,
          lastMealType: lastMeal && lastMeal.mealType,
          lastTimestamp: lastMeal && lastMeal.timestamp
        };
      }
    }

    var record = {
      mealId: body.mealId || Utilities.getUuid(),
      timestamp: nowIso_(),
      date: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      time: Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
      employeeId: body.employeeId,
      employeeName: body.employeeName || '',
      staffId: body.staffId || '',
      department: body.department || '',
      mealType: body.mealType,
      extraMeal: isExtra,
      siteName: body.siteName || '',
      operator: body.operator || '',
      matchScore: body.matchScore !== undefined ? body.matchScore : '',
      photo: truncatePhoto_(body.photo || ''),
      verified: body.verified !== false,
      note: body.note || ''
    };

    // Колонки date/time уже отформатированы как текст в getOrCreateSheet_,
    // поэтому appendRow не даст Google Sheets конвертировать их в Date.
    sheet.appendRow(objectToRow_(MEAL_HEADERS, record));

    logEvent_({
      type: 'meal',
      employeeId: record.employeeId, employeeName: record.employeeName,
      status: 'ok',
      message: record.mealType + (isExtra ? ' (доп.)' : '') + ' зарегистрирован',
      photo: record.photo,
      matchScore: record.matchScore
    });

    return { ok: true, mealId: record.mealId, meal: record };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function logEvent(body) { logEvent_(body || {}); return { ok: true }; }

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
  } catch (e) { /* журнал не должен ломать основной поток */ }
}

/**
 * Настройки точки (название точки/оператор), хранятся как пары
 * key/value на листе Settings, чтобы не терялись при очистке
 * localStorage браузера.
 */
function getSettings() {
  var sheet = getOrCreateSheet_(SHEET_SETTINGS, SETTINGS_HEADERS);
  var rows = sheetToObjects_(sheet);
  var out = {};
  rows.forEach(function (r) { out[r.key] = r.value; });
  return out;
}

function saveSettings(body) {
  if (!body) return { ok: false, status: 'error', message: 'Нет данных' };
  var sheet = getOrCreateSheet_(SHEET_SETTINGS, SETTINGS_HEADERS);
  var now = nowIso_();
  var toSave = {};
  if (body.siteName !== undefined) toSave.siteName = body.siteName;
  if (body.operator !== undefined) toSave.operator = body.operator;

  Object.keys(toSave).forEach(function (key) {
    var rowNum = findRowById_(sheet, 'key', key);
    var row = objectToRow_(SETTINGS_HEADERS, { key: key, value: toSave[key], updatedAt: now });
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, SETTINGS_HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });

  return { ok: true, settings: getSettings() };
}

function truncatePhoto_(photo) {
  if (!photo) return '';
  var s = String(photo);
  return s.length <= 45000 ? s : s.slice(0, 45000);
}

/** Запустите один раз из редактора Apps Script для создания листов. */
function setupSheets() {
  getOrCreateSheet_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  getOrCreateSheet_(SHEET_MEALS,     MEAL_HEADERS);
  getOrCreateSheet_(SHEET_LOGS,      LOG_HEADERS);
  getOrCreateSheet_(SHEET_SETTINGS,  SETTINGS_HEADERS);
  Logger.log('Готово: листы Employees, Meals, Logs, Settings созданы.');
}

/**
 * ОДНОРАЗОВЫЙ РЕМОНТ для тех, у кого уже есть данные, накопленные
 * СТАРОЙ версией скрипта (когда колонка date хранилась как тип Date
 * из-за авто-конвертации Google Sheets).
 *
 * Запустите вручную ОДИН РАЗ из редактора Apps Script (выберите
 * fixMealsDateColumn в выпадающем списке функций и нажмите ▶).
 * Скрипт пересчитает столбцы date/time из timestamp для всех
 * существующих строк и переформатирует сами колонки в текстовый
 * формат, чтобы проблема не повторялась.
 */
function fixMealsDateColumn() {
  var sheet = getOrCreateSheet_(SHEET_MEALS, MEAL_HEADERS); // заодно проставит текстовый формат
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('Нет данных для исправления.'); return; }

  var headers = data[0].map(String);
  var tsCol   = headers.indexOf('timestamp');
  var dateCol = headers.indexOf('date');
  var timeCol = headers.indexOf('time');
  if (tsCol < 0 || dateCol < 0 || timeCol < 0) {
    Logger.log('Не найдены нужные колонки.');
    return;
  }

  var tz = Session.getScriptTimeZone();
  var fixed = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ts = row[tsCol];
    var d;
    if (Object.prototype.toString.call(ts) === '[object Date]') {
      d = ts;
    } else if (ts) {
      var parsed = new Date(ts);
      if (!isNaN(parsed.getTime())) d = parsed;
    }
    if (!d) continue;

    var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(d, tz, 'HH:mm:ss');

    sheet.getRange(i + 1, dateCol + 1).setValue(dateStr);
    sheet.getRange(i + 1, timeCol + 1).setValue(timeStr);
    fixed++;
  }
  Logger.log('Исправлено строк: ' + fixed);
}
