/**
 * Приёмник заявок с лендинга OnlyStaff — пишет каждую заявку строкой в Google Sheet.
 *
 * УСТАНОВКА:
 * 1. Создайте новую Google Таблицу.
 * 2. Расширения → Apps Script, вставьте туда содержимое этого файла.
 * 3. (Опционально) впишите свой e-mail в NOTIFY_EMAIL — тогда на каждую заявку
 *    будет приходить письмо, помимо строки в таблице.
 * 4. Развернуть → Новое развёртывание → тип "Веб-приложение":
 *      - Выполнять от имени: Меня
 *      - У кого есть доступ: Все (Anyone)
 * 5. Скопируйте URL (заканчивается на /exec) → вставьте в index.html → CONFIG.formEndpoint.
 *
 * При изменении кода нужно новое развёртывание (Manage deployments → Edit → New version).
 */

var SHEET_NAME = 'Заявки';

var NOTIFY_EMAIL = '';

var COLUMNS = [
  'submitted_at', 'lang', 'name', 'company', 'phone', 'email',
  'country', 'city', 'roles', 'headcount', 'start',
  'languages', 'car_required', 'housing', 'notes',
  'page', 'referrer', 'attribution'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet_();

    var row = COLUMNS.map(function (key) { return formatField_(key, data); });
    sheet.appendRow(row);
    notify_(data);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function formatField_(key, data) {
  if (key === 'submitted_at') return data.submitted_at ? new Date(data.submitted_at) : new Date();

  if (key === 'roles') {
    if (!Array.isArray(data.roles)) return '';
    return data.roles
      .map(function (r) { return (r.label || r.specialty || '?') + ' ×' + (r.qty || 1); })
      .join(', ');
  }

  if (key === 'languages') {
    return Array.isArray(data.languages) ? data.languages.join(', ') : '';
  }

  if (key === 'attribution') {
    try { return JSON.stringify(data.attribution || {}); } catch (e) { return ''; }
  }

  return data[key] != null ? data[key] : '';
}

function notify_(data) {
  if (!NOTIFY_EMAIL) return;
  var roles = Array.isArray(data.roles)
    ? data.roles.map(function (r) { return (r.label || r.specialty || '?') + ' ×' + (r.qty || 1); }).join(', ')
    : '';
  var body =
    'Компания: ' + (data.company || '') + '\n' +
    'Контакт: ' + (data.name || '') + '\n' +
    'Телефон: ' + (data.phone || '') + '\n' +
    (data.email ? 'E-mail: ' + data.email + '\n' : '') +
    'Страна/город: ' + (data.country || '') + ' / ' + (data.city || '') + '\n' +
    'Нужны: ' + roles + ' (всего ' + (data.headcount || '') + ')\n' +
    'Начало: ' + (data.start || '') + '\n' +
    (data.notes ? 'Комментарий: ' + data.notes + '\n' : '');
  MailApp.sendEmail(NOTIFY_EMAIL, 'Новая заявка OnlyStaff — ' + (data.company || data.name || ''), body);
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
