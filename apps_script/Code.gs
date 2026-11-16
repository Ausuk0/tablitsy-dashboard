/**
 * Скрипт для Google Таблицы «Дашборд продаж».
 *
 * Что делает:
 *   1. Добавляет в таблицу меню «Отчёты»:
 *        • Обновить сводку — пересчёт, подсветка просрочек, отметка времени обновления;
 *        • Отправить сводку на почту — письмо с главными цифрами выбранного месяца;
 *        • Подсветить просроченные заявки;
 *        • Перенести выполненные в архив;
 *        • Включить / выключить автоматизацию (триггеры по расписанию).
 *   2. Подсвечивает строки на листе «Заявки»: просрочено — красным, срок сегодня — жёлтым,
 *      выполнено — зелёным. При ручной правке строки подсветка обновляется сразу.
 *   3. По триггеру (каждый вечер) переносит строки со статусом «Выполнено»
 *      с листа «Заявки» на лист «Архив» и ставит дату переноса.
 *
 * КАК УСТАНОВИТЬ (один раз, 3 минуты):
 *   1. Откройте Google Таблицу → меню «Расширения» → «Apps Script».
 *   2. Удалите всё, что есть в файле Code.gs (Код.gs), и вставьте этот текст целиком.
 *   3. Нажмите «Сохранить» (значок дискеты или Ctrl+S).
 *   4. В списке функций сверху выберите onOpen и нажмите «Выполнить».
 *      Google попросит разрешение: «Проверить разрешения» → выберите свой аккаунт →
 *      «Дополнительные настройки» → «Перейти на страницу …» → «Разрешить».
 *      Это стандартное окно для любого своего скрипта: он работает только с этой таблицей
 *      и отправляет письма от вашего имени.
 *   5. Вернитесь в таблицу и обновите страницу (F5) — появится меню «Отчёты».
 *   6. Чтобы перенос в архив и подсветка работали сами каждый день:
 *      «Отчёты» → «Включить автоматизацию».
 *
 * НАСТРОЙКА: всё, что может понадобиться поменять, собрано в блоке ниже.
 */

// ===================== НАСТРОЙКИ =====================

/** Кому отправлять сводку. Пусто — письмо уйдёт владельцу таблицы (тому, кто запустил скрипт).
 *  Несколько адресов — через запятую: 'boss@example.ru, sales@example.ru' */
const EMAIL_TO = '';

/** Отправлять сводку автоматически каждый понедельник в 9:00 (после «Включить автоматизацию»). */
const SEND_WEEKLY_EMAIL = false;

/** Во сколько часов (0–23) каждый день переносить выполненные заявки в архив. */
const ARCHIVE_HOUR = 20;

/** Названия листов — должны совпадать с ярлыками внизу таблицы. */
const SHEET_DASH = 'Дашборд';
const SHEET_ORDERS = 'Заявки';
const SHEET_ARCHIVE = 'Архив';

/** Лист «Заявки»: номера колонок (A = 1, B = 2, …). */
const COL_NUM = 2;       // B — № заявки
const COL_DEADLINE = 7;  // G — Срок
const COL_STATUS = 8;    // H — Статус
const ORDERS_COLS = 9;   // сколько колонок занимает заявка (A:I)
const HEADER_ROWS = 1;   // сколько строк занимает шапка

/** Какой статус считается выполненным. */
const STATUS_DONE = 'Выполнено';

/** Ячейки листа «Дашборд», из которых берутся цифры для письма и уведомления. */
const CELL_MONTH = 'C4';
const CELL_UPDATED = 'I4';
const CELL_OVERDUE = 'L4';
const CELL_REVENUE = 'B7';
const CELL_ORDERS = 'D7';
const CELL_AVG = 'H7';
const CELL_PLAN_PCT = 'J7';
const CELL_PLAN_SUM = 'L8';
const RANGE_TOP_PRODUCTS = 'B12:F16';  // место, товар, кол-во, выручка, доля
const RANGE_MANAGERS = 'H12:L17';      // менеджер, план, факт, % плана, отклонение

/** Цвета подсветки строк на листе «Заявки». */
const COLOR_OVERDUE = '#f6cfca';  // просрочено
const COLOR_TODAY = '#ffefb0';    // срок сегодня
const COLOR_DONE = '#d8eedf';     // выполнено (ждёт переноса в архив)

// ===================== МЕНЮ =====================

/** Запускается сама при открытии таблицы и добавляет меню «Отчёты». */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Отчёты')
    .addItem('Обновить сводку', 'updateSummary')
    .addItem('Отправить сводку на почту', 'sendSummaryEmail')
    .addSeparator()
    .addItem('Подсветить просроченные заявки', 'highlightOverdue')
    .addItem('Перенести выполненные в архив', 'archiveCompletedFromMenu')
    .addSeparator()
    .addItem('Включить автоматизацию', 'setupTriggers')
    .addItem('Выключить автоматизацию', 'removeTriggers')
    .addToUi();
}

/**
 * Запускается сама при любой ручной правке. Если правили лист «Заявки» —
 * перекрашиваем только затронутые строки (это быстро).
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_ORDERS) return;
  const first = Math.max(e.range.getRow(), HEADER_ROWS + 1);
  const last = e.range.getLastRow();
  if (last < first) return;
  paintRows_(sheet, first, last - first + 1);
}

// ===================== ОБНОВИТЬ СВОДКУ =====================

/**
 * «Отчёты» → «Обновить сводку».
 * Формулы в Google Таблицах пересчитываются сами; эта команда дополнительно
 * обновляет подсветку просрочек, ставит время обновления и показывает
 * короткую сводку во всплывающем уведомлении.
 */
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getSheet_(ss, SHEET_DASH);
  highlightOverdue();
  SpreadsheetApp.flush(); // дождаться пересчёта формул
  dash.getRange(CELL_UPDATED).setValue(
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm'));
  ss.toast(
    'Выручка: ' + dash.getRange(CELL_REVENUE).getDisplayValue() +
    '\nЗаказов: ' + dash.getRange(CELL_ORDERS).getDisplayValue() +
    '\nПлан выполнен на ' + dash.getRange(CELL_PLAN_PCT).getDisplayValue() +
    '\nПросрочено заявок: ' + dash.getRange(CELL_OVERDUE).getDisplayValue(),
    'Сводка за ' + dash.getRange(CELL_MONTH).getDisplayValue(), 8);
}

// ===================== ПИСЬМО СО СВОДКОЙ =====================

/** «Отчёты» → «Отправить сводку на почту»: спрашивает подтверждение и отправляет. */
function sendSummaryEmail() {
  const ui = SpreadsheetApp.getUi();
  const to = getRecipients_();
  const answer = ui.alert('Отправить сводку',
    'Отправить сводку за «' + getDashValue_(CELL_MONTH) + '» на адрес ' + to + '?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const sent = sendSummary_(to);
  SpreadsheetApp.getActiveSpreadsheet().toast('Письмо отправлено: ' + sent, 'Готово', 5);
}

/** То же самое без вопросов — для еженедельного триггера. */
function sendSummaryEmailAuto() {
  sendSummary_(getRecipients_());
}

function sendSummary_(to) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getSheet_(ss, SHEET_DASH);
  SpreadsheetApp.flush();
  const month = dash.getRange(CELL_MONTH).getDisplayValue();
  const kpi = [
    ['Выручка', dash.getRange(CELL_REVENUE).getDisplayValue()],
    ['Заказов', dash.getRange(CELL_ORDERS).getDisplayValue()],
    ['Средний чек', dash.getRange(CELL_AVG).getDisplayValue()],
    ['Выполнение плана', dash.getRange(CELL_PLAN_PCT).getDisplayValue() +
      ' (план ' + dash.getRange(CELL_PLAN_SUM).getDisplayValue() + ')'],
    ['Просрочено заявок', dash.getRange(CELL_OVERDUE).getDisplayValue()]
  ];
  // getDisplayValues — значения ровно так, как они видны в таблице (с ₽ и %)
  const products = dash.getRange(RANGE_TOP_PRODUCTS).getDisplayValues().filter(function (r) { return r[1] !== ''; });
  const managers = dash.getRange(RANGE_MANAGERS).getDisplayValues().filter(function (r) { return r[0] !== ''; });

  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1d2b26">' +
    '<h2 style="color:#1e6b57;margin:0 0 12px">Сводка продаж: ' + escapeHtml_(month) + '</h2>' +
    htmlTable_(['Показатель', 'Значение'], kpi) +
    '<h3 style="color:#1e6b57;margin:18px 0 8px">Топ-5 товаров</h3>' +
    htmlTable_(['Место', 'Товар', 'Кол-во', 'Выручка', 'Доля'], products) +
    '<h3 style="color:#1e6b57;margin:18px 0 8px">Менеджеры: план / факт</h3>' +
    htmlTable_(['Менеджер', 'План', 'Факт', '% плана', 'Отклонение'], managers) +
    '<p style="color:#6b7b75;font-size:12px;margin-top:16px">Письмо сформировано автоматически из таблицы «' +
    escapeHtml_(ss.getName()) + '». <a href="' + ss.getUrl() + '">Открыть таблицу</a></p></div>';

  const text = 'Сводка продаж: ' + month + '\n' +
    kpi.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n');

  MailApp.sendEmail({
    to: to,
    subject: 'Сводка продаж: ' + month,
    body: text,          // для почтовых программ без HTML
    htmlBody: html
  });
  return to;
}

// ===================== ПОДСВЕТКА ПРОСРОЧЕК =====================

/** Перекрашивает все строки листа «Заявки». Возвращает число просроченных. */
function highlightOverdue() {
  const sheet = getSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROWS) return 0;
  return paintRows_(sheet, HEADER_ROWS + 1, lastRow - HEADER_ROWS);
}

/**
 * Красит строки [firstRow; firstRow + count) по сроку и статусу.
 * Цвета ставятся одним вызовом setBackgrounds — так быстро даже на тысячах строк.
 */
function paintRows_(sheet, firstRow, count) {
  const range = sheet.getRange(firstRow, 1, count, ORDERS_COLS);
  const values = range.getValues();
  const today = startOfDay_(new Date());
  let overdue = 0;
  const colors = values.map(function (row) {
    const empty = row.every(function (v) { return v === '' || v === null; });
    let color = null; // null = убрать заливку
    if (!empty) {
      const status = String(row[COL_STATUS - 1]).trim();
      const deadline = toDate_(row[COL_DEADLINE - 1]);
      if (status === STATUS_DONE) {
        color = COLOR_DONE;
      } else if (deadline && deadline < today) {
        color = COLOR_OVERDUE;
        overdue++;
      } else if (deadline && deadline.getTime() === today.getTime()) {
        color = COLOR_TODAY;
      }
    }
    return new Array(ORDERS_COLS).fill(color);
  });
  range.setBackgrounds(colors);
  return overdue;
}

// ===================== ПЕРЕНОС В АРХИВ =====================

/** «Отчёты» → «Перенести выполненные в архив»: переносит и сообщает, сколько строк. */
function archiveCompletedFromMenu() {
  const moved = archiveCompleted();
  SpreadsheetApp.getUi().alert(moved
    ? 'Перенесено в «' + SHEET_ARCHIVE + '»: ' + moved + ' ' + plural_(moved, 'строка', 'строки', 'строк') + '.'
    : 'Нет заявок со статусом «' + STATUS_DONE + '».');
}

/**
 * Переносит все строки со статусом «Выполнено» с листа «Заявки» в конец листа «Архив»
 * и удаляет их из «Заявок». В архиве в последнюю колонку пишется дата переноса.
 * Эту функцию вызывает ежедневный триггер. Возвращает число перенесённых строк.
 */
function archiveCompleted() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return 0; // кто-то уже переносит — не мешаем
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_(ss, SHEET_ORDERS);
    const archive = getOrCreateArchive_(ss, sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROWS) return 0;

    const values = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, ORDERS_COLS).getValues();
    const now = new Date();
    const toArchive = [];
    const rowsToDelete = [];
    values.forEach(function (row, i) {
      if (String(row[COL_STATUS - 1]).trim() === STATUS_DONE) {
        toArchive.push(row.concat([now]));
        rowsToDelete.push(HEADER_ROWS + 1 + i);
      }
    });
    if (!toArchive.length) return 0;

    // 1) дописываем в архив одним блоком
    const start = archive.getLastRow() + 1;
    archive.getRange(start, 1, toArchive.length, ORDERS_COLS + 1).setValues(toArchive);
    archive.getRange(start, ORDERS_COLS + 1, toArchive.length, 1).setNumberFormat('dd.MM.yyyy HH:mm');

    // 2) Google не даёт удалить все незакреплённые строки листа — оставим одну пустую про запас
    if (sheet.getMaxRows() - rowsToDelete.length <= sheet.getFrozenRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }
    // 3) удаляем снизу вверх, соседние строки — пачками
    let i = rowsToDelete.length - 1;
    while (i >= 0) {
      let j = i;
      while (j > 0 && rowsToDelete[j - 1] === rowsToDelete[j] - 1) j--;
      sheet.deleteRows(rowsToDelete[j], i - j + 1);
      i = j - 1;
    }
    return toArchive.length;
  } finally {
    lock.releaseLock();
  }
}

// ===================== ТРИГГЕРЫ (АВТОМАТИЗАЦИЯ) =====================

/** Имена функций, которые запускаются по расписанию. */
const TRIGGER_HANDLERS = ['archiveCompleted', 'highlightOverdue', 'sendSummaryEmailAuto'];

/**
 * «Отчёты» → «Включить автоматизацию».
 * Каждый день в ARCHIVE_HOUR — перенос выполненных в архив,
 * каждое утро в 7:00 — подсветка просрочек (даты «сдвигаются» сами по себе),
 * по понедельникам в 9:00 — письмо со сводкой (если SEND_WEEKLY_EMAIL = true).
 * Повторное включение не плодит дубликаты: старые триггеры сначала удаляются.
 */
function setupTriggers() {
  deleteOurTriggers_();
  ScriptApp.newTrigger('archiveCompleted').timeBased().everyDays(1).atHour(ARCHIVE_HOUR).create();
  ScriptApp.newTrigger('highlightOverdue').timeBased().everyDays(1).atHour(7).create();
  if (SEND_WEEKLY_EMAIL) {
    ScriptApp.newTrigger('sendSummaryEmailAuto').timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  }
  SpreadsheetApp.getUi().alert('Автоматизация включена:\n' +
    '• каждый день около ' + ARCHIVE_HOUR + ':00 — перенос выполненных заявок в архив;\n' +
    '• каждое утро около 7:00 — подсветка просроченных заявок' +
    (SEND_WEEKLY_EMAIL ? ';\n• по понедельникам около 9:00 — сводка на почту.' : '.'));
}

/** «Отчёты» → «Выключить автоматизацию». */
function removeTriggers() {
  const n = deleteOurTriggers_();
  SpreadsheetApp.getUi().alert(n ? 'Автоматизация выключена.' : 'Автоматизация и так была выключена.');
}

function deleteOurTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

// ===================== ВСПОМОГАТЕЛЬНОЕ =====================

function getSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Не найден лист «' + name + '». Проверьте название ярлыка внизу таблицы.');
  return sheet;
}

/** Лист «Архив»; если его нет — создаётся с шапкой как у «Заявок» + «Перенесено в архив». */
function getOrCreateArchive_(ss, ordersSheet) {
  let archive = ss.getSheetByName(SHEET_ARCHIVE);
  if (!archive) {
    archive = ss.insertSheet(SHEET_ARCHIVE);
    const head = ordersSheet.getRange(1, 1, HEADER_ROWS, ORDERS_COLS).getValues();
    head[HEADER_ROWS - 1].push('Перенесено в архив');
    for (let r = 0; r < HEADER_ROWS - 1; r++) head[r].push('');
    archive.getRange(1, 1, HEADER_ROWS, ORDERS_COLS + 1).setValues(head).setFontWeight('bold');
    archive.setFrozenRows(HEADER_ROWS);
  }
  return archive;
}

function getDashValue_(a1) {
  return getSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_DASH).getRange(a1).getDisplayValue();
}

function getRecipients_() {
  return EMAIL_TO.trim() || Session.getEffectiveUser().getEmail();
}

/** Ячейка срока может быть датой или текстом «16.11.2026» — приводим к дате без времени. */
function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return startOfDay_(v);
  const m = String(v).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function htmlTable_(head, rows) {
  const th = head.map(function (h) {
    return '<th style="background:#1e6b57;color:#fff;padding:6px 10px;text-align:left">' + escapeHtml_(h) + '</th>';
  }).join('');
  const tr = rows.map(function (r, i) {
    const bg = i % 2 ? '#f2f8f5' : '#ffffff';
    return '<tr>' + r.map(function (c) {
      return '<td style="padding:6px 10px;border-bottom:1px solid #dbe7e1;background:' + bg + '">' + escapeHtml_(c) + '</td>';
    }).join('') + '</tr>';
  }).join('');
  return '<table style="border-collapse:collapse;min-width:420px">' + '<tr>' + th + '</tr>' + tr + '</table>';
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function plural_(n, one, few, many) {
  const n10 = n % 10, n100 = n % 100;
  if (n10 === 1 && n100 !== 11) return one;
  if (n10 >= 2 && n10 <= 4 && (n100 < 12 || n100 > 14)) return few;
  return many;
}
