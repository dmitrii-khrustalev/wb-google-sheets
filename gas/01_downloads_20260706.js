// ============================================
// WB API - ПОЛНАЯ АНАЛИТИКА
// ФИНАЛЬНАЯ РАБОЧАЯ ВЕРСИЯ (БЕЗ ЗАКАЗОВ)
// Файл 00_headers_ru_20260822.js должен быть в том же проекте (грузится первым).
// ============================================

// 1. МЕНЮ ПРИ ОТКРЫТИИ
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 WB Аналитика')
      .addItem('🔑 Ввести API токен', 'setApiToken')
      .addSeparator()
      .addItem('💰 Финансовый отчёт', 'showFinanceReportDialog')
      .addItem('📢 Рекламные расходы', 'showAdvDialog')
      .addItem('📈 Воронка продаж', 'probeSalesFunnel')
      .addItem('📦 Заказы', 'showOrdersDialog')
      .addItem('📦 Товары', 'loadAllProducts')
      .addItem('💵 История цен', 'savePriceHistory')
      .addItem('📦 Остатки', 'recordStocksFast')
      .addItem('📦 Остатки ЛК', 'recordLkStocks')
      .addItem('📦 Выгрузка FBS', 'showFbsExportDialog')
      .addSeparator()
      .addSubMenu(ui.createMenu('🚚 FBW')
        .addItem('📦 Поставки FBW', 'showSuppliesDialog')
        .addItem('🔄 Обновить поставки', 'updateSuppliesDaily')
        .addItem('🔄 Обновить статусы поставок', 'updateExistingSuppliesStatuses')
        .addItem('📦 Заказы в пути', 'runManualUpdate'))
      .addSubMenu(ui.createMenu('⏰ Утро')
        .addItem('🔄 Утренняя цепочка (4 шага)', 'runMorningChain')
        .addItem('⏰ Триггер цепочки 5:00', 'setupMorningChainTrigger')
        .addItem('🔄 Обновить все отчёты ч1', 'updateAllReports1')
        .addItem('🔄 Обновить все отчёты ч2', 'updateAllReports2'))
      .addSubMenu(ui.createMenu('🔧 Сервис')
        .addItem('📈 Еженедельная динамика продаж', 'showWeekPicker')
        .addItem('🗑️ Удалить дубликаты', 'removeDuplicatesFromSheet')
        .addItem('🔄 Проверить склады', 'generateWarehouseMasterList2')
        .addItem('🔄 Заменить склады', 'fixWarehouseNames2')
        .addItem('🔄 Заменить склады последняя дата', 'fixWarehouseNamesLastDate2')
        .addItem('🔑 Ключ базы', 'setDbKey')
        .addItem('📊 Заполненность ячеек', 'countFilledCellsOnly'))
      .addSubMenu(ui.createMenu('🧪 Пробы')
        .addItem('📦 Остатки CSV (проба)', 'probeStocksHistoryCsv')
        .addItem('💰 История цен (проба)', 'probePriceHistory'))
      .addToUi();
  } catch (e) {
    Logger.log('onOpen: ' + e.toString());
  }
}

// 2. РАБОТА С ТОКЕНОМ
// Короткий ui.prompt оставляет только первую строку вставки и молча затирает рабочий токен.
function setApiToken() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:16px;margin:0}' +
    'p{font-size:12px;line-height:1.45;color:#333;margin:0 0 10px}' +
    'textarea{width:100%;height:150px;box-sizing:border-box;font-family:Consolas,monospace;font-size:12px}' +
    'button{margin-top:10px;background:#4CAF50;color:#fff;padding:10px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e}' +
    '#st{margin-top:8px;font-size:12px;color:#555}' +
    '</style></head><body>' +
    '<p>Вставьте токен WB целиком. Категории: Статистика, Продвижение, Контент, Маркетплейс, Аналитика. Сохранится только строка из трёх частей через точку. Обрезанный вариант старый токен не затирает.</p>' +
    '<textarea id="t" placeholder="eyJ..."></textarea>' +
    '<button id="btn" onclick="go()">Сохранить</button>' +
    '<div id="st"></div>' +
    '<script>' +
    'function go(){' +
    'var v=document.getElementById("t").value;' +
    'document.getElementById("btn").disabled=true;' +
    'document.getElementById("st").textContent="Проверяю…";' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){document.getElementById("st").textContent=msg;alert(msg);google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("btn").disabled=false;document.getElementById("st").textContent=String(e.message||e);})' +
    '.saveWbApiToken(v);' +
    '}' +
    '</script></body></html>'
  ).setWidth(520).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'API токен Wildberries');
}

function normalizeWbToken_(raw) {
  var token = String(raw || '').replace(/\s+/g, '');
  if (token.indexOf('Bearer') === 0) token = token.substring(6);
  return token;
}

function saveWbApiToken(raw) {
  var token = normalizeWbToken_(raw);
  var parts = token ? token.split('.') : [];
  if (token.indexOf('eyJ') !== 0 || parts.length !== 3 || parts[0].length < 10 || parts[1].length < 10 || parts[2].length < 10) {
    throw new Error('Токен не сохранён: частей ' + parts.length + ' вместо 3, длина ' + token.length + '. Вставьте строку целиком, в ней две точки. Прежний токен не менял.');
  }
  PropertiesService.getUserProperties().setProperty('WB_API_TOKEN', token);
  return 'Токен сохранён. Длина ' + token.length + ', частей 3.';
}

function getApiToken() {
  var props = PropertiesService.getUserProperties();
  var token = normalizeWbToken_(props.getProperty('WB_API_TOKEN'));
  if (!token) {
    try {
      SpreadsheetApp.getUi().alert('⚠️ Токен не найден! Сначала введите его через меню.');
    } catch (e) {}
  }
  return token;
}

function isRunningFromTrigger() {
  try {
    SpreadsheetApp.getUi();
    return false;
  } catch(e) {
    return true;
  }
}

function setTelegramBot() {
  var ui = SpreadsheetApp.getUi();
  var t = ui.prompt(
    'Тот же бот, что /daily',
    'Токен из Supabase secrets: TELEGRAM_BOT_TOKEN (от @BotFather)',
    ui.ButtonSet.OK_CANCEL
  );
  if (t.getSelectedButton() !== ui.Button.OK) return;
  var token = String(t.getResponseText() || '').trim();
  if (!token) {
    ui.alert('Токен пустой — не сохранил');
    return;
  }
  var c = ui.prompt(
    'Тот же чат, что /daily',
    'TELEGRAM_CHAT_ID из секретов Supabase (число)',
    ui.ButtonSet.OK_CANCEL
  );
  if (c.getSelectedButton() !== ui.Button.OK) return;
  var chatId = String(c.getResponseText() || '').trim();
  if (!chatId) {
    ui.alert('chat id пустой — не сохранил');
    return;
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty('TG_BOT_TOKEN', token);
  props.setProperty('TG_CHAT_ID', chatId);
  var ok = tgSend_('Google Sheets: бот подключён. После каждого скрипта придёт отдельное сообщение.');
  ui.alert(ok
    ? 'Сохранил, тестовое сообщение ушло в тот же чат, что /daily'
    : 'Сохранил, но отправить не удалось. Проверьте токен и chat id.');
}

function testTelegramBot() {
  var ok = tgSendJob_('test', true, [
    'дата: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'тест: одно сообщение на скрипт, без ожидания остальных'
  ], Date.now());
  try {
    SpreadsheetApp.getUi().alert(ok
      ? 'Ушло одно тестовое сообщение в бот'
      : 'Не отправилось. Сначала меню «Telegram бот».');
  } catch (e) {
    Logger.log('testTelegramBot: ' + (ok ? 'ok' : 'fail'));
  }
}

function tgSend_(text) {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('TG_BOT_TOKEN');
    var chatId = props.getProperty('TG_CHAT_ID');
    if (!token || !chatId) {
      Logger.log('Telegram не настроен (меню → Telegram бот)');
      return false;
    }
    var msg = String(text || '');
    if (msg.length > 4000) msg = msg.substring(0, 3990) + '\n…';
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: msg,
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code !== 200) Logger.log('Telegram ' + code + ': ' + res.getContentText());
    return code === 200;
  } catch (e) {
    Logger.log('Telegram: ' + e);
    return false;
  }
}

function tgSheetRows_(name) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sh) return 0;
    return Math.max(0, sh.getLastRow() - 1);
  } catch (e) {
    return 0;
  }
}

function tgWatch_(lines, sheetName, fn, label) {
  var before = tgSheetRows_(sheetName);
  var ok = false;
  try {
    fn();
    ok = true;
  } finally {
    try { SpreadsheetApp.flush(); } catch (e1) {}
    var after = tgSheetRows_(sheetName);
    var d = after - before;
    var delta = d === 0 ? '0' : (d > 0 ? '+' + d : String(d));
    var name = label || sheetName;
    lines.push(name + ': ' + after + ' строк (' + delta + ')');
    if (typeof morningTgNote_ === 'function') {
      morningTgNote_(ok, name, after + ' (' + delta + ')');
    }
  }
}

function tgRunStep_(lines, title, fn) {
  try {
    fn();
  } catch (e) {
    lines.push('❌ ' + title + ': ' + e);
    Logger.log(title + ': ' + e);
    if (typeof morningTgNote_ === 'function') morningTgNote_(false, title, String(e));
  }
}

var TG_JOB_TITLE = {
  updateAllReports1: 'updateallreports1',
  updateAllReports2: 'updateallreports2',
  updateAllReports3: 'updateallreports3',
  updateFbsDaily: 'updateFBSdaily',
  fbsStocks: 'FBS остатки',
  composeFbsReports: 'сборка FBS',
  test: 'test'
};

function tgSendJob_(jobName, ok, lines, startedMs) {
  if (typeof morningChainOn_ === 'function' && morningChainOn_()) return true;
  var tz = Session.getScriptTimeZone();
  var sec = startedMs ? Math.round((Date.now() - startedMs) / 1000) : 0;
  var when = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  var title = TG_JOB_TITLE[jobName] || jobName;
  var out = [
    (ok ? 'готово' : 'ошибка') + ' · ' + title,
    when + ' · ' + sec + ' сек'
  ];
  var body = lines || [];
  for (var i = 0; i < body.length; i++) {
    if (body[i]) out.push(body[i]);
  }
  return tgSend_(out.join('\n'));
}

var TG_NOTES = [];
function tgNote_(line) {
  if (line) TG_NOTES.push(String(line));
}
function tgTakeNotes_() {
  var n = TG_NOTES.slice();
  TG_NOTES = [];
  return n;
}

// 3. ДИАЛОГ ФИНАНСОВОГО ОТЧЁТА
function showFinanceReportDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;height:280px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input,select{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:hover{background:#45a049}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '.info{background:#e3f2fd;padding:8px;margin:10px 0;border-radius:4px;font-size:11px}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">💰 Финансовый отчёт</h3>' +
    '<div class="info">📌 Ежедневный — детально по операциям<br>📌 Еженедельный — агрегированный</div>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo">' +
    '<label>Тип отчёта:</label>' +
    '<select id="periodType">' +
    '<option value="daily">📅 Ежедневный</option>' +
    '<option value="weekly">📊 Еженедельный</option>' +
    '</select>' +
    '<button id="btn" onclick="loadReport()">📥 Загрузить отчёт</button>' +
    '<script>' +
    'function loadReport(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'var period=document.getElementById("periodType").value;' +
    'if(!d1||!d2){alert("Выберите даты!");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);})' +
    '.runFinanceReport(d1,d2,period);' +
    '}' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Финансовый отчёт');
}

// 4. ДИАЛОГ РЕКЛАМЫ
function showAdvDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;height:220px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:hover{background:#45a049}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📢 Рекламные расходы</h3>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo">' +
    '<button id="btn" onclick="loadReport()">📥 Загрузить отчёт</button>' +
    '<script>' +
    'function loadReport(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'if(!d1||!d2){alert("Выберите даты!");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);})' +
    '.runAdvReport(d1,d2);' +
    '}' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Рекламные расходы');
}

function showFbsExportDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;margin:0}' +
    'h3{margin-top:0}' +
    '.row{margin:10px 0;font-size:14px}' +
    '.row label{cursor:pointer}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px;margin-top:8px}' +
    'button:hover{background:#45a049}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '.info{background:#e3f2fd;padding:8px;margin:0 0 10px;border-radius:4px;font-size:11px}' +
    '</style></head><body>' +
    '<h3>📦 Выгрузка FBS</h3>' +
    '<div class="info">Отметьте, что загрузить. Рабочие листы соберутся в конце.</div>' +
    '<div class="row"><label><input type="checkbox" id="orders" checked> Заказы</label></div>' +
    '<div class="row"><label><input type="checkbox" id="supplies" checked> Поставки</label></div>' +
    '<div class="row"><label><input type="checkbox" id="stocks" checked> Остатки</label></div>' +
    '<button id="btn" onclick="go()">📥 Выгрузить</button>' +
    '<script>' +
    'function go(){' +
    'var sel={orders:document.getElementById("orders").checked,supplies:document.getElementById("supplies").checked,stocks:document.getElementById("stocks").checked};' +
    'if(!sel.orders&&!sel.supplies&&!sel.stocks){alert("Выберите хотя бы один пункт");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);}).runFbsExport(sel);' +
    '}' +
    '</script></body></html>';
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Выгрузка FBS');
}

function runFbsExport(sel) {
  sel = sel || {};
  var orders = !!sel.orders;
  var supplies = !!sel.supplies;
  var stocks = !!sel.stocks;
  if (!orders && !supplies && !stocks) {
    SpreadsheetApp.getUi().alert('Ничего не выбрано');
    return;
  }
  var hadSilent = (typeof WB_SILENT_UI !== 'undefined') ? WB_SILENT_UI : false;
  if (typeof WB_SILENT_UI !== 'undefined') WB_SILENT_UI = true;
  var done = [];
  var opt = { skipCompose: true };
  try {
    if (orders) {
      if (typeof loadFbsOrders !== 'function') throw new Error('Нет функции loadFbsOrders (файл 06)');
      loadFbsOrders(opt);
      done.push('заказы');
    }
    if (supplies) {
      if (typeof loadFbsSupplies !== 'function') throw new Error('Нет функции loadFbsSupplies (файл 07)');
      loadFbsSupplies(opt);
      done.push('поставки');
    }
    if (stocks) {
      if (typeof loadSellerWarehouseStocks !== 'function') throw new Error('Нет функции loadSellerWarehouseStocks (файл 05)');
      loadSellerWarehouseStocks(opt);
      done.push('остатки');
    }
    if (typeof composeFbsReports === 'function') composeFbsReports(true);
  } finally {
    if (typeof WB_SILENT_UI !== 'undefined') WB_SILENT_UI = hadSilent;
  }
  try {
    SpreadsheetApp.getUi().alert('✅ FBS выгружено: ' + done.join(', '));
  } catch (eUi) {
    Logger.log('FBS выгружено: ' + done.join(', '));
  }
}

// 5. ВЫЗОВЫ ИЗ ДИАЛОГОВ
function runFinanceReport(dateFrom, dateTo, periodType) {
  var token = getApiToken();
  if (!token) return;
  loadFinancialReportWithPeriod(token, dateFrom, dateTo, periodType);
}

function runAdvReport(dateFrom, dateTo) {
  var token = getApiToken();
  if (!token) return;
  loadAdvExpenses(token, dateFrom, dateTo);
}

// ============================================
// ФИНАНСОВЫЙ ОТЧЁТ - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ
// ============================================

// API-поля, которые реально нужны скриптам (остальные не пишем на лист)
var REQUIRED_FINANCIAL_COLUMNS = [
  'rrd_id', 'nm_id', 'sa_name', 'barcode', 'doc_type_name',
  'quantity', 'retail_amount', 'office_name', 'supplier_oper_name',
  'rr_dt', 'shk_id',
  'commission_percent', 'site_country', 'delivery_rub', 'ppvz_for_pay', 'ppvz_kvw_prc', 'ppvz_kvw_prc_base'
];

var FINANCIAL_ID_COLUMNS = ['rrd_id', 'nm_id', 'barcode', 'shk_id'];

function financeIdAsText_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return ''; // из «1970-01-01» число уже не достать
  return String(val).trim();
}

function financeKvvVal_(val) {
  if (val === null || val === undefined || val === '') return '';
  var n = Number(val);
  return isNaN(n) ? val : n;
}

function financeCountry_(val) {
  if (val === null || val === undefined || val === '') return '';
  return String(val).trim();
}

function isFinanceIdHeader_(name) {
  var h = name != null ? String(name).trim() : '';
  for (var i = 0; i < FINANCIAL_ID_COLUMNS.length; i++) {
    if (FINANCIAL_ID_COLUMNS[i] === h) return true;
    if (typeof wbIs_ === 'function' && wbIs_(h, FINANCIAL_ID_COLUMNS[i])) return true;
  }
  return false;
}

function financeV1Money_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  var n = Number(String(val).replace(',', '.').replace(/\s/g, ''));
  return isNaN(n) ? 0 : n;
}

function financeV1Rows_(data) {
  if (Array.isArray(data)) return data;
  if (data && Array.isArray(data.data)) return data.data;
  return [];
}

/** Новый finance-api → старые ключи листа / отчётов. Сырой camelCase на лист не пишем. */
function financeV1RowToV5_(row) {
  row = row || {};
  return {
    rrd_id: row.rrdId,
    nm_id: row.nmId,
    sa_name: row.vendorCode,
    barcode: row.sku,
    doc_type_name: row.docTypeName,
    quantity: row.quantity,
    retail_amount: financeV1Money_(row.retailAmount),
    office_name: row.officeName,
    supplier_oper_name: row.sellerOperName,
    rr_dt: row.rrDate,
    shk_id: row.shkId,
    commission_percent: row.commissionPercent,
    site_country: financeCountry_(row.country || row.siteCountry || row.site_country),
    ppvz_kvw_prc: row.kvw,
    ppvz_kvw_prc_base: row.kvwBase,
    retail_price_withdisc_rub: financeV1Money_(row.retailPriceWithDisc),
    ppvz_sales_commission: financeV1Money_(row.ppvzSalesCommission),
    delivery_rub: financeV1Money_(row.deliveryService),
    ppvz_for_pay: financeV1Money_(row.forPay != null ? row.forPay : row.ppvz_for_pay)
  };
}

function fetchFinanceDetailedV1_(token, dateFrom, dateTo, period) {
  var endpoint = 'https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed';
  var allData = [];
  var rrdId = 0;
  var periodType = period || 'daily';
  var page = 0;
  var limit = 100000;

  while (true) {
    page++;
    if (page > 1) {
      Logger.log('Финансы v1: пауза 61 сек перед страницей ' + page);
      Utilities.sleep(61000);
    }

    var body = {
      dateFrom: dateFrom,
      dateTo: dateTo,
      limit: limit,
      rrdId: rrdId,
      period: periodType
    };
    var res = null;
    var attempt;
    for (attempt = 1; attempt <= 3; attempt++) {
      res = wrFetchJson_(token, endpoint, 'POST', body);
      Logger.log('Финансы v1 HTTP ' + res.code + ' page ' + page + ' try ' + attempt);
      if (res.code === 429) {
        Logger.log('Финансы v1: 429, ждём 61 сек');
        Utilities.sleep(61000);
        continue;
      }
      break;
    }

    if (res.code === 204) break;
    if (res.code !== 200) {
      throw new Error('Финансы v1: HTTP ' + res.code + '\n' + String(res.text || '').substring(0, 300));
    }

    var rows = financeV1Rows_(res.data);
    if (!rows.length) break;

    for (var i = 0; i < rows.length; i++) {
      allData.push(financeV1RowToV5_(rows[i]));
    }

    var last = rows[rows.length - 1] || {};
    var nextId = last.rrdId;
    if (!nextId || nextId === rrdId || rows.length < limit) break;
    rrdId = nextId;
  }

  return allData;
}

/** Текстовый формат ID-колонок ДО setValues (иначе Sheets превращает числа в даты) */
function formatFinancialIdColumns_(sheet, startRow, numRows) {
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var lastRow = sheet.getLastRow();
  var fromRow = startRow || 2;
  if (fromRow > lastRow) return;
  var rowCount = numRows || (lastRow - fromRow + 1);
  rowCount = Math.min(Math.max(rowCount, 1), lastRow - fromRow + 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (!isFinanceIdHeader_(headers[i])) continue;
    try {
      sheet.getRange(fromRow, i + 1, rowCount, 1).setNumberFormat('@');
    } catch (e) {}
  }
}

// ОСНОВНАЯ ФУНКЦИЯ ЗАГРУЗКИ (finance-api v1)
function loadFinancialReportWithPeriod(token, dateFrom, dateTo, periodType) {
  if (typeof renameWbColumnHeaders_ === 'function') renameWbColumnHeaders_();
  var allData = [];
  var period = periodType || 'daily';
  var sheetName = (period === 'weekly') ? 'Финансовые_недельные' : 'Финансовые_ежедневные';

  // Старый GET v5 (не вызывается):
  // var endpoint = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod';
  // while (true) {
  //   var url = endpoint + '?dateFrom=' + dateFrom + '&dateTo=' + dateTo + '&period=' + period;
  //   if (nextParam) url += '&' + nextParam;
  //   var response = UrlFetchApp.fetch(url, { method: 'GET', headers: { Authorization: token }, muteHttpExceptions: true });
  //   if (response.getResponseCode() === 204) break;
  //   if (response.getResponseCode() !== 200) throw new Error('Ошибка ' + response.getResponseCode());
  //   var data = JSON.parse(response.getContentText());
  //   if (Array.isArray(data) && data.length > 0) {
  //     allData = allData.concat(filterRequiredColumns(data));
  //     var rrdid = data[data.length - 1].rrd_id;
  //     if (rrdid) { nextParam = 'rrdid=' + rrdid; Utilities.sleep(500); continue; }
  //   }
  //   break;
  // }

  try {
    allData = filterRequiredColumns(fetchFinanceDetailedV1_(token, dateFrom, dateTo, period));
  } catch (error) {
    Logger.log(error.toString());
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('Ошибка: ' + error.toString());
    return;
  }

  if (allData.length > 0) {
    // Добавляем расчетные столбцы
    var enhancedData = addFinancialCalculatedColumns(allData);
    addDataToSheetOptimized(enhancedData, sheetName, 'rrd_id');
    if (!isRunningFromTrigger()) {
      SpreadsheetApp.getUi().alert('✅ Добавлено записей: ' + allData.length + '\n📊 Столбцов: ' + Object.keys(enhancedData[0]).length);
    }
    Logger.log('Добавлено записей: ' + allData.length);
  } else {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⚠️ Нет данных за период');
    Logger.log('Нет данных за период');
  }
}

// ФИЛЬТРАЦИЯ НУЖНЫХ СТОЛБЦОВ
function filterRequiredColumns(data) {
  var filtered = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var newRow = {};
    for (var j = 0; j < REQUIRED_FINANCIAL_COLUMNS.length; j++) {
      var colName = REQUIRED_FINANCIAL_COLUMNS[j];
      if (row.hasOwnProperty(colName)) {
        newRow[colName] = row[colName];
      } else {
        newRow[colName] = '';
      }
    }
    filtered.push(newRow);
  }
  return filtered;
}

// ДОБАВЛЕНИЕ РАСЧЕТНЫХ СТОЛБЦОВ (ДОБАВЛЯЕМ ТОЛЬКО 2 НОВЫХ СТОЛБЦА)
function addFinancialCalculatedColumns(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Загружаем данные из листа "Товары"
  // Столбцы: A - nmId, B - Имя товара
  var nmIdToProductName = {};
  var nmIdToPrice = {};  // НОВЫЙ СЛОВАРЬ ДЛЯ ЦЕН
  var productsSheet = ss.getSheetByName('Товары');
  
  if (productsSheet) {
    var productsData = productsSheet.getDataRange().getValues();
    var productsHeaders = productsData[0];
    
    // Находим индекс столбца с ценой со скидкой
    var priceWithDiscountIndex = -1;
    for (var i = 0; i < productsHeaders.length; i++) {
      if (productsHeaders[i] === 'Цена со скидкой, ₽') {
        priceWithDiscountIndex = i;
        break;
      }
    }
    
    console.log('Загружено товаров из "Товары": ' + (productsData.length - 1));
    
    for (var i = 1; i < productsData.length; i++) {
      var nmId = productsData[i][0];           // столбец A - nmId
      var productName = productsData[i][1];    // столбец B - Имя товара
      var priceWithDiscount = priceWithDiscountIndex !== -1 ? parseFloat(productsData[i][priceWithDiscountIndex]) || 0 : 0;
      
      if (nmId && productName) {
        nmIdToProductName[nmId.toString()] = productName;
      }
      if (nmId && priceWithDiscount) {
        nmIdToPrice[nmId.toString()] = priceWithDiscount;
      }
    }
  } else {
    console.log('❌ Лист "Товары" не найден');
  }
  
  // Загружаем веса из "Товар+Вес" (A - sa_name, ищем столбец "Вес, кг за 1 шт")
  var weightData = {};
  var weightSheet = ss.getSheetByName('Товар+Вес');
  
  if (weightSheet) {
    var weightSheetData = weightSheet.getDataRange().getValues();
    if (weightSheetData.length > 0) {
      var headers = weightSheetData[0];
      var weightColIndex = -1;
      
      for (var i = 0; i < headers.length; i++) {
        if (headers[i] === 'Вес, кг за 1 шт') {
          weightColIndex = i;
          break;
        }
      }
      
      for (var i = 1; i < weightSheetData.length; i++) {
        var saName = weightSheetData[i][0];
        if (saName && weightColIndex !== -1) {
          var key = saName.toString().trim().toLowerCase();
          var weight = parseFloat(weightSheetData[i][weightColIndex]);
          weightData[key] = isNaN(weight) ? 0 : weight;
        }
      }
      console.log('Загружено весов из "Товар+Вес": ' + Object.keys(weightData).length);
    }
  } else {
    console.log('❌ Лист "Товар+Вес" не найден');
  }
  
  // Функция форматирования даты
  function formatDate(value) {
    if (!value) return '';
    if (value instanceof Date) {
      var d = value.getDate().toString().padStart(2, '0');
      var m = (value.getMonth() + 1).toString().padStart(2, '0');
      var y = value.getFullYear();
      return d + '.' + m + '.' + y;
    }
    var str = value.toString();
    if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
      return str.substring(8, 10) + '.' + str.substring(5, 7) + '.' + str.substring(0, 4);
    }
    return str.substring(0, 10);
  }
  
  // Функция получения веса
  function getWeight(saName) {
    if (!saName) return 0;
    var key = saName.toString().trim().toLowerCase();
    return weightData[key] || 0;
  }
  
  // Обрабатываем каждую строку
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var docTypeName = row.doc_type_name || '';

    // Исходные поля
    var nmId = row.nm_id || '';
    var saName = row.sa_name || '';
    var quantity = row.quantity || 0;
    var retailAmount = row.retail_amount || 0;
    var rrDt = row.rr_dt || '';
    
    // Расчетные столбцы
    var productName = nmIdToProductName[nmId.toString()] || '';     // Имя товара (из столбца B)
    var dateObj = null;
    if (rrDt instanceof Date && !isNaN(rrDt.getTime())) {
      dateObj = rrDt;
    } else if (rrDt) {
      var ds = rrDt.toString();
      if (/^\d{4}-\d{2}-\d{2}/.test(ds)) dateObj = new Date(ds.substring(0, 10) + 'T00:00:00');
    }
    var weightPerUnit = getWeight(saName);                           // Вес, кг за 1 шт
    var totalWeight = weightPerUnit * quantity;                      // Вес, кг за 1 шт сумма
    var salesAmount = quantity * retailAmount;                       // Сумма продаж
    
    // НОВЫЙ СТОЛБЕЦ 1: Цена со скидкой из Товары
    var priceWithDiscount = nmIdToPrice[nmId.toString()] || 0;
    
    // НОВЫЙ СТОЛБЕЦ 2: Сумма по текущей цене (с учетом знака Продажа/Возврат)
    var amountAtCurrentPrice = 0;
    if (docTypeName === 'Продажа') {
      amountAtCurrentPrice = priceWithDiscount * quantity;
    } else if (docTypeName === 'Возврат') {
      amountAtCurrentPrice = -(priceWithDiscount * quantity);
    }
    
    // Сумма с учетом возвратов (Продажа: +, Возврат: -)
    var salesWithReturns = 0;
    var quantityWithReturns = 0;
    if (docTypeName === 'Продажа') {
      salesWithReturns = retailAmount;
      quantityWithReturns = quantity;
    } else if (docTypeName === 'Возврат') {
      salesWithReturns = -retailAmount;
      quantityWithReturns = -quantity;
    }
    
    var totalWeightWithReturns = weightPerUnit * quantityWithReturns; // Вес итог с учетом возвратов
    
    // Создаем новую строку (копируем исходные поля)
    var enhanced = {};
    for (var key in row) {
      enhanced[key] = row[key];
    }

    // ID только текстом: если в строке есть Date (rr_dt/Дата), Sheets
    // превращает nm_id и barcode в 1970-01-01 (unix ms как дата)
    enhanced['rrd_id'] = financeIdAsText_(enhanced['rrd_id']);
    enhanced['nm_id'] = financeIdAsText_(enhanced['nm_id']);
    enhanced['barcode'] = financeIdAsText_(enhanced['barcode']);
    enhanced['shk_id'] = financeIdAsText_(enhanced['shk_id']);
    enhanced['commission_percent'] = financeKvvVal_(row.commission_percent);
    enhanced['ppvz_kvw_prc'] = financeKvvVal_(row.ppvz_kvw_prc);
    enhanced['ppvz_kvw_prc_base'] = financeKvvVal_(row.ppvz_kvw_prc_base);
    enhanced['site_country'] = financeCountry_(row.site_country);
    enhanced['delivery_rub'] = financeV1Money_(row.delivery_rub);
    enhanced['ppvz_for_pay'] = financeV1Money_(row.ppvz_for_pay);
    
    // Дата для отчётов + rr_dt строкой (не Date), иначе ломаются ID рядом
    enhanced['Дата'] = dateObj || '';
    if (dateObj) {
      enhanced['rr_dt'] = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (rrDt) {
      enhanced['rr_dt'] = rrDt.toString().substring(0, 10);
    }
    enhanced['Вес, кг за 1 шт'] = weightPerUnit;
    enhanced['Вес, кг за 1 шт сумма'] = totalWeight;
    enhanced['Сумма продаж'] = salesAmount;
    enhanced['Имя товара'] = productName;
    enhanced['Сумма продаж с учетом возвратов'] = salesWithReturns;
    enhanced['Сумма упаковок с учетом возвратов'] = quantityWithReturns;
    enhanced['Вес итог, кг с учетом возвратов'] = totalWeightWithReturns;
    
    // ДОБАВЛЯЕМ ДВА НОВЫХ СТОЛБЦА В КОНЕЦ
    enhanced['Цена со скидкой, ₽ (из Товары)'] = priceWithDiscount;
    enhanced['Сумма по текущей цене'] = amountAtCurrentPrice;
    
    if (typeof wbRowToRu_ === 'function') enhanced = wbRowToRu_(enhanced);
    result.push(enhanced);
  }
  
  console.log('Обработано финансовых записей: ' + result.length);
  
  // Отладка: показываем первые 3 записи с результатами
  for (var i = 0; i < Math.min(3, result.length); i++) {
    console.log('Пример ' + (i+1) + ': nmId=' + result[i].nm_id + 
                ', Имя товара="' + result[i]['Имя товара'] + 
                '", Цена со скидкой=' + result[i]['Цена со скидкой, ₽ (из Товары)'] +
                '", Сумма по текущей цене=' + result[i]['Сумма по текущей цене']);
  }
  
  return result;
}

// ОПТИМИЗИРОВАННОЕ ДОБАВЛЕНИЕ В ЛИСТ
function addDataToSheetOptimized(data, sheetName, uniqueKeyField) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet && typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  if (!(typeof morningChainOn_ === 'function' && morningChainOn_()) &&
      typeof renameWbColumnHeaders_ === 'function') {
    renameWbColumnHeaders_(ss);
  }
  sheet = ss.getSheetByName(sheetName);
  var existingKeys = {};
  function finCol_(headers, name) {
    return typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [name]) : headers.indexOf(name);
  }
  function finPick_(item, name) {
    return typeof wbPickObj_ === 'function' ? wbPickObj_(item, [name]) : (item[name] != null ? item[name] : '');
  }
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = Object.keys(data[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    formatFinancialIdColumns_(sheet, 2, 1000);
  } else {
    ensureFinanceKvvColumns_(sheet);
    ensureFinanceCountryColumn_(sheet);
    ensureFinanceLogisticsColumn_(sheet);
    ensureFinanceForPayColumn_(sheet);
    // Собираем существующие ключи (rrd_id + nm_id + shk_id + doc_type_name)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var rrdIdIdx = finCol_(sheetHeaders, 'rrd_id');
      var nmIdIdx = finCol_(sheetHeaders, 'nm_id');
      var shkIdIdx = finCol_(sheetHeaders, 'shk_id');
      var docTypeIdx = finCol_(sheetHeaders, 'doc_type_name');
      
      if (rrdIdIdx !== -1 && nmIdIdx !== -1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (var i = 0; i < existingData.length; i++) {
          var row = existingData[i];
          var key = row[rrdIdIdx] + '|' + row[nmIdIdx] + '|' + (row[shkIdIdx] || '') + '|' + (row[docTypeIdx] || '');
          existingKeys[key] = true;
        }
      }
    }
  }
  
  // Фильтруем новые записи
  var newData = [];
  for (var i = 0; i < data.length; i++) {
    var item = data[i];
    var key = finPick_(item, 'rrd_id') + '|' + finPick_(item, 'nm_id') + '|' + (finPick_(item, 'shk_id') || '') + '|' + (finPick_(item, 'doc_type_name') || '');
    if (!existingKeys[key]) {
      newData.push(item);
      existingKeys[key] = true;
    }
  }
  
  if (newData.length === 0) {
    if (!isRunningFromTrigger()) {
      SpreadsheetApp.getUi().alert('⚠️ Новых записей не найдено');
    }
    wbPutSheetFilter_(sheet);
    return;
  }
  
  // Получаем заголовки
  var headers = Object.keys(newData[0]);
  
  // Проверяем и добавляем недостающие столбцы в существующем листе
  var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastCol = sheet.getLastColumn();
                          
  var requiredCalcColumns = ['Дата', 'Имя товара', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма', 
                              'Сумма продаж', 'Сумма продаж с учетом возвратов', 
                              'Сумма упаковок с учетом возвратов', 'Вес итог, кг с учетом возвратов',
                              'Цена со скидкой, ₽ (из Товары)', 'Сумма по текущей цене'];

  for (var i = 0; i < requiredCalcColumns.length; i++) {
    if (sheetHeaders.indexOf(requiredCalcColumns[i]) === -1) {
      sheet.getRange(1, lastCol + 1).setValue(requiredCalcColumns[i]);
      sheet.getRange(1, lastCol + 1).setFontWeight('bold');
      lastCol++;
      sheetHeaders.push(requiredCalcColumns[i]);
    }
  }
  
  // Формируем строки для записи
  var rows = [];
  for (var i = 0; i < newData.length; i++) {
    var row = [];
    for (var j = 0; j < sheetHeaders.length; j++) {
      var val = typeof wbCell_ === 'function' ? wbCell_(newData[i], sheetHeaders[j]) : newData[i][sheetHeaders[j]];
      if (val === null || val === undefined) val = '';
      if (isFinanceIdHeader_(sheetHeaders[j])) val = financeIdAsText_(val);
      row.push(val);
    }
    rows.push(row);
  }

  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1;
  // Текстовый формат ID-колонок ДО записи — иначе nm_id/barcode снова станут датами
  sheet.getRange(startRow, 1, rows.length, sheetHeaders.length).setValues(rows);
  formatFinancialIdColumns_(sheet, startRow, rows.length);
  
  // Форматируем столбец "Дата"
  var dateColIndex = sheetHeaders.indexOf('Дата') + 1;
  if (dateColIndex > 0 && sheet.getLastRow() > 1) {
    sheet.getRange(2, dateColIndex, sheet.getLastRow() - 1, 1).setNumberFormat("dd.mm.yyyy");
  }
  
  // Форматируем числовые столбцы
  var numberColumns = ['Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма', 'Сумма продаж', 
                        'Сумма продаж с учетом возвратов', 'Сумма упаковок с учетом возвратов', 
                        'Вес итог, кг с учетом возвратов',
                        'commission_percent', 'ppvz_kvw_prc', 'ppvz_kvw_prc_base',
                        'кВВ %', 'кВВ без НДС', 'кВВ база',
                        'delivery_rub', 'Логистика',
                        'ppvz_for_pay', 'К Перечислению', 'К перечислению'];
  
  for (var i = 0; i < numberColumns.length; i++) {
    var colIndex = (typeof wbColIdx_ === 'function' ? wbColIdx_(sheetHeaders, [numberColumns[i]]) : sheetHeaders.indexOf(numberColumns[i])) + 1;
    if (colIndex > 0 && sheet.getLastRow() > 1) {
      sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00");
    }
  }
  
  if (!isRunningFromTrigger()) {
    SpreadsheetApp.getUi().alert('✅ Добавлено ' + newData.length + ' новых записей');
  }
  Logger.log('Добавлено новых финансовых записей: ' + newData.length);
  wbPutSheetFilter_(sheet);
}

function ensureFinanceKvvColumns_(sheet) {
  if (!sheet) return;
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  function headerIndex_(name) {
    var last = Math.max(1, sheet.getLastColumn());
    var h = sheet.getRange(1, 1, 1, last).getValues()[0];
    if (typeof wbColIdx_ === 'function') return wbColIdx_(h, [name]);
    for (var c = 0; c < h.length; c++) {
      if (String(h[c] || '').trim() === name) return c;
    }
    return -1;
  }
  // Только в конец листа: insert посередине сдвигает буквы столбцов
  // и ломает формулы вида INDIRECT/QUERY/"столбец L".
  var names = [
    { en: 'commission_percent', ru: 'кВВ %' },
    { en: 'ppvz_kvw_prc', ru: 'кВВ без НДС' },
    { en: 'ppvz_kvw_prc_base', ru: 'кВВ база' }
  ];
  for (var i = 0; i < names.length; i++) {
    if (headerIndex_(names[i].en) >= 0 || headerIndex_(names[i].ru) >= 0) continue;
    var lastCol = Math.max(1, sheet.getLastColumn());
    try {
      sheet.insertColumnAfter(lastCol);
      sheet.getRange(1, lastCol + 1).setValue(names[i].ru).setFontWeight('bold');
    } catch (eIns) {
      Logger.log('ensureFinanceKvvColumns_ insert: ' + eIns);
    }
  }
}

function ensureFinanceCountryColumn_(sheet) {
  if (!sheet) return;
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  var last = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  var idx = typeof wbColIdx_ === 'function'
    ? wbColIdx_(headers, ['Страна продажи', 'site_country', 'country'])
    : -1;
  if (idx < 0) {
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim().toLowerCase();
      if (h === 'страна продажи' || h === 'site_country' || h === 'country') {
        idx = c;
        break;
      }
    }
  }
  if (idx >= 0) return;
  try {
    sheet.insertColumnAfter(last);
    sheet.getRange(1, last + 1).setValue('Страна продажи').setFontWeight('bold');
  } catch (eIns) {
    Logger.log('ensureFinanceCountryColumn_ insert: ' + eIns);
  }
}

function ensureFinanceLogisticsColumn_(sheet) {
  if (!sheet) return;
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  var last = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  var idx = typeof wbColIdx_ === 'function'
    ? wbColIdx_(headers, ['Логистика', 'delivery_rub', 'deliveryService'])
    : -1;
  if (idx < 0) {
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim().toLowerCase();
      if (h === 'логистика' || h === 'delivery_rub' || h === 'deliveryservice') {
        idx = c;
        break;
      }
    }
  }
  if (idx >= 0) return;
  try {
    sheet.insertColumnAfter(last);
    sheet.getRange(1, last + 1).setValue('Логистика').setFontWeight('bold');
  } catch (eIns) {
    Logger.log('ensureFinanceLogisticsColumn_ insert: ' + eIns);
  }
}

function ensureFinanceForPayColumn_(sheet) {
  if (!sheet) return;
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  var last = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  var idx = typeof wbColIdx_ === 'function'
    ? wbColIdx_(headers, ['К Перечислению', 'К перечислению', 'ppvz_for_pay', 'forPay'])
    : -1;
  if (idx < 0) {
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim().toLowerCase();
      if (h === 'к перечислению' || h === 'ppvz_for_pay' || h === 'forpay') {
        idx = c;
        break;
      }
    }
  }
  if (idx >= 0) return;
  try {
    sheet.insertColumnAfter(last);
    sheet.getRange(1, last + 1).setValue('К Перечислению').setFontWeight('bold');
  } catch (eIns) {
    Logger.log('ensureFinanceForPayColumn_ insert: ' + eIns);
  }
}

var DB_URL_DEFAULT_ = Utilities.newBlob(Utilities.base64Decode(
  'aHR0cHM6Ly9ldmdoY21lcWFlbGNrYWd5eGdtay5zdXBhYmFzZS5jbw=='
)).getDataAsString();

function setDbKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt(
    'Ключ базы',
    'JWT, начинается с eyJ.\nКороткий ключ вида sb_… не подойдёт: Apps Script шлёт User-Agent как браузер.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = String(r.getResponseText() || '').trim();
  if (!key) {
    ui.alert('Ключ пустой — не сохранил');
    return;
  }
  if (key.indexOf('sb_') === 0) {
    ui.alert('Этот формат ключа не подойдёт. Нужен JWT, начинается с eyJ.');
    return;
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty('DB_KEY', key);
  props.setProperty('DB_URL', DB_URL_DEFAULT_);
  ui.alert('Ключ базы сохранён');
}

function getDbConfig_() {
  var props = PropertiesService.getScriptProperties();
  var url = String(props.getProperty('DB_URL') || '').replace(/\/$/, '');
  var key = String(props.getProperty('DB_KEY') || '').trim();
  if (!url || !key) {
    var all = props.getProperties();
    for (var k in all) {
      var v = String(all[k] || '').trim();
      if (!key && v.indexOf('eyJ') === 0) key = v;
      if (!url && /^https:\/\//i.test(v)) url = v.replace(/\/$/, '');
    }
  }
  if (!url) url = DB_URL_DEFAULT_;
  return { url: url, key: key };
}

function dbRestHeaders_(key) {
  var headers = {
    'apikey': key,
    'User-Agent': 'WB-Sheets/1.0'
  };
  if (key.indexOf('eyJ') === 0) {
    headers['Authorization'] = 'Bearer ' + key;
  }
  return headers;
}

function dbFetchJson_(path, options) {
  var cfg = getDbConfig_();
  if (!cfg.key) throw new Error('Нет ключа базы. Меню: «Ключ базы».');
  options = options || {};
  var method = String(options.method || 'get').toLowerCase();
  var req = {
    method: method,
    muteHttpExceptions: true,
    headers: dbRestHeaders_(cfg.key)
  };
  if (options.body !== undefined) {
    req.contentType = 'application/json';
    req.payload = JSON.stringify(options.body);
  }
  var res = UrlFetchApp.fetch(cfg.url + path, req);
  var code = res.getResponseCode();
  var text = res.getContentText() || '';
  if (code < 200 || code >= 300) {
    throw new Error('БД HTTP ' + code + ' ' + text.substring(0, 300));
  }
  return text ? JSON.parse(text) : null;
}

// ============================================
// ВАША ОРИГИНАЛЬНАЯ ФУНКЦИЯ (НЕ МЕНЯЕТСЯ)
// ============================================
function getAdvertCampaigns(token) {
  var url = 'https://advert-api.wildberries.ru/adv/v1/promotion/count';
  var options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var allCampaigns = [];
      if (data.adverts && Array.isArray(data.adverts)) {
        for (var i = 0; i < data.adverts.length; i++) {
          var group = data.adverts[i];
          if (group.advert_list && Array.isArray(group.advert_list)) {
            for (var j = 0; j < group.advert_list.length; j++) {
              allCampaigns.push({
                id: group.advert_list[j].advertId,
                status: group.status,
                type: group.type
              });
            }
          }
        }
      }
      return allCampaigns.filter(function(c) {
        return [7, 9, 11].indexOf(c.status) !== -1;
      });
    }
  } catch (error) {
    Logger.log('Ошибка получения кампаний: ' + error.toString());
  }
  return [];
}

// ============================================
// ОСНОВНАЯ ФУНКЦИЯ - ТОЛЬКО ДОБАВЛЕНЫ 3 СТОЛБЦА
// ============================================
function loadAdvExpenses(token, dateFrom, dateTo) {
  if (typeof renameWbColumnHeaders_ === 'function') renameWbColumnHeaders_();
  var campaigns = getAdvertCampaigns(token);
  if (!campaigns || campaigns.length === 0) {
    if (!isRunningFromTrigger()) {
      SpreadsheetApp.getUi().alert('⚠️ Не найдено рекламных кампаний.');
    }
    Logger.log('Рекламные кампании не найдены');
    return;
  }
  
  var start = new Date(dateFrom);
  var end = new Date(dateTo);
  var daysDiff = Math.ceil((end - start) / (1000 * 60 * 60 * 24));
  
  if (daysDiff > 31) {
    if (!isRunningFromTrigger()) {
      SpreadsheetApp.getUi().alert('⚠️ Период не должен превышать 31 день!');
    }
    Logger.log('Период превышает 31 день');
    return;
  }
  
  // Загружаем данные
  var campaignIds = campaigns.map(function(c) { return c.id; }).join(',');
  var url = 'https://advert-api.wildberries.ru/adv/v3/fullstats?ids=' + campaignIds + '&beginDate=' + dateFrom + '&endDate=' + dateTo;
  var response = UrlFetchApp.fetch(url, { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true });
  
  if (response.getResponseCode() !== 200) {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('Ошибка API: ' + response.getResponseCode());
    return;
  }
  
  var data = JSON.parse(response.getContentText());
  
  // Функция для преобразования ISO даты в формат DD.MM.YYYY
  function formatDate(isoDateString) {
    if (!isoDateString) return '';
    var date = new Date(isoDateString);
    var day = date.getDate().toString().padStart(2, '0');
    var month = (date.getMonth() + 1).toString().padStart(2, '0');
    var year = date.getFullYear();
    return day + '.' + month + '.' + year;
  }
  
  // Разворачиваем в плоскую таблицу
  var allStats = [];
  for (var i = 0; i < data.length; i++) {
    var campaign = data[i];
    if (campaign.days) {
      for (var j = 0; j < campaign.days.length; j++) {
        var day = campaign.days[j];
        if (day.apps) {
          for (var k = 0; k < day.apps.length; k++) {
            var app = day.apps[k];
            if (app.nms && app.nms.length > 0) {
              for (var l = 0; l < app.nms.length; l++) {
                var nm = app.nms[l];
                allStats.push({
                  'advertId': campaign.advertId,
                  'date': day.date,
                  'appType': app.appType,
                  'nmId': nm.nmId,
                  'name': nm.name || nm.nmName || nm.title || '',
                  'views': nm.views || 0,
                  'clicks': nm.clicks || 0,
                  'sum': (nm.sum || 0) / 100,
                  'atbs': nm.atbs || 0,
                  'orders': nm.orders || 0,
                  'shks': nm.shks || 0,
                  'cpc': nm.cpc || 0,
                  'ctr': nm.ctr || 0,
                  'cr': nm.cr || 0
                });
              }
            } else {
              allStats.push({
                'advertId': campaign.advertId,
                'date': day.date,
                'appType': app.appType,
                'nmId': '',
                'name': 'ИТОГО по платформе',
                'views': app.views || 0,
                'clicks': app.clicks || 0,
                'sum': (app.sum || 0) / 100,
                'atbs': app.atbs || 0,
                'orders': app.orders || 0,
                'shks': app.shks || 0,
                'cpc': app.cpc || 0,
                'ctr': app.ctr || 0,
                'cr': app.cr || 0
              });
            }
          }
        }
      }
    }
  }
  
  // Получаем источник списания
  var urlUpd = 'https://advert-api.wildberries.ru/adv/v1/upd?from=' + dateFrom + '&to=' + dateTo;
  var updResponse = UrlFetchApp.fetch(urlUpd, { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true });
  var paymentTypeMap = {};
  if (updResponse.getResponseCode() === 200) {
    var updData = JSON.parse(updResponse.getContentText());
    for (var i = 0; i < updData.length; i++) {
      var item = updData[i];
      var key = item.advertId + '_' + (item.updTime ? item.updTime.substring(0,10) : '');
      paymentTypeMap[key] = item.paymentType || 'Неизвестно';
    }
  }
  
  for (var i = 0; i < allStats.length; i++) {
    var stat = allStats[i];
    var dateKey = stat.date ? stat.date.substring(0,10) : '';
    stat.paymentType = paymentTypeMap[stat.advertId + '_' + dateKey] || '';
  }
  
  // ============================================
  // ДОБАВЛЯЕМ ТРИ НОВЫХ ПОЛЯ В КОНЕЦ
  // ============================================
  
  // Артикул и название с листа «Товары».
  // fullstats больше не отдаёт nm.name — без каталога столбец «Название» пустой.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var productsSheet = ss.getSheetByName('Товары');
  var nmIdToArticle = {};
  var nmIdToName = {};
  
  if (productsSheet && productsSheet.getLastRow() > 1) {
    var productsData = productsSheet.getDataRange().getValues();
    var pHeaders = productsData[0];
    var nmCol = typeof wbColIdx_ === 'function' ? wbColIdx_(pHeaders, ['Артикул WB', 'nmId']) : 0;
    var nameCol = typeof wbColIdx_ === 'function' ? wbColIdx_(pHeaders, ['Название', 'name', 'title']) : 1;
    var artCol = typeof wbColIdx_ === 'function' ? wbColIdx_(pHeaders, ['Ваш артикул', 'Арт. продавца', 'vendorCode']) : 2;
    if (nmCol < 0) nmCol = 0;
    if (nameCol < 0) nameCol = 1;
    if (artCol < 0) artCol = 2;
    for (var i = 1; i < productsData.length; i++) {
      var nmIdVal = productsData[i][nmCol];
      if (!nmIdVal) continue;
      var nmKey = nmIdVal.toString();
      nmIdToArticle[nmKey] = productsData[i][artCol] || '';
      if (productsData[i][nameCol]) nmIdToName[nmKey] = String(productsData[i][nameCol]);
    }
  }
  
  // Добавляем три новых поля к каждой записи
  for (var i = 0; i < allStats.length; i++) {
    var nmIdStr = allStats[i].nmId ? allStats[i].nmId.toString() : '';
    
    allStats[i].article = nmIdToArticle[nmIdStr] || '';
    if (!allStats[i].name && nmIdStr && nmIdToName[nmIdStr]) {
      allStats[i].name = nmIdToName[nmIdStr];
    }
    allStats[i].dateFormatted = formatDate(allStats[i].date);
    allStats[i].advCost = (allStats[i].clicks || 0) * (allStats[i].cpc || 0);
  }
  
  // Формируем итоговый массив (原有 столбцы + 3 новых в конце)
  var finalStats = [];
  for (var i = 0; i < allStats.length; i++) {
    finalStats.push({
      'ID кампании': allStats[i].advertId,
      'Дата статы': formatDate(allStats[i].date),
      'Площадка': allStats[i].appType,
      'Артикул WB': allStats[i].nmId,
      'Название': allStats[i].name,
      'Показы': allStats[i].views,
      'Клики': allStats[i].clicks,
      'Сумма API': allStats[i].sum,
      'Корзина': allStats[i].atbs,
      'Заказы рек.': allStats[i].orders,
      'Выкупы рек.': allStats[i].shks,
      'CPC': allStats[i].cpc,
      'CTR %': allStats[i].ctr,
      'CR %': allStats[i].cr,
      'Источник списания': allStats[i].paymentType,
      'Артикул': allStats[i].article,
      'Дата': allStats[i].dateFormatted,
      'Расходы рекламы': allStats[i].advCost
    });
  }
  
  if (finalStats.length === 0) {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⚠️ Нет рекламных расходов за период');
    return;
  }
  
  // ============================================
  // ЗАПИСЬ В ЛИСТ
  // ============================================
  var sheetName = 'Рекламные_расходы';
  var sheet = ss.getSheetByName(sheetName);
  var existingKeys = {};
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = Object.keys(finalStats[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Собираем существующие ключи для избежания дублей (используем исходные ISO даты из API)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var id1Idx = typeof wbColIdx_ === 'function' ? wbColIdx_(sheetHeaders, ['advertId']) : sheetHeaders.indexOf('advertId');
      var id2Idx = typeof wbColIdx_ === 'function' ? wbColIdx_(sheetHeaders, ['date', 'Дата статы']) : sheetHeaders.indexOf('date');
      var id3Idx = typeof wbColIdx_ === 'function' ? wbColIdx_(sheetHeaders, ['nmId']) : sheetHeaders.indexOf('nmId');
      var id4Idx = typeof wbColIdx_ === 'function' ? wbColIdx_(sheetHeaders, ['appType']) : sheetHeaders.indexOf('appType');
      
      if (id1Idx !== -1 && id2Idx !== -1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (var i = 0; i < existingData.length; i++) {
          var row = existingData[i];
          var key = row[id1Idx] + '|' + adsDateKey_(row[id2Idx]) + '|' + (row[id3Idx] || '') + '|' + (row[id4Idx] || '');
          existingKeys[key] = true;
        }
      }
    }
  }
  
  // Фильтруем новые записи (используем исходные ISO даты для сравнения)
  var newStats = [];
  for (var i = 0; i < allStats.length; i++) {
    var originalItem = allStats[i];
    var key = originalItem.advertId + '|' + adsDateKey_(originalItem.date) + '|' + (originalItem.nmId || '') + '|' + (originalItem.appType || '');
    if (!existingKeys[key]) {
      newStats.push(finalStats[i]);  // добавляем уже отформатированную версию
      existingKeys[key] = true;
    }
  }
  
  var filledNames = adsFillEmptyNames_(sheet, nmIdToName);
  if (newStats.length === 0) {
    if (!isRunningFromTrigger()) {
      SpreadsheetApp.getUi().alert(filledNames
        ? '⚠️ Новых записей нет. Пустые названия заполнены: ' + filledNames
        : '⚠️ Новых рекламных записей не найдено');
    }
    Logger.log('Новых рекламных записей нет, названия заполнены: ' + filledNames);
    wbPutSheetFilter_(sheet);
    return;
  }
  
  // Находим последнюю реальную строку
  var lastRow = sheet.getLastRow();
  var realLastRow = lastRow;
  for (var i = lastRow; i >= 1; i--) {
    var cellValue = sheet.getRange(i, 1, 1, 1).getValue();
    if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
      realLastRow = i;
      break;
    }
  }
  
  var startRow = realLastRow + 1;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers || !headers.length || !headers[0]) {
    headers = Object.keys(newStats[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  var rows = [];
  for (var i = 0; i < newStats.length; i++) {
    rows.push(typeof wbFillRow_ === 'function' ? wbFillRow_(newStats[i], headers) : Object.keys(newStats[i]).map(function(k){ return newStats[i][k]; }));
  }
  
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  
  // Форматируем столбец "Расходы рекламы" как число
  var expenseColIndex = headers.indexOf('Расходы рекламы') + 1;
  if (expenseColIndex > 0 && sheet.getLastRow() > 1) {
    sheet.getRange(2, expenseColIndex, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00");
  }
  
  if (!isRunningFromTrigger()) {
    var msg = '✅ Добавлено ' + newStats.length + ' новых записей';
    if (filledNames) msg += '\nПустые названия заполнены: ' + filledNames;
    SpreadsheetApp.getUi().alert(msg);
  }
  Logger.log('Добавлено новых рекламных записей: ' + newStats.length + ', названия: ' + filledNames);
  wbPutSheetFilter_(sheet);
}

/** Уже записанные строки, где fullstats пришёл без названия. */
function adsFillEmptyNames_(sheet, nmIdToName) {
  if (!sheet || sheet.getLastRow() < 2 || !nmIdToName) return 0;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameIdx = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['Название', 'name']) : headers.indexOf('Название');
  var nmIdx = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['Артикул WB', 'nmId']) : headers.indexOf('Артикул WB');
  if (nameIdx < 0 || nmIdx < 0) return 0;
  var n = sheet.getLastRow() - 1;
  var names = sheet.getRange(2, nameIdx + 1, n, 1).getValues();
  var nms = sheet.getRange(2, nmIdx + 1, n, 1).getValues();
  var changed = false;
  var count = 0;
  var i, nm, title;
  for (i = 0; i < names.length; i++) {
    if (String(names[i][0] || '').trim()) continue;
    nm = nms[i][0] != null ? String(nms[i][0]).trim() : '';
    title = nm && nmIdToName[nm];
    if (!title) continue;
    names[i][0] = title;
    changed = true;
    count++;
  }
  if (changed) sheet.getRange(2, nameIdx + 1, n, 1).setValues(names);
  return count;
}


// 8. ТОВАРЫ С ЦЕНАМИ (исправленная версия)
function loadAllProducts() {
  if (!(typeof morningChainOn_ === 'function' && morningChainOn_()) &&
      typeof renameWbColumnHeaders_ === 'function') {
    renameWbColumnHeaders_();
  }
  var token = getApiToken();
  if (!token) {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⚠️ Токен не найден');
    Logger.log('Токен не найден');
    return;
  }

  // 1. Загружаем карточки товаров
  var allProducts = [];
  var cursor = null;
  var hasMore = true;

  while (hasMore) {
    var payload = {
      "settings": {
        "cursor": { "limit": 100 },
        "filter": { "withPhoto": -1 }
      }
    };
    
    if (cursor) {
      payload.settings.cursor.updatedAt = cursor.updatedAt;
      payload.settings.cursor.nmID = cursor.nmID;
    }
    
    var options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      var response = UrlFetchApp.fetch('https://content-api.wildberries.ru/content/v2/get/cards/list', options);
      if (response.getResponseCode() !== 200) break;
      
      var data = JSON.parse(response.getContentText());
      if (data.cards && data.cards.length > 0) {
        allProducts = allProducts.concat(data.cards);
      }
      
      if (data.cursor && data.cursor.updatedAt && data.cursor.nmID && data.cards.length === 100) {
        cursor = { updatedAt: data.cursor.updatedAt, nmID: data.cursor.nmID };
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } catch (error) {
      if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('Ошибка: ' + error.toString());
      Logger.log('Ошибка: ' + error.toString());
      return;
    }
  }
  
  if (allProducts.length === 0) {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⚠️ Товары не найдены.');
    Logger.log('Товары не найдены');
    return;
  }
  
  // 2. Загружаем цены через GET-запрос с пагинацией
  if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⏳ Загружено ' + allProducts.length + ' товаров. Теперь загружаем цены...');
  
  var priceMap = {};
  var offset = 0;
  var limit = 1000;
  var hasMorePrices = true;
  
  while (hasMorePrices) {
    var url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=' + limit + '&offset=' + offset;
    
    var priceOptions = {
      method: 'GET',
      headers: { 'Authorization': token },
      muteHttpExceptions: true
    };
    
    try {
      var priceResponse = UrlFetchApp.fetch(url, priceOptions);
      
      if (priceResponse.getResponseCode() === 429) {
        Logger.log('⚠️ Цены: 429, ждём 6 сек...');
        Utilities.sleep(6000);
        continue;
      }
      
      if (priceResponse.getResponseCode() === 200) {
        var priceData = JSON.parse(priceResponse.getContentText());
        
        // Проверяем структуру ответа
        if (priceData.data && priceData.data.listGoods && priceData.data.listGoods.length > 0) {
          var goodsList = priceData.data.listGoods;
          
          for (var i = 0; i < goodsList.length; i++) {
            var item = goodsList[i];
            var nmId = item.nmID;
            
            if (nmId && item.sizes && item.sizes.length > 0) {
              // Берем первый размер (обычно он один для товаров без размеров)
              var sizeInfo = item.sizes[0];
              
              // Цена в рублях (price уже в рублях, не нужно делить на 100)
              var priceRub = sizeInfo.price || 0;
              // Цена со скидкой
              var discountedPrice = sizeInfo.discountedPrice || 0;
              // Цена для WB Клуба
              var clubDiscountedPrice = sizeInfo.clubDiscountedPrice || 0;
              
              // Вычисляем процент скидки
              var discountPercent = 0;
              if (priceRub > 0 && discountedPrice > 0) {
                discountPercent = Math.round((1 - discountedPrice / priceRub) * 100);
              }
              
              // Вычисляем скидку WB Клуба
              var clubDiscountPercent = 0;
              if (priceRub > 0 && clubDiscountedPrice > 0) {
                clubDiscountPercent = Math.round((1 - clubDiscountedPrice / priceRub) * 100);
              }
              
              priceMap[nmId] = {
                price: priceRub,                          // Полная цена в рублях
                discountedPrice: discountedPrice,         // Цена со скидкой
                clubDiscountedPrice: clubDiscountedPrice, // Цена для WB Клуба
                discount: discountPercent,                // Скидка %
                clubDiscount: clubDiscountPercent,        // Скидка WB Клуба %
                currency: item.currencyIsoCode4217 || 'RUB',
                promoCodeDiscount: 0                      // Промокод скидка (не приходит в этом API)
              };
            }
          }
          
          // Если получили меньше, чем limit, значит это последняя страница
          if (goodsList.length < limit) {
            hasMorePrices = false;
          } else {
            offset += limit;
          }
        } else {
          hasMorePrices = false;
        }
      } else {
        Logger.log('❌ Ошибка API цен: ' + priceResponse.getResponseCode());
        hasMorePrices = false;
      }
      
      // Пауза между запросами для соблюдения лимитов
      Utilities.sleep(600);
      
    } catch(e) {
      Logger.log('❌ Ошибка при загрузке цен: ' + e);
      hasMorePrices = false;
    }
  }
  
  // 3. Сохраняем товары с ценами на лист
  saveProductsWithPricesToSheet(allProducts, priceMap);
  
  if (!isRunningFromTrigger()) {
    var priceCount = Object.keys(priceMap).length;
    SpreadsheetApp.getUi().alert('✅ Загружено товаров: ' + allProducts.length + '\nС ценами: ' + priceCount);
  }
  Logger.log('✅ Товары: ' + allProducts.length + ' карточек, ' + Object.keys(priceMap).length + ' с ценами');
}

// Сохранение товаров с ценами (обновленная версия)
function saveProductsWithPricesToSheet(products, priceMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Товары';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Очищаем лист
  sheet.clear();
  
  // Расширенные заголовки с ценами
  var headers = [
    'Артикул WB', 
    'Название', 
    'Ваш артикул', 
    'Бренд', 
    'Категория', 
    'Штрих-код',
    'Цена, ₽',
    'Цена со скидкой, ₽',
    'Скидка %',
    'Цена WB Клуба, ₽',
    'Скидка WB Клуба %',
    'Валюта',
    'Дата обновления цен'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  var today = new Date();
  var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  var rows = [];
  var withPriceCount = 0;
  
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    var nmID = p.nmID || '';
    var priceInfo = priceMap[nmID] || {};
    
    // Считаем сколько товаров с ценой
    if (priceInfo.price && priceInfo.price > 0) {
      withPriceCount++;
    }
    
    rows.push([
      nmID,
      p.title || '',
      p.vendorCode || '',
      p.brand || '',
      p.subjectName || '',
      p.barcode || '',
      priceInfo.price || '',
      priceInfo.discountedPrice || '',
      priceInfo.discount || '',
      priceInfo.clubDiscountedPrice || '',
      priceInfo.clubDiscount || '',
      priceInfo.currency || 'RUB',
      priceInfo.price ? dateStr : ''
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
    
    // Форматируем колонки с ценами как валюту
    var priceColumns = ['Цена, ₽', 'Цена со скидкой, ₽', 'Цена WB Клуба, ₽'];
    for (var c = 0; c < priceColumns.length; c++) {
      var colIndex = headers.indexOf(priceColumns[c]) + 1;
      if (colIndex > 0) {
        sheet.getRange(2, colIndex, rows.length, 1).setNumberFormat('#,##0.00 ₽');
      }
    }
    
    // Форматируем колонки со скидками как проценты
    var discountColumns = ['Скидка %', 'Скидка WB Клуба %'];
    for (var d = 0; d < discountColumns.length; d++) {
      var colIndex = headers.indexOf(discountColumns[d]) + 1;
      if (colIndex > 0) {
        sheet.getRange(2, colIndex, rows.length, 1).setNumberFormat('0"%"');
      }
    }
  }
  wbPutSheetFilter_(sheet);
}

// ============================================
// История цен: дневной снимок каталога.
// Тот же источник, что product_prices_daily в Supabase:
// GET discounts-prices-api /api/v2/list/goods/filter, одна строка на nm × день (МСК).
// Повтор за сегодня заменяет только этот день. Проба uploadID — отдельный лист.
// ============================================

var PRICE_HISTORY_SHEET_ = 'История цен';
var PRICE_HISTORY_HEADERS_ = [
  'Дата', 'Артикул WB', 'Арт. продавца', 'Название',
  'Цена, ₽', 'Цена со скидкой, ₽', 'Цена клуба, ₽',
  'Скидка прод. %', 'Валюта'
];

function savePriceHistory() {
  var token = getApiToken();
  if (!token) {
    if (!isRunningFromTrigger()) showAlert('⚠️ Токен не найден');
    Logger.log('История цен: нет токена');
    return { rows: 0 };
  }

  var goods = fetchWbListGoods_(token);
  var productsDict = getProductsDictionary();
  var dateStr = Utilities.formatDate(new Date(), sheetTimeZone_(), 'yyyy-MM-dd');
  var day = lkStocksDate_(dateStr);
  var rows = [];
  var i;
  for (i = 0; i < goods.length; i++) {
    var row = priceHistoryRow_(goods[i], day, productsDict);
    if (row) rows.push(row);
  }
  rows.sort(function(a, b) {
    var na = Number(a[1]);
    var nb = Number(b[1]);
    if (na !== nb) return na - nb;
    return String(a[1]).localeCompare(String(b[1]));
  });

  var written = writePriceHistory_(dateStr, rows);
  Logger.log('История цен: ' + written.rows + ' строк за ' + dateStr + ', снято повтор ' + written.removed);
  if (!isRunningFromTrigger()) {
    showAlert(
      '✅ История цен\n' +
      'Дата: ' + Utilities.formatDate(lkStocksDate_(dateStr), sheetTimeZone_(), 'dd.MM.yyyy') + '\n' +
      'Товаров: ' + written.rows + '\n' +
      'Лист: ' + PRICE_HISTORY_SHEET_ + '\n\n' +
      'Повтор за сегодня перезаписывает только эту дату.'
    );
  }
  return written;
}

function fetchWbListGoods_(token) {
  var host = 'https://discounts-prices-api.wildberries.ru';
  var goods = [];
  var offset = 0;
  var limit = 1000;
  var page;
  var retries = 0;
  for (page = 0; page < 50; page++) {
    var res = wrFetchJson_(token, host + '/api/v2/list/goods/filter?limit=' + limit + '&offset=' + offset, 'GET');
    if (res.code === 429) {
      retries++;
      if (retries > 5) throw new Error('WB лимит цен (429). Подождите минуту и повторите.');
      Logger.log('История цен: 429, пауза');
      Utilities.sleep(6000);
      page--;
      continue;
    }
    retries = 0;
    if (res.code === 403) {
      throw new Error('HTTP 403 — нужен токен с категорией «Цены и скидки».');
    }
    if (res.code !== 200) {
      throw new Error('Цены HTTP ' + res.code + ' — ' + String(res.text || '').substring(0, 200));
    }
    var data = (res.data && res.data.data) ? res.data.data : (res.data || {});
    var chunk = data.listGoods || [];
    if (!chunk.length) break;
    goods = goods.concat(chunk);
    if (chunk.length < limit) break;
    offset += limit;
    Utilities.sleep(600);
  }
  return goods;
}

function priceHistoryRow_(item, day, productsDict) {
  item = item || {};
  var nmId = item.nmID != null ? item.nmID : item.nmId;
  if (nmId == null || nmId === '') return null;
  var sizes = item.sizes || [];
  var size0 = sizes.length ? sizes[0] : {};
  var price = priceHistoryNum_(size0.price);
  var discounted = priceHistoryNum_(size0.discountedPrice);
  var club = priceHistoryNum_(size0.clubDiscountedPrice);
  var discount = '';
  if (price && discounted) {
    discount = Math.round((1 - discounted / price) * 100);
  }
  var product = lkProduct_(productsDict, nmId);
  var vendor = item.vendorCode != null ? String(item.vendorCode) : '';
  if (!vendor && product.vendorCode) vendor = String(product.vendorCode);
  return [
    day,
    nmId,
    vendor,
    product.title || '',
    price === null ? '' : price,
    discounted === null ? '' : discounted,
    club === null ? '' : club,
    discount,
    item.currencyIsoCode4217 || 'RUB'
  ];
}

function priceHistoryNum_(v) {
  if (v === undefined || v === null || v === '') return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function writePriceHistory_(dateStr, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRICE_HISTORY_SHEET_);
  if (!sheet) sheet = ss.insertSheet(PRICE_HISTORY_SHEET_);
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);

  var headers = PRICE_HISTORY_HEADERS_;
  var migrated = priceHistoryMigrate_(sheet);

  var removed = stocksRemoveDate_(sheet, headers, dateStr);
  var startRow = 0;
  if (rows.length) {
    startRow = sheet.getLastRow() + 1;
    var b, batch;
    for (b = 0; b < rows.length; b += 1000) {
      batch = rows.slice(b, b + 1000);
      sheet.getRange(startRow + b, 1, batch.length, headers.length).setValues(batch);
    }
  }
  if (migrated && sheet.getLastRow() > 1) priceHistoryFormat_(sheet, 2, sheet.getLastRow() - 1);
  else if (startRow) priceHistoryFormat_(sheet, startRow, rows.length);
  wbPutSheetFilter_(sheet);
  return { rows: rows.length, removed: removed };
}

function priceHistoryHeadersMatch_(sheet) {
  var headers = PRICE_HISTORY_HEADERS_;
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < headers.length) return false;
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var h;
  for (h = 0; h < headers.length; h++) {
    if (String(current[h] || '').trim() !== headers[h]) return false;
  }
  if (sheet.getLastColumn() > headers.length) {
    var extra = String(sheet.getRange(1, headers.length + 1).getValue() || '').trim();
    if (extra) return false;
  }
  return true;
}

/** Старый лист: «Дата цены» и столбец «Загружено». Данные сохраняются. */
function priceHistoryMigrate_(sheet) {
  var headers = PRICE_HISTORY_HEADERS_;
  if (priceHistoryHeadersMatch_(sheet)) return false;
  if (!sheet || sheet.getLastRow() <= 0 || sheet.getLastColumn() <= 0) return false;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (!data || !data.length || !data[0]) return false;
  var head = [];
  var h;
  for (h = 0; h < data[0].length; h++) head.push(String(data[0][h] || '').trim());
  function col(names) {
    var n, i;
    for (n = 0; n < names.length; n++) {
      i = head.indexOf(names[n]);
      if (i >= 0) return i;
    }
    return -1;
  }
  var map = [
    col(['Дата', 'Дата цены']),
    col(['Артикул WB']),
    col(['Арт. продавца', 'Артикул продавца']),
    col(['Название']),
    col(['Цена, ₽', 'Цена']),
    col(['Цена со скидкой, ₽']),
    col(['Цена клуба, ₽', 'Цена WB Клуба, ₽']),
    col(['Скидка прод. %', 'Скидка %']),
    col(['Валюта'])
  ];
  var out = [headers];
  var r, c, v, row, empty;
  for (r = 1; r < data.length; r++) {
    row = [];
    empty = true;
    for (c = 0; c < map.length; c++) {
      v = map[c] >= 0 ? data[r][map[c]] : '';
      if (c === 0 && v !== '' && v != null) v = lkStocksDate_(v);
      if (v !== '' && v != null) empty = false;
      row.push(v == null ? '' : v);
    }
    if (!empty) out.push(row);
  }
  sheet.clear();
  sheet.getRange(1, 1, out.length, headers.length).setValues(out);
  return true;
}

/** Формат только новых строк. Дата уже записана как полночь, столбец целиком не переписываем. */
function priceHistoryFormat_(sheet, startRow, rowCount) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var headers = PRICE_HISTORY_HEADERS_;
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#ffffff');
  stocksSafeFreezeRows_(sheet, 1);
  if (!startRow || startRow < 2 || !rowCount || rowCount < 1) return;
  sheet.getRange(startRow, 1, rowCount, headers.length).setFontFamily('Arial').setFontSize(10);
  sheet.getRange(startRow, 1, rowCount, 1).setNumberFormat('dd.mm.yyyy');
  var priceCol = headers.indexOf('Цена, ₽') + 1;
  var discCol = headers.indexOf('Скидка прод. %') + 1;
  if (priceCol > 0) sheet.getRange(startRow, priceCol, rowCount, 3).setNumberFormat('#,##0.00');
  if (discCol > 0) sheet.getRange(startRow, discCol, rowCount, 1).setNumberFormat('0');
  priceHistoryFillTitles_(sheet, startRow, rowCount);
}

function priceHistoryFillTitles_(sheet, startRow, rowCount) {
  var headers = PRICE_HISTORY_HEADERS_;
  var nmCol = headers.indexOf('Артикул WB') + 1;
  var titleCol = headers.indexOf('Название') + 1;
  if (nmCol < 1 || titleCol < 1 || !startRow || startRow < 2 || !rowCount || rowCount < 1) return;
  var nms = sheet.getRange(startRow, nmCol, rowCount, 1).getValues();
  var titles = sheet.getRange(startRow, titleCol, rowCount, 1).getValues();
  var dict = getProductsDictionary();
  var changed = false;
  var i, title, product;
  for (i = 0; i < titles.length; i++) {
    title = String(titles[i][0] || '').trim();
    if (title) continue;
    product = lkProduct_(dict, nms[i][0]);
    if (!product.title) continue;
    titles[i][0] = product.title;
    changed = true;
  }
  if (changed) sheet.getRange(startRow, titleCol, rowCount, 1).setValues(titles);
}

// Функция для тестирования API цен (для отладки)
function testPricesAPI() {
  var token = getApiToken();
  if (!token) {
    SpreadsheetApp.getUi().alert('❌ Токен не найден');
    return;
  }
  
  var url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=10&offset=0';
  var options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  
  // Парсим и показываем понятную информацию
  if (data.data && data.data.listGoods) {
    var message = '✅ API работает!\n\n';
    message += 'Найдено товаров: ' + data.data.listGoods.length + '\n\n';
    message += 'Пример первого товара:\n';
    var first = data.data.listGoods[0];
    if (first && first.sizes && first.sizes[0]) {
      message += 'Артикул WB: ' + first.nmID + '\n';
      message += 'Цена: ' + first.sizes[0].price + ' ₽\n';
      message += 'Цена со скидкой: ' + first.sizes[0].discountedPrice + ' ₽\n';
      message += 'Скидка: ' + first.discount + '%\n';
    }
    SpreadsheetApp.getUi().alert(message);
  } else {
    SpreadsheetApp.getUi().alert('❌ Неожиданный ответ API:\n' + JSON.stringify(data).substring(0, 500));
  }
}

// 9. ЕЖЕНЕДЕЛЬНАЯ ДИНАМИКА
function showWeekPicker() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;height:200px;margin:0}' +
    'label{display:block;font-weight:bold;margin-bottom:5px}' +
    'input{width:100%;padding:6px;margin-bottom:15px}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '.info{background:#e3f2fd;padding:8px;margin-bottom:10px;border-radius:4px;font-size:11px}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📈 Еженедельная динамика</h3>' +
    '<div class="info">📌 Выберите любой день недели</div>' +
    '<label>Выберите дату:</label><input type="date" id="selectedDate">' +
    '<button id="btn" onclick="loadDynamics()">📊 Сформировать</button>' +
    '<script>' +
    'function loadDynamics(){' +
    'var d=document.getElementById("selectedDate").value;' +
    'if(!d){alert("Выберите дату!");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);})' +
    '.calculateWeeklyDynamics(d);' +
    '}' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Еженедельная динамика');
}

function calculateWeeklyDynamics(selectedDateStr) {
  var token = getApiToken();
  if (!token) return;

  var selectedDate = new Date(selectedDateStr);
  var dayOfWeek = selectedDate.getDay();
  var daysToMonday = (dayOfWeek === 0 ? 6 : dayOfWeek - 1);
  
  var weekStart = new Date(selectedDate);
  weekStart.setDate(selectedDate.getDate() - daysToMonday);
  var weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  var weekNumber = getWeekNumber(weekStart);
  
  var currentData = fetchDataForPeriod(token, weekStart, weekEnd);
  var currentSales = aggregateSales(currentData);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Еженедельная динамика';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 1).setValue('Показатель');
    sheet.getRange(1, 1, 1, 1).setFontWeight('bold');
  }
  
  var lastCol = sheet.getLastColumn();
  var newCol = lastCol + 1;
  
  var startDateStr = formatDate(weekStart);
  var endDateStr = formatDate(weekEnd);
  var headerText = 'Неделя ' + weekNumber + ' (' + startDateStr + ' — ' + endDateStr + ')';
  
  if (lastCol === 1) {
    sheet.getRange(1, 2, 1, 1).setValue('Текущая неделя');
    sheet.getRange(1, 3, 1, 1).setValue('Динамика (%)');
    sheet.getRange(1, 2, 1, 2).setFontWeight('bold');
  } else {
    sheet.insertColumnAfter(lastCol);
    newCol = lastCol + 1;
    sheet.getRange(1, newCol, 1, 1).setValue(headerText);
    sheet.getRange(1, newCol, 1, 1).setFontWeight('bold');
    
    sheet.insertColumnAfter(newCol);
    var dynamicCol = newCol + 1;
    sheet.getRange(1, dynamicCol, 1, 1).setValue('Динамика (%)');
    sheet.getRange(1, dynamicCol, 1, 1).setFontWeight('bold');
    
    var prevCol = newCol - 2;
    var prevValues = sheet.getRange(2, prevCol, 5, 1).getValues();
    
    var metricsVals = [
      currentSales.revenue,
      currentSales.quantity,
      currentSales.commission,
      currentSales.logistics,
      currentSales.netProfit
    ];
    
    for (var i = 0; i < 5; i++) {
      var prev = prevValues[i][0] || 0;
      var current = metricsVals[i];
      var percent = prev !== 0 ? ((current - prev) / prev * 100).toFixed(2) : (current !== 0 ? 100 : 0);
      sheet.getRange(i + 2, dynamicCol, 1, 1).setValue(percent + '%');
    }
    
    newCol = dynamicCol;
  }
  
  var metrics = [
    'Выручка (руб)',
    'Количество продаж (шт)',
    'Комиссия WB (руб)',
    'Логистика (руб)',
    'Чистая прибыль (руб)'
  ];
  
  if (sheet.getLastRow() < 2) {
    for (var i = 0; i < metrics.length; i++) {
      sheet.getRange(i + 2, 1, 1, 1).setValue(metrics[i]);
    }
    sheet.getRange(2, 1, 5, 1).setFontWeight('bold');
  }
  
  var values = [
    currentSales.revenue,
    currentSales.quantity,
    currentSales.commission,
    currentSales.logistics,
    currentSales.netProfit
  ];
  
  for (var i = 0; i < values.length; i++) {
    sheet.getRange(i + 2, newCol, 1, 1).setValue(values[i]);
  }
  
  sheet.autoResizeColumns(1, newCol);
  sheet.getRange(2, 2, 5, newCol).setNumberFormat('# ##0.00');
  
  if (!isRunningFromTrigger()) {
    SpreadsheetApp.getUi().alert('✅ Динамика добавлена!\n\nНеделя ' + weekNumber + ' (' + startDateStr + ' — ' + endDateStr + ')');
  }
  Logger.log('Динамика добавлена: Неделя ' + weekNumber);
}

function getWeekNumber(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function fetchDataForPeriod(token, dateFrom, dateTo) {
  var dateFromStr = Utilities.formatDate(dateFrom, 'GMT+3', 'yyyy-MM-dd');
  var dateToStr = Utilities.formatDate(dateTo, 'GMT+3', 'yyyy-MM-dd');
  // Старый GET v5 (не вызывается):
  // var endpoint = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod';
  // while (true) {
  //   var url = endpoint + '?dateFrom=' + dateFromStr + '&dateTo=' + dateToStr;
  //   if (nextParam) url += '&' + nextParam;
  //   var response = UrlFetchApp.fetch(url, { method: 'GET', headers: { Authorization: token }, muteHttpExceptions: true });
  //   if (response.getResponseCode() === 204) break;
  //   if (response.getResponseCode() !== 200) break;
  //   var data = JSON.parse(response.getContentText());
  //   if (Array.isArray(data) && data.length > 0) {
  //     allData = allData.concat(data);
  //     var rrdid = data[data.length - 1].rrd_id;
  //     if (rrdid) { nextParam = 'rrdid=' + rrdid; Utilities.sleep(500); continue; }
  //   }
  //   break;
  // }
  try {
    return fetchFinanceDetailedV1_(token, dateFromStr, dateToStr, 'daily');
  } catch (e) {
    Logger.log('fetchDataForPeriod: ' + e);
    return [];
  }
}

function aggregateSales(data) {
  var result = { revenue: 0, quantity: 0, commission: 0, logistics: 0, netProfit: 0 };
  
  for (var i = 0; i < data.length; i++) {
    var item = data[i];
    if (item.doc_type_name === 'Продажа' || item.supplier_oper_name === 'Продажа') {
      result.revenue += parseFloat(item.retail_price_withdisc_rub || 0);
      result.quantity += parseFloat(item.quantity || 0);
      result.commission += parseFloat(item.ppvz_sales_commission || 0);
      result.logistics += parseFloat(item.delivery_rub || 0);
    }
  }
  result.netProfit = result.revenue - result.commission - result.logistics;
  return result;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'GMT+3', 'yyyy-MM-dd');
}

/** Ключ даты для антидублей рекламы: Date / ISO / дд.мм.гггг → yyyy-MM-dd */
function adsDateKey_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'GMT+3', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var iso = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var dmy = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (dmy) {
    return dmy[3] + '-' + ('0' + dmy[2]).slice(-2) + '-' + ('0' + dmy[1]).slice(-2);
  }
  return s.substring(0, 10);
}


function addDataToSheet(newData, sheetName, uniqueField) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var existingIds = {};
  
  // Функция для нахождения последней реально заполненной строки
  function getLastRealRow(sheet, column) {
    var col = column || 1; // по умолчанию колонка A
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 1;
    
    // Проверяем с конца до начала, ищем непустую ячейку
    for (var i = lastRow; i >= 1; i--) {
      var value = sheet.getRange(i, col, 1, 1).getValue();
      if (value !== null && value !== undefined && value !== '') {
        return i;
      }
    }
    return 1;
  }
  
  if (sheet) {
    var lastRow = getLastRealRow(sheet, 1);
    if (lastRow > 1) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var columnIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [uniqueField]) : headers.indexOf(uniqueField);
      
      if (columnIndex !== -1) {
        var existingData = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < existingData.length; i++) {
          var idValue = existingData[i][0];
          if (idValue) {
            existingIds[idValue.toString()] = true;
          }
        }
      }
    }
  } else {
    sheet = ss.insertSheet(sheetName);
    var headers = Object.keys(newData[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  var uniqueData = [];
  for (var i = 0; i < newData.length; i++) {
    var item = newData[i];
    var itemId = typeof wbPickObj_ === 'function' ? wbPickObj_(item, [uniqueField]) : item[uniqueField];
    if (itemId && !existingIds[itemId.toString()]) {
      uniqueData.push(item);
      existingIds[itemId.toString()] = true;
    }
  }
  
  if (uniqueData.length === 0) {
    if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('⚠️ Дубликаты не добавлены');
    wbPutSheetFilter_(sheet);
    return;
  }
  
  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = [];
  for (var i = 0; i < uniqueData.length; i++) {
    if (typeof wbFillRow_ === 'function') {
      rows.push(wbFillRow_(uniqueData[i], currentHeaders));
      continue;
    }
    var row = [];
    for (var j = 0; j < currentHeaders.length; j++) {
      var value = uniqueData[i][currentHeaders[j]];
      if (value === null || value === undefined) row.push('');
      else if (typeof value === 'object' && !(value instanceof Date)) row.push(JSON.stringify(value));
      else row.push(value);
    }
    rows.push(row);
  }
  
  if (rows.length > 0) {
    // Ищем последнюю реально заполненную строку
    var lastRealRow = getLastRealRow(sheet, 1);
    var startRow = lastRealRow + 1;
    
    // Если после последней реальной строки есть пустые строки с форматированием
    // удаляем их, чтобы не было разрывов
    var maxRow = sheet.getLastRow();
    if (maxRow > startRow) {
      sheet.deleteRows(startRow, maxRow - startRow + 1);
    }
    
    sheet.getRange(startRow, 1, rows.length, currentHeaders.length).setValues(rows);
  }
  
  if (!isRunningFromTrigger()) SpreadsheetApp.getUi().alert('✅ Добавлено: ' + uniqueData.length + ' записей');
  Logger.log('Добавлено: ' + uniqueData.length + ' записей в лист ' + sheetName);
  wbPutSheetFilter_(sheet);
}

function cleanEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Рекламные_расходы');
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  var deleted = 0;
  
  // Удаляем пустые строки между данными
  for (var i = lastRow; i >= 2; i--) {
    var range = sheet.getRange(i, 1, 1, sheet.getLastColumn());
    var values = range.getValues()[0];
    var isEmpty = true;
    for (var j = 0; j < values.length; j++) {
      if (values[j] !== null && values[j] !== undefined && values[j] !== '') {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) {
      sheet.deleteRow(i);
      deleted++;
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ Удалено пустых строк: ' + deleted);
}

// 12. УДАЛЕНИЕ ДУБЛИКАТОВ
function removeDuplicatesFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var response = ui.alert('Удаление дубликатов', 'Удалить дубликаты в листе "' + sheet.getName() + '"?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  var headers = data[0];
  var uniqueRows = {};
  var newData = [headers];
  
  for (var i = 1; i < data.length; i++) {
    var key = JSON.stringify(data[i]);
    if (!uniqueRows[key]) {
      uniqueRows[key] = true;
      newData.push(data[i]);
    }
  }
  
  var duplicatesCount = data.length - newData.length;
  if (duplicatesCount === 0) {
    ui.alert('✅ Дубликатов не найдено');
    return;
  }
  
  sheet.clear();
  sheet.getRange(1, 1, newData.length, headers.length).setValues(newData);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  ui.alert('✅ Удалено дубликатов: ' + duplicatesCount);
}

function протянутьФормулыЗаказыЛК() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "ЗаказыЛК";
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ Лист "' + sheetName + '" не найден');
    return;
  }
  
  // Снимаем фильтр
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  
  // Находим последнюю строку по столбцу A
  const colA = sheet.getRange("A:A").getValues();
  let lastRow = 0;
  for (let i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== "" && colA[i][0] !== null) {
      lastRow = i + 1;
      break;
    }
  }
  
  if (lastRow < 3) {
    wbPutSheetFilter_(sheet, 2);
    SpreadsheetApp.getUi().alert('⚠️ Нет данных для протягивания формул');
    return;
  }
  
  // Копируем формулу из I1 в диапазон I3:I{lastRow}
  const sourceRangeI = sheet.getRange("I1");
  const targetRangeI = sheet.getRange(3, 9, lastRow - 2, 1); // колонка I = 9
  sourceRangeI.copyTo(targetRangeI, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  
  // Копируем формулу из J1 в диапазон J3:J{lastRow}
  const sourceRangeJ = sheet.getRange("J1");
  const targetRangeJ = sheet.getRange(3, 10, lastRow - 2, 1); // колонка J = 10
  sourceRangeJ.copyTo(targetRangeJ, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  
  // Копируем формулу из K1 в диапазон K3:K{lastRow}
  const sourceRangeK = sheet.getRange("K1");
  const targetRangeK = sheet.getRange(3, 11, lastRow - 2, 1); // колонка K = 11
  sourceRangeK.copyTo(targetRangeK, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  // Шапка на строке 2: в строке 1 лежат формулы-образцы.
  wbPutSheetFilter_(sheet, 2);
  
  SpreadsheetApp.getUi().alert('✅ Формулы протянуты в листе "' + sheetName + '" до строки ' + lastRow);
}

// // ============================================
// // ЗАКАЗЫ - РАБОЧАЯ ВЕРСИЯ (statistics-api)
// // ============================================

// // ДИАЛОГ ВЫБОРА ПЕРИОДА

function showOrdersDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;height:200px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📦 Загрузка заказов за период</h3>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom" value="2026-04-01">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo" value="2026-04-01">' +
    '<button id="btn" onclick="loadOrders()">📥 Загрузить</button>' +
    '<script>' +
    'function loadOrders(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'if(!d1||!d2){alert("Выберите даты!");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);})' +
    '.loadOrdersRangeManual(d1,d2);' +
    '}' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Загрузка заказов');
}

// ФУНКЦИЯ ДЛЯ РУЧНОЙ ЗАГРУЗКИ ЗА ПЕРИОД (с расчетными столбцами)
function loadOrdersRangeStat(dateFromStr, dateToStr) {
  loadOrdersRangeManual(dateFromStr, dateToStr);
}

function loadOrdersRangeManual(dateFromStr, dateToStr) {
  if (typeof renameWbColumnHeaders_ === 'function') renameWbColumnHeaders_();
  var token = getApiToken();
  if (!token) {
    SpreadsheetApp.getUi().alert('⚠️ Токен не найден');
    return;
  }
  
  var startDate = new Date(dateFromStr);
  var endDate = new Date(dateToStr);
  var current = new Date(startDate);
  var allOrders = [];
  var ui = SpreadsheetApp.getUi();
  
  ui.alert('📦 Загрузка заказов\n\nПериод: ' + dateFromStr + ' — ' + dateToStr + '\n\nЗагрузка может занять время...');
  
  while (current <= endDate) {
    var year = current.getFullYear();
    var month = String(current.getMonth() + 1).padStart(2, '0');
    var day = String(current.getDate()).padStart(2, '0');
    var dateStr = year + '-' + month + '-' + day;
    
    var url = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=' + dateStr + '&flag=1';
    
    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: { 'Authorization': token },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        var orders = JSON.parse(response.getContentText());
        if (orders && orders.length > 0) {
          allOrders = allOrders.concat(orders);
          console.log(dateStr + ': ' + orders.length + ' заказов');
        }
      }
    } catch(e) {
      console.log('Ошибка ' + dateStr + ': ' + e);
    }
    
    current.setDate(current.getDate() + 1);
    Utilities.sleep(1000);
  }
  
  if (allOrders.length === 0) {
    ui.alert('⚠️ Заказов не найдено за период ' + dateFromStr + ' — ' + dateToStr);
    return;
  }
  
  // Добавляем расчетные столбцы
  var enhancedOrders = addOrderCalculatedColumns(allOrders);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Заказы');
  
  if (!sheet) {
    sheet = ss.insertSheet('Заказы');
  }
  
  // Сохраняем с проверкой на дубли
  saveOrdersToSheet(sheet, enhancedOrders);
  
  ui.alert('✅ Добавлено ' + enhancedOrders.length + ' заказов за период ' + dateFromStr + ' — ' + dateToStr);
}

// ФУНКЦИЯ ДЛЯ АВТООБНОВЛЕНИЯ (для триггера - только за вчера)
function autoLoadYesterdayOrders() {
  var token = getApiToken();
  if (!token) {
    console.log('Токен не найден');
    return;
  }
  
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var year = yesterday.getFullYear();
  var month = String(yesterday.getMonth() + 1).padStart(2, '0');
  var day = String(yesterday.getDate()).padStart(2, '0');
  var dateStr = year + '-' + month + '-' + day;
  
  console.log('Автозагрузка заказов за: ' + dateStr);
  
  var url = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=' + dateStr + '&flag=1';
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    console.log('Ошибка API: ' + response.getResponseCode());
    return;
  }
  
  var orders = JSON.parse(response.getContentText());
  if (!orders || orders.length === 0) {
    console.log('Нет заказов за ' + dateStr);
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Заказы');
  var existingIds = {};
  
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idColIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['srid', 'id', 'gNumber']) : headers.indexOf('id');
      if (idColIndex !== -1) {
        var existingData = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < existingData.length; i++) {
          if (existingData[i][0]) {
            existingIds[existingData[i][0].toString()] = true;
          }
        }
      }
    }
  } else {
    sheet = ss.insertSheet('Заказы');
  }
  
  var newOrders = [];
  for (var i = 0; i < orders.length; i++) {
    var orderId = orders[i].srid || orders[i].id || orders[i].gNumber;
    if (orderId && !existingIds[orderId.toString()]) {
      newOrders.push(orders[i]);
    }
  }
  
  if (newOrders.length === 0) {
    console.log('Новых заказов за ' + dateStr + ' нет');
    return;
  }
  
  var enhancedOrders = addOrderCalculatedColumns(newOrders);
  saveOrdersToSheet(sheet, enhancedOrders);
  
  console.log('✅ Добавлено ' + newOrders.length + ' новых заказов за ' + dateStr);
}

// ФУНКЦИЯ СОХРАНЕНИЯ (общая для всех)
function saveOrdersToSheet(sheet, enhancedOrders) {
  if (!enhancedOrders || enhancedOrders.length === 0) return;
  
  var newHeaders = Object.keys(enhancedOrders[0]);
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // var requiredColumns = ['Артикул', 'Дата', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  var requiredColumns = ['Артикул', 'Дата', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма', 'Регион поставки'];
  var lastCol = sheet.getLastColumn();
  
  for (var i = 0; i < requiredColumns.length; i++) {
    if (sheetHeaders.indexOf(requiredColumns[i]) === -1) {
      sheet.getRange(1, lastCol + 1).setValue(requiredColumns[i]);
      sheet.getRange(1, lastCol + 1).setFontWeight('bold');
      lastCol++;
      sheetHeaders.push(requiredColumns[i]);
    }
  }
  
  var rows = [];
  for (var i = 0; i < enhancedOrders.length; i++) {
    var row = [];
    for (var j = 0; j < sheetHeaders.length; j++) {
      var val = typeof wbCell_ === 'function' ? wbCell_(enhancedOrders[i], sheetHeaders[j]) : enhancedOrders[i][sheetHeaders[j]];
      row.push(val === null || val === undefined ? '' : val);
    }
    rows.push(row);
  }
  
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, sheetHeaders.length).setValues(rows);
  
  var dateColIndex = sheetHeaders.indexOf('Дата') + 1;
  if (dateColIndex > 0 && sheet.getLastRow() > 1) {
    sheet.getRange(2, dateColIndex, sheet.getLastRow() - 1, 1).setNumberFormat("dd.mm.yyyy");
  }
  
  var numberCols = ['Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  for (var i = 0; i < numberCols.length; i++) {
    var colIndex = sheetHeaders.indexOf(numberCols[i]) + 1;
    if (colIndex > 0 && sheet.getLastRow() > 1) {
      sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00");
    }
  }
  wbPutSheetFilter_(sheet);
}


// // ============================================
// // ФУНКЦИЯ ДЛЯ РАСЧЕТА НОВЫХ СТОЛБЦОВ (ВСТАВЬТЕ ЭТУ ФУНКЦИЮ)
// // ============================================

function addOrderCalculatedColumns(orders) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // === 1. Загружаем артикулы из листа "Товары" ===
  var nmIdToArticle = {};
  var productsSheet = ss.getSheetByName('Товары');
  if (productsSheet) {
    var productsData = productsSheet.getDataRange().getValues();
    if (productsData && productsData.length > 1) {
      for (var i = 1; i < productsData.length; i++) {
        var nmId = productsData[i][0];
        var article = productsData[i][1];
        if (nmId) nmIdToArticle[nmId.toString()] = article || '';
      }
    }
  }
  
  // === 2. Загружаем веса из листа "Товар+Вес" ===
  var weightData = {};
  var weightSheet = ss.getSheetByName('Товар+Вес');
  if (weightSheet) {
    var weightSheetData = weightSheet.getDataRange().getValues();
    if (weightSheetData && weightSheetData.length > 0) {
      var headers = weightSheetData[0];
      var weightColIndex = headers.indexOf('Вес, кг за 1 шт');
      if (weightColIndex !== -1) {
        for (var i = 1; i < weightSheetData.length; i++) {
          var supplierArticle = weightSheetData[i][0];
          if (supplierArticle) {
            var key = supplierArticle.toString().trim().toLowerCase();
            var weight = parseFloat(weightSheetData[i][weightColIndex]);
            weightData[key] = isNaN(weight) ? 0 : weight;
          }
        }
      }
    }
  }
  
  // === 3. Загружаем регионы из листа "Регион поставки" ===
  var warehouseToRegion = {};
  var regionSheet = ss.getSheetByName('Регион поставки');
  
  if (regionSheet) {
    var regionData = regionSheet.getDataRange().getValues();
    if (regionData && regionData.length > 1) {
      var rHeaders = regionData[0];
      var warehouseColIndex = rHeaders.indexOf('Склад');
      var regionColIndex = rHeaders.indexOf('Регион поставки');
      
      if (warehouseColIndex !== -1 && regionColIndex !== -1) {
        for (var i = 1; i < regionData.length; i++) {
          var warehouseName = regionData[i][warehouseColIndex];
          var region = regionData[i][regionColIndex];
          if (warehouseName) {
            var key = warehouseName.toString().trim().toLowerCase();
            warehouseToRegion[key] = region || '';
          }
        }
      }
    }
  }
  
  // === Вспомогательные функции ===
  
  // 🔥 ИЗМЕНЕНО: Теперь возвращает объект Date, а не строку
  function parseDate(value) {
    if (!value) return null;
    
    // Если уже объект Date
    if (value instanceof Date) return value;
    
    var str = value.toString();
    
    // Формат ISO: "2026-03-31T07:07:39" или "2026-03-31 7:07:39"
    // Заменяем пробел на T для надёжности парсинга, если нужно
    var isoStr = str.replace(' ', 'T');
    
    // Пробуем создать дату
    var d = new Date(isoStr);
    
    // Проверка на валидность даты
    if (isNaN(d.getTime())) {
      // Если не вышло, пробуем другой формат (например, DD.MM.YYYY)
      // Но API WB обычно отдаёт ISO или похожий формат
      return null; 
    }
    
    return d;
  }
  
  function getWeight(supplierArticle) {
    if (!supplierArticle) return 0;
    var key = supplierArticle.toString().trim().toLowerCase();
    return weightData[key] || 0;
  }
  
  function getRegion(warehouseName) {
    if (!warehouseName) return '';
    var key = warehouseName.toString().trim().toLowerCase();
    return warehouseToRegion[key] || '';
  }
  
  // === Обрабатываем каждый заказ ===
  var result = [];
  for (var i = 0; i < orders.length; i++) {
    var order = orders[i];
    
    var nmId = order.nmId || order.nm_id || '';
    
    // 🔥 Получаем дату как объект Date
    var dateValue = order.date || order.dt || order.create_dt || order.createdAt || '';
    var parsedDate = parseDate(dateValue);
    
    var supplierArticle = order.supplierArticle || order.sa_name || '';
    var quantity = order.quantity || order.qty || 1;
    var warehouseName = order.warehouseName || order.warehouse_name || order.warehouse || '';
    
    var article = nmIdToArticle[nmId.toString()] || '';
    var weightPerUnit = getWeight(supplierArticle);
    var totalWeight = weightPerUnit * quantity;
    var region = getRegion(warehouseName);
    
    var enhanced = {};
    for (var key in order) {
      enhanced[key] = order[key];
    }
    
    // Добавляем расчетные столбцы
    enhanced['Артикул'] = article;
    
    // 🔥 Записываем объект Date. Google Таблицы сами отформатируют его, 
    // если в saveOrdersToSheet задан setNumberFormat
    // enhanced['Дата'] = parsedDate; 
    enhanced['Дата'] = formatDateRU(parsedDate); 
    
    enhanced['Вес, кг за 1 шт'] = weightPerUnit;
    enhanced['Вес, кг за 1 шт сумма'] = totalWeight;
    enhanced['Регион поставки'] = region;
    
    if (typeof wbRowToRu_ === 'function') {
      enhanced = wbRowToRu_(enhanced, typeof WB_RU_ORD_EXTRA !== 'undefined' ? WB_RU_ORD_EXTRA : { date: 'Дата-время API' });
    }
    result.push(enhanced);
  }
  
  return result;
}

// ============================================
// ИСПРАВЛЕННАЯ ФУНКЦИЯ АВТООБНОВЛЕНИЯ (ЗАМЕНИТЕ СТАРУЮ)
// ============================================
// ИСПРАВЛЕННАЯ ВЕРСИЯ - добавляет заказы с расчетными столбцами
function loadOrdersForYesterdayAuto() {
  if (typeof renameWbColumnHeaders_ === 'function') renameWbColumnHeaders_();
  var token = getApiToken();
  if (!token) {
    console.log('Токен не найден');
    return;
  }
  
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var year = yesterday.getFullYear();
  var month = String(yesterday.getMonth() + 1).padStart(2, '0');
  var day = String(yesterday.getDate()).padStart(2, '0');
  var dateStr = year + '-' + month + '-' + day;
  
  console.log('Загружаем заказы за ' + dateStr);
  
  var url = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=' + dateStr + '&flag=1';
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    console.log('Ошибка API: ' + response.getResponseCode());
    return;
  }
  
  var orders = JSON.parse(response.getContentText());
  if (!orders || !orders.length) {
    console.log('Нет заказов за ' + dateStr);
    return;
  }

  console.log('API вернул ' + orders.length + ' заказов');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Заказы');
  if (!sheet) sheet = ss.insertSheet('Заказы');

  var existingIds = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idColIndex = typeof wbColIdx_ === 'function'
      ? wbColIdx_(headers, ['srid', 'id', 'gNumber'])
      : headers.indexOf('srid');
    if (idColIndex < 0) idColIndex = headers.indexOf('id');
    if (idColIndex >= 0) {
      var existingData = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existingData.length; i++) {
        if (existingData[i][0]) existingIds[String(existingData[i][0])] = true;
      }
    }
  }

  var newOrders = [];
  for (var o = 0; o < orders.length; o++) {
    var orderId = orders[o].srid || orders[o].id || orders[o].gNumber;
    if (orderId && !existingIds[String(orderId)]) {
      newOrders.push(orders[o]);
      existingIds[String(orderId)] = true;
    }
  }

  if (!newOrders.length) {
    console.log('Новых заказов за ' + dateStr + ' нет (дубли отсеяны)');
    return;
  }

  var enhancedOrders = addOrderCalculatedColumns(newOrders);
  saveOrdersToSheet(sheet, enhancedOrders);
  console.log('✅ Добавлено ' + newOrders.length + ' новых заказов за ' + dateStr);
}

// ============================================
// ФУНКЦИЯ СОХРАНЕНИЯ ОБОГАЩЕННЫХ ЗАКАЗОВ В ЛИСТ
// ============================================
function saveEnhancedOrdersToSheet(sheet, enhancedOrders) {
  if (!enhancedOrders || enhancedOrders.length === 0) return;
  
  // Получаем заголовки из первого обогащенного заказа
  var newHeaders = Object.keys(enhancedOrders[0]);
  
  // Если лист пустой — создаем заголовки
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  // Получаем текущие заголовки листа
  var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Добавляем недостающие столбцы (Артикул, Дата, Вес, кг за 1 шт, Вес, кг за 1 шт сумма)
  var requiredColumns = ['Артикул', 'Дата', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  var lastCol = sheet.getLastColumn();
  
  for (var i = 0; i < requiredColumns.length; i++) {
    if (sheetHeaders.indexOf(requiredColumns[i]) === -1) {
      sheet.getRange(1, lastCol + 1).setValue(requiredColumns[i]);
      sheet.getRange(1, lastCol + 1).setFontWeight('bold');
      lastCol++;
      sheetHeaders.push(requiredColumns[i]);
    }
  }
  
  // Формируем строки для записи
  var rows = [];
  for (var i = 0; i < enhancedOrders.length; i++) {
    var row = [];
    for (var j = 0; j < sheetHeaders.length; j++) {
      var val = typeof wbCell_ === 'function' ? wbCell_(enhancedOrders[i], sheetHeaders[j]) : enhancedOrders[i][sheetHeaders[j]];
      if (val === null || val === undefined) {
        row.push('');
      } else if (typeof val === 'object' && !(val instanceof Date)) {
        row.push(JSON.stringify(val));
      } else {
        row.push(val);
      }
    }
    rows.push(row);
  }
  
  // Добавляем в конец листа
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, sheetHeaders.length).setValues(rows);
  
  // Форматируем столбец "Дата"
  var dateColIndex = sheetHeaders.indexOf('Дата') + 1;
  if (dateColIndex > 0 && sheet.getLastRow() > 1) {
    sheet.getRange(2, dateColIndex, sheet.getLastRow() - 1, 1).setNumberFormat("dd.mm.yyyy");
  }
  
  // Форматируем числовые столбцы
  var numberCols = ['Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  for (var i = 0; i < numberCols.length; i++) {
    var colIndex = sheetHeaders.indexOf(numberCols[i]) + 1;
    if (colIndex > 0 && sheet.getLastRow() > 1) {
      sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00");
    }
  }
  
  sheet.autoResizeColumns(1, sheetHeaders.length);
  console.log('Сохранено заказов: ' + rows.length);
  wbPutSheetFilter_(sheet);
}

// 🔥 Форматирует объект Date в строку "дд.мм.гггг"
function formatDateRU(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  const day = String(dateObj.getDate()).padStart(2, '0');
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const year = dateObj.getFullYear();
  return `${day}.${month}.${year}`;
}

// ============================================
// ТЕСТОВАЯ ФУНКЦИЯ (ЗАПУСТИТЕ ДЛЯ ПРОВЕРКИ)
// ============================================
function testLoadOrdersForYesterdayAuto() {
  loadOrdersForYesterdayAuto();
}




function testAllOrderMethods() {
  const token = getApiToken();
  if (!token) {
    SpreadsheetApp.getUi().alert('Токен не найден');
    return;
  }

  const methods = [
    { name: '1. FBS Новые (v3)', url: 'https://marketplace-api.wildberries.ru/api/v3/orders/new', method: 'GET' },
    { name: '2. FBS Все заказы (v3)', url: 'https://marketplace-api.wildberries.ru/api/v3/orders?limit=10', method: 'GET' },
    { name: '3. FBS заказы за сегодня (v3)', url: 'https://marketplace-api.wildberries.ru/api/v3/orders?dateFrom=' + Math.floor(new Date().setHours(0,0,0,0)/1000) + '&limit=10', method: 'GET' },
    { name: '4. DBS Новые (v3)', url: 'https://marketplace-api.wildberries.ru/api/v3/dbs/orders/new', method: 'GET' },
    { name: '5. DBS Все заказы (v3)', url: 'https://marketplace-api.wildberries.ru/api/v3/dbs/orders?limit=10', method: 'GET' },
    { name: '6. DBW Новые (Курьером WB)', url: 'https://marketplace-api.wildberries.ru/api/v3/dbw/orders/new', method: 'GET' },
    { name: '7. DBW Завершённые (Курьером WB)', url: 'https://marketplace-api.wildberries.ru/api/v3/dbw/orders?limit=10', method: 'GET' },
    { name: '8. Самовывоз Новые', url: 'https://marketplace-api.wildberries.ru/api/v3/click-collect/orders/new', method: 'GET' },
    { name: '9. Статистика заказов (flag=1)', url: 'https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=2026-04-01&flag=1', method: 'GET' }
  ];

  let results = [];
  
  for (let m of methods) {
    try {
      const response = UrlFetchApp.fetch(m.url, {
        method: m.method,
        headers: { 'Authorization': token },
        muteHttpExceptions: true
      });
      results.push(`${m.name}: статус ${response.getResponseCode()}`);
    } catch(e) {
      results.push(`${m.name}: ошибка - ${e.toString()}`);
    }
  }
  
  SpreadsheetApp.getUi().alert('Результаты тестирования:\n\n' + results.join('\n'));
}

// ДИАЛОГ ДЛЯ ВЫБОРА ПЕРИОДА
function showOrdersStatDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;height:200px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px}' +
    'button{background:#4CAF50;color:white;padding:8px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📦 Загрузка заказов за период</h3>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom" value="2026-03-01">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo" value="' + new Date().toISOString().slice(0,10) + '">' +
    '<button id="btn" onclick="load()">📥 Загрузить</button>' +
    '<script>' +
    'function load(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'if(!d1||!d2){alert("Выберите даты!");return;}' +
    'var btn=document.getElementById("btn"); btn.disabled=true;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){btn.disabled=false;alert(e);})' +
    '.loadOrdersRangeManual(d1,d2);' +
    '}' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Загрузка заказов');
}


// ============================================
// ОСТАТКИ - С ДОБАВЛЕНИЕМ РЕГИОНА И ВЕСОВ
// ============================================

// Основная функция
function recordCurrentStocksEnriched(skipProductsUpdate) {
  var token = getApiToken();
  if (!token) { showAlert('⚠️ Токен не найден'); return; }
  
  if (!skipProductsUpdate) {
    loadAllProducts();
    Utilities.sleep(2000);
  }
  
  var today = new Date();
  var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  var stocks = fetchStocksWithPagination(token);
  if (!stocks || stocks.length === 0) { showAlert('⚠️ Нет данных по остаткам'); return; }
  
  var productsDict = getProductsDictionary();
  var regionMap = getRegionMap();
  var weightsMap = loadWeightsMapForStocks(); // ЗАГРУЖАЕМ ВЕСА

  // Формируем строки: Дата, Артикул продавца, Название, Артикул WB, Предмет, Бренд, Размер, Склад, Остаток, Регион поставки, Вес за шт, Общий вес
  // Фильтруем товары с остатком > 0
  var rows = stocks
    .filter(function(stock) {
      return stock.amount > 0;  // ← Исключаем товары с нулевым остатком
    })
    .map(function(stock) {
      var product = productsDict[stock.nmId] || {};
      var region = getRegionFromMap(stock.warehouse, regionMap);
      
      // Получаем вес по артикулу продавца
      var vendorCode = product.vendorCode || '';
      var weightPerUnit = 0;
      if (vendorCode && weightsMap[vendorCode.toString().trim()]) {
        weightPerUnit = weightsMap[vendorCode.toString().trim()];
      }
      
      // Вычисляем общий вес (вес за шт * количество остатков)
      var totalWeight = weightPerUnit * stock.amount;
      
      return [
        dateStr,
        vendorCode,
        product.title || '',
        stock.nmId,
        product.subject || '',
        product.brand || '',
        '',  // Размер (можно будет добавить позже)
        stock.warehouse,
        stock.amount,
        region,
        weightPerUnit,   // Вес, кг за 1 шт
        totalWeight      // Общий вес, кг
      ];
    });

  var added = appendToStockSheetFull(rows);
  showAlert('✅ Остатки за ' + dateStr + ': ' + added + ' строк (день перезаписан)');
}

// Запрос к API с пагинацией
function fetchStocksWithPagination(token) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses';
  var allItems = [];
  var offset = 0;
  var limit = 100000;
  var hasMore = true;
  
  while (hasMore) {
    var payload = { "limit": limit, "offset": offset };
    var options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 429) { Utilities.sleep(30000); continue; }
      if (response.getResponseCode() !== 200) break;
      
      var data = JSON.parse(response.getContentText());
      var items = data.data?.items || [];
      if (items.length === 0) break;
      
      allItems = allItems.concat(items);
      hasMore = items.length === limit;
      offset += limit;
      Utilities.sleep(200);
    } catch(e) { break; }
  }
  
  return allItems.map(function(item) {
    return { nmId: item.nmId || '', warehouse: item.warehouseName || '', amount: item.quantity || 0 };
  });
}

/** Шапка закрепляется только если ниже есть хотя бы одна строка. Иначе Таблицы бросают ошибку. */
function stocksSafeFreezeRows_(sheet, rows) {
  if (!sheet) return;
  var frozenCols = sheet.getFrozenColumns();
  if (frozenCols > 0 && frozenCols >= sheet.getMaxColumns()) sheet.setFrozenColumns(0);
  var want = rows > 0 ? rows : 0;
  if (want > 0 && sheet.getMaxRows() > want) sheet.setFrozenRows(want);
  else if (sheet.getFrozenRows() > 0) sheet.setFrozenRows(0);
}

/** Удаляет строки. Если это все незакреплённые строки листа, шапку на время открепляет. */
function stocksDeleteRows_(sheet, start, count) {
  if (!sheet || count < 1 || start < 1) return;
  var frozen = sheet.getFrozenRows();
  var deleteAllUnfrozen = frozen > 0 && start <= frozen + 1 && (start + count - 1) >= sheet.getMaxRows();
  if (deleteAllUnfrozen) sheet.setFrozenRows(0);
  sheet.deleteRows(start, count);
  if (deleteAllUnfrozen) stocksSafeFreezeRows_(sheet, frozen);
}

/** Снимает с листа «Остатки» все строки выбранной даты (обычно сегодняшний снимок). */
function stocksRemoveDate_(sheet, headers, dateStr) {
  if (!sheet || !dateStr || sheet.getLastRow() < 2) return 0;
  var dateCol = 0;
  if (headers && headers.length) {
    var di = headers.indexOf('Дата');
    if (di >= 0) dateCol = di;
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var oldDates = sheet.getRange(2, dateCol + 1, lastRow - 1, 1).getValues();
  var firstDrop = -1;
  var mixed = false;
  var i;
  for (i = 0; i < oldDates.length; i++) {
    var key = sheetDayKey_(oldDates[i][0]);
    var same = key === dateStr;
    if (same && firstDrop < 0) firstDrop = i;
    if (!same && firstDrop >= 0) mixed = true;
  }
  if (firstDrop < 0) return 0;
  if (!mixed) {
    stocksDeleteRows_(sheet, 2 + firstDrop, oldDates.length - firstDrop);
    return oldDates.length - firstDrop;
  }
  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var kept = [];
  for (i = 0; i < all.length; i++) {
    var k = sheetDayKey_(all[i][dateCol]);
    if (k !== dateStr) kept.push(all[i]);
  }
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, lastCol).setValues(kept);
  return all.length - kept.length;
}

// Сохранение в лист "Остатки" (добавляем Регион и Веса)
function appendToStockSheetFull(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Остатки');
  
  // Новые заголовки с весами
  var newHeaders = ['Дата','Артикул продавца','Название','Артикул WB','Предмет','Бренд','Размер','Склад','Остаток','Регион поставки','Вес, кг за 1 шт','Общий вес, кг'];
  
  // Если листа нет — создаём с полными заголовками
  if (!sheet) {
    sheet = ss.insertSheet('Остатки');
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Проверяем и добавляем недостающие колонки
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Добавляем колонку "Регион поставки" если нет
    if (existingHeaders.indexOf('Регион поставки') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Регион поставки');
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // Добавляем колонку "Вес, кг за 1 шт" если нет
    if (existingHeaders.indexOf('Вес, кг за 1 шт') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Вес, кг за 1 шт');
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // Добавляем колонку "Общий вес, кг" если нет
    if (existingHeaders.indexOf('Общий вес, кг') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Общий вес, кг');
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  }
  
  // Получаем актуальные заголовки после всех добавлений
  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);

  var dateStr = rows.length ? (typeof adsDateKey_ === 'function' ? adsDateKey_(rows[0][0]) : String(rows[0][0]).substring(0, 10)) : '';
  stocksRemoveDate_(sheet, currentHeaders, dateStr);

  var newRows = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i].slice();
    while (r.length < 12) r.push('');
    if (r.length > 12) r = r.slice(0, 12);
    newRows.push(r);
  }

  if (newRows.length === 0) {
    wbPutSheetFilter_(sheet);
    return 0;
  }
  
  // Записываем
  var startRow = sheet.getLastRow() + 1;
  for (var i = 0; i < newRows.length; i += 1000) {
    var batch = newRows.slice(i, i + 1000);
    sheet.getRange(startRow + i, 1, batch.length, 12).setValues(batch);
  }
  
  // Форматируем колонки с весами
  var weightColumnIndex = currentHeaders.indexOf('Вес, кг за 1 шт') + 1;
  var totalWeightColumnIndex = currentHeaders.indexOf('Общий вес, кг') + 1;
  if (weightColumnIndex > 0 && totalWeightColumnIndex > 0) {
    var lastRowAfterAdd = sheet.getLastRow();
    if (lastRowAfterAdd >= 2) {
      sheet.getRange(2, weightColumnIndex, lastRowAfterAdd - 1, 1).setNumberFormat('0.000');
      sheet.getRange(2, totalWeightColumnIndex, lastRowAfterAdd - 1, 1).setNumberFormat('0.00');
    }
  }

  wbPutSheetFilter_(sheet);
  return newRows.length;
}

// Справочник товаров
function getProductsDictionary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Товары');
  var dict = {};
  if (!sheet) return dict;
  
  // Расширяем диапазон для получения большего количества колонок
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(6, lastCol)).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      dict[row[0]] = { 
        vendorCode: row[2] || '', 
        title: row[1] || '', 
        brand: row[3] || '', 
        subject: row[4] || '' 
      };
    }
  }
  return dict;
}

// Загрузка весов для остатков (аналогично поставкам)
function loadWeightsMapForStocks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Товар+Вес');
  
  var weightsMap = {};
  
  if (!sheet) {
    Logger.log('⚠️ Лист "Товар+Вес" не найден');
    return weightsMap;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('⚠️ Лист "Товар+Вес" пуст');
    return weightsMap;
  }
  
  var headers = data[0];
  var artikulIndex = -1;
  var weightIndex = -1;
  
  // Ищем колонки (первая колонка - Артикул, вторая - Вес)
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i].toString().toLowerCase();
    if (header.indexOf('артикул') !== -1) {
      artikulIndex = i;
    }
    if (header.indexOf('вес') !== -1 && header.indexOf('шт') !== -1) {
      weightIndex = i;
    }
  }
  
  // Если не нашли по названиям, берем первые две колонки
  if (artikulIndex === -1) artikulIndex = 0;
  if (weightIndex === -1) weightIndex = 1;
  
  for (var i = 1; i < data.length; i++) {
    var artikul = data[i][artikulIndex];
    var weight = data[i][weightIndex];
    if (artikul && weight !== '' && weight !== null && weight !== undefined) {
      // Преобразуем артикул в строку для сравнения
      weightsMap[artikul.toString().trim()] = parseFloat(weight) || 0;
    }
  }
  
  return weightsMap;
}


// Ручной запуск
function updateAllStocksData() { recordCurrentStocksEnriched(false); }
function recordStocksFast() { recordCurrentStocksEnriched(true); }

function diagnoseWbWarehouses() {
  var token = getApiToken();
  if (!token) { showAlert('❌ Токен не найден'); return; }
  
  var response = UrlFetchApp.fetch('https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses', {
    method: 'POST', headers: { 'Authorization': token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ "limit": 10, "offset": 0 }), muteHttpExceptions: true
  });
  
  var status = response.getResponseCode();
  Logger.log('Статус: ' + status);
  if (status === 200) {
    var data = JSON.parse(response.getContentText());
    showAlert('✅ Статус 200\nНайдено позиций: ' + (data.data?.items?.length || 0));
  } else {
    showAlert('❌ Ошибка: ' + status);
  }
}

function getDateStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }



// ============================================
// СКЛАДЫ WB И ТРАНЗИТНЫЕ ТАРИФЫ (С РЕГИОНАМИ)
// ============================================

// 1. ЗАГРУЗКА СПИСКА СКЛАДОВ WB (с регионом)
/* отключено
function loadWarehousesList() {
  var token = getApiToken();
  if (!token) return;
  
  var url = 'https://supplies-api.wildberries.ru/api/v1/warehouses';
  var options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    showAlert('❌ Ошибка загрузки складов: ' + response.getResponseCode());
    return;
  }
  
  var warehouses = JSON.parse(response.getContentText());
  saveWarehousesToSheet(warehouses);
}

function saveWarehousesToSheet(warehouses) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Склады_WB';
  var sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  
  var headers = ['ID склада', 'Название', 'Адрес', 'Режим работы', 'Активен', 'Регион поставки'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // Загружаем справочник регионов
  var regionMap = getRegionMap();
  
  var rows = [];
  for (var i = 0; i < warehouses.length; i++) {
    var w = warehouses[i];
    var warehouseName = w.name || '';
    var region = getRegionFromMap(warehouseName, regionMap);
    
    rows.push([
      w.ID || '',
      warehouseName,
      w.address || '',
      w.workTime || '',
      w.isActive ? 'Да' : 'Нет',
      region
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }

  wbPutSheetFilter_(sheet);
  showAlert('✅ Загружено складов: ' + rows.length);
}

// 2. ЗАГРУЗКА ТРАНЗИТНЫХ ТАРИФОВ (с регионом)
function loadTransitTariffs() {
  var token = getApiToken();
  if (!token) return;
  
  var url = 'https://supplies-api.wildberries.ru/api/v1/transit-tariffs';
  var options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    showAlert('❌ Ошибка загрузки транзитных тарифов: ' + response.getResponseCode());
    return;
  }
  
  var tariffs = JSON.parse(response.getContentText());
  saveTransitTariffsToSheet(tariffs);
}

function saveTransitTariffsToSheet(tariffs) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Транзитные_тарифы';
  var sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  
  var headers = ['Откуда', 'Куда', 'Активен с', 'Тариф паллета', 'Регион (Откуда)', 'Регион (Куда)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // Загружаем справочник регионов
  var regionMap = getRegionMap();
  
  var rows = [];
  for (var i = 0; i < tariffs.length; i++) {
    var t = tariffs[i];
    var fromWarehouse = t.transitWarehouseName || '';
    var toWarehouse = t.destinationWarehouseName || '';
    
    rows.push([
      fromWarehouse,
      toWarehouse,
      t.activeFrom || '',
      t.palletTariff || '',
      getRegionFromMap(fromWarehouse, regionMap),
      getRegionFromMap(toWarehouse, regionMap)
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }
  
  showAlert('✅ Загружено транзитных направлений: ' + rows.length);
}
*/

// 3. ПОЛУЧЕНИЕ СПРАВОЧНИКА РЕГИОНОВ ИЗ ЛИСТА "Регионы"
function getRegionMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Регион поставки');
  var regionMap = {};
  
  if (!sheet) {
    Logger.log('Лист "Регионы" не найден. Создайте его с колонками: Город, Регион');
    return regionMap;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return regionMap;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var city = data[i][0];
    var region = data[i][1];
    if (city && region) {
      regionMap[city.toString().toLowerCase()] = region;
    }
  }
  
  return regionMap;
}

// 4. ПОЛУЧЕНИЕ РЕГИОНА ПО НАЗВАНИЮ СКЛАДА
function getRegionFromMap(warehouseName, regionMap) {
  if (!warehouseName || warehouseName === '') return 'Не указан';
  
  var name = warehouseName.toLowerCase();
  
  // Точное совпадение
  if (regionMap[name]) {
    return regionMap[name];
  }
  
  // Частичное совпадение (ищем вхождение)
  for (var key in regionMap) {
    if (name.indexOf(key) !== -1) {
      return regionMap[key];
    }
  }
  
  return 'Не определён';
}

// 5. ПОЛУЧЕНИЕ СПРАВОЧНИКА СКЛАДОВ (ДЛЯ ДРУГИХ ФУНКЦИЙ)
function getWarehousesDictionary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Склады_WB');
  var dict = {};
  
  if (!sheet) return dict;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return dict;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0];
    var name = data[i][1];
    if (id && name) {
      dict[id.toString()] = name;
    }
  }
  
  return dict;
}



// ============================================
// ПОСТАВКИ FBW - ПОЛНАЯ РАБОЧАЯ ВЕРСИЯ (БЕЗ ДУБЛЕЙ)
// ============================================

function formatDateISO(dateString) {
  if (!dateString) return '';
  if (dateString instanceof Date) {
    var day = dateString.getDate().toString().padStart(2, '0');
    var month = (dateString.getMonth() + 1).toString().padStart(2, '0');
    var year = dateString.getFullYear();
    return day + '.' + month + '.' + year;
  }
  var str = dateString.toString();
  if (str.indexOf('T') !== -1 && str.indexOf('-') !== -1) {
    var datePart = str.split('T')[0];
    var parts = datePart.split('-');
    if (parts.length === 3) {
      return parts[2] + '.' + parts[1] + '.' + parts[0];
    }
  }
  return str;
}

// 1. ДИАЛОГ
function showSuppliesDialog() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:450px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:8px;margin-bottom:12px;border:1px solid #ccc;border-radius:4px}' +
    'button{background:#4CAF50;color:white;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:5px}' +
    'button:hover{background:#45a049}' +
    'button:disabled{background:#9e9e9e !important;cursor:default}' +
    '.info{background:#e3f2fd;padding:10px;margin:10px 0;border-radius:4px;font-size:12px}' +
    '</style></head><body>' +
    '<h3>📦 Поставки FBW</h3>' +
    '<div class="info">📌 Выберите период до 90 дней назад</div>' +
    '<label>📅 Дата ОТ:</label><input type="date" id="dateFrom" value="' + getDefaultDateFrom() + '">' +
    '<label>📅 Дата ДО:</label><input type="date" id="dateTo" value="' + getDefaultDateTo() + '">' +
    '<button onclick="loadList()">📥 1. Загрузить список поставок</button>' +
    '<button onclick="loadAllDetailsAndGoods()" style="background:#FF9800">📦 2. Загрузить детали и товары (для всех поставок)</button>' +
    '<button onclick="createUnified()" style="background:#9C27B0">🔗 3. Создать объединённый лист "Поставки"</button>' +
    '<hr>' +
    '<label>🔍 Загрузить по ID:</label>' +
    '<input type="text" id="supplyId" placeholder="Например: 37805683">' +
    '<button onclick="loadDetailsById()">📋 Детали поставки</button>' +
    '<button onclick="loadGoodsById()">📦 Товары поставки</button>' +
    '<script>' +
    'function lockBtns(on){var all=document.querySelectorAll("button");for(var i=0;i<all.length;i++) all[i].disabled=!!on;}' +
    'function loadList(){ var d1=document.getElementById("dateFrom").value; var d2=document.getElementById("dateTo").value; if(!d1||!d2){alert("Выберите даты!");return;} lockBtns(true); google.script.run.withSuccessHandler(function(){lockBtns(false);alert("Список поставок загружен! Теперь нажмите Шаг 2.");}).withFailureHandler(function(e){lockBtns(false);alert(e);}).loadSuppliesListWithDates(d1,d2); }' +
    'function loadAllDetailsAndGoods(){ lockBtns(true); google.script.run.withSuccessHandler(function(){lockBtns(false);alert("Детали и товары загружены! Теперь нажмите Шаг 3.");}).withFailureHandler(function(e){lockBtns(false);alert(e);}).loadAllDetailsAndGoodsFromList(); }' +
    'function createUnified(){ lockBtns(true); google.script.run.withSuccessHandler(function(){lockBtns(false);alert("Объединённый лист Поставки создан!");}).withFailureHandler(function(e){lockBtns(false);alert(e);}).createUnifiedSuppliesSheet(); }' +
    'function loadDetailsById(){ var id=document.getElementById("supplyId").value; if(!id){alert("Введите ID!");return;} lockBtns(true); google.script.run.withSuccessHandler(function(){lockBtns(false);alert("Детали загружены");}).withFailureHandler(function(e){lockBtns(false);alert(e);}).loadSupplyDetailsById(id); }' +
    'function loadGoodsById(){ var id=document.getElementById("supplyId").value; if(!id){alert("Введите ID!");return;} lockBtns(true); google.script.run.withSuccessHandler(function(){lockBtns(false);alert("Товары загружены");}).withFailureHandler(function(e){lockBtns(false);alert(e);}).loadSupplyGoodsById(id); }' +
    '</script></body></html>';
  
  var dialog = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Поставки FBW');
}

// 2. ЗАГРУЗКА СПИСКА ПОСТАВОК ЗА ПЕРИОД
function loadSuppliesListWithDates(dateFrom, dateTo) {
  var token = getApiToken();
  if (!token) {
    showAlert('❌ Токен не найден');
    return;
  }
  
  showAlert('⏳ Загружаем список поставок за ' + dateFrom + ' — ' + dateTo);
  
  var allSupplies = [];
  var offset = 0;
  var limit = 1000;
  var hasMore = true;
  
  while (hasMore) {
    var url = 'https://supplies-api.wildberries.ru/api/v1/supplies?limit=' + limit + '&offset=' + offset;
    var payload = {
      "Type": "FBW",
      "dateFrom": dateFrom,
      "dateTo": dateTo
    };
    
    var options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) {
        showAlert('❌ Ошибка API: ' + response.getResponseCode());
        return;
      }
      
      var data = JSON.parse(response.getContentText());
      var supplies = Array.isArray(data) ? data : (data.data || data.supplies || []);
      
      if (supplies.length === 0) {
        hasMore = false;
      } else {
        allSupplies = allSupplies.concat(supplies);
        hasMore = supplies.length === limit;
        offset += limit;
        Utilities.sleep(200);
      }
    } catch(e) {
      Logger.log('Ошибка: ' + e);
      showAlert('❌ Ошибка: ' + e);
      return;
    }
  }
  
  if (allSupplies.length === 0) {
    showAlert('⚠️ Нет поставок за период ' + dateFrom + ' — ' + dateTo);
    return;
  }
  
  saveSuppliesListToSheet(allSupplies);
  showAlert('✅ Загружено поставок: ' + allSupplies.length);
}

function saveSuppliesListToSheet(supplies) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof renameWbColumnHeaders_ === 'function') renameWbColumnHeaders_(ss);
  var sheetName = 'Поставки_список';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clear(); else sheet = ss.insertSheet(sheetName);
  
  if (supplies.length === 0) {
    sheet.getRange(1, 1).setValue('Нет данных');
    wbPutSheetFilter_(sheet);
    return;
  }
  
  if (typeof wbRowToRu_ === 'function') {
    supplies = supplies.map(function(s) { return wbRowToRu_(s, typeof wbMapForSheet_ === 'function' ? wbMapForSheet_(sheetName) : null); });
  }
  var headers = Object.keys(supplies[0]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 🔥 Функция парсинга даты из строки API в объект Date
  function parseApiDate(dateStr) {
    if (!dateStr) return null;
    if (dateStr instanceof Date) return dateStr;
    
    var str = dateStr.toString();
    // Формат: "2025-03-18 7:07:39" или "2025-03-18T07:07:39"
    var cleanStr = str.replace(' ', 'T');
    var d = new Date(cleanStr);
    
    return isNaN(d.getTime()) ? null : d;
  }
  
  var rows = [];
  for (var i = 0; i < supplies.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = supplies[i][headers[j]];
      var headerName = headers[j];
      
      // 🔥 Превращаем строки дат в объекты Date
      // if ((headerName === 'createDate' || headerName === 'supplyDate' || 
      //      headerName === 'factDate' || headerName === 'updatedDate') && val) {
      //   row.push(parseApiDate(val));
      // Стало:
      if ((typeof wbIsApiDateField_ === 'function' ? wbIsApiDateField_(headerName) :
          (headerName === 'createDate' || headerName === 'supplyDate' ||
          headerName === 'factDate' || headerName === 'updatedDate' ||
          headerName.indexOf('Date') !== -1)) && val) {
        var dateObj = parseApiDate(val);  // ✅ Превращает в объект Date
        row.push(dateObj || val);
      } else if (val === null || val === undefined) {
        row.push('');
      } else if (typeof val === 'object') {
        row.push(JSON.stringify(val));
      } else {
        row.push(val);
      }
    }
    rows.push(row);
  }
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // 🔥 Форматируем столбцы с датами для отображения
  var dateCols = ['createDate', 'supplyDate', 'factDate', 'updatedDate'];
  for (var k = 0; k < dateCols.length; k++) {
    var colIndex = (typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [dateCols[k]]) : headers.indexOf(dateCols[k])) + 1;
    if (colIndex > 0) {
      sheet.getRange(2, colIndex, rows.length, 1).setNumberFormat('dd.mm.yyyy');
    }
  }
  
  sheet.autoResizeColumns(1, headers.length);
  wbPutSheetFilter_(sheet);
}

// 3. ЗАГРУЗКА ДЕТАЛЕЙ И ТОВАРОВ ДЛЯ ВСЕХ ПОСТАВОК
function loadAllDetailsAndGoodsFromList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Поставки_список');
  
  if (!sheet) {
    showAlert('❌ Лист "Поставки_список" не найден. Сначала загрузите список (Шаг 1).');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    showAlert('⚠️ Нет данных в листе "Поставки_список"');
    return;
  }
  
  var headers = data[0];
  var supplies = [];
  for (var i = 1; i < data.length; i++) {
    var supplyId = getValueByHeader(data[i], headers, 'supplyID');
    if (supplyId) {
      supplies.push({ supplyID: supplyId });
    }
  }
  
  if (supplies.length === 0) {
    showAlert('⚠️ Не найдено ID поставок');
    return;
  }
  
  showAlert('⏳ Загружаем детали и товары для ' + supplies.length + ' поставок...\n\nЭто займет около ' + Math.ceil(supplies.length * 2) + ' секунд.');
  
  var token = getApiToken();
  if (!token) return;
  
  var allDetails = [];
  var allGoods = [];
  
  for (var i = 0; i < supplies.length; i++) {
    var supplyId = supplies[i].supplyID;
    if (!supplyId) continue;
    
    Logger.log('Загружаем поставку ' + supplyId + ' (' + (i+1) + '/' + supplies.length + ')');
    
    if (i > 0) Utilities.sleep(1500);
    
    // ДЕТАЛИ
    var detailsUrl = 'https://supplies-api.wildberries.ru/api/v1/supplies/' + supplyId + '?isPreorderID=false';
    var detailsResponse = UrlFetchApp.fetch(detailsUrl, { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true });
    
    if (detailsResponse.getResponseCode() === 200) {
      var details = JSON.parse(detailsResponse.getContentText());
      details.supplyID = supplyId;
      allDetails.push(details);
      Logger.log('   ✅ Детали загружены, склад: ' + (details.warehouseName || details.actualWarehouseName || 'не указан'));
    } else if (detailsResponse.getResponseCode() === 429) {
      Logger.log('   ⏳ Ошибка 429, ждём 5 секунд...');
      Utilities.sleep(5000);
      i--;
      continue;
    } else {
      Logger.log('   ❌ Ошибка деталей: ' + detailsResponse.getResponseCode());
    }
    
    Utilities.sleep(800);
    
    // ТОВАРЫ
    var goodsUrl = 'https://supplies-api.wildberries.ru/api/v1/supplies/' + supplyId + '/goods?limit=100&offset=0&isPreorderID=false';
    var goodsResponse = UrlFetchApp.fetch(goodsUrl, { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true });
    
    if (goodsResponse.getResponseCode() === 200) {
      var goodsData = JSON.parse(goodsResponse.getContentText());
      var goodsArray = Array.isArray(goodsData) ? goodsData : (goodsData.data || goodsData.goods || []);
      for (var j = 0; j < goodsArray.length; j++) {
        goodsArray[j].supplyID = supplyId;
        allGoods.push(goodsArray[j]);
      }
      Logger.log('   📦 Товаров: ' + goodsArray.length);
    } else if (goodsResponse.getResponseCode() === 429) {
      Logger.log('   ⏳ Ошибка 429, ждём 5 секунд...');
      Utilities.sleep(5000);
      i--;
      continue;
    } else {
      Logger.log('   ❌ Ошибка товаров: ' + goodsResponse.getResponseCode());
    }
    
    Utilities.sleep(800);
    if ((i+1) % 10 === 0) Logger.log('📊 Прогресс: ' + (i+1) + '/' + supplies.length);
  }
  
  saveDetailsToSheet(allDetails);
  saveGoodsToSheet(allGoods);
  
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('✅ Готово!\n\nДеталей: ' + allDetails.length + '\nТоваров: ' + allGoods.length);
  } catch(e) {
    Logger.log('✅ Готово!\nДеталей: ' + allDetails.length + '\nТоваров: ' + allGoods.length);
  }
}

function saveDetailsToSheet(details) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Детали_поставок');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Детали_поставок');
  
  if (details.length === 0) { sheet.getRange(1,1).setValue('Нет данных'); wbPutSheetFilter_(sheet); return; }
  
  if (typeof wbRowToRu_ === 'function') {
    details = details.map(function(d) { return wbRowToRu_(d, typeof wbMapForSheet_ === 'function' ? wbMapForSheet_('Детали_поставок') : null); });
  }
  var headers = Object.keys(details[0]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 🔥 Функция парсинга даты
  function parseApiDate(dateStr) {
    if (!dateStr) return null;
    if (dateStr instanceof Date) return dateStr;
    var str = dateStr.toString();
    var cleanStr = str.replace(' ', 'T');
    var d = new Date(cleanStr);
    return isNaN(d.getTime()) ? null : d;
  }
  
  var rows = [];
  for (var i = 0; i < details.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = details[i][headers[j]];
      var headerName = headers[j];
      
      // 🔥 Превращаем строки дат в объекты Date
      if ((typeof wbIsApiDateField_ === 'function' ? wbIsApiDateField_(headerName) :
          (headerName.indexOf('Date') !== -1 ||
           headerName === 'createDate' || headerName === 'supplyDate' ||
           headerName === 'factDate' || headerName === 'updatedDate')) && val) {
        row.push(parseApiDate(val));
      } else if (val === null || val === undefined) {
        row.push('');
      } else if (typeof val === 'object') {
        row.push(JSON.stringify(val));
      } else {
        row.push(val);
      }
    }
    rows.push(row);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // 🔥 Форматируем все столбцы с датами
  for (var j = 0; j < headers.length; j++) {
    if (typeof wbIsApiDateField_ === 'function' ? wbIsApiDateField_(headers[j]) :
        (headers[j].indexOf('Date') !== -1 ||
        headers[j] === 'createDate' || headers[j] === 'supplyDate' ||
        headers[j] === 'factDate' || headers[j] === 'updatedDate')) {
      sheet.getRange(2, j + 1, rows.length, 1).setNumberFormat('dd.mm.yyyy');
    }
  }
  
  sheet.autoResizeColumns(1, headers.length);
  wbPutSheetFilter_(sheet);
}

// Сохранение товаров
function saveGoodsToSheet(goods) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Товары_поставок');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Товары_поставок');
  
  if (goods.length === 0) { sheet.getRange(1,1).setValue('Нет данных'); wbPutSheetFilter_(sheet); return; }
  
  if (typeof wbRowToRu_ === 'function') {
    goods = goods.map(function(g) { return wbRowToRu_(g, typeof wbMapForSheet_ === 'function' ? wbMapForSheet_('Товары_поставок') : null); });
  }
  var headers = Object.keys(goods[0]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  var rows = [];
  for (var i = 0; i < goods.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = goods[i][headers[j]];
      row.push((val === null || val === undefined) ? '' : (typeof val === 'object' ? JSON.stringify(val) : val));
    }
    rows.push(row);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  wbPutSheetFilter_(sheet);
}

// 4. ЗАГРУЗКА ДЕТАЛЕЙ ПО ID
function loadSupplyDetailsById(supplyId) {
  var token = getApiToken();
  if (!token) return;
  
  var url = 'https://supplies-api.wildberries.ru/api/v1/supplies/' + supplyId + '?isPreorderID=false';
  var options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) { showAlert('❌ Ошибка: ' + response.getResponseCode()); return; }
    
    var details = JSON.parse(response.getContentText());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Детали_поставки_' + supplyId);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet('Детали_поставки_' + supplyId);
    
    var headers = ['Поле', 'Значение'];
    sheet.getRange(1,1,1,2).setValues([headers]);
    sheet.getRange(1,1,1,2).setFontWeight('bold');
    
    var rows = [];
    for (var key in details) {
      var val = details[key];
      rows.push([key, (val === null || val === undefined) ? '' : (typeof val === 'object' ? JSON.stringify(val) : val)]);
    }
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    sheet.autoResizeColumns(1, 2);
    showAlert('✅ Детали поставки ' + supplyId + ' сохранены');
  } catch(e) { showAlert('❌ Ошибка: ' + e); }
}

// 5. ЗАГРУЗКА ТОВАРОВ ПО ID
function loadSupplyGoodsById(supplyId) {
  var token = getApiToken();
  if (!token) return;
  
  var url = 'https://supplies-api.wildberries.ru/api/v1/supplies/' + supplyId + '/goods?limit=100&offset=0&isPreorderID=false';
  var options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) { showAlert('❌ Ошибка: ' + response.getResponseCode()); return; }
    
    var goodsData = JSON.parse(response.getContentText());
    var goods = Array.isArray(goodsData) ? goodsData : (goodsData.data || goodsData.goods || []);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Товары_поставки_' + supplyId);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet('Товары_поставки_' + supplyId);
    
    if (goods.length === 0) { sheet.getRange(1,1).setValue('Нет товаров'); return; }
    
    var headers = Object.keys(goods[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    var rows = [];
    for (var i = 0; i < goods.length; i++) {
      var row = [];
      for (var j = 0; j < headers.length; j++) {
        var val = goods[i][headers[j]];
        row.push((val === null || val === undefined) ? '' : (typeof val === 'object' ? JSON.stringify(val) : val));
      }
      rows.push(row);
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
    showAlert('✅ Товаров в поставке ' + supplyId + ': ' + goods.length);
  } catch(e) { showAlert('❌ Ошибка: ' + e); }
}

// 6. ОБЪЕДИНЁННЫЙ ЛИСТ "ПОСТАВКИ" С РЕГИОНОМ И ВЕСАМИ (даты как числа)
function createUnifiedSuppliesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suppliesSheet = ss.getSheetByName('Поставки_список');
  if (!suppliesSheet) { showAlert('❌ Лист "Поставки_список" не найден'); return; }
  
  var suppliesData = suppliesSheet.getDataRange().getValues();
  var suppliesHeaders = suppliesData[0];
  var suppliesRows = suppliesData.slice(1);
  
  var detailsSheet = ss.getSheetByName('Детали_поставок');
  var detailsMap = {};
  var detailsHeaders = [];
  if (detailsSheet) {
    var detailsData = detailsSheet.getDataRange().getValues();
    detailsHeaders = detailsData[0];
    for (var i = 1; i < detailsData.length; i++) {
      var supplyId = getValueByHeader(detailsData[i], detailsHeaders, 'supplyID');
      if (supplyId) detailsMap[supplyId] = detailsData[i];
    }
  }
  
  var goodsSheet = ss.getSheetByName('Товары_поставок');
  var goodsMap = {};
  var goodsHeaders = [];
  if (goodsSheet) {
    var goodsData = goodsSheet.getDataRange().getValues();
    goodsHeaders = goodsData[0];
    for (var i = 1; i < goodsData.length; i++) {
      var supplyId = getValueByHeader(goodsData[i], goodsHeaders, 'supplyID');
      if (supplyId) {
        if (!goodsMap[supplyId]) goodsMap[supplyId] = [];
        goodsMap[supplyId].push(goodsData[i]);
      }
    }
  }
  
  var regionMap = getRegionMapFromSheet();
  var weightsMap = loadWeightsMap();
  
  var unifiedSheet = ss.getSheetByName('Поставки');
  if (unifiedSheet) unifiedSheet.clear();
  else unifiedSheet = ss.insertSheet('Поставки');
  
  var headers = [
    'ID поставки', 'ID заказа', 'Дата создания', 'Дата поставки', 'Факт дата',
    'Статус', 'Тип коробки', 'Склад', 'Регион поставки',
    'Баркод', 'Артикул продавца', 'Артикул WB', 'Размер', 'Цвет',
    'Количество', 'Принято', 'Готово к продаже',
    'Вес, кг за 1 шт', 'Общий вес, кг'
  ];
  unifiedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  unifiedSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  unifiedSheet.setFrozenRows(1);
  
  // Функция преобразования текстовой даты в объект Date
  function parseDateToNumber(dateString) {
    if (!dateString) return null;
    if (dateString instanceof Date) return dateString;
    if (typeof dateString === 'string') {
      var parts = dateString.split('.');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        return new Date(year, month, day);
      }
    }
    return null;
  }
  
  // Функция форматирования даты для отображения (только для заголовка)
  function formatDateForDisplay(dateObj) {
    if (!dateObj) return '';
    if (!(dateObj instanceof Date)) return dateObj;
    var day = dateObj.getDate().toString().padStart(2, '0');
    var month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    var year = dateObj.getFullYear();
    return day + '.' + month + '.' + year;
  }
  
  var allRows = [];
  for (var i = 0; i < suppliesRows.length; i++) {
    var supply = suppliesRows[i];
    var supplyId = getValueByHeader(supply, suppliesHeaders, 'supplyID');
    
    // Получаем сырые даты и преобразуем в объекты Date
    var createDateRaw = getValueByHeader(supply, suppliesHeaders, 'createDate');
    var supplyDateRaw = getValueByHeader(supply, suppliesHeaders, 'supplyDate');
    var factDateRaw = getValueByHeader(supply, suppliesHeaders, 'factDate');
    
    // Преобразуем в объекты Date (числовые значения)
    var createDateObj = parseDateToNumber(createDateRaw);
    var supplyDateObj = parseDateToNumber(supplyDateRaw);
    var factDateObj = parseDateToNumber(factDateRaw);
    
    // Для отображения в заголовке используем форматированную строку
    var createDateDisplay = formatDateForDisplay(createDateObj);
    var supplyDateDisplay = formatDateForDisplay(supplyDateObj);
    var factDateDisplay = formatDateForDisplay(factDateObj);
    
    var details = detailsMap[supplyId];
    var warehouseName = 'Не указан';
    if (details) {
      warehouseName = getValueByHeader(details, detailsHeaders, 'warehouseName') || 
                      getValueByHeader(details, detailsHeaders, 'actualWarehouseName') || 
                      'Не указан';
    }
    
    var region = getRegionFromMap(warehouseName, regionMap);
    var goods = goodsMap[supplyId] || [];
    
    var baseRow = [
      supplyId,
      getValueByHeader(supply, suppliesHeaders, 'preorderID'),
      createDateObj,    // Объект Date (число)
      supplyDateObj,    // Объект Date (число)
      factDateObj,      // Объект Date (число)
      getStatusName(getValueByHeader(supply, suppliesHeaders, 'statusID')),
      getBoxTypeName(getValueByHeader(supply, suppliesHeaders, 'boxTypeID')),
      warehouseName,
      region
    ];
    
    if (goods.length > 0) {
      for (var j = 0; j < goods.length; j++) {
        var g = goods[j];
        
        var vendorCode = getValueByHeader(g, goodsHeaders, 'vendorCode');
        if (vendorCode && typeof vendorCode === 'string') {
          vendorCode = vendorCode.trim();
        }
        
        var weightPerUnit = 0;
        if (vendorCode && weightsMap[vendorCode.toString()]) {
          weightPerUnit = weightsMap[vendorCode.toString()];
        }
        
        var quantity = getValueByHeader(g, goodsHeaders, 'quantity') || 0;
        var totalWeight = weightPerUnit * quantity;
        
        allRows.push(baseRow.concat([
          getValueByHeader(g, goodsHeaders, 'barcode'),
          vendorCode,
          getValueByHeader(g, goodsHeaders, 'nmID'),
          getValueByHeader(g, goodsHeaders, 'techSize'),
          getValueByHeader(g, goodsHeaders, 'color'),
          quantity,
          getValueByHeader(g, goodsHeaders, 'acceptedQuantity') || 0,
          getValueByHeader(g, goodsHeaders, 'readyForSaleQuantity') || 0,
          weightPerUnit,
          totalWeight
        ]));
      }
    } else {
      allRows.push(baseRow.concat(['', '', '', '', '', 0, 0, 0, 0, 0]));
    }
  }
  
  if (allRows.length > 0) {
    unifiedSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
    unifiedSheet.autoResizeColumns(1, headers.length);
    
    // Форматируем столбцы с весами
    var weightColumnIndex = headers.indexOf('Вес, кг за 1 шт') + 1;
    var totalWeightColumnIndex = headers.indexOf('Общий вес, кг') + 1;
    if (weightColumnIndex > 0 && totalWeightColumnIndex > 0) {
      unifiedSheet.getRange(2, weightColumnIndex, allRows.length, 1).setNumberFormat('0.000');
      unifiedSheet.getRange(2, totalWeightColumnIndex, allRows.length, 1).setNumberFormat('0.00');
    }
    
    // Форматируем столбцы с датами (числовой формат для отображения)
    var createDateColIndex = headers.indexOf('Дата создания') + 1;
    var supplyDateColIndex = headers.indexOf('Дата поставки') + 1;
    var factDateColIndex = headers.indexOf('Факт дата') + 1;
    
    if (createDateColIndex > 0) {
      unifiedSheet.getRange(2, createDateColIndex, allRows.length, 1).setNumberFormat('dd.mm.yyyy');
    }
    if (supplyDateColIndex > 0) {
      unifiedSheet.getRange(2, supplyDateColIndex, allRows.length, 1).setNumberFormat('dd.mm.yyyy');
    }
    if (factDateColIndex > 0) {
      unifiedSheet.getRange(2, factDateColIndex, allRows.length, 1).setNumberFormat('dd.mm.yyyy');
    }
  }
  
  // Сортируем лист "Поставки" по дате поставки
  if (unifiedSheet && unifiedSheet.getLastRow() > 1) {
    var headerRow = unifiedSheet.getRange(1, 1, 1, unifiedSheet.getLastColumn()).getValues()[0];
    var dateColIndex = headerRow.indexOf('Дата поставки');
    if (dateColIndex !== -1) {
      var range = unifiedSheet.getRange(2, 1, unifiedSheet.getLastRow() - 1, unifiedSheet.getLastColumn());
      range.sort({column: dateColIndex + 1, ascending: true});
    }
  }
  
  wbPutSheetFilter_(unifiedSheet);
  showAlert('✅ Объединённый лист "Поставки" создан!\nСтрок: ' + allRows.length);
}

// ============================================
// ЕЖЕДНЕВНОЕ ОБНОВЛЕНИЕ ПОСТАВОК
// ============================================

// 🔥 Универсальный парсер дат (ISO, DD.MM.YYYY, JS Date)
function parseToJSDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  
  var str = val.toString().trim();
  
  // Формат DD.MM.YYYY или DD.MM.YYYY HH:MM:SS
  var mDD = str.match(/^(\d{2})\.(\d{2})\.(\d{4})(\s+(\d{1,2}):(\d{2}):(\d{2}))?$/);
  if (mDD) {
    var d = new Date(parseInt(mDD[3]), parseInt(mDD[2]) - 1, parseInt(mDD[1]),
                     mDD[5] ? parseInt(mDD[5]) : 0, mDD[6] ? parseInt(mDD[6]) : 0, mDD[7] ? parseInt(mDD[7]) : 0);
    return isNaN(d.getTime()) ? null : d;
  }
  
  // Формат ISO или стандартный
  var d = new Date(str.replace(' ', 'T'));
  return isNaN(d.getTime()) ? null : d;
}

function updateSuppliesDaily() {
  var token = getApiToken();
  if (!token) return;
  
  var dateFrom = getDate30DaysAgo();
  var dateTo = getDefaultDateTo();
  
  var freshSupplies = fetchSuppliesList(dateFrom, dateTo);
  if (!freshSupplies || freshSupplies.length === 0) {
    Logger.log('✅ Поставки: список пуст (' + dateFrom + ' — ' + dateTo + ')');
    wbPutNamedFilters_(['Поставки_список', 'Детали_поставок', 'Товары_поставок', 'Поставки']);
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existingIds = {};
  var activeIds = {};
  
  // Сбор существующих ID
  ['Поставки_список', 'Детали_поставок', 'Товары_поставок'].forEach(function(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idCol = typeof wbColIdx_ === 'function' ? wbColIdx_(hdrs, ['supplyID']) : hdrs.indexOf('supplyID');
    if (idCol < 0) return;
    var statusCol = sheetName === 'Поставки_список'
      ? (typeof wbColIdx_ === 'function' ? wbColIdx_(hdrs, ['statusID']) : hdrs.indexOf('statusID'))
      : -1;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][idCol]) continue;
      var existingId = data[i][idCol].toString();
      existingIds[existingId] = true;
      if (statusCol >= 0) {
        var status = String(data[i][statusCol] == null ? '' : data[i][statusCol]).trim();
        if (status && status !== '5' && status !== 'Принято') activeIds[existingId] = true;
      }
    }
  });
  
  var newSupplies = freshSupplies.filter(function(s) { return s.supplyID && !existingIds[s.supplyID.toString()]; });
  var refreshIds = Object.keys(activeIds);
  var refreshSeen = {};
  for (var ri = 0; ri < refreshIds.length; ri++) refreshSeen[refreshIds[ri]] = true;
  for (var ni = 0; ni < newSupplies.length; ni++) {
    var newId = String(newSupplies[ni].supplyID || '');
    if (newId && !refreshSeen[newId]) {
      refreshSeen[newId] = true;
      refreshIds.push(newId);
    }
  }
  if (!refreshIds.length) {
    Logger.log('✅ Поставки: новых и активных нет (' + dateFrom + ' — ' + dateTo + ')');
    return;
  }
  Logger.log('✅ Поставки: новых ' + newSupplies.length + ', обновляем активных ' + Object.keys(activeIds).length);
  
  var allDetails = [], allGoods = [], refreshedDetailIds = [], refreshedGoodsIds = [];
  for (var i = 0; i < refreshIds.length; i++) {
    var sid = refreshIds[i];
    if (!sid) continue;
    Utilities.sleep(200);
    
    var detRes = UrlFetchApp.fetch('https://supplies-api.wildberries.ru/api/v1/supplies/' + sid + '?isPreorderID=false', { method:'GET', headers:{'Authorization':token}, muteHttpExceptions:true });
    if (detRes.getResponseCode() === 200) {
      var d = JSON.parse(detRes.getContentText());
      d.supplyID = sid;
      allDetails.push(d);
      refreshedDetailIds.push(sid);
    }
    
    Utilities.sleep(200);
    var goodsRes = UrlFetchApp.fetch('https://supplies-api.wildberries.ru/api/v1/supplies/' + sid + '/goods?limit=100&offset=0&isPreorderID=false', { method:'GET', headers:{'Authorization':token}, muteHttpExceptions:true });
    if (goodsRes.getResponseCode() === 200) { 
      var gData = JSON.parse(goodsRes.getContentText()); 
      var gArr = Array.isArray(gData)
        ? gData
        : (Array.isArray(gData.goods)
          ? gData.goods
          : (Array.isArray(gData.data)
            ? gData.data
            : (gData.data && Array.isArray(gData.data.goods) ? gData.data.goods : [])));
      gArr.forEach(function(g){ g.supplyID = sid; allGoods.push(g); });
      refreshedGoodsIds.push(sid);
    } else {
      Logger.log('⚠️ Товары поставки ' + sid + ': HTTP ' + goodsRes.getResponseCode() + ' ' + goodsRes.getContentText().slice(0, 300));
    }
  }
  
  appendSuppliesToListSheetNoDuplicates(newSupplies);
  appendDetailsToSheetNoDuplicates(allDetails, refreshedDetailIds);
  appendGoodsToSheetNoDuplicates(allGoods, refreshedGoodsIds);
  if (typeof morningChainOn_ === 'function' && morningChainOn_()) {
    Logger.log('Утро: пересборка листа «Поставки» отложена до статусов FBW');
  } else {
    rebuildUnifiedSheet();
    sortAllSheetsBySupplyDate();
  }
  wbPutNamedFilters_(['Поставки_список', 'Детали_поставок', 'Товары_поставок', 'Поставки']);
}

function appendSuppliesToListSheetNoDuplicates(newSupplies) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Поставки_список');
  if (!sheet) { saveSuppliesListToSheet(newSupplies); return; }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var existingIds = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i=0; i<ids.length; i++) if (ids[i][0]) existingIds[ids[i][0].toString()] = true;
  }
  
  var unique = newSupplies.filter(function(s){ return s.supplyID && !existingIds[s.supplyID.toString()]; });
  if (unique.length === 0) return;
  
  var dateCols = ['createDate','supplyDate','factDate','updatedDate'];
  var rows = unique.map(function(s){
    return headers.map(function(h){
      var val = typeof wbCell_ === 'function' ? wbCell_(s, h) : s[h];
      if ((typeof wbIsApiDateField_ === 'function' ? wbIsApiDateField_(h) : dateCols.indexOf(h) !== -1) && val) return parseToJSDate(val);
      if (val === null || val === undefined) return '';
      return typeof val === 'object' && !(val instanceof Date) ? JSON.stringify(val) : val;
    });
  });
  
  sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
  dateCols.forEach(function(col){
    var idx = (typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [col]) : headers.indexOf(col)) + 1;
    if (idx > 0) sheet.getRange(lastRow + 1, idx, rows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
  });
}

function appendDetailsToSheetNoDuplicates(newDetails, replaceSupplyIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Детали_поставок');
  if (!sheet) { saveDetailsToSheet(newDetails); return; }
  if (newDetails.length === 0) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idIdx = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['supplyID']) : headers.indexOf('supplyID');
  if (idIdx === -1) { saveDetailsToSheet(newDetails); return; }
  
  var replaceIds = {};
  for (var r = 0; r < (replaceSupplyIds || []).length; r++) {
    replaceIds[String(replaceSupplyIds[r])] = true;
  }
  var existingIds = {};
  var existingRows = [];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var oldRows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < oldRows.length; i++) {
      var oldId = oldRows[i][idIdx] == null ? '' : oldRows[i][idIdx].toString();
      if (oldId && replaceIds[oldId]) continue;
      existingRows.push(oldRows[i]);
      if (oldId) existingIds[oldId] = true;
    }
  }
  
  var unique = newDetails.filter(function(d){ return d.supplyID && !existingIds[d.supplyID.toString()]; });
  if (unique.length === 0 && !replaceSupplyIds) return;
  
  var rows = unique.map(function(d){
    return headers.map(function(h){
      var val = typeof wbCell_ === 'function' ? wbCell_(d, h) : d[h];
      if ((typeof wbIsApiDateField_ === 'function' ? wbIsApiDateField_(h) : h.indexOf('Date') !== -1) && val) return parseToJSDate(val);
      if (val === null || val === undefined) return '';
      return typeof val === 'object' ? JSON.stringify(val) : val;
    });
  });
  
  if (replaceSupplyIds && replaceSupplyIds.length) {
    var finalRows = existingRows.concat(rows);
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    if (finalRows.length) sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  } else if (rows.length) {
    sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
  }
  for (var j=0; j<headers.length; j++) {
    if (headers[j].indexOf('Date') !== -1) {
      var formatFrom = replaceSupplyIds && replaceSupplyIds.length ? 2 : lastRow + 1;
      var formatRows = replaceSupplyIds && replaceSupplyIds.length ? existingRows.length + rows.length : rows.length;
      if (formatRows) sheet.getRange(formatFrom, j + 1, formatRows, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
    }
  }
}

// Вспомогательная функция для получения существующих деталей
function getExistingDetails(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}


function appendGoodsToSheetNoDuplicates(newGoods, replaceSupplyIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Товары_поставок');
  if (!sheet) {
    saveGoodsToSheet(newGoods);
    return;
  }
  
  if (newGoods.length === 0 && (!replaceSupplyIds || !replaceSupplyIds.length)) return;
  
  // Используем существующие заголовки листа
  var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headers = existingHeaders;
  
  var supplyIdIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['supplyID']) : headers.indexOf('supplyID');
  var nmIdIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['nmID']) : headers.indexOf('nmID');
  var barcodeIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['barcode']) : headers.indexOf('barcode');
  
  if (supplyIdIndex === -1) {
    saveGoodsToSheet(newGoods.concat(getExistingGoods(sheet)));
    return;
  }
  
  // Собираем существующие ключи
  var existingKeys = {};
  var replaceIds = {};
  for (var ri = 0; ri < (replaceSupplyIds || []).length; ri++) {
    replaceIds[String(replaceSupplyIds[ri])] = true;
  }
  var keptRows = [];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var existingData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var supplyId = existingData[i][supplyIdIndex];
      if (supplyId && replaceIds[String(supplyId)]) continue;
      keptRows.push(existingData[i]);
      var nmId = nmIdIndex !== -1 ? existingData[i][nmIdIndex] : '';
      var barcode = barcodeIndex !== -1 ? existingData[i][barcodeIndex] : '';
      if (supplyId) {
        var key = supplyId + '|' + (nmId || '') + '|' + (barcode || '');
        existingKeys[key] = true;
      }
    }
  }
  
  // Фильтруем новые
  var uniqueNew = [];
  for (var i = 0; i < newGoods.length; i++) {
    var supplyId = newGoods[i].supplyID;
    var nmId = newGoods[i].nmID;
    var barcode = newGoods[i].barcode;
    var key = supplyId + '|' + (nmId || '') + '|' + (barcode || '');
    if (!existingKeys[key]) {
      uniqueNew.push(newGoods[i]);
      existingKeys[key] = true;
    }
  }
  
  if (uniqueNew.length === 0 && (!replaceSupplyIds || !replaceSupplyIds.length)) return;
  
  // Приводим к структуре существующих заголовков
  var rows = [];
  for (var i = 0; i < uniqueNew.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = typeof wbCell_ === 'function' ? wbCell_(uniqueNew[i], headers[j]) : uniqueNew[i][headers[j]];
      row.push((val === null || val === undefined) ? '' : (typeof val === 'object' ? JSON.stringify(val) : val));
    }
    rows.push(row);
  }
  
  if (replaceSupplyIds && replaceSupplyIds.length) {
    var finalRows = keptRows.concat(rows);
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    if (finalRows.length) sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  } else if (rows.length) {
    sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function getExistingGoods(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function rebuildUnifiedSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var supSh = ss.getSheetByName('Поставки_список');
  if (!supSh) return;
  var sData = supSh.getDataRange().getValues();
  var sHdr = sData[0];
  var sRows = sData.slice(1);
  
  var detSh = ss.getSheetByName('Детали_поставок');
  var detMap = {}; var detHdr = [];
  if (detSh && detSh.getLastRow() > 1) {
    var dData = detSh.getDataRange().getValues(); detHdr = dData[0];
    for (var i=1; i<dData.length; i++) { var id = getValueByHeader(dData[i], detHdr, 'supplyID'); if (id) detMap[id] = dData[i]; }
  }
  
  var goodsSh = ss.getSheetByName('Товары_поставок');
  var goodsMap = {}; var goodsHdr = [];
  if (goodsSh && goodsSh.getLastRow() > 1) {
    var gData = goodsSh.getDataRange().getValues(); goodsHdr = gData[0];
    for (var i=1; i<gData.length; i++) { var id = getValueByHeader(gData[i], goodsHdr, 'supplyID'); if (id) { if (!goodsMap[id]) goodsMap[id] = []; goodsMap[id].push(gData[i]); } }
  }
  
  var regionMap = getRegionMapFromSheet();
  var weightsMap = loadWeightsMap();
  
  var uniSh = ss.getSheetByName('Поставки');
  if (uniSh) uniSh.clear(); else uniSh = ss.insertSheet('Поставки');
  
  var uniHdr = ['ID поставки','ID заказа','Дата создания','Дата поставки','Факт дата','Статус','Тип коробки','Склад','Регион поставки','Баркод','Артикул продавца','Артикул WB','Размер','Цвет','Количество','Принято','Готово к продаже','Вес, кг за 1 шт','Общий вес, кг'];
  uniSh.getRange(1,1,1,uniHdr.length).setValues([uniHdr]).setFontWeight('bold');
  uniSh.setFrozenRows(1);
  
  var allRows = [];
  for (var i=0; i<sRows.length; i++) {
    var row = sRows[i];
    var sid = getValueByHeader(row, sHdr, 'supplyID');
    var det = detMap[sid];
    var whName = det ? (getValueByHeader(det, detHdr, 'warehouseName') || getValueByHeader(det, detHdr, 'actualWarehouseName') || 'Не указан') : 'Не указан';
    var region = getRegionFromMap(whName, regionMap);
    var goods = goodsMap[sid] || [];
    
    // 🔥 Читаем даты как объекты Date
    var createDate = parseToJSDate(getValueByHeader(row, sHdr, 'createDate'));
    var supplyDate = parseToJSDate(getValueByHeader(row, sHdr, 'supplyDate'));
    var factDate = parseToJSDate(getValueByHeader(row, sHdr, 'factDate'));
    
    var base = [sid, getValueByHeader(row, sHdr, 'preorderID'), createDate, supplyDate, factDate,
                getStatusName(getValueByHeader(row, sHdr, 'statusID')), getBoxTypeName(getValueByHeader(row, sHdr, 'boxTypeID')),
                whName, region];
                
    if (goods.length > 0) {
      for (var j=0; j<goods.length; j++) {
        var g = goods[j];
        var art = getValueByHeader(g, goodsHdr, 'vendorCode');
        if (art && typeof art === 'string') art = art.trim();
        var w = weightsMap[art] || 0;
        var qty = getValueByHeader(g, goodsHdr, 'quantity') || 0;
        allRows.push(base.concat([getValueByHeader(g, goodsHdr, 'barcode'), art, getValueByHeader(g, goodsHdr, 'nmID'),
                                  getValueByHeader(g, goodsHdr, 'techSize'), getValueByHeader(g, goodsHdr, 'color'),
                                  qty, getValueByHeader(g, goodsHdr, 'acceptedQuantity')||0, getValueByHeader(g, goodsHdr, 'readyForSaleQuantity')||0, w, w*qty]));
      }
    } else {
      allRows.push(base.concat(['','','','','',0,0,0,0,0]));
    }
  }
  
  if (allRows.length > 0) {
    uniSh.getRange(2,1,allRows.length,uniHdr.length).setValues(allRows);
    ['Дата создания','Дата поставки','Факт дата'].forEach(function(c){
      var idx = uniHdr.indexOf(c) + 1;
      if (idx > 0) uniSh.getRange(2, idx, allRows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
    });
    ['Вес, кг за 1 шт','Общий вес, кг'].forEach(function(c){
      var idx = uniHdr.indexOf(c) + 1;
      if (idx > 0) uniSh.getRange(2, idx, allRows.length, 1).setNumberFormat('0.000');
    });
    uniSh.autoResizeColumns(1, uniHdr.length);
  }
  Logger.log('✅ Поставки перестроены: ' + allRows.length + ' строк');
}

function fetchSuppliesList(dateFrom, dateTo) {
  var token = getApiToken();
  if (!token) return [];

  var allSupplies = [];
  var offset = 0;
  var limit = 1000;
  var hasMore = true;

  while (hasMore) {
    var url = 'https://supplies-api.wildberries.ru/api/v1/supplies?limit=' + limit + '&offset=' + offset;
    var payload = { 'Type': 'FBW', 'dateFrom': dateFrom, 'dateTo': dateTo };
    var options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) break;
      var data = JSON.parse(response.getContentText());
      var supplies = Array.isArray(data) ? data : (data.data || data.supplies || []);
      if (!supplies.length) {
        hasMore = false;
      } else {
        allSupplies = allSupplies.concat(supplies);
        hasMore = supplies.length === limit;
        offset += limit;
        if (hasMore) Utilities.sleep(200);
      }
    } catch (e) {
      Logger.log('Ошибка fetchSuppliesList: ' + e);
      break;
    }
  }

  return allSupplies;
}

function getDate30DaysAgo() {
  var d = new Date();
  d.setDate(d.getDate() - 30);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

function getDefaultDateFrom() {
  var d = new Date();
  d.setDate(d.getDate() - 90);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDefaultDateTo() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getValueByHeader(row, headers, fieldName) {
  var index = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [fieldName]) : headers.indexOf(fieldName);
  if (index === -1) return '';
  var val = row[index];
  return (val === null || val === undefined) ? '' : val;
}

function getStatusName(id) {
  var s = {1:'Не запланировано',2:'Запланировано',3:'Отгрузка разрешена',4:'Идёт приёмка',5:'Принято',6:'Отгружено на воротах'};
  return s[id] || 'Неизвестно';
}

function getBoxTypeName(id) {
  var t = {1:'Монокороб',2:'Короб',3:'Дисплей',4:'Паллета',5:'Монопаллета'};
  return t[id] || 'Неизвестно';
}

function getRegionMapFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Регион поставки');
  var map = {};
  if (!sheet) return map;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][1]) map[data[i][0].toString().toLowerCase()] = data[i][1];
  }
  return map;
}


function showAlert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) { Logger.log(msg); }
}

// Загрузка справочника весов — см. loadWeightsMap() ниже (единственная версия)

// ============================================
// СОРТИРОВКА ЛИСТОВ ПО ДАТЕ ПОСТАВКИ
// ============================================

function sortAllSheetsBySupplyDate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Сортируем лист "Поставки_список"
  var suppliesListSheet = ss.getSheetByName('Поставки_список');
  if (suppliesListSheet) {
    sortSheetByDate(suppliesListSheet, 'supplyDate');
  }
  
  // Сортируем лист "Поставки"
  var suppliesSheet = ss.getSheetByName('Поставки');
  if (suppliesSheet) {
    sortSheetByDate(suppliesSheet, 'Дата поставки');
  }
  
  // Сортируем лист "Детали_поставок"
  var detailsSheet = ss.getSheetByName('Детали_поставок');
  if (detailsSheet) {
    sortSheetByDate(detailsSheet, 'supplyDate');
  }
  
  // Сортируем лист "Товары_поставок"
  var goodsSheet = ss.getSheetByName('Товары_поставок');
  if (goodsSheet) {
    // Для товаров поставок сортируем по supplyID (связано с датой поставки)
    sortSheetByColumn(goodsSheet, 'supplyID');
  }
}

// Вспомогательная функция сортировки по дате
function sortSheetByDate(sheet, dateColumnName) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var dateColIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [dateColumnName]) : headers.indexOf(dateColumnName);
  
  if (dateColIndex === -1) {
    Logger.log('Колонка "' + dateColumnName + '" не найдена в листе ' + sheet.getName());
    return;
  }
  
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({column: dateColIndex + 1, ascending: true});
}

// Вспомогательная функция сортировки по колонке
function sortSheetByColumn(sheet, columnName) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIndex = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [columnName]) : headers.indexOf(columnName);
  
  if (colIndex === -1) return;
  
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({column: colIndex + 1, ascending: true});
}








// ============================================
// ЗАГРУЗКА ВЕСОВ С ЛИСТА "Товар+Вес"
// ============================================

function loadWeightsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Товар+Вес");
  if (!sheet) {
    Logger.log('❌ Лист "Товар+Вес" не найден');
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  
  const headers = data[0];
  
  let artikulIndex = 0;
  let weightIndex = 1;
  
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i]?.toString().toLowerCase() || '';
    if (header === 'артикул') artikulIndex = i;
    if (header === 'вес, кг за 1 шт') weightIndex = i;
  }
  
  const weightsMap = {};
  for (let i = 1; i < data.length; i++) {
    const artikul = data[i][artikulIndex];
    const weight = data[i][weightIndex];
    
    if (artikul && weight !== '' && weight !== null && weight !== undefined) {
      const artikulStr = artikul.toString().trim();
      weightsMap[artikulStr] = parseFloat(weight) || 0;
    }
  }
  
  return weightsMap;
}

// ============================================
// ЗАГРУЗКА ДАННЫХ С ЛИСТА "Товары"
// ============================================

function loadGoodsData_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Товары");
  if (!sheet) { Logger.log('❌ Лист "Товары" не найден'); return {}; }
  
  const data = sheet.getDataRange().getValues();
  const map = {};
  
  for (let i = 1; i < data.length; i++) {
    // 🔥 Жёстко читаем колонки по индексам
    const colA_nmId = data[i][0];   // Колонка A: nmID
    const colC_art = data[i][2];    // Колонка C: Ваш артикул
    
    const key = colA_nmId != null ? String(colA_nmId).trim() : '';
    if (!key) continue;
    
    map[key] = {
      vendorCode: colC_art ? String(colC_art).trim() : '',
      brand: data[i][3] || '',
      category: data[i][4] || '',
      barcode: data[i][5] || '',
      price: parseFloat(data[i][6]) || 0,
      discountPrice: parseFloat(data[i][7]) || 0
    };
  }
  
  return map;
}

// ============================================
// ПОЛУЧЕНИЕ ОСТАТКОВ (РАБОЧИЙ МЕТОД)
// ============================================

function getAllStocks() {
  const token = getApiToken();
  const url = 'https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses';
  
  let allItems = [];
  let offset = 0;
  const limit = 100000;
  
  while (true) {
    const payload = { limit: limit, offset: offset };
    const options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) break;
      
      const data = JSON.parse(response.getContentText());
      const items = data.data?.items || [];
      
      if (items.length === 0) break;
      
      allItems = allItems.concat(items);
      offset += limit;
      
      if (items.length < limit) break;
      
    } catch (error) {
      Logger.log('❌ Остатки API: ' + error);
      break;
    }
  }
  
  return allItems;
}

// ============================================
// СОХРАНЕНИЕ ДАННЫХ "В ПУТИ" (ПОЛНАЯ ЗАМЕНА)
// ============================================

function saveStocksHistory() {
  const stocks = getAllStocks();
  if (!stocks || stocks.length === 0) { Logger.log('❌ Нет данных об остатках'); return; }
  
  const goodsMap = loadGoodsData_v2(); // 🔥 Вызываем новую функцию
  const weightsMap = loadWeightsMap();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'В пути';
  let sheet = ss.getSheetByName(sheetName);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const headers = [
    'Дата', 'Артикул WB', 'Артикул продавца', 'Бренд', 'Категория', 'Штрих-код',
    'Цена, ₽', 'Цена со скидкой, ₽', 'Вес, кг за 1 шт', 'Склад', 'Регион',
    'Остатки (шт)', 'В пути к покупателю', 'В пути от покупателя',
    'В пути к покупателю Вес, кг', 'В пути от покупателя Вес, кг', 'Остатки вес, кг'
  ];
  
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clear();
  
  const newData = [headers];
  
  for (let i = 0; i < stocks.length; i++) {
    const item = stocks[i];
    const rawNmId = item.nmId;
    const nmId = rawNmId != null ? String(rawNmId).trim() : '';
    
    const goodsInfo = goodsMap[nmId] || {};
    const vendorCode = goodsInfo.vendorCode || '';
    const weightPerUnit = weightsMap[vendorCode] || 0;
    
    newData.push([
      today, nmId, vendorCode,
      goodsInfo.brand || '', goodsInfo.category || '', goodsInfo.barcode || '',
      goodsInfo.price || 0, goodsInfo.discountPrice || 0,
      weightPerUnit,
      item.warehouseName || '', item.regionName || '',
      item.quantity || 0, item.inWayToClient || 0, item.inWayFromClient || 0,
      weightPerUnit * (item.inWayToClient || 0),
      weightPerUnit * (item.inWayFromClient || 0),
      weightPerUnit * (item.quantity || 0)
    ]);
  }
  
  sheet.getRange(1, 1, newData.length, headers.length).setValues(newData);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f1f1');
  sheet.autoResizeColumns(1, headers.length);
  wbPutSheetFilter_(sheet);
  
  Logger.log('✅ В пути: ' + (newData.length - 1) + ' строк');
}

// ============================================
// РУЧНОЙ ЗАПУСК
// ============================================

function runManualUpdate() {
  saveStocksHistory();
  SpreadsheetApp.getUi().alert('✅ Готово', 'Лист "В пути" обновлён', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================
// Остатки ЛК: отчёт кабинета warehouse_remains.
// Длинный лист, как «Остатки»: день дописывается, повтор за сегодня перезаписывает этот день.
// Листы «Остатки» и «В пути» не меняет.
// ============================================
function probeWarehouseRemains() {
  recordLkStocks();
}

function recordLkStocks() {
  var token = getApiToken();
  if (!token) return;

  var host = 'https://seller-analytics-api.wildberries.ru';
  var dateStr = Utilities.formatDate(new Date(), sheetTimeZone_(), 'yyyy-MM-dd');
  Logger.log('Остатки ЛК, дата среза ' + dateStr);

  var createUrl = host + '/api/v1/warehouse_remains'
    + '?locale=ru&groupByNm=true&groupBySa=true&groupBySize=true';
  var created = wrFetchJson_(token, createUrl);
  Logger.log('create HTTP ' + created.code + ' ' + String(created.text || '').substring(0, 400));
  if (created.code === 429) {
    lkStocksFail_('WB лимит: 1 создание отчёта в минуту. Подождите минуту и запустите ещё раз.');
    return;
  }
  if (created.code !== 200) {
    lkStocksFail_('Не удалось создать отчёт: HTTP ' + created.code + '\n' + String(created.text || '').substring(0, 300));
    return;
  }

  var taskId = created.data && created.data.data && created.data.data.taskId;
  if (!taskId && created.data && created.data.taskId) taskId = created.data.taskId;
  if (!taskId) {
    lkStocksFail_('В ответе нет taskId:\n' + String(created.text || '').substring(0, 400));
    return;
  }
  Logger.log('taskId ' + taskId);

  var status = '';
  var i;
  for (i = 0; i < 24; i++) {
    Utilities.sleep(8000);
    var st = wrFetchJson_(token, host + '/api/v1/warehouse_remains/tasks/' + encodeURIComponent(taskId) + '/status');
    status = '';
    if (st.data && st.data.data && st.data.data.status) status = String(st.data.data.status);
    else if (st.data && st.data.status) status = String(st.data.status);
    Logger.log('status ' + (i + 1) + ': HTTP ' + st.code + ' ' + status + ' ' + String(st.text || '').substring(0, 200));
    if (st.code === 429) {
      Utilities.sleep(6000);
      continue;
    }
    if (String(status).toLowerCase() === 'done') break;
    if (/error|fail|cancel/i.test(status)) {
      lkStocksFail_('Отчёт не собрался, статус: ' + status);
      return;
    }
  }
  if (String(status).toLowerCase() !== 'done') {
    lkStocksFail_('Отчёт ещё не готов (статус: ' + (status || 'пусто') + '). Запустите «Остатки ЛК» позже.');
    return;
  }

  var dl = wrFetchJson_(token, host + '/api/v1/warehouse_remains/tasks/' + encodeURIComponent(taskId) + '/download');
  Logger.log('download HTTP ' + dl.code + ' bytes ' + String(dl.text || '').length);
  if (dl.code === 204 || !dl.text) {
    lkStocksFail_('Отчёт пустой (HTTP ' + dl.code + ')');
    return;
  }
  if (dl.code !== 200) {
    lkStocksFail_('Скачать отчёт не удалось: HTTP ' + dl.code + '\n' + String(dl.text || '').substring(0, 300));
    return;
  }

  var rowsJson = dl.data;
  if (rowsJson && !Array.isArray(rowsJson) && Array.isArray(rowsJson.data)) rowsJson = rowsJson.data;
  if (!Array.isArray(rowsJson) || !rowsJson.length) {
    lkStocksFail_('В отчёте нет строк. Ответ:\n' + String(dl.text || '').substring(0, 400));
    return;
  }

  var written = lkStocksWrite_(dateStr, rowsJson);
  showAlert(
    '✅ Остатки ЛК\n' +
    'Дата: ' + dateStr + '\n' +
    'Строк: ' + written.rows + '\n' +
    'Складов: ' + written.warehouses + '\n' +
    'Лист: Остатки ЛК\n\n' +
    'Повтор за сегодня перезаписывает только эту дату.' +
    (written.oldLeft ? '\n\nСтарый лист «Остатки_ЛК_проба» не трогал.' : '')
  );
}

function lkStocksFail_(msg) {
  Logger.log(msg);
  if (isRunningFromTrigger()) throw new Error(msg);
  showAlert(msg);
}

var LK_STOCKS_SHEET_ = 'Остатки ЛК';
var LK_STOCKS_SHEET_OLD_ = 'Остатки_ЛК_проба';
var LK_STOCKS_HEADERS_ = [
  'Дата', 'Артикул продавца', 'Название', 'Артикул WB', 'Предмет', 'Бренд', 'Размер',
  'Склад', 'Остаток', 'Регион поставки', 'Вес, кг за 1 шт', 'Общий вес, кг',
  'Объём', 'Пострадавший'
];
var LK_STOCKS_HEADERS_VICTIM_AFTER_REGION_ = [
  'Дата', 'Артикул продавца', 'Название', 'Артикул WB', 'Предмет', 'Бренд', 'Размер',
  'Склад', 'Остаток', 'Регион поставки', 'Пострадавший', 'Вес, кг за 1 шт', 'Общий вес, кг',
  'Объём'
];
var LK_STOCKS_HEADERS_MID_ = [
  'Дата', 'Артикул продавца', 'Название', 'Артикул WB', 'Предмет', 'Бренд', 'Размер',
  'Склад', 'Остаток', 'Регион поставки', 'Вес, кг за 1 шт', 'Общий вес, кг',
  'Объём'
];
var LK_STOCKS_HEADERS_OLD_ = [
  'Дата', 'Артикул продавца', 'Название', 'Артикул WB', 'Предмет', 'Бренд', 'Размер',
  'Склад', 'Остаток', 'Регион поставки', 'Вес, кг за 1 шт', 'Общий вес, кг',
  'Штрихкод', 'Объём', 'В пути к покупателю', 'В пути возврат'
];

/** Служебные имена отчёта ЛК — это не склады. */
function lkWhKind_(name) {
  var n = String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!n) return '';
  if (n.indexOf('в пути до получател') === 0) return 'toClient';
  if (n.indexOf('в пути возврат') === 0) return 'fromClient';
  if (n.indexOf('всего находится на складах') === 0) return 'total';
  return '';
}

function sheetTimeZone_() {
  try {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    if (tz) return tz;
  } catch (e) {}
  return Session.getScriptTimeZone() || 'Europe/Samara';
}

/** Календарный день в часовом поясе таблицы: yyyy-MM-dd. */
function sheetDayKey_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, sheetTimeZone_(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var iso = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var dmy = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (dmy) return dmy[3] + '-' + ('0' + dmy[2]).slice(-2) + '-' + ('0' + dmy[1]).slice(-2);
  return s.substring(0, 10);
}

/** Полночь в часовом поясе таблицы. Без 12:00:00, чтобы дата совпадала с ДАТА() в формулах. */
function lkStocksDate_(val) {
  var ymd = sheetDayKey_(val);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return val;
  return Utilities.parseDate(ymd + ' 00:00:00', sheetTimeZone_(), 'yyyy-MM-dd HH:mm:ss');
}

function lkVictimMap_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Регион поставки');
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var whIdx = -1;
  var vIdx = -1;
  var i, h;
  for (i = 0; i < headers.length; i++) {
    h = String(headers[i] || '').trim().toLowerCase();
    if (h === 'склад' && whIdx < 0) whIdx = i;
    if (h === 'пострадавший') vIdx = i;
  }
  if (whIdx < 0) whIdx = 0;
  if (vIdx < 0) {
    Logger.log('Остатки ЛК: на листе «Регион поставки» нет столбца Пострадавший');
    return map;
  }
  var name, raw, flag;
  for (i = 1; i < data.length; i++) {
    name = String(data[i][whIdx] || '').trim().toLowerCase();
    if (!name) continue;
    raw = data[i][vIdx];
    flag = String(raw || '').trim().toLowerCase();
    map[name] = (flag === 'да' || raw === true) ? 'да' : '';
  }
  return map;
}

function lkVictim_(warehouse, map) {
  if (!warehouse || !map) return '';
  var name = String(warehouse).trim().toLowerCase();
  if (!name) return '';
  if (Object.prototype.hasOwnProperty.call(map, name)) return map[name];
  var key;
  for (key in map) {
    if (name.indexOf(key) !== -1) return map[key];
  }
  return '';
}

function lkProduct_(dict, nmId) {
  if (!dict) return {};
  if (dict[nmId]) return dict[nmId];
  var s = nmId != null ? String(nmId).trim() : '';
  if (s && dict[s]) return dict[s];
  var n = Number(s);
  if (s && !isNaN(n) && dict[n]) return dict[n];
  return {};
}

function lkStocksEnsureSheet_(ss) {
  var sheet = ss.getSheetByName(LK_STOCKS_SHEET_);
  if (sheet) return sheet;
  var old = ss.getSheetByName(LK_STOCKS_SHEET_OLD_);
  if (old) {
    old.setName(LK_STOCKS_SHEET_);
    return old;
  }
  return ss.insertSheet(LK_STOCKS_SHEET_);
}

function lkStocksHeadersMatch_(sheet, expected) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < expected.length) return false;
  var h = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  var i;
  for (i = 0; i < expected.length; i++) {
    if (String(h[i] || '').trim() !== expected[i]) return false;
  }
  return true;
}

function lkStocksHeadersOk_(sheet) {
  return lkStocksHeadersMatch_(sheet, LK_STOCKS_HEADERS_);
}

function lkStocksRewrite_(sheet, out) {
  sheet.clear();
  sheet.getRange(1, 1, out.length, LK_STOCKS_HEADERS_.length).setValues(out);
  stocksSafeFreezeRows_(sheet, 1);
}

/** Снимает штрихкод, столбцы «в пути» и строки с пустым складом, дни сохраняет. */
function lkStocksMigrateOld_(sheet, victimMap) {
  var lastRow = sheet.getLastRow();
  var width = Math.max(sheet.getLastColumn(), LK_STOCKS_HEADERS_OLD_.length);
  var data = sheet.getRange(1, 1, lastRow, width).getValues();
  var out = [LK_STOCKS_HEADERS_];
  var i, row, wh;
  for (i = 1; i < data.length; i++) {
    row = data[i];
    wh = String(row[7] || '').trim();
    if (!wh) continue;
    out.push([
      lkStocksDate_(row[0]), row[1], row[2], row[3], row[4], row[5], row[6],
      wh, row[8], row[9], row[10], row[11],
      row[13], lkVictim_(wh, victimMap)
    ]);
  }
  lkStocksRewrite_(sheet, out);
  Logger.log('Остатки ЛК: старая шапка переписана, строк без пустого склада ' + (out.length - 1));
}

/** Добавляет «Пострадавший» в уже длинный лист без этого столбца. */
function lkStocksMigrateMid_(sheet, victimMap) {
  var lastRow = sheet.getLastRow();
  var width = Math.max(sheet.getLastColumn(), LK_STOCKS_HEADERS_MID_.length);
  var data = sheet.getRange(1, 1, lastRow, width).getValues();
  var out = [LK_STOCKS_HEADERS_];
  var i, row, wh;
  for (i = 1; i < data.length; i++) {
    row = data[i];
    wh = String(row[7] || '').trim();
    if (!wh) continue;
    out.push([
      lkStocksDate_(row[0]), row[1], row[2], row[3], row[4], row[5], row[6],
      wh, row[8], row[9], row[10], row[11], row[12],
      lkVictim_(wh, victimMap)
    ]);
  }
  lkStocksRewrite_(sheet, out);
  Logger.log('Остатки ЛК: добавлен столбец Пострадавший, строк ' + (out.length - 1));
}

/** Переносит «Пострадавший» из середины в последний столбец. */
function lkStocksMigrateVictimPos_(sheet) {
  var lastRow = sheet.getLastRow();
  var width = Math.max(sheet.getLastColumn(), LK_STOCKS_HEADERS_VICTIM_AFTER_REGION_.length);
  var data = sheet.getRange(1, 1, lastRow, width).getValues();
  var out = [LK_STOCKS_HEADERS_];
  var i, row, wh;
  for (i = 1; i < data.length; i++) {
    row = data[i];
    wh = String(row[7] || '').trim();
    if (!wh) continue;
    out.push([
      lkStocksDate_(row[0]), row[1], row[2], row[3], row[4], row[5], row[6],
      wh, row[8], row[9], row[11], row[12], row[13], row[10]
    ]);
  }
  lkStocksRewrite_(sheet, out);
  Logger.log('Остатки ЛК: столбец Пострадавший перенесён в конец, строк ' + (out.length - 1));
}

/** Формат только указанных строк. История уже оформлена, повторно её не переписываем. */
function lkStocksFormat_(sheet, startRow, rowCount) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastCol = LK_STOCKS_HEADERS_.length;
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#ffffff');
  stocksSafeFreezeRows_(sheet, 1);
  if (!startRow || startRow < 2 || !rowCount || rowCount < 1) return;
  var body = sheet.getRange(startRow, 1, rowCount, lastCol);
  body.setFontFamily('Arial').setFontSize(10);
  sheet.getRange(startRow, 1, rowCount, 1).setNumberFormat('dd.mm.yyyy');
  var wUnit = LK_STOCKS_HEADERS_.indexOf('Вес, кг за 1 шт') + 1;
  var wSum = LK_STOCKS_HEADERS_.indexOf('Общий вес, кг') + 1;
  var qty = LK_STOCKS_HEADERS_.indexOf('Остаток') + 1;
  if (qty > 0) sheet.getRange(startRow, qty, rowCount, 1).setNumberFormat('0');
  if (wUnit > 0) sheet.getRange(startRow, wUnit, rowCount, 1).setNumberFormat('0.000');
  if (wSum > 0) sheet.getRange(startRow, wSum, rowCount, 1).setNumberFormat('0.00');
}

function lkStocksWrite_(dateStr, items) {
  var productsDict = getProductsDictionary();
  var regionMap = getRegionMap();
  var weightsMap = loadWeightsMapForStocks();
  var victimMap = lkVictimMap_();
  var rows = [];
  var seenWh = {};
  var regionCache = {};
  var victimCache = {};
  var day = lkStocksDate_(dateStr);
  var i, w, it, whs, name, kind, q, product, vendorCode, weightPerUnit;
  var nmId, title, subject, brand, size, volume;
  var whQty, whOrder, amount, region, totalWeight;

  for (i = 0; i < items.length; i++) {
    it = items[i] || {};
    nmId = it.nmId != null ? it.nmId : '';
    product = lkProduct_(productsDict, nmId);
    vendorCode = it.vendorCode ? String(it.vendorCode).trim() : '';
    if (!vendorCode && product.vendorCode) vendorCode = String(product.vendorCode).trim();
    weightPerUnit = (vendorCode && weightsMap[vendorCode]) ? weightsMap[vendorCode] : 0;
    title = product.title || '';
    subject = it.subjectName || product.subject || '';
    brand = it.brand || product.brand || '';
    size = it.techSize != null && it.techSize !== '' ? String(it.techSize) : '';
    volume = it.volume != null && it.volume !== '' ? it.volume : '';

    whQty = {};
    whOrder = [];
    whs = it.warehouses || [];
    for (w = 0; w < whs.length; w++) {
      name = whs[w] && whs[w].warehouseName ? String(whs[w].warehouseName).trim() : '';
      if (!name) continue;
      kind = lkWhKind_(name);
      q = Number(whs[w].quantity);
      if (isNaN(q)) q = 0;
      if (kind) continue;
      if (whQty[name] == null) whOrder.push(name);
      whQty[name] = (whQty[name] || 0) + q;
    }

    for (w = 0; w < whOrder.length; w++) {
      name = whOrder[w];
      amount = whQty[name];
      if (!(amount > 0)) continue;
      seenWh[name] = true;
      if (!Object.prototype.hasOwnProperty.call(regionCache, name)) {
        regionCache[name] = getRegionFromMap(name, regionMap);
        victimCache[name] = lkVictim_(name, victimMap);
      }
      region = regionCache[name];
      totalWeight = weightPerUnit * amount;
      rows.push([
        day, vendorCode, title, nmId, subject, brand, size,
        name, amount, region, weightPerUnit, totalWeight,
        volume, victimCache[name]
      ]);
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = lkStocksEnsureSheet_(ss);
  var oldLeft = !!ss.getSheetByName(LK_STOCKS_SHEET_OLD_);
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  var rebuilt = false;
  if (lkStocksHeadersMatch_(sheet, LK_STOCKS_HEADERS_OLD_)) {
    lkStocksMigrateOld_(sheet, victimMap);
    rebuilt = true;
  } else if (lkStocksHeadersMatch_(sheet, LK_STOCKS_HEADERS_MID_)) {
    lkStocksMigrateMid_(sheet, victimMap);
    rebuilt = true;
  } else if (lkStocksHeadersMatch_(sheet, LK_STOCKS_HEADERS_VICTIM_AFTER_REGION_)) {
    lkStocksMigrateVictimPos_(sheet);
    rebuilt = true;
  }
  if (!lkStocksHeadersOk_(sheet)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, LK_STOCKS_HEADERS_.length).setValues([LK_STOCKS_HEADERS_]);
    stocksSafeFreezeRows_(sheet, 1);
    sheet.setTabColor(null);
    rebuilt = true;
  }

  var removed = stocksRemoveDate_(sheet, LK_STOCKS_HEADERS_, dateStr);
  var startRow = 0;
  if (rows.length) {
    startRow = sheet.getLastRow() + 1;
    var b, batch;
    for (b = 0; b < rows.length; b += 1000) {
      batch = rows.slice(b, b + 1000);
      sheet.getRange(startRow + b, 1, batch.length, LK_STOCKS_HEADERS_.length).setValues(batch);
    }
  }
  if (rebuilt && sheet.getLastRow() > 1) lkStocksFormat_(sheet, 2, sheet.getLastRow() - 1);
  else if (startRow) lkStocksFormat_(sheet, startRow, rows.length);
  wbPutSheetFilter_(sheet);

  var whNames = Object.keys(seenWh);
  whNames.sort(function(a, c) { return a.localeCompare(c, 'ru'); });
  Logger.log('Остатки ЛК: снято ' + removed + ', записано ' + rows.length + ', склады (' + whNames.length + '): ' + whNames.join(' | '));
  return {
    rows: rows.length,
    warehouses: whNames.length,
    removed: removed,
    oldLeft: oldLeft
  };
}

function wrFetchJson_(token, url, method, body) {
  var options = {
    method: method || 'GET',
    headers: { 'Authorization': token, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);
  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var text = response.getContentText() || '';
  var data = null;
  if (text) {
    try { data = JSON.parse(text); } catch (e) { data = null; }
  }
  return { code: code, text: text, data: data };
}

// ============================================
// ПРОБА: CSV STOCK_HISTORY_DAILY_CSV (история остатков analytics)
// Сравнить склады с листом Остатки ЛК и wb-warehouses.
// ============================================

function probeStocksHistoryCsv() {
  var tz = Session.getScriptTimeZone();
  var y = new Date();
  y.setDate(y.getDate() - 1);
  var dayStr = Utilities.formatDate(y, tz, 'yyyy-MM-dd');

  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box}' +
    'button{background:#4CAF50;color:white;padding:10px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e}' +
    '.info{background:#e8f5e9;padding:8px;margin:0 0 10px;border-radius:4px;font-size:11px;line-height:1.4}' +
    '#st{margin-top:10px;font-size:12px;color:#555}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📦 Остатки CSV (проба)</h3>' +
    '<div class="info">STOCK_HISTORY_DAILY_CSV — история analytics, до ~3 мес.<br>' +
    'Пишет листы «Остатки_CSV_проба» и «Остатки_склады_сравнение».</div>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom" value="' + dayStr + '">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo" value="' + dayStr + '">' +
    '<button id="btn" onclick="go()">📥 Загрузить</button>' +
    '<div id="st"></div>' +
    '<script>' +
    'function go(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'if(!d1||!d2){alert("Выберите даты");return;}' +
    'document.getElementById("btn").disabled=true;' +
    'document.getElementById("st").textContent="Отчёт генерируется… 1–3 мин";' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){document.getElementById("st").textContent=msg||"Готово";document.getElementById("btn").disabled=false;alert(msg||"Готово");google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("btn").disabled=false;document.getElementById("st").textContent="";alert(String(e));})' +
    '.probeStocksHistoryCsvRun(d1,d2);' +
    '}' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(360), 'Остатки CSV (проба)');
}

/** Без _ в конце: иначе HtmlService не вызывает функцию. */
function probeStocksHistoryCsvRun(dateFrom, dateTo) {
  var token = getApiToken();
  if (!token) throw new Error('Нет API токена. Сначала меню → Ввести API токен.');

  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('Книга занята. Закройте лишние вкладки и повторите.\n' + e);
  }

  var d1 = String(dateFrom || '').substring(0, 10);
  var d2 = String(dateTo || '').substring(0, 10);
  if (!d1 || !d2) throw new Error('Не выбраны даты');
  if (d1 > d2) {
    var tmp = d1;
    d1 = d2;
    d2 = tmp;
  }

  var host = 'https://seller-analytics-api.wildberries.ru';
  var downloadId = Utilities.getUuid();
  var createRes = wrFetchJson_(token, host + '/api/v2/nm-report/downloads', 'POST', {
    id: downloadId,
    reportType: 'STOCK_HISTORY_DAILY_CSV',
    userReportName: 'stocks csv probe ' + d1 + '..' + d2,
    params: {
      currentPeriod: { start: d1, end: d2 },
      stockType: '',
      skipDeletedNm: true
    }
  });
  Logger.log('csv create HTTP ' + createRes.code + ' ' + String(createRes.text || '').substring(0, 400));
  if (createRes.code === 429) throw new Error('WB лимит. Подождите 20 сек и повторите.');
  if (createRes.code === 403) throw new Error('HTTP 403 — нужен токен с категорией «Аналитика».\n' + String(createRes.text || '').substring(0, 300));
  if (createRes.code !== 200) {
    throw new Error('Не удалось создать CSV: HTTP ' + createRes.code + '\n' + String(createRes.text || '').substring(0, 400));
  }

  var status = shCsvPollUntilReady_(token, downloadId);
  if (status !== 'SUCCESS') throw new Error('CSV не готов: ' + (status || 'таймаут'));

  var csvText = shCsvFetchZipCsv_(token, downloadId);
  if (!csvText) throw new Error('ZIP без CSV или пустой ответ');

  var parsed = shCsvParse_(csvText);
  var written = shCsvWriteSheets_(ss, d1, d2, parsed);
  var cmp = shCsvCompareWarehouses_(ss, parsed.warehouses);

  return '✅ Остатки CSV (проба)\n' +
    'Период: ' + d1 + ' — ' + d2 + '\n' +
    'Формат CSV: ' + (parsed.format || '?') + '\n' +
    'Строк CSV: ' + parsed.rowsIn + ' → на лист ' + written.rows + '\n' +
    'Складов в CSV: ' + parsed.warehouses.length + '\n' +
    'Шапка: ' + parsed.headerLine + '\n\n' +
    'Листы: Остатки_CSV_проба, Остатки_склады_сравнение\n' +
    (cmp.lkFound ? ('Складов на ЛК-листе: ' + cmp.lkCount + ', совпало имён: ' + cmp.exact + '\n') : 'Лист «Остатки ЛК» не найден — сначала запустите «Остатки ЛК»\n') +
    'Склады CSV: ' + parsed.warehouses.slice(0, 12).join(', ') +
    (parsed.warehouses.length > 12 ? ' …' : '');
}

function shCsvPollUntilReady_(token, downloadId) {
  var host = 'https://seller-analytics-api.wildberries.ru';
  var url = host + '/api/v2/nm-report/downloads?filter[downloadIds]=' + encodeURIComponent(downloadId);
  var i, st, status;
  for (i = 0; i < 90; i++) {
    Utilities.sleep(i === 0 ? 5000 : 8000);
    st = wrFetchJson_(token, url, 'GET');
    if (st.code === 429) {
      Utilities.sleep(15000);
      continue;
    }
    if (st.code !== 200) {
      Logger.log('csv poll HTTP ' + st.code + ' ' + String(st.text || '').substring(0, 200));
      continue;
    }
    var item = (st.data && st.data.data && st.data.data[0]) ? st.data.data[0] : ((st.data && st.data[0]) ? st.data[0] : null);
    status = String((item && item.status) ? item.status : '').toUpperCase();
    Logger.log('csv poll ' + (i + 1) + ': ' + status);
    if (status === 'SUCCESS') return 'SUCCESS';
    if (status === 'FAILED') return 'FAILED';
  }
  return status || 'TIMEOUT';
}

function shCsvFetchZipCsv_(token, downloadId) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads/file/' + encodeURIComponent(downloadId);
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { Authorization: token },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  Logger.log('csv zip HTTP ' + code);
  if (code !== 200) throw new Error('Скачать ZIP не удалось: HTTP ' + code);
  var files = Utilities.unzip(response.getBlob());
  var f, name, text;
  for (f = 0; f < files.length; f++) {
    name = files[f].getName() || '';
    if (name.toLowerCase().indexOf('.csv') >= 0) {
      text = files[f].getDataAsString('UTF-8');
      if (text.charCodeAt(0) === 0xFEFF) text = text.substring(1);
      return text;
    }
  }
  return '';
}

function shCsvParseLine_(line) {
  var out = [];
  var cur = '';
  var q = false;
  var c, i;
  for (i = 0; i < line.length; i++) {
    c = line[i];
    if (c === '"') {
      q = !q;
      continue;
    }
    if (c === ',' && !q) {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += c;
  }
  out.push(cur);
  return out;
}

function shCsvHeaderIdx_(header, names) {
  var lower = header.map(function(h) { return String(h || '').trim().toLowerCase(); });
  var n, i;
  for (n = 0; n < names.length; n++) {
    i = lower.indexOf(String(names[n]).toLowerCase());
    if (i >= 0) return i;
  }
  return -1;
}

function shCsvLooksLikeDate_(s) {
  var t = String(s || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return true;
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(t)) return true;
  return false;
}

function shCsvDateYmd_(s) {
  var t = String(s || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.substring(0, 10);
  var m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  return '';
}

function shCsvParseWide_(header, lines) {
  var iNm = shCsvHeaderIdx_(header, ['nmID', 'nmId', 'nm_id']);
  var iWh = shCsvHeaderIdx_(header, ['officeName', 'office_name', 'warehouseName', 'warehouse_name', 'склад']);
  var iSize = shCsvHeaderIdx_(header, ['sizeName', 'techSize', 'size']);
  var iBrand = shCsvHeaderIdx_(header, ['brandName', 'brand']);
  var iVc = shCsvHeaderIdx_(header, ['vendorCode', 'supplierArticle']);
  if (iNm < 0 || iWh < 0) return null;

  var dateCols = [];
  var h;
  for (h = 0; h < header.length; h++) {
    if (shCsvLooksLikeDate_(header[h])) dateCols.push({ idx: h, ymd: shCsvDateYmd_(header[h]) });
  }
  if (!dateCols.length) return null;

  var whSeen = {};
  var warehouses = [];
  var rows = [];
  var rowsIn = 0;
  var i, cols, nm, wh, d, qty, key, prev, r, dc;

  for (i = 1; i < lines.length; i++) {
    cols = shCsvParseLine_(lines[i]);
    nm = String(cols[iNm] || '').trim();
    wh = String(cols[iWh] || '').trim();
    if (!nm || !wh) continue;
    if (!whSeen[wh]) {
      whSeen[wh] = true;
      warehouses.push(wh);
    }
    for (d = 0; d < dateCols.length; d++) {
      dc = dateCols[d];
      qty = cols[dc.idx];
      if (qty === '' || qty == null) continue;
      rowsIn++;
      key = dc.ymd + '|' + nm + '|' + wh;
      prev = null;
      for (r = 0; r < rows.length; r++) {
        if (rows[r]._key === key) {
          prev = rows[r];
          break;
        }
      }
      if (prev) {
        prev.quantity = Number(prev.quantity || 0) + Number(qty || 0);
      } else {
        rows.push({
          _key: key,
          snapshot_date: dc.ymd,
          nm_id: nm,
          warehouse_name: wh,
          quantity: Number(qty || 0),
          tech_size: iSize >= 0 ? cols[iSize] : '',
          vendor_code: iVc >= 0 ? cols[iVc] : '',
          brand: iBrand >= 0 ? cols[iBrand] : ''
        });
      }
    }
  }
  warehouses.sort(function(a, b) { return String(a).localeCompare(String(b), 'ru'); });
  return {
    rowsIn: rowsIn,
    rows: rows,
    warehouses: warehouses,
    headerLine: header.join(', '),
    format: 'wide'
  };
}

function shCsvParse_(csv) {
  var text = String(csv || '').trim();
  if (!text) return { rowsIn: 0, rows: [], warehouses: [], headerLine: '', format: 'empty' };
  var lines = text.split(/\r?\n/).filter(function(l) { return String(l).trim(); });
  if (lines.length < 2) return { rowsIn: 0, rows: [], warehouses: [], headerLine: lines[0] || '', format: 'empty' };

  var header = shCsvParseLine_(lines[0]).map(function(h) { return String(h).trim(); });
  var iDt = shCsvHeaderIdx_(header, ['dt', 'date', 'snapshot_date']);
  var iNm = shCsvHeaderIdx_(header, ['nmID', 'nmId', 'nm_id']);
  var iWh = shCsvHeaderIdx_(header, ['warehouseName', 'warehouse_name', 'officeName', 'office_name', 'склад']);
  var iQty = shCsvHeaderIdx_(header, ['quantity', 'stockQty', 'stockCount', 'qty', 'остаток']);
  var iSize = shCsvHeaderIdx_(header, ['techSize', 'sizeName', 'size', 'chrtID', 'chrtId']);
  var iBrand = shCsvHeaderIdx_(header, ['brandName', 'brand']);
  var iVc = shCsvHeaderIdx_(header, ['vendorCode', 'supplierArticle']);

  if (iDt < 0 || iNm < 0 || iWh < 0 || iQty < 0) {
    var wide = shCsvParseWide_(header, lines);
    if (wide) return wide;
    throw new Error('Не распознана шапка CSV:\n' + header.join(', '));
  }

  var whSeen = {};
  var warehouses = [];
  var rows = [];
  var rowsIn = 0;
  var i, cols, dt, nm, wh, qty, key, prev;

  for (i = 1; i < lines.length; i++) {
    cols = shCsvParseLine_(lines[i]);
    dt = shCsvDateYmd_(cols[iDt]) || String(cols[iDt] || '').substring(0, 10);
    nm = String(cols[iNm] || '').trim();
    wh = String(cols[iWh] || '').trim();
    qty = cols[iQty];
    if (!dt || !nm || !wh) continue;
    rowsIn++;
    if (!whSeen[wh]) {
      whSeen[wh] = true;
      warehouses.push(wh);
    }
    key = dt + '|' + nm + '|' + wh;
    prev = null;
    for (var r = 0; r < rows.length; r++) {
      if (rows[r]._key === key) {
        prev = rows[r];
        break;
      }
    }
    if (prev) {
      prev.quantity = Number(prev.quantity || 0) + Number(qty || 0);
    } else {
      rows.push({
        _key: key,
        snapshot_date: dt,
        nm_id: nm,
        warehouse_name: wh,
        quantity: Number(qty || 0),
        tech_size: iSize >= 0 ? cols[iSize] : '',
        vendor_code: iVc >= 0 ? cols[iVc] : '',
        brand: iBrand >= 0 ? cols[iBrand] : ''
      });
    }
  }
  warehouses.sort(function(a, b) { return String(a).localeCompare(String(b), 'ru'); });
  return { rowsIn: rowsIn, rows: rows, warehouses: warehouses, headerLine: header.join(', '), format: 'long' };
}

function shCsvWriteSheets_(ss, dateFrom, dateTo, parsed) {
  var sheetName = 'Остатки_CSV_проба';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = ['Период с', 'Период по', 'Дата', 'Артикул WB', 'Артикул продавца', 'Бренд', 'Размер', 'Склад', 'Остаток'];
  var out = [headers];
  var maxRows = 80000;
  var n = Math.min(parsed.rows.length, maxRows);
  var i, row;
  for (i = 0; i < n; i++) {
    row = parsed.rows[i];
    out.push([
      dateFrom, dateTo,
      row.snapshot_date, row.nm_id, row.vendor_code, row.brand, row.tech_size,
      row.warehouse_name, row.quantity
    ]);
  }
  sheet.getRange(1, 1, out.length, headers.length).setValues(out);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#d9ead3');
  wbPutSheetFilter_(sheet);
  return { rows: n, truncated: parsed.rows.length > maxRows };
}

function shCsvLkWarehouses_(ss) {
  var sheet = ss.getSheetByName('Остатки ЛК') || ss.getSheetByName('Остатки_ЛК_проба');
  if (!sheet || sheet.getLastRow() < 1) return [];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var whIdx = -1;
  var h, name, i;
  for (h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() === 'Склад') { whIdx = h; break; }
  }
  if (whIdx >= 0 && sheet.getLastRow() >= 2) {
    var vals = sheet.getRange(2, whIdx + 1, sheet.getLastRow() - 1, 1).getValues();
    var seen = {};
    var wh = [];
    for (i = 0; i < vals.length; i++) {
      name = String(vals[i][0] || '').trim();
      if (!name || seen[name] || lkWhKind_(name)) continue;
      seen[name] = true;
      wh.push(name);
    }
    return wh;
  }
  if (lastCol < 9) return [];
  var wide = [];
  for (h = 8; h < headers.length; h++) {
    name = String(headers[h] || '').trim();
    if (name && !lkWhKind_(name)) wide.push(name);
  }
  return wide;
}

function shCsvCompareWarehouses_(ss, csvWarehouses) {
  var lk = shCsvLkWarehouses_(ss);
  var lkSet = {};
  var csvSet = {};
  var i, name, out, onlyCsv, onlyLk, both;

  for (i = 0; i < lk.length; i++) lkSet[String(lk[i]).trim()] = true;
  for (i = 0; i < csvWarehouses.length; i++) csvSet[String(csvWarehouses[i]).trim()] = true;

  onlyCsv = [];
  onlyLk = [];
  both = 0;
  for (name in csvSet) {
    if (lkSet[name]) both++;
    else onlyCsv.push(name);
  }
  for (name in lkSet) {
    if (!csvSet[name]) onlyLk.push(name);
  }
  onlyCsv.sort(function(a, b) { return a.localeCompare(b, 'ru'); });
  onlyLk.sort(function(a, b) { return a.localeCompare(b, 'ru'); });

  var sheetName = 'Остатки_склады_сравнение';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(sheetName);

  out = [
    ['Только в CSV (' + onlyCsv.length + ')', 'Только в ЛК (' + onlyLk.length + ')', 'Совпало имён', both],
    ['', '', 'Всего CSV', csvWarehouses.length],
    ['', '', 'Всего ЛК', lk.length],
    []
  ];
  var maxLen = Math.max(onlyCsv.length, onlyLk.length, 1);
  for (i = 0; i < maxLen; i++) {
    out.push([
      onlyCsv[i] || '',
      onlyLk[i] || '',
      i === 0 ? 'Примеры CSV:' : '',
      i < csvWarehouses.length ? csvWarehouses[i] : ''
    ]);
  }
  sheet.getRange(1, 1, out.length, 4).setValues(out);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setTabColor('#fff2cc');

  return {
    lkFound: lk.length > 0,
    lkCount: lk.length,
    csvCount: csvWarehouses.length,
    exact: both,
    onlyCsv: onlyCsv.length,
    onlyLk: onlyLk.length
  };
}

// ============================================
// Воронка продаж
// API: POST /api/analytics/v3/sales-funnel/products (nmIds пустой = все).
// Лист «Воронка продаж»: выбранные дни перезаписываются (период с/по).
// Вес — оверлей в конец: из «Товар+Вес»; сумма = Заказы шт × вес.
// ============================================

var SF_SHEET_ = 'Воронка продаж';
var SF_SHEET_OLD_ = 'Аналитика_проба_дни';
var SF_HEADERS_ = [
  'Период с', 'Период по',
  'Артикул WB', 'Артикул продавца', 'Название', 'Бренд', 'Предмет',
  'Переходы', 'В корзину', 'Заказы шт', 'Заказы ₽',
  'Выкупы шт', 'Выкупы ₽', 'Ср. цена',
  'Отмены шт', 'Отмены ₽', 'В избранное', 'Валюта'
];
var SF_OVERLAY_COLS_ = ['Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма заказов'];

function probeSalesFunnel() {
  var tz = Session.getScriptTimeZone();
  var y = new Date();
  y.setDate(y.getDate() - 1);
  var dayStr = Utilities.formatDate(y, tz, 'yyyy-MM-dd');

  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:350px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box}' +
    'button{background:#4CAF50;color:white;padding:10px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e}' +
    '.info{background:#e3f2fd;padding:8px;margin:0 0 10px;border-radius:4px;font-size:11px}' +
    '#st{margin-top:10px;font-size:12px;color:#555}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">📈 Воронка продаж</h3>' +
    '<div class="info">Все товары по дням. Выбранные даты перезаписываются (данные воронки меняются). Вес — из «Товар+Вес».</div>' +
    '<label>Дата ОТ:</label><input type="date" id="dateFrom" value="' + dayStr + '">' +
    '<label>Дата ДО:</label><input type="date" id="dateTo" value="' + dayStr + '">' +
    '<button id="btn" onclick="go()">📥 Загрузить</button>' +
    '<div id="st"></div>' +
    '<script>' +
    'function go(){' +
    'var d1=document.getElementById("dateFrom").value;' +
    'var d2=document.getElementById("dateTo").value;' +
    'if(!d1||!d2){alert("Выберите даты");return;}' +
    'document.getElementById("btn").disabled=true;' +
    'document.getElementById("st").textContent="Загрузка… не закрывайте окно";' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){document.getElementById("st").textContent=msg||"Готово";document.getElementById("btn").disabled=false;alert(msg||"Готово");google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("btn").disabled=false;document.getElementById("st").textContent="";alert(String(e));})' +
    '.probeSalesFunnelRun(d1,d2);' +
    '}' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(360), 'Воронка продаж');
}

function probeSalesFunnelRun(dateFrom, dateTo) {
  var token = getApiToken();
  if (!token) throw new Error('Нет API токена. Сначала меню → Ввести API токен.');
  return loadSalesFunnelRange(token, dateFrom, dateTo);
}

function loadSalesFunnelRange(token, dateFrom, dateTo) {
  var d1 = String(dateFrom || '').substring(0, 10);
  var d2 = String(dateTo || '').substring(0, 10);
  if (!d1 || !d2) throw new Error('Не выбраны даты');
  if (d1 > d2) {
    var tmp = d1;
    d1 = d2;
    d2 = tmp;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var added = 0;
  var days = 0;
  var day = new Date(d1 + 'T00:00:00');
  var end = new Date(d2 + 'T00:00:00');
  while (day.getTime() <= end.getTime()) {
    if (days > 0) Utilities.sleep(21000);
    var dayStr = Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var fetched = sfFetchProducts_(token, dayStr, dayStr);
    added += sfAppendDailyRows_(ss, fetched.products, dayStr, dayStr, fetched.currency);
    days++;
    day.setDate(day.getDate() + 1);
  }
  var msg = '✅ ' + d1 + ' — ' + d2 + '\nДней: ' + days + '\nЗаписано строк: ' + added + '\nЛист: ' + SF_SHEET_;
  Logger.log(msg);
  return msg;
}

function sfFetchProducts_(token, d1, d2) {
  var url = 'https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products';
  var products = [];
  var currency = '';
  var offset = 0;
  var limit = 1000;
  var page;
  for (page = 0; page < 20; page++) {
    if (page > 0) Utilities.sleep(21000);
    var body = {
      selectedPeriod: { start: d1, end: d2 },
      nmIds: [],
      skipDeletedNm: true,
      limit: limit,
      offset: offset
    };
    var res = wrFetchJson_(token, url, 'POST', body);
    Logger.log('воронка ' + d1 + ' page ' + (page + 1) + ' HTTP ' + res.code);
    if (res.code === 429) throw new Error('WB лимит. Подождите минуту и повторите.');
    if (res.code === 402) throw new Error('HTTP 402 — возможно нужен Джем или другой токен.\n' + String(res.text || '').substring(0, 300));
    if (res.code === 403) throw new Error('HTTP 403 — нужен токен с категорией «Аналитика».\n' + String(res.text || '').substring(0, 300));
    if (res.code !== 200) throw new Error('HTTP ' + res.code + '\n' + String(res.text || '').substring(0, 400));
    var data = res.data && res.data.data ? res.data.data : res.data;
    var chunk = (data && data.products) ? data.products : [];
    if (data && data.currency) currency = data.currency;
    if (!chunk.length) break;
    products = products.concat(chunk);
    if (chunk.length < limit) break;
    offset += limit;
  }
  return { products: products, currency: currency };
}

function sfGetSheet_(ss) {
  var sheet = ss.getSheetByName(SF_SHEET_);
  if (sheet) return sheet;
  var old = ss.getSheetByName(SF_SHEET_OLD_);
  if (old) {
    old.setName(SF_SHEET_);
    return old;
  }
  sheet = ss.insertSheet(SF_SHEET_);
  sheet.getRange(1, 1, 1, SF_HEADERS_.length).setValues([SF_HEADERS_]);
  sheet.getRange(1, 1, 1, SF_HEADERS_.length).setFontWeight('bold').setBackground('#f1f1f1');
  sheet.setFrozenRows(1);
  return sheet;
}

function sfIso_(val) {
  if (val == null || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var m = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  return s.substring(0, 10);
}

function sfDate_(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || '';
  return new Date(iso + 'T00:00:00');
}

function sfWeightForArt_(weights, art) {
  if (!art) return 0;
  var key = String(art).trim();
  if (weights[key] != null) return weights[key];
  var low = key.toLowerCase();
  return weights[low] != null ? weights[low] : 0;
}

function sfEnsureOverlayCols_(sheet) {
  if (typeof wbClearFilter_ === 'function') wbClearFilter_(sheet);
  var last = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  var i;
  for (i = 0; i < SF_OVERLAY_COLS_.length; i++) {
    if (headers.indexOf(SF_OVERLAY_COLS_[i]) !== -1) continue;
    sheet.getRange(1, last + 1).setValue(SF_OVERLAY_COLS_[i]).setFontWeight('bold').setBackground('#f1f1f1');
    last++;
    headers.push(SF_OVERLAY_COLS_[i]);
  }
  return headers;
}

function sfAppendDailyRows_(ss, products, selStart, selEnd, currency) {
  var sheet = sfGetSheet_(ss);
  var headers = sfEnsureOverlayCols_(sheet);
  var lastRow = sheet.getLastRow();
  var startIdx = headers.indexOf('Период с');
  var endIdx = headers.indexOf('Период по');
  var keep = [];
  var r;
  if (lastRow > 1 && startIdx >= 0) {
    var old = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (r = 0; r < old.length; r++) {
      if (sfIso_(old[r][startIdx]) === selStart && sfIso_(old[r][endIdx]) === selEnd) continue;
      keep.push(old[r]);
    }
    if (keep.length !== old.length) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      if (keep.length) {
        sheet.getRange(2, 1, keep.length, headers.length).setValues(keep);
      }
    }
  }

  var weights = {};
  try { weights = loadWeightsMap() || {}; } catch (eW) {}
  var weightsLow = {};
  var wk;
  for (wk in weights) {
    if (Object.prototype.hasOwnProperty.call(weights, wk)) {
      weightsLow[String(wk).toLowerCase()] = weights[wk];
    }
  }

  var rows = [];
  var i, item, prod, sel, nm, art, orderCount, weight, obj, line, h;
  var seen = {};
  for (i = 0; i < products.length; i++) {
    item = products[i] || {};
    prod = item.product || {};
    sel = (item.statistic && item.statistic.selected) ? item.statistic.selected : {};
    nm = prod.nmId != null ? String(prod.nmId).trim() : '';
    var uniq = selStart + '|' + selEnd + '|' + nm;
    if (seen[uniq]) continue;
    seen[uniq] = true;
    art = prod.vendorCode || '';
    orderCount = sel.orderCount != null ? Number(sel.orderCount) || 0 : 0;
    weight = sfWeightForArt_(weights, art);
    if (!weight) weight = sfWeightForArt_(weightsLow, art);
    obj = {};
    obj['Период с'] = sfDate_(selStart);
    obj['Период по'] = sfDate_(selEnd);
    obj['Артикул WB'] = nm;
    obj['Артикул продавца'] = art;
    obj['Название'] = prod.title || '';
    obj['Бренд'] = prod.brandName || '';
    obj['Предмет'] = prod.subjectName || '';
    obj['Переходы'] = sel.openCount != null ? sel.openCount : '';
    obj['В корзину'] = sel.cartCount != null ? sel.cartCount : '';
    obj['Заказы шт'] = sel.orderCount != null ? sel.orderCount : '';
    obj['Заказы ₽'] = sel.orderSum != null ? sel.orderSum : '';
    obj['Выкупы шт'] = sel.buyoutCount != null ? sel.buyoutCount : '';
    obj['Выкупы ₽'] = sel.buyoutSum != null ? sel.buyoutSum : '';
    obj['Ср. цена'] = sel.avgPrice != null ? sel.avgPrice : '';
    obj['Отмены шт'] = sel.cancelCount != null ? sel.cancelCount : '';
    obj['Отмены ₽'] = sel.cancelSum != null ? sel.cancelSum : '';
    obj['В избранное'] = sel.addToWishlist != null ? sel.addToWishlist : '';
    obj['Валюта'] = currency || '';
    obj['Вес, кг за 1 шт'] = weight;
    obj['Вес, кг за 1 шт сумма заказов'] = weight * orderCount;
    line = [];
    for (h = 0; h < headers.length; h++) {
      line.push(obj[headers[h]] != null ? obj[headers[h]] : '');
    }
    rows.push(line);
  }

  if (!rows.length) return 0;
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  if (startIdx >= 0) {
    try { sheet.getRange(2, startIdx + 1, sheet.getLastRow() - 1, 1).setNumberFormat('dd.mm.yyyy'); } catch (e1) {}
  }
  if (endIdx >= 0) {
    try { sheet.getRange(2, endIdx + 1, sheet.getLastRow() - 1, 1).setNumberFormat('dd.mm.yyyy'); } catch (e2) {}
  }
  var wIdx = headers.indexOf('Вес, кг за 1 шт');
  var wsIdx = headers.indexOf('Вес, кг за 1 шт сумма заказов');
  if (wIdx >= 0) {
    try { sheet.getRange(2, wIdx + 1, sheet.getLastRow() - 1, 1).setNumberFormat('0.000'); } catch (e3) {}
  }
  if (wsIdx >= 0) {
    try { sheet.getRange(2, wsIdx + 1, sheet.getLastRow() - 1, 1).setNumberFormat('0.000'); } catch (e4) {}
  }
  wbPutSheetFilter_(sheet);
  return rows.length;
}

// ============================================
// ПРОБА: история изменений цен (history/tasks + history/goods/task)
// Нужен uploadID — отдельного «списка всех загрузок» в API нет.
// ============================================

function probePriceHistory() {
  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:15px;width:380px;margin:0}' +
    'label{display:block;margin:8px 0 4px;font-weight:bold}' +
    'input,textarea{width:100%;padding:6px;margin-bottom:10px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box}' +
    'button{background:#4CAF50;color:white;padding:10px 16px;border:none;border-radius:4px;cursor:pointer;width:100%;font-size:14px}' +
    'button:disabled{background:#9e9e9e}' +
    '.info{background:#fff3e0;padding:8px;margin:0 0 10px;border-radius:4px;font-size:11px;line-height:1.4}' +
    '#st{margin-top:10px;font-size:12px;color:#555}' +
    '</style></head><body>' +
    '<h3 style="margin-top:0">💰 История цен (проба)</h3>' +
    '<div class="info">Это не архив цен по дням, а детали конкретной загрузки цен/скидок.<br>' +
    'Нужен uploadID (из ответа upload/task или из кабинета).<br>' +
    'Можно несколько через запятую.</div>' +
    '<label>uploadID:</label>' +
    '<textarea id="ids" rows="3" placeholder="395643565 или 111,222,333"></textarea>' +
    '<button id="btn" onclick="go()">📥 Загрузить</button>' +
    '<div id="st"></div>' +
    '<script>' +
    'function go(){' +
    'var v=document.getElementById("ids").value.trim();' +
    'if(!v){alert("Введите uploadID");return;}' +
    'document.getElementById("btn").disabled=true;' +
    'document.getElementById("st").textContent="Загрузка…";' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){document.getElementById("st").textContent=msg||"Готово";document.getElementById("btn").disabled=false;alert(msg||"Готово");google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("btn").disabled=false;document.getElementById("st").textContent="";alert(String(e));})' +
    '.probePriceHistoryRun(v);' +
    '}' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(360), 'История цен (проба)');
}

/** Без _ в конце: иначе HtmlService не вызывает функцию. */
function probePriceHistoryRun(uploadIdsText) {
  var token = getApiToken();
  if (!token) throw new Error('Нет API токена. Сначала меню → Ввести API токен.');

  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('Книга занята. Закройте лишние вкладки и повторите.\n' + e);
  }

  var ids = phParseUploadIds_(uploadIdsText);
  if (!ids.length) throw new Error('Не указан uploadID');

  var host = 'https://discounts-prices-api.wildberries.ru';
  var allRows = [];
  var metaRows = [];
  var errors = [];
  var i, uploadId, taskRes, taskData, goods, offset, limit, page, goodsRes, goodsData, g, item;

  for (i = 0; i < ids.length; i++) {
    uploadId = ids[i];
    taskRes = wrFetchJson_(token, host + '/api/v2/history/tasks?uploadID=' + encodeURIComponent(uploadId), 'GET');
    Logger.log('history/tasks ' + uploadId + ' HTTP ' + taskRes.code + ' ' + String(taskRes.text || '').substring(0, 200));
    if (taskRes.code === 429) throw new Error('WB лимит. Подождите минуту и повторите.');
    if (taskRes.code === 403) throw new Error('HTTP 403 — нужен токен с категорией «Цены и скидки».\n' + String(taskRes.text || '').substring(0, 300));
    if (taskRes.code !== 200) {
      errors.push(uploadId + ': tasks HTTP ' + taskRes.code + ' — ' + String(taskRes.text || '').substring(0, 120));
      continue;
    }

    taskData = (taskRes.data && taskRes.data.data) ? taskRes.data.data : (taskRes.data || {});
    metaRows.push([
      uploadId,
      taskData.status != null ? taskData.status : '',
      taskData.uploadDate || '',
      taskData.activationDate || '',
      taskData.overAllGoodsNumber != null ? taskData.overAllGoodsNumber : '',
      taskData.successGoodsNumber != null ? taskData.successGoodsNumber : '',
      phStatusText_(taskData.status)
    ]);

    goods = [];
    offset = 0;
    limit = 1000;
    for (page = 0; page < 50; page++) {
      goodsRes = wrFetchJson_(token,
        host + '/api/v2/history/goods/task?uploadID=' + encodeURIComponent(uploadId) +
        '&limit=' + limit + '&offset=' + offset, 'GET');
      Logger.log('history/goods ' + uploadId + ' off=' + offset + ' HTTP ' + goodsRes.code);
      if (goodsRes.code === 429) throw new Error('WB лимит. Подождите минуту и повторите.');
      if (goodsRes.code !== 200) {
        errors.push(uploadId + ': goods HTTP ' + goodsRes.code + ' — ' + String(goodsRes.text || '').substring(0, 120));
        break;
      }
      goodsData = (goodsRes.data && goodsRes.data.data) ? goodsRes.data.data : (goodsRes.data || {});
      var chunk = goodsData.historyGoods || goodsData.goods || [];
      if (!chunk.length) break;
      goods = goods.concat(chunk);
      if (chunk.length < limit) break;
      offset += limit;
      Utilities.sleep(700);
    }

    for (g = 0; g < goods.length; g++) {
      item = goods[g] || {};
      var price = item.price != null ? item.price : '';
      var disc = item.discount != null ? item.discount : '';
      var priceWithDisc = '';
      if (price !== '' && disc !== '') {
        priceWithDisc = Math.round(Number(price) * (1 - Number(disc) / 100));
      }
      allRows.push([
        uploadId,
        taskData.uploadDate || '',
        taskData.activationDate || '',
        taskData.status != null ? taskData.status : '',
        item.nmID || item.nmId || '',
        item.vendorCode || '',
        item.sizeID || item.sizeId || '',
        item.techSizeName || '',
        price,
        disc,
        item.clubDiscount != null ? item.clubDiscount : '',
        priceWithDisc,
        item.currencyIsoCode4217 || 'RUB',
        item.status != null ? item.status : '',
        item.errorText || ''
      ]);
    }

    if (!goods.length && taskRes.code === 200) {
      allRows.push([
        uploadId, taskData.uploadDate || '', taskData.activationDate || '',
        taskData.status != null ? taskData.status : '',
        '', '', '', '', '', '', '', '', '', '', 'нет товаров в ответе goods/task'
      ]);
    }

    if (i < ids.length - 1) Utilities.sleep(700);
  }

  var counts;
  try {
    counts = phWritePriceHistorySheets_(ss, metaRows, allRows);
  } catch (e) {
    throw new Error('API ок, запись в книгу не прошла.\nПодождите минуту.\n' + e);
  }

  var msg = '✅ Загрузок: ' + counts.uploads + ', строк товаров: ' + counts.goods +
    '\nЛисты: История_цен_проба, История_цен_проба_загрузки';
  if (errors.length) msg += '\n\nОшибки:\n' + errors.join('\n');
  return msg;
}

function phParseUploadIds_(text) {
  var raw = String(text || '').split(/[\s,;]+/);
  var out = [];
  var i, n;
  for (i = 0; i < raw.length; i++) {
    n = String(raw[i] || '').trim();
    if (n && /^\d+$/.test(n)) out.push(n);
  }
  return out;
}

function phStatusText_(status) {
  var map = {
    3: 'обработана, без ошибок',
    4: 'отменена',
    5: 'обработана, есть ошибки',
    6: 'обработана, все с ошибками'
  };
  return map[status] || '';
}

function phWritePriceHistorySheets_(ss, metaRows, goodsRows) {
  var metaSheetName = 'История_цен_проба_загрузки';
  var goodsSheetName = 'История_цен_проба';
  var metaSheet = ss.getSheetByName(metaSheetName);
  if (!metaSheet) metaSheet = ss.insertSheet(metaSheetName);
  var goodsSheet = ss.getSheetByName(goodsSheetName);
  if (!goodsSheet) goodsSheet = ss.insertSheet(goodsSheetName);

  var metaHeaders = [
    'uploadID', 'Статус код', 'Дата загрузки', 'Дата активации',
    'Всего товаров', 'Без ошибок', 'Статус'
  ];
  var goodsHeaders = [
    'uploadID', 'Дата загрузки', 'Дата активации', 'Статус загрузки',
    'Артикул WB', 'Артикул продавца', 'sizeID', 'Размер',
    'Цена', 'Скидка %', 'Скидка клуб %', 'Цена со скидкой (расч.)',
    'Валюта', 'Статус товара', 'Ошибка'
  ];

  var metaOut = [metaHeaders];
  var goodsOut = [goodsHeaders];
  var i;
  for (i = 0; i < metaRows.length; i++) metaOut.push(metaRows[i]);
  for (i = 0; i < goodsRows.length; i++) goodsOut.push(goodsRows[i]);
  if (metaOut.length === 1) metaOut.push(['', '', 'нет данных', '', '', '', '']);
  if (goodsOut.length === 1) goodsOut.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', 'нет данных']);

  metaSheet.getRange(1, 1, metaOut.length, metaHeaders.length).setValues(metaOut);
  metaSheet.getRange(1, 1, 1, metaHeaders.length).setFontWeight('bold').setBackground('#f1f1f1');
  goodsSheet.getRange(1, 1, goodsOut.length, goodsHeaders.length).setValues(goodsOut);
  goodsSheet.getRange(1, 1, 1, goodsHeaders.length).setFontWeight('bold').setBackground('#f1f1f1');

  try {
    var oldMeta = metaSheet.getLastRow();
    if (oldMeta > metaOut.length) {
      metaSheet.getRange(metaOut.length + 1, 1, oldMeta - metaOut.length, metaHeaders.length).clearContent();
    }
    var oldGoods = goodsSheet.getLastRow();
    if (oldGoods > goodsOut.length) {
      goodsSheet.getRange(goodsOut.length + 1, 1, oldGoods - goodsOut.length, goodsHeaders.length).clearContent();
    }
  } catch (eClear) {
    Logger.log('хвост листа истории цен: ' + eClear);
  }

  return { uploads: Math.max(0, metaOut.length - 1), goods: Math.max(0, goodsOut.length - 1) };
}








// ============================================
// АВТООБНОВЛЕНИЕ - ЧАСТЬ 1 (5:00)
// Финансы + Реклама + Склады + Тарифы
// ============================================

function updateAllReports1() {
  var t0 = Date.now();
  var lines = [];
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateStr = Utilities.formatDate(yesterday, 'GMT+3', 'yyyy-MM-dd');
  lines.push('дата: ' + dateStr);

  var token;
  try { token = getApiToken(); } catch (eTok) { token = ''; }
  if (typeof morningTgPendingList_ === 'function') {
    morningTgPendingList_(['Финансы', 'Реклама', 'Воронка', 'Заказы']);
  }
  if (!token) {
    lines.push('❌ Токен WB не найден');
    if (typeof morningTgFailList_ === 'function') {
      morningTgFailList_(['Финансы', 'Реклама', 'Воронка', 'Заказы'], 'нет токена');
    }
    tgSendJob_('updateAllReports1', false, lines, t0);
    return;
  }

  Logger.log('🚀 [ЧАСТЬ 1] Начало автообновления за ' + dateStr);

  try {
    сброситьФильтрыНаВсехЛистах();
    if (typeof renameWbHeadersAllBooks_ === 'function') {
      try { renameWbHeadersAllBooks_(); } catch (eRen) { lines.push('rename: ' + eRen); }
    }

    tgRunStep_(lines, 'Финансы', function() {
      tgWatch_(lines, 'Финансовые_ежедневные', function() {
        loadFinancialReportWithPeriod(token, dateStr, dateStr, 'daily');
      }, 'Финансы');
    });
    Utilities.sleep(3000);

    tgRunStep_(lines, 'Реклама', function() {
      tgWatch_(lines, 'Рекламные_расходы', function() {
        loadAdvExpenses(token, dateStr, dateStr);
      }, 'Реклама');
    });
    Utilities.sleep(3000);

    tgRunStep_(lines, 'Воронка', function() {
      tgWatch_(lines, 'Воронка продаж', function() {
        loadSalesFunnelRange(token, dateStr, dateStr);
      }, 'Воронка');
    });
    Utilities.sleep(3000);

    tgRunStep_(lines, 'Заказы', function() {
      tgWatch_(lines, 'Заказы', function() {
        loadOrdersForYesterdayAuto();
      });
    });

    Logger.log('✅ [ЧАСТЬ 1] Завершена');
    поставитьФильтрыНаЛисты_(lines);
    tgSendJob_('updateAllReports1', true, lines, t0);
  } catch (error) {
    lines.push('❌ ' + error);
    Logger.log('❌ Ошибка в Part1: ' + error.toString());
    try { поставитьФильтрыНаЛисты_(lines); } catch (eF) {}
    tgSendJob_('updateAllReports1', false, lines, t0);
  }
}

// ============================================
// АВТООБНОВЛЕНИЕ - ЧАСТЬ 2 (6:00)
// Товары + Остатки + Заказы + Поставки
// ============================================

function updateAllReports2() {
  var t0 = Date.now();
  var lines = [];
  var token;
  try { token = getApiToken(); } catch (eTok) { token = ''; }
  if (typeof morningTgPendingList_ === 'function') {
    morningTgPendingList_(['Товары', 'История цен', 'Остатки', 'Остатки ЛК', 'В пути', 'Поставки']);
  }
  if (!token) {
    lines.push('❌ Токен WB не найден');
    if (typeof morningTgFailList_ === 'function') {
      morningTgFailList_(['Товары', 'История цен', 'Остатки', 'Остатки ЛК', 'В пути', 'Поставки'], 'нет токена');
    }
    tgSendJob_('updateAllReports2', false, lines, t0);
    return;
  }

  Logger.log('🚀 [ЧАСТЬ 2] Начало автообновления');

  try {
    сброситьФильтрыНаВсехЛистах();
    if (typeof renameWbHeadersAllBooks_ === 'function') {
      try { renameWbHeadersAllBooks_(); } catch (eRen) { lines.push('rename: ' + eRen); }
    }

    tgRunStep_(lines, 'Товары', function() {
      tgWatch_(lines, 'Товары', function() {
        loadAllProducts();
      });
    });
    Utilities.sleep(2000);

    tgRunStep_(lines, 'История цен', function() {
      tgWatch_(lines, 'История цен', function() {
        savePriceHistory();
      });
    });
    Utilities.sleep(2000);

    tgRunStep_(lines, 'Остатки', function() {
      tgWatch_(lines, 'Остатки', function() {
        recordCurrentStocksEnriched(true);
      });
    });
    Utilities.sleep(3000);

    tgRunStep_(lines, 'Остатки ЛК', function() {
      tgWatch_(lines, 'Остатки ЛК', function() {
        recordLkStocks();
      });
    });
    Utilities.sleep(2000);

    tgRunStep_(lines, 'В пути', function() {
      tgWatch_(lines, 'В пути', function() {
        saveStocksHistory();
      });
    });
    Utilities.sleep(3000);

    tgRunStep_(lines, 'Поставки', function() {
      tgWatch_(lines, 'Поставки_список', function() {
        updateSuppliesDaily();
      }, 'Поставки');
    });

    scheduleUpdateSuppliesStatuses_();
    lines.push('статусы поставок: запланированы отдельно');

    Logger.log('✅ [ЧАСТЬ 2] Завершена (статусы запланированы отдельно)');
    поставитьФильтрыНаЛисты_(lines);
    tgSendJob_('updateAllReports2', true, lines, t0);
  } catch (error) {
    lines.push('❌ ' + error);
    Logger.log('❌ Ошибка в Part2: ' + error.toString());
    try { поставитьФильтрыНаЛисты_(lines); } catch (eF) {}
    tgSendJob_('updateAllReports2', false, lines, t0);
  }
}

/** Удаляет одноразовые триггеры для handler и ставит новый через afterMs */
function scheduleUpdateSuppliesStatuses_() {
  if (typeof morningTgPending_ === 'function') morningTgPending_('Статусы FBW');
  clearTriggersForHandler_('runUpdateSuppliesStatusesJob_');
  ScriptApp.newTrigger('runUpdateSuppliesStatusesJob_')
    .timeBased()
    .after(20 * 1000)
    .create();
  Logger.log('⏰ Статусы поставок: запуск через ~20 сек');
}

function runUpdateSuppliesStatusesJob_() {
  clearTriggersForHandler_('runUpdateSuppliesStatusesJob_');
  if (!morningChainSheetLockOrReschedule_('runUpdateSuppliesStatusesJob_', 45 * 1000, function() {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('SUPPLIES_STATUS_NEXT_SID');
    props.deleteProperty('SUPPLIES_STATUS_IDX');
    updateExistingSuppliesStatuses();
  })) return;
}

function clearTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === handlerName) {
      try { ScriptApp.deleteTrigger(t); } catch (e) {}
    }
  });
}

function updateExistingSuppliesStatuses() {
  var token = getApiToken();
  if (!token) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName('Поставки_список');
  if (!listSheet || listSheet.getLastRow() < 2) {
    if (!isRunningFromTrigger()) showAlert('⚠️ Лист "Поставки_список" пуст или не найден');
    return;
  }

  var data = listSheet.getDataRange().getValues();
  var headers = data[0];
  var statusIdCol = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['statusID']) : headers.indexOf('statusID');
  var supplyIdCol = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, ['supplyID']) : headers.indexOf('supplyID');

  if (statusIdCol === -1 || supplyIdCol === -1) {
    if (!isRunningFromTrigger()) showAlert('❌ В листе не найдены колонки supplyID или statusID');
    return;
  }

  var statusColNum = statusIdCol + 1;
  var updatedCount = 0;
  var deletedCount = 0;
  var rowsToDelete = [];
  var statusUpdates = []; // {row, value} — пишем пачкой в конце чанка
  var toCheck = [];
  var skippedAccepted = 0;

  // Все поставки с supplyID, кроме «Принято» (5) — их не дергаем
  for (var i = 1; i < data.length; i++) {
    var sid = data[i][supplyIdCol];
    var curStatus = data[i][statusIdCol];
    if (!sid) continue;
    if (parseInt(curStatus, 10) === 5) {
      skippedAccepted++;
      continue;
    }
    toCheck.push({ row: i + 1, sid: sid, curStatus: curStatus });
  }

  if (toCheck.length === 0) {
    Logger.log('ℹ️ Нет поставок для проверки статусов (пропущено Принято: ' + skippedAccepted + ')');
    if (typeof morningChainOn_ === 'function' && morningChainOn_()) {
      rebuildUnifiedSheet();
      sortAllSheetsBySupplyDate();
      wbPutNamedFilters_(['Поставки_список', 'Детали_поставок', 'Товары_поставок', 'Поставки']);
    }
    if (!isRunningFromTrigger()) showAlert('ℹ️ Нет поставок для проверки');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var resumeSid = props.getProperty('SUPPLIES_STATUS_NEXT_SID');
  var startIdx = 0;
  if (resumeSid) {
    for (var ri = 0; ri < toCheck.length; ri++) {
      if (String(toCheck[ri].sid) === String(resumeSid)) {
        startIdx = ri;
        break;
      }
    }
  }

  Logger.log('🔄 Проверяем статусы: ' + toCheck.length + ' поставок (старт с #' + (startIdx + 1) + '), пропущено Принято: ' + skippedAccepted);

  var startMs = Date.now();
  var MAX_MS = 4.5 * 60 * 1000; // оставляем запас до лимита 6 мин
  var REQUEST_DELAY_MS = 800;   // чуть медленнее → меньше 429
  var finished = true;
  var stopAt = toCheck.length;

  for (var j = startIdx; j < toCheck.length; j++) {
    if (Date.now() - startMs > MAX_MS) {
      finished = false;
      stopAt = j;
      break;
    }

    var item = toCheck[j];
    Utilities.sleep(REQUEST_DELAY_MS);

    var url = 'https://supplies-api.wildberries.ru/api/v1/supplies/' + item.sid;
    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: { 'Authorization': token },
        muteHttpExceptions: true
      });

      var statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        var details = JSON.parse(response.getContentText());
        var newStatus = details.statusID;
        if (newStatus !== undefined && newStatus != item.curStatus) {
          statusUpdates.push({ row: item.row, value: newStatus });
          updatedCount++;
          Logger.log('  ✅ ' + item.sid + ': ' + getStatusName(item.curStatus) + ' → ' + getStatusName(newStatus));
        }
      } else if (statusCode === 400) {
        var errorText = response.getContentText();
        if (errorText.indexOf('поставка удалена') !== -1) {
          rowsToDelete.push(item.row);
          deletedCount++;
          Logger.log('  🗑️ ' + item.sid + ': ПОСТАВКА УДАЛЕНА');
        } else {
          Logger.log('  ⚠️ ' + item.sid + ': ошибка 400 - ' + errorText);
        }
      } else if (statusCode === 404) {
        rowsToDelete.push(item.row);
        deletedCount++;
        Logger.log('  🗑️ ' + item.sid + ': НЕ НАЙДЕНА (404)');
      } else if (statusCode === 429) {
        Logger.log('⏳ Лимит API (429). Ждём 5 сек...');
        Utilities.sleep(5000);
        j--;
        continue;
      } else {
        Logger.log('  ⚠️ ' + item.sid + ': ошибка ' + statusCode);
      }
    } catch (e) {
      Logger.log('  ❌ Ошибка для ' + item.sid + ': ' + e.toString());
    }

    if ((j + 1) % 25 === 0) Logger.log('📊 Прогресс статусов: ' + (j + 1) + '/' + toCheck.length);
  }

  // Пачечная запись статусов
  for (var u = 0; u < statusUpdates.length; u++) {
    listSheet.getRange(statusUpdates[u].row, statusColNum).setValue(statusUpdates[u].value);
  }

  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var k = 0; k < rowsToDelete.length; k++) {
      listSheet.deleteRow(rowsToDelete[k]);
    }
    Logger.log('🗑️ Удалено строк: ' + rowsToDelete.length);
  }

  if (!finished) {
    props.setProperty('SUPPLIES_STATUS_NEXT_SID', String(toCheck[stopAt].sid));
    clearTriggersForHandler_('continueUpdateSuppliesStatusesJob_');
    clearTriggersForHandler_('runUpdateSuppliesStatusesJob_');
    ScriptApp.newTrigger('continueUpdateSuppliesStatusesJob_')
      .timeBased()
      .after(30 * 1000)
      .create();
    Logger.log('⏸️ Статусы: пауза по времени, продолжим с #' + (stopAt + 1) + '/' + toCheck.length + ' (sid ' + toCheck[stopAt].sid + ') через 30 сек');
    return;
  }

  props.deleteProperty('SUPPLIES_STATUS_NEXT_SID');
  props.deleteProperty('SUPPLIES_STATUS_IDX');
  rebuildUnifiedSheet();
  wbPutNamedFilters_(['Поставки_список', 'Детали_поставок', 'Товары_поставок', 'Поставки']);

  var msg = '✅ Обновление статусов завершено! Изменено: ' + updatedCount + ', удалено: ' + deletedCount;
  Logger.log(msg);
  if (!isRunningFromTrigger()) showAlert(msg);
}

/** Продолжение проверки статусов без сброса индекса */
function continueUpdateSuppliesStatusesJob_() {
  clearTriggersForHandler_('continueUpdateSuppliesStatusesJob_');
  if (!morningChainSheetLockOrReschedule_('continueUpdateSuppliesStatusesJob_', 45 * 1000, function() {
    updateExistingSuppliesStatuses();
  })) return;
}

// ============================================
// АВТООБНОВЛЕНИЕ - ЧАСТЬ 3 (6:00)
// Товары + Остатки + Заказы + Поставки
// ============================================

function updateAllReports3() {
  var t0 = Date.now();
  var lines = [];
  var token;
  try { token = getApiToken(); } catch (eTok) { token = ''; }
  if (typeof morningTgPending_ === 'function') morningTgPending_('Склады');
  if (!token) {
    lines.push('❌ Токен WB не найден');
    if (typeof morningTgNote_ === 'function') morningTgNote_(false, 'Склады', 'нет токена');
    tgSendJob_('updateAllReports3', false, lines, t0);
    return;
  }

  Logger.log('🚀 [ЧАСТЬ 3] Начало автообновления');

  try {
    tgRunStep_(lines, 'Склады', function() {
      var stats = fixWarehouseNamesLastDate2() || { scanned: 0, replaced: 0 };
      lines.push('Проверено: ' + stats.scanned);
      lines.push('Заменено: ' + stats.replaced);
      if (typeof morningTgNote_ === 'function') {
        morningTgNote_(true, 'Склады', stats.scanned + ' | ' + stats.replaced);
      }
    });
    Logger.log('✅ [ЧАСТЬ 3] Завершена');
    tgSendJob_('updateAllReports3', true, lines, t0);
  } catch (error) {
    lines.push('❌ ' + error);
    Logger.log('❌ Ошибка в Part3: ' + error.toString());
    tgSendJob_('updateAllReports3', false, lines, t0);
  }
}

// ============================================
// УТРЕННЯЯ ЦЕПОЧКА (отдельно от старых триггеров)
// ч1 → ч2 → (ждём статусы поставок) → FBS → ч3
// Старые updateAllReports1/2/3 и updateFbsDaily не меняет.
// В триггерах Apps Script оставьте только runMorningChain;
// четыре старых ежедневных — выключите, иначе будет двойной прогон.
// ============================================

var MORNING_TG_STEPS_ = [
  'Финансы', 'Реклама', 'Воронка', 'Заказы',
  'Товары', 'История цен', 'Остатки', 'Остатки ЛК', 'В пути', 'Поставки',
  'Статусы FBW',
  'FBS заказы', 'FBS поставки', 'FBS остатки',
  'Сборка FBS',
  'Склады'
];

var MORNING_TG_NAME_ = {
  'Финансы': 'Финансы',
  'Финансовые_ежедневные': 'Финансы',
  'Реклама': 'Реклама',
  'Рекламные_расходы': 'Реклама',
  'Воронка': 'Воронка',
  'Воронка продаж': 'Воронка',
  'Аналитика_проба_дни': 'Воронка',
  'Заказы': 'Заказы',
  'Товары': 'Товары',
  'История цен': 'История цен',
  'Остатки': 'Остатки',
  'Остатки ЛК': 'Остатки ЛК',
  'В пути': 'В пути',
  'Поставки': 'Поставки',
  'Поставки_список': 'Поставки',
  'Статусы FBW': 'Статусы FBW',
  'FBS_Заказы': 'FBS заказы',
  'FBS заказы': 'FBS заказы',
  'FBS_Поставки': 'FBS поставки',
  'FBS поставки': 'FBS поставки',
  'FBS_Остатки': 'FBS остатки',
  'FBS остатки': 'FBS остатки',
  'Сборка FBS': 'Сборка FBS',
  'Склады': 'Склады'
};

function morningTgCanon_(name) {
  return MORNING_TG_NAME_[name] || name;
}

function morningTgInit_() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var data = {
    startedMs: Date.now(),
    date: Utilities.formatDate(now, tz, 'dd.MM.yyyy'),
    start: Utilities.formatDate(now, tz, 'HH:mm'),
    items: {}
  };
  PropertiesService.getScriptProperties().setProperty('WB_MORNING_TG', JSON.stringify(data));
}

function morningTgLoad_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('WB_MORNING_TG');
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function morningTgSave_(data) {
  PropertiesService.getScriptProperties().setProperty('WB_MORNING_TG', JSON.stringify(data));
}

function morningTgSet_(name, rec) {
  var data = morningTgLoad_();
  if (!data) return;
  if (!data.items) data.items = {};
  data.items[name] = rec;
  morningTgSave_(data);
}

function morningTgGet_(name) {
  var data = morningTgLoad_();
  if (!data || !data.items) return null;
  return data.items[morningTgCanon_(name)] || null;
}

function morningTgPending_(name) {
  if (typeof morningChainOn_ !== 'function' || !morningChainOn_()) return;
  var n = morningTgCanon_(name);
  if (MORNING_TG_STEPS_.indexOf(n) < 0) return;
  morningTgSet_(n, { ok: false, pending: true, detail: '' });
}

function morningTgPendingList_(names) {
  var i;
  for (i = 0; i < names.length; i++) morningTgPending_(names[i]);
}

function morningTgNote_(ok, name, detail) {
  if (typeof morningChainOn_ !== 'function' || !morningChainOn_()) return;
  var n = morningTgCanon_(name);
  if (MORNING_TG_STEPS_.indexOf(n) < 0) return;
  morningTgSet_(n, { ok: !!ok, pending: false, detail: detail ? String(detail) : '' });
}

function morningTgFailList_(names, detail) {
  var i;
  for (i = 0; i < names.length; i++) morningTgNote_(false, names[i], detail);
}

function morningTgFlush_() {
  var data = morningTgLoad_();
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_TG');
  if (!data) return;
  var tz = Session.getScriptTimeZone();
  var end = Utilities.formatDate(new Date(), tz, 'HH:mm');
  var min = data.startedMs ? Math.round((Date.now() - data.startedMs) / 60000) : 0;
  var out = [
    'утро · ' + (data.date || ''),
    (data.start || '') + ' → ' + end + ' · ' + min + ' мин',
    ''
  ];
  var items = data.items || {};
  var i, name, rec, line;
  for (i = 0; i < MORNING_TG_STEPS_.length; i++) {
    name = MORNING_TG_STEPS_[i];
    rec = items[name];
    if (!rec) {
      out.push('❌ ' + name + ' · не запускалась');
      continue;
    }
    if (rec.pending) {
      out.push('❌ ' + name + ' · не завершилась');
      continue;
    }
    if (rec.ok) {
      line = '✅ ' + name;
      if (rec.detail) line += ': ' + rec.detail;
      out.push(line);
    } else {
      line = '❌ ' + name;
      if (rec.detail) line += ' · ' + rec.detail;
      out.push(line);
    }
  }
  if (typeof tgSend_ === 'function') tgSend_(out.join('\n'));
}

var MORNING_CHAIN_HANDLERS_ = [
  'runMorningChainTick_',
  'runMorningChainWatchdog_',
  'runMorningChain2_',
  'runMorningChainWaitStatuses_',
  'runMorningChainFbs_',
  'runMorningChainWaitCompose_',
  'runMorningChain3_'
];

/** Свойство WB_MORNING_SKIP_STOCKS: 1=пропуск остатков FBW в утре, 0 или пусто=грузить. */
var MORNING_SKIP_FBW_STOCKS_DEFAULT_ = false;

function morningSkipFbwStocks_() {
  var p = PropertiesService.getScriptProperties().getProperty('WB_MORNING_SKIP_STOCKS');
  if (p === '0') return false;
  if (p === '1') return true;
  return MORNING_SKIP_FBW_STOCKS_DEFAULT_;
}

function morningChainJobTimeoutMs_(jobId) {
  var m = {
    finance: 8 * 60 * 1000,
    ads: 8 * 60 * 1000,
    funnel: 8 * 60 * 1000,
    orders: 10 * 60 * 1000,
    products: 12 * 60 * 1000,
    fbwStocks: 22 * 60 * 1000,
    lkStocks: 12 * 60 * 1000,
    transit: 10 * 60 * 1000,
    supplies: 22 * 60 * 1000,
    fbsOrders: 15 * 60 * 1000,
    fbsSupplies: 15 * 60 * 1000
  };
  return m[jobId] || 8 * 60 * 1000;
}

function morningChainSheetLockHeld_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(50)) return true;
  lock.releaseLock();
  return false;
}

/** Замок на запись в книгу; при занятости — отложить handler через afterMs. */
function morningChainSheetLockOrReschedule_(handlerName, afterMs, fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    Logger.log('Замок занят — ' + handlerName + ' через ' + afterMs + ' мс');
    morningChainSchedule_(handlerName, afterMs || 45 * 1000);
    return false;
  }
  try {
    fn();
    return true;
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}

function morningChainScheduleWatchdog_(jobId) {
  if (typeof clearTriggersForHandler_ === 'function') clearTriggersForHandler_('runMorningChainWatchdog_');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WB_MORNING_WATCHDOG_JOB', jobId);
  ScriptApp.newTrigger('runMorningChainWatchdog_')
    .timeBased()
    .after(morningChainJobTimeoutMs_(jobId))
    .create();
}

function runMorningChainWatchdog_() {
  if (typeof clearTriggersForHandler_ === 'function') clearTriggersForHandler_('runMorningChainWatchdog_');
  if (!morningChainOn_()) return;
  var props = PropertiesService.getScriptProperties();
  var wdJob = props.getProperty('WB_MORNING_WATCHDOG_JOB');
  var id = props.getProperty('WB_MORNING_JOB') || '';
  if (!wdJob || wdJob !== id) return;
  var job = morningChainJob_(id);
  if (!job) return;

  if (morningChainSheetLockHeld_()) {
    Logger.log('Утренняя цепочка: сторож — шаг ' + id + ' ещё выполняется, продлеваем');
    morningChainScheduleWatchdog_(id);
    return;
  }

  Logger.log('Утренняя цепочка: сторож — таймаут шага ' + id);
  if (typeof morningTgNote_ === 'function') {
    morningTgNote_(false, job.tg, 'время ожидания истекло');
  }
  props.deleteProperty('WB_MORNING_STEP_RUNNING');
  props.deleteProperty('WB_MORNING_WATCHDOG_JOB');
  morningChainAdvance_(job);
}

var MORNING_CHAIN_JOBS_ = [
  { id: 'finance', tg: 'Финансы' },
  { id: 'ads', tg: 'Реклама' },
  { id: 'funnel', tg: 'Воронка' },
  { id: 'orders', tg: 'Заказы' },
  { id: 'products', tg: 'Товары' },
  { id: 'fbwStocks', tg: 'Остатки' },
  { id: 'lkStocks', tg: 'Остатки ЛК' },
  { id: 'transit', tg: 'В пути' },
  { id: 'supplies', tg: 'Поставки', next: 'waitStatuses' },
  { id: 'fbsOrders', tg: 'FBS заказы' },
  { id: 'fbsSupplies', tg: 'FBS поставки', next: 'fbsRest' }
];

function setupMorningChainTrigger() {
  clearTriggersForHandler_('runMorningChain');
  ScriptApp.newTrigger('runMorningChain')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
  var msg = 'Утренняя цепочка: ежедневно 5:00 (' + Session.getScriptTimeZone() + ').\n' +
    'Четыре старых триггера (ч1, ч2, FBS, ч3) лучше выключить, функции не удалял.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function runMorningChain() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WB_MORNING_CHAIN', '1');
  props.deleteProperty('WB_MORNING_CHAIN_WAIT');
  props.deleteProperty('WB_MORNING_CHAIN_WAIT_FBS');
  props.deleteProperty('WB_MORNING_SAFETY');
  props.setProperty('WB_MORNING_JOB', 'finance');
  morningTgInit_();
  morningChainClearPending_();
  Logger.log('Утренняя цепочка: старт');
  morningChainSchedule_('runMorningChainTick_', 1000);
}

function runMorningChain2_() {
  morningChainClearSelf_('runMorningChain2_');
  if (!morningChainOn_()) return;
  Logger.log('Утренняя цепочка: ч2 (товары)');
  PropertiesService.getScriptProperties().setProperty('WB_MORNING_JOB', 'products');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_SAFETY');
  morningChainSchedule_('runMorningChainTick_', 1000);
}

function runMorningChainFbs_() {
  morningChainClearSelf_('runMorningChainFbs_');
  if (!morningChainOn_()) return;
  Logger.log('Утренняя цепочка: FBS');
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  PropertiesService.getScriptProperties().setProperty('WB_MORNING_JOB', 'fbsOrders');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_SAFETY');
  morningChainSchedule_('runMorningChainTick_', 1000);
}

function morningChainJob_(id) {
  var i;
  for (i = 0; i < MORNING_CHAIN_JOBS_.length; i++) {
    if (MORNING_CHAIN_JOBS_[i].id === id) return MORNING_CHAIN_JOBS_[i];
  }
  return null;
}

function morningChainClearSheetFilter_(name) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (sh && typeof wbClearFilter_ === 'function') wbClearFilter_(sh);
  } catch (e) {}
}

function runMorningChainTick_() {
  morningChainClearSelf_('runMorningChainTick_');
  if (!morningChainOn_()) return;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('WB_MORNING_JOB') || 'finance';
  var job = morningChainJob_(id);
  if (!job) {
    Logger.log('Утренняя цепочка: tick без рабочего шага (' + id + ')');
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    Logger.log('Утренняя цепочка: ждём замок, повтор шага ' + job.id);
    morningChainSchedule_('runMorningChainTick_', 30 * 1000);
    return;
  }

  try {
    if (typeof morningTgPending_ === 'function') morningTgPending_(job.tg);
    props.setProperty('WB_MORNING_STEP_RUNNING', job.id);
    morningChainScheduleWatchdog_(job.id);
    Logger.log('Утренняя цепочка: шаг ' + job.id);
    morningChainRunStep_(job.id, function() { morningChainRunJob_(job.id); });
    if (!morningChainOn_()) return;
    props.deleteProperty('WB_MORNING_STEP_RUNNING');
    props.deleteProperty('WB_MORNING_WATCHDOG_JOB');
    if (typeof clearTriggersForHandler_ === 'function') clearTriggersForHandler_('runMorningChainWatchdog_');
    morningChainAdvance_(job);
  } finally {
    try { lock.releaseLock(); } catch (eR) {}
  }
}

function morningChainStopTick_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('WB_MORNING_SAFETY');
  props.deleteProperty('WB_MORNING_STEP_RUNNING');
  props.deleteProperty('WB_MORNING_WATCHDOG_JOB');
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runMorningChainTick_');
    clearTriggersForHandler_('runMorningChainWatchdog_');
  }
}

function morningChainAdvance_(job) {
  var props = PropertiesService.getScriptProperties();
  if (job.next === 'waitStatuses') {
    morningChainStopTick_();
    props.setProperty('WB_MORNING_JOB', 'waitStatuses');
    scheduleUpdateSuppliesStatuses_();
    morningChainSchedule_('runMorningChainWaitStatuses_', 45 * 1000);
    return;
  }
  if (job.next === 'fbsRest') {
    morningChainStopTick_();
    props.setProperty('WB_MORNING_JOB', 'waitFbs');
    if (typeof scheduleFbsStocks_ === 'function') scheduleFbsStocks_(20 * 1000);
    morningChainSchedule_('runMorningChainWaitCompose_', 45 * 1000);
    return;
  }
  var i;
  for (i = 0; i < MORNING_CHAIN_JOBS_.length; i++) {
    if (MORNING_CHAIN_JOBS_[i].id === job.id) {
      if (i + 1 < MORNING_CHAIN_JOBS_.length) {
        props.setProperty('WB_MORNING_JOB', MORNING_CHAIN_JOBS_[i + 1].id);
        morningChainSchedule_('runMorningChainTick_', 20 * 1000);
      } else {
        morningChainStopTick_();
        morningChainSchedule_('runMorningChainWaitStatuses_', 45 * 1000);
      }
      return;
    }
  }
}

function morningChainRunJob_(id) {
  try {
    if (id === 'finance') morningJobFinance_();
    else if (id === 'ads') morningJobAds_();
    else if (id === 'funnel') morningJobFunnel_();
    else if (id === 'orders') morningJobOrders_();
    else if (id === 'products') morningJobProducts_();
    else if (id === 'fbwStocks') morningJobFbwStocks_();
    else if (id === 'lkStocks') morningJobLkStocks_();
    else if (id === 'transit') morningJobTransit_();
    else if (id === 'supplies') morningJobSupplies_();
    else if (id === 'fbsOrders') morningJobFbsOrders_();
    else if (id === 'fbsSupplies') morningJobFbsSupplies_();
  } finally {
    try { wbPutAlwaysFilters_(); } catch (eF) {}
  }
}

function morningYesterday_() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return Utilities.formatDate(yesterday, 'GMT+3', 'yyyy-MM-dd');
}

function morningJobToken_() {
  try { return getApiToken(); } catch (e) { return ''; }
}

function morningJobFinance_() {
  var token = morningJobToken_();
  if (!token) { morningTgNote_(false, 'Финансы', 'нет токена'); return; }
  morningChainClearSheetFilter_('Финансовые_ежедневные');
  var lines = [];
  var dateStr = morningYesterday_();
  tgRunStep_(lines, 'Финансы', function() {
    tgWatch_(lines, 'Финансовые_ежедневные', function() {
      loadFinancialReportWithPeriod(token, dateStr, dateStr, 'daily');
    }, 'Финансы');
  });
}

function morningJobAds_() {
  var token = morningJobToken_();
  if (!token) { morningTgNote_(false, 'Реклама', 'нет токена'); return; }
  morningChainClearSheetFilter_('Рекламные_расходы');
  var lines = [];
  var dateStr = morningYesterday_();
  tgRunStep_(lines, 'Реклама', function() {
    tgWatch_(lines, 'Рекламные_расходы', function() {
      loadAdvExpenses(token, dateStr, dateStr);
    }, 'Реклама');
  });
}

function morningJobFunnel_() {
  var token = morningJobToken_();
  if (!token) { morningTgNote_(false, 'Воронка', 'нет токена'); return; }
  morningChainClearSheetFilter_('Воронка продаж');
  morningChainClearSheetFilter_('Аналитика_проба_дни');
  var lines = [];
  var dateStr = morningYesterday_();
  tgRunStep_(lines, 'Воронка', function() {
    tgWatch_(lines, 'Воронка продаж', function() {
      loadSalesFunnelRange(token, dateStr, dateStr);
    }, 'Воронка');
  });
}

function morningJobOrders_() {
  if (!morningJobToken_()) { morningTgNote_(false, 'Заказы', 'нет токена'); return; }
  morningChainClearSheetFilter_('Заказы');
  var lines = [];
  tgRunStep_(lines, 'Заказы', function() {
    tgWatch_(lines, 'Заказы', function() {
      loadOrdersForYesterdayAuto();
    });
  });
}

function morningJobProducts_() {
  if (!morningJobToken_()) {
    morningTgNote_(false, 'Товары', 'нет токена');
    morningTgNote_(false, 'История цен', 'нет токена');
    return;
  }
  morningChainClearSheetFilter_('Товары');
  var lines = [];
  tgRunStep_(lines, 'Товары', function() {
    tgWatch_(lines, 'Товары', function() {
      loadAllProducts();
    });
  });
  morningChainClearSheetFilter_('История цен');
  tgRunStep_(lines, 'История цен', function() {
    tgWatch_(lines, 'История цен', function() {
      savePriceHistory();
    });
  });
}

function morningJobLkStocks_() {
  if (!morningJobToken_()) { morningTgNote_(false, 'Остатки ЛК', 'нет токена'); return; }
  morningChainClearSheetFilter_('Остатки ЛК');
  var lines = [];
  tgRunStep_(lines, 'Остатки ЛК', function() {
    tgWatch_(lines, 'Остатки ЛК', function() {
      recordLkStocks();
    });
  });
}

function morningJobFbwStocks_() {
  if (morningSkipFbwStocks_()) {
    if (typeof morningTgNote_ === 'function') morningTgNote_(true, 'Остатки', 'пропуск (ручная загрузка)');
    Logger.log('Утренняя цепочка: остатки FBW пропущены');
    return;
  }
  if (!morningJobToken_()) { morningTgNote_(false, 'Остатки', 'нет токена'); return; }
  morningChainClearSheetFilter_('Остатки');
  var lines = [];
  tgRunStep_(lines, 'Остатки', function() {
    tgWatch_(lines, 'Остатки', function() {
      recordCurrentStocksEnriched(true);
    });
  });
}

function morningJobTransit_() {
  morningChainClearSheetFilter_('В пути');
  var lines = [];
  tgRunStep_(lines, 'В пути', function() {
    tgWatch_(lines, 'В пути', function() {
      saveStocksHistory();
    });
  });
}

function morningJobSupplies_() {
  if (!morningJobToken_()) { morningTgNote_(false, 'Поставки', 'нет токена'); return; }
  morningChainClearSheetFilter_('Поставки_список');
  var lines = [];
  tgRunStep_(lines, 'Поставки', function() {
    tgWatch_(lines, 'Поставки_список', function() {
      updateSuppliesDaily();
    }, 'Поставки');
  });
}

function morningJobFbsOrders_() {
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  if (!morningJobToken_()) { morningTgNote_(false, 'FBS заказы', 'нет токена'); return; }
  var lines = [];
  try { WB_SILENT_UI = true; } catch (eS) {}
  try {
    tgRunStep_(lines, 'FBS заказы', function() {
      tgWatch_(lines, 'FBS_Заказы', function() {
        loadFbsOrders(true);
      });
    });
  } finally {
    try { WB_SILENT_UI = false; } catch (eS2) {}
  }
}

function morningJobFbsSupplies_() {
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  if (!morningJobToken_()) { morningTgNote_(false, 'FBS поставки', 'нет токена'); return; }
  var lines = [];
  try { WB_SILENT_UI = true; } catch (eS) {}
  try {
    tgRunStep_(lines, 'FBS поставки', function() {
      tgWatch_(lines, 'FBS_Поставки', function() {
        loadFbsSupplies(true);
      });
    });
  } finally {
    try { WB_SILENT_UI = false; } catch (eS2) {}
  }
}

function runMorningChainWaitStatuses_() {
  morningChainClearSelf_('runMorningChainWaitStatuses_');
  if (!morningChainOn_()) return;
  var busy = morningChainHasHandler_('runUpdateSuppliesStatusesJob_') ||
    morningChainHasHandler_('continueUpdateSuppliesStatusesJob_');
  var props = PropertiesService.getScriptProperties();
  var n = Number(props.getProperty('WB_MORNING_CHAIN_WAIT') || '0');
  if (busy && n < 20) {
    props.setProperty('WB_MORNING_CHAIN_WAIT', String(n + 1));
    Logger.log('Утренняя цепочка: ждём статусы поставок, круг ' + (n + 1));
    morningChainSchedule_('runMorningChainWaitStatuses_', 60 * 1000);
    return;
  }
  props.deleteProperty('WB_MORNING_CHAIN_WAIT');
  if (typeof morningTgGet_ === 'function' && morningTgGet_('Статусы FBW')) {
    morningTgNote_(!busy, 'Статусы FBW', busy ? 'не завершилась' : '');
  }
  morningChainSchedule_('runMorningChainFbs_', 90 * 1000);
}

function runMorningChainWaitCompose_() {
  morningChainClearSelf_('runMorningChainWaitCompose_');
  if (!morningChainOn_()) return;
  var busy = morningChainFbsBusy_();
  var props = PropertiesService.getScriptProperties();
  var n = Number(props.getProperty('WB_MORNING_CHAIN_WAIT_FBS') || '0');
  if (busy && n < 35) {
    props.setProperty('WB_MORNING_CHAIN_WAIT_FBS', String(n + 1));
    Logger.log('Утренняя цепочка: ждём сборку FBS, круг ' + (n + 1));
    morningChainSchedule_('runMorningChainWaitCompose_', 60 * 1000);
    return;
  }
  props.deleteProperty('WB_MORNING_CHAIN_WAIT_FBS');
  morningChainSchedule_('runMorningChain3_', 90 * 1000);
}

function runMorningChain3_() {
  morningChainClearSelf_('runMorningChain3_');
  if (!morningChainOn_()) return;
  Logger.log('Утренняя цепочка: ч3');
  if (!morningChainSheetLockOrReschedule_('runMorningChain3_', 60 * 1000, function() {
    morningChainRunStep_('ч3', function() { updateAllReports3(); });
  })) return;
  morningTgFlush_();
  try { поставитьФильтрыНаЛисты_([]); } catch (eF) {}
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_CHAIN');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_CHAIN_WAIT');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_CHAIN_WAIT_FBS');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_JOB');
  PropertiesService.getScriptProperties().deleteProperty('WB_MORNING_SAFETY');
  PropertiesService.getScriptProperties().deleteProperty('WB_FBS_COMPOSE_RUN');
  morningChainClearPending_();
  Logger.log('Утренняя цепочка: готово');
}

function morningChainOn_() {
  return PropertiesService.getScriptProperties().getProperty('WB_MORNING_CHAIN') === '1';
}

function morningChainRunStep_(name, fn) {
  try {
    fn();
  } catch (e) {
    Logger.log('Утренняя цепочка шаг ' + name + ': ' + e);
  }
}

function morningChainSchedule_(handler, afterMs) {
  if (typeof clearTriggersForHandler_ === 'function') clearTriggersForHandler_(handler);
  ScriptApp.newTrigger(handler).timeBased().after(afterMs).create();
  Logger.log('Утренняя цепочка: следующий ' + handler + ' через ' + afterMs + ' мс');
}

function morningChainClearSelf_(handler) {
  if (typeof clearTriggersForHandler_ === 'function') clearTriggersForHandler_(handler);
}

function morningChainClearPending_() {
  var i;
  for (i = 0; i < MORNING_CHAIN_HANDLERS_.length; i++) {
    if (typeof clearTriggersForHandler_ === 'function') {
      clearTriggersForHandler_(MORNING_CHAIN_HANDLERS_[i]);
    }
  }
}

function morningChainHasHandler_(name) {
  var list = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].getHandlerFunction() === name) return true;
  }
  return false;
}

function morningChainFbsBusy_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('WB_FBS_COMPOSE_BUSY') === '1') return true;
  if (props.getProperty('WB_FBS_COMPOSE_RUN') === '1') return true;
  var rec = typeof morningTgGet_ === 'function' ? morningTgGet_('Сборка FBS') : null;
  if (rec && rec.pending) return true;
  var job = props.getProperty('WB_MORNING_JOB');
  if ((job === 'fbsOrders' || job === 'fbsSupplies') &&
      morningChainHasHandler_('runMorningChainTick_')) {
    return true;
  }
  return morningChainHasHandler_('runFbsStocksJob_') ||
    morningChainHasHandler_('runComposeFbsJob_') ||
    morningChainHasHandler_('runComposeFbsCopyJob_');
}

/** Снимок фильтров: снять перед записью, вернуть в конце прогона. */
var WB_FILTER_SNAPSHOT_ = [];

/**
 * Снимает фильтры со всех листов (иначе запись шапки / insert колонки
 * даёт IllegalStateException). Диапазоны запоминаются и ставятся обратно
 * в конце updateAllReports1/2 через поставитьФильтрыНаЛисты_.
 */
function сброситьФильтрыНаВсехЛистах() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  WB_FILTER_SNAPSHOT_ = [];
  let resetCount = 0;

  for (const sheet of sheets) {
    try {
      const filter = sheet.getFilter();
      if (!filter) continue;
      const r = filter.getRange();
      WB_FILTER_SNAPSHOT_.push({
        name: sheet.getName(),
        startRow: r.getRow(),
        startCol: r.getColumn()
      });
      filter.remove();
      resetCount++;
    } catch (e) {
      Logger.log('⚠️ Фильтр «' + sheet.getName() + '»: ' + e.message);
    }
  }

  try { SpreadsheetApp.flush(); } catch (e2) {}
  Logger.log('✅ Фильтры сняты: ' + resetCount + ' (вернутся в конце прогона)');
}

/**
 * Ставит фильтры обратно.
 * Рабочие листы загрузок всегда получают фильтр с строки заголовка,
 * даже если снимок утром запомнил другую строку или фильтр уже висел ниже.
 * Отчёты (шапка не в строке 1) не двигаем, если фильтр уже стоит.
 */
var WB_FILTER_ALWAYS_ = [
  'Финансовые_ежедневные', 'Финансовые_недельные', 'Рекламные_расходы', 'Воронка продаж', 'Заказы', 'Товары',
  'Остатки', 'Остатки ЛК', 'В пути', 'Остатки_CSV_проба', 'Склады_WB',
  'Поставки_список', 'Детали_поставок', 'Товары_поставок',
  'Поставки', 'ЗаказыЛК',
  'FBS Остатки', 'FBS Заказы', 'FBS Поставки', 'FBS Поставка_состав',
  'FBS Офисы', 'FBS Склады'
];

/** Шапка не в первой строке. */
function wbFilterHeaderRow_(name) {
  if (name === 'ЗаказыЛК') return 2;
  return 1;
}

function wbPutFilter_(sheet, startRow, startCol, force) {
  if (!sheet) return false;
  startRow = startRow || 1;
  startCol = startCol || 1;
  var lastRow = Math.max(sheet.getLastRow(), startRow);
  var lastCol = Math.max(sheet.getLastColumn(), startCol);
  var numRows = lastRow - startRow + 1;
  var numCols = lastCol - startCol + 1;
  if (numRows < 1 || numCols < 1) return false;
  var filter = null;
  try { filter = sheet.getFilter(); } catch (e0) { filter = null; }
  if (filter) {
    var range = null;
    try { range = filter.getRange(); } catch (e1) { range = null; }
    if (range && range.getRow() === startRow && range.getColumn() === startCol &&
        range.getLastRow() >= lastRow && range.getLastColumn() >= lastCol) {
      return true;
    }
    if (!force) return true;
    try { filter.remove(); } catch (eR) {}
  }
  sheet.getRange(startRow, startCol, numRows, numCols).createFilter();
  return true;
}

function wbPutSheetFilter_(sheet, startRow) {
  if (!sheet) return;
  try {
    wbPutFilter_(sheet, startRow || 1, 1, true);
  } catch (e) {
    Logger.log('Фильтр «' + sheet.getName() + '»: ' + e);
  }
}

function wbPutNamedFilters_(names) {
  var ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e0) { return; }
  if (!ss || !names) return;
  var i, sh;
  for (i = 0; i < names.length; i++) {
    sh = ss.getSheetByName(names[i]);
    if (sh) wbPutSheetFilter_(sh, wbFilterHeaderRow_(names[i]));
  }
}

function wbPutAlwaysFilters_() {
  wbPutNamedFilters_(WB_FILTER_ALWAYS_);
  try {
    if (typeof fbsSidecar_ !== 'function') return;
    var sc = fbsSidecar_();
    var i, sh2;
    for (i = 0; i < WB_FILTER_ALWAYS_.length; i++) {
      sh2 = sc.getSheetByName(WB_FILTER_ALWAYS_[i]);
      if (sh2) wbPutSheetFilter_(sh2, wbFilterHeaderRow_(WB_FILTER_ALWAYS_[i]));
    }
  } catch (e) {}
}

function поставитьФильтрыНаЛисты_(lines) {
  try { SpreadsheetApp.flush(); } catch (e0) {}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var plan = {};
  var snapshot = WB_FILTER_SNAPSHOT_ || [];
  var i;
  for (i = 0; i < snapshot.length; i++) {
    plan[snapshot[i].name] = {
      startRow: snapshot[i].startRow || 1,
      startCol: snapshot[i].startCol || 1
    };
  }
  for (i = 0; i < WB_FILTER_ALWAYS_.length; i++) {
    if (!plan[WB_FILTER_ALWAYS_[i]]) {
      plan[WB_FILTER_ALWAYS_[i]] = { startRow: 1, startCol: 1 };
    }
  }

  var ok = 0;
  var fail = 0;
  var names = Object.keys(plan);
  for (i = 0; i < names.length; i++) {
    try {
      var sheet = ss.getSheetByName(names[i]);
      if (!sheet) continue;
      var startRow = plan[names[i]].startRow;
      var startCol = plan[names[i]].startCol;
      var force = false;
      if (WB_FILTER_ALWAYS_.indexOf(names[i]) !== -1) {
        startRow = wbFilterHeaderRow_(names[i]);
        startCol = 1;
        force = true;
      }
      if (wbPutFilter_(sheet, startRow, startCol, force)) ok++;
    } catch (e) {
      fail++;
      Logger.log('⚠️ Фильтр обратно «' + names[i] + '»: ' + e);
    }
  }
  try {
    if (typeof fbsSidecar_ === 'function') {
      var sc = fbsSidecar_();
      for (i = 0; i < WB_FILTER_ALWAYS_.length; i++) {
        var sh2 = sc.getSheetByName(WB_FILTER_ALWAYS_[i]);
        if (!sh2) continue;
        try {
          if (wbPutFilter_(sh2, wbFilterHeaderRow_(WB_FILTER_ALWAYS_[i]), 1, true)) ok++;
        } catch (e3) {
          fail++;
          Logger.log('⚠️ Фильтр sidecar «' + WB_FILTER_ALWAYS_[i] + '»: ' + e3);
        }
      }
    }
  } catch (e2) {}

  WB_FILTER_SNAPSHOT_ = [];
  var msg = 'фильтры: ' + ok + (fail ? ', ошибок: ' + fail : '');
  Logger.log('✅ ' + msg);
  if (lines) lines.push(msg);
  return ok;
}

/**
 *  Заменяет старые названия складов на новые в заданных колонках всех листов.
 * Запускать ПОСЛЕ обновления данных.
 */
function fixWarehouseNamesSmart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetName = "Склад Шушары";

  // 🔍 Ключевые слова для поиска (регистр и пробелы не важны)
  // Если в ячейке встречается любое из этих слов → она будет заменена
  const searchKeywords = ["шушары", "спб шушары", "шушары московское"];

  const config = [
    { sheet: "Финансовые_ежедневные", cols: ["office_name"] },
    { sheet: "Заказы", cols: ["warehouseName"] },
    { sheet: "Остатки", cols: ["Склад"] },
    { sheet: "Остатки ЛК", cols: ["Склад"] },
    { sheet: "В пути", cols: ["Склад"] },
    { sheet: "Склады_WB", cols: ["Название"] },
    { sheet: "Транзитные_тарифы", cols: ["Откуда", "Куда"] },
    { sheet: "Детали_поставок", cols: ["warehouseName"] },
    { sheet: "Поставки", cols: ["Склад"] }
  ];

  let totalScanned = 0;
  let totalReplaced = 0;
  const changeLog = [];

  for (const item of config) {
    const sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const colIndices = item.cols
      .map(col => findHeaderColIndex_(headers, col))
      .filter(idx => idx !== -1);

    if (colIndices.length === 0) continue;

    for (let r = 1; r < lastRow; r++) {
      for (const c of colIndices) {
        const rawVal = data[r][c];
        if (rawVal == null || rawVal === "") continue;

        const strVal = String(rawVal).trim();
        // Пропускаем, если уже целевое название
        if (strVal === targetName) continue;

        // Нормализуем: убираем регистр, лишние пробелы, дефисы, подчёркивания
        const lowerVal = strVal.toLowerCase().replace(/[\s\-\_]+/g, ' ');

        let matchFound = false;
        for (const kw of searchKeywords) {
          if (lowerVal.includes(kw.toLowerCase())) {
            matchFound = true;
            break;
          }
        }

        if (matchFound) {
          changeLog.push(`[${item.sheet}] Строка ${r+1}: "${strVal}"`);
          data[r][c] = targetName;
          totalReplaced++;
        }
        totalScanned++;
      }
    }

    if (totalReplaced > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).setValues(data);
    }
  }

  //  Вывод результатов
  if (changeLog.length > 0) {
    Logger.log(`✅ Заменено ${totalReplaced} значений. Примеры изменений:`);
    Logger.log(changeLog.slice(0, 20).join('\n'));
    if (changeLog.length > 20) Logger.log(`... и ещё ${changeLog.length - 20}`);
  } else {
    Logger.log("ℹ️ Совпадений по ключевым словам не найдено.");
  }

  SpreadsheetApp.getUi().alert(
    '✅ Готово',
    `Проверено ячеек: ${totalScanned}\nЗаменено: ${totalReplaced}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Заменяет названия складов на основе листа "Склады_проверка" ВСЕ СПИСКИ БЕЗ ДАТ
 * - Читает правила из столбцов: [1] Склад → [7] Объединение
 * - Заменяет ТОЛЬКО если "Объединение" НЕ пустое
 * - Работает во всех листах из конфига
 */
function fixWarehouseNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 🔹 1. Читаем правила замены из "Склады_проверка"
  const refSheet = ss.getSheetByName("Склады_проверка");
  if (!refSheet) {
    safeAlert('❌ Лист "Склады_проверка" не найден');
    return;
  }
  
  const refData = refSheet.getDataRange().getValues();
  const replaceMap = {}; // { "оригинал": "замена" }
  
  // Пропускаем заголовок, читаем столбцы 1 (А) и 7 (G)
  for (let i = 1; i < refData.length; i++) {
    const original = refData[i][0]?.toString().trim();
    const merged = refData[i][6]?.toString().trim();
    
    // 🔥 Заменяем ТОЛЬКО если "Объединение" НЕ пустое
    if (original && merged && merged !== "") {
      replaceMap[original] = merged;
    }
  }
  
  if (Object.keys(replaceMap).length === 0) {
    safeAlert('ℹ️ Нет правил замены (столбец "Объединение" пуст)');
    return;
  }
  
  // 🔹 2. Конфигурация листов и колонок для замены
  const config = [
    { sheet: "Финансовые_ежедневные", cols: ["office_name"] },
    { sheet: "Заказы", cols: ["warehouseName"] },
    { sheet: "Остатки", cols: ["Склад"] },
    { sheet: "Остатки ЛК", cols: ["Склад"] },
    { sheet: "В пути", cols: ["Склад"] },
    { sheet: "Склады_WB", cols: ["Название"] },
    { sheet: "Транзитные_тарифы", cols: ["Откуда", "Куда"] },
    { sheet: "Детали_поставок", cols: ["warehouseName"] },
    { sheet: "Поставки", cols: ["Склад"] }
  ];
  
  let totalScanned = 0;
  let totalReplaced = 0;
  const changeLog = [];
  
  // 🔹 3. Проходим по всем листам
  for (const item of config) {
    const sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    
    // Находим индексы целевых колонок
    const colIndices = item.cols
      .map(col => findHeaderColIndex_(headers, col))
      .filter(idx => idx !== -1);
    
    if (colIndices.length === 0) continue;
    
    let sheetReplaced = 0;
    
    // Проходим по всем строкам
    for (let r = 1; r < lastRow; r++) {
      for (const c of colIndices) {
        const rawVal = data[r][c];
        if (rawVal == null || rawVal === "") continue;
        
        const strVal = String(rawVal).trim();
        
        // 🔥 Если есть правило замены → применяем
        if (replaceMap[strVal]) {
          changeLog.push(`[${item.sheet}] Строка ${r+1}: "${strVal}" → "${replaceMap[strVal]}"`);
          data[r][c] = replaceMap[strVal];
          sheetReplaced++;
          totalReplaced++;
        }
        totalScanned++;
      }
    }
    
    // Сохраняем изменения только если были замены
    if (sheetReplaced > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).setValues(data);
    }
  }
  
  // 🔹 4. Вывод результатов
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  
  if (changeLog.length > 0) {
    Logger.log(`✅ Заменено ${totalReplaced} значений за ${duration} сек. Примеры:`);
    Logger.log(changeLog.slice(0, 20).join('\n'));
    if (changeLog.length > 20) Logger.log(`... и ещё ${changeLog.length - 20}`);
  } else {
    Logger.log(`ℹ️ Замен не найдено (проверено ${totalScanned} ячеек)`);
  }
  
  safeAlert(`✅ Готово!\nПроверено: ${totalScanned}\nЗаменено: ${totalReplaced}\nВремя: ${duration} сек`);
}

/**
 * Безопасный вывод алерта (работает везде)
 */
function safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

/**
 * Заменяет названия складов на основе листа "Склады_проверка" ВСЕ СПИСКИ БЕЗ ДАТ
 * 🔥 ОПТИМИЗИРОВАННАЯ ВЕРСИЯ: читает только нужные колонки, пишет только изменения
 */
function fixWarehouseNames2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 🔹 1. Читаем правила замены из "Склады_проверка"
  const refSheet = ss.getSheetByName("Склады_проверка");
  if (!refSheet) {
    safeAlert('❌ Лист "Склады_проверка" не найден');
    return;
  }
  
  const refData = refSheet.getDataRange().getValues();
  const replaceMap = new Map(); // 🔥 Map быстрее для поиска
  
  for (let i = 1; i < refData.length; i++) {
    const original = refData[i][0]?.toString().trim();
    const merged = refData[i][6]?.toString().trim();
    if (original && merged && merged !== "") {
      replaceMap.set(original, merged);
    }
  }
  
  if (replaceMap.size === 0) {
    safeAlert('ℹ️ Нет правил замены (столбец "Объединение" пуст)');
    return;
  }
  
  // 🔹 2. Конфигурация
  const config = [
    { sheet: "Финансовые_ежедневные", cols: ["office_name"] },
    { sheet: "Заказы", cols: ["warehouseName"] },
    { sheet: "Остатки", cols: ["Склад"] },
    { sheet: "Остатки ЛК", cols: ["Склад"] },
    { sheet: "В пути", cols: ["Склад"] },
    { sheet: "Склады_WB", cols: ["Название"] },
    { sheet: "Транзитные_тарифы", cols: ["Откуда", "Куда"] },
    { sheet: "Детали_поставок", cols: ["warehouseName"] },
    { sheet: "Поставки", cols: ["Склад"] }
  ];
  
  let totalScanned = 0;
  let totalReplaced = 0;
  
  // 🔹 3. Проходим по всем листам
  for (const item of config) {
    const sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    // 🔥 Читаем ТОЛЬКО заголовки
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Находим индексы целевых колонок
    const colIndices = item.cols
      .map(col => findHeaderColIndex_(headers, col))
      .filter(idx => idx !== -1);
    
    if (colIndices.length === 0) continue;
    
    // 🔥 Читаем ТОЛЬКО целевые колонки (не весь лист!)
    const minCol = Math.min(...colIndices);
    const maxCol = Math.max(...colIndices);
    const colCount = maxCol - minCol + 1;
    
    const data = sheet.getRange(2, minCol + 1, lastRow - 1, colCount).getValues();
    
    // 🔥 Собираем изменения для пакетной записи
    const updates = [];
    
    for (let r = 0; r < data.length; r++) {
      for (const c of colIndices) {
        const relativeCol = c - minCol;
        const rawVal = data[r][relativeCol];
        if (rawVal == null || rawVal === "") continue;
        
        const strVal = String(rawVal).trim();
        totalScanned++;
        
        if (replaceMap.has(strVal)) {
          updates.push({ row: r + 2, col: c + 1, value: replaceMap.get(strVal) });
          totalReplaced++;
        }
      }
    }
    
    // 🔥 Записываем только изменённые ячейки (гораздо быстрее чем весь диапазон)
    if (updates.length > 0) {
      for (const u of updates) {
        sheet.getRange(u.row, u.col).setValue(u.value);
      }
    }
  }
  
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  safeAlert(`✅ Готово!\nПроверено: ${totalScanned}\nЗаменено: ${totalReplaced}\nВремя: ${duration} сек`);
}

/**
 * Заменяет названия складов на основе листа "Склады_проверка" НА ПОСЛЕДНЮЮ ДАТУ
 * - Читает правила из столбцов: [1] Склад → [7] Объединение
 * - Заменяет ТОЛЬКО если "Объединение" НЕ пустое
 * - Для листов с датой: заменяет ТОЛЬКО в строках с ПОСЛЕДНЕЙ датой
 * - Для листов без даты: заменяет во ВСЕХ строках
 */
function fixWarehouseNamesLastDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 🔹 1. Читаем правила замены из "Склады_проверка"
  const refSheet = ss.getSheetByName("Склады_проверка");
  if (!refSheet) {
    safeAlert('❌ Лист "Склады_проверка" не найден');
    return;
  }
  
  const refData = refSheet.getDataRange().getValues();
  const replaceMap = {};
  
  for (let i = 1; i < refData.length; i++) {
    const original = refData[i][0]?.toString().trim();
    const merged = refData[i][6]?.toString().trim();
    if (original && merged && merged !== "") {
      replaceMap[original] = merged;
    }
  }
  
  if (Object.keys(replaceMap).length === 0) {
    safeAlert('ℹ️ Нет правил замены (столбец "Объединение" пуст)');
    return;
  }
  
  // 🔹 2. Конфигурация: лист → колонки + колонка даты (null = без фильтра)
  const config = [
    { sheet: "Финансовые_ежедневные", cols: ["office_name"], dateCol: "Дата" },
    { sheet: "Заказы", cols: ["warehouseName"], dateCol: "Дата" },
    { sheet: "Остатки", cols: ["Склад"], dateCol: "Дата" },
    { sheet: "Остатки ЛК", cols: ["Склад"], dateCol: "Дата" },
    { sheet: "В пути", cols: ["Склад"], dateCol: null },           // 🔥 без фильтра
    { sheet: "Склады_WB", cols: ["Название"], dateCol: null },     // 🔥 без фильтра
    { sheet: "Транзитные_тарифы", cols: ["Откуда", "Куда"], dateCol: null }, // 🔥 без фильтра
    { sheet: "Детали_поставок", cols: ["warehouseName"], dateCol: null },    // 🔥 без фильтра
    { sheet: "Поставки", cols: ["Склад"], dateCol: "Дата поставки" }
  ];
  
  let totalScanned = 0;
  let totalReplaced = 0;
  const changeLog = [];
  
  // 🔹 3. Проходим по всем листам
  for (const item of config) {
    const sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    
    // Находим индексы целевых колонок
    const colIndices = item.cols
      .map(col => findHeaderColIndex_(headers, col))
      .filter(idx => idx !== -1);
    
    if (colIndices.length === 0) continue;
    
    // 🔍 Определяем последнюю дату (если есть фильтр)
    let dateIdx = -1;
    let latestDate = null;
    
    if (item.dateCol) {
      dateIdx = findHeaderColIndex_(headers, item.dateCol);
      if (dateIdx !== -1) {
        for (let r = 1; r < lastRow; r++) {
          const cellDate = parseDate(data[r][dateIdx]);
          if (cellDate && (!latestDate || cellDate > latestDate)) {
            latestDate = cellDate;
          }
        }
      }
    }
    
    let sheetReplaced = 0;
    
    // Проходим по всем строкам
    for (let r = 1; r < lastRow; r++) {
      // 🔥 Если есть фильтр по дате — пропускаем строки не с последней датой
      if (dateIdx !== -1 && latestDate) {
        const rowDate = parseDate(data[r][dateIdx]);
        if (!rowDate || rowDate.getTime() !== latestDate.getTime()) {
          continue;
        }
      }
      
      for (const c of colIndices) {
        const rawVal = data[r][c];
        if (rawVal == null || rawVal === "") continue;
        
        const strVal = String(rawVal).trim();
        
        // 🔥 Если есть правило замены → применяем
        if (replaceMap[strVal]) {
          changeLog.push(`[${item.sheet}] Строка ${r+1}: "${strVal}" → "${replaceMap[strVal]}"`);
          data[r][c] = replaceMap[strVal];
          sheetReplaced++;
          totalReplaced++;
        }
        totalScanned++;
      }
    }
    
    // Сохраняем изменения только если были замены
    if (sheetReplaced > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).setValues(data);
    }
  }
  
  // 🔹 4. Вывод результатов
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  
  if (changeLog.length > 0) {
    Logger.log(`✅ Заменено ${totalReplaced} значений за ${duration} сек. Примеры:`);
    Logger.log(changeLog.slice(0, 20).join('\n'));
    if (changeLog.length > 20) Logger.log(`... и ещё ${changeLog.length - 20}`);
  } else {
    Logger.log(`ℹ️ Замен не найдено (проверено ${totalScanned} ячеек)`);
  }
  
  safeAlert(`✅ Готово!\nПроверено: ${totalScanned}\nЗаменено: ${totalReplaced}\nВремя: ${duration} сек`);
}

/**
 * Заменяет названия складов на основе листа "Склады_проверка" НА ПОСЛЕДНЮЮ ДАТУ
 * 🔥 ИСПРАВЛЕННАЯ ВЕРСИЯ: корректная запись значений
 */
function fixWarehouseNamesLastDate2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 🔹 1. Читаем правила замены из "Склады_проверка"
  const refSheet = ss.getSheetByName("Склады_проверка");
  if (!refSheet) {
    safeAlert('❌ Лист "Склады_проверка" не найден');
    return { scanned: 0, replaced: 0 };
  }
  
  const refData = refSheet.getDataRange().getValues();
  const replaceMap = new Map();
  
  for (let i = 1; i < refData.length; i++) {
    const original = refData[i][0]?.toString().trim();
    const merged = refData[i][6]?.toString().trim();
    if (original && merged && merged !== "") {
      replaceMap.set(original, merged);
    }
  }
  
  if (replaceMap.size === 0) {
    safeAlert('ℹ️ Нет правил замены (столбец "Объединение" пуст)');
    return { scanned: 0, replaced: 0 };
  }
  
  // 🔹 2. Конфигурация
  const config = [
    { sheet: "Финансовые_ежедневные", cols: ["office_name"], dateCol: "Дата" },
    { sheet: "Заказы", cols: ["warehouseName"], dateCol: "Дата" },
    { sheet: "Остатки", cols: ["Склад"], dateCol: "Дата" },
    { sheet: "Остатки ЛК", cols: ["Склад"], dateCol: "Дата" },
    { sheet: "В пути", cols: ["Склад"], dateCol: null },
    { sheet: "Склады_WB", cols: ["Название"], dateCol: null },
    { sheet: "Транзитные_тарифы", cols: ["Откуда", "Куда"], dateCol: null },
    { sheet: "Детали_поставок", cols: ["warehouseName"], dateCol: null },
    { sheet: "Поставки", cols: ["Склад"], dateCol: "Дата поставки" }
  ];
  
  let totalScanned = 0;
  let totalReplaced = 0;
  
  // 🔹 3. Проходим по всем листам
  for (const item of config) {
    const sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    // 🔥 Читаем ТОЛЬКО заголовки
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Находим индексы целевых колонок
    const colIndices = item.cols
      .map(col => findHeaderColIndex_(headers, col))
      .filter(idx => idx !== -1);
    
    if (colIndices.length === 0) continue;
    
    // 🔍 Определяем последнюю дату (если есть фильтр)
    let dateIdx = -1;
    let latestDate = null;
    
    if (item.dateCol) {
      dateIdx = findHeaderColIndex_(headers, item.dateCol);
      if (dateIdx !== -1) {
        // 🔥 Читаем ТОЛЬКО колонку даты
        const dateValues = sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
        for (let r = 0; r < dateValues.length; r++) {
          const cellDate = parseDate(dateValues[r][0]);
          if (cellDate && (!latestDate || cellDate > latestDate)) {
            latestDate = cellDate;
          }
        }
        if (!latestDate) continue;
      }
    }
    
    // 🔥 Читаем ТОЛЬКО нужные колонки
    const colsToRead = [dateIdx, ...colIndices].filter(idx => idx !== -1);
    const minCol = Math.min(...colsToRead);
    const maxCol = Math.max(...colsToRead);
    const colCount = maxCol - minCol + 1;
    
    const data = sheet.getRange(2, minCol + 1, lastRow - 1, colCount).getValues();
    
    // 🔥 Находим строки для обработки
    const rowsToProcess = [];
    if (item.dateCol && dateIdx !== -1 && latestDate) {
      const relativeDateIdx = dateIdx - minCol;
      for (let r = 0; r < data.length; r++) {
        const rowDate = parseDate(data[r][relativeDateIdx]);
        if (rowDate && rowDate.getTime() === latestDate.getTime()) {
          rowsToProcess.push(r);
        }
      }
    } else {
      for (let r = 0; r < data.length; r++) rowsToProcess.push(r);
    }
    
    // 🔥 Собираем изменения
    const updates = [];
    
    for (const r of rowsToProcess) {
      for (const c of colIndices) {
        const relativeCol = c - minCol;
        const rawVal = data[r][relativeCol];
        if (rawVal == null || rawVal === "") continue;
        
        const strVal = String(rawVal).trim();
        totalScanned++;
        
        if (replaceMap.has(strVal)) {
          updates.push({ row: r + 2, col: c + 1, value: replaceMap.get(strVal) });
          totalReplaced++;
        }
      }
    }
    
    // 🔥 🔧 ИСПРАВЛЕНИЕ: записываем изменения через setValue() для каждой ячейки
    if (updates.length > 0) {
      for (const u of updates) {
        sheet.getRange(u.row, u.col).setValue(u.value);
      }
    }
  }
  
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  safeAlert(`✅ Готово!\nПроверено: ${totalScanned}\nЗаменено: ${totalReplaced}\nВремя: ${duration} сек`);
  return { scanned: totalScanned, replaced: totalReplaced };
}

/**
 * Парсер дат для форматов:
 * - ДД.ММ.ГГГГ (13.04.2026)
 * - ДД.ММ.ГГГГ ЧЧ:мм:сс (26.03.2026 01:00:00)
 * - ГГГГ-ММ-ДД, ГГГГ-ММ-ДДТЧЧ:мм:сс (стандарт ISO)
 * - Date-объект из Google Sheets
 */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  
  const str = String(val).trim();
  
  // Формат: ДД.ММ.ГГГГ ЧЧ:мм:сс или ДД.ММ.ГГГГ
  const match = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?$/);
  if (match) {
    const [, d, m, y, h = 0, min = 0, s = 0] = match;
    return new Date(y, m - 1, d, h, min, s);
  }
  
  // Формат ISO: ГГГГ-ММ-ДД или ГГГГ-ММ-ДДТЧЧ:мм:сс
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    const d = new Date(str);
    return isNaN(d) ? null : d;
  }
  
  // Попытка распарсить как есть
  const d = new Date(str);
  return isNaN(d) ? null : d;
}



/**
 * Создаёт/обновляет лист "Склады_проверка"
 * - Берёт склады из "Регион поставки" (столбец 1)
 * - Сохраняет ручные правки в столбце "Объединение" и восстанавливает их
 * - Форматирование: столбец "Объединение" — ширина 150, заголовок по центру, значения слева
 */
function generateWarehouseMasterList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 1️⃣ Читаем склады из "Регион поставки" (столбец 1)
  const regionSheet = ss.getSheetByName("Регион поставки");
  if (!regionSheet) {
    safeAlert('❌ Лист "Регион поставки" не найден');
    return;
  }
  
  const regionData = regionSheet.getDataRange().getValues();
  const warehouseSet = new Set();
  for (let i = 1; i < regionData.length; i++) {
    const wh = regionData[i][0]?.toString().trim();
    if (wh && wh !== '') warehouseSet.add(wh);
  }
  
  if (warehouseSet.size === 0) {
    safeAlert('⚠️ В листе "Регион поставки" не найдены склады');
    return;
  }
  
  // 2️⃣ ЗАПОМИНАЕМ значения "Объединение" перед очисткой
  let savedMerges = {};
  let masterSheet = ss.getSheetByName("Склады_проверка");
  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    if (lastRow > 1) {
      const oldData = masterSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      oldData.forEach(row => {
        const key = row[0]?.toString().trim();
        const mergeVal = row[6];
        if (key && mergeVal !== "" && mergeVal !== null && mergeVal !== undefined) {
          savedMerges[key] = mergeVal;
        }
      });
    }
    masterSheet.clear();
  } else {
    masterSheet = ss.insertSheet("Склады_проверка");
  }
  
  // 3️⃣ Проверяем наличие в исходных листах
  const warehouseMap = {};
  const sources = [
    { sheet: "Финансовые_ежедневные", col: "office_name", key: "Финансовые" },
    { sheet: "Заказы", col: "warehouseName", key: "Заказы" },
    { sheet: "Остатки", col: "Склад", key: "Остатки" },
    { sheet: "Поставки", col: "Склад", key: "Поставки" }
  ];
  
  sources.forEach(src => {
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIdx = findHeaderColIndex_(headers, src.col);
    if (colIdx === -1) return;
    
    const sheetWhSet = new Set();
    for (let i = 1; i < data.length; i++) {
      const wh = data[i][colIdx]?.toString().trim();
      if (wh) sheetWhSet.add(wh);
    }
    
    warehouseSet.forEach(wh => {
      if (!warehouseMap[wh]) warehouseMap[wh] = {};
      warehouseMap[wh][src.key] = sheetWhSet.has(wh);
    });
  });
  
  // 4️⃣ Заголовки
  const headers = ["Склад", "Финансовые", "Заказы", "Остатки", "Поставки", "Регион поставки", "Объединение"];
  masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 5️⃣ Формируем данные
  const warehouses = Array.from(warehouseSet).sort();
  const outputData = warehouses.map(wh => {
    const flags = warehouseMap[wh] || {};
    const mergeValue = savedMerges[wh] || "";
    
    return [
      wh,
      flags.Финансовые ? "да" : "",
      flags.Заказы ? "да" : "",
      flags.Остатки ? "да" : "",
      flags.Поставки ? "да" : "",
      "",
      mergeValue
    ];
  });
  
  if (outputData.length > 0) {
    masterSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  
  // 6️⃣ Форматирование по ТЗ
  const lastDataRow = outputData.length + 1;
  
  // Шапка: жирная, фон, центр для всех
  masterSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#f1f1f1")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  // Столбцы 2-6: ширина 150, центр
  for (let c = 2; c <= 6; c++) {
    masterSheet.setColumnWidth(c, 150);
    masterSheet.getRange(1, c, lastDataRow, 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }
  
  // Столбец 1 "Склад": авто-ширина
  masterSheet.autoResizeColumn(1);
  
  // 🔥 Столбец 7 "Объединение": ширина 150, заголовок по центру, значения слева
  masterSheet.setColumnWidth(7, 150);
  masterSheet.getRange(1, 7).setHorizontalAlignment("center"); // Заголовок
  masterSheet.getRange(2, 7, lastDataRow - 1, 1).setHorizontalAlignment("left"); // Значения
  
  // Границы и шрифт
  masterSheet.getRange(1, 1, lastDataRow, headers.length)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false);
  
  masterSheet.setFrozenRows(1);
  
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  safeAlert(`✅ Обновлено ${warehouses.length} складов за ${duration} сек`);
}


/** Индекс колонки по заголовку (без учёта регистра). -1 если нет. */
function findHeaderColIndex_(headers, name) {
  if (typeof wbColIdx_ === 'function') return wbColIdx_(headers, [name]);
  const needle = name.toString().trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === needle) return i;
  }
  return -1;
}

/** Читает только одну колонку листа (без getDataRange на всю таблицу). */
function getSheetColumnValues_(sheet, headerName) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIdx = findHeaderColIndex_(headers, headerName);
  if (colIdx === -1) return null;
  return sheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
}

/**
 * Какие склады из warehouseSet встречаются в колонке.
 * Останавливается, когда найдены все — не сканирует хвост зря.
 */
function findWarehousesInColumn_(colValues, warehouseSet) {
  const remaining = new Set(warehouseSet);
  const found = new Set();
  if (!colValues) return found;
  for (let i = 0; i < colValues.length; i++) {
    const wh = colValues[i][0] != null ? colValues[i][0].toString().trim() : '';
    if (!wh || !remaining.has(wh)) continue;
    found.add(wh);
    remaining.delete(wh);
    if (remaining.size === 0) break;
  }
  return found;
}

/**
 * Создаёт/обновляет лист "Склады_проверка"
 * - Берёт склады из "Регион поставки" (столбец 1)
 * - Финансовые: "да" только если doc_type_name = "Продажа"
 * - Регион поставки: "да" если регион есть в справочнике
 * - Сохраняет ручные правки в столбце "Объединение"
 */
function generateWarehouseMasterList2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();
  
  // 1️⃣ Читаем справочник регионов и складов из "Регион поставки"
  const regionSheet = ss.getSheetByName("Регион поставки");
  if (!regionSheet) {
    safeAlert('❌ Лист "Регион поставки" не найден');
    return;
  }
  
  const regionData = regionSheet.getDataRange().getValues();
  const warehouseSet = new Set();
  const regionMap = new Map(); // склад -> регион
  const knownRegions = new Set(); // все регионы из справочника
  
  const regionHeaders = regionData[0];
  const whColIdx = findHeaderColIndex_(regionHeaders, 'склад');
  let regColIdx = findHeaderColIndex_(regionHeaders, 'регион');
  if (regColIdx === -1) regColIdx = findHeaderColIndex_(regionHeaders, 'регион поставки');
  
  for (let i = 1; i < regionData.length; i++) {
    const wh = regionData[i][whColIdx] != null ? regionData[i][whColIdx].toString().trim() : '';
    const reg = regColIdx !== -1 && regionData[i][regColIdx] != null
      ? regionData[i][regColIdx].toString().trim() : '';
    
    if (wh) {
      warehouseSet.add(wh);
      if (reg) {
        regionMap.set(wh, reg);
        knownRegions.add(reg);
      }
    }
  }
  
  if (warehouseSet.size === 0) {
    safeAlert('⚠️ В листе "Регион поставки" не найдены склады');
    return;
  }
  
  // 2️⃣ ЗАПОМИНАЕМ значения "Объединение" (без полной очистки листа)
  const savedMerges = {};
  let masterSheet = ss.getSheetByName("Склады_проверка");
  let oldLastRow = 0;
  if (masterSheet) {
    oldLastRow = masterSheet.getLastRow();
    if (oldLastRow > 1) {
      const oldData = masterSheet.getRange(2, 1, oldLastRow - 1, 7).getValues();
      for (let i = 0; i < oldData.length; i++) {
        const key = oldData[i][0] != null ? oldData[i][0].toString().trim() : '';
        const mergeVal = oldData[i][6];
        if (key && mergeVal !== '' && mergeVal != null) savedMerges[key] = mergeVal;
      }
    }
  } else {
    masterSheet = ss.insertSheet("Склады_проверка");
  }
  
  // 3️⃣ Проверяем наличие — читаем только нужные колонки
  const warehouseMap = {};
  warehouseSet.forEach(function(wh) { warehouseMap[wh] = {}; });
  
  // Финансовые: только «Продажа», две колонки вместо всего листа
  const financeSheet = ss.getSheetByName("Финансовые_ежедневные");
  if (financeSheet && financeSheet.getLastRow() > 1) {
    const lastRow = financeSheet.getLastRow();
    const lastCol = financeSheet.getLastColumn();
    const financeHeaders = financeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const whIdx = findHeaderColIndex_(financeHeaders, 'office_name');
    const docIdx = findHeaderColIndex_(financeHeaders, 'doc_type_name');
    if (whIdx !== -1) {
      const whCol = financeSheet.getRange(2, whIdx + 1, lastRow - 1, 1).getValues();
      const docCol = docIdx !== -1
        ? financeSheet.getRange(2, docIdx + 1, lastRow - 1, 1).getValues()
        : null;
      const remaining = new Set(warehouseSet);
      for (let i = 0; i < whCol.length; i++) {
        const wh = whCol[i][0] != null ? whCol[i][0].toString().trim() : '';
        if (!wh || !remaining.has(wh)) continue;
        const docType = docCol ? String(docCol[i][0] || '').trim() : '';
        if (docType !== 'Продажа') continue;
        warehouseMap[wh].Финансовые = true;
        remaining.delete(wh);
        if (remaining.size === 0) break;
      }
    }
  }
  
  const otherSources = [
    { sheet: "Заказы", col: "warehouseName", key: "Заказы" },
    { sheet: "Остатки", col: "Склад", key: "Остатки" },
    { sheet: "Поставки", col: "Склад", key: "Поставки" }
  ];
  for (let s = 0; s < otherSources.length; s++) {
    const src = otherSources[s];
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) continue;
    const found = findWarehousesInColumn_(getSheetColumnValues_(sheet, src.col), warehouseSet);
    found.forEach(function(wh) { warehouseMap[wh][src.key] = true; });
  }
  
  // 4️⃣–5️⃣ Заголовки и данные
  const headers = ["Склад", "Финансовые", "Заказы", "Остатки", "Поставки", "Регион поставки", "Объединение"];
  const warehouses = Array.from(warehouseSet).sort();
  const outputData = warehouses.map(function(wh) {
    const flags = warehouseMap[wh] || {};
    const region = regionMap.get(wh) || "";
    return [
      wh,
      flags.Финансовые ? "да" : "",
      flags.Заказы ? "да" : "",
      flags.Остатки ? "да" : "",
      flags.Поставки ? "да" : "",
      knownRegions.has(region) ? "да" : "",
      savedMerges[wh] || ""
    ];
  });
  
  masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (outputData.length > 0) {
    masterSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  const lastDataRow = outputData.length + 1;
  if (oldLastRow > lastDataRow) {
    masterSheet.getRange(lastDataRow + 1, 1, oldLastRow - lastDataRow, 7).clearContent();
  }
  
  // 6️⃣ Форматирование — пачками, без clear() всего листа
  masterSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#f1f1f1")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  masterSheet.setColumnWidths(2, 5, 150);
  masterSheet.getRange(1, 2, lastDataRow, 5)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  masterSheet.autoResizeColumn(1);
  masterSheet.setColumnWidth(7, 150);
  masterSheet.getRange(1, 7).setHorizontalAlignment("center");
  if (outputData.length > 0) {
    masterSheet.getRange(2, 7, outputData.length, 1).setHorizontalAlignment("left");
  }
  
  masterSheet.getRange(1, 1, lastDataRow, headers.length)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false);
  
  masterSheet.setFrozenRows(1);
  
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);
  Logger.log('✅ Склады_проверка: ' + warehouses.length + ' складов за ' + duration + ' сек');
  safeAlert('✅ Обновлено ' + warehouses.length + ' складов за ' + duration + ' сек');
}

var FILLED_CELLS_REPORT_SHEET = 'Заполненность';

/** Снимок заполненности: дата справа, прирост % и кол-во к предыдущей дате. */
function countFilledCellsOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var stampDate = new Date();
  var todayKey = filledDayKey_(stampDate);

  var current = {};
  var currentNames = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var name = sh.getName();
    if (name === FILLED_CELLS_REPORT_SHEET) continue;
    current[name] = countFilledOnSheet_(sh);
    currentNames.push(name);
  }

  var report = ss.getSheetByName(FILLED_CELLS_REPORT_SHEET);
  if (!report) report = ss.insertSheet(FILLED_CELLS_REPORT_SHEET);
  try { ss.setActiveSheet(report); ss.moveActiveSheet(ss.getSheets().length); } catch (e) {}

  var hist = filledReadHistory_(report);
  var names = hist.names.slice();
  var seen = {};
  for (var n = 0; n < names.length; n++) seen[names[n]] = true;
  for (var c = 0; c < currentNames.length; c++) {
    if (!seen[currentNames[c]]) {
      names.push(currentNames[c]);
      seen[currentNames[c]] = true;
    }
  }

  var overwrite = false;
  if (hist.dates.length) {
    var lastKey = filledDayKey_(hist.dates[hist.dates.length - 1]);
    overwrite = lastKey === todayKey;
  }

  var dateHeaders = [];
  var d;
  for (d = 0; d < hist.dates.length; d++) {
    if (overwrite && d === hist.dates.length - 1) dateHeaders.push(stampDate);
    else dateHeaders.push(filledHeaderDate_(hist.dates[d]));
  }
  if (!overwrite) dateHeaders.push(stampDate);
  var dateCount = dateHeaders.length;

  var output = [['Лист'].concat(dateHeaders, ['прирост %', 'прирост кол-во'])];
  var total = 0;
  for (var r = 0; r < names.length; r++) {
    var sheetName = names[r];
    var prev = hist.grid[sheetName] || [];
    var row = [sheetName];
    var lastHist = overwrite ? hist.dates.length - 1 : hist.dates.length;
    for (d = 0; d < lastHist; d++) {
      var oldVal = prev[d];
      row.push(oldVal === '' || oldVal == null ? 0 : oldVal);
    }
    var nowVal = current.hasOwnProperty(sheetName) ? current[sheetName] : 0;
    var prevVal = 0;
    if (row.length >= 2) prevVal = Number(row[row.length - 1]) || 0;
    row.push(nowVal);
    if (dateCount >= 2) {
      row.push(filledPct_(prevVal, nowVal));
      row.push(nowVal - prevVal);
    } else {
      row.push('', '');
    }
    total += nowVal;
    output.push(row);
  }
  var oldTotals = filledTotals_(hist, names);
  if (overwrite && oldTotals.length) oldTotals = oldTotals.slice(0, oldTotals.length - 1);
  var prevTotal = oldTotals.length ? Number(oldTotals[oldTotals.length - 1]) || 0 : 0;
  if (dateCount >= 2) {
    output.push(['ИТОГО'].concat(oldTotals, [total], [filledPct_(prevTotal, total), total - prevTotal]));
  } else {
    output.push(['ИТОГО'].concat(oldTotals, [total], ['', '']));
  }

  var width = dateCount + 3;
  report.clear();
  report.getRange(1, 1, output.length, width).setValues(output);

  report.setFrozenRows(1);
  report.setFrozenColumns(1);
  report.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#f1f1f1')
    .setHorizontalAlignment('center')
    .setWrap(true);
  report.setRowHeight(1, 36);
  report.setColumnWidth(1, 220);
  if (dateCount >= 1) {
    report.getRange(1, 2, 1, dateCount).setNumberFormat('dd.MM.yyyy');
    report.getRange(2, 2, output.length - 1, dateCount).setNumberFormat('#,##0');
    for (var col = 2; col <= dateCount + 1; col++) report.setColumnWidth(col, 100);
  }
  report.setColumnWidth(dateCount + 2, 95);
  report.setColumnWidth(dateCount + 3, 115);
  if (output.length > 1) {
    report.getRange(2, dateCount + 2, output.length - 1, 1).setNumberFormat('0.00%');
    report.getRange(2, dateCount + 3, output.length - 1, 1).setNumberFormat('+#,##0;-#,##0;0');
  }
  report.getRange(output.length, 1, 1, width).setFontWeight('bold').setBackground('#e8e8e8');
  filledColorGrowth_(report, output.length, dateCount);
  var filter = report.getFilter();
  if (filter) filter.remove();
  report.getRange(1, 1, output.length, width).createFilter();

  var stamp = Utilities.formatDate(stampDate, tz, 'dd.MM.yyyy');
  var msg = 'Заполненность: ' + stamp +
    '\nЛистов сейчас: ' + currentNames.length +
    '\nИтого ячеек: ' + total;
  try { SpreadsheetApp.getUi().alert(msg); } catch (e2) { Logger.log(msg); }
}

function countFilledOnSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return 0;
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var n = 0;
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined) n++;
    }
  }
  return n;
}

function filledHeaderDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof parseDate === 'function') {
    var parsed = parseDate(value);
    if (parsed) return parsed;
  }
  return value;
}

function filledDayKey_(value) {
  var d = filledHeaderDate_(value);
  if (!(d instanceof Date) || isNaN(d.getTime())) return String(value || '');
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Europe/Moscow', 'yyyy-MM-dd');
}

function filledReadHistory_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { names: [], dates: [], grid: {} };

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRow = 0;
  var a0 = String(values[0][0] || '').trim().toLowerCase();
  if (a0 === 'обновлено') headerRow = 1;

  var headers = values[headerRow] || [];
  var dates = [];
  var dateCols = [];
  for (var c = 1; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (!h) continue;
    if (h.indexOf('прирост') === 0) continue;
    dates.push(headers[c]);
    dateCols.push(c);
  }

  var names = [];
  var grid = {};
  for (var r = headerRow + 1; r < values.length; r++) {
    var name = String(values[r][0] || '').trim();
    if (!name) continue;
    var upper = name.toUpperCase();
    if (upper === 'ИТОГО' || upper === 'ИТОГО:') continue;
    names.push(name);
    var rowVals = [];
    for (var d = 0; d < dateCols.length; d++) {
      var v = values[r][dateCols[d]];
      rowVals.push(v === '' || v == null ? 0 : v);
    }
    grid[name] = rowVals;
  }
  return { names: names, dates: dates, grid: grid };
}

function filledPct_(prev, now) {
  prev = Number(prev) || 0;
  now = Number(now) || 0;
  if (prev === 0) return now === 0 ? 0 : '';
  return (now - prev) / prev;
}

function filledTotals_(hist, names) {
  var totals = [];
  for (var d = 0; d < hist.dates.length; d++) {
    var sum = 0;
    for (var i = 0; i < names.length; i++) {
      var row = hist.grid[names[i]] || [];
      var v = Number(row[d]);
      if (!isNaN(v)) sum += v;
    }
    totals.push(sum);
  }
  return totals;
}

function filledColorGrowth_(sheet, rowCount, dateCount) {
  if (dateCount < 2 || rowCount < 3) return;
  var lastCol = dateCount + 1;
  var prevCol = dateCount;
  var pctCol = dateCount + 2;
  var qtyCol = dateCount + 3;
  var dataRows = rowCount - 1;
  if (dataRows < 1) return;

  var lastVals = sheet.getRange(2, lastCol, dataRows, 1).getValues();
  var prevVals = sheet.getRange(2, prevCol, dataRows, 1).getValues();
  var colors = [];
  for (var i = 0; i < dataRows; i++) {
    var now = Number(lastVals[i][0]) || 0;
    var prev = Number(prevVals[i][0]) || 0;
    var color = null;
    if (now === 0 && prev > 0) color = '#f4cccc';
    else if (now > prev) color = '#d9ead3';
    else if (now < prev) color = '#fce5cd';
    colors.push([color]);
  }
  sheet.getRange(2, lastCol, dataRows, 1).setBackgrounds(colors);
  sheet.getRange(2, pctCol, dataRows, 1).setBackgrounds(colors);
  sheet.getRange(2, qtyCol, dataRows, 1).setBackgrounds(colors);
}