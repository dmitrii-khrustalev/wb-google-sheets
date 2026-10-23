// ============================================
// WB API — Остатки на складах продавца (FBS)
// Существующие файлы не меняет. Запускать из редактора скриптов.
//
// Документация:
// POST https://marketplace-api.wildberries.ru/api/v3/stocks/{warehouseId}
// Тело: { "chrtIds": [12345678] }   // до 1000 шт.
// Ответ: { "stocks": [{ "chrtId": 12345678, "amount": 10 }] }
//
// warehouseId — ID склада продавца из GET /api/v3/warehouses
// chrtId — ID размера из POST /content/v2/get/cards/list (поле sizes[].chrtID)
// Параметр skus с 20.05.2026 больше не принимается.
//
// Токен: категория Marketplace (+ Content, чтобы взять chrtId из карточек)
//
// Пишет сырьё в лёгкий файл «WB FBS Заказы», затем копирует листы в активную
// (главную) книгу. SpreadsheetApp напрямую в главную даёт timed out;
// Sheets REST API требует включения в Google Cloud — не используем.
// Рабочие сводки — «FBS Поставки» / «FBS Заказы» / «FBS Остатки», чтобы
// не затереть «Поставки» / «Заказы» / «Остатки» из файла 01.
//
// Функции:
//   loadSellerWarehouseStocks   — остатки + справочники FBS_Склады и FBS_Офисы
// ============================================

var FBS_STOCKS_SHEET = 'FBS_Остатки';
var WB_FBS_SIDECAR_ID = '1JkY8YrJnlI6HcZ0DbiDIvH_vIIeYmJHutAlJKTzLb6o';
var WB_FBS_SIDECAR_PROP = 'WB_FBS_ORDERS_SS_ID';
var FBS_STOCKS_HEADERS = [
  'Дата', 'warehouseId', 'warehouseName', 'officeId', 'deliveryType', 'deliveryType_name',
  'cargoType', 'cargoType_name', 'chrtId', 'amount', 'sku',
  'nmID', 'vendorCode', 'title', 'brand', 'subjectName', 'techSize', 'wbSize', 'skus'
];

var FBS_MP_HOST = 'https://marketplace-api.wildberries.ru';
var FBS_STOCKS_BATCH = 1000;
var FBS_REQUEST_PAUSE_MS = 250;

var FBS_DELIVERY_TYPE = {
  1: 'FBS — со склада продавца',
  2: 'DBS — доставка силами продавца',
  3: 'DBW — курьером WB',
  5: 'Самовывоз (Click & Collect)',
  6: 'EDBS — экспресс силами продавца'
};

var FBS_CARGO_TYPE = {
  1: 'МГТ',
  2: 'СГТ',
  3: 'КГТ+'
};

// --------------------------------------------
// ПОЛНАЯ ВЫГРУЗКА
// --------------------------------------------
function loadSellerWarehouseStocks(opts) {
  var daily = wbDaily_(opts);
  if (!daily && typeof renameWbHeadersAllBooks_ === 'function') renameWbHeadersAllBooks_();
  var token = getApiToken();
  if (!token) return;

  var dateStr = fbsToday_();
  var log = [];
  log.push(['Шаг', 'Результат']);
  log.push(['Дата снимка', dateStr]);

  var whRes = fbsFetchJson_(token, 'GET', FBS_MP_HOST + '/api/v3/warehouses', null);
  var warehouses = fbsExtractArray_(whRes.data, 'warehouses');
  log.push(['GET /api/v3/warehouses', whRes.code + ' | ' + warehouses.length]);
  if (warehouses.length) {
    fbsWriteSheetSafe_('FBS_Склады', fbsWarehousesToRows_(warehouses), log);
  }

  if (!daily) {
    var officesRes = fbsFetchJson_(token, 'GET', FBS_MP_HOST + '/api/v3/offices', null);
    var offices = fbsExtractArray_(officesRes.data, 'offices');
    log.push(['GET /api/v3/offices', officesRes.code + ' | ' + offices.length]);
    if (offices.length) {
      fbsWriteSheetSafe_('FBS_Офисы', fbsOfficesToRows_(offices), log);
    }
  } else {
    log.push(['GET /api/v3/offices', 'пропущен (ежедневный режим)']);
  }

  if (!warehouses.length) {
    log.push(['Остатки', 'Нет складов продавца — выгружать нечего']);
    var urlNoWh = fbsDumpDone_(log);
    fbsNotify_('Складов продавца нет', urlNoWh);
    return;
  }

  var sizes = [];
  if (daily) {
    sizes = fbsSizesFromStocksSheet_();
    log.push(['chrtId из прошлого снимка', String(sizes.length)]);
  }
  if (!sizes.length) {
    var cards = fbsFetchAllCards_(token, log);
    sizes = fbsCollectSizes_(cards);
    log.push(['Карточек / размеров', (cards ? cards.length : 0) + ' / ' + sizes.length]);
  }

  if (!sizes.length) {
    log.push(['Остатки', 'Нет chrtId в карточках']);
    var urlNoChrt = fbsDumpDone_(log);
    fbsNotify_('В карточках нет chrtId. Проверьте категорию Content у токена.', urlNoChrt);
    return;
  }

  var sizeMap = fbsSizeMap_(sizes);
  var chrtIds = sizes.map(function(s) { return s.chrtId; });
  var dataRows = [];
  var withQty = 0;

  for (var w = 0; w < warehouses.length; w++) {
    var wh = warehouses[w];
    var warehouseId = wh.id || wh.ID;
    var items = fbsFetchStocksAll_(token, warehouseId, chrtIds, log);
    var part = fbsStocksToRows_(warehouseId, wh, items, sizeMap, dateStr);
    for (var i = 1; i < part.length; i++) {
      dataRows.push(part[i]);
      if (Number(part[i][9]) > 0) withQty++;
    }
  }

  var added = fbsAppendStocksSheet_(dataRows, log);
  log.push(['Строк от API', String(dataRows.length)]);
  log.push(['Из них amount > 0', String(withQty)]);
  log.push(['Добавлено на лист', String(added)]);
  var url = fbsDumpDone_(log);
  if (!daily && !(opts && opts.skipCompose) && typeof composeFbsReports === 'function') {
    try { url = composeFbsReports(true) || url; } catch (e) { log.push(['Сборка рабочих листов', String(e)]); }
  }
  fbsNotify_(
    'Снимок ' + dateStr + ': с API ' + dataRows.length + ' строк, на лист ' + added +
    ' (день перезаписан)',
    url
  );
}

// --------------------------------------------
// HTTP
// --------------------------------------------
function fbsFetchJson_(token, method, url, body) {
  var options = {
    method: method,
    headers: { 'Authorization': token, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);

  for (var attempt = 0; attempt < 4; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var text = response.getContentText() || '';
      if (code === 429) {
        Utilities.sleep(6000);
        continue;
      }
      var data = null;
      if (text) {
        try { data = JSON.parse(text); } catch (e) { data = null; }
      }
      return { code: code, data: data, text: text };
    } catch (e) {
      if (attempt === 3) return { code: 0, data: null, text: String(e) };
      Utilities.sleep(2000);
    }
  }
  return { code: 429, data: null, text: 'Too many requests' };
}

function fbsFetchAllCards_(token, log) {
  var all = [];
  var cursor = null;
  var hasMore = true;
  var pages = 0;

  while (hasMore) {
    pages++;
    var payload = {
      settings: {
        cursor: { limit: 100 },
        filter: { withPhoto: -1 }
      }
    };
    if (cursor) {
      payload.settings.cursor.updatedAt = cursor.updatedAt;
      payload.settings.cursor.nmID = cursor.nmID;
    }

    var res = fbsFetchJson_(token, 'POST', 'https://content-api.wildberries.ru/content/v2/get/cards/list', payload);
    if (res.code !== 200) {
      if (log) log.push(['POST /content/v2/get/cards/list', res.code + ' | ' + (res.text || '').substring(0, 400)]);
      break;
    }

    var cards = (res.data && res.data.cards) ? res.data.cards : [];
    if (cards.length) all = all.concat(cards);

    if (res.data && res.data.cursor && res.data.cursor.updatedAt && res.data.cursor.nmID && cards.length === 100) {
      cursor = { updatedAt: res.data.cursor.updatedAt, nmID: res.data.cursor.nmID };
      Utilities.sleep(FBS_REQUEST_PAUSE_MS);
    } else {
      hasMore = false;
    }
  }

  if (log) log.push(['Страниц карточек', String(pages)]);
  return all;
}

function fbsFetchStocksAll_(token, warehouseId, chrtIds, log) {
  var items = [];
  for (var i = 0; i < chrtIds.length; i += FBS_STOCKS_BATCH) {
    var batch = chrtIds.slice(i, i + FBS_STOCKS_BATCH);
    var res = fbsFetchJson_(
      token,
      'POST',
      FBS_MP_HOST + '/api/v3/stocks/' + warehouseId,
      { chrtIds: batch }
    );
    if (log) log.push(['POST /api/v3/stocks/' + warehouseId + ' batch ' + (i / FBS_STOCKS_BATCH + 1), res.code + ' | chrtIds=' + batch.length]);
    if (res.code !== 200) {
      if (log && res.text) log.push(['Ошибка batch', res.text.substring(0, 800)]);
      continue;
    }
    items = items.concat(fbsExtractStocks_(res.data));
    Utilities.sleep(FBS_REQUEST_PAUSE_MS);
  }
  return items;
}

function fbsExtractStocks_(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (Array.isArray(data.stocks)) return data.stocks;
  return [];
}

function fbsExtractArray_(data, key) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (key && Array.isArray(data[key])) return data[key];
  return [];
}

// --------------------------------------------
// Карточки → размеры / chrtId
// --------------------------------------------
function fbsCollectSizes_(cards) {
  var out = [];
  var seen = {};
  for (var i = 0; i < cards.length; i++) {
    var card = cards[i];
    var sizes = card.sizes || [];
    for (var s = 0; s < sizes.length; s++) {
      var size = sizes[s];
      var chrtId = size.chrtID || size.chrtId || size.chrt_id;
      if (!chrtId || seen[chrtId]) continue;
      seen[chrtId] = true;
      out.push({
        chrtId: chrtId,
        nmID: card.nmID || card.nmId || '',
        vendorCode: card.vendorCode || '',
        title: card.title || '',
        brand: card.brand || '',
        subjectName: card.subjectName || '',
        techSize: size.techSize || '',
        wbSize: size.wbSize || '',
        skus: (size.skus || []).join(', '),
        sizeRaw: size
      });
    }
  }
  return out;
}

function fbsSizeMap_(sizes) {
  var map = {};
  for (var i = 0; i < sizes.length; i++) map[sizes[i].chrtId] = sizes[i];
  return map;
}

function fbsSizesFromStocksSheet_() {
  var values = wbReadSheet_(fbsBook_(), ['FBS_Остатки']);
  if (!values || values.length < 2) return [];

  var headers = values[0];
  var idx = {};
  for (var h = 0; h < headers.length; h++) idx[String(headers[h] || '')] = h;

  function col(row, name) {
    var i = typeof wbColIdx_ === 'function' ? wbColIdx_(headers, [name]) : idx[name];
    if (i === undefined || i === '' || i < 0) return '';
    return row[i] === undefined || row[i] === null ? '' : row[i];
  }

  var out = [];
  var seen = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var chrtId = col(row, 'chrtId') || col(row, 'chrtID');
    if (!chrtId || seen[chrtId]) continue;
    seen[chrtId] = true;
    out.push({
      chrtId: chrtId,
      nmID: col(row, 'nmID') || col(row, 'nmId'),
      vendorCode: col(row, 'vendorCode'),
      title: col(row, 'title'),
      brand: col(row, 'brand'),
      subjectName: col(row, 'subjectName'),
      techSize: col(row, 'techSize'),
      wbSize: col(row, 'wbSize'),
      skus: col(row, 'skus')
    });
  }
  return out;
}

// --------------------------------------------
// Строки для листов
// --------------------------------------------
function fbsWarehousesToRows_(warehouses) {
  var extraKeys = fbsExtraKeys_(warehouses, ['id', 'ID', 'name', 'officeId', 'cargoType', 'deliveryType', 'isDeleting', 'isProcessing']);
  var header = ['id', 'name', 'officeId', 'cargoType', 'cargoType_name', 'deliveryType', 'deliveryType_name', 'isDeleting', 'isProcessing'].concat(extraKeys);
  var rows = [header];
  for (var i = 0; i < warehouses.length; i++) {
    var w = warehouses[i];
    var row = [
      w.id || w.ID || '',
      w.name || '',
      w.officeId || '',
      w.cargoType !== undefined ? w.cargoType : '',
      FBS_CARGO_TYPE[w.cargoType] || '',
      w.deliveryType !== undefined ? w.deliveryType : '',
      FBS_DELIVERY_TYPE[w.deliveryType] || '',
      w.isDeleting === true ? 'Да' : (w.isDeleting === false ? 'Нет' : ''),
      w.isProcessing === true ? 'Да' : (w.isProcessing === false ? 'Нет' : '')
    ];
    for (var k = 0; k < extraKeys.length; k++) row.push(fbsScalar_(w[extraKeys[k]]));
    rows.push(row);
  }
  return rows;
}

function fbsOfficesToRows_(offices) {
  if (!Array.isArray(offices) || !offices.length) return [['(пусто)']];
  var extraKeys = fbsExtraKeys_(offices, ['id', 'name', 'address', 'city', 'federalDistrict', 'cargoType', 'deliveryType', 'selected', 'longitude', 'latitude']);
  var header = ['id', 'name', 'address', 'city', 'federalDistrict', 'cargoType', 'cargoType_name', 'deliveryType', 'deliveryType_name', 'selected', 'longitude', 'latitude'].concat(extraKeys);
  var rows = [header];
  for (var i = 0; i < offices.length; i++) {
    var o = offices[i];
    var row = [
      o.id || '',
      o.name || '',
      o.address || '',
      o.city || '',
      o.federalDistrict || '',
      o.cargoType !== undefined ? o.cargoType : '',
      FBS_CARGO_TYPE[o.cargoType] || '',
      o.deliveryType !== undefined ? o.deliveryType : '',
      FBS_DELIVERY_TYPE[o.deliveryType] || '',
      o.selected === true ? 'Да' : (o.selected === false ? 'Нет' : ''),
      o.longitude || '',
      o.latitude || ''
    ];
    for (var k = 0; k < extraKeys.length; k++) row.push(fbsScalar_(o[extraKeys[k]]));
    rows.push(row);
  }
  return rows;
}

function fbsSizesToRows_(sizes) {
  var rows = [['chrtId', 'nmID', 'vendorCode', 'title', 'brand', 'subjectName', 'techSize', 'wbSize', 'skus']];
  for (var i = 0; i < sizes.length; i++) {
    var s = sizes[i];
    rows.push([s.chrtId, s.nmID, s.vendorCode, s.title, s.brand, s.subjectName, s.techSize, s.wbSize, s.skus]);
  }
  return rows;
}

function fbsStocksToRows_(warehouseId, warehouse, items, sizeMap, dateStr) {
  var extraKeys = fbsExtraKeys_(items, ['chrtId', 'chrtID', 'amount', 'sku', 'skus']);
  var header = FBS_STOCKS_HEADERS.concat(extraKeys);
  var rows = [header];

  for (var i = 0; i < items.length; i++) {
    var st = items[i];
    var chrtId = st.chrtId || st.chrtID || '';
    var info = sizeMap[chrtId] || {};
    var row = [
      dateStr || fbsToday_(),
      warehouseId,
      (warehouse && warehouse.name) || '',
      (warehouse && warehouse.officeId) || '',
      warehouse && warehouse.deliveryType !== undefined ? warehouse.deliveryType : '',
      (warehouse && FBS_DELIVERY_TYPE[warehouse.deliveryType]) || '',
      warehouse && warehouse.cargoType !== undefined ? warehouse.cargoType : '',
      (warehouse && FBS_CARGO_TYPE[warehouse.cargoType]) || '',
      chrtId,
      st.amount !== undefined && st.amount !== null ? st.amount : '',
      st.sku || '',
      info.nmID || '',
      info.vendorCode || '',
      info.title || '',
      info.brand || '',
      info.subjectName || '',
      info.techSize || '',
      info.wbSize || '',
      info.skus || ''
    ];
    for (var k = 0; k < extraKeys.length; k++) row.push(fbsScalar_(st[extraKeys[k]]));
    rows.push(row);
  }
  return rows;
}

function wbDaily_(opts) {
  return opts === true || (opts && opts.daily === true);
}

function wbReuseSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    wbClearFilter_(sheet);
    sheet.clear();
    return sheet;
  }
  return ss.insertSheet(name);
}

function fbsSidecar_() {
  var props = PropertiesService.getUserProperties();
  var ids = [props.getProperty(WB_FBS_SIDECAR_PROP), WB_FBS_SIDECAR_ID];
  for (var i = 0; i < ids.length; i++) {
    if (!ids[i]) continue;
    try {
      var ss = SpreadsheetApp.openById(ids[i]);
      props.setProperty(WB_FBS_SIDECAR_PROP, ss.getId());
      return ss;
    } catch (e) {
      Logger.log('fbsSidecar_ ' + ids[i] + ': ' + e);
    }
  }
  var created = SpreadsheetApp.create('WB FBS сырьё');
  props.setProperty(WB_FBS_SIDECAR_PROP, created.getId());
  return created;
}

function wbReadSheetApp_(ss, names) {
  if (!ss) return [];
  if (typeof names === 'string') names = [names];
  for (var i = 0; i < names.length; i++) {
    try {
      var sheet = ss.getSheetByName(names[i]);
      if (sheet && sheet.getLastRow() > 0) return sheet.getDataRange().getValues();
    } catch (e) {
      Logger.log('wbReadSheetApp_ ' + names[i] + ': ' + e);
    }
  }
  return [];
}

function wbReadSheet_(ss, names) {
  if (typeof names === 'string') names = [names];
  var fromLight = wbReadSheetApp_(fbsSidecar_(), names);
  if (fromLight.length) return fromLight;
  return wbReadSheetApp_(ss, names);
}

function wbWriteSheetApp_(ss, name, rows) {
  rows = wbFormatSheetDates_(rows);
  var sheet = wbReuseSheet_(ss, name);
  wbClearFilter_(sheet);

  var width = 1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i] && rows[i].length > width) width = rows[i].length;
  }

  var CHUNK = 500;
  for (var start = 0; start < rows.length; start += CHUNK) {
    var end = Math.min(start + CHUNK, rows.length);
    var batch = [];
    for (var r = start; r < end; r++) {
      var row = rows[r] ? rows[r].slice() : [];
      while (row.length < width) row.push('');
      batch.push(row);
    }
    sheet.getRange(start + 1, 1, batch.length, width).setValues(batch);
  }
  sheet.setFrozenRows(1);
  wbEnsureFilter_(sheet);
}

function wbClearFilter_(sheet) {
  if (!sheet) return;
  try {
    var filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {
    Logger.log('wbClearFilter_ ' + sheet.getName() + ': ' + e);
  }
}

function wbEnsureFilter_(sheet) {
  if (!sheet) return;
  try {
    wbClearFilter_(sheet);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  } catch (e) {
    Logger.log('filter ' + sheet.getName() + ': ' + e);
  }
}

function wbEnsureFbsFilters_(ss, log) {
  if (!ss) return;
  var lightId = '';
  try { lightId = fbsSidecar_().getId(); } catch (e1) {}
  var allOnThis = ss.getId() === lightId;
  var sheets = ss.getSheets();
  var n = 0;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName() || '';
    if (!allOnThis && name.indexOf('FBS') !== 0) continue;
    wbEnsureFilter_(sheets[i]);
    n++;
  }
  if (log) log.push(['Фильтры ' + (allOnThis ? 'сырой файл' : 'главная'), String(n)]);
}

function wbCopySheetTo_(srcSs, name, destSs, log) {
  if (!srcSs || !destSs || srcSs.getId() === destSs.getId()) return;
  var src = srcSs.getSheetByName(name);
  if (!src) return;
  try {
    var old = destSs.getSheetByName(name);
    var keepIndex = 0;
    var keepColor = null;
    var keepHidden = false;
    if (old) {
      keepIndex = old.getIndex();
      try { keepColor = old.getTabColor(); } catch (e1) {}
      try { keepHidden = old.isSheetHidden(); } catch (e2) {}
    }
    var copied = src.copyTo(destSs);
    if (old && old.getSheetId() !== copied.getSheetId()) destSs.deleteSheet(old);
    copied.setName(name);
    if (keepColor) {
      try { copied.setTabColor(keepColor); } catch (e3) {}
    }
    if (keepIndex > 0) {
      try {
        destSs.setActiveSheet(copied);
        destSs.moveActiveSheet(keepIndex);
      } catch (e4) {}
    }
    if (keepHidden) {
      try { copied.hideSheet(); } catch (e5) {}
    }
    wbEnsureFilter_(copied);
    if (log) log.push(['Скопирован в главную', name]);
  } catch (e) {
    Logger.log('copyTo ' + name + ': ' + e);
    if (log) log.push(['Не скопировали в главную ' + name, String(e) + ' | сырьё: ' + srcSs.getUrl()]);
  }
}

function wbWriteSheetSafe_(ss, name, rows, log) {
  if (!rows || !rows.length) return;
  if (typeof wbMapHeaders_ === 'function') {
    rows = [wbMapHeaders_(rows[0], name)].concat(rows.slice(1));
  }
  var light = fbsSidecar_();
  var lastErr = null;
  for (var attempt = 0; attempt < 3; attempt++) {
    try {
      if (attempt) Utilities.sleep(3000);
      wbWriteSheetApp_(light, name, rows);
      if (log) log.push(['Записан лист ' + name, String(rows.length) + ' строк']);
      wbCopySheetTo_(light, name, ss, log);
      return;
    } catch (e) {
      lastErr = e;
      Logger.log('wbWriteSheet_ ' + name + ' попытка ' + (attempt + 1) + ': ' + e);
    }
  }
  if (log) log.push(['Ошибка записи ' + name, String(lastErr) + ' | сырьё: ' + light.getUrl()]);
  throw lastErr;
}

function wbHideSheets_(ss, names) {
  var hidden = [];
  for (var i = 0; i < names.length; i++) {
    try {
      var sheet = ss.getSheetByName(names[i]);
      if (!sheet) continue;
      sheet.hideSheet();
      hidden.push(names[i]);
    } catch (e) {
      Logger.log('hide ' + names[i] + ': ' + e);
    }
  }
  return hidden;
}

function wbDeleteSheets_(ss, names) {
  var dropped = [];
  for (var i = 0; i < names.length; i++) {
    try {
      var sheet = ss.getSheetByName(names[i]);
      if (!sheet) continue;
      if (ss.getSheets().length <= 1) break;
      ss.deleteSheet(sheet);
      dropped.push(names[i]);
    } catch (e) {
      Logger.log('delete ' + names[i] + ': ' + e);
    }
  }
  return dropped;
}

function wbMergeById_(oldTable, newTable, idKeys, sheetName) {
  oldTable = oldTable || [];
  newTable = newTable || [];
  if (typeof wbMapHeaders_ === 'function' && sheetName) {
    if (oldTable.length) oldTable = [wbMapHeaders_(oldTable[0], sheetName)].concat(oldTable.slice(1));
    if (newTable.length) newTable = [wbMapHeaders_(newTable[0], sheetName)].concat(newTable.slice(1));
  }
  var oldH = oldTable.length ? oldTable[0] : [];
  var newH = newTable.length ? newTable[0] : [];
  var headerMap = {};
  var headers = [];
  function addHeader(h) {
    var s = String(h || '').trim();
    if (!s) return;
    if (typeof WB_RU !== 'undefined' && WB_RU[s]) s = WB_RU[s];
    for (var i2 = 0; i2 < headers.length; i2++) {
      if (String(headers[i2]).toLowerCase() === s.toLowerCase()) return;
      if (typeof wbIs_ === 'function' && wbIs_(headers[i2], s)) return;
    }
    headerMap[s] = headers.length;
    headers.push(s);
  }
  var i;
  for (i = 0; i < newH.length; i++) addHeader(newH[i]);
  for (i = 0; i < oldH.length; i++) addHeader(oldH[i]);
  if (!headers.length) return [];

  function toObj(row, h) {
    var o = {};
    for (var c = 0; c < h.length; c++) {
      if (h[c]) o[String(h[c])] = row[c];
    }
    return o;
  }

  function rowId(obj) {
    var keys = idKeys.slice();
    if (keys.indexOf('id') >= 0) {
      keys = keys.concat(['ID задания', 'ID поставки', 'ID', 'ID поставки FBS', 'orderId']);
    }
    if (typeof wbPickObj_ === 'function') {
      var picked = wbPickObj_(obj, keys);
      if (picked !== '') return String(picked);
    }
    for (var k = 0; k < keys.length; k++) {
      var v = obj[keys[k]];
      if (v !== undefined && v !== null && v !== '') return String(v);
    }
    return '';
  }

  var byId = {};
  var order = [];
  function upsert(row, h, incomingWins) {
    var obj = toObj(row, h);
    var id = rowId(obj);
    if (!id) return;
    if (!byId[id]) {
      byId[id] = obj;
      order.push(id);
      return;
    }
    var prev = byId[id];
    var keys = Object.keys(obj);
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var v = obj[key];
      if (v === undefined || v === null || v === '') continue;
      if (incomingWins || prev[key] === undefined || prev[key] === null || prev[key] === '') prev[key] = v;
    }
  }

  for (i = 1; i < oldTable.length; i++) upsert(oldTable[i], oldH, false);
  for (i = 1; i < newTable.length; i++) upsert(newTable[i], newH, true);

  var out = [headers];
  for (i = 0; i < order.length; i++) {
    var rec = byId[order[i]];
    var row = [];
    for (var c = 0; c < headers.length; c++) {
      var val = typeof wbCell_ === 'function' ? wbCell_(rec, headers[c]) : rec[headers[c]];
      row.push(val === undefined || val === null ? '' : val);
    }
    out.push(row);
  }
  return out;
}

function wbMergePairRows_(oldTable, newTable, colA, colB) {
  var seen = {};
  var out = [[colA, colB]];
  function add(table) {
    if (!table || table.length < 2) return;
    var h = table[0];
    var ia = -1;
    var ib = -1;
    for (var i = 0; i < h.length; i++) {
      if (typeof wbIs_ === 'function') {
        if (wbIs_(h[i], colA)) ia = i;
        if (wbIs_(h[i], colB)) ib = i;
      } else {
        if (String(h[i]) === colA) ia = i;
        if (String(h[i]) === colB) ib = i;
      }
    }
    if (ia < 0 || ib < 0) return;
    for (var r = 1; r < table.length; r++) {
      var ka = String(table[r][ia] || '');
      var kb = String(table[r][ib] || '');
      if (!ka || !kb) continue;
      var key = ka + '|' + kb;
      if (seen[key]) continue;
      seen[key] = true;
      out.push([ka, kb]);
    }
  }
  add(oldTable);
  add(newTable);
  return out;
}

function wbTz_() {
  try {
    return Session.getScriptTimeZone() || 'Europe/Moscow';
  } catch (e) {
    return 'Europe/Moscow';
  }
}

function wbParseDate_(value) {
  if (value === undefined || value === null || value === '') return null;
  if (typeof value === 'number' || typeof value === 'boolean') return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  var s = String(value).trim();
  if (!s || s === 'null' || s === 'undefined') return null;

  var dmy = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (dmy) {
    return new Date(
      Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]),
      Number(dmy[4] || 0), Number(dmy[5] || 0), Number(dmy[6] || 0)
    );
  }

  var ymd = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));

  if (!/^\d{4}-\d{2}-\d{2}/.test(s) && !/^\d{1,2}\.\d{1,2}\.\d{4}/.test(s)) return null;
  var ms = Date.parse(s);
  return isNaN(ms) ? null : new Date(ms);
}

function wbHasTime_(value) {
  if (value instanceof Date) {
    return value.getHours() !== 0 || value.getMinutes() !== 0 || value.getSeconds() !== 0;
  }
  var s = String(value || '');
  return /T\d{1,2}:/.test(s) || / \d{1,2}:\d{2}/.test(s);
}

function wbDate_(value) {
  var d = wbParseDate_(value);
  if (!d) return value === 0 || value === false ? value : (value || '');
  return Utilities.formatDate(d, wbTz_(), 'dd.MM.yyyy');
}

function wbDateTime_(value) {
  var d = wbParseDate_(value);
  if (!d) return value === 0 || value === false ? value : (value || '');
  return Utilities.formatDate(d, wbTz_(), 'dd.MM.yyyy HH:mm');
}

function wbMaybeDate_(value) {
  var d = wbParseDate_(value);
  if (!d) return value === 0 || value === false ? value : (value || '');
  return wbHasTime_(value) ? wbDateTime_(value) : wbDate_(value);
}

function wbDateKey_(value) {
  var d = wbParseDate_(value);
  if (!d) return String(value || '');
  return Utilities.formatDate(d, wbTz_(), 'yyyy-MM-dd');
}

function wbDateColKind_(name) {
  var h = String(name || '').toLowerCase().replace(/\s+/g, '');
  if (!h) return '';
  if (h === 'дата' || h === 'date' || h.indexOf('дата') >= 0) return 'date';
  if (typeof wbIs_ === 'function') {
    if (wbIs_(name, 'createdAt') || wbIs_(name, 'closedAt') || wbIs_(name, 'scanDt') ||
        wbIs_(name, 'rejectDt') || wbIs_(name, 'sellerDate') || wbIs_(name, 'supplyCreatedAt') ||
        wbIs_(name, 'orderCreatedAt') || wbIs_(name, 'supplyClosedAt')) return 'datetime';
  }
  if (/createdat|closedat|scandt|rejectdt|sellerdate|deliverydate|status_date|statusdate|ordercreatedat|supplycreatedat|supplyclosedat/.test(h)) {
    return 'datetime';
  }
  if (/at$/.test(h) && h.length > 3) return 'datetime';
  if (/dt$/.test(h) && h.length > 2) return 'datetime';
  if (h.indexOf('date') >= 0) return 'maybe';
  return '';
}

function wbFormatSheetDates_(rows) {
  if (!rows || rows.length < 2) return rows;
  var headers = rows[0];
  var kinds = [];
  var any = false;
  for (var c = 0; c < headers.length; c++) {
    kinds[c] = wbDateColKind_(headers[c]);
    if (kinds[c]) any = true;
  }
  if (!any) return rows;

  var out = [headers];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r] ? rows[r].slice() : [];
    for (var i = 0; i < kinds.length; i++) {
      if (!kinds[i] || row[i] === undefined || row[i] === null || row[i] === '') continue;
      if (kinds[i] === 'date') row[i] = wbDate_(row[i]);
      else if (kinds[i] === 'datetime') row[i] = wbDateTime_(row[i]);
      else row[i] = wbMaybeDate_(row[i]);
    }
    out.push(row);
  }
  return out;
}

function fbsToday_() {
  return wbDate_(new Date());
}

function fbsNormDate_(value) {
  return wbDateKey_(value);
}

function fbsStockKey_(dateVal, warehouseId, chrtId) {
  return fbsNormDate_(dateVal) + '|' + String(warehouseId || '') + '|' + String(chrtId || '');
}

function fbsEnsureDateColumnValues_(values) {
  if (!values || !values.length) return [FBS_STOCKS_HEADERS.slice()];
  var firstHeader = String(values[0][0] || '').toLowerCase();
  if (firstHeader === 'дата' || firstHeader === 'date') return values;
  var today = fbsToday_();
  var out = [values[0].slice()];
  out[0].unshift('Дата');
  for (var i = 1; i < values.length; i++) {
    var row = (values[i] || []).slice();
    row.unshift(today);
    out.push(row);
  }
  return out;
}

function fbsAppendStocksSheet_(dataRows, log) {
  if (!dataRows || !dataRows.length) return 0;
  var dest = fbsBook_();
  var light = fbsSidecar_();
  var raw = wbReadSheet_(dest, [FBS_STOCKS_SHEET]);
  var needRewriteHeader = raw.length && String((raw[0] || [])[0] || '').toLowerCase() !== 'дата' &&
    String((raw[0] || [])[0] || '').toLowerCase() !== 'date';
  var values = fbsEnsureDateColumnValues_(raw);
  var headers = values[0];
  var dateIdx = 0;
  var whIdx = 1;
  var chrtIdx = 8;
  for (var h = 0; h < headers.length; h++) {
    var name = String(headers[h] || '');
    if (typeof wbIs_ === 'function') {
      if (wbIs_(name, 'warehouseId')) whIdx = h;
      if (wbIs_(name, 'chrtId')) chrtIdx = h;
    } else {
      if (name === 'warehouseId') whIdx = h;
      if (name === 'chrtId') chrtIdx = h;
    }
  }

  var existing = {};
  var snapDate = dataRows.length ? fbsNormDate_(dataRows[0][0]) : '';
  var kept = [headers];
  var removed = 0;
  for (var i = 1; i < values.length; i++) {
    if (snapDate && fbsNormDate_(values[i][dateIdx]) === snapDate) {
      removed++;
      continue;
    }
    kept.push(values[i]);
  }
  values = kept;
  if (log && removed) log.push(['Снято за ' + snapDate, String(removed)]);

  var colCount = Math.max(headers.length, FBS_STOCKS_HEADERS.length);
  var newRows = [];
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r].slice();
    var key = fbsStockKey_(row[0], row[1], row[8]);
    if (existing[key]) continue;
    existing[key] = true;
    while (row.length < colCount) row.push('');
    if (row.length > colCount) row = row.slice(0, colCount);
    newRows.push(row);
  }

  if (needRewriteHeader || values.length <= 1 || removed) {
    var all = (values.length <= 1 ? [headers] : values).concat(newRows);
    wbWriteSheetSafe_(dest, FBS_STOCKS_SHEET, all, log);
    if (log) log.push(['Запись', newRows.length ? 'перезаписан день, строк ' + newRows.length : 'снят день, новых нет']);
    return newRows.length;
  }

  if (!newRows.length) {
    if (log) log.push(['Дозапись', 'нет строк с API']);
    return 0;
  }

  var sheet = light.getSheetByName(FBS_STOCKS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    wbWriteSheetApp_(light, FBS_STOCKS_SHEET, values.concat(newRows));
  } else {
    wbClearFilter_(sheet);
    var start = sheet.getLastRow() + 1;
    for (var b = 0; b < newRows.length; b += 1000) {
      var batch = newRows.slice(b, b + 1000);
      sheet.getRange(start + b, 1, batch.length, colCount).setValues(batch);
    }
    wbEnsureFilter_(sheet);
  }
  wbCopySheetTo_(light, FBS_STOCKS_SHEET, dest, log);
  if (log) log.push(['Дозапись', 'добавлено ' + newRows.length]);
  return newRows.length;
}

// --------------------------------------------
// Утилиты
// --------------------------------------------
function fbsBook_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function fbsDumpDone_(log) {
  var ss = fbsBook_();
  var url = ss.getUrl();
  if (log) {
    log.push(['Главная книга', url]);
    try { log.push(['Сырой файл FBS', fbsSidecar_().getUrl()]); } catch (e) {}
  }
  return url;
}

function fbsNotify_(msg, url) {
  var text = msg + (url ? '\n\nОткройте файл:\n' + url : '');
  Logger.log(text);
  if (typeof tgNote_ === 'function') tgNote_(msg);
  if (typeof WB_SILENT_UI !== 'undefined' && WB_SILENT_UI) return;
  try {
    SpreadsheetApp.getUi().alert(text);
  } catch (e) {
    Logger.log('UI недоступен: ' + e);
  }
}

function fbsWriteSheetSafe_(name, rows, log) {
  wbWriteSheetSafe_(fbsBook_(), name, rows, log);
}

function fbsWriteSheet_(name, rows) {
  wbWriteSheetSafe_(fbsBook_(), name, rows, null);
}

function fbsUniqueKeys_(arr) {
  var keys = {};
  if (!Array.isArray(arr)) return [];
  for (var i = 0; i < arr.length; i++) {
    var obj = arr[i];
    if (!obj || typeof obj !== 'object') continue;
    var k = Object.keys(obj);
    for (var j = 0; j < k.length; j++) keys[k[j]] = true;
  }
  return Object.keys(keys);
}

function fbsExtraKeys_(arr, known) {
  var skip = {};
  for (var i = 0; i < known.length; i++) skip[known[i]] = true;
  return fbsUniqueKeys_(arr).filter(function(k) { return !skip[k]; });
}

function fbsScalar_(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return wbDateTime_(value);
  if (typeof value === 'object') return JSON.stringify(value);
  var parsed = wbParseDate_(value);
  if (parsed) return wbMaybeDate_(value);
  return value;
}
