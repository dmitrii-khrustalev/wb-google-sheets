// ============================================
// WB API — Сборочные задания FBS (Marketplace)
// Существующие файлы не меняет. Запускать из редактора скриптов.
// Токен: категория Marketplace. Берётся из getApiToken() (файл 01).
//
// Документация: https://dev.wildberries.ru/docs/openapi/orders-fbs
//
// Результат — в активную книгу (ту же, что 01–04). Сырой накопитель: FBS_Заказы.
// Не трогает лист «Заказы» из файла 01.
//
// Функции:
//   loadFbsOrders   — выгрузка: новые + 30 дней + статусы + стикеры + архив в накопитель
// ============================================

var ORD_MP = 'https://marketplace-api.wildberries.ru';
var ORD_PAUSE_MS = 250;
var ORD_LIMIT = 1000;
var ORD_STATUS_BATCH = 1000;
var ORD_STICKER_BATCH = 100;

var ORD_CARGO = { 1: 'МГТ', 2: 'СГТ', 3: 'КГТ+' };
var ORD_CARGO_STR = { mgt: 'МГТ', sgt: 'СГТ', kgtPlus: 'КГТ+' };

var ORD_SUPPLIER = {
  new: 'новое',
  confirm: 'на сборке',
  complete: 'в доставке',
  cancel: 'отменено продавцом',
  cancel_carrier: 'отменено перевозчиком'
};

var ORD_WB = {
  waiting: 'ожидает поступления на склад WB',
  sorted: 'отсортировано складом WB',
  sold: 'продано',
  canceled: 'отменено продавцом',
  canceled_by_client: 'отменено покупателем при получении',
  declined_by_client: 'отменено покупателем в первый час',
  defect: 'брак',
  ready_for_pickup: 'в ПВЗ, ждёт клиента',
  postponed_delivery: 'доставка курьером отложена',
  accepted_by_carrier: 'принято перевозчиком',
  sent_to_carrier: 'передано перевозчику',
  canceled_by_carrier: 'отменено перевозчиком'
};

var ORD_ORDER_HEADERS = [
  'id', 'rid', 'orderUid', 'createdAt', 'article', 'nmId', 'chrtId', 'skus',
  'warehouseId', 'officeId', 'offices', 'supplyId', 'deliveryType',
  'cargoType', 'cargoType_name', 'crossBorderType',
  'isZeroOrder', 'isPickupPointShipmentAllowed', 'comment', 'colorCode',
  'price', 'price_rub', 'finalPrice', 'finalPrice_rub',
  'convertedPrice', 'convertedPrice_rub', 'convertedFinalPrice', 'convertedFinalPrice_rub',
  'salePrice', 'salePrice_rub', 'scanPrice',
  'currencyCode', 'convertedCurrencyCode',
  'ddate', 'sellerDate', 'requiredMeta', 'optionalMeta',
  'address_full', 'longitude', 'latitude', 'isB2b', 'source'
];

var ORD_STATUS_HEADERS = [
  'id', 'supplierStatus', 'supplierStatus_name', 'wbStatus', 'wbStatus_name', 'isCancellable'
];

var ORD_ARCHIVE_HEADERS = [
  'id', 'rid', 'orderUid', 'createdAt', 'article', 'nmId', 'chrtId', 'skus',
  'warehouseId', 'supplyId', 'stickerId',
  'cargoType', 'cargoType_name', 'crossBorderType', 'parcel',
  'isZeroOrder', 'colorCode', 'isB2b',
  'price', 'price_rub', 'convertedPrice', 'convertedPrice_rub',
  'currencyCode', 'convertedCurrencyCode', 'scanPrice',
  'supplierStatus', 'wbStatus', 'year', 'month'
];

// --------------------------------------------
// ВЫГРУЗКА
// --------------------------------------------
function loadFbsOrders(opts) {
  if (typeof renameWbHeadersAllBooks_ === 'function') renameWbHeadersAllBooks_();
  var daily = wbDaily_(opts);
  var token = getApiToken();
  if (!token) return;

  var log = [];
  log.push(['Шаг', 'Результат']);

  var newRes = ordFetchJson_(token, 'GET', ORD_MP + '/api/v3/orders/new', null);
  var newOrders = ordExtractList_(newRes.data, 'orders');
  log.push(['GET /api/v3/orders/new', newRes.code + ' | ' + newOrders.length]);

  var listOrders = ordFetchAllOrders_(token, log);

  var allForStatus = newOrders.concat(listOrders);
  var ids = ordUniqueIds_(allForStatus);
  var statusItems = ordFetchStatuses_(token, ids, log);
  var statusMap = {};
  for (var s = 0; s < statusItems.length; s++) {
    if (statusItems[s] && statusItems[s].id != null) statusMap[String(statusItems[s].id)] = statusItems[s];
  }

  var confirmIds = ordIdsBySupplier_(statusItems, ['confirm', 'complete']);
  var stickers = [];
  if (!daily) {
    if (confirmIds.length) {
      var stickerUrl = ORD_MP + '/api/v3/orders/stickers?type=png&width=58&height=40';
      stickers = ordPostBatches_(token, stickerUrl, confirmIds.slice(0, 100), ORD_STICKER_BATCH, log, 'stickers');
    } else {
      log.push(['stickers', 'Нет заданий в confirm/complete — этикетки не запрашивались']);
    }
  } else {
    log.push(['stickers', 'пропущены (ежедневный режим)']);
  }

  var archive = ordFetchArchiveRange_(token, log, 4, 18, 8000);
  log.push(['Архив для накопителя', String(archive.length)]);

  var jobs = [];
  if (newOrders.length) jobs.push(['FBS_Заказы_новые', ordOrdersToRows_(newOrders, 'new')]);
  else log.push(['FBS_Заказы_новые', 'пусто — лист не трогали']);

  var incoming = [];
  if (listOrders.length) incoming = ordOrdersMergedToRows_(listOrders, statusMap, 'orders_30d');
  if (newOrders.length) incoming = wbMergeById_(incoming, ordOrdersToRows_(newOrders, 'new'), ['id'], 'FBS_Заказы');
  if (archive.length) incoming = wbMergeById_(incoming, ordArchiveToRows_(archive, '', ''), ['id'], 'FBS_Заказы');

  var ssStore = ordBook_();
  var existing = wbReadSheet_(ssStore, ['FBS_Заказы']);
  var merged = wbMergeById_(existing, incoming, ['id'], 'FBS_Заказы');
  log.push(['Заказов было / пришло / стало', Math.max(0, existing.length - 1) + ' / ' + Math.max(0, incoming.length - 1) + ' / ' + Math.max(0, merged.length - 1)]);

  var mergedIds = ordIdsFromTable_(merged);
  var needStatus = [];
  for (var m = 0; m < mergedIds.length; m++) {
    if (!statusMap[String(mergedIds[m])]) needStatus.push(mergedIds[m]);
  }
  var extraStatus = ordFetchStatuses_(token, needStatus, log);
  var statusAll = statusItems.concat(extraStatus);
  merged = ordPatchStatuses_(merged, statusAll);
  if (merged.length > 1) jobs.push(['FBS_Заказы', merged]);
  else log.push(['FBS_Заказы', 'пусто — лист не трогали']);

  if (!daily && statusAll.length) jobs.push(['FBS_Заказы_статусы', ordStatusesToRows_(statusAll)]);
  if (stickers.length) jobs.push(['FBS_Заказы_стикеры', ordStickersToRows_(stickers)]);

  var url = ordDumpSheets_(jobs, log);
  if (!daily && !(opts && opts.skipCompose) && typeof composeFbsReports === 'function') {
    try { url = composeFbsReports(true) || url; } catch (e) { log.push(['Сборка рабочих листов', String(e)]); }
  }
  ordNotify_(
    'Выгрузка: новых ' + newOrders.length +
    ', за 30 дней ' + listOrders.length +
    ', накопитель ' + Math.max(0, merged.length - 1) +
    ', статусов ' + statusAll.length +
    ', архив ' + archive.length,
    url
  );
}

// --------------------------------------------
// HTTP
// --------------------------------------------
function ordFetchJson_(token, method, url, body) {
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

function ordFetchAllOrders_(token, log) {
  var all = [];
  var next = 0;
  var pages = 0;
  var nowSec = Math.floor(Date.now() / 1000);
  var fromSec = nowSec - 30 * 24 * 3600;

  while (pages < 50) {
    pages++;
    var url = ORD_MP + '/api/v3/orders?limit=' + ORD_LIMIT +
      '&next=' + next + '&dateFrom=' + fromSec + '&dateTo=' + nowSec;
    var res = ordFetchJson_(token, 'GET', url, null);
    if (res.code !== 200) {
      if (log) log.push(['GET /api/v3/orders стр.' + pages, res.code + ' | ' + (res.text || '').substring(0, 400)]);
      break;
    }
    var orders = ordExtractList_(res.data, 'orders');
    if (orders.length) all = all.concat(orders);
    var newNext = res.data && res.data.next !== undefined && res.data.next !== null ? res.data.next : null;
    if (log) log.push(['GET /api/v3/orders стр.' + pages, res.code + ' | ' + orders.length + ' | next=' + newNext]);
    if (!orders.length || newNext === null || newNext === next) break;
    next = newNext;
    Utilities.sleep(ORD_PAUSE_MS);
  }
  if (log) log.push(['GET /api/v3/orders всего', String(all.length)]);
  return all;
}

function ordFetchStatuses_(token, ids, log) {
  var all = [];
  for (var i = 0; i < ids.length; i += ORD_STATUS_BATCH) {
    var batch = ids.slice(i, i + ORD_STATUS_BATCH);
    var res = ordFetchJson_(token, 'POST', ORD_MP + '/api/v3/orders/status', { orders: batch });
    var items = ordExtractList_(res.data, 'orders');
    if (log) log.push(['POST /status batch ' + (Math.floor(i / ORD_STATUS_BATCH) + 1), res.code + ' | ' + items.length]);
    if (res.code === 200) all = all.concat(items);
    else if (log && res.text) log.push(['Ошибка status', res.text.substring(0, 800)]);
    Utilities.sleep(ORD_PAUSE_MS);
  }
  return all;
}

function ordPostBatches_(token, url, ids, batchSize, log, label) {
  var all = [];
  for (var i = 0; i < ids.length; i += batchSize) {
    var batch = ids.slice(i, i + batchSize);
    var res = ordFetchJson_(token, 'POST', url, { orders: batch });
    var key = url.indexOf('stickers') >= 0 ? 'stickers' : 'orders';
    var items = ordExtractList_(res.data, key);
    if (log) log.push(['POST ' + label + ' batch ' + (Math.floor(i / batchSize) + 1), res.code + ' | ' + items.length]);
    if (res.code === 200) all = all.concat(items);
    else if (log && res.text) log.push(['Ошибка ' + label, res.text.substring(0, 800)]);
    Utilities.sleep(ORD_PAUSE_MS);
  }
  return all;
}

function ordFetchArchivePage_(token, year, month, next, limit) {
  var url = ORD_MP + '/api/marketplace/v3/fbs/orders/archive' +
    '?year=' + year + '&month=' + month + '&next=' + next + '&limit=' + limit;
  var res = ordFetchJson_(token, 'GET', url, null);
  return {
    res: res,
    orders: ordExtractList_(res.data, 'orders'),
    next: res.data && res.data.next !== undefined ? res.data.next : null
  };
}

function ordFetchArchiveAll_(token, year, month, log, maxRows) {
  var all = [];
  var next = 0;
  var pages = 0;
  var cap = maxRows || 5000;
  while (pages < 30 && all.length < cap) {
    pages++;
    var page = ordFetchArchivePage_(token, year, month, next, 1000);
    if (page.res.code !== 200) {
      if (log) log.push(['GET archive ' + year + '-' + month + ' стр.' + pages, page.res.code + ' | ' + (page.res.text || '').substring(0, 400)]);
      break;
    }
    if (page.orders.length) all = all.concat(page.orders);
    if (log) log.push(['GET archive ' + year + '-' + month + ' стр.' + pages, page.res.code + ' | ' + page.orders.length + ' | next=' + page.next]);
    if (!page.orders.length || page.next === null || page.next === next) break;
    next = page.next;
    Utilities.sleep(ORD_PAUSE_MS);
  }
  if (log) log.push(['archive ' + year + '-' + month + ' всего', String(all.length)]);
  return all;
}

function ordFetchArchiveRange_(token, log, startMonthsAgo, howMany, maxRows) {
  var all = [];
  var now = new Date();
  var emptyStreak = 0;
  var cap = maxRows || 8000;
  for (var i = startMonthsAgo; i < startMonthsAgo + howMany; i++) {
    if (all.length >= cap) break;
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    var part = ordFetchArchiveAll_(token, d.getFullYear(), d.getMonth() + 1, log, cap - all.length);
    if (!part.length) {
      emptyStreak++;
      if (emptyStreak >= 6) break;
    } else {
      emptyStreak = 0;
      all = all.concat(part);
    }
  }
  if (log) log.push(['archive диапазон всего', String(all.length)]);
  return all;
}

function ordIdsFromTable_(table) {
  var out = [];
  if (!table || table.length < 2) return out;
  var headers = table[0];
  var idCol = typeof wbColIdx_ === 'function'
    ? wbColIdx_(headers, ['id', 'orderId', 'orderID', 'ID задания'])
    : -1;
  if (idCol < 0) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '');
      if (h === 'id' || h === 'orderId' || h === 'orderID' || h === 'ID задания') {
        idCol = i;
        break;
      }
    }
  }
  if (idCol < 0) return out;
  var seen = {};
  for (var r = 1; r < table.length; r++) {
    var id = table[r][idCol];
    if (id === undefined || id === null || id === '') continue;
    var key = String(id);
    if (seen[key]) continue;
    seen[key] = true;
    out.push(id);
  }
  return out;
}

function ordPatchStatuses_(table, statusItems) {
  if (!table || table.length < 2 || !statusItems || !statusItems.length) return table;
  var h = table[0];
  var idx = {};
  for (var i = 0; i < h.length; i++) idx[String(h[i])] = i;
  var idCol = typeof wbColIdx_ === 'function' ? wbColIdx_(h, ['id', 'orderId', 'orderID', 'ID задания']) : -1;
  if (idCol < 0) {
    idCol = idx.id !== undefined ? idx.id : (idx.orderId !== undefined ? idx.orderId : idx.orderID);
  }
  if (idCol === undefined || idCol < 0) return table;

  function ensureCol(name) {
    var found = typeof wbColIdx_ === 'function' ? wbColIdx_(h, [name]) : -1;
    if (found >= 0) return found;
    if (idx[name] !== undefined) return idx[name];
    idx[name] = h.length;
    h.push(name);
    for (var r = 1; r < table.length; r++) table[r].push('');
    return idx[name];
  }

  var cStatus = ensureCol('supplierStatus');
  var cStatusN = ensureCol('supplierStatus_name');
  var cWb = ensureCol('wbStatus');
  var cWbN = ensureCol('wbStatus_name');
  var cCan = ensureCol('isCancellable');

  var map = {};
  for (var s = 0; s < statusItems.length; s++) {
    if (statusItems[s] && statusItems[s].id != null) map[String(statusItems[s].id)] = statusItems[s];
  }

  for (var r = 1; r < table.length; r++) {
    var st = map[String(table[r][idCol] || '')];
    if (!st) continue;
    table[r][cStatus] = st.supplierStatus || table[r][cStatus] || '';
    table[r][cStatusN] = ORD_SUPPLIER[st.supplierStatus] || table[r][cStatusN] || '';
    table[r][cWb] = st.wbStatus || table[r][cWb] || '';
    table[r][cWbN] = ORD_WB[st.wbStatus] || table[r][cWbN] || '';
    table[r][cCan] = st.isCancellable === true ? 'Да' : (st.isCancellable === false ? 'Нет' : (table[r][cCan] || ''));
  }
  return table;
}

// --------------------------------------------
// Строки
// --------------------------------------------
function ordOrdersToRows_(orders, source) {
  return ordOrdersMergedToRows_(orders, {}, source);
}

function ordOrdersMergedToRows_(orders, statusMap, source) {
  var extra = ordExtraKeys_(orders, [
    'id', 'rid', 'orderUid', 'createdAt', 'article', 'nmId', 'nmID', 'chrtId', 'skus',
    'warehouseId', 'officeId', 'offices', 'supplyId', 'deliveryType', 'cargoType',
    'crossBorderType', 'isZeroOrder', 'isPickupPointShipmentAllowed', 'comment', 'colorCode',
    'price', 'finalPrice', 'convertedPrice', 'convertedFinalPrice', 'salePrice', 'scanPrice',
    'currencyCode', 'convertedCurrencyCode', 'ddate', 'sellerDate', 'requiredMeta', 'optionalMeta',
    'address', 'options'
  ]);
  var header = ORD_ORDER_HEADERS.concat([
    'supplierStatus', 'supplierStatus_name', 'wbStatus', 'wbStatus_name', 'isCancellable'
  ]).concat(extra);
  var rows = [header];
  for (var i = 0; i < orders.length; i++) {
    var o = orders[i] || {};
    var addr = o.address || {};
    var opt = o.options || {};
    var st = (statusMap && o.id != null) ? (statusMap[o.id] || {}) : {};
    var row = [
      o.id || '',
      o.rid || '',
      o.orderUid || '',
      o.createdAt || '',
      o.article || '',
      o.nmId || o.nmID || '',
      o.chrtId || '',
      ordJoin_(o.skus),
      o.warehouseId || '',
      o.officeId || '',
      ordJoin_(o.offices),
      o.supplyId || '',
      o.deliveryType || '',
      o.cargoType !== undefined && o.cargoType !== null ? o.cargoType : '',
      ORD_CARGO[o.cargoType] || '',
      o.crossBorderType !== undefined && o.crossBorderType !== null ? o.crossBorderType : '',
      o.isZeroOrder === true ? 'Да' : (o.isZeroOrder === false ? 'Нет' : ''),
      o.isPickupPointShipmentAllowed === true ? 'Да' : (o.isPickupPointShipmentAllowed === false ? 'Нет' : ''),
      o.comment || '',
      o.colorCode || '',
      ordNum_(o.price),
      ordMoney_(o.price),
      ordNum_(o.finalPrice),
      ordMoney_(o.finalPrice),
      ordNum_(o.convertedPrice),
      ordMoney_(o.convertedPrice),
      ordNum_(o.convertedFinalPrice),
      ordMoney_(o.convertedFinalPrice),
      ordNum_(o.salePrice),
      ordMoney_(o.salePrice),
      o.scanPrice !== undefined && o.scanPrice !== null ? o.scanPrice : '',
      o.currencyCode || '',
      o.convertedCurrencyCode || '',
      o.ddate || '',
      o.sellerDate || '',
      ordJoin_(o.requiredMeta),
      ordJoin_(o.optionalMeta),
      addr.fullAddress || '',
      addr.longitude || '',
      addr.latitude || '',
      opt.isB2b === true || opt.isB2B === true ? 'Да' : (opt.isB2b === false || opt.isB2B === false ? 'Нет' : ''),
      source || ''
    ];
    row.push(
      st.supplierStatus || '',
      ORD_SUPPLIER[st.supplierStatus] || '',
      st.wbStatus || '',
      ORD_WB[st.wbStatus] || '',
      st.isCancellable === true ? 'Да' : (st.isCancellable === false ? 'Нет' : '')
    );
    for (var k = 0; k < extra.length; k++) row.push(ordScalar_(o[extra[k]]));
    rows.push(row);
  }
  return rows;
}

function ordStatusesToRows_(items) {
  var extra = ordExtraKeys_(items, ['id', 'supplierStatus', 'wbStatus', 'isCancellable']);
  var header = ORD_STATUS_HEADERS.concat(extra);
  var rows = [header];
  for (var i = 0; i < items.length; i++) {
    var o = items[i] || {};
    var row = [
      o.id || '',
      o.supplierStatus || '',
      ORD_SUPPLIER[o.supplierStatus] || '',
      o.wbStatus || '',
      ORD_WB[o.wbStatus] || '',
      o.isCancellable === true ? 'Да' : (o.isCancellable === false ? 'Нет' : '')
    ];
    for (var k = 0; k < extra.length; k++) row.push(ordScalar_(o[extra[k]]));
    rows.push(row);
  }
  return rows;
}

function ordStickersToRows_(items) {
  var extra = ordExtraKeys_(items, ['orderId', 'orderID', 'partA', 'partB', 'barcode', 'file', 'status', 'parcelId']);
  extra = extra.filter(function(k) { return k !== 'file'; });
  var header = ['orderId', 'partA', 'partB', 'barcode', 'status', 'parcelId', 'file_bytes'].concat(extra);
  var rows = [header];
  for (var i = 0; i < items.length; i++) {
    var o = items[i] || {};
    var fileLen = o.file ? String(o.file).length : 0;
    var row = [
      o.orderId || o.orderID || '',
      o.partA || '',
      o.partB || '',
      o.barcode || '',
      o.status || '',
      o.parcelId || '',
      fileLen ? fileLen : ''
    ];
    for (var k = 0; k < extra.length; k++) row.push(ordScalar_(o[extra[k]]));
    rows.push(row);
  }
  return rows;
}

function ordArchiveToRows_(items, year, month) {
  var extra = ordExtraKeys_(items, [
    'id', 'rid', 'orderUid', 'createdAt', 'warehouseId', 'supplyId', 'stickerId',
    'cargoType', 'crossBorderType', 'crossBorder', 'isZeroOrder', 'colorCode',
    'options', 'product', 'priceInfo', 'scanPrice', 'status', 'metaDetails'
  ]);
  var header = ORD_ARCHIVE_HEADERS.concat(extra);
  var rows = [header];
  for (var i = 0; i < items.length; i++) {
    var o = items[i] || {};
    var product = o.product || {};
    var price = o.priceInfo || {};
    var st = o.status || {};
    var opt = o.options || {};
    var xb = o.crossBorder || {};
    var row = [
      o.id || '',
      o.rid || '',
      o.orderUid || '',
      o.createdAt || '',
      product.article || o.article || '',
      product.nmId || o.nmId || '',
      product.chrtId || o.chrtId || '',
      ordJoin_(product.skus || o.skus),
      o.warehouseId || '',
      o.supplyId || '',
      o.stickerId || '',
      o.cargoType || '',
      ORD_CARGO_STR[o.cargoType] || ORD_CARGO[o.cargoType] || '',
      o.crossBorderType || '',
      xb.parcel || '',
      o.isZeroOrder === true ? 'Да' : (o.isZeroOrder === false ? 'Нет' : ''),
      o.colorCode || '',
      opt.isB2b === true || opt.isB2B === true ? 'Да' : (opt.isB2b === false || opt.isB2B === false ? 'Нет' : ''),
      ordNum_(price.price),
      ordMoney_(price.price),
      ordNum_(price.convertedPrice),
      ordMoney_(price.convertedPrice),
      price.currencyCode || '',
      price.convertedCurrencyCode || '',
      o.scanPrice !== undefined && o.scanPrice !== null ? o.scanPrice : '',
      st.supplierStatus || '',
      st.wbStatus || '',
      year || '',
      month || ''
    ];
    for (var k = 0; k < extra.length; k++) row.push(ordScalar_(o[extra[k]]));
    rows.push(row);
  }
  return rows;
}

// --------------------------------------------
// Утилиты
// --------------------------------------------
function ordExtractList_(data, key) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (key && Array.isArray(data[key])) return data[key];
  if (Array.isArray(data.orders)) return data.orders;
  if (Array.isArray(data.stickers)) return data.stickers;
  return [];
}

function ordUniqueIds_(orders) {
  var seen = {};
  var out = [];
  for (var i = 0; i < orders.length; i++) {
    var o = orders[i];
    if (!o) continue;
    var id = o.id || o.orderID || o.orderId;
    if (id == null || seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function ordIdsBySupplier_(statusItems, statuses) {
  var allow = {};
  for (var i = 0; i < statuses.length; i++) allow[statuses[i]] = true;
  var out = [];
  for (var s = 0; s < statusItems.length; s++) {
    var it = statusItems[s];
    if (it && it.id != null && allow[it.supplierStatus]) out.push(it.id);
  }
  return out;
}

function ordBook_() {
  return fbsBook_();
}

function ordDumpSheets_(jobs, log) {
  var ss = ordBook_();
  var url = ss.getUrl();
  log.push(['Файл выгрузки', url]);
  for (var i = 0; i < jobs.length; i++) {
    wbWriteSheetSafe_(ss, jobs[i][0], jobs[i][1], log);
  }
  return url;
}

function ordNotify_(msg, url) {
  var text = msg + (url ? '\n\nОткройте файл:\n' + url : '');
  Logger.log(text);
  if (typeof WB_SILENT_UI !== 'undefined' && WB_SILENT_UI) return;
  try {
    SpreadsheetApp.getUi().alert(text);
  } catch (e) {
    Logger.log('UI недоступен: ' + e);
  }
}

function ordUniqueKeys_(arr) {
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

function ordExtraKeys_(arr, known) {
  var skip = {};
  for (var i = 0; i < known.length; i++) skip[known[i]] = true;
  return ordUniqueKeys_(arr).filter(function(k) { return !skip[k]; });
}

function ordScalar_(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return wbDateTime_(value);
  if (typeof value === 'object') return JSON.stringify(value);
  var parsed = wbParseDate_(value);
  if (parsed) return wbMaybeDate_(value);
  return value;
}

function ordJoin_(arr) {
  return Array.isArray(arr) ? arr.join(', ') : '';
}

function ordNum_(value) {
  return value === undefined || value === null ? '' : value;
}

function ordMoney_(value) {
  if (value === undefined || value === null || value === '') return '';
  var n = Number(value);
  if (isNaN(n)) return '';
  return n / 100;
}
