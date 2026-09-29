// ============================================
// WB API — Поставки FBS (Marketplace)
// Существующие файлы не меняет. Запускать из редактора скриптов.
// Токен: категория Marketplace. Берётся из getApiToken() (файл 01).
// Книга выгрузки: активная (та же, что 01–04). Сырой накопитель: FBS_Поставки.
// Не трогает лист «Поставки» FBW из файла 01.
//
// Это НЕ FBW-поставки из supplies-api (файл 01, листы «Поставки»).
// Здесь поставки сборочных заданий: ID вида WB-GI-...
//
// Документация: https://dev.wildberries.ru/docs/openapi/orders-fbs
//
// Функции:
//   loadFbsSupplies   — полный список + карточка GET /supplies/{id}
//                       (closedAt/scanDt) + ID заданий + короба.
//                       QR и стикеры коробов — у закрытых, без картинок.
// ============================================

var SUP_MP = 'https://marketplace-api.wildberries.ru';
var SUP_PAUSE_MS = 250;
var SUP_LIMIT = 1000;
var SUP_DETAIL_MAX = 400;
var SUP_STICKER_BATCH = 100;

var SUP_CARGO = { 1: 'МГТ', 2: 'СГТ', 3: 'КГТ+' };
var SUP_CROSS = { 0: 'не кроссбордер', 1: 'кроссбордер' };

var SUP_HEADERS = [
  'id', 'name', 'done', 'done_name', 'isB2b',
  'createdAt', 'closedAt', 'scanDt',
  'cargoType', 'cargoType_name', 'crossBorderType', 'crossBorderType_name',
  'destinationOfficeId', 'ordersCount', 'trbxCount'
];

// --------------------------------------------
// ВЫГРУЗКА
// --------------------------------------------
function loadFbsSupplies(opts) {
  if (typeof renameWbHeadersAllBooks_ === 'function') renameWbHeadersAllBooks_();
  var daily = wbDaily_(opts);
  var token = getApiToken();
  if (!token) return;

  var log = [];
  log.push(['Шаг', 'Результат']);

  var supplies = supFetchAllSupplies_(token, log);
  var samplesForDetail = supplies.slice(0, SUP_DETAIL_MAX);
  if (supplies.length > SUP_DETAIL_MAX) {
    log.push([
      'Детали ограничены',
      'список полный (' + supplies.length + '), order-ids/trbx/QR только у первых ' + SUP_DETAIL_MAX
    ]);
  }

  var orderCount = {};
  var trbxCount = {};
  var orderRows = [];
  var trbxRows = [];
  var qrRows = [];
  var stickerRows = [];
  var supplyIndex = {};
  for (var si = 0; si < supplies.length; si++) {
    if (supplies[si] && supplies[si].id) supplyIndex[String(supplies[si].id)] = si;
  }
  var cardsOk = 0;
  var scanFilled = 0;

  for (var i = 0; i < samplesForDetail.length; i++) {
    var sid = samplesForDetail[i].id;
    if (!sid) continue;

    Utilities.sleep(SUP_PAUSE_MS);
    var detRes = supFetchJson_(token, 'GET',
      SUP_MP + '/api/v3/supplies/' + encodeURIComponent(sid), null);
    if (detRes.code === 200 && detRes.data && !Array.isArray(detRes.data)) {
      var card = supMergeSupply_(samplesForDetail[i], detRes.data);
      samplesForDetail[i] = card;
      if (supplyIndex[String(sid)] !== undefined) supplies[supplyIndex[String(sid)]] = card;
      cardsOk++;
      if (card.scanDt) scanFilled++;
    } else if (detRes.code !== 200) {
      log.push(['карточка ' + sid, detRes.code + ' | ' + String(detRes.text || '').substring(0, 300)]);
    }

    Utilities.sleep(SUP_PAUSE_MS);
    var oidRes = supFetchJson_(token, 'GET',
      SUP_MP + '/api/marketplace/v3/supplies/' + encodeURIComponent(sid) + '/order-ids', null);
    var orderIds = (oidRes.code === 200) ? supExtractIds_(oidRes.data, 'orderIds') : [];
    orderCount[sid] = orderIds.length;
    for (var o = 0; o < orderIds.length; o++) {
      orderRows.push({ supplyId: sid, orderId: orderIds[o] });
    }
    if (oidRes.code !== 200) {
      log.push(['order-ids ' + sid, oidRes.code + ' | ' + String(oidRes.text || '').substring(0, 300)]);
    }

    Utilities.sleep(SUP_PAUSE_MS);
    var trbxRes = supFetchJson_(token, 'GET',
      SUP_MP + '/api/v3/supplies/' + encodeURIComponent(sid) + '/trbx', null);
    var boxes = [];
    if (trbxRes.code === 200) {
      boxes = supExtractList_(trbxRes.data, 'trbxes');
      if (!boxes.length) boxes = supExtractList_(trbxRes.data, 'trbx');
    }
    var trbxIds = [];
    for (var b = 0; b < boxes.length; b++) {
      var boxId = boxes[b] && (boxes[b].id || boxes[b].trbxId);
      if (!boxId) continue;
      trbxIds.push(boxId);
      trbxRows.push({ supplyId: sid, trbxId: boxId });
    }
    trbxCount[sid] = trbxIds.length;
    if (trbxRes.code !== 200) {
      log.push(['trbx ' + sid, trbxRes.code + ' | ' + String(trbxRes.text || '').substring(0, 300)]);
    }

    if (!daily && samplesForDetail[i].done === true) {
      Utilities.sleep(SUP_PAUSE_MS);
      var qrRes = supFetchJson_(token, 'GET',
        SUP_MP + '/api/v3/supplies/' + encodeURIComponent(sid) + '/barcode?type=png', null);
      if (qrRes.code === 200 && qrRes.data) {
        qrRows.push({
          supplyId: sid,
          barcode: qrRes.data.barcode || sid,
          http: qrRes.code,
          fileLen: qrRes.data.file ? String(qrRes.data.file).length : 0
        });
      } else if (qrRes.code !== 409) {
        log.push(['barcode ' + sid, qrRes.code + ' | ' + String(qrRes.text || '').substring(0, 300)]);
      }
    } else if (daily && samplesForDetail[i].done === true) {
      qrRows.push({ supplyId: sid, barcode: sid, http: '', fileLen: '' });
    }

    if (!daily && trbxIds.length) {
      for (var t = 0; t < trbxIds.length; t += SUP_STICKER_BATCH) {
        var batch = trbxIds.slice(t, t + SUP_STICKER_BATCH);
        Utilities.sleep(SUP_PAUSE_MS);
        var stUrl = SUP_MP + '/api/v3/supplies/' + encodeURIComponent(sid) + '/trbx/stickers?type=png';
        var stRes = supFetchJson_(token, 'POST', stUrl, { trbxIds: batch });
        var stickers = (stRes.code === 200) ? supExtractList_(stRes.data, 'stickers') : [];
        for (var s = 0; s < stickers.length; s++) {
          stickerRows.push({
            supplyId: sid,
            barcode: stickers[s].barcode || '',
            trbxId: stickers[s].id || stickers[s].trbxId || '',
            fileLen: stickers[s].file ? String(stickers[s].file).length : 0
          });
        }
        if (stRes.code !== 200) {
          log.push(['stickers ' + sid, stRes.code + ' | ' + String(stRes.text || '').substring(0, 300)]);
        }
      }
    }

    if ((i + 1) % 20 === 0) {
      log.push(['Прогресс деталей', (i + 1) + '/' + samplesForDetail.length]);
    }
  }

  var ssStore = supBook_();
  var freshSupplies = supplies.length ? supSuppliesToRows_(supplies, orderCount, trbxCount) : [];
  var mergedSupplies = wbMergeById_(
    wbReadSheet_(ssStore, ['FBS_Поставки']),
    freshSupplies,
    ['id'],
    'FBS_Поставки'
  );
  var mergedOrders = wbMergePairRows_(
    wbReadSheet_(ssStore, ['FBS_Поставки_заказы']),
    orderRows.length ? supOrdersToRows_(orderRows) : [],
    'supplyId',
    'orderId'
  );
  var mergedTrbx = wbMergePairRows_(
    wbReadSheet_(ssStore, ['FBS_Поставки_короба']),
    trbxRows.length ? supTrbxToRows_(trbxRows) : [],
    'supplyId',
    'trbxId'
  );
  var mergedQr = wbMergeById_(
    wbReadSheet_(ssStore, ['FBS_Поставки_QR']),
    qrRows.length ? supQrToRows_(qrRows) : [],
    ['supplyId'],
    'FBS_Поставки_QR'
  );
  mergedSupplies = supApplyCounts_(mergedSupplies, mergedOrders, mergedTrbx);
  log.push(['Карточки GET /api/v3/supplies/{id}',
    cardsOk + ' ок, scanDt заполнен у ' + scanFilled + ' из ' + samplesForDetail.length]);
  log.push(['Поставок было / API / стало',
    Math.max(0, (wbReadSheet_(ssStore, ['FBS_Поставки']) || []).length - 1) + ' / ' + supplies.length + ' / ' + Math.max(0, mergedSupplies.length - 1)]);

  var jobs = [];
  if (mergedSupplies.length > 1) jobs.push(['FBS_Поставки', mergedSupplies]);
  else log.push(['FBS_Поставки', 'пусто — лист не трогали']);

  if (mergedOrders.length > 1) jobs.push(['FBS_Поставки_заказы', mergedOrders]);
  if (mergedTrbx.length > 1) jobs.push(['FBS_Поставки_короба', mergedTrbx]);
  if (mergedQr.length > 1) jobs.push(['FBS_Поставки_QR', mergedQr]);
  if (stickerRows.length) jobs.push(['FBS_Поставки_стикеры_коробов', supStickersToRows_(stickerRows)]);

  var url = supDumpSheets_(jobs, log);
  if (!daily && !(opts && opts.skipCompose) && typeof composeFbsReports === 'function') {
    try { url = composeFbsReports(true) || url; } catch (e) { log.push(['Сборка рабочих листов', String(e)]); }
  }
  supNotify_(
    'Выгрузка поставок: API ' + supplies.length +
    ', накопитель ' + Math.max(0, mergedSupplies.length - 1) +
    ', заданий ' + Math.max(0, mergedOrders.length - 1) +
    ', коробов ' + Math.max(0, mergedTrbx.length - 1),
    url
  );
}

// --------------------------------------------
// HTTP
// --------------------------------------------
function supFetchJson_(token, method, url, body) {
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

function supFetchAllSupplies_(token, log) {
  var all = [];
  var next = 0;
  var pages = 0;

  while (pages < 50) {
    pages++;
    var url = SUP_MP + '/api/v3/supplies?limit=' + SUP_LIMIT + '&next=' + next;
    var res = supFetchJson_(token, 'GET', url, null);
    if (res.code !== 200) {
      if (log) log.push(['GET /api/v3/supplies стр.' + pages, res.code + ' | ' + (res.text || '').substring(0, 400)]);
      break;
    }
    var items = supExtractList_(res.data, 'supplies');
    if (items.length) all = all.concat(items);
    var newNext = res.data && res.data.next !== undefined && res.data.next !== null ? res.data.next : null;
    if (log) log.push(['GET /api/v3/supplies стр.' + pages, res.code + ' | ' + items.length + ' | next=' + newNext]);
    if (!items.length || newNext === null || newNext === next) break;
    next = newNext;
    Utilities.sleep(SUP_PAUSE_MS);
  }
  if (log) log.push(['GET /api/v3/supplies всего', String(all.length)]);
  return all;
}

// --------------------------------------------
// Строки
// --------------------------------------------
function supSuppliesToRows_(items, orderCount, trbxCount) {
  orderCount = orderCount || {};
  trbxCount = trbxCount || {};
  var extra = supExtraKeys_(items, [
    'id', 'name', 'done', 'isB2b', 'createdAt', 'closedAt', 'scanDt',
    'cargoType', 'crossBorderType', 'destinationOfficeId'
  ]);
  var header = SUP_HEADERS.concat(extra);
  var rows = [header];
  for (var i = 0; i < items.length; i++) {
    var s = items[i] || {};
    var sid = s.id || '';
    var row = [
      sid,
      s.name || '',
      s.done === true ? 'Да' : (s.done === false ? 'Нет' : ''),
      s.done === true ? 'закрыта' : (s.done === false ? 'открыта' : ''),
      s.isB2b === true ? 'Да' : (s.isB2b === false ? 'Нет' : ''),
      supScalar_(s.createdAt),
      supScalar_(s.closedAt),
      supScalar_(s.scanDt),
      s.cargoType !== undefined && s.cargoType !== null ? s.cargoType : '',
      SUP_CARGO[s.cargoType] || '',
      s.crossBorderType !== undefined && s.crossBorderType !== null ? s.crossBorderType : '',
      SUP_CROSS[s.crossBorderType] || '',
      s.destinationOfficeId !== undefined && s.destinationOfficeId !== null ? s.destinationOfficeId : '',
      orderCount[sid] !== undefined ? orderCount[sid] : '',
      trbxCount[sid] !== undefined ? trbxCount[sid] : ''
    ];
    for (var k = 0; k < extra.length; k++) row.push(supScalar_(s[extra[k]]));
    rows.push(row);
  }
  return rows;
}

function supApplyCounts_(supplyTable, orderPairs, trbxPairs) {
  if (!supplyTable || supplyTable.length < 2) return supplyTable;
  var oc = {};
  var tc = {};
  if (orderPairs) {
    for (var i = 1; i < orderPairs.length; i++) {
      var sid = String(orderPairs[i][0] || '');
      if (sid) oc[sid] = (oc[sid] || 0) + 1;
    }
  }
  if (trbxPairs) {
    for (var j = 1; j < trbxPairs.length; j++) {
      var tid = String(trbxPairs[j][0] || '');
      if (tid) tc[tid] = (tc[tid] || 0) + 1;
    }
  }
  var h = supplyTable[0];
  var idCol = -1;
  var ocCol = -1;
  var tcCol = -1;
  for (var c = 0; c < h.length; c++) {
    if (typeof wbIs_ === 'function') {
      if (wbIs_(h[c], 'id') || wbIs_(h[c], 'ID поставки')) idCol = c;
      if (wbIs_(h[c], 'ordersCount')) ocCol = c;
      if (wbIs_(h[c], 'trbxCount')) tcCol = c;
    } else {
      if (h[c] === 'id') idCol = c;
      if (h[c] === 'ordersCount') ocCol = c;
      if (h[c] === 'trbxCount') tcCol = c;
    }
  }
  if (idCol < 0) return supplyTable;
  if (ocCol < 0) {
    ocCol = h.length;
    h.push('ordersCount');
    for (var r1 = 1; r1 < supplyTable.length; r1++) supplyTable[r1].push('');
  }
  if (tcCol < 0) {
    tcCol = h.length;
    h.push('trbxCount');
    for (var r2 = 1; r2 < supplyTable.length; r2++) supplyTable[r2].push('');
  }
  for (var r = 1; r < supplyTable.length; r++) {
    var id = String(supplyTable[r][idCol] || '');
    if (oc[id] !== undefined) supplyTable[r][ocCol] = oc[id];
    if (tc[id] !== undefined) supplyTable[r][tcCol] = tc[id];
  }
  return supplyTable;
}

function supOrdersToRows_(items) {
  var rows = [['supplyId', 'orderId']];
  for (var i = 0; i < items.length; i++) {
    rows.push([items[i].supplyId || '', items[i].orderId || '']);
  }
  return rows;
}

function supTrbxToRows_(items) {
  var rows = [['supplyId', 'trbxId']];
  for (var i = 0; i < items.length; i++) {
    rows.push([items[i].supplyId || '', items[i].trbxId || '']);
  }
  return rows;
}

function supQrToRows_(items) {
  var rows = [['supplyId', 'barcode', 'http', 'file_chars']];
  for (var i = 0; i < items.length; i++) {
    rows.push([
      items[i].supplyId || '',
      items[i].barcode || '',
      items[i].http || '',
      items[i].fileLen || ''
    ]);
  }
  return rows;
}

function supStickersToRows_(items) {
  var rows = [['supplyId', 'trbxId', 'barcode', 'file_chars']];
  for (var i = 0; i < items.length; i++) {
    rows.push([
      items[i].supplyId || '',
      items[i].trbxId || '',
      items[i].barcode || '',
      items[i].fileLen || ''
    ]);
  }
  return rows;
}

// --------------------------------------------
// Утилиты
// --------------------------------------------
function supExtractList_(data, key) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (key && Array.isArray(data[key])) return data[key];
  if (Array.isArray(data.supplies)) return data.supplies;
  if (Array.isArray(data.trbxes)) return data.trbxes;
  if (Array.isArray(data.stickers)) return data.stickers;
  return [];
}

function supExtractIds_(data, key) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (key && Array.isArray(data[key])) return data[key];
  if (Array.isArray(data.orderIds)) return data.orderIds;
  if (Array.isArray(data.orders)) return data.orders;
  return [];
}

function supBook_() {
  return fbsBook_();
}

function supDumpSheets_(jobs, log) {
  var ss = supBook_();
  var url = ss.getUrl();
  log.push(['Файл выгрузки', url]);
  for (var i = 0; i < jobs.length; i++) {
    wbWriteSheetSafe_(ss, jobs[i][0], jobs[i][1], log);
  }
  return url;
}

function supNotify_(msg, url) {
  var text = msg + (url ? '\n\nОткройте файл:\n' + url : '');
  Logger.log(text);
  if (typeof WB_SILENT_UI !== 'undefined' && WB_SILENT_UI) return;
  try {
    SpreadsheetApp.getUi().alert(text);
  } catch (e) {
    Logger.log('UI недоступен: ' + e);
  }
}

function supUniqueKeys_(arr) {
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

function supExtraKeys_(arr, known) {
  var skip = {};
  for (var i = 0; i < known.length; i++) skip[known[i]] = true;
  return supUniqueKeys_(arr).filter(function(k) { return !skip[k]; });
}

function supScalar_(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return wbDateTime_(value);
  if (typeof value === 'object') return JSON.stringify(value);
  var parsed = wbParseDate_(value);
  if (parsed) return wbMaybeDate_(value);
  return value;
}

function supHasVal_(value) {
  return value !== undefined && value !== null && value !== '';
}

function supMergeSupply_(listItem, card) {
  var out = {};
  var a = listItem || {};
  var b = card || {};
  var seen = {};
  var ka = Object.keys(a);
  var kb = Object.keys(b);
  var i;
  for (i = 0; i < ka.length; i++) seen[ka[i]] = true;
  for (i = 0; i < kb.length; i++) seen[kb[i]] = true;
  var keys = Object.keys(seen);
  for (i = 0; i < keys.length; i++) {
    var k = keys[i];
    out[k] = supHasVal_(b[k]) ? b[k] : a[k];
  }
  return out;
}
