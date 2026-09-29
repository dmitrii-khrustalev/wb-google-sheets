// ============================================
// WB FBS — сборка рабочих листов из сырой выгрузки
// Существующие файлы не меняет. API не дергает.
// Книга: пишет сырьё в лёгкий файл FBS, копии листов — в активную (главную).
//
// Берёт сырые FBS_* и собирает видимые (с префиксом FBS, чтобы не затереть 01):
//   FBS Поставки / FBS Поставка_состав / FBS Заказы / FBS Остатки
//   FBS Склады / FBS Офисы
// На «FBS Заказы», «FBS Остатки» и «FBS Поставка_состав» при сборке
// считаются: Артикул, Дата (где нет), Вес, кг за 1 шт, Вес, кг за 1 шт сумма.
// Сырые FBS_* и «FBS Служебное» скрываются. Пробы и неиспользуемые листы удаляются.
//
// Функции:
//   composeFbsReports     — собрать рабочие листы
//   unhideRawFbsSheets    — вернуть скрытые FBS_* на экран
//   updateFbsDaily        — ежедневное обновление: заказы → поставки → остатки
//   setupFbsDailyTrigger  — поставить триггер на 8:00 (после updateAllReports1/2)
// ============================================

var CMP_SUPPLY_COLS = [
  'id', 'name', 'done', 'done_name', 'isB2b',
  'createdAt', 'closedAt', 'scanDt',
  'cargoType', 'cargoType_name', 'crossBorderType', 'crossBorderType_name',
  'destinationOfficeId', 'ordersCount', 'trbxCount', 'trbxId',
  'barcode', 'rejectDt', 'recommendedWh', 'isPickupPointShipment'
];

var CMP_ORDER_COLS = [
  'id', 'rid', 'createdAt', 'article', 'nmId', 'chrtId', 'skus',
  'warehouseId', 'officeId', 'supplyId',
  'supplierStatus', 'supplierStatus_name', 'wbStatus', 'wbStatus_name',
  'price_rub', 'finalPrice_rub', 'salePrice_rub',
  'deliveryType', 'cargoType_name', 'isB2b',
  'sticker_partA', 'sticker_partB', 'sticker_barcode',
  'source', 'comment', 'isCancellable', 'orderUid', 'scanPrice',
  'Артикул', 'Дата', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма', 'Регион поставки'
];

var CMP_COMP_COLS = [
  'supplyId', 'supplyName', 'done_name', 'trbxId',
  'supplyCreatedAt', 'supplyClosedAt',
  'orderId', 'article', 'nmId', 'chrtId', 'skus',
  'supplierStatus_name', 'wbStatus_name',
  'price_rub', 'finalPrice_rub',
  'orderCreatedAt', 'warehouseId', 'есть_в_заказах',
  'Артикул', 'Дата', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'
];

var CMP_DROP = [
  'FBS_Поставки_Probe', 'FBS_Поставки_карта', 'FBS_Поставки_детали',
  'FBS_Probe',
  'FBS_Заказы_Probe', 'FBS_Заказы_карта', 'FBS_Заказы_повтор',
  'FBS_Заказы_клиент', 'FBS_Заказы_история', 'FBS_Заказы_архив'
];

var CMP_HIDE = [
  'FBS_Поставки', 'FBS_Поставки_заказы', 'FBS_Поставки_короба',
  'FBS_Поставки_QR', 'FBS_Поставки_стикеры_коробов',
  'FBS_Офисы', 'FBS_Склады', 'FBS_Остатки',
  'FBS_Заказы', 'FBS_Заказы_новые', 'FBS_Заказы_статусы', 'FBS_Заказы_стикеры',
  'FBS Служебное'
];

var CMP_VISIBLE = [
  'FBS Поставки', 'FBS Поставка_состав', 'FBS Заказы', 'FBS Остатки',
  'FBS Склады', 'FBS Офисы'
];

var CMP_SUPPLIER = {
  new: 'новое',
  confirm: 'на сборке',
  complete: 'в доставке',
  cancel: 'отменено продавцом',
  cancel_carrier: 'отменено перевозчиком'
};

var CMP_WB = {
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

var WB_SILENT_UI = false;

function updateFbsDaily() {
  WB_SILENT_UI = true;
  var t0 = Date.now();
  var lines = [];
  if (typeof tgTakeNotes_ === 'function') tgTakeNotes_();
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  try {
    Logger.log('updateFbsDaily старт');
    if (typeof morningTgPendingList_ === 'function') {
      morningTgPendingList_(['FBS заказы', 'FBS поставки']);
    }
    scheduleFbsStocks_(6.5 * 60 * 1000);

    if (typeof tgWatch_ === 'function') {
      tgRunStep_(lines, 'FBS заказы', function() {
        tgWatch_(lines, 'FBS_Заказы', function() {
          loadFbsOrders(true);
        });
      });
      Logger.log('заказы, мс: ' + (Date.now() - t0));

      tgRunStep_(lines, 'FBS поставки', function() {
        tgWatch_(lines, 'FBS_Поставки', function() {
          loadFbsSupplies(true);
        });
      });
      Logger.log('поставки, мс: ' + (Date.now() - t0));
    } else {
      loadFbsOrders(true);
      loadFbsSupplies(true);
    }

    scheduleFbsStocks_(20 * 1000);
    lines.push('FBS_Остатки: запланированы отдельно');
    lines.push('сборка FBS: после остатков');

    if (typeof tgTakeNotes_ === 'function') {
      var notes = tgTakeNotes_();
      for (var i = 0; i < notes.length; i++) lines.push(notes[i]);
    }
    Logger.log('готово, мс: ' + (Date.now() - t0));
    if (typeof tgSendJob_ === 'function') tgSendJob_('updateFbsDaily', true, lines, t0);
  } catch (e) {
    lines.push('❌ ' + e);
    if (typeof tgSendJob_ === 'function') tgSendJob_('updateFbsDaily', false, lines, t0);
  } finally {
    WB_SILENT_UI = false;
  }
}

function scheduleFbsStocks_(afterMs) {
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  if (typeof morningTgPending_ === 'function') morningTgPending_('FBS остатки');
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runFbsStocksJob_');
  }
  ScriptApp.newTrigger('runFbsStocksJob_')
    .timeBased()
    .after(afterMs || 20 * 1000)
    .create();
  Logger.log('FBS остатки: запуск через ' + (afterMs || 20000) + ' мс');
}

function runFbsStocksJob_() {
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runFbsStocksJob_');
  }
  var t0 = Date.now();
  var lines = [];
  WB_SILENT_UI = true;
  try {
    Logger.log('FBS остатки старт');
    if (typeof morningTgPending_ === 'function') morningTgPending_('FBS остатки');
    scheduleComposeFbs_(7 * 60 * 1000);
    if (typeof tgWatch_ === 'function') {
      tgWatch_(lines, 'FBS_Остатки', function() {
        loadSellerWarehouseStocks(true);
      }, 'FBS остатки');
    } else {
      loadSellerWarehouseStocks(true);
      if (typeof morningTgNote_ === 'function') morningTgNote_(true, 'FBS остатки', '');
    }
    scheduleComposeFbs_(20 * 1000);
    lines.push('сборка FBS: запланирована отдельно');
    if (typeof tgSendJob_ === 'function') tgSendJob_('fbsStocks', true, lines, t0);
  } catch (e) {
    lines.push('❌ ' + e);
    if (typeof morningTgNote_ === 'function') morningTgNote_(false, 'FBS остатки', String(e));
    if (typeof tgSendJob_ === 'function') tgSendJob_('fbsStocks', false, lines, t0);
    scheduleComposeFbs_(20 * 1000);
  } finally {
    WB_SILENT_UI = false;
  }
}

function scheduleComposeFbs_(afterMs) {
  PropertiesService.getScriptProperties().setProperty('WB_FBS_COMPOSE_BUSY', '1');
  if (typeof morningTgPending_ === 'function') morningTgPending_('Сборка FBS');
  PropertiesService.getScriptProperties().deleteProperty('FBS_COMPOSE_COPY_IDX');
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runComposeFbsJob_');
    clearTriggersForHandler_('runComposeFbsCopyJob_');
  }
  ScriptApp.newTrigger('runComposeFbsJob_')
    .timeBased()
    .after(afterMs || 20 * 1000)
    .create();
  Logger.log('FBS сборка: запуск через ' + (afterMs || 20000) + ' мс');
}

function runComposeFbsJob_() {
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runComposeFbsJob_');
  }
  WB_SILENT_UI = true;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WB_FBS_COMPOSE_BUSY', '1');
  props.setProperty('WB_FBS_COMPOSE_RUN', '1');
  try {
    Logger.log('FBS сборка старт (сырой файл)');
    composeFbsReports(true);
  } catch (e) {
    props.deleteProperty('WB_FBS_COMPOSE_BUSY');
    props.deleteProperty('WB_FBS_COMPOSE_RUN');
    if (typeof morningTgNote_ === 'function') morningTgNote_(false, 'Сборка FBS', String(e));
    if (typeof tgSendJob_ === 'function') {
      tgSendJob_('composeFbsReports', false, ['❌ ' + e], Date.now());
    }
  } finally {
    WB_SILENT_UI = false;
  }
}

function scheduleComposeFbsCopy_() {
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runComposeFbsCopyJob_');
  }
  ScriptApp.newTrigger('runComposeFbsCopyJob_')
    .timeBased()
    .after(20 * 1000)
    .create();
  Logger.log('FBS копия в главную: через ~20 сек');
}

function runComposeFbsCopyJob_() {
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runComposeFbsCopyJob_');
  }
  var props = PropertiesService.getScriptProperties();
  var names = CMP_VISIBLE.concat(['FBS Служебное']);
  var idx = Number(props.getProperty('FBS_COMPOSE_COPY_IDX') || '0');
  var t0 = Date.now();
  if (idx >= names.length) {
    composeFbsCopyFinish_(props, t0);
    return;
  }
  ScriptApp.newTrigger('runComposeFbsCopyJob_').timeBased().after(7 * 60 * 1000).create();
  var light = fbsSidecar_();
  var dest = fbsBook_();
  var MAX = 3.5 * 60 * 1000;
  var started = idx;
  try {
    while (idx < names.length) {
      if (idx > started && Date.now() - t0 > MAX) {
        props.setProperty('FBS_COMPOSE_COPY_IDX', String(idx));
        if (typeof clearTriggersForHandler_ === 'function') {
          clearTriggersForHandler_('runComposeFbsCopyJob_');
        }
        ScriptApp.newTrigger('runComposeFbsCopyJob_').timeBased().after(20 * 1000).create();
        Logger.log('FBS копия: пауза, следующий ' + names[idx]);
        return;
      }
      props.setProperty('FBS_COMPOSE_COPY_IDX', String(idx));
      Logger.log('FBS копия листа ' + names[idx]);
      wbCopySheetTo_(light, names[idx], dest, null);
      idx++;
    }
    composeFbsCopyFinish_(props, t0);
  } catch (e) {
    Logger.log('FBS копия: ' + e);
    props.deleteProperty('WB_FBS_COMPOSE_BUSY');
    props.deleteProperty('WB_FBS_COMPOSE_RUN');
    if (typeof morningTgNote_ === 'function') morningTgNote_(false, 'Сборка FBS', String(e));
    if (typeof tgSendJob_ === 'function') {
      tgSendJob_('composeFbsReports', false, ['❌ копия в главную: ' + e], t0);
    }
  }
}

function composeFbsCopyFinish_(props, t0) {
  if (typeof clearTriggersForHandler_ === 'function') {
    clearTriggersForHandler_('runComposeFbsCopyJob_');
  }
  props.deleteProperty('FBS_COMPOSE_COPY_IDX');
  try { cmpHideRaw_(fbsBook_(), null); } catch (eH) {}
  try { wbEnsureFbsFilters_(fbsBook_(), null); } catch (eF) {}
  props.deleteProperty('WB_FBS_COMPOSE_BUSY');
  props.deleteProperty('WB_FBS_COMPOSE_RUN');
  Logger.log('FBS копия: все листы в главной');
  if (typeof morningTgNote_ === 'function') morningTgNote_(true, 'Сборка FBS', '');
  if (typeof tgSendJob_ === 'function') {
    tgSendJob_('composeFbsReports', true, ['Сборка FBS: листы скопированы в главную книгу'], t0);
  }
}

function setupFbsDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateFbsDaily') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('updateFbsDaily')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  var msg = 'Ежедневный триггер updateFbsDaily на 8:00 (' + Session.getScriptTimeZone() + ').';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function composeFbsReports(silent) {
  var ss = fbsSidecar_();
  if (typeof renameWbColumnHeaders_ === 'function') {
    try { renameWbColumnHeaders_(ss); } catch (eRen) {}
  }
  var log = [];
  log.push(['Шаг', 'Результат']);
  log.push(['Источник', ss.getUrl()]);
  cmpDropJunk_(ss, log);

  var supplies = cmpBuildSupplies_(ss, log);
  var orders = cmpBuildOrders_(ss, log);
  var composition = cmpBuildComposition_(ss, supplies, orders, log);
  var stocks = cmpReadTable_(ss, ['FBS_Остатки']);
  var warehouses = cmpReadTable_(ss, ['FBS_Склады']);
  var offices = cmpReadTable_(ss, ['FBS_Офисы']);

  if (supplies.length > 1) cmpWrite_(ss, 'FBS Поставки', supplies, log);
  else log.push(['FBS Поставки', 'нет данных — лист не писали']);

  if (orders.length > 1) {
    cmpWrite_(ss, 'FBS Заказы', orders, log);
    cmpFormatOrderCalcCols_(ss, 'FBS Заказы');
  } else {
    log.push(['FBS Заказы', 'нет данных — лист не писали']);
  }

  if (composition.length > 1) {
    cmpWrite_(ss, 'FBS Поставка_состав', composition, log);
    cmpFormatOrderCalcCols_(ss, 'FBS Поставка_состав');
  } else {
    log.push(['FBS Поставка_состав', 'нет данных — лист не писали']);
  }

  if (stocks && stocks.length > 1) {
    stocks = cmpEnrichStocks_(ss, stocks, log);
    cmpWrite_(ss, 'FBS Остатки', stocks, log);
    cmpFormatOrderCalcCols_(ss, 'FBS Остатки');
  }
  if (warehouses && warehouses.length > 1) cmpWrite_(ss, 'FBS Склады', warehouses, log);
  if (offices && offices.length > 1) cmpWrite_(ss, 'FBS Офисы', offices, log);

  var service = cmpBuildService_(ss, log);
  cmpWrite_(ss, 'FBS Служебное', service, log);
  cmpHideRaw_(ss, log);
  try { wbEnsureFbsFilters_(ss, log); } catch (e) {
    log.push(['Фильтры', String(e)]);
  }
  scheduleComposeFbsCopy_();
  var url = ss.getUrl();
  Logger.log(log.map(function(r) { return r.join('\t'); }).join('\n'));
  if (!silent) {
    try {
      SpreadsheetApp.getUi().alert(
        'Собрано в сырой файл FBS. Копия в главную книгу через ~20 сек.\n' +
        'Поставки ' + Math.max(0, supplies.length - 1) +
        ', заказы ' + Math.max(0, orders.length - 1) +
        ', состав ' + Math.max(0, composition.length - 1) + '.\n\n' + url
      );
    } catch (e) {
      Logger.log('UI недоступен: ' + e);
    }
  }
  return url;
}

function unhideRawFbsSheets() {
  var ss = fbsBook_();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    try { sheets[i].showSheet(); } catch (e) {}
  }
}

function cmpBuildSupplies_(ss, log) {
  var raw = cmpReadMaps_(ss, 'FBS_Поставки');
  var qr = cmpIndexBy_(cmpReadMaps_(ss, 'FBS_Поставки_QR').rows, ['supplyId', 'id']);
  var boxes = cmpGroupBy_(cmpReadMaps_(ss, 'FBS_Поставки_короба').rows, ['supplyId']);
  var stickers = cmpIndexBy_(cmpReadMaps_(ss, 'FBS_Поставки_стикеры_коробов').rows, ['supplyId']);

  log.push(['Сырые поставки', String(raw.rows.length)]);
  log.push(['QR / короба / стикеры коробов', Object.keys(qr).length + ' / ' + Object.keys(boxes).length + ' / ' + Object.keys(stickers).length]);

  var rows = [CMP_SUPPLY_COLS.slice()];
  for (var i = 0; i < raw.rows.length; i++) {
    var s = raw.rows[i];
    var sid = String(cmpGet_(s, ['id', 'supplyId', 'ID поставки', 'supplyID', 'incomeID']) || '');
    if (!sid) continue;
    var boxList = boxes[sid] || [];
    var trbxIds = [];
    for (var b = 0; b < boxList.length; b++) {
      var tid = cmpGet_(boxList[b], ['trbxId', 'id']);
      if (tid !== '') trbxIds.push(String(tid));
    }
    var qrRow = qr[sid] || {};
    var stRow = stickers[sid] || {};
    var rec = {
      id: sid,
      name: cmpGet_(s, ['name']),
      done: cmpGet_(s, ['done']),
      done_name: cmpGet_(s, ['done_name']),
      isB2b: cmpGet_(s, ['isB2b']),
      createdAt: cmpGet_(s, ['createdAt']),
      closedAt: cmpGet_(s, ['closedAt']),
      scanDt: cmpGet_(s, ['scanDt']),
      cargoType: cmpGet_(s, ['cargoType']),
      cargoType_name: cmpGet_(s, ['cargoType_name']),
      crossBorderType: cmpGet_(s, ['crossBorderType']),
      crossBorderType_name: cmpGet_(s, ['crossBorderType_name']),
      destinationOfficeId: cmpGet_(s, ['destinationOfficeId']),
      ordersCount: cmpGet_(s, ['ordersCount']),
      trbxCount: trbxIds.length || cmpGet_(s, ['trbxCount']),
      trbxId: trbxIds.join(', '),
      barcode: cmpGet_(qrRow, ['barcode']) || cmpGet_(stRow, ['barcode']) || sid,
      rejectDt: cmpGet_(s, ['rejectDt']),
      recommendedWh: cmpGet_(s, ['recommendedWh']),
      isPickupPointShipment: cmpGet_(s, ['isPickupPointShipment'])
    };
    rows.push(cmpObjToRow_(rec, CMP_SUPPLY_COLS));
  }
  return rows;
}

function cmpBuildOrders_(ss, log) {
  var main = cmpReadMaps_(ss, 'FBS_Заказы');
  var news = cmpReadMaps_(ss, 'FBS_Заказы_новые');
  var statuses = cmpIndexBy_(cmpReadMaps_(ss, 'FBS_Заказы_статусы').rows, ['id']);
  var stickers = cmpIndexBy_(cmpReadMaps_(ss, 'FBS_Заказы_стикеры').rows, ['orderId', 'id']);
  var bridge = cmpReadMaps_(ss, 'FBS_Поставки_заказы').rows;

  var byId = {};
  var order = [];
  function addAll(list, fallbackSource) {
    for (var i = 0; i < list.length; i++) {
      var id = String(cmpGet_(list[i], ['id', 'orderId', 'orderID']) || '');
      if (!id) continue;
      if (!byId[id]) {
        byId[id] = list[i];
        if (fallbackSource && !cmpGet_(byId[id], ['source'])) byId[id].source = fallbackSource;
        order.push(id);
      }
    }
  }
  addAll(main.rows, 'orders_30d');
  addAll(news.rows, 'new');

  var supplyByOrder = {};
  for (var b = 0; b < bridge.length; b++) {
    var oid = String(cmpGet_(bridge[b], ['orderId', 'id']) || '');
    var sid = cmpGet_(bridge[b], ['supplyId']);
    if (oid && sid !== '') supplyByOrder[oid] = sid;
    if (oid && !byId[oid]) {
      byId[oid] = { id: oid, supplyId: sid, source: 'supply_order_ids' };
      order.push(oid);
    }
  }

  log.push(['Заказы журнал / новые / статусы / стикеры',
    main.rows.length + ' / ' + news.rows.length + ' / ' + Object.keys(statuses).length + ' / ' + Object.keys(stickers).length]);
  log.push(['Заказов после склейки', String(order.length)]);

  var calc = cmpLoadOrderCalcMaps_(ss);
  log.push(['Справочники артикул / вес / регион / склады',
    Object.keys(calc.nmIdToArticle).length + ' / ' +
    Object.keys(calc.weightByArticle).length + ' / ' +
    Object.keys(calc.warehouseToRegion).length + ' / ' +
    Object.keys(calc.warehouseIdToName).length]);

  var rows = [CMP_ORDER_COLS.slice()];
  for (var i = 0; i < order.length; i++) {
    var id = order[i];
    var o = byId[id] || {};
    var st = statuses[id] || {};
    var sk = stickers[id] || {};
    var supplier = cmpGet_(o, ['supplierStatus']) || cmpGet_(st, ['supplierStatus']);
    var wb = cmpGet_(o, ['wbStatus']) || cmpGet_(st, ['wbStatus']);
    var nmId = cmpGet_(o, ['nmId', 'nmID']);
    var sellerArticle = cmpGet_(o, ['article']);
    var createdAt = cmpGet_(o, ['createdAt']);
    var warehouseId = cmpGet_(o, ['warehouseId']);
    var officeId = cmpGet_(o, ['officeId']);
    var calcs = cmpCalcOrderRow_(calc, nmId, sellerArticle, createdAt, warehouseId, officeId, o);
    var rec = {
      id: id,
      rid: cmpGet_(o, ['rid']),
      createdAt: createdAt,
      article: sellerArticle,
      nmId: nmId,
      chrtId: cmpGet_(o, ['chrtId']),
      skus: cmpGet_(o, ['skus']),
      warehouseId: warehouseId,
      officeId: officeId,
      supplyId: cmpGet_(o, ['supplyId']) || supplyByOrder[id] || '',
      supplierStatus: supplier,
      supplierStatus_name: cmpGet_(o, ['supplierStatus_name']) || cmpGet_(st, ['supplierStatus_name']) || (CMP_SUPPLIER[supplier] || ''),
      wbStatus: wb,
      wbStatus_name: cmpGet_(o, ['wbStatus_name']) || cmpGet_(st, ['wbStatus_name']) || (CMP_WB[wb] || ''),
      price_rub: cmpGet_(o, ['price_rub']),
      finalPrice_rub: cmpGet_(o, ['finalPrice_rub']),
      salePrice_rub: cmpGet_(o, ['salePrice_rub']),
      deliveryType: cmpGet_(o, ['deliveryType']),
      cargoType_name: cmpGet_(o, ['cargoType_name']),
      isB2b: cmpGet_(o, ['isB2b']),
      sticker_partA: cmpGet_(sk, ['partA']),
      sticker_partB: cmpGet_(sk, ['partB']),
      sticker_barcode: cmpGet_(sk, ['barcode']),
      source: cmpGet_(o, ['source']),
      comment: cmpGet_(o, ['comment']),
      isCancellable: cmpGet_(o, ['isCancellable']) || cmpGet_(st, ['isCancellable']),
      orderUid: cmpGet_(o, ['orderUid']),
      scanPrice: cmpGet_(o, ['scanPrice']),
      'Артикул': calcs.article,
      'Дата': calcs.date,
      'Вес, кг за 1 шт': calcs.weightPerUnit,
      'Вес, кг за 1 шт сумма': calcs.totalWeight,
      'Регион поставки': calcs.region
    };
    rows.push(cmpObjToRow_(rec, CMP_ORDER_COLS));
  }
  return rows;
}

function cmpBuildComposition_(ss, supplyRows, orderRows, log) {
  var bridge = cmpReadMaps_(ss, 'FBS_Поставки_заказы').rows;
  var supplyMap = cmpIndexBy_(cmpRowsToMaps_(supplyRows), ['id']);
  var orderMap = cmpIndexBy_(cmpRowsToMaps_(orderRows), ['id']);

  var pairs = [];
  var seen = {};
  function addPair(sid, oid) {
    sid = String(sid || '');
    oid = String(oid || '');
    if (!sid || !oid) return;
    var key = sid + '|' + oid;
    if (seen[key]) return;
    seen[key] = true;
    pairs.push({ supplyId: sid, orderId: oid });
  }

  for (var i = 0; i < bridge.length; i++) {
    addPair(cmpGet_(bridge[i], ['supplyId']), cmpGet_(bridge[i], ['orderId', 'id']));
  }

  var orders = cmpRowsToMaps_(orderRows);
  for (var o = 0; o < orders.length; o++) {
    addPair(cmpGet_(orders[o], ['supplyId']), cmpGet_(orders[o], ['id']));
  }

  log.push(['Строк состава поставки', String(pairs.length)]);

  var rows = [CMP_COMP_COLS.slice()];
  for (var p = 0; p < pairs.length; p++) {
    var sid = pairs[p].supplyId;
    var oid = pairs[p].orderId;
    var s = supplyMap[sid] || {};
    var ord = orderMap[oid] || {};
    var rec = {
      supplyId: sid,
      supplyName: cmpGet_(s, ['name']),
      done_name: cmpGet_(s, ['done_name']),
      trbxId: cmpGet_(s, ['trbxId']),
      supplyCreatedAt: cmpGet_(s, ['createdAt']),
      supplyClosedAt: cmpGet_(s, ['closedAt']),
      orderId: oid,
      article: cmpGet_(ord, ['article']),
      nmId: cmpGet_(ord, ['nmId']),
      chrtId: cmpGet_(ord, ['chrtId']),
      skus: cmpGet_(ord, ['skus']),
      supplierStatus_name: cmpGet_(ord, ['supplierStatus_name']),
      wbStatus_name: cmpGet_(ord, ['wbStatus_name']),
      price_rub: cmpGet_(ord, ['price_rub']),
      finalPrice_rub: cmpGet_(ord, ['finalPrice_rub']),
      orderCreatedAt: cmpGet_(ord, ['createdAt']),
      warehouseId: cmpGet_(ord, ['warehouseId']),
      есть_в_заказах: cmpGet_(ord, ['article']) !== '' || cmpGet_(ord, ['nmId']) !== '' ? 'Да' : 'Нет',
      'Артикул': cmpGet_(ord, ['Артикул']),
      'Дата': cmpGet_(ord, ['Дата']),
      'Вес, кг за 1 шт': cmpGet_(ord, ['Вес, кг за 1 шт']),
      'Вес, кг за 1 шт сумма': cmpGet_(ord, ['Вес, кг за 1 шт сумма'])
    };
    rows.push(cmpObjToRow_(rec, CMP_COMP_COLS));
  }
  return rows;
}

function cmpBuildService_(ss, log) {
  var out = [['раздел', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8']];
  var width = 9;

  function addSection(title, table) {
    if (!table || table.length < 2) return;
    out.push([title]);
    for (var r = 0; r < table.length; r++) {
      var row = [r === 0 ? '  заголовок' : ''].concat(table[r]);
      while (row.length < width) row.push('');
      if (row.length > width) {
        width = row.length;
        for (var i = 0; i < out.length; i++) {
          while (out[i].length < width) out[i].push('');
        }
      }
      out.push(row);
    }
    out.push([]);
  }

  addSection('Сборка ' + new Date().toISOString(), log);
  return out;
}

function cmpReadMaps_(ss, name) {
  var table = cmpReadTable_(ss, [name]);
  if (!table || table.length < 2) return { name: name, headers: [], rows: [] };
  var headers = table[0].map(function(h) { return String(h || ''); });
  var rows = [];
  for (var i = 1; i < table.length; i++) {
    var obj = {};
    var empty = true;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = table[i][c];
      if (table[i][c] !== '' && table[i][c] !== null && table[i][c] !== undefined) empty = false;
    }
    if (!empty) rows.push(obj);
  }
  return { name: name, headers: headers, rows: rows };
}

function cmpReadTable_(ss, names) {
  var table = wbReadSheet_(ss, names);
  return table && table.length ? table : null;
}

function cmpRowsToMaps_(table) {
  if (!table || table.length < 2) return [];
  var headers = table[0];
  var out = [];
  for (var i = 1; i < table.length; i++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = table[i][c];
    out.push(obj);
  }
  return out;
}

function cmpIndexBy_(rows, keys) {
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var id = String(cmpGet_(rows[i], keys) || '');
    if (id) map[id] = rows[i];
  }
  return map;
}

function cmpGroupBy_(rows, keys) {
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var id = String(cmpGet_(rows[i], keys) || '');
    if (!id) continue;
    if (!map[id]) map[id] = [];
    map[id].push(rows[i]);
  }
  return map;
}

function cmpGet_(obj, names) {
  if (!obj) return '';
  if (typeof wbPickObj_ === 'function') return wbPickObj_(obj, names);
  if (typeof names === 'string') names = [names];
  var keys = Object.keys(obj);
  var lower = {};
  for (var k = 0; k < keys.length; k++) lower[String(keys[k]).toLowerCase()] = obj[keys[k]];

  for (var i = 0; i < names.length; i++) {
    var want = String(names[i]);
    var v = obj[want];
    if (v === undefined) v = lower[want.toLowerCase()];
    if (v === undefined || v === null || v === '') continue;
    return v;
  }
  for (var j = 0; j < names.length; j++) {
    var v2 = obj[names[j]];
    if (v2 === undefined) v2 = lower[String(names[j]).toLowerCase()];
    if (v2 === 0 || v2 === false) return v2;
  }
  return '';
}

function cmpObjToRow_(obj, cols) {
  var row = [];
  for (var i = 0; i < cols.length; i++) row.push(obj[cols[i]] !== undefined && obj[cols[i]] !== null ? obj[cols[i]] : '');
  return row;
}

function cmpLoadOrderCalcMaps_(ss) {
  var main = ss;
  try { main = SpreadsheetApp.getActiveSpreadsheet() || ss; } catch (e) {}
  var maps = {
    nmIdToArticle: {},
    weightByArticle: {},
    warehouseToRegion: {},
    warehouseIdToName: {},
    warehouseIdToOfficeId: {},
    officeById: {}
  };

  var products = cmpReadSheetValues_(main, 'Товары');
  if (products.length > 1) {
    var pH = products[0];
    var nmCol = cmpHeaderIndex_(pH, ['nmID', 'nmId', 'nm_id', 'Артикул WB'], 0);
    var artCol = cmpHeaderIndex_(pH, ['Ваш артикул', 'Арт. продавца', 'vendorCode'], 2);
    for (var i = 1; i < products.length; i++) {
      var nmId = cmpNormKey_(products[i][nmCol]);
      if (!nmId) continue;
      maps.nmIdToArticle[nmId] = products[i][artCol] || '';
    }
  }

  var weights = cmpReadSheetValues_(main, 'Товар+Вес');
  if (weights.length > 1) {
    var wH = weights[0];
    var artWCol = cmpHeaderIndex_(wH, ['Артикул', 'артикул'], 0);
    var weightCol = -1;
    for (var c = 0; c < wH.length; c++) {
      if (String(wH[c] || '') === 'Вес, кг за 1 шт') {
        weightCol = c;
        break;
      }
    }
    if (weightCol < 0) weightCol = cmpHeaderIndex_(wH, ['Вес, кг за 1 шт'], 1);
    for (var r = 1; r < weights.length; r++) {
      var art = cmpNormKey_(weights[r][artWCol]);
      if (!art) continue;
      var wgt = parseFloat(weights[r][weightCol]);
      maps.weightByArticle[art] = isNaN(wgt) ? 0 : wgt;
    }
  }

  var regions = cmpReadSheetValues_(main, 'Регион поставки');
  if (regions.length > 1) {
    var rH = regions[0];
    var whCol = cmpHeaderIndex_(rH, ['Склад'], 0);
    var regCol = cmpHeaderIndex_(rH, ['Регион поставки'], 1);
    for (var g = 1; g < regions.length; g++) {
      var whName = cmpNormKey_(regions[g][whCol]);
      if (!whName) continue;
      maps.warehouseToRegion[whName] = regions[g][regCol] || '';
    }
  }

  var whRows = cmpReadMaps_(ss, 'FBS_Склады').rows;
  for (var w = 0; w < whRows.length; w++) {
    var wid = cmpNormKey_(cmpGet_(whRows[w], ['id', 'ID', 'warehouseId']));
    if (!wid) continue;
    maps.warehouseIdToName[wid] = cmpGet_(whRows[w], ['name']) || '';
    maps.warehouseIdToOfficeId[wid] = cmpGet_(whRows[w], ['officeId']) || '';
  }

  var ofRows = cmpReadMaps_(ss, 'FBS_Офисы').rows;
  for (var o = 0; o < ofRows.length; o++) {
    var oid = cmpNormKey_(cmpGet_(ofRows[o], ['id', 'officeId']));
    if (!oid) continue;
    maps.officeById[oid] = {
      name: cmpGet_(ofRows[o], ['name']) || '',
      city: cmpGet_(ofRows[o], ['city']) || '',
      federalDistrict: cmpGet_(ofRows[o], ['federalDistrict']) || ''
    };
  }

  return maps;
}

function cmpCalcOrderRow_(maps, nmId, sellerArticle, createdAt, warehouseId, officeId, raw) {
  var article = maps.nmIdToArticle[cmpNormKey_(nmId)] || '';
  var dateVal = '';
  if (typeof wbDate_ === 'function') dateVal = wbDate_(createdAt);
  else dateVal = createdAt || '';

  var qty = Number(cmpGet_(raw, ['qty', 'quantity']) || 1);
  if (!qty) qty = 1;

  var weightPerUnit = cmpWeightByArticle_(maps.weightByArticle, sellerArticle);
  if (!weightPerUnit) weightPerUnit = cmpWeightByArticle_(maps.weightByArticle, article);

  var region = cmpRegionForOrder_(maps, warehouseId, officeId);

  return {
    article: article,
    date: dateVal,
    weightPerUnit: weightPerUnit,
    totalWeight: weightPerUnit * qty,
    region: region
  };
}

function cmpWeightByArticle_(weightByArticle, article) {
  var key = cmpNormKey_(article);
  if (!key) return 0;
  var w = weightByArticle[key];
  return w ? w : 0;
}

function cmpRegionForOrder_(maps, warehouseId, officeId) {
  var whKey = cmpNormKey_(warehouseId);
  var whName = maps.warehouseIdToName[whKey] || '';
  var region = cmpRegionFromName_(whName, maps.warehouseToRegion);
  if (region) return region;

  var oid = cmpNormKey_(officeId) || maps.warehouseIdToOfficeId[whKey] || '';
  var office = maps.officeById[oid];
  if (!office) return '';
  region = cmpRegionFromName_(office.name, maps.warehouseToRegion);
  if (region) return region;
  region = cmpRegionFromName_(office.city, maps.warehouseToRegion);
  if (region) return region;
  return office.federalDistrict || office.city || '';
}

function cmpRegionFromName_(name, warehouseToRegion) {
  if (!name) return '';
  var key = cmpNormKey_(name);
  if (warehouseToRegion[key]) return warehouseToRegion[key];
  for (var k in warehouseToRegion) {
    if (!k) continue;
    if (key.indexOf(k) !== -1 || k.indexOf(key) !== -1) return warehouseToRegion[k];
  }
  return '';
}

function cmpReadSheetValues_(ss, name) {
  if (!ss) return [];
  try {
    var sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 0) return sheet.getDataRange().getValues();
  } catch (e) {}
  return [];
}

function cmpHeaderIndex_(headers, names, fallback) {
  if (typeof wbColIdx_ === 'function') {
    var found = wbColIdx_(headers, names);
    if (found >= 0) return found;
  }
  for (var i = 0; i < names.length; i++) {
    var want = String(names[i]).toLowerCase();
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c] || '').toLowerCase() === want) return c;
    }
  }
  return fallback;
}

function cmpNormKey_(value) {
  if (value === undefined || value === null || value === '') return '';
  return String(value).trim().toLowerCase();
}

function cmpEnrichStocks_(ss, table, log) {
  if (!table || table.length < 2) return table;
  var calc = cmpLoadOrderCalcMaps_(ss);
  var srcHeaders = table[0];
  var extra = ['Артикул', 'Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  var headers = srcHeaders.slice();
  for (var e = 0; e < extra.length; e++) {
    if (headers.indexOf(extra[e]) < 0) headers.push(extra[e]);
  }
  var col = {};
  for (var c = 0; c < headers.length; c++) col[String(headers[c])] = c;
  var vendorCol = cmpHeaderIndex_(srcHeaders, ['vendorCode', 'article', 'Арт. продавца', 'Арт. в заказе'], -1);
  var nmCol = cmpHeaderIndex_(srcHeaders, ['nmID', 'nmId', 'nm_id', 'Артикул WB'], -1);
  var amountCol = cmpHeaderIndex_(srcHeaders, ['amount'], -1);
  var out = [headers];
  for (var i = 1; i < table.length; i++) {
    var row = table[i].slice();
    while (row.length < headers.length) row.push('');
    var vendor = vendorCol >= 0 ? row[vendorCol] : '';
    var nmId = nmCol >= 0 ? row[nmCol] : '';
    var amount = amountCol >= 0 ? Number(String(row[amountCol]).replace(',', '.')) || 0 : 0;
    var article = calc.nmIdToArticle[cmpNormKey_(nmId)] || vendor || '';
    var w = cmpWeightByArticle_(calc.weightByArticle, vendor);
    if (!w) w = cmpWeightByArticle_(calc.weightByArticle, article);
    row[col['Артикул']] = article;
    row[col['Вес, кг за 1 шт']] = w;
    row[col['Вес, кг за 1 шт сумма']] = w * amount;
    out.push(row);
  }
  if (log) log.push(['FBS Остатки расчёт', 'добавлены Артикул / вес']);
  return out;
}

function cmpFormatOrderCalcCols_(ss, name) {
  if (!ss) return;
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var last = sheet.getLastRow() - 1;
  var dateCol = headers.indexOf('Дата') + 1;
  if (dateCol > 0) sheet.getRange(2, dateCol, last, 1).setNumberFormat('dd.mm.yyyy');
  var numberCols = ['Вес, кг за 1 шт', 'Вес, кг за 1 шт сумма'];
  for (var i = 0; i < numberCols.length; i++) {
    var col = headers.indexOf(numberCols[i]) + 1;
    if (col > 0) sheet.getRange(2, col, last, 1).setNumberFormat('#,##0.00');
  }
}

function cmpWrite_(ss, name, rows, log) {
  wbWriteSheetSafe_(ss, name, rows, log);
}

function cmpHideRaw_(ss, log) {
  var hidden = wbHideSheets_(ss, CMP_HIDE);
  try { wbHideSheets_(fbsSidecar_(), CMP_HIDE); } catch (e) {}
  if (log) log.push(['Скрыты сырые листы', hidden.join(', ')]);
}

function cmpDropJunk_(ss, log) {
  var dropped = wbDeleteSheets_(ss, CMP_DROP);
  try { wbDeleteSheets_(fbsSidecar_(), CMP_DROP); } catch (e) {}
  if (log && dropped.length) log.push(['Удалены лишние листы', dropped.join(', ')]);
}

