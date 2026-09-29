// ============================================
// Русские названия колонок (алиасы + разовый rename)
// Не сливать date/rr_dt/createdAt с «Дата» и article с «Артикул».
// Отчёты (Вес продаж, К-во заказов…) свои шапки не меняют.
// ============================================

var WB_RU = {
  rrd_id: 'ID строки',
  nm_id: 'Артикул WB',
  nmId: 'Артикул WB',
  nmID: 'Артикул WB',
  sa_name: 'Арт. продавца',
  supplierArticle: 'Арт. продавца',
  vendorCode: 'Арт. продавца',
  barcode: 'Штрихкод',
  doc_type_name: 'Тип документа',
  quantity: 'Кол-во',
  retail_amount: 'Сумма розн.',
  office_name: 'Склад',
  warehouseName: 'Склад',
  supplier_oper_name: 'Операция',
  rr_dt: 'Дата ISO',
  shk_id: 'ID ШК',
  commission_percent: 'кВВ %',
  site_country: 'Страна продажи',
  country: 'Страна продажи',
  delivery_rub: 'Логистика',
  deliveryService: 'Логистика',
  ppvz_for_pay: 'К Перечислению',
  forPay: 'К Перечислению',
  ppvz_kvw_prc: 'кВВ без НДС',
  ppvz_kvw_prc_base: 'кВВ база',
  advertId: 'ID кампании',
  appType: 'Площадка',
  name: 'Название',
  views: 'Показы',
  clicks: 'Клики',
  sum: 'Сумма API',
  atbs: 'Корзина',
  atbs: 'Корзина',
  orders: 'Заказы рек.',
  shks: 'Выкупы рек.',
  cpc: 'CPC',
  ctr: 'CTR %',
  cr: 'CR %',
  paymentType: 'Источник списания',
  lastChangeDate: 'Изменено',
  warehouseType: 'Тип склада',
  countryName: 'Страна',
  oblastOkrugName: 'Округ',
  regionName: 'Регион WB',
  category: 'Категория',
  subject: 'Предмет',
  subjectName: 'Предмет',
  brand: 'Бренд',
  techSize: 'Размер',
  incomeID: 'ID поставки',
  isSupply: 'Поставка?',
  isRealization: 'Реализация?',
  totalPrice: 'Цена без скидок',
  discountPercent: 'Скидка прод. %',
  spp: 'Скидка WB %',
  finishedPrice: 'Цена итог',
  finishedPrice: 'Цена итог',
  priceWithDisc: 'Цена со скидкой',
  priceWithDisc: 'Цена со скидкой',
  isCancel: 'Отмена?',
  cancelDate: 'Дата отмены',
  orderType: 'Тип заказа',
  sticker: 'Стикер',
  gNumber: '№ заказа',
  srid: 'ID заказа',
  supplyID: 'ID поставки',
  preorderID: 'ID заказа',
  createDate: 'Создана',
  supplyDate: 'План поставки',
  factDate: 'Факт',
  statusID: 'Код статуса',
  boxTypeID: 'Код тары',
  actualWarehouseName: 'Склад факт',
  color: 'Цвет',
  acceptedQuantity: 'Принято',
  readyForSaleQuantity: 'К продаже',
  updatedDate: 'Обновлено',
  warehouseId: 'ID склада',
  officeId: 'ID офиса',
  deliveryType: 'Код схемы',
  deliveryType_name: 'Схема',
  cargoType: 'Код габарита',
  cargoType_name: 'Габарит',
  chrtId: 'ID размера',
  chrtId: 'ID размера',
  amount: 'Остаток',
  sku: 'SKU',
  title: 'Название',
  wbSize: 'Размер WB',
  skus: 'Штрихкоды',
  address: 'Адрес',
  address_full: 'Адрес',
  address_full: 'Адрес',
  city: 'Город',
  federalDistrict: 'Округ',
  isDeleting: 'Удаляется',
  isProcessing: 'В обработке',
  selected: 'Выбран',
  longitude: 'Долгота',
  latitude: 'Широта',
  rid: 'RID',
  orderUid: 'UID заказа',
  createdAt: 'Создано',
  article: 'Арт. в заказе',
  offices: 'Офисы',
  supplyId: 'ID поставки FBS',
  crossBorderType: 'Код кроссбордер',
  crossBorderType_name: 'Кроссбордер',
  price: 'Цена API',
  price_rub: 'Цена, ₽',
  price_rub: 'Цена, ₽',
  finalPrice_rub: 'Цена итог, ₽',
  salePrice_rub: 'Цена продажи, ₽',
  supplyName: 'Название поставки',
  supplyCreatedAt: 'Создана поставка',
  supplyClosedAt: 'Закрыта поставка',
  orderCreatedAt: 'Создан заказ',
  qty: 'Кол-во',
  barcode: 'Штрихкод',
  finalPrice: 'Цена итог API',
  finalPrice_rub: 'Цена итог, ₽',
  convertedPrice: 'Цена конв.',
  convertedPrice_rub: 'Цена конв., ₽',
  convertedFinalPrice: 'Цена конв. итог',
  convertedFinalPrice_rub: 'Цена конв. итог, ₽',
  salePrice: 'Цена продажи API',
  salePrice_rub: 'Цена продажи, ₽',
  scanPrice: 'Цена скана',
  currencyCode: 'Валюта',
  convertedCurrencyCode: 'Валюта конв.',
  ddate: 'Дата доставки',
  sellerDate: 'Дата продавца',
  comment: 'Комментарий',
  supplierStatus: 'Статус прод. код',
  supplierStatus_name: 'Статус продавца',
  wbStatus: 'Статус WB код',
  wbStatus_name: 'Статус WB',
  isCancellable: 'Можно отменить',
  isCancellable: 'Можно отменить',
  done: 'Закрыта',
  done: 'Закрыта',
  done_name: 'Статус',
  isB2b: 'B2B',
  closedAt: 'Закрыта когда',
  closedAt: 'Закрыта когда',
  scanDt: 'Сканирование',
  scanDt: 'Сканирование',
  destinationOfficeId: 'ID офиса',
  ordersCount: 'Заданий',
  ordersCount: 'Заданий',
  trbxCount: 'Коробов',
  trbxCount: 'Коробов',
  source: 'Источник',
  isZeroOrder: 'Нулевой заказ',
  isZeroOrder: 'Нулевой заказ',
  isPickupPointShipmentAllowed: 'Самовывоз разрешён',
  isPickupPointShipmentAllowed: 'Самовывоз разрешён',
  colorCode: 'Код цвета',
  requiredMeta: 'Мета обяз.',
  optionalMeta: 'Мета доп.',
  stickerId: 'ID стикера',
  stickerId: 'ID стикера',
  sticker_partA: 'Стикер A',
  sticker_partB: 'Стикер B',
  sticker_barcode: 'Штрихкод стикера',
  parcel: 'Посылка',
  year: 'Год',
  month: 'Месяц',
  trbxId: 'ID короба',
  rejectDt: 'Отказ',
  recommendedWh: 'Реком. склад',
  isPickupPointShipment: 'Самовывоз',
  orderId: 'ID задания'
};

var WB_RU_ADS_EXTRA = { date: 'Дата статы' };
var WB_RU_ORD_EXTRA = { date: 'Дата-время API' };

var WB_ALIAS_INDEX = null;

function wbBuildAliasIndex_() {
  if (WB_ALIAS_INDEX) return WB_ALIAS_INDEX;
  var groups = [
    ['ID строки', 'rrd_id'],
    ['Артикул WB', 'nm_id', 'nmId', 'nmID'],
    ['Арт. продавца', 'sa_name', 'supplierArticle', 'vendorCode', 'Ваш артикул'],
    ['Штрихкод', 'barcode', 'Штрих-код'],
    ['Тип документа', 'doc_type_name'],
    ['Кол-во', 'quantity'],
    ['Сумма розн.', 'retail_amount'],
    ['Склад', 'office_name', 'warehouseName'],
    ['Операция', 'supplier_oper_name'],
    ['Дата ISO', 'rr_dt'],
    ['ID ШК', 'shk_id'],
    ['кВВ %', 'commission_percent'],
    ['Страна продажи', 'site_country', 'country'],
    ['Логистика', 'delivery_rub', 'deliveryService'],
    ['К Перечислению', 'К перечислению', 'ppvz_for_pay', 'forPay'],
    ['кВВ без НДС', 'ppvz_kvw_prc'],
    ['кВВ база', 'ppvz_kvw_prc_base'],
    ['ID кампании', 'advertId'],
    ['Дата статы', 'date'],
    ['Дата-время API', 'date'],
    ['Площадка', 'appType'],
    ['Название', 'name', 'title'],
    ['Показы', 'views'],
    ['Клики', 'clicks'],
    ['Сумма API', 'sum'],
    ['Корзина', 'atbs', 'atbs'],
    ['Заказы рек.', 'orders'],
    ['Выкупы рек.', 'shks'],
    ['CPC', 'cpc'],
    ['CTR %', 'ctr'],
    ['CR %', 'cr'],
    ['Источник списания', 'paymentType'],
    ['Изменено', 'lastChangeDate'],
    ['Тип склада', 'warehouseType'],
    ['Страна', 'countryName'],
    ['Округ', 'oblastOkrugName', 'federalDistrict'],
    ['Регион WB', 'regionName'],
    ['Категория', 'category'],
    ['Предмет', 'subject', 'subjectName'],
    ['Бренд', 'brand'],
    ['Размер', 'techSize'],
    ['Размер WB', 'wbSize'],
    ['ID поставки', 'supplyID', 'incomeID'],
    ['Поставка?', 'isSupply'],
    ['Реализация?', 'isRealization'],
    ['Цена без скидок', 'totalPrice'],
    ['Скидка прод. %', 'discountPercent'],
    ['Скидка WB %', 'spp'],
    ['Цена итог', 'finishedPrice', 'finishedPrice'],
    ['Цена со скидкой', 'priceWithDisc', 'priceWithDisc'],
    ['Отмена?', 'isCancel'],
    ['Дата отмены', 'cancelDate'],
    ['Тип заказа', 'orderType'],
    ['Стикер', 'sticker'],
    ['№ заказа', 'gNumber'],
    ['ID заказа', 'srid', 'preorderID'],
    ['Создана', 'createDate'],
    ['План поставки', 'supplyDate'],
    ['Факт', 'factDate'],
    ['Код статуса', 'statusID'],
    ['Код тары', 'boxTypeID'],
    ['Склад факт', 'actualWarehouseName'],
    ['Цвет', 'color'],
    ['Принято', 'acceptedQuantity'],
    ['К продаже', 'readyForSaleQuantity'],
    ['Обновлено', 'updatedDate'],
    ['ID склада', 'warehouseId'],
    ['ID офиса', 'officeId', 'destinationOfficeId'],
    ['Код схемы', 'deliveryType'],
    ['Схема', 'deliveryType_name'],
    ['Код габарита', 'cargoType'],
    ['Габарит', 'cargoType_name'],
    ['ID размера', 'chrtId', 'chrtId'],
    ['Остаток', 'amount', 'qty'],
    ['SKU', 'sku'],
    ['Штрихкоды', 'skus'],
    ['Адрес', 'address', 'address_full', 'address_full'],
    ['Город', 'city'],
    ['Удаляется', 'isDeleting'],
    ['В обработке', 'isProcessing'],
    ['Выбран', 'selected'],
    ['Долгота', 'longitude'],
    ['Широта', 'latitude'],
    ['RID', 'rid'],
    ['UID заказа', 'orderUid'],
    ['Создано', 'createdAt'],
    ['Арт. в заказе', 'article'],
    ['Офисы', 'offices'],
    ['ID поставки FBS', 'supplyId'],
    ['Код кроссбордер', 'crossBorderType'],
    ['Кроссбордер', 'crossBorderType_name'],
    ['Цена, ₽', 'price_rub', 'price_rub'],
    ['Цена итог, ₽', 'finalPrice_rub'],
    ['Цена продажи, ₽', 'salePrice_rub'],
    ['Статус прод. код', 'supplierStatus'],
    ['Статус продавца', 'supplierStatus_name'],
    ['Статус WB код', 'wbStatus'],
    ['Статус WB', 'wbStatus_name'],
    ['Можно отменить', 'isCancellable', 'isCancellable'],
    ['Закрыта', 'done', 'done'],
    ['B2B', 'isB2b'],
    ['Закрыта когда', 'closedAt', 'closedAt'],
    ['Сканирование', 'scanDt', 'scanDt'],
    ['Заданий', 'ordersCount', 'ordersCount'],
    ['Коробов', 'trbxCount', 'trbxCount'],
    ['Источник', 'source'],
    ['ID задания', 'id', 'orderId'],
    ['ID', 'id'],
    ['Название поставки', 'supplyName'],
    ['Создана поставка', 'supplyCreatedAt'],
    ['Закрыта поставка', 'supplyClosedAt'],
    ['Создан заказ', 'orderCreatedAt'],
    ['Нулевой заказ', 'isZeroOrder', 'isZeroOrder'],
    ['Самовывоз разрешён', 'isPickupPointShipmentAllowed', 'isPickupPointShipmentAllowed'],
    ['Код цвета', 'colorCode'],
    ['Мета обяз.', 'requiredMeta'],
    ['Мета доп.', 'optionalMeta'],
    ['ID стикера', 'stickerId', 'stickerId'],
    ['Стикер A', 'sticker_partA'],
    ['Стикер B', 'sticker_partB'],
    ['Штрихкод стикера', 'sticker_barcode'],
    ['Посылка', 'parcel'],
    ['Год', 'year'],
    ['Месяц', 'month'],
    ['ID короба', 'trbxId'],
    ['Отказ', 'rejectDt'],
    ['Реком. склад', 'recommendedWh'],
    ['Самовывоз', 'isPickupPointShipment']
  ];
  var idx = {};
  for (var g = 0; g < groups.length; g++) {
    var arr = groups[g];
    for (var i = 0; i < arr.length; i++) {
      var key = String(arr[i]).toLowerCase();
      if (!idx[key]) idx[key] = [];
      for (var j = 0; j < arr.length; j++) {
        if (idx[key].indexOf(arr[j]) === -1) idx[key].push(arr[j]);
      }
    }
  }
  WB_ALIAS_INDEX = idx;
  return idx;
}

function wbAliasList_(name) {
  var n = String(name == null ? '' : name).trim();
  if (!n) return [];
  var idx = wbBuildAliasIndex_();
  var found = idx[n.toLowerCase()];
  return found && found.length ? found.slice() : [n];
}

function wbColIdx_(headers, names) {
  if (!headers || !names) return -1;
  if (typeof names === 'string') names = [names];
  for (var n = 0; n < names.length; n++) {
    var al = wbAliasList_(names[n]);
    var want = {};
    for (var j = 0; j < al.length; j++) want[String(al[j]).toLowerCase()] = true;
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c] != null ? String(headers[c]).trim().toLowerCase() : '';
      if (h && want[h]) return c;
    }
  }
  return -1;
}

function wbIs_(header, name) {
  return wbColIdx_([header], [name]) === 0;
}

function wbPickObj_(obj, names) {
  if (!obj) return '';
  if (typeof names === 'string') names = [names];
  var keys = Object.keys(obj);
  var lower = {};
  for (var k = 0; k < keys.length; k++) lower[String(keys[k]).toLowerCase()] = obj[keys[k]];
  for (var n = 0; n < names.length; n++) {
    var al = wbAliasList_(names[n]);
    for (var j = 0; j < al.length; j++) {
      var key = al[j];
      var v = obj[key];
      if (v === undefined) v = lower[String(key).toLowerCase()];
      if (v === undefined || v === null || v === '') continue;
      return v;
    }
  }
  for (var n2 = 0; n2 < names.length; n2++) {
    var al2 = wbAliasList_(names[n2]);
    for (var j2 = 0; j2 < al2.length; j2++) {
      var v2 = obj[al2[j2]];
      if (v2 === undefined) v2 = lower[String(al2[j2]).toLowerCase()];
      if (v2 === 0 || v2 === false) return v2;
    }
  }
  return '';
}

function wbRowToRu_(row, extraMap) {
  if (!row) return row;
  var out = {};
  for (var k in row) {
    if (!Object.prototype.hasOwnProperty.call(row, k)) continue;
    var nk = (extraMap && extraMap[k]) || WB_RU[k] || k;
    out[nk] = row[k];
  }
  return out;
}

function wbMapForSheet_(sheetName) {
  var m = {};
  for (var k in WB_RU) m[k] = WB_RU[k];
  if (sheetName === 'Рекламные_расходы') {
    m.date = 'Дата статы';
  } else if (sheetName === 'Заказы') {
    m.date = 'Дата-время API';
  }
  if (/Заказы/.test(sheetName) && sheetName.indexOf('Поставка') === -1) {
    m.id = 'ID задания';
    m.orderId = 'ID задания';
  } else if (/Поставка_состав/.test(sheetName)) {
    m.id = 'ID задания';
    m.orderId = 'ID задания';
    m.supplyName = 'Название поставки';
    m.supplyCreatedAt = 'Создана поставка';
    m.supplyClosedAt = 'Закрыта поставка';
    m.orderCreatedAt = 'Создан заказ';
  } else if (/Поставки/.test(sheetName) && /FBS/.test(sheetName)) {
    m.id = 'ID поставки';
    m.name = 'Название';
    m.done_name = 'Статус';
  } else if (/Склады|Офисы/.test(sheetName)) {
    m.id = 'ID';
  }
  return m;
}

function wbClearFilter_(sheet) {
  try {
    var f = sheet.getFilter();
    if (f) f.remove();
  } catch (e) {}
}

function wbRenameSheetHeaders_(sheet, map) {
  if (!sheet || sheet.getLastColumn() < 1) return 0;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var have = {};
  var i;
  for (i = 0; i < headers.length; i++) have[String(headers[i] || '').trim()] = true;
  var n = 0;
  var next = headers.slice();
  for (i = 0; i < headers.length; i++) {
    var oldName = String(headers[i] || '').trim();
    if (!oldName || !map[oldName]) continue;
    var neu = map[oldName];
    if (!neu || neu === oldName) continue;
    if (have[neu]) continue;
    next[i] = neu;
    have[neu] = true;
    delete have[oldName];
    n++;
  }
  if (n) {
    wbClearFilter_(sheet);
    sheet.getRange(1, 1, 1, lastCol).setValues([next]);
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');
  }
  return n;
}

function renameWbColumnHeaders_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return 0;
  var names = [
    'Финансовые_ежедневные', 'Финансовые_недельные', 'Товары', 'Рекламные_расходы',
    'Заказы', 'В пути', 'Поставки_список', 'Детали_поставок', 'Товары_поставок',
    'FBS_Остатки', 'FBS Остатки', 'FBS_Склады', 'FBS Склады', 'FBS_Офисы', 'FBS Офисы',
    'FBS_Заказы', 'FBS Заказы', 'FBS_Заказы_новые', 'FBS_Заказы_статусы', 'FBS_Заказы_стикеры',
    'FBS_Поставки', 'FBS Поставки', 'FBS Поставка_состав', 'FBS_Поставки_заказы',
    'FBS Заказы', 'FBS Остатки', 'FBS Склады', 'FBS Офисы',
    'FBS_Заказы_стикеры', 'FBS_Поставки_короба', 'FBS_Поставки_QR', 'FBS_Поставки_стикеры_коробов'
  ];
  var total = 0;
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    try {
      total += wbRenameSheetHeaders_(sh, wbMapForSheet_(names[i]));
    } catch (e) {
      Logger.log('rename «' + names[i] + '»: ' + e);
    }
  }
  return total;
}

function renameWbHeadersAllBooks_() {
  var n = 0;
  try { n += renameWbColumnHeaders_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) {}
  try {
    if (typeof fbsSidecar_ === 'function') n += renameWbColumnHeaders_(fbsSidecar_());
  } catch (e2) {}
  return n;
}

function wbCell_(obj, header) {
  if (!obj) return '';
  if (Object.prototype.hasOwnProperty.call(obj, header)) {
    var direct = obj[header];
    if (direct !== undefined && direct !== null && direct !== '') return direct;
    if (direct === 0 || direct === false) return direct;
  }
  return wbPickObj_(obj, [header]);
}

function wbMapHeaders_(headers, sheetName) {
  if (!headers) return [];
  var map = wbMapForSheet_(sheetName || '');
  var out = [];
  var seen = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (!h) { out.push(h); continue; }
    var neu = map[h] || h;
    var key = String(neu).toLowerCase();
    if (seen[key] && neu !== h) neu = h;
    seen[String(neu).toLowerCase()] = true;
    out.push(neu);
  }
  return out;
}

function wbFillRow_(obj, headers) {
  var row = [];
  for (var j = 0; j < headers.length; j++) {
    var val = wbCell_(obj, headers[j]);
    if (val === null || val === undefined) row.push('');
    else if (typeof val === 'object' && !(val instanceof Date)) row.push(JSON.stringify(val));
    else row.push(val);
  }
  return row;
}

function wbIsApiDateField_(header) {
  var h = String(header || '');
  var names = [
    'createDate', 'supplyDate', 'factDate', 'updatedDate', 'createdAt', 'closedAt',
    'scanDt', 'rr_dt', 'lastChangeDate', 'cancelDate', 'rejectDt', 'sellerDate', 'ddate'
  ];
  for (var i = 0; i < names.length; i++) {
    if (wbIs_(h, names[i])) return true;
  }
  return /Date/i.test(h) || /дата/i.test(h);
}

function русскиеНазванияКолонок() {
  var n = renameWbHeadersAllBooks_();
  try {
    SpreadsheetApp.getUi().alert('Переименовано заголовков: ' + n + '\nПовторный запуск безопасен: уже русские не трогает.');
  } catch (e) {
    Logger.log('rename headers: ' + n);
  }
  return n;
}
