// ============================================
// ГЛОБАЛЬНЫЕ КОНСТАНТЫ ОТЧЁТА "ПОСТАВКА"
// ============================================
const FIXED_COLS = 2;
const BASE_METRICS_COUNT = 9; // 5 базовых + 4 расчётных
const REGION_START_COL = FIXED_COLS + BASE_METRICS_COUNT + 1; // = 12 (Колонка L)
const WAREHOUSE_METRICS_COUNT = 8;
const BLOCK_SIZE = 8;

// ============================================
// ГЛОБАЛЬНАЯ ФУНКЦИЯ НОРМАЛИЗАЦИИ АРТИКУЛА
// ============================================
function normalizeArt(art) {
  if (!art) return '';
  let result = art.toString().toUpperCase();
  result = result.replace(/Х/g, 'X');
  return result.trim();
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ СБОРА ДАННЫХ
// ============================================
function loadGoodsNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const goodsMap = {};
  function extract(sheetName, artKeyword, nameKeyword) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    let aIdx = -1, nIdx = -1;
    if (typeof wbColIdx_ === 'function') {
      aIdx = wbColIdx_(headers, [artKeyword]);
      nIdx = nameKeyword ? wbColIdx_(headers, [nameKeyword]) : -1;
    }
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]?.toString().trim().toLowerCase();
      if (aIdx === -1 && h.includes(artKeyword.toLowerCase())) aIdx = i;
      if (nIdx === -1 && nameKeyword && h.includes(nameKeyword.toLowerCase())) nIdx = i;
    }
    if (aIdx === -1) return;
    for (let i = 1; i < data.length; i++) {
      const art = data[i][aIdx];
      const name = nIdx !== -1 ? data[i][nIdx] : '';
      if (!art) continue;
      const norm = normalizeArt(art);
      if (!goodsMap[norm] && name) goodsMap[norm] = name.toString().trim();
    }
  }
  // extract("Финансовые", "Артикул", "Наименование");
  // extract("Поставки", "Артикул", "Наименование");
  // extract("Заказы", "supplierArticle", "Артикул");
  // extract("Остатки", "Артикул продавца", "Название");

  extract("Финансовые_ежедневные", "sa_name", "Имя товара");
  extract("Поставки", "Артикул", "Наименование");
  extract("Заказы", "supplierArticle", "Артикул");
  extract("Остатки", "Артикул продавца", "Название");
  extract("Товары", "Ваш артикул", "Название"); // 🔥 НОВЫЙ: подстраховка

  const refSheet = ss.getSheetByName("Справочник соответствий");
  if (refSheet) {
    const data = refSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const art = data[i][0], name = data[i][1];
      if (art && name) {
        const norm = normalizeArt(art);
        if (!goodsMap[norm]) goodsMap[norm] = name.toString().trim();
      }
    }
  }
  return goodsMap;
}

function getUniqueRegionWarehousePairs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName("Регион поставки");
  const allRefWarehouses = new Map();
  const approvedWh = new Map();

  if (refSheet) {
    const refData = refSheet.getDataRange().getValues();
    if (refData.length > 1) {
      const refHeaders = refData[0];
      let whIdx = -1, regIdx = -1, flagIdx = -1;
      for (let i = 0; i < refHeaders.length; i++) {
        const h = refHeaders[i].toString().trim().toLowerCase();
        if (h === 'склад' || h === 'warehouse') whIdx = i;
        if (h === 'регион' || h === 'регион поставки') regIdx = i;
        if (h === 'выбор склада для анализа') flagIdx = i;
      }
      if (whIdx !== -1) {
        for (let i = 1; i < refData.length; i++) {
          const wh = refData[i][whIdx];
          const reg = regIdx !== -1 ? refData[i][regIdx] : '';
          const flag = flagIdx !== -1 ? refData[i][flagIdx] : '';
          if (wh) {
            const whNorm = wh.toString().trim().toLowerCase();
            if (reg) allRefWarehouses.set(whNorm, reg.toString().trim());
            if (flag.toString().trim().toLowerCase() === 'да') approvedWh.set(whNorm, wh.toString().trim());
          }
        }
      }
    }
  }

  const uniquePairs = new Map();

  // 🔥 ДОБАВЛЕНО: filterCol и filterVal только для Финансовых
  const sources = [
    { sheet: "Финансовые_ежедневные", region: "Регион", warehouse: "office_name", filterCol: "doc_type_name", filterVal: "Продажа" },
    { sheet: "Поставки", region: "Регион поставки", warehouse: "Склад" },
    { sheet: "Заказы", region: "Регион поставки", warehouse: "warehouseName" }
  ];

  sources.forEach(src => {
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    let rIdx = -1, wIdx = -1, fIdx = -1;
    if (typeof wbColIdx_ === 'function') {
      rIdx = wbColIdx_(headers, [src.region]);
      wIdx = wbColIdx_(headers, [src.warehouse]);
      if (src.filterCol) fIdx = wbColIdx_(headers, [src.filterCol]);
    }
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]?.toString().trim().toLowerCase();
      if (rIdx === -1 && h === (src.region || '').toLowerCase()) rIdx = i;
      if (wIdx === -1 && h === src.warehouse.toLowerCase()) wIdx = i;
      if (fIdx === -1 && src.filterCol && h === src.filterCol.toLowerCase()) fIdx = i;
    }
    if (wIdx === -1) return;

    for (let i = 1; i < data.length; i++) {
      const whRaw = data[i][wIdx];
      if (!whRaw) continue;

      // 🔥 ФИЛЬТР: если задана колонка фильтра, проверяем значение
      if (fIdx !== -1) {
        const docType = String(data[i][fIdx] || '').trim();
        if (docType !== src.filterVal) continue; // ← Пропускаем всё, кроме "Продажа"
      }

      const wh = whRaw.toString().trim();
      const whNorm = wh.toLowerCase();
      let reg = rIdx !== -1 ? (data[i][rIdx] || '').toString().trim() : '';
      if (!reg && allRefWarehouses.has(whNorm)) reg = allRefWarehouses.get(whNorm);
      if (!reg) reg = "Не указан";

      const key = `${reg}|${wh}`;
      if (!uniquePairs.has(key)) uniquePairs.set(key, { region: reg, warehouse: wh });
    }
  });

  // Добавляем склады с флагом "да" из справочника, если их нет в данных
  for (const [whNorm, originalName] of approvedWh.entries()) {
    let found = false;
    for (const val of uniquePairs.values()) {
      if (val.warehouse.toLowerCase() === whNorm) { found = true; break; }
    }
    if (!found && allRefWarehouses.has(whNorm)) {
      const region = allRefWarehouses.get(whNorm);
      const key = `${region}|${originalName}`;
      if (!uniquePairs.has(key)) uniquePairs.set(key, { region: region, warehouse: originalName });
    }
  }

  return Array.from(uniquePairs.values());
}

function getSupplyDataByRegionWarehouse(supplyDateFrom, supplyDateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suppliesSheet = ss.getSheetByName("Поставки");
  if (!suppliesSheet) return {};
  const suppliesData = suppliesSheet.getDataRange().getValues();
  const suppliesHeaders = suppliesData[0];
  let artIndex = -1, regionIndex = -1, warehouseIndex = -1, dateIndex = -1, statusIndex = -1, weightIndex = -1, quantityIndex = -1;
  for (let i = 0; i < suppliesHeaders.length; i++) {
    const header = suppliesHeaders[i];
    if (header === 'Артикул продавца') artIndex = i;
    if (header === 'Регион поставки') regionIndex = i;
    if (header === 'Склад') warehouseIndex = i;
    if (header === 'Дата поставки') dateIndex = i;
    if (header === 'Статус') statusIndex = i;
    if (header === 'Общий вес, кг') weightIndex = i;
    if (header === 'Количество') quantityIndex = i;
  }
  const supplyFromTime = supplyDateFrom.getTime();
  const supplyToTime = supplyDateTo.getTime();
  const result = {};
  function parseDate(dateVal) {
    if (!dateVal) return null;
    if (dateVal instanceof Date) return dateVal;
    if (typeof dateVal === 'string') {
      const parts = dateVal.split('.');
      if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return null;
  }
  for (let i = 1; i < suppliesData.length; i++) {
    const row = suppliesData[i];
    const artRaw = row[artIndex];
    if (!artRaw) continue;
    const art = normalizeArt(artRaw);
    const region = row[regionIndex];
    const warehouse = row[warehouseIndex];
    const dateVal = row[dateIndex];
    const status = row[statusIndex];
    const weight = parseFloat(row[weightIndex]) || 0;
    const quantity = (quantityIndex !== -1 ? parseFloat(row[quantityIndex]) : 1) || 0;
    if (!region || !warehouse) continue;
    const supplyDate = parseDate(dateVal);
    let inPeriod = false;
    if (supplyDate && !isNaN(supplyDate.getTime())) {
      const time = supplyDate.getTime();
      if (time >= supplyFromTime && time <= supplyToTime) inPeriod = true;
    }
    if (!inPeriod) continue;
    const key = `${art}|${region}|${warehouse}`;
    if (!result[key]) result[key] = { art: art, region: region, warehouse: warehouse, ordersQty: 0, inTransitQty: 0, inTransitWeight: 0, totalWeight: 0 };
    result[key].ordersQty += quantity;
    result[key].totalWeight += weight;
    if (status !== 'Принято') {
      result[key].inTransitQty += quantity;
      result[key].inTransitWeight += weight;
    }
  }
  return result;
}

function loadSalesData(fromTime, toTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const financeSheet = ss.getSheetByName("Финансовые_ежедневные");
  if (!financeSheet) return {};
  const salesData = {};
  const financeData = financeSheet.getDataRange().getValues();
  const financeHeaders = financeData[0];
  let artIndex = -1, nameIndex = -1, quantityIndex = -1, docTypeIndex = -1, dateIndex = -1;
  for (let i = 0; i < financeHeaders.length; i++) {
    const header = financeHeaders[i];
    if (typeof wbIs_ === 'function') {
      if (wbIs_(header, 'sa_name')) artIndex = i;
      if (header === 'Имя товара') nameIndex = i;
      if (wbIs_(header, 'quantity')) quantityIndex = i;
      if (wbIs_(header, 'doc_type_name')) docTypeIndex = i;
      if (header === 'Дата') dateIndex = i;
    } else {
      if (header === 'sa_name') artIndex = i;
      if (header === 'Имя товара') nameIndex = i;
      if (header === 'quantity') quantityIndex = i;
      if (header === 'doc_type_name') docTypeIndex = i;
      if (header === 'Дата') dateIndex = i;
    }
  }
  for (let i = 1; i < financeData.length; i++) {
    const row = financeData[i];
    const dt = row[dateIndex];
    if (!dt) continue;
    let dateObj = dt;
    if (typeof dt === 'string') {
      const parts = dt.split('.');
      if (parts.length === 3) dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
    }
    const time = dateObj.getTime();
    if (isNaN(time) || time < fromTime || time > toTime) continue;
    const artRaw = row[artIndex];
    if (!artRaw) continue;
    const art = normalizeArt(artRaw);
    const name = (row[nameIndex] ?? "").toString().trim().toUpperCase();
    const quantity = parseFloat(row[quantityIndex]) || 0;
    const docType = row[docTypeIndex];
    if (docType === 'Продажа') {
      const key = art + "|" + name;
      if (!salesData[key]) salesData[key] = { art: art, name: name, salesQuantity: 0 };
      salesData[key].salesQuantity += quantity;
    }
  }
  return salesData;
}

function parseDateStr(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
  const str = val.toString().trim();
  const m = str.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if (m) return new Date(m[3], m[2] - 1, m[1]);
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function getOrdersData(dateFrom, dateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Заказы");
  if (!sheet) return { byArt: {}, byRegionWarehouse: {} };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { byArt: {}, byRegionWarehouse: {} };
  const headers = data[0];
  let artIdx = -1, regIdx = -1, whIdx = -1, dateIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (typeof wbIs_ === 'function') {
      if (wbIs_(headers[i], 'supplierArticle')) artIdx = i;
      if (headers[i] === 'Регион поставки') regIdx = i;
      if (wbIs_(headers[i], 'warehouseName')) whIdx = i;
      if (headers[i] === 'Дата') dateIdx = i;
    } else {
      if (headers[i] === 'supplierArticle') artIdx = i;
      if (headers[i] === 'Регион поставки') regIdx = i;
      if (headers[i] === 'warehouseName') whIdx = i;
      if (headers[i] === 'Дата') dateIdx = i;
    }
  }
  const fromTime = dateFrom.getTime();
  const toTime = new Date(dateTo.getFullYear(), dateTo.getMonth(), dateTo.getDate(), 23, 59, 59).getTime();
  const result = { byArt: {}, byRegionWarehouse: {} };
  for (let i = 1; i < data.length; i++) {
    const rowDate = parseDateStr(data[i][dateIdx]);
    if (!rowDate) continue;
    const rowTime = rowDate.getTime();
    if (rowTime < fromTime || rowTime > toTime) continue;
    const art = normalizeArt(data[i][artIdx]);
    if (!art) continue;
    result.byArt[art] = (result.byArt[art] || 0) + 1;
    const region = regIdx !== -1 ? (data[i][regIdx] || '').toString().trim() : '';
    const warehouse = whIdx !== -1 ? (data[i][whIdx] || '').toString().trim() : '';
    if (region && warehouse) {
      const key = `${art}|${region}|${warehouse}`;
      result.byRegionWarehouse[key] = (result.byRegionWarehouse[key] || 0) + 1;
    }
  }
  return result;
}

function getStockData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Остатки");
  if (!sheet) return { byArt: {}, byRegionWarehouse: {} };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { byArt: {}, byRegionWarehouse: {} };
  const headers = data[0];
  let dIdx = -1, aIdx = -1, rIdx = -1, wIdx = -1, qIdx = -1, wtIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i]?.toString().trim().toLowerCase();
    if (h === 'дата') dIdx = i;
    if (h === 'артикул продавца') aIdx = i;
    if (h === 'регион поставки') rIdx = i;
    if (h === 'склад') wIdx = i;
    if (h === 'остаток') qIdx = i;
    if (h.includes('вес') || h.includes('общий вес')) wtIdx = i;
  }
  if (aIdx === -1 || dIdx === -1 || qIdx === -1) return { byArt: {}, byRegionWarehouse: {} };
  let latest = null;
  for (let i = 1; i < data.length; i++) {
    const dt = parseDateStr(data[i][dIdx]);
    if (dt && (!latest || dt > latest)) latest = dt;
  }
  if (!latest) return { byArt: {}, byRegionWarehouse: {} };
  const result = { byArt: {}, byRegionWarehouse: {} };
  const latestTime = latest.getTime();
  const parseNum = (val) => { if (val === null || val === undefined) return 0; return parseFloat(val.toString().replace(',', '.')) || 0; };
  for (let i = 1; i < data.length; i++) {
    const dt = parseDateStr(data[i][dIdx]);
    if (!dt || dt.getTime() !== latestTime) continue;
    const art = normalizeArt(data[i][aIdx]);
    if (!art) continue;
    const qty = parseNum(data[i][qIdx]);
    const weight = wtIdx !== -1 ? parseNum(data[i][wtIdx]) : 0;
    if (!result.byArt[art]) result.byArt[art] = { stockQty: 0, stockWeight: 0 };
    result.byArt[art].stockQty += qty;
    result.byArt[art].stockWeight += weight;
    const region = rIdx !== -1 ? (data[i][rIdx] || '').toString().trim() : '';
    const warehouse = wIdx !== -1 ? (data[i][wIdx] || '').toString().trim() : '';
    if (region && warehouse) {
      const key = `${art}|${region}|${warehouse}`;
      if (!result.byRegionWarehouse[key]) result.byRegionWarehouse[key] = { stockQty: 0, stockWeight: 0 };
      result.byRegionWarehouse[key].stockQty += qty;
      result.byRegionWarehouse[key].stockWeight += weight;
    }
  }
  return result;
}

// ============================================
// ОСНОВНАЯ ФУНКЦИЯ ОТЧЁТА "ПОСТАВКА"
// ============================================
function сформироватьОтчетПоставка() {
  const regionOrder = ["Склад WB", "Центральный", "Северо-Западный", "Приволжский", "Южный и Северо-Кавказский", "Уральский", "Дальневосточный и Сибирский", "Беларусь", "Казахстан", "Узбекистан", "Армения", "Грузия", "Ереван", "Не определён"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("Поставка");
  if (!reportSheet) { SpreadsheetApp.getUi().alert('❌ Лист "Поставка" не найден'); return; }

  // 🔥 ОЧИСТКА СТАРЫХ ПРАВИЛ УСЛОВНОГО ФОРМАТИРОВАНИЯ (чтобы не дублировались)
  reportSheet.setConditionalFormatRules([]);

  // Пачечное чтение A1:B5
  const abVals = reportSheet.getRange(1, 1, 5, 2).getValues();
  const financeDateFrom = new Date(abVals[0][0]);
  const financeDateTo = new Date(abVals[1][0]);
  const supplyDateFrom = new Date(abVals[2][0]);
  const supplyDateTo = new Date(abVals[3][0]);
  const daysCount = parseFloat(abVals[4][0]) || 30;
  if (isNaN(financeDateFrom.getTime()) || isNaN(financeDateTo.getTime()) || isNaN(supplyDateFrom.getTime()) || isNaN(supplyDateTo.getTime())) {
    SpreadsheetApp.getUi().alert('❌ Введите корректные даты в ячейки A1-A4'); return;
  }
  const dateDiff = Math.ceil((financeDateTo - financeDateFrom) / (1000 * 60 * 60 * 24)) + 1;

  const ordersData = getOrdersData(financeDateFrom, financeDateTo);
  const stockData = getStockData();
  const supplyDataByRegion = getSupplyDataByRegionWarehouse(supplyDateFrom, supplyDateTo);
  const regionWarehousePairs = getUniqueRegionWarehousePairs();
  const goodsNames = loadGoodsNames();

  // Индекс поставок по ключу art|region|warehouse и сумма «в пути» по артикулу
  const supplyByKey = {};
  const inTransitByArt = {};
  for (const k in supplyDataByRegion) {
    const s = supplyDataByRegion[k];
    const artN = normalizeArt(s.art);
    supplyByKey[`${artN}|${s.region}|${s.warehouse}`] = s;
    if (!inTransitByArt[artN]) inTransitByArt[artN] = { qty: 0, weight: 0 };
    inTransitByArt[artN].qty += s.inTransitQty;
    inTransitByArt[artN].weight += s.inTransitWeight;
  }

  const regionsMap = new Map();
  for (const pair of regionWarehousePairs) {
    if (!regionsMap.has(pair.region)) regionsMap.set(pair.region, []);
    regionsMap.get(pair.region).push(pair.warehouse);
  }
  const regions = Array.from(regionsMap.keys()).sort((a, b) => {
    const idxA = regionOrder.indexOf(a); const idxB = regionOrder.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b, 'ru');
  });

  // const BASE_METRICS = ['Заказы,\nшт', 'В пути,\nшт', 'В пути,\nкг', 'На складе ВБ,\nшт', 'На складе ВБ,\nкг', 'Ср. продажи,\nшт', 'Запас в днях', 'К-во на поставку', '✏️'];
  const BASE_METRICS = ['Заказы,\nшт', 'В пути,\nшт', 'В пути,\nкг', 'На складе ВБ,\nшт', 'Итого ВБ,\nшт', 'Ср. продажи,\nшт', 'Запас в днях', 'К-во на поставку', '✏️'];
  const BASE_COL_WIDTHS = [50, 50, 45, 45, 45, 45, 45, 45, 45];
  const WAREHOUSE_METRICS = ['Остатки, шт', 'Остатки, кг', 'Заказы, шт', '🚚', 'Ср. продажи, шт', 'Запас в днях', 'К-во на поставку', '✏️'];
  const WAREHOUSE_COL_WIDTHS = [50, 50, 45, 45, 45, 45, 45, 45];

  const metricsRow = ["Артикул продавца", "Наименование товара", ...BASE_METRICS];
  const regionRowForColH = [], subHeaderRowForColH = [];

  for (const region of regions) {
    const warehouses = regionsMap.get(region);
    regionRowForColH.push(region); subHeaderRowForColH.push("ИТОГО по региону");
    for (let i = 0; i < WAREHOUSE_METRICS_COUNT - 1; i++) { regionRowForColH.push(""); subHeaderRowForColH.push(""); }
    for (let i = 0; i < warehouses.length; i++) {
      regionRowForColH.push(""); subHeaderRowForColH.push(warehouses[i]);
      for (let j = 0; j < WAREHOUSE_METRICS_COUNT - 1; j++) { regionRowForColH.push(""); subHeaderRowForColH.push(""); }
    }
  }
  for (const region of regions) {
    const warehouses = regionsMap.get(region);
    for (let m = 0; m < WAREHOUSE_METRICS.length; m++) metricsRow.push(WAREHOUSE_METRICS[m]);
    for (let i = 0; i < warehouses.length; i++) for (let m = 0; m < WAREHOUSE_METRICS.length; m++) metricsRow.push(WAREHOUSE_METRICS[m]);
  }

  // reportSheet.getRange(4, 1, 20, reportSheet.getMaxColumns()).clear();
  // for (let row = 1; row <= 5; row++) {
  //   reportSheet.getRange(`A${row}`).setValue(savedValues[`A${row}`]);
  //   reportSheet.getRange(`B${row}`).setValue(savedValues[`B${row}`]);
  // }

  // Очищаем только реально использованный диапазон (не весь maxRows×maxCols)
  const oldLastRow = Math.max(reportSheet.getLastRow(), 7);
  const oldLastCol = Math.max(reportSheet.getLastColumn(), metricsRow.length);

  if (oldLastCol > 2) {
    reportSheet.getRange(4, 3, 3, oldLastCol - 2).clear();
  }
  const dataClearRows = oldLastRow - 6;
  if (dataClearRows > 0 && oldLastCol > 0) {
    reportSheet.getRange(7, 1, dataClearRows, oldLastCol).clear();
  }

  // Восстанавливаем A1:B5
  reportSheet.getRange(1, 1, 5, 2).setValues(abVals);

  // Хвосты справа от новой ширины
  if (metricsRow.length < oldLastCol) {
    reportSheet.getRange(4, metricsRow.length + 1, Math.max(1, oldLastRow - 3), oldLastCol - metricsRow.length).clear();
  }

  if (regionRowForColH.length > 0) reportSheet.getRange(4, REGION_START_COL, 1, regionRowForColH.length).setValues([regionRowForColH]);
  if (subHeaderRowForColH.length > 0) reportSheet.getRange(5, REGION_START_COL, 1, subHeaderRowForColH.length).setValues([subHeaderRowForColH]);
  reportSheet.getRange(6, 1, 1, metricsRow.length).setValues([metricsRow]);

  let startCol = REGION_START_COL;
  for (const region of regions) {
    const warehouses = regionsMap.get(region);
    const totalCols = WAREHOUSE_METRICS_COUNT + warehouses.length * WAREHOUSE_METRICS_COUNT;
    if (totalCols > 1 && startCol + totalCols - 1 <= metricsRow.length) try { reportSheet.getRange(4, startCol, 1, totalCols).merge(); } catch (e) { }
    startCol += totalCols;
  }
  startCol = REGION_START_COL;
  for (const region of regions) {
    const warehouses = regionsMap.get(region);
    if (WAREHOUSE_METRICS_COUNT > 1 && startCol + WAREHOUSE_METRICS_COUNT - 1 <= metricsRow.length) try { reportSheet.getRange(5, startCol, 1, WAREHOUSE_METRICS_COUNT).merge(); } catch (e) { }
    startCol += WAREHOUSE_METRICS_COUNT + warehouses.length * WAREHOUSE_METRICS_COUNT;
  }
  startCol = REGION_START_COL;
  for (const region of regions) {
    const warehouses = regionsMap.get(region); startCol += WAREHOUSE_METRICS_COUNT;
    for (let i = 0; i < warehouses.length; i++) {
      if (WAREHOUSE_METRICS_COUNT > 1 && startCol + WAREHOUSE_METRICS_COUNT - 1 <= metricsRow.length) try { reportSheet.getRange(5, startCol, 1, WAREHOUSE_METRICS_COUNT).merge(); } catch (e) { }
      startCol += WAREHOUSE_METRICS_COUNT;
    }
  }

  if (regionRowForColH.length > 0) reportSheet.getRange(4, REGION_START_COL, 2, metricsRow.length - REGION_START_COL + 1).setFontFamily("Calibri").setFontSize(9).setFontWeight("bold").setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");
  reportSheet.getRange(6, 1, 1, metricsRow.length).setFontFamily("Calibri").setFontSize(9).setFontWeight("bold").setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");
  for (let col = 1; col <= metricsRow.length; col++) {
    if (metricsRow[col - 1] === '🚚') reportSheet.getRange(6, col).setFontSize(15);
    if (metricsRow[col - 1] === '✏️') reportSheet.getRange(6, col).setFontSize(13);
  }
  reportSheet.setRowHeight(4, 30); reportSheet.setRowHeight(5, 30); reportSheet.setRowHeight(6, 45);

  reportSheet.setColumnWidth(1, 150); reportSheet.setColumnWidth(2, 350);
  for (let i = 0; i < BASE_METRICS_COUNT; i++) reportSheet.setColumnWidth(3 + i, BASE_COL_WIDTHS[i]);
  let cCol = REGION_START_COL;
  while (cCol <= metricsRow.length) {
    for (let k = 0; k < Math.min(WAREHOUSE_METRICS_COUNT, metricsRow.length - cCol + 1); k++) reportSheet.setColumnWidth(cCol + k, WAREHOUSE_COL_WIDTHS[k]);
    cCol += WAREHOUSE_METRICS_COUNT;
  }

  const allArts = new Map();
  function addArt(art, orig) { if (!art) return; const n = normalizeArt(art); if (!allArts.has(n)) allArts.set(n, orig || art); }
  for (const k in ordersData.byArt) addArt(k);
  for (const k in stockData.byArt) addArt(k);
  for (const k in supplyDataByRegion) addArt(supplyDataByRegion[k].art);
  const arts = Array.from(allArts.keys()).sort();
  const maxRows = 200;
  const displayArts = arts.length > maxRows ? arts.slice(0, maxRows) : arts;
  if (arts.length > maxRows) SpreadsheetApp.getUi().alert('⚠️ Найдено ' + arts.length + ' товаров. Выводим только ' + maxRows);

  const dataRows = [];
  const totalRow = new Array(metricsRow.length).fill(0);
  totalRow[0] = "ИТОГО"; totalRow[1] = "";

  for (const artNorm of displayArts) {
    const origArt = allArts.get(artNorm);
    let productName = goodsNames[artNorm] || '';
    const row = [origArt, productName];

    const baseOrders = ordersData.byArt[artNorm] || 0;
    const inT = inTransitByArt[artNorm] || { qty: 0, weight: 0 };
    const baseInTQty = inT.qty;
    const baseInTWeight = inT.weight;
    const totalStock = stockData.byArt[artNorm] || { stockQty: 0, stockWeight: 0 };

    // 🔥 5 БАЗОВЫХ СТОЛБЦОВ
    // row.push(baseOrders, baseInTQty, baseInTWeight, totalStock.stockQty, totalStock.stockWeight);
    // 🔥 5 БАЗОВЫХ СТОЛБЦОВ
    // 5-й столбец (индекс 6): Итого ВБ, шт = В пути, шт + На складе ВБ, шт
    const totalInWB = baseInTQty + totalStock.stockQty;
    row.push(baseOrders, baseInTQty, baseInTWeight, totalStock.stockQty, totalInWB);
    // 🔥 4 РАСЧЕТНЫХ СТОЛБЦА
    const avgSalesPerDay = dateDiff > 0 ? Math.round((baseOrders / dateDiff) * 100) / 100 : 0;
    row.push(avgSalesPerDay);
    const daysOfStock = avgSalesPerDay > 0 ? Math.round(((totalStock.stockQty + baseInTQty) / avgSalesPerDay) * 100) / 100 : 0;
    row.push(daysOfStock);
    let supplyVolume = Math.round((avgSalesPerDay * daysCount) - totalStock.stockQty - baseInTQty);
    row.push(supplyVolume <= 0 ? "" : supplyVolume);
    row.push(""); // ✏️ пусто

    // Регионы и склады
    for (const region of regions) {
      const warehouses = regionsMap.get(region);
      const whData = [];
      for (const wh of warehouses) {
        const key = `${artNorm}|${region}|${wh}`;
        const stock = stockData.byRegionWarehouse[key] || { stockQty: 0, stockWeight: 0 };
        const orders = ordersData.byRegionWarehouse[key] || 0;
        const supply = supplyByKey[key] || { inTransitQty: 0 };

        const avgSalesWh = dateDiff > 0 ? Math.round((orders / dateDiff) * 100) / 100 : 0;
        let volWh = Math.round((avgSalesWh * daysCount) - stock.stockQty - supply.inTransitQty);
        if (volWh <= 0) volWh = "";
        const daysWh = avgSalesWh > 0 ? (stock.stockQty + supply.inTransitQty) / avgSalesWh : 0;
        whData.push([stock.stockQty, stock.stockWeight, orders, supply.inTransitQty, avgSalesWh, daysWh, volWh, '']);
      }

      let rSQ = 0, rSW = 0, rOQ = 0, rITQ = 0;
      for (const wd of whData) { rSQ += wd[0]; rSW += wd[1]; rOQ += wd[2]; rITQ += wd[3]; }
      const regionAvgSales = dateDiff > 0 ? Math.round((rOQ / dateDiff) * 100) / 100 : 0;
      let regionSupplyVolume = Math.round((regionAvgSales * daysCount) - rSQ - rITQ);
      if (regionSupplyVolume <= 0) regionSupplyVolume = "";
      const regionDaysOfStock = regionAvgSales > 0 ? (rSQ + rITQ) / regionAvgSales : 0;

      row.push(rSQ, rSW, rOQ, rITQ, regionAvgSales, regionDaysOfStock, regionSupplyVolume, '');
      for (let d of whData) row.push(...d);
    }

    for (let i = 2; i < row.length; i++) {
      if (metricsRow[i - 1] === 'К-во на поставку' && row[i] === 0) row[i] = "";
    }
    dataRows.push(row);
    for (let i = 2; i < row.length; i++) {
      if (typeof row[i] === 'number' && metricsRow[i - 1] !== 'Запас в днях') totalRow[i] += row[i];
    }
  }

  if (dataRows.length > 0) {
    reportSheet.getRange(7, 1, 1, totalRow.length).setValues([totalRow]);
    reportSheet.getRange(7, 1, 1, totalRow.length).setFontWeight("bold").setBackground("#e8e8e8");
    reportSheet.getRange(8, 1, dataRows.length, dataRows[0].length).setValues(dataRows);

    for (let i = 0; i < metricsRow.length; i++) {
      if (metricsRow[i] === 'Запас в днях') reportSheet.getRange(7, i + 1).setValue("");
    }

    applyPostavkaMetricFormats_(reportSheet, metricsRow, 8, dataRows.length);
  }

  //  УСЛОВНОЕ ФОРМАТИРОВАНИЕ: низкие остатки и запасы (Упрощенные формулы)
  if (dataRows.length > 0) {
    const lastDataRow = 7 + dataRows.length;
    const LOW_STOCK_COLOR = "#ff9999";
    const rules = [];

    // === БАЗОВЫЙ БЛОК ===
    // "На складе ВБ, шт" = колонка 6 (F)
    // "Запас в днях" = колонка 9 (I)

    // Правило: Остатки <= 10 (включая 0, так как это критично)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=F8<=10`) //  Простая формула, как вы просили
      .setBackground(LOW_STOCK_COLOR)
      .setRanges([reportSheet.getRange(`F8:F${lastDataRow}`)])
      .build());

    // Правило: Запас в днях <= 30
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=I8<=30`)
      .setBackground(LOW_STOCK_COLOR)
      .setRanges([reportSheet.getRange(`I8:I${lastDataRow}`)])
      .build());

    // === РЕГИОНЫ И СКЛАДЫ ===
    let cursor = REGION_START_COL; // = 12
    for (const region of regions) {
      const warehouses = regionsMap.get(region);

      // --- ИТОГО по региону ---
      const rStockLet = getColumnLetter(cursor + 0);    // "Остатки, шт"
      const rDaysLet = getColumnLetter(cursor + 5);     // "Запас в днях"

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${rStockLet}8<=10`) // Простая формула
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(`${rStockLet}8:${rStockLet}${lastDataRow}`)])
        .build());

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${rDaysLet}8<=30`) // Простая формула
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(`${rDaysLet}8:${rDaysLet}${lastDataRow}`)])
        .build());

      cursor += WAREHOUSE_METRICS_COUNT;

      // --- Склады ---
      for (const wh of warehouses) {
        const wStockLet = getColumnLetter(cursor + 0);
        const wDaysLet = getColumnLetter(cursor + 5);

        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=${wStockLet}8<=10`)
          .setBackground(LOW_STOCK_COLOR)
          .setRanges([reportSheet.getRange(`${wStockLet}8:${wStockLet}${lastDataRow}`)])
          .build());

        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=${wDaysLet}8<=30`)
          .setBackground(LOW_STOCK_COLOR)
          .setRanges([reportSheet.getRange(`${wDaysLet}8:${wDaysLet}${lastDataRow}`)])
          .build());

        cursor += WAREHOUSE_METRICS_COUNT;
      }
    }

    // Применяем все правила за 1 обращение к серверу
    if (rules.length > 0) {
      reportSheet.setConditionalFormatRules(rules);
    }
  }

  // 🔥 ФОРМУЛЫ ✏️ В БАЗОВОМ БЛОКЕ: пачкой
  if (dataRows.length > 0) {
    const pencilColBase = 11;
    const lastDataRow = 7 + dataRows.length;
    const numDataRows = lastDataRow - 7;
    const pencilFormulas = [];

    // Собираем колонки ✏️ регионов один раз
    const regionPencilCols = [];
    let cursorP = REGION_START_COL;
    for (const region of regions) {
      const warehouses = regionsMap.get(region);
      regionPencilCols.push(getColumnLetter(cursorP + 7));
      cursorP += WAREHOUSE_METRICS_COUNT * (1 + warehouses.length);
    }

    if (regionPencilCols.length > 0) {
      for (let row = 8; row <= lastDataRow; row++) {
        pencilFormulas.push([`=${regionPencilCols.map(c => c + row).join("+")}`]);
      }
      reportSheet.getRange(8, pencilColBase, numDataRows, 1).setFormulas(pencilFormulas);
    }
  }

  SpreadsheetApp.flush();
  if (reportSheet.getFilter()) reportSheet.getFilter().remove();
  const filterLastRow = dataRows.length > 0 ? (7 + dataRows.length) : 7;
  try { reportSheet.getRange(6, 1, filterLastRow - 5, metricsRow.length).createFilter(); } catch (e) { }

  // 🔥 ФОРМУЛА ✏️ В СТРОКЕ ИТОГО (сумма ✏️ по всем регионам)
  // 🔥 ФОРМУЛЫ В СТРОКЕ ИТОГО (7) — ДОБАВЛЯЕМ ТОЛЬКО ТО, ЧТО ПРОСИЛИ
  // (Запас, К-во, ✏️ для всех блоков: базовый, регион, склад)

  if (dataRows.length > 0) {
    const lastDataRow = 7 + dataRows.length;

    // === БАЗОВЫЙ БЛОК (колонки 9, 10, 11 = Запас, К-во, ✏️) ===
    // Запас в днях: формула =IF(H7=0;"";(F7+D7)/H7)
    reportSheet.getRange(7, 9).setFormula('=IF(H7=0;"";(F7+D7)/H7)');
    reportSheet.getRange(7, 9).setNumberFormat("#,##0.00");

    // К-во на поставку: сумма по столбцу
    reportSheet.getRange(7, 10).setFormula(`=SUM(J8:J${lastDataRow})`);
    reportSheet.getRange(7, 10).setNumberFormat("0"); // 1 знак в расчёте, целое на экране

    // ✏️: сумма по столбцу
    reportSheet.getRange(7, 11).setFormula(`=SUM(K8:K${lastDataRow})`);
    reportSheet.getRange(7, 11).setNumberFormat("0");

    // === РЕГИОНЫ И СКЛАДЫ ===
    let cursor = REGION_START_COL; // = 12
    for (const region of regions) {
      const warehouses = regionsMap.get(region);

      // --- ИТОГО по региону ---
      // Позиции в блоке из 8 колонок: 0=Остатки, 3=🚚, 4=Ср.продажи, 5=Запас, 6=К-во, 7=✏️
      const rW = cursor + 4; // Ср.продажи
      const rD = cursor + 5; // Запас в днях
      const rK = cursor + 6; // К-во на поставку
      const rX = cursor + 7; // ✏️

      const rLetW = getColumnLetter(rW);

      // Запас в днях: формула
      reportSheet.getRange(7, rD).setFormula(`=IF(${rLetW}7=0;"";(${getColumnLetter(cursor)}7+${getColumnLetter(cursor + 3)}7)/${rLetW}7)`);
      reportSheet.getRange(7, rD).setNumberFormat("#,##0.00");

      // К-во на поставку: сумма по столбцу
      reportSheet.getRange(7, rK).setFormula(`=SUM(${getColumnLetter(rK)}8:${getColumnLetter(rK)}${lastDataRow})`);
      reportSheet.getRange(7, rK).setNumberFormat("0");

      // ✏️: сумма по столбцу
      reportSheet.getRange(7, rX).setFormula(`=SUM(${getColumnLetter(rX)}8:${getColumnLetter(rX)}${lastDataRow})`);
      reportSheet.getRange(7, rX).setNumberFormat("0");

      cursor += WAREHOUSE_METRICS_COUNT;

      // --- Склады ---
      for (const wh of warehouses) {
        const wW = cursor + 4;
        const wD = cursor + 5;
        const wK = cursor + 6;
        const wX = cursor + 7;

        const wLetW = getColumnLetter(wW);

        // Запас в днях: формула
        reportSheet.getRange(7, wD).setFormula(`=IF(${wLetW}7=0;"";(${getColumnLetter(cursor)}7+${getColumnLetter(cursor + 3)}7)/${wLetW}7)`);
        reportSheet.getRange(7, wD).setNumberFormat("#,##0.00");

        // К-во на поставку: сумма по столбцу
        reportSheet.getRange(7, wK).setFormula(`=SUM(${getColumnLetter(wK)}8:${getColumnLetter(wK)}${lastDataRow})`);
        reportSheet.getRange(7, wK).setNumberFormat("0");

        // ✏️: сумма по столбцу
        reportSheet.getRange(7, wX).setFormula(`=SUM(${getColumnLetter(wX)}8:${getColumnLetter(wX)}${lastDataRow})`);
        reportSheet.getRange(7, wX).setNumberFormat("0");

        cursor += WAREHOUSE_METRICS_COUNT;
      }
    }
  }

  calculateRequestForShipment();
  addRegionDropdown();
  formatPostavkaSheetFinal();
  // createWarehouseGroups();
  applyRegionColors();

  // 🔥 ФИНАЛЬНОЕ ПРИМЕНЕНИЕ ЦВЕТА #cee2e3 К ОБЪЕДИНЁННЫМ ЯЧЕЙКАМ (после ВСЕХ форматирований)
  const BASE_BLOCK_START = 3;  // Колонка "Заказы,\nшт"
  const BASE_BLOCK_END = 11;   // Колонка "✏️"
  const BASE_BG_COLOR = "#cee2e3";
  const BASE_WIDTH = BASE_BLOCK_END - BASE_BLOCK_START + 1; // = 9

  // Строка 4: "Все склады WB" (объединено)
  reportSheet.getRange(4, BASE_BLOCK_START, 1, BASE_WIDTH)
    .merge()
    .setValue("Все склады WB")
    .setFontFamily("Calibri")
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(BASE_BG_COLOR);

  // Строка 5: "Итого по всем регионам" (объединено)
  reportSheet.getRange(5, BASE_BLOCK_START, 1, BASE_WIDTH)
    .merge()
    .setValue("Итого по всем регионам")
    .setFontFamily("Calibri")
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(BASE_BG_COLOR);

  // Заголовки базового блока (строка 6, колонки 3-11)
  reportSheet.getRange(6, BASE_BLOCK_START, 1, BASE_WIDTH).setBackground(BASE_BG_COLOR);

  // Значения базового блока (строки 8+, колонки 3-11) — ИТОГО (строка 7) НЕ трогаем
  const lastRow = reportSheet.getLastRow();
  if (lastRow >= 8) {
    reportSheet.getRange(8, BASE_BLOCK_START, lastRow - 7, BASE_WIDTH).setBackground(BASE_BG_COLOR);
  }

  trimPostavkaExtraColumns_(reportSheet, metricsRow.length);

  SpreadsheetApp.getUi().alert('✅ Отчёт "Поставка" обновлён!');
}

// ============================================
// ФОРМАТИРОВАНИЕ И РАСЧЁТЫ
// ============================================
function formatPostavkaSheetFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Поставка");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 6) return;

  const headers = sheet.getRange(6, 1, 1, lastCol).getValues()[0];

  // Серый фон для Итого (цвета регионов применит applyRegionColors)
  sheet.getRange(7, 1, 1, lastCol).setFontWeight("bold").setBackground("#e8e8e8");

  if (lastCol >= REGION_START_COL) {
    sheet.getRange(4, REGION_START_COL, 2, lastCol - REGION_START_COL + 1).setFontFamily("Calibri").setFontSize(9);
  }

  sheet.getRange(6, 1, 1, lastCol).setFontFamily("Calibri").setFontSize(9).setFontWeight("bold").setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");

  if (lastRow >= 8) {
    sheet.getRange(8, 1, lastRow - 7, lastCol).setFontSize(9).setFontFamily("Calibri");
  }

  // Иконки и жирный ✏️ — только нужные колонки
  const pencilCols = [];
  for (let col = 1; col <= lastCol; col++) {
    const h = headers[col - 1]?.toString().trim();
    if (h === '🚚') sheet.getRange(6, col).setFontSize(17);
    if (h === '✏️') {
      sheet.getRange(6, col).setFontSize(13);
      pencilCols.push(col);
    }
  }
  if (lastRow >= 7 && pencilCols.length > 0) {
    for (let i = 0; i < pencilCols.length; i++) {
      sheet.getRange(7, pencilCols[i], lastRow - 6, 1).setFontWeight('bold');
    }
  }
}

function calculateRequestForShipment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Поставка");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 8) return;
  const headerRow = sheet.getRange(5, 1, 1, lastCol).getValues()[0];
  const REQUEST_COL_OFFSET = 7;
  const numRows = lastRow - 7;

  // Структура блоков не зависит от строки — собираем один раз, пишем пачкой
  for (let col = REGION_START_COL; col <= lastCol; col++) {
    if (headerRow[col - 1] !== "ИТОГО по региону") continue;
    const requestCol = col + REQUEST_COL_OFFSET;
    const warehouseRequestCols = [];
    let tempCol = col + BLOCK_SIZE;
    while (tempCol <= lastCol) {
      const nextHeader = headerRow[tempCol - 1];
      if (nextHeader === "ИТОГО по региону") break;
      if (nextHeader && nextHeader !== "") {
        warehouseRequestCols.push(tempCol + REQUEST_COL_OFFSET);
        tempCol += BLOCK_SIZE;
      } else { tempCol++; }
    }
    if (warehouseRequestCols.length < 1) continue;

    const formulas = [];
    for (let row = 8; row <= lastRow; row++) {
      if (warehouseRequestCols.length === 1) {
        const c = getColumnLetter(warehouseRequestCols[0]);
        formulas.push([`=IF(${c}${row}=""; 0; ${c}${row})`]);
      } else {
        formulas.push([`=${warehouseRequestCols.map(wc => getColumnLetter(wc) + row).join("+")}`]);
      }
    }
    sheet.getRange(8, requestCol, numRows, 1).setFormulas(formulas);
  }
}

function getColumnLetter(col) {
  let letter = '';
  while (col > 0) {
    const modulo = (col - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/** Пачками ставит numberFormat на список колонок (склеивает соседние) */
function setNumberFormatsGrouped_(sheet, startRow, numRows, cols, format) {
  if (!cols || cols.length === 0 || numRows < 1) return;
  const sorted = cols.slice().sort((a, b) => a - b);
  let runStart = sorted[0];
  let runEnd = sorted[0];
  for (let i = 1; i <= sorted.length; i++) {
    const c = sorted[i];
    if (c === runEnd + 1) {
      runEnd = c;
      continue;
    }
    sheet.getRange(startRow, runStart, numRows, runEnd - runStart + 1).setNumberFormat(format);
    runStart = c;
    runEnd = c;
  }
}

/** Форматы чисел отчёта Поставка по заголовкам строки 6 */
function applyPostavkaMetricFormats_(sheet, metricsRow, dataStartRow, numDataRows) {
  const kg = [], avg = [], intFmt = [], qtyBold = [];
  for (let col = 3; col <= metricsRow.length; col++) {
    const h = (metricsRow[col - 1] || '').toString();
    if (!h || h.includes('Артикул') || h.includes('Наименование') || h.includes('Заявка')) continue;
    if (h.includes('кг')) kg.push(col);
    else if (h.includes('Ср. продажи')) avg.push(col);
    else if (h.includes('К-во на поставку')) qtyBold.push(col);
    else intFmt.push(col);
  }
  if (numDataRows > 0) {
    setNumberFormatsGrouped_(sheet, dataStartRow, numDataRows, kg, "#,##0.000");
    setNumberFormatsGrouped_(sheet, dataStartRow, numDataRows, avg, "#,##0.00");
    setNumberFormatsGrouped_(sheet, dataStartRow, numDataRows, intFmt, "#,##0");
    setNumberFormatsGrouped_(sheet, dataStartRow, numDataRows, qtyBold, "#,##0");
    for (let i = 0; i < qtyBold.length; i++) {
      sheet.getRange(dataStartRow, qtyBold[i], numDataRows, 1).setFontWeight('bold');
    }
  }
  // строка ИТОГО (7)
  setNumberFormatsGrouped_(sheet, 7, 1, kg, "#,##0.000");
  setNumberFormatsGrouped_(sheet, 7, 1, avg, "#,##0.00");
  setNumberFormatsGrouped_(sheet, 7, 1, intFmt.concat(qtyBold), "#,##0");
}

/** Удаляет пустые столбцы справа — иначе Google считает их заполненными */
function trimPostavkaExtraColumns_(sheet, usedCols) {
  if (!sheet || usedCols < 1) return;
  var maxCols = sheet.getMaxColumns();
  if (maxCols <= usedCols) return;
  try { removeColumnGroupsInRange_(sheet, usedCols + 1, maxCols); } catch (e1) {}
  try {
    var rows = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow(), 10));
    sheet.getRange(1, usedCols + 1, rows, maxCols - usedCols).breakApart();
  } catch (e2) {}
  try {
    var filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e3) {}
  try {
    sheet.deleteColumns(usedCols + 1, maxCols - usedCols);
  } catch (e4) {
    Logger.log('trimPostavkaExtraColumns_: ' + e4);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 6) {
    try {
      sheet.getRange(6, 1, Math.max(1, lastRow - 5), usedCols).createFilter();
    } catch (e5) {}
  }
}

// ============================================
// ГРУППИРОВКА И ЦВЕТА
// ============================================

/** Быстрое удаление групп столбцов: снимаем самый глубокий уровень, не прыгаем зря */
function removeColumnGroupsInRange_(sheet, fromCol, toCol) {
  if (!sheet || toCol < fromCol) return;
  let col = fromCol;
  let guard = 0;
  const maxGuard = (toCol - fromCol + 1) * 4 + 100;
  while (col <= toCol && guard++ < maxGuard) {
    let depth = 0;
    try { depth = sheet.getColumnGroupDepth(col); } catch (e) { col++; continue; }
    if (depth < 1) { col++; continue; }
    try {
      const group = sheet.getColumnGroup(col, depth);
      if (group) group.remove();
      else col++;
    } catch (e) { col++; }
    // после remove остаёмся на той же col — снять следующий уровень вложенности
  }
}

function createWarehouseGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Поставка");
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < REGION_START_COL) return;

  const regionRow = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const subHeaderRow = sheet.getRange(5, 1, 1, lastCol).getValues()[0];

  removeColumnGroupsInRange_(sheet, REGION_START_COL, lastCol);

  let currentPos = REGION_START_COL;
  const regionGroups = [];
  const GROUPED_COLS = 7;
  while (currentPos <= lastCol) {
    const regionName = regionRow[currentPos - 1];
    if (regionName && regionName !== "") {
      currentPos += BLOCK_SIZE;
      const warehouses = [];
      while (currentPos <= lastCol) {
        const header = subHeaderRow[currentPos - 1];
        const nextRegionName = regionRow[currentPos - 1];
        if (nextRegionName && nextRegionName !== "" && nextRegionName !== regionName) break;
        if (header && header !== "" && header !== "ИТОГО по региону") {
          warehouses.push({ name: header, startCol: currentPos });
          currentPos += BLOCK_SIZE;
        } else { currentPos++; }
      }
      regionGroups.push({ name: regionName, warehouses: warehouses, regionEnd: currentPos - 1 });
    } else { currentPos++; }
  }

  for (let i = 0; i < regionGroups.length; i++) {
    const region = regionGroups[i];
    if (region.warehouses.length > 0) {
      const firstWarehouseCol = region.warehouses[0].startCol;
      const lastColRegion = region.regionEnd;
      if (firstWarehouseCol <= lastColRegion) {
        try { sheet.getRange(1, firstWarehouseCol, 1, lastColRegion - firstWarehouseCol + 1).shiftColumnGroupDepth(1); } catch (e) { }
      }
    }
  }
  for (let i = 0; i < regionGroups.length; i++) {
    const region = regionGroups[i];
    for (let j = 0; j < region.warehouses.length; j++) {
      const warehouse = region.warehouses[j];
      if (warehouse.startCol + GROUPED_COLS - 1 <= lastCol) {
        try { sheet.getRange(1, warehouse.startCol, 1, GROUPED_COLS).shiftColumnGroupDepth(1); } catch (e) { }
      }
    }
  }

  // Один вызов вместо обхода всех колонок × 2
  try { sheet.collapseAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= lastCol; col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && !g.isCollapsed()) g.collapse(); } catch (e2) { }
    }
  }
}

function expandAllWarehouseGroups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Поставка");
  if (!sheet) return;
  try { sheet.expandAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= sheet.getLastColumn(); col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && g.isCollapsed()) g.expand(); } catch (e2) { }
    }
  }
}
function collapseAllWarehouseGroups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Поставка");
  if (!sheet) return;
  try { sheet.collapseAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= sheet.getLastColumn(); col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && !g.isCollapsed()) g.collapse(); } catch (e2) { }
    }
  }
}
function removeAllGroups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Поставка");
  if (!sheet) return;
  removeColumnGroupsInRange_(sheet, REGION_START_COL, sheet.getLastColumn());
}

function создатьструктуру() {
  // createWarehouseGroups сам сбрасывает старые группы — removeAllGroups не дублируем
  createWarehouseGroups();
  try { SpreadsheetApp.getUi().alert('✅ Группировка обновлена'); } catch (e) { }
}

function applyRegionColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Поставка");
  if (!sheet) return;
  const regionColors = { "Склад WB": "#76d7c4", "Центральный": "#b7b7b7", "Северо-Западный": "#b4a7d6", "Приволжский": "#ffff00", "Южный и Северо-Кавказский": "#f2cfcd", "Уральский": "#6aa84f", "Дальневосточный и Сибирский": "#3881c9", "Беларусь": "#b6d7a8", "Казахстан": "#f9cb9c", "Узбекистан": "#8e7cc3", "Армения": "#ea9999", "Грузия": "#9fc5e8", "Ереван": "#ffe599", "Не определён": "#cccccc" };
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const regionRow = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  if (lastRow < 8) return;
  let currentRegion = null, currentRegionColor = null, regionStartCol = REGION_START_COL;

  for (let col = REGION_START_COL; col <= lastCol; col++) {
    const regionName = regionRow[col - 1];
    if (regionName && regionName !== "" && regionColors[regionName]) { currentRegion = regionName; currentRegionColor = regionColors[currentRegion]; regionStartCol = col; }
    const nextRegionName = (col + 1 <= lastCol) ? regionRow[col] : null;
    const isEndOfRegion = (nextRegionName && nextRegionName !== "" && regionColors[nextRegionName]) || col === lastCol;
    if (currentRegion && isEndOfRegion) {
      sheet.getRange(4, regionStartCol, lastRow - 3, col - regionStartCol + 1).setBackground(currentRegionColor);
      for (let blockStart = regionStartCol; blockStart <= col; blockStart += BLOCK_SIZE) {
        const blockEnd = Math.min(blockStart + BLOCK_SIZE - 1, col);
        sheet.getRange(4, blockStart, lastRow - 3, blockEnd - blockStart + 1).setBorder(true, true, true, true, false, false);
      }
      currentRegion = null;
    }
  }
  if (lastRow >= 6) { sheet.getRange(6, 1, 1, 7).setBackground("#f1f1f1"); sheet.getRange(6, 1, 1, lastCol).setFontWeight("bold"); }
  if (lastRow >= 7) sheet.getRange(7, 1, 1, lastCol).setBackground("#e8e8e8").setFontWeight("bold").setFontSize(9).setFontFamily("Calibri");
}

// ============================================
// ВЫПАДАЮЩИЕ СПИСКИ И ONEDIT
// ============================================
function addRegionDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Поставка");
  if (!sheet) return;
  const pairs = getUniqueRegionWarehousePairs();
  const regions = [...new Set(pairs.map(p => p.region))].sort();
  const regionRule = SpreadsheetApp.newDataValidation().requireValueInList(["Выберите регион", ...regions], true).setAllowInvalid(false).build();
  const b4 = sheet.getRange("B4"); b4.setDataValidation(regionRule); b4.setValue("Выберите регион"); formatDropdownCell(b4);
  const warehouseRule = SpreadsheetApp.newDataValidation().requireValueInList(["Выберите склад"], true).setAllowInvalid(false).build();
  const b5 = sheet.getRange("B5"); b5.setDataValidation(warehouseRule); b5.setValue("Выберите склад"); formatDropdownCell(b5);
  if (sheet.getColumnWidth(2) < 150) sheet.setColumnWidth(2, 180);
}
function formatDropdownCell(range) { range.setFontFamily("Calibri").setFontSize(8).setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, false, false); }

function getRegionWarehouseMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Поставка");
  if (!sheet || sheet.getLastRow() < 5) return {};
  const lastCol = sheet.getLastColumn();
  const rRow = sheet.getRange(4, REGION_START_COL, 1, Math.max(1, lastCol - REGION_START_COL + 1)).getValues()[0];
  const wRow = sheet.getRange(5, REGION_START_COL, 1, Math.max(1, lastCol - REGION_START_COL + 1)).getValues()[0];
  const map = {}; let currentRegion = "";
  for (let i = 0; i < rRow.length; i++) {
    const regCell = rRow[i]?.toString().trim();
    const whCell = wRow[i]?.toString().trim();
    if (regCell) { currentRegion = regCell; if (!map[currentRegion]) map[currentRegion] = []; }
    else if (currentRegion && whCell) map[currentRegion].push(whCell);
  }
  return map;
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  // Поставка_2: делегируем в отдельный обработчик
  if (sheetName === "Поставка_2") {
    handlePostavka2Edit(e);
    return;
  }
  if (sheetName !== "Поставка") return;
  if (e.range.getA1Notation() !== "B4") return;
  const b5 = sheet.getRange("B5");
  const region = e.value;
  b5.clearContent(); b5.clearDataValidations(); b5.setValue("⏳"); SpreadsheetApp.flush();
  if (!region || region === "Выберите регион") { setDropdownList(b5, ["Выберите склад"]); return; }
  const lastCol = sheet.getLastColumn();
  const rRow = sheet.getRange(4, REGION_START_COL, 1, Math.max(1, lastCol - REGION_START_COL + 1)).getValues()[0];
  const wRow = sheet.getRange(5, REGION_START_COL, 1, Math.max(1, lastCol - REGION_START_COL + 1)).getValues()[0];
  const warehouses = []; let currentReg = "";
  for (let i = 0; i < rRow.length; i++) {
    const regCell = rRow[i]?.toString().trim();
    const whCell = wRow[i]?.toString().trim();
    if (regCell) currentReg = regCell;
    else if (currentReg === region && whCell) warehouses.push(whCell);
  }
  setDropdownList(b5, warehouses.length > 0 ? warehouses : ["Нет складов для региона"]);
}
function setDropdownList(range, list) {
  range.clearContent();
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  range.setDataValidation(rule); range.setValue("Выберите склад"); SpreadsheetApp.flush();
}