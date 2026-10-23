function firstFinanceCol_(headers, names) {
  if (typeof wbColIdx_ === 'function') {
    var idx = wbColIdx_(headers, names);
    if (idx >= 0) return idx;
  }
  const lower = names.map(n => n.toString().trim().toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i] != null ? String(headers[i]).trim().toLowerCase() : '';
    if (lower.indexOf(h) !== -1) return i;
  }
  return -1;
}

function parseFinanceDate_(val) {
  if (val == null || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const str = String(val).trim();
  const m = str.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return new Date(str.substring(0, 10) + 'T00:00:00');
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function сформироватьОтчет_Вес() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Финансовые_ежедневные");
  const startSheet = ss.getSheetByName("Вес продаж");

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const data = source.getDataRange().getValues();
  const headers = data[0];
  const dateIdx = firstFinanceCol_(headers, ['Дата', 'rr_dt']);
  const artIdx = firstFinanceCol_(headers, ['sa_name']);
  const nameIdx = firstFinanceCol_(headers, ['Имя товара']);
  const valueCol = firstFinanceCol_(headers, ['Вес итог, кг с учетом возвратов']);
  if (dateIdx < 0 || artIdx < 0 || nameIdx < 0 || valueCol < 0) {
    SpreadsheetApp.getUi().alert('❌ На листе «Финансовые_ежедневные» нет нужных столбцов (Дата / sa_name / Имя товара / Вес итог)');
    return;
  }

  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "Вес продаж";

  const result = {};
  const datesObj = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = parseFinanceDate_(row[dateIdx]);
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;

    const art = row[artIdx];
    const name = row[nameIdx];
    const val = row[valueCol] || 0;
    const d = Utilities.formatDate(dt, tz, "dd.MM");

    // 🔧 ФИЛЬТРУЕМ ПУСТЫЕ АРТИКУЛЫ И НАИМЕНОВАНИЯ
    if (!art || art === '' || art === ' ' || art === '-' || !name || name === '' || name === ' ') {
      continue; // пропускаем пустые строки
    }

    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    result[key][d] = (result[key][d] || 0) + val;
  }

  // const dateList = Object.keys(datesObj).sort();
  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];
  
  // Заголовки
  output.push(["Артикул", "Наименование", ...dateList, "Итого"]);

  // Рассчитываем итоги по столбцам
  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0;
  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));

  // Сначала собираем все строки данных, чтобы посчитать итоги
  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split("|");
    const rowObj = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const v = rowObj[d] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    row.push(sumRow);
    dataRows.push(row);
  }

  // Добавляем строку Итого (после заголовка, перед данными)
  output.push(["Итого", "", ...colTotals, grand]);
  
  // Добавляем все строки данных
  for (const row of dataRows) {
    output.push(row);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (sheet) clearReportBody_(sheet);
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  // const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  // headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setTextRotation(90);
  // if (sheet.getFilter()) sheet.getFilter().remove();
  // headerRange.createFilter();

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center");

  // Только для колонок с датами (начиная с 3-й) делаем поворот
  if (output[0].length > 2) {
    const dateHeaderRange = sheet.getRange(3, 3, 1, output[0].length - 2);
    dateHeaderRange.setTextRotation(90);
  }

  // Строка "Итого" (вторая строка вывода) - жирный шрифт и фон
  sheet.getRange(4, 1, 1, output[0].length).setFontWeight("bold").setBackground("#e8e8e8");

  for (let col = 3; col <= output[0].length; col++) sheet.setColumnWidth(col, 40);
  sheet.setRowHeight(3, 45);

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Вес продаж' обновлён");
}

function сформироватьОтчет_Количество() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const source = ss.getSheetByName("Данные_для_анализа");
  const source = ss.getSheetByName("Финансовые_ежедневные");
  const startSheet = ss.getSheetByName("К-во продаж");

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const data = source.getDataRange().getValues();
  const headers = data[0];
  const dateIdx = firstFinanceCol_(headers, ['Дата', 'rr_dt']);
  const artIdx = firstFinanceCol_(headers, ['sa_name']);
  const nameIdx = firstFinanceCol_(headers, ['Имя товара']);
  const valueCol = firstFinanceCol_(headers, ['Сумма упаковок с учетом возвратов']);
  if (dateIdx < 0 || artIdx < 0 || nameIdx < 0 || valueCol < 0) {
    SpreadsheetApp.getUi().alert('❌ На листе «Финансовые_ежедневные» нет нужных столбцов (Дата / sa_name / Имя товара / Сумма упаковок)');
    return;
  }

  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "К-во продаж";

  const result = {};
  const datesObj = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = parseFinanceDate_(row[dateIdx]);
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;

    const art = row[artIdx];
    const name = row[nameIdx];
    const val = row[valueCol] || 0;
    const d = Utilities.formatDate(dt, tz, "dd.MM");

    // 🔧 ФИЛЬТРУЕМ ПУСТЫЕ АРТИКУЛЫ И НАИМЕНОВАНИЯ
    if (!art || art === '' || art === ' ' || art === '-' || !name || name === '' || name === ' ') {
      continue; // пропускаем пустые строки
    }

    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    result[key][d] = (result[key][d] || 0) + val;
  }

  // const dateList = Object.keys(datesObj).sort();
  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];
  
  // Заголовки
  output.push(["Артикул", "Наименование", ...dateList, "Итого"]);

  // Рассчитываем итоги по столбцам
  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0;
  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));

  // Сначала собираем все строки данных, чтобы посчитать итоги
  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split("|");
    const rowObj = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const v = rowObj[d] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    row.push(sumRow);
    dataRows.push(row);
  }

  // Добавляем строку Итого (после заголовка, перед данными)
  output.push(["Итого", "", ...colTotals, grand]);
  
  // Добавляем все строки данных
  for (const row of dataRows) {
    output.push(row);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (sheet) clearReportBody_(sheet);
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  // const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  // headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setTextRotation(90);
  // if (sheet.getFilter()) sheet.getFilter().remove();
  // headerRange.createFilter();

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center");

  // Только для колонок с датами (начиная с 3-й) делаем поворот
  if (output[0].length > 2) {
    const dateHeaderRange = sheet.getRange(3, 3, 1, output[0].length - 2);
    dateHeaderRange.setTextRotation(90);
  }

  // Строка "Итого" (вторая строка вывода) - жирный шрифт и фон
  sheet.getRange(4, 1, 1, output[0].length).setFontWeight("bold").setBackground("#e8e8e8");

  for (let col = 3; col <= output[0].length; col++) sheet.setColumnWidth(col, 30);
  sheet.setRowHeight(3, 45);

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Количество продаж' обновлён");
}

/* отключено
function сформироватьОтчет_ПродажиВозврат() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const source = ss.getSheetByName("Данные_для_анализа");
  const source = ss.getSheetByName("Финансовые_ежедневные");
  const startSheet = ss.getSheetByName("Продажи");

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const data = source.getDataRange().getValues();
  const headers = data[0];
  const dateIdx = firstFinanceCol_(headers, ['Дата', 'rr_dt']);
  const artIdx = firstFinanceCol_(headers, ['sa_name']);
  const nameIdx = firstFinanceCol_(headers, ['Имя товара']);
  const valueCol = firstFinanceCol_(headers, ['Сумма продаж с учетом возвратов']);
  if (dateIdx < 0 || artIdx < 0 || nameIdx < 0 || valueCol < 0) {
    SpreadsheetApp.getUi().alert('❌ На листе «Финансовые_ежедневные» нет нужных столбцов (Дата / sa_name / Имя товара / Сумма продаж)');
    return;
  }

  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "Продажи";

  const result = {};
  const datesObj = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = parseFinanceDate_(row[dateIdx]);
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;

    const art = row[artIdx];
    const name = row[nameIdx];
    const val = row[valueCol] || 0;
    const d = Utilities.formatDate(dt, tz, "dd.MM");

    // 🔧 ФИЛЬТРУЕМ ПУСТЫЕ АРТИКУЛЫ И НАИМЕНОВАНИЯ
    if (!art || art === '' || art === ' ' || art === '-' || !name || name === '' || name === ' ') {
      continue; // пропускаем пустые строки
    }

    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    result[key][d] = (result[key][d] || 0) + val;
  }

  // const dateList = Object.keys(datesObj).sort();

  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];
  
  // Заголовки
  output.push(["Артикул", "Наименование", ...dateList, "Итого"]);

  // Рассчитываем итоги по столбцам
  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0;
  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));

  // Сначала собираем все строки данных, чтобы посчитать итоги
  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split("|");
    const rowObj = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const v = rowObj[d] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    row.push(sumRow);
    dataRows.push(row);
  }

  // Добавляем строку Итого (после заголовка, перед данными)
  output.push(["Итого", "", ...colTotals, grand]);
  
  // Добавляем все строки данных
  for (const row of dataRows) {
    output.push(row);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (sheet) clearReportBody_(sheet);
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  // const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  // headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setTextRotation(90);
  // if (sheet.getFilter()) sheet.getFilter().remove();
  // headerRange.createFilter();

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center");

  // Только для колонок с датами (начиная с 3-й) делаем поворот
  if (output[0].length > 2) {
    const dateHeaderRange = sheet.getRange(3, 3, 1, output[0].length - 2);
    dateHeaderRange.setTextRotation(90);
  }
  
  // Строка "Итого" (вторая строка вывода) - жирный шрифт и фон
  sheet.getRange(4, 1, 1, output[0].length).setFontWeight("bold").setBackground("#e8e8e8");

  for (let col = 3; col <= output[0].length; col++) sheet.setColumnWidth(col, 50);
  sheet.setRowHeight(3, 45);

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Продажи' обновлён");
}
*/

function funnelOrdersReportSheet_(ss) {
  var sh = ss.getSheetByName('Воронка продаж');
  if (sh) return sh;
  if (typeof SF_SHEET_ === 'string') {
    sh = ss.getSheetByName(SF_SHEET_);
    if (sh) return sh;
  }
  return ss.getSheetByName('Аналитика_проба_дни');
}

function funnelOrdersReportFill_(ss, dateFrom, dateTo, valueHeader, result, datesObj) {
  var source = funnelOrdersReportSheet_(ss);
  if (!source || source.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ Лист «Воронка продаж» пуст или не найден. Сначала выгрузите воронку.');
    return false;
  }
  var data = source.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = firstFinanceCol_(headers, ['Период с']);
  var artIdx = firstFinanceCol_(headers, ['Артикул продавца']);
  var nameIdx = firstFinanceCol_(headers, ['Название']);
  var valueIdx = firstFinanceCol_(headers, [valueHeader]);
  if (dateIdx < 0 || artIdx < 0 || nameIdx < 0 || valueIdx < 0) {
    SpreadsheetApp.getUi().alert('❌ На листе «Воронка продаж» нет столбцов: Период с / Артикул продавца / Название / ' + valueHeader);
    return false;
  }
  var tz = Session.getScriptTimeZone();
  var fromD = new Date(dateFrom);
  var toD = new Date(dateTo);
  fromD.setHours(0, 0, 0, 0);
  toD.setHours(23, 59, 59, 999);
  var fromTime = fromD.getTime();
  var toTime = toD.getTime();
  var i, row, dt, art, name, val, d, key;
  for (i = 1; i < data.length; i++) {
    row = data[i];
    dt = parseFinanceDate_(row[dateIdx]);
    if (!dt) continue;
    if (dt.getTime() < fromTime || dt.getTime() > toTime) continue;
    art = row[artIdx];
    name = row[nameIdx];
    if (!art || art === '' || art === ' ' || art === '-' || !name || name === '' || name === ' ') continue;
    val = parseFloat(row[valueIdx]) || 0;
    d = Utilities.formatDate(dt, tz, 'dd.MM');
    datesObj[d] = true;
    key = art + '|' + name;
    if (!result[key]) {
      result[key] = { data: {}, returns: 0, stocks: 0, avgPrice: 0, currentPrice: 0, manualPrice: 0 };
    }
    result[key].data[d] = (result[key].data[d] || 0) + val;
  }
  return true;
}

function сформироватьОтчет_Заказы_Количество() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startSheet = ss.getSheetByName("К-во заказов");
  const financeSheet = ss.getSheetByName("Финансовые_ежедневные");
  const stocksSheet = ss.getSheetByName("Остатки");
  const goodsSheet = ss.getSheetByName("Товары");
  const manualPriceSheet = ss.getSheetByName("Товар+Вес"); // 🔥 ДОБАВЛЕНО

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "К-во заказов";
  
  const result = {};
  const datesObj = {};

  // ============ 1. ОСНОВНОЙ БЛОК (из Воронка продаж) ============
  if (!funnelOrdersReportFill_(ss, dateFrom, dateTo, 'Заказы шт', result, datesObj)) return;

  // ============ 2. ЗАГРУЗКА ИЗ ФИНАНСОВЫХ_ЕЖЕДНЕВНЫЕ ============
  if (financeSheet) {
    const financeValues = financeSheet.getDataRange().getValues();
    const financeHeaders = financeValues[0];
    
    const artIdx = firstFinanceCol_(financeHeaders, ['sa_name']);
    const qtyIdx = firstFinanceCol_(financeHeaders, ['quantity']);
    const docIdx = firstFinanceCol_(financeHeaders, ['doc_type_name']);
    const retailPriceIdx = firstFinanceCol_(financeHeaders, ['Сумма по текущей цене']);
    const dateIdx = firstFinanceCol_(financeHeaders, ['Дата', 'rr_dt']);
    if (artIdx < 0 || qtyIdx < 0 || docIdx < 0 || dateIdx < 0) {
      Logger.log('❌ Финансовые_ежедневные: нет sa_name / quantity / doc_type_name / Дата');
    } else {
    
    const financeTemp = {};
    
    for (let i = 1; i < financeValues.length; i++) {
      const row = financeValues[i];
      
      const dateObj = parseFinanceDate_(row[dateIdx]);
      if (!dateObj) continue;
      
      const time = dateObj.getTime();
      if (time < fromTime || time > toTime) continue;
      
      const art = row[artIdx];
      if (!art) continue;
      
      const quantity = parseFloat(row[qtyIdx]) || 0;
      const docType = row[docIdx];
      const retailPrice = retailPriceIdx >= 0 ? (parseFloat(row[retailPriceIdx]) || 0) : 0;
      
      if (!financeTemp[art]) {
        financeTemp[art] = {
          returns: 0,
          priceSum: 0,
          priceCount: 0
        };
      }
      
      if (docType === 'Возврат') {
        financeTemp[art].returns += Math.abs(quantity);
      }
      
      if (docType === 'Продажа' && retailPrice > 0) {
        financeTemp[art].priceSum += retailPrice;
        financeTemp[art].priceCount += 1;
      }
    }
    
    for (const art in financeTemp) {
      let foundKey = null;
      const artLower = art.toLowerCase();
      for (const key in result) {
        const keyArt = key.split('|')[0].toLowerCase();
        if (keyArt === artLower) {
          foundKey = key;
          break;
        }
      }
      
      if (!foundKey) continue;
      
      result[foundKey].returns = financeTemp[art].returns;
      
      if (financeTemp[art].priceCount > 0) {
        result[foundKey].avgPrice = Math.round((financeTemp[art].priceSum / financeTemp[art].priceCount) * 100) / 100;
      }
    }
    }
  }

  // ============ 3. ЗАГРУЗКА ТЕКУЩИХ ЦЕН С ЛИСТА "Товары" ============
  if (goodsSheet) {
    const goodsValues = goodsSheet.getDataRange().getValues();
    const goodsHeaders = goodsValues[0];
    
    let artIndex = -1;
    let priceIndex = -1;
    
    for (let i = 0; i < goodsHeaders.length; i++) {
      const header = goodsHeaders[i];
      if (header === 'Артикул продавца') artIndex = i;
      if (header === 'Цена со скидкой, ₽' || header === 'Цена со скидкой') priceIndex = i;
    }
    
    if (artIndex === -1) {
      artIndex = 2;
      Logger.log("Используем столбец 3 для артикула");
    }
    
    if (priceIndex === -1) {
      for (let i = 0; i < goodsHeaders.length; i++) {
        const header = goodsHeaders[i];
        if (header && header.toString().toLowerCase().indexOf('цена') !== -1) {
          priceIndex = i;
          Logger.log("Найдена колонка с ценой: " + header);
          break;
        }
      }
    }
    
    if (artIndex !== -1 && priceIndex !== -1) {
      const priceMap = {};
      for (let i = 1; i < goodsValues.length; i++) {
        let art = goodsValues[i][artIndex];
        if (!art) continue;
        
        art = art.toString().trim();
        
        const price = parseFloat(goodsValues[i][priceIndex]) || 0;
        priceMap[art.toLowerCase()] = price;
      }
      
      Logger.log(`Загружено цен с листа "Товары": ${Object.keys(priceMap).length}`);
      
      for (const key in result) {
        const art = key.split('|')[0];
        const artLower = art.toLowerCase();
        if (priceMap[artLower] !== undefined) {
          result[key].currentPrice = priceMap[artLower];
        }
      }
    } else {
      Logger.log("Не найдены колонки с артикулом или ценой на листе Товары");
    }
  }

  // ============ 4. 🔥 РУЧНАЯ ЦЕНА (из листа "Товар+Вес") ============
  if (manualPriceSheet) {
    const mValues = manualPriceSheet.getDataRange().getValues();
    const mHeaders = mValues[0];
    let aIdx = mHeaders.indexOf("Артикул");
    let pIdx = mHeaders.indexOf("Ручная цена");
    
    if (aIdx === -1) {
      for(let k=0; k<mHeaders.length; k++) {
        if(mHeaders[k].toString().toLowerCase().includes("артикул")) { aIdx = k; break; }
      }
    }
    if (pIdx === -1) {
      for(let k=0; k<mHeaders.length; k++) {
        if(mHeaders[k].toString().toLowerCase().includes("ручная цена")) { pIdx = k; break; }
      }
    }

    if (aIdx !== -1 && pIdx !== -1) {
      const map = {};
      for (let i = 1; i < mValues.length; i++) {
        let a = mValues[i][aIdx];
        if (!a) continue;
        const p = parseFloat(mValues[i][pIdx]);
        if (!isNaN(p)) map[a.toString().trim().toLowerCase()] = p;
      }
      for (const key in result) {
        const art = key.split('|')[0].toLowerCase();
        if (map[art] !== undefined) result[key].manualPrice = map[art];
      }
    }
  }

  // ============ 5. ЗАГРУЗКА ОСТАТКОВ ============
  if (stocksSheet) {
    const stocksValues = stocksSheet.getDataRange().getValues();
    const stocksHeaders = stocksValues[0];
    
    let artIndex = -1;
    let stockQtyIndex = -1;
    let dateIndex = -1;
    
    for (let i = 0; i < stocksHeaders.length; i++) {
      const header = stocksHeaders[i];
      if (header === 'Артикул продавца') artIndex = i;
      if (header === 'Остаток') stockQtyIndex = i;
      if (header === 'Дата') dateIndex = i;
    }
    
    if (artIndex !== -1 && stockQtyIndex !== -1 && dateIndex !== -1) {
      const stocksByArt = {};
      
      for (let i = 1; i < stocksValues.length; i++) {
        const art = stocksValues[i][artIndex];
        if (!art) continue;
        
        const qty = parseFloat(stocksValues[i][stockQtyIndex]) || 0;
        const dateValue = stocksValues[i][dateIndex];
        const dateTime = dateValue ? new Date(dateValue).getTime() : 0;
        
        const artKey = String(art).toLowerCase();
        
        if (!stocksByArt[artKey]) {
          stocksByArt[artKey] = {
            maxDate: -1,
            dateSums: {}
          };
        }
        
        const dateKey = dateTime;
        if (!stocksByArt[artKey].dateSums[dateKey]) {
          stocksByArt[artKey].dateSums[dateKey] = 0;
        }
        stocksByArt[artKey].dateSums[dateKey] += qty;
        
        if (dateTime > stocksByArt[artKey].maxDate) {
          stocksByArt[artKey].maxDate = dateTime;
        }
      }
      
      for (const key in result) {
        const art = key.split('|')[0];
        const artLower = art.toLowerCase();
        
        if (stocksByArt[artLower] && stocksByArt[artLower].maxDate !== -1) {
          const maxDate = stocksByArt[artLower].maxDate;
          const sumOnMaxDate = stocksByArt[artLower].dateSums[maxDate] || 0;
          result[key].stocks = sumOnMaxDate;
        }
      }
    }
  }

  // ============ 6. ФОРМИРОВАНИЕ ОТЧЕТА ============
  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];
  
  // 🔥 ДОБАВЛЕН СТОЛБЕЦ "Цена" в конец
  const headers = ["Артикул", "Наименование", ...dateList, "Итого", "Возвраты", "Остатки", "Ср. цена", "Текущ цена", "Цена"];
  output.push(headers);

  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0;
  let grandReturns = 0;
  let grandStocks = 0;
  
  let twAvg = 0, twCur = 0, twMan = 0, totalQty = 0;
  
  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));

  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split("|");
    const item = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    
    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const v = item.data[d] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    
    // 🔥 ДОБАВЛЕНА manualPrice
    row.push(sumRow);
    row.push(item.returns);
    row.push(item.stocks);
    row.push(item.avgPrice);
    row.push(item.currentPrice);
    row.push(item.manualPrice);
    
    dataRows.push(row);
    
    grandReturns += item.returns;
    grandStocks += item.stocks;
    
    // Накопление для средневзвешенных
    if (sumRow > 0) {
      twAvg += (item.avgPrice || 0) * sumRow;
      twCur += (item.currentPrice || 0) * sumRow;
      twMan += (item.manualPrice || 0) * sumRow;
      totalQty += sumRow;
    }
  }

  // Строка ИТОГО — цены без округления до целых, формат 0.00
  const avgT = totalQty > 0 ? twAvg / totalQty : 0;
  const curT = totalQty > 0 ? twCur / totalQty : 0;
  const manT = totalQty > 0 ? twMan / totalQty : 0;
  
  const totalRow = ["Итого", "", ...colTotals, grand, grandReturns, grandStocks, avgT, curT, manT];
  output.push(totalRow);
  
  for (const row of dataRows) {
    output.push(row);
  }

  // ============ 7. ВЫВОД НА ЛИСТ ============
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    let existingLastRow = sheet.getLastRow();
    if (existingLastRow >= 3) {
      sheet.getRange(3, 1, existingLastRow - 2, sheet.getLastColumn()).clear();
    }
  }
  
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  // Применяем форматирование ко всем заголовкам
  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center");
  headerRange.setWrap(true);
  
  // Поворачиваем колонки с датами
  if (dateList.length > 0) {
    const dateHeaderRange = sheet.getRange(3, 3, 1, dateList.length);
    dateHeaderRange.setTextRotation(90);
  }
  
  const итогоCol = dateList.length + 3;
  const возвратыCol = dateList.length + 4;
  const остаткиCol = dateList.length + 5;
  
  sheet.getRange(3, итогоCol, 1, 1).setTextRotation(90);
  sheet.getRange(3, возвратыCol, 1, 1).setTextRotation(0);
  sheet.getRange(3, остаткиCol, 1, 1).setTextRotation(0);
  
  if (sheet.getFilter()) sheet.getFilter().remove();
  headerRange.createFilter();

  // ============ 8. ФОРМАТИРОВАНИЕ СТРОКИ ИТОГО ============
  sheet.getRange(4, 1, 1, output[0].length)
    .setFontWeight("bold")
    .setBackground("#e8e8e8")
    .setFontFamily("Calibri")
    .setFontSize(10);

  // ============ 9. НАСТРОЙКА ШИРИНЫ КОЛОНОК ============
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 300);
  
  for (let col = 3; col <= dateList.length + 2; col++) {
    sheet.setColumnWidth(col, 30);
  }
  
  sheet.setColumnWidth(dateList.length + 3, 30);
  sheet.setColumnWidth(dateList.length + 4, 55);
  sheet.setColumnWidth(dateList.length + 5, 55);
  sheet.setColumnWidth(dateList.length + 6, 55);
  sheet.setColumnWidth(dateList.length + 7, 55);
  sheet.setColumnWidth(dateList.length + 8, 55); // 🔥 ДЛЯ "Цена"
  
  sheet.setRowHeight(3, 55);

  // ============ 10. ФОРМАТЫ ЧИСЕЛ ============
  const startR = 4;
  const cntR = output.length - 2;
  const colAvg = dateList.length + 6 + 1;
  const colCur = dateList.length + 7 + 1;
  const colMan = dateList.length + 8 + 1; // 🔥 НОВАЯ КОЛОНКА
  
  if (cntR > 0) {
    const totalRowNum = 4;
    
    // Столбцы с датами в строке ИТОГО - целые числа
    sheet.getRange(totalRowNum, 3, 1, dateList.length)
      .setNumberFormat("0");
    
    // Колонка "Итого" в строке ИТОГО - целое число
    sheet.getRange(totalRowNum, dateList.length + 3, 1, 1)
      .setNumberFormat("0");
    
    // Цены — всегда 0.00, без округления до целых
    sheet.getRange(totalRowNum, colAvg, 1, 1).setNumberFormat("#,##0.00");
    sheet.getRange(totalRowNum, colCur, 1, 1).setNumberFormat("#,##0.00");
    sheet.getRange(totalRowNum, colMan, 1, 1).setNumberFormat("#,##0.00");
    
    // Данные (строки 5 и ниже) - для цен с копейками
    if (cntR > 1) {
      const dataStartRow = 5;
      const dataRowsCount = cntR - 1;
      
      sheet.getRange(dataStartRow, colAvg, dataRowsCount, 1)
        .setNumberFormat("#,##0.00");
      sheet.getRange(dataStartRow, colCur, dataRowsCount, 1)
        .setNumberFormat("#,##0.00");
      sheet.getRange(dataStartRow, colMan, dataRowsCount, 1)
        .setNumberFormat("#,##0.00"); // 🔥 Цена с копейками для данных
    }
  }

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Количество заказов' обновлён");
}

function сформироватьОтчет_ЗаказыFBS_Количество() {
  сформироватьОтчет_ЗаказыFBS_('qty');
}

function сформироватьОтчет_ЗаказыFBS_Вес() {
  сформироватьОтчет_ЗаказыFBS_('weight');
}

function сформироватьОтчет_ЗаказыFBS_(mode) {
  const isWeight = mode === 'weight';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = isWeight ? 'Вес заказов FBS' : 'К-во Заказов FBS';
  const sheet = fbsQtyEnsureSheet_(ss, sheetName);
  const source = ss.getSheetByName('FBS Заказы');
  if (!source || source.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ Лист «FBS Заказы» пуст или не найден. Сначала выгрузите заказы FBS.');
    return;
  }

  const dateFrom = new Date(sheet.getRange('A1').getValue());
  const dateTo = new Date(sheet.getRange('A2').getValue());
  if (isNaN(dateFrom.getTime()) || isNaN(dateTo.getTime())) {
    SpreadsheetApp.getUi().alert('❌ Укажите период в ячейках A1 (с) и A2 (по) на листе «' + sheetName + '».');
    return;
  }
  dateFrom.setHours(0, 0, 0, 0);
  dateTo.setHours(23, 59, 59, 999);

  const tz = Session.getScriptTimeZone();
  const goodsNameByNm = fbsQtyGoodsNameMap_(ss);
  const goodsArtByNm = fbsQtyGoodsArtMap_(ss);
  const goodsNameByArt = fbsQtyGoodsNameByArtMap_(ss);

  const table = source.getDataRange().getValues();
  const headers = table[0];
  const artIdx = firstFinanceCol_(headers, ['article', 'Арт. в заказе']);
  const nameIdx = firstFinanceCol_(headers, ['Название', 'title', 'Имя товара']);
  const composedArtIdx = firstFinanceCol_(headers, ['Артикул']);
  const nmIdx = firstFinanceCol_(headers, ['nmId', 'nmID', 'Артикул WB']);
  const dateIdx = firstFinanceCol_(headers, ['Дата']);
  const createdIdx = firstFinanceCol_(headers, ['createdAt']);
  const statusNameIdx = firstFinanceCol_(headers, ['wbStatus_name']);
  const statusIdx = firstFinanceCol_(headers, ['wbStatus']);
  const weightIdx = firstFinanceCol_(headers, ['Вес, кг за 1 шт сумма']);
  if (artIdx < 0 && nmIdx < 0) {
    SpreadsheetApp.getUi().alert('❌ На «FBS Заказы» нет столбца article / nmId.');
    return;
  }
  if (isWeight && weightIdx < 0) {
    SpreadsheetApp.getUi().alert('❌ На «FBS Заказы» нет столбца «Вес, кг за 1 шт сумма». Сначала соберите лист FBS Заказы.');
    return;
  }

  const result = {};
  for (let i = 1; i < table.length; i++) {
    const row = table[i];
    const dt = parseFinanceDate_(dateIdx >= 0 ? row[dateIdx] : '') ||
      parseFinanceDate_(createdIdx >= 0 ? row[createdIdx] : '');
    if (!dt) continue;
    const time = dt.getTime();
    if (time < dateFrom.getTime() || time > dateTo.getTime()) continue;

    const nmId = nmIdx >= 0 ? row[nmIdx] : '';
    const nmKey = nmId != null && nmId !== '' ? String(nmId).trim().toLowerCase() : '';
    let art = artIdx >= 0 ? row[artIdx] : '';
    if (!art && composedArtIdx >= 0) art = row[composedArtIdx];
    if (!art && nmKey) art = goodsArtByNm[nmKey] || '';
    let name = nameIdx >= 0 ? row[nameIdx] : '';
    art = art != null ? String(art).trim() : '';
    name = name != null ? String(name).trim() : '';
    if (!name || name === '-' || name === art) {
      if (nmKey && goodsNameByNm[nmKey]) name = String(goodsNameByNm[nmKey] || '').trim();
    }
    if (!name || name === '-' || name === art) {
      var artKey = art.toLowerCase();
      if (goodsNameByArt[artKey]) name = String(goodsNameByArt[artKey] || '').trim();
      else if (typeof normalizeArt === 'function') {
        var nArt = String(normalizeArt(art) || '').toLowerCase();
        if (nArt && goodsNameByArt[nArt]) name = String(goodsNameByArt[nArt] || '').trim();
      }
    }
    if (!art || art === '-') continue;
    if (!name || name === '-') name = art;

    const val = isWeight ? (parseFloat(row[weightIdx]) || 0) : 1;
    const d = Utilities.formatDate(dt, tz, 'dd.MM');
    const key = art + '|' + name;
    if (!result[key]) result[key] = { data: {}, sold: 0, transit: 0, cancel: 0, stocks: 0 };
    result[key].data[d] = (result[key].data[d] || 0) + val;

    const kind = fbsQtyStatusKind_(
      statusNameIdx >= 0 ? row[statusNameIdx] : '',
      statusIdx >= 0 ? row[statusIdx] : ''
    );
    if (kind === 'sold') result[key].sold += val;
    else if (kind === 'transit') result[key].transit += val;
    else if (kind === 'cancel') result[key].cancel += val;
  }

  const dateList = [];
  const cur = new Date(dateFrom.getFullYear(), dateFrom.getMonth(), dateFrom.getDate());
  const end = new Date(dateTo.getFullYear(), dateTo.getMonth(), dateTo.getDate());
  while (cur.getTime() <= end.getTime()) {
    dateList.push(Utilities.formatDate(cur, tz, 'dd.MM'));
    cur.setDate(cur.getDate() + 1);
  }

  const stockMap = fbsQtyLoadStocks_(ss, isWeight);
  for (const key in result) {
    result[key].stocks = fbsQtyStockOf_(stockMap, key.split('|')[0]);
  }

  const output = [];
  const trailHeaders = ['Итого', 'Остатки', 'Продано', 'В пути', 'Отменено'];
  output.push(['Артикул', 'Наименование'].concat(dateList).concat(trailHeaders));

  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0, grandSold = 0, grandTransit = 0, grandCancel = 0, grandStocks = 0;
  const keys = Object.keys(result).sort(function(a, b) {
    return a.split('|')[0].localeCompare(b.split('|')[0], 'ru');
  });

  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split('|');
    const item = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    for (let j = 0; j < dateList.length; j++) {
      const v = item.data[dateList[j]] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    row.push(sumRow, item.stocks || 0, item.sold, item.transit, item.cancel);
    dataRows.push(row);
    grandSold += item.sold;
    grandTransit += item.transit;
    grandCancel += item.cancel;
    grandStocks += item.stocks || 0;
  }

  output.push(['Итого', ''].concat(colTotals).concat(
    [grand, grandStocks, grandSold, grandTransit, grandCancel]
  ));
  for (let r = 0; r < dataRows.length; r++) output.push(dataRows[r]);

  const existingLastRow = sheet.getLastRow();
  const existingLastCol = Math.max(sheet.getLastColumn(), 1);
  if (existingLastRow >= 3) {
    try { sheet.getRange(1, 3, 2, Math.max(existingLastCol - 2, 1)).breakApart(); } catch (e) {}
    sheet.getRange(3, 1, existingLastRow - 2, existingLastCol).clear();
  }

  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);
  outRange.setFontFamily('Calibri').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment('left');

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight('bold').setBackground('#f1f1f1').setHorizontalAlignment('center').setWrap(true);
  if (dateList.length > 0) {
    sheet.getRange(3, 3, 1, dateList.length).setNumberFormat('@').setTextRotation(90);
  }

  const итогоCol = dateList.length + 3;
  const trailCount = 4;
  sheet.getRange(3, итогоCol).setTextRotation(90);
  sheet.getRange(3, итогоCol + 1, 1, trailCount).setTextRotation(0);

  if (sheet.getFilter()) sheet.getFilter().remove();
  try { headerRange.createFilter(); } catch (eF) {}

  sheet.getRange(4, 1, 1, output[0].length)
    .setFontWeight('bold')
    .setBackground('#e8e8e8')
    .setFontFamily('Calibri')
    .setFontSize(10);

  const dateWidth = isWeight ? 40 : 30;
  const итогоWidth = isWeight ? 40 : 30;
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 300);
  for (let col = 3; col <= dateList.length + 2; col++) sheet.setColumnWidth(col, dateWidth);
  sheet.setColumnWidth(итогоCol, итогоWidth);
  sheet.setColumnWidth(итогоCol + 1, isWeight ? 55 : 70);
  sheet.setColumnWidth(итогоCol + 2, isWeight ? 55 : 70);
  sheet.setColumnWidth(итогоCol + 3, isWeight ? 55 : 70);
  sheet.setColumnWidth(итогоCol + 4, isWeight ? 55 : 80);
  sheet.setRowHeight(3, 55);

  const cntR = output.length - 1;
  if (cntR > 0 && dateList.length > 0) {
    sheet.getRange(4, 3, cntR, dateList.length + 1 + trailCount)
      .setNumberFormat(isWeight ? '#,##0.000' : '0');
  }

  ss.setActiveSheet(sheet);
  trimReportExtraColumns_(sheet);
  if (sheet.getLastRow() >= 5) formatCurrentSheet(sheet);

  var lastDataRow = 2 + output.length;
  if (isWeight) applyFbsWeightDateCF_(sheet, dateList.length, lastDataRow);
  else applyFbsQuantityDateCF_(sheet, dateList.length, lastDataRow);
  applyFbsDetailsHeader_(sheet, dateList.length);

  SpreadsheetApp.getUi().alert(isWeight
    ? '✅ Отчёт «Вес заказов FBS» обновлён'
    : '✅ Отчёт «К-во Заказов FBS» обновлён');
}

function fbsQtyEnsureSheet_(ss, sheetName) {
  sheetName = sheetName || 'К-во Заказов FBS';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet && sheetName === 'К-во Заказов FBS') {
    var old = ss.getSheetByName('Кол-во Заказов FBS');
    if (old) {
      old.setName('К-во Заказов FBS');
      return old;
    }
  }
  if (sheet) return sheet;
  sheet = ss.insertSheet(sheetName, ss.getSheets().length);
  var src = ss.getSheetByName('К-во Заказов FBS') ||
    ss.getSheetByName('Вес заказов') ||
    ss.getSheetByName('К-во заказов');
  if (src) {
    sheet.getRange('A1').setValue(src.getRange('A1').getValue());
    sheet.getRange('A2').setValue(src.getRange('A2').getValue());
  }
  sheet.getRange('A1:A2').setNumberFormat('dd.mm.yyyy');
  return sheet;
}

function fbsQtyLoadStocks_(ss, useWeight) {
  var map = {};
  var sheet = ss.getSheetByName('FBS Остатки');
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dIdx = firstFinanceCol_(headers, ['Дата']);
  var aIdx = firstFinanceCol_(headers, ['Артикул', 'vendorCode', 'Арт. продавца', 'article']);
  var vIdx = firstFinanceCol_(headers, ['vendorCode', 'Арт. продавца']);
  var qIdx = firstFinanceCol_(headers, ['amount', 'Остаток', 'quantity']);
  var wIdx = firstFinanceCol_(headers, ['Вес, кг за 1 шт сумма']);
  var valIdx = useWeight ? wIdx : qIdx;
  if (dIdx < 0 || (aIdx < 0 && vIdx < 0) || valIdx < 0) return map;

  var latest = null;
  var i;
  for (i = 1; i < data.length; i++) {
    var dt = parseFinanceDate_(data[i][dIdx]);
    if (dt && (!latest || dt.getTime() > latest.getTime())) latest = dt;
  }
  if (!latest) return map;
  var tz = Session.getScriptTimeZone();
  var latestDay = Utilities.formatDate(latest, tz, 'yyyy-MM-dd');

  for (i = 1; i < data.length; i++) {
    var dt2 = parseFinanceDate_(data[i][dIdx]);
    if (!dt2) continue;
    if (Utilities.formatDate(dt2, tz, 'yyyy-MM-dd') !== latestDay) continue;
    var art = aIdx >= 0 ? String(data[i][aIdx] || '').trim() : '';
    if (!art && vIdx >= 0) art = String(data[i][vIdx] || '').trim();
    if (!art) continue;
    var qty = parseFloat(String(data[i][valIdx]).replace(',', '.')) || 0;
    var k1 = art.toLowerCase();
    map[k1] = (map[k1] || 0) + qty;
    if (typeof normalizeArt === 'function') {
      var k2 = String(normalizeArt(art) || '').toLowerCase();
      if (k2 && k2 !== k1) map[k2] = (map[k2] || 0) + qty;
    }
  }
  return map;
}

function fbsQtyStockOf_(map, art) {
  if (!map) return 0;
  var a = String(art || '').trim();
  if (!a) return 0;
  if (map[a.toLowerCase()] != null) return map[a.toLowerCase()];
  if (typeof normalizeArt === 'function') {
    var n = String(normalizeArt(a) || '').toLowerCase();
    if (n && map[n] != null) return map[n];
  }
  return 0;
}

function fbsQtyStatusKind_(statusName, statusCode) {
  var s = String(statusName || '').toLowerCase();
  var c = String(statusCode || '').toLowerCase();
  if (c === 'sold' || s === 'продано') return 'sold';
  if (c === 'sorted' || s.indexOf('отсортировано') >= 0) return 'transit';
  if (c === 'ready_for_pickup' || s.indexOf('пвз') >= 0) return 'transit';
  if (c.indexOf('cancel') >= 0 || s.indexOf('отменено') >= 0) return 'cancel';
  return '';
}

function fbsQtyGoodsNameByArtMap_(ss) {
  var map = {};
  var sh = ss.getSheetByName('Товары');
  if (!sh || sh.getLastRow() < 2) return map;
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var artCol = firstFinanceCol_(h, ['vendorCode', 'Ваш артикул', 'Арт. продавца']);
  var nameCol = firstFinanceCol_(h, ['Название', 'title']);
  if (artCol < 0 || nameCol < 0) return map;
  for (var r = 1; r < data.length; r++) {
    var art = String(data[r][artCol] || '').trim();
    var name = String(data[r][nameCol] || '').trim();
    if (!art || !name) continue;
    map[art.toLowerCase()] = name;
    if (typeof normalizeArt === 'function') {
      var n = String(normalizeArt(art) || '').toLowerCase();
      if (n && !map[n]) map[n] = name;
    }
  }
  return map;
}

function fbsQtyGoodsNameMap_(ss) {
  var map = {};
  var sh = ss.getSheetByName('Товары');
  if (!sh || sh.getLastRow() < 2) return map;
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var nmCol = 0;
  var nameCol = 1;
  for (var i = 0; i < h.length; i++) {
    var t = String(h[i] || '');
    if (typeof wbIs_ === 'function') {
      if (wbIs_(t, 'nmID')) nmCol = i;
    } else if (t === 'nmID' || t === 'nmId') nmCol = i;
    if (t === 'Название') nameCol = i;
  }
  for (var r = 1; r < data.length; r++) {
    if (data[r][nmCol] == null || data[r][nmCol] === '') continue;
    map[String(data[r][nmCol]).trim().toLowerCase()] = data[r][nameCol] || '';
  }
  return map;
}

function fbsQtyGoodsArtMap_(ss) {
  var map = {};
  var sh = ss.getSheetByName('Товары');
  if (!sh || sh.getLastRow() < 2) return map;
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var nmCol = 0;
  var artCol = 2;
  for (var i = 0; i < h.length; i++) {
    var t = String(h[i] || '');
    if (typeof wbIs_ === 'function') {
      if (wbIs_(t, 'nmID')) nmCol = i;
    } else if (t === 'nmID' || t === 'nmId') nmCol = i;
    if (t === 'Ваш артикул' || (typeof wbIs_ === 'function' && wbIs_(t, 'vendorCode'))) artCol = i;
  }
  for (var r = 1; r < data.length; r++) {
    if (data[r][nmCol] == null || data[r][nmCol] === '') continue;
    map[String(data[r][nmCol]).trim().toLowerCase()] = data[r][artCol] || '';
  }
  return map;
}

/* отключено
function сформироватьОтчет_Заказы_Сумма() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("ЗаказыЛК");
  const startSheet = ss.getSheetByName("Сумма заказов");

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const data = source.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "Сумма заказов";
  const valueCol = 7; // 
  
  const result = {};
  const datesObj = {};

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const dt = row[2]; // колонка (дата)
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;

    const art = row[3];   // колонка (артикул)
    const name = row[8]; // колонка (наименование)
    const val = row[valueCol] || 0; // сумма из колонки 7
    const d = Utilities.formatDate(dt, tz, "dd.MM");

    // 🔧 ФИЛЬТРУЕМ ПУСТЫЕ АРТИКУЛЫ И НАИМЕНОВАНИЯ
    if (!art || art === '' || art === ' ' || art === '-' || !name || name === '' || name === ' ') {
      continue; // пропускаем пустые строки
    }
    
    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    result[key][d] = (result[key][d] || 0) + val;
  }

  // const dateList = Object.keys(datesObj).sort();
  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];
  
  // Заголовки
  output.push(["Артикул", "Наименование", ...dateList, "Итого"]);

  // Рассчитываем итоги по столбцам
  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0;
  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));

  // Собираем все строки данных
  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const parts = keys[i].split("|");
    const rowObj = result[keys[i]];
    const row = [parts[0], parts[1]];
    let sumRow = 0;
    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const v = rowObj[d] || 0;
      row.push(v);
      sumRow += v;
      colTotals[j] += v;
      grand += v;
    }
    row.push(sumRow);
    dataRows.push(row);
  }

  // Добавляем строку Итого (после заголовка, перед данными)
  output.push(["Итого", "", ...colTotals, grand]);
  
  // Добавляем все строки данных
  for (const row of dataRows) {
    output.push(row);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (sheet) clearReportBody_(sheet);
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  // const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  // headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setTextRotation(90);
  // if (sheet.getFilter()) sheet.getFilter().remove();
  // headerRange.createFilter();

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center");

  // Только для колонок с датами (начиная с 3-й) делаем поворот
  if (output[0].length > 2) {
    const dateHeaderRange = sheet.getRange(3, 3, 1, output[0].length - 2);
    dateHeaderRange.setTextRotation(90);
  }

  sheet.getRange(4, 1, 1, output[0].length).setFontWeight("bold").setBackground("#e8e8e8");

  for (let col = 3; col <= output[0].length; col++) sheet.setColumnWidth(col, 55);
  sheet.setRowHeight(3, 45);

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Сумма заказов' обновлён");
}
*/

function сформироватьОтчет_Заказы_Вес() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startSheet = ss.getSheetByName("Вес заказов");
  const financeSheet = ss.getSheetByName("Финансовые_ежедневные");
  const stocksSheet = ss.getSheetByName("Остатки");
  const goodsSheet = ss.getSheetByName("Товары");
  const manualPriceSheet = ss.getSheetByName("Товар+Вес"); // 🔥 Новый лист

  const dateFrom = new Date(startSheet.getRange("A1").getValue());
  const dateTo = new Date(startSheet.getRange("A2").getValue());

  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const sheetName = "Вес заказов";
  
  const result = {};
  const datesObj = {};

  // ============ 1. ОСНОВНОЙ БЛОК (из Воронка продаж) ============
  if (!funnelOrdersReportFill_(ss, dateFrom, dateTo, 'Вес, кг за 1 шт сумма заказов', result, datesObj)) return;

  // ============ 2. ФИНАНСОВЫЕ ДАННЫЕ ============
  if (financeSheet) {
    const fValues = financeSheet.getDataRange().getValues();
    const fHeaders = fValues[0];
    const artIdx = firstFinanceCol_(fHeaders, ['sa_name']);
    const qtyIdx = firstFinanceCol_(fHeaders, ['quantity']);
    const docIdx = firstFinanceCol_(fHeaders, ['doc_type_name']);
    const priceIdx = firstFinanceCol_(fHeaders, ['Сумма по текущей цене']);
    const dateIdx = firstFinanceCol_(fHeaders, ['Дата']);
    const finTemp = {};
    
    if (artIdx >= 0 && qtyIdx >= 0 && docIdx >= 0 && dateIdx >= 0) {
    for (let i = 1; i < fValues.length; i++) {
      const r = fValues[i];
      const dt = r[dateIdx];
      if (!dt || !(dt instanceof Date)) continue;
      if (dt.getTime() < fromTime || dt.getTime() > toTime) continue;
      
      const art = r[artIdx];
      if (!art) continue;
      if (!finTemp[art]) finTemp[art] = { ret: 0, pSum: 0, pCnt: 0 };
      
      const qty = Math.abs(parseFloat(r[qtyIdx]) || 0);
      if (r[docIdx] === 'Возврат') finTemp[art].ret += qty;
      if (r[docIdx] === 'Продажа') {
        const p = priceIdx >= 0 ? (parseFloat(r[priceIdx]) || 0) : 0;
        if (p > 0) { finTemp[art].pSum += p; finTemp[art].pCnt++; }
      }
    }
    
    for (const art in finTemp) {
      const artL = art.toLowerCase();
      for (const key in result) {
        if (key.split('|')[0].toLowerCase() === artL) {
          result[key].returns = finTemp[art].ret;
          if (finTemp[art].pCnt > 0) result[key].avgPrice = Math.round((finTemp[art].pSum / finTemp[art].pCnt) * 100) / 100;
          break;
        }
      }
    }
    }
  }

  // ============ 3. ТЕКУЩАЯ ЦЕНА (из "Товары") ============
  if (goodsSheet) {
    const gValues = goodsSheet.getDataRange().getValues();
    const gHeaders = gValues[0];
    let aIdx = -1, pIdx = -1;
    
    for (let i = 0; i < gHeaders.length; i++) {
      const h = gHeaders[i].toString().toLowerCase();
      if (h.includes('артикул продавца') || h === 'артикул') aIdx = i;
      if (h.includes('цена со скидкой') || h === 'цена') pIdx = i;
    }
    if (aIdx === -1) aIdx = 2;
    if (pIdx === -1) {
      for (let i = 0; i < gHeaders.length; i++) if (gHeaders[i].toString().toLowerCase().includes('цена')) { pIdx = i; break; }
    }
    
    if (aIdx !== -1 && pIdx !== -1) {
      const map = {};
      for (let i = 1; i < gValues.length; i++) {
        let a = gValues[i][aIdx];
        if (!a) continue;
        const p = parseFloat(gValues[i][pIdx]);
        if (!isNaN(p)) map[a.toString().trim().toLowerCase()] = p;
      }
      for (const key in result) {
        const art = key.split('|')[0].toLowerCase();
        if (map[art] !== undefined) result[key].currentPrice = map[art];
      }
    }
  }

  // ============ 4. 🔥 РУЧНАЯ ЦЕНА (из "Товар+Вес") ============
  if (manualPriceSheet) {
    const mValues = manualPriceSheet.getDataRange().getValues();
    const mHeaders = mValues[0];
    let aIdx = mHeaders.indexOf("Артикул");
    let pIdx = mHeaders.indexOf("Ручная цена");
    
    if (aIdx === -1) for(let k=0; k<mHeaders.length; k++) if(mHeaders[k].toString().toLowerCase().includes("артикул")) { aIdx = k; break; }
    if (pIdx === -1) for(let k=0; k<mHeaders.length; k++) if(mHeaders[k].toString().toLowerCase().includes("ручная цена")) { pIdx = k; break; }

    if (aIdx !== -1 && pIdx !== -1) {
      const map = {};
      for (let i = 1; i < mValues.length; i++) {
        let a = mValues[i][aIdx];
        if (!a) continue;
        const p = parseFloat(mValues[i][pIdx]);
        if (!isNaN(p)) map[a.toString().trim().toLowerCase()] = p;
      }
      for (const key in result) {
        const art = key.split('|')[0].toLowerCase();
        if (map[art] !== undefined) result[key].manualPrice = map[art];
      }
    }
  }

  // ============ 5. ОСТАТКИ ============
  if (stocksSheet) {
    const sValues = stocksSheet.getDataRange().getValues();
    const sHeaders = sValues[0];
    let aIdx = -1, qIdx = -1, dIdx = -1;
    for (let i = 0; i < sHeaders.length; i++) {
      if (sHeaders[i] === 'Артикул продавца') aIdx = i;
      if (sHeaders[i] === 'Остаток') qIdx = i;
      if (sHeaders[i] === 'Дата') dIdx = i;
    }
    
    if (aIdx !== -1 && qIdx !== -1 && dIdx !== -1) {
      const stMap = {};
      for (let i = 1; i < sValues.length; i++) {
        const a = sValues[i][aIdx];
        if (!a) continue;
        const q = parseFloat(sValues[i][qIdx]) || 0;
        const t = sValues[i][dIdx] ? new Date(sValues[i][dIdx]).getTime() : 0;
        const k = a.toString().toLowerCase();
        if (!stMap[k]) stMap[k] = { maxT: -1, sums: {} };
        if (!stMap[k].sums[t]) stMap[k].sums[t] = 0;
        stMap[k].sums[t] += q;
        if (t > stMap[k].maxT) stMap[k].maxT = t;
      }
      for (const key in result) {
        const art = key.split('|')[0].toLowerCase();
        if (stMap[art] && stMap[art].maxT !== -1) result[key].stocks = stMap[art].sums[stMap[art].maxT] || 0;
      }
    }
  }

  // ============ 6. ФОРМИРОВАНИЕ ОТЧЕТА ============
  const curYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const [da, ma] = a.split('.');
    const [db, mb] = b.split('.');
    return new Date(curYear, ma-1, da) - new Date(curYear, mb-1, db);
  });

  // 🔥 ДОБАВЛЕН СТОЛБЕЦ "Цена" В КОНЕЦ
  const headers = ["Артикул", "Наименование", ...dateList, "Итого", "Возвраты", "Остатки", "Ср. цена", "Текущ цена", "Цена"];
  const output = [headers];

  const colTotals = new Array(dateList.length).fill(0);
  let grand = 0, grandRet = 0, grandSt = 0;
  
  // 🔥 Переменные для средневзвешенных цен
  let twAvg = 0, twCur = 0, twMan = 0, totalQty = 0;

  const keys = Object.keys(result).sort((a, b) => a.split("|")[0].localeCompare(b.split("|")[0], 'ru'));
  const dataRows = [];

  for (const key of keys) {
    const [art, name] = key.split("|");
    const item = result[key];
    const row = [art, name];
    let rowSum = 0;
    
  // ✅ СТАЛО (используем индекс j):
  for (let j = 0; j < dateList.length; j++) {
    const d = dateList[j];
    const v = item.data[d] || 0;
    row.push(v);
    rowSum += v;
    colTotals[j] += v;  // ← теперь j определён и суммы считаются
    grand += v;
  }
    
    // 🔥 Добавляем manualPrice в строку
    row.push(rowSum, item.returns, item.stocks, item.avgPrice, item.currentPrice, item.manualPrice);
    dataRows.push(row);
    
//    grand += rowSum; 
    grandRet += item.returns; grandSt += item.stocks;
    
    // 🔥 Накопление для средневзвешенных (по весу, так как это отчёт по весу)
    if (rowSum > 0) {
      twAvg += (item.avgPrice || 0) * rowSum;
      twCur += (item.currentPrice || 0) * rowSum;
      twMan += (item.manualPrice || 0) * rowSum;
      totalQty += rowSum;
    }
  }

  // 🔥 Строка ИТОГО с расчётом средневзвешенной "Цены"
  const avgT = totalQty > 0 ? twAvg / totalQty : 0;
  const curT = totalQty > 0 ? twCur / totalQty : 0;
  const manT = totalQty > 0 ? twMan / totalQty : 0;
  
  const totalRow = ["Итого", "", ...colTotals, grand, grandRet, grandSt, avgT, curT, manT];
  
  output.push(totalRow);
  for (const r of dataRows) output.push(r);

  // ============ 7. ВЫВОД И ФОРМАТИРОВАНИЕ ============
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clear();
  
  const range = sheet.getRange(3, 1, output.length, headers.length);
  range.setValues(output).setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  
  const hdrRange = sheet.getRange(3, 1, 1, headers.length);
  hdrRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setWrap(true);
  if (dateList.length > 0) sheet.getRange(3, 3, 1, dateList.length).setTextRotation(90);
  sheet.getRange(3, dateList.length + 3, 1, 1).setTextRotation(90);
  
  if (sheet.getFilter()) sheet.getFilter().remove();
  hdrRange.createFilter();

  // Ширина колонок
  sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 300);
  for (let c = 3; c <= dateList.length + 2; c++) sheet.setColumnWidth(c, 40);
  sheet.setColumnWidth(dateList.length + 3, 40);  // Итого
  sheet.setColumnWidth(dateList.length + 4, 55);  // Возвраты
  sheet.setColumnWidth(dateList.length + 5, 55);  // Остатки
  sheet.setColumnWidth(dateList.length + 6, 55);  // Ср. цена
  sheet.setColumnWidth(dateList.length + 7, 55);  // Текущ цена
  sheet.setColumnWidth(dateList.length + 8, 55);  // 🔥 Цена
  sheet.setRowHeight(3, 55);

  // Форматы чисел
  const startR = 4;
  const cntR = output.length - 2;
  const colAvg = dateList.length + 6 + 1;
  const colCur = dateList.length + 7 + 1;
  const colMan = dateList.length + 8 + 1; // 🔥 Новая колонка
  
  if (cntR > 0) {
    sheet.getRange(startR, colAvg, cntR, 1).setNumberFormat("#,##0.00");
    sheet.getRange(startR, colCur, cntR, 1).setNumberFormat("#,##0.00");
    sheet.getRange(startR, colMan, cntR, 1).setNumberFormat("#,##0.00");
  }

  if (typeof formatCurrentSheet === 'function') formatCurrentSheet();
  SpreadsheetApp.getUi().alert("✅ Отчёт 'Вес заказов' обновлён");
}

function сформироватьОтчетРекламаИПродажи2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Реклама");
//   const adsSheet = ss.getSheetByName("Данные_для_анализа_реклама");
//   const salesSheet = ss.getSheetByName("Данные_для_анализа");
  const adsSheet = ss.getSheetByName("Рекламные_расходы");
  const salesSheet = ss.getSheetByName("Финансовые_ежедневные");

  const dateFrom = new Date(sheet.getRange("A1").getValue());
  const dateTo = new Date(sheet.getRange("A2").getValue());
  const tz = Session.getScriptTimeZone();
  const fromTime = dateFrom.getTime();
  const toTime = dateTo.getTime();

  const result = {};
  const datesObj = {};

function norm(text) {
  return (text ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{N} \.]/gu, "")  // ← добавили экранированную точку \.
    .toUpperCase();
}

  // Реклама
  const adsData = adsSheet.getDataRange().getValues();
  for (let i = 2; i < adsData.length; i++) {
    const row = adsData[i];
    const dtRaw = row[16];
    if (!dtRaw) continue;
    const dt = (dtRaw instanceof Date && !isNaN(dtRaw.getTime())) ? dtRaw : parseFinanceDate_(dtRaw);
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;
    const art = norm(row[15]);
    const name = norm(row[4]);
    const val = row[17] || 0;

    if (!art || art === '' || art === ' ' || art === '-' || art === '0' ||
        !name || name === '' || name === ' ' || name === '-' || name === '0') {
      continue;
    }

    const d = Utilities.formatDate(dt, tz, "dd.MM");
    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    if (!result[key][d]) result[key][d] = { ads: 0, sales: 0 };
    result[key][d].ads += val;
  }

  // Продажи — как раньше: дата из rr_dt (кол. 27), цикл с 3-й строки листа
  const salesData = salesSheet.getDataRange().getValues();
  const salesHeaders = salesData[0];
  const sDateIdx = firstFinanceCol_(salesHeaders, ['rr_dt', 'Дата']);
  const sArtIdx = firstFinanceCol_(salesHeaders, ['sa_name']);
  const sNameIdx = firstFinanceCol_(salesHeaders, ['Имя товара']);
  const sValIdx = firstFinanceCol_(salesHeaders, ['Сумма продаж с учетом возвратов']);
  if (sDateIdx < 0 || sArtIdx < 0 || sNameIdx < 0 || sValIdx < 0) {
    SpreadsheetApp.getUi().alert('❌ На листе «Финансовые_ежедневные» нет нужных столбцов для отчёта Реклама');
    return;
  }
  for (let i = 2; i < salesData.length; i++) {
    const row = salesData[i];
    const dtRaw = row[sDateIdx];
    if (!dtRaw) continue;
    const dt = (dtRaw instanceof Date && !isNaN(dtRaw.getTime())) ? dtRaw : parseFinanceDate_(dtRaw);
    if (!dt) continue;
    const time = dt.getTime();
    if (time < fromTime || time > toTime) continue;
    const art = norm(row[sArtIdx]);
    const name = norm(row[sNameIdx]);
    const val = row[sValIdx] || 0;

    if (!art || art === '' || art === ' ' || art === '-' || art === '0' ||
        !name || name === '' || name === ' ' || name === '-' || name === '0') {
      continue;
    }

    const d = Utilities.formatDate(dt, tz, "dd.MM");
    datesObj[d] = true;
    const key = art + "|" + name;
    if (!result[key]) result[key] = {};
    if (!result[key][d]) result[key][d] = { ads: 0, sales: 0 };
    result[key][d].sales += val;
  }

  // const dateList = Object.keys(datesObj).sort();
  const currentYear = new Date().getFullYear();
  const dateList = Object.keys(datesObj).sort((a, b) => {
    const partsA = a.split('.');
    const partsB = b.split('.');
    const dateA = new Date(currentYear, parseInt(partsA[1]) - 1, parseInt(partsA[0]));
    const dateB = new Date(currentYear, parseInt(partsB[1]) - 1, parseInt(partsB[0]));
    return dateA - dateB;
  });

  const output = [];

  // Заголовки (строка 3)
  const headerRow = ["Артикул", "Наименование"];
  for (let i = 0; i < dateList.length; i++) {
    headerRow.push(dateList[i] + "\nРасходы", dateList[i] + "\nПродажи", dateList[i] + "\nРеклама");
  }
  headerRow.push("Итого\nРасходы", "Итого\nПродажи", "Итого\nРеклама");
  output.push(headerRow);

  // Считаем общие итоги по столбцам
  const colTotals = new Array(dateList.length * 3).fill(0);

  const keys = Object.keys(result).sort((a, b) =>
    a.split("|")[0].localeCompare(b.split("|")[0], "ru")
  );

  // Сначала собираем данные по товарам
  const dataRows = [];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const parts = key.split("|");
    const rowObj = result[key];
    const row = [parts[0], parts[1]];
    let totalAds = 0;
    let totalSales = 0;

    for (let j = 0; j < dateList.length; j++) {
      const d = dateList[j];
      const ads = rowObj[d]?.ads || 0;
      const sales = rowObj[d]?.sales || 0;
      // const adsRub = ads / 100;
      // const salesRub = sales / 100;
      // const percent = salesRub !== 0 ? (adsRub / salesRub) : 0;

      const adsRub = ads;
      const salesRub = sales;
      const percent = salesRub !== 0 ? (adsRub / salesRub) : 0;

      row.push(adsRub);
      row.push(salesRub);
      row.push(percent);

      totalAds += adsRub;
      totalSales += salesRub;
      colTotals[j * 3] += adsRub;
      colTotals[j * 3 + 1] += salesRub;
    }

    const totalPercent = totalSales !== 0 ? (totalAds / totalSales) : 0;
    row.push(totalAds, totalSales, totalPercent);
    dataRows.push(row);
  }

  // 🔧 СТРОКА ИТОГО (вставляем ПЕРЕД данными)
  const totalRow = ["ИТОГО", ""];
  let grandTotalAds = 0;
  let grandTotalSales = 0;

  for (let j = 0; j < dateList.length; j++) {
    const totalAds = colTotals[j * 3];
    const totalSales = colTotals[j * 3 + 1];
    const totalPercent = totalSales !== 0 ? (totalAds / totalSales) : 0;
    totalRow.push(totalAds, totalSales, totalPercent);
    grandTotalAds += totalAds;
    grandTotalSales += totalSales;
  }

  const grandTotalPercent = grandTotalSales !== 0 ? (grandTotalAds / grandTotalSales) : 0;
  totalRow.push(grandTotalAds, grandTotalSales, grandTotalPercent);
  
  // Собираем финальный вывод: заголовок → итоги → данные
  output.push(totalRow);        // Итого сразу под заголовком
  for (const row of dataRows) { // Затем все товары
    output.push(row);
  }

  // Очистка и запись
  clearReportBody_(sheet);
  const outRange = sheet.getRange(3, 1, output.length, output[0].length);
  outRange.setValues(output);

  // Форматирование процентов (для всех строк)
  for (let j = 0; j < dateList.length; j++) {
    const percentCol = 3 + j * 3 + 2;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      sheet.getRange(4, percentCol, lastRow - 3, 1).setNumberFormat("0.00%");
    }
  }

  // Форматирование процентов в итоговых колонках
  const lastCol = output[0].length;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    sheet.getRange(4, lastCol, lastRow - 3, 1).setNumberFormat("0.00%");
  }

  // Форматирование расходов и продаж
  for (let j = 0; j < dateList.length; j++) {
    const adsCol = 3 + j * 3;
    const salesCol = 3 + j * 3 + 1;
    if (lastRow >= 4) {
      sheet.getRange(4, adsCol, lastRow - 3, 1).setNumberFormat("#,##0.00");
      sheet.getRange(4, salesCol, lastRow - 3, 1).setNumberFormat("#,##0.00");
    }
  }

  // Форматирование итоговых колонок (расходы, продажи)
  if (lastRow >= 4) {
    sheet.getRange(4, lastCol - 2, lastRow - 3, 1).setNumberFormat("#,##0.00");
    sheet.getRange(4, lastCol - 1, lastRow - 3, 1).setNumberFormat("#,##0.00");
  }

  // Стиль
  outRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(3, 1, output.length, 2).setHorizontalAlignment("left");

  const headerRange = sheet.getRange(3, 1, 1, output[0].length);
  headerRange.setFontWeight("bold").setBackground("#f1f1f1").setHorizontalAlignment("center").setWrap(true);
  sheet.setRowHeight(3, 45);

  if (sheet.getFilter()) sheet.getFilter().remove();
  headerRange.createFilter();

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 300);
  for (let col = 3; col <= output[0].length; col++) {
    sheet.setColumnWidth(col, 80);
  }

  // Итоговая строка (теперь это строка 4) — жирный шрифт и фон
  sheet.getRange(4, 1, 1, output[0].length)
    .setFontWeight("bold").setBackground("#e8e8e8");

  formatCurrentSheet();

  SpreadsheetApp.getUi().alert("✅ Отчёт 'Реклама' обновлён");
}

// Функция преобразования номера колонки в букву
function columnToLetter(column) {
  let result = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    column = Math.floor((column - 1) / 26);
  }
  return result;
}

function clearReportBody_(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;
  try { sheet.getRange(1, 3, 2, lastCol).breakApart(); } catch (e) {}
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(3, 1, lastRow - 2, lastCol).clear();
}

function trimReportExtraColumns_(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  var lastCol = sheet.getLastColumn();
  var maxCols = sheet.getMaxColumns();
  if (lastCol < 1) lastCol = 1;
  var used = 2;
  var headers = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] !== '' && headers[i] != null) used = i + 1;
  }
  if (maxCols <= used) return;
  try { sheet.getRange(1, 1, Math.min(2, lastRow), maxCols).breakApart(); } catch (e1) {}
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.deleteColumns(used + 1, maxCols - used);
}

// Основная функция форматирования
function formatCurrentSheet(optSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = optSheet || ss.getActiveSheet();
  const sheetName = sheet.getName();

  trimReportExtraColumns_(sheet);
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 5) {
    SpreadsheetApp.getUi().alert("⚠️ Недостаточно данных на листе " + sheetName);
    return;
  }
  
  const sheetsWithDetails = [
    'К-во заказов', 'Вес заказов', 
    'Вес продаж', 'К-во продаж', // 'Продажи', 'Сумма заказов', 
    // 'Реклама', 'Реклама'
    'Реклама',
    'К-во Заказов FBS',
    'Вес заказов FBS'
  ];
  const isFbsDetails = sheetName.indexOf('Заказов FBS') >= 0 || sheetName.indexOf('заказов FBS') >= 0;
  const isDetailsSheet = sheetsWithDetails.includes(sheetName) || isFbsDetails;
  
  const dataStartRow = 5;
  const dataStartCol = 3;
  
  const headersRow = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  
  // Находим колонки с датами
  const dateColumns = [];
  for (let col = dataStartCol; col <= lastCol; col++) {
    const headerValue = headersRow[col - 1];
    if (isReportDateHeader_(headerValue)) {
      dateColumns.push(col);
    }
  }
  
  if (dateColumns.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ Не найдены колонки с датами в строке 3");
    return;
  }
  
  // ============ 0. ОЧИСТКА СТРОК 1-2 ============
  if (lastCol >= 3) {
    try { sheet.getRange(1, 3, 2, lastCol - 2).breakApart(); } catch (e) {}
    sheet.getRange(1, 3, 1, lastCol - 2).clear();
    sheet.getRange(2, 3, 1, lastCol - 2).clear();
  }
  
  // ============ 1. ОБЪЕДИНЁННАЯ ЯЧЕЙКА НАД ДАТАМИ ============
  if (isDetailsSheet) {
    const firstDateCol = dateColumns[0];
    const lastDateCol = dateColumns[dateColumns.length - 1];
    
    const mergeRange = sheet.getRange(1, firstDateCol, 1, lastDateCol - firstDateCol + 1);
    mergeRange.merge();
    const detailsTitle = isFbsDetails
      ? 'Детализация заказов FBS по дням'
      : 'Детализация заказов по дням';
    sheet.getRange(1, firstDateCol).setValue(detailsTitle);
    sheet.getRange(1, firstDateCol).setFontWeight("bold")
      .setBackground("#0ee247")
      .setFontColor("#000000")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontSize(10);
    
    for (let row = 2; row <= 4; row++) {
      for (let i = 0; i < dateColumns.length; i++) {
        sheet.getRange(row, dateColumns[i]).setBackground("#0ee247");
      }
    }
  }
  
  // ============ 2. ФОРМАТИРОВАНИЕ ЧИСЕЛ ============
  let type = '';
  if (sheetName === 'К-во продаж' || sheetName === 'К-во заказов' || sheetName.indexOf('К-во Заказов FBS') >= 0) {
    type = 'quantity';
  } else if (sheetName === 'Вес продаж' || sheetName === 'Вес заказов' || sheetName.indexOf('Вес заказов FBS') >= 0) {
    type = 'weight';
  }
  
  if (type !== '') {
    for (let i = 0; i < dateColumns.length; i++) {
      const col = dateColumns[i];
      if (lastRow >= dataStartRow) {
        const dataRange = sheet.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1);
        if (type === 'weight') {
          dataRange.setNumberFormat("#,##0.000");
        } else if (type === 'quantity') {
          dataRange.setNumberFormat("#,##0");
        }
      }
    }
  }
  
  // ============ 3. ЧЕРЕДОВАНИЕ СТРОК ============
  for (let row = dataStartRow; row <= lastRow; row++) {
    sheet.getRange(row, 1, 1, lastCol).setBackground(row % 2 === 0 ? "#f5f5f5" : "#ffffff");
  }
  
  // ============ 4. ЗАЛИВКА КОЛОНОК "ИТОГО" ============
  const итогоColumns = [];
  for (let col = 1; col <= lastCol; col++) {
    if (String(headersRow[col - 1]).indexOf('Итого') !== -1) {
      итогоColumns.push(col);
    }
  }
  
  const fillFromRow = isFbsDetails ? 3 : 1;
  for (const col of итогоColumns) {
    sheet.getRange(fillFromRow, col, lastRow - fillFromRow + 1, 1).setBackground("#0ee247");
    if (lastRow >= 4) sheet.getRange(4, col).setFontWeight("bold");
  }
  
  // ============ 5. ЗАЛИВКА СЛУЖЕБНЫХ КОЛОНОК ============
  let возвратыCol = -1, остаткиCol = -1, средняяЦенаCol = -1, текущаяЦенаCol = -1, ценаCol = -1;
  let проданоCol = -1, вПутиCol = -1, отмененоCol = -1;
  
  for (let col = 1; col <= lastCol; col++) {
    const h = String(headersRow[col - 1]);
    if (h === 'Возвраты') возвратыCol = col;
    else if (h === 'Остатки') остаткиCol = col;
    else if (h === 'Ср. цена') средняяЦенаCol = col;
    else if (h === 'Текущ цена') текущаяЦенаCol = col;
    else if (h === 'Цена') ценаCol = col;
    else if (h === 'Продано') проданоCol = col;
    else if (h === 'В пути') вПутиCol = col;
    else if (h === 'Отменено') отмененоCol = col;
  }
  
  const colsToFill = [возвратыCol, остаткиCol, средняяЦенаCol, текущаяЦенаCol, ценаCol];
  const fillColors = ["#ffc000", "#ccffcc", "#fff2cc", "#a4c2f4", "#d5a6bd"];
  
  for (let i = 0; i < colsToFill.length; i++) {
    const col = colsToFill[i];
    if (col !== -1 && col <= lastCol) {
      sheet.getRange(3, col, lastRow - 2, 1).setBackground(fillColors[i]);
      sheet.getRange(4, col).setFontWeight("bold");
    }
  }

  // Ср. цена / Текущ цена / Цена — всегда два знака, без округления до целых
  const priceFmtCols = [средняяЦенаCol, текущаяЦенаCol, ценаCol];
  for (let i = 0; i < priceFmtCols.length; i++) {
    const col = priceFmtCols[i];
    if (col !== -1 && col <= lastCol && lastRow >= 4) {
      sheet.getRange(4, col, lastRow - 3, 1).setNumberFormat("#,##0.00");
    }
  }

  const fbsStatusCols = [проданоCol, вПутиCol, отмененоCol];
  const fbsStatusColors = ['#ffff00', '#00b0f0', '#ffc000'];
  for (let i = 0; i < fbsStatusCols.length; i++) {
    const col = fbsStatusCols[i];
    if (col !== -1 && col <= lastCol && lastRow >= 3) {
      sheet.getRange(3, col, lastRow - 2, 1).setBackground(fbsStatusColors[i]);
      if (lastRow >= 4) sheet.getRange(4, col).setFontWeight('bold');
    }
  }
  
  // ============ 🔥 6. УСЛОВНОЕ ФОРМАТИРОВАНИЕ ============
  sheet.clearConditionalFormatRules();
  const newRules = [];
  
  if (текущаяЦенаCol !== -1 && ценаCol !== -1 && lastRow >= dataStartRow) {
    const curLet = columnToLetter(текущаяЦенаCol);
    const manLet = columnToLetter(ценаCol);
    const range = sheet.getRange(dataStartRow, текущаяЦенаCol, lastRow - dataStartRow + 1, 1);
    const formula = `=${curLet}${dataStartRow}<${manLet}${dataStartRow}`;
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground("#ff0000")
        .setFontColor("#000000")
        .setRanges([range])
        .build()
    );
  }

  if (type === 'quantity') {
    addQuantityDateCF_(newRules, sheet, dateColumns, dataStartRow, lastRow);
  } else if (type === 'weight') {
    addWeightDateCF_(newRules, sheet, dateColumns, dataStartRow, lastRow);
  }

  if (newRules.length > 0) {
    sheet.setConditionalFormatRules(newRules);
  }
  
  // ============ 8-10. ЗАГОЛОВКИ, ШРИФТ, ФИЛЬТР ============
  if (lastRow >= 3) {
    sheet.getRange(3, 1, 1, 2).setBackground("#f1f1f1").setFontWeight("bold");
  }
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, lastCol).setFontSize(10).setFontFamily("Calibri");
    sheet.getRange(3, 1, 1, lastCol).setFontSize(10).setFontFamily("Calibri");
  }
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(3, 1, lastRow - 2, lastCol).createFilter();
}

function addQuantityDateCF_(newRules, sheet, dateColumns, dataStartRow, lastRow) {
  for (var i = 0; i < dateColumns.length; i++) {
    var col = dateColumns[i];
    var colLetter = columnToLetter(col);
    var range = sheet.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=" + colLetter + dataStartRow + "=0").setBackground("#ff9999").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(" + colLetter + dataStartRow + ">=1; " + colLetter + dataStartRow + "<=5)").setBackground("#ffcc99").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(" + colLetter + dataStartRow + ">=6; " + colLetter + dataStartRow + "<=10)").setBackground("#ccffcc").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=" + colLetter + dataStartRow + ">=11").setBackground("#99cc99").setRanges([range]).build());
  }
}

function addWeightDateCF_(newRules, sheet, dateColumns, dataStartRow, lastRow) {
  for (var i = 0; i < dateColumns.length; i++) {
    var col = dateColumns[i];
    var colLetter = columnToLetter(col);
    var range = sheet.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=" + colLetter + dataStartRow + "=0").setBackground("#ff9999").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(" + colLetter + dataStartRow + ">0; " + colLetter + dataStartRow + "<=1)").setBackground("#ffcc99").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=AND(" + colLetter + dataStartRow + ">1; " + colLetter + dataStartRow + "<=6)").setBackground("#ccffcc").setRanges([range]).build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=" + colLetter + dataStartRow + ">6").setBackground("#99cc99").setRanges([range]).build());
  }
}

function applyFbsWeightDateCF_(sheet, numDates, lastRow) {
  var start = 5;
  if (!sheet || !numDates || lastRow < start) return;
  sheet.clearConditionalFormatRules();
  var cols = [];
  for (var c = 3; c <= 2 + numDates; c++) cols.push(c);
  var rules = [];
  addWeightDateCF_(rules, sheet, cols, start, lastRow);
  if (rules.length) sheet.setConditionalFormatRules(rules);
}

function applyFbsQuantityDateCF_(sheet, numDates, lastRow) {
  var start = 5;
  if (!sheet || !numDates || lastRow < start) return;
  sheet.clearConditionalFormatRules();
  var cols = [];
  for (var c = 3; c <= 2 + numDates; c++) cols.push(c);
  var rules = [];
  addQuantityDateCF_(rules, sheet, cols, start, lastRow);
  if (rules.length) sheet.setConditionalFormatRules(rules);
}

function isReportDateHeader_(headerValue) {
  if (headerValue instanceof Date) return true;
  var s = String(headerValue == null ? '' : headerValue).trim();
  if (!s) return false;
  if (s.match(/^\d{1,2}\.\d{1,2}/)) return true;
  if (s.match(/^\d{1,2}\.\s*[а-яёА-ЯЁa-zA-Z]{2,}/)) return true;
  return false;
}

function applyFbsDetailsHeader_(sheet, numDates) {
  if (!sheet || !numDates) return;
  var first = 3;
  try { sheet.getRange(1, first, 2, numDates).breakApart(); } catch (e) {}
  sheet.getRange(1, first, 1, numDates).merge();
  sheet.getRange(1, first)
    .setValue('Детализация заказов FBS по дням')
    .setFontWeight('bold')
    .setBackground('#0ee247')
    .setFontColor('#000000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10);
  sheet.getRange(2, first, 3, numDates).setBackground('#0ee247');
  var trail = Math.max(1, sheet.getLastColumn() - (first + numDates) + 1);
  sheet.getRange(1, first + numDates, 2, trail).setBackground(null);
}