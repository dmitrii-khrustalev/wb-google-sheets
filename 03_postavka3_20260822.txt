// ============================================
// ОТЧЁТ "Поставка_new"
// Три блока:
//   1) Итого = Склад WB + Склад продавца (8 колонок, 🚚 только WB)
//   2) Склад WB — FBO (8 колонок)
//   3) Склад продавца — FBS (7 колонок, без 🚚)
// Хелперы: 03_postavka_20260706.js
// ============================================

const POSTAVKA3_SHEET = "Поставка_new";
const POSTAVKA3_WH_START = 11;  // K — блок «Склад WB»
const POSTAVKA3_FBS_START = 19; // S — блок «Склад продавца»
const POSTAVKA3_WH_WIDTH = 8;
const POSTAVKA3_FBS_WIDTH = 7;  // без колонки 🚚
const POSTAVKA3_WH_COLOR = "#76d7c4";
const POSTAVKA3_FBS_COLOR = "#fce5cd";

function сформироватьОтчетПоставка3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportSheet = ss.getSheetByName(POSTAVKA3_SHEET);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(POSTAVKA3_SHEET);
    const src = ss.getSheetByName("Поставка") || ss.getSheetByName("Поставка_2");
    if (src) reportSheet.getRange(1, 1, 5, 1).setValues(src.getRange(1, 1, 5, 1).getValues());
  }

  reportSheet.setConditionalFormatRules([]);

  const abVals = reportSheet.getRange(1, 1, 5, 2).getValues();
  const financeDateFrom = new Date(abVals[0][0]);
  const financeDateTo = new Date(abVals[1][0]);
  const supplyDateFrom = new Date(abVals[2][0]);
  const supplyDateTo = new Date(abVals[3][0]);
  const daysCount = parseFloat(abVals[4][0]) || 30;
  if (isNaN(financeDateFrom.getTime()) || isNaN(financeDateTo.getTime()) || isNaN(supplyDateFrom.getTime()) || isNaN(supplyDateTo.getTime())) {
    SpreadsheetApp.getUi().alert('❌ Введите корректные даты в ячейки A1-A4 на листе "Поставка_new"');
    return;
  }
  const dateDiff = Math.ceil((financeDateTo - financeDateFrom) / (1000 * 60 * 60 * 24)) + 1;

  const ordersData = getOrdersData(financeDateFrom, financeDateTo);
  const ordersWbOnly = getOrdersByWarehouseType3_(financeDateFrom, financeDateTo, 'Склад WB');
  const stockData = getStockData();
  const supplyDataByRegion = getSupplyDataByRegionWarehouse(supplyDateFrom, supplyDateTo);
  const goodsNames = loadGoodsNames();
  const fbsWeightByArt = getFbsWeightByArt3_();
  const fbsStock = getFbsStockData3_(fbsWeightByArt);
  const fbsOrders = getFbsOrdersData3_(financeDateFrom, financeDateTo);

  const inTransitByArt = {};
  for (const k in supplyDataByRegion) {
    const s = supplyDataByRegion[k];
    const artN = normalizeArt(s.art);
    if (!inTransitByArt[artN]) inTransitByArt[artN] = { qty: 0, weight: 0 };
    inTransitByArt[artN].qty += s.inTransitQty;
    inTransitByArt[artN].weight += s.inTransitWeight;
  }

  const WAREHOUSE_METRICS = ['Остатки, шт', 'Остатки, кг', 'Заказы, шт', '🚚', 'Ср. продажи, шт', 'Запас в днях', 'К-во на поставку', '✏️'];
  const FBS_METRICS = ['Остатки, шт', 'Остатки, кг', 'Заказы, шт', 'Ср. продажи, шт', 'Запас в днях', 'К-во на поставку', '✏️'];
  const WAREHOUSE_COL_WIDTHS = [50, 50, 45, 45, 45, 45, 45, 45];
  const FBS_COL_WIDTHS = [50, 50, 45, 45, 45, 45, 45];

  const metricsRow = ["Артикул продавца", "Наименование товара", ...WAREHOUSE_METRICS, ...WAREHOUSE_METRICS, ...FBS_METRICS];
  const usedCols = metricsRow.length;

  const oldLastRow = Math.max(reportSheet.getLastRow(), 7);
  const oldLastCol = Math.max(reportSheet.getLastColumn(), usedCols);

  if (oldLastCol > 2) reportSheet.getRange(4, 3, 3, oldLastCol - 2).clear();
  const dataClearRows = oldLastRow - 6;
  if (dataClearRows > 0 && oldLastCol > 0) {
    reportSheet.getRange(7, 1, dataClearRows, oldLastCol).clear();
  }
  reportSheet.getRange(1, 1, 5, 2).setValues(abVals);

  reportSheet.getRange(6, 1, 1, usedCols).setValues([metricsRow]);
  reportSheet.getRange(6, 1, 1, usedCols)
    .setFontFamily("Calibri").setFontSize(9).setFontWeight("bold")
    .setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");
  for (let col = 1; col <= usedCols; col++) {
    if (metricsRow[col - 1] === '🚚') reportSheet.getRange(6, col).setFontSize(15);
    if (metricsRow[col - 1] === '✏️') reportSheet.getRange(6, col).setFontSize(13);
  }
  reportSheet.setRowHeight(4, 30);
  reportSheet.setRowHeight(5, 30);
  reportSheet.setRowHeight(6, 45);
  reportSheet.setColumnWidth(1, 150);
  reportSheet.setColumnWidth(2, 350);
  for (let i = 0; i < WAREHOUSE_METRICS.length; i++) {
    reportSheet.setColumnWidth(3 + i, WAREHOUSE_COL_WIDTHS[i]);
    reportSheet.setColumnWidth(POSTAVKA3_WH_START + i, WAREHOUSE_COL_WIDTHS[i]);
  }
  for (let i = 0; i < FBS_METRICS.length; i++) {
    reportSheet.setColumnWidth(POSTAVKA3_FBS_START + i, FBS_COL_WIDTHS[i]);
  }

  const allArts = new Map();
  function addArt(art, orig) {
    if (!art) return;
    const n = normalizeArt(art);
    if (!allArts.has(n)) allArts.set(n, orig || art);
  }
  for (const k in ordersData.byArt) addArt(k);
  for (const k in ordersWbOnly) addArt(k);
  for (const k in stockData.byArt) addArt(k);
  for (const k in supplyDataByRegion) addArt(supplyDataByRegion[k].art);
  for (const k in fbsStock) addArt(k);
  for (const k in fbsOrders) addArt(k);

  const arts = Array.from(allArts.keys()).sort();
  const maxRows = 200;
  const displayArts = arts.length > maxRows ? arts.slice(0, maxRows) : arts;
  if (arts.length > maxRows) SpreadsheetApp.getUi().alert('⚠️ Найдено ' + arts.length + ' товаров. Выводим только ' + maxRows);

  const dataRows = [];
  const totalRow = new Array(usedCols).fill(0);
  totalRow[0] = "ИТОГО";
  totalRow[1] = "";

  for (const artNorm of displayArts) {
    const origArt = allArts.get(artNorm);
    const productName = goodsNames[artNorm] || '';
    const row = [origArt, productName];

    const inT = inTransitByArt[artNorm] || { qty: 0, weight: 0 };
    const totalStock = stockData.byArt[artNorm] || { stockQty: 0, stockWeight: 0 };

    const wbOrders = ordersWbOnly[artNorm] || 0;
    const wbAvg = dateDiff > 0 ? Math.round((wbOrders / dateDiff) * 100) / 100 : 0;
    const wbDays = wbAvg > 0
      ? Math.round(((totalStock.stockQty + inT.qty) / wbAvg) * 100) / 100
      : 0;
    let wbVol = Math.round((wbAvg * daysCount) - totalStock.stockQty - inT.qty);
    if (wbVol <= 0) wbVol = "";

    const fbsSt = fbsStock[artNorm] || { stockQty: 0, stockWeight: 0 };
    const fbsOrd = fbsOrders[artNorm] || 0;
    const fbsAvg = dateDiff > 0 ? Math.round((fbsOrd / dateDiff) * 100) / 100 : 0;
    const fbsDays = fbsAvg > 0 ? Math.round((fbsSt.stockQty / fbsAvg) * 100) / 100 : 0;
    let fbsVol = Math.round((fbsAvg * daysCount) - fbsSt.stockQty);
    if (fbsVol <= 0) fbsVol = "";

    const sumStockQty = totalStock.stockQty + fbsSt.stockQty;
    const sumStockKg = totalStock.stockWeight + fbsSt.stockWeight;
    const sumOrders = wbOrders + fbsOrd;
    const sumAvg = dateDiff > 0 ? Math.round((sumOrders / dateDiff) * 100) / 100 : 0;
    const sumDays = sumAvg > 0 ? Math.round(((sumStockQty + inT.qty) / sumAvg) * 100) / 100 : 0;
    let sumVol = Math.round((sumAvg * daysCount) - sumStockQty - inT.qty);
    if (sumVol <= 0) sumVol = "";

    row.push(sumStockQty, sumStockKg, sumOrders, inT.qty, sumAvg, sumDays, sumVol, "");
    row.push(totalStock.stockQty, totalStock.stockWeight, wbOrders, inT.qty, wbAvg, wbDays, wbVol, "");
    row.push(fbsSt.stockQty, fbsSt.stockWeight, fbsOrd, fbsAvg, fbsDays, fbsVol, "");

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

    const lastDataRow = 7 + dataRows.length;
    const LOW_STOCK_COLOR = "#ff9999";
    const rules = [];
    const stockDaysPairs = [
      ['C', 'H'],
      ['K', 'P'],
      ['S', 'W']
    ];
    for (const pair of stockDaysPairs) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + pair[0] + '8<=10')
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(pair[0] + '8:' + pair[0] + lastDataRow)])
        .build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + pair[1] + '8<=30')
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(pair[1] + '8:' + pair[1] + lastDataRow)])
        .build());
    }
    reportSheet.setConditionalFormatRules(rules);

    const pencilFormulas = [];
    for (let row = 8; row <= lastDataRow; row++) pencilFormulas.push([`=R${row}+Y${row}`]);
    reportSheet.getRange(8, 10, dataRows.length, 1).setFormulas(pencilFormulas);

    reportSheet.getRange(7, 8).setFormula('=IF(G7=0;"";(C7+F7)/G7)').setNumberFormat("#,##0.00");
    reportSheet.getRange(7, 9).setFormula(`=SUM(I8:I${lastDataRow})`).setNumberFormat("0");
    reportSheet.getRange(7, 10).setFormula(`=SUM(J8:J${lastDataRow})`).setNumberFormat("0");

    reportSheet.getRange(7, 16).setFormula('=IF(O7=0;"";(K7+N7)/O7)').setNumberFormat("#,##0.00");
    reportSheet.getRange(7, 17).setFormula(`=SUM(Q8:Q${lastDataRow})`).setNumberFormat("0");
    reportSheet.getRange(7, 18).setFormula(`=SUM(R8:R${lastDataRow})`).setNumberFormat("0");

    reportSheet.getRange(7, 23).setFormula('=IF(V7=0;"";S7/V7)').setNumberFormat("#,##0.00");
    reportSheet.getRange(7, 24).setFormula(`=SUM(X8:X${lastDataRow})`).setNumberFormat("0");
    reportSheet.getRange(7, 25).setFormula(`=SUM(Y8:Y${lastDataRow})`).setNumberFormat("0");
  }

  SpreadsheetApp.flush();
  if (reportSheet.getFilter()) reportSheet.getFilter().remove();
  const filterLastRow = dataRows.length > 0 ? (7 + dataRows.length) : 7;
  try { reportSheet.getRange(6, 1, filterLastRow - 5, usedCols).createFilter(); } catch (e) { }

  formatPostavkaSheetFinal3();
  paintPostavka3Blocks_(reportSheet);
  addPrintBlockDropdown3_(reportSheet);
  trimPostavkaExtraColumns_(reportSheet, usedCols);

  SpreadsheetApp.getUi().alert('✅ Отчёт "Поставка_new" обновлён!');
}

function formatPostavkaSheetFinal3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA3_SHEET);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 6) return;

  const headers = sheet.getRange(6, 1, 1, lastCol).getValues()[0];
  sheet.getRange(7, 1, 1, lastCol).setFontWeight("bold").setBackground("#e8e8e8");
  sheet.getRange(6, 1, 1, lastCol)
    .setFontFamily("Calibri").setFontSize(9).setFontWeight("bold")
    .setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");
  if (lastRow >= 8) {
    sheet.getRange(8, 1, lastRow - 7, lastCol).setFontSize(9).setFontFamily("Calibri");
  }

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

function addPrintBlockDropdown3_(sheet) {
  if (!sheet) return;
  const kept = String(sheet.getRange("B4").getValue() || "").trim();
  const list = ["Выберите блок", "Склад WB", "Склад продавца"];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  const b4 = sheet.getRange("B4");
  b4.setDataValidation(rule);
  b4.setValue(list.indexOf(kept) >= 0 && kept !== "Выберите блок" ? kept : "Выберите блок");
  if (typeof formatDropdownCell === "function") formatDropdownCell(b4);
  const b5 = sheet.getRange("B5");
  b5.clearDataValidations();
  if (String(b5.getValue() || "").indexOf("Выберите") >= 0) b5.clearContent();
  if (sheet.getColumnWidth(2) < 150) sheet.setColumnWidth(2, 180);
}

function paintPostavka3Blocks_(sheet) {
  const lastRow = sheet.getLastRow();
  paintPostavka3OneBlock_(sheet, 3, POSTAVKA3_WH_WIDTH, "Итого", "Склад WB + Склад продавца", "#cee2e3", lastRow);
  paintPostavka3OneBlock_(sheet, POSTAVKA3_WH_START, POSTAVKA3_WH_WIDTH, "Склад WB", "Итого по всем складам", POSTAVKA3_WH_COLOR, lastRow);
  paintPostavka3OneBlock_(sheet, POSTAVKA3_FBS_START, POSTAVKA3_FBS_WIDTH, "Склад продавца", "FBS", POSTAVKA3_FBS_COLOR, lastRow);
  if (lastRow >= 7) {
    sheet.getRange(7, 1, 1, sheet.getLastColumn()).setBackground("#e8e8e8").setFontWeight("bold");
  }
}

function paintPostavka3OneBlock_(sheet, startCol, width, title4, title5, color, lastRow) {
  sheet.getRange(4, startCol, 1, width)
    .merge()
    .setValue(title4)
    .setFontFamily("Calibri").setFontSize(9).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground(color);
  sheet.getRange(5, startCol, 1, width)
    .merge()
    .setValue(title5)
    .setFontFamily("Calibri").setFontSize(9).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground(color);
  if (startCol === 3) {
    sheet.getRange(6, startCol, 1, width).setBackground(color);
    if (lastRow >= 8) sheet.getRange(8, startCol, lastRow - 7, width).setBackground(color);
  } else if (lastRow >= 6) {
    sheet.getRange(6, startCol, lastRow - 5, width)
      .setBackground(color)
      .setBorder(true, true, true, true, false, false);
  }
}

function getOrdersByWarehouseType3_(dateFrom, dateTo, warehouseType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Заказы");
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let artIdx = postavka3Col_(headers, ['supplierArticle', 'sa_name']);
  let dateIdx = postavka3Col_(headers, ['Дата']);
  let typeIdx = postavka3Col_(headers, ['warehouseType', 'Тип склада']);
  if (artIdx < 0 || dateIdx < 0 || typeIdx < 0) return {};

  const want = String(warehouseType || '').trim().toLowerCase();
  const fromTime = dateFrom.getTime();
  const toTime = new Date(dateTo.getFullYear(), dateTo.getMonth(), dateTo.getDate(), 23, 59, 59).getTime();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const typeVal = String(data[i][typeIdx] || '').trim().toLowerCase();
    if (typeVal !== want) continue;
    const dt = parseDateStr(data[i][dateIdx]);
    if (!dt) continue;
    const t = dt.getTime();
    if (t < fromTime || t > toTime) continue;
    const art = normalizeArt(data[i][artIdx]);
    if (!art) continue;
    result[art] = (result[art] || 0) + 1;
  }
  return result;
}

function getFbsWeightByArt3_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Товар+Вес");
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let aIdx = postavka3Col_(headers, ['артикул', 'sa_name']);
  let wIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === 'Вес, кг за 1 шт') wIdx = i;
  }
  if (aIdx < 0 || wIdx < 0) return map;
  for (let i = 1; i < data.length; i++) {
    const art = normalizeArt(data[i][aIdx]);
    if (!art) continue;
    const w = parseFloat(String(data[i][wIdx]).replace(',', '.'));
    if (!isNaN(w)) map[art] = w;
  }
  return map;
}

function getFbsStockData3_(weightByArt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FBS Остатки");
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dIdx = postavka3Col_(headers, ['дата']);
  const aIdx = postavka3Col_(headers, ['vendorcode', 'артикул', 'article']);
  const qIdx = postavka3Col_(headers, ['amount', 'остаток', 'quantity']);
  const wSumIdx = postavka3Col_(headers, ['вес, кг за 1 шт сумма']);
  const wUnitIdx = postavka3Col_(headers, ['вес, кг за 1 шт']);
  if (dIdx < 0 || aIdx < 0 || qIdx < 0) return {};

  let latest = null;
  for (let i = 1; i < data.length; i++) {
    const dt = parseDateStr(data[i][dIdx]);
    if (dt && (!latest || dt > latest)) latest = dt;
  }
  if (!latest) return {};

  const latestTime = latest.getTime();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const dt = parseDateStr(data[i][dIdx]);
    if (!dt || dt.getTime() !== latestTime) continue;
    const art = normalizeArt(data[i][aIdx]);
    if (!art) continue;
    const qty = parseFloat(String(data[i][qIdx]).replace(',', '.')) || 0;
    let weight = 0;
    if (wSumIdx >= 0) weight = parseFloat(String(data[i][wSumIdx]).replace(',', '.')) || 0;
    else {
      const wUnit = (wUnitIdx >= 0 ? parseFloat(String(data[i][wUnitIdx]).replace(',', '.')) : 0) || (weightByArt[art] || 0);
      weight = qty * wUnit;
    }
    if (!result[art]) result[art] = { stockQty: 0, stockWeight: 0 };
    result[art].stockQty += qty;
    result[art].stockWeight += weight;
  }
  return result;
}

function getFbsOrdersData3_(dateFrom, dateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FBS Заказы");
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const aIdx = postavka3Col_(headers, ['article', 'артикул']);
  const dIdx = postavka3Col_(headers, ['дата']);
  const cIdx = postavka3Col_(headers, ['createdat']);
  if (aIdx < 0) return {};

  const fromTime = dateFrom.getTime();
  const toTime = new Date(dateTo.getFullYear(), dateTo.getMonth(), dateTo.getDate(), 23, 59, 59).getTime();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const dt = parseDateStr(dIdx >= 0 ? data[i][dIdx] : '') || parseDateStr(cIdx >= 0 ? data[i][cIdx] : '');
    if (!dt) continue;
    const t = dt.getTime();
    if (t < fromTime || t > toTime) continue;
    const art = normalizeArt(data[i][aIdx]);
    if (!art) continue;
    result[art] = (result[art] || 0) + 1;
  }
  return result;
}

function postavka3Col_(headers, names) {
  if (typeof wbColIdx_ === 'function') return wbColIdx_(headers, names);
  const lower = names.map(n => String(n).toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim().toLowerCase();
    if (lower.indexOf(h) >= 0) return i;
  }
  return -1;
}
