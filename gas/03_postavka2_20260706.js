// ============================================
// ОТЧЁТ "ПОСТАВКА_2" — копия "Поставка"
// Остатки по всем складам WB.
// Переиспользует константы и хелперы из 03_postavka_20260706.js
// ============================================

const POSTAVKA2_SHEET = "Поставка_2";

// ============================================
// ОСНОВНАЯ ФУНКЦИЯ ОТЧЁТА "ПОСТАВКА_2"
// ============================================
function сформироватьОтчетПоставка2() {
  const regionOrder = ["Склад WB", "Центральный", "Северо-Западный", "Приволжский", "Южный и Северо-Кавказский", "Уральский", "Дальневосточный и Сибирский", "Беларусь", "Казахстан", "Узбекистан", "Армения", "Грузия", "Ереван", "Не определён"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName(POSTAVKA2_SHEET);
  if (!reportSheet) { SpreadsheetApp.getUi().alert('❌ Лист "Поставка_2" не найден'); return; }

  reportSheet.setConditionalFormatRules([]);

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

  const oldLastRow = Math.max(reportSheet.getLastRow(), 7);
  const oldLastCol = Math.max(reportSheet.getLastColumn(), metricsRow.length);

  if (oldLastCol > 2) {
    reportSheet.getRange(4, 3, 3, oldLastCol - 2).clear();
  }
  const dataClearRows = oldLastRow - 6;
  if (dataClearRows > 0 && oldLastCol > 0) {
    reportSheet.getRange(7, 1, dataClearRows, oldLastCol).clear();
  }

  reportSheet.getRange(1, 1, 5, 2).setValues(abVals);

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

    const totalInWB = baseInTQty + totalStock.stockQty;
    row.push(baseOrders, baseInTQty, baseInTWeight, totalStock.stockQty, totalInWB);
    const avgSalesPerDay = dateDiff > 0 ? Math.round((baseOrders / dateDiff) * 100) / 100 : 0;
    row.push(avgSalesPerDay);
    const daysOfStock = avgSalesPerDay > 0 ? Math.round(((totalStock.stockQty + baseInTQty) / avgSalesPerDay) * 100) / 100 : 0;
    row.push(daysOfStock);
    let supplyVolume = Math.round((avgSalesPerDay * daysCount) - totalStock.stockQty - baseInTQty);
    row.push(supplyVolume <= 0 ? "" : supplyVolume);
    row.push("");

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

    applyPostavka2MetricFormats_(reportSheet, metricsRow, 8, dataRows.length);
  }

  if (dataRows.length > 0) {
    const lastDataRow = 7 + dataRows.length;
    const LOW_STOCK_COLOR = "#ff9999";
    const rules = [];

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=F8<=10`)
      .setBackground(LOW_STOCK_COLOR)
      .setRanges([reportSheet.getRange(`F8:F${lastDataRow}`)])
      .build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=I8<=30`)
      .setBackground(LOW_STOCK_COLOR)
      .setRanges([reportSheet.getRange(`I8:I${lastDataRow}`)])
      .build());

    let cursor = REGION_START_COL;
    for (const region of regions) {
      const warehouses = regionsMap.get(region);

      const rStockLet = getColumnLetter(cursor + 0);
      const rDaysLet = getColumnLetter(cursor + 5);

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${rStockLet}8<=10`)
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(`${rStockLet}8:${rStockLet}${lastDataRow}`)])
        .build());

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${rDaysLet}8<=30`)
        .setBackground(LOW_STOCK_COLOR)
        .setRanges([reportSheet.getRange(`${rDaysLet}8:${rDaysLet}${lastDataRow}`)])
        .build());

      cursor += WAREHOUSE_METRICS_COUNT;

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

    if (rules.length > 0) {
      reportSheet.setConditionalFormatRules(rules);
    }
  }

  if (dataRows.length > 0) {
    const pencilColBase = 11;
    const lastDataRow = 7 + dataRows.length;
    const numDataRows = lastDataRow - 7;
    const pencilFormulas = [];

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

  if (dataRows.length > 0) {
    const lastDataRow = 7 + dataRows.length;

    reportSheet.getRange(7, 9).setFormula('=IF(H7=0;"";(F7+D7)/H7)');
    reportSheet.getRange(7, 9).setNumberFormat("#,##0.00");

    reportSheet.getRange(7, 10).setFormula(`=SUM(J8:J${lastDataRow})`);
    reportSheet.getRange(7, 10).setNumberFormat("0");

    reportSheet.getRange(7, 11).setFormula(`=SUM(K8:K${lastDataRow})`);
    reportSheet.getRange(7, 11).setNumberFormat("0");

    let cursor = REGION_START_COL;
    for (const region of regions) {
      const warehouses = regionsMap.get(region);

      const rW = cursor + 4;
      const rD = cursor + 5;
      const rK = cursor + 6;
      const rX = cursor + 7;

      const rLetW = getColumnLetter(rW);

      reportSheet.getRange(7, rD).setFormula(`=IF(${rLetW}7=0;"";(${getColumnLetter(cursor)}7+${getColumnLetter(cursor + 3)}7)/${rLetW}7)`);
      reportSheet.getRange(7, rD).setNumberFormat("#,##0.00");

      reportSheet.getRange(7, rK).setFormula(`=SUM(${getColumnLetter(rK)}8:${getColumnLetter(rK)}${lastDataRow})`);
      reportSheet.getRange(7, rK).setNumberFormat("0");

      reportSheet.getRange(7, rX).setFormula(`=SUM(${getColumnLetter(rX)}8:${getColumnLetter(rX)}${lastDataRow})`);
      reportSheet.getRange(7, rX).setNumberFormat("0");

      cursor += WAREHOUSE_METRICS_COUNT;

      for (const wh of warehouses) {
        const wW = cursor + 4;
        const wD = cursor + 5;
        const wK = cursor + 6;
        const wX = cursor + 7;

        const wLetW = getColumnLetter(wW);

        reportSheet.getRange(7, wD).setFormula(`=IF(${wLetW}7=0;"";(${getColumnLetter(cursor)}7+${getColumnLetter(cursor + 3)}7)/${wLetW}7)`);
        reportSheet.getRange(7, wD).setNumberFormat("#,##0.00");

        reportSheet.getRange(7, wK).setFormula(`=SUM(${getColumnLetter(wK)}8:${getColumnLetter(wK)}${lastDataRow})`);
        reportSheet.getRange(7, wK).setNumberFormat("0");

        reportSheet.getRange(7, wX).setFormula(`=SUM(${getColumnLetter(wX)}8:${getColumnLetter(wX)}${lastDataRow})`);
        reportSheet.getRange(7, wX).setNumberFormat("0");

        cursor += WAREHOUSE_METRICS_COUNT;
      }
    }
  }

  calculateRequestForShipment2();
  addRegionDropdown2();
  formatPostavkaSheetFinal2();
  applyRegionColors2();

  const BASE_BLOCK_START = 3;
  const BASE_BLOCK_END = 11;
  const BASE_BG_COLOR = "#cee2e3";
  const BASE_WIDTH = BASE_BLOCK_END - BASE_BLOCK_START + 1;

  reportSheet.getRange(4, BASE_BLOCK_START, 1, BASE_WIDTH)
    .merge()
    .setValue("Все склады WB")
    .setFontFamily("Calibri")
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(BASE_BG_COLOR);

  reportSheet.getRange(5, BASE_BLOCK_START, 1, BASE_WIDTH)
    .merge()
    .setValue("Итого по всем регионам")
    .setFontFamily("Calibri")
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(BASE_BG_COLOR);

  reportSheet.getRange(6, BASE_BLOCK_START, 1, BASE_WIDTH).setBackground(BASE_BG_COLOR);

  const lastRow = reportSheet.getLastRow();
  if (lastRow >= 8) {
    reportSheet.getRange(8, BASE_BLOCK_START, lastRow - 7, BASE_WIDTH).setBackground(BASE_BG_COLOR);
  }

  trimPostavkaExtraColumns_(reportSheet, metricsRow.length);

  SpreadsheetApp.getUi().alert('✅ Отчёт "Поставка_2" обновлён!');
}

// ============================================
// ФОРМАТИРОВАНИЕ И РАСЧЁТЫ (Поставка_2)
// ============================================
function formatPostavkaSheetFinal2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 6) return;

  const headers = sheet.getRange(6, 1, 1, lastCol).getValues()[0];

  sheet.getRange(7, 1, 1, lastCol).setFontWeight("bold").setBackground("#e8e8e8");

  if (lastCol >= REGION_START_COL) {
    sheet.getRange(4, REGION_START_COL, 2, lastCol - REGION_START_COL + 1).setFontFamily("Calibri").setFontSize(9);
  }

  sheet.getRange(6, 1, 1, lastCol).setFontFamily("Calibri").setFontSize(9).setFontWeight("bold").setBackground("#f1f1f1").setWrap(true).setHorizontalAlignment("center");

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

function calculateRequestForShipment2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 8) return;
  const headerRow = sheet.getRange(5, 1, 1, lastCol).getValues()[0];
  const REQUEST_COL_OFFSET = 7;
  const numRows = lastRow - 7;

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

// ============================================
// ХЕЛПЕРЫ ФОРМАТОВ / ГРУПП (свои, без зависимости от 03_postavka)
// ============================================
function setNumberFormatsGrouped2_(sheet, startRow, numRows, cols, format) {
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

function applyPostavka2MetricFormats_(sheet, metricsRow, dataStartRow, numDataRows) {
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
    setNumberFormatsGrouped2_(sheet, dataStartRow, numDataRows, kg, "#,##0.000");
    setNumberFormatsGrouped2_(sheet, dataStartRow, numDataRows, avg, "#,##0.00");
    setNumberFormatsGrouped2_(sheet, dataStartRow, numDataRows, intFmt, "#,##0");
    setNumberFormatsGrouped2_(sheet, dataStartRow, numDataRows, qtyBold, "#,##0");
    for (let i = 0; i < qtyBold.length; i++) {
      sheet.getRange(dataStartRow, qtyBold[i], numDataRows, 1).setFontWeight('bold');
    }
  }
  setNumberFormatsGrouped2_(sheet, 7, 1, kg, "#,##0.000");
  setNumberFormatsGrouped2_(sheet, 7, 1, avg, "#,##0.00");
  setNumberFormatsGrouped2_(sheet, 7, 1, intFmt.concat(qtyBold), "#,##0");
}

function removeColumnGroupsInRange2_(sheet, fromCol, toCol) {
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
  }
}

// ============================================
// ГРУППИРОВКА И ЦВЕТА (Поставка_2)
// ============================================
function createWarehouseGroups2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < REGION_START_COL) return;

  const regionRow = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const subHeaderRow = sheet.getRange(5, 1, 1, lastCol).getValues()[0];

  removeColumnGroupsInRange2_(sheet, REGION_START_COL, lastCol);

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

  try { sheet.collapseAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= lastCol; col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && !g.isCollapsed()) g.collapse(); } catch (e2) { }
    }
  }
}

function expandAllWarehouseGroups2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  try { sheet.expandAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= sheet.getLastColumn(); col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && g.isCollapsed()) g.expand(); } catch (e2) { }
    }
  }
}
function collapseAllWarehouseGroups2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  try { sheet.collapseAllColumnGroups(); } catch (e) {
    for (let col = REGION_START_COL; col <= sheet.getLastColumn(); col++) {
      try { const g = sheet.getColumnGroup(col, 1); if (g && !g.isCollapsed()) g.collapse(); } catch (e2) { }
    }
  }
}
function removeAllGroups2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  removeColumnGroupsInRange2_(sheet, REGION_START_COL, sheet.getLastColumn());
}

function создатьструктуру2() {
  createWarehouseGroups2();
  try { SpreadsheetApp.getUi().alert('✅ Группировка Поставка_2 обновлена'); } catch (e) { }
}

function applyRegionColors2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA2_SHEET);
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
// ВЫПАДАЮЩИЕ СПИСКИ (Поставка_2)
// ============================================
function addRegionDropdown2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POSTAVKA2_SHEET);
  if (!sheet) return;
  const pairs = getUniqueRegionWarehousePairs();
  const regions = [...new Set(pairs.map(p => p.region))].sort();
  const regionRule = SpreadsheetApp.newDataValidation().requireValueInList(["Выберите регион", ...regions], true).setAllowInvalid(false).build();
  const b4 = sheet.getRange("B4"); b4.setDataValidation(regionRule); b4.setValue("Выберите регион"); formatDropdownCell(b4);
  const warehouseRule = SpreadsheetApp.newDataValidation().requireValueInList(["Выберите склад"], true).setAllowInvalid(false).build();
  const b5 = sheet.getRange("B5"); b5.setDataValidation(warehouseRule); b5.setValue("Выберите склад"); formatDropdownCell(b5);
  if (sheet.getColumnWidth(2) < 150) sheet.setColumnWidth(2, 180);
}

function getRegionWarehouseMap2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POSTAVKA2_SHEET);
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

/** Обработка B4 на листе Поставка_2 (вызывается из общего onEdit) */
function handlePostavka2Edit(e) {
  if (!e || !e.source || !e.range) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== POSTAVKA2_SHEET) return;
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
