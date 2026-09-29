// ============================================
// ПЕЧАТЬ / ВЫГРУЗКА С ЛИСТА "Поставка_2"
// Копия 04_print — функции с суффиксом 2, без конфликтов имён
// Переиспользует: loadGoodsData, loadWeightsData, loadBarcodesSimple, normalizeArt
// ============================================

const PRINT2_SOURCE_SHEET = "Поставка_2";

function downloadReportByWarehouse2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(PRINT2_SOURCE_SHEET);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "Поставка_2" не найден');
    return;
  }
  
  const selectedWarehouse = sourceSheet.getRange("B5").getValue();
  if (!selectedWarehouse || selectedWarehouse === "Выберите склад") {
    SpreadsheetApp.getUi().alert("❌ Пожалуйста, выберите склад в ячейке B5 на листе Поставка_2");
    return;
  }
  
  const goodsData = loadGoodsData();
  const weightsData = loadWeightsData();
  
  const lastCol = sourceSheet.getLastColumn();
  const regionRow = sourceSheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const subHeaderRow = sourceSheet.getRange(5, 1, 1, lastCol).getValues()[0];
  
  let requestCol = -1;
  let foundWarehouse = "";
  let foundRegion = "";
  let currentRegion = "";
  
  for (let col = 8; col <= lastCol; col++) {
    const region = regionRow[col - 1]?.toString().trim();
    const subHeader = subHeaderRow[col - 1]?.toString().trim();
    if (region) currentRegion = region;
    if (subHeader === "ИТОГО по региону") { col += 7; continue; }
    if (subHeader === selectedWarehouse) {
      requestCol = col + 7;
      foundWarehouse = subHeader;
      foundRegion = currentRegion || "Не указан";
      break;
    }
  }
  
  if (requestCol === -1) {
    SpreadsheetApp.getUi().alert(`❌ Для склада "${selectedWarehouse}" не найдена колонка "✏️"`);
    return;
  }
  
  const lastRow = sourceSheet.getLastRow();
  const orderData = [];
  if (lastRow >= 8) {
    const arts = sourceSheet.getRange(8, 1, lastRow - 7, 1).getValues();
    const vals = sourceSheet.getRange(8, requestCol, lastRow - 7, 1).getValues();
    for (let i = 0; i < arts.length; i++) {
      const art = arts[i][0];
      if (!art) continue;
      const value = vals[i][0];
      if (value && value !== 0 && value !== "" && !isNaN(parseFloat(value))) {
        orderData.push({ art: art.toString(), packs: parseFloat(value) });
      }
    }
  }
  
  if (orderData.length === 0) {
    SpreadsheetApp.getUi().alert(`❌ Нет товаров с заявкой по складу "${selectedWarehouse}"`);
    return;
  }
  
  const now = new Date();
  const dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') + ":" + String(now.getMinutes()).padStart(2, '0') + ":" + String(now.getSeconds()).padStart(2, '0');
  const fileName = `${selectedWarehouse}_${dateStr}_${timeStr}.xlsx`;
  
  const tempBook = SpreadsheetApp.create(`temp_${Date.now()}`);
  const tempSheet = tempBook.getActiveSheet();
  tempSheet.setName("УПАКОВКА");
  
  tempSheet.setRowHeight(1, 30);
  const currentDay = String(now.getDate()).padStart(2, '0');
  const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
  const currentYear = now.getFullYear();
  
  tempSheet.getRange(1, 1).setValue(`${currentDay}.${currentMonth}.${currentYear}`);
  tempSheet.getRange(1, 1).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold");
  tempSheet.getRange(1, 2).setValue("WB");
  tempSheet.getRange(1, 2).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold");
  tempSheet.getRange(1, 2).setHorizontalAlignment("center").setVerticalAlignment("middle");
  tempSheet.getRange(1, 3).setValue(selectedWarehouse);
  tempSheet.getRange(1, 3).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold");
  tempSheet.getRange(1, 3).setHorizontalAlignment("center").setVerticalAlignment("middle");
  tempSheet.setColumnWidth(3, 250);

  const widths = [190, 100, 550, 140, 100, 100, 100, 85, 190, 85, 85, 85, 100, 100];
  widths.forEach((w, i) => tempSheet.setColumnWidth(i + 1, w));
  
  const headers = ["Чертеж", "Вес 1 шт, кг", "Наименование товара", "Штрих код", "Объем коробки", "Объем упаковки", "Объем заказа", "№ коробки", "Артикул", "Кол-во, упаковок", "Кол-во в упак.", "Кол-во, шт.", "Общий вес, кг", "Факт. вес, кг"];
  tempSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  tempSheet.getRange(2, 1, 1, headers.length).setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold").setBackground("#ffff00").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true).setBorder(true, true, true, true, false, false);
  
  const BOX_VOL_LIMIT = 0.014;
  const BOX_WT_LIMIT = 23;
  
  const expandedData = [];
  let currentBoxNum = 1;
  let currentBoxVol = 0.0;
  let currentBoxWt = 0.0;

  for (const item of orderData) {
    const artRaw = item.art;
    const artKey = normalizeArt(artRaw);
    let packsLeft = item.packs;
    
    const wInfo = weightsData[artKey] || { weight: 0, volume: 0, piecesPerPackage: 1, drawing: "", barcode: "" };
    const gInfo = goodsData[artKey] || { name: "" };
    
    const piecesPerPackage = wInfo.piecesPerPackage || 1;
    
    const volPerPack = (parseFloat(wInfo.volume) || 0) * piecesPerPackage; 
    const wtPerPack = (parseFloat(wInfo.weight) || 0) * piecesPerPackage;

    while (packsLeft > 0) {
      let spaceVol = BOX_VOL_LIMIT - currentBoxVol;
      let spaceWt = BOX_WT_LIMIT - currentBoxWt;

      let maxByVol = volPerPack > 0 ? Math.floor((spaceVol + 1e-9) / volPerPack) : Infinity;
      let maxByWt = wtPerPack > 0 ? Math.floor((spaceWt + 1e-9) / wtPerPack) : Infinity;
      
      let fits = Math.min(maxByVol, maxByWt, packsLeft);

      if (fits > 0) {
        expandedData.push({
          drawing: wInfo.drawing || "",
          weight: wInfo.weight || 0,
          name: gInfo.name || "",
          barcode: wInfo.barcode || "",
          boxVolume: BOX_VOL_LIMIT,
          packVolume: wInfo.volume || 0,
          boxNumber: currentBoxNum,
          art: artRaw,
          packs: fits,
          piecesPerPackage: piecesPerPackage
        });
        currentBoxVol = Math.round((currentBoxVol + fits * volPerPack) * 1e6) / 1e6;
        currentBoxWt = Math.round((currentBoxWt + fits * wtPerPack) * 1e6) / 1e6;
        packsLeft -= fits;
      } else {
        if (currentBoxVol > 0 || currentBoxWt > 0) {
          currentBoxNum++;
          currentBoxVol = 0.0;
          currentBoxWt = 0.0;
        } else {
           expandedData.push({ drawing: wInfo.drawing||"", weight: wInfo.weight||0, name: gInfo.name||"", barcode: wInfo.barcode||"", boxVolume: BOX_VOL_LIMIT, packVolume: wInfo.volume||0, boxNumber: currentBoxNum, art: artRaw, packs: packsLeft, piecesPerPackage: piecesPerPackage });
           packsLeft = 0;
        }
      }
    }
  }
  
  const startRow = 3;
  const n = expandedData.length;
  const values = new Array(n);
  for (let i = 0; i < n; i++) {
    const d = expandedData[i];
    const row = startRow + i;
    values[i] = [
      d.drawing, d.weight, d.name, d.barcode, d.boxVolume, d.packVolume,
      `=F${row}*L${row}`, d.boxNumber, d.art, d.packs, d.piecesPerPackage,
      `=J${row}*K${row}`, `=B${row}*J${row}`, ""
    ];
  }
  const dataRange = tempSheet.getRange(startRow, 1, n, 14);
  dataRange.setValues(values);
  dataRange.setFontFamily("Times New Roman").setFontSize(12)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
  tempSheet.getRange(startRow, 3, n, 1).setHorizontalAlignment("left");
  tempSheet.getRange(startRow, 2, n, 1).setNumberFormat("0.000");
  tempSheet.getRange(startRow, 4, n, 1).setNumberFormat("0");
  tempSheet.getRange(startRow, 5, n, 1).setNumberFormat("0.000");
  tempSheet.getRange(startRow, 6, n, 1).setNumberFormat("0.000000");
  tempSheet.getRange(startRow, 7, n, 1).setNumberFormat("0.000000");
  tempSheet.getRange(startRow, 10, n, 1).setFontWeight("bold");
  tempSheet.getRange(startRow, 13, n, 1).setNumberFormat("0.000");
  tempSheet.getRange(startRow, 14, n, 1).setNumberFormat("0.000");
  
  const sumRow = startRow + n;
  tempSheet.getRange(sumRow, 7).setFormula(`=SUM(G${startRow}:G${sumRow - 1})`).setNumberFormat("0.000000");
  tempSheet.getRange(sumRow, 10).setFormula(`=SUM(J${startRow}:J${sumRow - 1})`).setNumberFormat("0");
  tempSheet.getRange(sumRow, 12).setFormula(`=SUM(L${startRow}:L${sumRow - 1})`).setNumberFormat("0");
  tempSheet.getRange(sumRow, 13).setFormula(`=SUM(M${startRow}:M${sumRow - 1})`).setNumberFormat("0.000");
  
  tempSheet.getRange(sumRow, 7, 1, 7).setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#00ff00").setBorder(true, true, true, true, false, false);

  const deliveryStartRow = sumRow + 2;
  tempSheet.getRange(deliveryStartRow, 2, 1, 3).merge();
  tempSheet.getRange(deliveryStartRow, 2).setValue("Отгрузка");
  tempSheet.getRange(deliveryStartRow, 2).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold").setBackground("#ffff00").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, false, false);
  
  const labels = ["Дата готовности:", "Дата отгрузки:", "Тайм слот:"];
  for (let i = 0; i < labels.length; i++) {
    const labelRow = deliveryStartRow + 1 + i;
    tempSheet.getRange(labelRow, 2, 1, 2).merge();
    tempSheet.getRange(labelRow, 2).setValue(labels[i]).setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle").setBorder(true, true, true, true, false, false);
    tempSheet.getRange(labelRow, 4).setValue("").setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold").setBorder(true, true, true, true, false, false);
    if (i < 2) tempSheet.getRange(labelRow, 4).setNumberFormat("dd.mm.yyyy");
  }
    
  tempSheet.getRange(1, 5, 1, 3).shiftColumnGroupDepth(1);
  tempSheet.getColumnGroup(5, 1).collapse();

  tempSheet.setFrozenRows(2);
  tempSheet.getRange(2, 1, 1, headers.length).createFilter();

  const maxBoxNum = Math.max(...expandedData.map(d => d.boxNumber));
  tempSheet.getRange(1, 9, 1, 6).merge();
  tempSheet.getRange(1, 9).setValue(`Коробов: ${maxBoxNum}`)
    .setFontFamily("Times New Roman")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
      
  const tempBookId = tempBook.getId();
  const tempFile = DriveApp.getFileById(tempBookId);
  tempFile.setName(fileName.replace('.xlsx', ''));
  
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempBookId}/export?format=xlsx`;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html><body style="font-family: Arial; text-align: center; padding: 20px;">
      <h3>✅ Файл готов к скачиванию (Поставка_2)</h3>
      <p><strong>${fileName}</strong></p>
      <p>📦 Строк (коробок): ${expandedData.length}</p>
      <p>🏭 Склад: ${foundWarehouse} (${foundRegion})</p>
      <br>
      <a href="${downloadUrl}" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;" 
         onclick="setTimeout(function(){ google.script.host.close(); }, 1000);">📥 Скачать Excel файл</a>
      <br><br>
      <button onclick="google.script.host.close()" style="padding: 5px 15px;">Закрыть</button>
      <br><br><small style="color: gray;">Файл будет удалён через ~30 сек после закрытия диалога</small>
    </body></html>
  `).setWidth(450).setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Скачивание отчета (Поставка_2)');
  Utilities.sleep(30000);
  try { tempFile.setTrashed(true); } catch(e) {}
}

/**
 * Простая выгрузка с Поставка_2: Штрих-код + Количество
 */
function downloadSimpleBarcodeList2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(PRINT2_SOURCE_SHEET);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "Поставка_2" не найден');
    return;
  }
  
  const selectedWarehouse = sourceSheet.getRange("B5").getValue();
  if (!selectedWarehouse || selectedWarehouse === "Выберите склад") {
    SpreadsheetApp.getUi().alert("❌ Пожалуйста, выберите склад в ячейке B5 на листе Поставка_2");
    return;
  }
  
  const barcodes = loadBarcodesSimple();
  
  const lastCol = sourceSheet.getLastColumn();
  const subHeaderRow = sourceSheet.getRange(5, 1, 1, lastCol).getValues()[0];
  
  let pencilCol = -1;
  
  for (let col = 8; col <= lastCol; col++) {
    const subHeader = subHeaderRow[col - 1]?.toString().trim();
    
    if (subHeader === "ИТОГО по региону") { 
      col += 7; 
      continue; 
    }
    
    if (subHeader === selectedWarehouse) {
      pencilCol = col + 7;
      break;
    }
  }
  
  if (pencilCol === -1) {
    SpreadsheetApp.getUi().alert(`❌ Для склада "${selectedWarehouse}" не найдена колонка "✏️"`);
    return;
  }
  
  const lastRow = sourceSheet.getLastRow();
  const exportData = [];
  
  if (lastRow >= 8) {
    const arts = sourceSheet.getRange(8, 1, lastRow - 7, 1).getValues();
    const qtys = sourceSheet.getRange(8, pencilCol, lastRow - 7, 1).getValues();
    for (let i = 0; i < arts.length; i++) {
      const art = arts[i][0];
      if (!art) continue;
      const quantity = qtys[i][0];
      if (quantity && quantity !== 0 && quantity !== "" && !isNaN(parseFloat(quantity))) {
        const barcode = barcodes[normalizeArt(art)];
        if (barcode) exportData.push([barcode, Math.round(parseFloat(quantity))]);
      }
    }
  }
  
  if (exportData.length === 0) {
    SpreadsheetApp.getUi().alert(`❌ Нет товаров с заявкой по складу "${selectedWarehouse}"`);
    return;
  }
  
  const now = new Date();
  const dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') + ":" + String(now.getMinutes()).padStart(2, '0') + ":" + String(now.getSeconds()).padStart(2, '0');
  const fileName = `${selectedWarehouse}_${dateStr}_${timeStr}.xlsx`;
  
  const tempBook = SpreadsheetApp.create(`temp_${Date.now()}`);
  const tempSheet = tempBook.getActiveSheet();
  
  tempSheet.getRange(1, 1).setValue("Баркод");
  tempSheet.getRange(1, 2).setValue("Количество");
  
  if (exportData.length > 0) {
    tempSheet.getRange(2, 1, exportData.length, 2).setValues(exportData);
  }
  
  const tempBookId = tempBook.getId();
  const tempFile = DriveApp.getFileById(tempBookId);
  tempFile.setName(fileName.replace('.xlsx', ''));
  
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  
  const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempBookId}/export?format=xlsx`;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html><body style="font-family: Arial; text-align: center; padding: 20px;">
      <h3>✅ Файл готов (Поставка_2)</h3>
      <p><strong>${fileName}</strong></p>
      <p>📦 Товаров: ${exportData.length}</p>
      <br>
      <a href="${downloadUrl}" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">📥 Скачать</a>
      <br><br>
      <button onclick="google.script.host.close()">Закрыть</button>
    </body></html>
  `).setWidth(400).setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Скачивание (Поставка_2)');
  
  Utilities.sleep(30000);
  try { tempFile.setTrashed(true); } catch(e) {}
}
