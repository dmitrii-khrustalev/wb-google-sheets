// ============================================
// ПЕЧАТЬ / ВЫГРУЗКА С ЛИСТА "Поставка_new"
// Выбор блока в B4: «Склад WB» или «Склад продавца».
// Итоговый блок не печатается.
// Переиспользует: loadGoodsData, loadWeightsData, loadBarcodesSimple, normalizeArt
// ============================================

const PRINT3_SOURCE_SHEET = "Поставка_new";
const PRINT3_BLOCKS = ["Склад WB", "Склад продавца"];

function downloadReportByWarehouse3() {
  try {
    downloadReportByWarehouse3_();
  } catch (e) {
    Logger.log('downloadReportByWarehouse3: ' + e);
    try { SpreadsheetApp.getUi().alert('❌ Печать Поставка_new: ' + e); } catch (e2) {}
  }
}

function downloadReportByWarehouse3_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(PRINT3_SOURCE_SHEET);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "Поставка_new" не найден');
    return;
  }

  const sel = print3GetBlock_(sourceSheet);
  if (!sel) return;

  const goodsData = loadGoodsData();
  const weightsData = loadWeightsData();
  const orderData = print3CollectOrders_(sourceSheet, sel.pencilCol);
  if (orderData.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Нет товаров с заявкой по блоку "' + sel.name + '"');
    return;
  }

  const now = new Date();
  const dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') + ":" + String(now.getMinutes()).padStart(2, '0') + ":" + String(now.getSeconds()).padStart(2, '0');
  const fileName = `${sel.name}_${dateStr}_${timeStr}.xlsx`;

  const tempBook = SpreadsheetApp.create(`temp_${Date.now()}`);
  const tempSheet = tempBook.getActiveSheet();
  tempSheet.setName("УПАКОВКА");

  tempSheet.setRowHeight(1, 30);
  const currentDay = String(now.getDate()).padStart(2, '0');
  const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
  const currentYear = now.getFullYear();

  tempSheet.getRange(1, 1).setValue(`${currentDay}.${currentMonth}.${currentYear}`);
  tempSheet.getRange(1, 1).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold");
  tempSheet.getRange(1, 2).setValue(sel.name === "Склад продавца" ? "FBS" : "WB");
  tempSheet.getRange(1, 2).setFontFamily("Times New Roman").setFontSize(18).setFontWeight("bold");
  tempSheet.getRange(1, 2).setHorizontalAlignment("center").setVerticalAlignment("middle");
  tempSheet.getRange(1, 3).setValue(sel.name);
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
      const spaceVol = BOX_VOL_LIMIT - currentBoxVol;
      const spaceWt = BOX_WT_LIMIT - currentBoxWt;
      const maxByVol = volPerPack > 0 ? Math.floor((spaceVol + 1e-9) / volPerPack) : Infinity;
      const maxByWt = wtPerPack > 0 ? Math.floor((spaceWt + 1e-9) / wtPerPack) : Infinity;
      const fits = Math.min(maxByVol, maxByWt, packsLeft);

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
          expandedData.push({
            drawing: wInfo.drawing || "",
            weight: wInfo.weight || 0,
            name: gInfo.name || "",
            barcode: wInfo.barcode || "",
            boxVolume: BOX_VOL_LIMIT,
            packVolume: wInfo.volume || 0,
            boxNumber: currentBoxNum,
            art: artRaw,
            packs: packsLeft,
            piecesPerPackage: piecesPerPackage
          });
          packsLeft = 0;
        }
      }
    }
  }

  const startRow = 3;
  const n = expandedData.length;
  if (n < 1) {
    SpreadsheetApp.getUi().alert('❌ Не удалось разложить заявку по коробкам');
    try { DriveApp.getFileById(tempBook.getId()).setTrashed(true); } catch (e0) {}
    return;
  }
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

  try {
    tempSheet.getRange(1, 5, 1, 3).shiftColumnGroupDepth(1);
    tempSheet.getColumnGroup(5, 1).collapse();
  } catch (eG) {}
  tempSheet.setFrozenRows(2);
  try { tempSheet.getRange(2, 1, 1, headers.length).createFilter(); } catch (eF) {}

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
      <h3>✅ Файл готов к скачиванию (Поставка_new)</h3>
      <p><strong>${fileName}</strong></p>
      <p>📦 Строк (коробок): ${expandedData.length}</p>
      <p>🏭 Блок: ${sel.name}</p>
      <br>
      <a href="${downloadUrl}" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;"
         onclick="setTimeout(function(){ google.script.host.close(); }, 1000);">📥 Скачать Excel файл</a>
      <br><br>
      <button onclick="google.script.host.close()" style="padding: 5px 15px;">Закрыть</button>
      <br><br><small style="color: gray;">Файл будет удалён через ~30 сек после закрытия диалога</small>
    </body></html>
  `).setWidth(450).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Скачивание отчета (Поставка_new)');
  Utilities.sleep(30000);
  try { tempFile.setTrashed(true); } catch (e) {}
}

function downloadSimpleBarcodeList3() {
  try {
    downloadSimpleBarcodeList3_();
  } catch (e) {
    Logger.log('downloadSimpleBarcodeList3: ' + e);
    try { SpreadsheetApp.getUi().alert('❌ Печать штрихкодов Поставка_new: ' + e); } catch (e2) {}
  }
}

function downloadSimpleBarcodeList3_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(PRINT3_SOURCE_SHEET);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "Поставка_new" не найден');
    return;
  }

  const sel = print3GetBlock_(sourceSheet);
  if (!sel) return;

  const barcodes = loadBarcodesSimple();
  const lastRow = sourceSheet.getLastRow();
  const exportData = [];
  let withQty = 0;
  let noBarcode = 0;
  if (lastRow >= 8) {
    const arts = sourceSheet.getRange(8, 1, lastRow - 7, 1).getValues();
    const qtys = sourceSheet.getRange(8, sel.pencilCol, lastRow - 7, 1).getValues();
    for (let i = 0; i < arts.length; i++) {
      const art = arts[i][0];
      if (!art || print3IsTotalRow_(art)) continue;
      const quantity = qtys[i][0];
      if (quantity && quantity !== 0 && quantity !== "" && !isNaN(parseFloat(quantity))) {
        withQty++;
        const barcode = barcodes[normalizeArt(art)];
        if (barcode) exportData.push([barcode, Math.round(parseFloat(quantity))]);
        else noBarcode++;
      }
    }
  }

  if (exportData.length === 0) {
    if (withQty > 0) {
      SpreadsheetApp.getUi().alert('❌ Заявка есть (' + withQty + '), но штрихкоды на «Товар+Вес» не найдены');
    } else {
      SpreadsheetApp.getUi().alert('❌ Нет товаров с заявкой по блоку "' + sel.name + '"');
    }
    return;
  }

  const now = new Date();
  const dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') + ":" + String(now.getMinutes()).padStart(2, '0') + ":" + String(now.getSeconds()).padStart(2, '0');
  const fileName = `${sel.name}_${dateStr}_${timeStr}.xlsx`;

  const tempBook = SpreadsheetApp.create(`temp_${Date.now()}`);
  const tempSheet = tempBook.getActiveSheet();
  tempSheet.getRange(1, 1).setValue("Баркод");
  tempSheet.getRange(1, 2).setValue("Количество");
  tempSheet.getRange(2, 1, exportData.length, 2).setValues(exportData);

  const tempBookId = tempBook.getId();
  const tempFile = DriveApp.getFileById(tempBookId);
  tempFile.setName(fileName.replace('.xlsx', ''));
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempBookId}/export?format=xlsx`;

  const htmlOutput = HtmlService.createHtmlOutput(`
    <html><body style="font-family: Arial; text-align: center; padding: 20px;">
      <h3>✅ Файл готов (Поставка_new)</h3>
      <p><strong>${fileName}</strong></p>
      <p>📦 Товаров: ${exportData.length}</p>
      <p>🏭 Блок: ${sel.name}</p>
      <br>
      <a href="${downloadUrl}" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">📥 Скачать</a>
      <br><br>
      <button onclick="google.script.host.close()">Закрыть</button>
    </body></html>
  `).setWidth(400).setHeight(280);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Скачивание (Поставка_new)');
  Utilities.sleep(30000);
  try { tempFile.setTrashed(true); } catch (e) {}
}

function print3GetBlock_(sheet) {
  const raw = String(sheet.getRange("B4").getValue() || "").replace(/\s+/g, " ").trim();
  if (!raw || raw === "Выберите блок" || PRINT3_BLOCKS.indexOf(raw) < 0) {
    SpreadsheetApp.getUi().alert("❌ Выберите блок в ячейке B4: «Склад WB» или «Склад продавца»");
    return null;
  }
  const pencilCol = print3FindPencilCol_(sheet, raw);
  if (pencilCol < 1) {
    SpreadsheetApp.getUi().alert('❌ Не найдена колонка ✏️ для блока "' + raw + '"');
    return null;
  }
  return { name: raw, pencilCol: pencilCol };
}

function print3IsPencil_(v) {
  return String(v || "").indexOf("✏") >= 0;
}

function print3IsTotalRow_(art) {
  return String(art || "").trim().toUpperCase() === "ИТОГО";
}

function print3FindPencilCol_(sheet, blockName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return print3PencilFallback_(blockName);
  const row4 = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const row6 = sheet.getRange(6, 1, 1, lastCol).getValues()[0];
  let start = -1;
  for (let c = 0; c < row4.length; c++) {
    if (String(row4[c] || "").trim() === blockName) {
      start = c;
      break;
    }
  }
  if (start < 0) return print3PencilFallback_(blockName);
  let found = -1;
  for (let c = start; c < row6.length; c++) {
    const title = String(row4[c] || "").trim();
    if (c > start && title && title !== blockName) break;
    if (print3IsPencil_(row6[c])) found = c + 1;
  }
  if (found > 0) return found;
  return print3PencilFallback_(blockName);
}

function print3PencilFallback_(blockName) {
  if (typeof POSTAVKA3_WH_START === "number" && blockName === "Склад WB") {
    return POSTAVKA3_WH_START + POSTAVKA3_WH_WIDTH - 1;
  }
  if (typeof POSTAVKA3_FBS_START === "number" && blockName === "Склад продавца") {
    return POSTAVKA3_FBS_START + POSTAVKA3_FBS_WIDTH - 1;
  }
  return -1;
}

function print3CollectOrders_(sheet, pencilCol) {
  const lastRow = sheet.getLastRow();
  const orderData = [];
  if (lastRow < 8) return orderData;
  const arts = sheet.getRange(8, 1, lastRow - 7, 1).getValues();
  const vals = sheet.getRange(8, pencilCol, lastRow - 7, 1).getValues();
  for (let i = 0; i < arts.length; i++) {
    const art = arts[i][0];
    if (!art || print3IsTotalRow_(art)) continue;
    const value = vals[i][0];
    if (value && value !== 0 && value !== "" && !isNaN(parseFloat(value))) {
      orderData.push({ art: art.toString(), packs: parseFloat(value) });
    }
  }
  return orderData;
}
