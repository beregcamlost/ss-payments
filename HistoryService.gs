/**
 * HistoryService.gs
 * Creates monthly history snapshots as separate tabs in the spreadsheet.
 */

/**
 * Archives the given month's data into a dedicated tab.
 * Tab name format: "2026-03 Marzo"
 *
 * @param {Date} [targetDate] - The date whose month to archive. Defaults to today.
 */
function archiveMonth(targetDate) {
  targetDate = targetDate || new Date();
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth(); // 0-based
  var monthStr = String(month + 1).length < 2 ? '0' + (month + 1) : String(month + 1);
  var prefix = year + '-' + monthStr;
  var tabName = prefix + ' ' + MONTH_NAMES_ES[month];

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Idempotent: skip if tab already exists
  if (ss.getSheetByName(tabName)) {
    Logger.log('HistoryService: tab "%s" ya existe, omitiendo.', tabName);
    return;
  }

  // Read and filter data from each sheet
  var gastosData = _filterRowsByMonth(ss.getSheetByName(SHEET_NAME), HEADERS, COL.FECHA, prefix);
  var incluidosData = _filterRowsByMonth(ss.getSheetByName(INCLUIDOS_SHEET_NAME), CLASSIFIED_ITEM_HEADERS, 0, prefix);
  var excluidosData = _filterRowsByMonth(ss.getSheetByName(EXCLUIDOS_SHEET_NAME), CLASSIFIED_ITEM_HEADERS, 0, prefix);
  var noFoodData = _filterRowsByMonth(ss.getSheetByName(NO_FOOD_SHEET_NAME), CLASSIFIED_ITEM_HEADERS, 0, prefix);

  if (gastosData.length === 0 && incluidosData.length === 0 &&
      excluidosData.length === 0 && noFoodData.length === 0) {
    Logger.log('HistoryService: sin datos para "%s", omitiendo.', tabName);
    return;
  }

  var sheet = ss.insertSheet(tabName);
  var currentRow = 1;
  var boldRows = [];

  // Write each section
  currentRow = _writeSection(sheet, currentRow, 'Gastos', HEADERS, gastosData, boldRows);
  currentRow = _writeSection(sheet, currentRow, 'Incluidos', CLASSIFIED_ITEM_HEADERS, incluidosData, boldRows);
  currentRow = _writeSection(sheet, currentRow, 'Excluidos', CLASSIFIED_ITEM_HEADERS, excluidosData, boldRows);
  currentRow = _writeSection(sheet, currentRow, 'No Comestible', CLASSIFIED_ITEM_HEADERS, noFoodData, boldRows);

  // Mini-Resumen section
  currentRow = _writeMiniResumen(sheet, currentRow, incluidosData, excluidosData, noFoodData, boldRows);

  // Bold formatting
  if (boldRows.length > 0) {
    sheet.getRangeList(boldRows.map(function(r) { return r; })).setFontWeight('bold');
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  Logger.log('HistoryService: tab "%s" creado con %s filas.', tabName, currentRow - 1);
  SpreadsheetApp.getActiveSpreadsheet().toast('Mes archivado: ' + tabName, 'Recibos', 5);
}

/**
 * Menu wrapper for archiving the current month.
 */
function archiveCurrentMonth() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Archivar Mes Actual',
    '¿Crear una pestaña con el resumen del mes actual?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  archiveMonth(new Date());
}

/**
 * Filters rows from a sheet where the date column starts with the given YYYY-MM prefix.
 *
 * @param {Sheet|null} sheet - The sheet to read from.
 * @param {string[]} headers - Header array for column count.
 * @param {number} dateColIdx - Column index of the date field.
 * @param {string} prefix - YYYY-MM prefix to filter by.
 * @returns {Array[]} Filtered rows (without headers).
 */
function _filterRowsByMonth(sheet, headers, dateColIdx, prefix) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return data.filter(function(row) {
    var fecha = String(row[dateColIdx]);
    return fecha.indexOf(prefix) === 0;
  });
}

/**
 * Writes a section (title + headers + data rows + blank separator) to the archive sheet.
 * Returns the next available row.
 */
function _writeSection(sheet, startRow, title, headers, rows, boldRows) {
  if (rows.length === 0) return startRow;

  // Section title
  sheet.getRange(startRow, 1).setValue(title);
  boldRows.push('A' + startRow + ':' + String.fromCharCode(64 + headers.length) + startRow);
  startRow++;

  // Headers
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  boldRows.push('A' + startRow + ':' + String.fromCharCode(64 + headers.length) + startRow);
  startRow++;

  // Data rows
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

  // CLP formatting on Total columns
  var totalColIdx = headers.indexOf('Total');
  if (totalColIdx >= 0) {
    sheet.getRange(startRow, totalColIdx + 1, rows.length, 1).setNumberFormat('$#,##0');
  }
  var precioColIdx = headers.indexOf('Precio Unitario');
  if (precioColIdx >= 0) {
    sheet.getRange(startRow, precioColIdx + 1, rows.length, 1).setNumberFormat('$#,##0');
  }

  startRow += rows.length;

  // Blank separator
  startRow++;

  return startRow;
}

/**
 * Writes a mini-resumen section using aggregated category totals.
 */
function _writeMiniResumen(sheet, startRow, ketoRows, noKetoRows, noFoodRows, boldRows) {
  var catIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Categoria');
  var totalIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Total');

  // Section title
  sheet.getRange(startRow, 1).setValue('Mini-Resumen');
  boldRows.push('A' + startRow + ':B' + startRow);
  startRow++;

  var grandTotal = 0;

  function addSubSection(label, rows) {
    sheet.getRange(startRow, 1).setValue(label);
    boldRows.push('A' + startRow + ':B' + startRow);
    startRow++;

    sheet.getRange(startRow, 1, 1, 2).setValues([['Categoria', 'Total']]);
    boldRows.push('A' + startRow + ':B' + startRow);
    startRow++;

    var totals = {};
    rows.forEach(function(row) {
      var cat = row[catIdx] || 'Sin Categoria';
      var amount = Number(row[totalIdx]) || 0;
      totals[cat] = (totals[cat] || 0) + amount;
    });

    var subtotal = 0;
    Object.keys(totals).sort().forEach(function(cat) {
      sheet.getRange(startRow, 1, 1, 2).setValues([[cat, totals[cat]]]);
      sheet.getRange(startRow, 2).setNumberFormat('$#,##0');
      subtotal += totals[cat];
      startRow++;
    });

    sheet.getRange(startRow, 1, 1, 2).setValues([['Subtotal', subtotal]]);
    sheet.getRange(startRow, 2).setNumberFormat('$#,##0');
    boldRows.push('A' + startRow + ':B' + startRow);
    grandTotal += subtotal;
    startRow++;
    startRow++; // blank separator
  }

  addSubSection('KETO (Incluidos)', ketoRows);
  addSubSection('NO KETO (Excluidos)', noKetoRows);
  addSubSection('NO COMESTIBLE', noFoodRows);

  sheet.getRange(startRow, 1, 1, 2).setValues([['TOTAL GENERAL', grandTotal]]);
  sheet.getRange(startRow, 2).setNumberFormat('$#,##0');
  boldRows.push('A' + startRow + ':B' + startRow);
  startRow++;

  return startRow;
}
