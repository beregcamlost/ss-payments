/**
 * SheetService.gs
 * Manages reading from and writing to all Google Sheets tabs.
 */

// ---------------------------------------------------------------------------
// Generic helpers
// ---------------------------------------------------------------------------

function _getOrCreate(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('SheetService: hoja "%s" creada.', name);
  }
  return sheet;
}

function _ensureHeaders(sheet, headers) {
  if (sheet.getRange(1, 1).getValue() !== '') return;
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Initializes all sheets and returns them with the keto dictionary loaded.
 * Keys are sorted longest-first for correct substring matching priority.
 */
function initAllSheets() {
  const gastos = _getOrCreate(SHEET_NAME);
  _ensureHeaders(gastos, HEADERS);

  const incluidos = _getOrCreate(INCLUIDOS_SHEET_NAME);
  _ensureHeaders(incluidos, CLASSIFIED_ITEM_HEADERS);

  const excluidos = _getOrCreate(EXCLUIDOS_SHEET_NAME);
  _ensureHeaders(excluidos, CLASSIFIED_ITEM_HEADERS);

  const ketoSheet = _getOrCreate(KETO_DICT_SHEET_NAME);
  _ensureHeaders(ketoSheet, KETO_DICT_HEADERS);
  ensureKetoDictData(ketoSheet);

  const ketoDict = loadKetoDictionary(ketoSheet);
  // Sort keys longest-first so "chocolate" matches before "te", "sandwich" before "pollo"
  const ketoKeys = Object.keys(ketoDict).sort(function (a, b) { return b.length - a.length; });

  return {
    gastos: gastos,
    incluidos: incluidos,
    excluidos: excluidos,
    ketoDict: ketoDict,
    ketoKeys: ketoKeys
  };
}

// ---------------------------------------------------------------------------
// Gastos sheet
// ---------------------------------------------------------------------------

function getProcessedFileIds(sheet) {
  const lastRow = sheet.getLastRow();
  const processed = new Set();
  if (lastRow < 2) return processed;

  const col = HEADERS.indexOf('File ID') + 1;
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  values.forEach(function (row) {
    if (row[0]) processed.add(String(row[0]));
  });
  Logger.log('SheetService: %s archivos ya procesados.', processed.size);
  return processed;
}

function parseAmount(value) {
  if (!value || value === '') return 0;
  const digits = String(value).replace(/[^\d]/g, '');
  const parsed = parseInt(digits, 10);
  return isNaN(parsed) ? 0 : parsed;
}

function buildReceiptRow(data, fileName, fileId) {
  return [
    data.fecha || '',
    data.tipo || '',
    data.comercio_destinatario || '',
    data.categoria || '',
    data.descripcion || '',
    parseAmount(data.total),
    data.banco_destino || '',
    fileName,
    fileId,
    new Date()
  ];
}

/**
 * Batch-writes accumulated receipt rows to the Gastos sheet.
 */
function flushReceiptRows(sheet, rows) {
  if (rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length)
    .setValues(rows);
  Logger.log('SheetService: %s filas de gastos escritas.', rows.length);
}

// ---------------------------------------------------------------------------
// Keto dictionary
// ---------------------------------------------------------------------------

function ensureKetoDictData(sheet) {
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, DEFAULT_KETO_DICT.length, 2).setValues(DEFAULT_KETO_DICT);
  Logger.log('SheetService: diccionario keto poblado con %s entradas.', DEFAULT_KETO_DICT.length);
}

function loadKetoDictionary(sheet) {
  const lastRow = sheet.getLastRow();
  const dict = {};
  if (lastRow < 2) return dict;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach(function (row) {
    const keyword = String(row[0]).trim().toLowerCase();
    const keto = String(row[1]).trim().toUpperCase();
    if (keyword) dict[keyword] = keto;
  });
  return dict;
}

/**
 * Checks if an item name matches any keyword using word-boundary matching.
 * Keywords are expected pre-sorted longest-first so "chocolate" beats "te".
 */
function isItemKeto(itemName, ketoDictKeys, ketoDict) {
  const name = ' ' + String(itemName).toLowerCase() + ' ';
  for (let i = 0; i < ketoDictKeys.length; i++) {
    var key = ketoDictKeys[i];
    // Word-boundary match: keyword must be surrounded by spaces or string edges
    var escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var re = new RegExp('(?:^|\\s)' + escaped + '(?:\\s|$)');
    if (re.test(name)) {
      return ketoDict[key] === 'SI';
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Incluidos / Excluidos — classified items
// ---------------------------------------------------------------------------

/**
 * Classifies items from a receipt as keto/not-keto.
 * Returns { ketoRows: [], noKetoRows: [] } for batch flushing.
 */
function classifyItems(data, ketoDict, ketoDictKeys) {
  const items = data.items;
  const ketoRows = [];
  const noKetoRows = [];
  if (!items || items.length === 0) return { ketoRows: ketoRows, noKetoRows: noKetoRows };

  const fecha = data.fecha || '';
  const comercio = data.comercio_destinatario || '';
  const categoria = data.categoria || '';

  items.forEach(function (item) {
    const qty = item.cantidad || 0;
    const unitPrice = item.precio_unitario || 0;
    const row = [fecha, comercio, categoria, item.nombre || '', qty, unitPrice, qty * unitPrice];

    if (isItemKeto(item.nombre, ketoDictKeys, ketoDict)) {
      ketoRows.push(row);
    } else {
      noKetoRows.push(row);
    }
  });

  return { ketoRows: ketoRows, noKetoRows: noKetoRows };
}

/**
 * Batch-writes accumulated classified rows to the Incluidos/Excluidos sheets.
 */
function flushClassifiedRows(incluidosSheet, excluidosSheet, allKetoRows, allNoKetoRows) {
  const cols = CLASSIFIED_ITEM_HEADERS.length;

  if (allKetoRows.length > 0) {
    incluidosSheet.getRange(incluidosSheet.getLastRow() + 1, 1, allKetoRows.length, cols)
      .setValues(allKetoRows);
    Logger.log('SheetService: %s items keto escritos.', allKetoRows.length);
  }

  if (allNoKetoRows.length > 0) {
    excluidosSheet.getRange(excluidosSheet.getLastRow() + 1, 1, allNoKetoRows.length, cols)
      .setValues(allNoKetoRows);
    Logger.log('SheetService: %s items no-keto escritos.', allNoKetoRows.length);
  }
}

/**
 * Flushes all accumulated rows (gastos + classified) and resets the buffers.
 * Used for incremental flushing to preserve partial progress.
 */
function flushAll(sheets, gastosRows, allKetoRows, allNoKetoRows) {
  flushReceiptRows(sheets.gastos, gastosRows);
  flushClassifiedRows(sheets.incluidos, sheets.excluidos, allKetoRows, allNoKetoRows);
  // Clear the arrays (mutate in place so caller sees the reset)
  gastosRows.length = 0;
  allKetoRows.length = 0;
  allNoKetoRows.length = 0;
}

// ---------------------------------------------------------------------------
// Resumen sheet — includes Gastos totals for full reconciliation
// ---------------------------------------------------------------------------

function refreshResumen() {
  const gastosSheet = _getOrCreate(SHEET_NAME);
  const incluidosSheet = _getOrCreate(INCLUIDOS_SHEET_NAME);
  const excluidosSheet = _getOrCreate(EXCLUIDOS_SHEET_NAME);
  const resumen = _getOrCreate(RESUMEN_SHEET_NAME);

  resumen.clear();

  // Aggregate from classified items
  const catIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Categoria');
  const totalIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Total');
  const ketoTotals = _aggregateByCategory(incluidosSheet, catIdx, totalIdx);
  const noKetoTotals = _aggregateByCategory(excluidosSheet, catIdx, totalIdx);

  // Aggregate from Gastos (the true total per category, includes transfers + item-less receipts)
  const gastosCatIdx = HEADERS.indexOf('Categoria');
  const gastosTotalIdx = HEADERS.indexOf('Total');
  const gastosTotals = _aggregateByCategory(gastosSheet, gastosCatIdx, gastosTotalIdx);

  // Collect all categories from all sources
  const allCats = new Set();
  Object.keys(ketoTotals).forEach(function (c) { allCats.add(c); });
  Object.keys(noKetoTotals).forEach(function (c) { allCats.add(c); });
  Object.keys(gastosTotals).forEach(function (c) { allCats.add(c); });

  var rows = [RESUMEN_HEADERS];
  var totalKeto = 0;
  var totalNoKeto = 0;
  var totalSinDetalle = 0;
  var totalGeneral = 0;
  const sortedCats = Array.from(allCats).sort();

  sortedCats.forEach(function (cat) {
    const k = ketoTotals[cat] || 0;
    const nk = noKetoTotals[cat] || 0;
    const gastosTotal = gastosTotals[cat] || 0;
    // "Sin Detalle" = Gastos total minus what's accounted for in items
    const sinDetalle = Math.max(0, gastosTotal - k - nk);
    totalKeto += k;
    totalNoKeto += nk;
    totalSinDetalle += sinDetalle;
    totalGeneral += gastosTotal;
    rows.push([cat, k, nk, sinDetalle, gastosTotal]);
  });

  rows.push(['', '', '', '', '']);
  rows.push(['TOTAL', totalKeto, totalNoKeto, totalSinDetalle, totalGeneral]);

  resumen.getRange(1, 1, rows.length, RESUMEN_HEADERS.length).setValues(rows);

  // Format
  resumen.getRange(1, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');
  resumen.setFrozenRows(1);
  resumen.getRange(rows.length, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');
  if (rows.length > 1) {
    resumen.getRange(2, 2, rows.length - 1, 4).setNumberFormat('#,##0');
  }

  Logger.log('SheetService: resumen actualizado con %s categorías.', sortedCats.length);
}

function _aggregateByCategory(sheet, catIdx, totalIdx) {
  const totals = {};
  if (!sheet) return totals;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return totals;

  // Read only the columns we need (widest possible header set)
  var numCols = Math.max(catIdx, totalIdx) + 1;
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  data.forEach(function (row) {
    const cat = row[catIdx] || 'Sin Categoria';
    const amount = Number(row[totalIdx]) || 0;
    totals[cat] = (totals[cat] || 0) + amount;
  });

  return totals;
}
