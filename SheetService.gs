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

// ---------------------------------------------------------------------------
// Gastos sheet
// ---------------------------------------------------------------------------

function getOrCreateSheet() { return _getOrCreate(SHEET_NAME); }
function ensureHeaders(sheet) { _ensureHeaders(sheet, HEADERS); }

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

function appendReceiptRow(sheet, data, fileName, fileId) {
  sheet.appendRow([
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
  ]);
  Logger.log('SheetService: fila añadida para "%s".', fileName);
}

// ---------------------------------------------------------------------------
// Keto dictionary sheet
// ---------------------------------------------------------------------------

function getOrCreateKetoDict() { return _getOrCreate(KETO_DICT_SHEET_NAME); }

function ensureKetoDictHeaders(sheet) {
  _ensureHeaders(sheet, KETO_DICT_HEADERS);
}

/**
 * Populates the Diccionario Keto sheet with defaults if it has no data rows.
 */
function ensureKetoDictData(sheet) {
  if (sheet.getLastRow() > 1) return; // already has data
  sheet.getRange(2, 1, DEFAULT_KETO_DICT.length, 2).setValues(DEFAULT_KETO_DICT);
  Logger.log('SheetService: diccionario keto poblado con %s entradas.', DEFAULT_KETO_DICT.length);
}

/**
 * Reads the keto dictionary and returns a map: lowercase keyword -> 'SI'|'NO'.
 */
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
 * Checks if an item name matches any keyword in the dictionary.
 * Returns true if keto, false if not. Defaults to false if no match.
 */
function isItemKeto(itemName, ketoDict) {
  const name = String(itemName).toLowerCase();
  const keywords = Object.keys(ketoDict);

  for (let i = 0; i < keywords.length; i++) {
    if (name.indexOf(keywords[i]) !== -1) {
      return ketoDict[keywords[i]] === 'SI';
    }
  }
  return false; // no match → not keto by default
}

// ---------------------------------------------------------------------------
// Incluidos / Excluidos sheets
// ---------------------------------------------------------------------------

function getOrCreateIncluidosSheet() { return _getOrCreate(INCLUIDOS_SHEET_NAME); }
function getOrCreateExcluidosSheet() { return _getOrCreate(EXCLUIDOS_SHEET_NAME); }

function ensureIncluidosHeaders(sheet) { _ensureHeaders(sheet, CLASSIFIED_ITEM_HEADERS); }
function ensureExcluidosHeaders(sheet) { _ensureHeaders(sheet, CLASSIFIED_ITEM_HEADERS); }

/**
 * Classifies items from a receipt as keto/not-keto and appends them
 * to Incluidos or Excluidos respectively.
 */
function appendClassifiedItems(incluidosSheet, excluidosSheet, data, fileId, ketoDict) {
  const items = data.items;
  if (!items || items.length === 0) return;

  const fecha = data.fecha || '';
  const comercio = data.comercio_destinatario || '';
  const categoria = data.categoria || '';

  const ketoRows = [];
  const noKetoRows = [];

  items.forEach(function (item) {
    const qty = item.cantidad || 0;
    const unitPrice = item.precio_unitario || 0;
    const row = [
      fecha,
      comercio,
      categoria,
      item.nombre || '',
      qty,
      unitPrice,
      qty * unitPrice,
      fileId
    ];

    if (isItemKeto(item.nombre, ketoDict)) {
      ketoRows.push(row);
    } else {
      noKetoRows.push(row);
    }
  });

  const cols = CLASSIFIED_ITEM_HEADERS.length;

  if (ketoRows.length > 0) {
    incluidosSheet.getRange(incluidosSheet.getLastRow() + 1, 1, ketoRows.length, cols)
      .setValues(ketoRows);
    Logger.log('SheetService: %s items keto añadidos.', ketoRows.length);
  }

  if (noKetoRows.length > 0) {
    excluidosSheet.getRange(excluidosSheet.getLastRow() + 1, 1, noKetoRows.length, cols)
      .setValues(noKetoRows);
    Logger.log('SheetService: %s items no-keto añadidos.', noKetoRows.length);
  }
}

// ---------------------------------------------------------------------------
// Resumen sheet
// ---------------------------------------------------------------------------

function getOrCreateResumenSheet() { return _getOrCreate(RESUMEN_SHEET_NAME); }

/**
 * Rebuilds the Resumen sheet from Incluidos + Excluidos data.
 * Aggregates totals by Categoria.
 */
function refreshResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resumen = getOrCreateResumenSheet();

  // Clear existing content
  resumen.clear();

  // Read data from both classified sheets
  const ketoTotals = _aggregateByCategory(ss.getSheetByName(INCLUIDOS_SHEET_NAME));
  const noKetoTotals = _aggregateByCategory(ss.getSheetByName(EXCLUIDOS_SHEET_NAME));

  // Collect all categories present in either sheet
  const allCats = new Set();
  Object.keys(ketoTotals).forEach(function (c) { allCats.add(c); });
  Object.keys(noKetoTotals).forEach(function (c) { allCats.add(c); });

  // Build output rows
  const rows = [RESUMEN_HEADERS];
  let totalKeto = 0;
  let totalNoKeto = 0;

  // Sort categories alphabetically
  const sortedCats = Array.from(allCats).sort();

  sortedCats.forEach(function (cat) {
    const k = ketoTotals[cat] || 0;
    const nk = noKetoTotals[cat] || 0;
    totalKeto += k;
    totalNoKeto += nk;
    rows.push([cat, k, nk, k + nk]);
  });

  // Totals row
  rows.push([]);
  rows.push(['TOTAL', totalKeto, totalNoKeto, totalKeto + totalNoKeto]);

  // Write
  resumen.getRange(1, 1, rows.length, RESUMEN_HEADERS.length).setValues(rows);

  // Format header
  resumen.getRange(1, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');
  resumen.setFrozenRows(1);

  // Format totals row
  const totalsRow = rows.length;
  resumen.getRange(totalsRow, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');

  // Format currency columns (B, C, D)
  if (rows.length > 1) {
    resumen.getRange(2, 2, rows.length - 1, 3).setNumberFormat('#,##0');
  }

  Logger.log('SheetService: resumen actualizado con %s categorías.', sortedCats.length);
}

/**
 * Reads a classified items sheet and returns { categoria: totalAmount }.
 */
function _aggregateByCategory(sheet) {
  const totals = {};
  if (!sheet) return totals;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return totals;

  const catCol = CLASSIFIED_ITEM_HEADERS.indexOf('Categoria') + 1;
  const subtotalCol = CLASSIFIED_ITEM_HEADERS.indexOf('Subtotal') + 1;
  const data = sheet.getRange(2, 1, lastRow - 1, CLASSIFIED_ITEM_HEADERS.length).getValues();

  data.forEach(function (row) {
    const cat = row[catCol - 1] || 'Sin Categoria';
    const amount = Number(row[subtotalCol - 1]) || 0;
    totals[cat] = (totals[cat] || 0) + amount;
  });

  return totals;
}
