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
 * Used by both processReceipts and processSpecificFile.
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

  return {
    gastos: gastos,
    incluidos: incluidos,
    excluidos: excluidos,
    ketoDict: loadKetoDictionary(ketoSheet)
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

function isItemKeto(itemName, ketoDictKeys, ketoDict) {
  const name = String(itemName).toLowerCase();
  for (let i = 0; i < ketoDictKeys.length; i++) {
    if (name.indexOf(ketoDictKeys[i]) !== -1) {
      return ketoDict[ketoDictKeys[i]] === 'SI';
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

// ---------------------------------------------------------------------------
// Resumen sheet
// ---------------------------------------------------------------------------

function refreshResumen() {
  const incluidosSheet = _getOrCreate(INCLUIDOS_SHEET_NAME);
  const excluidosSheet = _getOrCreate(EXCLUIDOS_SHEET_NAME);
  const resumen = _getOrCreate(RESUMEN_SHEET_NAME);

  resumen.clear();

  const catIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Categoria');
  const totalIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Total');
  const ketoTotals = _aggregateByCategory(incluidosSheet, catIdx, totalIdx);
  const noKetoTotals = _aggregateByCategory(excluidosSheet, catIdx, totalIdx);

  const allCats = new Set();
  Object.keys(ketoTotals).forEach(function (c) { allCats.add(c); });
  Object.keys(noKetoTotals).forEach(function (c) { allCats.add(c); });

  const rows = [RESUMEN_HEADERS];
  let totalKeto = 0;
  let totalNoKeto = 0;
  const sortedCats = Array.from(allCats).sort();

  sortedCats.forEach(function (cat) {
    const k = ketoTotals[cat] || 0;
    const nk = noKetoTotals[cat] || 0;
    totalKeto += k;
    totalNoKeto += nk;
    rows.push([cat, k, nk, k + nk]);
  });

  rows.push(['', '', '', '']);
  rows.push(['TOTAL', totalKeto, totalNoKeto, totalKeto + totalNoKeto]);

  resumen.getRange(1, 1, rows.length, RESUMEN_HEADERS.length).setValues(rows);

  // Format: bold header + totals row, currency columns
  resumen.getRange(1, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');
  resumen.setFrozenRows(1);
  resumen.getRange(rows.length, 1, 1, RESUMEN_HEADERS.length).setFontWeight('bold');
  if (rows.length > 1) {
    resumen.getRange(2, 2, rows.length - 1, 3).setNumberFormat('#,##0');
  }

  Logger.log('SheetService: resumen actualizado con %s categorías.', sortedCats.length);
}

function _aggregateByCategory(sheet, catIdx, totalIdx) {
  const totals = {};
  if (!sheet) return totals;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return totals;

  const data = sheet.getRange(2, 1, lastRow - 1, CLASSIFIED_ITEM_HEADERS.length).getValues();
  data.forEach(function (row) {
    const cat = row[catIdx] || 'Sin Categoria';
    const amount = Number(row[totalIdx]) || 0;
    totals[cat] = (totals[cat] || 0) + amount;
  });

  return totals;
}
