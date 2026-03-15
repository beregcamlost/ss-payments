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

  const noFood = _getOrCreate(NO_FOOD_SHEET_NAME);
  _ensureHeaders(noFood, CLASSIFIED_ITEM_HEADERS);

  const ketoSheet = _getOrCreate(KETO_DICT_SHEET_NAME);
  _ensureHeaders(ketoSheet, KETO_DICT_HEADERS);
  ensureKetoDictData(ketoSheet);

  initCacheSheet();

  const ketoDict = loadKetoDictionary(ketoSheet);
  const ketoPatterns = _buildKetoPatterns(ketoDict);

  return {
    gastos: gastos,
    incluidos: incluidos,
    excluidos: excluidos,
    noFood: noFood,
    ketoDict: ketoDict,
    ketoPatterns: ketoPatterns
  };
}

/**
 * Clears all data (keeping headers) from Gastos, Incluidos, Excluidos,
 * No Comestible, and Resumen so a full reprocess starts fresh.
 */
function clearAllData() {
  const names = [SHEET_NAME, INCLUIDOS_SHEET_NAME, EXCLUIDOS_SHEET_NAME, NO_FOOD_SHEET_NAME, RESUMEN_SHEET_NAME];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  names.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  });
  Logger.log('SheetService: todas las hojas limpiadas.');
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
  var origen = (data.tipo === 'transferencia') ? ORIGEN.TRANSFERENCIA : ORIGEN.RECIBO;
  return [
    data.fecha || '',
    data.tipo || '',
    data.comercio_destinatario || '',
    data.categoria || '',
    data.descripcion || '',
    parseAmount(data.total),
    fileName,
    fileId,
    new Date(),
    origen,
    ESTADO.SIN_EXTRACTO
  ];
}

/**
 * Builds a row array for a bank transaction (from CSV/Excel).
 */
function buildBankRow(tx, fileName, fileId) {
  return [
    tx.fecha || '',
    'extracto',
    tx.comercio || tx.descripcion || '',
    tx.categoria || 'Otro',
    tx.descripcion || '',
    tx.monto || 0,
    fileName,
    fileId,
    new Date(),
    ORIGEN.EXTRACTO,
    ESTADO.SIN_COMPROBANTE
  ];
}

/**
 * Returns a Set of bank transaction hashes already in Gastos.
 * Used for CSV deduplication.
 */
function getProcessedBankHashes(sheet) {
  var hashes = new Set();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return hashes;

  var origenCol = HEADERS.indexOf('Origen');
  var fechaCol = HEADERS.indexOf('Fecha');
  var descCol = HEADERS.indexOf('Descripcion');
  var totalCol = HEADERS.indexOf('Total');

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  data.forEach(function (row) {
    if (row[origenCol] === ORIGEN.EXTRACTO) {
      var tx = { fecha: row[fechaCol], descripcion: row[descCol], monto: row[totalCol] };
      hashes.add(getBankTxHash(tx));
    }
  });

  Logger.log('SheetService: %s transacciones bancarias ya procesadas.', hashes.size);
  return hashes;
}

/**
 * Batch-writes accumulated receipt rows to the Gastos sheet.
 */
function flushReceiptRows(sheet, rows) {
  if (rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, HEADERS.length)
    .setValues(rows);
  // CLP format on Total column (F = index 6, but after removing Banco Destino it's column 6)
  const totalCol = HEADERS.indexOf('Total') + 1;
  sheet.getRange(startRow, totalCol, rows.length, 1).setNumberFormat('$#,##0');
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
 * Pre-compiles regex patterns for all keto dictionary keys.
 * Sorted longest-first so "chocolate bitter" matches before "chocolate".
 */
function _buildKetoPatterns(ketoDict) {
  return Object.keys(ketoDict)
    .sort(function (a, b) { return b.length - a.length; })
    .map(function (key) {
      var escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      return { key: key, regex: new RegExp('(?:^|\\s)' + escaped + '(?:\\s|$)'), value: ketoDict[key] };
    });
}

/**
 * Classifies an item using pre-compiled keto patterns.
 * Returns 'SI', 'NO', 'NO_FOOD', or null if no match.
 */
function classifyItemKeto(itemName, ketoPatterns) {
  const name = ' ' + String(itemName).toLowerCase() + ' ';
  for (var i = 0; i < ketoPatterns.length; i++) {
    if (ketoPatterns[i].regex.test(name)) {
      return ketoPatterns[i].value;
    }
  }
  return null;
}

/**
 * Uses Gemini to classify items not found in the dictionary.
 * Returns a map: { itemName: 'SI' | 'NO' | 'NO_FOOD' }
 * Also writes new entries back to the Diccionario Keto sheet for future runs.
 */
function classifyUnknownItemsWithGemini(itemNames) {
  if (itemNames.length === 0) return {};

  const unique = Array.from(new Set(itemNames.map(function (n) { return String(n).trim(); })));
  Logger.log('SheetService: clasificando %s items desconocidos con Gemini.', unique.length);

  const prompt = `Clasifica cada item de la siguiente lista.
Para cada uno responde EXACTAMENTE con una de estas categorías:
- "SI" = es comida/bebida compatible con dieta keto (bajo en carbohidratos)
- "NO" = es comida/bebida pero NO es keto (alto en carbohidratos: pan, arroz, azúcar, bebidas dulces, etc.)
- "NO_FOOD" = NO es comida ni bebida (productos de limpieza, higiene, ropa, bolsas, etc.)

Lista de items:
${unique.map(function (name, i) { return (i + 1) + '. ' + name; }).join('\n')}`;

  const schema = {
    type: 'object',
    properties: {
      items: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            nombre: { type: 'string', description: 'Nombre del item tal como aparece en la lista' },
            clasificacion: { type: 'string', description: 'SI, NO, o NO_FOOD' },
            palabra_clave: { type: 'string', description: 'Palabra clave corta para el diccionario (ej: "merken", "alfajor", "bolsa plastica")' }
          },
          required: ['nombre', 'clasificacion', 'palabra_clave']
        }
      }
    },
    required: ['items']
  };

  const payload = {
    system_instruction: {
      parts: [{ text: 'Eres un nutricionista experto en dieta keto. Clasificas productos de supermercado y comercio chileno.' }]
    },
    contents: [{
      role: 'user',
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: schema,
      temperature: 0.1
    }
  };

  var parsed = callGemini(payload, 'clasificación keto');
  if (!parsed) return {};

  var results = {};
  var newDictEntries = [];
  var geminiItems = parsed.items || [];

  geminiItems.forEach(function (item) {
    var cls = String(item.clasificacion).trim().toUpperCase();
    if (cls !== 'SI' && cls !== 'NO' && cls !== 'NO_FOOD') cls = 'NO_FOOD';
    results[item.nombre] = cls;

    var keyword = String(item.palabra_clave || '').trim().toLowerCase();
    if (keyword) {
      newDictEntries.push([keyword, cls]);
    }
  });

  // Save new keywords to Diccionario Keto for future runs
  if (newDictEntries.length > 0) {
    var ketoSheet = _getOrCreate(KETO_DICT_SHEET_NAME);
    ketoSheet.getRange(ketoSheet.getLastRow() + 1, 1, newDictEntries.length, 2)
      .setValues(newDictEntries);
    Logger.log('SheetService: %s nuevas entradas añadidas al diccionario keto.', newDictEntries.length);
  }

  return results;
}

// ---------------------------------------------------------------------------
// Incluidos / Excluidos — classified items
// ---------------------------------------------------------------------------

/**
 * Classifies items from a receipt as keto/not-keto.
 * Returns { ketoRows: [], noKetoRows: [] } for batch flushing.
 */
function classifyItems(data, sheets) {
  const items = data.items;
  const ketoRows = [];
  const noKetoRows = [];
  const noFoodRows = [];
  if (!items || items.length === 0) return { ketoRows: ketoRows, noKetoRows: noKetoRows, noFoodRows: noFoodRows };

  const fecha = data.fecha || '';
  const comercio = data.comercio_destinatario || '';
  const categoria = data.categoria || '';

  // First pass: classify using pre-compiled patterns, collect unknowns
  var pending = [];
  var classified = [];

  items.forEach(function (item) {
    const qty = item.cantidad || 0;
    const unitPrice = item.precio_unitario || 0;
    const row = [fecha, comercio, categoria, item.nombre || '', qty, unitPrice, qty * unitPrice];
    const cls = classifyItemKeto(item.nombre, sheets.ketoPatterns);

    if (cls) {
      classified.push({ row: row, classification: cls });
    } else {
      pending.push({ row: row, name: item.nombre || '' });
    }
  });

  // Second pass: use Gemini for unknowns
  if (pending.length > 0) {
    var unknownNames = pending.map(function (p) { return p.name; });
    var geminiResults = classifyUnknownItemsWithGemini(unknownNames);

    pending.forEach(function (p) {
      var cls = geminiResults[p.name] || 'NO_FOOD';
      classified.push({ row: p.row, classification: cls });
    });

    // Update in-memory dict + patterns so subsequent receipts benefit
    var ketoSheet = _getOrCreate(KETO_DICT_SHEET_NAME);
    sheets.ketoDict = loadKetoDictionary(ketoSheet);
    sheets.ketoPatterns = _buildKetoPatterns(sheets.ketoDict);
  }

  // Sort into buckets
  classified.forEach(function (c) {
    if (c.classification === 'SI') {
      ketoRows.push(c.row);
    } else if (c.classification === 'NO') {
      noKetoRows.push(c.row);
    } else {
      noFoodRows.push(c.row);
    }
  });

  return { ketoRows: ketoRows, noKetoRows: noKetoRows, noFoodRows: noFoodRows };
}

/**
 * Batch-writes accumulated classified rows to the Incluidos/Excluidos sheets.
 */
// Precomputed column indices for classified item sheets
const CLASSIFIED_COLS = CLASSIFIED_ITEM_HEADERS.length;
const CLASSIFIED_PRECIO_COL = CLASSIFIED_ITEM_HEADERS.indexOf('Precio Unitario') + 1;
const CLASSIFIED_TOTAL_COL = CLASSIFIED_ITEM_HEADERS.indexOf('Total') + 1;

function _flushRowsToSheet(sheet, rows, label) {
  if (rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, CLASSIFIED_COLS).setValues(rows);
  sheet.getRange(startRow, CLASSIFIED_PRECIO_COL, rows.length, 1).setNumberFormat('$#,##0');
  sheet.getRange(startRow, CLASSIFIED_TOTAL_COL, rows.length, 1).setNumberFormat('$#,##0');
  Logger.log('SheetService: %s items %s escritos.', rows.length, label);
}

function flushClassifiedRows(incluidosSheet, excluidosSheet, noFoodSheet, allKetoRows, allNoKetoRows, allNoFoodRows) {
  _flushRowsToSheet(incluidosSheet, allKetoRows, 'keto');
  _flushRowsToSheet(excluidosSheet, allNoKetoRows, 'no-keto');
  _flushRowsToSheet(noFoodSheet, allNoFoodRows, 'no-comestible');
}

/**
 * Flushes all accumulated rows (gastos + classified) and resets the buffers.
 * Used for incremental flushing to preserve partial progress.
 */
function flushAll(sheets, gastosRows, allKetoRows, allNoKetoRows, allNoFoodRows) {
  flushReceiptRows(sheets.gastos, gastosRows);
  flushClassifiedRows(sheets.incluidos, sheets.excluidos, sheets.noFood, allKetoRows, allNoKetoRows, allNoFoodRows);
  // Clear the arrays (mutate in place so caller sees the reset)
  gastosRows.length = 0;
  allKetoRows.length = 0;
  allNoKetoRows.length = 0;
  allNoFoodRows.length = 0;
}

// ---------------------------------------------------------------------------
// Resumen sheet — includes Gastos totals for full reconciliation
// ---------------------------------------------------------------------------

function refreshResumen() {
  const gastosSheet = _getOrCreate(SHEET_NAME);
  const incluidosSheet = _getOrCreate(INCLUIDOS_SHEET_NAME);
  const excluidosSheet = _getOrCreate(EXCLUIDOS_SHEET_NAME);
  const noFoodSheet = _getOrCreate(NO_FOOD_SHEET_NAME);
  const resumen = _getOrCreate(RESUMEN_SHEET_NAME);

  resumen.clear();

  // Aggregate from classified items
  const catIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Categoria');
  const totalIdx = CLASSIFIED_ITEM_HEADERS.indexOf('Total');
  const ketoTotals = _aggregateByCategory(incluidosSheet, catIdx, totalIdx);
  const noKetoTotals = _aggregateByCategory(excluidosSheet, catIdx, totalIdx);
  const noFoodTotals = _aggregateByCategory(noFoodSheet, catIdx, totalIdx);

  // Aggregate from Gastos (the true total per category)
  const gastosCatIdx = HEADERS.indexOf('Categoria');
  const gastosTotalIdx = HEADERS.indexOf('Total');
  const gastosTotals = _aggregateByCategory(gastosSheet, gastosCatIdx, gastosTotalIdx);

  // Compute "Sin Detalle" per category (Gastos total minus all classified items)
  const sinDetalleTotals = {};
  Object.keys(gastosTotals).forEach(function (cat) {
    const k = ketoTotals[cat] || 0;
    const nk = noKetoTotals[cat] || 0;
    const nf = noFoodTotals[cat] || 0;
    const remainder = Math.max(0, gastosTotals[cat] - k - nk - nf);
    if (remainder > 0) sinDetalleTotals[cat] = remainder;
  });

  var rows = [];
  var boldRows = [];
  var grandTotal = 0;

  function addSection(title, subtotalLabel, totalsMap) {
    rows.push([title, '']);
    boldRows.push(rows.length);
    rows.push(['Categoria', 'Total']);
    boldRows.push(rows.length);
    var subtotal = 0;
    Object.keys(totalsMap).sort().forEach(function (cat) {
      rows.push([cat, totalsMap[cat]]);
      subtotal += totalsMap[cat];
    });
    rows.push([subtotalLabel, subtotal]);
    boldRows.push(rows.length);
    grandTotal += subtotal;
    rows.push(['', '']);
  }

  addSection('KETO (Incluidos)', 'SUBTOTAL KETO', ketoTotals);
  addSection('NO KETO (Excluidos)', 'SUBTOTAL NO KETO', noKetoTotals);
  addSection('NO COMESTIBLE', 'SUBTOTAL NO COMESTIBLE', noFoodTotals);
  addSection('SIN DETALLE', 'SUBTOTAL SIN DETALLE', sinDetalleTotals);

  rows.push(['TOTAL GENERAL', grandTotal]);
  boldRows.push(rows.length);

  // Write all rows
  resumen.getRange(1, 1, rows.length, 2).setValues(rows);

  // Bold formatting batched into a single service call
  resumen.getRangeList(boldRows.map(function (r) { return 'A' + r + ':B' + r; }))
    .setFontWeight('bold');

  // CLP format on all Total column values (column B)
  resumen.getRange(1, 2, rows.length, 1).setNumberFormat('$#,##0');

  Logger.log('SheetService: resumen actualizado (%s filas).', rows.length);
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
