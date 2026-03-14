/**
 * SheetService.gs
 * Manages reading from and writing to the Gastos Google Sheet.
 */

/**
 * Returns the target sheet, creating it (and its headers) if it does not exist.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Gastos sheet.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    Logger.log('SheetService: hoja "%s" creada.', SHEET_NAME);
  }

  return sheet;
}

/**
 * Writes the column headers in row 1 if the sheet is empty, then bolds them
 * and freezes the header row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet.
 */
function ensureHeaders(sheet) {
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '') {
    return; // Headers already present
  }

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  Logger.log('SheetService: encabezados escritos en la hoja "%s".', SHEET_NAME);
}

/**
 * Reads the "File ID" column (column N, index 14) and returns a Set of all
 * values found. Used for O(1) deduplication before processing new files.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet.
 * @returns {Set<string>} Set of already-processed Drive file IDs.
 */
function getProcessedFileIds(sheet) {
  const lastRow = sheet.getLastRow();
  const processed = new Set();

  if (lastRow < 2) {
    return processed; // No data rows yet
  }

  // Column N is index 14 (1-based), which is the "File ID" column
  const fileIdColumnIndex = HEADERS.indexOf('File ID') + 1; // 1-based
  const values = sheet
    .getRange(2, fileIdColumnIndex, lastRow - 1, 1)
    .getValues();

  values.forEach(function (row) {
    const id = row[0];
    if (id) {
      processed.add(String(id));
    }
  });

  Logger.log('SheetService: %s archivos ya procesados.', processed.size);
  return processed;
}

/**
 * Converts a string amount (possibly containing dots or commas as thousands
 * separators) to an integer. Returns 0 for empty or unparseable values.
 *
 * @param {string} value - The amount string, e.g. "15.990" or "15990" or "".
 * @returns {number} Integer CLP amount or 0.
 */
function parseAmount(value) {
  if (!value || value === '') {
    return 0;
  }
  // Strip any character that is not a digit
  const digits = String(value).replace(/[^\d]/g, '');
  const parsed = parseInt(digits, 10);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Maps the Gemini extraction result to a row array and appends it to the sheet.
 * Column order must match HEADERS exactly.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet    - The target sheet.
 * @param {object}                             data     - Extracted JSON from Gemini.
 * @param {string}                             fileName - Original Drive file name.
 * @param {string}                             fileId   - Drive file ID.
 */
function appendReceiptRow(sheet, data, fileName, fileId) {
  const row = [
    data.fecha || '',                       // A - Fecha
    data.tipo || '',                        // B - Tipo
    data.comercio_destinatario || '',       // C - Comercio / Destinatario
    data.rut || '',                         // D - RUT
    data.categoria || '',                   // E - Categoria
    data.descripcion || '',                 // F - Descripcion
    parseAmount(data.neto),                 // G - Neto (integer)
    parseAmount(data.iva),                  // H - IVA (integer)
    parseAmount(data.total),                // I - Total (integer)
    data.medio_pago || '',                  // J - Medio de Pago
    data.banco_destino || '',               // K - Banco Destino
    data.numero_operacion || '',            // L - N° Operacion
    fileName,                               // M - Archivo
    fileId,                                 // N - File ID
    new Date()                              // O - Procesado (timestamp)
  ];

  sheet.appendRow(row);
  Logger.log('SheetService: fila añadida para "%s".', fileName);
}
