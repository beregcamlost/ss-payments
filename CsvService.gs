/**
 * CsvService.gs
 * Parses CSV and Excel bank statements into structured transaction rows.
 */

/**
 * Processes a CSV or Excel file into bank transaction objects.
 *
 * @param {string} fileId - Drive file ID.
 * @param {string} strategy - 'csv' or 'excel'.
 * @param {string} fileHash - Cache key for Gemini calls.
 * @returns {object[]} Array of transaction objects: { fecha, descripcion, monto, banco, comercio, categoria }
 */
function parseBankFile(fileId, strategy, fileHash) {
  var rows;
  if (strategy === 'excel') {
    rows = excelToRows(fileId);
  } else {
    rows = csvToRows(fileId);
  }

  if (!rows || rows.length < 2) {
    Logger.log('CsvService: archivo con menos de 2 filas, omitido.');
    return [];
  }

  var mapping = detectCsvColumns(rows, fileHash);
  if (!mapping) {
    Logger.log('CsvService: no se pudo detectar columnas, archivo omitido.');
    return [];
  }

  var startRow = (mapping.header_row >= 0) ? mapping.header_row + 1 : 0;
  var dataRows = rows.slice(startRow);

  var transactions = [];
  var descriptions = [];

  dataRows.forEach(function (row) {
    var fechaRaw = String(row[mapping.fecha_col] || '').trim();
    var descripcion = String(row[mapping.descripcion_col] || '').trim();
    var montoRaw = String(row[mapping.monto_col] || '').trim();

    if (!fechaRaw || !descripcion || !montoRaw) return;

    var fecha = _parseBankDate(fechaRaw, mapping.date_format);
    var monto = parseAmount(montoRaw);

    if (monto === 0) return;

    transactions.push({
      fecha: fecha,
      descripcion: descripcion,
      monto: monto,
      banco: mapping.banco || ''
    });
    descriptions.push(descripcion);
  });

  if (transactions.length === 0) return [];

  var categories = categorizeBankDescriptions(descriptions, fileHash + ':categories');
  var catMap = {};
  if (categories && categories.items) {
    categories.items.forEach(function (item) {
      catMap[item.descripcion] = { categoria: item.categoria, comercio: item.comercio };
    });
  }

  transactions.forEach(function (tx) {
    var cat = catMap[tx.descripcion] || {};
    tx.categoria = cat.categoria || 'Otro';
    tx.comercio = cat.comercio || tx.descripcion;
  });

  Logger.log('CsvService: %s transacciones extraidas.', transactions.length);
  return transactions;
}

/**
 * Parses a date string using the detected format.
 * Returns YYYY-MM-DD format.
 */
function _parseBankDate(dateStr, format) {
  if (!dateStr || !format) return dateStr;

  var parts;
  format = format.toUpperCase();

  if (format.indexOf('DD/MM/YYYY') >= 0 || format.indexOf('DD-MM-YYYY') >= 0) {
    parts = dateStr.split(/[\/\-\.]/);
    if (parts.length >= 3) return parts[2] + '-' + _pad(parts[1]) + '-' + _pad(parts[0]);
  } else if (format.indexOf('MM/DD/YYYY') >= 0) {
    parts = dateStr.split(/[\/\-\.]/);
    if (parts.length >= 3) return parts[2] + '-' + _pad(parts[0]) + '-' + _pad(parts[1]);
  } else if (format.indexOf('YYYY-MM-DD') >= 0 || format.indexOf('YYYY/MM/DD') >= 0) {
    parts = dateStr.split(/[\/\-\.]/);
    if (parts.length >= 3) return parts[0] + '-' + _pad(parts[1]) + '-' + _pad(parts[2]);
  }

  return dateStr;
}

function _pad(s) {
  s = String(s);
  return s.length < 2 ? '0' + s : s;
}

/**
 * Generates a dedup key for a bank transaction.
 */
function getBankTxHash(tx) {
  var raw = tx.fecha + '|' + tx.descripcion + '|' + tx.monto;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
    .join('');
}
