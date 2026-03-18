/**
 * ReconciliationService.gs
 * Matches receipt rows to bank transaction rows in the Gastos sheet.
 */

/**
 * Runs the reconciliation process: matches receipts to bank transactions.
 * Updates Estado column for matched rows to "Conciliado".
 *
 * @param {Sheet} gastosSheet - The Gastos sheet.
 */
function matchAndMerge(gastosSheet) {
  var lastRow = gastosSheet.getLastRow();
  if (lastRow < 2) return;

  var data = gastosSheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  var origenCol = COL.ORIGEN;
  var estadoCol = COL.ESTADO;
  var fechaCol = COL.FECHA;
  var comercioCol = COL.COMERCIO___DESTINATARIO;
  var totalCol = COL.TOTAL;
  var descCol = COL.DESCRIPCION;
  var fileIdCol = COL.FILE_ID;

  var bankRows = [];
  var receiptRows = [];

  data.forEach(function (row, idx) {
    var entry = {
      rowIndex: idx + 2,
      fecha: String(row[fechaCol]),
      comercio: String(row[comercioCol]),
      total: Number(row[totalCol]),
      descripcion: String(row[descCol]),
      fileId: String(row[fileIdCol]),
      origen: String(row[origenCol]),
      estado: String(row[estadoCol])
    };

    if (entry.origen === ORIGEN.EXTRACTO && entry.estado === ESTADO.SIN_COMPROBANTE) {
      bankRows.push(entry);
    } else if ((entry.origen === ORIGEN.RECIBO || entry.origen === ORIGEN.TRANSFERENCIA) && entry.estado === ESTADO.SIN_EXTRACTO) {
      receiptRows.push(entry);
    }
  });

  if (bankRows.length === 0 || receiptRows.length === 0) {
    Logger.log('ReconciliationService: nada que conciliar (banco=%s, recibos=%s).', bankRows.length, receiptRows.length);
    return;
  }

  var matchedBankIndices = new Set();
  var updates = [];

  // Tier 1: Deterministic matching
  receiptRows.forEach(function (receipt) {
    if (matchedBankIndices.has('r' + receipt.rowIndex)) return;

    for (var i = 0; i < bankRows.length; i++) {
      if (matchedBankIndices.has(i)) continue;

      var bank = bankRows[i];
      if (_fuzzyStoreMatch(receipt.comercio, bank.descripcion) &&
          receipt.fecha === bank.fecha &&
          receipt.total === bank.total) {
        updates.push({ rowIndex: receipt.rowIndex, estado: ESTADO.CONCILIADO });
        updates.push({ rowIndex: bank.rowIndex, estado: ESTADO.CONCILIADO });
        matchedBankIndices.add(i);
        matchedBankIndices.add('r' + receipt.rowIndex);
        Logger.log('ReconciliationService: match deterministico — "%s" con "%s" ($%s)',
          receipt.comercio, bank.descripcion, receipt.total);
        break;
      }
    }
  });

  // Tier 2: Gemini fallback for unmatched receipts
  receiptRows.forEach(function (receipt) {
    if (matchedBankIndices.has('r' + receipt.rowIndex)) return;

    try {
      var candidates = [];
      var candidateIndices = [];
      bankRows.forEach(function (bank, i) {
        if (matchedBankIndices.has(i)) return;
        if (Math.abs(receipt.total - bank.total) <= receipt.total * RECONCILIATION_TOLERANCE) {
          candidates.push({ fecha: bank.fecha, descripcion: bank.descripcion, monto: bank.total });
          candidateIndices.push(i);
        }
      });

      if (candidates.length === 0) return;

      var receiptData = { comercio: receipt.comercio, fecha: receipt.fecha, total: receipt.total };
      var hash = 'match:' + receipt.fileId;
      var result = matchReceiptToTransaction(receiptData, candidates, hash);

      if (result && result.matched && result.index >= 0 && result.index < candidateIndices.length) {
        var bankIdx = candidateIndices[result.index];
        updates.push({ rowIndex: receipt.rowIndex, estado: ESTADO.CONCILIADO });
        updates.push({ rowIndex: bankRows[bankIdx].rowIndex, estado: ESTADO.CONCILIADO });
        matchedBankIndices.add(bankIdx);
        matchedBankIndices.add('r' + receipt.rowIndex);
        Logger.log('ReconciliationService: match Gemini — "%s" con "%s"',
          receipt.comercio, bankRows[bankIdx].descripcion);
      }
    } catch (e) {
      Logger.log('ReconciliationService: error en Tier 2 para "%s": %s', receipt.comercio, e.message);
    }
  });

  if (updates.length > 0) {
    updates.forEach(function (u) {
      gastosSheet.getRange(u.rowIndex, estadoCol + 1).setValue(u.estado);
    });
    Logger.log('ReconciliationService: %s filas actualizadas a Conciliado.', updates.length);
  }
}

/**
 * Fuzzy store name matching.
 * Normalizes both strings and checks if one contains the other.
 */
function _fuzzyStoreMatch(receiptName, bankDesc) {
  var a = _normalizeStoreName(receiptName);
  var b = _normalizeStoreName(bankDesc);
  if (!a || !b) return false;
  return a.indexOf(b) >= 0 || b.indexOf(a) >= 0;
}

/**
 * Normalizes a store name for fuzzy matching:
 * lowercase, remove accents, strip branch codes/numbers, collapse whitespace.
 */
function _normalizeStoreName(name) {
  if (!name) return '';
  return String(name)
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\b(spa|ltda|s\.?a\.?|suc\.?\s*\d*)\b/gi, '')
    .replace(/\d{3,}/g, '')
    .replace(/[^a-z\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
