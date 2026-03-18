/**
 * Code.gs
 * Orchestrator: menu setup, main processing loop, and utility entry points.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recibos')
    .addItem('Procesar Recibos', 'processReceipts')
    .addItem('Reprocesar Todo', 'reprocesarTodo')
    .addItem('Actualizar Resumen', 'updateResumen')
    .addSeparator()
    .addItem('Archivar Mes Actual', 'archiveCurrentMonth')
    .addSeparator()
    .addItem('Instalar Triggers Automáticos', 'installTriggers')
    .addItem('Desinstalar Triggers', 'uninstallTriggers')
    .addSeparator()
    .addItem('Limpiar Datos', 'confirmClearAllData')
    .addItem('Invalidar Cache', 'confirmClearCache')
    .addItem('Configurar API Key', 'setupApiKey')
    .addToUi();
}

/**
 * Menu wrapper for processReceipts. Handles UI errors (Drive access, etc.)
 * then delegates to _processReceiptsCore.
 */
function processReceipts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    _processReceiptsCore(ss);
  } catch (e) {
    Logger.log('processReceipts: error: %s', e.message);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Error durante el procesamiento:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Core processing logic (no UI.alert calls — safe for trigger context).
 * Three-phase processing:
 * 1. CSVs/Excel → bank transaction rows
 * 2. Images/PDFs → receipt rows with keto classification
 * 3. Reconciliation → match receipts to bank transactions
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 */
function _processReceiptsCore(ss) {
  // --- 1. Sheet setup + cache purge ---
  var sheets = initAllSheets();
  purgeExpiredCache();
  var geminiUrl = getGeminiUrl();

  // --- 2. Dedup ---
  var processedIds = getProcessedFileIds(sheets.gastos);
  var processedBankHashes = getProcessedBankHashes(sheets.gastos);

  // --- 3. Fetch all files ---
  var allFiles = getFiles(FOLDER_ID);

  // --- 4. Separate by strategy and filter processed ---
  var csvFiles = [];
  var receiptFiles = [];

  allFiles.forEach(function (file) {
    if (processedIds.has(file.id)) return;
    if (file.strategy === 'csv' || file.strategy === 'excel') {
      csvFiles.push(file);
    } else {
      receiptFiles.push(file);
    }
  });

  var totalPending = csvFiles.length + receiptFiles.length;
  if (totalPending === 0) {
    ss.toast('No hay archivos nuevos por procesar.', 'Recibos', 5);
    return;
  }

  ss.toast('Procesando ' + totalPending + ' archivo(s)...', 'Recibos', 5);

  // --- PHASE 1: Process CSVs/Excel ---
  var phase1 = _processCsvFiles(csvFiles, sheets, processedBankHashes, ss);

  // Flush bank rows before processing receipts
  flushReceiptRows(sheets.gastos, phase1.gastosRows);

  // --- PHASE 2: Process receipt images/PDFs ---
  var receiptBatch = receiptFiles.slice(0, BATCH_SIZE);
  var phase2 = _processReceiptFiles(receiptBatch, sheets, geminiUrl, ss);

  // --- Final flush ---
  flushAll(sheets, phase2.gastosRows, phase2.ketoRows, phase2.noKetoRows, phase2.noFoodRows);

  // --- PHASE 3: Reconciliation ---
  matchAndMerge(sheets.gastos);

  // --- Refresh summary ---
  refreshResumen();

  // --- Summary ---
  var successCount = phase1.successCount + phase2.successCount;
  var errorCount = phase1.errorCount + phase2.errorCount;
  var remaining = receiptFiles.length - receiptBatch.length;
  var summaryLines = [
    'Procesamiento completado.',
    '',
    'Exitosos:  ' + successCount,
    'Errores:   ' + errorCount
  ];
  if (remaining > 0) {
    summaryLines.push('');
    summaryLines.push('Quedan ' + remaining + ' archivos. Ejecuta de nuevo.');
  }
  ss.toast(summaryLines.join('\n'), 'Recibos', 15);
}

/**
 * Phase 1: Process CSV/Excel bank statement files.
 * @returns {{ successCount, errorCount, gastosRows }}
 */
function _processCsvFiles(csvFiles, sheets, processedBankHashes, ss) {
  var successCount = 0;
  var errorCount = 0;
  var gastosRows = [];

  csvFiles.forEach(function (file) {
    ss.toast('Procesando extracto: ' + file.name, 'Recibos', 30);
    Logger.log('processReceipts: procesando extracto "%s"', file.name);

    try {
      var fileHash = getDriveFileHash(file.id);
      var transactions = parseBankFile(file.id, file.strategy, fileHash);

      transactions.forEach(function (tx) {
        var txHash = getBankTxHash(tx);
        if (processedBankHashes.has(txHash)) return;
        processedBankHashes.add(txHash);
        gastosRows.push(buildBankRow(tx, file.name, file.id));
      });

      successCount++;
    } catch (e) {
      Logger.log('processReceipts: error en extracto "%s": %s', file.name, e.message);
      errorCount++;
    }

    Utilities.sleep(RATE_LIMIT_DELAY);
  });

  return { successCount: successCount, errorCount: errorCount, gastosRows: gastosRows };
}

/**
 * Phase 2: Process receipt image/PDF files.
 * @returns {{ successCount, errorCount, gastosRows, ketoRows, noKetoRows, noFoodRows }}
 */
function _processReceiptFiles(receiptBatch, sheets, geminiUrl, ss) {
  var successCount = 0;
  var errorCount = 0;
  var gastosRows = [];
  var ketoRows = [];
  var noKetoRows = [];
  var noFoodRows = [];

  for (var i = 0; i < receiptBatch.length; i++) {
    var file = receiptBatch[i];
    ss.toast('Procesando ' + (i + 1) + ' de ' + receiptBatch.length + ': ' + file.name, 'Recibos', 30);
    Logger.log('processReceipts: [%s/%s] procesando "%s"', i + 1, receiptBatch.length, file.name);

    try {
      var base64 = getFileBase64(file.id);
      var fileHash = getDriveFileHash(file.id);
      var data = extractReceiptData(base64, file.mimeType, file.name, fileHash, geminiUrl);

      if (data) {
        gastosRows.push(buildReceiptRow(data, file.name, file.id));
        var classified = classifyItems(data, sheets);
        classified.ketoRows.forEach(function (r) { ketoRows.push(r); });
        classified.noKetoRows.forEach(function (r) { noKetoRows.push(r); });
        classified.noFoodRows.forEach(function (r) { noFoodRows.push(r); });
        successCount++;
      } else {
        Logger.log('processReceipts: extraccion fallida para "%s", se omite.', file.name);
        errorCount++;
      }
    } catch (fileError) {
      Logger.log('processReceipts: error inesperado en "%s": %s', file.name, fileError.message);
      errorCount++;
    }

    if (gastosRows.length >= FLUSH_INTERVAL) {
      flushAll(sheets, gastosRows, ketoRows, noKetoRows, noFoodRows);
      Logger.log('processReceipts: flush parcial tras %s archivos.', i + 1);
    }

    if (i < receiptBatch.length - 1) {
      Utilities.sleep(RATE_LIMIT_DELAY);
    }
  }

  return {
    successCount: successCount,
    errorCount: errorCount,
    gastosRows: gastosRows,
    ketoRows: ketoRows,
    noKetoRows: noKetoRows,
    noFoodRows: noFoodRows
  };
}

function updateResumen() {
  refreshResumen();
  SpreadsheetApp.getActiveSpreadsheet().toast('Resumen actualizado.', 'Recibos', 5);
}

function confirmClearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Limpiar Datos',
    '¿Borrar todos los datos de Gastos, Incluidos, Excluidos, No Comestible y Resumen?\n\nLos encabezados se mantienen.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  clearAllData();
  ui.alert('Listo', 'Datos eliminados. Usa "Procesar Recibos" para reprocesar.', ui.ButtonSet.OK);
}

function confirmClearCache() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Invalidar Cache',
    '¿Borrar todo el cache de respuestas Gemini?\n\nLa próxima ejecución volverá a llamar a la API para todos los archivos.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  clearCache();
  ui.alert('Listo', 'Cache invalidado.', ui.ButtonSet.OK);
}

function reprocesarTodo() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reprocesar Todo',
    '¿Limpiar todos los datos y reprocesar todos los recibos desde cero?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  clearAllData();
  processReceipts();
}

function setupApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configurar API Key',
    'Ingresa tu clave de API de Gemini (Google AI Studio):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const key = result.getResponseText().trim();
  if (!key) {
    ui.alert('Error', 'La clave no puede estar vacía.', ui.ButtonSet.OK);
    return;
  }

  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('Listo', 'API Key guardada correctamente.', ui.ButtonSet.OK);
}

function processSpecificFile(fileId) {
  if (!fileId) return;

  const sheets = initAllSheets();

  var file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    Logger.log('processSpecificFile: archivo no encontrado: %s', fileId);
    return;
  }

  var mimeType = file.getMimeType();
  var strategy = FILE_STRATEGIES[mimeType];
  if (!strategy) {
    Logger.log('processSpecificFile: tipo no soportado: %s', mimeType);
    return;
  }

  var fileHash = getDriveFileHash(fileId);
  var geminiUrl = getGeminiUrl();

  if (strategy === 'csv' || strategy === 'excel') {
    var transactions = parseBankFile(fileId, strategy, fileHash);
    var rows = transactions.map(function (tx) { return buildBankRow(tx, file.getName(), fileId); });
    flushReceiptRows(sheets.gastos, rows);
  } else {
    var base64 = getFileBase64(fileId);
    var data = extractReceiptData(base64, mimeType, file.getName(), fileHash, geminiUrl);
    if (data) {
      flushReceiptRows(sheets.gastos, [buildReceiptRow(data, file.getName(), fileId)]);
      var classified = classifyItems(data, sheets);
      flushClassifiedRows(sheets.incluidos, sheets.excluidos, sheets.noFood,
        classified.ketoRows, classified.noKetoRows, classified.noFoodRows);
    }
  }

  matchAndMerge(sheets.gastos);
  refreshResumen();
  Logger.log('processSpecificFile: procesado "%s".', file.getName());
}
