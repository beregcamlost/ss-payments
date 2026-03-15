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
    .addItem('Limpiar Datos', 'confirmClearAllData')
    .addItem('Invalidar Cache', 'confirmClearCache')
    .addItem('Configurar API Key', 'setupApiKey')
    .addToUi();
}

/**
 * Main orchestrator. Three-phase processing:
 * 1. CSVs/Excel → bank transaction rows
 * 2. Images/PDFs → receipt rows with keto classification
 * 3. Reconciliation → match receipts to bank transactions
 */
function processReceipts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Sheet setup ---
  const sheets = initAllSheets();
  const geminiUrl = getGeminiUrl();

  // --- 2. Dedup ---
  const processedIds = getProcessedFileIds(sheets.gastos);
  const processedBankHashes = getProcessedBankHashes(sheets.gastos);

  // --- 3. Fetch all files ---
  let allFiles;
  try {
    allFiles = getFiles(FOLDER_ID);
  } catch (driveError) {
    Logger.log('processReceipts: error accediendo a Drive: %s', driveError.message);
    ui.alert('Error', 'No se pudo leer la carpeta de Drive:\n' + driveError.message, ui.ButtonSet.OK);
    return;
  }

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

  let successCount = 0;
  let errorCount = 0;
  const gastosRows = [];
  const allKetoRows = [];
  const allNoKetoRows = [];
  const allNoFoodRows = [];

  // --- PHASE 1: Process CSVs/Excel ---
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

  // Flush bank rows before processing receipts
  flushReceiptRows(sheets.gastos, gastosRows);
  gastosRows.length = 0;

  // --- PHASE 2: Process receipt images/PDFs ---
  var receiptBatch = receiptFiles.slice(0, BATCH_SIZE);

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
        classified.ketoRows.forEach(function (r) { allKetoRows.push(r); });
        classified.noKetoRows.forEach(function (r) { allNoKetoRows.push(r); });
        classified.noFoodRows.forEach(function (r) { allNoFoodRows.push(r); });
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
      flushAll(sheets, gastosRows, allKetoRows, allNoKetoRows, allNoFoodRows);
      Logger.log('processReceipts: flush parcial tras %s archivos.', i + 1);
    }

    if (i < receiptBatch.length - 1) {
      Utilities.sleep(RATE_LIMIT_DELAY);
    }
  }

  // --- Final flush ---
  flushAll(sheets, gastosRows, allKetoRows, allNoKetoRows, allNoFoodRows);

  // --- PHASE 3: Reconciliation ---
  matchAndMerge(sheets.gastos);

  // --- Refresh summary ---
  refreshResumen();

  // --- Summary ---
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

  if (strategy === 'csv' || strategy === 'excel') {
    var transactions = parseBankFile(fileId, strategy, fileHash);
    var rows = transactions.map(function (tx) { return buildBankRow(tx, file.getName(), fileId); });
    flushReceiptRows(sheets.gastos, rows);
  } else {
    var base64 = getFileBase64(fileId);
    var data = extractReceiptData(base64, mimeType, file.getName(), fileHash);
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
