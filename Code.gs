/**
 * Code.gs
 * Orchestrator: menu setup, main processing loop, and utility entry points.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recibos')
    .addItem('Procesar Recibos', 'processReceipts')
    .addItem('Actualizar Resumen', 'updateResumen')
    .addSeparator()
    .addItem('Configurar API Key', 'setupApiKey')
    .addToUi();
}

/**
 * Main orchestrator. Processes receipt images and classifies items as keto/not-keto.
 */
function processReceipts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Sheet setup ---
  const sheets = initAllSheets();
  const ketoKeys = Object.keys(sheets.ketoDict);
  const geminiUrl = getGeminiUrl(); // cache URL (and API key) for the entire batch

  // --- 2. Dedup ---
  const processedIds = getProcessedFileIds(sheets.gastos);

  // --- 3. Fetch image list ---
  let allFiles;
  try {
    allFiles = getImageFiles(FOLDER_ID);
  } catch (driveError) {
    Logger.log('processReceipts: error accediendo a Drive: %s', driveError.message);
    ui.alert('Error', 'No se pudo leer la carpeta de Drive:\n' + driveError.message, ui.ButtonSet.OK);
    return;
  }

  // --- 4. Filter already processed ---
  const pendingFiles = allFiles.filter(function (f) {
    return !processedIds.has(f.id);
  });

  if (pendingFiles.length === 0) {
    ui.alert(
      'Sin archivos nuevos',
      'Todos los archivos de la carpeta ya han sido procesados.',
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 5. Respect batch size ---
  const hasMore = pendingFiles.length > BATCH_SIZE;
  const batch = pendingFiles.slice(0, BATCH_SIZE);

  Logger.log(
    'processReceipts: %s archivos pendientes, procesando %s en este lote.',
    pendingFiles.length,
    batch.length
  );

  ss.toast('Iniciando procesamiento de ' + batch.length + ' archivo(s)...', 'Recibos', 5);

  // --- 6. Processing loop (accumulate rows for batch write) ---
  let successCount = 0;
  let errorCount = 0;
  const gastosRows = [];
  const allKetoRows = [];
  const allNoKetoRows = [];

  for (let i = 0; i < batch.length; i++) {
    const file = batch[i];
    ss.toast('Procesando ' + (i + 1) + ' de ' + batch.length + ': ' + file.name, 'Recibos', 30);
    Logger.log('processReceipts: [%s/%s] procesando "%s"', i + 1, batch.length, file.name);

    try {
      const base64 = getImageBase64(file.id);
      const data = extractReceiptData(base64, file.mimeType, file.name, geminiUrl);

      if (data) {
        gastosRows.push(buildReceiptRow(data, file.name, file.id));
        const classified = classifyItems(data, sheets.ketoDict, ketoKeys);
        classified.ketoRows.forEach(function (r) { allKetoRows.push(r); });
        classified.noKetoRows.forEach(function (r) { allNoKetoRows.push(r); });
        successCount++;
      } else {
        Logger.log('processReceipts: extracción fallida para "%s", se omite.', file.name);
        errorCount++;
      }
    } catch (fileError) {
      Logger.log('processReceipts: error inesperado en "%s": %s', file.name, fileError.message);
      errorCount++;
    }

    if (i < batch.length - 1) {
      Utilities.sleep(RATE_LIMIT_DELAY);
    }
  }

  // --- 7. Batch write all accumulated rows ---
  flushReceiptRows(sheets.gastos, gastosRows);
  flushClassifiedRows(sheets.incluidos, sheets.excluidos, allKetoRows, allNoKetoRows);

  // --- 8. Refresh summary ---
  refreshResumen();

  // --- 9. Summary toast ---
  const summaryLines = [
    'Procesamiento completado.',
    '',
    'Exitosos:  ' + successCount,
    'Con error: ' + errorCount,
    'Total procesados en este lote: ' + batch.length
  ];

  if (hasMore) {
    const remaining = pendingFiles.length - batch.length;
    summaryLines.push('');
    summaryLines.push(
      'Quedan ' + remaining + ' archivo(s) pendientes. Ejecuta "Procesar Recibos" nuevamente.'
    );
  }

  ss.toast(summaryLines.join('\n'), 'Recibos', 15);
  Logger.log('processReceipts: finalizado. Exitosos: %s, Errores: %s.', successCount, errorCount);
}

function updateResumen() {
  refreshResumen();
  SpreadsheetApp.getActiveSpreadsheet().toast('Resumen actualizado.', 'Recibos', 5);
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
  if (!fileId) {
    Logger.log('processSpecificFile: fileId requerido.');
    return;
  }

  const sheets = initAllSheets();
  const ketoKeys = Object.keys(sheets.ketoDict);

  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    Logger.log('processSpecificFile: archivo no encontrado: %s', fileId);
    return;
  }

  const mimeType = file.getMimeType();
  if (SUPPORTED_MIME_TYPES.indexOf(mimeType) === -1) {
    Logger.log('processSpecificFile: tipo MIME no soportado "%s" para "%s".', mimeType, file.getName());
    return;
  }

  const base64 = getImageBase64(fileId);
  const data = extractReceiptData(base64, mimeType, file.getName());

  if (data) {
    flushReceiptRows(sheets.gastos, [buildReceiptRow(data, file.getName(), fileId)]);
    const classified = classifyItems(data, sheets.ketoDict, ketoKeys);
    flushClassifiedRows(sheets.incluidos, sheets.excluidos, classified.ketoRows, classified.noKetoRows);
    refreshResumen();
    Logger.log('processSpecificFile: fila añadida para "%s".', file.getName());
  } else {
    Logger.log('processSpecificFile: extracción fallida para "%s".', file.getName());
  }
}
