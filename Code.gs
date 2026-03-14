/**
 * Code.gs
 * Orchestrator: menu setup, main processing loop, and utility entry points.
 */

/**
 * Adds the custom "Recibos" menu to the spreadsheet UI when the file is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recibos')
    .addItem('Procesar Recibos', 'processReceipts')
    .addSeparator()
    .addItem('Configurar API Key', 'setupApiKey')
    .addToUi();
}

/**
 * Main orchestrator. Reads new images from Drive, extracts data via Gemini,
 * and writes each result to the Gastos sheet.
 *
 * Flow:
 *  1. Prepare sheet and headers.
 *  2. Build dedup set of already-processed file IDs.
 *  3. Fetch all supported images from the Drive folder.
 *  4. Filter out already-processed files.
 *  5. Respect BATCH_SIZE to avoid the 6-minute execution limit.
 *  6. For each file: download → extract → write row → delay.
 *  7. Show completion summary.
 */
function processReceipts() {
  const ui = SpreadsheetApp.getUi();

  // --- 1. Sheet setup ---
  const sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  // --- 2. Dedup ---
  const processedIds = getProcessedFileIds(sheet);

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

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Iniciando procesamiento de ' + batch.length + ' archivo(s)...',
    'Recibos',
    5
  );

  // --- 6. Processing loop ---
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < batch.length; i++) {
    const file = batch[i];
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Procesando ' + (i + 1) + ' de ' + batch.length + ': ' + file.name,
      'Recibos',
      30
    );
    Logger.log(
      'processReceipts: [%s/%s] procesando "%s"',
      i + 1,
      batch.length,
      file.name
    );

    try {
      const base64 = getImageBase64(file.id);
      const data = extractReceiptData(base64, file.mimeType, file.name);

      if (data) {
        appendReceiptRow(sheet, data, file.name, file.id);
        successCount++;
      } else {
        Logger.log('processReceipts: extracción fallida para "%s", se omite.', file.name);
        errorCount++;
      }
    } catch (fileError) {
      Logger.log(
        'processReceipts: error inesperado en "%s": %s',
        file.name,
        fileError.message
      );
      errorCount++;
    }

    // Delay between API calls to respect rate limits (skip delay after last file)
    if (i < batch.length - 1) {
      Utilities.sleep(RATE_LIMIT_DELAY);
    }
  }

  // --- 7. Summary ---
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

  SpreadsheetApp.getActiveSpreadsheet().toast(
    summaryLines.join('\n'),
    'Recibos',
    15
  );

  Logger.log('processReceipts: finalizado. Exitosos: %s, Errores: %s.', successCount, errorCount);
}

/**
 * Prompts the user to enter their Gemini API key and stores it in
 * Script Properties (not visible in code, not committed to source).
 */
function setupApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configurar API Key',
    'Ingresa tu clave de API de Gemini (Google AI Studio):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const key = result.getResponseText().trim();
  if (!key) {
    ui.alert('Error', 'La clave no puede estar vacía.', ui.ButtonSet.OK);
    return;
  }

  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('Listo', 'API Key guardada correctamente.', ui.ButtonSet.OK);
  Logger.log('setupApiKey: API key guardada en Script Properties.');
}

/**
 * Reprocesses a single Drive file by ID. Useful for debugging extraction
 * issues with a specific screenshot without running the full batch.
 *
 * @param {string} fileId - The Drive file ID to reprocess.
 */
function processSpecificFile(fileId) {
  if (!fileId) {
    Logger.log('processSpecificFile: fileId requerido.');
    return;
  }

  const sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    Logger.log('processSpecificFile: archivo no encontrado: %s', fileId);
    return;
  }

  const mimeType = file.getMimeType();
  if (SUPPORTED_MIME_TYPES.indexOf(mimeType) === -1) {
    Logger.log(
      'processSpecificFile: tipo MIME no soportado "%s" para el archivo "%s".',
      mimeType,
      file.getName()
    );
    return;
  }

  const base64 = getImageBase64(fileId);
  const data = extractReceiptData(base64, mimeType, file.getName());

  if (data) {
    appendReceiptRow(sheet, data, file.getName(), fileId);
    Logger.log('processSpecificFile: fila añadida para "%s".', file.getName());
  } else {
    Logger.log('processSpecificFile: extracción fallida para "%s".', file.getName());
  }
}
