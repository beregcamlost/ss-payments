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

  // --- 1. Sheet setup ---
  const sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  const incluidosSheet = getOrCreateIncluidosSheet();
  ensureIncluidosHeaders(incluidosSheet);

  const excluidosSheet = getOrCreateExcluidosSheet();
  ensureExcluidosHeaders(excluidosSheet);

  const ketoSheet = getOrCreateKetoDict();
  ensureKetoDictHeaders(ketoSheet);
  ensureKetoDictData(ketoSheet);

  const ketoDict = loadKetoDictionary(ketoSheet);

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
        appendClassifiedItems(incluidosSheet, excluidosSheet, data, file.id, ketoDict);
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

    if (i < batch.length - 1) {
      Utilities.sleep(RATE_LIMIT_DELAY);
    }
  }

  // --- 7. Refresh summary ---
  refreshResumen();

  // --- 8. Summary toast ---
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
 * Menu entry point: manually refresh the Resumen tab from existing data.
 */
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

  const sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  const incluidosSheet = getOrCreateIncluidosSheet();
  ensureIncluidosHeaders(incluidosSheet);
  const excluidosSheet = getOrCreateExcluidosSheet();
  ensureExcluidosHeaders(excluidosSheet);

  const ketoSheet = getOrCreateKetoDict();
  ensureKetoDictHeaders(ketoSheet);
  ensureKetoDictData(ketoSheet);
  const ketoDict = loadKetoDictionary(ketoSheet);

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
      'processSpecificFile: tipo MIME no soportado "%s" para "%s".',
      mimeType,
      file.getName()
    );
    return;
  }

  const base64 = getImageBase64(fileId);
  const data = extractReceiptData(base64, mimeType, file.getName());

  if (data) {
    appendReceiptRow(sheet, data, file.getName(), fileId);
    appendClassifiedItems(incluidosSheet, excluidosSheet, data, fileId, ketoDict);
    refreshResumen();
    Logger.log('processSpecificFile: fila añadida para "%s".', file.getName());
  } else {
    Logger.log('processSpecificFile: extracción fallida para "%s".', file.getName());
  }
}
