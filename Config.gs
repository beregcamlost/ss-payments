/**
 * Config.gs
 * Centralized configuration and constants for the WhatsApp receipt extractor.
 */

/** @const {string} Google Drive folder ID containing receipt screenshots */
const FOLDER_ID = '1uTLNz0tl3lddwskVeE99x1eq-_JyDkki';

/** @const {string} Target sheet name */
const SHEET_NAME = 'Gastos';

/** @const {string} Gemini model to use for extraction */
const GEMINI_MODEL = 'gemini-2.5-flash';

/** @const {string} Base endpoint template (without API key) */
const GEMINI_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

/** @const {string[]} MIME types accepted for processing */
const SUPPORTED_MIME_TYPES = ['image/jpeg', 'image/png', 'image/webp'];

/** @const {string[]} Expense categories for classification */
const CATEGORIES = [
  'Alimentacion',
  'Farmacia',
  'Vestuario',
  'Hogar',
  'Transferencia Personal',
  'Arriendo',
  'Servicios',
  'Transporte',
  'Otro'
];

/**
 * Column headers for the Gastos sheet (15 columns).
 * Order must match the row array built in SheetService.appendReceiptRow().
 * @const {string[]}
 */
const HEADERS = [
  'Fecha',               // A - YYYY-MM-DD
  'Tipo',                // B - boleta | transferencia
  'Comercio / Destinatario', // C
  'RUT',                 // D
  'Categoria',           // E
  'Descripcion',         // F
  'Neto',                // G - integer CLP
  'IVA',                 // H - integer CLP
  'Total',               // I - integer CLP
  'Medio de Pago',       // J
  'Banco Destino',       // K
  'N° Operacion',        // L
  'Archivo',             // M - original file name
  'File ID',             // N - Drive file ID (used for dedup)
  'Procesado'            // O - processing timestamp
];

/**
 * Milliseconds to wait between Gemini API calls to avoid rate-limit errors.
 * @const {number}
 */
const RATE_LIMIT_DELAY = 4500;

/**
 * Maximum number of files to process in a single run to stay under
 * the 6-minute Apps Script execution limit.
 * @const {number}
 */
const BATCH_SIZE = 70;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Retrieves the Gemini API key from Script Properties.
 * @returns {string} The API key.
 * @throws {Error} If the property is not set.
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error(
      'API key no configurada. Usa el menú Recibos > Configurar API Key para guardarla.'
    );
  }
  return key;
}

/**
 * Builds the full Gemini API URL including the API key query parameter.
 * @returns {string} Complete request URL.
 */
function getGeminiUrl() {
  return `${GEMINI_ENDPOINT}?key=${getApiKey()}`;
}
