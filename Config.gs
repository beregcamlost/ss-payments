/**
 * Config.gs
 * Centralized configuration and constants for the WhatsApp receipt extractor.
 */

/** @const {string} Google Drive folder ID containing receipt screenshots */
const FOLDER_ID = '1uTLNz0tl3lddwskVeE99x1eq-_JyDkki';

/** @const {string} Gemini model to use for extraction */
const GEMINI_MODEL = 'gemini-2.5-flash';

/** @const {string} Base endpoint template (without API key) */
const GEMINI_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

/** @const {string[]} MIME types accepted for processing */
const SUPPORTED_MIME_TYPES = ['image/jpeg', 'image/png', 'image/webp'];

// ---------------------------------------------------------------------------
// Sheet names
// ---------------------------------------------------------------------------
const SHEET_NAME = 'Gastos';
const INCLUIDOS_SHEET_NAME = 'Incluidos';
const EXCLUIDOS_SHEET_NAME = 'Excluidos';
const RESUMEN_SHEET_NAME = 'Resumen';
const KETO_DICT_SHEET_NAME = 'Diccionario Keto';

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

// ---------------------------------------------------------------------------
// Column headers
// ---------------------------------------------------------------------------

const HEADERS = [
  'Fecha',               // A
  'Tipo',                // B
  'Comercio / Destinatario', // C
  'Categoria',           // D
  'Descripcion',         // E
  'Total',               // F
  'Banco Destino',       // G
  'Archivo',             // H
  'File ID',             // I
  'Procesado'            // J
];

const CLASSIFIED_ITEM_HEADERS = [
  'Fecha',           // A
  'Comercio',        // B
  'Categoria',       // C
  'Item',            // D
  'Cantidad',        // E
  'Precio Unitario', // F
  'Subtotal',        // G
  'File ID'          // H
];

const KETO_DICT_HEADERS = ['Palabra Clave', 'Keto'];

const RESUMEN_HEADERS = [
  'Categoria',
  'Gasto Keto',
  'Gasto No Keto',
  'Total'
];

// ---------------------------------------------------------------------------
// Default keto dictionary (pre-populated on first run)
// ---------------------------------------------------------------------------
const DEFAULT_KETO_DICT = [
  // --- Keto-friendly (SI) ---
  ['pollo', 'SI'], ['carne', 'SI'], ['cerdo', 'SI'], ['salmon', 'SI'],
  ['atun', 'SI'], ['pescado', 'SI'], ['marisco', 'SI'], ['camaron', 'SI'],
  ['huevo', 'SI'], ['jamon', 'SI'], ['tocino', 'SI'], ['chorizo', 'SI'],
  ['salchicha', 'SI'], ['costilla', 'SI'], ['lomo', 'SI'], ['filete', 'SI'],
  ['pechuga', 'SI'], ['res', 'SI'], ['vacuno', 'SI'], ['cordero', 'SI'],
  ['queso', 'SI'], ['mantequilla', 'SI'], ['crema', 'SI'], ['aceite', 'SI'],
  ['manteca', 'SI'], ['mayonesa', 'SI'], ['nata', 'SI'],
  ['lechuga', 'SI'], ['espinaca', 'SI'], ['brocoli', 'SI'], ['coliflor', 'SI'],
  ['pepino', 'SI'], ['apio', 'SI'], ['zapallo italiano', 'SI'], ['repollo', 'SI'],
  ['acelga', 'SI'], ['rucula', 'SI'], ['champiñon', 'SI'], ['hongos', 'SI'],
  ['palta', 'SI'], ['aceituna', 'SI'], ['aguacate', 'SI'],
  ['almendra', 'SI'], ['nuez', 'SI'], ['mani', 'SI'], ['semilla', 'SI'],
  ['tomate', 'SI'], ['cebolla', 'SI'], ['ajo', 'SI'], ['pimenton', 'SI'],
  ['agua mineral', 'SI'], ['cafe', 'SI'], ['te', 'SI'],
  // --- Not keto (NO) ---
  ['pan', 'NO'], ['arroz', 'NO'], ['fideos', 'NO'], ['pasta', 'NO'],
  ['harina', 'NO'], ['cereal', 'NO'], ['avena', 'NO'], ['maiz', 'NO'],
  ['tortilla', 'NO'], ['galleta', 'NO'], ['galletones', 'NO'],
  ['azucar', 'NO'], ['chocolate', 'NO'], ['dulce', 'NO'], ['mermelada', 'NO'],
  ['miel', 'NO'], ['torta', 'NO'], ['helado', 'NO'], ['caramelo', 'NO'],
  ['postre', 'NO'], ['kuchen', 'NO'],
  ['bebida', 'NO'], ['gaseosa', 'NO'], ['jugo', 'NO'], ['nectar', 'NO'],
  ['cerveza', 'NO'],
  ['papa', 'NO'], ['pure', 'NO'], ['choclo', 'NO'],
  ['empanada', 'NO'], ['completo', 'NO'], ['pizza', 'NO'], ['sandwich', 'NO'],
  ['sushi', 'NO'], ['hamburguesa', 'NO'],
  ['platano', 'NO'], ['uva', 'NO'], ['naranja', 'NO'],
  ['poroto', 'NO'], ['lenteja', 'NO'], ['garbanzo', 'NO']
];

// ---------------------------------------------------------------------------
// Processing limits
// ---------------------------------------------------------------------------

const RATE_LIMIT_DELAY = 4500;
const BATCH_SIZE = 70;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error(
      'API key no configurada. Usa el menú Recibos > Configurar API Key para guardarla.'
    );
  }
  return key;
}

function getGeminiUrl() {
  return `${GEMINI_ENDPOINT}?key=${getApiKey()}`;
}
