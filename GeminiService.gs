/**
 * GeminiService.gs
 * Calls the Gemini API to extract structured data from receipt screenshots.
 */

/**
 * System instruction sent to Gemini explaining the receipt domain.
 * Written in Spanish to match the receipts' language.
 * @const {string}
 */
const SYSTEM_INSTRUCTION = `Eres un asistente experto en extraer datos de comprobantes financieros chilenos.
Los documentos que recibirás son capturas de pantalla de WhatsApp y pueden ser de dos tipos:

1. BOLETAS ELECTRÓNICAS: Emitidas por comercios. Contienen nombre del comercio, RUT del emisor,
   fecha de emisión, listado de productos/servicios, monto neto, IVA (19%) y total.
   El medio de pago puede ser efectivo, débito, crédito u otro.

2. COMPROBANTES DE TRANSFERENCIA BANCARIA:
   - BCI: Pantalla titulada "Operación exitosa". Muestra monto, destinatario, banco destino,
     número de operación y fecha.
   - Scotiabank: Encabezado "Comprobante de Transferencia". Muestra monto, destinatario,
     banco destino, número de operación y fecha.
   Para transferencias, el campo "neto" e "iva" van vacíos; solo se llena "total".

Reglas generales:
- Las fechas deben estar en formato YYYY-MM-DD.
- Todos los montos son en pesos chilenos (CLP), sin decimales y sin separadores de miles
  (ni puntos ni comas). Ejemplo correcto: 15990. Ejemplo incorrecto: 15.990 o 15,990.
- Si un campo no aplica para el tipo de documento, devuelve cadena vacía "".
- Para "categoria", elige la opción más apropiada de esta lista exacta:
  ${CATEGORIES.join(', ')}.
- Para "descripcion", resume brevemente los artículos comprados (boleta) o el concepto
  de la transferencia (máximo 100 caracteres).
- Para "tipo", devuelve exactamente "boleta" o "transferencia".
- Para "items", extrae cada línea de detalle de la boleta como un objeto con nombre,
  cantidad y precio_unitario. Si es una transferencia o no hay detalle de items, devuelve
  un arreglo vacío [].`;

/**
 * JSON schema for the structured Gemini response.
 * All monetary fields are strings to avoid floating-point issues; they will be
 * parsed to integers in SheetService.
 * @const {object}
 */
const RESPONSE_SCHEMA = {
  type: 'object',
  properties: {
    fecha: { type: 'string', description: 'Fecha del comprobante en formato YYYY-MM-DD' },
    tipo: { type: 'string', description: 'Tipo de documento: boleta o transferencia' },
    comercio_destinatario: { type: 'string', description: 'Nombre del comercio o destinatario' },
    rut: { type: 'string', description: 'RUT del comercio o destinatario' },
    categoria: { type: 'string', description: 'Categoría del gasto' },
    descripcion: { type: 'string', description: 'Descripción breve del gasto o transferencia' },
    neto: { type: 'string', description: 'Monto neto en CLP, sin separadores' },
    iva: { type: 'string', description: 'IVA en CLP, sin separadores' },
    total: { type: 'string', description: 'Monto total en CLP, sin separadores' },
    medio_pago: { type: 'string', description: 'Medio de pago (efectivo, débito, crédito, transferencia, etc.)' },
    banco_destino: { type: 'string', description: 'Banco destino en caso de transferencia' },
    numero_operacion: { type: 'string', description: 'Número de operación o folio' },
    items: {
      type: 'array',
      description: 'Detalle de productos/servicios de la boleta. Vacío para transferencias.',
      items: {
        type: 'object',
        properties: {
          nombre: { type: 'string', description: 'Nombre del producto o servicio' },
          cantidad: { type: 'number', description: 'Cantidad comprada' },
          precio_unitario: { type: 'number', description: 'Precio unitario en CLP sin separadores' }
        },
        required: ['nombre', 'cantidad', 'precio_unitario']
      }
    }
  },
  required: [
    'fecha', 'tipo', 'comercio_destinatario', 'rut', 'categoria', 'descripcion',
    'neto', 'iva', 'total', 'medio_pago', 'banco_destino', 'numero_operacion', 'items'
  ]
};

/**
 * Sends a receipt image to the Gemini API and returns the extracted structured data.
 *
 * @param {string} base64    - Base64-encoded image content.
 * @param {string} mimeType  - MIME type of the image (e.g. 'image/jpeg').
 * @param {string} fileName  - Original file name, used only for logging.
 * @returns {object|null} Parsed JSON object with receipt fields, or null on failure.
 */
function extractReceiptData(base64, mimeType, fileName) {
  Logger.log('GeminiService: procesando "%s"', fileName);

  const payload = {
    system_instruction: {
      parts: [{ text: SYSTEM_INSTRUCTION }]
    },
    contents: [
      {
        role: 'user',
        parts: [
          {
            inline_data: {
              mime_type: mimeType,
              data: base64
            }
          },
          {
            text: `Analiza este comprobante y extrae los datos solicitados en formato JSON.
Recuerda:
- fecha en YYYY-MM-DD
- tipo: "boleta" o "transferencia"
- montos en CLP entero sin separadores (ej: 15990)
- categoria debe ser una de: ${CATEGORIES.join(', ')}
- campos que no apliquen: cadena vacía ""`
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: RESPONSE_SCHEMA,
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(getGeminiUrl(), options);
  } catch (networkError) {
    Logger.log('GeminiService: error de red para "%s": %s', fileName, networkError.message);
    return null;
  }

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode !== 200) {
    Logger.log(
      'GeminiService: respuesta HTTP %s para "%s". Body: %s',
      statusCode,
      fileName,
      responseText
    );
    return null;
  }

  let parsed;
  try {
    parsed = JSON.parse(responseText);
  } catch (parseError) {
    Logger.log('GeminiService: no se pudo parsear la respuesta para "%s".', fileName);
    return null;
  }

  // Navigate to the text content inside the Gemini response envelope
  const candidates = parsed.candidates;
  if (!candidates || candidates.length === 0) {
    Logger.log('GeminiService: sin candidatos en la respuesta para "%s".', fileName);
    return null;
  }

  const textContent = candidates[0].content &&
    candidates[0].content.parts &&
    candidates[0].content.parts[0] &&
    candidates[0].content.parts[0].text;

  if (!textContent) {
    Logger.log('GeminiService: contenido vacío en la respuesta para "%s".', fileName);
    return null;
  }

  let extracted;
  try {
    extracted = JSON.parse(textContent);
  } catch (jsonError) {
    Logger.log(
      'GeminiService: JSON inválido en contenido para "%s": %s',
      fileName,
      textContent
    );
    return null;
  }

  Logger.log('GeminiService: extracción exitosa para "%s".', fileName);
  return extracted;
}
