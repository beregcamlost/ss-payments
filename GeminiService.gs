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
 * Sends a payload to the Gemini API and returns the parsed JSON response.
 * Handles fetch, HTTP errors, and Gemini response envelope unwrapping.
 *
 * @param {object} payload  - Full Gemini API request body.
 * @param {string} label    - Label for logging (e.g. file name or operation).
 * @param {string} [url]    - Optional cached Gemini URL.
 * @returns {object|null} Parsed JSON from the response, or null on failure.
 */
function callGemini(payload, label, url) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response;
  try {
    response = UrlFetchApp.fetch(url || getGeminiUrl(), options);
  } catch (networkError) {
    Logger.log('GeminiService: error de red para "%s": %s', label, networkError.message);
    return null;
  }

  if (response.getResponseCode() !== 200) {
    Logger.log('GeminiService: HTTP %s para "%s": %s', response.getResponseCode(), label, response.getContentText());
    return null;
  }

  var envelope;
  try {
    envelope = JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('GeminiService: no se pudo parsear la respuesta para "%s".', label);
    return null;
  }

  var candidates = envelope.candidates;
  if (!candidates || candidates.length === 0) {
    Logger.log('GeminiService: sin candidatos en la respuesta para "%s".', label);
    return null;
  }

  var textContent = candidates[0].content &&
    candidates[0].content.parts &&
    candidates[0].content.parts[0] &&
    candidates[0].content.parts[0].text;

  if (!textContent) {
    Logger.log('GeminiService: contenido vacío en la respuesta para "%s".', label);
    return null;
  }

  try {
    return JSON.parse(textContent);
  } catch (e) {
    Logger.log('GeminiService: JSON inválido en contenido para "%s": %s', label, textContent);
    return null;
  }
}

/**
 * Sends a receipt image to the Gemini API and returns the extracted structured data.
 */
function extractReceiptData(base64, mimeType, fileName, fileHash, cachedUrl) {
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

  const result = cachedCallGemini(fileHash || fileName, 'receipt', payload, fileName);
  if (result) Logger.log('GeminiService: extracción exitosa para "%s".', fileName);
  return result;
}

/**
 * Sends the first rows of a CSV/Excel to Gemini to auto-detect column mapping.
 *
 * @param {string[][]} sampleRows - First 5 rows of the file.
 * @param {string} hash - Cache key for the file.
 * @returns {object|null} Column mapping: { fecha_col, descripcion_col, monto_col, banco, date_format, header_row }
 */
function detectCsvColumns(sampleRows, hash) {
  var sample = sampleRows.slice(0, 5).map(function (row) { return row.join(' | '); }).join('\n');

  var payload = {
    system_instruction: {
      parts: [{ text: 'Eres un experto en extractos bancarios chilenos. Identificas columnas y formatos de archivos CSV/Excel de bancos.' }]
    },
    contents: [{
      role: 'user',
      parts: [{ text: 'Identifica las columnas de este extracto bancario chileno. Las primeras filas son:\n\n' + sample +
        '\n\nIdentifica que columna (indice base 0) contiene: fecha, descripcion/glosa, monto/cargo, y el nombre del banco si se puede inferir del contenido.' }]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'object',
        properties: {
          fecha_col: { type: 'number', description: 'Indice de la columna de fecha (base 0)' },
          descripcion_col: { type: 'number', description: 'Indice de la columna de descripcion/glosa' },
          monto_col: { type: 'number', description: 'Indice de la columna de monto/cargo' },
          banco: { type: 'string', description: 'Nombre del banco inferido del contenido' },
          date_format: { type: 'string', description: 'Formato de fecha detectado (ej: DD/MM/YYYY, YYYY-MM-DD)' },
          header_row: { type: 'number', description: 'Indice de la fila de encabezados (base 0), -1 si no hay' },
          moneda: { type: 'string', description: 'Moneda (CLP, USD, etc.)' }
        },
        required: ['fecha_col', 'descripcion_col', 'monto_col', 'banco', 'date_format', 'header_row']
      },
      temperature: 0.1
    }
  };

  return cachedCallGemini(hash, 'csv_columns', payload, 'deteccion columnas CSV');
}

/**
 * Batch-categorizes bank transaction descriptions.
 *
 * @param {string[]} descriptions - Array of transaction descriptions.
 * @param {string} hash - Cache key for the batch.
 * @returns {object|null} { items: [{ descripcion, categoria, comercio }] }
 */
function categorizeBankDescriptions(descriptions, hash) {
  var unique = Array.from(new Set(descriptions));

  var payload = {
    system_instruction: {
      parts: [{ text: 'Eres un experto en finanzas personales chilenas. Categorizas transacciones bancarias.' }]
    },
    contents: [{
      role: 'user',
      parts: [{ text: 'Categoriza cada transaccion bancaria. Categorias validas: ' + CATEGORIES.join(', ') +
        '\n\nTransacciones:\n' + unique.map(function (d, i) { return (i + 1) + '. ' + d; }).join('\n') }]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'object',
        properties: {
          items: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                descripcion: { type: 'string', description: 'Descripcion original' },
                categoria: { type: 'string', description: 'Categoria asignada' },
                comercio: { type: 'string', description: 'Nombre normalizado del comercio' }
              },
              required: ['descripcion', 'categoria', 'comercio']
            }
          }
        },
        required: ['items']
      },
      temperature: 0.1
    }
  };

  return cachedCallGemini(hash, 'csv_categories', payload, 'categorizacion banco');
}

/**
 * Asks Gemini to match a receipt to candidate bank transactions.
 *
 * @param {object} receiptData - Extracted receipt data (comercio, fecha, total).
 * @param {object[]} candidates - Candidate bank rows [{fecha, descripcion, monto}].
 * @param {string} hash - Cache key.
 * @returns {object|null} { matched: true/false, index: candidateIndex }
 */
function matchReceiptToTransaction(receiptData, candidates, hash) {
  var payload = {
    system_instruction: {
      parts: [{ text: 'Eres un experto en conciliacion bancaria chilena. Determinas si un comprobante de compra corresponde a una transaccion bancaria.' }]
    },
    contents: [{
      role: 'user',
      parts: [{ text: 'Este comprobante es de:\n' +
        '- Comercio: ' + receiptData.comercio + '\n' +
        '- Fecha: ' + receiptData.fecha + '\n' +
        '- Total: $' + receiptData.total + '\n\n' +
        'Corresponde a alguna de estas transacciones bancarias?\n' +
        candidates.map(function (c, i) {
          return (i + 1) + '. Fecha: ' + c.fecha + ' | Descripcion: ' + c.descripcion + ' | Monto: $' + c.monto;
        }).join('\n') +
        '\n\nSi no hay coincidencia, responde matched=false.' }]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'object',
        properties: {
          matched: { type: 'boolean', description: 'true si el comprobante corresponde a alguna transaccion' },
          index: { type: 'number', description: 'Indice (base 0) de la transaccion que coincide, -1 si no hay' },
          confidence: { type: 'string', description: 'alta, media, o baja' }
        },
        required: ['matched', 'index']
      },
      temperature: 0.1
    }
  };

  return cachedCallGemini(hash, 'matching', payload, 'matching recibo-banco');
}
