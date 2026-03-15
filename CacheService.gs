/**
 * CacheService.gs
 * Caches Gemini API responses to avoid redundant calls.
 * Uses a hidden "Cache" sheet keyed by file metadata hash.
 */

const CACHE_HEADERS = ['Hash', 'Tipo', 'Resultado', 'Timestamp'];

/**
 * Initializes the Cache sheet (hidden).
 */
function initCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CACHE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CACHE_SHEET_NAME);
    sheet.hideSheet();
    const range = sheet.getRange(1, 1, 1, CACHE_HEADERS.length);
    range.setValues([CACHE_HEADERS]);
    range.setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('CacheService: hoja Cache creada y oculta.');
  }
  return sheet;
}

/**
 * Looks up a cached Gemini response.
 * @param {string} hash - Cache key (from getDriveFileHash).
 * @param {string} tipo - Cache entry type (receipt, csv_columns, etc.).
 * @returns {object|null} Parsed JSON result, or null if not cached.
 */
function getCached(hash, tipo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CACHE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CACHE_HEADERS.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === hash && data[i][1] === tipo) {
      try {
        return JSON.parse(data[i][2]);
      } catch (e) {
        Logger.log('CacheService: entrada corrupta para hash "%s", tipo "%s".', hash, tipo);
        return null;
      }
    }
  }
  return null;
}

/**
 * Stores a Gemini response in the cache.
 */
function putCache(hash, tipo, result) {
  var sheet = initCacheSheet();
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, CACHE_HEADERS.length)
    .setValues([[hash, tipo, JSON.stringify(result), new Date()]]);
}

/**
 * Cache-aware wrapper around callGemini().
 * Checks cache first; on miss, calls Gemini and stores the result.
 *
 * @param {string} hash - Cache key.
 * @param {string} tipo - Cache entry type.
 * @param {object} payload - Gemini API request body.
 * @param {string} label - Logging label.
 * @returns {object|null} Parsed result from cache or Gemini.
 */
function cachedCallGemini(hash, tipo, payload, label) {
  var cached = getCached(hash, tipo);
  if (cached) {
    Logger.log('CacheService: cache hit para "%s" (%s).', label, tipo);
    return cached;
  }

  var result = callGemini(payload, label);
  if (result) {
    putCache(hash, tipo, result);
    Logger.log('CacheService: resultado cacheado para "%s" (%s).', label, tipo);
  }
  return result;
}

/**
 * Clears all cached entries.
 */
function clearCache() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CACHE_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  Logger.log('CacheService: cache limpiado.');
}
