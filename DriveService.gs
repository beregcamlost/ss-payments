/**
 * DriveService.gs
 * Utilities for reading files from Google Drive.
 */

/**
 * Returns all supported files inside the given Drive folder,
 * tagged with a processing strategy.
 *
 * @param {string} folderId - The Drive folder ID to scan.
 * @returns {{ id: string, name: string, mimeType: string, strategy: string }[]}
 */
function getFiles(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const result = [];

  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    const strategy = FILE_STRATEGIES[mimeType];

    if (strategy) {
      result.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: mimeType,
        strategy: strategy
      });
    }
  }

  Logger.log('DriveService: encontrados %s archivos soportados en la carpeta.', result.length);
  return result;
}

/**
 * Downloads a Drive file and encodes its contents as a base64 string.
 *
 * @param {string} fileId - The Drive file ID.
 * @returns {string} Base64-encoded file content.
 */
function getFileBase64(fileId) {
  return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes());
}

/**
 * Computes a cache key from a Drive file's metadata.
 * Uses fileId + lastUpdated to avoid loading file content into memory.
 *
 * @param {string} fileId - The Drive file ID.
 * @returns {string} Cache hash key.
 */
function getDriveFileHash(fileId) {
  var file = DriveApp.getFileById(fileId);
  return file.getId() + ':' + file.getLastUpdated().getTime();
}

/**
 * Converts an Excel file to row arrays by uploading as a Google Sheet.
 *
 * @param {string} fileId - Drive file ID of the Excel file.
 * @returns {string[][]} Array of row arrays (including header row).
 */
function excelToRows(fileId) {
  var excelFile = DriveApp.getFileById(fileId);
  var blob = excelFile.getBlob();

  var resource = { title: 'temp_' + excelFile.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  var tempFile = Drive.Files.insert(resource, blob, { convert: true });

  try {
    var tempSheet = SpreadsheetApp.openById(tempFile.id).getSheets()[0];
    var lastRow = tempSheet.getLastRow();
    var lastCol = tempSheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return [];
    var data = tempSheet.getRange(1, 1, lastRow, lastCol).getValues();
    return data.map(function (row) {
      return row.map(function (cell) { return String(cell); });
    });
  } finally {
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    Logger.log('DriveService: archivo temporal Excel eliminado.');
  }
}

/**
 * Reads a CSV file from Drive and returns parsed rows.
 *
 * @param {string} fileId - Drive file ID of the CSV file.
 * @returns {string[][]} Array of row arrays.
 */
function csvToRows(fileId) {
  var content = DriveApp.getFileById(fileId).getBlob().getDataAsString();
  return Utilities.parseCsv(content);
}
