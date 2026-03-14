/**
 * DriveService.gs
 * Utilities for reading image files from Google Drive.
 */

/**
 * Returns all supported image files inside the given Drive folder.
 *
 * @param {string} folderId - The Drive folder ID to scan.
 * @returns {{ id: string, name: string, mimeType: string }[]} Array of file descriptors.
 */
function getImageFiles(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const result = [];

  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();

    if (SUPPORTED_MIME_TYPES.indexOf(mimeType) !== -1) {
      result.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: mimeType
      });
    }
  }

  Logger.log('DriveService: encontrados %s archivos de imagen en la carpeta.', result.length);
  return result;
}

/**
 * Downloads a Drive file and encodes its contents as a base64 string.
 *
 * @param {string} fileId - The Drive file ID.
 * @returns {string} Base64-encoded file content.
 */
function getImageBase64(fileId) {
  return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes());
}
