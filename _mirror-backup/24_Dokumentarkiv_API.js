/**
 * @file 24_Dokumentarkiv_API.js
 * @description Backend API for the Document Archive feature. Handles file uploads, listing, and deletion in Google Drive.
 */

// Helper function to get or create the main folder for the document archive
function getOrCreateDocumentArchiveFolder_() {
  const driveFolderName = "Dokumentarkiv - Sameiet";
  let folder;

  // Check if the folder already exists
  const folders = DriveApp.getFoldersByName(driveFolderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    // If not, create it
    folder = DriveApp.createFolder(driveFolderName);
  }
  return folder;
}

/**
 * Uploads a file to the designated Google Drive folder.
 * This function is called from the client-side script.
 *
 * @param {object} fileData - The file data from the client.
 * @param {string} fileData.fileName - The name of the file.
 * @param {string} fileData.mimeType - The MIME type of the file.
 * @param {string} fileData.data - The base64 encoded file data.
 * @returns {object} A result object indicating success or failure.
 */
function uploadFileToDrive(fileData) {
  try {
    const { fileName, mimeType, data } = fileData;
    const decodedData = Utilities.base64Decode(data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);

    const folder = getOrCreateDocumentArchiveFolder_();
    folder.createFile(blob);

    return { ok: true, message: 'File uploaded successfully.' };
  } catch (e) {
    console.error('Upload failed: ' + e.toString());
    return { ok: false, message: e.toString() };
  }
}

/**
 * Lists all files within the document archive folder.
 *
 * @returns {Array<object>} An array of file objects with details.
 */
function listFilesInDriveFolder() {
  try {
    const folder = getOrCreateDocumentArchiveFolder_();
    const files = folder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        mimeType: file.getMimeType(),
        size: file.getSize(),
        createdDate: file.getDateCreated().toISOString(),
      });
    }

    // Sort files by creation date, newest first
    fileList.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));

    return fileList;
  } catch (e) {
    console.error('Failed to list files: ' + e.toString());
    return []; // Return empty list on error
  }
}

/**
 * Deletes a file from the document archive folder by its ID.
 *
 * @param {string} fileId The ID of the file to delete.
 * @returns {object} A result object indicating success or failure.
 */
function deleteFileFromDrive(fileId) {
    try {
        const file = DriveApp.getFileById(fileId);
        // To be safe, let's check if the file is in our archive folder
        const folder = getOrCreateDocumentArchiveFolder_();
        const parents = file.getParents();
        let inFolder = false;
        while(parents.hasNext()){
            if(parents.next().getId() === folder.getId()){
                inFolder = true;
                break;
            }
        }

        if (!inFolder) {
            return { ok: false, message: "File is not in the archive folder." };
        }

        file.setTrashed(true); // Move to trash instead of permanent delete
        return { ok: true, message: "File deleted successfully." };
    } catch (e) {
        console.error('Failed to delete file: ' + e.toString());
        return { ok: false, message: e.toString() };
    }
}

/**
 * Handles GET requests for the standalone Document Archive web app.
 * @param {object} e - The event parameter from doGet.
 * @returns {HtmlOutput} The HTML page for the document archive.
 */
function handleDocumentArchiveRequest(e) {
  // Permission check
  if (typeof hasPermission === 'function' && !hasPermission('VIEW_DOCUMENT_ARCHIVE')) {
    return HtmlService.createHtmlOutput('<h1>Tilgang nektet</h1><p>Du har ikke tilgang til å se denne siden.</p>');
  }

  const key = 'DOKUMENTARKIV';
  const cfg = globalThis.UI_FILES?.[key];
  if (!cfg) {
    return HtmlService.createHtmlOutput('<h1>Feil</h1><p>UI-konfigurasjon for dokumentarkivet ble ikke funnet.</p>');
  }

  const templateName = String(cfg.file).replace(/\.html?$/i, '');
  const template = HtmlService.createTemplateFromFile(templateName);

  template.FILE = cfg.file;
  template.VERSION = APP.VERSION;
  template.UPDATED = APP.BUILD;
  template.PARAMS = e?.parameter || {};

  const output = template.evaluate()
    .setTitle(cfg.title || APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return output;
}