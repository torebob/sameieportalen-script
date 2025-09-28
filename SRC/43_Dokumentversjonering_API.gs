/**
 * ==================================================================
 *   FILE: 43_Dokumentversjonering_API.gs
 *   VERSION: 1.0.0
 *   UPDATED: 2025-09-28
 *   PURPOSE: API for document versioning.
 * ==================================================================
 */

const DOK_LOGG_SHEET = 'Styringsdokumenter_Logg';

/**
 * Extracts the Google Drive file ID from a URL.
 * @param {string} url The Google Drive file URL.
 * @returns {string|null} The file ID or null if not found.
 */
function getFileIdFromUrl_(url) {
  if (!url || typeof url !== 'string') return null;
  const match = url.match(/\/d\/(.+?)\//);
  return match ? match[1] : null;
}

/**
 * Retrieves the list of manageable documents from the log sheet.
 * @returns {Array<Object>} A list of document objects.
 */
function getVersionableDocuments() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DOK_LOGG_SHEET);
    if (!sheet) {
      throw new Error(`Sheet "${DOK_LOGG_SHEET}" not found.`);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const masterUrlIndex = headers.indexOf('MasterURL');
    const kategoriIndex = headers.indexOf('Kategori');

    if (masterUrlIndex === -1 || kategoriIndex === -1) {
      throw new Error('Required columns (MasterURL, Kategori) not found in sheet.');
    }

    return data.map((row, index) => {
      const masterUrl = row[masterUrlIndex];
      const fileId = getFileIdFromUrl_(masterUrl);
      return {
        rowNum: index + 2, // 1-based index for the sheet row
        kategori: row[kategoriIndex],
        masterUrl: masterUrl,
        fileId: fileId,
      };
    }).filter(doc => doc.fileId); // Only include documents with a valid file ID
  } catch (e) {
    Logger.log(`Error in getVersionableDocuments: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Updates a file in Google Drive with new content, creating a new version.
 * NOTE: This requires the Drive API advanced service to be enabled.
 *
 * @param {string} fileId The ID of the file to update.
 * @param {string} base64Data The new content encoded in base64.
 * @param {string} mimeType The MIME type of the file.
 * @param {string} fileName The name of the file.
 * @param {number} rowNum The row number in the log sheet to update.
 * @returns {Object} A result object.
 */
function updateDocumentVersion(fileId, base64Data, mimeType, fileName, rowNum) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);

    // This call updates the file content and creates a new revision
    const updatedFile = Drive.Files.update({ title: fileName }, fileId, blob, { mimeType: mimeType });

    // Update the log sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DOK_LOGG_SHEET);
    const lastEditedCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('SistEndret') + 1;
    const versionCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('SisteVersjon') + 1;

    if (lastEditedCol > 0) {
      sheet.getRange(rowNum, lastEditedCol).setValue(new Date());
    }
    if (versionCol > 0) {
      const currentVersion = sheet.getRange(rowNum, versionCol).getValue();
      // A simple version increment, can be made more sophisticated
      const newVersion = (typeof currentVersion === 'number') ? currentVersion + 0.1 : 1.1;
      sheet.getRange(rowNum, versionCol).setValue(newVersion.toFixed(1));
    }

    return { ok: true, message: `File "${fileName}" updated successfully to version ${updatedFile.version}.` };
  } catch (e) {
    Logger.log(`Error in updateDocumentVersion: ${e.message}`);
    // Check for common error if Drive API is not enabled
    if (e.message.includes("Drive is not defined")) {
       return { error: "The Google Drive API advanced service is not enabled. Please enable it in the script editor." };
    }
    return { error: e.message };
  }
}

/**
 * Retrieves the revision history for a specific file.
 * NOTE: This requires the Drive API advanced service to be enabled.
 *
 * @param {string} fileId The ID of the file to get history for.
 * @returns {Object} An object containing the list of revisions or an error.
 */
function getDocumentRevisionHistory(fileId) {
  try {
    const revisions = Drive.Revisions.list(fileId);
    const history = revisions.items.map(revision => {
      return {
        id: revision.id,
        modifiedDate: revision.modifiedDate,
        user: revision.lastModifyingUser ? revision.lastModifyingUser.displayName : 'Unknown',
        fileSize: revision.fileSize,
        version: revision.id, // The revision ID can serve as a version identifier
      };
    }).sort((a, b) => new Date(b.modifiedDate) - new Date(a.modifiedDate)); // Sort descending by date

    return { ok: true, history: history };
  } catch (e) {
    Logger.log(`Error in getDocumentRevisionHistory: ${e.message}`);
    if (e.message.includes("Drive is not defined")) {
       return { error: "The Google Drive API advanced service is not enabled. Please enable it in the script editor." };
    }
    return { error: e.message };
  }
}