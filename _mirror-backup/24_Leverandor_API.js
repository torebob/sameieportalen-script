/* ====================== Leverandør - API ======================
 * FILE: 24_Leverandor_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Håndtere logikk for leverandørprofiler og dokumentopplasting.
 * ================================================================== */

const LEVERANDOR_SHEET = 'LEVERANDØRER';
const AVTALE_FOLDER_NAME = 'Avtaledokumenter (Leverandører)';

/**
 * Henter detaljer for en spesifikk leverandør basert på radnummer.
 * @param {string|number} vendorId - Radnummeret til leverandøren.
 * @returns {Object} Et objekt med leverandørdetaljer.
 */
function getLeverandorDetails(vendorId) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(LEVERANDOR_SHEET);
    if (!sheet) {
      throw new Error(`Arket '${LEVERANDOR_SHEET}' ble ikke funnet.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(parseInt(vendorId, 10), 1, 1, headers.length).getValues()[0];

    const vendorDetails = {};
    headers.forEach((header, i) => {
      // Clean header name to be used as a key
      const key = header.replace(/\s/g, '').replace('/', '');
      vendorDetails[key] = data[i];
    });
     // For frontend compatibility
    vendorDetails.Navn = vendorDetails.Navn || 'Ukjent';
    vendorDetails.Kategori = vendorDetails['KategoriSystem'] || '';
    vendorDetails.AvtaleURL = vendorDetails['AvtaledokumentURL'] || '';


    return vendorDetails;
  } catch (e) {
    Logger.log(`getLeverandorDetails feil: ${e.message}`);
    throw new Error(`Kunne ikke hente leverandørdetaljer: ${e.message}`);
  }
}

/**
 * Laster opp et avtaledokument for en leverandør.
 * @param {string|number} vendorId - Radnummeret til leverandøren.
 * @param {Object} fileData - Filobjekt fra klienten.
 * @param {string} fileData.fileName - Filnavn.
 * @param {string} fileData.mimeType - Mime-type.
 * @param {string} fileData.content - Base64-kodet filinnhold.
 * @returns {string} URL-en til den opplastede filen.
 */
function uploadAvtaleDokument(vendorId, fileData) {
  try {
    // 1. Finn eller opprett mappen i Google Drive
    let folder = getOrCreateFolder_(AVTALE_FOLDER_NAME);

    // 2. Dekode og last opp filen
    const decodedContent = Utilities.base64Decode(fileData.content);
    const blob = Utilities.newBlob(decodedContent, fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();

    // 3. Oppdater regnearket med lenken til filen
    const sheet = SpreadsheetApp.getActive().getSheetByName(LEVERANDOR_SHEET);
    if (!sheet) {
      throw new Error(`Arket '${LEVERANDOR_SHEET}' ble ikke funnet.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const urlColumnIndex = headers.indexOf('Avtaledokument URL') + 1;
    if (urlColumnIndex === 0) {
      throw new Error("Kolonnen 'Avtaledokument URL' mangler i arket.");
    }

    sheet.getRange(parseInt(vendorId, 10), urlColumnIndex).setValue(fileUrl);

    return fileUrl;
  } catch (e) {
    Logger.log(`uploadAvtaleDokument feil: ${e.message}`);
    throw new Error(`Opplasting feilet: ${e.message}`);
  }
}

/**
 * Hjelpefunksjon for å finne eller opprette en mappe i Drive.
 * @param {string} folderName - Navnet på mappen.
 * @returns {GoogleAppsScript.Drive.Folder} Mappe-objektet.
 */
function getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Henter alle leverandører fra regnearket.
 * @returns {Array<Object>} En liste med leverandørobjekter.
 */
function getAlleLeverandorer() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(LEVERANDOR_SHEET);
    if (!sheet || sheet.getLastRow() < 2) {
      return []; // Tomt ark eller bare header
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    return data.map((row, index) => {
      const vendor = {};
      headers.forEach((header, i) => {
        const key = header.replace(/\s/g, '').replace('/', '');
        vendor[key] = row[i];
      });
      vendor.id = index + 2; // Radnummer er ID
      vendor.Navn = vendor.Navn || 'Ukjent';
      vendor.Kategori = vendor['KategoriSystem'] || '';
      return vendor;
    });
  } catch (e) {
    Logger.log(`getAlleLeverandorer feil: ${e.message}`);
    throw new Error(`Kunne ikke hente leverandørlisten: ${e.message}`);
  }
}