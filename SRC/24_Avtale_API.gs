/**
 * @OnlyCurrentDoc
 *
 * Backend for managing suppliers and their service agreements.
 */

// --- CONFIGURATION ---
const LEVERANDOR_SHEET_NAME = 'Leverandører';
const AVTALER_SHEET_NAME = 'Avtaler';

/**
 * Fetches all suppliers from the 'Leverandører' sheet.
 * @returns {object} A response object with the list of suppliers.
 */
function avtaleGetLeverandorer() {
  try {
    _validateConfig(); // Assuming a global _validateConfig exists
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(LEVERANDOR_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${LEVERANDOR_SHEET_NAME}" not found.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header

    const suppliers = data.map(row => {
      const supplier = {};
      headers.forEach((header, i) => {
        supplier[header] = row[i];
      });
      return supplier;
    }).filter(s => s.ID); // Ensure supplier has an ID

    return { ok: true, suppliers };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/**
 * Fetches all agreements for a specific supplier ID.
 * @param {string} leverandorId - The ID of the supplier.
 * @returns {object} A response object with the list of agreements.
 */
function avtaleGetAvtaler(leverandorId) {
  if (!leverandorId) return { ok: false, message: 'Leverandør-ID er påkrevd.' };
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(AVTALER_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${AVTALER_SHEET_NAME}" not found.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const leverandorIdIndex = headers.indexOf('LeverandorID');

    if (leverandorIdIndex === -1) {
      throw new Error("Kolonnen 'LeverandorID' ble ikke funnet i Avtaler-arket.");
    }

    const agreements = data.map(row => {
      const agreement = {};
      headers.forEach((header, i) => {
        agreement[header] = row[i];
      });
      return agreement;
    }).filter(a => a.LeverandorID == leverandorId);

    return { ok: true, agreements };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/**
 * Saves an agreement.
 * @param {object} payload The agreement data.
 * @returns {object} A response object.
 */
function avtaleLagre(payload) {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(AVTALER_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${AVTALER_SHEET_NAME}" not found.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (payload.AvtaleID) {
      // Update
      const data = sheet.getDataRange().getValues();
      const idIndex = headers.indexOf('AvtaleID');
      const rowIndex = data.findIndex(row => row[idIndex] == payload.AvtaleID);

      if (rowIndex > 0) {
        const row = headers.map(header => payload[header] !== undefined ? payload[header] : data[rowIndex][headers.indexOf(header)]);
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);
      } else {
        throw new Error(`Avtale med ID ${payload.AvtaleID} ikke funnet.`);
      }
    } else {
      // Create
      payload.AvtaleID = Utilities.getUuid();
      const newRow = headers.map(header => payload[header] || '');
      sheet.appendRow(newRow);
    }

    return { ok: true, AvtaleID: payload.AvtaleID };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}