/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file access for this script
 * to only the current document. This is a best practice for security.
 */

// --- USER ACTION REQUIRED ---
// To use the new "Henvendelse" feature, please add the following columns
// to your 'Tasks' sheet in your Google Sheet:
// - kategori
// - innsendt_av
// --------------------------

// --- CONFIGURATION ---
const DB_SHEET_ID = 'YOUR_SHEET_ID_HERE'; // Replace with the actual ID of the Google Sheet
const TASKS_SHEET_NAME = 'Tasks';
const USERS_SHEET_NAME = 'Users';
const SUPPLIERS_SHEET_NAME = 'Suppliers';
const MESSAGES_SHEET_NAME = 'Messages';
const SECTIONS_SHEET_NAME = 'Sections';
const OWNERS_SHEET_NAME = 'Owners';
const TENANTS_SHEET_NAME = 'Tenants';
const ATTACHMENTS_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; // Replace with the ID of the Google Drive folder for attachments

/**
 * Creates the Messages sheet if it doesn't already exist.
 * @private
 */
function _createMessagesSheetIfNotExist() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  if (!ss.getSheetByName(MESSAGES_SHEET_NAME)) {
    const sheet = ss.insertSheet(MESSAGES_SHEET_NAME);
    sheet.appendRow(['id', 'timestamp', 'title', 'content', 'recipients', 'attachmentUrl']);
  }
}

/**
 * Creates the sheets for the Beboerregister module if they don't already exist.
 * @private
 */
function _createBeboerregisterSheetsIfNotExist() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);

  // Create Sections sheet
  if (!ss.getSheetByName(SECTIONS_SHEET_NAME)) {
    const sheet = ss.insertSheet(SECTIONS_SHEET_NAME);
    sheet.appendRow(['id', 'seksjonsnummer', 'adresse', 'areal', 'antall_rom', 'etg']);
  }

  // Create Owners sheet
  if (!ss.getSheetByName(OWNERS_SHEET_NAME)) {
    const sheet = ss.insertSheet(OWNERS_SHEET_NAME);
    sheet.appendRow(['id', 'sectionId', 'navn', 'fodselsdato_orgnr', 'epost', 'telefon']);
  }

  // Create Tenants sheet
  if (!ss.getSheetByName(TENANTS_SHEET_NAME)) {
    const sheet = ss.insertSheet(TENANTS_SHEET_NAME);
    sheet.appendRow(['id', 'sectionId', 'navn', 'epost', 'telefon', 'leieperiodeStart', 'leieperiodeSlutt']);
  }
}

/**
 * Validates that the script has been configured.
 * @private
 */
function _validateConfig() {
  if (DB_SHEET_ID.startsWith('YOUR_') || ATTACHMENTS_FOLDER_ID.startsWith('YOUR_')) {
    throw new Error('Skriptet er ikke konfigurert. Vennligst følg SETUP_INSTRUCTIONS.md.');
  }
  _createMessagesSheetIfNotExist();
  _createBeboerregisterSheetsIfNotExist();
  _createHmsSheetsIfNotExist(); // Added this line
}

/**
 * Retrieves the list of tasks.
 * @returns {object} A response object with the list of tasks.
 */
function gjoremalGet() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(TASKS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${TASKS_SHEET_NAME}" ble ikke funnet. Vennligst sjekk arknavnet.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row

    const tasks = data.map(row => {
      const task = {};
      headers.forEach((header, i) => {
        task[header] = row[i];
      });
      return task;
    });

    return { ok: true, tasks: tasks };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

// --- Beboerregister Functions ---

/**
 * A generic helper function to fetch all data from a given sheet.
 * @private
 * @param {string} sheetName - The name of the sheet to read.
 * @returns {Array<Object>} An array of objects representing the rows.
 */
function _getSheetData(sheetName) {
  _validateConfig();
  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Arket "${sheetName}" ble ikke funnet.`);
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift();
  return data.map(row => {
    const record = {};
    headers.forEach((header, i) => {
      record[header] = row[i];
    });
    return record;
  });
}

/**
 * Fetches all data for the Beboerregister (Sections, Owners, Tenants).
 * @returns {object} A response object containing all the data.
 */
function getBeboerregisterData() {
  try {
    const sections = _getSheetData(SECTIONS_SHEET_NAME);
    const owners = _getSheetData(OWNERS_SHEET_NAME);
    const tenants = _getSheetData(TENANTS_SHEET_NAME);

    // Combine data for easier use on the frontend
    const combinedData = sections.map(section => {
      return {
        ...section,
        owners: owners.filter(o => o.sectionId === section.id),
        tenants: tenants.filter(t => t.sectionId === section.id),
      };
    });

    return { ok: true, data: combinedData };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke hente beboerregister: ${e.message}` };
  }
}

/**
 * Generic function to save a record (create or update) to a specified sheet.
 * @param {object} payload - The data object to save. Must include 'sheetName'.
 * @returns {object} A response object indicating success or failure.
 */
function saveBeboerRecord(payload) {
  try {
    _validateConfig();
    const { sheetName, ...record } = payload;
    if (!sheetName) throw new Error("Arknavn er påkrevd.");

    const validSheetNames = [SECTIONS_SHEET_NAME, OWNERS_SHEET_NAME, TENANTS_SHEET_NAME];
    if (!validSheetNames.includes(sheetName)) {
        throw new Error(`Ugyldig arknavn: ${sheetName}`);
    }

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
    if (!sheet) throw new Error(`Arket "${sheetName}" ble ikke funnet.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (record.id) {
      // Update existing record
      const data = sheet.getDataRange().getValues();
      const rowIndex = data.findIndex(row => row[0] == record.id);

      if (rowIndex > 0) {
        const rowData = data[rowIndex];
        const newRow = headers.map((header, i) => record[header] !== undefined ? record[header] : rowData[i]);
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
      } else {
        throw new Error(`Post med ID ${record.id} ble ikke funnet i ${sheetName}.`);
      }
    } else {
      // Create new record
      record.id = Utilities.getUuid();
      const newRow = headers.map(header => record[header] !== undefined ? record[header] : '');
      sheet.appendRow(newRow);
    }

    return { ok: true, id: record.id };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Serverfeil: ${e.message}` };
  }
}

/**
 * Generic function to delete a record from a specified sheet.
 * @param {object} payload - Must contain 'sheetName' and 'id'.
 * @returns {object} A response object indicating success or failure.
 */
function deleteBeboerRecord(payload) {
  try {
    _validateConfig();
    const { sheetName, id } = payload;
    if (!sheetName || !id) throw new Error("Arknavn og ID er påkrevd.");

    const validSheetNames = [SECTIONS_SHEET_NAME, OWNERS_SHEET_NAME, TENANTS_SHEET_NAME];
    if (!validSheetNames.includes(sheetName)) {
        throw new Error(`Ugyldig arknavn: ${sheetName}`);
    }

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
    if (!sheet) throw new Error(`Arket "${sheetName}" ble ikke funnet.`);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] == id);

    if (rowIndex > 0) {
        sheet.deleteRow(rowIndex + 1);
        return { ok: true };
    } else {
      return { ok: false, error: `Post med ID ${id} ble ikke funnet i ${sheetName}.` };
    }
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Serverfeil: ${e.message}` };
  }
}

/**
 * Exports beboer data to a CSV file and returns its base64 representation.
 * @param {string} type - The type of list to export ('all', 'owners', 'tenants').
 * @returns {object} A response object with the CSV data.
 */
function exportBeboerliste(type = 'all') {
    try {
        const sections = _getSheetData(SECTIONS_SHEET_NAME);
        const owners = _getSheetData(OWNERS_SHEET_NAME);
        const tenants = _getSheetData(TENANTS_SHEET_NAME);

        let csvContent = "";
        let fileName = "";

        if (type === 'owners' || type === 'all') {
            csvContent += "Eiere\\n";
            csvContent += "Seksjonsnr,Navn,E-post,Telefon\\n";
            owners.forEach(owner => {
                const section = sections.find(s => s.id === owner.sectionId);
                csvContent += `${section ? section.seksjonsnummer : 'N/A'},${owner.navn},${owner.epost},${owner.telefon}\\n`;
            });
            if (type === 'all') csvContent += "\\n";
        }

        if (type === 'tenants' || type === 'all') {
            csvContent += "Leietakere\\n";
            csvContent += "Seksjonsnr,Navn,E-post,Telefon,Leieperiode Start,Leieperiode Slutt\\n";
            tenants.forEach(tenant => {
                const section = sections.find(s => s.id === tenant.sectionId);
                csvContent += `${section ? section.seksjonsnummer : 'N/A'},${tenant.navn},${tenant.epost},${tenant.telefon},${tenant.leieperiodeStart},${tenant.leieperiodeSlutt}\\n`;
            });
        }

        switch(type) {
            case 'owners': fileName = 'eierliste.csv'; break;
            case 'tenants': fileName = 'leietakerliste.csv'; break;
            default: fileName = 'beboerliste.csv'; break;
        }

        const base64Csv = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);

        return { ok: true, file: { name: fileName, base64: base64Csv, mimeType: 'text/csv' } };

    } catch (e) {
        Logger.log(e);
        return { ok: false, error: `Kunne ikke eksportere liste: ${e.message}` };
    }
}

/**
 * Retrieves the list of recipient groups.
 * @returns {Array<string>} A list of recipient groups.
 */
function getRecipientGroups() {
  // In a real application, this could be dynamic based on user roles or properties.
  return ['Alle beboere', 'Kun eiere', 'Kun leietakere', 'Styret'];
}

/**
 * Sends a new message (oppslag) to the specified recipients.
 * @param {object} payload The message data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function sendOppslag(payload) {
  try {
    _validateConfig();
    const { tittel, innhold, maalgruppe, attachment } = payload;

    if (!tittel || !innhold || !maalgruppe) {
      throw new Error('Mangler tittel, innhold eller målgruppe.');
    }

    let attachmentUrl = '';
    if (attachment) {
      const { base64, mimeType, name } = attachment;
      const decoded = Utilities.base64Decode(base64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, mimeType, name);
      const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      const file = folder.createFile(blob);
      attachmentUrl = file.getUrl();
    }

    // Store message in history
    const messageSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(MESSAGES_SHEET_NAME);
    const newId = Utilities.getUuid();
    const timestamp = new Date();
    messageSheet.appendRow([newId, timestamp, tittel, innhold, maalgruppe, attachmentUrl]);

    // Fetch recipient emails (this is a simplified example)
    const userSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
    const usersData = userSheet.getDataRange().getValues();
    usersData.shift(); // remove headers
    const recipients = usersData.map(row => row[1]).filter(email => email); // Get all user emails for this example

    // Send email notifications
    const subject = `Nytt oppslag: ${tittel}`;
    let body = `<p><b>${tittel}</b></p><p>${innhold.replace(/\n/g, '<br>')}</p>`;
    if (attachmentUrl) {
      body += `<p>Se vedlegg: <a href="${attachmentUrl}">Klikk her</a></p>`;
    }

    // In a real app, you would filter recipients based on `maalgruppe`
    recipients.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body,
        name: "Styret"
      });
    });

    return { ok: true, message: 'Oppslaget ble sendt!' };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `En feil oppstod: ${e.message}` };
  }
}

/**
 * Retrieves the history of sent messages.
 * @returns {object} A response object with the list of messages.
 */
function getMessageHistory() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(MESSAGES_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${MESSAGES_SHEET_NAME}" ble ikke funnet.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift() || [];

    const messages = data.map(row => {
      const message = {};
      headers.forEach((header, i) => {
        message[header] = row[i];
      });
      return message;
    }).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Sort by newest first

    return { ok: true, messages: messages };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke hente meldingshistorikk: ${e.message}` };
  }
}

/**
 * Handles a new inquiry (henvendelse) from a resident and creates a task from it.
 * @param {object} payload The inquiry data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function sendHenvendelse(payload) {
  try {
    _validateConfig();
    const { tittel, innhold, kategori, attachment } = payload;

    if (!tittel || !innhold || !kategori) {
      throw new Error('Mangler tittel, innhold eller kategori.');
    }

    // BSM.11.4: Log timestamp and sender
    const innsendtTid = new Date();
    const innsendtAv = Session.getEffectiveUser().getEmail();

    // Prepare a task object that matches the structure of the 'Tasks' sheet
    const taskPayload = {
      description: tittel, // Using 'description' for the task title
      notes: innhold,      // Using 'notes' for the message body
      kategori: kategori,
      innsendt_av: innsendtAv,
      status: 'Open',      // Default status for new tasks
      created_at: innsendtTid,
      attachment: attachment // Pass attachment through if it exists
    };

    // BSM.11.5: Use existing function to save the inquiry as a task
    const result = gjoremalSave(taskPayload);

    if (!result.ok) {
      throw new Error(result.message || 'Klarte ikke å lagre henvendelsen som en oppgave.');
    }

    return { ok: true };

  } catch (e) {
    Logger.log(`Feil i sendHenvendelse: ${e.message}`);
    return { ok: false, error: e.message };
  }
}

/**
 * Retrieves the list of suppliers.
 * @returns {object} A response object with the list of suppliers.
 */
function getSuppliers() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(SUPPLIERS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${SUPPLIERS_SHEET_NAME}" ble ikke funnet.`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, suppliers: [] }; // No data rows is a valid state
    const headers = data.shift();

    const suppliers = data.map(row => {
      const supplier = {};
      headers.forEach((header, i) => {
        supplier[header] = row[i];
      });
      return supplier;
    });

    return { ok: true, suppliers: suppliers };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/**
 * Saves a supplier (creates a new one or updates an existing one).
 * @param {object} payload The supplier data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function saveSupplier(payload) {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(SUPPLIERS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${SUPPLIERS_SHEET_NAME}" ble ikke funnet.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (payload.id) {
      // Update existing supplier
      const data = sheet.getDataRange().getValues();
      const rowIndex = data.findIndex(row => row[0] == payload.id); // Assumes ID is in the first column

      if (rowIndex > 0) { // rowIndex > 0 means it's not the header
        const rowData = data[rowIndex];
        const newRow = headers.map((header, i) => payload[header] !== undefined ? payload[header] : rowData[i]);
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
      } else {
        throw new Error(`Leverandør med ID ${payload.id} ble ikke funnet.`);
      }
    } else {
      // Create new supplier
      payload.id = Utilities.getUuid();
      const newRow = headers.map(header => payload[header] !== undefined ? payload[header] : '');
      sheet.appendRow(newRow);
    }

    return { ok: true, id: payload.id };
  } catch (e) {
    return { ok: false, message: `Serverfeil: ${e.message}` };
  }
}

/**
 * Deletes a supplier by their ID.
 * @param {string} id The ID of the supplier to delete.
 * @returns {object} A response object indicating success or failure.
 */
function deleteSupplier(id) {
  try {
    _validateConfig();
    if (!id) throw new Error("Leverandør-ID er påkrevd for sletting.");

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(SUPPLIERS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${SUPPLIERS_SHEET_NAME}" ble ikke funnet.`);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] == id); // Assumes ID is in the first column

    if (rowIndex > 0) { // rowIndex > 0 means it's not the header
        sheet.deleteRow(rowIndex + 1); // sheet rows are 1-indexed, so rowIndex+1 is the correct row number
        return { ok: true };
    } else {
      return { ok: false, message: `Leverandør med ID ${id} ble ikke funnet.` };
    }
  } catch (e) {
    return { ok: false, message: `Serverfeil: ${e.message}` };
  }
}

/**
 * Saves a task (creates a new one or updates an existing one).
 * @param {object} payload The task data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function gjoremalSave(payload) {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(TASKS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${TASKS_SHEET_NAME}" ble ikke funnet. Vennligst sjekk arknavnet.`);

    // Handle attachment upload
    if (payload.attachment) {
      const { base64, mimeType, name } = payload.attachment;
      const decoded = Utilities.base64Decode(base64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, mimeType, name);

      const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      const file = folder.createFile(blob);

      payload.attachmentUrl = file.getUrl(); // Add URL to payload for sheet storage
    }
    delete payload.attachment; // Remove base64 data before saving to sheet

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (payload.id) {
      // Update existing task
      const data = sheet.getDataRange().getValues();
      const rowIndex = data.findIndex(row => row[0] == payload.id); // Assuming ID is in the first column

      if (rowIndex > 0) {
        const rowData = data[rowIndex];
        const row = headers.map((header, index) => payload[header] !== undefined ? payload[header] : rowData[index]);
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);
      } else {
        throw new Error(`Oppgave med ID ${payload.id} ble ikke funnet.`);
      }
    } else {
      // Create new task
      const newId = Utilities.getUuid();
      payload.id = newId;
      payload.status = 'Open';

      const newRow = headers.map(header => {
        // Use the value from the payload if it exists, otherwise use an empty string.
        // This is more robust than `|| ''` as it correctly handles `false` or `0` values.
        return payload[header] !== undefined ? payload[header] : '';
      });

      sheet.appendRow(newRow);
    }

    return { ok: true };
  } catch (e) {
    return { ok: false, message: `Serverfeil: ${e.message}` };
  }
}

/**
 * Retrieves the list of users.
 * For this example, we'll use a sheet. In a real scenario, this could come from another source.
 * @returns {object} A response object with the list of users.
 */
function gjoremalGetUsers() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Arket "${USERS_SHEET_NAME}" ble ikke funnet. Vennligst sjekk arknavnet.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Assumes headers: 'name', 'email'
    if (!headers || headers.length < 2) {
      throw new Error(`Arket "${USERS_SHEET_NAME}" må ha minst kolonnene 'name' og 'email'.`);
    }

    const users = data.map(row => ({ name: row[0], email: row[1] }));

    return { ok: true, users: users };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}