/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file access for this script
 * to only the current document. This is a best practice for security.
 */

// --- CONFIGURATION ---
const DB_SHEET_ID = 'YOUR_SHEET_ID_HERE'; // Replace with the actual ID of the Google Sheet
const TASKS_SHEET_NAME = 'Tasks';
const USERS_SHEET_NAME = 'Users';
const SUPPLIERS_SHEET_NAME = 'Suppliers';
const MESSAGES_SHEET_NAME = 'Messages';
const INVOICES_SHEET_NAME = 'Invoices'; // New sheet for invoices
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
 * Creates the Invoices sheet if it doesn't already exist.
 * @private
 */
function _createInvoicesSheetIfNotExist() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  if (!ss.getSheetByName(INVOICES_SHEET_NAME)) {
    const sheet = ss.insertSheet(INVOICES_SHEET_NAME);
    // Headers for the invoice sheet
    sheet.appendRow(['id', 'supplierName', 'amount', 'dueDate', 'invoiceDate', 'description', 'attachmentUrl', 'status', 'attestationHistory', 'rules']);
  }
}

/**
 * Validates that the script has been configured.
 * @private
 */
function _validateConfig() {
  if (DB_SHEET_ID.startsWith('YOUR_') || ATTACHMENTS_FOLDER_ID.startsWith('YOUR_')) {
    throw new Error('Script not configured. Please follow SETUP_INSTRUCTIONS.md.');
  }
  _createMessagesSheetIfNotExist();
  _createInvoicesSheetIfNotExist(); // Ensure the invoices sheet exists
}

/**
 * Retrieves the list of tasks.
 * @returns {object} A response object with the list of tasks.
 */
function gjoremalGet() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(TASKS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${TASKS_SHEET_NAME}" not found. Please check sheet name.`);

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

/**
 * Gets the email of the active user.
 * @returns {string} The user's email address.
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// --- INVOICE ATTESTATION FUNCTIONS ---

/**
 * Retrieves all invoices from the Invoices sheet.
 * @returns {object} A response object with the list of invoices.
 */
function getInvoices() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${INVOICES_SHEET_NAME}" not found.`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, invoices: [] }; // No data is valid
    const headers = data.shift();

    const invoices = data.map(row => {
      const invoice = {};
      headers.forEach((header, i) => {
        invoice[header] = row[i];
      });
      // Attempt to parse history if it's a string
      if (invoice.attestationHistory && typeof invoice.attestationHistory === 'string') {
        try {
          invoice.attestationHistory = JSON.parse(invoice.attestationHistory);
        } catch (e) {
          invoice.attestationHistory = []; // Or handle error appropriately
        }
      }
      return invoice;
    }).sort((a, b) => new Date(b.invoiceDate) - new Date(a.invoiceDate)); // Sort by newest first

    return { ok: true, invoices: invoices };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Could not retrieve invoices: ${e.message}` };
  }
}

/**
 * Adds a new invoice and notifies the board.
 * This is a helper function for demonstration; in a real-world scenario,
 * this would be triggered by an external system (e.g., email from an accountant).
 * @param {object} payload The invoice data.
 * @returns {object} A response object indicating success or failure.
 */
function addInvoice(payload) {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICES_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const newInvoice = {
      id: Utilities.getUuid(),
      status: 'Pending',
      attestationHistory: JSON.stringify([]),
      ...payload
    };

    // Set default rules if not provided
    if (!newInvoice.rules) {
        newInvoice.rules = JSON.stringify({ requiredApprovals: newInvoice.amount > 5000 ? 2 : 1 });
    }

    const newRow = headers.map(header => newInvoice[header] || '');
    sheet.appendRow(newRow);

    // --- Notify Board Members ---
    const userSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(USERS_SHEET_NAME);
    const usersData = userSheet.getDataRange().getValues();
    usersData.shift(); // remove headers
    const boardEmails = usersData
      .filter(row => row[2] === 'Styret') // Assuming column C indicates role
      .map(row => row[1])
      .filter(email => email);

    if (boardEmails.length > 0) {
      const subject = `Ny faktura til attestering: ${newInvoice.supplierName}`;
      const body = `
        <p>En ny faktura krever din oppmerksomhet.</p>
        <p><b>Leverandør:</b> ${newInvoice.supplierName}</p>
        <p><b>Beløp:</b> ${newInvoice.amount} kr</p>
        <p><b>Forfallsdato:</b> ${new Date(newInvoice.dueDate).toLocaleDateString()}</p>
        <p>Vennligst logg inn i applikasjonen for å behandle den.</p>
      `;
      MailApp.sendEmail({
        to: boardEmails.join(','),
        subject: subject,
        htmlBody: body,
        name: "System"
      });
    }

    return { ok: true, invoice: newInvoice };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Could not add invoice: ${e.message}` };
  }
}


/**
 * Attests (approves or rejects) an invoice.
 * @param {object} payload Contains invoiceId, action ('Approve' or 'Reject'), and userEmail.
 * @returns {object} A response object indicating success or failure.
 */
function attestInvoice(payload) {
  try {
    _validateConfig();
    const { invoiceId, action, userEmail } = payload;
    if (!invoiceId || !action || !userEmail) {
      throw new Error("Mangler invoiceId, action, eller userEmail.");
    }

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICES_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf('id');
    const rowIndex = data.findIndex(row => row[idColIndex] === invoiceId);

    if (rowIndex === -1) {
      throw new Error(`Faktura med ID ${invoiceId} ble ikke funnet.`);
    }

    const rowData = data[rowIndex];
    const invoice = {};
    headers.forEach((header, i) => {
      invoice[header] = rowData[i];
    });

    // --- Update Attestation History ---
    let history = invoice.attestationHistory ? JSON.parse(invoice.attestationHistory) : [];

    // Prevent duplicate attestations
    if (history.some(entry => entry.user === userEmail)) {
        return { ok: false, message: "Du har allerede attestert denne fakturaen." };
    }

    history.push({
      user: userEmail,
      action: action,
      timestamp: new Date().toISOString()
    });

    // --- Update Status based on Rules ---
    const rules = invoice.rules ? JSON.parse(invoice.rules) : { requiredApprovals: 1 };
    const requiredApprovals = rules.requiredApprovals || 1;
    const currentApprovals = history.filter(h => h.action === 'Approve').length;

    let newStatus = invoice.status;
    if (action === 'Reject') {
      newStatus = 'Rejected';
    } else if (action === 'Approve') {
      if (currentApprovals >= requiredApprovals) {
        newStatus = 'Approved';
      } else {
        newStatus = 'Partially Approved';
      }
    }

    // --- Write Updates to Sheet ---
    const statusColIndex = headers.indexOf('status');
    const historyColIndex = headers.indexOf('attestationHistory');

    sheet.getRange(rowIndex + 2, statusColIndex + 1).setValue(newStatus);
    sheet.getRange(rowIndex + 2, historyColIndex + 1).setValue(JSON.stringify(history));

    return { ok: true, newStatus: newStatus, newHistory: history };

  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Attestering feilet: ${e.message}` };
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
    if (!sheet) throw new Error(`Sheet "${MESSAGES_SHEET_NAME}" not found.`);

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
 * Retrieves the list of suppliers.
 * @returns {object} A response object with the list of suppliers.
 */
function getSuppliers() {
  try {
    _validateConfig();
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(SUPPLIERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SUPPLIERS_SHEET_NAME}" not found.`);

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
    if (!sheet) throw new Error(`Sheet "${SUPPLIERS_SHEET_NAME}" not found.`);

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
        throw new Error(`Supplier with ID ${payload.id} not found.`);
      }
    } else {
      // Create new supplier
      payload.id = Utilities.getUuid();
      const newRow = headers.map(header => payload[header] !== undefined ? payload[header] : '');
      sheet.appendRow(newRow);
    }

    return { ok: true, id: payload.id };
  } catch (e) {
    return { ok: false, message: `Server error: ${e.message}` };
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
    if (!id) throw new Error("Supplier ID is required for deletion.");

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(SUPPLIERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SUPPLIERS_SHEET_NAME}" not found.`);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] == id); // Assumes ID is in the first column

    if (rowIndex > 0) { // rowIndex > 0 means it's not the header
        sheet.deleteRow(rowIndex + 1); // sheet rows are 1-indexed, so rowIndex+1 is the correct row number
        return { ok: true };
    } else {
      return { ok: false, message: `Supplier with ID ${id} not found.` };
    }
  } catch (e) {
    return { ok: false, message: `Server error: ${e.message}` };
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
    if (!sheet) throw new Error(`Sheet "${TASKS_SHEET_NAME}" not found. Please check sheet name.`);

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
        throw new Error(`Task with ID ${payload.id} not found.`);
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
    return { ok: false, message: `Server error: ${e.message}` };
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
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found. Please check sheet name.`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Assumes headers: 'name', 'email'
    if (!headers || headers.length < 2) {
      throw new Error(`Sheet "${USERS_SHEET_NAME}" must have at least 'name' and 'email' columns.`);
    }

    const users = data.map(row => ({ name: row[0], email: row[1] }));

    return { ok: true, users: users };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}