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
const ATTACHMENTS_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; // Replace with the ID of the Google Drive folder for attachments

/**
 * Creates the Messages sheet if it doesn't already exist.
 * @private
 */
// --- NEW INVOICING CONFIGURATION ---
const INVOICES_SHEET_NAME = 'Fakturaer';
const INVOICE_TEMPLATES_SHEET_NAME = 'Fakturamaler';
const SECTIONS_SHEET_NAME = 'Seksjoner';

/**
 * Creates the necessary sheets for the invoicing module if they don't already exist.
 * @private
 */
function _createInvoicingSheetsIfNotExist() {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);

    const sheets = {
        [INVOICES_SHEET_NAME]: ['id', 'seksjonsnummer', 'belop', 'forfallsdato', 'status', 'opprettetDato', 'beskrivelse', 'betaltBelop'],
        [INVOICE_TEMPLATES_SHEET_NAME]: ['id', 'navn', 'belop', 'beskrivelse'],
        [SECTIONS_SHEET_NAME]: ['seksjonsnummer', 'eierNavn', 'eierEpost', 'andel']
    };

    for (const sheetName in sheets) {
        if (!ss.getSheetByName(sheetName)) {
            const sheet = ss.insertSheet(sheetName);
            sheet.appendRow(sheets[sheetName]);
            if (sheetName === SECTIONS_SHEET_NAME) {
                 // Add sample data for sections to facilitate testing
                sheet.appendRow(['101', 'Ola Nordmann', 'ola.nordmann@example.com', '1.0']);
                sheet.appendRow(['102', 'Kari Nordmann', 'kari.nordmann@example.com', '1.2']);
            }
        }
    }
}

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
 * Validates that the script has been configured.
 * @private
 */
function _validateConfig() {
  if (DB_SHEET_ID.startsWith('YOUR_') || ATTACHMENTS_FOLDER_ID.startsWith('YOUR_')) {
    throw new Error('Script not configured. Please follow SETUP_INSTRUCTIONS.md.');
  }
  _createMessagesSheetIfNotExist();
  _createInvoicingSheetsIfNotExist(); // Add invoicing sheets
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

// --- INVOICING FUNCTIONS ---

/**
 * Retrieves all invoice templates.
 * @returns {object} A response object with the list of templates.
 */
function getInvoiceTemplates() {
    try {
        _validateConfig();
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICE_TEMPLATES_SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${INVOICE_TEMPLATES_SHEET_NAME}" not found.`);

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return { ok: true, templates: [] };
        const headers = data.shift();

        const templates = data.map(row => {
            const template = {};
            headers.forEach((header, i) => {
                template[header] = row[i];
            });
            return template;
        });

        return { ok: true, templates: templates };
    } catch (e) {
        return { ok: false, message: `Server error: ${e.message}` };
    }
}

/**
 * Sends an invoice email to the resident.
 * @param {string} invoiceId The ID of the invoice to send.
 * @returns {object} A response object indicating success or failure.
 */
function sendInvoiceEmail(invoiceId) {
    try {
        _validateConfig();
        if (!invoiceId) throw new Error("Invoice ID is required.");

        const { invoices } = getInvoices();
        const invoice = invoices.find(inv => inv.id === invoiceId);

        if (!invoice) throw new Error("Invoice not found.");
        if (!invoice.eierEpost) throw new Error("Recipient email not found for this section.");

        const subject = `Faktura fra Sameiet: ${invoice.beskrivelse}`;
        const body = `
            <p>Hei ${invoice.eierNavn},</p>
            <p>Vedlagt følger faktura for ${invoice.beskrivelse}.</p>
            <ul>
                <li><strong>Fakturabeløp:</strong> ${invoice.belop.toFixed(2)} kr</li>
                <li><strong>Forfallsdato:</strong> ${new Date(invoice.forfallsdato).toLocaleDateString()}</li>
            </ul>
            <p>Vennligst betal innen forfall.</p>
            <p>Med vennlig hilsen,<br>Styret</p>
        `;

        MailApp.sendEmail({
            to: invoice.eierEpost,
            subject: subject,
            htmlBody: body,
            name: "Styret i Sameiet"
        });

        // Optionally, update invoice status to 'Sendt'
        // This would require finding the row and updating the status column.

        return { ok: true, message: "Fakturaen ble sendt på e-post." };
    } catch (e) {
        return { ok: false, message: `E-post-feil: ${e.message}` };
    }
}

/**
 * Sends a payment reminder for an overdue invoice.
 * @param {string} invoiceId The ID of the overdue invoice.
 * @returns {object} A response object indicating success or failure.
 */
function sendPaymentReminder(invoiceId) {
    try {
        _validateConfig();
        if (!invoiceId) throw new Error("Invoice ID is required.");

        const { invoices } = getInvoices();
        const invoice = invoices.find(inv => inv.id === invoiceId);

        if (!invoice) throw new Error("Invoice not found.");
        if (!invoice.eierEpost) throw new Error("Recipient email not found for this section.");

        const subject = `Betalingspåminnelse: Faktura ${invoice.beskrivelse}`;
        const body = `
            <p>Hei ${invoice.eierNavn},</p>
            <p>Vi minner om ubetalt faktura for ${invoice.beskrivelse}.</p>
            <ul>
                <li><strong>Fakturabeløp:</strong> ${invoice.belop.toFixed(2)} kr</li>
                <li><strong>Forfallsdato:</strong> ${new Date(invoice.forfallsdato).toLocaleDateString()}</li>
            </ul>
            <p>Vi ber deg vennligst om å betale denne så snart som mulig.</p>
            <p>Med vennlig hilsen,<br>Styret</p>
        `;

        MailApp.sendEmail({
            to: invoice.eierEpost,
            subject: subject,
            htmlBody: body,
            name: "Styret i Sameiet"
        });

        return { ok: true, message: "Påminnelsen ble sendt." };
    } catch (e) {
        return { ok: false, message: `E-post-feil: ${e.message}` };
    }
}

/**
 * Retrieves all invoices, along with owner details from the Sections sheet.
 * @returns {object} A response object with the list of invoices.
 */
function getInvoices() {
    try {
        _validateConfig();
        const ss = SpreadsheetApp.openById(DB_SHEET_ID);
        const invoiceSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
        const sectionSheet = ss.getSheetByName(SECTIONS_SHEET_NAME);

        if (!invoiceSheet || !sectionSheet) {
            throw new Error("Invoicing sheets are not available.");
        }

        // Get invoice data
        const invoiceData = invoiceSheet.getDataRange().getValues();
        if (invoiceData.length < 2) return { ok: true, invoices: [] };
        const invoiceHeaders = invoiceData.shift();

        const invoices = invoiceData.map(row => {
            const invoice = {};
            invoiceHeaders.forEach((header, i) => {
                invoice[header] = row[i];
            });
            return invoice;
        });

        // Get section data to enrich invoices
        const sectionData = sectionSheet.getDataRange().getValues();
        const sectionHeaders = sectionData.shift();
        const sections = sectionData.reduce((acc, row) => {
            const section = {};
            sectionHeaders.forEach((header, i) => {
                section[header] = row[i];
            });
            acc[section.seksjonsnummer] = section;
            return acc;
        }, {});

        // Combine data
        const enrichedInvoices = invoices.map(invoice => {
            const sectionInfo = sections[invoice.seksjonsnummer] || {};
            return {
                ...invoice,
                eierNavn: sectionInfo.eierNavn || 'N/A',
                eierEpost: sectionInfo.eierEpost || 'N/A'
            };
        }).sort((a, b) => new Date(b.opprettetDato) - new Date(a.opprettetDato));

        return { ok: true, invoices: enrichedInvoices };
    } catch (e) {
        return { ok: false, message: `Server error: ${e.message}` };
    }
}

/**
 * Generates invoices for all sections from a template.
 * @param {object} payload The generation request data, including templateId and dueDate.
 * @returns {object} A response object indicating success or failure.
 */
function generateInvoicesFromTemplate(payload) {
    try {
        _validateConfig();
        const { templateId, forfallsdato } = payload;
        if (!templateId || !forfallsdato) {
            throw new Error("Template ID and due date are required.");
        }

        const ss = SpreadsheetApp.openById(DB_SHEET_ID);
        const templateSheet = ss.getSheetByName(INVOICE_TEMPLATES_SHEET_NAME);
        const sectionSheet = ss.getSheetByName(SECTIONS_SHEET_NAME);
        const invoiceSheet = ss.getSheetByName(INVOICES_SHEET_NAME);

        // Find the template
        const templateData = templateSheet.getDataRange().getValues();
        const templateRow = templateData.find(row => row[0] == templateId);
        if (!templateRow) throw new Error("Template not found.");
        const template = { id: templateRow[0], navn: templateRow[1], belop: parseFloat(templateRow[2]), beskrivelse: templateRow[3] };


        // Get all sections
        const sectionData = sectionSheet.getDataRange().getValues();
        sectionData.shift(); // Remove headers

        const newInvoices = [];
        const today = new Date();

        sectionData.forEach(row => {
            const section = { seksjonsnummer: row[0], andel: parseFloat(row[3] || 1) };
            const finalAmount = template.belop * section.andel;

            newInvoices.push([
                Utilities.getUuid(),
                section.seksjonsnummer,
                finalAmount,
                new Date(forfallsdato),
                'Ubetalt',
                today,
                `${template.navn} - ${template.beskrivelse}`,
                0
            ]);
        });

        if (newInvoices.length > 0) {
            invoiceSheet.getRange(invoiceSheet.getLastRow() + 1, 1, newInvoices.length, newInvoices[0].length).setValues(newInvoices);
        }

        return { ok: true, count: newInvoices.length };
    } catch (e) {
        return { ok: false, message: `Server error: ${e.message}` };
    }
}

/**
 * Registers a payment for a specific invoice.
 * @param {object} payload The payment data, including invoiceId and amount.
 * @returns {object} A response object indicating success or failure.
 */
function registerPayment(payload) {
    try {
        _validateConfig();
        const { invoiceId, amount } = payload;
        if (!invoiceId || amount === undefined) {
            throw new Error("Invoice ID and amount are required.");
        }

        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICES_SHEET_NAME);
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const rowIndex = data.findIndex(row => row[0] == invoiceId);

        if (rowIndex === -1) throw new Error("Invoice not found.");

        const row = data[rowIndex];
        const totalAmount = parseFloat(row[headers.indexOf('belop')]);
        let paidAmount = parseFloat(row[headers.indexOf('betaltBelop')] || 0);
        paidAmount += parseFloat(amount);

        let status = 'Delvis betalt';
        if (paidAmount >= totalAmount) {
            status = 'Betalt';
            paidAmount = totalAmount; // Ensure it does not exceed the total
        }

        sheet.getRange(rowIndex + 2, headers.indexOf('betaltBelop') + 1).setValue(paidAmount);
        sheet.getRange(rowIndex + 2, headers.indexOf('status') + 1).setValue(status);

        return { ok: true, newStatus: status };
    } catch (e) {
        return { ok: false, message: `Server error: ${e.message}` };
    }
}

/**
 * Saves an invoice template (creates or updates).
 * @param {object} payload The template data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function saveInvoiceTemplate(payload) {
    try {
        _validateConfig();
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(INVOICE_TEMPLATES_SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${INVOICE_TEMPLATES_SHEET_NAME}" not found.`);

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        if (payload.id) {
            // Update existing template
            const data = sheet.getDataRange().getValues();
            const rowIndex = data.findIndex(row => row[0] == payload.id);
            if (rowIndex > 0) {
                const rowData = data[rowIndex];
                const newRow = headers.map((header, i) => payload[header] !== undefined ? payload[header] : rowData[i]);
                sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
            } else {
                throw new Error(`Template with ID ${payload.id} not found.`);
            }
        } else {
            // Create new template
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