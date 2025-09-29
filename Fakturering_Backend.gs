/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file access for this script
 * to only the current document. This is a best practice for security.
 */

// --- CONFIGURATION ---
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
 * Retrieves all invoice templates.
 * @returns {object} A response object with the list of templates.
 */
function getInvoiceTemplates() {
    try {
        _createInvoicingSheetsIfNotExist();
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID_FAKTURA).getSheetByName(INVOICE_TEMPLATES_SHEET_NAME);
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
 * Saves an invoice template (creates or updates).
 * @param {object} payload The template data from the client.
 * @returns {object} A response object indicating success or failure.
 */
function saveInvoiceTemplate(payload) {
    try {
        _createInvoicingSheetsIfNotExist();
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID_FAKTURA).getSheetByName(INVOICE_TEMPLATES_SHEET_NAME);
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
 * Retrieves all invoices, along with owner details from the Sections sheet.
 * @returns {object} A response object with the list of invoices.
 */
function getInvoices() {
    try {
        _createInvoicingSheetsIfNotExist();
        const ss = SpreadsheetApp.openById(DB_SHEET_ID_FAKTURA);
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
        _createInvoicingSheetsIfNotExist();
        const { templateId, forfallsdato } = payload;
        if (!templateId || !forfallsdato) {
            throw new Error("Template ID and due date are required.");
        }

        const ss = SpreadsheetApp.openById(DB_SHEET_ID_FAKTURA);
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
        _createInvoicingSheetsIfNotExist();
        const { invoiceId, amount } = payload;
        if (!invoiceId || amount === undefined) {
            throw new Error("Invoice ID and amount are required.");
        }

        const sheet = SpreadsheetApp.openById(DB_SHEET_ID_FAKTURA).getSheetByName(INVOICES_SHEET_NAME);
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
 * Sends an invoice email to the resident.
 * @param {string} invoiceId The ID of the invoice to send.
 * @returns {object} A response object indicating success or failure.
 */
function sendInvoiceEmail(invoiceId) {
    try {
        _createInvoicingSheetsIfNotExist();
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
        _createInvoicingSheetsIfNotExist();
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