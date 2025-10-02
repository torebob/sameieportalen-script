/**
 * Main entry point for the web app.
 * @param {object} e The event parameter for a web app doGet request.
 * @returns {HtmlOutput} The HTML output for the page.
 */
function doGet(e) {
    const action = e.parameter.action;

    if (action === 'admin') {
        return HtmlService.createHtmlOutputFromFile('Admin_Panel')
            .setTitle('Admin Panel');
    }

    if (action === 'edit') {
        const pageId = e.parameter.page;
        if (!pageId) {
            return HtmlService.createHtmlOutput('Side-ID er påkrevd for redigering.');
        }
        const template = HtmlService.createTemplateFromFile('Edit_Page');
        template.pageId = pageId;
        return template.evaluate().setTitle(`Redigerer: ${pageId}`);
    }

    const page = e.parameter.page || 'home';
    const template = HtmlService.createTemplateFromFile('Website_Template');
    template.page = page;
    return template.evaluate()
        .setTitle('Sameiet Hjemmeside')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Gets the list of news articles.
 * @returns {Array<object>} A list of news articles.
 */
function getNewsFeed() {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    if (!sheet) {
      const newSheet = SpreadsheetApp.openById(DB_SHEET_ID).insertSheet('News');
      newSheet.appendRow(['id', 'title', 'content', 'publishedDate']);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    return data.map(row => {
      const article = {};
      headers.forEach((header, i) => article[header] = row[i]);
      return article;
    }).sort((a, b) => new Date(b.publishedDate) - new Date(a.publishedDate));
  } catch (e) {
    console.error("Error in getNewsFeed: " + e.message);
    return [];
  }
}

/**
 * Gets the list of documents.
 * @returns {Array<object>} A list of documents.
 */
function getDocuments() {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Documents');
    if (!sheet) {
      const newSheet = SpreadsheetApp.openById(DB_SHEET_ID).insertSheet('Documents');
      newSheet.appendRow(['id', 'title', 'url', 'description']);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    return data.map(row => {
      const doc = {};
      headers.forEach((header, i) => doc[header] = row[i]);
      return doc;
    });
  } catch (e) {
    console.error("Error in getDocuments: " + e.message);
    return [];
  }
}

/**
 * Gets the content for a specific page from the spreadsheet.
 * @param {string} pageId The ID of the page to retrieve.
 * @returns {object} The page content or null if not found.
 */
function getPageContent(pageId, password) {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
    if (!sheet) {
      const newSheet = SpreadsheetApp.openById(DB_SHEET_ID).insertSheet('WebsitePages');
      newSheet.appendRow(['pageId', 'title', 'content', 'password']);
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const pageIdIndex = headers.indexOf('pageId');
    const passwordIndex = headers.indexOf('password');

    for (const row of data) {
      if (row[pageIdIndex] === pageId) {
        const page = {};
        const pagePassword = row[passwordIndex];

        if (pagePassword && pagePassword !== password) {
          return { authRequired: true };
        }

        headers.forEach((header, i) => {
          if (header !== 'password') {
            page[header] = row[i];
          }
        });
        return page;
      }
    }
    return null;
  } catch (e) {
    console.error("Error in getPageContent: " + e.message);
    return null;
  }
}

function verifyPassword(pageId, password) {
    const pageContent = getPageContent(pageId, password);
    if (pageContent && !pageContent.authRequired) {
        return pageContent;
    }
    return { ok: false, message: 'Ugyldig passord' };
}

/**
 * Includes the content of another HTML file.
 * Used for including CSS and JS files in the main template.
 * @param {string} filename The name of the file to include.
 * @returns {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Saves the content of a page to the spreadsheet.
 * @param {string} pageId The ID of the page to save.
 * @param {string} content The new HTML content of the page.
 * @returns {object} A success or error object.
 */
function savePageContent(pageId, content) {
  try {
    requireAuth(['admin', 'board_member', 'board_leader']);
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
    if (!sheet) throw new Error("'WebsitePages' sheet not found.");

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const pageIdIndex = headers.indexOf('pageId');
    const contentIndex = headers.indexOf('content');

    let rowIndex = data.findIndex(row => row[pageIdIndex] === pageId);

    if (rowIndex !== -1) {
      // Update existing page
      sheet.getRange(rowIndex + 2, contentIndex + 1).setValue(content);
    } else {
      // Create new page
      const newRow = headers.map(h => {
        if (h === 'pageId') return pageId;
        if (h === 'content') return content;
        if (h === 'title') return `Ny side (${pageId})`; // Default title
        return '';
      });
      sheet.appendRow(newRow);
    }
    return { ok: true };
  } catch (e) {
    console.error("Error in savePageContent: " + e.message);
    return { ok: false, message: e.message };
  }
}


// --- Booking System Functions ---

function _getOrCreateSheet(sheetName, headers) {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(headers);
    }
    return sheet;
}

function listResources() {
    try {
        const sheet = _getOrCreateSheet('CommonResources', ['id', 'name', 'description', 'maxBookingHours', 'price', 'cancellationDeadline']);
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const resources = data.map(row => {
            const resource = {};
            headers.forEach((h, i) => resource[h] = row[i]);
            return resource;
        });
        return { ok: true, resources: resources };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function getBookings(resourceId, year, month) {
    try {
        requireAuth();
        const sheet = _getOrCreateSheet('Bookings', ['id', 'resourceId', 'startTime', 'endTime', 'userEmail', 'userName', 'createdAt']);
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const resourceIdIndex = headers.indexOf('resourceId');
        const startTimeIndex = headers.indexOf('startTime');

        const bookings = data.filter(row => {
            if (row[resourceIdIndex] !== resourceId) {
                return false;
            }
            const bookingDate = new Date(row[startTimeIndex]);
            // Filter by month and year to reduce data transfer
            return bookingDate.getFullYear() === year && bookingDate.getMonth() === month;
        }).map(row => {
            const booking = {};
            headers.forEach((h, i) => booking[h] = row[i]);
            return booking;
        });

        return { ok: true, bookings: bookings };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function createBooking(bookingDetails) {
    const lock = LockService.getScriptLock();

    try {
        // KRITISK: Autentiser bruker først
        const user = requireAuth(); // Fra Auth.gs

        lock.waitLock(30000); // Vent maks 30 sek

        const { resourceId, startTime, endTime } = bookingDetails;

        // VIKTIG: Bruk kun server-side brukerinfo - ALDRI stol på klient!
        const userName = user.name;
        const userEmail = user.email;

        const start = new Date(startTime);
        const end = new Date(endTime);

        // Validering
        if (!resourceId || !startTime || !endTime) {
            return { ok: false, message: "Alle felter er påkrevd" };
        }

        if (start >= end) {
            return { ok: false, message: "Starttid må være før sluttid" };
        }

        // Konfliktsjekk (nå innenfor lock)
        const bookingsSheet = _getOrCreateSheet('Bookings',
            ['id', 'resourceId', 'startTime', 'endTime', 'userEmail', 'userName', 'createdAt']
        );
        const data = bookingsSheet.getDataRange().getValues();
        const headers = data.shift();
        const resourceIdIndex = headers.indexOf('resourceId');
        const startTimeIndex = headers.indexOf('startTime');
        const endTimeIndex = headers.indexOf('endTime');

        const conflictingBooking = data.find(row => {
            if (row[resourceIdIndex] !== resourceId) return false;
            const existingStart = new Date(row[startTimeIndex]);
            const existingEnd = new Date(row[endTimeIndex]);
            return start < existingEnd && end > existingStart;
        });

        if (conflictingBooking) {
            return { ok: false, message: "Tiden er allerede booket. Vennligst velg en annen tid." };
        }

        // Opprett booking
        const id = Utilities.getUuid();
        const createdAt = new Date().toISOString();
        bookingsSheet.appendRow([id, resourceId, startTime, endTime, userEmail, userName, createdAt]);

        // Logger handling
        logAuditEvent('CREATE_BOOKING', 'Bookings', {
            bookingId: id,
            resourceId: resourceId
        });

        // Send e-post
        const resourceSheet = _getOrCreateSheet('CommonResources', ['id', 'name']);
        const resourceData = resourceSheet.getDataRange().getValues();
        const resourceHeaders = resourceData.shift();
        const resIdIndex = resourceHeaders.indexOf('id');
        const resourceNameIndex = resourceHeaders.indexOf('name');
        const resourceRow = resourceData.find(r => r[resIdIndex] === resourceId);
        const resourceName = resourceRow ? resourceRow[resourceNameIndex] : 'Ukjent Ressurs';

        try {
            MailApp.sendEmail(userEmail, "Booking bekreftelse", `
                Hei ${userName},

                Din booking er bekreftet:
                Ressurs: ${resourceName}
                Starttid: ${start.toLocaleString('no-NO')}
                Sluttid: ${end.toLocaleString('no-NO')}

                Takk!
            `);
        } catch(e) {
            console.error("Kunne ikke sende e-post: " + e.message);
        }

        return { ok: true, id: id };

    } catch (e) {
        console.error("Error in createBooking: " + e.message);
        console.error(e.stack);
        return { ok: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Gets the raw HTML for the booking page.
 * @returns {string} The HTML content of Booking.html.
 */
function getBookingPageHtml() {
    return HtmlService.createHtmlOutputFromFile('Booking.html').getContent();
}

// --- SIKKER DOKUMENT-SLETTING ---

function deleteDocument(docId) {
    try {
        // Kun admin eller styremedlemmer kan slette dokumenter
        requireAuth(['admin', 'board_member', 'board_leader']);

        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Documents');
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const idIndex = headers.indexOf('id');
        const urlIndex = headers.indexOf('url');

        const rowIndex = data.findIndex(row => row[idIndex] == docId);

        if (rowIndex !== -1) {
            const fileUrl = data[rowIndex][urlIndex];
            const fileId = fileUrl.match(/id=([^&]+)/)[1];
            if (fileId) {
                DriveApp.getFileById(fileId).setTrashed(true);
            }
            sheet.deleteRow(rowIndex + 2);

            // Logger sletting
            logAuditEvent('DELETE_DOCUMENT', 'Documents', { documentId: docId });

            return { ok: true };
        }
        return { ok: false, message: "Dokument ikke funnet" };
    } catch(e) {
        console.error("Error in deleteDocument: " + e.message);
        return { ok: false, message: e.message };
    }
}

// --- SIKKER SIDE-SLETTING ---

function deletePage(pageId) {
    try {
        // Kun admin kan slette sider
        requireAuth(['admin', 'board_leader']);

        if (!pageId) throw new Error("Side-ID er påkrevd.");
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) throw new Error("'WebsitePages'-arket ble ikke funnet.");

        const data = sheet.getDataRange().getValues();
        const pageIdIndex = data[0].indexOf('pageId');
        if (pageIdIndex === -1) throw new Error("'pageId'-kolonnen ble ikke funnet.");

        const rowIndex = data.findIndex(row => row[pageIdIndex] == pageId);

        if (rowIndex > 0) {
            sheet.deleteRow(rowIndex + 1);

            // Logger sletting
            logAuditEvent('DELETE_PAGE', 'WebsitePages', { pageId: pageId });

            return { ok: true };
        } else {
            return { ok: false, message: "Siden ble ikke funnet." };
        }
    } catch (e) {
        console.error("Error in deletePage: " + e.message);
        return { ok: false, message: e.message };
    }
}

// --- NYE GDPR-FUNKSJONER ---

/**
 * Eksporterer brukerens egne data (GDPR Art. 15 - Rett til innsyn)
 */
function exportMyData() {
    try {
        const user = getCurrentUser();

        // Hent brukerens data fra alle relevante sheets
        const myData = {
            profile: getUserInfo(user.email),
            bookings: getMyBookings(user.email),
            auditLog: getMyAuditLog(user.email)
        };

        return { ok: true, data: myData };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function getMyBookings(email) {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Bookings');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const emailIndex = headers.indexOf('userEmail');

    return data.filter(row => row[emailIndex] === email).map(row => {
        const booking = {};
        headers.forEach((h, i) => booking[h] = row[i]);
        return booking;
    });
}

function getMyAuditLog(email) {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('AuditLog');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const emailIndex = headers.indexOf('userEmail');

    return data.filter(row => row[emailIndex] === email).map(row => {
        const log = {};
        headers.forEach((h, i) => log[h] = row[i]);
        return log;
    });
}

/**
 * Sletter/anonymiserer brukerens data (GDPR Art. 17 - Rett til sletting)
 * MERK: Noe data må beholdes lovpålagt (økonomi i 5 år)
 */
function requestDataDeletion() {
    try {
        const user = getCurrentUser();

        // Anonymiser bookinger (ikke slett - nødvendig for statistikk)
        const bookingsSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Bookings');
        if (bookingsSheet) {
            const data = bookingsSheet.getDataRange().getValues();
            const headers = data[0];
            const emailIndex = headers.indexOf('userEmail');
            const nameIndex = headers.indexOf('userName');

            for (let i = 1; i < data.length; i++) {
                if (data[i][emailIndex] === user.email) {
                    bookingsSheet.getRange(i + 1, emailIndex + 1).setValue('anonymisert@slettet.local');
                    bookingsSheet.getRange(i + 1, nameIndex + 1).setValue('Anonymisert bruker');
                }
            }
        }

        // Logger sletting før bruker fjernes
        logAuditEvent('USER_DATA_DELETION', 'Users', { email: user.email });

        // Slett bruker fra Users-sheet
        const usersSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Users');
        if (usersSheet) {
            const data = usersSheet.getDataRange().getValues();
            const emailIndex = data[0].indexOf('email');
            const rowIndex = data.findIndex(row => row[emailIndex] === user.email);
            if (rowIndex > 0) {
                usersSheet.deleteRow(rowIndex + 1);
            }
        }

        return {
            ok: true,
            message: "Dine data er slettet/anonymisert. Du vil bli logget ut."
        };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}