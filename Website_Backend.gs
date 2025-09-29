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
    try {
        const { resourceId, startTime, endTime, userName, userEmail } = bookingDetails;
        const start = new Date(startTime);
        const end = new Date(endTime);

        // --- Conflict Check ---
        const bookingsSheet = _getOrCreateSheet('Bookings', ['id', 'resourceId', 'startTime', 'endTime', 'userEmail', 'userName', 'createdAt']);
        const data = bookingsSheet.getDataRange().getValues();
        const headers = data.shift();
        const resourceIdIndex = headers.indexOf('resourceId');
        const startTimeIndex = headers.indexOf('startTime');
        const endTimeIndex = headers.indexOf('endTime');

        const conflictingBooking = data.find(row => {
            if (row[resourceIdIndex] !== resourceId) return false;
            const existingStart = new Date(row[startTimeIndex]);
            const existingEnd = new Date(row[endTimeIndex]);
            // Check for overlap: (StartA < EndB) and (EndA > StartB)
            return start < existingEnd && end > existingStart;
        });

        if (conflictingBooking) {
            return { ok: false, message: "Tiden er allerede booket. Vennligst velg en annen tid." };
        }

        // --- Create Booking ---
        const id = Utilities.getUuid();
        const createdAt = new Date().toISOString();
        bookingsSheet.appendRow([id, resourceId, startTime, endTime, userEmail, userName, createdAt]);

        // --- Get Resource Name for Email ---
        const resourceSheet = _getOrCreateSheet('CommonResources', ['id', 'name']);
        const resourceData = resourceSheet.getDataRange().getValues();
        const resourceHeaders = resourceData.shift();
        const resIdIndex = resourceHeaders.indexOf('id');
        const resourceNameIndex = resourceHeaders.indexOf('name');
        const resourceRow = resourceData.find(r => r[resIdIndex] === resourceId);
        const resourceName = resourceRow ? resourceRow[resourceNameIndex] : 'Ukjent Ressurs';

        // --- Send Confirmation Email ---
        const subject = "Booking bekreftelse";
        const body = `
            Hei ${userName},

            Din booking er bekreftet:
            Ressurs: ${resourceName}
            Starttid: ${start.toLocaleString('no-NO')}
            Sluttid: ${end.toLocaleString('no-NO')}

            Takk!
        `;
        // Using a try-catch for the email in case of permission issues,
        // so it doesn't block the booking itself.
        try {
            MailApp.sendEmail(userEmail, subject, body);
        } catch(e) {
            console.error("Kunne ikke sende bekreftelses-epost: " + e.message);
            // Don't fail the whole operation, just log the error.
        }

        return { ok: true, id: id };
    } catch (e) {
        return { ok: false, message: "En feil oppstod under bookingen: " + e.message };
    }
}

/**
 * Gets the raw HTML for the booking page.
 * @returns {string} The HTML content of Booking.html.
 */
function getBookingPageHtml() {
    return HtmlService.createHtmlOutputFromFile('Booking.html').getContent();
}