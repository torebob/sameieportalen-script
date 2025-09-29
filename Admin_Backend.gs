/**
 * Initializes the website with default pages.
 * This function should only be run once.
 */
function initializeSite() {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
    if (!sheet) {
        throw new Error("The 'WebsitePages' sheet does not exist. Please ensure the main website has been loaded at least once.");
    }

    // Check if pages already exist to prevent duplicates
    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) { // >1 to account for header row
        // throw new Error("Site appears to be already initialized.");
        // We can silently ignore instead of throwing an error, to make it more user friendly.
        return { ok: true, message: "Site already initialized." };
    }

    const defaultPages = [
      ['home', 'Velkommen til Vårt Sameie', 'Dette er forsiden. Bruk redigeringsverktøyet til å endre denne teksten.', ''],
      ['about', 'Om Oss', 'Her kan dere skrive om sameiet, styret, og historien.', ''],
      ['rules', 'Husordensregler', 'Her legger dere inn husordensreglene.', ''],
      ['contact', 'Kontaktinformasjon', 'Styrets kontaktinformasjon kan legges inn her.', '']
    ];

    // Append the default page data
    defaultPages.forEach(page => {
      sheet.appendRow(page);
    });

    return { ok: true };
  } catch (e) {
    console.error("Error in initializeSite: " + e.message);
    return { ok: false, message: e.message };
  }
}

// --- News Management Functions ---

function listNewsArticles() {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    if (!sheet) return { ok: true, articles: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const articles = data.map(row => {
        const article = {};
        headers.forEach((h, i) => article[h] = row[i]);
        return article;
    });
    return { ok: true, articles: articles.sort((a, b) => new Date(b.publishedDate) - new Date(a.publishedDate)) };
  } catch(e) { return { ok: false, message: e.message }; }
}

function addNewsArticle(article) {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    article.id = Utilities.getUuid();
    article.publishedDate = new Date().toISOString();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(h => article[h] || '');
    sheet.appendRow(newRow);
    return { ok: true, id: article.id };
  } catch(e) { return { ok: false, message: e.message }; }
}

function updateNewsArticle(article) {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idIndex = headers.indexOf('id');
    const rowIndex = data.findIndex(row => row[idIndex] == article.id);
    if (rowIndex === -1) throw new Error("Article not found");

    const newRow = headers.map(h => article[h] || '');
    sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([newRow]);
    return { ok: true };
  } catch(e) { return { ok: false, message: e.message }; }
}

function deleteNewsArticle(articleId) {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    const data = sheet.getDataRange().getValues();
    const idIndex = data[0].indexOf('id');
    const rowIndex = data.findIndex(row => row[idIndex] == articleId);
    if (rowIndex > 0) {
        sheet.deleteRow(rowIndex + 1);
        return { ok: true };
    }
    return { ok: false, message: "Article not found" };
  } catch(e) { return { ok: false, message: e.message }; }
}

// --- Document Management Functions ---

function listDocuments() {
    try {
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Documents');
        if (!sheet) return { ok: true, documents: [] };
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const documents = data.map(row => {
            const doc = {};
            headers.forEach((h, i) => doc[h] = row[i]);
            return doc;
        });
        return { ok: true, documents: documents };
    } catch(e) { return { ok: false, message: e.message }; }
}

function addDocument(fileObject, title, description) {
    try {
        if (!fileObject) throw new Error("File data is missing.");

        const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
        const decoded = Utilities.base64Decode(fileObject.base64, Utilities.Charset.UTF_8);
        const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.name);
        const file = folder.createFile(blob);

        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Documents');
        const docId = Utilities.getUuid();
        sheet.appendRow([docId, title, file.getUrl(), description]);

        return { ok: true, id: docId };
    } catch(e) {
        console.error("Error in addDocument: " + e.message);
        return { ok: false, message: e.message };
    }
}

// --- Common Resource Management ---

function addResource(resource) {
    try {
        const sheet = _getOrCreateSheet('CommonResources', ['id', 'name', 'description', 'maxBookingHours', 'price', 'cancellationDeadline']);
        const id = Utilities.getUuid();
        sheet.appendRow([
            id,
            resource.name,
            resource.description,
            resource.maxBookingHours || '',
            resource.price || '',
            resource.cancellationDeadline || ''
        ]);
        return { ok: true, id: id };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function deleteResource(resourceId) {
    try {
        const sheet = _getOrCreateSheet('CommonResources', ['id', 'name', 'description', 'maxBookingHours', 'price', 'cancellationDeadline']);
        const data = sheet.getDataRange().getValues();
        const idIndex = data[0].indexOf('id');
        const rowIndex = data.findIndex(row => row[idIndex] == resourceId);

        if (rowIndex > 0) {
            sheet.deleteRow(rowIndex + 1);
            return { ok: true };
        }
        return { ok: false, message: "Resource not found" };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function deleteDocument(docId) {
    try {
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
            sheet.deleteRow(rowIndex + 2); // +2 because of header and 0-based index
            return { ok: true };
        }
        return { ok: false, message: "Document not found" };
    } catch(e) {
        console.error("Error in deleteDocument: " + e.message);
        return { ok: false, message: e.message };
    }
}

/**
 * Lists all pages from the WebsitePages sheet.
 * @returns {object} An object containing the list of pages or an error.
 */
function listPages() {
    try {
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) {
            return { ok: true, pages: [] }; // No sheet means no pages
        }
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const pages = data.map(row => {
            const page = {};
            headers.forEach((header, i) => {
                page[header] = row[i];
            });
            return page;
        });
        return { ok: true, pages: pages };
    } catch (e) {
        console.error("Error in listPages: " + e.message);
        return { ok: false, message: e.message };
    }
}

/**
 * Deletes a page from the WebsitePages sheet.
 * @param {string} pageId The ID of the page to delete.
 * @returns {object} A success or error object.
 */
function deletePage(pageId) {
    try {
        if (!pageId) throw new Error("Page ID is required.");
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) throw new Error("'WebsitePages' sheet not found.");

        const data = sheet.getDataRange().getValues();
        const pageIdIndex = data[0].indexOf('pageId');
        if (pageIdIndex === -1) throw new Error("'pageId' column not found.");

        const rowIndex = data.findIndex(row => row[pageIdIndex] == pageId);

        if (rowIndex > 0) { // > 0 to not delete header
            sheet.deleteRow(rowIndex + 1);
            return { ok: true };
        } else {
            return { ok: false, message: "Page not found." };
        }
    } catch (e) {
        console.error("Error in deletePage: " + e.message);
        return { ok: false, message: e.message };
    }
}

/**
 * Sets or changes the password for a page.
 * @param {string} pageId The ID of the page.
 * @param {string} password The new password. An empty string removes the password.
 * @returns {object} A success or error object.
 */
function setPagePassword(pageId, password) {
    try {
        if (!pageId) throw new Error("Page ID is required.");
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) throw new Error("'WebsitePages' sheet not found.");

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const pageIdIndex = headers.indexOf('pageId');
        let passwordIndex = headers.indexOf('password');

        // If 'password' column doesn't exist, add it
        if (passwordIndex === -1) {
            sheet.getRange(1, headers.length + 1).setValue('password');
            passwordIndex = headers.length;
        }

        const rowIndex = data.findIndex(row => row[pageIdIndex] == pageId);

        if (rowIndex > 0) { // > 0 to not affect header
            sheet.getRange(rowIndex + 1, passwordIndex + 1).setValue(password);
            return { ok: true };
        } else {
            return { ok: false, message: "Page not found." };
        }
    } catch (e) {
        console.error("Error in setPagePassword: " + e.message);
        return { ok: false, message: e.message };
    }
}