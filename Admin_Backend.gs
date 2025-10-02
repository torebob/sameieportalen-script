/**
 * Initializes the website with default pages.
 * This function should only be run once.
 */
function initializeSite() {
  try {
    requireAuth(['admin']); // Only admins can initialize
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
    if (!sheet) {
        throw new Error("'WebsitePages'-arket finnes ikke. Sørg for at hovednettstedet er lastet inn minst én gang.");
    }

    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) {
        return { ok: true, message: "Nettstedet er allerede initialisert." };
    }

    const defaultPages = [
      ['home', 'Velkommen til Vårt Sameie', 'Dette er forsiden. Bruk redigeringsverktøyet til å endre denne teksten.', ''],
      ['about', 'Om Oss', 'Her kan dere skrive om sameiet, styret, og historien.', ''],
      ['rules', 'Husordensregler', 'Her legger dere inn husordensreglene.', ''],
      ['contact', 'Kontaktinformasjon', 'Styrets kontaktinformasjon kan legges inn her.', '']
    ];

    defaultPages.forEach(page => {
      sheet.appendRow(page);
    });

    logAuditEvent('INITIALIZE_SITE', 'System', { success: true });
    return { ok: true };
  } catch (e) {
    console.error("Error in initializeSite: " + e.message);
    return { ok: false, message: e.message };
  }
}

// --- News Management Functions ---

function listNewsArticles() {
  try {
    requireAuth(['admin', 'board_member', 'board_leader']);
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
    requireAuth(['admin', 'board_member', 'board_leader']);
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    article.id = Utilities.getUuid();
    article.publishedDate = new Date().toISOString();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(h => article[h] || '');
    sheet.appendRow(newRow);
    logAuditEvent('ADD_NEWS', 'News', { articleId: article.id, title: article.title });
    return { ok: true, id: article.id };
  } catch(e) { return { ok: false, message: e.message }; }
}

function updateNewsArticle(article) {
  try {
    requireAuth(['admin', 'board_member', 'board_leader']);
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idIndex = headers.indexOf('id');
    const rowIndex = data.findIndex(row => row[idIndex] == article.id);
    if (rowIndex === -1) throw new Error("Artikkelen ble ikke funnet");

    const newRow = headers.map(h => article[h] || '');
    sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([newRow]);
    logAuditEvent('UPDATE_NEWS', 'News', { articleId: article.id });
    return { ok: true };
  } catch(e) { return { ok: false, message: e.message }; }
}

function deleteNewsArticle(articleId) {
  try {
    requireAuth(['admin', 'board_member', 'board_leader']);
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('News');
    const data = sheet.getDataRange().getValues();
    const idIndex = data[0].indexOf('id');
    const rowIndex = data.findIndex(row => row[idIndex] == articleId);
    if (rowIndex > 0) {
        sheet.deleteRow(rowIndex + 1);
        logAuditEvent('DELETE_NEWS', 'News', { articleId: articleId });
        return { ok: true };
    }
    return { ok: false, message: "Artikkelen ble ikke funnet" };
  } catch(e) { return { ok: false, message: e.message }; }
}

// --- Document Management Functions ---

function listDocuments() {
    try {
        requireAuth(['admin', 'board_member', 'board_leader']);
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
        requireAuth(['admin', 'board_member', 'board_leader']);
        if (!fileObject) throw new Error("Fildata mangler.");

        const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
        const decoded = Utilities.base64Decode(fileObject.base64, Utilities.Charset.UTF_8);
        const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.name);
        const file = folder.createFile(blob);

        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('Documents');
        const docId = Utilities.getUuid();
        sheet.appendRow([docId, title, file.getUrl(), description]);
        logAuditEvent('ADD_DOCUMENT', 'Documents', { documentId: docId, title: title });
        return { ok: true, id: docId };
    } catch(e) {
        console.error("Error in addDocument: " + e.message);
        return { ok: false, message: e.message };
    }
}

// --- Common Resource Management ---

function addResource(resource) {
    try {
        requireAuth(['admin', 'board_leader']);
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
        logAuditEvent('ADD_RESOURCE', 'CommonResources', { resourceId: id, name: resource.name });
        return { ok: true, id: id };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function deleteResource(resourceId) {
    try {
        requireAuth(['admin', 'board_leader']);
        const sheet = _getOrCreateSheet('CommonResources', ['id', 'name', 'description', 'maxBookingHours', 'price', 'cancellationDeadline']);
        const data = sheet.getDataRange().getValues();
        const idIndex = data[0].indexOf('id');
        const rowIndex = data.findIndex(row => row[idIndex] == resourceId);

        if (rowIndex > 0) {
            sheet.deleteRow(rowIndex + 1);
            logAuditEvent('DELETE_RESOURCE', 'CommonResources', { resourceId: resourceId });
            return { ok: true };
        }
        return { ok: false, message: "Ressurs ikke funnet" };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

/**
 * Lists all pages from the WebsitePages sheet.
 * @returns {object} An object containing the list of pages or an error.
 */
function listPages() {
    try {
        requireAuth(['admin', 'board_member', 'board_leader']);
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) {
            return { ok: true, pages: [] };
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
 * Sets or changes the password for a page.
 * @param {string} pageId The ID of the page.
 * @param {string} password The new password. An empty string removes the password.
 * @returns {object} A success or error object.
 */
function setPagePassword(pageId, password) {
    try {
        requireAuth(['admin', 'board_leader']);
        if (!pageId) throw new Error("Side-ID er påkrevd.");
        const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName('WebsitePages');
        if (!sheet) throw new Error("'WebsitePages'-arket ble ikke funnet.");

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const pageIdIndex = headers.indexOf('pageId');
        let passwordIndex = headers.indexOf('password');

        if (passwordIndex === -1) {
            sheet.getRange(1, headers.length + 1).setValue('password');
            passwordIndex = headers.length;
        }

        const rowIndex = data.findIndex(row => row[pageIdIndex] == pageId);

        if (rowIndex > 0) {
            sheet.getRange(rowIndex + 1, passwordIndex + 1).setValue(password);
            logAuditEvent('SET_PAGE_PASSWORD', 'WebsitePages', { pageId: pageId });
            return { ok: true };
        } else {
            return { ok: false, message: "Siden ble ikke funnet." };
        }
    } catch (e) {
        console.error("Error in setPagePassword: " + e.message);
        return { ok: false, message: e.message };
    }
}