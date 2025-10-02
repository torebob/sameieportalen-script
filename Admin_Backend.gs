/**
 * Initializes the website with default pages.
 * This function should only be run once.
 */
function initializeSite() {
    try {
        const existingPages = DB.query('WebsitePages');
        if (existingPages.length > 0) {
            return { ok: true, message: "Nettstedet er allerede initialisert." };
        }

        const defaultPages = [
            { pageId: 'home', title: 'Velkommen til Vårt Sameie', content: 'Dette er forsiden. Bruk redigeringsverktøyet til å endre denne teksten.', password: '' },
            { pageId: 'about', title: 'Om Oss', content: 'Her kan dere skrive om sameiet, styret, og historien.', password: '' },
            { pageId: 'rules', title: 'Husordensregler', content: 'Her legger dere inn husordensreglene.', password: '' },
            { pageId: 'contact', title: 'Kontaktinformasjon', content: 'Styrets kontaktinformasjon kan legges inn her.', password: '' }
        ];

        defaultPages.forEach(page => {
            DB.insert('WebsitePages', page);
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
    const articles = DB.query('News');
    return { ok: true, articles: articles.sort((a, b) => new Date(b.publishedDate) - new Date(a.publishedDate)) };
  } catch(e) { return { ok: false, message: e.message }; }
}

function addNewsArticle(article) {
  try {
    const user = getCurrentUser(); // Sikrer at forfatter er logget inn
    article.author = user.name;
    article.publishedDate = new Date().toISOString();

    const newArticle = DB.insert('News', article);
    return { ok: true, id: newArticle.id };
  } catch(e) { return { ok: false, message: e.message }; }
}

function updateNewsArticle(article) {
  try {
    DB.update('News', article.id, article);
    return { ok: true };
  } catch(e) { return { ok: false, message: e.message }; }
}

function deleteNewsArticle(articleId) {
  try {
    const success = DB.delete('News', articleId);
    if (success) {
        return { ok: true };
    }
    return { ok: false, message: "Artikkelen ble ikke funnet" };
  } catch(e) { return { ok: false, message: e.message }; }
}

// --- Document Management Functions ---

function listDocuments() {
    try {
        const documents = DB.query('Documents');
        return { ok: true, documents: documents };
    } catch(e) { return { ok: false, message: e.message }; }
}

function addDocument(fileObject, title, description) {
    try {
        if (!fileObject) throw new Error("Fildata mangler.");
        const user = getCurrentUser();

        const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
        const decoded = Utilities.base64Decode(fileObject.base64, Utilities.Charset.UTF_8);
        const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.name);
        const file = folder.createFile(blob);

        const newDoc = {
            title: title,
            description: description,
            url: file.getUrl(),
            uploadedBy: user.email,
            uploadedAt: new Date().toISOString()
        };

        const insertedDoc = DB.insert('Documents', newDoc);
        return { ok: true, id: insertedDoc.id };
    } catch(e) {
        console.error("Error in addDocument: " + e.message);
        return { ok: false, message: e.message };
    }
}

// --- Common Resource Management ---

function addResource(resource) {
    try {
        const newResource = DB.insert('CommonResources', resource);
        return { ok: true, id: newResource.id };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function deleteResource(resourceId) {
    try {
        const success = DB.delete('CommonResources', resourceId);
        if (success) {
            return { ok: true };
        }
        return { ok: false, message: "Ressurs ikke funnet" };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function deleteDocument(docId) {
    try {
        const doc = DB.getById('Documents', docId);

        if (doc && doc.url) {
            const fileId = doc.url.match(/id=([^&]+)/)[1];
            if (fileId) {
                DriveApp.getFileById(fileId).setTrashed(true);
            }
            DB.delete('Documents', docId);
            return { ok: true };
        }
        return { ok: false, message: "Dokument ikke funnet" };
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
        const pages = DB.query('WebsitePages');
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
        if (!pageId) throw new Error("Side-ID er påkrevd.");

        // The delete function in our provider uses the 'id' field by default.
        // For 'WebsitePages', the unique identifier is 'pageId'.
        // We need to fetch the item first to get its row index for deletion.
        // This is a limitation of the current SheetsProvider implementation.
        // A better implementation would allow specifying a key for deletion.
        const success = DB.delete('WebsitePages', pageId);

        if (success) {
            return { ok: true };
        } else {
            return { ok: false, message: "Siden ble ikke funnet." };
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
        if (!pageId) throw new Error("Side-ID er påkrevd.");

        // The update function has a workaround for non-'id' keys.
        // We pass the key in the data payload itself.
        const success = DB.update('WebsitePages', pageId, { pageId: pageId, password: password });

        if (success) {
            return { ok: true };
        } else {
            return { ok: false, message: "Siden ble ikke funnet." };
        }
    } catch (e) {
        console.error("Error in setPagePassword: " + e.message);
        return { ok: false, message: e.message };
    }
}