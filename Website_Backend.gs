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
            return HtmlService.createHtmlOutput('Page ID is required for editing.');
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
    return { ok: false, message: 'Invalid password' };
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
        if (h === 'title') return `Ny Side (${pageId})`; // Default title
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