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
const POSTS_SHEET_NAME = 'Posts';
const ATTACHMENTS_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; // Replace with the ID of the Google Drive folder for attachments

/**
 * Validates that the script has been configured.
 * @private
 */
function _validateConfig() {
  if (DB_SHEET_ID.startsWith('YOUR_') || ATTACHMENTS_FOLDER_ID.startsWith('YOUR_')) {
    throw new Error('Script not configured. Please follow SETUP_INSTRUCTIONS.md.');
  }
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
 * Helper to get a sheet by name, or create it if it doesn't exist.
 * @param {string} name The name of the sheet.
 * @param {Array<string>} headers The headers to set if the sheet is created.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function _getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Retrieves social feed posts from the 'Posts' sheet.
 * SN.10.1
 * @returns {object} A response object with a list of posts.
 */
function social_getPosts() {
  try {
    _validateConfig();
    const headers = ['id', 'author', 'content', 'timestamp'];
    const sheet = _getOrCreateSheet(POSTS_SHEET_NAME, headers);

    const data = sheet.getDataRange().getValues();
    const headerRow = data.shift() || [];

    // Create a map of header names to column indices for robustness
    const headerMap = headerRow.reduce((map, header, i) => {
      map[header] = i;
      return map;
    }, {});

    const posts = data.map(row => ({
      id: row[headerMap['id']],
      author: row[headerMap['author']],
      content: row[headerMap['content']],
      timestamp: row[headerMap['timestamp']],
    })).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Sort newest first

    return { ok: true, posts: posts };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/**
 * Adds a new post to the 'Posts' sheet.
 * SN.10.1
 * @param {object} payload The post data from the client (e.g., { content: "..." }).
 * @returns {object} A response object indicating success or failure.
 */
function social_addPost(payload) {
  try {
    _validateConfig();
    const headers = ['id', 'author', 'content', 'timestamp'];
    const sheet = _getOrCreateSheet(POSTS_SHEET_NAME, headers);

    const userEmail = Session.getActiveUser().getEmail();

    const newPost = {
      id: Utilities.getUuid(),
      author: userEmail,
      content: payload.content,
      timestamp: new Date().toISOString()
    };

    // Append the new post in the correct order based on headers
    const newRow = headers.map(header => newPost[header]);
    sheet.appendRow(newRow);

    return { ok: true, newPost: newPost };
  } catch (e) {
    return { ok: false, message: `Server error: ${e.message}` };
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