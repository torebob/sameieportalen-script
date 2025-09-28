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