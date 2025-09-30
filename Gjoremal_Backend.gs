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
 * Updates the status of a single task.
 * @param {string} taskId The ID of the task to update.
 * @param {string} newStatus The new status ('Open' or 'Completed').
 * @returns {object} A response object indicating success or failure.
 */
function gjoremalUpdateStatus(taskId, newStatus) {
  try {
    _validateConfig();
    if (!taskId || !newStatus) {
      throw new Error("Task ID and new status are required.");
    }
    if (newStatus !== 'Open' && newStatus !== 'Completed') {
        throw new Error("Invalid status provided.");
    }

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(TASKS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${TASKS_SHEET_NAME}" not found.`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColumnIndex = headers.indexOf('id');
    const statusColumnIndex = headers.indexOf('status');

    if (idColumnIndex === -1 || statusColumnIndex === -1) {
      throw new Error("Required columns ('id', 'status') not found in the Tasks sheet.");
    }

    const rowIndexToUpdate = data.findIndex((row, index) => index > 0 && row[idColumnIndex] == taskId);

    if (rowIndexToUpdate !== -1) {
      // +1 for 1-based index, +1 for header row
      sheet.getRange(rowIndexToUpdate + 1, statusColumnIndex + 1).setValue(newStatus);
      return { ok: true };
    } else {
      return { ok: false, message: `Task with ID ${taskId} not found.` };
    }
  } catch (e) {
    console.error(`Error in gjoremalUpdateStatus: ${e.message}`);
    return { ok: false, message: e.message };
  }
}

/**
 * Deletes a task by its ID.
 * @param {string} taskId The ID of the task to delete.
 * @returns {object} A response object indicating success or failure.
 */
function gjoremalDeleteTask(taskId) {
  try {
    _validateConfig();
    if (!taskId) {
      throw new Error("Task ID is required for deletion.");
    }

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(TASKS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${TASKS_SHEET_NAME}" not found.`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColumnIndex = headers.indexOf('id');

    if (idColumnIndex === -1) {
      throw new Error("Column 'id' not found in the Tasks sheet.");
    }

    const rowIndexToDelete = data.findIndex((row, index) => index > 0 && row[idColumnIndex] == taskId);

    if (rowIndexToDelete !== -1) {
      sheet.deleteRow(rowIndexToDelete + 1);
      return { ok: true };
    } else {
      return { ok: false, message: `Task with ID ${taskId} not found.` };
    }
  } catch (e) {
    console.error(`Error in gjoremalDeleteTask: ${e.message}`);
    return { ok: false, message: e.message };
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