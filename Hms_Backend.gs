/**
 * @OnlyCurrentDoc
 *
 * This script contains the backend logic for the HSE (Helse, Miljø og Sikkerhet) module.
 */

// --- CONFIGURATION ---
const HMS_RISK_ASSESSMENT_TEMPLATES_SHEET_NAME = 'HmsRiskAssessmentTemplates';
const HMS_RISK_ASSESSMENTS_SHEET_NAME = 'HmsRiskAssessments';
const HMS_CHECKLISTS_SHEET_NAME = 'HmsChecklists';
const HMS_DEVIATIONS_SHEET_NAME = 'HmsDeviations';
const HMS_ACTIVITY_LOG_SHEET_NAME = 'HmsActivityLog';

/**
 * Creates the necessary sheets for the HMS module if they don't already exist.
 * This function should be called from a setup or validation function.
 * @private
 */
function _createHmsSheetsIfNotExist() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);

  // Sheet for Risk Assessment Templates
  if (!ss.getSheetByName(HMS_RISK_ASSESSMENT_TEMPLATES_SHEET_NAME)) {
    const sheet = ss.insertSheet(HMS_RISK_ASSESSMENT_TEMPLATES_SHEET_NAME);
    sheet.appendRow(['id', 'name', 'description', 'itemsJson']);
    // Pre-populate with a playground template
    sheet.appendRow([
      'tpl_playground_01',
      'Risikovurdering Lekeplass',
      'Standard mal for årlig sjekk av lekeplassområdet.',
      JSON.stringify([
        { "id": "item_01", "area": "Generelt", "question": "Er området rent og ryddig, fritt for søppel og farlige gjenstander?", "risk": "", "measure": "" },
        { "id": "item_02", "area": "Husker", "question": "Er huskestativet stabilt og uten synlig rust eller skade?", "risk": "", "measure": "" },
        { "id": "item_03", "area": "Husker", "question": "Er kjettinger og seter hele og uten sprekker?", "risk": "", "measure": "" },
        { "id": "item_04", "area": "Sklie", "question": "Er sklien hel, uten sprekker eller skarpe kanter?", "risk": "", "measure": "" },
        { "id": "item_05", "area": "Sklie", "question": "Er underlaget ved enden av sklien støtdempende og tilstrekkelig?", "risk": "", "measure": "" },
        { "id": "item_06", "area": "Sandkasse", "question": "Er sanden ren og fri for fremmedlegemer?", "risk": "", "measure": "" }
      ])
    ]);
  }

  // Sheet for started Risk Assessments
  if (!ss.getSheetByName(HMS_RISK_ASSESSMENTS_SHEET_NAME)) {
    const sheet = ss.insertSheet(HMS_RISK_ASSESSMENTS_SHEET_NAME);
    sheet.appendRow(['id', 'templateId', 'areaName', 'status', 'createdAt', 'completedAt', 'assessmentJson']);
  }

  // Sheet for user-defined Checklists
  if (!ss.getSheetByName(HMS_CHECKLISTS_SHEET_NAME)) {
    const sheet = ss.insertSheet(HMS_CHECKLISTS_SHEET_NAME);
    sheet.appendRow(['id', 'name', 'description', 'frequency', 'itemsJson']);
  }

  // Sheet for Deviations/Incidents
  if (!ss.getSheetByName(HMS_DEVIATIONS_SHEET_NAME)) {
    const sheet = ss.insertSheet(HMS_DEVIATIONS_SHEET_NAME);
    sheet.appendRow(['id', 'timestamp', 'description', 'location', 'reportedBy', 'attachmentUrl', 'linkedTaskId']);
  }

  // Sheet for logging all HMS activities
  if (!ss.getSheetByName(HMS_ACTIVITY_LOG_SHEET_NAME)) {
    const sheet = ss.insertSheet(HMS_ACTIVITY_LOG_SHEET_NAME);
    sheet.appendRow(['id', 'timestamp', 'user', 'activityType', 'description']);
  }
}

/**
 * Logs an activity to the HMS Activity Log.
 * @param {string} activityType - The type of activity (e.g., 'RISK_ASSESSMENT_STARTED', 'DEVIATION_REPORTED').
 * @param {string} description - A detailed description of the activity.
 * @private
 */
function _hmsLogActivity(activityType, description) {
  try {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(HMS_ACTIVITY_LOG_SHEET_NAME);
    const user = Session.getActiveUser().getEmail(); // Get current user's email
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      user,
      activityType,
      description
    ]);
  } catch (e) {
    // If logging fails, we don't want to stop the main operation.
    // Log the error to the Apps Script console instead.
    console.error(`Failed to log HMS activity: ${e.message}`);
  }
}

// --- PUBLIC FUNCTIONS ---

/**
 * Fetches all available risk assessment templates.
 * Corresponds to requirement HMS.20.1.
 * @returns {object} A response object with the list of templates.
 */
function hmsGetRiskAssessmentTemplates() {
  try {
    _validateConfig();
    const templates = _getSheetData(HMS_RISK_ASSESSMENT_TEMPLATES_SHEET_NAME);
    return { ok: true, templates: templates };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke hente maler for risikovurdering: ${e.message}` };
  }
}

/**
 * Starts a new risk assessment based on a template.
 * Corresponds to requirement HMS.20.1.
 * @param {string} templateId - The ID of the template to use.
 * @param {string} areaName - The specific area being assessed (e.g., "Lekeplass ved blokk A").
 * @returns {object} A response object with the ID of the new assessment.
 */
function hmsStartRiskAssessment(templateId, areaName) {
  try {
    _validateConfig();
    if (!templateId || !areaName) {
      throw new Error("Mal-ID og områdenavn er påkrevd.");
    }

    const templateSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(HMS_RISK_ASSESSMENT_TEMPLATES_SHEET_NAME);
    const templateData = templateSheet.getDataRange().getValues();
    const templateRow = templateData.find(row => row[0] === templateId);

    if (!templateRow) {
      throw new Error(`Mal med ID ${templateId} ble ikke funnet.`);
    }

    const assessmentSheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(HMS_RISK_ASSESSMENTS_SHEET_NAME);
    const newId = Utilities.getUuid();
    const assessmentJson = templateRow[3]; // itemsJson from template

    assessmentSheet.appendRow([
      newId,
      templateId,
      areaName,
      'Started', // Initial status
      new Date(),
      '', // completedAt is empty
      assessmentJson
    ]);

    _hmsLogActivity('RISK_ASSESSMENT_STARTED', `Started new risk assessment '${templateRow[1]}' for area '${areaName}'.`);

    return { ok: true, newAssessmentId: newId };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke starte risikovurdering: ${e.message}` };
  }
}

/**
 * Fetches all inspection checklists.
 * Corresponds to requirement HMS.20.2.
 * @returns {object} A response object with the list of checklists.
 */
function hmsGetChecklists() {
  try {
    _validateConfig();
    const checklists = _getSheetData(HMS_CHECKLISTS_SHEET_NAME);
    return { ok: true, checklists: checklists };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke hente sjekklister: ${e.message}` };
  }
}

/**
 * Saves a checklist (creates or updates).
 * Corresponds to requirement HMS.20.2.
 * @param {object} checklistData - The checklist object to save. It should include an 'id' for updates.
 * @returns {object} A response object indicating success or failure.
 */
function hmsSaveChecklist(checklistData) {
  try {
    _validateConfig();
    const result = saveRecord_(HMS_CHECKLISTS_SHEET_NAME, checklistData);
    const activity = checklistData.id ? 'CHECKLIST_UPDATED' : 'CHECKLIST_CREATED';
    _hmsLogActivity(activity, `Saved checklist '${checklistData.name}'.`);
    return result;
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke lagre sjekkliste: ${e.message}` };
  }
}

/**
 * Deletes a checklist.
 * Corresponds to requirement HMS.20.2.
 * @param {string} checklistId - The ID of the checklist to delete.
 * @returns {object} A response object indicating success or failure.
 */
function hmsDeleteChecklist(checklistId) {
  try {
    _validateConfig();
    // We need the name for the log before deleting
    const checklist = getRecordById_(HMS_CHECKLISTS_SHEET_NAME, checklistId);
    const result = deleteRecord_(HMS_CHECKLISTS_SHEET_NAME, checklistId);
    if (result.ok && checklist) {
       _hmsLogActivity('CHECKLIST_DELETED', `Deleted checklist '${checklist.name}'.`);
    }
    return result;
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke slette sjekkliste: ${e.message}` };
  }
}

/**
 * Registers a new deviation/incident.
 * Corresponds to requirements HMS.20.3 and HMS.20.4.
 * It also handles file attachments and creates a linked task in the "Gjøremål" system.
 * @param {object} deviationData - The deviation data from the client, including an optional 'attachment' object.
 * @returns {object} A response object indicating success or failure.
 */
function hmsRegisterDeviation(deviationData) {
  try {
    _validateConfig();
    const { description, location, attachment } = deviationData;
    if (!description || !location) {
      throw new Error("Beskrivelse og sted er påkrevd for å registrere et avvik.");
    }

    // 1. Handle file attachment
    let attachmentUrl = '';
    if (attachment && attachment.base64) {
      const { base64, mimeType, name } = attachment;
      const decoded = Utilities.base64Decode(base64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, mimeType, name);
      const folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
      const file = folder.createFile(blob);
      attachmentUrl = file.getUrl();
    }

    // 2. Create a task in the Gjøremål system
    const taskPayload = {
      beskrivelse: `HMS-AVVIK: ${description}`,
      kategori: 'HMS',
      prioritet: 'High',
      status: 'Open',
      opprettet_av: Session.getActiveUser().getEmail(),
      notater: `Avvik registrert på lokasjon: ${location}. Se HMS-modul for detaljer.`
    };
    const taskResult = gjoremalSave(taskPayload); // Assuming gjoremalSave is available and returns { ok: true, id: '...' }
    if (!taskResult.ok) {
      throw new Error(`Klarte ikke å opprette tilknyttet oppgave: ${taskResult.message}`);
    }
    const linkedTaskId = taskResult.id;

    // 3. Save the deviation record
    const deviationPayload = {
      timestamp: new Date(),
      description: description,
      location: location,
      reportedBy: Session.getActiveUser().getEmail(),
      attachmentUrl: attachmentUrl,
      linkedTaskId: linkedTaskId
    };
    const deviationResult = saveRecord_(HMS_DEVIATIONS_SHEET_NAME, deviationPayload);

    // 4. Log the activity
    _hmsLogActivity('DEVIATION_REPORTED', `New deviation reported at '${location}': "${description}"`);

    return { ok: true, deviationId: deviationResult.id, linkedTaskId: linkedTaskId };

  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke registrere avvik: ${e.message}` };
  }
}

/**
 * Retrieves the history of all HMS-related activities.
 * Corresponds to requirement HMS.20.5.
 * @param {object} [filters] - Optional filters to apply to the history. (Not implemented yet)
 * @returns {object} A response object with the list of history items.
 */
function hmsGetHistory(filters) {
  try {
    _validateConfig();
    const history = _getSheetData(HMS_ACTIVITY_LOG_SHEET_NAME);
    // Sort by newest first
    const sortedHistory = history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { ok: true, history: sortedHistory };
  } catch (e) {
    Logger.log(e);
    return { ok: false, error: `Kunne ikke hente historikk: ${e.message}` };
  }
}

/**
 * Generates a summary report of all HMS activities.
 * Corresponds to requirement HMS.20.6.
 * @returns {object} A response object containing the report data.
 */
function hmsGenerateReport() {
    try {
        _validateConfig();
        const deviations = _getSheetData(HMS_DEVIATIONS_SHEET_NAME);
        const assessments = _getSheetData(HMS_RISK_ASSESSMENTS_SHEET_NAME);
        const log = _getSheetData(HMS_ACTIVITY_LOG_SHEET_NAME).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

        const report = {
            deviations: deviations,
            assessments: assessments,
            log: log
        };

        return { ok: true, report: report };

    } catch (e) {
        Logger.log(e);
        return { ok: false, error: `Kunne ikke generere rapport: ${e.message}` };
    }
}

// --- HELPER FUNCTIONS ---

/**
 * A generic helper function to fetch all data from a given sheet.
 * @private
 * @param {string} sheetName - The name of the sheet to read.
 * @returns {Array<Object>} An array of objects representing the rows.
 */
function _getSheetData(sheetName) {
  _validateConfig();
  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Arket "${sheetName}" ble ikke funnet.`);
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift();
  return data.map(row => {
    const record = {};
    headers.forEach((header, i) => {
      record[header] = row[i];
    });
    return record;
  });
}

/**
 * A generic helper function to get a single record by its ID from a sheet.
 * @private
 * @param {string} sheetName - The name of the sheet.
 * @param {string} id - The ID of the record to find.
 * @returns {object|null} The record object or null if not found.
 */
function getRecordById_(sheetName, id) {
    const data = _getSheetData(sheetName);
    return data.find(item => item.id === id) || null;
}


/**
 * Generic function to save a record (create or update) to a specified sheet.
 * @private
 * @param {string} sheetName - The name of the sheet.
 * @param {object} record - The data object to save.
 * @returns {object} A response object indicating success or failure.
 */
function saveRecord_(sheetName, record) {
    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
    if (!sheet) throw new Error(`Arket "${sheetName}" ble ikke funnet.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (record.id) {
      // Update
      const data = sheet.getDataRange().getValues();
      const rowIndex = data.findIndex(row => row[0] == record.id);

      if (rowIndex > 0) {
        const newRow = headers.map(header => record[header] !== undefined ? record[header] : data[rowIndex][headers.indexOf(header)]);
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
      } else {
        throw new Error(`Post med ID ${record.id} ble ikke funnet i ${sheetName}.`);
      }
    } else {
      // Create
      record.id = Utilities.getUuid();
      const newRow = headers.map(header => record[header] !== undefined ? record[header] : '');
      sheet.appendRow(newRow);
    }

    return { ok: true, id: record.id };
}

/**
 * Generic function to delete a record from a specified sheet.
 * @private
 * @param {string} sheetName - The name of the sheet.
 * @param {string} id - The ID of the record to delete.
 * @returns {object} A response object indicating success or failure.
 */
function deleteRecord_(sheetName, id) {
    if (!id) throw new Error("ID er påkrevd for sletting.");

    const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(sheetName);
    if (!sheet) throw new Error(`Arket "${sheetName}" ble ikke funnet.`);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] == id);

    if (rowIndex > 0) {
        sheet.deleteRow(rowIndex + 1);
        return { ok: true };
    } else {
      return { ok: false, error: `Post med ID ${id} ble ikke funnet i ${sheetName}.` };
    }
}