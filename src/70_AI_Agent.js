// --- CONFIGURATION ---
// IMPORTANT: Set these properties in your script's properties.
// Go to Project Settings > Script Properties and add the following keys:
// 1. GEMINI_API_KEY: Your API key for the Google AI for Developers.
// 2. DRIVE_FOLDER_ID: The ID of the Google Drive folder containing your documentation.

/**
 * Main function called from the frontend to ask a question to the AI.
 * @param {string} question The user's question.
 * @returns {object} An object with either an 'answer' or 'error' key.
 */
function askAI(question) {
  try {
    const context = getDocumentationContext();
    if (!context) {
      return { error: 'Ingen dokumentasjon funnet. Sjekk at riktig mappe-ID er angitt i innstillingene og at mappen inneholder lesbare dokumenter.' };
    }

    const answer = queryGemini(context, question);
    return { answer: answer };

  } catch (e) {
    Logger.log('Error in askAI: ' + e.toString());
    // Provide a user-friendly error if configuration is missing.
    if (e.message.includes('is not set in Script Properties')) {
      return { error: 'Konfigurasjon mangler. En administrator må angi API-nøkkel og mappe-ID i Q&A-innstillingene før agenten kan brukes.' };
    }
    return { error: e.message };
  }
}

/**
 * Retrieves and concatenates the text content from all supported files
 * in the specified Google Drive folder.
 * @returns {string} The combined text from all documents.
 */
function getDocumentationContext() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const FOLDER_ID = scriptProperties.getProperty('DRIVE_FOLDER_ID');

  if (!FOLDER_ID) {
    throw new Error('DRIVE_FOLDER_ID is not set in Script Properties.');
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let context = '';

  while (files.hasNext()) {
    const file = files.next();
    try {
      let text = '';
      const mimeType = file.getMimeType();

      if (mimeType === MimeType.GOOGLE_DOCS) {
        text = DocumentApp.openById(file.getId()).getBody().getText();
      } else if (mimeType.startsWith('text/')) {
        text = file.getBlob().getDataAsString('UTF-8');
      } else {
        // You can add support for other file types like PDFs here if needed.
        // For now, we'll just log it.
        Logger.log('Skipping unsupported file type: ' + file.getName() + ' (' + mimeType + ')');
        continue;
      }

      context += `--- DOC: ${file.getName()} ---\n${text}\n\n`;

    } catch (e) {
      Logger.log(`Could not process file "${file.getName()}": ${e.message}`);
    }
  }

  return context;
}

/**
 * Sends a query to the Gemini API with the provided context and question.
 * @param {string} context The documentation text.
 * @param {string} question The user's question.
 * @returns {string} The AI-generated answer.
 */
function queryGemini(context, question) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const API_KEY = scriptProperties.getProperty('GEMINI_API_KEY');

  if (!API_KEY) {
    throw new Error('GEMINI_API_KEY is not set in Script Properties.');
  }

  const API_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${API_KEY}`;

  const prompt = `Basert på følgende dokumentasjon, svar på spørsmålet. Svar kun basert på informasjonen som er gitt. Hvis svaret ikke finnes i dokumentasjonen, si "Jeg fant ikke svaret i dokumentasjonen."\n\nDokumentasjon:\n---\n${context}\n---\n\nSpørsmål: ${question}`;

  const requestBody = {
    contents: [{
      parts: [{ text: prompt }]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true // Important to handle API errors gracefully
  };

  const response = UrlFetchApp.fetch(API_URL, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`API request failed with status ${responseCode}: ${responseBody}`);
  }

  const jsonResponse = JSON.parse(responseBody);

  // Safely navigate the response structure
  if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
      jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
      jsonResponse.candidates[0].content.parts.length > 0) {
    return jsonResponse.candidates[0].content.parts[0].text;
  } else {
    // This can happen if the content is filtered due to safety settings
    return "Kunne ikke generere et svar. Dette kan skyldes innholdsfiltre eller en uventet API-respons.";
  }
}

/**
 * Saves the Q&A agent settings.
 * @param {object} settings An object containing GEMINI_API_KEY and DRIVE_FOLDER_ID.
 * @returns {object} A success or error object.
 */
function saveQnASettings(settings) {
  try {
    const user = getCurrentUserInfo(); // Assumes this function exists and returns user info
    if (!user.isAdmin) {
      return { success: false, error: 'Bare administratorer kan endre innstillinger.' };
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperties({
      'GEMINI_API_KEY': settings.GEMINI_API_KEY,
      'DRIVE_FOLDER_ID': settings.DRIVE_FOLDER_ID
    });

    return { success: true };
  } catch (e) {
    Logger.log('Error in saveQnASettings: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Retrieves the Q&A agent settings.
 * Only returns the API key if the user is an admin.
 * @returns {object} An object with the settings.
 */
function getQnASettings() {
  try {
    const user = getCurrentUserInfo();
    if (!user.isAdmin) {
      throw new Error('Du har ikke tilgang til å se disse innstillingene.');
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    const properties = scriptProperties.getProperties();

    return {
      GEMINI_API_KEY: properties.GEMINI_API_KEY || '',
      DRIVE_FOLDER_ID: properties.DRIVE_FOLDER_ID || ''
    };
  } catch (e) {
    Logger.log('Error in getQnASettings: ' + e.toString());
    // Re-throw to be caught by the frontend's withFailureHandler
    throw new Error('En feil oppstod ved henting av innstillinger: ' + e.message);
  }
}