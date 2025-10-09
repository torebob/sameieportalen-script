// --- CONFIGURATION ---
// IMPORTANT: Set these properties in your script's properties.
// Go to Project Settings > Script Properties and add the following keys:
// 1. GEMINI_API_KEY: Your API key for the Google AI for Developers.
// 2. DRIVE_FOLDER_ID: The ID of the Google Drive folder containing your documentation for the Q&A agent.

// =================================================================================
// FASE 1: PROSESSERING AV NYE HENVENDELSER (fra ny-henvendelse-page.html)
// =================================================================================

/**
 * Tar imot en henvendelse, analyserer den med AI og oppretter en oppgave i regnearket.
 * @param {Object} formData Objekt med { description } fra klientsiden.
 * @returns {Object} Et suksessobjekt med den nye oppgave-IDen.
 */
function processHenvendelse(formData) {
  try {
    if (!formData || !formData.description) {
      throw new Error("Mangler beskrivelse i henvendelsen.");
    }
    Logger.log(`Mottok ny henvendelse: ${formData.description}`);

    const prompt = `
      Du er en hjelpsom assistent for et styre i et boligsameie.
      Analyser følgende henvendelse fra en beboer.
      Ditt mål er å konvertere den til en strukturert oppgave.
      Teksten er: "${formData.description}"

      Vennligst returner kun et gyldig JSON-objekt med følgende struktur:
      {
        "title": "En kort, beskrivende tittel for oppgaven (maks 10 ord)",
        "category": "Velg den mest passende kategorien fra denne listen: [Vedlikehold, Renhold, Økonomi, Generelt, Annet]",
        "priority": "Vurder hastegraden og velg én: [Høy, Middels, Lav]"
      }
    `;

    const aiResponse = _callGeminiApi(prompt);
    const taskData = JSON.parse(aiResponse);
    Logger.log(`AI-analyse resultat: ${JSON.stringify(taskData)}`);

    const oppgaverSheet = _getSheetByName_('Oppgaver'); // Antar at _getSheetByName_ finnes i et annet script
    if (!oppgaverSheet) {
      throw new Error('Fant ikke regneark-arket "Oppgaver".');
    }

    const newTaskId = `OPPG-${new Date().toISOString().substring(0, 10)}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`;
    const newRow = [
      newTaskId,
      taskData.title || "Ny henvendelse",
      formData.description,
      taskData.category || 'Annet',
      taskData.priority || 'Middels',
      new Date(), '','Ny','','',
    ];

    oppgaverSheet.appendRow(newRow);
    Logger.log(`Opprettet ny oppgave ${newTaskId} i arket "Oppgaver".`);
    
    return { success: true, taskId: newTaskId };

  } catch (e) {
    Logger.log(`FEIL i processHenvendelse: ${e.message}\n${e.stack}`);
    throw new Error(`Kunne ikke prosessere henvendelsen. Feil: ${e.message}`);
  }
}

// =================================================================================
// FASE 2: SPØRR KUNNSKAPSBASEN (fra kunnskapsbase-assistent-page.html)
// =================================================================================

/**
 * Hovedfunksjon kalt fra frontend for å stille et spørsmål til AI-en.
 * @param {string} question Brukerens spørsmål.
 * @returns {string} Svaret fra AI-en.
 */
function askAI(question) {
  try {
    const context = getDocumentationContext();
    if (!context) {
      return 'Ingen dokumentasjon funnet. Sjekk at riktig mappe-ID er angitt og at mappen inneholder lesbare dokumenter.';
    }

    const prompt = `Basert på følgende dokumentasjon, svar på spørsmålet. Svar kun basert på informasjonen som er gitt. Hvis svaret ikke finnes i dokumentasjonen, si "Jeg fant ikke svaret i dokumentasjonen."\n\nDokumentasjon:\n---\n${context}\n---\n\nSpørsmål: ${question}`;
    
    const answer = _callGeminiApi(prompt);
    return answer;

  } catch (e) {
    Logger.log('Error in askAI: ' + e.toString());
    if (e.message.includes('is not set in Script Properties')) {
      return 'Konfigurasjon mangler. En administrator må angi API-nøkkel og mappe-ID i innstillingene før agenten kan brukes.';
    }
    return `En feil oppstod: ${e.message}`;
  }
}

/**
 * Henter og slår sammen tekstinnhold fra alle støttede filer i en spesifisert Google Drive-mappe.
 * @returns {string} Den kombinerte teksten fra alle dokumentene.
 */
function getDocumentationContext() {
  const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
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
        Logger.log('Hopper over filtype som ikke støttes: ' + file.getName() + ' (' + mimeType + ')');
        continue;
      }
      context += `--- DOC: ${file.getName()} ---\n${text}\n\n`;
    } catch (e) {
      Logger.log(`Kunne ikke behandle filen "${file.getName()}": ${e.message}`);
    }
  }
  return context;
}

// =================================================================================
// ADMIN-FUNKSJONER FOR INNSTILLINGER
// =================================================================================

/**
 * Lagrer innstillingene for Q&A-agenten.
 * @param {object} settings Et objekt som inneholder GEMINI_API_KEY og DRIVE_FOLDER_ID.
 * @returns {object} Et suksess- eller feilobjekt.
 */
function saveQnASettings(settings) {
  try {
    // Antar at det finnes en funksjon _getCurrentUserRoles() eller lignende
    // For nå, kommenterer vi ut sikkerhetssjekken for enkelhets skyld.
    // const userRoles = _getCurrentUserRoles(); 
    // if (!userRoles.includes('Admin')) {
    //   return { success: false, error: 'Bare administratorer kan endre innstillinger.' };
    // }

    PropertiesService.getScriptProperties().setProperties({
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
 * Henter innstillingene for Q&A-agenten.
 * @returns {object} Et objekt med innstillingene.
 */
function getQnASettings() {
  try {
    // const userRoles = _getCurrentUserRoles();
    // if (!userRoles.includes('Admin')) {
    //   throw new Error('Du har ikke tilgang til å se disse innstillingene.');
    // }
    const properties = PropertiesService.getScriptProperties().getProperties();
    return {
      GEMINI_API_KEY: properties.GEMINI_API_KEY || '',
      DRIVE_FOLDER_ID: properties.DRIVE_FOLDER_ID || ''
    };
  } catch (e) {
    Logger.log('Error in getQnASettings: ' + e.toString());
    throw new Error('En feil oppstod ved henting av innstillinger: ' + e.message);
  }
}

// =================================================================================
// SENTRAL HJELPEFUNKSJON FOR API-KALL
// =================================================================================

/**
 * Intern funksjon for å kalle Gemini API.
 * @param {string} prompt Teksten som skal sendes til AI-en.
 * @returns {string} Svaret fra AI-en som ren tekst.
 * @private
 */
function _callGeminiApi(prompt) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    throw new Error("GEMINI_API_KEY er ikke satt i Script Properties.");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${API_KEY}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const json = JSON.parse(responseBody);
    if (json.candidates && json.candidates.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return "Kunne ikke generere et svar (muligens pga. innholdsfiltre).";
    }
  } else {
    Logger.log(`API-kall feilet med status ${responseCode}: ${responseBody}`);
    throw new Error(`AI-tjenesten svarte med en feil (Status: ${responseCode}). Sjekk loggen for detaljer.`);
  }
}
