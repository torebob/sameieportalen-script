/**
 * @OnlyCurrentDoc
 *
 * FILE: 22_Mail_Processor.js
 * VERSION: 1.0.0
 * AUTHOR: Jules
 * DATE: 2025-09-28
 *
 * DESCRIPTION:
 * Handles the processing of incoming emails for the AI assistant.
 * This includes fetching, categorizing, and logging emails.
 */

/**
 * Initializes the Email Assistant feature by ensuring all necessary sheets,
 * including the 'E-post-Logg', are created. This function calls the main
 * workbook setup routine.
 */
function initializeEmailFeature() {
  try {
    if (typeof setupWorkbook !== 'function') {
      throw new Error('Kritisk funksjon `setupWorkbook` ble ikke funnet. Sjekk at `01_Setup_og_Vedlikehold.js` er lastet.');
    }
    // This will create 'E-post-Logg' from the definition in 00_App_Core.js
    setupWorkbook();
    showToast('E-postassistenten ble initialisert.', 'Suksess');
  } catch (e) {
    showAlert(`En feil oppstod under initialisering av E-postassistenten: ${e.message}`, 'Feil');
    safeLog('Email_Assistant_Init_Fail', e.message);
  }
}

const EMAIL_CONFIG = {
  processingLabel: 'Sameie/Til Behandling',
  processedLabel: 'Sameie/Behandlet',
  processingBatchSize: 10 // Max emails to process in one run
};

/**
 * Main function to fetch and process incoming emails.
 * Searches for emails with a specific label, processes them, and logs them to the sheet.
 */
function processIncomingEmails() {
  try {
    const { processingLabel, processedLabel, processingBatchSize } = EMAIL_CONFIG;

    // Ensure labels exist
    _ensureGmailLabelExists(processingLabel);
    _ensureGmailLabelExists(processedLabel);

    const threads = GmailApp.search(`label:${processingLabel} -label:${processedLabel}`, 0, processingBatchSize);

    if (threads.length === 0) {
      Logger.log('Ingen nye e-poster å behandle.');
      return;
    }

    const logSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPOST_LOGG);
    if (!logSheet) {
      throw new Error(`Arket '${SHEETS.EPOST_LOGG}' ble ikke funnet.`);
    }

    const processingTimestamp = new Date();

    threads.forEach(thread => {
      const message = thread.getMessages()[0]; // Process first message in thread
      const emailData = {
        id: message.getId(),
        receivedDate: message.getDate(),
        sender: message.getFrom(),
        subject: message.getSubject(),
        body: message.getPlainBody(),
        threadId: thread.getId()
      };

      // 1. Categorize Email
      const category = _categorizeEmail(emailData.subject, emailData.body);

      // 2. Generate Reply Suggestion
      const replySuggestion = _generateReplySuggestion(category, emailData);

      // 3. Log to Sheet
      logSheet.appendRow([
        emailData.id,
        emailData.receivedDate,
        emailData.sender,
        emailData.subject,
        category,
        'Ny', // Status
        replySuggestion,
        emailData.body.substring(0, 500), // Log a snippet of the body
        emailData.threadId
      ]);

      // 4. Mark as processed
      thread.addLabel(GmailApp.getUserLabelByName(processedLabel));
      thread.removeLabel(GmailApp.getUserLabelByName(processingLabel));
      GmailApp.markThreadAsRead(thread);
    });

    safeLog('Email_Processing', `Behandlet ${threads.length} e-posttråder.`);
    showToast(`${threads.length} e-poster ble behandlet og loggført.`);

  } catch (e) {
    safeLog('Email_Processing_Fail', e.message);
    showAlert(`En feil oppstod under behandling av e-poster: ${e.message}`);
  }
}

/**
 * Ensures a specific Gmail label exists.
 * @param {string} labelName The name of the label to create.
 */
function _ensureGmailLabelExists(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * Rule-based engine to categorize an email based on its content.
 * @param {string} subject The email subject.
 * @param {string} body The email plain text body.
 * @returns {string} The determined category.
 */
function _categorizeEmail(subject, body) {
  const content = `${subject.toLowerCase()} ${body.toLowerCase()}`;

  // Rule mapping: keyword -> category
  const rules = {
    'faktura': 'Økonomi - Faktura',
    'invoice': 'Økonomi - Faktura',
    'parkering': 'Parkering',
    'lading': 'Elbil-lading',
    'elbil': 'Elbil-lading',
    'strøm': 'Strøm & Energi',
    'klage': 'Beboer-henvendelse',
    'nabovarsel': 'Byggesak/Nabovarsel',
    'vedlikehold': 'Vedlikehold',
    'heis': 'Vedlikehold - Heis',
    'dugnad': 'Dugnad',
    'årsmøte': 'Generalforsamling',
    'generalforsamling': 'Generalforsamling',
    'salg': 'Eierskifte',
    'eierskifte': 'Eierskifte'
  };

  for (const keyword in rules) {
    if (content.includes(keyword)) {
      return rules[keyword];
    }
  }

  return 'Generelt'; // Default category
}

/**
 * Generates a standard reply suggestion based on the email category.
 * @param {string} category The email category.
 * @param {object} emailData The extracted email data.
 * @returns {string} A suggested reply text.
 */
function _generateReplySuggestion(category, emailData) {
  const senderName = emailData.sender.split('<')[0].trim() || 'Beboer';

  const templates = {
    'Økonomi - Faktura': `Hei,\n\nTakk for din henvendelse.\n\nFakturaen er mottatt og vil bli behandlet.\n\nMvh,\nStyret`,
    'Parkering': `Hei ${senderName},\n\nTakk for din henvendelse angående parkering. Vi ser på saken og kommer tilbake til deg.\n\nMvh,\nStyret`,
    'Beboer-henvendelse': `Hei ${senderName},\n\nTakk for din henvendelse. Vi har mottatt din melding og vil ta tak i saken så snart som mulig.\n\nMvh,\nStyret`,
    'Generelt': `Hei ${senderName},\n\nTakk for din henvendelse. Vi har mottatt meldingen din og kommer tilbake til deg ved anledning.\n\nMvh,\nStyret`
  };

  return templates[category] || templates['Generelt'];
}

/**
 * Creates a time-based trigger to run the email processing function periodically.
 * It will run every hour. The function avoids creating duplicate triggers.
 */
function createEmailProcessingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExists = triggers.some(t => t.getHandlerFunction() === 'processIncomingEmails');

  if (!triggerExists) {
    ScriptApp.newTrigger('processIncomingEmails')
      .timeBased()
      .everyHours(1)
      .create();
    safeLog('Trigger_Creation', 'Opprettet timebasert trigger for e-postbehandling.');
    showToast('Automatisk e-postbehandling er aktivert (kjører hver time).');
  } else {
    showToast('Automatisk e-postbehandling er allerede aktivert.');
  }
}

/**
 * Tests the accuracy of the email categorization logic.
 * This function runs a series of predefined test cases against the
 * _categorizeEmail function and logs the accuracy score.
 */
function testEmailCategorizationAccuracy() {
  const testCases = [
    { subject: 'Faktura for strøm', body: 'Vedlagt er faktura for strømforbruk i fellesareal.', expected: 'Økonomi - Faktura' },
    { subject: 'Invoice #12345', body: 'Please find attached the invoice for services.', expected: 'Økonomi - Faktura' },
    { subject: 'Parkering - Gjest', body: 'Hvordan kan gjester parkere?', expected: 'Parkering' },
    { subject: 'Problem med lading av elbil', body: 'Laderen i garasjen virker ikke.', expected: 'Elbil-lading' },
    { subject: 'Spørsmål om strømavtale', body: 'Hvem er vår leverandør av strøm?', expected: 'Strøm & Energi' },
    { subject: 'Støy fra nabo', body: 'Jeg vil gjerne sende en formell klage på støy.', expected: 'Beboer-henvendelse' },
    { subject: 'Nabovarsel - Riving av garasje', body: 'Vi planlegger arbeid på vår tomt.', expected: 'Byggesak/Nabovarsel' },
    { subject: 'Defekt lyspære i oppgang A', body: 'Kan vaktmester bytte en lyspære?', expected: 'Vedlikehold' },
    { subject: 'Heisen står fast', body: 'Heisen har stoppet mellom etasjene.', expected: 'Vedlikehold - Heis' },
    { subject: 'Påmelding til dugnad', body: 'Jeg vil gjerne være med på lørdag.', expected: 'Dugnad' },
    { subject: 'Innkalling til årsmøte', body: 'Her er dokumentene til generalforsamlingen.', expected: 'Generalforsamling' },
    { subject: 'Salg av min leilighet', body: 'Jeg skal selge seksjon 101, trenger info.', expected: 'Eierskifte' },
    { subject: 'Nøkkelbrikke', body: 'Jeg har mistet nøkkelbrikken min.', expected: 'Generelt' },
    { subject: 'Sykkelparkering', body: 'Er det planer om flere sykkelstativ?', expected: 'Generelt' },
    { subject: 'Protokoll fra styremøte?', body: 'Hei, hvor finner jeg siste protokoll?', expected: 'Generelt' },
  ];

  let correctCount = 0;
  const results = [];

  testCases.forEach(test => {
    const actual = _categorizeEmail(test.subject, test.body);
    const isCorrect = actual === test.expected;
    if (isCorrect) {
      correctCount++;
    }
    results.push({
      subject: test.subject,
      expected: test.expected,
      actual: actual,
      correct: isCorrect
    });
  });

  const accuracy = (correctCount / testCases.length) * 100;
  const message = `Kategoriseringstest fullført.\n\nNøyaktighet: ${accuracy.toFixed(2)}% (${correctCount}/${testCases.length} korrekte).\n\nKrav: 80%.\n\nResultat: ${accuracy >= 80 ? 'GODKJENT' : 'FEILET'}`;

  Logger.log(message);
  Logger.log(JSON.stringify(results, null, 2));

  showAlert(message, 'Testresultat');

  return {
    accuracy: accuracy,
    results: results
  };
}