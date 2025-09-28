/* ====================== AI-assistent for Kommunikasjon ======================
 * FILE: 22_AI_Kommunikasjon_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Backend for AI-drevet analyse av innkommende e-post.
 * ========================================================================== */

/**
 * Henter de 20 siste uleste e-posttrådene fra styrets innboks.
 * @returns {Array<Object>} En liste med e-postobjekter.
 */
function getEmailsForProcessing() {
  try {
    // Autentisering trengs ikke her siden det er en intern funksjon,
    // men vi kan legge til en sjekk for admin-rolle om nødvendig.
    // requirePermission('USE_AI_ASSISTANT');
    const appConfig = getAppConfig();
    const query = `is:unread in:inbox label:${appConfig.AI_ASSISTANT.GMAIL_LABEL}`;
    const threads = GmailApp.search(query, 0, 20);
    const emails = [];

    threads.forEach(thread => {
      const firstMessage = thread.getMessages()[0];
      if (firstMessage) {
        emails.push({
          threadId: thread.getId(),
          subject: firstMessage.getSubject(),
          from: firstMessage.getFrom(),
          date: firstMessage.getDate().toISOString(),
        });
      }
    });

    return emails;

  } catch (e) {
    safeLog('AI_Assistent_Feil', `getEmailsForProcessing: ${e.message}`);
    throw new Error(`Kunne ikke hente e-poster: ${e.message}`);
  }
}

/**
 * Sender e-postinnhold til en AI-tjeneste for analyse.
 * @param {string} threadId ID-en til e-posttråden som skal analyseres.
 * @returns {Object} Et objekt med klassifisering, oppsummering og svarforslag.
 */
function getAiAssistance(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      throw new Error("Fant ikke e-posttråden.");
    }

    const message = thread.getMessages()[0]; // Analyserer kun første melding for nå
    const emailContent = message.getPlainBody();
    const emailSubject = message.getSubject();

    // Placeholder for AI-kall. Dette vil bli erstattet med en ekte implementering.
    const aiResponse = _callGenerativeAi_({
      prompt: `Analyser følgende e-post og gi en klassifisering, en kort oppsummering, og et forslag til svar.

      Emne: ${emailSubject}
      Innhold:
      ${emailContent}

      Formatér svaret som en JSON-objekt med nøklene "classification", "summary", og "replySuggestion".`
    });

    // Marker e-posten som lest etter analyse for å unngå at den dukker opp igjen.
    thread.markRead();

    return { ok: true, ...aiResponse };

  } catch (e) {
    safeLog('AI_Assistent_Feil', `getAiAssistance: ${e.message}`);
    throw new Error(`AI-analyse feilet: ${e.message}`);
  }
}

/**
 * Kaller en ekstern AI-tjeneste (f.eks. Google Gemini) for å analysere e-postinnhold.
 * @private
 * @param {Object} payload Data som inneholder prompten for AI-tjenesten.
 * @returns {Object} Et parset JSON-objekt med analysen fra AI-en.
 * @throws {Error} Hvis API-kall feiler eller API-nøkkel mangler.
 */
function _callGenerativeAi_(payload) {
  const appConfig = getAppConfig();
  const apiKey = appConfig.AI_ASSISTANT.API_KEY;

  if (!apiKey) {
    throw new Error("API-nøkkel for AI-tjenesten er ikke konfigurert i 'Konfig'-arket.");
  }

  // MERK: For å overholde GDPR/Schrems II, sørg for at AI-leverandøren
  // prosesserer data i EU/EØS. Gemini API kan konfigureres for dette.
  const apiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;

  const requestBody = {
    contents: [{
      parts: [{ text: payload.prompt }],
    }],
    generationConfig: {
      temperature: 0.3,
      topK: 1,
      topP: 1,
      maxOutputTokens: 2048,
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true, // For å fange opp feil og gi bedre feilmeldinger
  };

  const response = UrlFetchApp.fetch(apiEndpoint, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    safeLog('AI_API_ERROR', `Code: ${responseCode}, Body: ${responseBody}`);
    throw new Error(`AI-tjenesten svarte med en feil (HTTP ${responseCode}). Sjekk loggen for detaljer.`);
  }

  try {
    const json = JSON.parse(responseBody);
    const textContent = json.candidates[0].content.parts[0].text;
    // Renser og parser JSON-innholdet som AI-en returnerer
    const cleanedJsonString = textContent.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(cleanedJsonString);
  } catch (e) {
    safeLog('AI_PARSE_ERROR', `Failed to parse AI response: ${responseBody}`);
    throw new Error(`Kunne ikke tolke svaret fra AI-tjenesten. Rådata: ${responseBody}`);
  }
}