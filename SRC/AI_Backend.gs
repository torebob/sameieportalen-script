/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file authorization now that the Drive API is being used.
 *
 * To learn more about restricting authorization scopes, see:
 * https://developers.google.com/apps-script/guides/services/authorization
 */

/**
 * Gets an answer from the AI assistant based on internal documents.
 * This function is called from the frontend chat interface.
 *
 * @param {string} question The user's question from the chat.
 * @return {string} The answer to be displayed to the user.
 */
function getAiFaqAnswer(question) {
  // Access Control: Ensure the user is logged in.
  if (!Session.getActiveUser().getEmail()) {
    return "Tilgang nektet. Du må være logget inn for å bruke denne funksjonen.";
  }

  try {
    // Retrieve the Folder ID from script properties, which can be configured by an admin.
    const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('AI_ASSISTANT_FOLDER_ID');

    if (!FOLDER_ID) {
      return "AI-assistenten er ikke konfigurert. En administrator må angi en Google Drive-mappe-ID.";
    }

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    let context = "";

    // Extract text from all processable documents in the folder.
    while (files.hasNext()) {
      let file = files.next();
      let mimeType = file.getMimeType();

      try {
        if (mimeType === MimeType.GOOGLE_DOCS) {
          context += DocumentApp.openById(file.getId()).getBody().getText() + "\n\n";
        } else if (mimeType === MimeType.PLAIN_TEXT || mimeType.includes("text")) {
          context += file.getBlob().getDataAsString() + "\n\n";
        }
        // Note: Add support for other file types like PDFs if needed, though it requires more advanced parsing.
      } catch (e) {
        console.warn("Could not process file: " + file.getName() + " (" + e.message + ")");
      }
    }

    if (context.trim() === "") {
      return "Beklager, jeg kunne ikke finne noen dokumenter å søke i. Sjekk at mappen er riktig konfigurert og inneholder tekst-dokumenter.";
    }

    // --- Basic Keyword Search Implementation ---
    // This is a simple proof-of-concept. For a true "AI" experience,
    // this should be replaced with a more advanced solution like semantic search or a call to a private LLM.
    const keywords = question.toLowerCase().split(/\s+/).filter(k => k.length > 2);
    const sentences = context.split(/[.!?]/); // Split context into sentences

    let bestSentence = "Beklager, jeg fant ikke et klart svar på spørsmålet ditt i dokumentene.";
    let maxMatchCount = 0;

    sentences.forEach(sentence => {
      let currentMatches = 0;
      const lowerSentence = sentence.toLowerCase();

      keywords.forEach(keyword => {
        if (lowerSentence.includes(keyword)) {
          currentMatches++;
        }
      });

      if (currentMatches > maxMatchCount) {
        maxMatchCount = currentMatches;
        bestSentence = sentence.trim();
      }
    });

    if (maxMatchCount > 0) {
      return "Jeg fant følgende informasjon som kan være relevant: \"" + bestSentence + ".\"";
    }

    return bestSentence;

  } catch (e) {
    console.error("Error in getAiFaqAnswer: " + e.toString());
    // Return a user-friendly error message
    return "Det oppstod en teknisk feil under søket. Vennligst sjekk loggene eller prøv igjen senere.";
  }
}