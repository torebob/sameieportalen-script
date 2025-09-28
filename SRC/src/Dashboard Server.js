/**
 * Dashboard Server (v1.0.0)
 * Main backend for the modal dashboard UI.
 * Provides metrics and module navigation.
 */

// Define the mapping of module keys to their HTML files and properties.
const UI_FILES = {
  DASHBOARD_HTML:  { file: '37_Dashboard.html', title: 'Sameieportal — Dashbord', w: 1280, h: 840 },
  MOTEOVERSIKT:    { file: '30_Moteoversikt.html', title: 'Møteoversikt & Protokoller', w: 1100, h: 760 },
  MOTE_SAK_EDITOR: { file: '31_MoteSakerEditor.html', title: 'Møtesaker – Editor', w: 1100, h: 760 },
  QNA_AGENT:       { file: '40_QnA_Agent.html', title: 'Spørsmål & Svar Agent', w: 900, h: 700 },
  QNA_SETTINGS:    { file: '41_Settings.html', title: 'Q&A Innstillinger', w: 800, h: 500 },
  // ... add all your other UI files here
};


/**
 * Fetches key metrics for the dashboard.
 * @returns {Object} Data object for the UI.
 */
function dashMetrics() {
  const functionName = 'dashMetrics';
  try {
    const user = getCurrentUserInfo();
    
    // TODO: Replace with your actual logic to count items from sheets
    const counts = {
      upcomingMeetings: 2,
      openTasks: 15,
      myTasks: (user.email === 'tore.sveinson@gmail.com' ? 3 : 0), // Example logic
      pendingApprovals: 1
    };
    
    return { ok: true, user: user, counts: counts };
  } catch (e) {
    Logger.error(functionName, 'Failed to fetch dashboard metrics.', { errorMessage: e.message });
    return { ok: false, error: e.message, counts: {} };
  }
}


/**
 * Opens a UI module based on a key. Called from the dashboard.
 * @param {string} key The key identifying the module to open (e.g., 'MOTEOVERSIKT').
 */
function dashOpen(key) {
  const functionName = 'dashOpen';
  const user = getCurrentUserInfo();
  
  Logger.info(functionName, `User opening module: ${key}`, { user: user.email });

  try {
    const ADMIN_ONLY_MODULES = ['MOTE_SAK_EDITOR', 'EIERSKIFTE', 'QNA_SETTINGS'];
    if (ADMIN_ONLY_MODULES.includes(key) && !user.isAdmin) {
      throw new Error(`Du mangler ADMIN-rettigheter for å åpne "${key}".`);
    }

    const spec = UI_FILES[key];
    if (!spec) {
      throw new Error(`Ukjent UI-nøkkel: ${key}`);
    }

    const htmlOutput = HtmlService.createTemplateFromFile(spec.file)
      .evaluate()
      .setTitle(spec.title)
      .setWidth(spec.w)
      .setHeight(spec.h);
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, spec.title);
    
  } catch (e) {
    Logger.error(functionName, `Failed to open module: ${key}`, { errorMessage: e.message });
    throw e; // Re-throw to be caught by the frontend's withFailureHandler
  }
}