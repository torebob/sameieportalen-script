/* ====================== Web App Central Router ======================
 * FILE: 00a_WebApp_Router.js | VERSION: 1.0.0 | UPDATED: 2025-09-26
 * FORMÅL: En enkelt, sentral doGet(e) for hele web-appen.
 * Denne funksjonen ruter forespørsler til riktig handler basert på
 * URL-parameteren 'page'. Dette løser konflikten med flere doGet-funksjoner.
 * ================================================================== */

function doGet(e) {
  try {
    const page = e && e.parameter && e.parameter.page;

    if (!page) {
      // Standard handling hvis ingen side er spesifisert.
      // For øyeblikket, la oss anta at dette skal åpne budsjett-appen som standard.
      if (typeof handleBudgetAppRequest === 'function') {
        return handleBudgetAppRequest(e);
      }
      return HtmlService.createHtmlOutput('<h1>Velkommen</h1><p>Ingen side spesifisert.</p>');
    }

    switch (page) {
      case 'protokoll':
        if (typeof handleProtokollApprovalRequest === 'function') {
          return handleProtokollApprovalRequest(e);
        }
        break;

      case 'tracking':
        if (typeof handleTrackingPixelRequest === 'function') {
          return handleTrackingPixelRequest(e);
        }
        break;

      case 'budget':
        if (typeof handleBudgetAppRequest === 'function') {
          return handleBudgetAppRequest(e);
        }
        break;

      default:
        return HtmlService.createHtmlOutput(`<h1>Ukjent side</h1><p>Siden '${page}' finnes ikke.</p>`);
    }

    // Fallback hvis funksjonen ikke ble funnet
    throw new Error(`Handler for page '${page}' is not defined.`);

  } catch (err) {
    Logger.log(`doGet Router Error: ${err.message}`);
    return HtmlService.createHtmlOutput(`<h1>En feil oppstod</h1><p>${err.message}</p>`);
  }
}