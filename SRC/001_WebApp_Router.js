/* ====================== Web App Central Router ======================
 * FILE: 00a_WebApp_Router.js | VERSION: 2.0.0 | UPDATED: 2025-09-26
 * FORMÅL: En enkelt, sentral doGet(e) for hele web-appen.
 * Denne funksjonen ruter forespørsler til riktig handler basert på
 * URL-parameteren 'page'. Dette løser konflikten med flere doGet-funksjoner.
 *
 * ENDRINGER v2.0.0:
 *  - Introdusert en handler-mapping for enklere vedlikehold.
 *  - Forbedret logging for ukjente sider og manglende handlere.
 *  - Returnerer en mer informativ feilmelding til brukeren.
 * ================================================================== */

function doGet(e) {
  try {
    // Bruker optional chaining for å trygt hente 'page'-parameteren.
    const page = e?.parameter?.page;

    // Definerer en mapping fra 'page'-parameter til en handler-funksjon.
    const pageHandlers = {
      protokoll: handleProtokollApprovalRequest,
      'protokoll-signering': renderProtokollSigneringPage,
      tracking: handleTrackingPixelRequest,
      budget: handleBudgetAppRequest,
      faq: handleFaqRequest, // Ny rute for FAQ-siden
      leverandor: renderLeverandorProfilPage, // Rute for leverandørprofil
      leverandorliste: renderLeverandorListePage, // Rute for leverandørliste
      minside: renderMinSidePage, // Rute for Min Side
    };

    // Hvis ingen side er spesifisert, bruk 'budget' som standard.
    const handlerKey = page || 'budget';
    const handler = pageHandlers[handlerKey];

    if (typeof handler === 'function') {
      return handler(e);
    }

    // Håndter ukjente sider eller manglende funksjoner
    const errorMessage = `Handler for siden '${handlerKey}' er ikke definert eller funnet.`;
    Logger.log(`doGet Router Warning: ${errorMessage}`);

    if (page) {
        return HtmlService.createHtmlOutput(`<h1>Ukjent side</h1><p>Siden '${escapeHtml(page)}' finnes ikke.</p>`);
    } else {
        // Fallback hvis standard-handleren (budget) mangler.
        return HtmlService.createHtmlOutput('<h1>Velkommen</h1><p>Standard-siden kunne ikke lastes.</p>');
    }

  } catch (err) {
    const errorMessage = err?.message || String(err);
    Logger.log(`doGet Router Error: ${errorMessage}`);
    return HtmlService.createHtmlOutput(`<h1>En feil oppstod</h1><p>${escapeHtml(errorMessage)}</p>`);
  }
}

function renderProtokollSigneringPage(e) {
  const moteId = e?.parameter?.moteId;
  if (!moteId) {
    return HtmlService.createHtmlOutput('<h1>Feil</h1><p>Møte-ID mangler.</p>').setTitle('Feil');
  }

  const template = HtmlService.createTemplateFromFile('35_ProtokollGodkjenningSkjema');
  template.moteId = moteId;

  return template.evaluate().setTitle('Signer Protokoll');
}

/**
 * Viser leverandørprofil-siden.
 * Krever 'id' (radnummer) som URL-parameter.
 */
function renderLeverandorProfilPage(e) {
  const vendorId = e?.parameter?.id;
  if (!vendorId) {
    return HtmlService.createHtmlOutput('<h1>Feil</h1><p>Leverandør-ID mangler.</p>').setTitle('Feil');
  }

  const template = HtmlService.createTemplateFromFile('44_Leverandor_Profil');
  template.vendorId = vendorId;

  return template.evaluate().setTitle('Leverandørprofil');
}

/**
 * Viser leverandørliste-siden.
 */
function renderLeverandorListePage(e) {
  return HtmlService.createTemplateFromFile('45_Leverandor_Liste')
    .evaluate()
    .setTitle('Leverandørliste');
}

/**
 * Viser "Min Side" for den innloggede brukeren.
 */
function renderMinSidePage(e) {
  const template = HtmlService.createTemplateFromFile('46_MinSide');
  template.appName = APP.NAME; // Gi app-navnet til malen
  return template.evaluate().setTitle('Min Side');
}