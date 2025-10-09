/**
 * FIL: 00_WebApp_Router.js
 * VERSJON: 1.0.6 (Endelig Fiks: Tilbake til enkle filnavn)
 * BESKRIVELSE: Hovedinngang (doGet) og ruting for webappen. Håndterer også klientkall.
 */
/**
 * FIL: 00_WebApp_Router.gs
 * FORMÅL: Hovedinngang (doGet) + ruting for webappen.
 * - ?gid & token            -> handleProtocolClick(e)
 * - ?oppslagId (& personId) -> handleTrackingPixel(e)
 * - ellers                  -> index.html (frontend)
 *
 * MERK:
 * - index.html MÅ nå bruke enkle filnavn (uten mapper):
 * <?!= include('public_css_app'); ?>
 * <?!= include('public_js_app'); ?>
 */

/** Inkluder en HTML-delfil uten escaping.
 * Bruk i HTML:  <?!= include('sti/til/fil'); ?>
 * NB: 'sti/til/fil' peker til en HTML-fil i prosjektet, uten .html-suffiks.
 */
function includeHtml(path) {
  try {
    // Denne gangen forventer vi at stien er det faktiske filnavnet
    // Siden vi bruker enkle filnavn i index.html, er dette den riktige måten å kalle det på.
    return HtmlService.createTemplateFromFile(path).evaluate().getContent();
  } catch (e) {
    // Logger feilen og returnerer HTML-kommentar for debugging i nettleseren
    Logger.log('includeHtml ERROR: Klarte ikke laste filen: %s. Feil: %s', path, e.message);
    return '<!-- includeHtml ERROR: Klarte ikke laste filen: ' + path + ' :: ' + e.message + ' -->';
  }
}

/**
 * Hovedruter for Web App.
 * Denne versjonen sikrer at HTML serveres som en mal,
 * slik at google.script.run blir tilgjengelig for klienten.
 */
function doGet(e) {
  // Standard: server frontend (index.html) som en TEMPLATE.
  // .evaluate() er den kritiske delen som bygger "broen" til serveren.
  return HtmlService
    .createTemplateFromFile('index') 
    .evaluate() 
    .setTitle('Sameieportalen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Inkluderer en HTML-delfil. Brukes slik i HTML: <?!= include('filnavn'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Håndterer alle asynkrone kall fra klientsiden. */
function doPost(e) {
  // Sjekker om e er en gyldig hendelse og inneholder data
  if (!e || !e.postData || e.postData.type !== 'application/json') {
    Logger.log('Invalid request: Missing postData or incorrect content type.');
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid request' })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const args = request.args;

    // Vi bruker 'global' for å referere til toppnivåfunksjoner i Apps Script
    if (typeof global[action] === 'function') {
      const result = global[action].apply(null, args);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: result })).setMimeType(ContentService.MimeType.JSON);
    } else {
      Logger.log('doPost Error: Action function not found: %s', action);
      return ContentService.createTextOutput(JSON.stringify({ error: 'Funksjonen ' + action + ' ble ikke funnet på serveren.' })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log('doPost Execution Error: %s', error.message);
    return ContentService.createTextOutput(JSON.stringify({ error: 'Serverfeil: ' + error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

/** Kalles av frontend for å hente innholdet i en HTML-del. */
function getHtmlContent(filename) {
  // Filnavnet må være uten sti for Apps Script
  const htmlFile = filename.split('/').pop();
  
  try {
    return HtmlService.createTemplateFromFile(htmlFile).evaluate().getContent();
  } catch (e) {
    // Hvis filen mangler, vil feilen her fanges.
    Logger.log('getHtmlContent Error: Failed to load: %s. Error: %s', htmlFile, e.message);
    return '<h2>Feil ved lasting av side</h2><p class="text-red-600">Klarte ikke å laste filen: <b>' + htmlFile + '</b>. Sjekk at filen eksisterer med riktig navn og path i prosjektet ditt.</p>';
  }
}
