/**
 * FIL: 00_WebApp_Router.gs
 * FORMÅL: Hovedinngang (doGet) + ruting for webappen.
 * - ?gid & token            -> handleProtocolClick(e)
 * - ?oppslagId (& personId) -> handleTrackingPixel(e)
 * - ellers                  -> index.html (frontend)
 *
 * MERK:
 *  - index.html må bruke Apps Script templates:
 *       <?!= include('public/css/app'); ?>
 *       <?!= include('public/js/app'); ?>
 *  - Derfor MÅ vi bruke createTemplateFromFile(...).evaluate() i doGet().
 */

/** Inkluder en HTML-delfil uten escaping.
 *  Bruk i HTML:  <?!= include('sti/til/fil'); ?>
 *  NB: 'sti/til/fil' peker til en HTML-fil i prosjektet, uten .html-suffiks.
 */
/** Inkluder en HTML-del uten escaping. Eneste sannhetskilde. */
function includeHtml(path) {
  try {
    return HtmlService.createHtmlOutputFromFile(path).getContent();
  } catch (e) {
    // Hjelper ved feilsøking i UI
    return '<!-- includeHtml ERROR: ' + path + ' :: ' + e.message + ' -->';
  }
}

/** Hovedruter for Web App */
function doGet(e) {
  e = e || { parameter: {} };

  // 1) Protokoll-godkjenning fra e-post
  if (e.parameter.gid && e.parameter.token) {
    if (typeof handleProtocolClick === 'function') {
      return handleProtocolClick(e);
    }
    return HtmlService.createHtmlOutput('<h3>Protokoll</h3><p>Handler mangler i prosjektet.</p>');
  }

  // 2) Sporingspiksel fra oppslag (e-poståpning)
  if (e.parameter.oppslagId) {
    if (typeof handleTrackingPixel === 'function') {
      return handleTrackingPixel(e);
    }
    // Fallback: returner en transparent 1x1 GIF så pikselen ikke feiler
    var gif = Utilities.base64Decode('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7');
    return ContentService.createImage(gif).setMimeType(ContentService.MimeType.GIF);
  }

  // 3) Standard: server frontend (index.html) som TEMPLATE
  return HtmlService
    .createTemplateFromFile('index')  // <- viktig for <?!= include(...) ?>
    .evaluate()
    .setTitle('Sameieportalen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------
   (Valgfritt) Hvis du ikke allerede har disse handlere i andre filer,
   kan du midlertidig bruke stubber under for å teste frontenden.
   Slett stubber når ekte implementasjoner finnes.
------------------------------------------------------------------- */

// STUB: handleProtocolClick
/*
function handleProtocolClick(e) {
  var p = e.parameter || {};
  var msg = 'Stub: Protokoll-lenke mottatt.'
          + '<br>gid=' + (p.gid || '?')
          + ' token=' + (p.token || '?')
          + ' action=' + (p.action || '?');
  return HtmlService.createHtmlOutput('<h3>Protokoll</h3><p>' + msg + '</p>');
}
*/

// STUB: handleTrackingPixel
/*
function handleTrackingPixel(e) {
  var gif = Utilities.base64Decode('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7');
  return ContentService.createImage(gif).setMimeType(ContentService.MimeType.GIF);
}
*/
