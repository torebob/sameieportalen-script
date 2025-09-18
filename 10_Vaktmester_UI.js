/* ====================== Vaktmester UI (Launcher) ======================
 * FILE: 10_Vaktmester_UI.gs | VERSION: 1.0.0
 * Åpner VaktmesterVisning.html + meny.
 * ===================================================================== */

function openVaktmesterUI() {
  var html = HtmlService.createHtmlOutputFromFile('VaktmesterVisning')
    .setTitle('Vaktmester')
    .setWidth(1100)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vaktmester');
}

/** Legg til meny (slås sammen med eksisterende om du har en onOpen fra før) */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('Sameieportalen');
    // legg gjerne til andre punkter her også (Møteoversikt osv.)
    menu.addItem('Vaktmester', 'openVaktmesterUI');
    if (typeof openMeetingsUI === 'function') {
      menu.addItem('Møteoversikt & Protokoller', 'openMeetingsUI');
    }
    menu.addToUi();
  } catch(_) {}
}
