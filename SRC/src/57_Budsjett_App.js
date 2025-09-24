// =============================================================================
// Budsjett – Dashboard App (launcher)
// FILE: 57_Budsjett_App.gs
// VERSION: 1.0.1
// UPDATED: 2025-09-15
// NOTE: Viser HTML-fila 51_Budsjett_App.html som dashboard/webapp
// =============================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('51_Budsjett_App')
    .evaluate()
    .setTitle('Budsjett')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // tillat innbygging i dashbord
}

// Valgfri meny i Sheets (åpner webapp-URL)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Økonomi')
    .addItem('Åpne Budsjett (webapp)', 'openBudgetWebapp')
    .addToUi();
}

function openBudgetWebapp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Webapp-URL:\n' + url + '\n\nPubliser som webapp om du får 404.');
}
