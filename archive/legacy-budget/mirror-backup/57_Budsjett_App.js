// =============================================================================
// Budsjett – Dashboard App (launcher)
// FILE: 57_Budsjett_App.gs
// VERSION: 1.0.1
// UPDATED: 2025-09-15
// NOTE: Viser HTML-fila 51_Budsjett_App.html som dashboard/webapp
// =============================================================================

function handleBudgetAppRequest(e) {
  return HtmlService.createTemplateFromFile('51_Budsjett_App')
    .evaluate()
    .setTitle('Budsjett')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // tillat innbygging i dashbord
}

function openBudgetWebapp() {
  let url = ScriptApp.getService().getUrl();
  if (!url) {
    _ui()?.alert('Budsjett-appen er ikke publisert som en webapp ennå.');
    return;
  }
  url += '?page=budget';

  // Bruker en liten HTML-dialog til å trigge åpning av ny fane via JavaScript.
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}", "_blank"); google.script.host.close();</script>`
  ).setHeight(10).setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Åpner...');
}
