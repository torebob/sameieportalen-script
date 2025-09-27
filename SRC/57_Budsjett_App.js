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

// MERK: onOpen() er fjernet herfra. Menyen "Økonomi" er integrert i 00_App_Core.js.

function openBudgetWebapp() {
  let url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Budsjett-appen er ikke publisert som en webapp ennå.');
    return;
  }

  // Legg til page-parameter for routeren
  url += '?page=budget';

  // Viser en enkel HTML-side med en lenke som kan åpnes i ny fane.
  const html = `
    <p>Klikk på lenken for å åpne budsjett-appen:</p>
    <a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a>
    <script>document.querySelector('a').click(); setTimeout(google.script.host.close, 500);</script>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(100),
    'Åpner Budsjett Webapp...'
  );
}
