/**
 * Sameieportal – Utviklerverktøy
 * FILE: 99_Dev_Tools.gs
 * VERSION: 1.2.0
 * UPDATED: 2025-09-28
 *
 * FORMÅL:
 * - Inneholder verktøy og funksjoner som kun er ment for utviklere/administratorer.
 * - Hjelper med å validere prosjektoppsettet og feilsøke vanlige feil.
 *
 * SIKKERHET:
 * - Kritiske funksjoner som endrer systemtilstand er nå beskyttet av _requireAdmin_().
 */

// ============================================================================
//  PUBLIC FACING FUNCTIONS - SECURED
// ============================================================================

/**
 * Aktiverer utviklermodus. Krever admin.
 */
function adminEnableDevTools() {
  _requireAdmin_();
  adminEnableDevTools_();
}

/**
 * Deaktiverer utviklermodus. Krever admin.
 */
function adminDisableDevTools() {
  _requireAdmin_();
  adminDisableDevTools_();
}

/**
 * Kjører validering av UI-filer og viser resultatet. Krever admin.
 */
function checkUIFilesExist() {
    _requireAdmin_();
    checkUIFilesExist_();
}

// ============================================================================
//  INTERNAL IMPLEMENTATION FUNCTIONS
// ============================================================================

function adminEnableDevTools_() {
  try {
    PropertiesService.getScriptProperties().setProperty('DEV_TOOLS_ENABLED', 'true');
    SpreadsheetApp.getActive().toast('Utvikler-verktøy er PÅ. Last regnearket på nytt for å oppdatere menyen.');
  } catch (e) {
    Logger.log('adminEnableDevTools: Kunne ikke aktivere utvikler-verktøy. Feil: ' + e.message);
    SpreadsheetApp.getUi().alert('Kunne ikke aktivere utvikler-verktøy.');
  }
}

function adminDisableDevTools_() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('DEV_TOOLS_ENABLED');
    SpreadsheetApp.getActive().toast('Utvikler-verktøy er AV. Last regnearket på nytt for å oppdatere menyen.');
  } catch (e) {
    Logger.log('adminDisableDevTools: Kunne ikke deaktivere utvikler-verktøy. Feil: ' + e.message);
    SpreadsheetApp.getUi().alert('Kunne ikke deaktivere utvikler-verktøy.');
  }
}

function validateUIFiles_() {
  const missing = [];
  const uiConfig = (typeof CONFIG_UI !== 'undefined') ? CONFIG_UI : (globalThis.UI_FILES || {});
  const entries = Object.entries(uiConfig);

  for (let i = 0; i < entries.length; i++) {
    const [key, cfg] = entries[i];
    if (!cfg || !cfg.file) continue;
    try {
      HtmlService.createTemplateFromFile(String(cfg.file).replace(/\.html?$/i, ''));
    } catch (e) {
      missing.push({ key, file: cfg.file, error: String(e.message || e) });
    }
  }
  return missing;
}

function checkUIFilesExist_() {
  const functionName = 'checkUIFilesExist';
  try {
    const missingFiles = validateUIFiles_();
    if (!missingFiles.length) {
      SpreadsheetApp.getActive().toast('Alle UI-filer er gyldige og funnet.');
      return true;
    }

    const htmlRows = missingFiles.map(m => `<tr><td>${m.key}</td><td>${m.file || '(ukjent)'}</td><td>${m.error}</td></tr>`).join('');
    const output = HtmlService.createHtmlOutput(`
      <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f6f6f6}</style>
      <h3>Mangler UI-filer</h3>
      <table><thead><tr><th>Nøkkel</th><th>Filnavn</th><th>Feilmelding</th></tr></thead><tbody>${htmlRows}</tbody></table>
    `).setWidth(700).setHeight(420);

    SpreadsheetApp.getUi().showModalDialog(output, 'Validering av UI-filer');
    return false;
  } catch (e) {
    Logger.log('En feil oppstod under validering av UI-filer: ' + e.message);
    SpreadsheetApp.getUi().alert('En feil oppstod under validering av UI-filer.');
    return false;
  }
}

function runSmokeCheck_() {
  const functionName = 'runSmokeCheck_';
  try {
    if (typeof MENU_CONFIG === 'undefined') {
      SpreadsheetApp.getActive().toast('FEIL: MENU_CONFIG er ikke definert. Kan ikke kjøre røyk-test.');
      return [];
    }

    const mainMenuFunctions = MENU_CONFIG.mainMenu.items.map(item => item.function);
    const adminMenuFunctions = MENU_CONFIG.adminMenu.items.map(item => item.function);
    const functionsToCheck = [...new Set([...mainMenuFunctions, ...adminMenuFunctions])];

    const results = functionsToCheck.map(fn => ({
      functionName: fn,
      isDefined: (typeof globalThis[fn] === 'function')
    }));

    const allOk = results.every(r => r.isDefined);
    Logger.log(`Røyk-test ${allOk ? 'OK' : 'FEILET'}`, results);
    SpreadsheetApp.getActive().toast(`Røyk-test ${allOk ? 'OK' : 'FEILET'} – se loggen for detaljer.`);
    return results;

  } catch (e) {
    Logger.log('En feil oppstod under røyk-test: ' + e.message);
    SpreadsheetApp.getActive().toast('En feil oppstod under røyk-test. Sjekk loggen.');
    return [];
  }
}