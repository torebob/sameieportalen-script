/**
 * Sameieportal – Utviklerverktøy
 * FILE: 99_Dev_Tools.gs
 * VERSION: 1.1.0
 * UPDATED: 2025-09-23
 *
 * FORMÅL:
 * - Inneholder verktøy og funksjoner som kun er ment for utviklere/administratorer.
 * - Hjelper med å validere prosjektoppsettet og feilsøke vanlige feil.
 * - Bør ikke inneholde kjernefunksjonalitet som vanlige brukere er avhengige av.
 *
 * HOVEDFUNKSJONER:
 * - adminEnableDevTools(), adminDisableDevTools(): Skrur av/på en global "utviklermodus"-bryter via PropertiesService.
 * - checkUIFilesExist(): Validerer at alle HTML-filer definert i CONFIG_UI (fra 00_Config.gs) faktisk finnes i prosjektet.
 * - runSmokeCheck_(): En dynamisk "røyk-test" som automatisk sjekker at alle funksjoner lenket til i MENU_CONFIG (fra MenuBuilder.gs) er definert.
 *
 * BRUK:
 * - Funksjonene kalles typisk fra en "Admin"- eller "TESTING"-meny i regnearket. Se instruksjoner over.
 */

/**
 * Aktiverer utviklermodus ved å sette en egenskap i PropertiesService.
 */
function adminEnableDevTools() {
  try {
    PropertiesService.getScriptProperties().setProperty(PROP_KEYS.DEV_TOOLS_ENABLED, 'true');
    _toast_('Utvikler-verktøy er PÅ. Last regnearket på nytt for å oppdatere menyen.');
  } catch (e) {
    Logger.error('adminEnableDevTools', 'Failed to enable dev tools.', { errorMessage: e.message });
    _alert_('Kunne ikke aktivere utvikler-verktøy.');
  }
}

/**
 * Deaktiverer utviklermodus.
 */
function adminDisableDevTools() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(PROP_KEYS.DEV_TOOLS_ENABLED);
    _toast_('Utvikler-verktøy er AV. Last regnearket på nytt for å oppdatere menyen.');
  } catch (e) {
    Logger.error('adminDisableDevTools', 'Failed to disable dev tools.', { errorMessage: e.message });
    _alert_('Kunne ikke deaktivere utvikler-verktøy.');
  }
}

/**
 * Validerer at alle UI-filer definert i CONFIG_UI eksisterer.
 * @returns {Array<object>} En liste over manglende filer.
 */
function validateUIFiles() {
  const missing = [];
  // Bruker CONFIG_UI direkte, som definert i 00_Config.gs
  const uiConfig = (typeof CONFIG_UI !== 'undefined') ? CONFIG_UI : (globalThis.UI_FILES || {});
  const entries = Object.entries(uiConfig);

  for (let i = 0; i < entries.length; i++) {
    const [key, cfg] = entries[i];
    if (!cfg || !cfg.file) continue;

    try {
      const base = String(cfg.file).replace(/\.html?$/i, '');
      HtmlService.createTemplateFromFile(base); // Kaster feil hvis filen ikke finnes
    } catch (e) {
      missing.push({ key, file: cfg.file, error: String(e.message || e) });
    }
  }
  return missing;
}

/**
 * Kjører validering av UI-filer og viser resultatet i en dialogboks.
 * @returns {boolean} True hvis alle filer er OK, ellers false.
 */
function checkUIFilesExist() {
  const functionName = 'checkUIFilesExist';
  try {
    const missingFiles = validateUIFiles();

    if (!missingFiles.length) {
      _toast_('Alle UI-filer er gyldige og funnet.');
      return true;
    }

    Logger.error(functionName, 'One or more UI files are missing.', { missing: missingFiles });
    const ui = _ui();
    const htmlRows = missingFiles.map(m => `<tr><td>${m.key}</td><td>${m.file || '(ukjent)'}</td><td>${m.error}</td></tr>`).join('');
    const output = HtmlService.createHtmlOutput(`
      <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f6f6f6}</style>
      <h3>Mangler UI-filer</h3>
      <table><thead><tr><th>Nøkkel</th><th>Filnavn</th><th>Feilmelding</th></tr></thead><tbody>${htmlRows}</tbody></table>
    `).setWidth(700).setHeight(420);

    if (ui) ui.showModalDialog(output, 'Validering av UI-filer');
    return false;
  } catch (e) {
    Logger.error(functionName, 'An error occurred during UI file validation.', { errorMessage: e.message });
    _alert_('En feil oppstod under validering av UI-filer.');
    return false;
  }
}

/**
 * Kjører en "røyk-test" som dynamisk sjekker at alle funksjoner
 * definert i MENU_CONFIG faktisk eksisterer i det globale skopet.
 * @returns {Array<object>} En liste med resultater fra sjekken.
 */
function runSmokeCheck_() {
  const functionName = 'runSmokeCheck_';
  try {
    if (typeof MENU_CONFIG === 'undefined') {
      _toast_('FEIL: MENU_CONFIG er ikke definert. Kan ikke kjøre røyk-test.');
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
    Logger.info(functionName, 'Smoke test results.', { results });
    _toast_(`Røyk-test ${allOk ? 'OK' : 'FEILET'} – se loggen for detaljer.`);
    return results;

  } catch (e) {
    Logger.error(functionName, 'An error occurred during smoke test.', { errorMessage: e.message });
    _toast_('En feil oppstod under røyk-test. Sjekk loggen.');
    return [];
  }
}