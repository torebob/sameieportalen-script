/* ====================== App Core & Constants ======================
 * FILE: 00_App_Core.gs | VERSION: 2.0.0 | UPDATED: 2025-09-26
 * FORMÅL: App-oppstart, konstanter, meny, generiske UI-åpnere.
 *
 * ENDRINGER v2.0.0:
 *  - Modernisert til `let`/`const` og arrow functions.
 *  - Forbedret `validateUIFiles` med `for...of` og optional chaining.
 *  - Tydeliggjort kommentarer og logikk i UI-åpner og menybygger.
 * ================================================================== */

const APP = Object.freeze({
  NAME: 'Sameieportalen',
  VERSION: '2.0.0',
  BUILD: '2025-09-26'
});

const SHEETS = Object.freeze({
  SEKSJONER: 'Seksjoner',
  PERSONER: 'Personer',
  EIERSKAP: 'Eierskap',
  LEIE: 'Leieforhold',
  TASKS: 'Oppgaver',
  BOARD: 'Styret',
  LOGG: 'Hendelseslogg',
  KONFIG: 'Konfig',
  VEDLEGG: 'Vedlegg',
  REPORT: 'Rapport',
  HMS: 'HMS_Egenkontroll',
  MOTER: 'Møter',
  MOTE_SAKER: 'MøteSaker',
  MOTE_KOMMENTARER: 'MøteSakKommentarer'
});

const PROPS = PropertiesService.getScriptProperties();
const PROP_KEYS = Object.freeze({
  TASK_ID_SEQ: 'TASK_ID_SEQ',
  DEV_TOOLS: 'DEV_TOOLS'
});

/* ---------- Felles mapping til alle UI-filer (nummerert) ---------- */
globalThis.UI_FILES = Object.freeze({
  DASHBOARD_HTML:        { file: '37_Dashboard.html',                title: 'Sameieportal — Dashbord',   w: 1280, h: 840 },
  MOTEOVERSIKT:          { file: '30_Moteoversikt.html',             title: 'Møteoversikt & Protokoller', w: 1100, h: 760 },
  MOTE_SAK_EDITOR:       { file: '31_MoteSakEditor.html',            title: 'Møtesaker – Editor',         w: 1100, h: 760 },
  EIERSKIFTE:            { file: '34_EierskifteSkjema.html',         title: 'Eierskifteskjema',           w: 980,  h: 760 },
  PROTOKOLL_GODKJENNING: { file: '35_ProtokollGodkjenningSkjema.html', title: 'Protokoll-godkjenning',   w: 980,  h: 760 },
  SEKSJON_HISTORIKK:     { file: '32_SeksjonHistorikk.html',         title: 'Seksjonshistorikk',          w: 1100, h: 760 },
  VAKTMESTER:            { file: '33_VaktmesterVisning.html',        title: 'Vaktmester',                 w: 1100, h: 800 }
});

/*
 * MERK: UI-hjelpere (getUi, safeLog, showAlert, showToast) er flyttet til
 * den sentrale verktøyfilen 00b_Utils.js for å unngå duplisering.
 */

/* ---------- Generisk, sikker UI-åpner ---------- */
function _openHtmlFromMap_(key, target = 'modal', params = {}) {
  try {
    const ui = getUi();
    if (!ui) {
      Logger.log(`UI not available for key: ${key}`);
      return;
    }

    // Hent UI-konfigurasjon fra den globale mappingen
    const cfg = globalThis.UI_FILES?.[key];
    if (!cfg) {
      throw new Error(`Ukjent UI key: ${key}`);
    }

    // Opprett mal fra fil, og fjern .html-endingen
    const templateName = String(cfg.file).replace(/\.html?$/i, '');
    const template = HtmlService.createTemplateFromFile(templateName);

    // Injisér standard metadata og eventuelle egendefinerte parametere
    template.FILE = cfg.file;
    template.VERSION = APP.VERSION;
    template.UPDATED = APP.BUILD;
    template.PARAMS = params;

    // Evaluer malen og konfigurer vinduet
    const output = template.evaluate()
      .setTitle(cfg.title || APP.NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    // Vis som enten sidebar eller modal dialog
    if (target === 'sidebar') {
      ui.showSidebar(output);
    } else {
      output.setWidth(cfg.w || 1000).setHeight(cfg.h || 720);
      ui.showModalDialog(output, cfg.title || APP.NAME);
    }

    safeLog('UI_OPEN', `Åpnet ${key}`, { file: cfg.file, target });
    return output;
  } catch (e) {
    const errorMessage = e?.message || String(e);
    Logger.log(`Failed to open UI (${key}): ${errorMessage}`);
    safeLog('UI_OPEN_FAIL', `Feil ved åpning av ${key}`, { error: errorMessage });
    showAlert(`Kunne ikke åpne ${key}: ${errorMessage}`, 'Feil');
  }
}

/* ---------- Dashboard-åpnere ---------- */
function openDashboardModal() {
  return _openHtmlFromMap_('DASHBOARD_HTML', 'modal');
}

function openDashboardSidebar() {
  // Delegrer til den avanserte implementasjonen i 05_Dashboard_UI.js hvis den finnes
  if (typeof globalThis.openDashboard === 'function') {
    return globalThis.openDashboard();
  }
  // Fallback for å sikre at noe vises hvis 05-filen mangler
  return openDashboardModal();
}

/**
 * Adaptiv dashbord-åpner som viser modal for vanlige brukere
 * og et sidepanel for administratorer.
 */
function openDashboardAuto() {
  if (typeof hasPermission === 'function' && !hasPermission('VIEW_USER_DASHBOARD')) {
    return showAlert('Du har ikke tilgang til Dashbord.', 'Tilgang nektet');
  }
  const isAdmin = (typeof hasPermission === 'function' && hasPermission('VIEW_ADMIN_MENU'));
  return isAdmin ? openDashboardSidebar() : openDashboardModal();
}

/* ---------- Meny / oppstart ---------- */
function onOpen(e) {
  const ui = getUi();
  if (!ui) {
    Logger.log('onOpen: UI ikke tilgjengelig (headless).');
    return;
  }

  const menu = ui.createMenu(APP.NAME);

  // Hjelpefunksjon for å legge til menyelementer bare hvis funksjonen eksisterer globalt.
  const addIf = (label, fn) => {
    if (typeof globalThis[fn] === 'function') {
      menu.addItem(label, fn);
    } else {
      Logger.log(`Mangler meny-funksjon: ${fn} (skipper "${label}")`);
    }
  };

  // Hovedmeny (alle)
  addIf('Dashbord', 'openDashboardAuto');
  menu.addSeparator();
  addIf('Møteoversikt & Protokoller…', 'openMeetingsUI');
  addIf('Møtesaker (editor)…', 'openMoteSakEditor');
  addIf('Registrer eierskifte…', 'openOwnershipForm');
  addIf('Søk i seksjonshistorikk…', 'openSectionHistory');

  // Vaktmester (rollebasert)
  if (typeof hasPermission === 'function' && hasPermission('VIEW_VAKTMESTER_UI')) {
    menu.addSeparator();
    addIf('Mine Oppgaver (Vaktmester)', 'openVaktmesterUI');
  }

  // Rapporter (styret)
  if (typeof hasPermission === 'function' && hasPermission('GENERATE_REPORTS')) {
    menu.addSeparator();
    const reports = ui.createMenu('Rapporter');
    reports.addItem('Åpne saker per kategori', 'generateOpenCasesReport');
    menu.addSubMenu(reports);
  }

  // Økonomi (styret)
  if (typeof hasPermission === 'function' && hasPermission('VIEW_BUDGET_MENU')) {
    menu.addSeparator();
    const ekonomi = ui.createMenu('Økonomi');
    ekonomi.addItem('Åpne Budsjett (webapp)', 'openBudgetWebapp');
    menu.addSubMenu(ekonomi);
  }

  // Admin
  if (typeof hasPermission === 'function' && hasPermission('VIEW_ADMIN_MENU')) {
    menu.addSeparator();
    const admin = ui.createMenu('Admin');
    addIf('Opprett basisfaner', 'createBaseSheets');
    addIf('Kjør kvalitetssjekk', 'runAllChecks');
    admin.addSeparator();
    addIf('Synkroniser årshjul til kalender', 'syncYearWheelToCalendar');
    admin.addSeparator();
    addIf('Åpne Adminpanel (sidepanel)', 'openDashboardSidebar');
    admin.addSeparator();
    addIf('Skru PÅ Utvikler-verktøy', 'adminEnableDevTools');
    addIf('Skru AV Utvikler-verktøy', 'adminDisableDevTools');
    admin.addSeparator();
    admin.addItem('Valider UI-filer', 'checkUIFilesExist');

    // Analyse-verktøy
    if (typeof generateDiscoveryReportInDoc === 'function') {
      const analyse = ui.createMenu('Analyse');
      analyse.addItem('Generer Discovery-rapport (Doc)', 'generateDiscoveryReportInDoc');
      analyse.addItem('Åpne Discovery-dokument', 'openDiscoveryDocQuick');
      analyse.addItem('Foreslå nye krav → «Krav»-arket', 'discoverySuggestToKravQuick');

      if (typeof rg_menu_openWizard === 'function') {
        analyse.addSeparator();
        analyse.addItem('Krav Generator (wizard)', 'rg_menu_openWizard');
        analyse.addItem('Kjør krav-generator (uten UI)', 'rg_menu_runAllQuick');
        analyse.addItem('Åpne Requirements-arket', 'rg_menu_openReqSheet');
      }

      if (typeof rsp_menu_firstRunWizard === 'function') {
        analyse.addSeparator();
        analyse.addItem('Krav Sync (wizard)', 'rsp_menu_firstRunWizard');
        analyse.addItem('Valider Krav-system', 'rsp_menu_validateSystemState');
        analyse.addItem('Push Krav (Sheet → Doc)', 'rsp_menu_pushRun');
        analyse.addItem('Åpne Krav-dokument', 'rsp_menu_openDoc');
      }
      admin.addSeparator();
      admin.addSubMenu(analyse);
    }

    if (typeof runCoreAnalysis_Smoke === 'function') {
      const coreAnalyse = ui.createMenu('Core Analysis');
      coreAnalyse.addItem('Run Analysis (log)', 'runCoreAnalysis_Smoke');
      coreAnalyse.addItem('Open Dashboard (Mermaid)', 'ae_showDashboard');
      admin.addSeparator();
      admin.addSubMenu(coreAnalyse);
    }
    menu.addSubMenu(admin);
  }

  // TESTING-submeny
  if (typeof buildTestingSubmenu_ === 'function') {
    buildTestingSubmenu_(menu);
  } else {
    const test = ui.createMenu('TESTING');
    test.addItem('Valider UI-filer', 'checkUIFilesExist');
    test.addItem('Røyk-test menyfunksjoner', 'runSmokeCheck_');
    menu.addSubMenu(test);
  }

  menu.addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/* ---------- Admin/Dev togglers ---------- */
function adminEnableDevTools() {
  PROPS.setProperty(PROP_KEYS.DEV_TOOLS, 'true');
  showToast('Utvikler-verktøy er PÅ. Last regnearket på nytt for å oppdatere menyen.');
}

function adminDisableDevTools() {
  PROPS.setProperty(PROP_KEYS.DEV_TOOLS, 'false');
  showToast('Utvikler-verktøy er AV. Last regnearket på nytt for å oppdatere menyen.');
}

/* ---------- Manglende ÅPNERE (defineres kun hvis de ikke finnes) ---------- */
if (typeof globalThis.openMeetingsUI !== 'function') {
  globalThis.openMeetingsUI = () => _openHtmlFromMap_('MOTEOVERSIKT', 'modal');
}
if (typeof globalThis.openMoteSakEditor !== 'function') {
  globalThis.openMoteSakEditor = () => _openHtmlFromMap_('MOTE_SAK_EDITOR', 'modal');
}
if (typeof globalThis.openOwnershipForm !== 'function') {
  globalThis.openOwnershipForm = () => _openHtmlFromMap_('EIERSKIFTE', 'modal');
}
if (typeof globalThis.openSectionHistory !== 'function') {
  globalThis.openSectionHistory = () => _openHtmlFromMap_('SEKSJON_HISTORIKK', 'modal');
}
if (typeof globalThis.openVaktmesterUI !== 'function') {
  globalThis.openVaktmesterUI = () => _openHtmlFromMap_('VAKTMESTER', 'modal');
}

/* ---------- Utvikler-verktøy: valider at HTML-filer finnes ---------- */
function validateUIFiles() {
  const missing = [];
  const entries = Object.entries(globalThis.UI_FILES || {});
  for (const [key, cfg] of entries) {
    try {
      const base = String(cfg.file).replace(/\.html?$/i, '');
      HtmlService.createTemplateFromFile(base); // throws if not found
    } catch (e) {
      missing.push({ key, file: cfg?.file, error: String(e?.message || e) });
    }
  }
  return missing;
}

function checkUIFilesExist() {
  const missingFiles = validateUIFiles();
  if (missingFiles.length === 0) {
    showToast('Alle UI-filer er tilgjengelige.');
    return true;
  }

  Logger.log('Mangler UI-filer: ' + JSON.stringify(missingFiles));
  const ui = getUi();
  const htmlRows = missingFiles.map(m => `<tr><td>${m.key}</td><td>${m.file || '(ukjent)'}</td><td>${m.error}</td></tr>`).join('');
  const output = HtmlService.createHtmlOutput(`
    <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f6f6f6}</style>
    <h3>Mangler UI-filer</h3>
    <table><thead><tr><th>Key</th><th>Fil</th><th>Feil</th></tr></thead><tbody>${htmlRows}</tbody></table>
  `).setWidth(700).setHeight(420);

  ui?.showModalDialog(output, 'Validering av UI-filer');
  return false;
}

/* ---------- Enkel røyk-test for menyfunksjoner ---------- */
function runSmokeCheck_() {
  const fns = ['openDashboardAuto', 'openMeetingsUI', 'openMoteSakEditor', 'openOwnershipForm', 'openSectionHistory', 'openVaktmesterUI', 'checkUIFilesExist'];
  const res = fns.map(fn => ({ fn, ok: typeof globalThis[fn] === 'function' }));
  const allOk = res.every(r => r.ok);
  Logger.log('Smoke check: ' + JSON.stringify(res));
  showToast((allOk ? 'OK' : 'Feil i') + ' røyk-test – se Logg.');
  return res;
}
