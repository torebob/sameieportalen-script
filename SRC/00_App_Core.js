/* ====================== App Core & Constants ======================
 * FILE: 00_App_Core.gs | VERSION: 1.9.1 | UPDATED: 2025-09-14
 * FORMÅL: App-oppstart, konstanter, meny, generiske UI-åpnere.
 *
 * NYTT v1.9.1:
 *  - Sikkert, generisk _openHtmlFromMap_() med param-støtte + robust feillogging
 *  - Full pakke med manglende "åpne"-funksjoner (defineres kun hvis de ikke finnes fra før)
 *  - validateUIFiles() + checkUIFilesExist() for å finne manglende HTML-filer under utvikling
 *  - Forbedret onOpen(): logger manglende callback-funksjoner i meny, kobler inn TESTING-submeny hvis tilgjengelig
 *  - openDashboardAuto(): permission-sjekk (VIEW_USER_DASHBOARD) og adaptive visning (modal for brukere / sidepanel for admin)
 * ================================================================== */

const APP = Object.freeze({
  NAME: 'Sameieportalen',
  VERSION: '1.9.1',
  BUILD: '2025-09-14'
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
  // Møtemodul
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
  // Dashbord (stor modal for brukere)
  DASHBOARD_HTML:            { file:'37_Dashboard.html',                title:'Sameieportal — Dashbord',   w:1280, h:840 },

  // Styremodul / møter
  MOTEOVERSIKT:              { file:'30_Moteoversikt.html',             title:'Møteoversikt & Protokoller', w:1100, h:760 },
  MOTE_SAK_EDITOR:           { file:'31_MoteSakEditor.html',            title:'Møtesaker – Editor',         w:1100, h:760 },

  // Skjema/visninger
  EIERSKIFTE:                { file:'32_Eierskifteskjema.html',         title:'Eierskifteskjema',           w:980,  h:760 },
  PROTOKOLL_GODKJENNING:     { file:'33_ProtokollGodkjenningSkjema.html', title:'Protokoll-godkjenning',   w:980,  h:760 },
  SEKSJON_HISTORIKK:         { file:'34_SeksjonHistorikk.html',         title:'Seksjonshistorikk',          w:1100, h:760 },
  VAKTMESTER:                { file:'35_VaktmesterVisning.html',        title:'Vaktmester',                 w:1100, h:800 }
});

/* ---------- UI helpers ---------- */
function _ui(){ try { return SpreadsheetApp.getUi(); } catch(_) { return null; } }
function _safeLog_(topic, msg, extra){
  try { if (typeof _logEvent === 'function') _logEvent(topic, msg, extra || {}); } catch(_) {}
}
function _alert_(msg, title){
  try {
    const ui = _ui();
    if (ui) ui.alert(title || APP.NAME, String(msg), ui.ButtonSet.OK);
    else Logger.log(`ALERT [${title||APP.NAME}]: ${msg}`);
  } catch(e) {
    Logger.log(`ALERT failed: ${e && e.message} | ${msg}`);
  }
}
function _toast_(msg){
  try { SpreadsheetApp.getActive().toast(String(msg)); }
  catch(e){ Logger.log('Toast failed: ' + (e && e.message) + ' | Message: ' + msg); }
}

/* ---------- Generisk, sikker UI-åpner ---------- */
function _openHtmlFromMap_(key, target = 'modal', params){
  try {
    const ui = _ui();
    if (!ui) { Logger.log('UI not available for key: ' + key); return; }

    const cfg = (globalThis.UI_FILES && globalThis.UI_FILES[key]) || null;
    if (!cfg) throw new Error('Ukjent UI key: ' + key);

    // Templatenavn uten .html
    const base = String(cfg.file).replace(/\.html?$/i,'');
    const tpl = HtmlService.createTemplateFromFile(base);

    // Injektér standard meta + custom params til HTML
    tpl.FILE    = cfg.file;
    tpl.VERSION = APP.VERSION;
    tpl.UPDATED = APP.BUILD;
    tpl.PARAMS  = params || {};

    const out = tpl.evaluate().setTitle(cfg.title || APP.NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    if (target === 'sidebar') {
      ui.showSidebar(out);
    } else {
      out.setWidth(cfg.w || 1000).setHeight(cfg.h || 720);
      ui.showModalDialog(out, cfg.title || APP.NAME);
    }
    _safeLog_('UI_OPEN', `Åpnet ${key}`, { file: cfg.file, target });
    return out;
  } catch(e) {
    Logger.log('Failed to open UI ('+key+'): ' + (e && e.message));
    _safeLog_('UI_OPEN_FAIL', 'Feil ved åpning av '+key, { error: String(e && e.message || e) });
    _alert_('Kunne ikke åpne ' + key + ': ' + (e && e.message ? e.message : e), 'Feil');
  }
}

/* ---------- Dashboard-åpnere ---------- */
function openDashboardModal(){ return _openHtmlFromMap_('DASHBOARD_HTML','modal'); }
function openDashboardSidebar(){ 
  // Adminpanel i sidepanel (krever 05_Dashboard_UI.gs med globalThis.openDashboard)
  if (typeof globalThis.openDashboard === 'function') return globalThis.openDashboard();
  // Fallback: vis dashbord-modal hvis sidepanel-koden ikke finnes
  return openDashboardModal();
}

/** Adaptiv dashbord-åpner (bruker modal for brukere / sidepanel for admin). */
function openDashboardAuto(){
  if (typeof hasPermission === 'function' && !hasPermission('VIEW_USER_DASHBOARD')) {
    _safeLog_('SECURITY', 'Dashbord nektet (mangler VIEW_USER_DASHBOARD).');
    return _alert_('Du har ikke tilgang til Dashbord.', 'Tilgang nektet');
  }
  const isAdmin = (typeof hasPermission === 'function' && hasPermission('VIEW_ADMIN_MENU'));
  return isAdmin ? openDashboardSidebar() : openDashboardModal();
}

/* ---------- Meny / oppstart ---------- */
function onOpen(e){
  const ui = _ui();
  if (!ui){ Logger.log('onOpen: UI ikke tilgjengelig (headless).'); return; }

  const G = globalThis;
  const menu = ui.createMenu(APP.NAME);

  const addIf = (label, fn) => {
    if (typeof G[fn] === 'function') menu.addItem(label, fn);
    else Logger.log('Mangler meny-funksjon: ' + fn + ' (skipper "' + label + '")');
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

  // Admin
  if (typeof hasPermission === 'function' && hasPermission('VIEW_ADMIN_MENU')) {
    menu.addSeparator();
    const admin = ui.createMenu('Admin');
    addIf('Opprett basisfaner', 'createBaseSheets');
    addIf('Kjør kvalitetssjekk', 'runAllChecks');
    admin.addSeparator();
    addIf('Synkroniser årshjul til kalender', 'syncYearWheelToCalendar');
    admin.addSeparator();
    // Eksplisitt adminpanel i sidepanel (05_Dashboard_UI.gs)
    addIf('Åpne Adminpanel (sidepanel)', 'openDashboardSidebar');
    admin.addSeparator();
    addIf('Skru PÅ Utvikler-verktøy', 'adminEnableDevTools');
    addIf('Skru AV Utvikler-verktøy', 'adminDisableDevTools');
    // Utviklerhjelp: sjekk at HTML-filer finnes
    admin.addSeparator();
    admin.addItem('Valider UI-filer', 'checkUIFilesExist');
    menu.addSubMenu(admin);
  }

  // TESTING-submeny hvis testmodul finnes (00_Testing_AddOn.gs)
  if (typeof buildTestingSubmenu_ === 'function') {
    buildTestingSubmenu_(menu);
  } else {
    // Enkel fallback for utvikling
    const test = ui.createMenu('TESTING');
    test.addItem('Valider UI-filer', 'checkUIFilesExist');
    test.addItem('Røyk-test menyfunksjoner', 'runSmokeCheck_');
    menu.addSubMenu(test);
  }

  menu.addToUi();
}
function onInstall(e){ onOpen(e); }

/* ---------- Admin/Dev togglers ---------- */
function adminEnableDevTools(){
  PROPS.setProperty(PROP_KEYS.DEV_TOOLS,'true');
  _toast_('Utvikler-verktøy er PÅ. Last regnearket på nytt for å oppdatere menyen.');
}
function adminDisableDevTools(){
  PROPS.setProperty(PROP_KEYS.DEV_TOOLS,'false');
  _toast_('Utvikler-verktøy er AV. Last regnearket på nytt for å oppdatere menyen.');
}

/* ---------- Manglende ÅPNERE (defineres kun hvis de ikke finnes) ---------- */
if (typeof globalThis.openMeetingsUI !== 'function') {
  globalThis.openMeetingsUI = function(){ _openHtmlFromMap_('MOTEOVERSIKT','modal'); };
}
if (typeof globalThis.openMoteSakEditor !== 'function') {
  globalThis.openMoteSakEditor = function(){ _openHtmlFromMap_('MOTE_SAK_EDITOR','modal'); };
}
if (typeof globalThis.openOwnershipForm !== 'function') {
  globalThis.openOwnershipForm = function(){ _openHtmlFromMap_('EIERSKIFTE','modal'); };
}
if (typeof globalThis.openSectionHistory !== 'function') {
  globalThis.openSectionHistory = function(){ _openHtmlFromMap_('SEKSJON_HISTORIKK','modal'); };
}
if (typeof globalThis.openVaktmesterUI !== 'function') {
  globalThis.openVaktmesterUI = function(){ _openHtmlFromMap_('VAKTMESTER','modal'); };
}

/* ---------- Utvikler-verktøy: valider at HTML-filer finnes ---------- */
function validateUIFiles(){
  const missing = [];
  const entries = Object.entries(UI_FILES || {});
  for (let i=0;i<entries.length;i++){
    const [key, cfg] = entries[i];
    try {
      const base = String(cfg.file).replace(/\.html?$/i,'');
      HtmlService.createTemplateFromFile(base); // kaster hvis ikke finnes
    } catch(e) {
      missing.push({ key, file: cfg && cfg.file, error: String(e && e.message || e) });
    }
  }
  return missing;
}
function checkUIFilesExist(){
  const miss = validateUIFiles();
  if (!miss.length) { _toast_('Alle UI-filer OK.'); return true; }
  Logger.log('Mangler UI-filer: ' + JSON.stringify(miss));
  const ui = _ui();
  const htmlRows = miss.map(m => `<tr><td>${m.key}</td><td>${m.file||'(ukjent)'}</td><td>${m.error}</td></tr>`).join('');
  const out = HtmlService.createHtmlOutput(`
    <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f6f6f6}</style>
    <h3>Mangler UI-filer</h3>
    <table><thead><tr><th>Key</th><th>Fil</th><th>Feil</th></tr></thead><tbody>${htmlRows}</tbody></table>
  `).setWidth(700).setHeight(420);
  ui && ui.showModalDialog(out, 'Validering av UI-filer');
  return false;
}

/* ---------- Enkel røyk-test for menyfunksjoner ---------- */
function runSmokeCheck_(){
  const fns = ['openDashboardAuto','openMeetingsUI','openMoteSakEditor','openOwnershipForm','openSectionHistory','openVaktmesterUI','checkUIFilesExist'];
  const res = fns.map(fn => ({ fn, ok: (typeof globalThis[fn] === 'function') }));
  const allOk = res.every(r=>r.ok);
  Logger.log('Smoke check: ' + JSON.stringify(res));
  _toast_((allOk ? 'OK' : 'Feil i') + ' røyk-test – se Logg.');
  return res;
}
