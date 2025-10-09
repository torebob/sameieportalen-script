/** ================= MENY-STABILIZER v2 (Sameieportalen) =======================
 * Stopper flimring ved å (1) fjerne gamle triggere, (2) throttle menybygging,
 * og (3) bygge Prosjekt/TESTING-meny sikkert uten skjøre avhengigheter.
 * --------------------------------------------------------------------------- */

var MENUS_LOCK_KEY = 'MENUS_LAST_BUILD_MS';
var MENUS_LOCK_WINDOW_MS = 5000; // minst 5 sek mellom forsøk

/** Engangskjøring: fjern triggere som bygger menyer og nullstill lås. */
function menus_stopFlickerQuick() {
  _menus_deleteSpreadsheetTriggers_([
    'spOnOpen',               // denne filen
    'registerProjectMenu_',   // v97
    'uiBootstrap', 'onOpen',  // evt. andre byggere
    'menus_repairQuick'
  ]);
  PropertiesService.getDocumentProperties().deleteProperty(MENUS_LOCK_KEY);
  SpreadsheetApp.getActive().toast('Menytriggere fjernet og lås nullstilt.');
}

/** Bygg alle menyer nå – men bare hvis låsen tillater det. */
function menus_repairQuick() {
  if (!_menus_shouldBuildNow_()) return;
  try { _menus_safeCall_('uiBootstrap'); } catch(e){}
  try { _menus_safeCall_('addDashboardMenu'); } catch(e){}
  try { _menus_safeCall_('registerProjectMenu_'); } catch(e){} // fra v97
  try { _menus_buildTestingMenuSafe_(); } catch(e){}
  try { _menus_safeCall_('forceShowMenu'); } catch(e){}
  SpreadsheetApp.getActive().toast('Menyene er (re)bygget.');
}

/** Installer idempotent onOpen-trigger som kaller spOnOpen (trygt). */
function menus_installOnOpen() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'spOnOpen' &&
        t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('spOnOpen').forSpreadsheet(ss).onOpen().create();
  SpreadsheetApp.getActive().toast('onOpen-trigger (spOnOpen) installert. Last arket på nytt.');
}

/** Aggregator som kjøres på åpning – beskyttet av “lock/throttle”. */
function spOnOpen(e) {
  if (!_menus_shouldBuildNow_()) return;
  try { _menus_safeCall_('uiBootstrap'); } catch(e){}
  try { _menus_safeCall_('addDashboardMenu'); } catch(e){}
  try { _menus_safeCall_('registerProjectMenu_'); } catch(e){}
  try { _menus_buildTestingMenuSafe_(); } catch(e){}
  try { _menus_safeCall_('forceShowMenu'); } catch(e){}
}

/* ----------------------------- HJELPEFUNKSJONER ----------------------------- */

function _menus_shouldBuildNow_() {
  var dp = PropertiesService.getDocumentProperties();
  var now = Date.now();
  var last = Number(dp.getProperty(MENUS_LOCK_KEY) || 0);
  if (now - last < MENUS_LOCK_WINDOW_MS) {
    Logger.log('Menu build throttled (lock).');
    return false;
  }
  dp.setProperty(MENUS_LOCK_KEY, String(now));
  return true;
}

function _menus_safeCall_(name) {
  try {
    var fn = globalThis[name];
    if (typeof fn === 'function') fn();
    else Logger.log('Menybygger finnes ikke: ' + name);
  } catch (err) {
    Logger.log('Menybygger feilet: ' + name + ' → ' + err);
  }
}

/** Slett alle SPREADSHEETS-triggere for oppgitte handler-navn (hvis de finnes). */
function _menus_deleteSpreadsheetTriggers_(handlers) {
  var set = Array.isArray(handlers) ? new Set(handlers) : null;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getTriggerSource() !== ScriptApp.TriggerSource.SPREADSHEETS) return;
    if (!set || set.has(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** Robust TESTING-meny – kaller IKKE buildTestingSubmenu_ direkte. */
function _menus_buildTestingMenuSafe_() {
  try {
    var ui = SpreadsheetApp.getUi();
    var m = ui.createMenu('TESTING');

    // Legg til bare det som finnes:
    if (typeof runSmokeCheck_ === 'function')            m.addItem('Smoke check', 'runSmokeCheck_');
    if (typeof runAllTestsQuick_ === 'function')         m.addItem('Kjør alle tester (rask)', 'runAllTestsQuick_');
    if (typeof runAllTestsAndShowReport_ === 'function') m.addItem('Test-rapport', 'runAllTestsAndShowReport_');
    if (typeof getTestingTests_ === 'function')          m.addItem('Liste over tester', 'getTestingTests_');
    if (typeof testDocsHtml_ === 'function')             m.addItem('Vis test-dokumenter', 'testDocsHtml_');

    try { m.addSeparator(); } catch(_) {}

    if (typeof adminEnableDevTools === 'function')       m.addItem('Aktiver dev-verktøy', 'adminEnableDevTools');
    if (typeof adminDisableDevTools === 'function')      m.addItem('Deaktiver dev-verktøy', 'adminDisableDevTools');

    m.addToUi();
  } catch (err) {
    Logger.log('TESTING-meny feilet: ' + err);
  }
}

/** (Valgfritt) Finn og logg alle triggere – nyttig ved feilsøking. */
function menus_listTriggers() {
  var rows = ScriptApp.getProjectTriggers().map(function(t){
    return [t.getHandlerFunction(), String(t.getTriggerSource()), String(t.getEventType())];
  });
  Logger.log(JSON.stringify(rows));
  return rows;
}
