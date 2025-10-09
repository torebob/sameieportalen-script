/** ================= MENY-REPARATØR (Sameieportalen) — SAFE v2 ================ */
/** Kall denne én gang for å bygge alle menypunkter med en gang. */
function menus_repairQuick() {
  var ss = SpreadsheetApp.getActive();
  // 1) Bygg kjerne-/prosjektmenyer
  _menus_safeCallNoArgs_('uiBootstrap');          // hvis du har en sentral menykonstruktør
  _menus_safeCallNoArgs_('addDashboardMenu');     // Økonomi/Dashboard-meny
  _menus_safeCallNoArgs_('registerProjectMenu_'); // Prosjekt-meny (fra v97)
  _menus_safeCallNoArgs_('forceShowMenu');        // evt. prosjektspesifikk

  // 2) Bygg TESTING-meny robust (IKKE kall buildTestingSubmenu_ direkte)
  _menus_buildTestingMenuSafe_();

  // 3) Til slutt: gi beskjed
  ss.toast('Menyene er (re)bygget.');
}

/** Installerer idempotent onOpen-trigger som bygger menyene hver åpning. */
function menus_installOnOpen() {
  var ss = SpreadsheetApp.getActive();
  // Fjern tidligere like triggere
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'spOnOpen' &&
        t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('spOnOpen').forSpreadsheet(ss).onOpen().create();
  ss.toast('onOpen-trigger installert (spOnOpen). Last arket på nytt.');
}

/** Valgfritt: fjern onOpen-triggeren igjen. */
function menus_uninstallOnOpen() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'spOnOpen') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast('onOpen-trigger fjernet.');
}

/** Aggregator som kjøres ved åpning – robust og demper feil. */
function spOnOpen(e) {
  try { _menus_safeCallNoArgs_('uiBootstrap'); } catch(e){}
  try { _menus_safeCallNoArgs_('addDashboardMenu'); } catch(e){}
  try { _menus_safeCallNoArgs_('registerProjectMenu_'); } catch(e){}
  try { _menus_buildTestingMenuSafe_(); } catch(e){}
  try { _menus_safeCallNoArgs_('forceShowMenu'); } catch(e){}
}

/** Hjelper: kall funksjon uten argumenter hvis den finnes, demp feil. */
function _menus_safeCallNoArgs_(name) {
  try {
    var fn = globalThis[name];
    if (typeof fn === 'function') { fn(); }
  } catch (err) {
    Logger.log('Menybygger feilet: '+ name +' → ' + err);
  }
}

/** Bygg TESTING-meny uten å kalle buildTestingSubmenu_ (som krever parent-menu). */
function _menus_buildTestingMenuSafe_() {
  try {
    var ui = SpreadsheetApp.getUi();
    var m = ui.createMenu('TESTING');

    // Legg til det som faktisk finnes i prosjektet ditt – alt er valgfritt:
    if (typeof runSmokeCheck_ === 'function')      m.addItem('Smoke check', 'runSmokeCheck_');
    if (typeof runAllTestsQuick_ === 'function')   m.addItem('Kjør alle tester (rask)', 'runAllTestsQuick_');
    if (typeof runAllTestsAndShowReport_ === 'function')
                                                   m.addItem('Test-rapport', 'runAllTestsAndShowReport_');
    if (typeof getTestingTests_ === 'function')    m.addItem('Liste over tester', 'getTestingTests_');
    if (typeof testDocsHtml_ === 'function')       m.addItem('Vis test-dokumenter', 'testDocsHtml_');

    // Separator bare hvis vi har lagt til noe først
    try { m.addSeparator(); } catch(_) {}

    // Nyttige dev-verktøy hvis de finnes:
    if (typeof adminEnableDevTools === 'function')  m.addItem('Aktiver dev-verktøy', 'adminEnableDevTools');
    if (typeof adminDisableDevTools === 'function') m.addItem('Deaktiver dev-verktøy', 'adminDisableDevTools');

    m.addToUi();
  } catch (err) {
    Logger.log('TESTING-meny feilet: ' + err);
  }
}
