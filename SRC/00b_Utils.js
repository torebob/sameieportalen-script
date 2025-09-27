/* ====================== Central Utilities ======================
 * FILE: 00b_Utils.js | VERSION: 1.0.0 | UPDATED: 2025-09-26
 * FORMÅL: En sentral samling av gjenbrukbare hjelpefunksjoner.
 * Dette forhindrer duplikatkode og gjør vedlikehold enklere.
 * ================================================================== */

/**
 * Trygt henter UI-objektet.
 * @returns {Ui} Google Apps Script UI-objektet, eller null.
 */
function _ui(){
  try {
    return SpreadsheetApp.getUi();
  } catch(_) {
    // Kjører i en context uten UI, f.eks. en trigger
    return null;
  }
}

/**
 * Logger en hendelse trygt, forutsatt at en _logEvent-funksjon eksisterer.
 * @param {string} topic - Emnet for loggen.
 * @param {string} msg - Loggmeldingen.
 * @param {object} [extra] - Valgfrie ekstra data.
 */
function _safeLog_(topic, msg, extra){
  try {
    if (typeof _logEvent === 'function') {
      _logEvent(topic, msg, extra || {});
    }
  } catch(_) {
    // Ignorer feil hvis logging feiler
  }
}

/**
 * Viser en UI-alert, med fallback til Logger.log hvis UI ikke er tilgjengelig.
 * @param {string} msg - Meldingen som skal vises.
 * @param {string} [title] - Tittelen på alert-boksen.
 */
function _alert_(msg, title){
  try {
    const ui = _ui();
    const appName = (typeof APP !== 'undefined' && APP.NAME) ? APP.NAME : 'Sameieportalen';
    if (ui) {
      ui.alert(title || appName, String(msg), ui.ButtonSet.OK);
    } else {
      Logger.log(`ALERT [${title || appName}]: ${msg}`);
    }
  } catch(e) {
    Logger.log(`ALERT failed: ${e && e.message} | ${msg}`);
  }
}

/**
 * Viser en toast-melding nederst i regnearket.
 * @param {string} msg - Meldingen som skal vises.
 */
function _toast_(msg){
  try {
    SpreadsheetApp.getActive().toast(String(msg));
  } catch(e){
    Logger.log('Toast failed: ' + (e && e.message) + ' | Message: ' + msg);
  }
}

/**
 * Henter den aktive brukerens e-post.
 * @returns {string} Brukerens e-post, eller en tom streng.
 */
function _currentEmail_(){
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch(e) {
    return '';
  }
}

/**
 * Henter tidssonen for skriptet.
 * @returns {string} Tidssonen, f.eks. 'Europe/Oslo'.
 */
function _tz_() {
  try {
    return Session.getScriptTimeZone() || 'Europe/Oslo';
  } catch(e) {
    return 'Europe/Oslo';
  }
}