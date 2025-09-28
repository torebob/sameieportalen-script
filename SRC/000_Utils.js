/* ====================== Central Utilities ======================
 * FILE: 00b_Utils.js | VERSION: 1.1.0 | UPDATED: 2025-09-28
 * FORMÅL: En sentral samling av gjenbrukbare hjelpefunksjoner.
 * Dette forhindrer duplikatkode og gjør vedlikehold enklere.
 * ENDRINGER v1.1.0: Standardisert navn (fjernet ledende underscore).
 * ================================================================== */

/**
 * Trygt henter UI-objektet.
 * @returns {Ui} Google Apps Script UI-objektet, eller null.
 */
function getUi(){
  try {
    return SpreadsheetApp.getUi();
  } catch(_) {
    // Kjører i en context uten UI, f.eks. en trigger
    return null;
  }
}

/**
 * Logger en hendelse trygt, forutsatt at en logEvent-funksjon eksisterer.
 * @param {string} topic - Emnet for loggen.
 * @param {string} msg - Loggmeldingen.
 * @param {object} [extra] - Valgfrie ekstra data.
 */
function safeLog(topic, msg, extra){
  try {
    if (typeof logEvent === 'function') {
      logEvent(topic, msg, extra || {});
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
function showAlert(msg, title){
  try {
    const ui = getUi();
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
function showToast(msg){
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
function getCurrentEmail(){
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
function getScriptTimezone() {
  try {
    return Session.getScriptTimeZone() || 'Europe/Oslo';
  } catch(e) {
    return 'Europe/Oslo';
  }
}

/**
 * Robust dato-parser som håndterer yyyy-MM-dd, dd.MM.yyyy, og Date-objekter.
 * @param {*} value - Verdien som skal parses.
 * @returns {Date|null} Et Date-objekt, eller null hvis ugyldig.
 */
function normalizeDate(value) {
  if (value instanceof Date && !isNaN(value)) return value;

  const s = String(value || '').trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  return null;
}

/**
 * Returnerer et Date-objekt satt til midnatt.
 * @param {Date} d - Datoen som skal nullstilles.
 * @returns {Date|null}
 */
function getMidnight(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  const newDate = new Date(d);
  newDate.setHours(0, 0, 0, 0);
  return newDate;
}

/**
 * Beregner antall dager mellom to datoer.
 * @param {Date} from - Startdato.
 * @param {Date} to - Sluttdato.
 * @returns {number} Antall dager.
 */
function getDaysDiff(from, to) {
  if (!from || !to) return NaN;
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  return Math.round((to.getTime() - from.getTime()) / MS_PER_DAY);
}

/**
 * Formaterer en dato til dd.MM.yyyy-format.
 * @param {Date} d - Datoen som skal formateres.
 * @param {string} [tz] - Tidssonen.
 * @returns {string} Den formaterte datostrengen.
 */
function formatDate(d, tz) {
  if (!d) return '';
  try {
    return Utilities.formatDate(d, tz || getScriptTimezone(), 'dd.MM.yyyy');
  } catch (e) {
    return d.toISOString().slice(0, 10);
  }
}

/**
 * Creates a map of header names to their 1-based column index.
 * @param {string[]} headerRow - The array of header strings.
 * @returns {Object.<string, number>} A map of lowercase header names to column indices.
 */
function createHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((header, i) => {
    const key = String(header || '').trim().toLowerCase();
    if (key) {
      map[key] = i + 1; // 1-based index
    }
  });
  return map;
}

/**
 * Safely stringifies a JavaScript object, handling circular references.
 * @param {*} obj - The object to stringify.
 * @returns {string} The JSON string.
 */
function stringifySafe(obj) {
  const seen = new WeakSet();
  try {
    return JSON.stringify(obj, (key, value) => {
      if (typeof value === 'object' && value !== null) {
        if (seen.has(value)) {
          return '[Circular]';
        }
        seen.add(value);
      }
      return value;
    });
  } catch (e) {
    return `<<JSON Error: ${e.message}>>`;
  }
}

/**
 * Safely sets a value in a specific cell.
 * @param {Sheet} sh - The sheet object.
 * @param {number} row - The 1-based row index.
 * @param {number} col - The 1-based column index.
 * @param {*} v - The value to set.
 */
function setCell(sh, row, col, v) {
  if (sh && row && col) {
    try {
      sh.getRange(row, col).setValue(v);
    } catch (e) {
      safeLog('setCell_Error', `Writing to cell (${row},${col}) failed: ${e.message}`);
    }
  }
}

/**
 * Gets a list of board member emails from the 'Styret' sheet.
 * @returns {string[]} An array of valid email addresses.
 */
function getBoardEmails() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEETS.BOARD);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 2, sh.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .map(v => String(v || '').trim())
      .filter(v => v.includes('@')); // Basic validation
  } catch (e) {
    safeLog('getBoardEmails_Error', e.message);
    return [];
  }
}