/* ====================== Tasks Setup Patch ====================== */
/* FILE: 15_Tasks_Setup_Patch.gs | VERSION: 1.0 | UPDATED: 2025-09-14 */

const _TASKS_SHEET_NAME_ = (typeof SHEETS !== 'undefined' && SHEETS.TASKS) ? SHEETS.TASKS : 'Oppgaver';
const _TASKS_HEADERS_ = ['OppgaveID','Tittel','Ansvarlig','Status','Seksjonsnr','Frist'];

/** Opprett Oppgaver-arket hvis det mangler (med riktige kolonneoverskrifter). */
function ensureTasksSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(_TASKS_SHEET_NAME_);
  if (!sh) {
    sh = ss.insertSheet(_TASKS_SHEET_NAME_);
  }
  // Sørg for headers i rad 1
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const hasHeaders = lastRow >= 1 && lastCol >= _TASKS_HEADERS_.length;
  if (!hasHeaders) {
    sh.clear();
    sh.getRange(1, 1, 1, _TASKS_HEADERS_.length).setValues([_TASKS_HEADERS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  Logger.log('Oppgaver-ark OK: ' + sh.getName());
  return sh.getName();
}

/** (Valgfritt) Legg inn én testoppgave tildelt deg selv, så Vaktmester-UI viser noe. */
function seedMyTask_Example() {
  const name = ensureTasksSheet();
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const email = (Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '').toLowerCase();
  const id = 'TASK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Oslo', 'yyyyMMdd-HHmmss');
  const row = [id, 'Sjekke lekkasje i garasje', email, 'Ny', 'G-1', new Date()];
  sh.appendRow(row);
  Logger.log('La inn testoppgave ' + id + ' til ' + email);
}

/** Kjør alt i ett: opprett arket og legg inn en testoppgave. */
function setupTasksAndSeed() {
  ensureTasksSheet();
  seedMyTask_Example();
}
