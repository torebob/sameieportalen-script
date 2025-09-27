/* ====================== Forms Scheduler (reminders & auto-close) ======================
 * FILE: 04_Forms_Scheduler.gs  |  VERSION: 2.0.0  |  UPDATED: 2025-09-26
 * FORMÅL: Daglig tidsstyrt håndtering av frister/påminnelser/auto-stenging for Google Forms.
 * ENDRINGER v2.0.0:
 *  - Modernisert til let/const og arrow functions.
 *  - Fjernet lokale hjelpefunksjoner; bruker nå 000_Utils.js.
 *  - Forbedret kodestruktur for lesbarhet.
 * ================================================================================ */

const FORMS_REG_SHEET = globalThis.SHEETS?.SKJEMA_REG || 'SkjemaRegister';

function setupFormsDailySchedulerTrigger() {
  _removeFormsDailySchedulerTrigger_();
  ScriptApp.newTrigger('runFormsDailyScheduler').timeBased().atHour(9).nearMinute(0).everyDays(1).create();
  _safeLog_('FormsScheduler', 'Daglig trigger satt kl. 09:00.');
}

function clearFormsDailySchedulerTrigger() {
  _removeFormsDailySchedulerTrigger_();
  _safeLog_('FormsScheduler', 'Daglig trigger fjernet.');
}

function runFormsDailyScheduler() {
  const today = _midnight_(new Date());
  const sh = _ensureFormsRegisterSheet_();
  if (!sh) {
    _safeLog_('FormsScheduler', 'Klarte ikke å opprette/finne registerark.');
    return;
  }

  const map = _headerMap_(sh);
  if (!map.skjemaNavn || !(map.formId || map.formURL) || !map.status) {
    _safeLog_('FormsScheduler', 'Mangler minimumskolonner (SkjemaNavn, FormId/URL, Status).');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    _safeLog_('FormsScheduler', 'Ingen rader i Skjema-register.');
    return;
  }

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  let processed = 0, reminders = 0, closed = 0, statsUpdated = 0;

  for (const [i, row] of values.entries()) {
    const rObj = _rowToObj_(row, map);
    if (!rObj.SkjemaNavn) continue;

    const status = String(rObj.Status || '').trim().toUpperCase();
    const isActive = ['ACTIVE', 'TRUE', 'ON', 'ÅPEN', ''].includes(status);

    if (!isActive || (rObj.StartDato && today < _midnight_(rObj.StartDato))) {
      continue;
    }

    processed++;
    const dueHasDate = !!rObj.FristDato;
    const daysLeft = dueHasDate ? _daysDiff_(today, _midnight_(rObj.FristDato)) : NaN;
    const isReminderDay = dueHasDate && rObj.PaaminnelseDager?.includes(daysLeft);
    const isAutoCloseDay = rObj.AutoSteng && dueHasDate && today > _midnight_(rObj.FristDato);

    if (!isReminderDay && !isAutoCloseDay && !isActive) continue;

    let form;
    try {
      const formId = rObj.FormId || _extractFormIdFromUrl_(rObj.FormURL);
      if (formId) form = FormApp.openById(formId);
    } catch (e) {
      _safeLog_('FormsScheduler', `Åpning av form "${rObj.SkjemaNavn}" feilet: ${e.message}`);
    }
    if (!form) continue;

    if (isActive) {
      const stats = _calcFormStats_(form);
      if (_writeBackStats_(sh, i + 2, map, stats, rObj)) statsUpdated++;
    }

    if (isReminderDay) {
      reminders += _sendReminderForForm_(sh, i + 2, rObj, daysLeft, map);
    }

    if (isAutoCloseDay) {
      try {
        if (form.isAcceptingResponses()) {
          form.setAcceptingResponses(false);
          if (map.status) _setCell_(sh, i + 2, map.status, 'CLOSED');
          _safeLog_('FormsScheduler', `Auto-stengte "${rObj.SkjemaNavn}" (frist utløpt).`);
          closed++;
        }
      } catch (e) {
        _safeLog_('FormsScheduler', `Auto-stenging feilet for "${rObj.SkjemaNavn}": ${e.message}`);
      }
    }
  }

  if (map.sistOppdatert) _setCell_(sh, 1, map.sistOppdatert, new Date());
  _safeLog_('FormsScheduler', `Kjøring ferdig. Aktive: ${processed}. Påminnelser: ${reminders}. Auto-stengt: ${closed}. Stats oppdatert: ${statsUpdated}.`);
}

function _sendReminderForForm_(sh, rowIndex, rObj, daysLeft, map) {
  const markerKey = `D-${daysLeft}`;
  if (_reminderAlreadySent_(rObj.RemindersSent, markerKey)) return 0;

  const seg = (String(rObj.Segment || '').trim().toUpperCase() || 'STYRET');
  const to = (seg === 'STYRET') ? _getBoardEmails_() : [];
  if (to.length === 0) {
    _safeLog_('FormsScheduler', `Påminnelse (ikke sendt – segment=${seg}, ingen mottakere) for "${rObj.SkjemaNavn}".`);
    return 0;
  }

  const subject = `[${APP.NAME}] Påminnelse: ${rObj.SkjemaNavn} (frist ${_fmtDate_(rObj.FristDato)})`;
  const body = `Hei,\n\nDette er en påminnelse om skjemaet "${rObj.SkjemaNavn}".\n${rObj.FormURL ? `Skjema: ${rObj.FormURL}\n` : ''}${rObj.FristDato ? `Frist: ${_fmtDate_(rObj.FristDato)}\n` : ''}\nMvh\n${APP.NAME}`;

  try {
    MailApp.sendEmail({ to: to.join(','), subject, body });
    _safeLog_('FormsScheduler', `Påminnelse D${daysLeft >= 0 ? '-' : ''}${daysLeft} sendt til STYRET (${to.length}) for "${rObj.SkjemaNavn}".`);
    _appendReminderMarker_(sh, rowIndex, map.remindersSent, markerKey);
    _setCell_(sh, rowIndex, map.sisteUts, new Date());
    return 1;
  } catch (e) {
    _safeLog_('FormsScheduler', `E-post feilet for "${rObj.SkjemaNavn}": ${e.message}`);
    return 0;
  }
}

function _ensureFormsRegisterSheet_() {
  try {
    const headers = ['SkjemaNavn', 'FormId', 'FormURL', 'Status', 'StartDato', 'FristDato', 'PaaminnelseDager', 'AutoSteng', 'Segment', 'SvarTeller', 'SistSvarTs', 'SisteUtsendelse', 'RemindersSent', 'SistOppdatert'];
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(FORMS_REG_SHEET);
    if (!sh) sh = ss.insertSheet(FORMS_REG_SHEET);

    const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (sh.getLastRow() === 0 || JSON.stringify(cur) !== JSON.stringify(headers)) {
      sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).clearContent();
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
    }
    return sh;
  } catch (e) {
    _safeLog_('FormsScheduler', `Klarte ikke sikre registerark: ${e.message}`);
    return null;
  }
}

function _removeFormsDailySchedulerTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFormsDailyScheduler') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function _headerMap_(sh) {
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const norm = s => String(s || '').toLowerCase().replace(/[\s_]+/g, '').trim();
  const want = {
    skjemaNavn: ['skjemanavn', 'navn'],
    formId: ['formid', 'id'],
    formURL: ['formurl', 'url'],
    status: ['status', 'aktiv'],
    startDato: ['startdato', 'start', 'aktivfra'],
    fristDato: ['fristdato', 'deadline', 'due'],
    paaminnelseDager: ['paaminnelsedager', 'påminnelsedager', 'reminderdays'],
    autoSteng: ['autosteng', 'autostengning', 'autoclose'],
    segment: ['segment', 'målgruppe', 'malgruppe'],
    svarTeller: ['svarteller', 'antallsvar', 'responses', 'count'],
    sistSvarTs: ['sistsvarts', 'lastresponse', 'lastrapporttid', 'sistsvar'],
    sisteUts: ['sisteutsendelse', 'lastsent', 'notified'],
    remindersSent: ['reminderssent', 'sendtepåminnelser', 'påminnelserlogg'],
    sistOppdatert: ['sistoppdatert', 'oppdatert', 'lastupdated']
  };
  return Object.keys(want).reduce((res, key) => {
    const foundIndex = hdr.findIndex(h => want[key].includes(norm(h)));
    if (foundIndex !== -1) res[key] = foundIndex + 1;
    return res;
  }, {});
}

function _rowToObj_(row, map) {
  const get = (idx) => (idx ? row[idx - 1] : '');
  const parseBool = (v) => /^(true|1|ja|yes)$/i.test(String(v).trim());
  const parseCSVints = (v) => String(v || '').split(',').map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n));
  const parseDate = (v) => (v instanceof Date && !isNaN(v)) ? v : (v ? new Date(v) : null);

  return {
    SkjemaNavn: get(map.skjemaNavn),
    FormId: get(map.formId),
    FormURL: get(map.formURL),
    Status: get(map.status),
    StartDato: parseDate(get(map.startDato)),
    FristDato: parseDate(get(map.fristDato)),
    PaaminnelseDager: parseCSVints(get(map.paaminnelseDager)),
    AutoSteng: parseBool(get(map.autoSteng)),
    Segment: get(map.segment),
    SvarTeller: get(map.svarTeller),
    SistSvarTs: parseDate(get(map.sistSvarTs)),
    RemindersSent: get(map.remindersSent)
  };
}

function _calcFormStats_(form) {
  if (!form) return { count: null, last: null };
  try {
    const res = form.getResponses();
    return { count: res.length, last: res.length ? res[res.length - 1].getTimestamp() : null };
  } catch (e) {
    _safeLog_('FormsScheduler', `getResponses() feilet: ${e.message}`);
    return { count: null, last: null };
  }
}

function _writeBackStats_(sh, rowIndex, map, newStats, oldRowObj) {
  let changed = false;
  if (map.svarTeller && newStats.count != null && String(newStats.count) !== String(oldRowObj.SvarTeller)) {
    _setCell_(sh, rowIndex, map.svarTeller, newStats.count);
    changed = true;
  }
  if (map.sistSvarTs && newStats.last && newStats.last.getTime() !== oldRowObj.SistSvarTs?.getTime()) {
    _setCell_(sh, rowIndex, map.sistSvarTs, newStats.last);
    changed = true;
  }
  if (changed && map.sistOppdatert) {
    _setCell_(sh, rowIndex, map.sistOppdatert, new Date());
  }
  return changed;
}

function _appendReminderMarker_(sh, rowIndex, colIndex, markerKey) {
  if (!colIndex) return;
  const cur = String(sh.getRange(rowIndex, colIndex).getValue() || '').trim();
  const stamp = _fmtDate_(new Date());
  const add = `${markerKey}|${stamp}`;
  const nextVal = cur ? `${cur};${add}` : add;
  sh.getRange(rowIndex, colIndex).setValue(nextVal);
}

function _reminderAlreadySent_(val, markerKey) {
  return String(val || '').split(';').some(p => p.trim().startsWith(`${markerKey}|`));
}

function _extractFormIdFromUrl_(url) {
  const m = String(url || '').match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

function _isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function _getBoardEmails_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.BOARD);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat().map(v => String(v || '').trim()).filter(_isValidEmail_);
}

function _setCell_(sh, row, col, v) {
  if (row && col) {
    try {
      sh.getRange(row, col).setValue(v);
    } catch (e) {
      _safeLog_('FormsScheduler', `Skriving til celle (${row},${col}) feilet: ${e.message}`);
    }
  }
}