/* global Sameie */
/* ====================== Forms Scheduler (reminders & auto-close) ======================
 * FILE: 04_Forms_Scheduler.gs  |  VERSION: 1.7.0  |  UPDATED: 2025-09-13
 * FORMÅL: Daglig tidsstyrt håndtering av frister/påminnelser/auto-stenging for Google Forms
 *         basert på rader i Skjema-register-fanen. Oppdaterer også svarstatistikk.
 * Endringer v1.7.0: Robust fallback for arknavn, automatisk opprettelse av registerark,
 *                   router med tilstedeværelses-sjekk, forbedret logging og validering.
 * Avhenger av: _logEvent(), _tz_(), SHEETS (valgfritt for SKJEMA_REG)
 * ================================================================================ */

/* -------------------- Lokale konstanter (failsafe) -------------------- */
const FORMS_REG_SHEET = (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.SKJEMA_REG)
  ? SHEETS.SKJEMA_REG
  : 'SkjemaRegister';

/* ====================== Offentlige "entrypoints" ====================== */

/** Opprett én daglig trigger kl. 09:00 som kjører runFormsDailyScheduler(). */
function setupFormsDailySchedulerTrigger(){
  _removeFormsDailySchedulerTrigger_();
  ScriptApp.newTrigger('runFormsDailyScheduler')
    .timeBased()
    .atHour(9).nearMinute(0) // lokal tidssone fra prosjektet
    .everyDays(1)
    .create();
  if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Daglig trigger satt kl. 09:00.');
}

/** Fjern ev. eksisterende triggere for runFormsDailyScheduler(). */
function clearFormsDailySchedulerTrigger(){
  _removeFormsDailySchedulerTrigger_();
  if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Daglig trigger fjernet.');
}

/** Kjør scheduler nå (manuelt fra editor/meny). */
function runFormsDailyScheduler(){
  const tz = (typeof _tz_ === 'function') ? _tz_() : (Session.getScriptTimeZone() || 'Europe/Oslo');
  const today = _midnight_(new Date(), tz);

  // Sørg for at register-arket finnes og har riktige kolonner
  const sh = _ensureFormsRegisterSheet_();
  if (!sh){ if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Klarte ikke å opprette/finne registerark.'); return; }

  const map = _headerMap_(sh);
  if (!map.skjemaNavn || !(map.formId || map.formURL) || !map.status){
    if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Mangler minimumskolonner (SkjemaNavn, FormId/URL, Status).');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2){ if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Ingen rader i Skjema-register.'); return; }

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  let processed = 0, reminders = 0, closed = 0, statsUpdated = 0;

  for (let i=0; i<values.length; i++){
    const row = values[i];
    const rObj = _rowToObj_(row, map);

    if (!rObj.SkjemaNavn) continue;

    // Status-normalisering
    const status = String(rObj.Status || '').trim().toUpperCase();
    const isActive = (status === 'ACTIVE' || status === 'TRUE' || status === 'ON' || status === 'ÅPEN' || status === '');

    // Hopp over lukkede eller fremtidige skjemaer for å spare ytelse
    if (!isActive || (rObj.StartDato && today < _midnight_(rObj.StartDato, tz))) {
      continue;
    }

    processed++;
    const dueHasDate = !!rObj.FristDato;
    const daysLeft = dueHasDate ? _daysDiff_(today, _midnight_(rObj.FristDato, tz)) : NaN;

    // Handlinger som kan være aktuelle i dag
    const isReminderDay = dueHasDate && Array.isArray(rObj.PaaminnelseDager) && rObj.PaaminnelseDager.includes(daysLeft);
    const isAutoCloseDay = rObj.AutoSteng === true && dueHasDate && today > _midnight_(rObj.FristDato, tz);
    const shouldUpdateStats = isActive; // oppdater kun aktive skjema

    // Åpne form kun dersom nødvendig
    if (!isReminderDay && !isAutoCloseDay && !shouldUpdateStats) continue;

    let form = null;
    const formId = rObj.FormId || _extractFormIdFromUrl_(rObj.FormURL);
    if (formId){
      try { form = FormApp.openById(formId); }
      catch(e){ if (typeof _logEvent === 'function') _logEvent('FormsScheduler','Åpning av form "' + rObj.SkjemaNavn + '" feilet: ' + e.message); }
    }
    if (!form) continue;

    // Oppdater svarstatistikk
    if (shouldUpdateStats) {
      const stats = _calcFormStats_(form);
      const statsChanged = _writeBackStats_(sh, i+2, map, stats, rObj);
      if (statsChanged) statsUpdated++;
    }

    // Påminnelse
    if (isReminderDay){
      const markerKey = 'D-' + daysLeft;
      if (!_reminderAlreadySent_(rObj.RemindersSent, markerKey)){
        const seg = (String(rObj.Segment||'').trim().toUpperCase() || 'STYRET');
        let to = [];
        if (seg === 'STYRET') to = _getBoardEmails_();
        const subject = `[${APP.NAME}] Påminnelse: ${rObj.SkjemaNavn} (frist ${_fmtDate_(rObj.FristDato, tz)})`;
        const body =
          `Hei,\n\nDette er en påminnelse om skjemaet "${rObj.SkjemaNavn}".\n` +
          (rObj.FormURL ? `Skjema: ${rObj.FormURL}\n` : '') +
          (rObj.FristDato ? `Frist: ${_fmtDate_(rObj.FristDato, tz)}\n` : '') +
          `\nMvh\n${APP.NAME}`;

        if (to.length){
          try {
            MailApp.sendEmail({ to: to.join(','), subject, body });
            if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Påminnelse D${daysLeft>=0?'-':''}${daysLeft} sendt til STYRET (${to.length}) for "${rObj.SkjemaNavn}".`);
            reminders++;
            _appendReminderMarker_(sh, i+2, map.remindersSent, markerKey, tz);
            _setCell_(sh, i+2, map.sisteUts, new Date());
          } catch(e){
            if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `E-post feilet for "${rObj.SkjemaNavn}": ` + e.message);
          }
        } else {
          if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Påminnelse (ikke sendt – segment=${seg}, ingen mottakere) for "${rObj.SkjemaNavn}".`);
        }
      }
    }

    // Auto-steng
    if (isAutoCloseDay){
      try {
        if (form.isAcceptingResponses()){
          form.setAcceptingResponses(false);
          if (map.status) _setCell_(sh, i+2, map.status, 'CLOSED');
          if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Auto-stengte "${rObj.SkjemaNavn}" (frist utløpt).`);
          closed++;
        }
      } catch(e){
        if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Auto-stenging feilet for "${rObj.SkjemaNavn}": ` + e.message);
      }
    }
  }

  // Enkel "last touched"-markør (NB: skriver i header-cellen)
  if (map.sistOppdatert) _setCell_(sh, 1, map.sistOppdatert, new Date());
  if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Kjøring ferdig. Aktive: ${processed}. Påminnelser: ${reminders}. Auto-stengt: ${closed}. Stats oppdatert: ${statsUpdated}.`);
}

/* ====================== Hjelpere (internt) ====================== */

/** Sørger for at register-arket finnes og har riktige kolonner. */
function _ensureFormsRegisterSheet_(){
  try{
    const headers = [
      'SkjemaNavn','FormId','FormURL','Status',
      'StartDato','FristDato','PaaminnelseDager','AutoSteng','Segment',
      'SvarTeller','SistSvarTs','SisteUtsendelse','RemindersSent','SistOppdatert'
    ];
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(FORMS_REG_SHEET);
    if (!sh) sh = ss.insertSheet(FORMS_REG_SHEET);

    const cur = sh.getRange(1,1,1,headers.length).getValues()[0];
    Sameie.Sheets.ensureHeader(sh, headers);
    return sh;
  } catch(e){
    if (typeof _logEvent === 'function') _logEvent('FormsScheduler', 'Klarte ikke sikre registerark: ' + e.message);
    return null;
  }
}

function _removeFormsDailySchedulerTrigger_(){
  const triggers = ScriptApp.getProjectTriggers() || [];
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'runFormsDailyScheduler'){
      ScriptApp.deleteTrigger(t);
    }
  });
}

function _headerMap_(sh){
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const norm = s => String(s||'').toLowerCase().replace(/[\s_]+/g,'').trim();
  const want = {
    skjemaNavn: ['skjemanavn','navn'],
    formId:     ['formid','id'],
    formURL:    ['formurl','url'],
    status:     ['status','aktiv'],
    startDato:  ['startdato','start','aktivfra'],
    fristDato:  ['fristdato','deadline','due'],
    paaminnelseDager: ['paaminnelsedager','påminnelsedager','reminderdays'],
    autoSteng:  ['autosteng','autostengning','autoclose'],
    segment:    ['segment','målgruppe','malgruppe'],
    svarTeller: ['svarteller','antallsvar','responses','count'],
    sistSvarTs: ['sistsvarts','lastresponse','lastrapporttid','sistsvar'],
    sisteUts:   ['sisteutsendelse','lastsent','notified'],
    remindersSent: ['reminderssent','sendtepåminnelser','påminnelserlogg'],
    sistOppdatert: ['sistoppdatert','oppdatert','lastupdated']
  };
  const res = {};
  for (let c=0; c<hdr.length; c++){
    const n = norm(hdr[c]);
    if (!n) continue;
    for (const key of Object.keys(want)){
      if (res[key] != null) continue;
      if (want[key].some(alias => n === alias)){ res[key] = c+1; }
    }
  }
  return res;
}

function _rowToObj_(row, map){
  const get = (idx) => (idx ? row[idx-1] : '');
  const parseBool = (v) => {
    const s = String(v).trim().toLowerCase();
    return (s === 'true' || s === '1' || s === 'ja' || s === 'yes' || v === true);
  };
  const parseCSVints = (v) => {
    if (v == null || v === '') return [];
    return String(v).split(',').map(x => parseInt(String(x).trim(),10)).filter(n => !isNaN(n));
  };
  const parseDate = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    if (!v) return null;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  };
  return {
    SkjemaNavn: get(map.skjemaNavn),
    FormId:     get(map.formId),
    FormURL:    get(map.formURL),
    Status:     get(map.status),
    StartDato:  parseDate(get(map.startDato)),
    FristDato:  parseDate(get(map.fristDato)),
    PaaminnelseDager: parseCSVints(get(map.paaminnelseDager)),
    AutoSteng:  parseBool(get(map.autoSteng)),
    Segment:    get(map.segment),
    SvarTeller: get(map.svarTeller),
    SistSvarTs: parseDate(get(map.sistSvarTs)),
    RemindersSent: get(map.remindersSent)
  };
}

function _calcFormStats_(form){
  if (!form) return { count:null, last:null };
  try{
    const res = form.getResponses();
    const n = res.length;
    const last = n ? res[n-1].getTimestamp() : null;
    return { count:n, last:last };
  }catch(e){
    if (typeof _logEvent === 'function') _logEvent('FormsScheduler','getResponses() feilet: ' + e.message);
    return { count:null, last:null };
  }
}

/** Skriv tilbake statistikk, returnerer true hvis noe ble endret. */
function _writeBackStats_(sh, rowIndex, map, newStats, oldRowObj){
  let changed = false;
  if (map.svarTeller && newStats.count != null && String(newStats.count) !== String(oldRowObj.SvarTeller)) {
    _setCell_(sh, rowIndex, map.svarTeller, newStats.count);
    changed = true;
  }
  if (map.sistSvarTs && newStats.last) {
    const oldTime = oldRowObj.SistSvarTs ? oldRowObj.SistSvarTs.getTime() : 0;
    if (newStats.last.getTime() !== oldTime) {
       _setCell_(sh, rowIndex, map.sistSvarTs, newStats.last);
       changed = true;
    }
  }
  if (changed && map.sistOppdatert){
    _setCell_(sh, rowIndex, map.sistOppdatert, new Date());
  }
  return changed;
}

function _appendReminderMarker_(sh, rowIndex, colIndex, markerKey, tz){
  if (!colIndex) return;
  const cur = String(sh.getRange(rowIndex, colIndex).getValue() || '').trim();
  const stamp = _fmtDate_(new Date(), tz);
  const add = `${markerKey}|${stamp}`;
  const nextVal = cur ? (cur + ';' + add) : add;
  sh.getRange(rowIndex, colIndex).setValue(nextVal);
}

function _reminderAlreadySent_(val, markerKey){
  if (!val) return false;
  const parts = String(val).split(';').map(s => String(s||'').trim());
  return parts.some(p => p.startsWith(markerKey + '|'));
}

function _extractFormIdFromUrl_(url){
  if (!url) return '';
  const m = String(url).match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

/** Enkel regex for å validere e-postformat. */
function _isValidEmail_(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

/** Hent e-poster fra Styret-fanen (filtrert til gyldige adresser). */
function _getBoardEmails_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.BOARD);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2,2, sh.getLastRow()-1, 1).getValues().flat();
  return vals.map(v => String(v || '').trim()).filter(_isValidEmail_);
}

/** Sett en enkelt celleverdi (1-basert), med logging ved feil. */
function _setCell_(sh, row, col, v){
  if (!row || !col) return;
  try { sh.getRange(row, col).setValue(v); }
  catch(e) { if (typeof _logEvent === 'function') _logEvent('FormsScheduler', `Skriving til celle (${row},${col}) feilet: ${e.message}`); }
}

/* -------------------- Datohjelpere -------------------- */
function _midnight_(d){
  if (!(d instanceof Date) || isNaN(d)) return null;
  const y = d.getFullYear(), m = d.getMonth(), day = d.getDate();
  return new Date(y, m, day, 0,0,0,0);
}
function _daysDiff_(from, to){
  if (!from || !to) return NaN;
  const MS = 24*60*60*1000;
  return Math.round((to.getTime() - from.getTime()) / MS);
}
function _fmtDate_(d, tz){
  if (!d) return '';
  try { return Utilities.formatDate(d, tz || 'Europe/Oslo', 'dd.MM.yyyy'); }
  catch(_) { return d.toISOString().slice(0,10); }
}

/* -------------------- Router for onFormSubmit -------------------- */
/**
 * Kjør én trigger "onFormSubmit" og la denne funksjonen rute videre.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function routeFormSubmit(e) {
  try {
    const svar = (e && e.namedValues) ? e.namedValues : {};

    // Eksempel: HMS-skjema (gjenkjent via unikt felt)
    if (svar['Sjekkpunkt: Røykvarsler'] || svar['Seksjonsnummer']) {
      if (typeof _logEvent === 'function') _logEvent('FormRouter', 'Ruter til handleHmsFormSubmit');
      if (typeof handleHmsFormSubmit === 'function') return handleHmsFormSubmit(e);
      if (typeof _logEvent === 'function') _logEvent('FormRouter_Mangler', 'handleHmsFormSubmit er ikke definert.');
      return;
    }

    // Eksempel: Support-skjema
    if (svar['Kategori for henvendelse']) {
      if (typeof _logEvent === 'function') _logEvent('FormRouter', 'Ruter til handleSupportFormSubmit');
      if (typeof handleSupportFormSubmit === 'function') return handleSupportFormSubmit(e);
      if (typeof _logEvent === 'function') _logEvent('FormRouter_Mangler', 'handleSupportFormSubmit er ikke definert.');
      return;
    }

    // Legg til flere matcher ved behov …

    if (typeof _logEvent === 'function') _logEvent('FormRouter_Warning', 'Skjemainnsending matchet ingen kjent rute.');
  } catch (err) {
    if (typeof _logEvent === 'function') _logEvent('FormRouter_KRITISK_FEIL', `Router-funksjonen feilet: ${err.message}`);
  }
}
