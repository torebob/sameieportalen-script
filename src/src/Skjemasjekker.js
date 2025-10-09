/**
 * FormChecksPlus (Skjemasjekker)
 * Versjon: 1.2.0
 *
 * Funksjoner:
 *  - runFormsDailyScheduler()      : Daglig sjekk av registrerte skjemaer (Skjema-register)
 *  - routeFormSubmit(e)            : Inngangspunkt ved skjema-innsending (Forms/Sheets)
 *  - sendScheduledReminders()      : Tidsstyrte påminnelser basert på Skjema-register
 *  - tests_onSubmitRunner(e)       : Enkelt testrun for onSubmit-hendelser
 *  - setupFormsDailySchedulerTrigger() / clearFormsDailySchedulerTrigger()
 *
 * Forventede ark:
 *  - "Skjema-register": SkjemaNavn | Form ID | URL | Målark | Handler | Status | SistOppdatert
 *  - "Rapport": Kj.Dato | Kategori | Nøkkel | Status | Detaljer
 *
 * Valgfritt:
 *  - LoggerPlus / getAppLogger_(): Om definert i prosjektet brukes den til strukturert logging.
 *    Fallback til console.* dersom ikke tilgjengelig.
 */

///////////////////////////// Konfigurasjon /////////////////////////////////////

const FORM_CHECKS_CFG = {
  VERSION: '1.2.0',
  SHEET_NAMES: {
    REGISTER: 'Skjema-register',
    RAPPORT: 'Rapport'
  },
  REGISTER_HEADERS: [
    'SkjemaNavn', 'Form ID', 'URL', 'Målark', 'Handler', 'Status', 'SistOppdatert'
  ],
  RAPPORT_HEADERS: ['Kj.Dato', 'Kategori', 'Nøkkel', 'Status', 'Detaljer'],
  // Kolonnenavn som brukes internt (mapping lages fra header)
  COLS: {
    NAVN: 'SkjemaNavn',
    FORM_ID: 'Form ID',
    URL: 'URL',
    MAALARK: 'Målark',
    HANDLER: 'Handler',
    STATUS: 'Status',
    SIST: 'SistOppdatert'
  },
  // Påminnelser / markører
  REMINDER_MARKER_PREFIX: 'REMINDER_SENT_',
  // Valg for daglig planlegger
  DAILY_TRIGGER_HOUR: 6, // 06:00 lokaltid
  // Rapportering
  REPORT_CATEGORY: 'Skjemasjekk'
};

///////////////////////////// Logging (trygg) ///////////////////////////////////

function _formLog_() {
  try {
    if (typeof getAppLogger_ === 'function') return getAppLogger_();
  } catch (_) {}
  return {
    info: (fn, msg, d) => { try { console.log('[INFO]', fn || '', msg || '', d || ''); } catch (_) {} },
    warn: (fn, msg, d) => { try { console.warn('[WARN]', fn || '', msg || '', d || ''); } catch (_) {} },
    error: (fn, msg, d) => { try { console.error('[ERROR]', fn || '', msg || '', d || ''); } catch (_) {} }
  };
}

///////////////////////////// Offentlige API ////////////////////////////////////

/**
 * Daglig scheduler: validerer register, sjekker skjemaer, oppdaterer status og skriver rapportlinjer.
 */
function runFormsDailyScheduler() {
  var log = _formLog_(); var fn = 'runFormsDailyScheduler';
  var started = Date.now();
  try {
    var ss = SpreadsheetApp.getActive();
    var register = _ensureFormsRegisterSheet_(ss);
    var rows = _readRegisterAsObjects_(register);
    var results = [];

    rows.forEach(function(r, idx) {
      var itemKey = r[FORM_CHECKS_CFG.COLS.NAVN] || ('#' + (idx + 2));
      try {
        var formId = _extractFormIdFromUrl_(r[FORM_CHECKS_CFG.COLS.URL], r[FORM_CHECKS_CFG.COLS.FORM_ID]);
        var status = _computeFormStatus_(formId, r);
        _writeRegisterStatus_(register, r._rowIndex, status);
        results.push({ key: itemKey, status: 'OK', details: status });
      } catch (e1) {
        log.warn(fn, 'Skjema feilet under sjekk', { key: itemKey, error: e1 && e1.message });
        _writeRegisterStatus_(register, r._rowIndex, 'FEIL: ' + (e1 && e1.message));
        results.push({ key: itemKey, status: 'FAIL', details: e1 && e1.message });
      }
    });

    _appendRapport_(
      ss,
      FORM_CHECKS_CFG.REPORT_CATEGORY,
      'Daglig',
      'OK',
      'Antall sjekket: ' + results.length
    );

    log.info(fn, 'Ferdig daglig forms-sjekk', {
      ms: Date.now() - started,
      count: results.length
    });
  } catch (e) {
    _appendRapport_(SpreadsheetApp.getActive(), FORM_CHECKS_CFG.REPORT_CATEGORY, 'Daglig', 'FAIL', e && e.message);
    log.error(fn, 'Kritisk feil i runFormsDailyScheduler()', { error: e && e.message, stack: e && e.stack });
  }
}

/**
 * Inngangspunkt ved innsending av skjema (Forms/Sheets onSubmit).
 * Ruter innsendingen til angitt handler i registeret (kolonnen "Handler").
 * @param {Object} e Event-objektet fra onFormSubmit/onSubmit.
 */
function routeFormSubmit(e) {
  var log = _formLog_(); var fn = 'routeFormSubmit';
  try {
    var ss = SpreadsheetApp.getActive();
    var register = _ensureFormsRegisterSheet_(ss);
    var rows = _readRegisterAsObjects_(register);
    var formId = _inferSubmitFormId_(e) || '';

    if (!formId) {
      log.warn(fn, 'Fant ikke formId fra event. Hopper over.', {});
      return;
    }
    // Finn første match i registeret på Form ID
    var entry = rows.find(function(r) {
      var rid = _extractFormIdFromUrl_(r[FORM_CHECKS_CFG.COLS.URL], r[FORM_CHECKS_CFG.COLS.FORM_ID]);
      return rid && rid === formId;
    });

    if (!entry) {
      log.warn(fn, 'Ingen match i Skjema-register for formId', { formId: formId });
      return;
    }
    var handlerName = (entry[FORM_CHECKS_CFG.COLS.HANDLER] || '').trim();
    if (handlerName && typeof this[handlerName] === 'function') {
      // Kall handler med e
      this[handlerName](e);
      log.info(fn, 'Handler utført', { handler: handlerName, formId: formId });
    } else {
      log.warn(fn, 'Handler ikke definert eller ikke en funksjon', { handler: handlerName, formId: formId });
    }

    // Oppdater registerets SistOppdatert
    _writeRegisterUpdated_(register, entry._rowIndex, new Date());

  } catch (err) {
    log.error(fn, 'Feil under routeFormSubmit', { error: err && err.message, stack: err && err.stack });
  }
}

/**
 * Sender planlagte påminnelser basert på registeret.
 * Forenklet eksempel: sjekker etter rader med Status ~ 'PÅMINNELSE' og markerer utsending.
 */
function sendScheduledReminders() {
  var log = _formLog_(); var fn = 'sendScheduledReminders';
  try {
    var ss = SpreadsheetApp.getActive();
    var register = _ensureFormsRegisterSheet_(ss);
    var rows = _readRegisterAsObjects_(register);
    var sent = 0;

    rows.forEach(function(r) {
      var status = (r[FORM_CHECKS_CFG.COLS.STATUS] || '').toString().toUpperCase();
      var formId = _extractFormIdFromUrl_(r[FORM_CHECKS_CFG.COLS.URL], r[FORM_CHECKS_CFG.COLS.FORM_ID]);
      if (!formId) return;

      // Eksempel-policy: send påminnelse dersom status inneholder "PÅMINNELSE" og ikke allerede sendt i dag
      if (status.indexOf('PÅMINNELSE') >= 0 && !_reminderAlreadySent_(register, r._rowIndex)) {
        // … her kunne du sende e-post / lag varsling / oppslag …
        _appendReminderMarker_(register, r._rowIndex);
        sent++;
      }
    });

    _appendRapport_(ss, FORM_CHECKS_CFG.REPORT_CATEGORY, 'Påminnelse', 'OK', 'Sendt: ' + sent);
    log.info(fn, 'Påminnelser ferdig', { sent: sent });
  } catch (e) {
    _appendRapport_(SpreadsheetApp.getActive(), FORM_CHECKS_CFG.REPORT_CATEGORY, 'Påminnelse', 'FAIL', e && e.message);
    log.error(fn, 'Feil i sendScheduledReminders', { error: e && e.message, stack: e && e.stack });
  }
}

/**
 * Enkel testrunner for onSubmit-logikk.
 * @param {Object} e Event-objekt (kan være syntetisk i test)
 */
function tests_onSubmitRunner(e) {
  var log = _formLog_(); var fn = 'tests_onSubmitRunner';
  try {
    // Kjør routeFormSubmit med et syntetisk event hvis e mangler
    var evt = e || {
      namedValues: {},
      range: null,
      source: SpreadsheetApp.getActive(),
      triggerUid: 'TEST_' + new Date().getTime()
    };
    routeFormSubmit(evt);
    log.info(fn, 'Test gjennomført', { ok: true });
  } catch (err) {
    log.error(fn, 'Test feilet', { error: err && err.message, stack: err && err.stack });
  }
}

/** Installer daglig trigger (06:00 lokaltid). */
function setupFormsDailySchedulerTrigger() {
  var log = _formLog_(); var fn = 'setupFormsDailySchedulerTrigger';
  try {
    clearFormsDailySchedulerTrigger(); // rydd opp duplikater
    ScriptApp.newTrigger('runFormsDailyScheduler')
      .timeBased()
      .atHour(FORM_CHECKS_CFG.DAILY_TRIGGER_HOUR)
      .everyDays(1)
      .create();
    log.info(fn, 'Daglig trigger installert', { hour: FORM_CHECKS_CFG.DAILY_TRIGGER_HOUR });
  } catch (e) {
    log.error(fn, 'Klarte ikke installere daglig trigger', { error: e && e.message });
  }
}

/** Fjern daglig trigger for runFormsDailyScheduler(). */
function clearFormsDailySchedulerTrigger() {
  var log = _formLog_(); var fn = 'clearFormsDailySchedulerTrigger';
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'runFormsDailyScheduler') {
        ScriptApp.deleteTrigger(t);
      }
    });
    log.info(fn, 'Daglig trigger fjernet');
  } catch (e) {
    log.error(fn, 'Feil ved fjerning av trigger', { error: e && e.message });
  }
}

///////////////////////////// Private helpers ///////////////////////////////////

/** Sørg for at "Skjema-register" finnes og har headere. */
function _ensureFormsRegisterSheet_(ss) {
  var sh = ss.getSheetByName(FORM_CHECKS_CFG.SHEET_NAMES.REGISTER);
  if (!sh) {
    sh = ss.insertSheet(FORM_CHECKS_CFG.SHEET_NAMES.REGISTER);
    sh.appendRow(FORM_CHECKS_CFG.REGISTER_HEADERS);
    sh.setFrozenRows(1);
  } else {
    // sikre headere, men ikke overskriv eksisterende
    var hdr = (sh.getLastRow() > 0) ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : [];
    if (!hdr || hdr.length < FORM_CHECKS_CFG.REGISTER_HEADERS.length) {
      // forsøk å sette minst de forventede
      var to = FORM_CHECKS_CFG.REGISTER_HEADERS.slice();
      sh.getRange(1,1,1,to.length).setValues([to]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/** Lese registeret til objekter med header-mapping. Fester også _rowIndex. */
function _readRegisterAsObjects_(sh) {
  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return [];
  var header = vals[0];
  var hmap = _headerMap_(header);
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    if (_emptyRow_(row)) continue;
    var obj = _rowToObj_(row, hmap);
    obj._rowIndex = r + 1; // 1-basert
    out.push(obj);
  }
  return out;
}

/** Skriv status + sist oppdatert i registerrad. */
function _writeRegisterStatus_(sh, rowIndex, statusText) {
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var hmap = _headerMap_(header);
  var statusIdx = hmap[FORM_CHECKS_CFG.COLS.STATUS];
  var sistIdx = hmap[FORM_CHECKS_CFG.COLS.SIST];

  var updates = [];
  var lastCol = sh.getLastColumn();

  if (typeof statusIdx === 'number') {
    updates.push({ col: statusIdx + 1, val: statusText });
  }
  if (typeof sistIdx === 'number') {
    updates.push({ col: sistIdx + 1, val: new Date() });
  }
  if (updates.length) {
    var rng = sh.getRange(rowIndex, 1, 1, lastCol);
    var row = rng.getValues()[0];
    updates.forEach(function(u){ row[u.col - 1] = u.val; });
    rng.setValues([row]);
  }
}

/** Skriv kun SistOppdatert. */
function _writeRegisterUpdated_(sh, rowIndex, when) {
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idx = _headerMap_(header)[FORM_CHECKS_CFG.COLS.SIST];
  if (typeof idx !== 'number') return;
  sh.getRange(rowIndex, idx + 1).setValue(when || new Date());
}

/** Ekstraher Form ID fra URL (eller bruk gitt formId) */
function _extractFormIdFromUrl_(url, fallbackFormId) {
  var u = (url || '').toString().trim();
  if (u) {
    // Støtter både /forms/d/e/<FORM_ID>/viewform og /d/<ID>/edit
    var m = u.match(/\/forms\/d\/e\/([a-zA-Z0-9_-]+)/) || u.match(/\/d\/([a-zA-Z0-9_-]+)\//);
    if (m && m[1]) return m[1];
  }
  return (fallbackFormId || '').toString().trim() || '';
}

/** Avgjør status for et skjema. (Forenklet eksempel – kan utvides etter behov.) */
function _computeFormStatus_(formId, registerRowObj) {
  if (!formId) return 'UKJENT_ID';
  // Eksempel: sjekk at skjema eksisterer og er aktivt (enkelt-check via UrlFetch kan være begrenset i GAS).
  // Her setter vi status “AKTIV” hvis formId ikke er tomt; du kan utvide med faktiske kall/valideringer.
  return 'AKTIV';
}

/** Finn formId fra onSubmit-event (best effort). */
function _inferSubmitFormId_(e) {
  try {
    // Hvis innsending kommer fra Google Forms → e.source er vanligvis et Spreadsheet (destinasjon),
    // ikke direkte Form. Vi forsøker å matche mot registeret via URL senere.
    // Om du legger ved formId i et skjema-felt, plukk det her:
    if (e && e.namedValues) {
      // typisk felt navn: "Form ID" eller "Skjema ID"
      var keys = Object.keys(e.namedValues);
      for (var i=0; i<keys.length; i++) {
        var k = (keys[i] || '').toLowerCase();
        if (k.indexOf('form id') >= 0 || k.indexOf('skjema id') >= 0) {
          var v = e.namedValues[keys[i]];
          if (Array.isArray(v) && v[0]) return (v[0] || '').toString().trim();
          if (v) return (v || '').toString().trim();
        }
      }
    }
  } catch (_) {}
  return '';
}

/** Sjekk om påminnelse allerede sendt “i dag” til rad (bruker en marker-kolonne dynamisk). */
function _reminderAlreadySent_(sh, rowIndex) {
  var todayKey = FORM_CHECKS_CFG.REMINDER_MARKER_PREFIX + _yyyymmdd_(new Date());
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var map = _headerMap_(header);
  var idx = map[todayKey];
  if (typeof idx !== 'number') return false;
  var val = sh.getRange(rowIndex, idx + 1).getValue();
  return !!val;
}

/** Marker at påminnelse er sendt i dag (oppretter kolonne om nødvendig). */
function _appendReminderMarker_(sh, rowIndex) {
  var todayKey = FORM_CHECKS_CFG.REMINDER_MARKER_PREFIX + _yyyymmdd_(new Date());
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var map = _headerMap_(header);
  var idx = map[todayKey];
  if (typeof idx !== 'number') {
    // legg til ny kolonne på slutten
    var lastCol = sh.getLastColumn();
    sh.getRange(1, lastCol + 1).setValue(todayKey);
    idx = lastCol; // null-basert
  }
  sh.getRange(rowIndex, idx + 1).setValue(true);
}

/** Legg linje til "Rapport". Oppretter arket ved behov. */
function _appendRapport_(ss, kategori, nokkel, status, detaljer) {
  var sh = ss.getSheetByName(FORM_CHECKS_CFG.SHEET_NAMES.RAPPORT);
  if (!sh) {
    sh = ss.insertSheet(FORM_CHECKS_CFG.SHEET_NAMES.RAPPORT);
    sh.appendRow(FORM_CHECKS_CFG.RAPPORT_HEADERS);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), kategori, nokkel, status, detaljer || '']);
}

///////////////////////////// Små verktøy ///////////////////////////////////////

/** Header → index-map (case-sensitive match på eksakte headertekster). */
function _headerMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = (headers[i] || '').toString().trim();
    if (key) map[key] = i;
  }
  return map;
}

/** Rad → objekt iht. headerMap. */
function _rowToObj_(row, hmap) {
  var obj = {};
  Object.keys(hmap).forEach(function(h) {
    obj[h] = row[hmap[h]];
  });
  return obj;
}

/** Sjekk om rad er "tom". */
function _emptyRow_(row) {
  for (var i=0; i<row.length; i++) {
    var v = row[i];
    if (v !== '' && v !== null && typeof v !== 'undefined') return false;
  }
  return true;
}

/** Dato → YYYYMMDD. */
function _yyyymmdd_(d) {
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return '' + y + m + day;
}
