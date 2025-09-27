/* ====================== Vaktmester API (Backend) ======================
 * FILE: 18_Vaktmester_API.gs | VERSION: 1.3.0 | UPDATED: 2025-09-14
 * FORMÅL: Komplett og sikker backend for Vaktmester-UI.
 * - Henter aktive oppgaver og historikk
 * - Statusendringer + kommentering
 * - Vaktmester kan opprette egne saker
 * - Sikkerhet: kun ansvarlig kan endre
 * - Ytelse: rad-indeks i PropertiesService for O(1) oppslag
 * MERK: Bruker globalThis.SHEETS fra 00_App_Core.gs (ingen redeklarasjon).
 * ====================================================================== */

(function () {
  /* ---------- Avhengigheter / aliaser (ingen globale redeklarasjoner) ---------- */
  var SH = (globalThis.SHEETS || { TASKS: 'Oppgaver', BOARD: 'Styret' });
  var PROPS = PropertiesService.getScriptProperties();
  var TZ = Session.getScriptTimeZone() || 'Europe/Oslo';

  /* ---------- Små hjelpere ---------- */
  /*
   * MERK: Den lokale _safeLog_()-funksjonen er fjernet. Den globale
   * versjonen fra 00b_Utils.js blir nå brukt i stedet.
   */
  function _normalizeEmail_(s) {
    if (!s) return '';
    var str = String(s).trim();
    var m = str.match(/<([^>]+)>/); // f.eks. "Navn <mail@domene.no>"
    if (m) str = m[1];
    str = str.replace(/^mailto:/i, '');
    return str.toLowerCase();
  }
  function _hasAccess_() {
    // Tillat hvis RBAC ikke er aktiv, ellers krev eksplisitt rettighet
    if (typeof hasPermission === 'function') return !!hasPermission('VIEW_VAKTMESTER_UI');
    return true;
  }
  function _ensureTaskSheet_() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SH.TASKS);
    if (!sh) {
      sh = ss.insertSheet(SH.TASKS);
      var HDR = ['OppgaveID','Tittel','Beskrivelse','Seksjonsnr','Frist','Opprettet','Status','Prioritet','Ansvarlig','Kommentarer','Kilde','Kategori'];
      sh.getRange(1,1,1,HDR.length).setValues([HDR]).setFontWeight('bold'); sh.setFrozenRows(1);
    }
    return sh;
  }
  function _headersMap_(H) {
    var m = {}; for (var i=0;i<H.length;i++) m[String(H[i]||'')] = i; return m;
  }

  /* ---------- Rad-indeks (OppgaveID -> rad) ---------- */
  var VM_IDX = {
    key: function(){ return 'IDX::' + (SH.TASKS || 'Oppgaver'); },
    get: function() {
      var raw = PROPS.getProperty(this.key());
      if (!raw) return this.rebuild();
      try { var o = JSON.parse(raw); return (o && typeof o==='object') ? o : this.rebuild(); }
      catch(_) { return this.rebuild(); }
    },
    put: function(id, row){
      var m = this.get(); m[String(id)] = row; PROPS.setProperty(this.key(), JSON.stringify(m));
    },
    del: function(id){
      var m = this.get(); delete m[String(id)]; PROPS.setProperty(this.key(), JSON.stringify(m));
    },
    rebuild: function(){
      var sh = SpreadsheetApp.getActive().getSheetByName(SH.TASKS);
      var map = {};
      if (sh && sh.getLastRow() > 1) {
        var data = sh.getDataRange().getValues(); var H = data.shift();
        var cId = H.indexOf('OppgaveID');
        if (cId >= 0) {
          for (var i=0;i<data.length;i++) { var id = data[i][cId]; if (id) map[id] = i+2; }
        }
      }
      PROPS.setProperty(this.key(), JSON.stringify(map));
      _safeLog_('VaktmesterIndex', 'Rebuild OK (' + Object.keys(map).length + ' nøkler)');
      return map;
    }
  };

  /* ---------- API ---------- */

  /** Profil (navn slås opp i Styret-ark om mulig) */
  function getCurrentVaktmesterProfile() {
    try {
      var email = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());
      var name = '';
      try {
        var boardSheet = SpreadsheetApp.getActive().getSheetByName(SH.BOARD);
        if (boardSheet && boardSheet.getLastRow() > 1) {
          var vals = boardSheet.getRange(2,1,boardSheet.getLastRow()-1,2).getValues();
          for (var i=0;i<vals.length;i++){
            if (_normalizeEmail_(vals[i][1]) === email) { name = vals[i][0]; break; }
          }
        }
      } catch(_) {}
      return { name: name, email: email };
    } catch (e) {
      return { name: '', email: 'Ukjent' };
    }
  }

  /** Liste oppgaver for innlogget (kind: 'active' | 'history') */
  function getTasksForVaktmester(kind) {
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      kind = (String(kind||'active').toLowerCase());
      var userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());
      if (!userEmail) throw new Error('Kunne ikke identifisere brukeren.');

      var sh = _ensureTaskSheet_();
      if (sh.getLastRow() < 2) return { ok:true, items:[] };

      var data = sh.getDataRange().getValues(); var H = data.shift(); var c = _headersMap_(H);

      var targetStatuses = (kind === 'active')
        ? ['ny','påbegynt','venter']
        : ['fullført','avvist','lukket','ferdig'];

      var out = [];
      for (var i=0;i<data.length;i++){
        var r = data[i];
        var ansvarlig = _normalizeEmail_(r[c['Ansvarlig']]);
        var status = String(r[c['Status']]||'').toLowerCase();
        if (ansvarlig !== userEmail) continue;
        if (targetStatuses.indexOf(status) === -1) continue;

        var opprettet = (r[c['Opprettet']] instanceof Date) ? r[c['Opprettet']].toISOString() : '';
        var frist     = (r[c['Frist']]     instanceof Date) ? r[c['Frist']].toISOString()     : '';

        out.push({
          id: r[c['OppgaveID']],
          tittel: r[c['Tittel']],
          beskrivelse: r[c['Beskrivelse']],
          status: r[c['Status']],
          opprettetISO: opprettet,
          fristISO: frist,
          seksjon: r[c['Seksjonsnr']],
          prioritet: r[c['Prioritet']]
        });
      }

      out.sort(function(a,b){
        var av = a.opprettetISO ? new Date(a.opprettetISO).getTime() : 0;
        var bv = b.opprettetISO ? new Date(b.opprettetISO).getTime() : 0;
        return bv - av;
      });

      return { ok:true, items: out };
    } catch (e) {
      _safeLog_('VaktmesterAPI_Feil', 'getTasksForVaktmester: ' + e.message);
      return { ok:false, error:'Kunne ikke hente oppgavelisten.' };
    }
  }

  /** Oppdater status (Fullført/Avvist) og/eller legg til kommentar. Kun ansvarlig. */
  function updateTaskStatusByVaktmester(taskId, newStatus, comment) {
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      var userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());

      var norm = String(newStatus||'').trim();
      var valid = ['Fullført','Avvist'];
      if (norm && valid.indexOf(norm) === -1) throw new Error('Ugyldig status.');

      var sh = _ensureTaskSheet_();
      var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      var c = _headersMap_(H);

      // Slå opp via indeks (O(1)), fall tilbake til TextFinder
      var idx = VM_IDX.get();
      var rowNum = idx[taskId];
      if (!rowNum) {
        var tf = sh.createTextFinder(String(taskId)).matchEntireCell(true).findNext();
        if (!tf) throw new Error('Fant ikke oppgave med ID: ' + taskId);
        rowNum = tf.getRow();
        VM_IDX.rebuild();
      }

      var row = sh.getRange(rowNum,1,1,sh.getLastColumn()).getValues()[0];
      if (_normalizeEmail_(row[c['Ansvarlig']]) !== userEmail) {
        throw new Error('Tilgang nektet. Du er ikke ansvarlig for denne oppgaven.');
      }

      if (norm) sh.getRange(rowNum, c['Status']+1).setValue(norm);

      var trimmed = String(comment||'').trim();
      if (trimmed && c['Kommentarer'] > -1) {
        var existing = row[c['Kommentarer']] || '';
        var ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
        var line = '[' + ts + ' - ' + userEmail + ']: ' + trimmed;
        var joined = existing ? (existing + '\n' + line) : line;
        sh.getRange(rowNum, c['Kommentarer']+1).setValue(joined);
      }

      _safeLog_('Oppgave_Status', 'Vaktmester ' + userEmail + ' oppdaterte ' + taskId);
      return { ok:true, message:'Oppgave ' + taskId + ' er oppdatert.' };
    } catch (e) {
      _safeLog_('VaktmesterAPI_Feil', 'updateTaskStatus: ' + e.message);
      throw e;
    } finally { lock.releaseLock(); }
  }

  /** Legg til kommentar (uten å endre status) */
  function addTaskCommentByVaktmester(taskId, comment) {
    return updateTaskStatusByVaktmester(taskId, null, comment);
  }

  /** Opprett ny oppgave tildelt innlogget vaktmester */
  function createVaktmesterTask(payload) {
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      var user = getCurrentVaktmesterProfile();
      if (!payload || !String(payload.tittel||'').trim()) throw new Error('Tittel er påkrevd.');

      var sh = _ensureTaskSheet_();
      var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      var c = _headersMap_(H);

      var newId = (typeof _nextTaskId_ === 'function') ? _nextTaskId_() : ('TASK-' + new Date().getTime());
      var row = new Array(H.length).fill('');

      row[c['OppgaveID']]   = newId;
      row[c['Tittel']]      = payload.tittel;
      row[c['Beskrivelse']] = payload.beskrivelse || '';
      row[c['Seksjonsnr']]  = payload.seksjonsnr || '';
      row[c['Frist']]       = payload.frist ? new Date(payload.frist) : '';
      row[c['Opprettet']]   = new Date();
      row[c['Status']]      = 'Ny';
      row[c['Prioritet']]   = payload.prioritet || 'Medium';
      row[c['Ansvarlig']]   = user.email;
      if (c['Kilde'] > -1)    row[c['Kilde']]    = 'Vaktmester-UI';
      if (c['Kategori'] > -1) row[c['Kategori']] = 'Vaktmester';

      sh.appendRow(row);
      VM_IDX.put(newId, sh.getLastRow());

      _safeLog_('Oppgave_Opprettet', 'Vaktmester ' + user.email + ' opprettet ' + newId);
      return { ok:true, id:newId };
    } catch (e) {
      _safeLog_('VaktmesterAPI_Feil', 'createVaktmesterTask: ' + e.message);
      throw e;
    } finally { lock.releaseLock(); }
  }

  /* ---------- Eksporter globale API-navn ---------- */
  globalThis.getCurrentVaktmesterProfile   = getCurrentVaktmesterProfile;
  globalThis.getTasksForVaktmester         = getTasksForVaktmester;
  globalThis.updateTaskStatusByVaktmester  = updateTaskStatusByVaktmester;
  globalThis.addTaskCommentByVaktmester    = addTaskCommentByVaktmester;
  globalThis.createVaktmesterTask          = createVaktmesterTask;
})();
