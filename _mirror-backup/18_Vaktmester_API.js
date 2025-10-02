/* ====================== Vaktmester API (Backend) ======================
 * FILE: 18_Vaktmester_API.gs | VERSION: 2.2.0 | UPDATED: 2025-09-28
 * FORMÅL: Komplett og sikker backend for Vaktmester-UI.
 * - Modernisert med let/const, arrow functions, og forbedret lesbarhet.
 * - Henter aktive oppgaver og historikk
 * - Statusendringer + kommentering
 * - Vaktmester kan opprette egne saker
 * - Sikkerhet: kun ansvarlig kan endre
 * - Ytelse: rad-indeks i PropertiesService for O(1) oppslag
 * ENDRINGER v2.2.0:
 *  - Lagt til funksjonalitet for opplasting av vedlegg til oppgaver.
 *  - Opprettet hjelpefunksjon for å hente/opprette vedleggsmappe i Drive.
 *  - Utvidet `getTasksForVaktmester` til å returnere vedlegg.
 * ENDRINGER v2.1.0:
 *  - Byttet til standardiserte navn på hjelpefunksjoner (safeLog).
 * ====================================================================== */

(() => {
  const SH = globalThis.SHEETS || { TASKS: 'Oppgaver', BOARD: 'Styret' };
  const PROPS = PropertiesService.getScriptProperties();
  const TZ = Session.getScriptTimeZone() || 'Europe/Oslo';
  const ATTACHMENT_FOLDER_NAME = 'Vaktmester Vedlegg';

  /* ---------- Hjelpefunksjoner ---------- */
  const _normalizeEmail_ = (s) => {
    if (!s) return '';
    let str = String(s).trim();
    const m = str.match(/<([^>]+)>/); // f.eks. "Navn <mail@domene.no>"
    if (m) str = m[1];
    str = str.replace(/^mailto:/i, '');
    return str.toLowerCase();
  };

  const _hasAccess_ = () => {
    if (typeof hasPermission === 'function') return !!hasPermission('VIEW_VAKTMESTER_UI');
    return true; // Fallback hvis RBAC-modul ikke er tilgjengelig
  };

  const _ensureTaskSheet_ = () => {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SH.TASKS);
    if (!sh) {
      sh = ss.insertSheet(SH.TASKS);
      const HDR = ['OppgaveID', 'Tittel', 'Beskrivelse', 'Seksjonsnr', 'Frist', 'Opprettet', 'Status', 'Prioritet', 'Ansvarlig', 'Kommentarer', 'Vedlegg', 'Kilde', 'Kategori'];
      sh.getRange(1, 1, 1, HDR.length).setValues([HDR]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  };

  const _getAttachmentFolder_ = () => {
    const folders = DriveApp.getFoldersByName(ATTACHMENT_FOLDER_NAME);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(ATTACHMENT_FOLDER_NAME);
  };

  const _headersMap_ = (H) => H.reduce((acc, header, i) => {
    const key = String(header || '').trim();
    if (key) acc[key] = i;
    return acc;
  }, {});

  /* ---------- Rad-indeks (OppgaveID -> rad) ---------- */
  const VM_IDX = {
    key: () => `IDX::${SH.TASKS || 'Oppgaver'}`,
    get() {
      const raw = PROPS.getProperty(this.key());
      if (!raw) return this.rebuild();
      try {
        const o = JSON.parse(raw);
        return (o && typeof o === 'object') ? o : this.rebuild();
      } catch (e) {
        return this.rebuild();
      }
    },
    put(id, row) {
      const m = this.get();
      m[String(id)] = row;
      PROPS.setProperty(this.key(), JSON.stringify(m));
    },
    del(id) {
      const m = this.get();
      delete m[String(id)];
      PROPS.setProperty(this.key(), JSON.stringify(m));
    },
    rebuild() {
      const sh = _ensureTaskSheet_();
      const map = {};
      if (sh.getLastRow() > 1) {
        const data = sh.getDataRange().getValues();
        const H = data.shift();
        const cId = H.indexOf('OppgaveID');
        if (cId >= 0) {
          data.forEach((row, i) => {
            const id = row[cId];
            if (id) map[id] = i + 2;
          });
        }
      }
      PROPS.setProperty(this.key(), JSON.stringify(map));
      safeLog('VaktmesterIndex', `Rebuild OK (${Object.keys(map).length} nøkler)`);
      return map;
    },
  };

  /* ---------- API ---------- */
  function getCurrentVaktmesterProfile() {
    try {
      const email = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());
      let name = '';
      try {
        const boardSheet = SpreadsheetApp.getActive().getSheetByName(SH.BOARD);
        if (boardSheet && boardSheet.getLastRow() > 1) {
          const vals = boardSheet.getRange(2, 1, boardSheet.getLastRow() - 1, 2).getValues();
          const match = vals.find(row => _normalizeEmail_(row[1]) === email);
          if (match) name = match[0];
        }
      } catch (e) {
        // Ignorer feil ved lesing av Styret-arket
      }
      return { name, email };
    } catch (e) {
      return { name: '', email: 'Ukjent' };
    }
  }

  function getTasksForVaktmester(kind = 'active') {
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');

      const userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());
      if (!userEmail) throw new Error('Kunne ikke identifisere brukeren.');

      const sh = _ensureTaskSheet_();
      if (sh.getLastRow() < 2) return { ok: true, items: [] };

      const data = sh.getDataRange().getValues();
      const H = data.shift();
      const c = _headersMap_(H);

      const targetStatuses = (kind.toLowerCase() === 'active')
        ? new Set(['ny', 'påbegynt', 'venter'])
        : new Set(['fullført', 'avvist', 'lukket', 'ferdig']);

      const items = data
        .filter(r => {
          const ansvarlig = _normalizeEmail_(r[c.Ansvarlig]);
          const status = String(r[c.Status] || '').toLowerCase();
          return ansvarlig === userEmail && targetStatuses.has(status);
        })
        .map(r => {
          let attachments = [];
          if (c.Vedlegg > -1 && r[c.Vedlegg]) {
            try {
              attachments = JSON.parse(r[c.Vedlegg]);
              if (!Array.isArray(attachments)) attachments = [];
            } catch (e) { /* ignore parse error */ }
          }
          return {
            id: r[c.OppgaveID],
            tittel: r[c.Tittel],
            beskrivelse: r[c.Beskrivelse],
            status: r[c.Status],
            opprettetISO: (r[c.Opprettet] instanceof Date) ? r[c.Opprettet].toISOString() : '',
            fristISO: (r[c.Frist] instanceof Date) ? r[c.Frist].toISOString() : '',
            seksjon: r[c.Seksjonsnr],
            prioritet: r[c.Prioritet] || '—',
            attachments: attachments,
          };
        })
        .sort((a, b) => (new Date(b.opprettetISO).getTime() || 0) - (new Date(a.opprettetISO).getTime() || 0));

      return { ok: true, items };
    } catch (e) {
      safeLog('VaktmesterAPI_Feil', `getTasksForVaktmester: ${e.message}`);
      return { ok: false, error: 'Kunne ikke hente oppgavelisten.' };
    }
  }

  function updateTaskStatusByVaktmester(taskId, newStatus, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      const userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());

      const normStatus = String(newStatus || '').trim();
      const validStatuses = ['Fullført', 'Avvist'];
      if (normStatus && !validStatuses.includes(normStatus)) {
        throw new Error('Ugyldig status.');
      }

      const sh = _ensureTaskSheet_();
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const c = _headersMap_(H);

      const idx = VM_IDX.get();
      let rowNum = idx[taskId];
      if (!rowNum) {
        const tf = sh.createTextFinder(String(taskId)).matchEntireCell(true).findNext();
        if (!tf) throw new Error(`Fant ikke oppgave med ID: ${taskId}`);
        rowNum = tf.getRow();
        VM_IDX.rebuild();
      }

      const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn());
      const rowVals = row.getValues()[0];
      if (_normalizeEmail_(rowVals[c.Ansvarlig]) !== userEmail) {
        throw new Error('Tilgang nektet. Du er ikke ansvarlig for denne oppgaven.');
      }

      if (normStatus) rowVals[c.Status] = normStatus;

      const trimmedComment = String(comment || '').trim();
      if (trimmedComment && c.Kommentarer > -1) {
        const existing = rowVals[c.Kommentarer] || '';
        const timestamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
        const line = `[${timestamp} - ${userEmail}]: ${trimmedComment}`;
        rowVals[c.Kommentarer] = existing ? `${existing}\n${line}` : line;
      }

      row.setValues([rowVals]);

      safeLog('Oppgave_Status', `Vaktmester ${userEmail} oppdaterte ${taskId}`);
      return { ok: true, message: `Oppgave ${taskId} er oppdatert.` };
    } catch (e) {
      safeLog('VaktmesterAPI_Feil', `updateTaskStatus: ${e.message}`);
      throw e;
    } finally {
      lock.releaseLock();
    }
  }

  function addTaskCommentByVaktmester(taskId, comment) {
    return updateTaskStatusByVaktmester(taskId, null, comment);
  }

  function uploadAttachmentByVaktmester(taskId, fileObject) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      if (!fileObject || !fileObject.fileName || !fileObject.data || !fileObject.mimeType) {
        throw new Error('Ugyldig filobjekt.');
      }
      const userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail());

      const sh = _ensureTaskSheet_();
      let H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      let c = _headersMap_(H);

      // Sjekk om bruker er ansvarlig
      const idx = VM_IDX.get();
      const rowNum = idx[taskId];
      if (!rowNum) throw new Error(`Fant ikke oppgave-rad for ID ${taskId}`);
      const ansvarligEmail = _normalizeEmail_(sh.getRange(rowNum, c.Ansvarlig + 1).getValue());
      if (ansvarligEmail !== userEmail) {
        throw new Error('Tilgang nektet. Du er ikke ansvarlig for denne oppgaven.');
      }

      // Håndter filopplasting
      const folder = _getAttachmentFolder_();
      const decodedData = Utilities.base64Decode(fileObject.data, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decodedData, fileObject.mimeType, fileObject.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const newAttachment = { name: file.getName(), url: file.getUrl() };

      // Legg til 'Vedlegg' kolonne hvis den mangler
      if (c.Vedlegg === undefined) {
        sh.insertColumnAfter(c.Kommentarer + 1);
        sh.getRange(1, c.Kommentarer + 2).setValue('Vedlegg').setFontWeight('bold');
        H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
        c = _headersMap_(H);
      }

      const cell = sh.getRange(rowNum, c.Vedlegg + 1);
      let attachments = [];
      const existingVal = cell.getValue();
      if (existingVal) {
        try {
          attachments = JSON.parse(existingVal);
          if (!Array.isArray(attachments)) attachments = [];
        } catch (e) { /* ignore */ }
      }
      attachments.push(newAttachment);
      cell.setValue(JSON.stringify(attachments));

      safeLog('Vedlegg_LastetOpp', `Vaktmester ${userEmail} lastet opp ${file.getName()} til ${taskId}`);
      return { ok: true, attachment: newAttachment };

    } catch (e) {
      safeLog('VaktmesterAPI_Feil', `uploadAttachment: ${e.message}`);
      throw e;
    } finally {
      lock.releaseLock();
    }
  }

  function createVaktmesterTask(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (!_hasAccess_()) throw new Error('Ingen tilgang til vaktmester-modulen.');
      const user = getCurrentVaktmesterProfile();
      if (!payload || !String(payload.tittel || '').trim()) {
        throw new Error('Tittel er påkrevd.');
      }

      const sh = _ensureTaskSheet_();
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const c = _headersMap_(H);

      const newId = (typeof _nextTaskId_ === 'function') ? _nextTaskId_() : `TASK-${Date.now()}`;
      const row = new Array(H.length).fill('');

      row[c.OppgaveID] = newId;
      row[c.Tittel] = payload.tittel;
      row[c.Beskrivelse] = payload.beskrivelse || '';
      row[c.Seksjonsnr] = payload.seksjonsnr || '';
      row[c.Frist] = payload.frist ? new Date(payload.frist) : '';
      row[c.Opprettet] = new Date();
      row[c.Status] = 'Ny';
      row[c.Prioritet] = payload.prioritet || 'Medium';
      row[c.Ansvarlig] = user.email;
      if (c.Kilde > -1) row[c.Kilde] = 'Vaktmester-UI';
      if (c.Kategori > -1) row[c.Kategori] = 'Vaktmester';

      sh.appendRow(row);
      VM_IDX.put(newId, sh.getLastRow());

      safeLog('Oppgave_Opprettet', `Vaktmester ${user.email} opprettet ${newId}`);
      return { ok: true, id: newId };
    } catch (e) {
      safeLog('VaktmesterAPI_Feil', `createVaktmesterTask: ${e.message}`);
      throw e;
    } finally {
      lock.releaseLock();
    }
  }

  globalThis.getCurrentVaktmesterProfile = getCurrentVaktmesterProfile;
  globalThis.getTasksForVaktmester = getTasksForVaktmester;
  globalThis.updateTaskStatusByVaktmester = updateTaskStatusByVaktmester;
  globalThis.addTaskCommentByVaktmester = addTaskCommentByVaktmester;
  globalThis.uploadAttachmentByVaktmester = uploadAttachmentByVaktmester;
  globalThis.createVaktmesterTask = createVaktmesterTask;
})();