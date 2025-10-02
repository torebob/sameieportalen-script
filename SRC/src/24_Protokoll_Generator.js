/* ================== Protokoll-Generator (API) ==================
 * FILE: 24_Protokoll_Generator.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Generere en Google Docs-protokoll basert på møtedata,
 *         saksliste, innspill og stemmeresultater.
 * ================================================================== */

(function (global) {
  // Antar at disse er definert globalt, f.eks. i 00_App_Core.js
  const SHEETS = {
    MOTER: "Møter",
    MOTE_SAKER: "Møtesaker",
    MOTE_KOMMENTARER: "Møtekommentarer",
    MOTE_STEMMER: "Møtestemmer"
  };

  /* ---------- Hjelpere ---------- */
  function _log_(topic, msg) {
    try {
      if (typeof _logEvent === 'function') _logEvent(topic, msg);
      else Logger.log(`[${topic}] ${msg}`);
    } catch (_) {}
  }

  function _findMoteData_(moteId) {
    if (typeof listMeetings_ !== 'function') {
      throw new Error("listMeetings_ er ikke tilgjengelig.");
    }
    const allMeetings = listMeetings_({ scope: 'all' });
    const meetingData = allMeetings.find(m => m.id === moteId);
    return meetingData;
  }

  function _updateMeetingInSheet_(moteId, protocolUrl) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEETS.MOTER);
    if (!sh) throw new Error(`Finner ikke arket: ${SHEETS.MOTER}`);

    const data = sh.getDataRange().getValues();
    const headers = data.shift();
    const idCol = headers.indexOf('id');
    const statusCol = headers.indexOf('status');
    const urlCol = headers.indexOf('Protokoll-URL'); // Endret til å matche 11_Protokoll_API.js

    if (idCol === -1 || statusCol === -1 || urlCol === -1) {
      throw new Error(`Mangler påkrevde kolonner (id, status, Protokoll-URL) i ${SHEETS.MOTER}-arket.`);
    }

    for (let i = 0; i < data.length; i++) {
      if (data[i][idCol] === moteId) {
        const rowNum = i + 2; // +1 for header, +1 for 0-indeks
        sh.getRange(rowNum, statusCol + 1).setValue('Til godkjenning');
        sh.getRange(rowNum, urlCol + 1).setValue(protocolUrl);
        return true;
      }
    }
    return false;
  }


  /**
   * Genererer en protokoll for et gitt møte.
   * @param {string} moteId - ID for møtet.
   * @returns {{ok: boolean, message: string, url?: string}}
   */
  function generateProtocol(moteId) {
    if (!moteId) {
      return { ok: false, message: "Møte-ID er påkrevd." };
    }
    try {
      // 1. Hent møtedetaljer
      const meeting = _findMoteData_(moteId);
      if (!meeting) {
        return { ok: false, message: `Fant ikke møte med ID: ${moteId}` };
      }

      // 2. Hent saksliste
      const saker = listAgenda(moteId);

      // 3. Opprett Google Doc
      const docName = `Protokoll - ${meeting.tittel} - ${new Date(meeting.dato).toLocaleDateString('no-NO')}`;
      const doc = DocumentApp.create(docName);
      const body = doc.getBody();

      // 4. Fyll inn møteinfo
      body.appendParagraph(docName).setHeading(DocumentApp.Attribute.HEADING1);
      body.appendParagraph(`Dato: ${new Date(meeting.dato).toLocaleDateString('no-NO')}`);
      body.appendParagraph(`Tid: ${meeting.start} - ${meeting.slutt || ''}`);
      body.appendParagraph(`Sted: ${meeting.sted}`);
      body.appendHorizontalRule();

      // 5. Loop gjennom saker og fyll inn data
      saker.forEach(sak => {
        body.appendParagraph(`${sak.saksnr}: ${sak.tittel}`).setHeading(DocumentApp.Attribute.HEADING2);

        // Forslag til vedtak
        if (sak.forslag) {
          body.appendParagraph("Forslag til vedtak:").setBold(true);
          body.appendParagraph(sak.forslag);
        }

        // Innspill (notater)
        const innspill = listInnspill(sak.sakId);
        if (innspill && innspill.length > 0) {
          body.appendParagraph("Notater/Innspill:").setBold(true);
          const list = body.appendList();
          innspill.forEach(item => {
            const text = `[${new Date(item.ts).toLocaleString('no-NO')}] ${item.from}: ${item.text}`;
            list.addItem(text);
          });
        }

        // Avstemmingsresultater
        const votes = getVoteSummary(sak.sakId);
        body.appendParagraph("Avstemming:").setBold(true);
        body.appendParagraph(`For: ${votes.JA || 0}, Mot: ${votes.NEI || 0}, Blanke: ${votes.BLANK || 0}`);

        // Endelig vedtak
        body.appendParagraph("Vedtak:").setBold(true);
        body.appendParagraph(sak.vedtak || "Ikke spesifisert.");

        body.appendHorizontalRule();
      });

      doc.saveAndClose();
      const url = doc.getUrl();

      // 6. Oppdater møtestatus og protokoll-URL i arket
      const updated = _updateMeetingInSheet_(moteId, url);
      if (!updated) {
        _log_('ProtokollGenerator', `Klarte ikke oppdatere status for møte ${moteId} i arket.`);
      }

      _log_('ProtokollGenerator', `Protokoll generert for ${moteId}: ${url}`);
      return { ok: true, message: "Protokoll generert!", url: url };

    } catch (e) {
      _log_('ProtokollGenerator_FEIL', e.message);
      return { ok: false, message: `En feil oppstod: ${e.message}` };
    }
  }

  // Eksporter funksjonen for global tilgang
  // global.generateProtocol = generateProtocol;

})(this);