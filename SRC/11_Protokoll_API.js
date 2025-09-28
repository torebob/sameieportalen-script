/* ================== Elektronisk Protokollgodkjenning (API) ==================
 * FILE: 11_Protokoll_API.gs | VERSION: 3.0.0 | UPDATED: 2025-09-26
 * FORMÅL: Sende, spore og motta godkjenninger/avvisninger for møteprotokoller.
 * ENDRINGER v3.0.0:
 *  - Modernisert til let/const og arrow functions.
 *  - Bruker sentrale hjelpefunksjoner fra 000_Utils.js.
 *  - Forbedret kodestruktur og lesbarhet.
 * ========================================================================== */

const _getHeadersFromConfig_ = (key, fallback) => {
  if (typeof globalThis.getConfigValue !== 'function' || typeof globalThis.parseCsvString !== 'function') {
    return fallback;
  }
  const fromConfig = globalThis.getConfigValue(key);
  if (fromConfig) {
    const parsed = globalThis.parseCsvString(fromConfig);
    if (parsed.length > 0) return parsed;
  }
  return fallback;
};

const _PG_HEADERS_ = _getHeadersFromConfig_('HEADERS_PROTOKOLL_GODKJENNING',
  ['Godkjenning-ID', 'Møte-ID', 'Navn', 'E-post', 'Token', 'Utsendt-Dato', 'Status', 'Svar-Dato', 'Kommentar', 'Protokoll-URL']
);

const _MOTE_HEADERS_FALLBACK_ = _getHeadersFromConfig_('HEADERS_MOTER',
  ['Møte-ID', 'Type', 'Dato', 'Starttid', 'Sluttid', 'Sted', 'Tittel', 'Agenda', 'Protokoll-URL', 'Deltakere', 'Kalender-ID', 'Status']
);

// Private helper functions, prefixed with _ to indicate local scope.
const _hdrIdxMap_Protokoll_ = (headers, names) => names.reduce((acc, name) => {
    acc[name] = headers.indexOf(name);
    return acc;
  }, {});

const _uuid8_Protokoll_ = () => Utilities.getUuid().replace(/-/g, '').slice(0, 8);

const _getBoardList_Protokoll_ = () => {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.BOARD);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return vals.map(([navn, mail]) => ({ navn: String(navn || '').trim(), email: String(mail || '').trim() })).filter(p => p.navn && p.email);
};

const _findMoteRow_Protokoll_ = (moteId) => {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.MOTER);
  if (!sh || sh.getLastRow() < 2) return null;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H = _hdrIdxMap_Protokoll_(headers, _MOTE_HEADERS_FALLBACK_);
  if (H['Møte-ID'] === -1) return null;

  const idColRange = sh.getRange(2, H['Møte-ID'] + 1, sh.getLastRow() - 1, 1);
  const finder = idColRange.createTextFinder(String(moteId)).matchEntireCell(true);
  const hit = finder.findNext();
  if (!hit) return null;

  return { sheet: sh, row: hit.getRow(), headers, H };
};

const _ensureSheet_Protokoll_ = (name, headers) => {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
  } else {
      const curHeaders = sh.getRange(1, 1, 1, headers.length).getValues()[0];
      const headersMatch = JSON.stringify(curHeaders) === JSON.stringify(headers);
      if (sh.getLastRow() === 0 || !headersMatch) {
          sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).clearContent();
          sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
          if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
      }
  }
  return sh;
};

function sendProtokollForGodkjenning(moteId, protokollUrl) {
  try {
    if (typeof hasPermission === 'function' && !hasPermission('VIEW_ADMIN_MENU')) {
      throw new Error('Tilgang nektet. (Krever VIEW_ADMIN_MENU)');
    }

    if (!moteId) throw new Error('Mangler Møte-ID.');
    const mote = _findMoteRow_Protokoll_(moteId);
    if (!mote) throw new Error(`Fant ikke møtet i "${SHEETS.MOTER}".`);

    let url = String(protokollUrl || '').trim();
    if (!url) {
      const cProt = mote.headers.indexOf('Protokoll-URL');
      if (cProt === -1) throw new Error('Møter-arket mangler kolonnen "Protokoll-URL".');
      url = String(mote.sheet.getRange(mote.row, cProt + 1).getValue() || '').trim();
    }
    if (!/^https:\/\/docs\.google\.com\/document\//.test(url)) {
      throw new Error('Ugyldig protokoll-URL (må være Google Docs).');
    }

    const godkjSh = _ensureSheet_Protokoll_(SHEETS.PROTOKOLL_GODKJENNING, _PG_HEADERS_);
    const board = _getBoardList_Protokoll_();
    if (board.length === 0) throw new Error(`Fant ingen styremedlemmer i "${SHEETS.BOARD}".`);

    const gid = `G-${String(moteId).replace(/[^A-Za-z0-9_-]/g, '')}-${_uuid8_Protokoll_()}`;
    const now = new Date();

    const h = godkjSh.getRange(1, 1, 1, godkjSh.getLastColumn()).getValues()[0];
    const H = _hdrIdxMap_Protokoll_(h, _PG_HEADERS_);
    const rows = board.map(member => {
      const token = `${_uuid8_Protokoll_()}${_uuid8_Protokoll_()}`;
      const r = new Array(h.length).fill('');
      r[H['Godkjenning-ID']] = gid;
      r[H['Møte-ID']] = moteId;
      r[H['Navn']] = member.navn;
      r[H['E-post']] = member.email;
      r[H['Token']] = token;
      r[H['Utsendt-Dato']] = now;
      r[H['Status']] = 'Sendt';
      r[H['Protokoll-URL']] = url;
      return r;
    });

    if (rows.length) {
      godkjSh.getRange(godkjSh.getLastRow() + 1, 1, rows.length, h.length).setValues(rows);
    }

    if (mote.H['Status'] !== -1) {
      mote.sheet.getRange(mote.row, mote.H['Status'] + 1).setValue('Til godkjenning');
    }

    const webAppUrl = ScriptApp.getService().getUrl();

    rows.forEach((row, j) => {
      const tokenCell = row[H['Token']];
      const email = board[j].email;
      const approveUrl = webAppUrl ? `${webAppUrl}?page=protokoll&gid=${encodeURIComponent(gid)}&token=${encodeURIComponent(tokenCell)}&action=approve` : url;
      const rejectUrl = webAppUrl ? `${webAppUrl}?page=protokoll&gid=${encodeURIComponent(gid)}&token=${encodeURIComponent(tokenCell)}&action=reject` : url;

      const subject = `[Sameieportalen] Til godkjenning: Protokoll ${moteId}`;
      const body = `<p>Hei ${escapeHtml(board[j].navn)},</p><p>Protokollen for møtet <b>${escapeHtml(moteId)}</b> er klar for godkjenning.</p><p><a href="${url}" target="_blank" rel="noopener"><b>Les protokollen her</b></a></p>${webAppUrl ? `<p>Registrer ditt valg:</p><div style="margin:12px 0"><a href="${approveUrl}" style="background:#16a34a;color:#fff;padding:10px 14px;border-radius:6px;text-decoration:none;margin-right:8px">Godkjenn</a><a href="${rejectUrl}" style="background:#dc2626;color:#fff;padding:10px 14px;border-radius:6px;text-decoration:none">Avvis</a></div>` : '<p>(WebApp-URL mangler, kontakt administrator.)</p>'}<p>— Sameieportalen</p>`;

      try {
        MailApp.sendEmail({ to: email, subject, htmlBody: body });
      } catch (mailErr) {
        safeLog('Protokoll_MailFeil', `E-post til ${email} feilet: ${mailErr.message}`);
      }
    });

    safeLog('Protokoll', `Sendte godkjenning ${gid} for Møte-ID ${moteId} til ${board.length} mottakere.`);
    return { ok: true, message: `Protokoll sendt til ${board.length} styremedlemmer.`, gid, count: board.length };
  } catch (e) {
    safeLog('Protokoll_Feil', `sendProtokollForGodkjenning: ${e.message}`);
    throw e;
  }
}

function getProtocolPreviewUrl(moteId) {
  try {
    if (!moteId) throw new Error('Mangler Møte-ID.');

    const mote = _findMoteRow_Protokoll_(moteId);
    if (!mote) throw new Error(`Fant ikke møtet med ID "${moteId}".`);

    const protokollUrlCol = mote.H['Protokoll-URL'];
    if (protokollUrlCol === -1) {
      throw new Error('Fant ikke kolonnen "Protokoll-URL" i møte-arket.');
    }

    const url = mote.sheet.getRange(mote.row, protokollUrlCol + 1).getValue();
    if (!url) {
      throw new Error('Protokoll-URL er ikke registrert for dette møtet.');
    }

    return { url: String(url).trim() };
  } catch (e) {
    safeLog('Protokoll_Preview_Feil', `getProtocolPreviewUrl: ${e.message}`);
    throw new Error(e.message);
  }
}

function handleProtokollApprovalRequest(e) {
  const title = 'Protokoll';
  const page = msg => HtmlService.createHtmlOutput(`<h3>${title}</h3><p>${msg}</p>`);

  try {
    const p = e?.parameter || {};
    const gid = String(p.gid || '').trim();
    const token = String(p.token || '').trim();
    const action = String(p.action || '').trim().toLowerCase();
    const comment = String(p.comment || '').trim();

    if (!gid || !token || !['approve', 'reject'].includes(action)) {
      return page('Lenken mangler nødvendig informasjon.').setTitle('Ugyldig forespørsel');
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEETS.PROTOKOLL_GODKJENNING);
    if (!sh || sh.getLastRow() < 2) return page('Godkjenningsarket mangler.').setTitle('Feil');

    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const H = _hdrIdxMap_Protokoll_(h, _PG_HEADERS_);
    if (['Godkjenning-ID', 'Token', 'Status', 'Svar-Dato', 'Kommentar', 'Møte-ID'].some(key => H[key] === -1)) {
      return page('Godkjenningsarket har ikke forventede kolonner.').setTitle('Feil');
    }

    const tokenColRange = sh.getRange(2, H.Token + 1, sh.getLastRow() - 1, 1);
    const finder = tokenColRange.createTextFinder(token).matchEntireCell(true);
    const hit = finder.findNext();
    if (!hit) return page('Fant ikke denne godkjenningsforespørselen (token).').setTitle('Ugyldig lenke');

    const rowIdx = hit.getRow();
    const rowVals = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    if (String(rowVals[H['Godkjenning-ID']] || '').trim() !== gid) {
      return page('Godkjennings-ID passer ikke.').setTitle('Ugyldig lenke');
    }

    const prevStatus = String(rowVals[H.Status] || '').trim();
    if (['Godkjent', 'Avvist'].includes(prevStatus)) {
      return page(`Ditt svar er allerede registrert: <b>${prevStatus}</b>.`).setTitle('Allerede behandlet');
    }

    const newStatus = (action === 'approve') ? 'Godkjent' : 'Avvist';
    sh.getRange(rowIdx, H.Status + 1).setValue(newStatus);
    sh.getRange(rowIdx, H['Svar-Dato'] + 1).setValue(new Date());
    if (comment) sh.getRange(rowIdx, H.Kommentar + 1).setValue(comment);

    const dataRange = sh.getDataRange().getValues();
    dataRange.shift();
    const rowsForGid = dataRange.filter(r => String(r[H['Godkjenning-ID']] || '').trim() === gid);
    const anyRejected = rowsForGid.some(r => String(r[H.Status] || '').trim() === 'Avvist');
    const allApproved = rowsForGid.length > 0 && rowsForGid.every(r => String(r[H.Status] || '').trim() === 'Godkjent');

    const moteId = String(rowVals[H['Møte-ID']] || '').trim();
    const mote = moteId ? _findMoteRow_Protokoll_(moteId) : null;
    if (mote && mote.H.Status !== -1) {
      let moteStatus = 'Til godkjenning';
      if (anyRejected) moteStatus = 'Avvist';
      else if (allApproved) moteStatus = 'Godkjent';
      mote.sheet.getRange(mote.row, mote.H.Status + 1).setValue(moteStatus);
    }

    safeLog('Protokoll', `Mottok ${newStatus} for ${gid} (Møte ${moteId || '?'}).`);

    const finalTitle = (newStatus === 'Godkjent') ? 'Takk for godkjenningen!' : 'Avvisning registrert';
    const msg = (newStatus === 'Godkjent') ? 'Din godkjenning er registrert.' : 'Din avvisning er registrert. Referent/styret blir varslet ved behov.';
    return page(msg).setTitle(finalTitle);

  } catch (err) {
    safeLog('Protokoll_WebApp_Feil', err.message);
    return page(escapeHtml(err.message)).setTitle('En feil oppstod');
  }
}