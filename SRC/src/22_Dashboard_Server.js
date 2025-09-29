/* ====================== Dashboard Server (Backend) ======================
 * FILE: 22_Dashboard_Server.gs | VERSION: 1.2.0 | UPDATED: 2025-09-14
 * PURPOSE: KPI-er og sikre åpner-funksjoner for 37_Dashboard.html
 * - dashMetrics(forceFresh?)  → returnerer {ok,user,counts,ts}
 * - dashOpen(key, params?)    → åpner whitelista UI-moduler trygt
 * - Ytelse: CacheService (45s) pr bruker
 * ====================================================================== */

(function () {
  const PROPS = PropertiesService.getScriptProperties();

  function _email_() {
    try {
      const u = Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '';
      return String(u).toLowerCase();
    } catch (_) { return ''; }
  }
  function _sheet_(name) {
    try { return SpreadsheetApp.getActive().getSheetByName(name); } catch (_) { return null; }
  }
  function _readAll_(sheetName) {
    const sh = _sheet_(sheetName);
    if (!sh || sh.getLastRow() < 1) return { H: [], rows: [] };
    const range = sh.getDataRange();
    const values = range.getValues();
    const H = values.shift() || [];
    return { H, rows: values };
  }
  function _idx_(H) {
    const m = {};
    H.forEach((h, i) => { m[String(h)] = i; });
    return m;
  }
  function _isAdmin_() {
    try {
      if (typeof hasPermission === 'function') return !!hasPermission('VIEW_ADMIN_MENU');
      return false;
    } catch (_) { return false; }
  }
  function _cacheKey_(email) {
    return 'DASH:METRICS:v1:' + (email || 'anon');
  }

  /** KPI-er til dashbordet. forceFresh=true hopper over cache. */
  function dashMetrics(forceFresh) {
    const email = _email_();
    const cache = CacheService.getScriptCache();
    const key = _cacheKey_(email);

    if (!forceFresh) {
      const hit = cache.get(key);
      if (hit) try { return JSON.parse(hit); } catch (_) {}
    }

    // --- KPI: Møter (kommende)
    let upcomingMeetings = 0;
    try {
      const { H, rows } = _readAll_(SHEETS.MOTER);
      const I = _idx_(H);
      const iDato = I['dato'], iStatus = I['status'];
      const today = new Date(); today.setHours(0,0,0,0);
      if (iDato >= 0) {
        upcomingMeetings = rows.reduce((acc, r) => {
          const d = r[iDato] instanceof Date ? r[iDato] : (r[iDato] ? new Date(r[iDato]) : null);
          const st = String((r[iStatus] || '')).toLowerCase();
          if (d && d >= today && st !== 'slettet' && st !== 'arkivert') acc++;
          return acc;
        }, 0);
      }
    } catch (_) {}

    // --- KPI: Oppgaver
    let openTasks = 0, myTasks = 0;
    try {
      const { H, rows } = _readAll_(SHEETS.TASKS);
      const I = _idx_(H);
      const iStatus = I['Status'], iAnsvarlig = I['Ansvarlig'];
      const openSet = new Set(['ny','påbegynt','venter','ny ','påbegynt ','venter '].map(s=>s.trim()));
      rows.forEach(r => {
        const st = String(r[iStatus] || '').toLowerCase().trim();
        if (openSet.has(st)) {
          openTasks++;
          if (String(r[iAnsvarlig] || '').toLowerCase() === email) myTasks++;
        }
      });
    } catch (_) {}

    // --- KPI: Godkjenninger (placeholder = 0 inntil kilde er klar)
    const pendingApprovals = 0;

    const result = {
      ok: true,
      ts: new Date().toISOString(),
      user: { email, isAdmin: _isAdmin_() },
      counts: { upcomingMeetings, openTasks, myTasks, pendingApprovals }
    };

    try { cache.put(key, JSON.stringify(result), 45); } catch (_) {}
    return result;
  }

  /** Sikker åpner – whitelist av lovlige nøkler + valgfri RBAC. */
  function dashOpen(key, params) {
    const map = globalThis.UI_FILES || {};
    const entry = map[key];
    if (!entry || !entry.file) throw new Error('Ukjent modulnøkkel: ' + key);

    // RBAC (valgfri – koble nøkler til permissions)
    const permMap = {
      MOTEOVERSIKT: 'OPEN_MEETINGS_UI',
      MOTE_SAK_EDITOR: 'OPEN_MEETING_EDITOR',
      EIERSKIFTE: 'OPEN_OWNERSHIP_FORM',
      PROTOKOLL_GODKJENNING: 'OPEN_PROTOCOL_APPROVAL',
      SEKSJON_HISTORIKK: 'OPEN_SECTION_HISTORY',
      VAKTMESTER: 'VIEW_VAKTMESTER_UI',
      LEVERANDOR_OVERSIKT: 'VIEW_SUPPLIER_REGISTRY',
      BEBOERREGISTER: 'OPEN_BEBOERREGISTER_UI'
    };
    const need = permMap[key];
    if (need && typeof hasPermission === 'function' && !hasPermission(need)) {
      throw new Error('Du har ikke tilgang til denne modulen.');
    }

    const html = HtmlService.createHtmlOutputFromFile(entry.file)
      .setTitle(entry.title || 'Modul')
      .setWidth(entry.w || 1100)
      .setHeight(entry.h || 760)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    const ui = (typeof globalThis._ui === 'function') ? globalThis._ui() : SpreadsheetApp.getUi();
    ui.showModalDialog(html, entry.title || 'Modul');
    return true;
  }

  // Eksporter
  globalThis.dashMetrics = dashMetrics;
  globalThis.dashOpen = dashOpen;
})();
