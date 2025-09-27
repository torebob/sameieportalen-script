/* ====================== Dashboard (Sidebar & Modal) - OPTIMIZED (IIFE) ======================
 * FILE: 05_Dashboard_UI.gs  |  VERSION: 2.4.0  |  UPDATED: 2025-09-14
 *
 * FORMÅL:
 *  - Admin: dashboard i sidepanel (openDashboard).
 *  - Brukere: stort MODAL-dashbord (openDashboardModal / openDashboardAuto).
 *  - Én felles "åpne" mekanisme (dashOpen) for alle moduler via UI_FILES-mapping.
 *  - Enkle nøkkeltall til dashbordet (dashMetrics).
 *
 * AVHENGIGHETER:
 *  - globalThis.SHEETS og globalThis.APP (fra 00_App_Core.gs)
 *
 * KONFIG (ark: SHEETS.KONFIG, kolonner A:B = Nøkkel:Verdi):
 *  - ADMIN_WHITELIST      : Kommadelt liste med admin-e-poster (viser admin-knapper)
 *  - ADMIN_NOTIFY_EMAILS  : Kommadelt liste med e-poster som varsles ved systemfeil-oppgaver
 *
 * MENY (anbefalt i 00_App_Core.gs → onOpen):
 *   if (typeof addDashboardMenu === 'function') addDashboardMenu();
 *   // eller:
 *   ui.createMenu('Sameieportalen')
 *     .addItem('Dashbord', 'openDashboardAuto')        // bruker: modal, admin: sidebar
 *     .addItem('Adminpanel (sidepanel)', 'openDashboard')
 *     .addToUi();
 * ============================================================================================ */

(function(){
  /* ---------- Private constants & cache ---------- */
  const DASHBOARD_CONFIG = {
    CACHE_DURATION: 5 * 60 * 1000, // 5 minutter
    MAX_RETRIES: 3,
    SIDEBAR_WIDTH: 280,
    SIDEBAR_HEIGHT: 650,
    REQUIRED_GLOBALS: ['SHEETS', 'APP']
  };

  // Standard UI-files (kan overstyres av globalThis.UI_FILES)
  const UI_MAP_DEFAULT = {
    // Dashbord (modal)
    DASHBOARD_HTML: { file:'36_Dashboard.html', title:'Sameieportal — Dashbord', w:1280, h:840 },

    // Moduler (bruk dine nummererte filer)
    MOTEOVERSIKT:           { file:'30_Moteoversikt.html',               title:'Møteoversikt & Protokoller', w:1100, h:760 },
    MOTE_SAK_EDITOR:        { file:'31_MoteSakerEditor.html',            title:'Møtesaker – Editor',          w:1100, h:760 },
    EIERSKIFTE:             { file:'32_Eierskifteskjema.html',           title:'Eierskifteskjema',            w:980,  h:760 },
    PROTOKOLL_GODKJENNING:  { file:'33_ProtokollGodkjenningSkjema.html', title:'Protokoll-godkjenning',       w:980,  h:760 },
    SEKSJON_HISTORIKK:      { file:'34_SeksjonHistorikk.html',           title:'Seksjonshistorikk',           w:1100, h:760 },
    VAKTMESTER:             { file:'35_VaktmesterVisning.html',          title:'Vaktmester',                  w:1100, h:800 }
  };

  // Eksponér default hvis ikke satt i 00_App_Core
  if (!globalThis.UI_FILES) globalThis.UI_FILES = UI_MAP_DEFAULT;

  let _dashboardCache = {
    config: null,
    configTime: 0,
    userInfo: null,
    userTime: 0
  };

  /* ---------- Utils ---------- */
  function _validateDependencies_() {
    const missing = DASHBOARD_CONFIG.REQUIRED_GLOBALS.filter(dep => typeof globalThis[dep] === 'undefined');
    if (missing.length > 0) throw new Error(`Missing required dependencies: ${missing.join(', ')}`);
  }
  function _escape_(value) {
    if (value == null) return '';
    return String(value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
  function _withRetry_(operation, maxRetries = DASHBOARD_CONFIG.MAX_RETRIES, delayMs = 100) {
    let lastError;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try { return operation(); }
      catch (error) {
        lastError = error;
        if (attempt < maxRetries) Utilities.sleep(delayMs * attempt);
      }
    }
    throw lastError;
  }
  function _todayMidnight_(){
    const d = new Date(); d.setHours(0,0,0,0); return d;
  }
  function _ui_(){
    return (typeof globalThis._ui === 'function') ? globalThis._ui() : SpreadsheetApp.getUi();
  }

  /*
   * MERK: User/auth og Konfig-hjelpere er fjernet fra denne filen.
   * De er nå definert og håndtert sentralt i Config_Service Dashbord.js,
   * som eksponerer dem til det globale skopet.
   */
  function _isAdminUser_(userInfo) {
    const whitelist = _parseEmailList_(_getConfigValue_('ADMIN_WHITELIST',''));
    return !!(userInfo && userInfo.email && whitelist.includes(userInfo.email));
  }

  /* ---------- HTML/CSS/JS for sidepanel ---------- */
  function _generateCSS_() {
    return `
      <style>
        :root{--bg:#0b1220;--card:#0f172a;--muted:#94a3b8;--border:#1f2937;--primary:#3b82f6;--success:#10b981;--warning:#f59e0b;--danger:#ef4444}
        *{box-sizing:border-box} body{margin:0;font-family:Inter,-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,system-ui,sans-serif;background:var(--bg);color:#fff;line-height:1.5}
        .dashboard-container{display:flex;min-height:100vh}
        .sidebar{width:240px;background:#0b1020;border-right:1px solid var(--border);padding:16px;flex-shrink:0}
        .main-content{flex:1;display:flex;flex-direction:column;min-width:0}
        .header{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid var(--border);background:#0c1526;position:sticky;top:0;z-index:10}
        .title{font-weight:600;font-size:18px;letter-spacing:-.025em}
        .subtitle{font-size:13px;color:var(--muted);margin-top:2px}
        .badge{font-size:12px;color:var(--muted);padding:4px 8px;background:var(--card);border-radius:6px}
        .content-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;padding:20px}
        .card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px;transition:border-color .2s ease}
        .card:hover{border-color:var(--primary)}
        .card-title{margin:0 0 8px;font-size:16px;font-weight:600;color:#e2e8f0}
        .card-content{color:var(--muted);font-size:14px;line-height:1.6}
        .kpi-container{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}
        .kpi-pill{border:1px solid var(--border);background:#0b1b33;padding:6px 12px;border-radius:20px;font-size:12px;color:#cbd5e1;font-weight:500}
        .btn-group{display:flex;gap:8px;flex-wrap:wrap}
        .btn{border:1px solid var(--border);background:var(--card);color:#e2e8f0;padding:8px 16px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;transition:all .2s ease;text-decoration:none;display:inline-flex;align-items:center;gap:6px}
        .btn:hover:not(:disabled){background:var(--primary);border-color:var(--primary);transform:translateY(-1px)}
        .btn:disabled{opacity:.5;cursor:not-allowed}
        .btn.success{border-color:var(--success);background:rgba(16,185,129,.1)}
        .btn.success:hover:not(:disabled){background:var(--success)}
        .warning-banner{margin:16px 20px 0;padding:12px 16px;border-radius:8px;border:1px solid var(--warning);background:rgba(245,158,11,.1);color:#fcd34d;font-size:13px}
        .nav-list{list-style:none;margin:16px 0 0;padding:0}
        .nav-item{padding:12px;border-radius:8px;color:#cbd5e1;cursor:pointer;transition:all .2s ease;font-size:14px}
        .nav-item:hover{background:#101a33;color:#fff}
        .footer{padding:16px 20px;border-top:1px solid var(--border);color:var(--muted);font-size:12px;text-align:center}
        .footer a{color:#93c5fd;text-decoration:none}.footer a:hover{text-decoration:underline}
        .loading{opacity:.7;pointer-events:none}
      </style>
    `;
  }
  function _generateJavaScript_() {
    return `
      <script>
        const Dashboard = {
          init(){ this.setupEventListeners(); this.loadInitialData(); },
          setupEventListeners(){
            document.querySelectorAll('.nav-item').forEach(item=>{
              item.addEventListener('click', e=>this.handleNavigation(e.currentTarget.dataset.tab));
            });
            document.querySelectorAll('.admin-btn').forEach(btn=>{
              btn.addEventListener('click', e=>this.handleAdminAction(e.currentTarget.dataset.action, e.currentTarget));
            });
          },
          handleNavigation(tab){
            google.script.host.setHeight(${DASHBOARD_CONFIG.SIDEBAR_HEIGHT});
            google.script.host.setWidth(${DASHBOARD_CONFIG.SIDEBAR_WIDTH + 400});
          },
          handleAdminAction(action, button){
            if (!action || button.disabled) return;
            const originalText = button.textContent;
            button.disabled = true; button.classList.add('loading'); button.textContent = 'Arbeider...';
            google.script.run
              .withSuccessHandler(msg=>{
                button.disabled=false; button.classList.remove('loading'); button.textContent=originalText;
                this.showNotification(msg||'Utført!', 'success');
              })
              .withFailureHandler(error=>{
                button.disabled=false; button.classList.remove('loading'); button.textContent=originalText;
                this.showNotification('Feil: ' + (error && error.message ? error.message : error), 'error');
                console.error(error);
              })[action]();
          },
          showNotification(message, type='info'){
            const el=document.createElement('div');
            el.textContent=message;
            el.style.cssText='position:fixed;top:20px;right:20px;background:var(--'+(type==='success'?'success':type==='error'?'danger':'primary')+');color:#fff;padding:12px 16px;border-radius:8px;z-index:1000;box-shadow:0 4px 12px rgba(0,0,0,.3)';
            document.body.appendChild(el); setTimeout(()=>el.remove(), 3000);
          },
          loadInitialData(){ /* hook for future */ }
        };
        document.addEventListener('DOMContentLoaded', ()=>Dashboard.init());
      </script>
    `;
  }
  function _generateAdminControls_(isAdmin) {
    if (!isAdmin) return '<span class="badge">Ingen admin-tilgang</span>';
    const actions = [
      { action:'runAllChecks', label:'Kjør systemsjekk', class:'success' },
      { action:'adminEnableDevTools', label:'Aktiver utviklerverktøy' },
      { action:'adminDisableDevTools', label:'Deaktiver utviklerverktøy' },
      { action:'adminLogDummyAction', label:'Demo: Logg admin-hendelse' }
    ];
    return `<div class="btn-group">${
      actions.map(a=>`<button class="btn admin-btn ${a.class||''}" data-action="${a.action}">${a.label}</button>`).join('')
    }</div>`;
  }

  /* ---------- Core: build & show SIDE-PANEL dashboard (admin) ---------- */
  function _openDashboardImpl_() {
    const userInfo = _getCurrentUserInfo_();
    const isAdmin = _isAdminUser_(userInfo);
    const showWarning = userInfo.hasEditAccess && !isAdmin;
    const appName = (globalThis.APP && globalThis.APP.NAME) || 'Sameieportalen';
    const appVersion = (globalThis.APP && globalThis.APP.VERSION) || '2.4.0';

    const warningBanner = showWarning ? `
      <div class="warning-banner">
        <strong>Advarsel:</strong> Du har redigeringstilgang, men er ikke i admin-whitelist.<br>
        Legg til <code>${_escape_(userInfo.email)}</code> i Konfig → ADMIN_WHITELIST.
      </div>` : '';

    const html = `
      <!DOCTYPE html><html lang="no"><head>
        <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
        <title>${_escape_(appName)} Dashboard</title>
        ${_generateCSS_()}
      </head><body>
        <div class="dashboard-container">
          <aside class="sidebar">
            <div><div class="title">${_escape_(appName)}</div><div class="subtitle">v${_escape_(appVersion)}</div></div>
            <nav><ul class="nav-list">
              <li class="nav-item" data-tab="inbox">📥 Krever handling</li>
              <li class="nav-item" data-tab="saker">📋 Saker</li>
              <li class="nav-item" data-tab="hms">⚠️ HMS</li>
              <li class="nav-item" data-tab="dokumenter">📄 Dokumenter</li>
              <li class="nav-item" data-tab="rapporter">📊 Rapporter</li>
            </ul></nav>
          </aside>
          <main class="main-content">
            <header class="header">
              <div><div class="title">Dashboard</div><div class="subtitle">Innlogget: ${_escape_(userInfo.email||'Ukjent')}</div></div>
              ${_generateAdminControls_(isAdmin)}
            </header>
            ${warningBanner}
            <div class="content-grid">
              <div class="card"><h3 class="card-title">Krever handling</h3><div class="card-content">Oppgaver og henvendelser som trenger oppfølging.</div>
                <div class="kpi-container"><div class="kpi-pill">Nye oppgaver: 0</div><div class="kpi-pill">Support: 0</div><div class="kpi-pill">Avvik: 0</div></div>
              </div>
              <div class="card"><h3 class="card-title">Mine oppgaver</h3><div class="card-content">Oppgaver tildelt til deg basert på rolle og tilganger.</div>
                <div class="kpi-container"><div class="kpi-pill">Aktive: 0</div><div class="kpi-pill">Forfalt: 0</div></div>
              </div>
              <div class="card"><h3 class="card-title">HMS og sikkerhet</h3><div class="card-content">Avvik registreres og konverteres til oppgaver innen 60 sekunder.</div>
                <div class="kpi-container"><div class="kpi-pill">Åpne avvik: 0</div><div class="kpi-pill">Varsler: 0</div></div>
              </div>
            </div>
            <footer class="footer">
              <p>Filstruktur Kap. 31 • Responsivt design • WCAG 2.1 AA<br>
                <a href="#" onclick="Dashboard.showNotification('Versjon: ${_escape_(appVersion)}', 'info')">Systeminfo</a>
              </p>
            </footer>
          </main>
        </div>
        ${_generateJavaScript_()}
      </body></html>`;

    const out = HtmlService.createHtmlOutput(html)
      .setTitle(`${appName} – Dashboard`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    const ui = _ui_();
    if (ui) ui.showSidebar(out);
    return out;
  }

  /* ---------- Modal dashboard (bruker) + auto-velger ---------- */
  function openDashboardModal(params) {
    const spec = (globalThis.UI_FILES && globalThis.UI_FILES.DASHBOARD_HTML) || UI_MAP_DEFAULT.DASHBOARD_HTML;
    const t = HtmlService.createTemplateFromFile(spec.file);
    t.PARAMS = params || {};
    const out = t.evaluate()
      .setTitle(spec.title)
      .setWidth(spec.w)
      .setHeight(spec.h)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    _ui_().showModalDialog(out, spec.title);
  }
  function openDashboardAuto() {
    const u = _getCurrentUserInfo_();
    return _isAdminUser_(u) ? openDashboard() : openDashboardModal();
  }

  /* ---------- Felles åpner for alle moduler ---------- */
  function dashOpen(key, params) {
    const spec = (globalThis.UI_FILES && globalThis.UI_FILES[key]) || UI_MAP_DEFAULT[key];
    if (!spec) throw new Error('Ukjent UI-nøkkel: ' + key);
    const t = HtmlService.createTemplateFromFile(spec.file);
    t.PARAMS = params || {};
    const out = t.evaluate()
      .setTitle(spec.title)
      .setWidth(spec.w)
      .setHeight(spec.h)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    _ui_().showModalDialog(out, spec.title);
  }

  /* ---------- Enkle nøkkeltall til dashbordet ---------- */
  function dashMetrics() {
    const counts = { upcomingMeetings: 0, openTasks: 0, myTasks: 0, pendingApprovals: 0 };
    const userEmail = (_getCurrentUserInfo_().email || '').toLowerCase();
    try {
      const ss = SpreadsheetApp.getActive();

      // Møter
      const meetingName = globalThis.SHEETS.MOTER || globalThis.SHEETS.MØTER || globalThis.SHEETS.MEETINGS || 'Møter';
      const shM = ss.getSheetByName(meetingName);
      if (shM && shM.getLastRow() > 1) {
        const data = shM.getDataRange().getValues();
        const H = data.shift();
        const iDato = H.indexOf('dato') > -1 ? H.indexOf('dato') : H.indexOf('Dato');
        const iStatus = H.indexOf('status') > -1 ? H.indexOf('status') : H.indexOf('Status');
        const today = _todayMidnight_();
        counts.upcomingMeetings = data.filter(r=>{
          const d = r[iDato]; const st = (r[iStatus]||'').toString().toLowerCase();
          const okStatus = st !== 'slettet' && st !== 'arkivert';
          const dd = (d instanceof Date) ? d : (d ? new Date(d) : null);
          return okStatus && dd && dd >= today;
        }).length;
      }

      // Oppgaver (Tasks)
      const tasksName = globalThis.SHEETS.TASKS || 'Oppgaver';
      const shT = ss.getSheetByName(tasksName);
      if (shT && shT.getLastRow() > 1) {
        const data = shT.getDataRange().getValues();
        const H = data.shift();
        const iStatus = H.indexOf('Status');
        const iAnsvar = H.indexOf('Ansvarlig');
        const activeSet = new Set(['ny','påbegynt','paabegynt','venter','åpen','apen','open']);
        data.forEach(r=>{
          const st = (r[iStatus]||'').toString().toLowerCase();
          if (activeSet.has(st)) {
            counts.openTasks++;
            const ansv = (r[iAnsvar]||'').toString().toLowerCase();
            if (userEmail && ansv === userEmail) counts.myTasks++;
          }
        });
      }

      // Godkjenninger (best effort – teller "venter"/"avventer")
      const cand = [
        globalThis.SHEETS.PROTOKOLL_GODKJENNINGER,
        globalThis.SHEETS.PROTOKOLL_GODKJENNING,
        'ProtokollGodkjenninger','ProtokollGodkjenning','Godkjenninger'
      ].filter(Boolean);
      let shG = null; for (var i=0;i<cand.length;i++){ shG = ss.getSheetByName(cand[i]); if (shG) break; }
      if (shG && shG.getLastRow() > 1) {
        const data = shG.getDataRange().getValues();
        const H = data.shift();
        const iStatus = H.indexOf('Status');
        counts.pendingApprovals = data.filter(r=>{
          const st = (r[iStatus]||'').toString().toLowerCase();
          return st.includes('venter') || st.includes('avventer') || st.includes('til godkjenning');
        }).length;
      }

      return { ok:true, counts };
    } catch (e) {
      return { ok:false, error:e.message, counts };
    }
  }

  /* ---------- Fallback: oppgave + e-post ---------- */
  function _createSystemTask_(title, details) {
    if (!globalThis.SHEETS || !globalThis.SHEETS.TASKS) throw new Error('SHEETS.TASKS mangler.');
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(globalThis.SHEETS.TASKS);
    // Sørg for HEADERS lik 01_Setup_og_Vedlikehold.gs
    const TASK_HDR = ['OppgaveID','Tittel','Beskrivelse','Kategori','Prioritet','Opprettet','Frist','Status','Ansvarlig','Seksjonsnr','Relatert','Kommentarer'];
    if (!sh) sh = ss.insertSheet(globalThis.SHEETS.TASKS);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,TASK_HDR.length).setValues([TASK_HDR]).setFontWeight('bold'); sh.setFrozenRows(1);
    } else {
      const cur = sh.getRange(1,1,1,TASK_HDR.length).getValues()[0];
      if (JSON.stringify(cur) !== JSON.stringify(TASK_HDR)) {
        sh.getRange(1,1,1,Math.max(TASK_HDR.length, sh.getLastColumn())).clearContent();
        sh.getRange(1,1,1,TASK_HDR.length).setValues([TASK_HDR]).setFontWeight('bold'); sh.setFrozenRows(1);
      }
    }
    // TaskID via Script Properties
    let nextId;
    try {
      const PROPS = PropertiesService.getScriptProperties();
      const raw = PROPS.getProperty('TASK_ID_SEQ');
      const cur = raw ? parseInt(raw,10) : 0;
      nextId = isNaN(cur) ? 1 : cur + 1;
      PROPS.setProperty('TASK_ID_SEQ', String(nextId));
    } catch (_) {
      nextId = sh.getLastRow(); // fallback
    }
    const frist = new Date(); frist.setDate(frist.getDate()+3);
    const row = [
      'TASK-' + Utilities.formatString('%04d', nextId),
      title || 'Systemfeil',
      details || '',
      'System',
      'Høy',
      new Date(),
      frist,
      'Ny',
      '',
      '',
      'Dashboard',
      ''
    ];
    sh.appendRow(row);
    // Formater datoer
    const cOpprettet = TASK_HDR.indexOf('Opprettet') + 1;
    const cFrist = TASK_HDR.indexOf('Frist') + 1;
    if (cOpprettet) sh.getRange(sh.getLastRow(), cOpprettet).setNumberFormat('dd.MM.yyyy "kl." HH:mm');
    if (cFrist) sh.getRange(sh.getLastRow(), cFrist).setNumberFormat('dd.MM.yyyy');
    return row[0]; // OppgaveID
  }
  function _notifyOnSystemTask_(taskId, title, details) {
    const recipients = _parseEmailList_(_getConfigValue_('ADMIN_NOTIFY_EMAILS',''));
    if (!recipients.length) return;
    const subject = `[Sameieportalen] ${title} (${taskId})`;
    const body = `${title}\n\nOppgaveID: ${taskId}\nOpprettet: ${new Date().toLocaleString()}\n\nDetaljer:\n${details}\n\nÅpne regnearket for å se Oppgaver.\n--\n${(globalThis.APP&&globalThis.APP.NAME)||'Sameieportalen'} v${(globalThis.APP&&globalThis.APP.VERSION)||''}`;
    try { recipients.forEach(rcpt => MailApp.sendEmail(rcpt, subject, body)); } catch (_) {}
  }

  /* ---------- Exporterte (globale) funksjoner ---------- */
  globalThis.openDashboard = function openDashboard() {
    try { return _openDashboardImpl_(); }
    catch (error) {
      return HtmlService.createHtmlOutput(`<h2>Dashboard-feil</h2><p>${_escape_(error.message)}</p>`);
    }
  };
  // Sidepanel-alias (bakoverkompatibelt)
  globalThis.openDashboardSidebar = function openDashboardSidebar(){ return globalThis.openDashboard(); };

  // Modal-variant + auto-velger
  globalThis.openDashboardModal = openDashboardModal;
  globalThis.openDashboardAuto  = openDashboardAuto;

  // Felles åpner + nøkkeltall til modal-dashbord
  globalThis.dashOpen    = dashOpen;
  globalThis.dashMetrics = dashMetrics;

  /*
   * MERK: adminEnableDevTools og adminDisableDevTools er fjernet herfra.
   * De er nå definert sentralt i 00_App_Core.js og er tilgjengelige globalt.
   */
  globalThis.adminLogDummyAction = function adminLogDummyAction() {
    const email = (Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '').toLowerCase();
    try {
      if (!globalThis.SHEETS || !globalThis.SHEETS.LOGG) {
        throw new Error('SHEETS.LOGG mangler – kontroller 00_App_Core.gs.');
      }
      const ss = SpreadsheetApp.getActive();
      let sh = ss.getSheetByName(globalThis.SHEETS.LOGG);
      if (!sh) sh = ss.insertSheet(globalThis.SHEETS.LOGG);
      if (sh.getLastRow() === 0) sh.appendRow(['Tid','Type','Bruker','Detaljer']);
      sh.appendRow([new Date(),'ADMIN_DEMO', email || '(ukjent)', 'Demo-hendelse fra Dashboard-knappen']);
      return 'Hendelsen ble logget i Hendelseslogg.';
    } catch (err) {
      const details = 'Forsøk på å logge ADMIN_DEMO feilet.\nFeil: ' + (err && err.message ? err.message : String(err)) +
                      '\nBruker: ' + (email || '(ukjent)') + '\nTid: ' + new Date().toISOString();
      const taskId = _createSystemTask_('Systemfeil: Admin-demo feilet', details);
      _notifyOnSystemTask_(taskId, 'Systemfeil: Admin-demo feilet', details);
      return `Feil ved logging. Oppgave opprettet (${taskId}) og varsling sendt.`;
    }
  };

  // Valgfri: legg meny (kall fra 00_App_Core.gs → onOpen)
  globalThis.addDashboardMenu = function addDashboardMenu() {
    const ui = _ui_();
    if (!ui) return;
    ui.createMenu('Sameieportalen')
      .addItem('Dashbord', 'openDashboardAuto')            // bruker=modal, admin=sidebar
      .addSeparator()
      .addItem('Adminpanel (sidepanel)', 'openDashboard')  // eksplisitt sidepanel
      .addToUi();
  };

  // Cache-verktøy
  globalThis.clearDashboardCache = function clearDashboardCache(){
    _dashboardCache={config:null,configTime:0,userInfo:null,userTime:0};
    return 'Dashboard-cache tømt.';
  };
  globalThis.getDashboardStatus = function getDashboardStatus(){
    const userInfo=_getCurrentUserInfo_(); const config=_getCachedConfig_();
    return {timestamp:new Date().toISOString(),userEmail:userInfo.email,isAdmin:_isAdminUser_(userInfo),
      hasEditAccess:userInfo.hasEditAccess,configKeys:Object.keys(config).length,
      cacheStatus:{configAge:Date.now()-_dashboardCache.configTime,userAge:Date.now()-_dashboardCache.userTime}};
  };
})();
