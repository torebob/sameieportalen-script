/* ====================== Dashboard (Sidebar & Modal) - MODERNIZED ======================
 * FILE: 05_Dashboard_UI.gs  |  VERSION: 3.0.0  |  UPDATED: 2025-09-26
 *
 * FORMÅL:
 *  - Sentralisert logikk for å bygge og vise dashboard-UI.
 *  - Bruker moderne JavaScript (let/const, arrow functions) og sentraliserte hjelpere.
 *
 * ENDRINGER v3.0.0:
 *  - Modernisert til let/const og arrow functions.
 *  - Fjernet lokale hjelpefunksjoner; bruker nå 000_Utils.js og Config_Service Dashbord.js.
 *  - Forbedret kodestruktur for lesbarhet og vedlikehold.
 * ============================================================================================ */

(() => {
  const DASHBOARD_CONFIG = {
    CACHE_DURATION: 5 * 60 * 1000,
    MAX_RETRIES: 3,
    SIDEBAR_WIDTH: 280,
    SIDEBAR_HEIGHT: 650,
    REQUIRED_GLOBALS: ['SHEETS', 'APP'],
  };

  const UI_MAP_DEFAULT = {
    DASHBOARD_HTML: { file: '37_Dashboard.html', title: 'Sameieportal — Dashbord', w: 1280, h: 840 },
    MOTEOVERSIKT: { file: '30_Moteoversikt.html', title: 'Møteoversikt & Protokoller', w: 1100, h: 760 },
    MOTE_SAK_EDITOR: { file: '31_MoteSakEditor.html', title: 'Møtesaker – Editor', w: 1100, h: 760 },
    EIERSKIFTE: { file: '34_EierskifteSkjema.html', title: 'Eierskifteskjema', w: 980, h: 760 },
    PROTOKOLL_GODKJENNING: { file: '35_ProtokollGodkjenningSkjema.html', title: 'Protokoll-godkjenning', w: 980, h: 760 },
    SEKSJON_HISTORIKK: { file: '32_SeksjonHistorikk.html', title: 'Seksjonshistorikk', w: 1100, h: 760 },
    VAKTMESTER: { file: '33_VaktmesterVisning.html', title: 'Vaktmester', w: 1100, h: 800 },
  };

  if (!globalThis.UI_FILES) {
    globalThis.UI_FILES = UI_MAP_DEFAULT;
  }

  const _validateDependencies_ = () => {
    const missing = DASHBOARD_CONFIG.REQUIRED_GLOBALS.filter(dep => typeof globalThis[dep] === 'undefined');
    if (missing.length > 0) throw new Error(`Missing required dependencies: ${missing.join(', ')}`);
  };

  const _escape_ = (value) => {
    if (value == null) return '';
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  };

  const _withRetry_ = (operation, maxRetries = DASHBOARD_CONFIG.MAX_RETRIES, delayMs = 100) => {
    let lastError;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return operation();
      } catch (error) {
        lastError = error;
        if (attempt < maxRetries) Utilities.sleep(delayMs * attempt);
      }
    }
    throw lastError;
  };

  let _dashboardCache = { userInfo: null, userTime: 0 };

  const getCurrentUserInfo = () => {
    const now = Date.now();
    if (_dashboardCache.userInfo && (now - _dashboardCache.userTime < DASHBOARD_CONFIG.CACHE_DURATION)) {
      return _dashboardCache.userInfo;
    }

    const user = {
      email: Session.getActiveUser().getEmail(),
      isDev: false,
      permissions: {},
    };

    try {
      const devStatus = globalThis.getUserDevStatus?.();
      if (devStatus) {
        user.isDev = devStatus.isDev;
        user.permissions = devStatus.permissions || {};
      }
    } catch(e) {
      // Ignore if dev status is not available
    }

    _dashboardCache.userInfo = user;
    _dashboardCache.userTime = now;
    return user;
  };

  const _getConfigValue_ = (key, defaultValue = '') => {
    const config = globalThis.getCachedConfig?.() || {};
    return config[key] || defaultValue;
  };

  const _parseEmailList_ = (listString) => {
    if (!listString || typeof listString !== 'string') return [];
    return listString.split(',').map(e => e.trim().toLowerCase()).filter(e => e);
  };

  const _isAdminUser_ = (userInfo) => {
    const whitelist = _parseEmailList_(_getConfigValue_('ADMIN_WHITELIST', ''));
    return !!userInfo?.email && whitelist.includes(userInfo.email);
  };

  const _generateCSS_ = () => `
    <style>
      :root{--bg:#0b1220;--card:#0f172a;--muted:#94a3b8;--border:#1f2937;--primary:#3b82f6;--success:#10b981;--warning:#f59e0b;--danger:#ef4444}
      *{box-sizing:border-box} body{margin:0;font-family:Inter,-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,system-ui,sans-serif;background:var(--bg);color:#fff;line-height:1.5}
      .btn{border:1px solid var(--border);background:var(--card);color:#e2e8f0;padding:8px 16px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;transition:all .2s ease;text-decoration:none;display:inline-flex;align-items:center;gap:6px}
      .btn:hover:not(:disabled){background:var(--primary);border-color:var(--primary);transform:translateY(-1px)}
      .btn:disabled{opacity:.5;cursor:not-allowed}
    </style>`;

  const _generateJavaScript_ = () => `
    <script>
      const Dashboard = {
        init() { this.setupEventListeners(); },
        setupEventListeners() {
          document.querySelectorAll('.admin-btn').forEach(btn => {
            btn.addEventListener('click', e => this.handleAdminAction(e.currentTarget.dataset.action, e.currentTarget));
          });
        },
        handleAdminAction(action, button) {
          if (!action || button.disabled) return;
          const originalText = button.textContent;
          button.disabled = true;
          button.textContent = 'Arbeider...';
          google.script.run
            .withSuccessHandler(msg => {
              button.disabled = false;
              button.textContent = originalText;
              this.showNotification(msg || 'Utført!', 'success');
            })
            .withFailureHandler(error => {
              button.disabled = false;
              button.textContent = originalText;
              this.showNotification('Feil: ' + (error?.message || error), 'error');
              console.error(error);
            })[action]();
        },
        showNotification(message, type = 'info') {
          const el = document.createElement('div');
          el.textContent = message;
          el.style.cssText = 'position:fixed;top:20px;right:20px;background:var(--' + (type === 'success' ? 'success' : type === 'error' ? 'danger' : 'primary') + ');color:#fff;padding:12px 16px;border-radius:8px;z-index:1000;box-shadow:0 4px 12px rgba(0,0,0,.3)';
          document.body.appendChild(el);
          setTimeout(() => el.remove(), 3000);
        }
      };
      document.addEventListener('DOMContentLoaded', () => Dashboard.init());
    </script>`;

  const _generateAdminControls_ = (isAdmin) => {
    if (!isAdmin) return '<span class="badge">Ingen admin-tilgang</span>';
    const actions = [
      { action: 'runAllChecks', label: 'Kjør systemsjekk', class: 'success' },
      { action: 'adminEnableDevTools', label: 'Aktiver utviklerverktøy' },
      { action: 'adminDisableDevTools', label: 'Deaktiver utviklerverktøy' },
      { action: 'adminLogDummyAction', label: 'Demo: Logg admin-hendelse' }
    ];
    return `<div class="btn-group">${actions.map(a => `<button class="btn admin-btn ${a.class || ''}" data-action="${a.action}">${a.label}</button>`).join('')}</div>`;
  };

  const _openDashboardImpl_ = () => {
    const userInfo = getCurrentUserInfo();
    const isAdmin = _isAdminUser_(userInfo);
    const appName = globalThis.APP?.NAME || 'Sameieportalen';
    const appVersion = globalThis.APP?.VERSION || '3.0.0';

    const html = `
      <!DOCTYPE html><html lang="no"><head>
        <meta charset="utf-8"><title>${_escape_(appName)} Dashboard</title>
        ${_generateCSS_()}
      </head><body>
        <header><h1>Dashboard</h1></header>
        <main>
          <p>Innlogget som: ${_escape_(userInfo.email || 'Ukjent')}</p>
          ${_generateAdminControls_(isAdmin)}
        </main>
        ${_generateJavaScript_()}
      </body></html>`;

    const output = HtmlService.createHtmlOutput(html).setTitle(`${appName} – Dashboard`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    _ui()?.showSidebar(output);
    return output;
  };

  const dashMetrics = () => {
    const counts = { upcomingMeetings: 0, openTasks: 0, myTasks: 0, pendingApprovals: 0 };
    const userEmail = getCurrentUserInfo().email || '';
    try {
      const ss = SpreadsheetApp.getActive();
      const meetingName = globalThis.SHEETS.MOTER || 'Møter';
      const shM = ss.getSheetByName(meetingName);
      if (shM && shM.getLastRow() > 1) {
        const data = shM.getDataRange().getValues();
        const H = data.shift();
        const iDato = H.indexOf('Dato');
        const iStatus = H.indexOf('Status');
        const today = _midnight_(new Date());
        counts.upcomingMeetings = data.filter(r => {
          const d = r[iDato];
          const st = (r[iStatus] || '').toString().toLowerCase();
          const okStatus = !['slettet', 'arkivert'].includes(st);
          const dd = (d instanceof Date) ? d : (d ? new Date(d) : null);
          return okStatus && dd && dd >= today;
        }).length;
      }
      return { ok: true, counts };
    } catch (e) {
      return { ok: false, error: e.message, counts };
    }
  };

  globalThis.openDashboard = () => {
    try {
      return _openDashboardImpl_();
    } catch (error) {
      return HtmlService.createHtmlOutput(`<h2>Dashboard-feil</h2><p>${_escape_(error.message)}</p>`);
    }
  };

  globalThis.dashMetrics = dashMetrics;

  globalThis.clearDashboardCache = () => {
    _dashboardCache = { config: null, configTime: 0, userInfo: null, userTime: 0 };
    return 'Dashboard-cache tømt.';
  };
})();