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
    AI_ASSISTENT: { file: '40_AI_Assistent.html', title: 'AI Assistent', w: 800, h: 700 },
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

  const _openDashboardImpl_ = () => {
    const userInfo = getCurrentUserInfo();
    const isAdmin = _isAdminUser_(userInfo);
    const appName = globalThis.APP?.NAME || 'Sameieportalen';

    // Fetch and parse admin buttons from config, with a hardcoded fallback.
    const adminButtonsConfig = _getConfigValue_(
      'DASHBOARD_ADMIN_BUTTONS',
      'Kjør systemsjekk|runAllChecks|success,Aktiver utviklerverktøy|adminEnableDevTools,Deaktiver utviklerverktøy|adminDisableDevTools,Sett AI Mappe-ID|setAiAssistantFolderId'
    );

    const adminButtons = adminButtonsConfig.split(',').map(s => {
        const parts = s.split('|');
        if (parts.length < 2) return null;
        return {
            label: parts[0].trim(),
            action: parts[1].trim(),
            class: parts[2] ? parts[2].trim() : ''
        };
    }).filter(Boolean);

    const template = HtmlService.createTemplateFromFile('SRC/37_Dashboard.html');
    template.userInfo = userInfo;
    template.isAdmin = isAdmin;
    template.appName = appName;
    template.adminButtons = adminButtons;

    const output = template.evaluate()
      .setTitle(`${appName} – Dashboard`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

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

  globalThis.setAiAssistantFolderId = () => {
    const userInfo = getCurrentUserInfo();
    if (!_isAdminUser_(userInfo)) {
      throw new Error("Tilgang nektet. Kun administratorer kan utføre denne handlingen.");
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Konfigurer AI-assistent',
      'Vennligst skriv inn ID-en til Google Drive-mappen som inneholder dokumentene for AI-assistenten:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
      const folderId = response.getResponseText().trim();
      if (folderId) {
        PropertiesService.getScriptProperties().setProperty('AI_ASSISTANT_FOLDER_ID', folderId);
        return `Mappe-ID er satt til: ${folderId}`;
      } else {
        return "Ingen ID ble angitt. Handlingen ble avbrutt.";
      }
    }
    return "Handlingen ble avbrutt av brukeren.";
  };

  globalThis.handleAiAssistantRequest = (e) => {
    try {
      const uiInfo = UI_FILES.AI_ASSISTENT;
      if (!uiInfo) throw new Error("AI Assistant UI configuration not found.");

      const template = HtmlService.createTemplateFromFile(`SRC/${uiInfo.file}`);
      // The AI assistant page doesn't need any template variables for now.

      const output = template.evaluate()
        .setTitle(uiInfo.title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      return output;

    } catch (error) {
      Logger.log(`handleAiAssistantRequest Error: ${error.message}`);
      return HtmlService.createHtmlOutput(`<h2>Feil ved lasting av AI Assistent</h2><p>Kunne ikke laste siden. Feil: ${error.message}</p>`);
    }
  };
})();