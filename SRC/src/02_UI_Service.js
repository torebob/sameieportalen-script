/**
 * Sameieportal – UI Service
 * FILE: 02_UI_Service.gs | VERSION: 2.0.0 | UPDATED: 2025-09-23
 *
 * PURPOSE:
 * - Provides a centralized, reusable service for displaying all user interface elements.
 * - Standardizes the opening of modal dialogs and sidebars.
 * - Includes robust error handling and fallback mechanisms.
 *
 * KEY FUNCTIONS:
 * - _openHtmlTemplate_(): A powerful, generic function that creates and displays any UI from a configuration key.
 * - openDashboardAuto(): Intelligently selects the correct dashboard (modal/sidebar) based on user role.
 * - Fallback Functions (e.g., openMeetingsUI): Provides default UI openers to prevent menu errors if a main module is missing.
 *
 * DEPENDENCIES:
 * - 00_Config.gs: Uses APP constants and the getUIConfig() utility.
 * - Logger.gs: For robust, centralized logging.
 * - An RBAC function (e.g., hasPermission) for role-based UI decisions.
 */

const UI_SERVICE_CONFIG = Object.freeze({
  DEFAULT_MODAL_WIDTH: 1000,
  DEFAULT_MODAL_HEIGHT: 720,
});

/**
 * Safely gets the Spreadsheet UI object.
 * @returns {GoogleAppsScript.Base.Ui | null} The UI object, or null if not available.
 */
function _ui() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

/**
 * Displays a standardized alert dialog.
 * @param {string} message The message to display.
 * @param {string} [title=APP.NAME] The title of the dialog.
 */
function _alert_(message, title) {
  const functionName = '_alert_';
  try {
    const ui = _ui();
    if (ui) {
      ui.alert(title || APP.NAME, String(message), ui.ButtonSet.OK);
    } else {
      Logger.warn(functionName, `UI not available. Alert message: ${message}`);
    }
  } catch (e) {
    Logger.error(functionName, 'Failed to show alert.', { errorMessage: e.message, originalMessage: message });
  }
}

/**
 * Displays a standardized toast notification.
 * @param {string} message The message to display.
 */
function _toast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(message));
  } catch (e) {
    Logger.error('_toast_', 'Failed to show toast.', { errorMessage: e.message, originalMessage: message });
  }
}

/**
 * Generic and secure UI opener for all HTML-based interfaces.
 * @param {string} key The key from CONFIG_UI identifying which UI to open.
 * @param {'modal' | 'sidebar'} [target='modal'] The type of UI to display.
 * @param {object} [params={}] An object of parameters to pass to the HTML template.
 * @returns {GoogleAppsScript.HTML.HtmlOutput | undefined} The HtmlOutput object, or undefined on failure.
 */
function _openHtmlTemplate_(key, target = 'modal', params = {}) {
  const functionName = '_openHtmlTemplate_';
  try {
    const ui = _ui();
    if (!ui) {
      Logger.warn(functionName, `UI not available to open key: ${key}`);
      return;
    }

    // Safely get configuration using the utility from 00_Config.gs
    const cfg = getUIConfig(key);
    
    const template = HtmlService.createTemplateFromFile(cfg.file);
    
    // Inject standard variables and custom parameters into the template
    template.APP_INFO = { VERSION: APP.VERSION, BUILD: APP.BUILD };
    template.PARAMS = params || {};

    const htmlOutput = template.evaluate()
      .setTitle(cfg.title || APP.NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    if (target === 'sidebar') {
      ui.showSidebar(htmlOutput);
    } else {
      htmlOutput
        .setWidth(cfg.width || UI_SERVICE_CONFIG.DEFAULT_MODAL_WIDTH)
        .setHeight(cfg.height || UI_SERVICE_CONFIG.DEFAULT_MODAL_HEIGHT);
      ui.showModalDialog(htmlOutput, cfg.title || APP.NAME);
    }
    
    Logger.info(functionName, `Opened UI for key: ${key}`, { file: cfg.file, target });
    return htmlOutput;

  } catch (e) {
    Logger.error(functionName, `Failed to open UI for key: ${key}`, { errorMessage: e.message });
    _alert_(`Kunne ikke åpne grensesnittet for '${key}': ${e.message}`, 'UI Feil');
  }
}

// ========== DASHBOARD OPENERS ==========

/** Opens the main user-facing modal dashboard. */
function openDashboardModal() {
  return _openHtmlTemplate_('DASHBOARD_HTML', 'modal');
}

/** Opens the admin-specific sidebar dashboard. */
function openDashboardSidebar() {
  // This function assumes a global `openDashboard` function exists in an admin-specific module.
  // This provides loose coupling, allowing the admin module to be optional.
  if (typeof globalThis.openDashboard === 'function') {
    return globalThis.openDashboard();
  }
  // Fallback to the user modal if the admin sidebar function isn't found.
  Logger.warn('openDashboardSidebar', 'Admin function "openDashboard" not found. Falling back to modal.');
  return openDashboardModal();
}

/**
 * Intelligently opens the correct dashboard (modal for users, sidebar for admins).
 */
function openDashboardAuto() {
  // Assumes a global hasPermission function exists for RBAC
  if (typeof hasPermission === 'function') {
    if (!hasPermission('VIEW_USER_DASHBOARD')) {
      return _alert_('Du har ikke tilgang til å se dashbordet.', 'Tilgang Nektet');
    }
    const isAdmin = hasPermission('VIEW_ADMIN_MENU');
    return isAdmin ? openDashboardSidebar() : openDashboardModal();
  }
  
  // Fallback if no permission system is found
  return openDashboardModal();
}


// ========== FALLBACK UI OPENERS ==========
// These functions provide default implementations for menu items.
// This ensures the menu doesn't break if the main script for a feature is missing.

if (typeof globalThis.openMeetingsUI !== 'function') {
  globalThis.openMeetingsUI = function() { return _openHtmlTemplate_('MOTEOVERSIKT', 'modal'); };
}
if (typeof globalThis.openMoteSakEditor !== 'function') {
  globalThis.openMoteSakEditor = function() { return _openHtmlTemplate_('MOTE_SAK_EDITOR', 'modal'); };
}
if (typeof globalThis.openOwnershipForm !== 'function') {
  globalThis.openOwnershipForm = function() { return _openHtmlTemplate_('EIERSKIFTE', 'modal'); };
}
if (typeof globalThis.openSectionHistory !== 'function') {
  globalThis.openSectionHistory = function() { return _openHtmlTemplate_('SEKSJON_HISTORIKK', 'modal'); };
}
if (typeof globalThis.openVaktmesterUI !== 'function') {
  globalThis.openVaktmesterUI = function() { return _openHtmlTemplate_('VAKTMESTER', 'modal'); };
}