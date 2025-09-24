/**
 * Config & Auth Service (v1.0.0)
 * Handles loading configuration from the KONFIG sheet and managing user session data.
 * Provides cached, reusable functions for the rest of the application.
 */

const CONFIG_CACHE_DURATION_MS = 5 * 60 * 1000; // 5 minutes

const _configCache = {
  config: null,
  configTime: 0,
  userInfo: null,
  userTime: 0
};

/**
 * Gets the current user's info (email, admin status). Caches result.
 * @returns {{email: string, isAdmin: boolean, hasEditAccess: boolean}}
 */
function getCurrentUserInfo() {
  const now = Date.now();
  if (_configCache.userInfo && (now - _configCache.userTime) < CONFIG_CACHE_DURATION_MS) {
    return _configCache.userInfo;
  }
  
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const allEditors = SpreadsheetApp.getActive().getEditors().map(u => u.getEmail().toLowerCase());
    const hasEditAccess = allEditors.includes(email);
    
    const whitelist = _parseEmailList_(_getConfigValue_('ADMIN_WHITELIST'));
    const isAdmin = whitelist.includes(email);

    const userInfo = { email, isAdmin, hasEditAccess };
    _configCache.userInfo = userInfo;
    _configCache.userTime = now;
    return userInfo;
  } catch (e) {
    return { email: '', isAdmin: false, hasEditAccess: false };
  }
}

/**
 * Fetches a specific value from the Konfig sheet.
 * @private
 */
function _getConfigValue_(key, fallback = '') {
  const now = Date.now();
  if (!_configCache.config || (now - _configCache.configTime) > CONFIG_CACHE_DURATION_MS) {
    _configCache.config = _loadAllConfig_();
    _configCache.configTime = now;
  }
  return _configCache.config[String(key || '').toUpperCase()] || fallback;
}

/**
 * Loads all key-value pairs from the Konfig sheet.
 * @private
 */
function _loadAllConfig_() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Konfig');
    if (!sheet || sheet.getLastRow() < 2) return {};
    
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const config = {};
    values.forEach(([key, value]) => {
      if (key) config[String(key).trim().toUpperCase()] = String(value || '').trim();
    });
    return config;
  } catch (e) {
    return {};
  }
}

/**
 * Parses a comma-separated string into an array of lowercase emails.
 * @private
 */
function _parseEmailList_(rawList) {
  if (!rawList) return [];
  return String(rawList).split(/[,;\s]+/).map(s => s.trim().toLowerCase()).filter(s => s.includes('@'));
}