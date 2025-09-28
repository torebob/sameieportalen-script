/* ============================================================================
 * Admin Security Module
 * FILE: 00_Admin_Security.js
 *
 * PURPOSE:
 * - Provides a centralized function to enforce administrator-only access
 *   for sensitive server-side functions.
 * ========================================================================== */

/**
 * Checks if the current user is an administrator.
 * Throws an error if the user is not authorized.
 *
 * This function is the designated security gate for all administrative actions.
 */
function _requireAdmin_() {
  const userEmail = Session.getActiveUser().getEmail();
  const adminWhitelistStr = PropertiesService.getScriptProperties().getProperty('ADMIN_WHITELIST') || '';
  const adminWhitelist = adminWhitelistStr.split(',').map(e => e.trim().toLowerCase()).filter(e => e);

  if (!adminWhitelist.includes(userEmail.toLowerCase())) {
    throw new Error('Unauthorized: You do not have permission to perform this action.');
  }
}