/**
 * Escapes HTML special characters to prevent XSS.
 * @param {*} s The string to escape.
 * @returns {string} The escaped string.
 */
function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}