/* ============================================================================
 * Board-Level Task Management API
 * FILE: 19_Tasks_API.js
 *
 * PURPOSE:
 * - Provides board members with the necessary tools to manage task statuses.
 * - This API is distinct from the Vaktmester API and has different permissions.
 * ========================================================================== */

(() => {
  const SH = globalThis.SHEETS || { TASKS: 'Oppgaver', BOARD: 'Styret' };
  const CACHE = CacheService.getScriptCache();
  const CACHE_KEY = 'board_members_emails_list';
  const CACHE_EXPIRATION_SECONDS = 3600; // 1 hour

  /**
   * Normalizes an email address.
   * @param {string} s The email string.
   * @returns {string} The normalized email.
   */
  const _normalizeEmail_ = (s) => {
    if (!s) return '';
    let str = String(s).trim();
    const m = str.match(/<([^>]+)>/);
    if (m) str = m[1];
    str = str.replace(/^mailto:/i, '');
    return str.toLowerCase();
  };

  /**
   * Checks if the current active user is a board member.
   * Caches the list of board members for performance.
   * @returns {boolean} True if the user is a board member.
   */
  const _isBoardMember_ = () => {
    const userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail());
    if (!userEmail) return false;

    let boardEmailsJson = CACHE.get(CACHE_KEY);
    let boardEmails;

    if (boardEmailsJson) {
      boardEmails = JSON.parse(boardEmailsJson);
    } else {
      try {
        const boardSheet = SpreadsheetApp.getActive().getSheetByName(SH.BOARD);
        if (!boardSheet || boardSheet.getLastRow() < 2) {
          boardEmails = [];
        } else {
          boardEmails = boardSheet
            .getRange(2, 2, boardSheet.getLastRow() - 1, 1) // Assumes email is in column B
            .getValues()
            .flat()
            .map(_normalizeEmail_)
            .filter(Boolean);
        }
        CACHE.put(CACHE_KEY, JSON.stringify(boardEmails), CACHE_EXPIRATION_SECONDS);
      } catch (e) {
        console.error(`Error reading board member sheet: ${e.message}`);
        boardEmails = [];
      }
    }
    return boardEmails.includes(userEmail);
  };

  /**
   * Updates the status of a task. Accessible only by board members.
   * @param {string} taskId The ID of the task to update.
   * @param {string} newStatus The new status to set.
   * @returns {object} An object indicating success or failure.
   */
  function updateTaskStatusByBoard(taskId, newStatus) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      if (!_isBoardMember_()) {
        throw new Error('Authorization failed: User is not a board member.');
      }

      const normStatus = String(newStatus || '').trim();
      const validStatuses = ['Ny', 'Pågår', 'Venter', 'Fullført', 'Avvist'];
      if (!normStatus || !validStatuses.includes(normStatus)) {
        throw new Error(`Invalid status provided: "${newStatus}".`);
      }

      const sh = SpreadsheetApp.getActive().getSheetByName(SH.TASKS);
      if (!sh) throw new Error('Task sheet not found.');

      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const c = H.reduce((acc, header, i) => {
        acc[String(header).trim()] = i;
        return acc;
      }, {});

      const idCol = H.indexOf('OppgaveID');
      if (idCol === -1) throw new Error('OppgaveID column not found.');

      const tf = sh.getRange(2, idCol + 1, sh.getLastRow() -1).createTextFinder(String(taskId)).matchEntireCell(true).findNext();
      if (!tf) throw new Error(`Task with ID ${taskId} not found.`);

      const rowNum = tf.getRow();
      const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn());
      const rowVals = row.getValues()[0];

      // Update status
      rowVals[c.Status] = normStatus;

      // Add audit comment
      const userEmail = _normalizeEmail_(Session.getActiveUser()?.getEmail());
      const trimmedComment = `Status endret til '${normStatus}' av styremedlem.`;
      if (c.Kommentarer > -1) {
        const existing = rowVals[c.Kommentarer] || '';
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        const line = `[${timestamp} - ${userEmail}]: ${trimmedComment}`;
        rowVals[c.Kommentarer] = existing ? `${existing}\n${line}` : line;
      }

      row.setValues([rowVals]);

      console.log(`Task ${taskId} updated to ${normStatus} by ${userEmail}`);
      return { ok: true, message: `Task ${taskId} updated successfully.` };

    } catch (e) {
      console.error(`updateTaskStatusByBoard Error: ${e.message}`);
      return { ok: false, error: e.message };
    } finally {
      lock.releaseLock();
    }
  }

  // Expose the function to be callable from the client-side
  globalThis.updateTaskStatusByBoard = updateTaskStatusByBoard;

})();