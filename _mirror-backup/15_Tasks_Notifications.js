/* ================== Task Notifications ==================
 * FILE: 15_Tasks_Notifications.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * PURPOSE: Handles automated notifications for task deadlines and completions.
 * ========================================================================== */

const TASK_REMINDER_DAYS_BEFORE = 3; // Reminder will be sent 3 days before the deadline

/**
 * Sends reminders for tasks with upcoming deadlines.
 * This function is intended to be run by a daily time-based trigger.
 */
function checkTaskDeadlines() {
  try {
    const taskSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.TASKS);
    if (!taskSheet) {
      safeLog('TaskNotifications', 'Could not find the task sheet.');
      return;
    }

    const data = taskSheet.getDataRange().getValues();
    const headers = data.shift();
    const c = {
      tittel: headers.indexOf('Tittel'),
      ansvarlig: headers.indexOf('Ansvarlig'),
      status: headers.indexOf('Status'),
      frist: headers.indexOf('Frist'),
    };

    if ([c.tittel, c.ansvarlig, c.status, c.frist].includes(-1)) {
      throw new Error("One or more required columns are missing in the Tasks sheet.");
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const reminderLimit = new Date(today.getTime() + TASK_REMINDER_DAYS_BEFORE * 24 * 60 * 60 * 1000);

    let remindersSent = 0;
    data.forEach(row => {
      const taskStatus = row[c.status];
      const deadline = new Date(row[c.frist]);

      if (['Ny', 'Pågår'].includes(taskStatus) && deadline >= today && deadline <= reminderLimit) {
        const recipient = row[c.ansvarlig];
        if (recipient && recipient.includes('@')) {
          const subject = `Påminnelse: Oppgave "${row[c.tittel]}" har frist snart`;
          const body = `
            <p>Hei,</p>
            <p>Dette er en automatisk påminnelse om at følgende oppgave du er tildelt nærmer seg fristen:</p>
            <ul>
              <li><b>Tittel:</b> ${row[c.tittel]}</li>
              <li><b>Frist:</b> ${deadline.toLocaleDateString('nb-NO')}</li>
              <li><b>Status:</b> ${taskStatus}</li>
            </ul>
            <p>Vennligst se over oppgaven og oppdater statusen ved behov.</p>
            <p>Med vennlig hilsen,<br>${APP.NAME}</p>
          `;
          GmailApp.sendEmail(recipient, subject, '', { htmlBody: body, name: APP.NAME });
          remindersSent++;
        }
      }
    });

    if (remindersSent > 0) {
      safeLog('TaskNotifications', `Sent ${remindersSent} deadline reminders.`);
    }
  } catch (e) {
    safeLog('TaskNotifications_Error', `Error in checkTaskDeadlines: ${e.message}`);
  }
}

/**
 * Sends a notification when a task is marked as complete.
 * This function is called from the main onEdit trigger.
 * @param {Object} e The event object from the onEdit trigger.
 */
function handleTaskStatusChange(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== SHEETS.TASKS) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const c = {
      status: headers.indexOf('Status'),
      tittel: headers.indexOf('Tittel'),
      ansvarlig: headers.indexOf('Ansvarlig'),
    };

    // Exit if the edited column is not 'Status' or if it's the header row
    if (range.getColumn() !== c.status + 1 || range.getRow() === 1) return;

    const newValue = e.value;
    if (newValue !== 'Fullført') return;

    const row = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    const taskTitle = row[c.tittel];
    const assignedUser = row[c.ansvarlig];
    const editor = Session.getActiveUser().getEmail();

    // Get board members to notify
    const personerSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.PERSONER);
    if (!personerSheet) throw new Error("Could not find Personer sheet.");

    const personData = personerSheet.getDataRange().getValues();
    const personHeaders = personData.shift();
    const pc = {
      rolle: personHeaders.indexOf('rolle'),
      epost: personHeaders.indexOf('epost'),
    };

    const boardEmails = personData
      .filter(pRow => ['styremedlem', 'kjernebruker'].includes(String(pRow[pc.rolle]).toLowerCase()))
      .map(pRow => pRow[pc.epost])
      .filter(email => email && email.includes('@')); // Ensure valid emails

    if (boardEmails.length > 0) {
      const subject = `Oppgave fullført: "${taskTitle}"`;
      const body = `
        <p>Hei,</p>
        <p>En oppgave har blitt markert som fullført:</p>
        <ul>
          <li><b>Tittel:</b> ${taskTitle}</li>
          <li><b>Ansvarlig:</b> ${assignedUser}</li>
          <li><b>Fullført av:</b> ${editor}</li>
        </ul>
        <p>Dette er kun til orientering.</p>
        <p>Mvh,<br>${APP.NAME}</p>
      `;

      // Send to all board members
      GmailApp.sendEmail(boardEmails.join(','), subject, '', { htmlBody: body, name: APP.NAME });
      safeLog('TaskNotifications', `Sent completion notification for task "${taskTitle}" to ${boardEmails.length} board members.`);
    }

  } catch (err) {
    safeLog('TaskNotifications_Error', `Error in handleTaskStatusChange: ${err.message}`);
  }
}

/**
 * Installs or verifies the daily trigger for deadline checks.
 * Can be run manually by an admin from the script editor.
 */
function setupTaskNotifications() {
  const functionName = 'checkTaskDeadlines';
  const triggers = ScriptApp.getProjectTriggers();

  const triggerExists = triggers.some(t => t.getHandlerFunction() === functionName);

  if (!triggerExists) {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyDays(1)
      .atHour(8) // Runs every day at 8 AM
      .create();
    SpreadsheetApp.getUi().alert(`✅ Trigger for oppgavepåminnelser er installert og vil kjøre daglig kl. 08:00.`);
    safeLog('TaskNotifications', `Daily trigger for ${functionName} was created.`);
  } else {
    SpreadsheetApp.getUi().alert(`👍 Trigger for oppgavepåminnelser er allerede installert.`);
    safeLog('TaskNotifications', `Trigger for ${functionName} already exists.`);
  }
}