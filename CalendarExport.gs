/**
 * @OnlyCurrentDoc
 *
 * NOTE FOR THE HUMAN REVIEWER:
 * This script needs to be deployed as a Web App for the calendar export to work.
 * Deployment steps:
 * 1. In the Apps Script editor, go to "Deploy" > "New deployment".
 * 2. Select "Web app" as the type.
 * 3. In the configuration:
 *    - Description: "Calendar Export Feed"
 *    - Execute as: "Me (your Google account)"
 *    - Who has access: "Anyone" (This is crucial for external services to access the feed)
 * 4. Click "Deploy".
 * 5. Copy the provided Web app URL. This is the URL that will be used for calendar subscriptions.
 *
 * The `getCalendarExportUrl` function below will programmatically return this URL once deployed.
 * This script assumes that a function `listMeetings_({scope: 'all'})` is available in the global
 * scope of the Google Apps Script project.
 */

/**
 * Web App endpoint to generate and serve the iCalendar (.ics) feed.
 * This function is called when a user accesses the deployed script's URL.
 * @param {object} e - The event parameter for the web app.
 * @returns {ContentService.TextOutput} The calendar feed.
 */
function doGet(e) {
  try {
    // Fetch both meetings and tasks
    const meetings = listMeetings_({ scope: 'all' });
    const tasks = listTasks_(); // Using the new function from Gjoremal_Backend.gs

    // Generate the calendar with both data sources
    const cal = generateICal(meetings, tasks);

    return ContentService.createTextOutput(cal)
      .setMimeType(ContentService.MimeType.ICAL);
  } catch (err) {
    // Log the error for easier debugging from the Apps Script dashboard.
    console.error("Error in doGet for CalendarExport: " + err.toString());
    return ContentService.createTextOutput("Error generating calendar feed: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Returns the public URL of the deployed Web App for the client-side UI.
 * @returns {object} A response object with the URL or an error message.
 */
function getCalendarExportUrl() {
  try {
    // This will only work after the script has been deployed as a web app.
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      return { ok: false, message: "Script has not been deployed as a web app yet. See instructions in CalendarExport.gs." };
    }
    return { ok: true, url: url };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/**
 * Generates an iCalendar (.ics) formatted string from meetings and tasks.
 * @param {Array<object>} meetings - The list of meeting objects.
 * @param {Array<object>} tasks - The list of task objects.
 * @returns {string} The iCalendar data as a string.
 */
function generateICal(meetings, tasks) {
  const calName = "Styrekalender";
  const timeZone = Session.getScriptTimeZone();

  // Helper to format dates for iCal (YYYYMMDDTHHMMSSZ or YYYYMMDD for all-day)
  const toICalDate = (date, allDay = false) => {
    if (allDay) {
      // Format as YYYYMMDD for all-day events
      return date.toISOString().slice(0, 10).replace(/-/g, '');
    }
    // Format as YYYYMMDDTHHMMSSZ for timed events
    return date.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
  };

  let icalString = `BEGIN:VCALENDAR
PRODID:-//Styre-App//NONSGML v1.0//EN
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-CALNAME:${calName}
X-WR-TIMEZONE:${timeZone}
`;

  // Process Meetings
  (meetings || []).forEach(m => {
    if (!m.dato || !m.tittel) return;

    const startDate = new Date(m.dato);
    const endDate = new Date(m.dato);

    if (m.start && typeof m.start === 'string' && m.start.includes(':')) {
      const [startHour, startMinute] = m.start.split(':');
      startDate.setHours(startHour, startMinute, 0, 0);
    } else {
      startDate.setHours(0, 0, 0, 0); // Default to start of the day
    }

    if (m.slutt && typeof m.slutt === 'string' && m.slutt.includes(':')) {
      const [endHour, endMinute] = m.slutt.split(':');
      endDate.setHours(endHour, endMinute, 0, 0);
    } else {
      endDate.setHours(startDate.getHours() + 1, startDate.getMinutes(), 0, 0); // Default 1h duration
    }

    icalString += `BEGIN:VEVENT
UID:${m.id || Utilities.getUuid()}@styreapp.dev
DTSTAMP:${toICalDate(new Date())}
DTSTART:${toICalDate(startDate)}
DTEND:${toICalDate(endDate)}
SUMMARY:${m.tittel}
LOCATION:${m.sted || ''}
DESCRIPTION:Møtetype: ${m.type || 'Ikke spesifisert'}. Status: ${m.status || 'Planlagt'}.
END:VEVENT
`;
  });

  // Process Tasks as all-day events
  (tasks || []).forEach(task => {
    if (!task.dueDate || !task.title) return;

    const dueDate = new Date(task.dueDate);

    // For all-day events, DTSTART is the date, and DTEND is the next day.
    const nextDay = new Date(dueDate);
    nextDay.setDate(nextDay.getDate() + 1);

    icalString += `BEGIN:VEVENT
UID:${task.id || Utilities.getUuid()}@styreapp.dev
DTSTAMP:${toICalDate(new Date())}
DTSTART;VALUE=DATE:${toICalDate(dueDate, true)}
DTEND;VALUE=DATE:${toICalDate(nextDay, true)}
SUMMARY:${task.title}
DESCRIPTION:Oppgave tildelt: ${task.assignee || 'N/A'}. Status: ${task.status || 'Open'}. \\n${task.description || ''}
END:VEVENT
`;
  });

  icalString += 'END:VCALENDAR';
  return icalString;
}