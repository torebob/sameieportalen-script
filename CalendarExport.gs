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
    // This calls the existing listMeetings_ function from the main backend script.
    const meetings = listMeetings_({ scope: 'all' });
    const cal = generateICal(meetings);

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
 * Generates an iCalendar (.ics) formatted string from a list of meetings.
 * @param {Array<object>} meetings - The list of meeting objects from listMeetings_.
 * @returns {string} The iCalendar data as a string.
 */
function generateICal(meetings) {
  const calName = "Styrekalender";
  const timeZone = Session.getScriptTimeZone();

  let icalString = `BEGIN:VCALENDAR
PRODID:-//Styre-App//NONSGML v1.0//EN
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-CALNAME:${calName}
X-WR-TIMEZONE:${timeZone}
`;

  (meetings || []).forEach(m => {
    if (!m.dato || !m.tittel) {
      return; // Skip meetings without a date or title.
    }

    // Ensure 'dato' is a Date object.
    const startDate = m.dato instanceof Date ? new Date(m.dato.getTime()) : new Date(m.dato);
    const endDate = m.dato instanceof Date ? new Date(m.dato.getTime()) : new Date(m.dato);

    // Handle time. Assumes m.start and m.slutt are "HH:mm" strings.
    if (m.start && typeof m.start === 'string' && m.start.includes(':')) {
      const [startHour, startMinute] = m.start.split(':');
      startDate.setHours(startHour, startMinute, 0, 0);
    } else {
      // If no start time, it's an all-day event. We'll handle this by not including time in the DTSTART/DTEND.
      // However, for simplicity and compatibility, we'll set a default time.
      // A true all-day event would format the date as VALUE=DATE:YYYYMMDD.
      startDate.setHours(0, 0, 0, 0);
    }

    if (m.slutt && typeof m.slutt === 'string' && m.slutt.includes(':')) {
      const [endHour, endMinute] = m.slutt.split(':');
      endDate.setHours(endHour, endMinute, 0, 0);
    } else {
      // If no end time, default to a 1-hour duration from the start time.
      endDate.setHours(startDate.getHours() + 1, startDate.getMinutes(), 0, 0);
    }

    // Format dates to the required iCal UTC format (YYYYMMDDTHHMMSSZ).
    const toUTC = (date) => {
      return date.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
    }

    const uid = `${m.id || Utilities.getUuid()}@styreapp.dev`; // Generate a UUID if id is missing.
    const summary = m.tittel;
    const location = m.sted || '';
    const description = `Møtetype: ${m.type || 'Ikke spesifisert'}. Status: ${m.status || 'Planlagt'}.`;
    const dtstamp = toUTC(new Date()); // Timestamp for when the event was created/modified.

    icalString += `BEGIN:VEVENT
UID:${uid}
DTSTAMP:${dtstamp}
DTSTART:${toUTC(startDate)}
DTEND:${toUTC(endDate)}
SUMMARY:${summary}
LOCATION:${location}
DESCRIPTION:${description}
END:VEVENT
`;
  });

  icalString += 'END:VCALENDAR';
  return icalString;
}