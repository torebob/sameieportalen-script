/**
 * Main entry point for the web app.
 * @param {object} e The event parameter for a web app doGet request.
 * @returns {HtmlOutput} The HTML output for the page.
 */
function doGet(e) {
    const action = e.parameter.action;

    if (action === 'admin') {
        return HtmlService.createHtmlOutputFromFile('Admin_Panel')
            .setTitle('Admin Panel');
    }

    if (action === 'edit') {
        const pageId = e.parameter.page;
        if (!pageId) {
            return HtmlService.createHtmlOutput('Side-ID er påkrevd for redigering.');
        }
        const template = HtmlService.createTemplateFromFile('Edit_Page');
        template.pageId = pageId;
        return template.evaluate().setTitle(`Redigerer: ${pageId}`);
    }

    const page = e.parameter.page || 'home';
    const template = HtmlService.createTemplateFromFile('Website_Template');
    template.page = page;
    return template.evaluate()
        .setTitle('Sameiet Hjemmeside')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Gets the list of news articles.
 * @returns {Array<object>} A list of news articles.
 */
function getNewsFeed() {
  try {
    const articles = DB.query('News');
    return articles.sort((a, b) => new Date(b.publishedDate) - new Date(a.publishedDate));
  } catch (e) {
    console.error("Error in getNewsFeed: " + e.message);
    return [];
  }
}

/**
 * Gets the list of documents.
 * @returns {Array<object>} A list of documents.
 */
function getDocuments() {
  try {
    return DB.query('Documents');
  } catch (e) {
    console.error("Error in getDocuments: " + e.message);
    return [];
  }
}

/**
 * Gets the content for a specific page from the spreadsheet.
 * @param {string} pageId The ID of the page to retrieve.
 * @returns {object} The page content or null if not found.
 */
function getPageContent(pageId, password) {
  try {
    const results = DB.query('WebsitePages', { pageId: pageId });
    if (results.length === 0) {
      return null;
    }

    const page = results[0];
    const pagePassword = page.password;

    if (pagePassword && pagePassword !== password) {
      return { authRequired: true };
    }

    // Fjern passord fra objektet som sendes til klienten
    delete page.password;
    return page;

  } catch (e) {
    console.error("Error in getPageContent: " + e.message);
    return null;
  }
}


function verifyPassword(pageId, password) {
    const pageContent = getPageContent(pageId, password);
    if (pageContent && !pageContent.authRequired) {
        return pageContent;
    }
    return { ok: false, message: 'Ugyldig passord' };
}

/**
 * Includes the content of another HTML file.
 * Used for including CSS and JS files in the main template.
 * @param {string} filename The name of the file to include.
 * @returns {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Saves the content of a page to the spreadsheet.
 * @param {string} pageId The ID of the page to save.
 * @param {string} content The new HTML content of the page.
 * @returns {object} A success or error object.
 */
function savePageContent(pageId, content) {
    try {
        const existingPage = DB.query('WebsitePages', { pageId: pageId });

        if (existingPage.length > 0) {
            DB.update('WebsitePages', pageId, { content: content });
        } else {
            const newPage = {
                pageId: pageId,
                title: `Ny side (${pageId})`,
                content: content,
                password: ''
            };
            DB.insert('WebsitePages', newPage);
        }
        return { ok: true };
    } catch (e) {
        console.error("Error in savePageContent: " + e.message);
        return { ok: false, message: e.message };
    }
}


// --- Booking System Functions ---

function listResources() {
    try {
        const resources = DB.query('CommonResources');
        return { ok: true, resources: resources };
    } catch (e) {
        console.error("Error in listResources: " + e.message);
        return { ok: false, message: e.message };
    }
}

function getBookings(resourceId, year, month) {
    try {
        // Hent alle bookinger for ressursen
        const allBookings = DB.query('Bookings', { resourceId: resourceId });

        // Filtrer i minnet basert på år og måned
        const bookings = allBookings.filter(booking => {
            const bookingDate = new Date(booking.startTime);
            return bookingDate.getFullYear() === year && bookingDate.getMonth() === month;
        });

        return { ok: true, bookings: bookings };
    } catch (e) {
        console.error("Error in getBookings: " + e.message);
        return { ok: false, message: e.message };
    }
}

function createBooking(bookingDetails) {
    try {
        // Get the currently logged-in user. This is the security fix.
        const user = getCurrentUser();
        const { email: userEmail, name: userName } = user;

        const { resourceId, startTime, endTime } = bookingDetails;
        const start = new Date(startTime);
        const end = new Date(endTime);

        // --- Conflict Check ---
        const allBookings = DB.query('Bookings', { resourceId: resourceId });
        const conflictingBooking = allBookings.find(booking => {
            const existingStart = new Date(booking.startTime);
            const existingEnd = new Date(booking.endTime);
            // Check for overlap: (StartA < EndB) and (EndA > StartB)
            return start < existingEnd && end > existingStart;
        });

        if (conflictingBooking) {
            return { ok: false, message: "Tiden er allerede booket. Vennligst velg en annen tid." };
        }

        // --- Create Booking ---
        const newBooking = {
            resourceId: resourceId,
            startTime: startTime,
            endTime: endTime,
            userEmail: userEmail,
            userName: userName,
            status: 'Confirmed' // Eksempel på status
        };
        const insertedBooking = DB.insert('Bookings', newBooking);


        // Log the audit event for GDPR compliance and tracking
        logAuditEvent('CREATE_BOOKING', 'Bookings', { resourceId, startTime, endTime });

        // --- Get Resource Name for Email ---
        const resource = DB.getById('CommonResources', resourceId);
        const resourceName = resource ? resource.name : 'Ukjent Ressurs';


        // --- Send Confirmation Email ---
        const subject = "Booking bekreftelse";
        const body = `
            Hei ${userName},

            Din booking er bekreftet:
            Ressurs: ${resourceName}
            Starttid: ${start.toLocaleString('no-NO')}
            Sluttid: ${end.toLocaleString('no-NO')}

            Takk!
        `;
        // Using a try-catch for the email in case of permission issues,
        // so it doesn't block the booking itself.
        try {
            // Send to the authenticated user's email
            MailApp.sendEmail(userEmail, subject, body);
        } catch(e) {
            console.error("Kunne ikke sende bekreftelses-epost: " + e.message);
            // Don't fail the whole operation, just log the error.
        }

        return { ok: true, id: insertedBooking.id };
    } catch (e) {
        // Provide a more specific error message if not authenticated.
        if (e.message.includes("Ikke autentisert")) {
            return { ok: false, message: "Du må være logget inn for å booke." };
        }
        console.error("Error in createBooking: " + e.message);
        return { ok: false, message: "En serverfeil oppstod under bookingen: " + e.message };
    }
}

/**
 * Gets the raw HTML for the booking page.
 * @returns {string} The HTML content of Booking.html.
 */
function getBookingPageHtml() {
    return HtmlService.createHtmlOutputFromFile('Booking.html').getContent();
}