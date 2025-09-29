/**
 * ==============================================================================
 * DRAFT: Accounting Integration Backend
 *
 * Dette skriptet håndterer integrasjon mot eksterne regnskapssystemer.
 * Funksjoner her vil dekke lagring av API-nøkler, import av data (budsjett)
 * og eksport av data (faktura).
 *
 * RI.17.1, RI.17.2, RI.17.3, RI.17.4
 * ==============================================================================
 */

// --- Constants ---
const ACCOUNTING_SYSTEMS = {
  XLEDGER: 'Xledger',
  VISMA: 'Visma e-conomic',
  FORTNOX: 'Fortnox'
};

const MOCK_API_ENDPOINT = 'https://api.mocki.io/v2/a3d3e6f3'; // For testing purposes

/**
 * ==============================================================================
 * SECTION: Settings Management (RI.17.1)
 * ==============================================================================
 */

/**
 * Saves accounting integration settings securely.
 * @param {object} settings The settings object, e.g., { system: 'Xledger', apiKey: '...' }
 * @returns {object} A response object.
 */
function saveAccountingSettings(settings) {
  try {
    if (!settings || !settings.system || !settings.apiKey) {
      throw new Error('Mangler system eller API-nøkkel.');
    }
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('accounting_system', settings.system);
    userProperties.setProperty('accounting_api_key', settings.apiKey);
    return { ok: true, message: 'Innstillinger lagret.' };
  } catch (e) {
    Logger.log(`Feil ved lagring av regnskapsinnstillinger: ${e.message}`);
    return { ok: false, error: `Kunne ikke lagre innstillinger: ${e.message}` };
  }
}

/**
 * Retrieves the currently saved accounting integration settings.
 * @returns {object} A response object with the settings.
 */
function getAccountingSettings() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const system = userProperties.getProperty('accounting_system');
    const apiKey = userProperties.getProperty('accounting_api_key');
    return { ok: true, settings: { system, apiKey } };
  } catch (e) {
    Logger.log(`Feil ved henting av regnskapsinnstillinger: ${e.message}`);
    return { ok: false, error: `Kunne ikke hente innstillinger: ${e.message}` };
  }
}

/**
 * ==============================================================================
 * SECTION: Data Import (RI.17.2)
 * ==============================================================================
 */

/**
 * Imports budget data from the external accounting system.
 * For this draft, it uses a mock API.
 * @param {number} year The year to import the budget for.
 * @returns {object} A response object with the imported budget data.
 */
function importBudgetFromAccounting(year) {
  const settings = getAccountingSettings();
  if (!settings.ok || !settings.settings.system || !settings.settings.apiKey) {
    return { ok: false, error: 'Regnskapssystem er ikke konfigurert.' };
  }

  try {
    // In a real scenario, the URL and options would depend on the selected system.
    const url = `${MOCK_API_ENDPOINT}/budget?year=${year}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${settings.settings.apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    if (responseCode >= 400) {
      throw new Error(`API-feil (${responseCode}): ${content}`);
    }

    const budgetData = JSON.parse(content);

    // Here, you would typically process `budgetData` and save it to your budget sheet.
    // For this draft, we just return the data.
    return { ok: true, data: budgetData.items };

  } catch (e) {
    Logger.log(`Feil ved import av budsjett: ${e.message}`);
    notifyAdminOfError(`Feil ved import av budsjett for år ${year}`, e.message);
    return { ok: false, error: `Import feilet: ${e.message}` };
  }
}


/**
 * ==============================================================================
 * SECTION: Data Export (RI.17.3)
 * ==============================================================================
 */

/**
 * Exports an invoice to the external accounting system.
 * @param {object} invoiceData The invoice data to export.
 * @returns {object} A response object.
 */
function exportInvoiceToAccounting(invoiceData) {
  const settings = getAccountingSettings();
  if (!settings.ok || !settings.settings.system || !settings.settings.apiKey) {
    return { ok: false, error: 'Regnskapssystem er ikke konfigurert.' };
  }

  try {
    const url = `${MOCK_API_ENDPOINT}/invoices`;
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${settings.settings.apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(invoiceData),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    if (responseCode >= 400) {
      throw new Error(`API-feil (${responseCode}): ${content}`);
    }

    return { ok: true, message: 'Faktura eksportert.', data: JSON.parse(content) };
  } catch (e) {
    Logger.log(`Feil ved eksport av faktura: ${e.message}`);
    notifyAdminOfError('Feil ved eksport av faktura', e.message);
    return { ok: false, error: `Eksport feilet: ${e.message}` };
  }
}


/**
 * ==============================================================================
 * SECTION: Error Handling & Notifications (RI.17.4)
 * ==============================================================================
 */

/**
 * Notifies an administrator about a critical integration error.
 * @param {string} subject The subject of the notification.
 * @param {string} details The error details.
 */
function notifyAdminOfError(subject, details) {
  try {
    // In a real app, get this from a config or user database
    const adminEmail = Session.getActiveUser().getEmail();

    MailApp.sendEmail({
      to: adminEmail,
      subject: `[Varsel] Regnskapsintegrasjon: ${subject}`,
      body: `Det oppstod en feil i integrasjonen med regnskapssystemet.\n\nDetaljer:\n${details}`
    });
  } catch (e) {
    Logger.log(`Kunne ikke sende varsel-e-post: ${e.message}`);
  }
}