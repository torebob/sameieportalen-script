/* ====================== Rapporter & Årshjul/Calendar ======================
 * FILE: 07_Reporting_YearWheel.gs  |  VERSION: 1.8.0  |  UPDATED: 2025-09-14
 * FORMÅL: Generere rapporter og synkronisere årshjul til kalender.
 * NYTT v1.8.0: Lagt til rapport for åpne saker per kategori (K17-02).
 * ====================================================================== */

/* -------- Lokale fallbacks og konstanter -------- */
// ... (eksisterende hjelpefunksjoner og konstanter forblir uendret) ...

/* ===================== Gevinst-rapport ===================== */

function generateAnnualBenefitsReport(){
  // ... (eksisterende funksjon forblir uendret) ...
}

/* ===================== NY: Rapport - Åpne Saker per Kategori ===================== */

/**
 * K17-02: Genererer en rapport over åpne saker, gruppert etter kategori.
 * Viser resultatet i en pen dialogboks for enkel oversikt.
 */
function generateOpenCasesReport() {
  const ui = (typeof _ui === 'function') ? _ui() : SpreadsheetApp.getUi();
  try {
    // Bruker den sentrale tilgangskontrollen
    if (typeof requirePermission === 'function') {
      requirePermission('GENERATE_REPORTS');
    } else if (typeof hasPermission !== 'function' || !hasPermission('GENERATE_REPORTS')) {
      throw new Error("Du har ikke tilgang til å generere rapporter.");
    }

    const tasksSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.TASKS);
    if (!tasksSheet || tasksSheet.getLastRow() < 2) {
      return _alert_('Fant ingen oppgaver å rapportere på.', 'Rapport: Åpne Saker');
    }

    const data = tasksSheet.getDataRange().getValues();
    const headers = data.shift(); // Fjerner og returnerer header-raden
    const cStatus = headers.indexOf('Status');
    const cKategori = headers.indexOf('Kategori');

    if (cStatus === -1 || cKategori === -1) {
      throw new Error("Mangler påkrevde kolonner ('Status', 'Kategori') i Oppgaver-arket.");
    }

    const openStatuses = new Set(['ny', 'påbegynt', 'venter']);
    const categoryCounts = {};

    data.forEach(row => {
      const status = String(row[cStatus] || '').toLowerCase();
      if (openStatuses.has(status)) {
        const category = String(row[cKategori] || 'Ukategorisert').trim();
        categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      }
    });

    // Bygg HTML for en pen rapport
    let reportHtml = `
      <style>
        body { font-family: Inter, sans-serif; padding: 15px; color: #1f2937; }
        h2 { margin-top: 0; }
        table { border-collapse: collapse; width: 100%; font-size: 14px; }
        th, td { border: 1px solid #e5e7eb; padding: 10px; text-align: left; }
        th { background-color: #f9fafb; font-weight: 600; }
      </style>
      <h2>Rapport: Åpne Saker per Kategori</h2>
      <table>
        <tr><th>Kategori</th><th>Antall åpne saker</th></tr>
    `;

    const sortedCategories = Object.keys(categoryCounts).sort();
    
    if (sortedCategories.length === 0) {
        reportHtml += '<tr><td colspan="2" style="color: #6b7280;">Fant ingen åpne saker.</td></tr>';
    } else {
        for (const category of sortedCategories) {
          reportHtml += `<tr><td>${category}</td><td>${categoryCounts[category]}</td></tr>`;
        }
    }
    
    reportHtml += '</table>';

    const htmlOutput = HtmlService.createHtmlOutput(reportHtml)
      .setWidth(450)
      .setHeight(350);
    
    ui.showModalDialog(htmlOutput, 'Rapport: Åpne Saker');
    
    if (typeof _logEvent === 'function') _logEvent('Rapport', 'Genererte rapport for åpne saker.');

  } catch (e) {
    _alert_(e.message, 'Feil ved rapportgenerering');
    if (typeof _logEvent === 'function') _logEvent('Rapport_Feil', e.message);
  }
}


/* ===================== Årshjul → Kalender ===================== */

function syncYearWheelToCalendar(){
  // ... (eksisterende funksjon forblir uendret) ...
}

/* ===================== Daglige triggere ===================== */
// ... (resten av filen er uendret) ...

