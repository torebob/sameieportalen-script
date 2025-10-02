/* ======================= AI-Støttet Risikohåndtering =======================
 * FILE: 70_RiskAnalysis_API.js | VERSION: 1.0.0 | CREATED: 2025-09-28
 * FORMÅL: Analysere oppgaver, avvik og møtereferater for å foreslå risikoområder.
 * KRAV: AI-03
 * ============================================================================== */

((global) => {
  const { SHEETS } = global; // Assuming SHEETS is a global object with sheet names

  // Fallback if global SHEETS is not defined
  const AppSheets = SHEETS || {
    TASKS: 'Oppgaver',
    MOTE_SAKER: 'MOTE_SAKER',
    MOTE_KOMMENTARER: 'MOTE_KOMMENTARER',
  };

  const RiskAnalysis = {};

  /**
   * Helper function to safely read all data from a sheet.
   * @param {string} sheetName The name of the sheet to read.
   * @returns {Array<Array<string>>} 2D array of sheet data, or empty array on error.
   */
  const _getSheetData = (sheetName) => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) {
        console.warn(`RiskAnalysis: Sheet '${sheetName}' is missing or empty.`);
        return [];
      }
      return sheet.getDataRange().getValues();
    } catch (e) {
      console.error(`RiskAnalysis: Failed to read sheet '${sheetName}': ${e.message}`);
      return [];
    }
  };

  /**
   * Collects data from all relevant sources for analysis.
   * @returns {object} An object containing arrays of tasks, meeting items, and comments.
   */
  RiskAnalysis.collectData = () => {
    const tasksData = _getSheetData(AppSheets.TASKS);
    const sakerData = _getSheetData(AppSheets.MOTE_SAKER);
    const commentsData = _getSheetData(AppSheets.MOTE_KOMMENTARER);

    const tasks = tasksData.length > 1 ? tasksData.slice(1).map(row => ({
      id: row[0],
      title: row[1],
      description: row[2],
      status: row[6],
      priority: row[7],
      comments: row[9],
      category: row[11],
    })) : [];

    const saker = sakerData.length > 1 ? sakerData.slice(1).map(row => ({
      meetingId: row[0],
      caseId: row[1],
      title: row[3],
      proposal: row[4],
      decision: row[5],
    })) : [];

    const comments = commentsData.length > 1 ? commentsData.slice(1).map(row => ({
      caseId: row[0],
      timestamp: row[1],
      from: row[2],
      text: row[3],
    })) : [];

    return { tasks, saker, comments };
  };

  /**
   * A rule-based engine to analyze text for potential risks.
   * @param {object} data The collected data from sheets.
   * @returns {Array<object>} A list of identified risks.
   */
  RiskAnalysis.analyze = (data) => {
    const identifiedRisks = [];
    const riskKeywords = {
      'Vann/Fukt': ['lekkasje', 'vann', 'fukt', 'kondens', 'avløp', 'tett', 'rør'],
      'Brannsikkerhet': ['brann', 'røyk', 'varsler', 'slukkeutstyr', 'nødutgang', 'brannmur'],
      'Elektrisk': ['strømbrudd', 'jordfeil', 'sikring', 'elektrisk', 'støt'],
      'Bygningsmessig': ['sprekk', 'skade', 'murpuss', 'tak', 'grunnmur', 'fasade', 'vindu', 'dør'],
      'Skadedyr': ['mus', 'rotter', 'skadedyr', 'insekt', 'maur'],
      'HMS': ['ulykke', 'skade', 'fall', 'sikkerhet', 'avvik'],
    };

    const searchInText = (text, category, source, reference) => {
      if (!text) return;
      const keywords = riskKeywords[category];
      keywords.forEach(keyword => {
        if (new RegExp(`\\b${keyword}\\b`, 'i').test(text)) {
          identifiedRisks.push({
            category: category,
            keyword: keyword,
            source: source,
            reference: reference,
            text: text.substring(0, 200) + (text.length > 200 ? '...' : ''),
          });
        }
      });
    };

    // Analyze Tasks
    data.tasks.forEach(task => {
      const fullText = `${task.title} ${task.description} ${task.comments}`;
      for (const category in riskKeywords) {
        searchInText(fullText, category, 'Oppgave', `ID: ${task.id}`);
      }
    });

    // Analyze Meeting Agendas and Decisions
    data.saker.forEach(sak => {
      const fullText = `${sak.title} ${sak.proposal} ${sak.decision}`;
      for (const category in riskKeywords) {
        searchInText(fullText, category, 'Møtesak', `Sak: ${sak.caseId}`);
      }
    });

    // Analyze Meeting Comments
    data.comments.forEach(comment => {
      for (const category in riskKeywords) {
        searchInText(comment.text, category, 'Innspill/Kommentar', `Sak: ${comment.caseId}`);
      }
    });

    return identifiedRisks;
  };

  /**
   * Main entry point for running the risk analysis.
   * @returns {object} A structured object containing the analysis results.
   */
  RiskAnalysis.run = () => {
    const data = RiskAnalysis.collectData();
    const identifiedRisks = RiskAnalysis.analyze(data);

    return {
      ok: true,
      message: "Risk analysis completed successfully.",
      results: {
        identifiedRisks: identifiedRisks,
        summary: `Analyse fullført. Fant ${identifiedRisks.length} potensielle risikoer.`,
      }
    };
  };

  // Expose the RiskAnalysis namespace to the global scope
  global.runRiskAnalysis = RiskAnalysis.run;

})(this);