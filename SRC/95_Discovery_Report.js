/* =============================================================================
 * Assisted Discovery / Kravsrapport – Generator + Meny + Forslag til "Krav"
 * FILE: 95_Discovery_Report.gs
 * VERSION: 4.4.0
 * UPDATED: 2025-09-16
 *
 * PURPOSE:
 * - Automatisert "reverse engineering": finn funksjoner, triggere, ark og felter.
 * - Generer en Google Doc-rapport med krav, avvik, avhengigheter og teknisk info.
 * - Konfigurerbar via «Konfig»-arket. Skalerer til store prosjekter (progress/tidsvern).
 *
 * ----------------------------------------------------------------------------- 
 * VERSJONSHISTORIKK
 * ----------------------------------------------------------------------------- 
 * v4.4.0 (2025-09-16):
 * - MERGE: 4.2.1 (din) + 4.3.0 (min) → ny forbedret utgave.
 * - ADD: Kilde-type, -detalj, -lenke kolonner i «Krav» + automatisk utfylling.
 * - ADD: Klikkbare «Åpne»-lenker i rapport for ark, header-rad og skript-IDE.
 * - ADD: Detaljert kildevisning i rapport for kravforslag (funksjon/ark/felt/trigger).
 * - FIX: Robust DocumentApp-tabellbygging (ingen array-append på rader).
 * - FIX: Validering av dokument-ID + fallback til ny opprettelse hvis korrupt.
 * - IMPROVE: Trygge batch-størrelser, bedre timeout-sjekk og statusmeldinger.
 * - IMPROVE: Konfig leses uten sideeffekt (ingen tvungen skriving).
 *
 * v4.2.1 (2025-09-16):
 * - ADD: Detaljert kilde for kravforslag i rapporten (funksjonsnavn, arknavn, feltnavn).
 *
 * v4.2.0 (2025-09-16):
 * - MERGE 4.1.1-varianter. Korrekt tabellbygging, auto «Hensikt & mål»,
 *   openById-fallback, trygg ark-batching, bedre feilmeldinger/status.
 * ============================================================================ */


/* -------------------------- Namespace & Konfigurasjon ----------------------- */
(function (glob) {
  const loadedConfig = loadConfiguration_(); // les fra «Konfig», fallback til defaults

  const CONFIG = {
    KONFIG_SHEET: 'Konfig',
    KRAV_SHEET: 'Krav',
    DISCOVERY_DOC_KEY: 'DISCOVERY_DOC_ID',
    MENU_NAME: 'Analyse',
    KRAV_HEADERS: [
      'ID','Tittel','Kravtekst','Prioritet','Hensikt & mål','Verifikasjon / Test',
      'Kilde-type','Kilde-detalj','Kilde-lenke',
      'Fremdrift 1','Fremdrift 2','Fremdrift 3','Fremdrift 4','Fremdrift 5'
    ],
    DEFAULT_PRIORITY: 'BØR',

    // Konfigverdier (kan settes i «Konfig»)
    DEDUPE_JACCARD: loadedConfig.DEDUPE_JACCARD,
    MAX_SUGGESTIONS_IN_DOC: loadedConfig.MAX_SUGGESTIONS_IN_DOC,
    BATCH_SIZE: loadedConfig.BATCH_SIZE,
    MAX_EXECUTION_TIME: loadedConfig.MAX_EXECUTION_TIME, // ms
    MAX_SHEET_COLUMNS: loadedConfig.MAX_SHEET_COLUMNS,
    LARGE_SHEET_WARNING_THRESHOLD: loadedConfig.LARGE_SHEET_WARNING_THRESHOLD,
    LARGE_FUNCS_THRESHOLD: loadedConfig.LARGE_FUNCS_THRESHOLD,
    LARGE_SHEETS_THRESHOLD: loadedConfig.LARGE_SHEETS_THRESHOLD
  };

  const MESSAGES = {
    REPORT_READY: '✅ Discovery-rapport (v4.4) er klar!',
    REPORT_FAIL: 'Feil ved generering: {err}',
    ANALYZING: 'Analyserer prosjekt… Dette kan ta flere minutter.',
    SUGGEST_OK: '💡 {n} kravforslag lagt til i «Krav».',
    SUGGEST_NONE: 'Ingen nye forslag å legge til.',
    OPENING_DOC: 'Åpner dokument…',
    PROGRESS_FUNCTIONS: 'Analyserer funksjoner... ({count} funnet)',
    PROGRESS_TRIGGERS: 'Analyserer triggere... ({count} aktive)',
    PROGRESS_SHEETS: 'Analyserer ark... (Batch {current}/{total})',
    TIMEOUT_WARNING: '⚠️ Analyse stoppet pga. tidsbegrensning. Rapport kan være ufullstendig.',
    PERMISSION_ERROR: 'Mangler tilgang til å opprette dokumenter. Kontakt administrator.',
    LARGE_PROJECT_WARNING: 'Dette er et stort prosjekt ({sheets} ark, {functions} funksjoner). Analysen kan ta 3–6 minutter.'
  };

  const NS = glob.DISCOVERY_REPORT || {};
  NS.VERSION = '4.4.0';
  NS.CONFIG = CONFIG;
  NS.MESSAGES = MESSAGES;
  glob.DISCOVERY_REPORT = NS;
})(globalThis);


/* -------------------------------- Meny (UI) - DEPRECATED -------------------------------- */
/*
 * MERK: Menyopprettelse for Discovery-verktøyene er flyttet til 00_App_Core.js.
 * Funksjonene discoveryInstallMenuOnOpen, discoveryMenuBuildQuick, og discoveryRegisterMenu_
 * er fjernet for å unngå onOpen-konflikter. Menyen vises nå under Admin-menyen.
 */


/* ---------------------------- Hovedkommandoer ------------------------------ */
function generateDiscoveryReportInDoc() {
  const startTime = Date.now();
  let timeoutOccurred = false;

  try {
    const projectSize = assessProjectSize_();
    if (projectSize.isLarge) {
      const message = DISCOVERY_REPORT.MESSAGES.LARGE_PROJECT_WARNING
        .replace('{sheets}', projectSize.sheets)
        .replace('{functions}', projectSize.functions);
      const response = SpreadsheetApp.getUi().alert(
        'Stort prosjekt oppdaget',
        message + '\n\nVil du fortsette?',
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (response !== SpreadsheetApp.getUi().Button.YES) return;
    } else {
      SpreadsheetApp.getUi().alert(
        'Starter analyse...',
        'Klikk OK for å fortsette. Følg med på statusmeldingene nederst.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

    SpreadsheetApp.getActive().toast(DISCOVERY_REPORT.MESSAGES.ANALYZING);

    const analysis = performComprehensiveAnalysis_(startTime);

    // Tidsvern
    if (Date.now() - startTime > DISCOVERY_REPORT.CONFIG.MAX_EXECUTION_TIME) {
      timeoutOccurred = true;
      SpreadsheetApp.getActive().toast(DISCOVERY_REPORT.MESSAGES.TIMEOUT_WARNING);
    }

    const existingReqs = readRequirementsForGapAnalysis_();
    const candidates  = generateRequirementCandidates_(analysis);
    const deduped     = dedupeCandidates_(candidates, existingReqs, DISCOVERY_REPORT.CONFIG.DEDUPE_JACCARD);
    const gap         = performGapAnalysis_(analysis, existingReqs, deduped);

    // Dokument-ID med validering + fallback
    let docId = getOrCreateDiscoveryDocId_();
    if (!docId) {
      throw new Error('Klarte ikke hente eller opprette rapportdokumentet. Sjekk «Konfig» og tillatelser.');
    }

    let doc;
    try {
      doc = DocumentApp.openById(docId);
    } catch (e) {
      Logger.log('openById feilet, prøver å opprette nytt dokument. Årsak: ' + e.message);
      docId = getOrCreateDiscoveryDocId_(true);
      doc = DocumentApp.openById(docId);
    }

    buildDiscoveryReportDocument_(doc.getBody().clear(), analysis, existingReqs, deduped, gap, timeoutOccurred);

    doc.saveAndClose();
    SpreadsheetApp.getActive().toast(DISCOVERY_REPORT.MESSAGES.REPORT_READY);
    openInNewTab_(doc.getUrl());

    if (projectSize.isLarge) Utilities.sleep(500); // hint
  } catch (e) {
    const errorMsg = handleError_(e);
    SpreadsheetApp.getUi().alert('En feil oppstod', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Discovery report generation failed: ' + (e && e.stack || e));
    throw e;
  }
}

function openDiscoveryDocQuick() {
  try {
    const id = getOrCreateDiscoveryDocId_();
    const url = 'https://docs.google.com/document/d/' + id + '/edit';
    SpreadsheetApp.getActive().toast(DISCOVERY_REPORT.MESSAGES.OPENING_DOC);
    openInNewTab_(url);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Feil', 'Kunne ikke åpne dokument: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function discoverySuggestToKravQuick() {
  try {
    const sh = ensureKravSheet_();
    SpreadsheetApp.getActive().toast('Genererer kravforslag...');

    const analysis   = performComprehensiveAnalysis_();
    const existing   = readRequirementsForGapAnalysis_();
    const candidates = generateRequirementCandidates_(analysis);
    const deduped    = dedupeCandidates_(candidates, existing, DISCOVERY_REPORT.CONFIG.DEDUPE_JACCARD);

    if (deduped.length === 0) {
      SpreadsheetApp.getActive().toast(DISCOVERY_REPORT.MESSAGES.SUGGEST_NONE);
      return;
    }

    const nextIdStart = nextKravId_(existing, sh);
    const rows = [];
    let seq = 0;

    for (let i = 0; i < deduped.length && i < DISCOVERY_REPORT.CONFIG.MAX_SUGGESTIONS_IN_DOC; i++) {
      const c = deduped[i];
      const title = guessShortTitle_(c.text) || c.text;
      const kravId = formatKravId_(nextIdStart + (seq++));
      const priority = guessPriority_(c);
      const intent = generateIntentGoalText_(c.text || '', title);

      const srcType = String(c.source || '');
      const srcDetail =
        srcType === 'funksjon' ? (c.extra && c.extra.name) || '' :
        srcType === 'trigger'  ? ((c.extra && c.extra.event ? c.extra.event + ' → ' : '') + (c.extra && c.extra.handler || '')) :
        srcType === 'ark'      ? (c.extra && c.extra.sheet) || '' :
        srcType === 'felt'     ? ((c.extra && c.extra.sheet ? c.extra.sheet + ' · ' : '') + (c.extra && c.extra.field || '')) : '';
      const srcLink = (c.extra && c.extra.link) || '';

      rows.push([
        kravId, title, c.text, priority, intent, '',
        srcType, srcDetail, srcLink,
        '', '', '', '', ''
      ]);
    }

    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    SpreadsheetApp.getActive().toast(
      DISCOVERY_REPORT.MESSAGES.SUGGEST_OK.replace('{n}', String(rows.length))
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Feil ved generering av forslag', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/* ------------------------------- Analyse-kjerne ---------------------------- */
function performComprehensiveAnalysis_(startTime) {
  const ss = SpreadsheetApp.getActive();

  SpreadsheetApp.getActive().toast('Analyserer triggere...');
  const triggers = analyzeTriggers_();

  const functions = analyzeFunctions_(); // viser egen teller

  const sheets = analyzeSheets_(DISCOVERY_REPORT.CONFIG.BATCH_SIZE, startTime);

  return {
    metadata: {
      timestamp: new Date(),
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl().replace(/#.*$/, ''),
      user: Session.getActiveUser().getEmail(),
      tz: Session.getScriptTimeZone() || 'UTC',
      analysisTimeMs: Date.now() - startTime,
      version: DISCOVERY_REPORT.VERSION
    },
    triggers, functions, sheets
  };
}

function analyzeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  SpreadsheetApp.getActive().toast(
    DISCOVERY_REPORT.MESSAGES.PROGRESS_TRIGGERS.replace('{count}', triggers.length)
  );
  return {
    count: triggers.length,
    details: triggers.map(t => ({
      handler: t.getHandlerFunction(),
      eventType: String(t.getEventType && t.getEventType() || ''),
      source: String(t.getTriggerSource && t.getTriggerSource() || '')
    }))
  };
}

function analyzeFunctions_() {
  const globalFuncs = [], privateFuncs = [];
  Object.keys(globalThis).forEach(key => {
    if (typeof globalThis[key] === 'function') {
      const info = { name: key, type: 'Function' };
      if (key.endsWith('_')) privateFuncs.push(info);
      else globalFuncs.push(info);
    }
  });
  const total = globalFuncs.length + privateFuncs.length;
  SpreadsheetApp.getActive().toast(
    DISCOVERY_REPORT.MESSAGES.PROGRESS_FUNCTIONS.replace('{count}', total)
  );
  return { global: globalFuncs, private: privateFuncs };
}

function analyzeSheets_(batchSize, startTime) {
  const ss = SpreadsheetApp.getActive();
  const baseUrl = ss.getUrl().replace(/#.*$/, '');
  const allSheets = ss.getSheets();
  const safeBatch = Math.max(1, Number(batchSize || 15));
  const totalBatches = Math.ceil(allSheets.length / safeBatch);
  const allDetails = [];
  let processedCount = 0;

  for (let i = 0; i < allSheets.length; i += safeBatch) {
    if (startTime && (Date.now() - startTime > DISCOVERY_REPORT.CONFIG.MAX_EXECUTION_TIME)) {
      Logger.log('Sheet analysis stopped due to timeout');
      break;
    }
    const currentBatch = Math.floor(i / safeBatch) + 1;
    SpreadsheetApp.getActive().toast(
      DISCOVERY_REPORT.MESSAGES.PROGRESS_SHEETS
        .replace('{current}', currentBatch)
        .replace('{total}', totalBatches)
    );

    const batch = allSheets.slice(i, i + safeBatch);
    const batchDetails = batch.map(s => {
      const validation = validateSheetStructure_(s);
      if (!validation.valid) {
        Logger.log(`Skipping sheet "${s.getName()}": ${validation.reason}`);
        return null;
      }
      processedCount++;
      const gid = s.getSheetId();
      const url = baseUrl + '#gid=' + gid;
      return {
        name: s.getName(),
        gid: gid,
        url: url,
        rows: s.getLastRow(),
        columns: s.getLastColumn(),
        isHidden: s.isSheetHidden(),
        headerPreview: getHeaderPreview_(s),
        warnings: validation.warnings || []
      };
    }).filter(Boolean);

    allDetails.push(...batchDetails);
    if (safeBatch > 10 && currentBatch < totalBatches) Utilities.sleep(80);
  }

  return {
    count: allDetails.length,
    sheets: allDetails,
    processedCount,
    totalSheets: allSheets.length,
    skippedSheets: allSheets.length - processedCount
  };
}

function validateSheetStructure_(sheet) {
  const validation = { valid: true, warnings: [] };
  if (!sheet || sheet.getLastRow() === 0) return { valid: false, reason: 'Empty sheet' };

  const rows = sheet.getLastRow();
  const cols = sheet.getLastColumn();

  if (cols > DISCOVERY_REPORT.CONFIG.MAX_SHEET_COLUMNS) {
    return { valid: false, reason: `Too many columns (${cols} > ${DISCOVERY_REPORT.CONFIG.MAX_SHEET_COLUMNS})` };
  }
  if (rows > DISCOVERY_REPORT.CONFIG.LARGE_SHEET_WARNING_THRESHOLD) {
    validation.warnings.push(`Large sheet: ${rows} rows`);
    Logger.log(`Large sheet detected: ${sheet.getName()} (${rows} rows)`);
  }
  if (cols > 50) validation.warnings.push(`Many columns: ${cols}`);

  return validation;
}

function getHeaderPreview_(sheet) {
  try {
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const headerRange = sheet.getRange(1, 1, 1, Math.min(10, sheet.getLastColumn()));
      return headerRange.getValues()[0]
        .map(v => String(v || '').substring(0, 30))
        .join(' | ');
    }
  } catch (e) {
    Logger.log(`Could not get header preview for ${sheet.getName()}: ${e.message}`);
  }
  return '';
}


/* -------------------------- Lesing av "Krav"-arket ------------------------- */
function readRequirementsForGapAnalysis_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(DISCOVERY_REPORT.CONFIG.KRAV_SHEET);
  if (!sh) return [];

  try {
    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return [];

    const hdr = vals[0].map(h => String(h || '').toLowerCase());
    const idx = {
      id: idxOfAny_(hdr, ['id', 'kravid']),
      title: idxOfAny_(hdr, ['tittel', 'krav', 'kravtekst', 'tekst']),
      priority: idxOfAny_(hdr, ['prioritet', 'prio']),
      p1: idxOfAny_(hdr, ['fremdrift 1','fremdrift1']),
      p2: idxOfAny_(hdr, ['fremdrift 2','fremdrift2']),
      p3: idxOfAny_(hdr, ['fremdrift 3','fremdrift3']),
      p4: idxOfAny_(hdr, ['fremdrift 4','fremdrift4']),
      p5: idxOfAny_(hdr, ['fremdrift 5','fremdrift5'])
    };

    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const id = safeCell_(row, idx.id);
      const text = safeCell_(row, idx.title);
      if (!(id || text)) continue;

      const checks = [idx.p1, idx.p2, idx.p3, idx.p4, idx.p5]
        .map(c => normalizeCheck_(safeCell_(row, c)))
        .filter(c => c !== null);

      const done = checks.filter(Boolean).length;
      const progressPct = checks.length > 0 ? Math.round((done / checks.length) * 100) : 0;

      out.push({
        id: id,
        text: text,
        priority: safeCell_(row, idx.priority),
        progressPct: progressPct
      });
    }
    return out;
  } catch (e) {
    Logger.log('Error reading requirements: ' + e.message);
    return [];
  }
}

function idxOfAny_(hdr, names) {
  for (const name of names) {
    const j = hdr.indexOf(name);
    if (j >= 0) return j;
  }
  return -1;
}

function normalizeCheck_(v) {
  if (v === '' || v === null) return null;
  const s = String(v).trim().toLowerCase();
  return (s === 'true' || s === 'x' || s === '[x]' || s === '✓' || s === '1');
}

function safeCell_(row, i) {
  return (isFinite(i) && i >= 0 && i < row.length) ? row[i] : '';
}


/* --------------------------- Avviksanalyse / Gap --------------------------- */
function performGapAnalysis_(analysis, requirements, newSuggestions) {
  const reqs = requirements || [];
  const unimplemented = reqs.filter(r => (r.progressPct || 0) === 0);

  const allTextLC = reqs.map(r => String(r.text || '').toLowerCase()).join(' ');
  const undocumented = analysis.functions.global.filter(f => {
    const n = f.name || '';
    const isStandard = /^on(Open|Edit|Install|Submit)|^install|^generate/i.test(n);
    const mentioned = allTextLC.indexOf(n.toLowerCase()) >= 0;
    return !isStandard && !mentioned;
  });

  return {
    unimplementedRequirements: unimplemented,
    undocumentedFunctions: undocumented,
    newSuggestions: newSuggestions || [],
    coverageStats: {
      totalRequirements: reqs.length,
      implementedRequirements: reqs.filter(r => (r.progressPct || 0) > 0).length,
      fullyImplementedRequirements: reqs.filter(r => (r.progressPct || 0) >= 100).length,
      documentedFunctions: analysis.functions.global.length - undocumented.length,
      totalFunctions: analysis.functions.global.length
    }
  };
}


/* ----------------------- Krav-kandidater (auto-forslag) -------------------- */
function generateRequirementCandidates_(analysis) {
  const out = [];
  const allFuncs = []
    .concat(analysis.functions.global || [])
    .concat(analysis.functions.private || []);
  const scriptUrl = getScriptEditorUrl_();

  // Funksjonsnavn → krav
  allFuncs.forEach(f => {
    const text = functionNameToRequirement_(f.name);
    if (text) out.push({
      text, source: 'funksjon', score: 0.9,
      extra: { name: f.name, link: scriptUrl }
    });
  });

  // Triggere → krav
  (analysis.triggers.details || []).forEach(t => {
    const text = triggerToRequirement_(t);
    if (text) out.push({
      text, source: 'trigger', score: 0.95,
      extra: { event: t.eventType, handler: t.handler, link: scriptUrl }
    });
  });

  // Ark/felter → krav
  (analysis.sheets.sheets || []).forEach(s => {
    if (s && s.name) {
      out.push({
        text: `Systemet skal ha et register/ark for «${s.name}».`,
        source: 'ark', score: 0.85,
        extra: { sheet: s.name, link: s.url }
      });
      if (s.headerPreview) {
        const cols = s.headerPreview.split('|').map(x => String(x || '').trim()).filter(Boolean).slice(0, 5);
        cols.forEach(c => {
          if (c.length > 2 && c.length < 50) {
            out.push({
              text: `Datafeltet «${c}» skal finnes og forvaltes i «${s.name}».`,
              source: 'felt', score: 0.8,
              extra: { sheet: s.name, field: c, link: s.url + '&range=1:1' }
            });
          }
        });
      }
    }
  });

  return out.sort((a, b) => (b.score || 0) - (a.score || 0));
}

function functionNameToRequirement_(name) {
  if (!name) return '';
  if (/^onOpen|^onEdit|^doGet|^doPost|^include/i.test(name)) return '';

  const pretty = name
    .replace(/^open/i,    'åpne ')
    .replace(/^get/i,     'hente ')
    .replace(/^set/i,     'sette ')
    .replace(/^save/i,    'lagre ')
    .replace(/^create/i,  'opprette ')
    .replace(/^update/i,  'oppdatere ')
    .replace(/^delete/i,  'slette ')
    .replace(/^send/i,    'sende ')
    .replace(/^process/i, 'prosessere ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!pretty || pretty.length < 3) return '';
  return `Systemet skal kunne ${pretty.toLowerCase()}.`;
}

function triggerToRequirement_(t) {
  if (!t || !t.handler) return '';
  const evt = String(t.eventType || '').toUpperCase();

  if (evt.indexOf('ON_FORM_SUBMIT') >= 0) return `Ved innsending av skjema skal systemet prosessere via «${t.handler}».`;
  if (evt.indexOf('ON_EDIT') >= 0)        return `Ved endring i ark skal systemet prosessere via «${t.handler}».`;
  if (evt.indexOf('CLOCK') >= 0)          return `Systemet skal periodisk kjøre «${t.handler}» (tidsstyrt).`;
  if (evt.indexOf('ON_OPEN') >= 0)        return `Ved åpning av regnearket skal systemet kjøre «${t.handler}».`;

  return `Systemet skal støtte hendelsen «${t.eventType}» via «${t.handler}».`;
}

function dedupeCandidates_(candidates, existingReqs, jaccardThreshold) {
  const maxN = DISCOVERY_REPORT.CONFIG.MAX_SUGGESTIONS_IN_DOC;
  const existingTexts = (existingReqs || []).map(r => normalizeText_(r.text || ''));
  const out = [];
  const added = new Set();

  for (const c of (candidates || [])) {
    const norm = normalizeText_(c.text || '');
    if (!norm) continue;

    let skip = false;
    for (const ex of existingTexts) {
      if (jaccard_(tokenize_(norm), tokenize_(ex)) >= (jaccardThreshold || 0.78)) { skip = true; break; }
    }
    if (skip || added.has(norm)) continue;
    added.add(norm);
    out.push(c);
    if (out.length >= maxN) break;
  }
  return out;
}


/* ---------------------- Dokumentbygger (Google Document) ------------------- */
function buildDiscoveryReportDocument_(body, analysis, existingReqs, newSuggestions, gap, timeoutOccurred) {
  body.appendParagraph(`Teknisk Analyse & Kravsrapport: "${analysis.metadata.spreadsheetName}"`)
      .setHeading(DocumentApp.ParagraphHeading.TITLE);

  const meta = analysis.metadata;
  body.appendParagraph(`Rapport generert: ${Utilities.formatDate(meta.timestamp, meta.tz, 'dd.MM.yyyy HH:mm:ss')} (v${meta.version})`);
  body.appendParagraph(`Analysetid: ${Math.round(meta.analysisTimeMs / 1000)} sekunder`);
  body.appendParagraph(`URL: ${analysis.metadata.spreadsheetUrl}`);
  body.appendParagraph(`Bruker: ${meta.user}`);

  if (timeoutOccurred) {
    body.appendParagraph('⚠️ ADVARSEL: Analysen ble avbrutt pga. tidsbegrensning. Rapporten kan være ufullstendig.')
        .editAsText()
        .setForegroundColor('#D93025')
        .setBold(true);
  }
  body.appendParagraph('');

  buildRequirementsChapter_(body, analysis, existingReqs || [], newSuggestions || []);
  buildGapAnalysisSection_(body, gap);
  buildDependencySection_(body, analysis);
  buildSummaryAndTriggersSection_(body, analysis);
  buildFunctionsSection_(body, analysis.functions);
  buildSheetsSection_(body, analysis.sheets);
  buildManualAnalysisSection_(body);
  buildUsageAppendix_(body);
}

function buildRequirementsChapter_(body, analysis, requirements, newSuggestions) {
  body.appendParagraph('1. Krav (Eksisterende + Forslag)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const count = requirements.length;
  const done  = requirements.filter(r => (r.progressPct || 0) >= 100).length;
  const zero  = requirements.filter(r => (r.progressPct || 0) === 0).length;
  body.appendParagraph(`Antall eksisterende krav: ${count} (fullført: ${done}, 0%: ${zero})`);

  body.appendParagraph('1.1 Eksisterende krav')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (count === 0) {
    body.appendParagraph('Ingen krav funnet i arket «Krav».');
  } else {
    const table = body.appendTable([['ID', 'Krav', 'Prio.', '%', 'Hensikt & mål (manuelt)', 'Verifikasjon / Test (manuelt)']]);
    table.getRow(0).editAsText().setBold(true);
    requirements.forEach(r => {
      const row = table.appendTableRow();
      row.appendTableCell(String(r.id || ''));
      row.appendTableCell(String(r.text || ''));
      row.appendTableCell(String(r.priority || ''));
      row.appendTableCell(isFinite(r.progressPct) ? (r.progressPct + '%') : '');
      row.appendTableCell('');
      row.appendTableCell('');
    });
  }

  body.appendParagraph('');
  body.appendParagraph('1.2 Nye kravforslag (auto-oppdaget)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (!newSuggestions || newSuggestions.length === 0) {
    body.appendParagraph('Ingen nye forslag denne gangen.');
  } else {
    const t2 = body.appendTable([['Foreslått krav', 'Prio. (auto)', 'Kilde', 'Kilde-detalj', 'Kilde-lenke', 'Hensikt & mål (auto)', 'Verifikasjon / Test (manuelt)']]);
    t2.getRow(0).editAsText().setBold(true);
    newSuggestions.forEach(c => {
      const row = t2.appendTableRow();
      row.appendTableCell(c.text || '');
      row.appendTableCell(guessPriority_(c));
      row.appendTableCell(String(c.source || 'auto'));

      const detail =
        c.source === 'funksjon' ? (c.extra && c.extra.name) || '' :
        c.source === 'trigger'  ? ((c.extra && c.extra.event ? c.extra.event + ' → ' : '') + (c.extra && c.extra.handler || '')) :
        c.source === 'ark'      ? (c.extra && c.extra.sheet) || '' :
        c.source === 'felt'     ? ((c.extra && c.extra.sheet ? c.extra.sheet + ' · ' : '') + (c.extra && c.extra.field || '')) : '';
      row.appendTableCell(detail);

      const linkCell = row.appendTableCell(c.extra && c.extra.link ? 'Åpne' : '');
      if (c.extra && c.extra.link) linkCell.editAsText().setLinkUrl(c.extra.link);

      row.appendTableCell(generateIntentGoalText_(c.text || '', guessShortTitle_(c.text || '')));
      row.appendTableCell('');
    });
  }
  body.appendParagraph('');
  body.appendHorizontalRule();
}

function buildGapAnalysisSection_(body, gap) {
  body.appendParagraph('2. Avviksanalyse (Kode vs. Krav)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('📝 Uimplementerte krav (0% fremdrift)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const miss = (gap && gap.unimplementedRequirements) || [];
  if (miss.length === 0) {
    body.appendParagraph('Ingen krav med 0% fremdrift.');
  } else {
    const table = body.appendTable([['KravID','Prioritet','Kravtekst']]);
    table.getRow(0).editAsText().setBold(true);
    miss.forEach(r => {
      const row = table.appendTableRow();
      row.appendTableCell(String(r.id || ''));
      row.appendTableCell(String(r.priority || ''));
      row.appendTableCell(String(r.text || ''));
    });
  }
  body.appendParagraph('');

  body.appendParagraph('🚨 Udokumenterte funksjoner')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const undoc = (gap && gap.undocumentedFunctions) || [];
  if (undoc.length === 0) {
    body.appendParagraph('Ingen åpenbart udokumenterte funksjoner.');
  } else {
    const t2 = body.appendTable([['Funksjon']]);
    t2.getRow(0).editAsText().setBold(true);
    undoc.forEach(f => t2.appendTableRow().appendTableCell(String(f.name || '')));
  }
  body.appendParagraph('');
  body.appendHorizontalRule();
}

function buildDependencySection_(body, analysis) {
  body.appendParagraph('3. Avhengigheter (avledet)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const table = body.appendTable([['Type','Fra','Til / Detalj','Lenke']]);
  table.getRow(0).editAsText().setBold(true);

  const ideUrl = getScriptEditorUrl_();

  (analysis.triggers.details || []).forEach(t => {
    const row = table.appendTableRow();
    row.appendTableCell('Trigger');
    row.appendTableCell(String(t.eventType || ''));
    row.appendTableCell(String(t.handler || ''));
    const c = row.appendTableCell('Åpne');
    c.editAsText().setLinkUrl(ideUrl);
  });

  (analysis.sheets.sheets || []).forEach(s => {
    if (s.headerPreview) {
      const cols = s.headerPreview.split('|').map(x => String(x || '').trim()).filter(Boolean).slice(0, 5);
      cols.forEach(cn => {
        const r = table.appendTableRow();
        r.appendTableCell('Datafelt');
        r.appendTableCell(s.name);
        r.appendTableCell(cn);
        const lc = r.appendTableCell('Åpne');
        lc.editAsText().setLinkUrl(s.url + '&range=1:1');
      });
    }
  });

  body.appendParagraph('');
}

function buildSummaryAndTriggersSection_(body, analysis) {
  body.appendParagraph('4. Teknisk Sammendrag')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const t = body.appendTable([['Aspekt','Antall']]);
  t.getRow(0).editAsText().setBold(true);
  let r = t.appendTableRow(); r.appendTableCell('Ark'); r.appendTableCell(String(analysis.sheets.count));
  r = t.appendTableRow();     r.appendTableCell('Funksjoner (offentlig + privat)'); r.appendTableCell(String(analysis.functions.global.length + analysis.functions.private.length));
  r = t.appendTableRow();     r.appendTableCell('Automatiske triggere'); r.appendTableCell(String(analysis.triggers.count));
  body.appendParagraph('');

  body.appendParagraph('Automatiske triggere')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (analysis.triggers.count === 0) {
    body.appendParagraph('Ingen automatiske triggere funnet.');
  } else {
    const table = body.appendTable([['Hendelse','Kilde','Funksjon','Lenke']]);
    table.getRow(0).editAsText().setBold(true);
    const ideUrl = getScriptEditorUrl_();
    (analysis.triggers.details || []).forEach(tg => {
      const row = table.appendTableRow();
      row.appendTableCell(String(tg.eventType || ''));
      row.appendTableCell(String(tg.source || ''));
      row.appendTableCell(String(tg.handler || ''));
      const c = row.appendTableCell('Åpne');
      c.editAsText().setLinkUrl(ideUrl);
    });
  }
  body.appendParagraph('');
}

function buildFunctionsSection_(body, funcs) {
  body.appendParagraph('5. Funksjoner')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Offentlige og private (suffix «_» indikerer privat helper).');

  const p = body.appendParagraph('Åpne skriptprosjektet (IDE)');
  p.setItalic(true);
  p.editAsText().setLinkUrl(getScriptEditorUrl_());
  body.appendParagraph('');

  const table = body.appendTable([['Navn','Synlighet']]);
  table.getRow(0).editAsText().setBold(true);

  (funcs.global || []).forEach(f => {
    const row = table.appendTableRow();
    row.appendTableCell(String(f.name || ''));
    row.appendTableCell('Offentlig');
  });
  (funcs.private || []).forEach(f => {
    const row = table.appendTableRow();
    row.appendTableCell(String(f.name || ''));
    row.appendTableCell('Privat');
  });
  body.appendParagraph('');
}

function buildSheetsSection_(body, sheets) {
  body.appendParagraph('6. Ark i Regneark')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const table = body.appendTable([['Arknavn','Rader','Kolonner','Skjult?','Header (utdrag)','Lenke']]);
  table.getRow(0).editAsText().setBold(true);

  (sheets.sheets || []).forEach(s => {
    const row = table.appendTableRow();
    row.appendTableCell(String(s.name || ''));
    row.appendTableCell(String(s.rows || 0));
    row.appendTableCell(String(s.columns || 0));
    row.appendTableCell(s.isHidden ? 'Ja' : 'Nei');
    row.appendTableCell(String(s.headerPreview || ''));
    const c = row.appendTableCell('Åpne');
    if (s.url) c.editAsText().setLinkUrl(s.url);
  });
  body.appendParagraph('');
}

function buildManualAnalysisSection_(body) {
  body.appendHorizontalRule();
  body.appendParagraph('7. Manuelt Arbeidsområde')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Bruk tabellen til å beskrive formål/akseptanse for de viktigste funksjonene.');

  const table = body.appendTable([['Funksjon / Trigger','Forretningsformål (med egne ord)']]);
  table.getRow(0).editAsText().setBold(true);
  for (let i = 0; i < 15; i++) {
    const row = table.appendTableRow();
    row.appendTableCell('');
    row.appendTableCell('');
  }
  body.appendParagraph('');
}

function buildUsageAppendix_(body) {
  body.appendParagraph('8. Bruksanvisning (hurtigstart)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('1) Meny «Analyse» → 🔎 Generer Discovery-rapport (Doc). Dokumentet åpnes automatisk.');
  body.appendParagraph('2) Hvis du har et «Krav»-ark: oppdater prioritet og fremdrift.');
  body.appendParagraph('3) Meny «Analyse» → 💡 Foreslå nye krav → «Krav»-arket for å fylle på forslag.');
  body.appendParagraph('4) Lenken til dokumentet lagres i «Konfig» (DISCOVERY_DOC_ID).');
  body.appendParagraph('');
  body.appendParagraph('Om funksjons-oppdagelse (reverse engineering)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Funksjons-oppdagelse kartlegger hva systemet faktisk gjør – via kode, triggere og datamodell – og oversetter det til eksplisitte krav. Denne rapporten foreslår krav automatisk, slik at du kan dokumentere først og prioritere utvikling etterpå.');
  body.appendParagraph('');
  body.appendParagraph('Begrensninger')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('• Heuristikk: krav-kandidater baseres på navn/strukturer – kan gi falske positive.\n• Apps Script gir ikke direkte lenke til spesifikk funksjon i IDE, men «Åpne»-lenken tar deg til prosjektet.\n• Fremdriftsberegning forutsetter 5 fremdriftskolonner i «Krav» (kan tilpasses).');
}


/* ----------------------------- Hjelpefunksjoner ---------------------------- */
function ensureKravSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DISCOVERY_REPORT.CONFIG.KRAV_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DISCOVERY_REPORT.CONFIG.KRAV_SHEET);
    sh.getRange(1, 1, 1, DISCOVERY_REPORT.CONFIG.KRAV_HEADERS.length)
      .setValues([DISCOVERY_REPORT.CONFIG.KRAV_HEADERS]);
  } else {
    const need = DISCOVERY_REPORT.CONFIG.KRAV_HEADERS;
    const have = (sh.getLastRow() > 0 && sh.getLastColumn() > 0)
      ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), need.length)).getValues()[0]
      : [];
    if (!have || !have[0]) {
      sh.getRange(1, 1, 1, need.length).setValues([need]);
    } else if (have.length < need.length) {
      const newRow = need.map((h, i) => have[i] || h);
      sh.getRange(1, 1, 1, need.length).setValues([newRow]);
    }
  }
  return sh;
}

function nextKravId_(existingReqs, sh) {
  let maxNum = 0;
  const all = (existingReqs || []).map(r => String(r.id || ''));
  try {
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) all.push(String(vals[i][0] || ''));
  } catch (_) {}
  all.forEach(id => {
    const m = id.match(/(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  });
  return maxNum + 1;
}

function formatKravId_(n) {
  const s = String(n);
  return 'KR-' + s.padStart(4, '0');
}

function getOrCreateDiscoveryDocId_(forceNew) {
  const ss = SpreadsheetApp.getActive();
  const key = DISCOVERY_REPORT.CONFIG.DISCOVERY_DOC_KEY;

  if (!forceNew) {
    const existing = upsertKonfigLocal_(key); // get
    if (existing) {
      try { DriveApp.getFileById(existing); return existing; } catch (_) { /* fall-through */ }
    }
  }

  const doc = DocumentApp.create('Analyse & Kravsrapport – ' + ss.getName());
  const newId = doc.getId();
  upsertKonfigLocal_(key, newId, 'Discovery Google Doc ID');
  return newId;
}

function upsertKonfigLocal_(key, value, desc) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = DISCOVERY_REPORT.CONFIG.KONFIG_SHEET;
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, 3).setValues([['Nøkkel','Verdi','Beskrivelse']]);
  }
  const last = sh.getLastRow();
  if (last <= 1) {
    if (value === undefined) return null;
    sh.appendRow([key, value, desc || '']);
    return value;
  }
  const keys = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || ''));
  let row = -1;
  for (let i = 0; i < keys.length; i++) if (keys[i] === key) { row = i + 2; break; }

  if (value === undefined) return (row > -1) ? sh.getRange(row, 2).getValue() : null;

  if (row > -1) {
    sh.getRange(row, 2).setValue(value);
    if (desc) sh.getRange(row, 3).setValue(desc);
  } else {
    sh.appendRow([key, value, desc || '']);
  }
  return value;
}

function openInNewTab_(url) {
  const clean = String(url || '').replace(/[<>"']/g, '');
  const html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial;padding:10px">' +
    'Åpner dokumentet…<br>' +
    '<a target="_blank" href="' + clean + '">Klikk her hvis det ikke åpner</a>' +
    '<script>window.open("' + clean + '","_blank");google.script.host.close();</script>' +
    '</body></html>'
  ).setWidth(300).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Åpner dokument…');
}


/* ---------- Tekstheuristikk (prioritet, hensikt) + dedupe utils ----------- */
function formatSourceDetails_(candidate) {
  if (!candidate || !candidate.source || !candidate.extra) return (candidate && candidate.source) || 'auto';
  switch (candidate.source) {
    case 'funksjon': return `Funksjon: ${candidate.extra.name}`;
    case 'trigger':  return `Trigger: ${(candidate.extra.event || '').toString().toUpperCase()} → ${candidate.extra.handler || ''}`;
    case 'ark':      return `Ark: ${candidate.extra.sheet}`;
    case 'felt':     return `Felt: ${candidate.extra.field} (i ark: ${candidate.extra.sheet})`;
    default:         return candidate.source;
  }
}

function guessPriority_(candidate) {
  const s = String(candidate.text || '').toLowerCase();
  if (/trigger|on_open|on edit|on form submit|tidsstyrt|periodisk/.test(s)) return 'MÅ';
  if (/tilgang|rolle|permission|rbac|sikker/.test(s)) return 'MÅ';
  if (/hms|varsle|avvik|forfall|kritisk/.test(s)) return 'MÅ';
  if (/budsjett|faktura|regnskap|økonomi/.test(s)) return 'BØR';
  if (/rapport|eksport|csv|analyse/.test(s)) return 'KAN';
  return DISCOVERY_REPORT.CONFIG.DEFAULT_PRIORITY;
}

function guessShortTitle_(txt) {
  const s = String(txt || '').trim();
  if (!s) return '';
  const m = s.match(/^(.{3,80}?)([.:;]|$)/);
  return (m && m[1]) ? m[1].trim() : s.slice(0, 80);
}

function generateIntentGoalText_(rawText, shortTitle) {
  const text = String(rawText || '').trim();
  const title = String(shortTitle || '').trim();
  const basis = title || text;
  if (!basis) return 'Hensikt: beskrive, kvalitetssikre og tydeliggjøre funksjonen.\nMål: redusere manuelt arbeid og sikre etterlevelse.';

  const low = basis.toLowerCase();
  const verb =
    /varsle|notifier|notify/.test(low) ? 'varsle' :
    /synk|sync|synkronis/.test(low) ? 'synkronisere' :
    /registrer|lagre|save|oppdat/.test(low) ? 'registrere og oppdatere' :
    /beregn|kalkuler|calculate/.test(low) ? 'beregne' :
    /rapport|export|eksport|report/.test(low) ? 'rapportere' :
    /valider|sjekk|check|validate/.test(low) ? 'validere' :
    /tilgang|rolle|autoriser|permission|rbac/.test(low) ? 'styre tilgang' :
    /plan|årshjul|kalender/.test(low) ? 'planlegge' :
    /oppgave|task/.test(low) ? 'håndtere oppgaver' : 'støtte';

  const goal =
    /hms/.test(low) ? 'sikre etterlevelse av HMS-rutiner' :
    /budsjett|regnskap|faktura/.test(low) ? 'forbedre økonomikontroll og sporbarhet' :
    /møte|protokoll|avstem|vote|protokoll/.test(low) ? 'effektivisere møtearenaer og beslutninger' :
    /kalender|årshjul/.test(low) ? 'gi forutsigbar gjennomføring i året' :
    /oppgave|task/.test(low) ? 'redusere manuelt arbeid og sikre gjennomføring' :
    /beboer|seksjon|eier|leie|person/.test(low) ? 'ha oppdatert og korrekt grunnlagsdata' :
    'redusere manuell oppfølging og sikre kvalitet';

  return (
    'Hensikt: ' + capitalizeNo_('å ' + verb + ' ' + basis + '.').replace(/\.\.$/, '.') +
    '\nMål: ' + capitalizeNo_(goal + '.')
  );
}

function capitalizeNo_(s) { s = String(s || ''); return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

function normalizeText_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9æøåäöáéíóúýüñç\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokenize_(s) {
  return normalizeText_(s).split(/\s+/).filter(Boolean);
}

function jaccard_(aTokens, bTokens) {
  const A = new Set(aTokens);
  const B = new Set(bTokens);
  const inter = [...A].filter(x => B.has(x)).length;
  const uni = new Set([...A, ...B]).size;
  return uni ? inter / uni : 0;
}


/* ----------------------------- Konfig / Utils ------------------------------ */
function loadConfiguration_() {
  // Les fra «Konfig» (ingen skriving/sideeffekt her). Fallback til defaults.
  const defaults = {
    DEDUPE_JACCARD: 0.78,
    MAX_SUGGESTIONS_IN_DOC: 250,
    BATCH_SIZE: 15,
    MAX_EXECUTION_TIME: 330000,               // ~5.5 min
    MAX_SHEET_COLUMNS: 200,
    LARGE_SHEET_WARNING_THRESHOLD: 5000,
    LARGE_FUNCS_THRESHOLD: 200,
    LARGE_SHEETS_THRESHOLD: 40
  };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Konfig');
    if (!sh || sh.getLastRow() < 2) return defaults;
    const data = sh.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim();
      const v = data[i][1];
      if (!k) continue;
      map[k] = v;
    }
    return {
      DEDUPE_JACCARD: Number(map.DEDUPE_JACCARD || defaults.DEDUPE_JACCARD),
      MAX_SUGGESTIONS_IN_DOC: Math.max(10, Number(map.MAX_SUGGESTIONS_IN_DOC || defaults.MAX_SUGGESTIONS_IN_DOC)),
      BATCH_SIZE: Math.max(1, Number(map.BATCH_SIZE || defaults.BATCH_SIZE)),
      MAX_EXECUTION_TIME: Math.max(60000, Number(map.MAX_EXECUTION_TIME || defaults.MAX_EXECUTION_TIME)),
      MAX_SHEET_COLUMNS: Math.max(20, Number(map.MAX_SHEET_COLUMNS || defaults.MAX_SHEET_COLUMNS)),
      LARGE_SHEET_WARNING_THRESHOLD: Math.max(1000, Number(map.LARGE_SHEET_WARNING_THRESHOLD || defaults.LARGE_SHEET_WARNING_THRESHOLD)),
      LARGE_FUNCS_THRESHOLD: Math.max(50, Number(map.LARGE_FUNCS_THRESHOLD || defaults.LARGE_FUNCS_THRESHOLD)),
      LARGE_SHEETS_THRESHOLD: Math.max(10, Number(map.LARGE_SHEETS_THRESHOLD || defaults.LARGE_SHEETS_THRESHOLD))
    };
  } catch (e) {
    Logger.log('loadConfiguration_ failed; using defaults: ' + e.message);
    return defaults;
  }
}

function assessProjectSize_() {
  try {
    const sheets = SpreadsheetApp.getActive().getSheets().length;
    let totalFuncs = 0;
    Object.keys(globalThis).forEach(k => { if (typeof globalThis[k] === 'function') totalFuncs++; });
    const isLarge =
      sheets >= DISCOVERY_REPORT.CONFIG.LARGE_SHEETS_THRESHOLD ||
      totalFuncs >= DISCOVERY_REPORT.CONFIG.LARGE_FUNCS_THRESHOLD;
    return { isLarge, sheets, functions: totalFuncs };
  } catch (e) {
    return { isLarge: false, sheets: 0, functions: 0 };
  }
}

function handleError_(e) {
  const msg = (e && e.message) ? e.message : String(e);
  if (/You do not have permission|Cannot call DocumentApp|Service unavailable/i.test(msg)) {
    return DISCOVERY_REPORT.MESSAGES.PERMISSION_ERROR + '\n\n' + msg;
  }
  if (/Service invoked too many times/i.test(msg)) {
    return 'Tjenesten er kalt for mange ganger på kort tid. Vent litt og prøv igjen.\n\n' + msg;
  }
  return msg;
}

function getScriptEditorUrl_() {
  return 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit';
}

// Valgfritt: slå på menyen direkte ved åpning
// function onOpen() { discoveryRegisterMenu_(); }
