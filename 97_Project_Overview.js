/* =============================================================================
 * Project Overview / Scorecard – DIAG + Google Doc + Menu + Triggers
 * FILE: 97_Project_Overview.gs
 * VERSION: 2.2.0
 * UPDATED: 2025-09-15
 * PURPOSE:
 *  - Skanne kjerneark og skrive en lettlest oversikt til DIAG_PROJECT
 *  - Eksportere samme status til et Google-dokument (lenkes i DIAG_PROJECT)
 *  - Legge "Prosjekt"-meny (Oppdater, Eksporter, Åpne dokument, Installer ukentlig)
 *  - Installere auto-triggere (daglig/ukentlig/onOpen) – idempotent
 *
 * INKLUDERER PATCHER:
 *  - (Kritisk) Doc-tabell: ikke bruk appendTableRow([]) med arrays
 *  - (Kritisk) Lenke: bruk RichText i stedet for HYPERLINK()-formel (#ERROR!)
 *  - (Liten) Auto-åpne Google Doc etter eksport
 *  - (Liten) Prosjekt-meny via engangsknapp + installérbar onOpen-trigger
 * ============================================================================ */

/* -------------------------- Namespace & Konfigurasjon ----------------------- */
(function(glob){
  const CONFIG = {
    DIAG_SHEET: 'DIAG_PROJECT',
    KONFIG_SHEET: 'Konfig',
    KONFIG_KEY_SCORECARD_DOC: 'SCORECARD_DOC_ID',
    MENU_NAME: 'Prosjekt',
    CORE_DEFS: [
      { label:'HMS_PLAN (plan)',                 names:['HMS_PLAN'] },
      { label:'TASKS / Oppgaver (oppgaver)',     names:['Oppgaver','TASKS'] },
      { label:'TILGANG (roller)',                names:['TILGANG'] },
      { label:'BEBOERE (beboerregister)',        names:['BEBOERE','Beboere'] },
      { label:'LEVERANDØRER (leverandører)',     names:['LEVERANDØRER'] }
    ],
    DEFAULTS: { TRIGGER_HOUR: 8, PREVIEW_MAX_COLS: 10 }
  };

  const NS = glob.PROJ_OVERVIEW || {};
  NS.VERSION = '2.2.0';
  NS.UPDATED = '2025-09-15';
  NS.CONFIG = CONFIG;
  glob.PROJ_OVERVIEW = NS;
})(globalThis);

/* ------------------------------- Utilities ---------------------------------- */
const Validators = {
  isValidHour: (h) => Number.isInteger(Number(h)) && Number(h) >= 0 && Number(h) <= 23
};

const Utils = {
  withRetry(fn, attempts = 3, delayMs = 800){
    for (let i=1;i<=attempts;i++){
      try { return fn(); }
      catch(e){
        if (i===attempts) throw e;
        Utilities.sleep(delayMs * i);
      }
    }
  },
  logError(op, e, ctx){
    try{ Logger.log('['+op+'] '+(e && e.message||e)+' | ctx='+JSON.stringify(ctx||{})); }catch(_){}
  },
  openInNewTab(url){
    const html = HtmlService.createHtmlOutput(
      '<html><body style="font-family:Arial;padding:10px">Åpner dokumentet… '+
      '<a target="_blank" href="'+url+'">Klikk her hvis det ikke åpner</a>'+
      '<script>window.open("'+url+'","_blank");google.script.host.close();</script>'+
      '</body></html>'
    ).setWidth(320).setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(html, 'Scorecard');
  }
};

/* -------------------------------- Meny (UI) -------------------------------- */
/** Bygg Prosjekt-menyen nå (engangsknapp). */
function projectMenuBuildQuick(){
  try { registerProjectMenu_(); SpreadsheetApp.getActive().toast('Prosjekt-meny lagt til.'); }
  catch(e){ Utils.logError('projectMenuBuildQuick', e); }
}

/** Installer idempotent onOpen-trigger som bygger menyen. */
function installProjectMenuOnOpenTrigger(){
  const ssId = SpreadsheetApp.getActive().getId();
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction()==='registerProjectMenu_' &&
        t.getTriggerSource()===ScriptApp.TriggerSource.SPREADSHEETS){
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('registerProjectMenu_').forSpreadsheet(ssId).onOpen().create();
  SpreadsheetApp.getActive().toast('Prosjekt-meny koblet til onOpen. Last arket på nytt.');
}

/** Selve menybyggeren. */
function registerProjectMenu_(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(PROJ_OVERVIEW.CONFIG.MENU_NAME)
    .addItem('🚀 Kjør ALT nå (setup→oversikt→åpne doc)', 'projectRunAllQuick')
    .addSeparator()
    .addItem('Oppdater prosjektoversikt', 'projectOverviewQuick')
    .addItem('Eksporter scorecard → Google Doc', 'projectScorecardToDocQuick')
    .addItem('Åpne scorecard-dokument…', 'openScorecardDoc')
    .addSeparator()
    .addSubMenu(ui.createMenu('Installer triggere')
      .addItem('Daglig 08:00 (Kjør ALT)', 'installRunAllDailyTrigger')
      .addItem('Ved åpning (Kjør ALT)', 'installRunAllOnOpenTrigger')
      .addItem('Ukentlig status (oversikt)', 'installProjectOverviewTrigger'))
    .addSubMenu(ui.createMenu('Fjern triggere')
      .addItem('Daglig (Kjør ALT)', 'uninstallRunAllDailyTrigger')
      .addItem('Ved åpning (Kjør ALT)', 'uninstallRunAllOnOpenTrigger')
      .addItem('Ukentlig status', 'uninstallProjectOverviewTrigger'))
    .addToUi();
}

/* ---------------------------- Hoved-kommandoer ----------------------------- */
/** Orkestrerer: setupWorkbook → projectOverviewQuick → openScorecardDoc. */
function projectRunAllQuick(){
  const ss = SpreadsheetApp.getActive();
  const res = {setup:false, overview:false, open:false};
  try { if (typeof setupWorkbook==='function'){ setupWorkbook(); res.setup=true; } } catch(e){ Utils.logError('setupWorkbook', e); }
  let summary='';
  try { summary=projectOverviewQuick(); res.overview=true; } catch(e){ Utils.logError('projectOverviewQuick', e); }
  try { openScorecardDoc(); res.open=true; } catch(e){ Utils.logError('openScorecardDoc', e); }
  ss.toast('Kjør alt ferdig. ('+Object.values(res).filter(Boolean).length+'/3) '+(summary||''));
  return {summary:summary||'OK', res};
}

/** Oppdater DIAG + legg scorecard-lenke. */
function projectOverviewQuick(){
  const scan = projectScan_();
  projectOverviewWrite_(scan);
  projectOverviewAddScorecardUrl_();
  return scan.summary;
}

/** Eksporter til Doc og åpne automatisk. */
function projectScorecardToDocQuick(){
  const out = Utils.withRetry(()=>projectExportScorecardToDoc_());
  Utils.openInNewTab(out.url);
  SpreadsheetApp.getActive().toast('Scorecard eksportert og åpnet.');
  return out;
}

/** Åpne scorecard-dokument (opprett hvis mangler). */
function openScorecardDoc(){
  const id = Utils.withRetry(()=>getOrCreateScorecardDocId_());
  const url = 'https://docs.google.com/document/d/'+id+'/edit';
  Utils.openInNewTab(url);
  return {ok:true, docId:id, url:url};
}

/* ------------------------------ Trigger-styring ---------------------------- */
function installRunAllDailyTrigger(hour){
  const h = Validators.isValidHour(Number(hour)) ? Number(hour) : PROJ_OVERVIEW.CONFIG.DEFAULTS.TRIGGER_HOUR;
  _uninstallTrigger({handler:'projectRunAllQuick', source:'CLOCK'});
  ScriptApp.newTrigger('projectRunAllQuick').timeBased().everyDays(1).atHour(h).create();
  SpreadsheetApp.getActive().toast('Daglig "Kjør alt" installert kl. '+String(h).padStart(2,'0')+':00.');
}
function uninstallRunAllDailyTrigger(){
  _uninstallTrigger({handler:'projectRunAllQuick', source:'CLOCK'});
  SpreadsheetApp.getActive().toast('Daglig "Kjør alt"-trigger fjernet.');
}
function installRunAllOnOpenTrigger(){
  _uninstallTrigger({handler:'projectRunAllQuick', source:'SPREADSHEETS'});
  ScriptApp.newTrigger('projectRunAllQuick').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  SpreadsheetApp.getActive().toast('"Kjør alt" ved åpning er aktivert.');
}
function uninstallRunAllOnOpenTrigger(){
  _uninstallTrigger({handler:'projectRunAllQuick', source:'SPREADSHEETS'});
  SpreadsheetApp.getActive().toast('"Kjør alt" ved åpning er deaktivert.');
}
function installProjectOverviewTrigger(){
  _uninstallTrigger({handler:'projectOverviewQuick', source:'CLOCK'});
  ScriptApp.newTrigger('projectOverviewQuick').timeBased().everyWeeks(1).create();
  SpreadsheetApp.getActive().toast('Ukentlig statusoppdatering installert.');
}
function uninstallProjectOverviewTrigger(){
  _uninstallTrigger({handler:'projectOverviewQuick', source:'CLOCK'});
  SpreadsheetApp.getActive().toast('Ukentlig statusoppdatering-trigger fjernet.');
}
function _uninstallTrigger({handler, source}){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction()===handler &&
        (source ? t.getTriggerSource()===ScriptApp.TriggerSource[source] : true)) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/* ------------------------------ Skanning/Status ---------------------------- */
function projectScan_(){
  const ss = SpreadsheetApp.getActive();
  const now = new Date();
  const meta = {
    moduleVersion: PROJ_OVERVIEW.VERSION,
    moduleUpdated: PROJ_OVERVIEW.UPDATED,
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    tz: Session.getScriptTimeZone() || 'UTC',
    user: (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'ukjent')
  };

  const checks = PROJ_OVERVIEW.CONFIG.CORE_DEFS.map(def=>{
    const found = findSheetByCandidates_(def.names);
    return {
      label: def.label,
      status: found ? 'OK' : 'Mangler',
      icon: found ? '✅' : '❌',
      rows: found ? found.getLastRow() : 0,
      header: found ? firstRowPreview_(found, PROJ_OVERVIEW.CONFIG.DEFAULTS.PREVIEW_MAX_COLS) : '',
      actualName: found ? found.getName() : null
    };
  });

  const ok = checks.filter(c=>c.status==='OK').length;
  const fail = checks.filter(c=>c.status==='Mangler').length;
  const progressPct = Math.round((ok / checks.length) * 100);
  const okList = checks.filter(c=>c.status==='OK').map(c=>'OK '+c.label.split(' ')[0]+(c.actualName?'; '+c.actualName:''));
  const summary = 'runAllChecks: OK='+ok+', WARN=0, FAIL='+fail+(okList.length?(' | '+okList.join('; ')):'');
  return { ts:now, meta, checks, ok, warn:0, fail, progressPct, summary };
}

/* ------------------------------ Skriv til DIAG ----------------------------- */
function projectOverviewWrite_(scan){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PROJ_OVERVIEW.CONFIG.DIAG_SHEET) || ss.insertSheet(PROJ_OVERVIEW.CONFIG.DIAG_SHEET);
  sh.clear();

  let r=1;
  const H=(t)=>{ sh.getRange(r,1).setValue(t).setFontWeight('bold'); r++; };
  const KV=(k,v)=>{ sh.getRange(r,1,1,2).setValues([[k,v]]); r++; };
  const BL=()=>{ r++; };

  H('Prosjekt');
  KV('Versjon (modul)', scan.meta.moduleVersion);
  KV('Oppdatert (modul)', scan.meta.moduleUpdated);
  KV('Spreadsheet navn', scan.meta.spreadsheetName);
  KV('Spreadsheet ID', scan.meta.spreadsheetId);
  KV('Spreadsheet URL', scan.meta.spreadsheetUrl);
  KV('Tidssone', scan.meta.tz);
  KV('Aktiv bruker', scan.meta.user);
  KV('Skannet', Utilities.formatDate(scan.ts, scan.meta.tz, 'dd.MM.yyyy HH:mm:ss'));
  BL();

  H('Kjernesjekk');
  sh.getRange(r,1,1,4).setValues([['Navn','Status','Rader','Header (utdrag)']]).setFontWeight('bold'); r++;
  scan.checks.forEach(c=>{
    sh.getRange(r,1,1,4).setValues([[
      c.label + (c.actualName ? '  ('+c.actualName+')' : ''),
      c.icon + ' ' + (c.actualName ? '('+c.actualName+')' : c.status),
      c.rows,
      c.header
    ]]); r++;
  });
  BL();
  KV('Score', scan.ok+' OK, '+scan.warn+' WARN, '+scan.fail+' FAIL');
  KV('Fremdrift', scan.progressPct + '%');

  try { sh.autoResizeColumns(1, 4); } catch(_){}
}

/* ----------------------------- Google Doc eksport -------------------------- */
function projectExportScorecardToDoc_(){
  const scan = projectScan_();
  const docId = getOrCreateScorecardDocId_();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody(); body.clear();

  body.appendParagraph('Sameieportalen – Prosjektstatus / Scorecard').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Oppdatert: ' + Utilities.formatDate(scan.ts, scan.meta.tz, 'dd.MM.yyyy HH:mm'));
  body.appendParagraph('Regneark: ' + scan.meta.spreadsheetName + ' (' + scan.meta.spreadsheetId + ')');
  body.appendParagraph('URL: ' + scan.meta.spreadsheetUrl);
  body.appendParagraph('Modul: v' + scan.meta.moduleVersion + ' – sist oppdatert ' + scan.meta.moduleUpdated);
  body.appendParagraph('Aktiv bruker: ' + scan.meta.user);
  body.appendParagraph('');

  body.appendParagraph('Sammendrag').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Score: ' + scan.ok+' OK, '+scan.warn+' WARN, '+scan.fail+' FAIL');
  body.appendParagraph('Fremdrift: ' + scan.progressPct + '%');
  body.appendParagraph(scan.summary);
  body.appendParagraph('');

  body.appendParagraph('Kjernesjekk').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const table = body.appendTable([['Navn','Status','Rader','Header (utdrag)']]);
  table.getRow(0).editAsText().setBold(true);
  scan.checks.forEach(function(c){
    const row = table.appendTableRow();               // PATCH (kritisk): ikke send array
    row.appendTableCell(c.label + (c.actualName ? '  ('+c.actualName+')' : ''));
    row.appendTableCell(c.icon + ' ' + (c.actualName ? '('+c.actualName+')' : c.status));
    row.appendTableCell(String(c.rows));
    row.appendTableCell(c.header || '');
  });

  doc.saveAndClose();
  const url = doc.getUrl();

  // Hold DIAG oppdatert med lenke (PATCH: RichText, ikke HYPERLINK-formel)
  try { projectOverviewAddScorecardUrl_(); } catch(_){}
  return { ok:true, docId: docId, url: url };
}

/** Hent Doc-ID fra Konfig eller opprett nytt dokument og lagre ID. */
function getOrCreateScorecardDocId_(){
  const ss = SpreadsheetApp.getActive();
  const id = upsertKonfigLocal_(PROJ_OVERVIEW.CONFIG.KONFIG_KEY_SCORECARD_DOC);
  if (id){
    try { DriveApp.getFileById(id); return id; } catch(_){}
  }
  const doc = DocumentApp.create('Sameieportalen – Scorecard ('+ss.getName()+')');
  upsertKonfigLocal_(PROJ_OVERVIEW.CONFIG.KONFIG_KEY_SCORECARD_DOC, doc.getId(), 'Scorecard Google Doc ID');
  return doc.getId();
}

/** PATCH (kritisk): skriv RichText-lenke "Scorecard-dokument" i DIAG. */
function projectOverviewAddScorecardUrl_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PROJ_OVERVIEW.CONFIG.DIAG_SHEET) || ss.insertSheet(PROJ_OVERVIEW.CONFIG.DIAG_SHEET);
  const docId = getOrCreateScorecardDocId_();
  const url = 'https://docs.google.com/document/d/'+docId+'/edit';
  const label = 'Scorecard-dokument';

  const last = Math.max(1, sh.getLastRow());
  let target = 0;
  const colA = sh.getRange(1,1,last,1).getValues().map(r=>String(r[0]||'').trim());
  for (let i=0;i<colA.length;i++){ if (colA[i]===label){ target=i+1; break; } }
  if (!target){
    target = last + 1;
    sh.insertRowAfter(last);
    sh.getRange(target,1).setValue(label);
  }
  try{
    const rich = SpreadsheetApp.newRichTextValue().setText('Åpne dokument').setLinkUrl(url).build();
    sh.getRange(target,2).setRichTextValue(rich);
  }catch(_){
    sh.getRange(target,2).setValue(url);
  }
}

/* --------------------------------- Helpers --------------------------------- */
function findSheetByCandidates_(names){
  const ss = SpreadsheetApp.getActive();
  for (const n of names){ const s = ss.getSheetByName(n); if (s) return s; }
  const low = names.map(n=>String(n).toLowerCase());
  for (const s of ss.getSheets()){ if (low.includes(s.getName().toLowerCase())) return s; }
  return null;
}
function firstRowPreview_(sheet, maxCols){
  try{
    const c = Math.min(sheet.getLastColumn(), maxCols || 10);
    if (sheet.getLastRow() < 1 || c < 1) return '';
    const vals = sheet.getRange(1,1,1,c).getValues()[0];
    return vals.map(v=>String(v||'').trim()).filter(Boolean).join(' | ');
  }catch(e){ return ''; }
}
/** Lokal Konfig get/set. value===undefined => GET; ellers SET. */
function upsertKonfigLocal_(key, value, desc){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PROJ_OVERVIEW.CONFIG.KONFIG_SHEET);
  if (!sh){ sh = ss.insertSheet(PROJ_OVERVIEW.CONFIG.KONFIG_SHEET); sh.appendRow(['Nøkkel','Verdi','Beskrivelse']); }
  const last = sh.getLastRow();
  if (last<=1){
    if (value===undefined) return null;
    sh.appendRow([key, value, desc||'']); return value;
  }
  const keys = sh.getRange(2,1,last-1,1).getValues().map(r=>String(r[0]||'').trim());
  const idx = keys.findIndex(k=>k===key);
  if (value===undefined){ return idx>=0 ? sh.getRange(idx+2,2).getValue() : null; }
  if (idx>=0){
    sh.getRange(idx+2,2).setValue(value);
    if (desc) sh.getRange(idx+2,3).setValue(desc);
  }else{
    sh.appendRow([key, value, desc||'']);
  }
  return value;
}
