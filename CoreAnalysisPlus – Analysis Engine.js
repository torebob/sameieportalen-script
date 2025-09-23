/**
 * CoreAnalysisPlus - Analysis Engine (All-in-One v1.9.2)
 *
 * - Safe dependency bridges (works with or without external helpers)
 * - Comprehensive analysis orchestrator with progress callbacks
 * - Requirements reader, Jaccard dedupe, gap analysis + recommendations
 * - Mermaid graph builder (Triggers, Menus, Functions, Sheets) with similarity edges
 * - Historical metrics logging (Analysis_Log)
 * - Minimal fallback collectors (when DataCollector is not present)
 * - Lightweight dashboard (HtmlService) to render Mermaid graph
 * - Sheets menu: Run Analysis, Open Dashboard
 */

var CORE_ANALYSIS_CFG = (typeof CORE_ANALYSIS_CFG === 'object' && CORE_ANALYSIS_CFG) ? CORE_ANALYSIS_CFG : {
  VERSION: '1.9.2',
  NAMES: {
    kravSheet: ['Requirements','Krav','Kravliste']
  },
  HEADERS: {
    krav: {
      id:       ['kravid','krav id','krav-id','id'],
      text:     ['krav','beskrivelse','tekst','requirement','description','text'],
      priority: ['prioritet','prio','priority'],
      progress: ['fremdrift %','fremdrift%','fremdrift','progress','progress %','progress%']
    }
  },
  LARGE_DATA_SHEETS: 25,
  LARGE_DATA_MAXCOLS: 60,
  LARGE_DATA_TOTALROWS: 30000, // removed numeric underscore
  TOKEN_MIN_LEN: 2,
  DEFAULT_JACCARD_THRESHOLD: 0.78,
  GRAPH_SIM_THRESHOLD: 0.60,
  GRAPH_MAX_EDGES: 400,
  GRAPH_MAX_NODES: 300,
  ANALYSIS_LOG_SHEET: 'Analysis_Log'
};

/* --------------------------- Safe dependency bridges ------------------------ */

function __ae_log_(){
  try{
    if (typeof _getLoggerPlus_ === 'function') return _getLoggerPlus_();
  }catch(_){}
  return {
    debug: function(fn,msg,d){ try{ console.log('[DEBUG]',fn||'',msg||'',d||''); }catch(_){ } },
    info:  function(fn,msg,d){ try{ console.log('[INFO]', fn||'',msg||'',d||''); }catch(_){ } },
    warn:  function(fn,msg,d){ try{ console.warn('[WARN]', fn||'',msg||'',d||''); }catch(_){ } },
    error: function(fn,msg,d){ try{ console.error('[ERROR]',fn||'',msg||'',d||''); }catch(_){ } }
  };
}
function __ae_cfgGet_(key, fallback){
  try{
    if (typeof _cfgGet_ === 'function') return _cfgGet_(key, fallback);
  }catch(_){}
  try{
    if (typeof CORE_ANALYSIS_CFG !== 'undefined' && CORE_ANALYSIS_CFG &&
        Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)){
      return CORE_ANALYSIS_CFG[key];
    }
  }catch(_){}
  return fallback;
}
function __ae_numCfg_(key, fallback){
  var v = Number(__ae_cfgGet_(key, fallback));
  return isNaN(v) ? Number(fallback) : v;
}
function __ae_boolCfg_(key, fallback){
  var v = __ae_cfgGet_(key, fallback);
  if (typeof v === 'boolean') return v;
  var s = String(v).trim().toLowerCase();
  return (s === 'true' || s === '1' || s === 'ja' || s === 'on' || s === 'enabled');
}

/* -------------------------------- Constants -------------------------------- */

var __AE_CONST = {
  TOKEN_MIN_LEN: __ae_numCfg_('TOKEN_MIN_LEN', CORE_ANALYSIS_CFG.TOKEN_MIN_LEN),
  JACCARD_TH:   __ae_numCfg_('DEFAULT_JACCARD_THRESHOLD', CORE_ANALYSIS_CFG.DEFAULT_JACCARD_THRESHOLD),
  LD_SHEETS:    __ae_numCfg_('LARGE_DATA_SHEETS', CORE_ANALYSIS_CFG.LARGE_DATA_SHEETS),
  LD_MAXCOLS:   __ae_numCfg_('LARGE_DATA_MAXCOLS', CORE_ANALYSIS_CFG.LARGE_DATA_MAXCOLS),
  LD_TOTALROWS: __ae_numCfg_('LARGE_DATA_TOTALROWS', CORE_ANALYSIS_CFG.LARGE_DATA_TOTALROWS),
  GRAPH_SIM_TH:   __ae_numCfg_('GRAPH_SIM_THRESHOLD', CORE_ANALYSIS_CFG.GRAPH_SIM_THRESHOLD),
  GRAPH_MAX_EDGES:__ae_numCfg_('GRAPH_MAX_EDGES', CORE_ANALYSIS_CFG.GRAPH_MAX_EDGES),
  GRAPH_MAX_NODES:__ae_numCfg_('GRAPH_MAX_NODES', CORE_ANALYSIS_CFG.GRAPH_MAX_NODES),
  ANALYSIS_LOG_SHEET: String(__ae_cfgGet_('ANALYSIS_LOG_SHEET', CORE_ANALYSIS_CFG.ANALYSIS_LOG_SHEET) || 'Analysis_Log')
};

/* ---------------------------- Tokenization & Diff --------------------------- */

var __ae_tokenCache = Object.create(null);
function __ae_safeStr_(v){ return (v === null || v === undefined) ? '' : String(v); }
function __ae_tokens_(s){
  s = __ae_safeStr_(s).toLowerCase();
  if (__ae_tokenCache[s]) return __ae_tokenCache[s];
  var parts = s.split(/[^a-z0-9æøå]+/).filter(Boolean);
  var min = __AE_CONST.TOKEN_MIN_LEN;
  if (min > 1) parts = parts.filter(function(t){ return t.length >= min; });
  __ae_tokenCache[s] = parts;
  return parts;
}
function __ae_jaccard_(a, b, minTh){
  var A = __ae_tokens_(a), B = __ae_tokens_(b);
  if (A.length === 0 && B.length === 0) return 1.0;
  if (A.length === 0 || B.length === 0) return 0.0;
  var minLen = Math.min(A.length, B.length), maxLen = Math.max(A.length, B.length);
  var upper = minLen / maxLen;
  if (typeof minTh === 'number' && upper < minTh) return 0.0;
  var setA = Object.create(null);
  for (var i=0;i<A.length;i++) setA[A[i]] = true;
  var inter = 0;
  var setB = Object.create(null);
  for (var j=0;j<B.length;j++){ setB[B[j]] = true; if (setA[B[j]]) inter++; }
  var union = Object.keys(setA).length + Object.keys(setB).length - inter;
  return inter / (union || 1);
}
function __ae_headerTokens_(headerPreview){
  return String(headerPreview||'')
    .split('|')
    .map(function(s){ return String(s||'').trim(); })
    .filter(Boolean);
}

/* ----------------------------- Fallback Collectors -------------------------- */

function __ae_collectMetadata_fallback_(){
  var ss = SpreadsheetApp.getActive();
  return {
    spreadsheetName: ss.getName(),
    spreadsheetUrl:  ss.getUrl(),
    spreadsheetId:   ss.getId(),
    timeZone:        ss.getSpreadsheetTimeZone && ss.getSpreadsheetTimeZone(),
    locale:          ss.getSpreadsheetLocale && ss.getSpreadsheetLocale(),
    sheetsCount:     (ss.getSheets()||[]).length,
    user:            (function(){ try{ return Session.getActiveUser().getEmail(); }catch(_){ return ''; } })()
  };
}
function __ae_collectTriggers_fallback_(){
  var out = [];
  try{
    var trig = ScriptApp.getProjectTriggers() || [];
    trig.forEach(function(t){
      var eventType='', source='', handler='';
      try{ handler = String(t.getHandlerFunction()||''); }catch(_){}
      try{ eventType = String(t.getEventType && t.getEventType()); }catch(_){}
      try{ source = String(t.getTriggerSource && t.getTriggerSource()); }catch(_){}
      out.push({ handler:handler, eventType:(eventType||'UNKNOWN'), source:(source||'UNKNOWN'), raw:{eventType,source} });
    });
  }catch(_){}
  return out;
}
function __ae_collectMenus_fallback_(){ return []; }
function __ae_collectDataModel_fallback_(){
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets() || [];
  var out = [];
  var headerDup = {};
  for (var i=0;i<sheets.length;i++){
    var sh = sheets[i];
    var cols = sh.getLastColumn();
    var hdr = cols > 0 ? (sh.getRange(1,1,1,cols).getValues()[0] || []) : [];
    var preview = hdr.slice(0,10).map(function(h){ return String(h||'').trim(); }).filter(Boolean).join(' | ');
    hdr.forEach(function(h, idx){
      var norm = String(h||'').trim().toLowerCase();
      if (!norm) return;
      if (!headerDup[norm]) headerDup[norm] = [];
      headerDup[norm].push({ sheet: sh.getName(), col: idx+1 });
    });
    out.push({
      name: sh.getName(),
      rows: sh.getLastRow(),
      columns: cols,
      hidden: (typeof sh.isSheetHidden === 'function') ? sh.isSheetHidden() : false,
      headerPreview: preview,
      typesByHeader: {}
    });
  }
  var duplicates = [];
  Object.keys(headerDup).forEach(function(h){
    if (headerDup[h] && headerDup[h].length > 1) {
      duplicates.push({ header: h, occurrences: headerDup[h] });
    }
  });
  return { sheets: out, headerDuplicates: duplicates };
}

/* ------------------------------ Public: Analysis ---------------------------- */

function performComprehensiveAnalysis_(options){
  var opt = options || {};
  var log = __ae_log_(), fn = 'performComprehensiveAnalysis_';
  var started = Date.now();

  var meta     = (typeof _collectMetadata_       === 'function') ? _collectMetadata_()        : __ae_collectMetadata_fallback_();
  var triggers = (typeof _collectTriggers_       === 'function') ? _collectTriggers_()        : __ae_collectTriggers_fallback_();
  var menus    = (typeof _collectMenuFunctions_  === 'function') ? _collectMenuFunctions_()   : __ae_collectMenus_fallback_();
  var data     = (typeof _collectDataModel_      === 'function') ? _collectDataModel_(opt.progressCb) : __ae_collectDataModel_fallback_();

  var sheetsArr = data.sheets || [];
  var sheetsScanned = sheetsArr.length;
  var totalRows = sheetsArr.reduce(function(sum,s){ return sum + (s.rows||0); }, 0);
  var maxCols = sheetsArr.reduce(function(m,s){ return Math.max(m, s.columns||0); }, 0);
  var durationMs = Date.now() - started;

  var isLarge = (sheetsScanned >= __AE_CONST.LD_SHEETS) || (maxCols >= __AE_CONST.LD_MAXCOLS) || (totalRows >= __AE_CONST.LD_TOTALROWS);
  if (isLarge){
    log.info(fn, 'Large dataset detected', { sheetsScanned: sheetsScanned, maxCols: maxCols, totalRows: totalRows, thresholds: {S:__AE_CONST.LD_SHEETS,C:__AE_CONST.LD_MAXCOLS,R:__AE_CONST.LD_TOTALROWS} });
  }

  var functions = _ae_mergeFunctionInventory_(triggers, menus);

  var result = {
    metadata: meta,
    triggers: { count: Array.isArray(triggers)?triggers.length:0, details: triggers },
    menus:    { fromSheets: menus },
    functions:{ global: functions, private: [] },
    sheets:   { count: sheetsArr.length, sheets: sheetsArr, headerDuplicates: data.headerDuplicates||[] },
    performanceMetrics: { sheetsScanned: sheetsScanned, totalRows: totalRows, maxCols: maxCols, scanDurationMs: durationMs },
    version: __ae_cfgGet_('VERSION', CORE_ANALYSIS_CFG.VERSION)
  };

  try { ae_writeAnalysisLog_(result); } catch(e){ log.warn(fn, 'Failed to write Analysis_Log', {error: e && e.message}); }

  log.info(fn, 'Full analysis complete', {
    ms: durationMs,
    sheets: result.sheets.count,
    triggers: result.triggers.count,
    functions: result.functions.global.length
  });
  return result;
}

/* ----------------------- Public: Requirements reader ------------------------ */

function readRequirementsForGapAnalysis_(){
  var log = __ae_log_(), fn = 'readRequirementsForGapAnalysis_';
  try{
    var names = __ae_cfgGet_('NAMES', CORE_ANALYSIS_CFG.NAMES);
    var sh = _ae_getSheetByAnyName_(names.kravSheet);
    if (!sh) return [];
    var vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return [];

    var headers = vals[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
    var KH = __ae_cfgGet_('HEADERS', CORE_ANALYSIS_CFG.HEADERS).krav;
    var idIdx = _ae_indexOfHeaderAny_(headers, KH.id);
    var textIdx = _ae_indexOfHeaderAny_(headers, KH.text);
    var prioIdx = _ae_indexOfHeaderAny_(headers, KH.priority);
    var progIdx = _ae_indexOfHeaderAny_(headers, KH.progress);

    var out = [];
    for (var r=1;r<vals.length;r++){
      var row = vals[r] || [];
      out.push({
        id: (idIdx>=0 ? row[idIdx] : '') || '',
        text: String((textIdx>=0 ? row[textIdx] : '') || ''),
        priority: String((prioIdx>=0 ? row[prioIdx] : '') || ''),
        progressPct: Number((progIdx>=0 ? row[progIdx] : 0) || 0)
      });
    }
    return out;
  }catch(e){
    log.error(fn, 'Failed to read requirements', { error: e.message, stack: e.stack });
    return [];
  }
}

/* -------------------------- Public: Dedupe + Gap ---------------------------- */

function dedupeCandidates_(candidates, existing, threshold){
  var log = __ae_log_(), fn = 'dedupeCandidates_';
  if (!Array.isArray(candidates)) { log.warn(fn, 'candidates must be array'); return []; }
  if (!Array.isArray(existing))   { existing = []; }

  var th = _ae_getJaccardThreshold_(threshold);
  var existTexts = existing.map(function(e){ return String(e.text||''); });
  var seen = [];
  var out = [];

  for (var i=0;i<candidates.length;i++){
    var c = candidates[i];
    var t = String(c && c.text || '');
    if (!t) continue;
    var dupExisting = existTexts.some(function(et){ return __ae_jaccard_(t, et, th) >= th; });
    if (dupExisting) continue;
    var dupNew = seen.some(function(s){ return __ae_jaccard_(t, s, th) >= th; });
    if (dupNew) continue;
    seen.push(t);
    out.push(c);
  }
  return out;
}

function performGapAnalysis_(analysis, existing, newDeduped){
  var A = analysis || {}, E = existing || [], N = newDeduped || [];
  var zero = E.filter(function(r){ return Number(r.progressPct||0) === 0; });
  var partial = E.filter(function(r){ var p=Number(r.progressPct||0); return p>0 && p<100; });
  var full = E.filter(function(r){ return Number(r.progressPct||0) >= 100; });

  var publicFns = (A.functions && A.functions.global || []).map(function(f){ return String(f.name||''); }).filter(Boolean);
  var kravTexts = E.map(function(r){ return String(r.text||'').toLowerCase(); });

  var undocumented = [];
  for (var i=0;i<publicFns.length;i++){
    var fn = publicFns[i], low = fn.toLowerCase();
    var inKrav = kravTexts.some(function(kt){ return kt.indexOf(low) >= 0; });
    var inNew  = N.some(function(c){ return String(c.text||'').toLowerCase().indexOf(low) >= 0; });
    if (!inKrav && !inNew){
      undocumented.push({ function: fn });
    }
  }

  var totalReq = E.length;
  var implementedPct = totalReq ? Math.round((full.length/totalReq)*100) : 0;
  var coverageScore = Math.round((implementedPct * 0.7) + ((E.length ? (1 - undocumented.length/Math.max(1,publicFns.length))*100 : 0) * 0.3));

  var recommendations = [];
  if (undocumented.length > 0) recommendations.push('Dokumenter funksjoner uten krav (se "undocumentedFunctions").');
  if (zero.length > 0)        recommendations.push('Start med MA-kra v med 0% fremdrift.');
  if (partial.length > 0)     recommendations.push('Fullfor delvis implementerte krav for rask gevinst.');
  if ((A.sheets && A.sheets.headerDuplicates || []).length > 0) recommendations.push('Rydd opp i duplikate headere pa tvers av ark.');

  return {
    requirements: { unimplemented: zero, partial: partial, complete: full },
    undocumentedFunctions: undocumented,
    coverage: {
      totalRequirements: totalReq,
      implementedPct: implementedPct,
      codeHealthScore: coverageScore
    },
    recommendations: recommendations
  };
}

/* -------------------------------- Graph Builder ----------------------------- */

function ae_buildMermaid_(analysis){
  var A = analysis || {};
  var lines = ['graph LR'];
  function safeId(s){ return (String(s||'').replace(/[^a-zA-Z0-9_]/g,'_') || 'X'); }
  function add(line){ lines.push(line); }

  var sheets = (A.sheets && A.sheets.sheets) || [];
  var funcs  = (A.functions && A.functions.global) || [];
  var trigs  = (A.triggers && A.triggers.details) || [];
  var menus  = (A.menus && A.menus.fromSheets) || [];

  var MAXN = __AE_CONST.GRAPH_MAX_NODES;
  if (sheets.length + funcs.length + trigs.length + menus.length > MAXN){
    sheets = sheets.slice(0, Math.floor(MAXN*0.35));
    funcs  = funcs.slice(0, Math.floor(MAXN*0.35));
    trigs  = trigs.slice(0, Math.floor(MAXN*0.15));
    menus  = menus.slice(0, Math.floor(MAXN*0.15));
  }

  var nodeCount=0, edgeCount=0, MAXE=__AE_CONST.GRAPH_MAX_EDGES;

  add('subgraph Triggers');
  for (var k=0;k<trigs.length;k++){
    var t = trigs[k];
    var et = String(t.eventType||'UNKNOWN');
    var hf = String(t.handler||'');
    if (!hf) continue;
    add(safeId('trig_'+hf)+'["Trig: '+et+'\\n'+hf+'"]:::trig'); nodeCount++;
    if (nodeCount>=MAXN) break;
  }
  add('end');

  add('subgraph Menus');
  for (var m=0;m<menus.length;m++){
    var mi = menus[m];
    var ttl = String(mi.title || mi.functionName || '');
    if (!ttl) continue;
    add(safeId('menu_'+ttl)+'["Menu: '+ttl+'"]:::menu'); nodeCount++;
    if (nodeCount>=MAXN) break;
  }
  add('end');

  add('subgraph Functions');
  for (var j=0;j<funcs.length;j++){
    var fnn = String(funcs[j].name||'');
    if (!fnn) continue;
    add(safeId('fn_'+fnn)+'(("Fn: '+fnn+'")):::fn'); nodeCount++;
    if (nodeCount>=MAXN) break;
  }
  add('end');

  add('subgraph Sheets');
  for (var i=0;i<sheets.length;i++){
    var sn = String(sheets[i].name||'');
    if (!sn) continue;
    add(safeId('sheet_'+sn)+'["Sheet: '+sn+'"]:::sheet'); nodeCount++;
    if (nodeCount>=MAXN) break;
  }
  add('end');

  for (var t2=0;t2<trigs.length && edgeCount<MAXE;t2++){
    var tr = trigs[t2], h = String(tr.handler||''), et2 = String(tr.eventType||'UNKNOWN');
    if (!h) continue;
    add(safeId('trig_'+h)+' -- "'+et2+'" --> '+safeId('fn_'+h));
    edgeCount++;
  }

  for (var mm=0;mm<menus.length && edgeCount<MAXE;mm++){
    var me = menus[mm];
    var fcall = String(me.functionName||'');
    var ttl2  = String(me.title || me.functionName || '');
    if (!fcall || !ttl2) continue;
    add(safeId('menu_'+ttl2)+' -- "Menu" --> '+safeId('fn_'+fcall));
    edgeCount++;
  }

  var SIM_TH = __AE_CONST.GRAPH_SIM_TH;
  for (var f=0; f<funcs.length && edgeCount<MAXE; f++){
    var fname = String(funcs[f].name||'');
    if (!fname) continue;
    for (var s=0; s<sheets.length && edgeCount<MAXE; s++){
      var sh  = sheets[s];
      var sname = String(sh.name||'');
      var hp = String(sh.headerPreview||'');
      var hdrTokens = __ae_headerTokens_(hp);

      var sim = __ae_jaccard_(fname, sname, 0.30);
      for (var hix=0; hix<hdrTokens.length; hix++){
        var tok = hdrTokens[hix];
        sim = Math.max(sim, __ae_jaccard_(fname, tok, 0.30));
        if (sim >= 0.999) break;
      }
      if (sim >= SIM_TH){
        var lbl = 'sim:'+String(Math.round(sim*100)/100);
        add(safeId('fn_'+fname)+' -- "'+lbl+'" --> '+safeId('sheet_'+sname));
        edgeCount++;
        if (edgeCount>=MAXE) break;
      }
    }
  }

  add('Legend[/"Legend"\\nEdges: event type, "Menu", or sim:0.6+\\nFn->Sheet = semantic match/]:::legend');
  add('classDef sheet fill:#ECFCCB,stroke:#84CC16,color:#1a1a1a;');
  add('classDef fn fill:#E0F2FE,stroke:#38BDF8,color:#1a1a1a;');
  add('classDef trig fill:#FFE4E6,stroke:#FB7185,color:#1a1a1a;');
  add('classDef menu fill:#F3E8FF,stroke:#A78BFA,color:#1a1a1a;');
  add('classDef legend fill:#F8FAFC,stroke:#94A3B8,color:#334155;');

  if (edgeCount>=MAXE || nodeCount>=MAXN){
    add('NoteTrunc[/"Graf avkortet"\\nNoder:'+nodeCount+' / '+MAXN+'\\nKanter:'+edgeCount+' / '+MAXE+'/]:::legend');
  }
  return lines.join('\n');
}

/* ----------------------------- Analysis Logging ----------------------------- */

function ae_writeAnalysisLog_(analysis){
  var ss = SpreadsheetApp.getActive();
  var name = __AE_CONST.ANALYSIS_LOG_SHEET;
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0){
    sh.appendRow(['Timestamp','SpreadsheetId','SheetsScanned','TotalRows','MaxCols','ScanMs','Triggers','Functions','UndocumentedFns','CoveragePct','HealthScore','CommitHash']);
  }

  var perf = analysis.performanceMetrics || {};
  var undocumentedCount = 0, coveragePct = null, health = null;

  try{
    var gap = performGapAnalysis_(analysis, readRequirementsForGapAnalysis_(), []);
    undocumentedCount = (gap.undocumentedFunctions||[]).length;
    coveragePct = (gap.coverage && gap.coverage.implementedPct) || 0;
    health = (gap.coverage && gap.coverage.codeHealthScore) || 0;
  }catch(_){}

  var commit = (function(){
    try{
      var p = PropertiesService.getScriptProperties().getProperty('COMMIT_HASH');
      return p || '';
    }catch(_){ return ''; }
  })();

  sh.appendRow([
    new Date(),
    (analysis.metadata && analysis.metadata.spreadsheetId) || '',
    perf.sheetsScanned||0, perf.totalRows||0, perf.maxCols||0, perf.scanDurationMs||0,
    (analysis.triggers && analysis.triggers.count)||0,
    (analysis.functions && analysis.functions.global && analysis.functions.global.length)||0,
    undocumentedCount, coveragePct, health, commit
  ]);
}

/* ------------------------------- Dashboard ---------------------------------- */

function ae_showDashboard(){
  var analysis = performComprehensiveAnalysis_();
  var mer = ae_buildMermaid_(analysis);

  var html = HtmlService
    .createHtmlOutput(
      '<html><head>' +
      '<meta charset="utf-8"/>' +
      '<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>' +
      '<style>body{font-family:Inter,system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:12px} .wrap{white-space:pre} .meta{margin:8px 0 16px;color:#475569}</style>' +
      '</head><body>' +
      '<h2>CoreAnalysisPlus - Graph</h2>' +
      '<div class="meta">Version: '+(analysis.version||'')+' • Sheets: '+(analysis.sheets && analysis.sheets.count || 0)+' • Triggers: '+(analysis.triggers && analysis.triggers.count || 0)+'</div>' +
      '<div class="mermaid">'+
      mer.replace(/</g,'&lt;').replace(/>/g,'&gt;') +
      '</div>' +
      '<script>mermaid.initialize({startOnLoad:true, theme:"default"});</script>' +
      '</body></html>'
    )
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'CoreAnalysisPlus - Dashboard');
}

/* --------------------------------- Helpers ---------------------------------- */

function _ae_mergeFunctionInventory_(triggers, menus){
  var set = Object.create(null), out=[];
  (triggers||[]).forEach(function(t){
    var n = String(t.handler||'').trim();
    if (!n || set[n]) return;
    set[n]=true; out.push({ name:n, source:'trigger', eventType: t.eventType||'' });
  });
  (menus||[]).forEach(function(m){
    var n = String(m.functionName||'').trim();
    if (!n || set[n]) return;
    set[n]=true; out.push({ name:n, source:'menu', title: m.title||'' });
  });
  return out;
}
function _ae_indexOfHeaderAny_(headersLower, alts){
  if (!Array.isArray(headersLower) || !Array.isArray(alts)) return -1;
  var wants = alts.map(function(x){ return String(x||'').trim().toLowerCase(); });
  for (var i=0;i<headersLower.length;i++){
    var h = String(headersLower[i]||'').trim().toLowerCase();
    for (var j=0;j<wants.length;j++){ if (h === wants[j]) return i; }
  }
  return -1;
}
function _ae_getSheetByAnyName_(candidates){
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets() || [];
  var cand = (Array.isArray(candidates) ? candidates : [candidates])
    .map(function(s){ return String(s||'').toLowerCase().replace(/\s+/g,'').replace(/_/g,''); });
  for (var i=0;i<sheets.length;i++){
    var n = String(sheets[i].getName()||'').toLowerCase().replace(/\s+/g,'').replace(/_/g,'');
    for (var j=0;j<cand.length;j++) if (n === cand[j]) return sheets[i];
  }
  return null;
}
function _ae_getJaccardThreshold_(provided){
  if (typeof provided === 'number' && !isNaN(provided)) return Math.max(0, Math.min(1, provided));
  try{
    if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS){
      if (typeof CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD === 'number') return CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD;
      if (typeof CONFIG_PLUS.DEDUPLE_JACCARD === 'number') return CONFIG_PLUS.DEDUPLE_JACCARD;
    }
  }catch(_){}
  return __AE_CONST.JACCARD_TH;
}

/* --------------------------------- Smoke/Test -------------------------------- */

function runCoreAnalysis_Smoke(){
  var res = performComprehensiveAnalysis_();
  var m = res.performanceMetrics || {};
  try { __ae_log_().info('runCoreAnalysis_Smoke', 'Core analysis metrics', m); } catch(_){}
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Integritetsrapport') || ss.getSheetByName('Rapport') || ss.insertSheet('Integritetsrapport');
  if (sh.getLastRow() === 0) sh.appendRow(['Kj.Dato','Kategori','Nokkel','Status','Detaljer']);
  sh.appendRow([new Date(), 'Analyse', 'CoreAnalysisPlus', 'OK', JSON.stringify(m)]);
}

/* ----------------------------------- Menu ----------------------------------- */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Analysis')
    .addItem('Run Analysis (log)', 'runCoreAnalysis_Smoke')
    .addItem('Open Dashboard (Mermaid)', 'ae_showDashboard')
    .addToUi();
}
